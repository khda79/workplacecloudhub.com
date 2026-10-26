#Requires -Version 7.0

<#
.SYNOPSIS
    Sends one SmartM365 test mail through Send-SmartM365Mail with the selected tenant profile.

.DESCRIPTION
    Loads the tenant profile (global, tenant and mail profile files), shows the effective
    SendMailMode, From, To and Cc, then sends a short HTML test mail through the same
    Send-SmartM365Mail path as the inventory scripts (Graph app-only or SMTP).

    No collection, upload or configuration change is performed.

.PARAMETER Tenant
    Tenant profile key. Defaults to test.

.PARAMETER From
    Sender mailbox. When omitted, the profile value is used.

.PARAMETER To
    Recipients, separated by ';' or ','. When omitted, the profile To (or ErrorMailTo) is used.

.PARAMETER Cc
    Optional Cc recipients. When omitted, the profile value is used.

.PARAMETER AsScript
    Script name whose Scripts.<name> entry in <tenant>.mail.local.json.txt is applied,
    for example SmartM365-Licences-Inventory. Only From and Cc come from that entry.

.PARAMETER Subject
    Mail subject.

.EXAMPLE
    .\Setup\Test-SmartM365-SendMail.ps1 -Tenant prod -From "sender@contoso.com" -To "admin@contoso.com"

.EXAMPLE
    .\Setup\Test-SmartM365-SendMail.ps1 -Tenant prod -AsScript SmartM365-Licences-Inventory -To "admin@contoso.com" -WhatIf

.VERSION
1.0
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [string]$Tenant = 'test',
    [string]$From,
    [string]$To,
    [string]$Cc,
    [string]$AsScript = '',
    [string]$Subject = 'SmartM365 mail test'
)
$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidates = @(
            (Join-Path -Path $d -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $d -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )
        foreach ($p in $candidates) {
            if (Test-Path -LiteralPath $p) { return $p }
        }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $tenantContextPath
$script:SmartM365RootPath = Find-SmartM365Root -StartPath $PSScriptRoot
Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot -ScriptName $AsScript | Out-Null

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$modulePath = Join-Path -Path $script:SmartM365RootPath -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
if (-not (Test-Path -LiteralPath $modulePath)) {
    throw "SmartM365.Core module was not found at '$modulePath'."
}
Import-Module $modulePath -MinimumVersion '1.0.58' -Force -ErrorAction Stop

function Get-ProfileMailValue {
    param([Parameter(Mandatory)][string]$Name)
    $config = $global:SmartM365GlobalConfig
    $value = $null
    if ($config -is [System.Collections.IDictionary]) {
        if ($config.Contains($Name)) { $value = $config[$Name] }
    }
    elseif ($null -ne $config -and $config.PSObject.Properties[$Name]) {
        $value = $config.$Name
    }
    $text = if ($null -eq $value) { '' } else { ([string]$value).Trim() }
    if ($text -in @('__USE_GLOBAL__', 'USE_GLOBAL')) { return '' }
    return $text
}

# Same resolution order as Send-SmartM365Mail: explicit parameter, then profile value (To falls back to ErrorMailTo).
$effective = [ordered]@{}
foreach ($name in @('SendMailMode', 'SmtpServer', 'From', 'To', 'Cc')) {
    if ($PSBoundParameters.ContainsKey($name)) { $effective[$name] = [string]$PSBoundParameters[$name]; continue }
    $effective[$name] = Get-ProfileMailValue -Name $name
}
if ([string]::IsNullOrWhiteSpace($effective['To'])) { $effective['To'] = Get-ProfileMailValue -Name 'ErrorMailTo' }
if ([string]::IsNullOrWhiteSpace($effective['SendMailMode'])) {
    $effective['SendMailMode'] = if ([string]::IsNullOrWhiteSpace($effective['SmtpServer'])) { 'Graph' } else { 'SMTP' }
}

Write-Information ("Tenant profile : {0}" -f $Tenant) -InformationAction Continue
Write-Information ("Mail entry     : {0}" -f $(if ($AsScript) { "Scripts.$AsScript" } else { '(none)' })) -InformationAction Continue
foreach ($name in $effective.Keys) {
    Write-Information ("{0,-14} : {1}" -f $name, $(if ($effective[$name]) { $effective[$name] } else { '(empty)' })) -InformationAction Continue
}
if ([string]::IsNullOrWhiteSpace($effective['From']) -or [string]::IsNullOrWhiteSpace($effective['To'])) {
    throw 'From and To are required: pass -From/-To or set them in the tenant profile.'
}

$sentAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
$encode = { param($Value) [System.Net.WebUtility]::HtmlEncode([string]$Value) }
$bodyHtml = @"
<p>SmartM365 test mail sent by Test-SmartM365-SendMail.ps1.</p>
<table>
<tr><td>Tenant profile</td><td>$(& $encode $Tenant)</td></tr>
<tr><td>Mail entry</td><td>$(& $encode $(if ($AsScript) { "Scripts.$AsScript" } else { '(none)' }))</td></tr>
<tr><td>Send mode</td><td>$(& $encode $effective['SendMailMode'])</td></tr>
<tr><td>From</td><td>$(& $encode $effective['From'])</td></tr>
<tr><td>Computer</td><td>$(& $encode $env:COMPUTERNAME)</td></tr>
<tr><td>Sent at</td><td>$(& $encode $sentAt)</td></tr>
</table>
"@

$mailParams = @{ Subject = $Subject; BodyHtml = $bodyHtml; MailPurpose = 'Report' }
foreach ($name in @('From', 'To', 'Cc')) {
    if ($PSBoundParameters.ContainsKey($name)) { $mailParams[$name] = $PSBoundParameters[$name] }
}

if ($PSCmdlet.ShouldProcess($effective['To'], ("Send test mail from {0}" -f $effective['From']))) {
    Send-SmartM365Mail @mailParams
    Write-Information ("[OK] Test mail sent from {0} to {1}." -f $effective['From'], $effective['To']) -InformationAction Continue
}
