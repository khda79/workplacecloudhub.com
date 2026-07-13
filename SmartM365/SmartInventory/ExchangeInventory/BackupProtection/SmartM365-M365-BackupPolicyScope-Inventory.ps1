#Requires -Version 7.0
<#
.SYNOPSIS
    Microsoft 365 Backup policy scope mailbox coverage inventory.

.DESCRIPTION
    Reads the configured Entra group used as the Microsoft 365 Backup mailbox policy scope,
    compares its user members with the latest Exchange Online mailbox inventory, and exports
    Power BI friendly CSV files. This is an expected-scope inventory, not proof that Microsoft
    365 Backup has protected each mailbox.

.PARAMETER Tenant
    Tenant profile key to load from Config/Tenants. Defaults to test.

.PARAMETER BackupPolicyScopeGroupId
    Entra group object id used by the Microsoft 365 Backup policy scope.

.PARAMETER BackupPolicyScopeGroupDisplayName
    Optional display name used in logs and CSVs.

.PARAMETER DataLastFolder
    Optional DATA-LAST folder override. Defaults to LatestCsvFolderPath from config.

.PARAMETER OutputPath
    Optional output directory override. Defaults to ScriptCsvLogFolderPath from config.

.PARAMETER DisableSharePointUpload
    Skips SharePoint upload for generated latest CSV files.

.PARAMETER MaxItems
    Limits group members and mailbox rows for smoke tests. Generated CSV names include _MAXITEMS.

.VERSION
1.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Minimum application permissions: GroupMember.Read.All, User.Read.All.
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$BackupPolicyScopeGroupId = '',
    [string]$BackupPolicyScopeGroupDisplayName = '',
    [string]$DataLastFolder = '',
    [string]$OutputPath = '',
    [switch]$DisableSharePointUpload,
    [int]$MaxItems = 0
)

if ($PSBoundParameters.ContainsKey('MaxItems') -and $MaxItems -gt 0) {
    $global:SmartM365MaxItems = [int]$MaxItems
    $global:SmartM365TestMaxItems = [int]$MaxItems
    $global:SmartM365IsMaxItemsRun = $true
}

$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidates = @(
            (Join-Path -Path $d -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $d -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )
        foreach ($p in $candidates) { if (Test-Path -LiteralPath $p) { return $p } }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $tenantContextPath
$script:SmartM365EffectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$MaximumFunctionCount = 32768
$ScriptVersion = '1.0'
$TaskName = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion"
$script:SmartM365GlobalConfig = $null
$script:LogPath = ''

function Import-SmartM365CoreModule {
    $searchRoot = $PSScriptRoot
    while ($searchRoot) {
        $candidate = Join-Path -Path $searchRoot -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
        if (Test-Path -LiteralPath $candidate) {
            Import-Module -Name $candidate -MinimumVersion '1.0.24' -Force -ErrorAction Stop
            return
        }
        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }
    throw 'SmartM365.Core module manifest not found.'
}

function Get-ScriptLocalConfig {
    $configPath = Join-Path -Path $PSScriptRoot -ChildPath ("{0}.local.json" -f [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))
    if (-not (Test-Path -LiteralPath $configPath)) {
        $templatePath = '{0}.template' -f $configPath
        if (Get-Command Initialize-SmartM365LocalJsonFromTemplate -ErrorAction SilentlyContinue) {
            Initialize-SmartM365LocalJsonFromTemplate -Path $configPath -TemplatePath $templatePath -ConfigDescription 'script local configuration' | Out-Null
        }
        else {
            if (-not (Test-Path -LiteralPath $templatePath)) { throw "Missing local config and template: $configPath" }
            Copy-Item -LiteralPath $templatePath -Destination $configPath -ErrorAction Stop
        }
    }
    return Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
}

function Get-SmartM365GlobalConfig {
    if ($null -ne $script:SmartM365GlobalConfig) { return $script:SmartM365GlobalConfig }
    if ($null -ne $script:SmartM365EffectiveConfig) {
        $script:SmartM365GlobalConfig = $script:SmartM365EffectiveConfig
        return $script:SmartM365GlobalConfig
    }
    $script:SmartM365GlobalConfig = [pscustomobject]@{}
    return $script:SmartM365GlobalConfig
}

function Resolve-SmartM365ConfigValue {
    param([AllowNull()]$Value)
    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) { return $Value }
    if ($Value -notmatch '\{\{[^}]+\}\}') { return $Value }
    $globalConfig = Get-SmartM365GlobalConfig
    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $matches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($matches.Count -eq 0) { break }
        $changed = $false
        foreach ($match in $matches) {
            $property = $globalConfig.PSObject.Properties[$match.Groups['Name'].Value]
            if ($null -eq $property -or $null -eq $property.Value) { continue }
            $tokenValue = Resolve-SmartM365ConfigValue -Value $property.Value
            if ($null -eq $tokenValue) { continue }
            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }
        if (-not $changed) { break }
    }
    return $resolved
}

function Get-ScriptLocalConfigValue {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$Name,
        $DefaultValue
    )
    $property = $Config.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        if ($property.Value -isnot [string]) { return Resolve-SmartM365ConfigValue -Value $property.Value }
        $text = $property.Value.Trim()
        if ($text -and $text -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) { return Resolve-SmartM365ConfigValue -Value $property.Value }
    }
    $globalConfig = Get-SmartM365GlobalConfig
    $globalProperty = $globalConfig.PSObject.Properties[$Name]
    if ($null -ne $globalProperty -and $null -ne $globalProperty.Value) {
        if ($globalProperty.Value -is [string] -and [string]::IsNullOrWhiteSpace($globalProperty.Value)) { return $DefaultValue }
        return Resolve-SmartM365ConfigValue -Value $globalProperty.Value
    }
    return $DefaultValue
}

function Write-ScopeLog {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO'
    )
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    if (-not [string]::IsNullOrWhiteSpace($script:LogPath)) {
        Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
    }
}

function Ensure-GraphAuthModule {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        throw 'Microsoft.Graph.Authentication module is required.'
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
}

function Connect-GraphAppOnly {
    param(
        [Parameter(Mandatory = $true)][string]$AppId,
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $true)][string]$Thumbprint
    )
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-ScopeLog -Message 'Connecting to Microsoft Graph with app-only certificate authentication.'
    Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome -ErrorAction Stop | Out-Null
    Write-ScopeLog -Message 'Microsoft Graph connected.' -Level SUCCESS
}

function Invoke-GraphGetCollection {
    param([Parameter(Mandatory = $true)][string]$Uri)
    $items = New-Object System.Collections.Generic.List[object]
    $nextUri = $Uri
    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject -ErrorAction Stop
        foreach ($item in @($response.value)) { $items.Add($item) | Out-Null }
        $nextProperty = $response.PSObject.Properties['@odata.nextLink']
        $nextUri = if ($null -ne $nextProperty) { [string]$nextProperty.Value } else { '' }
        if ($MaxItems -gt 0 -and $items.Count -ge $MaxItems) { break }
    }
    if ($MaxItems -gt 0 -and $items.Count -gt $MaxItems) { return @($items | Select-Object -First $MaxItems) }
    return @($items)
}

function Get-ObjectValue {
    param([AllowNull()]$Object, [Parameter(Mandatory = $true)][string[]]$Names)
    if ($null -eq $Object) { return $null }
    foreach ($name in $Names) {
        $property = $Object.PSObject.Properties[$name]
        if ($null -ne $property) { return $property.Value }
    }
    return $null
}

function Normalize-MailAddress {
    param([AllowNull()][string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    return $Value.Trim().ToLowerInvariant()
}

function Build-MailboxAddressIndex {
    param([object[]]$Mailboxes)
    $index = @{}
    foreach ($mailbox in $Mailboxes) {
        foreach ($propertyName in @('UserPrincipalName','UPN','PrimarySmtpAddress','WindowsEmailAddress','ExternalDirectoryObjectId')) {
            $value = Normalize-MailAddress ([string](Get-ObjectValue -Object $mailbox -Names @($propertyName)))
            if ($value -and -not $index.ContainsKey($value)) { $index[$value] = $mailbox }
        }
        $aliases = [string](Get-ObjectValue -Object $mailbox -Names @('EmailAddresses','ProxyAddresses'))
        foreach ($alias in @($aliases -split '[;|,]')) {
            $clean = (Normalize-MailAddress $alias) -replace '^smtp:', ''
            if ($clean -and -not $index.ContainsKey($clean)) { $index[$clean] = $mailbox }
        }
    }
    return $index
}

function ConvertTo-GroupMemberRow {
    param($Member, [string]$RunId, [string]$RunDateUtc, [string]$TenantName, [string]$GroupId, [string]$GroupDisplayName)
    $upn = [string](Get-ObjectValue -Object $Member -Names @('userPrincipalName'))
    $mail = [string](Get-ObjectValue -Object $Member -Names @('mail'))
    $enabled = Get-ObjectValue -Object $Member -Names @('accountEnabled')
    [pscustomobject]@{
        RunId = $RunId
        RunDateUtc = $RunDateUtc
        TenantName = $TenantName
        GroupId = $GroupId
        GroupDisplayName = $GroupDisplayName
        MemberId = [string](Get-ObjectValue -Object $Member -Names @('id'))
        MemberDisplayName = [string](Get-ObjectValue -Object $Member -Names @('displayName'))
        MemberUserPrincipalName = $upn
        MemberMail = $mail
        MemberType = [string](Get-ObjectValue -Object $Member -Names @('@odata.type'))
        AccountEnabled = if ($null -eq $enabled) { '' } else { [string]$enabled }
        Status = 'OK'
        NumericValue = 1
        TextValue = $upn
        Threshold = ''
        Details = 'Group member in expected Microsoft 365 Backup policy scope.'
    }
}

function ConvertTo-CoverageRows {
    param(
        [object[]]$Members,
        [hashtable]$MailboxIndex,
        [string]$RunId,
        [string]$RunDateUtc,
        [string]$TenantName,
        [string]$GroupId,
        [string]$GroupDisplayName
    )
    foreach ($member in $Members) {
        $keys = @(
            Normalize-MailAddress ([string](Get-ObjectValue -Object $member -Names @('userPrincipalName'))),
            Normalize-MailAddress ([string](Get-ObjectValue -Object $member -Names @('mail'))),
            Normalize-MailAddress ([string](Get-ObjectValue -Object $member -Names @('id')))
        ) | Where-Object { $_ }
        $mailbox = $null
        foreach ($key in $keys) {
            if ($MailboxIndex.ContainsKey($key)) { $mailbox = $MailboxIndex[$key]; break }
        }
        $hasMailbox = $null -ne $mailbox
        [pscustomobject]@{
            RunId = $RunId
            RunDateUtc = $RunDateUtc
            TenantName = $TenantName
            GroupId = $GroupId
            GroupDisplayName = $GroupDisplayName
            MemberId = [string](Get-ObjectValue -Object $member -Names @('id'))
            MemberDisplayName = [string](Get-ObjectValue -Object $member -Names @('displayName'))
            MemberUserPrincipalName = [string](Get-ObjectValue -Object $member -Names @('userPrincipalName'))
            MemberMail = [string](Get-ObjectValue -Object $member -Names @('mail'))
            MailboxFoundInInventory = $hasMailbox
            MailboxDisplayName = if ($hasMailbox) { [string](Get-ObjectValue -Object $mailbox -Names @('DisplayName','Name')) } else { '' }
            MailboxUserPrincipalName = if ($hasMailbox) { [string](Get-ObjectValue -Object $mailbox -Names @('UserPrincipalName','UPN')) } else { '' }
            PrimarySmtpAddress = if ($hasMailbox) { [string](Get-ObjectValue -Object $mailbox -Names @('PrimarySmtpAddress','WindowsEmailAddress')) } else { '' }
            RecipientTypeDetails = if ($hasMailbox) { [string](Get-ObjectValue -Object $mailbox -Names @('RecipientTypeDetails')) } else { '' }
            ExpectedPolicyScopeStatus = if ($hasMailbox) { 'InScopeMailbox' } else { 'GroupMemberWithoutMailboxInInventory' }
            ProtectionEvidenceStatus = 'NotVerifiedByThisScript'
            Status = if ($hasMailbox) { 'OK' } else { 'Warning' }
            NumericValue = if ($hasMailbox) { 1 } else { 0 }
            TextValue = if ($hasMailbox) { 'Mailbox matched in Exchange inventory' } else { 'No mailbox match in Exchange inventory' }
            Threshold = 'MailboxFoundInInventory=True'
            Details = 'Expected scope comparison only; use Microsoft 365 Backup protection-unit inventory for protection evidence.'
        }
    }
}

function Export-InventoryCsv {
    param(
        [Parameter(Mandatory = $true)][string]$BaseName,
        [Parameter(Mandatory = $true)][object[]]$Rows,
        [string[]]$Columns = @()
    )
    $name = if ($MaxItems -gt 0) { '{0}_MAXITEMS' -f $BaseName } else { $BaseName }
    Export-SmartM365Csv -BaseFileName $name -OutputPath $ScriptCsvLogFolderPath -GlobalPath $LatestCsvFolderPath -Data $Rows -Columns $Columns -Encoding UTF8 -Delimiter ',' -NoSharePointUpload:($DisableSharePointUpload.IsPresent) | Out-Null
}

Import-SmartM365CoreModule
Ensure-GraphAuthModule
$ScriptLocalConfig = Get-ScriptLocalConfig

$global:RetentionMaxCSV = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxLogs' -DefaultValue 30)
$global:EnableSharePointUpload = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointTargetFolderPath' -DefaultValue ''
$AppId = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AppId' -DefaultValue '')
$TenantId = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue '')
$Thumb = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumb' -DefaultValue '')
$OrgDomain = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue $Tenant)
$ScriptCsvLogFolderPath = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ScriptCsvLogFolderPath' -DefaultValue '')
$LatestCsvFolderPath = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '')
$LogAllRootPath = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LogAllRootPath' -DefaultValue '')
$InputDataLastFolder = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'InputDataLastFolder' -DefaultValue $LatestCsvFolderPath)

if ([string]::IsNullOrWhiteSpace($BackupPolicyScopeGroupId)) { $BackupPolicyScopeGroupId = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'BackupPolicyScopeGroupId' -DefaultValue '') }
if ([string]::IsNullOrWhiteSpace($BackupPolicyScopeGroupDisplayName)) { $BackupPolicyScopeGroupDisplayName = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'BackupPolicyScopeGroupDisplayName' -DefaultValue '') }
if (-not [string]::IsNullOrWhiteSpace($DataLastFolder)) { $InputDataLastFolder = $DataLastFolder }
if (-not [string]::IsNullOrWhiteSpace($OutputPath)) { $ScriptCsvLogFolderPath = $OutputPath }
if ([string]::IsNullOrWhiteSpace($ScriptCsvLogFolderPath)) { $ScriptCsvLogFolderPath = Join-Path -Path $PSScriptRoot -ChildPath 'Output' }
if ([string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) { $LatestCsvFolderPath = $ScriptCsvLogFolderPath }
if ([string]::IsNullOrWhiteSpace($InputDataLastFolder)) { $InputDataLastFolder = $LatestCsvFolderPath }
if ([string]::IsNullOrWhiteSpace($BackupPolicyScopeGroupId)) { throw 'BackupPolicyScopeGroupId is required.' }

$runId = [guid]::NewGuid().ToString()
$runDateUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logRoot = if ([string]::IsNullOrWhiteSpace($LogAllRootPath)) { Join-Path $ScriptCsvLogFolderPath 'Logs' } else { Join-Path $LogAllRootPath 'SmartM365-M365-BackupPolicyScope-Inventory' }
$script:LogPath = Join-Path -Path $logRoot -ChildPath ("SmartM365-M365-BackupPolicyScope-Inventory_{0}.log" -f $timestamp)
foreach ($folder in @($ScriptCsvLogFolderPath, $LatestCsvFolderPath, $logRoot)) {
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
}
Set-SmartM365CoreContext -RunId $runId -RunOutputRoot $ScriptCsvLogFolderPath -LatestOutputRoot $LatestCsvFolderPath -LogPath $script:LogPath
$global:LogTextFile = $script:LogPath

try {
    Write-ScopeLog -Message "Starting $TaskName. Tenant=$Tenant RunId=$runId"
    Write-ScopeLog -Message "This inventory reports expected backup policy scope membership; it does not verify Microsoft 365 Backup protection state."
    if ($MaxItems -gt 0) { Write-ScopeLog -Message "MaxItems test mode enabled: $MaxItems" -Level WARN }

    $mailboxCsvPath = Join-Path -Path $InputDataLastFolder -ChildPath 'Exchange_EXO_Mailboxes_AllDomains.csv'
    $resolvedMailboxCsvPath = Resolve-SmartM365CsvPathWithSharePointFallback -Path $mailboxCsvPath -Description 'Exchange Online mailbox inventory' -Required
    Write-ScopeLog -Message "Loading mailbox inventory: $resolvedMailboxCsvPath"
    $mailboxes = @(Import-Csv -LiteralPath $resolvedMailboxCsvPath -ErrorAction Stop)
    if ($MaxItems -gt 0) { $mailboxes = @($mailboxes | Select-Object -First $MaxItems) }
    Write-ScopeLog -Message "Mailbox inventory rows loaded: $($mailboxes.Count)"

    Connect-GraphAppOnly -AppId $AppId -TenantId $TenantId -Thumbprint $Thumb

    Invoke-SmartM365Preflight `
        -ScriptName $TaskName `
        -RequiredModules @('Microsoft.Graph.Authentication') `
        -OutputPaths @($ScriptCsvLogFolderPath, $LatestCsvFolderPath) `
        -RequiredGraphApplicationPermissions @('GroupMember.Read.All','User.Read.All') `
        -GraphProbeUris @("https://graph.microsoft.com/v1.0/groups/$BackupPolicyScopeGroupId/members?`$top=1") | Out-Null

    $select = 'id,displayName,userPrincipalName,mail,accountEnabled'
    $membersUri = "https://graph.microsoft.com/v1.0/groups/$BackupPolicyScopeGroupId/transitiveMembers/microsoft.graph.user?`$select=$select&`$top=999"
    Write-ScopeLog -Message "Retrieving transitive user members for group $BackupPolicyScopeGroupId."
    $members = @(Invoke-GraphGetCollection -Uri $membersUri)
    Write-ScopeLog -Message "Group user members retrieved: $($members.Count)"

    $mailboxIndex = Build-MailboxAddressIndex -Mailboxes $mailboxes
    $memberRows = @($members | ForEach-Object { ConvertTo-GroupMemberRow -Member $_ -RunId $runId -RunDateUtc $runDateUtc -TenantName $OrgDomain -GroupId $BackupPolicyScopeGroupId -GroupDisplayName $BackupPolicyScopeGroupDisplayName })
    $coverageRows = @(ConvertTo-CoverageRows -Members $members -MailboxIndex $mailboxIndex -RunId $runId -RunDateUtc $runDateUtc -TenantName $OrgDomain -GroupId $BackupPolicyScopeGroupId -GroupDisplayName $BackupPolicyScopeGroupDisplayName)
    $missingRows = @($coverageRows | Where-Object { $_.MailboxFoundInInventory -ne $true })
    $summaryRows = @(
        [pscustomobject]@{ RunId=$runId; RunDateUtc=$runDateUtc; TenantName=$OrgDomain; GroupId=$BackupPolicyScopeGroupId; GroupDisplayName=$BackupPolicyScopeGroupDisplayName; Metric='GroupUserMembers'; NumericValue=$members.Count; TextValue=[string]$members.Count; Threshold=''; Status='OK'; Details='Transitive user members in expected backup policy scope group.' },
        [pscustomobject]@{ RunId=$runId; RunDateUtc=$runDateUtc; TenantName=$OrgDomain; GroupId=$BackupPolicyScopeGroupId; GroupDisplayName=$BackupPolicyScopeGroupDisplayName; Metric='MatchedMailboxes'; NumericValue=(@($coverageRows | Where-Object { $_.MailboxFoundInInventory -eq $true })).Count; TextValue=[string](@($coverageRows | Where-Object { $_.MailboxFoundInInventory -eq $true })).Count; Threshold=''; Status='OK'; Details='Group users matched to Exchange Online mailbox inventory.' },
        [pscustomobject]@{ RunId=$runId; RunDateUtc=$runDateUtc; TenantName=$OrgDomain; GroupId=$BackupPolicyScopeGroupId; GroupDisplayName=$BackupPolicyScopeGroupDisplayName; Metric='MembersWithoutMailboxInInventory'; NumericValue=$missingRows.Count; TextValue=[string]$missingRows.Count; Threshold='0'; Status=$(if ($missingRows.Count -gt 0) { 'Warning' } else { 'OK' }); Details='Group users not matched to Exchange Online mailbox inventory.' }
    )

    Export-InventoryCsv -BaseName 'M365_BackupPolicyScope_GroupMembers' -Rows $memberRows
    Export-InventoryCsv -BaseName 'M365_BackupPolicyScope_MailboxCoverage' -Rows $coverageRows
    Export-InventoryCsv -BaseName 'M365_BackupPolicyScope_GroupMembersWithoutMailbox' -Rows $missingRows
    Export-InventoryCsv -BaseName 'M365_BackupPolicyScope_Summary' -Rows $summaryRows

    Write-ScopeLog -Message ("Backup policy scope inventory complete. Members={0}; MatchedMailboxes={1}; MissingMailbox={2}" -f $members.Count, (@($coverageRows | Where-Object { $_.MailboxFoundInInventory -eq $true })).Count, $missingRows.Count) -Level SUCCESS
    exit 0
}
catch {
    Write-ScopeLog -Message $_.Exception.Message -Level ERROR
    exit 1
}
finally {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
}


# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBg/68OnKaqiwc0
# 6fYy73tLG+EYVTqa8lZumjl3VXd0HKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCDwHIg+Xzdy7Hx9EinlR1YgGy6wTLh6uDgSnelezYmYMTANBgkqhkiG9w0B
# AQEFAASCAYCzAuHATUf8A1X1EmY1Eu8vSljI2qL6Vfwt8nUNtVUvEO+1XkKocYgy
# n53/H4jWtHvexCddoDazzgNicNmOSFSK0weROCSyWYoicQIkLiY2HLsEUOEYw2oq
# p0Rll10L3a24sF7DTg77SxWd2lzH0S2cKTG1Fqinn/Ws8Fha0s/rEyoLCFgjL9g5
# A3hcBgDjGKYnDMHSYsnMacBh1cwM0oEsQdDdTuYHelJXuHbi02nHamdYVrypVfGs
# auX73WorxiEEMa2Y+3ObJ0tqU3uA/6SpE5c3lPAK96UKBXXqcGIXIyFZ19///g9z
# APY0mRL6UJeOJZRfYx4s9pngO+nQUEcH1ZdynvISykcwYGuDmFuXWL/1OgnfrwMm
# NHJt7aM3hnxXd3whXlrrb/V7tnKY4/7xalF2qTPayXbDjBEdYbItGpXVMu4O5fDX
# Z5i5qsSifziTFGbdyZDHRBDB71pa01/+CEJktJ3hmoLXl4dQ6z5j6Ec7YswekvYK
# 3yY/YICtKJk=
# SIG # End signature block
