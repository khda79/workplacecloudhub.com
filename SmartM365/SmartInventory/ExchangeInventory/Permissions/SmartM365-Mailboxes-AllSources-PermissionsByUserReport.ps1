<#
.SYNOPSIS
    Builds a mailbox permissions report grouped by delegate user across Exchange Online and on-premises mailbox exports.

.DESCRIPTION
    Reads the latest SmartInventory Exchange mailbox CSV exports, normalizes FullAccess, SendAs, and SendOnBehalf
    delegates, then writes one row per resolved delegate with the mailboxes where permissions are granted. The
    report is read-only and highlights cross-premises permissions for migration and cleanup review.

.PARAMETER Tenant
    Tenant profile key to load from Config/Tenants. Defaults to test.

.PARAMETER LocalMailboxesCsvPath
    Optional override for the Exchange on-premises mailbox inventory CSV.

.PARAMETER ExoMailboxesCsvPath
    Optional override for the Exchange Online mailbox inventory CSV.

.PARAMETER OutputPath
    Optional output directory override. If omitted, ScriptCsvLogFolderPath from local JSON is used.

.REQUIREMENTS
    PowerShell 7+, SmartM365.Core module, read access to both source mailbox permission CSV files,
    and write access to configured DATA-ALL, DATA-LAST, and LOG-ALL folders. This generator does not
    query AD, EXO, Graph, or Exchange on-prem directly; data completeness depends on the supplied
    Exchange_OnPrem and Exchange_EXO mailbox permission exports.

.VERSION
1.11

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$LocalMailboxesCsvPath,
    [string]$ExoMailboxesCsvPath,
    [string]$OutputPath,
    [int]$MaxItems = 0
)
if ($PSBoundParameters.ContainsKey('MaxItems') -and $MaxItems -gt 0) {
    $global:SmartM365MaxItems = [int]$MaxItems
    $global:SmartM365TestMaxItems = [int]$MaxItems
    $global:SmartM365IsMaxItemsRun = $true
    foreach ($smartM365LimitName in @('TopUsers','TopMailboxes','MaxDevices','MaxSites','MaxTeams','MaxApps','MaxPolicies','Limit','MaxPages')) {
        $smartM365LimitVariable = Get-Variable -Name $smartM365LimitName -Scope Script -ErrorAction SilentlyContinue
        if ($smartM365LimitVariable -and -not $PSBoundParameters.ContainsKey($smartM365LimitName) -and $null -ne $smartM365LimitVariable.Value) {
            Set-Variable -Name $smartM365LimitName -Value ([int]$MaxItems) -Scope Script
        }
    }
}

$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidates = @((Join-Path $d 'SmartM365-TenantContext.ps1'), (Join-Path $d 'Config\SmartM365-TenantContext.ps1'))
        foreach ($p in $candidates) { if (Test-Path -LiteralPath $p) { return $p } }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $tenantContextPath
Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot | Out-Null

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ScriptVersion = "1.11"
$TaskName = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion"
$CurrentOperation = 'Initialize'

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host 'This script requires PowerShell 7 or later.' -ForegroundColor Red
    exit 1
}

function Get-ScriptLocalConfig {
    $configPath = Join-Path $PSScriptRoot ("{0}.local.json" -f [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))
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
    $script:SmartM365GlobalConfig = [pscustomobject]@{}
    $searchRoot = $PSScriptRoot
    while ($searchRoot) {
        $globalConfigPath = Join-Path $searchRoot 'Config\SmartM365.global.local.json'
        if (Test-Path -LiteralPath $globalConfigPath) {
            $script:SmartM365GlobalConfig = Get-Content -LiteralPath $globalConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            break
        }
        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }
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
            $tokenProperty = $globalConfig.PSObject.Properties[$match.Groups['Name'].Value]
            if ($null -eq $tokenProperty -or $null -eq $tokenProperty.Value) { continue }
            $tokenValue = Resolve-SmartM365ConfigValue -Value $tokenProperty.Value
            if ($null -eq $tokenValue) { continue }
            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }
        if (-not $changed) { break }
    }
    return $resolved
}

function Get-ScriptLocalConfigValue {
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$Name, $DefaultValue)
    $property = $Config.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        if ($property.Value -is [string]) {
            $value = $property.Value.Trim()
            if ($value -and $value -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) { return Resolve-SmartM365ConfigValue -Value $property.Value }
        }
        else { return Resolve-SmartM365ConfigValue -Value $property.Value }
    }
    $globalProperty = (Get-SmartM365GlobalConfig).PSObject.Properties[$Name]
    if ($null -ne $globalProperty -and $null -ne $globalProperty.Value) { return Resolve-SmartM365ConfigValue -Value $globalProperty.Value }
    return $DefaultValue
}

function Import-SmartM365CoreModule {
    $modulePath = & {
        $d = $PSScriptRoot
        while ($d) {
            foreach ($relative in @('Modules\SmartM365.Core\SmartM365.Core.psd1')) {
                $candidate = Join-Path $d $relative
                if (Test-Path -LiteralPath $candidate) { return $candidate }
            }
            $parent = Split-Path -Path $d -Parent
            if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
            $d = $parent
        }
        throw 'SmartM365.Core module not found.'
    }
    Import-Module -Name $modulePath -MinimumVersion '1.0.24' -Force -ErrorAction Stop
}

function Get-CsvPropertyValue {
    param(
        [Parameter(Mandatory)]$Row,
        [Parameter(Mandatory)][string[]]$Names,
        [string]$DefaultValue = ''
    )

    foreach ($name in $Names) {
        $property = $Row.PSObject.Properties[$name]
        if ($null -ne $property -and $null -ne $property.Value) {
            return [string]$property.Value
        }
    }
    return $DefaultValue
}
function Split-PermissionList {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return @() }
    return @([string]$Value -split '[;|]' | ForEach-Object { $_.Trim() } | Where-Object { $_ -and $_ -notin @('NotChecked', 'None', 'NT AUTHORITY\SELF') })
}

function Add-LookupValue {
    param([hashtable]$Map, [string]$Key, [hashtable]$Entry)
    if ([string]::IsNullOrWhiteSpace($Key)) { return }
    if (-not $Map.ContainsKey($Key)) { $Map[$Key] = $Entry }
}

function Resolve-DelegateEntry {
    param([string]$Delegate)
    if ([string]::IsNullOrWhiteSpace($Delegate)) { return $null }
    $value = $Delegate.Trim()
    foreach ($map in @($script:IdentityToUser, $script:SmtpToUser, $script:DomainSamToUser, $script:SamToUser)) {
        if ($map.ContainsKey($value)) { return $map[$value] }
    }
    if ($value -match '^[^\\]+\\(?<Sam>.+)$') {
        $sam = $Matches['Sam']
        if ($script:SamToUser.ContainsKey($sam)) { return $script:SamToUser[$sam] }
    }
    return @{ Smtp = $value; DN = ''; ExtDirObjId = ''; Source = 'Unresolved' }
}

function Add-PermissionGrant {
    param(
        [Parameter(Mandatory)][string]$PermissionType,
        [Parameter(Mandatory)][string]$Delegate,
        [Parameter(Mandatory)][string]$TargetMailbox,
        [Parameter(Mandatory)][string]$TargetSource
    )
    $resolved = Resolve-DelegateEntry -Delegate $Delegate
    if ($null -eq $resolved) { return }
    $key = if (-not [string]::IsNullOrWhiteSpace($resolved.Smtp)) { $resolved.Smtp } else { $Delegate }
    if (-not $script:PermissionsReport.ContainsKey($key)) {
        $script:PermissionsReport[$key] = [pscustomobject]@{
            ResolvedUser = $key
            PrimarySMTPaddress = $resolved.Smtp
            DistinguishedName = $resolved.DN
            ExternalDirectoryObjectId = $resolved.ExtDirObjId
            Source = $resolved.Source
            FullAccessOn = [System.Collections.Generic.List[string]]::new()
            SendAsOn = [System.Collections.Generic.List[string]]::new()
            SendOnBehalfOn = [System.Collections.Generic.List[string]]::new()
        }
    }
    $script:PermissionsReport[$key].$PermissionType.Add("[$TargetSource] $TargetMailbox") | Out-Null
}

function Send-PermissionsTeamsInfo {
    param([string]$Message, [hashtable]$Facts)
    if (Get-Command Send-SmartM365TeamsNotification -ErrorAction SilentlyContinue) {
        Send-SmartM365TeamsNotification -Title 'Mailbox permissions by user report success' -Message $Message -Level SUCCESS -Channel Infos -ResultSummary $Message -Facts $Facts | Out-Null
    }
}

$config = Get-ScriptLocalConfig
$global:RetentionMaxCSV = [int](Get-ScriptLocalConfigValue -Config $config -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-ScriptLocalConfigValue -Config $config -Name 'RetentionMaxLogs' -DefaultValue 30)
$global:EnableSharePointUpload = [bool](Get-ScriptLocalConfigValue -Config $config -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = Get-ScriptLocalConfigValue -Config $config -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-ScriptLocalConfigValue -Config $config -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-ScriptLocalConfigValue -Config $config -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-ScriptLocalConfigValue -Config $config -Name 'SharePointTargetFolderPath' -DefaultValue ''
$LatestCsvFolderPath = Get-ScriptLocalConfigValue -Config $config -Name 'LatestCsvFolderPath' -DefaultValue ''
$ScriptCsvLogFolderPath = Get-ScriptLocalConfigValue -Config $config -Name 'ScriptCsvLogFolderPath' -DefaultValue ''
$LogAllRootPath = Get-ScriptLocalConfigValue -Config $config -Name 'LogAllRootPath' -DefaultValue ''
if (-not $LocalMailboxesCsvPath) { $LocalMailboxesCsvPath = Get-ScriptLocalConfigValue -Config $config -Name 'LocalMailboxesCsvPath' -DefaultValue '' }
if (-not $ExoMailboxesCsvPath) { $ExoMailboxesCsvPath = Get-ScriptLocalConfigValue -Config $config -Name 'ExoMailboxesCsvPath' -DefaultValue '' }
if (-not [string]::IsNullOrWhiteSpace($OutputPath)) { $ScriptCsvLogFolderPath = $OutputPath }
if ([string]::IsNullOrWhiteSpace($ScriptCsvLogFolderPath)) { $ScriptCsvLogFolderPath = Join-Path $PSScriptRoot 'Output' }
if ([string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) { $LatestCsvFolderPath = $ScriptCsvLogFolderPath }

Import-SmartM365CoreModule
$runId = [guid]::NewGuid().ToString()
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$logRoot = if ([string]::IsNullOrWhiteSpace($LogAllRootPath)) { Join-Path $ScriptCsvLogFolderPath 'Logs' } else { Join-Path $LogAllRootPath 'SmartM365-Exchange-MailboxPermissionsByUser' }
$logPath = Join-Path $logRoot ("SmartM365-Mailboxes-AllSources-PermissionsByUserReport_{0}.log" -f $timestamp)
foreach ($folder in @($ScriptCsvLogFolderPath, $LatestCsvFolderPath, $logRoot)) { if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null } }
Set-SmartM365CoreContext -RunId $runId -RunOutputRoot $ScriptCsvLogFolderPath -LatestOutputRoot $LatestCsvFolderPath -LogPath $logPath
$global:LogTextFile = $logPath

try {
    $CurrentOperation = 'InitializeScriptEnvironment'
    $initializedOutput = InitializeScriptEnvironment -OutputPath $ScriptCsvLogFolderPath -LogFileName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))
    $ScriptCsvLogFolderPath = $initializedOutput
    $timestampedCsvPath = Join-Path $ScriptCsvLogFolderPath ("Exchange_Mailboxes_AllSources_PermissionsByUser_{0}.csv" -f $timestamp)
    $latestCsvPath = Join-Path $LatestCsvFolderPath 'Exchange_Mailboxes_AllSources_PermissionsByUser.csv'
    Start-Transcript -Path $global:logTranscriptFile -Append | Out-Null
    WriteLog -Message "Starting $TaskName." -Level INFO
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($ScriptCsvLogFolderPath,$LatestCsvFolderPath,$logRoot) | Out-Null

    $CurrentOperation = 'LoadSourceCsv'
    $sourceRows = [System.Collections.Generic.List[object]]::new()
    if (-not [string]::IsNullOrWhiteSpace($LocalMailboxesCsvPath)) {
        $localMailboxRows = @(Import-SmartM365CsvWithSharePointFallback -Path $LocalMailboxesCsvPath -Description 'Local mailbox CSV')
        if ($localMailboxRows.Count -gt 0) {
            foreach ($row in $localMailboxRows) {
                $sourceRows.Add([pscustomobject]@{
                    Source = 'Local'
                    Identity = Get-CsvPropertyValue -Row $row -Names @('Identity','Name','DisplayName')
                    PrimarySMTPaddress = Get-CsvPropertyValue -Row $row -Names @('PrimarySMTPaddress','PrimarySmtpAddress','PrimarySMTPAddress','PrimaryEmailAddress')
                    DistinguishedName = Get-CsvPropertyValue -Row $row -Names @('DistinguishedName')
                    ExternalDirectoryObjectId = Get-CsvPropertyValue -Row $row -Names @('ExternalDirectoryObjectId','ExternalDirectoryObjectID','Guid')
                    DomainName = Get-CsvPropertyValue -Row $row -Names @('DomainName','Domain')
                    SamAccountName = Get-CsvPropertyValue -Row $row -Names @('SamAccountName','SAMAccountName','Sam')
                    FullAccess = Get-CsvPropertyValue -Row $row -Names @('FullAccess','FullAccessPermissions')
                    SendAs = Get-CsvPropertyValue -Row $row -Names @('SendAs','SendAsPermissions')
                    GrantSendOnBehalfTo = Get-CsvPropertyValue -Row $row -Names @('GrantSendOnBehalfTo','SendOnBehalf','SendOnBehalfTo')
                }) | Out-Null
            }
        }
        else { throw "Local mailbox CSV not found locally or in SharePoint: $LocalMailboxesCsvPath" }
    }
    else { throw 'Local mailbox CSV path is not configured.' }

    if (-not [string]::IsNullOrWhiteSpace($ExoMailboxesCsvPath)) {
        $exoMailboxRows = @(Import-SmartM365CsvWithSharePointFallback -Path $ExoMailboxesCsvPath -Description 'EXO mailbox CSV')
        if ($exoMailboxRows.Count -gt 0) {
            foreach ($row in $exoMailboxRows) {
                $sourceRows.Add([pscustomobject]@{
                    Source = 'EXO'
                    Identity = Get-CsvPropertyValue -Row $row -Names @('UserPrincipalName','Identity','DisplayName')
                    PrimarySMTPaddress = Get-CsvPropertyValue -Row $row -Names @('PrimarySmtpAddress','PrimarySMTPaddress','PrimarySMTPAddress','PrimaryEmailAddress')
                    DistinguishedName = Get-CsvPropertyValue -Row $row -Names @('DistinguishedName')
                    ExternalDirectoryObjectId = Get-CsvPropertyValue -Row $row -Names @('ExternalDirectoryObjectId','ExternalDirectoryObjectID','ExternalDirectoryObjectGuid','Guid')
                    DomainName = Get-CsvPropertyValue -Row $row -Names @('DomainName','Domain')
                    SamAccountName = Get-CsvPropertyValue -Row $row -Names @('SamAccountName','SAMAccountName','Sam','Alias')
                    FullAccess = Get-CsvPropertyValue -Row $row -Names @('FullAccess','FullAccessPermissions')
                    SendAs = Get-CsvPropertyValue -Row $row -Names @('SendAs','SendAsPermissions')
                    GrantSendOnBehalfTo = Get-CsvPropertyValue -Row $row -Names @('GrantSendOnBehalfTo','SendOnBehalf','SendOnBehalfTo')
                }) | Out-Null
            }
        }
        else { throw "EXO mailbox CSV not found locally or in SharePoint: $ExoMailboxesCsvPath" }
    }
    else { throw 'EXO mailbox CSV path is not configured.' }

    if ($sourceRows.Count -eq 0) { throw 'No source mailbox rows were loaded. Check LocalMailboxesCsvPath and ExoMailboxesCsvPath.' }
    if ($MaxItems -gt 0 -and $sourceRows.Count -gt $MaxItems) {
        $originalSourceCount = $sourceRows.Count
        $limitedRows = @($sourceRows.ToArray() | Sort-Object Source, PrimarySMTPaddress, Identity | Select-Object -First $MaxItems)
        $sourceRows = [System.Collections.Generic.List[object]]::new()
        foreach ($limitedRow in $limitedRows) { $sourceRows.Add($limitedRow) | Out-Null }
        WriteLog -Message ("MaxItems enabled: restricted source mailbox rows from {0} to {1}." -f $originalSourceCount, $sourceRows.Count) -Level WARN
    }

    $script:DomainSamToUser = @{}
    $script:SamToUser = @{}
    $script:SmtpToUser = @{}
    $script:IdentityToUser = @{}
    $script:PermissionsReport = @{}

    foreach ($row in $sourceRows) {
        $entry = @{ Smtp = $row.PrimarySMTPaddress; DN = $row.DistinguishedName; ExtDirObjId = $row.ExternalDirectoryObjectId; Source = $row.Source }
        Add-LookupValue -Map $script:SmtpToUser -Key $row.PrimarySMTPaddress -Entry $entry
        Add-LookupValue -Map $script:IdentityToUser -Key $row.Identity -Entry $entry
        Add-LookupValue -Map $script:SamToUser -Key $row.SamAccountName -Entry $entry
        if (-not [string]::IsNullOrWhiteSpace($row.DomainName) -and -not [string]::IsNullOrWhiteSpace($row.SamAccountName)) {
            Add-LookupValue -Map $script:DomainSamToUser -Key ("{0}\{1}" -f $row.DomainName, $row.SamAccountName) -Entry $entry
        }
    }

    $CurrentOperation = 'AnalyzePermissions'
    foreach ($mb in $sourceRows) {
        $targetMailbox = if (-not [string]::IsNullOrWhiteSpace($mb.PrimarySMTPaddress)) { $mb.PrimarySMTPaddress } else { $mb.Identity }
        foreach ($delegate in Split-PermissionList $mb.FullAccess) { Add-PermissionGrant -PermissionType 'FullAccessOn' -Delegate $delegate -TargetMailbox $targetMailbox -TargetSource $mb.Source }
        foreach ($delegate in Split-PermissionList $mb.SendAs) { Add-PermissionGrant -PermissionType 'SendAsOn' -Delegate $delegate -TargetMailbox $targetMailbox -TargetSource $mb.Source }
        foreach ($delegate in Split-PermissionList $mb.GrantSendOnBehalfTo) { Add-PermissionGrant -PermissionType 'SendOnBehalfOn' -Delegate $delegate -TargetMailbox $targetMailbox -TargetSource $mb.Source }
    }

    $reportRows = @($script:PermissionsReport.Values | Sort-Object ResolvedUser | ForEach-Object {
        $delegateSource = $_.Source
        $sourcePattern = "\[$([regex]::Escape($delegateSource))\]"
        $full = @($_.FullAccessOn | Sort-Object -Unique)
        $sendAs = @($_.SendAsOn | Sort-Object -Unique)
        $sendOnBehalf = @($_.SendOnBehalfOn | Sort-Object -Unique)
        $crossFull = @($full | Where-Object { $_ -match '\[(Local|EXO)\]' -and $_ -notmatch $sourcePattern })
        $crossSendAs = @($sendAs | Where-Object { $_ -match '\[(Local|EXO)\]' -and $_ -notmatch $sourcePattern })
        $crossSendOnBehalf = @($sendOnBehalf | Where-Object { $_ -match '\[(Local|EXO)\]' -and $_ -notmatch $sourcePattern })
        [pscustomobject]@{
            ResolvedUser = $_.ResolvedUser
            PrimarySMTPaddress = $_.PrimarySMTPaddress
            DistinguishedName = $_.DistinguishedName
            ExternalDirectoryObjectId = $_.ExternalDirectoryObjectId
            Source = $delegateSource
            FullAccessOn = ($full -join '; ')
            SendAsOn = ($sendAs -join '; ')
            SendOnBehalfOn = ($sendOnBehalf -join '; ')
            CrossPremisesFullAccessOn = ($crossFull -join '; ')
            CrossPremisesSendAsOn = ($crossSendAs -join '; ')
            CrossPremisesSendOnBehalfOn = ($crossSendOnBehalf -join '; ')
            HasCrossPremisesPermissions = (($crossFull.Count + $crossSendAs.Count + $crossSendOnBehalf.Count) -gt 0)
            RunId = $runId
            ExportDateTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        }
    })

    Export-SmartM365Csv -Data $reportRows -TimestampedPath $timestampedCsvPath -LatestPath $latestCsvPath | Out-Null

    $message = "Mailbox permissions by user report completed. Source mailboxes: {0}; delegated users: {1}." -f $sourceRows.Count, $reportRows.Count
    WriteLog -Message $message -Level SUCCESS
    Send-PermissionsTeamsInfo -Message $message -Facts @{ Script = $MyInvocation.MyCommand.Name; SourceRows = $sourceRows.Count; DelegatedUsers = $reportRows.Count; LatestCsvPath = $latestCsvPath; LogFile = $global:LogTextFile }
    try { Stop-Transcript | Out-Null } catch {}
    try { RemoveOldFiles -Path $ScriptCsvLogFolderPath -Filter '*.csv' -KeepCount $global:RetentionMaxCSV -LogFile $global:LogTextFile } catch {}
    try { RemoveOldFiles -Path $logRoot -Filter '*.log' -KeepCount $global:RetentionMaxLogs -LogFile $global:LogTextFile } catch {}
    Complete-SmartM365ExecutionContext -Status Auto
}
catch {
    $globalError = $_
    WriteLog -Message ("Global error during {0}: {1}" -f $CurrentOperation, $globalError.Exception.Message) -Level ERROR
    try { Stop-Transcript | Out-Null } catch {}
    try { Complete-SmartM365ExecutionContext -Status Failed } catch {}
    exit 1
}
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDn6SiiPGQsfWyU
# mNeSDwyC3Ly3wcPjfE7bTIdpxsaXMqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
# s0Q4yPEDH+JoMA0GCSqGSIb3DQEBCwUAME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91
# ZGh1Yi5jb20wHhcNMjYwNzEzMDgyMjM1WhcNMjkwNzEzMDgzMjI5WjBOMR4wHAYD
# VQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRh
# Y3RAd29ya3BsYWNlY2xvdWRodWIuY29tMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEAse6XztERSyHn9DVqj8Rdv0qjc5owqvgAIGaYxBmfiQuoM48Fo4Xt
# 1ovi9brLUtf55G4XgthNPCoanxfCRRg30IVRxaDfdPXJzYmgsM5tXlsuNU49lE7E
# PJk3+jEOgSCt8NKzmVPKpNRG0NmK0a8wm12cceYZOZlSYE0+ZtT6wy5PQQjMUqIx
# XnGjt4H0nfgZZa7D4FyARKOVg/Xr9sUq5jIn3zszvg4jjeb4b0DKJtfbHukhWc2Y
# oVFgswxVBXCWIaBnfF/cjqMfK/CaToT2trVb4hG4qcQ31s1nR4keoRaOw/vyd6ap
# rEtCsT22N/Jx0dz7fIo1tVyvIaVcHdN9LW3chn0en0OKZ6Ke1OH9wf2prl4KA6Ww
# VzrAZrOlXTAItdK7D9kKO/HeJd4PZvO53oy1LdmMGLSz3OLB9e5q7yo8rfqi5Ka9
# KzM2CrSzz1yphn/H90wz7Q2pm4FIlWdcj86A/0kmhYg+5Wqqbg1drrPXu4nEBwWN
# /dzoGtKZKHTdAgMBAAGjgZYwgZMwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMD8GA1UdEQQ4MDaBHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIu
# Y29tghV3b3JrcGxhY2VjbG91ZGh1Yi5jb20wDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQUXIOOADQM78XfPAncirgCECedg9gwDQYJKoZIhvcNAQELBQADggGBADhZUB2R
# 5J/Jw030xodhEWeCQ0vnJRaiEsjOxuArQREKH3lCrQ3UsUVl292d6LnQUSTH/jF7
# rovEZ+JN2GQ/LCrXRaCuwCEGZKzlSEbtYWhfwDyj6GpIPq8Y4SeXyjdq4/rrI1bm
# iTK4Sq7EoBlGJuX6l2nfvx1tTioSr11FoDfllJR7EYawRj9hBFJ0gG0b2SuYZMgW
# gaDKefcnJDmOwcRNAZUII0ss8EeyANukWSkNN5ILZ+iKDpQgZxgDLPTiRguCyx45
# PI5wrVTjV/pR7IrtSIfq8UladlrSZJyyDn3NV2ATvIZ6wNxbTmPFcE0uMg/EYzwd
# Tek+CgXL3TxUKeldJM4YDWPimNBRhOPXzBDiOQIj6WNswt/KM1oDLnA00CNtciPN
# dn+dXlneMvTEUah9wyt8o8tkLpoBw+KN+Bq/K0O1qPtS7umi70l45pPiej+mwbwq
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIMFOFEn5o4OGQZmoHRsi1bkdma2jYISLAVtrQS/+diIzMA0GCSqG
# SIb3DQEBAQUABIIBgGXriO0znyuOINud5dlrnANkNTEQLoUSoN3kYkPl39DmTYFh
# VURMuVUucU8JXWarCEuANVz9ylVTSFrkmzHE2LhTlkbEViUBw6nA8HaGHMyaPsqV
# 6PpeheOOB0oqWTPJUO0zI97/ran3nHKgfR3VqHc+/DvEdLtSZgWCuIWvBCsAAWWf
# Y5CIiYIRF2QI0W3PE17v+SNXDKfq7SV1G6dl6jxdjc/5xx9RON55sSySHcpDFF8V
# DhyFC1w1Bl7S3zVffn8H+54IRzkZFM77orBIiddqJnNnkmz+h2rYQqmaPFB0xAiV
# jHwDuJ1E4zpdxZcb8k7hra9N9QK9+4sXvxj7mYgV64G3uNZujzC3THxQKY2uob8Z
# hZMlAddNBxNKwMweO6yxUOcnROcRTf7xHnNXH9J3u9yeiFH1QfoKjkImmizokDEr
# mUxxA59cJAINxsEWcWx9gxY1uQyTS+xuNm/KR4sp0RcqFsjL5D5TuuU5InTG/cf+
# n9fUHaOYgrkMNsJT3aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjAyMTA5
# MzJaMC8GCSqGSIb3DQEJBDEiBCBkY0RoMrqRAopCIQZJHZQN1U/bdSbFE8qJzk3b
# ai0vhzANBgkqhkiG9w0BAQEFAASCAgCbpLuDl1KtcUzvbL1GSozFA3dH+LxGXlGw
# 6e2quBGlC6HF99B9gDxE2qIWprqBSDCaxQeILJcEs1hXiy1lsN6JzWz/zGIcuxNf
# mVdldIY+yk1kOtzrLtYZEKPkkkBQ0+5QicyvdSIXu3kB0ctxYOw17vsQi4o+An7W
# OAQBofC5oosGTCFS/VCB++zKVaWL2apYLmfVn3H1u2X0O13/iIdxNq6ekXaDEHiK
# kmh40TWd9vWQyp2CVFzWEkW7qn0rJVtU+kY3xdUCJ7dINrKePa2r+HjVNgrOTPin
# WklcsRUAC/hQxvwmO+IiKtZf06+T9l5ifG8/Lq8mAotP5Ji5kuUft5cCIQ/psfHf
# GzBc0SKtGCttz97hHIt1hB5zaoZeer/t0/mEgd6LZHv/uuYqJ8eT4FwK1EliUKc9
# 50bK9uf1qoKBDNLGQ4FjkbVIvfb1uttXR1LwGJAANj6RCb6+QvKlKQzya65ecRTR
# 4FsTSfhniIdwikhDMkVEncfV3YZnp3Hd7hmCGcSaRiqEJUdr79qv0pRmAWomD64Z
# Qop12bYbqoI31iQ0HTiCRAAClYB+y/zr42fJ61qUXUz3WgokAK0hYMI5Awk6z2fS
# SqM8ElYlzYIQBT1thHqWAp0dUXZmuP8EcscYMWwNAFtYzh6HejNiIUJIXdBClm74
# EZJlLlF5RA==
# SIG # End signature block
