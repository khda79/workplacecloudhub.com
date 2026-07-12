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
1.8

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
$ScriptVersion = "1.8"
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
    Import-Module -Name $modulePath -MinimumVersion '1.0.23' -Force -ErrorAction Stop
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
$logRoot = if ([string]::IsNullOrWhiteSpace($LogAllRootPath)) { Join-Path $ScriptCsvLogFolderPath 'Logs' } else { Join-Path $LogAllRootPath 'Exchange-MailboxPermissionsByUser' }
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
    Complete-SmartM365ExecutionContext -Status Auto
    try { RemoveOldFiles -Path $ScriptCsvLogFolderPath -Filter '*.csv' -KeepCount $global:RetentionMaxCSV -LogFile $global:LogTextFile } catch {}
    try { RemoveOldFiles -Path $logRoot -Filter '*.log' -KeepCount $global:RetentionMaxLogs -LogFile $global:LogTextFile } catch {}
}
catch {
    $globalError = $_
    WriteLog -Message ("Global error during {0}: {1}" -f $CurrentOperation, $globalError.Exception.Message) -Level ERROR
    try { Stop-Transcript | Out-Null } catch {}
    try { Complete-SmartM365ExecutionContext -Status Auto } catch {}
    exit 1
}
# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCCqs6TtxCPpzqI
# 90lm+yjkPzaLchu7xKu8+BMJ8j3UPKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDa5LYyWSDNczMeqQWyuBHCifHVdYGaOUHVVfacCR2xQDANBgkqhkiG9w0B
# AQEFAASCAYCF81faYxfjPQCvSiZPko/GDY/O7hKDfFm/4aWt6tNGb95xKO1JRJVW
# QC3sGaN5x7LbfwNMD1+oD9p74q1a+/hmbwvkXqdoUH6mjVfPY8VZhi9CiHpKKT0X
# MM/XzfRaLAOG+6OV0jgDNUrFr2xl41TeLMC+pzofB+gvpjoa6oNZ+pbf4AIaDk0r
# KCOLPu3sDk7PEW0E06rplpzXwasDVGdLUbQ9W5eUJK/uSQEoFXdRsUOLtOcUYrm7
# 1DOqK+VQOj9xZhrXY1z/IlgfRf/+vcdU+LoxTce/D18Dx0q87torO/zZvEVE0XnW
# VTXYDE5vBvJGoaU4FXJa2IJvI4a26wmm5BgNjknQ+GoBdXeCOXSp4NttDboxDFj4
# +nNkSxovNh2SUKT/Vl6kqQh3jEQldlJRf8gxGI5SbxD1VC+lW48BNuOoQZ2gxDPq
# 6XoJ9xMrgk3e3n5gUbulMkTo7IIA2P/ttFdLmvFileZclh1xMZQpjTBEuCsfS5vi
# SQ5yY3SCEoE=
# SIG # End signature block
