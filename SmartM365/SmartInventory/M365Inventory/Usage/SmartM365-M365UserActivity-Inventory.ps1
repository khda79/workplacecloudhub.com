<#
.SYNOPSIS
Exports Microsoft 365 user and workload usage details from Microsoft Graph reports.

.DESCRIPTION
Downloads selected Microsoft 365 usage reports, normalizes the active user detail
report, and publishes stable CSV files into the tenant DATA-LAST folder for
SmartFinOps and downstream inventory analysis.

.VERSION
1.7


.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication; Microsoft.Graph.Reports.
    Minimum Graph application permissions: Reports.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
Author: https://github.com/khda79/workplacecloudhub.com
Requires: PowerShell 7+, Microsoft.Graph.Reports, SmartM365.Core.psd1
Minimum application permissions: Reports.Read.All
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [ValidateSet('D7', 'D30', 'D90', 'D180')]
    [string]$Period = 'D180',
    [ValidateSet('All', 'Office365ActiveUserDetail', 'MailboxUsageDetail', 'OneDriveUsageAccountDetail', 'SharePointSiteUsageDetail', 'Office365ActivationUserDetail', 'TeamsUserActivityUserDetail', 'EmailActivityUserDetail')]
    [string[]]$Reports = @('Office365ActiveUserDetail'),
    [string]$OutputPath,
    [string]$LatestCsvFolderPath,
    [switch]$Connect,
    [switch]$InteractiveAuth,
    [switch]$ValidateOnly,
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

$ErrorActionPreference = 'Stop'
$ScriptVersion = "1.7"
$runId = Get-Date -Format 'yyyyMMdd_HHmmss'

if ($PSVersionTable.PSVersion.Major -lt 7) {
    throw "This script requires PowerShell 7 or later. Current version: $($PSVersionTable.PSVersion)"
}

$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidates = @(
            (Join-Path -Path $d -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $d -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )
        foreach ($candidate in $candidates) {
            if (Test-Path -LiteralPath $candidate) { return $candidate }
        }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $tenantContextPath
$script:SmartM365EffectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot

function Resolve-SmartM365TokenValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) { return $Value }
    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $tokenMatches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($tokenMatches.Count -eq 0) { break }
        $changed = $false
        foreach ($match in $tokenMatches) {
            $property = $script:SmartM365EffectiveConfig.PSObject.Properties[$match.Groups['Name'].Value]
            if ($null -eq $property -or $null -eq $property.Value) { continue }
            $tokenValue = Resolve-SmartM365TokenValue -Value $property.Value
            if ($null -eq $tokenValue) { continue }
            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }
        if (-not $changed) { break }
    }
    return $resolved
}

function Get-SmartM365ConfigValue {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name, [AllowNull()]$DefaultValue = $null)

    $property = $script:SmartM365EffectiveConfig.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $DefaultValue }
    return Resolve-SmartM365TokenValue -Value $property.Value
}

function ConvertTo-DateOrNull {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return $null }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    try { return [datetime]$text } catch { return $null }
}

function Get-SourcePropertyValue {
    [CmdletBinding()]
    param([AllowNull()]$Row, [Parameter(Mandatory)][string[]]$Names)

    if ($null -eq $Row) { return $null }
    foreach ($name in $Names) {
        $property = $Row.PSObject.Properties[$name]
        if ($null -ne $property) { return $property.Value }
    }
    return $null
}

function ConvertTo-ReportBool {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return $false }
    $text = ([string]$Value).Trim()
    return ($text -match '^(true|yes|1)$')
}

function Get-LatestActivity {
    [CmdletBinding()]
    param([hashtable]$ActivityByWorkload)

    $latestDate = $null
    $workloads = New-Object System.Collections.Generic.List[string]
    foreach ($key in $ActivityByWorkload.Keys) {
        $date = $ActivityByWorkload[$key]
        if ($null -eq $date) { continue }
        if ($null -eq $latestDate -or $date -gt $latestDate) {
            $latestDate = $date
            $workloads.Clear()
            $workloads.Add($key) | Out-Null
        }
        elseif ($date -eq $latestDate) {
            $workloads.Add($key) | Out-Null
        }
    }

    return [pscustomobject]@{
        Date = $latestDate
        Workload = ($workloads.ToArray() -join ';')
    }
}

function ConvertFrom-M365UserActivityReport {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows)

    foreach ($row in $Rows) {
        $activityByWorkload = @{
            Exchange = ConvertTo-DateOrNull (Get-SourcePropertyValue -Row $row -Names @('Exchange Last Activity Date'))
            OneDrive = ConvertTo-DateOrNull (Get-SourcePropertyValue -Row $row -Names @('OneDrive Last Activity Date'))
            SharePoint = ConvertTo-DateOrNull (Get-SourcePropertyValue -Row $row -Names @('SharePoint Last Activity Date'))
            SkypeForBusiness = ConvertTo-DateOrNull (Get-SourcePropertyValue -Row $row -Names @('Skype For Business Last Activity Date', 'Skype For Business Last Activity date'))
            Yammer = ConvertTo-DateOrNull (Get-SourcePropertyValue -Row $row -Names @('Yammer Last Activity Date'))
            Teams = ConvertTo-DateOrNull (Get-SourcePropertyValue -Row $row -Names @('Teams Last Activity Date'))
        }
        $latest = Get-LatestActivity -ActivityByWorkload $activityByWorkload

        [pscustomobject]@{
            RunId = $runId
            ReportPeriod = $Period
            ReportRefreshDate = [string](Get-SourcePropertyValue -Row $row -Names @('Report Refresh Date'))
            UserPrincipalName = [string](Get-SourcePropertyValue -Row $row -Names @('User Principal Name', 'UserPrincipalName'))
            DisplayName = [string](Get-SourcePropertyValue -Row $row -Names @('Display Name', 'DisplayName'))
            IsDeleted = ConvertTo-ReportBool (Get-SourcePropertyValue -Row $row -Names @('Is Deleted'))
            DeletedDate = [string](Get-SourcePropertyValue -Row $row -Names @('Deleted Date'))
            HasExchangeLicense = ConvertTo-ReportBool (Get-SourcePropertyValue -Row $row -Names @('Has Exchange License'))
            HasOneDriveLicense = ConvertTo-ReportBool (Get-SourcePropertyValue -Row $row -Names @('Has OneDrive License'))
            HasSharePointLicense = ConvertTo-ReportBool (Get-SourcePropertyValue -Row $row -Names @('Has SharePoint License'))
            HasSkypeForBusinessLicense = ConvertTo-ReportBool (Get-SourcePropertyValue -Row $row -Names @('Has Skype For Business License'))
            HasYammerLicense = ConvertTo-ReportBool (Get-SourcePropertyValue -Row $row -Names @('Has Yammer License'))
            HasTeamsLicense = ConvertTo-ReportBool (Get-SourcePropertyValue -Row $row -Names @('Has Teams License'))
            ExchangeLastActivityDate = if ($activityByWorkload.Exchange) { $activityByWorkload.Exchange.ToString('yyyy-MM-dd') } else { '' }
            OneDriveLastActivityDate = if ($activityByWorkload.OneDrive) { $activityByWorkload.OneDrive.ToString('yyyy-MM-dd') } else { '' }
            SharePointLastActivityDate = if ($activityByWorkload.SharePoint) { $activityByWorkload.SharePoint.ToString('yyyy-MM-dd') } else { '' }
            SkypeForBusinessLastActivityDate = if ($activityByWorkload.SkypeForBusiness) { $activityByWorkload.SkypeForBusiness.ToString('yyyy-MM-dd') } else { '' }
            YammerLastActivityDate = if ($activityByWorkload.Yammer) { $activityByWorkload.Yammer.ToString('yyyy-MM-dd') } else { '' }
            TeamsLastActivityDate = if ($activityByWorkload.Teams) { $activityByWorkload.Teams.ToString('yyyy-MM-dd') } else { '' }
            LastActivityDate = if ($latest.Date) { $latest.Date.ToString('yyyy-MM-dd') } else { '' }
            LastActivityWorkload = $latest.Workload
            DaysSinceLastActivity = if ($latest.Date) { [int]((Get-Date).Date - $latest.Date.Date).TotalDays } else { '' }
            HasAnyM365Activity = ($null -ne $latest.Date)
            AssignedProducts = [string](Get-SourcePropertyValue -Row $row -Names @('Assigned Products'))
        }
    }
}

function Get-M365UsageReportDefinition {
    [CmdletBinding()]
    param()

    [ordered]@{
        Office365ActiveUserDetail = [pscustomobject]@{
            Name = 'Office365ActiveUserDetail'
            Command = 'Get-MgReportOffice365ActiveUserDetail'
            Endpoint = 'getOffice365ActiveUserDetail'
            BaseFileName = 'M365_Users_Activity'
            SupportsPeriod = $true
            Normalize = 'ActiveUserDetail'
            Description = 'Microsoft 365 active user detail by workload.'
        }
        MailboxUsageDetail = [pscustomobject]@{
            Name = 'MailboxUsageDetail'
            Command = 'Get-MgReportMailboxUsageDetail'
            Endpoint = 'getMailboxUsageDetail'
            BaseFileName = 'M365_Mailbox_Usage'
            SupportsPeriod = $true
            Normalize = 'Generic'
            Description = 'Mailbox usage, storage, quotas, deleted items, and archive flag.'
        }
        OneDriveUsageAccountDetail = [pscustomobject]@{
            Name = 'OneDriveUsageAccountDetail'
            Command = 'Get-MgReportOneDriveUsageAccountDetail'
            Endpoint = 'getOneDriveUsageAccountDetail'
            BaseFileName = 'M365_OneDrive_Usage'
            SupportsPeriod = $true
            Normalize = 'Generic'
            Description = 'OneDrive usage by account, storage, file counts, and owner.'
        }
        SharePointSiteUsageDetail = [pscustomobject]@{
            Name = 'SharePointSiteUsageDetail'
            Command = 'Get-MgReportSharePointSiteUsageDetail'
            Endpoint = 'getSharePointSiteUsageDetail'
            BaseFileName = 'M365_SharePoint_SiteUsage'
            SupportsPeriod = $true
            Normalize = 'Generic'
            Description = 'SharePoint site usage, storage, file counts, and activity.'
        }
        Office365ActivationUserDetail = [pscustomobject]@{
            Name = 'Office365ActivationUserDetail'
            Command = 'Get-MgReportOffice365ActivationUserDetail'
            Endpoint = 'getOffice365ActivationsUserDetail'
            BaseFileName = 'M365_Apps_Activations'
            SupportsPeriod = $false
            Normalize = 'Generic'
            Description = 'Microsoft 365 Apps / Office activations by user and platform.'
        }
        TeamsUserActivityUserDetail = [pscustomobject]@{
            Name = 'TeamsUserActivityUserDetail'
            Command = 'Get-MgReportTeamUserActivityUserDetail'
            Endpoint = 'getTeamsUserActivityUserDetail'
            BaseFileName = 'M365_Teams_UserActivity'
            SupportsPeriod = $true
            Normalize = 'Generic'
            Description = 'Teams user activity detail.'
        }
        EmailActivityUserDetail = [pscustomobject]@{
            Name = 'EmailActivityUserDetail'
            Command = 'Get-MgReportEmailActivityUserDetail'
            Endpoint = 'getEmailActivityUserDetail'
            BaseFileName = 'M365_Email_Activity'
            SupportsPeriod = $true
            Normalize = 'Generic'
            Description = 'Email activity user detail, including send/read/receive counts.'
        }
    }
}

function Resolve-M365UsageReportSelection {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$RequestedReports)

    $definitions = Get-M365UsageReportDefinition
    if ($RequestedReports -contains 'All') { return @($definitions.Values) }

    $selected = New-Object System.Collections.Generic.List[object]
    foreach ($reportName in $RequestedReports) {
        if (-not $definitions.Contains($reportName)) { throw "Unsupported report selection: $reportName" }
        $selected.Add($definitions[$reportName]) | Out-Null
    }
    return @($selected.ToArray() | Sort-Object Name -Unique)
}

function ConvertFrom-M365GenericReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][string]$ReportName
    )

    foreach ($row in $Rows) {
        $ordered = [ordered]@{
            RunId = $runId
            ReportName = $ReportName
            ReportPeriodRequested = $Period
        }
        foreach ($property in $row.PSObject.Properties) {
            if (-not $ordered.Contains($property.Name)) { $ordered[$property.Name] = $property.Value }
        }
        [pscustomobject]$ordered
    }
}

function Invoke-M365UsageReportDownload {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$ReportDefinition,
        [Parameter(Mandatory)][string]$RawPath
    )

    $endpoint = [string]$ReportDefinition.Endpoint
    if ([string]::IsNullOrWhiteSpace($endpoint)) {
        throw "Missing Graph reports endpoint for $($ReportDefinition.Name)."
    }

    if ($ReportDefinition.SupportsPeriod) {
        $uri = "https://graph.microsoft.com/v1.0/reports/$endpoint(period='$Period')"
    }
    else {
        $uri = "https://graph.microsoft.com/v1.0/reports/$endpoint"
    }

    Invoke-MgGraphRequest -Method GET -Uri $uri -OutputFilePath $RawPath -ProgressAction SilentlyContinue -ErrorAction Stop | Out-Null
}

function Export-M365UsageReport {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$ReportDefinition)

    $rawPath = Join-Path -Path $runOutputRoot -ChildPath ("{0}_Raw_{1}_{2}.csv" -f $ReportDefinition.BaseFileName, $Period, $runId)
    $effectivePeriod = if ($ReportDefinition.SupportsPeriod) { $Period } else { 'not applicable' }
    WriteLog ("Downloading {0}. Period={1}" -f $ReportDefinition.Name, $effectivePeriod) 'INFO'
    Invoke-M365UsageReportDownload -ReportDefinition $ReportDefinition -RawPath $rawPath

    $rawRows = @(Import-Csv -LiteralPath $rawPath)
    if ($ReportDefinition.Normalize -eq 'ActiveUserDetail') {
        $normalizedRows = @(ConvertFrom-M365UserActivityReport -Rows $rawRows)
        $columns = @(
            'RunId', 'ReportPeriod', 'ReportRefreshDate', 'UserPrincipalName', 'DisplayName', 'IsDeleted', 'DeletedDate',
            'HasExchangeLicense', 'HasOneDriveLicense', 'HasSharePointLicense', 'HasSkypeForBusinessLicense', 'HasYammerLicense', 'HasTeamsLicense',
            'ExchangeLastActivityDate', 'OneDriveLastActivityDate', 'SharePointLastActivityDate', 'SkypeForBusinessLastActivityDate', 'YammerLastActivityDate', 'TeamsLastActivityDate',
            'LastActivityDate', 'LastActivityWorkload', 'DaysSinceLastActivity', 'HasAnyM365Activity', 'AssignedProducts'
        )
    }
    else {
        $normalizedRows = @(ConvertFrom-M365GenericReport -Rows $rawRows -ReportName $ReportDefinition.Name)
        $columns = @()
    }

    $exportResult = Export-SmartM365Csv `
        -BaseFileName $ReportDefinition.BaseFileName `
        -OutputPath $runOutputRoot `
        -GlobalPath $LatestCsvFolderPath `
        -Data $normalizedRows `
        -Columns $columns

    [pscustomobject]@{
        ReportName = $ReportDefinition.Name
        Rows = $normalizedRows.Count
        LatestPath = $exportResult.LatestPath
        RawPath = $rawPath
    }
}

$dataAllRoot = Resolve-SmartM365TokenValue -Value (Get-SmartM365ConfigValue -Name 'DataAllRootPath' -DefaultValue '')
$logAllRoot = Resolve-SmartM365TokenValue -Value (Get-SmartM365ConfigValue -Name 'LogAllRootPath' -DefaultValue '')
if ([string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) {
    $LatestCsvFolderPath = Resolve-SmartM365TokenValue -Value (Get-SmartM365ConfigValue -Name 'LatestCsvFolderPath' -DefaultValue '')
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path -Path $dataAllRoot -ChildPath 'M365\Usage'
}
$runOutputRoot = Join-Path -Path $OutputPath -ChildPath $runId
$logFolder = Join-Path -Path $logAllRoot -ChildPath 'SmartM365-M365UserActivity-Inventory'
$logPath = Join-Path -Path $logFolder -ChildPath ("SmartM365-M365UserActivity-Inventory_{0}.log" -f $runId)

$modulePath = Join-Path -Path ([string](Get-SmartM365ConfigValue -Name 'SmartM365RootPath' -DefaultValue (Split-Path -Path $PSScriptRoot -Parent))) -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
Import-Module -Name $modulePath -MinimumVersion '1.0.24' -Force -ErrorAction Stop

$global:RetentionMaxCSV = [int](Get-SmartM365ConfigValue -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-SmartM365ConfigValue -Name 'RetentionMaxLogs' -DefaultValue 30)
$global:EnableSharePointUpload = [bool](Get-SmartM365ConfigValue -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = Get-SmartM365ConfigValue -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-SmartM365ConfigValue -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-SmartM365ConfigValue -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-SmartM365ConfigValue -Name 'SharePointTargetFolderPath' -DefaultValue ''

Set-SmartM365CoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestCsvFolderPath -LogPath $logPath
$global:LogTextFile = $logPath
$global:logTextFile = $logPath
$global:LogPath = $logFolder
$global:logTranscriptFile = Join-Path -Path $logFolder -ChildPath ("SmartM365-M365UserActivity-Inventory_{0}_Transcript.log" -f $runId)

function Stop-SmartM365UsageTranscript {
    [CmdletBinding()]
    param()
    try {
        Stop-Transcript | Out-Null
        if ($global:logTranscriptFile -and (Get-Command Update-SmartM365TimestampedTranscript -ErrorAction SilentlyContinue)) {
            Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile
        }
    }
    catch {}
}
$selectedReports = Resolve-M365UsageReportSelection -RequestedReports $Reports
$selectedReportCommands = @($selectedReports | ForEach-Object { $_.Command })
$requiredReportCommands = @(($selectedReportCommands + @('Invoke-MgGraphRequest')) | Sort-Object -Unique)

try {
    foreach ($folder in @($runOutputRoot, $LatestCsvFolderPath, $logFolder)) {
        if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
    }
    Start-Transcript -Path $global:logTranscriptFile -Append | Out-Null

    Invoke-SmartM365Preflight `
        -ScriptName 'SmartM365-M365UserActivity-Inventory' `
        -RequiredModules @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Reports') `
        -RequiredCommands $requiredReportCommands `
        -RequiredGraphApplicationPermissions @('Reports.Read.All') `
        -OutputPaths @($runOutputRoot, $LatestCsvFolderPath) | Out-Null

    if ($ValidateOnly) {
        WriteLog "Validation completed. Tenant=$Tenant; Period=$Period; Reports=$(($selectedReports.Name -join ',')); OutputPath=$runOutputRoot; LatestCsvFolderPath=$LatestCsvFolderPath" 'SUCCESS'
        return
    }

    if ($Connect -or $null -eq (Get-MgContext -ErrorAction SilentlyContinue)) {
        $connectParams = @{
            Graph = $true
            ExchangeOnline = $false
            GraphScopes = @('Reports.Read.All')
        }
        if (-not $InteractiveAuth) {
            $connectParams.AppId = [string](Get-SmartM365ConfigValue -Name 'AppId' -DefaultValue '')
            $connectParams.Thumbprint = [string](Get-SmartM365ConfigValue -Name 'Thumbprint' -DefaultValue '')
            $connectParams.TenantId = [string](Get-SmartM365ConfigValue -Name 'TenantId' -DefaultValue '')
        }

        $connectResult = Connect-SmartM365CloudSession @connectParams
        if (-not $connectResult.GraphConnected) {
            throw 'Microsoft Graph connection failed. Check app-only certificate settings or use -InteractiveAuth.'
        }
    }

    $reportResults = New-Object System.Collections.Generic.List[object]
    foreach ($reportDefinition in $selectedReports) {
        $reportResults.Add((Export-M365UsageReport -ReportDefinition $reportDefinition)) | Out-Null
    }

    $summaryParts = @($reportResults | ForEach-Object { "{0}={1}" -f $_.ReportName, $_.Rows })
    $summary = "Tenant=$Tenant; Period=$Period; Reports=$($summaryParts -join '; '); LatestFolder=$LatestCsvFolderPath"
    WriteLog ("Microsoft 365 usage reports inventory completed. {0}" -f $summary) 'SUCCESS'
    Send-SmartM365TeamsNotification -Level SUCCESS -Channel Infos -Title 'SmartM365 M365 usage reports inventory completed' -Message $summary -ResultSummary $summary -Facts @{
        Tenant = $Tenant
        Period = $Period
        Reports = ($selectedReports.Name -join ',')
        LatestFolder = $LatestCsvFolderPath
        RunId = $runId
    } | Out-Null
    Stop-SmartM365UsageTranscript
    Complete-SmartM365ExecutionContext -Status Auto
}
catch {
    $message = $_.Exception.Message
    WriteLog ("Microsoft 365 usage reports inventory failed: {0}" -f $message) 'ERROR'
    try {
        Send-SmartM365TeamsNotification -Level ERROR -Channel Alerts -Title 'SmartM365 M365 usage reports inventory failed' -Message $message -Facts @{
            Tenant = $Tenant
            Period = $Period
            Reports = ($selectedReports.Name -join ',')
            RunId = $runId
            LogPath = $logPath
        } | Out-Null
    }
    catch { WriteLog ("Teams alert notification failed: {0}" -f $_.Exception.Message) 'WARNING' }
    Stop-SmartM365UsageTranscript
    try { Complete-SmartM365ExecutionContext -Status Failed -FailureStage 'M365UsageInventory' } catch {}
    throw
}

finally {
    try { RemoveOldFiles -Path $runOutputRoot -Filter '*.csv' -KeepCount $global:RetentionMaxCSV -LogFile $global:LogTextFile } catch {}
    try { RemoveOldFiles -Path $logFolder -Filter '*.log' -KeepCount $global:RetentionMaxLogs -LogFile $global:LogTextFile } catch {}
    try { Stop-SmartM365UsageTranscript } catch {}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDHrZOyDZyNBJdl
# 2A/O9DizrSti+h04ZBGulAg2hkbxzKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMhtZqMPB14+Me06lp9dlgvoyQD3If12NVEU99xQ2UO3MA0GCSqG
# SIb3DQEBAQUABIIBgC6MuDmoxaQuz/cGTaGWyrYEDyHy8HJXobxkNqu59SQ32FTb
# n3BgPfLrtqFQ1EjGRbrs5nvoJ3+oEn74A2eNX4d8UH6NmiqNfuabGvzLDeogptHe
# wZKUxlHoaiYNYt5dnIDPt+na2hCcqgZyosppmCZnloa/+ZfzSjnL4l3XEFN/P3QQ
# GwJqHylfKXis7bbn2MDjoguTeiq9vLdSusfad/vAz9dEw/LMnXa8kYGe+jRpZ2Qk
# 4sNRJws3FBTEcKSiaDvOXPe+Zbuti7M6nRSRGlFYzjpHN7IxgO9wVrchGlZfGdp6
# 8DPKkcxTJIs08P2aPBUoA2Bg1uwPJltbTmjRoHVgpH+w6fkANV0hN+WIe9uVJbs1
# MjeID5lIuqcBJGBEGYNBrXvlD+0tzCumTvgGmPSciK5YsvYDz3g3nuWqmT3929VW
# D7K0qcpfZzBeLUm7tZKu26u7gzatDmjRx7icbx30Z/K3rdb24Qo0jC/qWmf55Emk
# KKUbehvXDPTomZGyM6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODUw
# MzZaMC8GCSqGSIb3DQEJBDEiBCDWAJsiR1F5iJ+6ohZ4ZFQ+eVrCrtKhwjzEuRmv
# Sv8MoDANBgkqhkiG9w0BAQEFAASCAgDD9SOYvPonEbn+Sh0YACr5gkf1kbMLpba1
# sHN6DzIgRhuhceKsNSTXHgiPUNIvh/tHRQXqvrllvzNc13GhmnzDPJIIDSryJI+J
# VUzS6VjrfvhIU4XyvRxscaHOPDjmwtKkzDajBVQMu+vGI2ZX1m1Dkv447JR2P1zN
# g1NM54dh6/QAqvvXaVog+UuH8myFhNyjqd95X2b8ZkH0llfNpq/ND3BaNwPzeM3L
# u78l+ZG4QQ6zANQ60tMex7+TFoj6b2Y9Gr4KY+jqsv7xp0rzrMbp7hAUnUo8j4WH
# dod3fHgcbgG4KOVE6oTIeKjSzGbM9RoAPEHihpKREuvWPyt5T6YAVAxt2jSq2V7f
# FfrbKwQuvxr83rJMz8/QkMaSEUKUjLnsqui5ADxQnbI0zwfoHeb2bNRRwDEfo5v8
# 7MhPA3Mjaf+AlhnZ7ghlbSMfMw1skkEA4ysHHJBCHbg7u2fm8yZrZaHhQpWubEYY
# qZbgt0EW4QpT04eASzToom/24XWf3mCw/Zfq049XUWP3GwpfPwsxOHFSCfq7q9U8
# wbEIziFBCvRdXnGAFKKKfK4TEOCsxU7HSuVuRGauvoNrcJDu473EUplYvogUDKyZ
# /h6odxnMglh306+4fYIiXvNdUxUvySpzWOqK222X19aylEpS94MI4VaTnuVyJgzl
# BgsSyx2r7w==
# SIG # End signature block
