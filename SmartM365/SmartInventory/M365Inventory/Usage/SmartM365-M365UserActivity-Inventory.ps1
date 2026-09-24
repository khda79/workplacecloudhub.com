<#
.SYNOPSIS
Exports Microsoft 365 user and workload usage details from Microsoft Graph reports.

.DESCRIPTION
Downloads selected Microsoft 365 usage reports, normalizes the active user detail
report, and publishes stable CSV files into the tenant DATA-LAST folder for
SmartFinOps and downstream inventory analysis.

.VERSION
1.17


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
    [ValidateSet('All', 'Office365ActiveUserDetail', 'MailboxUsageDetail', 'OneDriveUsageAccountDetail', 'SharePointSiteUsageDetail', 'SharePointActivityUserDetail', 'Office365ActivationUserDetail', 'M365AppUserDetailD30', 'M365AppUserDetailD180', 'TeamsUserActivityUserDetail', 'TeamsDeviceUsageUserDetail', 'EmailActivityUserDetail')]
    [string[]]$Reports = @('Office365ActiveUserDetail', 'M365AppUserDetailD30', 'M365AppUserDetailD180'),
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
$ScriptVersion = "1.17"
$TaskName = "SmartM365-M365UserActivity-Inventory v$ScriptVersion"
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
        SharePointActivityUserDetail = [pscustomobject]@{
            Name = 'SharePointActivityUserDetail'
            Command = 'Get-MgReportSharePointActivityUserDetail'
            Endpoint = 'getSharePointActivityUserDetail'
            BaseFileName = 'M365_SharePoint_UserActivity'
            SupportsPeriod = $true
            Normalize = 'Generic'
            Columns = @(
                'RunId', 'ReportName', 'ReportPeriodRequested',
                'Report Refresh Date', 'User Principal Name', 'Is Deleted', 'Deleted Date', 'Last Activity Date',
                'Viewed Or Edited File Count', 'Synced File Count', 'Shared Internally File Count',
                'Shared Externally File Count', 'Visited Page Count', 'Assigned Products', 'Report Period'
            )
            Description = 'SharePoint file, sharing, sync, and page activity by user.'
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
        M365AppUserDetailD30 = [pscustomobject]@{
            Name = 'M365AppUserDetailD30'
            Command = 'Get-MgReportM365AppUserDetail'
            Endpoint = 'getM365AppUserDetail'
            BaseFileName = 'M365_Apps_Usage_30D'
            SupportsPeriod = $true
            PeriodOverride = 'D30'
            Normalize = 'Generic'
            Description = 'Observed Microsoft 365 Apps usage by user and platform over 30 days.'
        }
        M365AppUserDetailD180 = [pscustomobject]@{
            Name = 'M365AppUserDetailD180'
            Command = 'Get-MgReportM365AppUserDetail'
            Endpoint = 'getM365AppUserDetail'
            BaseFileName = 'M365_Apps_Usage_180D'
            SupportsPeriod = $true
            PeriodOverride = 'D180'
            Normalize = 'Generic'
            Description = 'Observed Microsoft 365 Apps usage by user and platform over 180 days.'
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
        TeamsDeviceUsageUserDetail = [pscustomobject]@{
            Name = 'TeamsDeviceUsageUserDetail'
            Command = 'Get-MgReportTeamDeviceUsageUserDetail'
            Endpoint = 'getTeamsDeviceUsageUserDetail'
            BaseFileName = 'M365_Teams_DeviceUsage'
            SupportsPeriod = $true
            Normalize = 'Generic'
            Columns = @(
                'RunId', 'ReportName', 'ReportPeriodRequested',
                'Report Refresh Date', 'User Id', 'User Principal Name', 'Last Activity Date', 'Is Deleted',
                'Deleted Date', 'Used Web', 'Used Windows Phone', 'Used iOS', 'Used Mac',
                'Used Android Phone', 'Used Windows', 'Used Chrome OS', 'Used Linux', 'Is Licensed', 'Report Period'
            )
            Description = 'Teams device and operating system usage by user.'
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
        [Parameter(Mandatory)][string]$ReportName,
        [Parameter(Mandatory)][string]$RequestedPeriod
    )

    foreach ($row in $Rows) {
        $ordered = [ordered]@{
            RunId = $runId
            ReportName = $ReportName
            ReportPeriodRequested = $RequestedPeriod
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

    $requestedPeriod = if ($ReportDefinition.PSObject.Properties['PeriodOverride']) { [string]$ReportDefinition.PeriodOverride } else { $Period }
    if ($ReportDefinition.SupportsPeriod) {
        $uri = "https://graph.microsoft.com/v1.0/reports/$endpoint(period='$requestedPeriod')"
    }
    else {
        $uri = "https://graph.microsoft.com/v1.0/reports/$endpoint"
    }

    try {
        Invoke-MgGraphRequest -Method GET -Uri $uri -OutputFilePath $RawPath -ProgressAction SilentlyContinue -ErrorAction Stop | Out-Null
    }
    catch {
        $restError = $_.Exception.Message
        $commandName = [string]$ReportDefinition.Command
        if ([string]::IsNullOrWhiteSpace($commandName)) {
            throw
        }

        WriteLog ("Graph REST download failed for {0}; retrying with {1}. Error={2}" -f $ReportDefinition.Name, $commandName, $restError) 'WARNING'
        if (Test-Path -LiteralPath $RawPath) {
            Remove-Item -LiteralPath $RawPath -Force -ErrorAction SilentlyContinue
        }

        $commandParameters = @{
            OutFile = $RawPath
            ProgressAction = 'SilentlyContinue'
            ErrorAction = 'Stop'
        }
        if ($ReportDefinition.SupportsPeriod) {
            $commandParameters.Period = $requestedPeriod
        }

        try {
            & $commandName @commandParameters | Out-Null
        }
        catch {
            throw "Report download failed via Graph REST ($restError) and $commandName ($($_.Exception.Message))."
        }
    }

    if (-not (Test-Path -LiteralPath $RawPath) -or (Get-Item -LiteralPath $RawPath).Length -le 0) {
        throw "Graph report download produced no file content for $($ReportDefinition.Name)."
    }
}

function Export-M365UsageReport {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$ReportDefinition)

    $requestedPeriod = if ($ReportDefinition.PSObject.Properties['PeriodOverride']) { [string]$ReportDefinition.PeriodOverride } else { $Period }
    $rawPath = Join-Path -Path $runOutputRoot -ChildPath ("{0}_Raw_{1}_{2}.csv" -f $ReportDefinition.BaseFileName, $requestedPeriod, $runId)
    $effectivePeriod = if ($ReportDefinition.SupportsPeriod) { $requestedPeriod } else { 'not applicable' }
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
        $normalizedRows = @(ConvertFrom-M365GenericReport -Rows $rawRows -ReportName $ReportDefinition.Name -RequestedPeriod $requestedPeriod)
        $columns = if ($ReportDefinition.PSObject.Properties['Columns']) { @($ReportDefinition.Columns) } else { @() }
    }

    $exportData = if ($normalizedRows.Count -eq 0 -and $columns.Count -gt 0) { @($null) } else { $normalizedRows }
    $exportParameters = @{
        BaseFileName = $ReportDefinition.BaseFileName
        OutputPath = $runOutputRoot
        GlobalPath = $LatestCsvFolderPath
        Data = $exportData
        NoWeeklyHistory = $true
    }
    if ($columns.Count -gt 0) {
        $exportParameters.Columns = $columns
    }
    $exportResult = Export-SmartM365Csv @exportParameters

    [pscustomobject]@{
        ReportName = $ReportDefinition.Name
        Rows = $normalizedRows.Count
        LatestPath = $exportResult.LatestPath
        PublishedPath = $exportResult.PublishedPath
        TimestampedPath = $exportResult.TimestampedPath
        RawPath = $rawPath
    }
}

function Publish-M365UsageWeeklyHistory {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$ReportResults)

    if (Test-SmartM365MaxItemsMode) { return }

    $weeklyHistoryEnabled = [bool](Get-SmartM365ConfigValue -Name 'EnableWeeklyHistory' -DefaultValue $true)
    if (-not $weeklyHistoryEnabled) {
        WriteLog 'Consolidated WeeklyHistory publication is disabled by configuration.' 'INFO'
        return
    }

    $sourceFiles = @(
        $ReportResults |
            ForEach-Object { [string]$_.PublishedPath } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Sort-Object -Unique
    )
    if ($sourceFiles.Count -eq 0) {
        WriteLog 'Consolidated WeeklyHistory publication skipped: no published CSV file found.' 'WARNING'
        return
    }

    $historyRootPath = [string](Get-SmartM365ConfigValue -Name 'WeeklyHistoryFolderPath' -DefaultValue '')
    if ([string]::IsNullOrWhiteSpace($historyRootPath)) {
        $historyRootPath = Join-Path -Path $runOutputRoot -ChildPath 'WeeklyHistory'
    }
    $retentionWeeks = [int](Get-SmartM365ConfigValue -Name 'WeeklyHistoryRetentionWeeks' -DefaultValue 52)

    WriteLog ("Publishing consolidated WeeklyHistory for {0} report(s)." -f $sourceFiles.Count) 'INFO'
    Save-SmartM365WeeklyInventoryHistory `
        -SourceFiles $sourceFiles `
        -HistoryRootPath $historyRootPath `
        -RetentionWeeks $retentionWeeks
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
$logFileBaseName = 'SmartM365-UserActivity-Inventory'
$logFolder = Join-Path -Path $logAllRoot -ChildPath $logFileBaseName
$logPath = Join-Path -Path $logFolder -ChildPath ("{0}_{1}.log" -f $logFileBaseName,$runId)

$modulePath = Join-Path -Path ([string](Get-SmartM365ConfigValue -Name 'SmartM365RootPath' -DefaultValue (Split-Path -Path $PSScriptRoot -Parent))) -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
Import-Module -Name $modulePath -MinimumVersion '1.0.57' -Force -ErrorAction Stop

$global:RetentionMaxCSV = [int](Get-SmartM365ConfigValue -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-SmartM365ConfigValue -Name 'RetentionMaxLogs' -DefaultValue 30)
$global:EnableSharePointUpload = [bool](Get-SmartM365ConfigValue -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = Get-SmartM365ConfigValue -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-SmartM365ConfigValue -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-SmartM365ConfigValue -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-SmartM365ConfigValue -Name 'SharePointTargetFolderPath' -DefaultValue ''

$global:SmartM365ExecutionStartTime = Get-Date
$global:SmartM365ExecutionSummaryWritten = $false
$global:SmartM365ScriptName = $TaskName
Set-SmartM365CoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestCsvFolderPath -LogPath $logPath
$global:LogTextFile = $logPath
$global:logTextFile = $logPath
$global:LogPath = $logFolder
$global:logTranscriptFile = Join-Path -Path $logFolder -ChildPath ("{0}_{1}_Transcript.log" -f $logFileBaseName,$runId)

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
    WriteLog ("Starting {0}. Tenant={1}; Period={2}; Reports={3}; RunId={4}" -f $TaskName, $Tenant, $Period, ($selectedReports.Name -join ','), $runId) 'INFO'

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

    Invoke-SmartM365Preflight `
        -ScriptName $TaskName `
        -RequiredModules @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Reports') `
        -RequiredCommands $requiredReportCommands `
        -RequiredGraphApplicationPermissions @('Reports.Read.All') `
        -OutputPaths @($runOutputRoot, $LatestCsvFolderPath) | Out-Null

    if ($ValidateOnly) {
        WriteLog "Validation completed. Tenant=$Tenant; Period=$Period; Reports=$(($selectedReports.Name -join ',')); OutputPath=$runOutputRoot; LatestCsvFolderPath=$LatestCsvFolderPath" 'SUCCESS'
        return
    }

    $reportResults = New-Object System.Collections.Generic.List[object]
    foreach ($reportDefinition in $selectedReports) {
        $reportResults.Add((Export-M365UsageReport -ReportDefinition $reportDefinition)) | Out-Null
    }
    Publish-M365UsageWeeklyHistory -ReportResults @($reportResults.ToArray())

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
    try { Remove-SmartM365TimestampedFilesOlderThan -FolderPath $runOutputRoot -FilePattern '*.csv' -RetentionDays 7 -RequireCurrentRunPublication -LogFile $global:LogTextFile } catch {}
    try { RemoveOldFiles -Path $logFolder -Filter '*.log' -KeepCount $global:RetentionMaxLogs -LogFile $global:LogTextFile } catch {}
    try { Stop-SmartM365UsageTranscript } catch {}
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAuM+0hD8yEwlO7
# SlvRsLC3jL5PmfQxME1/ojtCWSveYqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAjlw/uW/03II7CSkHu7Cxn
# MsbjWV40AYt4IBnnRpdnNTANBgkqhkiG9w0BAQEFAASCAYBLLlLa7l9/QHtGJgmS
# 5+az/IsBqi5/tv6P+fdy14tFd9q6lRhcEelcxlyLxOPz7I0mla4LA2PMbLXnA6pt
# xw7l5MsI2c88gUQZtY3W+I1E3Z/4yf4UtoyHK0xMrucqwzBPN00Ru1O8cJtBHPkl
# 1Pc7cEgUp/xdCad4eieqpEf1T0gc3nXLNatPRZIk74aBuk8CHvqXHTcSczAF4DeE
# km/ktd8jzYvpN8nzyPs26+hQBSyqQjBdLhuzxoRBkONg4eTcErc79Rsx5v+ZBtlL
# OvtKHTgVrSjw9QpbH3lYcipzN84kTIZfCxZeaxlHr6eKGlBJi7XGTEQ88eg1fzHZ
# 9L0KZTTx/HGYTxl55wW+niVCP/ghdSnTQHsiAK7/iq6K5OYcnvlVtraUosMWfr7A
# PdvIG1W9nsN+d5gzZUg/4q6Yscwh7QCC3DD3EhsMdq7aisvOBgPDuZ6xoY4/jO2s
# Qxxx7uIRHB27AJdAiBnGbg+mKEXU7W3rq3aM4IUka2jaBwM=
# SIG # End signature block
