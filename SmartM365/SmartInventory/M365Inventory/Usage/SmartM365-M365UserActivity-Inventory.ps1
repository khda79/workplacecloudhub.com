<#
.SYNOPSIS
Exports Microsoft 365 user and workload usage details from Microsoft Graph reports.

.DESCRIPTION
Downloads selected Microsoft 365 usage reports, normalizes the active user detail
report, and publishes stable CSV files into the tenant DATA-LAST folder for
SmartFinOps and downstream inventory analysis.

.VERSION
1.1

.NOTES
Author: https://github.com/khda79/workplacecloudhub.com
Requires: PowerShell 7+, Microsoft.Graph.Reports, SmartM365.Core.psd1
Scopes: Reports.Read.All
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
    [switch]$ValidateOnly
)

$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.1'
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
$logFolder = Join-Path -Path $logAllRoot -ChildPath 'M365\Usage'
$logPath = Join-Path -Path $logFolder -ChildPath ("SmartM365-M365UserActivity-Inventory_{0}.log" -f $runId)

$modulePath = Join-Path -Path ([string](Get-SmartM365ConfigValue -Name 'SmartM365RootPath' -DefaultValue (Split-Path -Path $PSScriptRoot -Parent))) -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
Import-Module $modulePath -Force -ErrorAction Stop

Set-SmartM365CoreContext -RunId $runId -RunOutputRoot $runOutputRoot -LatestOutputRoot $LatestCsvFolderPath -LogPath $logPath
$selectedReports = Resolve-M365UsageReportSelection -RequestedReports $Reports
$selectedReportCommands = @($selectedReports | ForEach-Object { $_.Command })
$requiredReportCommands = @(($selectedReportCommands + @('Invoke-MgGraphRequest')) | Sort-Object -Unique)

try {
    foreach ($folder in @($runOutputRoot, $LatestCsvFolderPath, $logFolder)) {
        if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
    }

    Invoke-SmartM365Preflight `
        -ScriptName 'SmartM365-M365UserActivity-Inventory' `
        -RequiredModules @('Microsoft.Graph.Authentication', 'Microsoft.Graph.Reports') `
        -RequiredCommands $requiredReportCommands `
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
    throw
}
