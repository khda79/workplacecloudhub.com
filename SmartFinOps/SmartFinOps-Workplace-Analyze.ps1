<#
.SYNOPSIS
Builds a SmartFinOps Workplace report from SmartM365 SmartInventory CSV exports.

.DESCRIPTION
Reads existing SmartM365 DATA-LAST CSV files and produces FinOps-oriented CSV outputs
and a standalone HTML report. The script is read-only and does not connect to Microsoft
Graph, Azure, Citrix, or Azure Virtual Desktop.

.NOTES
Version: 1.0
Author: https://github.com/khda79/workplacecloudhub.com
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$SmartM365LatestCsvFolderPath,
    [string]$OutputRoot,
    [string]$LatestOutputRoot,
    [int]$StaleUserDays = 0,
    [int]$StaleDeviceDays = 0,
    [string]$ReportTitle,
    [switch]$ValidateOnly
)

$ErrorActionPreference = 'Stop'
$runId = Get-Date -Format 'yyyyMMdd_HHmmss'

$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidate = Join-Path -Path $d -ChildPath 'Config\SmartFinOps-TenantContext.ps1'
        if (Test-Path -LiteralPath $candidate) { return $candidate }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartFinOps-TenantContext.ps1 not found.'
}
. $tenantContextPath

$script:SmartFinOpsEffectiveConfig = Initialize-SmartFinOpsTenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$ScriptLocalConfig = Get-SmartFinOpsScriptLocalConfig -ScriptPath $PSCommandPath

if ([string]::IsNullOrWhiteSpace($SmartM365LatestCsvFolderPath)) { $SmartM365LatestCsvFolderPath = Get-SmartFinOpsScriptConfigValue -Config $ScriptLocalConfig -Name 'SmartM365LatestCsvFolderPath' -DefaultValue '' }
if ([string]::IsNullOrWhiteSpace($OutputRoot)) { $OutputRoot = Get-SmartFinOpsScriptConfigValue -Config $ScriptLocalConfig -Name 'OutputRoot' -DefaultValue '' }
if ([string]::IsNullOrWhiteSpace($LatestOutputRoot)) { $LatestOutputRoot = Get-SmartFinOpsScriptConfigValue -Config $ScriptLocalConfig -Name 'LatestOutputRoot' -DefaultValue '' }
if ($StaleUserDays -le 0) { $StaleUserDays = [int](Get-SmartFinOpsScriptConfigValue -Config $ScriptLocalConfig -Name 'StaleUserDays' -DefaultValue 90) }
if ($StaleDeviceDays -le 0) { $StaleDeviceDays = [int](Get-SmartFinOpsScriptConfigValue -Config $ScriptLocalConfig -Name 'StaleDeviceDays' -DefaultValue 60) }
if ([string]::IsNullOrWhiteSpace($ReportTitle)) { $ReportTitle = [string](Get-SmartFinOpsScriptConfigValue -Config $ScriptLocalConfig -Name 'ReportTitle' -DefaultValue 'SmartFinOps Workplace Report') }

$resolvedOutputRoots = Resolve-SmartFinOpsOutputRoots -OutputRoot $OutputRoot -LatestOutputRoot $LatestOutputRoot -AreaPath 'Workplace'
$OutputRoot = $resolvedOutputRoots.OutputRoot
$LatestOutputRoot = $resolvedOutputRoots.LatestOutputRoot
$runOutputRoot = Join-Path -Path $OutputRoot -ChildPath $runId
$logRoot = Resolve-SmartFinOpsConfigTokenValue -Value '{{LogAllRootPath}}'
$logFolder = Join-Path -Path $logRoot -ChildPath 'Workplace'
$logPath = Join-Path -Path $logFolder -ChildPath ("SmartFinOps-Workplace-Analyze_{0}.log" -f $runId)

$script:RunId = $runId
$script:RunOutputRoot = $runOutputRoot
$script:LatestOutputRoot = $LatestOutputRoot
$script:LogPath = $logPath
$script:SmartM365LatestCsvFolderPath = $SmartM365LatestCsvFolderPath

function Write-SmartFinOpsLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')][string]$Level = 'INFO'
    )
    $line = "{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Write-Host $line
    $folder = Split-Path -Path $script:LogPath -Parent
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
    Add-Content -LiteralPath $script:LogPath -Value $line -Encoding UTF8
}

function Get-RowPropertyValue {
    [CmdletBinding()]
    param([AllowNull()]$Row, [Parameter(Mandatory)][string[]]$Names)
    if ($null -eq $Row) { return $null }
    foreach ($name in $Names) {
        $property = $Row.PSObject.Properties[$name]
        if ($null -ne $property) { return $property.Value }
    }
    return $null
}

function ConvertTo-DateTimeOrNull {
    [CmdletBinding()]
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return $null }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    try { return [datetime]$text } catch { return $null }
}

function ConvertTo-BoolOrNull {
    [CmdletBinding()]
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return $null }
    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    if ($text -match '^(true|yes|1)$') { return $true }
    if ($text -match '^(false|no|0)$') { return $false }
    return $null
}

function Resolve-FirstExistingCsv {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$FolderPath, [Parameter(Mandatory)][string[]]$FileNames)
    foreach ($fileName in $FileNames) {
        $path = Join-Path -Path $FolderPath -ChildPath $fileName
        if (Test-Path -LiteralPath $path -PathType Leaf) { return $path }
    }
    return ''
}

function Import-SmartFinOpsSourceCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SourceName,
        [Parameter(Mandatory)][string[]]$FileNames,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.Generic.List[object]]$DataQualityRows
    )
    $path = Resolve-FirstExistingCsv -FolderPath $script:SmartM365LatestCsvFolderPath -FileNames $FileNames
    if ([string]::IsNullOrWhiteSpace($path)) {
        $DataQualityRows.Add([pscustomobject]@{ RunId = $script:RunId; SourceName = $SourceName; Status = 'Missing'; Path = ($FileNames -join ' | '); RowCount = 0; LastWriteTime = ''; Notes = 'Source CSV not found in SmartM365 DATA-LAST.' }) | Out-Null
        return @()
    }
    try {
        $rows = @(Import-Csv -LiteralPath $path)
        $item = Get-Item -LiteralPath $path
        $DataQualityRows.Add([pscustomobject]@{ RunId = $script:RunId; SourceName = $SourceName; Status = 'Loaded'; Path = $path; RowCount = $rows.Count; LastWriteTime = $item.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'); Notes = '' }) | Out-Null
        Write-SmartFinOpsLog -Message ("Loaded source {0}: {1} row(s)" -f $SourceName, $rows.Count)
        return $rows
    }
    catch {
        $DataQualityRows.Add([pscustomobject]@{ RunId = $script:RunId; SourceName = $SourceName; Status = 'Error'; Path = $path; RowCount = 0; LastWriteTime = ''; Notes = $_.Exception.Message }) | Out-Null
        Write-SmartFinOpsLog -Level WARN -Message ("Failed to load source {0}: {1}" -f $SourceName, $_.Exception.Message)
        return @()
    }
}

function Get-SmartFinOpsSummaryRow {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Category, [Parameter(Mandatory)][string]$Metric, [Parameter(Mandatory)][object]$Value, [string]$Unit = 'count', [string]$Interpretation = '')
    [pscustomobject]@{ RunId = $script:RunId; Category = $Category; Metric = $Metric; Value = $Value; Unit = $Unit; Interpretation = $Interpretation }
}

function Export-SmartFinOpsCsv {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows)
    $historyPath = Join-Path -Path $script:RunOutputRoot -ChildPath ("{0}_{1}.csv" -f $Name, $script:RunId)
    $latestPath = Join-Path -Path $script:LatestOutputRoot -ChildPath ("{0}.csv" -f $Name)
    foreach ($folder in @((Split-Path -Path $historyPath -Parent), (Split-Path -Path $latestPath -Parent))) {
        if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
    }
    $Rows | Export-Csv -LiteralPath $historyPath -NoTypeInformation -Encoding UTF8
    Copy-Item -LiteralPath $historyPath -Destination $latestPath -Force
    Write-SmartFinOpsLog -Message ("Exported {0} row(s): {1}" -f @($Rows).Count, $latestPath)
    return $latestPath
}

function Get-PriceModel {
    [CmdletBinding()]
    param()
    $model = Get-SmartFinOpsScriptConfigValue -Config $ScriptLocalConfig -Name 'PriceModel' -DefaultValue $null
    if ($null -eq $model) { return [pscustomobject]@{ Currency = ''; MonthlyUnitPriceBySkuPartNumber = [pscustomobject]@{} } }
    return $model
}

function Get-MonthlySkuPrice {
    [CmdletBinding()]
    param([AllowNull()]$PriceModel, [string]$SkuPartNumber)
    if ($null -eq $PriceModel -or [string]::IsNullOrWhiteSpace($SkuPartNumber)) { return $null }
    $mapProperty = $PriceModel.PSObject.Properties['MonthlyUnitPriceBySkuPartNumber']
    if ($null -eq $mapProperty -or $null -eq $mapProperty.Value) { return $null }
    $priceProperty = $mapProperty.Value.PSObject.Properties[$SkuPartNumber]
    if ($null -eq $priceProperty -or $null -eq $priceProperty.Value) { return $null }
    try { return [decimal]$priceProperty.Value } catch { return $null }
}

function ConvertTo-HtmlEncoded {
    [CmdletBinding()]
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Convert-RowsToHtmlTable {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows, [int]$MaxRows = 20)
    if (-not $Rows -or $Rows.Count -eq 0) { return '<p class="empty">No rows.</p>' }
    $sample = @($Rows | Select-Object -First $MaxRows)
    $columns = @($sample[0].PSObject.Properties | ForEach-Object { $_.Name })
    $html = New-Object System.Collections.Generic.List[string]
    $html.Add('<table>') | Out-Null
    $html.Add('<thead><tr>') | Out-Null
    foreach ($column in $columns) { $html.Add(('<th>{0}</th>' -f (ConvertTo-HtmlEncoded $column))) | Out-Null }
    $html.Add('</tr></thead><tbody>') | Out-Null
    foreach ($row in $sample) {
        $html.Add('<tr>') | Out-Null
        foreach ($column in $columns) { $html.Add(('<td>{0}</td>' -f (ConvertTo-HtmlEncoded (Get-RowPropertyValue -Row $row -Names @($column))))) | Out-Null }
        $html.Add('</tr>') | Out-Null
    }
    $html.Add('</tbody></table>') | Out-Null
    if ($Rows.Count -gt $MaxRows) { $html.Add(('<p class="note">Showing {0} of {1} rows.</p>' -f $MaxRows, $Rows.Count)) | Out-Null }
    return ($html -join [Environment]::NewLine)
}

function Write-SmartFinOpsHtmlReport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$SummaryRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$LicenseRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$DeviceRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$DataQualityRows
    )
    $loadedSources = @($DataQualityRows | Where-Object { $_.Status -eq 'Loaded' }).Count
    $missingSources = @($DataQualityRows | Where-Object { $_.Status -eq 'Missing' }).Count
    $errorSources = @($DataQualityRows | Where-Object { $_.Status -eq 'Error' }).Count
    $pricedRows = @($LicenseRows | Where-Object { $null -ne $_.EstimatedMonthlyWaste -and $_.EstimatedMonthlyWaste -ne '' })
    $totalPotentialMonthly = if ($pricedRows.Count -gt 0) { ($pricedRows | Measure-Object -Property EstimatedMonthlyWaste -Sum).Sum } else { $null }
    $potentialText = if ($null -ne $totalPotentialMonthly) { ('{0:N2}' -f $totalPotentialMonthly) } else { 'Not priced' }

    $html = @"
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>$(ConvertTo-HtmlEncoded $ReportTitle)</title>
  <style>
    body { margin: 0; font-family: Segoe UI, Arial, sans-serif; background: #f5f8fb; color: #1f2937; }
    header { background: #ffffff; border-bottom: 1px solid #dde7f0; padding: 24px 32px; }
    main { padding: 24px 32px 40px; }
    h1 { margin: 0 0 8px; font-size: 26px; font-weight: 650; }
    h2 { margin: 28px 0 12px; font-size: 18px; }
    .subtitle { color: #5f6b7a; margin: 0; }
    .grid { display: grid; grid-template-columns: repeat(4, minmax(160px, 1fr)); gap: 12px; margin-top: 18px; }
    .metric { background: #fff; border: 1px solid #dde7f0; border-radius: 8px; padding: 14px; }
    .metric .label { color: #5f6b7a; font-size: 12px; }
    .metric .value { font-size: 24px; font-weight: 650; margin-top: 4px; }
    section { background: #fff; border: 1px solid #dde7f0; border-radius: 8px; padding: 16px; margin-top: 16px; }
    table { border-collapse: collapse; width: 100%; font-size: 12px; }
    th { text-align: left; background: #eef5fb; color: #1f2937; border-bottom: 1px solid #dde7f0; padding: 8px; }
    td { border-bottom: 1px solid #edf2f7; padding: 7px 8px; vertical-align: top; }
    .note, .empty { color: #5f6b7a; font-size: 12px; }
    .pill { display: inline-block; border-radius: 999px; background: #e6f4ff; color: #005a9e; padding: 3px 9px; font-size: 12px; font-weight: 600; }
  </style>
</head>
<body>
  <header>
    <div class="pill">SmartFinOps Workplace</div>
    <h1>$(ConvertTo-HtmlEncoded $ReportTitle)</h1>
    <p class="subtitle">Tenant: $(ConvertTo-HtmlEncoded $Tenant) | RunId: $(ConvertTo-HtmlEncoded $runId) | Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</p>
    <div class="grid">
      <div class="metric"><div class="label">Loaded sources</div><div class="value">$loadedSources</div></div>
      <div class="metric"><div class="label">Missing sources</div><div class="value">$missingSources</div></div>
      <div class="metric"><div class="label">Source errors</div><div class="value">$errorSources</div></div>
      <div class="metric"><div class="label">Potential monthly waste</div><div class="value">$potentialText</div></div>
    </div>
  </header>
  <main>
    <section><h2>Summary</h2>$(Convert-RowsToHtmlTable -Rows $SummaryRows -MaxRows 80)</section>
    <section><h2>License Optimization</h2>$(Convert-RowsToHtmlTable -Rows $LicenseRows -MaxRows 40)</section>
    <section><h2>Device Optimization</h2>$(Convert-RowsToHtmlTable -Rows $DeviceRows -MaxRows 40)</section>
    <section><h2>Data Quality</h2>$(Convert-RowsToHtmlTable -Rows $DataQualityRows -MaxRows 80)</section>
  </main>
</body>
</html>
"@
    $folder = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
    Set-Content -LiteralPath $Path -Value $html -Encoding UTF8
}

try {
    New-Item -Path $runOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $LatestOutputRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $logFolder -ItemType Directory -Force | Out-Null

    Write-SmartFinOpsLog -Message "Starting SmartFinOps Workplace analysis. RunId=$runId"
    Write-SmartFinOpsLog -Message ("SmartM365 latest CSV folder: {0}" -f $SmartM365LatestCsvFolderPath)

    if ($ValidateOnly) {
        Write-SmartFinOpsLog -Level SUCCESS -Message 'Validation completed.'
        return
    }

    $dataQualityRows = New-Object System.Collections.Generic.List[object]
    $summaryRows = New-Object System.Collections.Generic.List[object]
    $licenseOptimizationRows = New-Object System.Collections.Generic.List[object]
    $deviceOptimizationRows = New-Object System.Collections.Generic.List[object]

    $m365Users = Import-SmartFinOpsSourceCsv -SourceName 'M365 active users' -FileNames @('M365_Users_Active.csv') -DataQualityRows $dataQualityRows
    $m365UserActivity = Import-SmartFinOpsSourceCsv -SourceName 'M365 user activity' -FileNames @('M365_Users_Activity.csv') -DataQualityRows $dataQualityRows
    $m365MailboxUsage = Import-SmartFinOpsSourceCsv -SourceName 'M365 mailbox usage' -FileNames @('M365_Mailbox_Usage.csv') -DataQualityRows $dataQualityRows
    $m365OneDriveUsage = Import-SmartFinOpsSourceCsv -SourceName 'M365 OneDrive usage' -FileNames @('M365_OneDrive_Usage.csv') -DataQualityRows $dataQualityRows
    $m365SharePointSiteUsage = Import-SmartFinOpsSourceCsv -SourceName 'M365 SharePoint site usage' -FileNames @('M365_SharePoint_SiteUsage.csv') -DataQualityRows $dataQualityRows
    $m365AppsActivations = Import-SmartFinOpsSourceCsv -SourceName 'M365 Apps activations' -FileNames @('M365_Apps_Activations.csv') -DataQualityRows $dataQualityRows
    $m365TeamsUserActivity = Import-SmartFinOpsSourceCsv -SourceName 'M365 Teams user activity' -FileNames @('M365_Teams_UserActivity.csv') -DataQualityRows $dataQualityRows
    $m365EmailActivity = Import-SmartFinOpsSourceCsv -SourceName 'M365 email activity' -FileNames @('M365_Email_Activity.csv') -DataQualityRows $dataQualityRows
    $m365LicenseUsers = Import-SmartFinOpsSourceCsv -SourceName 'M365 license user assignments' -FileNames @('M365_Licenses_Users.csv') -DataQualityRows $dataQualityRows
    $m365LicenseTenant = Import-SmartFinOpsSourceCsv -SourceName 'M365 tenant licenses' -FileNames @('M365_Licenses_Tenant.csv') -DataQualityRows $dataQualityRows
    $intuneDevices = Import-SmartFinOpsSourceCsv -SourceName 'Intune devices' -FileNames @('Intune_Devices_Inventory.csv') -DataQualityRows $dataQualityRows
    $intuneCompliance = Import-SmartFinOpsSourceCsv -SourceName 'Intune device compliance' -FileNames @('Intune_Devices_Compliance.csv') -DataQualityRows $dataQualityRows
    $entraDevices = Import-SmartFinOpsSourceCsv -SourceName 'Entra devices' -FileNames @('M365_Entra_Devices.csv') -DataQualityRows $dataQualityRows
    $adUsers = Import-SmartFinOpsSourceCsv -SourceName 'Active Directory users' -FileNames @('AD_Users_AllDomains.csv') -DataQualityRows $dataQualityRows
    $adComputers = Import-SmartFinOpsSourceCsv -SourceName 'Active Directory computers' -FileNames @('AD_Computers_AllDomains.csv') -DataQualityRows $dataQualityRows
    $autopilotDevices = Import-SmartFinOpsSourceCsv -SourceName 'Windows Autopilot devices' -FileNames @('Intune_Autopilot_Devices.csv') -DataQualityRows $dataQualityRows
    $upgradeEligibility = Import-SmartFinOpsSourceCsv -SourceName 'Windows upgrade eligibility' -FileNames @('Intune_Devices_UpgradeEligibility.csv', 'Intune_Devices_Windows11UpgradeEligibility.csv') -DataQualityRows $dataQualityRows
    $exoMailboxes = Import-SmartFinOpsSourceCsv -SourceName 'Exchange Online mailboxes' -FileNames @('Exchange_EXO_Mailboxes_AllDomains.csv', 'Exchange_EXO_Mailboxes_AllDomains_Stats.csv') -DataQualityRows $dataQualityRows
    $backupUnprotected = Import-SmartFinOpsSourceCsv -SourceName 'Exchange backup unprotected mailboxes' -FileNames @('Exchange_EXO_BackupProtection_UnprotectedMailboxes.csv') -DataQualityRows $dataQualityRows

    $now = Get-Date
    $staleUserCutoff = $now.AddDays(-1 * [math]::Abs($StaleUserDays))
    $staleDeviceCutoff = $now.AddDays(-1 * [math]::Abs($StaleDeviceDays))
    $priceModel = Get-PriceModel
    $currency = if ($priceModel -and $priceModel.PSObject.Properties['Currency']) { [string]$priceModel.Currency } else { '' }

    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'Sources' -Metric 'Loaded source files' -Value @($dataQualityRows | Where-Object { $_.Status -eq 'Loaded' }).Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'Sources' -Metric 'Missing source files' -Value @($dataQualityRows | Where-Object { $_.Status -eq 'Missing' }).Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'Users' -Metric 'M365 active users rows' -Value $m365Users.Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'Users' -Metric 'M365 user activity rows' -Value $m365UserActivity.Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'M365 Usage' -Metric 'Mailbox usage rows' -Value $m365MailboxUsage.Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'M365 Usage' -Metric 'OneDrive usage rows' -Value $m365OneDriveUsage.Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'M365 Usage' -Metric 'SharePoint site usage rows' -Value $m365SharePointSiteUsage.Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'M365 Usage' -Metric 'M365 Apps activation rows' -Value $m365AppsActivations.Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'M365 Usage' -Metric 'Teams user activity rows' -Value $m365TeamsUserActivity.Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'M365 Usage' -Metric 'Email activity rows' -Value $m365EmailActivity.Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'Users' -Metric 'AD users rows' -Value $adUsers.Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'Devices' -Metric 'Intune devices rows' -Value $intuneDevices.Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'Devices' -Metric 'Entra devices rows' -Value $entraDevices.Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'Devices' -Metric 'AD computers rows' -Value $adComputers.Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'Licenses' -Metric 'License assignment rows' -Value $m365LicenseUsers.Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'Licenses' -Metric 'Tenant SKU rows' -Value $m365LicenseTenant.Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'Exchange' -Metric 'Mailbox rows' -Value $exoMailboxes.Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'Exchange' -Metric 'Unprotected mailbox rows' -Value $backupUnprotected.Count)) | Out-Null

    $m365UserByUpn = @{}
    foreach ($user in $m365Users) {
        $upn = [string](Get-RowPropertyValue -Row $user -Names @('User principal name', 'UserPrincipalName'))
        if (-not [string]::IsNullOrWhiteSpace($upn)) { $m365UserByUpn[$upn.ToLowerInvariant()] = $user }
    }
    $m365ActivityByUpn = @{}
    foreach ($activityRow in $m365UserActivity) {
        $upn = [string](Get-RowPropertyValue -Row $activityRow -Names @('UserPrincipalName', 'User Principal Name'))
        if (-not [string]::IsNullOrWhiteSpace($upn)) { $m365ActivityByUpn[$upn.ToLowerInvariant()] = $activityRow }
    }
    $adUserByUpn = @{}
    foreach ($user in $adUsers) {
        $upn = [string](Get-RowPropertyValue -Row $user -Names @('UserPrincipalName', 'User principal name'))
        if (-not [string]::IsNullOrWhiteSpace($upn)) { $adUserByUpn[$upn.ToLowerInvariant()] = $user }
    }

    $uniqueLicensedUserKeys = New-Object System.Collections.Generic.HashSet[string]
    foreach ($licenseRow in $m365LicenseUsers) {
        $upn = [string](Get-RowPropertyValue -Row $licenseRow -Names @('User principal name', 'UserPrincipalName'))
        $skuPartNumber = [string](Get-RowPropertyValue -Row $licenseRow -Names @('SkuPartNumber', 'SKU part number'))
        $skuName = [string](Get-RowPropertyValue -Row $licenseRow -Names @('SKU name', 'SkuName'))
        if (-not [string]::IsNullOrWhiteSpace($upn)) { [void]$uniqueLicensedUserKeys.Add($upn.ToLowerInvariant()) }
        if ([string]::IsNullOrWhiteSpace($upn) -or [string]::IsNullOrWhiteSpace($skuPartNumber)) { continue }

        $userKey = $upn.ToLowerInvariant()
        $m365User = if ($m365UserByUpn.ContainsKey($userKey)) { $m365UserByUpn[$userKey] } else { $null }
        $m365Activity = if ($m365ActivityByUpn.ContainsKey($userKey)) { $m365ActivityByUpn[$userKey] } else { $null }
        $adUser = if ($adUserByUpn.ContainsKey($userKey)) { $adUserByUpn[$userKey] } else { $null }
        $lastM365Activity = ConvertTo-DateTimeOrNull (Get-RowPropertyValue -Row $m365Activity -Names @('LastActivityDate'))
        $lastM365ActivityWorkload = [string](Get-RowPropertyValue -Row $m365Activity -Names @('LastActivityWorkload'))
        $lastSignIn = ConvertTo-DateTimeOrNull (Get-RowPropertyValue -Row $m365User -Names @('LastSignInDateTime', 'Last sign-in time'))
        $lastAdLogon = ConvertTo-DateTimeOrNull (Get-RowPropertyValue -Row $adUser -Names @('LastLogonDate', 'Last logon date'))
        $accountEnabled = ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $m365User -Names @('AccountEnabled'))
        if ($null -eq $accountEnabled) {
            $blockCredential = ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $m365User -Names @('Block credential'))
            if ($null -ne $blockCredential) { $accountEnabled = -not $blockCredential }
        }
        $adEnabled = ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $adUser -Names @('Enabled'))

        $reason = ''
        if ($null -ne $accountEnabled -and -not $accountEnabled) { $reason = 'Licensed M365 user is disabled or blocked.' }
        elseif ($null -ne $adEnabled -and -not $adEnabled) { $reason = 'Licensed user is disabled in Active Directory.' }
        elseif ($lastM365Activity -and $lastM365Activity -lt $staleUserCutoff) { $reason = "Licensed user has no recent M365 workload activity within $StaleUserDays days." }
        elseif ($lastSignIn -and $lastSignIn -lt $staleUserCutoff) { $reason = "Licensed user has no recent M365 sign-in within $StaleUserDays days." }
        elseif ($lastAdLogon -and $lastAdLogon -lt $staleUserCutoff) { $reason = "Licensed user has no recent AD logon within $StaleUserDays days." }
        elseif (-not $m365User -and -not $adUser) { $reason = 'Licensed user was not found in M365 active users or AD user exports.' }

        if (-not [string]::IsNullOrWhiteSpace($reason)) {
            $monthlyPrice = Get-MonthlySkuPrice -PriceModel $priceModel -SkuPartNumber $skuPartNumber
            $licenseOptimizationRows.Add([pscustomobject]@{
                RunId = $runId
                UserPrincipalName = $upn
                SkuPartNumber = $skuPartNumber
                SkuName = $skuName
                Source = [string](Get-RowPropertyValue -Row $licenseRow -Names @('Source'))
                Reason = $reason
                LastM365Activity = if ($lastM365Activity) { $lastM365Activity.ToString('yyyy-MM-dd') } else { '' }
                LastM365ActivityWorkload = $lastM365ActivityWorkload
                LastM365SignIn = if ($lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd') } else { '' }
                LastADLogon = if ($lastAdLogon) { $lastAdLogon.ToString('yyyy-MM-dd') } else { '' }
                AccountEnabled = $accountEnabled
                ADEnabled = $adEnabled
                Currency = $currency
                EstimatedMonthlyWaste = $monthlyPrice
            }) | Out-Null
        }
    }
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'Licenses' -Metric 'Unique licensed users' -Value $uniqueLicensedUserKeys.Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'Licenses' -Metric 'License optimization findings' -Value $licenseOptimizationRows.Count)) | Out-Null

    foreach ($device in $intuneDevices) {
        $name = [string](Get-RowPropertyValue -Row $device -Names @('Device name', 'DeviceName', 'Name'))
        $lastCheckIn = ConvertTo-DateTimeOrNull (Get-RowPropertyValue -Row $device -Names @('Last check-in', 'LastSyncDateTime', 'Last check in'))
        $compliance = [string](Get-RowPropertyValue -Row $device -Names @('Compliance', 'ComplianceState'))
        $owner = [string](Get-RowPropertyValue -Row $device -Names @('Primary user UPN', 'UserPrincipalName'))
        $os = [string](Get-RowPropertyValue -Row $device -Names @('OS', 'OperatingSystem'))
        $osVersion = [string](Get-RowPropertyValue -Row $device -Names @('OS version', 'OsVersion'))
        $reason = ''
        if ($lastCheckIn -and $lastCheckIn -lt $staleDeviceCutoff) { $reason = "Intune device has no recent check-in within $StaleDeviceDays days." }
        elseif ($compliance -and $compliance -notmatch '^(compliant|true)$') { $reason = 'Intune device is not compliant.' }
        elseif ([string]::IsNullOrWhiteSpace($owner)) { $reason = 'Intune device has no primary user in the export.' }
        if (-not [string]::IsNullOrWhiteSpace($reason)) {
            $deviceOptimizationRows.Add([pscustomobject]@{ RunId = $runId; Source = 'Intune'; DeviceName = $name; Owner = $owner; OperatingSystem = $os; OperatingSystemVersion = $osVersion; LastActivity = if ($lastCheckIn) { $lastCheckIn.ToString('yyyy-MM-dd') } else { '' }; State = $compliance; Reason = $reason }) | Out-Null
        }
    }

    foreach ($device in $entraDevices) {
        $name = [string](Get-RowPropertyValue -Row $device -Names @('DisplayName', 'DeviceName'))
        $lastSignIn = ConvertTo-DateTimeOrNull (Get-RowPropertyValue -Row $device -Names @('ApproximateLastSignInDateTime'))
        $trustType = [string](Get-RowPropertyValue -Row $device -Names @('TrustType'))
        $os = [string](Get-RowPropertyValue -Row $device -Names @('OperatingSystem'))
        $reason = ''
        if (-not $lastSignIn) { $reason = 'Entra device has never signed in or has no approximate last sign-in date.' }
        elseif ($lastSignIn -lt $staleDeviceCutoff) { $reason = "Entra device has no recent sign-in within $StaleDeviceDays days." }
        if (-not [string]::IsNullOrWhiteSpace($reason)) {
            $deviceOptimizationRows.Add([pscustomobject]@{ RunId = $runId; Source = 'Entra'; DeviceName = $name; Owner = ''; OperatingSystem = $os; OperatingSystemVersion = [string](Get-RowPropertyValue -Row $device -Names @('OperatingSystemVersion')); LastActivity = if ($lastSignIn) { $lastSignIn.ToString('yyyy-MM-dd') } else { '' }; State = $trustType; Reason = $reason }) | Out-Null
        }
    }

    foreach ($computer in $adComputers) {
        $name = [string](Get-RowPropertyValue -Row $computer -Names @('Name', 'SamAccountName', 'ComputerName'))
        $lastLogon = ConvertTo-DateTimeOrNull (Get-RowPropertyValue -Row $computer -Names @('LastLogonDate', 'Last logon date'))
        $enabled = ConvertTo-BoolOrNull (Get-RowPropertyValue -Row $computer -Names @('Enabled'))
        $reason = ''
        if ($null -ne $enabled -and -not $enabled) { $reason = 'AD computer account is disabled.' }
        elseif ($lastLogon -and $lastLogon -lt $staleDeviceCutoff) { $reason = "AD computer has no recent logon within $StaleDeviceDays days." }
        if (-not [string]::IsNullOrWhiteSpace($reason)) {
            $deviceOptimizationRows.Add([pscustomobject]@{ RunId = $runId; Source = 'Active Directory'; DeviceName = $name; Owner = ''; OperatingSystem = [string](Get-RowPropertyValue -Row $computer -Names @('OperatingSystem')); OperatingSystemVersion = [string](Get-RowPropertyValue -Row $computer -Names @('operatingSystemVersion', 'OperatingSystemVersion')); LastActivity = if ($lastLogon) { $lastLogon.ToString('yyyy-MM-dd') } else { '' }; State = if ($null -eq $enabled) { '' } else { "Enabled=$enabled" }; Reason = $reason }) | Out-Null
        }
    }

    $nonCompliantComplianceRows = @($intuneCompliance | Where-Object {
        $state = [string](Get-RowPropertyValue -Row $_ -Names @('ComplianceState', 'Compliance', 'State'))
        $state -and $state -notmatch '^(compliant|true)$'
    })
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'Devices' -Metric 'Device optimization findings' -Value $deviceOptimizationRows.Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'Devices' -Metric 'Non-compliant rows from compliance export' -Value $nonCompliantComplianceRows.Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'Devices' -Metric 'Autopilot device rows' -Value $autopilotDevices.Count)) | Out-Null
    $summaryRows.Add((Get-SmartFinOpsSummaryRow -Category 'Devices' -Metric 'Upgrade eligibility rows' -Value $upgradeEligibility.Count)) | Out-Null

    $summaryPath = Export-SmartFinOpsCsv -Name 'SmartFinOps_Workplace_Summary' -Rows $summaryRows.ToArray()
    $licensePath = Export-SmartFinOpsCsv -Name 'SmartFinOps_Workplace_LicenseOptimization' -Rows $licenseOptimizationRows.ToArray()
    $devicePath = Export-SmartFinOpsCsv -Name 'SmartFinOps_Workplace_DeviceOptimization' -Rows $deviceOptimizationRows.ToArray()
    $dataQualityPath = Export-SmartFinOpsCsv -Name 'SmartFinOps_Workplace_DataQuality' -Rows $dataQualityRows.ToArray()

    $historyHtmlPath = Join-Path -Path $runOutputRoot -ChildPath ("SmartFinOps_Workplace_Report_{0}.html" -f $runId)
    $latestHtmlPath = Join-Path -Path $LatestOutputRoot -ChildPath 'SmartFinOps_Workplace_Report.html'
    Write-SmartFinOpsHtmlReport -Path $historyHtmlPath -SummaryRows $summaryRows.ToArray() -LicenseRows $licenseOptimizationRows.ToArray() -DeviceRows $deviceOptimizationRows.ToArray() -DataQualityRows $dataQualityRows.ToArray()
    Copy-Item -LiteralPath $historyHtmlPath -Destination $latestHtmlPath -Force

    Write-SmartFinOpsLog -Level SUCCESS -Message ("SmartFinOps Workplace analysis completed. Report={0}" -f $latestHtmlPath)
    Write-SmartFinOpsLog -Message ("CSV outputs: {0}; {1}; {2}; {3}" -f $summaryPath, $licensePath, $devicePath, $dataQualityPath)
}
catch {
    Write-SmartFinOpsLog -Level ERROR -Message $_.Exception.Message
    throw
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBOhUwbcexV/8nK
# Z02G5N8xgNELNcLjewFfFRJWLqVfhqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGj8ScMqjkH4395F9gIhvuLpTI7ZAbuFlTYTXzJ2TFfIMA0GCSqG
# SIb3DQEBAQUABIIBgHeNtupNBc6QHh0oIOvmHX5E4oDxNUIkx3I26iLIk6n8hfoo
# 8vLZnpZ5GMSB9Xu+TybD6Ld1Y2/MJnMAJbcd0fW/ZZzw6VPO3VLweZHLFqfbsnza
# B13eZXUyzICccdrnzt8y2J9hP6aMpzS4ghLacaTYz0bR34YGnPizz+mAiRJbbMwM
# C4r/aGsk15IpT0WKaRxC2cBb1HLSoIVgZOXE6QFSOakWCaMfCwSXiCdxB+v0idNV
# zAV0Kpg+I8Qf4k2SvKR7gQK3kqvaonLn0/iwTTWD+VDjmbWCIv+ufLR55JNJ+slw
# FYPNJjV5RoYrM7NLjhtrmBxu9FKDjpAAFPr3kY6LdFSw7v3bHIQO2X/GMTjWwc10
# h53VpwnMdTs5yr/O1P7BMLCYUG4NCaIlB+Pfil/Jn05W7MglJpFU/dy0TPnMV/lU
# T3CLLQddhn1xvDRGv7Sbm7N6a33vOTd3CbNuQRWSorJck4xs3vmqiWG51afYzqru
# sIynIfMs1heB2soiNaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# MjJaMC8GCSqGSIb3DQEJBDEiBCCUXSW1omx4mvKLqlG+HlJAChFBDsihul759tF1
# gbmISzANBgkqhkiG9w0BAQEFAASCAgBnOmga3/EWqPjln4lUIc6lQwRb05suU5VF
# q/ie6CGO9a9Xmo107mkwa88XQi2AuZR/haIEU+JuauFZQb066mirJd1aEa0DULe1
# LkHum5sDF2gq9peetOy2sHdaLOiauipiBG6HdJm3OYoRYoN06hnMGCkAkRJdialm
# Zh9x5KWMY7y9V/W6Yg+WM6oz//uMqoJU+ydOsYkfA61JuiQy9lN6a7QdPV9dojQT
# PdW/qCZ9tdLIi0Dku01MBeFufNoXSBMUHwyZ3L5gwna8dsoFKl93IEJkso7xcxRr
# 86yh4wouOhgCgLj+KgVIC4gj6cJnn3UVBzAZDBIdTDPiZnmCLC7C4UUxDAm2PUvI
# LRi0kfKiUWL9fLxxSGNl1m+oWqOaSJfuXFd0IlSCfAzW7mxdueT7F6bYV2bbDvza
# cpV3i3lgxPxliVdKdaSzwcd4x7CtCZp0Dz5VV+KcVTpSq2Xiyqz9iCMEteC2Lxsd
# mt/tIHhSMiA1LM4oD987wK/37An3TaLXTpbhsLSlUP9MnLaJ42mbqvziIHGtV5eK
# +O6aUAF5aAG4a79fm05UQH4BTbwue+Q0zxi5p7UwVgqlqF/wM8dgk7L+KbIE6jUu
# 9Uqytb10Vxznr6VnSgMZscnJpqVBEe1arkNaNphQBR6z2DcycLLHBx1nbmBto1Bj
# LCWhbDzKHg==
# SIG # End signature block
