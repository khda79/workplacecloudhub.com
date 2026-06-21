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
        $adUser = if ($adUserByUpn.ContainsKey($userKey)) { $adUserByUpn[$userKey] } else { $null }
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
