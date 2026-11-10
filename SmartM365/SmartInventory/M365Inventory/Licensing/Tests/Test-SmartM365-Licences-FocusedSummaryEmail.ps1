<#
.SYNOPSIS
  Offline checks for the license overview and recovery email.
.VERSION
1.1.10
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$inventoryPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-Licences-Inventory.ps1')).Path
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($inventoryPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count -gt 0) { throw ($parseErrors | Out-String) }
$inventoryText = Get-Content -LiteralPath $inventoryPath -Raw
$coreModulePath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..\..\Modules\SmartM365.Core\SmartM365.Core.psd1')).Path
$coreModule = Import-Module -Name $coreModulePath -MinimumVersion '1.0.79' -Force -PassThru -ErrorAction Stop
# Workbook regression checks may load the installed module, but never repair/install it.
function Initialize-SmartM365ImportExcel {
  SmartM365.Core\Initialize-SmartM365ImportExcel -CheckOnly | Out-Null
}
if (-not (Get-Command -Name Complete-SmartM365CmdbSourceReceipt -Module $coreModule.Name -ErrorAction SilentlyContinue)) {
  throw 'SmartM365.Core must export the source receipt completion command used before the email.'
}
$reportIndex = $inventoryText.IndexOf('$licenseRunReport = Publish-LicensesRunReport -Folder $LatestCsvFolderPath')
$completeIndex = $inventoryText.IndexOf("Complete-SmartM365CmdbSourceReceipt -Status 'Completed'")
$mailIndex = $inventoryText.IndexOf('[void](Send-LicensesFocusedSummaryEmail -CsvFolderPath $LatestCsvFolderPath')
if ($reportIndex -lt 0 -or $completeIndex -lt 0 -or $mailIndex -lt 0 -or
    $reportIndex -gt $completeIndex -or $completeIndex -gt $mailIndex -or
    $inventoryText.IndexOf('Licensing source receipt could not be completed; summary email was not sent') -lt 0) {
  throw 'The full collector must publish its report, then qualify its source receipt, before sending the email.'
}
$names = @(
  'Get-LicensesSuiteProduct', 'Get-LicensesOptionalProperty', 'Get-LicensesFocusedSummaryRows', 'Get-LicensesAdditionalOverviewRows', 'Get-LicensesOtherProduct', 'New-LicensesOverviewCardHtml',
  'ConvertTo-LicensesActivityDate', 'ConvertTo-LicensesMailboxSizeGb', 'Get-LicensesQualifiedPartialAdCoverage', 'Get-LicensesCsvSource', 'Get-LicensesPublicationPath',
  'Import-LicensesSourceCsv', 'Read-LicensesIndexedSource', 'Get-LicensesMailboxGapSummary', 'Get-LicensesAdAccountActivitySummary', 'Get-LicensesFocusedUsageRows', 'Get-LicensesSharedConversionBatch', 'ConvertTo-LicensesSignInUtc', 'Get-LicensesSharedConversionDetails', 'Get-LicensesSharedConversionCounts', 'Import-LicensesSourceCsvCached', 'Get-LicensesUserAdScopes', 'Get-LicensesUserScope', 'Get-LicensesScopeBreakdown', 'New-LicensesScopeTablesHtml', 'New-LicensesSummaryMailContent', 'Send-LicensesSummaryMailScope', 'ConvertTo-LicensesMailSchedule', 'Get-LicensesMailSchedule', 'Get-LicensesDomainMailSettings', 'Get-LicensesDomainMailStatePath', 'Get-LicensesDomainPublishedReport', 'Format-LicensesMetric',
  'Get-LicensesAlertSummary', 'New-LicensesRecoveryWorkbook', 'Publish-LicensesReportCsv', 'Get-LicensesStandalonePlan2Stock', 'Get-LicensesEquivalentPlan2Stock', 'Get-LicensesSharedMigrationEvidence', 'Publish-LicensesReportSnapshot', 'Write-LicensesDailyMailState', 'Enter-LicensesDailyMailGate',
  'Send-LicensesFocusedSummaryEmail', 'Read-LicensesTenantSnapshot', 'ConvertFrom-LicensesJsonElement', 'ConvertTo-LicensesCountTable',
  'Read-LicensesPublishedReport'
)
$definitions = @($ast.FindAll({
  param($node)
  $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $names -contains $node.Name
}, $true))
if ($definitions.Count -ne $names.Count) { throw 'Focused summary functions are missing.' }
foreach ($definition in $definitions) { Invoke-Expression $definition.Extent.Text }

function Assert-Equal {
  param($Actual, $Expected, [string]$Label)
  if ($Actual -ne $Expected) { throw "$Label expected '$Expected', got '$Actual'." }
}
$testStockSku = [pscustomobject]@{ TenantSkuPartNumber='EXCHANGEENTERPRISE'; Id='sku-plan2'; CapabilityStatus='Enabled'; TenantPrepaidEnabled='12'; TenantConsumedUnits='7' }
$testStock = Get-LicensesStandalonePlan2Stock -TenantRows @($testStockSku) -CollectedAtUtc '2026-10-06T10:00:00Z'
Assert-Equal $testStock.Available 5 'Standalone Plan 2 stock'
Assert-Equal $testStock.CollectedAtUtc '2026-10-06T10:00:00Z' 'Stock keeps the collection timestamp'
Assert-Equal (Get-LicensesStandalonePlan2Stock -TenantRows @() -CollectedAtUtc '2026-10-06T10:00:00Z') $null 'Absent SKU is unknown'
Assert-Equal (Get-LicensesStandalonePlan2Stock -TenantRows @($testStockSku,$testStockSku) -CollectedAtUtc '2026-10-06T10:00:00Z') $null 'Duplicate SKU is unknown'
$testStockSku.TenantConsumedUnits='13'
Assert-Equal (Get-LicensesStandalonePlan2Stock -TenantRows @($testStockSku) -CollectedAtUtc '2026-10-06T10:00:00Z') $null 'Negative availability is unknown'
$testStockSku.TenantConsumedUnits='7'; $testStockSku.CapabilityStatus='Suspended'
Assert-Equal (Get-LicensesStandalonePlan2Stock -TenantRows @($testStockSku) -CollectedAtUtc '2026-10-06T10:00:00Z') $null 'Suspended SKU is unknown'
$equivalentSku=[pscustomobject]@{Id='sku-e3';TenantSkuPartNumber='SPE_E3';CapabilityStatus='Enabled';TenantPrepaidEnabled='12';TenantConsumedUnits='7'}
$unknownSku=[pscustomobject]@{Id='sku-unknown';TenantSkuPartNumber='OTHER';CapabilityStatus='Enabled';TenantPrepaidEnabled='4';TenantConsumedUnits='1'}
$plan=[pscustomobject]@{TenantKey='prod';SkuId='sku-e3';PlanName='EXCHANGE_S_ENTERPRISE'}
$equivalent=Get-LicensesEquivalentPlan2Stock -TenantRows @($equivalentSku,$unknownSku,$testStockSku) -CatalogRows @($plan) -TenantKey 'prod'
Assert-Equal @($equivalent.Rows).Count 1 'Equivalent SKU counted once and standalone excluded'
Assert-Equal $equivalent.Rows[0].Available 5 'Equivalent SKU available units'
Assert-Equal $equivalent.UnknownSkuCount 1 'Missing catalog entry stays unknown'
function Get-ScriptLocalConfigValue {
  param($Config, [string]$Name, $DefaultValue)
  if ($Config.ContainsKey($Name)) { return $Config[$Name] }
  return $DefaultValue
}
function Test-SmartM365MaxItemsMode { return $script:Sampled }
function WriteLog { param([string]$Message, [string]$Level) }
function Send-SmartM365Mail {
  param([string]$From, [string]$To, [string]$Cc, [string]$ExpectedSingleRecipient, [string]$Subject, [string]$BodyHtml, [string]$MailPurpose, [string[]]$Attachments, [switch]$AllowAttachments, [switch]$SuppressAttachmentLinks)
  if ($script:FailMail) { throw 'Simulated mail transport failure.' }
  if (-not $AllowAttachments -or -not $SuppressAttachmentLinks -or $Attachments.Count -ne 1 -or -not (Test-Path -LiteralPath $Attachments[0])) { throw 'Recovery workbook was not attached privately.' }
  $bookSummary = @(Import-Excel -Path $Attachments[0] -WorksheetName Overview -StartRow 14 -EndRow 18 -StartColumn 2 -EndColumn 8)
  $bookCandidates = @(Import-Excel -Path $Attachments[0] -WorksheetName 'Recovery candidates' -WarningAction SilentlyContinue)
  $bookDowngrade = @(Import-Excel -Path $Attachments[0] -WorksheetName 'E3 to F3 review' -WarningAction SilentlyContinue)
  $bookDates = @{}
  $package = Open-ExcelPackage -Path $Attachments[0] -ErrorAction Stop
  try {
    $sheet = $package.Workbook.Worksheets['Recovery candidates']
    for ($row=2; $row -le $sheet.Dimension.End.Row; $row++) {
      $bookDates[[string]$sheet.Cells[$row,2].Text] = [pscustomobject]@{
        AdValue = $sheet.Cells[$row,6].Value
        AdFormat = $sheet.Cells[$row,6].Style.Numberformat.Format
        M365Value = $sheet.Cells[$row,7].Value
        M365Format = $sheet.Cells[$row,7].Style.Numberformat.Format
      }
    }
  }
  finally { $package.Dispose() }
  $script:SentMail.Add([pscustomobject]@{From=$From;To=$To;Cc=$Cc;CcBound=$PSBoundParameters.ContainsKey('Cc');ExpectedSingleRecipient=$ExpectedSingleRecipient;Subject=$Subject;BodyHtml=$BodyHtml;MailPurpose=$MailPurpose;BookSummary=$bookSummary;BookCandidates=$bookCandidates;BookDowngrade=$bookDowngrade;BookDates=$bookDates;AttachmentPath=$Attachments[0]}) | Out-Null
}
function Update-TestLicensesReceipt {
  param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string[]]$CsvPaths)
  # Records the report CSVs as optional files of the fixture licensing receipt, as Complete-SmartM365CmdbSourceReceipt does.
  $path = Join-Path $Folder 'SmartInventory_SmartM365-Licences-Inventory.current.json.txt'
  $receipt = if (Test-Path -LiteralPath $path) { Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -AsHashtable } else { @{Status='Completed';IsPartialInventory=$false;Files=@()} }
  $names = @($CsvPaths | ForEach-Object { [IO.Path]::GetFileName($_) })
  $files = @(@($receipt.Files) | Where-Object { $_ -and [string]$_.File -notin $names })
  foreach ($csv in $CsvPaths) {
    $files += @{File=[IO.Path]::GetFileName($csv);Status='Success';IsPartialInventory=$false;Required=$false;SHA256=(Get-FileHash -LiteralPath $csv -Algorithm SHA256).Hash}
  }
  $receipt.Files = $files
  $receipt.TenantKey = 'prod'
  $receipt.Producer = 'SmartM365-Licences-Inventory.ps1'
  if (-not $receipt.Status) { $receipt.Status = 'Completed' }
  $receipt | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $path
}
function Update-TestUserActivityReceipt {
  param([Parameter(Mandatory)][string]$Folder)
  # Licences uses a usage report only when the completed UserActivity receipt lists it with the same hash.
  $files = foreach ($name in @('M365_Users_Activity.csv','M365_Mailbox_Usage.csv','M365_Email_Activity.csv','M365_Apps_Usage_180D.csv','M365_OneDrive_Usage.csv')) {
    $csv = Join-Path $Folder $name
    if (Test-Path -LiteralPath $csv) { @{File=$name;Status='Success';IsPartialInventory=$false;Required=($name -eq 'M365_Users_Activity.csv');SHA256=(Get-FileHash -LiteralPath $csv -Algorithm SHA256).Hash} }
  }
  @{Status='Completed';IsPartialInventory=$false;TenantKey='prod';Producer='SmartM365-M365UserActivity-Inventory.ps1';Files=@($files)} | ConvertTo-Json -Depth 6 |
    Set-Content -LiteralPath (Join-Path $Folder 'SmartInventory_SmartM365-M365UserActivity-Inventory.current.json.txt')
}
function Publish-TestLicenseReport {
  param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][object[]]$TenantRows, [Parameter(Mandatory)][string]$CollectedAtUtc,
        [switch]$ForceAdCsvAnalysis, [switch]$WithoutUsage)
  # Publishes the report from completed fixture receipts, as the collector does before its receipt completes.
  $summaryRows = @(Get-LicensesFocusedSummaryRows -TenantRows $TenantRows)
  $additionalRows = @(Get-LicensesAdditionalOverviewRows -TenantRows $TenantRows)
  $usage = $null
  if (-not $WithoutUsage) {
    try {
      $usage = Get-LicensesFocusedUsageRows -CsvFolderPath $Folder -ExpectedTenantKey 'prod' -LicenseSnapshotUtc ([datetimeoffset]::Parse($CollectedAtUtc,[Globalization.CultureInfo]::InvariantCulture)) -ForceAdCsvAnalysis:$ForceAdCsvAnalysis
    }
    catch { $usage = $null }
  }
  $snapshot = Publish-LicensesReportSnapshot -Folder $Folder -TenantKey 'prod' -CollectedAtUtc $CollectedAtUtc -SummaryRows $summaryRows -AdditionalRows $additionalRows -Usage $usage -AssignmentIssues $null
  Update-TestLicensesReceipt -Folder $Folder -CsvPaths $snapshot.CsvPaths
  return $snapshot
}
function New-SmartM365EmailBody {
  param([string]$Title, [string]$Category, [string]$HostName, [string]$GeneratedAt, [string]$BodyHtml, [string]$Footer)
  $script:TemplateCalls.Add([pscustomobject]@{HostName=$HostName;GeneratedAt=$GeneratedAt;Footer=$Footer}) | Out-Null
  return "<!-- SmartM365EmailTemplate:v1 --><html><body><header>$Title</header><main>$BodyHtml</main><footer>$Footer</footer></body></html>"
}

$script:OrgDomain = 'example.invalid'
$script:ScriptLocalConfig = @{To='reports@example.invalid';From='sender@example.invalid';EnableLicenseSummaryEmail=$true}
$script:Sampled = $false
$script:FailMail = $false
$script:SentMail = [System.Collections.Generic.List[object]]::new()
$script:TemplateCalls = [System.Collections.Generic.List[object]]::new()
$mailGateRoot = Join-Path $env:TEMP ('SmartM365-LicenseMailGate-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $mailGateRoot -Force | Out-Null
$script:MailStatePath = Join-Path $mailGateRoot 'M365_Licenses_SummaryEmail_SendState.json.txt'
$basicReportRoot = Join-Path $mailGateRoot 'basic-report'
New-Item -ItemType Directory -Path $basicReportRoot -Force | Out-Null
$rows = @(
  [pscustomobject]@{TenantSkuPartNumber='M365_F1';TenantPrepaidEnabled=10;TenantConsumedUnits=8}
  [pscustomobject]@{TenantSkuPartNumber='M365_F1_COMM';TenantPrepaidEnabled=5;TenantConsumedUnits=4}
  [pscustomobject]@{TenantSkuPartNumber='SPE_F1';TenantPrepaidEnabled=20;TenantConsumedUnits=17}
  [pscustomobject]@{TenantSkuPartNumber='SPE_E3';TenantPrepaidEnabled=30;TenantConsumedUnits=28}
  [pscustomobject]@{TenantSkuPartNumber='Microsoft_365_Copilot';TenantPrepaidEnabled=10;TenantConsumedUnits=8}
  [pscustomobject]@{TenantSkuPartNumber='VIRTUAL_AGENT_USL';TenantPrepaidEnabled=1;TenantConsumedUnits=1}
  [pscustomobject]@{TenantSkuPartNumber='DYN365_FINANCE';TenantPrepaidEnabled=20;TenantConsumedUnits=12}
  [pscustomobject]@{TenantSkuPartNumber='Dyn365_Operations_Activity';TenantPrepaidEnabled=5;TenantConsumedUnits=3}
  [pscustomobject]@{TenantSkuPartNumber='Dynamics_365_for_Operations_Sandbox_Tier2_SKU';TenantPrepaidEnabled=50;TenantConsumedUnits=0}
  [pscustomobject]@{TenantSkuPartNumber='PROJECT_MADEIRA_PREVIEW_IW_SKU';TenantPrepaidEnabled=10000;TenantConsumedUnits=1}
  [pscustomobject]@{TenantSkuPartNumber='POWER_BI_PRO';TenantPrepaidEnabled=10;TenantConsumedUnits=7}
  [pscustomobject]@{TenantSkuPartNumber='PBI_PREMIUM_PER_USER';TenantPrepaidEnabled=4;TenantConsumedUnits=4}
  [pscustomobject]@{TenantSkuPartNumber='POWER_BI_STANDARD';TenantPrepaidEnabled=1000000;TenantConsumedUnits=1000}
  [pscustomobject]@{TenantSkuPartNumber='OTHER_SKU';TenantPrepaidEnabled=99;TenantConsumedUnits=99}
)
$basicCollectedAt = '2026-10-06T10:00:00.0000000+00:00'
$rows | ForEach-Object {
  [pscustomobject]@{ TenantKey='prod'; TenantSkuPartNumber=$_.TenantSkuPartNumber; TenantPrepaidEnabled=$_.TenantPrepaidEnabled
    TenantConsumedUnits=$_.TenantConsumedUnits; CollectedAtUtc=$basicCollectedAt }
} | Export-Csv -LiteralPath (Join-Path $basicReportRoot 'M365_Licenses_Tenant.csv') -NoTypeInformation
$summary = @(Get-LicensesFocusedSummaryRows -TenantRows $rows)
Assert-Equal $summary.Count 4 'Product count'
Assert-Equal $summary[0].Enabled 15 'F1 enabled sum'
Assert-Equal $summary[0].Consumed 12 'F1 consumed sum'
Assert-Equal $summary[1].Product 'Microsoft 365 F3' 'F3 mapping'
Assert-Equal $summary[1].Enabled 20 'F3 enabled count'
Assert-Equal $summary[2].Enabled 30 'E3 enabled count'
Assert-Equal $summary[3].Subscribed $false 'Missing E5 status'
Assert-Equal $summary[3].Enabled 0 'Missing E5 enabled count'
$additional = @(Get-LicensesAdditionalOverviewRows -TenantRows $rows)
Assert-Equal $additional.Count 3 'Additional product count'
Assert-Equal $additional[0].Product 'Microsoft 365 Copilot' 'Copilot family'
Assert-Equal $additional[0].Enabled 10 'Copilot excludes Studio licenses'
Assert-Equal $additional[0].Consumed 8 'Copilot used units'
Assert-Equal $additional[1].Enabled 25 'Dynamics user units exclude sandbox and preview'
Assert-Equal $additional[1].Consumed 15 'Dynamics used units'
Assert-Equal $additional[2].Enabled 14 'Power BI Pro and Premium Per User units exclude free Standard'
Assert-Equal $additional[2].Consumed 11 'Power BI used units'
if ((New-LicensesOverviewCardHtml -Row $additional[1] -Width 33 -Accent '#7c3aed') -notmatch 'font-size:26px;line-height:30px;[^>]*>60%</div>.*<strong>15</strong> used of <strong>25</strong> enabled') {
  throw 'Dynamics overview card must emphasize the utilization percentage above the license counts.'
}
if ((New-LicensesOverviewCardHtml -Row $summary[3] -Width 25 -Accent '#475569') -notmatch 'font-size:26px;line-height:30px;[^>]*>N/A</div>.*Not subscribed') {
  throw 'An unsubscribed overview card must show N/A as its primary value.'
}

Publish-TestLicenseReport -Folder $basicReportRoot -TenantRows $rows -CollectedAtUtc $basicCollectedAt -WithoutUsage | Out-Null
Send-LicensesFocusedSummaryEmail -CsvFolderPath $basicReportRoot -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath | Out-Null
Assert-Equal $script:SentMail.Count 1 'Sent mail count'
$sentState = Get-Content -LiteralPath $script:MailStatePath -Raw | ConvertFrom-Json
Assert-Equal $sentState.Status 'Sent' 'Daily send state'
Send-LicensesFocusedSummaryEmail -CsvFolderPath $basicReportRoot -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath | Out-Null
Assert-Equal $script:SentMail.Count 1 'Second send on the same Paris day is skipped'
Assert-Equal $script:SentMail[0].BookSummary.Count 4 'Workbook summary has four target suites'
Assert-Equal $script:SentMail[0].BookCandidates.Count 0 'Workbook with unavailable sources has no invented candidates'
Assert-Equal (Test-Path -LiteralPath $script:SentMail[0].AttachmentPath) $false 'Temporary attachment removed after send'
Assert-Equal $script:SentMail[0].MailPurpose 'Report' 'Mail purpose'
Assert-Equal $script:SentMail[0].To 'reports@example.invalid' 'Report recipient'
if ($script:SentMail[0].BodyHtml -notmatch '(?s)License alerts.*>N/D</div>.*Awaiting capacity.*F1 N/D · F3 N/D · E3 N/D · E5 N/D.*>N/D</div>.*Licensing errors.*F1 N/D · F3 N/D · E3 N/D · E5 N/D') {
  throw 'Unqualified assignment sources must show N/D in both top alert cards.'
}
foreach ($product in @('F1','F3','E3','E5')) {
  if ($script:SentMail[0].BodyHtml -notmatch ('>' + $product + '</td>')) { throw "Missing product '$product' in mail body." }
}
if ($script:SentMail[0].BodyHtml -like '*OTHER_SKU*' -or $script:SentMail[0].BodyHtml -match '>\s*99\s*<') {
  throw 'Unrelated SKU leaked into the focused summary email.'
}

$script:Sampled = $true
Send-LicensesFocusedSummaryEmail -CsvFolderPath $basicReportRoot -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath | Out-Null
Assert-Equal $script:SentMail.Count 1 'Sampled run sends no mail'
$script:Sampled = $false
$script:ScriptLocalConfig.EnableLicenseSummaryEmail = $false
Send-LicensesFocusedSummaryEmail -CsvFolderPath $basicReportRoot -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath | Out-Null
Assert-Equal $script:SentMail.Count 1 'Disabled summary sends no mail'
$script:Sampled = $true
Send-LicensesFocusedSummaryEmail -CsvFolderPath $basicReportRoot -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath -Manual -ForceLicenseSummaryEmail | Out-Null
Assert-Equal $script:SentMail.Count 2 'Explicit email-only mode sends mail'
if ($script:SentMail[1].Subject -ne 'Microsoft 365 license overview and recovery' -or $script:SentMail[1].BodyHtml -notlike '*No new inventory was run*') {
  throw 'Email-only mode did not identify the existing CSV snapshot.'
}
if ($script:SentMail[1].BodyHtml.IndexOf('Source: existing published CSV') -le $script:SentMail[1].BodyHtml.IndexOf('Source freshness') -or
    $script:SentMail[1].BodyHtml.IndexOf('Source: existing published CSV') -ge $script:SentMail[1].BodyHtml.IndexOf('<footer>Host:')) {
  throw 'Existing CSV provenance is not at the bottom of the email content.'
}
$script:Sampled = $false

$dayStatePath = Join-Path $mailGateRoot 'synthetic-day-state.json.txt'
$dayStart = [datetimeoffset]::Parse('2026-10-24T22:30:00Z')
$pendingGate = Enter-LicensesDailyMailGate -Path $dayStatePath -TenantKey 'prod' -NowUtc $dayStart
Assert-Equal $pendingGate.Day '2026-10-25' 'Paris calendar day'
$pendingGate.Stream.Dispose()
$pendingRetry = Enter-LicensesDailyMailGate -Path $dayStatePath -TenantKey 'prod' -NowUtc $dayStart
Assert-Equal $pendingRetry.Skip $true 'Pending delivery prevents an automatic duplicate'
$forcedGate = Enter-LicensesDailyMailGate -Path $dayStatePath -TenantKey 'prod' -NowUtc $dayStart -Force
Assert-Equal $forcedGate.Skip $false 'Force switch bypasses the daily limit'
Write-LicensesDailyMailState -Stream $forcedGate.Stream -Bytes ([Text.Encoding]::UTF8.GetBytes('{"TenantKey":"prod","LocalDate":"2026-10-25","Status":"Sent"}'))
$forcedGate.Stream.Dispose()
$nextDayGate = Enter-LicensesDailyMailGate -Path $dayStatePath -TenantKey 'prod' -NowUtc ([datetimeoffset]::Parse('2026-10-25T23:30:00Z'))
Assert-Equal $nextDayGate.Skip $false 'New Paris day allows another send'
$nextDayGate.Stream.Dispose()

# Weekly schedule (1.64): one email per Monday-to-Sunday Paris week, due from WeekDay, caught up later in the week.
function Set-SyntheticSent([string]$Path, [string]$Day) {
  [IO.File]::WriteAllText($Path, ('{{"TenantKey":"prod","LocalDate":"{0}","Status":"Sent"}}' -f $Day))
}
function Enter-WeeklyGate([string]$Path, [string]$NowUtc, [DayOfWeek]$WeekDay = [DayOfWeek]::Monday, [switch]$IgnoreWeekDay) {
  $gate = Enter-LicensesDailyMailGate -Path $Path -TenantKey 'prod' -NowUtc ([datetimeoffset]::Parse($NowUtc)) -Frequency Weekly -WeekDay $WeekDay -IgnoreWeekDay:$IgnoreWeekDay
  if ($gate.Stream) { $gate.Stream.Dispose() }
  return $gate
}
$weekPath = Join-Path $mailGateRoot 'synthetic-week-state.json.txt'
$mondayGate = Enter-WeeklyGate $weekPath '2026-10-26T08:00:00Z'
Assert-Equal $mondayGate.Skip $false 'Weekly email is sent on Monday'
Assert-Equal $mondayGate.Day '2026-10-26' 'Weekly gate records the Paris day'
$weeklyPending = Enter-WeeklyGate $weekPath '2026-10-28T08:00:00Z'
Assert-Equal $weeklyPending.Skip $true 'Pending weekly delivery prevents a duplicate later in the week'
Set-SyntheticSent $weekPath '2026-10-26'
Assert-Equal (Enter-WeeklyGate $weekPath '2026-10-27T08:00:00Z').Skip $true 'Weekly email is not sent again on Tuesday'
Assert-Equal (Enter-WeeklyGate $weekPath '2026-11-01T20:00:00Z').Skip $true 'Weekly email is not sent again on Sunday evening'
Assert-Equal (Enter-WeeklyGate $weekPath '2026-11-01T23:30:00Z').Skip $false 'Paris Monday 00:30 opens the next week'
Set-SyntheticSent $weekPath '2026-10-26'
$forcedWeekly = Enter-LicensesDailyMailGate -Path $weekPath -TenantKey 'prod' -NowUtc ([datetimeoffset]::Parse('2026-10-27T08:00:00Z')) -Frequency Weekly -Force
Assert-Equal $forcedWeekly.Skip $false 'Force bypasses the weekly limit'
$forcedWeekly.Stream.Dispose()
$catchUpPath = Join-Path $mailGateRoot 'synthetic-catchup-state.json.txt'
Set-SyntheticSent $catchUpPath '2026-10-19'
Assert-Equal (Enter-WeeklyGate $catchUpPath '2026-10-28T08:00:00Z').Skip $false 'A missed Monday is caught up on Wednesday'
$dueDayPath = Join-Path $mailGateRoot 'synthetic-dueday-state.json.txt'
$notDue = Enter-WeeklyGate $dueDayPath '2026-10-27T08:00:00Z' -WeekDay Wednesday
Assert-Equal $notDue.Skip $true 'Weekly email waits for its week day'
if ($notDue.Reason -notmatch 'due from Wednesday') { throw "Not-due reason is unclear: $($notDue.Reason)" }
if (Test-Path -LiteralPath $dueDayPath) { throw 'A not-due weekly email must not write a send state.' }
Assert-Equal (Enter-WeeklyGate $dueDayPath '2026-10-27T08:00:00Z' -WeekDay Wednesday -IgnoreWeekDay).Skip $false 'Manual runs do not wait for the week day'
Set-SyntheticSent $dueDayPath '2026-10-20'
Assert-Equal (Enter-WeeklyGate $dueDayPath '2026-10-29T08:00:00Z' -WeekDay Wednesday).Skip $false 'Weekly email is sent after its week day'
$dailyPath = Join-Path $mailGateRoot 'synthetic-daily-state.json.txt'
Set-SyntheticSent $dailyPath '2026-10-26'
$dailyGate = Enter-LicensesDailyMailGate -Path $dailyPath -TenantKey 'prod' -NowUtc ([datetimeoffset]::Parse('2026-10-27T08:00:00Z')) -Frequency Daily
Assert-Equal $dailyGate.Skip $false 'Daily schedule still sends the next day'
$dailyGate.Stream.Dispose()

$defaultSchedule = Get-LicensesMailSchedule -Config @{}
Assert-Equal $defaultSchedule.Frequency 'Weekly' 'Default frequency'
Assert-Equal $defaultSchedule.WeekDay ([DayOfWeek]::Monday) 'Default week day'
$dailySchedule = Get-LicensesMailSchedule -Config @{LicenseSummaryEmailFrequency='Daily';LicenseSummaryEmailWeekDay='friday'}
Assert-Equal $dailySchedule.Frequency 'Daily' 'Configured frequency'
Assert-Equal $dailySchedule.WeekDay ([DayOfWeek]::Friday) 'Configured week day is case-insensitive'
foreach ($badSchedule in @(@{LicenseSummaryEmailFrequency='Monthly'}, @{LicenseSummaryEmailWeekDay='lundi'}, @{LicenseSummaryEmailWeekDay='1'})) {
  $threw = $false
  try { Get-LicensesMailSchedule -Config $badSchedule | Out-Null } catch { $threw = $true }
  if (-not $threw) { throw "Invalid schedule was accepted: $($badSchedule | ConvertTo-Json -Compress)" }
}
$domainSchedules = @(Get-LicensesDomainMailSettings -Config @{LicenseSummaryEmailWeekDay='Tuesday';LicenseSummaryDomainMails=@(
  @{Domain='a.example.invalid';To='a@example.invalid'}
  @{Domain='b.example.invalid';To='b@example.invalid';Frequency='Daily'}
  @{Domain='c.example.invalid';To='c@example.invalid';WeekDay='Thursday'})})
Assert-Equal "$($domainSchedules[0].Schedule.Frequency)/$($domainSchedules[0].Schedule.WeekDay)" 'Weekly/Tuesday' 'Domain entry inherits the tenant schedule'
Assert-Equal $domainSchedules[1].Schedule.Frequency 'Daily' 'Domain entry overrides the frequency'
Assert-Equal "$($domainSchedules[2].Schedule.Frequency)/$($domainSchedules[2].Schedule.WeekDay)" 'Weekly/Thursday' 'Domain entry overrides the week day'
$threw = $false
try { Get-LicensesDomainMailSettings -Config @{LicenseSummaryDomainMails=@(@{Domain='a.example.invalid';To='a@example.invalid';Frequency='Hourly'})} | Out-Null } catch { $threw = $true }
if (-not $threw) { throw 'Invalid domain frequency was accepted.' }

$badRows = @([pscustomobject]@{TenantSkuPartNumber='SPE_E5';TenantPrepaidEnabled=$null;TenantConsumedUnits=1})
$threw = $false
try { Get-LicensesFocusedSummaryRows -TenantRows $badRows | Out-Null } catch { $threw = $true }
Assert-Equal $threw $true 'Missing count is rejected'
$failedPreparationPath = Join-Path $mailGateRoot 'failed-preparation.json.txt'
$threw = $false
$unpublishedRoot = Join-Path $mailGateRoot 'unpublished-report'
New-Item -ItemType Directory -Path $unpublishedRoot -Force | Out-Null
try { Send-LicensesFocusedSummaryEmail -CsvFolderPath $unpublishedRoot -ExpectedTenantKey 'prod' -MailStatePath $failedPreparationPath -Manual | Out-Null } catch { $threw = $true }
Assert-Equal $threw $true 'Missing published report is reported'
Assert-Equal (Get-Item -LiteralPath $failedPreparationPath).Length 0 'Preparation failure leaves no daily send marker'
Send-LicensesFocusedSummaryEmail -CsvFolderPath $basicReportRoot -ExpectedTenantKey 'prod' -MailStatePath $failedPreparationPath -Manual | Out-Null
Assert-Equal $script:SentMail.Count 3 'Preparation failure can be retried without force'

$uncertainPath = Join-Path $mailGateRoot 'uncertain-send.json.txt'
$script:FailMail = $true
$threw = $false
try { Send-LicensesFocusedSummaryEmail -CsvFolderPath $basicReportRoot -ExpectedTenantKey 'prod' -MailStatePath $uncertainPath -Manual | Out-Null } catch { $threw = $true }
$script:FailMail = $false
Assert-Equal $threw $true 'Transport failure is reported'
Assert-Equal (Get-Content -LiteralPath $uncertainPath -Raw | ConvertFrom-Json).Status 'Pending' 'Uncertain delivery blocks automatic retry'
Send-LicensesFocusedSummaryEmail -CsvFolderPath $basicReportRoot -ExpectedTenantKey 'prod' -MailStatePath $uncertainPath -Manual | Out-Null
Assert-Equal $script:SentMail.Count 3 'Uncertain delivery is not automatically repeated'
Send-LicensesFocusedSummaryEmail -CsvFolderPath $basicReportRoot -ExpectedTenantKey 'prod' -MailStatePath $uncertainPath -Manual -ForceLicenseSummaryEmail | Out-Null
Assert-Equal $script:SentMail.Count 4 'Explicit force retries uncertain delivery'

$adRecentRow = [pscustomobject]@{Enabled='True';LastLogonDate=([datetime]::UtcNow.AddDays(-2).ToString('yyyy-MM-dd'))}
$adOldRow = [pscustomobject]@{Enabled='True';LastLogonDate=([datetime]::UtcNow.AddDays(-100).ToString('yyyy-MM-dd'))}
$adDisabledRow = [pscustomobject]@{Enabled='False';LastLogonDate=''}
$adNoDateRow = [pscustomobject]@{Enabled='True';LastLogonDate=''}
$section06Members = @{
  'a'=[pscustomobject]@{OnPremisesImmutableId='imm-a';'User principal name'='a@example.invalid'}
  'b'=[pscustomobject]@{OnPremisesImmutableId='';'User principal name'='b@example.invalid'}
  'c'=[pscustomobject]@{OnPremisesImmutableId='';'User principal name'='c@example.invalid'}
  'd'=[pscustomobject]@{OnPremisesImmutableId='';'User principal name'='d@example.invalid'}
  'e'=[pscustomobject]@{OnPremisesImmutableId='';'User principal name'='e@example.invalid'}
  'f'=[pscustomobject]@{OnPremisesImmutableId='';'User principal name'='f@example.invalid'}
}
$adGap = Get-LicensesAdAccountActivitySummary `
  -MailboxGap ([pscustomobject]@{NoUserMailbox=[pscustomobject]@{Available=$true;Members=$section06Members}}) `
  -AdSource ([pscustomobject]@{Ready=$true;Forced=$false;Reason=''}) `
  -AdByImmutable @{'imm-a'=$adRecentRow} `
  -AdByUpn @{'a@example.invalid'=$adRecentRow;'b@example.invalid'=$adOldRow;'c@example.invalid'=$adDisabledRow;'f@example.invalid'=$adNoDateRow} `
  -DuplicateAdUpns @{'e@example.invalid'=$true} -Cutoff ([datetime]::UtcNow.Date.AddDays(-90))
Assert-Equal $adGap.Members 6 'Section 07 AD member population'
Assert-Equal $adGap.AdObserved 4 'Unique AD matches'
Assert-Equal $adGap.AdEnabledRecent 1 'Recent AD logon'
Assert-Equal $adGap.AdEnabledInactive 1 'Inactive AD logon'
Assert-Equal $adGap.AdEnabledNoDate 1 'Missing AD logon date'
Assert-Equal $adGap.AdDisabled 1 'Disabled AD account'
Assert-Equal $adGap.NoObservedMatch 1 'Unmatched AD account kept separate'
Assert-Equal $adGap.Ambiguous 1 'Ambiguous AD account kept separate'

$testRoot = Join-Path $env:TEMP ('SmartM365-LicenseSummary-' + [guid]::NewGuid().ToString('N'))
try {
  New-Item -ItemType Directory -Path $testRoot -Force | Out-Null
  $csvPath = Join-Path $testRoot 'M365_Licenses_Tenant.csv'
  @(
    [pscustomobject]@{TenantKey='prod';TenantSkuPartNumber='SPE_F1';TenantPrepaidEnabled=20;TenantConsumedUnits=17;CollectedAtUtc='2026-10-06T10:00:00Z'}
    [pscustomobject]@{TenantKey='prod';TenantSkuPartNumber='SPE_E3';TenantPrepaidEnabled=30;TenantConsumedUnits=28;CollectedAtUtc='2026-10-06T10:00:00Z'}
  ) | Export-Csv -LiteralPath $csvPath -NoTypeInformation
  $snapshot = Read-LicensesTenantSnapshot -Path $csvPath -ExpectedTenantKey 'prod'
  Assert-Equal $snapshot.Rows.Count 2 'Existing CSV row count'
  Assert-Equal $snapshot.CollectedAtUtc '2026-10-06T10:00:00.0000000+00:00' 'Existing CSV collection time'
  $threw = $false
  try { Read-LicensesTenantSnapshot -Path $csvPath -ExpectedTenantKey 'test' | Out-Null } catch { $threw = $true }
  Assert-Equal $threw $true 'Cross-tenant CSV is rejected'
  $csvRows = @(Import-Csv -LiteralPath $csvPath)
  $csvRows[1].CollectedAtUtc = '2026-10-05T10:00:00Z'
  $csvRows | Export-Csv -LiteralPath $csvPath -NoTypeInformation
  $threw = $false
  try { Read-LicensesTenantSnapshot -Path $csvPath -ExpectedTenantKey 'prod' | Out-Null } catch { $threw = $true }
  Assert-Equal $threw $true 'Mixed snapshot timestamps are rejected'
  $tenantCollectedAt = [datetimeoffset]::UtcNow.ToString('o')
  $rows | ForEach-Object {
    [pscustomobject]@{ TenantKey='prod'; TenantSkuPartNumber=$_.TenantSkuPartNumber;
      TenantPrepaidEnabled=$_.TenantPrepaidEnabled; TenantConsumedUnits=$_.TenantConsumedUnits;
      CollectedAtUtc=$tenantCollectedAt }
  } | Export-Csv -LiteralPath $csvPath -NoTypeInformation

  $today = [datetime]::UtcNow.Date
  $recent = $today.AddDays(-2).ToString('yyyy-MM-dd')
  $old = $today.AddDays(-100).ToString('yyyy-MM-dd')
  $adRecent = $today.AddDays(-2).ToString('dd/MM/yyyy HH:mm:ss')
  $adOld = $today.AddDays(-100).ToString('dd/MM/yyyy HH:mm:ss')
  $refresh = $today.AddDays(-1).ToString('yyyy-MM-dd')
  @(
    [pscustomobject]@{TenantKey='prod';UserId='u1';SkuPartNumber='M365_F1'}
    [pscustomobject]@{TenantKey='prod';UserId='u1';SkuPartNumber='M365_F1_COMM'}
    [pscustomobject]@{TenantKey='prod';UserId='u2';SkuPartNumber='SPE_F1'}
    [pscustomobject]@{TenantKey='prod';UserId='u2';SkuPartNumber='VISIOCLIENT'}
    [pscustomobject]@{TenantKey='prod';UserId='u3';SkuPartNumber='SPE_E3'}
    [pscustomobject]@{TenantKey='prod';UserId='u3';SkuPartNumber='POWER_BI_PRO'}
    [pscustomobject]@{TenantKey='prod';UserId='u4';SkuPartNumber='SPE_E5'}
    [pscustomobject]@{TenantKey='prod';UserId='u4';SkuPartNumber='SPE_F1'}
    [pscustomobject]@{TenantKey='prod';UserId='u4';SkuPartNumber='SPE_F1'}
    [pscustomobject]@{TenantKey='prod';UserId='u5';SkuPartNumber='M365_F1_COMM'}
    [pscustomobject]@{TenantKey='prod';UserId='u6';SkuPartNumber='SPE_E3'}
    [pscustomobject]@{TenantKey='prod';UserId='u6';SkuPartNumber='POWER_BI_PRO'}
    [pscustomobject]@{TenantKey='prod';UserId='u7';SkuPartNumber='SPE_E3'}
    [pscustomobject]@{TenantKey='prod';UserId='u8';SkuPartNumber='SPE_E3'}
    [pscustomobject]@{TenantKey='prod';UserId='u9';SkuPartNumber='POWER_BI_PRO'}
    [pscustomobject]@{TenantKey='prod';UserId='u11';SkuPartNumber='VISIOCLIENT'}
  ) | Export-Csv -LiteralPath (Join-Path $testRoot 'M365_Licenses_Users.csv') -NoTypeInformation
  $licenseUsersPath = Join-Path $testRoot 'M365_Licenses_Users.csv'
  $licenseManifestPath = Join-Path $testRoot 'SmartInventory_SmartM365-Licences-Inventory.current.json.txt'
  $licenseHash = (Get-FileHash -LiteralPath $licenseUsersPath -Algorithm SHA256).Hash
  @{Status='Completed';IsPartialInventory=$false;Files=@(@{File='M365_Licenses_Users.csv';Status='Success';IsPartialInventory=$false;SHA256=$licenseHash})} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $licenseManifestPath
  $localGuid = [guid]::Parse('00112233-4455-6677-8899-aabbccddeeff')
  $localImmutable = [Convert]::ToBase64String($localGuid.ToByteArray())
  @(
    [pscustomobject]@{TenantKey='prod';'Object Id'='u1';'User principal name'='u1@example.invalid';UserType='Member';AccountEnabled='False';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=''}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u2';'User principal name'='u2@example.invalid';UserType='Member';AccountEnabled='True';OnPremisesImmutableId='a2';LastSuccessfulSignInDateTime=$old}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u3';'User principal name'='u3@example.invalid';UserType='Member';AccountEnabled='True';OnPremisesImmutableId='a3';LastSuccessfulSignInDateTime=$recent}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u4';'User principal name'='u4@example.invalid';UserType='Member';AccountEnabled='True';OnPremisesImmutableId='a4';LastSuccessfulSignInDateTime=$old}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u6';'User principal name'='u6@example.invalid';UserType='Member';AccountEnabled='True';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=$recent}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u7';'User principal name'='u7@example.invalid';UserType='Member';AccountEnabled='False';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=''}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u8';'User principal name'='u8@example.invalid';UserType='Member';AccountEnabled='False';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=''}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u9';'User principal name'='u9@example.invalid';UserType='Member';AccountEnabled='False';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=''}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u10';'User principal name'='u10@example.invalid';UserType='Member';AccountEnabled='True';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=''}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u11';'User principal name'='u11@example.invalid';UserType='Guest';AccountEnabled='True';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=''}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u12';'User principal name'='u12@example.invalid';UserType='Member';AccountEnabled='False';OnPremisesImmutableId=$localImmutable;LastSuccessfulSignInDateTime=''}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u13';'User principal name'='u13@example.invalid';UserType='Member';AccountEnabled='False';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=''}
    [pscustomobject]@{TenantKey='prod';'Object Id'='u15';'User principal name'='u15@example.invalid';UserType='Member';AccountEnabled='False';OnPremisesImmutableId='';LastSuccessfulSignInDateTime=''}
  ) | Export-Csv -LiteralPath (Join-Path $testRoot 'M365_Users_Active.csv') -NoTypeInformation
  $onPremPath = Join-Path $testRoot 'Exchange_OnPrem_Mailboxes_AllDomains.csv'
  @([pscustomobject]@{TenantKey='prod';ObjectGUID=$localGuid.ToString();UserPrincipalName='u12@example.invalid';RecipientType='UserMailbox';NativeIdentityStatus='Observed'}) |
    Export-Csv -LiteralPath $onPremPath -NoTypeInformation
  $onPremManifestPath = Join-Path $testRoot 'SmartInventory_SmartM365-Exchange-Local-Mailboxes-Inventory.current.json.txt'
  $onPremHash = (Get-FileHash -LiteralPath $onPremPath -Algorithm SHA256).Hash
  @{Status='Completed';IsPartialInventory=$false;ConsumerScopeQualified=$true;Files=@(@{File='Exchange_OnPrem_Mailboxes_AllDomains.csv';Status='Success';IsPartialInventory=$false;SHA256=$onPremHash})} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $onPremManifestPath
  @(
    [pscustomobject]@{TenantKey='prod';ImmutableId_AD='a2';UserPrincipalName='u2@example.invalid';Enabled='True';LastLogonDate=$adOld}
    [pscustomobject]@{TenantKey='prod';ImmutableId_AD='a3';UserPrincipalName='u3@example.invalid';Enabled='True';LastLogonDate=$adRecent}
    [pscustomobject]@{TenantKey='prod';ImmutableId_AD='a4';UserPrincipalName='u4@example.invalid';Enabled='True';LastLogonDate=$adOld}
  ) | Export-Csv -LiteralPath (Join-Path $testRoot 'AD_Users_AllDomains.csv') -NoTypeInformation
  @(
    [pscustomobject]@{TenantKey='prod';UserPrincipalName='u2@example.invalid';ReportPeriod='D180';ReportRefreshDate=$refresh;LastActivityDate=$old;IsDeleted='False'}
    [pscustomobject]@{TenantKey='prod';UserPrincipalName='u3@example.invalid';ReportPeriod='D180';ReportRefreshDate=$refresh;LastActivityDate=$recent;IsDeleted='False'}
    [pscustomobject]@{TenantKey='prod';UserPrincipalName='u4@example.invalid';ReportPeriod='D180';ReportRefreshDate=$refresh;LastActivityDate=$old;IsDeleted='False'}
    [pscustomobject]@{TenantKey='prod';UserPrincipalName='u6@example.invalid';ReportPeriod='D180';ReportRefreshDate=$refresh;LastActivityDate=$old;IsDeleted='False'}
  ) | Export-Csv -LiteralPath (Join-Path $testRoot 'M365_Users_Activity.csv') -NoTypeInformation
  @(
    [pscustomobject]@{TenantKey='prod';'User Principal Name'='u2@example.invalid';'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'='';'Is Deleted'='False'}
    [pscustomobject]@{TenantKey='prod';'User Principal Name'='u3@example.invalid';'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'=$recent;'Is Deleted'='False'}
    [pscustomobject]@{TenantKey='prod';'User Principal Name'='u4@example.invalid';'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'='';'Is Deleted'='False'}
  ) | Export-Csv -LiteralPath (Join-Path $testRoot 'M365_Mailbox_Usage.csv') -NoTypeInformation
  @(
    [pscustomobject]@{TenantKey='prod';'User Principal Name'='u2@example.invalid';'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'='';'Is Deleted'='False';'Send Count'=0;'Read Count'=0}
    [pscustomobject]@{TenantKey='prod';'User Principal Name'='u3@example.invalid';'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'=$recent;'Is Deleted'='False';'Send Count'=1;'Read Count'=0}
    [pscustomobject]@{TenantKey='prod';'User Principal Name'='u4@example.invalid';'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'='';'Is Deleted'='False';'Send Count'=0;'Read Count'=0}
  ) | Export-Csv -LiteralPath (Join-Path $testRoot 'M365_Email_Activity.csv') -NoTypeInformation
  $appsPath = Join-Path $testRoot 'M365_Apps_Usage_180D.csv'
  @(
    [pscustomobject]@{TenantKey='prod';'User Principal Name'='u3@example.invalid';'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'=$recent;Windows='Yes';Mac='No'}
    [pscustomobject]@{TenantKey='prod';'User Principal Name'='u4@example.invalid';'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'=$old;Windows='No';Mac='No'}
    [pscustomobject]@{TenantKey='prod';'User Principal Name'='u6@example.invalid';'Report Period'='180';'Report Refresh Date'=$refresh;'Last Activity Date'=$recent;Windows='No';Mac='No'}
  ) | Export-Csv -LiteralPath $appsPath -NoTypeInformation
  $sharedPath = Join-Path $testRoot 'Exchange_EXO_Mailboxes_AllDomains.csv'
  @(
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u1';RecipientTypeDetails='SharedMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='12,5';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u2';RecipientTypeDetails='UserMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u3';RecipientTypeDetails='SharedMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='49,99';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u4';RecipientTypeDetails='SharedMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='12,0';ArchiveStatus='Active';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u6';RecipientTypeDetails='UserMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='1,4';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u7';RecipientTypeDetails='UserMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u8';RecipientTypeDetails='SharedMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='12,0';ArchiveStatus='Active';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u9';RecipientTypeDetails='UserMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u10';RecipientTypeDetails='UserMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u13';RecipientTypeDetails='SharedMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='1';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
    [pscustomobject]@{TenantKey='prod';ExternalDirectoryObjectId='u15';RecipientTypeDetails='RoomMailbox';NativeIdentityStatus='Observed';TotalItemSizeGB='1';ArchiveStatus='None';LitigationHoldEnabled='False';RetentionHoldEnabled='False'}
  ) | Export-Csv -LiteralPath $sharedPath -NoTypeInformation

  $oneDrivePath = Join-Path $testRoot 'M365_OneDrive_Usage.csv'
  @(
    [pscustomobject]@{TenantKey='prod';'Owner Principal Name'='u6@example.invalid';'Is Deleted'='False';'Storage Used (Byte)'='1073741824';'Report Refresh Date'=$refresh;'Report Period'='180'}
  ) | Export-Csv -LiteralPath $oneDrivePath -NoTypeInformation
  Update-TestUserActivityReceipt -Folder $testRoot

  $intunePath = Join-Path $testRoot 'Intune_Devices_Inventory.csv'
  @(
    [pscustomobject]@{TenantKey='prod';'Device ID'='pc-u2-a';OS='Windows';UserId='u2';'Primary user UPN'='u2@example.invalid'}
    [pscustomobject]@{TenantKey='prod';'Device ID'='pc-u2-b';OS='Windows';UserId='u2';'Primary user UPN'='u2@example.invalid'}
    [pscustomobject]@{TenantKey='prod';'Device ID'='pc-u3';OS='Windows';UserId='u3';'Primary user UPN'='u3@example.invalid'}
    [pscustomobject]@{TenantKey='prod';'Device ID'='phone-u1';OS='iOS';UserId='u1';'Primary user UPN'='u1@example.invalid'}
    [pscustomobject]@{TenantKey='prod';'Device ID'='pc-u4';OS='Windows';UserId='u4';'Primary user UPN'='u4@example.invalid'}
  ) | Export-Csv -LiteralPath $intunePath -NoTypeInformation

  $usage = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $usage.MailboxGap.UserMailboxes.Total 2 'User mailboxes without target suite'
  Assert-Equal $usage.MailboxGap.UserMailboxes.Universe 5 'Qualified EXO UserMailbox denominator'
  Assert-Equal $usage.MailboxGap.UserMailboxes.OtherSkus 1 'User mailbox with another SKU'
  Assert-Equal $usage.MailboxGap.UserMailboxes.NoSkus 1 'User mailbox with no SKU'
  Assert-Equal $usage.MailboxGap.UserMailboxes.EntraStateAvailable $true 'User mailbox Entra state is qualified'
  Assert-Equal $usage.MailboxGap.UserMailboxes.EntraEnabled 1 'Enabled Entra account with unlicensed UserMailbox'
  Assert-Equal $usage.MailboxGap.UserMailboxes.EntraDisabled 1 'Disabled Entra account with unlicensed UserMailbox'
  Assert-Equal $usage.MailboxGap.UserMailboxes.EntraStateUnknown 0 'Unlicensed UserMailbox Entra state reconciles'
  $activePath = Join-Path $testRoot 'M365_Users_Active.csv'
  $activeRows = @(Import-Csv -LiteralPath $activePath)
  $candidateWithoutState = @($activeRows | Where-Object { $_.'Object Id' -eq 'u10' })[0]
  $candidateWithoutState.AccountEnabled = ''
  $activeRows | Export-Csv -LiteralPath $activePath -NoTypeInformation
  $unknownState = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $unknownState.MailboxGap.UserMailboxes.Total 2 'Unknown Entra state retains UserMailbox total'
  Assert-Equal $unknownState.MailboxGap.UserMailboxes.EntraStateUnknown 1 'Blank Entra state is N/D'
  $candidateWithoutState.AccountEnabled = 'True'
  $activeRows | Export-Csv -LiteralPath $activePath -NoTypeInformation
  Move-Item -LiteralPath $activePath -Destination ($activePath + '.missing')
  try {
    $missingEntra = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
    Assert-Equal $missingEntra.MailboxGap.UserMailboxes.Total 2 'Missing Entra source retains UserMailbox total'
    Assert-Equal $missingEntra.MailboxGap.UserMailboxes.EntraStateAvailable $false 'Missing Entra source leaves state N/D'
  }
  finally { Move-Item -LiteralPath ($activePath + '.missing') -Destination $activePath }
  Assert-Equal $usage.MailboxGap.NoUserMailbox.Total 1 'On-premises UserMailbox is excluded from mailbox-free accounts'
  Assert-Equal $usage.MailboxGap.NoUserMailbox.Universe 7 'Qualified nontechnical Entra account denominator'
  Assert-Equal $usage.MailboxGap.NoUserMailbox.OtherSkus 1 'Mailbox-free account with another SKU'
  Assert-Equal $usage.MailboxGap.NoUserMailbox.NoSkus 0 'On-premises mailbox account with no SKU is excluded'
  Assert-Equal $usage.MailboxGap.NoUserMailbox.Guests 1 'Guest is identified within the total'
  Assert-Equal $usage.MailboxGap.NoUserMailbox.MemberEnabled 0 'Enabled member segment'
  Assert-Equal $usage.MailboxGap.NoUserMailbox.MemberDisabled 0 'On-premises mailbox disabled member is excluded'
  $localAccount = @($activeRows | Where-Object { $_.'Object Id' -eq 'u12' })[0]
  $localAccount.OnPremisesImmutableId = ''
  $activeRows | Export-Csv -LiteralPath $activePath -NoTypeInformation
  $upnFallback = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $upnFallback.MailboxGap.NoUserMailbox.Total 1 'Unique UPN fallback excludes on-premises UserMailbox'
  $localAccount.OnPremisesImmutableId = $localImmutable
  $activeRows | Export-Csv -LiteralPath $activePath -NoTypeInformation
  $byProduct = @{}
  foreach ($item in $usage.Rows) { $byProduct[$item.Product] = $item.Counts }
  Assert-Equal $byProduct['Microsoft 365 F1'].Assigned 2 'F1 distinct users'
  Assert-Equal $byProduct['Microsoft 365 F1'].Multiple 0 'F1 variants count as one target suite'
  Assert-Equal $byProduct['Microsoft 365 F1'].MultipleAll 0 'Shared F1 mailbox is excluded from user overlap'
  Assert-Equal $byProduct['Microsoft 365 F1'].Disabled 0 'Disabled shared mailbox is excluded from disabled users'
  Assert-Equal $byProduct['Microsoft 365 F1'].DisabledUnknown 1 'F1 unmatched account unknown'
  Assert-Equal $byProduct['Microsoft 365 F1'].SharedEligible 1 'F1 disabled shared mailbox under 50 GB'
  Assert-Equal $byProduct['Microsoft 365 F3'].Assigned 2 'F3 duplicate assignment path deduplicated'
  Assert-Equal $byProduct['Microsoft 365 F3'].Multiple 0 'Shared F3 mailbox is excluded from target-suite overlap'
  Assert-Equal $byProduct['Microsoft 365 F3'].MultipleAll 1 'F3 non-shared user with an add-on SKU'
  Assert-Equal $byProduct['Microsoft 365 F3'].AdEntraInactive 1 'Shared mailbox is excluded from AD and Entra inactivity'
  Assert-Equal $byProduct['Microsoft 365 F3'].MailboxInactive 1 'Shared mailbox is excluded from mailbox inactivity'
  Assert-Equal $byProduct['Microsoft 365 F3'].M365Inactive 1 'Shared mailbox is excluded from M365 inactivity'
  Assert-Equal $byProduct['Microsoft 365 F3'].SharedUnder50 1 'F3 licensed shared mailbox under 50 GB'
  Assert-Equal $byProduct['Microsoft 365 F3'].SharedEligible 0 'F3 archived shared mailbox excluded'
  Assert-Equal $byProduct['Microsoft 365 F3'].RecoveryCandidates 1 'Archived shared mailbox excluded from recovery union'
  Assert-Equal $byProduct['Microsoft 365 F3'].RecoveryPrimaryPc 1 'Two Windows PCs count once for a recovery candidate'
  Assert-Equal $byProduct['Microsoft 365 F1'].RecoveryPrimaryPc 0 'An iOS device is not an Intune PC'
  Assert-Equal $byProduct['Microsoft 365 E3'].LocalAppsInactive 1 'E3 local Apps inactive'
  Assert-Equal $byProduct['Microsoft 365 E3'].MailboxUnknown 1 'E3 missing mailbox row unknown'
  Assert-Equal $byProduct['Microsoft 365 E3'].M365Inactive 0 'Recent Apps activity prevents M365 false inactivity'
  Assert-Equal $byProduct['Microsoft 365 E3'].SharedEligible 1 'E3 active shared mailbox is a recovery candidate'
  Assert-Equal $byProduct['Microsoft 365 E3'].SharedLicensed 2 'Both E3 shared mailboxes remain in the dedicated shared table'
  Assert-Equal $byProduct['Microsoft 365 E3'].Disabled 1 'Only a disabled non-shared E3 account is counted'
  Assert-Equal $byProduct['Microsoft 365 E3'].RecoveryCandidates 2 'E3 shared and disabled user candidates are distinct'
  Assert-Equal $byProduct['Microsoft 365 E3'].RecoveryPrimaryPc 1 'Shared recovery candidate can be primary on an Intune PC'
  Assert-Equal $byProduct['Microsoft 365 E3'].Multiple 0 'E3 plus Power BI is one target suite'
  Assert-Equal $byProduct['Microsoft 365 E3'].MultipleAll 1 'Only non-shared E3 plus Power BI is counted'
  Assert-Equal $byProduct['Microsoft 365 E5'].Multiple 0 'Shared E5 mailbox is excluded from target-suite overlap'
  Assert-Equal $byProduct['Microsoft 365 E5'].LocalAppsInactive 0 'Shared E5 mailbox is excluded from local Apps inactivity'
  Assert-Equal $byProduct['Microsoft 365 E5'].RecoveryCandidates 0 'E5 archived shared mailbox is not recoverable'
  Assert-Equal $byProduct['Microsoft 365 E5'].RecoveryPrimaryPc 0 'A PC on a blocked shared mailbox is not a recovery intersection'
  Assert-Equal @($usage.RecoveryDetails).Count 4 'One detail row per qualified recovery candidate and target suite'
  Assert-Equal @($usage.RecoveryDetails | Where-Object { $_.License -eq 'Microsoft 365 F1' -and $_.UserId -eq 'u1' -and $_.RecoveryReason -eq 'Eligible shared mailbox under 50 GB' }).Count 1 'F1 shared mailbox recovery detail'
  Assert-Equal @($usage.RecoveryDetails | Where-Object { $_.License -eq 'Microsoft 365 F3' -and $_.UserId -eq 'u2' -and $_.RecoveryReason -eq 'No M365 activity in 90 days' -and $_.PrimaryOnIntuneWindowsPc -eq 'Yes' }).Count 1 'F3 inactive user and Intune detail'
  Assert-Equal @($usage.RecoveryDetails | Where-Object { $_.License -eq 'Microsoft 365 E3' -and $_.UserId -eq 'u7' -and $_.RecoveryReason -eq 'Disabled account' }).Count 1 'E3 disabled user recovery detail'
  $u2Detail = @($usage.RecoveryDetails | Where-Object { $_.UserId -eq 'u2' })[0]
  Assert-Equal $u2Detail.LastAdActivityDate $old 'F3 recovery detail has last AD logon date'
  Assert-Equal $u2Detail.LastM365ActivityDate $old 'F3 recovery detail has last M365 activity date'
  $u3Detail = @($usage.RecoveryDetails | Where-Object { $_.UserId -eq 'u3' })[0]
  Assert-Equal $u3Detail.LastAdActivityDate 'N/D' 'Shared mailbox has no AD activity date'
  Assert-Equal $u3Detail.LastM365ActivityDate 'N/D' 'Shared mailbox has no M365 user activity date'
  Assert-Equal $usage.DowngradeReview.Candidates 1 'One qualified E3 to F3 review candidate'
  Assert-Equal @($usage.DowngradeDetails).Count 1 'E3 to F3 detail count matches candidate KPI'
  Assert-Equal $usage.DowngradeDetails[0].UserId 'u6' 'E3 to F3 review lists the qualified user'
  Assert-Equal $usage.DowngradeDetails[0].MailboxSizeGB ([decimal]1.4) 'E3 to F3 detail keeps mailbox size'
  Assert-Equal $usage.DowngradeDetails[0].OneDriveUsedBytes 1073741824 'E3 to F3 detail keeps exact OneDrive storage'
  Assert-Equal $usage.DowngradeReview.RecoveryExcluded 2 'Recovery candidates excluded from downgrade review'
  Assert-Equal $usage.DowngradeReview.Unknown 0 'All E3 downgrade inputs are qualified'
  Assert-Equal $usage.DowngradeReview.Excluded 1 'Archived shared mailbox is excluded from downgrade review'
  $incompleteReview = $usage | Select-Object *
  $incompleteReview.DowngradeDetails = @()
  $reviewMismatchRejected = $false
  try {
    [void](New-LicensesRecoveryWorkbook -Path (Join-Path $testRoot 'IncompleteReview.xlsx') -SummaryRows $summary -Usage $incompleteReview -CollectedAtUtc ([datetimeoffset]::UtcNow.ToString('o')))
  }
  catch { $reviewMismatchRejected = $_.Exception.Message -like 'E3 to F3 review detail count differs from the email KPI*' }
  Assert-Equal $reviewMismatchRejected $true 'E3 to F3 workbook cannot omit email KPI candidates'
  Assert-Equal ((ConvertTo-LicensesActivityDate '01/09/2026 20:00:00').ToString('yyyy-MM-dd')) '2026-09-01' 'AD day/month parsing'
  Assert-Equal (ConvertTo-LicensesMailboxSizeGb '49,99') ([decimal]49.99) 'French mailbox size parsing'
  Assert-Equal (ConvertTo-LicensesMailboxSizeGb '50,00') ([decimal]50) '50 GB boundary parsing'

  $script:SentMail.Clear()
  $script:TemplateCalls.Clear()
  Publish-TestLicenseReport -Folder $testRoot -TenantRows $rows -CollectedAtUtc ([datetimeoffset]::UtcNow.ToString('o')) | Out-Null; Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath -Manual -ForceLicenseSummaryEmail | Out-Null
  Assert-Equal $script:SentMail.Count 1 'Enriched email sent'
  $reportSnapshotPath = Join-Path $testRoot 'M365_Licenses_ReportSnapshot.json.txt'
  $reportSnapshot = Get-Content -LiteralPath $reportSnapshotPath -Raw | ConvertFrom-Json
  Assert-Equal $reportSnapshot.TenantKey 'prod' 'Snapshot tenant'
  Assert-Equal @($reportSnapshot.Products).Count 4 'Snapshot target products'
  Assert-Equal @($reportSnapshot.OtherProducts).Count 3 'Snapshot paid overview products'
  Assert-Equal @($reportSnapshot.RecoveryCandidates).Count 4 'Snapshot recovery detail reconciles with workbook'
  Assert-Equal @($reportSnapshot.DowngradeCandidates).Count 1 'Snapshot downgrade detail reconciles with workbook'
  Assert-Equal @($reportSnapshot.Products | Where-Object Product -eq 'Microsoft 365 E3')[0].Counts.RecoveryCandidates 2 'Snapshot E3 recovery KPI'
  Assert-Equal @($reportSnapshot.Products | Where-Object Product -eq 'Microsoft 365 E3')[0].Counts.SharedE3Plan2Review 0 'Snapshot includes E3 Plan 2 review count'
  if ($script:SentMail[0].BodyHtml -notlike '*E3 to Exchange Online Plan 2 review: 0*') { throw 'Email omits the E3 Plan 2 review count.' }
  if ($script:SentMail[0].BodyHtml -notlike '*N/D available*') { throw 'Missing standalone Plan 2 stock must be displayed as N/D.' }
  $isolatedTestState = Join-Path $testRoot 'isolated-test-state.json.txt'
  [void](Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $isolatedTestState -Manual -TenantOnly -TestTo 'only@example.invalid')
  Assert-Equal $script:SentMail.Count 2 'Exactly one isolated test message'
  Assert-Equal $script:SentMail[1].To 'only@example.invalid' 'Isolated test recipient'
  Assert-Equal $script:SentMail[1].CcBound $true 'Isolated test explicitly overrides global Cc'
  Assert-Equal $script:SentMail[1].Cc '' 'Isolated test has no Cc'
  Assert-Equal $script:SentMail[1].ExpectedSingleRecipient 'only@example.invalid' 'Final transport route is constrained to the test recipient'
  if ($script:SentMail[1].Subject -notlike '[[]TEST]*') { throw 'Isolated test mail is not labelled.' }
  $script:SentMail.RemoveAt(1)
  $script:TemplateCalls.RemoveAt($script:TemplateCalls.Count - 1)
  Assert-Equal $reportSnapshot.E3ToF3Review.Candidates 1 'Snapshot E3 downgrade KPI'
  $reportSummaryCsv = @(Import-Csv -LiteralPath (Join-Path $testRoot 'M365_Licenses_ReportSummary.csv'))
  $reportCandidatesCsv = @(Import-Csv -LiteralPath (Join-Path $testRoot 'M365_Licenses_ReportCandidates.csv'))
  $reportGapsCsv = @(Import-Csv -LiteralPath (Join-Path $testRoot 'M365_Licenses_ReportGaps.csv'))
  $reportSourcesCsv = @(Import-Csv -LiteralPath (Join-Path $testRoot 'M365_Licenses_ReportSources.csv'))
  Assert-Equal $reportSummaryCsv.Count 7 'Power BI summary has all email products'
  $e3Csv = @($reportSummaryCsv | Where-Object Product -eq 'Microsoft 365 E3')[0]
  Assert-Equal $e3Csv.E3ToF3ReviewAssigned $reportSnapshot.E3ToF3Review.Assigned 'CSV E3 review assigned KPI'
  Assert-Equal $e3Csv.E3ToF3ReviewExcluded $reportSnapshot.E3ToF3Review.Excluded 'CSV E3 review excluded KPI'
  Assert-Equal $e3Csv.E3ToF3ReviewRecoveryExcluded $reportSnapshot.E3ToF3Review.RecoveryExcluded 'CSV E3 review already-in-recovery KPI'
  Assert-Equal $e3Csv.NoM365ActivityUnknown (@($reportSnapshot.Products | Where-Object Product -eq 'Microsoft 365 E3')[0].Counts.M365Unknown) 'CSV E3 unqualified activity KPI'
  Assert-Equal $e3Csv.AvailableUnits ([long]$e3Csv.EnabledUnits - [long]$e3Csv.ConsumedUnits) 'CSV available capacity KPI'
  Assert-Equal @($reportCandidatesCsv | Where-Object CandidateType -eq 'Recovery').Count 4 'Power BI recovery rows match the workbook'
  Assert-Equal @($reportCandidatesCsv | Where-Object CandidateType -eq 'E3 to F3 review').Count 1 'Power BI review rows match the workbook'
  Assert-Equal $reportGapsCsv.Count 25 'Power BI gaps have all mail indicators'
  Assert-Equal @($reportSourcesCsv | Where-Object Name -eq 'M365_Licenses_Tenant.csv').Count 1 'Power BI source freshness includes tenant source'
  foreach ($csvRow in @($reportSummaryCsv) + @($reportCandidatesCsv) + @($reportGapsCsv) + @($reportSourcesCsv)) {
    Assert-Equal $csvRow.SnapshotId $reportSnapshot.SnapshotId 'Power BI CSV snapshot identity'
    Assert-Equal $csvRow.TenantKey 'prod' 'Power BI CSV tenant identity'
  }
  if ($script:SentMail[0].BodyHtml -notmatch [regex]::Escape("Snapshot: $($reportSnapshot.SnapshotId)")) { throw 'Email does not identify its Power BI snapshot.' }
  Assert-Equal $script:SentMail[0].BookCandidates.Count 4 'Workbook row count matches recovery totals'
  Assert-Equal @($script:SentMail[0].BookCandidates | Where-Object License -eq 'Microsoft 365 E3').Count 2 'Workbook E3 rows match the email KPI'
  Assert-Equal @($script:SentMail[0].BookSummary | Where-Object License -eq 'Microsoft 365 E3')[0].'Recovery candidates' 2 'Workbook Summary E3 total'
  Assert-Equal @($script:SentMail[0].BookSummary | Where-Object License -eq 'Microsoft 365 E3')[0].'E3 to F3 review' 1 'Workbook Summary E3 review total'
  Assert-Equal $script:SentMail[0].BookDowngrade.Count 1 'Workbook E3 to F3 rows match the email KPI'
  Assert-Equal $script:SentMail[0].BookDowngrade[0].'User ID' 'u6' 'Workbook E3 to F3 review identity'
  Assert-Equal $script:SentMail[0].BookDowngrade[0].'Mailbox (GB)' ([decimal]1.4) 'Workbook E3 to F3 review mailbox size'
  Assert-Equal $script:SentMail[0].BookDowngrade[0].'OneDrive (bytes)' 1073741824 'Workbook E3 to F3 review exact OneDrive bytes'
  $bookU2 = $script:SentMail[0].BookDates['u2']
  Assert-Equal ([datetime]::FromOADate([double]$bookU2.AdValue)).ToString('yyyy-MM-dd') $old 'Workbook has AD activity date'
  Assert-Equal ([datetime]::FromOADate([double]$bookU2.M365Value)).ToString('yyyy-MM-dd') $old 'Workbook has M365 activity date'
  if ($script:SentMail[0].BookDates['u2'].AdValue -is [string] -or
      $script:SentMail[0].BookDates['u2'].M365Value -is [string] -or
      $null -eq $script:SentMail[0].BookDates['u2'].AdValue -or
      $null -eq $script:SentMail[0].BookDates['u2'].M365Value) {
    throw 'Workbook activity dates are stored as text or missing.'
  }
  Assert-Equal $script:SentMail[0].BookDates['u2'].AdFormat 'yyyy-mm-dd' 'AD date uses sortable Excel date format'
  Assert-Equal $script:SentMail[0].BookDates['u2'].M365Format 'yyyy-mm-dd' 'M365 date uses sortable Excel date format'
  Assert-Equal $script:SentMail[0].BookDates['u3'].AdValue $null 'Shared mailbox AD date cell is blank'
  Assert-Equal $script:SentMail[0].BookDates['u3'].M365Value $null 'Shared mailbox M365 date cell is blank'
  Assert-Equal (Test-Path -LiteralPath $script:SentMail[0].AttachmentPath) $false 'Temporary recovery workbook removed after send'
  if ($script:SentMail[0].BodyHtml -notlike '*Multiple assigned SKUs*' -or $script:SentMail[0].BodyHtml -notlike '*Multiple target suites*' -or $script:SentMail[0].BodyHtml -notlike '*Recovery candidates*' -or $script:SentMail[0].BodyHtml -notlike '*Removal candidates after archive and hold checks*') {
    throw 'Enriched KPI headers are missing from email.'
  }
  if ($script:SentMail[0].BodyHtml -notlike '*Candidates primary on Intune PC*' -or
      $script:SentMail[0].BodyHtml -notlike '*Intune_Devices_Inventory.csv*') {
    throw 'Intune PC recovery indicator or source freshness is missing.'
  }
  if ($script:SentMail[0].BodyHtml -notlike '*AD accounts and activity among section 07 Members*' -or
      $script:SentMail[0].BodyHtml -notlike '*No AD match means no unique match in the available AD export*') {
    throw 'AD account activity qualification is missing from section 07.'
  }
  if ($script:SentMail[0].BodyHtml -notlike '*03 &nbsp; E3 to F3 downgrade review*' -or
      $script:SentMail[0].BodyHtml -notlike '*E3 to F3 review tab listing every qualified review candidate*' -or
      $script:SentMail[0].BodyHtml -notlike '*OneDrive storage below 2 GB*' -or
      $script:SentMail[0].BodyHtml -notlike '*Activity dates are sortable Excel dates*') {
    throw 'Downgrade review or workbook activity dates are not described.'
  }
  $oneDriveRows = @(Import-Csv -LiteralPath $oneDrivePath)
  $oneDriveRows[0].'Storage Used (Byte)' = '2147483648'
  $oneDriveRows | Export-Csv -LiteralPath $oneDrivePath -NoTypeInformation
  Update-TestUserActivityReceipt -Folder $testRoot
  $oneDriveBoundary = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $oneDriveBoundary.DowngradeReview.Candidates 0 'OneDrive at 2 GB is excluded'
  Assert-Equal @($oneDriveBoundary.DowngradeDetails).Count 0 'Boundary user is absent from E3 to F3 detail'
  Assert-Equal $oneDriveBoundary.DowngradeReview.Excluded 2 'OneDrive at 2 GB is counted as excluded'
  $emptyReviewPath = Join-Path $testRoot 'NoE3ToF3Review.xlsx'
  [void](New-LicensesRecoveryWorkbook -Path $emptyReviewPath -SummaryRows $summary -Usage $oneDriveBoundary -CollectedAtUtc ([datetimeoffset]::UtcNow.ToString('o')))
  $emptyReviewRows = @(Import-Excel -Path $emptyReviewPath -WorksheetName 'E3 to F3 review' -WarningAction SilentlyContinue)
  Assert-Equal $emptyReviewRows.Count 0 'Empty E3 to F3 review sheet has no candidate rows'
  $emptyPackage = Open-ExcelPackage -Path $emptyReviewPath
  try { Assert-Equal $emptyPackage.Workbook.Worksheets['E3 to F3 review'].Cells[1,1].Text 'User ID' 'Empty E3 to F3 review sheet retains headers' }
  finally { $emptyPackage.Dispose() }
  Remove-Item -LiteralPath $emptyReviewPath -Force
  $oneDriveRows[0].'Storage Used (Byte)' = ''
  $oneDriveRows | Export-Csv -LiteralPath $oneDrivePath -NoTypeInformation
  Update-TestUserActivityReceipt -Folder $testRoot
  $oneDriveMissing = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $oneDriveMissing.DowngradeReview.Unknown 1 'Missing OneDrive size is N/D'
  $oneDriveRows[0].'Storage Used (Byte)' = '1073741824'
  $oneDriveRows | Export-Csv -LiteralPath $oneDrivePath -NoTypeInformation
  Update-TestUserActivityReceipt -Folder $testRoot
  if ($script:SentMail[0].BodyHtml -notmatch '<td style="[^"]*">E3</td><td style="[^"]*">2</td><td style="[^"]*">1</td>' -or
      $script:SentMail[0].BodyHtml -notlike '*Disabled users and activity/overlap indicators exclude identified shared mailboxes*') {
    throw 'The email does not separate disabled user accounts from shared mailboxes.'
  }
  if ($script:SentMail[0].BodyHtml -notmatch 'License overview and recovery' -or $script:SentMail[0].BodyHtml -notmatch 'F3/F1 adds both suite counts') {
    throw 'KPI banner and assignment-grain note are missing.'
  }
  $overviewBody = $script:SentMail[0].BodyHtml
  if ($overviewBody.IndexOf('>License overview</h2>') -lt 0 -or
      $overviewBody -notmatch '(?s)>License overview</h2>\s*<h3[^>]*>License alerts</h3>.*<h2[^>]*>License recovery overview</h2>' -or
      $overviewBody.IndexOf('>License recovery overview</h2>') -ge $overviewBody.IndexOf('MICROSOFT 365 SUITES') -or
      $overviewBody.IndexOf('Recovery candidates F3/F1') -ge $overviewBody.IndexOf('MICROSOFT 365 SUITES') -or
      $overviewBody.IndexOf('MICROSOFT 365 SUITES') -ge $overviewBody.IndexOf('COPILOT, DYNAMICS 365 AND POWER BI') -or
      $overviewBody.IndexOf('COPILOT, DYNAMICS 365 AND POWER BI') -ge $overviewBody.IndexOf('01 &nbsp; License capacity') -or
      $overviewBody -notlike '*MICROSOFT 365 SUITES*' -or
      $overviewBody -notlike '*COPILOT, DYNAMICS 365 AND POWER BI*') {
    throw 'License alerts and recovery KPIs are not above the license bands.'
  }
  foreach ($label in @('F1','F3','E3','E5','Copilot','Dynamics 365','Power BI')) {
    if ($overviewBody -notmatch ('>' + [regex]::Escape($label) + '</div>')) { throw "Missing overview card '$label'." }
  }
  if ($script:SentMail[0].BodyHtml -notmatch '>0</div><div[^>]*>Recovery candidates E5' -or
      $script:SentMail[0].BodyHtml -notmatch '>2</div><div[^>]*>Recovery candidates E3' -or
      $script:SentMail[0].BodyHtml -notmatch '>2 \(N/D: 1\)</div><div[^>]*>Recovery candidates F3/F1' -or
      $script:SentMail[0].BodyHtml -notmatch '>1</div><div[^>]*>E3 to F3 downgrade review') {
    throw 'Recovery and E3 downgrade KPI cards do not match the detailed sections.'
  }
  $body = $script:SentMail[0].BodyHtml
  if ($body.IndexOf('01 &nbsp; License capacity') -lt 0 -or
      $body.IndexOf('02 &nbsp; Recovery by license') -le $body.IndexOf('01 &nbsp; License capacity') -or
      $body.IndexOf('03 &nbsp; E3 to F3 downgrade review') -le $body.IndexOf('02 &nbsp; Recovery by license') -or
      $body.IndexOf('04 &nbsp; Activity and overlap') -le $body.IndexOf('03 &nbsp; E3 to F3 downgrade review') -or
      $body -notlike '*Consumed (used %)*' -or $body -notlike '*Available (free %)*') {
    throw 'License capacity order or percentages are missing.'
  }
  if ($body.IndexOf('06 &nbsp; User mailboxes without licence F1/F3/E3/E5') -le $body.IndexOf('05 &nbsp; Licensed shared mailboxes') -or
      $body.IndexOf('07 &nbsp; Without User mailboxes + without licence F1/F3/E3/E5') -le $body.IndexOf('06 &nbsp; User mailboxes without licence F1/F3/E3/E5') -or
      $body -notmatch '>2</strong> of 5 qualified EXO UserMailbox.*?<strong>40%</strong>' -or
      $body -notmatch 'Entra enabled</th>.*?Entra disabled</th>.*?Entra state N/D</th>.*?<td[^>]*>1</td><td[^>]*>1</td><td[^>]*>0</td>' -or
       $body -notmatch '>1</strong> of 7 qualified Entra accounts.*?<strong>14.3%</strong>' -or
       $body -notmatch 'Member enabled</th>.*?<td[^>]*>0</td><td[^>]*>0</td><td[^>]*>1</td><td[^>]*>1</td><td[^>]*>0</td><td[^>]*>0</td>') {
    throw 'Mailbox and account gap sections are missing or have incorrect counts.'
  }
  Assert-Equal $script:TemplateCalls.Count 1 'Focused email uses one shared template'
  Assert-Equal $script:TemplateCalls[0].HostName '' 'Host metadata removed from template header'
  Assert-Equal $script:TemplateCalls[0].GeneratedAt '' 'Generated metadata removed from template header'
  if ($script:TemplateCalls[0].Footer -notmatch '^Host: .+ \| Generated: \d{4}-\d{2}-\d{2}') { throw 'Host and generated time are missing from template footer.' }
  if ($body.IndexOf('<footer>Host:') -le $body.IndexOf('Source freshness')) { throw 'Host and generated time are not at the bottom of the email.' }
  if ($script:SentMail[0].BodyHtml -match '<table border="1"' -or $script:SentMail[0].BodyHtml -match '<th>License</th>') {
    throw 'Legacy unstyled table remains in email.'
  }
  if ($script:SentMail[0].BodyHtml -notmatch '<td style="[^"]*">F3</td><td style="[^"]*">20</td><td style="[^"]*">17 \(85%\)</td><td style="[^"]*">3 \(15%\)</td><td style="[^"]*">2</td><td style="[^"]*">Subscribed</td>') {
    throw 'F3 license capacity and percentages are missing from email.'
  }

  $mailboxRowsWithId = @(Import-Csv -LiteralPath $sharedPath)
  @($mailboxRowsWithId | Where-Object ExternalDirectoryObjectId -eq 'u10')[0].ExternalDirectoryObjectId = ''
  $mailboxRowsWithId | Export-Csv -LiteralPath $sharedPath -NoTypeInformation
  $unjoined = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $unjoined.MailboxGap.UserMailboxes.Unknown 1 'User mailbox without Entra ID is unqualified'
  Assert-Equal $unjoined.MailboxGap.NoUserMailbox.Available $false 'Missing mailbox identity prevents proving account has no mailbox'
  @($mailboxRowsWithId | Where-Object RecipientTypeDetails -eq 'UserMailbox' | Where-Object { -not $_.ExternalDirectoryObjectId })[0].ExternalDirectoryObjectId = 'u10'
  $mailboxRowsWithId | Export-Csv -LiteralPath $sharedPath -NoTypeInformation

  $intuneManifestPath = Join-Path $testRoot 'SmartInventory_SmartM365-Devices-Inventory.current.json.txt'
  @{Status='Completed';IsPartialInventory=$false;Files=@(@{File='Intune_Devices_Inventory.csv';Status='Failed';IsPartialInventory=$true})} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $intuneManifestPath
  $partialIntune = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $partialIntune.IntuneSourceReady $false 'Partial Intune file receipt is rejected'
  Assert-Equal @($partialIntune.Rows | Where-Object Product -eq 'Microsoft 365 F3')[0].Counts.RecoveryCandidates 1 'Intune source does not change recovery count'
  $script:SentMail.Clear()
  Publish-TestLicenseReport -Folder $testRoot -TenantRows $rows -CollectedAtUtc ([datetimeoffset]::UtcNow.ToString('o')) | Out-Null; Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath -Manual -ForceLicenseSummaryEmail | Out-Null
  if ($script:SentMail[0].BodyHtml -notmatch '<td style="[^"]*">F3</td><td style="[^"]*">1</td><td style="[^"]*">0</td><td style="[^"]*">1</td><td style="[^"]*">N/D</td>') {
    throw 'Unqualified Intune PC source is not marked N/D in recovery table.'
  }
  Remove-Item -LiteralPath $intuneManifestPath -Force

  $misaligned = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -LicenseSnapshotUtc ([datetimeoffset]::UtcNow.AddDays(-3)) -AsOfUtc $today
  $f3Misaligned = @($misaligned.Rows | Where-Object Product -eq 'Microsoft 365 F3')[0]
  Assert-Equal $f3Misaligned.Available $false 'Misaligned license snapshots are not qualified'
  Assert-Equal $misaligned.Sources[0].Ready $false 'Misaligned license user source is not ready'
  Assert-Equal $misaligned.MailboxGap $null 'Misaligned license source makes both new sections N/D'
  $script:SentMail.Clear()
  Publish-TestLicenseReport -Folder $testRoot -TenantRows $rows -CollectedAtUtc ([datetimeoffset]::UtcNow.AddDays(-3).ToString('o')) | Out-Null; Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath -Manual -ForceLicenseSummaryEmail | Out-Null
  $misalignedSnapshot = Get-Content -LiteralPath $reportSnapshotPath -Raw | ConvertFrom-Json
  $misalignedF3 = @($misalignedSnapshot.Products | Where-Object Product -eq 'Microsoft 365 F3')[0]
  Assert-Equal $misalignedF3.UsageAvailable $false 'Snapshot rejects misaligned license assignments'
  Assert-Equal $misalignedF3.Counts $null 'Snapshot leaves unqualified metrics null, not zero'
  $misalignedSummaryCsv = @(Import-Csv -LiteralPath (Join-Path $testRoot 'M365_Licenses_ReportSummary.csv'))
  $misalignedF3Csv = @($misalignedSummaryCsv | Where-Object Product -eq 'Microsoft 365 F3')[0]
  Assert-Equal $misalignedF3Csv.EvidenceStatus 'N/D' 'CSV marks unqualified usage N/D'
  Assert-Equal $misalignedF3Csv.RecoveryCandidates '' 'CSV leaves unqualified recovery count blank'
  Assert-Equal $script:SentMail.Count 1 'Unqualified usage still sends a stock summary'
  if ($script:SentMail[0].BodyHtml -notlike '*N/D*') { throw 'Unqualified usage is not marked N/D in email.' }

  $sharedRows = @(Import-Csv -LiteralPath $sharedPath)
  $u4Shared = @($sharedRows | Where-Object ExternalDirectoryObjectId -eq 'u4')[0]
  $u4Shared.ArchiveStatus = 'None'
  $u4Shared.LitigationHoldEnabled = 'True'
  $sharedRows | Export-Csv -LiteralPath $sharedPath -NoTypeInformation
  $held = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  $heldE5 = @($held.Rows | Where-Object Product -eq 'Microsoft 365 E5')[0]
  Assert-Equal $heldE5.Counts.SharedEligible 0 'Litigation hold blocks shared recovery'
  Assert-Equal $heldE5.Counts.RecoveryCandidates 0 'Litigation hold blocks recovery union'

  $u4Shared.LitigationHoldEnabled = 'False'
  $u4Shared.TotalItemSizeGB = ''
  $sharedRows | Export-Csv -LiteralPath $sharedPath -NoTypeInformation
  $unknownSize = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  $unknownE5 = @($unknownSize.Rows | Where-Object Product -eq 'Microsoft 365 E5')[0]
  Assert-Equal $unknownE5.Counts.SharedUnknown 1 'Missing shared mailbox size is unknown'
  Assert-Equal $unknownE5.Counts.RecoveryUnknown 1 'Missing size cannot become a recovery candidate'

  $u3Shared = @($sharedRows | Where-Object ExternalDirectoryObjectId -eq 'u3')[0]
  $u3Shared.TotalItemSizeGB = '50,00'
  $sharedRows | Export-Csv -LiteralPath $sharedPath -NoTypeInformation
  $boundary = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  $boundaryE3 = @($boundary.Rows | Where-Object Product -eq 'Microsoft 365 E3')[0]
  Assert-Equal $boundaryE3.Counts.SharedEligible 0 '50 GB shared mailbox is excluded'
  Assert-Equal $boundaryE3.Counts.RecoveryCandidates 1 '50 GB shared mailbox is excluded while disabled user remains'
  Assert-Equal $boundaryE3.Counts.SharedE3Plan2Review 0 'Rounded 50 GB boundary is not counted as above 50 GB'

  $u3Shared.TotalItemSizeGB = '75,00'
  $sharedRows | Export-Csv -LiteralPath $sharedPath -NoTypeInformation
  $large = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal @($large.Rows | Where-Object Product -eq 'Microsoft 365 E3')[0].Counts.SharedE3Plan2Review 1 'Large shared mailbox with E3 requires Plan 2 replacement review'
  Assert-Equal @($large.Rows | Where-Object Product -eq 'Microsoft 365 F3')[0].Counts.SharedE3Plan2Review 0 'F3 assignment is not an E3 replacement'
  $u3Shared.TotalItemSizeGB = '100,00'
  $sharedRows | Export-Csv -LiteralPath $sharedPath -NoTypeInformation
  $upperBoundary = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal @($upperBoundary.Rows | Where-Object Product -eq 'Microsoft 365 E3')[0].Counts.SharedE3Plan2Review 0 'Rounded 100 GB boundary is unqualified'

  $u3Shared.TotalItemSizeGB = '49,99'
  $u4Shared.TotalItemSizeGB = '12,0'
  $u4Shared.ArchiveStatus = 'Active'
  $sharedRows | Export-Csv -LiteralPath $sharedPath -NoTypeInformation
  $exoManifestPath = Join-Path $testRoot 'SmartInventory_SmartM365-EXO-Mailboxes-Inventory.current.json.txt'
  @{Status='Completed';IsPartialInventory=$false;Files=@(@{File='Exchange_EXO_Mailboxes_AllDomains.csv';Status='Failed';IsPartialInventory=$true})} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $exoManifestPath
  $partialExo = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $partialExo.SharedSourceReady $false 'Partial EXO file receipt is rejected'
  Assert-Equal $partialExo.MailboxGap.UserMailboxes.Available $false 'Partial EXO source makes user mailbox gap N/D'
  Assert-Equal $partialExo.MailboxGap.NoUserMailbox.Available $false 'Partial EXO source makes no-mailbox gap N/D'
  $partialE3 = @($partialExo.Rows | Where-Object Product -eq 'Microsoft 365 E3')[0]
  Assert-Equal $partialE3.Counts.RecoveryCandidates 0 'Unavailable EXO source cannot yield recovery candidates'
  Assert-Equal $partialE3.Counts.RecoveryUnknown 4 'Unavailable EXO source makes all E3 mailbox types unknown'
  Assert-Equal $partialE3.Counts.MultipleAllUnknown 4 'Unqualified mailbox type is excluded from overlap counts'
  Remove-Item -LiteralPath $exoManifestPath -Force

  $adManifestPath = Join-Path $testRoot 'SmartInventory_SmartM365-ActiveDirectory-Inventory.current.json.txt'
  $qualifiedAdPath = Join-Path $testRoot 'AD_Users_AllDomains.csv'
  $originalAdCsv = Get-Content -LiteralPath $qualifiedAdPath -Raw
  $coverage = @{
    Kind='ADDomainCoverage'; Status='PartialAccepted'; Reason='NonBlockingDomainErrors'
    ExpectedDomains=@('observed.example.invalid','allowed.example.invalid','unavailable.example.invalid')
    CollectedDomains=@('observed.example.invalid','allowed.example.invalid')
    UnavailableDomains=@('unavailable.example.invalid')
    NonBlockingDomainErrors=@('ALLOWED.example.invalid','UNAVAILABLE.example.invalid')
  }
  $scope = 'CMDB:ad_users,ad_computers,ad_domains,ad_groups,ad_objects,ad_members'
  $qualifiedAdRows = @(Import-Csv -LiteralPath $qualifiedAdPath)
  foreach ($adRow in $qualifiedAdRows) { $adRow | Add-Member -NotePropertyName DomainName -NotePropertyValue 'allowed.example.invalid' -Force }
  $qualifiedAdRows | Export-Csv -LiteralPath $qualifiedAdPath -NoTypeInformation
  $qualifiedProof = @{
    Producer='SmartM365-ActiveDirectory-Inventory.ps1'; Status='Completed'; Errors=0; Scope=$scope
    ScopeQualification='ConsumerScope'; ConsumerScopeQualified=$true; FullInventoryQualified=$false; IsPartialInventory=$true
    RequiredFiles=@('AD_Users_AllDomains.csv'); DomainCoverage=$coverage
    Files=@(@{File='AD_Users_AllDomains.csv'; Status='Success'; Errors=0; Scope=$scope; IsPartialInventory=$true
      SHA256=(Get-FileHash -LiteralPath $qualifiedAdPath).Hash; DomainCoverage=$coverage})
  }
  try {
    $qualifiedProof | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $adManifestPath
    $strictSource = Get-LicensesCsvSource -Folder $testRoot -FileName 'AD_Users_AllDomains.csv' -Columns @('TenantKey','DomainName') -AsOfUtc $today -CollectorManifestName 'SmartInventory_SmartM365-ActiveDirectory-Inventory.current.json.txt'
    Assert-Equal $strictSource.Ready $false 'Partial acceptance is opt-in for the AD consumer only'
    $acceptedAd = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
    $acceptedSource = @($acceptedAd.Sources | Where-Object Name -eq 'AD_Users_AllDomains.csv')[0]
    Assert-Equal $acceptedSource.Ready $true 'Qualified partial AD source is usable'
    Assert-Equal $acceptedSource.PartialScopeAccepted $true 'Partial AD scope remains recorded'
    Assert-Equal $acceptedSource.Forced $false 'Qualified partial AD is not a receipt bypass'
    Assert-Equal $acceptedAd.AdSourceForced $false 'Partial acceptance does not enable ForceAdCsvAnalysis'
    Assert-Equal @($acceptedAd.Rows | Where-Object Product -eq 'Microsoft 365 F3')[0].Counts.AdEntraInactive 1 'Available allowlisted domain remains included'
    if ($acceptedSource.Reason -notlike '*2/3 domains collected*unavailable.example.invalid*') { throw 'Accepted AD coverage limitation is missing.' }
    Assert-Equal $acceptedAd.MailboxGap.NoUserMailbox.Total $usage.MailboxGap.NoUserMailbox.Total 'AD coverage does not alter mailbox-free population'
    $partialSummary = Get-LicensesAdAccountActivitySummary -MailboxGap ([pscustomobject]@{NoUserMailbox=[pscustomobject]@{Available=$true;Members=$section06Members}}) `
      -AdSource $acceptedSource -AdByImmutable @{'imm-a'=$adRecentRow} `
      -AdByUpn @{'a@example.invalid'=$adRecentRow;'b@example.invalid'=$adOldRow;'c@example.invalid'=$adDisabledRow;'f@example.invalid'=$adNoDateRow} `
      -DuplicateAdUpns @{'e@example.invalid'=$true} -Cutoff $today.AddDays(-90)
    Assert-Equal $partialSummary.Provisional $true 'Partial AD counts are provisional'
    Assert-Equal $partialSummary.AdObserved 4 'Observed AD matches remain counted'
    Assert-Equal $partialSummary.UnmatchedScopeUnknown 1 'Unmatched AD domain remains unassessed'
    if ($null -ne $partialSummary.NoObservedMatch) { throw 'Partial AD scope must not confirm missing AD matches.' }

    $script:SentMail.Clear()
    Publish-TestLicenseReport -Folder $testRoot -TenantRows $rows -CollectedAtUtc ([datetimeoffset]::UtcNow.ToString('o')) | Out-Null; Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath -Manual -ForceLicenseSummaryEmail | Out-Null
    if ($script:SentMail[0].BodyHtml -notlike '*PROVISIONAL (*2/3 domains collected*' -or
        $script:SentMail[0].BodyHtml -notlike '*AD activity scope:*' -or
        $script:SentMail[0].BodyHtml -like '*FORCED AD CSV*') { throw 'Accepted partial AD scope is not visibly distinguished from a forced bypass.' }
    $partialSnapshot = Get-Content -LiteralPath (Join-Path $testRoot 'M365_Licenses_ReportSnapshot.json.txt') -Raw | ConvertFrom-Json
    $partialSnapshotSource = @($partialSnapshot.Sources | Where-Object Name -eq 'AD_Users_AllDomains.csv')[0]
    Assert-Equal $partialSnapshotSource.Provisional $true 'Snapshot preserves partial AD qualification'
    Assert-Equal $partialSnapshotSource.DomainCoverage.Status 'PartialAccepted' 'Snapshot retains structured domain coverage'
    if ($null -ne $partialSnapshot.AdGapActivity.NoObservedMatch) { throw 'Snapshot converted unassessed AD absence to zero.' }
    $partialGapCsv = @(Import-Csv -LiteralPath (Join-Path $testRoot 'M365_Licenses_ReportGaps.csv'))
    $partialNoMatch = @($partialGapCsv | Where-Object {$_.Section -eq 'AdGapActivity' -and $_.Metric -eq 'NoObservedMatch'})[0]
    Assert-Equal $partialNoMatch.Value '' 'CSV keeps unassessed AD absence empty'
    Assert-Equal $partialNoMatch.EvidenceStatus 'N/D' 'Unassessed AD absence is not qualified'
    $partialObserved = @($partialGapCsv | Where-Object {$_.Section -eq 'AdGapActivity' -and $_.Metric -eq 'AdObserved'})[0]
    Assert-Equal $partialObserved.EvidenceStatus 'Provisional' 'Observed counts disclose partial AD coverage'
    $partialSourceCsv = @(Import-Csv -LiteralPath (Join-Path $testRoot 'M365_Licenses_ReportSources.csv') | Where-Object Name -eq 'AD_Users_AllDomains.csv')[0]
    Assert-Equal $partialSourceCsv.Provisional 'True' 'CSV source qualification remains provisional'

    $validProofJson = $qualifiedProof | ConvertTo-Json -Depth 12
    $rejectedCases = @(
      @{Label='Unqualified consumer'; Mutate={param($p) $p.ConsumerScopeQualified=$false}},
      @{Label='String consumer flag'; Mutate={param($p) $p.ConsumerScopeQualified='False'}},
      @{Label='Failed producer'; Mutate={param($p) $p.Status='Failed'}},
      @{Label='Producer error'; Mutate={param($p) $p.Errors=1}},
      @{Label='Another producer'; Mutate={param($p) $p.Producer='Other-Inventory.ps1'}},
      @{Label='Missing domain coverage'; Mutate={param($p) $p.DomainCoverage=$null}},
      @{Label='Untolerated unavailable domain'; Mutate={param($p) $p.DomainCoverage.NonBlockingDomainErrors=@('allowed.example.invalid')}},
      @{Label='Empty collected scope'; Mutate={param($p) $p.DomainCoverage.CollectedDomains=@()}},
      @{Label='Overlapping domains'; Mutate={param($p) $p.DomainCoverage.CollectedDomains+=@('unavailable.example.invalid')}},
      @{Label='Unaccounted expected domain'; Mutate={param($p) $p.DomainCoverage.ExpectedDomains+=@('extra.example.invalid')}},
      @{Label='Repeated coverage domain'; Mutate={param($p) $p.DomainCoverage.CollectedDomains+=@('ALLOWED.example.invalid')}},
      @{Label='Mismatched file coverage'; Mutate={param($p) $p.Files[0].DomainCoverage.CollectedDomains=@('observed.example.invalid')}},
      @{Label='Failed file'; Mutate={param($p) $p.Files[0].Status='Failed'}},
      @{Label='File error'; Mutate={param($p) $p.Files[0].Errors=1}},
      @{Label='Wrong file hash'; Mutate={param($p) $p.Files[0].SHA256='BADHASH'}},
      @{Label='Missing required file'; Mutate={param($p) $p.Files=@()}}
    )
    foreach ($case in $rejectedCases) {
      $badProof = $validProofJson | ConvertFrom-Json
      & $case.Mutate $badProof
      $badProof | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $adManifestPath
      $rejectedAd = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
      Assert-Equal @($rejectedAd.Sources | Where-Object Name -eq 'AD_Users_AllDomains.csv')[0].Ready $false $case.Label
    }
    $qualifiedProof | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $adManifestPath
    $qualifiedAdRows[0].DomainName='unavailable.example.invalid'
    $qualifiedAdRows | Export-Csv -LiteralPath $qualifiedAdPath -NoTypeInformation
    $qualifiedProof.Files[0].SHA256=(Get-FileHash -LiteralPath $qualifiedAdPath).Hash
    $qualifiedProof | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $adManifestPath
    $outsideAd = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
    Assert-Equal @($outsideAd.Sources | Where-Object Name -eq 'AD_Users_AllDomains.csv')[0].Ready $false 'Rows in unavailable domains are not accepted'
    $qualifiedAdRows[0].DomainName='allowed.example.invalid'
    $qualifiedAdRows[0].TenantKey='another-tenant'
    $qualifiedAdRows | Export-Csv -LiteralPath $qualifiedAdPath -NoTypeInformation
    $qualifiedProof.Files[0].SHA256=(Get-FileHash -LiteralPath $qualifiedAdPath).Hash
    $qualifiedProof | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $adManifestPath
    $foreignAd = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
    Assert-Equal @($foreignAd.Sources | Where-Object Name -eq 'AD_Users_AllDomains.csv')[0].Ready $false 'Partial AD scope preserves tenant isolation'
  }
  finally {
    Set-Content -LiteralPath $qualifiedAdPath -Value $originalAdCsv -NoNewline
    Remove-Item -LiteralPath $adManifestPath -Force -ErrorAction SilentlyContinue
  }
  @{Status='Failed';IsPartialInventory=$true} | ConvertTo-Json | Set-Content -LiteralPath $adManifestPath
  $partialAd = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  $f3PartialAd = @($partialAd.Rows | Where-Object Product -eq 'Microsoft 365 F3')[0]
  Assert-Equal $f3PartialAd.Counts.AdEntraInactive 0 'Partial AD export cannot prove inactivity'
  Assert-Equal $f3PartialAd.Counts.AdEntraUnknown 1 'Only non-shared F3 account has unknown AD activity'
  Assert-Equal @($partialAd.Sources | Where-Object Name -eq 'AD_Users_AllDomains.csv')[0].Ready $false 'Failed AD receipt is rejected'
  $forcedAd = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today -ForceAdCsvAnalysis
  $forcedAdSource = @($forcedAd.Sources | Where-Object Name -eq 'AD_Users_AllDomains.csv')[0]
  $forcedF3 = @($forcedAd.Rows | Where-Object Product -eq 'Microsoft 365 F3')[0]
  Assert-Equal $forcedAdSource.Ready $true 'Fresh AD CSV is accepted with explicit override'
  Assert-Equal $forcedAdSource.Forced $true 'AD override is recorded in source metadata'
  Assert-Equal $forcedAd.AdSourceForced $true 'AD override is reported to the email builder'
  Assert-Equal $forcedF3.Counts.AdEntraInactive 1 'Forced AD CSV contributes only non-shared user inactivity'
  $script:SentMail.Clear()
  Publish-TestLicenseReport -Folder $testRoot -TenantRows $rows -CollectedAtUtc ([datetimeoffset]::UtcNow.ToString('o')) -ForceAdCsvAnalysis | Out-Null; Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath -Manual -ForceLicenseSummaryEmail | Out-Null
  if ($script:SentMail[0].BodyHtml -notlike '*Provisional AD/Entra indicator*' -or
      $script:SentMail[0].BodyHtml -notlike '*PROVISIONAL (*FORCED AD CSV*') {
    throw 'Forced AD usage is not visibly qualified in the email.'
  }
  $adPath = Join-Path $testRoot 'AD_Users_AllDomains.csv'
  $adLastWriteUtc = (Get-Item -LiteralPath $adPath).LastWriteTimeUtc
  [System.IO.File]::SetLastWriteTimeUtc($adPath, $today.AddDays(-20))
  $staleAd = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today -ForceAdCsvAnalysis
  Assert-Equal @($staleAd.Sources | Where-Object Name -eq 'AD_Users_AllDomains.csv')[0].Ready $false 'AD override does not bypass CSV freshness'
  Assert-Equal @($staleAd.Rows | Where-Object Product -eq 'Microsoft 365 F3')[0].Counts.AdEntraInactive 0 'Stale AD CSV does not prove inactivity'
  [System.IO.File]::SetLastWriteTimeUtc($adPath, $adLastWriteUtc)
  $adRows = @(Import-Csv -LiteralPath $adPath)
  $adRows[0].TenantKey = 'another-tenant'
  $adRows | Export-Csv -LiteralPath $adPath -NoTypeInformation
  $crossTenantAd = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today -ForceAdCsvAnalysis
  Assert-Equal @($crossTenantAd.Sources | Where-Object Name -eq 'AD_Users_AllDomains.csv')[0].Ready $false 'AD override does not bypass tenant isolation'
  Assert-Equal $crossTenantAd.AdSourceForced $false 'Rejected AD CSV is not presented as forced evidence'
  $adRows[0].TenantKey = 'prod'
  $adRows | Export-Csv -LiteralPath $adPath -NoTypeInformation
  Remove-Item -LiteralPath $adManifestPath -Force

  $staleApps = @(Import-Csv -LiteralPath $appsPath)
  foreach ($row in $staleApps) { $row.'Report Refresh Date' = $today.AddDays(-20).ToString('yyyy-MM-dd') }
  $staleApps | Export-Csv -LiteralPath $appsPath -NoTypeInformation
  Update-TestUserActivityReceipt -Folder $testRoot
  $usage = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  $e5 = @($usage.Rows | Where-Object Product -eq 'Microsoft 365 E5')[0]
  Assert-Equal $e5.Counts.LocalAppsInactive 0 'Stale Apps report is not treated as no use'
  Assert-Equal $e5.Counts.LocalAppsUnknown 0 'Shared E5 mailbox remains outside local Apps usage'

  $conflictGuid = [guid]::Parse('10213243-5465-7687-98a9-bacbdcedfe0f')
  $activeRows = @(Import-Csv -LiteralPath $activePath)
  @($activeRows | Where-Object { $_.'Object Id' -eq 'u11' })[0].OnPremisesImmutableId = [Convert]::ToBase64String($conflictGuid.ToByteArray())
  $activeRows | Export-Csv -LiteralPath $activePath -NoTypeInformation
  @(
    [pscustomobject]@{TenantKey='prod';ObjectGUID=$localGuid.ToString();UserPrincipalName='u12@example.invalid';RecipientType='UserMailbox';NativeIdentityStatus='Observed'}
    [pscustomobject]@{TenantKey='prod';ObjectGUID=$conflictGuid.ToString();UserPrincipalName='u12@example.invalid';RecipientType='RoomMailbox';NativeIdentityStatus='Observed'}
    [pscustomobject]@{TenantKey='prod';ObjectGUID=([guid]::NewGuid().ToString());UserPrincipalName='u8@example.invalid';RecipientType='UserMailbox';NativeIdentityStatus='Observed'}
  ) | Export-Csv -LiteralPath $onPremPath -NoTypeInformation
  $conflictHash = (Get-FileHash -LiteralPath $onPremPath -Algorithm SHA256).Hash
  @{Status='Completed';IsPartialInventory=$false;ConsumerScopeQualified=$true;Files=@(@{File='Exchange_OnPrem_Mailboxes_AllDomains.csv';Status='Success';IsPartialInventory=$false;SHA256=$conflictHash})} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $onPremManifestPath
  $ambiguousOnPrem = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $ambiguousOnPrem.MailboxGap.NoUserMailbox.Available $true 'Unrelated accounts remain qualified when local identities conflict'
  Assert-Equal $ambiguousOnPrem.MailboxGap.NoUserMailbox.Unknown 2 'Identity conflicts qualify affected unlicensed accounts as unknown'
  Assert-Equal $ambiguousOnPrem.MailboxGap.NoUserMailbox.Total 0 'Conflicting local mailbox identities are not counted as mailbox-free'
  $activeRows = @(Import-Csv -LiteralPath $activePath)
  @($activeRows | Where-Object { $_.'Object Id' -eq 'u11' })[0].OnPremisesImmutableId = ''
  $activeRows | Export-Csv -LiteralPath $activePath -NoTypeInformation
  @([pscustomobject]@{TenantKey='prod';ObjectGUID=$localGuid.ToString();UserPrincipalName='u12@example.invalid';RecipientType='UserMailbox';NativeIdentityStatus='Observed'}) |
    Export-Csv -LiteralPath $onPremPath -NoTypeInformation

  @{Status='Failed';IsPartialInventory=$true;ConsumerScopeQualified=$false;Files=@()} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $onPremManifestPath
  $partialOnPrem = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $partialOnPrem.MailboxGap.UserMailboxes.Available $true 'EXO UserMailbox section remains available'
  Assert-Equal $partialOnPrem.MailboxGap.NoUserMailbox.Available $false 'Partial on-premises receipt makes section 07 N/D'
  $script:SentMail.Clear()
  Publish-TestLicenseReport -Folder $testRoot -TenantRows $rows -CollectedAtUtc ([datetimeoffset]::UtcNow.ToString('o')) | Out-Null; Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath $script:MailStatePath -Manual -ForceLicenseSummaryEmail | Out-Null
  if ($script:SentMail[0].BodyHtml -notlike '*Section 07 N/D*') { throw 'Unqualified on-premises source is not visible in the email.' }
  @{Status='Completed';IsPartialInventory=$false;ConsumerScopeQualified=$true;Files=@(@{File='Exchange_OnPrem_Mailboxes_AllDomains.csv';Status='Success';IsPartialInventory=$false;SHA256='BADHASH'})} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $onPremManifestPath
  $changedOnPrem = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $changedOnPrem.MailboxGap.NoUserMailbox.Available $false 'Changed on-premises CSV hash makes section 07 N/D'
  @{Status='Completed';IsPartialInventory=$false;ConsumerScopeQualified=$true;Files=@(@{File='Exchange_OnPrem_Mailboxes_AllDomains.csv';Status='Success';IsPartialInventory=$false;SHA256=$onPremHash})} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $onPremManifestPath

  $licenseHashNow = (Get-FileHash -LiteralPath $licenseUsersPath -Algorithm SHA256).Hash
  @{Status='Completed';IsPartialInventory=$false;Files=@()} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $licenseManifestPath
  $oldReceipt = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today
  Assert-Equal $oldReceipt.Rows[0].Available $false 'A licensing receipt without the file record is not qualified'
  @{Status='Running';IsPartialInventory=$true;Files=@()} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $licenseManifestPath
  $licenseFullPath = [IO.Path]::GetFullPath($licenseUsersPath)
  $runProof = [pscustomobject]@{ RunId='current-run-test'; Hashes=@{ $licenseFullPath=$licenseHashNow } }
  $currentRun = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today -CurrentRunPublication $runProof
  Assert-Equal $currentRun.Rows[0].Available $true 'Current-run publication qualifies the licensing file while its receipt is running'
  Assert-Equal $currentRun.Sources[0].Provenance 'current run publication' 'Current-run provenance is explicit'
  Assert-Equal $currentRun.Sources[0].Forced $false 'Current-run publication is not a receipt bypass'
  $wrongRun = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today `
    -CurrentRunPublication ([pscustomobject]@{ RunId='current-run-test'; Hashes=@{ $licenseFullPath='BADHASH' } })
  Assert-Equal $wrongRun.Rows[0].Available $false 'Current-run proof rejects a changed licensing file'
  Assert-Equal $wrongRun.Sources[0].Reason 'file hash differs from the current-run publication' 'Changed licensing file reason'
  $notPublished = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today `
    -CurrentRunPublication ([pscustomobject]@{ RunId='current-run-test'; Hashes=@{} })
  Assert-Equal $notPublished.Rows[0].Available $false 'Current-run proof requires a publication by this run'
  $onPremCurrentHash = (Get-FileHash -LiteralPath $onPremPath -Algorithm SHA256).Hash
  @{Status='Failed';IsPartialInventory=$true;ConsumerScopeQualified=$false;Files=@()} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $onPremManifestPath
  $foreignSource = Get-LicensesCsvSource -Folder $testRoot -FileName 'Exchange_OnPrem_Mailboxes_AllDomains.csv' -Columns @('TenantKey','ObjectGUID') -AsOfUtc $today `
    -CollectorManifestName 'SmartInventory_SmartM365-Exchange-Local-Mailboxes-Inventory.current.json.txt' -RequireFileReceipt `
    -CurrentRunPublication ([pscustomobject]@{ RunId='current-run-test'; Hashes=@{ ([IO.Path]::GetFullPath($onPremPath))=$onPremCurrentHash } })
  Assert-Equal $foreignSource.Ready $false 'Current-run proof never replaces another producer receipt'
  @{Status='Completed';IsPartialInventory=$false;ConsumerScopeQualified=$true;Files=@(@{File='Exchange_OnPrem_Mailboxes_AllDomains.csv';Status='Success';IsPartialInventory=$false;SHA256=$onPremCurrentHash})} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $onPremManifestPath
  $licenseLastWriteUtc = (Get-Item -LiteralPath $licenseUsersPath).LastWriteTimeUtc
  [System.IO.File]::SetLastWriteTimeUtc($licenseUsersPath, $today.AddDays(-20))
  $staleRun = Get-LicensesFocusedUsageRows -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -AsOfUtc $today -CurrentRunPublication $runProof
  Assert-Equal $staleRun.Rows[0].Available $false 'Current-run proof does not bypass freshness'
  [System.IO.File]::SetLastWriteTimeUtc($licenseUsersPath, $licenseLastWriteUtc)

  # Email-only mode renders only a report qualified by the current licensing receipt.
  @{Status='Completed';IsPartialInventory=$false;Files=@(@{File='M365_Licenses_Users.csv';Status='Success';IsPartialInventory=$false;SHA256=$licenseHashNow})} |
    ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $licenseManifestPath
  $qualifiedReport = Publish-TestLicenseReport -Folder $testRoot -TenantRows $rows -CollectedAtUtc ([datetimeoffset]::UtcNow.ToString('o'))
  $publishedRead = Read-LicensesPublishedReport -Folder $testRoot -TenantKey 'prod'
  Assert-Equal $publishedRead.SnapshotId $qualifiedReport.Id 'Published report is read back by snapshot identity'
  Assert-Equal @($publishedRead.SummaryRows).Count 4 'Published report keeps the four target suites'
  Assert-Equal @($publishedRead.AdditionalRows).Count 3 'Published report keeps the three paid overview products'
  Assert-Equal $publishedRead.Usage.Rows[2].Counts.GetType().Name 'Hashtable' 'Snapshot counts are restored as lookup tables'
  $publishedRead.SharedMailboxMigrationEvidence=[pscustomobject]@{BatchId='20261103T070757019Z-f084240f';BatchCreatedUtc='2026-11-03T07:10:00Z';ExceptionSnapshotDate='2026-11-03';Counts=[pscustomobject]@{Required=3;AssignedCoverage=1;ToComplete=1;CoverageReview=1;Excluded=9;IdentityOrRegisterReview=2;SizeReview=1;Under50=6;E3ToPlan2Review=1;E3Under50Review=2;F3Under50Review=1}}
  $migrationPreview=New-LicensesSummaryMailContent -PublishedReport $publishedRead
  if($migrationPreview.Body -notlike '*Known minimum: 3 required; 1 assigned coverage; 1 to complete; 1 coverage review*' -or $migrationPreview.Body -notlike '*Excluded/retained: 9*' -or $migrationPreview.Body -notlike '*20261103T070757019Z-f084240f*'){throw 'Prepared migration evidence is not accurately rendered.'}
  $publishedRead.SharedMailboxMigrationEvidence=$null
  $receiptText = Get-Content -LiteralPath $licenseManifestPath -Raw
  $snapshotBytes = [IO.File]::ReadAllBytes($reportSnapshotPath)
  $candidatesPath = Join-Path $testRoot 'M365_Licenses_ReportCandidates.csv'
  $candidatesBytes = [IO.File]::ReadAllBytes($candidatesPath)
  $rejections = @(
    @{ Label='Changed report CSV'; Message='*differs from the current licensing receipt*'; Tenant='prod'
       Arrange={ [IO.File]::WriteAllBytes($candidatesPath, [byte[]]($candidatesBytes + [byte[]](13,10))) } }
    @{ Label='Running licensing receipt'; Message='*not a completed receipt*'; Tenant='prod'
       Arrange={ ($receiptText | ConvertFrom-Json -AsHashtable) | ForEach-Object { $_.Status='Running'; $_ } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $licenseManifestPath } }
    @{ Label='Another tenant'; Message='*not a completed receipt*'; Tenant='test'; Arrange={} }
    @{ Label='Report CSV absent from the receipt'; Message='*not recorded by the current licensing receipt*'; Tenant='prod'
       Arrange={ ($receiptText | ConvertFrom-Json -AsHashtable) | ForEach-Object { $_.Files=@($_.Files | Where-Object { $_.File -ne 'M365_Licenses_ReportGaps.csv' }); $_ } | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $licenseManifestPath } }
    @{ Label='Snapshot of another report'; Message='*belongs to another snapshot*'; Tenant='prod'
       Arrange={ [IO.File]::WriteAllText($reportSnapshotPath, ([Text.Encoding]::UTF8.GetString($snapshotBytes).Replace($qualifiedReport.Id, [guid]::NewGuid().ToString('N')))) } }
  )
  foreach ($case in $rejections) {
    & $case.Arrange
    $rejectedGate = Join-Path $mailGateRoot ('rejected-' + [guid]::NewGuid().ToString('N') + '.json.txt')
    $rejectedMessage = ''
    try { Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey $case.Tenant -MailStatePath $rejectedGate -Manual | Out-Null }
    catch { $rejectedMessage = $_.Exception.Message }
    if ($rejectedMessage -notlike $case.Message) { throw "$($case.Label) was not rejected as expected: '$rejectedMessage'." }
    Assert-Equal (Get-Item -LiteralPath $rejectedGate).Length 0 "$($case.Label) leaves no daily send marker"
    Set-Content -LiteralPath $licenseManifestPath -Value $receiptText -NoNewline
    [IO.File]::WriteAllBytes($reportSnapshotPath, $snapshotBytes)
    [IO.File]::WriteAllBytes($candidatesPath, $candidatesBytes)
  }
  $script:SentMail.Clear()
  Send-LicensesFocusedSummaryEmail -CsvFolderPath $testRoot -ExpectedTenantKey 'prod' -MailStatePath (Join-Path $mailGateRoot 'qualified-report.json.txt') -Manual | Out-Null
  Assert-Equal $script:SentMail.Count 1 'Qualified published report is sent in email-only mode'
  if ($script:SentMail[0].BodyHtml -notmatch [regex]::Escape("Snapshot: $($qualifiedReport.Id)")) { throw 'Email-only mode does not identify the published snapshot.' }
}
finally {
  if (Test-Path -LiteralPath $testRoot) { Remove-Item -LiteralPath $testRoot -Recurse -Force }
  if (Test-Path -LiteralPath $mailGateRoot) { Remove-Item -LiteralPath $mailGateRoot -Recurse -Force }
}
'PASS: Focused license summary email offline checks.'

# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBBqtbXdwCq8cPy
# Csr7dm1TJEHhtwxSy6aGhNQyScK9oaCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgSyEwbrRs64cj
# Xx39jVnW6z/GMSuXgOIfA800DSTo8i4wDQYJKoZIhvcNAQEBBQAEggGAofA8s+HR
# 0u6jT8dXTt1UDnJrtoKTqeEKnrwYD1ux+j38clPTZ6WaPK/nJSOLOH5rIHD95s0j
# 1BUkhsl6AKLzEflDto6VI3DoMDWqIfTDwAvQ/1ML3VXsWuEvE/CSX9HgNwdM64Tm
# Gt3W/lu637xCX8WoOPs8cIeZ4Xr/9W9CrtSTUr3jTxVK6xiKtgNMsREIFd5FMbIn
# fkj9KCKp3SEBODm6s2ab0I7s73mnjsguOU9TZ99LqJd8P5Wre0nmw1W5q0+Pj/wp
# NMLChKcY8H5Ti3SlU0ftDw0zuIN/QGTXbmA6pZ2DDFyOe9F59JRJk0BDBHzy1PMg
# 8E1xCtpuxKeAh+EdUVHrPTCKVKN3XJ0YWP3Vm8lrThsZusoFxPEEZCk9hJE6FT+o
# uuDswLZ5K3jmvh1xK5NS3e6jKZ/yspD0pWe9GxBshVifzFYZrKznlf1vyyLyThcK
# 0wJh8HD/cpT5eiK/zB/SlGdMtm++iIkqtAj1iK02keZ+7cQIyBg5ZCaooYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTA0MTYxNDE0WjA/BgkqhkiG9w0BCQQxMgQwd2BvfBzSdxJOjgD7KrkWKN0f
# 1VkxZ5u3QsVV8NvXmc/h8RVd4WMju3FZVuM19wGxMA0GCSqGSIb3DQEBAQUABIIC
# AF8zO25LfTFwmFXp1d4uM+3p6TeIYon4JChZDzLFTUUWdAYqneI2NRW54d/NjJWn
# RjQ3ZthlA0vForS9+VXFjZu3BRcAVGVOTYm1EdP/tYQuBJBZwDKMdwdnHFAG9Nkz
# kNDXK0fkSsjtJLpHn7/WCjiRguY00r9E/GnHdbzqzSuwy/hH5bP/umbyJe7hozZi
# SljHgP2sncMMUJ2lVBC6mS4n2vlrszOJBm17Fie4+ooUVTELStPmTFD/IrNjkeES
# /Lmz19ivGZMnrHmOPOb2SKU1pXl7VDbeSvhwjPqrgMYS2I2qkj8PALwQa8HOySXZ
# j1N6zjQsBDjlfruATww5L8G8DJDEQCe64Xyw6nfihmybSALV3XR/QtNh8XL2hFe6
# kn2YvdCGTG+pNPbfLomXjTNJzwRfVeADYswak7xlvB4yuitXGqHm0NSsQOCBV//7
# hDCFpSGkRQEJ5bJBxfzUH4ObiczjQTtLWEqi+DHpbEqEYEnbwuKHob1x1IVM/eLN
# MqgLmnywFXunHyaIroWNBAN/OvJocr5b+ET2MmvdrfNO0SRvnv0YCiFTkp0GD+Rl
# 7E8oABknWFIzczPbSARuTrWNLw1TdmDHVnAojj0+cAC7wQ+c/DFhLVtFBoJg3WAi
# nVYWe4wo21dGtYbjwa861PNyerjaM1pVdrpuuZBZ2hK3
# SIG # End signature block
