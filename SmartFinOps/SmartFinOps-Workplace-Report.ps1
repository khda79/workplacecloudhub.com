#Requires -Version 7.0
<#
.SYNOPSIS
  SmartFinOps Workplace V2 executive HTML report (read-only).
.DESCRIPTION
  Pure consumer of the valuation engine: it reads the seven FinOps_* CSVs and the
  SmartFinOps_Workplace-Valuation receipt, qualifies them, then displays, sorts and groups them for reading.
  No decision and no amount is computed: every amount comes from FinOps_Summary, FinOps_SkuValuation,
  FinOps_Showback or FinOps_Realization. A failed or inconsistent run gives a "Report not available" page
  with the reason and no figure; it replaces the current report so that no older figure stays visible.
  Layout (2.5.0): a summary on top (four published totals with their certainty, top 3 decisions, source states), then
  the detail in folding sections; SmartM365 GUI charter, WorkplaceCloudHub logo embedded, no client logo.
.VERSION
2.5.0
.NOTES
  Input and output: the SmartM365 DATA-LAST of the profile (SmartFinOps_Workplace-Report.html), a copy in
  DATA-ALL\FinOps\Workplace\<FinOpsRunId>\ when that run is in the history, and the log in
  LOG-ALL\SmartFinOps-Workplace-Report\; mirrored to SharePoint when upload is enabled (SmartM365 setting by default).
  Daily email (Report.EnableDailyEmail, default true): once per Europe/Paris day and profile (-ForceEmail sends again),
  aggregates only, the page attached, through the SmartM365 mail transport of the profile; never in -ValidateOnly.
  Aggregates only: no user identifier, no external resource, no script, every value HTML-encoded. Same inputs give the
  same HTML except the generation time.
#>
[CmdletBinding()]
param(
  [string]$Tenant = 'test',
  [string]$SmartM365LatestCsvFolderPath,
  [string]$SmartM365DataAllRootPath,
  [string]$SmartM365LogAllRootPath,
  [string]$SmartM365TenantKey,
  [switch]$ForceEmail,
  [switch]$ValidateOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:ScriptVersion = '2.5.0'
# Run with -File, this script scope is the global scope: SmartM365 session values (LogPath, SmartM365TenantKey, ...)
# set by the mail and upload helpers would overwrite script variables of the same name. Names here never collide.
$script:TenantKeyParameter = [string]$SmartM365TenantKey
$script:LogLines = [System.Collections.Generic.List[string]]::new()
$script:Invariant = [Globalization.CultureInfo]::InvariantCulture
$script:ReceiptName = 'SmartFinOps_Workplace-Valuation.current.json.txt'
$script:ReportName = 'SmartFinOps_Workplace-Report.html'
$script:DailyEmailStateName = 'SmartFinOps_Workplace-Report_DailyEmail_SendState.json.txt'

#region Input contract (engine 2.1.4 to 2.2.x)
$script:CommonColumns = @('TenantKey','FinOpsRunId','SnapshotId','ReceiptRunId')
$script:InputColumns = [ordered]@{
  'FinOps_Sources.csv' = @('SourceFile','Producer','ReceiptFile','SourceReceiptRunId','ReceiptStatus','ReceiptCompletedAtUtc','Required','PublishedInReceipt','ExpectedSHA256','ActualSHA256','ExpectedRows','ActualRows','FreshnessThresholdHours','AgeHours','QualificationStatus','Reason','MissingDomains','ProvisionalLevers')
  'FinOps_Opportunities.csv' = @('OpportunityKey','Lever','ValueClass','ValuationStatus','StatusReason','SourceFile','SourceCandidateType','SourceRecommendation','SourceReason','MailboxGuard','DependencyStatus','CountedIn','ReportProduct','SkuId','SkuPartNumber','TargetSkuPartNumber','UserId','UserPrincipalName','Units','UnitPriceMonthly','TargetUnitPriceMonthly','MonthlyAmount','AnnualAmount','Currency','PriceRowKey','PricingReceiptRunId','PriceSource','PriceIsEstimated','CapRank','ContractKey','RenewalDate','RealizableFrom','AdDomain','Country','Entity','Department','ShowbackJoinStatus')
  'FinOps_SkuValuation.csv' = @('SkuId','SkuPartNumber','SkuDisplayName','ReportProduct','Excluded','EnabledUnits','ConsumedUnits','AvailableUnits','ReserveUnits','MinimumUnits','CandidateUnitsR','PoolUnits','ReusedUnits','ReducibleUnits','CapUnits','RealizableUnits','CapacityRealizableUnits','CappedUnits','ReviewUnits','NotMeasuredUnits','BlockedUnits','ReferencedUnits','UnitPriceMonthly','PriceRowKey','PriceIsEstimated','RealizableAnnualAmount','AvoidedAnnualAmount','ValuationStatus','ContractKey','ContractStatus','ReductionPolicy','RenewalDate','RealizableFrom','UsageEvidence')
  'FinOps_Showback.csv' = @('Dimension','DimensionValue','Lever','ValueClass','ValuationStatus','Users','Units','AnnualAmount','Currency','Provisional')
  'FinOps_Summary.csv' = @('KpiKey','ValueClass','Lever','ValuationStatus','Users','Units','AnnualAmount','Currency','Provisional','MissingDomains','SnapshotGeneratedAtUtc','PricingReceiptRunId','Note')
  'FinOps_Realization.csv' = @('SkuId','SkuPartNumber','BaselineFinOpsRunId','BaselineSnapshotId','EnabledUnitsBefore','EnabledUnitsAfter','Decrease','RemovedCandidateLicenses','BaselineCapacityUnits','AttributableUnits','RealizedUnits','NotAttributedUnits','UnitPriceMonthly','PriceRowKey','RealizedAnnualAmount','Status')
  'FinOps_Controls.csv' = @('Control','Key','SourceCount','FinOpsCount','Status','Detail')
}
# The opportunity columns the HTML may use; identifiers (UserId, UserPrincipalName, OpportunityKey, SkuId, CountedIn) are never rendered.
$script:Totals = [ordered]@{
  TOTAL_R = 'Realizable reduction (R)'
  TOTAL_R_CONFIRMED = 'R confirmed'
  TOTAL_R_ENABLED_INACTIVE = 'R to confirm: enabled accounts without M365 activity in 90 days'
  TOTAL_R_DISABLED_AND_SHARED = 'of which disabled accounts and eligible shared mailboxes'
  TOTAL_R_COPILOT = 'of which Copilot without activity'
  TOTAL_R_CAPACITY = 'of which unassigned units'
  TOTAL_RC_E3_TO_F3 = 'Conditional E3 to F3 (not in the R total)'
  TOTAL_E_REUSE = 'Cost avoidance: reuse for users waiting for capacity'
  TOTAL_E_ESU = 'Cost avoidance: Windows 10 ESU'
}
$script:Levers = [ordered]@{
  L0 = @('Report candidate (counted in the suite levers)', '')
  L1 = @('Disabled accounts', 'Remove the suite license from disabled accounts')
  L2 = @('Eligible shared mailboxes under 50 GB', 'Convert to shared mailboxes and remove the license')
  L3 = @('Enabled accounts without M365 activity in 90 days', 'Confirm with the business, then remove or reassign')
  L4 = @('E3 to F3 review', 'Confirm Frontline eligibility, then downgrade at renewal')
  L5 = @('Disabled accounts (per SKU)', 'Remove the SKU from disabled accounts')
  L6 = @('Copilot without activity', 'Reassign or remove unused Copilot licenses')
  L7 = @('Unassigned units', 'Reduce unassigned units at contract or renewal')
  L8 = @('Reuse for users waiting for capacity', 'Assign freed units before buying new ones')
  L9 = @('Over-assigned SKUs', '')
  L10 = @('Dynamics 365 usage not measured', '')
  L11 = @('License assignment issues', '')
  L12 = @('Mailbox and AD gap metrics', '')
  L13 = @('Windows 10 ESU', '')
}
$script:ReservedValues = @('Not attributed', 'Not measured')
#endregion

#region Helpers
function Write-ReportLog {
  param([Parameter(Mandatory)][AllowEmptyString()][string]$Message, [ValidateSet('INFO','WARNING','ERROR')][string]$Level = 'INFO')
  foreach ($part in ($Message -split "`r?`n")) {
    $line = '{0} [{1}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Level, $part
    $script:LogLines.Add($line)
    $color = switch ($Level) { 'WARNING' { 'Yellow' } 'ERROR' { 'Red' } default { 'Gray' } }
    Write-Host $line -ForegroundColor $color
  }
}

function Get-ReportProperty {
  param([AllowNull()]$InputObject, [Parameter(Mandatory)][string]$Name)
  if ($null -eq $InputObject) { return $null }
  if ($InputObject -is [System.Collections.IDictionary]) { if ($InputObject.Contains($Name)) { return $InputObject[$Name] } return $null }
  $property = $InputObject.PSObject.Properties[$Name]
  if ($property) { return $property.Value }
  return $null
}

function ConvertTo-ReportDecimal {
  param([AllowNull()]$Value)
  $text = ([string]$Value).Trim()
  # The engine writes invariant decimals ('1234.50'); anything else is not a measured amount.
  if ($text -notmatch '^-?\d+(?:\.\d+)?$') { return $null }
  return [decimal]::Parse($text, [Globalization.NumberStyles]::AllowLeadingSign -bor [Globalization.NumberStyles]::AllowDecimalPoint, $script:Invariant)
}

function ConvertTo-ReportLong {
  param([AllowNull()]$Value)
  $text = ([string]$Value).Trim()
  if ($text -notmatch '^-?\d+$') { return $null }
  return [long]::Parse($text, $script:Invariant)
}

function ConvertTo-ReportJsonText {
  param([AllowNull()]$Value, [string]$Format = 'o')
  # ConvertFrom-Json turns ISO 8601 strings into DateTime: format them back, independently of the host culture.
  if ($Value -is [datetime]) { return $Value.ToUniversalTime().ToString($Format, $script:Invariant) }
  return [string]$Value
}

function ConvertTo-ReportHtml {
  param([AllowNull()]$Value)
  return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function ConvertTo-ReportSafeText {
  param([AllowNull()]$Value)
  # Free text from a source may quote an identifier: never render an e-mail address or a GUID.
  $text = [string]$Value
  $text = [regex]::Replace($text, '[^\s@<>"]+@[^\s@<>"]+', '[user]')
  return [regex]::Replace($text, '(?i)\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b', '[id]')
}

function Format-ReportCount {
  param([AllowNull()]$Value)
  $number = ConvertTo-ReportLong $Value
  if ($null -eq $number) { return '' }
  return $number.ToString('#,##0', $script:Invariant)
}

function Format-ReportEur {
  param([AllowNull()]$Value)
  $number = ConvertTo-ReportDecimal $Value
  if ($null -eq $number) { return '' }
  $rounded = [math]::Round($number, 0, [MidpointRounding]::AwayFromZero)
  return ('<span class="amt" title="EUR {0}">EUR {1}</span>' -f $number.ToString('0.00', $script:Invariant), $rounded.ToString('#,##0', $script:Invariant))
}

function Get-ReportBadge {
  param([AllowNull()][string]$Status)
  $map = @{
    'Valued'=@('ok','Valued'); 'Provisional'=@('prov','Provisional'); 'Not valued'=@('nv','Not valued'); 'Not measured'=@('nm','Not measured')
    'Review'=@('rev','Review'); 'Blocked'=@('blk','Blocked'); 'Capped'=@('cap','Capped'); 'Reused'=@('reu','Reused (see cost avoidance)')
    'Referenced'=@('ref','Counted elsewhere'); 'Signal'=@('sig','Signal'); 'Realized'=@('ok','Realized'); 'No reduction observed'=@('sig','No reduction observed')
  }
  $entry = $map[[string]$Status]
  if ($null -eq $entry) { return ('<span class="badge sig">{0}</span>' -f (ConvertTo-ReportHtml (ConvertTo-ReportSafeText $Status))) }
  return ('<span class="badge {0}">{1}</span>' -f $entry[0], $entry[1])
}

function Get-ReportAmountCell {
  param([AllowNull()]$Amount, [AllowNull()][string]$Status)
  $text = Format-ReportEur $Amount
  if ($text) { if ($Status -ceq 'Provisional') { return "$text $(Get-ReportBadge 'Provisional')" } return $text }
  # An empty amount is never shown as zero: it is a status, or no opportunity at all when measured.
  if (-not $Status -or $Status -ceq 'Valued') { return '<span class="none">None identified</span>' }
  return (Get-ReportBadge $Status)
}

function Get-ReportPriceOrigin {
  param([AllowNull()]$PriceIsEstimated)
  switch (([string]$PriceIsEstimated).Trim()) { 'True' { return 'Catalogue estimate' } 'False' { return 'Contract price' } default { return 'No price' } }
}

function Get-ReportSorted {
  param([AllowEmptyCollection()][object[]]$Items, [Parameter(Mandatory)][scriptblock]$Key, [switch]$Descending)
  # Deterministic: explicit keys compared ordinally, never the culture of the host.
  # -Descending applies to the first key, which must then be a number; ties keep the other keys ascending.
  $list = [System.Collections.Generic.List[object]]::new()
  foreach ($item in $Items) {
    if ($null -eq $item) { continue }
    $keys = @(& $Key $item)
    if ($Descending) { $keys[0] = -1 * [decimal]$keys[0] }
    $list.Add([pscustomobject]@{ Keys=$keys; Item=$item })
  }
  $list.Sort([Comparison[object]]{
    param($a, $b)
    for ($i = 0; $i -lt $a.Keys.Count; $i++) {
      $x = $a.Keys[$i]; $y = $b.Keys[$i]
      $result = if ($x -is [string] -or $y -is [string]) { [StringComparer]::OrdinalIgnoreCase.Compare([string]$x, [string]$y) } else { ([decimal]$x).CompareTo([decimal]$y) }
      if ($result -ne 0) { return $result }
    }
    return 0
  })
  foreach ($entry in $list) { $entry.Item }
}
#endregion

#region Qualification
function Read-ReportInputs {
  param([Parameter(Mandatory)][string]$Root, [Parameter(Mandatory)][string]$ProfileKey, [Parameter(Mandatory)][string]$TenantKey)
  $result = [pscustomobject]@{ Available=$false; Reason=''; Receipt=$null; RunId=''; CompletedAtUtc=''; Data=@{}; Checks=[System.Collections.Generic.List[object]]::new() }
  $rule = ''
  $pass = { param([string]$Rule, [string]$Detail) $result.Checks.Add([pscustomobject]@{ Rule=$Rule; Status='Pass'; Detail=$Detail }) }
  try {
    $rule = 'Q1 Receipt'
    $receiptPath = Join-Path $Root $script:ReceiptName
    if (-not (Test-Path -LiteralPath $receiptPath -PathType Leaf)) { throw 'The valuation receipt is missing.' }
    if ((Get-Item -LiteralPath $receiptPath).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw 'The valuation receipt is a link.' }
    try { $receipt = [Text.Encoding]::UTF8.GetString([IO.File]::ReadAllBytes($receiptPath)).TrimStart([char]0xFEFF) | ConvertFrom-Json }
    catch { throw 'The valuation receipt is not valid JSON.' }
    $result.Receipt = $receipt
    $result.RunId = [string](Get-ReportProperty $receipt 'RunId')
    $result.CompletedAtUtc = ConvertTo-ReportJsonText (Get-ReportProperty $receipt 'CompletedAtUtc')
    if ([string](Get-ReportProperty $receipt 'Owner') -cne 'SmartFinOps-Receipt' -or [string](Get-ReportProperty $receipt 'ContractVersion') -cne '1.1') { throw 'The valuation receipt has an unsupported owner or contract version (1.1 required: run SmartFinOps-Workplace-Valuation.ps1 2.1.4 or later).' }
    $version = [string](Get-ReportProperty $receipt 'ScriptVersion')
    # Same input contract from engine 2.1.4 to 2.2.x (2.2.0 moved the files to the SmartM365 data tree, not their content).
    $supported = ($version -match '^2\.1\.(\d+)$' -and [int]$Matches[1] -ge 4) -or $version -match '^2\.2\.\d+$'
    if (-not $supported) { throw "The valuation engine version '$version' is not supported (2.1.4 or later, up to 2.2.x)." }
    if ($result.RunId -notmatch '^\d{8}T\d{9}Z-[0-9a-f]{8}$') { throw 'The valuation receipt RunId is invalid.' }
    & $pass $rule "Owner SmartFinOps-Receipt 1.1, engine $version"

    $rule = 'Q2 Status'
    $status = [string](Get-ReportProperty $receipt 'Status')
    if ($status -cne 'Completed') {
      $runError = [string](Get-ReportProperty $receipt 'Error')
      throw ("The valuation run is {0}{1}." -f $(if ($status) { $status } else { 'in an unknown state' }), $(if ($runError) { ": $runError" } else { '' }))
    }
    & $pass $rule 'Completed'

    $rule = 'Q3 Tenant'
    # The organization TenantKey is checked, not only the profile key used for the data folder.
    if ([string](Get-ReportProperty $receipt 'TenantKey') -cne $TenantKey) { throw 'The valuation receipt belongs to another tenant (TenantKey differs).' }
    if ([string](Get-ReportProperty $receipt 'ProfileKey') -cne $ProfileKey) { throw 'The valuation receipt belongs to another profile (ProfileKey differs).' }
    & $pass $rule "TenantKey $TenantKey, profile $ProfileKey"

    $rule = 'Q4 Files'
    $files = @(Get-ReportProperty $receipt 'Files' | Where-Object { $null -ne $_ })
    $names = @($files | ForEach-Object { [string](Get-ReportProperty $_ 'File') } | Sort-Object)
    $expected = @($script:InputColumns.Keys | Sort-Object)
    if (($names -join '|') -cne ($expected -join '|')) { throw 'The valuation receipt does not list exactly the seven FinOps files.' }
    $texts = @{}
    foreach ($file in $files) {
      $name = [string](Get-ReportProperty $file 'File')
      $path = Join-Path $Root $name
      if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "File missing: $name" }
      if ((Get-Item -LiteralPath $path).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "File is a link: $name" }
      # Hash and parse the same bytes: the file cannot change between the check and the read.
      $bytes = [IO.File]::ReadAllBytes($path)
      $hash = [Convert]::ToHexString([Security.Cryptography.SHA256]::HashData($bytes))
      if ($hash -ine [string](Get-ReportProperty $file 'SHA256')) { throw "File hash mismatch: $name" }
      $text = [Text.Encoding]::UTF8.GetString($bytes).TrimStart([char]0xFEFF)
      $rows = @(if ($text.Trim()) { $text | ConvertFrom-Csv })
      if ([string]$rows.Count -cne [string](Get-ReportProperty $file 'Rows')) { throw "Row count mismatch: $name" }
      $texts[$name] = $text
      $result.Data[$name] = $rows
    }
    & $pass $rule '7 files, SHA256 and row counts match the receipt'

    $rule = 'Q5 Schema'
    foreach ($name in $script:InputColumns.Keys) {
      $firstLine = ($texts[$name] -split "`r?`n", 2)[0]
      $header = @($firstLine.Split(',') | ForEach-Object { $_.Trim('"') })
      if (($header -join ',') -cne (@($script:CommonColumns + $script:InputColumns[$name]) -join ',')) { throw "Schema mismatch: $name" }
    }
    & $pass $rule 'Exact headers'

    $rule = 'Q6 Run'
    $snapshot = [string](Get-ReportProperty $receipt 'SnapshotId'); $licencesRun = [string](Get-ReportProperty $receipt 'LicencesReceiptRunId')
    foreach ($name in $script:InputColumns.Keys) {
      foreach ($row in $result.Data[$name]) {
        if ($row.TenantKey -cne $TenantKey -or $row.FinOpsRunId -cne $result.RunId -or $row.SnapshotId -cne $snapshot -or $row.ReceiptRunId -cne $licencesRun) { throw "Mixed run in $name" }
      }
    }
    & $pass $rule "Every row belongs to run $($result.RunId)"

    $rule = 'Q7 Controls'
    $controls = Get-ReportProperty $receipt 'Controls'
    $measured = [bool](Get-ReportProperty $receipt 'LicencesMeasured')
    $controlStatus = [string](Get-ReportProperty $controls 'Status')
    $okStatus = if ($measured) { $controlStatus -ceq 'Pass' } else { $controlStatus -ceq 'Not applicable' }
    if (-not $okStatus -or [string](Get-ReportProperty $controls 'FailedCount') -cne '0' -or @($result.Data['FinOps_Controls.csv'] | Where-Object Status -ceq 'Fail').Count) { throw 'The equivalence controls failed or are inconsistent with the receipt.' }
    & $pass $rule $(if ($measured) { 'Equivalence controls Pass' } else { 'Not applicable (Licences report not measured)' })
    $result.Available = $true
  }
  catch {
    $result.Available = $false
    $result.Reason = $_.Exception.Message
    $result.Checks.Add([pscustomobject]@{ Rule=$rule; Status='Fail'; Detail=$result.Reason })
    $result.Data = @{}
  }
  return $result
}
#endregion

#region HTML building blocks
# SmartM365 GUI charter: background #f5f8fb, white cards bordered #dde7f0 with 8px corners, accent #0078d4,
# accent soft #e6f4ff, text #1f2937, muted #5f6b7a. Sections are <details> (no script): only Financial potential is open.
$script:Css = @'
    body { margin: 0; font-family: Segoe UI, Arial, sans-serif; background: #f5f8fb; color: #1f2937; font-size: 14px; }
    .page { max-width: 1200px; margin: 0 auto; padding: 24px 24px 32px; }
    .card { background: #ffffff; border: 1px solid #dde7f0; border-radius: 8px; }
    .head { display: flex; justify-content: space-between; align-items: center; gap: 20px; padding: 20px 24px; }
    .logo { flex: none; border: 1px solid #dde7f0; border-radius: 8px; padding: 8px 12px; background: #ffffff; }
    .logo img { display: block; width: 120px; height: auto; }
    h1 { margin: 10px 0 6px; font-size: 26px; font-weight: 600; }
    h2 { margin: 0; font-size: 17px; font-weight: 600; }
    h3 { margin: 18px 0 8px; font-size: 14px; }
    .subtitle { color: #5f6b7a; margin: 2px 0 0; font-size: 13px; }
    .statusbar { display: flex; flex-wrap: wrap; align-items: center; gap: 8px 16px; padding: 12px 24px; margin-top: 12px; font-size: 13px; }
    .grid { display: grid; grid-template-columns: repeat(4, minmax(170px, 1fr)); gap: 12px; margin-top: 12px; }
    .metric { background: #ffffff; border: 1px solid #dde7f0; border-left: 4px solid #0078d4; border-radius: 8px; padding: 14px; }
    .metric.to-confirm { border-left-color: #d97706; }
    .metric.conditional { border-left-color: #0f766e; }
    .metric.avoidance { border-left-color: #64748b; }
    .metric .label { color: #5f6b7a; font-size: 12px; }
    .metric .value { font-size: 24px; font-weight: 600; margin: 6px 0; }
    .metric .detail { color: #5f6b7a; font-size: 12px; margin-top: 6px; }
    section { background: #ffffff; border: 1px solid #dde7f0; border-radius: 8px; padding: 16px 20px; margin-top: 16px; overflow-x: auto; }
    section.fold { padding: 0; }
    section.fold > details > summary { cursor: pointer; padding: 14px 20px; list-style: none; }
    section.fold > details > summary::-webkit-details-marker { display: none; }
    section.fold > details > summary h2 { display: inline; }
    section.fold > details > summary::before { content: '+'; display: inline-block; width: 18px; color: #0078d4; font-weight: 700; }
    section.fold > details[open] > summary::before { content: '-'; }
    section.fold > details[open] > summary { border-bottom: 1px solid #dde7f0; }
    .hint { color: #5f6b7a; font-size: 12px; margin-left: 8px; }
    .fold-body { padding: 4px 20px 16px; overflow-x: auto; }
    table { border-collapse: collapse; width: 100%; font-size: 12px; }
    th { text-align: left; background: #e6f4ff; color: #1f2937; border-bottom: 1px solid #dde7f0; padding: 8px; }
    td { border-bottom: 1px solid #edf2f7; padding: 7px 8px; vertical-align: top; }
    td.num, th.num { text-align: right; white-space: nowrap; }
    .note, .empty, .none { color: #5f6b7a; font-size: 12px; }
    .pill { display: inline-block; border-radius: 999px; background: #e6f4ff; color: #005a9e; padding: 3px 9px; font-size: 12px; font-weight: 600; }
    .chips { display: flex; flex-wrap: wrap; gap: 8px; }
    .chip { border: 1px solid #dde7f0; border-radius: 999px; padding: 3px 10px; font-size: 12px; background: #ffffff; }
    .banner { border-radius: 8px; padding: 10px 14px; margin-top: 12px; font-size: 13px; border: 1px solid #f0c36d; background: #fff8e6; }
    .banner.stop { border-color: #e3a4a4; background: #fdeeee; }
    .badge { display: inline-block; border-radius: 4px; padding: 1px 6px; font-size: 11px; font-weight: 600; border: 1px solid; white-space: nowrap; }
    .badge.ok { color: #1d6b3a; border-color: #9fd2b1; background: #eaf7ef; }
    .badge.prov, .badge.tc { color: #8a5a00; border-color: #f0c36d; background: #fff8e6; }
    .badge.nv { color: #4b5563; border-color: #c7ccd3; background: #f3f4f6; }
    .badge.nm { color: #4b5563; border-color: #9ca3af; background: repeating-linear-gradient(45deg, #f3f4f6, #f3f4f6 4px, #e5e7eb 4px, #e5e7eb 8px); }
    .badge.rev { color: #1e4f91; border-color: #9cbbe6; background: #edf4fd; }
    .badge.blk { color: #9b1c1c; border-color: #e3a4a4; background: #fdeeee; }
    .badge.cap { color: #5b2a86; border-color: #c3a6de; background: #f5eefb; }
    .badge.cond { color: #0f5e66; border-color: #9fd0d4; background: #ebf7f8; }
    .badge.reu, .badge.ref, .badge.sig, .badge.av { color: #374151; border-color: #d1d5db; background: #ffffff; }
    .bars text { font-size: 11px; fill: #1f2937; }
    .bars rect { fill: #0078d4; }
    details.sub > summary { cursor: pointer; font-weight: 600; margin: 12px 0 6px; }
    footer { color: #5f6b7a; font-size: 12px; text-align: center; padding: 16px 0 0; }
    @media (max-width: 860px) { .grid { grid-template-columns: repeat(2, minmax(0, 1fr)); } .head { flex-direction: column-reverse; align-items: flex-start; } }
    @media (max-width: 520px) { .grid { grid-template-columns: 1fr; } .page { padding: 12px; } }
    @media print { body { background: #ffffff; } .page { max-width: none; padding: 0; } section { break-inside: avoid; } }
'@

# WorkplaceCloudHub lockup reduced for the report: 240 x 193 px, 256-color PNG (15,958 bytes), derived from the
# WorkplaceCloudHub-lockup-WPF.png GUI asset. Embedded so that the page stays self-contained; no client logo.
$script:LogoPngBase64 = @'
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
'@

function Get-ReportLogoHtml {
  $data = $script:LogoPngBase64 -replace '\s', ''
  return ('<img alt="WorkplaceCloudHub" width="120" height="97" src="data:image/png;base64,{0}">' -f $data)
}

function Add-ReportTable {
  param([Parameter(Mandatory)][System.Text.StringBuilder]$Builder, [Parameter(Mandatory)][string[]]$Headers, [AllowEmptyCollection()][object[]]$Rows, [int[]]$Numeric = @(), [string]$Empty = 'Nothing to report.')
  if (-not $Rows -or $Rows.Count -eq 0) { [void]$Builder.AppendLine(('<p class="empty">{0}</p>' -f (ConvertTo-ReportHtml $Empty))); return }
  # A single row built with ,@(...) arrives unrolled as its cells.
  if ($Rows[0] -isnot [array]) { $Rows = ,$Rows }
  [void]$Builder.Append('<table><thead><tr>')
  for ($i = 0; $i -lt $Headers.Count; $i++) { [void]$Builder.Append(('<th{0}>{1}</th>' -f $(if ($i -in $Numeric) { ' class="num"' } else { '' }), (ConvertTo-ReportHtml $Headers[$i]))) }
  [void]$Builder.AppendLine('</tr></thead><tbody>')
  foreach ($row in $Rows) {
    [void]$Builder.Append('<tr>')
    for ($i = 0; $i -lt $Headers.Count; $i++) { [void]$Builder.Append(('<td{0}>{1}</td>' -f $(if ($i -in $Numeric) { ' class="num"' } else { '' }), $row[$i])) }
    [void]$Builder.AppendLine('</tr>')
  }
  [void]$Builder.AppendLine('</tbody></table>')
}

function Add-ReportBars {
  param([Parameter(Mandatory)][System.Text.StringBuilder]$Builder, [Parameter(Mandatory)][string]$Title, [AllowEmptyCollection()][object[]]$Items)
  # Inline SVG, proportional bars only; the table below it carries the values.
  $bars = @($Items | Where-Object { $null -ne $_.Value -and $_.Value -gt 0 })
  if ($bars.Count -eq 0) { return }
  $max = [decimal]0; foreach ($bar in $bars) { if ($bar.Value -gt $max) { $max = $bar.Value } }
  $height = 26 * $bars.Count + 6
  [void]$Builder.AppendLine(('<svg class="bars" role="img" aria-label="{0}" viewBox="0 0 760 {1}" width="100%" preserveAspectRatio="xMinYMin meet">' -f (ConvertTo-ReportHtml $Title), $height))
  $y = 4
  foreach ($bar in $bars) {
    $width = [math]::Max(1, [math]::Round([double]($bar.Value / $max) * 380, 1))
    $label = [string]$bar.Label; if ($label.Length -gt 42) { $label = $label.Substring(0, 41) + '...' }
    [void]$Builder.AppendLine(('<text x="0" y="{0}">{1}</text><rect x="250" y="{2}" width="{3}" height="16" rx="2"></rect><text x="{4}" y="{0}">{5}</text>' -f ($y + 12), (ConvertTo-ReportHtml $label), $y, $width.ToString($script:Invariant), (256 + $width).ToString($script:Invariant), (ConvertTo-ReportHtml $bar.Text)))
    $y += 26
  }
  [void]$Builder.AppendLine('</svg>')
}

function Get-ReportEurText {
  param([AllowNull()]$Value)
  $number = ConvertTo-ReportDecimal $Value
  if ($null -eq $number) { return '' }
  return 'EUR ' + ([math]::Round($number, 0, [MidpointRounding]::AwayFromZero)).ToString('#,##0', $script:Invariant)
}

function Get-ReportLeverLabel {
  param([string]$Lever)
  # Business label first, lever code in brackets.
  $entry = $script:Levers[$Lever]
  if ($null -eq $entry) { return (ConvertTo-ReportHtml $Lever) }
  return ('{0} ({1})' -f (ConvertTo-ReportHtml $entry[0]), (ConvertTo-ReportHtml $Lever))
}

function Get-ReportLeverNumber { param([string]$Lever) $n = 0; if ([int]::TryParse(([string]$Lever).TrimStart('L'), [ref]$n)) { return $n } return 99 }

function Get-ReportCertaintyBadge {
  param([AllowNull()][string]$Certainty)
  $map = @{ 'Final'='ok'; 'To confirm'='tc'; 'Conditional - outside R total'='cond'; 'Cost avoidance'='av' }
  if ($map.ContainsKey([string]$Certainty)) { return ('<span class="badge {0}">{1}</span>' -f $map[[string]$Certainty], (ConvertTo-ReportHtml $Certainty)) }
  return (Get-ReportBadge $Certainty)
}

function Get-ReportSectionStart {
  param([Parameter(Mandatory)][string]$Id, [Parameter(Mandatory)][string]$Title, [string]$Hint, [switch]$Open)
  return ('<section id="{0}" class="fold"><details{1}><summary><h2>{2}</h2><span class="hint">{3}</span></summary><div class="fold-body">' -f $Id, $(if ($Open) { ' open' } else { '' }), (ConvertTo-ReportHtml $Title), (ConvertTo-ReportHtml $Hint))
}
$script:SectionEnd = '</div></details></section>'

function New-ReportPage {
  param([Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][string]$Header, [Parameter(Mandatory)][string]$Body)
  return @"
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>$(ConvertTo-ReportHtml $Title)</title>
  <style>
$($script:Css)
  </style>
</head>
<body>
<div class="page">
<header>
$Header
</header>
<main>
$Body
</main>
<footer>SmartFinOps Workplace by WorkplaceCloudHub | Report $($script:ScriptVersion) | Aggregates only: the detail stays in the FinOps CSV files.</footer>
</div>
</body>
</html>
"@
}

function Get-ReportGeneratedLine {
  param([Parameter(Mandatory)][string]$Tenant, [string]$RunId)
  return ('<p class="subtitle">Profile: {0} | FinOps run: {1} | Generated: <span id="generated">{2}</span> UTC | Report {3}</p>' -f (ConvertTo-ReportHtml $Tenant), (ConvertTo-ReportHtml $(if ($RunId) { $RunId } else { 'unknown' })), [datetime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss', $script:Invariant), $script:ScriptVersion)
}

function Get-ReportHeadCard {
  param([Parameter(Mandatory)][string]$Title, [Parameter(Mandatory)][string[]]$Lines)
  $text = ('  <div><div class="pill">SmartFinOps Workplace</div>' + "`n" + "  <h1>$(ConvertTo-ReportHtml $Title)</h1>" + "`n" + (($Lines | ForEach-Object { "  $_" }) -join "`n") + "`n  </div>")
  return ('<div class="card head">' + "`n" + $text + "`n" + ('  <div class="logo">{0}</div>' -f (Get-ReportLogoHtml)) + "`n</div>")
}
#endregion

#region Summary model (shared by the page and the daily email)
function Get-ReportCertainty {
  param([AllowNull()][string]$Status)
  switch -CaseSensitive ([string]$Status) { 'Valued' { return 'Final' } '' { return 'Not measured' } default { return $Status } }
}

function Get-ReportSummaryModel {
  param([Parameter(Mandatory)]$Inputs)
  # Published values only: the cards are FinOps_Summary totals, the decisions the published DETAIL rows in the
  # Priority decisions order, the source states the FinOps_Sources rows. Nothing is added or subtracted.
  $receipt = $Inputs.Receipt
  $summary = @($Inputs.Data['FinOps_Summary.csv'])
  $kpi = [System.Collections.Generic.Dictionary[string,object]]::new([StringComparer]::Ordinal)
  foreach ($row in $summary) { $kpi[[string]$row.KpiKey] = $row }
  $measured = [bool](Get-ReportProperty $receipt 'LicencesMeasured')
  $overall = if (-not $measured) { 'Not measured' } elseif ([bool](Get-ReportProperty $receipt 'Provisional')) { 'Provisional' } else { 'Final' }
  $cards = foreach ($spec in @(
      @{ Key='TOTAL_R_CONFIRMED'; Class='confirmed'; Tag=''; Detail='Disabled accounts, eligible shared mailboxes, Copilot without activity and unassigned units; at contract or renewal' }
      @{ Key='TOTAL_R_ENABLED_INACTIVE'; Class='to-confirm'; Tag='To confirm'; Detail='Confirm with the business before removing or reassigning' }
      @{ Key='TOTAL_RC_E3_TO_F3'; Class='conditional'; Tag='Conditional - outside R total'; Detail='If Frontline eligibility is confirmed' }
      @{ Key='TOTAL_E_REUSE'; Class='avoidance'; Tag='Cost avoidance'; Detail='Purchase avoided; never added to R' })) {
    $row = if ($kpi.ContainsKey($spec.Key)) { $kpi[$spec.Key] } else { $null }
    $status = if ($row) { [string]$row.ValuationStatus } else { 'Not measured' }
    # TOTAL_R_CONFIRMED is published from engine 2.2.3; an older run shows it as Not measured, never as a difference.
    $detail = if ($null -eq $row -and $spec.Key -ceq 'TOTAL_R_CONFIRMED') { 'Published from valuation engine 2.2.3: run the valuation again' } else { $spec.Detail }
    [pscustomobject]@{ Key=$spec.Key; Label=$script:Totals[$spec.Key]; Class=$spec.Class; Row=$row; Status=$status; Certainty=(Get-ReportCertainty $status); Tag=$spec.Tag; Detail=$detail }
  }
  $detailRows = @($summary | Where-Object { ([string]$_.KpiKey).StartsWith('DETAIL|', [StringComparison]::Ordinal) })
  $ranked = @(Get-ReportSorted @($detailRows | Where-Object { $_.ValuationStatus -cin @('Valued','Provisional') -and $_.AnnualAmount }) { param($x) @((ConvertTo-ReportDecimal $x.AnnualAmount), (Get-ReportLeverNumber $x.Lever)) } -Descending)
  $decisions = for ($i = 0; $i -lt $ranked.Count; $i++) {
    $row = $ranked[$i]
    $certainty = if ([string]$row.Lever -ceq 'L3') { 'To confirm' } elseif ([string]$row.ValueClass -ceq 'Rc') { 'Conditional - outside R total' } elseif ([string]$row.ValueClass -ceq 'E') { 'Cost avoidance' } else { Get-ReportCertainty $row.ValuationStatus }
    $entry = $script:Levers[[string]$row.Lever]
    [pscustomobject]@{ Rank=$i + 1; Row=$row; Lever=[string]$row.Lever; Label=$(if ($entry) { $entry[0] } else { [string]$row.Lever }); Action=$(if ($entry) { $entry[1] } else { '' }); Certainty=$certainty }
  }
  $sources = @($Inputs.Data['FinOps_Sources.csv'])
  $qualified = @($sources | Where-Object { [string]$_.QualificationStatus -ceq 'Qualified' }).Count
  $provisional = @($sources | Where-Object { [string]$_.QualificationStatus -ceq 'Provisional' }).Count
  return [pscustomobject]@{
    Measured=$measured; Overall=$overall; Cards=@($cards); TotalR=$(if ($kpi.ContainsKey('TOTAL_R')) { $kpi['TOTAL_R'] } else { $null }); Esu=$(if ($kpi.ContainsKey('TOTAL_E_ESU')) { $kpi['TOTAL_E_ESU'] } else { $null })
    Decisions=@($decisions); Top=@($decisions | Select-Object -First 3)
    Sources=[pscustomobject]@{ Qualified=$qualified; Provisional=$provisional; NotMeasured=($sources.Count - $qualified - $provisional) }
    ProvisionalLevers=@(Get-ReportProperty $receipt 'ProvisionalLevers' | Where-Object { $_ })
  }
}

function Get-ReportSourcesText {
  param([Parameter(Mandatory)]$Model)
  return ('Sources: {0} qualified, {1} provisional, {2} not measured' -f $Model.Sources.Qualified, $Model.Sources.Provisional, $Model.Sources.NotMeasured)
}
#endregion

#region Not available page
function New-ReportUnavailablePage {
  param([Parameter(Mandatory)]$Inputs, [Parameter(Mandatory)][string]$Tenant)
  $header = (Get-ReportHeadCard -Title 'Report not available' -Lines @((Get-ReportGeneratedLine -Tenant $Tenant -RunId $Inputs.RunId)))
  $body = [System.Text.StringBuilder]::new()
  [void]$body.AppendLine('<section><h2>Why</h2>')
  [void]$body.AppendLine(('<div class="banner stop">{0}</div>' -f (ConvertTo-ReportHtml (ConvertTo-ReportSafeText $Inputs.Reason))))
  [void]$body.AppendLine('<p class="note">The latest valuation run is not qualified, so no figure is shown, not even an older one. Fix the cause and run SmartFinOps-Workplace-Valuation.ps1 again, then this report.</p>')
  if ($Inputs.CompletedAtUtc) { [void]$body.AppendLine(('<p class="note">Valuation receipt completed at: {0}</p>' -f (ConvertTo-ReportHtml $Inputs.CompletedAtUtc))) }
  [void]$body.AppendLine('<h3>Report qualification</h3>')
  Add-ReportTable $body @('Rule','Result','Detail') @($Inputs.Checks | ForEach-Object { ,@((ConvertTo-ReportHtml $_.Rule), (ConvertTo-ReportHtml $_.Status), (ConvertTo-ReportHtml (ConvertTo-ReportSafeText $_.Detail))) })
  [void]$body.AppendLine('</section>')
  return New-ReportPage -Title 'SmartFinOps Workplace - Report not available' -Header $header -Body $body.ToString()
}
#endregion

#region Report page
function New-ReportPageContent {
  param([Parameter(Mandatory)]$Inputs, [Parameter(Mandatory)][string]$Tenant, [Parameter(Mandatory)][int]$MinimumGroupUsers)
  $receipt = $Inputs.Receipt
  $summary = @($Inputs.Data['FinOps_Summary.csv'])
  $skus = @($Inputs.Data['FinOps_SkuValuation.csv'])
  $opps = @($Inputs.Data['FinOps_Opportunities.csv'])
  $showback = @($Inputs.Data['FinOps_Showback.csv'])
  $realization = @($Inputs.Data['FinOps_Realization.csv'])
  $sources = @($Inputs.Data['FinOps_Sources.csv'])
  $controls = @($Inputs.Data['FinOps_Controls.csv'])
  $kpi = [System.Collections.Generic.Dictionary[string,object]]::new([StringComparer]::Ordinal)
  foreach ($row in $summary) { $kpi[[string]$row.KpiKey] = $row }
  $get = { param([string]$Key) if ($kpi.ContainsKey($Key)) { $kpi[$Key] } else { $null } }
  $model = Get-ReportSummaryModel -Inputs $Inputs
  $measured = $model.Measured
  $skuNames = @{}; foreach ($sku in $skus) { $skuNames[[string]$sku.SkuId] = [string]$sku.SkuPartNumber }
  $first = @($summary | Where-Object { $_.SnapshotGeneratedAtUtc } | Select-Object -First 1)
  $snapshotDate = if ($first.Count) { [string]$first[0].SnapshotGeneratedAtUtc } else { '' }
  $pricingRun = if ($first.Count) { [string]$first[0].PricingReceiptRunId } else { '' }

  # Header card, status bar and banners.
  $header = [System.Text.StringBuilder]::new()
  [void]$header.AppendLine((Get-ReportHeadCard -Title 'Workplace FinOps report' -Lines @(
    (Get-ReportGeneratedLine -Tenant $Tenant -RunId $Inputs.RunId),
    ('<p class="subtitle">Licences snapshot: {0} | Amounts: EUR excluding VAT, annual run-rate</p>' -f (ConvertTo-ReportHtml $(if ($snapshotDate) { $snapshotDate } else { 'not measured' }))))))
  $overallText = switch -CaseSensitive ($model.Overall) { 'Final' { 'Final figures' } 'Provisional' { 'Provisional figures' } default { 'Licences report not measured' } }
  [void]$header.AppendLine(('<div class="card statusbar">{0} <span>{1}</span> <span class="note">{2}</span></div>' -f (Get-ReportCertaintyBadge $(if ($model.Overall -ceq 'Final') { 'Final' } else { $model.Overall })), (ConvertTo-ReportHtml $overallText), (ConvertTo-ReportHtml (Get-ReportSourcesText $model))))
  if (-not $measured) { [void]$header.AppendLine('<div class="banner stop">Licences report not qualified: every figure is <strong>Not measured</strong>. See the technical appendix for the source qualification.</div>') }
  if ([bool](Get-ReportProperty $receipt 'Provisional')) { [void]$header.AppendLine(('<div class="banner">Provisional figures for levers {0}: a source they use is provisional (see How to read this report).</div>' -f (ConvertTo-ReportHtml ($model.ProvisionalLevers -join ', ')))) }
  $missingDomains = [string](Get-ReportProperty $receipt 'MissingDomains')
  if ($missingDomains) { [void]$header.AppendLine(('<div class="banner">AD domains not collected: {0}. Their users are Not measured in the showback.</div>' -f (ConvertTo-ReportHtml $missingDomains))) }

  $body = [System.Text.StringBuilder]::new()
  $detailRows = @($summary | Where-Object { ([string]$_.KpiKey).StartsWith('DETAIL|', [StringComparison]::Ordinal) })
  $licenseLevers = @('L1','L2','L3','L4','L5','L6','L7')

  #region Summary (always visible)
  [void]$body.AppendLine('<section id="summary"><h2>Summary</h2>')
  [void]$body.AppendLine('<div class="grid">')
  foreach ($card in $model.Cards) {
    $row = $card.Row
    $value = if ($row -and (Format-ReportEur $row.AnnualAmount)) { Format-ReportEur $row.AnnualAmount } else { Get-ReportAmountCell '' $card.Status }
    $badges = @(Get-ReportCertaintyBadge $card.Certainty) + @($(if ($card.Tag) { Get-ReportCertaintyBadge $card.Tag }))
    $meta = if ($row -and ($row.Units -or ([string]$row.Users -and [string]$row.Users -ne '0'))) { 'Units: {0} | Users: {1}' -f $(if ($row.Units) { Format-ReportCount $row.Units } else { '-' }), (Format-ReportCount $row.Users) } else { '' }
    [void]$body.AppendLine(('<div class="metric {0}" data-kpi="{1}"><div class="label">{2}</div><div class="value">{3}</div><div>{4}</div><div class="detail">{5}</div><div class="detail">{6}</div></div>' -f $card.Class, $card.Key, (ConvertTo-ReportHtml $card.Label), $value, ($badges -join ' '), (ConvertTo-ReportHtml $meta), (ConvertTo-ReportHtml $card.Detail)))
  }
  [void]$body.AppendLine('</div>')
  $totalR = $model.TotalR
  $totalText = if ($totalR -and (Format-ReportEur $totalR.AnnualAmount)) { Format-ReportEur $totalR.AnnualAmount } else { Get-ReportAmountCell '' $(if ($totalR) { [string]$totalR.ValuationStatus } else { 'Not measured' }) }
  [void]$body.AppendLine(('<p class="note">Realizable reduction (R) published: {0} = R confirmed + R to confirm (control G7). Conditional E3 to F3 and cost avoidance are never added to R.</p>' -f $totalText))
  [void]$body.AppendLine('<h3>Top 3 decisions</h3>')
  Add-ReportTable $body @('#','Decision','Action','Annual amount','Certainty') @($model.Top | ForEach-Object {
    ,@([string]$_.Rank, (Get-ReportLeverLabel $_.Lever), (ConvertTo-ReportHtml $_.Action), (Format-ReportEur $_.Row.AnnualAmount), (Get-ReportCertaintyBadge $_.Certainty))
  }) @(3) $(if ($measured) { 'No valued decision.' } else { 'Not measured: the Licences report is not qualified.' })
  [void]$body.AppendLine('<h3>License units by status (levers L1 to L7)</h3>')
  $chips = foreach ($status in @('Provisional','Review','Blocked','Capped','Not valued','Not measured')) {
    $units = [long]0; $found = $false
    foreach ($row in @($detailRows | Where-Object { $_.ValuationStatus -ceq $status -and $_.Lever -in $licenseLevers })) { $n = ConvertTo-ReportLong $row.Units; if ($null -ne $n) { $units += $n; $found = $true } }
    '<span class="chip">{0} {1}</span>' -f (Get-ReportBadge $status), $(if ($found) { $units.ToString('#,##0', $script:Invariant) } else { 'None' })
  }
  [void]$body.AppendLine(('<div class="chips">{0}</div>' -f ($chips -join ' ')))
  [void]$body.AppendLine('<p class="note">Provisional units are in the totals with their status; Review, Blocked, Capped, Not valued and Not measured units are not.</p>')
  [void]$body.AppendLine('</section>')
  #endregion

  #region 1. Financial potential
  [void]$body.AppendLine((Get-ReportSectionStart -Id 'financial-potential' -Title '1. Financial potential' -Hint 'Totals by class, price origin, lever and SKU' -Open))
  $kpiRows = foreach ($key in $script:Totals.Keys) {
    $row = & $get $key
    if ($null -eq $row) { ,@((ConvertTo-ReportHtml $script:Totals[$key]), '', '', '', '', (Get-ReportBadge 'Not measured'), ''); continue }
    ,@((ConvertTo-ReportHtml $script:Totals[$key]), (ConvertTo-ReportHtml $row.ValueClass), (ConvertTo-ReportHtml ($row.Lever -replace ';', ', ')), (Format-ReportCount $row.Users), (Format-ReportCount $row.Units), (Get-ReportAmountCell $row.AnnualAmount $row.ValuationStatus), (ConvertTo-ReportHtml (ConvertTo-ReportSafeText $row.Note)))
  }
  Add-ReportTable $body @('Indicator','Class','Levers','Users','Units','Annual amount','Note') @($kpiRows) @(3,4,5)
  [void]$body.AppendLine('<p class="note">R = R confirmed + R to confirm (control G7). The "of which" lines detail R confirmed and are never added to it. E3 to F3 is conditional and cost avoidance is not a reduction.</p>')
  [void]$body.AppendLine('<h3>Price origin</h3>')
  $origins = @('Estimated','Contract','Unknown')
  $priceRows = foreach ($key in @($script:Totals.Keys | Where-Object { $_ -ne 'TOTAL_E_ESU' })) {
    $cells = @((ConvertTo-ReportHtml $script:Totals[$key]))
    foreach ($origin in $origins) { $row = & $get "PRICE|$key|$origin"; $cells += $(if ($row) { Get-ReportAmountCell $row.AnnualAmount $row.ValuationStatus } else { '<span class="none">-</span>' }) }
    ,$cells
  }
  Add-ReportTable $body @('Indicator','Catalogue estimate','Contract price','No price flag') @($priceRows) @(1,2,3)
  [void]$body.AppendLine('<p class="note">Catalogue estimates are public list prices flagged IsEstimated by LicensePricing; contract prices come from the private pricing override.</p>')
  [void]$body.AppendLine('<h3>By lever and status</h3>')
  $classOrder = @{ R=0; Rc=1; E=2; S=3 }
  $sortedDetail = @(Get-ReportSorted $detailRows { param($x) @($(if ($classOrder.ContainsKey([string]$x.ValueClass)) { $classOrder[[string]$x.ValueClass] } else { 9 }), (Get-ReportLeverNumber $x.Lever), [string]$x.ValuationStatus) })
  Add-ReportTable $body @('Class','Lever','Status','Users','Units','Annual amount') @($sortedDetail | Where-Object { $_.Lever -ne 'L12' } | ForEach-Object {
    ,@((ConvertTo-ReportHtml $_.ValueClass), (Get-ReportLeverLabel $_.Lever), (Get-ReportBadge $_.ValuationStatus), (Format-ReportCount $_.Users), (Format-ReportCount $_.Units), $(if ($_.AnnualAmount) { Get-ReportAmountCell $_.AnnualAmount $_.ValuationStatus } else { '' }))
  }) @(3,4,5) 'No opportunity: the Licences report is not measured.'
  [void]$body.AppendLine('<h3>By SKU</h3>')
  $skuSorted = @(Get-ReportSorted @($skus | Where-Object { $_.Excluded -ne 'True' -and ((ConvertTo-ReportLong $_.CandidateUnitsR) -gt 0 -or (ConvertTo-ReportLong $_.CapacityRealizableUnits) -gt 0 -or (ConvertTo-ReportLong $_.ReusedUnits) -gt 0 -or (ConvertTo-ReportLong $_.CappedUnits) -gt 0) }) { param($x) @($(if ($null -ne (ConvertTo-ReportDecimal $x.RealizableAnnualAmount)) { ConvertTo-ReportDecimal $x.RealizableAnnualAmount } else { [decimal]-1 }), [string]$x.SkuPartNumber) } -Descending)
  Add-ReportBars $body 'Realizable reduction by SKU (top 10)' @($skuSorted | Select-Object -First 10 | ForEach-Object { [pscustomobject]@{ Label=$_.SkuPartNumber; Value=(ConvertTo-ReportDecimal $_.RealizableAnnualAmount); Text=(Get-ReportEurText $_.RealizableAnnualAmount) } })
  Add-ReportTable $body @('SKU','Product','Realizable units','Capped','Reused','Unit price / month','Price origin','Realizable annual','Avoided annual','Status') @($skuSorted | ForEach-Object {
    $price = ConvertTo-ReportDecimal $_.UnitPriceMonthly
    ,@(('{0}<br><span class="note">{1}</span>' -f (ConvertTo-ReportHtml $_.SkuPartNumber), (ConvertTo-ReportHtml $_.SkuDisplayName)), (ConvertTo-ReportHtml $_.ReportProduct), (Format-ReportCount $_.RealizableUnits), (Format-ReportCount $_.CappedUnits), (Format-ReportCount $_.ReusedUnits),
      $(if ($null -ne $price) { 'EUR ' + $price.ToString('0.00', $script:Invariant) } else { '' }), (ConvertTo-ReportHtml (Get-ReportPriceOrigin $_.PriceIsEstimated)), (Get-ReportAmountCell $_.RealizableAnnualAmount $_.ValuationStatus), (Format-ReportEur $_.AvoidedAnnualAmount), (Get-ReportBadge $_.ValuationStatus))
  }) @(2,3,4,5,7,8) 'No SKU with a reduction or reuse.'
  [void]$body.AppendLine($script:SectionEnd)
  #endregion

  #region 2. Priority decisions
  [void]$body.AppendLine((Get-ReportSectionStart -Id 'priority-decisions' -Title '2. Priority decisions' -Hint 'Every valued decision, largest annual amount first'))
  Add-ReportBars $body 'Annual amount by lever' @($model.Decisions | ForEach-Object { [pscustomobject]@{ Label=('{0} ({1})' -f $_.Label, $_.Lever); Value=(ConvertTo-ReportDecimal $_.Row.AnnualAmount); Text=(Get-ReportEurText $_.Row.AnnualAmount) } })
  Add-ReportTable $body @('#','Decision','Class','Action','Units','Annual amount','Certainty') @($model.Decisions | ForEach-Object {
    ,@([string]$_.Rank, (Get-ReportLeverLabel $_.Lever), (ConvertTo-ReportHtml $_.Row.ValueClass), (ConvertTo-ReportHtml $_.Action), (Format-ReportCount $_.Row.Units), (Get-ReportAmountCell $_.Row.AnnualAmount $_.Row.ValuationStatus), (Get-ReportCertaintyBadge $_.Certainty))
  }) @(4,5) 'No valued decision.'
  [void]$body.AppendLine('<p class="note">Rc (E3 to F3) is conditional and E is cost avoidance: neither is part of the R total.</p>')
  [void]$body.AppendLine('<h3>Top SKUs</h3>')
  Add-ReportTable $body @('SKU','Realizable annual','Realizable from','Contract status','Reduction policy') @($skuSorted | Where-Object { $null -ne (ConvertTo-ReportDecimal $_.RealizableAnnualAmount) -and (ConvertTo-ReportDecimal $_.RealizableAnnualAmount) -gt 0 } | Select-Object -First 10 | ForEach-Object {
    ,@((ConvertTo-ReportHtml $_.SkuPartNumber), (Get-ReportAmountCell $_.RealizableAnnualAmount $_.ValuationStatus), (ConvertTo-ReportHtml $_.RealizableFrom), (ConvertTo-ReportHtml $_.ContractStatus), (ConvertTo-ReportHtml $_.ReductionPolicy))
  }) @(1) 'No SKU with a valued reduction.'
  [void]$body.AppendLine($script:SectionEnd)
  #endregion

  #region 3. Decisions not in the totals
  [void]$body.AppendLine((Get-ReportSectionStart -Id 'trade-offs' -Title '3. Decisions not in the totals' -Hint 'Review, blocked, not measured, not valued, capped, signals and gap metrics'))
  [void]$body.AppendLine('<p class="note">These decisions are not in the totals: they wait for a review, a dependency, a measure or a price.</p>')
  $held = @('Review','Blocked','Not measured','Not valued','Capped','Reused','Signal')
  Add-ReportTable $body @('Lever','Class','Status','Users','Units') @($sortedDetail | Where-Object { $_.ValuationStatus -cin $held -and $_.Lever -ne 'L12' } | ForEach-Object {
    ,@((Get-ReportLeverLabel $_.Lever), (ConvertTo-ReportHtml $_.ValueClass), (Get-ReportBadge $_.ValuationStatus), (Format-ReportCount $_.Users), (Format-ReportCount $_.Units))
  }) @(3,4)
  [void]$body.AppendLine('<h3>Reasons</h3>')
  $reasonGroups = @{}
  foreach ($row in @($opps | Where-Object { $_.ValuationStatus -cin @('Review','Blocked','Not measured','Not valued') -and $_.Lever -ne 'L12' })) {
    $reason = ConvertTo-ReportSafeText $row.StatusReason
    if ($reason -match '\[(user|id)\]') { $reason = 'Other reason' }
    $key = '{0}|{1}|{2}' -f $row.Lever, $row.ValuationStatus, $reason
    if (-not $reasonGroups.ContainsKey($key)) { $reasonGroups[$key] = [pscustomobject]@{ Lever=[string]$row.Lever; Status=[string]$row.ValuationStatus; Reason=$reason; Rows=0; Units=[long]0; HasUnits=$false } }
    $group = $reasonGroups[$key]; $group.Rows++
    $n = ConvertTo-ReportLong $row.Units; if ($null -ne $n) { $group.Units += $n; $group.HasUnits = $true }
  }
  Add-ReportTable $body @('Lever','Status','Reason','Rows','Units') @(Get-ReportSorted @($reasonGroups.Values) { param($x) @((Get-ReportLeverNumber $x.Lever), $x.Status, $x.Reason) } | ForEach-Object {
    ,@((Get-ReportLeverLabel $_.Lever), (Get-ReportBadge $_.Status), (ConvertTo-ReportHtml $(if ($_.Reason) { $_.Reason } else { 'No reason given' })), $_.Rows.ToString('#,##0', $script:Invariant), $(if ($_.HasUnits) { $_.Units.ToString('#,##0', $script:Invariant) } else { '' }))
  }) @(3,4)
  [void]$body.AppendLine('<h3>Over-assigned SKUs</h3>')
  Add-ReportTable $body @('SKU','Over-assigned units','Enabled','Consumed') @(Get-ReportSorted @($skus | Where-Object { (ConvertTo-ReportLong $_.AvailableUnits) -lt 0 }) { param($x) @([string]$x.SkuPartNumber) } | ForEach-Object {
    ,@((ConvertTo-ReportHtml $_.SkuPartNumber), ([math]::Abs([long](ConvertTo-ReportLong $_.AvailableUnits))).ToString('#,##0', $script:Invariant), (Format-ReportCount $_.EnabledUnits), (Format-ReportCount $_.ConsumedUnits))
  }) @(1,2,3) 'No over-assigned SKU.'
  [void]$body.AppendLine('<h3>Mailbox and AD gap metrics (never added together)</h3>')
  Add-ReportTable $body @('Section','Metric','Value','Status') @(Get-ReportSorted @($summary | Where-Object { ([string]$_.KpiKey).StartsWith('GAP|', [StringComparison]::Ordinal) }) { param($x) @([string]$x.KpiKey) } | ForEach-Object {
    $parts = ([string]$_.KpiKey).Split('|', 3)
    ,@((ConvertTo-ReportHtml $parts[1]), (ConvertTo-ReportHtml $parts[2]), $(if ($_.Units -ne '') { Format-ReportCount $_.Units } else { Get-ReportBadge 'Not measured' }), (Get-ReportBadge $_.ValuationStatus))
  }) @(2) 'No gap metric.'
  [void]$body.AppendLine($script:SectionEnd)
  #endregion

  #region 4. Capacity, renewals and forecast
  [void]$body.AppendLine((Get-ReportSectionStart -Id 'capacity-forecast' -Title '4. Capacity, renewals and forecast' -Hint 'When each amount becomes realizable, capacity and contracts per SKU'))
  $forecastKeys = @('TOTAL_R','TOTAL_RC_E3_TO_F3','TOTAL_E_REUSE')
  $dates = @($summary | ForEach-Object { [string]$_.KpiKey } | Where-Object { $_ -match '^FORECAST\|(TOTAL_R|TOTAL_RC_E3_TO_F3|TOTAL_E_REUSE)\|' } | ForEach-Object { $_.Split('|', 3)[2] } | Sort-Object -Unique)
  $dates = @(Get-ReportSorted $dates { param($x) @($(if ($x -ceq 'Renewal unknown') { 1 } else { 0 }), [string]$x) })
  Add-ReportBars $body 'Realizable reduction by date' @($dates | ForEach-Object { $row = & $get "FORECAST|TOTAL_R|$_"; if ($row) { [pscustomobject]@{ Label=$_; Value=(ConvertTo-ReportDecimal $row.AnnualAmount); Text=(Get-ReportEurText $row.AnnualAmount) } } })
  Add-ReportTable $body @('Realizable from','Reduction (R)','Conditional E3 to F3','Cost avoidance (reuse)') @($dates | ForEach-Object {
    $date = $_
    $cells = @($(if ($date -ceq 'Renewal unknown') { '<strong>Renewal unknown</strong>' } else { ConvertTo-ReportHtml $date }))
    foreach ($key in $forecastKeys) { $row = & $get "FORECAST|$key|$date"; $cells += $(if ($row) { Get-ReportAmountCell $row.AnnualAmount $row.ValuationStatus } else { '<span class="none">-</span>' }) }
    ,$cells
  }) @(1,2,3) 'No valued amount to schedule.'
  [void]$body.AppendLine('<p class="note">Realizable from: renewal date for AtRenewal or Anniversary contracts, run date for Monthly, otherwise Renewal unknown (complete the contracts file).</p>')
  [void]$body.AppendLine('<h3>Capacity and contracts</h3>')
  $active = @(Get-ReportSorted @($skus | Where-Object { $_.Excluded -ne 'True' }) { param($x) @([string]$x.SkuPartNumber) })
  Add-ReportTable $body @('SKU','Enabled','Consumed','Available','Reserve','Minimum','Cap','Contract','Contract status','Policy','Renewal','Realizable from','Usage evidence') @($active | ForEach-Object {
    $available = ConvertTo-ReportLong $_.AvailableUnits
    ,@((ConvertTo-ReportHtml $_.SkuPartNumber), (Format-ReportCount $_.EnabledUnits), (Format-ReportCount $_.ConsumedUnits),
      $(if ($null -ne $available -and $available -lt 0) { 'Over-assigned ' + ([math]::Abs($available)).ToString('#,##0', $script:Invariant) } else { Format-ReportCount $_.AvailableUnits }),
      (Format-ReportCount $_.ReserveUnits), (Format-ReportCount $_.MinimumUnits), (Format-ReportCount $_.CapUnits), (ConvertTo-ReportHtml $_.ContractKey), (ConvertTo-ReportHtml $_.ContractStatus), (ConvertTo-ReportHtml $_.ReductionPolicy),
      (ConvertTo-ReportHtml $_.RenewalDate), $(if ($_.RealizableFrom -ceq 'Renewal unknown') { '<strong>Renewal unknown</strong>' } else { ConvertTo-ReportHtml $_.RealizableFrom }), (ConvertTo-ReportHtml $_.UsageEvidence))
  }) @(1,2,3,4,5,6) 'No SKU measured.'
  $excluded = @($skus | Where-Object Excluded -eq 'True').Count
  if ($excluded) { [void]$body.AppendLine(('<p class="note">{0} excluded SKU(s) (free or system) are not shown.</p>' -f $excluded)) }
  [void]$body.AppendLine($script:SectionEnd)
  #endregion

  #region 5. Realization
  [void]$body.AppendLine((Get-ReportSectionStart -Id 'realization' -Title '5. Realization' -Hint 'Savings observed since a qualified baseline run'))
  $measuredRows = @($realization | Where-Object { $_.Status -cin @('Realized','No reduction observed') })
  if ($measuredRows.Count) {
    Add-ReportTable $body @('SKU','Baseline run','Enabled before','Enabled after','Decrease','Removed candidate licenses','Baseline capacity','Realized units','Not attributed','Realized annual','Status') @(Get-ReportSorted $measuredRows { param($x) @([string]$x.SkuPartNumber) } | ForEach-Object {
      ,@((ConvertTo-ReportHtml $_.SkuPartNumber), (ConvertTo-ReportHtml $_.BaselineFinOpsRunId), (Format-ReportCount $_.EnabledUnitsBefore), (Format-ReportCount $_.EnabledUnitsAfter), (Format-ReportCount $_.Decrease), (Format-ReportCount $_.RemovedCandidateLicenses), (Format-ReportCount $_.BaselineCapacityUnits), (Format-ReportCount $_.RealizedUnits), (Format-ReportCount $_.NotAttributedUnits), (Format-ReportEur $_.RealizedAnnualAmount), (Get-ReportBadge $_.Status))
    }) @(2,3,4,5,6,7,8,9)
  }
  else { [void]$body.AppendLine('<p class="empty">No qualified baseline run yet: realized savings are measured from the next qualified runs.</p>') }
  $byStatus = @($realization | Group-Object -Property Status -CaseSensitive | ForEach-Object { [pscustomobject]@{ Status=$_.Name; Count=$_.Count } })
  Add-ReportTable $body @('Realization status','SKUs') @(Get-ReportSorted $byStatus { param($x) @([string]$x.Status) } | ForEach-Object { ,@((ConvertTo-ReportHtml (ConvertTo-ReportSafeText $_.Status)), $_.Count.ToString('#,##0', $script:Invariant)) }) @(1) 'No SKU measured.'
  [void]$body.AppendLine($script:SectionEnd)
  #endregion

  #region 6. Showback
  [void]$body.AppendLine((Get-ReportSectionStart -Id 'showback' -Title '6. Showback' -Hint 'By AD domain, country, entity and department; small groups combined'))
  [void]$body.AppendLine(('<p class="note">Values with fewer than {0} users are combined into "Other (groups under {0} users)" so that no group can be deduced; the detail stays in the FinOps CSV files. Users are distinct users with at least one decision row.</p>' -f $MinimumGroupUsers))
  $columns = @(@('TOTAL_R','Reduction (R)'), @('TOTAL_R_CONFIRMED','R confirmed'), @('TOTAL_R_ENABLED_INACTIVE','R to confirm'), @('TOTAL_RC_E3_TO_F3','E3 to F3 (Rc)'))
  $countStatuses = @('Review','Blocked','Not measured','Not valued')
  foreach ($dimension in @('AdDomain','Country','Entity','Department')) {
    $label = @{ AdDomain='AD domain'; Country='Country'; Entity='Entity'; Department='Department' }[$dimension]
    $groups = @{}
    $ensure = { param([string]$Value) if (-not $groups.ContainsKey($Value)) { $groups[$Value] = [pscustomobject]@{ Value=$Value; Users=$null; Amounts=@{}; Statuses=@{}; Counts=@{}; Members=1 } }; $groups[$Value] }
    foreach ($row in @($showback | Where-Object Dimension -ceq $dimension)) {
      $group = & $ensure ([string]$row.DimensionValue)
      if ($row.ValuationStatus -cin $countStatuses) { $n = ConvertTo-ReportLong $row.Units; if ($null -ne $n) { $group.Counts[[string]$row.ValuationStatus] = [long]$(if ($group.Counts.ContainsKey([string]$row.ValuationStatus)) { $group.Counts[[string]$row.ValuationStatus] } else { 0 }) + $n } }
    }
    foreach ($row in $summary) {
      $key = [string]$row.KpiKey
      if (-not $key.StartsWith("SHOWBACK|$dimension|", [StringComparison]::Ordinal)) { continue }
      $parts = $key.Split('|', 4)
      $group = & $ensure $parts[3]
      if ($parts[2] -ceq 'USERS') { $group.Users = ConvertTo-ReportLong $row.Users; continue }
      $amount = ConvertTo-ReportDecimal $row.AnnualAmount
      if ($null -ne $amount) { $group.Amounts[$parts[2]] = $amount; $group.Statuses[$parts[2]] = [string]$row.ValuationStatus }
    }
    $all = @($groups.Values)
    $small = [System.Collections.Generic.List[object]]::new()
    foreach ($group in @($all | Where-Object { $_.Value -cne 'SKU level' -and [long]$_.Users -lt $MinimumGroupUsers })) { $small.Add($group) }
    if ($small.Count) {
      # The combined row must itself reach the threshold: add the smallest remaining groups until it does.
      $smallUsers = [long]0; foreach ($group in $small) { $smallUsers += [long]$group.Users }
      $rest = @(Get-ReportSorted @($all | Where-Object { $_.Value -cne 'SKU level' -and $small -notcontains $_ }) { param($x) @([decimal][long]$x.Users, [string]$x.Value) })
      foreach ($group in $rest) { if ($smallUsers -ge $MinimumGroupUsers) { break }; $small.Add($group); $smallUsers += [long]$group.Users }
      $other = [pscustomobject]@{ Value=("Other (groups under {0} users)" -f $MinimumGroupUsers); Users=$smallUsers; Amounts=@{}; Statuses=@{}; Counts=@{}; Members=$small.Count }
      foreach ($group in $small) {
        foreach ($k in $group.Amounts.Keys) { $other.Amounts[$k] = [decimal]$(if ($other.Amounts.ContainsKey($k)) { $other.Amounts[$k] } else { 0 }) + $group.Amounts[$k]; if ($group.Statuses[$k] -ceq 'Provisional' -or -not $other.Statuses.ContainsKey($k)) { $other.Statuses[$k] = $group.Statuses[$k] } }
        foreach ($k in $group.Counts.Keys) { $other.Counts[$k] = [long]$(if ($other.Counts.ContainsKey($k)) { $other.Counts[$k] } else { 0 }) + $group.Counts[$k] }
      }
      $all = @($all | Where-Object { $small -notcontains $_ }) + @($other)
    }
    $rank = { param($x) if ($x.Value -cin $script:ReservedValues) { 1 } elseif ($x.Value -ceq 'SKU level') { 2 } elseif ($x.Value -like 'Other (groups under*') { 3 } else { 0 } }
    $sortedGroups = @(Get-ReportSorted $all { param($x) @((& $rank $x), $(if ($x.Amounts.ContainsKey('TOTAL_R')) { -1 * $x.Amounts['TOTAL_R'] } else { [decimal]1 }), [string]$x.Value) })
    [void]$body.AppendLine(('<details class="sub"><summary>By {0}</summary>' -f (ConvertTo-ReportHtml $label)))
    Add-ReportTable $body (@($label, 'Users') + @($columns | ForEach-Object { $_[1] }) + $countStatuses) @($sortedGroups | ForEach-Object {
      $group = $_
      $cells = @((ConvertTo-ReportHtml (ConvertTo-ReportSafeText $group.Value)), $(if ($null -ne $group.Users) { ([long]$group.Users).ToString('#,##0', $script:Invariant) } else { '' }))
      foreach ($column in $columns) { $cells += $(if ($group.Amounts.ContainsKey($column[0])) { Get-ReportAmountCell $group.Amounts[$column[0]].ToString('0.00', $script:Invariant) $group.Statuses[$column[0]] } else { '' }) }
      foreach ($status in $countStatuses) { $cells += $(if ($group.Counts.ContainsKey($status)) { ([long]$group.Counts[$status]).ToString('#,##0', $script:Invariant) } else { '' }) }
      ,$cells
    }) @(1,2,3,4,5,6,7,8,9) 'Not measured for this dimension.'
    [void]$body.AppendLine('</details>')
  }
  [void]$body.AppendLine('<p class="note">SKU level: unassigned units and reuse, not allocated to users. Not attributed: no AD or Entra match. Not measured: AD domain not collected or source not qualified.</p>')
  [void]$body.AppendLine($script:SectionEnd)
  #endregion

  #region 7. How to read this report
  $settings = Get-ReportProperty $receipt 'Settings'
  [void]$body.AppendLine((Get-ReportSectionStart -Id 'assumptions' -Title '7. How to read this report' -Hint 'Assumptions, classes, statuses and levers'))
  [void]$body.AppendLine('<ul>')
  [void]$body.AppendLine('<li>Amounts are an annual run-rate in EUR excluding VAT: monthly unit price x 12, from the date shown in Realizable from.</li>')
  [void]$body.AppendLine('<li>R is realizable at contract or renewal and capped by the contract minimum: R confirmed needs no business check, R to confirm (enabled accounts without activity) does. Rc (E3 to F3) needs a Frontline eligibility check; E is a purchase avoided and is never added to R; signals are never valued.</li>')
  [void]$body.AppendLine('<li>Not measured means a source or a guard was not qualified: no amount, never zero. Not valued means no single effective EUR price.</li>')
  [void]$body.AppendLine(('<li>Price date (snapshot): {0}. Freshness: daily sources within {1} h, weekly within {2} days. Contracts file SHA256: {3}.</li>' -f (ConvertTo-ReportHtml (ConvertTo-ReportJsonText (Get-ReportProperty $settings 'AsOfDate') 'yyyy-MM-dd')), (ConvertTo-ReportHtml (Get-ReportProperty $settings 'FreshnessDailyHours')), (ConvertTo-ReportHtml (Get-ReportProperty $settings 'FreshnessWeeklyDays')), (ConvertTo-ReportHtml $(if ([string](Get-ReportProperty $settings 'ContractsFileSHA256')) { ([string](Get-ReportProperty $settings 'ContractsFileSHA256')).Substring(0, 12) } else { 'none' }))))
  $priced = @($skus | Where-Object { $_.UnitPriceMonthly })
  [void]$body.AppendLine(('<li>SKU prices: {0} catalogue estimate(s), {1} contract price(s), {2} SKU(s) without price.</li>' -f @($priced | Where-Object PriceIsEstimated -eq 'True').Count, @($priced | Where-Object PriceIsEstimated -eq 'False').Count, @($skus | Where-Object { -not $_.UnitPriceMonthly -and $_.Excluded -ne 'True' }).Count))
  foreach ($source in @(Get-ReportSorted @($sources | Where-Object { $_.ProvisionalLevers }) { param($x) @([string]$x.SourceFile) })) {
    [void]$body.AppendLine(('<li>Provisional source {0} affects {1}: {2}</li>' -f (ConvertTo-ReportHtml $source.SourceFile), (ConvertTo-ReportHtml ($source.ProvisionalLevers -replace ';', ', ')), (ConvertTo-ReportHtml (ConvertTo-ReportSafeText $(if ($source.Reason) { $source.Reason } else { 'no reason given' })))))
  }
  if ($missingDomains) { [void]$body.AppendLine(('<li>AD domains not collected: {0}.</li>' -f (ConvertTo-ReportHtml $missingDomains))) }
  [void]$body.AppendLine('</ul>')
  [void]$body.AppendLine('<h3>Certainty and statuses</h3>')
  Add-ReportTable $body @('Label','Meaning') @(
    ,@((Get-ReportCertaintyBadge 'Final'), 'Valued from qualified sources and a single effective price.')
    ,@((Get-ReportCertaintyBadge 'To confirm'), 'Valued, but the business must confirm before the license is removed or reassigned.')
    ,@((Get-ReportCertaintyBadge 'Conditional - outside R total'), 'E3 to F3: valued only if Frontline eligibility is confirmed; outside the R total.')
    ,@((Get-ReportCertaintyBadge 'Cost avoidance'), 'Purchase avoided by reusing units; never a reduction.')
    ,@((Get-ReportBadge 'Provisional'), 'A source used by the lever is provisional: amount shown with this status.')
    ,@((Get-ReportBadge 'Not measured'), 'Source or guard not qualified: no amount, never zero.')
    ,@((Get-ReportBadge 'Not valued'), 'No single effective EUR price.')
    ,@((Get-ReportBadge 'Review'), 'Waits for a review; not in the totals.')
    ,@((Get-ReportBadge 'Blocked'), 'Blocked by a dependency (for example a mailbox hold); not in the totals.')
    ,@((Get-ReportBadge 'Capped'), 'Beyond the contract minimum: not realizable.')
    ,@((Get-ReportBadge 'Signal'), 'Information never monetized.'))
  [void]$body.AppendLine('<h3>Levers</h3>')
  Add-ReportTable $body @('Lever','Meaning','Action') @($script:Levers.Keys | ForEach-Object { ,@((ConvertTo-ReportHtml $_), (ConvertTo-ReportHtml $script:Levers[$_][0]), (ConvertTo-ReportHtml $script:Levers[$_][1])) })
  [void]$body.AppendLine($script:SectionEnd)
  #endregion

  #region 8. Technical appendix
  [void]$body.AppendLine((Get-ReportSectionStart -Id 'appendix' -Title '8. Technical appendix' -Hint 'Report qualification, sources and equivalence controls'))
  [void]$body.AppendLine(('<p class="note">Licences snapshot: {0} | Pricing run: {1}</p>' -f (ConvertTo-ReportHtml $(if ($snapshotDate) { $snapshotDate } else { 'not measured' })), (ConvertTo-ReportHtml $(if ($pricingRun) { $pricingRun } else { 'not qualified' }))))
  [void]$body.AppendLine('<h3>Report qualification</h3>')
  Add-ReportTable $body @('Rule','Result','Detail') @($Inputs.Checks | ForEach-Object { ,@((ConvertTo-ReportHtml $_.Rule), (ConvertTo-ReportHtml $_.Status), (ConvertTo-ReportHtml (ConvertTo-ReportSafeText $_.Detail))) })
  [void]$body.AppendLine('<h3>Sources</h3>')
  Add-ReportTable $body @('Source','Producer','Receipt run','Receipt','Completed (UTC)','Age h / max','Rows','SHA256','Qualification','Reason','Provisional levers') @(Get-ReportSorted $sources { param($x) @([string]$x.SourceFile) } | ForEach-Object {
    $hash = [string]$_.ActualSHA256
    ,@((ConvertTo-ReportHtml $_.SourceFile), (ConvertTo-ReportHtml $_.Producer), (ConvertTo-ReportHtml (ConvertTo-ReportSafeText $_.SourceReceiptRunId)), (ConvertTo-ReportHtml $_.ReceiptStatus), (ConvertTo-ReportHtml $_.ReceiptCompletedAtUtc),
      (ConvertTo-ReportHtml ('{0} / {1}' -f $_.AgeHours, $_.FreshnessThresholdHours)), (ConvertTo-ReportHtml ('{0} / {1}' -f $_.ActualRows, $_.ExpectedRows)), (ConvertTo-ReportHtml $(if ($hash.Length -ge 12) { $hash.Substring(0, 12) } else { $hash })),
      $(switch -CaseSensitive ([string]$_.QualificationStatus) { 'Qualified' { '<span class="badge ok">Qualified</span>' } 'Provisional' { Get-ReportBadge 'Provisional' } default { Get-ReportBadge 'Not measured' } }), (ConvertTo-ReportHtml (ConvertTo-ReportSafeText $_.Reason)), (ConvertTo-ReportHtml $_.ProvisionalLevers))
  }) @() 'No source recorded.'
  [void]$body.AppendLine('<h3>Equivalence controls</h3>')
  $controlSummary = @($controls | Group-Object -Property Control -CaseSensitive | ForEach-Object { [pscustomobject]@{ Control=$_.Name; Pass=@($_.Group | Where-Object Status -ceq 'Pass').Count; Other=@($_.Group | Where-Object Status -cne 'Pass').Count } })
  Add-ReportTable $body @('Control','Pass','Fail or not applicable') @(Get-ReportSorted $controlSummary { param($x) @([string]$x.Control) } | ForEach-Object { ,@((ConvertTo-ReportHtml $_.Control), $_.Pass.ToString($script:Invariant), $_.Other.ToString($script:Invariant)) }) @(1,2)
  $mapKey = { param([string]$Text) (@(([string]$Text).Split('|') | ForEach-Object { if ($skuNames.ContainsKey($_)) { $skuNames[$_] } else { $_ } }) -join '|') }
  [void]$body.AppendLine('<details class="sub"><summary>Every control</summary>')
  Add-ReportTable $body @('Control','Key','Source','FinOps','Status','Detail') @(Get-ReportSorted $controls { param($x) @([string]$x.Control, [string]$x.Key) } | ForEach-Object {
    ,@((ConvertTo-ReportHtml $_.Control), (ConvertTo-ReportHtml (ConvertTo-ReportSafeText (& $mapKey $_.Key))), (ConvertTo-ReportHtml $_.SourceCount), (ConvertTo-ReportHtml $_.FinOpsCount), (ConvertTo-ReportHtml $_.Status), (ConvertTo-ReportHtml (ConvertTo-ReportSafeText $_.Detail)))
  }) @(2,3)
  [void]$body.AppendLine('</details>')
  [void]$body.AppendLine($script:SectionEnd)
  #endregion

  return New-ReportPage -Title 'SmartFinOps Workplace - FinOps report' -Header $header.ToString().TrimEnd() -Body $body.ToString()
}
#endregion

#region Daily email
function Write-ReportDailyMailState {
  param([Parameter(Mandatory)][System.IO.FileStream]$Stream, [Parameter(Mandatory)][AllowEmptyCollection()][byte[]]$Bytes)
  $Stream.SetLength(0)
  $Stream.Position = 0
  $Stream.Write($Bytes, 0, $Bytes.Length)
  $Stream.Flush($true)
}

function Enter-ReportDailyMailGate {
  param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$TenantKey, [switch]$Force, [datetimeoffset]$NowUtc = [datetimeoffset]::UtcNow)
  # Same gate as the SmartM365 Licences summary: one mail per Europe/Paris day, exclusive state file, Pending before sending.
  if (-not (Test-Path -LiteralPath (Split-Path -Path $Path -Parent) -PathType Container)) { throw "Daily email state folder does not exist: $Path" }
  $timeZone = try { [TimeZoneInfo]::FindSystemTimeZoneById('Romance Standard Time') } catch { [TimeZoneInfo]::FindSystemTimeZoneById('Europe/Paris') }
  $day = [TimeZoneInfo]::ConvertTime($NowUtc, $timeZone).ToString('yyyy-MM-dd', $script:Invariant)
  $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
  try {
    if ($stream.Length -gt 16384) { throw "Daily email state is unexpectedly large: $Path" }
    $previous = [byte[]]::new([int]$stream.Length)
    if ($previous.Length -gt 0) {
      $stream.Position = 0
      $read = 0
      while ($read -lt $previous.Length) {
        $count = $stream.Read($previous, $read, $previous.Length - $read)
        if ($count -eq 0) { throw "Daily email state could not be read: $Path" }
        $read += $count
      }
      $state = $null
      try {
        $state = [Text.Encoding]::UTF8.GetString($previous) | ConvertFrom-Json -AsHashtable -ErrorAction Stop
        if ($state -isnot [System.Collections.IDictionary] -or -not $state.Contains('TenantKey') -or -not $state.Contains('LocalDate') -or
            [string]$state.Status -notin @('Sent', 'Pending') -or [string]$state.LocalDate -notmatch '^\d{4}-\d{2}-\d{2}$') { throw 'Invalid daily email state structure.' }
      }
      catch {
        if (-not $Force) { throw "Daily email state is unreadable; use -ForceEmail after checking delivery: $Path" }
        $state = $null
      }
      if ($state -and [string]$state.TenantKey -cne $TenantKey) { throw "Daily email state belongs to another tenant: $Path" }
      if ($state -and [string]$state.LocalDate -eq $day -and -not $Force) {
        $reason = if ([string]$state.Status -eq 'Sent') { "already sent on $day" } else { "delivery is pending or uncertain on $day" }
        $stream.Dispose()
        return [pscustomobject]@{ Skip=$true; Reason=$reason; Stream=$null; Day=$day; PreviousBytes=$previous }
      }
    }
    $pending = [ordered]@{ TenantKey=$TenantKey; LocalDate=$day; Status='Pending'; LastAttemptUtc=$NowUtc.ToString('o') }
    Write-ReportDailyMailState -Stream $stream -Bytes ([Text.UTF8Encoding]::new($false).GetBytes(($pending | ConvertTo-Json -Compress)))
    return [pscustomobject]@{ Skip=$false; Reason=''; Stream=$stream; Day=$day; PreviousBytes=$previous }
  }
  catch {
    $stream.Dispose()
    throw
  }
}

function Test-ReportIdentifierFree {
  param([AllowNull()][string]$Text)
  # The mail and its attachment carry aggregates only: no e-mail address or UPN, no GUID.
  if ([string]::IsNullOrEmpty($Text)) { return $true }
  if ($Text -match '[A-Za-z0-9._%+''-]+@[A-Za-z0-9-]+(\.[A-Za-z0-9-]+)+') { return $false }
  return -not ($Text -match '(?i)\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b')
}

function Get-ReportDailyEmailContent {
  param([Parameter(Mandatory)]$Inputs, [Parameter(Mandatory)][string]$Tenant, [string]$ReportUrl)
  # Short body built from the same summary model as the page (published values only, nothing added); the SmartM365
  # transport wraps it in the SmartM365 mail frame. Inline styles and tables only, for mail clients.
  $p = 'margin:0 0 10px;font-size:13px;line-height:19px;color:#1f2937;'
  $muted = 'margin:6px 0 14px;font-size:11px;line-height:16px;color:#5f6b7a;'
  $h = 'margin:16px 0 8px;font-size:15px;line-height:21px;color:#1f2937;'
  $badgeColors = @{ 'Final'=@('#1d6b3a','#eaf7ef','#9fd2b1'); 'Provisional'=@('#8a5a00','#fff8e6','#f0c36d'); 'To confirm'=@('#8a5a00','#fff8e6','#f0c36d'); 'Conditional - outside R total'=@('#0f5e66','#ebf7f8','#9fd0d4'); 'Cost avoidance'=@('#374151','#ffffff','#d1d5db') }
  $badge = { param([string]$Text)
    $c = if ($badgeColors.ContainsKey($Text)) { $badgeColors[$Text] } else { @('#4b5563','#f3f4f6','#c7ccd3') }
    '<span style="display:inline-block;padding:1px 6px;border:1px solid {2};background:{1};color:{0};font-size:11px;font-weight:700;">{3}</span>' -f $c[0], $c[1], $c[2], (ConvertTo-ReportHtml $Text) }
  $link = if ($ReportUrl) { '<p style="{0}">Full report on SharePoint: <a href="{1}" style="color:#0078d4;">{2}</a> (also attached).</p>' -f $p, (ConvertTo-ReportHtml $ReportUrl), (ConvertTo-ReportHtml $script:ReportName) } else { '<p style="{0}">The full report is attached.</p>' -f $p }
  if (-not $Inputs.Available) {
    $reason = ConvertTo-ReportHtml (ConvertTo-ReportSafeText $Inputs.Reason)
    return [pscustomobject]@{ Status='NotAvailable'; Subject='SmartFinOps Workplace - Daily valuation - report not available'; Attach=$false
      BodyHtml=('<p style="{0}"><strong>Report not available.</strong> No figure is published for profile {1} today.</p><p style="{0}">Reason: {2}</p><p style="{3}">The last valuation run did not qualify; no older figure is reused.</p>' -f $p, (ConvertTo-ReportHtml $Tenant), $reason, $muted) }
  }
  $model = Get-ReportSummaryModel -Inputs $Inputs
  $sourcesLine = '<p style="{0}">{1}.</p>' -f $p, (ConvertTo-ReportHtml (Get-ReportSourcesText $model))
  if (-not $model.Measured) {
    return [pscustomobject]@{ Status='NotMeasured'; Subject='SmartFinOps Workplace - Daily valuation - Licences report not measured'; Attach=$true
      BodyHtml=('<p style="{0}">{1} <strong>Licences report not qualified.</strong> Every figure of run {2} is Not measured; no amount is published and no older value is reused.</p>{3}{4}' -f $p, (& $badge 'Not measured'), (ConvertTo-ReportHtml $Inputs.RunId), $sourcesLine, $link) }
  }
  $amountText = { param($Row, [string]$Status)
    $text = if ($Row) { Get-ReportEurText $Row.AnnualAmount } else { '' }
    if ($text) { $text } elseif ($Status -and $Status -cne 'Valued') { $Status } else { 'None identified' } }
  $accent = @{ 'confirmed'='#0078d4'; 'to-confirm'='#d97706'; 'conditional'='#0f766e'; 'avoidance'='#64748b' }
  $cell = { param($Card)
    $badges = (& $badge $Card.Certainty) + $(if ($Card.Tag) { ' ' + (& $badge $Card.Tag) } else { '' })
    '<td width="50%" style="width:50%;padding:6px;vertical-align:top;"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#ffffff;border:1px solid #dde7f0;border-left:4px solid {0};"><tr><td style="padding:12px 10px;"><div style="font-size:12px;line-height:17px;color:#5f6b7a;">{1}</div><div style="font-size:22px;line-height:28px;font-weight:700;color:#1f2937;margin:4px 0;">{2}</div><div>{3}</div><div style="font-size:11px;line-height:16px;color:#5f6b7a;margin-top:6px;">{4}</div></td></tr></table></td>' -f $accent[$Card.Class], (ConvertTo-ReportHtml $Card.Label), (ConvertTo-ReportHtml (& $amountText $Card.Row $Card.Status)), $badges, (ConvertTo-ReportHtml $Card.Detail) }
  $cards = $model.Cards
  $cardsHtml = '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;background:#f5f8fb;border:1px solid #dde7f0;"><tr>{0}{1}</tr><tr>{2}{3}</tr></table>' -f (& $cell $cards[0]), (& $cell $cards[1]), (& $cell $cards[2]), (& $cell $cards[3])
  $th = 'padding:7px 8px;background:#e6f4ff;border-bottom:1px solid #dde7f0;text-align:left;font-size:11px;color:#1f2937;'
  $td = 'padding:7px 8px;border-bottom:1px solid #edf2f7;font-size:12px;color:#1f2937;vertical-align:top;'
  $topRows = @($model.Top | ForEach-Object { '<tr><td style="{0}">{1}</td><td style="{0}">{2} ({3})<div style="font-size:11px;color:#5f6b7a;">{4}</div></td><td style="{0}text-align:right;white-space:nowrap;">{5}</td><td style="{0}">{6}</td></tr>' -f $td, $_.Rank, (ConvertTo-ReportHtml $_.Label), (ConvertTo-ReportHtml $_.Lever), (ConvertTo-ReportHtml $_.Action), (ConvertTo-ReportHtml (Get-ReportEurText $_.Row.AnnualAmount)), (& $badge $_.Certainty) })
  $topHtml = if ($topRows.Count) { '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;"><tr><th style="{0}">#</th><th style="{0}">Decision</th><th style="{0}text-align:right;">Annual amount</th><th style="{0}">Certainty</th></tr>{1}</table>' -f $th, ($topRows -join '') } else { '<p style="{0}">No valued decision.</p>' -f $p }
  $provisional = $model.Overall -ceq 'Provisional'
  $banner = if ($provisional) { '<p style="margin:0 0 12px;font-size:12px;line-height:18px;padding:8px 10px;background:#fff8e6;border-left:3px solid #d97706;color:#7c2d12;">Provisional figures for levers {0}: a source they use is provisional.</p>' -f (ConvertTo-ReportHtml ($model.ProvisionalLevers -join ', ')) } else { '' }
  $totalR = $model.TotalR
  $esu = $model.Esu
  $body = ('<p style="{0}">{1} Daily valuation of profile {2}, run {3}. Amounts in EUR excluding VAT, annual run-rate.</p>{4}' -f $p, (& $badge $(if ($provisional) { 'Provisional' } else { 'Final' })), (ConvertTo-ReportHtml $Tenant), (ConvertTo-ReportHtml $Inputs.RunId), $banner) +
    $cardsHtml +
    ('<p style="{0}">Realizable reduction (R) published: {1} = R confirmed + R to confirm. Conditional E3 to F3 and cost avoidance are never added to R. Windows 10 ESU: {2}.</p>' -f $muted, (ConvertTo-ReportHtml (& $amountText $totalR $(if ($totalR) { [string]$totalR.ValuationStatus } else { 'Not measured' }))), (ConvertTo-ReportHtml (& $amountText $esu $(if ($esu) { [string]$esu.ValuationStatus } else { 'Not measured' })))) +
    ('<h3 style="{0}">Top 3 decisions</h3>' -f $h) + $topHtml +
    $sourcesLine + $link
  return [pscustomobject]@{ Status=$(if ($provisional) { 'Provisional' } else { 'Final' }); Subject=$(if ($provisional) { 'SmartFinOps Workplace - Daily valuation - provisional figures' } else { 'SmartFinOps Workplace - Daily valuation' }); Attach=$true; BodyHtml=$body }
}

function Send-ReportDailyEmail {
  param(
    [Parameter(Mandatory)]$Inputs, [Parameter(Mandatory)][string]$Tenant, [Parameter(Mandatory)][string]$TenantKey,
    [Parameter(Mandatory)][string]$PagePath, [Parameter(Mandatory)][string]$StatePath, [string]$ReportUrl,
    [Parameter(Mandatory)][string]$From, [Parameter(Mandatory)][string]$To, [switch]$Force,
    [Parameter(Mandatory)][scriptblock]$Send, [datetimeoffset]$NowUtc = [datetimeoffset]::UtcNow
  )
  # Returns Sent, Skipped or Failed; a failure is a warning (the page stays written), never a run error.
  $gate = $null; $sendStarted = $false
  try {
    $gate = Enter-ReportDailyMailGate -Path $StatePath -TenantKey $TenantKey -Force:$Force -NowUtc $NowUtc
    if ($gate.Skip) { Write-ReportLog ("Daily email skipped: {0} (Europe/Paris). Use -ForceEmail to send again." -f $gate.Reason); return 'Skipped' }
    $content = Get-ReportDailyEmailContent -Inputs $Inputs -Tenant $Tenant -ReportUrl $ReportUrl
    $attachments = @()
    if ($content.Attach) {
      if (-not (Test-Path -LiteralPath $PagePath -PathType Leaf)) { throw "Report page not found for the attachment: $PagePath" }
      if (-not (Test-ReportIdentifierFree ([IO.File]::ReadAllText($PagePath)))) { throw 'The report page contains an identifier pattern (e-mail address or GUID); nothing is sent.' }
      $attachments = @($PagePath)
    }
    if (-not (Test-ReportIdentifierFree ($content.Subject + $content.BodyHtml))) { throw 'The daily email body contains an identifier pattern (e-mail address or GUID); nothing is sent.' }
    $sendStarted = $true
    & $Send ([pscustomobject]@{ From=$From; To=$To; Subject=$content.Subject; BodyHtml=$content.BodyHtml; Attachments=$attachments; Status=$content.Status })
    $sent = [ordered]@{ TenantKey=$TenantKey; LocalDate=$gate.Day; Status='Sent'; SentAtUtc=[datetimeoffset]::UtcNow.ToString('o'); Content=$content.Status }
    Write-ReportDailyMailState -Stream $gate.Stream -Bytes ([Text.UTF8Encoding]::new($false).GetBytes(($sent | ConvertTo-Json -Compress)))
    Write-ReportLog ("Daily email sent ({0}){1}." -f $content.Status, $(if ($attachments.Count) { ' with the report attached' } else { '' }))
    return 'Sent'
  }
  catch {
    if ($gate -and -not $gate.Skip -and -not $sendStarted) {
      try { Write-ReportDailyMailState -Stream $gate.Stream -Bytes $gate.PreviousBytes }
      catch { Write-ReportLog ("Daily email state could not be restored: {0}" -f $_.Exception.Message) 'WARNING' }
    }
    Write-ReportLog ("Daily email failed: {0}{1}" -f $_.Exception.Message, $(if ($sendStarted) { ' (state kept Pending: no automatic resend today; use -ForceEmail after checking delivery)' } else { '' })) 'WARNING'
    return 'Failed'
  }
  finally { if ($gate -and $gate.Stream) { $gate.Stream.Dispose() } }
}
#endregion

#region Output
function Write-ReportFile {
  param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$Content)
  $temporary = $Path + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
  try {
    [IO.File]::WriteAllText($temporary, $Content, [Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $temporary -Destination $Path -Force
  }
  finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force } }
}
#endregion

#region Main
function Send-ReportSharePointFile {
  param([Parameter(Mandatory)]$SharePoint, [Parameter(Mandatory)][string]$Path)
  # SmartM365 upload to the same DATA-LAST / DATA-ALL / LOG-ALL relative path; the helper reports failure by returning nothing.
  $parameters = $SharePoint.Parameters
  $record = $null
  try { $record = SmartM365.Core\Invoke-SmartM365SharePointCsvUpload -LocalFilePath $Path -EnsureParentFolders @parameters }
  catch { Write-ReportLog ("SharePoint upload failed for {0}: {1}" -f $Path, $_.Exception.Message) 'WARNING'; return $false }
  if ($null -eq $record -or -not ($record.PSObject.Properties['SharePointPath'] -and $record.SharePointPath)) {
    Write-ReportLog "SharePoint upload failed for ${Path}: the SmartM365 upload returned no result (see the messages above)." 'WARNING'
    return $false
  }
  Write-ReportLog "SharePoint uploaded: $($record.SharePointPath)"
  return $record
}

if ($MyInvocation.InvocationName -eq '.') { return }

$exitCode = 0
$startedAt = Get-Date
$root = $null; $archiveDir = $null; $runLogPath = ''; $sharePoint = $null
try {
  . (Join-Path $PSScriptRoot 'Config\SmartFinOps-TenantContext.ps1')
  Write-SmartFinOpsStartupBanner
  Write-ReportLog ("Starting SmartFinOps Workplace Report v{0} - profile '{1}'{2}" -f $script:ScriptVersion, $Tenant, $(if ($ValidateOnly) { ' - ValidateOnly (no file written)' } else { '' }))
  Write-ReportLog ("Tenant context {0}" -f $script:SmartFinOpsTenantContextVersion)
  Initialize-SmartFinOpsTenantContext -Tenant $Tenant -StartPath $PSScriptRoot | Out-Null
  # SmartM365 data tree of the profile: valuation outputs and the page in DATA-LAST, run history in DATA-ALL, log in LOG-ALL.
  $layout = Get-SmartFinOpsDataLayout -DataLast $SmartM365LatestCsvFolderPath -DataAll $SmartM365DataAllRootPath -LogAll $SmartM365LogAllRootPath
  $root = $layout.DataLast
  $runLogPath = Join-Path (Join-Path $layout.LogAll 'SmartFinOps-Workplace-Report') ('SmartFinOps-Workplace-Report-{0}.log' -f $startedAt.ToString('yyyy-MM-dd_HH-mm-ss', $script:Invariant))
  $tenantKey = if ($script:TenantKeyParameter) { $script:TenantKeyParameter.Trim() } else { [string](Get-SmartFinOpsConfigValue -Name 'SmartM365TenantKey') }
  Assert-SmartFinOpsResolvedPath -Name 'SmartM365TenantKey' -Value $tenantKey
  if (-not $tenantKey) { throw 'SmartM365TenantKey is empty: set it in the SmartFinOps configuration, pass -SmartM365TenantKey, or leave __USE_SMARTM365__.' }
  Write-ReportLog ("Profile '{0}' - TenantKey '{1}'{2}" -f $Tenant, $tenantKey, $(if ($script:TenantKeyParameter) { ' (parameter)' } else { '' }))
  $reportConfig = Get-SmartFinOpsConfigValue -Name 'Report'
  $minimum = Get-ReportProperty $reportConfig 'ShowbackMinimumGroupUsers'
  $minimumUsers = if ($null -eq $minimum -or [string]$minimum -eq '') { 5 } else { [int]$minimum }
  if ($minimumUsers -lt 1) { throw 'Report.ShowbackMinimumGroupUsers must be a positive integer.' }
  Write-ReportLog "Inputs and page: $root"
  Write-ReportLog "Log: $runLogPath"
  if (-not $ValidateOnly) {
    $sharePoint = Get-SmartFinOpsSharePointConfiguration -TenantKey $tenantKey
    if (-not $sharePoint.Enabled) { Write-ReportLog 'SharePoint upload disabled.'; $sharePoint = $null }
    elseif ($sharePoint.Missing.Count) { Write-ReportLog ("SharePoint upload enabled but the SmartM365 configuration is incomplete (missing: {0}); no file is uploaded." -f ($sharePoint.Missing -join ', ')) 'WARNING'; $sharePoint = $null }
    else {
      Import-Module (Join-Path (Split-Path -Path $PSScriptRoot -Parent) 'SmartM365\Modules\SmartM365.Core\SmartM365.Core.psd1') -MinimumVersion '1.0.90' -ErrorAction Stop | Out-Null
      Write-ReportLog ("SharePoint upload enabled (SmartM365 context of profile '{0}'): {1}{2}, library '{3}', data root '{4}' (DATA-LAST, DATA-ALL, LOG-ALL)." -f $sharePoint.ProfileKey, $sharePoint.Parameters.SiteHostname, $sharePoint.Parameters.SitePath, $sharePoint.Parameters.LibraryDisplayName, (SmartM365.Core\ConvertTo-SmartM365SharePointDataRootPath -TargetFolderPath $sharePoint.Parameters.TargetFolderPath))
    }
  }

  $inputs = Read-ReportInputs -Root $root -ProfileKey $Tenant -TenantKey $tenantKey
  foreach ($check in $inputs.Checks) { Write-ReportLog ('{0}: {1} - {2}' -f $check.Rule, $check.Status, $check.Detail) $(if ($check.Status -eq 'Pass') { 'INFO' } else { 'WARNING' }) }
  if ($inputs.Available) {
    $html = New-ReportPageContent -Inputs $inputs -Tenant $Tenant -MinimumGroupUsers $minimumUsers
    Write-ReportLog ("Report built for run {0} (Licences measured: {1}; showback minimum group: {2} users)." -f $inputs.RunId, [bool](Get-ReportProperty $inputs.Receipt 'LicencesMeasured'), $minimumUsers)
    # The page is complete, but its figures are not final: the run completes with warnings.
    if (-not [bool](Get-ReportProperty $inputs.Receipt 'LicencesMeasured')) { Write-ReportLog ("Licences report not measured in run {0}: every figure is shown as Not measured." -f $inputs.RunId) 'WARNING' }
    if ([bool](Get-ReportProperty $inputs.Receipt 'Provisional')) { Write-ReportLog ("Provisional figures in run {0} for levers {1}." -f $inputs.RunId, (@(Get-ReportProperty $inputs.Receipt 'ProvisionalLevers' | Where-Object { $_ }) -join ', ')) 'WARNING' }
  }
  else {
    $html = New-ReportUnavailablePage -Inputs $inputs -Tenant $Tenant
    Write-ReportLog "Report not available: $($inputs.Reason)" 'ERROR'
    $exitCode = 1
  }
  if ($ValidateOnly) { Write-ReportLog 'ValidateOnly: no file written.' }
  else {
    Write-ReportFile -Path (Join-Path $root $script:ReportName) -Content $html
    if ($inputs.RunId -cmatch '^\d{8}T\d{9}Z-[0-9a-f]{8}$') {
      $candidate = Join-Path $layout.HistoryRoot $inputs.RunId
      if (Test-Path -LiteralPath $candidate -PathType Container) { $archiveDir = $candidate; Write-ReportFile -Path (Join-Path $archiveDir $script:ReportName) -Content $html }
    }
    Write-ReportLog ("Report written: {0}{1}" -f (Join-Path $root $script:ReportName), $(if ($archiveDir) { ' (copy in the run history)' } else { '' }))
    $pageRecord = $null
    if ($sharePoint) {
      if ($archiveDir) { [void](Send-ReportSharePointFile -SharePoint $sharePoint -Path (Join-Path $archiveDir $script:ReportName)) }
      $pageRecord = Send-ReportSharePointFile -SharePoint $sharePoint -Path (Join-Path $root $script:ReportName)
    }
    # Daily email, after the page is written (and uploaded): a mail problem is a warning, the page stays.
    $emailSetting = Get-ReportProperty $reportConfig 'EnableDailyEmail'
    $emailEnabled = $true
    if ($null -ne $emailSetting -and [string]$emailSetting -ne '' -and -not [bool]::TryParse([string]$emailSetting, [ref]$emailEnabled)) { throw "Report.EnableDailyEmail must be true or false (got '$emailSetting')." }
    if (-not $emailEnabled) { Write-ReportLog 'Daily email disabled (Report.EnableDailyEmail).' }
    else {
      try {
        $effective = Initialize-SmartFinOpsSmartM365Context -ProfileKey $Tenant -TenantKey $tenantKey
        Import-Module (Join-Path (Split-Path -Path $PSScriptRoot -Parent) 'SmartM365\Modules\SmartM365.Core\SmartM365.Core.psd1') -MinimumVersion '1.0.91' -ErrorAction Stop | Out-Null
        $recipients = @{}
        foreach ($pair in @(@('DailyEmailTo', 'To'), @('DailyEmailFrom', 'From'))) {
          $value = [string](Get-ReportProperty $reportConfig $pair[0])
          if ([string]::IsNullOrWhiteSpace($value) -or $value.Trim() -in @('__USE_SMARTM365__', 'USE_SMARTM365')) {
            $property = $effective.PSObject.Properties[$pair[1]]
            $value = if ($null -ne $property) { [string]$property.Value } else { '' }
          }
          if ($value.Trim() -in @('__USE_GLOBAL__', 'USE_GLOBAL') -or [string]::IsNullOrWhiteSpace($value)) { throw ("no {0}: set Report.{1} or the SmartM365 {0} of profile '{2}'." -f $pair[1], $pair[0], $Tenant) }
          $recipients[$pair[1]] = $value.Trim()
        }
        # The SmartM365 mail copy is written next to the report log; the footer names this script.
        $mailLogFolder = Split-Path -Path $runLogPath -Parent
        New-Item -ItemType Directory -Path $mailLogFolder -Force | Out-Null
        foreach ($session in @(@('LogPath', $mailLogFolder), @('SmartM365ScriptFileName', 'SmartFinOps-Workplace-Report.ps1'), @('SmartM365ScriptVersion', $script:ScriptVersion))) { Set-Variable -Name $session[0] -Scope Global -Value $session[1] }
        $transport = { param($Message) [void](SmartM365.Core\Send-SmartM365Mail -From $Message.From -To $Message.To -Subject $Message.Subject -BodyHtml $Message.BodyHtml -MailPurpose Report -Attachments $Message.Attachments -AllowAttachments -SuppressAttachmentLinks) }
        $reportUrl = if ($pageRecord -and $pageRecord.PSObject.Properties['WebUrl']) { [string]$pageRecord.WebUrl } else { '' }
        [void](Send-ReportDailyEmail -Inputs $inputs -Tenant $Tenant -TenantKey $tenantKey -PagePath (Join-Path $root $script:ReportName) -StatePath (Join-Path $root $script:DailyEmailStateName) -ReportUrl $reportUrl -From $recipients.From -To $recipients.To -Force:$ForceEmail -Send $transport)
      }
      catch { Write-ReportLog ("Daily email skipped: {0}" -f $_.Exception.Message) 'WARNING' }
    }
  }
}
catch {
  $exitCode = 1
  Write-ReportLog ("Report failed: {0}" -f $_.Exception.Message) 'ERROR'
  Write-ReportLog ("Stack: {0}" -f ($_.ScriptStackTrace -replace '\r?\n', ' | ')) 'ERROR'
}
finally {
  if (-not $ValidateOnly -and $runLogPath) {
    try {
      New-Item -ItemType Directory -Path (Split-Path $runLogPath -Parent) -Force | Out-Null
      Write-ReportFile -Path $runLogPath -Content (($script:LogLines -join [Environment]::NewLine) + [Environment]::NewLine)
      # The SharePoint copy is the log without the completion banner, which is appended locally right after.
      if ($sharePoint -and -not (Send-ReportSharePointFile -SharePoint $sharePoint -Path $runLogPath)) {
        Write-ReportFile -Path $runLogPath -Content (($script:LogLines -join [Environment]::NewLine) + [Environment]::NewLine)
      }
    }
    catch { Write-Host "Report log could not be written: $($_.Exception.Message)" -ForegroundColor Red; $runLogPath = '' }
  }
  else { $runLogPath = '' }
  # Completion banner on success and on failure; "Report not available" (exit 1) is a failed run.
  # Exit code contract (orchestrator): 0 success, 3 completed with warnings, 1 failure.
  $warnings = @($script:LogLines | Where-Object { $_ -match '^\S+ \S+ \[WARNING\]' }).Count
  $errors = @($script:LogLines | Where-Object { $_ -match '^\S+ \S+ \[ERROR\]' }).Count
  if ($exitCode -eq 0 -and $errors -gt 0) { $exitCode = 1 }
  elseif ($exitCode -eq 0 -and $warnings -gt 0) { $exitCode = 3 }
  $bannerStatus = switch ($exitCode) { 0 { 'Success' } 3 { 'CompletedWithWarnings' } default { 'Failed' } }
  if (Get-Command -Name Write-SmartFinOpsCompletionBanner -ErrorAction SilentlyContinue) {
    Write-SmartFinOpsCompletionBanner -Status $bannerStatus -ScriptName 'SmartFinOps-Workplace-Report.ps1' -StartedAt $startedAt -WarningCount $warnings -ErrorCount $errors -GeneratedCsvFiles 0 -LogPath $runLogPath
  }
  else { Write-Host ("SmartFinOps-Workplace-Report.ps1: {0} (tenant context not loaded)" -f $(switch ($exitCode) { 0 { 'SUCCESS' } 3 { 'COMPLETED WITH WARNINGS' } default { 'FAILED' } })) }
}
exit $exitCode
#endregion

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBDfWhv7eraKTUo
# ptexru7AhSffQjEFnT7NRUduhRW1VqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIB+Fxl/NcVIZVCPIxDHFUtF8yM6eqGeIOC2fpnoDvuMRMA0GCSqG
# SIb3DQEBAQUABIIBgFg9sjOPeZjPGZ5z2rGHKvtPhZx8pBlWDQP/pOwY7l9xh4yV
# 7V8viUjExNxlVHhDVhRDJ4h9Wo3OOSQvrfoKhYCcKtVPg2GVAsddOyNcC5FfkmlZ
# gA3vja8g/uFCM71F0Ts3TA9Wzg/jaaQ/5xRLA6aS9Qyr9t0CtKXhKtmiXCn8eV/Y
# Kx4oQEwdSUaNW9Yihuj26Aosc/lyvBl2S+yyiY9TiRZszxjKZHA6+xU5iQKY8StA
# /pu+UMe2fMdKOiFHtY4folNOwy+NNWiGEP1670TTER7pnll8AohEc7oo0fS8IIVi
# dvp9vIwhv+5piPpl1mXhC8i2K8tF9Xe61uyDQRWymBfTK36lvC3Euiw5LUH0oDZH
# ygmVmqI7zTNTrquxaVAXSh+E7Xgx/BQOH89GIUmU/5rYBYUwFIS/KsW67tY+S9cP
# dcSzKUGmtbSnSUy+uiVilSpK4byisBSwWkTv4oMpi83YuUvRijEMP5z2Qn6w5uEn
# dBy1HM4v9OnyigHDlKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTkxNDA5
# NThaMC8GCSqGSIb3DQEJBDEiBCB0X3bYPKqeDXomESJTEISbbuwns6pICoBaHayL
# B59gHjANBgkqhkiG9w0BAQEFAASCAgAqdfGKtPhp/CrKcpX1F1pLOp8p7DIO1NcC
# ei+G7lJwGIjO3hk6zIFASjnI0aSFK/h1pKkeGNUywIIPQwSpS25jJ+GYIYWKPamY
# DrAC8hIOMZBVW7/v7b/2PvsLpwC4W9Znnj9oc3QAHCo9FU1QAbTqyKxIQI8fP8qD
# 3mZQDAfvQmSDIYBZ+QttbdLyoToTf5T/QOlw3eFNhQAUTNbnbqMmvxXseTgGMpLy
# 6jlOfoDB47l+yJtd8x9JEaOloJCx3+ruqk2orqK2qAXqZuRRKsajMSdzGJa1waBc
# 6YHhBsJE91hYOCF+KXN6IfNTuViCV6B5h/49QaPVU6qiF39OcVdfD+wBgGvMDEIL
# be1cPavd7T00bgBnQpJL9OqJfTtsK34nZ3xlswF61j08qTEZ3GFRQBIpnLQ5Q9RL
# 4xtZtC3u7TMlTdKhsE6ldGy6axsjZW14ow4dXI4VNy1pQdP9NWDVVGb2D/AYkp7S
# xy7+y3CGN1cyUa3dsg4/4QjWnZ4YfQEi6T3TIA9Ub/TbaTl2j7HaznodWpOznPua
# FsWA5MkugOJ3009lDEbKgfrcz5aN40rWuEXtkb9Ik0MD9gXhKmVeSg5k7jRONYb9
# hlpuh+enoR2cbPbJ7CSiWNHZE+I5Q/Dh6kHVS4JqSG+oRooyQd8RhfUq1NkqCyx2
# TuJKB9/Qjw==
# SIG # End signature block
