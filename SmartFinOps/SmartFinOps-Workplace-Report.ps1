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
.VERSION
2.2.2
.NOTES
  Output: Data\<TenantKey>\SmartFinOps_Workplace-Report.html and its log; copy of both in
  Archive\<FinOpsRunId>\ when that run is archived. Aggregates only: no user identifier, no external
  resource, no script, every value HTML-encoded. Same inputs give the same HTML except the generation time.
#>
[CmdletBinding()]
param(
  [string]$Tenant = 'test',
  [string]$InputRoot,
  [switch]$ValidateOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:ScriptVersion = '2.2.2'
$script:LogLines = [System.Collections.Generic.List[string]]::new()
$script:Invariant = [Globalization.CultureInfo]::InvariantCulture
$script:ReceiptName = 'SmartFinOps_Workplace-Valuation.current.json.txt'
$script:ReportName = 'SmartFinOps_Workplace-Report.html'
$script:LogName = 'SmartFinOps-Workplace-Report.log'

#region Input contract (engine 2.1.1 or later 2.1.x)
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
  TOTAL_R_DISABLED_AND_SHARED = 'of which disabled accounts and eligible shared mailboxes'
  TOTAL_R_ENABLED_INACTIVE = 'of which enabled accounts without M365 activity in 90 days'
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
  param([Parameter(Mandatory)][string]$Root, [Parameter(Mandatory)][string]$TenantKey)
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
    if ([string](Get-ReportProperty $receipt 'Owner') -cne 'SmartFinOps-Receipt' -or [string](Get-ReportProperty $receipt 'ContractVersion') -cne '1.0') { throw 'The valuation receipt has an unsupported owner or contract version.' }
    $version = [string](Get-ReportProperty $receipt 'ScriptVersion')
    if ($version -notmatch '^2\.1\.(\d+)$' -or [int]$Matches[1] -lt 1) { throw "The valuation engine version '$version' is not supported (2.1.1 or later 2.1.x required)." }
    if ($result.RunId -notmatch '^\d{8}T\d{9}Z-[0-9a-f]{8}$') { throw 'The valuation receipt RunId is invalid.' }
    & $pass $rule "Owner SmartFinOps-Receipt 1.0, engine $version"

    $rule = 'Q2 Status'
    $status = [string](Get-ReportProperty $receipt 'Status')
    if ($status -cne 'Completed') {
      $runError = [string](Get-ReportProperty $receipt 'Error')
      throw ("The valuation run is {0}{1}." -f $(if ($status) { $status } else { 'in an unknown state' }), $(if ($runError) { ": $runError" } else { '' }))
    }
    & $pass $rule 'Completed'

    $rule = 'Q3 Tenant'
    if ([string](Get-ReportProperty $receipt 'TenantKey') -cne $TenantKey) { throw 'The valuation receipt belongs to another tenant.' }
    & $pass $rule $TenantKey

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
    & $pass $rule $(if ($measured) { 'G1-G6 Pass' } else { 'Not applicable (Licences report not measured)' })
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
$script:Css = @'
    body { margin: 0; font-family: Segoe UI, Arial, sans-serif; background: #f5f8fb; color: #1f2937; }
    header { background: #ffffff; border-bottom: 1px solid #dde7f0; padding: 24px 32px; }
    main { padding: 24px 32px 40px; }
    h1 { margin: 0 0 8px; font-size: 26px; font-weight: 650; }
    h2 { margin: 0 0 12px; font-size: 18px; }
    h3 { margin: 18px 0 8px; font-size: 14px; }
    .subtitle { color: #5f6b7a; margin: 0; }
    .grid { display: grid; grid-template-columns: repeat(4, minmax(160px, 1fr)); gap: 12px; margin-top: 18px; }
    .metric { background: #fff; border: 1px solid #dde7f0; border-radius: 8px; padding: 14px; }
    .metric .label { color: #5f6b7a; font-size: 12px; }
    .metric .value { font-size: 24px; font-weight: 650; margin-top: 4px; }
    .metric .detail { color: #5f6b7a; font-size: 12px; margin-top: 6px; }
    section { background: #fff; border: 1px solid #dde7f0; border-radius: 8px; padding: 16px; margin-top: 16px; }
    table { border-collapse: collapse; width: 100%; font-size: 12px; }
    th { text-align: left; background: #eef5fb; color: #1f2937; border-bottom: 1px solid #dde7f0; padding: 8px; }
    td { border-bottom: 1px solid #edf2f7; padding: 7px 8px; vertical-align: top; }
    td.num, th.num { text-align: right; white-space: nowrap; }
    .note, .empty, .none { color: #5f6b7a; font-size: 12px; }
    .pill { display: inline-block; border-radius: 999px; background: #e6f4ff; color: #005a9e; padding: 3px 9px; font-size: 12px; font-weight: 600; }
    .banner { border-radius: 8px; padding: 10px 14px; margin-top: 12px; font-size: 13px; border: 1px solid #f0c36d; background: #fff8e6; }
    .banner.stop { border-color: #e3a4a4; background: #fdeeee; }
    .badge { display: inline-block; border-radius: 4px; padding: 1px 6px; font-size: 11px; font-weight: 600; border: 1px solid; white-space: nowrap; }
    .badge.ok { color: #1d6b3a; border-color: #9fd2b1; background: #eaf7ef; }
    .badge.prov { color: #8a5a00; border-color: #f0c36d; background: #fff8e6; }
    .badge.nv { color: #4b5563; border-color: #c7ccd3; background: #f3f4f6; }
    .badge.nm { color: #4b5563; border-color: #9ca3af; background: repeating-linear-gradient(45deg, #f3f4f6, #f3f4f6 4px, #e5e7eb 4px, #e5e7eb 8px); }
    .badge.rev { color: #1e4f91; border-color: #9cbbe6; background: #edf4fd; }
    .badge.blk { color: #9b1c1c; border-color: #e3a4a4; background: #fdeeee; }
    .badge.cap { color: #5b2a86; border-color: #c3a6de; background: #f5eefb; }
    .badge.reu, .badge.ref, .badge.sig { color: #374151; border-color: #d1d5db; background: #fff; }
    .bars text { font-size: 11px; fill: #1f2937; }
    .bars rect { fill: #3b82c4; }
    details summary { cursor: pointer; font-weight: 600; margin: 8px 0; }
    @media print { body { background: #fff; } section { break-inside: avoid; } }
'@

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
  $entry = $script:Levers[$Lever]
  if ($null -eq $entry) { return (ConvertTo-ReportHtml $Lever) }
  return ('{0} {1}' -f (ConvertTo-ReportHtml $Lever), (ConvertTo-ReportHtml $entry[0]))
}

function Get-ReportLeverNumber { param([string]$Lever) $n = 0; if ([int]::TryParse(([string]$Lever).TrimStart('L'), [ref]$n)) { return $n } return 99 }

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
<header>
$Header
</header>
<main>
$Body
</main>
</body>
</html>
"@
}

function Get-ReportGeneratedLine {
  param([Parameter(Mandatory)][string]$Tenant, [string]$RunId)
  return ('<p class="subtitle">Tenant: {0} | FinOps run: {1} | Generated: <span id="generated">{2}</span> UTC | Report {3}</p>' -f (ConvertTo-ReportHtml $Tenant), (ConvertTo-ReportHtml $(if ($RunId) { $RunId } else { 'unknown' })), [datetime]::UtcNow.ToString('yyyy-MM-dd HH:mm:ss', $script:Invariant), $script:ScriptVersion)
}
#endregion

#region Not available page
function New-ReportUnavailablePage {
  param([Parameter(Mandatory)]$Inputs, [Parameter(Mandatory)][string]$Tenant)
  $header = "  <div class=""pill"">SmartFinOps Workplace</div>`n  <h1>Report not available</h1>`n  $(Get-ReportGeneratedLine -Tenant $Tenant -RunId $Inputs.RunId)"
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
  $amountOf = { param([string]$Key) $row = & $get $Key; if ($row) { Get-ReportAmountCell $row.AnnualAmount $row.ValuationStatus } else { Get-ReportBadge 'Not measured' } }
  $measured = [bool](Get-ReportProperty $receipt 'LicencesMeasured')
  $skuNames = @{}; foreach ($sku in $skus) { $skuNames[[string]$sku.SkuId] = [string]$sku.SkuPartNumber }
  $first = @($summary | Where-Object { $_.SnapshotGeneratedAtUtc } | Select-Object -First 1)
  $snapshotDate = if ($first.Count) { [string]$first[0].SnapshotGeneratedAtUtc } else { '' }
  $pricingRun = if ($first.Count) { [string]$first[0].PricingReceiptRunId } else { '' }

  # Header and status banners.
  $header = [System.Text.StringBuilder]::new()
  [void]$header.AppendLine('  <div class="pill">SmartFinOps Workplace</div>')
  [void]$header.AppendLine('  <h1>Workplace FinOps report</h1>')
  [void]$header.AppendLine('  ' + (Get-ReportGeneratedLine -Tenant $Tenant -RunId $Inputs.RunId))
  [void]$header.AppendLine(('  <p class="subtitle">Licences snapshot: {0} | Pricing run: {1} | Amounts: EUR excluding VAT, annual run-rate</p>' -f (ConvertTo-ReportHtml $(if ($snapshotDate) { $snapshotDate } else { 'not measured' })), (ConvertTo-ReportHtml $(if ($pricingRun) { $pricingRun } else { 'not qualified' }))))
  if (-not $measured) { [void]$header.AppendLine('  <div class="banner stop">Licences report not qualified: every figure is <strong>Not measured</strong>. See the technical appendix for the source qualification.</div>') }
  $provisionalLevers = @(Get-ReportProperty $receipt 'ProvisionalLevers' | Where-Object { $_ })
  if ([bool](Get-ReportProperty $receipt 'Provisional')) { [void]$header.AppendLine(('  <div class="banner">Provisional figures for levers {0}: a source they use is provisional (see Assumptions).</div>' -f (ConvertTo-ReportHtml ($provisionalLevers -join ', ')))) }
  $missingDomains = [string](Get-ReportProperty $receipt 'MissingDomains')
  if ($missingDomains) { [void]$header.AppendLine(('  <div class="banner">AD domains not collected: {0}. Their users are Not measured in the showback.</div>' -f (ConvertTo-ReportHtml $missingDomains))) }

  # Executive cards.
  $card = {
    param([string]$Key, [string]$Detail)
    $row = & $get $Key
    $meta = if ($row -and ($row.Units -or ([string]$row.Users -and [string]$row.Users -ne '0'))) { 'Units: {0} | Users: {1}' -f $(if ($row.Units) { Format-ReportCount $row.Units } else { '-' }), (Format-ReportCount $row.Users) } else { '' }
    '<div class="metric"><div class="label">{0}</div><div class="value">{1}</div><div class="detail">{2}</div><div class="detail">{3}</div></div>' -f (ConvertTo-ReportHtml $script:Totals[$Key]), (& $amountOf $Key), (ConvertTo-ReportHtml $meta), (ConvertTo-ReportHtml $Detail)
  }
  [void]$header.AppendLine('  <div class="grid">')
  [void]$header.AppendLine('    ' + (& $card 'TOTAL_R' 'At contract or renewal; excludes E3 to F3'))
  [void]$header.AppendLine('    ' + (& $card 'TOTAL_R_ENABLED_INACTIVE' 'Included in R; to confirm with the business'))
  [void]$header.AppendLine('    ' + (& $card 'TOTAL_RC_E3_TO_F3' 'If Frontline eligibility is confirmed'))
  [void]$header.AppendLine('    ' + (& $card 'TOTAL_E_REUSE' 'Purchase avoided; never added to R'))
  [void]$header.AppendLine('  </div>')

  $body = [System.Text.StringBuilder]::new()
  $detailRows = @($summary | Where-Object { ([string]$_.KpiKey).StartsWith('DETAIL|', [StringComparison]::Ordinal) })
  $licenseLevers = @('L1','L2','L3','L4','L5','L6','L7')

  #region 1. Executive summary
  [void]$body.AppendLine('<section id="executive-summary"><h2>1. Executive summary</h2>')
  $r = & $get 'TOTAL_R'; $inactive = & $get 'TOTAL_R_ENABLED_INACTIVE'; $rc = & $get 'TOTAL_RC_E3_TO_F3'; $e = & $get 'TOTAL_E_REUSE'
  $sentence = { param($Row, [string]$Valued, [string]$Other) if ($Row -and (Get-ReportEurText $Row.AnnualAmount)) { $Valued -f (Get-ReportEurText $Row.AnnualAmount) } else { $Other -f $(if ($Row -and $Row.ValuationStatus -and $Row.ValuationStatus -cne 'Valued') { $Row.ValuationStatus } else { 'none identified' }) } }
  [void]$body.AppendLine('<ul>')
  [void]$body.AppendLine(('<li>{0}{1}</li>' -f (ConvertTo-ReportHtml (& $sentence $r 'Realizable license reduction: {0} per year at contract or renewal' 'Realizable license reduction: {0}')), (ConvertTo-ReportHtml $(if ($inactive -and (Get-ReportEurText $inactive.AnnualAmount)) { ", of which $(Get-ReportEurText $inactive.AnnualAmount) from enabled accounts without activity in 90 days." } else { '.' }))))
  [void]$body.AppendLine(('<li>{0}</li>' -f (ConvertTo-ReportHtml (& $sentence $rc 'Conditional E3 to F3 downgrades: {0} per year if Frontline eligibility is confirmed; not included in the total.' 'Conditional E3 to F3 downgrades: {0}.'))))
  [void]$body.AppendLine(('<li>{0}</li>' -f (ConvertTo-ReportHtml (& $sentence $e 'Cost avoidance: {0} per year by reusing freed units for users waiting for capacity; never added to the reduction.' 'Cost avoidance from reuse: {0}.'))))
  [void]$body.AppendLine('</ul>')
  [void]$body.AppendLine('<h3>Decision status (license units, levers L1 to L7)</h3>')
  $strip = foreach ($status in @('Provisional','Review','Blocked','Capped','Not valued','Not measured')) {
    $units = [long]0; $found = $false
    foreach ($row in @($detailRows | Where-Object { $_.ValuationStatus -ceq $status -and $_.Lever -in $licenseLevers })) { $n = ConvertTo-ReportLong $row.Units; if ($null -ne $n) { $units += $n; $found = $true } }
    ,@((Get-ReportBadge $status), $(if ($found) { $units.ToString('#,##0', $script:Invariant) } else { '<span class="none">None</span>' }))
  }
  Add-ReportTable $body @('Status','License units') @($strip) @(1)
  [void]$body.AppendLine('</section>')
  #endregion

  #region 2. Financial potential
  [void]$body.AppendLine('<section id="financial-potential"><h2>2. Financial potential</h2>')
  $kpiRows = foreach ($key in $script:Totals.Keys) {
    $row = & $get $key
    if ($null -eq $row) { ,@((ConvertTo-ReportHtml $script:Totals[$key]), '', '', '', '', (Get-ReportBadge 'Not measured'), ''); continue }
    ,@((ConvertTo-ReportHtml $script:Totals[$key]), (ConvertTo-ReportHtml $row.ValueClass), (ConvertTo-ReportHtml ($row.Lever -replace ';', ', ')), (Format-ReportCount $row.Users), (Format-ReportCount $row.Units), (Get-ReportAmountCell $row.AnnualAmount $row.ValuationStatus), (ConvertTo-ReportHtml (ConvertTo-ReportSafeText $row.Note)))
  }
  Add-ReportTable $body @('Indicator','Class','Levers','Users','Units','Annual amount','Note') @($kpiRows) @(3,4,5)
  [void]$body.AppendLine('<p class="note">The "of which" lines are part of the R total and are never added to it. E3 to F3 is conditional and cost avoidance is not a reduction.</p>')
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
  [void]$body.AppendLine('</section>')
  #endregion

  #region 3. Priority decisions
  [void]$body.AppendLine('<section id="priority-decisions"><h2>3. Priority decisions</h2>')
  $valuedDetail = @(Get-ReportSorted @($detailRows | Where-Object { $_.ValuationStatus -cin @('Valued','Provisional') -and $_.AnnualAmount }) { param($x) @((ConvertTo-ReportDecimal $x.AnnualAmount), (Get-ReportLeverNumber $x.Lever)) } -Descending)
  Add-ReportBars $body 'Annual amount by lever' @($valuedDetail | ForEach-Object { [pscustomobject]@{ Label=('{0} {1}' -f $_.Lever, $script:Levers[[string]$_.Lever][0]); Value=(ConvertTo-ReportDecimal $_.AnnualAmount); Text=(Get-ReportEurText $_.AnnualAmount) } })
  $rank = 0
  Add-ReportTable $body @('#','Lever','Class','Action','Units','Annual amount') @($valuedDetail | ForEach-Object {
    $rank++
    ,@([string]$rank, (Get-ReportLeverLabel $_.Lever), (ConvertTo-ReportHtml $_.ValueClass), (ConvertTo-ReportHtml $script:Levers[[string]$_.Lever][1]), (Format-ReportCount $_.Units), (Get-ReportAmountCell $_.AnnualAmount $_.ValuationStatus))
  }) @(4,5) 'No valued decision.'
  [void]$body.AppendLine('<p class="note">Rc (E3 to F3) is conditional and E is cost avoidance: neither is part of the R total.</p>')
  [void]$body.AppendLine('<h3>Top SKUs</h3>')
  Add-ReportTable $body @('SKU','Realizable annual','Realizable from','Contract status','Reduction policy') @($skuSorted | Where-Object { $null -ne (ConvertTo-ReportDecimal $_.RealizableAnnualAmount) -and (ConvertTo-ReportDecimal $_.RealizableAnnualAmount) -gt 0 } | Select-Object -First 10 | ForEach-Object {
    ,@((ConvertTo-ReportHtml $_.SkuPartNumber), (Get-ReportAmountCell $_.RealizableAnnualAmount $_.ValuationStatus), (ConvertTo-ReportHtml $_.RealizableFrom), (ConvertTo-ReportHtml $_.ContractStatus), (ConvertTo-ReportHtml $_.ReductionPolicy))
  }) @(1) 'No SKU with a valued reduction.'
  [void]$body.AppendLine('</section>')
  #endregion

  #region 4. Trade-offs
  [void]$body.AppendLine('<section id="trade-offs"><h2>4. Cost, quality and risk trade-offs</h2>')
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
  [void]$body.AppendLine('</section>')
  #endregion

  #region 5. Capacity, renewals and forecast
  [void]$body.AppendLine('<section id="capacity-forecast"><h2>5. Capacity, renewals and forecast</h2>')
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
  [void]$body.AppendLine('</section>')
  #endregion

  #region 6. Realization
  [void]$body.AppendLine('<section id="realization"><h2>6. Realization</h2>')
  $measuredRows = @($realization | Where-Object { $_.Status -cin @('Realized','No reduction observed') })
  if ($measuredRows.Count) {
    Add-ReportTable $body @('SKU','Baseline run','Enabled before','Enabled after','Decrease','Removed candidate licenses','Baseline capacity','Realized units','Not attributed','Realized annual','Status') @(Get-ReportSorted $measuredRows { param($x) @([string]$x.SkuPartNumber) } | ForEach-Object {
      ,@((ConvertTo-ReportHtml $_.SkuPartNumber), (ConvertTo-ReportHtml $_.BaselineFinOpsRunId), (Format-ReportCount $_.EnabledUnitsBefore), (Format-ReportCount $_.EnabledUnitsAfter), (Format-ReportCount $_.Decrease), (Format-ReportCount $_.RemovedCandidateLicenses), (Format-ReportCount $_.BaselineCapacityUnits), (Format-ReportCount $_.RealizedUnits), (Format-ReportCount $_.NotAttributedUnits), (Format-ReportEur $_.RealizedAnnualAmount), (Get-ReportBadge $_.Status))
    }) @(2,3,4,5,6,7,8,9)
  }
  else { [void]$body.AppendLine('<p class="empty">No qualified baseline run yet: realized savings are measured from the next qualified runs.</p>') }
  $byStatus = @($realization | Group-Object -Property Status -CaseSensitive | ForEach-Object { [pscustomobject]@{ Status=$_.Name; Count=$_.Count } })
  Add-ReportTable $body @('Realization status','SKUs') @(Get-ReportSorted $byStatus { param($x) @([string]$x.Status) } | ForEach-Object { ,@((ConvertTo-ReportHtml (ConvertTo-ReportSafeText $_.Status)), $_.Count.ToString('#,##0', $script:Invariant)) }) @(1) 'No SKU measured.'
  [void]$body.AppendLine('</section>')
  #endregion

  #region 7. Showback
  [void]$body.AppendLine('<section id="showback"><h2>7. Showback</h2>')
  [void]$body.AppendLine(('<p class="note">Values with fewer than {0} users are combined into "Other (groups under {0} users)" so that no group can be deduced; the detail stays in the FinOps CSV files. Users are distinct users with at least one decision row.</p>' -f $MinimumGroupUsers))
  $columns = @(@('TOTAL_R','Reduction (R)'), @('TOTAL_R_ENABLED_INACTIVE','of which inactive'), @('TOTAL_RC_E3_TO_F3','E3 to F3 (Rc)'))
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
    [void]$body.AppendLine(('<details{0}><summary>By {1}</summary>' -f $(if ($dimension -ne 'Department') { ' open' } else { '' }), (ConvertTo-ReportHtml $label)))
    Add-ReportTable $body (@($label, 'Users') + @($columns | ForEach-Object { $_[1] }) + $countStatuses) @($sortedGroups | ForEach-Object {
      $group = $_
      $cells = @((ConvertTo-ReportHtml (ConvertTo-ReportSafeText $group.Value)), $(if ($null -ne $group.Users) { ([long]$group.Users).ToString('#,##0', $script:Invariant) } else { '' }))
      foreach ($column in $columns) { $cells += $(if ($group.Amounts.ContainsKey($column[0])) { Get-ReportAmountCell $group.Amounts[$column[0]].ToString('0.00', $script:Invariant) $group.Statuses[$column[0]] } else { '' }) }
      foreach ($status in $countStatuses) { $cells += $(if ($group.Counts.ContainsKey($status)) { ([long]$group.Counts[$status]).ToString('#,##0', $script:Invariant) } else { '' }) }
      ,$cells
    }) @(1,2,3,4,5,6,7,8) 'Not measured for this dimension.'
    [void]$body.AppendLine('</details>')
  }
  [void]$body.AppendLine('<p class="note">SKU level: unassigned units and reuse, not allocated to users. Not attributed: no AD or Entra match. Not measured: AD domain not collected or source not qualified.</p>')
  [void]$body.AppendLine('</section>')
  #endregion

  #region 8. Assumptions
  $settings = Get-ReportProperty $receipt 'Settings'
  [void]$body.AppendLine('<section id="assumptions"><h2>8. Assumptions</h2><ul>')
  [void]$body.AppendLine('<li>Amounts are an annual run-rate in EUR excluding VAT: monthly unit price x 12, from the date shown in Realizable from.</li>')
  [void]$body.AppendLine('<li>R is realizable at contract or renewal and capped by the contract minimum; Rc (E3 to F3) needs a Frontline eligibility check; E is a purchase avoided and is never added to R; signals are never valued.</li>')
  [void]$body.AppendLine('<li>Not measured means a source or a guard was not qualified: no amount, never zero. Not valued means no single effective EUR price.</li>')
  [void]$body.AppendLine(('<li>Price date (snapshot): {0}. Freshness: daily sources within {1} h, weekly within {2} days. Contracts file SHA256: {3}.</li>' -f (ConvertTo-ReportHtml (ConvertTo-ReportJsonText (Get-ReportProperty $settings 'AsOfDate') 'yyyy-MM-dd')), (ConvertTo-ReportHtml (Get-ReportProperty $settings 'FreshnessDailyHours')), (ConvertTo-ReportHtml (Get-ReportProperty $settings 'FreshnessWeeklyDays')), (ConvertTo-ReportHtml $(if ([string](Get-ReportProperty $settings 'ContractsFileSHA256')) { ([string](Get-ReportProperty $settings 'ContractsFileSHA256')).Substring(0, 12) } else { 'none' }))))
  $priced = @($skus | Where-Object { $_.UnitPriceMonthly })
  [void]$body.AppendLine(('<li>SKU prices: {0} catalogue estimate(s), {1} contract price(s), {2} SKU(s) without price.</li>' -f @($priced | Where-Object PriceIsEstimated -eq 'True').Count, @($priced | Where-Object PriceIsEstimated -eq 'False').Count, @($skus | Where-Object { -not $_.UnitPriceMonthly -and $_.Excluded -ne 'True' }).Count))
  foreach ($source in @(Get-ReportSorted @($sources | Where-Object { $_.ProvisionalLevers }) { param($x) @([string]$x.SourceFile) })) {
    [void]$body.AppendLine(('<li>Provisional source {0} affects {1}: {2}</li>' -f (ConvertTo-ReportHtml $source.SourceFile), (ConvertTo-ReportHtml ($source.ProvisionalLevers -replace ';', ', ')), (ConvertTo-ReportHtml (ConvertTo-ReportSafeText $(if ($source.Reason) { $source.Reason } else { 'no reason given' })))))
  }
  if ($missingDomains) { [void]$body.AppendLine(('<li>AD domains not collected: {0}.</li>' -f (ConvertTo-ReportHtml $missingDomains))) }
  [void]$body.AppendLine('</ul></section>')
  #endregion

  #region 9. Technical appendix
  [void]$body.AppendLine('<section id="appendix"><h2>9. Technical appendix</h2><details><summary>Source qualification and controls</summary>')
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
  Add-ReportTable $body @('Control','Key','Source','FinOps','Status','Detail') @(Get-ReportSorted $controls { param($x) @([string]$x.Control, [string]$x.Key) } | ForEach-Object {
    ,@((ConvertTo-ReportHtml $_.Control), (ConvertTo-ReportHtml (ConvertTo-ReportSafeText (& $mapKey $_.Key))), (ConvertTo-ReportHtml $_.SourceCount), (ConvertTo-ReportHtml $_.FinOpsCount), (ConvertTo-ReportHtml $_.Status), (ConvertTo-ReportHtml (ConvertTo-ReportSafeText $_.Detail)))
  }) @(2,3)
  [void]$body.AppendLine('</details></section>')
  #endregion

  return New-ReportPage -Title 'SmartFinOps Workplace - FinOps report' -Header $header.ToString().TrimEnd() -Body $body.ToString()
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
if ($MyInvocation.InvocationName -eq '.') { return }

$exitCode = 0
$startedAt = Get-Date
$root = $null; $archiveDir = $null
try {
  . (Join-Path $PSScriptRoot 'Config\SmartFinOps-TenantContext.ps1')
  Write-SmartFinOpsStartupBanner
  Write-ReportLog ("SmartFinOps Workplace Report {0} - tenant '{1}'{2}" -f $script:ScriptVersion, $Tenant, $(if ($ValidateOnly) { ' - ValidateOnly (no file written)' } else { '' }))
  Write-ReportLog ("Tenant context {0}" -f $script:SmartFinOpsTenantContextVersion)
  Initialize-SmartFinOpsTenantContext -Tenant $Tenant -StartPath $PSScriptRoot | Out-Null
  $resolvedRoot = if ($InputRoot) { $InputRoot } else { [string](Resolve-SmartFinOpsConfigTokenValue -Value '{{TenantDataRootPath}}') }
  Assert-SmartFinOpsResolvedPath -Name 'TenantDataRootPath' -Value $resolvedRoot
  $root = ConvertTo-SmartFinOpsFileSystemPath $resolvedRoot
  $reportConfig = Get-SmartFinOpsConfigValue -Name 'Report'
  $minimum = Get-ReportProperty $reportConfig 'ShowbackMinimumGroupUsers'
  $minimumUsers = if ($null -eq $minimum -or [string]$minimum -eq '') { 5 } else { [int]$minimum }
  if ($minimumUsers -lt 1) { throw 'Report.ShowbackMinimumGroupUsers must be a positive integer.' }
  Write-ReportLog "Inputs: $root"
  if (-not (Test-Path -LiteralPath $root -PathType Container)) { throw "FinOps data folder not found: $root" }

  $inputs = Read-ReportInputs -Root $root -TenantKey $Tenant
  foreach ($check in $inputs.Checks) { Write-ReportLog ('{0}: {1} - {2}' -f $check.Rule, $check.Status, $check.Detail) $(if ($check.Status -eq 'Pass') { 'INFO' } else { 'WARNING' }) }
  if ($inputs.Available) {
    $html = New-ReportPageContent -Inputs $inputs -Tenant $Tenant -MinimumGroupUsers $minimumUsers
    Write-ReportLog ("Report built for run {0} (Licences measured: {1}; showback minimum group: {2} users)." -f $inputs.RunId, [bool](Get-ReportProperty $inputs.Receipt 'LicencesMeasured'), $minimumUsers)
  }
  else {
    $html = New-ReportUnavailablePage -Inputs $inputs -Tenant $Tenant
    Write-ReportLog "Report not available: $($inputs.Reason)" 'WARNING'
    $exitCode = 2
  }
  if ($ValidateOnly) { Write-ReportLog 'ValidateOnly: no file written.' }
  else {
    Write-ReportFile -Path (Join-Path $root $script:ReportName) -Content $html
    if ($inputs.RunId -match '^\d{8}T\d{9}Z-[0-9a-f]{8}$') {
      $candidate = Join-Path (Join-Path $root 'Archive') $inputs.RunId
      if (Test-Path -LiteralPath $candidate -PathType Container) { $archiveDir = $candidate; Write-ReportFile -Path (Join-Path $archiveDir $script:ReportName) -Content $html }
    }
    Write-ReportLog ("Report written: {0}{1}" -f (Join-Path $root $script:ReportName), $(if ($archiveDir) { ' (archived with the run)' } else { '' }))
  }
}
catch {
  $exitCode = 1
  Write-ReportLog ("Report failed: {0}" -f $_.Exception.Message) 'ERROR'
  Write-ReportLog ("Stack: {0}" -f ($_.ScriptStackTrace -replace '\r?\n', ' | ')) 'ERROR'
}
finally {
  $logPath = ''
  if (-not $ValidateOnly -and $root -and (Test-Path -LiteralPath $root -PathType Container)) {
    $content = ($script:LogLines -join [Environment]::NewLine) + [Environment]::NewLine
    try {
      Write-ReportFile -Path (Join-Path $root $script:LogName) -Content $content
      $logPath = Join-Path $root $script:LogName
    }
    catch { Write-Host "Report log could not be written: $($_.Exception.Message)" -ForegroundColor Red }
  }
  # Completion banner on success and on failure; "Report not available" (exit 2) is a failed run.
  $bannerStatus = if ($exitCode -ne 0) { 'Failed' } else { 'Auto' }
  $warnings = @($script:LogLines | Where-Object { $_ -match '^\S+ \S+ \[WARNING\]' }).Count
  $errors = @($script:LogLines | Where-Object { $_ -match '^\S+ \S+ \[ERROR\]' }).Count
  if (Get-Command -Name Write-SmartFinOpsCompletionBanner -ErrorAction SilentlyContinue) {
    Write-SmartFinOpsCompletionBanner -Status $bannerStatus -ScriptName 'SmartFinOps-Workplace-Report.ps1' -StartedAt $startedAt -WarningCount $warnings -ErrorCount $errors -GeneratedCsvFiles 0 -LogPath $logPath
  }
  else { Write-Host ("SmartFinOps-Workplace-Report.ps1: {0} (tenant context not loaded)" -f $(if ($exitCode -ne 0) { 'FAILED' } else { 'SUCCESS' })) }
  # The archived log is a copy of the current one, banner included.
  if ($logPath -and $archiveDir) {
    try { Copy-Item -LiteralPath $logPath -Destination (Join-Path $archiveDir $script:LogName) -Force }
    catch { Write-Host "Archived report log could not be written: $($_.Exception.Message)" -ForegroundColor Red }
  }
}
exit $exitCode
#endregion

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBVvwvjdOqYszdQ
# q1K9tlXO2U3UIVBzSHHWhgo3iO8726CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICwFmVmyHRcdES5CDjUGHnygRb/zq7v3wcrv68yahE+bMA0GCSqG
# SIb3DQEBAQUABIIBgA6k6QnOmnEv/mPs590lBoT4nxgtLQvXBDe0qlde2CQZlkDx
# HaLn3hs4VqhNYIKF6KBJO1YbIW3UjeGwpDzO/2deLJCd8b78KzJZ6arcQWFQu/WN
# qM+TLqgT/lttdhaiRqAJK8oq6lJzLqQ8qrOXZUyT18gI5/OnzwJC3JHHWx+nE9/x
# 5hoVnYU/AnNmVgv3U0SYjwGOEF8frdpC1VF02uPWZXNCvNeKr4Va0aLTL6V/Q8zm
# clxHQZx7LdjmlkNI+GdSKAoRi6T0PQmz5yLie2hYvM9Nneo2/7gjOgmwrJuivBIG
# mA6i0AnJt05gTSy3VPD9pD3AvbF0Z5L/0V6TxoYRDQpY17P/JeeMa2h4Dvkp/xxl
# SGqzs4twMp5sLTPILzG0WXOa35OoVUfoM3DtNg7ueh30z4HmkZPRwEe5xYzTSwwy
# Fq9c+hwEKEEMaMp6GlDsyoV4VHhn/zlAs37I6o67y2HAYbjXRHqjcxb04JRRwxt3
# wJIShYWFLIEl9SaPk6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTgxMDE1
# MTVaMC8GCSqGSIb3DQEJBDEiBCDXhPRU6piCSb6sG3UGkNMRU2RdXFgJFP4y3hKL
# 3TwWaTANBgkqhkiG9w0BAQEFAASCAgBND1lnXjpZISnZaSMWEG0qwFksv1EUj5mb
# USRPj+AQ0+ZuAorNtNXP1cpfQfSgF8FStWUNJ2UVKT3IzYOvAHnHCR/n2pBeYwHw
# ZKsUftUftWTpzkBOIskYNmDcune93Pocgb6CFpOlkOrVPOGwtrKp0rYs1wA4DQYV
# U0XcjPIWJznBeBVLT5O+52chBLWHBP2cNDs3Xqf4hsf2dQlXtk9S+gk+AKY258DZ
# HuvzHtxMFoZ17gT0R1vChzb4Hy/nEVGM8avH4QjPGjLATNRjASlKMnpjtYfzfeYx
# 5EJ4J8fQD+0dzWiRuO+xd4e5ayXVH0zoomihDczRlEMS0DoPkQAwzFz5i4+lFjCb
# ItTO0Fz4tgkcLY981ASVJdlYz51dswJBra7mA4gn28asxdgg/Vm9zxTI/NYL4K6A
# p6Av/UEVRsd0OOpaFvSPCC9LNPqBWzBM/qJr5JNAMnV4eJquBbQW6sm56f0k9wb7
# rH/EmPg24JEfdBNC6aq7atyYT0Iz+SCWuzL+/fyIk2bGKVM6jNrXTE345BSJfU8u
# quIQn/9bJkyxo/N46TupkKLeO02xuj/MKupLRmyNWhYCd6EOlZfjPMWwtDnMKjmF
# 4kiIwBHU/syrKG0mj70Jq0n6AkCM4fPeXwczIs0x64e88H1rDJ/J6wDB+RUqfGAO
# YkS5qjEtmA==
# SIG # End signature block
