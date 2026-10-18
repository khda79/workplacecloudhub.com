#Requires -Version 7.0
<#
.SYNOPSIS
  SmartFinOps Workplace V2 valuation engine (read-only).
.DESCRIPTION
  Values the license decisions already published by SmartInventory. It never recomputes a decision:
  it reads the Licences report (five CSVs and snapshot), the per-SKU decision CSVs, license prices
  and optional Windows ESU prices, each qualified by its producer source receipt, then writes
  FinOps_* CSVs with their own receipt.
  Unqualified input is reported as Not measured, never as zero or as an older value.
.VERSION
2.1.0
.NOTES
  Outputs: Data\<TenantKey>\FinOps_*.csv, SmartFinOps_Workplace-Valuation.current.json.txt and
  Archive\<FinOpsRunId>\ (same files plus the run log). Consumers must check the receipt
  (Status Completed, TenantKey, SHA256 per file) before reading any FinOps_* CSV.
#>
[CmdletBinding()]
param(
  [string]$Tenant = 'test',
  [string]$SmartM365LatestCsvFolderPath,
  [string]$OutputRoot,
  [string]$ContractsPath,
  [string]$AsOfUtc,
  [switch]$ValidateOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:ScriptVersion = '2.1.0'
$script:LogLines = [System.Collections.Generic.List[string]]::new()
$script:Invariant = [Globalization.CultureInfo]::InvariantCulture

#region Helpers
function Write-FinOpsLog {
  param([Parameter(Mandatory)][AllowEmptyString()][string]$Message, [ValidateSet('INFO','WARNING','ERROR')][string]$Level = 'INFO')
  foreach ($part in ($Message -split "`r?`n")) {
    $line = '{0} [{1}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Level, $part
    $script:LogLines.Add($line)
    $color = switch ($Level) { 'WARNING' { 'Yellow' } 'ERROR' { 'Red' } default { 'Gray' } }
    Write-Host $line -ForegroundColor $color
  }
}

function Get-FinOpsProperty {
  param([AllowNull()]$InputObject, [Parameter(Mandatory)][string]$Name)
  if ($null -eq $InputObject) { return $null }
  if ($InputObject -is [System.Collections.IDictionary]) { if ($InputObject.Contains($Name)) { return $InputObject[$Name] } return $null }
  $property = $InputObject.PSObject.Properties[$Name]
  if ($property) { return $property.Value }
  return $null
}

function ConvertTo-FinOpsDecimal {
  param([AllowNull()]$Value)
  if ($null -eq $Value) { return $null }
  if ($Value -is [decimal] -or $Value -is [int] -or $Value -is [long] -or $Value -is [double]) { return [decimal]$Value }
  $text = ([string]$Value).Trim()
  # One optional decimal separator, '.' or ',' (collectors write prices in the host culture); no group separator,
  # so '37,78' is never read as 3778.
  if ($text -notmatch '^-?\d+(?:[.,]\d+)?$') { return $null }
  $number = [decimal]0
  if ([decimal]::TryParse($text.Replace(',', '.'), [Globalization.NumberStyles]::AllowLeadingSign -bor [Globalization.NumberStyles]::AllowDecimalPoint, $script:Invariant, [ref]$number)) { return $number }
  return $null
}

function ConvertTo-FinOpsLong {
  param([AllowNull()]$Value)
  $number = ConvertTo-FinOpsDecimal $Value
  if ($null -eq $number -or $number -ne [decimal]::Truncate($number)) { return $null }
  return [long]$number
}

function Format-FinOpsDecimal {
  param([AllowNull()]$Value)
  if ($null -eq $Value) { return '' }
  return ([decimal]$Value).ToString('0.00', $script:Invariant)
}

function ConvertTo-FinOpsUtc {
  param([AllowNull()]$Value)
  if ($null -eq $Value) { return $null }
  if ($Value -is [datetime]) { return $Value.ToUniversalTime() }
  if ($Value -is [datetimeoffset]) { return $Value.UtcDateTime }
  $parsed = [datetimeoffset]::MinValue
  if ([datetimeoffset]::TryParse(([string]$Value).Trim(), $script:Invariant, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsed)) { return $parsed.UtcDateTime }
  return $null
}

function ConvertTo-FinOpsDate {
  param([AllowNull()]$Value)
  $text = ([string]$Value).Trim()
  if (-not $text) { return $null }
  $parsed = [datetime]::MinValue
  if ([datetime]::TryParseExact($text, 'yyyy-MM-dd', $script:Invariant, [Globalization.DateTimeStyles]::None, [ref]$parsed)) { return $parsed.Date }
  return $null
}

function Test-FinOpsTrue { param([AllowNull()]$Value) return ([string]$Value).Trim() -ieq 'true' }

function Get-FinOpsSha256 { param([Parameter(Mandatory)][string]$Path) return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash }

function Get-FinOpsCsvHeader {
  param([Parameter(Mandatory)][string]$Path, [string]$Delimiter = ',')
  $line = Get-Content -LiteralPath $Path -TotalCount 1 -Encoding utf8
  if ($null -eq $line) { return @() }
  return @(([string]$line).TrimStart([char]0xFEFF).Split($Delimiter) | ForEach-Object { $_.Trim('"') })
}

function Read-FinOpsJsonText {
  param([Parameter(Mandatory)][string]$Path)
  if ((Get-Item -LiteralPath $Path).Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Linked file is not accepted: $Path" }
  return [IO.File]::ReadAllText($Path)
}

function Get-FinOpsSortedRows {
  param([AllowEmptyCollection()][object[]]$Rows, [Parameter(Mandatory)][string[]]$KeyColumns)
  # Ordinal, case-insensitive sort on the key: the Licences row order changes between runs.
  $list = @($Rows | Where-Object { $null -ne $_ })
  if ($list.Count -le 1) { return ,$list }
  $pairs = [System.Collections.Generic.List[object]]::new()
  foreach ($item in $list) {
    $pairs.Add([pscustomobject]@{ Key=(@(foreach ($column in $KeyColumns) { [string](Get-FinOpsProperty $item $column) }) -join [char]31); Item=$item })
  }
  $pairs.Sort([Comparison[object]]{ param($left, $right) [StringComparer]::OrdinalIgnoreCase.Compare($left.Key, $right.Key) })
  return ,[object[]]@($pairs | ForEach-Object { $_.Item })
}

function Write-FinOpsCsv {
  param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string[]]$Columns, [AllowEmptyCollection()][object[]]$Rows)
  $builder = [System.Text.StringBuilder]::new()
  [void]$builder.AppendLine((@($Columns | ForEach-Object { '"' + $_ + '"' }) -join ','))
  foreach ($row in $Rows) {
    $values = foreach ($column in $Columns) {
      $value = Get-FinOpsProperty $row $column
      $text = if ($null -eq $value) { '' } elseif ($value -is [bool]) { if ($value) { 'True' } else { 'False' } } elseif ($value -is [decimal]) { Format-FinOpsDecimal $value } else { [string]$value }
      '"' + $text.Replace('"', '""') + '"'
    }
    [void]$builder.AppendLine(($values -join ','))
  }
  $temporary = $Path + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
  try {
    [IO.File]::WriteAllText($temporary, $builder.ToString(), [Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $temporary -Destination $Path -Force
  }
  finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force } }
}

function Write-FinOpsJson {
  param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)]$Document)
  $temporary = $Path + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
  try {
    [IO.File]::WriteAllText($temporary, ($Document | ConvertTo-Json -Depth 12), [Text.UTF8Encoding]::new($false))
    Move-Item -LiteralPath $temporary -Destination $Path -Force
  }
  finally { if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force } }
}
#endregion

#region Output contract
$script:CommonColumns = @('TenantKey','FinOpsRunId','SnapshotId','ReceiptRunId')
$script:OutputSpecs = [ordered]@{
  'FinOps_Sources.csv' = @{ Key=@('SourceFile'); Columns=@('SourceFile','Producer','ReceiptFile','SourceReceiptRunId','ReceiptStatus','ReceiptCompletedAtUtc','Required','PublishedInReceipt','ExpectedSHA256','ActualSHA256','ExpectedRows','ActualRows','FreshnessThresholdHours','AgeHours','QualificationStatus','Reason','MissingDomains','ProvisionalLevers') }
  'FinOps_Opportunities.csv' = @{ Key=@('OpportunityKey'); Columns=@('OpportunityKey','Lever','ValueClass','ValuationStatus','StatusReason','SourceFile','SourceCandidateType','SourceRecommendation','SourceReason','MailboxGuard','DependencyStatus','CountedIn','ReportProduct','SkuId','SkuPartNumber','TargetSkuPartNumber','UserId','UserPrincipalName','Units','UnitPriceMonthly','TargetUnitPriceMonthly','MonthlyAmount','AnnualAmount','Currency','PriceRowKey','PricingReceiptRunId','PriceSource','PriceIsEstimated','CapRank','ContractKey','RenewalDate','RealizableFrom','AdDomain','Country','Entity','Department','ShowbackJoinStatus') }
  'FinOps_SkuValuation.csv' = @{ Key=@('SkuId'); Columns=@('SkuId','SkuPartNumber','SkuDisplayName','ReportProduct','Excluded','EnabledUnits','ConsumedUnits','AvailableUnits','ReserveUnits','MinimumUnits','CandidateUnitsR','PoolUnits','ReusedUnits','ReducibleUnits','CapUnits','RealizableUnits','CapacityRealizableUnits','CappedUnits','ReviewUnits','NotMeasuredUnits','BlockedUnits','ReferencedUnits','UnitPriceMonthly','PriceRowKey','PriceIsEstimated','RealizableAnnualAmount','AvoidedAnnualAmount','ValuationStatus','ContractKey','ContractStatus','ReductionPolicy','RenewalDate','RealizableFrom','UsageEvidence') }
  'FinOps_Showback.csv' = @{ Key=@('Dimension','DimensionValue','Lever','ValueClass','ValuationStatus'); Columns=@('Dimension','DimensionValue','Lever','ValueClass','ValuationStatus','Users','Units','AnnualAmount','Currency','Provisional') }
  'FinOps_Summary.csv' = @{ Key=@('KpiKey'); Columns=@('KpiKey','ValueClass','Lever','ValuationStatus','Users','Units','AnnualAmount','Currency','Provisional','MissingDomains','SnapshotGeneratedAtUtc','PricingReceiptRunId','Note') }
  'FinOps_Realization.csv' = @{ Key=@('SkuId','BaselineFinOpsRunId'); Columns=@('SkuId','SkuPartNumber','BaselineFinOpsRunId','BaselineSnapshotId','EnabledUnitsBefore','EnabledUnitsAfter','Decrease','RemovedCandidateLicenses','BaselineCapacityUnits','AttributableUnits','RealizedUnits','NotAttributedUnits','UnitPriceMonthly','PriceRowKey','RealizedAnnualAmount','Status') }
  'FinOps_Controls.csv' = @{ Key=@('Control','Key'); Columns=@('Control','Key','SourceCount','FinOpsCount','Status','Detail') }
}
$script:ReceiptName = 'SmartFinOps_Workplace-Valuation.current.json.txt'
$script:LogName = 'SmartFinOps-Workplace-Valuation.log'
$script:RecoveryLevers = @{ 'Disabled account'='L1'; 'Eligible shared mailbox under 50 GB'='L2'; 'No M365 activity in 90 days'='L3' }
$script:CapRankByLever = @{ L1=1; L5=2; L2=3; L6=4; L3=5 }
$script:ValuedStatuses = @('Valued','Provisional')
# Which valued levers use a Licences report source; a provisional source makes only these levers Provisional.
# Any other provisional source makes every valued lever Provisional (fail-safe).
$script:ValuedLevers = @('L1','L2','L3','L4','L5','L6','L7','L8')
$script:ProvisionalLeversBySource = @{
  'M365_OneDrive_Usage.csv' = @('L4')
  'AD_Users_AllDomains.csv' = @('L1','L2','L3','L4','L5','L6')
}
#endregion

#region Source qualification
function New-FinOpsSourceRecord {
  param([string]$FileName, [string]$Producer, [string]$ReceiptFile, [bool]$Required, [int]$FreshnessHours)
  [pscustomobject][ordered]@{
    SourceFile=$FileName; Producer=$Producer; ReceiptFile=$ReceiptFile; SourceReceiptRunId=''; ReceiptStatus=''; ReceiptCompletedAtUtc=''
    Required=$Required; PublishedInReceipt=$false; ExpectedSHA256=''; ActualSHA256=''; ExpectedRows=''; ActualRows=''
    FreshnessThresholdHours=$FreshnessHours; AgeHours=''; QualificationStatus='NotMeasured'; Reason=''; MissingDomains=''; ProvisionalLevers=''
    Rows=$null; Path=''; Receipt=$null; Delimiter=','
  }
}

function Get-FinOpsSource {
  param(
    [Parameter(Mandatory)][string]$Folder,
    [Parameter(Mandatory)][string]$FileName,
    [Parameter(Mandatory)][string]$ReceiptFile,
    [Parameter(Mandatory)][string]$Producer,
    [Parameter(Mandatory)][string]$TenantKey,
    [Parameter(Mandatory)][datetime]$RunStartUtc,
    [Parameter(Mandatory)][int]$FreshnessHours,
    [string[]]$RequiredColumns = @(),
    [switch]$Required,
    [switch]$AllowPartialAd
  )
  $record = New-FinOpsSourceRecord -FileName $FileName -Producer $Producer -ReceiptFile $ReceiptFile -Required ([bool]$Required) -FreshnessHours $FreshnessHours
  $receiptPath = Join-Path $Folder $ReceiptFile
  $path = Join-Path $Folder $FileName
  $record.Path = $path
  try {
    if (-not (Test-Path -LiteralPath $receiptPath -PathType Leaf)) { throw 'source receipt is missing' }
    $receiptHash = Get-FinOpsSha256 $receiptPath
    $receipt = (Read-FinOpsJsonText $receiptPath) | ConvertFrom-Json
    $record.Receipt = $receipt
    $record.SourceReceiptRunId = [string](Get-FinOpsProperty $receipt 'RunId')
    $record.ReceiptStatus = [string](Get-FinOpsProperty $receipt 'Status')
    $completedAt = ConvertTo-FinOpsUtc (Get-FinOpsProperty $receipt 'CompletedAtUtc')
    if ($completedAt) { $record.ReceiptCompletedAtUtc = $completedAt.ToString('o') }
    if ((Get-FinOpsProperty $receipt 'Owner') -cne 'SmartInventory-SourceReceipt' -or [string](Get-FinOpsProperty $receipt 'ContractVersion') -cne '1.2') { throw 'unsupported source receipt owner or version' }
    if ([string](Get-FinOpsProperty $receipt 'Producer') -cne $Producer) { throw 'source receipt producer differs' }
    if ([string](Get-FinOpsProperty $receipt 'TenantKey') -cne $TenantKey) { throw 'source receipt tenant differs' }
    if ($record.ReceiptStatus -cne 'Completed') { throw "source receipt is $($record.ReceiptStatus)" }
    if (-not $record.SourceReceiptRunId) { throw 'source receipt RunId is missing' }
    # A newer run publishing over the qualified files makes them unreadable until it completes.
    try { Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof $receipt }
    catch { throw "source publication is in progress or unqualified: $($_.Exception.Message)" }
    $optional = @(Get-FinOpsProperty $receipt 'OptionalFiles' | Where-Object { $_ -and [string](Get-FinOpsProperty $_ 'File') -ceq $FileName })
    $fileRecords = @(Get-FinOpsProperty $receipt 'Files' | Where-Object { $_ -and [string](Get-FinOpsProperty $_ 'File') -ceq $FileName })
    if ($fileRecords.Count -eq 0) {
      if ($optional.Count -eq 1 -and -not [bool](Get-FinOpsProperty $optional[0] 'Published')) { throw 'not published by the current source receipt' }
      throw 'file is not recorded by the current source receipt'
    }
    if ($fileRecords.Count -ne 1) { throw 'file is recorded more than once by the source receipt' }
    $fileRecord = $fileRecords[0]
    $record.PublishedInReceipt = $true
    $record.ExpectedSHA256 = [string](Get-FinOpsProperty $fileRecord 'SHA256')
    $record.ExpectedRows = [string](Get-FinOpsProperty $fileRecord 'Rows')
    $delimiter = [string](Get-FinOpsProperty $fileRecord 'Delimiter'); if ($delimiter) { $record.Delimiter = $delimiter }
    if ([string](Get-FinOpsProperty $fileRecord 'Status') -cne 'Success') { throw 'file receipt is not successful' }
    # Partial inventories are accepted only for AD PartialAccepted coverage, as Provisional.
    $partial = [bool](Get-FinOpsProperty $receipt 'IsPartialInventory') -or [bool](Get-FinOpsProperty $fileRecord 'IsPartialInventory')
    $provisional = $false
    if ($partial) {
      $coverage = Get-FinOpsProperty $receipt 'DomainCoverage'
      if (-not $AllowPartialAd -or [string](Get-FinOpsProperty $coverage 'Status') -cne 'PartialAccepted' -or -not [bool](Get-FinOpsProperty $receipt 'ConsumerScopeQualified')) { throw 'partial source receipt is not accepted' }
      $record.MissingDomains = (@(Get-FinOpsProperty $coverage 'UnavailableDomains' | Sort-Object) -join ';')
      $provisional = $true
    }
    if (-not $completedAt) { throw 'source receipt completion date is missing' }
    $age = ($RunStartUtc - $completedAt).TotalHours
    $record.AgeHours = [math]::Round($age, 1).ToString($script:Invariant)
    if ($age -gt $FreshnessHours) { throw ('source is stale: {0:0.0} h > {1} h' -f $age, $FreshnessHours) }
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw 'file is missing' }
    $before = Get-FinOpsSha256 $path
    $record.ActualSHA256 = $before
    if ($before -cne $record.ExpectedSHA256) { throw 'file hash differs from the source receipt' }
    $header = Get-FinOpsCsvHeader -Path $path -Delimiter $record.Delimiter
    foreach ($column in $RequiredColumns) { if ($header -notcontains $column) { throw "missing column: $column" } }
    $rows = @(Import-Csv -LiteralPath $path -Delimiter $record.Delimiter -Encoding utf8)
    $record.ActualRows = [string]$rows.Count
    if ($record.ExpectedRows -ne [string]$rows.Count) { throw 'row count differs from the source receipt' }
    if ((Get-FinOpsSha256 $path) -cne $before) { throw 'file changed during read' }
    if ((Get-FinOpsSha256 $receiptPath) -cne $receiptHash) { throw 'source receipt changed during read' }
    if ($header -contains 'TenantKey') { foreach ($row in $rows) { if ([string]$row.TenantKey -cne $TenantKey) { throw 'a row belongs to another tenant' } } }
    $record.Rows = $rows
    $record.QualificationStatus = if ($provisional) { 'Provisional' } else { 'Qualified' }
    if ($provisional) { $record.Reason = 'Partial AD coverage accepted (PartialAccepted)' }
  }
  catch {
    $record.QualificationStatus = 'NotMeasured'
    $record.Reason = $_.Exception.Message
    $record.Rows = $null
  }
  return $record
}

function Set-FinOpsSourceNotMeasured {
  param([Parameter(Mandatory)]$Record, [Parameter(Mandatory)][string]$Reason)
  $Record.QualificationStatus = 'NotMeasured'
  $Record.Reason = if ($Record.Reason -and $Record.Reason -notlike 'Partial AD*') { "$($Record.Reason); $Reason" } else { $Reason }
  $Record.Rows = $null
}

function Read-FinOpsLicencesSnapshot {
  param([Parameter(Mandatory)][string]$Path)
  $document = [System.Text.Json.JsonDocument]::Parse((Read-FinOpsJsonText $Path))
  try {
    $root = $document.RootElement
    $value = { param($element, [string]$name) $child = [System.Text.Json.JsonElement]::new(); if ($element.ValueKind -eq 'Object' -and $element.TryGetProperty($name, [ref]$child)) { return $child } return $null }
    $snapshotId = & $value $root 'SnapshotId'; $tenant = & $value $root 'TenantKey'; $schema = & $value $root 'SchemaVersion'
    $generated = & $value $root 'GeneratedAtUtc'; $run = & $value $root 'ReceiptRun'
    $runId = if ($null -ne $run) { & $value $run 'RunId' } else { $null }
    [pscustomobject]@{
      SnapshotId=$(if ($null -ne $snapshotId -and $snapshotId.ValueKind -eq 'String') { $snapshotId.GetString() } else { '' })
      TenantKey=$(if ($null -ne $tenant -and $tenant.ValueKind -eq 'String') { $tenant.GetString() } else { '' })
      SchemaVersion=$(if ($null -ne $schema -and $schema.ValueKind -eq 'Number') { $schema.GetInt32() } else { 0 })
      GeneratedAtUtc=$(if ($null -ne $generated -and $generated.ValueKind -eq 'String') { $generated.GetString() } else { '' })
      ReceiptRunId=$(if ($null -ne $runId -and $runId.ValueKind -eq 'String') { $runId.GetString() } else { '' })
    }
  }
  finally { $document.Dispose() }
}
#endregion

#region Contracts
function Get-FinOpsContracts {
  param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][string]$TemplatePath)
  if (-not (Test-Path -LiteralPath $TemplatePath -PathType Leaf)) { throw "Contracts template not found: $TemplatePath" }
  $template = (Read-FinOpsJsonText $TemplatePath) | ConvertFrom-Json
  $changed = $false
  if (Test-Path -LiteralPath $Path -PathType Leaf) { $config = (Read-FinOpsJsonText $Path) | ConvertFrom-Json }
  else {
    $config = (Read-FinOpsJsonText $TemplatePath) | ConvertFrom-Json
    $changed = $true
    Write-FinOpsLog "Contracts file created from template: $Path"
  }
  $added = @(Add-SmartFinOpsMissingJsonTemplateProperties -Target $config -Template $template)
  if ($added.Count) { $changed = $true; Write-FinOpsLog ('Contracts file completed from template; added keys: {0}' -f ($added -join ', ')) }
  if ([string](Get-FinOpsProperty $config 'SchemaVersion') -ne '1') { throw 'Contracts file SchemaVersion must be 1.' }
  $defaults = Get-FinOpsProperty $config 'Defaults'
  $percent = ConvertTo-FinOpsDecimal (Get-FinOpsProperty $defaults 'ReserveUnitsPercent')
  if ($null -eq $percent -or $percent -lt 0 -or $percent -gt 100) { throw 'Contracts Defaults.ReserveUnitsPercent must be between 0 and 100.' }
  $policies = @('AtRenewal','Anniversary','Monthly','Unknown')
  $defaultPolicy = [string](Get-FinOpsProperty $defaults 'ReductionPolicy')
  if ($defaultPolicy -cnotin $policies) { throw 'Contracts Defaults.ReductionPolicy is invalid.' }
  $bySku = @{}
  foreach ($contract in @(Get-FinOpsProperty $config 'Contracts')) {
    if ($null -eq $contract) { continue }
    $key = [string](Get-FinOpsProperty $contract 'ContractKey')
    if (-not $key -or $key -ceq '__CONTRACT_KEY__') { continue }
    $policy = [string](Get-FinOpsProperty $contract 'ReductionPolicy'); if (-not $policy) { $policy = $defaultPolicy }
    if ($policy -cnotin $policies) { throw "Contract '$key': ReductionPolicy is invalid." }
    $dates = @{}
    foreach ($field in @('StartDate','RenewalDate')) {
      $text = ([string](Get-FinOpsProperty $contract $field)).Trim()
      $date = ConvertTo-FinOpsDate $text
      if ($text -and $null -eq $date) { throw "Contract '$key': $field must be yyyy-MM-dd." }
      $dates[$field] = $date
    }
    foreach ($sku in @(Get-FinOpsProperty $contract 'Skus')) {
      $part = ([string](Get-FinOpsProperty $sku 'SkuPartNumber')).Trim()
      if (-not $part) { continue }
      $upper = $part.ToUpperInvariant()
      if ($bySku.ContainsKey($upper)) { throw "SKU '$part' belongs to more than one contract." }
      $minimum = Get-FinOpsProperty $sku 'MinimumUnits'; $reserve = Get-FinOpsProperty $sku 'ReserveUnits'
      $minimumValue = if ($null -eq $minimum -or [string]$minimum -eq '') { $null } else { ConvertTo-FinOpsLong $minimum }
      $reserveValue = if ($null -eq $reserve -or [string]$reserve -eq '') { $null } else { ConvertTo-FinOpsLong $reserve }
      if (($null -ne $minimum -and [string]$minimum -ne '' -and ($null -eq $minimumValue -or $minimumValue -lt 0)) -or
          ($null -ne $reserve -and [string]$reserve -ne '' -and ($null -eq $reserveValue -or $reserveValue -lt 0))) { throw "Contract '$key', SKU '$part': MinimumUnits and ReserveUnits must be non-negative integers." }
      $bySku[$upper] = [pscustomobject]@{ ContractKey=$key; ReductionPolicy=$policy; RenewalDate=$dates.RenewalDate; MinimumUnits=$minimumValue; ReserveUnits=$reserveValue }
    }
  }
  return [pscustomobject]@{ Config=$config; BySku=$bySku; ReserveUnitsPercent=$percent; DefaultPolicy=$defaultPolicy; Changed=$changed; Path=$Path }
}

function Save-FinOpsContracts {
  param([Parameter(Mandatory)]$Contracts, [string[]]$UnmappedSkus = @())
  $config = $Contracts.Config
  $existing = @(Get-FinOpsProperty $config 'UnmappedSkus' | Where-Object { $_ } | ForEach-Object { [string]$_ })
  $merged = @($existing + $UnmappedSkus | Sort-Object -Unique)
  if (@(Compare-Object -ReferenceObject @($existing | Sort-Object -Unique) -DifferenceObject $merged -SyncWindow 0).Count -gt 0) {
    $config.UnmappedSkus = $merged
    $Contracts.Changed = $true
    Write-FinOpsLog ('Contracts file: {0} paid SKU(s) without contract listed in UnmappedSkus.' -f ($merged.Count - $existing.Count))
  }
  if ($Contracts.Changed) {
    New-Item -ItemType Directory -Path (Split-Path -Path $Contracts.Path -Parent) -Force | Out-Null
    Write-FinOpsJson -Path $Contracts.Path -Document $config
  }
}

function Get-FinOpsContractTerms {
  param([Parameter(Mandatory)]$Contracts, [Parameter(Mandatory)][string]$SkuPartNumber, [long]$EnabledUnits, [datetime]$RunDate)
  $entry = $Contracts.BySku[$SkuPartNumber.ToUpperInvariant()]
  $reserve = if ($entry -and $null -ne $entry.ReserveUnits) { [long]$entry.ReserveUnits } else { [long][math]::Floor([double]($EnabledUnits * $Contracts.ReserveUnitsPercent / 100)) }
  if (-not $entry) {
    return [pscustomobject]@{ ContractKey=''; ContractStatus='No contract'; ReductionPolicy='Unknown'; RenewalDate=''; RealizableFrom='Renewal unknown'; MinimumUnits=$null; ReserveUnits=$reserve }
  }
  $renewal = if ($entry.RenewalDate) { $entry.RenewalDate.ToString('yyyy-MM-dd') } else { '' }
  $from = switch ($entry.ReductionPolicy) {
    'Monthly' { $RunDate.ToString('yyyy-MM-dd') }
    { $_ -in @('AtRenewal','Anniversary') } { if ($renewal) { $renewal } else { 'Renewal unknown' } }
    default { 'Renewal unknown' }
  }
  [pscustomobject]@{
    ContractKey=$entry.ContractKey; ContractStatus=$(if ($null -eq $entry.MinimumUnits) { 'Commitment unknown' } else { 'Covered' })
    ReductionPolicy=$entry.ReductionPolicy; RenewalDate=$renewal; RealizableFrom=$from; MinimumUnits=$entry.MinimumUnits; ReserveUnits=$reserve
  }
}
#endregion

#region Prices
function New-FinOpsPriceIndex {
  param([AllowNull()]$Source, [Parameter(Mandatory)][datetime]$AsOfDate)
  $index = @{ Ready=$false; Reason=''; BySku=@{}; RunId='' }
  if ($null -eq $Source -or $Source.QualificationStatus -eq 'NotMeasured') { $index.Reason = 'license price source is not qualified'; return $index }
  $index.Ready = $true
  $index.RunId = $Source.SourceReceiptRunId
  foreach ($row in $Source.Rows) {
    if ([string]$row.Currency -cne 'EUR') { continue }
    $from = ConvertTo-FinOpsDate $row.EffectiveFrom; $to = ConvertTo-FinOpsDate $row.EffectiveTo
    if ($null -eq $from -or $from -gt $AsOfDate -or ($null -ne $to -and $to -le $AsOfDate)) { continue }
    $sku = ([string]$row.SkuPartNumber).Trim().ToUpperInvariant()
    if (-not $index.BySku.ContainsKey($sku)) { $index.BySku[$sku] = [System.Collections.Generic.List[object]]::new() }
    $index.BySku[$sku].Add($row)
  }
  return $index
}

function Get-FinOpsPrice {
  param([Parameter(Mandatory)]$PriceIndex, [AllowEmptyString()][string]$SkuPartNumber)
  if (-not $PriceIndex.Ready) { return [pscustomobject]@{ Valued=$false; Reason=$PriceIndex.Reason; Price=$null; RowKey=''; Source=''; Estimated=''; RunId='' } }
  if (-not $SkuPartNumber) { return [pscustomobject]@{ Valued=$false; Reason='SKU not resolved'; Price=$null; RowKey=''; Source=''; Estimated=''; RunId='' } }
  $rows = $PriceIndex.BySku[$SkuPartNumber.ToUpperInvariant()]
  if ($null -eq $rows -or $rows.Count -eq 0) { return [pscustomobject]@{ Valued=$false; Reason="no EUR price effective for $SkuPartNumber"; Price=$null; RowKey=''; Source=''; Estimated=''; RunId='' } }
  if ($rows.Count -gt 1) { return [pscustomobject]@{ Valued=$false; Reason="more than one EUR price effective for $SkuPartNumber"; Price=$null; RowKey=''; Source=''; Estimated=''; RunId='' } }
  $row = $rows[0]
  $price = ConvertTo-FinOpsDecimal $row.UnitPriceMonthly
  if ($null -eq $price -or $price -lt 0) { return [pscustomobject]@{ Valued=$false; Reason="invalid price for $SkuPartNumber"; Price=$null; RowKey=''; Source=''; Estimated=''; RunId='' } }
  [pscustomobject]@{
    Valued=$true; Reason=''; Price=$price; RowKey=('{0}|{1}|{2}|{3}' -f ([string]$row.SkuPartNumber).Trim(), $row.Currency, $row.EffectiveFrom, $row.EffectiveTo)
    Source=[string]$row.PriceSource; Estimated=[string]$row.IsEstimated; RunId=$PriceIndex.RunId
  }
}
#endregion

#region Opportunities
function New-FinOpsOpportunity {
  param([Parameter(Mandatory)][hashtable]$Values)
  $row = [ordered]@{}
  foreach ($column in $script:OutputSpecs['FinOps_Opportunities.csv'].Columns) { $row[$column] = '' }
  $row.Units = 1; $row.Currency = 'EUR'
  foreach ($key in $Values.Keys) { $row[$key] = $Values[$key] }
  return [pscustomobject]$row
}

function Set-FinOpsOpportunityValue {
  param([Parameter(Mandatory)]$Row, [Parameter(Mandatory)]$Price, [bool]$Provisional, [AllowNull()]$TargetPrice)
  $Row.PriceRowKey = $Price.RowKey; $Row.PricingReceiptRunId = $Price.RunId; $Row.PriceSource = $Price.Source; $Row.PriceIsEstimated = $Price.Estimated
  if (-not $Price.Valued) { $Row.ValuationStatus = 'Not valued'; $Row.StatusReason = $Price.Reason; return }
  $Row.UnitPriceMonthly = $Price.Price
  $unit = [decimal]$Price.Price
  if ($null -ne $TargetPrice) {
    if (-not $TargetPrice.Valued) { $Row.ValuationStatus = 'Not valued'; $Row.StatusReason = "target: $($TargetPrice.Reason)"; return }
    $Row.TargetUnitPriceMonthly = $TargetPrice.Price
    $unit = [decimal]$Price.Price - [decimal]$TargetPrice.Price
    if ($unit -le 0) { $Row.ValuationStatus = 'Not valued'; $Row.StatusReason = 'target price is not lower than the current price'; return }
  }
  $units = [decimal]$Row.Units
  $Row.MonthlyAmount = [decimal]($unit * $units)
  $Row.AnnualAmount = [decimal]($unit * $units * 12)
  $Row.ValuationStatus = if ($Provisional) { 'Provisional' } else { 'Valued' }
}

function Clear-FinOpsOpportunityValue {
  param([Parameter(Mandatory)]$Row, [Parameter(Mandatory)][string]$Status, [AllowEmptyString()][string]$Reason = '')
  $Row.MonthlyAmount = ''; $Row.AnnualAmount = ''
  $Row.ValuationStatus = $Status; $Row.StatusReason = $Reason
}
#endregion

#region Realization and archive
function Get-FinOpsArchiveRuns {
  param([Parameter(Mandatory)][string]$ArchiveRoot, [Parameter(Mandatory)][string]$TenantKey)
  $runs = [System.Collections.Generic.List[object]]::new()
  if (-not (Test-Path -LiteralPath $ArchiveRoot -PathType Container)) { return ,$runs }
  foreach ($directory in @(Get-ChildItem -LiteralPath $ArchiveRoot -Directory | Where-Object { $_.Name -match '^\d{8}T\d{9}Z-[0-9a-f]{8}$' } | Sort-Object Name)) {
    $receiptPath = Join-Path $directory.FullName $script:ReceiptName
    $entry = [pscustomobject]@{ RunId=$directory.Name; Path=$directory.FullName; Qualified=$false; Sku=@{}; Pairs=@{}; SnapshotId='' }
    try {
      if (-not (Test-Path -LiteralPath $receiptPath -PathType Leaf)) { throw 'receipt missing' }
      $receipt = (Read-FinOpsJsonText $receiptPath) | ConvertFrom-Json
      if ($receipt.Owner -cne 'SmartFinOps-Receipt' -or $receipt.Status -cne 'Completed' -or $receipt.TenantKey -cne $TenantKey -or $receipt.RunId -cne $directory.Name -or -not [bool]$receipt.LicencesMeasured) { throw 'not a qualified run' }
      foreach ($name in @('FinOps_SkuValuation.csv','FinOps_Opportunities.csv')) {
        $file = @($receipt.Files | Where-Object File -ceq $name)
        $path = Join-Path $directory.FullName $name
        if ($file.Count -ne 1 -or -not (Test-Path -LiteralPath $path -PathType Leaf) -or (Get-FinOpsSha256 $path) -cne $file[0].SHA256) { throw "$name does not match its receipt" }
      }
      foreach ($row in @(Import-Csv -LiteralPath (Join-Path $directory.FullName 'FinOps_SkuValuation.csv') -Encoding utf8)) {
        $entry.SnapshotId = [string]$row.SnapshotId
        $entry.Sku[[string]$row.SkuId] = $row
      }
      foreach ($row in @(Import-Csv -LiteralPath (Join-Path $directory.FullName 'FinOps_Opportunities.csv') -Encoding utf8)) {
        if ($row.ValueClass -cne 'R' -or $row.Lever -notin @('L1','L2','L3','L5','L6') -or $row.ValuationStatus -notin $script:ValuedStatuses -or -not $row.UserId -or -not $row.SkuId) { continue }
        $entry.Pairs[('{0}|{1}' -f $row.UserId.ToLowerInvariant(), $row.SkuId.ToLowerInvariant())] = $row.SkuId
      }
      $entry.Qualified = $true
    }
    catch { $entry.Qualified = $false }
    $runs.Add($entry)
  }
  return ,$runs
}

function Get-FinOpsRealization {
  param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$SkuRows, [Parameter(Mandatory)][AllowEmptyCollection()]$ArchiveRuns, [AllowNull()]$UsersSource)
  $results = [System.Collections.Generic.List[object]]::new()
  $protected = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
  $assigned = $null
  if ($UsersSource -and $UsersSource.QualificationStatus -ne 'NotMeasured') {
    $assigned = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $UsersSource.Rows) { [void]$assigned.Add(('{0}|{1}' -f [string]$row.UserId, [string]$row.SkuId)) }
  }
  $qualified = @($ArchiveRuns | Where-Object Qualified)
  foreach ($sku in $SkuRows) {
    $skuId = [string]$sku.SkuId
    $history = @($qualified | Where-Object { $_.Sku.ContainsKey($skuId) })
    $row = [ordered]@{ SkuId=$skuId; SkuPartNumber=$sku.SkuPartNumber; BaselineFinOpsRunId=''; BaselineSnapshotId=''; EnabledUnitsBefore=''; EnabledUnitsAfter=$sku.EnabledUnits
      Decrease=''; RemovedCandidateLicenses=''; BaselineCapacityUnits=''; AttributableUnits=''; RealizedUnits=''; NotAttributedUnits=''; UnitPriceMonthly=''; PriceRowKey=''; RealizedAnnualAmount=''; Status='' }
    if ($history.Count -eq 0) { $row.Status = 'Not measured: no qualified baseline run'; $results.Add([pscustomobject]$row); continue }
    # Plateau = contiguous most recent runs with the same EnabledUnits; its first run is the baseline.
    $latest = $history[-1]
    $level = ConvertTo-FinOpsLong $latest.Sku[$skuId].EnabledUnits
    $plateau = [System.Collections.Generic.List[object]]::new()
    for ($i = $history.Count - 1; $i -ge 0; $i--) {
      if ((ConvertTo-FinOpsLong $history[$i].Sku[$skuId].EnabledUnits) -ne $level) { break }
      $plateau.Insert(0, $history[$i])
    }
    $baseline = $plateau[0]
    [void]$protected.Add($baseline.RunId)
    $baselineSku = $baseline.Sku[$skuId]
    $row.BaselineFinOpsRunId = $baseline.RunId; $row.BaselineSnapshotId = $baseline.SnapshotId; $row.EnabledUnitsBefore = $level
    $decrease = [long]$level - [long]$sku.EnabledUnits
    if ($decrease -le 0) { $row.Decrease = 0; $row.RealizedUnits = 0; $row.Status = 'No reduction observed'; $results.Add([pscustomobject]$row); continue }
    $row.Decrease = $decrease
    if ($null -eq $assigned) { $row.Status = 'Not measured: license assignments not qualified'; $results.Add([pscustomobject]$row); continue }
    $pairs = @{}
    foreach ($run in $plateau) { foreach ($key in $run.Pairs.Keys) { if ($run.Pairs[$key] -ieq $skuId) { $pairs[$key] = $true } } }
    $removed = @($pairs.Keys | Where-Object { -not $assigned.Contains($_) }).Count
    $capacity = [long](ConvertTo-FinOpsLong $baselineSku.CapacityRealizableUnits)
    $capUnits = [long](@(foreach ($run in $plateau) { ConvertTo-FinOpsLong $run.Sku[$skuId].RealizableUnits }) | Measure-Object -Maximum).Maximum
    $attributable = [math]::Min([long]$removed + $capacity, $capUnits)
    $realized = [math]::Min($attributable, $decrease)
    $row.RemovedCandidateLicenses = $removed; $row.BaselineCapacityUnits = $capacity; $row.AttributableUnits = $attributable
    $row.RealizedUnits = $realized; $row.NotAttributedUnits = $decrease - $realized
    $price = ConvertTo-FinOpsDecimal $baselineSku.UnitPriceMonthly
    $row.PriceRowKey = $baselineSku.PriceRowKey
    if ($null -eq $price) { $row.Status = 'Not valued: no baseline price' }
    else { $row.UnitPriceMonthly = $price; $row.RealizedAnnualAmount = [decimal]($price * $realized * 12); $row.Status = 'Realized' }
    $results.Add([pscustomobject]$row)
  }
  return [pscustomobject]@{ Rows=$results.ToArray(); ProtectedRuns=$protected }
}

function Invoke-FinOpsArchiveRetention {
  param([Parameter(Mandatory)][string]$ArchiveRoot, [Parameter(Mandatory)][int]$KeepRuns, [Parameter(Mandatory)]$ProtectedRuns, [Parameter(Mandatory)][string]$CurrentRunId)
  if (-not (Test-Path -LiteralPath $ArchiveRoot -PathType Container)) { return }
  $directories = @(Get-ChildItem -LiteralPath $ArchiveRoot -Directory | Where-Object { $_.Name -match '^\d{8}T\d{9}Z-[0-9a-f]{8}$' } | Sort-Object Name -Descending)
  $kept = 0
  foreach ($directory in $directories) {
    $kept++
    if ($kept -le $KeepRuns -or $directory.Name -ceq $CurrentRunId -or $ProtectedRuns.Contains($directory.Name)) { continue }
    Remove-Item -LiteralPath $directory.FullName -Recurse -Force
    Write-FinOpsLog "Archive retention: removed run $($directory.Name)"
  }
}
#endregion

#region Engine
function Invoke-FinOpsValuation {
  param(
    [Parameter(Mandatory)][hashtable]$Sources,
    [Parameter(Mandatory)][datetime]$RunStartUtc,
    [Parameter(Mandatory)]$Contracts,
    [AllowNull()]$Snapshot
  )
  $lic = @('M365_Licenses_ReportSummary.csv','M365_Licenses_ReportCandidates.csv','M365_Licenses_ReportAssignmentIssues.csv','M365_Licenses_ReportGaps.csv','M365_Licenses_ReportSources.csv','M365_Licenses_ReportSkuCandidates.csv','M365_Licenses_ReportSkuCapacity.csv')
  $measured = -not @($lic | Where-Object { $Sources[$_].QualificationStatus -eq 'NotMeasured' })
  $snapshotId = if ($measured) { $Snapshot.SnapshotId } else { '' }
  $receiptRunId = if ($measured) { $Sources['M365_Licenses_ReportSummary.csv'].SourceReceiptRunId } else { '' }
  $asOfDate = if ($measured) { (ConvertTo-FinOpsUtc $Snapshot.GeneratedAtUtc).Date } else { $RunStartUtc.Date }
  $opportunities = [System.Collections.Generic.List[object]]::new()
  $skuRows = [System.Collections.Generic.List[object]]::new()
  $unmapped = [System.Collections.Generic.List[string]]::new()
  $provisionalReasons = [System.Collections.Generic.List[string]]::new()
  $provisionalLevers = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
  $provisionalTrace = [System.Collections.Generic.List[object]]::new()
  $missingDomains = ''
  $priceIndex = New-FinOpsPriceIndex -Source $Sources['M365_License_Prices.csv'] -AsOfDate $asOfDate

  if ($measured) {
    foreach ($row in $Sources['M365_Licenses_ReportSources.csv'].Rows) {
      if (-not (Test-FinOpsTrue $row.Provisional)) { continue }
      $levers = if ($script:ProvisionalLeversBySource.ContainsKey([string]$row.Name)) { $script:ProvisionalLeversBySource[[string]$row.Name] } else { $script:ValuedLevers }
      foreach ($lever in $levers) { [void]$provisionalLevers.Add($lever) }
      $provisionalTrace.Add([pscustomobject]@{ Source=[string]$row.Name; Levers=(@($levers) -join ';'); Reason=[string]$row.Reason })
      $provisionalReasons.Add(('{0} -> {1}: {2}' -f $row.Name, (@($levers) -join ';'), $row.Reason))
    }
    $adRecord = $Sources['AD_Users_AllDomains.csv']
    if ($adRecord.MissingDomains) { $missingDomains = $adRecord.MissingDomains }
  }
  $provisional = $provisionalLevers.Count -gt 0

  if ($measured) {
    $capacity = @{}
    foreach ($row in $Sources['M365_Licenses_ReportSkuCapacity.csv'].Rows) { $capacity[[string]$row.SkuId] = $row }
    $skusByProduct = @{}
    foreach ($row in $capacity.Values) {
      $product = [string]$row.ReportProduct
      if (-not $product) { continue }
      if (-not $skusByProduct.ContainsKey($product)) { $skusByProduct[$product] = [System.Collections.Generic.List[object]]::new() }
      $skusByProduct[$product].Add($row)
    }
    $resolveProductSku = {
      param([string]$Product)
      $candidates = @($skusByProduct[$Product] | Where-Object { $_ -and (ConvertTo-FinOpsLong $_.AssignedUsers) -gt 0 })
      if ($candidates.Count -eq 1) { return $candidates[0] }
      return $null
    }

    # Referenced rows of the per-SKU file point at ReportCandidates; they give the SKU of a suite candidate.
    $skuCandidates = @($Sources['M365_Licenses_ReportSkuCandidates.csv'].Rows)
    $referencedSku = @{}
    foreach ($row in $skuCandidates) {
      if (-not $row.CountedIn) { continue }
      $product = if ($capacity.ContainsKey([string]$row.SkuId)) { [string]$capacity[[string]$row.SkuId].ReportProduct } else { '' }
      $key = '{0}|{1}|{2}' -f ([string]$row.UserId).ToLowerInvariant(), $row.CountedIn, $product
      if ($referencedSku.ContainsKey($key)) { $referencedSku[$key] = $null } else { $referencedSku[$key] = $row }
    }

    # L1-L4 from ReportCandidates.
    foreach ($row in $Sources['M365_Licenses_ReportCandidates.csv'].Rows) {
      $type = [string]$row.CandidateType; $product = [string]$row.Product
      $values = @{ SourceFile='M365_Licenses_ReportCandidates.csv'; SourceCandidateType=$type; SourceReason=[string]$row.Reason; ReportProduct=$product; UserId=[string]$row.UserId; UserPrincipalName=[string]$row.UserPrincipalName }
      if ($type -ceq 'Recovery') {
        $lever = $script:RecoveryLevers[[string]$row.Reason]
        if (-not $lever) { throw "Unknown recovery reason in ReportCandidates: '$($row.Reason)'" }
        $countedIn = 'ReportCandidates:Recovery'
        $values.Lever = $lever; $values.ValueClass = 'R'
      }
      elseif ($type -ceq 'E3 to F3 review') {
        $countedIn = 'ReportCandidates:E3 to F3 review'
        $values.Lever = 'L4'; $values.ValueClass = 'Rc'
      }
      else { throw "Unknown CandidateType in ReportCandidates: '$type'" }
      $refKey = '{0}|{1}|{2}' -f ([string]$row.UserId).ToLowerInvariant(), $countedIn, $product
      $skuRow = $null; $skuReason = ''
      if ($referencedSku.ContainsKey($refKey)) {
        if ($null -ne $referencedSku[$refKey]) { $skuRow = $capacity[[string]$referencedSku[$refKey].SkuId] } else { $skuReason = 'SKU ambiguous: several referenced SKUs' }
      }
      else { $skuRow = & $resolveProductSku $product; if (-not $skuRow) { $skuReason = "SKU ambiguous or absent for product $product" } }
      if ($skuRow) { $values.SkuId = [string]$skuRow.SkuId; $values.SkuPartNumber = [string]$skuRow.SkuPartNumber }
      $values.OpportunityKey = '{0}|{1}|{2}' -f $values.Lever, $values.UserId, $(if ($skuRow) { $values.SkuId } else { $product })
      $opportunity = New-FinOpsOpportunity $values
      if (-not $skuRow) { Clear-FinOpsOpportunityValue $opportunity 'Not valued' $skuReason }
      elseif ($values.Lever -eq 'L4') {
        $targets = @($skusByProduct['Microsoft 365 F3'] | Where-Object { $_ })
        if ($targets.Count -ne 1) { Clear-FinOpsOpportunityValue $opportunity 'Not valued' 'target F3 SKU ambiguous or absent in the tenant' }
        else {
          $opportunity.TargetSkuPartNumber = [string]$targets[0].SkuPartNumber
          Set-FinOpsOpportunityValue $opportunity (Get-FinOpsPrice $priceIndex $values.SkuPartNumber) ($provisionalLevers.Contains('L4')) (Get-FinOpsPrice $priceIndex $opportunity.TargetSkuPartNumber)
          if ($opportunity.ValuationStatus -in $script:ValuedStatuses) { $opportunity.StatusReason = 'Conditional: Frontline eligibility and E3 features to confirm' }
        }
      }
      else { Set-FinOpsOpportunityValue $opportunity (Get-FinOpsPrice $priceIndex $values.SkuPartNumber) ($provisionalLevers.Contains($values.Lever)) $null }
      $opportunities.Add($opportunity)
    }

    # L0, L5, L6 from ReportSkuCandidates.
    foreach ($row in $skuCandidates) {
      $type = [string]$row.CandidateType; $recommendation = [string]$row.Recommendation
      $lever = switch -CaseSensitive ($type) { 'Disabled account' { 'L5' } 'Copilot without activity' { 'L6' } 'Report candidate' { 'L0' } default { throw "Unknown CandidateType in ReportSkuCandidates: '$type'" } }
      if ($recommendation -cnotin @('Recommended','Review','Blocked by dependency','Referenced')) { throw "Unknown Recommendation in ReportSkuCandidates: '$recommendation'" }
      $skuInfo = $capacity[[string]$row.SkuId]
      $opportunity = New-FinOpsOpportunity @{
        OpportunityKey=('{0}|{1}|{2}' -f $lever, $row.UserId, $row.SkuId); Lever=$lever; SourceFile='M365_Licenses_ReportSkuCandidates.csv'
        SourceCandidateType=$type; SourceRecommendation=$recommendation; SourceReason=[string]$row.Reason; MailboxGuard=[string]$row.MailboxGuard
        DependencyStatus=[string]$row.DependencyStatus; CountedIn=[string]$row.CountedIn; ReportProduct=$(if ($skuInfo) { [string]$skuInfo.ReportProduct } else { '' })
        SkuId=[string]$row.SkuId; SkuPartNumber=[string]$row.SkuPartNumber; UserId=[string]$row.UserId; UserPrincipalName=[string]$row.UserPrincipalName
      }
      switch -CaseSensitive ($recommendation) {
        'Recommended' { $opportunity.ValueClass = 'R'; Set-FinOpsOpportunityValue $opportunity (Get-FinOpsPrice $priceIndex $opportunity.SkuPartNumber) ($provisionalLevers.Contains($lever)) $null }
        'Referenced' { $opportunity.ValueClass = 'R'; Clear-FinOpsOpportunityValue $opportunity 'Referenced' "Counted in $($row.CountedIn)" }
        'Blocked by dependency' { $opportunity.ValueClass = 'S'; Clear-FinOpsOpportunityValue $opportunity 'Blocked' ([string]$row.Reason) }
        'Review' {
          $opportunity.ValueClass = 'S'
          if (([string]$row.MailboxGuard).StartsWith('Unknown', [StringComparison]::Ordinal)) { Clear-FinOpsOpportunityValue $opportunity 'Not measured' "Mailbox state not measured: $($row.MailboxGuard)" }
          else { Clear-FinOpsOpportunityValue $opportunity 'Review' ([string]$row.Reason) }
        }
      }
      $opportunities.Add($opportunity)
    }

    # Per-SKU pool, reuse and cap; L6/L7/L8/L9/L10 SKU-level rows.
    $waitByProduct = @{}
    foreach ($row in $Sources['M365_Licenses_ReportSummary.csv'].Rows) {
      if ([string]$row.Category -ceq 'Suite') { $wait = ConvertTo-FinOpsLong $row.AssignmentCapacityWait; if ($null -ne $wait -and $wait -gt 0) { $waitByProduct[[string]$row.Product] = $wait } }
    }
    $rRowsBySku = @{}
    foreach ($opportunity in $opportunities) {
      if ($opportunity.ValueClass -ne 'R' -or $opportunity.Lever -notin @('L1','L2','L3','L5','L6') -or -not $opportunity.SkuId -or $opportunity.ValuationStatus -notin @('Valued','Provisional','Not valued')) { continue }
      if (-not $rRowsBySku.ContainsKey($opportunity.SkuId)) { $rRowsBySku[$opportunity.SkuId] = [System.Collections.Generic.List[object]]::new() }
      $rRowsBySku[$opportunity.SkuId].Add($opportunity)
    }
    $rows = @(foreach ($skuId in ($capacity.Keys | Sort-Object)) { $capacity[$skuId] })
    foreach ($sku in $rows) {
      $skuId = [string]$sku.SkuId; $part = [string]$sku.SkuPartNumber; $product = [string]$sku.ReportProduct
      $excluded = Test-FinOpsTrue $sku.Excluded
      $enabled = [long](ConvertTo-FinOpsLong $sku.EnabledUnits); $consumed = [long](ConvertTo-FinOpsLong $sku.ConsumedUnits); $available = [long](ConvertTo-FinOpsLong $sku.AvailableUnits)
      $terms = Get-FinOpsContractTerms -Contracts $Contracts -SkuPartNumber $part -EnabledUnits $enabled -RunDate $RunStartUtc.Date
      if (-not $excluded -and $terms.ContractStatus -eq 'No contract' -and $enabled -gt 0) { $unmapped.Add($part) }
      $price = Get-FinOpsPrice $priceIndex $part
      $candidates = @($(if ($rRowsBySku.ContainsKey($skuId)) { $rRowsBySku[$skuId] }) | Sort-Object @{ Expression={ $script:CapRankByLever[$_.Lever] } }, @{ Expression={ $_.UserId.ToLowerInvariant() } })
      $capacityUnits = if ($excluded) { 0 } else { [math]::Max([long]0, [math]::Max([long]0, $available) - [long]$terms.ReserveUnits) }
      $pool = [long]$capacityUnits + $candidates.Count
      $wait = [long]$(if ($product -and $waitByProduct.ContainsKey($product) -and $skusByProduct.ContainsKey($product) -and $skusByProduct[$product].Count -eq 1) { $waitByProduct[$product] } else { 0 })
      $reused = [long][math]::Min([long]$pool, [long]$wait)
      $reducible = $pool - $reused
      $cap = if ($null -eq $terms.MinimumUnits) { $null } else { [math]::Max([long]0, $enabled - [long]$terms.MinimumUnits) }
      $realizable = if ($null -eq $cap) { $reducible } else { [math]::Min($reducible, [long]$cap) }
      $reusedFromCapacity = [math]::Min([long]$capacityUnits, $reused)
      $rFromCapacity = [math]::Min([long]$capacityUnits - $reusedFromCapacity, $realizable)
      $cappedFromCapacity = [long]$capacityUnits - $reusedFromCapacity - $rFromCapacity
      $reuseLeft = $reused - $reusedFromCapacity; $rLeft = $realizable - $rFromCapacity
      $rank = 0
      foreach ($candidate in $candidates) {
        $rank++; $candidate.CapRank = $rank
        $candidate.ContractKey = $terms.ContractKey; $candidate.RenewalDate = $terms.RenewalDate; $candidate.RealizableFrom = $terms.RealizableFrom
        if ($reuseLeft -gt 0) { $reuseLeft--; Clear-FinOpsOpportunityValue $candidate 'Reused' 'Unit reused for users waiting for capacity (see L8)'; continue }
        if ($rLeft -gt 0) { $rLeft--; continue }
        Clear-FinOpsOpportunityValue $candidate 'Capped' 'Beyond the contract reduction cap'
      }
      $skuCommon = @{ SkuId=$skuId; SkuPartNumber=$part; ReportProduct=$product; SourceFile='M365_Licenses_ReportSkuCapacity.csv'; ContractKey=$terms.ContractKey; RenewalDate=$terms.RenewalDate; RealizableFrom=$terms.RealizableFrom; CapRank=0 }
      if ($rFromCapacity -gt 0) {
        $o = New-FinOpsOpportunity ($skuCommon + @{ OpportunityKey="L7|$skuId|Realizable"; Lever='L7'; ValueClass='R'; Units=$rFromCapacity; SourceReason='Unassigned units after reserve' })
        Set-FinOpsOpportunityValue $o $price ($provisionalLevers.Contains('L7')) $null; $opportunities.Add($o)
      }
      if ($cappedFromCapacity -gt 0) {
        $o = New-FinOpsOpportunity ($skuCommon + @{ OpportunityKey="L7|$skuId|Capped"; Lever='L7'; ValueClass='R'; Units=$cappedFromCapacity; SourceReason='Unassigned units after reserve' })
        Clear-FinOpsOpportunityValue $o 'Capped' 'Beyond the contract reduction cap'; $opportunities.Add($o)
      }
      if ($reused -gt 0) {
        $reuseValues = $skuCommon.Clone(); $reuseValues.SourceFile = 'M365_Licenses_ReportSummary.csv'
        $o = New-FinOpsOpportunity ($reuseValues + @{ OpportunityKey="L8|$skuId|Reuse"; Lever='L8'; ValueClass='E'; Units=$reused; SourceReason=("{0} user(s) waiting for {1} capacity" -f $wait, $product) })
        Set-FinOpsOpportunityValue $o $price ($provisionalLevers.Contains('L8')) $null; $opportunities.Add($o)
      }
      if ($available -lt 0) {
        $o = New-FinOpsOpportunity ($skuCommon + @{ OpportunityKey="L9|$skuId|OverAssigned"; Lever='L9'; ValueClass='S'; Units=[math]::Abs($available) })
        Clear-FinOpsOpportunityValue $o 'Signal' 'Over-assigned'; $opportunities.Add($o)
      }
      if ([string]$sku.UsageEvidence -ceq 'Not measured') {
        $lever = if ($product -ceq 'Microsoft 365 Copilot') { 'L6' } else { 'L10' }
        $o = New-FinOpsOpportunity ($skuCommon + @{ OpportunityKey="$lever|$skuId|UsageNotMeasured"; Lever=$lever; ValueClass='S'; Units=[long](ConvertTo-FinOpsLong $sku.ActiveUsersUsageNotMeasured) })
        Clear-FinOpsOpportunityValue $o 'Not measured' ([string]$sku.UsageEvidenceReason); $opportunities.Add($o)
      }
      $skuValued = @($opportunities | Where-Object { $_.SkuId -ceq $skuId -and $_.ValueClass -eq 'R' -and $_.Lever -in @('L1','L2','L3','L5','L6','L7') -and $_.ValuationStatus -in $script:ValuedStatuses })
      $reviewRows = @($opportunities | Where-Object { $_.SkuId -ceq $skuId -and $_.SourceFile -eq 'M365_Licenses_ReportSkuCandidates.csv' })
      $skuRows.Add([pscustomobject][ordered]@{
        SkuId=$skuId; SkuPartNumber=$part; SkuDisplayName=[string]$sku.SkuDisplayName; ReportProduct=$product; Excluded=$excluded
        EnabledUnits=$enabled; ConsumedUnits=$consumed; AvailableUnits=$available; ReserveUnits=$terms.ReserveUnits; MinimumUnits=$(if ($null -eq $terms.MinimumUnits) { '' } else { $terms.MinimumUnits })
        CandidateUnitsR=$candidates.Count; PoolUnits=$pool; ReusedUnits=$reused; ReducibleUnits=$reducible; CapUnits=$(if ($null -eq $cap) { '' } else { $cap })
        RealizableUnits=$realizable; CapacityRealizableUnits=$rFromCapacity; CappedUnits=($pool - $reused - $realizable)
        ReviewUnits=@($reviewRows | Where-Object ValuationStatus -eq 'Review').Count; NotMeasuredUnits=@($reviewRows | Where-Object ValuationStatus -eq 'Not measured').Count
        BlockedUnits=@($reviewRows | Where-Object ValuationStatus -eq 'Blocked').Count; ReferencedUnits=@($reviewRows | Where-Object ValuationStatus -eq 'Referenced').Count
        UnitPriceMonthly=$(if ($price.Valued) { $price.Price } else { '' }); PriceRowKey=$price.RowKey; PriceIsEstimated=$price.Estimated
        RealizableAnnualAmount=$(if ($price.Valued) { [decimal](@($skuValued | ForEach-Object { [decimal]$_.AnnualAmount }) | Measure-Object -Sum).Sum } else { '' })
        AvoidedAnnualAmount=$(if ($price.Valued -and $reused -gt 0) { [decimal]($price.Price * $reused * 12) } else { '' })
        ValuationStatus=$(if (-not $price.Valued) { 'Not valued' } elseif (@($opportunities | Where-Object { $_.SkuId -ceq $skuId -and $_.ValuationStatus -eq 'Provisional' }).Count) { 'Provisional' } else { 'Valued' })
        ContractKey=$terms.ContractKey; ContractStatus=$terms.ContractStatus; ReductionPolicy=$terms.ReductionPolicy; RenewalDate=$terms.RenewalDate; RealizableFrom=$terms.RealizableFrom
        UsageEvidence=[string]$sku.UsageEvidence
      })
    }

    # L11 assignment issues and L12 gaps: signals only.
    $issues = @{}
    foreach ($row in $Sources['M365_Licenses_ReportAssignmentIssues.csv'].Rows) {
      $key = '{0}|{1}' -f $row.Product, $row.Category
      if (-not $issues.ContainsKey($key)) { $issues[$key] = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) }
      [void]$issues[$key].Add([string]$row.UserId)
    }
    foreach ($key in $issues.Keys) {
      $parts = $key.Split('|', 2)
      $o = New-FinOpsOpportunity @{ OpportunityKey="L11|$key"; Lever='L11'; ValueClass='S'; ReportProduct=$parts[0]; SourceFile='M365_Licenses_ReportAssignmentIssues.csv'; SourceCandidateType=$parts[1]; Units=$issues[$key].Count; CapRank=0 }
      Clear-FinOpsOpportunityValue $o 'Signal' $parts[1]; $opportunities.Add($o)
    }
    foreach ($row in $Sources['M365_Licenses_ReportGaps.csv'].Rows) {
      $value = ConvertTo-FinOpsLong $row.Value
      $o = New-FinOpsOpportunity @{ OpportunityKey=('L12|{0}|{1}' -f $row.Section, $row.Metric); Lever='L12'; ValueClass='S'; SourceFile='M365_Licenses_ReportGaps.csv'; SourceCandidateType=[string]$row.Section; SourceReason=[string]$row.Metric; Units=$(if ($null -eq $value) { '' } else { $value }); CapRank=0 }
      if ($null -eq $value) { Clear-FinOpsOpportunityValue $o 'Not measured' 'N/D in the Licences report' } else { Clear-FinOpsOpportunityValue $o 'Signal' ([string]$row.EvidenceStatus) }
      $opportunities.Add($o)
    }
  }

  # L13 ESU: no SmartInventory decision defines the device population yet.
  $esu = New-FinOpsOpportunity @{ OpportunityKey='L13|WIN10_ESU'; Lever='L13'; ValueClass='E'; SourceFile='Windows_ESU_Prices.csv'; Units=''; CapRank=0 }
  Clear-FinOpsOpportunityValue $esu 'Not measured' $(if ($Sources['Windows_ESU_Prices.csv'].QualificationStatus -eq 'NotMeasured') { "No ESU device population published by SmartInventory; ESU prices: $($Sources['Windows_ESU_Prices.csv'].Reason)" } else { 'No ESU device population published by SmartInventory' })
  $opportunities.Add($esu)

  # Showback join (allocation only): Entra object -> immutable ID -> AD, else unique UPN.
  $entra = $Sources['M365_Users_Active.csv']; $ad = $Sources['AD_Users_AllDomains.csv']
  $entraById = @{}
  if ($entra.QualificationStatus -ne 'NotMeasured') { foreach ($row in $entra.Rows) { $entraById[([string]$row.'Object Id').ToLowerInvariant()] = $row } }
  $adByImmutable = @{}; $adByUpn = @{}
  if ($ad.QualificationStatus -ne 'NotMeasured') {
    foreach ($row in $ad.Rows) {
      $immutable = ([string]$row.ImmutableId_AD).Trim().ToLowerInvariant(); $upn = ([string]$row.UserPrincipalName).Trim().ToLowerInvariant()
      if ($immutable) { $adByImmutable[$immutable] = if ($adByImmutable.ContainsKey($immutable)) { $null } else { $row } }
      if ($upn) { $adByUpn[$upn] = if ($adByUpn.ContainsKey($upn)) { $null } else { $row } }
    }
  }
  $orEmpty = { param($value) $text = ([string]$value).Trim(); if ($text) { $text } else { 'Not attributed' } }
  foreach ($o in $opportunities) {
    if (-not $o.UserId) { $o.AdDomain = 'SKU level'; $o.Country = 'SKU level'; $o.Entity = 'SKU level'; $o.Department = 'SKU level'; $o.ShowbackJoinStatus = 'SKU level'; continue }
    $entraRow = $entraById[$o.UserId.ToLowerInvariant()]
    $o.Department = if ($entra.QualificationStatus -eq 'NotMeasured') { 'Not measured' } elseif ($entraRow) { & $orEmpty $entraRow.Department } else { 'Not attributed' }
    if ($ad.QualificationStatus -eq 'NotMeasured') { $o.AdDomain = 'Not measured'; $o.Country = 'Not measured'; $o.Entity = 'Not measured'; $o.ShowbackJoinStatus = 'AD not measured'; continue }
    $adRow = $null; $join = 'Not joined'
    if ($entraRow) {
      $immutable = ([string]$entraRow.OnPremisesImmutableId).Trim().ToLowerInvariant()
      if ($immutable -and $adByImmutable.ContainsKey($immutable) -and $adByImmutable[$immutable]) { $adRow = $adByImmutable[$immutable]; $join = 'ImmutableId' }
    }
    if (-not $adRow) {
      $upn = ([string]$o.UserPrincipalName).Trim().ToLowerInvariant()
      if ($upn -and $adByUpn.ContainsKey($upn) -and $adByUpn[$upn]) { $adRow = $adByUpn[$upn]; $join = 'UserPrincipalName' }
    }
    $o.ShowbackJoinStatus = $join
    if ($adRow) { $o.AdDomain = & $orEmpty $adRow.DomainName; $o.Country = & $orEmpty $adRow.Country; $o.Entity = & $orEmpty $adRow.EntityLabel }
    else {
      # With partial AD coverage an unmatched user may sit in a domain that was not collected.
      $fallback = if ($ad.QualificationStatus -eq 'Provisional') { 'Not measured' } else { 'Not attributed' }
      $o.AdDomain = $fallback; $o.Country = $fallback; $o.Entity = $fallback
    }
  }
  return [pscustomobject]@{
    Measured=$measured; SnapshotId=$snapshotId; ReceiptRunId=$receiptRunId; AsOfDate=$asOfDate; Provisional=$provisional; ProvisionalReasons=$provisionalReasons.ToArray(); ProvisionalLevers=@($provisionalLevers | Sort-Object); ProvisionalTrace=$provisionalTrace.ToArray()
    MissingDomains=$missingDomains; Opportunities=$opportunities.ToArray(); SkuRows=$skuRows.ToArray(); Unmapped=@($unmapped | Sort-Object -Unique); PriceIndex=$priceIndex
  }
}

function Get-FinOpsAggregate {
  param([AllowEmptyCollection()][object[]]$Rows)
  $units = [long]0; $hasUnits = $false; $amount = [decimal]0; $hasAmount = $false
  $users = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach ($row in $Rows) {
    # L12 gap metrics are heterogeneous: never summed.
    if ($row.Lever -ne 'L12') { $u = ConvertTo-FinOpsLong $row.Units; if ($null -ne $u) { $units += $u; $hasUnits = $true } }
    if ($row.UserId) { [void]$users.Add($row.UserId) }
    if ($null -ne $row.AnnualAmount -and [string]$row.AnnualAmount -ne '') { $amount += [decimal]$row.AnnualAmount; $hasAmount = $true }
  }
  return [pscustomobject]@{ Users=$users.Count; Units=$(if ($hasUnits) { $units } else { '' }); Amount=$(if ($hasAmount) { $amount } else { '' }) }
}

function Get-FinOpsShowback {
  param([AllowEmptyCollection()][object[]]$Opportunities)
  $result = [System.Collections.Generic.List[object]]::new()
  foreach ($dimension in @('AdDomain','Country','Entity','Department')) {
    foreach ($group in @($Opportunities | Group-Object -Property @{ Expression={ '{0}|{1}|{2}|{3}' -f $_.$dimension, $_.Lever, $_.ValueClass, $_.ValuationStatus } })) {
      $first = $group.Group[0]
      $aggregate = Get-FinOpsAggregate $group.Group
      $result.Add([pscustomobject][ordered]@{ Dimension=$dimension; DimensionValue=$first.$dimension; Lever=$first.Lever; ValueClass=$first.ValueClass; ValuationStatus=$first.ValuationStatus
        Users=$aggregate.Users; Units=$aggregate.Units; AnnualAmount=$aggregate.Amount; Currency='EUR'; Provisional=($first.ValuationStatus -eq 'Provisional') })
    }
  }
  return $result.ToArray()
}

function Get-FinOpsSummary {
  param([Parameter(Mandatory)]$Valuation)
  $rows = [System.Collections.Generic.List[object]]::new()
  $common = @{ Currency='EUR'; MissingDomains=$Valuation.MissingDomains
    SnapshotGeneratedAtUtc=$(if ($Valuation.Measured) { $Valuation.AsOfDate.ToString('yyyy-MM-dd') } else { '' }); PricingReceiptRunId=$Valuation.PriceIndex.RunId }
  $add = {
    param([string]$Key, [string]$Class, [string]$Lever, [string]$Status, [AllowEmptyCollection()][object[]]$Items, [string]$Note)
    $aggregate = Get-FinOpsAggregate $Items
    if ($Status -eq 'Auto') { $Status = if (@($Items | Where-Object { $null -ne $_ -and $_.ValuationStatus -eq 'Provisional' }).Count) { 'Provisional' } else { 'Valued' } }
    $values = [ordered]@{ KpiKey=$Key; ValueClass=$Class; Lever=$Lever; ValuationStatus=$Status; Users=$aggregate.Users; Units=$aggregate.Units; AnnualAmount=$aggregate.Amount }
    foreach ($k in $common.Keys) { $values[$k] = $common[$k] }
    $values.Provisional = @($Items | Where-Object { $null -ne $_ -and $_.ValuationStatus -eq 'Provisional' }).Count -gt 0
    $values.Note = $Note
    $rows.Add([pscustomobject]$values)
  }
  if (-not $Valuation.Measured) {
    foreach ($key in @('TOTAL_R','TOTAL_R_DISABLED_AND_SHARED','TOTAL_R_ENABLED_INACTIVE','TOTAL_R_COPILOT','TOTAL_R_CAPACITY','TOTAL_RC_E3_TO_F3','TOTAL_E_REUSE')) {
      $values = [ordered]@{ KpiKey=$key; ValueClass=''; Lever=''; ValuationStatus='Not measured'; Users=''; Units=''; AnnualAmount='' }
      foreach ($k in $common.Keys) { $values[$k] = $common[$k] }
      $values.Provisional = $false
      $values.Note = 'Licences report not qualified'
      $rows.Add([pscustomobject]$values)
    }
    & $add 'TOTAL_E_ESU' 'E' 'L13' 'Not measured' @($Valuation.Opportunities | Where-Object Lever -eq 'L13') 'No ESU device population published by SmartInventory'
    return $rows.ToArray()
  }
  $opps = $Valuation.Opportunities
  $valued = { param([string[]]$Levers, [string]$Class) @($opps | Where-Object { $_.Lever -in $Levers -and $_.ValueClass -eq $Class -and $_.ValuationStatus -in $script:ValuedStatuses }) }
  $status = 'Auto'
  & $add 'TOTAL_R' 'R' 'L1;L2;L3;L5;L6;L7' $status (& $valued @('L1','L2','L3','L5','L6','L7') 'R') 'Realizable at contract or renewal; excludes conditional E3 to F3'
  & $add 'TOTAL_R_DISABLED_AND_SHARED' 'R' 'L1;L2;L5' $status (& $valued @('L1','L2','L5') 'R') 'Disabled accounts and eligible shared mailboxes'
  & $add 'TOTAL_R_ENABLED_INACTIVE' 'R' 'L3' $status (& $valued @('L3') 'R') 'R - enabled accounts without M365 activity in 90 days'
  & $add 'TOTAL_R_COPILOT' 'R' 'L6' $status (& $valued @('L6') 'R') 'Copilot without licensed activity'
  & $add 'TOTAL_R_CAPACITY' 'R' 'L7' $status (& $valued @('L7') 'R') 'Unassigned units after reserve and cap'
  & $add 'TOTAL_RC_E3_TO_F3' 'Rc' 'L4' $status (& $valued @('L4') 'Rc') 'Conditional: outside TOTAL_R'
  & $add 'TOTAL_E_REUSE' 'E' 'L8' $status (& $valued @('L8') 'E') 'Purchase avoided by reusing units for users waiting for capacity'
  & $add 'TOTAL_E_ESU' 'E' 'L13' 'Not measured' @($opps | Where-Object Lever -eq 'L13') 'No ESU device population published by SmartInventory'
  & $add 'NOT_MEASURED_MAILBOX' 'S' 'L5' 'Not measured' @($opps | Where-Object { $_.Lever -eq 'L5' -and $_.ValuationStatus -eq 'Not measured' }) 'Mailbox state unknown during the Licences run (expected only when EXO failed)'
  foreach ($group in @($opps | Group-Object -Property @{ Expression={ '{0}|{1}|{2}' -f $_.ValueClass, $_.Lever, $_.ValuationStatus } })) {
    $first = $group.Group[0]
    & $add ('DETAIL|' + $group.Name) $first.ValueClass $first.Lever $first.ValuationStatus $group.Group ''
  }
  # One row per gap metric, value as published, never summed.
  foreach ($gap in @($opps | Where-Object Lever -eq 'L12')) {
    $values = [ordered]@{ KpiKey=('GAP|{0}|{1}' -f $gap.SourceCandidateType, $gap.SourceReason); ValueClass='S'; Lever='L12'; ValuationStatus=$gap.ValuationStatus; Users=''; Units=$gap.Units; AnnualAmount='' }
    foreach ($k in $common.Keys) { $values[$k] = $common[$k] }
    $values.Provisional = $false
    $values.Note = 'Licences report gap metric'
    $rows.Add([pscustomobject]$values)
  }
  return $rows.ToArray()
}

function Get-FinOpsControls {
  param([Parameter(Mandatory)][hashtable]$Sources, [Parameter(Mandatory)]$Valuation, [AllowEmptyCollection()][object[]]$Showback)
  $controls = [System.Collections.Generic.List[object]]::new()
  $addCheck = { param([string]$Control, [string]$Key, $SourceCount, $FinOpsCount, [string]$Detail)
    $controls.Add([pscustomobject][ordered]@{ Control=$Control; Key=$Key; SourceCount=[string]$SourceCount; FinOpsCount=[string]$FinOpsCount; Status=$(if ([string]$SourceCount -ceq [string]$FinOpsCount) { 'Pass' } else { 'Fail' }); Detail=$Detail }) }
  if (-not $Valuation.Measured) {
    $controls.Add([pscustomobject][ordered]@{ Control='G0'; Key='Licences'; SourceCount=''; FinOpsCount=''; Status='Not applicable'; Detail='Licences report not qualified: nothing valued' })
    return $controls.ToArray()
  }
  $opps = $Valuation.Opportunities
  $compare = {
    param([string]$Control, [hashtable]$Source, [hashtable]$FinOps, [string]$Detail)
    foreach ($key in @($Source.Keys + $FinOps.Keys | Sort-Object -Unique)) { & $addCheck $Control $key $(if ($Source.ContainsKey($key)) { $Source[$key] } else { 0 }) $(if ($FinOps.ContainsKey($key)) { $FinOps[$key] } else { 0 }) $Detail }
  }
  $countBy = { param([object[]]$Items, [scriptblock]$Key) $table = @{}; foreach ($row in $Items) { $k = & $Key $row; $table[$k] = 1 + $(if ($table.ContainsKey($k)) { $table[$k] } else { 0 }) }; $table }
  # G1 ReportCandidates CandidateType x Product.
  & $compare 'G1' (& $countBy @($Sources['M365_Licenses_ReportCandidates.csv'].Rows) { param($r) '{0}|{1}' -f $r.CandidateType, $r.Product }) (& $countBy @($opps | Where-Object SourceFile -eq 'M365_Licenses_ReportCandidates.csv') { param($r) '{0}|{1}' -f $r.SourceCandidateType, $r.ReportProduct }) 'ReportCandidates rows by CandidateType x Product'
  # G2 ReportSkuCandidates CandidateType x SKU x Recommendation.
  $skuSource = @($Sources['M365_Licenses_ReportSkuCandidates.csv'].Rows)
  & $compare 'G2' (& $countBy $skuSource { param($r) '{0}|{1}|{2}' -f $r.CandidateType, $r.SkuPartNumber, $r.Recommendation }) (& $countBy @($opps | Where-Object SourceFile -eq 'M365_Licenses_ReportSkuCandidates.csv') { param($r) '{0}|{1}|{2}' -f $r.SourceCandidateType, $r.SkuPartNumber, $r.SourceRecommendation }) 'ReportSkuCandidates rows by CandidateType x SKU x Recommendation'
  # G3 Summary RecoveryCandidates per product equals Recovery rows (source coherence).
  $summary = @{}; foreach ($row in $Sources['M365_Licenses_ReportSummary.csv'].Rows) { $n = ConvertTo-FinOpsLong $row.RecoveryCandidates; if ($null -ne $n) { $summary[[string]$row.Product] = $n } }
  $recovery = & $countBy @($Sources['M365_Licenses_ReportCandidates.csv'].Rows | Where-Object CandidateType -ceq 'Recovery') { param($r) [string]$r.Product }
  foreach ($product in $summary.Keys) { if (-not $recovery.ContainsKey($product)) { $recovery[$product] = 0 } }
  & $compare 'G3' $summary $recovery 'ReportSummary RecoveryCandidates vs Recovery rows'
  # G4 ReportSkuCapacity counters per SKU equal the per-SKU candidate rows.
  $capacityCounts = @{}
  foreach ($row in $Sources['M365_Licenses_ReportSkuCapacity.csv'].Rows) {
    foreach ($pair in @(@('Recommended','RecommendedCandidates'),@('Review','ReviewCandidates'),@('Blocked by dependency','BlockedCandidates'),@('Referenced','ReferencedCandidates'))) {
      $n = [long](ConvertTo-FinOpsLong (Get-FinOpsProperty $row $pair[1])); if ($n -ne 0) { $capacityCounts['{0}|{1}' -f $row.SkuId, $pair[0]] = $n }
    }
  }
  & $compare 'G4' $capacityCounts (& $countBy $skuSource { param($r) '{0}|{1}' -f $r.SkuId, $r.Recommendation }) 'ReportSkuCapacity counters vs ReportSkuCandidates'
  # G5 every Referenced row has its ReportCandidates target; at most one valued row per user and SKU.
  $targets = @{}
  foreach ($row in $Sources['M365_Licenses_ReportCandidates.csv'].Rows) { $targets[('{0}|ReportCandidates:{1}|{2}' -f ([string]$row.UserId).ToLowerInvariant(), $row.CandidateType, $row.Product)] = $true }
  $missingTargets = @($opps | Where-Object { $_.ValuationStatus -eq 'Referenced' -and -not $targets.ContainsKey(('{0}|{1}|{2}' -f $_.UserId.ToLowerInvariant(), $_.CountedIn, $_.ReportProduct)) })
  & $addCheck 'G5' 'ReferencedTargets' 0 $missingTargets.Count 'Referenced rows without a ReportCandidates target'
  $valuedPairs = @($opps | Where-Object { $_.UserId -and $_.SkuId -and $_.ValuationStatus -in $script:ValuedStatuses } | Group-Object -Property @{ Expression={ '{0}|{1}' -f $_.UserId.ToLowerInvariant(), $_.SkuId.ToLowerInvariant() } } | Where-Object Count -gt 1)
  & $addCheck 'G5' 'ValuedPairs' 0 $valuedPairs.Count 'User x SKU pairs valued more than once'
  $keys = @($opps | Group-Object OpportunityKey | Where-Object Count -gt 1)
  & $addCheck 'G5' 'OpportunityKeys' 0 $keys.Count 'Duplicate OpportunityKey'
  # G6 showback per dimension reconciles with the opportunity totals.
  $total = @{}
  foreach ($group in @($opps | Group-Object -Property @{ Expression={ '{0}|{1}|{2}' -f $_.Lever, $_.ValueClass, $_.ValuationStatus } })) {
    $aggregate = Get-FinOpsAggregate $group.Group; $total[$group.Name] = '{0}|{1}' -f $aggregate.Units, (Format-FinOpsDecimal $(if ([string]$aggregate.Amount -eq '') { $null } else { $aggregate.Amount }))
  }
  foreach ($dimension in @('AdDomain','Country','Entity','Department')) {
    $side = @{}
    foreach ($group in @($Showback | Where-Object Dimension -eq $dimension | Group-Object -Property @{ Expression={ '{0}|{1}|{2}' -f $_.Lever, $_.ValueClass, $_.ValuationStatus } })) {
      $unitValues = @($group.Group | Where-Object { [string]$_.Units -ne '' } | ForEach-Object { [long]$_.Units })
      $units = if ($unitValues.Count) { [long]($unitValues | Measure-Object -Sum).Sum } else { '' }
      $amounts = @($group.Group | Where-Object { $null -ne $_.AnnualAmount -and [string]$_.AnnualAmount -ne '' } | ForEach-Object { [decimal]$_.AnnualAmount })
      $side[$group.Name] = '{0}|{1}' -f $units, (Format-FinOpsDecimal $(if ($amounts.Count) { [decimal]($amounts | Measure-Object -Sum).Sum } else { $null }))
    }
    foreach ($key in @($total.Keys + $side.Keys | Sort-Object -Unique)) { & $addCheck 'G6' "$dimension|$key" $total[$key] $side[$key] 'Showback units|amount vs opportunity totals' }
  }
  return $controls.ToArray()
}
#endregion

#region Main
if ($MyInvocation.InvocationName -eq '.') { return }

$exitCode = 0
$runStart = if ($AsOfUtc) { ConvertTo-FinOpsUtc $AsOfUtc } else { [datetime]::UtcNow }
if ($null -eq $runStart) { throw "Invalid -AsOfUtc value: $AsOfUtc" }
$runId = '{0}-{1}' -f $runStart.ToString('yyyyMMddTHHmmssfffZ', $script:Invariant), [guid]::NewGuid().ToString('N').Substring(0, 8)
$currentRoot = $null; $archiveDir = $null
try {
  Write-FinOpsLog ("SmartFinOps Workplace Valuation {0} - tenant '{1}' - run {2}{3}" -f $script:ScriptVersion, $Tenant, $runId, $(if ($ValidateOnly) { ' - ValidateOnly (no file written)' } else { '' }))
  . (Join-Path $PSScriptRoot 'Config\SmartFinOps-TenantContext.ps1')
  $config = Initialize-SmartFinOpsTenantContext -Tenant $Tenant -StartPath $PSScriptRoot
  $repositoryRoot = Split-Path -Path $PSScriptRoot -Parent
  $coreManifest = Join-Path $repositoryRoot 'SmartM365\Modules\SmartM365.Core\SmartM365.Core.psd1'
  Import-Module $coreManifest -MinimumVersion '1.0.90' -ErrorAction Stop | Out-Null

  $dataLast = if ($SmartM365LatestCsvFolderPath) { $SmartM365LatestCsvFolderPath } else { [string](Get-SmartFinOpsConfigValue -Name 'SmartM365LatestCsvFolderPath') }
  if (-not $dataLast -or -not (Test-Path -LiteralPath $dataLast -PathType Container)) { throw "SmartM365 DATA-LAST folder not found: $dataLast" }
  $currentRoot = if ($OutputRoot) { $OutputRoot } else { [string](Resolve-SmartFinOpsConfigTokenValue -Value '{{TenantDataRootPath}}') }
  $archiveRoot = Join-Path $currentRoot 'Archive'
  $valuation = Get-SmartFinOpsConfigValue -Name 'Valuation'
  $dailyHours = [int]$(if ($null -ne (Get-FinOpsProperty $valuation 'FreshnessDailyHours')) { Get-FinOpsProperty $valuation 'FreshnessDailyHours' } else { 48 })
  $weeklyDays = [int]$(if ($null -ne (Get-FinOpsProperty $valuation 'FreshnessWeeklyDays')) { Get-FinOpsProperty $valuation 'FreshnessWeeklyDays' } else { 8 })
  $keepRuns = [int]$(if ($null -ne (Get-FinOpsProperty $valuation 'ArchiveRetentionRuns')) { Get-FinOpsProperty $valuation 'ArchiveRetentionRuns' } else { 30 })
  if ($dailyHours -lt 1 -or $weeklyDays -lt 1 -or $keepRuns -lt 1) { throw 'Valuation freshness and retention settings must be positive.' }
  $contractsFile = if ($ContractsPath) { $ContractsPath } else {
    $configured = [string](Get-FinOpsProperty $valuation 'ContractsPath')
    if ($configured) { [string](Resolve-SmartFinOpsConfigTokenValue -Value $configured) } else { Join-Path $PSScriptRoot 'Config\SmartFinOps-Contracts.local.json' }
  }
  Write-FinOpsLog "Sources: $dataLast"
  Write-FinOpsLog "Outputs: $currentRoot"
  $contracts = Get-FinOpsContracts -Path $contractsFile -TemplatePath (Join-Path $PSScriptRoot 'Config\SmartFinOps-Contracts.local.json.template')

  $licReceipt = 'SmartInventory_SmartM365-Licences-Inventory.current.json.txt'; $licProducer = 'SmartM365-Licences-Inventory.ps1'
  $specs = @(
    @{ File='M365_Licenses_ReportSummary.csv'; Receipt=$licReceipt; Producer=$licProducer; Required=$true; Columns=@('SnapshotId','Product','Category','AssignmentCapacityWait','RecoveryCandidates','GeneratedAtUtc') }
    @{ File='M365_Licenses_ReportCandidates.csv'; Receipt=$licReceipt; Producer=$licProducer; Required=$true; Columns=@('SnapshotId','CandidateType','Product','UserId','UserPrincipalName','Reason') }
    @{ File='M365_Licenses_ReportAssignmentIssues.csv'; Receipt=$licReceipt; Producer=$licProducer; Required=$true; Columns=@('SnapshotId','Product','Category','UserId') }
    @{ File='M365_Licenses_ReportGaps.csv'; Receipt=$licReceipt; Producer=$licProducer; Required=$true; Columns=@('SnapshotId','Section','Metric','Value','EvidenceStatus') }
    @{ File='M365_Licenses_ReportSources.csv'; Receipt=$licReceipt; Producer=$licProducer; Required=$true; Columns=@('SnapshotId','Name','Ready','Provisional','Reason') }
    @{ File='M365_Licenses_ReportSkuCandidates.csv'; Receipt=$licReceipt; Producer=$licProducer; Required=$true; Columns=@('SnapshotId','ReceiptRunId','CandidateType','Recommendation','Reason','SkuId','SkuPartNumber','UserId','UserPrincipalName','MailboxGuard','DependencyStatus','CountedIn') }
    @{ File='M365_Licenses_ReportSkuCapacity.csv'; Receipt=$licReceipt; Producer=$licProducer; Required=$true; Columns=@('SnapshotId','ReceiptRunId','SkuId','SkuPartNumber','SkuDisplayName','ReportProduct','Excluded','EnabledUnits','ConsumedUnits','AvailableUnits','AssignedUsers','ActiveUsersUsageNotMeasured','RecommendedCandidates','ReviewCandidates','BlockedCandidates','ReferencedCandidates','UsageEvidence','UsageEvidenceReason') }
    @{ File='M365_Licenses_Users.csv'; Receipt=$licReceipt; Producer=$licProducer; Required=$true; Columns=@('UserId','SkuId') }
    @{ File='M365_License_Prices.csv'; Receipt='SmartInventory_SmartM365-LicensePricing-Inventory.current.json.txt'; Producer='SmartM365-LicensePricing-Inventory.ps1'; Required=$true; Columns=@('SkuPartNumber','Currency','UnitPriceMonthly','EffectiveFrom','EffectiveTo','PriceSource','IsEstimated') }
    @{ File='Windows_ESU_Prices.csv'; Receipt='SmartInventory_SmartM365-LicensePricing-Inventory.current.json.txt'; Producer='SmartM365-LicensePricing-Inventory.ps1'; Required=$false; Columns=@('ProgramKey','CoverageYear','Currency','UnitPricePerDeviceYear') }
    @{ File='AD_Users_AllDomains.csv'; Receipt='SmartInventory_SmartM365-ActiveDirectory-Inventory.current.json.txt'; Producer='SmartM365-ActiveDirectory-Inventory.ps1'; Required=$false; Columns=@('DomainName','ImmutableId_AD','UserPrincipalName','Country','EntityLabel'); PartialAd=$true }
    @{ File='M365_Users_Active.csv'; Receipt='SmartInventory_SmartM365-ActiveUsers-Inventory.current.json.txt'; Producer='SmartM365-ActiveUsers-Inventory.ps1'; Required=$false; Columns=@('Object Id','User principal name','OnPremisesImmutableId','Department') }
  )
  $sources = @{}
  foreach ($spec in $specs) {
    $record = Get-FinOpsSource -Folder $dataLast -FileName $spec.File -ReceiptFile $spec.Receipt -Producer $spec.Producer -TenantKey $Tenant -RunStartUtc $runStart -FreshnessHours $dailyHours -RequiredColumns $spec.Columns -Required:$spec.Required -AllowPartialAd:([bool]$spec['PartialAd'])
    $sources[$spec.File] = $record
    Write-FinOpsLog ('Source {0}: {1}{2}' -f $spec.File, $record.QualificationStatus, $(if ($record.Reason) { " ($($record.Reason))" } else { '' })) $(if ($record.QualificationStatus -eq 'NotMeasured' -and $spec.Required) { 'WARNING' } else { 'INFO' })
  }

  # Licences report coherence: one snapshot, one receipt run, one tenant.
  $licFiles = @($specs | Where-Object { $_.Receipt -eq $licReceipt -and $_.File -ne 'M365_Licenses_Users.csv' } | ForEach-Object { $_.File })
  $snapshotPath = Join-Path $dataLast 'M365_Licenses_ReportSnapshot.json.txt'
  $snapshotRecord = New-FinOpsSourceRecord -FileName 'M365_Licenses_ReportSnapshot.json.txt' -Producer $licProducer -ReceiptFile $licReceipt -Required $true -FreshnessHours $dailyHours
  $snapshot = $null
  try {
    if (@($licFiles | Where-Object { $sources[$_].QualificationStatus -eq 'NotMeasured' }).Count) { throw 'Licences report files are not all qualified' }
    if (-not (Test-Path -LiteralPath $snapshotPath -PathType Leaf)) { throw 'snapshot is missing' }
    $snapshotRecord.ActualSHA256 = Get-FinOpsSha256 $snapshotPath
    $snapshot = Read-FinOpsLicencesSnapshot -Path $snapshotPath
    $receiptRun = $sources['M365_Licenses_ReportSummary.csv'].SourceReceiptRunId
    $snapshotRecord.SourceReceiptRunId = $receiptRun; $snapshotRecord.ReceiptStatus = 'Completed'; $snapshotRecord.PublishedInReceipt = $true
    $snapshotRecord.ReceiptCompletedAtUtc = $sources['M365_Licenses_ReportSummary.csv'].ReceiptCompletedAtUtc; $snapshotRecord.AgeHours = $sources['M365_Licenses_ReportSummary.csv'].AgeHours
    if ($snapshot.SchemaVersion -ne 2) { throw "unsupported snapshot SchemaVersion $($snapshot.SchemaVersion)" }
    if ($snapshot.TenantKey -cne $Tenant) { throw 'snapshot tenant differs' }
    if (-not $snapshot.SnapshotId) { throw 'snapshot ID is missing' }
    if ($snapshot.ReceiptRunId -cne $receiptRun) { throw 'snapshot ReceiptRun.RunId differs from the Licences receipt RunId' }
    if ($null -eq (ConvertTo-FinOpsUtc $snapshot.GeneratedAtUtc)) { throw 'snapshot GeneratedAtUtc is invalid' }
    foreach ($file in $licFiles) {
      $ids = @($sources[$file].Rows | ForEach-Object { [string]$_.SnapshotId } | Sort-Object -Unique)
      if ($ids.Count -gt 1 -or ($ids.Count -eq 1 -and $ids[0] -cne $snapshot.SnapshotId)) { throw "mixed or foreign SnapshotId in $file" }
      if ($file -like '*Sku*') { if (@($sources[$file].Rows | Where-Object { [string]$_.ReceiptRunId -cne $receiptRun }).Count) { throw "ReceiptRunId differs from the Licences receipt in $file" } }
    }
    $snapshotRecord.QualificationStatus = 'Qualified'
  }
  catch {
    $snapshotRecord.QualificationStatus = 'NotMeasured'; $snapshotRecord.Reason = $_.Exception.Message
    foreach ($file in $licFiles) { if ($sources[$file].QualificationStatus -ne 'NotMeasured') { Set-FinOpsSourceNotMeasured $sources[$file] "Licences report not coherent: $($_.Exception.Message)" } }
    Write-FinOpsLog "Licences report not measured: $($_.Exception.Message)" 'WARNING'
    $snapshot = $null
  }
  $sources[$snapshotRecord.SourceFile] = $snapshotRecord

  $archiveRuns = Get-FinOpsArchiveRuns -ArchiveRoot $archiveRoot -TenantKey $Tenant
  $result = Invoke-FinOpsValuation -Sources $sources -RunStartUtc $runStart -Contracts $contracts -Snapshot $snapshot
  $realization = if ($result.Measured) { Get-FinOpsRealization -SkuRows $result.SkuRows -ArchiveRuns $archiveRuns -UsersSource $sources['M365_Licenses_Users.csv'] } else { [pscustomobject]@{ Rows=@(); ProtectedRuns=[System.Collections.Generic.HashSet[string]]::new() } }
  foreach ($trace in $result.ProvisionalTrace) {
    if ($sources.ContainsKey($trace.Source)) { $sources[$trace.Source].ProvisionalLevers = $trace.Levers; continue }
    $record = New-FinOpsSourceRecord -FileName $trace.Source -Producer $licProducer -ReceiptFile 'M365_Licenses_ReportSources.csv' -Required $false -FreshnessHours $dailyHours
    $record.SourceReceiptRunId = $result.ReceiptRunId; $record.QualificationStatus = 'Provisional'; $record.ProvisionalLevers = $trace.Levers
    $record.Reason = ('Provisional in the Licences report{0}' -f $(if ($trace.Reason) { ": $($trace.Reason)" } else { '' }))
    $sources[$trace.Source] = $record
  }
  if ($sources['AD_Users_AllDomains.csv'].QualificationStatus -eq 'Provisional' -and -not $sources['AD_Users_AllDomains.csv'].ProvisionalLevers) { $sources['AD_Users_AllDomains.csv'].ProvisionalLevers = 'Showback' }
  $showback = Get-FinOpsShowback -Opportunities $result.Opportunities
  $summary = Get-FinOpsSummary -Valuation $result
  $controls = Get-FinOpsControls -Sources $sources -Valuation $result -Showback $showback
  $failed = @($controls | Where-Object Status -eq 'Fail')
  foreach ($control in $failed) { Write-FinOpsLog ('Control {0} failed: {1} source={2} finops={3}' -f $control.Control, $control.Key, $control.SourceCount, $control.FinOpsCount) 'ERROR' }
  $notMeasuredMailbox = @($result.Opportunities | Where-Object { $_.Lever -eq 'L5' -and $_.ValuationStatus -eq 'Not measured' }).Count
  if ($notMeasuredMailbox) { Write-FinOpsLog "$notMeasuredMailbox disabled-account license(s) not measured: mailbox state unknown during the Licences run (EXO not qualified)." 'WARNING' }
  if ($result.Provisional) { Write-FinOpsLog ('Provisional amounts: {0}' -f ($result.ProvisionalReasons -join ' | ')) 'WARNING' }
  $totalR = @($summary | Where-Object KpiKey -eq 'TOTAL_R')
  Write-FinOpsLog ('Licences measured: {0}; opportunities: {1}; TOTAL_R annual: {2} EUR; controls failed: {3}' -f $result.Measured, $result.Opportunities.Count, $(if ($totalR.Count -and [string]$totalR[0].AnnualAmount -ne '') { Format-FinOpsDecimal $totalR[0].AnnualAmount } else { 'n/a' }), $failed.Count)

  $common = [ordered]@{ TenantKey=$Tenant; FinOpsRunId=$runId; SnapshotId=$result.SnapshotId; ReceiptRunId=$result.ReceiptRunId }
  $outputs = [ordered]@{
    'FinOps_Sources.csv' = @($sources.Values)
    'FinOps_Opportunities.csv' = $result.Opportunities
    'FinOps_SkuValuation.csv' = $result.SkuRows
    'FinOps_Showback.csv' = $showback
    'FinOps_Summary.csv' = $summary
    'FinOps_Realization.csv' = $realization.Rows
    'FinOps_Controls.csv' = $controls
  }
  $status = if ($failed.Count) { 'Failed' } else { 'Completed' }
  if ($failed.Count) { $exitCode = 1 }
  if ($ValidateOnly) {
    Write-FinOpsLog "ValidateOnly: run status would be $status; no file written."
  }
  else {
    Save-FinOpsContracts -Contracts $contracts -UnmappedSkus $result.Unmapped
    $archiveDir = Join-Path $archiveRoot $runId
    New-Item -ItemType Directory -Path $archiveDir -Force | Out-Null
    $files = [System.Collections.Generic.List[object]]::new()
    foreach ($name in $outputs.Keys) {
      $spec = $script:OutputSpecs[$name]
      $columns = @($script:CommonColumns + $spec.Columns)
      $rows = @(foreach ($row in (Get-FinOpsSortedRows -Rows @($outputs[$name]) -KeyColumns $spec.Key)) { $values = [ordered]@{} + $common; foreach ($column in $spec.Columns) { $values[$column] = Get-FinOpsProperty $row $column }; [pscustomobject]$values })
      $path = Join-Path $archiveDir $name
      Write-FinOpsCsv -Path $path -Columns $columns -Rows $rows
      $files.Add([ordered]@{ File=$name; Rows=$rows.Count; SHA256=(Get-FinOpsSha256 $path); Required=$true })
    }
    $receipt = [ordered]@{
      Owner='SmartFinOps-Receipt'; ContractVersion='1.0'; TenantKey=$Tenant; Producer='SmartFinOps-Workplace-Valuation.ps1'; ScriptVersion=$script:ScriptVersion
      RunId=$runId; StartedAtUtc=$runStart.ToString('o'); CompletedAtUtc=[datetime]::UtcNow.ToString('o'); Status=$status
      LicencesMeasured=$result.Measured; SnapshotId=$result.SnapshotId; LicencesReceiptRunId=$result.ReceiptRunId; Provisional=$result.Provisional; ProvisionalLevers=@($result.ProvisionalLevers); MissingDomains=$result.MissingDomains
      Inputs=@(foreach ($record in ($sources.Values | Sort-Object SourceFile)) { [ordered]@{ SourceFile=$record.SourceFile; SourceReceiptRunId=$record.SourceReceiptRunId; SHA256=$record.ActualSHA256; QualificationStatus=$record.QualificationStatus } })
      Settings=[ordered]@{ FreshnessDailyHours=$dailyHours; FreshnessWeeklyDays=$weeklyDays; ArchiveRetentionRuns=$keepRuns; AsOfDate=$result.AsOfDate.ToString('yyyy-MM-dd'); ContractsFileSHA256=$(if (Test-Path -LiteralPath $contractsFile) { Get-FinOpsSha256 $contractsFile } else { '' }) }
      Files=$files.ToArray()
      Controls=[ordered]@{ Status=$(if ($failed.Count) { 'Fail' } elseif ($result.Measured) { 'Pass' } else { 'Not applicable' }); FailedCount=$failed.Count }
    }
    Write-FinOpsJson -Path (Join-Path $archiveDir $script:ReceiptName) -Document $receipt
    New-Item -ItemType Directory -Path $currentRoot -Force | Out-Null
    $currentReceipt = Join-Path $currentRoot $script:ReceiptName
    if ($status -eq 'Completed') {
      # Disqualify the current files while they are replaced, then publish the completed receipt.
      $publishing = [ordered]@{} + $receipt; $publishing.Status = 'Publishing'
      Write-FinOpsJson -Path $currentReceipt -Document $publishing
      foreach ($name in $outputs.Keys) { Copy-Item -LiteralPath (Join-Path $archiveDir $name) -Destination (Join-Path $currentRoot $name) -Force }
      foreach ($file in $files) { if ((Get-FinOpsSha256 (Join-Path $currentRoot $file.File)) -cne $file.SHA256) { throw "Published file differs from the archive: $($file.File)" } }
      Write-FinOpsJson -Path $currentReceipt -Document $receipt
      Write-FinOpsLog "Published FinOps outputs and receipt ($status)."
    }
    else {
      # Failed guard: current files stay but the receipt disqualifies them.
      Write-FinOpsJson -Path $currentReceipt -Document $receipt
      Write-FinOpsLog 'Equivalence guard failed: current FinOps files kept but disqualified by the receipt.' 'ERROR'
    }
    # Without a measured Licences report the realization baselines are unknown: no purge.
    if ($result.Measured) { Invoke-FinOpsArchiveRetention -ArchiveRoot $archiveRoot -KeepRuns $keepRuns -ProtectedRuns $realization.ProtectedRuns -CurrentRunId $runId }
  }
}
catch {
  $exitCode = 1
  $runError = $_.Exception.Message
  Write-FinOpsLog ("Run failed: {0}" -f $runError) 'ERROR'
  Write-FinOpsLog ("Stack: {0}" -f ($_.ScriptStackTrace -replace '\r?\n', ' | ')) 'ERROR'
  if (-not $ValidateOnly -and $currentRoot) {
    try {
      New-Item -ItemType Directory -Path $currentRoot -Force | Out-Null
      $failedReceipt = [ordered]@{ Owner='SmartFinOps-Receipt'; ContractVersion='1.0'; TenantKey=$Tenant; Producer='SmartFinOps-Workplace-Valuation.ps1'; ScriptVersion=$script:ScriptVersion
        RunId=$runId; StartedAtUtc=$runStart.ToString('o'); CompletedAtUtc=[datetime]::UtcNow.ToString('o'); Status='Failed'; LicencesMeasured=$false; Error=$runError; Files=@() }
      Write-FinOpsJson -Path (Join-Path $currentRoot $script:ReceiptName) -Document $failedReceipt
      if (-not $archiveDir) { $archiveDir = Join-Path (Join-Path $currentRoot 'Archive') $runId; New-Item -ItemType Directory -Path $archiveDir -Force | Out-Null }
      Write-FinOpsJson -Path (Join-Path $archiveDir $script:ReceiptName) -Document $failedReceipt
    }
    catch { Write-FinOpsLog ("Failed receipt could not be written: {0}" -f $_.Exception.Message) 'ERROR' }
  }
}
finally {
  if (-not $ValidateOnly -and $archiveDir -and (Test-Path -LiteralPath $archiveDir)) {
    [IO.File]::WriteAllLines((Join-Path $archiveDir $script:LogName), $script:LogLines, [Text.UTF8Encoding]::new($false))
  }
}
exit $exitCode
#endregion
