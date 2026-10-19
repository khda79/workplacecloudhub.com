#Requires -Version 7.0
<#
.SYNOPSIS
  Offline tests of the SmartFinOps Workplace V2 executive HTML report.
.VERSION
1.3.1
.NOTES
  Synthetic FinOps_* fixtures and receipts only, in an isolated temporary runtime (no local configuration,
  export or tenant). The input contract is checked against the valuation engine output specification.
#>
[CmdletBinding()]
param([switch]$KeepEvidence)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$project = Split-Path $PSScriptRoot -Parent
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartFinOps-Report-' + [guid]::NewGuid().ToString('N'))
$runtime = Join-Path $root 'runtime'
$report = Join-Path $runtime 'SmartFinOps\SmartFinOps-Workplace-Report.ps1'
$pwsh = (Get-Process -Id $PID).Path
# Profile 'prod' whose organization TenantKey is 'contoso-prod'.
$tenant = 'contoso-prod'
$profileKey = 'prod'
$baseConfig = '{ "SmartM365TenantKey": "contoso-prod", "EnableSharePointUpload": false }'
$runId = '20261018T000000000Z-abcdef12'
$skuGuid = '11111111-2222-3333-4444-555555555555'
$userGuid = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee'
$script:checks = 0
$script:failures = [System.Collections.Generic.List[string]]::new()

function Check {
  param([bool]$Condition, [string]$Message)
  $script:checks++
  if (-not $Condition) { $script:failures.Add($Message); Write-Host "FAIL $Message" -ForegroundColor Red }
}

foreach ($file in @('SmartFinOps-Workplace-Report.ps1','SmartFinOps-Workplace-Valuation.ps1','SmartFinOps.global.local.json.template','Config\SmartFinOps-TenantContext.ps1')) {
  $target = Join-Path $runtime "SmartFinOps\$file"
  New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force | Out-Null
  Copy-Item -LiteralPath (Join-Path $project $file) -Destination $target
}
$baseConfig | Set-Content -LiteralPath (Join-Path $runtime 'SmartFinOps\SmartFinOps.global.local.json') -Encoding utf8NoBOM

#region Fixture
$columns = [ordered]@{
  'FinOps_Sources.csv' = @('SourceFile','Producer','ReceiptFile','SourceReceiptRunId','ReceiptStatus','ReceiptCompletedAtUtc','Required','PublishedInReceipt','ExpectedSHA256','ActualSHA256','ExpectedRows','ActualRows','FreshnessThresholdHours','AgeHours','QualificationStatus','Reason','MissingDomains','ProvisionalLevers')
  'FinOps_Opportunities.csv' = @('OpportunityKey','Lever','ValueClass','ValuationStatus','StatusReason','SourceFile','SourceCandidateType','SourceRecommendation','SourceReason','MailboxGuard','DependencyStatus','CountedIn','ReportProduct','SkuId','SkuPartNumber','TargetSkuPartNumber','UserId','UserPrincipalName','Units','UnitPriceMonthly','TargetUnitPriceMonthly','MonthlyAmount','AnnualAmount','Currency','PriceRowKey','PricingReceiptRunId','PriceSource','PriceIsEstimated','CapRank','ContractKey','RenewalDate','RealizableFrom','AdDomain','Country','Entity','Department','ShowbackJoinStatus')
  'FinOps_SkuValuation.csv' = @('SkuId','SkuPartNumber','SkuDisplayName','ReportProduct','Excluded','EnabledUnits','ConsumedUnits','AvailableUnits','ReserveUnits','MinimumUnits','CandidateUnitsR','PoolUnits','ReusedUnits','ReducibleUnits','CapUnits','RealizableUnits','CapacityRealizableUnits','CappedUnits','ReviewUnits','NotMeasuredUnits','BlockedUnits','ReferencedUnits','UnitPriceMonthly','PriceRowKey','PriceIsEstimated','RealizableAnnualAmount','AvoidedAnnualAmount','ValuationStatus','ContractKey','ContractStatus','ReductionPolicy','RenewalDate','RealizableFrom','UsageEvidence')
  'FinOps_Showback.csv' = @('Dimension','DimensionValue','Lever','ValueClass','ValuationStatus','Users','Units','AnnualAmount','Currency','Provisional')
  'FinOps_Summary.csv' = @('KpiKey','ValueClass','Lever','ValuationStatus','Users','Units','AnnualAmount','Currency','Provisional','MissingDomains','SnapshotGeneratedAtUtc','PricingReceiptRunId','Note')
  'FinOps_Realization.csv' = @('SkuId','SkuPartNumber','BaselineFinOpsRunId','BaselineSnapshotId','EnabledUnitsBefore','EnabledUnitsAfter','Decrease','RemovedCandidateLicenses','BaselineCapacityUnits','AttributableUnits','RealizedUnits','NotAttributedUnits','UnitPriceMonthly','PriceRowKey','RealizedAnnualAmount','Status')
  'FinOps_Controls.csv' = @('Control','Key','SourceCount','FinOpsCount','Status','Detail')
}
$common = @('TenantKey','FinOpsRunId','SnapshotId','ReceiptRunId')
$escapePayload = '<script>alert(1)</script> & "q" ''x'''

function Write-FixtureCsv {
  param([string]$Path, [string[]]$Columns, [object[]]$Rows, [hashtable]$CommonValues)
  $builder = [System.Text.StringBuilder]::new()
  [void]$builder.AppendLine((@($common + $Columns | ForEach-Object { '"' + $_ + '"' }) -join ','))
  foreach ($row in $Rows) {
    $values = foreach ($column in @($common + $Columns)) {
      $value = if ($column -in $common -and -not $row.Contains($column)) { $CommonValues[$column] } elseif ($row.Contains($column)) { [string]$row[$column] } else { '' }
      '"' + ([string]$value).Replace('"', '""') + '"'
    }
    [void]$builder.AppendLine(($values -join ','))
  }
  [IO.File]::WriteAllText($Path, $builder.ToString(), [Text.UTF8Encoding]::new($false))
}

function Get-FixtureData {
  param([switch]$NotMeasured)
  $kpi = { param($key, $class, $lever, $status, $users, $units, $amount, $note = '') [ordered]@{ KpiKey=$key; ValueClass=$class; Lever=$lever; ValuationStatus=$status; Users=$users; Units=$units; AnnualAmount=$amount; Currency='EUR'; Provisional=$(if ($status -eq 'Provisional') { 'True' } else { 'False' }); MissingDomains=''; SnapshotGeneratedAtUtc=$(if ($NotMeasured) { '' } else { '2026-10-17' }); PricingReceiptRunId='pricerun0001'; Note=$note } }
  if ($NotMeasured) {
    $summary = @(foreach ($key in @('TOTAL_R','TOTAL_R_DISABLED_AND_SHARED','TOTAL_R_ENABLED_INACTIVE','TOTAL_R_COPILOT','TOTAL_R_CAPACITY','TOTAL_RC_E3_TO_F3','TOTAL_E_REUSE')) { & $kpi $key '' '' 'Not measured' '' '' '' 'Licences report not qualified' }; & $kpi 'TOTAL_E_ESU' 'E' 'L13' 'Not measured' 0 '' '' 'No ESU device population')
    return @{
      'FinOps_Sources.csv' = @([ordered]@{ SourceFile='M365_Licenses_ReportSummary.csv'; QualificationStatus='NotMeasured'; Reason='source receipt is Failed' })
      'FinOps_Opportunities.csv' = @([ordered]@{ OpportunityKey='L13|WIN10_ESU'; Lever='L13'; ValueClass='E'; ValuationStatus='Not measured'; StatusReason='No ESU device population'; Units=''; Currency='EUR' })
      'FinOps_SkuValuation.csv' = @(); 'FinOps_Showback.csv' = @(); 'FinOps_Summary.csv' = $summary; 'FinOps_Realization.csv' = @()
      'FinOps_Controls.csv' = @([ordered]@{ Control='G0'; Key='Licences'; Status='Not applicable'; Detail='Licences report not qualified: nothing valued' })
    }
  }
  $summary = @(
    & $kpi 'TOTAL_R' 'R' 'L1;L2;L3;L5;L6;L7' 'Valued' 6 9 '12345.67' 'Realizable at contract or renewal'
    & $kpi 'TOTAL_R_DISABLED_AND_SHARED' 'R' 'L1;L2;L5' 'Valued' 3 3 '3000.00'
    & $kpi 'TOTAL_R_ENABLED_INACTIVE' 'R' 'L3' 'Valued' 2 2 '2222.22'
    & $kpi 'TOTAL_R_COPILOT' 'R' 'L6' 'Valued' 0 '' ''
    & $kpi 'TOTAL_R_CAPACITY' 'R' 'L7' 'Valued' 0 4 '7123.45'
    & $kpi 'TOTAL_RC_E3_TO_F3' 'Rc' 'L4' 'Provisional' 1 1 '349.44'
    & $kpi 'TOTAL_E_REUSE' 'E' 'L8' 'Valued' 0 1 '103.92'
    & $kpi 'TOTAL_E_ESU' 'E' 'L13' 'Not measured' 0 '' '' 'No ESU device population'
    & $kpi 'NOT_MEASURED_MAILBOX' 'S' 'L5' 'Not measured' 1 1 ''
    & $kpi 'DETAIL|R|L1|Valued' 'R' 'L1' 'Valued' 1 1 '1000.00'
    & $kpi 'DETAIL|R|L3|Valued' 'R' 'L3' 'Valued' 2 2 '2222.22'
    & $kpi 'DETAIL|R|L3|Capped' 'R' 'L3' 'Capped' 1 1 ''
    & $kpi 'DETAIL|R|L5|Not valued' 'R' 'L5' 'Not valued' 1 1 ''
    & $kpi 'DETAIL|R|L7|Valued' 'R' 'L7' 'Valued' 0 4 '7123.45'
    & $kpi 'DETAIL|Rc|L4|Provisional' 'Rc' 'L4' 'Provisional' 1 1 '349.44'
    & $kpi 'DETAIL|S|L5|Review' 'S' 'L5' 'Review' 1 1 ''
    & $kpi 'DETAIL|S|L5|Blocked' 'S' 'L5' 'Blocked' 1 1 ''
    & $kpi 'DETAIL|S|L9|Signal' 'S' 'L9' 'Signal' 0 2 ''
    & $kpi 'GAP|UserMailboxes|NoSkus' 'S' 'L12' 'Signal' '' 4 ''
    & $kpi 'GAP|AdGapActivity|NoObservedMatch' 'S' 'L12' 'Not measured' '' '' ''
    & $kpi 'PRICE|TOTAL_R|Estimated' 'R' 'L1;L2;L3;L5;L6;L7' 'Valued' 5 7 '9000.00'
    & $kpi 'PRICE|TOTAL_R|Contract' 'R' 'L1;L2;L3;L5;L6;L7' 'Valued' 1 2 '3345.67'
    & $kpi 'PRICE|TOTAL_R|Unknown' 'R' 'L1;L2;L3;L5;L6;L7' 'Valued' 0 0 '0.00'
    & $kpi 'FORECAST|TOTAL_R|2027-03-31' 'R' 'L1;L2;L3;L5;L6;L7' 'Valued' 4 6 '10000.00'
    & $kpi 'FORECAST|TOTAL_R|Renewal unknown' 'R' 'L1;L2;L3;L5;L6;L7' 'Valued' 2 3 '2345.67'
    & $kpi 'SHOWBACK|Department|TOTAL_R|Finance' 'R' 'x' 'Valued' 3 3 '5000.00'
    & $kpi 'SHOWBACK|Department|TOTAL_R|Tiny' 'R' 'x' 'Valued' 1 1 '777.77'
    & $kpi 'SHOWBACK|Department|TOTAL_R|Small2' 'R' 'x' 'Valued' 2 2 '111.11'
    & $kpi "SHOWBACK|Department|TOTAL_R|$escapePayload" 'R' 'x' 'Valued' 7 7 '222.22'
    & $kpi 'SHOWBACK|Department|USERS|Finance' '' '' '' 6 '' ''
    & $kpi 'SHOWBACK|Department|USERS|Tiny' '' '' '' 1 '' ''
    & $kpi 'SHOWBACK|Department|USERS|Small2' '' '' '' 2 '' ''
    & $kpi "SHOWBACK|Department|USERS|$escapePayload" '' '' '' 7 '' ''
    & $kpi 'SHOWBACK|Country|TOTAL_R|FR' 'R' 'x' 'Valued' 5 5 '8000.00'
    & $kpi 'SHOWBACK|Country|TOTAL_R|LU' 'R' 'x' 'Valued' 1 1 '4321.98'
    & $kpi 'SHOWBACK|Country|TOTAL_R|SKU level' 'R' 'L7' 'Valued' 0 4 '7123.45'
    & $kpi 'SHOWBACK|Country|USERS|FR' '' '' '' 10 '' ''
    & $kpi 'SHOWBACK|Country|USERS|LU' '' '' '' 3 '' ''
  )
  $opp = { param($key, $lever, $class, $status, $reason, $units, $amount) [ordered]@{ OpportunityKey=$key; Lever=$lever; ValueClass=$class; ValuationStatus=$status; StatusReason=$reason; SkuId=$skuGuid; SkuPartNumber='SPE_E3'; UserId=$userGuid; UserPrincipalName='alice.secret@contoso.invalid'; Units=$units; AnnualAmount=$amount; Currency='EUR'; AdDomain='north.contoso.invalid'; Country='FR'; Entity='Site 1'; Department='Finance'; ShowbackJoinStatus='ImmutableId' } }
  return @{
    'FinOps_Sources.csv' = @(
      [ordered]@{ SourceFile='M365_Licenses_ReportSummary.csv'; Producer='SmartM365-Licences-Inventory.ps1'; SourceReceiptRunId='licrun0001'; ReceiptStatus='Completed'; ReceiptCompletedAtUtc='2026-10-17T00:14:00.0000000Z'; ExpectedRows='5'; ActualRows='5'; FreshnessThresholdHours='48'; AgeHours='3.2'; ActualSHA256='ABCDEF0123456789ABCDEF'; QualificationStatus='Qualified' }
      [ordered]@{ SourceFile='M365_OneDrive_Usage.csv'; Producer='SmartM365-Licences-Inventory.ps1'; QualificationStatus='Provisional'; Reason='Provisional in the Licences report: owner carol@contoso.invalid'; ProvisionalLevers='L4' }
    )
    'FinOps_Opportunities.csv' = @(
      & $opp "L1|$userGuid|$skuGuid" 'L1' 'R' 'Valued' '' 1 '1000.00'
      & $opp "L5|$userGuid|b" 'L5' 'S' 'Blocked' 'Mailbox retention: litigation hold for bob@contoso.invalid' 1 ''
      & $opp "L5|$userGuid|c" 'L5' 'S' 'Review' 'Entra disabled; AD state is not measured' 1 ''
      & $opp "L5|$userGuid|d" 'L5' 'R' 'Not valued' "no EUR price effective for $escapePayload" 1 ''
    )
    'FinOps_SkuValuation.csv' = @(
      [ordered]@{ SkuId=$skuGuid; SkuPartNumber='SPE_E3'; SkuDisplayName='<b>E3 & co</b>'; ReportProduct='Microsoft 365 E3'; Excluded='False'; EnabledUnits='100'; ConsumedUnits='96'; AvailableUnits='4'; ReserveUnits='0'; MinimumUnits=''; CandidateUnitsR='5'; ReusedUnits='0'; CapUnits=''; RealizableUnits='9'; CapacityRealizableUnits='4'; CappedUnits='1'; UnitPriceMonthly='37.78'; PriceIsEstimated='True'; RealizableAnnualAmount='9000.00'; AvoidedAnnualAmount=''; ValuationStatus='Valued'; ContractKey='EA-CONTOSO'; ContractStatus='Commitment unknown'; ReductionPolicy='AtRenewal'; RenewalDate='2027-03-31'; RealizableFrom='2027-03-31'; UsageEvidence='Not applicable' }
      [ordered]@{ SkuId='22222222-2222-3333-4444-555555555555'; SkuPartNumber='Microsoft_365_Copilot'; SkuDisplayName='Copilot'; ReportProduct='Microsoft 365 Copilot'; Excluded='False'; EnabledUnits='10'; ConsumedUnits='10'; AvailableUnits='0'; ReserveUnits='0'; CandidateUnitsR='1'; ReusedUnits='0'; RealizableUnits='1'; CappedUnits='0'; UnitPriceMonthly='26.00'; PriceIsEstimated='False'; RealizableAnnualAmount='3345.67'; ValuationStatus='Valued'; ContractStatus='No contract'; ReductionPolicy='Unknown'; RealizableFrom='Renewal unknown'; UsageEvidence='Measured' }
      [ordered]@{ SkuId='33333333-2222-3333-4444-555555555555'; SkuPartNumber='DYN365_FINANCE'; SkuDisplayName='Dynamics'; ReportProduct='Dynamics 365'; Excluded='False'; EnabledUnits='2'; ConsumedUnits='4'; AvailableUnits='-2'; ReserveUnits='0'; CandidateUnitsR='1'; ReusedUnits='0'; RealizableUnits='1'; CappedUnits='0'; UnitPriceMonthly=''; PriceIsEstimated=''; RealizableAnnualAmount=''; ValuationStatus='Not valued'; ContractStatus='No contract'; ReductionPolicy='Unknown'; RealizableFrom='Renewal unknown'; UsageEvidence='Not measured' }
      [ordered]@{ SkuId='44444444-2222-3333-4444-555555555555'; SkuPartNumber='FLOW_FREE'; SkuDisplayName='Flow free'; Excluded='True'; EnabledUnits='10000'; ConsumedUnits='5'; AvailableUnits='9995'; ValuationStatus='Not valued' }
    )
    'FinOps_Showback.csv' = @(
      [ordered]@{ Dimension='Department'; DimensionValue='Finance'; Lever='L5'; ValueClass='S'; ValuationStatus='Review'; Users='1'; Units='1'; Currency='EUR'; Provisional='False' }
      [ordered]@{ Dimension='Department'; DimensionValue='Tiny'; Lever='L5'; ValueClass='S'; ValuationStatus='Blocked'; Users='1'; Units='1'; Currency='EUR'; Provisional='False' }
      [ordered]@{ Dimension='Country'; DimensionValue='FR'; Lever='L1'; ValueClass='R'; ValuationStatus='Valued'; Users='1'; Units='1'; AnnualAmount='1000.00'; Currency='EUR'; Provisional='False' }
    )
    'FinOps_Summary.csv' = $summary
    'FinOps_Realization.csv' = @([ordered]@{ SkuId=$skuGuid; SkuPartNumber='SPE_E3'; EnabledUnitsAfter='100'; Status='Not measured: no qualified baseline run' })
    'FinOps_Controls.csv' = @(
      [ordered]@{ Control='G1'; Key='Recovery|Microsoft 365 E3'; SourceCount='3'; FinOpsCount='3'; Status='Pass'; Detail='ReportCandidates vs FinOps' }
      [ordered]@{ Control='G4'; Key="$skuGuid|Review"; SourceCount='1'; FinOpsCount='1'; Status='Pass'; Detail='Capacity counters' }
    )
  }
}

function New-ReportFixture {
  param([string]$Name, [switch]$NotMeasured, [scriptblock]$ChangeData, [scriptblock]$ChangeReceipt)
  # SmartM365 data tree of the profile: the valuation outputs in DATA-LAST, the run history in DATA-ALL, logs in LOG-ALL.
  $folder = Join-Path $root "$Name\DATA-LAST"
  New-Item -ItemType Directory -Path $folder, (Join-Path $root "$Name\DATA-ALL\FinOps\Workplace\$runId"), (Join-Path $root "$Name\LOG-ALL") -Force | Out-Null
  $data = Get-FixtureData -NotMeasured:$NotMeasured
  if ($ChangeData) { & $ChangeData $data }
  $values = @{ TenantKey=$tenant; FinOpsRunId=$runId; SnapshotId=$(if ($NotMeasured) { '' } else { 'snap-1' }); ReceiptRunId=$(if ($NotMeasured) { '' } else { 'licrun0001' }) }
  $files = foreach ($name in $columns.Keys) {
    $path = Join-Path $folder $name
    Write-FixtureCsv -Path $path -Columns $columns[$name] -Rows @($data[$name]) -CommonValues $values
    [ordered]@{ File=$name; Rows=@($data[$name]).Count; SHA256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash; Required=$true }
  }
  $receipt = [ordered]@{
    Owner='SmartFinOps-Receipt'; ContractVersion='1.1'; ProfileKey=$profileKey; TenantKey=$tenant; Producer='SmartFinOps-Workplace-Valuation.ps1'; ScriptVersion='2.1.4'
    RunId=$runId; StartedAtUtc='2026-10-18T00:00:00.0000000Z'; CompletedAtUtc='2026-10-18T00:02:00.0000000Z'; Status='Completed'
    LicencesMeasured=(-not $NotMeasured); SnapshotId=$values.SnapshotId; LicencesReceiptRunId=$values.ReceiptRunId; Provisional=(-not $NotMeasured); ProvisionalLevers=@($(if (-not $NotMeasured) { 'L4' })); MissingDomains=''
    Inputs=@(); Settings=[ordered]@{ FreshnessDailyHours=48; FreshnessWeeklyDays=8; ArchiveRetentionRuns=30; AsOfDate='2026-10-17'; ContractsFileSHA256='0123456789ABCDEF0123' }
    Files=@($files); Controls=[ordered]@{ Status=$(if ($NotMeasured) { 'Not applicable' } else { 'Pass' }); FailedCount=0 }
  }
  if ($ChangeReceipt) { & $ChangeReceipt $receipt $folder }
  if ($null -ne $receipt) { $receipt | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath (Join-Path $folder 'SmartFinOps_Workplace-Valuation.current.json.txt') -Encoding utf8NoBOM }
  return $folder
}

function Get-TenantFolder { param([string]$Folder, [string]$Name) return (Join-Path (Split-Path $Folder -Parent) $Name) }
function Get-ReportLogPath { param([string]$Folder) return @(Get-ChildItem -LiteralPath (Get-TenantFolder $Folder 'LOG-ALL\SmartFinOps-Workplace-Report') -Filter 'SmartFinOps-Workplace-Report-*.log' -ErrorAction SilentlyContinue | Sort-Object Name | ForEach-Object FullName | Select-Object -Last 1) }

function Invoke-Report {
  param([string]$Folder, [string]$Culture, [switch]$ValidateOnly)
  $arguments = @('-NoProfile','-ExecutionPolicy','Bypass')
  $dataAll = Get-TenantFolder $Folder 'DATA-ALL'; $logAll = Get-TenantFolder $Folder 'LOG-ALL'
  if ($Culture) {
    $command = "[Globalization.CultureInfo]::CurrentCulture = '$Culture'; & '$report' -Tenant '$profileKey' -SmartM365LatestCsvFolderPath '$Folder' -SmartM365DataAllRootPath '$dataAll' -SmartM365LogAllRootPath '$logAll'$(if ($ValidateOnly) { ' -ValidateOnly' }); exit `$LASTEXITCODE"
    $arguments += @('-Command', $command)
  }
  else {
    $arguments += @('-File', $report, '-Tenant', $profileKey, '-SmartM365LatestCsvFolderPath', $Folder, '-SmartM365DataAllRootPath', $dataAll, '-SmartM365LogAllRootPath', $logAll)
    if ($ValidateOnly) { $arguments += '-ValidateOnly' }
  }
  $log = & $pwsh @arguments 2>&1 | Out-String
  $path = Join-Path $Folder 'SmartFinOps_Workplace-Report.html'
  return [pscustomobject]@{ ExitCode=$LASTEXITCODE; Log=$log; Path=$path; Html=$(if (Test-Path -LiteralPath $path) { Get-Content -LiteralPath $path -Raw } else { '' }) }
}

function Get-NormalizedHtml { param([string]$Html) return ($Html -replace '<span id="generated">[^<]*</span>', '<span id="generated"></span>') }
#endregion

try {
  #region 1. Input contract equals the engine output specification
  . (Join-Path $runtime 'SmartFinOps\SmartFinOps-Workplace-Valuation.ps1') -Tenant $tenant
  $engineSpecs = $script:OutputSpecs; $engineCommon = $script:CommonColumns
  . $report -Tenant $tenant
  Check ((@($engineCommon) -join ',') -ceq (@($script:CommonColumns) -join ',')) 'common columns match the engine'
  foreach ($name in $engineSpecs.Keys) { Check ((@($engineSpecs[$name].Columns) -join ',') -ceq (@($script:InputColumns[$name]) -join ',')) "report input contract matches the engine for $name" }
  Check (@($engineSpecs.Keys).Count -eq 7 -and @($script:InputColumns.Keys).Count -eq 7) 'seven FinOps files'
  foreach ($name in $columns.Keys) { Check ((@($columns[$name]) -join ',') -ceq (@($script:InputColumns[$name]) -join ',')) "test fixture columns match the report for $name" }
  Check ((ConvertTo-ReportHtml $escapePayload) -ceq '&lt;script&gt;alert(1)&lt;/script&gt; &amp; &quot;q&quot; &#39;x&#39;') 'HTML encoding of < > & " and single quote'
  Check ((ConvertTo-ReportSafeText 'hold for bob@contoso.invalid on 11111111-2222-3333-4444-555555555555') -ceq 'hold for [user] on [id]') 'free text never keeps an e-mail address or a GUID'
  Check ((Format-ReportEur '1234.50') -ceq '<span class="amt" title="EUR 1234.50">EUR 1,235</span>' -and (Format-ReportEur '') -ceq '' -and (Format-ReportEur '12,5') -ceq '') 'invariant amount format, empty is not zero'
  Check ((Get-ReportAmountCell '' 'Not measured') -like '*Not measured*' -and (Get-ReportAmountCell '0.00' 'Valued') -like '*EUR 0<*' -and (Get-ReportAmountCell '' 'Valued') -like '*None identified*') 'empty amount shows its status, measured zero shows EUR 0'
  #endregion

  #region 2. Nominal report
  $nominal = New-ReportFixture 'nominal'
  $run = Invoke-Report $nominal
  $html = $run.Html
  Check ($run.ExitCode -eq 3) "nominal exit code 3, provisional L4 (got $($run.ExitCode)) $($run.Log)"
  $historyPage = Get-TenantFolder $nominal "DATA-ALL\FinOps\Workplace\$runId\SmartFinOps_Workplace-Report.html"
  Check ((Test-Path -LiteralPath $historyPage) -and (Get-Content -LiteralPath $historyPage -Raw) -ceq $html) 'report copied into the run history (DATA-ALL\FinOps\Workplace\<RunId>)'
  $order = @('executive-summary','financial-potential','priority-decisions','trade-offs','capacity-forecast','realization','showback','assumptions','appendix' | ForEach-Object { $html.IndexOf("id=""$_""", [StringComparison]::Ordinal) })
  Check (@($order | Where-Object { $_ -lt 0 }).Count -eq 0 -and (($order | Sort-Object) -join ',') -eq ($order -join ',')) 'nine sections in the RULES order'
  Check ($html.Contains('title="EUR 12345.67">EUR 12,346</span>') -and $html.Contains('EUR 2,222') -and $html.Contains('EUR 349') -and $html.Contains('EUR 104')) 'totals displayed from FinOps_Summary'
  Check ($html.Contains('None identified')) 'empty valued total shown as none identified, not zero'
  Check ($html.Contains('title="EUR 0.00">EUR 0</span>')) 'measured zero displayed as EUR 0'
  foreach ($label in @('Provisional','Not valued','Not measured','Review','Blocked','Capped','Signal')) { Check ($html.Contains(">$label</span>")) "status badge $label" }
  foreach ($label in @('Catalogue estimate','Contract price','No price')) { Check ($html.Contains($label)) "price origin $label" }
  Check ($html.Contains('Over-assigned 2') -and $html.Contains('<strong>Renewal unknown</strong>') -and $html.Contains('No qualified baseline run yet')) 'over-assigned, renewal unknown and no baseline shown'
  Check ($html.Contains('Provisional figures for levers L4')) 'provisional banner with its levers'
  Check ($html.Contains('Profile: prod |') -and $run.Log.Contains("Profile 'prod' - TenantKey 'contoso-prod'")) 'profile shown, TenantKey checked and logged'
  Check ($html.Contains('<svg class="bars"') -and $html.Contains('<table>')) 'inline SVG charts with tables'
  Check (-not $html.Contains('FLOW_FREE') -and $html.Contains('1 excluded SKU(s)')) 'excluded SKUs not listed'
  Check ($html.Contains('SPE_E3|Review') -and -not $html.Contains("$skuGuid|Review")) 'control keys show the SKU part number'
  #endregion

  #region 3. Confidentiality, autonomy, escaping
  Check (-not $html.Contains('alice.secret') -and -not $html.Contains($userGuid) -and -not ($html -match '(?i)[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}')) 'no UPN, UserId or e-mail address in the HTML'
  Check (-not ($html -match '(?i)\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b')) 'no GUID in the HTML'
  Check (-not ($html -match '(?i)https?:|//|<link|<script|\bsrc=|@import|url\(|<iframe|<img')) 'no external resource, no script'
  Check ($html.Contains('&lt;script&gt;alert(1)&lt;/script&gt; &amp; &quot;q&quot; &#39;x&#39;') -and $html.Contains('&lt;b&gt;E3 &amp; co&lt;/b&gt;')) 'CSV values HTML-encoded'
  Check ($html.Contains('Other reason')) 'reason quoting a user replaced'
  #endregion

  #region 4. Showback small groups
  Check ($html.Contains('Other (groups under 5 users)') -and -not $html.Contains('>Tiny<') -and -not $html.Contains('>Small2<') -and -not $html.Contains('EUR 778') -and -not $html.Contains('EUR 111')) 'department groups under 5 users combined, values not shown'
  Check ($html.Contains('EUR 5,889') -and -not $html.Contains('>Finance<')) 'combined row reaches the threshold by adding the smallest group'
  Check (-not $html.Contains('>LU<') -and -not $html.Contains('EUR 4,322') -and -not $html.Contains('>FR<') -and $html.Contains('EUR 12,322')) 'a single small country is never shown alone'
  $configured = New-ReportFixture 'threshold'
  '{ "SmartM365TenantKey": "contoso-prod", "EnableSharePointUpload": false, "Report": { "ShowbackMinimumGroupUsers": 1 } }' | Set-Content -LiteralPath (Join-Path $runtime 'SmartFinOps\SmartFinOps.global.local.json') -Encoding utf8NoBOM
  $runThreshold = Invoke-Report $configured
  $baseConfig | Set-Content -LiteralPath (Join-Path $runtime 'SmartFinOps\SmartFinOps.global.local.json') -Encoding utf8NoBOM
  Check ($runThreshold.ExitCode -eq 3 -and $runThreshold.Html.Contains('>Tiny<') -and $runThreshold.Html.Contains('EUR 778')) 'threshold read from the local configuration'
  #endregion

  #region 5. Determinism
  $again = Invoke-Report $nominal
  Check ((Get-NormalizedHtml $again.Html) -ceq (Get-NormalizedHtml $html)) 'same inputs give the same HTML except the generation time'
  $french = Invoke-Report $nominal -Culture 'fr-FR'
  Check ($french.ExitCode -eq 3 -and (Get-NormalizedHtml $french.Html) -ceq (Get-NormalizedHtml $html)) 'host culture does not change the HTML'
  #endregion

  #region 6. Not available: qualification failures, never a figure
  $figures = @('12,346', '9,000', '7,123', '2,222')
  $cases = [ordered]@{
    'receipt-missing' = @{ Drop=$true; Reason='receipt is missing' }
    'receipt-json' = @{ After={ param($f) Set-Content -LiteralPath (Join-Path $f 'SmartFinOps_Workplace-Valuation.current.json.txt') -Value '{ not json' }; Reason='not valid JSON' }
    'receipt-owner' = @{ Receipt={ param($r) $r.Owner = 'Other' }; Reason='unsupported owner' }
    'engine-2.1.3' = @{ Receipt={ param($r) $r.ScriptVersion = '2.1.3' }; Reason='not supported' }
    'engine-2.3.0' = @{ Receipt={ param($r) $r.ScriptVersion = '2.3.0' }; Reason='not supported' }
    'contract-1.0' = @{ Receipt={ param($r) $r.ContractVersion = '1.0' }; Reason='unsupported owner or contract version' }
    'tenantkey-is-profile' = @{ Receipt={ param($r) $r.TenantKey = 'prod' }; Reason='TenantKey differs' }
    'other-profile' = @{ Receipt={ param($r) $r.ProfileKey = 'test' }; Reason='ProfileKey differs' }
    'receipt-failed' = @{ Receipt={ param($r) $r.Status = 'Failed'; $r.Error = '<b>boom</b>' }; Reason='Failed: &lt;b&gt;boom&lt;/b&gt;' }
    'publishing' = @{ Receipt={ param($r) $r.Status = 'Publishing' }; Reason='Publishing' }
    'tenant' = @{ Receipt={ param($r) $r.TenantKey = 'fabrikam' }; Reason='another tenant' }
    'hash' = @{ After={ param($f) Add-Content -LiteralPath (Join-Path $f 'FinOps_Summary.csv') -Value '"x"' }; Reason='File hash mismatch: FinOps_Summary.csv' }
    'rows' = @{ Receipt={ param($r) $r.Files[4].Rows = 1 }; Reason='Row count mismatch: FinOps_Summary.csv' }
    'file-missing' = @{ After={ param($f) Remove-Item -LiteralPath (Join-Path $f 'FinOps_Controls.csv') }; Reason='File missing: FinOps_Controls.csv' }
    'files-list' = @{ Receipt={ param($r) $r.Files = @($r.Files | Select-Object -First 6) }; Reason='exactly the seven' }
    'mixed-run' = @{ Data={ param($d) $d['FinOps_Summary.csv'][0]['FinOpsRunId'] = '20261017T000000000Z-00000000' }; Reason='Mixed run in FinOps_Summary.csv' }
    'controls-fail' = @{ Data={ param($d) $d['FinOps_Controls.csv'][0]['Status'] = 'Fail' }; Reason='controls failed' }
    'controls-receipt' = @{ Receipt={ param($r) $r.Controls.FailedCount = 1 }; Reason='controls failed' }
  }
  foreach ($case in $cases.Keys) {
    $spec = $cases[$case]
    $folder = New-ReportFixture $case -ChangeData $spec['Data'] -ChangeReceipt $spec['Receipt']
    if ($spec['Drop']) { Remove-Item -LiteralPath (Join-Path $folder 'SmartFinOps_Workplace-Valuation.current.json.txt') -Force }
    if ($spec['After']) { & $spec['After'] $folder }
    $caseRun = Invoke-Report $folder
    Check ($caseRun.ExitCode -eq 1 -and $caseRun.Html.Contains('<h1>Report not available</h1>')) "$case : Report not available page (exit $($caseRun.ExitCode))"
    Check ($caseRun.Html.Contains($spec.Reason)) "$case : reason shown ($($spec.Reason))"
    Check (@($figures | Where-Object { $caseRun.Html.Contains($_) }).Count -eq 0 -and $caseRun.Html -notmatch 'EUR \d') "$case : no figure shown"
  }
  # Schema drift with consistent hashes.
  $schema = New-ReportFixture 'schema' -ChangeReceipt { param($r, $f)
    $path = Join-Path $f 'FinOps_Showback.csv'
    (Get-Content -LiteralPath $path -Raw).Replace('"Provisional"', '"IsProvisional"') | Set-Content -LiteralPath $path -Encoding utf8NoBOM -NoNewline
    $r.Files[3].SHA256 = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
  }
  $schemaRun = Invoke-Report $schema
  Check ($schemaRun.ExitCode -eq 1 -and $schemaRun.Html.Contains('Schema mismatch: FinOps_Showback.csv')) 'schema drift refused'
  # A previous report is replaced: no older figure stays visible.
  $replace = New-ReportFixture 'replace'
  $first = Invoke-Report $replace
  Set-Content -LiteralPath (Join-Path $replace 'SmartFinOps_Workplace-Valuation.current.json.txt') -Value ((Get-Content -LiteralPath (Join-Path $replace 'SmartFinOps_Workplace-Valuation.current.json.txt') -Raw).Replace('"Completed"', '"Failed"')) -Encoding utf8NoBOM
  $second = Invoke-Report $replace
  Check ($first.ExitCode -eq 3 -and $first.Html.Contains('EUR 12,346') -and $second.ExitCode -eq 1 -and -not $second.Html.Contains('12,346')) 'a failed run replaces the previous report'
  #endregion

  #region 7. Licences not measured: report generated, everything Not measured
  $notMeasured = New-ReportFixture 'not-measured' -NotMeasured
  $nmRun = Invoke-Report $notMeasured
  Check ($nmRun.ExitCode -eq 3 -and $nmRun.Html.Contains('Licences report not qualified') -and $nmRun.Html.Contains('<h1>Workplace FinOps report</h1>')) 'not measured: report generated with a status banner'
  Check ($nmRun.Html -notmatch 'EUR \d' -and $nmRun.Html.Contains('>Not measured</span>')) 'not measured: no amount, never EUR 0'
  #endregion

  #region 8. ValidateOnly, logs, PowerShell 7 only
  $validate = New-ReportFixture 'validate'
  $validateRun = Invoke-Report $validate -ValidateOnly
  Check ($validateRun.ExitCode -eq 3 -and -not (Test-Path -LiteralPath (Join-Path $validate 'SmartFinOps_Workplace-Report.html')) -and @(Get-ChildItem -LiteralPath (Get-TenantFolder $validate 'LOG-ALL') -Recurse -File).Count -eq 0) 'ValidateOnly writes nothing'
  $nominalLog = Get-ReportLogPath $nominal
  Check ($null -ne $nominalLog -and (Split-Path $nominalLog -Leaf) -cmatch '^SmartFinOps-Workplace-Report-\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\.log$') "report log in LOG-ALL\SmartFinOps-Workplace-Report ($nominalLog)"
  Check (@(Get-ChildItem -LiteralPath $nominal -Filter '*.log').Count -eq 0 -and @(Get-ChildItem -LiteralPath (Get-TenantFolder $nominal 'DATA-ALL') -Recurse -Filter '*.log').Count -eq 0) 'no report log in DATA-LAST or in the run history'
  $logLines = @(Get-Content -LiteralPath $nominalLog)
  $bannerStart = [array]::IndexOf($logLines, '================================================================================')
  Check ($bannerStart -gt 0 -and @($logLines[0..($bannerStart - 1)] | Where-Object { $_ -notmatch '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} \[' }).Count -eq 0) 'every log line before the completion banner is timestamped'
  Check (($logLines -join "`n").Contains(' SmartFinOps by WorkplaceCloudHub - Execution completed') -and ($logLines -join "`n").Contains(' Status    : COMPLETED WITH WARNINGS') -and ($logLines -join "`n").Contains(' CSV files : 0')) 'completion banner appended to the report log'
  Check ($run.Log.Contains(' SmartFinOps by WorkplaceCloudHub') -and $run.Log.Contains(' GitHub  : https://github.com/khda79/workplacecloudhub.com') -and $run.Log -match 'Starting SmartFinOps Workplace Report v2\.3\.1 ' -and $run.Log -match 'Tenant context 1.4.1') 'startup banner and versions on the console'
  $unavailable = Invoke-Report (New-ReportFixture 'banner-failed' -ChangeReceipt { param($r) $r.Status = 'Failed' })
  Check ($unavailable.ExitCode -eq 1 -and $unavailable.Log.Contains(' SmartFinOps by WorkplaceCloudHub - Execution failed') -and $unavailable.Log.Contains(' Status    : FAILED')) 'report not available: completion banner Failed'
  Check ($unavailable.Log -match '\[ERROR\] Report not available: ') 'report not available: logged as an error'
  Check ($nmRun.Log -match '\[WARNING\] Licences report not measured in run ' -and $nmRun.Log.Contains(' Status    : COMPLETED WITH WARNINGS')) 'not measured: completed with warnings'
  Check ($run.Log -match '\[WARNING\] Provisional figures in run \S+ for levers L4\.') 'provisional run: warning lists the levers'
  # Orchestrator contract: default success-evidence log pattern, 0 success, 3 completed with warnings, 1 failure.
  $orchestratorEvidence = 'Execution context:|Environment initialized successfully|Starting .* v|(?s:Execution summary:.*Status:\s*(?:Success|CompletedWithWarnings).*Errors:\s*0)'
  $clean = Invoke-Report (New-ReportFixture 'clean' -ChangeReceipt { param($r) $r.Provisional = $false; $r.ProvisionalLevers = @() })
  Check ($clean.ExitCode -eq 0 -and $clean.Log.Contains(' Status    : SUCCESS') -and $clean.Log.Contains(' Warnings  : 0') -and $clean.Log -match $orchestratorEvidence) "final figures: exit 0 and Success banner (got $($clean.ExitCode))"
  Check ($run.Log -match $orchestratorEvidence) 'startup line matches the orchestrator success evidence'
  $engine22 = Invoke-Report (New-ReportFixture 'engine-2.2.0' -ChangeReceipt { param($r) $r.ScriptVersion = '2.2.0' })
  Check ($engine22.ExitCode -eq 3 -and $engine22.Html.Contains('<h1>Workplace FinOps report</h1>')) "engine 2.2.x receipt accepted (exit $($engine22.ExitCode))"
  $missing = Invoke-Report (Join-Path $root 'no-such-tenant\DATA-LAST')
  Check ($missing.ExitCode -eq 1 -and $missing.Log.Contains('SmartM365LatestCsvFolderPath folder not found') -and $missing.Log.Contains('never creates the SmartM365 data roots') -and $missing.Log.Contains(' Status    : FAILED')) 'report error: missing DATA-LAST refused, Failed banner and logged reason'
  Check (@(Get-ChildItem -LiteralPath $nominal -Filter '*.tmp').Count -eq 0) 'no temporary file left'
  $windowsPowerShell = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
  if (Test-Path -LiteralPath $windowsPowerShell) {
    $ps5 = & $windowsPowerShell -NoProfile -ExecutionPolicy Bypass -File $report -Tenant $tenant -ValidateOnly 2>&1 | Out-String
    Check ($LASTEXITCODE -ne 0 -and $ps5 -match '7\.0') 'Windows PowerShell 5.1 is refused'
  }
  #endregion
}
finally {
  if (-not $KeepEvidence -and (Test-Path -LiteralPath $root)) { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }
  elseif ($KeepEvidence) { Write-Host "Evidence kept: $root" }
}

if ($script:failures.Count) { Write-Host ("SmartFinOps report tests: {0} failure(s) / {1} checks" -f $script:failures.Count, $script:checks) -ForegroundColor Red; exit 1 }
Write-Host ("SmartFinOps report tests: PASS ({0} checks)" -f $script:checks) -ForegroundColor Green

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDcwTd9N8MK7R3A
# E2JI61g8orV5rf6zEqz33mBqP6WF66CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIA3czl4Iu5Tmb/3BMnf+YMUvxkd+P2dfGN6WTmNKgvcTMA0GCSqG
# SIb3DQEBAQUABIIBgEoeQf8l+gmayUAk1Icn9M3POt/STJrE0rVu3uw6+hYfFA+W
# i+69w524jMEyoJQE+S2ds0CR8/FecYfMOwpxXF+WKsrSwzh1ZM/epLI8my7JnTnc
# lkPZvIocjPd31q2GklYXJ9awgJ9l8KiATkHYqHdt2nPW9CbZoy1y4HdFBRNt0Ehy
# GLbF0AoWIgsbTQt26dXV9cNpoDDNfDmNdMPnIkJrMwdCls3jOaXS8FWQ3fIbp7UQ
# u7e1d1sOUM8z4YmSFR2/a0bpAWFmy7vgiTrHPSJzdXtninJwGIUVCE9RYTaeB/nX
# 1Iqg/FZJed7rbEogXNnE/OE9tmrU834hFzk+sQxDW7gRxCJ7vLJB0tkeWH3YWAJq
# E336eLIsQ38fPqWpBciXXRRIZqupPP5abfje9JYGETURvrdfVqjq+6nbWkKyROeO
# WzJvTLDBzSKe3WF41hOnm8VUbVXsFhMpsspLtLaLFPjfDAgNhPooxx80r5qEiN9m
# +CxvvzSWQKQgKMSRFKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTkwMDI3
# MzZaMC8GCSqGSIb3DQEJBDEiBCAk419nyP0u6j9mYjUo96OvsMJr+5/LcpmvZB97
# ORfNmjANBgkqhkiG9w0BAQEFAASCAgBsG7i9myhy0WPVRrI+iA69mzu9GRB5TUX4
# 8scq4xFYx42Lbv8BW2iyIT/7ak1bREvBGzhd48uMKkdhb0B2YwY90MNg06gBOn0P
# Z7pkCiFT6Gq26eDQGcgPsYfQzYMGXKETHr0Zk84YbScf4okh3rD66LlEHsTESko8
# EAvIof+kYvdZCTgCDx8N7hr/vG68/M0HGkqbwL76eicjn6cWOm//4munudHvOXAR
# M5vmqWTJEC9R2PhgEIhO+2UQ2n4/RL8OhdDlEBRWBYZi24EP+EafOw7XqP5OSDY/
# bSQ/gX1a/Y9EwcRkvgS7vQt8BX9AGLr6oRILEDTcSHeWXBG6sIgD8wz56vMW16MY
# F1GqP29YDd5Ddo8UCw1mM0NCvPrxHJAB1Am2G6fb8BFyf5H6ieXb2q20sBay+l5F
# YEkJbf74PQw67zTrU5O5OQ2tJvy5eulLY+GDlb5IkXIjLv1p9DrmJWewLvXj4JZ9
# 5AUB2f8xhOSj1745jngd6L/mc7zMPIhNSAMLMQZYxQVci5Qb4X+DI2G6yEgw3w6f
# h6SzHHln1IfL1xe9DntkB5tJpnIox3TuTuIR1WADM302tW60x6XYy9XfeJC+S0Pc
# I/Hua9nGF3CmYOsZu3ZqJQxQXfvjAnnqOSlvZZDAE7GNvf9Ksm5t/+6TPCTXOPe9
# mhbNMev7Qw==
# SIG # End signature block
