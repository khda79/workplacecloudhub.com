#Requires -Version 7.0
<#
.SYNOPSIS
  Offline tests of the SmartFinOps Workplace V2 valuation engine.
.VERSION
1.4.2
.NOTES
  Synthetic contoso data only, in an isolated temporary runtime (no local configuration, export or tenant).
  Source receipts are produced by the real SmartM365.Core module commands.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic Core identity globals are saved and restored in finally.')]
param([switch]$KeepEvidence)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$project = Split-Path $PSScriptRoot -Parent
$repository = Split-Path $project -Parent
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartFinOps-Valuation-' + [guid]::NewGuid().ToString('N'))
$runtime = Join-Path $root 'runtime'
$engine = Join-Path $runtime 'SmartFinOps\SmartFinOps-Workplace-Valuation.ps1'
$pwsh = (Get-Process -Id $PID).Path
# Profile 'prod' (Data\<profile>) whose organization TenantKey, carried by the receipts, is 'contoso-prod'.
$tenant = 'contoso-prod'
$profileKey = 'prod'
$script:checks = 0
$script:failures = [System.Collections.Generic.List[string]]::new()

function Check {
  param([bool]$Condition, [string]$Message)
  $script:checks++
  if (-not $Condition) { $script:failures.Add($Message); Write-Host "FAIL $Message" -ForegroundColor Red }
}

# Isolated runtime: engine, its config helpers and templates, and the real SmartM365.Core module.
foreach ($file in @('SmartFinOps-Workplace-Valuation.ps1','SmartFinOps.global.local.json.template','Config\SmartFinOps-TenantContext.ps1','Config\SmartFinOps-Contracts.local.json.template')) {
  $target = Join-Path $runtime "SmartFinOps\$file"
  New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force | Out-Null
  Copy-Item -LiteralPath (Join-Path $project $file) -Destination $target
}
New-Item -ItemType Directory -Path (Join-Path $runtime 'SmartM365\Modules') -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $repository 'SmartM365\Modules\SmartM365.Core') -Destination (Join-Path $runtime 'SmartM365\Modules') -Recurse
# The TenantKey is fixed in the SmartFinOps configuration (no SmartM365 context in this runtime).
'{ "SmartM365TenantKey": "contoso-prod", "EnableSharePointUpload": false }' | Set-Content -LiteralPath (Join-Path $runtime 'SmartFinOps\SmartFinOps.global.local.json') -Encoding utf8NoBOM
$coreManifest = Join-Path $repository 'SmartM365\Modules\SmartM365.Core\SmartM365.Core.psd1'
Import-Module $coreManifest -MinimumVersion '1.0.90' -Force | Out-Null

$scriptPaths = @{
  Licences = Join-Path $repository 'SmartM365\SmartInventory\M365Inventory\Licensing\SmartM365-Licences-Inventory.ps1'
  Pricing = Join-Path $repository 'SmartM365\SmartInventory\M365Inventory\Licensing\SmartM365-LicensePricing-Inventory.ps1'
  AD = Join-Path $repository 'SmartM365\SmartInventory\ActiveDirectoryInventory\SmartM365-ActiveDirectory-Inventory.ps1'
  Users = Join-Path $repository 'SmartM365\SmartInventory\M365Inventory\Users\SmartM365-ActiveUsers-Inventory.ps1'
}
$scopes = @{ Licences='CMDB:skus,license_paths,plans,user_plans,groups'; AD='CMDB:ad_users,ad_computers,ad_domains,ad_groups,ad_objects,ad_members'; Users='CMDB:users' }

$saved = @{}
foreach ($name in @('SmartM365TenantKey','SmartM365OrganizationKey','SmartM365EnvironmentKey','SmartM365TenantId','csvGeneratedPaths')) {
  $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
  $saved[$name] = @{ Exists=($null -ne $variable); Value=$(if ($variable) { $variable.Value } else { $null }) }
}

#region Fixture
function Write-FixtureCsv {
  param([string]$Folder, [string]$Name, [object[]]$Rows, [switch]$Reverse)
  $path = Join-Path $Folder $Name
  $ordered = if ($Reverse) { [object[]]@($Rows)[($Rows.Count - 1)..0] } else { $Rows }
  $ordered | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding utf8NoBOM
  return $path
}

function New-Receipt {
  param([string]$Folder, [string]$Producer, [string[]]$Paths, [hashtable]$DomainCoverage, [scriptblock]$BeforeComplete)
  $global:csvGeneratedPaths = $null
  if ($Producer -eq 'Pricing') { Start-SmartM365SourceReceipt -ScriptPath $scriptPaths.Pricing -SourceRootPath $Folder }
  else { Start-SmartM365CmdbSourceReceipt -ScriptPath $scriptPaths[$Producer] -SourceRootPath $Folder }
  $publication = Get-SmartM365CurrentSourcePublication
  foreach ($path in $Paths) { Register-SmartM365SourceCsv -Path $path }
  if ($BeforeComplete) { & $BeforeComplete $publication }
  if ($Producer -ne 'Pricing') {
    if ($DomainCoverage) { Set-SmartM365CmdbSourceScope -CompleteScope $false -Scope $scopes[$Producer] -DomainCoverage $DomainCoverage }
    else { Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope $scopes[$Producer] }
  }
  $completed = Complete-SmartM365CmdbSourceReceipt -Status 'Completed'
  if ($completed -notlike '*.current.json.txt') { throw "Fixture receipt failed for $Producer" }
  return $publication.RunId
}

function New-Fixture {
  param([string]$Name, [hashtable]$Options = @{})
  $folder = Join-Path $root "$Name\DATA-LAST"
  New-Item -ItemType Directory -Path $folder -Force | Out-Null
  $t = $tenant; $snapshotId = 'snap-' + $Name
  $reverse = [bool]$Options['Reverse']
  $e3Enabled = if ($Options.ContainsKey('E3Enabled')) { $Options.E3Enabled } else { 10 }
  $e3Consumed = if ($Options.ContainsKey('E3Consumed')) { $Options.E3Consumed } else { 8 }
  $removed = @(if ($Options.ContainsKey('RemovedUsers')) { $Options.RemovedUsers })
  $e3Recovery = @(@('u1','Disabled account'),@('u2','Eligible shared mailbox under 50 GB'),@('u3','No M365 activity in 90 days') | Where-Object { $_[0] -notin $removed })
  $collected = [datetimeoffset]::UtcNow.ToString('o')

  # Licences canonical files (required by its receipt).
  $users = @(
    foreach ($pair in @(@('u1','sku-e3'),@('u2','sku-e3'),@('u3','sku-e3'),@('u4','sku-e3'),@('u5','sku-dyn'),@('u6','sku-e3'),@('u7','sku-visio'),@('u8','sku-e3'),@('u9','sku-f3'),@('u10','sku-cop'),@('u11','sku-cop'))) {
      if ($pair[0] -in $removed) { continue }
      [pscustomobject][ordered]@{ TenantKey=$t; UserId=$pair[0]; SkuId=$pair[1]; SkuPartNumber=''; 'User principal name'="$($pair[0])@contoso.invalid" }
    }
  )
  $licencePaths = @(
    Write-FixtureCsv $folder 'M365_Licenses_Tenant.csv' @([pscustomobject]@{ TenantKey=$t; Id='sku-e3'; TenantSkuPartNumber='SPE_E3' })
    Write-FixtureCsv $folder 'M365_Licenses_AssignmentPaths.csv' @([pscustomobject]@{ TenantKey=$t; UserId='u1'; SkuId='sku-e3' })
    Write-FixtureCsv $folder 'M365_Licenses_Users.csv' $users -Reverse:$reverse
    Write-FixtureCsv $folder 'M365_Licenses_ServicePlans_Catalog.csv' @([pscustomobject]@{ TenantKey=$t; SkuId='sku-e3'; PlanName='EXCHANGE_S_ENTERPRISE' })
    Write-FixtureCsv $folder 'M365_Licenses_UserServicePlanStates.csv' @([pscustomobject]@{ TenantKey=$t; UserId='u1'; StateCode='A' })
    Write-FixtureCsv $folder 'M365_EntraGroups_All.csv' @([pscustomobject]@{ TenantKey=$t; GroupId='g1' })
  )
  $suite = { param($product, $recovery, $wait) [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; GeneratedAtUtc=$collected; Product=$product; Category='Suite'; AssignmentCapacityWait=$wait; RecoveryCandidates=$recovery } }
  $summary = @(
    & $suite 'Microsoft 365 E3' $(if ($Options.ContainsKey('SummaryE3Recovery')) { $Options.SummaryE3Recovery } else { $e3Recovery.Count }) 0
    & $suite 'Microsoft 365 F3' 0 1
    & $suite 'Microsoft 365 E5' 0 0
    & $suite 'Microsoft 365 F1' 0 0
    [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; GeneratedAtUtc=$collected; Product='Microsoft 365 Copilot'; Category='Other paid product'; AssignmentCapacityWait=''; RecoveryCandidates='' }
  )
  $candidates = @(
    foreach ($item in $e3Recovery) { [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; CandidateType='Recovery'; Product='Microsoft 365 E3'; UserId=$item[0]; UserPrincipalName="$($item[0])@contoso.invalid"; Reason=$item[1] } }
    [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; CandidateType=$(if ($Options['UnknownType']) { 'Unknown type' } else { 'E3 to F3 review' }); Product='Microsoft 365 E3'; UserId='u4'; UserPrincipalName='u4@contoso.invalid'; Reason='Manual downgrade review' }
  )
  if ($Options['MixedSnapshot']) { $candidates[0].SnapshotId = 'snap-other' }
  $issues = @([pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; Product='Microsoft 365 F3'; Category='Capacity wait'; UserId='u12' })
  $gaps = @(
    [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; Section='UserMailboxes'; Metric='NoSkus'; Value='4'; EvidenceStatus='Qualified' }
    [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; Section='AdGapActivity'; Metric='NoObservedMatch'; Value=''; EvidenceStatus='N/D' }
  )
  $provisional = [bool]$Options['Provisional']
  $reportSources = @(
    [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; Name='AD_Users_AllDomains.csv'; Ready='True'; Provisional=$(if ($provisional) { 'True' } else { 'False' }); Reason=$(if ($provisional) { 'Qualified partial AD scope' } else { '' }) }
    [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; Name='Exchange_EXO_Mailboxes_AllDomains.csv'; Ready='True'; Provisional='False'; Reason='' }
    [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; Name='M365_OneDrive_Usage.csv'; Ready='True'; Provisional=$(if ($Options['OneDriveProvisional']) { 'True' } else { 'False' }); Reason='' }
  )
  if ($Options['ProvisionalSource']) { $reportSources += [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; Name=$Options.ProvisionalSource; Ready='True'; Provisional='True'; Reason='No receipt' } }
  $sku = { param($user, $skuId, $part, $type, $recommendation, $reason, $guard, $countedIn)
    [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; ReceiptRunId='RUNID'; CandidateType=$type; Recommendation=$recommendation; Reason=$reason; SkuId=$skuId; SkuPartNumber=$part; UserId=$user; UserPrincipalName="$user@contoso.invalid"; MailboxGuard=$guard; DependencyStatus='Clear'; CountedIn=$countedIn } }
  $skuCandidates = @(
    foreach ($item in $e3Recovery) { & $sku $item[0] 'sku-e3' 'SPE_E3' $(if ($item[0] -eq 'u1') { 'Disabled account' } else { 'Report candidate' }) 'Referenced' 'Counted in ReportCandidates:Recovery' 'Not applicable' 'ReportCandidates:Recovery' }
    & $sku 'u4' 'sku-e3' 'SPE_E3' 'Report candidate' 'Referenced' 'Counted' 'Not applicable' 'ReportCandidates:E3 to F3 review'
    & $sku 'u5' 'sku-dyn' 'DYN365_FINANCE' 'Disabled account' 'Recommended' 'Disabled in Entra and AD' 'Not applicable' ''
    & $sku 'u6' 'sku-e3' 'SPE_E3' 'Disabled account' 'Review' 'EXO mailbox source is not qualified' $(if ($Options['CleanMailbox']) { 'Not applicable' } else { 'Unknown: EXO mailbox source is not qualified' }) ''
    & $sku 'u7' 'sku-visio' 'VISIOCLIENT' 'Disabled account' 'Review' 'Entra disabled; AD state is not measured' 'Not applicable' ''
    & $sku 'u8' 'sku-e3' 'SPE_E3' 'Disabled account' 'Blocked by dependency' 'Mailbox retention: litigation hold' 'Blocked: Mailbox retention: litigation hold' ''
    & $sku 'u9' 'sku-f3' 'SPE_F1' 'Disabled account' 'Recommended' 'Disabled in Entra and AD' 'Clear' ''
    & $sku 'u10' 'sku-cop' 'Microsoft_365_Copilot' 'Copilot without activity' 'Recommended' 'No Microsoft 365 Copilot activity in 60 days' 'Not applicable' ''
    & $sku 'u11' 'sku-cop' 'Microsoft_365_Copilot' 'Copilot without activity' 'Review' 'Copilot Chat only' 'Not applicable' ''
  )
  $e3Referenced = @($skuCandidates | Where-Object { $_.SkuId -eq 'sku-e3' -and $_.Recommendation -eq 'Referenced' }).Count
  $cap = { param($skuId, $part, $product, $excluded, $enabled, $consumed, $assigned, $rec, $rev, $blk, $ref, $usage, $notMeasured)
    [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; ReceiptRunId='RUNID'; SkuId=$skuId; SkuPartNumber=$part; SkuDisplayName=$part; ReportProduct=$product; Excluded=$excluded
      EnabledUnits=$enabled; ConsumedUnits=$consumed; AvailableUnits=($enabled - $consumed); AssignedUsers=$assigned; ActiveUsersUsageNotMeasured=$notMeasured
      RecommendedCandidates=$rec; ReviewCandidates=$rev; BlockedCandidates=$blk; ReferencedCandidates=$ref; UsageEvidence=$usage; UsageEvidenceReason=$(if ($usage -eq 'Not measured') { 'No Dynamics 365 product telemetry; active users are a signal only' } else { '' }) } }
  $capacity = @(
    & $cap 'sku-e3' 'SPE_E3' 'Microsoft 365 E3' 'False' $e3Enabled $e3Consumed $e3Consumed 0 1 1 $e3Referenced 'Not applicable' 0
    & $cap 'sku-f3' 'SPE_F1' 'Microsoft 365 F3' 'False' 5 5 5 1 0 0 0 'Not applicable' 0
    & $cap 'sku-cop' 'Microsoft_365_Copilot' 'Microsoft 365 Copilot' 'False' 3 3 3 1 1 0 0 'Measured' 0
    & $cap 'sku-dyn' 'DYN365_FINANCE' 'Dynamics 365' 'False' 2 3 3 1 0 0 0 'Not measured' 2
    & $cap 'sku-free' 'FLOW_FREE' '' 'True' 10000 5 5 0 0 0 0 'Not applicable' 0
    & $cap 'sku-visio' 'VISIOCLIENT' '' 'False' 4 2 2 0 1 0 0 'Not applicable' 0
  )
  if ($Options['G4Mismatch']) { $capacity[0].ReviewCandidates = 5 }
  $reportPaths = [ordered]@{}
  $licRunId = New-Receipt -Folder $folder -Producer 'Licences' -Paths $licencePaths -BeforeComplete {
    param($publication)
    foreach ($row in @($skuCandidates) + @($capacity)) { $row.ReceiptRunId = $publication.RunId }
    if ($Options['BadSkuRunId']) { $capacity[0].ReceiptRunId = 'other-run' }
    $reportPaths['M365_Licenses_ReportSummary.csv'] = Write-FixtureCsv $folder 'M365_Licenses_ReportSummary.csv' $summary -Reverse:$reverse
    $reportPaths['M365_Licenses_ReportCandidates.csv'] = Write-FixtureCsv $folder 'M365_Licenses_ReportCandidates.csv' $candidates -Reverse:$reverse
    $reportPaths['M365_Licenses_ReportAssignmentIssues.csv'] = Write-FixtureCsv $folder 'M365_Licenses_ReportAssignmentIssues.csv' $issues
    $reportPaths['M365_Licenses_ReportGaps.csv'] = Write-FixtureCsv $folder 'M365_Licenses_ReportGaps.csv' $gaps -Reverse:$reverse
    $reportPaths['M365_Licenses_ReportSources.csv'] = Write-FixtureCsv $folder 'M365_Licenses_ReportSources.csv' $reportSources
    $reportPaths['M365_Licenses_ReportSkuCandidates.csv'] = Write-FixtureCsv $folder 'M365_Licenses_ReportSkuCandidates.csv' $skuCandidates -Reverse:$reverse
    if (-not $Options['NoSkuCapacity']) { $reportPaths['M365_Licenses_ReportSkuCapacity.csv'] = Write-FixtureCsv $folder 'M365_Licenses_ReportSkuCapacity.csv' $capacity -Reverse:$reverse }
    foreach ($path in $reportPaths.Values) { Register-SmartM365SourceCsv -Path $path }
    $snapshot = [ordered]@{ SchemaVersion=2; LogicVersion='1.44'; SnapshotId=$snapshotId; TenantKey=$t; GeneratedAtUtc=$collected
      ReceiptRun=[ordered]@{ Producer='SmartM365-Licences-Inventory.ps1'; RunId=$(if ($Options['BadSnapshotRun']) { 'other-run' } else { $publication.RunId }); StartedAtUtc=$publication.StartedAtUtc } }
    $snapshot | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $folder 'M365_Licenses_ReportSnapshot.json.txt') -Encoding utf8NoBOM
  }
  if ($Options['NoSkuCapacity']) {
    # An older decision file left in DATA-LAST that the current receipt declares unpublished.
    Write-FixtureCsv $folder 'M365_Licenses_ReportSkuCapacity.csv' $capacity | Out-Null
  }

  if (-not $Options['NoPricing']) {
    $prices = @(
      [pscustomobject][ordered]@{ TenantKey=$t; SkuPartNumber='SPE_E3'; Currency='EUR'; UnitPriceMonthly='30.00'; EffectiveFrom='2020-01-01'; EffectiveTo='2026-01-01'; PriceSource='Old price'; IsEstimated='True' }
      [pscustomobject][ordered]@{ TenantKey=$t; SkuPartNumber='SPE_E3'; Currency='EUR'; UnitPriceMonthly='37.78'; EffectiveFrom='2026-01-01'; EffectiveTo=''; PriceSource='Microsoft France public list price'; IsEstimated='True' }
      [pscustomobject][ordered]@{ TenantKey=$t; SkuPartNumber='SPE_F1'; Currency='EUR'; UnitPriceMonthly='8,66'; EffectiveFrom='2026-01-01'; EffectiveTo=''; PriceSource='Microsoft France public list price'; IsEstimated='True' }
      [pscustomobject][ordered]@{ TenantKey=$t; SkuPartNumber='Microsoft_365_Copilot'; Currency='EUR'; UnitPriceMonthly='26.00'; EffectiveFrom='2026-01-01'; EffectiveTo=''; PriceSource='Contract'; IsEstimated='False' }
      [pscustomobject][ordered]@{ TenantKey=$t; SkuPartNumber='DYN365_FINANCE'; Currency='EUR'; UnitPriceMonthly='182.00'; EffectiveFrom='2026-01-01'; EffectiveTo=''; PriceSource='Microsoft France public list price'; IsEstimated='True' }
      [pscustomobject][ordered]@{ TenantKey=$t; SkuPartNumber='VISIOCLIENT'; Currency='USD'; UnitPriceMonthly='15.00'; EffectiveFrom='2026-01-01'; EffectiveTo=''; PriceSource='USD only'; IsEstimated='True' }
    )
    $pricePath = Write-FixtureCsv $folder 'M365_License_Prices.csv' $prices
    New-Receipt -Folder $folder -Producer 'Pricing' -Paths @($pricePath) | Out-Null
  }

  $activeUsers = @(foreach ($index in 1..11) { [pscustomobject][ordered]@{ TenantKey=$t; 'Object Id'="u$index"; 'User principal name'="u$index@contoso.invalid"; OnPremisesImmutableId="imm$index"; Department=$(if ($index % 2) { 'Finance' } else { 'Care' }) } })
  New-Receipt -Folder $folder -Producer 'Users' -Paths @(Write-FixtureCsv $folder 'M365_Users_Active.csv' $activeUsers) | Out-Null
  $adUsers = @(foreach ($index in 1..11) { if ($index -eq 7) { continue }; [pscustomobject][ordered]@{ TenantKey=$t; DomainName=$(if ($index -le 5) { 'north.contoso.invalid' } else { 'south.contoso.invalid' }); ImmutableId_AD="imm$index"; UserPrincipalName="u$index@contoso.invalid"; Country=$(if ($index -le 5) { 'FR' } else { 'ES' }); EntityLabel="Site $($index % 3)" } })
  $adPaths = @(Write-FixtureCsv $folder 'AD_Users_AllDomains.csv' $adUsers)
  foreach ($name in @('AD_Computers_AllDomains.csv','AD_Domains_AllDomains.csv','AD_Groups_AllDomains.csv','AD_DirectoryObjects_AllDomains.csv','AD_GroupMemberships_AllDomains.csv')) { $adPaths += Write-FixtureCsv $folder $name @([pscustomobject]@{ TenantKey=$t; Name='x' }) }
  $coverage = if ($Options['PartialAd']) { @{ ExpectedDomains=@('north','south','west'); CollectedDomains=@('north','south'); UnavailableDomains=@('west'); NonBlockingDomainErrors=@('west') } } else { $null }
  New-Receipt -Folder $folder -Producer 'AD' -Paths $adPaths -DomainCoverage $coverage | Out-Null
  return [pscustomobject]@{ Folder=$folder; LicencesRunId=$licRunId; SnapshotId=$snapshotId }
}

function Set-ReceiptField {
  param([string]$Folder, [string]$Receipt, [scriptblock]$Change)
  $path = Join-Path $Folder $Receipt
  $document = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json -AsHashtable
  & $Change $document
  $document | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $path -Encoding utf8NoBOM
}

function New-Contracts {
  param([string]$Path, [object[]]$Skus)
  $document = [ordered]@{ SchemaVersion=1; Defaults=[ordered]@{ ReserveUnitsPercent=0; ReductionPolicy='Unknown' }
    Contracts=@([ordered]@{ ContractKey='EA-CONTOSO'; Type='EA'; Reference=''; StartDate='2024-04-01'; RenewalDate='2027-03-31'; NoticeDays=90; ReductionPolicy='AtRenewal'; Skus=$Skus }); UnmappedSkus=@() }
  New-Item -ItemType Directory -Path (Split-Path $Path -Parent) -Force | Out-Null
  $document | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $Path -Encoding utf8NoBOM
}

function Invoke-Engine {
  param([string]$DataLast, [string]$Output, [string]$Contracts, [string]$AsOfUtc, [switch]$ValidateOnly, [string[]]$Extra = @())
  # $Output is a tenant data root (DATA-LAST, DATA-ALL, LOG-ALL): the fixture sources are copied into its DATA-LAST,
  # where the engine also publishes its current outputs, exactly as in the SmartM365 data tree.
  if ($DataLast.Contains('{{')) { $last = $DataLast; $all = Join-Path $Output 'DATA-ALL'; $logs = Join-Path $Output 'LOG-ALL' }
  else {
    $last = Join-Path $Output 'DATA-LAST'; $all = Join-Path $Output 'DATA-ALL'; $logs = Join-Path $Output 'LOG-ALL'
    New-Item -ItemType Directory -Path $last, $all, $logs -Force | Out-Null
    Get-ChildItem -LiteralPath $DataLast -File | Copy-Item -Destination $last -Force
  }
  $arguments = @('-NoProfile','-ExecutionPolicy','Bypass','-File',$engine,'-Tenant',$profileKey,'-SmartM365LatestCsvFolderPath',$last,'-SmartM365DataAllRootPath',$all,'-SmartM365LogAllRootPath',$logs,'-ContractsPath',$Contracts)
  if ($AsOfUtc) { $arguments += @('-AsOfUtc', $AsOfUtc) }
  if ($ValidateOnly) { $arguments += '-ValidateOnly' }
  $arguments += $Extra
  $log = & $pwsh @arguments 2>&1 | Out-String
  return [pscustomobject]@{ ExitCode=$LASTEXITCODE; Log=$log }
}

function Read-Output { param([string]$Output, [string]$Name) return @(Import-Csv -LiteralPath (Join-Path $Output "DATA-LAST\$Name") -Encoding utf8) }
function Read-Receipt { param([string]$Output) return (Get-Content -LiteralPath (Join-Path $Output 'DATA-LAST\SmartFinOps_Workplace-Valuation.current.json.txt') -Raw | ConvertFrom-Json) }
function Get-HistoryPath { param([string]$Output, [string]$RunId, [string]$Name) return (Join-Path $Output "DATA-ALL\FinOps\Workplace\$RunId\$Name") }
function Get-RunLogPath { param([string]$Output, $Receipt) return (Join-Path $Output ([string]$Receipt.LogFile).Replace('/', '\')) }
function Get-Opportunity { param([object[]]$Rows, [string]$Key) return @($Rows | Where-Object OpportunityKey -eq $Key) }
function Get-Kpi { param([object[]]$Rows, [string]$Key) return @($Rows | Where-Object KpiKey -eq $Key)[0] }
#endregion

try {
  $global:SmartM365TenantKey = $tenant; $global:SmartM365OrganizationKey = 'contoso'; $global:SmartM365EnvironmentKey = 'prod'; $global:SmartM365TenantId = 'synthetic-tenant'
  $baseSkus = @([ordered]@{ SkuPartNumber='SPE_E3'; MinimumUnits=$null; ReserveUnits=$null }, [ordered]@{ SkuPartNumber='SPE_F1'; MinimumUnits=0; ReserveUnits=$null })

  #region 1. Nominal run: classes, statuses, amounts, contracts, outputs
  $base = New-Fixture 'base'
  $contracts = Join-Path $root 'base\contracts.local.json'
  New-Contracts $contracts $baseSkus
  $out = Join-Path $root 'base\out'
  $run = Invoke-Engine -DataLast $base.Folder -Output $out -Contracts $contracts
  Check ($run.ExitCode -eq 3) "nominal run exit code 3, L5 mailbox not measured (got $($run.ExitCode)) $($run.Log)"
  $receipt = Read-Receipt $out
  Check ($receipt.Status -eq 'Completed' -and $receipt.LicencesMeasured -and $receipt.SnapshotId -eq $base.SnapshotId -and $receipt.LicencesReceiptRunId -eq $base.LicencesRunId) 'nominal receipt Completed with snapshot and Licences RunId'
  Check ($receipt.ContractVersion -eq '1.1' -and $receipt.ProfileKey -ceq $profileKey -and $receipt.TenantKey -ceq $tenant) 'receipt 1.1 carries the profile key and the organization TenantKey'
  Check ($run.Log.Contains("Profile 'prod' - TenantKey 'contoso-prod'")) 'profile and TenantKey logged at startup'
  foreach ($file in $receipt.Files) {
    Check ((Get-FileHash -LiteralPath (Join-Path $out "DATA-LAST\$($file.File)") -Algorithm SHA256).Hash -eq $file.SHA256) "receipt hash matches $($file.File) in DATA-LAST"
    Check ((Get-FileHash -LiteralPath (Get-HistoryPath $out $receipt.RunId $file.File) -Algorithm SHA256).Hash -eq $file.SHA256) "history copy in DATA-ALL\FinOps\Workplace matches $($file.File)"
  }
  Check (([string]$receipt.LogFile) -cmatch '^LOG-ALL/SmartFinOps-Workplace-Valuation/SmartFinOps-Workplace-Valuation-\d{4}-\d{2}-\d{2}_\d{2}-\d{2}-\d{2}\.log$' -and (Test-Path -LiteralPath (Get-RunLogPath $out $receipt))) "run log in LOG-ALL\SmartFinOps-Workplace-Valuation, named by the receipt ($($receipt.LogFile))"
  Check (-not (Test-Path -LiteralPath (Get-HistoryPath $out $receipt.RunId 'SmartFinOps-Workplace-Valuation.log')) -and @(Get-ChildItem -LiteralPath (Join-Path $out 'DATA-LAST') -Filter '*.log').Count -eq 0) 'no log in DATA-LAST or in the run history'
  $logLines = @(Get-Content -LiteralPath (Get-RunLogPath $out $receipt))
  $bannerStart = [array]::IndexOf($logLines, '================================================================================')
  Check ($bannerStart -gt 0 -and @($logLines[0..($bannerStart - 1)] | Where-Object { $_ -notmatch '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} \[' }).Count -eq 0) 'every log line before the completion banner is timestamped'
  $banner = ($logLines[$bannerStart..($logLines.Count - 1)] -join "`n")
  Check ($banner.Contains(' SmartFinOps by WorkplaceCloudHub - Execution completed') -and $banner.Contains(' Script    : SmartFinOps-Workplace-Valuation.ps1') -and $banner.Contains(' CSV files : 7') -and $banner -match ' Duration  : \d{2}:\d{2}:\d{2}' -and $banner -match ' Status    : (SUCCESS|COMPLETED WITH WARNINGS)') "completion banner appended to the run log ($banner)"
  Check ($run.Log.Contains(' SmartFinOps by WorkplaceCloudHub') -and $run.Log.Contains(' Website : https://workplacecloudhub.com') -and $run.Log.Contains(' Errors    : 0')) 'startup and completion banners on the console'
  Check ($run.Log -match 'Starting SmartFinOps Workplace Valuation v2\.2\.2 ' -and $run.Log -match 'Tenant context 1.4.1') 'engine and context versions logged at runtime'
  # Orchestrator contract: default success-evidence log pattern, 0 success, 3 completed with warnings.
  $orchestratorEvidence = 'Execution context:|Environment initialized successfully|Starting .* v|(?s:Execution summary:.*Status:\s*(?:Success|CompletedWithWarnings).*Errors:\s*0)'
  Check ($run.Log -match $orchestratorEvidence -and $banner.Contains(' Status    : COMPLETED WITH WARNINGS') -and $banner.Contains(' SmartFinOps by WorkplaceCloudHub - Execution completed with warnings')) 'warning run: exit 3, Completed with warnings banner, orchestrator success evidence'
  $clean = New-Fixture 'clean' @{ CleanMailbox=$true }
  $cleanContracts = Join-Path $root 'clean\contracts.local.json'
  New-Contracts $cleanContracts $baseSkus
  $cleanRun = Invoke-Engine -DataLast $clean.Folder -Output (Join-Path $root 'clean\out') -Contracts $cleanContracts
  Check ($cleanRun.ExitCode -eq 0 -and $cleanRun.Log.Contains(' Status    : SUCCESS') -and $cleanRun.Log.Contains(' Warnings  : 0') -and $cleanRun.Log -match $orchestratorEvidence) "run without warning: exit 0 and Success banner (got $($cleanRun.ExitCode))"
  $expectedHeaders = @{
    'FinOps_Sources.csv'='TenantKey,FinOpsRunId,SnapshotId,ReceiptRunId,SourceFile,Producer,ReceiptFile,SourceReceiptRunId,ReceiptStatus,ReceiptCompletedAtUtc,Required,PublishedInReceipt,ExpectedSHA256,ActualSHA256,ExpectedRows,ActualRows,FreshnessThresholdHours,AgeHours,QualificationStatus,Reason,MissingDomains,ProvisionalLevers'
    'FinOps_Opportunities.csv'='TenantKey,FinOpsRunId,SnapshotId,ReceiptRunId,OpportunityKey,Lever,ValueClass,ValuationStatus,StatusReason,SourceFile,SourceCandidateType,SourceRecommendation,SourceReason,MailboxGuard,DependencyStatus,CountedIn,ReportProduct,SkuId,SkuPartNumber,TargetSkuPartNumber,UserId,UserPrincipalName,Units,UnitPriceMonthly,TargetUnitPriceMonthly,MonthlyAmount,AnnualAmount,Currency,PriceRowKey,PricingReceiptRunId,PriceSource,PriceIsEstimated,CapRank,ContractKey,RenewalDate,RealizableFrom,AdDomain,Country,Entity,Department,ShowbackJoinStatus'
    'FinOps_SkuValuation.csv'='TenantKey,FinOpsRunId,SnapshotId,ReceiptRunId,SkuId,SkuPartNumber,SkuDisplayName,ReportProduct,Excluded,EnabledUnits,ConsumedUnits,AvailableUnits,ReserveUnits,MinimumUnits,CandidateUnitsR,PoolUnits,ReusedUnits,ReducibleUnits,CapUnits,RealizableUnits,CapacityRealizableUnits,CappedUnits,ReviewUnits,NotMeasuredUnits,BlockedUnits,ReferencedUnits,UnitPriceMonthly,PriceRowKey,PriceIsEstimated,RealizableAnnualAmount,AvoidedAnnualAmount,ValuationStatus,ContractKey,ContractStatus,ReductionPolicy,RenewalDate,RealizableFrom,UsageEvidence'
    'FinOps_Showback.csv'='TenantKey,FinOpsRunId,SnapshotId,ReceiptRunId,Dimension,DimensionValue,Lever,ValueClass,ValuationStatus,Users,Units,AnnualAmount,Currency,Provisional'
    'FinOps_Summary.csv'='TenantKey,FinOpsRunId,SnapshotId,ReceiptRunId,KpiKey,ValueClass,Lever,ValuationStatus,Users,Units,AnnualAmount,Currency,Provisional,MissingDomains,SnapshotGeneratedAtUtc,PricingReceiptRunId,Note'
    'FinOps_Realization.csv'='TenantKey,FinOpsRunId,SnapshotId,ReceiptRunId,SkuId,SkuPartNumber,BaselineFinOpsRunId,BaselineSnapshotId,EnabledUnitsBefore,EnabledUnitsAfter,Decrease,RemovedCandidateLicenses,BaselineCapacityUnits,AttributableUnits,RealizedUnits,NotAttributedUnits,UnitPriceMonthly,PriceRowKey,RealizedAnnualAmount,Status'
    'FinOps_Controls.csv'='TenantKey,FinOpsRunId,SnapshotId,ReceiptRunId,Control,Key,SourceCount,FinOpsCount,Status,Detail'
  }
  foreach ($name in $expectedHeaders.Keys) {
    $header = (Get-Content -LiteralPath (Join-Path $out "DATA-LAST\$name") -TotalCount 1).Replace('"','')
    Check ($header -ceq $expectedHeaders[$name]) "exact header $name"
  }
  $opps = Read-Output $out 'FinOps_Opportunities.csv'
  Check (@($opps | Where-Object { $_.TenantKey -ne $tenant -or $_.SnapshotId -ne $base.SnapshotId -or $_.ReceiptRunId -ne $base.LicencesRunId -or $_.FinOpsRunId -ne $receipt.RunId }).Count -eq 0) 'every row carries TenantKey, FinOpsRunId, SnapshotId and ReceiptRunId'
  $expect = @(
    @('L1|u1|sku-e3','R','Valued','453.36'), @('L2|u2|sku-e3','R','Valued','453.36'), @('L3|u3|sku-e3','R','Valued','453.36'),
    @('L4|u4|sku-e3','Rc','Valued','349.44'), @('L5|u5|sku-dyn','R','Valued','2184.00'), @('L5|u1|sku-e3','R','Referenced',''),
    @('L0|u2|sku-e3','R','Referenced',''), @('L5|u6|sku-e3','S','Not measured',''), @('L5|u7|sku-visio','S','Review',''),
    @('L5|u8|sku-e3','S','Blocked',''), @('L5|u9|sku-f3','R','Reused',''), @('L6|u10|sku-cop','R','Valued','312.00'), @('L6|u11|sku-cop','S','Review',''),
    @('L7|sku-e3|Realizable','R','Valued','906.72'), @('L7|sku-visio|Realizable','R','Not valued',''), @('L8|sku-f3|Reuse','E','Valued','103.92'),
    @('L9|sku-dyn|OverAssigned','S','Signal',''), @('L10|sku-dyn|UsageNotMeasured','S','Not measured',''), @('L13|WIN10_ESU','E','Not measured',''),
    @('L11|Microsoft 365 F3|Capacity wait','S','Signal',''), @('L12|AdGapActivity|NoObservedMatch','S','Not measured',''), @('L12|UserMailboxes|NoSkus','S','Signal','')
  )
  foreach ($item in $expect) {
    $row = Get-Opportunity $opps $item[0]
    Check ($row.Count -eq 1) "opportunity $($item[0]) present once"
    if ($row.Count -eq 1) { Check ($row[0].ValueClass -ceq $item[1] -and $row[0].ValuationStatus -ceq $item[2] -and $row[0].AnnualAmount -ceq $item[3]) "opportunity $($item[0]) = $($item[1])/$($item[2])/$($item[3]) (got $($row[0].ValueClass)/$($row[0].ValuationStatus)/$($row[0].AnnualAmount))" }
  }
  Check (@($opps | Where-Object { $_.ValuationStatus -notin @('Valued','Provisional') -and $_.AnnualAmount -ne '' }).Count -eq 0) 'amounts only on Valued/Provisional rows'
  Check (@($opps | Where-Object { $_.AnnualAmount -eq '0.00' }).Count -eq 0) 'no zero amount'
  $l1 = (Get-Opportunity $opps 'L1|u1|sku-e3')[0]
  Check ($l1.PriceRowKey -eq 'SPE_E3|EUR|2026-01-01|' -and $l1.PricingReceiptRunId -and $l1.PriceIsEstimated -eq 'True' -and $l1.ContractKey -eq 'EA-CONTOSO' -and $l1.RealizableFrom -eq '2027-03-31') 'price row traced and contract terms on R row'
  Check ($l1.AdDomain -eq 'north.contoso.invalid' -and $l1.Country -eq 'FR' -and $l1.Department -eq 'Finance' -and $l1.ShowbackJoinStatus -eq 'ImmutableId') 'showback joined by immutable ID'
  $u7 = (Get-Opportunity $opps 'L5|u7|sku-visio')[0]
  Check ($u7.AdDomain -eq 'Not attributed' -and $u7.Department -eq 'Finance') 'unmatched AD user is Not attributed when AD is complete'
  Check ((Get-Opportunity $opps 'L5|u7|sku-visio')[0].StatusReason -like 'Entra disabled*') 'review reason carried'
  Check ((Get-Opportunity $opps 'L4|u4|sku-e3')[0].TargetSkuPartNumber -eq 'SPE_F1') 'E3 to F3 target SKU from ReportProduct'
  $summary = Read-Output $out 'FinOps_Summary.csv'
  foreach ($pair in @(@('TOTAL_R','4762.80'),@('TOTAL_R_DISABLED_AND_SHARED','3090.72'),@('TOTAL_R_ENABLED_INACTIVE','453.36'),@('TOTAL_R_COPILOT','312.00'),@('TOTAL_R_CAPACITY','906.72'),@('TOTAL_RC_E3_TO_F3','349.44'),@('TOTAL_E_REUSE','103.92'))) {
    $kpi = Get-Kpi $summary $pair[0]
    Check ($kpi.AnnualAmount -eq $pair[1] -and $kpi.ValuationStatus -eq 'Valued') "KPI $($pair[0]) = $($pair[1]) (got $($kpi.AnnualAmount))"
  }
  Check ((Get-Kpi $summary 'NOT_MEASURED_MAILBOX').Units -eq '1' -and (Get-Kpi $summary 'NOT_MEASURED_MAILBOX').AnnualAmount -eq '') 'mailbox Not measured KPI counted without amount'
  Check ((Get-Kpi $summary 'TOTAL_E_ESU').ValuationStatus -eq 'Not measured') 'ESU Not measured in V2.1'
  Check ((Get-Kpi $summary 'GAP|UserMailboxes|NoSkus').Units -eq '4' -and (Get-Kpi $summary 'GAP|AdGapActivity|NoObservedMatch').ValuationStatus -eq 'Not measured') 'one summary row per gap metric with its own value'
  Check (@($summary | Where-Object { $_.Lever -eq 'L12' -and $_.KpiKey -like 'DETAIL*' -and $_.Units -ne '' }).Count -eq 0) 'gap metrics never summed in summary'
  Check (@(Read-Output $out 'FinOps_Showback.csv' | Where-Object { $_.Lever -eq 'L12' -and $_.Units -ne '' }).Count -eq 0) 'gap metrics never summed in showback'
  # 2.1.1 breakdowns: price origin, realizable date, showback value, distinct users per value.
  foreach ($item in @(
      @('PRICE|TOTAL_R|Estimated','6','4450.80',''), @('PRICE|TOTAL_R|Contract','1','312.00',''), @('PRICE|TOTAL_RC_E3_TO_F3|Estimated','1','349.44',''),
      @('FORECAST|TOTAL_R|2027-03-31','5','2266.80',''), @('FORECAST|TOTAL_R|Renewal unknown','2','2496.00',''),
      @('FORECAST|TOTAL_RC_E3_TO_F3|2027-03-31','1','349.44',''), @('FORECAST|TOTAL_E_REUSE|2027-03-31','1','103.92',''),
      @('SHOWBACK|Country|TOTAL_R|FR','4','3544.08','4'), @('SHOWBACK|Country|TOTAL_R|ES','1','312.00','1'), @('SHOWBACK|Country|TOTAL_R|SKU level','2','906.72','0'),
      @('SHOWBACK|Department|TOTAL_R|Finance','3','3090.72','3'), @('SHOWBACK|Department|TOTAL_R_ENABLED_INACTIVE|Finance','1','453.36','1'),
      @('SHOWBACK|Country|USERS|FR','','','5'), @('SHOWBACK|Country|USERS|ES','','','5'), @('SHOWBACK|Country|USERS|Not attributed','','','1'),
      @('SHOWBACK|Department|USERS|Finance','','','6'), @('SHOWBACK|Department|USERS|Care','','','5'))) {
    $kpi = @($summary | Where-Object KpiKey -ceq $item[0])
    Check ($kpi.Count -eq 1 -and $kpi[0].Units -eq $item[1] -and $kpi[0].AnnualAmount -eq $item[2] -and ($item[3] -eq '' -or $kpi[0].Users -eq $item[3])) "summary breakdown $($item[0]) = $($item[1])/$($item[2])/$($item[3]) (got $(if ($kpi.Count) { '{0}/{1}/{2}' -f $kpi[0].Units, $kpi[0].AnnualAmount, $kpi[0].Users } else { 'none' }))"
  }
  Check (@($summary | Where-Object { $_.KpiKey -like 'SHOWBACK|*|USERS|*' -and ($_.Units -ne '' -or $_.AnnualAmount -ne '') }).Count -eq 0) 'distinct-user rows carry no units or amount'
  Check (@($summary | Where-Object { $_.KpiKey -like 'PRICE|TOTAL_R|*' }).Count -eq 2 -and @($summary | Where-Object { $_.KpiKey -like 'SHOWBACK|*|USERS|SKU level' }).Count -eq 0) 'breakdown keys well formed, no user count for SKU-level rows'
  Check ((Get-Opportunity $opps 'L4|u4|sku-e3')[0].RealizableFrom -eq '2027-03-31' -and (Get-Opportunity $opps 'L4|u4|sku-e3')[0].ContractKey -eq 'EA-CONTOSO') 'E3 to F3 carries the contract terms of its SKU'
  $skus = Read-Output $out 'FinOps_SkuValuation.csv'
  $e3 = @($skus | Where-Object SkuId -eq 'sku-e3')[0]
  Check ($e3.PoolUnits -eq '5' -and $e3.RealizableUnits -eq '5' -and $e3.CapacityRealizableUnits -eq '2' -and $e3.CappedUnits -eq '0' -and $e3.ContractStatus -eq 'Commitment unknown' -and $e3.ReserveUnits -eq '0' -and $e3.NotMeasuredUnits -eq '1' -and $e3.BlockedUnits -eq '1' -and $e3.ReferencedUnits -eq '4') 'E3 SKU valuation pool and counters'
  Check ($e3.RealizableAnnualAmount -eq '2266.80') "E3 realizable amount (got $($e3.RealizableAnnualAmount))"
  $f3 = @($skus | Where-Object SkuId -eq 'sku-f3')[0]
  Check ($f3.ReusedUnits -eq '1' -and $f3.AvoidedAnnualAmount -eq '103.92' -and $f3.ContractStatus -eq 'Covered' -and $f3.CapUnits -eq '5') 'F3 reuse served before reduction'
  Check (@($skus | Where-Object SkuId -eq 'sku-free')[0].PoolUnits -eq '0') 'excluded SKU has no capacity lever'
  Check (@($skus | Where-Object SkuId -eq 'sku-dyn')[0].ContractStatus -eq 'No contract' -and @($skus | Where-Object SkuId -eq 'sku-dyn')[0].RealizableFrom -eq 'Renewal unknown') 'SKU without contract: renewal unknown'
  $contractsAfter = Get-Content -LiteralPath $contracts -Raw | ConvertFrom-Json
  Check ((@($contractsAfter.UnmappedSkus) -join ',') -eq 'DYN365_FINANCE,Microsoft_365_Copilot,VISIOCLIENT') "UnmappedSkus completed without invented dates (got $(@($contractsAfter.UnmappedSkus) -join ','))"
  $controls = Read-Output $out 'FinOps_Controls.csv'
  Check (@($controls | Where-Object Status -ne 'Pass').Count -eq 0 -and @($controls | Where-Object Control -eq 'G6').Count -gt 0) 'all equivalence controls pass, G1-G6 present'
  foreach ($control in @('G1','G2','G3','G4','G5','G6')) { Check (@($controls | Where-Object Control -eq $control).Count -gt 0) "control $control present" }
  foreach ($key in @('PRICE|TOTAL_R','FORECAST|TOTAL_RC_E3_TO_F3','SHOWBACK|Country|TOTAL_R','SHOWBACK|Department|TOTAL_E_REUSE','SHOWBACK|Entity|USERS')) {
    Check (@($controls | Where-Object { $_.Control -eq 'G6' -and $_.Key -eq $key -and $_.Status -eq 'Pass' }).Count -eq 1) "G6 reconciles summary breakdown $key"
  }
  $showback = Read-Output $out 'FinOps_Showback.csv'
  Check (@($showback | Where-Object { $_.Dimension -eq 'Department' -and $_.Lever -eq 'L3' -and $_.ValueClass -eq 'R' }).Count -ge 1) 'enabled inactive accounts shown as their own showback lever'
  Check (@($showback | Where-Object { $_.DimensionValue -eq 'SKU level' }).Count -gt 0) 'SKU-level rows allocated to SKU level'
  $sources = Read-Output $out 'FinOps_Sources.csv'
  Check (@($sources | Where-Object { $_.SourceFile -like 'M365_Licenses_Report*' -and $_.QualificationStatus -ne 'Qualified' }).Count -eq 0 -and @($sources | Where-Object SourceFile -eq 'Windows_ESU_Prices.csv')[0].QualificationStatus -eq 'NotMeasured') 'Licences sources qualified; absent ESU Not measured'
  $realization = Read-Output $out 'FinOps_Realization.csv'
  Check (@($realization | Where-Object { $_.Status -notlike 'Not measured: no qualified baseline run' }).Count -eq 0) 'first run: realization not measured (no baseline)'
  #endregion

  #region 2. Determinism: permuted source rows give identical outputs
  $perm = New-Fixture 'permuted' @{ Reverse=$true }
  $contractsPerm = Join-Path $root 'permuted\contracts.local.json'; New-Contracts $contractsPerm $baseSkus
  $outPerm = Join-Path $root 'permuted\out'
  $runPerm = Invoke-Engine -DataLast $perm.Folder -Output $outPerm -Contracts $contractsPerm
  Check ($runPerm.ExitCode -eq 3) 'permuted run exit 3'
  foreach ($name in @('FinOps_Opportunities.csv','FinOps_SkuValuation.csv','FinOps_Showback.csv','FinOps_Summary.csv','FinOps_Controls.csv','FinOps_Realization.csv')) {
    $normalize = { param($path, $fixture) (Get-Content -LiteralPath $path -Raw) -replace '"\d{8}T\d{9}Z-[0-9a-f]{8}"', '"RUN"' -replace [regex]::Escape($fixture.LicencesRunId), 'LICRUN' -replace [regex]::Escape($fixture.SnapshotId), 'SNAP' -replace '"[0-9a-f]{32}"', '"PRICERUN"' }
    Check ((& $normalize (Join-Path $out "DATA-LAST\$name") $base) -ceq (& $normalize (Join-Path $outPerm "DATA-LAST\$name") $perm)) "deterministic sort: $name identical with permuted sources"
  }
  #endregion

  #region 3. ValidateOnly writes nothing
  $outValidate = Join-Path $root 'validate\out'
  $contractsValidate = Join-Path $root 'validate\contracts.local.json'
  $runValidate = Invoke-Engine -DataLast $base.Folder -Output $outValidate -Contracts $contractsValidate -ValidateOnly
  Check ($runValidate.ExitCode -eq 3 -and @(Get-ChildItem -LiteralPath (Join-Path $outValidate 'DATA-LAST') -Filter 'FinOps_*').Count -eq 0 -and -not (Test-Path -LiteralPath (Join-Path $outValidate 'DATA-LAST\SmartFinOps_Workplace-Valuation.current.json.txt')) -and -not (Test-Path -LiteralPath (Join-Path $outValidate 'DATA-ALL\FinOps')) -and @(Get-ChildItem -LiteralPath (Join-Path $outValidate 'LOG-ALL') -Recurse -File).Count -eq 0 -and -not (Test-Path -LiteralPath $contractsValidate)) 'ValidateOnly writes no output, history, log or contracts file'
  #endregion

  #region 4. Contracts file created from template when absent
  $outCreate = Join-Path $root 'create\out'
  $contractsCreate = Join-Path $root 'create\contracts.local.json'
  $runCreate = Invoke-Engine -DataLast $base.Folder -Output $outCreate -Contracts $contractsCreate
  $created = Get-Content -LiteralPath $contractsCreate -Raw | ConvertFrom-Json
  Check ($runCreate.ExitCode -eq 3 -and $created.SchemaVersion -eq 1 -and @($created.UnmappedSkus).Count -eq 5) 'contracts file created from template with UnmappedSkus'
  #endregion

  #region 5. Qualification failures: Licences Not measured, never zero
  $cases = [ordered]@{
    'receipt-failed' = { param($f) Set-ReceiptField $f.Folder 'SmartInventory_SmartM365-Licences-Inventory.current.json.txt' { param($d) $d.Status = 'Failed' } }
    'receipt-tenant' = { param($f) Set-ReceiptField $f.Folder 'SmartInventory_SmartM365-Licences-Inventory.current.json.txt' { param($d) $d.TenantKey = 'fabrikam' } }
    'hash-mismatch' = { param($f) Add-Content -LiteralPath (Join-Path $f.Folder 'M365_Licenses_ReportCandidates.csv') -Value '"snap","contoso","Recovery","Microsoft 365 E3","u99","u99@contoso.invalid","Disabled account"' }
    'publishing' = { param($f) $p = Join-Path $f.Folder 'SmartInventory_SmartM365-Licences-Inventory.run.json.txt'; $d = Get-Content $p -Raw | ConvertFrom-Json -AsHashtable; $d.Status = 'Publishing'; $d | ConvertTo-Json -Depth 8 | Set-Content $p -Encoding utf8NoBOM }
  }
  foreach ($case in $cases.Keys) {
    $fixture = New-Fixture $case
    & $cases[$case] $fixture
    $caseOut = Join-Path $root "$case\out"
    $caseRun = Invoke-Engine -DataLast $fixture.Folder -Output $caseOut -Contracts (Join-Path $root "$case\contracts.local.json")
    $caseReceipt = Read-Receipt $caseOut
    $caseOpps = Read-Output $caseOut 'FinOps_Opportunities.csv'
    $caseSummary = Read-Output $caseOut 'FinOps_Summary.csv'
    Check ($caseRun.ExitCode -eq 3 -and $caseReceipt.Status -eq 'Completed' -and -not $caseReceipt.LicencesMeasured) "$case : run completed with Licences not measured"
    Check (@($caseOpps | Where-Object Lever -ne 'L13').Count -eq 0 -and (Get-Kpi $caseSummary 'TOTAL_R').ValuationStatus -eq 'Not measured' -and (Get-Kpi $caseSummary 'TOTAL_R').AnnualAmount -eq '') "$case : no opportunity, TOTAL_R Not measured without amount"
    Check (@($caseSummary | Where-Object { $_.KpiKey -match '^(PRICE|FORECAST|SHOWBACK)\|' }).Count -eq 0) "$case : no breakdown row when Licences is not measured"
  }
  $flagCases = [ordered]@{ 'optional-unpublished'=@{ NoSkuCapacity=$true }; 'snapshot-run'=@{ BadSnapshotRun=$true }; 'mixed-snapshot'=@{ MixedSnapshot=$true }; 'sku-runid'=@{ BadSkuRunId=$true } }
  foreach ($case in $flagCases.Keys) {
    $fixture = New-Fixture $case $flagCases[$case]
    $caseOut = Join-Path $root "$case\out"
    $caseRun = Invoke-Engine -DataLast $fixture.Folder -Output $caseOut -Contracts (Join-Path $root "$case\contracts.local.json")
    $caseReceipt = Read-Receipt $caseOut
    Check ($caseRun.ExitCode -eq 3 -and -not $caseReceipt.LicencesMeasured -and @(Read-Output $caseOut 'FinOps_Opportunities.csv' | Where-Object Lever -ne 'L13').Count -eq 0) "$case : Licences not measured"
  }
  # The profile key is not the TenantKey: comparing the receipts with it measures nothing.
  $profileOut = Join-Path $root 'profile-as-tenant\out'
  $profileRun = Invoke-Engine -DataLast $base.Folder -Output $profileOut -Contracts (Join-Path $root 'profile-as-tenant\contracts.local.json') -Extra @('-SmartM365TenantKey', $profileKey)
  Check ($profileRun.ExitCode -eq 3 -and -not (Read-Receipt $profileOut).LicencesMeasured -and @(Read-Output $profileOut 'FinOps_Sources.csv' | Where-Object { $_.SourceFile -eq 'M365_Licenses_ReportSummary.csv' -and $_.Reason -eq 'source receipt tenant differs' }).Count -eq 1) 'receipts compared with the TenantKey, not with the profile key'
  Check ($profileRun.Log.Contains("Profile 'prod' - TenantKey 'prod' (parameter)")) '-SmartM365TenantKey overrides the configuration'
  $staleOut = Join-Path $root 'stale\out'
  $staleRun = Invoke-Engine -DataLast $base.Folder -Output $staleOut -Contracts (Join-Path $root 'stale\contracts.local.json') -AsOfUtc ([datetime]::UtcNow.AddHours(49).ToString('o'))
  Check ($staleRun.ExitCode -eq 3 -and -not (Read-Receipt $staleOut).LicencesMeasured -and @(Read-Output $staleOut 'FinOps_Sources.csv' | Where-Object { $_.SourceFile -eq 'M365_Licenses_ReportSummary.csv' -and $_.Reason -like 'source is stale*' }).Count -eq 1) 'stale sources (> 48 h) are Not measured'
  #endregion

  #region 6. Pricing not qualified: Not valued, never zero
  $noPrice = New-Fixture 'no-pricing' @{ NoPricing=$true }
  $noPriceOut = Join-Path $root 'no-pricing\out'
  $noPriceRun = Invoke-Engine -DataLast $noPrice.Folder -Output $noPriceOut -Contracts (Join-Path $root 'no-pricing\contracts.local.json')
  $noPriceOpps = Read-Output $noPriceOut 'FinOps_Opportunities.csv'
  Check ($noPriceRun.ExitCode -eq 3 -and (Get-Opportunity $noPriceOpps 'L1|u1|sku-e3')[0].ValuationStatus -eq 'Not valued' -and (Get-Kpi (Read-Output $noPriceOut 'FinOps_Summary.csv') 'TOTAL_R').AnnualAmount -eq '') 'pricing missing: Not valued, empty totals'
  #endregion

  #region 7. Partial AD: Provisional amounts, missing domains, showback Not measured
  $partial = New-Fixture 'partial' @{ PartialAd=$true; Provisional=$true }
  $partialOut = Join-Path $root 'partial\out'
  $partialRun = Invoke-Engine -DataLast $partial.Folder -Output $partialOut -Contracts (Join-Path $root 'partial\contracts.local.json')
  $partialOpps = Read-Output $partialOut 'FinOps_Opportunities.csv'
  $partialReceipt = Read-Receipt $partialOut
  Check ($partialRun.ExitCode -eq 3 -and $partialReceipt.Provisional -and $partialReceipt.MissingDomains -eq 'west') 'partial AD: run Provisional with missing domain'
  Check ((Get-Opportunity $partialOpps 'L1|u1|sku-e3')[0].ValuationStatus -eq 'Provisional' -and (Get-Opportunity $partialOpps 'L1|u1|sku-e3')[0].AnnualAmount -eq '453.36') 'partial AD: amounts Provisional'
  Check ((Get-Opportunity $partialOpps 'L5|u7|sku-visio')[0].AdDomain -eq 'Not measured') 'partial AD: unmatched user showback Not measured'
  Check ((Get-Kpi (Read-Output $partialOut 'FinOps_Summary.csv') 'TOTAL_R').ValuationStatus -eq 'Provisional') 'partial AD: TOTAL_R Provisional'
  Check ((Get-Opportunity $partialOpps 'L7|sku-e3|Realizable')[0].ValuationStatus -eq 'Valued' -and (Get-Opportunity $partialOpps 'L8|sku-f3|Reuse')[0].ValuationStatus -eq 'Valued') 'partial AD: capacity levers not using AD stay Valued'
  Check (@(Read-Output $partialOut 'FinOps_Sources.csv' | Where-Object SourceFile -eq 'AD_Users_AllDomains.csv')[0].ProvisionalLevers -eq 'L1;L2;L3;L4;L5;L6') 'partial AD: provisional levers traced on the AD source'
  #endregion

  #region 7b. Provisional per lever: OneDrive only affects E3 to F3; an unknown source affects every valued lever
  $oneDrive = New-Fixture 'onedrive' @{ OneDriveProvisional=$true }
  $oneDriveOut = Join-Path $root 'onedrive\out'
  $oneDriveRun = Invoke-Engine -DataLast $oneDrive.Folder -Output $oneDriveOut -Contracts (Join-Path $root 'onedrive\contracts.local.json')
  $oneDriveOpps = Read-Output $oneDriveOut 'FinOps_Opportunities.csv'
  $oneDriveSummary = Read-Output $oneDriveOut 'FinOps_Summary.csv'
  $oneDriveSource = @(Read-Output $oneDriveOut 'FinOps_Sources.csv' | Where-Object SourceFile -eq 'M365_OneDrive_Usage.csv')
  Check ($oneDriveRun.ExitCode -eq 3 -and (Get-Opportunity $oneDriveOpps 'L4|u4|sku-e3')[0].ValuationStatus -eq 'Provisional' -and (Get-Opportunity $oneDriveOpps 'L1|u1|sku-e3')[0].ValuationStatus -eq 'Valued') 'OneDrive provisional: only L4 Provisional'
  Check ((Get-Kpi $oneDriveSummary 'TOTAL_R').ValuationStatus -eq 'Valued' -and (Get-Kpi $oneDriveSummary 'TOTAL_R').Provisional -eq 'False' -and (Get-Kpi $oneDriveSummary 'TOTAL_RC_E3_TO_F3').ValuationStatus -eq 'Provisional') 'OneDrive provisional: TOTAL_R Valued, E3 to F3 Provisional'
  Check ($oneDriveSource.Count -eq 1 -and $oneDriveSource[0].QualificationStatus -eq 'Provisional' -and $oneDriveSource[0].ProvisionalLevers -eq 'L4') 'OneDrive provisional: traced in FinOps_Sources with its lever'
  Check ((@((Read-Receipt $oneDriveOut).ProvisionalLevers) -join ';') -eq 'L4') 'OneDrive provisional: receipt lists the provisional levers'
  $sbRows = @(Read-Output $oneDriveOut 'FinOps_Showback.csv' | Where-Object { $_.Lever -eq 'L4' })
  Check ($sbRows.Count -gt 0 -and @($sbRows | Where-Object Provisional -ne 'True').Count -eq 0 -and @(Read-Output $oneDriveOut 'FinOps_Showback.csv' | Where-Object { $_.Lever -eq 'L1' -and $_.Provisional -ne 'False' }).Count -eq 0) 'OneDrive provisional: showback flag per lever'
  $unknown = New-Fixture 'unknown-provisional' @{ ProvisionalSource='Some_New_Source.csv' }
  $unknownOut = Join-Path $root 'unknown-provisional\out'
  $unknownRun = Invoke-Engine -DataLast $unknown.Folder -Output $unknownOut -Contracts (Join-Path $root 'unknown-provisional\contracts.local.json')
  $unknownOpps = Read-Output $unknownOut 'FinOps_Opportunities.csv'
  Check ($unknownRun.ExitCode -eq 3 -and @($unknownOpps | Where-Object { $_.AnnualAmount -ne '' -and $_.ValuationStatus -ne 'Provisional' }).Count -eq 0 -and @(Read-Output $unknownOut 'FinOps_Sources.csv' | Where-Object SourceFile -eq 'Some_New_Source.csv')[0].ProvisionalLevers -eq 'L1;L2;L3;L4;L5;L6;L7;L8') 'unknown provisional source: every valued lever Provisional (fail-safe)'
  #endregion

  #region 8. Cap and reserve from the contracts file
  $capOut = Join-Path $root 'cap\out'; $capContracts = Join-Path $root 'cap\contracts.local.json'
  New-Contracts $capContracts @([ordered]@{ SkuPartNumber='SPE_E3'; MinimumUnits=6; ReserveUnits=1 })
  $capRun = Invoke-Engine -DataLast $base.Folder -Output $capOut -Contracts $capContracts
  $capOpps = Read-Output $capOut 'FinOps_Opportunities.csv'
  $capE3 = @(Read-Output $capOut 'FinOps_SkuValuation.csv' | Where-Object SkuId -eq 'sku-e3')[0]
  Check ($capRun.ExitCode -eq 3 -and $capE3.ReserveUnits -eq '1' -and $capE3.PoolUnits -eq '4' -and $capE3.CapUnits -eq '4' -and $capE3.RealizableUnits -eq '4' -and $capE3.ContractStatus -eq 'Covered') 'reserve 1 and cap EnabledUnits - MinimumUnits'
  Check ((Get-Opportunity $capOpps 'L7|sku-e3|Realizable')[0].Units -eq '1' -and (Get-Opportunity $capOpps 'L3|u3|sku-e3')[0].ValuationStatus -eq 'Valued') 'reserve reduces capacity units'
  $capOut2 = Join-Path $root 'cap2\out'; $capContracts2 = Join-Path $root 'cap2\contracts.local.json'
  New-Contracts $capContracts2 @([ordered]@{ SkuPartNumber='SPE_E3'; MinimumUnits=7; ReserveUnits=0 })
  $capRun2 = Invoke-Engine -DataLast $base.Folder -Output $capOut2 -Contracts $capContracts2
  $cap2Opps = Read-Output $capOut2 'FinOps_Opportunities.csv'
  Check ($capRun2.ExitCode -eq 3 -and (Get-Opportunity $cap2Opps 'L2|u2|sku-e3')[0].ValuationStatus -eq 'Capped' -and (Get-Opportunity $cap2Opps 'L3|u3|sku-e3')[0].ValuationStatus -eq 'Capped' -and (Get-Opportunity $cap2Opps 'L1|u1|sku-e3')[0].ValuationStatus -eq 'Valued') 'cap: deterministic order capacity, L1, then L2/L3 capped'
  Check ((Get-Kpi (Read-Output $capOut2 'FinOps_Summary.csv') 'TOTAL_R_ENABLED_INACTIVE').AnnualAmount -eq '') 'capped rows carry no amount'
  #endregion

  #region 9. Equivalence guard failure keeps files disqualified
  $guardOut = Join-Path $root 'guard\out'; $guardContracts = Join-Path $root 'guard\contracts.local.json'
  New-Contracts $guardContracts $baseSkus
  $guardOk = Invoke-Engine -DataLast $base.Folder -Output $guardOut -Contracts $guardContracts
  $hashBefore = (Get-FileHash -LiteralPath (Join-Path $guardOut 'DATA-LAST\FinOps_Opportunities.csv')).Hash
  foreach ($case in @(@{ SummaryE3Recovery=2 }, @{ G4Mismatch=$true })) {
    $guardFixture = New-Fixture ('guard-' + [guid]::NewGuid().ToString('N').Substring(0,6)) $case
    $guardRun = Invoke-Engine -DataLast $guardFixture.Folder -Output $guardOut -Contracts $guardContracts -AsOfUtc ([datetime]::UtcNow.AddMinutes(5).ToString('o'))
    $guardReceipt = Read-Receipt $guardOut
    Check ($guardOk.ExitCode -eq 3 -and $guardRun.ExitCode -eq 1 -and $guardReceipt.Status -eq 'Failed') "guard failure ($(@($case.Keys)[0])): exit 1 and current receipt Failed"
    Check ($guardRun.Log.Contains(' SmartFinOps by WorkplaceCloudHub - Execution failed') -and $guardRun.Log.Contains(' Status    : FAILED')) 'guard failure: completion banner Failed'
    Check ((Get-Content -LiteralPath (Get-RunLogPath $guardOut $guardReceipt) -Raw).Contains(' Status    : FAILED')) 'guard failure: Failed banner in the run log'
    Check ((Get-FileHash -LiteralPath (Join-Path $guardOut 'DATA-LAST\FinOps_Opportunities.csv')).Hash -eq $hashBefore) 'guard failure: current files kept (disqualified by receipt)'
    Check (@(Import-Csv -LiteralPath (Get-HistoryPath $guardOut $guardReceipt.RunId 'FinOps_Controls.csv') | Where-Object Status -eq 'Fail').Count -ge 1) 'guard failure: controls kept in the run history'
  }
  $schema = New-Fixture 'schema-drift' @{ UnknownType=$true }
  $schemaOut = Join-Path $root 'schema-drift\out'
  $schemaRun = Invoke-Engine -DataLast $schema.Folder -Output $schemaOut -Contracts (Join-Path $root 'schema-drift\contracts.local.json')
  Check ($schemaRun.ExitCode -eq 1 -and (Read-Receipt $schemaOut).Status -eq 'Failed') 'unknown CandidateType fails the run'
  Check ($schemaRun.Log.Contains(' Status    : FAILED') -and $schemaRun.Log.Contains('Run failed: Unknown CandidateType')) 'run error: Failed banner and logged reason'
  $tokenRun = Invoke-Engine -DataLast (Join-Path $root '{{NoSuchToken}}\DATA-LAST') -Output (Join-Path $root 'token\out') -Contracts (Join-Path $root 'token\contracts.local.json')
  Check ($tokenRun.ExitCode -eq 1 -and $tokenRun.Log.Contains('unresolved token {{NoSuchToken}}') -and $tokenRun.Log.Contains(' Status    : FAILED')) 'a DATA-LAST path with an unresolved token is refused'
  #endregion

  #region 10. Realization against the plateau baseline
  $realOut = Join-Path $root 'real\out'; $realContracts = Join-Path $root 'real\contracts.local.json'
  New-Contracts $realContracts $baseSkus
  $run1 = Invoke-Engine -DataLast $base.Folder -Output $realOut -Contracts $realContracts
  $run1Id = (Read-Receipt $realOut).RunId
  $same = Invoke-Engine -DataLast $base.Folder -Output $realOut -Contracts $realContracts -AsOfUtc ([datetime]::UtcNow.AddMinutes(1).ToString('o'))
  $sameRows = Read-Output $realOut 'FinOps_Realization.csv'
  Check ($same.ExitCode -eq 3 -and @($sameRows | Where-Object { $_.SkuId -eq 'sku-e3' -and $_.Status -eq 'No reduction observed' -and $_.RealizedUnits -eq '0' -and $_.BaselineFinOpsRunId -eq $run1Id }).Count -eq 1) 'no reduction: measured zero against the plateau baseline'
  # u1 license removed, E3 enabled units 10 -> 6: removed 1 + baseline capacity 2 = 3 attributable, 1 not attributed.
  $after = New-Fixture 'after' @{ RemovedUsers=@('u1'); E3Enabled=6; E3Consumed=6 }
  $run2 = Invoke-Engine -DataLast $after.Folder -Output $realOut -Contracts $realContracts -AsOfUtc ([datetime]::UtcNow.AddMinutes(2).ToString('o'))
  $realRow = @(Read-Output $realOut 'FinOps_Realization.csv' | Where-Object SkuId -eq 'sku-e3')[0]
  Check ($run1.ExitCode -eq 3 -and $run2.ExitCode -eq 3) 'realization runs exit 3'
  Check ($realRow.BaselineFinOpsRunId -eq $run1Id -and $realRow.Decrease -eq '4' -and $realRow.RemovedCandidateLicenses -eq '1' -and $realRow.BaselineCapacityUnits -eq '2' -and $realRow.AttributableUnits -eq '3' -and $realRow.RealizedUnits -eq '3' -and $realRow.NotAttributedUnits -eq '1' -and $realRow.RealizedAnnualAmount -eq '1360.08' -and $realRow.Status -eq 'Realized') "realization capped and attributed (got $($realRow | ConvertTo-Json -Compress))"
  #endregion

  #region 11. Guarded retention: only SmartFinOps runs of this TenantKey under DATA-ALL\FinOps\Workplace, never a baseline
  . $engine -Tenant $tenant
  . (Join-Path $runtime 'SmartFinOps\Config\SmartFinOps-TenantContext.ps1')
  $retTenantRoot = Join-Path $root 'retention'
  $retRoot = Join-Path $retTenantRoot 'DATA-ALL\FinOps\Workplace'
  $writeRunReceipt = { param($name, $key) $folder = Join-Path $retRoot $name; New-Item -ItemType Directory -Path $folder -Force | Out-Null
    ([ordered]@{ Owner='SmartFinOps-Receipt'; ContractVersion='1.1'; TenantKey=$key; RunId=$name; Status='Completed' } | ConvertTo-Json) | Set-Content -LiteralPath (Join-Path $folder 'SmartFinOps_Workplace-Valuation.current.json.txt') -Encoding utf8NoBOM }
  foreach ($index in 1..6) { & $writeRunReceipt ('2026010{0}T000000000Z-0000000{0}' -f $index) $tenant }
  & $writeRunReceipt '20251201T000000000Z-000000aa' 'fabrikam-prod'
  New-Item -ItemType Directory -Path (Join-Path $retRoot '20251202T000000000Z-000000bb'), (Join-Path $retRoot 'not-a-run'), (Join-Path $retRoot '20251203T000000000Z-000000CC'), (Join-Path $retRoot '20251204T000000000Z-000000dd-copy') -Force | Out-Null
  & $writeRunReceipt '20251205T000000000Z-000000ee' $tenant
  Set-Content -LiteralPath (Join-Path $retRoot '20251205T000000000Z-000000ee\SmartFinOps_Workplace-Valuation.current.json.txt') -Value ((Get-Content -LiteralPath (Join-Path $retRoot '20251205T000000000Z-000000ee\SmartFinOps_Workplace-Valuation.current.json.txt') -Raw).Replace('20251205T000000000Z-000000ee', '20251205T000000000Z-000000ff')) -Encoding utf8NoBOM
  & $writeRunReceipt 'not-a-run' $tenant
  foreach ($other in @('DATA-LAST', 'LOG-ALL\SmartFinOps-Workplace-Valuation')) { New-Item -ItemType Directory -Path (Join-Path $retTenantRoot "$other\20251206T000000000Z-00000011") -Force | Out-Null }
  $protected = [System.Collections.Generic.HashSet[string]]::new([string[]]@('20260101T000000000Z-00000001'))
  $remoteCalls = [System.Collections.Generic.List[string]]::new()
  $script:LogLines.Clear()
  Invoke-FinOpsArchiveRetention -ArchiveRoot $retRoot -KeepRuns 2 -ProtectedRuns $protected -CurrentRunId '20260106T000000000Z-00000006' -TenantKey $tenant -RemoveRemote { param($folder, $name) $remoteCalls.Add("$name|$folder") }
  $left = @(Get-ChildItem -LiteralPath $retRoot -Directory | ForEach-Object Name | Sort-Object)
  $expectedLeft = @('20251201T000000000Z-000000aa','20251202T000000000Z-000000bb','20251203T000000000Z-000000CC','20251204T000000000Z-000000dd-copy','20251205T000000000Z-000000ee','20260101T000000000Z-00000001','20260105T000000000Z-00000005','20260106T000000000Z-00000006','not-a-run') | Sort-Object
  Check (($left -join ',') -ceq ($expectedLeft -join ',')) "retention removes only old SmartFinOps runs of this TenantKey; baseline, other tenant, missing or foreign receipt and non-conforming folders stay (got $($left -join ','))"
  Check ((@($remoteCalls | ForEach-Object { $_.Split('|')[0] }) | Sort-Object) -join ',' -ceq '20260102T000000000Z-00000002,20260103T000000000Z-00000003,20260104T000000000Z-00000004' -and @($remoteCalls | Where-Object { $_.Split('|')[1] -ceq (Join-Path $retRoot $_.Split('|')[0]) }).Count -eq 3) "SharePoint removal requested for exactly the locally removed run folders (got $(@($remoteCalls) -join ','))"
  $removedLines = @($script:LogLines | Where-Object { $_ -match 'Archive retention: removed run (\S+) \(' })
  Check ($removedLines.Count -eq 3 -and @($removedLines | Where-Object { $_ -match [regex]::Escape($retRoot) }).Count -eq 3) 'each removal logged with its full path'
  Check ((Test-Path -LiteralPath (Join-Path $retTenantRoot 'DATA-LAST\20251206T000000000Z-00000011')) -and (Test-Path -LiteralPath (Join-Path $retTenantRoot 'LOG-ALL\SmartFinOps-Workplace-Valuation\20251206T000000000Z-00000011'))) 'nothing removed in DATA-LAST or LOG-ALL'
  $refused = $false
  try { Invoke-FinOpsArchiveRetention -ArchiveRoot (Join-Path $retTenantRoot 'DATA-LAST') -KeepRuns 0 -ProtectedRuns $protected -CurrentRunId 'x' -TenantKey $tenant } catch { $refused = $_.Exception.Message -like 'Archive retention refused*' }
  Check ($refused -and (Test-Path -LiteralPath (Join-Path $retTenantRoot 'DATA-LAST\20251206T000000000Z-00000011'))) 'retention refused outside DATA-ALL\FinOps\Workplace'
  #endregion

  #region 12. Unit checks: sort, prices, contracts
  $sorted = Get-FinOpsSortedRows -Rows @([pscustomobject]@{ K='b' }, [pscustomobject]@{ K='A' }, [pscustomobject]@{ K='a2' }) -KeyColumns @('K')
  Check ((@($sorted | ForEach-Object K) -join ',') -eq 'A,a2,b') 'ordinal case-insensitive sort'
  $index = New-FinOpsPriceIndex -Source ([pscustomobject]@{ QualificationStatus='Qualified'; SourceReceiptRunId='r'; Rows=@(
    [pscustomobject]@{ SkuPartNumber='X'; Currency='EUR'; UnitPriceMonthly='1'; EffectiveFrom='2026-01-01'; EffectiveTo=''; PriceSource='a'; IsEstimated='True' }
    [pscustomobject]@{ SkuPartNumber='X'; Currency='EUR'; UnitPriceMonthly='2'; EffectiveFrom='2026-06-01'; EffectiveTo=''; PriceSource='b'; IsEstimated='True' }
    [pscustomobject]@{ SkuPartNumber='Y'; Currency='EUR'; UnitPriceMonthly='3'; EffectiveFrom='2027-01-01'; EffectiveTo=''; PriceSource='c'; IsEstimated='True' }) }) -AsOfDate ([datetime]'2026-10-18')
  Check (-not (Get-FinOpsPrice $index 'X').Valued -and (Get-FinOpsPrice $index 'X').Reason -like 'more than one*') 'two effective prices: Not valued'
  Check (-not (Get-FinOpsPrice $index 'Y').Valued) 'future price only: Not valued'
  Check ((ConvertTo-FinOpsDecimal '37,78') -eq [decimal]37.78 -and (ConvertTo-FinOpsDecimal '37.78') -eq [decimal]37.78 -and (ConvertTo-FinOpsDecimal '182') -eq [decimal]182) 'decimal comma and point read as decimal separators'
  Check ($null -eq (ConvertTo-FinOpsDecimal '1,234.56') -and $null -eq (ConvertTo-FinOpsDecimal '1 234,5') -and $null -eq (ConvertTo-FinOpsDecimal '12a')) 'group separators and text rejected'
  $commaIndex = New-FinOpsPriceIndex -Source ([pscustomobject]@{ QualificationStatus='Qualified'; SourceReceiptRunId='r'; Rows=@([pscustomobject]@{ SkuPartNumber='SPE_E3'; Currency='EUR'; UnitPriceMonthly='37,78'; EffectiveFrom='2026-01-01'; EffectiveTo=''; PriceSource='a'; IsEstimated='True' }) }) -AsOfDate ([datetime]'2026-10-18')
  Check ((Get-FinOpsPrice $commaIndex 'SPE_E3').Price -eq [decimal]37.78) 'culture-formatted price 37,78 valued as 37.78'
  Check ((Get-FinOpsPriceOrigin 'True') -eq 'Estimated' -and (Get-FinOpsPriceOrigin 'False') -eq 'Contract' -and (Get-FinOpsPriceOrigin '') -eq 'Unknown') 'price origin labels'
  # G6 (2.1.1) detects a summary breakdown that no longer reconciles with its total.
  $g6Opp = New-FinOpsOpportunity @{ OpportunityKey='L1|g6|s'; Lever='L1'; ValueClass='R'; ValuationStatus='Valued'; UserId='g6'; SkuId='s'; AnnualAmount=[decimal]10; PriceIsEstimated='True'; RealizableFrom='2027-03-31'; AdDomain='d'; Country='c'; Entity='e'; Department='p' }
  $g6Valuation = [pscustomobject]@{ Measured=$true; Opportunities=@($g6Opp); MissingDomains=''; AsOfDate=[datetime]'2026-10-18'; PriceIndex=@{ RunId='r' } }
  $g6Sources = @{}; foreach ($name in @('M365_Licenses_ReportCandidates.csv','M365_Licenses_ReportSkuCandidates.csv','M365_Licenses_ReportSummary.csv','M365_Licenses_ReportSkuCapacity.csv')) { $g6Sources[$name] = [pscustomobject]@{ Rows=@() } }
  $g6Showback = Get-FinOpsShowback -Opportunities @($g6Opp)
  $g6Summary = Get-FinOpsSummary -Valuation $g6Valuation
  Check (@(Get-FinOpsControls -Sources $g6Sources -Valuation $g6Valuation -Showback $g6Showback -SummaryRows $g6Summary | Where-Object Status -eq 'Fail').Count -eq 0) 'G6 unit: consistent breakdowns pass'
  @($g6Summary | Where-Object KpiKey -ceq 'PRICE|TOTAL_R|Estimated')[0].AnnualAmount = [decimal]9
  @($g6Summary | Where-Object KpiKey -ceq 'SHOWBACK|Country|USERS|c')[0].Users = 2
  $g6Failed = @(Get-FinOpsControls -Sources $g6Sources -Valuation $g6Valuation -Showback $g6Showback -SummaryRows $g6Summary | Where-Object Status -eq 'Fail' | ForEach-Object Key)
  Check (($g6Failed -join ',') -eq 'PRICE|TOTAL_R,SHOWBACK|Country|USERS') "G6 unit: tampered breakdowns fail (got $($g6Failed -join ','))"
  $template = Join-Path $project 'Config\SmartFinOps-Contracts.local.json.template'
  foreach ($bad in @(
      @{ Name='date'; Skus=@([ordered]@{ SkuPartNumber='SPE_E3' }); Renewal='31/03/2027' },
      @{ Name='duplicate'; Skus=@([ordered]@{ SkuPartNumber='SPE_E3' }, [ordered]@{ SkuPartNumber='spe_e3' }); Renewal='2027-03-31' },
      @{ Name='negative'; Skus=@([ordered]@{ SkuPartNumber='SPE_E3'; MinimumUnits=-1 }); Renewal='2027-03-31' })) {
    $path = Join-Path $root "contracts-$($bad.Name).json"
    New-Contracts $path $bad.Skus
    if ($bad.Name -eq 'date') { (Get-Content $path -Raw).Replace('2027-03-31', $bad.Renewal) | Set-Content $path -Encoding utf8NoBOM }
    $failed = $false; try { Get-FinOpsContracts -Path $path -TemplatePath $template | Out-Null } catch { $failed = $true }
    Check $failed "invalid contracts rejected: $($bad.Name)"
  }
  #endregion

  #region 13. PowerShell 7 only
  $windowsPowerShell = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
  if (Test-Path -LiteralPath $windowsPowerShell) {
    $ps5 = & $windowsPowerShell -NoProfile -ExecutionPolicy Bypass -File $engine -Tenant $tenant -ValidateOnly 2>&1 | Out-String
    Check ($LASTEXITCODE -ne 0 -and $ps5 -match '7\.0') 'Windows PowerShell 5.1 is refused'
  }
  #endregion
}
finally {
  foreach ($name in $saved.Keys) {
    if ($saved[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $saved[$name].Value } else { Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue }
  }
  if (-not $KeepEvidence -and (Test-Path -LiteralPath $root)) { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }
  elseif ($KeepEvidence) { Write-Host "Evidence kept: $root" }
}

if ($script:failures.Count) { Write-Host ("SmartFinOps valuation tests: {0} failure(s) / {1} checks" -f $script:failures.Count, $script:checks) -ForegroundColor Red; exit 1 }
Write-Host ("SmartFinOps valuation tests: PASS ({0} checks)" -f $script:checks) -ForegroundColor Green

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAHbx63ZlSz6mNp
# PO1ZP8roQsKv1DuHwkbkiYqsi+1v56CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAJ1O5gLclMf5BJEI1cxFm/pvCNcdJqKQT0LnNStF2JlMA0GCSqG
# SIb3DQEBAQUABIIBgJv+htF0fbPCPdsTETJxUM8RlUdysNX0dBP0uz+N8aZPSz4m
# Fy3dySblGrMd279JyGW2SVSFBMr5RGMl5/WzaGCK7+EZATYrTI7xZz3dsVnhPpdj
# 1wscpnw8Q1Bdfe/lD+1Vg6/6MgpKcsbTd06rpxFLjB1oR1Amni+qvp8PfLT9NRnP
# 6w627O52mAQ6qmvqAHS+7dfAYDwG9WhiZtB0/djukU7scCzda289gdUrbh6SKtxm
# ATLumF8Kw7Su9f9QC8KFbyzxbIMpxaLqQpbQlcL028ByP3yd6t/XgAn3bn+ZLE6V
# qjSG7aDxtePpiGRis1qDUmnanPTIiLTc/bKkpCsMv5ZX/4VZ+gk6c05AzXRcyLXS
# u28erKDYnXIp+qW4b1gBbMNzVkBSBl9ajjfgbaU7t7clZCnLzJ/rWRoMWCrtGbs2
# Op3Xiwbz0j8lirC9hPoVgSW86IfulZ2/78QgakHkMAWqbnvqUqohNu8b/AIgxRE5
# zX1o0U8wYwtxRlIaHKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTkwOTMx
# NDNaMC8GCSqGSIb3DQEJBDEiBCDfrsKLv0QqFrj2p/lscjJHxRUPcnOMhtTdxvqz
# sMo84zANBgkqhkiG9w0BAQEFAASCAgBkTcpIfk3oBhHo0zyRfd6J5QyNnVWu8/wk
# BpknzP5hRChbHTRQrdjWITR+3TXwBCQAAOanXVM3shodquM9bGP5ZrZ+Hm22pIQ7
# BMlixAd7mMKsqHK6JPx/n9xB7DEPUL8Rehrx7ivZ9PGYs+tXC0ITyIn3HZQrQGdF
# 6Fn1DxNN3V/o6B7cGIANQwJ8KqMY0yWvhEF62GDbc3pveobb8S91Vkehvyx3B6yc
# f3PgVm5qyRp6ZGt1JE9l4HsIyYYcwHGQc7MzdLq9itKYQFRomD1n9L0RKUGP3Nt6
# D+mWbwLxLMe+4ETkaRbVhQ8VyRECfDPas3Iy9bVmOXb8pFM6jdDqzwp1RusenjDP
# TYZ0JZEOjUOGuGv1OOzsf/X1GdRrdM1hqTC96dAT1vICxhPqrlisfxfHTiw/Onuu
# L0r94/RAqvoN6pvcMoKVAVUb2UR4qn7yig2O4zIJgshTYXF6BklqdGO3ry+CP9+p
# 0rfodJs4Xo5T47xfdyVlfF8RA03XEt+tsJWu7P2ICOR84r8gb2iNH8t6u4qkVihL
# 47My9rU/cBQV/L5F42a8zdFcMs1DBUWbT00P5TT9fP0M4gkzpLBJ8mvIjH76IpNO
# fLcUmGtmVokcSy925VraeJBj6bVK+NBa+jdPzwLqVBaFH2BqWFEVWqXZR/jEQPJe
# HNzTjs7Q5A==
# SIG # End signature block
