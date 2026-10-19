#Requires -Version 7.0
<#
.SYNOPSIS
  Synthetic SmartInventory DATA-LAST for the SmartFinOps Workplace demo report.
.DESCRIPTION
  Dot-sourced by SmartFinOps-Workplace-DemoReport.ps1. New-SmartFinOpsDemoFixture writes the Licences report,
  per-SKU decisions, license prices, active users and Active Directory users of a fictional organization
  (contoso.invalid, 240 users) and publishes them with the real SmartM365.Core source receipt commands, exactly
  as the SmartInventory collectors do. Nothing is read from a tenant, a directory or a local export.
  Stage 'Baseline' is the previous plateau (four more disabled E3 accounts, 186 enabled E3 units); stage 'Current'
  is the day after the reduction (180 units), so the report also shows a realized saving.
  Statuses covered: Valued, Provisional (partial AD scope, domain west not collected), Not valued (USD price only),
  Review, Blocked by dependency, Not measured (mailbox state unknown), Capped (contract minimum), Reused (capacity wait).
.VERSION
1.0.0
.NOTES
  Fictional data only: user IDs u001..u243, UPNs under contoso.invalid, no GUID, no address, no real name.
#>

Set-StrictMode -Version Latest

$script:DemoFixtureVersion = '1.0.0'

function Write-SmartFinOpsDemoCsv {
  param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$Name, [Parameter(Mandatory)][object[]]$Rows)
  $path = Join-Path $Folder $Name
  $Rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding utf8NoBOM
  return $path
}

function Publish-SmartFinOpsDemoReceipt {
  # One SmartInventory producer run: start the receipt, register the files, declare the scope, complete it.
  [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='SmartM365.Core reads csvGeneratedPaths; the generator saves and restores it.')]
  param([Parameter(Mandatory)][string]$Folder, [Parameter(Mandatory)][string]$ScriptPath, [string[]]$Paths, [string]$Scope, [hashtable]$DomainCoverage, [switch]$ConfiguredOutputs, [scriptblock]$BeforeComplete)
  $global:csvGeneratedPaths = $null
  if ($ConfiguredOutputs) { Start-SmartM365SourceReceipt -ScriptPath $ScriptPath -SourceRootPath $Folder }
  else { Start-SmartM365CmdbSourceReceipt -ScriptPath $ScriptPath -SourceRootPath $Folder }
  $publication = Get-SmartM365CurrentSourcePublication
  foreach ($path in $Paths) { Register-SmartM365SourceCsv -Path $path }
  if ($BeforeComplete) { & $BeforeComplete $publication }
  if (-not $ConfiguredOutputs) {
    if ($DomainCoverage) { Set-SmartM365CmdbSourceScope -CompleteScope $false -Scope $Scope -DomainCoverage $DomainCoverage }
    else { Set-SmartM365CmdbSourceScope -CompleteScope $true -Scope $Scope }
  }
  $completed = Complete-SmartM365CmdbSourceReceipt -Status 'Completed'
  if ($completed -notlike '*.current.json.txt') { throw "Demo receipt failed for $([IO.Path]::GetFileName($ScriptPath))." }
  return $publication.RunId
}

function Get-SmartFinOpsDemoUserId { param([int]$Index) return ('u{0:000}' -f $Index) }

function New-SmartFinOpsDemoFixture {
  param(
    [Parameter(Mandatory)][string]$Folder,
    [Parameter(Mandatory)][string]$TenantKey,
    [Parameter(Mandatory)][ValidateSet('Baseline','Current')][string]$Stage,
    [Parameter(Mandatory)][hashtable]$ProducerScripts
  )
  New-Item -ItemType Directory -Path $Folder -Force | Out-Null
  $t = $TenantKey
  $snapshotId = 'demo-snapshot-' + $Stage.ToLowerInvariant()
  $collected = [datetimeoffset]::UtcNow.ToString('o')
  $id = { param([int]$Index) Get-SmartFinOpsDemoUserId $Index }
  $upn = { param([string]$UserId) "$UserId@contoso.invalid" }

  # Population: 240 users in four AD domains (west is not collected by the AD inventory: partial scope).
  $domainOf = { param([int]$Index) if ($Index -le 90) { 'north' } elseif ($Index -le 160) { 'south' } elseif ($Index -le 215) { 'east' } else { 'west' } }
  $countryOf = @{ north='FR'; south='ES'; east='DE'; west='IT' }
  $departments = @('Finance','Care','Sales','IT','Operations','Human Resources')

  # Candidate sets (user indexes). Baseline adds four disabled E3 accounts removed before the current stage.
  $disabled = [System.Collections.Generic.List[int]]::new([int[]]@(8,19,33,47,72,96,118,141,160))
  $removedAtReduction = @(6,17,54,90)
  if ($Stage -eq 'Baseline') { $disabled.AddRange([int[]]$removedAtReduction) }
  $shared = @(12,58,99,133,155)
  $inactive = @(5,15,27,39,52,66,81,93,107,122,137,149,158,166)
  $downgrade = @(2,23,44,61,77,88,113,127,139,146,152,163)
  $present = { param([int]$Index) $Stage -eq 'Baseline' -or $Index -notin $removedAtReduction }
  $e3Users = @(1..168 | Where-Object { & $present $_ })
  $e3Enabled = if ($Stage -eq 'Baseline') { 186 } else { 180 }

  # Licences canonical files (required by the Licences receipt).
  $assignments = [System.Collections.Generic.List[object]]::new()
  $addAssignment = { param([int]$Index, [string]$SkuId) $userId = & $id $Index; $assignments.Add([pscustomobject][ordered]@{ TenantKey=$t; UserId=$userId; SkuId=$SkuId; SkuPartNumber=''; 'User principal name'=(& $upn $userId) }) }
  foreach ($index in $e3Users) { & $addAssignment $index 'sku-e3' }
  foreach ($index in 169..228) { & $addAssignment $index 'sku-f3' }
  foreach ($index in @(1..36 | Where-Object { & $present $_ })) { & $addAssignment $index 'sku-copilot' }
  foreach ($index in 37..45) { & $addAssignment $index 'sku-dynamics' }
  foreach ($index in 46..51) { & $addAssignment $index 'sku-visio' }
  foreach ($index in 1..5) { & $addAssignment $index 'sku-flow-free' }
  $licencePaths = @(
    Write-SmartFinOpsDemoCsv $Folder 'M365_Licenses_Tenant.csv' @(
      [pscustomobject][ordered]@{ TenantKey=$t; Id='sku-e3'; TenantSkuPartNumber='SPE_E3' }
      [pscustomobject][ordered]@{ TenantKey=$t; Id='sku-f3'; TenantSkuPartNumber='SPE_F1' }
      [pscustomobject][ordered]@{ TenantKey=$t; Id='sku-copilot'; TenantSkuPartNumber='Microsoft_365_Copilot' }
      [pscustomobject][ordered]@{ TenantKey=$t; Id='sku-dynamics'; TenantSkuPartNumber='DYN365_FINANCE' }
      [pscustomobject][ordered]@{ TenantKey=$t; Id='sku-visio'; TenantSkuPartNumber='VISIOCLIENT' }
      [pscustomobject][ordered]@{ TenantKey=$t; Id='sku-flow-free'; TenantSkuPartNumber='FLOW_FREE' })
    Write-SmartFinOpsDemoCsv $Folder 'M365_Licenses_AssignmentPaths.csv' @($assignments | ForEach-Object { [pscustomobject][ordered]@{ TenantKey=$t; UserId=$_.UserId; SkuId=$_.SkuId } })
    Write-SmartFinOpsDemoCsv $Folder 'M365_Licenses_Users.csv' @($assignments)
    Write-SmartFinOpsDemoCsv $Folder 'M365_Licenses_ServicePlans_Catalog.csv' @([pscustomobject][ordered]@{ TenantKey=$t; SkuId='sku-e3'; PlanName='EXCHANGE_S_ENTERPRISE' })
    Write-SmartFinOpsDemoCsv $Folder 'M365_Licenses_UserServicePlanStates.csv' @([pscustomobject][ordered]@{ TenantKey=$t; UserId=(& $id 1); StateCode='A' })
    Write-SmartFinOpsDemoCsv $Folder 'M365_EntraGroups_All.csv' @([pscustomobject][ordered]@{ TenantKey=$t; GroupId='group-licensing-e3' })
  )

  # Licences report: suite summary, suite candidates, assignment issues, gap metrics and source states.
  $suite = { param([string]$Product, $Recovery, $Wait) [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; GeneratedAtUtc=$collected; Product=$Product; Category='Suite'; AssignmentCapacityWait=$Wait; RecoveryCandidates=$Recovery } }
  $candidates = [System.Collections.Generic.List[object]]::new()
  $addCandidate = { param([int]$Index, [string]$Type, [string]$Reason) $userId = & $id $Index; $candidates.Add([pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; CandidateType=$Type; Product='Microsoft 365 E3'; UserId=$userId; UserPrincipalName=(& $upn $userId); Reason=$Reason }) }
  foreach ($index in $disabled) { & $addCandidate $index 'Recovery' 'Disabled account' }
  foreach ($index in $shared) { & $addCandidate $index 'Recovery' 'Eligible shared mailbox under 50 GB' }
  foreach ($index in $inactive) { & $addCandidate $index 'Recovery' 'No M365 activity in 90 days' }
  foreach ($index in $downgrade) { & $addCandidate $index 'E3 to F3 review' 'Frontline usage profile' }
  $recoveryCount = @($candidates | Where-Object CandidateType -eq 'Recovery').Count
  $summary = @(
    & $suite 'Microsoft 365 E3' $recoveryCount 0
    & $suite 'Microsoft 365 F3' 0 3
    & $suite 'Microsoft 365 E5' 0 0
    & $suite 'Microsoft 365 F1' 0 0
    [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; GeneratedAtUtc=$collected; Product='Microsoft 365 Copilot'; Category='Other paid product'; AssignmentCapacityWait=''; RecoveryCandidates='' }
  )
  $issues = @(foreach ($index in 241..243) { [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; Product='Microsoft 365 F3'; Category='Capacity wait'; UserId=(& $id $index) } })
  $gaps = @(
    [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; Section='UserMailboxes'; Metric='NoSkus'; Value='4'; EvidenceStatus='Qualified' }
    [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; Section='SharedMailboxes'; Metric='Licensed'; Value='7'; EvidenceStatus='Qualified' }
    [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; Section='AdGapActivity'; Metric='NoObservedMatch'; Value=''; EvidenceStatus='N/D' }
  )
  $reportSources = @(
    [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; Name='AD_Users_AllDomains.csv'; Ready='True'; Provisional='True'; Reason='Qualified partial AD scope: domain west not collected' }
    [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; Name='Exchange_EXO_Mailboxes_AllDomains.csv'; Ready='True'; Provisional='False'; Reason='' }
    [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; Name='M365_OneDrive_Usage.csv'; Ready='True'; Provisional='False'; Reason='' }
  )

  # Per-SKU decisions. Suite candidates are only referenced here (counted once, in the suite levers).
  $skuCandidates = [System.Collections.Generic.List[object]]::new()
  $skuParts = @{ 'sku-e3'='SPE_E3'; 'sku-f3'='SPE_F1'; 'sku-copilot'='Microsoft_365_Copilot'; 'sku-dynamics'='DYN365_FINANCE'; 'sku-visio'='VISIOCLIENT' }
  $addSku = { param([int]$Index, [string]$SkuId, [string]$Type, [string]$Recommendation, [string]$Reason, [string]$Guard, [string]$CountedIn)
    $userId = & $id $Index
    $skuCandidates.Add([pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; ReceiptRunId='RUNID'; CandidateType=$Type; Recommendation=$Recommendation; Reason=$Reason; SkuId=$SkuId; SkuPartNumber=$skuParts[$SkuId]; UserId=$userId; UserPrincipalName=(& $upn $userId); MailboxGuard=$Guard; DependencyStatus='Clear'; CountedIn=$CountedIn }) }
  foreach ($candidate in $candidates) {
    $index = [int]$candidate.UserId.Substring(1)
    $countedIn = if ($candidate.CandidateType -eq 'Recovery') { 'ReportCandidates:Recovery' } else { 'ReportCandidates:E3 to F3 review' }
    $type = if ($candidate.Reason -eq 'Disabled account') { 'Disabled account' } else { 'Report candidate' }
    & $addSku $index 'sku-e3' $type 'Referenced' 'Counted in ReportCandidates' 'Not applicable' $countedIn
  }
  & $addSku 41 'sku-e3' 'Disabled account' 'Blocked by dependency' 'Mailbox retention: litigation hold' 'Blocked: Mailbox retention: litigation hold' ''
  & $addSku 85 'sku-e3' 'Disabled account' 'Blocked by dependency' 'Mailbox retention: litigation hold' 'Blocked: Mailbox retention: litigation hold' ''
  & $addSku 130 'sku-e3' 'Disabled account' 'Review' 'EXO mailbox source is not qualified' 'Unknown: EXO mailbox source is not qualified' ''
  foreach ($index in @(8,19,33) + @(if ($Stage -eq 'Baseline') { 6,17 })) { & $addSku $index 'sku-copilot' 'Disabled account' 'Recommended' 'Disabled in Entra and AD' 'Not applicable' '' }
  foreach ($index in 3,11,21,29,35) { & $addSku $index 'sku-copilot' 'Copilot without activity' 'Recommended' 'No Microsoft 365 Copilot activity in 60 days' 'Not applicable' '' }
  foreach ($index in 14,30) { & $addSku $index 'sku-copilot' 'Copilot without activity' 'Review' 'Copilot Chat only' 'Not applicable' '' }
  foreach ($index in 175,199) { & $addSku $index 'sku-f3' 'Disabled account' 'Recommended' 'Disabled in Entra and AD' 'Clear' '' }
  & $addSku 40 'sku-dynamics' 'Disabled account' 'Recommended' 'Disabled in Entra and AD' 'Not applicable' ''
  & $addSku 48 'sku-visio' 'Disabled account' 'Review' 'Entra disabled; AD state is not measured' 'Not applicable' ''

  # Per-SKU capacity: counters are derived from the decisions above so that the equivalence controls reconcile.
  $assignedBySku = @{}; foreach ($row in $assignments) { $assignedBySku[$row.SkuId] = 1 + [int]$assignedBySku[$row.SkuId] }
  $count = { param([string]$SkuId, [string]$Recommendation) @($skuCandidates | Where-Object { $_.SkuId -eq $SkuId -and $_.Recommendation -eq $Recommendation }).Count }
  $capacityRow = { param([string]$SkuId, [string]$Part, [string]$Product, [string]$Excluded, [int]$Enabled, [string]$Usage, [int]$NotMeasured)
    $consumed = [int]$assignedBySku[$SkuId]
    [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; ReceiptRunId='RUNID'; SkuId=$SkuId; SkuPartNumber=$Part; SkuDisplayName=$Part; ReportProduct=$Product; Excluded=$Excluded
      EnabledUnits=$Enabled; ConsumedUnits=$consumed; AvailableUnits=($Enabled - $consumed); AssignedUsers=$consumed; ActiveUsersUsageNotMeasured=$NotMeasured
      RecommendedCandidates=(& $count $SkuId 'Recommended'); ReviewCandidates=(& $count $SkuId 'Review'); BlockedCandidates=(& $count $SkuId 'Blocked by dependency'); ReferencedCandidates=(& $count $SkuId 'Referenced')
      UsageEvidence=$Usage; UsageEvidenceReason=$(if ($Usage -eq 'Not measured') { 'No Dynamics 365 product telemetry; active users are a signal only' } else { '' }) } }
  $capacity = @(
    & $capacityRow 'sku-e3' 'SPE_E3' 'Microsoft 365 E3' 'False' $e3Enabled 'Not applicable' 0
    & $capacityRow 'sku-f3' 'SPE_F1' 'Microsoft 365 F3' 'False' 60 'Not applicable' 0
    & $capacityRow 'sku-copilot' 'Microsoft_365_Copilot' 'Microsoft 365 Copilot' 'False' 40 'Measured' 0
    & $capacityRow 'sku-dynamics' 'DYN365_FINANCE' 'Dynamics 365' 'False' 8 'Not measured' 9
    & $capacityRow 'sku-visio' 'VISIOCLIENT' '' 'False' 9 'Not applicable' 0
    & $capacityRow 'sku-flow-free' 'FLOW_FREE' '' 'True' 10000 'Not applicable' 0
  )

  $licencesRunId = Publish-SmartFinOpsDemoReceipt -Folder $Folder -ScriptPath $ProducerScripts.Licences -Paths $licencePaths -Scope 'CMDB:skus,license_paths,plans,user_plans,groups' -BeforeComplete {
    param($publication)
    foreach ($row in @($skuCandidates) + @($capacity)) { $row.ReceiptRunId = $publication.RunId }
    $reportPaths = @(
      Write-SmartFinOpsDemoCsv $Folder 'M365_Licenses_ReportSummary.csv' $summary
      Write-SmartFinOpsDemoCsv $Folder 'M365_Licenses_ReportCandidates.csv' @($candidates)
      Write-SmartFinOpsDemoCsv $Folder 'M365_Licenses_ReportAssignmentIssues.csv' $issues
      Write-SmartFinOpsDemoCsv $Folder 'M365_Licenses_ReportGaps.csv' $gaps
      Write-SmartFinOpsDemoCsv $Folder 'M365_Licenses_ReportSources.csv' $reportSources
      Write-SmartFinOpsDemoCsv $Folder 'M365_Licenses_ReportSkuCandidates.csv' @($skuCandidates)
      Write-SmartFinOpsDemoCsv $Folder 'M365_Licenses_ReportSkuCapacity.csv' $capacity
    )
    foreach ($path in $reportPaths) { Register-SmartM365SourceCsv -Path $path }
    $snapshot = [ordered]@{ SchemaVersion=2; LogicVersion='1.44'; SnapshotId=$snapshotId; TenantKey=$t; GeneratedAtUtc=$collected
      ReceiptRun=[ordered]@{ Producer='SmartM365-Licences-Inventory.ps1'; RunId=$publication.RunId; StartedAtUtc=$publication.StartedAtUtc } }
    $snapshot | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $Folder 'M365_Licenses_ReportSnapshot.json.txt') -Encoding utf8NoBOM
  }

  # License prices: public list estimates, one contract price (Copilot), one USD-only SKU (Visio: Not valued).
  $price = { param([string]$Part, [string]$Amount, [string]$From, [string]$To, [string]$Source, [string]$Estimated, [string]$Currency = 'EUR')
    [pscustomobject][ordered]@{ TenantKey=$t; SkuPartNumber=$Part; Currency=$Currency; UnitPriceMonthly=$Amount; EffectiveFrom=$From; EffectiveTo=$To; PriceSource=$Source; IsEstimated=$Estimated } }
  $prices = @(
    & $price 'SPE_E3' '33.75' '2024-01-01' '2026-01-01' 'Former list price' 'True'
    & $price 'SPE_E3' '37.78' '2026-01-01' '' 'Public list price (demo)' 'True'
    & $price 'SPE_F1' '8.66' '2026-01-01' '' 'Public list price (demo)' 'True'
    & $price 'Microsoft_365_Copilot' '26.00' '2026-01-01' '' 'Contract price (demo)' 'False'
    & $price 'DYN365_FINANCE' '182.00' '2026-01-01' '' 'Public list price (demo)' 'True'
    & $price 'VISIOCLIENT' '15.00' '2026-01-01' '' 'USD list price only (demo)' 'True' 'USD'
  )
  Publish-SmartFinOpsDemoReceipt -Folder $Folder -ScriptPath $ProducerScripts.Pricing -Paths @(Write-SmartFinOpsDemoCsv $Folder 'M365_License_Prices.csv' $prices) -ConfiguredOutputs | Out-Null

  # Showback dimensions: department from the active users, domain, country and entity from Active Directory.
  $population = @(1..240 | Where-Object { & $present $_ })
  $activeUsers = @(foreach ($index in $population) { $userId = & $id $index
      [pscustomobject][ordered]@{ TenantKey=$t; 'Object Id'=$userId; 'User principal name'=(& $upn $userId); OnPremisesImmutableId="imm-$userId"; Department=$departments[$index % $departments.Count] } })
  Publish-SmartFinOpsDemoReceipt -Folder $Folder -ScriptPath $ProducerScripts.Users -Paths @(Write-SmartFinOpsDemoCsv $Folder 'M365_Users_Active.csv' $activeUsers) -Scope 'CMDB:users' | Out-Null
  $adUsers = @(foreach ($index in $population) { $domain = & $domainOf $index; if ($domain -eq 'west') { continue }; $userId = & $id $index
      [pscustomobject][ordered]@{ TenantKey=$t; DomainName="$domain.contoso.invalid"; ImmutableId_AD="imm-$userId"; UserPrincipalName=(& $upn $userId); Country=$countryOf[$domain]; EntityLabel=('Site {0}' -f [char](65 + ($index % 5))) } })
  $adPaths = @(Write-SmartFinOpsDemoCsv $Folder 'AD_Users_AllDomains.csv' $adUsers)
  foreach ($name in @('AD_Computers_AllDomains.csv','AD_Domains_AllDomains.csv','AD_Groups_AllDomains.csv','AD_DirectoryObjects_AllDomains.csv','AD_GroupMemberships_AllDomains.csv')) {
    $adPaths += Write-SmartFinOpsDemoCsv $Folder $name @([pscustomobject][ordered]@{ TenantKey=$t; Name='demo' })
  }
  $coverage = @{ ExpectedDomains=@('north','south','east','west'); CollectedDomains=@('north','south','east'); UnavailableDomains=@('west'); NonBlockingDomainErrors=@('west') }
  Publish-SmartFinOpsDemoReceipt -Folder $Folder -ScriptPath $ProducerScripts.AD -Paths $adPaths -Scope 'CMDB:ad_users,ad_computers,ad_domains,ad_groups,ad_objects,ad_members' -DomainCoverage $coverage | Out-Null

  return [pscustomobject]@{ Folder=$Folder; SnapshotId=$snapshotId; LicencesRunId=$licencesRunId; Stage=$Stage }
}

function New-SmartFinOpsDemoContract {
  # Contract terms: an EA renewed in March 2027 (E3 minimum 140 units, Copilot minimum 28) and a monthly CSP agreement.
  param([Parameter(Mandatory)][string]$Path)
  $document = [ordered]@{ SchemaVersion=1; Defaults=[ordered]@{ ReserveUnitsPercent=0; ReductionPolicy='Unknown' }
    Contracts=@(
      [ordered]@{ ContractKey='EA-CONTOSO'; Type='EA'; Reference=''; StartDate='2024-04-01'; RenewalDate='2027-03-31'; NoticeDays=90; ReductionPolicy='AtRenewal'
        Skus=@([ordered]@{ SkuPartNumber='SPE_E3'; MinimumUnits=140; ReserveUnits=2 }, [ordered]@{ SkuPartNumber='SPE_F1'; MinimumUnits=0; ReserveUnits=0 }, [ordered]@{ SkuPartNumber='Microsoft_365_Copilot'; MinimumUnits=28; ReserveUnits=0 }) }
      [ordered]@{ ContractKey='CSP-CONTOSO'; Type='CSP-NCE'; Reference=''; StartDate='2026-01-01'; RenewalDate=''; NoticeDays=$null; ReductionPolicy='Monthly'
        Skus=@([ordered]@{ SkuPartNumber='DYN365_FINANCE'; MinimumUnits=$null; ReserveUnits=$null }) })
    UnmappedSkus=@() }
  New-Item -ItemType Directory -Path (Split-Path $Path -Parent) -Force | Out-Null
  $document | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $Path -Encoding utf8NoBOM
}
