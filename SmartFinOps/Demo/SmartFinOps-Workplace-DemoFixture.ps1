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
  Every source receipt is complete (all four AD domains collected, no provisional Licences report source), so the
  figures are Final (1.1.0). Statuses covered: Valued, Not valued (USD price only), Review (Copilot Chat only, AD state
  not measured, room mailbox), Blocked by dependency, Not measured (Windows 10 ESU, Dynamics 365 usage),
  Capped (contract minimum), Reused (capacity wait).
.VERSION
1.1.0
.NOTES
  Fictional data only: user IDs u001..u243, UPNs under contoso.invalid, no GUID, no address, no real name.
#>

Set-StrictMode -Version Latest

$script:DemoFixtureVersion = '1.1.0'

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
    [pscustomobject][ordered]@{ SnapshotId=$snapshotId; TenantKey=$t; Name='AD_Users_AllDomains.csv'; Ready='True'; Provisional='False'; Reason='' }
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
  & $addSku 130 'sku-e3' 'Disabled account' 'Review' 'Room mailbox: resource account' 'Not applicable' ''
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
  $adUsers = @(foreach ($index in $population) { $domain = & $domainOf $index; $userId = & $id $index
      [pscustomobject][ordered]@{ TenantKey=$t; DomainName="$domain.contoso.invalid"; ImmutableId_AD="imm-$userId"; UserPrincipalName=(& $upn $userId); Country=$countryOf[$domain]; EntityLabel=('Site {0}' -f [char](65 + ($index % 5))) } })
  $adPaths = @(Write-SmartFinOpsDemoCsv $Folder 'AD_Users_AllDomains.csv' $adUsers)
  foreach ($name in @('AD_Computers_AllDomains.csv','AD_Domains_AllDomains.csv','AD_Groups_AllDomains.csv','AD_DirectoryObjects_AllDomains.csv','AD_GroupMemberships_AllDomains.csv')) {
    $adPaths += Write-SmartFinOpsDemoCsv $Folder $name @([pscustomobject][ordered]@{ TenantKey=$t; Name='demo' })
  }
  Publish-SmartFinOpsDemoReceipt -Folder $Folder -ScriptPath $ProducerScripts.AD -Paths $adPaths -Scope 'CMDB:ad_users,ad_computers,ad_domains,ad_groups,ad_objects,ad_members' | Out-Null

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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBNBzXQ+UmkFbvc
# h3EqBJn3UQT87IiYVMYL5z7+gJBcvqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGZdakpS2rQAQDztfmtirFK3odql5BQ2hnDePuWuUIHzMA0GCSqG
# SIb3DQEBAQUABIIBgI4bF7uOTr8aO+H1QOGA/VjDcitmPoqcDyiScAyjZ/eOhXIE
# YK27khNlb4zbgUc7kISQx/LeOXsCvbKJA2lGgYkrVAcuQJP+kj8BA7k+yqHg/25B
# 3OXvzF+TaxS2k60TqjGAoIvsh4Dm40mv+UoOd752tCyg/eBmWN3YBvF7WM429nBi
# D6TW+c7pkqtsyM9rOvaIwtaMp9znDAwzJnArcYN5692HnnkInmt0q5h/1q+7zf9V
# fPz1XvjeL6320NHDulSjWUK1xQFEGTOCmS+v7vQFCYb2MVf4cQ5tx3ODz9MIowuM
# YRoyTeMfPdBxhvL5f7ZW43q4OuJmJqNsWi3NuwmphBxrAR2BIQfNLkCRjPAV5FnS
# pXOOlEGg1Jc5g2kBAtH5iijyZE5OpgVKeOfEHQdLJmVX/d9iZajtVm0t3YoUqro+
# oFzIKl6yNBvXIiSXJtwiwzAgApkGutJ874RfP/qaRBeS+uXMZ+s7Zgcf8dWpGlrg
# LuIk4xj+OGrClhAkbKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTkxMjUz
# MDZaMC8GCSqGSIb3DQEJBDEiBCAHeIzGkFIXnXC2Yqw8feTdgmXgx0VR0k7fvlUd
# 3i0baTANBgkqhkiG9w0BAQEFAASCAgBlfFC5kfeB3oPm1Uq+aBdQq/Q6KV+b8tvD
# Zh1AbLfGMU9n28unbxCIaQd8pl4qKWNCi6gsnVtqbDYsmhyqamv2ZQAhy8WMGMcY
# 5PaPH0CGFjSeGIAJT/ad5xhCEXquzgbjQ5YFVHMAr+hUYERTxj8gMweNVmHhxzTe
# IaNsR6Yg0O5/DOyK4dG9QBG3zR7/6is85dJdpruHqJfhpooUfFTGOQ+w/CCns8DQ
# eBJKTv2lM48tgHeWGfBY1fkMyqFaZuJYTl5lUsuTMwxZBVNegnoYnj1YcZsMs9wN
# iwOxGJuzPXoUPoXtg98fLP02Jr34nFqGcgGFgXw/O8P197ADRshJDouG99WSnS6N
# Dp3CJ48WBxOYtuGK3LnHmULq57U0xi+DEn+T6FsBHNtOs22v//TkJLnqk/wQuWHv
# foWDvaCuTKncWFWZLliedRDFOBsnpceP+kXOQa+/KXN88A2EbTIUPJpalV8i1bxb
# rNpHCqUOEe40WrPgqrPBScq11FB0zLY0K4hOl6Y0eK45hQ4lOBsIshLSJxWB6HGo
# AlrNyYj6kx+oo62cqGB14iLK6GrOqMZ5kF7cZn24nu92ZpjSlraFmnenOwFimOSK
# tgiVzVkhh1kEDPMCUWMO2PTF5xl9qHijBrsAHlx1Sy8EhYo4UJc/OqtwyM5NRW46
# VzfRuTIEKg==
# SIG # End signature block
