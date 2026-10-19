#Requires -Version 7.0
<#
.SYNOPSIS
Offline contract tests for the retired CMDB integration candidate.
.DESCRIPTION
CmdbEvidence-Prepare is retired: the unified WorkplaceEvidence-Prepare job carries the
former CMDB producer dependencies. Reads repository contracts, merges jobs in memory and
calls the real pipeline selector and cutover planner. No tenant configuration, API,
submission, scheduling or file write.
.VERSION
1.1.0
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$smartRoot = Split-Path $PSScriptRoot -Parent
$repoRoot = Split-Path $smartRoot -Parent
$inventoryRoot = Join-Path $smartRoot 'SmartInventory'
$candidatePath = Join-Path $inventoryRoot 'PreparedEvidence/cmdb-orchestrator-integration.json.txt'
$jobsPath = Join-Path $inventoryRoot 'Orchestrator/Orchestrator-Jobs.json.template'
$sourcePath = Join-Path $smartRoot 'Modules/SmartM365.Core/SmartM365-CmdbSources.json.txt'
$intelligencePath = Join-Path $repoRoot 'SmartWorkplaceIntelligence/config/prepared-evidence-contract.json.txt'
$entryPath = Join-Path $inventoryRoot 'PreparedEvidence/SmartM365-CmdbEvidence-Orchestrator.ps1'
$hashes = @{}
foreach ($path in @($candidatePath,$jobsPath,$sourcePath,$intelligencePath,$entryPath)) { $hashes[$path] = (Get-FileHash -LiteralPath $path).Hash }
$candidate = Get-Content -LiteralPath $candidatePath -Raw | ConvertFrom-Json
$template = Get-Content -LiteralPath $jobsPath -Raw | ConvertFrom-Json
$producers = @((Get-Content -LiteralPath $sourcePath -Raw | ConvertFrom-Json).Producers)
$intelligence = Get-Content -LiteralPath $intelligencePath -Raw | ConvertFrom-Json
Import-Module (Join-Path $inventoryRoot 'Orchestrator/SmartM365.Orchestrator.Pipeline.psm1') -Force
Import-Module (Join-Path $inventoryRoot 'Orchestrator/SmartM365.WorkplaceCutover.psm1') -Force
function Assert-True { param([bool]$Condition, [string]$Label) if (-not $Condition) { throw $Label } }
function ConvertTo-CompactJson { param($Value) ConvertTo-Json -InputObject $Value -Depth 50 -Compress }
$testCount = 0
function Test-Case {
    param([string]$Name, [scriptblock]$Body)
    & $Body
    $script:testCount++
    Write-Output "PASS: $Name"
}
$retired = @($candidate.ProposedJobs | Where-Object Name -eq 'CmdbEvidence-Prepare')[0]
$unified = @($template.Jobs | Where-Object Name -eq 'WorkplaceEvidence-Prepare')
Test-Case 'Retired candidate stays an inactive review-only fixture' {
    Assert-True ($candidate.CandidateOnly -and -not $candidate.PSObject.Properties['Jobs']) 'Executable candidate manifest.'
    Assert-True (@($candidate.ProposedJobs).Count -eq 2) 'Unexpected integration expansion.'
    foreach ($job in $candidate.ProposedJobs) {
        Assert-True (-not $job.Enabled -and $job.AssignmentMode -eq 'Elected') 'Job activated or misleadingly Manual.'
        Assert-True (-not $job.PSObject.Properties['Schedule'] -and -not $job.PSObject.Properties['TimeoutMinutes']) 'Unqualified schedule/timeout.'
    }
    Assert-True (@($candidate.JobArgumentAdditions).Count -eq 1) 'Unrelated argument changes.'
    $addition = $candidate.JobArgumentAdditions[0]
    Assert-True ($addition.Name -eq 'Intune-DiscoveredApps-Inventory' -and $addition.Arguments -eq '-DeviceDetailMode All -FreshDeviceDetails') 'Fresh arguments missing.'
    foreach ($argument in @('-ResetResume','-MaxApps','-MaxItems','-DeviceDetailMode')) { Assert-True ($addition.RejectConflictingArguments -contains $argument) "Conflict not declared: $argument" }
    $tokens = $null; $errors = $null
    $entry = [Management.Automation.Language.Parser]::ParseFile($entryPath,[ref]$tokens,[ref]$errors)
    Assert-True ($errors.Count -eq 0) 'Retired CMDB entry point does not parse.'
    $first = $entry.EndBlock.Statements[0]
    Assert-True ($first -is [Management.Automation.Language.ThrowStatementAst] -and $first.Extent.Text -match 'CMDB activation is retired') 'Retired activation entry point can run.'
}
Test-Case 'Template keeps WorkplaceScope active and the retired CMDB job absent' {
    Assert-True (@($template.Jobs | Where-Object Name -eq 'CmdbEvidence-Prepare').Count -eq 0) 'Retired CMDB job remains in the active template.'
    Assert-True (@($template.Jobs | Where-Object { 'CmdbEvidence-Prepare' -in @($_.DependsOn) }).Count -eq 0) 'A job still depends on the retired CMDB job.'
    $scope = @($template.Jobs | Where-Object Name -eq 'M365-WorkplaceScope-Inventory')
    Assert-True ($scope.Count -eq 1 -and $scope[0].Enabled -and $scope[0].AssignmentMode -ne 'Manual' -and -not $scope[0].PSObject.Properties['RequiresExplicitActivation']) 'M365-WorkplaceScope-Inventory is not enabled in the template.'
    $proposedScope = @($candidate.ProposedJobs | Where-Object Name -eq 'M365-WorkplaceScope-Inventory')[0]
    Assert-True ($scope[0].ScriptPath -eq $proposedScope.ScriptPath) 'WorkplaceScope script path drifted.'
    Assert-True (Test-Path -LiteralPath (Join-Path $inventoryRoot $scope[0].ScriptPath) -PathType Leaf) 'WorkplaceScope script missing.'
    Assert-True ($unified.Count -eq 1 -and $unified[0].Enabled -and $unified[0].AssignmentMode -ne 'Manual' -and $unified[0].DependencyMode -eq 'FreshSuccess') 'Unified preparation is missing, disabled or ungated.'
}
Test-Case 'Unified preparation covers all 17 native producers' {
    $dependsOn = @($unified[0].DependsOn)
    Assert-True ($producers.Count -eq 17 -and @($retired.DependsOn).Count -eq 17) 'Producer registry changed; explicit review needed.'
    Assert-True (@($dependsOn | Sort-Object -Unique).Count -eq $dependsOn.Count) 'Duplicate dependency.'
    foreach ($name in $retired.DependsOn) { Assert-True ($dependsOn -contains $name) "Retired CMDB dependency not carried by the unified job: $name" }
    $dependencies = @(foreach ($name in $retired.DependsOn) {
        $matches = @($template.Jobs | Where-Object Name -eq $name)
        Assert-True ($matches.Count -eq 1 -and $matches[0].Enabled) "Missing, ambiguous or disabled CMDB producer job: $name"
        $matches[0]
    })
    # A producer script can back several jobs (EXO mailbox details and permissions).
    foreach ($producer in $producers) {
        Assert-True (@($dependencies | Where-Object { [IO.Path]::GetFileName($_.ScriptPath) -eq $producer.Script }).Count -ge 1) "Producer coverage: $($producer.Script)"
    }
    Assert-True ($dependsOn -contains 'EXO-Mailboxes-Inventory-Fast' -and $dependsOn -notcontains 'EXO-Mailboxes-Inventory') 'Acquisition must use daily mailbox details, not weekly stats.'
    $contract = Get-Content -LiteralPath (Join-Path $inventoryRoot 'PreparedEvidence/cmdb-prepared-contract.json.txt') -Raw | ConvertFrom-Json
    Assert-True ($contract.maxAgeHours -eq 48 -and $contract.maxCollectionSpanHours -eq 48) 'Core acquisition gates weakened.'
    Assert-True ($contract.freshnessGroups.Count -eq 1 -and $contract.freshnessGroups[0].maxAgeHours -eq 240 -and $contract.freshnessGroups[0].warningAgeHours -eq 168) 'Apps acquisition gates differ.'
}
Test-Case 'Real selector rejects the retired job and selects every producer for the unified job' {
    $rejected = $false
    try { Get-SmartM365OrchestratorPipelineSelection -JobsDocument $template -JobName CmdbEvidence-Prepare -IncludeDependencies | Out-Null }
    catch { $rejected = $_.Exception.Message -like '*absent from the jobs manifest*' }
    Assert-True $rejected 'Retired CMDB job can be requested.'
    $selection = Get-SmartM365OrchestratorPipelineSelection -JobsDocument $template -JobName WorkplaceEvidence-Prepare -IncludeDependencies
    foreach ($name in $retired.DependsOn) { Assert-True (@($selection.SelectedJobs | Where-Object Name -eq $name).Count -eq 1) "Not selected: $name" }
}
Test-Case 'A re-merged legacy CMDB job is removed by the cutover plan, enabled or disabled' {
    foreach ($enabled in @($false,$true)) {
        $merged = ConvertTo-CompactJson $template | ConvertFrom-Json -Depth 50
        $legacy = ConvertTo-CompactJson $retired | ConvertFrom-Json -Depth 50
        $legacy.Enabled = $enabled
        $merged.Jobs = @($merged.Jobs) + @($legacy)
        $plan = New-SmartM365WorkplaceCutoverPlan $merged
        Assert-True (@($plan.Jobs.Jobs | Where-Object Name -eq 'CmdbEvidence-Prepare').Count -eq 0) 'Re-merged CMDB job survived cutover.'
        Assert-True ((ConvertTo-CompactJson $plan.Jobs) -ceq (ConvertTo-CompactJson $template)) 'Cutover changed unrelated template jobs.'
    }
}
Test-Case 'Every Intelligence History/Trend retains its explicit stable key columns' {
    $dailyKey = 'TenantKey|OrganizationKey|EnvironmentKey|TenantId|Date Key|Source Name|Domain Key|Metric Name'
    $expected = @{
        'User Activity History Evidence'='Week Label'; 'Workforce Trend Evidence'='Week Label'
        'Windows Lifecycle Trend Evidence'='Snapshot Week'; 'Endpoint Experience Trend Evidence'='Snapshot Week'
        'Application Trend Evidence'='Snapshot Week'; 'Content Storage Trend Evidence'='Snapshot Week'
        'Collaboration Trend Evidence'='Snapshot Week|Service'; 'Executive KPI Trends'='Date Key|Metric Name'
        'Daily Directory Trends'=$dailyKey; 'Daily Mailbox Trends'=$dailyKey
    }
    $tables = @($intelligence.tables | Where-Object table -match 'History|Trend')
    Assert-True ($tables.Count -eq $expected.Count) 'History/Trend set changed; explicit review needed.'
    foreach ($table in $tables) {
        Assert-True ($table.PSObject.Properties['historyKey'] -and @($table.historyKey).Count -gt 0) "Missing historyKey: $($table.table)"
        Assert-True ((@($table.historyKey) -join '|') -eq $expected[$table.table]) "Stable key changed: $($table.table)"
        foreach ($key in $table.historyKey) { Assert-True (@($table.columns.name) -contains $key) "Key column missing: $($table.table)/$key" }
    }
}
Test-Case 'Inspected files unchanged' {
    foreach ($path in $hashes.Keys) { Assert-True ((Get-FileHash -LiteralPath $path).Hash -eq $hashes[$path]) "File mutated: $path" }
    Assert-True (@($candidate.ProposedJobs | Where-Object Enabled).Count -eq 0) 'Real candidate enabled.'
}
Write-Output "CMDB post-retirement contract tests: $testCount passed. No requests or external actions."
