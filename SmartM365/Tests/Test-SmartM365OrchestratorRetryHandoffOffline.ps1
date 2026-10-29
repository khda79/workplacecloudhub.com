<#
.SYNOPSIS
Offline test of the retry handoff: an elected job that failed on SRV-A is retried by a free,
eligible server, and exactly one server runs the retry.
.VERSION
1.0.0
.NOTES
Real Distributed module claims in a temporary folder; the orchestrator offer and adoption
functions run in one module instance per simulated server. No tenant initialization,
collectors, scheduling or publication.
#>
#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
$script:passed = 0
function Check($Value, $Message) { if (-not $Value) { throw $Message }; $script:passed++ }

Import-Module (Join-Path $base 'SmartM365.Orchestrator.Distributed.psm1') -Force
$transport = Get-Module SmartM365.JsonTransport
& $transport { function script:Get-SmartM365JsonTransportPolicy { @{ Mode = 'Readers'; QualifiedUncRoots = @(); QualifiedLocalDataRoots = @() } } }

$orchestratorPath = Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'
$ast = [Management.Automation.Language.Parser]::ParseFile($orchestratorPath, [ref]$null, [ref]$null)
function Get-FunctionText([string]$Name) {
    $definition = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
    if ($null -eq $definition) { throw "Missing orchestrator function: $Name" }
    $definition.Extent.Text
}

# Wiring: offered only for elected jobs once the claim is RetryScheduled; the failing server waits.
$complete = Get-FunctionText 'Complete-JobRun'
Check ($complete -match "claimStatus -eq 'RetryScheduled'.*AssignmentMode -eq 'Elected'" -and $complete -match 'Write-OrchestratorRetryOffer') 'Only an elected retry with a RetryScheduled claim may be offered.'
$launch = Get-FunctionText 'Invoke-LaunchPhase'
Check ($launch -match 'Invoke-RetryHandoffAdoption -Now \$Now' -and $launch -match '\$Now -lt \$notBefore\.AddSeconds\(\$script:RetryHandoffGraceSeconds\)') 'Peers must adopt first and the failing server must wait its grace period.'
Check ($launch -match '-RetryHandoffFrom \$retryHandoffFrom' -and $launch -match 'Remove-OrchestratorRetryOffer -JobName \$name -OfferId \$retryOfferId') 'The launch must take the claim over atomically and remove the offer.'

$functions = @('ConvertTo-OrchestratorUtcTime', 'Get-OrchestratorRetryOfferPath', 'Write-OrchestratorRetryOffer', 'Get-OrchestratorRetryOffers', 'Remove-OrchestratorRetryOffer', 'Test-OrchestratorRetryHandoffEligible', 'Invoke-RetryHandoffAdoption', 'Test-JobAllowedOnServer', 'ConvertFrom-StateTime', 'ConvertTo-StateTime')
$stubs = @'
function Get-JobState { param($JobName) if (-not $script:State.Jobs.ContainsKey($JobName)) { $script:State.Jobs[$JobName] = @{ PendingRetry = $null } }; $script:State.Jobs[$JobName] }
function Save-OrchestratorState { }
function Write-OrchestratorLog { param($Message, $Level = 'INFO') $script:Logs.Add($Message) }
function Write-OrchestratorRuntimeUpdateWarning { param($Key, $Message, $Now, $Level) $script:Logs.Add($Message) }
function Test-JobSelected { param($JobName) $true }
function Get-ElectedJobOwner { param($JobName) 'SRV-A' }
'@
$functionText = (@($functions | ForEach-Object { Get-FunctionText $_ }) -join "`n") + "`n" + $stubs
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-RetryHandoff-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
$claims = Join-Path $root 'Claims'
New-Item -ItemType Directory -Path $claims -Force | Out-Null
$originalComputer = $env:COMPUTERNAME
$job = [pscustomobject]@{ Name = 'Graph-Job'; AssignmentMode = 'Elected'; RequiredCapabilities = @('Graph'); RequiredGraphAppRoles = @(); TimeoutMinutes = 60 }
function New-Server([string]$Name, [string[]]$Ready, [bool]$Maintenance = $false) {
    $module = New-Module -Name ('Server-' + $Name) -ScriptBlock ([scriptblock]::Create($functionText))
    & $module {
        param($Root, $Claims, $Job, $Ready, $Maintenance)
        $script:Logs = [Collections.Generic.List[string]]::new()
        $script:State = @{ Jobs = @{} }
        $script:RunningJobs = @{}
        $script:Settings = [pscustomobject]@{ SharedDataFolderPath = $Root; ElectionClaimsPath = $Claims; DistributedSchedulingEnabled = $true; MaxConcurrency = 2; ServerJobPolicies = @{} }
        $script:ElectionPlan = [pscustomobject]@{ PlanId = 'plan-test' }
        $script:ServerCapabilities = [pscustomobject]@{ ReadyCapabilities = $Ready; GraphAppRoles = @() }
        $script:Manifest = [pscustomobject]@{ JobsByName = @{ $Job.Name = $Job } }
        $script:MaintenanceControl = [pscustomobject]@{ Enabled = $Maintenance }
    } $root $claims $job $Ready $Maintenance
    $module
}
function Invoke-As([string]$Server, $Module, [scriptblock]$Block, [object[]]$Arguments = @()) {
    $env:COMPUTERNAME = $Server
    try { & $Module $Block @Arguments } finally { $env:COMPUTERNAME = $originalComputer }
}
function New-FailedOccurrence($ServerA, [datetime]$Occurrence, [datetime]$NotBefore, [string]$BatchId = 'BATCH-1') {
    # SRV-A ran the job, it failed with a retry left: claim RetryScheduled, local PendingRetry, offer.
    Invoke-As 'SRV-A' $ServerA {
        param($Occurrence, $NotBefore, $BatchId, $Claims)
        $claim = Enter-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $Claims -JobName 'Graph-Job' -Occurrence $Occurrence -OwnerServer 'SRV-A' -PlanId 'plan-test'
        $null = Set-SmartM365OrchestratorOccurrenceClaim -ClaimPath $claim.ClaimPath -OwnerServer 'SRV-A' -Status RetryScheduled -Attempt 0
        $pending = @{ NotBefore = ConvertTo-StateTime -Value $NotBefore; Attempt = 1; ScheduledOccurrence = ConvertTo-StateTime -Value $Occurrence; PipelineBatchId = $BatchId; PipelineStatusPath = '' }
        $pending['HandoffOfferId'] = Write-OrchestratorRetryOffer -JobName 'Graph-Job' -PendingRetry $pending -ClaimPath $claim.ClaimPath
        $pending['HandoffOffered'] = $true
        $script:State.Jobs['Graph-Job'] = @{ PendingRetry = $pending }
        [pscustomobject]@{ ClaimPath = $claim.ClaimPath; OfferId = $pending.HandoffOfferId }
    } @($Occurrence, $NotBefore, $BatchId, $claims)
}
function Enter-Claim($Module, [string]$Server, [datetime]$Occurrence, [string]$HandoffFrom = '') {
    Invoke-As $Server $Module { param($Occurrence, $Server, $HandoffFrom, $Claims) Enter-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $Claims -JobName 'Graph-Job' -Occurrence $Occurrence -OwnerServer $Server -PlanId 'plan-test' -RetryHandoffFrom $HandoffFrom } @($Occurrence, $Server, $HandoffFrom, $claims)
}
try {
    $serverA = New-Server 'SRV-A' @('Graph')
    $serverB = New-Server 'SRV-B' @('Graph')
    $serverC = New-Server 'SRV-C' @('AD')
    $serverD = New-Server 'SRV-D' @('Graph')
    $now = Get-Date
    $occurrence = $now.Date.AddHours(5).AddMinutes(10).ToUniversalTime()

    # Failure on SRV-A; the retry delay is still running: nobody adopts it yet.
    $failed = New-FailedOccurrence $serverA $occurrence $now.AddMinutes(5)
    Invoke-As 'SRV-B' $serverB { param($n) Invoke-RetryHandoffAdoption -Now $n } @($now)
    Check ($null -eq (& $serverB { (Get-JobState -JobName 'Graph-Job').PendingRetry })) 'A retry must not be adopted before RetryDelaySeconds.'

    # After the delay: SRV-B (free, eligible) adopts; SRV-C lacks Graph; SRV-A ignores its own offer.
    $later = $now.AddMinutes(6)
    foreach ($pair in @(@('SRV-B', $serverB), @('SRV-C', $serverC), @('SRV-D', $serverD), @('SRV-A', $serverA))) { Invoke-As $pair[0] $pair[1] { param($n) Invoke-RetryHandoffAdoption -Now $n } @($later) }
    $adoptedB = & $serverB { (Get-JobState -JobName 'Graph-Job').PendingRetry }
    Check ($null -ne $adoptedB -and $adoptedB.HandoffFrom -eq 'SRV-A' -and $adoptedB.Attempt -eq 1 -and $adoptedB.PipelineBatchId -eq 'BATCH-1' -and $adoptedB.HandoffOfferId -eq $failed.OfferId) 'SRV-B must adopt the retry with its attempt and pipeline batch.'
    Check ($null -eq (& $serverC { (Get-JobState -JobName 'Graph-Job').PendingRetry })) 'A server without the required capability must not adopt the retry.'
    Check ((Invoke-As 'SRV-B' $serverB { Test-JobAllowedOnServer -Job $script:Manifest.JobsByName['Graph-Job'] }) -and -not (Invoke-As 'SRV-C' $serverC { Test-JobAllowedOnServer -Job $script:Manifest.JobsByName['Graph-Job'] })) 'Only the adopting server may launch the elected job it does not own.'

    # Race: SRV-B and SRV-D both adopted; the atomic claim lets exactly one server run the retry.
    $claimB = Enter-Claim $serverB 'SRV-B' $occurrence 'SRV-A'
    $claimD = Enter-Claim $serverD 'SRV-D' $occurrence 'SRV-A'
    $claimA = Enter-Claim $serverA 'SRV-A' $occurrence
    $winners = @(@($claimB, $claimD, $claimA) | Where-Object Acquired)
    $stored = (Get-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $claims -JobName 'Graph-Job' -Occurrence $occurrence).Claim
    Check ($winners.Count -eq 1 -and $claimB.Acquired -and $stored.OwnerServer -eq 'SRV-B' -and $stored.Status -eq 'Claimed' -and $stored.HandoffFrom -eq 'SRV-A') "Exactly one server must win the retry (B=$($claimB.Acquired), D=$($claimD.Acquired), A=$($claimA.Acquired))."
    # The losers drop their copy of the retry when their claim is refused (Invoke-LaunchPhase).
    & $serverD { $script:State.Jobs['Graph-Job'].PendingRetry = $null }

    # The launcher removes the offer; peers then ignore the occurrence.
    Invoke-As 'SRV-B' $serverB { param($id) Remove-OrchestratorRetryOffer -JobName 'Graph-Job' -OfferId $id } @($failed.OfferId)
    Check (@(Invoke-As 'SRV-D' $serverD { Get-OrchestratorRetryOffers }).Count -eq 0) 'The offer must be removed once the retry is launched.'

    # No free eligible peer: SRV-A takes its own retry back atomically, and a late peer is refused.
    $occurrence2 = $occurrence.AddDays(-1)
    $null = New-FailedOccurrence $serverA $occurrence2 $now.AddMinutes(-3) ''
    $ownA = Enter-Claim $serverA 'SRV-A' $occurrence2
    $lateB = Enter-Claim $serverB 'SRV-B' $occurrence2 'SRV-A'
    $stored2 = (Get-SmartM365OrchestratorOccurrenceClaim -ClaimsRootPath $claims -JobName 'Graph-Job' -Occurrence $occurrence2).Claim
    Check ($ownA.Acquired -and -not $lateB.Acquired -and $stored2.OwnerServer -eq 'SRV-A' -and $stored2.Status -eq 'Claimed') 'Without a peer, the failing server must keep its retry and a late peer must be refused.'

    # The offer of that occurrence is now stale (claim no longer RetryScheduled): any server removes it.
    Invoke-As 'SRV-D' $serverD { param($n) Invoke-RetryHandoffAdoption -Now $n } @($later)
    Check (@(Invoke-As 'SRV-D' $serverD { Get-OrchestratorRetryOffers }).Count -eq 0 -and $null -eq (& $serverD { (Get-JobState -JobName 'Graph-Job').PendingRetry })) 'A stale offer must be cleaned and never adopted.'

    # Maintenance: a scheduled (non-pipeline) retry is not adopted while maintenance is on.
    $serverM = New-Server 'SRV-M' @('Graph') $true
    $occurrence3 = $occurrence.AddDays(-2)
    $null = New-FailedOccurrence $serverA $occurrence3 $now.AddMinutes(-3) ''
    Invoke-As 'SRV-M' $serverM { param($n) Invoke-RetryHandoffAdoption -Now $n } @($later)
    Check ($null -eq (& $serverM { (Get-JobState -JobName 'Graph-Job').PendingRetry })) 'Maintenance must keep a scheduled retry from being adopted.'
}
finally {
    $env:COMPUTERNAME = $originalComputer
    Get-Module -Name 'Server-*' | Remove-Module -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
"PASS: retry handoff, $($script:passed) offline checks."
