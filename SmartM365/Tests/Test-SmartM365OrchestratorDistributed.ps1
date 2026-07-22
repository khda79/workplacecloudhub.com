<#
.SYNOPSIS
Validates SmartM365 Orchestrator distributed election and claim behavior.
.VERSION
1.0.2
#>

#Requires -Version 7.0
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$distributedModuleFile = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'SmartInventory\Orchestrator\SmartM365.Orchestrator.Distributed.psm1'
Import-Module -Name $distributedModuleFile -Force -ErrorAction Stop

function Assert-True {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )
    if (-not $Condition) { throw $Message }
}

$jobs = @(
    [pscustomobject]@{
        Name = 'GraphHeavy'
        Enabled = $true
        AssignmentMode = 'Elected'
        DependsOn = @()
        RequiredCapabilities = @('SharedRuntime', 'Graph')
        RequiredGraphAppRoles = @('Reports.Read.All')
        EstimatedDurationMinutes = 60
        Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('01:00', '02:00'); DaysOfWeek = @() }
    }
    [pscustomobject]@{
        Name = 'Parent'
        Enabled = $true
        AssignmentMode = 'Elected'
        DependsOn = @()
        RequiredCapabilities = @('SharedRuntime', 'Graph')
        RequiredGraphAppRoles = @()
        EstimatedDurationMinutes = 10
        Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('03:00'); DaysOfWeek = @() }
    }
    [pscustomobject]@{
        Name = 'Child'
        Enabled = $true
        AssignmentMode = 'Elected'
        DependsOn = @('Parent')
        RequiredCapabilities = @('SharedRuntime', 'Graph')
        RequiredGraphAppRoles = @()
        EstimatedDurationMinutes = 10
        Schedule = [pscustomobject]@{ Type = 'Daily'; Times = @('03:10'); DaysOfWeek = @() }
    }
)
$capabilities = @(
    [pscustomobject]@{
        ServerName = 'SERVER-A'
        ReadyCapabilities = @('SharedRuntime', 'Graph')
        GraphAppRoles = @('Reports.Read.All')
    }
    [pscustomobject]@{
        ServerName = 'SERVER-B'
        ReadyCapabilities = @('SharedRuntime', 'Graph')
        GraphAppRoles = @('Reports.Read.All')
    }
)

$plan = Get-SmartM365OrchestratorElectionPlan `
    -Jobs $jobs `
    -ServerCapabilities $capabilities `
    -ServerWeights @{ 'SERVER-A' = 1.0; 'SERVER-B' = 1.1 }

$parentOwner = [string]($plan.Assignments | Where-Object JobName -eq 'Parent').OwnerServer
$childOwner = [string]($plan.Assignments | Where-Object JobName -eq 'Child').OwnerServer
Assert-True -Condition ($parentOwner -eq $childOwner) -Message 'DependsOn jobs were not assigned to the same server.'
Assert-True -Condition (@($plan.UnassignedGroups).Count -eq 0) -Message 'Eligible mock jobs were unexpectedly left unassigned.'
Assert-True -Condition ([int]$plan.SchemaVersion -eq 2) -Message 'The election plan schema was not upgraded to version 2.'
Assert-True -Condition ((@($plan.EligibleServers) -join ',') -eq 'SERVER-A,SERVER-B') -Message 'The election plan does not record its eligible server set.'

# Startup-race regression: a plan produced while only one server is ready must not
# pin every job to that first server when the rest of the cluster becomes visible.
$singleServerPlan = Get-SmartM365OrchestratorElectionPlan `
    -Jobs $jobs `
    -ServerCapabilities @($capabilities | Where-Object ServerName -eq 'SERVER-A') `
    -ServerWeights @{ 'SERVER-A' = 1.0 }
Assert-True `
    -Condition (-not (Test-SmartM365OrchestratorCanPreserveOwners -PreviousPlan $singleServerPlan -ServerCapabilities $capabilities)) `
    -Message 'A partial startup plan was incorrectly considered safe to preserve after SERVER-B joined.'
$expandedPlan = Get-SmartM365OrchestratorElectionPlan `
    -Jobs $jobs `
    -ServerCapabilities $capabilities `
    -ServerWeights @{ 'SERVER-A' = 1.0; 'SERVER-B' = 1.1 } `
    -PreviousPlan $singleServerPlan
$expandedOwners = @($expandedPlan.Assignments.OwnerServer | Sort-Object -Unique)
Assert-True -Condition ($expandedOwners.Count -eq 2) -Message 'The full rebalance after server discovery did not distribute jobs across both servers.'
Assert-True `
    -Condition (Test-SmartM365OrchestratorCanPreserveOwners -PreviousPlan $expandedPlan -ServerCapabilities $capabilities) `
    -Message 'A complete plan with an unchanged server set was not considered safe to preserve.'
$legacyPlan = [pscustomobject]@{
    SchemaVersion = 1
    Assignments = @($expandedPlan.Assignments)
    UnassignedGroups = @()
    ServerLoads = @($expandedPlan.ServerLoads)
}
Assert-True `
    -Condition (-not (Test-SmartM365OrchestratorCanPreserveOwners -PreviousPlan $legacyPlan -ServerCapabilities $capabilities)) `
    -Message 'A legacy plan was incorrectly allowed to preserve potentially biased owners.'
$incompletePlan = $expandedPlan | Select-Object *
$incompletePlan.UnassignedGroups = @([pscustomobject]@{ GroupKey = 'Synthetic'; Reason = 'Startup capability race' })
Assert-True `
    -Condition (-not (Test-SmartM365OrchestratorCanPreserveOwners -PreviousPlan $incompletePlan -ServerCapabilities $capabilities)) `
    -Message 'An incomplete plan was incorrectly allowed to preserve owners.'


$missingRoleMatch = Test-SmartM365OrchestratorCapabilityMatch `
    -ServerCapabilities ([pscustomobject]@{
        ReadyCapabilities = @('SharedRuntime', 'Graph')
        GraphAppRoles = @()
    }) `
    -RequiredCapabilities @('SharedRuntime', 'Graph') `
    -RequiredGraphAppRoles @('Reports.Read.All')
Assert-True -Condition (-not $missingRoleMatch.Eligible) -Message 'A server missing a required Graph app role was accepted.'

# Integration guard for the committed production template and the validated server policy.
$manifestPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'SmartInventory\Orchestrator\Orchestrator-Jobs.json.template'
$manifest = Get-Content -LiteralPath $manifestPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
$allGraphRoles = @($manifest.Jobs.RequiredGraphAppRoles | ForEach-Object { $_ } | Where-Object { $_ } | Sort-Object -Unique)
$productionCapabilities = @(
    [pscustomobject]@{
        ServerName = 'CPPV-CAPTSE-001'
        ReadyCapabilities = @('SharedRuntime', 'Graph', 'EXO', 'AD')
        GraphAppRoles = $allGraphRoles
    }
    [pscustomobject]@{
        ServerName = 'CPPV-CAPTSE-002'
        ReadyCapabilities = @('SharedRuntime', 'Graph', 'EXO', 'AD', 'TeamsPowerShell')
        GraphAppRoles = $allGraphRoles
    }
    [pscustomobject]@{
        ServerName = 'CPPV-EXCSRV-113'
        # The policy must remain authoritative even when every technical prerequisite exists.
        ReadyCapabilities = @('SharedRuntime', 'Graph', 'EXO', 'AD', 'ExchangeOnPrem', 'TeamsPowerShell')
        GraphAppRoles = $allGraphRoles
    }
)
$productionPlan = Get-SmartM365OrchestratorElectionPlan `
    -Jobs $manifest.Jobs `
    -ServerCapabilities $productionCapabilities `
    -ServerWeights @{
        'CPPV-CAPTSE-001' = 1.0
        'CPPV-CAPTSE-002' = 1.1
        'CPPV-EXCSRV-113' = 1.0
    } `
    -ServerJobPolicies @{
        'CPPV-EXCSRV-113' = @{ OnlyJobsRequiring = @('ExchangeOnPrem') }
    }
Assert-True -Condition (@($productionPlan.UnassignedGroups).Count -eq 0) -Message 'The production template has an unassigned eligible job group in the mock topology.'
$load001 = [double]($productionPlan.ServerLoads | Where-Object ServerName -eq 'CPPV-CAPTSE-001').LoadMinutesPerDay
$load002 = [double]($productionPlan.ServerLoads | Where-Object ServerName -eq 'CPPV-CAPTSE-002').LoadMinutesPerDay
Assert-True -Condition ($load002 -gt $load001) -Message 'CPPV-CAPTSE-002 is not slightly more loaded than CPPV-CAPTSE-001.'
$exchangeOnPremJobNames = @($manifest.Jobs | Where-Object { $_.Enabled -and $_.AssignmentMode -eq 'Elected' -and @($_.RequiredCapabilities) -contains 'ExchangeOnPrem' } | ForEach-Object Name)
$exchangeOnPremOwners = @($productionPlan.Assignments | Where-Object JobName -in $exchangeOnPremJobNames | ForEach-Object OwnerServer | Sort-Object -Unique)
Assert-True -Condition ($exchangeOnPremOwners.Count -eq 1 -and $exchangeOnPremOwners[0] -eq 'CPPV-EXCSRV-113') -Message 'Exchange on-premises jobs are not exclusively assigned to CPPV-EXCSRV-113.'
$server113Assignments = @($productionPlan.Assignments | Where-Object OwnerServer -eq 'CPPV-EXCSRV-113')
$server113PolicyViolations = @(
    foreach ($assignment in $server113Assignments) {
        $assignedJob = $manifest.Jobs | Where-Object Name -eq $assignment.JobName | Select-Object -First 1
        if ($null -eq $assignedJob -or @($assignedJob.RequiredCapabilities) -notcontains 'ExchangeOnPrem') { $assignment.JobName }
    }
)
Assert-True -Condition ($server113Assignments.Count -gt 0) -Message 'CPPV-EXCSRV-113 received no Exchange on-premises jobs.'
Assert-True -Condition ($server113PolicyViolations.Count -eq 0) -Message ("CPPV-EXCSRV-113 received jobs outside the ExchangeOnPrem policy: {0}" -f ($server113PolicyViolations -join ', '))
$teamsPhoneOwner = [string]($productionPlan.Assignments | Where-Object JobName -eq 'M365-TeamsPhonePstnUsage-Inventory').OwnerServer
Assert-True -Condition ($teamsPhoneOwner -eq 'CPPV-CAPTSE-002') -Message 'The Teams Phone PSTN usage job is not assigned to CPPV-CAPTSE-002.'
Assert-True -Condition (@($productionPlan.Assignments | Where-Object JobName -eq 'M365-PowerBIFabricActivity-Inventory').Count -eq 0) -Message 'The manual Power BI Fabric activity job was included in the scheduled election plan.'

$temporaryRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ('SmartM365-OrchestratorDistributed-' + [guid]::NewGuid().ToString('N'))
try {
    $claimsRoot = Join-Path -Path $temporaryRoot -ChildPath 'Claims'
    $heartbeatRoot = Join-Path -Path $temporaryRoot -ChildPath 'Heartbeats'
    $occurrence = [datetime]'2026-07-19T12:00:00'

    $firstClaim = Enter-SmartM365OrchestratorOccurrenceClaim `
        -ClaimsRootPath $claimsRoot `
        -JobName 'GraphHeavy' `
        -Occurrence $occurrence `
        -OwnerServer 'SERVER-A' `
        -PlanId 'plan-a' `
        -HeartbeatRootPath $heartbeatRoot
    $duplicateClaim = Enter-SmartM365OrchestratorOccurrenceClaim `
        -ClaimsRootPath $claimsRoot `
        -JobName 'GraphHeavy' `
        -Occurrence $occurrence `
        -OwnerServer 'SERVER-B' `
        -PlanId 'plan-b' `
        -HeartbeatRootPath $heartbeatRoot
    Assert-True -Condition $firstClaim.Acquired -Message 'The first atomic occurrence claim was not acquired.'
    Assert-True -Condition (-not $duplicateClaim.Acquired) -Message 'A duplicate occurrence claim was accepted.'

    $claimDocument = Get-Content -LiteralPath $firstClaim.ClaimPath -Raw | ConvertFrom-Json
    $claimDocument.SafeUntilUtc = [datetime]::UtcNow.AddMinutes(-5).ToString('o')
    [System.IO.File]::WriteAllText(
        $firstClaim.ClaimPath,
        ($claimDocument | ConvertTo-Json -Depth 6),
        [System.Text.UTF8Encoding]::new($false)
    )
    $takeoverClaim = Enter-SmartM365OrchestratorOccurrenceClaim `
        -ClaimsRootPath $claimsRoot `
        -JobName 'GraphHeavy' `
        -Occurrence $occurrence `
        -OwnerServer 'SERVER-B' `
        -PlanId 'plan-b' `
        -HeartbeatRootPath $heartbeatRoot `
        -HeartbeatStaleMinutes 5
    Assert-True -Condition $takeoverClaim.Acquired -Message 'An expired claim with a stale/missing owner heartbeat was not safely taken over.'
    Assert-True -Condition ($takeoverClaim.Claim.OwnerServer -eq 'SERVER-B') -Message 'The failover claim owner is incorrect.'

    # Regression guard: a PID re-adopted from pre-claim state must claim its original occurrence.
    $orchestratorPath = Join-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -ChildPath 'SmartInventory\Orchestrator\SmartM365-Inventory-Orchestrator.ps1'
    $orchestratorTokens = $null
    $orchestratorErrors = $null
    $orchestratorAst = [System.Management.Automation.Language.Parser]::ParseFile(
        $orchestratorPath,
        [ref]$orchestratorTokens,
        [ref]$orchestratorErrors
    )
    Assert-True -Condition ($orchestratorErrors.Count -eq 0) -Message 'The orchestrator could not be parsed for the Restore-RunningJobs regression test.'
    foreach ($functionName in @('Get-RunningJobTimeoutWindow', 'Restore-RunningJobs', 'ConvertTo-OrchestratorUtcTime', 'Get-OrchestratorConcurrencyBlockState')) {
        $functionAst = $orchestratorAst.Find(
            {
                param($node)
                $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and
                    $node.Name -eq $functionName
            },
            $true
        )
        Assert-True -Condition ($null -ne $functionAst) -Message ("{0} was not found in the orchestrator." -f $functionName)
        . ([scriptblock]::Create($functionAst.Extent.Text))
    }

    function ConvertFrom-StateTime {
        param([string]$Text)
        if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
        return [datetime]::Parse($Text, [cultureinfo]::InvariantCulture)
    }
    function Test-ProcessMatchesRecord {
        return Get-Process -Id $PID
    }
    function Save-OrchestratorState {
        $script:SavedStateCount++
    }
    function Write-OrchestratorLog {
        param([string]$Message, [string]$Level = 'INFO')
    }
    function Write-OrchestratorRuntimeUpdateWarning {
        param([string]$Key, [string]$Message, [datetime]$Now)
    }
    function Complete-JobRun {
        throw 'Complete-JobRun must not be called while the synthetic process is re-adopted.'
    }

    $recoveryOccurrence = [datetime]'2026-07-19T14:00:00'
    $script:SavedStateCount = 0
    $script:RunningJobs = @{}
    $script:Settings = [pscustomobject]@{
        ElectionClaimsPath = $claimsRoot
        ElectionClaimGraceMinutes = 15
        SharedDataFolderPath = $heartbeatRoot
        PeerHeartbeatStaleMinutes = 5
        ConcurrencyLeasesPath = Join-Path -Path $temporaryRoot -ChildPath 'Concurrency'
    }
    $script:ElectionPlan = [pscustomobject]@{ PlanId = 'recovery-plan' }
    $recoveryJob = [pscustomobject]@{
        Name = 'RecoveryJob'
        AssignmentMode = 'Elected'
        TimeoutMinutes = 240
        ConcurrencyKey = 'RecoveryLease'
    }
    $script:Manifest = [pscustomobject]@{
        OrderedJobs = @($recoveryJob)
        JobsByName = @{ RecoveryJob = $recoveryJob }
    }

    # Regression guard: when PendingJobs briefly lags behind the shared lease,
    # peer monitoring must still recognize the active concurrency blocker.
    $blockedJob = [pscustomobject]@{ Name = 'BlockedFastJob'; ConcurrencyKey = 'EXOMailboxes' }
    $blockerLease = Enter-SmartM365OrchestratorConcurrencyLease `
        -LeasesRootPath $script:Settings.ConcurrencyLeasesPath `
        -ConcurrencyKey $blockedJob.ConcurrencyKey `
        -JobName 'FullMailboxJob' `
        -Occurrence (Get-Date).AddMinutes(-5) `
        -OwnerServer 'SERVER-A' `
        -SafeMinutes 60
    Assert-True -Condition $blockerLease.Acquired -Message 'The active concurrency lease for the peer-monitoring regression test was not acquired.'
    $activeBlock = Get-OrchestratorConcurrencyBlockState -Job $blockedJob -Now (Get-Date)
    Assert-True -Condition ($null -ne $activeBlock -and $activeBlock.Active) -Message 'Peer monitoring did not recognize a valid shared concurrency lease.'
    Assert-True -Condition ($activeBlock.BlockingJob -eq 'FullMailboxJob') -Message 'Peer monitoring reported the wrong concurrency blocker.'

    $staleLeaseDocument = Get-Content -LiteralPath $blockerLease.LeasePath -Raw | ConvertFrom-Json
    $staleLeaseDocument.SafeUntilUtc = [datetime]::UtcNow.AddMinutes(-1).ToString('o')
    [System.IO.File]::WriteAllText(
        $blockerLease.LeasePath,
        ($staleLeaseDocument | ConvertTo-Json -Depth 6),
        [System.Text.UTF8Encoding]::new($false)
    )
    $staleBlock = Get-OrchestratorConcurrencyBlockState -Job $blockedJob -Now (Get-Date)
    Assert-True -Condition ($null -ne $staleBlock -and -not $staleBlock.Active) -Message 'Peer monitoring did not distinguish an expired concurrency lease.'

    $script:State = @{
        Jobs = @{
            RecoveryJob = @{
                Running = @{
                    StartTime = (Get-Date).AddMinutes(-5).ToString('o')
                    ScheduledOccurrence = $recoveryOccurrence.ToString('o')
                    Pid = $PID
                    ProcessName = 'pwsh'
                    LogPath = 'synthetic.log'
                    Attempt = 1
                    TimeoutMinutes = 240
                }
            }
        }
    }

    Restore-RunningJobs
    $restoredRecord = $script:State.Jobs.RecoveryJob.Running
    Assert-True -Condition (-not [string]::IsNullOrWhiteSpace([string]$restoredRecord.ClaimPath)) -Message 'Restore-RunningJobs did not persist the recovered ClaimPath.'
    Assert-True -Condition (Test-Path -LiteralPath $restoredRecord.ClaimPath) -Message 'Restore-RunningJobs did not create the recovered claim file.'
    $restoredClaim = Get-Content -LiteralPath $restoredRecord.ClaimPath -Raw | ConvertFrom-Json
    Assert-True -Condition ($restoredClaim.Status -eq 'Running') -Message 'The recovered claim was not marked Running.'
    Assert-True -Condition (-not [string]::IsNullOrWhiteSpace([string]$restoredRecord.ConcurrencyLeasePath)) -Message 'Restore-RunningJobs did not persist the recovered ConcurrencyLeasePath.'
    Assert-True -Condition (-not [string]::IsNullOrWhiteSpace([string]$restoredRecord.ConcurrencyLeaseId)) -Message 'Restore-RunningJobs did not persist the recovered ConcurrencyLeaseId.'
    Assert-True -Condition (Test-Path -LiteralPath $restoredRecord.ConcurrencyLeasePath) -Message 'Restore-RunningJobs did not create the recovered concurrency lease.'
    $restoredLease = Get-Content -LiteralPath $restoredRecord.ConcurrencyLeasePath -Raw | ConvertFrom-Json
    $expectedLeaseDeadline = ([datetime]$restoredRecord.StartTime).ToUniversalTime().AddMinutes(255)
    $actualLeaseDeadline = ([datetime]$restoredLease.SafeUntilUtc).ToUniversalTime()
    Assert-True -Condition ([string]$restoredLease.LeaseId -eq [string]$restoredRecord.ConcurrencyLeaseId) -Message 'The recovered concurrency lease ID does not match the persisted state.'
    Assert-True -Condition ([string]$restoredLease.OwnerServer -eq $env:COMPUTERNAME.ToUpperInvariant()) -Message 'The recovered concurrency lease owner is incorrect.'
    Assert-True -Condition ([int]$restoredLease.OrchestratorPid -eq $PID) -Message 'The recovered concurrency lease was not re-bound to the current orchestrator PID.'
    Assert-True -Condition ([math]::Abs(($actualLeaseDeadline - $expectedLeaseDeadline).TotalSeconds) -lt 2) -Message 'The recovered concurrency lease was not aligned with the original timeout deadline.'
    Assert-True -Condition ($script:SavedStateCount -eq 1) -Message 'The recovered claim and concurrency lease were not saved exactly once.'
    Assert-True -Condition $script:RunningJobs.ContainsKey('RecoveryJob') -Message 'The synthetic PID was not re-adopted after its claim and concurrency lease were recovered.'
}
finally {
    if (Test-Path -LiteralPath $temporaryRoot) {
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Output 'SmartM365 Orchestrator distributed scheduling mock tests passed.'

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAh26m9MFTZDTQW
# wHDUwZzBUBPguEQrzZKpZStA4wpTCKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEINkEUoilKI9ijkvvzx0U+gEhC0+LuGvuzDX8C5XL5f/MMA0GCSqG
# SIb3DQEBAQUABIIBgHkuxtledM9l3m90elZo2DLRijh32San3aZ6VKR2q+ud+xmW
# pKT2Y+m36g9xjF57RvZ5OzB/j9IxqMKF8zSw3FLDtFRiQYF7RraLXTKS0+vdlXlY
# UbcEqBM8tHr9AEUpCgZL9uTGL/0jRKCjWh+suRrq2iYBbsk+uoe/bMvlInjXY4q5
# ki/RGa9irGsIhXlWyIqpOjnBYY45ghdXImGbcq2maAmTjBsow148ugNBh54zO7cr
# ybbRbQgdoB1e84IvsksW48kBytWijUZlyNsbbdqnbYGd8CIvohBKmsJdcGBC1EO7
# HYNx0iNeWVX2EFW6XoAWeUVonqMUsPA5s2aOqopmpB7SLtiT/kcCtIZ3K0S/oK1S
# NMnTwEGudlvn2AaIMCP6XUvqvFLSJe+aqK56EavTduRiVt2aiEcCbxugnMg69ggb
# zDahNWNqcKq5qCFAu4bhijCxzbT3UgAn8+6u0X2X4NGi9bqpOdwNemv8uGSeLtce
# Bbc5yVL/4m1oHxL1kKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjExMTMy
# NTRaMC8GCSqGSIb3DQEJBDEiBCB8kMLmIFkngKGm6QB1cf3TMzwVc+hup+PfnVte
# XFG2qjANBgkqhkiG9w0BAQEFAASCAgAVL8adyKFFf0FzJwfvHB7xVXkTbwXA75Eh
# phqFI4IAsu/EkKSXnpob04kgkLSLyGG1FIyOIrAfK28b19Gb1/UtZnbY5Azx/9aM
# FnJNkMvrxSd9vAlH5Q2s3IPotp8OqRn35+zzLfHLrAxz6pyN9feH+RwYozvPgN+w
# CFCqFppXvrg/hBDtJmtXqM5ImRC431RpekpSn9pQTL/TtbXpJEgNJPpAgPaUzD09
# N+V8x8JlMiv8rNDnZ3AOtL9LLVxrOlk+kl81n9UlcnaJeAskN+NqGhkDvBEcFowV
# o5tn72RFkQjAgvD+Gu8fFdwrJpHgrUDv/OjaRlYoZKqsPwdhoS5MnI+0rjUpKaAr
# xfgDXcENiC8JLb1ZGo7dd3J6qwA46KBypLtgIDjDm7+9JXhK4+bIXd27BJi0oFsj
# qvf+GmbdU8C8c5EvDYSak+KK2Df71giVSFX7VX8H93eQxqtMBXzOP1Ez1KjMBXhy
# qyWHKVzA6I9t11L+ZR10yj0wtZWw99XMt9HdwFxTlwLBgQTHo73ucnE8n6i85YH6
# 67/k6PbFaL5DJWvEU3pOrJ1IM566iqDEvOlP6aBhlwBQR3HG899sHDYpqH1Rtrmv
# qcDRyFlexNhCO4HRDmTAXyzmN/HTSJunV89avr1QvzLS/4+Nppcv0kb9gcr1rAhW
# Ws/y8UPOHw==
# SIG # End signature block
