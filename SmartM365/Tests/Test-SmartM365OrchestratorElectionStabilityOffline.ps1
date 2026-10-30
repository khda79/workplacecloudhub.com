<#
.SYNOPSIS
Offline tests of election stability and peer job audit (orchestrator 1.5.79): no JobNotStarted when
another server still holds the occurrence, exclusion window for stale heartbeats, last known state on
unreadable capability/heartbeat/plan files and short retry on a sharing violation.
.VERSION
1.0.0
.NOTES
Runs the real orchestrator functions extracted from the script with the real JSON transport module,
in a temporary folder. Sharing violations are simulated by opening the files without sharing. Logs,
claims, leases and the planner are stubbed: no mail, no real peer, no collection.
#>
#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$base = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
$script:passed = 0
function Check($Value, $Message) { if (-not $Value) { throw $Message }; $script:passed++ }

$ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $base 'SmartM365-Inventory-Orchestrator.ps1'), [ref]$null, [ref]$null)
function Get-FunctionText([string]$Name) {
    $definition = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
    if ($null -eq $definition) { throw "Missing orchestrator function: $Name" }
    $definition.Extent.Text
}
Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Force
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Management.psm1') -Force
foreach ($name in @('Get-OrchestratorPeerHealthSnapshot', 'Find-OrchestratorClusterOccurrenceActivity', 'New-OrchestratorPeerIssue', 'ConvertTo-OrchestratorUtcTime',
        'ConvertFrom-StateTime', 'ConvertTo-StateTime', 'Get-OrchestratorPeerRecycleGrace', 'Get-OrchestratorElectionExclusionMinutes', 'Read-OrchestratorSharedJsonDocument',
        'Test-OrchestratorTransientReadError', 'Test-JobAllowedOnNamedServer', 'Get-ElectedJobOwner', 'Test-OrchestratorLocalServerActive', 'Get-OrchestratorConcurrencyBlockState',
        'Get-LiveOrchestratorServerCapabilities', 'Read-SharedElectionPlan', 'Update-OrchestratorElectionPlan')) {
    . ([scriptblock]::Create((Get-FunctionText $name)))
}
$script:Logs = [Collections.Generic.List[string]]::new()
function Write-OrchestratorLog { param($Message, $Level = 'INFO') $script:Logs.Add("[$Level] $Message") }
function Write-OrchestratorRuntimeUpdateWarning { param($Key, $Message, $Now, $Level = 'WARN') $script:Logs.Add("[$Level] $Message") }
function Get-JobEffectiveAllowedServers { param($Job) @() }
function Get-SmartM365OrchestratorOccurrenceClaim { param($ClaimsRootPath, $JobName, $Occurrence) $null }
function Get-SmartM365OrchestratorConcurrencyLease { param($LeasesRootPath, $ConcurrencyKey) $null }
function Get-LatestPastOccurrence { param($Job, $Now) $script:Occurrence }

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-Election-' + [guid]::NewGuid().ToString('N').Substring(0, 12))
New-Item -ItemType Directory -Path $root -Force | Out-Null
function Write-Json([string]$Path, $Value) {
    New-Item -ItemType Directory -Path (Split-Path $Path -Parent) -Force | Out-Null
    [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
}
function Write-Peer([string]$Server, [double]$AgeMinutes, [object[]]$Pending = @(), [object[]]$Running = @(), [string]$Lifecycle = 'Running') {
    $folder = Join-Path $root $Server
    Write-Json (Join-Path $folder 'Orchestrator-Heartbeat.json') ([ordered]@{ Timestamp = [datetime]::UtcNow.AddMinutes(-$AgeMinutes).ToString('o'); Lifecycle = $Lifecycle; Pid = 42; RunningJobs = @($Running); PendingJobs = @($Pending) })
    Write-Json (Join-Path $folder 'Orchestrator-State.json') ([ordered]@{ Jobs = [ordered]@{} })
    Write-Json (Join-Path $folder 'Orchestrator-Capabilities.json') ([ordered]@{ ServerName = $Server; GeneratedAtUtc = [datetime]::UtcNow.ToString('o'); ReadyCapabilities = @('Graph') })
}
function Lock-File([string]$Path) { [IO.File]::Open($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None) }

$originalComputer = $env:COMPUTERNAME
try {
    $env:COMPUTERNAME = 'SRV-A'
    $now = Get-Date
    $script:Occurrence = $now.AddMinutes(-60)
    $occurrenceText = ConvertTo-StateTime -Value $script:Occurrence
    $script:RunningJobs = @{}
    $script:MaintenanceHealthy = $false
    $script:MaintenanceControl = $null
    $script:Manifest = [pscustomobject]@{ OrderedJobs = @([pscustomobject]@{ Name = 'Job-Valuation'; Enabled = $true; AssignmentMode = 'Elected'; DependencyWaitTimeoutMinutes = 360; ConcurrencyKey = '' }) }
    $script:ElectionPlan = [pscustomobject]@{ PlanId = 'P2'; Assignments = @([pscustomobject]@{ JobName = 'Job-Valuation'; OwnerServer = 'SRV-B' }) }
    $script:Settings = [pscustomobject]@{
        SharedDataFolderPath = $root; ExpectedOrchestratorServers = @('SRV-A', 'SRV-B', 'SRV-C'); PeerJobMonitoringEnabled = $true
        PeerHeartbeatStaleMinutes = 5; PeerRecycleGraceMinutes = 15; ElectionServerExclusionMinutes = 15; PeerJobStartGraceMinutes = 15
        DependencyWaitTimeoutMinutes = 360; ElectionClaimsPath = (Join-Path $root 'Claims'); ConcurrencyLeasesPath = (Join-Path $root 'Leases')
        CapabilityMaxAgeMinutes = 60; DistributedSchedulingEnabled = $true; ElectionPlanRefreshSeconds = 60
        ElectionPlanPath = (Join-Path $root 'Election/Orchestrator-ElectionPlan.json'); ElectionPlanLockPath = (Join-Path $root 'Election/Orchestrator-ElectionPlan.lock')
    }
    function Get-Issues { @((Get-OrchestratorPeerHealthSnapshot -Now $now).Issues) }
    function Get-Keys { @(Get-Issues | ForEach-Object { [string]$_.Key }) }
    $waiting = [ordered]@{ Name = 'Job-Valuation'; ScheduledOccurrence = $occurrenceText; Reason = 'WaitingDependencies'; FirstSeen = (ConvertTo-StateTime -Value $now.AddMinutes(-40)); Details = 'Waiting for M365-Licences-Inventory.' }

    # 1. The 30/10 case: the new owner SRV-B has nothing, the previous owner SRV-C (heartbeat 6 min) waits for dependencies.
    Write-Peer 'SRV-B' 1
    Write-Peer 'SRV-C' 6 -Pending @($waiting)
    $keys = Get-Keys
    Check ($keys -notcontains 'JobNotStarted|SRV-B|Job-Valuation') "A wait held by the previous owner must not raise JobNotStarted ($($keys -join ', '))."
    Check ($keys -contains 'HeartbeatStale|SRV-C') 'The stale heartbeat of SRV-C is still reported.'

    # 2. Control: nobody holds the occurrence.
    Write-Peer 'SRV-C' 1
    Check ((Get-Keys) -contains 'JobNotStarted|SRV-B|Job-Valuation') 'An occurrence held by nobody must raise JobNotStarted.'

    # 3. Every waiting reason and a run elsewhere suppress it.
    foreach ($reason in @('PendingRetry', 'WaitingConcurrency', 'BlockedByServerConcurrency', 'LowDiskSpace')) {
        Write-Peer 'SRV-C' 1 -Pending @([ordered]@{ Name = 'Job-Valuation'; ScheduledOccurrence = $occurrenceText; Reason = $reason; Details = '' })
        Check ((Get-Keys) -notcontains 'JobNotStarted|SRV-B|Job-Valuation') "A $reason wait on another server must not raise JobNotStarted."
    }
    Write-Peer 'SRV-C' 1 -Pending @([ordered]@{ Name = 'Job-Valuation'; ScheduledOccurrence = $occurrenceText; Reason = 'BlockedByConcurrencyKey'; BlockSafeUntilUtc = [datetime]::UtcNow.AddHours(1).ToString('o'); Details = '' })
    Check ((Get-Keys) -notcontains 'JobNotStarted|SRV-B|Job-Valuation') 'A valid concurrency block on another server must not raise JobNotStarted.'
    Write-Peer 'SRV-C' 1 -Running @([ordered]@{ Name = 'Job-Valuation'; ScheduledOccurrence = $occurrenceText })
    Check ((Get-Keys) -notcontains 'JobNotStarted|SRV-B|Job-Valuation') 'A run on another server must not raise JobNotStarted.'
    Write-Peer 'SRV-C' 1
    $script:RunningJobs = @{ 'Job-Valuation' = @{ Occurrence = $script:Occurrence } }
    Check ((Get-Keys) -notcontains 'JobNotStarted|SRV-B|Job-Valuation') 'A run on the observer itself must not raise JobNotStarted.'
    $script:RunningJobs = @{}

    # 4. A not-started reason elsewhere, an older occurrence or a server out of the election do not hide it.
    Write-Peer 'SRV-C' 1 -Pending @([ordered]@{ Name = 'Job-Valuation'; ScheduledOccurrence = $occurrenceText; Reason = 'DueNotStarted'; Details = '' })
    Check ((Get-Keys) -contains 'JobNotStarted|SRV-B|Job-Valuation') 'DueNotStarted on another server is still a missed start.'
    Write-Peer 'SRV-C' 1 -Running @([ordered]@{ Name = 'Job-Valuation'; ScheduledOccurrence = (ConvertTo-StateTime -Value $script:Occurrence.AddDays(-1)) })
    Check ((Get-Keys) -contains 'JobNotStarted|SRV-B|Job-Valuation') 'A run of an older occurrence does not cover the expected one.'
    Write-Peer 'SRV-C' 20 -Pending @($waiting)
    Check ((Get-Keys) -contains 'JobNotStarted|SRV-B|Job-Valuation') 'A server out of the election (heartbeat 20 min) does not cover the occurrence.'

    # 5. A dependency wait beyond its timeout is JobWaitingTooLong on the server that holds it, reported once.
    $tooLong = [ordered]@{}; foreach ($key in $waiting.Keys) { $tooLong[$key] = $waiting[$key] }; $tooLong.FirstSeen = ConvertTo-StateTime -Value $now.AddMinutes(-400)
    Write-Peer 'SRV-C' 1 -Pending @($tooLong)
    $issues = Get-Issues
    $waitIssues = @($issues | Where-Object Key -eq 'JobDependencyTimeout|SRV-C|Job-Valuation')
    Check ($waitIssues.Count -eq 1 -and $waitIssues[0].Type -eq 'JobWaitingTooLong' -and $waitIssues[0].Server -eq 'SRV-C') 'A wait beyond the timeout is JobWaitingTooLong on SRV-C.'
    Check (@($issues | Where-Object Type -eq 'JobNotStarted').Count -eq 0) 'A wait beyond the timeout is not JobNotStarted.'

    # 6. Election eligibility: exclusion window instead of the stale threshold.
    $script:ServerCapabilities = [pscustomobject]@{ ServerName = 'SRV-A'; ReadyCapabilities = @('Graph') }
    $script:LiveCapabilityCache = @{}
    Write-Peer 'SRV-B' 7; Write-Peer 'SRV-C' 1
    $names = @(Get-LiveOrchestratorServerCapabilities | ForEach-Object ServerName | Sort-Object)
    Check (($names -join ',') -eq 'SRV-A,SRV-B,SRV-C') "A heartbeat of 7 min keeps the server in the election ($($names -join ','))."
    Write-Peer 'SRV-B' 16
    $names = @(Get-LiveOrchestratorServerCapabilities | ForEach-Object ServerName | Sort-Object)
    Check (($names -join ',') -eq 'SRV-A,SRV-C') 'A heartbeat older than ElectionServerExclusionMinutes leaves the election.'
    $script:Settings.ElectionServerExclusionMinutes = 2
    Write-Peer 'SRV-B' 4
    Check ((Get-OrchestratorElectionExclusionMinutes) -eq 5 -and @(Get-LiveOrchestratorServerCapabilities | Where-Object ServerName -eq 'SRV-B').Count -eq 1) 'The window is never shorter than PeerHeartbeatStaleMinutes.'
    $script:Settings.ElectionServerExclusionMinutes = 15

    # 7. Unreadable heartbeat: last known state kept, then excluded once that state is too old.
    Write-Peer 'SRV-B' 1
    $null = Get-LiveOrchestratorServerCapabilities
    $lock = Lock-File (Join-Path $root 'SRV-B/Orchestrator-Heartbeat.json')
    try {
        $script:Logs.Clear()
        $names = @(Get-LiveOrchestratorServerCapabilities | ForEach-Object ServerName | Sort-Object)
        Check (($names -join ',') -eq 'SRV-A,SRV-B,SRV-C' -and @($script:Logs | Where-Object { $_ -like '`[INFO`] Server SRV-B kept in the election from its last known state*' }).Count -eq 1) 'A locked heartbeat keeps the last known state (INFO).'
        $script:LiveCapabilityCache['SRV-B'].HeartbeatUtc = [datetime]::UtcNow.AddMinutes(-16)
        $script:Logs.Clear()
        $names = @(Get-LiveOrchestratorServerCapabilities | ForEach-Object ServerName | Sort-Object)
        Check (($names -join ',') -eq 'SRV-A,SRV-C' -and @($script:Logs | Where-Object { $_ -like '`[WARN`] Server SRV-B excluded from election*' }).Count -eq 1) 'An old last known state no longer keeps the server.'
    }
    finally { $lock.Dispose() }

    # 8. Short retry: a lock released after 0.6 s is read without error.
    $heartbeatPath = Join-Path $root 'SRV-B/Orchestrator-Heartbeat.json'
    $holder = Start-ThreadJob -ScriptBlock { param($Path) $s = [IO.File]::Open($Path, 'Open', 'Read', 'None'); try { Start-Sleep -Milliseconds 600 } finally { $s.Dispose() } } -ArgumentList $heartbeatPath
    $deadline = [datetime]::UtcNow.AddSeconds(10)
    while ([datetime]::UtcNow -lt $deadline) { try { [IO.File]::ReadAllBytes($heartbeatPath) | Out-Null; Start-Sleep -Milliseconds 20 } catch { break } }
    $document = Read-OrchestratorSharedJsonDocument -Path $heartbeatPath
    $holder | Wait-Job | Remove-Job
    Check ($null -ne $document -and $document.Lifecycle -eq 'Running') 'A sharing violation shorter than the retry is read.'
    Check (-not (Test-OrchestratorTransientReadError -Exception ([InvalidOperationException]::new('x')))) 'A non-IO error is not retried.'

    # 9. Shared plan: last known plan on a sharing violation, fail closed without one, no replan.
    Write-Json $script:Settings.ElectionPlanPath ([ordered]@{ PlanId = 'P1'; GeneratedAtUtc = [datetime]::UtcNow.AddMinutes(-5).ToString('o'); Assignments = @() })
    $script:LastReadElectionPlan = $null
    Check ((Read-SharedElectionPlan).PlanId -eq 'P1' -and -not $script:SharedElectionPlanReadFailed) 'The plan is read.'
    $lock = Lock-File $script:Settings.ElectionPlanPath
    try {
        $script:Logs.Clear()
        $plan = Read-SharedElectionPlan
        Check ($plan.PlanId -eq 'P1' -and $script:SharedElectionPlanReadFailed -and @($script:Logs | Where-Object { $_ -like '`[INFO`] Shared election plan cannot be read; the last known plan P1 is kept*' }).Count -eq 1) 'A locked plan returns the last known plan (INFO).'

        $script:CapabilityCalls = 0; $script:PlanWrites = 0
        function Get-LiveOrchestratorServerCapabilities { $script:CapabilityCalls++; @() }
        function Write-FileAtomically { param($Path, $Content) $script:PlanWrites++ }
        function Read-SharedElectionRebalanceRequest { $null }
        function Test-OrchestratorElectionRebalanceRequestPending { param($Request, $Plan) $false }
        $script:ElectionRebalanceRequested = $false
        $script:Manifest = [pscustomobject]@{ OrderedJobs = @([pscustomobject]@{ Name = 'Job-Valuation'; Enabled = $true; AssignmentMode = 'Elected' }) }
        $script:ElectionPlan = $null
        Update-OrchestratorElectionPlan
        Check ($script:ElectionPlan.PlanId -eq 'P1' -and $script:CapabilityCalls -eq 0 -and $script:PlanWrites -eq 0) 'The planner keeps the last known plan and never replans from an unreadable plan.'
        Check (-not (Test-Path -LiteralPath $script:Settings.ElectionPlanLockPath)) 'The planner lock is released.'

        $script:LastReadElectionPlan = $null
        $script:Logs.Clear()
        Check ($null -eq (Read-SharedElectionPlan) -and @($script:Logs | Where-Object { $_ -like '`[WARN`] Shared election plan cannot be read; elected jobs fail closed*' }).Count -eq 1) 'Without a known plan, elected jobs fail closed.'
    }
    finally { $lock.Dispose() }
}
finally {
    $env:COMPUTERNAME = $originalComputer
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
Write-Host ("Election stability offline tests passed: {0} checks." -f $script:passed) -ForegroundColor Green
