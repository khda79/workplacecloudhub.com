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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCiaLnb7ZTr0EEx
# jU1V8Kc4ddMVMUoDnjybH4oda2y4zqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILeZhWyEWFO+mkJtdAm1srh9WN4uIrqA2h1QfxAwMeb/MA0GCSqG
# SIb3DQEBAQUABIIBgA3IQv8HzyJvhsNUmJfjENg8fWQiEwf4DLza6zzgV720rKKF
# ES5eJYmWyBbXJfNrDoKVKI5CKdrwvbEnxDvbS3JxL0MmXc4XmfOOwlwyv3U+wsv6
# Xaj/bACNZ189phCvN1L/5Oyw0GOYTtv8g9iOvWv7RxJWUFUj21pvxjuUOeJDNx4G
# d93DCqdhJQsPSmJL35UMjvNfKEc2r6j3/DJJmp1p63JC5Tr7fwHtVnEVj16UEjbe
# dh+o4e1HBUayPgFtajViN7ng0/oxTesC2wgPixiL0qRkWdih7jiylEPsYpWSjcSl
# 8eIdi212WaHlYPP0ySlGbYW8d2HhwWfd5odg4lGRGzFh00ExkJ4LAmWY2Vmeatly
# +5TvLyvevHBXB5TXf6CviJl9277+GiLcCczr8vY8r2mk2/xD1lRDTGjuRIY0HXiV
# JovpaCv3lv6zuc7ERBNBwCp8RHAcmseLjwxkt40dGIp3TXxjEH5GtrHxJU1eilP/
# 4QP3TZ+Y9ojHzMYioKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzAwODMw
# MjlaMC8GCSqGSIb3DQEJBDEiBCD6ENMjdIQKJdkGdLJSrI5MRY25RudTUA8PiN5i
# y/IA7jANBgkqhkiG9w0BAQEFAASCAgBE2ufn81pf0sw/jR9XqnH6f1ItVv5cFrur
# B3YM2xZcRTj/NHiI/cVFrhe0zYZPDXknqc/2qMCAX3L10MyP7L0NdS87vc4HpeSs
# q+z43qm+v6jGGU9F7jVcBlX+ecm3sbQtn3dLsMlj7bYCXCAnbZpckKUVFiHYu9bh
# UX5SCJuxIvZ5omneoamDFKDiTfJYMAkgVsfkzM7ZaWmDUCpKWsFVb7VW2BT4HsLt
# o5n5yFjyL9LQNoeIrOpgV+HUMh59GA/6lbTuumOfg3zEkdiFgVPncZ70VBZC3Z9v
# puj4BNUc+aDifO/TiF5tGW1pR3VZPuYgtQvoUkq7k+pRM1a5bPFhiBmrDdESI+GJ
# zUxLTmoEzWviRmzBVsaHd36u9inbePx4njViKY9NxL/UhD3aEES3hIk1yaoRE2ol
# +ckjojXgVt5QoU6jZ5o07CnejsjIOsETl+RoxyKLjGLzXQeBEf8qko+LY1FCtdEk
# igYfdNj4JUCWeuYqswjbT24JnMQYQ/4IjOOP2myu43IxISUsaa/Ix7ol/QDI2yk1
# qErxarJWirRmU8OmafiOO25zKTJIQMJB0xnQtxE40KugpJ+wa+/pJeGLZU101HbN
# tLIHC+bJBAPJV02uTg07v9QIWWBWMJErLQvhkoAeo2VVv7kfWM19hfgxUAL2d07A
# eRPVcM7wPQ==
# SIG # End signature block
