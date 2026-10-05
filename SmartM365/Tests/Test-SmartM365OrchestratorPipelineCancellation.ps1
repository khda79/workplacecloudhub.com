#Requires -Version 7.0
<#
.SYNOPSIS
Offline cancellation tests. Synthetic temporary shared state; no collectors or tenant calls.
.VERSION
1.0.1
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$folder = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
# Standalone orchestrator script module, without a versioned module manifest.
$pipelineModuleFile = (Resolve-Path (Join-Path $folder 'SmartM365.Orchestrator.Pipeline.psm1')).Path
Import-Module $pipelineModuleFile -Force
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-Cancellation-Test-' + [guid]::NewGuid().ToString('N'))
$script:Cases = 0
$worker = $null
$guard = $null
function Assert-Case { param([bool]$Condition, [string]$Message) $script:Cases++; if (-not $Condition) { throw $Message } }
function Assert-Throws {
    param([scriptblock]$Body, [string]$Pattern)
    $failed = $false
    try { & $Body | Out-Null } catch { if ($_.Exception.Message -notlike $Pattern) { throw }; $failed = $true }
    Assert-Case $failed "Expected refusal: $Pattern"
}
function Write-Fixture {
    param([string]$Path, $Document)
    $null = New-Item -ItemType Directory -Path (Split-Path $Path -Parent) -Force
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes(($Document | ConvertTo-Json -Depth 30))
    $null = Write-SmartM365JsonBytesAtomically -Path $Path -Bytes $bytes -Validate { param($value) if ($value -isnot [pscustomobject]) { throw 'Synthetic fixture must be an object.' } }
}
function Write-Heartbeat {
    param([int]$Protocol = 1, [int]$Age = 0, [string]$Lifecycle = 'Running', [string]$Tenant = 'test')
    Write-Fixture (Join-Path $root 'WorkerB/Orchestrator-Heartbeat.json.txt') ([pscustomobject]@{
        Timestamp = [datetime]::UtcNow.AddMinutes(-$Age).ToString('o'); Lifecycle = $Lifecycle
        Tenant = $Tenant; PipelineCancellationProtocol = $Protocol; ScriptVersion = '1.5.41'
    })
}
function New-Request {
    New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $root -Tenant test -Pipeline Jobs -Selection $script:Selection -ManifestHash 'synthetic'
}
function Set-State {
    param([string]$Batch, [string]$Job, [string]$Status)
    $null = Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $root -BatchId $Batch -JobName $Job -Status $Status -OwnerServer WorkerA
}
function Stop-Request {
    param([string]$Batch, [switch]$ValidateOnly)
    Stop-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $root -BatchId $Batch -Tenant test -Reason 'Synthetic operator test' -ValidateOnly:$ValidateOnly
}
try {
    $null = New-Item -ItemType Directory -Path $root
    Write-Fixture (Join-Path $root 'Config/Orchestrator-Cluster.json.txt') ([pscustomobject]@{
        ExpectedOrchestratorServers = @('WorkerA', 'WorkerB'); PeerHeartbeatStaleMinutes = 5
    })
    Write-Heartbeat
    # Create the second synthetic heartbeat folder explicitly.
    Write-Fixture (Join-Path $root 'WorkerA/Orchestrator-Heartbeat.json.txt') ([pscustomobject]@{
        Timestamp = [datetime]::UtcNow.ToString('o'); Lifecycle = 'Running'; Tenant = 'test'; PipelineCancellationProtocol = 1
    })
    $script:Selection = [pscustomobject]@{ SelectedJobs = @(
        [pscustomobject]@{ Name = 'Finished' }, [pscustomobject]@{ Name = 'Waiting' }, [pscustomobject]@{ Name = 'Retry' }
    ) }
    $run = New-Request
    $batch = $run.BatchId
    Set-State $batch Finished Success
    Set-State $batch Retry RetryScheduled
    $successPath = Get-SmartM365OrchestratorPipelineJobStatusPath -SharedDataFolderPath $root -BatchId $batch -JobName Finished
    $successHash = (Get-FileHash $successPath).Hash
    $requestPath = $run.RequestPath
    $requestHash = (Get-FileHash $requestPath).Hash
    $preview = Stop-Request $batch -ValidateOnly
    Assert-Case ($preview.OverallStatus -eq 'Running' -and (Get-FileHash $requestPath).Hash -eq $requestHash) 'Preview wrote state.'
    $cli = Join-Path $folder 'SmartM365-Inventory-Pipeline.ps1'
    $previewOutput = & pwsh -NoProfile -ExecutionPolicy Bypass -File $cli -Tenant test -SharedDataFolderPath $root -JobsManifestPath (Join-Path $root 'unused-manifest.json.txt') -Cancel -BatchId $batch -Reason 'Synthetic CLI preview' -ValidateOnly 2>&1
    Assert-Case ($LASTEXITCODE -eq 0 -and ($previewOutput -join "`n") -match 'no write' -and (Get-FileHash $requestPath).Hash -eq $requestHash) 'CLI preview failed or modified the request.'
    foreach ($bad in @(
        @{ Input = @{ Protocol = 0 }; Detail = '*protocol 1*' },
        @{ Input = @{ Age = 10 }; Detail = '*stale*' },
        @{ Input = @{ Age = -10 }; Detail = '*future*' },
        @{ Input = @{ Lifecycle = 'Starting' }; Detail = '*Lifecycle=Starting*' },
        @{ Input = @{ Lifecycle = 'Recycling' }; Detail = '*Lifecycle=Recycling*' },
        @{ Input = @{ Tenant = 'different' }; Detail = '*tenant*' }
    )) {
        $inputParameters = $bad.Input
        Write-Heartbeat @inputParameters
        $readiness = @(Get-SmartM365OrchestratorPipelineCancellationReadiness -SharedDataFolderPath $root -Tenant test | Where-Object Server -eq WorkerB)
        Assert-Case ($readiness.Count -eq 1 -and -not $readiness[0].Ready -and $readiness[0].Detail -like $bad.Detail) 'Cancellation refusal did not identify the failing readiness check.'
        Assert-Throws { Stop-Request $batch } '*Cancellation unavailable*'
        Assert-Case ((Get-FileHash $requestPath).Hash -eq $requestHash) 'Readiness failure changed the request.'
    }
    Remove-Item -LiteralPath (Join-Path $root 'WorkerB/Orchestrator-Heartbeat.json.txt')
    Assert-Throws { Stop-Request $batch } '*Cancellation unavailable*'
    Write-Heartbeat
    Assert-Throws { Stop-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $root -BatchId $batch -Tenant other -Reason test } '*tenant*'
    Assert-Throws { Stop-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $root -BatchId $batch -Tenant test -Reason '  ' } '*reason*'
    Assert-Throws { Stop-Request '../escape' } '*Invalid pipeline BatchId*'
    $cancelled = Stop-Request $batch
    Assert-Case ($cancelled.OverallStatus -eq 'Cancelled' -and $cancelled.IsTerminal -and $cancelled.CancelledCount -eq 2) 'Pending/retry cancellation aggregation failed.'
    Assert-Case ((Get-FileHash $successPath).Hash -eq $successHash) 'Completed result bytes changed.'
    Assert-Case ($cancelled.Request.Cancellation.Reason -eq 'Synthetic operator test' -and $cancelled.Request.Cancellation.RequestedBy) 'Cancellation audit missing.'
    $auditHash = (Get-FileHash $requestPath).Hash
    $null = Stop-Request $batch
    Assert-Case ((Get-FileHash $requestPath).Hash -eq $auditHash) 'Idempotent cancellation rewrote audit.'
    $waitingPath = Get-SmartM365OrchestratorPipelineJobStatusPath -SharedDataFolderPath $root -BatchId $batch -JobName Waiting
    $partial = (Read-SmartM365JsonDocument -Path $waitingPath).Document
    $partial.Status = 'Pending'
    Write-Fixture $waitingPath $partial
    $resumed = Stop-Request $batch
    Assert-Case ($resumed.CancelledCount -eq 2 -and (Get-FileHash $requestPath).Hash -eq $auditHash) 'Partial publication did not resume without rewriting audit.'
    foreach ($status in @('Pending', 'Starting', 'Running', 'RetryScheduled', 'Success', 'Failed')) {
        Set-State $batch Waiting $status
        $after = Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $root -BatchId $batch
        Assert-Case (@($after.Jobs | Where-Object { $_.JobName -eq 'Waiting' -and $_.Status -eq 'Cancelled' }).Count -eq 1) "Cancelled job resurrected as $status."
    }
    $guard = Enter-SmartM365OrchestratorPipelineLaunchGuard -SharedDataFolderPath $root -BatchId $batch -JobName Waiting
    Assert-Case (-not $guard.Allowed) 'Stale queue passed the launch fence.'
    Exit-SmartM365OrchestratorPipelineLaunchGuard $guard; $guard = $null

    # A real competing runspace must wait for the launch fence, then preserve the started job.
    $run = New-Request; $batch = $run.BatchId
    # Resource refusal releases the reservation without dropping a pending retry.
    Set-State $batch Retry RetryScheduled
    $guard = Enter-SmartM365OrchestratorPipelineLaunchGuard -SharedDataFolderPath $root -BatchId $batch -JobName Retry
    $reserved = Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $root -BatchId $batch
    Assert-Case (@($reserved.Jobs | Where-Object { $_.JobName -eq 'Retry' -and $_.Status -eq 'Starting' }).Count -eq 1) 'Process admission was not persisted before launch.'
    Exit-SmartM365OrchestratorPipelineLaunchGuard $guard; $guard = $null
    $restored = Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $root -BatchId $batch
    Assert-Case (@($restored.Jobs | Where-Object { $_.JobName -eq 'Retry' -and $_.Status -eq 'RetryScheduled' }).Count -eq 1) 'No-launch reservation lost its original retry.'
    $guard = Enter-SmartM365OrchestratorPipelineLaunchGuard -SharedDataFolderPath $root -BatchId $batch -JobName Waiting
    Assert-Case $guard.Allowed 'Pending job was denied.'
    $worker = [powershell]::Create()
    $null = $worker.AddScript({
        param($Module, $Root, $Batch)
        Import-Module $Module -Force
        Stop-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $Root -BatchId $Batch -Tenant test -Reason 'Synthetic concurrent cancellation' -LockTimeoutSeconds 15
    }).AddArgument($pipelineModuleFile).AddArgument($root).AddArgument($batch)
    $handle = $worker.BeginInvoke()
    # Wait until the worker holds the submission lock, proving that it reached cancellation.
    $paths = Get-SmartM365OrchestratorPipelinePaths -SharedDataFolderPath $root -BatchId $batch
    $deadline = [datetime]::UtcNow.AddSeconds(10)
    while (-not (Test-Path $paths.SubmissionLockPath) -and [datetime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 50 }
    Assert-Case ((Test-Path $paths.SubmissionLockPath) -and -not $handle.IsCompleted) 'Competing cancellation did not wait for process admission.'
    Set-State $batch Waiting Running
    Exit-SmartM365OrchestratorPipelineLaunchGuard $guard; $guard = $null
    Assert-Case ($handle.AsyncWaitHandle.WaitOne(15000)) 'Cancellation worker did not finish.'
    $result = @($worker.EndInvoke($handle))
    if ($worker.HadErrors) { throw ($worker.Streams.Error | Out-String) }
    $worker.Dispose(); $worker = $null
    $after = Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $root -BatchId $batch
    Assert-Case ($after.OverallStatus -eq 'Cancelling' -and $after.PendingCount -eq 1 -and $after.CancelledCount -eq 2) 'Running collector was cancelled or aggregate was terminal too early.'
    Assert-Throws { New-Request } '*active pipeline request*'

    # Execute the real resident launch/completion functions with mocked process/transport.
    $tokens = $null; $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile((Join-Path $folder 'SmartM365-Inventory-Orchestrator.ps1'), [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors.Count) { throw 'Resident parser errors.' }
    $definitions = foreach ($name in @('Invoke-LaunchPhase', 'Complete-JobRun', 'ConvertFrom-StateTime', 'ConvertTo-StateTime')) {
        $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
        if ($null -eq $node) { throw "Missing resident function: $name" }
        $node.Extent.Text
    }
    $runtime = New-Module -ScriptBlock ([scriptblock]::Create($definitions -join "`n"))
    try {
        $runtimeResult = & $runtime {
            param($Root, $Batch)
            $script:Settings = @{ SharedDataFolderPath = $Root; MaxConcurrency = 4; JobMailMode = 'Never' }
            $script:MaintenanceHealthy = $true; $script:StatePersistenceHealthy = $true
            $script:MaintenanceControl = [pscustomobject]@{ Enabled = $true }
            $script:Starts = 0; $script:Leases = 0; $script:RowStatus = ''; $script:PipelineStatus = ''
            $script:ForcedPending = @(); $script:RunningJobs = @{}; $script:ElectionPlan = $null
            $job = [pscustomobject]@{ Name = 'Retry'; Enabled = $true; AssignmentMode = 'Legacy'; DependsOn = @(); ConcurrencyKey = 'synthetic'; TimeoutMinutes = 60; MaxRetries = 2; RetryDelaySeconds = 60 }
            $script:Manifest = @{ OrderedJobs = @($job); JobsByName = @{ Retry = $job } }
            $script:LocalState = @{ PendingRetry = $null; LastScheduledOccurrence = ''; Running = $null }
            $script:PipelinePending = @{ Retry = @{ BatchId = $Batch; Occurrence = (Get-Date); Attempt = 0; CreatedAtUtc = [datetime]::UtcNow; StatusPath = 'synthetic' } }
            $script:PipelineBatchJobNames = @{}
            function script:Get-JobState { param($JobName) $script:LocalState }
            function script:Test-JobSelected { param($JobName) $true }
            function script:Test-JobAllowedOnServer { param($Job,[switch]$AllowManual) $true }
            function script:Save-OrchestratorState {}
            function script:Write-OrchestratorLog { param($Message,$Level) }
            function script:Write-OrchestratorRuntimeUpdateWarning { param($Key,$Message,$Now) }
            function script:Clear-DependencyWaitLog { param($JobName) }
            function script:Enter-SmartM365OrchestratorMaintenanceGate { param($SharedDataFolderPath) [IO.MemoryStream]::new() }
            function script:Get-SmartM365OrchestratorMaintenanceState { param($SharedDataFolderPath) $script:MaintenanceControl }
            function script:Test-SmartM365OrchestratorMaintenanceLaunch { param($State,$Origin,$Occurrence) $true }
            function script:Enter-SmartM365OrchestratorConcurrencyLease { param($LeasesRootPath,$ConcurrencyKey,$JobName,$Occurrence,$OwnerServer,$SafeMinutes,$HeartbeatRootPath,$HeartbeatStaleMinutes) $script:Leases++; throw 'Unexpected lease acquisition for cancelled work.' }
            function script:Start-InventoryJob { param($Job,$Occurrence,$Attempt,$ClaimPath,$ConcurrencyLeasePath,$ConcurrencyLeaseId,$PipelineBatchId,$PipelineStatusPath) $script:Starts++; throw 'Unexpected collector launch.' }
            Invoke-LaunchPhase (Get-Date)
            $script:PipelinePending = @{}
            $script:LocalState.PendingRetry = @{ PipelineBatchId = $Batch; Attempt = 1; ScheduledOccurrence = (Get-Date).ToString('o'); NotBefore = (Get-Date).ToString('o') }
            Invoke-LaunchPhase (Get-Date)
            $cleared = $null -eq $script:LocalState.PendingRetry
            function script:Set-OrchestratorPipelineJobStatus { param($BatchId,$JobName,$Status,$Attempt,$Detail,$NotBeforeUtc) $script:PipelineStatus = $Status }
            function script:Add-JobRunCsvRow { param($JobName,$ScheduledTime,$StartTime,$EndTime,$DurationSec,$ExitCode,$Status,$RetryCount,$LogPath) $script:RowStatus = $Status }
            function script:Invoke-OrchestratorSharePointUpload { param($LocalFilePath,$Reason,[switch]$Force) $false }
            function script:Get-JobRunsCsvPath { 'synthetic.csv' }
            $info = @{ PipelineBatchId = $Batch; Attempt = 0; StartTime = (Get-Date).AddMinutes(-1); Occurrence = (Get-Date); LogPath = 'synthetic.log' }
            Complete-JobRun -JobName Retry -RunInfo $info -StatusHint Exited -ExitCode 1 -EndTime (Get-Date)
            [pscustomobject]@{ Starts = $script:Starts; Leases = $script:Leases; RetryCleared = $cleared; Retry = $script:LocalState.PendingRetry; CsvStatus = $script:RowStatus; PipelineStatus = $script:PipelineStatus }
        } $root $batch
        Assert-Case ($runtimeResult.Starts -eq 0 -and $runtimeResult.Leases -eq 0) 'Real resident launched cached cancelled work or acquired a lease.'
        Assert-Case ($runtimeResult.RetryCleared -and $null -eq $runtimeResult.Retry) 'Resident retained or scheduled a cancelled pipeline retry.'
        Assert-Case ($runtimeResult.CsvStatus -eq 'Failed' -and $runtimeResult.PipelineStatus -eq 'Failed') 'Resident concealed real failure after cancellation.'
    }
    finally { Remove-Module $runtime }
    Set-State $batch Waiting Success
    $after = Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $root -BatchId $batch
    Assert-Case ($after.IsTerminal -and $after.OverallStatus -eq 'Cancelled' -and @($after.Jobs | Where-Object Status -eq Success).Count -eq 1) 'Running completion was not preserved.'
    Import-Module (Join-Path $folder 'SmartM365.Orchestrator.Insights.psm1') -Force
    $rows = @(Get-SmartM365OrchestratorRecentPipelineRuns -SharedDataFolderPath $root)
    Assert-Case (@($rows | Where-Object { $_.BatchId -eq $batch -and $_.Status -eq 'Cancelled' -and $_.Cancelled -eq 2 }).Count -eq 1) 'GUI aggregation differs from pipeline.'

    $run = New-Request; $batch = $run.BatchId
    Set-State $batch Finished Failed
    $after = Stop-Request $batch
    Assert-Case ($after.OverallStatus -eq 'Failed' -and $after.FailedCount -eq 1 -and $after.CancelledCount -eq 2) 'Cancellation concealed an existing failure.'
    $run = New-Request; $batch = $run.BatchId
    foreach ($name in @('Finished', 'Waiting', 'Retry')) { Set-State $batch $name Success }
    Assert-Throws { Stop-Request $batch } '*already terminal*'
    Write-Output "PASS cancellation offline: $script:Cases assertions; synthetic data only."
}
finally {
    if ($null -ne $guard) { Exit-SmartM365OrchestratorPipelineLaunchGuard $guard }
    if ($null -ne $worker) { $worker.Stop(); $worker.Dispose() }
    # Exact GUID-named fixture under the temporary directory, never a workspace/data root.
    if ([IO.Path]::GetFullPath($root).StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()), [StringComparison]::OrdinalIgnoreCase) -and (Split-Path $root -Leaf) -match '^SmartM365-Cancellation-Test-[0-9a-f]{32}$') {
        Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBJ1Ujf/1i/Hz94
# l4J7DghLEUqo5VxWF/JHUWGKel8SOqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKKpioRTHHSfc4gB1zvzGlpftfJngBEO7v5sou2QHY7rMA0GCSqG
# SIb3DQEBAQUABIIBgEU3iH5skuSUn4m6tj/v3v8iQD//B21m9ncEmzIKCAT4Ha/N
# SA52fEJBlkxhuXTxLGMblgoa67DG4/eYhAuyJhv9Hv60NbAwyPZKjWWMDrrDLD0N
# Q7l7pyj1zWOWxiXrZWwitJ9XOn81Qp7mzDwteONa0Gsr4oPJWmEBd/LzAaYnGpy6
# 1GeEVtsRTUti3QmmrFcuc7Cgm9nRa+bsaFWQERO/SVbFNpPqbGDxwk/idIPtpoRV
# 51mCpThOkBeIriEr5xGpcbFLfs895txXN8aNEUoB1IF9fM6p08Uk4Jqq/v0lKRoe
# JkZIWQTdmdWKuuxF+0jCEPOpKwcv2GhlhCmoG4fNx6ot4okG1aQ+4wLQqOL8QZN9
# JndLLRApvw+NprMFyfwDJVjcBjJd7Q/FE74Nx5OPceoPO5XE2KU8cQfVywBamGHT
# rzg8cUSZw1ejnErRLKa5p4a0TaE3gHQHSEXJ0wF+HwXBprz7D/UCUQucV0okN3oT
# 64UOzmgjtHRgsInkRKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDUxOTM0
# NTFaMC8GCSqGSIb3DQEJBDEiBCBAl5KbtEwcPKKHAu8IvqnJkw1XqmEs0D3jhBHv
# WUDlnzANBgkqhkiG9w0BAQEFAASCAgCQXu/Xk2N17IMqAgL+k1VVZiKn9fwXNfRz
# IC3n55gUJT01N+h767axBDS/lcFrh9G3IYrGbhNTydWnKUyW5htlDrVeOdxXSFFe
# meg4VMy6lXVOhVOeT08RRKNfbnAgKHARqsrCqF+YIaNZ4FRSruG6Mw3R//dHvJk1
# DOh1Jmq8DdOhMg03DMQfc+IDWymVbMozv6BcD9mEwQhqrw6RvHZCLWWB/1tVBut2
# NTKmO5pJWV99YAfW1jO1XceRRgE4lPXEPMqE6pN1qxA7j3ArKS40g8qcb2F9yV/U
# +pSwNhSs6ETifkX/bSD70RsTLir1Uv+zWlIxfI5IE8KTMuw4gfi96sUPcwZls7Os
# Z9g1psfSwg8aWQubfyf9lY1ZVXWzjGfIymh57Vf4/wyqKWkTcFKqxQ86SqjyC9tJ
# nZNC8leEdXCLWWB/czLBTg8B1jNfksGox6nWo1P+Rh8GYl56SlHTeDYr63jXkiUL
# qEb3qbBxdm75NIuMgOCrgIvHaB+Porc4UaeVoWzipoqkpzMcfZFbmkvkBQ48UpPI
# iyF6blEpMraH7EdNlaEw9L7C193wzfNbGJRR2U7M6x5bx9WeHitdOlwqEN0x616J
# fELXjbjGrmImfMcXmjpd5Kc8krHlVAN6MW7nwaScvhj/medoJBZOc8uAUnLmMJwt
# y7C/NCNVJw==
# SIG # End signature block
