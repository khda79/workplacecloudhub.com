<#
.SYNOPSIS
Runs offline management and central-manifest migration tests for the SmartM365 orchestrator.
.VERSION
1.1.2
#>
#Requires -Version 7.0
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$smartM365Root = Split-Path -Path $PSScriptRoot -Parent
$managementModuleFile = Join-Path -Path $smartM365Root -ChildPath 'SmartInventory\Orchestrator\SmartM365.Orchestrator.Management.psm1'
$jobsTemplatePath = Join-Path -Path $smartM365Root -ChildPath 'SmartInventory\Orchestrator\Orchestrator-Jobs.json.template'
Import-Module -Name $managementModuleFile -Force -ErrorAction Stop

function Assert-True {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )
    if (-not $Condition) { throw $Message }
}

$temporaryRoot = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ('SmartM365-OrchestratorManagement-' + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path $temporaryRoot -Force | Out-Null
    $cluster = [pscustomobject][ordered]@{
        SchemaVersion = 1
        ExpectedOrchestratorServers = @('SERVER-A', 'SERVER-B')
        ElectionWeightsByServer = [pscustomobject]@{ 'SERVER-A' = 1.0; 'SERVER-B' = 1.1 }
        ServerJobPolicies = [pscustomobject]@{
            'SERVER-B' = [pscustomobject]@{ OnlyJobsRequiring = @('ExchangeOnPrem') }
        }
        PeerMonitoringEnabled = $true
        PeerJobMonitoringEnabled = $true
        PeerMonitoringCheckIntervalSeconds = 60
        PeerHeartbeatStaleMinutes = 5
        PeerMonitoringConfirmationChecks = 2
        PeerJobStartGraceMinutes = 15
        PeerAlertReminderMinutes = 240
        PeerAlertMailRetryMinutes = 15
        PeerRecoveryEmailEnabled = $true
    }
    $snapshot = Initialize-SmartM365OrchestratorCentralConfiguration -SharedDataFolderPath $temporaryRoot -BootstrapJobsPath $jobsTemplatePath -BootstrapClusterDocument $cluster
    Assert-True -Condition (Test-Path -LiteralPath $snapshot.Paths.JobsPath) -Message 'Central jobs configuration was not initialized.'
    Assert-True -Condition (Test-Path -LiteralPath $snapshot.Paths.ClusterPath) -Message 'Central cluster configuration was not initialized.'
    Assert-True -Condition ((Test-SmartM365OrchestratorJobsDocument -Document $snapshot.Jobs).Valid) -Message 'Production jobs template failed management validation.'
    Assert-True -Condition ((Test-SmartM365OrchestratorClusterDocument -Document $snapshot.Cluster).Valid) -Message 'Mock cluster configuration failed validation.'

    $firstRebalanceRequest = Request-SmartM365OrchestratorRebalance -SharedDataFolderPath $temporaryRoot -Reason 'Management test'
    $savedRebalanceRequest = Read-SmartM365OrchestratorJson -Path $firstRebalanceRequest.RequestPath
    Assert-True -Condition ([string]$savedRebalanceRequest.RequestId -ceq [string]$firstRebalanceRequest.RequestId) -Message 'The rebalance request was not persisted atomically.'
    Assert-True -Condition ([string]$savedRebalanceRequest.Reason -eq 'Management test') -Message 'The rebalance request reason was not preserved.'
    $secondRebalanceRequest = Request-SmartM365OrchestratorRebalance -SharedDataFolderPath $temporaryRoot -Reason 'Second management test'
    Assert-True -Condition ([string]$secondRebalanceRequest.RequestId -cne [string]$firstRebalanceRequest.RequestId) -Message 'A second rebalance request did not receive a unique identifier.'
    Assert-True -Condition ([string](Read-SmartM365OrchestratorJson -Path $secondRebalanceRequest.RequestPath).RequestId -ceq [string]$secondRebalanceRequest.RequestId) -Message 'The latest rebalance request did not replace the earlier request.'

    $managementLogPath = Join-Path -Path $temporaryRoot -ChildPath 'Logs\SmartM365-Orchestrator-GUI_TEST.log'
    Write-SmartM365OrchestratorManagementLog -Path $managementLogPath -Message "First line`nSecond line" -Level WARN
    $managementLogLines = @(Get-Content -LiteralPath $managementLogPath)
    Assert-True -Condition ($managementLogLines.Count -eq 2) -Message 'The persistent GUI log did not preserve one timestamped record per physical line.'
    Assert-True -Condition ($managementLogLines[0] -match '^\[\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d{3}\]\[WARN\] First line$') -Message 'The persistent GUI log prefix is invalid.'

    $readOnlyPath = Join-Path -Path $temporaryRoot -ChildPath 'ReadOnly-Replacement.json'
    Write-SmartM365OrchestratorJsonAtomically -Path $readOnlyPath -Document ([pscustomobject]@{ Value = 'before' })
    (Get-Item -LiteralPath $readOnlyPath).IsReadOnly = $true
    try {
        Write-SmartM365OrchestratorJsonAtomically -Path $readOnlyPath -Document ([pscustomobject]@{ Value = 'after' }) -RetrySeconds 0
    }
    finally {
        if (Test-Path -LiteralPath $readOnlyPath) { (Get-Item -LiteralPath $readOnlyPath).IsReadOnly = $false }
    }
    Assert-True -Condition ((Read-SmartM365OrchestratorJson -Path $readOnlyPath).Value -eq 'after') -Message 'The SMB-compatible replacement did not replace a read-only destination with -Force.'

    $persistentlyLockedPath = Join-Path -Path $temporaryRoot -ChildPath 'Persistently-Locked.json'
    Write-SmartM365OrchestratorJsonAtomically -Path $persistentlyLockedPath -Document ([pscustomobject]@{ Value = 'preserved' })
    $lockedHash = Get-SmartM365OrchestratorFileHash -Path $persistentlyLockedPath
    $persistentLock = [IO.File]::Open($persistentlyLockedPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    $persistentLockError = ''
    try {
        try {
            Write-SmartM365OrchestratorJsonAtomically -Path $persistentlyLockedPath -Document ([pscustomobject]@{ Value = 'must-not-publish' }) -RetrySeconds 0
        }
        catch {
            $persistentLockError = $_.Exception.Message
        }
    }
    finally {
        $persistentLock.Dispose()
    }
    Assert-True -Condition ($persistentLockError -like "*Atomic replacement failed for '$persistentlyLockedPath'*") -Message 'A persistent destination lock did not report the exact target path.'
    Assert-True -Condition ((Get-SmartM365OrchestratorFileHash -Path $persistentlyLockedPath) -eq $lockedHash) -Message 'A failed locked-file replacement changed the valid destination.'
    Assert-True -Condition (@(Get-ChildItem -LiteralPath $temporaryRoot -Filter 'Persistently-Locked.json.*.tmp' -File).Count -eq 0) -Message 'A failed locked-file replacement left a temporary file behind.'

    $transientlyLockedPath = Join-Path -Path $temporaryRoot -ChildPath 'Transiently-Locked.json'
    $transientReadyPath = Join-Path -Path $temporaryRoot -ChildPath 'Transiently-Locked.ready'
    Write-SmartM365OrchestratorJsonAtomically -Path $transientlyLockedPath -Document ([pscustomobject]@{ Value = 'before' })
    $transientLockJob = Start-Job -ScriptBlock {
        $lockPath = [string]$args[0]
        $readySignalPath = [string]$args[1]
        $stream = [IO.File]::Open($lockPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
        try {
            [IO.File]::WriteAllText($readySignalPath, 'ready')
            Start-Sleep -Milliseconds 900
        }
        finally {
            $stream.Dispose()
        }
    } -ArgumentList $transientlyLockedPath, $transientReadyPath
    try {
        $readyDeadline = (Get-Date).AddSeconds(10)
        while (-not (Test-Path -LiteralPath $transientReadyPath) -and (Get-Date) -lt $readyDeadline) { Start-Sleep -Milliseconds 50 }
        Assert-True -Condition (Test-Path -LiteralPath $transientReadyPath) -Message 'The transient-lock test did not acquire its destination handle.'
        Write-SmartM365OrchestratorJsonAtomically -Path $transientlyLockedPath -Document ([pscustomobject]@{ Value = 'after-retry' }) -RetrySeconds 5
    }
    finally {
        Wait-Job -Job $transientLockJob -Timeout 10 | Out-Null
        Receive-Job -Job $transientLockJob -ErrorAction SilentlyContinue | Out-Null
        Remove-Job -Job $transientLockJob -Force -ErrorAction SilentlyContinue
    }
    Assert-True -Condition ((Read-SmartM365OrchestratorJson -Path $transientlyLockedPath).Value -eq 'after-retry') -Message 'The bounded retry did not recover after a transient destination lock.'

    $upgradeRoot = Join-Path -Path $temporaryRoot -ChildPath 'ManifestUpgrade'
    New-Item -ItemType Directory -Path $upgradeRoot -Force | Out-Null
    $upgradeManifestPath = Join-Path -Path $upgradeRoot -ChildPath 'Orchestrator-Jobs.json'
    $upgradeTemplatePath = Join-Path -Path $upgradeRoot -ChildPath 'Orchestrator-Jobs.json.template'
    $existingJob = $snapshot.Jobs.Jobs[0] | ConvertTo-Json -Depth 100 | ConvertFrom-Json -Depth 100
    $existingJob.Enabled = $false
    $existingJob.Schedule.Times = @('23:00')
    $existingJob.Arguments = '-CustomArgument example'
    $newJob = $snapshot.Jobs.Jobs[1] | ConvertTo-Json -Depth 100 | ConvertFrom-Json -Depth 100
    $upgradeDocument = [pscustomobject][ordered]@{ SchemaVersion = 1; Jobs = @($existingJob) }
    $upgradeTemplate = [pscustomobject][ordered]@{
        SchemaVersion = 1
        Jobs = @(
            ($snapshot.Jobs.Jobs[0] | ConvertTo-Json -Depth 100 | ConvertFrom-Json -Depth 100),
            $newJob
        )
    }
    $upgradeTemplate.Jobs[0].Arguments = '-EnableConfiguredExternalActions'
    Write-SmartM365OrchestratorJsonAtomically -Path $upgradeManifestPath -Document $upgradeDocument
    Write-SmartM365OrchestratorJsonAtomically -Path $upgradeTemplatePath -Document $upgradeTemplate
    $upgradeResult = Sync-SmartM365OrchestratorJobsManifest -Path $upgradeManifestPath -TemplatePath $upgradeTemplatePath
    $upgradedDocument = Read-SmartM365OrchestratorJson -Path $upgradeManifestPath
    Assert-True -Condition ($upgradeResult.Updated -and -not $upgradeResult.Created) -Message 'The existing manifest was not upgraded.'
    Assert-True -Condition (@($upgradeResult.AddedJobNames).Count -eq 1 -and $upgradeResult.AddedJobNames[0] -eq $newJob.Name) -Message 'The missing job was not reported correctly.'
    Assert-True -Condition (@($upgradedDocument.Jobs).Count -eq 2) -Message 'The missing job was not appended.'
    Assert-True -Condition (-not [bool]$upgradedDocument.Jobs[0].Enabled) -Message 'The existing Enabled override was overwritten.'
    Assert-True -Condition ($upgradedDocument.Jobs[0].Schedule.Times[0] -eq '23:00') -Message 'The existing schedule override was overwritten.'
    Assert-True -Condition ([string]$upgradedDocument.Jobs[0].Arguments -eq '-CustomArgument example -EnableConfiguredExternalActions') -Message 'The required external-action opt-in was not merged with existing arguments.'
    Assert-True -Condition (@($upgradeResult.UpdatedJobNames) -contains $existingJob.Name) -Message 'The external-action argument migration was not reported.'
    $secondUpgradeResult = Sync-SmartM365OrchestratorJobsManifest -Path $upgradeManifestPath -TemplatePath $upgradeTemplatePath
    Assert-True -Condition (-not $secondUpgradeResult.Updated -and @($secondUpgradeResult.AddedJobNames).Count -eq 0 -and @($secondUpgradeResult.UpdatedJobNames).Count -eq 0) -Message 'A second manifest synchronization was not idempotent.'
    $convertedPolicies = ConvertTo-SmartM365OrchestratorHashtable -InputObject $snapshot.Cluster.ServerJobPolicies
    Assert-True -Condition ($convertedPolicies['SERVER-B']['OnlyJobsRequiring'] -eq 'ExchangeOnPrem') -Message 'String values were corrupted while converting the server policy to a hashtable.'
    $invalidPinnedJobs = $snapshot.Jobs | ConvertTo-Json -Depth 100 | ConvertFrom-Json -Depth 100
    $invalidPinnedJobs.Jobs[0].AssignmentMode = 'Pinned'
    $invalidPinnedJobs.Jobs[0].AllowedServers = @('SERVER-NOT-IN-CLUSTER')
    $consistency = Test-SmartM365OrchestratorConfigurationConsistency -JobsDocument $invalidPinnedJobs -ClusterDocument $snapshot.Cluster
    Assert-True -Condition (-not $consistency.Valid) -Message 'A pinned job targeting a non-cluster server was accepted.'

    $invalidJobs = $snapshot.Jobs | ConvertTo-Json -Depth 100 | ConvertFrom-Json -Depth 100
    $invalidJobs.Jobs[0].Schedule.Times = @('25:99')
    Assert-True -Condition (-not (Test-SmartM365OrchestratorJobsDocument -Document $invalidJobs).Valid) -Message 'An invalid time was accepted.'

    $publishedJobs = $snapshot.Jobs | ConvertTo-Json -Depth 100 | ConvertFrom-Json -Depth 100
    $publishedJobs.Jobs[0].Enabled = -not [bool]$publishedJobs.Jobs[0].Enabled
    $unchangedClusterLock = [IO.File]::Open($snapshot.Paths.ClusterPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    try {
        $publishResult = Publish-SmartM365OrchestratorConfiguration `
            -SharedDataFolderPath $temporaryRoot `
            -JobsDocument $publishedJobs `
            -ClusterDocument $snapshot.Cluster `
            -ExpectedJobsHash $snapshot.JobsHash `
            -ExpectedClusterHash $snapshot.ClusterHash `
            -ChangeSummary 'Management test' `
            -AtomicWriteRetrySeconds 0
    }
    finally {
        $unchangedClusterLock.Dispose()
    }
    Assert-True -Condition (Test-Path -LiteralPath $publishResult.VersionFolderPath) -Message 'No configuration version was created.'
    Assert-True -Condition (Test-Path -LiteralPath $snapshot.Paths.AuditPath) -Message 'No configuration audit CSV was created.'
    Assert-True -Condition ($publishResult.JobsChanged -and -not $publishResult.ClusterChanged) -Message 'A jobs-only publication did not skip the unchanged locked cluster document.'

    $conflictDetected = $false
    try {
        Publish-SmartM365OrchestratorConfiguration `
            -SharedDataFolderPath $temporaryRoot `
            -JobsDocument $publishedJobs `
            -ClusterDocument $snapshot.Cluster `
            -ExpectedJobsHash $snapshot.JobsHash `
            -ExpectedClusterHash $snapshot.ClusterHash `
            -ChangeSummary 'Expected conflict' | Out-Null
    }
    catch { $conflictDetected = $_.Exception.Message -like '*changed after it was loaded*' }
    Assert-True -Condition $conflictDetected -Message 'Optimistic concurrency did not reject a stale publication.'

    $rollbackSnapshot = Get-SmartM365OrchestratorConfigurationSnapshot -SharedDataFolderPath $temporaryRoot
    $rollbackJobs = $rollbackSnapshot.Jobs | ConvertTo-Json -Depth 100 | ConvertFrom-Json -Depth 100
    $rollbackJobs.Jobs[1].Enabled = -not [bool]$rollbackJobs.Jobs[1].Enabled
    $rollbackCluster = $rollbackSnapshot.Cluster | ConvertTo-Json -Depth 100 | ConvertFrom-Json -Depth 100
    $rollbackCluster.PeerAlertReminderMinutes = [int]$rollbackCluster.PeerAlertReminderMinutes + 1
    $versionsBeforeFailure = @(Get-ChildItem -LiteralPath $rollbackSnapshot.Paths.VersionsFolderPath -Directory | ForEach-Object Name)
    $lockedClusterStream = [IO.File]::Open($rollbackSnapshot.Paths.ClusterPath, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::Read)
    $rollbackFailureMessage = ''
    try {
        try {
            Publish-SmartM365OrchestratorConfiguration `
                -SharedDataFolderPath $temporaryRoot `
                -JobsDocument $rollbackJobs `
                -ClusterDocument $rollbackCluster `
                -ExpectedJobsHash $rollbackSnapshot.JobsHash `
                -ExpectedClusterHash $rollbackSnapshot.ClusterHash `
                -ChangeSummary 'Expected rollback test' `
                -AtomicWriteRetrySeconds 0 | Out-Null
        }
        catch {
            $rollbackFailureMessage = $_.Exception.Message
        }
    }
    finally {
        $lockedClusterStream.Dispose()
    }
    $afterRollbackFailure = Get-SmartM365OrchestratorConfigurationSnapshot -SharedDataFolderPath $temporaryRoot
    Assert-True -Condition ($rollbackFailureMessage -like "*failed at 'Orchestrator-Cluster.json'*rolled back*") -Message 'A second-file publication failure did not report successful rollback.'
    Assert-True -Condition ($afterRollbackFailure.JobsHash -eq $rollbackSnapshot.JobsHash -and $afterRollbackFailure.ClusterHash -eq $rollbackSnapshot.ClusterHash) -Message 'A second-file publication failure did not restore the exact prior shared configuration.'
    $failedVersionFolder = @(Get-ChildItem -LiteralPath $rollbackSnapshot.Paths.VersionsFolderPath -Directory | Where-Object Name -notin $versionsBeforeFailure | Sort-Object Name -Descending)[0]
    # The record follows the JSON transport policy (Publication-Failed.json.txt in JsonText mode).
    $failureRecordPath = if ($null -ne $failedVersionFolder) { Get-SmartM365JsonReadPath (Join-Path $failedVersionFolder.FullName 'Publication-Failed.json') -Optional } else { $null }
    Assert-True -Condition (-not [string]::IsNullOrWhiteSpace([string]$failureRecordPath)) -Message 'A failed publication did not leave an explicit failure record in its version folder.'
    $failureRecord = (Read-SmartM365JsonDocument -Path $failureRecordPath).Document
    Assert-True -Condition ([bool]$failureRecord.RollbackSucceeded -and [string]$failureRecord.FailedStage -eq 'Orchestrator-Cluster.json') -Message 'The failed-publication record does not describe the rollback outcome.'

    $jobRunsFolder = Join-Path -Path $temporaryRoot -ChildPath 'SERVER-A\JobRuns'
    New-Item -ItemType Directory -Path $jobRunsFolder -Force | Out-Null
    @(
        [pscustomobject]@{
            JobName = 'ExampleJob'
            ScheduledTime = (Get-Date).AddMinutes(-10).ToString('o')
            StartTime = (Get-Date).AddMinutes(-9).ToString('o')
            EndTime = (Get-Date).AddMinutes(-8).ToString('o')
            DurationSec = 60
            ExitCode = 0
            Status = 'Success'
            RetryCount = 0
            LogPath = 'C:\Logs\ExampleJob.log'
        }
    ) | Export-Csv -LiteralPath (Join-Path $jobRunsFolder ('Orchestrator_JobRuns_{0}.csv' -f (Get-Date).ToString('yyyyMMdd'))) -NoTypeInformation -Encoding utf8
    $history = @(Get-SmartM365OrchestratorHistory -SharedDataFolderPath $temporaryRoot -From (Get-Date).AddDays(-1) -To (Get-Date).AddDays(1))
    Assert-True -Condition ($history.Count -eq 1 -and $history[0].Server -eq 'SERVER-A') -Message 'All-server history aggregation failed.'

    foreach ($server in @('SERVER-A', 'SERVER-B')) {
        $serverFolder = Join-Path -Path $temporaryRoot -ChildPath $server
        New-Item -ItemType Directory -Path $serverFolder -Force | Out-Null
        [pscustomobject]@{ Timestamp = [datetime]::UtcNow.ToString('o'); Pid = 1234 } |
            ConvertTo-Json |
            Set-Content -LiteralPath (Join-Path $serverFolder 'Orchestrator-Heartbeat.json') -Encoding utf8
        [pscustomobject]@{ ReadyCapabilities = @('SharedRuntime', 'ExchangeOnPrem') } |
            ConvertTo-Json |
            Set-Content -LiteralPath (Join-Path $serverFolder 'Orchestrator-Capabilities.json') -Encoding utf8
    }
    $serverStatus = @(Get-SmartM365OrchestratorServerStatus -SharedDataFolderPath $temporaryRoot -ClusterDocument $snapshot.Cluster)
    $serverBStatus = $serverStatus | Where-Object Server -EQ 'SERVER-B'
    Assert-True -Condition ($serverStatus.Count -eq 2 -and @($serverStatus | Where-Object Online).Count -eq 2) -Message 'Current Timestamp heartbeats were not reported online.'
    Assert-True -Condition ($serverBStatus.Policy -eq 'ExchangeOnPrem') -Message 'The Exchange on-premises server policy was not rendered correctly.'

    $restoreSnapshot = Get-SmartM365OrchestratorConfigurationSnapshot -SharedDataFolderPath $temporaryRoot
    $restoreResult = Restore-SmartM365OrchestratorConfigurationVersion `
        -SharedDataFolderPath $temporaryRoot `
        -VersionFolderPath $publishResult.VersionFolderPath `
        -Snapshot Before `
        -ExpectedJobsHash $restoreSnapshot.JobsHash `
        -ExpectedClusterHash $restoreSnapshot.ClusterHash
    Assert-True -Condition (-not [string]::IsNullOrWhiteSpace($restoreResult.VersionId)) -Message 'Rollback did not publish a new auditable version.'
}
finally {
    if (Test-Path -LiteralPath $temporaryRoot) {
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ("[{0}] SmartM365 Orchestrator management tests passed." -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) -ForegroundColor Green
