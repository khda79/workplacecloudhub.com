<#
.SYNOPSIS
Runs offline management and central-manifest migration tests for the SmartM365 orchestrator.
.VERSION
1.1.0
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
    Assert-True -Condition ($null -ne $failedVersionFolder -and (Test-Path -LiteralPath (Join-Path $failedVersionFolder.FullName 'Publication-Failed.json'))) -Message 'A failed publication did not leave an explicit failure record in its version folder.'
    $failureRecord = Get-Content -LiteralPath (Join-Path $failedVersionFolder.FullName 'Publication-Failed.json') -Raw | ConvertFrom-Json
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBAqTMYm5ugkoqd
# sidXe4M4CIZleH12OGuAbUhU+PN0b6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEII+6GhQ1hYIV0+KRJBHlRejFQtrjIxxPfZCNsW9wLGiaMA0GCSqG
# SIb3DQEBAQUABIIBgHAj82Y9ipCz/P72PoAmxlDf0v4LMBO44mfZEU/FuQK8DIO+
# r3VTS8toGQ4z415mBif25LPyGBVzpARWhiCybCE7tIfAlYP55JfU9nlgBUXXB4dK
# 88H6qZFXQ1x9qptBTGXjAKjCiSgf9SgeGp/CngynXQz58l45wMp2s5bGTbOEDNqo
# wBoKXAD50zCE1FwFQfUaR6wTs7J3Nzs9WtMiX2j+FaPI7t9mgk4LFEdz/QpoY/Ay
# A50XVQWAfv965CMSBrQ9P3WucjBxvCwRi12TWYADQco6dEE4UvZCmbZAU0iK+kb+
# dmBzmF9lDhL8AZlVPPbwr2Q8v0/eMml9Z53pH0Ej0DVLZIiuQWVM20FrrcMvSHX2
# fOOYPiR3zoxg068olPggEGM9qVzPEoaV/aRtGJKtWgmZETo66cMqweEXaBxZME94
# SW3PGqtcLKH/9Z2a5l1U4u8qF9uKBwUxqqvgMm2wP+WmLq0AyDzNHRwNB/dojI9+
# Mkud5iB1IajZa4K3MqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjMwNzU3
# NTZaMC8GCSqGSIb3DQEJBDEiBCB2zXBRbSGI0M6gWi9wcJwZJE3ug7/8NLeN2ZD+
# iBqDJDANBgkqhkiG9w0BAQEFAASCAgBSelDa5LtbIArqxuP3CTeeflaIvOT7lGuG
# 6V6S/xXqO7lONNtw8s4irUXXLtPxB2D1lH42a4CvU7+QvWha3SRgq2i/v04AjvDB
# O5oV9cITjM7hmkbkZ9emBw7m9BQBKVBkNwu1qvdDXRGnKbm784enKOneVVBo62Ty
# FVL0mnnwA+az8xkKZ53+gpeC+P0U+K7T++l/rI5ANqHShK7ltKcTA0nxu6LIvHug
# neXqSXwaux0JY+0aZbTyPgdNrfIZJA5cV+nn1sRCpqQT4oQBMhC76Lv4EaOTIbNQ
# 244zib9YYDOgkEbRpAWZ6fj4UZ7ysY+eI9lqxEPAL6tYdSQs5543cyKTFSVgHw6t
# tErxNsdZWriN6W3212CIUcu/6x/k09QLbDYYwHlypTmh+oqioqSyT4IlTBKS5XUU
# iRrth20iZD76bSmi8IEOWVabCzqnNm6i7Hdxy/caICPlC0CeUMPS5BiCqqfGvED0
# ltmQlxvR1khc1fD42m1fI3dirahMb8RSH2N5+2eBq/e3VuVXO1VaBOKMzkNem8MR
# e19N5bIYgjxY1tFRB173tkSyfyq8rhne7QMMZLo44hjtiWyvs5QXFyHGAA1g3Q8x
# o3PsUmHxNDXV9pxxOrYfc3ebRFG8yFPyJk61RxoQ0JcK9qKafWWMmYOiDBG+A9lG
# pmVeFXi1vw==
# SIG # End signature block
