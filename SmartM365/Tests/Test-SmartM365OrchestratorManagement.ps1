#Requires -Version 7.0
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$smartM365Root = Split-Path -Path $PSScriptRoot -Parent
$modulePath = Join-Path -Path $smartM365Root -ChildPath 'SmartInventory\Orchestrator\SmartM365.Orchestrator.Management.psm1'
$jobsTemplatePath = Join-Path -Path $smartM365Root -ChildPath 'SmartInventory\Orchestrator\Orchestrator-Jobs.json.template'
Import-Module -Name $modulePath -Force -ErrorAction Stop

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
    $publishResult = Publish-SmartM365OrchestratorConfiguration `
        -SharedDataFolderPath $temporaryRoot `
        -JobsDocument $publishedJobs `
        -ClusterDocument $snapshot.Cluster `
        -ExpectedJobsHash $snapshot.JobsHash `
        -ExpectedClusterHash $snapshot.ClusterHash `
        -ChangeSummary 'Management test'
    Assert-True -Condition (Test-Path -LiteralPath $publishResult.VersionFolderPath) -Message 'No configuration version was created.'
    Assert-True -Condition (Test-Path -LiteralPath $snapshot.Paths.AuditPath) -Message 'No configuration audit CSV was created.'

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
