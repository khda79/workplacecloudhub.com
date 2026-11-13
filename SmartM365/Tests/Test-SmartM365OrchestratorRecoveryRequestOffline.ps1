#requires -Version 7.0
$ErrorActionPreference='Stop'
$base=Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $base 'SmartM365.Orchestrator.PendingWork.psm1') -Force
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Maintenance.psm1')
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Pipeline.psm1')
Import-Module (Join-Path $base 'SmartM365.Orchestrator.Distributed.psm1')
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-RecoveryRequest-' + [guid]::NewGuid().ToString('N'))
$raceWorker=$null
function Write-Fixture($Path,$Value) {
    $null=New-Item -ItemType Directory -Path (Split-Path $Path -Parent) -Force
    [IO.File]::WriteAllText($Path,($Value | ConvertTo-Json -Depth 15),[Text.UTF8Encoding]::new($false))
}
try {
    $null=New-Item -ItemType Directory -Path (Join-Path $root 'Config') -Force
    $null=New-Item -ItemType Directory -Path (Join-Path $root 'Election/Concurrency'),(Join-Path $root 'Election/Claims') -Force
    Write-Fixture (Join-Path $root 'Config/Orchestrator-Cluster.json.txt') ([pscustomobject]@{ExpectedOrchestratorServers=@('SERVER-A');PeerHeartbeatStaleMinutes=5})
    $jobsPath=Join-Path $root 'Config/Orchestrator-Jobs.json.txt'
    Write-Fixture $jobsPath ([pscustomobject]@{Jobs=@([pscustomobject]@{Name='Example';Group='M365';Enabled=$true;AssignmentMode='Elected';DependsOn=@();RequiredCapabilities=@('EXO')})})
    $control=Initialize-SmartM365OrchestratorMaintenance $root
    $heartbeat=[pscustomobject]@{Tenant='test';ScriptVersion='1.5.90';Timestamp=[datetime]::UtcNow.ToString('o');Lifecycle='Running';RunningJobs=@();MaintenanceProtocol=1;MaintenanceCatchUpProtocol=1;MaintenanceResetProtocol=1;RecoveryGuardProtocol=2;PipelineCancellationProtocol=1;PipelineQueueProtocol=1;StatePersistenceHealthy=$true;MaintenanceHealthy=$true;MaintenanceEnabled=$false;MaintenanceRevision=$control.Revision}
    Write-Fixture (Join-Path $root 'SERVER-A/Orchestrator-Heartbeat.json.txt') $heartbeat
    Write-Fixture (Join-Path $root 'SERVER-A/Orchestrator-State.json.txt') ([pscustomobject]@{Jobs=[pscustomobject]@{}})
    $hash=(Get-FileHash -LiteralPath $jobsPath -Algorithm SHA256).Hash
    $guiAst=[Management.Automation.Language.Parser]::ParseFile((Join-Path $base 'SmartM365-Inventory-Orchestrator-GUI.ps1'),[ref]$null,[ref]$null)
    $workerFunction=$guiAst.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-GuiCancellationWork'},$true)
    . ([scriptblock]::Create($workerFunction.Extent.Text))
    $work=Get-GuiCancellationWork
    $progress=[hashtable]::Synchronized(@{Message='';LogError=''})
    $inputData=[pscustomobject]@{Kind='RecoveryWork';Phase='Validate';Root=$root;Tenant='test';All=$false;JobNames=@('Example');JobsPath='';ManifestHash='';PreviewSinceUtc=[datetime]::MinValue;LogPath='';ModulePaths=@(
        (Join-Path $base 'SmartM365.Orchestrator.Management.psm1'),(Join-Path $base 'SmartM365.Orchestrator.Pipeline.psm1'),
        (Join-Path $base 'SmartM365.Orchestrator.Maintenance.psm1'),(Join-Path $base 'SmartM365.Orchestrator.Insights.psm1'),(Join-Path $base 'SmartM365.Orchestrator.PendingWork.psm1'))}
    $runsFolder=Join-Path $root 'SERVER-A/JobRuns'; $null=New-Item -ItemType Directory -Path $runsFolder -Force
    [pscustomobject]@{JobName='Example';ScheduledTime=(Get-Date).AddMinutes(-5).ToString('yyyy-MM-dd HH:mm:ss');StartTime='';EndTime=(Get-Date).ToString('yyyy-MM-dd HH:mm:ss');DurationSec=0;ExitCode='';Status='BlockedDependencyTimeout';RetryCount=0;LogPath=''} | Export-Csv -LiteralPath (Join-Path $runsFolder ('Orchestrator_JobRuns_{0}.csv' -f (Get-Date -Format yyyyMMdd))) -NoTypeInformation
    $inputData.All=$true
    $allPreview=& $work $progress $inputData
    if ($allPreview.Phase -ne 'RecoveryValidate' -or 'Example' -notin @($allPreview.Preview.RequestedJobs)) { throw 'All-blocked preview omitted a terminal dependency block.' }
    $inputData.All=$false
    $preview=& $work $progress $inputData
    if ($preview.Phase -ne 'RecoveryValidate' -or -not $preview.Preview.Jobs[0].OwnerUnavailable) { throw 'GUI recovery preview did not report the missing owner.' }
    $inputData.Phase='Publish';$inputData.JobsPath=$preview.JobsPath;$inputData.ManifestHash=$preview.ManifestHash;$inputData.PreviewSinceUtc=$preview.PreviewSinceUtc
    $published=& $work $progress $inputData
    if ($published.Phase -ne 'RecoveryPublish') { throw 'GUI recovery worker did not return the published request.' }
    $request=$published.Result
    if (-not $request.Submitted -or -not (Test-SmartM365OrchestratorRecoveryReservation -SharedDataFolderPath $root -JobName Example -Tenant test)) { throw 'Recovery was not durably reserved without an elected owner.' }
    $differentSelection=[pscustomobject]@{SelectedJobs=@([pscustomobject]@{Name='Example'},[pscustomobject]@{Name='Other'})}
    $overlapNormal=$false
    try { $null=New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $root -Tenant test -Pipeline Jobs -Selection $differentSelection -ManifestHash $hash } catch { $overlapNormal=$_.Exception.Message -match 'overlaps' }
    if (-not $overlapNormal) { throw 'An ordinary request overlapping recovery must be refused.' }
    $overlap=$false
    try { $null=New-SmartM365OrchestratorRecoveryRequest -SharedDataFolderPath $root -JobsPath $jobsPath -Tenant test -JobName Example -ExpectedManifestHash $hash -PreviewSinceUtc ([datetime]::UtcNow) } catch { $overlap=$true }
    if (-not $overlap) { throw 'Overlapping recovery request was not refused.' }
    $null=Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $root -BatchId $request.BatchId -JobName Example -Status Success
    # A dependent must not be queued against an already-running prerequisite:
    # cross-server dependency checks cannot safely infer its completion yet.
    Write-Fixture $jobsPath ([pscustomobject]@{Jobs=@(
        [pscustomobject]@{Name='Example';Group='M365';Enabled=$true;AssignmentMode='Elected';DependsOn=@();RequiredCapabilities=@('EXO')},
        [pscustomobject]@{Name='Child';Group='M365';Enabled=$true;AssignmentMode='Elected';DependsOn=@('Example');RequiredCapabilities=@('EXO')})})
    $dependentHash=(Get-FileHash -LiteralPath $jobsPath -Algorithm SHA256).Hash
    $heartbeat.RunningJobs=@([pscustomobject]@{Name='Example';Pid=999})
    Write-Fixture (Join-Path $root 'SERVER-A/Orchestrator-Heartbeat.json.txt') $heartbeat
    $dependencyRefused=$false
    try { $null=New-SmartM365OrchestratorRecoveryRequest -SharedDataFolderPath $root -JobsPath $jobsPath -Tenant test -JobName Child -ExpectedManifestHash $dependentHash -PreviewSinceUtc ([datetime]::UtcNow) } catch { $dependencyRefused=$_.Exception.Message -match 'depends on work already' }
    if (-not $dependencyRefused) { throw 'Recovery must refuse a dependent while its prerequisite is already running.' }
    if (Test-SmartM365OrchestratorRecoveryReservation -SharedDataFolderPath $root -JobName Child -Tenant test) { throw 'A refused dependent must not be reserved.' }
    $heartbeat.RunningJobs=@(); Write-Fixture (Join-Path $root 'SERVER-A/Orchestrator-Heartbeat.json.txt') $heartbeat
    Write-Fixture $jobsPath ([pscustomobject]@{Jobs=@([pscustomobject]@{Name='Example';Group='M365';Enabled=$true;AssignmentMode='Elected';DependsOn=@();RequiredCapabilities=@('EXO')})})
    $hash=(Get-FileHash -LiteralPath $jobsPath -Algorithm SHA256).Hash
    Write-Fixture (Join-Path $root 'SERVER-A/Orchestrator-State.json.txt') ([pscustomobject]@{Jobs=[pscustomobject]@{Example=[pscustomobject]@{Running=$null;PendingRetry=$null;LastStatus='Success';LastRunEnd=[datetime]::UtcNow.ToString('o')}}})
    $finished=New-SmartM365OrchestratorRecoveryRequest -SharedDataFolderPath $root -JobsPath $jobsPath -Tenant test -JobName Example -ExpectedManifestHash $hash -PreviewSinceUtc ([datetime]::UtcNow.AddMinutes(-1))
    if ($finished.Submitted -or 'Example' -notin $finished.AlreadyActive) { throw 'A job that succeeded after preview must not be submitted again.' }
    Write-Fixture (Join-Path $root 'SERVER-A/Orchestrator-State.json.txt') ([pscustomobject]@{Jobs=[pscustomobject]@{}})
    $heartbeat.RunningJobs=@([pscustomobject]@{Name='Example';Pid=999})
    Write-Fixture (Join-Path $root 'SERVER-A/Orchestrator-Heartbeat.json.txt') $heartbeat
    $running=New-SmartM365OrchestratorRecoveryRequest -SharedDataFolderPath $root -JobsPath $jobsPath -Tenant test -JobName Example -ExpectedManifestHash $hash -PreviewSinceUtc ([datetime]::UtcNow)
    if ($running.Submitted -or 'Example' -notin $running.AlreadyActive) { throw 'An already running job must finish rather than be submitted again.' }
    $heartbeat.RunningJobs=@(); Write-Fixture (Join-Path $root 'SERVER-A/Orchestrator-Heartbeat.json.txt') $heartbeat
    $leasePath=Join-Path $root 'Election/Concurrency/Example.json.txt'
    Write-Fixture $leasePath ([pscustomobject]@{JobName='Example';LeaseId='fixture';OwnerServer='SERVER-A';ConcurrencyKey='Example'})
    $leased=New-SmartM365OrchestratorRecoveryRequest -SharedDataFolderPath $root -JobsPath $jobsPath -Tenant test -JobName Example -ExpectedManifestHash $hash -PreviewSinceUtc ([datetime]::UtcNow)
    if ($leased.Submitted -or 'Example' -notin $leased.AlreadyActive) { throw 'A job with a shared running lease must not be submitted again.' }
    Remove-Item -LiteralPath $leasePath -Force
    $claimPath=Join-Path $root ('Election/Claims/Example/{0}.json.txt' -f ([datetime]::UtcNow.ToString('yyyyMMddTHHmmssfffZ')))
    Write-Fixture $claimPath ([pscustomobject]@{JobName='Example';Status='RetryScheduled';SafeUntilUtc=[datetime]::UtcNow.AddHours(1).ToString('o')})
    $retry=New-SmartM365OrchestratorRecoveryRequest -SharedDataFolderPath $root -JobsPath $jobsPath -Tenant test -JobName Example -ExpectedManifestHash $hash -PreviewSinceUtc ([datetime]::UtcNow)
    if ($retry.Submitted -or 'Example' -notin $retry.AlreadyActive) { throw 'A shared retry claim must not be duplicated.' }
    Remove-Item -LiteralPath $claimPath -Force
    # A scheduled launch wins the gate just before recovery: its lease must be
    # visible to the recovery submitter after the gate is released.
    $gate=Enter-SmartM365OrchestratorMaintenanceGate -SharedDataFolderPath $root
    try {
        $signal=Join-Path $root 'race-started.txt'
        $modulePath=(Resolve-Path (Join-Path $base 'SmartM365.Orchestrator.PendingWork.psm1')).Path
        $raceWorker=Start-Job -ArgumentList $modulePath,$root,$jobsPath,$hash,$signal -ScriptBlock {
            param($modulePath,$root,$jobsPath,$hash,$signal)
            $ErrorActionPreference='Stop'; Import-Module $modulePath -Force
            [IO.File]::WriteAllText($signal,'ready')
            New-SmartM365OrchestratorRecoveryRequest -SharedDataFolderPath $root -JobsPath $jobsPath -Tenant test -JobName Example -ExpectedManifestHash $hash -PreviewSinceUtc ([datetime]::UtcNow)
        }
        $deadline=[datetime]::UtcNow.AddSeconds(10)
        while (-not (Test-Path -LiteralPath $signal) -and [datetime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 100 }
        if (-not (Test-Path -LiteralPath $signal)) { throw 'Race worker did not start.' }
        $lease=Enter-SmartM365OrchestratorConcurrencyLease -LeasesRootPath (Join-Path $root 'Election/Concurrency') -ConcurrencyKey Example -JobName Example -Occurrence (Get-Date) -OwnerServer SERVER-A
        if (-not $lease.Acquired) { throw 'Synthetic scheduled launch could not reserve its lease.' }
    } finally { $gate.Dispose() }
    if (-not (Wait-Job $raceWorker -Timeout 10)) { throw 'Race worker did not finish.' }
    $race=Receive-Job $raceWorker -ErrorAction Stop
    if ($race.Submitted -or 'Example' -notin $race.AlreadyActive) { throw 'Recovery duplicated a scheduled launch that won the shared gate.' }
    Remove-Job $raceWorker -Force; $raceWorker=$null
    $heartbeat.RecoveryGuardProtocol=0; Write-Fixture (Join-Path $root 'SERVER-A/Orchestrator-Heartbeat.json.txt') $heartbeat
    $unsupported=$false
    try { $null=New-SmartM365OrchestratorRecoveryRequest -SharedDataFolderPath $root -JobsPath $jobsPath -Tenant test -JobName Example -ExpectedManifestHash $hash -PreviewSinceUtc ([datetime]::UtcNow) } catch { $unsupported=$true }
    if (-not $unsupported) { throw 'An incompatible server must not receive a recovery request.' }
    'PASS: recovery reservation, overlap and running-job protection'
}
finally { if ($raceWorker) { Stop-Job $raceWorker -ErrorAction SilentlyContinue; Remove-Job $raceWorker -Force -ErrorAction SilentlyContinue }; if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force } }
