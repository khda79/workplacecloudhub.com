# Coordinated pending-work cancellation. Running jobs are not stopped.
Set-StrictMode -Version 2.0
Import-Module (Join-Path $PSScriptRoot '../../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -MinimumVersion '1.0.0' -Global -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.Maintenance.psm1') -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.Pipeline.psm1') -ErrorAction Stop
Import-Module (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.Distributed.psm1') -ErrorAction Stop

function ConvertTo-PendingWorkUtc {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetimeoffset]) { return $Value.UtcDateTime }
    if ($Value -is [datetime]) { return $Value.ToUniversalTime() }
    $parsed=[datetimeoffset]::MinValue
    if ([datetimeoffset]::TryParse([string]$Value,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::RoundtripKind,[ref]$parsed)) { return $parsed.UtcDateTime }
    return $null
}

function Assert-SmartM365PendingWorkResetReady {
    param([Parameter(Mandatory)][string]$SharedDataFolderPath, [Parameter(Mandatory)][string]$Tenant, [Parameter(Mandatory)]$Control, [switch]$RequireRecoveryGuard)
    $cluster = (Read-SmartM365JsonDocument -Path (Join-Path $SharedDataFolderPath 'Config/Orchestrator-Cluster.json')).Document
    $servers = @(Get-SmartM365OrchestratorMaintenanceReadiness -SharedDataFolderPath $SharedDataFolderPath -ClusterDocument $cluster -State $Control)
    if (-not $servers.Count -or @($servers | Where-Object Status -ne 'Applied').Count) { throw 'All expected servers must be online and acknowledge the current maintenance revision.' }
    foreach ($server in $servers) {
        $heartbeat = (Read-SmartM365JsonDocument -Path (Join-Path $SharedDataFolderPath "$($server.Server)/Orchestrator-Heartbeat.json")).Document
        if (-not $heartbeat.PSObject.Properties['Tenant'] -or [string]$heartbeat.Tenant -ne $Tenant -or
            -not $heartbeat.PSObject.Properties['MaintenanceResetProtocol'] -or [int]$heartbeat.MaintenanceResetProtocol -ne 1 -or
            -not $heartbeat.PSObject.Properties['PipelineCancellationProtocol'] -or [int]$heartbeat.PipelineCancellationProtocol -ne 1 -or
            -not $heartbeat.PSObject.Properties['StatePersistenceHealthy'] -or $heartbeat.StatePersistenceHealthy -ne $true) {
            throw "Server $($server.Server) has an incompatible tenant, protocol or state-persistence status."
        }
        if ($RequireRecoveryGuard -and (-not $heartbeat.PSObject.Properties['RecoveryGuardProtocol'] -or [int]$heartbeat.RecoveryGuardProtocol -ne 2)) { throw "Server $($server.Server) does not support recovery reservations and shared schedule receipts." }
    }
    return $servers
}

function New-SmartM365OrchestratorRecoveryRequest {
    <# The same gate guards scheduled process creation and recovery request publication. #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SharedDataFolderPath, [Parameter(Mandatory)][string]$JobsPath,
        [Parameter(Mandatory)][string]$Tenant, [Parameter(Mandatory)][string[]]$JobName,
        [Parameter(Mandatory)][string]$ExpectedManifestHash, [Parameter(Mandatory)][datetime]$PreviewSinceUtc)
    $previewUtc = $PreviewSinceUtc.ToUniversalTime()
    if ($previewUtc -gt [datetime]::UtcNow.AddMinutes(1) -or $previewUtc -lt [datetime]::UtcNow.AddMinutes(-15)) { throw 'Recovery preview has expired or has an invalid clock; refresh it.' }
    $gate = Enter-SmartM365OrchestratorMaintenanceGate -SharedDataFolderPath $SharedDataFolderPath
    try {
        $control = Get-SmartM365OrchestratorMaintenanceState -SharedDataFolderPath $SharedDataFolderPath
        if ($control.Enabled) { throw 'Recovery cannot be submitted while maintenance is active.' }
        $servers = @(Assert-SmartM365PendingWorkResetReady -SharedDataFolderPath $SharedDataFolderPath -Tenant $Tenant -Control $control -RequireRecoveryGuard)
        $publishedJobs = Get-SmartM365JsonReadPath -Path (Join-Path $SharedDataFolderPath 'Config/Orchestrator-Jobs.json') -Optional
        if (-not $publishedJobs -or -not [string]::Equals([IO.Path]::GetFullPath($JobsPath), [IO.Path]::GetFullPath($publishedJobs), [StringComparison]::OrdinalIgnoreCase)) { throw 'Recovery requires the published jobs configuration from this orchestrator root.' }
        $hash = (Get-FileHash -LiteralPath $JobsPath -Algorithm SHA256 -ErrorAction Stop).Hash
        if ($hash -ine $ExpectedManifestHash) { throw 'Published jobs configuration changed; refresh the recovery preview.' }
        $document = (Read-SmartM365JsonDocument -Path $JobsPath).Document
        $selection = Get-SmartM365OrchestratorPipelineSelection -JobsDocument $document -JobName $JobName -IncludeDependencies
        $leasesRoot = Join-Path $SharedDataFolderPath 'Election/Concurrency'
        $claimsRoot = Join-Path $SharedDataFolderPath 'Election/Claims'
        foreach ($folder in @($leasesRoot,$claimsRoot)) {
            $null = Get-Item -LiteralPath $folder -ErrorAction Stop
            Get-ChildItem -LiteralPath $folder -Force -ErrorAction Stop | Select-Object -First 1 | Out-Null
        }
        $busy = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($server in $servers) {
            $folder = Join-Path $SharedDataFolderPath ([string]$server.Server)
            $heartbeat = (Read-SmartM365JsonDocument -Path (Join-Path $folder 'Orchestrator-Heartbeat.json')).Document
            if (-not $heartbeat.PSObject.Properties['RunningJobs']) { throw "Server $($server.Server) has no running-job snapshot." }
            foreach ($run in @($heartbeat.RunningJobs)) { if ($run.Name) { [void]$busy.Add([string]$run.Name) } }
            $state = (Read-SmartM365JsonDocument -Path (Join-Path $folder 'Orchestrator-State.json')).Document
            if (-not $state.PSObject.Properties['Jobs'] -or $null -eq $state.Jobs) { throw "Server $($server.Server) has no readable job state." }
            foreach ($job in @($selection.SelectedJobs)) {
                $entry = $state.Jobs.PSObject.Properties[[string]$job.Name]
                if ($null -eq $entry -or $null -eq $entry.Value) { continue }
                $jobState = $entry.Value
                if (($jobState.PSObject.Properties['Running'] -and $null -ne $jobState.Running) -or
                    ($jobState.PSObject.Properties['PendingRetry'] -and $null -ne $jobState.PendingRetry)) { [void]$busy.Add([string]$job.Name) }
                if ($jobState.PSObject.Properties['LastStatus'] -and [string]$jobState.LastStatus -in @('Success','CompletedWithWarnings') -and $jobState.PSObject.Properties['LastRunEnd'] -and $jobState.LastRunEnd) {
                    $finished = ConvertTo-PendingWorkUtc $jobState.LastRunEnd
                    if ($null -eq $finished) { throw "Invalid last completion time for $($job.Name) on $($server.Server)." }
                    if ($finished -ge $previewUtc) { [void]$busy.Add([string]$job.Name) }
                }
            }
        }
        foreach ($job in @($selection.SelectedJobs)) {
            $definition = @($document.Jobs | Where-Object { [string]$_.Name -eq [string]$job.Name })[0]
            $key = if ($definition.PSObject.Properties['ConcurrencyKey'] -and [string]$definition.ConcurrencyKey) { [string]$definition.ConcurrencyKey } else { [string]$job.Name }
            $lease = Get-SmartM365OrchestratorConcurrencyLease -LeasesRootPath $leasesRoot -ConcurrencyKey $key
            if ($null -ne $lease -and [string]$lease.Lease.JobName -eq [string]$job.Name) { [void]$busy.Add([string]$job.Name) }
            if ([string]$job.AssignmentMode -eq 'Elected') {
                if ([string]$job.Name -notmatch '^[A-Za-z0-9._-]+$') { throw 'Invalid job name in recovery selection.' }
                $claimFolder = Join-Path $claimsRoot ([string]$job.Name)
                if (Test-Path -LiteralPath $claimFolder -PathType Container -ErrorAction Stop) {
                    foreach ($file in @(Get-ChildItem -LiteralPath $claimFolder -File -ErrorAction Stop | Where-Object { $_.Name -match '^\d{8}T\d{9}Z\.json(?:\.txt)?$' })) {
                        $claim = (Read-SmartM365JsonDocument -Path $file.FullName).Document
                        if (-not $claim.PSObject.Properties['Status'] -or -not $claim.PSObject.Properties['JobName'] -or [string]$claim.JobName -ne [string]$job.Name) { throw "Invalid occurrence claim for $($job.Name)." }
                        if ([string]$claim.Status -eq 'RetryScheduled') { [void]$busy.Add([string]$job.Name); break }
                        if ([string]$claim.Status -in @('Claimed','Running')) {
                            $until = if ($claim.PSObject.Properties['SafeUntilUtc']) { ConvertTo-PendingWorkUtc $claim.SafeUntilUtc } else { $null }
                            if ($null -eq $until) { throw "Invalid active occurrence claim for $($job.Name)." }
                            if ($until -gt [datetime]::UtcNow) { [void]$busy.Add([string]$job.Name); break }
                        }
                    }
                }
            }
        }
        $remaining = @($selection.SelectedJobs | Where-Object { -not $busy.Contains([string]$_.Name) })
        $skipped = @($selection.SelectedJobs | Where-Object { $busy.Contains([string]$_.Name) } | ForEach-Object Name)
        if (-not $remaining.Count) { return [pscustomobject]@{ Submitted=$false; BatchId=''; AlreadyActive=$skipped; TotalCount=0; QueuePosition=0 } }
        $dependsOnBusy = @($remaining | Where-Object { @($_.DependsOn | Where-Object { $busy.Contains([string]$_) }).Count } | ForEach-Object Name)
        if ($dependsOnBusy.Count) { throw "Recovery depends on work already running, retrying or completed since the preview. Let it finish and refresh before submitting: $($skipped -join ', ')." }
        $remainingNames = @($remaining | ForEach-Object Name)
        foreach ($job in $remaining) { $job.ExternalDependencies = @($job.DependsOn | Where-Object { $_ -notin $remainingNames }) }
        $selection.SelectedJobs = $remaining
        $request = New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $SharedDataFolderPath -Tenant $Tenant -Pipeline Jobs -Selection $selection -ManifestHash $hash -RequestedBy ([Environment]::UserName) -RequestedFrom ([Environment]::MachineName) -Recovery
        return [pscustomobject]@{ Submitted=$true; BatchId=[string]$request.BatchId; AlreadyActive=$skipped; TotalCount=[int]$request.TotalCount; QueuePosition=[int]$request.QueuePosition }
    }
    finally { $gate.Dispose() }
}

function Invoke-SmartM365CancelPendingWork {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$SharedDataFolderPath,
        [Parameter(Mandatory)][string]$Tenant,
        [Parameter(Mandatory)][long]$ExpectedRevision,
        [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$Reason,
        [ValidateRange(1,600)][int]$AcknowledgementTimeoutSeconds = 180,
        [scriptblock]$OnProgress = $null
    )
    if ([string]::IsNullOrWhiteSpace($Reason)) { throw 'A reason is required.' }
    $control = Get-SmartM365OrchestratorMaintenanceState -SharedDataFolderPath $SharedDataFolderPath
    if ($control.Enabled -or $control.Revision -ne $ExpectedRevision) { throw 'Maintenance is active or changed; refresh the preview.' }
    $null = Assert-SmartM365PendingWorkResetReady -SharedDataFolderPath $SharedDataFolderPath -Tenant $Tenant -Control $control
    $enabled = $null
    try {
        $enabled = Set-SmartM365OrchestratorMaintenance -SharedDataFolderPath $SharedDataFolderPath -Enabled $true -ExpectedRevision $ExpectedRevision -Reason $Reason -PendingWorkReset
        $deadline = [datetime]::UtcNow.AddSeconds($AcknowledgementTimeoutSeconds)
        while ($true) {
            if ($OnProgress) { & $OnProgress 'Waiting for servers to acknowledge the pause.' | Out-Null }
            try { $null = Assert-SmartM365PendingWorkResetReady -SharedDataFolderPath $SharedDataFolderPath -Tenant $Tenant -Control $enabled; break }
            catch {
                if ([datetime]::UtcNow -ge $deadline) { throw "Servers did not acknowledge the pause before the timeout: $($_.Exception.Message)" }
                Start-Sleep -Seconds 2
            }
        }
        $cancelled = [Collections.Generic.List[string]]::new()
        $runs = @(Get-SmartM365OrchestratorActivePipelineRuns -SharedDataFolderPath $SharedDataFolderPath)
        foreach ($run in $runs) {
            if ([string]$run.Tenant -ne $Tenant) { throw 'An active pipeline request belongs to another tenant.' }
            if ($run.CancellationRequested) { continue }
            if ($OnProgress) { & $OnProgress ("Cancelling request {0}." -f $run.BatchId) | Out-Null }
            $gate = Enter-SmartM365OrchestratorMaintenanceGate -SharedDataFolderPath $SharedDataFolderPath
            try {
                $liveControl = Get-SmartM365OrchestratorMaintenanceState -SharedDataFolderPath $SharedDataFolderPath
                if (-not $liveControl.Enabled -or $liveControl.Revision -ne $enabled.Revision) { throw 'Maintenance changed while cancellation was in progress.' }
                $null = Stop-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $SharedDataFolderPath -BatchId ([string]$run.BatchId) -Tenant $Tenant -Reason $Reason
                $cancelled.Add([string]$run.BatchId)
            }
            catch {
                $current = Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $SharedDataFolderPath -BatchId ([string]$run.BatchId)
                if (-not $current.IsTerminal -and -not $current.CancellationRequested) { throw }
            }
            finally { $gate.Dispose() }
        }
        if ($OnProgress) { & $OnProgress 'Discarding pending scheduled work and resuming future schedules.' | Out-Null }
        $resumed = Set-SmartM365OrchestratorMaintenance -SharedDataFolderPath $SharedDataFolderPath -Enabled $false -ExpectedRevision $enabled.Revision -Reason $Reason -DiscardPendingCatchUp
        return [pscustomobject]@{ Revision=$resumed.Revision; ResumeAfterUtc=$resumed.ResumeAfterUtc; CancelledRequests=@($cancelled); RunningJobsStopped=$false }
    }
    catch {
        if ($null -ne $enabled) { throw "Pending-work cancellation did not complete. Maintenance revision $($enabled.Revision) was enabled; refresh its actual state before resuming. Running jobs were not stopped. $($_.Exception.Message)" }
        throw
    }
}
Export-ModuleMember -Function Assert-SmartM365PendingWorkResetReady, Invoke-SmartM365CancelPendingWork, New-SmartM365OrchestratorRecoveryRequest
