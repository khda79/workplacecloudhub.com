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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCfXB4JlJaGNCCp
# CFlnfteMbcdVmj0NJsq+aO+YOWOMkaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILuOtBgkdpfTlqcMkmkVK0iZPZFCWDY0+4InQOsgPnVUMA0GCSqG
# SIb3DQEBAQUABIIBgALOqITLL08+mWpeg5VH5RRCfOw4YHlvYFaOod3+c/UTNWdP
# Z7L1j6PmmhFTuShxvGfhgKGrqxH7sJu9NPdMbwMN1qNJAvqleezkHLKdwVdb09RY
# IBW/GrVrVpAybuZwkIn9NlCYzO/PfHQuJKDFhATl0ZrjN2GrJNmp8hynP95y83h9
# racYFYSXM8ocHAC8cFJYa6GXYGpJZFWLfs68D6UQ1JPwRDCTCqjvZHUbqpKU3p3e
# KQ8TblnMH4TdDn2aDFyivgAx7ov4FZWVuqm04WNYqoCeRWIifeh2Gep5MNtoG+pj
# AEYHWDOaP8VcHjUFPFXn4mK5zpZNaj1sWJ27ShYf8/SUh30Y7lGjVpGacV1J52wq
# 3OAlSL2fEXONFkwB6I2lZ/rNxzGVsNOdLNQ5q1/FVJct8cmywH375UuP+qBKmlvc
# 4/wV2Jc0TtyX0lrv0vMR1ZF7tJRlc1yaW0uLru10CT6cRqdxu/JqlIKGwPKC1Uee
# pT3L90iUsxZficT7qqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTMxMzE2
# MjFaMC8GCSqGSIb3DQEJBDEiBCBfY2FT4+lO4zNMR1hBiq2opQUJtS6rhlgxD/o+
# fn55fDANBgkqhkiG9w0BAQEFAASCAgCMS9W+whOl04b8JuClzlMyX3sJ13Ircj4q
# 9noFwx15bXgvRvAjeM1gqUZ/Qg4w7xkz892wUDvVGOcvYMfzs9oNDsVfNfJbIwg0
# pBJ3iuJTvcRk88tWVtJe20ong3+pzpi77WMvx6J+BrZWiUThrOjU1zbnedJEjwBz
# g/D1G+YbE+OqBxBPyTgDamPC8sr53JCoFrs8EJqk4QC4x4DqAfNijoxOv1FqZtFt
# UhGeVI8P0bdkSvkQXM6BNxc549MCnm2d2i3LQNE+Ykwuon26r79DGkRkX6I4pwmb
# Xu/uR7iPnXMpNMf0gOCiZCtpIZCMkm+sk1tJHr0AF3Pdx2kuZMQGMnpF569/aVRY
# DgrhE8ViQL69BxptdtddTa6NzBzYvWoCaW5YWF00ez4raNyjfyEpzE97FVs2TGtE
# KmUdAy7e2SeYHTR/ITbv0RnZ514K2Qg1UWInrFSB8rX6oVRs1P85TCV3hca3hdjK
# U/r85cP8pusYBuXlo7cLsj79F/+kTpHmxkveccHACO+JnCY0WmlMArl27SFfSKjP
# ZnGAON9kbs3zvxQziqgL/agcszmxAMf0aMncMmT0x+ddsZhw1SQ+nNo1X9FKbUzV
# 4L+dM69bTv5OOu5IZXnbcqQRXK8vMNjxSHC4PqTPxB01GrpCj5ycSPIXtRTjak9c
# oFwYLtVKog==
# SIG # End signature block
