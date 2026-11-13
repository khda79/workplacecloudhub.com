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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD1+7pYyK7/L9SY
# AYiKXjoFoT0Tjp4jCl9Yz//kUnHn16CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGqKZWP0UaQyZYL9b7Wo/ndbwjPGeoO2hROIPMLb29gzMA0GCSqG
# SIb3DQEBAQUABIIBgFMipDkdWzUSfCCAmgoLj6RkA1z8Wb9o5Wq+I+1mmKtuDqHl
# jQokYm24MkLQZEYMxvGfmRHHm5TDmQDS//VW41Yds6zqU4CHf+OvWiVfXLPIl4Y6
# 30ju8gtMan40UNzY4J62HL2n+Sl44sPLl4Sq68JiM+aYB7DDQaHln2MTY00rr/gj
# 5z5ggWLBk31uW0L0IlplRmQtLxPaKhjYj0J/cPzU0sf4W/cilSRNf9+GB/p9LptC
# 3tb5WbIo4J16SPH3CAxLHSAZ2F5i/BkNHLXkOLD5ZH7tn7yKi6DpKXnQrxKDOT10
# QfCdlpvNSW+XLWlDwMqx575J+KBJa6Ouw6RBLMU+b6+bi+aOpIB9u7Un5RvC5BPU
# jW9Zs6NfuWiOx1qyt0EqQ3ArJnpBoDug3lhTHVuw0LZjajjt1X9VZNsCdD7c4a1c
# 1twsN626ZoAhhvZm0QTKqQeLg/TI/ZIuGcFFscQ7Zprf8JU2NoOMmAdAS3PrxjEm
# K3hoYwqlyMjWYgcdY6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTMxMzE2
# MjRaMC8GCSqGSIb3DQEJBDEiBCATe9JCDv2Uwqk6/BMHH6syEZQogSFpP6/foWKz
# 3+cGXDANBgkqhkiG9w0BAQEFAASCAgA8MFyi8SkiMMszaDxYINdACLcfxnYxwjcl
# 0c9oaIBZHZIVtAES42ulz5+u65S/h+SLl/UhRnj6O0IkNZYcI4m+YFc2SPpqu0wQ
# gCElNYkP1A0BqyfFSbXrb22PYvKPCse49I9+vcZePwsKVECBFMGx44dKtenUKn/Z
# smj0E6xT5qVkpYRJ3zcDhgndMVkLmD6aGalq+8S3DFsYRip5INP17ghlUcqnbe7+
# JqIF4HVFrI3tE2BraxAPFKO3JhBK9GElzJ2xqEQeppwj3j+bAPnhsQyxp+RosD6m
# RxmTJvEILaLjzx7f90i9JZHzXuXMDoQ+l97mME1/qKBW2Wqpdtx4E0697KB6V3gg
# io3RU9TcV8n/TVLVpbTmpq55bOnOnlLyUTuG2uW1orJtee2uQUulNJercDukupOA
# NDcQOeC9pAZPFRX+Jft6An4/b8qrFjcKK9+50FC7vm+icrcFntSE8RC9UL7rFJsp
# e4wp9DSyemgDZ+7j8tLWg3bKsYTfQFcuvzMT91Z4wjdMNuQLGQQ4SGb40h4iIgo2
# Qv7Nc0GEbQ3wiXbCqAoTdZXRlBRm2urRjmNCUUlYYGe28KMzZDFXLFrwQDyAKiJH
# ElXy9zicDHe6EQJU5h204JFotPw4805HVnEvVaGVOBWHRmfHUqsjiXG7pGqugTAx
# TcgZRq+tnA==
# SIG # End signature block
