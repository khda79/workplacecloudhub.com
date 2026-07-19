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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBHp7DRZc8ly8p0
# 1rBgzPWo2XHUzcmSEPMlhI1dF4MjCqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIFPYkxhsDLzg4Mz8/JtAZXIyEt+aAguBoD50OW+z2mwnMA0GCSqG
# SIb3DQEBAQUABIIBgGXAL04BDqUz9J+5xoWfFjHl2V0VcZVGfXg5eKS57xyWWAJI
# y2g+DiFZzYR/cQF/kK4QYviSbTForfTVPnJ8uIbaTdUH46M7FjXTL2jq3RHoFgC/
# 5EgDx8Pjjns68jxsoBbBWSn15heXpa1Pek52kAIh9YKQyWWCx+rxMErMRIBMLjkz
# hrp9gYRRaxjo7fZ2NA1qj197YzRD7Bf+dP8WQ5syQlOZ8bsc9IBbUm/4WGPe1N7/
# NboCGuuBHk0uE6vO0J41oSWc3slawuYX8crC4Za5pkXAbggrZKgcZD6GFMx2L3fr
# 6AexiiFXiyP2uqvMcz45387efpU95CsjBeKZCnH67virrFXV6sn64NqWlnZcEvit
# +RdWR/830L3vzNfrO+Pw8oEyKj5sDZHaY9Dq+Ju1YLgYLh8/qblUv41oPOfWwBxD
# kaoi/RI5HSQqMtR8mqN1ewngl57DVSSRaZMgQoFJKdvobXk9RwDPfH/mqIrlZzkN
# O5RRIA0Ia8j+9dI5gKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkxMzQ2
# MDBaMC8GCSqGSIb3DQEJBDEiBCBPeDi00XfZW4Dz1BwqGG+9OCLnD54xb2J/aImW
# bNzrZjANBgkqhkiG9w0BAQEFAASCAgCxXoUsifaNq3uCHoNRidfofuvq6sejxwQF
# ewWuOgj+XRS/6bN+Zwm/G2v7d+iSBFwUKrCgwAZfFcFUF2qc2ff1Dzn4l1QSWA2u
# 4kasWjXqfrUCd0dLqwJ9rfU1Fd6WrmviKxufvZfxNtUZOqmfd0BJeOTcPpxBZ/Gj
# XPeo1WhudFvm30kn6qZWz9s4S5FE8a2J5tL+L8GDMqVsVoVo5MzmidTCjW49Dxco
# /DMm86Kv9G/zi67TDVviXZnwf16c0amzITjanxPmCtsiOHj5J3iw6MSO742exzrV
# eICv7tTpLxew4te2fE8U9DMWFTUndg5YzoSD3Zrvr/G3tFIzPLhEx5t2SSoBfYYs
# sYIdLDVvqxgLYrmdhShfJlzsBTFQ8ingPOVoJ3GhMWWFvQExPSG+ZwGa15+tzddE
# F1v31/g4XJS92zFwyvtt/O7Zzx3wVx3YPRF8TmrAiwuWDlksLwpUh9TPbUU+Qmxu
# 5IRUmwKEhIEbH7mYbxmLTNEWSWDrbbvARUk4dHHGHuATl0C/4mvheYDvPiAza7o9
# kwcPKjfxmf4zRfo8c/PgSfhEuCmk5cct7AedEsqdXkw5IZuW7oUoBOSF8emZ481U
# FxzHsS4mnMLYJHTrXbeVNbbC2rMZkFb+huTZdNdc6CO0jzWRsUF+cmEAkN//AfC6
# CHeQlSjc+g==
# SIG # End signature block
