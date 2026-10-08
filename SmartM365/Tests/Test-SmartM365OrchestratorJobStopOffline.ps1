# Offline job-stop contract test. No inventory collector or production share is used.
$ErrorActionPreference = 'Stop'
$folder = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
Import-Module (Join-Path $folder 'SmartM365.Orchestrator.Management.psm1') -Force
Import-Module (Join-Path $folder 'SmartM365.Orchestrator.JobStop.psm1') -Force
Import-Module (Join-Path $folder 'SmartM365.Orchestrator.Pipeline.psm1') -Force
$source = Join-Path $folder 'SmartM365-Inventory-Orchestrator.ps1'
$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($source, [ref]$tokens, [ref]$errors)
if ($errors.Count) { throw ($errors.Message -join '; ') }
$names = @('Update-OrchestratorJobStopRequests', 'Complete-JobRun')
$definitions = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -in $names }, $true))
if ($definitions.Count -ne $names.Count) { throw 'Required runtime functions were not found.' }
$testModule = New-Module -Name OrchestratorJobStopOffline -ScriptBlock ([scriptblock]::Create((@($definitions | ForEach-Object { $_.Extent.Text }) -join "`n") + @'
function Get-JobState { param($JobName) return $script:State.Jobs[$JobName] }
function ConvertTo-StateTime { param([datetime]$Value) return $Value.ToString('o') }
function Save-OrchestratorState { $script:Saved = $true }
function Test-OrchestratorStatePersistenceReady { return [bool]$script:Saved }
function Write-OrchestratorLog { param($Message, $Level) }
function Set-OrchestratorPipelineJobStatus { throw 'Unexpected pipeline status call.' }
function Add-JobRunCsvRow { param($JobName,$ScheduledTime,$StartTime,$EndTime,$DurationSec,$ExitCode,$Status,$RetryCount,$LogPath) $script:CsvStatus = $Status }
function Get-JobRunsCsvPath { return 'synthetic.csv' }
function Invoke-OrchestratorSharePointUpload { param($LocalFilePath,$Reason,[switch]$Force) return $false }
function Send-JobResultEmail { throw 'Unexpected mail send.' }
Export-ModuleMember -Function Update-OrchestratorJobStopRequests, Complete-JobRun
'@))
Import-Module $testModule -Force

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-jobstop-offline-' + [guid]::NewGuid().ToString('N'))
try {
    $server = 'TEST-SERVER'
    $jobName = 'Synthetic-Inventory'
    $pidValue = 8123
    $started = (Get-Date).AddMinutes(-10)
    $startedText = $started.ToString('o')
    $heartbeatPath = Join-Path (Join-Path $tempRoot $server) 'Orchestrator-Heartbeat.json'
    $heartbeat = [pscustomobject]@{
        Timestamp = (Get-Date).ToString('o'); Lifecycle = 'Running'; Tenant = 'test'; JobStopProtocol = 1
        RunningJobs = @([pscustomobject]@{ Name = $jobName; Pid = $pidValue; StartTime = $startedText })
    }
    Write-SmartM365OrchestratorJsonAtomically -Path $heartbeatPath -Document $heartbeat
    $request = Request-SmartM365OrchestratorJobStop -SharedDataFolderPath $tempRoot -Server $server -JobName $jobName -ProcessId $pidValue -StartTime $startedText -Reason 'Operator test stop' -Tenant test
    $savedRequest = Read-SmartM365OrchestratorJson -Path $request.Path
    if ($savedRequest.Status -ne 'Requested' -or $savedRequest.Reason -ne 'Operator test stop') { throw 'Stop request was not durably recorded.' }
    $rejected = $false
    try { Request-SmartM365OrchestratorJobStop -SharedDataFolderPath $tempRoot -Server $server -JobName $jobName -ProcessId 9999 -StartTime $startedText -Reason 'Wrong process' -Tenant test | Out-Null }
    catch { $rejected = $true }
    if (-not $rejected) { throw 'A changed PID was accepted.' }
    $rejected = $false
    try { Request-SmartM365OrchestratorJobStop -SharedDataFolderPath $tempRoot -Server $server -JobName $jobName -ProcessId $pidValue -StartTime $started.AddSeconds(-10).ToString('o') -Reason 'Wrong start' -Tenant test | Out-Null }
    catch { $rejected = $true }
    if (-not $rejected) { throw 'A changed process start time was accepted.' }
    $heartbeat.Timestamp = (Get-Date).AddMinutes(-10).ToString('o')
    Write-SmartM365OrchestratorJsonAtomically -Path $heartbeatPath -Document $heartbeat
    $rejected = $false
    try { Request-SmartM365OrchestratorJobStop -SharedDataFolderPath $tempRoot -Server $server -JobName $jobName -ProcessId $pidValue -StartTime $startedText -Reason 'Stale owner' -Tenant test | Out-Null }
    catch { $rejected = $true }
    if (-not $rejected) { throw 'A stale owner heartbeat was accepted.' }

    & $testModule {
        param($root,$serverName,$name,$id,$start)
        $env:COMPUTERNAME = $serverName
        $script:Tenant = 'test'
        $script:Settings = @{ SharedDataFolderPath = $root; JobMailMode = 'Never' }
        $script:State = @{ Jobs = @{ $name = @{ Running = @{ Pid = $id; StartTime = $start }; LastStatus = 'Running'; PendingRetry = $null } } }
        $script:Manifest = @{ JobsByName = @{ $name = [pscustomobject]@{ MaxRetries = 3 } } }
        $script:RunningJobs = @{ $name = @{ Process = [pscustomobject]@{ Id = $id }; StartTime = [datetime]::Parse($start); Occurrence = Get-Date; LogPath = ''; Attempt = 0; ProcessName = 'pwsh' } }
        $script:Saved = $false
    } $tempRoot $server $jobName $pidValue $startedText
    Update-OrchestratorJobStopRequests
    $ack = Read-SmartM365OrchestratorJson -Path $request.Path
    if ($ack.Status -ne 'Stopping') { throw "Owner did not accept the stop request: $($ack.Status)" }
    & $testModule {
        param($name,$path)
        if (-not $script:Saved -or -not $script:RunningJobs[$name].OperatorStopRequested -or
            $script:State.Jobs[$name].Running.OperatorStopRequestPath -ne $path) { throw 'Stop intent was not persisted before acknowledgement.' }
    } $jobName $request.Path
    & $testModule {
        param($name)
        $run = $script:RunningJobs[$name]
        Complete-JobRun -JobName $name -RunInfo $run -StatusHint Cancelled -ExitCode $null -EndTime (Get-Date) -ErrorText 'Stopped by operator for offline test.'
        if ($script:State.Jobs[$name].LastStatus -ne 'Cancelled' -or $script:State.Jobs[$name].PendingRetry -or $script:CsvStatus -ne 'Cancelled') { throw 'Cancelled run was retried or recorded incorrectly.' }
    } $jobName
    $final = Read-SmartM365OrchestratorJson -Path $request.Path
    if ($final.Status -ne 'Stopped') { throw 'Stop result was not recorded.' }
    $selection = [pscustomobject]@{ SelectedJobs = @([pscustomobject]@{ Name = $jobName }) }
    $batch = New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $tempRoot -Tenant test -Pipeline Jobs -Selection $selection -ManifestHash synthetic
    $null = Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $tempRoot -BatchId $batch.BatchId -JobName $jobName -Status Running -OwnerServer $server
    $blocked = Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $tempRoot -BatchId $batch.BatchId -JobName $jobName -Status Cancelled -OwnerServer $server
    if ($blocked.Status -ne 'Running') { throw 'Ordinary batch cancellation changed a running job.' }
    $accepted = Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $tempRoot -BatchId $batch.BatchId -JobName $jobName -Status Cancelled -OwnerServer $server -AllowRunningCancellation
    $summary = Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $tempRoot -BatchId $batch.BatchId
    if ($accepted.Status -ne 'Cancelled' -or $summary.OverallStatus -ne 'Failed') { throw 'Operator stop did not finalize the running pipeline job.' }
    'PASS: exact-run request, stale PID/start/owner rejection, owner acknowledgement, Cancelled result without retry, and running pipeline finalization.'
}
finally {
    $resolved = [IO.Path]::GetFullPath($tempRoot)
    $tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    if ($resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($resolved) -like 'SmartM365-jobstop-offline-*' -and
        (Test-Path -LiteralPath $resolved)) { Remove-Item -LiteralPath $resolved -Recurse -Force }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDahrrcwwgYSW2Z
# k850kE7hdL+yGjRxFeGcvsbhGjT8haCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIC+JQ/ltfwuGFIprUO0uTzRiCepnw0h5eI0au/rqlXKtMA0GCSqG
# SIb3DQEBAQUABIIBgB4u9g7/ib1h1UiQTDXZbhqVS0h+mTCsLuII400YxM/tO90Q
# +n7Bweoe6+QhZedA8lrvXdsLIPMVJ/maIftE9dB/itExxVLws/y5kFLiiiTbJ+Lj
# 7WL9rhng6B8+5PeRKKyUP5t/E55dPSwgLtm3Rop9+SiO5rSjzLQ4Hp5k2KQrACm7
# 5bICT0l2h91UOLQmdtg5pX+Zwdd7xIhuZDokTf9bGkQ7j3d53aiZXuANeT8YsWKv
# 4NsSr5O+iyNKtD2eCetbdRcjjzjtEuuSZ0F64IXM20J39qQvbKh5JmY5pLWzzvim
# J9df06hT0WHP0CWp7bwYuYPXpueUwHsPXHdiMhU8MHiUcfhy3GAivXPRScJUr5Qq
# U16EIXiE2x3H7Iol7NKGya5DunB9ptw7ANLRAcOuCf5Bi0E7wIJ7SFfDueZK6NLe
# Vzw/uGn4+4ceYxvFmdRg1hBdn5k3NPTo0Gl/VGB4Bvk/nlyen+VRLnH1xFD4Hcme
# TLgh3/6NKCC6pyBNZKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDgyMDM2
# MTlaMC8GCSqGSIb3DQEJBDEiBCBTqBQmKOurUw2Xm+JCSMe6Q5zKD1gtvct64wNf
# feUrODANBgkqhkiG9w0BAQEFAASCAgABYc0PggWZ9xdqRub+N0vVCjgJJlWw7wA/
# SjOFvyKR2mIlGzZtMtf3Ur1rYwKIe9N0pw2zQT0vASNaUvpyo1KYbVqRHFfGr0Fj
# /ycsf2ek1kdF/T72ClooCbFkPsAcRgwyR2I6RHXdWkz/ZCv7XvukpYVihGmyov0A
# 6ok1Dc6KlTyl8UQI2aXGbZa/DB9AYAmKPKvwVr6Gv7PoTFMKNvAmFp5SNy4lkcOb
# KGnPWW8mHW9KynyNXVMnRGd1YPqRA90u4KAuqsGszPe3jjPa2hEf2O8VRAXj8NFu
# Ub0P4t1ZczehEMabIgYSvzdWwV5BGi1gBkJxlPybykgf2AGApBSb2tk4590lwPL4
# hzHqOAseZTJy+K14TFvhMaB29+yPpiu7bvnlm23S8kwois8sQdCab+rvibzKKC48
# Zukf2he6aVO1GMrmuRuTfHnyyqpPMXcmgza7qJbtbRSk6fKDe+ANRB0P6iUGbVwz
# X8ia6bOZe4mA0nRyzp/buBYqtPwSzV8B/60zu0pE8yQ1iOgwepYB4sTA3u3GuO+Y
# cQ9LpuXlv3O46uJ8eXe7wvxR54bPi6gCwpkBGch2kAIAAPyrzCGrkVa2gnzjwR/l
# NAKsCYxkhTqRqp19s5U6h+vaCTENGkavpktco+2U3AUVFTh1Vgb5pBIUZ5T2IDSj
# pvH5FT7jtg==
# SIG # End signature block
