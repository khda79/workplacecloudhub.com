#requires -Version 7.0
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365.Orchestrator.PendingWork.psm1') -Force
Import-Module (Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365.Orchestrator.Maintenance.psm1')
Import-Module (Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365.Orchestrator.Pipeline.psm1')
$root=Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-CancelPending-' + [guid]::NewGuid().ToString('N'))
$worker=$null
function Write-Fixture($Path,$Value) {
    $null=New-Item -ItemType Directory -Path (Split-Path $Path -Parent) -Force
    [IO.File]::WriteAllText($Path,($Value | ConvertTo-Json -Depth 15),[Text.UTF8Encoding]::new($false))
}
function New-Heartbeat($control) {
    [pscustomobject]@{ Tenant='test';ScriptVersion='1.5.90';Timestamp=[datetime]::UtcNow.ToString('o');Lifecycle='Running';RunningJobs=@([pscustomobject]@{Name='AlreadyRunning';Pid=999})
        MaintenanceProtocol=1;MaintenanceCatchUpProtocol=1;MaintenanceResetProtocol=1;PipelineCancellationProtocol=1;StatePersistenceHealthy=$true;MaintenanceHealthy=$true
        MaintenanceEnabled=$control.Enabled;MaintenanceRevision=$control.Revision }
}
try {
    $null=New-Item -ItemType Directory -Path (Join-Path $root 'Config') -Force
    Write-Fixture (Join-Path $root 'Config/Orchestrator-Cluster.json.txt') ([pscustomobject]@{ExpectedOrchestratorServers=@('SERVER-A');PeerHeartbeatStaleMinutes=5})
    $zero=Initialize-SmartM365OrchestratorMaintenance $root
    Write-Fixture (Join-Path $root 'SERVER-A/Orchestrator-Heartbeat.json.txt') (New-Heartbeat $zero)
    $incompatible=New-Heartbeat $zero; $incompatible.MaintenanceResetProtocol=0
    Write-Fixture (Join-Path $root 'SERVER-A/Orchestrator-Heartbeat.json.txt') $incompatible
    $refused=$false
    try { $null=Invoke-SmartM365CancelPendingWork -SharedDataFolderPath $root -Tenant test -ExpectedRevision 0 -Reason 'unsupported test' -AcknowledgementTimeoutSeconds 1 } catch { $refused=$true }
    if (-not $refused -or (Get-SmartM365OrchestratorMaintenanceState $root).Enabled) { throw 'An incompatible resident must be refused before pausing.' }
    Write-Fixture (Join-Path $root 'SERVER-A/Orchestrator-Heartbeat.json.txt') (New-Heartbeat $zero)
    $selection=[pscustomobject]@{SelectedJobs=@([pscustomobject]@{Name='AlreadyRunning'},[pscustomobject]@{Name='StillPending'})}
    $request=New-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $root -Tenant test -Pipeline Jobs -Selection $selection -ManifestHash 'fixture'
    $null=Set-SmartM365OrchestratorPipelineJobStatus -SharedDataFolderPath $root -BatchId $request.BatchId -JobName 'AlreadyRunning' -Status Running -OwnerServer 'SERVER-A'
    $orchestratorFolder=Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
    $guiAst=[Management.Automation.Language.Parser]::ParseFile((Join-Path $orchestratorFolder 'SmartM365-Inventory-Orchestrator-GUI.ps1'),[ref]$null,[ref]$null)
    $workerFunction=$guiAst.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-GuiCancellationWork'},$true)
    . ([scriptblock]::Create($workerFunction.Extent.Text))
    $work=Get-GuiCancellationWork
    $progress=[hashtable]::Synchronized(@{Message='';LogError=''})
    $inputData=[pscustomobject]@{Kind='PendingWork';Phase='Validate';Root=$root;Tenant='test';Reason='offline test';Revision=[long]-1;LogPath='';ModulePaths=@(
        (Join-Path $orchestratorFolder 'SmartM365.Orchestrator.Management.psm1'),(Join-Path $orchestratorFolder 'SmartM365.Orchestrator.Pipeline.psm1'),
        (Join-Path $orchestratorFolder 'SmartM365.Orchestrator.Maintenance.psm1'),(Join-Path $orchestratorFolder 'SmartM365.Orchestrator.Insights.psm1'),
        (Join-Path $orchestratorFolder 'SmartM365.Orchestrator.PendingWork.psm1'))}
    $validation=& $work $progress $inputData
    if ($validation.Phase -ne 'PendingWorkValidate' -or $validation.Revision -ne 0 -or @($validation.BatchIds).Count -ne 1) { throw 'GUI cancellation preview did not enumerate the active request.' }
    $inputData.Tenant='wrong-tenant'; $refused=$false
    try { $null=& $work $progress $inputData } catch { $refused=$_.Exception.Message -match 'tenant' }
    if (-not $refused -or (Get-SmartM365OrchestratorMaintenanceState $root).Enabled) { throw 'GUI must preserve the preflight error and not pause the wrong tenant.' }
    $inputData.Tenant='test'
    $inputData.Phase='Publish'; $inputData.Revision=$validation.Revision
    $worker=Start-Job -ArgumentList $root -ScriptBlock {
        param($root)
        $controlPath=Join-Path $root 'Config/Orchestrator-Maintenance.json.txt'
        $heartbeatPath=Join-Path $root 'SERVER-A/Orchestrator-Heartbeat.json.txt'
        $end=[datetime]::UtcNow.AddSeconds(20)
        while ([datetime]::UtcNow -lt $end) {
            try {
                $control=Get-Content -LiteralPath $controlPath -Raw | ConvertFrom-Json
                if ($control.Enabled) {
                    $heartbeat=Get-Content -LiteralPath $heartbeatPath -Raw | ConvertFrom-Json
                    $heartbeat.MaintenanceEnabled=$true; $heartbeat.MaintenanceRevision=$control.Revision; $heartbeat.Timestamp=[datetime]::UtcNow.ToString('o')
                    $temp=$heartbeatPath+'.testtmp'; [IO.File]::WriteAllText($temp,($heartbeat | ConvertTo-Json -Depth 10),[Text.UTF8Encoding]::new($false))
                    [IO.File]::Move($temp,$heartbeatPath,$true)
                    break
                }
            } catch { }
            Start-Sleep -Milliseconds 100
        }
    }
    $published=& $work $progress $inputData
    if ($published.Phase -ne 'PendingWorkPublish') { throw 'GUI cancellation did not return the published result.' }
    $result=$published.Result
    if ($result.RunningJobsStopped -or $result.Revision -ne 2 -or (Get-SmartM365OrchestratorMaintenanceState $root).Enabled) { throw 'Reset did not resume correctly or claimed to stop running jobs.' }
    $requestAfter=Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $root -BatchId $request.BatchId
    if (-not $requestAfter.CancellationRequested -or @($requestAfter.Jobs | Where-Object { $_.JobName -eq 'StillPending' -and $_.Status -eq 'Cancelled' }).Count -ne 1 -or @($requestAfter.Jobs | Where-Object { $_.JobName -eq 'AlreadyRunning' -and $_.Status -eq 'Running' }).Count -ne 1) { throw 'Cancellation did not preserve running work and cancel pending request work.' }
    $hb=Get-Content (Join-Path $root 'SERVER-A/Orchestrator-Heartbeat.json.txt') -Raw | ConvertFrom-Json
    if (@($hb.RunningJobs).Count -ne 1) { throw 'The running job was not preserved in the synthetic heartbeat.' }
    $timeoutRoot=Join-Path $root 'TimeoutCase'
    $null=New-Item -ItemType Directory -Path (Join-Path $timeoutRoot 'Config') -Force
    Write-Fixture (Join-Path $timeoutRoot 'Config/Orchestrator-Cluster.json.txt') ([pscustomobject]@{ExpectedOrchestratorServers=@('SERVER-A');PeerHeartbeatStaleMinutes=5})
    $timeoutZero=Initialize-SmartM365OrchestratorMaintenance $timeoutRoot
    Write-Fixture (Join-Path $timeoutRoot 'SERVER-A/Orchestrator-Heartbeat.json.txt') (New-Heartbeat $timeoutZero)
    $failed=$false
    try { $null=Invoke-SmartM365CancelPendingWork -SharedDataFolderPath $timeoutRoot -Tenant test -ExpectedRevision 0 -Reason 'timeout test' -AcknowledgementTimeoutSeconds 1 } catch { $failed=$true }
    if (-not $failed -or -not (Get-SmartM365OrchestratorMaintenanceState $timeoutRoot).Enabled) { throw 'A missing acknowledgement must fail and leave the safe pause active.' }
    'PASS: coordinated pending-work cancellation leaves running jobs alone'
}
finally {
    if ($worker) { Stop-Job $worker -ErrorAction SilentlyContinue; Remove-Job $worker -Force -ErrorAction SilentlyContinue }
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDnThtyyIyRcqnC
# gsRJ2y5FD7eEBZuJPNURm87GPTWNOaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIKpfxKylg4QtDlJkusMwmzMoAyve02Hoa6mOQSSo+KUrMA0GCSqG
# SIb3DQEBAQUABIIBgIEhlb0GPrgMCI8ilC4Q3lldTpyRMCJkIR0a2aqPtxFkmj8R
# K1RYXl5qsT38UpSKaJbaZg3huhKQ0V36ozJLD5sUFy9orITNHG6BeK3Hmubp24xX
# IV6G2H5z30RnB7Mj2PfqVNqmNlEyDUGAK0n4S5ZXiihycL4pPcPjNfFlh6xful6b
# m3QHHwB7GRrMWO/XkEbAW+XtBVYiy/zz8Fnm5MkD917UIeFrOH5EPEIu4t0ez7xJ
# WeX6uI9SijeVjftaQ4s0S07C1jizIgl9IUNiZMV9oW6cy7W4nmIBeUgQ5Wgw2a6N
# 9ANc2JORvKbW1TdlCXbPHI5xbxeMvOq74ns/aaPCjUQ/8quudYkTSAESp/oaoM/T
# IO/7YuC7wO9uE/pq6F7fL7pn9L1J7panv4AgbDIRmEJYv3HldQlNHMIS+dRxViYJ
# bwvjIkpuOj6e+kMB7r/k3KWXV+CVKmditTrAr7E6R61ATuqf/pZBLvvkhAX/4BgO
# LmrNDOruPhfWlU2fEaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTMxMzE2
# MjJaMC8GCSqGSIb3DQEJBDEiBCAvJiylhjLIvlTq7N/WCSq1iWGgBSJhySw+33mq
# 8UOxCzANBgkqhkiG9w0BAQEFAASCAgBk7rJFudIuvbaCyZSeUnmpm+otxRkGEA35
# 7rx2OD8lcPx8127d9FFFZGJZ93IUoElmjRGQiB9gi8N/U/X4mxY0paX4GLd8/cks
# 2cwQVY3DMfLrsePSCUoG1YCrlxfHrFKjfao3U2Tzv4aYL1ImJ05JmEkhtS07pikM
# tsD7anEOPrB/tMaQXxAigZ2rMnzcR2GOFFOKU3quTaAdUfagP0di1bMrH0zEYlkd
# FACll7kAMpVlQtTSuoXBcZX35hDXFsMTaz9Eru9xV92uIvM3mh+w9hTDQVJxvhrk
# UoBP2klppGHkFWfQutURkiD1T/Q8gV3XKJzeLQolRUzwcZ4tZlumNdgE0ic02ceV
# J1bAvmw2ACWy5VIG5z5ZBhj16WYGT9Ehr+GztMqF9XNVLwZduZM3vClZFHFQMZt6
# lkVF2doGzjvM5MtTL7mWcno7kuUEglYgzZBKQ7J2NOT2RoGcG07f1w1j5Wi5HC1p
# DpfAMSAU0vH/ywb9NBB7zZx9r/Ry4+70h4KJRDlxmZMxsmdgIukkn8vyz8xXkfwJ
# wHST++47Wny2+Va9qHkW1AKop7Bq8oTSYrYXawJsF+j0KPSFTsE/8kNbHaX78kN0
# Epw3zkClxHty+E042TZaQaC5IlWkacXkP2GtAD0qwfkIxePvWuI5Yd9WYYW8SaOB
# KDqsnmbGpg==
# SIG # End signature block
