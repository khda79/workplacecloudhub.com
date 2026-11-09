# SmartInventory prerequisite audit adapter regression, version 0.6.0.
$ErrorActionPreference='Stop'
$assistantRoot=Split-Path -Parent $PSScriptRoot
. (Join-Path $assistantRoot 'Lib/SmartM365.ConfigurationChecks.ps1')
. (Join-Path $assistantRoot 'Lib/SmartM365.ConfigurationAudit.ps1')
$fixture=Join-Path ([IO.Path]::GetTempPath()) ('si-check-' + [guid]::NewGuid().ToString('N'))
try {
    $null=New-Item -ItemType Directory -Path (Join-Path $fixture 'Setup') -Force
    [IO.File]::WriteAllText((Join-Path $fixture 'Setup/Install-SmartM365-SmartInventoryPrerequisites.ps1'),'throw "must not run"')
    $ok=Invoke-SiPrerequisiteAudit -SmartM365Root $fixture -Runner { param($p) [pscustomobject]@{ExitCode=0;Lines=@('AVAILABLE synthetic')} }
    if ($ok.State -notlike '*imports not verified*' -or $ok.Detail -ne 'AVAILABLE synthetic') { throw 'Incorrect success interpretation.' }
    $missing=Invoke-SiPrerequisiteAudit -SmartM365Root $fixture -Runner { param($p) [pscustomobject]@{ExitCode=1;Lines=@('MISSING synthetic')} }
    if ($missing.State -notlike '*missing*') { throw 'Missing modules were not reported.' }
    try { Invoke-SiPrerequisiteAudit -SmartM365Root $fixture -Runner { param($p) [pscustomobject]@{ExitCode=2;Lines=@()} } | Out-Null; throw 'Unexpected exit accepted.' }
    catch { if ($_.Exception.Message -eq 'Unexpected exit accepted.') { throw } }
    [IO.File]::WriteAllText((Join-Path $fixture 'Setup/Install-SmartM365-SmartInventoryPrerequisites.ps1'), 'param([switch]$AuditOnly) if (-not $AuditOnly) { exit 3 }; Write-Output "SYNTHETIC AUDIT"; exit 0')
    $child=Invoke-SiPrerequisiteAudit -SmartM365Root $fixture
    if ($child.Detail -notmatch 'SYNTHETIC AUDIT') { throw 'Child audit output missing.' }
    $path=Join-Path $fixture 'shared-data'
    $null=New-Item -ItemType Directory -Path $path
    [IO.File]::WriteAllText((Join-Path $path 'fixture.txt'),'preserve')
    $before=(Get-FileHash (Join-Path $path 'fixture.txt')).Hash
    $read=Test-SiSharedPathRead -Path $path
    if ($read.State -ne 'Readable' -or $read.Detail -notmatch 'service-account access.*not verified') { throw 'Read-only path probe misreported.' }
    if ((Get-FileHash (Join-Path $path 'fixture.txt')).Hash -ne $before -or @(Get-ChildItem $path -Force).Count -ne 1) { throw 'Path probe modified fixtures.' }
    if ((Test-SiSharedPathRead -Path (Join-Path $fixture 'missing')).State -ne 'Unavailable') { throw 'Missing path was not detected.' }
    try { Test-SiSharedPathRead -Path 'relative\path' | Out-Null; throw 'Relative path accepted.' }
    catch { if ($_.Exception.Message -eq 'Relative path accepted.') { throw } }
    $one={ [pscustomobject]@{TaskName='SmartM365 Inventory Orchestrator - test';TaskPath='\WCH\';Principal=[pscustomobject]@{UserId='TEST\svc'}} }
    $observed=Get-SiOrchestratorTaskObservation -Tenant 'test' -ListTasks $one
    if ($observed.State -ne 'Observed' -or $observed.Detail -notmatch 'require verification') { throw 'Task observation overstates qualification.' }
    if ((Get-SiOrchestratorTaskObservation -Tenant 'prod' -ListTasks $one).State -ne 'NotObserved') { throw 'Tenant task mismatch.' }
    $customTask={ [pscustomobject]@{TaskName='SmartM365 Inventory Orchestrator - client_eu-prod';TaskPath='\WCH\';Principal=[pscustomobject]@{UserId='TEST\svc'}} }
    if ((Get-SiOrchestratorTaskObservation -Tenant 'client_eu-prod' -ListTasks $customTask).State -ne 'Observed') {throw 'Custom profile task was not observed.'}
    $two={ & $one; [pscustomobject]@{TaskName='SmartM365 Inventory Orchestrator - test';TaskPath='\';Principal=$null} }.GetNewClosure()
    if ((Get-SiOrchestratorTaskObservation -Tenant 'test' -ListTasks $two).State -ne 'Conflict') { throw 'Duplicate tasks not detected.' }
    $standardScript=Join-Path $fixture 'SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator.ps1'
    $standardArgs='-NoProfile -ExecutionPolicy Bypass -File "{0}" -Tenant test -Connect' -f $standardScript
    $standard={ [pscustomobject]@{TaskName='SmartM365 Inventory Orchestrator - test';TaskPath='\WCH\';Principal=[pscustomobject]@{UserId='TEST\svc'};Actions=@([pscustomobject]@{Execute='C:\synthetic\pwsh.exe';Arguments=$standardArgs})} }.GetNewClosure()
    $assessment=Get-SiOrchestratorTaskObservation -Tenant test -SmartM365Root $fixture -AllowedExecutables @('C:\synthetic\pwsh.exe') -ListTasks $standard
    if ($assessment.State -ne 'Observed' -or $assessment.Detail -notmatch 'Triggers, rights and execution require verification') { throw 'Standard action incorrectly assessed.' }
    $assessment=Get-SiOrchestratorTaskObservation -Tenant test -SmartM365Root $fixture -AllowedExecutables @('C:\another\pwsh.exe') -ListTasks $standard
    if ($assessment.State -ne 'ActionMismatch') { throw 'Unexpected executable was not reported.' }
    $tenantDir=Join-Path $fixture 'Config/Tenants'
    $null=New-Item -ItemType Directory -Path $tenantDir -Force
    $thumb='A' * 40
    [IO.File]::WriteAllText((Join-Path $tenantDir 'test.local.json'), ('{"Thumb":"'+$thumb+'","Thumbprint":"'+$thumb+'"}'))
    $certs={ [pscustomobject]@{Thumbprint=$thumb;NotBefore=(Get-Date).AddDays(-1);NotAfter=(Get-Date).AddDays(1);HasPrivateKey=$true} }.GetNewClosure()
    $observation=Get-SiCertificateObservation -Root $fixture -Tenant 'test' -ListCertificates $certs
    if ($observation.State -ne 'Observed' -or $observation.Detail -notmatch 'not verified') { throw 'Certificate observation overstates private-key access.' }
    $none={ @() }
    if ((Get-SiCertificateObservation -Root $fixture -Tenant 'test' -ListCertificates $none).State -ne 'NotObserved') { throw 'Missing certificate misreported.' }
    $expired={ [pscustomobject]@{Thumbprint=$thumb;NotBefore=(Get-Date).AddDays(-2);NotAfter=(Get-Date).AddDays(-1);HasPrivateKey=$true} }.GetNewClosure()
    if ((Get-SiCertificateObservation -Root $fixture -Tenant 'test' -ListCertificates $expired).State -ne 'OutsideValidity') { throw 'Expired certificate misreported.' }
    [IO.File]::WriteAllText((Join-Path $tenantDir 'test.local.json'), ('{"Thumb":"'+$thumb+'","Thumbprint":"'+('B' * 40)+'"}'))
    try { Get-SiCertificateObservation -Root $fixture -Tenant 'test' -ListCertificates $certs | Out-Null; throw 'Differing thumbprints accepted.' }
    catch { if ($_.Exception.Message -eq 'Differing thumbprints accepted.') { throw } }
    'Prerequisite adapter fixtures passed.'
} finally { if (Test-Path -LiteralPath $fixture) { Remove-Item -LiteralPath $fixture -Recurse -Force } }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAHgWL0kCMADbpv
# KiZLHkR8t1yuvtHTP6LoIsRQ20GzQKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICUsYtjGN89CwxeDH+4SlnbUSqXk/LQb08nrH90gZS+lMA0GCSqG
# SIb3DQEBAQUABIIBgCDR7ZSBHv1p0tPLKVH4sT4Hs7Fi/l2UeA1TIC6cBq61UvVo
# hqXV921++EGq5Ujd7UNBy+rzxs3tlA+ZKWcuu46Xb+oSqB1UAEz2jMZwcWkrFVIC
# pwKD9FMKbVCd1aflQIO58v6Qn0mr4GYE9qMLWmvj+pTgzH8/qaM2ebgO2GWjwu4Y
# 9QoBCJqNTfOyCpdx78+aKAhmYBIuny/kVIKUzmgBXqP8C5ScJ8k87rOPy6U74k3x
# CKsgmTHZHcAIOAd2zLNtH5+RxOYTG90nKJdgEVwb90E8n2WM975TrtUN+CuoPNs+
# VWoY6AD3/GEYc/bmmr8eVYXjPltL5d557CgHMyrI9QldY/kfJeHUQHzDgDK5kceB
# 83dFbhui/YX9jipP1VSBc3aWvaMXM6j9g8/ghXdlyAPmXgsQNZb1Ea0q8dVTHDTI
# cnrX1xNdNKGwM8Z/XTnkHg7ROqSYI9U4AlvJ978AFkteqHmTwmgBZLrFt3vfxkla
# r2G45mDhiSPyYMnC16GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDkyMTA5
# MDNaMC8GCSqGSIb3DQEJBDEiBCCZTv6BDEMU9C4W7bBU0uJYDlgG763VAfNKu3OL
# X7I/JzANBgkqhkiG9w0BAQEFAASCAgC2SY0TeInTw6DYaHH78i5Rki3uOaEYetXv
# GJP/ytArukLKf0uLXfF7ItB7cS58kGy4tNwGaNpjdPNdbM0lVxkX36kShNxU01U3
# eqnVvRD7yIhoDoKAIyzgSWiPsn0q1FsZTBe9aoZPzTO5lszjkXrFfgbWgBtJEGe5
# mf01AcmBa96ggKUGFKxv4dx3x6UIrf4WSOWtqUjzeYAtK21QZ/+MKLWiHZH4u3QT
# LnZFqLAv6DFU2GgpRvB8fiHOpuxWWg6R+Qzc3yN3PR6bCVvilbwpwDobgZe7KZMx
# +bADefnkBRDkVg25VnLegTs1luqMpYm2pgrIYPqlKS1EJKpEvQNHMZxO5bvxtyV3
# kKUjV0yh0bgy7pIlKqTI7LolsFyVE7hdWmPsUvJQvffBI/Vgtc4F83OhLP/5G/K9
# X0DtFato4Xvt3QRgSQe1OqS3iPnOkBJ2PUgnAJvI89437nF18taSncqjHd/FlijF
# qJqHbH8UkLVVkGWN7sGWTg//K7HJmxTI11c2D+11sDDE+jj2ZwZ9a2K0PIqXzHui
# vzX4qtavdMnAOOst73Iuo8LuObiBwk+PuEp8gNfHs6fKceM7073FqzJ76yq2YYnD
# rrGfOQ47Rroei1nIAHDral4rmGl9vSj8Z5wMAiyPZeczsT5yIKce7C3JksTeSG7i
# luE0+FidBg==
# SIG # End signature block
