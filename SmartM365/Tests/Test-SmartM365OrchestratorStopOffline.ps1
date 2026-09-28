#requires -Version 7.0
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$source = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator.ps1'
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($source,[ref]$null,[ref]$errors)
if ($errors.Count) { throw 'Source parsing failed.' }
$definitions = foreach ($name in @('Test-OrchestratorProcessStartTime','Get-OrchestratorLockOwner','Get-OrchestratorProcessCandidate','Request-OrchestratorStop','Test-OrchestratorStopRequested')) {
    $node = $ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    if (-not $node) { throw "Missing $name" }
    $node.Extent.Text
}
$module = New-Module -ScriptBlock ([scriptblock]::Create($definitions -join "`n"))
$passed = 0
function Check($Value,$Message) { if (-not $Value) { throw $Message }; $script:passed++ }
try {
    & $module {
        $script:Tenant = 'prod'
        $script:Settings = [pscustomobject]@{LockPath='synthetic.lock';StopRequestPath='synthetic.json.txt'}
        $script:When = [datetime]::UtcNow.AddHours(-1)
        $script:LockExists = $true
        $script:Alive = $true
        $script:ExitDuringWait = $true
        $script:Request = $null
        $script:Writes = 0; $script:Consumed = 0
        $script:ActualOwner = (Get-Command Get-OrchestratorLockOwner).ScriptBlock
        $script:ActualCandidates = (Get-Command Get-OrchestratorProcessCandidate).ScriptBlock
        $script:Candidate = [pscustomobject]@{ProcessId=12345;Name='pwsh.exe';CreationDate=$script:When;CommandLine='pwsh -File "C:\Code\SmartM365-Inventory-Orchestrator.ps1" -Tenant prod -Connect'}
        $script:Candidates = @($script:Candidate)
        function script:Test-Path {param($LiteralPath) return $script:LockExists}
        function script:Get-Content {param($LiteralPath,[switch]$Raw,$ErrorAction) @{Pid=12345;StartTimeUtc=$script:When.ToString('o')} | ConvertTo-Json}
        function script:Get-Process {param($Id,$ErrorAction) if($script:Alive){[pscustomobject]@{ProcessName='pwsh';StartTime=$script:When}}}
        function script:Get-CimInstance {param($ClassName,$Filter,$ErrorAction) $script:Candidates}
        function script:Write-OrchestratorLog {param($Message,$Level)}
        function script:Write-Host {param($Object,$ForegroundColor)}
        function script:Get-OrchestratorRunUserName {'SYNTHETIC'}
        function script:Get-SmartM365JsonReadPath {param($Path,[switch]$Optional) if($script:Request){'synthetic.json.txt'}}
        function script:Read-SmartM365JsonDocument {param($Path) [pscustomobject]@{Document=$script:Request;SHA256='synthetic'}}
        function script:Write-FileAtomically {param($Path,$Content) $script:Writes++; $script:Request=$Content|ConvertFrom-Json}
        function script:Complete-SmartM365JsonConsumption {param($Path,$Owner,$ExpectedSHA256) $script:Consumed++;$script:Request=$null}
        function script:Stop-OrchestratorScheduledTaskIfRunning {throw 'Must never force scheduled task stop.'}
        function script:Start-Sleep {param($Seconds) if($script:ExitDuringWait){$script:Alive=$false}else{Microsoft.PowerShell.Utility\Start-Sleep -Milliseconds 1100}}
    }
    Check ((& $module {Get-OrchestratorLockOwner}).Pid -eq 12345) 'Valid resident owner rejected by filename regex.'
    Check ((& $module {Request-OrchestratorStop -TimeoutSeconds 1}) -eq 0) 'Normal graceful stop failed.'
    & $module {$script:LockExists=$false;$script:Alive=$true;$script:Request=$null}
    Check ((& $module {Request-OrchestratorStop -TimeoutSeconds 1}) -eq 0) 'Single candidate without lock not stopped gracefully.'
    Check ((& $module {$script:Writes}) -eq 2) 'Stop request was not published for candidate.'
    & $module {$script:Alive=$true;$script:ExitDuringWait=$false}
    Check ((& $module {Request-OrchestratorStop -TimeoutSeconds 1}) -eq 1) 'Missing lock falsely reported stopped process.'
    Check ((& $module {$script:Writes}) -eq 2) 'Existing matching request was rewritten.'
    Check ((& $module {$null -ne $script:Request}) ) 'Timeout removed request.'
    & $module {$script:Candidate.CreationDate=$script:When.AddTicks(-6);$script:Request.TargetStartTimeUtc=$script:When.AddTicks(-6).ToString('o')}
    Check ((& $module {Request-OrchestratorStop -TimeoutSeconds 1}) -eq 1) 'CIM precision mismatch falsely reported process exit.'
    Check ((& $module {([datetime]$script:Request.TargetStartTimeUtc).ToUniversalTime() -eq $script:When.ToUniversalTime()})) 'CIM request was not normalized for old consumers.'
    # Keep the existing assertions relative to their earlier publication count.
    & $module {$script:Writes=2;$script:Candidate.CreationDate=$script:When}
    & $module {$script:Request.TargetPid=99999}
    $rejected=$false
    try { & $module {Request-OrchestratorStop -TimeoutSeconds 1} } catch {$rejected=$true}
    Check $rejected 'Existing request for another process overwritten.'
    Check ((& $module {$script:Request.TargetPid}) -eq 99999) 'Foreign pending request modified.'
    & $module {$script:Request.TargetPid=12345}
    & $module {$script:Candidates=@($script:Candidate,$script:Candidate)}
    Check ((& $module {Request-OrchestratorStop -TimeoutSeconds 1}) -eq 1) 'Ambiguous candidates accepted.'
    Check ((& $module {$script:Writes}) -eq 2) 'Ambiguous request changed.'
    & $module {$script:Candidates=@();$script:Alive=$false}
    Check ((& $module {Request-OrchestratorStop -TimeoutSeconds 1}) -eq 0) 'Absent process not recognized.'
    Check ((& $module {$null -ne $script:Request})) 'Absent lock/process erased existing request.'
    & $module {$script:Candidate.CommandLine='pwsh -File C:\Code\SmartM365-Inventory-Orchestrator.ps1 -Tenant production';$script:Candidates=@($script:Candidate)}
    Check (@(& $module {Get-OrchestratorProcessCandidate}).Count -eq 0) 'Tenant prefix selected wrong tenant.'
    & $module {$script:LockExists=$true;$script:Alive=$true}
    Check ($null -eq (& $module {Get-OrchestratorLockOwner})) 'Resident owner from another tenant accepted.'
    & $module {$script:Request=[pscustomobject]@{TargetPid=$PID+100;TargetStartTimeUtc=$script:When.ToString('o')}}
    Check (-not (& $module {Test-OrchestratorStopRequested})) 'Request for another PID consumed.'
    Check ((& $module {$script:Consumed}) -eq 0) 'Foreign request deleted.'
    & $module {$script:Request=[pscustomobject]@{TargetPid=$PID;TargetStartTimeUtc=$script:When.AddHours(-1).ToString('o')};$script:Alive=$true}
    Check (-not (& $module {Test-OrchestratorStopRequested})) 'Reused PID consumed stale request.'
    & $module {$script:Request=[pscustomobject]@{TargetPid=$PID;TargetStartTimeUtc=$script:When.ToString('o')}}
    Check (& $module {Test-OrchestratorStopRequested}) 'Matching request was not consumed.'
    & $module {$script:Request=[pscustomobject]@{TargetPid=$PID;TargetStartTimeUtc=$script:When.AddTicks(-6).ToString('o')}}
    Check (& $module {Test-OrchestratorStopRequested}) 'CIM precision request was not consumed.'
    Check (-not (& $module {Test-OrchestratorProcessStartTime -Actual $script:When -Expected $script:When.AddMilliseconds(-1)})) 'Tolerance accepted a different process start.'
    $realProcess=Get-Process -Id $PID
    $realCim=Get-CimInstance Win32_Process -Filter "ProcessId=$PID"
    Check (& $module {param($a,$e) Test-OrchestratorProcessStartTime -Actual $a -Expected $e} $realProcess.StartTime $realCim.CreationDate) 'Actual local CIM/native process identity mismatch.'
    & $module {function script:Get-CimInstance {throw 'Synthetic process enumeration failure.'};$script:LockExists=$false}
    $rejected=$false
    try { & $module {Request-OrchestratorStop -TimeoutSeconds 1} } catch {$rejected=$true}
    Check $rejected 'Process enumeration failure reported success.'
    [pscustomobject]@{Passed=$passed;Scope='Synthetic stop scenarios plus read-only local CIM identity check; no real process or task stopped'}
} finally { Remove-Module $module -Force }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAKkL+usUlDHHg9
# yl4zgcQXBUGjx7efS5eD41Ag/qbxl6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIESlacUqeCMQyQitqx3mBANpQgDjEXLwuU17eYSnMz6UMA0GCSqG
# SIb3DQEBAQUABIIBgJ/4hft7n69lhsXQXCggl9MdI5yxP95wQFS4unFbyuR3etS9
# /NQmyhP7/eeiIOhHtWKXoZN0BSY8d5rdI+R1kPvVppEefY6agFzV8fvav4+ALs0K
# pQWWxD30q7UDs7K4FowE9EUYT6Hvz2ww7mGugQXfrG6V3S2ZCUVWzeTd9kTLSnql
# tCx4r/iS2lb4CJ51+8lLcyoJSUTnyX3OFZeXEmwQXeaSxz2dE5UlvEKWG7+d+KXb
# j1O1t1kKWdgaWiHhaF2s4ZvmRX6f4NLiZVFFntNEFb4kZzz/BDD0hDI6KbKq0uDr
# 78Kb9cNsTsBBwqakDkRmHX9YrBldQvLDpAiGbyVtiDwCYyipoY0MPeGLx+EwbMBR
# Hi2818icWQIR3eCzqjF+YzL9iZwdHnPGWVY3XnuELy+JExoNs3KKHjCmH1hAeoKu
# 8xpgjf51bWf81d3Qv8jGyUXySRQzrz+DuhtyvRLOmZNKiUwPri7r5JBRFPR33buF
# 3jYymN+0lvADWOB6OqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjgxOTI5
# MDVaMC8GCSqGSIb3DQEJBDEiBCDF31hDuZieRysbRQ8Use1yS5uRh1CtwPuoCOWW
# wNjcyTANBgkqhkiG9w0BAQEFAASCAgB5M76XNHQz6gl5WZWUGf/R29A3uDe2D/2q
# 6yBpJydm6HZuP+6ZE/WrsBFo1Kp4ZnF/n3n84zCejfhJ+FoExzuBTrmwi23PHKoW
# gQeQBLu7J3aMkhTtw0JazSJPPKvKYlL1P37u7qEeDMGWJtE5E5QbHI/dib3PP+JS
# X7hmAnSUDH5SFYWSiMd5VWSl6YbElfQMlLY4YelY6eufrWFH09oebg+lywK6mJ5p
# YNH+YFXEU72uff+uD7kP1xSKs7zanKFDQZunlEGWwFFy+tH/UuOZJZAr63z3f605
# LqcQN/vmYt5PRXJS+SCWxDi5MaLbxTT3bsKUXsxW1D5gYuZC9PrMqQImLZHN0PBU
# U5BjcLZb/FPDNZDIU/FzCZr5tZpP9nlUCaFHRNqC4f2npMRsUOfhB5mCHUtk/7Fc
# Ofr1KjXY6vD/l4ifTkedd2PYZTHhIDc99yKBH8JUGhbCFhfFeKHCy+i3a9X4jp8w
# QRTK8F57+iqT4JR93so/ixbv7+5KmK2/91U02jm2UKFy5E6HBY1sdnJFpiL8yWzB
# ILfJwyWw5WPOODNYYE0jOylzz9xWA/LpGy6eT77wgE9M2jxHH5AGLLVL2KYdzRet
# hy9BY0y3cGyIr5tPMUleMr/vtvsJAoCF6RyFyBNZLYradlm2eSkxrnz/kXwj8PdT
# PFY/DWHOSQ==
# SIG # End signature block
