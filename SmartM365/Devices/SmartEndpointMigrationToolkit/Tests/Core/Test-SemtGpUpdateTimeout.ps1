<#
.SYNOPSIS
    Offline gpupdate timeout tests using simulated processes.
.VERSION
    1.0.0
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$hj = Get-Content -LiteralPath (Join-Path $root 'Campaigns\HybridJoin\Scripts\SmartM365-Invoke-EndpointMigration-HybridJoin.ps1') -Raw
$w11 = Get-Content -LiteralPath (Join-Path $root 'Campaigns\Windows11Upgrade\Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1') -Raw
$passed = 0
$failed = 0
function Assert-Test([string]$Name, [bool]$Condition) {
    if ($Condition) { $script:passed++; Write-Host "PASS: $Name" }
    else { $script:failed++; Write-Host "FAIL: $Name" }
}
function Import-Function([string]$Source, [string]$Name) {
    $tokens = $null; $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseInput($Source,[ref]$tokens,[ref]$errors)
    if ($errors.Count) { throw ('Parse error in endpoint: ' + $errors[0].Message) }
    $f = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name },$true)
    if (-not $f) { throw "Missing function $Name" }
    return $f.Extent.Text
}
. ([scriptblock]::Create((Import-Function $hj 'Invoke-ProcessWithTimeout')))
. ([scriptblock]::Create((Import-Function $w11 'Invoke-SemtGpUpdateWithTimeout')))
function Start-Process {
    param($FilePath,$ArgumentList,[switch]$PassThru,$WindowStyle,$RedirectStandardOutput,$RedirectStandardError)
    Set-Content -LiteralPath $RedirectStandardOutput -Value 'gpupdate output'
    Set-Content -LiteralPath $RedirectStandardError -Value 'diagnostic'
    $script:lastCall = @{ FilePath=$FilePath; Arguments=@($ArgumentList) }
    $p = [pscustomobject]@{ ExitCode=7; Killed=$false; WaitCount=0 }
    $p | Add-Member ScriptMethod WaitForExit { param($ms) $this.WaitCount++; if ($this.WaitCount -eq 1) { return -not $script:simulateTimeout }; return $true }
    $p | Add-Member ScriptMethod Kill { if ($script:denyKill) { throw 'termination denied' }; $this.Killed=$true }
    $p | Add-Member ScriptMethod Refresh { }
    $script:lastProcess = $p
    return $p
}
$temp = Join-Path $env:TEMP ('SEMT-test-gpupdate-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temp | Out-Null
try {
    foreach ($campaign in @('HybridJoin','Windows11')) {
        foreach ($timeout in @($false,$true)) {
            $script:simulateTimeout = $timeout
            $script:denyKill = $false
            $outputFile = Join-Path $temp ($campaign + $timeout + '.txt')
            $result = if ($campaign -eq 'HybridJoin') {
                Invoke-ProcessWithTimeout -FilePath 'gpupdate.exe' -Arguments @('/target:computer','/force') -TimeoutSeconds 1 -OutputFile $outputFile
            } else { Invoke-SemtGpUpdateWithTimeout -TimeoutSeconds 1 -OutputFile $outputFile }
            Assert-Test "$campaign timeout=$timeout preserves output and arguments" ($result.TimedOut -eq $timeout -and (Get-Content -LiteralPath $outputFile -Raw) -match 'gpupdate output' -and $script:lastCall.FilePath -eq 'gpupdate.exe' -and ($script:lastCall.Arguments -join ' ') -eq '/target:computer /force')
            Assert-Test "$campaign timeout=$timeout handles process" ($(if ($timeout) { $script:lastProcess.Killed -and $result.ExitCode -eq '' } else { -not $script:lastProcess.Killed -and $result.ExitCode -eq 7 }))
        }
        $script:simulateTimeout = $true
        $script:denyKill = $true
        $outputFile = Join-Path $temp ($campaign + '-denied.txt')
        $result = if ($campaign -eq 'HybridJoin') {
            Invoke-ProcessWithTimeout -FilePath 'gpupdate.exe' -Arguments @('/target:computer','/force') -TimeoutSeconds 1 -OutputFile $outputFile
        } else { Invoke-SemtGpUpdateWithTimeout -TimeoutSeconds 1 -OutputFile $outputFile }
        Assert-Test "$campaign reports a failed termination" ($result.TimedOut -and $result.StopDetail -match 'termination denied')
    }
    # Execute the real initial policy-refresh branch with a fake process. The next
    # statement represents continuation toward the identity and MDM checks.
    $start = $hj.IndexOf('    if ($AuditOnly) {', $hj.IndexOf('    $domainPreflightComplete = $true'))
    $end = $hj.IndexOf('    if ($domainPreflightComplete) {', $start)
    if ($start -lt 0 -or $end -lt 0) { throw 'Initial policy-refresh branch not found.' }
    $policyBranch = [scriptblock]::Create($hj.Substring($start, $end - $start))
    function Write-RunLog { param($Message) }
    function Invoke-ProcessWithTimeout {
        param($FilePath,$Arguments,$TimeoutSeconds,$OutputFile)
        $script:gpCalls++
        return [pscustomobject]@{ ExitCode=0; Output=''; TimedOut=$script:mockTimeout; StopDetail='' }
    }
    $OutputDir = $temp; $ComputerName = 'PC'; $RunId = 'test'; $GpUpdateTimeoutSeconds = 180
    $GpUpdateWaitSeconds = 0; $gpUpdateTimedOut = $false; $gpUpdateAttempted = $false
    $script:gpCalls = 0; $script:mockTimeout = $false; $AuditOnly = $true; $continued = $false
    . $policyBranch
    $continued = $true
    Assert-Test 'Hybrid Join audit skips gpupdate and continues to identity checks' ($script:gpCalls -eq 0 -and -not $gpUpdateAttempted -and $continued)
    $AuditOnly = $false; $gpUpdateAttempted = $false
    . $policyBranch
    Assert-Test 'Hybrid Join repair still runs gpupdate' ($script:gpCalls -eq 1 -and $gpUpdateAttempted)
    Assert-Test 'Hybrid Join final gpupdate also excludes audit' ($hj -match '(?m)^\s*if \(-not \$AuditOnly -and -not \$gpUpdateTimedOut -and \$finalGpUpdateReasons\.Count -gt 0\)')
    $script:mockTimeout = $true; $gpUpdateTimedOut = $false; $status = 'SUCCESS'; $ExitCode = 0
    . $policyBranch
    Assert-Test 'Hybrid Join continues after initial gpupdate timeout without changing enrollment status' ($script:gpCalls -eq 2 -and $gpUpdateTimedOut -and $status -eq 'SUCCESS' -and $ExitCode -eq 0)
    $startFinal = $hj.IndexOf('    $finalGpUpdateReasons = New-Object')
    $endFinal = $hj.IndexOf('    # Enrollment events describe observations', $startFinal)
    if ($startFinal -lt 0 -or $endFinal -lt 0) { throw 'Final policy-refresh branch not found.' }
    $finalBranch = [scriptblock]::Create($hj.Substring($startFinal, $endFinal - $startFinal))
    $leaveAttempted = $true; $nonIntuneMdmRemovalAttempted = $false; $gpUpdateMdmPolicyWarning = $false
    $intuneEnrolled = $false; $rebootAttempted = $false; $status = 'INTUNE_PENDING'; $ExitCode = 3
    $finalGpUpdateTimedOut = $false; $AuditOnly = $false; $gpUpdateTimedOut = $false
    . $finalBranch
    Assert-Test 'Final gpupdate timeout preserves prior enrollment status' ($finalGpUpdateTimedOut -and $status -eq 'INTUNE_PENDING' -and $ExitCode -eq 3)
    $priorCalls = $script:gpCalls; $gpUpdateTimedOut = $true
    . $finalBranch
    Assert-Test 'No second gpupdate after an initial timeout' ($script:gpCalls -eq $priorCalls)
    Assert-Test 'Hybrid Join records both timeouts and preserves the final enrollment status' ($hj -match 'final gpupdate exceeded' -and $hj -notmatch "status = 'FINAL_GPUPDATE_TIMEOUT'" -and $hj -match 'GpUpdateTimedOut' -and $hj -match 'FinalGpUpdateTimedOut')
    # Exercise the Windows 11 decision before the upgrade gate with simulated policy repair.
    $startW11 = $w11.IndexOf('    # Repair the existing policy gate before choosing the upgrade path.')
    $endW11 = $w11.IndexOf("    if (`$os.MajorFamily -eq 'Windows11') {", $startW11)
    if ($startW11 -lt 0 -or $endW11 -lt 0) { throw 'Windows 11 policy preflight not found.' }
    $w11Branch = [scriptblock]::Create($w11.Substring($startW11, $endW11 - $startW11))
    function Write-SmartLog { param($Message,$Level) }
    function Repair-WindowsUpdatePolicyBlockers { $script:repairCalls++; [pscustomobject]@{TimedOut=$true;StopDetail='';OutputFile='test.txt'} }
    function Get-WindowsUpdatePolicySummary { [pscustomobject]@{HasLegacyBlocker=$script:blockerRemains;Issues='test'} }
    $os=[pscustomobject]@{MajorFamily='Windows10'}; $hardwareReadinessBlocksUpgrade=$false
    $DirectSetupUpgrade=$false; $intune=[pscustomobject]@{IsIntuneEnrolled=$true}
    $indicators=[pscustomobject]@{ActionableBlocking=$false}; $freeGb=50; $MinimumFreeDiskGB=10
    $pendingReboot=$false; $policy=[pscustomobject]@{HasLegacyBlocker=$true;Issues='test'}
    $AllowPolicyRepair=$true; $AuditOnly=$false; $policyGpUpdateTimedOut=$false
    $policyRepairCompletedNormally=$false; $script:repairCalls=0; $script:blockerRemains=$false
    . $w11Branch
    Assert-Test 'Windows 11 rechecks policy and can continue after gpupdate timeout' ($script:repairCalls -eq 1 -and $policyGpUpdateTimedOut -and -not $policy.HasLegacyBlocker)
    $policy=[pscustomobject]@{HasLegacyBlocker=$true;Issues='test'}; $script:blockerRemains=$true
    . $w11Branch
    Assert-Test 'Windows 11 retains a real policy blocker after gpupdate timeout' ($policy.HasLegacyBlocker -and $policyGpUpdateTimedOut)
    Assert-Test 'Windows 11 does not replace upgrade result with gpupdate timeout' ($w11 -notmatch "status = 'GPUPDATE_TIMEOUT'" -and $w11 -match 'PolicyGpUpdateTimedOut')

    # Real process: Windows PowerShell 5.1 loses ExitCode after WaitForExit(timeout) unless a handle is open.
    Remove-Item -LiteralPath Function:\Start-Process
    . ([scriptblock]::Create((Import-Function $hj 'Invoke-ProcessWithTimeout')))
    $real = Invoke-ProcessWithTimeout -FilePath (Join-Path $env:WINDIR 'System32\cmd.exe') -Arguments @('/c','exit','3') -TimeoutSeconds 60 -OutputFile (Join-Path $temp 'real-exit.txt')
    Assert-Test 'Hybrid Join process helper returns the real exit code (EXECUTED cmd.exe /c exit 3)' (-not $real.TimedOut -and [string]$real.ExitCode -eq '3')
    $handleLine = [regex]::Escape('$null = $process.Handle')
    Assert-Test 'Windows 11 gpupdate helper opens the process handle before waiting' ((Import-Function $w11 'Invoke-SemtGpUpdateWithTimeout') -match $handleLine)
} finally { Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue }
Write-Host "Summary: Passed=$passed; Failed=$failed"
if ($failed) { exit 1 }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA8pA3GsxITb6fI
# dUypSjebUNGavUCcnC30bg4L/Yz9lqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIK05Z2qFXtGwIkDNuZgCMws5mtoGr+yvIIrIuo9daDv+MA0GCSqG
# SIb3DQEBAQUABIIBgFhkVn1v2NbXLupydCnNWxOvUap1Pciylw3B9as6LVbxfQe+
# NxnwM4gq9fZMFHoIWynOwLvb9n/GXFcano2TwKRLVgTOXUiDmw3f5fRzazB+yGhb
# 4IQRrf1jVdldkhc5Ib3rEs0WzhMwAxgTPimsUghU0CAZNAwYCwBI6b2WkwIpI2pD
# 0NW2LwoySMY3ZbWqS36/1L+bShlZIkkr8PV7i4B+WxV8dHZwBDZQgsxuOB31MCnK
# KaXgnQfasKR1gY107LxwMtrU+5nJ7YrrUZN79u5IFNLw49NJtHhiUC1fckUnYwVs
# 9KfRq24zUXesBAwPbLTJ1RTGZZaLxTExlX08Q3/5qig/4aeJHG+I8SPMVVK1E4bb
# vzbvSUMfqt/w3mRlwXreDMG8iE1D+V5YXlnNhxbaFIOPkQVSGIYhLvjhb5TpNGQl
# D4B9MRX8Ru3OjbjK+b82Cwtsdn2QYisiUlqTcTTAhpgyXoDRm8KdcdpWKn1GbE4f
# 3CD9vbWUpw5bW9otgqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDUyMTQ4
# MTZaMC8GCSqGSIb3DQEJBDEiBCCzGRR3xhL4mqxZseGF+WzTW5DLs3yWviScWGcP
# 9Ab59DANBgkqhkiG9w0BAQEFAASCAgARqkPiq0Q1yECVJKOoBU6jf+Si4EAz9cIy
# t5BE61YxgKTbD+B78yWYjGMxdNrqnMWEJtahzsRzoU/9iylFsRIvzVTCxtwZWHVq
# vUzf6m29SEFo4JRR4Wz3wLjDJ/qdlSYCfcrOVP8EaQB+5bdZIJpj6zrgxWYCDMrM
# RVay0OIO54T8VXd+PbYlDlUieERb09ORA2bCT/Wny4+sGG9vC3neGmJgHL2/N6jC
# CqlOJUX6lBdAHTjkKt/9GlrIG/mZNkAyRdabdCGCy7bQt8fYIvoiUT+6fhW8AbrS
# zjFD7A3t6GOXIGxqX1z2E4o7odVOwdY1ksnu0yigRH9tj8FV69xLoW0llsLqUHZJ
# tUUy+EH7OzmPvAxZ/ngNf1qTFbTVRGIEtdjMlhpLFOpZOBQDoOwBmUBZe88vNfPG
# +f4LqDrrUu64DNuetb7tKSUlWTMDf2TxaMnT2Ug8IhjzGMg5fFa5MfARTfqROK3q
# 1a2oLR9Vrnxo9vjQ+x5yj36aTrozSWeIXYRI7zHRPxthF/OrPIpKdONkZmKrj3F/
# qKd7Y2F2ag6+q0t/shnxrWyO07hN214q8A0eZZ/w7cCyLbtsJZ9Xt2GMxAdRVAtv
# JvptI0z0lCAryVdS66YrquC9dpcZVoUNykJrwrKbagAWwUqXiScJGM3u0dAe0SYr
# gQR64BHTZQ==
# SIG # End signature block
