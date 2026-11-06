<#
.SYNOPSIS
    Offline tests of the Windows 11 upgrade decision for computers not enrolled in Intune.
.DESCRIPTION
    Extracts the real decision block of the Windows11Upgrade endpoint script (from the Intune
    gate comment to the end of the upgrade decision chain) and runs it with stubbed actions
    (disk cleanup, policy repair, setup, Windows Update). Checks that a computer not enrolled in
    Intune is upgraded by setup through the usual gates, and that NOT_INTUNE_ENROLLED is reported
    only when setup upgrade is disabled.
    Prints one PASS/FAIL line per assertion and "Summary: Passed=N; Failed=M". Exits 1 on failure.
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

$ErrorActionPreference = 'Stop'
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$w11 = Get-Content -LiteralPath (Join-Path $root 'Campaigns\Windows11Upgrade\Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1') -Raw
$passed = 0
$failed = 0
function Assert-Test([string]$Name, [bool]$Condition) {
    if ($Condition) { $script:passed++; Write-Host "PASS: $Name" }
    else { $script:failed++; Write-Host "FAIL: $Name" }
}

$start = $w11.IndexOf('    # Intune enrollment does not gate the setup upgrade')
$end = $w11.IndexOf("`r`n}`r`ncatch [System.OperationCanceledException]", $start)
if ($start -lt 0 -or $end -lt 0) { throw 'Windows 11 upgrade decision block not found.' }
$decision = [scriptblock]::Create($w11.Substring($start, $end - $start))

function Write-SmartLog { param($Message,$Level = 'INFO') [void]$script:logs.Add(('{0}|{1}' -f $Level,$Message)) }
function Invoke-SafeDiskCleanup { $script:calls['DiskCleanup']++; $script:DiskCleanupAction = 'SafeCleanup'; return $script:freeAfterCleanup }
function Invoke-OldUpgradeFolderCleanup { $script:calls['AdvancedCleanup']++; return $script:freeAfterCleanup }
function Invoke-DismComponentCleanup { $script:calls['DismCleanup']++; return $script:freeAfterCleanup }
function Repair-WindowsUpdatePolicyBlockers { $script:calls['PolicyRepair']++; [pscustomobject]@{ TimedOut = $false; StopDetail = ''; OutputFile = 'gpupdate.txt' } }
function Get-WindowsUpdatePolicySummary { [pscustomobject]@{ HasLegacyBlocker = $false; Issues = '' } }
function Resolve-PendingRebootOutcome { param($PendingRebootInfo,$RebootReason) [pscustomobject]@{ Status = 'PENDING_REBOOT_USER_CONNECTED'; NextAction = 'WAIT'; Detail = ''; ExitCode = 3; ActionResult = '' } }
function Resolve-SetupUpgradeExecutable { $script:calls['ResolveSetup']++; return 'C:\SetupMedia\setup.exe' }
function Invoke-SetupUpgrade { param($SetupExePath) $script:calls['Setup']++; return 0 }
function Get-SetupExitCodeInfo { param($ExitCode) [pscustomobject]@{ ExitCode = $ExitCode } }
function Format-SetupExitCodeInfo { param($Info) return ('ExitCode={0}' -f $Info.ExitCode) }
function Resolve-SetupUpgradeSuccessOutcome { param($Prefix,$SetupExitCode) [pscustomobject]@{ Status = "${Prefix}_STARTED"; NextAction = 'MONITOR_SETUP_AND_REBOOT'; ExitCode = 0; ActionResultSuffix = '' } }
function Save-SetupReturnCheckpoint { param($Status,$NextAction,$Detail,$ActionResult,$ExitCode,$SetupExePath) }
function Resolve-SetupFailureOutcome { param($SetupExitInfo) [pscustomobject]@{ Handled = $false } }
function Get-SetupFailureNextAction { param($Info) 'CHECK_SETUP_LOGS' }
function Reset-WindowsUpdateComponents { $script:calls['WUReset']++ }
function Invoke-AssignedUpdateInstall { $script:calls['ForceUpgrade']++; return 'Triggered' }

function Invoke-Decision {
    param(
        [bool]$Enrolled,
        [bool]$SetupUpgrade = $true,
        [bool]$Audit = $false,
        [int]$FreeGb = 50,
        [int]$FreeAfterCleanup = 50,
        [bool]$PolicyBlocker = $false
    )
    $script:logs = New-Object System.Collections.ArrayList
    $script:calls = @{ DiskCleanup = 0; AdvancedCleanup = 0; DismCleanup = 0; PolicyRepair = 0; ResolveSetup = 0; Setup = 0; WUReset = 0; ForceUpgrade = 0 }
    $script:freeAfterCleanup = $FreeAfterCleanup
    $script:DiskCleanupAction = ''; $script:AdvancedDiskCleanupAction = ''; $script:DismCleanupAction = ''
    $script:HardwareReadiness = [pscustomobject]@{ Result = 'CAPABLE'; Reason = ''; Logging = '' }
    $os = [pscustomobject]@{ MajorFamily = 'Windows10' }
    $hardwareReadinessBlocksUpgrade = $false
    $intune = [pscustomobject]@{ IsIntuneEnrolled = $Enrolled; StrongEvidence = ''; WeakEvidence = 'none' }
    $indicators = [pscustomobject]@{ ActionableBlocking = $false; BlockingReasons = '' }
    $policy = [pscustomobject]@{ HasLegacyBlocker = $PolicyBlocker; Issues = $(if ($PolicyBlocker) { 'TargetReleaseVersion' } else { '' }) }
    $freeGb = $FreeGb
    $MinimumFreeDiskGB = 32
    $pendingReboot = $false
    $pendingRebootInfo = [pscustomobject]@{ IsPending = $false; Source = '' }
    $AllowSetupUpgrade = $SetupUpgrade
    $DirectSetupUpgrade = $false
    $AuditOnly = $Audit
    $AllowDiskCleanup = $true
    $AllowAdvancedDiskCleanup = $false
    $AllowDismComponentCleanup = $false
    $AllowPolicyRepair = $true
    $AllowReboot = $true
    $ScheduleRetryAfterReboot = $true
    $AllowWUReset = $true
    $AllowForceUpgrade = $true
    $policyGpUpdateTimedOut = $false
    $policyRepairCompletedNormally = $false
    $status = 'UNKNOWN'; $nextAction = ''; $detail = ''; $actionResult = ''; $exitCode = 1
    . $decision
    return [pscustomobject]@{ Status = $status; NextAction = $nextAction; ExitCode = $exitCode; ActionResult = $actionResult; Calls = $script:calls.Clone(); Logs = @($script:logs) }
}

$warnPattern = '^WARN\|Device is not enrolled in Intune'

$r = Invoke-Decision -Enrolled $false
Assert-Test 'not enrolled + setup upgrade: setup.exe runs' ($r.Calls['Setup'] -eq 1 -and $r.Status -eq 'SETUP_UPGRADE_STARTED' -and $r.ExitCode -eq 0)
Assert-Test 'not enrolled + setup upgrade: no NOT_INTUNE_ENROLLED' ($r.Status -ne 'NOT_INTUNE_ENROLLED')
Assert-Test 'not enrolled + setup upgrade: warning logged once' (@($r.Logs | Where-Object { $_ -match $warnPattern }).Count -eq 1 -and @($r.Logs | Where-Object { $_ -match $warnPattern -and $_ -match 'continues through the usual gates' }).Count -eq 1)
Assert-Test 'not enrolled + setup upgrade: no Windows Update reset or forced update' ($r.Calls['WUReset'] -eq 0 -and $r.Calls['ForceUpgrade'] -eq 0)

$r = Invoke-Decision -Enrolled $true
Assert-Test 'enrolled + setup upgrade: setup.exe runs as before' ($r.Calls['Setup'] -eq 1 -and $r.Status -eq 'SETUP_UPGRADE_STARTED')
Assert-Test 'enrolled: no Intune warning' (@($r.Logs | Where-Object { $_ -match $warnPattern }).Count -eq 0)

$r = Invoke-Decision -Enrolled $false -SetupUpgrade $false
Assert-Test 'not enrolled + setup disabled: NOT_INTUNE_ENROLLED' ($r.Status -eq 'NOT_INTUNE_ENROLLED' -and $r.NextAction -eq 'FIX_INTUNE_ENROLLMENT_FIRST' -and $r.ExitCode -eq 3)
Assert-Test 'not enrolled + setup disabled: nothing runs' ($r.Calls['Setup'] -eq 0 -and $r.Calls['WUReset'] -eq 0 -and $r.Calls['ForceUpgrade'] -eq 0 -and $r.Calls['ResolveSetup'] -eq 0)
Assert-Test 'not enrolled + setup disabled: warning explains why' (@($r.Logs | Where-Object { $_ -match $warnPattern -and $_ -match 'setup upgrade is disabled' }).Count -eq 1)

$r = Invoke-Decision -Enrolled $false -SetupUpgrade $false -FreeGb 10 -FreeAfterCleanup 40
Assert-Test 'not enrolled + setup disabled: no disk cleanup (unchanged)' ($r.Calls['DiskCleanup'] -eq 0 -and $r.Status -eq 'NOT_INTUNE_ENROLLED')

$r = Invoke-Decision -Enrolled $false -FreeGb 10 -FreeAfterCleanup 40
Assert-Test 'not enrolled + low disk: safe disk cleanup runs, then setup' ($r.Calls['DiskCleanup'] -eq 1 -and $r.Calls['Setup'] -eq 1 -and $r.Status -eq 'SETUP_UPGRADE_STARTED')

$r = Invoke-Decision -Enrolled $false -FreeGb 10 -FreeAfterCleanup 20
Assert-Test 'not enrolled + disk still low after cleanup: INSUFFICIENT_DISK_AFTER_CLEANUP' ($r.Calls['DiskCleanup'] -eq 1 -and $r.Calls['Setup'] -eq 0 -and $r.Status -eq 'INSUFFICIENT_DISK_AFTER_CLEANUP')

$r = Invoke-Decision -Enrolled $false -PolicyBlocker $true
Assert-Test 'not enrolled + policy blocker: policy repair runs, setup waits for the next run' ($r.Calls['PolicyRepair'] -eq 1 -and $r.Calls['Setup'] -eq 0 -and $r.Status -eq 'WU_POLICY_BLOCKER' -and $r.ActionResult -eq 'PolicyRepairCompleted')

$r = Invoke-Decision -Enrolled $false -Audit $true
Assert-Test 'not enrolled + audit: SETUP_UPGRADE_READY without setup' ($r.Status -eq 'SETUP_UPGRADE_READY' -and $r.Calls['Setup'] -eq 0 -and $r.ActionResult -eq 'SetupValidatedAuditOnly')

$r = Invoke-Decision -Enrolled $false -Audit $true -FreeGb 10 -FreeAfterCleanup 40
Assert-Test 'not enrolled + audit + low disk: no cleanup, INSUFFICIENT_DISK' ($r.Calls['DiskCleanup'] -eq 0 -and $r.Status -eq 'INSUFFICIENT_DISK')

Write-Host "Summary: Passed=$passed; Failed=$failed"
if ($failed) { exit 1 }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBE14zFin2wrfWS
# nhiDD8SVrSJ9XL+R8PEvzXY25FwcH6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJUt10feBFaLGc5cLiOJ4RjkDnIZh5KP1DsSdSG5jdgZMA0GCSqG
# SIb3DQEBAQUABIIBgA6wueRig4Yr5cy4LBOlac1Mk4QV+c+MmcTubJRRcFZQbLaw
# lZoXKucNzbkp8hm2CYBF+2WdhIkJoPM8ym6FcpYNPC5AkCo7ohXvTXAXLd7A7VYV
# 39tiIuYasWotfVA0nS6WLqejSRikiz09eociElqz+TLymLt4hM3K/LIjgs9K8/zG
# ZKiFc67piJvawmB6xODGDZJWJZMiv4C5dMTacfEX9SlYhLZPKyZIwY1EQW14wUyG
# V0h24M0GwRPGGoHlsVkUw0KbIPwjdClkCst8h7nJDNsoe2epxnDL03IKc49Kqti+
# BLnXsjVyEHIzkcOayRM6KCVcYo9rI0bzrUAHKK31Ua35xC6gy2BEl1G0goxtEncS
# 1lxy9h9e/08iC9Wdtve4e6cq966D9WPE5Sn1RRzbunD5Zi/JA+ZUen8E98zP31Ev
# pGvNeKAZdwvC3ZRmKDyK/H0FO1xMwkQH7LvbpfvDVzC9OKkzJX977rn7KY4HA5u2
# XF0tFuRzm7lMq2E11aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDYwOTAx
# NDZaMC8GCSqGSIb3DQEJBDEiBCBJoga+PdG/h9JiPUYuS//L4E2s9jEvga9vBlaY
# qkAlnTANBgkqhkiG9w0BAQEFAASCAgAJMHFnv8H3hZ5BfdzHHVL597k925K2r+Se
# ffMXNYHvEGPKunHwX7U7wJOGPNLjiG+Hh1+qN1SvKakANYOWAO++AAkLRSrh+0nD
# YbnCySNVpx73hYClo91kxZ1dVEbY3xSozTexnclv7cnXitC2sSbInNJ3HNQU5GyQ
# Qcg87EHPa89nxnb1Au8mOKr+BeM4edT/p5GjuAHgX4fcICzX0R/orTRdQ6o1FBvX
# 0Z2YazUXqnqlxXSNPtHnFTm2Zr8SVAx5YR5KkmOm0rMEWEivz4l16nh7/rPis3HP
# jLaR3Kr7SMRIayX7dTOgbuEmpdRVU/L3cl8+o8mJ8s9y0og9akkxEy3Wt/bbox6i
# 4pI0ADA0F2ynbMipOO6vnAANALsg9A4rfkVvflp8cj2DGsR/pYY6xZWf7Tpd1Pqz
# 993HOSg+i+xQ4qOYpPWA51/5zE+0zeERBL5NOOron27p3XYOwbUO91u5PTIWFtIj
# zWb2Pd7MagbprTL0M2TPCPZw8SFgNUZI1GvFviFFssvO6jL6Y5auGxX52jvugdC6
# BmVyRjgOI2C1ZC0xQJ2/huAuF17HZ1UJxSYz6lhDHmCaxp5cWrwkKmGa3bJ7JP8H
# 8HuhpLgnkppODSg9fRdwTRtH9SjEHxUPUaXMJG3LD9ork58WqldtUUhKQ5xJjF2C
# PgZ2jWgNOg==
# SIG # End signature block
