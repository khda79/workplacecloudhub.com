<#
.SYNOPSIS
    Regression test: a Single PC launch from the GUI must apply the GUI options like a LOT launch.
.DESCRIPTION
    The GUI passes its options as EHJIR_* environment values. A LOT launch goes through
    Run-IntuneHybridJoinRepairWithPsExec-Lot.cmd, which turns them into launcher parameters.
    A Single PC launch called the PsExec launcher directly with the environment only, which the
    launcher does not read: unchecked "Allow dsregcmd /leave" still ran with the launcher default
    (-AllowDsregLeave = $true) and "Skip virtual machines", reboot, cleanup, delay, night pause
    and timeout choices were ignored.
    This test extracts Start-ToolkitSingleComputer from the GUI with stubbed helpers and checks
    the generated launcher command line. Nothing is launched.
    Prints one PASS/FAIL line per assertion and "Summary: Passed=N; Failed=M". Exits 1 on failure.
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:Passed = 0
$script:Failed = 0
$toolkitRoot = Split-Path -Path $PSScriptRoot -Parent
$guiPath = Join-Path $toolkitRoot 'Scripts\SmartM365-IntuneHybridJoinRepair-LotLauncher-GUI.ps1'

function Write-TestResult {
    param([bool]$Condition, [string]$Message, [string]$Detail = '')
    if ($Condition) { $script:Passed++; Write-Host ("PASS: {0}" -f $Message) }
    else {
        $script:Failed++
        if ([string]::IsNullOrWhiteSpace($Detail)) { Write-Host ("FAIL: {0}" -f $Message) } else { Write-Host ("FAIL: {0} ({1})" -f $Message,$Detail) }
    }
}

$tokens = $null; $errors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$tokens, [ref]$errors)
Write-TestResult -Condition (@($errors).Count -eq 0) -Message 'GUI parses without error'
foreach ($name in @('Start-ToolkitSingleComputer','Assert-CmdLiteralValue','ConvertTo-CmdArgument','ConvertTo-CmdSetCommand')) {
    $definition = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
    if ($definition) { . ([scriptblock]::Create($definition.Extent.Text)) }
}
$mapping = $ast.Find({ param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-SingleComputerOptionArguments' }, $true)
if ($mapping) { . ([scriptblock]::Create($mapping.Extent.Text)) }

$tempRoot = Join-Path ([IO.Path]::GetTempPath()) ('EHJIR-SinglePcTest-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path (Join-Path $tempRoot 'Scripts') -Force | Out-Null
Set-Content -LiteralPath (Join-Path $tempRoot 'Scripts\SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1') -Value '# stub' -Encoding ASCII
$script:CapturedCommands = @()
function New-SingleComputerRunContext { param([string]$ToolkitRoot, [string]$ComputerName) [pscustomobject]@{ Root = $tempRoot; ComputersPath = (Join-Path $tempRoot 'Computers.txt'); ComputerName = $ComputerName; Name = $ComputerName; LogRoot = $tempRoot; ReportRoot = $tempRoot; CentralLogRoot = $tempRoot; ArchiveRoot = $tempRoot } }
function Resolve-GuiPsExecPath { param([string]$ToolkitRoot) return 'C:\Tools\PsExec64.exe' }
function Resolve-GuiPowerShellPath { return 'C:\Program Files\PowerShell\7\pwsh.exe' }
function New-GuiLaunchCommandFile { param([string]$WorkingDirectory, [string[]]$Commands, [string]$NamePrefix, [string]$WindowTitle, [switch]$PauseWhenDone) $script:CapturedCommands = @($Commands); return (Join-Path $tempRoot 'launch.cmd') }
function Start-GuiLaunchCommandFile { param([string]$LaunchCommandPath, [string]$WorkingDirectory) }

function Get-LaunchLine {
    param([hashtable]$Environment, [string[]]$ExtraArguments = @(), [string]$Mode = 'Once')
    $script:CapturedCommands = @()
    $fictionalComputer = 'PC-0142'
    [void](Start-ToolkitSingleComputer -ToolkitRoot $tempRoot -ComputerName $fictionalComputer -Mode $Mode -ExtraArguments $ExtraArguments -Environment $Environment)
    return [string](@($script:CapturedCommands | Where-Object { $_ -like '*SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1*' }) | Select-Object -Last 1)
}

$guiOff = @{
    EHJIR_THROTTLE = '10'; EHJIR_GLOBAL_CONCURRENCY_LIMIT = '15'; EHJIR_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES = '0'
    EHJIR_DELAY_BETWEEN_CYCLES_MINUTES = '7'; EHJIR_DISABLE_NIGHT_PAUSE = '1'; EHJIR_NIGHT_PAUSE_START_HOUR = '21'; EHJIR_NIGHT_PAUSE_END_HOUR = '6'
    EHJIR_INTUNE_RETRY_SLEEP_MINUTES = '4'; EHJIR_INTUNE_RETRY_MAX_RETRIES = '2'; EHJIR_STALE_CLEANUP_DELAY_SECONDS = '45'; EHJIR_REBOOT_DELAY_SECONDS = '240'; EHJIR_PSEXEC_TIMEOUT_MINUTES = '90'
    EHJIR_ALLOW_DSREG_LEAVE = '0'; EHJIR_ALLOW_REMOVE_STALE_INTUNE_ENROLLMENT = '0'; EHJIR_ALLOW_REBOOT_WHEN_NO_INTERACTIVE_USER = '0'; EHJIR_ALLOW_REBOOT_AFTER_DSREG_LEAVE = '0'; EHJIR_SKIP_VIRTUAL_MACHINES = '1'
    EHJIR_USE_TECHNICIAN_RUN_GUARD_HISTORY = '1'; EHJIR_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY = '0'; EHJIR_TECHNICIAN_RUN_GUARD_HOURS = '12'; EHJIR_CENTRAL_LOG_COLLECTION_MODE = 'Standard'
}
try {
    $line = Get-LaunchLine -Environment $guiOff
    # The generated CMD quotes this token; PowerShell 7 -File binds "-AllowDsregLeave:$false" as false either way.
    Write-TestResult -Condition ($line -match '(^| )"?-AllowDsregLeave:\$false"?( |$)') -Message 'unchecked "Allow dsregcmd /leave" passes -AllowDsregLeave:$false' -Detail $line
    Write-TestResult -Condition ($line -notmatch '-AllowRemoveStaleIntuneEnrollment|-AllowRebootWhenNoInteractiveUser|-AllowRebootAfterDsregLeave') -Message 'unchecked cleanup and reboot authorizations are not passed'
    Write-TestResult -Condition ($line -match '(^| )-SkipVirtualMachines( |$)') -Message 'checked "Skip virtual machines" passes -SkipVirtualMachines'
    Write-TestResult -Condition ($line -match '(^| )-DisableNightPause( |$)' -and $line -match '-NightPauseStartHour 21' -and $line -match '-NightPauseEndHour 6') -Message 'night pause choice and hours are passed'
    Write-TestResult -Condition ($line -match '-DelayBetweenCyclesMinutes 7' -and $line -match '-IntuneRetrySleepMinutes 4' -and $line -match '-IntuneRetryMaxRetries 2' -and $line -match '-StaleCleanupDelaySeconds 45' -and $line -match '-RebootDelaySeconds 240' -and $line -match '-PsExecTimeoutMinutes 90') -Message 'delay, retry, cleanup, reboot delay and PsExec timeout values are passed'
    Write-TestResult -Condition ($line -match '-ThrottleLimit 1( |$)' -and $line -notmatch '-ThrottleLimit 10') -Message 'Single PC keeps one worker'

    $guiOn = $guiOff.Clone()
    $guiOn.EHJIR_ALLOW_DSREG_LEAVE = '1'; $guiOn.EHJIR_ALLOW_REMOVE_STALE_INTUNE_ENROLLMENT = '1'; $guiOn.EHJIR_ALLOW_REBOOT_WHEN_NO_INTERACTIVE_USER = '1'; $guiOn.EHJIR_ALLOW_REBOOT_AFTER_DSREG_LEAVE = '1'; $guiOn.EHJIR_SKIP_VIRTUAL_MACHINES = '0'; $guiOn.EHJIR_DISABLE_NIGHT_PAUSE = '0'
    $line = Get-LaunchLine -Environment $guiOn
    Write-TestResult -Condition ($line -match '(^| )-AllowDsregLeave( |$)' -and $line -match '-AllowRemoveStaleIntuneEnrollment' -and $line -match '-AllowRebootWhenNoInteractiveUser' -and $line -match '-AllowRebootAfterDsregLeave') -Message 'checked authorizations are passed'
    Write-TestResult -Condition ($line -notmatch '-SkipVirtualMachines|-DisableNightPause') -Message 'unchecked "Skip virtual machines" and enabled night pause are not passed'

    $line = Get-LaunchLine -Environment $guiOff -ExtraArguments @('-DryRun','-AuditOnly') -Mode 'OnceIgnoreRunGuard'
    Write-TestResult -Condition ($line -match '-RunOnce' -and $line -match '-IgnoreRunGuard' -and $line -match '-DryRun' -and $line -match '-AuditOnly') -Message 'mode and extra GUI arguments are still passed'
}
finally {
    Remove-Item -LiteralPath $tempRoot -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ("Summary: Passed={0}; Failed={1}" -f $script:Passed,$script:Failed)
if ($script:Failed -gt 0) { exit 1 }
exit 0

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC0L/3tdcpoTK9t
# buZWDlEyaJVuyr/Yudr/ii1jFdpodqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFqEHfocALVjECNqLleCHzwAXGNwAYdOIfupNEVPL2HxMA0GCSqG
# SIb3DQEBAQUABIIBgGU5BWCtOgxwQ8PqTM7E39dB60HyKReR74ztG9xTbJky/FpE
# AZYR0/qh9civ7h4m8VXn3tYioUGDU5iBPEaj/JAR1VWVnSfN99g/sNgBkpkGXB/u
# fiCDTixGaiJ1muV3n9ISvK7smpuZMkqb4+ckPUA7otRxT4x6jQQlvjLNC96HqFmD
# y/l1E0QKiUBPLDuuCzTfUTv97EJHYQ9Fv91uPLAYb1Fqfm45COUGiUpnEs+a4JTT
# 2O1tW7Ie8h1vJsXZ2b8sg1bdEsq0/WR0BUBYQFvM9eHcmo69yfNN2d/jRdkwXGx6
# CAGp6qHrS91rtZ2fGOoj3CbPhwWLQRGKn1327g4zUnzm3RDFR4zKvcwD7JE+11m+
# f1sORSMCSaYUxfxmsc9F3BuiOkq1v3wto4DLiWF0vX4Pw7Ha/u6TwZlOG7NBgkEJ
# USEMht3XoJXH32xXtt8GXj06Z2dMcMxXpgk3MdyFC+jqlEl/5f1wpocYZ9R4/Ac7
# iKkLW0iZ1QMkMBg5baGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMzE2
# NDZaMC8GCSqGSIb3DQEJBDEiBCBguKP0Dr1eCI8H12GljoVbXSQuRpNrsZ4qJqFX
# REU3oDANBgkqhkiG9w0BAQEFAASCAgA5WD7OHhwG+2ICOcH5O3FcPz5mKFCIXrne
# L3xLkpW4xzfsNE/txAnm4mORGzNrICxuXjL7pD3olb6K7yFyAXmJRh3f9uoaP6Ko
# DOIbZjb95L286QSbT9MyLoP5yABRHnOMrTvgTPmR0ibhUfY8ICFpb/SsjZnRa+Hb
# ehBdtzf2DIYv/Mb8m1m/rdddXXk+BzpVzOeWCIvd8SlPsjoKfgHpOMrT9evzmQjQ
# H/tyKyHo9Z5Xf9ytRwhpd1QPnc4etWgq0JbWJInUoif9ZLgY7AcDZPBAMoKn6S9P
# 7ZQvKHvWEg6Gl8CW9wSjLPT8ovxATjuzFoYUOMaBEjpEkt5q8lJsjLgZfULdXyhE
# +RJr9sygf8iMNyCjCAggFzyBdFGaRbaSm13rbueRRvhRTiWot20DmHuDB5+CSX+Y
# 63R+7od8gN1tw57LDorhB2Y6CAO3VYqwzBsUTY5vxgO4/OvGBTpZsq1lNODyLY1r
# SfR87eJE8M+Wi6n+BHSdGkY9+st8kcD6yF7a1xj2DdZo53SgeoDMzv23i1TuMRjn
# GVDe75jSQxgpqKw5hNsnUgiVnoe5lS6frp2gYTr4lIscqxW7L7pObdx5yeMmOfss
# ten6+bJ3sBLDC2v5wcZrKmnIgCpqUm+xbR0rPEref6JAKnw2KLXkvHDAC106PS2p
# r65X2v1E/g==
# SIG # End signature block
