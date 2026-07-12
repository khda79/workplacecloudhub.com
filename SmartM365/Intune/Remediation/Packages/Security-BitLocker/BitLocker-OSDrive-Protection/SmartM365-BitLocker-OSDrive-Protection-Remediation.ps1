<#
.SYNOPSIS
    Version: 1.0
    Enables or repairs BitLocker protection for BitLocker-OSDrive-Protection.
.DESCRIPTION
    Validates TPM readiness, adds a recovery password protector when needed, enables BitLocker on the operating system volume with used-space-only encryption, resumes protection when suspended, and attempts to back up the recovery key to Entra ID.
#>
[CmdletBinding()]
param(
    [string]$MountPoint = 'C:'
)

$ErrorActionPreference = 'Stop'
$Scenario = 'BitLocker-OSDrive-Protection'
$LogRoot = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path $LogRoot "$Scenario-Remediation.log"
function Write-SmartM365Log { param([string]$Message) $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Scenario, $Message; Write-Output $line; Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8 }

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-SmartM365Log 'Remediation started.'

    $tpm = Get-Tpm -ErrorAction Stop
    if (-not $tpm.TpmPresent -or -not $tpm.TpmReady) { throw 'TPM is not present or not ready. BitLocker cannot be enabled safely by this remediation.' }

    $volume = Get-BitLockerVolume -MountPoint $MountPoint -ErrorAction Stop
    if ($volume.ProtectionStatus -eq 'On' -and $volume.EncryptionPercentage -eq 100) {
        Write-SmartM365Log 'BitLocker is already fully enabled and protected.'
        exit 0
    }

    $recoveryProtector = $volume.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } | Select-Object -First 1
    if (-not $recoveryProtector) {
        $protector = Add-BitLockerKeyProtector -MountPoint $MountPoint -RecoveryPasswordProtector -ErrorAction Stop
        $recoveryProtector = $protector.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } | Select-Object -First 1
        Write-SmartM365Log 'Recovery password protector added.'
    }

    if ($volume.VolumeStatus -eq 'FullyDecrypted') {
        Enable-BitLocker -MountPoint $MountPoint -TpmProtector -UsedSpaceOnly -SkipHardwareTest -ErrorAction Stop
        Write-SmartM365Log 'BitLocker enablement started with TPM protector and used-space-only encryption.'
    }
    elseif ($volume.ProtectionStatus -ne 'On') {
        Resume-BitLocker -MountPoint $MountPoint -ErrorAction SilentlyContinue
        Write-SmartM365Log 'BitLocker protection resume requested.'
    }

    if ($recoveryProtector) {
        try {
            BackupToAAD-BitLockerKeyProtector -MountPoint $MountPoint -KeyProtectorId $recoveryProtector.KeyProtectorId -ErrorAction Stop
            Write-SmartM365Log 'Recovery key backup to Entra ID requested.'
        }
        catch {
            Write-SmartM365Log "Recovery key backup to Entra ID failed or is not supported on this device: $($_.Exception.Message)"
        }
    }

    Write-SmartM365Log 'Remediation completed.'
    exit 0
}
catch {
    Write-SmartM365Log "Remediation failed: $($_.Exception.Message)"
    exit 1
}


# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDHvLLgy7053Gci
# zuzR25rQ1a4PO5wk7wODETiimHXCzqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCAubvEaXPFB5jzC+4TLIv45sWcHvnKJS7B5f2tiqiSluDANBgkqhkiG9w0B
# AQEFAASCAYCyIscq1FcLXOVVVPyfurY8Ex4dTPmzdSYqsSAiKSzu/y/ZvnIGitGo
# ulhBcHplopT8sEVHay/YeJDxW26NbigfUGWeEW6K6Q7fHmsLRe4XvzNmv2R2aVdA
# UvnCwzTv9MnYVNdkHJuhttsQNCLWKgWAha/HFIfEY2oqH4oMotr7zKATamtQ1+ur
# HaE7NpkVH2vcYeDJYz81UpwG+777khzmPcPvQcGKvvSdcKlQ3w4RqCMRj4MQmFoJ
# FGmU3b6RPbR/hY78yZKjVwcPNRPkDUefzKZfGfCDe5nCd93ZCg4UquPF55POEhM0
# jD8Djkv0FdO9/QAacd8Xk8FDewcr1mHVFgAM0udvLQEJfC9JWpI9vyNQcGEiy21l
# QYvS/JELxqPu7BmHfKiuHSz0QV2jlQAY7WNSuNhPoXQtzau6e0H2RFN6kPHE+OK6
# Bj4Sv4YE+aBTc8D0X/H9ILsXkiMsQls8fmBf3VNM/Ko3NFXe5ooz9q8uT9mEwIdy
# Noyx5kHAhbk=
# SIG # End signature block
