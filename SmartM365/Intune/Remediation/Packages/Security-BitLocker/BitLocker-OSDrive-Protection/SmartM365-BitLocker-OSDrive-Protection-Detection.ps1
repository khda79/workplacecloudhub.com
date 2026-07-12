<#
.SYNOPSIS
    Version: 1.0
    Detects BitLocker protection health on the operating system drive.
.DESCRIPTION
    Validates that the operating system drive exists, BitLocker protection is enabled,
    encryption is complete, and a recovery password protector is present.
#>
[CmdletBinding()]
param(
    [string]$MountPoint = 'C:'
)

$ErrorActionPreference = 'Stop'

try {
    $volume = Get-BitLockerVolume -MountPoint $MountPoint -ErrorAction Stop

    if ($null -eq $volume) {
        Write-Output "Status=BitLockerVolumeNotFound"
        exit 1
    }

    $issues = New-Object System.Collections.Generic.List[string]

    if ($volume.ProtectionStatus -ne 'On') {
        $issues.Add("ProtectionStatus=$($volume.ProtectionStatus)")
    }

    if ($volume.EncryptionPercentage -lt 100) {
        $issues.Add("EncryptionPercentage=$($volume.EncryptionPercentage)")
    }

    $recoveryProtector = $volume.KeyProtector |
        Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } |
        Select-Object -First 1

    if (-not $recoveryProtector) {
        $issues.Add('RecoveryPasswordProtector=Missing')
    }

    if ($issues.Count -gt 0) {
        Write-Output "Status=BitLockerNonCompliant"
        Write-Output ("Issues=" + ($issues -join '; '))
        exit 1
    }

    Write-Output "Status=BitLockerProtected"
    Write-Output "MountPoint=$MountPoint"
    Write-Output "EncryptionPercentage=$($volume.EncryptionPercentage)"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAF5xjl8uFnFiTR
# Alyj6l5+o7Qt/LtGHwpoHxgNZTbhhqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDe0oaFoW8gpPOP0M/L5TXUFLbdeNa063U2NBMuHn59BjANBgkqhkiG9w0B
# AQEFAASCAYAXyoeaRqInmX5wk8V7v9Up2hg1WwaClJG5f76BsjrRsoO15bM95ELa
# mVDyd8IiCs25vT8rcUP19CYojJcHUACPg2fgTJ3vGL4ntCEF1z7bj7V7IB+b9yO+
# 9ycwRn9YunMmRf7w8IurACAnxlmqflopUPXtu6B24hrCW4Y8/LzU1eMPiWOdyC03
# f8Hbg2CFnGhgiNaWKqz9/jt4zk1aRUo6EhAAlBksWlxOqxpoDWZZcySMKQ6P7/2V
# j8PYHDSavxTovIXGvJM+hXLWvQ3K62abImOhHkRdDAfIdgmlCkxg46xfgZ1B7eUs
# qRF0v7MVOgGel5zRMd224t9eCpA/qgy+1aIBGbkutxnaiwJoMAthH2p85EstVOHU
# JRSWfA6bWJh3Y12aV2MFvtXO/1beSrdiIcwZYOc5gOX0xtFX+P2BEdSnB/iF+X4/
# bB9tBkT32OYcHzcMDhbDSiLcO9l5J04XL5MQgp7XiWOcJGHbuaNLO6W0X1X+BA5S
# vT2jyIxu0Pk=
# SIG # End signature block
