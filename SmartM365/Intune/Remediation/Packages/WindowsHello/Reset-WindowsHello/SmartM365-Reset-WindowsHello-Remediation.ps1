
<#
.SYNOPSIS
    Version: 1.0
    Intune remediation script for Windows Hello Key Trust issue.
.DESCRIPTION
    This script:
    - Logs actions to C:\ProgramData\SmartM365\IntuneRemediation\Logs\Remediate-WindowsHello.log.
    - Stops NgcSvc service.
    - Takes ownership of NGC folder and deletes it.
    - Restarts NgcSvc.
    - Optionally forces dsregcmd /join if Hybrid Join is broken.
    - Runs silently and exits with code 0 for success.
.NOTES
    Run as SYSTEM via Intune.
#>

$ErrorActionPreference = "Stop"
$Scenario = "Reset-WindowsHello"
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogPath = Join-Path -Path $LogRoot -ChildPath "$Scenario-Remediation.log"

function Write-SmartM365Log {
    param([string]$Message)
    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Scenario, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogPath -Value $line -Encoding utf8
}

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-SmartM365Log "RemediationStarted"

    $ngcPath = "C:\\Windows\\ServiceProfiles\\LocalService\\AppData\\Local\\Microsoft\\NGC"
    $ngcService = Get-Service -Name NgcSvc -ErrorAction SilentlyContinue

    if ($ngcService -and $ngcService.Status -ne "Stopped") {
        Write-SmartM365Log "Stopping NgcSvc service."
        Stop-Service -Name NgcSvc -Force -ErrorAction SilentlyContinue
    }

    if (Test-Path -LiteralPath $ngcPath) {
        Write-SmartM365Log "Taking ownership of NGC folder."
        takeown.exe /F $ngcPath /A /R | Out-Null
        icacls.exe $ngcPath /grant administrators:F /T | Out-Null
        Write-SmartM365Log "Deleting NGC folder."
        Remove-Item -LiteralPath $ngcPath -Recurse -Force -ErrorAction Stop
    }
    else {
        Write-SmartM365Log "NGC folder not found; nothing to delete."
    }

    if ($ngcService) {
        Write-SmartM365Log "Starting NgcSvc service."
        Start-Service -Name NgcSvc -ErrorAction SilentlyContinue
    }

    $dsreg = dsregcmd /status | Out-String
    if ($dsreg -notmatch "AzureAdJoined.*YES" -or $dsreg -notmatch "DomainJoined.*YES") {
        Write-SmartM365Log "Device join state is incomplete; requesting dsregcmd /join."
        dsregcmd /join | Out-Null
    }

    Write-SmartM365Log "RemediationCompleted"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    try { Write-SmartM365Log "RemediationFailed Message=$($_.Exception.Message)" } catch { Write-Output "LogWriteFailed=$($_.Exception.Message)" }
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBIorMPQduQ8GBT
# rp8kSRHb2VFu3QaxD7JqzDc4yzJ9raCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCCF5Wa+cQUN4gJpCpaNpDuyPkWXFQfaiCeGBnzszKd8JTANBgkqhkiG9w0B
# AQEFAASCAYB8EUgEgdf1HXz7LrrfecmHvaQ97ceJRETM+uONZBTKlVCBfeLO09yM
# FHfh47PuB/XNvyBxcK2A7wcz+LDl3U5qAu7i23ye4nJRBRCVBENB808ebmDmjKbV
# BA4gxabQvKhf8ij6CAE+ZV7L+iPvT69hf9rbJWQ+2P5N0MT5r7E7gfiM0UDHxCpz
# Dr9fpp8Dpl0i6j4syRQjqh334Z5lNP3MHlUuTe288v2S33RG4arwqSLTrWcHX/2H
# cdD82oHZnwDk/Xptz2MC2FrI/k9mwSifLMBuWnyjTxdhsxMZFd8aGKgZjWxJtpyf
# 5T2b+7lh1Mep8meQY/UjQr9oQqvbTnwt0ODfsRU7/I8FaIreHmfPcIb1zweWNZSi
# 5Qvg/t9yYsarOaekPZzIiY8flma1DjoFn3vsU1fx0ulcpbwxGMH/oMyjCq+JxATo
# 5qJDcjuCxmtK9Q3HZxU/uYhA0cYlL4FY+KrZUWrgUllpFJCXeEvI/D2qh4ZmV++H
# /yG9YDF+MWY=
# SIG # End signature block
