<#
.SYNOPSIS
    Version: 1.0
    Repairs inconsistent WinHTTP proxy state for WinHTTP-Proxy.
.DESCRIPTION
    Reads the current WinHTTP proxy state and resets it to direct access when the configuration is unreadable or inconsistent. This does not change user-level browser proxy settings.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$Scenario = 'WinHTTP-Proxy'
$LogRoot = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path $LogRoot "$Scenario-Remediation.log"
function Write-SmartM365Log { param([string]$Message) $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Scenario, $Message; Write-Output $line; Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8 }

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-SmartM365Log 'Remediation started.'
    $proxyOutput = netsh winhttp show proxy 2>&1
    $proxyText = ($proxyOutput -join ' ')
    Write-SmartM365Log "Current WinHTTP proxy output: $proxyText"
    if ($LASTEXITCODE -ne 0 -or ($proxyText -notmatch 'Direct access|Acc.s direct|Acces direct' -and $proxyText -notmatch 'Proxy Server|Serveur proxy')) {
        netsh winhttp reset proxy | Out-Null
        Write-SmartM365Log 'WinHTTP proxy reset to direct access.'
    }
    else {
        Write-SmartM365Log 'WinHTTP proxy configuration is already valid; no reset required.'
    }
    exit 0
}
catch {
    Write-SmartM365Log "Remediation failed: $($_.Exception.Message)"
    exit 1
}


# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB+DYbR+h6+ZWyb
# cclGliA4ic4GUh6GcMBqMY8zIA3ZaaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCB6L+L0JJL59zzlmtbQdT4H1wocnVnY2ueh1U6FoAdFLjANBgkqhkiG9w0B
# AQEFAASCAYAZtTO7s1hXaOJRs6EZGhxboc5uknJJU6+gosWKbQAd5IeEQkbC6dNi
# sknntogcZ3C6FJZobgvDo9SVT/bwQUaUAzN0PxZGpNj2ZR4UkzwmiD2EHPXqnLan
# IOmXq7y/MZKsWAxZGJ4hIcl+VgtDpmSNJWPnVR3M1VwJ2ctmLTdpRqHtmHyOeu0u
# xx0k5SH5xyYgV2bslaayQX23PKj1Lp+IxfU6/HEppYJ5GwtobRgKh6WrA6aUHKy3
# 1qdm3kBBBWIVKPQEd5BkS5PeZfxhRUZMgf08Y8q2FOHVEcK1x7QeNuPgkXxScRde
# CFw93KvdOGFRdpIMoAt7KygkvdDGizdpFkFnidnQsk2bfiBusKHcIodDuF60d9yi
# 3tjnRxqiMbRrPuaAKBO0BvUYh1lOWZBGodEr1dt1O4vOOZ8uRTdxrFlzMEpJ5IcI
# cq0mD6t4dnyJbw4Bt22lM8pIsF4ngC9e77roj9zcxU+h3RIqtnPce4OrYFpryLuu
# BlKk1U79ZAI=
# SIG # End signature block
