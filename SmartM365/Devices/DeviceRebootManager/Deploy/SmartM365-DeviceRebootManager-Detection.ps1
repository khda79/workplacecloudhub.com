#Requires -Version 5.1

<#
.SYNOPSIS
    Intune Win32 detection script for SmartM365 Device Reboot Manager.
#>

[CmdletBinding()]
param(
    [string]$InstallPath = "$env:ProgramData\SmartM365\DeviceRebootManager",
    [string]$TaskPath = '\SmartM365\',
    [string]$TaskName = 'Device Reboot Manager'
)

$ErrorActionPreference = 'Stop'

$requiredFiles = @(
    'SmartM365-DeviceRebootManager-GUI.ps1'
    'SmartM365-DeviceRebootManager-GUI.strings.psd1'
    'SmartM365-DeviceRebootManager-GUI.config.json'
    'SmartM365.GuiSplash.ps1'
    'WorkplaceCloudHub.ico'
    'WorkplaceCloudHub-lockup-WPF.png'
)

$missing = New-Object System.Collections.Generic.List[string]

foreach ($fileName in $requiredFiles) {
    $path = Join-Path -Path $InstallPath -ChildPath $fileName
    if (-not (Test-Path -LiteralPath $path)) {
        $missing.Add($path)
    }
}

$task = Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction SilentlyContinue
if ($null -eq $task) {
    $missing.Add(("{0}{1}" -f $TaskPath, $TaskName))
}

if ($missing.Count -gt 0) {
    Write-Output 'SmartM365 Device Reboot Manager is not detected.'
    foreach ($item in $missing) {
        Write-Output ("Missing: {0}" -f $item)
    }
    exit 1
}

Write-Output 'SmartM365 Device Reboot Manager detected.'
exit 0

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCApxm1XqWefivlm
# +2h6BES7s9GM1xD4o0sdJDiaW/O49qCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDt09vhYNIxvyPJ2zHysOalGGSoS9X1Nbyp3RtAssfJOTANBgkqhkiG9w0B
# AQEFAASCAYCKAPBrN0yGK3SmzY8YlaeJq0piMY9G7798j0EGHPlZzrTyuR6oA3Hb
# 8FN7XpvDFdiDm0A9sYPJ9ISuyXWihLdSVLX16dECi4Fszt6MO9mK3y0eyJ1jrB62
# XeG7ncf3l7VGmO1pb6MGDUEVqVB1vYaBPAo6/yFsE+AD8E+3u4r6OwZgMonPgaoc
# 3bYRq5ydh+KX4MCt9U6p08GFzqLKWLItayCbuODTTsaYPCcbTczAO7uTDa9fZVOo
# Mjp1T6Nj5z/8DrEJO4Spc0TUCwO1SKN8iDd0XJq+UoyxZdaqU9NkB0xtwLcsAVhP
# BQetoPUVoJaRO/oHJSgMGBF9pkNU44BJOjvVeOg3Vnp7A9jTxMKfWunfL1LuMonW
# TuUZ8G1uIUULpGohiPnTBWtSEtkuXArXjL7dSn8Qn3Ha6Pb5jSn3+sn0NQv3xvnI
# lw5KDL2JQvFSRzdQZw6/hM9lamnXYW1mzSQe/Tvs2BQPIuXovr9AZdttZZkdYm6J
# rb2e7fUrGKc=
# SIG # End signature block
