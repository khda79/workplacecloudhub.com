#Requires -Version 5.1

<#
.SYNOPSIS
    Uninstalls SmartM365 Device Reboot Manager.
#>

[CmdletBinding()]
param(
    [string]$InstallPath = "$env:ProgramData\SmartM365\DeviceRebootManager",
    [string]$TaskPath = '\SmartM365\',
    [string]$TaskName = 'Device Reboot Manager',
    [switch]$KeepConfig
)

$ErrorActionPreference = 'Stop'

try {
    $task = Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($null -ne $task) {
        Unregister-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -Confirm:$false
    }
}
catch {
    Write-Warning ("Unable to remove scheduled task. {0}" -f $_.Exception.Message)
}

if (Test-Path -LiteralPath $InstallPath) {
    if ($KeepConfig) {
        $configPath = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager-GUI.config.json'
        Get-ChildItem -LiteralPath $InstallPath -Force |
            Where-Object { $_.FullName -ne $configPath } |
            ForEach-Object { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
    }
    else {
        Remove-Item -LiteralPath $InstallPath -Recurse -Force
    }
}

Write-Output 'SmartM365 Device Reboot Manager uninstalled.'

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCClP6rvNAHvK4Nz
# hLYLglCkZHt+cmtzsE+7nY8LK9Ri1KCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCAD89s8CW7DhZM4XXR/K5xo48Pv5AqvsG0hs16H6urSfDANBgkqhkiG9w0B
# AQEFAASCAYATC7TCAzlQ4iZU6p1sbx8dolFblhTzw6TVBCv8bOS5pzVh9sPO7HW6
# lM07vuBBq27nC3KJB2Mi2yaYZfXLNkEDJiCSA4oaNWkuF5s/9vDU5pinWqY5lXj2
# uf0nZOoVxjB525soziAOth6Pujn0sjwbFv1KBxTRe8asN5y5jNVBm5k+DwgqNjVk
# KZk+JX94KL0MRB2JlbSanhkwrpsDddZ21GOfOJfhmNrDW92eYiL8V2lrw5ft9eyP
# Er87d3nACpbMycSh4seL9P0jXu7p7xfPFz0syAvagOtrYFWWOw9rSjKyV03I6Sp0
# XBCmNvXOWjjzoHy6i68qjC44RchX53K4G+0XHQAeNgwfsyVAacqi/ktT19x3BjRY
# v5FqoYa8RymAORa4riUGkT3OBc8JsaADkG/X27NiQbrvndMugngUhmTAkVPrtCPW
# VQIMuH75q81ySz/Kg0QL2dgTWR8/DolqVez3qllH2uJSC0uAjPZqfNEI1U+oG+FR
# yrm+WAOGJ/k=
# SIG # End signature block
