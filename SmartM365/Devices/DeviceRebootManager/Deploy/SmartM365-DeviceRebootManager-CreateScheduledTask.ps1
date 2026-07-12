#Requires -Version 5.1

<#
.SYNOPSIS
    Creates the SmartM365 Device Reboot Manager scheduled task.

.DESCRIPTION
    Registers a user-interactive scheduled task that starts the WPF GUI at user
    logon and then regularly while a user session is available. This script is
    intended for Intune Win32 deployments running as SYSTEM.
#>

[CmdletBinding()]
param(
    [string]$InstallPath = "$env:ProgramData\SmartM365\DeviceRebootManager",
    [string]$TaskPath = '\SmartM365\',
    [string]$TaskName = 'Device Reboot Manager',
    [int]$RepeatIntervalMinutes = 240,
    [string]$PowerShellPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe",
    [string]$ConfigPath = '',
    [switch]$RunOnceNow
)

$ErrorActionPreference = 'Stop'

function Quote-Argument {
    param([Parameter(Mandatory = $true)][string]$Value)
    return ('"{0}"' -f ($Value -replace '"', '\"'))
}

if ($RepeatIntervalMinutes -lt 15) {
    throw 'RepeatIntervalMinutes must be at least 15.'
}

if (-not (Test-Path -LiteralPath $PowerShellPath)) {
    $PowerShellPath = 'powershell.exe'
}

$appScriptPath = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager-GUI.ps1'
if (-not (Test-Path -LiteralPath $appScriptPath)) {
    throw "Device Reboot Manager script not found: $appScriptPath"
}

if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $candidateConfigPath = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager-GUI.config.json'
    if (Test-Path -LiteralPath $candidateConfigPath) {
        $ConfigPath = $candidateConfigPath
    }
}

$taskArguments = @(
    '-STA'
    '-NoProfile'
    '-WindowStyle'
    'Hidden'
    '-ExecutionPolicy'
    'Bypass'
    '-File'
    (Quote-Argument -Value $appScriptPath)
)

if (-not [string]::IsNullOrWhiteSpace($ConfigPath)) {
    $taskArguments += @('-ConfigPath', (Quote-Argument -Value $ConfigPath))
}

$action = New-ScheduledTaskAction -Execute $PowerShellPath -Argument ($taskArguments -join ' ')
$logonTrigger = New-ScheduledTaskTrigger -AtLogOn
$repeatTrigger = New-ScheduledTaskTrigger -Daily -At (Get-Date)
$repeatTrigger.Repetition.Interval = New-TimeSpan -Minutes $RepeatIntervalMinutes
$repeatTrigger.Repetition.Duration = New-TimeSpan -Days 1

$principal = New-ScheduledTaskPrincipal -GroupId 'S-1-5-32-545' -LogonType Interactive -RunLevel LeastPrivilege
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1)

Register-ScheduledTask `
    -TaskPath $TaskPath `
    -TaskName $TaskName `
    -Action $action `
    -Trigger @($logonTrigger, $repeatTrigger) `
    -Principal $principal `
    -Settings $settings `
    -Description 'Starts the SmartM365 Device Reboot Manager GUI in the interactive user session.' `
    -Force | Out-Null

if ($RunOnceNow) {
    Start-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName
}

Write-Output ("Scheduled task registered: {0}{1}" -f $TaskPath, $TaskName)

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAQlLAqi4WISK7e
# nlAaMEGf7QDnqFJmc6dU1gd8i3qgCqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCCg1Iw8ta+ztIQnYzQedTBnTwFk7phgVquaxDo/RXAJtzANBgkqhkiG9w0B
# AQEFAASCAYBLSFfw4dkZ94UUXWZtaHMS9C/8TDKW0ODFcJYzUYiW000fdOmcph2V
# dYfgIJelwWRuNKjUKODk/60XWCkqzFwDvzqZsdXjyoE7s7phRRqTQmuGGvX2in+a
# fD3H7NOi9VD5TC/G7vR4xzi+XmMcNmOnBDubDjuBLPlzweFmG4Ah85G7TWRYjouF
# SjSYwU1+0c4PE2cfv6pOKwcYqF+ygR8KJyIm3Kqt98I/y2eHJ1IHPVCYm+F6Zq8G
# RtUI1ky/n5dojts2aaSVmejGDLJbbO/TYJ/23wAmBItSKBZgkICK0cN+pBxDoLUA
# KZZC7GKu86jo2wD/vywh3C5Y/UqXgMNlEbbnFTxC/xuCZ1uun7bwD/7nhRFNINb/
# cS+mMClKVjOFbs7Y6Kv0WXrSWCCt7J4cKX9uU/oGsaxeiK4CxOkPO4osrsXnLaW9
# 4s9C/NltMeBzykmZTOsxUKmfnNoams5WpLHRJtzgIbsV7U1pnpb/cZyV26+9/qZ9
# 7VdTsj1NaQE=
# SIG # End signature block
