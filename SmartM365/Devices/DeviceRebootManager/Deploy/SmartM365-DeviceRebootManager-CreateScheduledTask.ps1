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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAQlLAqi4WISK7e
# nlAaMEGf7QDnqFJmc6dU1gd8i3qgCqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCg1Iw8ta+ztIQnYzQe
# dTBnTwFk7phgVquaxDo/RXAJtzANBgkqhkiG9w0BAQEFAASCAYBVyFiZN69NqOPV
# iW8jtSBbCKmOxeLLm0/+j1IxXBZkIW8nO690J/B/BJl8cYbm6htVFIe4FiszsvfL
# a1rn+NTB4IDyeKTdqTJSuR4kQnrzXtVgNvG0D/Du1kP8GP/+Qt22s2c0P8R9CUS0
# gkObjhxgCj4BD/7DbxUOBQx+xDGPPch/nfHEdc9xBSMxPqmIg5cJ/0wykR4WEBH+
# ZAEV/zMuheo/eZ1P1n2jCRONkIw1Mkrg3MvwmFevZnK+PEXlc/rXurhlKJWlYmsP
# J791alCwXW2DX8rveM4xkHMFs/ffGYR7UrTAO4l3ewYQK+Xke4+4nNNNk4tdYB5M
# mXUSJD8G3S8CZYteEv9K6VNh7A2m1J2K2KwJzHoXqBQ20V1PAr7sa9yMVqisqm4N
# hR+cGHp5miYFrdpFd+cGAi4RPkazSPeWcFQdfn/21aRUzU0uuFapB1PzR8EHN/S6
# cUwhFigEWVbqXJvuIepWTWlaMtdCdQFp4gDWT4aU9nbe+C5PDpc=
# SIG # End signature block
