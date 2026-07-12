#Requires -Version 5.1

<#
.SYNOPSIS
    Installs SmartM365 Device Reboot Manager for Intune Win32 deployment.

.DESCRIPTION
    Copies runtime files to ProgramData and creates the scheduled task used to
    launch the GUI in the interactive user session.
#>

[CmdletBinding()]
param(
    [string]$InstallPath = "$env:ProgramData\SmartM365\DeviceRebootManager",
    [string]$ConfigSourcePath = '',
    [switch]$ForceConfig,
    [switch]$SkipScheduledTask,
    [string]$TaskPath = '\SmartM365\',
    [string]$TaskName = 'Device Reboot Manager',
    [int]$RepeatIntervalMinutes = 240
)

$ErrorActionPreference = 'Stop'

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdministrator)) {
    throw 'Installation must run elevated. Intune Win32 install should run as SYSTEM.'
}

$deployRoot = $PSScriptRoot
$sourceRoot = Split-Path -Path $deployRoot -Parent

$requiredFiles = @(
    'SmartM365-DeviceRebootManager-GUI.ps1'
    'SmartM365-DeviceRebootManager-GUI.strings.psd1'
    'SmartM365-DeviceRebootManager-GUI.config.json.template'
    'SmartM365.GuiSplash.ps1'
    'WorkplaceCloudHub.ico'
    'WorkplaceCloudHub-lockup-WPF.png'
    'Start-SmartM365-DeviceRebootManager-GUI.cmd'
    'Start-SmartM365-DeviceRebootManager-GUI-Test.cmd'
)

New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null

foreach ($fileName in $requiredFiles) {
    $sourcePath = Join-Path -Path $sourceRoot -ChildPath $fileName
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        throw "Required package file not found: $sourcePath"
    }

    Copy-Item -LiteralPath $sourcePath -Destination (Join-Path -Path $InstallPath -ChildPath $fileName) -Force
}

$runtimeConfigPath = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager-GUI.config.json'
if (-not [string]::IsNullOrWhiteSpace($ConfigSourcePath)) {
    if (-not (Test-Path -LiteralPath $ConfigSourcePath)) {
        throw "Config source file not found: $ConfigSourcePath"
    }

    Copy-Item -LiteralPath $ConfigSourcePath -Destination $runtimeConfigPath -Force
}
elseif ($ForceConfig -or -not (Test-Path -LiteralPath $runtimeConfigPath)) {
    $templatePath = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager-GUI.config.json.template'
    Copy-Item -LiteralPath $templatePath -Destination $runtimeConfigPath -Force
}

if (-not $SkipScheduledTask) {
    $taskScriptPath = Join-Path -Path $deployRoot -ChildPath 'SmartM365-DeviceRebootManager-CreateScheduledTask.ps1'
    if (-not (Test-Path -LiteralPath $taskScriptPath)) {
        throw "Scheduled task helper not found: $taskScriptPath"
    }

    & $taskScriptPath `
        -InstallPath $InstallPath `
        -TaskPath $TaskPath `
        -TaskName $TaskName `
        -RepeatIntervalMinutes $RepeatIntervalMinutes `
        -ConfigPath $runtimeConfigPath
}

Write-Output "SmartM365 Device Reboot Manager installed to: $InstallPath"

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDvu8FDtxCEaAwH
# 25TwuLjUTEyx6qCAk3xLaAnLQgTQRqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBBg3VU6JMa4XNgsh7x
# cuqIj8NVAmeANly7AG6f9uP6HjANBgkqhkiG9w0BAQEFAASCAYBGD+dTJCQ0VYpo
# hrnILuICJ5ahXipfIoqn1zY94YWjOGVtzz4MQ5TvDFcSSMkfYHI6TViwxl6FHjyX
# Vy9krzdCvtKhGgkqYJHegyQ8P9d9rPEYICyNUkjwu2FfkS2+wUv2o/2OfZpPHchH
# x1fWQnH6hFoPkLSwGm/pJiZMpc2lqwPq8q3o3IqPvaxuNh/rAJYFmwAXkjKm/mMP
# VwTTrnoPJXeAIy+FYYsLSm/qS2WKRRovKVBsFZaTZZF30ivEu7C2mPhPBKzgDQop
# L6YYm1S6IGpJuMdTFIV2h3eAJ41nEmFZDA+3D3dXd5RJQ//Z/YTjB53CSbgeoYqP
# YKvipta6yPbJnkyNdp5bMsjs9NIuf1QFizU7Z0v8vGYkctkxcWXW2pT5CjntbjOg
# VhWjeoCJQuqtacgmEYoaGnRIetldpWsQQpPuHt1ZLVeb6emTCPIzdiO7Qip1SmSb
# aS0255NuuToCqoGuG0jx69B19Pxu4Qe2Lc0ewa/gbXGlubyYtxA=
# SIG # End signature block
