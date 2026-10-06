#Requires -Version 5.1

<#
.SYNOPSIS
    Installs SmartM365 Device Reboot Manager for Intune Win32 deployment.

.DESCRIPTION
    Copies runtime files to ProgramData and creates the scheduled task used to
    launch the GUI in the interactive user session. Intune is the default
    deployment channel and removes the optional PowerShell Gallery update task.
#>

[CmdletBinding()]
param(
    [string]$InstallPath = "$env:ProgramData\SmartM365\DeviceRebootManager",
    [string]$ConfigSourcePath = '',
    [switch]$ForceConfig,
    [switch]$SkipScheduledTask,
    [string]$TaskPath = '\SmartM365\',
    [string]$TaskName = 'Device Reboot Manager',
    [int]$RepeatIntervalMinutes = 240,
    [ValidateSet('Intune', 'PowerShellGallery', 'Local')]
    [string]$PackageSource = 'Intune',
    [string]$UpdateTaskPath = '\SmartM365\',
    [string]$UpdateTaskName = 'Device Reboot Manager Update'
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
    'SmartM365-DeviceRebootManager.version.json.txt'
    'SmartM365-DeviceRebootManager-GUI.ps1'
    'SmartM365-DeviceRebootManager-GUI.strings.psd1'
    'SmartM365-DeviceRebootManager-GUI.config.json.template'
    'SmartM365.GuiSplash.ps1'
    'WorkplaceCloudHub.ico'
    'WorkplaceCloudHub-lockup-WPF.png'
    'Start-SmartM365-DeviceRebootManager-GUI.cmd'
    'Start-SmartM365-DeviceRebootManager-GUI-Test.cmd'
)

$versionSourcePath = Join-Path -Path $sourceRoot -ChildPath 'SmartM365-DeviceRebootManager.version.json.txt'
if (-not (Test-Path -LiteralPath $versionSourcePath -PathType Leaf)) {
    throw "Package version manifest not found: $versionSourcePath"
}

try {
    $packageVersionManifest = Get-Content -LiteralPath $versionSourcePath -Raw | ConvertFrom-Json
}
catch {
    throw "Package version manifest is invalid: $($_.Exception.Message)"
}

if ([int]$packageVersionManifest.SchemaVersion -ne 1 -or
    [string]::IsNullOrWhiteSpace([string]$packageVersionManifest.ProductName) -or
    [string]::IsNullOrWhiteSpace([string]$packageVersionManifest.PackageVersion)) {
    throw 'Package version manifest is incomplete or uses an unsupported schema.'
}

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

if ($PackageSource -eq 'Intune') {
    $galleryUpdateTask = Get-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -ErrorAction SilentlyContinue
    if ($null -ne $galleryUpdateTask) {
        Unregister-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -Confirm:$false
    }
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

$installationMetadata = [ordered]@{
    SchemaVersion  = 1
    ProductName    = [string]$packageVersionManifest.ProductName
    PackageVersion = [string]$packageVersionManifest.PackageVersion
    PackageSource  = $PackageSource
    InstalledAtUtc = [DateTime]::UtcNow.ToString('o')
}
$installationMetadata |
    ConvertTo-Json -Depth 3 |
    Set-Content -LiteralPath (Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager.installation.json') -Encoding UTF8

Write-Output ("SmartM365 Device Reboot Manager {0} installed to: {1} (source={2})" -f
    $packageVersionManifest.PackageVersion,$InstallPath,$PackageSource)

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAgicnOh9QT80an
# KG5Q8A+Jv9qkOu+4aH2MewH3rIJ88qCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCB/6fVUKmDr9TOgM8hOPbVv
# GLKEpk0UYtlfm1E4vlughjANBgkqhkiG9w0BAQEFAASCAYCsqi++cOgB4HZ1SS22
# B9uMIPcX5Yi2C3kqdozDxZTWeAdtwm4NfxeKiqb9GOwJcqIBLCmUY+fDbVMdep0h
# YsHZZLHiVS++Qj3VQyY2qDQVg2jspfx02km4rInK1DelCh9sqU+lH4pTGCHT6shi
# CyrVu0mibtTt4hAQMnBb7u6dG5zBaTtlpaoytMnmqzCHE+fSxcHIdAgQKPezGxOo
# ZbG8uC+ySIR84I3pffGlfuJd2hdT75ltZlsgHr7XJvaiWFZBqK4mq3yd820Vvi52
# vK5sNHWKaycWx63FF6AWPpctej8zOyLGAriIf5kNIMzWJPkjaQobjvvzqz0dcRc3
# oqu9u3WKLHxWckkWLfZeZIbyYcG0gCkniKB5u3lI/kyBQK9J/bNKlXGW9RK15iJU
# kO4GYfAXXAKejGtU0l4S5pw9nm9WJbIwF2NZJgz2/HTQB9o0J24yk/h20QvXKLmk
# HwQ54jDkE4t7rlkL4zdNtiE/hot59oq3j4HsFBX+U8njxKI=
# SIG # End signature block
