#Requires -Version 5.1

<#
.SYNOPSIS
    Intune Win32 detection script for SmartM365 Device Reboot Manager.
#>

[CmdletBinding()]
param(
    [string]$InstallPath = "$env:ProgramData\SmartM365\DeviceRebootManager",
    [string]$TaskPath = '\SmartM365\',
    [string]$TaskName = 'Device Reboot Manager',
    [string]$ExpectedVersion = '0.1.0',
    [string]$UpdateTaskPath = '\SmartM365\',
    [string]$UpdateTaskName = 'Device Reboot Manager Update',
    [bool]$RequireGalleryAutomaticUpdateDisabled = $true
)

$ErrorActionPreference = 'Stop'

$requiredFiles = @(
    'SmartM365-DeviceRebootManager.version.json.txt'
    'SmartM365-DeviceRebootManager.installation.json'
    'SmartM365-DeviceRebootManager-GUI.ps1'
    'SmartM365-DeviceRebootManager-GUI.strings.psd1'
    'SmartM365-DeviceRebootManager-GUI.config.json'
    'SmartM365.GuiSplash.ps1'
    'WorkplaceCloudHub.ico'
    'WorkplaceCloudHub-lockup-WPF.png'
)

$issues = New-Object System.Collections.Generic.List[string]

foreach ($fileName in $requiredFiles) {
    $path = Join-Path -Path $InstallPath -ChildPath $fileName
    if (-not (Test-Path -LiteralPath $path)) {
        $issues.Add(("Missing file: {0}" -f $path))
    }
}

$task = Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction SilentlyContinue
if ($null -eq $task) {
    $issues.Add(("Missing scheduled task: {0}{1}" -f $TaskPath,$TaskName))
}

$installedVersion = ''
$versionPath = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager.version.json.txt'
if (Test-Path -LiteralPath $versionPath -PathType Leaf) {
    try {
        $versionManifest = Get-Content -LiteralPath $versionPath -Raw | ConvertFrom-Json
        $installedVersion = [string]$versionManifest.PackageVersion
        if ([int]$versionManifest.SchemaVersion -ne 1) {
            $issues.Add(("Unsupported version manifest schema: {0}" -f $versionManifest.SchemaVersion))
        }
        if ($installedVersion -ne $ExpectedVersion) {
            $issues.Add(("Version mismatch: installed={0}; expected={1}" -f $installedVersion,$ExpectedVersion))
        }
    }
    catch {
        $issues.Add(("Invalid version manifest: {0}" -f $_.Exception.Message))
    }
}

$installationMetadataPath = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager.installation.json'
if (Test-Path -LiteralPath $installationMetadataPath -PathType Leaf) {
    try {
        $installationMetadata = Get-Content -LiteralPath $installationMetadataPath -Raw | ConvertFrom-Json
        if ([string]$installationMetadata.PackageVersion -ne $ExpectedVersion) {
            $issues.Add(("Installation metadata version mismatch: installed={0}; expected={1}" -f
                $installationMetadata.PackageVersion,$ExpectedVersion))
        }
        if ([string]$installationMetadata.PackageSource -ne 'Intune') {
            $issues.Add(("Unexpected installation source: {0}" -f $installationMetadata.PackageSource))
        }
    }
    catch {
        $issues.Add(("Invalid installation metadata: {0}" -f $_.Exception.Message))
    }
}

if ($RequireGalleryAutomaticUpdateDisabled) {
    $galleryUpdateTask = Get-ScheduledTask -TaskPath $UpdateTaskPath -TaskName $UpdateTaskName -ErrorAction SilentlyContinue
    if ($null -ne $galleryUpdateTask) {
        $issues.Add(("Unexpected PowerShell Gallery update task: {0}{1}" -f $UpdateTaskPath,$UpdateTaskName))
    }
}

if ($issues.Count -gt 0) {
    Write-Output 'SmartM365 Device Reboot Manager is not detected.'
    foreach ($issue in $issues) {
        Write-Output $issue
    }
    exit 1
}

Write-Output ("SmartM365 Device Reboot Manager {0} detected; PowerShell Gallery automatic update is disabled." -f
    $installedVersion)
exit 0

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDKumLTvNZW3qEo
# 9Ar10lBvQUDmVBUBRihjrtFah+NjGqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAh63mH+4LC1M/BTX4+h0fV
# kjocw5TOES/+7z0soP6t7DANBgkqhkiG9w0BAQEFAASCAYA5SINfJWUfUY+u+0Uj
# s+PoLeFjvihwo4BByHNiSpOlD/Irm1EJ7LQsNi6ZOd8VQngOK+hCfTzQ3rVe5GUO
# 1HlQRbLL+nwzi8ZHRQyqcwuDWoSGwu/not3WugP2QQ7IexLhAIUe46Q2heuUqlhq
# sByDXitivLrCV5BFhBqd0kxxWKSEQCXWQYvU1rp8cZqv5uyA9uyJ2eoRmMgFBrs8
# R7K59OPutkL5QpiyEPHWiy52sqiSbIi64zb7b+JJQM68cGg8dJm5jXzf8ke9HsXJ
# feFRtEdoIzbMhxowKAk845DXU26JmPzgxM0qhA75DdAITsrX7WtuyKFKEuIIon3K
# qXAQGP3g+Pc1uCAT8hV0650mZ3b5V6zKnVb6DZ1b8PlER9sdn6D8Xy3TrU2U2oBF
# /UkyX0esUijt8Sm+lp/MwdcKek68UAm01gglAMhLRvrtEkVpLhthVjMrV/cCwe9Z
# /1JR3vPhJKtB1fDJxH/4siZXUWiz6GTvEoBdMiJH7o0E3SA=
# SIG # End signature block
