#Requires -Version 5.1

<#
.SYNOPSIS
Builds the local PowerShell Gallery package for SmartM365 Device Reboot Manager.

.DESCRIPTION
Creates a clean module folder from an explicit allow-list, validates PowerShell
syntax and the pinned WorkplaceCloudHub Authenticode signer, and never copies
local runtime configuration or LOCAL_MEMORY.md.
#>

[CmdletBinding()]
param(
    [string]$SourceRoot = (Split-Path -Path $PSScriptRoot -Parent),
    [string]$OutputRoot = (Join-Path ([IO.Path]::GetTempPath()) 'SmartM365\PowerShellGallery'),
    [string]$ExpectedSignerThumbprint = 'D70ECB7B00377EBFB76B304C08DFC6620584E114',
    [switch]$Force,
    [switch]$SkipAuthenticodeValidation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-NormalizedThumbprint {
    param([string]$Value)
    return ([string]$Value).Replace(' ', '').ToUpperInvariant()
}

function Assert-SafeChildPath {
    param(
        [Parameter(Mandatory = $true)][string]$Root,
        [Parameter(Mandatory = $true)][string]$Child
    )

    $resolvedRoot = [IO.Path]::GetFullPath($Root).TrimEnd('\', '/')
    $resolvedChild = [IO.Path]::GetFullPath($Child).TrimEnd('\', '/')
    $prefix = $resolvedRoot + [IO.Path]::DirectorySeparatorChar
    if ($resolvedChild -eq $resolvedRoot -or -not $resolvedChild.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Unsafe build target outside OutputRoot: $resolvedChild"
    }
}

function Assert-PowerShellSyntax {
    param([Parameter(Mandatory = $true)][string]$Path)

    $parseErrors = $null
    [void][Management.Automation.PSParser]::Tokenize(
        (Get-Content -LiteralPath $Path -Raw),
        [ref]$parseErrors
    )
    if ($parseErrors.Count -gt 0) {
        $messages = $parseErrors | ForEach-Object { $_.Message }
        throw ("PowerShell syntax validation failed for {0}: {1}" -f $Path,($messages -join '; '))
    }
}

function Assert-PinnedSignature {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Thumbprint
    )

    $signature = Get-AuthenticodeSignature -LiteralPath $Path
    $actualThumbprint = if ($signature.SignerCertificate) {
        Get-NormalizedThumbprint $signature.SignerCertificate.Thumbprint
    }
    else {
        ''
    }

    if ($signature.Status -ne 'Valid' -or $actualThumbprint -ne (Get-NormalizedThumbprint $Thumbprint)) {
        throw ("Authenticode validation failed: file={0}; status={1}; signer={2}" -f $Path,$signature.Status,$actualThumbprint)
    }
}

$moduleTemplateRoot = Join-Path -Path $PSScriptRoot -ChildPath 'Module'
$manifestSourcePath = Join-Path -Path $moduleTemplateRoot -ChildPath 'SmartM365.DeviceRebootManager.psd1'
if (-not (Test-Path -LiteralPath $manifestSourcePath -PathType Leaf)) {
    throw "Module manifest not found: $manifestSourcePath"
}

$manifest = Test-ModuleManifest -Path $manifestSourcePath
$moduleName = [string]$manifest.Name
$moduleVersion = [string]$manifest.Version
$prerelease = [string]$manifest.PrivateData.PSData['Prerelease']
$packageVersion = if ([string]::IsNullOrWhiteSpace($prerelease)) {
    $moduleVersion
}
else {
    '{0}-{1}' -f $moduleVersion,$prerelease
}

$versionManifestSourcePath = Join-Path -Path $SourceRoot -ChildPath 'SmartM365-DeviceRebootManager.version.json.txt'
if (-not (Test-Path -LiteralPath $versionManifestSourcePath -PathType Leaf)) {
    throw "Package version manifest not found: $versionManifestSourcePath"
}
$versionManifest = Get-Content -LiteralPath $versionManifestSourcePath -Raw | ConvertFrom-Json
if ([int]$versionManifest.SchemaVersion -ne 1 -or [string]$versionManifest.PackageVersion -ne $packageVersion) {
    throw ("Package version mismatch: manifest={0}; module={1}" -f $versionManifest.PackageVersion,$packageVersion)
}

$packagePath = Join-Path -Path (Join-Path -Path $OutputRoot -ChildPath $moduleName) -ChildPath $moduleVersion
Assert-SafeChildPath -Root $OutputRoot -Child $packagePath

if (Test-Path -LiteralPath $packagePath) {
    if (-not $Force) {
        throw "Build target already exists. Use -Force to replace it: $packagePath"
    }
    Remove-Item -LiteralPath $packagePath -Recurse -Force
}

New-Item -ItemType Directory -Path $packagePath -Force | Out-Null

$moduleFiles = @(
    'SmartM365.DeviceRebootManager.psd1'
    'SmartM365.DeviceRebootManager.psm1'
    'Tools\SmartM365-DeviceRebootManager-GalleryUpdate.ps1'
)

$runtimeFiles = @(
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

$deployFiles = @(
    'SmartM365-DeviceRebootManager-CreateScheduledTask.ps1'
    'SmartM365-DeviceRebootManager-Detection.ps1'
    'SmartM365-DeviceRebootManager-Install.ps1'
    'SmartM365-DeviceRebootManager-Uninstall.ps1'
)

foreach ($relativePath in $moduleFiles) {
    $sourcePath = Join-Path -Path $moduleTemplateRoot -ChildPath $relativePath
    $destinationPath = Join-Path -Path $packagePath -ChildPath $relativePath
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Required module file not found: $sourcePath"
    }
    New-Item -ItemType Directory -Path (Split-Path -Path $destinationPath -Parent) -Force | Out-Null
    Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Force
}

foreach ($relativePath in $runtimeFiles) {
    $sourcePath = Join-Path -Path $SourceRoot -ChildPath $relativePath
    $destinationPath = Join-Path -Path (Join-Path $packagePath 'Runtime') -ChildPath $relativePath
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Required runtime file not found: $sourcePath"
    }
    New-Item -ItemType Directory -Path (Split-Path -Path $destinationPath -Parent) -Force | Out-Null
    Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Force
}

foreach ($relativePath in $deployFiles) {
    $sourcePath = Join-Path -Path (Join-Path $SourceRoot 'Deploy') -ChildPath $relativePath
    $destinationPath = Join-Path -Path (Join-Path $packagePath 'Runtime\Deploy') -ChildPath $relativePath
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Required deployment file not found: $sourcePath"
    }
    New-Item -ItemType Directory -Path (Split-Path -Path $destinationPath -Parent) -Force | Out-Null
    Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Force
}

$packagedFiles = @(Get-ChildItem -LiteralPath $packagePath -Recurse -File)
$powerShellFiles = @($packagedFiles | Where-Object { $_.Extension -in @('.ps1', '.psm1', '.psd1') })
foreach ($file in $powerShellFiles) {
    Assert-PowerShellSyntax -Path $file.FullName
    if (-not $SkipAuthenticodeValidation) {
        Assert-PinnedSignature -Path $file.FullName -Thumbprint $ExpectedSignerThumbprint
    }
}

$builtManifestPath = Join-Path -Path $packagePath -ChildPath 'SmartM365.DeviceRebootManager.psd1'
[void](Test-ModuleManifest -Path $builtManifestPath)

[pscustomobject]@{
    ModuleName               = $moduleName
    Version                  = $moduleVersion
    Prerelease               = $prerelease
    PackageVersion           = $packageVersion
    PackagePath              = $packagePath
    IntuneSourcePath         = Join-Path -Path $packagePath -ChildPath 'Runtime'
    FileCount                = $packagedFiles.Count
    PowerShellFileCount      = $powerShellFiles.Count
    SignaturesValidated      = (-not $SkipAuthenticodeValidation)
    ExpectedSignerThumbprint = Get-NormalizedThumbprint $ExpectedSignerThumbprint
    Ready                    = $true
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAucoBhU54VFuYf
# moge4LEtkmzUHlO4pw//ZLnaSwq0YKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBHG7WApDL4+ErosRySZ9w9
# aSB9XeUStqqxDjoTLL8ddDANBgkqhkiG9w0BAQEFAASCAYCMVzcc5CeMCkdULTA7
# ursRQJu5lE2OkPrrF/babxR1Hi9THYot0udZIaY7tTXvflaN3GH01iLWZ2kmjJav
# O1EzlRMZ7AzKP7wrb0qHu1TLCmFDfpJz6e9G3uF8IxSyyaHhSGU41vBe2dwBPqI7
# 2DIytbnuZD2j3u9HLEBoKq9jCNa59i2AqP6AFMLFtLpsoXr3zhDrgjp0dvPfPzBW
# QEr2Sr5GEiQG68h7w90/vkRlVJJYf18lL7WYOoEQbiqwiJtVA8FD0iv1f3azSla5
# k+S3l3+nhFJIqrYVAAnhWS5rKfN8XnT+aK2zFnGAug8My0u93B29vNIBZ8Z0/adT
# XbA9ccwBv5avzS+H3PU4kCvw2DocAiw9yybSfrhp+sLcKDSdXtRbh1lGSmmnU8A2
# Jqzfq0HBeXaTIHf8JF2kCo9nPtxYu8GVgoJNimK4x++d/+nVKM29DH+4ZFvbGGcM
# Qg++OCG5spnYUftL72f0N26Dne3ptf6Ip1AvVLJnIPYfMh4=
# SIG # End signature block
