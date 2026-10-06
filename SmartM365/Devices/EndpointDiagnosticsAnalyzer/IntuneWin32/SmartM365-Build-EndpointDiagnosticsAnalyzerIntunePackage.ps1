#Requires -Version 5.1

[CmdletBinding()]
param(
    [string]$IntuneWinAppUtilPath = 'C:\tmp\SmartM365Tools\IntuneWinAppUtil.exe',
    [string]$OutputRoot = (Join-Path ([IO.Path]::GetTempPath()) 'SmartM365\IntuneWin32\EndpointDiagnosticsAnalyzer'),
    [string]$GalleryBuildRoot = (Join-Path ([IO.Path]::GetTempPath()) 'SmartM365\PowerShellGallery'),
    [switch]$Force,
    [switch]$SkipAuthenticodeValidation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$applicationRoot = Split-Path -Parent $PSScriptRoot
$versionPath = Join-Path $applicationRoot 'SmartM365-EndpointDiagnosticsAnalyzer.version.json.txt'
$version = Get-Content -LiteralPath $versionPath -Raw | ConvertFrom-Json
$packageVersion = [string]$version.PackageVersion
if ([string]::IsNullOrWhiteSpace($packageVersion)) {
    throw 'PackageVersion is missing from the version manifest.'
}
if (-not (Test-Path -LiteralPath $IntuneWinAppUtilPath -PathType Leaf)) {
    throw "IntuneWinAppUtil.exe not found: $IntuneWinAppUtilPath"
}

$buildScript = Join-Path $applicationRoot 'PowerShellGallery\SmartM365-Build-EndpointDiagnosticsAnalyzerGalleryPackage.ps1'
$buildParameters = @{
    OutputRoot = $GalleryBuildRoot
    Force = $Force
    SkipAuthenticodeValidation = $SkipAuthenticodeValidation
}
$galleryBuild = & $buildScript @buildParameters
$sourcePath = [string]$galleryBuild.IntuneSourcePath
$setupFile = 'Deploy\SmartM365-EndpointDiagnosticsAnalyzer-Install.ps1'
if (-not (Test-Path -LiteralPath (Join-Path $sourcePath $setupFile) -PathType Leaf)) {
    throw "Intune setup file is missing from the staging source: $setupFile"
}

$resolvedOutputRoot = [IO.Path]::GetFullPath($OutputRoot)
$outputPath = Join-Path $resolvedOutputRoot $packageVersion
$normalizedRoot = $resolvedOutputRoot.TrimEnd('\') + '\'
if (-not ([IO.Path]::GetFullPath($outputPath).StartsWith($normalizedRoot, [StringComparison]::OrdinalIgnoreCase))) {
    throw "Unsafe Intune output path: $outputPath"
}
if (Test-Path -LiteralPath $outputPath) {
    if (-not $Force) { throw "Intune build target already exists: $outputPath" }
    Remove-Item -LiteralPath $outputPath -Recurse -Force
}
New-Item -ItemType Directory -Path $outputPath -Force | Out-Null

& $IntuneWinAppUtilPath -c $sourcePath -s $setupFile -o $outputPath -q
if ($LASTEXITCODE -ne 0) {
    throw "IntuneWinAppUtil failed with exit code $LASTEXITCODE."
}
$generatedName = ([IO.Path]::GetFileNameWithoutExtension($setupFile) + '.intunewin')
$generatedPath = Join-Path $outputPath $generatedName
if (-not (Test-Path -LiteralPath $generatedPath -PathType Leaf)) {
    $generatedPath = Get-ChildItem -LiteralPath $outputPath -Filter '*.intunewin' -File |
        Select-Object -First 1 -ExpandProperty FullName
}
if ([string]::IsNullOrWhiteSpace([string]$generatedPath) -or -not (Test-Path -LiteralPath $generatedPath -PathType Leaf)) {
    throw 'IntuneWinAppUtil did not produce an .intunewin file.'
}
$finalPath = Join-Path $outputPath ("SmartM365-EndpointDiagnosticsAnalyzer-{0}.intunewin" -f $packageVersion)
if (-not [IO.Path]::GetFullPath($generatedPath).Equals([IO.Path]::GetFullPath($finalPath), [StringComparison]::OrdinalIgnoreCase)) {
    Move-Item -LiteralPath $generatedPath -Destination $finalPath -Force
}

[pscustomobject]@{
    Result = 'PASS'
    PackageVersion = $packageVersion
    SourcePath = $sourcePath
    SetupFile = $setupFile
    IntuneWinPath = $finalPath
    Sha256 = (Get-FileHash -LiteralPath $finalPath -Algorithm SHA256).Hash
    AutomaticUpdateEnabled = $false
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCgr7GXwPDsPdMO
# tYN1SOm69PBLAaizg1FVMTZmFxXuKaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCB1xMc6Z6xPOQexV2TqtoF+
# EDewXFzf1dv4nfnEyS0wUzANBgkqhkiG9w0BAQEFAASCAYCYBXLqW01n1ZEdqTkm
# UWSRhzN44UEK2v/huMD4AsRuSp1yzSl1LyOSFUkX5OCpQy0r8rovLaqIEXF6Edar
# yj0AdS1rKgyK8sXJfzQeF3xDxjOeiA16Zy+TsfFu4SqYKCLHxmFktB95TaNuACco
# tSInKwBcT2zbkDaJ98sAqNiHe6lOQq9P1+Bwy3sPfzKzHaDOmaTfQeJ2II6356tJ
# RDIuqWVBiwmKyKSg9eadZ22GnKpBW1RSEqfE/B8BF+5ReCJxIKgikcPjbJFM/uAb
# v9mrjVOcerNXjTJgffI8XS92BQflwAPlpxkGukrzLv3mSENZmsNYGO2pOoTFrp32
# BIdv6kMQgod9xgLxSuyAdkcKl/lsJnDL1qUMX0iwfYi5/Iz3oZjjdjfxx/GBI9bd
# pCrz3qsmfzfEBcekeKB/UOxG5TKWkBpyi4RaXFyRkkrh6JKyFc+XSMzXX//XNGAk
# z1nzw1XMsRrQUbmlnKGDGe7dZYPhXxvAep7X3aef0duLCik=
# SIG # End signature block
