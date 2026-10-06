#Requires -Version 5.1

[CmdletBinding()]
param(
    [string]$OutputRoot = 'C:\tmp',
    [string]$IntuneWinAppUtilPath = 'C:\tmp\SmartM365Tools\IntuneWinAppUtil.exe',
    [switch]$Force,
    [switch]$SkipAuthenticodeValidation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$applicationRoot = Split-Path -Parent $PSScriptRoot
$version = Get-Content -LiteralPath (Join-Path $applicationRoot 'SmartM365-EndpointDiagnosticsAnalyzer.version.json.txt') -Raw | ConvertFrom-Json
$packageVersion = [string]$version.PackageVersion
if ([string]::IsNullOrWhiteSpace($packageVersion)) { throw 'PackageVersion is missing.' }

$resolvedOutputRoot = [IO.Path]::GetFullPath($OutputRoot)
$releaseChannel = if ([string]::IsNullOrWhiteSpace([string]$version.Prerelease)) { 'Stable' } else { 'Preview' }
$bundlePath = Join-Path $resolvedOutputRoot ("SmartM365-EndpointDiagnosticsAnalyzer-{0}-{1}" -f $releaseChannel,$packageVersion)
$normalizedRoot = $resolvedOutputRoot.TrimEnd('\') + '\'
if (-not ([IO.Path]::GetFullPath($bundlePath).StartsWith($normalizedRoot,[StringComparison]::OrdinalIgnoreCase))) {
    throw "Unsafe bundle path: $bundlePath"
}
if (Test-Path -LiteralPath $bundlePath) {
    if (-not $Force) { throw "Bundle already exists: $bundlePath" }
    Remove-Item -LiteralPath $bundlePath -Recurse -Force
}
$galleryBundlePath = Join-Path $bundlePath 'GalleryPackage'
$adminBundlePath = Join-Path $bundlePath 'AdminBundle'
New-Item -ItemType Directory -Path $galleryBundlePath,$adminBundlePath -Force | Out-Null

$stagingRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365\BundleStaging\EndpointDiagnosticsAnalyzer\' + [guid]::NewGuid().ToString('N'))
try {
    $galleryBuild = & (Join-Path $applicationRoot 'PowerShellGallery\SmartM365-Build-EndpointDiagnosticsAnalyzerGalleryPackage.ps1') `
        -OutputRoot (Join-Path $stagingRoot 'Gallery') `
        -Force `
        -SkipAuthenticodeValidation:$SkipAuthenticodeValidation
    $moduleDestination = Join-Path $galleryBundlePath 'SmartM365.EndpointDiagnosticsAnalyzer\0.3.0'
    New-Item -ItemType Directory -Path (Split-Path -Parent $moduleDestination) -Force | Out-Null
    Copy-Item -LiteralPath $galleryBuild.PackagePath -Destination $moduleDestination -Recurse -Force

    $intuneBuildOutput = & (Join-Path $applicationRoot 'IntuneWin32\SmartM365-Build-EndpointDiagnosticsAnalyzerIntunePackage.ps1') `
        -IntuneWinAppUtilPath $IntuneWinAppUtilPath `
        -OutputRoot (Join-Path $stagingRoot 'Intune') `
        -GalleryBuildRoot (Join-Path $stagingRoot 'IntuneGallery') `
        -Force `
        -SkipAuthenticodeValidation:$SkipAuthenticodeValidation
    $intuneBuild = @($intuneBuildOutput | Where-Object { $_.PSObject.Properties.Name -contains 'IntuneWinPath' } | Select-Object -Last 1)
    if ($intuneBuild.Count -ne 1) { throw 'The Intune package builder did not return a result object.' }
    $intuneBuild = $intuneBuild[0]
    Copy-Item -LiteralPath $intuneBuild.IntuneWinPath -Destination $adminBundlePath -Force

    foreach ($file in @(
        'Deploy\SmartM365-EndpointDiagnosticsAnalyzer-PublishIntune.ps1',
        'Deploy\SmartM365-EndpointDiagnosticsAnalyzer-Detection.ps1',
        'PowerShellGallery\SmartM365-Publish-EndpointDiagnosticsAnalyzerGalleryPackage.ps1',
        'SmartM365-EndpointDiagnosticsAnalyzer.version.json.txt',
        'README.md'
    )) {
        Copy-Item -LiteralPath (Join-Path $applicationRoot $file) -Destination $adminBundlePath -Force
    }

    $intuneFileName = Split-Path -Leaf $intuneBuild.IntuneWinPath
    $prereleaseArgument = if ($releaseChannel -eq 'Preview') { ' -AllowPrereleasePublication' } else { '' }
    $commands = @"
PowerShell Gallery preview (no publication):
.\SmartM365-Publish-EndpointDiagnosticsAnalyzerGalleryPackage.ps1 -PackagePath ..\GalleryPackage\SmartM365.EndpointDiagnosticsAnalyzer\0.3.0

PowerShell Gallery publication (explicit execution):
.\SmartM365-Publish-EndpointDiagnosticsAnalyzerGalleryPackage.ps1 -PackagePath ..\GalleryPackage\SmartM365.EndpointDiagnosticsAnalyzer\0.3.0$prereleaseArgument -Execute

Intune preview (no Graph connection):
.\SmartM365-EndpointDiagnosticsAnalyzer-PublishIntune.ps1 -IntuneWinPath .\$intuneFileName -CreatePilotGroup -AssignPilotGroup

Intune publication with an Available pilot assignment:
.\SmartM365-EndpointDiagnosticsAnalyzer-PublishIntune.ps1 -IntuneWinPath .\$intuneFileName -CreatePilotGroup -AssignPilotGroup -Execute
"@
    [IO.File]::WriteAllText((Join-Path $adminBundlePath 'Admin-Commands.txt'),$commands,[Text.UTF8Encoding]::new($false))

    $hashes = Get-ChildItem -LiteralPath $bundlePath -Recurse -File |
        Where-Object { $_.Name -ne 'SHA256SUMS.txt' } |
        Sort-Object FullName |
        ForEach-Object {
            $relative = $_.FullName.Substring($bundlePath.Length).TrimStart('\')
            '{0}  {1}' -f (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash,$relative
        }
    [IO.File]::WriteAllLines((Join-Path $bundlePath 'SHA256SUMS.txt'),[string[]]$hashes,[Text.UTF8Encoding]::new($false))

    [pscustomobject]@{
        Result = 'PASS'
        PackageVersion = $packageVersion
        BundlePath = $bundlePath
        GalleryPackagePath = $moduleDestination
        AdminBundlePath = $adminBundlePath
        IntuneWinPath = Join-Path $adminBundlePath $intuneFileName
        IntuneWinSha256 = [string]$intuneBuild.Sha256
        AutomaticUpdateEnabled = $false
        PublicationAttempted = $false
    }
}
finally {
    if (Test-Path -LiteralPath $stagingRoot) { Remove-Item -LiteralPath $stagingRoot -Recurse -Force }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCYiDca6YN9gkFu
# PJMjFAWx1ajJtmI0aLybqApo1CTdhqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCD9svu3U+BkhQMVBUWTtVwO
# pIVTL3fql/h2LpaJ24z5ZjANBgkqhkiG9w0BAQEFAASCAYAjwxLlpGN8YV+TtzL8
# hM1t6Xcqt1M13wr6aEzyk5SN4OH6aoQx7XUeh400o/0DnETWttoSUkltDJJfmtN1
# /4Ng9C9z3Sv2Xh7/JFAnDmjQ/wkPqwyT2ta1GBk9pfA87dM7jpR0xnunLGPEb+Jc
# LL4ibo9ecKe6FAeSjJNMPd8yoV/nyP69sirDqv6MCxPasmq7HyLjqNwEVtr2jyjK
# MSCNXXXSgJ5Gvi3xgS9Ly1EQ/zpgrMr73aGceIjpwzGRoME3+vvD01bEMXyxUymE
# PSEWE2X64i93sAD3Qv0qCCJH/vtXg14+wZnI1Zw30Xk3gJzbTnlK7NHMPmk7vI/7
# 9dRK/uIfz++o7UaB6Gxns50QDaJS8Fl5OtXKgdEyPXyVZwEn2QiiqPNNmYsiLi3F
# fMYdpFtVbNUg2EbvKsijcxvcI1dYEcvefEMkrtwq59a4Jm467fdwsBLQDbEpkuU8
# 3kpb92IHRjBdBM6A6KBQyMup+iZLMRQ4/yE01BWH9jeI0F4=
# SIG # End signature block
