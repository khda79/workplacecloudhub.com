#Requires -Version 5.1

[CmdletBinding()]
param(
    [switch]$SkipAuthenticodeValidation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-True {
    param(
        [Parameter(Mandatory = $true)][bool]$Condition,
        [Parameter(Mandatory = $true)][string]$Message
    )

    if (-not $Condition) {
        throw "Assertion failed: $Message"
    }
}

function Get-RelativeFilePaths {
    param([Parameter(Mandatory = $true)][string]$Root)

    return @(
        Get-ChildItem -LiteralPath $Root -Recurse -File |
            ForEach-Object { $_.FullName.Substring($Root.Length).TrimStart('\', '/') } |
            Sort-Object
    )
}

$productRoot = Split-Path -Path $PSScriptRoot -Parent
$builderPath = Join-Path -Path $productRoot -ChildPath 'Release\SmartM365-Build-DeviceRebootManagerRelease.ps1'
$testRoot = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath (
    'SmartM365-DeviceRebootManager-ReleaseTest-{0}' -f [guid]::NewGuid().ToString('N')
)

$expectedPackagePaths = @(
    'Deploy\SmartM365-DeviceRebootManager-CreateScheduledTask.ps1'
    'Deploy\SmartM365-DeviceRebootManager-Detection.ps1'
    'Deploy\SmartM365-DeviceRebootManager-Install.ps1'
    'Deploy\SmartM365-DeviceRebootManager-PublishIntune.ps1'
    'Deploy\SmartM365-DeviceRebootManager-Uninstall.ps1'
    'LICENSE'
    'NOTICE'
    'PowerShellGallery\Module\SmartM365.DeviceRebootManager.psd1'
    'PowerShellGallery\Module\SmartM365.DeviceRebootManager.psm1'
    'PowerShellGallery\Module\Tools\SmartM365-DeviceRebootManager-GalleryUpdate.ps1'
    'PowerShellGallery\SmartM365-Build-DeviceRebootManagerGalleryPackage.ps1'
    'PowerShellGallery\SmartM365-New-DeviceRebootManagerGalleryVmBundle.ps1'
    'PowerShellGallery\SmartM365-Publish-DeviceRebootManagerGalleryPackage.ps1'
    'README.md'
    'SmartM365-DeviceRebootManager-GUI.config.json.template'
    'SmartM365-DeviceRebootManager-GUI.ps1'
    'SmartM365-DeviceRebootManager-GUI.strings.psd1'
    'SmartM365-DeviceRebootManager.version.json.txt'
    'SmartM365.GuiSplash.ps1'
    'Start-SmartM365-DeviceRebootManager-GUI-Test.cmd'
    'Start-SmartM365-DeviceRebootManager-GUI.cmd'
    'Tests\SmartM365-Test-DeviceRebootManagerGalleryVm.ps1'
    'Tests\SmartM365-Test-DeviceRebootManagerIntunePublisher.ps1'
    'Tests\SmartM365-Test-DeviceRebootManagerPowerShellGallery.ps1'
    'WorkplaceCloudHub-lockup-WPF.png'
    'WorkplaceCloudHub.ico'
) | Sort-Object

try {
    $previewRoot = Join-Path -Path $testRoot -ChildPath 'Preview'
    $preview = & $builderPath `
        -OutputRoot $previewRoot `
        -SkipAuthenticodeValidation:$SkipAuthenticodeValidation

    Assert-True ($preview.Mode -eq 'Preview') 'Default mode must be Preview.'
    Assert-True (-not $preview.BuildAttempted) 'Preview unexpectedly attempted a build.'
    Assert-True (-not $preview.PublicationAttempted) 'Preview unexpectedly attempted publication.'
    Assert-True ($preview.FileCount -eq 26) 'Preview allow-list file count is incorrect.'
    Assert-True ($preview.PowerShellFileCount -eq 17) 'Preview PowerShell file count is incorrect.'
    Assert-True $preview.ReadmeValidated 'Preview did not validate README references.'
    Assert-True $preview.DetectionVersionValidated 'Preview did not validate detection version consistency.'
    Assert-True (-not (Test-Path -LiteralPath $previewRoot)) 'Preview created an output directory.'

    $buildRoot = Join-Path -Path $testRoot -ChildPath 'Build'
    $build = & $builderPath `
        -OutputRoot $buildRoot `
        -Build `
        -SkipAuthenticodeValidation:$SkipAuthenticodeValidation

    Assert-True ($build.Mode -eq 'LocalBuild') 'Build mode is incorrect.'
    Assert-True $build.BuildAttempted 'Build did not report an attempted local build.'
    Assert-True (-not $build.PublicationAttempted) 'Local build unexpectedly attempted publication.'
    Assert-True ($build.PackageVersion -eq '0.1.0') 'Unexpected package version.'
    Assert-True ($build.ExpectedReleaseTag -eq 'device-reboot-manager-v0.1.0') 'Unexpected release tag.'
    Assert-True ($build.FileCount -eq 26) 'Built package file count is incorrect.'
    Assert-True ($build.PowerShellFileCount -eq 17) 'Built package PowerShell file count is incorrect.'
    Assert-True (-not $build.IntuneWinIncluded) 'Build unexpectedly included an IntuneWin.'
    Assert-True (Test-Path -LiteralPath $build.ZipPath -PathType Leaf) 'ZIP was not created.'
    Assert-True (Test-Path -LiteralPath $build.ChecksumPath -PathType Leaf) 'Checksum file was not created.'
    Assert-True ([string]::IsNullOrWhiteSpace($build.IntuneWinPath)) 'Build reported an unexpected IntuneWin path.'
    Assert-True ((Get-FileHash -LiteralPath $build.ZipPath -Algorithm SHA256).Hash -eq $build.ZipSHA256) `
        'Reported ZIP hash is incorrect.'

    $checksumText = Get-Content -LiteralPath $build.ChecksumPath -Raw
    Assert-True ($checksumText -match ("(?m)^{0}  {1}$" -f
        [regex]::Escape($build.ZipSHA256),
        [regex]::Escape([IO.Path]::GetFileName($build.ZipPath)))) `
        'Checksum file does not contain the ZIP hash.'
    Assert-True ($checksumText -notmatch '\.intunewin') 'Checksum file unexpectedly references an IntuneWin.'

    $expandedRoot = Join-Path -Path $testRoot -ChildPath 'Expanded'
    Expand-Archive -LiteralPath $build.ZipPath -DestinationPath $expandedRoot
    $packageRoot = Join-Path -Path $expandedRoot -ChildPath 'SmartM365-DeviceRebootManager-0.1.0'
    Assert-True (Test-Path -LiteralPath $packageRoot -PathType Container) 'Expected ZIP root folder is missing.'

    $actualPackagePaths = Get-RelativeFilePaths -Root $packageRoot
    Assert-True ($actualPackagePaths.Count -eq 26) 'ZIP does not contain exactly 26 files.'
    Assert-True (@(Compare-Object -ReferenceObject $expectedPackagePaths -DifferenceObject $actualPackagePaths).Count -eq 0) `
        'ZIP paths differ from the independently defined expected package topology.'

    $sourceReadme = Join-Path -Path $productRoot -ChildPath 'README.md'
    $packagedReadme = Join-Path -Path $packageRoot -ChildPath 'README.md'
    Assert-True ((Get-FileHash -LiteralPath $sourceReadme -Algorithm SHA256).Hash -eq
        (Get-FileHash -LiteralPath $packagedReadme -Algorithm SHA256).Hash) `
        'Packaged README does not exactly match the source README.'
    $packagedReadmeText = Get-Content -LiteralPath $packagedReadme -Raw
    Assert-True ($packagedReadmeText -notmatch '\.\\Devices\\DeviceRebootManager\\') `
        'Packaged README contains obsolete repository-relative standalone commands.'

    $secondBuildBlocked = $false
    try {
        & $builderPath `
            -OutputRoot $buildRoot `
            -Build `
            -SkipAuthenticodeValidation:$SkipAuthenticodeValidation | Out-Null
    }
    catch {
        $secondBuildBlocked = $_.Exception.Message -like '*Release output already exists*'
        if (-not $secondBuildBlocked) { throw }
    }
    Assert-True $secondBuildBlocked 'Builder did not guard existing release artifacts.'

    $forceBuild = & $builderPath `
        -OutputRoot $buildRoot `
        -Build `
        -Force `
        -SkipAuthenticodeValidation:$SkipAuthenticodeValidation
    Assert-True $forceBuild.Ready 'Force rebuild did not complete successfully.'
    Assert-True (-not $forceBuild.PublicationAttempted) 'Force rebuild unexpectedly attempted publication.'

    [pscustomobject]@{
        Result                    = 'PASS'
        PreviewMode               = $preview.Mode
        PreviewCreatedOutput      = (Test-Path -LiteralPath $previewRoot)
        PackageVersion            = $build.PackageVersion
        PackageFileCount          = $build.FileCount
        PowerShellFileCount       = $build.PowerShellFileCount
        ReadmeExactMatch          = $true
        ExistingArtifactGuard     = $secondBuildBlocked
        ForceRebuildReady         = $forceBuild.Ready
        PublicationAttempted      = $forceBuild.PublicationAttempted
        SignaturesValidated       = $forceBuild.SignaturesValidated
    }
}
finally {
    if (Test-Path -LiteralPath $testRoot) {
        $resolvedTestRoot = [IO.Path]::GetFullPath($testRoot)
        $resolvedTempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\', '/')
        $expectedPrefix = $resolvedTempRoot + [IO.Path]::DirectorySeparatorChar
        if ($resolvedTestRoot.StartsWith($expectedPrefix, [StringComparison]::OrdinalIgnoreCase) -and
            [IO.Path]::GetFileName($resolvedTestRoot) -like 'SmartM365-DeviceRebootManager-ReleaseTest-*') {
            Remove-Item -LiteralPath $resolvedTestRoot -Recurse -Force
        }
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD5nJw0cXewgXu0
# V0+BYLv9xypXdCML9v8TI3+RyujhZaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBXi/ekftR2LoVURP3e1d+R
# pTUb3lvpZmVWC2AQqaVJ2DANBgkqhkiG9w0BAQEFAASCAYAI5czrhSn16WM89BHq
# DyHdvMuyyFz/Z9hyD7bG25T/xAEpieS4Zw8XzUSdMk02qaKwzOqyV4i5glhvH6e1
# LnpPiOGo3UoZLjZ5ghIbV+Bb7kgFUDHGIozOJ/YM7qWtwnjwGLoZTo5TpN0YW0+U
# TykqC1Tfa0yAgEYRxuCD999DNbrdn9x88e6iupAiaLLcEyoXQN+OsrGN09+m6vKY
# cQORuPhVXZClGobjkgIde5+vwpDqZZB/UcA56EIa7rskx+XpFhXG0sDSgcWlgTLB
# RGYJeWGk7DSQEjUAnA1MFCLZnT+pePcxhep3cJioJdyCazp9qiPidpz9bmWbSEk+
# jCk5qFT21QLqmpoHSU/8V3htog29p05C8KjjVUZZft7dBMAnCToNXetgzbzcXpN/
# fBGyTBvgaVHTs6MzlqsXQhDm1IIe6pjwJyrcauOTvl9WCbaE6AL8NIPvgJk+BWxK
# V5/KXKbxTipxce//nD4VHx40SSplvnXWVg2zJw/mLCvYAvU=
# SIG # End signature block
