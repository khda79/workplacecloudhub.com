#Requires -Version 5.1

[CmdletBinding()]
param(
    [switch]$KeepTestArtifacts,
    [switch]$SkipAuthenticodeValidation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Assert-True {
    param([bool]$Condition,[string]$Message)
    if (-not $Condition) { throw "Assertion failed: $Message" }
}

$applicationRoot = Split-Path -Parent $PSScriptRoot
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365\Tests\EndpointDiagnosticsAnalyzer\' + [guid]::NewGuid().ToString('N'))
$galleryRoot = Join-Path $testRoot 'Gallery'
$installPath = Join-Path $testRoot 'Install'
$module = $null
try {
    $version = Get-Content -LiteralPath (Join-Path $applicationRoot 'SmartM365-EndpointDiagnosticsAnalyzer.version.json.txt') -Raw | ConvertFrom-Json
    Assert-True ([string]$version.PackageVersion -eq '0.3.0') 'Unexpected package version.'

    $guiPath = Join-Path $applicationRoot 'SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1'
    $guiContent = Get-Content -LiteralPath $guiPath -Raw
    Assert-True ($guiContent -match "AppVersion = '0\.3\.0'") 'GUI version is not aligned.'
    Assert-True ($guiContent -match 'LocalDataRoot = Join-Path .*LOCALAPPDATA.*SmartM365\\EndpointDiagnosticsAnalyzer') 'Local data root is not under LocalAppData.'
    Assert-True ($guiContent -match 'api_key_dpapi') 'DPAPI-backed AI configuration is missing.'
    Assert-True ($guiContent -notmatch 'api_key\s*=\s*\$Config\.ApiKey') 'Plaintext AI key persistence is present.'
    Assert-True ($guiContent -match "entry\.Key -in @\('ValidateOnly', 'AIApiKey'\)") 'Plaintext API key is not excluded from elevation arguments.'

    $scripts = @(Get-ChildItem -LiteralPath $applicationRoot -Recurse -File | Where-Object { $_.Extension -in @('.ps1','.psm1','.psd1') })
    foreach ($scriptFile in $scripts) {
        $tokens = $null
        $errors = $null
        [void][Management.Automation.Language.Parser]::ParseFile($scriptFile.FullName,[ref]$tokens,[ref]$errors)
        Assert-True (@($errors).Count -eq 0) ("PowerShell parse error in {0}: {1}" -f $scriptFile.FullName,(@($errors | ForEach-Object { $_.Message }) -join ' | '))
    }
    $scheduledTaskReferences = @($scripts | Where-Object { $_.FullName -notlike (Join-Path $PSScriptRoot '*') -and (Get-Content -LiteralPath $_.FullName -Raw) -match '(?i)Register-ScheduledTask|New-ScheduledTask' })
    Assert-True ($scheduledTaskReferences.Count -eq 0) 'The application package must not create scheduled tasks.'

    $builder = Join-Path $applicationRoot 'PowerShellGallery\SmartM365-Build-EndpointDiagnosticsAnalyzerGalleryPackage.ps1'
    $build = & $builder -OutputRoot $galleryRoot -Force -SkipAuthenticodeValidation:$SkipAuthenticodeValidation
    Assert-True ([bool]$build.Ready) 'Gallery build is not ready.'
    Assert-True ([string]$build.PackageVersion -eq [string]$version.PackageVersion) 'Gallery build version mismatch.'

    $module = Import-Module -Name (Join-Path $build.PackagePath 'SmartM365.EndpointDiagnosticsAnalyzer.psd1') -Force -PassThru
    foreach ($commandName in @(
        'Get-SmartM365EndpointDiagnosticsAnalyzer',
        'Install-SmartM365EndpointDiagnosticsAnalyzer',
        'Update-SmartM365EndpointDiagnosticsAnalyzer',
        'Uninstall-SmartM365EndpointDiagnosticsAnalyzer'
    )) {
        Assert-True ($null -ne (Get-Command -Name $commandName -Module $module.Name -ErrorAction SilentlyContinue)) "Missing exported command: $commandName"
    }

    $install = Install-SmartM365EndpointDiagnosticsAnalyzer -InstallScope CurrentUser -InstallPath $installPath -SkipShortcut -Confirm:$false
    Assert-True ([string]$install.Result -eq 'PASS') 'Gallery installation failed.'
    Assert-True (-not [bool]$install.AutomaticUpdateEnabled) 'Automatic updates must be disabled by default.'
    $status = Get-SmartM365EndpointDiagnosticsAnalyzer -InstallScope CurrentUser -InstallPath $installPath
    Assert-True ([bool]$status.Installed) 'Installed runtime was not detected by the module.'
    Assert-True ([string]$status.PackageVersion -eq [string]$version.PackageVersion) 'Installed runtime version mismatch.'

    $detection = Join-Path $build.IntuneSourcePath 'Deploy\SmartM365-EndpointDiagnosticsAnalyzer-Detection.ps1'
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $detection -InstallPath $installPath -ExpectedVersion ([string]$version.PackageVersion) -ExpectedPackageSource PowerShellGallery | Out-Null
    Assert-True ($LASTEXITCODE -eq 0) 'Detection script rejected the controlled Gallery installation.'

    # Preflight rejection must leave the existing installation byte-for-byte intact.
    $installedGui=Join-Path $installPath 'SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1'
    $before=(Get-FileHash -LiteralPath $installedGui).Hash
    $broken=Join-Path $testRoot 'BrokenRuntime'
    Copy-Item -LiteralPath $build.IntuneSourcePath -Destination $broken -Recurse
    $asset=Join-Path $broken 'WorkplaceCloudHub.ico'
    [IO.File]::WriteAllBytes($asset,[byte[]]@(0))
    $rejected=$false
    try { & (Join-Path $broken 'Deploy\SmartM365-EndpointDiagnosticsAnalyzer-Install.ps1') -InstallScope CurrentUser -InstallPath $installPath -SkipShortcut | Out-Null } catch { $rejected=$_.Exception.Message -match 'hash|integrity|signature' }
    Assert-True ($rejected -and (Get-FileHash -LiteralPath $installedGui).Hash -eq $before) 'Broken package modified the previous installation.'

    # Simulate activation failing after the previous tree has been moved aside.
    function Move-Item {
        param([string]$LiteralPath,[string]$Destination)
        if($LiteralPath -like ($installPath+'.stage-*') -and $Destination -eq $installPath){throw 'Synthetic activation failure'}
        Microsoft.PowerShell.Management\Move-Item -LiteralPath $LiteralPath -Destination $Destination
    }
    $activationRejected=$false
    try { & (Join-Path $build.IntuneSourcePath 'Deploy\SmartM365-EndpointDiagnosticsAnalyzer-Install.ps1') -InstallScope CurrentUser -InstallPath $installPath -PackageSource PowerShellGallery -SkipShortcut | Out-Null } catch { $activationRejected=$_.Exception.Message -match 'Synthetic activation failure' }
    finally { Remove-Item Function:\Move-Item }
    Assert-True ($activationRejected -and (Get-FileHash -LiteralPath $installedGui).Hash -eq $before) 'Failed activation did not restore the previous installation.'

    # Detection and the public status command must reject damaged payloads.
    $installedAsset=Join-Path $installPath 'WorkplaceCloudHub.ico'
    $assetBytes=[IO.File]::ReadAllBytes($installedAsset)
    [IO.File]::WriteAllBytes($installedAsset,[byte[]]@(0))
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $detection -InstallPath $installPath -ExpectedVersion ([string]$version.PackageVersion) -ExpectedPackageSource PowerShellGallery | Out-Null
    Assert-True ($LASTEXITCODE -ne 0) 'Detection accepted a corrupt runtime.'
    Assert-True (-not(Get-SmartM365EndpointDiagnosticsAnalyzer -InstallScope CurrentUser -InstallPath $installPath).Installed) 'Status accepted a corrupt runtime.'
    [IO.File]::WriteAllBytes($installedAsset,$assetBytes)

    # Intune ownership must stop a Gallery update before any repository operation.
    $metadataPath=Join-Path $installPath 'SmartM365-EndpointDiagnosticsAnalyzer.installation.json'
    $metadataText=[IO.File]::ReadAllText($metadataPath)
    $metadata=$metadataText|ConvertFrom-Json;$metadata.PackageSource='Intune'
    [IO.File]::WriteAllText($metadataPath,($metadata|ConvertTo-Json -Depth 5),[Text.Encoding]::UTF8)
    $ownedRejected=$false
    try { Update-SmartM365EndpointDiagnosticsAnalyzer -InstallScope CurrentUser -InstallPath $installPath -Confirm:$false | Out-Null } catch { $ownedRejected=$_.Exception.Message -match 'Intune' }
    Assert-True $ownedRejected 'Gallery update did not reject the Intune-owned target.'
    [IO.File]::WriteAllText($metadataPath,$metadataText,[Text.Encoding]::UTF8)

    Uninstall-SmartM365EndpointDiagnosticsAnalyzer -InstallScope CurrentUser -InstallPath $installPath -Confirm:$false | Out-Null
    Assert-True (-not (Test-Path -LiteralPath $installPath)) 'Controlled installation was not removed.'

    [pscustomobject]@{
        Result = 'PASS'
        PackageVersion = [string]$version.PackageVersion
        PowerShellEdition = $PSVersionTable.PSEdition
        PowerShellVersion = [string]$PSVersionTable.PSVersion
        ParsedFileCount = $scripts.Count
        GalleryPackagePath = [string]$build.PackagePath
        DefaultAutomaticUpdate = $false
        ScheduledTaskCreated = $false
        UserDataRemovalRequested = $false
    }
}
finally {
    if ($module) { Remove-Module -ModuleInfo $module -Force -ErrorAction SilentlyContinue }
    if (-not $KeepTestArtifacts -and (Test-Path -LiteralPath $testRoot)) {
        $resolved=[IO.Path]::GetFullPath($testRoot)
        $allowed=[IO.Path]::GetFullPath((Join-Path ([IO.Path]::GetTempPath()) 'SmartM365\Tests\EndpointDiagnosticsAnalyzer')).TrimEnd('\')+'\'
        if(-not$resolved.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path -Leaf $resolved) -notmatch '^[0-9a-f]{32}$'){throw 'Unsafe fixture cleanup target'}
        Remove-Item -LiteralPath $testRoot -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC81SXHIIimC9Bd
# 3cXZdw4ScPNxa16y4R1tuEZa+u8v46CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAhUF2LibedtcO5bC+3/9g2
# +pGuHXa3HvxBVqXKL/2/XzANBgkqhkiG9w0BAQEFAASCAYCdal1fuYVlw0zxtElT
# wLLlU01EQq2xIpomFB4SyTWaj25Le7BHj+kliP+PYXRGfO3o+vCfPYYhQcwIjyMT
# NVF9/D53jM6rvLsZDVCZXl+wGzYCtFarEcbxmj5er1kC5KKjDUg+0nHMDV+QBIGK
# zAiWsj29+hZksPE+apV6sivURAfKcogmdZwCJySyBni3qvMLzB+Sy+E3F1U5En55
# J7InzNP8gu/l5Lq5gpv5Iklu8ovI07z1xpK7AzNZMmGhhlXKvxCWFYURZk335xEw
# mfLSq8JBCM75cRR/PPczwmF+aNGjPiXWYPqWldIhgVy7RwYo8zCTcmUYlbrfQhdF
# Hf6uk/GrtiLaLA3QKbIgppt+4XPtR/j1v1N2uAW5Q2D5zNtm8HjgGSNb1tMTin6Q
# HEAMCvgSCDsvpQyWzI31iRvVGsMF0eeUAyzMJhzdqoKUo+GOkfr4cEwnNsCskWVh
# LrGM7jGDYckSFiznQQhUy1+86PmNGSJoFK9vbhtXhfLD988=
# SIG # End signature block
