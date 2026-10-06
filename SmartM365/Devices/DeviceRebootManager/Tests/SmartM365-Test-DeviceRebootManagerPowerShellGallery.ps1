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

$productRoot = Split-Path -Path $PSScriptRoot -Parent
$galleryRoot = Join-Path -Path $productRoot -ChildPath 'PowerShellGallery'
$moduleSourceRoot = Join-Path -Path $galleryRoot -ChildPath 'Module'
$testRoot = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ("SmartM365-GalleryTest-{0}" -f [guid]::NewGuid().ToString('N'))

try {
    $sourcePowerShellFiles = @(Get-ChildItem -LiteralPath $galleryRoot -Recurse -File |
        Where-Object { $_.Extension -in @('.ps1', '.psm1', '.psd1') })
    Assert-True ($sourcePowerShellFiles.Count -ge 5) 'Expected Gallery source files were not found.'

    foreach ($file in $sourcePowerShellFiles) {
        $parseErrors = $null
        [void][Management.Automation.PSParser]::Tokenize(
            (Get-Content -LiteralPath $file.FullName -Raw),
            [ref]$parseErrors
        )
        Assert-True ($parseErrors.Count -eq 0) ("PowerShell syntax errors found in {0}" -f $file.FullName)
    }

    $moduleSourcePath = Join-Path -Path $moduleSourceRoot -ChildPath 'SmartM365.DeviceRebootManager.psm1'
    $moduleSourceText = Get-Content -LiteralPath $moduleSourcePath -Raw
    $updateHelperPath = Join-Path -Path $moduleSourceRoot -ChildPath 'Tools\SmartM365-DeviceRebootManager-GalleryUpdate.ps1'
    $updateHelperText = Get-Content -LiteralPath $updateHelperPath -Raw
    Assert-True ($moduleSourceText -match '\[bool\]\$EnableAutomaticUpdate\s*=\s*\$false') 'Installation must disable automatic updates by default.'
    Assert-True ($updateHelperText -match '\[bool\]\$EnableAutomaticUpdate\s*=\s*\$false') 'Update helper must disable automatic updates by default.'
    Assert-True ($moduleSourceText -match '\[Nullable\[bool\]\]\$EnableAutomaticUpdate\s*=\s*\$null') `
        'Manual update must preserve the existing automatic-update state by default.'
    Assert-True ($moduleSourceText -match '\[bool\]\$IncludePrerelease\s*=\s*\$false') `
        'Stable module must not select prerelease updates by default.'
    Assert-True ($updateHelperText -match '\[bool\]\$IncludePrerelease\s*=\s*\$false') `
        'Stable update helper must not select prerelease updates by default.'

    $scheduledTaskScriptPath = Join-Path -Path $productRoot -ChildPath 'Deploy\SmartM365-DeviceRebootManager-CreateScheduledTask.ps1'
    $scheduledTaskScriptText = Get-Content -LiteralPath $scheduledTaskScriptPath -Raw
    Assert-True ($scheduledTaskScriptText -notmatch '\.Repetition\.(Interval|Duration)\s*=') `
        'Main scheduled-task script must not mutate a missing Repetition property.'
    Assert-True ($updateHelperText -notmatch '\.Repetition\.(Interval|Duration)\s*=') `
        'Gallery update helper must not mutate a missing Repetition property.'
    Assert-True ($scheduledTaskScriptText -match '(?s)New-ScheduledTaskTrigger\s+`\s*-Once.*-RepetitionInterval.*-RepetitionDuration') `
        'Main repeat trigger must use the compatible Once parameter set.'
    Assert-True ($scheduledTaskScriptText -notmatch 'LeastPrivilege') `
        'Main scheduled-task script must use a valid RunLevel enumerator.'
    Assert-True ($scheduledTaskScriptText -notmatch '-GroupId[^\r\n]+-LogonType') `
        'Group principal must not mix incompatible GroupId and LogonType parameter sets.'
    Assert-True ($scheduledTaskScriptText -match "-GroupId 'S-1-5-32-545' -RunLevel Limited") `
        'Main task principal must use the Users group with Limited run level.'
    Assert-True ($updateHelperText -match '(?s)New-ScheduledTaskTrigger\s+`\s*-Once.*-RepetitionInterval.*-RepetitionDuration') `
        'Gallery update trigger must use the compatible Once parameter set.'

    $sourceManifestPath = Join-Path -Path $moduleSourceRoot -ChildPath 'SmartM365.DeviceRebootManager.psd1'
    $sourceManifest = Test-ModuleManifest -Path $sourceManifestPath
    Assert-True ($sourceManifest.Name -eq 'SmartM365.DeviceRebootManager') 'Unexpected module name.'
    $sourcePrerelease = [string]$sourceManifest.PrivateData.PSData['Prerelease']
    Assert-True ([string]::IsNullOrWhiteSpace($sourcePrerelease)) 'Stable manifest must not contain a prerelease label.'

    $versionManifestPath = Join-Path -Path $productRoot -ChildPath 'SmartM365-DeviceRebootManager.version.json.txt'
    $versionManifest = Get-Content -LiteralPath $versionManifestPath -Raw | ConvertFrom-Json
    Assert-True ([int]$versionManifest.SchemaVersion -eq 1) 'Unexpected package version schema.'
    Assert-True ([string]$versionManifest.PackageVersion -eq '0.1.0') 'Version manifest and stable module version differ.'

    $buildScript = Join-Path -Path $galleryRoot -ChildPath 'SmartM365-Build-DeviceRebootManagerGalleryPackage.ps1'
    $buildResult = & $buildScript `
        -OutputRoot $testRoot `
        -Force `
        -SkipAuthenticodeValidation:$SkipAuthenticodeValidation
    Assert-True $buildResult.Ready 'Build did not report Ready.'
    Assert-True ($buildResult.PackageVersion -eq '0.1.0') 'Stable build package version is incorrect.'
    Assert-True (Test-Path -LiteralPath $buildResult.IntuneSourcePath -PathType Container) 'Intune source path was not created.'

    $requiredPackagePaths = @(
        'SmartM365.DeviceRebootManager.psd1'
        'SmartM365.DeviceRebootManager.psm1'
        'Tools\SmartM365-DeviceRebootManager-GalleryUpdate.ps1'
        'Runtime\SmartM365-DeviceRebootManager.version.json.txt'
        'Runtime\SmartM365-DeviceRebootManager-GUI.ps1'
        'Runtime\SmartM365-DeviceRebootManager-GUI.config.json.template'
        'Runtime\Deploy\SmartM365-DeviceRebootManager-Install.ps1'
        'Runtime\Deploy\SmartM365-DeviceRebootManager-Uninstall.ps1'
    )
    foreach ($relativePath in $requiredPackagePaths) {
        Assert-True (Test-Path -LiteralPath (Join-Path $buildResult.PackagePath $relativePath) -PathType Leaf) `
            ("Required package file is missing: {0}" -f $relativePath)
    }

    $forbiddenNames = @(
        'LOCAL_MEMORY.md'
        'SmartM365-DeviceRebootManager-GUI.config.json'
    )
    $packagedFiles = @(Get-ChildItem -LiteralPath $buildResult.PackagePath -Recurse -File)
    foreach ($forbiddenName in $forbiddenNames) {
        Assert-True (-not ($packagedFiles.Name -contains $forbiddenName)) ("Private file was packaged: {0}" -f $forbiddenName)
    }
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $buildResult.PackagePath 'Logs'))) 'Logs directory was packaged.'

    $vmValidationScript = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-Test-DeviceRebootManagerGalleryVm.ps1'
    $vmPreview = & $vmValidationScript -PackagePath $buildResult.PackagePath
    Assert-True ($vmPreview.Mode -eq 'Preview') 'VM validation script did not stay in Preview mode.'
    Assert-True (-not $vmPreview.ChangesAttempted) 'VM preview unexpectedly attempted changes.'
    Assert-True (-not $vmPreview.AutomaticUpdateDefault) 'VM preview does not report automatic updates disabled by default.'
    Assert-True $vmPreview.AllPinnedSignerMatches 'VM preview did not validate the pinned package signer.'
    Assert-True $vmPreview.ScheduledTaskApiCompatible 'VM preview ScheduledTasks API compatibility preflight failed.'

    $builtManifestPath = Join-Path $buildResult.PackagePath 'SmartM365.DeviceRebootManager.psd1'
    Import-Module -Name $builtManifestPath -Force
    $exportedNames = @((Get-Module SmartM365.DeviceRebootManager).ExportedFunctions.Keys)
    foreach ($expectedCommand in @(
        'Get-SmartM365DeviceRebootManager'
        'Install-SmartM365DeviceRebootManager'
        'Uninstall-SmartM365DeviceRebootManager'
        'Update-SmartM365DeviceRebootManager'
    )) {
        Assert-True ($exportedNames -contains $expectedCommand) ("Exported command is missing: {0}" -f $expectedCommand)
    }

    $fakeInstallPath = Join-Path -Path $testRoot -ChildPath 'NotInstalled'
    $status = Get-SmartM365DeviceRebootManager -InstallPath $fakeInstallPath
    Assert-True (-not $status.Installed) 'Status should report an absent installation.'

    $publishScript = Join-Path -Path $galleryRoot -ChildPath 'SmartM365-Publish-DeviceRebootManagerGalleryPackage.ps1'
    $preview = & $publishScript -PackagePath $buildResult.PackagePath
    Assert-True ($preview.Mode -eq 'Preview') 'Publish script did not stay in Preview mode.'
    Assert-True (-not $preview.PublicationAttempted) 'Preview unexpectedly attempted publication.'
    Assert-True $preview.PublicMetadataComplete 'Stable package must include approved public license metadata.'

    $missingApiKeyVariable = "SMARTM365_TEST_MISSING_API_KEY_$([guid]::NewGuid().ToString('N'))"
    $publicationGuardTriggered = $false
    try {
        & $publishScript `
            -PackagePath $buildResult.PackagePath `
            -ApiKeyEnvironmentVariable $missingApiKeyVariable `
            -Execute `
            -Confirm:$false | Out-Null
    }
    catch {
        $publicationGuardTriggered = $_.Exception.Message -like '*API key environment variable is empty*'
        if (-not $publicationGuardTriggered) { throw }
    }
    Assert-True $publicationGuardTriggered 'Execute mode was not blocked when the API key was absent.'

    [pscustomobject]@{
        Result                 = 'PASS'
        ModuleName             = $buildResult.ModuleName
        Version                = $buildResult.Version
        Prerelease             = $buildResult.Prerelease
        PackageFileCount       = $buildResult.FileCount
        PowerShellFileCount    = $buildResult.PowerShellFileCount
        SignaturesValidated    = $buildResult.SignaturesValidated
        PublicationAttempted   = $preview.PublicationAttempted
        PublicMetadataComplete = $preview.PublicMetadataComplete
        PublicationGuard       = $publicationGuardTriggered
        VmPreviewChangesAttempted = $vmPreview.ChangesAttempted
    }
}
finally {
    Remove-Module -Name SmartM365.DeviceRebootManager -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $testRoot) {
        $resolvedTestRoot = [IO.Path]::GetFullPath($testRoot)
        $resolvedTempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\', '/')
        if ($resolvedTestRoot.StartsWith($resolvedTempRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
            Remove-Item -LiteralPath $resolvedTestRoot -Recurse -Force
        }
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCd2YI4HrrUW56u
# +tyaHVSN0CMrys1dD9cyAg5DAvrPKqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCTnpvnP7JbtnbkbKqaCn0r
# crFBu8/BV3nu2LlmbGiLQDANBgkqhkiG9w0BAQEFAASCAYCJRD9vrCxya3aPf1JL
# l9Wx6BCMmVNhictlA3CQTT/DU8z+FuEvud+9m3izYzck0DdHOzuoNG1rveZ4ZRw+
# KDHtT/T1x9O49QkZRUKgK+lOCRoOuCnQtz0SycEIi2KZX1ZQUfpY9UISv+6kOQyn
# /rVBmC0ltZDtulRXpBv/x1ySMo71STj8g+JDKcpGNZp2lUZgSa2I6ExezysLpI1N
# zDJQmuyr0poCYKePZfFtmXkto2SAYXOuW/jfR8rFA3upSDRX0G7BXS2e0piSYbxt
# /f3rjSLSWNhg5VhXOb2t/OhsoyaTi1piWRydXlhSm+aQudhuU9OzHJhkBsKqx3Dc
# TrmqymVBpKJTatSTEb4GlIP96N4gi6U+xzLan+jorzkAJFgnkBYxdHUALAkA1UJm
# LpX/gpWeACQcw1fxNWncg2pXpUVqFS646WUG94vuXWuUa/2aataLTGhTp21JSIOv
# 8vSd/H9WCoU953fXE32QUyYXO2gXPeZ5jkMkWwtucMAnfOw=
# SIG # End signature block
