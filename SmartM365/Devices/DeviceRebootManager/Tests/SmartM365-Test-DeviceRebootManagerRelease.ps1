#Requires -Version 5.1
<#
.SYNOPSIS
Offline tests for the Device Reboot Manager standalone release builder (preview, ZIP build, IntuneWin validation).
.DESCRIPTION
Runs the release builder in preview and local build modes into a temporary
folder and checks the allow-listed ZIP topology, forward-slash entry names, the
checksum file, the packaged demo and the existing-artifact guard. Synthetic
IntuneWin packages (real AES-256-CBC + HMAC-SHA256 layout) verify that a package
matching the current runtime is accepted and that stale or corrupted packages
are refused. The expected version is read from the version manifest.
.VERSION
1.1.0
#>

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

function New-TestIntuneWin {
    # Builds an IntuneWin with the IntuneWinAppUtil layout: outer ZIP with
    # IntuneWinPackage/Metadata/Detection.xml and IntuneWinPackage/Contents/
    # IntunePackage.intunewin = HMAC-SHA256 | IV | AES-256-CBC(inner ZIP).
    param(
        [Parameter(Mandatory = $true)][string]$SourceRoot,
        [Parameter(Mandatory = $true)][string[]]$RelativePaths,
        [Parameter(Mandatory = $true)][string]$DestinationPath,
        [hashtable]$ContentOverrides = @{},
        [switch]$CorruptCiphertext
    )

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $innerStream = New-Object IO.MemoryStream
    $innerZip = New-Object IO.Compression.ZipArchive($innerStream, [IO.Compression.ZipArchiveMode]::Create, $true)
    try {
        foreach ($relativePath in $RelativePaths) {
            $entry = $innerZip.CreateEntry($relativePath.Replace('\', '/'))
            $bytes = if ($ContentOverrides.ContainsKey($relativePath)) {
                [Text.Encoding]::UTF8.GetBytes([string]$ContentOverrides[$relativePath])
            }
            else {
                [IO.File]::ReadAllBytes((Join-Path -Path $SourceRoot -ChildPath $relativePath))
            }
            $entryStream = $entry.Open()
            try { $entryStream.Write($bytes, 0, $bytes.Length) } finally { $entryStream.Dispose() }
        }
    }
    finally {
        $innerZip.Dispose()
    }
    $plain = $innerStream.ToArray()

    $aes = [Security.Cryptography.Aes]::Create()
    $aes.KeySize = 256
    $aes.GenerateKey()
    $aes.GenerateIV()
    $encryptor = $aes.CreateEncryptor()
    $cipher = $encryptor.TransformFinalBlock($plain, 0, $plain.Length)
    $macKey = New-Object byte[] 32
    $random = [Security.Cryptography.RandomNumberGenerator]::Create()
    $random.GetBytes($macKey)
    $ivAndCipher = New-Object byte[] (16 + $cipher.Length)
    [Array]::Copy($aes.IV, 0, $ivAndCipher, 0, 16)
    [Array]::Copy($cipher, 0, $ivAndCipher, 16, $cipher.Length)
    $hmac = New-Object Security.Cryptography.HMACSHA256(,$macKey)
    $mac = $hmac.ComputeHash($ivAndCipher)
    $encrypted = New-Object byte[] (32 + $ivAndCipher.Length)
    [Array]::Copy($mac, 0, $encrypted, 0, 32)
    [Array]::Copy($ivAndCipher, 0, $encrypted, 32, $ivAndCipher.Length)
    if ($CorruptCiphertext) {
        $encrypted[$encrypted.Length - 1] = $encrypted[$encrypted.Length - 1] -bxor 0xFF
    }
    $digest = [Security.Cryptography.SHA256]::Create().ComputeHash($plain)

    $detectionXml = @(
        '<?xml version="1.0" encoding="utf-8"?>'
        '<ApplicationInfo xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ToolVersion="1.8.6.0">'
        '  <Name>SmartM365-DeviceRebootManager-Install.ps1</Name>'
        ('  <UnencryptedContentSize>{0}</UnencryptedContentSize>' -f $plain.Length)
        '  <FileName>IntunePackage.intunewin</FileName>'
        '  <SetupFile>Deploy\SmartM365-DeviceRebootManager-Install.ps1</SetupFile>'
        '  <EncryptionInfo>'
        ('    <EncryptionKey>{0}</EncryptionKey>' -f [Convert]::ToBase64String($aes.Key))
        ('    <MacKey>{0}</MacKey>' -f [Convert]::ToBase64String($macKey))
        ('    <InitializationVector>{0}</InitializationVector>' -f [Convert]::ToBase64String($aes.IV))
        ('    <Mac>{0}</Mac>' -f [Convert]::ToBase64String($mac))
        '    <ProfileIdentifier>ProfileVersion1</ProfileIdentifier>'
        ('    <FileDigest>{0}</FileDigest>' -f [Convert]::ToBase64String($digest))
        '    <FileDigestAlgorithm>SHA256</FileDigestAlgorithm>'
        '  </EncryptionInfo>'
        '</ApplicationInfo>'
    ) -join "`r`n"
    $aes.Dispose()

    New-Item -ItemType Directory -Path (Split-Path -Path $DestinationPath -Parent) -Force | Out-Null
    $outer = [IO.Compression.ZipFile]::Open($DestinationPath, [IO.Compression.ZipArchiveMode]::Create)
    try {
        $contentEntry = $outer.CreateEntry('IntuneWinPackage/Contents/IntunePackage.intunewin')
        $stream = $contentEntry.Open()
        try { $stream.Write($encrypted, 0, $encrypted.Length) } finally { $stream.Dispose() }
        $metadataEntry = $outer.CreateEntry('IntuneWinPackage/Metadata/Detection.xml')
        $metadataBytes = [Text.Encoding]::UTF8.GetBytes($detectionXml)
        $stream = $metadataEntry.Open()
        try { $stream.Write($metadataBytes, 0, $metadataBytes.Length) } finally { $stream.Dispose() }
    }
    finally {
        $outer.Dispose()
    }
}

function Assert-PackagedDemo {
    param([string]$RuntimeRoot)
    $entry=(Join-Path $RuntimeRoot 'SmartM365-DeviceRebootManager-GUI.ps1').Replace("'","''")
    $command='$ErrorActionPreference=''Stop''; $state=. '''+$entry+''' -Demo -ValidateOnly; if($state.Mode -ne ''Demo'' -or $null -eq $state.Window -or $state.Controls.Count -eq 0){throw ''Packaged demo did not return a window and controls''}; $state.Window.Close(); Write-Output ''DEMO-PASS'''
    $encoded=[Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($command))
    $output=& powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -EncodedCommand $encoded 2>&1
    Assert-True ($LASTEXITCODE -eq 0 -and $output -contains 'DEMO-PASS') ('Packaged demo failed: '+($output -join ' | '))
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
    'SmartM365-DeviceRebootManager-Demo.ps1'
    'Start-SmartM365-DeviceRebootManager-Demo.cmd'
    'SmartM365-DeviceRebootManager-GUI.strings.psd1'
    'SmartM365-DeviceRebootManager.version.json.txt'
    'SmartM365.GuiSplash.ps1'
    'Start-SmartM365-DeviceRebootManager-GUI-Test.cmd'
    'Start-SmartM365-DeviceRebootManager-GUI.cmd'
    'Tests\SmartM365-Test-DeviceRebootManagerGalleryVm.ps1'
    'Tests\SmartM365-Test-DeviceRebootManagerIntunePublisher.ps1'
    'Tests\SmartM365-Test-DeviceRebootManagerPowerShellGallery.ps1'
    'Tests\SmartM365-Test-DeviceRebootManagerGuiLogic.ps1'
    'WorkplaceCloudHub-lockup-WPF.png'
    'WorkplaceCloudHub.ico'
) | Sort-Object

$intuneRuntimePaths = @(
    'SmartM365-DeviceRebootManager.version.json.txt'
    'SmartM365-DeviceRebootManager-GUI.ps1'
    'SmartM365-DeviceRebootManager-Demo.ps1'
    'Start-SmartM365-DeviceRebootManager-Demo.cmd'
    'SmartM365-DeviceRebootManager-GUI.strings.psd1'
    'SmartM365-DeviceRebootManager-GUI.config.json.template'
    'SmartM365.GuiSplash.ps1'
    'WorkplaceCloudHub.ico'
    'WorkplaceCloudHub-lockup-WPF.png'
    'Start-SmartM365-DeviceRebootManager-GUI.cmd'
    'Start-SmartM365-DeviceRebootManager-GUI-Test.cmd'
    'Deploy\SmartM365-DeviceRebootManager-CreateScheduledTask.ps1'
    'Deploy\SmartM365-DeviceRebootManager-Detection.ps1'
    'Deploy\SmartM365-DeviceRebootManager-Install.ps1'
    'Deploy\SmartM365-DeviceRebootManager-Uninstall.ps1'
)

$versionManifest = Get-Content -LiteralPath (Join-Path $productRoot 'SmartM365-DeviceRebootManager.version.json.txt') -Raw | ConvertFrom-Json
$expectedPackageVersion = [string]$versionManifest.PackageVersion
$artifactBaseName = "SmartM365-DeviceRebootManager-$expectedPackageVersion"

try {
    Assert-True (-not [string]::IsNullOrWhiteSpace($expectedPackageVersion)) 'Version manifest has no PackageVersion.'
    $previewRoot = Join-Path -Path $testRoot -ChildPath 'Preview'
    $preview = & $builderPath `
        -OutputRoot $previewRoot `
        -SkipAuthenticodeValidation:$SkipAuthenticodeValidation

    Assert-True ($preview.Mode -eq 'Preview') 'Default mode must be Preview.'
    Assert-True (-not $preview.BuildAttempted) 'Preview unexpectedly attempted a build.'
    Assert-True (-not $preview.PublicationAttempted) 'Preview unexpectedly attempted publication.'
    Assert-True ($preview.FileCount -eq 29) 'Preview allow-list file count is incorrect.'
    Assert-True ($preview.PowerShellFileCount -eq 19) 'Preview PowerShell file count is incorrect.'
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
    Assert-True ($build.PackageVersion -eq $expectedPackageVersion) 'Unexpected package version.'
    Assert-True ($build.ExpectedReleaseTag -eq "device-reboot-manager-v$expectedPackageVersion") 'Unexpected release tag.'
    Assert-True ($build.FileCount -eq 29) 'Built package file count is incorrect.'
    Assert-True ($build.PowerShellFileCount -eq 19) 'Built package PowerShell file count is incorrect.'
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

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zipArchive = [IO.Compression.ZipFile]::OpenRead($build.ZipPath)
    try {
        $zipEntryNames = @($zipArchive.Entries | ForEach-Object { $_.FullName })
    }
    finally {
        $zipArchive.Dispose()
    }
    Assert-True (@($zipEntryNames | Where-Object { $_.Contains('\') }).Count -eq 0) 'ZIP entry names must use forward slashes.'
    Assert-True (@($zipEntryNames | Where-Object { -not $_.StartsWith("$artifactBaseName/") }).Count -eq 0) `
        'Every ZIP entry must be under the versioned root folder.'

    $expandedRoot = Join-Path -Path $testRoot -ChildPath 'Expanded'
    Expand-Archive -LiteralPath $build.ZipPath -DestinationPath $expandedRoot
    $packageRoot = Join-Path -Path $expandedRoot -ChildPath $artifactBaseName
    Assert-True (Test-Path -LiteralPath $packageRoot -PathType Container) 'Expected ZIP root folder is missing.'

    Assert-PackagedDemo -RuntimeRoot $packageRoot

    $actualPackagePaths = Get-RelativeFilePaths -Root $packageRoot
    Assert-True ($actualPackagePaths.Count -eq 29) 'ZIP does not contain exactly 29 files.'
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

    # IntuneWin validation: a package matching the current runtime is accepted;
    # stale content and corrupted ciphertext are refused before any output.
    $intuneWinRoot = Join-Path -Path $testRoot -ChildPath 'IntuneWin'
    $currentIntuneWin = Join-Path -Path (Join-Path $intuneWinRoot 'Current') -ChildPath "$artifactBaseName.intunewin"
    New-TestIntuneWin -SourceRoot $productRoot -RelativePaths $intuneRuntimePaths -DestinationPath $currentIntuneWin
    $intuneBuild = & $builderPath `
        -OutputRoot (Join-Path $testRoot 'BuildWithIntuneWin') `
        -Build `
        -IntuneWinPath $currentIntuneWin `
        -SkipAuthenticodeValidation:$SkipAuthenticodeValidation
    Assert-True $intuneBuild.IntuneWinIncluded 'Current IntuneWin was not included.'
    Assert-True $intuneBuild.IntuneWinContentValidated 'Current IntuneWin content was not validated.'
    Assert-True ((Get-Content -LiteralPath $intuneBuild.ChecksumPath -Raw) -match [regex]::Escape("$artifactBaseName.intunewin")) `
        'Checksum file does not reference the validated IntuneWin.'

    $staleIntuneWin = Join-Path -Path (Join-Path $intuneWinRoot 'Stale') -ChildPath "$artifactBaseName.intunewin"
    New-TestIntuneWin `
        -SourceRoot $productRoot `
        -RelativePaths $intuneRuntimePaths `
        -DestinationPath $staleIntuneWin `
        -ContentOverrides @{ 'Deploy\SmartM365-DeviceRebootManager-Detection.ps1' = "[string]`$ExpectedVersion = '0.0.1'" }
    $missingFileIntuneWin = Join-Path -Path (Join-Path $intuneWinRoot 'MissingFile') -ChildPath "$artifactBaseName.intunewin"
    New-TestIntuneWin `
        -SourceRoot $productRoot `
        -RelativePaths @($intuneRuntimePaths | Where-Object { $_ -ne 'WorkplaceCloudHub.ico' }) `
        -DestinationPath $missingFileIntuneWin
    $corruptIntuneWin = Join-Path -Path (Join-Path $intuneWinRoot 'Corrupt') -ChildPath "$artifactBaseName.intunewin"
    New-TestIntuneWin -SourceRoot $productRoot -RelativePaths $intuneRuntimePaths -DestinationPath $corruptIntuneWin -CorruptCiphertext

    $intuneWinRejections = 0
    foreach ($rejectedCase in @(
        @{ Path = $staleIntuneWin; Expected = '*does not match the current runtime*changed=Deploy\SmartM365-DeviceRebootManager-Detection.ps1*' }
        @{ Path = $missingFileIntuneWin; Expected = '*does not match the current runtime*missing=WorkplaceCloudHub.ico*' }
        @{ Path = $corruptIntuneWin; Expected = '*HMAC does not match*' }
    )) {
        $rejected = $false
        try {
            & $builderPath `
                -OutputRoot (Join-Path $testRoot 'RejectedIntuneWin') `
                -IntuneWinPath $rejectedCase.Path `
                -SkipAuthenticodeValidation:$SkipAuthenticodeValidation | Out-Null
        }
        catch {
            $rejected = $_.Exception.Message -like $rejectedCase.Expected
            if (-not $rejected) { throw }
        }
        Assert-True $rejected ("Release builder accepted an invalid IntuneWin: {0}" -f $rejectedCase.Path)
        $intuneWinRejections++
    }
    Assert-True (-not (Test-Path -LiteralPath (Join-Path $testRoot 'RejectedIntuneWin'))) 'Rejected IntuneWin produced output.'

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
        IntuneWinCurrentAccepted  = $intuneBuild.IntuneWinContentValidated
        IntuneWinRejectedCases    = $intuneWinRejections
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBEg0CCOf/9zUl+
# qexeo1vUiZm13dF7HWy255V0tDsRBKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIMjGRSbRNs86/bO0jgXG/Tobsv0vddQbv6GSavjbgvkEMA0GCSqG
# SIb3DQEBAQUABIIBgEfiHfUZ98yP/PtSWVG22G+PiNU9l2LFXTOEC5exMG7ouNLS
# 4skArFsSoj4l2S2vGSEjMmEMDKRmmQaBVL3FeZv7t9D55BLWcXGAUpbwNEfvKVyG
# GmvysxkyNSkvGjMkAF3liGqxzRuqzJ51Wyi/TqDmjnrmGm1F0aZ26B+lepRz6b/+
# FCSoS3Q0pYb1mD+nbsFxK9Py7t+AM3xoPbs526uKz+R1JATQYtb+PYlml8+kbJIK
# LPBFUSUGp4Tj4ZzLcKnfxLC7zi3DgjNt8OXIF4XEVtDXqRzEKwwnoLQeMYYZ0zqz
# aPGQmtguLKS8mQ3/yEv2h54o+Fo4JsNEz0rYIqxCNT+gD2qL8yH7ybKVJaTcSNRk
# yswRxF0AxVdn9NwdD6i9Jy+iep3nEoIJcX9Zdm1nBC9t+ni9MGkR2Jy7Hk11yQd+
# gvlYBVEek0JXHsoPC0Kb2dgN1kMKAowwEJpbzO0ylRgoACiS5nmTtNbN8WFgQ6Sb
# sPyO9c9IAlqR69/Xz6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcxMjQy
# MjRaMC8GCSqGSIb3DQEJBDEiBCCgo91oPEjnQRZY3oEOGuy+MXLEg5j9nts2IQcp
# +r+ZXTANBgkqhkiG9w0BAQEFAASCAgCpQ295RPoJh4ndtikCT05kErxiq3t3d85E
# Ltoqv9/GcJeN5yOINN8LbJTlOJMcnBcD01VBv3LgU2HA1ik+Qv6NJrJuJ6oKIrtl
# msE9FI+ElnTuzqvT4jHDN0QNUd9F/QAoho26+Zd9igqjLJrEyXZfeqB8fmg7+ir5
# XnlllhNRpCi9dSPEHe9PgJzmWKsdvy0knUlxSxk3TJ9gQsJdqB/QygzflxVrwpVH
# HGyc+nR5F/tkoI94vvGFFzfjdF7kjPTpVK1EButrcicrlNpvPF+9jLkhVqOGW0Fb
# nmyXG9shJXNlENsu5ZavaxOLczy/ig0H0KkU0eNzDQe/X9Y2tP9E9cEXXExDm2ds
# m3NVKSnc+JKklrp3zEOMSEDkhAYJU+rezxFZb9o+zlFvPVnKQd4sq1M/zmJQkP/k
# y9/Z65s0NzThVMfLK3ZgH+h1MV9ERCH721H7NvUA68LlDHC8CtpbQOOumkvNTFOo
# LuhsgzaVmT9+K1pa9laKsOjeuO113Qk65IohoITCqDvovHINGAjxYsb0xjEtT/kS
# 96tcwgRG3C/WqJyA0GkCnQu+hCkAUHJAkz175WnX593/KcNLIxAeC2+3Oh5WKRAn
# EBDGwZ6ksL1kphjkJKFuSa+E6kGz/NIA4r+EYm5wGnW0ey38Wo6NSbVaA1krqzFv
# diwI1/1oyA==
# SIG # End signature block
