#Requires -Version 5.1

<#
.SYNOPSIS
Validates and optionally builds the standalone Device Reboot Manager release artifacts.

.DESCRIPTION
Uses an explicit 26-file allow-list, validates version consistency, PowerShell
syntax, the pinned WorkplaceCloudHub Authenticode signer, and README release
references. Preview is the default and writes nothing. Use -Build to create the
ZIP and SHA-256 file locally. An optional existing IntuneWin can be copied and
included in the checksum file.

This script never publishes to GitHub, PowerShell Gallery, or Microsoft Intune.

.EXAMPLE
.\SmartM365-Build-DeviceRebootManagerRelease.ps1

.EXAMPLE
.\SmartM365-Build-DeviceRebootManagerRelease.ps1 `
    -Build `
    -IntuneWinPath C:\Temp\SmartM365-DeviceRebootManager-0.1.0.intunewin `
    -OutputRoot C:\Temp\DeviceRebootManager-Release
#>

[CmdletBinding()]
param(
    [string]$SourceRoot,
    [string]$RepositoryRoot,
    [string]$OutputRoot = (Join-Path ([IO.Path]::GetTempPath()) 'SmartM365\DeviceRebootManager\Release'),
    [string]$IntuneWinPath,
    [string]$ExpectedSignerThumbprint = 'D70ECB7B00377EBFB76B304C08DFC6620584E114',
    [switch]$Build,
    [switch]$Force,
    [switch]$SkipAuthenticodeValidation
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($SourceRoot)) {
    $SourceRoot = Split-Path -Path $PSScriptRoot -Parent
}

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
    if ($resolvedChild -eq $resolvedRoot -or
        -not $resolvedChild.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Unsafe release path outside the expected root: $resolvedChild"
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
        $messages = @($parseErrors | ForEach-Object { $_.Message })
        throw ("PowerShell syntax validation failed: file={0}; errors={1}" -f $Path,($messages -join '; '))
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

    if ($signature.Status -ne 'Valid' -or
        $actualThumbprint -ne (Get-NormalizedThumbprint $Thumbprint)) {
        throw ("Authenticode validation failed: file={0}; status={1}; signer={2}" -f
            $Path,$signature.Status,$actualThumbprint)
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

if (-not (Test-Path -LiteralPath $SourceRoot -PathType Container)) {
    throw "Product source root not found: $SourceRoot"
}
$SourceRoot = (Resolve-Path -LiteralPath $SourceRoot).Path

if ([string]::IsNullOrWhiteSpace($RepositoryRoot)) {
    $RepositoryRoot = Split-Path -Path (Split-Path -Path (Split-Path -Path $SourceRoot -Parent) -Parent) -Parent
}
if (-not (Test-Path -LiteralPath $RepositoryRoot -PathType Container)) {
    throw "Repository root not found: $RepositoryRoot"
}
$RepositoryRoot = (Resolve-Path -LiteralPath $RepositoryRoot).Path

$versionManifestPath = Join-Path -Path $SourceRoot -ChildPath 'SmartM365-DeviceRebootManager.version.json'
if (-not (Test-Path -LiteralPath $versionManifestPath -PathType Leaf)) {
    throw "Package version manifest not found: $versionManifestPath"
}
$versionManifest = Get-Content -LiteralPath $versionManifestPath -Raw | ConvertFrom-Json
if ([int]$versionManifest.SchemaVersion -ne 1 -or
    [string]::IsNullOrWhiteSpace([string]$versionManifest.PackageVersion)) {
    throw "Invalid package version manifest: $versionManifestPath"
}
$packageVersion = [string]$versionManifest.PackageVersion

$moduleManifestPath = Join-Path -Path $SourceRoot -ChildPath 'PowerShellGallery\Module\SmartM365.DeviceRebootManager.psd1'
$moduleManifest = Test-ModuleManifest -Path $moduleManifestPath
$modulePrerelease = [string]$moduleManifest.PrivateData.PSData['Prerelease']
$modulePackageVersion = if ([string]::IsNullOrWhiteSpace($modulePrerelease)) {
    [string]$moduleManifest.Version
}
else {
    '{0}-{1}' -f $moduleManifest.Version,$modulePrerelease
}
if ($modulePackageVersion -ne $packageVersion) {
    throw ("Package version mismatch: manifest={0}; module={1}" -f $packageVersion,$modulePackageVersion)
}

$detectionPath = Join-Path -Path $SourceRoot -ChildPath 'Deploy\SmartM365-DeviceRebootManager-Detection.ps1'
$detectionText = Get-Content -LiteralPath $detectionPath -Raw
$detectionVersionMatch = [regex]::Match(
    $detectionText,
    '\[string\]\$ExpectedVersion\s*=\s*''([^'']+)'''
)
if (-not $detectionVersionMatch.Success -or
    $detectionVersionMatch.Groups[1].Value -ne $packageVersion) {
    throw ("Detection default version does not match package version: package={0}; detection={1}" -f
        $packageVersion,$detectionVersionMatch.Groups[1].Value)
}

$artifactBaseName = "SmartM365-DeviceRebootManager-$packageVersion"
$expectedReleaseTag = "device-reboot-manager-v$packageVersion"
$readmePath = Join-Path -Path $SourceRoot -ChildPath 'README.md'
$readmeText = Get-Content -LiteralPath $readmePath -Raw
$requiredReadmeValues = @(
    "releases/tag/$expectedReleaseTag"
    "$artifactBaseName.zip"
    "$artifactBaseName.intunewin"
    "$artifactBaseName.sha256"
    'root of the extracted standalone release'
)
$missingReadmeValues = @($requiredReadmeValues | Where-Object { -not $readmeText.Contains($_) })
if ($missingReadmeValues.Count -gt 0) {
    throw ("README release references are incomplete for {0}: {1}" -f
        $packageVersion,($missingReadmeValues -join ', '))
}
if ($readmeText -match '\.\\Devices\\DeviceRebootManager\\') {
    throw 'README contains repository-relative commands that are invalid in the standalone ZIP.'
}

$releaseFiles = @(
    [pscustomobject]@{ Base = 'Repository'; Source = 'LICENSE'; Destination = 'LICENSE' }
    [pscustomobject]@{ Base = 'Repository'; Source = 'NOTICE'; Destination = 'NOTICE' }
    [pscustomobject]@{ Base = 'Product'; Source = 'README.md'; Destination = 'README.md' }
    [pscustomobject]@{ Base = 'Product'; Source = 'SmartM365-DeviceRebootManager.version.json'; Destination = 'SmartM365-DeviceRebootManager.version.json' }
    [pscustomobject]@{ Base = 'Product'; Source = 'SmartM365-DeviceRebootManager-GUI.ps1'; Destination = 'SmartM365-DeviceRebootManager-GUI.ps1' }
    [pscustomobject]@{ Base = 'Product'; Source = 'SmartM365-DeviceRebootManager-GUI.strings.psd1'; Destination = 'SmartM365-DeviceRebootManager-GUI.strings.psd1' }
    [pscustomobject]@{ Base = 'Product'; Source = 'SmartM365-DeviceRebootManager-GUI.config.json.template'; Destination = 'SmartM365-DeviceRebootManager-GUI.config.json.template' }
    [pscustomobject]@{ Base = 'Product'; Source = 'SmartM365.GuiSplash.ps1'; Destination = 'SmartM365.GuiSplash.ps1' }
    [pscustomobject]@{ Base = 'Product'; Source = 'Start-SmartM365-DeviceRebootManager-GUI.cmd'; Destination = 'Start-SmartM365-DeviceRebootManager-GUI.cmd' }
    [pscustomobject]@{ Base = 'Product'; Source = 'Start-SmartM365-DeviceRebootManager-GUI-Test.cmd'; Destination = 'Start-SmartM365-DeviceRebootManager-GUI-Test.cmd' }
    [pscustomobject]@{ Base = 'Product'; Source = 'WorkplaceCloudHub.ico'; Destination = 'WorkplaceCloudHub.ico' }
    [pscustomobject]@{ Base = 'Product'; Source = 'WorkplaceCloudHub-lockup-WPF.png'; Destination = 'WorkplaceCloudHub-lockup-WPF.png' }
    [pscustomobject]@{ Base = 'Product'; Source = 'Deploy\SmartM365-DeviceRebootManager-CreateScheduledTask.ps1'; Destination = 'Deploy\SmartM365-DeviceRebootManager-CreateScheduledTask.ps1' }
    [pscustomobject]@{ Base = 'Product'; Source = 'Deploy\SmartM365-DeviceRebootManager-Detection.ps1'; Destination = 'Deploy\SmartM365-DeviceRebootManager-Detection.ps1' }
    [pscustomobject]@{ Base = 'Product'; Source = 'Deploy\SmartM365-DeviceRebootManager-Install.ps1'; Destination = 'Deploy\SmartM365-DeviceRebootManager-Install.ps1' }
    [pscustomobject]@{ Base = 'Product'; Source = 'Deploy\SmartM365-DeviceRebootManager-PublishIntune.ps1'; Destination = 'Deploy\SmartM365-DeviceRebootManager-PublishIntune.ps1' }
    [pscustomobject]@{ Base = 'Product'; Source = 'Deploy\SmartM365-DeviceRebootManager-Uninstall.ps1'; Destination = 'Deploy\SmartM365-DeviceRebootManager-Uninstall.ps1' }
    [pscustomobject]@{ Base = 'Product'; Source = 'PowerShellGallery\SmartM365-Build-DeviceRebootManagerGalleryPackage.ps1'; Destination = 'PowerShellGallery\SmartM365-Build-DeviceRebootManagerGalleryPackage.ps1' }
    [pscustomobject]@{ Base = 'Product'; Source = 'PowerShellGallery\SmartM365-New-DeviceRebootManagerGalleryVmBundle.ps1'; Destination = 'PowerShellGallery\SmartM365-New-DeviceRebootManagerGalleryVmBundle.ps1' }
    [pscustomobject]@{ Base = 'Product'; Source = 'PowerShellGallery\SmartM365-Publish-DeviceRebootManagerGalleryPackage.ps1'; Destination = 'PowerShellGallery\SmartM365-Publish-DeviceRebootManagerGalleryPackage.ps1' }
    [pscustomobject]@{ Base = 'Product'; Source = 'PowerShellGallery\Module\SmartM365.DeviceRebootManager.psd1'; Destination = 'PowerShellGallery\Module\SmartM365.DeviceRebootManager.psd1' }
    [pscustomobject]@{ Base = 'Product'; Source = 'PowerShellGallery\Module\SmartM365.DeviceRebootManager.psm1'; Destination = 'PowerShellGallery\Module\SmartM365.DeviceRebootManager.psm1' }
    [pscustomobject]@{ Base = 'Product'; Source = 'PowerShellGallery\Module\Tools\SmartM365-DeviceRebootManager-GalleryUpdate.ps1'; Destination = 'PowerShellGallery\Module\Tools\SmartM365-DeviceRebootManager-GalleryUpdate.ps1' }
    [pscustomobject]@{ Base = 'Product'; Source = 'Tests\SmartM365-Test-DeviceRebootManagerGalleryVm.ps1'; Destination = 'Tests\SmartM365-Test-DeviceRebootManagerGalleryVm.ps1' }
    [pscustomobject]@{ Base = 'Product'; Source = 'Tests\SmartM365-Test-DeviceRebootManagerIntunePublisher.ps1'; Destination = 'Tests\SmartM365-Test-DeviceRebootManagerIntunePublisher.ps1' }
    [pscustomobject]@{ Base = 'Product'; Source = 'Tests\SmartM365-Test-DeviceRebootManagerPowerShellGallery.ps1'; Destination = 'Tests\SmartM365-Test-DeviceRebootManagerPowerShellGallery.ps1' }
)

if ($releaseFiles.Count -ne 26) {
    throw "Release allow-list must contain exactly 26 files; found $($releaseFiles.Count)."
}
$duplicateDestinations = @(
    $releaseFiles |
        Group-Object { ([string]$_.Destination).ToLowerInvariant() } |
        Where-Object Count -gt 1
)
if ($duplicateDestinations.Count -gt 0) {
    throw ("Duplicate release destinations: {0}" -f ($duplicateDestinations.Name -join ', '))
}

$resolvedFiles = @()
foreach ($entry in $releaseFiles) {
    $basePath = if ($entry.Base -eq 'Repository') { $RepositoryRoot } else { $SourceRoot }
    $sourcePath = Join-Path -Path $basePath -ChildPath $entry.Source
    if (-not (Test-Path -LiteralPath $sourcePath -PathType Leaf)) {
        throw "Required release file not found: $sourcePath"
    }
    $resolvedFiles += [pscustomobject]@{
        SourcePath      = (Resolve-Path -LiteralPath $sourcePath).Path
        DestinationPath = [string]$entry.Destination
    }
}

$powerShellFiles = @(
    $resolvedFiles |
        Where-Object { [IO.Path]::GetExtension($_.DestinationPath) -in @('.ps1', '.psm1', '.psd1') }
)
if ($powerShellFiles.Count -ne 17) {
    throw "Release allow-list must contain exactly 17 PowerShell files; found $($powerShellFiles.Count)."
}
foreach ($file in $powerShellFiles) {
    Assert-PowerShellSyntax -Path $file.SourcePath
    if (-not $SkipAuthenticodeValidation) {
        Assert-PinnedSignature -Path $file.SourcePath -Thumbprint $ExpectedSignerThumbprint
    }
}

$resolvedIntuneWinPath = ''
if (-not [string]::IsNullOrWhiteSpace($IntuneWinPath)) {
    if (-not (Test-Path -LiteralPath $IntuneWinPath -PathType Leaf)) {
        throw "IntuneWin file not found: $IntuneWinPath"
    }
    $resolvedIntuneWinPath = (Resolve-Path -LiteralPath $IntuneWinPath).Path
    $expectedIntuneWinName = "$artifactBaseName.intunewin"
    if ([IO.Path]::GetFileName($resolvedIntuneWinPath) -ne $expectedIntuneWinName) {
        throw ("Unexpected IntuneWin filename: expected={0}; actual={1}" -f
            $expectedIntuneWinName,[IO.Path]::GetFileName($resolvedIntuneWinPath))
    }
}

$zipFileName = "$artifactBaseName.zip"
$checksumFileName = "$artifactBaseName.sha256"
$zipOutputPath = Join-Path -Path $OutputRoot -ChildPath $zipFileName
$checksumOutputPath = Join-Path -Path $OutputRoot -ChildPath $checksumFileName
$intuneWinOutputPath = Join-Path -Path $OutputRoot -ChildPath "$artifactBaseName.intunewin"

if (-not $Build) {
    return [pscustomobject]@{
        Mode                       = 'Preview'
        BuildAttempted             = $false
        PublicationAttempted       = $false
        PackageVersion             = $packageVersion
        ExpectedReleaseTag         = $expectedReleaseTag
        PackageRootName            = $artifactBaseName
        FileCount                  = $releaseFiles.Count
        PowerShellFileCount        = $powerShellFiles.Count
        SignaturesValidated        = (-not $SkipAuthenticodeValidation)
        ExpectedSignerThumbprint   = Get-NormalizedThumbprint $ExpectedSignerThumbprint
        ReadmeValidated            = $true
        DetectionVersionValidated  = $true
        IntuneWinIncluded          = (-not [string]::IsNullOrWhiteSpace($resolvedIntuneWinPath))
        OutputRoot                 = [IO.Path]::GetFullPath($OutputRoot)
        ZipPath                    = $zipOutputPath
        ChecksumPath               = $checksumOutputPath
        IntuneWinPath              = if ($resolvedIntuneWinPath) { $intuneWinOutputPath } else { '' }
        Ready                      = $true
    }
}

$outputFullPath = [IO.Path]::GetFullPath($OutputRoot)
$plannedOutputPaths = @($zipOutputPath,$checksumOutputPath,$intuneWinOutputPath)
$existingOutputPaths = @($plannedOutputPaths | Where-Object { Test-Path -LiteralPath $_ })
if ($existingOutputPaths.Count -gt 0 -and -not $Force) {
    throw ("Release output already exists. Use -Force to replace only the version-specific artifacts: {0}" -f
        ($existingOutputPaths -join ', '))
}

$temporaryRoot = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath (
    'SmartM365-DeviceRebootManager-ReleaseBuild-{0}' -f [guid]::NewGuid().ToString('N')
)
$temporaryRoot = [IO.Path]::GetFullPath($temporaryRoot)
$temporaryBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\', '/')
Assert-SafeChildPath -Root $temporaryBase -Child $temporaryRoot

try {
    $packageRoot = Join-Path -Path $temporaryRoot -ChildPath $artifactBaseName
    $candidateRoot = Join-Path -Path $temporaryRoot -ChildPath 'Artifacts'
    $verificationRoot = Join-Path -Path $temporaryRoot -ChildPath 'Verify'
    New-Item -ItemType Directory -Path $packageRoot -Force | Out-Null
    New-Item -ItemType Directory -Path $candidateRoot -Force | Out-Null

    foreach ($file in $resolvedFiles) {
        $destinationPath = Join-Path -Path $packageRoot -ChildPath $file.DestinationPath
        New-Item -ItemType Directory -Path (Split-Path -Path $destinationPath -Parent) -Force | Out-Null
        Copy-Item -LiteralPath $file.SourcePath -Destination $destinationPath -Force
    }

    $packagedRelativePaths = Get-RelativeFilePaths -Root $packageRoot
    $expectedRelativePaths = @($releaseFiles.Destination | Sort-Object)
    if ($packagedRelativePaths.Count -ne $expectedRelativePaths.Count -or
        @(Compare-Object -ReferenceObject $expectedRelativePaths -DifferenceObject $packagedRelativePaths).Count -ne 0) {
        throw 'Staged package contents do not exactly match the release allow-list.'
    }

    $stagedReadmePath = Join-Path -Path $packageRoot -ChildPath 'README.md'
    if ((Get-FileHash -LiteralPath $stagedReadmePath -Algorithm SHA256).Hash -ne
        (Get-FileHash -LiteralPath $readmePath -Algorithm SHA256).Hash) {
        throw 'Staged README does not exactly match the repository README.'
    }

    $candidateZipPath = Join-Path -Path $candidateRoot -ChildPath $zipFileName
    Compress-Archive -LiteralPath $packageRoot -DestinationPath $candidateZipPath -CompressionLevel Optimal
    Expand-Archive -LiteralPath $candidateZipPath -DestinationPath $verificationRoot

    $verifiedPackageRoot = Join-Path -Path $verificationRoot -ChildPath $artifactBaseName
    $verifiedRelativePaths = Get-RelativeFilePaths -Root $verifiedPackageRoot
    if ($verifiedRelativePaths.Count -ne $expectedRelativePaths.Count -or
        @(Compare-Object -ReferenceObject $expectedRelativePaths -DifferenceObject $verifiedRelativePaths).Count -ne 0) {
        throw 'ZIP contents do not exactly match the release allow-list.'
    }

    $verifiedReadmePath = Join-Path -Path $verifiedPackageRoot -ChildPath 'README.md'
    if ((Get-FileHash -LiteralPath $verifiedReadmePath -Algorithm SHA256).Hash -ne
        (Get-FileHash -LiteralPath $readmePath -Algorithm SHA256).Hash) {
        throw 'ZIP README does not exactly match the repository README.'
    }

    $verifiedPowerShellFiles = @(
        Get-ChildItem -LiteralPath $verifiedPackageRoot -Recurse -File |
            Where-Object { $_.Extension -in @('.ps1', '.psm1', '.psd1') }
    )
    if ($verifiedPowerShellFiles.Count -ne 17) {
        throw "ZIP must contain exactly 17 PowerShell files; found $($verifiedPowerShellFiles.Count)."
    }
    foreach ($file in $verifiedPowerShellFiles) {
        Assert-PowerShellSyntax -Path $file.FullName
        if (-not $SkipAuthenticodeValidation) {
            Assert-PinnedSignature -Path $file.FullName -Thumbprint $ExpectedSignerThumbprint
        }
    }

    $candidateIntuneWinPath = ''
    if ($resolvedIntuneWinPath) {
        $candidateIntuneWinPath = Join-Path -Path $candidateRoot -ChildPath "$artifactBaseName.intunewin"
        Copy-Item -LiteralPath $resolvedIntuneWinPath -Destination $candidateIntuneWinPath -Force
        if ((Get-FileHash -LiteralPath $candidateIntuneWinPath -Algorithm SHA256).Hash -ne
            (Get-FileHash -LiteralPath $resolvedIntuneWinPath -Algorithm SHA256).Hash) {
            throw 'Copied IntuneWin does not exactly match the supplied source file.'
        }
    }

    $zipHash = (Get-FileHash -LiteralPath $candidateZipPath -Algorithm SHA256).Hash
    $checksumLines = @("$zipHash  $zipFileName")
    $intuneWinHash = ''
    if ($candidateIntuneWinPath) {
        $intuneWinHash = (Get-FileHash -LiteralPath $candidateIntuneWinPath -Algorithm SHA256).Hash
        $checksumLines += "$intuneWinHash  $artifactBaseName.intunewin"
    }
    $candidateChecksumPath = Join-Path -Path $candidateRoot -ChildPath $checksumFileName
    [IO.File]::WriteAllText(
        $candidateChecksumPath,
        (($checksumLines -join "`n") + "`n"),
        (New-Object Text.UTF8Encoding($false))
    )

    New-Item -ItemType Directory -Path $outputFullPath -Force | Out-Null
    foreach ($path in $plannedOutputPaths) {
        Assert-SafeChildPath -Root $outputFullPath -Child $path
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
        }
    }
    Copy-Item -LiteralPath $candidateZipPath -Destination $zipOutputPath -Force
    Copy-Item -LiteralPath $candidateChecksumPath -Destination $checksumOutputPath -Force
    if ($candidateIntuneWinPath) {
        Copy-Item -LiteralPath $candidateIntuneWinPath -Destination $intuneWinOutputPath -Force
    }

    return [pscustomobject]@{
        Mode                       = 'LocalBuild'
        BuildAttempted             = $true
        PublicationAttempted       = $false
        PackageVersion             = $packageVersion
        ExpectedReleaseTag         = $expectedReleaseTag
        PackageRootName            = $artifactBaseName
        FileCount                  = $packagedRelativePaths.Count
        PowerShellFileCount        = $verifiedPowerShellFiles.Count
        SignaturesValidated        = (-not $SkipAuthenticodeValidation)
        ExpectedSignerThumbprint   = Get-NormalizedThumbprint $ExpectedSignerThumbprint
        ReadmeValidated            = $true
        DetectionVersionValidated  = $true
        IntuneWinIncluded          = [bool]$candidateIntuneWinPath
        OutputRoot                 = $outputFullPath
        ZipPath                    = $zipOutputPath
        ZipSHA256                  = $zipHash
        ChecksumPath               = $checksumOutputPath
        ChecksumSHA256             = (Get-FileHash -LiteralPath $checksumOutputPath -Algorithm SHA256).Hash
        IntuneWinPath              = if ($candidateIntuneWinPath) { $intuneWinOutputPath } else { '' }
        IntuneWinSHA256            = $intuneWinHash
        Ready                      = $true
    }
}
finally {
    if (Test-Path -LiteralPath $temporaryRoot) {
        $resolvedTemporaryRoot = [IO.Path]::GetFullPath($temporaryRoot)
        $temporaryPrefix = $temporaryBase + [IO.Path]::DirectorySeparatorChar
        if ($resolvedTemporaryRoot.StartsWith($temporaryPrefix, [StringComparison]::OrdinalIgnoreCase) -and
            [IO.Path]::GetFileName($resolvedTemporaryRoot) -like 'SmartM365-DeviceRebootManager-ReleaseBuild-*') {
            Remove-Item -LiteralPath $resolvedTemporaryRoot -Recurse -Force
        }
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAHSsxCEuR3Dhkk
# QZAv04+yH2nvOr6G6IIALm5q+GBSyqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICCS8GWPEUv1K1HTlklau6jno6KOIjAn/LGL5ej11QhHMA0GCSqG
# SIb3DQEBAQUABIIBgAASEoCaKhT0aW4EbARfgid/x43BmLNC8A53EOMl6pz4B7QL
# vbeLcEHEIDWXRbi9QzDVClPfth/w96shvqYF2IGXH9k5hsCBbql8aQP9jeujDbx/
# f3zPNitFbZ1ekR7aPBp6gTKp9W/xqBhGYi3e/PyvGmjv7+XMo/bTys0SyTIvspZJ
# MZYT3z815VmwFyekWcXM5f734fBDAYUY8hWOV7oMjg3BdmWs/DLnfzRzP6OTq/TB
# cmYI0EpKiDE2C5kKFOnsBQ/dYAYGfU0bgPoNGQDzQEBNiTI+neD2+7BXAwZ47cg6
# JlEY9A8c0Ff9is+tLM4Dy/J1dK539pgs21smlIwWkSaSdn3uxhDfKoKJDIVxZfgU
# UweC1X7nKvJ8oHlRWEMLUxO0ZCWHN2jmX6rgAQMk0wxEVW6SrlET6vQBWUU5v+cv
# Ki5Kztc9sirFOZ/HGCypBPqN6CTPhflC8yAJHLL0KakMC1oiCTnK7wZOcrNeb5O+
# 3cptkwzF0XlA3GEDGqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MDcxMDUz
# MzNaMC8GCSqGSIb3DQEJBDEiBCBLzTjLzSK5NC1+pyBPI6qsJbReP/mvSrzMiX32
# g5W67zANBgkqhkiG9w0BAQEFAASCAgBdsFOeXsNrveDGEj8p5/0D7yxexhSUBZAk
# Squ6aB8K1LuKmO8ky39CzZJSGbPFYAvollqhWDCADFsrnBCqei+vONxTE05m20IB
# 1Ap/6x7BD3RvjJYffyKU0k8rZdkBYy6UhHjF9u04RX+I3g5mr4fwhizzYlT0Y7CZ
# fBmKSPmv9MbeXgvLjS7IBuiYmdxkG83JobMHd0hivLNnh/Vq2TQVgdbuQzy4+iFb
# mu1h5lUgg2wON5GJgjezAsgfHu3ApxlYy5wMfrZPti6zWy+zrccY9SMxmhjLv9pb
# HGORfK0555DX4tZfz4jTqSUdBCQH+nDzSNSwfrnKsHZuWwsP7+YLSrKPiS6FbSQ8
# wJnkCOkZBKFtrY4iwu549gVV41tm4rQog5dOUsGLzIIVKOo2iy6ZaoJwBM3K/A+5
# L5Alh3af+W2/m4mltM/oCQKrYLGaIRzAWiON5kwA7iLl+GQsURHTJXgtUAA9L8Ol
# kmtrvdYJvEclKGEYzZ1WxjDEyjcWE8OzYOV0uks/gKyltLb/A3PYb8YyrN70ndIl
# DsfuXuDUFxhrfJaj1X8v88WpgPqNzJJBPdCdGFpf/UI4uTkCPeym4kQTwp5XLZvd
# y4de65JACyhsmcICV6Ogd16RFyMLcfusNxE/A8pm9ap7/rgsLYvegSww8U1rjaDo
# tYRydGFhWQ==
# SIG # End signature block
