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

$versionManifestPath = Join-Path -Path $SourceRoot -ChildPath 'SmartM365-DeviceRebootManager.version.json.txt'
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
    [pscustomobject]@{ Base = 'Product'; Source = 'SmartM365-DeviceRebootManager.version.json.txt'; Destination = 'SmartM365-DeviceRebootManager.version.json.txt' }
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCDOISoRMJUJBjM
# Z7pOo/YEINBem/tpHSL/HPfo0eylxaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBN1frsD1PvkjVYcpM70oMX
# Y3pLr+52tuyEe/kD30g+WjANBgkqhkiG9w0BAQEFAASCAYAZJFHX7fQGiKLH0Gi7
# G2PL8A5eJN9lfU3RImxwatSIb2HDFkt0le/ajDmOF/3z+/JaQVeYcCAy/Vmv4Cyo
# phNh8Km4tznLHV6YsBtDhxiq1lvRlB0MUSts7tVgRL+YW1mkcRMPE+3Pqsaf39bz
# LkwyrxazbfzcU/hPq/sjE2kFGlHsFoQKLhpPKUA3HC/V9AhNfrfrvpxRo9BJkdQP
# Bhjh2btoXC8YmJyZHFrHjZ995epnsm69Sjtni6mg3ZBWWt+vuRDUxY8MBbMsxHjd
# KjV5tL0KdcT0Y8pH3Eh3KOwD9H7/W0BOBlL5+Kj1bn6a0FBXOnNtqOS/EaW+KSE5
# uvGC3dyJ3vrsacIS8vB8Alua6BFY5Uhf+qKf371/BgSEUXiJIZXJuVCgMOef54+3
# 422JD8iQ0s3i3F+ZbOE6zBN5Mt96lSgoHWBsm5mNpG90forB6j4Aj+riW3I6gAnF
# lCoTzVB1JHXKxS9sNCSteQW8QIUNnnPtbVd5mA2pEm94eR0=
# SIG # End signature block
