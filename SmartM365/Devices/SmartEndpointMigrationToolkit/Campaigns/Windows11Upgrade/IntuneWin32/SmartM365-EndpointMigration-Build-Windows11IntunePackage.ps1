<#
.SYNOPSIS
    Builds one Smart Endpoint Migration Toolkit Windows 11 Upgrade Intune Win32 package per setup language.
.DESCRIPTION
    Creates a staging source folder containing the endpoint script, Intune installer scripts, package manifest, and optionally one language-specific Windows setup media cache, then optionally runs IntuneWinAppUtil.exe.
    Package ids use the SmartM365-SEMT-Windows11Upgrade- prefix and app names use 'Smart Endpoint Migration Toolkit - Windows 11 Upgrade (<language>)' so they never collide with the previous Windows11UpgradeToolkit packages.
    Provenance: copied from SmartM365/Devices/Windows11UpgradeToolkit/IntuneWin32/Build-SmartM365Windows11IntunePackage.ps1 1.0.2 (see Docs\PROVENANCE.json.txt).
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][ValidatePattern('^[a-z]{2}-[A-Z]{2}$')][string]$Language,
    [string]$MediaFolder,
    [string]$MediaId = 'Win11',
    [string]$SetupSourceRoot,
    [string]$PackageVersion,
    [string]$IntuneWinAppUtilPath,
    [string]$ContentPrepRoot = 'C:\tmp\SmartM365-SEMT-W11-ContentPrep',
    [switch]$WithCacheOnly,
    [switch]$AllowSetupMediaDuplication,
    [switch]$SkipIntuneWinBuild,
    [switch]$Force
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$toolkitRoot = Split-Path -Parent $scriptDir
$sourceTemplateRoot = Join-Path $scriptDir 'Source'
$endpointScriptFileName = 'SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'
$installerFileName = 'SmartM365-EndpointMigration-Install.ps1'
$runnerFileName = 'SmartM365-EndpointMigration-Run-IntuneUpgrade.ps1'
$integrityHelperFileName = 'SmartM365-EndpointMigration-SetupMediaIntegrity.ps1'
$detectTemplateFileName = 'SmartM365-EndpointMigration-Detect-Template.ps1'
$detectOutputFileName = 'SmartM365-EndpointMigration-Detect.ps1'
$endpointScript = Join-Path $toolkitRoot ('Scripts\{0}' -f $endpointScriptFileName)
$integrityHelperSource = Join-Path $sourceTemplateRoot $integrityHelperFileName
if (-not (Test-Path -LiteralPath $integrityHelperSource -PathType Leaf)) { throw "Setup media integrity helper not found: $integrityHelperSource" }
. $integrityHelperSource

function Get-EndpointVersion {
    param([string]$Path)
    $content = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    $match = [regex]::Match($content, '(?m)^\$script:ScriptVersion\s*=\s*''(?<Version>[^'']+)''')
    if (-not $match.Success) { throw "Unable to read endpoint script version from $Path" }
    return $match.Groups['Version'].Value
}

function Get-DefaultMediaFolder {
    param([string]$Value)
    $parts = $Value -split '-'
    if ($parts.Count -ne 2) { throw "Invalid language: $Value" }
    return ('{0}-{1}' -f $parts[0].ToUpperInvariant(), $parts[1].ToLowerInvariant())
}

if ([string]::IsNullOrWhiteSpace($SetupSourceRoot)) { $SetupSourceRoot = Join-Path $toolkitRoot 'SetupSource' }
if ([string]::IsNullOrWhiteSpace($MediaFolder)) { $MediaFolder = Get-DefaultMediaFolder -Value $Language }
$endpointVersion = Get-EndpointVersion -Path $endpointScript
if ([string]::IsNullOrWhiteSpace($PackageVersion)) { $PackageVersion = $endpointVersion }

$packageSuffix = if ($WithCacheOnly) { '-WithCacheOnly' } else { '' }
$packageId = "SmartM365-SEMT-Windows11Upgrade-$MediaId-$Language$packageSuffix"
$displayName = if ($WithCacheOnly) { "Smart Endpoint Migration Toolkit - Windows 11 Upgrade ($Language, WithCacheOnly)" } else { "Smart Endpoint Migration Toolkit - Windows 11 Upgrade ($Language)" }
$cacheFolder = "$MediaId-$Language"
$packageMode = if ($WithCacheOnly) { 'WithCacheOnly' } else { 'WithMedia' }
$mediaRoot = Join-Path $SetupSourceRoot $MediaFolder

if (-not $WithCacheOnly) {
    if (-not (Test-Path -LiteralPath (Join-Path $mediaRoot 'setup.exe') -PathType Leaf)) { throw "setup.exe not found in media root: $mediaRoot" }
    if (-not (Test-Path -LiteralPath (Join-Path $mediaRoot 'sources\install.wim') -PathType Leaf)) { throw "sources\install.wim not found in media root: $mediaRoot" }
    if (-not (Test-Path -LiteralPath (Join-Path $mediaRoot 'SmartM365-SetupMediaManifest.sha256.csv') -PathType Leaf)) { throw "SmartM365-SetupMediaManifest.sha256.csv not found in media root: $mediaRoot" }
    $sourceIntegrity = Test-SmartM365SetupMediaIntegrity -MediaRoot $mediaRoot
    Write-Host ("Validated setup source integrity. Files={0}; Bytes={1}; Root={2}" -f $sourceIntegrity.Files,$sourceIntegrity.Bytes,$sourceIntegrity.MediaRoot)
}

$buildRoot = Join-Path $scriptDir ("Build\$packageId")
$packageSource = Join-Path $buildRoot 'Source'
$outputRoot = Join-Path $scriptDir ("Output\$packageId")
$mediaDest = Join-Path $packageSource ("SetupMedia\$cacheFolder")

if ((Test-Path -LiteralPath $buildRoot) -and -not $Force) { throw "Build folder already exists: $buildRoot. Use -Force to recreate it." }
if (Test-Path -LiteralPath $buildRoot) { Remove-Item -LiteralPath $buildRoot -Recurse -Force }
New-Item -ItemType Directory -Path $packageSource,$outputRoot -Force | Out-Null
if (-not $WithCacheOnly) { New-Item -ItemType Directory -Path $mediaDest -Force | Out-Null }

Copy-Item -LiteralPath (Join-Path $sourceTemplateRoot $installerFileName) -Destination (Join-Path $packageSource $installerFileName) -Force
Copy-Item -LiteralPath (Join-Path $sourceTemplateRoot $runnerFileName) -Destination (Join-Path $packageSource $runnerFileName) -Force
Copy-Item -LiteralPath $integrityHelperSource -Destination (Join-Path $packageSource $integrityHelperFileName) -Force
Copy-Item -LiteralPath $endpointScript -Destination (Join-Path $packageSource $endpointScriptFileName) -Force

$manifest = [ordered]@{
    PackageId = $packageId
    PackageVersion = $PackageVersion
    DisplayName = $displayName
    MediaId = $MediaId
    Language = $Language
    MediaFolder = $MediaFolder
    SetupCacheFolder = $cacheFolder
    PackageMode = $packageMode
    RequiresExistingSetupCache = [bool]$WithCacheOnly
    AllowSetupMediaDuplication = [bool]$AllowSetupMediaDuplication
    EndpointVersion = $endpointVersion
    CreatedUtc = (Get-Date).ToUniversalTime().ToString('o')
}
$manifestPath = Join-Path $packageSource 'PackageManifest.json'
$manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $manifestPath -Encoding ASCII
Copy-Item -LiteralPath $manifestPath -Destination (Join-Path $outputRoot 'PackageManifest.json') -Force

$robocopy = Join-Path $env:SystemRoot 'System32\robocopy.exe'
if (-not $WithCacheOnly) {
    Write-Host "Copying setup media: $mediaRoot -> $mediaDest"
    & $robocopy $mediaRoot $mediaDest /MIR /R:2 /W:5 /NP /NFL /NDL | Out-Null
    $copyExit = [int]$LASTEXITCODE
    if ($copyExit -gt 7) { throw "Robocopy media copy failed with exit code $copyExit." }
    $stagedIntegrity = Test-SmartM365SetupMediaIntegrity -MediaRoot $mediaDest
    Write-Host ("Validated staged setup media integrity. Files={0}; Bytes={1}; Root={2}" -f $stagedIntegrity.Files,$stagedIntegrity.Bytes,$stagedIntegrity.MediaRoot)
}
else {
    Write-Host "Building cache-only package. Setup media is not embedded; endpoint will require local cache: $cacheFolder"
}

$detectTemplate = Get-Content -LiteralPath (Join-Path $sourceTemplateRoot $detectTemplateFileName) -Raw
foreach ($placeholder in @('__PACKAGE_ID__','__PACKAGE_VERSION__')) {
    if (-not $detectTemplate.Contains($placeholder)) { throw "Detection template placeholder $placeholder not found: $detectTemplateFileName" }
}
$detect = $detectTemplate.Replace('__PACKAGE_ID__', $packageId).Replace('__PACKAGE_VERSION__', $PackageVersion)
if ($detect.Contains('__PACKAGE_')) { throw "Generated detection script still contains a placeholder: $detectOutputFileName" }
$detectOutputPath = Join-Path $outputRoot $detectOutputFileName
$detect | Set-Content -LiteralPath $detectOutputPath -Encoding ASCII

$commands = @"
Install command:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\$installerFileName

Uninstall command:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\$installerFileName -Uninstall

Detection script:
$detectOutputPath
"@
$commands | Set-Content -LiteralPath (Join-Path $outputRoot 'Intune-App-Commands.txt') -Encoding ASCII

$intuneWinPath = ''
if (-not $SkipIntuneWinBuild) {
    if ([string]::IsNullOrWhiteSpace($IntuneWinAppUtilPath)) { throw 'IntuneWinAppUtilPath is required unless -SkipIntuneWinBuild is used.' }
    if (-not (Test-Path -LiteralPath $IntuneWinAppUtilPath -PathType Leaf)) { throw "IntuneWinAppUtil.exe not found: $IntuneWinAppUtilPath" }

    $prepRoot = Join-Path $ContentPrepRoot $packageId
    $prepSource = Join-Path $prepRoot 'Source'
    $prepOutput = Join-Path $prepRoot 'Output'
    if (Test-Path -LiteralPath $prepRoot) { Remove-Item -LiteralPath $prepRoot -Recurse -Force }
    New-Item -ItemType Directory -Path $prepSource,$prepOutput -Force | Out-Null

    Write-Host "Preparing short IntuneWinAppUtil source path: $prepSource"
    & $robocopy $packageSource $prepSource /MIR /R:2 /W:5 /NP /NFL /NDL | Out-Null
    $prepCopyExit = [int]$LASTEXITCODE
    if ($prepCopyExit -gt 7) { throw "Robocopy content-prep copy failed with exit code $prepCopyExit." }
    if (-not $WithCacheOnly) {
        $prepMediaRoot = Join-Path $prepSource ("SetupMedia\{0}" -f $cacheFolder)
        $prepIntegrity = Test-SmartM365SetupMediaIntegrity -MediaRoot $prepMediaRoot
        Write-Host ("Validated content-prep setup media integrity. Files={0}; Bytes={1}; Root={2}" -f $prepIntegrity.Files,$prepIntegrity.Bytes,$prepIntegrity.MediaRoot)
    }

    & $IntuneWinAppUtilPath -c $prepSource -s $installerFileName -o $prepOutput -q
    if ($LASTEXITCODE -ne 0) { throw "IntuneWinAppUtil failed with exit code $LASTEXITCODE." }

    $generated = Join-Path $prepOutput ('{0}.intunewin' -f [System.IO.Path]::GetFileNameWithoutExtension($installerFileName))
    if (-not (Test-Path -LiteralPath $generated -PathType Leaf)) { throw "Expected IntuneWinAppUtil output not found: $generated" }
    $intuneWinPath = Join-Path $outputRoot ("{0}.intunewin" -f $packageId)
    Copy-Item -LiteralPath $generated -Destination $intuneWinPath -Force
}

[pscustomobject]@{
    PackageId = $packageId
    DisplayName = $displayName
    PackageMode = $packageMode
    Source = $packageSource
    Output = $outputRoot
    IntuneWin = $intuneWinPath
    DetectionScript = $detectOutputPath
}
