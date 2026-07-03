<#
.SYNOPSIS
    Builds one SmartM365 Windows 11 Upgrade Toolkit Intune Win32 package per setup language.
.DESCRIPTION
    Creates a staging source folder containing the endpoint script, Intune installer scripts, package manifest, and optionally one language-specific Windows setup media cache, then optionally runs IntuneWinAppUtil.exe.
.VERSION
    1.0.1
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
    [string]$ContentPrepRoot = 'C:\tmp\SmartM365-W11UT-ContentPrep',
    [switch]$WithCacheOnly,
    [switch]$SkipIntuneWinBuild,
    [switch]$Force
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$toolkitRoot = Split-Path -Parent $scriptDir
$sourceTemplateRoot = Join-Path $scriptDir 'Source'
$endpointScript = Join-Path $toolkitRoot 'Scripts\SmartM365-Invoke-Windows11UpgradeRepair.ps1'

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
$packageId = "SmartM365-Windows11UpgradeToolkit-$MediaId-$Language$packageSuffix"
$displayName = "Windows11UpgradeToolkit-$Language$packageSuffix"
$cacheFolder = "$MediaId-$Language"
$packageMode = if ($WithCacheOnly) { 'WithCacheOnly' } else { 'WithMedia' }
$mediaRoot = Join-Path $SetupSourceRoot $MediaFolder

if (-not $WithCacheOnly) {
    if (-not (Test-Path -LiteralPath (Join-Path $mediaRoot 'setup.exe') -PathType Leaf)) { throw "setup.exe not found in media root: $mediaRoot" }
    if (-not (Test-Path -LiteralPath (Join-Path $mediaRoot 'sources\install.wim') -PathType Leaf)) { throw "sources\install.wim not found in media root: $mediaRoot" }
    if (-not (Test-Path -LiteralPath (Join-Path $mediaRoot 'SmartM365-SetupMediaManifest.sha256.csv') -PathType Leaf)) { throw "SmartM365-SetupMediaManifest.sha256.csv not found in media root: $mediaRoot" }
}

$buildRoot = Join-Path $scriptDir ("Build\$packageId")
$packageSource = Join-Path $buildRoot 'Source'
$outputRoot = Join-Path $scriptDir ("Output\$packageId")
$mediaDest = Join-Path $packageSource ("SetupMedia\$cacheFolder")

if ((Test-Path -LiteralPath $buildRoot) -and -not $Force) { throw "Build folder already exists: $buildRoot. Use -Force to recreate it." }
if (Test-Path -LiteralPath $buildRoot) { Remove-Item -LiteralPath $buildRoot -Recurse -Force }
New-Item -ItemType Directory -Path $packageSource,$outputRoot -Force | Out-Null
if (-not $WithCacheOnly) { New-Item -ItemType Directory -Path $mediaDest -Force | Out-Null }

Copy-Item -LiteralPath (Join-Path $sourceTemplateRoot 'Install.ps1') -Destination (Join-Path $packageSource 'Install.ps1') -Force
Copy-Item -LiteralPath (Join-Path $sourceTemplateRoot 'Run-IntuneUpgrade.ps1') -Destination (Join-Path $packageSource 'Run-IntuneUpgrade.ps1') -Force
Copy-Item -LiteralPath $endpointScript -Destination (Join-Path $packageSource 'SmartM365-Invoke-Windows11UpgradeRepair.ps1') -Force

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
}
else {
    Write-Host "Building cache-only package. Setup media is not embedded; endpoint will require local cache: $cacheFolder"
}

$detectTemplate = Get-Content -LiteralPath (Join-Path $sourceTemplateRoot 'Detect-Template.ps1') -Raw
$detect = $detectTemplate.Replace('__PACKAGE_ID__', $packageId).Replace('__PACKAGE_VERSION__', $PackageVersion)
$detect | Set-Content -LiteralPath (Join-Path $outputRoot 'Detect.ps1') -Encoding ASCII

$commands = @"
Install command:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1

Uninstall command:
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -Uninstall

Detection script:
$outputRoot\Detect.ps1
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

    & $IntuneWinAppUtilPath -c $prepSource -s 'Install.ps1' -o $prepOutput -q
    if ($LASTEXITCODE -ne 0) { throw "IntuneWinAppUtil failed with exit code $LASTEXITCODE." }

    $generated = Join-Path $prepOutput 'Install.intunewin'
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
    DetectionScript = Join-Path $outputRoot 'Detect.ps1'
}
