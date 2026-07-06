<#
.SYNOPSIS
    Builds SmartM365 Windows 11 Upgrade Toolkit Intune Win32 packages for every setup media folder.

.DESCRIPTION
    Detects language media folders under SetupSource, optionally refreshes missing media manifests,
    and calls Build-SmartM365Windows11IntunePackage.ps1 once per language.

.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>
[CmdletBinding()]
param(
    [string]$SetupSourceRoot,
    [string]$MediaId = 'Win11',
    [string]$PackageVersion,
    [string]$IntuneWinAppUtilPath,
    [string]$ContentPrepRoot = 'C:\tmp\SmartM365-W11UT-ContentPrep',
    [switch]$WithCacheOnly,
    [switch]$SkipIntuneWinBuild,
    [switch]$GenerateMissingManifest,
    [switch]$Force,
    [switch]$KeepStaging,
    [string[]]$IncludeMediaFolder,
    [string[]]$ExcludeMediaFolder
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$toolkitRoot = Split-Path -Parent $scriptDir
$builderPath = Join-Path $scriptDir 'Build-SmartM365Windows11IntunePackage.ps1'
$manifestScriptPath = Join-Path $toolkitRoot 'Scripts\New-SmartM365SetupMediaManifest.ps1'

if ([string]::IsNullOrWhiteSpace($SetupSourceRoot)) { $SetupSourceRoot = Join-Path $toolkitRoot 'SetupSource' }
if (-not (Test-Path -LiteralPath $SetupSourceRoot -PathType Container)) { throw "SetupSourceRoot not found: $SetupSourceRoot" }
if (-not (Test-Path -LiteralPath $builderPath -PathType Leaf)) { throw "Package builder not found: $builderPath" }

function Convert-MediaFolderToLanguage {
    param([Parameter(Mandatory = $true)][string]$Name)

    $match = [regex]::Match($Name, '^(?<Lang>[A-Z]{2})-(?<Region>[a-z]{2})$')
    if (-not $match.Success) { return '' }
    return ('{0}-{1}' -f $match.Groups['Lang'].Value.ToLowerInvariant(), $match.Groups['Region'].Value.ToUpperInvariant())
}

function Test-MediaFolderReady {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath (Join-Path $Path 'setup.exe') -PathType Leaf)) { return $false }
    if (-not (Test-Path -LiteralPath (Join-Path $Path 'sources\install.wim') -PathType Leaf) -and -not (Test-Path -LiteralPath (Join-Path $Path 'sources\install.esd') -PathType Leaf)) { return $false }
    return $true
}

if ($GenerateMissingManifest -and -not (Test-Path -LiteralPath $manifestScriptPath -PathType Leaf)) {
    throw "Manifest generator not found: $manifestScriptPath"
}

$includeSet = @{}
foreach ($name in @($IncludeMediaFolder)) {
    if (-not [string]::IsNullOrWhiteSpace($name)) { $includeSet[$name.ToUpperInvariant()] = $true }
}
$excludeSet = @{}
foreach ($name in @($ExcludeMediaFolder)) {
    if (-not [string]::IsNullOrWhiteSpace($name)) { $excludeSet[$name.ToUpperInvariant()] = $true }
}

$mediaFolders = @(
    Get-ChildItem -LiteralPath $SetupSourceRoot -Directory -ErrorAction Stop |
        Sort-Object Name |
        Where-Object {
            $upperName = $_.Name.ToUpperInvariant()
            (($includeSet.Count -eq 0) -or $includeSet.ContainsKey($upperName)) -and -not $excludeSet.ContainsKey($upperName)
        }
)

if ($mediaFolders.Count -eq 0) { throw "No setup media folders found under $SetupSourceRoot." }

$results = New-Object System.Collections.Generic.List[object]
foreach ($folder in $mediaFolders) {
    $language = Convert-MediaFolderToLanguage -Name $folder.Name
    if ([string]::IsNullOrWhiteSpace($language)) {
        Write-Host ("Skipping {0}: folder name does not match language media pattern like FR-fr." -f $folder.Name) -ForegroundColor Yellow
        continue
    }

    if (-not $WithCacheOnly -and -not (Test-MediaFolderReady -Path $folder.FullName)) {
        Write-Host ("Skipping {0}: setup.exe or sources\install.wim/install.esd is missing." -f $folder.Name) -ForegroundColor Yellow
        continue
    }

    $manifestPath = Join-Path $folder.FullName 'SmartM365-SetupMediaManifest.sha256.csv'
    if (-not $WithCacheOnly -and -not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
        if ($GenerateMissingManifest) {
            Write-Host ("Generating setup media manifest for {0}" -f $folder.Name) -ForegroundColor Cyan
            & $manifestScriptPath -MediaRoot $folder.FullName -Force
            if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "Manifest generation failed for $($folder.Name): $manifestPath" }
        }
        else {
            throw "Manifest missing for $($folder.Name): $manifestPath. Run with -GenerateMissingManifest or generate manifests first."
        }
    }

    $builderParameters = @{
        Language = $language
        MediaFolder = $folder.Name
        MediaId = $MediaId
        SetupSourceRoot = $SetupSourceRoot
        ContentPrepRoot = $ContentPrepRoot
    }
    if (-not [string]::IsNullOrWhiteSpace($PackageVersion)) { $builderParameters['PackageVersion'] = $PackageVersion }
    if (-not [string]::IsNullOrWhiteSpace($IntuneWinAppUtilPath)) { $builderParameters['IntuneWinAppUtilPath'] = $IntuneWinAppUtilPath }
    if ($WithCacheOnly) { $builderParameters['WithCacheOnly'] = $true }
    if ($SkipIntuneWinBuild) { $builderParameters['SkipIntuneWinBuild'] = $true }
    if ($Force) { $builderParameters['Force'] = $true }

    Write-Host ("Building Intune package for {0} ({1})" -f $folder.Name,$language) -ForegroundColor Cyan
    $buildResult = & $builderPath @builderParameters
    foreach ($item in @($buildResult)) { [void]$results.Add($item) }

    if (-not $KeepStaging) {
        $packageId = "SmartM365-Windows11UpgradeToolkit-$MediaId-$language$(if ($WithCacheOnly) { '-WithCacheOnly' } else { '' })"
        $buildRoot = Join-Path $scriptDir ("Build\$packageId")
        $prepRoot = Join-Path $ContentPrepRoot $packageId
        foreach ($stagingPath in @($buildRoot, $prepRoot)) {
            if (Test-Path -LiteralPath $stagingPath) {
                Write-Host ("Removing staging folder: {0}" -f $stagingPath) -ForegroundColor DarkGray
                Remove-Item -LiteralPath $stagingPath -Recurse -Force -ErrorAction Stop
            }
        }
    }
}

if ($results.Count -eq 0) { throw 'No package was built.' }

$results.ToArray()
