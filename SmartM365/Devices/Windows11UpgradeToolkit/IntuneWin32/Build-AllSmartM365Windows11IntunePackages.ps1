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

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDFCh8ewgan8iBS
# hRgrv7OyGC+aqbA7tbiD/HvZSKOYiqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCBgkyfvGBObtrrmxvd3kw4tVv6x+E0J8ZEGI0J17u2ZZTANBgkqhkiG9w0B
# AQEFAASCAYAb6vBDHR1XkktPNckh4EC4VMNCfmVT2co/iwNh/mAoxbhrcmeI+4bg
# UiLrh795F8XwV0VennKct7eLTLHVCacmvnm9yklvBbcH17CrRU6PPx+SXDNJWUPl
# Kb+rZSQ0Mz1xwBIgSK60uiF1xcrrLcp6Z9kjghMSB5Rz1qnBxvaHaekyrLZHfe2E
# 8P8H8/FA3Ywns87D2WgV9/CtKJMA71eajvA1u/u4f+RjH/FpOoRZyfDqR+aP06T5
# CfNZWNcasFUyh/zY3TJROpQq4MXLlXE0VOYQwGrMXO9CWUSGEk7d6U79AtzLIrqj
# Oi8hf5RlWk+Gd7jCdX4DcYcBdQmpSmEZHyCzgv1/APHnkMNXqX6meA82PJVIBGET
# qKNgAW021f8xSSIp+G3sfSGriy8jgCyxgymjR9vTzEODo9p2ksLEhv1gcaX9G3XA
# WhI+oYas7/Mlnb5ohHTLFrjmMhCZL4DWosnSY9ZlYwv8UWYDYFTd7JZuRGOPeHHs
# f98YRjsUaao=
# SIG # End signature block
