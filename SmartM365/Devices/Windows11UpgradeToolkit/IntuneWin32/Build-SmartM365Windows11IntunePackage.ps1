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

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAJM6EgrhZtU7EB
# U6s9VS5VfJIJhPgtgUEUtGC4tbufaaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDgj9ml8dupta5ITUEY885kM6gSaoUMQB73p01KgjPixTANBgkqhkiG9w0B
# AQEFAASCAYAPxEaXc0qfbnXt2tlowLjDAYKGaj5gxQgl1FsdkHznp53NbcEv21sb
# VoL6boETqDy098XpPD5MUE9Fq9SsHaVihPOkBuwOS8Ty+uzKYaL+zNB139NgdwRY
# QNGm4MF6dFCxzb3b6U9HxE/4wFXZqDEbf1RJslqrJ5DJyogyAgLsgLtDuBJzPrGR
# m9iKDGm8EwfYzVs1SeVTW16G8gqAjYk9pVEgTjKMqIrUNa2DR0om+Kak+PblkWdu
# WfmG3Lm5Ng9NgLjrgVb+YFYk053Xl0Np7nzAn+4dmDP4aCOWLH8f12FzCqz5D2kP
# IDxhoaoKcSZnn57DcozxWs7twTmbAv3lFW+P6JIEd1ZQFJZDmjB8+DS3TPFgeemf
# RAxzGzyYjL/Bqsyem+LT5JcvlRkSyXgK2Hzr937L8bLBQCUy5Wb0zgIlt44zaxry
# Q9YGdlZreBktGikWJCS0lCAs03r6He/wd8AXM8FlppP+1le0dDcJx/PCQSTqqYCc
# S/v657KPl6o=
# SIG # End signature block
