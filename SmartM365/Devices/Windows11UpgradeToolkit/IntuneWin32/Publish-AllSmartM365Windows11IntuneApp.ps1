<#
.SYNOPSIS
    Publishes multiple SmartM365 Windows 11 Upgrade Toolkit Intune Win32 packages.

.DESCRIPTION
    Discovers generated .intunewin packages under the IntuneWin32 Output folder and calls
    Publish-SmartM365Windows11IntuneApp.ps1 once per package. By default, only full media
    packages are selected so existing WithCacheOnly packages are not republished accidentally.

.VERSION
    1.0.3
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$OutputRoot,
    [string]$PublisherScriptPath,
    [ValidateSet('WithMedia','WithCacheOnly','All')][string]$PackageMode = 'WithMedia',
    [ValidatePattern('^[a-z]{2}-[A-Z]{2}$')][string[]]$IncludeLanguage,
    [ValidatePattern('^[a-z]{2}-[A-Z]{2}$')][string[]]$ExcludeLanguage,
    [string[]]$IncludePackageId,
    [string[]]$ExcludePackageId,
    [string]$PackageVersion,
    [switch]$ForceCreateNew,
    [switch]$UpdateMetadataOnly,
    [switch]$UpdateDetectionRules,
    [switch]$DisableLanguageRequirementRule,
    [ValidateRange(-1, 2147483647)][int]$MinimumFreeDiskSpaceInMB = -1,
    [int]$UploadBlockSizeMB = 16,
    [int]$AzureUploadMaxRetries = 5,
    [int]$PollSeconds = 10,
    [int]$PollTimeoutMinutes = 45,
    [string]$GraphBaseUri = 'https://graph.microsoft.com/beta',
    [switch]$NoConnect
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if ([string]::IsNullOrWhiteSpace($OutputRoot)) { $OutputRoot = Join-Path $scriptDir 'Output' }
if ([string]::IsNullOrWhiteSpace($PublisherScriptPath)) { $PublisherScriptPath = Join-Path $scriptDir 'Publish-SmartM365Windows11IntuneApp.ps1' }

$OutputRoot = [System.IO.Path]::GetFullPath($OutputRoot)
$PublisherScriptPath = [System.IO.Path]::GetFullPath($PublisherScriptPath)

if (-not (Test-Path -LiteralPath $OutputRoot -PathType Container)) { throw "OutputRoot not found: $OutputRoot" }
if (-not (Test-Path -LiteralPath $PublisherScriptPath -PathType Leaf)) { throw "Publisher script not found: $PublisherScriptPath" }

function Write-Step {
    param([string]$Message)
    Write-Host ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message)
}

function ConvertTo-Set {
    param([AllowNull()][string[]]$Values)

    $set = @{}
    foreach ($value in @($Values)) {
        if (-not [string]::IsNullOrWhiteSpace($value)) { $set[$value.ToLowerInvariant()] = $true }
    }
    return $set
}

function Read-PackageManifest {
    param([Parameter(Mandatory = $true)][string]$IntuneWinPath)

    $packageDir = Split-Path -Parent $IntuneWinPath
    foreach ($candidate in @((Join-Path $packageDir 'PackageManifest.json'), (Join-Path (Split-Path -Parent $packageDir) 'Source\PackageManifest.json'))) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            try { return (Get-Content -LiteralPath $candidate -Raw | ConvertFrom-Json) }
            catch { throw "Unable to read package manifest: $candidate. $($_.Exception.Message)" }
        }
    }

    return $null
}

function Get-PackageLanguageFromName {
    param([string]$Name)

    $match = [regex]::Match($Name, 'Win11-(?<Language>[a-z]{2}-[A-Z]{2})(?:-WithCacheOnly)?\.intunewin$', 'IgnoreCase')
    if ($match.Success) { return $match.Groups['Language'].Value }
    return ''
}

function Get-PackageIdFromName {
    param([string]$Name)

    if ($Name.EndsWith('.intunewin', [System.StringComparison]::OrdinalIgnoreCase)) {
        return $Name.Substring(0, $Name.Length - '.intunewin'.Length)
    }
    return $Name
}

$includeLanguageSet = ConvertTo-Set -Values $IncludeLanguage
$excludeLanguageSet = ConvertTo-Set -Values $ExcludeLanguage
$includePackageSet = ConvertTo-Set -Values $IncludePackageId
$excludePackageSet = ConvertTo-Set -Values $ExcludePackageId

$packages = @(
    Get-ChildItem -LiteralPath $OutputRoot -Filter '*.intunewin' -Recurse -File -ErrorAction Stop |
        Sort-Object FullName |
        ForEach-Object {
            $manifest = Read-PackageManifest -IntuneWinPath $_.FullName
            $packageId = if ($manifest -and $manifest.PackageId) { [string]$manifest.PackageId } else { Get-PackageIdFromName -Name $_.Name }
            $language = if ($manifest -and $manifest.Language) { [string]$manifest.Language } else { Get-PackageLanguageFromName -Name $_.Name }
            $mode = if ($manifest -and $manifest.PackageMode) { [string]$manifest.PackageMode } elseif ($packageId -match '-WithCacheOnly$') { 'WithCacheOnly' } else { 'WithMedia' }
            $version = if ($manifest -and $manifest.PackageVersion) { [string]$manifest.PackageVersion } else { '' }
            $displayName = if ($manifest -and $manifest.DisplayName) { [string]$manifest.DisplayName } else { '' }

            [pscustomobject]@{
                IntuneWinPath = $_.FullName
                PackageId = $packageId
                PackageVersion = $version
                Language = $language
                PackageMode = $mode
                DisplayName = $displayName
            }
        }
)

$selected = @(
    $packages | Where-Object {
        $languageKey = ([string]$_.Language).ToLowerInvariant()
        $packageKey = ([string]$_.PackageId).ToLowerInvariant()
        (($PackageMode -eq 'All') -or ([string]$_.PackageMode -eq $PackageMode)) -and
        (($includeLanguageSet.Count -eq 0) -or $includeLanguageSet.ContainsKey($languageKey)) -and
        (-not $excludeLanguageSet.ContainsKey($languageKey)) -and
        (($includePackageSet.Count -eq 0) -or $includePackageSet.ContainsKey($packageKey)) -and
        (-not $excludePackageSet.ContainsKey($packageKey))
    }
)

if ($selected.Count -eq 0) {
    throw "No .intunewin package matched the requested filters. OutputRoot=$OutputRoot; PackageMode=$PackageMode"
}

Write-Step ("Publishing/updating {0} package(s). OutputRoot={1}; PackageMode={2}; UpdateMetadataOnly={3}" -f $selected.Count,$OutputRoot,$PackageMode,[bool]$UpdateMetadataOnly)
foreach ($package in $selected) {
    Write-Step ("Package: {0}; Language={1}; Mode={2}; Version={3}" -f $package.PackageId,$package.Language,$package.PackageMode,$package.PackageVersion)

    $publishParams = @{
        IntuneWinPath = $package.IntuneWinPath
        UploadBlockSizeMB = $UploadBlockSizeMB
        AzureUploadMaxRetries = $AzureUploadMaxRetries
        PollSeconds = $PollSeconds
        PollTimeoutMinutes = $PollTimeoutMinutes
        MinimumFreeDiskSpaceInMB = $MinimumFreeDiskSpaceInMB
        GraphBaseUri = $GraphBaseUri
    }
    if ($ForceCreateNew) { $publishParams['ForceCreateNew'] = $true }
    if ($UpdateMetadataOnly) { $publishParams['UpdateMetadataOnly'] = $true }
    if ($UpdateDetectionRules) { $publishParams['UpdateDetectionRules'] = $true }
    if (-not [string]::IsNullOrWhiteSpace($PackageVersion)) { $publishParams['PackageVersion'] = $PackageVersion }
    if ($DisableLanguageRequirementRule) { $publishParams['DisableLanguageRequirementRule'] = $true }
    if ($NoConnect) { $publishParams['NoConnect'] = $true }

    if ($WhatIfPreference) {
        & $PublisherScriptPath @publishParams -WhatIf
        continue
    }

    $actionText = if ($UpdateMetadataOnly) { 'Update SmartM365 Windows 11 Intune app metadata' } else { 'Publish SmartM365 Windows 11 Intune app' }
    if ($PSCmdlet.ShouldProcess($package.PackageId, $actionText)) {
        & $PublisherScriptPath @publishParams
    }
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDiQ6qbBOz7NieW
# PihPr3RQcAM/vaF0qGk11x0EFh91kKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCCy4ubbJXPjU2K7fAt7sMZ7q7ZYrRw/UQohkmBkAgI1uzANBgkqhkiG9w0B
# AQEFAASCAYCEsUu6DS//gD/ntgep8wYaGI+Ml/yZw+On9HKcwJdt16Gte+h2xLi2
# gwP7SaIp6hiy6xRcXUxDeam70NaVAQcXiwMuCHUC5M7f4PQ6nOYhsVQCmzARGpuc
# GUiUB/3bfgVjUVZYCLaZhGHWgNaa1U/HcVkdVrJK7i8kwWBoiyruNZKkaFo3QLQh
# dL9VIyldVZcxwJVI/Klo7mmB5WDpgHL0I1orvWhfUY6mWp7atPnuR5rGf5N8RFCE
# bgTRtWy/SsU/90Vxs5AstbEJHLuNUxBtCKzfC7CqTL7UUlZjiXHVImAWnvnmTynj
# j6YdC72IF99M6tjGSUtGrZtTvoCYyPRjuhPNBtzL8jU5HaPPo3XvUDEqnOg5w+pm
# eutmqcjVEn1jhp+JNV+4Ii2kenzRsGfOZfRwBMoyuU1Vywjk5oGWn54TGGUNmy/o
# Dz2WSzFVGxzQSXy6kO1JCfvPoEIv+Dx27u4feb9bd8+cb+USauED1VAREy/Q86Bd
# Qkx5bfPhOYM=
# SIG # End signature block
