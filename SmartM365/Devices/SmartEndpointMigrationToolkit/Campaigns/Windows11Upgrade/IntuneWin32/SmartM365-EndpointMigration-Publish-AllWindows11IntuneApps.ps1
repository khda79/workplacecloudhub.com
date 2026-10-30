<#
.SYNOPSIS
    Publishes multiple Smart Endpoint Migration Toolkit Windows 11 Upgrade Intune Win32 packages.

.DESCRIPTION
    Discovers generated .intunewin packages under the IntuneWin32 Output folder and calls
    SmartM365-EndpointMigration-Publish-Windows11IntuneApp.ps1 once per package. By default, only full media
    packages are selected so existing WithCacheOnly packages are not republished accidentally.
    Only packages whose PackageId starts with SmartM365-SEMT-Windows11Upgrade- are published. No assignment is created.
    Provenance: copied from SmartM365/Devices/Windows11UpgradeToolkit/IntuneWin32/Publish-AllSmartM365Windows11IntuneApp.ps1 1.0.3 (see Docs\PROVENANCE.json.txt).

.VERSION
    1.0.0
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
if ([string]::IsNullOrWhiteSpace($PublisherScriptPath)) { $PublisherScriptPath = Join-Path $scriptDir 'SmartM365-EndpointMigration-Publish-Windows11IntuneApp.ps1' }

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
    foreach ($candidate in @((Join-Path $packageDir 'PackageManifest.json.txt'), (Join-Path (Split-Path -Parent $packageDir) 'Source\PackageManifest.json.txt'))) {
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

$packageIdPrefix = 'SmartM365-SEMT-Windows11Upgrade-'
foreach ($foreignPackage in @($packages | Where-Object { -not ([string]$_.PackageId).StartsWith($packageIdPrefix, [System.StringComparison]::OrdinalIgnoreCase) })) {
    Write-Step ("Skipping package without the {0} prefix: {1}" -f $packageIdPrefix,$foreignPackage.IntuneWinPath)
}
$packages = @($packages | Where-Object { ([string]$_.PackageId).StartsWith($packageIdPrefix, [System.StringComparison]::OrdinalIgnoreCase) })

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

    $actionText = if ($UpdateMetadataOnly) { 'Update Smart Endpoint Migration Toolkit Windows 11 Upgrade Intune app metadata' } else { 'Publish Smart Endpoint Migration Toolkit Windows 11 Upgrade Intune app' }
    if ($PSCmdlet.ShouldProcess($package.PackageId, $actionText)) {
        & $PublisherScriptPath @publishParams
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAzU5gSWk/kgHZ7
# 0gOR3QjN4xcVJxRYSNKV5YcCzUSg3KCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCx3OhefE/utrN6wjHqEk79
# RZs9k7EOzqowCksB5lAarTANBgkqhkiG9w0BAQEFAASCAYAbnv8FUI/8+jA+p6w5
# gI8A2n3FftIT3joeXaJh4h/0njoS0sqeR0sb3MuBaplMWdClzVBq/96MUjJmkScp
# hE/3GfZNYrlh7vjyIHnquigQ0TkvmvMEdSHJ30v6+HkvNQ4ZHnSGrh391gMCcFus
# PwwMHZlVuRgIvFW5/PeqV1n1B5Ru+tMz/lt0NckGjF6DqKVYVqiagFywISu1ODBI
# m7AsbyYAjm1p7eWIHO0dV9FjryenMqtvqxYpzMwy+Zsadj6cm1CVQGAn4HFcEGww
# uIMiX2nWO1bPo4VsIe3Sv1xOjiesoZOwi1dDy1szxVhQ+qQm96zIEDnfMZ8qm2w8
# //xWq/4l78J0ot1H8yS5Tv3T7LcubNrSiJT+ZEBO7bNjHdkZHrlN0iZySGgJs4EA
# tXs3zbvpk0MdKYHIKKbxfVx22GdCI4smq1B2C8rRGqk2CZqnZ9wtRJFHpmRfUDg7
# MzYNxD4qu9A3Jh+HZKfUc+49xoKSYHeLfuug1Y0XeffUwx0=
# SIG # End signature block
