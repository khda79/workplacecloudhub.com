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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDiQ6qbBOz7NieW
# PihPr3RQcAM/vaF0qGk11x0EFh91kKCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCy4ubbJXPjU2K7fAt7
# sMZ7q7ZYrRw/UQohkmBkAgI1uzANBgkqhkiG9w0BAQEFAASCAYBkELj/g6ERNfqF
# NKs0isC+xFJdhYTE8rGpQk2RszyP3vJbk9dBPuNq5bKOQs15ejNcyMpSB0oANbPC
# Q+bk5uS+B8+Pcd+b3sqobdOZuPumrX3A/YFFTW3dibMj/X5cpzrDRMZ0UngLqOLK
# T0Iekv+JqBBVftPQoXtR1+m+/XMP0t8A2q1Rg4SBC/QqEvuqY5n+EcH4/39M1xxx
# zam6MhWXPPqzuPkHGlx7trCtTsE/uNkVPVZuW3LkOZwHaAhnjBLQNngVQNAx6hfF
# wfq/hH3rhpF/4aM4iRI7AlLPqYcjdV7/+Q6aKLZ/G6cpAtv8V6SW8jb301kkWoCG
# eHXV54YeazM/b/qmEUrcQFk4EdJEzWRnNeYPtNXhq6Vbq5GvYtOJEe8iXeIun44S
# LnOF1e/R7OAbrSmbt3KUWbBPg42AH42kfWskP4EGB3v/D/ZKdEOkxn31Rsk4QLnp
# c2VeTptvzLaOmkLyU0mOVZOAOEOVmgn2NUnJjFcHe00jf1Wd56k=
# SIG # End signature block
