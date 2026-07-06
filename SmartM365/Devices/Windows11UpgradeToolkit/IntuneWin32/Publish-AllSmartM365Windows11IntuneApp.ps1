<#
.SYNOPSIS
    Publishes multiple SmartM365 Windows 11 Upgrade Toolkit Intune Win32 packages.

.DESCRIPTION
    Discovers generated .intunewin packages under the IntuneWin32 Output folder and calls
    Publish-SmartM365Windows11IntuneApp.ps1 once per package. By default, only full media
    packages are selected so existing WithCacheOnly packages are not republished accidentally.

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
    [switch]$ForceCreateNew,
    [switch]$DisableLanguageRequirementRule,
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

Write-Step ("Publishing {0} package(s). OutputRoot={1}; PackageMode={2}" -f $selected.Count,$OutputRoot,$PackageMode)
foreach ($package in $selected) {
    Write-Step ("Package: {0}; Language={1}; Mode={2}; Version={3}" -f $package.PackageId,$package.Language,$package.PackageMode,$package.PackageVersion)

    $publishParams = @{
        IntuneWinPath = $package.IntuneWinPath
        UploadBlockSizeMB = $UploadBlockSizeMB
        AzureUploadMaxRetries = $AzureUploadMaxRetries
        PollSeconds = $PollSeconds
        PollTimeoutMinutes = $PollTimeoutMinutes
        GraphBaseUri = $GraphBaseUri
    }
    if ($ForceCreateNew) { $publishParams['ForceCreateNew'] = $true }
    if ($DisableLanguageRequirementRule) { $publishParams['DisableLanguageRequirementRule'] = $true }
    if ($NoConnect) { $publishParams['NoConnect'] = $true }

    if ($WhatIfPreference) {
        & $PublisherScriptPath @publishParams -WhatIf
        continue
    }

    if ($PSCmdlet.ShouldProcess($package.PackageId, 'Publish SmartM365 Windows 11 Intune app')) {
        & $PublisherScriptPath @publishParams
    }
}
