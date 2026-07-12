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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDFCh8ewgan8iBS
# hRgrv7OyGC+aqbA7tbiD/HvZSKOYiqCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBgkyfvGBObtrrmxvd3
# kw4tVv6x+E0J8ZEGI0J17u2ZZTANBgkqhkiG9w0BAQEFAASCAYBEF3GDqXWm1/Mw
# nFqfjLoVIPsqzQ7xwWMum0+3kSDw73DifhX5wKOkFAhT96a9LPXJQlYPWlYOgmZG
# nASAgQMU+H6QqvAqk0KK4R6kmU9q14zP3RTH0w/uVaZeKNmAGAw7P8lL4IyM64QS
# Q1rJ1EJrfv1bxd9FVQZimFdbCTM4l0EsxqFgkx3MZroEfRNPnzUQUtWWGMo3sIdj
# l4yQK/UZEvHnqfTOKwSFS8/msF5gE0CG0bLgUIJvGCPKb0me8snB44vKiLpx3Vx9
# Dy5q0xGtwpJL72OjzQHEGBmuqzh+hqmcAA6blyi3BYvRKVHQqvR/U8suw8WDlZ3j
# fB11WWHviw1X0e1IISzTFFYnfgaXkxjnALfpi/DZeqZciwpgMPLzD3eMs8AhjxoU
# j8AtKJWVlR0AOadV3I4GBCT6yn2ti2Hi4ugMrHIzSUwzdDB5W+duj8bz8fB0ODV0
# tmd5kNdJW06Ww3fI2owJSVFSw6rHBJxOtlhXRYAEZXjfw48CsEo=
# SIG # End signature block
