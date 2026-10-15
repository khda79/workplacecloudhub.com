[CmdletBinding()]
# Version 1.2.0: preserve application trends across compact summary history.
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DataRoot,

    [string]$InventoryOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\ApplicationInventoryEvidence.csv'),

    [string]$TrendOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\ApplicationTrendEvidence.csv')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path (Split-Path $PSScriptRoot -Parent) '../SmartM365/SmartInventory/Common/SmartM365.WorkplaceSource.psd1') -MinimumVersion '1.0.0' -ErrorAction Stop

function Convert-ToInvariantDecimalText {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    return ([double]$Value).ToString('0.0########', [Globalization.CultureInfo]::InvariantCulture)
}

function Get-ApplicationProfile {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory = $true)][datetime]$SnapshotDate,
        [Parameter(Mandatory = $true)][string]$SnapshotWeek,
        [switch]$IncludeDetail,
        [hashtable]$ProductFootprint = @{},
        [switch]$LegacyDefinition
    )

    $detail = [System.Collections.Generic.List[object]]::new()
    $fragmentedApplications = 0
    $highlyFragmentedApplications = 0
    $installationsOutsideDominant = [int64]0
    $longTailVersions = 0
    $unknownPublisherVersions = 0

    # Historical observations keep their original definition; do not rebase them.
    $groups = @(if ($LegacyDefinition) { @($Rows | Group-Object AppName) } else {
        @($Rows | Group-Object -Property { Get-WorkplaceApplicationProductKey $_ })
    })
    foreach ($group in $groups) {
        $nameText = if ($LegacyDefinition) { $group.Name } else { [string]$group.Group[0].AppName }
        $applicationName = if ([string]::IsNullOrWhiteSpace($nameText)) { 'Unknown application' } else { $nameText.Trim() }
        $productKey = Get-WorkplaceApplicationProductKey $group.Group[0]
        $distinctDevices = if ($ProductFootprint.ContainsKey($productKey)) { $ProductFootprint[$productKey].Count } else { $null }
        $versions = @($group.Group | ForEach-Object { [string]$_.AppVersion } | Sort-Object -Unique)
        $versionCount = $versions.Count
        $productTotal = [int64](($group.Group | Measure-Object DeviceCount -Sum).Sum)
        $dominant = $group.Group | Sort-Object -Property @{ Expression = { [int64]$_.DeviceCount }; Descending = $true }, @{ Expression = { [string]$_.AppVersion }; Descending = $true }, @{ Expression = { [string]$_.AppId }; Descending = $false } | Select-Object -First 1
        $dominantCount = [int64]$dominant.DeviceCount
        $dominantShare = if ($productTotal -gt 0) { $dominantCount / $productTotal } else { $null }
        $standardizationState = if ($versionCount -ge 10) { 'Highly fragmented' } elseif ($versionCount -gt 1) { 'Fragmented' } else { 'Single version' }
        $riskPriority = if ($versionCount -ge 20) { 'High' } elseif ($versionCount -ge 10) { 'Medium' } elseif ($versionCount -gt 1) { 'Low' } else { 'Standardized' }
        $recommendedAction = switch ($riskPriority) {
            'High' { 'Prioritize packaging and migration to a smaller approved version set.' }
            'Medium' { 'Review deployment rings and reduce avoidable version spread.' }
            'Low' { 'Confirm the preferred version and monitor remaining secondary versions.' }
            default { 'Maintain the approved version baseline.' }
        }

        if ($versionCount -gt 1) { $fragmentedApplications++ }
        if ($versionCount -ge 10) { $highlyFragmentedApplications++ }

        foreach ($row in $group.Group) {
            $deviceCount = [int64]$row.DeviceCount
            $isDominant = ([string]$row.AppId -eq [string]$dominant.AppId)
            $outsideDominant = if ($isDominant) { [int64]0 } else { $deviceCount }
            $publisher = if ([string]::IsNullOrWhiteSpace([string]$row.AppPublisher)) { 'Unknown publisher' } else { ([string]$row.AppPublisher).Trim() }
            $publisherState = if ($publisher -match '(?i)^unknown') { 'Unknown' } else { 'Known' }
            $footprintState = if ($isDominant) { 'Dominant version' } elseif ($deviceCount -le 5) { 'Long tail version' } else { 'Secondary version' }

            $installationsOutsideDominant += $outsideDominant
            if ($footprintState -eq 'Long tail version') { $longTailVersions++ }
            if ($publisherState -eq 'Unknown') { $unknownPublisherVersions++ }

            if ($IncludeDetail) {
                $detail.Add([pscustomobject][ordered]@{
                    'Application Source ID' = [string]$row.AppId
                    'Application Product Key' = $productKey
                    'Application Name' = $applicationName
                    'Application Version' = if ([string]::IsNullOrWhiteSpace([string]$row.AppVersion)) { 'Unknown version' } else { ([string]$row.AppVersion).Trim() }
                    'Publisher' = $publisher
                    'Publisher State' = $publisherState
                    'Platform' = if ([string]::IsNullOrWhiteSpace([string]$row.Platform)) { 'Unknown' } else { ([string]$row.Platform).Trim().ToLowerInvariant() }
                    'Device Count' = $deviceCount
                    'Product Total Installs' = $productTotal
                    'Product Distinct Devices' = $distinctDevices
                    'Metric Definition' = if ($LegacyDefinition) { 'Legacy name / version observations' } else { 'Name + publisher + platform / distinct device footprint' }
                    'Collection Scope' = if ($LegacyDefinition) { 'Legacy Windows' } else { 'All platforms' }
                    'Product Version Count' = $versionCount
                    'Dominant Version Share (%)' = Convert-ToInvariantDecimalText $dominantShare
                    'Installs Outside Dominant Version' = $outsideDominant
                    'Standardization State' = $standardizationState
                    'Version Footprint State' = $footprintState
                    'Risk Priority' = $riskPriority
                    'Recommended Action' = $recommendedAction
                    'Snapshot Date' = $SnapshotDate.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
                    'Snapshot Week' = $SnapshotWeek
                    'Evidence Status' = 'Observed'
                })
            }
        }
    }

    $totalInstallations = if ($Rows.Count) { [int64](($Rows | Measure-Object DeviceCount -Sum).Sum) } else { [int64]0 }
    $publishers = @($Rows | ForEach-Object { [string]$_.AppPublisher } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique).Count
    $standardizationRate = if ($groups.Count -gt 0) { $fragmentedApplications / $groups.Count } else { $null }
    $dominantCoverage = if ($totalInstallations -gt 0) { 1 - ($installationsOutsideDominant / $totalInstallations) } else { $null }

    return [pscustomobject]@{
        Detail = @($detail)
        Trend = [pscustomobject][ordered]@{
            'Snapshot Date' = $SnapshotDate.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
            'Date Key' = [int]$SnapshotDate.ToString('yyyyMMdd', [Globalization.CultureInfo]::InvariantCulture)
            'Snapshot Week' = $SnapshotWeek
            'Applications' = $groups.Count
            'Application Versions' = $Rows.Count
            'Publishers' = $publishers
            'Application Install Observations' = $totalInstallations
            'Fragmented Applications' = $fragmentedApplications
            'Highly Fragmented Applications' = $highlyFragmentedApplications
            'Installations Outside Dominant Version' = $installationsOutsideDominant
            'Long Tail Versions' = $longTailVersions
            'Unknown Publisher Versions' = $unknownPublisherVersions
            'Standardization Opportunity Rate (%)' = Convert-ToInvariantDecimalText $standardizationRate
            'Dominant Version Coverage (%)' = Convert-ToInvariantDecimalText $dominantCoverage
            'Evidence Status' = 'Observed'
            'Metric Definition' = if ($LegacyDefinition) { 'Legacy name / version observations' } else { 'Name + publisher + platform / distinct device footprint' }
            'Collection Scope' = if ($LegacyDefinition) { 'Legacy Windows' } else { 'All platforms' }
        }
    }
}

$currentPath = Join-Path $DataRoot 'DATA-LAST\Intune_DiscoveredApps_Summary.csv'
$historyRoot = Join-Path $DataRoot 'DATA-ALL\Intune\Applications\DiscoveredApps\WeeklyHistory'
$summaryHistoryRoot = Join-Path $DataRoot 'DATA-ALL\Intune\Applications\DiscoveredApps\SummaryHistory'
if (-not (Test-Path -LiteralPath $currentPath -PathType Leaf)) { throw "Required source was not found: $currentPath" }
if (-not (Test-Path -LiteralPath $historyRoot -PathType Container)) { throw "Required history folder was not found: $historyRoot" }

$currentItem = Get-Item -LiteralPath $currentPath
$currentRows = @(Import-Csv -LiteralPath $currentPath)
$relationPath = Join-Path $DataRoot 'DATA-LAST/Intune_DiscoveredApps_AppDeviceRelations.csv'
$currentFootprint = Get-WorkplaceApplicationFootprint -Applications $currentRows -RelationPath $relationPath
$latestHistoryFolder = Get-ChildItem -LiteralPath $historyRoot -Directory | Sort-Object Name -Descending | Select-Object -First 1
$currentWeek = if ($latestHistoryFolder) { $latestHistoryFolder.Name } else { 'Current' }
$currentLegacy = @($currentRows | Where-Object { (Get-WorkplaceSourceValue $_ @('CollectionScope')) -ne 'AllPlatforms' }).Count -gt 0
$currentProfile = Get-ApplicationProfile -Rows $currentRows -SnapshotDate $currentItem.LastWriteTime.Date -SnapshotWeek $currentWeek -IncludeDetail -ProductFootprint $currentFootprint -LegacyDefinition:$currentLegacy

$summaryByWeek = @{}
foreach ($root in @($summaryHistoryRoot, $historyRoot)) {
    if (-not (Test-Path -LiteralPath $root -PathType Container)) { continue }
    foreach ($weekFolder in (Get-ChildItem -LiteralPath $root -Directory | Where-Object { $_.Name -match '^\d{4}-W\d{2}$' })) {
        $summaryPath = Join-Path $weekFolder.FullName 'Intune_DiscoveredApps_Summary.csv'
        if (Test-Path -LiteralPath $summaryPath -PathType Leaf) {
            $summaryByWeek[$weekFolder.Name] = $summaryPath
        }
    }
}

$trend = foreach ($week in @($summaryByWeek.Keys | Sort-Object)) {
    $summaryPath = $summaryByWeek[$week]
    $summaryItem = Get-Item -LiteralPath $summaryPath
    $snapshotDate = $summaryItem.LastWriteTime.Date
    if ([IO.Path]::GetDirectoryName([IO.Path]::GetDirectoryName($summaryPath)) -eq $summaryHistoryRoot) {
        $manifestPath = Join-Path $summaryItem.DirectoryName 'summary.manifest.json.txt'
        if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "Archived application summary manifest is missing: $manifestPath" }
        $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
        if ($manifest.Week -ne $week -or $manifest.FileName -ne $summaryItem.Name -or
            $manifest.SHA256 -ne (Get-FileHash -LiteralPath $summaryPath -Algorithm SHA256).Hash) {
            throw "Archived application summary failed provenance verification: $summaryPath"
        }
        $sourceTime = [datetime]::Parse([string]$manifest.SourceLastWriteTimeUtc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)
        $snapshotDate = $sourceTime.ToLocalTime().Date
    }
    $weeklyRows = @(Import-Csv -LiteralPath $summaryPath)
    $legacy = @($weeklyRows | Where-Object { (Get-WorkplaceSourceValue $_ @('CollectionScope')) -ne 'AllPlatforms' }).Count -gt 0
    $weeklyProfile = Get-ApplicationProfile -Rows $weeklyRows -SnapshotDate $snapshotDate -SnapshotWeek $week -LegacyDefinition:$legacy
    $weeklyProfile.Trend
}

foreach ($outputPath in @($InventoryOutputPath, $TrendOutputPath)) {
    $outputDirectory = Split-Path -Parent $outputPath
    if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $outputDirectory | Out-Null
    }
}

if ($currentProfile.Detail.Count) {
    $currentProfile.Detail | Export-Csv -LiteralPath $InventoryOutputPath -NoTypeInformation -Encoding utf8NoBOM
} else {
    $empty = [ordered]@{}
    foreach ($column in @('Application Source ID','Application Product Key','Application Name','Application Version','Publisher','Publisher State','Platform','Device Count','Product Total Installs','Product Distinct Devices','Metric Definition','Collection Scope','Product Version Count','Dominant Version Share (%)','Installs Outside Dominant Version','Standardization State','Version Footprint State','Risk Priority','Recommended Action','Snapshot Date','Snapshot Week','Evidence Status')) { $empty[$column] = '' }
    $header = @([pscustomobject]$empty | ConvertTo-Csv -NoTypeInformation)[0]
    [IO.File]::WriteAllText($InventoryOutputPath,$header+[Environment]::NewLine,[Text.UTF8Encoding]::new($false))
}
$trend | Export-Csv -LiteralPath $TrendOutputPath -NoTypeInformation -Encoding utf8NoBOM

[pscustomobject][ordered]@{
    InventoryOutputPath = $InventoryOutputPath
    InventoryRows = $currentProfile.Detail.Count
    TrendOutputPath = $TrendOutputPath
    TrendSnapshots = @($trend).Count
    Applications = $currentProfile.Trend.Applications
    ApplicationVersions = $currentProfile.Trend.'Application Versions'
    ApplicationInstallObservations = $currentProfile.Trend.'Application Install Observations'
    FragmentedApplications = $currentProfile.Trend.'Fragmented Applications'
    InstallationsOutsideDominantVersion = $currentProfile.Trend.'Installations Outside Dominant Version'
    DominantVersionCoverage = $currentProfile.Trend.'Dominant Version Coverage (%)'
} | ConvertTo-Json -Depth 3

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDUbY8z2fC1e86W
# BzKbt6GcrsnCTxr1qV27i6JlZnFBT6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCA4NbMcY8fNXjlX2iAY3PZ+
# DvsCelEJs3Gkb1rBWeXKnDANBgkqhkiG9w0BAQEFAASCAYAoOG75CDD41CsdSzbu
# LANCf0F2mb+Q7IXOQfljUT2L/OKDGrCu6mN11EPkjN5M/XCa3KjIBPzZkjLmi8QB
# BjWiNnda3tiChBwy4UJcqWAmXMyf6FEakQjiraYTDs6Ls6P5dH+xm4yhlKcG6Iwj
# JFq8TLV4AFxXSf8VATyv5i8KdbkBgRn3JqPSoZ253op7WvxCSKOVMWVpduXvZAfO
# rUp5Ni2EfGxF1a2a/mIAEcvywRRDR3xLeykAG0FD2WvzDjPlhpwwf//Y0CUNGDzf
# 4Wwe0V4UrHSWgwjPrQxoqyqdwc/K+4F1nEvYAQI7XAgvQxrQdz5T7OzYsJi2v1Yv
# JwVJYHIHB93rZjwxC0LQjzeJ+kb3gxLbtI8VSErhHs1YFjLIkjMRlmV9u0GdDMak
# QpXUPWbOUa1nqNu27nWnZYCqN//EsB1CGSj8OcSlAyoNqlWMhB9JOfZ2P2Okcfdm
# LnvK9CzZcdf2Q/NAipe/uvghB/Q8ej9END7sTSmAdYSHx2s=
# SIG # End signature block
