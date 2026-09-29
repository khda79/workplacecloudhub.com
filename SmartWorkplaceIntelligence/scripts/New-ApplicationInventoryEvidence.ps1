[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DataRoot,

    [string]$InventoryOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\ApplicationInventoryEvidence.csv'),

    [string]$TrendOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\ApplicationTrendEvidence.csv')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Convert-ToInvariantDecimalText {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    return ([double]$Value).ToString('0.0########', [Globalization.CultureInfo]::InvariantCulture)
}

function Get-ApplicationProfile {
    param(
        [Parameter(Mandatory = $true)][object[]]$Rows,
        [Parameter(Mandatory = $true)][datetime]$SnapshotDate,
        [Parameter(Mandatory = $true)][string]$SnapshotWeek,
        [switch]$IncludeDetail
    )

    $detail = [System.Collections.Generic.List[object]]::new()
    $fragmentedApplications = 0
    $highlyFragmentedApplications = 0
    $installationsOutsideDominant = [int64]0
    $longTailVersions = 0
    $unknownPublisherVersions = 0

    $groups = @($Rows | Group-Object AppName)
    foreach ($group in $groups) {
        $applicationName = if ([string]::IsNullOrWhiteSpace($group.Name)) { 'Unknown application' } else { $group.Name.Trim() }
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
                    'Application Name' = $applicationName
                    'Application Version' = if ([string]::IsNullOrWhiteSpace([string]$row.AppVersion)) { 'Unknown version' } else { ([string]$row.AppVersion).Trim() }
                    'Publisher' = $publisher
                    'Publisher State' = $publisherState
                    'Platform' = if ([string]::IsNullOrWhiteSpace([string]$row.Platform)) { 'Unknown' } else { ([string]$row.Platform).Trim().ToLowerInvariant() }
                    'Device Count' = $deviceCount
                    'Product Total Installs' = $productTotal
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

    $totalInstallations = [int64](($Rows | Measure-Object DeviceCount -Sum).Sum)
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
        }
    }
}

$currentPath = Join-Path $DataRoot 'DATA-LAST\Intune_DiscoveredApps_Summary.csv'
$historyRoot = Join-Path $DataRoot 'DATA-ALL\Intune\Applications\DiscoveredApps\WeeklyHistory'
if (-not (Test-Path -LiteralPath $currentPath -PathType Leaf)) { throw "Required source was not found: $currentPath" }
if (-not (Test-Path -LiteralPath $historyRoot -PathType Container)) { throw "Required history folder was not found: $historyRoot" }

$currentItem = Get-Item -LiteralPath $currentPath
$currentRows = @(Import-Csv -LiteralPath $currentPath)
$latestHistoryFolder = Get-ChildItem -LiteralPath $historyRoot -Directory | Sort-Object Name -Descending | Select-Object -First 1
$currentWeek = if ($latestHistoryFolder) { $latestHistoryFolder.Name } else { 'Current' }
$currentProfile = Get-ApplicationProfile -Rows $currentRows -SnapshotDate $currentItem.LastWriteTime.Date -SnapshotWeek $currentWeek -IncludeDetail

$trend = foreach ($weekFolder in (Get-ChildItem -LiteralPath $historyRoot -Directory | Sort-Object Name)) {
    $summaryPath = Join-Path $weekFolder.FullName 'Intune_DiscoveredApps_Summary.csv'
    if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf)) { continue }
    $summaryItem = Get-Item -LiteralPath $summaryPath
    $weeklyProfile = Get-ApplicationProfile -Rows @(Import-Csv -LiteralPath $summaryPath) -SnapshotDate $summaryItem.LastWriteTime.Date -SnapshotWeek $weekFolder.Name
    $weeklyProfile.Trend
}

foreach ($outputPath in @($InventoryOutputPath, $TrendOutputPath)) {
    $outputDirectory = Split-Path -Parent $outputPath
    if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $outputDirectory | Out-Null
    }
}

$currentProfile.Detail | Export-Csv -LiteralPath $InventoryOutputPath -NoTypeInformation -Encoding utf8NoBOM
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
