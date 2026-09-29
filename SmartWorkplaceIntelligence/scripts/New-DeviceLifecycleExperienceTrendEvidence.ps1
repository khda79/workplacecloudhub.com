[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
    [string]$DataRoot,

    [string]$WindowsOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\WindowsLifecycleTrendEvidence.csv'),

    [string]$EndpointOutputPath = (Join-Path (Split-Path -Parent $PSScriptRoot) '_private\EndpointExperienceTrendEvidence.csv')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Convert-ToDoubleOrNull {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return $null }

    $parsed = 0.0
    if ([double]::TryParse([string]$Value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
        return $parsed
    }
    if ([double]::TryParse([string]$Value, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::GetCultureInfo('fr-FR'), [ref]$parsed)) {
        return $parsed
    }
    return $null
}

function Convert-ToInvariantDecimalText {
    param([AllowNull()][object]$Value)
    if ($null -eq $Value) { return $null }
    return ([double]$Value).ToString('0.0########', [Globalization.CultureInfo]::InvariantCulture)
}

function Get-WindowsGeneration {
    param([AllowNull()][object]$OperatingSystem, [AllowNull()][object]$OperatingSystemVersion)
    if ([string]$OperatingSystem -notmatch '(?i)windows') { return 'Not Windows' }
    $parts = ([string]$OperatingSystemVersion).Split('.')
    $build = 0
    if ($parts.Count -ge 3 -and [int]::TryParse($parts[2], [ref]$build)) {
        if ($build -ge 22000) { return 'Windows 11' }
        if ($build -ge 10240) { return 'Windows 10' }
    }
    return 'Windows legacy or unknown'
}

function Get-LatestRowsByDevice {
    param(
        [object[]]$Rows,
        [string]$PreferredProperty
    )

    $index = @{}
    foreach ($row in $Rows) {
        $key = ([string]$row.DeviceId).Trim().ToLowerInvariant()
        if (-not $key) { continue }
        if (-not $index.ContainsKey($key)) {
            $index[$key] = $row
            continue
        }

        $candidateHasValue = -not [string]::IsNullOrWhiteSpace([string]$row.$PreferredProperty)
        $currentHasValue = -not [string]::IsNullOrWhiteSpace([string]$index[$key].$PreferredProperty)
        $candidateDate = [datetime]::MinValue
        $currentDate = [datetime]::MinValue
        [void][datetime]::TryParse([string]$row.ReportRefreshDate, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$candidateDate)
        [void][datetime]::TryParse([string]$index[$key].ReportRefreshDate, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$currentDate)
        if (($candidateHasValue -and -not $currentHasValue) -or
            ($candidateHasValue -eq $currentHasValue -and $candidateDate -gt $currentDate)) {
            $index[$key] = $row
        }
    }
    return @($index.Values)
}

function Get-AverageOrNull {
    param([object[]]$Values)
    $numbers = @($Values | Where-Object { $null -ne $_ })
    if ($numbers.Count -eq 0) { return $null }
    return [math]::Round((($numbers | Measure-Object -Average).Average), 1)
}

$windowsHistoryRoot = Join-Path $DataRoot 'DATA-ALL\Intune\Devices\Inventory\WeeklyHistory'
$endpointHistoryRoot = Join-Path $DataRoot 'DATA-ALL\Intune\EndpointAnalytics\WeeklyHistory'
foreach ($requiredRoot in @($windowsHistoryRoot, $endpointHistoryRoot)) {
    if (-not (Test-Path -LiteralPath $requiredRoot -PathType Container)) {
        throw "Required weekly-history folder was not found: $requiredRoot"
    }
}

$windowsTrend = foreach ($weekFolder in (Get-ChildItem -LiteralPath $windowsHistoryRoot -Directory | Sort-Object Name)) {
    $inventoryPath = Join-Path $weekFolder.FullName 'Intune_Devices_Inventory.csv'
    if (-not (Test-Path -LiteralPath $inventoryPath -PathType Leaf)) { continue }

    $inventory = @(Import-Csv -LiteralPath $inventoryPath)
    $windowsRows = @($inventory | Where-Object { $_.OS -match '(?i)windows' })
    $windows10 = @($windowsRows | Where-Object { (Get-WindowsGeneration $_.OS $_.'OS version') -eq 'Windows 10' }).Count
    $windows11 = @($windowsRows | Where-Object { (Get-WindowsGeneration $_.OS $_.'OS version') -eq 'Windows 11' }).Count
    $knownModernWindows = $windows10 + $windows11
    $snapshotDate = (Get-Item -LiteralPath $inventoryPath).LastWriteTime.Date

    [pscustomobject][ordered]@{
        'Snapshot Date' = $snapshotDate.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        'Snapshot Week' = $weekFolder.Name
        'Managed Devices' = $inventory.Count
        'Windows Devices' = $windowsRows.Count
        'Windows 10 Devices' = $windows10
        'Windows 11 Devices' = $windows11
        'Windows 11 Adoption (%)' = Convert-ToInvariantDecimalText $(if ($knownModernWindows -gt 0) { $windows11 / $knownModernWindows } else { $null })
        'Evidence Status' = 'Observed'
    }
}

$endpointTrend = foreach ($weekFolder in (Get-ChildItem -LiteralPath $endpointHistoryRoot -Directory | Sort-Object Name)) {
    $performancePath = Join-Path $weekFolder.FullName 'Intune_EndpointAnalytics_DevicePerformance.csv'
    $startupPath = Join-Path $weekFolder.FullName 'Intune_EndpointAnalytics_StartupDevices.csv'
    if (-not (Test-Path -LiteralPath $performancePath -PathType Leaf) -or -not (Test-Path -LiteralPath $startupPath -PathType Leaf)) { continue }

    $performance = Get-LatestRowsByDevice -Rows @(Import-Csv -LiteralPath $performancePath) -PreferredProperty 'EndpointAnalyticsScore'
    $startup = Get-LatestRowsByDevice -Rows @(Import-Csv -LiteralPath $startupPath) -PreferredProperty 'StartupScore'
    $endpointScores = @($performance | ForEach-Object { Convert-ToDoubleOrNull $_.EndpointAnalyticsScore } | Where-Object { $null -ne $_ })
    $startupScores = @($startup | ForEach-Object { Convert-ToDoubleOrNull $_.StartupScore } | Where-Object { $null -ne $_ })
    $appScores = @($performance | ForEach-Object { Convert-ToDoubleOrNull $_.AppReliabilityScore } | Where-Object { $null -ne $_ -and $_ -ge 0 })
    $observedDates = @($performance + $startup | ForEach-Object {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse([string]$_.ReportRefreshDate, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsed)) { $parsed }
    })
    $snapshotDate = if ($observedDates.Count -gt 0) { ($observedDates | Sort-Object -Descending | Select-Object -First 1).Date } else { (Get-Item -LiteralPath $performancePath).LastWriteTime.Date }

    [pscustomobject][ordered]@{
        'Snapshot Date' = $snapshotDate.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        'Snapshot Week' = $weekFolder.Name
        'Endpoint Analytics Covered Devices' = $endpointScores.Count
        'Endpoint Analytics Score' = Convert-ToInvariantDecimalText (Get-AverageOrNull $endpointScores)
        'Endpoint Analytics Below 50 Devices' = @($endpointScores | Where-Object { $_ -lt 50 }).Count
        'Startup Covered Devices' = $startupScores.Count
        'Startup Score' = Convert-ToInvariantDecimalText (Get-AverageOrNull $startupScores)
        'Startup Below 50 Devices' = @($startupScores | Where-Object { $_ -lt 50 }).Count
        'App Reliability Covered Devices' = $appScores.Count
        'App Reliability Score' = Convert-ToInvariantDecimalText (Get-AverageOrNull $appScores)
        'Devices with Stop Errors' = @($startup | Where-Object { $null -ne (Convert-ToDoubleOrNull $_.StopErrorCount) -and (Convert-ToDoubleOrNull $_.StopErrorCount) -gt 0 }).Count
        'Devices with Boot Time over 60s' = @($startup | Where-Object { $null -ne (Convert-ToDoubleOrNull $_.CoreBootTime) -and (Convert-ToDoubleOrNull $_.CoreBootTime) -gt 60 }).Count
        'Devices with Sign-in Time over 30s' = @($startup | Where-Object { $null -ne (Convert-ToDoubleOrNull $_.CoreSignInTime) -and (Convert-ToDoubleOrNull $_.CoreSignInTime) -gt 30 }).Count
        'Evidence Status' = 'Observed'
    }
}

foreach ($outputPath in @($WindowsOutputPath, $EndpointOutputPath)) {
    $outputDirectory = Split-Path -Parent $outputPath
    if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
        New-Item -ItemType Directory -Path $outputDirectory | Out-Null
    }
}

$windowsTrend | Export-Csv -LiteralPath $WindowsOutputPath -NoTypeInformation -Encoding utf8NoBOM
$endpointTrend | Export-Csv -LiteralPath $EndpointOutputPath -NoTypeInformation -Encoding utf8NoBOM

[pscustomobject][ordered]@{
    WindowsOutputPath = $WindowsOutputPath
    WindowsSnapshots = $windowsTrend.Count
    EndpointOutputPath = $EndpointOutputPath
    EndpointSnapshots = $endpointTrend.Count
    LatestWindowsSnapshot = if ($windowsTrend.Count -gt 0) { $windowsTrend[-1].'Snapshot Date' } else { $null }
    LatestEndpointSnapshot = if ($endpointTrend.Count -gt 0) { $endpointTrend[-1].'Snapshot Date' } else { $null }
} | ConvertTo-Json -Depth 3
