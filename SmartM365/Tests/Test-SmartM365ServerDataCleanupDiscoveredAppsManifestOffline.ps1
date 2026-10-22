<#
.SYNOPSIS
Offline check of the Discovered Apps summary manifest comparison of Invoke-SmartM365-ServerDataCleanup.
.VERSION
1.0.0
.NOTES
Builds a synthetic tenant DATA root under TEMP with five recent Discovered Apps weeks and five older
ones, then runs the cleanup script with -Execute in a child PowerShell 7 process. An older week
without a summary manifest, or with an existing correct one (as written by an earlier pass), is
removed; a manifest with a different SHA256, a different date or an unreadable date fails and keeps
the week intact. No tenant, network or synchronized data access.
#>
#requires -Version 7.2
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$smartM365Root = Split-Path $PSScriptRoot -Parent
$cleanupScript = Join-Path $smartM365Root 'SmartInventory/Maintenance/Invoke-SmartM365-ServerDataCleanup.ps1'
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }

$base = Join-Path ([IO.Path]::GetTempPath()) ('sm-apps-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$root = Join-Path $base 'prod'
$reportFolder = Join-Path $base 'report'
$apps = Join-Path $root 'DATA-ALL/Intune/Applications/DiscoveredApps'
$weekly = Join-Path $apps 'WeeklyHistory'
$summaryHistory = Join-Path $apps 'SummaryHistory'
$summaryName = 'Intune_DiscoveredApps_Summary.csv'
# Seven fractional digits: the manifest date must round-trip at full tick precision.
$sourceTime = [datetime]::new(2025, 1, 6, 8, 1, 2, [DateTimeKind]::Utc).AddTicks(1234567)

function Write-SyntheticWeek {
    param([string]$Week, [string]$Content)
    $folder = Join-Path $weekly $Week
    [void][IO.Directory]::CreateDirectory($folder)
    [IO.File]::WriteAllText((Join-Path $folder 'Intune_DiscoveredApps.csv'), 'detail')
    $summary = Join-Path $folder $summaryName
    [IO.File]::WriteAllText($summary, $Content)
    [IO.File]::SetLastWriteTimeUtc($summary, $sourceTime)
    return $folder
}

function Write-SummaryManifest {
    param([string]$Week, [string]$Sha256, [string]$SourceLastWriteTimeUtc)
    $folder = Join-Path $summaryHistory $Week
    [void][IO.Directory]::CreateDirectory($folder)
    $manifest = [ordered]@{
        SchemaVersion = 1; Week = $Week; FileName = $summaryName
        SHA256 = $Sha256; SourceLastWriteTimeUtc = $SourceLastWriteTimeUtc
    }
    $path = Join-Path $folder 'summary.manifest.json.txt'
    $manifest | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $path -Encoding utf8 -NoNewline
    return $path
}

try {
    [void][IO.Directory]::CreateDirectory((Join-Path $root 'DATA-LAST'))
    foreach ($offset in 0..4) {
        $date = (Get-Date).AddDays(-7 * $offset)
        $null = Write-SyntheticWeek ('{0}-W{1:00}' -f [Globalization.ISOWeek]::GetYear($date), [Globalization.ISOWeek]::GetWeekOfYear($date)) 'recent'
    }
    $weeks = [ordered]@{
        NoManifest     = '2025-W02'
        CorrectExisting = '2025-W03'
        OtherSha       = '2025-W04'
        OtherDate      = '2025-W05'
        UnreadableDate = '2025-W06'
    }
    $folders = @{}
    foreach ($name in $weeks.Keys) { $folders[$name] = Write-SyntheticWeek $weeks[$name] ('summary ' + $name) }
    $hashes = @{}
    foreach ($name in $weeks.Keys) { $hashes[$name] = (Get-FileHash -LiteralPath (Join-Path $folders[$name] $summaryName) -Algorithm SHA256).Hash }
    $isoTime = $sourceTime.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
    $correctManifest = Write-SummaryManifest -Week $weeks.CorrectExisting -Sha256 $hashes.CorrectExisting -SourceLastWriteTimeUtc $isoTime
    $correctManifestHash = (Get-FileHash -LiteralPath $correctManifest -Algorithm SHA256).Hash
    $null = Write-SummaryManifest -Week $weeks.OtherSha -Sha256 ('0' * 64) -SourceLastWriteTimeUtc $isoTime
    $null = Write-SummaryManifest -Week $weeks.OtherDate -Sha256 $hashes.OtherDate -SourceLastWriteTimeUtc $sourceTime.AddSeconds(1).ToString('o', [Globalization.CultureInfo]::InvariantCulture)
    $null = Write-SummaryManifest -Week $weeks.UnreadableDate -Sha256 $hashes.UnreadableDate -SourceLastWriteTimeUtc 'not a date'

    $arguments = @('-NoProfile', '-NonInteractive', '-File', $cleanupScript, '-Tenant', 'prod', '-DataRootPath', $root,
        '-ReportDirectory', $reportFolder, '-EffectiveJobsPath', (Join-Path $base 'no-jobs.json'), '-Execute')
    $output = & pwsh @arguments 2>&1 | Out-String
    Check ($LASTEXITCODE -ne 0) "Execute did not report the mismatched manifests: $output"
    $status = @{}
    foreach ($row in @(Import-Csv (Join-Path $reportFolder 'verified-actions.csv') | Where-Object Category -eq 'DiscoveredAppsWeek')) {
        $status[(Split-Path $row.RelativePath -Leaf)] = $row.Status
    }
    Check ($status.Count -eq 5) "Planned Discovered Apps weeks: $(@($status.Keys) -join ', ')"

    foreach ($name in 'NoManifest', 'CorrectExisting') {
        Check ($status[$weeks[$name]] -eq 'Removed') "$name week status: $($status[$weeks[$name]])"
        Check (-not (Test-Path -LiteralPath $folders[$name])) "$name week folder was not removed."
        $kept = Join-Path (Join-Path $summaryHistory $weeks[$name]) $summaryName
        Check ((Test-Path -LiteralPath $kept) -and (Get-FileHash -LiteralPath $kept -Algorithm SHA256).Hash -eq $hashes[$name]) "$name summary was not preserved."
        Check ((Get-Item -LiteralPath $kept).LastWriteTimeUtc -eq $sourceTime) "$name summary lost its source time."
    }
    $written = Get-Content -LiteralPath (Join-Path (Join-Path $summaryHistory $weeks.NoManifest) 'summary.manifest.json.txt') -Raw
    Check ($written -match [regex]::Escape('"SourceLastWriteTimeUtc": "' + $isoTime + '"')) 'The new manifest does not record the source time.'
    Check ((Get-FileHash -LiteralPath $correctManifest -Algorithm SHA256).Hash -eq $correctManifestHash) 'The existing correct manifest was rewritten.'

    foreach ($name in 'OtherSha', 'OtherDate', 'UnreadableDate') {
        Check ($status[$weeks[$name]] -eq 'Failed: Discovered Apps summary manifest mismatch.') "$name week status: $($status[$weeks[$name]])"
        Check (Test-Path -LiteralPath (Join-Path $folders[$name] 'Intune_DiscoveredApps.csv')) "$name week detail was removed."
        Check ((Get-FileHash -LiteralPath (Join-Path $folders[$name] $summaryName) -Algorithm SHA256).Hash -eq $hashes[$name]) "$name week summary changed."
    }
}
finally {
    if (Test-Path -LiteralPath $base) { Remove-Item -LiteralPath $base -Recurse -Force }
}
"Test-SmartM365ServerDataCleanupDiscoveredAppsManifestOffline: $script:checks checks passed."
