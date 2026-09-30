<#
.SYNOPSIS
Removes the timestamped run copies of inventory CSVs that WeeklyHistory already covers.

.DESCRIPTION
Up to SmartM365.Core 1.0.60, Publish-SmartM365Csv uploaded every timestamped run copy
(<prefix>_yyyyMMdd_HHmmss.csv) to SharePoint while the seven-day retention only ran on the server,
so DATA-ALL on SharePoint kept every run. Core 1.0.61 applies the retention on SharePoint too;
this script removes the copies accumulated before.

A timestamped CSV is removed only when all of these hold:
- its run timestamp is older than RetentionDays;
- its folder has a WeeklyHistory folder;
- the ISO week of the timestamp is recorded there: the week manifest names the folder's week,
  lists <prefix>.csv, and <prefix>.csv exists in the week folder with content;
- <prefix>.csv exists under the DATA-LAST folder next to DATA-ALL (current copy).
Never considered: WeeklyHistory, Corrections, Archive, Orchestrator and run folders
(<yyyyMMdd_HHmmss>, for example the M365 Usage runs), non-timestamped files and non-CSV files.
Every file that is not removed is reported with its reason.

Without -Execute the script is read-only and prints the plan. With -Execute it deletes the planned
files and writes a CSV report of every decision. On a OneDrive/SharePoint synchronized copy,
deleted files go to the SharePoint recycle bin.

.PARAMETER DataAllRootPath
DATA-ALL root, for example \\server\share\...\Tenants\prod\DATA-ALL or the synchronized copy
...\SMART-M365\DATA\DATA-ALL.

.PARAMETER RetentionDays
Minimum age of a removed copy, from its file-name timestamp. Default 7 (collector retention).

.PARAMETER ReportPath
CSV report of every decision. Defaults to %TEMP%\SmartM365-ExpiredTimestampedCsv_<timestamp>.csv.

.PARAMETER Execute
Deletes the planned files. Omit for preview only.

.VERSION
1.0

.NOTES
Author: https://github.com/khda79/workplacecloudhub.com
#>
#requires -Version 7.2

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$DataAllRootPath,

    [ValidateRange(1, 3650)]
    [int]$RetentionDays = 7,

    [string]$ReportPath = '',

    [switch]$Execute
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.0'

$tenantContextPath = Join-Path -Path $PSScriptRoot -ChildPath '..\..\Config\SmartM365-TenantContext.ps1'
if (-not (Test-Path -LiteralPath $tenantContextPath -PathType Leaf)) { throw "SmartM365 tenant context not found: $tenantContextPath" }
. $tenantContextPath
Write-SmartM365StartupBanner

$dataAllRoot = [IO.Path]::GetFullPath($DataAllRootPath).TrimEnd('\', '/')
if (-not (Test-Path -LiteralPath $dataAllRoot -PathType Container)) { throw "DATA-ALL root not found: $dataAllRoot" }
if ((Split-Path -Path $dataAllRoot -Leaf) -ne 'DATA-ALL') { throw "DataAllRootPath must be a DATA-ALL folder: $dataAllRoot" }
$dataLastRoot = Join-Path -Path (Split-Path -Path $dataAllRoot -Parent) -ChildPath 'DATA-LAST'
if (-not (Test-Path -LiteralPath $dataLastRoot -PathType Container)) { throw "DATA-LAST folder not found next to DATA-ALL: $dataLastRoot" }
if ([string]::IsNullOrWhiteSpace($ReportPath)) {
    $ReportPath = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ('SmartM365-ExpiredTimestampedCsv_{0}.csv' -f (Get-Date).ToString('yyyyMMdd_HHmmss'))
}

$now = Get-Date
$cutoff = $now.AddDays(-$RetentionDays)
$namePattern = '^(?<Prefix>.+)_(?<Stamp>\d{8}[-_]\d{6})\.csv$'
$stampFormats = [string[]]@('yyyyMMdd_HHmmss', 'yyyyMMdd-HHmmss')
$excludedSegments = @('WeeklyHistory', 'Corrections', 'Archive', 'Orchestrator')

function Get-IsoWeekName {
    # Same ISO-8601 rule as the internal SmartM365.Core helper used by the collectors.
    param([Parameter(Mandatory = $true)][datetime]$Date)
    $calendar = [Globalization.CultureInfo]::InvariantCulture.Calendar
    $dayOfWeek = $calendar.GetDayOfWeek($Date)
    if ($dayOfWeek -ge [DayOfWeek]::Monday -and $dayOfWeek -le [DayOfWeek]::Wednesday) { $Date = $Date.AddDays(3) }
    $week = $calendar.GetWeekOfYear($Date, [Globalization.CalendarWeekRule]::FirstFourDayWeek, [DayOfWeek]::Monday)
    return '{0}-W{1:00}' -f $Date.Year, $week
}

$script:WeekCache = @{}
function Get-RecordedWeek {
    # Plain read of the week manifest (.json.txt or .json); works on a share and on a synchronized copy.
    param([Parameter(Mandatory = $true)][string]$HistoryRoot, [Parameter(Mandatory = $true)][string]$Week)
    $key = $HistoryRoot + '|' + $Week
    if ($script:WeekCache.ContainsKey($key)) { return $script:WeekCache[$key] }
    $weekFolder = Join-Path -Path $HistoryRoot -ChildPath $Week
    $record = [pscustomobject]@{ Folder = $weekFolder; Files = @(); Error = 'WeekNotRecorded' }
    $manifestPath = @('manifest.json.txt', 'manifest.json' | ForEach-Object { Join-Path $weekFolder $_ } | Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } | Select-Object -First 1)
    if ($manifestPath.Count) {
        try {
            $manifest = Get-Content -LiteralPath $manifestPath[0] -Raw -Encoding utf8 | ConvertFrom-Json
            if ($manifest.PSObject.Properties['Week'] -and [string]$manifest.Week -eq $Week -and $manifest.PSObject.Properties['Files']) {
                $record = [pscustomobject]@{ Folder = $weekFolder; Files = @($manifest.Files | ForEach-Object { [string]$_ }); Error = '' }
            }
            else { $record.Error = 'InvalidManifest' }
        }
        catch { $record.Error = 'UnreadableManifest' }
    }
    $script:WeekCache[$key] = $record
    return $record
}

Write-Host ("Scanning {0} (retention {1} days, cutoff {2:yyyy-MM-dd HH:mm})..." -f $dataAllRoot, $RetentionDays, $cutoff)
$latestNames = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($latest in @(Get-ChildItem -LiteralPath $dataLastRoot -Recurse -File -Filter '*.csv' -ErrorAction Stop)) { [void]$latestNames.Add($latest.Name) }

$decisions = [Collections.Generic.List[object]]::new()
foreach ($file in @(Get-ChildItem -LiteralPath $dataAllRoot -Recurse -File -Filter '*.csv' -ErrorAction Stop)) {
    $match = [regex]::Match($file.Name, $namePattern, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if (-not $match.Success) { continue }
    $relativeFolder = [IO.Path]::GetRelativePath($dataAllRoot, $file.DirectoryName)
    $segments = @($relativeFolder -split '[\\/]')
    $reason = ''
    $week = ''
    if (@($segments | Where-Object { $_ -in $excludedSegments -or $_ -match '^\d{8}[-_]\d{6}$' }).Count) { $reason = 'ExcludedFolder' }
    $stamp = [datetime]::MinValue
    if (-not $reason -and -not [datetime]::TryParseExact($match.Groups['Stamp'].Value, $stampFormats, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$stamp)) { $reason = 'InvalidTimestamp' }
    if (-not $reason -and $stamp -ge $cutoff) { $reason = 'Recent' }
    $latestName = $match.Groups['Prefix'].Value + '.csv'
    if (-not $reason) {
        $historyRoot = Join-Path -Path $file.DirectoryName -ChildPath 'WeeklyHistory'
        if (-not (Test-Path -LiteralPath $historyRoot -PathType Container)) { $reason = 'NoWeeklyHistory' }
        else {
            $week = Get-IsoWeekName -Date $stamp
            $recorded = Get-RecordedWeek -HistoryRoot $historyRoot -Week $week
            if ($recorded.Error) { $reason = $recorded.Error }
            elseif (-not @($recorded.Files | Where-Object { $_ -ieq $latestName }).Count) { $reason = 'NotInWeekManifest' }
            else {
                $weekCopy = Get-Item -LiteralPath (Join-Path $recorded.Folder $latestName) -ErrorAction SilentlyContinue
                if ($null -eq $weekCopy -or $weekCopy.Length -eq 0) { $reason = 'WeekCopyMissing' }
            }
        }
    }
    if (-not $reason -and -not $latestNames.Contains($latestName)) { $reason = 'NoLatestCopy' }
    $decisions.Add([pscustomobject]@{
        Action = if ($reason) { 'Keep' } else { 'Remove' }; Reason = $reason; Folder = $relativeFolder; Name = $file.Name
        Week = $week; SizeBytes = $file.Length; Result = ''; FullName = $file.FullName
    })
}

$toRemove = @($decisions | Where-Object Action -eq 'Remove')
$kept = @($decisions | Where-Object Action -eq 'Keep')
Write-Host ''
Write-Host ("Timestamped CSV found: {0}; to remove: {1} ({2:N0} MB); kept: {3}." -f $decisions.Count, $toRemove.Count, (($toRemove | Measure-Object SizeBytes -Sum).Sum / 1MB), $kept.Count)
if ($toRemove.Count) {
    $toRemove | Group-Object Folder | ForEach-Object {
        [pscustomobject]@{ Folder = $_.Name; Files = $_.Count; MB = [math]::Round(($_.Group | Measure-Object SizeBytes -Sum).Sum / 1MB, 1); Weeks = (@($_.Group.Week | Sort-Object -Unique) -join ',') }
    } | Sort-Object MB -Descending | Format-Table -AutoSize | Out-String -Width 240 | Write-Host
}
if ($kept.Count) {
    Write-Host 'Kept, by reason:'
    $kept | Group-Object Reason | ForEach-Object {
        [pscustomobject]@{ Reason = $_.Name; Files = $_.Count; MB = [math]::Round(($_.Group | Measure-Object SizeBytes -Sum).Sum / 1MB, 1); Folders = (@($_.Group.Folder | Sort-Object -Unique | Select-Object -First 4) -join '; ') }
    } | Sort-Object Files -Descending | Format-Table -AutoSize | Out-String -Width 240 | Write-Host
}

if ($Execute) {
    $failed = 0
    foreach ($decision in $toRemove) {
        try { Remove-Item -LiteralPath $decision.FullName -Force -ErrorAction Stop; $decision.Result = 'Removed' }
        catch { $decision.Result = 'Failed: ' + $_.Exception.Message; $failed++ }
    }
    Write-Host ("Removed: {0}; failed: {1}." -f ($toRemove.Count - $failed), $failed)
}
else {
    Write-Host 'Preview only. Run again with -Execute to delete the files listed as to remove.'
}
$decisions | Select-Object Action, Reason, Result, Folder, Name, Week, SizeBytes | Export-Csv -LiteralPath $ReportPath -NoTypeInformation -Encoding utf8
Write-Host ("Report: {0} (script v{1})" -f $ReportPath, $ScriptVersion)
if ($Execute -and $failed -gt 0) { exit 1 }
