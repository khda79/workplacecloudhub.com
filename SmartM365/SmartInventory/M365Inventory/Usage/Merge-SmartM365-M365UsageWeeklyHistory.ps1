<#
.SYNOPSIS
Consolidates the per-run M365 Usage weekly history folders into the collector WeeklyHistory root.

.DESCRIPTION
SmartM365-M365UserActivity-Inventory up to v1.18 wrote its weekly history under each run folder
(<Usage>\<yyyyMMdd_HHmmss>\WeeklyHistory\<week>) instead of <Usage>\WeeklyHistory, so every run
kept an isolated copy and no consolidated history was built.

This script rebuilds the consolidated history with the collector snapshot policy: for every ISO
week and every report CSV, the earliest captured snapshot is kept. Files already present in the
consolidated history are never overwritten; missing reports are added and the week manifest is
written through the SmartM365 JSON transport (owner 'WeeklyHistory:SmartM365 inventory').
Manifest content only depends on the source folders, so a production share and a synchronized
copy consolidated with the same RecordedHistoryRootPath get identical manifests.

A consolidated file already present without a recorded capture time (for example after an
interrupted run) keeps the provenance of the planned snapshot when both SHA256 hashes match.

Without -Execute the script is read-only and prints the plan.
With -Execute it copies each selected CSV, verifies its SHA256 and writes the week manifests.
With -Execute -RemoveRunHistory it then deletes a run WeeklyHistory folder only when every CSV in it
is present in the consolidated week and an identical CSV remains in a run folder.

.PARAMETER UsageRootPath
M365 Usage output folder, for example \\server\share\...\Tenants\prod\DATA-ALL\M365\Usage.

.PARAMETER RecordedHistoryRootPath
HistoryRootPath recorded in the manifests. Defaults to <UsageRootPath>\WeeklyHistory. Pass the
production UNC root when consolidating a synchronized copy.

.PARAMETER SynchronizedCopy
For a OneDrive/SharePoint synchronized copy only. Its folders are reparse points that the JSON
transport refuses, so manifests are written as plain files (temporary file then rename, with the
same bytes the share gets). Refused on UNC paths; requires -RecordedHistoryRootPath.

.PARAMETER Execute
Copies the planned files and writes the manifests. Omit for preview only.

.PARAMETER RemoveRunHistory
With -Execute, removes the per-run WeeklyHistory folders that are fully covered.

.VERSION
1.1

.NOTES
Author: https://github.com/khda79/workplacecloudhub.com
Run outside the collector schedule (M365-Usage-Inventory runs daily at 04:50).
#>
#requires -Version 7.2

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$UsageRootPath,

    [string]$RecordedHistoryRootPath = '',

    [switch]$SynchronizedCopy,

    [switch]$Execute,

    [switch]$RemoveRunHistory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.1'

if ($RemoveRunHistory -and -not $Execute) { throw '-RemoveRunHistory requires -Execute.' }
if ($SynchronizedCopy) {
    if ($UsageRootPath.StartsWith('\\')) { throw '-SynchronizedCopy is only for a local synchronized copy; the share must use the JSON transport.' }
    if ([string]::IsNullOrWhiteSpace($RecordedHistoryRootPath)) { throw '-SynchronizedCopy requires -RecordedHistoryRootPath (the production UNC history root).' }
}

$tenantContextPath = Join-Path -Path $PSScriptRoot -ChildPath '..\..\..\Config\SmartM365-TenantContext.ps1'
if (-not (Test-Path -LiteralPath $tenantContextPath -PathType Leaf)) { throw "SmartM365 tenant context not found: $tenantContextPath" }
. $tenantContextPath
Write-SmartM365StartupBanner

$coreModulePath = Join-Path -Path $PSScriptRoot -ChildPath '..\..\..\Modules\SmartM365.Core\SmartM365.Core.psd1'
Import-Module -Name $coreModulePath -MinimumVersion '1.0.58' -Force -ErrorAction Stop

$usageRoot = [IO.Path]::GetFullPath($UsageRootPath).TrimEnd('\', '/')
if (-not (Test-Path -LiteralPath $usageRoot -PathType Container)) { throw "Usage root not found: $usageRoot" }
$historyRoot = Join-Path -Path $usageRoot -ChildPath 'WeeklyHistory'
$recordedRoot = if ([string]::IsNullOrWhiteSpace($RecordedHistoryRootPath)) { $historyRoot } else { $RecordedHistoryRootPath.TrimEnd('\', '/') }
$historyLabel = 'SmartM365 inventory'
$manifestOwner = 'WeeklyHistory:' + $historyLabel
$weekPattern = '^\d{4}-W(?:0[1-9]|[1-4]\d|5[0-3])$'

function Get-IsoWeekName {
    # Same ISO-8601 rule as the internal SmartM365.Core helper used by the collectors.
    param([Parameter(Mandatory = $true)][datetime]$Date)
    $calendar = [Globalization.CultureInfo]::InvariantCulture.Calendar
    $dayOfWeek = $calendar.GetDayOfWeek($Date)
    if ($dayOfWeek -ge [DayOfWeek]::Monday -and $dayOfWeek -le [DayOfWeek]::Wednesday) { $Date = $Date.AddDays(3) }
    $week = $calendar.GetWeekOfYear($Date, [Globalization.CalendarWeekRule]::FirstFourDayWeek, [DayOfWeek]::Monday)
    return '{0}-W{1:00}' -f $Date.Year, $week
}

function Get-RunCaptureUtc {
    param([Parameter(Mandatory = $true)][string]$RunId)
    $local = [datetime]::ParseExact($RunId, 'yyyyMMdd_HHmmss', [Globalization.CultureInfo]::InvariantCulture)
    return [datetime]::SpecifyKind($local, [DateTimeKind]::Local).ToUniversalTime()
}

function ConvertTo-CaptureText {
    param([Parameter(Mandatory = $true)][datetime]$Utc)
    return $Utc.ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss.fffffff'Z'", [Globalization.CultureInfo]::InvariantCulture)
}

function ConvertTo-StoredTime {
    # Keeps recorded times as ISO UTC text even when ConvertFrom-Json returned a DateTime.
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) { return ConvertTo-CaptureText -Utc $Value }
    return [string]$Value
}

function Get-FileSnapshotText {
    # Per-file capture time recorded by the run manifest; the run id time otherwise.
    param($Manifest, [string]$FileName, [datetime]$RunCaptureUtc)
    if ($null -ne $Manifest -and $Manifest.PSObject.Properties['FileSnapshotCreatedAtUtc'] -and $null -ne $Manifest.FileSnapshotCreatedAtUtc) {
        $property = $Manifest.FileSnapshotCreatedAtUtc.PSObject.Properties[$FileName]
        if ($null -ne $property -and $null -ne $property.Value) {
            $value = $property.Value
            if ($value -is [datetime]) { return ConvertTo-CaptureText -Utc $value }
            $parsed = [datetime]::MinValue
            if ([datetime]::TryParse([string]$value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsed)) {
                return ConvertTo-CaptureText -Utc $parsed
            }
        }
    }
    return ConvertTo-CaptureText -Utc $RunCaptureUtc
}

function Write-SynchronizedManifest {
    # Plain atomic write for a synchronized copy; fails if the manifest changed since it was read.
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][byte[]]$Bytes, [Parameter(Mandatory = $true)][string]$ExpectedSHA256, [Parameter(Mandatory = $true)][string]$Week)
    $document = [Text.Encoding]::UTF8.GetString($Bytes) | ConvertFrom-Json
    if ($document.Week -ne $Week) { throw 'Weekly manifest week mismatch.' }
    $current = if (Test-Path -LiteralPath $Path -PathType Leaf) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash } else { 'ABSENT' }
    if ($current -ne $ExpectedSHA256) { throw "Weekly manifest changed since it was read: $Path" }
    $temporary = '{0}.{1}.tmp' -f $Path, [guid]::NewGuid().ToString('N')
    try {
        [IO.File]::WriteAllBytes($temporary, $Bytes)
        Move-Item -LiteralPath $temporary -Destination $Path -Force
    }
    finally {
        if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Force }
    }
}

function Read-OptionalManifest {
    param([Parameter(Mandatory = $true)][string]$WeekFolder)
    $legacy = Join-Path -Path $WeekFolder -ChildPath 'manifest.json'
    if (-not (Get-SmartM365JsonReadPath $legacy -Optional)) { return $null }
    return Read-SmartM365JsonDocument $legacy
}

# --- Inventory of the per-run history snapshots -------------------------------------------------
$candidates = [Collections.Generic.List[object]]::new()
$runHistoryFolders = [Collections.Generic.List[object]]::new()
foreach ($runFolder in @(Get-ChildItem -LiteralPath $usageRoot -Directory | Where-Object { $_.Name -match '^\d{8}_\d{6}$' } | Sort-Object Name)) {
    $runHistory = Join-Path -Path $runFolder.FullName -ChildPath 'WeeklyHistory'
    if (-not (Test-Path -LiteralPath $runHistory -PathType Container)) { continue }
    $runHistoryFolders.Add([pscustomobject]@{ RunId = $runFolder.Name; RunFolder = $runFolder.FullName; HistoryFolder = $runHistory })
    $runCaptureUtc = Get-RunCaptureUtc -RunId $runFolder.Name
    foreach ($weekFolder in @(Get-ChildItem -LiteralPath $runHistory -Directory | Where-Object { $_.Name -match $weekPattern })) {
        $expectedWeek = Get-IsoWeekName -Date $runCaptureUtc.ToLocalTime()
        if ($expectedWeek -ne $weekFolder.Name) {
            Write-Warning ("Run {0} stores week {1} but its run time belongs to {2}; the folder week is kept." -f $runFolder.Name, $weekFolder.Name, $expectedWeek)
        }
        $manifest = $null
        try { $document = Read-OptionalManifest -WeekFolder $weekFolder.FullName; if ($null -ne $document) { $manifest = $document.Document } }
        catch { Write-Warning ("Run {0} week {1}: manifest unreadable, run time used: {2}" -f $runFolder.Name, $weekFolder.Name, $_.Exception.Message) }
        foreach ($csv in @(Get-ChildItem -LiteralPath $weekFolder.FullName -File -Filter '*.csv')) {
            $candidates.Add([pscustomobject]@{
                Week = $weekFolder.Name
                FileName = $csv.Name
                RunId = $runFolder.Name
                SourcePath = $csv.FullName
                Length = $csv.Length
                CaptureUtc = Get-FileSnapshotText -Manifest $manifest -FileName $csv.Name -RunCaptureUtc $runCaptureUtc
            })
        }
    }
}

# --- Plan: earliest snapshot per week and report; existing consolidated files are kept ----------
$plan = [Collections.Generic.List[object]]::new()
foreach ($weekGroup in @($candidates | Group-Object Week | Sort-Object Name)) {
    $targetWeek = Join-Path -Path $historyRoot -ChildPath $weekGroup.Name
    foreach ($fileGroup in @($weekGroup.Group | Group-Object FileName | Sort-Object Name)) {
        $first = @($fileGroup.Group | Sort-Object CaptureUtc, RunId)[0]
        $targetPath = Join-Path -Path $targetWeek -ChildPath $first.FileName
        $action = if (Test-Path -LiteralPath $targetPath -PathType Leaf) { 'KeepExisting' } else { 'Copy' }
        $plan.Add([pscustomobject]@{
            Week = $weekGroup.Name
            FileName = $first.FileName
            SourceRun = $first.RunId
            CaptureUtc = $first.CaptureUtc
            SizeMB = [math]::Round($first.Length / 1MB, 1)
            Snapshots = $fileGroup.Count
            Action = $action
            SourcePath = $first.SourcePath
            TargetPath = $targetPath
        })
    }
}

Write-Output ("Merge-SmartM365-M365UsageWeeklyHistory v{0} - {1}" -f $ScriptVersion, $(if ($Execute) { 'EXECUTE' } else { 'PREVIEW (read-only)' }))
Write-Output ("Usage root: {0}" -f $usageRoot)
Write-Output ("Consolidated history root: {0} (recorded as {1})" -f $historyRoot, $recordedRoot)
Write-Output ("Run history folders: {0}; snapshots found: {1}; weeks: {2}" -f $runHistoryFolders.Count, $candidates.Count, @($plan | Select-Object -ExpandProperty Week -Unique).Count)
$plan | Format-Table Week, FileName, SourceRun, CaptureUtc, SizeMB, Snapshots, Action -AutoSize | Out-String -Width 220 | Write-Output

if (Test-Path -LiteralPath $historyRoot -PathType Container) {
    foreach ($existingWeek in @(Get-ChildItem -LiteralPath $historyRoot -Directory | Where-Object { $_.Name -match $weekPattern } | Sort-Object Name)) {
        try {
            $existing = Read-OptionalManifest -WeekFolder $existingWeek.FullName
            $root = if ($null -ne $existing -and $existing.Document.PSObject.Properties['HistoryRootPath']) { [string]$existing.Document.HistoryRootPath } else { '<no manifest>' }
            if ($root -ne $recordedRoot) { Write-Warning ("Existing consolidated week {0} was recorded for another root: {1}" -f $existingWeek.Name, $root) }
        }
        catch { Write-Warning ("Existing consolidated week {0} has an unreadable manifest: {1}" -f $existingWeek.Name, $_.Exception.Message) }
    }
}

if (-not $Execute) {
    Write-Output 'Preview only. Re-run with -Execute to consolidate, then with -Execute -RemoveRunHistory to remove covered per-run folders.'
    return
}

# --- Execution: copy, verify, write manifests ---------------------------------------------------
$copied = 0
foreach ($weekPlan in @($plan | Group-Object Week | Sort-Object Name)) {
    $weekFolder = Join-Path -Path $historyRoot -ChildPath $weekPlan.Name
    New-Item -ItemType Directory -Path $weekFolder -Force | Out-Null

    $previous = $null
    $previousHash = 'ABSENT'
    if ($SynchronizedCopy) {
        $manifestPath = Join-Path -Path $weekFolder -ChildPath 'manifest.json.txt'
        if (Test-Path -LiteralPath (Join-Path -Path $weekFolder -ChildPath 'manifest.json') -PathType Leaf) {
            throw "Legacy manifest.json present in $weekFolder; consolidate the share first, it migrates legacy manifests."
        }
        if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
            $previous = Read-SmartM365JsonDocument $manifestPath
            $previousHash = (Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash
        }
    }
    else {
        $manifestPath = Resolve-SmartM365OwnedJsonPath -Path (Join-Path -Path $weekFolder -ChildPath 'manifest.json') -Owner $manifestOwner -Validate {
            param($document)
            if ($document -isnot [pscustomobject]) { throw 'Weekly manifest must be an object.' }
        }
        if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
            $previous = Read-SmartM365JsonDocument $manifestPath
            $previousHash = $previous.SHA256
        }
    }

    $fileTimes = [ordered]@{}
    $sources = [ordered]@{}
    if ($null -ne $previous -and $previous.Document.PSObject.Properties['FileSnapshotCreatedAtUtc'] -and $null -ne $previous.Document.FileSnapshotCreatedAtUtc) {
        foreach ($property in $previous.Document.FileSnapshotCreatedAtUtc.PSObject.Properties) { $fileTimes[$property.Name] = ConvertTo-StoredTime -Value $property.Value }
    }
    if ($null -ne $previous -and $previous.Document.PSObject.Properties['ConsolidatedFromRuns'] -and $null -ne $previous.Document.ConsolidatedFromRuns) {
        foreach ($property in $previous.Document.ConsolidatedFromRuns.PSObject.Properties) { $sources[$property.Name] = [string]$property.Value }
    }

    $changed = $null -eq $previous
    # Recover the provenance of consolidated files left without a recorded time (interrupted run).
    foreach ($item in @($weekPlan.Group | Where-Object Action -eq 'KeepExisting')) {
        if ($fileTimes.Contains($item.FileName)) { continue }
        if ((Get-FileHash -LiteralPath $item.SourcePath -Algorithm SHA256).Hash -eq (Get-FileHash -LiteralPath $item.TargetPath -Algorithm SHA256).Hash) {
            $fileTimes[$item.FileName] = $item.CaptureUtc
            $sources[$item.FileName] = $item.SourceRun
            $changed = $true
        }
        else {
            Write-Warning ("Week {0}: existing {1} differs from the earliest snapshot; its capture time stays unknown." -f $weekPlan.Name, $item.FileName)
        }
    }
    foreach ($item in @($weekPlan.Group | Where-Object Action -eq 'Copy')) {
        Copy-SmartM365FileAtomically -SourcePath $item.SourcePath -DestinationPath $item.TargetPath
        if ((Get-FileHash -LiteralPath $item.SourcePath -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $item.TargetPath -Algorithm SHA256).Hash) {
            throw "Copy verification failed: $($item.TargetPath)"
        }
        $fileTimes[$item.FileName] = $item.CaptureUtc
        $sources[$item.FileName] = $item.SourceRun
        $copied++
        $changed = $true
    }
    if (-not $changed) { continue }

    $files = @(Get-ChildItem -LiteralPath $weekFolder -File -Filter '*.csv' | Sort-Object Name | ForEach-Object { $_.Name })
    $sortedTimes = [ordered]@{}
    foreach ($name in $files) { if ($fileTimes.Contains($name)) { $sortedTimes[$name] = $fileTimes[$name] } }
    $sortedSources = [ordered]@{}
    foreach ($name in $files) { if ($sources.Contains($name)) { $sortedSources[$name] = $sources[$name] } }
    $knownTimes = @($sortedTimes.Values | Sort-Object)
    $allTimed = $knownTimes.Count -eq $files.Count
    $snapshotCreated = if ($null -ne $previous -and $previous.Document.PSObject.Properties['SnapshotCreatedAtUtc'] -and $previous.Document.SnapshotCreatedAtUtc) {
        ConvertTo-StoredTime -Value $previous.Document.SnapshotCreatedAtUtc
    }
    elseif ($allTimed -and $knownTimes.Count -gt 0) { $knownTimes[0] } else { $null }

    $manifest = [pscustomobject][ordered]@{
        UpdatedAt                   = $(if ($knownTimes.Count -gt 0) { $knownTimes[-1] } else { $null })
        Week                        = $weekPlan.Name
        HistoryLabel                = $historyLabel
        HistoryRootPath             = $recordedRoot
        SnapshotPolicy              = 'FirstSnapshotPerIsoWeekUnlessOverwriteExisting'
        SnapshotCreatedAtUtc        = $snapshotCreated
        SnapshotTimestampStatus     = $(if ($snapshotCreated) { 'Recorded' } else { 'UnknownLegacy' })
        FileSnapshotCreatedAtUtc    = $sortedTimes
        Files                       = $files
        ConsolidatedFromRuns        = $sortedSources
        ConsolidatedBy              = "Merge-SmartM365-M365UsageWeeklyHistory v$ScriptVersion"
    }
    $week = $weekPlan.Name
    $bytes = [Text.UTF8Encoding]::new($false).GetBytes(($manifest | ConvertTo-Json -Depth 7))
    if ($SynchronizedCopy) {
        Write-SynchronizedManifest -Path $manifestPath -Bytes $bytes -ExpectedSHA256 $previousHash -Week $week
    }
    else {
        $null = Write-SmartM365JsonBytesAtomically -Path $manifestPath -Bytes $bytes -ExpectedSHA256 $previousHash -Validate {
            param($document)
            if ($document.Week -ne $week) { throw 'Weekly manifest week mismatch.' }
        }.GetNewClosure()
    }
    Write-Output ("Week {0}: {1} file(s) in consolidated history; manifest {2}" -f $weekPlan.Name, $files.Count, $manifestPath)
}
Write-Output ("Copied and verified: {0} file(s)." -f $copied)

if (-not $RemoveRunHistory) {
    Write-Output 'Per-run WeeklyHistory folders kept. Re-run with -Execute -RemoveRunHistory to remove the covered ones.'
    return
}

# --- Removal of fully covered per-run history folders -------------------------------------------
# Run CSVs indexed by size; hashes are computed lazily. Concurrent manual runs can publish the same
# DATA-LAST file, so a history snapshot may match the CSV of another run rather than its own.
$runCsvBySize = @{}
foreach ($run in $runHistoryFolders) {
    foreach ($runCsv in @(Get-ChildItem -LiteralPath $run.RunFolder -File -Filter '*.csv')) {
        if (-not $runCsvBySize.ContainsKey($runCsv.Length)) { $runCsvBySize[$runCsv.Length] = [Collections.Generic.List[object]]::new() }
        $runCsvBySize[$runCsv.Length].Add([pscustomobject]@{ Path = $runCsv.FullName; RunFolder = $run.RunFolder; Hash = $null })
    }
}
function Test-IdenticalRunCsv {
    param([Parameter(Mandatory = $true)][IO.FileInfo]$HistoryCsv, [Parameter(Mandatory = $true)][string]$OwnRunFolder)
    if (-not $runCsvBySize.ContainsKey($HistoryCsv.Length)) { return $false }
    $hash = (Get-FileHash -LiteralPath $HistoryCsv.FullName -Algorithm SHA256).Hash
    # Own run first: it is the expected match and avoids hashing other runs.
    $entries = @($runCsvBySize[$HistoryCsv.Length])
    foreach ($ownRunPass in @($true, $false)) {
        foreach ($entry in $entries) {
            if (($entry.RunFolder -eq $OwnRunFolder) -ne $ownRunPass) { continue }
            if ($null -eq $entry.Hash) { $entry.Hash = (Get-FileHash -LiteralPath $entry.Path -Algorithm SHA256).Hash }
            if ($entry.Hash -eq $hash) { return $true }
        }
    }
    return $false
}

$removed = 0
$kept = 0
foreach ($run in $runHistoryFolders) {
    $reason = ''
    foreach ($historyCsv in @(Get-ChildItem -LiteralPath $run.HistoryFolder -File -Filter '*.csv' -Recurse)) {
        $week = $historyCsv.Directory.Name
        if (-not (Test-Path -LiteralPath (Join-Path -Path (Join-Path -Path $historyRoot -ChildPath $week) -ChildPath $historyCsv.Name) -PathType Leaf)) {
            $reason = "$week\$($historyCsv.Name) missing from the consolidated history"
            break
        }
        if (-not (Test-IdenticalRunCsv -HistoryCsv $historyCsv -OwnRunFolder $run.RunFolder)) {
            $reason = "$week\$($historyCsv.Name) has no identical CSV in any run folder"
            break
        }
    }
    if ($reason) {
        Write-Warning ("Run {0}: WeeklyHistory kept ({1})." -f $run.RunId, $reason)
        $kept++
        continue
    }
    Remove-Item -LiteralPath $run.HistoryFolder -Recurse -Force
    $removed++
}
Write-Output ("Per-run WeeklyHistory folders removed: {0}; kept: {1}." -f $removed, $kept)
