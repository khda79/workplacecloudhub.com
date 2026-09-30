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

-SnapshotPolicy Latest keeps the latest snapshot of each week and report instead, which is what the
prepared Workforce history reads and what the published batches contain. Its sources also include
the timestamped CSVs left in each run folder, so it works after the per-run WeeklyHistory folders
were removed, and it replaces a consolidated CSV that differs from the latest snapshot.

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

.PARAMETER SnapshotPolicy
First (default): earliest snapshot per week and report; existing consolidated files are kept.
Latest: latest snapshot per week and report, from per-run history and run-folder CSVs; a
consolidated CSV with different content is replaced.

.PARAMETER SynchronizedCopy
For a OneDrive/SharePoint synchronized copy only. Its folders are reparse points that the JSON
transport refuses, so manifests are written as plain files (temporary file then rename, with the
same bytes the share gets). Refused on UNC paths; requires -RecordedHistoryRootPath.

.PARAMETER Execute
Copies the planned files and writes the manifests. Omit for preview only.

.PARAMETER RemoveRunHistory
With -Execute, removes the per-run WeeklyHistory folders that are fully covered.

.VERSION
1.2

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

    [ValidateSet('First', 'Latest')]
    [string]$SnapshotPolicy = 'First',

    [switch]$SynchronizedCopy,

    [switch]$Execute,

    [switch]$RemoveRunHistory
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.2'

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
    $runCaptureUtc = Get-RunCaptureUtc -RunId $runFolder.Name
    if ($SnapshotPolicy -eq 'Latest') {
        # The run's own timestamped CSVs are its snapshots (the collector copies them to the history).
        $runWeek = Get-IsoWeekName -Date $runCaptureUtc.ToLocalTime()
        foreach ($csv in @(Get-ChildItem -LiteralPath $runFolder.FullName -File -Filter '*.csv')) {
            $historyName = $csv.Name -replace '_\d{8}[-_]\d{4,6}(?=\.csv$)', ''
            if ($historyName -eq $csv.Name -or $historyName -match '_Raw_') { continue }
            $candidates.Add([pscustomobject]@{
                Week = $runWeek
                FileName = $historyName
                RunId = $runFolder.Name
                SourcePath = $csv.FullName
                Length = $csv.Length
                CaptureUtc = ConvertTo-CaptureText -Utc $runCaptureUtc
                FromRunFolder = $true
            })
        }
    }
    if (-not (Test-Path -LiteralPath $runHistory -PathType Container)) { continue }
    $runHistoryFolders.Add([pscustomobject]@{ RunId = $runFolder.Name; RunFolder = $runFolder.FullName; HistoryFolder = $runHistory })
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

if ($SnapshotPolicy -eq 'Latest') {
    # Run-folder CSVs only re-select snapshots inside weeks that already have a history (consolidated
    # or per-run); they never create a week, such as runs from before the history was published.
    $historyWeeks = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($candidate in $candidates) { if (-not $candidate.PSObject.Properties['FromRunFolder']) { [void]$historyWeeks.Add($candidate.Week) } }
    if (Test-Path -LiteralPath $historyRoot -PathType Container) {
        foreach ($folder in @(Get-ChildItem -LiteralPath $historyRoot -Directory | Where-Object { $_.Name -match $weekPattern })) { [void]$historyWeeks.Add($folder.Name) }
    }
    $skippedWeeks = @($candidates | Where-Object { -not $historyWeeks.Contains($_.Week) } | Select-Object -ExpandProperty Week -Unique | Sort-Object)
    if ($skippedWeeks.Count -gt 0) { Write-Warning ("Weeks without history ignored by the Latest policy: {0}" -f ($skippedWeeks -join ', ')) }
    $retainedCandidates = @($candidates | Where-Object { $historyWeeks.Contains($_.Week) })
    $candidates = [Collections.Generic.List[object]]::new()
    foreach ($candidate in $retainedCandidates) { $candidates.Add($candidate) }
}

# --- Plan: one snapshot per week and report; existing consolidated files are kept (First) ------
$plan = [Collections.Generic.List[object]]::new()
foreach ($weekGroup in @($candidates | Group-Object Week | Sort-Object Name)) {
    $targetWeek = Join-Path -Path $historyRoot -ChildPath $weekGroup.Name
    foreach ($fileGroup in @($weekGroup.Group | Group-Object FileName | Sort-Object Name)) {
        $first = if ($SnapshotPolicy -eq 'Latest') {
            @($fileGroup.Group | Sort-Object RunId, CaptureUtc -Descending)[0]
        }
        else {
            @($fileGroup.Group | Sort-Object CaptureUtc, RunId)[0]
        }
        $targetPath = Join-Path -Path $targetWeek -ChildPath $first.FileName
        $action = if (-not (Test-Path -LiteralPath $targetPath -PathType Leaf)) { 'Copy' }
        elseif ($SnapshotPolicy -eq 'Latest' -and (Get-FileHash -LiteralPath $targetPath -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $first.SourcePath -Algorithm SHA256).Hash) { 'Replace' }
        else { 'KeepExisting' }
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

Write-Output ("Merge-SmartM365-M365UsageWeeklyHistory v{0} - {1} - snapshot policy {2}" -f $ScriptVersion, $(if ($Execute) { 'EXECUTE' } else { 'PREVIEW (read-only)' }), $SnapshotPolicy)
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
    foreach ($item in @($weekPlan.Group | Where-Object Action -in @('Copy', 'Replace'))) {
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
    $snapshotCreated = if ($SnapshotPolicy -eq 'Latest') {
        # Replaced files make the previous first-capture time meaningless; derive it from the files.
        if ($allTimed -and $knownTimes.Count -gt 0) { $knownTimes[0] } else { $null }
    }
    elseif ($null -ne $previous -and $previous.Document.PSObject.Properties['SnapshotCreatedAtUtc'] -and $previous.Document.SnapshotCreatedAtUtc) {
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
        ConsolidatedBy              = "Merge-SmartM365-M365UsageWeeklyHistory v$ScriptVersion ($SnapshotPolicy snapshot)"
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDf3LjQgwxvVeZ/
# 7p8oLSWKYCNR2XmVzGFkbgftOO519qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIAmyk+PpSTRtlXCFdmQ8mYQCiOfHF2GiAJMjIkeslw6DMA0GCSqG
# SIb3DQEBAQUABIIBgLBpVBbUFnsLwn6iLguOP40IpY5CKENXABM/mfXeVm/VIIrY
# Babnpz/snQTLPBU4ZfmpT4igr02O8A4NSl6xmjZsX0LbEHfl5DSi7ZwWjPg5RvCw
# KmZySrv1o/PjrRMp6EL3xJGQabjhvW8pQWeeC9zC6A7Mb2erRZwFTfTY5KPjdelU
# STsQ6pGh7FVAqaxYEAZYEcIDM3moiNSuhIxXrX6WxWAb2/8z6Q7wF9DN6myu8Bre
# 7hiRAdHbv6eRK6CJKaJQ4ODCAD/YxrNSxomCooHhHPWf3OxSgAM952Z2AgTaP1kX
# aBHGHsZuKopccpt1P8nomespSIiix8NNJ9LT8GJyVwl9cSOb99LxOPZ63sN2ZOta
# YgkQfCcbA1/MaHRCK7VTYKv1VZUiZoPsbOU9BPCHOZZa+/JGPvJbgxJkIr4XrLx6
# CUFdsr0gK003NoIjTA8wjTi2CR2ah5U2U/8hwFetLJbQR6OzBA0YLPRIMKGYYDJ4
# ucRVr9YEaHmpVrnpo6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MzAxODIz
# MjlaMC8GCSqGSIb3DQEJBDEiBCDIwEeTjbnWwAlAXSDSwx8Nzc3A88tq1mDFRbAm
# iDNwNTANBgkqhkiG9w0BAQEFAASCAgBG8KHem6ZRLAyMrksYBI1DdRHyVi37lLKF
# eo01e3lquNVdasp/s8MGK4hu2yd+Q9eEpAym8iuZmoQxwAW5wXeu5yTWZ9NUqrVf
# zvw2VOSLlv0uMAxlvPBMxvQQ0oyckxzXVSG6gO8QeDmGcKv48rjRdYMdyOjhut8j
# IF98AdY7Pf2+OLMydQcsVITFQuH+3SY74JVTWBBqGRic7qoCAjjmxSiqiTj79A/S
# HwfUBQcc/Zgj0nzhwqgJK2xue9125IQC+SRngkAKRo9m53wRJLJDaolcnmY6ypdB
# CLWZ/X772Q/Mv8C4779UYH/ge4F+d8GkKXttfEeJCIQfC5Edbo7HqodOmTLDS0UA
# 18KQs8BfaqdfeXWj6v4ctwMCTT+nbRqU1BuRcyXvSv3VfIxQMn3qf+Fjm1aEByOn
# GvfnmUcEufJih9fGt+X01TazChU4gUsmLgNudRpeHvt5ysek+0OJaQHCunGbIFI9
# GNjMHGJysHZdr/FxGn9cM8C6+C5ffae0zbGZYb8dEM8gz0vrWTmC5XenX+ow4LCS
# UsyzqEVPjGi2zugrYBGtcF2kMWaMXygGOXroBDODHp2uR3YKxKq+KiHwJQVnNglu
# 7Zw0GFAvzaXXifERieTSZPzfsAPMSK6k87blgjdc4Kr3DYEC/IynXfl+hmhbcDnL
# aOdqL+ETnA==
# SIG # End signature block
