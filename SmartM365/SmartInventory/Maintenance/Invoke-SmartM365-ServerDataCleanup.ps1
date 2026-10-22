<#
.SYNOPSIS
Audits a complete SmartM365 tenant DATA root and retires verified redundant files.

.DESCRIPTION
Preview is read-only. Execute removes only verified legacy Windows 11 readiness
archives, old Discovered Apps detail weeks after preserving their summaries,
byte-identical old M365 Usage run folders, timestamped CSV copies qualified by
weekly manifests, orphan atomic-write temporaries left by interrupted writers
and, only when enabled, linked obsolete prepared batches after retaining their
receipts. Every other file is inventoried for review. The script is standalone.

The server share keeps the prepared-batch history by default: linked batches
older than current and previous are inventoried as kept. -RetireLinkedPreparedBatches
retires them, beyond -RetainedLinkedPreparedBatchCount further batches and only
once older than -LinkedPreparedBatchMinimumAgeDays (age from the batch id).

Orphan temporaries are removed only when the name exactly matches a known
writer pattern, the file is older than the minimum age, and no JSON transport
lock in its folder is held. The JSON transport keeps its staging and previous
copies on failure; nothing resumes them, so they only age out here.

.VERSION
1.2.0

.EXAMPLE
./Invoke-SmartM365-ServerDataCleanup.ps1 -Tenant prod -DataRootPath '<server DATA root>'
./Invoke-SmartM365-ServerDataCleanup.ps1 -Tenant prod -DataRootPath '<server DATA root>' -Execute
./Invoke-SmartM365-ServerDataCleanup.ps1 -Tenant prod -DataRootPath '<server DATA root>' -RetireLinkedPreparedBatches -RetainedLinkedPreparedBatchCount 10 -LinkedPreparedBatchMinimumAgeDays 90
#>
#requires -Version 7.2
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$DataRootPath,
    [string]$ReportDirectory = '',
    [string]$EffectiveJobsPath = '',
    # A JSON transport .previous copy can be the only prior version: it gets a longer age.
    [ValidateRange(1, 3650)][int]$OrphanTemporaryMinimumAgeDays = 7,
    [ValidateRange(1, 3650)][int]$OrphanPreviousMinimumAgeDays = 30,
    # The share keeps the prepared-batch history unless retirement is enabled explicitly.
    [switch]$RetireLinkedPreparedBatches,
    [ValidateRange(0, 100000)][int]$RetainedLinkedPreparedBatchCount = 0,
    [ValidateRange(0, 36500)][int]$LinkedPreparedBatchMinimumAgeDays = 0,
    [switch]$Execute
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
# The same script also runs from a temporary server folder without the repo.
$tenantContextPath = Join-Path $PSScriptRoot '..\..\Config\SmartM365-TenantContext.ps1'
if (Test-Path -LiteralPath $tenantContextPath -PathType Leaf) {
    . $tenantContextPath
    Write-SmartM365StartupBanner
}

$root = [IO.Path]::GetFullPath($DataRootPath).TrimEnd('\', '/')
if (-not (Test-Path -LiteralPath $root -PathType Container)) { throw "DATA root not found: $root" }
if (-not [string]::Equals((Split-Path $root -Leaf), $Tenant, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'The DATA root leaf must match -Tenant.'
}
if ((Get-Item -LiteralPath $root -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
    throw 'A reparse-point DATA root is not supported.'
}
if ([string]::IsNullOrWhiteSpace($ReportDirectory)) {
    $ReportDirectory = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-ServerDataCleanup_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
}
$reportRoot = [IO.Path]::GetFullPath($ReportDirectory).TrimEnd('\', '/')
$rootPrefix = $root + [IO.Path]::DirectorySeparatorChar
if ($reportRoot.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase) -or
    [string]::Equals($reportRoot, $root, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'ReportDirectory must be outside the DATA root.'
}
New-Item -ItemType Directory -Path $reportRoot -Force | Out-Null

function Assert-OwnedPath {
    param([Parameter(Mandatory)][string]$Path)
    $full = [IO.Path]::GetFullPath($Path)
    if (-not $full.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Cleanup target is outside the DATA root: $full"
    }
    $relative = [IO.Path]::GetRelativePath($root, $full)
    $cursor = $root
    foreach ($part in ($relative -split '[\\/]')) {
        $cursor = Join-Path $cursor $part
        if ((Test-Path -LiteralPath $cursor) -and
            ((Get-Item -LiteralPath $cursor -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)) {
            throw "Cleanup target contains a reparse point: $cursor"
        }
    }
    return $full
}

function Get-Hash {
    param([string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
}

function Get-IsoWeek {
    param([datetime]$Date)
    return '{0}-W{1:00}' -f [Globalization.ISOWeek]::GetYear($Date), [Globalization.ISOWeek]::GetWeekOfYear($Date)
}

function Get-DirectoryState {
    param([string]$Path)
    $entries = @(Get-ChildItem -LiteralPath $Path -Recurse -Force |
        Sort-Object FullName |
        ForEach-Object {
            [pscustomobject]@{
                Path = [IO.Path]::GetRelativePath($Path, $_.FullName)
                Directory = $_.PSIsContainer
                Bytes = if ($_.PSIsContainer) { 0 } else { $_.Length }
                WriteTicks = $_.LastWriteTimeUtc.Ticks
            }
        })
    return (ConvertTo-Json -InputObject $entries -Compress -Depth 3)
}

function Get-TimestampedCsvDecisions {
    $cutoff = (Get-Date).AddDays(-7)
    $latestNames = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    # Source staging holds unpromoted copies: they do not prove a latest export.
    foreach ($latest in @(Get-ChildItem -LiteralPath $dataLast -File -Recurse -Filter '*.csv' | Where-Object { $_.FullName -notmatch '[\\/]\.staging[\\/]' })) {
        [void]$latestNames.Add($latest.Name)
    }
    $weekCache = @{}
    foreach ($file in @(Get-ChildItem -LiteralPath $dataAll -File -Recurse -Filter '*.csv')) {
        $match = [regex]::Match($file.Name, '^(?<Prefix>.+)_(?<Stamp>\d{8}[-_]\d{6})\.csv$',
            [Text.RegularExpressions.RegexOptions]::IgnoreCase)
        if (-not $match.Success) { continue }
        $relativeFolder = [IO.Path]::GetRelativePath($dataAll, $file.DirectoryName)
        $segments = @($relativeFolder -split '[\\/]')
        $reason = ''
        $week = ''
        if (@($segments | Where-Object { $_ -in @('WeeklyHistory','Corrections','Archive','Orchestrator') -or
            $_ -match '^\d{8}[-_]\d{6}$' }).Count) { $reason = 'ExcludedFolder' }
        $stamp = [datetime]::MinValue
        if (-not $reason -and -not [datetime]::TryParseExact($match.Groups['Stamp'].Value,
            [string[]]@('yyyyMMdd_HHmmss','yyyyMMdd-HHmmss'), [Globalization.CultureInfo]::InvariantCulture,
            [Globalization.DateTimeStyles]::None, [ref]$stamp)) { $reason = 'InvalidTimestamp' }
        if (-not $reason -and $stamp -ge $cutoff) { $reason = 'Recent' }
        $latestName = $match.Groups['Prefix'].Value + '.csv'
        if (-not $reason) {
            $historyRoot = Join-Path $file.DirectoryName 'WeeklyHistory'
            if (-not (Test-Path -LiteralPath $historyRoot -PathType Container)) { $reason = 'NoWeeklyHistory' }
            else {
                $week = Get-IsoWeek $stamp
                $key = $historyRoot + '|' + $week
                if (-not $weekCache.ContainsKey($key)) {
                    $weekFolder = Join-Path $historyRoot $week
                    $manifestPath = @('manifest.json.txt','manifest.json' |
                        ForEach-Object { Join-Path $weekFolder $_ } |
                        Where-Object { Test-Path -LiteralPath $_ -PathType Leaf } |
                        Select-Object -First 1)
                    $record = [pscustomobject]@{ Folder = $weekFolder; Files = @(); Error = 'WeekNotRecorded' }
                    if ($manifestPath.Count) {
                        try {
                            $manifest = Get-Content -LiteralPath $manifestPath[0] -Raw -Encoding utf8 | ConvertFrom-Json
                            if ($manifest.PSObject.Properties['Week'] -and $manifest.PSObject.Properties['Files'] -and
                                [string]$manifest.Week -eq $week) {
                                $record = [pscustomobject]@{ Folder = $weekFolder; Files = @($manifest.Files | ForEach-Object { [string]$_ }); Error = '' }
                            } else { $record.Error = 'InvalidManifest' }
                        } catch { $record.Error = 'UnreadableManifest' }
                    }
                    $weekCache[$key] = $record
                }
                $recorded = $weekCache[$key]
                if ($recorded.Error) { $reason = $recorded.Error }
                elseif (-not @($recorded.Files | Where-Object { $_ -ieq $latestName }).Count) { $reason = 'NotInWeekManifest' }
                else {
                    $weekCopy = Get-Item -LiteralPath (Join-Path $recorded.Folder $latestName) -ErrorAction SilentlyContinue
                    if ($null -eq $weekCopy -or $weekCopy.Length -eq 0) { $reason = 'WeekCopyMissing' }
                }
            }
        }
        if (-not $reason -and -not $latestNames.Contains($latestName)) { $reason = 'NoLatestCopy' }
        if (-not $reason) {
            try { $null = Assert-OwnedPath $file.FullName }
            catch { $reason = 'ReparseOrOutsideRoot' }
        }
        [pscustomobject]@{
            Action = if ($reason) { 'Keep' } else { 'Remove' }
            Reason = $reason; Result = ''; Folder = $relativeFolder
            Name = $file.Name; Week = $week; SizeBytes = $file.Length
            FullName = $file.FullName; LastWriteTimeUtc = $file.LastWriteTimeUtc
        }
    }
}

# Exact names written by the Core, JsonTransport, CmdbReceipt, DailyCaptureRetention and the
# Windows PowerShell 5.1 module. Legacy forms predate the short ~sm names (JsonTransport 1.0.8).
# The legacy <name>.<guid><extension> CSV form is ambiguous and stays a review item.
$orphanTemporaryPatterns = @(
    [pscustomobject]@{ Name = 'Short'; Regex = '^~sm[0-9a-f]{12}\.(?<Kind>pending|previous|tmp|partial|download|csv)$' }
    [pscustomobject]@{ Name = 'LegacyTenantKey'; Regex = '^.+\.tenantkey\.[0-9a-f]{32}\.(?<Kind>tmp)$' }
    [pscustomobject]@{ Name = 'LegacySibling'; Regex = '^.+\.[0-9a-f]{32}\.(?<Kind>pending|previous|tmp|partial|download)$' }
    [pscustomobject]@{ Name = 'LegacyWriteProbe'; Regex = '^\.smartm365-write-test-[0-9a-f]{32}\.(?<Kind>tmp)$' }
    [pscustomobject]@{ Name = 'LegacyPreflightProbe'; Regex = '^\.smartm365-preflight-[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\.(?<Kind>tmp)$' }
)

function Open-FolderTransportLockSet {
    # A temporary name does not carry its target, so every JSON transport lock of the
    # folder is held: no transport writer of that folder can run while they are open.
    param([Parameter(Mandatory)][string]$Folder)
    $handles = [Collections.Generic.List[IO.FileStream]]::new()
    try {
        foreach ($lock in @(Get-ChildItem -LiteralPath $Folder -File -Force -Filter '*.transport.lock' |
            Where-Object { $_.Name.EndsWith('.transport.lock', [StringComparison]::OrdinalIgnoreCase) })) {
            if ($lock.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Linked transport lock: $($lock.FullName)" }
            $handles.Add([IO.File]::Open($lock.FullName, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None))
        }
    } catch {
        foreach ($handle in $handles) { $handle.Dispose() }
        throw
    }
    return ,$handles
}

function Get-OrphanTemporaryDecision {
    $now = [datetime]::UtcNow
    $lockStates = @{}
    foreach ($file in @(Get-ChildItem -LiteralPath $root -File -Recurse -Force)) {
        # Weekly upload markers are publication state, never a temporary.
        if ($file.Name -ieq 'upload.pending') { continue }
        $pattern = ''
        $kind = ''
        foreach ($candidate in $orphanTemporaryPatterns) {
            $match = [regex]::Match($file.Name, $candidate.Regex, [Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if ($match.Success) { $pattern = $candidate.Name; $kind = $match.Groups['Kind'].Value.ToLowerInvariant(); break }
        }
        if (-not $pattern) { continue }
        # A copied file keeps an old write time; its creation time still shows recent activity.
        $lastActivity = if ($file.CreationTimeUtc -gt $file.LastWriteTimeUtc) { $file.CreationTimeUtc } else { $file.LastWriteTimeUtc }
        $ageDays = ($now - $lastActivity).TotalDays
        $minimumDays = if ($kind -eq 'previous') { $OrphanPreviousMinimumAgeDays } else { $OrphanTemporaryMinimumAgeDays }
        $reason = ''
        if ($ageDays -lt $minimumDays) { $reason = 'Recent' }
        if (-not $reason) {
            try {
                $null = Assert-OwnedPath $file.FullName
                if ($file.Attributes -band [IO.FileAttributes]::ReparsePoint) { $reason = 'ReparseOrOutsideRoot' }
            } catch { $reason = 'ReparseOrOutsideRoot' }
        }
        if (-not $reason) {
            if (-not $lockStates.ContainsKey($file.DirectoryName)) {
                $state = ''
                try { foreach ($handle in (Open-FolderTransportLockSet $file.DirectoryName)) { $handle.Dispose() } }
                catch [IO.IOException] { $state = 'TransportLockHeld' }
                catch { $state = 'TransportLockUnavailable' }
                $lockStates[$file.DirectoryName] = $state
            }
            $reason = $lockStates[$file.DirectoryName]
        }
        [pscustomobject]@{
            Action = if ($reason) { 'Keep' } else { 'Remove' }
            Reason = $reason; Result = ''; Pattern = $pattern; Kind = $kind
            Folder = [IO.Path]::GetRelativePath($root, $file.DirectoryName)
            Name = $file.Name; SizeBytes = $file.Length; MinimumAgeDays = $minimumDays
            AgeDays = [Math]::Round($ageDays, 1); FullName = $file.FullName
            LastWriteTimeUtc = $file.LastWriteTimeUtc
        }
    }
}

function Write-CleanupLog {
    param([string]$Message)
    Write-Information ('[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Message) -InformationAction Continue
}

$actions = [Collections.Generic.List[object]]::new()
function Add-Action {
    param([string]$Category, [string]$Path, [int64]$Bytes, [string]$Reason, [object]$Evidence)
    $full = Assert-OwnedPath $Path
    $actions.Add([pscustomobject]@{
        Category = $Category
        Path = $full
        RelativePath = [IO.Path]::GetRelativePath($root, $full)
        Bytes = $Bytes
        Reason = $Reason
        Evidence = $Evidence
        Status = 'Planned'
    })
}

$dataAll = Join-Path $root 'DATA-ALL'
$dataLast = Join-Path $root 'DATA-LAST'
if (-not (Test-Path -LiteralPath $dataAll -PathType Container) -or
    -not (Test-Path -LiteralPath $dataLast -PathType Container)) {
    throw 'DATA-ALL and DATA-LAST must exist below the DATA root.'
}

# Readiness archives are a retired output; current detail, summary, and weekly
# history must exist before the old daily archive is eligible.
$readiness = Join-Path $dataAll 'Intune\WindowsUpdate\Windows11ReadinessIssues'
$archive = Join-Path $readiness 'Archive'
$readyLatest = Join-Path $dataLast 'Intune_Windows11_Readiness_Issues.csv'
$readySummary = Join-Path $dataLast 'Intune_Windows11_Readiness_Issues_Summary.csv'
$readyWeekly = Join-Path $readiness 'WeeklyHistory'
if ((Test-Path -LiteralPath $archive -PathType Container) -and
    (Test-Path -LiteralPath $readyLatest -PathType Leaf) -and
    (Test-Path -LiteralPath $readySummary -PathType Leaf) -and
    (Test-Path -LiteralPath $readyWeekly -PathType Container) -and
    (Get-Item -LiteralPath $readyLatest).Length -gt 0 -and
    (Get-Item -LiteralPath $readySummary).Length -gt 0 -and
    @(Get-ChildItem -LiteralPath $readyWeekly -File -Recurse |
        Where-Object { $_.Name -like 'Intune_Windows11_Readiness_Issues*.csv' -and $_.Length -gt 0 }).Count) {
    foreach ($file in @(Get-ChildItem -LiteralPath $archive -File -Force)) {
        if ($file.Name -match '^Intune_Windows11_Readiness_Issues_\d{8}[-_]\d{6}\.csv$' -and
            -not ($file.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
            Add-Action 'ReadinessArchive' $file.FullName $file.Length 'Legacy daily archive; current detail, summary and weekly history exist.' @{
                Length = $file.Length; LastWriteTimeUtc = $file.LastWriteTimeUtc
            }
        }
    }
    if (-not @(Get-ChildItem -LiteralPath $archive -Directory -Force).Count -and
        -not @(Get-ChildItem -LiteralPath $archive -File -Force |
            Where-Object { $_.Name -notmatch '^Intune_Windows11_Readiness_Issues_\d{8}[-_]\d{6}\.csv$' -or
                ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) }).Count) {
        Add-Action 'EmptyReadinessArchive' $archive 0 'Remove the retired archive directory after its daily CSV files are gone.' @{}
    }
}

# Preserve the weekly summary and its source timestamp before deleting detail.
$apps = Join-Path $dataAll 'Intune\Applications\DiscoveredApps'
$appsWeekly = Join-Path $apps 'WeeklyHistory'
$appsSummary = Join-Path $apps 'SummaryHistory'
if (Test-Path -LiteralPath $appsWeekly -PathType Container) {
    $weekFolders = @(Get-ChildItem -LiteralPath $appsWeekly -Directory -Force |
        Where-Object { $_.Name -match '^\d{4}-W\d{2}$' })
    $currentWeek = Get-IsoWeek (Get-Date)
    $keepWeeks = @(@($weekFolders.Name) + $currentWeek | Sort-Object -Unique -Descending | Select-Object -First 5)
    foreach ($weekFolder in @($weekFolders | Where-Object { $_.Name -notin $keepWeeks } | Sort-Object Name)) {
        if ($weekFolder.Attributes -band [IO.FileAttributes]::ReparsePoint) { continue }
        if (Test-Path -LiteralPath (Join-Path $weekFolder.FullName 'upload.pending')) { continue }
        $source = Join-Path $weekFolder.FullName 'Intune_DiscoveredApps_Summary.csv'
        if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { continue }
        $summaryFile = Get-Item -LiteralPath $source -Force
        if ($summaryFile.Attributes -band [IO.FileAttributes]::ReparsePoint) { continue }
        $hash = Get-Hash $source
        $destination = Join-Path (Join-Path $appsSummary $weekFolder.Name) $summaryFile.Name
        if ((Test-Path -LiteralPath $destination) -and (Get-Hash $destination) -ne $hash) { continue }
        $allFiles = @(Get-ChildItem -LiteralPath $weekFolder.FullName -File -Recurse -Force)
        $subdirs = @(Get-ChildItem -LiteralPath $weekFolder.FullName -Directory -Recurse -Force)
        if (@($allFiles | Where-Object { $_.Attributes -band [IO.FileAttributes]::ReparsePoint }).Count -or
            @($subdirs | Where-Object { $_.Attributes -band [IO.FileAttributes]::ReparsePoint }).Count) { continue }
        $bytes = [int64](($allFiles | Measure-Object Length -Sum).Sum)
        Add-Action 'DiscoveredAppsWeek' $weekFolder.FullName $bytes 'Older than the five retained detail weeks; summary will be SHA256-preserved.' @{
            Week = $weekFolder.Name; SourceHash = $hash; SourceTime = $summaryFile.LastWriteTimeUtc
            DirectoryState = (Get-DirectoryState $weekFolder.FullName)
        }
    }
}

# A Usage run is redundant only when every file matches its same-week copy.
$usage = Join-Path $dataAll 'M365\Usage'
if (Test-Path -LiteralPath $usage -PathType Container) {
    foreach ($run in @(Get-ChildItem -LiteralPath $usage -Directory -Force |
        Where-Object { $_.Name -match '^\d{8}_\d{6}$' })) {
        if ($run.Attributes -band [IO.FileAttributes]::ReparsePoint) { continue }
        $runDate = [datetime]::MinValue
        if (-not [datetime]::TryParseExact($run.Name, 'yyyyMMdd_HHmmss',
            [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$runDate)) { continue }
        if ($runDate -ge (Get-Date).AddDays(-7)) { continue }
        $children = @(Get-ChildItem -LiteralPath $run.FullName -Force)
        if (@($children | Where-Object { $_.PSIsContainer -or ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) }).Count) { continue }
        $week = Get-IsoWeek $runDate
        $weekRoot = Join-Path (Join-Path $usage 'WeeklyHistory') $week
        $matchedFiles = [Collections.Generic.List[object]]::new()
        $valid = $true
        foreach ($file in $children) {
            if ($file.Name -notmatch '^(?<Base>.+)_\d{8}[-_]\d{6}\.csv$') { $valid = $false; break }
            $weekly = Join-Path $weekRoot ($Matches.Base + '.csv')
            if (-not (Test-Path -LiteralPath $weekly -PathType Leaf) -or
                $file.Length -ne (Get-Item -LiteralPath $weekly).Length -or
                (Get-Hash $file.FullName) -ne (Get-Hash $weekly)) { $valid = $false; break }
            $matchedFiles.Add([pscustomobject]@{ File = $file.FullName; Weekly = $weekly; Hash = (Get-Hash $file.FullName) })
        }
        if ($valid) {
            $runBytes = if ($children.Count) { [int64](($children | Measure-Object Length -Sum).Sum) } else { [int64]0 }
            Add-Action 'DuplicateUsageRun' $run.FullName $runBytes 'Every run file is SHA256-identical to its same-week history; or run folder is empty.' @{
                Files = $matchedFiles.ToArray(); Count = $children.Count
            }
        }
    }
}

# The former CMDB lock is eligible only when the effective, not template, job
# manifest contains no legacy preparation reference and no process owns it.
$cmdbLock = Join-Path $root '.cmdb-preparation.lock'
if ([string]::IsNullOrWhiteSpace($EffectiveJobsPath)) {
    $EffectiveJobsPath = Join-Path $PSScriptRoot '..\Orchestrator\Orchestrator-Jobs.json'
}
if ((Test-Path -LiteralPath $cmdbLock -PathType Leaf) -and
    (Test-Path -LiteralPath $EffectiveJobsPath -PathType Leaf)) {
    $jobsText = Get-Content -LiteralPath $EffectiveJobsPath -Raw
    $null = $jobsText | ConvertFrom-Json
    if ($jobsText -notmatch 'CmdbEvidence-Prepare') {
        try {
            $handle = [IO.File]::Open($cmdbLock, 'Open', 'ReadWrite', 'None')
            $handle.Dispose()
            $file = Get-Item -LiteralPath $cmdbLock -Force
            if (-not ($file.Attributes -band [IO.FileAttributes]::ReparsePoint)) {
                Add-Action 'RetiredCmdbLock' $cmdbLock $file.Length 'Effective jobs contain no legacy CMDB preparation and lock is not held.' @{
                    JobsHash = (Get-Hash $EffectiveJobsPath); LockHash = (Get-Hash $cmdbLock)
                }
            }
        } catch {
            Write-Warning ('CMDB lock retained: ' + $_.Exception.Message)
        }
    }
}

# Follow only the published prepared-batch chain. Unreferenced folders remain
# review items because they can be failed runs or deliberate recovery snapshots.
$prepared = Join-Path $root 'DATA-POWERBI'
$pointerPath = Join-Path $prepared 'current.json.txt'
$pointerHash = ''
$keptLinkedBatches = [Collections.Generic.List[object]]::new()
$linkedBatchPosition = 0
if (Test-Path -LiteralPath $pointerPath -PathType Leaf) {
    $pointerHash = Get-Hash $pointerPath
    $pointer = Get-Content -LiteralPath $pointerPath -Raw | ConvertFrom-Json
    $batchPattern = '^\d{8}T\d{9}Z-[a-f0-9]{8}$'
    $currentId = [string]$pointer.BatchId
    $previousId = [string]$pointer.PreviousBatchId
    if ($currentId -match $batchPattern -and $previousId -match $batchPattern) {
        $previousFolder = Join-Path (Join-Path $prepared 'batches') $previousId
        $previousReceiptPath = Join-Path $previousFolder 'current.json.txt'
        if (Test-Path -LiteralPath $previousReceiptPath -PathType Leaf) {
            $previousReceipt = Get-Content -LiteralPath $previousReceiptPath -Raw | ConvertFrom-Json
            if ($previousReceipt.BatchId -eq $previousId -and $previousReceipt.TenantKey -eq $pointer.TenantKey) {
                $candidateId = [string]$previousReceipt.PreviousBatchId
                $seen = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
                while ($candidateId -match $batchPattern -and $candidateId -ne $currentId -and
                    $candidateId -ne $previousId -and $seen.Add($candidateId)) {
                    $folder = Join-Path (Join-Path $prepared 'batches') $candidateId
                    if (-not (Test-Path -LiteralPath $folder -PathType Container)) { break }
                    $receiptPath = Join-Path $folder 'current.json.txt'
                    $manifestPath = Join-Path $folder 'batch.json.txt'
                    $validationPath = Join-Path $folder 'validation.json.txt'
                    if (-not ((Test-Path $receiptPath -PathType Leaf) -and
                        (Test-Path $manifestPath -PathType Leaf) -and
                        (Test-Path $validationPath -PathType Leaf))) { break }
                    $receipt = Get-Content -LiteralPath $receiptPath -Raw | ConvertFrom-Json
                    $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
                    if ($receipt.BatchId -ne $candidateId -or $manifest.BatchId -ne $candidateId -or
                        $receipt.TenantKey -ne $pointer.TenantKey -or $manifest.TenantKey -ne $pointer.TenantKey -or
                        (Get-Hash $manifestPath) -ne $receipt.ManifestSHA256) { break }
                    $files = @(Get-ChildItem -LiteralPath $folder -File -Recurse -Force)
                    $dirs = @(Get-ChildItem -LiteralPath $folder -Directory -Recurse -Force)
                    if (@($files | Where-Object { $_.Attributes -band [IO.FileAttributes]::ReparsePoint }).Count -or
                        @($dirs | Where-Object { $_.Attributes -band [IO.FileAttributes]::ReparsePoint }).Count) { break }
                    $linkedBatchPosition++
                    $batchBytes = [int64](($files | Measure-Object Length -Sum).Sum)
                    $batchTime = [datetime]::MinValue
                    $hasBatchTime = [datetime]::TryParseExact($candidateId.Substring(0, 19), 'yyyyMMddTHHmmssfffZ',
                        [Globalization.CultureInfo]::InvariantCulture,
                        [Globalization.DateTimeStyles]::AdjustToUniversal -bor [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$batchTime)
                    $keepReason = ''
                    if (-not $RetireLinkedPreparedBatches) { $keepReason = 'share retains prepared-batch history (retirement not enabled)' }
                    elseif ($linkedBatchPosition -le $RetainedLinkedPreparedBatchCount) { $keepReason = "within the $RetainedLinkedPreparedBatchCount retained older batches" }
                    elseif ($LinkedPreparedBatchMinimumAgeDays -gt 0 -and
                        (-not $hasBatchTime -or ([datetime]::UtcNow - $batchTime).TotalDays -lt $LinkedPreparedBatchMinimumAgeDays)) {
                        $keepReason = "younger than $LinkedPreparedBatchMinimumAgeDays days"
                    }
                    if ($keepReason) {
                        $keptLinkedBatches.Add([pscustomobject]@{
                            Path = (Assert-OwnedPath $folder); Bytes = $batchBytes
                            Reason = 'Linked prepared batch kept: ' + $keepReason + '.'
                        })
                        $candidateId = [string]$receipt.PreviousBatchId
                        continue
                    }
                    Add-Action 'LinkedPreparedBatch' $folder $batchBytes 'Older linked batch; current and previous stay; three audit receipts will be preserved.' @{
                        Id = $candidateId; ReceiptHash = (Get-Hash $receiptPath)
                        ManifestHash = (Get-Hash $manifestPath); ValidationHash = (Get-Hash $validationPath)
                        DirectoryState = (Get-DirectoryState $folder)
                    }
                    $candidateId = [string]$receipt.PreviousBatchId
                }
            }
        }
    }
}
$keptLinkedBytes = [int64]0
foreach ($kept in $keptLinkedBatches) { $keptLinkedBytes += $kept.Bytes }
Write-CleanupLog ('Linked prepared batches beyond current and previous: {0} found, {1} planned for retirement, {2} kept ({3:N3} GiB); retirement {4}.' -f
    $linkedBatchPosition, @($actions | Where-Object Category -eq 'LinkedPreparedBatch').Count, $keptLinkedBatches.Count,
    ($keptLinkedBytes / 1GB),
    $(if ($RetireLinkedPreparedBatches) { "enabled (retain $RetainedLinkedPreparedBatchCount, minimum age $LinkedPreparedBatchMinimumAgeDays days)" } else { 'not enabled' }))

# Inventory every file in the tenant root, including items outside the known
# cleanup categories. Paths, not file age alone, drive deletion decisions.
$allFiles = @(Get-ChildItem -LiteralPath $root -File -Recurse -Force)
$timestampedReport = Join-Path $reportRoot 'timestamped-csv.csv'
$timestampedDecisions = @(Get-TimestampedCsvDecisions)
$timestampedDecisions | Select-Object Action,Reason,Result,Folder,Name,Week,SizeBytes |
    Export-Csv -LiteralPath $timestampedReport -NoTypeInformation -Encoding utf8
$timestampedCandidates = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($decision in @($timestampedDecisions | Where-Object Action -eq 'Remove')) {
    [void]$timestampedCandidates.Add($decision.FullName)
}
$orphanReport = Join-Path $reportRoot 'orphan-temporaries.csv'
$orphanDecisions = @(Get-OrphanTemporaryDecision)
$orphanDecisions | Select-Object Action,Reason,Result,Pattern,Kind,Folder,Name,SizeBytes,AgeDays,MinimumAgeDays |
    Export-Csv -LiteralPath $orphanReport -NoTypeInformation -Encoding utf8
$orphanByPath = [Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($decision in $orphanDecisions) { $orphanByPath[$decision.FullName] = $decision }
Write-CleanupLog ('Orphan temporaries: {0} matched, {1} eligible (minimum age {2} days, {3} days for .previous).' -f
    $orphanDecisions.Count, @($orphanDecisions | Where-Object Action -eq 'Remove').Count,
    $OrphanTemporaryMinimumAgeDays, $OrphanPreviousMinimumAgeDays)
$plannedFolders = @($actions | Where-Object { $_.Category -in 'DiscoveredAppsWeek','DuplicateUsageRun','LinkedPreparedBatch','EmptyReadinessArchive' })
$plannedFileReasons = [Collections.Generic.Dictionary[string,string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($item in @($actions | Where-Object { $_.Category -in 'ReadinessArchive','RetiredCmdbLock' })) {
    $plannedFileReasons[$item.Path] = $item.Reason
}
$inventory = foreach ($file in $allFiles) {
    $relative = [IO.Path]::GetRelativePath($root, $file.FullName)
    $action = 'Review'
    $reason = 'No deletion rule proved this file redundant.'
    if ($plannedFileReasons.ContainsKey($file.FullName)) { $action = 'PlannedRemoval'; $reason = $plannedFileReasons[$file.FullName] }
    elseif ($timestampedCandidates.Contains($file.FullName)) { $action = 'PlannedRemoval'; $reason = 'Timestamped CSV qualified by weekly manifest and current DATA-LAST copy.' }
    elseif ($orphanByPath.ContainsKey($file.FullName)) {
        $orphan = $orphanByPath[$file.FullName]
        if ($orphan.Action -eq 'Remove') { $action = 'PlannedRemoval'; $reason = 'Orphan atomic-write temporary past its minimum age; no transport lock held.' }
        else { $action = 'Keep'; $reason = 'Orphan atomic-write temporary kept: ' + $orphan.Reason }
    }
    foreach ($folder in $plannedFolders) {
        if ($file.FullName.StartsWith(($folder.Path + [IO.Path]::DirectorySeparatorChar), [StringComparison]::OrdinalIgnoreCase)) {
            $action = 'PlannedRemoval'; $reason = $folder.Reason; break
        }
    }
    if ($action -eq 'Review') {
        foreach ($kept in $keptLinkedBatches) {
            if ($file.FullName.StartsWith(($kept.Path + [IO.Path]::DirectorySeparatorChar), [StringComparison]::OrdinalIgnoreCase)) {
                $action = 'Keep'; $reason = $kept.Reason; break
            }
        }
    }
    if ($action -eq 'Review') {
        if ($relative -match '^DATA-LAST[\\/]\.staging[\\/]') { $action = 'Keep'; $reason = 'Source staging of a running or interrupted producer run; its producer removes it.' }
        elseif ($relative -match '^DATA-LAST[\\/]') { $action = 'Keep'; $reason = 'Canonical latest export.' }
        elseif ($relative -match '^DATA-POWERBI-CMDB[\\/]') { $action = 'Keep'; $reason = 'Legacy CMDB rollback cohort pending acceptance.' }
        elseif ($relative -match '^DATA-POWERBI[\\/]batches[\\/]') { $reason = 'Prepared batch outside verified retirement chain.' }
        elseif ($relative -match '^DATA-REPAIR-BACKUPS[\\/]') { $reason = 'Recovery evidence requires audit review.' }
        elseif ($relative -match '^DATA-ALL[\\/]M365[\\/]Devices[\\/]EntraDevices[\\/].*RemoveCandidates.*\.ps1$') { $reason = 'Generated manual deletion script; operator-use review required.' }
        elseif ($relative -match '^DATA-ALL[\\/]M365[\\/]WorkplaceScope[\\/]') { $reason = 'WorkplaceScope run output has no general weekly-history guarantee.' }
        elseif ($relative -match '^LOG-ALL[\\/]') { $reason = 'Operational log; retention policy not established.' }
        elseif ($relative -match '^mappings[\\/]') { $reason = 'Mapping acquisition state; review before cleanup.' }
        elseif ($relative -match '^\.prepared-source\.lock$') { $action = 'Keep'; $reason = 'Active prepared-source lock path.' }
        elseif ($relative -match '^\.cmdb-preparation\.lock$') { $reason = 'Retired CMDB lock; effective job and lock state need checking.' }
        elseif ($relative -match '^SmartWorkplaceIntelligence-.*Classification\.xlsx$') { $action = 'Keep'; $reason = 'Persistent classification workbook.' }
        elseif ($relative -match '[\\/]WeeklyHistory[\\/]') { $action = 'Keep'; $reason = 'Historical consumer source; only explicit retention can remove it.' }
    }
    [pscustomobject]@{
        RelativePath = $relative; Bytes = $file.Length
        LastWriteTimeUtc = $file.LastWriteTimeUtc.ToString('o')
        Decision = $action; Reason = $reason
    }
}
$inventoryPath = Join-Path $reportRoot 'all-files.csv'
$actionsPath = Join-Path $reportRoot 'verified-actions.csv'
$inventory | Export-Csv -LiteralPath $inventoryPath -NoTypeInformation -Encoding utf8
$inventory | ForEach-Object {
    $parts = $_.RelativePath -split '[\\/]'
    [pscustomobject]@{
        Area = ($parts[0..([Math]::Min(2, $parts.Length - 1))] -join '\')
        Decision = $_.Decision; Reason = $_.Reason; Bytes = $_.Bytes
    }
} | Group-Object Area,Decision,Reason | ForEach-Object {
    [pscustomobject]@{
        Area = $_.Group[0].Area; Decision = $_.Group[0].Decision
        Reason = $_.Group[0].Reason; Files = $_.Count
        GiB = [Math]::Round((($_.Group | Measure-Object Bytes -Sum).Sum / 1GB), 3)
    }
} | Sort-Object Decision,Area,Reason |
    Export-Csv -LiteralPath (Join-Path $reportRoot 'summary.csv') -NoTypeInformation -Encoding utf8
$actions | Select-Object Category,RelativePath,Bytes,Reason,Status |
    Export-Csv -LiteralPath $actionsPath -NoTypeInformation -Encoding utf8

if ($Execute) {
    foreach ($action in $actions) {
        try {
            $path = Assert-OwnedPath $action.Path
            if (-not (Test-Path -LiteralPath $path)) { $action.Status = 'AlreadyAbsent'; continue }
            switch ($action.Category) {
                'ReadinessArchive' {
                    $file = Get-Item -LiteralPath $path -Force
                    if ($file.Length -ne $action.Evidence.Length -or
                        $file.LastWriteTimeUtc -ne $action.Evidence.LastWriteTimeUtc -or
                        -not (Test-Path $readyLatest -PathType Leaf) -or
                        -not (Test-Path $readySummary -PathType Leaf) -or
                        -not (Test-Path $readyWeekly -PathType Container) -or
                        -not @(Get-ChildItem -LiteralPath $readyWeekly -File -Recurse |
                            Where-Object { $_.Name -like 'Intune_Windows11_Readiness_Issues*.csv' -and $_.Length -gt 0 }).Count) { throw 'Readiness evidence changed.' }
                    Remove-Item -LiteralPath $path -Force -ErrorAction Stop
                }
                'RetiredCmdbLock' {
                    if ((Get-Hash $EffectiveJobsPath) -ne $action.Evidence.JobsHash -or
                        (Get-Hash $path) -ne $action.Evidence.LockHash) { throw 'CMDB lock evidence changed.' }
                    $handle = [IO.File]::Open($path, 'Open', 'ReadWrite', 'None')
                    $handle.Dispose()
                    Remove-Item -LiteralPath $path -Force -ErrorAction Stop
                }
                'EmptyReadinessArchive' {
                    if (@(Get-ChildItem -LiteralPath $path -Force).Count) { throw 'Readiness archive is not empty.' }
                    Remove-Item -LiteralPath $path -Force -ErrorAction Stop
                }
                'DiscoveredAppsWeek' {
                    $source = Join-Path $path 'Intune_DiscoveredApps_Summary.csv'
                    if ((Get-Hash $source) -ne $action.Evidence.SourceHash -or
                        (Get-Item $source).LastWriteTimeUtc -ne $action.Evidence.SourceTime -or
                        (Test-Path (Join-Path $path 'upload.pending')) -or
                        (Get-DirectoryState $path) -cne $action.Evidence.DirectoryState) { throw 'Discovered Apps source changed.' }
                    $destinationFolder = Assert-OwnedPath (Join-Path $appsSummary $action.Evidence.Week)
                    New-Item -ItemType Directory -Path $destinationFolder -Force | Out-Null
                    $destination = Assert-OwnedPath (Join-Path $destinationFolder 'Intune_DiscoveredApps_Summary.csv')
                    if (-not (Test-Path $destination)) { Copy-Item -LiteralPath $source -Destination $destination -ErrorAction Stop }
                    if ((Get-Hash $destination) -ne $action.Evidence.SourceHash) { throw 'Discovered Apps summary hash mismatch.' }
                    [IO.File]::SetLastWriteTimeUtc($destination, $action.Evidence.SourceTime)
                    $manifestPath = Assert-OwnedPath (Join-Path $destinationFolder 'summary.manifest.json.txt')
                    $expected = [ordered]@{
                        SchemaVersion = 1; Week = $action.Evidence.Week
                        FileName = 'Intune_DiscoveredApps_Summary.csv'
                        SHA256 = $action.Evidence.SourceHash
                        SourceLastWriteTimeUtc = $action.Evidence.SourceTime.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
                    }
                    if (Test-Path $manifestPath) {
                        $existing = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
                        if ($existing.Week -ne $expected.Week -or $existing.SHA256 -ne $expected.SHA256 -or
                            $existing.SourceLastWriteTimeUtc -ne $expected.SourceLastWriteTimeUtc) { throw 'Discovered Apps summary manifest mismatch.' }
                    } else {
                        $expected | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $manifestPath -Encoding utf8 -NoNewline
                    }
                    if (@(Get-ChildItem -LiteralPath $path -Recurse -Force | Where-Object { $_.Attributes -band [IO.FileAttributes]::ReparsePoint }).Count) {
                        throw 'Discovered Apps week gained a reparse point.'
                    }
                    Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop
                }
                'DuplicateUsageRun' {
                    $children = @(Get-ChildItem -LiteralPath $path -Force)
                    if ($children.Count -ne $action.Evidence.Count) { throw 'Usage run contents changed.' }
                    foreach ($match in $action.Evidence.Files) {
                        if ((Get-Hash $match.File) -ne $match.Hash -or (Get-Hash $match.Weekly) -ne $match.Hash) {
                            throw 'Usage duplicate evidence changed.'
                        }
                    }
                    Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop
                }
                'LinkedPreparedBatch' {
                    $receiptPath = Join-Path $path 'current.json.txt'
                    $manifestPath = Join-Path $path 'batch.json.txt'
                    $validationPath = Join-Path $path 'validation.json.txt'
                    if ((Get-Hash $pointerPath) -ne $pointerHash -or
                        (Get-DirectoryState $path) -cne $action.Evidence.DirectoryState -or
                        (Get-Hash $receiptPath) -ne $action.Evidence.ReceiptHash -or
                        (Get-Hash $manifestPath) -ne $action.Evidence.ManifestHash -or
                        (Get-Hash $validationPath) -ne $action.Evidence.ValidationHash) { throw 'Prepared batch receipts changed.' }
                    $audit = Assert-OwnedPath (Join-Path (Join-Path $prepared 'retired') $action.Evidence.Id)
                    New-Item -ItemType Directory -Path $audit -Force | Out-Null
                    foreach ($name in 'current','batch','validation') {
                        $source = Join-Path $path ($name + '.json.txt')
                        $target = Assert-OwnedPath (Join-Path $audit ($name + '.json.txt'))
                        if (-not (Test-Path $target)) { Copy-Item -LiteralPath $source -Destination $target -ErrorAction Stop }
                        if ((Get-Hash $source) -ne (Get-Hash $target)) { throw 'Prepared batch audit copy mismatch.' }
                    }
                    if (@(Get-ChildItem -LiteralPath $path -Recurse -Force | Where-Object { $_.Attributes -band [IO.FileAttributes]::ReparsePoint }).Count) {
                        throw 'Prepared batch gained a reparse point.'
                    }
                    Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop
                }
            }
            $action.Status = 'Removed'
        } catch {
            $action.Status = 'Failed: ' + $_.Exception.Message
        }
    }
}

$actions | Select-Object Category,RelativePath,Bytes,Reason,Status |
    Export-Csv -LiteralPath $actionsPath -NoTypeInformation -Encoding utf8
if ($Execute) {
    $rechecked = [Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($decision in @(Get-TimestampedCsvDecisions | Where-Object Action -eq 'Remove')) {
        $rechecked[$decision.FullName] = $decision
    }
    foreach ($decision in @($timestampedDecisions | Where-Object Action -eq 'Remove')) {
        try {
            if (-not $rechecked.ContainsKey($decision.FullName)) { throw 'Timestamped CSV is no longer qualified.' }
            $current = $rechecked[$decision.FullName]
            if ($current.SizeBytes -ne $decision.SizeBytes -or
                $current.LastWriteTimeUtc -ne $decision.LastWriteTimeUtc) { throw 'Timestamped CSV changed after preview.' }
            $null = Assert-OwnedPath $decision.FullName
            Remove-Item -LiteralPath $decision.FullName -Force -ErrorAction Stop
            $decision.Result = 'Removed'
        } catch {
            $decision.Result = 'Failed: ' + $_.Exception.Message
        }
    }
    $timestampedDecisions | Select-Object Action,Reason,Result,Folder,Name,Week,SizeBytes |
        Export-Csv -LiteralPath $timestampedReport -NoTypeInformation -Encoding utf8

    # Requalify from disk, then delete while every transport lock of the folder is held.
    $orphanRechecked = [Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($decision in @(Get-OrphanTemporaryDecision)) { $orphanRechecked[$decision.FullName] = $decision }
    foreach ($decision in @($orphanDecisions | Where-Object Action -eq 'Remove')) {
        $relativePath = Join-Path $decision.Folder $decision.Name
        try {
            if (-not (Test-Path -LiteralPath $decision.FullName)) { $decision.Result = 'AlreadyAbsent'; continue }
            if (-not $orphanRechecked.ContainsKey($decision.FullName)) { throw 'Orphan temporary no longer matches a known pattern.' }
            $current = $orphanRechecked[$decision.FullName]
            if ($current.Action -ne 'Remove') { $decision.Result = 'Skipped: ' + $current.Reason }
            elseif ($current.SizeBytes -ne $decision.SizeBytes -or
                $current.LastWriteTimeUtc -ne $decision.LastWriteTimeUtc) { $decision.Result = 'Skipped: changed after preview' }
            else {
                $path = Assert-OwnedPath $decision.FullName
                $handles = $null
                try { $handles = Open-FolderTransportLockSet (Split-Path -Path $path -Parent) }
                catch [IO.IOException] { $decision.Result = 'Skipped: TransportLockHeld' }
                if ($null -ne $handles) {
                    try { Remove-Item -LiteralPath $path -Force -ErrorAction Stop }
                    finally { foreach ($handle in $handles) { $handle.Dispose() } }
                    $decision.Result = 'Removed'
                    Write-CleanupLog ('Removed orphan temporary {0} ({1} bytes, {2}, {3} days old).' -f $relativePath, $decision.SizeBytes, $decision.Pattern, $decision.AgeDays)
                }
            }
            if ($decision.Result -like 'Skipped:*') { Write-CleanupLog ('Orphan temporary kept {0}: {1}' -f $relativePath, $decision.Result) }
        } catch {
            $decision.Result = 'Failed: ' + $_.Exception.Message
            Write-CleanupLog ('Orphan temporary removal failed for {0}: {1}' -f $relativePath, $_.Exception.Message)
        }
    }
    $orphanDecisions | Select-Object Action,Reason,Result,Pattern,Kind,Folder,Name,SizeBytes,AgeDays,MinimumAgeDays |
        Export-Csv -LiteralPath $orphanReport -NoTypeInformation -Encoding utf8
}

$result = [pscustomobject]@{
    DataRootPath = $root
    Mode = if ($Execute) { 'Execute' } else { 'Preview' }
    InventoryFiles = $allFiles.Count
    PlannedActions = $actions.Count + $timestampedCandidates.Count + @($orphanDecisions | Where-Object Action -eq 'Remove').Count
    RemovedActions = @($actions | Where-Object Status -eq 'Removed').Count + @($timestampedDecisions | Where-Object Result -eq 'Removed').Count +
        @($orphanDecisions | Where-Object Result -eq 'Removed').Count
    FailedActions = @($actions | Where-Object Status -like 'Failed:*').Count + @($timestampedDecisions | Where-Object Result -like 'Failed:*').Count +
        @($orphanDecisions | Where-Object Result -like 'Failed:*').Count
    SkippedOrphanTemporaries = @($orphanDecisions | Where-Object Result -like 'Skipped:*').Count
    ReportDirectory = $reportRoot
}
$result
if ($result.FailedActions -gt 0) { throw 'One or more verified cleanup actions failed; inspect verified-actions.csv, timestamped-csv.csv and orphan-temporaries.csv.' }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCCm7DWNCrfKrvM
# TSaU+NIaxmZyzhUxxUw+fAEbjCPIeqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGK3jxnWxFL6lqMhhEerei0kb+gAfJl8JQpQyHkjffYXMA0GCSqG
# SIb3DQEBAQUABIIBgEfsaemX64f17cONJRigDRjftyIj+hUyVFqJIYbmBJtW0aaa
# nDLBBLXkzlhHrLMX6tqx5fOP+UV4flcjQ9CJxQggBh7FmVWx0eQPMx3AWovg4CdZ
# uqnFMylz7LULCZUZ8zLr+TzaqIeF2xOLn8BlufcvswOnwSET0KWbGNGcQvtw0ct6
# YdziG8W0JI1r744NgilPYLZQrkTsLnzLGS+TO4Gxd3uf4fpuYWGPOY7Z4kPyqBs5
# v8MruN8KqRAddrzJNc4PXAzJtcSbKGA7/q7Kdyujjgxxo4Xs9c7vTCwCcN6WEkPh
# 8KyYD+s4dS4hFdv45Ey1yi/x0oA9UJ5R96RBlLDAGx++IS2EXhdRDCmGaP2yZY4h
# uAP4txPiJvhyZmfv6DGxv0NSsZlUFHXwZEx8CTaL6xHaWHtSN1wsacvDugFGPoFn
# EpnDK+ClLybBZ9oUZ8JOlvz6+qBW9AphHNmMbUYv3aMOZX33h0f2557UlV/zkwE2
# /kNw3VyZ8efKyIkpSqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjIxNTM4
# MzdaMC8GCSqGSIb3DQEJBDEiBCDaReEfMTpV+0t09yWVvjQfHu0SLJYdpZTOI2Ic
# QKJDmjANBgkqhkiG9w0BAQEFAASCAgCrYvCNr35wdJIJDnLPRzQiLAmjawlDpCD+
# TvVqgeJmEZaV6pqkPQzlKLuk73rG1WNIC/Xuj0vm9VEmsqoS/CJuz/hws1YnWWT7
# nxZp6ZSO2SgvvlmYF3oYvXehSC+XvmMkdqxjfsdCkZxBqFa4dqRyej5lbaOktoSR
# dZSrZ/WMyg2nTh2iNwUN91pngrjqj/4uMdQcF7Wl3iG9smo8l6GYLCr8a/g8PoaX
# cxUO3XzKgWMi5gjwBfIXC5d+yHI14668JSZ0IsKkb/3sHeeHhVNmT9Y0Chc0IuVn
# C4Lejguv+qETNu3Mt9yFCc6RZHt198vYG2dI7A0DjPvv7K90Kaeq42Ph/HYuEaBe
# gKdr7AAEd2Z0uYzgmhBs1O3gDVCD4gyRyd/6DiwVgbHmsxV7urkpJrmFcHJGjio5
# SkB5fVhs9Gyc1VxRdrEnJN/B/soCPtDDkDljI4a8N5qFepj9bLljWau/VjKLNT/k
# zI7Lv7SKSDLIS36ZP5ePzBFXvzT5EnNwUVROND/bY/SbPYw03S+wSPuboZkM5OJr
# 6E5NiXDEdG8G1x+8phfX+2gp69IPxZVVS4Hf4KPGu6E7ZrKjmxTN4fJZkUHR4Ss5
# oaxPXKmsjEwrBNIL+4Y/E7qwnJLOdT8fy0IONOODRC6fXwB9feUSsHHkZEst5A2a
# kWXQ0uFI9Q==
# SIG # End signature block
