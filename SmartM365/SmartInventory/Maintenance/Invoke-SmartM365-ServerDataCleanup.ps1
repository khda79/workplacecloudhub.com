<#
.SYNOPSIS
Audits a complete SmartM365 tenant DATA root and retires verified redundant files.

.DESCRIPTION
Preview is read-only. Retired readiness data is preserved. Execute removes
old Discovered Apps detail weeks after preserving their summaries,
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
1.2.2

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

function Get-UtcTick {
    # ConvertFrom-Json turns an ISO 'o' string into a [datetime] before 7.5 offers -DateKind;
    # both forms compare as UTC ticks. Anything unreadable returns -1 and never matches.
    param([object]$Value)
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [DateTimeKind]::Unspecified) { return [datetime]::SpecifyKind($Value, [DateTimeKind]::Utc).Ticks }
        return $Value.ToUniversalTime().Ticks
    }
    if ($Value -is [DateTimeOffset]) { return $Value.UtcTicks }
    $parsed = [DateTimeOffset]::MinValue
    if ($Value -is [string] -and [DateTimeOffset]::TryParse($Value, [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$parsed)) { return $parsed.UtcTicks }
    return [int64]-1
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
        # Retired readiness producers: preserve all existing CSV generations and history.
        if ($match.Groups['Prefix'].Value -match '^(Exchange_HybridIdentity_Issues|Intune_Windows11_Readiness_Issues)(?:_|$)') { continue }
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

# Retired readiness CSVs, archives and WeeklyHistory are preserved.

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
$plannedFolders = @($actions | Where-Object { $_.Category -in 'DiscoveredAppsWeek','DuplicateUsageRun','LinkedPreparedBatch' })
$plannedFileReasons = [Collections.Generic.Dictionary[string,string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($item in @($actions | Where-Object { $_.Category -in 'RetiredCmdbLock' })) {
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
                            -not $existing.PSObject.Properties['SourceLastWriteTimeUtc'] -or
                            (Get-UtcTick $existing.SourceLastWriteTimeUtc) -ne $action.Evidence.SourceTime.ToUniversalTime().Ticks) { throw 'Discovered Apps summary manifest mismatch.' }
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
