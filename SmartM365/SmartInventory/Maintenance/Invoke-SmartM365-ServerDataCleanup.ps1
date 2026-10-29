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

Daily capture roots (folders with DailyCaptureRecords) keep their own seven-day
retention, which only removes the CSVs it recorded. Timestamped CSVs at the top
of such a root that no record owns (written before adoption or by an unqualified
run) are retired once older than the seven-day window, and only when a recorded
capture inside the window still holds the same export. CSVs of run folders
(yyyyMMdd_HHmmss, such as the Usage and Copilot runs) that no record owns, raw
extracts included, are retired the same way once the producer has a recorded
capture inside the window; a run folder left empty is removed. Every capture
retention lock of the root is held while they are removed. A file another rule
already removes is reported as PlannedByAnotherRule and counted once.

.VERSION
1.5.0

.EXAMPLE
./Invoke-SmartM365-ServerDataCleanup.ps1 -Tenant prod -DataRootPath '<server DATA root>'
./Invoke-SmartM365-ServerDataCleanup.ps1 -Tenant prod -DataRootPath '<server DATA root>' -Execute
./Invoke-SmartM365-ServerDataCleanup.ps1 -Tenant prod -DataRootPath '<server DATA root>' -RetireLinkedPreparedBatches -RetainedLinkedPreparedBatchCount 10 -LinkedPreparedBatchMinimumAgeDays 90
./Invoke-SmartM365-ServerDataCleanup.ps1 -Tenant prod -Execute

.NOTES
Without -DataRootPath the tenant DATA root is the parent of the tenant DataAllRootPath of the
global configuration: this is how the orchestrator job Maintenance-ServerDataCleanup runs it.
Without -ReportDirectory the reports go to a dated TEMP folder; such default folders older than
-ReportRetentionDays are removed at the start of the next run.
#>
#requires -Version 7.2
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$DataRootPath = '',
    [string]$ReportDirectory = '',
    [ValidateRange(1, 3650)][int]$ReportRetentionDays = 30,
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
    if ([string]::IsNullOrWhiteSpace($DataRootPath)) {
        # Same tenant DATA-ALL as the producers: its parent holds DATA-ALL, DATA-LAST and LOG-ALL.
        $tenantConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath (Join-Path $PSScriptRoot '..\..')
        $tokens = @{
            ProfileKey = [string]$tenantConfig.ProfileKey
            WorkspaceRootPath = [string]$tenantConfig.WorkspaceRootPath
            SmartM365RootPath = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..')).TrimEnd('\')
        }
        $dataAllSetting = [string]$tenantConfig.DataAllRootPath
        for ($pass = 0; $pass -lt 5 -and $dataAllSetting -match '\{\{'; $pass++) {
            foreach ($match in [regex]::Matches($dataAllSetting, '\{\{([A-Za-z0-9_.-]+)\}\}')) {
                if (-not $tokens.ContainsKey($match.Groups[1].Value)) { throw "DataAllRootPath token $($match.Value) cannot be resolved; pass -DataRootPath." }
                $dataAllSetting = $dataAllSetting.Replace($match.Value, $tokens[$match.Groups[1].Value])
            }
        }
        if ($dataAllSetting -match '\{\{' -or -not [IO.Path]::IsPathRooted($dataAllSetting)) { throw "DataAllRootPath does not resolve to an absolute folder: $dataAllSetting" }
        $DataRootPath = Split-Path -Path ([IO.Path]::GetFullPath($dataAllSetting).TrimEnd('\', '/')) -Parent
        Write-Information ('[{0}] DATA root resolved from the tenant configuration: {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $DataRootPath) -InformationAction Continue
    }
    else {
        Write-SmartM365StartupBanner
    }
}
if ([string]::IsNullOrWhiteSpace($DataRootPath)) { throw 'DataRootPath is required when the SmartM365 configuration is not available.' }

$root = [IO.Path]::GetFullPath($DataRootPath).TrimEnd('\', '/')
if (-not (Test-Path -LiteralPath $root -PathType Container)) { throw "DATA root not found: $root" }
if (-not [string]::Equals((Split-Path $root -Leaf), $Tenant, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'The DATA root leaf must match -Tenant.'
}
if ((Get-Item -LiteralPath $root -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) {
    throw 'A reparse-point DATA root is not supported.'
}
if ([string]::IsNullOrWhiteSpace($ReportDirectory)) {
    # A scheduled run reports to TEMP: earlier default report folders age out here.
    $reportParent = [IO.Path]::GetTempPath()
    $reportCutoff = (Get-Date).AddDays(-$ReportRetentionDays)
    foreach ($oldReport in @(Get-ChildItem -LiteralPath $reportParent -Directory -Filter 'SmartM365-ServerDataCleanup_*' -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^SmartM365-ServerDataCleanup_\d{8}_\d{6}$' -and $_.LastWriteTime -lt $reportCutoff -and
            -not ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) })) {
        try { Remove-Item -LiteralPath $oldReport.FullName -Recurse -Force -ErrorAction Stop }
        catch { Write-Warning ('Old cleanup report kept: {0}: {1}' -f $oldReport.FullName, $_.Exception.Message) }
    }
    $ReportDirectory = Join-Path $reportParent ('SmartM365-ServerDataCleanup_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
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

$script:captureRootCache = @{}
function Test-InDailyCaptureRoot {
    # A daily capture root (a folder with DailyCaptureRecords) is cleaned only by the capture rules:
    # deleting a recorded payload there would make its own retention refuse every later purge.
    param([Parameter(Mandatory)][string]$Folder)
    $cursor = [IO.Path]::GetFullPath($Folder).TrimEnd('\', '/')
    $visited = [Collections.Generic.List[string]]::new()
    $result = $false
    while ($cursor -and $cursor.StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)) {
        if ($script:captureRootCache.ContainsKey($cursor)) { $result = $script:captureRootCache[$cursor]; break }
        $visited.Add($cursor)
        if (Test-Path -LiteralPath (Join-Path $cursor 'DailyCaptureRecords') -PathType Container) { $result = $true; break }
        $cursor = [IO.Path]::GetDirectoryName($cursor)
    }
    foreach ($folder in $visited) { $script:captureRootCache[$folder] = $result }
    return $result
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
        if (-not $reason -and (Test-InDailyCaptureRoot $file.DirectoryName)) { $reason = 'DailyCaptureRoot' }
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

function Get-ParisDay {
    # Same calendar as SmartM365-DailyCaptureRetention: Europe/Paris days.
    param([datetime]$Utc)
    try { $zone = [TimeZoneInfo]::FindSystemTimeZoneById('Europe/Paris') }
    catch { $zone = [TimeZoneInfo]::FindSystemTimeZoneById('Romance Standard Time') }
    return [TimeZoneInfo]::ConvertTimeFromUtc([datetime]::SpecifyKind($Utc, [DateTimeKind]::Utc), $zone).ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
}

function Open-CaptureRetentionLockSet {
    # DailyCaptureRetention holds DailyCaptureRecords\<producer>\.retention.lock exclusively while it
    # retires captures; holding every lock of the root keeps it out while legacy files are removed.
    param([Parameter(Mandatory)][string]$CaptureRoot)
    $handles = [Collections.Generic.List[IO.FileStream]]::new()
    try {
        foreach ($lock in @(Get-ChildItem -LiteralPath (Join-Path $CaptureRoot 'DailyCaptureRecords') -File -Force -Recurse -Filter '.retention.lock')) {
            if ($lock.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Linked capture retention lock: $($lock.FullName)" }
            $handles.Add([IO.File]::Open($lock.FullName, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None))
        }
    } catch {
        foreach ($handle in $handles) { $handle.Dispose() }
        throw
    }
    return ,$handles
}

function Get-DailyCaptureLegacyDecision {
    # Window of DailyCaptureRetention: today plus the six preceding Europe/Paris days.
    $today = Get-ParisDay ([datetime]::UtcNow)
    $cutoff = [datetime]::ParseExact($today, 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture).AddDays(-6).ToString('yyyy-MM-dd')
    foreach ($recordsFolder in @(Get-ChildItem -LiteralPath $dataAll -Directory -Recurse -Force -Filter 'DailyCaptureRecords')) {
        $captureRoot = $recordsFolder.Parent.FullName
        $rootState = ''
        $owned = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        # Export names with a recorded, not retired capture inside the window whose payload is present.
        $recentExports = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        $recentCaptureCount = 0
        try {
            $null = Assert-OwnedPath $recordsFolder.FullName
            foreach ($meta in @(Get-ChildItem -LiteralPath $recordsFolder.FullName -File -Recurse -Force -Filter '*.json.txt')) {
                if ($meta.Attributes -band [IO.FileAttributes]::ReparsePoint) { throw "Linked capture record: $($meta.Name)" }
                $record = Get-Content -LiteralPath $meta.FullName -Raw -Encoding utf8 | ConvertFrom-Json
                if ([string]$record.Owner -cne 'SmartM365-DailyCapture' -or [int]$record.Version -ne 1 -or [string]$record.Status -cne 'Completed' -or
                    [string]$record.Day -notmatch '^\d{4}-\d{2}-\d{2}$' -or @($record.Files).Count -eq 0) { throw "Unrecognised capture record: $($meta.Name)" }
                $removed = @($record.RemovedFiles | ForEach-Object { [string]$_ })
                foreach ($file in @($record.Files)) {
                    $relative = [string]$file.Relative
                    if ([string]::IsNullOrWhiteSpace($relative) -or [string]::IsNullOrWhiteSpace([string]$file.File)) { throw "Capture record without file identity: $($meta.Name)" }
                    $full = [IO.Path]::GetFullPath((Join-Path $captureRoot $relative))
                    [void]$owned.Add($full)
                    if ([string]$record.Day -lt $cutoff -or $relative -in $removed) { continue }
                    $payload = Get-Item -LiteralPath $full -Force -ErrorAction SilentlyContinue
                    if ($null -ne $payload -and $payload.Length -eq [int64]$file.Bytes) { [void]$recentExports.Add([string]$file.File); $recentCaptureCount++ }
                }
            }
        } catch { $rootState = 'UnreadableCaptureRecords' }
        # Top-level timestamped CSVs, then CSVs of run folders (yyyyMMdd_HHmmss) such as the Usage and
        # Copilot runs: their raw extracts and pre-adoption exports are never recorded by a capture.
        $candidates = [Collections.Generic.List[object]]::new()
        foreach ($file in @(Get-ChildItem -LiteralPath $captureRoot -File -Force -Filter '*.csv')) {
            $match = [regex]::Match($file.Name, '^(?<Base>.+)_(?<Stamp>\d{8})[-_]\d{6}\.csv$', [Text.RegularExpressions.RegexOptions]::IgnoreCase)
            if ($match.Success) { $candidates.Add([pscustomobject]@{ File = $file; Stamp = $match.Groups['Stamp'].Value; Export = $match.Groups['Base'].Value + '.csv'; RunFolder = '' }) }
        }
        foreach ($run in @(Get-ChildItem -LiteralPath $captureRoot -Directory -Force | Where-Object { $_.Name -match '^\d{8}_\d{6}$' })) {
            if ($run.Attributes -band [IO.FileAttributes]::ReparsePoint) { continue }
            foreach ($file in @(Get-ChildItem -LiteralPath $run.FullName -File -Force -Filter '*.csv')) {
                $candidates.Add([pscustomobject]@{ File = $file; Stamp = $run.Name.Substring(0, 8); Export = ''; RunFolder = $run.FullName })
            }
        }
        foreach ($candidate in $candidates) {
            $file = $candidate.File
            $stampDay = ''
            $stampDate = [datetime]::MinValue
            if ([datetime]::TryParseExact($candidate.Stamp, 'yyyyMMdd', [Globalization.CultureInfo]::InvariantCulture,
                [Globalization.DateTimeStyles]::None, [ref]$stampDate)) { $stampDay = $stampDate.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture) }
            $reason = $rootState
            if (-not $reason -and $owned.Contains($file.FullName)) { $reason = 'CaptureOwned' }
            if (-not $reason -and $file.Name -match '(?i)receipt|manifest') { $reason = 'ReceiptOrManifest' }
            if (-not $reason -and -not $stampDay) { $reason = 'InvalidTimestamp' }
            if (-not $reason -and $stampDay -ge $cutoff) { $reason = 'WithinCaptureWindow' }
            if (-not $reason -and (Get-ParisDay $file.LastWriteTimeUtc) -ge $cutoff) { $reason = 'RecentWrite' }
            # A top-level copy needs a newer capture of the same export; a run file needs a newer capture of the producer.
            if (-not $reason -and $candidate.Export -and -not $recentExports.Contains($candidate.Export)) { $reason = 'NoRecentRecordedCapture' }
            if (-not $reason -and -not $candidate.Export -and $recentCaptureCount -eq 0) { $reason = 'NoRecentRecordedCapture' }
            if (-not $reason) {
                try {
                    $null = Assert-OwnedPath $file.FullName
                    if ($file.Attributes -band [IO.FileAttributes]::ReparsePoint) { $reason = 'ReparseOrOutsideRoot' }
                } catch { $reason = 'ReparseOrOutsideRoot' }
            }
            [pscustomobject]@{
                Action = if ($reason) { 'Keep' } else { 'Remove' }
                Reason = $reason; Result = ''
                CaptureRoot = [IO.Path]::GetRelativePath($root, $captureRoot); CaptureRootPath = $captureRoot
                Name = [IO.Path]::GetRelativePath($captureRoot, $file.FullName); StampDay = $stampDay; WindowStartDay = $cutoff; SizeBytes = $file.Length
                FullName = $file.FullName; LastWriteTimeUtc = $file.LastWriteTimeUtc; RunFolderPath = $candidate.RunFolder
            }
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
# Once Usage keeps daily captures, its run folders hold recorded payloads: the capture rules own them.
if ((Test-Path -LiteralPath $usage -PathType Container) -and (Test-InDailyCaptureRoot $usage)) {
    Write-CleanupLog 'Usage run folders: daily capture root; duplicate-run retirement left to the capture rules.'
}
elseif (Test-Path -LiteralPath $usage -PathType Container) {
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
$legacyCaptureReport = Join-Path $reportRoot 'daily-capture-legacy.csv'
$legacyCaptureDecisions = @(Get-DailyCaptureLegacyDecision)
# A file another rule already removes is reported, not planned twice.
$coveringFolders = @($actions | Where-Object { $_.Category -in 'DiscoveredAppsWeek','DuplicateUsageRun','LinkedPreparedBatch' } | ForEach-Object { $_.Path + [IO.Path]::DirectorySeparatorChar })
foreach ($decision in @($legacyCaptureDecisions | Where-Object Action -eq 'Remove')) {
    if ($timestampedCandidates.Contains($decision.FullName) -or
        @($coveringFolders | Where-Object { $decision.FullName.StartsWith($_, [StringComparison]::OrdinalIgnoreCase) }).Count) {
        $decision.Action = 'Keep'; $decision.Reason = 'PlannedByAnotherRule'
    }
}
$legacyCaptureDecisions | Select-Object Action,Reason,Result,CaptureRoot,Name,StampDay,WindowStartDay,SizeBytes |
    Export-Csv -LiteralPath $legacyCaptureReport -NoTypeInformation -Encoding utf8
$legacyCaptureByPath = [Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($decision in $legacyCaptureDecisions) { $legacyCaptureByPath[$decision.FullName] = $decision }
$legacyCaptureEligible = @($legacyCaptureDecisions | Where-Object Action -eq 'Remove')
$legacyCaptureBytes = [int64]0
foreach ($decision in $legacyCaptureEligible) { $legacyCaptureBytes += $decision.SizeBytes }
Write-CleanupLog ('Daily capture legacy CSVs: {0} recorded and {1} unrecorded timestamped or run-folder file(s) in {2} capture root(s); {3} eligible ({4:N3} GiB), {5} left to another rule.' -f
    @($legacyCaptureDecisions | Where-Object Reason -eq 'CaptureOwned').Count, @($legacyCaptureDecisions | Where-Object Reason -ne 'CaptureOwned').Count,
    @($legacyCaptureDecisions | Select-Object -ExpandProperty CaptureRoot -Unique).Count,
    $legacyCaptureEligible.Count, ($legacyCaptureBytes / 1GB), @($legacyCaptureDecisions | Where-Object Reason -eq 'PlannedByAnotherRule').Count)
$script:removedRunFolders = 0
$script:failedRunFolders = 0
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
    elseif ($legacyCaptureByPath.ContainsKey($file.FullName)) {
        $legacy = $legacyCaptureByPath[$file.FullName]
        if ($legacy.Action -eq 'Remove') { $action = 'PlannedRemoval'; $reason = 'Unrecorded daily capture CSV older than the seven-day window; a recorded capture holds the same export.' }
        elseif ($legacy.Reason -eq 'CaptureOwned') { $action = 'Keep'; $reason = 'Recorded daily capture; its own retention removes it.' }
        else { $action = 'Keep'; $reason = 'Unrecorded daily capture CSV kept: ' + $legacy.Reason }
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

    # One capture root at a time: hold its retention locks, requalify from disk, then delete.
    foreach ($group in @($legacyCaptureEligible | Group-Object CaptureRootPath)) {
        $handles = $null
        try { $handles = Open-CaptureRetentionLockSet $group.Name }
        catch [IO.IOException] { foreach ($decision in $group.Group) { $decision.Result = 'Skipped: CaptureRetentionLockHeld' } }
        catch { foreach ($decision in $group.Group) { $decision.Result = 'Failed: ' + $_.Exception.Message } }
        if ($null -eq $handles) { continue }
        try {
            $legacyRechecked = [Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($decision in @(Get-DailyCaptureLegacyDecision | Where-Object CaptureRootPath -eq $group.Name)) { $legacyRechecked[$decision.FullName] = $decision }
            foreach ($decision in $group.Group) {
                $relativePath = Join-Path $decision.CaptureRoot $decision.Name
                try {
                    if (-not (Test-Path -LiteralPath $decision.FullName)) { $decision.Result = 'AlreadyAbsent'; continue }
                    if (-not $legacyRechecked.ContainsKey($decision.FullName)) { throw 'Daily capture legacy CSV is no longer listed.' }
                    $current = $legacyRechecked[$decision.FullName]
                    if ($current.Action -ne 'Remove') { $decision.Result = 'Skipped: ' + $current.Reason }
                    elseif ($current.SizeBytes -ne $decision.SizeBytes -or
                        $current.LastWriteTimeUtc -ne $decision.LastWriteTimeUtc) { $decision.Result = 'Skipped: changed after preview' }
                    else {
                        Remove-Item -LiteralPath (Assert-OwnedPath $decision.FullName) -Force -ErrorAction Stop
                        $decision.Result = 'Removed'
                        Write-CleanupLog ('Removed daily capture legacy CSV {0} ({1} bytes, day {2}).' -f $relativePath, $decision.SizeBytes, $decision.StampDay)
                    }
                    if ($decision.Result -like 'Skipped:*') { Write-CleanupLog ('Daily capture legacy CSV kept {0}: {1}' -f $relativePath, $decision.Result) }
                } catch {
                    $decision.Result = 'Failed: ' + $_.Exception.Message
                    Write-CleanupLog ('Daily capture legacy CSV removal failed for {0}: {1}' -f $relativePath, $_.Exception.Message)
                }
            }
            # A run folder left without any entry is removed; anything still inside keeps it.
            foreach ($runFolder in @($group.Group | Where-Object { $_.RunFolderPath -and $_.Result -in 'Removed','AlreadyAbsent' } | Select-Object -ExpandProperty RunFolderPath -Unique)) {
                try {
                    if (-not (Test-Path -LiteralPath $runFolder -PathType Container)) { continue }
                    $folderItem = Get-Item -LiteralPath (Assert-OwnedPath $runFolder) -Force
                    if ($folderItem.Attributes -band [IO.FileAttributes]::ReparsePoint) { continue }
                    if (@(Get-ChildItem -LiteralPath $runFolder -Force).Count) { continue }
                    Remove-Item -LiteralPath $runFolder -Force -ErrorAction Stop
                    $script:removedRunFolders++
                    Write-CleanupLog ('Removed empty daily capture run folder {0}.' -f [IO.Path]::GetRelativePath($root, $runFolder))
                } catch {
                    $script:failedRunFolders++
                    Write-CleanupLog ('Daily capture run folder removal failed for {0}: {1}' -f $runFolder, $_.Exception.Message)
                }
            }
        } finally { foreach ($handle in $handles) { $handle.Dispose() } }
    }
    $legacyCaptureDecisions | Select-Object Action,Reason,Result,CaptureRoot,Name,StampDay,WindowStartDay,SizeBytes |
        Export-Csv -LiteralPath $legacyCaptureReport -NoTypeInformation -Encoding utf8
}

$result = [pscustomobject]@{
    DataRootPath = $root
    Mode = if ($Execute) { 'Execute' } else { 'Preview' }
    InventoryFiles = $allFiles.Count
    PlannedActions = $actions.Count + $timestampedCandidates.Count + @($orphanDecisions | Where-Object Action -eq 'Remove').Count +
        $legacyCaptureEligible.Count
    RemovedActions = @($actions | Where-Object Status -eq 'Removed').Count + @($timestampedDecisions | Where-Object Result -eq 'Removed').Count +
        @($orphanDecisions | Where-Object Result -eq 'Removed').Count + @($legacyCaptureDecisions | Where-Object Result -eq 'Removed').Count
    FailedActions = @($actions | Where-Object Status -like 'Failed:*').Count + @($timestampedDecisions | Where-Object Result -like 'Failed:*').Count +
        @($orphanDecisions | Where-Object Result -like 'Failed:*').Count + @($legacyCaptureDecisions | Where-Object Result -like 'Failed:*').Count +
        $script:failedRunFolders
    SkippedOrphanTemporaries = @($orphanDecisions | Where-Object Result -like 'Skipped:*').Count
    SkippedDailyCaptureLegacy = @($legacyCaptureDecisions | Where-Object Result -like 'Skipped:*').Count
    RemovedRunFolders = $script:removedRunFolders
    ReportDirectory = $reportRoot
}
$result
if ($result.FailedActions -gt 0) { throw 'One or more verified cleanup actions failed; inspect verified-actions.csv, timestamped-csv.csv, orphan-temporaries.csv and daily-capture-legacy.csv.' }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBEYRMg1kGBdfbr
# EXi2e+r9BPqTS5GpcwjD/05UDb3bK6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIA1ia6/fqrG9UaCEESGiATxi3xw5OEYs3dOdeKiio/QZMA0GCSqG
# SIb3DQEBAQUABIIBgH9BhapDazzzX4j9LUlpUzwAy874937ToQt6wuupqK90X2fP
# oud9ErpDFGK1RZM5WP5l/6vbDYSEJD2DgIBzvt54EPJ7VGSa29m7Y9UdINCphuXc
# cHTEXX1ITBCPrbYLNlzWpnT1LDkYoUMuF18bmRM6+hUNjNinHqhlVaCK8cLbYyJt
# qnuxND3GMp1x4yNpjPFYGzXyFmPh0usWGjaDTGn1iHAKagC3r6oiN5dLfnA+0X6T
# LMAj/EFEDWAJ9N0bjyN7WjYVXyFsQvNj40ByyiSyh2Y1n8Np8A++s9gAXsuEHkFM
# kA4IuAybwbTUaC0cxE/Z+cc8t8LkK0tYyvzso77iYNQNzv60/4mEHz2AhiCEc1gI
# ykA7SUViaYUqIzpQcMtxWbi4L3rgGr3+CdZcRGJnMT87fQQesSj7SF0xXBm5LpeK
# MBEpBBPhAdsr5sXe2w8PVjs9J7Q9ISys+ppb4/klRaYa8W2jNVlItDbK4YNMDDE7
# Uhb3Y31E+o5g7xiA76GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkyMDM1
# NDJaMC8GCSqGSIb3DQEJBDEiBCCND3toapCpR2grHXCMuG5ULsRUaAMnfhYaOKpN
# iH3xNTANBgkqhkiG9w0BAQEFAASCAgCo6uVCoT77pMeiyX72lPgBvWkMPrDVEfbQ
# Egz920qlBgPO3Zx0EBJHiNYQKv2W1G6+aBxLMOZmPHm5BNtkaOWg3Y6ki8x0+EZl
# ZgH/wkhrB4ylj5llQ0TMA0K0vDUPdREI6pcHvoyaA7Drubm+MyPExQCNLVyisKLI
# a/R+Q+i42M3q06YTiQ+bu8qmrY5y7H+oyH9lETsx+cJbrq9G5nsieAwbXjGLY172
# ZGlBXsLRJBCvnF+2mpy7cOIGKdSuC1sXlCEnOl3XmYf5R+6gKPlsbgXBerzhG8BA
# LBIG/ZqCp093UJLwLFcCvEuALC2fnNZsKIR9uF/aaRpxS0oAvly+hQ0hVqobfnDL
# Psc51uMEu/VwPu6A2T7kmg5E8M7qgRnHfN5YVP3IsN3GgEF8fRBRUAwWGjdocud8
# fNMg5oH5NE0EZjJlpQuxbcUZ8ra+y21siiPvEhn9JkOiCVWn9OajukcX29xXeG/a
# wcLK1qb1d/SqjeFKr1SNLCfDal+Yt1qZW9i37yFABD2HOVKBmB1AmvVjUIalTUNC
# +r/MgJC7nIvA2y51aIYWpNNwU/VG0BScqxiaw61c4WJ6vPG9Zl7TQSdlYnzC+H/S
# ZL1l8BlTf99/Y9v0hLRuqNnrxGOmfBHwflmPCoADXDaAI7aJIN/Kcc+Xi4iTe4S1
# yEqB4+e40w==
# SIG # End signature block
