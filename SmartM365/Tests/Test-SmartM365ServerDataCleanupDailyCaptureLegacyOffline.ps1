<#
.SYNOPSIS
Offline check of the daily capture legacy rule of Invoke-SmartM365-ServerDataCleanup.
.VERSION
1.1.0
.NOTES
Builds a synthetic tenant DATA root under TEMP with daily capture roots and runs the cleanup script
in a child PowerShell 7 process, so a capture retention lock held here is a real lock from another
process. Checks that an unrecorded timestamped CSV older than the seven-day window is removed only
when a recorded capture inside the window holds the same export; recorded captures, files inside
the window or recently written, other exports, sub-folders, non-timestamped files and roots with an
unrecognised record are kept; a held retention lock defers removal. Run folders (Usage, Copilot) lose
their unrecorded CSVs once the producer has a recorded capture in the window and are removed when
empty; recorded payloads are never removed by the timestamped-CSV or duplicate Usage run rules.
No tenant or network access.
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

$base = Join-Path ([IO.Path]::GetTempPath()) ('sm-capture-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$root = Join-Path $base 'prod'
$reportFolder = Join-Path $base 'report'
$scope = 'DATA-ALL/M365/WorkplaceScope'
$broken = 'DATA-ALL/Intune/WorkplaceScope/Configurations'
$mismatch = 'DATA-ALL/M365/Licensing/Licenses'
$producer = 'SmartM365-WorkplaceScope-Inventory'
$usage = 'DATA-ALL/M365/Usage'
$copilot = 'DATA-ALL/M365/Usage/Copilot'

function Get-Stamp { param([int]$DaysAgo, [string]$Time = '040000') return (Get-Date).AddDays(-$DaysAgo).ToString('yyyyMMdd') + '_' + $Time }
function Get-Day { param([int]$DaysAgo) return (Get-Date).AddDays(-$DaysAgo).ToString('yyyy-MM-dd') }
function Get-IsoWeek { param([int]$DaysAgo) $date = (Get-Date).AddDays(-$DaysAgo); return '{0}-W{1:00}' -f [Globalization.ISOWeek]::GetYear($date), [Globalization.ISOWeek]::GetWeekOfYear($date) }

function Write-SyntheticFile {
    param([string]$RelativePath, [int]$WrittenDaysAgo, [string]$Content = 'synthetic')
    $path = Join-Path $root $RelativePath
    [void][IO.Directory]::CreateDirectory((Split-Path $path -Parent))
    [IO.File]::WriteAllText($path, $Content)
    $time = (Get-Date).AddDays(-$WrittenDaysAgo)
    [IO.File]::SetCreationTime($path, $time)
    [IO.File]::SetLastWriteTime($path, $time)
    return $path
}

function Write-CaptureRecord {
    param([string]$CaptureRoot, [string]$RunId, [string]$Day, [object[]]$Files, [string[]]$RemovedFiles = @(), [string]$Owner = 'SmartM365-DailyCapture')
    $folder = Join-Path (Join-Path $root $CaptureRoot) "DailyCaptureRecords/$producer"
    [void][IO.Directory]::CreateDirectory($folder)
    $record = [ordered]@{
        Owner = $Owner; Version = 1; OutputRoot = (Join-Path $root $CaptureRoot); RunId = $RunId; Status = 'Completed'
        TimeZone = 'Europe/Paris'; Day = $Day; CompletedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
        Files = @($Files); RemovedFiles = @($RemovedFiles); Producer = "$producer.ps1"
    }
    [IO.File]::WriteAllText((Join-Path $folder "$RunId.json.txt"), ($record | ConvertTo-Json -Depth 6))
    $lock = Join-Path $folder '.retention.lock'
    if (-not (Test-Path -LiteralPath $lock)) { [IO.File]::WriteAllText($lock, '') }
    return $lock
}

function Get-CaptureFile {
    param([string]$File, [string]$Path, [int64]$Bytes = -1, [string]$RunFolder = '')
    $item = Get-Item -LiteralPath $Path
    $relative = if ($RunFolder) { "$RunFolder\$($item.Name)" } else { $item.Name }
    return [ordered]@{ File = $File; Relative = $relative; Bytes = if ($Bytes -ge 0) { $Bytes } else { $item.Length }; SHA256 = ('0' * 64) }
}

function Invoke-Cleanup {
    param([switch]$Execute)
    $arguments = @('-NoProfile', '-NonInteractive', '-File', $cleanupScript, '-Tenant', 'prod', '-DataRootPath', $root,
        '-ReportDirectory', $reportFolder, '-EffectiveJobsPath', (Join-Path $base 'no-jobs.json'))
    if ($Execute) { $arguments += '-Execute' }
    $output = & pwsh @arguments 2>&1 | Out-String
    return [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = $output }
}

$lockHandle = $null
$oldRun = Get-Stamp 12 '044500'
$ownedRun = Get-Stamp 10 '044500'
$dupRun = Get-Stamp 11 '044500'
$recentRun = Get-Stamp 1 '044500'
$mixedRun = Get-Stamp 13 '044500'
try {
    [void][IO.Directory]::CreateDirectory((Join-Path $root 'DATA-LAST'))
    $files = [ordered]@{
        RecentCapture   = Write-SyntheticFile "$scope/Scope_Users_$(Get-Stamp 1).csv" 1
        OwnedOld        = Write-SyntheticFile "$scope/Scope_Users_$(Get-Stamp 10).csv" 10
        LegacyOld       = Write-SyntheticFile "$scope/Scope_Users_$(Get-Stamp 12).csv" 12
        LegacyOlder     = Write-SyntheticFile "$scope/Scope_Users_$(Get-Stamp 20 '231500').csv" 20
        OtherExport     = Write-SyntheticFile "$scope/Scope_Groups_$(Get-Stamp 12).csv" 12
        InWindow        = Write-SyntheticFile "$scope/Scope_Users_$(Get-Stamp 3 '050000').csv" 3
        RecentWrite     = Write-SyntheticFile "$scope/Scope_Users_$(Get-Stamp 25).csv" 1
        Latest          = Write-SyntheticFile "$scope/Scope_Users.csv" 12
        SubFolder       = Write-SyntheticFile "$scope/Archive/Scope_Users_$(Get-Stamp 12).csv" 12
        BrokenRecent    = Write-SyntheticFile "$broken/Config_Policies_$(Get-Stamp 1).csv" 1
        BrokenLegacy    = Write-SyntheticFile "$broken/Config_Policies_$(Get-Stamp 12).csv" 12
        MismatchRecent  = Write-SyntheticFile "$mismatch/Licences_Users_$(Get-Stamp 1).csv" 1
        MismatchLegacy  = Write-SyntheticFile "$mismatch/Licences_Users_$(Get-Stamp 12).csv" 12
        UsageOldRaw     = Write-SyntheticFile "$usage/$oldRun/Usage_Users_Raw_D7_$oldRun.csv" 12
        UsageOldExport  = Write-SyntheticFile "$usage/$oldRun/Usage_Users_$oldRun.csv" 12
        UsageOwnedOld   = Write-SyntheticFile "$usage/$ownedRun/Usage_Users_$ownedRun.csv" 10
        UsageOwnedRaw   = Write-SyntheticFile "$usage/$ownedRun/Usage_Users_Raw_D7_$ownedRun.csv" 10
        UsageDuplicate  = Write-SyntheticFile "$usage/$dupRun/Usage_Users_$dupRun.csv" 11
        UsageRecent     = Write-SyntheticFile "$usage/$recentRun/Usage_Users_$recentRun.csv" 1
        UsageRecentRaw  = Write-SyntheticFile "$usage/$recentRun/Usage_Users_Raw_D7_$recentRun.csv" 1
        UsageMixedCsv   = Write-SyntheticFile "$usage/$mixedRun/Usage_Users_$mixedRun.csv" 13
        UsageMixedNote  = Write-SyntheticFile "$usage/$mixedRun/notes.json" 13
        UsageWeekCopy   = Write-SyntheticFile "$usage/WeeklyHistory/$(Get-IsoWeek 11)/Usage_Users.csv" 11
        ScopeWeekCopy   = Write-SyntheticFile "$scope/WeeklyHistory/$(Get-IsoWeek 10)/Scope_Users.csv" 10
        ScopeLatest     = Write-SyntheticFile "DATA-LAST/Scope_Users.csv" 1
        CopilotOld      = Write-SyntheticFile "$copilot/$oldRun/Copilot_Source_$oldRun.csv" 12
        CopilotOwnedOld = Write-SyntheticFile "$copilot/$ownedRun/Copilot_Users_$ownedRun.csv" 10
    }
    $weekManifest = [ordered]@{ Week = (Get-IsoWeek 10); Files = @('Scope_Users.csv') } | ConvertTo-Json
    [IO.File]::WriteAllText((Join-Path $root "$scope/WeeklyHistory/$(Get-IsoWeek 10)/manifest.json.txt"), $weekManifest)
    $retiredName = "Scope_Users_$(Get-Stamp 9).csv"
    $scopeLock = Write-CaptureRecord $scope 'run-recent' (Get-Day 1) @(Get-CaptureFile 'Scope_Users.csv' $files.RecentCapture)
    $null = Write-CaptureRecord $scope 'run-old' (Get-Day 10) @((Get-CaptureFile 'Scope_Users.csv' $files.OwnedOld))
    $null = Write-CaptureRecord $scope 'run-retired' (Get-Day 9) @([ordered]@{ File = 'Scope_Users.csv'; Relative = $retiredName; Bytes = 9; SHA256 = ('0' * 64) }) -RemovedFiles @($retiredName)
    $null = Write-CaptureRecord $broken 'run-broken' (Get-Day 1) @(Get-CaptureFile 'Config_Policies.csv' $files.BrokenRecent) -Owner 'Someone-Else'
    $null = Write-CaptureRecord $mismatch 'run-mismatch' (Get-Day 1) @(Get-CaptureFile 'Licences_Users.csv' $files.MismatchRecent 999)
    $null = Write-CaptureRecord $usage 'usage-recent' (Get-Day 1) @(Get-CaptureFile 'Usage_Users.csv' $files.UsageRecent -RunFolder $recentRun)
    $null = Write-CaptureRecord $usage 'usage-old' (Get-Day 10) @(Get-CaptureFile 'Usage_Users.csv' $files.UsageOwnedOld -RunFolder $ownedRun)
    $null = Write-CaptureRecord $usage 'usage-dup' (Get-Day 11) @(Get-CaptureFile 'Usage_Users.csv' $files.UsageDuplicate -RunFolder $dupRun)
    $null = Write-CaptureRecord $copilot 'copilot-old' (Get-Day 10) @(Get-CaptureFile 'Copilot_Users.csv' $files.CopilotOwnedOld -RunFolder $ownedRun)
    $removed = @('LegacyOld', 'LegacyOlder', 'UsageOldRaw', 'UsageOldExport', 'UsageOwnedRaw', 'UsageMixedCsv')
    $expectedReasons = @{
        RecentCapture = 'CaptureOwned'; OwnedOld = 'CaptureOwned'; OtherExport = 'NoRecentRecordedCapture'
        InWindow = 'WithinCaptureWindow'; RecentWrite = 'RecentWrite'; BrokenRecent = 'UnreadableCaptureRecords'
        BrokenLegacy = 'UnreadableCaptureRecords'; MismatchRecent = 'CaptureOwned'; MismatchLegacy = 'NoRecentRecordedCapture'
        UsageOwnedOld = 'CaptureOwned'; UsageDuplicate = 'CaptureOwned'; UsageRecent = 'CaptureOwned'; UsageRecentRaw = 'WithinCaptureWindow'
        CopilotOld = 'NoRecentRecordedCapture'; CopilotOwnedOld = 'CaptureOwned'
    }

    # Preview never deletes and plans exactly the eligible names.
    $preview = Invoke-Cleanup
    Check ($preview.ExitCode -eq 0) "Preview failed: $($preview.Output)"
    foreach ($name in $files.Keys) { Check (Test-Path -LiteralPath $files[$name]) "Preview removed $name." }
    $report = @(Import-Csv (Join-Path $reportFolder 'daily-capture-legacy.csv'))
    $planned = @($report | Where-Object Action -eq 'Remove' | ForEach-Object { Join-Path $_.CaptureRoot $_.Name })
    $expectedNames = @($removed | ForEach-Object { [IO.Path]::GetRelativePath($root, $files[$_]) })
    Check (@(Compare-Object ($planned | Sort-Object) ($expectedNames | Sort-Object)).Count -eq 0) "Preview planned: $($planned -join ', ')"
    foreach ($name in $expectedReasons.Keys) {
        $row = @($report | Where-Object { (Join-Path $_.CaptureRoot $_.Name) -eq [IO.Path]::GetRelativePath($root, $files[$name]) })
        Check ($row.Count -eq 1 -and $row[0].Reason -eq $expectedReasons[$name]) "$name reason: $(@($row | ForEach-Object Reason) -join ', ')"
    }
    Check (-not @($report | Where-Object Name -eq 'Scope_Users.csv').Count) 'A non-timestamped latest file appears in the legacy report.'
    # SubFolder shares the LegacyOld name: a single row proves the sub-folder is not scanned.
    Check (@($report | Where-Object Name -eq (Split-Path $files.SubFolder -Leaf)).Count -eq 1) 'A capture root sub-folder was scanned.'
    Check (-not @($report | Where-Object CaptureRoot -like '*Archive*').Count) 'A sub-folder was treated as a capture root.'
    Check (-not @($report | Where-Object Name -like '*notes.json').Count) 'A non-CSV run file appears in the legacy report.'
    Check ($preview.Output -match 'Daily capture legacy CSVs: \d+ recorded and \d+ unrecorded timestamped or run-folder file\(s\) in 5 capture root\(s\); 6 eligible \([0-9.,]+ GiB\), 0 left to another rule') "Preview summary line missing: $($preview.Output)"
    # Recorded payloads stay out of the timestamped-CSV and duplicate Usage run rules.
    $timestamped = @(Import-Csv (Join-Path $reportFolder 'timestamped-csv.csv'))
    $ownedRow = @($timestamped | Where-Object Name -eq (Split-Path $files.OwnedOld -Leaf))
    Check ($ownedRow.Count -eq 1 -and $ownedRow[0].Reason -eq 'DailyCaptureRoot') "A recorded capture is not protected from the timestamped rule: $(@($ownedRow | ForEach-Object Reason) -join ', ')"
    Check (-not @($timestamped | Where-Object Action -eq 'Remove').Count) 'The timestamped rule plans a file inside a capture root.'
    Check (-not @(Import-Csv (Join-Path $reportFolder 'verified-actions.csv') | Where-Object Category -eq 'DuplicateUsageRun').Count) 'A duplicate Usage run was planned inside a capture root.'
    Check ($preview.Output -match 'duplicate-run retirement left to the capture rules') 'The Usage duplicate-run skip was not logged.'
    Check ($preview.Output -match 'PlannedActions\s+:\s+6\b') "Planned actions are not counted once: $($preview.Output)"
    $inventory = @(Import-Csv (Join-Path $reportFolder 'all-files.csv'))
    Check (@($inventory | Where-Object Decision -eq 'PlannedRemoval').Count -eq 6) "Inventory does not plan exactly the six legacy files: $(@($inventory | Where-Object Decision -eq 'PlannedRemoval' | ForEach-Object RelativePath) -join ', ')"
    Check (@($inventory | Where-Object { $_.RelativePath -like ('*' + (Split-Path $files.OwnedOld -Leaf)) -and $_.Decision -eq 'Keep' }).Count -eq 1) 'Inventory does not keep a recorded capture.'

    # Execute while another process holds the WorkplaceScope retention lock: that root is untouched, the others proceed.
    $lockHandle = [IO.File]::Open($scopeLock, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    $run = Invoke-Cleanup -Execute
    Check ($run.ExitCode -eq 0) "Execute with held lock failed: $($run.Output)"
    foreach ($name in $files.Keys) {
        $exists = Test-Path -LiteralPath $files[$name]
        if ($files[$name] -like "*$($scope.Replace('/', [IO.Path]::DirectorySeparatorChar))*" -or $name -notin $removed) { Check $exists "$name was removed while the WorkplaceScope lock was held." }
        else { Check (-not $exists) "$name in an unlocked root was not removed." }
    }
    Check (-not (Test-Path -LiteralPath (Join-Path $root "$usage/$oldRun"))) 'An emptied run folder was not removed.'
    foreach ($kept in $ownedRun, $recentRun, $mixedRun, $dupRun) { Check (Test-Path -LiteralPath (Join-Path $root "$usage/$kept")) "Run folder $kept was removed although not empty." }
    Check ($run.Output -match 'RemovedRunFolders\s+:\s+1\b') "Removed run folder count wrong: $($run.Output)"
    Check ($run.Output -match 'RemovedActions\s+:\s+4\b') "Removed actions wrong with one root locked: $($run.Output)"
    $held = @(Import-Csv (Join-Path $reportFolder 'daily-capture-legacy.csv') | Where-Object Action -eq 'Remove')
    Check (@($held | Where-Object Result -eq 'Skipped: CaptureRetentionLockHeld').Count -eq 2) "Held lock not reported: $(@($held | ForEach-Object Result) -join ', ')"

    # Once the lock is released, exactly the legacy files go.
    $lockHandle.Dispose(); $lockHandle = $null
    $run = Invoke-Cleanup -Execute
    Check ($run.ExitCode -eq 0) "Execute failed: $($run.Output)"
    foreach ($name in $files.Keys) {
        $exists = Test-Path -LiteralPath $files[$name]
        if ($name -in $removed) { Check (-not $exists) "$name was not removed." }
        else { Check $exists "$name was removed." }
    }
    foreach ($name in 'LegacyOld', 'LegacyOlder') {
        Check ($run.Output -match ([regex]::Escape('Removed daily capture legacy CSV') + '.*' + [regex]::Escape((Split-Path $files[$name] -Leaf)))) "Removal of $name was not logged."
    }
    foreach ($kept in $ownedRun, $recentRun, $mixedRun, $dupRun) { Check (Test-Path -LiteralPath (Join-Path $root "$usage/$kept")) "Run folder $kept was removed although not empty." }
    Check ($run.Output -match 'RemovedActions\s+:\s+2\b') "Removed actions not counted once: $($run.Output)"
    Check (Test-Path -LiteralPath $scopeLock) 'The capture retention lock file was removed.'
    Check (@(Get-ChildItem -LiteralPath (Join-Path $root "$scope/DailyCaptureRecords/$producer") -Filter '*.json.txt').Count -eq 3) 'Capture records were changed.'

    # A further run finds nothing left to do.
    $run = Invoke-Cleanup -Execute
    Check ($run.ExitCode -eq 0) "Third execute failed: $($run.Output)"
    Check (-not @(Import-Csv (Join-Path $reportFolder 'daily-capture-legacy.csv') | Where-Object Action -eq 'Remove').Count) 'Third run still plans removals.'
}
finally {
    if ($null -ne $lockHandle) { $lockHandle.Dispose() }
    if (Test-Path -LiteralPath $base) { Remove-Item -LiteralPath $base -Recurse -Force }
}
"Test-SmartM365ServerDataCleanupDailyCaptureLegacyOffline: $script:checks checks passed."

