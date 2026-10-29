<#
.SYNOPSIS
Offline check of the daily capture legacy rule of Invoke-SmartM365-ServerDataCleanup.
.VERSION
1.0.0
.NOTES
Builds a synthetic tenant DATA root under TEMP with daily capture roots and runs the cleanup script
in a child PowerShell 7 process, so a capture retention lock held here is a real lock from another
process. Checks that an unrecorded timestamped CSV older than the seven-day window is removed only
when a recorded capture inside the window holds the same export; recorded captures, files inside
the window or recently written, other exports, sub-folders, non-timestamped files and roots with an
unrecognised record are kept; a held retention lock defers removal. No tenant or network access.
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

function Get-Stamp { param([int]$DaysAgo, [string]$Time = '040000') return (Get-Date).AddDays(-$DaysAgo).ToString('yyyyMMdd') + '_' + $Time }
function Get-Day { param([int]$DaysAgo) return (Get-Date).AddDays(-$DaysAgo).ToString('yyyy-MM-dd') }

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
    param([string]$File, [string]$Path, [int64]$Bytes = -1)
    $item = Get-Item -LiteralPath $Path
    return [ordered]@{ File = $File; Relative = $item.Name; Bytes = if ($Bytes -ge 0) { $Bytes } else { $item.Length }; SHA256 = ('0' * 64) }
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
try {
    [void][IO.Directory]::CreateDirectory((Join-Path $root 'DATA-LAST'))
    $files = [ordered]@{
        RecentCapture   = Write-SyntheticFile "$scope/Scope_Users_$(Get-Stamp 1).csv" 1
        OwnedOld        = Write-SyntheticFile "$scope/Scope_Users_$(Get-Stamp 10).csv" 10
        LegacyOld       = Write-SyntheticFile "$scope/Scope_Users_$(Get-Stamp 12).csv" 12
        LegacyOlder     = Write-SyntheticFile "$scope/Scope_Users_$(Get-Stamp 20 '231500').csv" 20
        OtherExport     = Write-SyntheticFile "$scope/Scope_Groups_$(Get-Stamp 12).csv" 12
        InWindow        = Write-SyntheticFile "$scope/Scope_Users_$(Get-Stamp 3 '050000').csv" 3
        RecentWrite     = Write-SyntheticFile "$scope/Scope_Users_$(Get-Stamp 15).csv" 1
        Latest          = Write-SyntheticFile "$scope/Scope_Users.csv" 12
        SubFolder       = Write-SyntheticFile "$scope/Archive/Scope_Users_$(Get-Stamp 12).csv" 12
        BrokenRecent    = Write-SyntheticFile "$broken/Config_Policies_$(Get-Stamp 1).csv" 1
        BrokenLegacy    = Write-SyntheticFile "$broken/Config_Policies_$(Get-Stamp 12).csv" 12
        MismatchRecent  = Write-SyntheticFile "$mismatch/Licences_Users_$(Get-Stamp 1).csv" 1
        MismatchLegacy  = Write-SyntheticFile "$mismatch/Licences_Users_$(Get-Stamp 12).csv" 12
    }
    $retiredName = "Scope_Users_$(Get-Stamp 9).csv"
    $scopeLock = Write-CaptureRecord $scope 'run-recent' (Get-Day 1) @(Get-CaptureFile 'Scope_Users.csv' $files.RecentCapture)
    $null = Write-CaptureRecord $scope 'run-old' (Get-Day 10) @((Get-CaptureFile 'Scope_Users.csv' $files.OwnedOld))
    $null = Write-CaptureRecord $scope 'run-retired' (Get-Day 9) @([ordered]@{ File = 'Scope_Users.csv'; Relative = $retiredName; Bytes = 9; SHA256 = ('0' * 64) }) -RemovedFiles @($retiredName)
    $null = Write-CaptureRecord $broken 'run-broken' (Get-Day 1) @(Get-CaptureFile 'Config_Policies.csv' $files.BrokenRecent) -Owner 'Someone-Else'
    $null = Write-CaptureRecord $mismatch 'run-mismatch' (Get-Day 1) @(Get-CaptureFile 'Licences_Users.csv' $files.MismatchRecent 999)
    $removed = @('LegacyOld', 'LegacyOlder')
    $expectedReasons = @{
        RecentCapture = 'CaptureOwned'; OwnedOld = 'CaptureOwned'; OtherExport = 'NoRecentRecordedCapture'
        InWindow = 'WithinCaptureWindow'; RecentWrite = 'RecentWrite'; BrokenRecent = 'UnreadableCaptureRecords'
        BrokenLegacy = 'UnreadableCaptureRecords'; MismatchRecent = 'CaptureOwned'; MismatchLegacy = 'NoRecentRecordedCapture'
    }

    # Preview never deletes and plans exactly the eligible names.
    $preview = Invoke-Cleanup
    Check ($preview.ExitCode -eq 0) "Preview failed: $($preview.Output)"
    foreach ($name in $files.Keys) { Check (Test-Path -LiteralPath $files[$name]) "Preview removed $name." }
    $report = @(Import-Csv (Join-Path $reportFolder 'daily-capture-legacy.csv'))
    $planned = @($report | Where-Object Action -eq 'Remove' | ForEach-Object Name)
    $expectedNames = @($removed | ForEach-Object { Split-Path $files[$_] -Leaf })
    Check (@(Compare-Object ($planned | Sort-Object) ($expectedNames | Sort-Object)).Count -eq 0) "Preview planned: $($planned -join ', ')"
    foreach ($name in $expectedReasons.Keys) {
        $row = @($report | Where-Object Name -eq (Split-Path $files[$name] -Leaf))
        if ($name -eq 'RecentCapture') { $row = @($row | Where-Object CaptureRoot -like '*WorkplaceScope') }
        Check ($row.Count -eq 1 -and $row[0].Reason -eq $expectedReasons[$name]) "$name reason: $(@($row | ForEach-Object Reason) -join ', ')"
    }
    Check (-not @($report | Where-Object Name -eq 'Scope_Users.csv').Count) 'A non-timestamped latest file appears in the legacy report.'
    # SubFolder shares the LegacyOld name: a single row proves the sub-folder is not scanned.
    Check (@($report | Where-Object Name -eq (Split-Path $files.SubFolder -Leaf)).Count -eq 1) 'A capture root sub-folder was scanned.'
    Check (-not @($report | Where-Object CaptureRoot -like '*Archive*').Count) 'A sub-folder was treated as a capture root.'
    Check ($preview.Output -match 'Daily capture legacy CSVs: \d+ unrecorded timestamped file\(s\) in 3 capture root\(s\), 2 eligible') "Preview summary line missing: $($preview.Output)"
    $inventory = @(Import-Csv (Join-Path $reportFolder 'all-files.csv'))
    Check (@($inventory | Where-Object { $_.Decision -eq 'PlannedRemoval' -and $_.RelativePath -like '*Scope_Users_*' }).Count -eq 2) 'Inventory does not plan exactly the two legacy files.'
    Check (@($inventory | Where-Object { $_.RelativePath -like ('*' + (Split-Path $files.OwnedOld -Leaf)) -and $_.Decision -eq 'Keep' }).Count -eq 1) 'Inventory does not keep a recorded capture.'

    # Execute while another process holds the capture retention lock: nothing is removed.
    $lockHandle = [IO.File]::Open($scopeLock, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    $run = Invoke-Cleanup -Execute
    Check ($run.ExitCode -eq 0) "Execute with held lock failed: $($run.Output)"
    foreach ($name in $files.Keys) { Check (Test-Path -LiteralPath $files[$name]) "$name was removed while the capture lock was held." }
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
    foreach ($name in $removed) {
        Check ($run.Output -match ([regex]::Escape('Removed daily capture legacy CSV') + '.*' + [regex]::Escape((Split-Path $files[$name] -Leaf)))) "Removal of $name was not logged."
    }
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
