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


# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCEIB4dT0ATABuU
# O0HHQxRC+ZLTJboSB3Cq5v6Qvu3H0qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIMERZaUU3vh1oBQMh6BWJZHlZhI07He+AaZnZfGvTaCMA0GCSqG
# SIb3DQEBAQUABIIBgC6V/WcfmRynb9U4+WqRmrdyIqKfcW+9zVIh7OO0TvPM44WZ
# im6/c595f/zobjb62VaAFrSZsIJIAAAC6Da1DkPdFUXBXnbKMsrfpozMQzeHi+Ej
# c//+P0h2tchBKp8W2r0FBIncCB8zZ79RQvAyxezM9/VCTe67H2t9+6YcxCPEOrLi
# cpSkAYWwkC1cs6m8N6so6zCp0uNDo1IaXnmn/arLhCjhQ35x04cImm+TlR1AbNYu
# d9YDhypIpB2mHY4IJbCddEtTLGO/4dl7zJ0f16cp/7DLMrez1zWEGEbseM3N807K
# SFzkd/C6InB+QNEfXRINhxKhtd6UiB6W4stIsS9xyCj6+hcCWx/n5HltzMzBC6Mf
# kRtsmkTmcw6ElAM29BZ+/xVzcgssbQ9KPpVcooyRqfAbI13YHO6GB26kjlJ8Mebf
# 7OKRZlsjbXoKfW67iMzfaUDIQ0kNrY5M6e6dwewfylR0zzTicB2rgW1CqQoTu64v
# BPugKqrDspQpnhieLKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkxOTU5
# MjJaMC8GCSqGSIb3DQEJBDEiBCCEiki2gx/LrZAAiiZ/lwHpU6WPRLbwSE7+lzD8
# DZmPqzANBgkqhkiG9w0BAQEFAASCAgCFbO/GwKJh1DJkBZQ7zH0LOtI0EmpVfqfc
# XmZ5pgjbtHvTXkFEmi0mIJaAMevETxyTDwNv51q9oW0gBBttrFMM1NlGFPyuwozd
# 39fIdCeWbwwAvTa4n81pl+t1/88JN3TLM5PvW2sIuiAbmsx5lMWFyfMcUweI5a9D
# Ehj+nbt6ZAfUMzOD1dNTEmLClYj2UxJEAiTMq0DivwuPorBSlISGKhJNjQM7pka0
# Tiu/IB2Imxj+VTEntIhHjtV7CS09DLeBi98la8FWu9n5g3hJ6apq1RLY+tVlU122
# CmpEoXjmsxSYhLTwQNt7h4gO/YXqBuaGyrQdGX8Us+lRWSW086MDVtAckmNZHgZZ
# WhlN9mmN7ZUqJcvYc1HJHnDKUebvlb1Yg5/t9NaoxmjjlcDR267PihL5rdheoUdC
# VX9MMKTW4SBcM/1wuhvFC9G1KmFi5IL8aVXTdsCoe4mrBwtJQwCifyDLThShtAaF
# qXas4Mcs0gUqugOCoHrgR4374l2KMCiGutY1k7WhOskaprBkN6icnE8SzMSR2jge
# hMRnd2foBsG/DJ5KEOypgOHaZtMq0dVtRR9KHKGjSaMOoJ/m7bWbNHiQJsT0CAWZ
# oqUQ8y+lNTCFV/QwWFKPSTp83nopcCN7nmx5Y1WrKGCCRI5y61QdSvM3HY2JbqJS
# +qOLZKHAQQ==
# SIG # End signature block
