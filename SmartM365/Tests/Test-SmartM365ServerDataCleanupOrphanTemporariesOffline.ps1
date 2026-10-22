<#
.SYNOPSIS
Offline check of the orphan atomic-write temporary rule of Invoke-SmartM365-ServerDataCleanup.
.VERSION
1.0.0
.NOTES
Builds a synthetic tenant DATA root under TEMP and runs the cleanup script in a child PowerShell 7
process, so a transport lock held here is a real lock from another process. Checks that an old
orphan is removed, a recent one and an old one beside a held transport lock are kept, a .previous
copy uses its longer age, weekly upload.pending markers and near-miss names are never touched,
and every removal is logged. No tenant, network or synchronized data access.
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

$base = Join-Path ([IO.Path]::GetTempPath()) ('sm-orphan-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$root = Join-Path $base 'prod'
$reportFolder = Join-Path $base 'report'
$old = (Get-Date).AddDays(-10)
$older = (Get-Date).AddDays(-40)
$recent = (Get-Date).AddDays(-1)

function Write-SyntheticFile {
    param([string]$RelativePath, [datetime]$Time)
    $path = Join-Path $root $RelativePath
    [void][IO.Directory]::CreateDirectory((Split-Path $path -Parent))
    [IO.File]::WriteAllText($path, 'synthetic')
    [IO.File]::SetCreationTime($path, $Time)
    [IO.File]::SetLastWriteTime($path, $Time)
    return $path
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
    $files = [ordered]@{
        OldOrphan       = Write-SyntheticFile 'DATA-LAST/~sm0123456789ab.pending' $old
        OldCsvOrphan    = Write-SyntheticFile 'DATA-LAST/~sm0123456789ac.csv' $old
        RecentOrphan    = Write-SyntheticFile 'DATA-LAST/~smaaaaaaaaaaaa.tmp' $recent
        RecentPrevious  = Write-SyntheticFile 'DATA-ALL/Orchestrator/SRV/~smcccccccccccc.previous' $old
        OldPrevious     = Write-SyntheticFile 'DATA-ALL/Orchestrator/SRV/~smdddddddddddd.previous' $older
        LegacyOrphan    = Write-SyntheticFile ('DATA-ALL/Config/state.json.txt.' + ('e' * 32) + '.pending') $old
        LegacyProbe     = Write-SyntheticFile ('DATA-ALL/.smartm365-write-test-' + ('f' * 32) + '.tmp') $old
        LockedOrphan    = Write-SyntheticFile 'DATA-ALL/Locked/~smbbbbbbbbbbbb.pending' $old
        LockFile        = Write-SyntheticFile 'DATA-ALL/Locked/state.json.transport.lock' $old
        UploadMarker    = Write-SyntheticFile 'DATA-ALL/Intune/Apps/WeeklyHistory/2026-W30/upload.pending' $older
        BadHex          = Write-SyntheticFile 'DATA-LAST/~smXYZXYZXYZXYZ.pending' $old
        LongHex         = Write-SyntheticFile 'DATA-LAST/~sm0123456789abc.tmp' $old
        OtherExtension  = Write-SyntheticFile 'DATA-LAST/~sm0123456789ab.exe' $old
        LegacyCsvGuid   = Write-SyntheticFile ('DATA-LAST/Export.' + ('a' * 32) + '.csv') $old
        CanonicalCsv    = Write-SyntheticFile 'DATA-LAST/Export.csv' $old
    }
    $removed = @('OldOrphan', 'OldCsvOrphan', 'OldPrevious', 'LegacyOrphan', 'LegacyProbe')

    # Preview never deletes and plans exactly the eligible names (no lock is held yet).
    $preview = Invoke-Cleanup
    Check ($preview.ExitCode -eq 0) "Preview failed: $($preview.Output)"
    foreach ($name in $files.Keys) { Check (Test-Path -LiteralPath $files[$name]) "Preview removed $name." }
    $planned = @(Import-Csv (Join-Path $reportFolder 'orphan-temporaries.csv') | Where-Object Action -eq 'Remove' | ForEach-Object Name)
    $expectedNames = @(@($removed) + 'LockedOrphan' | ForEach-Object { Split-Path $files[$_] -Leaf })
    Check (@(Compare-Object ($planned | Sort-Object) ($expectedNames | Sort-Object)).Count -eq 0) "Preview planned: $($planned -join ', ')"
    $report = @(Import-Csv (Join-Path $reportFolder 'orphan-temporaries.csv'))
    Check (-not @($report | Where-Object Name -eq 'upload.pending').Count) 'upload.pending appears in the orphan report.'
    Check (@($report | Where-Object { $_.Name -eq '~smcccccccccccc.previous' -and $_.Reason -eq 'Recent' -and $_.MinimumAgeDays -eq '30' }).Count -eq 1) 'A 10-day .previous copy is not kept by its 30-day age.'

    # Execute while another process holds the transport lock of one folder.
    $lockHandle = [IO.File]::Open($files.LockFile, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    $run = Invoke-Cleanup -Execute
    Check ($run.ExitCode -eq 0) "Execute failed: $($run.Output)"
    foreach ($name in $files.Keys) {
        $exists = Test-Path -LiteralPath $files[$name]
        if ($name -in $removed) { Check (-not $exists) "$name was not removed." }
        else { Check $exists "$name was removed." }
    }
    foreach ($name in $removed) {
        Check ($run.Output -match [regex]::Escape('Removed orphan temporary') + '.*' + [regex]::Escape((Split-Path $files[$name] -Leaf))) "Removal of $name was not logged."
    }
    $locked = @(Import-Csv (Join-Path $reportFolder 'orphan-temporaries.csv') | Where-Object Name -eq '~smbbbbbbbbbbbb.pending')
    Check ($locked.Count -eq 1 -and $locked[0].Reason -eq 'TransportLockHeld') 'An orphan beside a held transport lock was not kept as TransportLockHeld.'

    # Once the lock is released, the same old orphan becomes eligible; upload.pending still stays.
    $lockHandle.Dispose(); $lockHandle = $null
    $run = Invoke-Cleanup -Execute
    Check ($run.ExitCode -eq 0) "Second execute failed: $($run.Output)"
    Check (-not (Test-Path -LiteralPath $files.LockedOrphan)) 'Orphan was not removed after its folder lock was released.'
    Check (Test-Path -LiteralPath $files.LockFile) 'The transport lock file itself was removed.'
    Check (Test-Path -LiteralPath $files.UploadMarker) 'upload.pending was removed.'
    Check (Test-Path -LiteralPath $files.RecentOrphan) 'A recent orphan was removed.'
}
finally {
    if ($null -ne $lockHandle) { $lockHandle.Dispose() }
    if (Test-Path -LiteralPath $base) { Remove-Item -LiteralPath $base -Recurse -Force }
}
"Test-SmartM365ServerDataCleanupOrphanTemporariesOffline: $script:checks checks passed."
