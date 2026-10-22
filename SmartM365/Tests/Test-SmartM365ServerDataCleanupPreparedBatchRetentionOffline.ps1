<#
.SYNOPSIS
Offline check of the linked prepared-batch retention of Invoke-SmartM365-ServerDataCleanup.
.VERSION
1.0.0
.NOTES
Builds a synthetic tenant DATA root under TEMP with a published chain current -> previous -> three
older batches and runs the cleanup script in a child PowerShell 7 process. Checks that by default
no linked batch is planned or removed and all are inventoried as kept; that enabled retirement
honours the retained count and the minimum age, never touches current and previous, and preserves
SHA256-identical audit receipts. No tenant, network or synchronized data access.
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

$base = Join-Path ([IO.Path]::GetTempPath()) ('sm-batches-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$root = Join-Path $base 'prod'
$reportFolder = Join-Path $base 'report'
$prepared = Join-Path $root 'DATA-POWERBI'
$tenantKey = 'synthetic-tenant-key'

function Get-SyntheticBatchId {
    param([int]$AgeDays, [string]$Suffix)
    return [datetime]::UtcNow.AddDays(-$AgeDays).ToString('yyyyMMddTHHmmssfff', [Globalization.CultureInfo]::InvariantCulture) + 'Z-' + $Suffix
}

function Write-Json {
    param([string]$Path, [object]$Value)
    [void][IO.Directory]::CreateDirectory((Split-Path $Path -Parent))
    [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 4))
}

function Write-SyntheticBatch {
    param([string]$Id, [string]$PreviousId)
    $folder = Join-Path (Join-Path $prepared 'batches') $Id
    Write-Json (Join-Path $folder 'batch.json.txt') ([ordered]@{ BatchId = $Id; TenantKey = $tenantKey })
    Write-Json (Join-Path $folder 'validation.json.txt') ([ordered]@{ BatchId = $Id; Status = 'Valid' })
    Write-Json (Join-Path $folder 'tables/Table.csv') 'synthetic'
    $manifestHash = (Get-FileHash -LiteralPath (Join-Path $folder 'batch.json.txt') -Algorithm SHA256).Hash
    Write-Json (Join-Path $folder 'current.json.txt') ([ordered]@{
        BatchId = $Id; PreviousBatchId = $PreviousId; TenantKey = $tenantKey; ManifestSHA256 = $manifestHash })
    return $folder
}

function Invoke-Cleanup {
    param([string[]]$Extra = @())
    $arguments = @('-NoProfile', '-NonInteractive', '-File', $cleanupScript, '-Tenant', 'prod', '-DataRootPath', $root,
        '-ReportDirectory', $reportFolder, '-EffectiveJobsPath', (Join-Path $base 'no-jobs.json')) + $Extra
    $output = & pwsh @arguments 2>&1 | Out-String
    return [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = $output }
}

function Get-PlannedBatchId {
    return @(Import-Csv (Join-Path $reportFolder 'verified-actions.csv') |
        Where-Object Category -eq 'LinkedPreparedBatch' | ForEach-Object { Split-Path $_.RelativePath -Leaf } | Sort-Object)
}

try {
    [void][IO.Directory]::CreateDirectory((Join-Path $root 'DATA-ALL'))
    [void][IO.Directory]::CreateDirectory((Join-Path $root 'DATA-LAST'))
    $ids = [ordered]@{
        Current  = Get-SyntheticBatchId 1 'aaaaaaa1'
        Previous = Get-SyntheticBatchId 2 'aaaaaaa2'
        Older1   = Get-SyntheticBatchId 5 'aaaaaaa3'
        Older2   = Get-SyntheticBatchId 100 'aaaaaaa4'
        Older3   = Get-SyntheticBatchId 200 'aaaaaaa5'
    }
    $folders = [ordered]@{
        Current  = Write-SyntheticBatch $ids.Current $ids.Previous
        Previous = Write-SyntheticBatch $ids.Previous $ids.Older1
        Older1   = Write-SyntheticBatch $ids.Older1 $ids.Older2
        Older2   = Write-SyntheticBatch $ids.Older2 $ids.Older3
        Older3   = Write-SyntheticBatch $ids.Older3 ''
    }
    Write-Json (Join-Path $prepared 'current.json.txt') ([ordered]@{
        BatchId = $ids.Current; PreviousBatchId = $ids.Previous; TenantKey = $tenantKey })

    # Default: the share keeps the whole linked history, in Preview and Execute.
    $preview = Invoke-Cleanup
    Check ($preview.ExitCode -eq 0) "Default preview failed: $($preview.Output)"
    Check (@(Get-PlannedBatchId).Count -eq 0) 'Default preview planned a linked prepared batch.'
    Check ($preview.Output -match '3 found, 0 planned for retirement, 3 kept .*retirement not enabled') "Default retention was not logged: $($preview.Output)"
    $inventory = @(Import-Csv (Join-Path $reportFolder 'all-files.csv'))
    foreach ($name in 'Older1', 'Older2', 'Older3') {
        $rows = @($inventory | Where-Object { $_.RelativePath -like ('DATA-POWERBI\batches\' + $ids[$name] + '\*') })
        Check ($rows.Count -eq 4) "Inventory misses the $name batch files."
        Check (-not @($rows | Where-Object { $_.Decision -ne 'Keep' -or $_.Reason -notlike 'Linked prepared batch kept: share retains*' }).Count) "$name batch is not inventoried as kept by default."
    }
    $run = Invoke-Cleanup @('-Execute')
    Check ($run.ExitCode -eq 0) "Default execute failed: $($run.Output)"
    foreach ($name in $folders.Keys) { Check (Test-Path -LiteralPath $folders[$name]) "Default execute removed the $name batch." }
    Check (-not (Test-Path -LiteralPath (Join-Path $prepared 'retired'))) 'Default execute wrote retired receipts.'

    # Enabled: the retained count and the minimum age both keep batches.
    $preview = Invoke-Cleanup @('-RetireLinkedPreparedBatches', '-RetainedLinkedPreparedBatchCount', '1')
    Check ($preview.ExitCode -eq 0) "Count preview failed: $($preview.Output)"
    Check (@(Compare-Object @(Get-PlannedBatchId) (@($ids.Older2, $ids.Older3) | Sort-Object)).Count -eq 0) "Count preview planned: $(@(Get-PlannedBatchId) -join ', ')"
    $kept = @(Import-Csv (Join-Path $reportFolder 'all-files.csv') | Where-Object { $_.RelativePath -like ('DATA-POWERBI\batches\' + $ids.Older1 + '\*') })
    Check (-not @($kept | Where-Object { $_.Decision -ne 'Keep' -or $_.Reason -notlike '*within the 1 retained older batches*' }).Count) 'Retained-count batch is not inventoried as kept.'

    $retireOldest = @('-RetireLinkedPreparedBatches', '-RetainedLinkedPreparedBatchCount', '1', '-LinkedPreparedBatchMinimumAgeDays', '150')
    $preview = Invoke-Cleanup $retireOldest
    Check ($preview.ExitCode -eq 0) "Age preview failed: $($preview.Output)"
    Check (@(Compare-Object @(Get-PlannedBatchId) @($ids.Older3)).Count -eq 0) "Age preview planned: $(@(Get-PlannedBatchId) -join ', ')"

    $receiptHashes = @{}
    foreach ($receipt in 'current', 'batch', 'validation') {
        $receiptHashes[$receipt] = (Get-FileHash -LiteralPath (Join-Path $folders.Older3 ($receipt + '.json.txt')) -Algorithm SHA256).Hash
    }
    $run = Invoke-Cleanup ($retireOldest + '-Execute')
    Check ($run.ExitCode -eq 0) "Age execute failed: $($run.Output)"
    Check (-not (Test-Path -LiteralPath $folders.Older3)) 'The oldest batch was not retired.'
    foreach ($name in 'Current', 'Previous', 'Older1', 'Older2') { Check (Test-Path -LiteralPath $folders[$name]) "Age execute removed the $name batch." }
    foreach ($receipt in $receiptHashes.Keys) {
        $copy = Join-Path (Join-Path (Join-Path $prepared 'retired') $ids.Older3) ($receipt + '.json.txt')
        Check ((Test-Path -LiteralPath $copy) -and (Get-FileHash -LiteralPath $copy -Algorithm SHA256).Hash -eq $receiptHashes[$receipt]) "Audit receipt $receipt is missing or differs."
    }

    # Fully enabled: every older linked batch goes, current and previous always stay.
    $run = Invoke-Cleanup @('-RetireLinkedPreparedBatches', '-Execute')
    Check ($run.ExitCode -eq 0) "Full execute failed: $($run.Output)"
    foreach ($name in 'Older1', 'Older2') { Check (-not (Test-Path -LiteralPath $folders[$name])) "Full execute kept the $name batch." }
    foreach ($name in 'Current', 'Previous') { Check (Test-Path -LiteralPath (Join-Path $folders[$name] 'tables/Table.csv')) "Full execute touched the $name batch." }
    foreach ($name in 'Older1', 'Older2', 'Older3') {
        Check (@(Get-ChildItem -LiteralPath (Join-Path (Join-Path $prepared 'retired') $ids[$name]) -File).Count -eq 3) "Retired receipts of $name are incomplete."
    }
}
finally {
    if (Test-Path -LiteralPath $base) { Remove-Item -LiteralPath $base -Recurse -Force }
}
"Test-SmartM365ServerDataCleanupPreparedBatchRetentionOffline: $script:checks checks passed."
