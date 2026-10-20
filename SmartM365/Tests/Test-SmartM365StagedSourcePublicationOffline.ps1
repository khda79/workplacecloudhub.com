#Requires -Version 7.0
<#
.SYNOPSIS
Synthetic checks of staged source publication: a running or failed run never changes DATA-LAST,
a complete run promotes its CSVs under the Publishing fence and writes the receipt last.
.VERSION
1.0.0
.NOTES
Synthetic TEMP data only. SharePoint upload and WeeklyHistory are replaced by recorders.
No tenant, API, mail or synchronized data access.
#>
[CmdletBinding()]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='Synthetic Core identity globals are saved and restored in finally.')]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$core = Import-Module (Join-Path $root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.93' -Force -PassThru
Import-Module (Join-Path $root 'SmartInventory/Orchestrator/SmartM365.Orchestrator.Distributed.psm1') -Force
$producer = Get-ChildItem (Join-Path $root 'SmartInventory') -Recurse -File -Filter 'SmartM365-IntuneManagedApps-Inventory.ps1' | Select-Object -First 1
$files = @('Intune_ManagedApps.csv', 'Intune_ManagedAppAssignments.csv', 'Intune_ManagedAppInstallSummary.csv', 'Intune_ManagedAppDeviceStates.csv')
$temporary = Join-Path ([IO.Path]::GetTempPath()) ('SmartInventory-StagedPublication-' + [guid]::NewGuid().ToString('N'))
$latest = Join-Path $temporary 'DATA-LAST'
$output = Join-Path $temporary 'DATA-ALL/Intune/WorkplaceScope/ManagedApps'
$shared = Join-Path $temporary 'DATA-ALL/Orchestrator'
New-Item -ItemType Directory -Path $latest, $output, $shared | Out-Null
$saved = @{}
foreach ($name in @('SmartM365TenantKey', 'SmartM365OrganizationKey', 'SmartM365EnvironmentKey', 'SmartM365TenantId', 'csvGeneratedPaths')) {
    $variable = Get-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue
    $saved[$name] = @{ Exists = ($null -ne $variable); Value = $(if ($variable) { $variable.Value } else { $null }) }
}
$global:SmartM365TenantKey = 'synthetic'; $global:SmartM365OrganizationKey = 'contoso'
$global:SmartM365EnvironmentKey = 'test'; $global:SmartM365TenantId = 'synthetic-tenant'
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }
$receiptPath = Join-Path $latest 'SmartInventory_SmartM365-IntuneManagedApps-Inventory.current.json.txt'
$runPath = Join-Path $latest 'SmartInventory_SmartM365-IntuneManagedApps-Inventory.run.json.txt'
$stagingProducer = Join-Path $latest '.staging/SmartM365-IntuneManagedApps-Inventory'
function Get-LatestHashTable { $result = @{}; foreach ($file in $files) { $path = Join-Path $latest $file; $result[$file] = if (Test-Path -LiteralPath $path) { (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash } else { '' } }; return $result }
function Read-Json([string]$Path) { Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json }
function Publish-Set([string]$Label) {
    foreach ($file in $files) {
        $rows = @([pscustomobject]@{ AppId = [guid]::NewGuid().ToString(); Name = "$Label $file" })
        $stamp = (Get-Date).ToString('yyyyMMdd_HHmmss') + [guid]::NewGuid().ToString('N').Substring(0, 4)
        Publish-SmartM365Csv -Data $rows -Columns @('AppId', 'Name') -TimestampedPath (Join-Path $output ("{0}_{1}.csv" -f [IO.Path]::GetFileNameWithoutExtension($file), $stamp)) `
            -LatestPath (Join-Path $latest $file) -RetentionMaxCsv 0 -SingleSerialization | Out-Null
    }
}
try {
    # Recorders replace SharePoint upload and WeeklyHistory inside the module; they also capture
    # the publication state seen at each promotion step.
    & $core {
        $script:Recorded = New-Object 'Collections.Generic.List[object]'
        function script:Invoke-SmartM365SharePointCsvUpload { param([string]$LocalFilePath) $script:Recorded.Add([pscustomobject]@{ Kind = 'Upload'; Path = $LocalFilePath }) }
        function script:Invoke-SmartM365WeeklyInventoryHistoryForCsv { param([string[]]$SourceFiles, [string]$TimestampedPath) $script:Recorded.Add([pscustomobject]@{ Kind = 'History'; Path = $SourceFiles[0]; Timestamped = $TimestampedPath }) }
        function script:WriteLog {
            param($Message, $Level)
            $null = $Level
            if ([string]$Message -match 'promoted from source staging') {
                $run = Get-ChildItem -LiteralPath ([IO.Path]::GetDirectoryName(($Message -replace '^CSV latest copy written to: (.+) \(promoted.*$', '$1'))) -Filter '*.run.json.txt' | Select-Object -First 1
                $script:Recorded.Add([pscustomobject]@{ Kind = 'Promoted'; Path = $Message; RunStatus = (Get-Content -LiteralPath $run.FullName -Raw | ConvertFrom-Json).Status })
            }
        }
    }
    function Get-Recorded { & $core { $script:Recorded.ToArray() } }
    function Clear-Recorded { & $core { $script:Recorded.Clear() } }

    # Baseline: a direct, unstaged complete run publishes the previous good set.
    $global:csvGeneratedPaths = $null
    Start-SmartM365SourceReceipt -ScriptPath $producer.FullName -SourceRootPath $latest
    Publish-Set 'baseline'
    Complete-SmartM365SourceReceipt -Status 'Completed' | Out-Null
    $baselineReceipt = Read-Json $receiptPath
    $baselineHashes = Get-LatestHashTable
    Check ($baselineReceipt.Status -eq 'Completed' -and @($baselineReceipt.Files).Count -eq 4) 'Baseline receipt not completed.'
    Check (-not (Read-Json $runPath).PSObject.Properties['StagedPublication']) 'An unstaged run declared staging.'
    Clear-Recorded

    # Resolve leaves paths unchanged without staging.
    Check ((Resolve-SmartM365SourcePublicationPath -Path (Join-Path $latest 'X.csv')) -eq (Join-Path $latest 'X.csv')) 'Resolve changed a path without an active staging.'

    # Staged run, interrupted before completion: an orphan from a crashed run is cleaned first.
    New-Item -ItemType Directory -Path (Join-Path $stagingProducer 'crashedrun') -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $stagingProducer 'crashedrun/Intune_ManagedApps.csv') -Value 'TenantKey,Name'
    $global:csvGeneratedPaths = $null
    Start-SmartM365SourceReceipt -ScriptPath $producer.FullName -SourceRootPath $latest
    Enable-SmartM365StagedSourcePublication
    Check (-not (Test-Path -LiteralPath (Join-Path $stagingProducer 'crashedrun'))) 'An orphaned staging folder survived the next run start.'
    $run = Read-Json $runPath
    Check ($run.Status -eq 'Collecting' -and $run.StagedPublication -eq $true) 'Run state does not declare staged collection.'
    $state = Get-SmartM365OrchestratorSourceRunState -SharedDataFolderPath $shared -Job ([pscustomobject]@{ ScriptPath = $producer.FullName })
    Check ($state.Known -and $state.StagedPublication -and $state.Status -eq 'Collecting') 'Orchestrator does not read the staged run state.'
    Check ((Resolve-SmartM365SourcePublicationPath -Path (Join-Path $latest 'X.txt')) -eq (Join-Path $latest 'X.txt')) 'A non-CSV path was staged.'
    Check ((Resolve-SmartM365SourcePublicationPath -Path (Join-Path $output 'X.csv')) -eq (Join-Path $output 'X.csv')) 'A CSV outside DATA-LAST was staged.'
    Publish-Set 'running'
    $stagedFolder = @(Get-ChildItem -LiteralPath $stagingProducer -Directory)
    Check ($stagedFolder.Count -eq 1 -and @(Get-ChildItem -LiteralPath $stagedFolder[0].FullName -Filter '*.csv').Count -eq 4) 'Run CSVs were not written to its staging folder.'
    $during = Get-LatestHashTable
    Check (@($files | Where-Object { $during[$_] -ne $baselineHashes[$_] }).Count -eq 0) 'A running staged run changed DATA-LAST.'
    Check ((Read-Json $receiptPath).RunId -eq $baselineReceipt.RunId) 'A running staged run changed the receipt.'
    Check ((Read-Json $runPath).Status -eq 'Collecting') 'A staged write switched the run to Publishing.'
    Check (@(Get-Recorded | Where-Object Kind -in @('Upload', 'History') | Where-Object { $_.Path -like "$latest*" -and $_.Path -notlike '*DATA-ALL*' }).Count -eq 0) 'A staged DATA-LAST copy was uploaded or historized before promotion.'
    Check (@(Get-Recorded | Where-Object { $_.Path -like '*.staging*' }).Count -eq 0) 'A staging path reached SharePoint or WeeklyHistory.'

    # Failed run: DATA-LAST, receipt and previous qualification are untouched; staging is removed.
    Complete-SmartM365SourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null
    $after = Get-LatestHashTable
    Check (@($files | Where-Object { $after[$_] -ne $baselineHashes[$_] }).Count -eq 0) 'A failed staged run changed DATA-LAST.'
    Check ((Read-Json $receiptPath).RunId -eq $baselineReceipt.RunId) 'A failed staged run replaced the receipt.'
    $run = Read-Json $runPath
    Check ($run.Status -eq 'Failed' -and -not $run.UnqualifiedPublication -and -not $run.PublicationStarted) 'A failed staged run disqualified the previous publication.'
    Check (@(Get-ChildItem -LiteralPath $stagingProducer -Force).Count -eq 0) 'A failed run left its staging folder.'
    Check (@(Get-Recorded | Where-Object { $_.Kind -eq 'Promoted' -or $_.Path -like "$latest\Intune*" }).Count -eq 0) 'A failed run promoted or uploaded DATA-LAST copies.'
    Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof (Read-Json $receiptPath)
    $script:checks++

    # Complete run: promotion under Publishing, receipt last, follow-ups after the receipt.
    Clear-Recorded
    $global:csvGeneratedPaths = $null
    Start-SmartM365SourceReceipt -ScriptPath $producer.FullName -SourceRootPath $latest
    Enable-SmartM365StagedSourcePublication
    Publish-Set 'complete'
    $stagedHashes = @{}
    foreach ($file in $files) { $stagedHashes[$file] = (Get-FileHash -LiteralPath (Join-Path (Get-ChildItem -LiteralPath $stagingProducer -Directory)[0].FullName $file) -Algorithm SHA256).Hash }
    Complete-SmartM365SourceReceipt -Status 'Completed' | Out-Null
    $receipt = Read-Json $receiptPath
    $final = Get-LatestHashTable
    Check ($receipt.Status -eq 'Completed' -and $receipt.RunId -ne $baselineReceipt.RunId) 'A complete staged run did not publish its receipt.'
    foreach ($file in $files) {
        $entry = @($receipt.Files | Where-Object File -eq $file)
        Check ($final[$file] -eq $stagedHashes[$file] -and $entry.Count -eq 1 -and $entry[0].SHA256 -eq $final[$file]) "Promoted CSV and receipt disagree: $file"
    }
    $promoted = @(Get-Recorded | Where-Object Kind -eq 'Promoted')
    Check ($promoted.Count -eq 4 -and @($promoted | Where-Object RunStatus -ne 'Publishing').Count -eq 0) 'A CSV was promoted outside the Publishing fence.'
    $run = Read-Json $runPath
    Check ($run.Status -eq 'Completed' -and -not $run.UnqualifiedPublication -and $run.StagedPublication) 'Completed staged run state is not qualified.'
    $uploads = @(Get-Recorded | Where-Object { $_.Kind -eq 'Upload' -and $_.Path -like "$latest*" })
    $history = @(Get-Recorded | Where-Object Kind -eq 'History')
    Check ($uploads.Count -eq 4 -and @($uploads | Where-Object { $_.Path -like '*.staging*' }).Count -eq 0) 'Promoted DATA-LAST copies were not uploaded by canonical path.'
    Check ($history.Count -eq 4 -and @($history | Where-Object { $_.Path -like '*.staging*' }).Count -eq 0) 'WeeklyHistory did not follow the promoted canonical copies.'
    $order = @(Get-Recorded | ForEach-Object Kind)
    Check ([array]::LastIndexOf($order, 'Promoted') -lt [array]::IndexOf($order, 'History')) 'WeeklyHistory ran before the promotion.'
    Check (@(Get-ChildItem -LiteralPath $stagingProducer -Force).Count -eq 0) 'A completed run left its staging folder.'
    Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof $receipt
    $script:checks++

    # Promotion failure: some CSVs replaced, receipt unchanged, publication disqualified (fail-closed).
    $previousReceipt = $receipt
    $global:csvGeneratedPaths = $null
    Start-SmartM365SourceReceipt -ScriptPath $producer.FullName -SourceRootPath $latest
    Enable-SmartM365StagedSourcePublication
    Publish-Set 'promotion-failure'
    $lockedFile = (@($files | Sort-Object))[-1]
    $lock = [IO.File]::Open((Join-Path $latest $lockedFile), [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]::None)
    try { Complete-SmartM365SourceReceipt -Status 'Completed' | Out-Null }
    finally { $lock.Dispose() }
    Check ((Read-Json $receiptPath).RunId -eq $previousReceipt.RunId) 'A failed promotion published a receipt.'
    $run = Read-Json $runPath
    Check ($run.Status -eq 'Failed' -and $run.UnqualifiedPublication -and $run.Error -match 'promotion') 'A failed promotion was not disqualified.'
    $rejected = $false
    try { Assert-SmartM365SourcePublication -ReceiptPath $receiptPath -Proof (Read-Json $receiptPath) } catch { $rejected = $true }
    Check $rejected 'Readers accepted a partially promoted set.'
    $state = Get-SmartM365OrchestratorSourceRunState -SharedDataFolderPath $shared -Job ([pscustomobject]@{ ScriptPath = $producer.FullName })
    Check ($state.Known -and $state.Status -eq 'Failed') 'Orchestrator does not see the failed promotion.'

    # A direct DATA-LAST write voids the staging declaration for the rest of the run.
    $global:csvGeneratedPaths = $null
    Start-SmartM365SourceReceipt -ScriptPath $producer.FullName -SourceRootPath $latest
    Enable-SmartM365StagedSourcePublication
    Write-SmartM365CsvAtomically -Data @([pscustomobject]@{ TenantKey = 'synthetic'; Name = 'direct' }) -Path (Join-Path $latest 'Intune_ManagedApps.csv') -NoTenantKey
    $run = Read-Json $runPath
    Check ($run.Status -eq 'Publishing' -and -not $run.PSObject.Properties['StagedPublication']) 'A direct write kept the staging declaration.'
    Complete-SmartM365SourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null

    # Staging cannot start after a direct publication and is a no-op without an active receipt.
    $global:csvGeneratedPaths = $null
    Start-SmartM365SourceReceipt -ScriptPath $producer.FullName -SourceRootPath $latest
    Write-SmartM365CsvAtomically -Data @([pscustomobject]@{ TenantKey = 'synthetic'; Name = 'direct' }) -Path (Join-Path $latest 'Intune_ManagedApps.csv') -NoTenantKey
    $refused = $false
    try { Enable-SmartM365StagedSourcePublication } catch { $refused = $true }
    Check $refused 'Staging was enabled after a direct DATA-LAST write.'
    Complete-SmartM365SourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null
    Enable-SmartM365StagedSourcePublication
    $script:checks++
}
finally {
    # Releases the producer lock if a check failed while a receipt was active.
    try { Complete-SmartM365SourceReceipt -Status 'Failed' -ErrorCount 1 | Out-Null } catch { Write-Verbose $_.Exception.Message }
    foreach ($name in $saved.Keys) {
        if ($saved[$name].Exists) { Set-Variable -Name $name -Scope Global -Value $saved[$name].Value }
        else { Remove-Variable -Name $name -Scope Global -ErrorAction SilentlyContinue }
    }
    if (Test-Path -LiteralPath $temporary) { Remove-Item -LiteralPath $temporary -Recurse -Force }
}
"PASS: $script:checks staged source publication checks; PowerShell $($PSVersionTable.PSVersion)."
