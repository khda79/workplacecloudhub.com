<#
.SYNOPSIS
Offline tests of the weekly history JSON transition and changed-only SharePoint publication.

.VERSION
1.1
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot '../Modules/SmartM365.Core/SmartM365.JsonTransport.psd1') -Force
$transport = Get-Module SmartM365.JsonTransport
$policy = & $transport { (Get-Command Get-SmartM365JsonTransportPolicy).ScriptBlock }
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-JsonWeekly-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory $root
$script:passed = 0
function Check([bool]$Value, [string]$Message) { if (-not $Value) { throw $Message }; $script:passed++ }
function Reject([scriptblock]$Action, [string]$Message) { $failed=$false; try { & $Action | Out-Null } catch { $failed=$true }; Check $failed $Message }
function WriteLog { param($Message,$Level) }
function Get-SmartM365IsoWeekName { '2026-W39' }
function Get-SmartM365WeeklyHistoryFileName { param($Path) [IO.Path]::GetFileName($Path) }
function Copy-SmartM365FileAtomically { param($SourcePath,$DestinationPath) Copy-Item -LiteralPath $SourcePath -Destination $DestinationPath }
function Invoke-SmartM365SharePointCsvUpload { param($LocalFilePath) $script:uploads += $LocalFilePath; if ($script:failUploads) { return $null }; return 'receipt' }
try {
    foreach ($relative in @('../Modules/SmartM365.Core/SmartM365.Core.psm1','../Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psm1')) {
        $tokens=$null;$errors=$null
        $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot $relative),[ref]$tokens,[ref]$errors)
        if ($errors.Count) { throw ($errors | Out-String) }
        $definition=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Save-SmartM365WeeklyInventoryHistory'},$true)
        . ([scriptblock]::Create($definition.Extent.Text))
        $case = Join-Path $root ([guid]::NewGuid().ToString('N'))
        $history = Join-Path $case 'WeeklyHistory'
        $week = Join-Path $history '2026-W39'
        $oldWeek = Join-Path $history '2025-W01'
        $null = New-Item -ItemType Directory $week,$oldWeek -Force
        $source = Join-Path $case 'Inventory.csv'
        [IO.File]::WriteAllText($source,"Id`r`n42`r`n")
        Copy-Item $source $week
        Copy-Item $source $oldWeek
        foreach ($folder in @($week,$oldWeek)) {
            $doc=[ordered]@{Week=[IO.Path]::GetFileName($folder);HistoryLabel='Synthetic';HistoryRootPath=$history;Files=@('Inventory.csv');SnapshotCreatedAtUtc='2025-01-01T00:00:00Z';FileSnapshotCreatedAtUtc=@{};UpdatedAt='2025-01-01T00:00:00Z'}
            [IO.File]::WriteAllText((Join-Path $folder 'manifest.json'),($doc | ConvertTo-Json -Depth 7),[Text.UTF8Encoding]::new($true))
        }
        $legacy = Join-Path $week 'manifest.json'
        $oldLegacy = Join-Path $oldWeek 'manifest.json'
        $hash = (Get-FileHash $legacy).Hash
        $oldHash = (Get-FileHash $oldLegacy).Hash
        & $transport { function script:Get-SmartM365JsonTransportPolicy { @{Mode='Readers';QualifiedUncRoots=@()} } }
        $script:uploads=@()
        Save-SmartM365WeeklyInventoryHistory -SourceFiles $source -HistoryRootPath $history -HistoryLabel Synthetic -RetentionWeeks 0
        Check ((Get-FileHash $legacy).Hash -eq $hash) 'Reader rollout rewrote an unchanged historical manifest.'
        & $transport { function script:Get-SmartM365JsonTransportPolicy { @{Mode='JsonText';QualifiedUncRoots=@()} } }
        $script:uploads=@()
        Save-SmartM365WeeklyInventoryHistory -SourceFiles $source -HistoryRootPath $history -HistoryLabel Synthetic -RetentionWeeks 0 -UploadChangedFilesOnly
        Check ((Get-FileHash "$legacy.txt").Hash -eq $hash -and -not (Test-Path $legacy)) 'Current manifest bytes changed during migration.'
        Check ((Get-FileHash "$oldLegacy.txt").Hash -eq $oldHash -and -not (Test-Path $oldLegacy)) 'Older manifest was lost or recalculated.'
        Check ($script:uploads -contains "$oldLegacy.txt" -and $script:uploads -contains "$legacy.txt") 'Converted manifests absent from upload candidates.'
        Check (@($script:uploads | Where-Object { $_ -match '\.json$' }).Count -eq 0) 'Legacy manifest selected for upload.'
        $script:uploads=@()
        Save-SmartM365WeeklyInventoryHistory -SourceFiles $source -HistoryRootPath $history -HistoryLabel Synthetic -RetentionWeeks 0 -UploadChangedFilesOnly
        Check ($script:uploads -contains "$oldLegacy.txt") 'Restart after migration did not retry remote publication.'
        Check ((Get-FileHash "$legacy.txt").Hash -eq $hash) 'Idempotent execution rewrote historical metadata.'
        Copy-Item "$legacy.txt" $legacy
        [IO.File]::WriteAllText("$legacy.txt",'{')
        Reject { Save-SmartM365WeeklyInventoryHistory -SourceFiles $source -HistoryRootPath $history -HistoryLabel Synthetic -RetentionWeeks 1 } 'Invalid preferred manifest fell back or was rebuilt.'
        Check (Test-Path "$oldLegacy.txt") 'Retention ran after failed manifest validation.'
        Copy-Item $legacy "$legacy.txt" -Force
        $foreign=Get-Content "$legacy.txt" -Raw | ConvertFrom-Json
        $foreign.HistoryLabel='Foreign'
        [IO.File]::WriteAllText("$legacy.txt",($foreign | ConvertTo-Json -Depth 7))
        Reject { Resolve-SmartM365WeeklyManifestPaths -HistoryRootPath $history -HistoryLabel Synthetic } 'Divergent or foreign manifest accepted.'
        Check ((Get-FileHash $legacy).Hash -eq $hash) 'Failure destroyed the last valid legacy manifest.'
        # Changed-only publication recovers from a failed SharePoint upload through the pending marker.
        $pendingCase = Join-Path $root ([guid]::NewGuid().ToString('N'))
        $pendingHistory = Join-Path $pendingCase 'WeeklyHistory'
        $null = New-Item -ItemType Directory $pendingHistory -Force
        $pendingSource = Join-Path $pendingCase 'Inventory.csv'
        [IO.File]::WriteAllText($pendingSource, "Id`r`n7`r`n")
        $pendingMarker = Join-Path $pendingHistory '2026-W39\upload.pending'
        $pendingCsv = Join-Path $pendingHistory '2026-W39\Inventory.csv'
        Set-Variable -Name EnableSharePointUpload -Scope Global -Value $true
        try {
            $script:failUploads = $true; $script:uploads = @()
            Reject { Save-SmartM365WeeklyInventoryHistory -SourceFiles $pendingSource -HistoryRootPath $pendingHistory -HistoryLabel Synthetic -RetentionWeeks 0 -UploadChangedFilesOnly } 'Failed weekly publication was not reported.'
            Check ((Test-Path $pendingCsv) -and (Test-Path $pendingMarker)) 'Failed publication left no pending marker.'
            $script:failUploads = $false; $script:uploads = @()
            Save-SmartM365WeeklyInventoryHistory -SourceFiles $pendingSource -HistoryRootPath $pendingHistory -HistoryLabel Synthetic -RetentionWeeks 0 -UploadChangedFilesOnly
            Check ($script:uploads -contains $pendingCsv) 'Pending week was not republished.'
            Check (-not (Test-Path $pendingMarker)) 'Pending marker kept after a complete publication.'
            Check (@($script:uploads | Where-Object { $_ -like '*upload.pending' }).Count -eq 0) 'Pending marker selected for upload.'
            $script:uploads = @()
            Save-SmartM365WeeklyInventoryHistory -SourceFiles $pendingSource -HistoryRootPath $pendingHistory -HistoryLabel Synthetic -RetentionWeeks 0 -UploadChangedFilesOnly
            Check (-not ($script:uploads -contains $pendingCsv)) 'Unchanged week CSV uploaded again without a pending marker.'
        }
        finally { Set-Variable -Name EnableSharePointUpload -Scope Global -Value $false; $script:failUploads = $false }
        $disabledHistory = Join-Path (Join-Path $root ([guid]::NewGuid().ToString('N'))) 'WeeklyHistory'
        $null = New-Item -ItemType Directory $disabledHistory -Force
        Save-SmartM365WeeklyInventoryHistory -SourceFiles $pendingSource -HistoryRootPath $disabledHistory -HistoryLabel Synthetic -RetentionWeeks 0 -UploadChangedFilesOnly
        Check (-not (Test-Path (Join-Path $disabledHistory '2026-W39\upload.pending'))) 'Pending marker written while SharePoint publication is disabled.'
    }
    [pscustomobject]@{Passed=$script:passed;FixtureRoot=$root;Evidence='Synthetic; SharePoint upload mocked; no collector entry point'}
} finally { & $transport { param($original) Set-Item Function:script:Get-SmartM365JsonTransportPolicy -Value $original } $policy }
