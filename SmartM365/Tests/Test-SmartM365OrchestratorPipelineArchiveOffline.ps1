<#
.SYNOPSIS
Runs offline tests for pipeline request archiving, the terminal request cache, the JSON
unlinked-path parent cache, transcript, expired lease and mail HTML copy retention of the orchestrator.
.VERSION
1.0.0
#>
#Requires -Version 7.0
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$smartM365Root = Split-Path -Path $PSScriptRoot -Parent
$orchestratorRoot = Join-Path -Path $smartM365Root -ChildPath 'SmartInventory\Orchestrator'
$orchestratorPath = Join-Path -Path $orchestratorRoot -ChildPath 'SmartM365-Inventory-Orchestrator.ps1'
$pipelineModule = Import-Module -Name (Join-Path $orchestratorRoot 'SmartM365.Orchestrator.Pipeline.psm1') -Force -PassThru -ErrorAction Stop
$insightsModule = Import-Module -Name (Join-Path $orchestratorRoot 'SmartM365.Orchestrator.Insights.psm1') -Force -PassThru -ErrorAction Stop
$transportModule = Get-Module -Name SmartM365.JsonTransport
$checks = 0

function Assert-True {
    param([Parameter(Mandatory)][bool]$Condition, [Parameter(Mandatory)][string]$Message)
    if (-not $Condition) { throw $Message }
    $script:checks++
}

function Write-TestJson {
    param([string]$Path, $Document)
    $folder = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    [IO.File]::WriteAllText($Path, ($Document | ConvertTo-Json -Depth 20), [Text.UTF8Encoding]::new($false))
}

function New-TestBatch {
    # Synthetic request folder with explicit creation time and job statuses.
    param([string]$Root, [datetime]$CreatedUtc, [hashtable]$Statuses, [datetime]$UpdatedUtc, [string]$Folder = 'PipelineRuns')
    $batchId = '{0}-{1}' -f $CreatedUtc.ToString('yyyyMMddTHHmmssfffZ'), [guid]::NewGuid().ToString('N').Substring(0, 8)
    $batchPath = Join-Path (Join-Path $Root $Folder) $batchId
    $names = @($Statuses.Keys | Sort-Object)
    Write-TestJson -Path (Join-Path $batchPath 'request.json') -Document ([pscustomobject][ordered]@{
        SchemaVersion = 1; BatchId = $batchId; Tenant = 'test'; Pipeline = 'Jobs'
        CreatedAtUtc = $CreatedUtc.ToString('o'); OccurrenceUtc = $CreatedUtc.ToString('o'); ManifestHash = ''
        SelectedJobs = @($names | ForEach-Object { [pscustomobject]@{ Name = $_ } })
    })
    foreach ($name in $names) {
        Write-TestJson -Path (Join-Path (Join-Path $batchPath 'Jobs') ($name + '.json')) -Document ([pscustomobject][ordered]@{
            SchemaVersion = 1; BatchId = $batchId; JobName = $name; Status = $Statuses[$name]; Attempt = 1; OwnerServer = 'SERVER-A'
            CreatedAtUtc = $CreatedUtc.ToString('o'); UpdatedAtUtc = $UpdatedUtc.ToString('o'); Detail = ''; NotBeforeUtc = ''
        })
    }
    [pscustomobject]@{ BatchId = $batchId; Path = $batchPath }
}

function Get-ExtractedFunctions {
    param([string[]]$Names, [string[]]$Extra = @())
    $tokens = $null; $parseErrors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($orchestratorPath, [ref]$tokens, [ref]$parseErrors)
    if (@($parseErrors).Count -gt 0) { throw 'Orchestrator source failed parsing.' }
    $definitions = foreach ($name in $Names) {
        $node = $ast.Find({ param($candidate) $candidate -is [Management.Automation.Language.FunctionDefinitionAst] -and $candidate.Name -eq $name }, $true)
        if (-not $node) { throw "Missing orchestrator function: $name" }
        $node.Extent.Text
    }
    New-Module -ScriptBlock ([scriptblock]::Create((@($definitions) + $Extra) -join [Environment]::NewLine))
}

$temporaryRoot = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath ('SmartM365-PipelineArchive-' + [guid]::NewGuid().ToString('N'))
try {
    $root = Join-Path $temporaryRoot 'DATA-ALL\Orchestrator'
    New-Item -ItemType Directory -Path (Join-Path $root 'PipelineRuns') -Force | Out-Null
    $now = [datetime]::UtcNow

    # 1. Terminal cache: a terminal request is read once, then skipped by the active scan.
    $done = New-TestBatch -Root $root -CreatedUtc $now.AddDays(-20) -UpdatedUtc $now.AddDays(-19) -Statuses @{ JobA = 'Success'; JobB = 'Failed' }
    $running = New-TestBatch -Root $root -CreatedUtc $now.AddDays(-15) -UpdatedUtc $now.AddDays(-15) -Statuses @{ JobA = 'Running' }
    $progress = [ref]0
    $active = @(Get-SmartM365OrchestratorActivePipelineRuns -SharedDataFolderPath $root -OnProgress { $progress.Value++ })
    Assert-True ($active.Count -eq 1 -and $active[0].BatchId -eq $running.BatchId) 'The active scan did not return only the running request.'
    Assert-True ($progress.Value -eq 2) "The active scan did not report progress for each request read ($($progress.Value))."
    Assert-True ((& $pipelineModule { $script:PipelineTerminalBatches.Count }) -eq 1) 'The terminal request was not cached.'
    $progress.Value = 0
    $active = @(Get-SmartM365OrchestratorActivePipelineRuns -SharedDataFolderPath $root -OnProgress { $progress.Value++ })
    Assert-True ($active.Count -eq 1 -and $progress.Value -eq 1) 'The cached terminal request was read again.'

    # A terminal request with a missing status file may be a transient share error: never cached.
    $missing = New-TestBatch -Root $root -CreatedUtc $now.AddDays(-18) -UpdatedUtc $now.AddDays(-18) -Statuses @{ JobA = 'Success'; JobB = 'Success' }
    Remove-Item -LiteralPath (Join-Path $missing.Path 'Jobs\JobB.json') -Force
    $null = Get-SmartM365OrchestratorActivePipelineRuns -SharedDataFolderPath $root
    Assert-True ((& $pipelineModule { $script:PipelineTerminalBatches.Count }) -eq 1) 'A request with a missing job status was cached as terminal.'

    # 2. Archiving: only terminal requests whose last activity is older than the delay move.
    $recent = New-TestBatch -Root $root -CreatedUtc $now.AddDays(-12) -UpdatedUtc $now.AddDays(-1) -Statuses @{ JobA = 'Success' }
    $locked = New-TestBatch -Root $root -CreatedUtc $now.AddDays(-11) -UpdatedUtc $now.AddDays(-11) -Statuses @{ JobA = 'Success' }
    Set-Content -LiteralPath (Join-Path $locked.Path 'Control.lock') -Value 'busy'
    $young = New-TestBatch -Root $root -CreatedUtc $now.AddDays(-2) -UpdatedUtc $now.AddDays(-2) -Statuses @{ JobA = 'Success' }

    $expired = Move-SmartM365OrchestratorPipelineRunsToArchive -SharedDataFolderPath $root -ArchiveAfterDays 7 -RetentionDays 90 -DeadlineUtc $now.AddMinutes(-1)
    Assert-True ($expired.BudgetExhausted -and $expired.Archived -eq 0) 'An exhausted time budget still archived requests.'

    $refused = $false
    try { $null = Move-SmartM365OrchestratorPipelineRunsToArchive -SharedDataFolderPath $root -ArchiveAfterDays 7 -RetentionDays 7 } catch { $refused = $true }
    Assert-True $refused 'A retention not longer than the archive delay was accepted.'

    # The submission lock serializes archiving with submissions and cancellations.
    $paths = Get-SmartM365OrchestratorPipelinePaths -SharedDataFolderPath $root
    $held = [IO.File]::Open($paths.SubmissionLockPath, 'CreateNew', 'ReadWrite', 'None')
    try { $busy = Move-SmartM365OrchestratorPipelineRunsToArchive -SharedDataFolderPath $root -ArchiveAfterDays 7 -RetentionDays 90 -LockTimeoutSeconds 1 }
    finally { $held.Dispose(); Remove-Item -LiteralPath $paths.SubmissionLockPath -Force }
    Assert-True ($busy.Archived -eq 0 -and @($busy.Messages | Where-Object { $_ -like '*lock busy*' }).Count -eq 1) ("Archiving did not wait for the submission lock: " + ($busy | ConvertTo-Json -Depth 3 -Compress))

    $progress.Value = 0
    $result = Move-SmartM365OrchestratorPipelineRunsToArchive -SharedDataFolderPath $root -ArchiveAfterDays 7 -RetentionDays 90 -OnProgress { $progress.Value++ }
    Assert-True ($result.Archived -eq 1 -and (Test-Path -LiteralPath (Join-Path $paths.ArchiveRootPath $done.BatchId)) -and -not (Test-Path -LiteralPath $done.Path)) 'The old terminal request was not archived.'
    Assert-True ($result.KeptActive -eq 3 -and $result.KeptRecent -eq 1) "Unexpected kept counts: active=$($result.KeptActive); recent=$($result.KeptRecent)."
    foreach ($kept in @($running, $missing, $recent, $locked, $young)) { Assert-True (Test-Path -LiteralPath $kept.Path) "A request that must stay live was moved: $($kept.BatchId)." }
    Assert-True ($progress.Value -ge 5) 'Archiving did not report progress.'
    Assert-True ((& $pipelineModule { $script:PipelineTerminalBatches.Count }) -eq 0) 'The archived request stayed in the terminal cache.'
    $active = @(Get-SmartM365OrchestratorActivePipelineRuns -SharedDataFolderPath $root)
    Assert-True ($active.Count -eq 1 -and $active[0].BatchId -eq $running.BatchId) 'The active scan changed after archiving.'

    # A request already present in the archive is left in place and reported.
    Remove-Item -LiteralPath (Join-Path $locked.Path 'Control.lock') -Force
    New-Item -ItemType Directory -Path (Join-Path $paths.ArchiveRootPath $locked.BatchId) -Force | Out-Null
    $duplicate = Move-SmartM365OrchestratorPipelineRunsToArchive -SharedDataFolderPath $root -ArchiveAfterDays 7 -RetentionDays 90
    Assert-True ($duplicate.Archived -eq 0 -and $duplicate.Failures -eq 1 -and (Test-Path -LiteralPath $locked.Path)) 'A request whose archive name already exists was moved or not reported.'
    Remove-Item -LiteralPath (Join-Path $paths.ArchiveRootPath $locked.BatchId) -Recurse -Force

    # 3. The GUI history lists live and archived requests, newest first, the live copy winning.
    $history = @(Get-SmartM365OrchestratorRecentPipelineRuns -SharedDataFolderPath $root -Count 20)
    Assert-True ($history.Count -eq 6) "The history did not list live and archived requests ($($history.Count))."
    Assert-True ((@($history.BatchId) -join '|') -eq (@($history.BatchId | Sort-Object -Descending) -join '|')) 'The history is not sorted newest first.'
    $archivedRow = @($history | Where-Object BatchId -eq $done.BatchId)
    Assert-True ($archivedRow.Count -eq 1 -and $archivedRow[0].Status -eq 'Failed' -and $archivedRow[0].Jobs -eq 2) 'The archived request is not shown with its final status.'
    Assert-True (@(Get-SmartM365OrchestratorRecentPipelineRuns -SharedDataFolderPath $root -Count 2).Count -eq 2) 'The history count limit is not applied.'

    # 4. Purge of archived requests older than the retention, never through a link.
    $old = New-TestBatch -Root $root -Folder 'PipelineRuns-Archive' -CreatedUtc $now.AddDays(-120) -UpdatedUtc $now.AddDays(-119) -Statuses @{ JobA = 'Success' }
    $linked = New-TestBatch -Root $root -Folder 'PipelineRuns-Archive' -CreatedUtc $now.AddDays(-110) -UpdatedUtc $now.AddDays(-109) -Statuses @{ JobA = 'Success' }
    $outside = Join-Path $temporaryRoot 'Outside'
    New-Item -ItemType Directory -Path $outside -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $outside 'sentinel.txt') -Value 'keep'
    New-Item -ItemType Junction -Path (Join-Path $linked.Path 'Link') -Target $outside | Out-Null
    $purge = Move-SmartM365OrchestratorPipelineRunsToArchive -SharedDataFolderPath $root -ArchiveAfterDays 7 -RetentionDays 90
    Assert-True ($purge.Purged -eq 1 -and -not (Test-Path -LiteralPath $old.Path)) 'The expired archived request was not purged.'
    Assert-True ((Test-Path -LiteralPath $linked.Path) -and (Test-Path -LiteralPath (Join-Path $outside 'sentinel.txt')) -and $purge.Failures -eq 1) 'An archived folder containing a link was purged or followed.'
    Assert-True (Test-Path -LiteralPath (Join-Path $paths.ArchiveRootPath $done.BatchId)) 'An archived request inside the retention was purged.'
    $keepAll = Move-SmartM365OrchestratorPipelineRunsToArchive -SharedDataFolderPath $root -ArchiveAfterDays 7 -RetentionDays 0
    Assert-True ($keepAll.Purged -eq 0) 'RetentionDays 0 purged archived requests.'
    (Get-Item -LiteralPath (Join-Path $linked.Path 'Link')).Delete()

    # 5. Unlinked-path cache: parents are trusted for a short time, the target never is.
    $chain = Join-Path $temporaryRoot 'Chain\Level1\Level2'
    New-Item -ItemType Directory -Path $chain -Force | Out-Null
    & $transportModule { $script:UnlinkedDirectoryCache.Clear() }
    & $transportModule { param($Path) Assert-SmartM365JsonUnlinkedPath $Path } (Join-Path $chain 'state.json')
    $cached = & $transportModule { @($script:UnlinkedDirectoryCache.Keys) }
    Assert-True (@($cached | Where-Object { $_ -like '*\Chain\Level1\Level2' }).Count -eq 1 -and @($cached | Where-Object { $_ -like '*\Chain' }).Count -eq 1) 'Verified parent directories were not cached.'
    $leafLink = Join-Path $chain 'linked.json'
    New-Item -ItemType Junction -Path $leafLink -Target $outside | Out-Null
    $leafRefused = $false
    try { & $transportModule { param($Path) Assert-SmartM365JsonUnlinkedPath $Path } $leafLink } catch { $leafRefused = $_.Exception.Message -like 'Linked path*' }
    Assert-True $leafRefused 'A linked target was accepted because its parents were cached.'
    (Get-Item -LiteralPath $leafLink).Delete()
    # A parent replaced by a link is detected once the cache entry expires (here: disabled).
    $level2Moved = Join-Path $temporaryRoot 'Chain\Level1\Level2-real'
    Move-Item -LiteralPath $chain -Destination $level2Moved
    New-Item -ItemType Junction -Path $chain -Target $level2Moved | Out-Null
    $parentRefused = $false
    try {
        & $transportModule { $script:UnlinkedDirectoryCacheSeconds = 0 }
        & $transportModule { param($Path) Assert-SmartM365JsonUnlinkedPath $Path } (Join-Path $chain 'state.json')
    }
    catch { $parentRefused = $_.Exception.Message -like 'Linked path*' }
    finally { & $transportModule { $script:UnlinkedDirectoryCacheSeconds = 300; $script:UnlinkedDirectoryCache.Clear() } }
    Assert-True $parentRefused 'A parent replaced by a link was accepted after the cache expired.'
    (Get-Item -LiteralPath $chain).Delete()

    # 6. Hourly archiving in the resident: one server at a time, heartbeat kept fresh.
    $archiveModule = Get-ExtractedFunctions -Names @('Invoke-OrchestratorPipelineRunArchive', 'Enter-OrchestratorSharePointMirrorLock', 'Exit-OrchestratorSharePointMirrorLock') -Extra @(
        'function Update-OrchestratorHeartbeatDuringLongOperation { $script:Heartbeats++ }',
        'function Write-OrchestratorLog { param([string]$Message, [string]$Level = ''INFO'') $script:Logs.Add("$Level|$Message") }',
        'function Write-OrchestratorRuntimeUpdateWarning { param($Key, $Message, $Now, $Level = ''WARN'') if (-not $script:Warned.Contains($Key)) { [void]$script:Warned.Add($Key); $script:Logs.Add("$Level|$Message") } }'
    )
    $back = New-TestBatch -Root $root -CreatedUtc $now.AddDays(-30) -UpdatedUtc $now.AddDays(-30) -Statuses @{ JobA = 'Success' }
    $archiveRun = & $archiveModule {
        param($Root)
        $script:Settings = [pscustomobject]@{
            SharedDataFolderPath = $Root; PipelineRunsFolderPath = (Join-Path $Root 'PipelineRuns')
            PipelineRunArchiveDays = 7; PipelineRunArchiveRetentionDays = 90; ElectionClaimRetentionMaxSecondsPerRun = 120
        }
        $script:LastPipelineRunArchiveRun = [datetime]::MinValue
        $script:Heartbeats = 0; $script:Logs = [Collections.Generic.List[string]]::new(); $script:Warned = [Collections.Generic.HashSet[string]]::new()
        $lockPath = Join-Path $Root 'Orchestrator-PipelineRunArchive.lock'
        $peer = [IO.File]::Open($lockPath, 'CreateNew', 'ReadWrite', 'None')
        try { Invoke-OrchestratorPipelineRunArchive -Now (Get-Date) } finally { $peer.Dispose(); Remove-Item -LiteralPath $lockPath -Force }
        $blockedLogs = $script:Logs.Count
        $script:LastPipelineRunArchiveRun = [datetime]::MinValue
        Invoke-OrchestratorPipelineRunArchive -Now (Get-Date)
        $firstLogs = @($script:Logs.ToArray())
        Invoke-OrchestratorPipelineRunArchive -Now (Get-Date).AddMinutes(10)
        [pscustomobject]@{ BlockedLogs = $blockedLogs; Logs = $firstLogs; LogsAfterGate = $script:Logs.Count; Heartbeats = $script:Heartbeats; LockLeft = Test-Path -LiteralPath $lockPath }
    } $root
    Assert-True ($archiveRun.BlockedLogs -eq 0) 'Archiving ran while another server held the archive lock.'
    Assert-True (-not (Test-Path -LiteralPath $back.Path) -and (Test-Path -LiteralPath (Join-Path $paths.ArchiveRootPath $back.BatchId))) 'The resident did not archive the old request.'
    Assert-True (@($archiveRun.Logs | Where-Object { $_ -like 'INFO|Pipeline request archiving (after 7 days, purge after 90 days): archived=1;*' }).Count -eq 1) "The archiving summary was not logged: $($archiveRun.Logs -join ' / ')"
    Assert-True ($archiveRun.LogsAfterGate -eq $archiveRun.Logs.Count -and $archiveRun.Heartbeats -ge 1 -and -not $archiveRun.LockLeft) 'The hourly gate, heartbeat or lock release failed.'

    # 7. Transcript retention and expired concurrency lease archives.
    $logRoot = Join-Path $temporaryRoot 'Logs'
    $transcripts = Join-Path $logRoot 'Transcripts'
    New-Item -ItemType Directory -Path $transcripts -Force | Out-Null
    $oldTranscript = Join-Path $transcripts 'SmartM365-Inventory-Orchestrator_SERVER-A_20260901-010000_100.transcript.log'
    $newTranscript = Join-Path $transcripts 'SmartM365-Inventory-Orchestrator_SERVER-A_20261029-010000_200.transcript.log'
    $other = Join-Path $transcripts 'notes.log'
    foreach ($file in @($oldTranscript, $newTranscript, $other)) { Set-Content -LiteralPath $file -Value 'x' }
    foreach ($file in @($oldTranscript, $other)) { (Get-Item -LiteralPath $file).LastWriteTime = (Get-Date).AddDays(-30) }
    $retentionModule = Get-ExtractedFunctions -Names @('Invoke-RetentionCleanup') -Extra @(
        '$ScriptName = ''SmartM365-Inventory-Orchestrator''',
        'function Update-OrchestratorHeartbeatDuringLongOperation { }',
        'function Write-OrchestratorLog { param([string]$Message, [string]$Level) }'
    )
    & $retentionModule {
        param($LogRoot)
        $script:Settings = [pscustomobject]@{
            OrchestratorLogFolderPath = $LogRoot; OrchestratorLogRetentionDays = 30; JobLogFolderPath = (Join-Path $LogRoot 'Jobs'); JobLogRetentionDays = 30
            JobRunsFolderPath = (Join-Path $LogRoot 'JobRuns'); JobRunsCsvRetentionDays = 90
            OrchestratorTranscriptFolderPath = (Join-Path $LogRoot 'Transcripts'); OrchestratorTranscriptRetentionDays = 14
        }
        Invoke-RetentionCleanup
    } $logRoot
    Assert-True (-not (Test-Path -LiteralPath $oldTranscript) -and (Test-Path -LiteralPath $newTranscript) -and (Test-Path -LiteralPath $other)) 'Transcript retention removed the wrong files.'

    $claims = Join-Path $root 'Election\Claims'
    $leases = Join-Path $root 'Election\Concurrency'
    New-Item -ItemType Directory -Path $claims, $leases -Force | Out-Null
    $liveLease = Join-Path $leases 'SharedRuntime.json.txt'
    $oldStale = Join-Path $leases ('SharedRuntime.json.stale.{0}.json.txt' -f ('a' * 32))
    $recentStale = Join-Path $leases ('SharedRuntime.json.stale.{0}.json.txt' -f ('b' * 32))
    foreach ($file in @($liveLease, $oldStale, $recentStale)) { Set-Content -LiteralPath $file -Value '{}' }
    foreach ($file in @($liveLease, $oldStale)) { (Get-Item -LiteralPath $file).LastWriteTime = (Get-Date).AddDays(-40) }
    $claimModule = Get-ExtractedFunctions -Names @('Invoke-OrchestratorClaimRetention', 'Get-OrchestratorClaimOccurrenceUtc', 'Enter-OrchestratorSharePointMirrorLock', 'Exit-OrchestratorSharePointMirrorLock') -Extra @(
        'function Update-OrchestratorHeartbeatDuringLongOperation { }',
        'function Write-OrchestratorLog { param([string]$Message, [string]$Level = ''INFO'') $script:Logs.Add("$Level|$Message") }'
    )
    $claimLogs = & $claimModule {
        param($Root)
        $script:Settings = [pscustomobject]@{
            SharedDataFolderPath = $Root; ElectionClaimsPath = (Join-Path $Root 'Election\Claims'); ConcurrencyLeasesPath = (Join-Path $Root 'Election\Concurrency')
            ElectionClaimRetentionDays = 14; ElectionClaimRetentionMaxSecondsPerRun = 120
        }
        $script:LastClaimRetentionRun = [datetime]::MinValue
        $script:Logs = [Collections.Generic.List[string]]::new()
        Invoke-OrchestratorClaimRetention -Now (Get-Date)
        @($script:Logs.ToArray())
    } $root
    Assert-True (-not (Test-Path -LiteralPath $oldStale) -and (Test-Path -LiteralPath $recentStale) -and (Test-Path -LiteralPath $liveLease)) 'Expired lease archive retention removed the wrong files.'
    Assert-True (@($claimLogs | Where-Object { $_ -like '*expired lease archives removed=1;*' }).Count -eq 1) "Expired lease archive removal was not logged: $($claimLogs -join ' / ')"

    # 8. Mail HTML copies older than 24 h are removed from every script folder of LOG-ALL;
    # recent copies, other extensions and linked folders are kept.
    $logAll = Join-Path $temporaryRoot 'LOG-ALL'
    $oldMails = @(
        (Join-Path $logAll 'SmartM365-Orchestrator\SERVER-A\2026-10-01_08-00-00_[SMART 365] - test - Orchestrator.htm'),
        (Join-Path $logAll 'SmartM365-WorkplaceEvidence-Prepare\2026-10-02_09-00-00_[SMART 365] - test - Evidence.htm')
    )
    $recentMail = Join-Path $logAll 'SmartM365-WorkplaceEvidence-Prepare\2026-10-29_09-00-00_[SMART 365] - test - Evidence.htm'
    $oldHtml = Join-Path $logAll 'SmartM365-Reports\report.html'
    $oldLog = Join-Path $logAll 'SmartM365-Reports\run.log'
    $linkedTarget = Join-Path $temporaryRoot 'LinkedLogs'
    $linkedMail = Join-Path $linkedTarget 'old.htm'
    foreach ($file in @($oldMails + $recentMail, $oldHtml, $oldLog, $linkedMail)) {
        New-Item -ItemType Directory -Path (Split-Path $file -Parent) -Force | Out-Null
        Set-Content -LiteralPath $file -Value '<html></html>'
    }
    foreach ($file in @($oldMails + $oldHtml, $oldLog, $linkedMail)) { (Get-Item -LiteralPath $file).LastWriteTime = (Get-Date).AddHours(-30) }
    (Get-Item -LiteralPath $recentMail).LastWriteTime = (Get-Date).AddHours(-20)
    New-Item -ItemType Junction -Path (Join-Path $logAll 'Linked') -Target $linkedTarget | Out-Null
    $mailModule = Get-ExtractedFunctions -Names @('Invoke-OrchestratorMailHtmlRetention', 'Enter-OrchestratorSharePointMirrorLock', 'Exit-OrchestratorSharePointMirrorLock') -Extra @(
        'function Update-OrchestratorHeartbeatDuringLongOperation { $script:Heartbeats++ }',
        'function Write-OrchestratorLog { param([string]$Message, [string]$Level = ''INFO'') $script:Logs.Add("$Level|$Message") }',
        'function Write-OrchestratorRuntimeUpdateWarning { param($Key, $Message, $Now, $Level = ''WARN'') $script:Logs.Add("$Level|$Message") }'
    )
    $mailRun = & $mailModule {
        param($Root, $LogAll)
        $script:Settings = [pscustomobject]@{ SharedDataFolderPath = $Root; LogAllRootPath = $LogAll; MailHtmlCopyRetentionHours = 24; ElectionClaimRetentionMaxSecondsPerRun = 120 }
        $script:LastMailHtmlRetentionRun = [datetime]::MinValue
        $script:Heartbeats = 0; $script:Logs = [Collections.Generic.List[string]]::new()
        Invoke-OrchestratorMailHtmlRetention -Now (Get-Date)
        $first = @($script:Logs.ToArray())
        Invoke-OrchestratorMailHtmlRetention -Now (Get-Date).AddMinutes(5)
        [pscustomobject]@{ Logs = $first; LogsAfterGate = $script:Logs.Count; Heartbeats = $script:Heartbeats; LockLeft = Test-Path -LiteralPath (Join-Path $Root 'Orchestrator-MailHtmlRetention.lock') }
    } $root $logAll
    Assert-True (@($oldMails | Where-Object { Test-Path -LiteralPath $_ }).Count -eq 0) 'Mail HTML copies older than 24 hours were not removed from every script folder.'
    Assert-True ((Test-Path -LiteralPath $recentMail) -and (Test-Path -LiteralPath $oldHtml) -and (Test-Path -LiteralPath $oldLog)) 'Mail HTML retention removed a recent copy or a non-.htm file.'
    Assert-True (Test-Path -LiteralPath $linkedMail) 'Mail HTML retention followed a link.'
    Assert-True (@($mailRun.Logs | Where-Object { $_ -like 'INFO|Mail HTML copy retention (24 h, LOG-ALL): removed=2;*' }).Count -eq 1) "Mail HTML retention summary missing: $($mailRun.Logs -join ' / ')"
    Assert-True ($mailRun.LogsAfterGate -eq $mailRun.Logs.Count -and $mailRun.Heartbeats -ge 1 -and -not $mailRun.LockLeft) 'Mail HTML retention hourly gate, heartbeat or lock release failed.'

    # 9. Resident wiring: transcript, archive pass, heartbeat during the request scan.
    $source = Get-Content -LiteralPath $orchestratorPath -Raw
    foreach ($hook in @(
        'Start-SmartM365RunTranscript -Path $transcriptPath -UseMinimalHeader',
        'Invoke-OrchestratorPipelineRunArchive -Now $now',
        'Invoke-OrchestratorMailHtmlRetention -Now $now',
        "'MailHtmlCopyRetentionHours' -DefaultValue 24",
        '-OnProgress { Update-OrchestratorHeartbeatDuringLongOperation }',
        "'PipelineRunArchiveDays' -DefaultValue 2",
        "'PipelineRunArchiveRetentionDays' -DefaultValue 7",
        "'OrchestratorTranscriptRetentionDays' -DefaultValue 14"
    )) {
        Assert-True ($source.Contains($hook)) "Orchestrator hook is missing: $hook"
    }
}
finally {
    if (Test-Path -LiteralPath $temporaryRoot) {
        # Remove junctions first so the recursive cleanup never follows them.
        Get-ChildItem -LiteralPath $temporaryRoot -Recurse -Force -Attributes ReparsePoint -ErrorAction SilentlyContinue | ForEach-Object { $_.Delete() }
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}

Write-Host ("[{0}] PASS: {1} offline pipeline archive, terminal cache, unlinked-path cache and retention checks. Synthetic data only." -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $checks) -ForegroundColor Green



# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA38H6MrbdMgQeA
# VMDYTovZOvbZjyFlLlEPKUzZm6tN16CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIC4+cVFXO7w0+dQsR1VEI5nJNRQf5kSGnZiWkmKP8Ze1MA0GCSqG
# SIb3DQEBAQUABIIBgCuhedEUyESB6Mip5IA7bKs3b9rDTo8VKJeJBmEYOrAKpFRg
# sQAZhAJkO3/YlGoTsrDZxT4Xx+rPivPjJg2bJhE/Vt7uJtBf4Mg03OUZ8kqN/CpL
# 0uRdhNbXQso0asz1mELsaihnCqYVI9zzBZS48qSBBxwHCaWSb13OzQRMjrhiOF6y
# XfS3FGGk5x5LsPZLo6TQmgdaovJaUDVlhgx5uZRBseArPVt744jwvnXahvvz1BjJ
# gmP7XUOSAyv12mCKaIyl4zYoFq6WYypT82D4rndooewCLHvmKoRfpNSD5VHO6Ayr
# OG6YfX+pFaFFuGI5YexOKmkStrS2vWwldGu7AOZXUAvmKgDPqfouafBSwHuk48sF
# u5aEoRHDI7zo2c4CHBTp0ppNlCuRhIKU9+IQ7HVTYkBVzVRhDT6pAS7CyapULADg
# PwsJkvNDokKJrTMlsWs5UiFbjzmOeWrglnJp/v/ufjbtT/F/wsB7JMCjOARfNnuG
# Hm9CBFaAWudkc2jNzKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkxODIz
# MzhaMC8GCSqGSIb3DQEJBDEiBCA9IVu2idu+UkVSEPiOD6JbX4Flt4vWwRptnDRa
# UFh+nzANBgkqhkiG9w0BAQEFAASCAgBSttde+vBuQhG9uirRV8AlEkuiZJb01wk6
# DyGSdjL2q9nBEtio9pbgjBplJ0uE8oznASbNM8+XjokuZd43aOiRQ8f/bFOhyqL8
# hNTSg4pz2PxO2ZneCq4nJyFTNMNdDJac9WfR4AdCCIcbzKb3iCCMTxF1GmK0Ju0A
# aDG7jS8jA9kZDwnDtqu0AzlzYTDGInfBXuResOtMnuVWiRIjDN7HrqbBxDJCVaPD
# DlFiF2Z10HFxKDUXjZuucuM0v7siylMNK0yIQplisLuBKLRFnbzJShsffWd4ivg4
# ggfOvvdZ/c6OUFFxNJxwVl5Xe8Di8dCf2J/EcsN8hZ7GluMwOrpNzKMaWPjJXgRW
# 4rzfAaRPxwCHgoRRFegqo64BHfJ3s6S3uMvLytfpAur9sSVHYQxkhgrsLLFZ8KnM
# YOkEcb0FynchGc26SfkTx7cPiNoIOVz8TOSEXNuhNWkmh1CoyNqNKSiXOD81+pSj
# c3/zVSHeH7DHNL1J45gCa+QpV2xCDHD5xXBsFcpF0xnuMAWN/kC4jRDQ6n95gWX7
# BEeY7kweBYOgXbBrBWLwipoDLp2hyuV72XUuwR8wKAWwSPBceZP0IM2LhqJTl/5F
# RHUrm/3MHlNC1u4ZuUtFFKEtFGtyn3AKEkEayWcGYD82hjagj8cgens/QO8LCa19
# E/jZbSaEyA==
# SIG # End signature block
