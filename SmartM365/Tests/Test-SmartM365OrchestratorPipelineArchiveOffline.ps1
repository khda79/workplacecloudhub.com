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


