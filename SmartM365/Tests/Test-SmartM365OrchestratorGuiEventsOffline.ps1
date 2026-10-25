#Requires -Version 7.0
<#
.SYNOPSIS
Checks the GUI Events tab data layer on synthetic resident orchestrator logs, offline.
.DESCRIPTION
Builds fictional daily logs for three servers, then verifies parsing, categories, period
and display filters, repeat grouping, incremental cached reads, shared reads of a file
held open by a writer, read-only behaviour, and the background GUI refresh contract.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
$folder = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
$insightsModule = (Resolve-Path (Join-Path $folder 'SmartM365.Orchestrator.Insights.psm1')).Path
$guiPath = Join-Path $folder 'SmartM365-Inventory-Orchestrator-GUI.ps1'
Import-Module $insightsModule -Force
Import-Module (Join-Path $folder 'SmartM365.Orchestrator.GuiWorker.psm1') -Force
$script:Checks = 0
function Assert-Events {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-GuiEvents-' + [guid]::NewGuid().ToString('N'))
$logRoot = Join-Path $root 'LOG-ALL/SmartM365-Orchestrator'
$now = [datetime]'2026-10-25 12:00:00'
$utf8 = [Text.UTF8Encoding]::new($false)
function Get-LogPath { param([string]$Server, [datetime]$Day) Join-Path (Join-Path $logRoot $Server) ('SmartM365-Inventory-Orchestrator_{0}_{1}.log' -f $Server, $Day.ToString('yyyyMMdd')) }
function Write-SyntheticLog {
    param([string]$Path, [string[]]$Lines, [switch]$Bom, [string]$Tail = '')
    $null = New-Item -ItemType Directory -Path (Split-Path $Path -Parent) -Force
    $text = ($Lines -join "`r`n") + "`r`n" + $Tail
    $bytes = $utf8.GetBytes($text)
    if ($Bom) { $bytes = [byte[]](0xEF, 0xBB, 0xBF) + $bytes }
    [IO.File]::WriteAllBytes($Path, $bytes)
}

try {
    $workerAToday = Get-LogPath 'WorkerA' $now
    $workerAYesterday = Get-LogPath 'WorkerA' $now.AddDays(-1)
    $workerBToday = Get-LogPath 'WorkerB' $now
    Write-SyntheticLog -Path $workerAYesterday -Lines @(
        '[2026-10-24 09:00:00][ERROR] Synthetic error older than one day.'
        '[2026-10-24 13:00:00][WARN] Failed to write heartbeat: synthetic share timeout.'
        '[2026-10-24 23:59:00][INFO] SmartM365-Inventory-Orchestrator exiting normally (exit code 0).'
    )
    Write-SyntheticLog -Path $workerAToday -Bom -Lines @(
        '[2026-10-25 00:00:05][INFO] SmartM365-Inventory-Orchestrator v9.9.1 starting (PID 101, tenant demo, MaxConcurrency 2, MaxLifetimeHours 24).'
        '[2026-10-25 00:01:00][INFO] Heartbeat: alive; runningJobs=0/2; enabledJobs=3; running=none.'
        '[2026-10-25 01:00:00][WARN] Failed to write heartbeat: synthetic share timeout.'
        '[2026-10-25 01:01:00][WARN] Failed to write heartbeat: synthetic share timeout.'
        '[2026-10-25 01:02:00][WARN] Failed to write heartbeat: synthetic share timeout.'
        '[2026-10-25 01:03:00][ERROR] State persistence unavailable after 60 second(s) of UNC write retries; the scheduler remains alive but new job launches are paused. Path: X. Error: synthetic.'
        '[2026-10-25 01:04:00][ERROR] Central cluster configuration reload rejected; the last valid cluster settings stay in effect. Synthetic validation error.'
        '[2026-10-25 01:30:00][INFO] State persistence recovered after 4 failed write(s); new job launches are enabled again.'
        '[2026-10-25 02:00:00][INFO] Runtime update candidate detected; stability check 1/2. Orchestrator 9.9.1 -> 9.9.2.'
        '[2026-10-25 02:01:00][INFO] Email sent via Graph: [SMART 365] - [demo] - [ Orchestrator] - Runtime update detected: 9.9.1 -> 9.9.2'
        '[2026-10-25 02:02:00][INFO] Validated runtime update is stable and signed. Recycling cleanly so Task Scheduler can start it.'
        '[2026-10-25 02:02:01][INFO] Heartbeat published as Recycling (RuntimeUpdate); peers tolerate it for 15 minute(s).'
        '[2026-10-25 02:02:02][INFO] SmartM365-Inventory-Orchestrator exiting normally (exit code 0).'
        '[2026-10-25 02:05:00][INFO] Maintenance control applied: revision=4; enabled=False; catchUpAll=True; reason=Synthetic.'
        '2026-10-25 02:06:00 [WARNING] SharePoint upload failed but script continues: synthetic.'
        '2026-10-25 02:07:00 [INFO] SharePoint file uploaded: synthetic (ignored).'
        'Continuation line without timestamp (ignored).'
        '[2026-10-25 03:00:00][INFO] Job JobA: finished with status Success (exit code 0, duration 5s).'
    )
    Write-SyntheticLog -Path $workerBToday -Lines @(
        '[2026-10-25 04:00:00][INFO] Manual stop request consumed. RequestedBy=operator; RequestedFrom=PC; RequestedAt=2026-10-25T03:59:00Z. No running inventory child job will be killed.'
        '[2026-10-25 04:10:00][WARN] Orphan stop request retired: target PID 202 started 2026-10-24T00:00:00Z no longer runs. RequestedBy=operator; RequestedAt=2026-10-24.'
        '[2026-10-25 11:30:00][ERROR] Jobs manifest reload rejected; the last valid manifest stays in effect. Synthetic.'
    ) -Tail '[2026-10-25 11:59:00][ERROR] Partial line still being written'
    $hashesBefore = @{}
    foreach ($file in @(Get-ChildItem -LiteralPath $root -Recurse -File)) { $hashesBefore[$file.FullName] = '{0}|{1}' -f (Get-FileHash -LiteralPath $file.FullName).Hash, $file.LastWriteTimeUtc.Ticks }

    # Categories and levels
    Assert-Events ((Get-SmartM365OrchestratorLogEventCategory -Level INFO -Message 'Email sent via Graph: Runtime update detected') -eq $null) 'Mail lines must not be events.'
    Assert-Events ((Get-SmartM365OrchestratorLogEventCategory -Level INFO -Message 'Heartbeat: alive; runningJobs=0/2') -eq $null) 'Periodic heartbeat must not be an event.'
    Assert-Events ((Get-SmartM365OrchestratorLogEventCategory -Level WARN -Message 'Something unusual') -eq 'Warning') 'Uncategorized WARN must be kept.'
    Assert-Events ((Get-SmartM365OrchestratorLogEventCategory -Level ERROR -Message 'Job stop request scan failed: x') -eq 'StopRequest') 'Job stop request errors must be StopRequest.'

    $cache = @{}
    $result = Get-SmartM365OrchestratorLogEvents -LogFolderPath $logRoot -Server @('WorkerB', 'WorkerA', 'WorkerC', 'WorkerA') -Hours 24 -Now $now -Cache $cache
    $events = @($result.Events)
    $byCategory = @{}; foreach ($group in @($events | Group-Object Category)) { $byCategory[$group.Name] = $group.Count }
    Assert-Events ($events.Count -eq 18) "24 h window must hold 18 events, got $($events.Count)."
    Assert-Events ($byCategory['Lifecycle'] -eq 6) 'Lifecycle count (two exits, start, candidate, validated, recycling).'
    Assert-Events ($byCategory['Heartbeat'] -eq 4 -and $byCategory['StatePersistence'] -eq 2 -and $byCategory['ConfigReload'] -eq 2) 'Heartbeat, state persistence and reload counts.'
    Assert-Events ($byCategory['StopRequest'] -eq 2 -and $byCategory['Maintenance'] -eq 1 -and $byCategory['Warning'] -eq 1) 'Stop request, maintenance and warning counts.'
    Assert-Events (-not $byCategory.ContainsKey('Error')) 'No uncategorized ERROR inside the window.'
    Assert-Events (@($events | Where-Object { $_.Message -like '*older than one day*' }).Count -eq 0) 'Events older than the period must be excluded.'
    Assert-Events (@($events | Where-Object { $_.Message -like '*Partial line*' }).Count -eq 0) 'An unterminated line must wait for its end of line.'
    Assert-Events (@($events | Where-Object { $_.Level -eq 'WARN' -and $_.Message -like 'SharePoint upload failed*' }).Count -eq 1) 'WARNING must be normalized to WARN.'
    Assert-Events ($events[0].Time -ge $events[-1].Time -and $events[0].Server -eq 'WorkerB') 'Events must be sorted newest first.'
    Assert-Events ($events[-1].Message -like 'Failed to write heartbeat*' -and $events[-1].Time -eq [datetime]'2026-10-24 13:00:00') 'Previous day file must be read for a 24 h window.'
    Assert-Events (@($events | Where-Object { $_.Category -eq 'Lifecycle' -and $_.Message -like '*v9.9.1 starting*' }).Count -eq 1) 'BOM start line must be parsed.'
    $sources = @{}; foreach ($source in @($result.Sources)) { $sources[$source.Server] = $source }
    Assert-Events (@($result.Sources).Count -eq 3) 'Servers must be de-duplicated.'
    Assert-Events ($sources['WorkerA'].Status -eq 'OK' -and $sources['WorkerA'].Files -eq 2 -and $sources['WorkerC'].Status -eq 'Log folder not found') 'Source status per server.'
    Assert-Events ($cache.Count -eq 3) 'Cache must hold one entry per read file.'

    $shortWindow = Get-SmartM365OrchestratorLogEvents -LogFolderPath $logRoot -Server @('WorkerA', 'WorkerB') -Hours 1 -Now $now -Cache @{}
    Assert-Events (@($shortWindow.Events).Count -eq 1 -and @($shortWindow.Events)[0].Category -eq 'ConfigReload') '1 h window filter.'

    # Display filters and grouping
    $all = @($result.Events)
    $grouped = Select-SmartM365OrchestratorLogEvents -Events $all -GroupRepeats -Since $result.Since
    $heartbeatRow = @($grouped.Rows | Where-Object Category -eq 'Heartbeat')
    Assert-Events ($heartbeatRow.Count -eq 1 -and $heartbeatRow[0].Count -eq 4 -and $heartbeatRow[0].First -eq '2026-10-24 13:00:00' -and $heartbeatRow[0].Time -eq '2026-10-25 01:02:00') 'Repeats grouped with count, first and last time.'
    Assert-Events ($grouped.Matching -eq 18 -and $grouped.Total -eq 14) 'Grouped totals.'
    Assert-Events ((Select-SmartM365OrchestratorLogEvents -Events $all).Total -eq 18) 'Ungrouped rows.'
    Assert-Events ((Select-SmartM365OrchestratorLogEvents -Events $all -Server WorkerB).Total -eq 3) 'Server filter.'
    Assert-Events ((Select-SmartM365OrchestratorLogEvents -Events $all -Level ERROR).Total -eq 3) 'ERROR filter.'
    Assert-Events ((Select-SmartM365OrchestratorLogEvents -Events $all -Level ErrorsAndWarnings).Total -eq 9) 'Errors and warnings filter.'
    Assert-Events ((Select-SmartM365OrchestratorLogEvents -Events $all -Level INFO).Total -eq 9) 'INFO filter.'
    Assert-Events ((Select-SmartM365OrchestratorLogEvents -Events $all -Category StopRequest).Total -eq 2) 'Category filter.'
    Assert-Events ((Select-SmartM365OrchestratorLogEvents -Events $all -Text 'ORPHAN').Total -eq 1) 'Case-insensitive text filter.'
    Assert-Events ((Select-SmartM365OrchestratorLogEvents -Events $all -Text 'statepersistence').Total -eq 2) 'Text filter matches the category.'
    Assert-Events ((Select-SmartM365OrchestratorLogEvents -Events $all -Since ([datetime]'2026-10-25 04:00:00')).Total -eq 3) 'Since filter.'
    $capped = Select-SmartM365OrchestratorLogEvents -Events $all -MaxRows 5
    Assert-Events ($capped.Truncated -and @($capped.Rows).Count -eq 5 -and $capped.Total -eq 18 -and $capped.Rows[0].Time -eq '2026-10-25 11:30:00') 'Row cap keeps the newest rows.'
    Assert-Events ((Select-SmartM365OrchestratorLogEvents -Events @()).Total -eq 0 -and (Select-SmartM365OrchestratorLogEvents -Events $null).Total -eq 0) 'Empty input.'

    # Read-only: nothing written, renamed or touched.
    foreach ($file in @(Get-ChildItem -LiteralPath $root -Recurse -File)) {
        Assert-Events ($hashesBefore.ContainsKey($file.FullName)) "Unexpected file created: $($file.FullName)"
        Assert-Events ($hashesBefore[$file.FullName] -eq ('{0}|{1}' -f (Get-FileHash -LiteralPath $file.FullName).Hash, $file.LastWriteTimeUtc.Ticks)) "Log modified: $($file.FullName)"
    }

    # Incremental read while the resident still writes (writer holds the file open).
    $offsetBefore = [long]$cache[$workerBToday].Offset
    $writer = [IO.FileStream]::new($workerBToday, [IO.FileMode]::Append, [IO.FileAccess]::Write, [IO.FileShare]::Read)
    try {
        $append = $utf8.GetBytes(" now complete.`r`n[2026-10-25 11:59:30][INFO] Maintenance control applied: revision=5; enabled=True; reason=Synthetic.`r`n[2026-10-25 11:59:40][WARN] Half")
        $writer.Write($append, 0, $append.Length); $writer.Flush()
        $incremental = Get-SmartM365OrchestratorLogEvents -LogFolderPath $logRoot -Server @('WorkerA', 'WorkerB') -Hours 24 -Now $now -Cache $cache
    }
    finally { $writer.Dispose() }
    $newEvents = @($incremental.Events | Where-Object Server -eq 'WorkerB')
    Assert-Events ($newEvents.Count -eq 5) "Completed and appended lines must be read through a shared handle, got $($newEvents.Count)."
    Assert-Events (@($newEvents | Where-Object Message -eq 'Partial line still being written now complete.').Count -eq 1) 'Completed partial line parsed once.'
    Assert-Events ([long]$cache[$workerBToday].Offset -gt $offsetBefore -and @($cache[$workerBToday].Events).Count -eq 5) 'Cache offset advanced without duplicates.'
    Assert-Events (@($incremental.Events | Where-Object Server -eq 'WorkerA').Count -eq 15) 'Unchanged file served from cache.'
    $again = Get-SmartM365OrchestratorLogEvents -LogFolderPath $logRoot -Server @('WorkerA', 'WorkerB') -Hours 24 -Now $now -Cache $cache
    Assert-Events (@($again.Events).Count -eq @($incremental.Events).Count) 'Repeated read without change is stable.'
    Write-SyntheticLog -Path $workerBToday -Lines @('[2026-10-25 05:00:00][ERROR] Replaced shorter file.')
    $reset = Get-SmartM365OrchestratorLogEvents -LogFolderPath $logRoot -Server @('WorkerB') -Hours 24 -Now $now -Cache $cache
    Assert-Events (@($reset.Events).Count -eq 1 -and @($reset.Events)[0].Category -eq 'Error') 'A shrunk file is re-read from the start.'
    Assert-Events (-not $cache.ContainsKey($workerAToday)) 'Cache entries of servers no longer read are dropped.'
    $locked = [IO.FileStream]::new($workerAToday, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    try { $lockedRead = Get-SmartM365OrchestratorLogEvents -LogFolderPath $logRoot -Server @('WorkerA') -Hours 24 -Now $now -Cache @{} }
    finally { $locked.Dispose() }
    Assert-Events (@($lockedRead.Sources)[0].Status -like 'Read error:*' -and @($lockedRead.Events).Count -eq 2) 'An exclusively locked log is reported without failing the other files.'

    # Static read-only guard on the log functions.
    $tokens = $null; $errors = $null
    $moduleAst = [Management.Automation.Language.Parser]::ParseFile($insightsModule, [ref]$tokens, [ref]$errors)
    Assert-Events ($errors.Count -eq 0) 'Insights module parses.'
    foreach ($name in @('Get-SmartM365OrchestratorLogEventCategory', 'Read-InsightsLogTail', 'Get-SmartM365OrchestratorLogEvents', 'Select-SmartM365OrchestratorLogEvents')) {
        $definition = $moduleAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name }, $true)
        Assert-Events ($null -ne $definition) "$name is defined."
        Assert-Events ($definition.Extent.Text -notmatch '(?i)Set-Content|Add-Content|Out-File|Remove-Item|New-Item|Move-Item|Copy-Item|WriteAll|FileAccess\]::(Read)?Write|FileMode\]::(Create|Append|Truncate)') "$name must stay read-only."
    }

    # GUI wiring and background refresh contract.
    $guiAst = [Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$tokens, [ref]$errors)
    Assert-Events ($errors.Count -eq 0) 'GUI parses.'
    $guiText = [IO.File]::ReadAllText($guiPath)
    Assert-Events ($guiText -match '<TabItem Header="Events">' -and $guiText -match "'Activity','Events'\)\]") 'Events tab and demo view exist.'
    $startRefresh = $guiAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Start-GuiRefresh' }, $true)
    Assert-Events ($startRefresh.Extent.Text -match "LogEventsHours = \`$\(if \(\`$Mode -ne 'History' -and \(Test-GuiEventsTabVisible\)\)") 'Logs are read only by live/full refreshes while the Events tab is visible.'
    $updateView = $guiAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Update-GuiEventsView' }, $true)
    Assert-Events ($updateView.Extent.Text -notmatch 'Get-SmartM365OrchestratorLogEvents') 'Display filters never read files on the UI thread.'
    $definition = $guiAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-GuiRefreshWork' }, $true)
    . ([scriptblock]::Create($definition.Extent.Text))
    $shared = Join-Path $root 'Shared'
    $null = New-Item -ItemType Directory -Path $shared
    $managementPath = Join-Path $root 'Management.psm1'
    $stubInsightsPath = Join-Path $root 'Insights.psm1'
    @'
function Read-SmartM365OrchestratorJson { param($Path) [pscustomobject]@{ExpectedOrchestratorServers=@('WorkerA','WorkerB')} }
Export-ModuleMember -Function *
'@ | Set-Content -LiteralPath $managementPath
    @"
Import-Module '$insightsModule' -Global -Function Get-SmartM365OrchestratorLogEvents
function Get-SmartM365OrchestratorOperations { param(`$SharedDataFolderPath,`$ClusterDocument,`$MailFolderPath,`$MailHours) Start-Sleep -Milliseconds 900; [pscustomobject]@{Servers=@();Running=@();Pending=@();Incidents=@();Mails=@()} }
function Get-SmartM365OrchestratorRecentPipelineRuns { param(`$SharedDataFolderPath,`$Count) @() }
Export-ModuleMember -Function Get-SmartM365OrchestratorOperations, Get-SmartM365OrchestratorRecentPipelineRuns
"@ | Set-Content -LiteralPath $stubInsightsPath
    function Invoke-RefreshWorker {
        param([string]$Mode, [int]$Hours, $Cache)
        $context = [pscustomobject]@{ Worker = $null; Progress = [hashtable]::Synchronized(@{ Message = '' }) }
        $inputData = [pscustomobject]@{ Mode = $Mode; Root = $shared; MailFolderPath = $logRoot; ManagementModulePath = $managementPath; InsightsModulePath = $stubInsightsPath
            HistoryFrom = $now.Date; HistoryTo = $now; HistoryServer = ''; HistoryJob = ''; HistoryStatus = ''; OnlyFailures = $false; LogEventsHours = $Hours; LogEventsCache = $Cache }
        $watch = [Diagnostics.Stopwatch]::StartNew()
        Assert-Events (Start-SmartM365OrchestratorGuiWorker -Context $context -Work (Get-GuiRefreshWork) -InputObject $inputData) 'Worker started.'
        $watch.Stop()
        Assert-Events ($watch.ElapsedMilliseconds -lt 700) "Refresh must not block the caller ($($watch.ElapsedMilliseconds) ms)."
        $deadline = [datetime]::UtcNow.AddSeconds(20)
        while (-not $context.Worker.Handle.IsCompleted -and [datetime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 30 }
        Assert-Events $context.Worker.Handle.IsCompleted 'Worker completed in time.'
        $received = Receive-SmartM365OrchestratorGuiWorker -Context $context
        Assert-Events (-not $received.Error -and @($received.Output).Count -eq 1) "Worker result: $($received.Error)"
        return $received.Output[0]
    }
    $guiCache = @{}
    $live = Invoke-RefreshWorker -Mode Live -Hours 24 -Cache $guiCache
    Assert-Events ($null -ne $live.LogEvents -and -not $live.LogEventsError -and @($live.LogEvents.Events).Count -gt 0) 'Live refresh returns events when the tab is visible.'
    Assert-Events ($guiCache.Count -gt 0) 'Worker fills the GUI-owned cache.'
    $hidden = Invoke-RefreshWorker -Mode Live -Hours 0 -Cache @{}
    Assert-Events ($null -eq $hidden.LogEvents -and -not $hidden.LogEventsError) 'No log read while the Events tab is hidden.'
    $broken = Invoke-RefreshWorker -Mode Live -Hours 24 -Cache 'not-a-cache'
    Assert-Events ($null -eq $broken.LogEvents -and $broken.LogEventsError -and $null -ne $broken.Operations) 'A log read failure is isolated from live operations.'

    "[{0}] PASS: Orchestrator GUI events, {1} offline checks." -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $script:Checks
}
finally {
    if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force }
}
