#Requires -Version 7.0
<#
.SYNOPSIS
Checks the GUI Events tab data layer on synthetic resident orchestrator logs, offline.
.DESCRIPTION
Builds fictional daily logs for three servers, then verifies parsing, categories, period
and display filters, repeat grouping, incremental cached reads, shared reads of a file
held open by a writer, read-only behaviour, and the background GUI refresh contract.
.VERSION
1.0.1
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
    Assert-Events ($guiText -match '<TabItem Header="Events">' -and $guiText -match "'Activity','Events'[,)]") 'Events tab and demo view exist.'
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
    # The worker reads the 24 hours before the real clock: add a line written now.
    $realNow = Get-Date
    Write-SyntheticLog -Path (Get-LogPath 'WorkerB' $realNow) -Lines @(('[{0}][ERROR] Central cluster configuration reload rejected; synthetic line written now.' -f $realNow.AddMinutes(-1).ToString('yyyy-MM-dd HH:mm:ss')))
    $guiCache = @{}
    $live = Invoke-RefreshWorker -Mode Live -Hours 24 -Cache $guiCache
    Assert-Events ($null -ne $live.LogEvents -and -not $live.LogEventsError -and @($live.LogEvents.Events).Count -gt 0) ("Live refresh returns events when the tab is visible: " + $live.LogEventsError)
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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAJneJiKLksLmoH
# t6b5vcDALs3uaYbgiDLuDG0UnITaJaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIA5S/Pyz6BJK/teNj/P7sElASj5DY/B6R8rdhgxJwwMnMA0GCSqG
# SIb3DQEBAQUABIIBgAXqheV7DfZCpQHgu2IKhR6lLEh90EJhozo1zc9BGn/LELp7
# IE4/pUVnYmCPnxOeA7AKit6jpegn9PzYwZ5IR0nPzKGwa719BDtAcIfH38yCEScl
# LchkfVF7oiodEYy/F1cQRu8+8ffBXnpa+362lKmY8hzCCw/rh8Ww04RaIOu/H48H
# fj6jn27sjYx3JPSjnxw7zvHFRsYNXTZmxnYO5D+8D14S9uMsXTsdVKKNctBBKCNK
# 6lr67L7A1d1JUS7ijMNsmiGSViIErvlzPVx8stEsl/bTC1u8seZzwymkMCZTc8g3
# HmihX9lcxVgR+6WXzC8WCkU/HbpL4ry/55ADp6hZ0fYJAGl1RpSIXlgXrGlx+Lnt
# TNnSY5qsIGAXI42xwctDlvQyXrMtYamD7Wpl7Xf1/7JGWYhLThSmwaXLmbzucMUJ
# vPM3B8fjyUTPoIqrzxfmN1HHrzGAbtKd6xANnpI8xSQQkvmn9PzpWQK6ne0Z/a1d
# Nf0pa/5HK257UwOCGqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjYyMjI2
# MDBaMC8GCSqGSIb3DQEJBDEiBCA2rzrMYm8zbpeXKtDTEY49FCMZUWPJr2vyKepI
# 58nJDTANBgkqhkiG9w0BAQEFAASCAgBxaL5zSC4k7F5YjtfB7ew6czKIbAn/Jb3A
# Qw5MVbN7X19hdhe+o6TMhsrMGgV5cqxm0AvRVdNLUp8/ukxysFt6Uyk6q8Vz41L0
# HfqoIT/y9sdS9DlGT3HyxPNfohktei9JiDkcOk3xeBMiVgU0fEq1d7vfC1ojpK90
# DrkvEyz/kQ+Yw8hT6TS6/qut8HwYuvvN6YqFBoD25Bf42vWUeUvQOOsgrfpghFsY
# lKqf0Vj9K2wKpkiMO+nEAtQjlw+ETtGGRMBZd+Ch7eN0tEDNAk1bQ5RvllduGJXb
# CC1i0rVNStim1c70m8eUG3dUQ9fCbk8lObY0RveeVeWMcULFQ00LkJQrZa4sLcez
# wAJ8xqzmrcUu7YcZPz77LmzxqUaAG0lQoohU7ReowZt+J5vszxTcaP6FesAoy1lN
# BDpOzaEZ966EXCoVdDMG0aQ29+EiI/0QfKQyTPVJOnarrGKw04ZBsRcBJN4k7JIo
# H3l2jnJ0e6q/DHE7ZrlCfoWoCDi3iOAaDDUhZ8qzau8c0G5SvSG9Rc8EVA+JaDwn
# yjR0+lY206Ymd/CD7EixC6dujsMhOzI/ryRCnA4ZSFijVryxIhdion3BqOl9HSc2
# G1EqMhTgwA6/bOJL1chYCIoLIBvRc2KU7E46Gzq5FufD1zChxCTgQQVXl+2LkSUv
# 0pzZ9/Na4w==
# SIG # End signature block
