#Requires -Version 7.0
<#
.SYNOPSIS
Checks the orchestrator GUI Dashboard offline: aggregation, sources, worker, loading overlay and rendering.
.DESCRIPTION
Uses synthetic operations, job runs, election plan, requests and log events only. Verifies the
banner, tiles, items to review, server rows and 24-hour timeline, the background refresh contract,
the delayed non-blocking loading overlay and the WPF rendering of tiles and timeline. Run with -STA.
.VERSION
1.0.0
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version 2.0
if ([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) { throw 'Run with -STA.' }
$folder = Join-Path $PSScriptRoot '../SmartInventory/Orchestrator'
$insightsModule = (Resolve-Path (Join-Path $folder 'SmartM365.Orchestrator.Insights.psm1')).Path
$guiPath = Join-Path $folder 'SmartM365-Inventory-Orchestrator-GUI.ps1'
Import-Module $insightsModule -Force
Import-Module (Join-Path $folder 'SmartM365.Orchestrator.GuiWorker.psm1') -Force
Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase
$script:Checks = 0
function Assert-Dashboard {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw "FAIL: $Message" }
    $script:Checks++
}

$now = [datetime]'2026-10-25 12:00:00'
function New-Job {
    param([string]$Name, [string]$Mode = 'Elected', [bool]$Enabled = $true, [string[]]$Times = @('13:00'), [string[]]$Allowed = @(), [int]$Estimated = 30)
    [pscustomobject]@{ Name = $Name; Enabled = $Enabled; AssignmentMode = $Mode; AllowedServers = $Allowed; EstimatedDurationMinutes = $Estimated
        Schedule = [pscustomobject]@{ Type = 'Daily'; Times = $Times; DaysOfWeek = @() } }
}
$jobsDocument = [pscustomobject]@{ Jobs = @(
    New-Job 'JobOk' -Times @('13:00', '23:00')
    New-Job 'JobFailing' -Times @('14:00')
    New-Job 'JobStale' -Times @('15:00')
    New-Job 'JobOrphan' -Times @('16:00')
    New-Job 'JobPinned' -Mode 'Pinned' -Allowed @('SRV02') -Times @('17:00')
    New-Job 'JobLegacy' -Mode 'Legacy' -Times @('18:00')
    New-Job 'JobManual' -Mode 'Manual' -Times @('19:00')
    New-Job 'JobDisabled' -Enabled $false -Times @('20:00')
) }
function New-Server {
    param([string]$Name, [string]$State = 'Online', [string]$Persistence = 'OK', $Days = 200, [string]$Blocks = '')
    [pscustomobject]@{ Server = $Name; State = $State; HeartbeatAgeMinutes = 0.5; Version = '9.9.9'; Running = 0; Pending = 0; RecycleIn = '10h00'
        StatePersistence = $Persistence; ReadyCapabilities = 'AD, Graph, SharedRuntime'; CertificateDaysRemaining = $Days; PrerequisiteBlocks = $Blocks; Maintenance = 'OFF' }
}
$operations = [pscustomobject]@{
    Servers = @((New-Server 'SRV01'), (New-Server 'SRV02' -Days 20), (New-Server 'SRV03' -State 'Stale' -Persistence 'Launches paused: share' -Days 5 -Blocks 'Certificate: missing'))
    Running = @([pscustomobject]@{ Server = 'SRV01'; Job = 'JobOk'; Started = '2026-10-25 11:30'; Pid = '1' })
    Pending = @([pscustomobject]@{ Server = 'SRV02'; Job = 'JobPinned'; Reason = 'Dependency' }, [pscustomobject]@{ Server = 'SRV02'; Job = 'JobLegacy'; Reason = 'Dependency' })
    Incidents = @([pscustomobject]@{ ObservedBy = 'SRV01'; Issue = 'SRV03: HeartbeatStale' })
    Mails = @(); Maintenance = [pscustomobject]@{ Enabled = $false; Reason = '' }; MaintenanceError = ''
}
$health = @(
    [pscustomobject]@{ Name = 'JobOk'; Health = 'OK'; LastStatus = 'Success'; LastRun = '2026-10-25 01:00'; LastSuccess = '2026-10-25 01:10'; SuccessAgeHours = 11; ExpectedMaxAgeHours = 26; AverageDurationMinutes = 45 }
    [pscustomobject]@{ Name = 'JobFailing'; Health = 'Failing'; LastStatus = 'Failed'; LastRun = '2026-10-25 02:00'; LastSuccess = '2026-10-24 02:10'; SuccessAgeHours = 34; ExpectedMaxAgeHours = 26; AverageDurationMinutes = 10 }
    [pscustomobject]@{ Name = 'JobStale'; Health = 'Stale'; LastStatus = 'Success'; LastRun = '2026-10-23 03:00'; LastSuccess = '2026-10-23 03:10'; SuccessAgeHours = 57; ExpectedMaxAgeHours = 26; AverageDurationMinutes = $null }
)
$runs = @(
    [pscustomobject]@{ Server = 'SRV01'; JobName = 'JobFailing'; StartTime = [datetime]'2026-10-25 02:00'; EndTime = [datetime]'2026-10-25 02:10'; Status = 'Failed' }
    [pscustomobject]@{ Server = 'SRV01'; JobName = 'JobOld'; StartTime = [datetime]'2026-10-23 02:00'; EndTime = [datetime]'2026-10-23 02:10'; Status = 'TimedOut' }
    [pscustomobject]@{ Server = 'SRV02'; JobName = 'JobOk'; StartTime = [datetime]'2026-10-25 01:00'; EndTime = [datetime]'2026-10-25 01:10'; Status = 'Success' }
)
$plan = [pscustomobject]@{
    Assignments = @([pscustomobject]@{ JobName = 'JobOk'; OwnerServer = 'SRV01' }, [pscustomobject]@{ JobName = 'JobFailing'; OwnerServer = 'SRV02' }, [pscustomobject]@{ JobName = 'JobStale'; OwnerServer = 'SRV02' })
    UnassignedGroups = @([pscustomobject]@{ GroupKey = 'JobOrphan'; Jobs = @('JobOrphan'); Reason = 'No live server satisfies every required capability.' })
}
$requests = @(
    [pscustomobject]@{ BatchId = 'OLD-BATCH'; Status = 'Running'; Created = '2026-10-25 02:00'; Pending = 2 }
    [pscustomobject]@{ BatchId = 'NEW-BATCH'; Status = 'Running'; Created = '2026-10-25 11:00'; Pending = 1 }
)
$logEvents = [pscustomobject]@{ Events = @(
    [pscustomobject]@{ Time = [datetime]'2026-10-25 11:40'; Server = 'SRV03'; Level = 'WARN'; Category = 'Heartbeat'; Message = 'Failed to write heartbeat: synthetic.' }
    [pscustomobject]@{ Time = [datetime]'2026-10-25 11:50'; Server = 'SRV03'; Level = 'WARN'; Category = 'Heartbeat'; Message = 'Failed to write heartbeat: synthetic.' }
    [pscustomobject]@{ Time = [datetime]'2026-10-25 09:00'; Server = 'SRV01'; Level = 'ERROR'; Category = 'Error'; Message = 'Synthetic error.' }
    [pscustomobject]@{ Time = [datetime]'2026-10-25 11:55'; Server = 'SRV01'; Level = 'INFO'; Category = 'Lifecycle'; Message = 'Synthetic start.' }
    [pscustomobject]@{ Time = [datetime]'2026-10-23 09:00'; Server = 'SRV01'; Level = 'ERROR'; Category = 'Error'; Message = 'Older than 24 h.' }
) }

# Aggregation.
$dashboard = Get-SmartM365OrchestratorDashboard -JobsDocument $jobsDocument -Operations $operations -Health $health -Runs $runs -Plan $plan -Requests $requests -LogEvents $logEvents -Now $now
Assert-Dashboard ($dashboard.Banner.Level -eq 'Critical' -and $dashboard.Banner.Text -like '*2/3 server(s) online*1 running*2 pending*1 failure(s) in 24 h*') "Banner: $($dashboard.Banner.Text)"
$tiles = @{}; foreach ($tile in @($dashboard.Tiles)) { $tiles[$tile.Label] = $tile }
Assert-Dashboard (@($dashboard.Tiles).Count -eq 9) 'Nine tiles.'
Assert-Dashboard ($tiles['Jobs'].Value -eq '8' -and $tiles['Enabled'].Value -eq '7' -and $tiles['Jobs'].Detail -eq '1 manual, 1 disabled') 'Job tiles.'
Assert-Dashboard ($tiles['Servers online'].Value -eq '2/3' -and $tiles['Servers online'].Level -eq 'Critical' -and $tiles['Servers online'].Detail -eq 'SRV03 Stale') 'Server tile.'
Assert-Dashboard ($tiles['Running'].Value -eq '1' -and $tiles['Pending'].Value -eq '2' -and $tiles['Pending'].Detail -eq 'mostly Dependency') 'Running and pending tiles.'
Assert-Dashboard ($tiles['Job health'].Value -eq '2' -and $tiles['Job health'].Level -eq 'Critical' -and $tiles['Job health'].Detail -eq '1 failing, 1 stale, 0 no success') 'Job health tile.'
Assert-Dashboard ($tiles['Failures 24 h'].Value -eq '1' -and $tiles['Failures 24 h'].Detail -eq 'JobFailing' -and $tiles['Failures 24 h'].Target -eq 'Failures24h') 'Failures tile (TimedOut older than 24 h excluded).'
Assert-Dashboard ($tiles['Log errors 24 h'].Value -eq '1' -and $tiles['Log errors 24 h'].Detail -eq '2 warning(s)' -and $tiles['Log errors 24 h'].Target -eq 'Events') 'Log tile.'
Assert-Dashboard ($tiles['Certificate'].Value -eq '5 d' -and $tiles['Certificate'].Level -eq 'Critical') 'Certificate tile uses the lowest server.'
$actions = @($dashboard.Actions)
function Get-Action([string]$Area, [string]$Item) { @($actions | Where-Object { $_.Area -eq $Area -and $_.Item -eq $Item }) }
Assert-Dashboard ((Get-Action 'Server' 'SRV03')[0].Severity -eq 'Critical') 'Stale server is critical.'
Assert-Dashboard ((Get-Action 'State persistence' 'SRV03')[0].Severity -eq 'Critical' -and (Get-Action 'Launch prerequisite' 'SRV03')[0].Severity -eq 'Critical') 'Persistence and prerequisite blocks are critical.'
Assert-Dashboard ((Get-Action 'Certificate' 'SRV03')[0].Severity -eq 'Critical' -and (Get-Action 'Certificate' 'SRV02')[0].Severity -eq 'Warning' -and @(Get-Action 'Certificate' 'SRV01').Count -eq 0) 'Certificate thresholds.'
Assert-Dashboard ((Get-Action 'Job Failing' 'JobFailing')[0].Severity -eq 'Critical' -and (Get-Action 'Job Stale' 'JobStale')[0].Severity -eq 'Warning' -and (Get-Action 'Job Stale' 'JobStale')[0].Target -eq 'Planning') 'Job health items.'
Assert-Dashboard ((Get-Action 'Election' 'JobOrphan')[0].Detail -like 'Unassigned:*') 'Unassigned election group.'
Assert-Dashboard (@(Get-Action 'Peer monitoring' 'SRV01').Count -eq 1) 'Peer incident.'
Assert-Dashboard (@(Get-Action 'Pipeline request' 'OLD-BATCH').Count -eq 1 -and @(Get-Action 'Pipeline request' 'NEW-BATCH').Count -eq 0) 'Only requests older than 6 h are listed.'
Assert-Dashboard ((Get-Action 'Log Heartbeat' 'SRV03')[0].Detail -like '2 line(s) in the last hour*' -and (Get-Action 'Log Heartbeat' 'SRV03')[0].Target -eq 'Events') 'Recent heartbeat log lines are grouped.'
Assert-Dashboard (@($actions | Where-Object Area -eq 'Maintenance').Count -eq 0) 'No maintenance item while maintenance is off.'
$ranks = @($actions | ForEach-Object { @{ Critical = 0; Warning = 1; Info = 2 }[$_.Severity] })
Assert-Dashboard ((@($ranks) -join ',') -eq (@($ranks | Sort-Object) -join ',')) 'Items sorted by severity.'
$servers = @{}; foreach ($row in @($dashboard.Servers)) { $servers[$row.Server] = $row }
Assert-Dashboard ($servers['SRV01'].Owned -eq 1 -and $servers['SRV02'].Owned -eq 2 -and $servers['SRV03'].Owned -eq 0 -and $servers['SRV02'].Certificate -eq '20 d') 'Server rows: elected jobs and certificate.'
$timeline = $dashboard.Timeline
Assert-Dashboard ($timeline.From -eq $now -and $timeline.To -eq $now.AddHours(24)) 'Timeline window.'
Assert-Dashboard ((@($timeline.Lanes) -join ',') -eq 'SRV01,SRV02,SRV03,Any server,Unassigned') "Timeline lanes: $(@($timeline.Lanes) -join ',')"
$items = @($timeline.Items)
Assert-Dashboard (@($items | Where-Object { $_.Job -in @('JobManual', 'JobDisabled') }).Count -eq 0) 'Manual and disabled jobs are not planned.'
Assert-Dashboard (@($items | Where-Object Job -eq 'JobOk' | Where-Object Kind -eq 'Planned').Count -eq 2) 'Both remaining daily occurrences are planned.'
$okItem = @($items | Where-Object { $_.Job -eq 'JobOk' -and $_.Kind -eq 'Planned' })[0]
Assert-Dashboard ($okItem.Lane -eq 'SRV01' -and ($okItem.End - $okItem.Start).TotalMinutes -eq 45) 'Elected owner lane and average duration.'
Assert-Dashboard ((@($items | Where-Object Job -eq 'JobStale')[0].End - @($items | Where-Object Job -eq 'JobStale')[0].Start).TotalMinutes -eq 30) 'Estimated duration fallback.'
Assert-Dashboard (@($items | Where-Object Job -eq 'JobOrphan')[0].Kind -eq 'Unassigned' -and @($items | Where-Object Job -eq 'JobPinned')[0].Lane -eq 'SRV02' -and @($items | Where-Object Job -eq 'JobLegacy')[0].Lane -eq 'Any server') 'Unassigned, pinned and legacy lanes.'
$runningItem = @($items | Where-Object Kind -eq 'Running')
Assert-Dashboard ($runningItem.Count -eq 1 -and $runningItem[0].Start -eq [datetime]'2026-10-25 11:30' -and $runningItem[0].End -eq [datetime]'2026-10-25 12:15') 'Running job bar.'

$operations.Maintenance.Enabled = $true; $operations.Maintenance.Reason = 'Synthetic window'
$maintenance = Get-SmartM365OrchestratorDashboard -JobsDocument $jobsDocument -Operations $operations -Health $health -Runs $runs -Plan $plan -Now $now
Assert-Dashboard (@($maintenance.Actions | Where-Object Area -eq 'Maintenance')[0].Detail -like '*Synthetic window*' -and @($maintenance.Timeline.Items | Where-Object Kind -eq 'Suspended').Count -gt 0) 'Maintenance item and suspended occurrences.'
Assert-Dashboard (@($maintenance.Tiles | Where-Object Label -eq 'Log errors 24 h')[0].Value -eq '-') 'Log tile without resident logs.'
$operations.Maintenance.Enabled = $false
$healthy = Get-SmartM365OrchestratorDashboard -JobsDocument ([pscustomobject]@{ Jobs = @(New-Job 'OnlyJob') }) -Operations ([pscustomobject]@{ Servers = @(New-Server 'SRV01'); Running = @(); Pending = @(); Incidents = @(); Mails = @(); Maintenance = $null; MaintenanceError = '' }) -Now $now
Assert-Dashboard ($healthy.Banner.Level -eq 'Normal' -and $healthy.Banner.Text -like 'Everything looks healthy.*' -and @($healthy.Actions).Count -eq 0) 'Healthy banner.'

# Sources: runs CSV and election plan read from a synthetic shared folder.
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-GuiDashboard-' + [guid]::NewGuid().ToString('N'))
try {
    $null = New-Item -ItemType Directory -Path (Join-Path $root 'SRV01/JobRuns'), (Join-Path $root 'Election'), (Join-Path $root 'Config')
    @(
        [pscustomobject]@{ JobName = 'JobOk'; ScheduledTime = '2026-10-25 01:00:00'; StartTime = '2026-10-25 01:00:00'; EndTime = '2026-10-25 01:20:00'; DurationSec = 1200; ExitCode = 0; Status = 'Success'; RetryCount = 0; LogPath = '' }
        [pscustomobject]@{ JobName = 'JobFailing'; ScheduledTime = '2026-10-25 02:00:00'; StartTime = '2026-10-25 02:00:00'; EndTime = '2026-10-25 02:05:00'; DurationSec = 300; ExitCode = 1; Status = 'Failed'; RetryCount = 0; LogPath = '' }
    ) | Export-Csv -LiteralPath (Join-Path $root 'SRV01/JobRuns/Orchestrator_JobRuns_20261025.csv') -NoTypeInformation
    $plan | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath (Join-Path $root 'Election/Orchestrator-ElectionPlan.json')
    $data = Get-SmartM365OrchestratorDashboardData -SharedDataFolderPath $root -JobsDocument $jobsDocument -Operations $operations -Requests @() -Now $now
    Assert-Dashboard (@($data.Actions | Where-Object { $_.Area -eq 'Job Failing' -and $_.Item -eq 'JobFailing' }).Count -eq 1) 'Job health computed from the shared runs.'
    Assert-Dashboard (@($data.Tiles | Where-Object Label -eq 'Failures 24 h')[0].Value -eq '1' -and @($data.Actions | Where-Object Area -eq 'Election').Count -eq 1) 'Failures and plan read from the shared folder.'
    $okFromRuns = @($data.Timeline.Items | Where-Object { $_.Job -eq 'JobOk' -and $_.Kind -eq 'Planned' })[0]
    Assert-Dashboard (($okFromRuns.End - $okFromRuns.Start).TotalMinutes -eq 20 -and $okFromRuns.Lane -eq 'SRV01') 'Average duration from the shared runs.'

    # Background refresh contract.
    $tokens = $null; $errors = $null
    $guiAst = [Management.Automation.Language.Parser]::ParseFile($guiPath, [ref]$tokens, [ref]$errors)
    Assert-Dashboard ($errors.Count -eq 0) 'GUI parses.'
    function Get-GuiFunction([string]$Name) { $guiAst.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $Name }, $true) }
    . ([scriptblock]::Create((Get-GuiFunction 'Get-GuiRefreshWork').Extent.Text))
    $managementPath = Join-Path $root 'Management.psm1'; $stubInsightsPath = Join-Path $root 'Insights.psm1'
    @'
function Read-SmartM365OrchestratorJson { param($Path) if ($Path -like '*Jobs.json') { if (Test-Path (Join-Path (Split-Path (Split-Path $Path)) 'jobs-fail.flag')) { throw 'Synthetic jobs read failure' }; return $script:Jobs } [pscustomobject]@{ExpectedOrchestratorServers=@('SRV01','SRV02')} }
Export-ModuleMember -Function *
'@ | Set-Content -LiteralPath $managementPath
    $jobsJsonPath = Join-Path $root 'jobs.json'
    $jobsDocument | ConvertTo-Json -Depth 10 | Set-Content -LiteralPath $jobsJsonPath
    (Get-Content -LiteralPath $managementPath -Raw).Replace('return $script:Jobs', "return (Get-Content -LiteralPath '$jobsJsonPath' -Raw | ConvertFrom-Json)") | Set-Content -LiteralPath $managementPath
    @"
Import-Module '$insightsModule' -Global -Function Get-SmartM365OrchestratorDashboardData, Get-SmartM365OrchestratorLogEvents
function Get-SmartM365OrchestratorOperations { param(`$SharedDataFolderPath,`$ClusterDocument,`$MailFolderPath,`$MailHours) Start-Sleep -Milliseconds 800; [pscustomobject]@{Servers=@([pscustomobject]@{Server='SRV01';State='Online';HeartbeatAgeMinutes=0.2;Version='9';Running=0;Pending=0;RecycleIn='';StatePersistence='OK';ReadyCapabilities='';CertificateDaysRemaining=100;PrerequisiteBlocks=''});Running=@();Pending=@();Incidents=@();Mails=@();Maintenance=`$null;MaintenanceError=''} }
function Get-SmartM365OrchestratorRecentPipelineRuns { param(`$SharedDataFolderPath,`$Count) @() }
Export-ModuleMember -Function Get-SmartM365OrchestratorOperations, Get-SmartM365OrchestratorRecentPipelineRuns
"@ | Set-Content -LiteralPath $stubInsightsPath
    function Invoke-DashboardWorker {
        param([bool]$IncludeDashboard)
        $context = [pscustomobject]@{ Worker = $null; Progress = [hashtable]::Synchronized(@{ Message = '' }) }
        $inputData = [pscustomobject]@{ Mode = 'Live'; Root = $root; MailFolderPath = ''; ManagementModulePath = $managementPath; InsightsModulePath = $stubInsightsPath
            HistoryFrom = $now.Date; HistoryTo = $now; HistoryServer = ''; HistoryJob = ''; HistoryStatus = ''; OnlyFailures = $false; LogEventsHours = 0; LogEventsCache = @{}; IncludeDashboard = $IncludeDashboard }
        $watch = [Diagnostics.Stopwatch]::StartNew()
        Assert-Dashboard (Start-SmartM365OrchestratorGuiWorker -Context $context -Work (Get-GuiRefreshWork) -InputObject $inputData) 'Worker started.'
        $watch.Stop()
        Assert-Dashboard ($watch.ElapsedMilliseconds -lt 700) "Dashboard refresh must not block the caller ($($watch.ElapsedMilliseconds) ms)."
        $deadline = [datetime]::UtcNow.AddSeconds(30)
        while (-not $context.Worker.Handle.IsCompleted -and [datetime]::UtcNow -lt $deadline) { Start-Sleep -Milliseconds 30 }
        $received = Receive-SmartM365OrchestratorGuiWorker -Context $context
        Assert-Dashboard (-not $received.Error -and @($received.Output).Count -eq 1) "Worker result: $($received.Error)"
        return $received.Output[0]
    }
    $withDashboard = Invoke-DashboardWorker -IncludeDashboard $true
    Assert-Dashboard ($null -ne $withDashboard.Dashboard -and -not $withDashboard.DashboardError -and @($withDashboard.Dashboard.Tiles).Count -eq 9) "Worker builds the dashboard: $($withDashboard.DashboardError)"
    $without = Invoke-DashboardWorker -IncludeDashboard $false
    Assert-Dashboard ($null -eq $without.Dashboard -and -not $without.DashboardError) 'No dashboard work when the tab is hidden.'
    $null = New-Item -ItemType File -Path (Join-Path $root 'jobs-fail.flag')
    $failed = Invoke-DashboardWorker -IncludeDashboard $true
    Assert-Dashboard ($null -eq $failed.Dashboard -and $failed.DashboardError -like '*Synthetic jobs read failure*' -and $null -ne $failed.Operations) 'A dashboard failure is isolated from live operations.'

    # Loading overlay: delayed, non-modal, only on the Dashboard tab, released when the read ends.
    foreach ($name in @('Test-GuiDashboardTabVisible', 'Update-GuiDashboardLoading', 'Stop-GuiDashboardLoading')) { . ([scriptblock]::Create((Get-GuiFunction $name).Extent.Text)) }
    $tab = [pscustomobject]@{ Header = 'Dashboard' }
    $script:Controls = @{ MainTabs = [pscustomobject]@{ SelectedItem = $tab }; DashboardLoadingOverlay = [pscustomobject]@{ Visibility = 'Collapsed' }; DashboardLoadingText = [pscustomobject]@{ Text = '' } }
    $script:RefreshContext = [pscustomobject]@{ Progress = @{ Message = 'Building dashboard: synthetic...' } }
    $script:DashboardLoading = [pscustomobject]@{ Pending = $true; Clock = [Diagnostics.Stopwatch]::StartNew() }
    $script:PendingRefreshRequest = $null
    Update-GuiDashboardLoading
    Assert-Dashboard ($script:Controls.DashboardLoadingOverlay.Visibility -eq 'Collapsed') 'No overlay flash for a fast read.'
    Start-Sleep -Milliseconds 400
    Update-GuiDashboardLoading
    Assert-Dashboard ($script:Controls.DashboardLoadingOverlay.Visibility -eq 'Visible' -and $script:Controls.DashboardLoadingText.Text -like 'Building dashboard: synthetic...*s elapsed*') 'Overlay shown with progress after 350 ms.'
    $tab.Header = 'Planning'; Update-GuiDashboardLoading
    Assert-Dashboard ($script:Controls.DashboardLoadingOverlay.Visibility -eq 'Collapsed') 'Overlay hidden on other tabs.'
    $tab.Header = 'Dashboard'
    $script:PendingRefreshRequest = @{ Mode = 'Live'; ShowLoading = $true }
    Stop-GuiDashboardLoading
    Assert-Dashboard $script:DashboardLoading.Pending 'Overlay kept while a queued dashboard read follows.'
    $script:PendingRefreshRequest = $null
    Stop-GuiDashboardLoading
    Update-GuiDashboardLoading
    Assert-Dashboard (-not $script:DashboardLoading.Pending -and $script:Controls.DashboardLoadingOverlay.Visibility -eq 'Collapsed') 'Overlay released when the read ends.'
    $guiText = [IO.File]::ReadAllText($guiPath)
    Assert-Dashboard ($guiText.Contains('$script:RefreshTimer.Add_Tick({ Receive-GuiRefresh; Update-GuiDashboardLoading })')) 'Overlay driven by the existing non-blocking refresh timer.'
    Assert-Dashboard ((Get-GuiFunction 'Invoke-AutoRefresh').Extent.Text -notmatch 'ShowLoading') 'Auto-refresh never shows the overlay.'
    Assert-Dashboard ($guiText -match "if \(Test-GuiDashboardTabVisible\) \{ Start-GuiRefresh -Mode Live -ShowLoading \}") 'Opening the Dashboard tab requests the overlay.'
    Assert-Dashboard ($guiText -match '<Border x:Name="DashboardLoadingOverlay" Grid.RowSpan="4"') 'Overlay covers the Dashboard tab only (not modal).'

    # WPF rendering of tiles, items and timeline.
    $colorLines = @([regex]::Matches($guiText, '(?m)^\$script:Dashboard(Level|Banner|Timeline)Colors = .*$') | ForEach-Object Value)
    Assert-Dashboard ($colorLines.Count -eq 3) 'Dashboard colour tables.'
    foreach ($line in $colorLines) { . ([scriptblock]::Create($line)) }
    foreach ($name in @('ConvertTo-OrchestratorBrush', 'Set-OrchestratorDashboardView', 'Update-OrchestratorDashboardTimeline')) { . ([scriptblock]::Create((Get-GuiFunction $name).Extent.Text)) }
    $canvas = [System.Windows.Controls.Canvas]::new()
    $controls = @{
        DashboardBannerBorder = [System.Windows.Controls.Border]::new(); DashboardBannerText = [System.Windows.Controls.TextBlock]::new()
        DashboardTilesList = [System.Windows.Controls.ItemsControl]::new(); DashboardActionsGrid = [System.Windows.Controls.DataGrid]::new()
        DashboardActionsGroup = [System.Windows.Controls.GroupBox]::new(); DashboardGrid = [System.Windows.Controls.DataGrid]::new(); DashboardTimelineCanvas = $canvas
    }
    $canvas.Measure([System.Windows.Size]::new(1200, 300)); $canvas.Arrange([System.Windows.Rect]::new(0, 0, 1200, 300))
    Set-OrchestratorDashboardView -Controls $controls -Dashboard $dashboard
    Assert-Dashboard ($controls.DashboardBannerText.Text -eq $dashboard.Banner.Text -and ([string]$controls.DashboardBannerBorder.Background) -eq '#FFFDE7E9') 'Banner text and critical colour.'
    $renderedTiles = @($controls.DashboardTilesList.ItemsSource)
    Assert-Dashboard ($renderedTiles.Count -eq 9 -and @($renderedTiles | Where-Object Label -eq 'Failures 24 h')[0].Background -eq '#FDE7E9' -and @($renderedTiles | Where-Object Label -eq 'Running')[0].Background -eq '#FFFFFF') 'Tile colours by level.'
    Assert-Dashboard ([string]$controls.DashboardActionsGroup.Header -eq ('To review ({0})' -f @($dashboard.Actions).Count) -and @($controls.DashboardGrid.ItemsSource).Count -eq 3) 'Items and servers bound.'
    $bars = @($canvas.Children | Where-Object { $_ -is [System.Windows.Shapes.Rectangle] -and $null -ne $_.ToolTip })
    Assert-Dashboard ($bars.Count -eq @($dashboard.Timeline.Items).Count -and $canvas.Height -eq (22 + 5 * 24 + 4)) "Timeline bars ($($bars.Count)) and lane height."
    Assert-Dashboard (@($bars | Where-Object { [string]$_.Fill -eq '#FF005A9E' }).Count -eq 1 -and @($bars | Where-Object { [string]$_.Fill -eq '#FFF7B955' }).Count -ge 1) 'Running and unassigned bar colours.'
    Assert-Dashboard (@($canvas.Children | Where-Object { $_ -is [System.Windows.Controls.TextBlock] -and $_.Text -eq '14:00' }).Count -eq 1) 'Two-hour ticks.'
    $narrow = [System.Windows.Controls.Canvas]::new(); $narrow.Tag = $dashboard.Timeline
    Update-OrchestratorDashboardTimeline -Canvas $narrow
    Assert-Dashboard (@($narrow.Children).Count -eq 0) 'Not drawn before layout; redrawn on SizeChanged.'
    Assert-Dashboard ($guiText.Contains('$script:Controls.DashboardTimelineCanvas.Add_SizeChanged(')) 'Timeline redraw on resize.'
}
finally { if (Test-Path -LiteralPath $root) { Remove-Item -LiteralPath $root -Recurse -Force } }

"[{0}] PASS: Orchestrator GUI dashboard, {1} offline checks." -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $script:Checks

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCGAB488cAWIxF6
# hR4AZcRBkTZA1ah12NErKdhiS+rgIqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIFDmhBzFLBzSKWYFnePsponvR/dyMrhp/Bo0irsfgLwhMA0GCSqG
# SIb3DQEBAQUABIIBgCbQqtHgi1KgpK2aE3daur4PfH2msyUxDUdCt3EPhhNBBykQ
# VobSTA1ZXOMwY05r9Ckhd/qtkUu+0ZvFHDQT+4iKwqzZi7DHk91E60+xs1YmRlP1
# 5QzARySiIpeu50VuW+Ws6XG9DjapSS/7Oo+KqjFd9U44eO0B6SkkCHYHfFC8sxyA
# wVyI1Rw4vSNuiicbWo5oLcKu+iTFtKSkm/CzrGsxqUclSnogX+ZqJo6+sUxj/5uL
# ojH0XxQciB7dtoqViUjNcvZZKTg2hwS9vLgA3zyOHdVR3EaLuoZO6Ig/TFUwLYWS
# 5rrbnTuSWoO2GLBkbFoOtKXK+5iMMZOzzSWXjhRcDkd2/S2KUdiHG2V46Nv8mE8Z
# IhWunqEESANsdO7ydyEvPPKtG+9AS9qOjvEcM6Uiz+8QWzMByZj72RPgz7N22aOn
# 0zgx/nF1OdeZGGRfiq2v3KcBFeJHadTK7XDpHYcWXqfdAhMNKbxRff0v0qChdyu4
# v2rbBjg4/TLHbiYbrKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxOTI2
# MzBaMC8GCSqGSIb3DQEJBDEiBCBSZyYYvBo/slWB/wFviULkENMYYmqFYCLvzMCs
# F63m5DANBgkqhkiG9w0BAQEFAASCAgBEUlu7f1+cWyZTdc5sfL0Q98/Afxh1wZv5
# JZtoiYXNfWEfnI5eRilc5m8cYhQknTSV14H45vyTmMUcXJFN0qwLyrtHyb5+An8I
# i0+78vKN7yyld697ZZhl+rzZahHqO1wDkij+fhGEqse2f/YpeUf/3dkJ5fbrvinY
# Y+eTc3grhSzwkVyHqTWtMaPV7SHf+uHutD4GHW7wqhZ8Thcj/9oxtFBmzqzEH70R
# Ri4u49ToLWnqsXzTqkrTosHPwSEh/ThXuXAU4IFemuW2Jh8Qr+KhD4UDfpJMee75
# d0svJj8kCyOPJ6tRAr5bfzeya01dkCQeJW5j33zK0qL37zPzdo78RAePrAeZPN6U
# DAxJLiXqKVRb3n9Cme4U2Wqwu9rskaA6Z3BVlM+ed/TZTye9nUwcEmrVC11mqOK+
# e2zc1u1k0GmqmNcuUzKrNjfLxDL9X0T5dlChix243JfWE5fwaaG/tMOt2ds6WwWf
# bUSzAgFmtR4kNHDiLtkIqGtrnrS0zN6NEsbvQVZ4oiLY0Bq3BlZA5o7GZXJ+HDDQ
# wrMsdhO8BEnTYzFno2cL5SmIfwZ0ymQBb5DuZicDvyyZT0N93YtG6Gwb1JDyrRg5
# y8YQPh5OHebCntnksLIzhjcurfv0lXH9+Dm8xEegGx4P5bll/GimrTUvAJoD6QGm
# HPiI9ocvOA==
# SIG # End signature block
