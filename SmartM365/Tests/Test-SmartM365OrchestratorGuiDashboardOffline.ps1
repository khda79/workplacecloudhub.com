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
