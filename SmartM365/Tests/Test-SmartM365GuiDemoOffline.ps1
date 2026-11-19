#Requires -Version 5.1
<#
.SYNOPSIS
Exercises the two pilot demos with WPF events and forbidden operational commands.
.VERSION
1.0.12
.NOTES
Run RebootManager on PS5.1 or PS7; Orchestrator requires PS7.
#>
[CmdletBinding()]
param(
    [ValidateSet('RebootManager','Orchestrator')][string]$Application='RebootManager',
    [string]$PreviewFolder='',
    [ValidateRange(1180,3840)][int]$DemoWindowWidth=1500,
    [ValidateRange(700,2160)][int]$DemoWindowHeight=900,
    [switch]$InteractiveSmoke
)
$ErrorActionPreference='Stop'
if([Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA){throw 'Run with -STA.'}
$script:DemoCheckCount=0
function Assert-Demo {
    param([bool]$Condition,[string]$Message)
    if(-not $Condition){throw $Message}
    $script:DemoCheckCount++
}
function Invoke-DemoClick {
    param($Control)
    $Control.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Button]::ClickEvent))
}
function Export-DemoPreview {
    param($Window,[string]$Name)
    if(-not $PreviewFolder){return}
    [void][IO.Directory]::CreateDirectory($PreviewFolder)
    $visual=$Window.Content
    $width=[int]$Window.Width; $height=[int]$Window.Height-40
    if([string]$Window.SizeToContent -eq 'Height'){
        # Windows that size to their content are captured at their content height.
        $visual.Measure([System.Windows.Size]::new($width,[double]::PositiveInfinity))
        $height=[int][Math]::Ceiling($visual.DesiredSize.Height)
    }
    $visual.Measure([System.Windows.Size]::new($width,$height))
    $visual.Arrange([System.Windows.Rect]::new(0,0,$width,$height))
    $visual.UpdateLayout()
    $visual.Dispatcher.Invoke([action]{},[System.Windows.Threading.DispatcherPriority]::ApplicationIdle)
    $image=[System.Windows.Media.Imaging.RenderTargetBitmap]::new($width,$height,96,96,[System.Windows.Media.PixelFormats]::Pbgra32)
    $background=[System.Windows.Media.DrawingVisual]::new()
    $drawing=$background.RenderOpen()
    $drawing.DrawRectangle($Window.Background,$null,[System.Windows.Rect]::new(0,0,$width,$height))
    $drawing.Close()
    $image.Render($background)
    $image.Render($visual)
    $encoder=[System.Windows.Media.Imaging.PngBitmapEncoder]::new()
    $encoder.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($image))
    $stream=[IO.File]::Create((Join-Path $PreviewFolder ($Name+'.png')))
    try{$encoder.Save($stream)}finally{$stream.Dispose()}
}
function Get-DemoProtectedHashes {
    param([string[]]$Paths)
    $result=@{}
    foreach($path in $Paths){
        $result[$path]=if(Test-Path -LiteralPath $path -PathType Leaf){(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash}else{'ABSENT'}
    }
    return $result
}
# Any escape into platform discovery, auth, management or process launch fails the suite.
function Get-CimInstance {throw 'Forbidden demo device discovery'}
function Get-NetIPAddress {throw 'Forbidden demo IP discovery'}
function Get-ItemProperty {throw 'Forbidden demo registry discovery'}
function Restart-Computer {throw 'Forbidden demo restart'}
function Start-Process {throw 'Forbidden demo process launch'}
function Import-Module {throw 'Forbidden demo runtime module import'}
function Initialize-SmartM365TenantContext {throw 'Forbidden demo tenant context'}
function Start-SmartM365InventoryRun {throw 'Forbidden demo operational lifecycle'}
function Connect-MgGraph {throw 'Forbidden demo authentication'}

$smartRoot=Join-Path $PSScriptRoot '..'
$rebootFolder=Join-Path $smartRoot 'Devices/DeviceRebootManager'
$rebootVersion=[string](Get-Content -LiteralPath (Join-Path $rebootFolder 'SmartM365-DeviceRebootManager.version.json.txt') -Raw | ConvertFrom-Json).PackageVersion
$orchestratorFolder=Join-Path $smartRoot 'SmartInventory/Orchestrator'
$protected=@(
    (Join-Path $rebootFolder 'SmartM365-DeviceRebootManager-GUI.config.json'),
    (Join-Path $env:APPDATA 'SmartM365/DeviceRebootManager/SmartM365-DeviceRebootManager.state.json'),
    (Join-Path $env:APPDATA 'SmartM365/DeviceRebootManager/SmartM365-DeviceRebootManager.preferences.json'),
    (Join-Path $env:APPDATA 'SmartM365/DeviceRebootManager/SmartM365-DeviceRebootManager.log'),
    (Join-Path $orchestratorFolder 'SmartM365-Inventory-Orchestrator.local.json'),
    (Join-Path $orchestratorFolder 'SmartM365-Inventory-Orchestrator.local.json.txt')
)
$before=Get-DemoProtectedHashes $protected
try {
    if($Application -eq 'RebootManager'){
        $result=. (Join-Path $rebootFolder 'SmartM365-DeviceRebootManager-GUI.ps1') -Demo -ValidateOnly -ConfigPath 'Z:\DO-NOT-READ\private.json' -ResetRunState
        Assert-Demo ($script:Ui.ProductVersionText.Text -eq "v$rebootVersion") "Reboot product version is not $rebootVersion."
        Assert-Demo ($script:Window.Title -notlike "*$rebootVersion*") 'Reboot title must not show the version.'
        Assert-Demo ($script:Window.Title -like '*DEMO*') 'Reboot demo identification missing.'
        Assert-Demo ($script:Ui.DeviceIdentityText.Text -like '*CTSO-LAP-*' -and $script:Ui.DeviceIdentityText.Text -notlike "*$env:COMPUTERNAME*") 'Real device identity leaked.'
        Assert-Demo (-not $script:Ui.ContainsKey('ActivityText')) 'Reboot activity log is still shown.'
        $frozen=$script:SecondsRemaining
        Start-Sleep -Milliseconds 1100
        Assert-Demo ($script:SecondsRemaining -eq $frozen) 'Demo countdown is running.'
        foreach($scene in @('Required','Recommended','Postponed')){
            $script:Ui.DemoScenarioCombo.SelectedItem=$scene
            Assert-Demo ($script:ForceMode -eq ($scene -ne 'Recommended')) "Wrong force mode for $scene"
            Assert-Demo ($script:RebootDemoScenario -eq $scene) 'Scenario selector did not update.'
            Assert-Demo (($script:Ui.CountdownPanel.Visibility -eq 'Visible') -eq ($scene -ne 'Recommended')) "Wrong countdown visibility for $scene"
            Export-DemoPreview $script:Window ('reboot-'+$scene.ToLowerInvariant())
        }
        $script:Ui.DemoScenarioCombo.SelectedItem='Required'
        Invoke-DemoClick $script:Ui.RestartButton
        Assert-Demo ($script:Ui.NextActionText.Text -like '*not restarted*') 'Restart simulation did not complete.'
        Invoke-DemoClick $script:Ui.PostponeButton
        Assert-Demo ($script:PostponeCount -eq 1 -and $script:SecondsRemaining -eq 1500) 'Postponement simulation failed.'
        Invoke-DemoClick $script:Ui.PostponeButton
        Assert-Demo (-not $script:Ui.PostponeButton.IsEnabled) 'Postponement limit was not applied.'
        $script:Ui.DemoScenarioCombo.SelectedItem='Recommended'
        Invoke-DemoClick $script:Ui.LaterButton
        Assert-Demo ($script:Ui.NextActionText.Text -like '*shown later*') 'Reminder simulation failed.'
        $script:Ui.LanguageCombo.SelectedItem=@($script:Ui.LanguageCombo.Items | Where-Object Code -EQ 'fr')[0]
        Assert-Demo ($script:SelectedLanguage -eq 'fr') 'Demo language switching failed.'
        Assert-Demo ($script:Window.Title -like '*DEMO*') 'Language switching removed demo identification.'
        $script:Window.Close()
    } else {
        if($PSVersionTable.PSVersion.Major -lt 7){throw 'Orchestrator requires PowerShell 7.'}
        $result=. (Join-Path $orchestratorFolder 'SmartM365-Inventory-Orchestrator-GUI.ps1') -Demo -ValidateOnly -DemoWindowWidth $DemoWindowWidth -DemoWindowHeight $DemoWindowHeight -Tenant 'DO-NOT-LOAD' -SharedDataFolderPath 'Z:\DO-NOT-READ' -GuiLogFolderPath 'Z:\DO-NOT-WRITE'
        $s=$script:OrchestratorDemo; $c=$s.Controls
        $catalog=Get-Content -LiteralPath (Join-Path $orchestratorFolder 'Orchestrator-Jobs.json.template') -Raw|ConvertFrom-Json
        $expectedJobs=$catalog.Jobs.Count
        $expectedHistory=$expectedJobs*7
        $originalTimes=$s.Data.Jobs[0].Times
        Assert-Demo ($result.Mode -eq 'Demo') 'Demo validation did not return its model.'
        Assert-Demo ($s.Data.Jobs.Count -eq $expectedJobs -and $s.Data.Servers.Count -eq 4 -and @($s.Data.Servers | Where-Object Mode -eq 'Inactive').Count -eq 1) 'Sample model incomplete.'
        Assert-Demo (@(Compare-Object @($catalog.Jobs.Name|Sort-Object) @($s.Data.Jobs.Name|Sort-Object)).Count -eq 0) 'Planning does not match the complete shipped catalog.'
        Assert-Demo ($c.PlanningGrid.Items.Count -eq $expectedJobs) 'Planning grid omitted catalog jobs.'
        Assert-Demo (@($s.Data.Jobs|Where-Object {-not $_.Enabled}).Count -eq @($catalog.Jobs|Where-Object {-not $_.Enabled}).Count) 'Disabled catalog jobs are missing.'
        Assert-Demo (@($s.Data.Jobs|Where-Object Name -EQ 'M365-Usage-Inventory')[0].Arguments -eq '-Reports All') 'Catalog arguments were lost.'
        foreach($sourceJob in $catalog.Jobs){
            $sampleJob=@($s.Data.Jobs|Where-Object Name -EQ $sourceJob.Name)[0]
            Assert-Demo ($sampleJob.Enabled -eq $sourceJob.Enabled -and $sampleJob.Times -eq ($sourceJob.Schedule.Times -join ', ') -and $sampleJob.Frequency -eq $sourceJob.Schedule.Type) ('Catalog schedule/state mismatch: '+$sourceJob.Name)
            Assert-Demo ($sampleJob.DependsOn -eq ($sourceJob.DependsOn -join ', ')) ('Catalog dependencies lost: '+$sourceJob.Name)
        }
        Assert-Demo ($s.FilteredHistory.Count -eq $expectedHistory) 'Sample history incomplete.'
        Assert-Demo (@($s.Data.History|Where-Object {[datetime]$_.EventTime -gt $s.Data.Clock}).Count -eq 0) 'History contains future sample executions.'
        Assert-Demo (-not (Test-Path -LiteralPath $s.OutputRoot)) 'Demo wrote files during initialization.'
        Assert-Demo ($c.RequestJobsGrid.Items.Count -eq 3) 'Request detail selection failed.'
        foreach($view in @('Dashboard','Operations','Planning','History','Requests','Servers and versions','Events','Configuration')){
            if($view -eq 'Requests'){
                # Requests is a sub-tab of Operations since GUI 1.6.3.
                Select-OrchestratorDemoTarget 'Requests'
                Assert-Demo ($c.MainTabs.SelectedItem.Header -eq 'Operations' -and $c.OperationsTabs.SelectedItem -eq $c.OperationsRequestsTab -and @($c.MainTabs.Items | Where-Object Header -EQ 'Requests').Count -eq 0) 'Requests is not under Operations.'
            }
            elseif($view -eq 'Servers and versions'){
                # Servers and versions is a sub-tab of Configuration since GUI 1.6.13; Activity is removed.
                Select-OrchestratorDemoTarget 'Servers and versions'
                Assert-Demo ($c.MainTabs.SelectedItem.Header -eq 'Configuration' -and $c.ConfigurationTabs.SelectedItem -eq $c.ConfigurationServersTab -and @($c.MainTabs.Items | Where-Object { $_.Header -in @('Servers and versions','Activity') }).Count -eq 0) 'Servers and versions is not under Configuration.'
                Assert-Demo (@($c.ServersGrid.ItemsSource | Where-Object { $_.Os -and $_.Cpu -and $_.MemoryGB -and $_.SystemDriveSizeGB }).Count -eq @($c.ServersGrid.ItemsSource).Count -and @($c.ServersGrid.ItemsSource | Where-Object SystemDriveLow).Count -eq 1) 'Server system information missing in the demo.'
            }
            else{
                # Servers and versions is captured on its own sub-tab; Configuration shows Settings.
                if($view -eq 'Configuration'){$c.ConfigurationSettingsTab.IsSelected=$true}
                foreach($tab in $c.MainTabs.Items){if($tab.Header -eq $view){$c.MainTabs.SelectedItem=$tab}}
                Assert-Demo ($c.MainTabs.SelectedItem.Header -eq $view) "Cannot navigate to $view"
            }
            Export-DemoPreview $s.Window ('orchestrator-'+($view -replace ' ','-').ToLowerInvariant())
        }
        Assert-Demo (@($c.DashboardTilesList.ItemsSource).Count -eq 9 -and @($c.DashboardGrid.ItemsSource).Count -eq 4 -and @($c.DashboardTimelineCanvas.Tag.Items).Count -gt 0) 'Dashboard tiles, servers or timeline missing.'
        Assert-Demo ($c.DashboardBannerText.Text -like '*DEMO*' -and @($c.DashboardActionsGrid.ItemsSource).Count -eq @($s.Dashboard.Actions).Count) 'Dashboard banner or items to review missing.'
        Select-OrchestratorDemoTarget 'Failures24h'
        Assert-Demo ($c.MainTabs.SelectedItem.Header -eq 'History' -and @($s.FilteredHistory | Where-Object Status -NE 'Failed').Count -eq 0) 'Dashboard failure tile navigation failed.'
        $c.HistoryStatusCombo.SelectedItem='All';Update-OrchestratorDemoHistory
        $tenantRows=@($c.ConfigTenantGrid.ItemsSource)
        Assert-Demo (@($c.ConfigGlobalGrid.ItemsSource).Count -gt 0 -and $tenantRows.Count -gt 0 -and @($c.ConfigMailGrid.ItemsSource).Count -gt 0) 'Configuration grids are empty.'
        Assert-Demo (@($tenantRows | Where-Object { $_.Sensitive -and -not $_.Masked }).Count -eq 0 -and @($tenantRows | Where-Object Parameter -eq 'Thumbprint')[0].Value -notlike 'DEMO0*') 'Sensitive demo values are visible by default.'
        $tenantView=@(@($s.Configuration.Files | Where-Object Kind -eq 'Tenant')[0].Rows)
        Assert-Demo (@($tenantView | Where-Object Model -eq 'Not in template').Count -eq 1 -and @($tenantRows | Where-Object Model -eq 'Not in template').Count -eq 0 -and @($tenantRows | Where-Object Model -eq 'Missing from file').Count -eq 1) 'Template comparison missing in the demo, or keys not in the template are shown.'
        Assert-Demo (@($tenantRows | Where-Object Parameter -eq 'MailClientName')[0].Source -eq 'global' -and @($tenantRows | Where-Object Parameter -eq 'MailClientName')[0].Effective -eq 'Contoso (DEMO)') 'Inherited tenant value not resolved from global.'
        Invoke-DemoClick $c.ConfigShowSensitiveButton
        Assert-Demo (@($c.ConfigTenantGrid.ItemsSource | Where-Object Parameter -eq 'Thumbprint')[0].Value -like 'DEMO0*') 'Show did not reveal sensitive demo values.'
        Invoke-DemoClick $c.ConfigShowSensitiveButton
        $c.ConfigFilterBox.Text='DEMO0'
        Assert-Demo (@($c.ConfigTenantGrid.ItemsSource).Count -eq 0) 'The filter matched a hidden value.'
        $c.ConfigFilterBox.Text='Licences'
        Assert-Demo (@($c.ConfigMailGrid.ItemsSource).Count -eq 2 -and @($c.ConfigMailGrid.ItemsSource | Where-Object Source -eq 'mail').Count -eq 2) 'Configuration filter or mail source failed.'
        $c.ConfigFilterBox.Text=''
        foreach($tab in $c.MainTabs.Items){if($tab.Header -eq 'Configuration'){$c.MainTabs.SelectedItem=$tab}}
        $c.ConfigFileTabs.SelectedIndex=1
        Export-DemoPreview $s.Window 'orchestrator-configuration-tenant'
        $c.ConfigFileTabs.SelectedIndex=0
        $c.ReportCadenceTab.IsSelected=$true
        $c.ReportCadenceGrid.SelectedItem = $c.ReportCadenceGrid.ItemsSource | Where-Object Script -EQ 'SmartM365-Licences-Inventory'
        Assert-Demo ($c.ReportCadenceMode.SelectedValue -eq 'SelectedDays' -and $c.ReportCadenceDays.Children[0].IsChecked) 'Licensing report cadence did not load.'
        Export-DemoPreview $s.Window 'orchestrator-report-mail-cadence'
        Assert-Demo ($c.ReportReplayAudience.SelectedValue -eq 'Administrators' -and $c.ReportReplayButton.IsEnabled) 'Replay defaults to administrators only.'
        Invoke-DemoClick $c.ReportReplayButton
        $c.ReportReplayAudience.SelectedValue='Configured'
        Invoke-DemoClick $c.ReportReplayButton
        Assert-Demo ($c.ReportReplayAudience.SelectedValue -eq 'Configured') 'The configured-recipient choice is unavailable.'
        $c.ReportCadenceDays.Children[0].IsChecked=$false
        Assert-Demo (-not $c.ReportCadenceSave.IsEnabled) 'Empty cadence day selection can be saved.'
        $c.ReportCadenceDays.Children[2].IsChecked=$true
        Invoke-DemoClick $c.ReportCadenceSave
        Assert-Demo ($c.ReportCadenceGrid.SelectedItem.Weekdays -eq 'Wednesday') 'Demo cadence edit did not persist in memory.'
        $c.ReportCadenceMode.SelectedValue='Daily'
        Assert-Demo (-not $c.ReportCadenceDays.IsEnabled -and $c.ReportCadenceSave.IsEnabled) 'Daily mode did not disable day selection.'
        Invoke-DemoClick $c.ReportCadenceSave
        Assert-Demo ($c.ReportCadenceGrid.SelectedItem.Mode -eq 'Daily') 'Demo Daily mode did not save.'
        $c.ConfigurationSettingsTab.IsSelected=$true
        $c.HistoryServerCombo.SelectedItem='CTSO-INV-01'
        Invoke-DemoClick $c.HistoryRefreshButton
        Assert-Demo ($s.FilteredHistory.Count -eq (@($s.Data.Jobs|Where-Object Server -EQ 'CTSO-INV-01').Count*7)) 'History server filter failed.'
        $c.HistoryStatusCombo.SelectedItem='Failed'
        Invoke-DemoClick $c.HistoryRefreshButton
        Assert-Demo (@($s.FilteredHistory | Where-Object Status -NE 'Failed').Count -eq 0) 'History status filter failed.'
        $c.HistoryServerCombo.SelectedItem='All';$c.HistoryStatusCombo.SelectedItem='All'
        Invoke-DemoClick $c.HistoryRefreshButton
        Invoke-DemoClick $c.ExportHtmlButton
        $html=Join-Path $s.OutputRoot 'SmartM365-Orchestrator-History-DEMO.html'
        Assert-Demo (Test-Path -LiteralPath $html) 'Demo HTML export missing.'
        $htmlText=[IO.File]::ReadAllText($html)
        Assert-Demo ($htmlText -like '*DEMO - Sample data*' -and $htmlText -like '*CTSO-INV-01*') 'HTML lacks sample identification or rows.'
        Assert-Demo ([regex]::Matches($htmlText,'<tr>').Count -eq ($expectedHistory+1)) 'HTML report row count mismatch.'
        Invoke-DemoClick $c.ExportCsvButton
        Assert-Demo ((Import-Csv (Join-Path $s.OutputRoot 'SmartM365-Orchestrator-History-DEMO.csv')).Count -eq $expectedHistory) 'CSV report row count mismatch.'
        Invoke-DemoClick $c.EnableMaintenanceButton
        Assert-Demo $s.Data.Maintenance 'Maintenance simulation failed.'
        Invoke-DemoClick $c.DisableMaintenanceButton
        Assert-Demo (-not $s.Data.Maintenance) 'Maintenance resume simulation failed.'
        $c.TimesBox.Text='06:45';Invoke-DemoClick $c.ApplyJobButton
        Assert-Demo ($s.Data.Jobs[0].Times -eq '06:45') 'Job draft simulation failed.'
        Invoke-DemoClick $c.RequestRunButton
        Assert-Demo ($s.Data.Requests.Count -eq 2) 'Request simulation failed.'
        Invoke-DemoClick $c.CancelAllRequestsButton
        Assert-Demo (@($s.Data.Requests | Where-Object Pending -GT 0).Count -eq 0) 'Cancellation simulation failed.'
        $runningBefore=@($s.Data.Running).Count
        Invoke-DemoClick $c.CancelPendingWorkButton
        Assert-Demo (@($s.Data.Pending).Count -eq 0 -and @($s.Data.Running).Count -eq $runningBefore) 'Pending-work cancellation must preserve running sample jobs.'
        Invoke-DemoClick $c.StopRunningJobButton
        Assert-Demo ($s.Data.Running.Count -eq 1) 'Stop simulation failed.'
        Invoke-DemoClick $c.PublishButton
        Assert-Demo ($s.Data.Versions.Count -eq 3) 'Publication simulation failed.'
        Invoke-DemoClick $c.RebalanceButton
        Invoke-DemoClick $c.RollbackButton
        Assert-Demo ($s.Data.Jobs[0].Times -eq $originalTimes -and $s.Data.Jobs.Count -eq $expectedJobs) 'Demo reset failed.'
        Assert-Demo ((New-OrchestratorDemoData -Scenario Issues).Incidents.Count -gt 0) 'Issues scene missing.'
        Assert-Demo (New-OrchestratorDemoData -Scenario Maintenance).Maintenance 'Maintenance scene missing.'
        $healthy=New-OrchestratorDemoData -Scenario Healthy
        Assert-Demo (@($healthy.History | Where-Object Status -NE 'Success').Count -eq 0 -and @($healthy.LogEvents | Where-Object Level -NE 'INFO').Count -eq 0 -and $healthy.Incidents.Count -eq 0 -and -not $healthy.Maintenance) 'Healthy scene contains issues.'
        Assert-Demo ($healthy.Running.Count -gt 0 -and $healthy.Pending.Count -gt 0) 'Healthy scene lost its running and pending samples.'
        $healthyDashboard=New-OrchestratorDemoDashboard -Data $healthy
        Assert-Demo ($healthyDashboard.Banner.Level -eq 'Normal' -and @($healthyDashboard.Actions).Count -eq 0 -and @($healthyDashboard.Tiles | Where-Object Level -In @('Warning','Critical')).Count -eq 0 -and @($healthyDashboard.Tiles | Where-Object Label -EQ 'Failures 24 h')[0].Value -eq '0') 'Healthy dashboard is not clean.'
        $c.DemoScenarioCombo.SelectedItem='Healthy'
        Assert-Demo ($s.Data.Scenario -eq 'Healthy' -and $c.DashboardBannerText.Text -like 'Everything looks healthy. DEMO*' -and @($c.DashboardActionsGrid.ItemsSource).Count -eq 0) 'Healthy selector failed.'
        $c.MainTabs.SelectedIndex=0
        Export-DemoPreview $s.Window 'orchestrator-dashboard-healthy'
        $c.DemoScenarioCombo.SelectedItem='Issues'
        Assert-Demo ($c.OperationsIncidentsGrid.Items.Count -eq 1) 'Issues selector failed.'
        Assert-Demo ($c.EventsGrid.Items.Count -eq 17) 'Issues events missing.'
        $c.EventsLevelCombo.SelectedIndex=2
        Assert-Demo (@($s.FilteredEvents | Where-Object Level -NE 'ERROR').Count -eq 0 -and $s.FilteredEvents.Count -eq 2) 'Events level filter failed.'
        $c.EventsLevelCombo.SelectedIndex=0;$c.EventsServerCombo.SelectedItem='CTSO-INV-03';$c.EventsCategoryCombo.SelectedItem='StopRequest'
        Assert-Demo ($c.EventsGrid.Items.Count -eq 1) 'Events server/category filter failed.'
        $c.EventsCategoryCombo.SelectedIndex=0;$c.EventsServerCombo.SelectedIndex=0;$c.EventsTextBox.Text='heartbeat';$c.EventsPeriodCombo.SelectedIndex=0
        Assert-Demo ($c.EventsGrid.Items.Count -eq 1) 'Events text/period filter failed.'
        $c.EventsTextBox.Text='';$c.EventsPeriodCombo.SelectedIndex=2
        $c.MainTabs.SelectedIndex=7
        Export-DemoPreview $s.Window 'orchestrator-events-issues'
        $c.MainTabs.SelectedIndex=1
        Export-DemoPreview $s.Window 'orchestrator-issues'
        $c.DemoScenarioCombo.SelectedItem='Maintenance'
        Assert-Demo ($s.Data.Maintenance -and -not $c.EnableMaintenanceButton.IsEnabled) 'Maintenance selector failed.'
        if($PreviewFolder){Copy-Item -LiteralPath $html -Destination (Join-Path $PreviewFolder 'orchestrator-history-demo.html')}
        $s.Window.Close()
    }
    if($InteractiveSmoke){
        # Exercise the real interactive entry path, including its six-second splash.
        $script:DemoSmokeObserved=$false
        $script:DemoSmokeApplication=$Application
        $script:DemoSmokeClock=[Diagnostics.Stopwatch]::StartNew()
        $script:DemoSmokeTimer=[System.Windows.Threading.DispatcherTimer]::new()
        $script:DemoSmokeTimer.Interval=[timespan]::FromMilliseconds(250)
        $script:DemoSmokeTimer.Add_Tick({
            $target=if($script:DemoSmokeApplication -eq 'RebootManager'){$script:Window}else{$script:OrchestratorDemo.Window}
            if($target -and $target.IsVisible -and $script:DemoSmokeClock.Elapsed.TotalSeconds -ge 8){
                Export-DemoPreview $target ($script:DemoSmokeApplication.ToLowerInvariant()+'-interactive')
                $script:DemoSmokeObserved=$true
                $script:DemoSmokeTimer.Stop()
                $target.Close()
            }
        })
        $script:DemoSmokeTimer.Start()
        try{
            if($Application -eq 'RebootManager'){
                . (Join-Path $rebootFolder 'SmartM365-DeviceRebootManager-GUI.ps1') -Demo
            }else{
                . (Join-Path $orchestratorFolder 'SmartM365-Inventory-Orchestrator-GUI.ps1') -Demo
            }
        }finally{$script:DemoSmokeTimer.Stop()}
        Assert-Demo $script:DemoSmokeObserved 'The interactive demo window did not become visible.'
    }
    $after=Get-DemoProtectedHashes $protected
    foreach($path in $protected){Assert-Demo ($before[$path] -eq $after[$path]) 'Operational configuration/state/log changed.'}
    '[{0}] PASS: {1} demo, {2} offline WPF and isolation checks.' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Application,$script:DemoCheckCount
} catch {
    [Console]::Error.WriteLine($_.Exception.ToString()+[Environment]::NewLine+$_.ScriptStackTrace)
    exit 1
}


# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAh/fg4QZ/sLzwm
# 5X5PmGZmHqjLfS99olCnnVEnZ1l8faCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgC/tCn1CF3Zum
# bgCSk1/fwEu9WX2Rqq3p4UYjJca3gS0wDQYJKoZIhvcNAQEBBQAEggGAKqfw819O
# anjEooyoScXNQzoct3k4fIYgaOpTqlg/74A+npu3sbrAsFjQhZ4djXSaEWGGGve4
# l0OPHnJkIQrxlo3JRprekETcUR89mhk5UcSBgPIQn3Pxkmj8B3EV+W/A8I3MqUsF
# Kb6dg8S5rUNrSdHeF2cnow0Ldfpc75CnHBbzZ5JCtuOiVgofARqZTyMNEe7yZUDH
# dj2v5sWj4E+Z0ggaNBlLk11RCUfDGizrXa9txunhSQQ8+HZ8/jvGvOi/YsonHyOL
# k0fdeGPzbTrqfx10xxfFlb0c87U7egY9R+unBmObvMoCQtM8/MVw4TEbASRECHQu
# k1Dd/36BR1CkPiEm2oLp+uYH9GkB9s0r0NquzCbP8m5yBsVoc3YdZR13vzYyic2U
# pCtf16lPfuwSE+AaPfj/1cTdxQDqE44VWgyATozgJeBBDqR6Pfxq+0B9lItaTg9G
# qWLQq8wAVjiCr0A/uMjq3RsMYnFwV+szJP/RBF26vXh2ZSEY4DiIWJ9FoYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTE5MDgxODMwWjA/BgkqhkiG9w0BCQQxMgQwcvjyH2N0Gb6dj1uKlXhXXhBz
# ZYXcLzxTmY0xuwQT1IpEP8r3TDK/ES9TShLWjc+8MA0GCSqGSIb3DQEBAQUABIIC
# ACz9wcodY9AxK5KfMWog0whO8NNIsEJtzXtSTW+o6PUfHMof6d8pXKaDXeadWP3I
# ZdPNKZwrDJsZT+M90AG2PoQaoeb7apA+vapf4cKazsBRefWXQ4+ZWWsDwb7yXBS8
# 31qUxxQ0AaB6HzsocCghX1xbOdI/CP9jknvuagQjCGvFViOaj9lmMo1VW72BE+4f
# 3YWLlVBsWBsDf6g7rBDkSv23Zb2sQjy8dvmLVKqrxnmt0EOZQKX3bOIf2xua0UH1
# ql6Tg1Tzu7KypLmf9fC7Qc8RCPqs5Lyvndf/xOvkxJVDu8lTHFid356Z+6OsEvwd
# cP+TydleVOqsExWaH0OiCUCbipFaWLuhHTSS5grIouUyBoDQzUSwp+xSiWOktP1h
# Zlqr1KddozJj9VyKyold1Srs3YTb0WazEVnaQQoXXik+a4mZsn+GnOSALnft86FN
# 5i96Udpkgf1Y0hpNF1+WYGJf8NjeJLKt6ZstjdAlCNyrnOur3OzIq4+BYFoXvAaf
# HUrwLcPhCbDLmA/atTxkeVUuxNsDDhThp2Yy+bags8M6UNeS6ApF/0bxpj17GkKx
# 8ccq3S1bN8w5qqoN7MbXq7MR/4DyZEuL4HdkxbKxlptfJ7j/ObTcQd5d40MfJBCV
# /g1MdVRLx1Drrlmt+0PmWhNrDv9aet5nvp17H3a9YYjJ
# SIG # End signature block
