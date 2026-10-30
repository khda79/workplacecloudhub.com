#Requires -Version 5.1
<#
.SYNOPSIS
Exercises the two pilot demos with WPF events and forbidden operational commands.
.VERSION
1.0.9
.NOTES
Run RebootManager on PS5.1 or PS7; Orchestrator requires PS7.
#>
[CmdletBinding()]
param(
    [ValidateSet('RebootManager','Orchestrator')][string]$Application='RebootManager',
    [string]$PreviewFolder='',
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
        $result=. (Join-Path $orchestratorFolder 'SmartM365-Inventory-Orchestrator-GUI.ps1') -Demo -ValidateOnly -Tenant 'DO-NOT-LOAD' -SharedDataFolderPath 'Z:\DO-NOT-READ' -GuiLogFolderPath 'Z:\DO-NOT-WRITE'
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB2lVqM2KNmKybY
# VacXH8qcx1iLZsMGAz/pv1VqgU0TD6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGKXnZkOsdO+2RsLSM77Vj8gZEmlJ1k8HWxKoOkftjhbMA0GCSqG
# SIb3DQEBAQUABIIBgB6gYCEiF0acjQbtlPnR/kMpiutN/1FW8bZb8svikY3tRMFL
# hfu3bKQkyYpTfNkyigkrQA17ORM+qKizP51YxRSzbrU5Np1fSJbEi7y5mLsHeaP2
# AA4e0ZU+qXFnnkTSpFqvG3QsDPBJfR5KHnB2/4gOTBrlhg+ntQ7T0aosWjcs2oLv
# wWVNbNhEiZTRJzjxzxIN9rytryqrO5KbUI+Yu+zKSQzNy8Wr81rYKTZMwC0goU8E
# ImzOCmXpvle/H65vuq/wZZruir6qLWtvv62lBXsRywl2GoFFcvQmbUsBorSluh0q
# y/khBp50d2ObtwmC7772/ziAYyj3CwwA8bY5wy8ltIH+YBvRWq/vb3dCwWo15xym
# +2onSh4EUWa4Gn7uKs68ztmCT4E66lLgKvlwwydCe5y1vdE9tu/tDWa/XhCOOAjg
# 1Uf1jUKTSmJLLefvmvIKlHmsQegMrbj/ElVc2v5+BhhnJIF6GYoyYaZezdLIBW2C
# 39flvRXhZdPqVnv+baGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzAwMDIx
# MDVaMC8GCSqGSIb3DQEJBDEiBCANmPqUCGLpnoHxI56STFsxr+by6oVv/IcA8PBz
# dvlqXzANBgkqhkiG9w0BAQEFAASCAgAHLsW1WaKeJ2xmSDA/XE+z1UMI/Y8FxwTa
# rpFz+bYdqz/z7uAJgwHJb+GinI3SiU5lsHooMYMg6Tnl2u4pfzsvMTA6y7/vaqF3
# 65VISq626FsH81zx8M20KGPkqCEXizLSahts4Wb3kxo9CpBUT724dTI9SuB9E12i
# FOiLg2saBcDj3QWh3oiAITlYsAcRVt69FAZjGgi1BGIIo0XTjP5Tl2SW/LodWQoX
# SgGye09h2qiOlvwPKfsSFBSJC1G0/LRh+Pr4siHnxzXuOcvu5zwsiut2qLZ26bkX
# 5WP65gOYNknnfRANv+eZ0fw+mDJ4Qae6b2MT8a2BarygddEZMDUfBeqF7jzAESay
# F8aNUWgPymw2187lkWp0QJiUVkgA3TybLA1XVBvt/aImJlvCbete+pV+NYBDXsYH
# kLmE3nrtz4Jad2P6/8WwKDfS+LkBfcXKLjLN+mRr0MMMxeHGWpOWJT1XAOgX3kOq
# qw89K+Q7s6CooR8ywD/mWKS/vaGcCxeTnBQ85JHVmBu7O1BVwALkV/AsyJtLdMZP
# WDp5O1jTd7ewrLZmJcS36MJxmL4Ir90smJnHg5ZYkUHxtBJzZDTqht87lfmgXt/J
# DuktjCDe0vuCtxx+JrlKhEjTWcLFTOiua7B+2r+OgVRIJWraVdpHcax2PJCU4u7S
# iD7mAD7L1w==
# SIG # End signature block
