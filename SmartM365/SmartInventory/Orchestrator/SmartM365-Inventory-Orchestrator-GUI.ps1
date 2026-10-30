#Requires -Version 7.0
<#
.SYNOPSIS
Central WPF management console for the SmartM365 Inventory Orchestrator.

.DESCRIPTION
Edits the shared jobs and cluster configuration with validation, optimistic
concurrency, atomic publication, version history and an audit CSV. It also
shows the live operations of every server, job health, dependency readiness,
the all-server execution history and the current election plan, and submits
job run requests that the resident orchestrators execute.
The Operations tab also runs catalog jobs and manual launches on this computer
(Invoke-SmartM365InventoryJob.ps1 in a new console), sets the execution mode of
the on-premises jobs in the draft, and starts, stops or restarts the orchestrator,
the full collect and the execution summary.

.PARAMETER Tenant
SmartM365 tenant profile. Defaults to test.

.PARAMETER SharedDataFolderPath
Optional direct path to the shared Orchestrator data folder. When omitted,
the path is resolved from the tenant and local Orchestrator configuration.

.PARAMETER GuiLogFolderPath
Optional direct folder for persistent GUI logs. When omitted, logs are written
under LogAllRootPath\SmartM365-Orchestrator-GUI\<computer>.

.PARAMETER ValidateOnly
Parses the XAML, imports the management module and validates the committed
configuration templates without opening a window or touching runtime data.

.PARAMETER SmokeTest
Loads the complete WPF data model without showing the splash or main window.
Intended only for isolated tests with SharedDataFolderPath pointing to a temporary folder.

.VERSION
1.6.33
#>
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$SharedDataFolderPath = '',
    [string]$GuiLogFolderPath = '',
    [switch]$ValidateOnly,
    [switch]$SmokeTest,
    [switch]$Demo,
    [ValidateSet('Overview','Issues','Maintenance','Healthy')]
    [string]$DemoScenario = 'Overview',
    [ValidateSet('Dashboard','Operations','Planning','History','Requests','Servers and versions','Events','Configuration')]
    [string]$DemoView = 'Dashboard',
    [ValidateRange(1180,3840)][int]$DemoWindowWidth = 1500,
    [ValidateRange(720,2160)][int]$DemoWindowHeight = 900
)

# Local lifecycle only; no collection or publication is performed by this guard.
$smartM365RunGuard = $null
if (-not $Demo) {
$smartM365LifecycleRoot=$PSScriptRoot
while ($smartM365LifecycleRoot -and -not (Test-Path -LiteralPath (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1') -PathType Leaf)) {
    $smartM365LifecycleParent=Split-Path $smartM365LifecycleRoot -Parent
    if($smartM365LifecycleParent -eq $smartM365LifecycleRoot){break}
    $smartM365LifecycleRoot=$smartM365LifecycleParent
}
if(-not $smartM365LifecycleRoot){throw 'SmartM365 lifecycle support is missing; update the complete runtime.'}
. (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1')
$smartM365RunGuard=Start-SmartM365InventoryRun -ScriptPath $PSCommandPath -RequireReturn
}
try {
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:AppVersion = '1.6.33'
$script:PrepareJobName = 'WorkplaceEvidence-Prepare'
$script:StartupClock = [Diagnostics.Stopwatch]::StartNew()
$script:Snapshot = $null
$script:DraftJobs = $null
$script:DraftCluster = $null
# History exports: same columns as the grid (Time = start, or scheduled time when the run never started).
$script:HistoryExportColumns = @(@{ Name = 'Time'; Expression = { $_.EventTime } }, 'Scheduled', 'Server', 'JobName', 'Status', 'DurationMinutes', 'ExitCode', 'RetryCount', 'LogPath')
$script:ServerSelectionRestoring = $false
$script:PublishedChangeWarned = $false
# The smoke test publishes a test server only on an explicitly given (copied) shared folder.
$script:SmokePublishAllowed = $PSBoundParameters.ContainsKey('SharedDataFolderPath')
$script:PlanningRows = @()
$script:HistoryRows = @()
$script:InitialHistoryLoaded = $false
$script:RecentRuns = @()
$script:HealthByName = @{}
# SWI validity per job, computed in the Full refresh worker (Get-SmartM365OrchestratorSwiValidityForRoot).
$script:SwiByName = @{}
$script:SwiValidityError = ''
# Job script .VERSION read in the Full refresh worker; the cache keeps unchanged files from being re-read.
$script:ScriptVersions = @{}
$script:ScriptVersionCache = [hashtable]::Synchronized(@{})
$script:RequestRows = @()
$script:MailFolderPath = ''
$script:LogEventsCache = @{}
$script:LogEventsResult = $null
$script:LogEventsLastError = ''
$script:DashboardData = $null
$script:ConfigurationView = $null
$script:ConfigurationShowUntil = [datetime]::MinValue
$script:ConfigurationTimer = $null
$script:DashboardLastError = ''
$script:DashboardLoading = [pscustomobject]@{ Pending = $false; Clock = [Diagnostics.Stopwatch]::new() }
$script:AutoRefreshTimer = $null
$script:CancellationTimer = $null
$script:CancellationBusy = $false
$script:CancellationUiState = @{}
$script:CancellationContext = [pscustomobject]@{
    Worker = $null; Progress = [hashtable]::Synchronized(@{ Message = '' })
    Input = $null; Clock = [Diagnostics.Stopwatch]::new()
}
$script:RequestSubmitTimer = $null
$script:RequestSubmitBusy = $false
$script:RequestSubmitUiState = @{}
$script:RequestSubmitDialog = $null
$script:SendSummaryTimer = $null
$script:SendSummaryRun = $null
$script:SendSummaryTimeoutMinutes = 10
$script:RequestSubmitContext = [pscustomobject]@{
    Worker = $null; Progress = [hashtable]::Synchronized(@{ Message = ''; Stage = ''; LogError = '' })
    Input = $null; Clock = [Diagnostics.Stopwatch]::new()
}
$script:RefreshContext = [pscustomobject]@{
    Worker = $null; Progress = [hashtable]::Synchronized(@{ Message = '' })
    Input = $null; Clock = [Diagnostics.Stopwatch]::new()
}
$script:RefreshTimer = $null
$script:PendingRefreshRequest = $null
$script:LastSuccessfulRefresh = $null
$script:HistoryFailureStatuses = @('Failed', 'TimedOut', 'Interrupted', 'BlockedDependencyFailed', 'BlockedDependencyTimeout', 'BlockedPrerequisite')
$script:Controls = $null
$script:GuiLogPath = ''
$script:GuiLogWriteWarningShown = $false

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SmartM365 Orchestrator" Width="1500" Height="900" WindowState="Maximized"
        MinWidth="1180" MinHeight="720" WindowStartupLocation="CenterScreen"
        Background="#F3F6FA" FontFamily="Segoe UI" UseLayoutRounding="True">
    <Window.Resources>
        <SolidColorBrush x:Key="AccentBrush" Color="#0078D4"/>
        <SolidColorBrush x:Key="DarkBrush" Color="#17324D"/>
        <SolidColorBrush x:Key="MutedBrush" Color="#5D6B78"/>
        <SolidColorBrush x:Key="PanelBrush" Color="White"/>
        <SolidColorBrush x:Key="BorderBrushSoft" Color="#D5DEE8"/>
        <SolidColorBrush x:Key="ControlBorderBrush" Color="#C3CFDC"/>
        <SolidColorBrush x:Key="HoverBrush" Color="#EAF1F8"/>
        <!-- Flat rounded button; a dark overlay gives hover and pressed states to any background colour. -->
        <Style TargetType="Button">
            <Setter Property="Padding" Value="14,7"/>
            <Setter Property="Margin" Value="4,0"/>
            <Setter Property="MinHeight" Value="32"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="Foreground" Value="{StaticResource DarkBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource ControlBorderBrush}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="HorizontalContentAlignment" Value="Center"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Grid>
                            <Border x:Name="Chrome" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6"/>
                            <Border x:Name="Overlay" Background="#17324D" Opacity="0" CornerRadius="6"/>
                            <ContentPresenter Margin="{TemplateBinding Padding}" HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="{TemplateBinding VerticalContentAlignment}" RecognizesAccessKey="True"/>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Overlay" Property="Opacity" Value="0.07"/></Trigger>
                            <Trigger Property="IsPressed" Value="True"><Setter TargetName="Overlay" Property="Opacity" Value="0.15"/></Trigger>
                            <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Chrome" Property="BorderBrush" Value="{StaticResource AccentBrush}"/></Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.45"/><Setter Property="Cursor" Value="Arrow"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Margin" Value="0,3,0,8"/>
            <Setter Property="BorderBrush" Value="{StaticResource ControlBorderBrush}"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Padding" Value="5,3"/>
            <Setter Property="Margin" Value="0,3,0,8"/>
            <Setter Property="BorderBrush" Value="{StaticResource ControlBorderBrush}"/>
        </Style>
        <!-- Tabs: header row on a separator line, content on a rounded white card. -->
        <Style TargetType="TabControl">
            <Setter Property="Padding" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabControl">
                        <Grid KeyboardNavigation.TabNavigation="Local">
                            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                            <Border BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="0,0,0,1" Margin="0,0,0,8">
                                <TabPanel x:Name="HeaderPanel" IsItemsHost="True" Background="Transparent" KeyboardNavigation.TabIndex="1"/>
                            </Border>
                            <Border Grid.Row="1" Background="White" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8">
                                <ContentPresenter x:Name="PART_SelectedContentHost" ContentSource="SelectedContent" Margin="{TemplateBinding Padding}"/>
                            </Border>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="TabItem">
            <Setter Property="Foreground" Value="Black"/>
            <Setter Property="Padding" Value="14,7"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="TabBorder" Background="Transparent" BorderBrush="Transparent" BorderThickness="0,0,0,2" CornerRadius="6,6,0,0" Margin="0,0,2,0" Cursor="Hand">
                            <ContentPresenter x:Name="TabHeader" ContentSource="Header" Margin="{TemplateBinding Padding}" HorizontalAlignment="Center" VerticalAlignment="Center" RecognizesAccessKey="True" TextElement.Foreground="{StaticResource MutedBrush}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="{StaticResource HoverBrush}"/>
                                <Setter TargetName="TabHeader" Property="TextElement.Foreground" Value="{StaticResource DarkBrush}"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
                                <Setter TargetName="TabBorder" Property="Background" Value="White"/>
                                <Setter TargetName="TabHeader" Property="TextElement.Foreground" Value="{StaticResource AccentBrush}"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.5"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="DataGrid">
            <Setter Property="AutoGenerateColumns" Value="False"/>
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="CanUserAddRows" Value="False"/>
            <Setter Property="CanUserDeleteRows" Value="False"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrushSoft}"/>
            <Setter Property="AlternatingRowBackground" Value="#F7FAFD"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#E6EDF4"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="RowHeaderWidth" Value="0"/>
        </Style>
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#F3F6FA"/>
            <Setter Property="Foreground" Value="{StaticResource DarkBrush}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrushSoft}"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
        </Style>
        <!-- Frames: rounded card with a bold header. -->
        <Style TargetType="GroupBox">
            <Setter Property="Margin" Value="0,0,0,12"/>
            <Setter Property="Padding" Value="12,6,12,12"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrushSoft}"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="GroupBox">
                        <Border Background="White" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="8">
                            <DockPanel>
                                <ContentPresenter DockPanel.Dock="Top" ContentSource="Header" Margin="12,10,12,0" RecognizesAccessKey="True" TextElement.FontWeight="SemiBold" TextElement.Foreground="{StaticResource DarkBrush}"/>
                                <ContentPresenter Margin="{TemplateBinding Padding}"/>
                            </DockPanel>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="18">
        <Grid.RowDefinitions>
            <RowDefinition Height="88"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="44"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="{StaticResource DarkBrush}" CornerRadius="10" Padding="18,10">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="78"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Image x:Name="HeaderLogo" Width="62" Height="62" Stretch="Uniform"/>
                <StackPanel Grid.Column="1" VerticalAlignment="Center" Margin="12,0">
                    <TextBlock Text="SmartM365 Orchestrator" Foreground="White" FontSize="25" FontWeight="SemiBold"/>
                    <TextBlock x:Name="SharedPathText" Foreground="#DCEBFA" FontSize="12" TextTrimming="CharacterEllipsis"/>
                </StackPanel>
                <StackPanel Grid.Column="2" VerticalAlignment="Center">
                    <TextBlock x:Name="ConnectionText" Foreground="White" HorizontalAlignment="Right" FontWeight="SemiBold"/>
                    <TextBlock x:Name="LastRefreshText" Foreground="#DCEBFA" HorizontalAlignment="Right" FontSize="12"/>
                </StackPanel>
            </Grid>
        </Border>

        <Grid Grid.Row="1" Margin="0,10,0,8">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel VerticalAlignment="Center">
                <TextBlock x:Name="StatusText" Foreground="{StaticResource MutedBrush}" Text="Ready"/>
                <TextBlock x:Name="MaintenanceBannerText" Foreground="#92400E" FontWeight="SemiBold" TextWrapping="Wrap" Text="Maintenance: checking shared control..."/>
            </StackPanel>
            <StackPanel Grid.Column="1" Orientation="Horizontal">
                <CheckBox x:Name="AutoRefreshCheck" Content="Auto-refresh (60 s)" VerticalAlignment="Center" Margin="4,0,10,0" IsChecked="True"/>
                <Button x:Name="RefreshButton" Content="Refresh"/>
                <Button x:Name="RebalanceButton" Content="Rebalance now" Background="#FFF4CE"/>
                <Button x:Name="RequestPrepareButton" Content="Refresh Power BI data" Background="#FFF4CE" IsEnabled="False" ToolTip="Asks the orchestrators to run WorkplaceEvidence-Prepare once with the current data (published configuration; the inventory jobs it depends on are not rerun)."/>
                <Button x:Name="PublishButton" Content="Publish changes" Background="{StaticResource AccentBrush}" BorderBrush="{StaticResource AccentBrush}" Foreground="White" FontWeight="SemiBold"/>
            </StackPanel>
        </Grid>

        <TabControl x:Name="MainTabs" Grid.Row="2">
            <TabItem Header="Dashboard">
                <Grid Margin="12">
                    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="98"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                    <Border x:Name="DashboardBannerBorder" Background="#F3F6FA" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="7" Padding="12,8" Margin="4,0,4,8">
                        <TextBlock x:Name="DashboardBannerText" FontWeight="SemiBold" TextWrapping="Wrap" Text="Loading dashboard..."/>
                    </Border>
                    <ItemsControl x:Name="DashboardTilesList" Grid.Row="1" Margin="0,0,0,8" ToolTip="Click a tile to open the related view.">
                        <ItemsControl.ItemsPanel><ItemsPanelTemplate><UniformGrid Rows="1"/></ItemsPanelTemplate></ItemsControl.ItemsPanel>
                        <ItemsControl.ItemTemplate>
                            <DataTemplate>
                                <Border Background="{Binding Background}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="7" Margin="4" Padding="10,6" Cursor="Hand" ToolTip="{Binding Detail}">
                                    <StackPanel>
                                        <TextBlock Text="{Binding Label}" Foreground="{StaticResource MutedBrush}" TextTrimming="CharacterEllipsis"/>
                                        <TextBlock Text="{Binding Value}" FontSize="24" FontWeight="SemiBold"/>
                                        <TextBlock Text="{Binding Detail}" Foreground="{StaticResource MutedBrush}" FontSize="11" TextTrimming="CharacterEllipsis"/>
                                    </StackPanel>
                                </Border>
                            </DataTemplate>
                        </ItemsControl.ItemTemplate>
                    </ItemsControl>
                    <Grid Grid.Row="2">
                        <Grid.ColumnDefinitions><ColumnDefinition Width="2*"/><ColumnDefinition Width="12"/><ColumnDefinition Width="3*"/></Grid.ColumnDefinitions>
                        <GroupBox x:Name="DashboardActionsGroup" Header="To review" Margin="4,0,0,8">
                            <DataGrid x:Name="DashboardActionsGrid" SelectionMode="Single" ToolTip="Double-click a row to open the related view.">
                                <DataGrid.RowStyle>
                                    <Style TargetType="DataGridRow">
                                        <Style.Triggers>
                                            <DataTrigger Binding="{Binding Severity}" Value="Critical"><Setter Property="Background" Value="#FDE7E9"/></DataTrigger>
                                            <DataTrigger Binding="{Binding Severity}" Value="Warning"><Setter Property="Background" Value="#FFF4CE"/></DataTrigger>
                                            <DataTrigger Binding="{Binding Severity}" Value="Info"><Setter Property="Background" Value="#E5F1FB"/></DataTrigger>
                                        </Style.Triggers>
                                    </Style>
                                </DataGrid.RowStyle>
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Severity" Binding="{Binding Severity}" Width="70"/>
                                    <DataGridTextColumn Header="Area" Binding="{Binding Area}" Width="135"/>
                                    <DataGridTextColumn Header="Item" Binding="{Binding Item}" Width="190"/>
                                    <DataGridTextColumn Header="Detail" Binding="{Binding Detail}" Width="*"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </GroupBox>
                        <GroupBox Grid.Column="2" Header="Servers" Margin="0,0,4,8">
                            <DataGrid x:Name="DashboardGrid">
                                <DataGrid.RowStyle>
                                    <Style TargetType="DataGridRow">
                                        <Style.Triggers>
                                            <DataTrigger Binding="{Binding State}" Value="Stale"><Setter Property="Background" Value="#FDE7E9"/></DataTrigger>
                                            <DataTrigger Binding="{Binding State}" Value="No heartbeat"><Setter Property="Background" Value="#FDE7E9"/></DataTrigger>
                                            <DataTrigger Binding="{Binding State}" Value="Recycling"><Setter Property="Background" Value="#E5F1FB"/></DataTrigger>
                                        </Style.Triggers>
                                    </Style>
                                </DataGrid.RowStyle>
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Server" Binding="{Binding Server}" Width="125"/>
                                    <DataGridTextColumn Header="State" Binding="{Binding State}" Width="Auto"/>
                                    <DataGridTextColumn Header="Version" Binding="{Binding Version}" Width="Auto"/>
                                    <DataGridTextColumn Header="CPU %" Binding="{Binding CpuPercent}" Width="Auto"/>
                                    <DataGridTextColumn Header="RAM %" Binding="{Binding MemoryPercent}" Width="Auto"/>
                                    <DataGridTextColumn Header="Running" Binding="{Binding Running}" Width="Auto"/>
                                    <DataGridTextColumn Header="Pending" Binding="{Binding Pending}" Width="Auto"/>
                                    <DataGridTextColumn Header="Recycle in" Binding="{Binding RecycleIn}" Width="Auto"/>
                                    <DataGridTextColumn Header="Elected jobs" Binding="{Binding Owned}" Width="Auto"/>
                                    <DataGridTextColumn Header="Certificate" Binding="{Binding Certificate}" Width="Auto"/>
                                    <DataGridTextColumn Header="State persistence" Binding="{Binding StatePersistence}" Width="Auto" MaxWidth="160"/>
                                    <DataGridTextColumn Header="Ready capabilities" Binding="{Binding Capabilities}" Width="*" MinWidth="160"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </GroupBox>
                    </Grid>
                    <GroupBox Grid.Row="3" Header="Next 24 hours by server (elected owner, average duration; dark blue = running, orange = no owner, grey = suspended by maintenance)" Margin="4,0,4,0">
                        <Canvas x:Name="DashboardTimelineCanvas" Height="110" ClipToBounds="True" Background="White"/>
                    </GroupBox>
                    <GroupBox Grid.Row="4" Header="CPU and memory, last 12 hours (blue = CPU, green = RAM; sampled every 5 minutes by each orchestrator)" Margin="4,8,4,0">
                        <Canvas x:Name="DashboardMetricsCanvas" Height="130" ClipToBounds="True" Background="White"/>
                    </GroupBox>
                    <Border x:Name="DashboardLoadingOverlay" Grid.RowSpan="5" Background="#B3F3F6FA" Visibility="Collapsed">
                        <Border Background="White" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="10" Padding="28,20" HorizontalAlignment="Center" VerticalAlignment="Center">
                            <Border.Effect><DropShadowEffect BlurRadius="16" ShadowDepth="2" Opacity="0.25"/></Border.Effect>
                            <StackPanel Width="320">
                                <TextBlock Text="Loading dashboard..." FontSize="16" FontWeight="SemiBold" HorizontalAlignment="Center"/>
                                <ProgressBar IsIndeterminate="True" Height="6" Margin="0,12,0,8"/>
                                <TextBlock x:Name="DashboardLoadingText" Foreground="{StaticResource MutedBrush}" HorizontalAlignment="Center" TextAlignment="Center" TextWrapping="Wrap" Text="Reading shared orchestrator data in the background. The window stays usable."/>
                            </StackPanel>
                        </Border>
                    </Border>
                </Grid>
            </TabItem>

            <TabItem Header="Operations">
                <Grid Margin="12,8,12,12">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                        <TextBlock Text="Orchestrator on this computer" VerticalAlignment="Center" FontWeight="SemiBold" Margin="0,0,10,0"/>
                        <Button x:Name="OrchestratorStartButton" Content="Start" ToolTip="Starts a resident orchestrator in a new console on this computer."/>
                        <Button x:Name="OrchestratorStopButton" Content="Stop" ToolTip="Asks the resident orchestrator of this computer to stop cleanly."/>
                        <Button x:Name="OrchestratorRestartButton" Content="Restart" ToolTip="Stops the orchestrator cleanly, then starts its scheduled task."/>
                        <Button x:Name="FullCollectButton" Content="Full collect" Background="#FFF4CE" ToolTip="Requests the full pipeline from the resident orchestrators and follows it in a new console."/>
                        <Button x:Name="SendSummaryButton" Content="Send summary" ToolTip="Sends the execution summary email now."/>
                    </StackPanel>
                    <TabControl x:Name="OperationsTabs" Grid.Row="1">
                        <TabItem Header="Live">
                        <Grid Margin="12">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="150"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <!-- Maintenance text left, reason and buttons right on the same rows (GUI 1.6.29). -->
                            <Grid Margin="0,0,0,8">
                                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                <TextBlock x:Name="MaintenanceDetailText" TextWrapping="Wrap" VerticalAlignment="Center" Margin="0,0,12,0"/>
                                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                                    <TextBlock Text="Reason" VerticalAlignment="Center" Margin="0,0,8,0"/>
                                    <TextBox x:Name="MaintenanceReasonBox" Width="300" MaxLength="1000" ToolTip="Required to enable maintenance. Does not publish the configuration draft."/>
                                    <Button x:Name="EnableMaintenanceButton" Content="Enable maintenance" Background="#FFF4CE" IsEnabled="False"/>
                                    <Button x:Name="DisableMaintenanceButton" Content="Disable maintenance" IsEnabled="False"/>
                                </StackPanel>
                            </Grid>
                            <DataGrid x:Name="OperationsServersGrid" Grid.Row="1">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Server" Binding="{Binding Server}" Width="160"/>
                                    <DataGridTextColumn Header="State" Binding="{Binding State}" Width="100"/>
                                    <DataGridTextColumn Header="Jobs" Binding="{Binding Mode}" Width="110">
                                        <DataGridTextColumn.ElementStyle>
                                            <Style TargetType="TextBlock">
                                                <Setter Property="ToolTip" Value="Active = takes jobs. Inactive = orchestrator running, no job launched (Configuration &gt; Servers and versions)."/>
                                                <Style.Triggers><DataTrigger Binding="{Binding Mode}" Value="Inactive"><Setter Property="Foreground" Value="#8A5A00"/><Setter Property="FontWeight" Value="SemiBold"/></DataTrigger></Style.Triggers>
                                            </Style>
                                        </DataGridTextColumn.ElementStyle>
                                    </DataGridTextColumn>
                                    <DataGridTextColumn Header="Heartbeat age (min)" Binding="{Binding HeartbeatAgeMinutes}" Width="140"/>
                                    <DataGridTextColumn Header="Version" Binding="{Binding Version}" Width="80"/>
                                    <DataGridTextColumn Header="Maintenance" Binding="{Binding Maintenance}" Width="100"/>
                                    <DataGridTextColumn Header="Acknowledgement" Binding="{Binding MaintenanceAcknowledgement}" Width="145"/>
                                    <DataGridTextColumn Header="Running" Binding="{Binding Running}" Width="70"/>
                                    <DataGridTextColumn Header="Pending" Binding="{Binding Pending}" Width="70"/>
                                    <DataGridTextColumn Header="Recycle in" Binding="{Binding RecycleIn}" Width="90"/>
                                    <DataGridTextColumn Header="PID" Binding="{Binding Pid}" Width="70"/>
                                    <DataGridTextColumn Header="DATA free %" Binding="{Binding DataFreePercent}" Width="90"/>
                                    <DataGridTextColumn Header="Disk space" Binding="{Binding DiskSpace}" Width="120"/>
                                    <DataGridTextColumn Header="State persistence" Binding="{Binding StatePersistence}" Width="*"/>
                                </DataGrid.Columns>
                            </DataGrid>
                            <!-- Running and pending jobs stacked at full width (GUI 1.6.21). -->
                            <Grid Grid.Row="2" Margin="0,10,0,0">
                                <Grid.RowDefinitions><RowDefinition Height="*"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                                <GroupBox Margin="0,0,0,6">
                                    <GroupBox.Header>
                                        <!-- Stop reason on the right of the title (GUI 1.6.29). -->
                                        <Grid>
                                            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                            <TextBlock Text="Running jobs" VerticalAlignment="Center"/>
                                            <StackPanel Grid.Column="1" Orientation="Horizontal" TextElement.FontWeight="Normal">
                                            <TextBlock Text="Stop reason" VerticalAlignment="Center" Margin="0,0,6,0" FontWeight="Normal"/>
                                            <TextBox x:Name="JobStopReasonBox" Width="260" MaxLength="1000" ToolTip="Required. Select one running job, then request its owner to stop it."/>
                                            <Button x:Name="StopRunningJobButton" Content="Stop selected job" IsEnabled="False" Margin="8,0,0,0" Background="#FDE7E9" FontWeight="Normal" ToolTip="The owning orchestrator verifies the selected process and records the stop. The job is not retried automatically."/>
                                            </StackPanel>
                                        </Grid>
                                    </GroupBox.Header>
                                    <Grid>
                                        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                                    <DataGrid x:Name="OperationsRunningGrid" Grid.Row="1">
                                        <DataGrid.RowStyle>
                                            <Style TargetType="DataGridRow">
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding Activity}" Value="Stalled"><Setter Property="Background" Value="#FDE7E9"/></DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </DataGrid.RowStyle>
                                        <DataGrid.Columns>
                                            <DataGridTextColumn Header="Server" Binding="{Binding Server}" Width="120"/>
                                            <!-- Server, Job and Server state share their widths with Pending jobs so the columns align (GUI 1.6.23). -->
                                            <DataGridTextColumn Header="Job" Binding="{Binding Job}" Width="320"/>
                                            <DataGridTextColumn Header="Server state" Binding="{Binding ServerState}" Width="85"/>
                                            <DataGridTextColumn Header="PID" Binding="{Binding Pid}" Width="55"/>
                                            <DataGridTextColumn Header="Started" Binding="{Binding Started}" Width="110"/>
                                            <DataGridTextColumn Header="Duration (min)" Binding="{Binding DurationMinutes}" Width="95"/>
                                            <DataGridTextColumn Header="Est. remaining" Binding="{Binding EstimatedRemaining}" Width="130">
                                                <DataGridTextColumn.ElementStyle>
                                                    <Style TargetType="TextBlock"><Setter Property="ToolTip" Value="{Binding EstimateSource}"/></Style>
                                                </DataGridTextColumn.ElementStyle>
                                            </DataGridTextColumn>
                                            <DataGridTextColumn Header="Est. end" Binding="{Binding EstimatedEnd}" Width="150">
                                                <DataGridTextColumn.ElementStyle>
                                                    <Style TargetType="TextBlock"><Setter Property="ToolTip" Value="{Binding EstimateSource}"/></Style>
                                                </DataGridTextColumn.ElementStyle>
                                            </DataGridTextColumn>
                                            <DataGridTextColumn Header="Activity" Binding="{Binding Activity}" Width="65"/>
                                            <DataGridTextColumn Header="Silent (min)" Binding="{Binding SilentMinutes}" Width="80"/>
                                            <DataGridTextColumn Header="CPU (s)" Binding="{Binding CpuSeconds}" Width="65"/>
                                            <DataGridTextColumn Header="CPU s/min" Binding="{Binding CpuSecondsPerMinute}" Width="*" MinWidth="70"/>
                                        </DataGrid.Columns>
                                    </DataGrid>
                                    </Grid>
                                </GroupBox>
                                <GroupBox Grid.Row="1" Header="Pending jobs (why they wait)" Margin="0,6,0,0">
                                    <DataGrid x:Name="OperationsPendingGrid">
                                        <DataGrid.Columns>
                                            <!-- Same leading columns and widths as Running jobs: Server, Job, Server state (GUI 1.6.23). -->
                                            <DataGridTextColumn Header="Server" Binding="{Binding Server}" Width="120"/>
                                            <DataGridTextColumn Header="Job" Binding="{Binding Job}" Width="320"/>
                                            <DataGridTextColumn Header="Server state" Binding="{Binding ServerState}" Width="85"/>
                                            <DataGridTextColumn Header="Reason" Binding="{Binding Reason}" Width="125"/>
                                            <DataGridTextColumn Header="Waiting (min)" Binding="{Binding WaitingMinutes}" Width="85"/>
                                            <DataGridTextColumn Header="Est. duration (min)" Binding="{Binding EstimatedMinutes}" Width="120">
                                                <DataGridTextColumn.ElementStyle>
                                                    <Style TargetType="TextBlock"><Setter Property="ToolTip" Value="{Binding EstimateSource}"/></Style>
                                                </DataGridTextColumn.ElementStyle>
                                            </DataGridTextColumn>
                                            <DataGridTextColumn Header="Details" Binding="{Binding Details}" Width="*">
                                                <DataGridTextColumn.ElementStyle>
                                                    <Style TargetType="TextBlock">
                                                        <Setter Property="TextWrapping" Value="Wrap"/>
                                                        <Setter Property="ToolTip" Value="{Binding Details}"/>
                                                    </Style>
                                                </DataGridTextColumn.ElementStyle>
                                            </DataGridTextColumn>
                                        </DataGrid.Columns>
                                    </DataGrid>
                                </GroupBox>
                            </Grid>
                        </Grid>
                        </TabItem>
                        <TabItem Header="Jobs">
                            <Grid Margin="8">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="*"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="150"/>
                                </Grid.RowDefinitions>
                                <StackPanel Margin="0,0,0,8">
                                    <StackPanel Orientation="Horizontal">
                                        <TextBlock Text="On-premises execution mode (OnPremExecutionMode):" VerticalAlignment="Center" Margin="0,0,6,0"/>
                                        <TextBlock x:Name="OnPremModeText" VerticalAlignment="Center" FontWeight="SemiBold"/>
                                    </StackPanel>
                                    <StackPanel Orientation="Horizontal" Margin="0,6,0,0">
                                        <TextBlock Text="Mode of the selected on-premises job" VerticalAlignment="Center" Margin="0,0,6,0"/>
                                        <ComboBox x:Name="JobExecutionModeCombo" Width="130" IsEnabled="False">
                                            <ComboBoxItem Content="(default)"/>
                                            <ComboBoxItem Content="LocalCache"/>
                                            <ComboBoxItem Content="Direct"/>
                                            <ComboBoxItem Content="Auto"/>
                                        </ComboBox>
                                        <Button x:Name="ApplyJobExecutionModeButton" Content="Apply to draft" IsEnabled="False"/>
                                        <Button x:Name="ApplyModeToAllOnPremButton" Content="Apply to all on-premises jobs"/>
                                        <TextBlock Text="Then publish the changes. (default) follows OnPremExecutionMode." VerticalAlignment="Center" Foreground="{StaticResource MutedBrush}" Margin="8,0,0,0"/>
                                    </StackPanel>
                                </StackPanel>
                                <DataGrid x:Name="LaunchJobsGrid" Grid.Row="1">
                                    <DataGrid.RowStyle>
                                        <Style TargetType="DataGridRow">
                                            <Style.Triggers>
                                                <DataTrigger Binding="{Binding SwiBlocks}" Value="True"><Setter Property="Background" Value="#FDE7E9"/></DataTrigger>
                                            </Style.Triggers>
                                        </Style>
                                    </DataGrid.RowStyle>
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Job" Binding="{Binding Name}" Width="290"/>
                                        <DataGridTextColumn Header="Version" Binding="{Binding Version}" Width="70"/>
                                        <DataGridTextColumn Header="Group" Binding="{Binding Group}" Width="90"/>
                                        <DataGridCheckBoxColumn Header="Enabled" Binding="{Binding Enabled}" Width="65"/>
                                        <DataGridTextColumn Header="Runtime" Binding="{Binding Runtime}" Width="130"/>
                                        <DataGridTextColumn Header="Mode" Binding="{Binding Mode}" Width="85"/>
                                        <DataGridTextColumn Header="Mode source" Binding="{Binding ModeSource}" Width="180"/>
                                        <DataGridTextColumn Header="Draft mode" Binding="{Binding DraftMode}" Width="90"/>
                                        <DataGridTextColumn Header="Last success" Binding="{Binding LastSuccess}" Width="125"/>
                                        <DataGridTextColumn Header="SWI validity" Binding="{Binding Swi}" Width="150">
                                            <DataGridTextColumn.ElementStyle>
                                                <Style TargetType="TextBlock">
                                                    <Setter Property="ToolTip" Value="{Binding SwiDetail}"/>
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding SwiBlocks}" Value="True"><Setter Property="Foreground" Value="#A4262C"/><Setter Property="FontWeight" Value="SemiBold"/></DataTrigger>
                                                </Style.Triggers>
                                                </Style>
                                            </DataGridTextColumn.ElementStyle>
                                        </DataGridTextColumn>
                                        <DataGridTextColumn Header="SWI days left" Binding="{Binding SwiDaysLeft}" Width="85"/>
                                        <DataGridTextColumn Header="Cache" Binding="{Binding CacheName}" Width="*"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                                <StackPanel Grid.Row="2" Orientation="Horizontal" Margin="0,8,0,6">
                                    <Button x:Name="RunJobNowButton" Content="Run now on this computer" IsEnabled="False" ToolTip="Runs the job in a new console on this computer, outside the orchestrators (no election, lock or dependency check)."/>
                                    <Button x:Name="RunJobAsServiceAccountButton" Content="Run now on this computer with service account" IsEnabled="False" ToolTip="Same as Run now, started with runas under the account of this computer's orchestrator scheduled task. The runas console asks for the account password; the GUI never sees it."/>
                                    <Button x:Name="RequestJobButton" Content="Request (orchestrators)" Background="#FFF4CE" IsEnabled="False" ToolTip="Submits the job to the resident orchestrators: elected server, shared locks and dependency rules."/>
                                    <CheckBox x:Name="LaunchIncludeDependenciesCheck" Content="Also run the enabled dependencies (Request)" VerticalAlignment="Center" Margin="10,0,0,0"/>
                                </StackPanel>
                                <TextBox x:Name="LaunchJobCommandBox" Grid.Row="3" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" FontFamily="Consolas"/>
                            </Grid>
                        </TabItem>
                        <TabItem x:Name="OperationsRequestsTab" Header="Requests">
                            <Grid Margin="8">
                                <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                                <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                                    <TextBlock Text="Cancellation reason" VerticalAlignment="Center" Margin="0,0,8,0"/>
                                    <TextBox x:Name="CancellationReasonBox" Width="340" Margin="0,0,8,0"/>
                                    <Button x:Name="CancelRequestButton" Content="Cancel remaining jobs" IsEnabled="False" ToolTip="Cancel pending jobs and retries in the selected request. Running collectors continue. All residents must support cancellation; an offline resident applies it when it returns."/>
                                    <Button x:Name="CancelAllRequestsButton" Content="Cancel All remaining Jobs" IsEnabled="False" Margin="8,0,0,0" ToolTip="Cancel pending jobs and retries in all active pipeline requests for this tenant. Running collectors and automatic schedules are unchanged."/>
                                </StackPanel>
                                <TextBlock x:Name="CancellationProgressText" Grid.Row="1" Margin="0,0,0,8" Foreground="#5D6B78" TextWrapping="Wrap"/>
                                <GroupBox Grid.Row="2" Header="Pipeline and job run requests (newest first)">
                                    <DataGrid x:Name="RequestsGrid" SelectionMode="Single">
                                        <DataGrid.Columns>
                                            <DataGridTextColumn Header="Batch" Binding="{Binding BatchId}" Width="260"/>
                                            <DataGridTextColumn Header="Selection" Binding="{Binding Selection}" Width="90"/>
                                            <DataGridTextColumn Header="Created" Binding="{Binding Created}" Width="125"/>
                                            <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="150"/>
                                            <DataGridTextColumn Header="Jobs" Binding="{Binding Jobs}" Width="55"/>
                                            <DataGridTextColumn Header="Pending" Binding="{Binding Pending}" Width="65"/>
                                            <DataGridTextColumn Header="Failed" Binding="{Binding Failed}" Width="60"/>
                                            <DataGridTextColumn Header="Cancelled" Binding="{Binding Cancelled}" Width="70"/>
                                            <DataGridTextColumn Header="Requested by" Binding="{Binding RequestedBy}" Width="*"/>
                                        </DataGrid.Columns>
                                    </DataGrid>
                                </GroupBox>
                                <GroupBox Grid.Row="3" Header="Jobs of the selected request">
                                    <DataGrid x:Name="RequestJobsGrid">
                                        <DataGrid.Columns>
                                            <DataGridTextColumn Header="Job" Binding="{Binding JobName}" Width="260"/>
                                            <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="170"/>
                                            <DataGridTextColumn Header="Server" Binding="{Binding OwnerServer}" Width="140"/>
                                            <DataGridTextColumn Header="Updated (UTC)" Binding="{Binding UpdatedAtUtc}" Width="170"/>
                                            <DataGridTextColumn Header="Detail" Binding="{Binding Detail}" Width="*"/>
                                        </DataGrid.Columns>
                                    </DataGrid>
                                </GroupBox>
                            </Grid>
                        </TabItem>
                        <TabItem Header="Manual launches">
                            <Grid Margin="8">
                                <Grid.RowDefinitions>
                                    <RowDefinition Height="*"/>
                                    <RowDefinition Height="Auto"/>
                                    <RowDefinition Height="150"/>
                                </Grid.RowDefinitions>
                                <DataGrid x:Name="ManualLaunchesGrid">
                                    <DataGrid.Columns>
                                        <DataGridTextColumn Header="Launch" Binding="{Binding Name}" Width="320"/>
                                        <DataGridTextColumn Header="Runtime" Binding="{Binding Runtime}" Width="130"/>
                                        <DataGridTextColumn Header="Mode" Binding="{Binding Mode}" Width="85"/>
                                        <DataGridCheckBoxColumn Header="Interactive" Binding="{Binding Interactive}" Width="80"/>
                                        <DataGridCheckBoxColumn Header="Confirmation" Binding="{Binding RequiresConfirmation}" Width="90"/>
                                        <DataGridTextColumn Header="Description" Binding="{Binding Description}" Width="*"/>
                                    </DataGrid.Columns>
                                </DataGrid>
                                <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,8,0,6">
                                    <Button x:Name="RunManualLaunchButton" Content="Run on this computer" IsEnabled="False" ToolTip="Runs the entry in a new console on this computer."/>
                                    <TextBlock Text="Manual entries never run through the orchestrators." VerticalAlignment="Center" Foreground="{StaticResource MutedBrush}" Margin="8,0,0,0"/>
                                </StackPanel>
                                <TextBox x:Name="ManualLaunchCommandBox" Grid.Row="2" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" FontFamily="Consolas"/>
                            </Grid>
                        </TabItem>
                    </TabControl>
                </Grid>
            </TabItem>

            <TabItem Header="Planning">
                <Grid Margin="12">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="360"/>
                    </Grid.ColumnDefinitions>
                    <DataGrid x:Name="PlanningGrid" SelectionMode="Single">
                        <DataGrid.RowStyle>
                            <Style TargetType="DataGridRow">
                                <Style.Triggers>
                                    <DataTrigger Binding="{Binding SwiBlocks}" Value="True"><Setter Property="Background" Value="#FDE7E9"/></DataTrigger>
                                </Style.Triggers>
                            </Style>
                        </DataGrid.RowStyle>
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Job" Binding="{Binding Name}" SortMemberPath="Name" SortDirection="Ascending" Width="250"/>
                            <DataGridTextColumn Header="Version" Binding="{Binding Version}" Width="70"/>
                            <DataGridTextColumn Header="Arguments" Binding="{Binding Arguments}" Width="180"/>
                            <DataGridCheckBoxColumn Header="Enabled" Binding="{Binding Enabled}" Width="70"/>
                            <DataGridTextColumn Header="SWI validity" Binding="{Binding Swi}" Width="150">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock">
                                        <Setter Property="ToolTip" Value="{Binding SwiDetail}"/>
                                    <Style.Triggers>
                                        <DataTrigger Binding="{Binding SwiBlocks}" Value="True"><Setter Property="Foreground" Value="#A4262C"/><Setter Property="FontWeight" Value="SemiBold"/></DataTrigger>
                                    </Style.Triggers>
                                    </Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn Header="SWI days left" Binding="{Binding SwiDaysLeft}" Width="85"/>
                            <DataGridTextColumn Header="Frequency" Binding="{Binding Frequency}" Width="90"/>
                            <DataGridTextColumn Header="Times" Binding="{Binding Times}" Width="155"/>
                            <DataGridTextColumn Header="Days" Binding="{Binding Days}" Width="150"/>
                            <DataGridTextColumn Header="Assignment" Binding="{Binding Assignment}" Width="90"/>
                            <DataGridTextColumn Header="Server / owner" Binding="{Binding Server}" Width="155"/>
                            <DataGridTextColumn Header="Next run" Binding="{Binding NextRun}" Width="145"/>
                            <DataGridTextColumn Header="Group" Binding="{Binding Group}" Width="100"/>
                            <DataGridTextColumn Header="Health" Binding="{Binding Health}" Width="115"/>
                            <DataGridTextColumn Header="Last status" Binding="{Binding LastStatus}" Width="130"/>
                            <DataGridTextColumn Header="Last success" Binding="{Binding LastSuccess}" Width="125"/>
                            <DataGridTextColumn Header="Age (h)" Binding="{Binding SuccessAgeHours}" Width="65"/>
                            <DataGridTextColumn Header="Max (h)" Binding="{Binding ExpectedMaxAgeHours}" Width="65"/>
                            <DataGridTextColumn Header="Avg (min)" Binding="{Binding AverageDurationMinutes}" Width="70"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    <Grid Grid.Column="1" Margin="14,0,0,0">
                        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                        <StackPanel Margin="0,0,0,8">
                            <Button x:Name="ApplyJobButton" Content="Apply to draft" Background="#E5F1FB" Margin="0,0,0,8"/>
                            <GroupBox Header="Run through the orchestrator">
                                <StackPanel>
                                    <CheckBox x:Name="IncludeDependenciesCheck" Content="Also run the enabled dependencies" Margin="0,0,0,8"/>
                                    <Button x:Name="RequestRunButton" Content="Request run" Background="#FFF4CE"/>
                                    <TextBlock Text="The request uses the published configuration. The orchestrators run the job on its elected server with their locks and dependency rules; nothing is started by this GUI." Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap" Margin="0,8,0,0"/>
                                </StackPanel>
                            </GroupBox>
                        </StackPanel>
                        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                            <StackPanel>
                                <GroupBox Header="Selected job">
                                    <StackPanel>
                                        <TextBlock x:Name="SelectedJobText" FontWeight="SemiBold" FontSize="15" Text="Select a job"/>
                                        <TextBlock Text="Arguments (read-only)" Margin="0,8,0,0"/>
                                        <TextBox x:Name="JobArgumentsBox" IsReadOnly="True" TextWrapping="Wrap" MinHeight="44" VerticalScrollBarVisibility="Auto"/>
                                        <CheckBox x:Name="JobEnabledCheck" Content="Enabled" Margin="0,10,0,8"/>
                                        <TextBlock Text="Frequency"/>
                                        <ComboBox x:Name="ScheduleTypeCombo"><ComboBoxItem Content="Daily"/><ComboBoxItem Content="Weekly"/></ComboBox>
                                        <TextBlock Text="Times (HH:mm, comma separated)"/>
                                        <TextBox x:Name="TimesBox"/>
                                        <TextBlock Text="Days"/>
                                        <UniformGrid x:Name="DaysPanel" Columns="2" Margin="0,3,0,8">
                                            <CheckBox x:Name="MondayCheck" Content="Monday" Margin="0,3"/>
                                            <CheckBox x:Name="TuesdayCheck" Content="Tuesday" Margin="0,3"/>
                                            <CheckBox x:Name="WednesdayCheck" Content="Wednesday" Margin="0,3"/>
                                            <CheckBox x:Name="ThursdayCheck" Content="Thursday" Margin="0,3"/>
                                            <CheckBox x:Name="FridayCheck" Content="Friday" Margin="0,3"/>
                                            <CheckBox x:Name="SaturdayCheck" Content="Saturday" Margin="0,3"/>
                                            <CheckBox x:Name="SundayCheck" Content="Sunday" Margin="0,3"/>
                                        </UniformGrid>
                                        <TextBlock Text="Missed run policy"/>
                                        <ComboBox x:Name="MissedPolicyCombo"><ComboBoxItem Content="RunOnce"/><ComboBoxItem Content="Skip"/></ComboBox>
                                        <TextBlock Text="Assignment"/>
                                        <ComboBox x:Name="AssignmentCombo"><ComboBoxItem Content="Elected"/><ComboBoxItem Content="Pinned"/><ComboBoxItem Content="Manual"/><ComboBoxItem Content="Legacy"/></ComboBox>
                                        <TextBlock Text="Pinned server"/>
                                        <ComboBox x:Name="PinnedServerCombo" IsEditable="False" IsEnabled="False"/>
                                        <UniformGrid Columns="2">
                                            <StackPanel Margin="0,0,6,0"><TextBlock Text="Timeout (min)"/><TextBox x:Name="TimeoutBox"/></StackPanel>
                                            <StackPanel Margin="6,0,0,0"><TextBlock Text="Retries"/><TextBox x:Name="RetriesBox"/></StackPanel>
                                        </UniformGrid>
                                        <UniformGrid Columns="2">
                                            <StackPanel Margin="0,0,6,0"><TextBlock Text="Retry delay (sec)"/><TextBox x:Name="RetryDelayBox"/></StackPanel>
                                            <StackPanel Margin="6,0,0,0"><TextBlock Text="Estimated (min)"/><TextBox x:Name="DurationBox"/></StackPanel>
                                        </UniformGrid>
                                        <TextBlock Text="Elected owners are read-only and come from the shared election plan." Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap" Margin="0,4,0,0"/>
                                    </StackPanel>
                                </GroupBox>
                                <GroupBox Header="Dependencies">
                                    <StackPanel>
                                        <TextBlock Text="Depends on (job names, comma separated)"/>
                                        <TextBox x:Name="DependsOnBox" TextWrapping="Wrap" AcceptsReturn="False" MinHeight="44"/>
                                        <UniformGrid Columns="2">
                                            <StackPanel Margin="0,0,6,0"><TextBlock Text="Rule"/><ComboBox x:Name="DependencyModeCombo"><ComboBoxItem Content="LatestOccurrence"/><ComboBoxItem Content="FreshSuccess"/></ComboBox></StackPanel>
                                            <StackPanel Margin="6,0,0,0"><TextBlock Text="Max age (h, 0 = auto)"/><TextBox x:Name="DependencyMaxAgeBox"/></StackPanel>
                                        </UniformGrid>
                                        <TextBlock x:Name="DependentsText" TextWrapping="Wrap" Foreground="{StaticResource MutedBrush}" Margin="0,0,0,6"/>
                                        <TextBlock Text="Why this job waits (published data, current rule)" FontWeight="SemiBold"/>
                                        <DataGrid x:Name="ReadinessGrid" Height="200" Margin="0,4,0,0">
                                            <DataGrid.Columns>
                                                <DataGridTextColumn Header="Dependency" Binding="{Binding Dependency}" Width="*"/>
                                                <DataGridTextColumn Header="State" Binding="{Binding State}" Width="60"/>
                                                <DataGridTextColumn Header="Age (h)" Binding="{Binding AgeHours}" Width="55"/>
                                                <DataGridTextColumn Header="Max" Binding="{Binding MaxAgeHours}" Width="45"/>
                                                <DataGridTextColumn Header="Detail" Binding="{Binding Detail}" Width="*"/>
                                            </DataGrid.Columns>
                                        </DataGrid>
                                    </StackPanel>
                                </GroupBox>
                            </StackPanel>
                        </ScrollViewer>
                    </Grid>
                </Grid>
            </TabItem>

            <TabItem Header="History">
                <Grid Margin="12">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="150"/><ColumnDefinition Width="150"/><ColumnDefinition Width="180"/><ColumnDefinition Width="250"/><ColumnDefinition Width="150"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0" Margin="0,0,8,0"><TextBlock Text="From"/><DatePicker x:Name="HistoryFromPicker"/></StackPanel>
                        <StackPanel Grid.Column="1" Margin="0,0,8,0"><TextBlock Text="To"/><DatePicker x:Name="HistoryToPicker"/></StackPanel>
                        <StackPanel Grid.Column="2" Margin="0,0,8,0"><TextBlock Text="Server"/><ComboBox x:Name="HistoryServerCombo"/></StackPanel>
                        <StackPanel Grid.Column="3" Margin="0,0,8,0"><TextBlock Text="Job"/><ComboBox x:Name="HistoryJobCombo" IsEditable="True"/></StackPanel>
                        <StackPanel Grid.Column="4" Margin="0,0,8,0"><TextBlock Text="Status"/><ComboBox x:Name="HistoryStatusCombo"/></StackPanel>
                        <Button x:Name="HistoryRefreshButton" Grid.Column="5" Content="Filter" VerticalAlignment="Bottom" Margin="4,0,4,8"/>
                        <Button x:Name="ExportCsvButton" Grid.Column="6" Content="Export CSV" VerticalAlignment="Bottom" Margin="4,0,4,8"/>
                        <Button x:Name="ExportHtmlButton" Grid.Column="7" Content="Export HTML" VerticalAlignment="Bottom" Margin="4,0,4,8"/>
                        <Button x:Name="Failures24hButton" Grid.Column="8" Content="Failures 24 h" VerticalAlignment="Bottom" Margin="4,0,4,8" Background="#FDE7E9"/>
                    </Grid>
                    <DataGrid x:Name="HistoryGrid" Grid.Row="1" Margin="0,10,0,0">
                        <DataGrid.Columns>
                            <!-- One time column (GUI 1.6.31): the start, or the scheduled time for a run that never started. -->
                            <DataGridTextColumn Header="Time" Binding="{Binding EventTime}" Width="145">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock"><Setter Property="ToolTip" Value="Start time of the run; for a skipped or blocked occurrence (no start), its scheduled time."/></Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn Header="Scheduled" Binding="{Binding Scheduled}" Width="115">
                                <DataGridTextColumn.ElementStyle>
                                    <Style TargetType="TextBlock"><Setter Property="ToolTip" Value="Scheduled occurrence the run belongs to (differs from Time for a catch-up, retry, forced or pipeline run)."/></Style>
                                </DataGridTextColumn.ElementStyle>
                            </DataGridTextColumn>
                            <DataGridTextColumn Header="Server" Binding="{Binding Server}" Width="150"/>
                            <DataGridTextColumn Header="Job" Binding="{Binding JobName}" Width="250"/>
                            <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="100"/>
                            <DataGridTextColumn Header="Duration (min)" Binding="{Binding DurationMinutes}" Width="105"/>
                            <DataGridTextColumn Header="Exit" Binding="{Binding ExitCode}" Width="60"/>
                            <DataGridTextColumn Header="Retry" Binding="{Binding RetryCount}" Width="60"/>
                            <DataGridTextColumn Header="Log" Binding="{Binding LogPath}" Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </TabItem>

            <TabItem Header="Events">
                <Grid Margin="12">
                    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="90"/><RowDefinition Height="220"/></Grid.RowDefinitions>
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="180"/><ColumnDefinition Width="160"/><ColumnDefinition Width="160"/><ColumnDefinition Width="110"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0" Margin="0,0,8,0"><TextBlock Text="Server"/><ComboBox x:Name="EventsServerCombo"/></StackPanel>
                        <StackPanel Grid.Column="1" Margin="0,0,8,0"><TextBlock Text="Level"/><ComboBox x:Name="EventsLevelCombo"/></StackPanel>
                        <StackPanel Grid.Column="2" Margin="0,0,8,0"><TextBlock Text="Category"/><ComboBox x:Name="EventsCategoryCombo"/></StackPanel>
                        <StackPanel Grid.Column="3" Margin="0,0,8,0"><TextBlock Text="Period"/><ComboBox x:Name="EventsPeriodCombo"/></StackPanel>
                        <StackPanel Grid.Column="4" Margin="0,0,8,0"><TextBlock Text="Text"/><TextBox x:Name="EventsTextBox" ToolTip="Case-insensitive search in the message and category."/></StackPanel>
                        <CheckBox x:Name="EventsGroupCheck" Grid.Column="5" Content="Group repeats" IsChecked="True" VerticalAlignment="Bottom" Margin="4,0,8,14" ToolTip="Merge identical messages of one server; Count and First show the repetitions."/>
                        <Button x:Name="EventsRefreshButton" Grid.Column="6" Content="Read logs" VerticalAlignment="Bottom" Margin="4,0,4,8"/>
                    </Grid>
                    <TextBlock x:Name="EventsSummaryText" Grid.Row="1" Margin="0,0,0,6" Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap" Text="Resident orchestrator logs are read while this tab is open (read-only)."/>
                    <DataGrid x:Name="EventsGrid" Grid.Row="2" SelectionMode="Single">
                        <DataGrid.RowStyle>
                            <Style TargetType="DataGridRow">
                                <Style.Triggers>
                                    <DataTrigger Binding="{Binding Level}" Value="ERROR"><Setter Property="Background" Value="#FDE7E9"/></DataTrigger>
                                    <DataTrigger Binding="{Binding Level}" Value="WARN"><Setter Property="Background" Value="#FFF4CE"/></DataTrigger>
                                </Style.Triggers>
                            </Style>
                        </DataGrid.RowStyle>
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Time" Binding="{Binding Time}" Width="145"/>
                            <DataGridTextColumn Header="Server" Binding="{Binding Server}" Width="150"/>
                            <DataGridTextColumn Header="Level" Binding="{Binding Level}" Width="60"/>
                            <DataGridTextColumn Header="Category" Binding="{Binding Category}" Width="120"/>
                            <DataGridTextColumn Header="Count" Binding="{Binding Count}" Width="55"/>
                            <DataGridTextColumn Header="First" Binding="{Binding First}" Width="145"/>
                            <DataGridTextColumn Header="Message" Binding="{Binding Message}" Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    <TextBox x:Name="EventsDetailBox" Grid.Row="3" Margin="0,8,0,0" FontFamily="Consolas" FontSize="12" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
                    <!-- Moved from Operations > Live in GUI 1.6.17 so the running and pending job grids get the height. -->
                    <Grid Grid.Row="4" Margin="0,10,0,0">
                        <Grid.ColumnDefinitions><ColumnDefinition Width="9*"/><ColumnDefinition Width="11*"/></Grid.ColumnDefinitions>
                        <GroupBox Header="Active peer-monitoring incidents" Margin="0,0,6,0">
                            <DataGrid x:Name="OperationsIncidentsGrid">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Observed by" Binding="{Binding ObservedBy}" Width="130"/>
                                    <DataGridTextColumn Header="Issue" Binding="{Binding Issue}" Width="*"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </GroupBox>
                        <GroupBox Grid.Column="1" Header="Orchestrator mails (24 h) - double-click to open" Margin="6,0,0,0">
                            <DataGrid x:Name="OperationsMailsGrid">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Time" Binding="{Binding Time}" Width="120"/>
                                    <DataGridTextColumn Header="Server" Binding="{Binding Server}" Width="120"/>
                                    <DataGridTextColumn Header="Subject" Binding="{Binding Subject}" Width="*"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </GroupBox>
                    </Grid>
                </Grid>
            </TabItem>

            <TabItem Header="Configuration">
                <Grid Margin="12">
                    <TabControl x:Name="ConfigurationTabs">
                        <!-- Servers and versions first and selected by default (GUI 1.6.28). -->
                        <TabItem x:Name="ConfigurationServersTab" Header="Servers and versions">
                            <Grid Margin="12">
                                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="390"/></Grid.ColumnDefinitions>
                                <!-- Servers grid on the full width, chart and server panel below (GUI 1.6.29). -->
                                <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                                    <DataGrid x:Name="ServersGrid" Grid.ColumnSpan="2" MinHeight="210" MaxHeight="340" ToolTip="OS, CPU, RAM and drive C: are published by each orchestrator when it starts.">
                                        <DataGrid.RowStyle>
                                            <Style TargetType="DataGridRow">
                                                <Style.Triggers>
                                                    <DataTrigger Binding="{Binding State}" Value="Inactive"><Setter Property="Foreground" Value="#6B6B6B"/></DataTrigger>
                                                    <DataTrigger Binding="{Binding SystemDriveLow}" Value="True"><Setter Property="Background" Value="#FDE7E9"/></DataTrigger>
                                                </Style.Triggers>
                                            </Style>
                                        </DataGrid.RowStyle>
                                        <DataGrid.Columns>
                                            <DataGridTextColumn Header="Server" Binding="{Binding Server}" Width="150"/>
                                            <DataGridTextColumn Header="State" Binding="{Binding State}" Width="70">
                                                <DataGridTextColumn.ElementStyle>
                                                    <Style TargetType="TextBlock"><Setter Property="ToolTip" Value="Active = takes jobs. Inactive = runs as an orchestrator (heartbeat, monitoring, maintenance) but takes no job. A new server is Inactive."/></Style>
                                                </DataGridTextColumn.ElementStyle>
                                            </DataGridTextColumn>
                                            <DataGridCheckBoxColumn Header="Online" Binding="{Binding Online}" Width="60"/>
                                            <DataGridTextColumn Header="Orchestrator" Binding="{Binding Version}" Width="90"/>
                                            <DataGridTextColumn Header="Weight" Binding="{Binding Weight}" Width="60"/>
                                            <DataGridTextColumn Header="Policy" Binding="{Binding Policy}" Width="160"/>
                                            <DataGridTextColumn Header="Capabilities" Binding="{Binding Capabilities}" Width="220"/>
                                            <DataGridTextColumn Header="OS" Binding="{Binding Os}" Width="230"/>
                                            <DataGridTextColumn Header="CPU" Binding="{Binding Cpu}" Width="220"/>
                                            <DataGridTextColumn Header="CPU %" Binding="{Binding CpuPercent}" Width="60"/>
                                            <DataGridTextColumn Header="RAM (GB)" Binding="{Binding MemoryGB}" Width="75"/>
                                            <DataGridTextColumn Header="RAM %" Binding="{Binding MemoryPercent}" Width="60"/>
                                            <DataGridTextColumn Header="C: size (GB)" Binding="{Binding SystemDriveSizeGB}" Width="90"/>
                                            <DataGridTextColumn Header="C: free %" Binding="{Binding SystemDriveFreePercent}" Width="75"/>
                                            <DataGridTextColumn Header="Usage at" Binding="{Binding UsageAt}" Width="120"/>
                                            <DataGridTextColumn Header="Started at (static)" Binding="{Binding SystemInfoAt}" Width="130"/>
                                            <DataGridTextColumn Header="Jobs" Binding="{Binding AssignedJobs}" Width="50"/>
                                        </DataGrid.Columns>
                                    </DataGrid>
                                <Grid Grid.Row="1" Margin="0,10,0,0">
                                    <Grid.RowDefinitions><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                                    <GroupBox Header="CPU and memory, last 12 hours (blue = CPU, green = RAM; sampled every 5 minutes by each orchestrator)" Margin="0,10,0,0">
                                        <!-- No fixed height: the chart fills the free height (GUI 1.6.30). -->
                                        <Canvas x:Name="ServersMetricsCanvas" MinHeight="120" ClipToBounds="True" Background="White"/>
                                    </GroupBox>
                                    <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,10,0,0">
                                        <TextBox x:Name="NewServerBox" Width="220" Margin="0,0,8,0"/>
                                        <Button x:Name="AddServerButton" Content="Add server"/>
                                        <Button x:Name="RemoveServerButton" Content="Remove selected"/>
                                    </StackPanel>
                                </Grid>
                                <!-- Right panel scrolls when the window is too low for the versions list (GUI 1.6.29). -->
                                <ScrollViewer Grid.Row="1" Grid.Column="1" Margin="14,10,0,0" VerticalScrollBarVisibility="Auto">
                                <Grid>
                                    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*" MinHeight="230"/></Grid.RowDefinitions>
                                        <GroupBox Header="Selected server">
                                            <StackPanel>
                                                <TextBlock x:Name="SelectedServerText" FontWeight="SemiBold" Text="Select a server"/>
                                                <CheckBox x:Name="ServerActiveCheck" Content="Active (takes jobs)" Margin="0,8,0,0" ToolTip="Inactive: the orchestrator keeps running (heartbeat, peer monitoring, maintenance, disk space guard, mails) but launches no job and is never elected. A new server is Inactive until you activate it here and publish."/>
                                                <TextBlock Text="Election weight" Margin="0,8,0,0"/>
                                                <TextBox x:Name="ServerWeightBox"/>
                                                <TextBlock Text="Only jobs requiring capabilities, comma separated (blank = unrestricted)"/>
                                                <ComboBox x:Name="ServerPolicyCombo" IsEditable="True">
                                                    <ComboBoxItem Content=""/>
                                                    <ComboBoxItem Content="ExchangeOnPrem"/>
                                                    <ComboBoxItem Content="SharePointOnPrem"/>
                                                    <ComboBoxItem Content="AD"/>
                                                    <ComboBoxItem Content="Graph"/>
                                                    <ComboBoxItem Content="EXO"/>
                                                    <ComboBoxItem Content="TeamsPowerShell"/>
                                                    <ComboBoxItem Content="Adaxes"/>
                                                    <ComboBoxItem Content="Python"/>
                                                </ComboBox>
                                                <Button x:Name="ApplyServerButton" Content="Apply server settings to draft" Background="#E5F1FB" Margin="0,6,0,0"/>
                                                <!-- Acceptance test before activation (GUI 1.6.32): it reads from the computer it runs on. -->
                                                <TextBlock Text="Acceptance test (limited real reads, nothing kept)" Margin="0,12,0,0"/>
                                                <TextBlock x:Name="ServerTestHintText" Foreground="#5F6B7A" TextWrapping="Wrap" Text="Select a server"/>
                                                <Button x:Name="TestServerButton" Content="Test this server" IsEnabled="False" Margin="0,6,0,0" ToolTip="Runs the Orchestrator-ServerAcceptance manual launch in a new console on this computer."/>
                                                <Button x:Name="TestServerAsServiceAccountButton" Content="Test as service account" IsEnabled="False" Margin="0,6,0,0" ToolTip="Same test started with runas under the account of the orchestrator scheduled task of this computer."/>
                                            </StackPanel>
                                        </GroupBox>
                                        <GroupBox Grid.Row="1" Header="Configuration versions" Margin="0,10,0,0">
                                            <Grid>
                                                <Grid.RowDefinitions><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                                                <DataGrid x:Name="VersionsGrid">
                                                    <DataGrid.Columns>
                                                        <DataGridTextColumn Header="Version" Binding="{Binding VersionId}" Width="*"/>
                                                        <DataGridTextColumn Header="Created" Binding="{Binding Created}" Width="135"/>
                                                    </DataGrid.Columns>
                                                </DataGrid>
                                                <Button x:Name="RollbackButton" Grid.Row="1" Content="Rollback to state before selected version" Margin="0,8,0,0"/>
                                            </Grid>
                                        </GroupBox>
                                </Grid>
                                </ScrollViewer>
                            </Grid>
                        </TabItem>
                        <TabItem x:Name="ConfigurationSettingsTab" Header="Settings">
                            <Grid Margin="8">
                                <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                                <Grid>
                                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                                    <StackPanel Margin="0,0,8,0"><TextBlock Text="Filter (parameter, description, source, visible value)"/><TextBox x:Name="ConfigFilterBox"/></StackPanel>
                                    <Button x:Name="ConfigShowSensitiveButton" Grid.Column="1" Content="Show" VerticalAlignment="Bottom" Margin="4,0,4,8" ToolTip="Shows hidden sensitive values for 30 seconds."/>
                                    <Button x:Name="ConfigRefreshButton" Grid.Column="2" Content="Refresh" VerticalAlignment="Bottom" Margin="4,0,4,8"/>
                                    <Button x:Name="ConfigOpenFolderButton" Grid.Column="3" Content="Open folder" VerticalAlignment="Bottom" Margin="4,0,4,8" ToolTip="Opens the folder of the selected file in Explorer."/>
                                </Grid>
                                <TextBlock x:Name="ConfigSummaryText" Grid.Row="1" Margin="0,0,0,6" Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap" Text="Read-only view of the configuration of the current tenant profile. Nothing is written."/>
                                <TabControl x:Name="ConfigFileTabs" Grid.Row="2">
                                    <TabItem x:Name="ConfigGlobalTab" Header="Global">
                                        <Grid Margin="8">
                                            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                                            <TextBlock x:Name="ConfigGlobalStatusText" Margin="0,0,0,6" Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap"/>
                                            <DataGrid x:Name="ConfigGlobalGrid" Grid.Row="1" SelectionMode="Single" ClipboardCopyMode="None">
                                                <DataGrid.RowStyle>
                                                    <Style TargetType="DataGridRow">
                                                        <Style.Triggers>
                                                            <DataTrigger Binding="{Binding Model}" Value="Not in template"><Setter Property="Background" Value="#FFF4CE"/></DataTrigger>
                                                            <DataTrigger Binding="{Binding Model}" Value="Missing from file"><Setter Property="Background" Value="#FDE7E9"/></DataTrigger>
                                                            <DataTrigger Binding="{Binding Masked}" Value="True"><Setter Property="FontStyle" Value="Italic"/></DataTrigger>
                                                        </Style.Triggers>
                                                    </Style>
                                                </DataGrid.RowStyle>
                                                <DataGrid.Columns>
                                                    <DataGridTextColumn Header="Parameter" Binding="{Binding Parameter}" Width="300"/>
                                                    <DataGridTextColumn Header="Value" Binding="{Binding Value}" Width="220"/>
                                                    <DataGridTextColumn Header="Effective value" Binding="{Binding Effective}" Width="220"/>
                                                    <DataGridTextColumn Header="Source" Binding="{Binding Source}" Width="150"/>
                                                    <DataGridTextColumn Header="Template" Binding="{Binding Model}" Width="115"/>
                                                    <DataGridTextColumn Header="Description" Binding="{Binding Description}" Width="*"/>
                                                </DataGrid.Columns>
                                            </DataGrid>
                                        </Grid>
                                    </TabItem>
                                    <TabItem x:Name="ConfigTenantTab" Header="Tenant">
                                        <Grid Margin="8">
                                            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                                            <TextBlock x:Name="ConfigTenantStatusText" Margin="0,0,0,6" Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap"/>
                                            <DataGrid x:Name="ConfigTenantGrid" Grid.Row="1" SelectionMode="Single" ClipboardCopyMode="None">
                                                <DataGrid.RowStyle>
                                                    <Style TargetType="DataGridRow">
                                                        <Style.Triggers>
                                                            <DataTrigger Binding="{Binding Model}" Value="Not in template"><Setter Property="Background" Value="#FFF4CE"/></DataTrigger>
                                                            <DataTrigger Binding="{Binding Model}" Value="Missing from file"><Setter Property="Background" Value="#FDE7E9"/></DataTrigger>
                                                            <DataTrigger Binding="{Binding Masked}" Value="True"><Setter Property="FontStyle" Value="Italic"/></DataTrigger>
                                                        </Style.Triggers>
                                                    </Style>
                                                </DataGrid.RowStyle>
                                                <DataGrid.Columns>
                                                    <DataGridTextColumn Header="Parameter" Binding="{Binding Parameter}" Width="300"/>
                                                    <DataGridTextColumn Header="Value" Binding="{Binding Value}" Width="220"/>
                                                    <DataGridTextColumn Header="Effective value" Binding="{Binding Effective}" Width="220"/>
                                                    <DataGridTextColumn Header="Source" Binding="{Binding Source}" Width="150"/>
                                                    <DataGridTextColumn Header="Template" Binding="{Binding Model}" Width="115"/>
                                                    <DataGridTextColumn Header="Description" Binding="{Binding Description}" Width="*"/>
                                                </DataGrid.Columns>
                                            </DataGrid>
                                        </Grid>
                                    </TabItem>
                                    <TabItem x:Name="ConfigMailTab" Header="Mail">
                                        <Grid Margin="8">
                                            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                                            <TextBlock x:Name="ConfigMailStatusText" Margin="0,0,0,6" Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap"/>
                                            <DataGrid x:Name="ConfigMailGrid" Grid.Row="1" SelectionMode="Single" ClipboardCopyMode="None">
                                                <DataGrid.RowStyle>
                                                    <Style TargetType="DataGridRow">
                                                        <Style.Triggers>
                                                            <DataTrigger Binding="{Binding Model}" Value="Not in template"><Setter Property="Background" Value="#FFF4CE"/></DataTrigger>
                                                            <DataTrigger Binding="{Binding Model}" Value="Missing from file"><Setter Property="Background" Value="#FDE7E9"/></DataTrigger>
                                                            <DataTrigger Binding="{Binding Masked}" Value="True"><Setter Property="FontStyle" Value="Italic"/></DataTrigger>
                                                        </Style.Triggers>
                                                    </Style>
                                                </DataGrid.RowStyle>
                                                <DataGrid.Columns>
                                                    <DataGridTextColumn Header="Parameter" Binding="{Binding Parameter}" Width="300"/>
                                                    <DataGridTextColumn Header="Value" Binding="{Binding Value}" Width="220"/>
                                                    <DataGridTextColumn Header="Effective value" Binding="{Binding Effective}" Width="220"/>
                                                    <DataGridTextColumn Header="Source" Binding="{Binding Source}" Width="150"/>
                                                    <DataGridTextColumn Header="Template" Binding="{Binding Model}" Width="115"/>
                                                    <DataGridTextColumn Header="Description" Binding="{Binding Description}" Width="*"/>
                                                </DataGrid.Columns>
                                            </DataGrid>
                                        </Grid>
                                    </TabItem>
                                </TabControl>
                            </Grid>
                        </TabItem>
                    </TabControl>
                </Grid>
            </TabItem>
        </TabControl>

        <Border Grid.Row="3" Background="White" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="7" Margin="0,9,0,0" Padding="10,6">
            <Grid><TextBlock x:Name="FooterText" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center"/><TextBlock x:Name="VersionText" Text="v1.1.0" HorizontalAlignment="Right" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center"/></Grid>
        </Border>
    </Grid>
</Window>
'@

function ConvertFrom-OrchestratorGuiXaml {
    param([Parameter(Mandatory = $true)][string]$Text)
    $reader = [System.Xml.XmlNodeReader]::new([xml]$Text)
    return [System.Windows.Markup.XamlReader]::Load($reader)
}

function Show-OrchestratorDialog {
    # Branded replacement of System.Windows.MessageBox: same buttons, icons and results (OK, Yes, No, Cancel).
    param(
        [AllowNull()][AllowEmptyString()][string]$Message,
        [string]$Title = 'SmartM365 Orchestrator',
        [ValidateSet('OK', 'OKCancel', 'YesNo', 'YesNoCancel')][string]$Buttons = 'OK',
        [string]$Icon = 'None'
    )
    $style = switch -Regex ($Icon) {
        '^(Information|Asterisk)$' { @{ Glyph = 0xE946; Color = '#0078D4'; Tint = '#E5F1FB' }; break }
        '^(Warning|Exclamation)$' { @{ Glyph = 0xE7BA; Color = '#92400E'; Tint = '#FFF4CE' }; break }
        '^(Error|Hand|Stop)$' { @{ Glyph = 0xEA39; Color = '#A4262C'; Tint = '#FDE7E9' }; break }
        '^Question$' { @{ Glyph = 0xE9CE; Color = '#0078D4'; Tint = '#E5F1FB' }; break }
        default { @{ Glyph = 0xE946; Color = '#5D6B78'; Tint = '#F3F6FA' } }
    }
    $dialogXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent" ResizeMode="NoResize"
        SizeToContent="WidthAndHeight" ShowInTaskbar="False" FontFamily="Segoe UI" UseLayoutRounding="True">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Padding" Value="16,7"/>
            <Setter Property="Margin" Value="8,0,0,0"/>
            <Setter Property="MinHeight" Value="32"/>
            <Setter Property="MinWidth" Value="92"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="Foreground" Value="#17324D"/>
            <Setter Property="BorderBrush" Value="#C3CFDC"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Grid>
                            <Border x:Name="Chrome" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6"/>
                            <Border x:Name="Overlay" Background="#17324D" Opacity="0" CornerRadius="6"/>
                            <ContentPresenter Margin="{TemplateBinding Padding}" HorizontalAlignment="Center" VerticalAlignment="Center" RecognizesAccessKey="True"/>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Overlay" Property="Opacity" Value="0.07"/></Trigger>
                            <Trigger Property="IsPressed" Value="True"><Setter TargetName="Overlay" Property="Opacity" Value="0.15"/></Trigger>
                            <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Chrome" Property="BorderBrush" Value="#0078D4"/></Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Border Margin="18" Background="White" BorderBrush="#D5DEE8" BorderThickness="1" CornerRadius="10">
        <Border.Effect><DropShadowEffect BlurRadius="18" ShadowDepth="2" Opacity="0.28"/></Border.Effect>
        <Grid MinWidth="420" MaxWidth="640">
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
                <RowDefinition Height="Auto"/>
            </Grid.RowDefinitions>
            <Border x:Name="DialogHeader" Background="#17324D" CornerRadius="9,9,0,0" Padding="16,10" Cursor="SizeAll">
                <DockPanel>
                    <Image x:Name="DialogLogo" Width="26" Height="26" Margin="0,0,10,0" DockPanel.Dock="Left" Stretch="Uniform"/>
                    <TextBlock x:Name="DialogTitle" Foreground="White" FontSize="15" FontWeight="SemiBold" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
                </DockPanel>
            </Border>
            <Grid Grid.Row="1" Margin="22,20,22,16">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Border x:Name="DialogIconBadge" Width="42" Height="42" CornerRadius="21" VerticalAlignment="Top" Margin="0,0,16,0">
                    <TextBlock x:Name="DialogIconGlyph" FontFamily="Segoe MDL2 Assets" FontSize="20" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                </Border>
                <ScrollViewer Grid.Column="1" MaxHeight="380" VerticalScrollBarVisibility="Auto" VerticalAlignment="Center">
                    <TextBox x:Name="DialogMessage" IsReadOnly="True" BorderThickness="0" Background="Transparent" TextWrapping="Wrap" FontSize="13" Foreground="#17324D" Padding="0"/>
                </ScrollViewer>
            </Grid>
            <Border Grid.Row="2" Background="#F7FAFD" BorderBrush="#E3E9F0" BorderThickness="0,1,0,0" CornerRadius="0,0,9,9" Padding="16,10">
                <StackPanel x:Name="DialogButtons" Orientation="Horizontal" HorizontalAlignment="Right"/>
            </Border>
        </Grid>
    </Border>
</Window>
'@
    $dialog = ConvertFrom-OrchestratorGuiXaml -Text $dialogXaml
    $dialog.Title = $Title
    $dialog.FindName('DialogTitle').Text = $Title
    $dialog.FindName('DialogMessage').Text = [string]$Message
    $converter = [System.Windows.Media.BrushConverter]::new()
    $dialog.FindName('DialogIconBadge').Background = $converter.ConvertFromString($style.Tint)
    $glyph = $dialog.FindName('DialogIconGlyph')
    $glyph.Text = [string][char]$style.Glyph
    $glyph.Foreground = $converter.ConvertFromString($style.Color)
    $logoFile = Join-Path $PSScriptRoot 'WorkplaceCloudHub-lockup-WPF.png'
    if (Test-Path -LiteralPath $logoFile) {
        $logo = [System.Windows.Media.Imaging.BitmapImage]::new()
        $logo.BeginInit(); $logo.CacheOption = 'OnLoad'; $logo.UriSource = [Uri]$logoFile; $logo.EndInit(); $logo.Freeze()
        $dialog.FindName('DialogLogo').Source = $logo
    }
    $dialog.FindName('DialogHeader').Add_MouseLeftButtonDown({ param($source) [System.Windows.Window]::GetWindow($source).DragMove() })

    $labels = switch ($Buttons) { 'OK' { @('OK') } 'OKCancel' { @('OK', 'Cancel') } 'YesNo' { @('Yes', 'No') } 'YesNoCancel' { @('Yes', 'No', 'Cancel') } }
    # Escape and the window close answer like MessageBox: Cancel when offered, else No, else OK.
    $cancelResult = if ($labels -contains 'Cancel') { 'Cancel' } elseif ($labels -contains 'No') { 'No' } else { 'OK' }
    $accent = $converter.ConvertFromString('#0078D4')
    $panel = $dialog.FindName('DialogButtons')
    $primary = $null
    foreach ($label in $labels) {
        $button = [System.Windows.Controls.Button]::new()
        $button.Content = $label
        $button.Tag = $label
        if ($null -eq $primary) {
            $button.Background = $accent; $button.BorderBrush = $accent; $button.Foreground = [System.Windows.Media.Brushes]::White
            $button.FontWeight = [System.Windows.FontWeights]::SemiBold; $button.IsDefault = $true
            $primary = $button
        }
        if ($label -eq $cancelResult) { $button.IsCancel = $true }
        $button.Add_Click({ param($source) $owner = [System.Windows.Window]::GetWindow($source); $owner.Tag = [string]$source.Tag; $owner.Close() })
        [void]$panel.Children.Add($button)
    }
    $mainWindow = Get-Variable -Name window -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    if ($mainWindow -is [System.Windows.Window] -and $mainWindow.IsVisible) {
        $dialog.Owner = $mainWindow
        $dialog.WindowStartupLocation = 'CenterOwner'
        if ($mainWindow.Icon) { $dialog.Icon = $mainWindow.Icon }
    }
    else {
        $dialog.WindowStartupLocation = 'CenterScreen'
        $dialog.Topmost = $true
    }
    $dialog.Add_ContentRendered({ param($source) [void]$source.FindName('DialogButtons').Children[0].Focus() })
    [void]$dialog.ShowDialog()
    if ([string]::IsNullOrEmpty([string]$dialog.Tag)) { return $cancelResult }
    return [string]$dialog.Tag
}

# Same call shape as $script:OrchestratorDialog.Show(message, title, buttons, icon).
$script:OrchestratorDialog = [pscustomobject]@{}
Add-Member -InputObject $script:OrchestratorDialog -MemberType ScriptMethod -Name Show -Value {
    param($Message, $Title = 'SmartM365 Orchestrator', $Buttons = 'OK', $Icon = 'None')
    Show-OrchestratorDialog -Message ([string]$Message) -Title ([string]$Title) -Buttons ([string]$Buttons) -Icon ([string]$Icon)
}

function Get-OrchestratorGuiPropertyValue {
    param(
        [AllowNull()]$Object,
        [Parameter(Mandatory = $true)][string]$Name,
        [AllowNull()]$DefaultValue
    )
    if ($null -eq $Object -or -not $Object.PSObject.Properties[$Name] -or $null -eq $Object.$Name) {
        return $DefaultValue
    }
    return $Object.$Name
}
$script:DashboardLevelColors = @{ Normal = '#FFFFFF'; Info = '#E5F1FB'; Warning = '#FFF4CE'; Critical = '#FDE7E9' }
$script:DashboardBannerColors = @{ Normal = '#DFF6DD'; Info = '#E5F1FB'; Warning = '#FFF4CE'; Critical = '#FDE7E9' }
$script:DashboardTimelineColors = @{ Running = '#005A9E'; Planned = '#8DB6E0'; Suspended = '#B8B8B8'; Unassigned = '#F7B955' }

function ConvertTo-OrchestratorBrush {
    param([Parameter(Mandatory = $true)][string]$Color)
    return [System.Windows.Media.BrushConverter]::new().ConvertFromString($Color)
}

function Set-OrchestratorDashboardView {
    # Renders a dashboard object (Get-SmartM365OrchestratorDashboard shape) into the Dashboard tab.
    param([Parameter(Mandatory = $true)]$Controls, [AllowNull()]$Dashboard)
    if ($null -eq $Dashboard) { return }
    $level = [string]$Dashboard.Banner.Level
    $Controls.DashboardBannerBorder.Background = ConvertTo-OrchestratorBrush $(if ($script:DashboardBannerColors.ContainsKey($level)) { $script:DashboardBannerColors[$level] } else { '#F3F6FA' })
    $Controls.DashboardBannerText.Text = [string]$Dashboard.Banner.Text
    $Controls.DashboardTilesList.ItemsSource = @($Dashboard.Tiles | ForEach-Object {
        $tileLevel = [string]$_.Level
        $_ | Select-Object *, @{ Name = 'Background'; Expression = { if ($script:DashboardLevelColors.ContainsKey($tileLevel)) { $script:DashboardLevelColors[$tileLevel] } else { '#FFFFFF' } } }
    })
    $actions = @($Dashboard.Actions)
    $Controls.DashboardActionsGrid.ItemsSource = $actions
    $Controls.DashboardActionsGroup.Header = if ($actions.Count) { 'To review ({0})' -f $actions.Count } else { 'To review (nothing to review)' }
    $Controls.DashboardGrid.ItemsSource = @($Dashboard.Servers)
    $Controls.DashboardTimelineCanvas.Tag = $Dashboard.Timeline
    Update-OrchestratorDashboardTimeline -Canvas $Controls.DashboardTimelineCanvas
    $Controls.DashboardMetricsCanvas.Tag = if ($Dashboard.PSObject.Properties['Metrics']) { @($Dashboard.Metrics) } else { @() }
    Update-OrchestratorServerMetricsChart -Canvas $Controls.DashboardMetricsCanvas
}

function Update-OrchestratorDashboardTimeline {
    # Draws one lane per server for the timeline stored in Canvas.Tag; redrawn on resize.
    param([Parameter(Mandatory = $true)]$Canvas)
    $Canvas.Children.Clear()
    $timeline = $Canvas.Tag
    if ($null -eq $timeline) { return }
    $lanes = @($timeline.Lanes)
    $labelWidth = 150.0; $laneHeight = 24.0; $top = 22.0
    $Canvas.Height = $top + [math]::Max(1, $lanes.Count) * $laneHeight + 4
    $width = [double]$Canvas.ActualWidth
    if ($width -lt ($labelWidth + 120)) { return }
    $plotWidth = $width - $labelWidth - 10
    $from = [datetime]$timeline.From; $to = [datetime]$timeline.To
    $totalMinutes = [math]::Max(1.0, ($to - $from).TotalMinutes)
    $gridBrush = ConvertTo-OrchestratorBrush '#E3E9F0'; $mutedBrush = ConvertTo-OrchestratorBrush '#5D6B78'
    for ($index = 0; $index -lt $lanes.Count; $index++) {
        if ($index % 2 -eq 1) {
            $band = [System.Windows.Shapes.Rectangle]::new(); $band.Width = $width; $band.Height = $laneHeight; $band.Fill = ConvertTo-OrchestratorBrush '#F7FAFD'
            [System.Windows.Controls.Canvas]::SetLeft($band, 0); [System.Windows.Controls.Canvas]::SetTop($band, $top + $index * $laneHeight); [void]$Canvas.Children.Add($band)
        }
        $label = [System.Windows.Controls.TextBlock]::new(); $label.Text = [string]$lanes[$index]; $label.Width = $labelWidth - 8; $label.TextTrimming = 'CharacterEllipsis'
        if ([string]$lanes[$index] -eq 'Unassigned') { $label.Foreground = ConvertTo-OrchestratorBrush '#B45309' }
        [System.Windows.Controls.Canvas]::SetLeft($label, 4); [System.Windows.Controls.Canvas]::SetTop($label, $top + $index * $laneHeight + 4); [void]$Canvas.Children.Add($label)
    }
    $tick = $from.Date.AddHours($from.Hour + 1)
    if ($tick.Hour % 2 -eq 1) { $tick = $tick.AddHours(1) }
    while ($tick -lt $to) {
        $x = $labelWidth + ($tick - $from).TotalMinutes / $totalMinutes * $plotWidth
        $line = [System.Windows.Shapes.Line]::new(); $line.X1 = $x; $line.X2 = $x; $line.Y1 = $top - 4; $line.Y2 = $Canvas.Height; $line.Stroke = $gridBrush
        [void]$Canvas.Children.Add($line)
        $hour = [System.Windows.Controls.TextBlock]::new(); $hour.Text = $tick.ToString('HH:mm'); $hour.FontSize = 11; $hour.Foreground = $mutedBrush
        [System.Windows.Controls.Canvas]::SetLeft($hour, $x - 15); [System.Windows.Controls.Canvas]::SetTop($hour, 2); [void]$Canvas.Children.Add($hour)
        $tick = $tick.AddHours(2)
    }
    $nowLine = [System.Windows.Shapes.Line]::new(); $nowLine.X1 = $labelWidth; $nowLine.X2 = $labelWidth; $nowLine.Y1 = $top - 4; $nowLine.Y2 = $Canvas.Height
    $nowLine.Stroke = ConvertTo-OrchestratorBrush '#C42B1C'; $nowLine.StrokeThickness = 2; [void]$Canvas.Children.Add($nowLine)
    foreach ($item in @($timeline.Items)) {
        $laneIndex = [array]::IndexOf([object[]]$lanes, [string]$item.Lane)
        if ($laneIndex -lt 0) { continue }
        $start = if ([datetime]$item.Start -lt $from) { $from } else { [datetime]$item.Start }
        $end = if ([datetime]$item.End -gt $to) { $to } else { [datetime]$item.End }
        if ($end -le $from -or $start -ge $to) { continue }
        $x1 = $labelWidth + ($start - $from).TotalMinutes / $totalMinutes * $plotWidth
        $x2 = $labelWidth + ($end - $from).TotalMinutes / $totalMinutes * $plotWidth
        $kind = [string]$item.Kind
        $bar = [System.Windows.Shapes.Rectangle]::new(); $bar.Width = [math]::Max(4.0, $x2 - $x1); $bar.Height = $laneHeight - 8; $bar.RadiusX = 2; $bar.RadiusY = 2; $bar.Opacity = 0.9
        $bar.Fill = ConvertTo-OrchestratorBrush $(if ($script:DashboardTimelineColors.ContainsKey($kind)) { $script:DashboardTimelineColors[$kind] } else { '#8DB6E0' })
        $bar.Stroke = [System.Windows.Media.Brushes]::White; $bar.StrokeThickness = 0.5
        $bar.ToolTip = '{0}{1}{2:yyyy-MM-dd HH:mm} - {3:HH:mm} | {4} | {5}' -f $item.Job, [Environment]::NewLine, [datetime]$item.Start, [datetime]$item.End, $kind, $item.Lane
        [System.Windows.Controls.Canvas]::SetLeft($bar, $x1); [System.Windows.Controls.Canvas]::SetTop($bar, $top + $laneIndex * $laneHeight + 4); [void]$Canvas.Children.Add($bar)
    }
}

function Update-OrchestratorServerMetricsChart {
    # Draws, for each server stored in Canvas.Tag, CPU and RAM in percent over the last 12 hours;
    # a gap longer than three samples breaks the line. Redrawn on resize.
    param([Parameter(Mandatory = $true)]$Canvas)
    $Canvas.Children.Clear()
    $servers = @($Canvas.Tag | Where-Object { $null -ne $_ })
    $width = [double]$Canvas.ActualWidth
    $mutedBrush = ConvertTo-OrchestratorBrush '#5D6B78'; $gridBrush = ConvertTo-OrchestratorBrush '#E3E9F0'
    if ($servers.Count -eq 0) {
        $empty = [System.Windows.Controls.TextBlock]::new(); $empty.Text = 'No CPU or memory sample yet (orchestrator 1.5.69 or later samples every 5 minutes).'; $empty.Foreground = $mutedBrush
        [System.Windows.Controls.Canvas]::SetLeft($empty, 8); [System.Windows.Controls.Canvas]::SetTop($empty, 8); [void]$Canvas.Children.Add($empty)
        return
    }
    $gap = 12.0; $panelWidth = ($width - $gap * ($servers.Count - 1)) / $servers.Count
    if ($panelWidth -lt 140) { return }
    $axisWidth = 30.0; $top = 20.0; $bottom = 16.0
    # A canvas without a fixed height (Servers and versions) uses its rendered height.
    $canvasHeight = if ([double]::IsNaN([double]$Canvas.Height)) { [double]$Canvas.ActualHeight } else { [double]$Canvas.Height }
    $plotHeight = $canvasHeight - $top - $bottom
    if ($plotHeight -lt 40) { return }
    for ($index = 0; $index -lt $servers.Count; $index++) {
        $metric = $servers[$index]
        $left = $index * ($panelWidth + $gap); $plotLeft = $left + $axisWidth; $plotWidth = $panelWidth - $axisWidth - 4
        $from = [datetime]$metric.From; $to = [datetime]$metric.To
        $totalMinutes = [math]::Max(1.0, ($to - $from).TotalMinutes)
        $samples = @($metric.Samples)
        $last = if ($samples.Count) { $samples[$samples.Count - 1] } else { $null }
        $title = [System.Windows.Controls.TextBlock]::new(); $title.FontWeight = 'SemiBold'; $title.Width = $panelWidth; $title.TextTrimming = 'CharacterEllipsis'
        $title.Text = if ($null -ne $last) { '{0}  CPU {1}%  RAM {2}%' -f $metric.Server, $last.Cpu, $last.Memory } else { '{0}  no sample' -f $metric.Server }
        [System.Windows.Controls.Canvas]::SetLeft($title, $left); [System.Windows.Controls.Canvas]::SetTop($title, 0); [void]$Canvas.Children.Add($title)
        foreach ($percent in @(0, 50, 100)) {
            $y = $top + $plotHeight - $percent / 100.0 * $plotHeight
            $line = [System.Windows.Shapes.Line]::new(); $line.X1 = $plotLeft; $line.X2 = $plotLeft + $plotWidth; $line.Y1 = $y; $line.Y2 = $y; $line.Stroke = $gridBrush
            [void]$Canvas.Children.Add($line)
            $label = [System.Windows.Controls.TextBlock]::new(); $label.Text = "$percent%"; $label.FontSize = 10; $label.Foreground = $mutedBrush
            [System.Windows.Controls.Canvas]::SetLeft($label, $left); [System.Windows.Controls.Canvas]::SetTop($label, $y - 7); [void]$Canvas.Children.Add($label)
        }
        $tick = $from.Date.AddHours($from.Hour + 1)
        while ($tick -lt $to) {
            if ($tick.Hour % 3 -eq 0) {
                $x = $plotLeft + ($tick - $from).TotalMinutes / $totalMinutes * $plotWidth
                $hour = [System.Windows.Controls.TextBlock]::new(); $hour.Text = $tick.ToString('HH:mm'); $hour.FontSize = 10; $hour.Foreground = $mutedBrush
                [System.Windows.Controls.Canvas]::SetLeft($hour, $x - 14); [System.Windows.Controls.Canvas]::SetTop($hour, $top + $plotHeight + 1); [void]$Canvas.Children.Add($hour)
            }
            $tick = $tick.AddHours(1)
        }
        $maxGapMinutes = 3 * [math]::Max(1, [int]$metric.SampleMinutes) + 1
        foreach ($series in @(@{ Name = 'Cpu'; Color = '#0078D4' }, @{ Name = 'Memory'; Color = '#107C10' })) {
            $line = $null; $previous = $null
            foreach ($sample in $samples) {
                $value = $sample.($series.Name)
                if ($null -eq $value) { $previous = $null; continue }
                $time = [datetime]$sample.Time
                if ($null -eq $line -or $null -eq $previous -or ($time - $previous).TotalMinutes -gt $maxGapMinutes) {
                    $line = [System.Windows.Shapes.Polyline]::new(); $line.Stroke = ConvertTo-OrchestratorBrush $series.Color; $line.StrokeThickness = 1.6
                    $line.ToolTip = '{0} {1}' -f $metric.Server, $(if ($series.Name -eq 'Cpu') { 'CPU %' } else { 'RAM %' })
                    [void]$Canvas.Children.Add($line)
                }
                $x = $plotLeft + ($time - $from).TotalMinutes / $totalMinutes * $plotWidth
                $y = $top + $plotHeight - [math]::Min(100.0, [math]::Max(0.0, [double]$value)) / 100.0 * $plotHeight
                $line.Points.Add([System.Windows.Point]::new($x, $y))
                $previous = $time
            }
        }
    }
}

if ($Demo) {
    . (Join-Path $PSScriptRoot 'SmartM365-Inventory-Orchestrator-Demo.ps1')
    Invoke-OrchestratorDemo -Xaml $xaml -Scenario $DemoScenario -View $DemoView -Width $DemoWindowWidth -Height $DemoWindowHeight -ValidateOnly:$ValidateOnly
    return
}

$script:GuiSplash = $null
if (-not $ValidateOnly -and -not $SmokeTest) {
    if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne [System.Threading.ApartmentState]::STA) {
        throw 'The GUI must be launched from an STA PowerShell process. Use the provided launcher.'
    }
    $splashPath = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365.GuiSplash.ps1'
    if (Test-Path -LiteralPath $splashPath) {
        . $splashPath
        $script:GuiSplash = Start-SmartM365GuiSplash `
            -ProductName 'SmartM365 Orchestrator' `
            -Subtitle 'Central planning and execution history' `
            -LogoPath (Join-Path $PSScriptRoot 'WorkplaceCloudHub-lockup-WPF.png') `
            -WindowIconPath (Join-Path $PSScriptRoot 'WorkplaceCloudHub.ico')
    }
}
$script:SplashReadyMs = $script:StartupClock.ElapsedMilliseconds
$managementModulePath = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365.Orchestrator.Management.psm1'
Import-Module -Name $managementModulePath -Force -ErrorAction Stop
Import-Module -Name (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.JobStop.psm1') -Force -ErrorAction Stop
Import-Module -Name (Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365.Orchestrator.Insights.psm1') -Force -ErrorAction Stop
Import-Module -Name (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.Maintenance.psm1') -ErrorAction Stop
Import-Module -Name (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.Pipeline.psm1') -ErrorAction Stop
Import-Module -Name (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.GuiWorker.psm1') -ErrorAction Stop
Import-Module -Name (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.ConfigView.psm1') -Force -ErrorAction Stop
Import-Module -Name (Join-Path $PSScriptRoot 'SmartM365.Inventory.Launch.psm1') -Force -ErrorAction Stop
$script:SmartInventoryRoot = Split-Path -Path $PSScriptRoot -Parent

if ($ValidateOnly) {
    $validationWindow = ConvertFrom-OrchestratorGuiXaml -Text $xaml
    foreach ($controlName in @('MainTabs', 'PlanningGrid', 'JobArgumentsBox', 'HistoryGrid', 'ServersGrid', 'VersionsGrid', 'PublishButton', 'RebalanceButton', 'RequestPrepareButton', 'ApplyJobButton', 'ApplyServerButton', 'ServerActiveCheck', 'TestServerButton', 'TestServerAsServiceAccountButton', 'ServerTestHintText', 'RollbackButton', 'DaysPanel', 'MondayCheck', 'TuesdayCheck', 'WednesdayCheck', 'ThursdayCheck', 'FridayCheck', 'SaturdayCheck', 'SundayCheck',
        'AutoRefreshCheck', 'OperationsServersGrid', 'OperationsRunningGrid', 'JobStopReasonBox', 'StopRunningJobButton', 'OperationsPendingGrid', 'OperationsIncidentsGrid', 'OperationsMailsGrid',
        'DependsOnBox', 'DependencyModeCombo', 'DependencyMaxAgeBox', 'DependentsText', 'ReadinessGrid', 'IncludeDependenciesCheck', 'RequestRunButton',
        'Failures24hButton', 'RequestsGrid', 'RequestJobsGrid', 'CancellationReasonBox', 'CancellationProgressText', 'CancelRequestButton', 'CancelAllRequestsButton', 'MaintenanceBannerText', 'MaintenanceDetailText',
        'MaintenanceReasonBox', 'EnableMaintenanceButton', 'DisableMaintenanceButton',
        'EventsServerCombo', 'EventsLevelCombo', 'EventsCategoryCombo', 'EventsPeriodCombo', 'EventsTextBox', 'EventsGroupCheck', 'EventsRefreshButton', 'EventsSummaryText', 'EventsGrid', 'EventsDetailBox',
        'DashboardBannerBorder', 'DashboardBannerText', 'DashboardTilesList', 'DashboardActionsGroup', 'DashboardActionsGrid', 'DashboardGrid', 'DashboardTimelineCanvas', 'DashboardMetricsCanvas', 'ServersMetricsCanvas', 'DashboardLoadingOverlay', 'DashboardLoadingText',
        'ConfigFilterBox', 'ConfigShowSensitiveButton', 'ConfigRefreshButton', 'ConfigOpenFolderButton', 'ConfigSummaryText', 'ConfigFileTabs', 'ConfigGlobalTab', 'ConfigTenantTab', 'ConfigMailTab', 'ConfigGlobalStatusText', 'ConfigTenantStatusText', 'ConfigMailStatusText', 'ConfigGlobalGrid', 'ConfigTenantGrid', 'ConfigMailGrid',
        'OrchestratorStartButton', 'OrchestratorStopButton', 'OrchestratorRestartButton', 'FullCollectButton', 'SendSummaryButton', 'OperationsTabs', 'OperationsRequestsTab', 'OnPremModeText', 'JobExecutionModeCombo', 'ApplyJobExecutionModeButton', 'ApplyModeToAllOnPremButton', 'LaunchJobsGrid', 'RunJobNowButton', 'RunJobAsServiceAccountButton', 'RequestJobButton', 'LaunchIncludeDependenciesCheck', 'LaunchJobCommandBox', 'ManualLaunchesGrid', 'RunManualLaunchButton', 'ManualLaunchCommandBox')) {
        if (-not $validationWindow.FindName($controlName)) { throw "Required XAML control not found: $controlName" }
    }
    $source = [IO.File]::ReadAllText($PSCommandPath)
    $registration = [regex]::Match($source, '(?s)foreach \(\$name in @\((?<Names>.*?)\)\) \{\s*\$script:Controls\[\$name\]')
    if (-not $registration.Success) { throw 'GUI control registration list could not be inspected.' }
    $registeredNames = @([regex]::Matches($registration.Groups['Names'].Value, "'(?<Name>[A-Za-z][A-Za-z0-9_]*)'") | ForEach-Object { $_.Groups['Name'].Value })
    $referencedNames = @([regex]::Matches($source, '\$script:Controls\.(?<Name>[A-Za-z][A-Za-z0-9_]*)') | ForEach-Object { $_.Groups['Name'].Value } | Sort-Object -Unique)
    foreach ($controlName in $referencedNames) {
        if ($controlName -notin $registeredNames) { throw "GUI control referenced but not registered: $controlName" }
    }
    $jobsTemplate = Read-SmartM365OrchestratorJson -Path (Join-Path $PSScriptRoot 'Orchestrator-Jobs.json.template')
    $clusterTemplate = Read-SmartM365OrchestratorJson -Path (Join-Path $PSScriptRoot 'Orchestrator-Cluster.json.template')
    $jobsValidation = Test-SmartM365OrchestratorJobsDocument -Document $jobsTemplate
    $clusterValidation = Test-SmartM365OrchestratorClusterDocument -Document $clusterTemplate
    if (-not $jobsValidation.Valid) { throw "Jobs template validation failed: $($jobsValidation.Errors -join '; ')" }
    if (-not $clusterValidation.Valid) { throw "Cluster template validation failed: $($clusterValidation.Errors -join '; ')" }
    foreach ($job in @($jobsTemplate.Jobs)) {
        [void](Get-OrchestratorGuiPropertyValue -Object $job.Schedule -Name 'DaysOfWeek' -DefaultValue @())
        [void](Get-OrchestratorGuiPropertyValue -Object $job.Schedule -Name 'MissedRunPolicy' -DefaultValue 'RunOnce')
        [void](Get-OrchestratorGuiPropertyValue -Object $job -Name 'AssignmentMode' -DefaultValue 'Legacy')
        [void](Get-OrchestratorGuiPropertyValue -Object $job -Name 'AllowedServers' -DefaultValue @())
    }
    "[{0}] VALIDATION_OK SmartM365 Orchestrator GUI v{1}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $script:AppVersion
    $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
}

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne [System.Threading.ApartmentState]::STA) {
    throw 'The GUI must be launched from an STA PowerShell process. Use the provided launcher.'
}

function Write-GuiActivity {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')][string]$Level = 'INFO',
        [switch]$DisplayOnly
    )
    foreach ($physicalLine in @($Message -split '\r?\n')) {
        # The Activity tab was removed in GUI 1.6.13: lines go to the console and the GUI log only.
        $line = '[{0}][{1}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Level, $physicalLine
        Microsoft.PowerShell.Utility\Write-Host $line
        if (-not $DisplayOnly -and -not [string]::IsNullOrWhiteSpace($script:GuiLogPath)) {
            try {
                Write-SmartM365OrchestratorManagementLog -Path $script:GuiLogPath -Message $physicalLine -Level $Level
            }
            catch {
                if (-not $script:GuiLogWriteWarningShown) {
                    $script:GuiLogWriteWarningShown = $true
                    Microsoft.PowerShell.Utility\Write-Host ("[{0}][WARN] Persistent GUI log write failed for '{1}': {2}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $script:GuiLogPath, $_.Exception.Message)
                }
            }
        }
    }
}

function Write-GuiException {
    param(
        [Parameter(Mandatory = $true)][string]$Context,
        [Parameter(Mandatory = $true)][System.Management.Automation.ErrorRecord]$ErrorRecord
    )

    Write-GuiActivity -Message ("{0}: {1}" -f $Context, $ErrorRecord.Exception.Message) -Level ERROR
    $hresult = '0x{0:X8}' -f ($ErrorRecord.Exception.HResult -band 0xffffffffL)
    Write-GuiActivity -Message ("ExceptionType={0}; HResult={1}" -f $ErrorRecord.Exception.GetType().FullName, $hresult) -Level ERROR
    if ($ErrorRecord.Exception.InnerException) {
        Write-GuiActivity -Message ("InnerException={0}: {1}" -f $ErrorRecord.Exception.InnerException.GetType().FullName, $ErrorRecord.Exception.InnerException.Message) -Level ERROR
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$ErrorRecord.ScriptStackTrace)) {
        Write-GuiActivity -Message ("ScriptStackTrace={0}" -f ([string]$ErrorRecord.ScriptStackTrace -replace '\r?\n', ' | ')) -Level ERROR
    }
}

function Initialize-GuiLogPath {
    param([string]$PreferredFolderPath)

    $folders = [Collections.Generic.List[string]]::new()
    if (-not [string]::IsNullOrWhiteSpace($PreferredFolderPath)) { $folders.Add($PreferredFolderPath) | Out-Null }
    $localFallback = Join-Path -Path $env:LOCALAPPDATA -ChildPath 'SmartM365\Logs\SmartM365-Orchestrator-GUI'
    if ($localFallback -notin $folders) { $folders.Add($localFallback) | Out-Null }

    foreach ($folder in $folders) {
        try {
            New-Item -ItemType Directory -Path $folder -Force -ErrorAction Stop | Out-Null
            $probePath = Join-Path -Path $folder -ChildPath ('.write-probe-{0}.tmp' -f [guid]::NewGuid().ToString('N'))
            try {
                [IO.File]::WriteAllText($probePath, 'probe', [Text.UTF8Encoding]::new($false))
            }
            finally {
                Remove-Item -LiteralPath $probePath -Force -ErrorAction SilentlyContinue
            }
            return Join-Path -Path $folder -ChildPath ('SmartM365-Orchestrator-GUI_{0}.log' -f (Get-Date).ToString('yyyyMMdd'))
        }
        catch {
            Microsoft.PowerShell.Utility\Write-Host ("[{0}][WARN] GUI log folder is unavailable: {1}; {2}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $folder, $_.Exception.Message)
        }
    }
    return ''
}

function Set-JsonProperty {
    param(
        [Parameter(Mandatory = $true)]$Object,
        [Parameter(Mandatory = $true)][string]$Name,
        [AllowNull()]$Value
    )
    if ($Object.PSObject.Properties[$Name]) { $Object.$Name = $Value }
    else { $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value }
}

function Copy-JsonDocument {
    param([Parameter(Mandatory = $true)]$Document)
    return ($Document | ConvertTo-Json -Depth 100 | ConvertFrom-Json -Depth 100)
}

function Get-ConfigValue {
    param($Config, [string]$Name, $DefaultValue)
    $property = $Config.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        if ($property.Value -isnot [string]) { return $property.Value }
        if ($property.Value.Trim() -and $property.Value.Trim() -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) { return $property.Value }
    }
    $globalProperty = $script:EffectiveConfig.PSObject.Properties[$Name]
    if ($null -ne $globalProperty -and $null -ne $globalProperty.Value) { return $globalProperty.Value }
    return $DefaultValue
}

function Resolve-ConfigTokens {
    param([AllowNull()]$Value)
    if ($Value -isnot [string]) { return $Value }
    $result = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $matches = [regex]::Matches($result, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($matches.Count -eq 0) { break }
        $changed = $false
        foreach ($match in $matches) {
            $property = $script:EffectiveConfig.PSObject.Properties[$match.Groups['Name'].Value]
            if ($null -eq $property -or $null -eq $property.Value) { continue }
            $result = $result.Replace($match.Value, [string]$property.Value)
            $changed = $true
        }
        if (-not $changed) { break }
    }
    return $result
}

function Get-ComboText {
    param($Combo)
    if ($null -eq $Combo.SelectedItem) { return ([string]$Combo.Text).Trim() }
    if ($Combo.SelectedItem -is [System.Windows.Controls.ComboBoxItem]) { return ([string]$Combo.SelectedItem.Content).Trim() }
    return ([string]$Combo.SelectedItem).Trim()
}

function Select-ComboText {
    param($Combo, [string]$Text)
    $matched = $false
    foreach ($item in @($Combo.Items)) {
        $value = if ($item -is [System.Windows.Controls.ComboBoxItem]) { [string]$item.Content } else { [string]$item }
        if ($value -ieq $Text) { $Combo.SelectedItem = $item; $matched = $true; break }
    }
    if (-not $matched) { $Combo.Text = $Text }
}

function Get-ElectionOwners {
    $owners = @{}
    $planPath = Join-Path -Path $script:SharedDataFolderPath -ChildPath 'Election\Orchestrator-ElectionPlan.json'
    $planPath=Get-SmartM365JsonReadPath $planPath -Optional
    if (-not $planPath) { return $owners }
    try {
        $plan = Read-SmartM365OrchestratorJson -Path $planPath
        foreach ($assignment in @($plan.Assignments)) { $owners[[string]$assignment.JobName] = [string]$assignment.OwnerServer }
    }
    catch { Write-GuiException -Context 'Election plan could not be read' -ErrorRecord $_ }
    return $owners
}

function Get-NextRunText {
    param($Job)
    if (-not [bool]$Job.Enabled) { return 'Disabled' }
    if ([string](Get-OrchestratorGuiPropertyValue -Object $Job -Name 'AssignmentMode' -DefaultValue 'Legacy') -eq 'Manual') { return 'Manual' }
    $now = Get-Date
    for ($offset = 0; $offset -le 8; $offset++) {
        $day = $now.Date.AddDays($offset)
        $scheduleType = [string](Get-OrchestratorGuiPropertyValue -Object $Job.Schedule -Name 'Type' -DefaultValue 'Daily')
        $daysOfWeek = @(Get-OrchestratorGuiPropertyValue -Object $Job.Schedule -Name 'DaysOfWeek' -DefaultValue @())
        if ($scheduleType -eq 'Weekly' -and [string]$day.DayOfWeek -notin $daysOfWeek) { continue }
        foreach ($timeText in @(Get-OrchestratorGuiPropertyValue -Object $Job.Schedule -Name 'Times' -DefaultValue @() | Sort-Object)) {
            $time = [timespan]::Zero
            if (-not [timespan]::TryParseExact([string]$timeText, 'hh\:mm', [System.Globalization.CultureInfo]::InvariantCulture, [ref]$time)) { continue }
            $candidate = $day.Add($time)
            if ($candidate -ge $now) { return $candidate.ToString('yyyy-MM-dd HH:mm') }
        }
    }
    return 'No occurrence'
}

function Set-PlanningDefaultSort {
    $grid = $script:Controls.PlanningGrid
    foreach ($column in @($grid.Columns)) { $column.SortDirection = $null }
    $jobColumn = @($grid.Columns | Where-Object { [string]$_.Header -eq 'Job' })[0]
    if ($null -ne $jobColumn) { $jobColumn.SortDirection = [System.ComponentModel.ListSortDirection]::Ascending }

    $view = [System.Windows.Data.CollectionViewSource]::GetDefaultView($grid.ItemsSource)
    if ($null -ne $view -and $view.CanSort) {
        $view.SortDescriptions.Clear()
        $view.SortDescriptions.Add([System.ComponentModel.SortDescription]::new('Name', [System.ComponentModel.ListSortDirection]::Ascending))
        $view.Refresh()
    }
}

function Update-PinnedServerControlState {
    $mode = Get-ComboText -Combo $script:Controls.AssignmentCombo
    $isPinned = $mode -eq 'Pinned'
    $script:Controls.PinnedServerCombo.IsEnabled = $isPinned
    if (-not $isPinned) {
        $script:Controls.PinnedServerCombo.SelectedIndex = -1
        $script:Controls.PinnedServerCombo.Text = ''
    }
}

function Get-ScheduleDayDefinitions {
    return @(
        [pscustomobject]@{ Day = 'Monday'; Control = 'MondayCheck' }
        [pscustomobject]@{ Day = 'Tuesday'; Control = 'TuesdayCheck' }
        [pscustomobject]@{ Day = 'Wednesday'; Control = 'WednesdayCheck' }
        [pscustomobject]@{ Day = 'Thursday'; Control = 'ThursdayCheck' }
        [pscustomobject]@{ Day = 'Friday'; Control = 'FridayCheck' }
        [pscustomobject]@{ Day = 'Saturday'; Control = 'SaturdayCheck' }
        [pscustomobject]@{ Day = 'Sunday'; Control = 'SundayCheck' }
    )
}

function Get-SelectedScheduleDays {
    return @(
        foreach ($definition in @(Get-ScheduleDayDefinitions)) {
            if ([bool]$script:Controls[$definition.Control].IsChecked) { [string]$definition.Day }
        }
    )
}

function Set-SelectedScheduleDays {
    param([string[]]$Days = @())

    foreach ($definition in @(Get-ScheduleDayDefinitions)) {
        $script:Controls[$definition.Control].IsChecked = [string]$definition.Day -in @($Days)
    }
}

function Update-ScheduleDaysControlState {
    $isWeekly = (Get-ComboText -Combo $script:Controls.ScheduleTypeCombo) -eq 'Weekly'
    $script:Controls.DaysPanel.IsEnabled = $isWeekly
    if (-not $isWeekly) { Set-SelectedScheduleDays -Days @() }
}

function Get-GuiJobScriptPaths {
    param([AllowNull()]$JobsDocument)
    $paths = foreach ($job in @(Get-OrchestratorGuiPropertyValue -Object $JobsDocument -Name 'Jobs' -DefaultValue @())) { [string](Get-OrchestratorGuiPropertyValue -Object $job -Name 'ScriptPath' -DefaultValue '') }
    return @($paths | Where-Object { $_ })
}

function Get-GuiJobScriptVersion {
    # Version of the script present on this computer; empty until the first Full refresh read it.
    param([AllowNull()]$Job)
    $path = [string](Get-OrchestratorGuiPropertyValue -Object $Job -Name 'ScriptPath' -DefaultValue '')
    if (-not $path -or -not $script:ScriptVersions.ContainsKey($path.Trim())) { return '' }
    return [string]$script:ScriptVersions[$path.Trim()]
}

function Refresh-PlanningView {
    param([AllowNull()]$Owners = $null)
    $owners = if ($null -ne $Owners) { $Owners } else { Get-ElectionOwners }
    $rows = foreach ($job in @($script:DraftJobs.Jobs)) {
        $mode = if ($job.PSObject.Properties['AssignmentMode']) { [string]$job.AssignmentMode } else { 'Legacy' }
        $arguments = [string](Get-OrchestratorGuiPropertyValue -Object $job -Name 'Arguments' -DefaultValue '')
        $allowedServers = @(Get-OrchestratorGuiPropertyValue -Object $job -Name 'AllowedServers' -DefaultValue @())
        $server = switch ($mode) {
            'Pinned' { $allowedServers -join ', ' }
            'Elected' { if ($owners.ContainsKey([string]$job.Name)) { $owners[[string]$job.Name] } else { 'Not elected' } }
            'Manual' { 'Manual' }
            default { $allowedServers -join ', ' }
        }
        $health = if ($script:HealthByName.ContainsKey([string]$job.Name)) { $script:HealthByName[[string]$job.Name] } else { $null }
        $swi = Get-GuiSwiRow -JobName ([string]$job.Name)
        [pscustomobject]@{
            Name = [string]$job.Name
            Version = Get-GuiJobScriptVersion -Job $job
            Arguments = if ([string]::IsNullOrWhiteSpace($arguments)) { '(none)' } else { $arguments }
            Enabled = [bool]$job.Enabled
            Frequency = [string]$job.Schedule.Type
            Times = @($job.Schedule.Times) -join ', '
            Days = @(Get-OrchestratorGuiPropertyValue -Object $job.Schedule -Name 'DaysOfWeek' -DefaultValue @()) -join ', '
            Assignment = $mode
            Server = $server
            NextRun = Get-NextRunText -Job $job
            Group = [string]$job.Group
            Health = if ($health) { $health.Health } else { '' }
            LastStatus = if ($health) { $health.LastStatus } else { '' }
            LastSuccess = if ($health) { $health.LastSuccess } else { '' }
            SuccessAgeHours = if ($health) { $health.SuccessAgeHours } else { $null }
            ExpectedMaxAgeHours = if ($health) { $health.ExpectedMaxAgeHours } else { $null }
            AverageDurationMinutes = if ($health) { $health.AverageDurationMinutes } else { $null }
            Swi = $swi.Swi; SwiDaysLeft = $swi.DaysLeft; SwiDetail = $swi.Detail; SwiBlocks = [bool]$swi.Blocks
        }
    }
    $script:PlanningRows = @($rows | Sort-Object -Property Name)
    $script:Controls.PlanningGrid.ItemsSource = $script:PlanningRows
    Set-PlanningDefaultSort
}

function Refresh-ServersView {
    param([AllowNull()]$Servers = $null)
    if ($null -eq $Servers) { $Servers = @(Get-SmartM365OrchestratorServerStatus -SharedDataFolderPath $script:SharedDataFolderPath -ClusterDocument $script:DraftCluster) }
    $servers = @($Servers)
    $selected = $script:Controls.ServersGrid.SelectedItem
    $selectedName = if ($null -ne $selected) { [string]$selected.Server } else { '' }
    $script:ServerSelectionRestoring = $true
    try {
        $script:Controls.ServersGrid.ItemsSource = $servers
        if ($selectedName) { $script:Controls.ServersGrid.SelectedItem = $servers | Where-Object { [string]$_.Server -eq $selectedName } | Select-Object -First 1 }
    }
    finally { $script:ServerSelectionRestoring = $false }
    $script:Controls.PinnedServerCombo.ItemsSource = @($script:DraftCluster.ExpectedOrchestratorServers)
    return $servers
}

function Refresh-HistoryView {
    Start-GuiRefresh -Mode History
}

function Set-GuiSwiValidity {
    # Keeps the previous values when a refresh cannot read the SWI inputs; one message per new error.
    param([AllowNull()]$Validity, [string]$ErrorText = '')
    if ($null -ne $Validity) {
        $byName = @{}
        foreach ($row in @($Validity.Rows)) { $byName[[string]$row.Name] = $row }
        $script:SwiByName = $byName
        if (@($Validity.Unmapped).Count) { Write-GuiActivity -Level WARN -Message ("SWI inputs without a producer job: {0}" -f (@($Validity.Unmapped) -join ', ')) }
        $script:SwiValidityError = ''
        return
    }
    if ($ErrorText -and $ErrorText -ne $script:SwiValidityError) { Write-GuiActivity -Level WARN -Message "SWI validity unavailable: $ErrorText" }
    $script:SwiValidityError = $ErrorText
}

function Get-GuiSwiRow {
    param([string]$JobName)
    if ($script:SwiByName.ContainsKey($JobName)) { return $script:SwiByName[$JobName] }
    $text = if ($script:SwiValidityError) { 'Unavailable' } else { '' }
    return [pscustomobject]@{ Swi = $text; DaysLeft = $null; Blocks = $false; Detail = $(if ($script:SwiValidityError) { "SWI validity unavailable: $($script:SwiValidityError)" } else { '' }) }
}

function Get-GuiJobDurationEstimate {
    # Same source as the Dashboard timeline: average successful duration of the recent runs
    # (job health), else the catalog EstimatedDurationMinutes.
    param([Parameter(Mandatory = $true)][string]$JobName)
    if ($script:HealthByName.ContainsKey($JobName)) {
        $average = Get-OrchestratorGuiPropertyValue -Object $script:HealthByName[$JobName] -Name 'AverageDurationMinutes' -DefaultValue $null
        if ($null -ne $average -and [double]$average -gt 0) { return [pscustomobject]@{ Minutes = [math]::Max(1, [math]::Round([double]$average)); Source = 'Average of the recent successful runs' } }
    }
    $job = if ($null -ne $script:DraftJobs) { @(Get-OrchestratorGuiPropertyValue -Object $script:DraftJobs -Name 'Jobs' -DefaultValue @() | Where-Object { [string]$_.Name -eq $JobName })[0] } else { $null }
    $estimate = Get-OrchestratorGuiPropertyValue -Object $job -Name 'EstimatedDurationMinutes' -DefaultValue $null
    if ($null -ne $estimate -and [double]$estimate -gt 0) { return [pscustomobject]@{ Minutes = [math]::Max(1, [math]::Round([double]$estimate)); Source = 'Catalog EstimatedDurationMinutes (no recent successful run)' } }
    return $null
}

function Add-GuiJobDurationEstimates {
    # Running rows get the estimated time left, pending rows the estimated duration.
    param([AllowNull()]$Operations)
    if ($null -eq $Operations) { return }
    foreach ($row in @(Get-OrchestratorGuiPropertyValue -Object $Operations -Name 'Running' -DefaultValue @())) {
        $estimate = Get-GuiJobDurationEstimate -JobName ([string]$row.Job)
        $elapsed = Get-OrchestratorGuiPropertyValue -Object $row -Name 'DurationMinutes' -DefaultValue $null
        $text = if ($null -eq $estimate) { '' }
            elseif ($null -eq $elapsed) { '~{0} min' -f $estimate.Minutes }
            elseif ([double]$elapsed -le $estimate.Minutes) { '~{0} min' -f [math]::Max(0, $estimate.Minutes - [double]$elapsed) }
            else { 'over by {0} min' -f [math]::Round([double]$elapsed - $estimate.Minutes) }
        $row | Add-Member -NotePropertyName EstimatedRemaining -NotePropertyValue $text -Force
        # Estimated end: start time plus the estimate (same date format as Started); past it, the run is overdue.
        $endText = ''
        if ($null -ne $estimate) {
            $startedAt = [datetime]::MinValue
            $endAt = if ([datetime]::TryParseExact([string](Get-OrchestratorGuiPropertyValue -Object $row -Name 'Started' -DefaultValue ''), 'yyyy-MM-dd HH:mm', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$startedAt)) { $startedAt.AddMinutes($estimate.Minutes) }
                elseif ($null -ne $elapsed) { (Get-Date).AddMinutes($estimate.Minutes - [double]$elapsed) }
                else { (Get-Date).AddMinutes($estimate.Minutes) }
            $endText = if ($endAt -lt (Get-Date)) { 'overdue since ' + $endAt.ToString('yyyy-MM-dd HH:mm') } else { $endAt.ToString('yyyy-MM-dd HH:mm') }
        }
        $row | Add-Member -NotePropertyName EstimatedEnd -NotePropertyValue $endText -Force
        $row | Add-Member -NotePropertyName EstimateSource -NotePropertyValue $(if ($null -ne $estimate) { '{0}: {1} min' -f $estimate.Source, $estimate.Minutes } else { 'No duration history or catalog estimate' }) -Force
    }
    foreach ($row in @(Get-OrchestratorGuiPropertyValue -Object $Operations -Name 'Pending' -DefaultValue @())) {
        $estimate = Get-GuiJobDurationEstimate -JobName ([string]$row.Job)
        $row | Add-Member -NotePropertyName EstimatedMinutes -NotePropertyValue $(if ($null -ne $estimate) { $estimate.Minutes } else { $null }) -Force
        $row | Add-Member -NotePropertyName EstimateSource -NotePropertyValue $(if ($null -ne $estimate) { $estimate.Source } else { 'No duration history or catalog estimate' }) -Force
    }
}

function Update-JobHealth {
    param([AllowNull()]$Runs = $null)
    if ($null -eq $Runs) { $Runs = @(Get-SmartM365OrchestratorRecentRuns -SharedDataFolderPath $script:SharedDataFolderPath -Days 21) }
    $script:RecentRuns = @($Runs)
    $script:HealthByName = @{}
    foreach ($health in @(Get-SmartM365OrchestratorJobHealth -JobsDocument $script:DraftJobs -Runs $script:RecentRuns)) { $script:HealthByName[[string]$health.Name] = $health }
}

function Refresh-OperationsView {
    try {
        $publishedCluster = Read-SmartM365OrchestratorJson -Path (Join-Path $script:SharedDataFolderPath 'Config/Orchestrator-Cluster.json')
        $operations = Get-SmartM365OrchestratorOperations -SharedDataFolderPath $script:SharedDataFolderPath -ClusterDocument $publishedCluster -MailFolderPath $script:MailFolderPath -MailHours 24
        Refresh-MaintenanceView -Operations $operations
        Add-GuiJobDurationEstimates -Operations $operations
        $script:Controls.OperationsServersGrid.ItemsSource = @($operations.Servers)
        $script:Controls.OperationsRunningGrid.ItemsSource = @($operations.Running)
        $script:Controls.StopRunningJobButton.IsEnabled = $false
        $script:Controls.OperationsPendingGrid.ItemsSource = @($operations.Pending)
        $script:Controls.OperationsIncidentsGrid.ItemsSource = @($operations.Incidents)
        $script:Controls.OperationsMailsGrid.ItemsSource = @($operations.Mails)
        return $operations
    }
    catch {
        Refresh-MaintenanceView
        Write-GuiException -Context 'Operations refresh failed' -ErrorRecord $_
        return $null
    }
}

function Request-SelectedRunningJobStop {
    $row = $script:Controls.OperationsRunningGrid.SelectedItem
    if ($null -eq $row -or -not $row.StopSupported) { throw 'Select a running job on an online server that supports job stop requests.' }
    $reason = $script:Controls.JobStopReasonBox.Text.Trim()
    if ($reason.Length -lt 3) { throw 'Enter a stop reason (at least 3 characters).' }
    $message = "Stop this running job?`n`nJob: $($row.Job)`nServer: $($row.Server)`nPID: $($row.Pid)`nStarted: $($row.Started)`n`nReason: $reason`n`nThe owning orchestrator will verify this exact process, stop its process tree, and record Cancelled. This occurrence will not be retried automatically."
    if ($script:OrchestratorDialog.Show($message, 'Stop running job', 'YesNo', 'Warning') -ne 'Yes') { return }
    $request = Request-SmartM365OrchestratorJobStop -SharedDataFolderPath $script:SharedDataFolderPath -Server ([string]$row.Server) -JobName ([string]$row.Job) -ProcessId ([int]$row.Pid) -StartTime ([string]$row.StartTime) -Reason $reason -Tenant $Tenant
    Write-GuiActivity -Message ("Job stop requested: job={0}; server={1}; PID={2}; request={3}; reason={4}." -f $row.Job, $row.Server, $row.Pid, $request.RequestId, $reason) -Level WARN
    $script:Controls.JobStopReasonBox.Clear()
    $script:OrchestratorDialog.Show("Stop request $($request.RequestId) was submitted. The owning orchestrator will process it on its next cycle. Refresh Operations or History to verify the final status.", 'Stop requested', 'OK', 'Information') | Out-Null
}

function Set-GuiMaintenanceBannerStyle {
    # Normal = scheduling runs normally (muted); Attention = paused or in transition (amber); Error = control unreadable (red).
    param([ValidateSet('Normal', 'Attention', 'Error')][string]$Level)
    $banner = $script:Controls.MaintenanceBannerText
    $color = switch ($Level) { 'Normal' { '#5D6B78' } 'Attention' { '#92400E' } 'Error' { '#A4262C' } }
    $banner.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($color)
    $banner.FontWeight = if ($Level -eq 'Normal') { [System.Windows.FontWeights]::Normal } else { [System.Windows.FontWeights]::SemiBold }
}

function Refresh-MaintenanceView {
    param($Operations = $null)
    try {
        if ($null -ne $Operations) {
            if ($Operations.MaintenanceError) { throw $Operations.MaintenanceError }
            $control = $Operations.Maintenance; $servers = @($Operations.MaintenanceServers)
        } else {
            $control = Get-SmartM365OrchestratorMaintenanceState $script:SharedDataFolderPath
            # Published cluster only. Editing a draft cannot hide an unacknowledged server.
            $cluster = Read-SmartM365OrchestratorJson -Path (Join-Path $script:SharedDataFolderPath 'Config/Orchestrator-Cluster.json')
            $servers = @(Get-SmartM365OrchestratorMaintenanceReadiness $script:SharedDataFolderPath $cluster $control)
        }
        $applied = $servers.Count -gt 0 -and @($servers | Where-Object Status -ne 'Applied').Count -eq 0
        # Offline servers whose last heartbeat supports maintenance apply it when they return.
        $activationReady = Test-SmartM365OrchestratorMaintenanceActivationReady -Servers $servers
        $offlineCount = @($servers | Where-Object Status -like 'Offline (*').Count
        $script:MaintenanceViewState = $control
        $script:Controls.EnableMaintenanceButton.IsEnabled = (-not $control.Enabled -and $activationReady)
        $script:Controls.DisableMaintenanceButton.IsEnabled = $control.Enabled
        $stateText = if ($control.Enabled) { if ($applied) { 'ACTIVE' } elseif ($activationReady) { "ACTIVE ($offlineCount offline server(s) apply it on return)" } else { 'ACTIVATION REQUESTED' } }
            elseif ($applied) { 'INACTIVE' } elseif ($activationReady) { "INACTIVE ($offlineCount offline server(s))" } else { 'ACKNOWLEDGEMENT PENDING' }
        $script:Controls.MaintenanceBannerText.Text = "Maintenance: $stateText | revision $($control.Revision) | scheduled launches only; manual Pipeline requests remain available"
        # Normal scheduling is not a problem: neutral when inactive, amber only while it pauses or is in transition.
        Set-GuiMaintenanceBannerStyle -Level $(if (-not $control.Enabled -and $activationReady) { 'Normal' } else { 'Attention' })
        $running =($servers | Measure-Object Running -Sum).Sum
        $stamp = if ($control.ChangedAtUtc) { ([datetimeoffset]::Parse($control.ChangedAtUtc)).LocalDateTime.ToString('yyyy-MM-dd HH:mm:ss') } else { 'not changed' }
        $script:Controls.MaintenanceDetailText.Text = "Changed: $stamp | By: $($control.ChangedBy) | Reason: $($control.Reason) | Running jobs: $running`n" + (($servers | ForEach-Object { "$($_.Server): $($_.Status)" }) -join ' | ')
    }
    catch {
        $script:MaintenanceViewState = $null
        $script:Controls.EnableMaintenanceButton.IsEnabled = $false
        $script:Controls.DisableMaintenanceButton.IsEnabled = $false
        $script:Controls.MaintenanceBannerText.Text = 'Maintenance: CONTROL UNAVAILABLE - do not assume scheduling is paused or resumed'
        Set-GuiMaintenanceBannerStyle -Level 'Error'
        $script:Controls.MaintenanceDetailText.Text = $_.Exception.Message
    }
}

function Set-GuiMaintenance {
    param([bool]$Enabled)
    if ($null -eq $script:MaintenanceViewState) { throw 'Refresh the maintenance control before changing it.' }
    $reason = $script:Controls.MaintenanceReasonBox.Text.Trim()
    if (-not $reason -and $Enabled) { throw 'Enter a reason before enabling maintenance.' }
    if (-not $reason) { $reason = 'Scheduled planning resumed from GUI' }
    $action = if ($Enabled) { 'Enable' } else { 'Disable' }
    $catchUp = $false
    if ($Enabled) {
        $message = "Enable shared maintenance for this tenant?`n`nAutomatic launches and retries will pause. Running jobs remain supervised and manual Pipeline requests remain available.`n`nReason: $reason"
        if ($script:OrchestratorDialog.Show($message, 'Enable maintenance', 'YesNo', 'Warning') -ne 'Yes') { return }
    }
    else {
        $started = ([datetimeoffset]::Parse($script:MaintenanceViewState.ChangedAtUtc)).ToLocalTime().ToString('yyyy-MM-dd HH:mm')
        $message = "Disable shared maintenance for this tenant?`n`nSuspended since: $started`n`nYES: replay every missed scheduled occurrence for currently enabled jobs, including jobs marked Skip. Each run collects current data; this does not reconstruct historical snapshots. Independent jobs may run in parallel; normal ownership, dependency and concurrency checks apply. This can queue many jobs.`n`nNO: skip occurrences from this maintenance interval and resume future schedules. Earlier confirmed catch-up work remains queued.`n`nCANCEL: keep maintenance enabled.`n`nReason: $reason"
        $choice = $script:OrchestratorDialog.Show($message, 'Disable maintenance', 'YesNoCancel', 'Warning')
        if ($choice -eq 'Cancel') { return }
        $catchUp = $choice -eq 'Yes'
    }
    $result = Set-SmartM365OrchestratorMaintenance -SharedDataFolderPath $script:SharedDataFolderPath -Enabled $Enabled -ExpectedRevision $script:MaintenanceViewState.Revision -Reason $reason -CatchUpMissedOccurrences:$catchUp
    Write-GuiActivity -Message ("Maintenance transition published: revision={0}; enabled={1}; catchUpAll={2}; reason={3}. Awaiting server acknowledgement." -f $result.Revision, $result.Enabled, $result.CatchUpAll, $result.Reason)
    Invoke-AutoRefresh
}

function Update-SelectedPipelineRequest {
    $row = $script:Controls.RequestsGrid.SelectedItem
    # Do not assign an if-expression: its pipeline unwraps a singleton array.
    $jobRows = @()
    if ($null -ne $row) { $jobRows = @($row.JobRows) }
    $script:Controls.RequestJobsGrid.ItemsSource = $jobRows
    $script:Controls.CancelRequestButton.IsEnabled = -not $script:CancellationBusy -and -not $script:RequestSubmitBusy -and $null -ne $row -and $row.Status -in @('Running', 'Queued', 'Cancelling')
}

function Set-GuiRequestRows {
    param([AllowEmptyCollection()][object[]]$Rows)
    $selectedBatch = if ($null -ne $script:Controls.RequestsGrid.SelectedItem) { [string]$script:Controls.RequestsGrid.SelectedItem.BatchId } else { '' }
    $script:RequestRows = @($Rows)
    $script:Controls.RequestsGrid.ItemsSource = $script:RequestRows
    $selectedRow = @($script:RequestRows | Where-Object { [string]$_.BatchId -eq $selectedBatch })
    if ($selectedRow.Count) { $script:Controls.RequestsGrid.SelectedItem = $selectedRow[0] }
    Update-SelectedPipelineRequest
    $script:Controls.CancelAllRequestsButton.IsEnabled = -not $script:CancellationBusy -and -not $script:RequestSubmitBusy -and @($script:RequestRows | Where-Object Status -in @('Running', 'Queued', 'Cancelling')).Count -gt 0
}

function Refresh-RequestsView {
    try {
        if ($script:CancellationBusy -or $script:RequestSubmitBusy) { return }
        Set-GuiRequestRows -Rows @(Get-SmartM365OrchestratorRecentPipelineRuns -SharedDataFolderPath $script:SharedDataFolderPath -Count 30)
    }
    catch { Write-GuiException -Context 'Requests refresh failed' -ErrorRecord $_ }
}

function Confirm-PipelineCancellation {
    param([string]$Message, [string]$Title)
    return $script:OrchestratorDialog.Show($Message, $Title, 'YesNo', 'Warning') -eq 'Yes'
}

function Get-GuiCancellationWork {
    # Only plain input/progress crosses the runspace boundary; no WPF references.
    return {
        param($Progress, $InputData)
        $ErrorActionPreference = 'Stop'
        Set-StrictMode -Version 2.0
        foreach ($path in $InputData.ModulePaths) { Import-Module $path -ErrorAction Stop }
        function Write-CancellationLog {
            param([string]$Message, [string]$Level = 'INFO')
            if ($InputData.LogPath) {
                try { $null = Write-SmartM365OrchestratorManagementLog -Path $InputData.LogPath -Message $Message -Level $Level }
                catch { $Progress.LogError = $_.Exception.Message }
            }
        }
        try {
            if ($InputData.Phase -eq 'Validate') {
                $Progress.Message = 'Reading active requests and checking cancellation readiness...'
                $runs = @()
                if ($InputData.All) { $runs = @(Get-SmartM365OrchestratorActivePipelineRuns -SharedDataFolderPath $InputData.Root) }
                else { $runs = @(Get-SmartM365OrchestratorPipelineRunStatus -SharedDataFolderPath $InputData.Root -BatchId $InputData.BatchIds[0]) }
                if (-not $runs.Count) { throw 'No active pipeline requests remain.' }
                $offlineServers = @()
                foreach ($run in $runs) {
                    $Progress.Message = "Checking batch $($run.BatchId)..."
                    $preview = Stop-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $InputData.Root -BatchId $run.BatchId -Tenant $InputData.Tenant -Reason $InputData.Reason -ValidateOnly
                    if ($preview.PSObject.Properties['OfflineServers']) { $offlineServers += @($preview.OfflineServers) }
                }
                return [pscustomobject]@{ Phase = 'Validate'; BatchIds = @($runs | ForEach-Object BatchId); OfflineServers = @($offlineServers | Sort-Object -Unique) }
            }
            $results = @(); $failure = ''; $refreshError = ''
            foreach ($batch in $InputData.BatchIds) {
                $Progress.Message = "Publishing cancellation for batch $batch..."
                try {
                    # The existing API rechecks live readiness and serializes publication.
                    $result = Stop-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $InputData.Root -BatchId $batch -Tenant $InputData.Tenant -Reason $InputData.Reason
                    $results += $result
                    Write-CancellationLog -Level SUCCESS -Message "Cancellation published. Batch=$batch; Status=$($result.OverallStatus); Cancelled=$($result.CancelledCount); Remaining=$($result.PendingCount)."
                }
                catch { $failure = $_.Exception.Message; Write-CancellationLog -Level ERROR -Message "Cancellation failed for batch $batch : $failure"; break }
            }
            $Progress.Message = 'Refreshing request states...'
            $rows = @()
            try { $rows = @(Get-SmartM365OrchestratorRecentPipelineRuns -SharedDataFolderPath $InputData.Root -Count 30) }
            catch { $refreshError = $_.Exception.Message; Write-CancellationLog -Level WARN -Message "Requests refresh failed: $refreshError" }
            [pscustomobject]@{ Phase = 'Publish'; Results = $results; Error = $failure; Rows = $rows; RefreshError = $refreshError }
        }
        catch { Write-CancellationLog -Level ERROR -Message $_.Exception.Message; throw }
    }
}

function Set-GuiCancellationBusy {
    param([bool]$Busy)
    $script:CancellationBusy = $Busy
    $names = @('RefreshButton','RebalanceButton','PublishButton','HistoryRefreshButton',
        'RequestRunButton','RequestPrepareButton','EnableMaintenanceButton','DisableMaintenanceButton','RollbackButton',
        'CancelRequestButton','CancelAllRequestsButton','CancellationReasonBox')
    if ($Busy) {
        $script:CancellationUiState = @{}
        foreach ($name in $names) {
            if ($script:Controls[$name]) {
                $script:CancellationUiState[$name] = $script:Controls[$name].IsEnabled
                $script:Controls[$name].IsEnabled = $false
            }
        }
        $script:CancellationContext.Clock.Restart()
    } else {
        $script:CancellationContext.Clock.Stop()
        foreach ($name in $script:CancellationUiState.Keys) { $script:Controls[$name].IsEnabled = $script:CancellationUiState[$name] }
        Update-SelectedPipelineRequest
        $script:Controls.CancelAllRequestsButton.IsEnabled = @($script:RequestRows | Where-Object Status -in @('Running', 'Queued', 'Cancelling')).Count -gt 0
    }
}

function Start-GuiPipelineCancellation {
    param([switch]$All, [string]$BatchId = '')
    if ($script:CancellationBusy) { return }
    if ($script:RequestSubmitBusy) { throw 'A run request is being submitted. Wait for its result before cancelling.' }
    $reason = $script:Controls.CancellationReasonBox.Text.Trim()
    if (-not $reason) { throw 'Enter a cancellation reason.' }
    $workerFolder = Split-Path $managementModulePath -Parent
    $script:CancellationContext.Input = [pscustomobject]@{
        Phase = 'Validate'; All = [bool]$All; BatchIds = @($BatchId)
        Root = $script:SharedDataFolderPath; Tenant = $Tenant; Reason = $reason; LogPath = $script:GuiLogPath
        ModulePaths = @($managementModulePath, (Join-Path $workerFolder 'SmartM365.Orchestrator.Pipeline.psm1'),
            (Join-Path $workerFolder 'SmartM365.Orchestrator.Insights.psm1'))
    }
    Set-GuiCancellationBusy $true
    $script:CancellationContext.Progress.Message = 'Starting cancellation checks...'
    $script:CancellationContext.Progress.LogError = ''
    $script:Controls.CancellationProgressText.Text = $script:CancellationContext.Progress.Message
    try {
        if (-not (Start-SmartM365OrchestratorGuiWorker -Context $script:CancellationContext -Work (Get-GuiCancellationWork) -InputObject $script:CancellationContext.Input)) { throw 'A cancellation worker is already active.' }
    }
    catch { Set-GuiCancellationBusy $false; throw }
}

function Stop-SelectedPipelineRequest {
    if ($script:CancellationBusy) { return }
    $row = $script:Controls.RequestsGrid.SelectedItem
    if ($null -eq $row -or $row.Status -notin @('Running', 'Queued', 'Cancelling')) { throw 'Select an active request.' }
    Start-GuiPipelineCancellation -BatchId $row.BatchId
}

function Stop-AllPipelineRequests {
    Start-GuiPipelineCancellation -All
}

function Show-GuiCancellationError {
    param([string]$Message)
    Write-GuiActivity -Message $Message -Level ERROR -DisplayOnly
    $script:OrchestratorDialog.Show($Message, 'Cancellation failed', 'OK', 'Error') | Out-Null
}

function Receive-GuiPipelineCancellation {
    if (-not $script:CancellationBusy) { return }
    $context = $script:CancellationContext
    $script:Controls.CancellationProgressText.Text = '{0} ({1:N0}s)' -f $context.Progress.Message, $context.Clock.Elapsed.TotalSeconds
    $received = Receive-SmartM365OrchestratorGuiWorker -Context $context
    if ($null -eq $received) { return }
    $continue = $false
    try {
        if ($received.Error) { throw $received.Error }
        if ($received.Output.Count -ne 1) { throw 'Cancellation worker returned an invalid result; review live request state before retrying.' }
        $result = $received.Output[0]
        if ($result.Phase -eq 'Validate') {
            $context.Input.BatchIds = @($result.BatchIds)
            $context.Progress.Message = 'Awaiting confirmation...'
            $title = if ($context.Input.All) { 'Cancel All remaining Jobs' } else { 'Cancel remaining jobs' }
            $message = "Cancel remaining jobs in these requests?`n$($context.Input.BatchIds -join "`n")`nPending jobs and retries will not launch. Running collectors continue; completed results and automatic schedules stay unchanged.`nReason: $($context.Input.Reason)"
            if (@($result.OfflineServers).Count) { $message += "`n`nOffline server(s): $(@($result.OfflineServers) -join ', '). Their last heartbeat supports cancellation: they apply it before any launch when they return." }
            if (Confirm-PipelineCancellation -Message $message -Title $title) {
                $context.Input.Phase = 'Publish'
                $context.Progress.Message = 'Starting cancellation publication...'
                $continue = Start-SmartM365OrchestratorGuiWorker -Context $context -Work (Get-GuiCancellationWork) -InputObject $context.Input
                if (-not $continue) { throw 'Cancellation publication worker could not start.' }
            } else { $script:Controls.CancellationProgressText.Text = 'Cancellation declined; no changes published.' }
        } else {
            if (-not $result.RefreshError) { Set-GuiRequestRows -Rows @($result.Rows) }
            foreach ($item in $result.Results) {
                Write-GuiActivity -Level SUCCESS -DisplayOnly -Message "Cancellation published. Batch=$($item.BatchId); Status=$($item.OverallStatus); Cancelled=$($item.CancelledCount); Remaining=$($item.PendingCount)."
            }
            if ($result.Error) { throw "$($result.Error) Already published cancellations remain audited; no running collector was stopped." }
            $script:Controls.CancellationReasonBox.Clear()
            $script:Controls.CancellationProgressText.Text = 'Cancellation published; running collectors continue.'
            if ($result.RefreshError) { $script:Controls.CancellationProgressText.Text += " Display refresh failed; refresh before relying on the grid: $($result.RefreshError)" }
        }
        if ($context.Progress.LogError) { Write-GuiActivity -Level WARN -DisplayOnly -Message "Persistent cancellation log unavailable: $($context.Progress.LogError)" }
    }
    catch {
        $problem = $_
        $script:Controls.CancellationProgressText.Text = 'Cancellation operation failed; review live request state before retrying.'
        Show-GuiCancellationError -Message $problem.Exception.Message
    }
    finally { if (-not $continue) { Set-GuiCancellationBusy $false } }
}

function Invoke-AutoRefresh {
    # No automatic shared reads while a cancellation or a run request is being published.
    if ($script:CancellationBusy -or $script:RequestSubmitBusy) { return }
    Start-GuiRefresh -Mode Live
}

function Show-SelectedJobDependencies {
    param([Parameter(Mandatory = $true)]$Job)
    $name = [string]$Job.Name
    $dependents = @(Get-SmartM365OrchestratorDependents -JobsDocument $script:DraftJobs -JobName $name)
    $enabledDependents = @(Get-SmartM365OrchestratorDependents -JobsDocument $script:DraftJobs -JobName $name -EnabledOnly)
    $script:Controls.DependentsText.Text = if ($dependents.Count) { 'Used by {0} job(s) ({1} enabled): {2}' -f $dependents.Count, $enabledDependents.Count, ($dependents -join ', ') } else { 'No job depends on this job.' }
    try {
        $script:Controls.ReadinessGrid.ItemsSource = @(Get-SmartM365OrchestratorDependencyReadiness -SharedDataFolderPath $script:SharedDataFolderPath -JobsDocument $script:DraftJobs -JobName $name -Runs $script:RecentRuns)
    }
    catch {
        $script:Controls.ReadinessGrid.ItemsSource = @()
        Write-GuiException -Context "Dependency readiness of '$name' could not be read" -ErrorRecord $_
    }
}

function Get-PublishedJobsPath {
    $jobsPath = (Get-SmartM365OrchestratorConfigurationPaths -SharedDataFolderPath $script:SharedDataFolderPath).JobsPath
    $selected = Get-SmartM365JsonReadPath $jobsPath -Optional
    if (-not $selected) { throw "Published jobs configuration not found: $jobsPath" }
    return $selected
}

function Test-JobDraftChanged {
    param([Parameter(Mandatory = $true)][string]$JobName)
    $draftJob = @($script:DraftJobs.Jobs | Where-Object { [string]$_.Name -eq $JobName })
    $publishedJob = @($script:Snapshot.Jobs.Jobs | Where-Object { [string]$_.Name -eq $JobName })
    if ($draftJob.Count -ne 1 -or $publishedJob.Count -ne 1) { return $true }
    return (($draftJob[0] | ConvertTo-Json -Depth 100 -Compress) -cne ($publishedJob[0] | ConvertTo-Json -Depth 100 -Compress))
}

function Request-SelectedJobRun {
    param([switch]$SkipConfirmation, [string]$JobName = '', [AllowNull()]$IncludeDependencies = $null)

    # Every GUI entry (Planning and Operations > Jobs) passes here: one submission at a time.
    if ($script:RequestSubmitBusy) { return $null }
    if ($script:CancellationBusy) { throw 'A cancellation is being checked or published. Wait for its result before requesting a run.' }
    if ([string]::IsNullOrWhiteSpace($JobName)) {
        $row = $script:Controls.PlanningGrid.SelectedItem
        if ($null -eq $row) { throw 'Select a job first.' }
        $JobName = [string]$row.Name
    }
    $jobName = $JobName
    $includeDependencies = if ($null -ne $IncludeDependencies) { [bool]$IncludeDependencies } else { [bool]$script:Controls.IncludeDependenciesCheck.IsChecked }
    if ($SkipConfirmation) {
        # Non-interactive path (smoke test): synchronous, same API and contract.
        # Queueing, duplicate and queue-limit rules are applied under the shared submission lock.
        $request = New-SmartM365OrchestratorJobRunRequest -SharedDataFolderPath $script:SharedDataFolderPath -JobsPath (Get-PublishedJobsPath) -JobName @($jobName) -Tenant $Tenant -IncludeDependencies:$includeDependencies
        Write-GuiActivity -Message ("Run request recorded. Job={0}; IncludeDependencies={1}; BatchId={2}; Jobs={3}; QueuePosition={4}" -f $jobName, $includeDependencies, $request.BatchId, $request.TotalCount, $request.QueuePosition) -Level SUCCESS
        $script:Controls.StatusText.Text = "Run request recorded: $($request.BatchId)"
        Refresh-RequestsView
        return $request
    }
    # Displayed rows only: no shared read on the WPF thread. The submission API decides under
    # the shared submission lock (queue, duplicate, queue limit, resident queue protocol).
    $activeRequests = @($script:RequestRows | Where-Object Status -in @('Running', 'Cancelling', 'Queued'))
    $message = "Ask the orchestrators to run '$jobName' now?"
    if ($includeDependencies) { $message += "`n`nIts enabled dependencies are included and run first." }
    else { $message += "`n`nDependencies are not run; the job starts on its elected server when its dependency rule allows it." }
    if (Test-JobDraftChanged -JobName $jobName) { $message += "`n`nThis job has unpublished draft changes. The request uses the PUBLISHED configuration." }
    if ($activeRequests.Count) { $message += "`n`n$($activeRequests.Count) request(s) are active or queued (last refresh): this request is queued and starts automatically when they have finished." }
    else { $message += "`n`nIf another request becomes active first, this one is queued behind it." }
    # Displayed server rows only (last refresh): a paused server keeps the request queued.
    $serverRows = if ($null -ne $script:Controls['OperationsServersGrid']) { @($script:Controls['OperationsServersGrid'].ItemsSource) } else { @() }
    $diskPausedServers = @($serverRows | Where-Object { $null -ne $_ -and $_.PSObject.Properties['DiskSpace'] -and [string]$_.DiskSpace -eq 'Launches paused' } | ForEach-Object { [string]$_.Server })
    if ($diskPausedServers.Count) { $message += "`n`nLow DATA disk space: new launches are paused on $($diskPausedServers -join ', '). The request waits until free space reaches the resume threshold." }
    if (-not (Confirm-GuiJobRunRequest -Message $message)) { return $null }
    [void](Start-GuiJobRunSubmission -JobName $jobName -IncludeDependencies $includeDependencies)
    return $null
}

function Confirm-GuiJobRunRequest {
    param([string]$Message)
    return $script:OrchestratorDialog.Show($Message, 'Request run', 'YesNo', 'Question') -eq 'Yes'
}

function Show-GuiJobRunRequestResult {
    param([string]$Message, [string]$Title, [ValidateSet('Information', 'Warning', 'Error')][string]$Icon)
    $script:OrchestratorDialog.Show($Message, $Title, 'OK', $Icon) | Out-Null
}

function Get-GuiJobRunSubmissionWork {
    # Only plain input/progress crosses the runspace boundary; no WPF references.
    return {
        param($Progress, $InputData)
        $ErrorActionPreference = 'Stop'
        Set-StrictMode -Version 2.0
        function Write-SubmissionLog {
            param([string]$Message, [string]$Level = 'INFO')
            if ($InputData.LogPath) {
                try { $null = Write-SmartM365OrchestratorManagementLog -Path $InputData.LogPath -Message $Message -Level $Level }
                catch { $Progress.LogError = $_.Exception.Message }
            }
        }
        function Get-BatchFolderNames {
            $runs = Join-Path $InputData.Root 'PipelineRuns'
            if (-not (Test-Path -LiteralPath $runs)) { return @() }
            @(Get-ChildItem -LiteralPath $runs -Directory -ErrorAction Stop | ForEach-Object Name)
        }
        $Progress.Stage = 'Preparing'
        $Progress.Message = 'Loading orchestrator modules...'
        foreach ($path in $InputData.ModulePaths) { Import-Module $path -ErrorAction Stop }
        $Progress.Message = 'Reading the published jobs configuration...'
        $jobsPath = (Get-SmartM365OrchestratorConfigurationPaths -SharedDataFolderPath $InputData.Root).JobsPath
        $publishedJobsPath = Get-SmartM365JsonReadPath $jobsPath -Optional
        if (-not $publishedJobsPath) { throw "Published jobs configuration not found: $jobsPath" }
        $before = @(Get-BatchFolderNames)
        $Progress.Stage = 'Submitting'
        $Progress.Message = 'Recording the request in the shared folder...'
        try {
            $request = New-SmartM365OrchestratorJobRunRequest -SharedDataFolderPath $InputData.Root -JobsPath $publishedJobsPath -JobName @($InputData.JobName) -Tenant $InputData.Tenant -IncludeDependencies:([bool]$InputData.IncludeDependencies)
        }
        catch {
            $failure = $_.Exception.Message
            Write-SubmissionLog -Level ERROR -Message ("Run request failed. Job={0}; Error={1}" -f $InputData.JobName, $failure)
            # Never retried here: report whether a new batch appeared so the operator can decide.
            $Progress.Message = 'Checking whether a request was recorded...'
            $newBatchIds = @(); $checkError = ''
            try { $newBatchIds = @(Get-BatchFolderNames | Where-Object { $_ -notin $before }) }
            catch { $checkError = $_.Exception.Message }
            return [pscustomobject]@{ Outcome = 'Failed'; Error = $failure; NewBatchIds = $newBatchIds; CheckError = $checkError }
        }
        $Progress.Stage = 'Recorded'
        Write-SubmissionLog -Level SUCCESS -Message ("Run request recorded. Job={0}; IncludeDependencies={1}; BatchId={2}; Jobs={3}; QueuePosition={4}" -f $InputData.JobName, [bool]$InputData.IncludeDependencies, $request.BatchId, $request.TotalCount, $request.QueuePosition)
        $Progress.Message = 'Refreshing request states...'
        $rows = @(); $refreshError = ''
        try { $rows = @(Get-SmartM365OrchestratorRecentPipelineRuns -SharedDataFolderPath $InputData.Root -Count 30) }
        catch { $refreshError = $_.Exception.Message; Write-SubmissionLog -Level WARN -Message "Requests refresh failed: $refreshError" }
        [pscustomobject]@{ Outcome = 'Recorded'; BatchId = [string]$request.BatchId; TotalCount = [int]$request.TotalCount; QueuePosition = [int]$request.QueuePosition; QueuedBehind = @($request.QueuedBehind); Rows = $rows; RefreshError = $refreshError }
    }
}

function New-GuiWaitDialog {
    # Modeless and owned: the WPF thread stays free, so the indeterminate bar keeps moving.
    param([string]$Title, [string]$Message, [string]$Detail)
    $dialogXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Please wait" WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        ResizeMode="NoResize" SizeToContent="WidthAndHeight" ShowInTaskbar="False" FontFamily="Segoe UI" UseLayoutRounding="True">
    <Border Margin="18" Background="White" BorderBrush="#D5DEE8" BorderThickness="1" CornerRadius="10" Padding="28,20">
        <Border.Effect><DropShadowEffect BlurRadius="16" ShadowDepth="2" Opacity="0.25"/></Border.Effect>
        <StackPanel Width="380">
            <TextBlock x:Name="WaitTitleText" FontSize="16" FontWeight="SemiBold" Foreground="#17324D" HorizontalAlignment="Center"/>
            <TextBlock x:Name="SubmitMessageText" Margin="0,8,0,0" TextWrapping="Wrap" TextAlignment="Center"/>
            <Border x:Name="SubmitProgressTrack" Height="6" Margin="0,12,0,8" CornerRadius="3" Background="#E6EDF4" ClipToBounds="True">
                <Border x:Name="SubmitProgressBar" Width="110" HorizontalAlignment="Left" CornerRadius="3" Background="#0078D4">
                    <Border.RenderTransform><TranslateTransform X="-110"/></Border.RenderTransform>
                </Border>
            </Border>
            <TextBlock x:Name="SubmitDetailText" Foreground="#5D6B78" FontSize="12" TextWrapping="Wrap" TextAlignment="Center"/>
        </StackPanel>
    </Border>
</Window>
'@
    $dialog = ConvertFrom-OrchestratorGuiXaml -Text $dialogXaml
    $dialog.FindName('WaitTitleText').Text = $Title
    $dialog.FindName('SubmitMessageText').Text = $Message
    $dialog.FindName('SubmitDetailText').Text = $Detail
    # Explicit animation: the theme's indeterminate ProgressBar stays still when Windows
    # client-area animations are off (common on servers and remote sessions).
    $sweep = [System.Windows.Media.Animation.DoubleAnimation]::new(-110, 380, [System.Windows.Duration]::new([TimeSpan]::FromMilliseconds(1400)))
    $sweep.RepeatBehavior = [System.Windows.Media.Animation.RepeatBehavior]::Forever
    $dialog.FindName('SubmitProgressBar').RenderTransform.BeginAnimation([System.Windows.Media.TranslateTransform]::XProperty, $sweep)
    $owner = [System.Windows.Window]::GetWindow($script:Controls.StatusText)
    if ($null -ne $owner -and $owner.IsVisible) {
        $dialog.Owner = $owner
        $dialog.WindowStartupLocation = 'CenterOwner'
    } else { $dialog.WindowStartupLocation = 'CenterScreen' }
    # No cancel: the work may already be done when the operator would press it. Only Close-GuiWaitDialog closes it.
    $dialog.Add_Closing({ param($sender, $eventArgs) if ([string]$sender.Tag -ne 'Closing') { $eventArgs.Cancel = $true } })
    $dialog.Show()
    return $dialog
}

function Close-GuiWaitDialog {
    param($Dialog)
    if ($null -eq $Dialog) { return }
    $Dialog.Tag = 'Closing'
    $Dialog.Close()
}

function Show-GuiRequestSubmitDialog {
    param([string]$JobName)
    Close-GuiRequestSubmitDialog
    $script:RequestSubmitDialog = New-GuiWaitDialog -Title 'Run request' -Message ('Submitting request for {0}{1} Please wait.' -f $JobName, [char]0x2026) -Detail 'Recording the request in the shared folder. The job starts later, on its orchestrator.'
}

function Close-GuiRequestSubmitDialog {
    $dialog = $script:RequestSubmitDialog
    $script:RequestSubmitDialog = $null
    Close-GuiWaitDialog -Dialog $dialog
}

function Update-GuiPrepareRequestButton {
    # Header shortcut: enabled only when Prepare is in the published configuration.
    if (-not $script:Controls['RequestPrepareButton']) { return }
    $rows = if (Test-Path -LiteralPath 'variable:script:LaunchJobRows') { @($script:LaunchJobRows) } else { @() }
    $prepare = @($rows | Where-Object { $null -ne $_ -and [string]$_.Name -eq $script:PrepareJobName -and [bool]$_.InManifest })
    $script:Controls.RequestPrepareButton.IsEnabled = $prepare.Count -gt 0 -and -not $script:RequestSubmitBusy -and -not $script:CancellationBusy
}

function Set-GuiRequestSubmitBusy {
    param([bool]$Busy, [string]$JobName = '')
    $script:RequestSubmitBusy = $Busy
    # Both request entries, plus the actions that would read or change requests concurrently.
    $names = @('RequestRunButton', 'RequestJobButton', 'RequestPrepareButton', 'RefreshButton', 'CancelRequestButton', 'CancelAllRequestsButton')
    if ($Busy) {
        $script:RequestSubmitUiState = @{}
        foreach ($name in $names) {
            if ($script:Controls[$name]) {
                $script:RequestSubmitUiState[$name] = $script:Controls[$name].IsEnabled
                $script:Controls[$name].IsEnabled = $false
            }
        }
        $script:RequestSubmitContext.Clock.Restart()
        Show-GuiRequestSubmitDialog -JobName $JobName
    } else {
        $script:RequestSubmitContext.Clock.Stop()
        Close-GuiRequestSubmitDialog
        foreach ($name in $script:RequestSubmitUiState.Keys) { $script:Controls[$name].IsEnabled = $script:RequestSubmitUiState[$name] }
        $script:RequestSubmitUiState = @{}
        if ($script:Controls['LaunchJobsGrid'] -and $script:Controls['RequestJobButton']) {
            $launchRow = $script:Controls.LaunchJobsGrid.SelectedItem
            $script:Controls.RequestJobButton.IsEnabled = $null -ne $launchRow -and [bool]$launchRow.InManifest
        }
        Update-GuiPrepareRequestButton
        Update-SelectedPipelineRequest
        $script:Controls.CancelAllRequestsButton.IsEnabled = -not $script:CancellationBusy -and @($script:RequestRows | Where-Object Status -in @('Running', 'Queued', 'Cancelling')).Count -gt 0
    }
}

function Start-GuiJobRunSubmission {
    param([Parameter(Mandatory = $true)][string]$JobName, [bool]$IncludeDependencies)
    if ($script:RequestSubmitBusy) { return $false }
    $workerFolder = Split-Path $managementModulePath -Parent
    $context = $script:RequestSubmitContext
    $context.Input = [pscustomobject]@{
        JobName = $JobName; IncludeDependencies = $IncludeDependencies
        Root = $script:SharedDataFolderPath; Tenant = $Tenant; LogPath = $script:GuiLogPath
        ModulePaths = @($managementModulePath, (Join-Path $workerFolder 'SmartM365.Orchestrator.Pipeline.psm1'),
            (Join-Path $workerFolder 'SmartM365.Orchestrator.Insights.psm1'))
    }
    $context.Progress.Message = 'Starting...'
    $context.Progress.Stage = 'Starting'
    $context.Progress.LogError = ''
    Set-GuiRequestSubmitBusy $true -JobName $JobName
    $script:Controls.StatusText.Text = "Submitting run request for $JobName..."
    try {
        if (-not (Start-SmartM365OrchestratorGuiWorker -Context $context -Work (Get-GuiJobRunSubmissionWork) -InputObject $context.Input)) { throw 'A run request worker is already active.' }
    }
    catch {
        Set-GuiRequestSubmitBusy $false
        $script:Controls.StatusText.Text = 'Run request not submitted.'
        throw
    }
    return $true
}

function Receive-GuiJobRunSubmission {
    if (-not $script:RequestSubmitBusy) { return }
    $context = $script:RequestSubmitContext
    if ($null -ne $script:RequestSubmitDialog) {
        $script:RequestSubmitDialog.FindName('SubmitDetailText').Text = '{0} ({1:N0} s)' -f $context.Progress.Message, $context.Clock.Elapsed.TotalSeconds
    }
    $received = Receive-SmartM365OrchestratorGuiWorker -Context $context
    if ($null -eq $received) { return }
    $jobName = [string]$context.Input.JobName
    $stage = [string]$context.Progress.Stage
    $result = $null
    if (-not $received.Error -and $received.Output.Count -eq 1) { $result = $received.Output[0] }
    # Close the wait window and restore the controls before any result dialog.
    Set-GuiRequestSubmitBusy $false
    $title = 'Run request failed'; $icon = 'Error'
    if ($null -ne $result -and [string]$result.Outcome -eq 'Recorded') {
        if (-not $result.RefreshError) { Set-GuiRequestRows -Rows @($result.Rows) }
        Write-GuiActivity -Level SUCCESS -DisplayOnly -Message ("Run request recorded. Job={0}; IncludeDependencies={1}; BatchId={2}; Jobs={3}" -f $jobName, [bool]$context.Input.IncludeDependencies, $result.BatchId, $result.TotalCount)
        if ($result.QueuePosition -gt 0) {
            $script:Controls.StatusText.Text = "Run request queued: $($result.BatchId) (position $($result.QueuePosition))"
            $message = "Request queued.`nBatch: $($result.BatchId)`nJobs: $($result.TotalCount)`nQueue position: $($result.QueuePosition), behind $(@($result.QueuedBehind) -join ', ')`n`nThe request is recorded and starts automatically when the earlier requests have finished. Its dependency and prerequisite waits count from that start.`n`nFollow it in Operations > Requests."
        }
        else {
            $script:Controls.StatusText.Text = "Run request recorded: $($result.BatchId) (not started yet)"
            $message = "Request submitted.`nBatch: $($result.BatchId)`nJobs: $($result.TotalCount)`n`nThe request is recorded; no job has started yet. The orchestrators start it on its elected server when its dependency rule allows it.`n`nFollow it in Operations > Requests."
        }
        if ($result.RefreshError) { $message += "`n`nThe Requests list could not be refreshed ($($result.RefreshError)). Refresh it before relying on it." }
        $title = 'Run requested'; $icon = 'Information'
    }
    elseif ($null -ne $result -and [string]$result.Outcome -eq 'Failed' -and -not @($result.NewBatchIds).Count -and -not $result.CheckError) {
        $message = "The run request for '$jobName' was not recorded.`n`n$($result.Error)`n`nNo new request appeared in the shared folder. Fix the cause, then request again."
        Write-GuiActivity -Level ERROR -DisplayOnly -Message "Run request not recorded. Job=$jobName; Error=$($result.Error)"
        $script:Controls.StatusText.Text = 'Run request not recorded.'
    }
    elseif ($null -ne $result -and [string]$result.Outcome -eq 'Failed') {
        $seen = if (@($result.NewBatchIds).Count) { "New request(s) appeared meanwhile: $(@($result.NewBatchIds) -join ', ')." } else { "The shared folder could not be checked: $($result.CheckError)" }
        $message = "The run request for '$jobName' reported an error and may still have been recorded.`n`n$($result.Error)`n`n$seen`n`nCheck Operations > Requests (Refresh) before requesting again. Nothing is resubmitted automatically."
        Write-GuiActivity -Level WARN -DisplayOnly -Message "Run request result uncertain. Job=$jobName; Error=$($result.Error); $seen"
        $script:Controls.StatusText.Text = 'Run request result uncertain: check Operations > Requests.'
        $title = 'Run request uncertain'; $icon = 'Warning'
    }
    elseif ($stage -in @('Submitting', 'Recorded')) {
        $detail = if ($received.Error) { $received.Error } else { 'The submission worker returned an invalid result.' }
        $message = "The run request for '$jobName' was interrupted while it was being recorded and may have been recorded.`n`n$detail`n`nCheck Operations > Requests (Refresh) before requesting again. Nothing is resubmitted automatically."
        Write-GuiActivity -Level WARN -DisplayOnly -Message "Run request result uncertain. Job=$jobName; Stage=$stage; Error=$detail"
        $script:Controls.StatusText.Text = 'Run request result uncertain: check Operations > Requests.'
        $title = 'Run request uncertain'; $icon = 'Warning'
    }
    else {
        $detail = if ($received.Error) { $received.Error } else { 'The submission worker returned an invalid result.' }
        $message = "The run request for '$jobName' was not submitted.`n`n$detail"
        Write-GuiActivity -Level ERROR -DisplayOnly -Message "Run request not submitted. Job=$jobName; Error=$detail"
        $script:Controls.StatusText.Text = 'Run request not submitted.'
    }
    if ($context.Progress.LogError) { Write-GuiActivity -Level WARN -DisplayOnly -Message "Persistent run request log unavailable: $($context.Progress.LogError)" }
    # Live refresh first (read-only, background): the grids update while the result is read.
    Invoke-AutoRefresh
    Show-GuiJobRunRequestResult -Message $message -Title $title -Icon $icon
}

function Get-GuiLaunchCatalog {
    # Catalog of the committed template; the published jobs give the Arguments and Execution.Mode that really run.
    $catalog = Get-SmartM365InventoryCatalog -SmartInventoryRoot $script:SmartInventoryRoot
    $published = @{}
    if ($null -ne $script:Snapshot) { foreach ($job in @($script:Snapshot.Jobs.Jobs)) { $published[[string]$job.Name] = $job } }
    foreach ($entry in @($catalog.Entries)) {
        if ($entry.Kind -eq 'Job' -and $published.ContainsKey($entry.Name)) { $entry.Overlay = $published[$entry.Name] }
    }
    return $catalog
}

function Refresh-LaunchView {
    try {
        $catalog = Get-GuiLaunchCatalog
        $script:LaunchCatalog = $catalog
        $globalMode = [string](Get-SmartM365InventoryGlobalSetting -SmartM365Root $catalog.SmartM365Root -Name 'OnPremExecutionMode')
        if ([string]::IsNullOrWhiteSpace($globalMode)) { $globalMode = 'LocalCache' }
        $script:LaunchGlobalMode = $globalMode
        $script:Controls.OnPremModeText.Text = $globalMode
        $draft = @{}
        foreach ($job in @($script:DraftJobs.Jobs)) { $draft[[string]$job.Name] = $job }
        $jobRows = foreach ($entry in @($catalog.Entries | Where-Object { $_.Kind -eq 'Job' })) {
            $definition = $entry.Definition
            $execution = Get-OrchestratorGuiPropertyValue -Object $definition -Name 'Execution' -DefaultValue $null
            $runtime = [string](Get-OrchestratorGuiPropertyValue -Object $definition -Name 'Runtime' -DefaultValue (Get-OrchestratorGuiPropertyValue -Object $definition -Name 'PowerShellEdition' -DefaultValue 'PowerShell7'))
            $mode = Resolve-SmartM365InventoryExecutionMode -Entry $entry -Catalog $catalog -DefaultOnPremExecutionMode $globalMode
            $draftJob = if ($draft.ContainsKey($entry.Name)) { $draft[$entry.Name] } else { $null }
            $draftMode = [string](Get-OrchestratorGuiPropertyValue -Object (Get-OrchestratorGuiPropertyValue -Object $draftJob -Name 'Execution' -DefaultValue $null) -Name 'Mode' -DefaultValue '')
            $health = if ($script:HealthByName.ContainsKey([string]$entry.Name)) { $script:HealthByName[[string]$entry.Name] } else { $null }
            $swi = Get-GuiSwiRow -JobName ([string]$entry.Name)
            [pscustomobject]@{
                Name = $entry.Name
                Version = Get-GuiJobScriptVersion -Job $definition
                Group = [string](Get-OrchestratorGuiPropertyValue -Object $draftJob -Name 'Group' -DefaultValue (Get-OrchestratorGuiPropertyValue -Object $definition -Name 'Group' -DefaultValue ''))
                Enabled = [bool](Get-OrchestratorGuiPropertyValue -Object $draftJob -Name 'Enabled' -DefaultValue $false)
                Runtime = $runtime
                OnPremises = $null -ne $execution
                Mode = $mode.Mode
                ModeSource = if ($mode.Zone) { '{0}, zone {1}' -f $mode.Source, $mode.Zone } else { $mode.Source }
                DraftMode = if ($null -eq $execution) { '' } elseif ($draftMode) { $draftMode } else { '(default)' }
                CacheName = [string](Get-OrchestratorGuiPropertyValue -Object $execution -Name 'CacheName' -DefaultValue '')
                InManifest = $null -ne $draftJob
                LastSuccess = if ($health) { $health.LastSuccess } else { '' }
                Swi = $swi.Swi; SwiDaysLeft = $swi.DaysLeft; SwiDetail = $swi.Detail; SwiBlocks = [bool]$swi.Blocks
            }
        }
        # HideInGui entries (orchestrator actions already offered as buttons, setup tools) stay in SmartInventory.cmd only.
        $manualRows = foreach ($entry in @($catalog.Entries | Where-Object { $_.Kind -eq 'Manual' -and -not [bool](Get-OrchestratorGuiPropertyValue -Object $_.Definition -Name 'HideInGui' -DefaultValue $false) })) {
            $definition = $entry.Definition
            $mode = Resolve-SmartM365InventoryExecutionMode -Entry $entry -Catalog $catalog -DefaultOnPremExecutionMode $globalMode
            [pscustomobject]@{
                Name = $entry.Name
                Runtime = [string](Get-OrchestratorGuiPropertyValue -Object $definition -Name 'Runtime' -DefaultValue 'PowerShell7')
                Mode = $mode.Mode
                Interactive = [bool](Get-OrchestratorGuiPropertyValue -Object $definition -Name 'Interactive' -DefaultValue $false)
                RequiresConfirmation = [bool](Get-OrchestratorGuiPropertyValue -Object $definition -Name 'RequiresConfirmation' -DefaultValue $false)
                Description = [string](Get-OrchestratorGuiPropertyValue -Object $definition -Name 'Description' -DefaultValue '')
            }
        }
        $script:LaunchJobRows = @($jobRows | Sort-Object -Property Name)
        $script:ManualLaunchRows = @($manualRows | Sort-Object -Property Name)
        $script:Controls.LaunchJobsGrid.ItemsSource = $script:LaunchJobRows
        $script:Controls.ManualLaunchesGrid.ItemsSource = $script:ManualLaunchRows
        Update-SelectedLaunchPreview -Kind Job
        Update-SelectedLaunchPreview -Kind Manual
        Update-GuiPrepareRequestButton
    }
    catch { Write-GuiException -Context 'Launch catalog refresh failed' -ErrorRecord $_ }
}

function Get-GuiLaunchPlanText {
    param([Parameter(Mandatory = $true)][string]$Name)
    try {
        $plan = Resolve-SmartM365InventoryLaunch -Catalog $script:LaunchCatalog -Name $Name -Tenant $Tenant -DefaultOnPremExecutionMode $script:LaunchGlobalMode
        return [pscustomobject]@{ Plan = $plan; Text = (Format-SmartM365InventoryPlan -Plan $plan -Catalog $script:LaunchCatalog) }
    }
    catch { return [pscustomobject]@{ Plan = $null; Text = ('Command not resolved: ' + $_.Exception.Message) } }
}

function Update-SelectedLaunchPreview {
    param([Parameter(Mandatory = $true)][ValidateSet('Job', 'Manual')][string]$Kind)
    if ($Kind -eq 'Job') {
        $row = $script:Controls.LaunchJobsGrid.SelectedItem
        $script:Controls.RunJobNowButton.IsEnabled = $null -ne $row
        $script:Controls.RunJobAsServiceAccountButton.IsEnabled = $null -ne $row
        $script:Controls.RequestJobButton.IsEnabled = $null -ne $row -and [bool]$row.InManifest -and -not $script:RequestSubmitBusy
        $script:Controls.JobExecutionModeCombo.IsEnabled = $null -ne $row -and [bool]$row.OnPremises -and [bool]$row.InManifest
        $script:Controls.ApplyJobExecutionModeButton.IsEnabled = $script:Controls.JobExecutionModeCombo.IsEnabled
        if ($null -eq $row) { $script:Controls.LaunchJobCommandBox.Text = ''; return }
        if ($row.OnPremises) { Select-ComboText -Combo $script:Controls.JobExecutionModeCombo -Text ([string]$row.DraftMode) }
        $script:Controls.LaunchJobCommandBox.Text = (Get-GuiLaunchPlanText -Name ([string]$row.Name)).Text
        return
    }
    $row = $script:Controls.ManualLaunchesGrid.SelectedItem
    $script:Controls.RunManualLaunchButton.IsEnabled = $null -ne $row
    $script:Controls.ManualLaunchCommandBox.Text = if ($null -eq $row) { '' } else { (Get-GuiLaunchPlanText -Name ([string]$row.Name)).Text }
}

function Start-GuiConsoleProcess {
    # Starts a script in its own console window; the GUI does not wait for it.
    param(
        [Parameter(Mandatory = $true)][ValidateSet('PowerShell7', 'WindowsPowerShell')][string]$Runtime,
        [Parameter(Mandatory = $true)][string]$ScriptPath,
        [string[]]$Arguments = @(),
        [switch]$KeepOpen,
        [ValidateSet('Normal', 'Hidden', 'Minimized')][string]$WindowStyle = 'Normal'
    )
    $engine = Get-SmartM365InventoryEnginePath -Runtime $Runtime
    $all = @('-NoProfile', '-ExecutionPolicy', 'Bypass')
    if ($KeepOpen) { $all += '-NoExit' }
    $all += @('-File', $ScriptPath) + @($Arguments)
    $commandLine = (@($all) | ForEach-Object { ConvertTo-SmartM365InventoryCommandLineArgument -Value $_ }) -join ' '
    $process = Start-Process -FilePath $engine -ArgumentList $commandLine -WindowStyle $WindowStyle -PassThru
    Write-GuiActivity -Message ("Console started on {0} (PID {1}): {2} {3}" -f $env:COMPUTERNAME, $process.Id, $engine, $commandLine)
    return $process
}

function Get-GuiOrchestratorActionCommand {
    # Same commands as the former Launchers\Orchestrator .cmd files, with the tenant of this GUI.
    param([Parameter(Mandatory = $true)][ValidateSet('Start', 'Stop', 'Restart', 'FullCollect', 'SendSummary')][string]$Action)
    $folder = Join-Path $script:SmartInventoryRoot 'Orchestrator'
    $orchestrator = Join-Path $folder 'SmartM365-Inventory-Orchestrator.ps1'
    switch ($Action) {
        'Start' { [pscustomobject]@{ Title = 'Start the orchestrator'; ScriptPath = $orchestrator; Arguments = @('-Tenant', $Tenant, '-Connect'); Icon = 'Question'; KeepOpen = $false
                Message = "Start a resident orchestrator in a new console on this computer ($env:COMPUTERNAME)?`n`nUse Restart to run it through its scheduled task instead." } }
        'Stop' { [pscustomobject]@{ Title = 'Stop the orchestrator'; ScriptPath = $orchestrator; Arguments = @('-Tenant', $Tenant, '-Stop', '-StopTimeoutSeconds', '180'); Icon = 'Warning'; KeepOpen = $true
                Message = "Ask the resident orchestrator of this computer ($env:COMPUTERNAME) to stop cleanly?`n`nIt saves its state and releases its lock; it stops waiting for jobs after 180 seconds." } }
        'Restart' { [pscustomobject]@{ Title = 'Restart the orchestrator'; ScriptPath = (Join-Path $folder 'Restart-SmartM365-Inventory-Orchestrator.ps1'); Arguments = @('-Tenant', $Tenant); Icon = 'Warning'; KeepOpen = $true
                Message = "Restart the orchestrator of this computer ($env:COMPUTERNAME)?`n`nIt is stopped cleanly, then its scheduled task is started." } }
        'FullCollect' { [pscustomobject]@{ Title = 'Full collect'; ScriptPath = (Join-Path $folder 'SmartM365-Inventory-Pipeline.ps1'); Arguments = @('-Tenant', $Tenant, '-Pipeline', 'Full', '-Collect'); Icon = 'Warning'; KeepOpen = $true
                Message = "Request the full pipeline from the resident orchestrators?`n`nEvery enabled job runs on its elected server with the shared locks and dependency rules. If another request is active, this one is queued behind it; follow it in the new console or Operations > Requests." } }
        'SendSummary' { [pscustomobject]@{ Title = 'Send summary'; ScriptPath = $orchestrator; Arguments = @('-Tenant', $Tenant, '-SendExecutionSummary'); Icon = 'Question'; KeepOpen = $true
                Message = 'Send the execution summary email now? The resident orchestrator and its jobs are not affected.' } }
    }
}

function Start-GuiSendSummary {
    # Hidden console: a wait window replaces it until the process ends (Receive-GuiSendSummary).
    param([Parameter(Mandatory = $true)]$Command)
    $process = Start-GuiConsoleProcess -Runtime PowerShell7 -ScriptPath $Command.ScriptPath -Arguments $Command.Arguments -WindowStyle Hidden
    $null = $process.Handle  # keeps the exit code readable after the process ends
    $script:SendSummaryRun = [pscustomobject]@{ Process = $process; Clock = [Diagnostics.Stopwatch]::StartNew(); Dialog = $null }
    $script:Controls.SendSummaryButton.IsEnabled = $false
    $script:SendSummaryRun.Dialog = New-GuiWaitDialog -Title 'Send summary' -Message ('Sending the execution summary email{0} Please wait.' -f [char]0x2026) -Detail 'Reading the job runs of this tenant. The resident orchestrator and its jobs are not affected.'
    $script:Controls.StatusText.Text = 'Send summary: sending in the background.'
}

function Receive-GuiSendSummary {
    $run = $script:SendSummaryRun
    if ($null -eq $run) { return }
    $process = $run.Process
    $timedOut = $run.Clock.Elapsed.TotalMinutes -ge $script:SendSummaryTimeoutMinutes
    if (-not $process.HasExited -and -not $timedOut) {
        if ($null -ne $run.Dialog) { $run.Dialog.FindName('SubmitDetailText').Text = 'Reading the job runs and sending the email ({0:N0} s). The resident orchestrator and its jobs are not affected.' -f $run.Clock.Elapsed.TotalSeconds }
        return
    }
    $script:SendSummaryRun = $null
    Close-GuiWaitDialog -Dialog $run.Dialog
    $script:Controls.SendSummaryButton.IsEnabled = $true
    if (-not $process.HasExited) {
        Write-GuiActivity -Message ("Send summary still running after {0} minutes (PID {1}); the GUI stops waiting." -f $script:SendSummaryTimeoutMinutes, $process.Id) -Level WARN
        $script:Controls.StatusText.Text = 'Send summary: still running in the background.'
        Show-GuiJobRunRequestResult -Message ("The execution summary is still being sent after {0} minutes (PID {1}).`n`nIt continues in the background; check the orchestrator log of this computer." -f $script:SendSummaryTimeoutMinutes, $process.Id) -Title 'Send summary' -Icon Warning
        return
    }
    $exitCode = $process.ExitCode
    if ($exitCode -eq 0) {
        Write-GuiActivity -Message ("Execution summary email sent ({0:N0} s)." -f $run.Clock.Elapsed.TotalSeconds) -Level SUCCESS
        $script:Controls.StatusText.Text = 'Send summary: email sent.'
        Show-GuiJobRunRequestResult -Message 'The execution summary email was sent.' -Title 'Send summary' -Icon Information
    } else {
        Write-GuiActivity -Message ("Execution summary failed (exit code {0})." -f $exitCode) -Level ERROR
        $script:Controls.StatusText.Text = 'Send summary: failed.'
        Show-GuiJobRunRequestResult -Message ("The execution summary email could not be sent (exit code {0}).`n`nCheck the mail settings and the orchestrator log of this computer." -f $exitCode) -Title 'Send summary' -Icon Error
    }
}

function Invoke-GuiOrchestratorAction {
    param([Parameter(Mandatory = $true)][string]$Action)
    $command = Get-GuiOrchestratorActionCommand -Action $Action
    $preview = (@($command.ScriptPath) + @($command.Arguments) | ForEach-Object { ConvertTo-SmartM365InventoryCommandLineArgument -Value $_ }) -join ' '
    if ($Action -eq 'SendSummary' -and $null -ne $script:SendSummaryRun) { throw 'The execution summary is already being sent. Wait for its result.' }
    if ($script:OrchestratorDialog.Show(("{0}`n`nCommand (PowerShell 7):`n{1}" -f $command.Message, $preview), $command.Title, 'YesNo', $command.Icon) -ne 'Yes') { return }
    if ($Action -eq 'SendSummary') { Start-GuiSendSummary -Command $command; return }
    [void](Start-GuiConsoleProcess -Runtime PowerShell7 -ScriptPath $command.ScriptPath -Arguments $command.Arguments -KeepOpen:([bool]$command.KeepOpen))
    $script:Controls.StatusText.Text = "$($command.Title): started in a new console."
}

$script:ServerAcceptanceLaunchName = 'Orchestrator-ServerAcceptance'

function Update-ServerTestButtons {
    # The acceptance test reads from the computer it runs on: it can only test the computer of this GUI.
    param([AllowEmptyString()][string]$Server)
    $local = -not [string]::IsNullOrWhiteSpace($Server) -and [string]::Equals($Server, $env:COMPUTERNAME, [StringComparison]::OrdinalIgnoreCase)
    $script:Controls.TestServerButton.IsEnabled = $local
    $script:Controls.TestServerAsServiceAccountButton.IsEnabled = $local
    $script:Controls.ServerTestHintText.Text = if ([string]::IsNullOrWhiteSpace($Server)) { 'Select a server' }
    elseif ($local) { 'Runs on this computer; also in Manual launches (Orchestrator-ServerAcceptance).' }
    else { "Open this GUI on $Server to test it: the test runs on the computer that starts it." }
}

function Get-GuiOrchestratorServiceAccount {
    # Account of this computer's orchestrator scheduled task (installer name: 'SmartM365 Inventory Orchestrator - <Tenant>').
    # Never hard-coded: the repository is public.
    param([AllowNull()][object[]]$Tasks = $null)
    if ($null -eq $Tasks) { $Tasks = @(Get-ScheduledTask -TaskName 'SmartM365 Inventory Orchestrator*' -ErrorAction SilentlyContinue) }
    $tasks = @($Tasks | Where-Object { $null -ne $_ })
    if ($tasks.Count -eq 0) { throw "No 'SmartM365 Inventory Orchestrator' scheduled task on this computer ($env:COMPUTERNAME): the service account cannot be determined." }
    $preferred = @($tasks | Where-Object { [string]$_.TaskName -eq ('SmartM365 Inventory Orchestrator - ' + $Tenant) })
    $task = if ($preferred.Count) { $preferred[0] } else { $tasks[0] }
    $account = [string]$task.Principal.UserId
    if ([string]::IsNullOrWhiteSpace($account)) { throw "Scheduled task '$($task.TaskName)' has no account." }
    if ($account.EndsWith('$')) { throw "Scheduled task '$($task.TaskName)' runs as a managed account ($account): it has no password, runas cannot use it." }
    if ($account -notmatch '[\\@]') { $account = '{0}\{1}' -f $env:USERDOMAIN, $account }
    return [pscustomobject]@{ Account = $account; TaskName = [string]$task.TaskName }
}

function Get-GuiRunAsArgumentList {
    # runas.exe /user:<account> "<engine> <arguments>": the inner command is quoted with the same
    # CommandLineToArgvW rules (inner quotes become \").
    param([Parameter(Mandatory = $true)][string]$Account, [Parameter(Mandatory = $true)][string]$Engine, [string[]]$Arguments = @())
    $inner = (@($Engine) + @($Arguments) | ForEach-Object { ConvertTo-SmartM365InventoryCommandLineArgument -Value $_ }) -join ' '
    return '/user:{0} {1}' -f (ConvertTo-SmartM365InventoryCommandLineArgument -Value $Account), (ConvertTo-SmartM365InventoryCommandLineArgument -Value $inner)
}

function Invoke-GuiLaunch {
    # Runs a catalog job or a manual entry now on this computer, like SmartInventory.cmd <Name>.
    # -AsServiceAccount: same command started with runas under the orchestrator task account.
    # -Name runs that catalog entry instead of the grid selection (server acceptance test buttons).
    param([Parameter(Mandatory = $true)][ValidateSet('Job', 'Manual')][string]$Kind, [switch]$AsServiceAccount, [string]$Name = '')
    $row = if ($Name) { [pscustomobject]@{ Name = $Name } } elseif ($Kind -eq 'Job') { $script:Controls.LaunchJobsGrid.SelectedItem } else { $script:Controls.ManualLaunchesGrid.SelectedItem }
    if ($null -eq $row) { throw 'Select an entry first.' }
    $preview = Get-GuiLaunchPlanText -Name ([string]$row.Name)
    if ($null -eq $preview.Plan) { throw $preview.Text }
    $icon = 'Question'
    $message = "Run '{0}' now on this computer ({1}), outside the orchestrators?`n`n{2}" -f $row.Name, $env:COMPUTERNAME, $preview.Text
    if ($Kind -eq 'Job') { $message += "`n`nNo election, shared lock or dependency check. Use Request to run it through the orchestrators." }
    $serviceAccount = $null
    if ($AsServiceAccount) {
        $serviceAccount = Get-GuiOrchestratorServiceAccount
        $message += "`n`nAccount: {0} (scheduled task '{1}'). A runas console asks for its password; the GUI never sees it." -f $serviceAccount.Account, $serviceAccount.TaskName
    }
    if ($preview.Plan.Interactive) { $message += "`n`nThis entry needs you in the console (interactive sign-in or prompts)." }
    if ($preview.Plan.RequiresConfirmation) {
        $icon = 'Warning'
        $message += "`n`nTHIS ENTRY MODIFIES DATA. The console will also ask you to type WRITE before it starts."
    }
    if ($script:OrchestratorDialog.Show($message, ('Run ' + $row.Name), 'YesNo', $icon) -ne 'Yes') { return }
    $engineScript = Join-Path $script:SmartInventoryRoot 'Orchestrator\Invoke-SmartM365InventoryJob.ps1'
    $engineArguments = @([string]$row.Name, '-Now', '-Tenant', $Tenant)
    if ($null -ne $serviceAccount) {
        $engine = Get-SmartM365InventoryEnginePath -Runtime WindowsPowerShell
        $runAsArguments = Get-GuiRunAsArgumentList -Account $serviceAccount.Account -Engine $engine -Arguments (@('-NoProfile', '-ExecutionPolicy', 'Bypass', '-NoExit', '-File', $engineScript) + $engineArguments)
        $process = Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\runas.exe') -ArgumentList $runAsArguments -PassThru
        Write-GuiActivity -Message ("Run now as {0}: runas console started (PID {1}): runas.exe {2}" -f $serviceAccount.Account, $process.Id, $runAsArguments)
    }
    else { [void](Start-GuiConsoleProcess -Runtime WindowsPowerShell -ScriptPath $engineScript -Arguments $engineArguments -KeepOpen) }
    Write-GuiActivity -Message ("Run now: {0} ({1}) started on {2}." -f $row.Name, $Kind, $env:COMPUTERNAME) -Level SUCCESS
    $script:Controls.StatusText.Text = "$($row.Name): started in a new console."
}

function Set-GuiJobExecutionMode {
    # Writes Execution.Mode of on-premises jobs in the draft; (default) removes it so OnPremExecutionMode applies.
    param([Parameter(Mandatory = $true)][string[]]$JobName, [Parameter(Mandatory = $true)][ValidateSet('(default)', 'Direct', 'LocalCache', 'Auto')][string]$Mode)
    $changed = @()
    foreach ($name in $JobName) {
        $job = @($script:DraftJobs.Jobs | Where-Object { [string]$_.Name -eq $name })[0]
        if ($null -eq $job) { throw "Job '$name' is not in the published jobs configuration." }
        $execution = Get-OrchestratorGuiPropertyValue -Object $job -Name 'Execution' -DefaultValue $null
        if ($Mode -eq '(default)') {
            if ($null -eq $execution) { continue }
            if ($execution.PSObject.Properties['Mode']) { $execution.PSObject.Properties.Remove('Mode') }
            if (@($execution.PSObject.Properties).Count -eq 0) { $job.PSObject.Properties.Remove('Execution') }
        }
        elseif ($null -eq $execution) { Set-JsonProperty -Object $job -Name 'Execution' -Value ([pscustomobject]@{ Mode = $Mode }) }
        else { Set-JsonProperty -Object $execution -Name 'Mode' -Value $Mode }
        $changed += $name
    }
    if ($changed.Count) {
        Write-GuiActivity -Message ("Draft execution mode set to {0} for: {1}. Publish the changes to apply." -f $Mode, ($changed -join ', ')) -Level WARN
        $script:Controls.StatusText.Text = "Execution mode changed in the draft ($($changed.Count) job(s)); publish to apply."
    }
    Refresh-LaunchView
    return $changed
}

function Show-FailuresLast24Hours {
    $from = (Get-Date).AddHours(-24)
    $script:Controls.HistoryFromPicker.SelectedDate = $from.Date
    $script:Controls.HistoryToPicker.SelectedDate = (Get-Date).Date
    $script:Controls.HistoryServerCombo.SelectedIndex = 0
    $script:Controls.HistoryJobCombo.SelectedIndex = 0
    $script:Controls.HistoryStatusCombo.SelectedIndex = 0
    Start-GuiRefresh -Mode History -HistoryFrom $from -HistoryTo (Get-Date) -OnlyFailures
}

function Set-GuiEventsServers {
    $selected = Get-ComboText -Combo $script:Controls.EventsServerCombo
    $script:Controls.EventsServerCombo.ItemsSource = @('All') + @($script:DraftCluster.ExpectedOrchestratorServers | ForEach-Object { [string]$_ } | Sort-Object -Unique)
    if ($selected -and $selected -in @($script:Controls.EventsServerCombo.ItemsSource)) { $script:Controls.EventsServerCombo.SelectedItem = $selected } else { $script:Controls.EventsServerCombo.SelectedIndex = 0 }
}

function Test-GuiConfigurationTabVisible {
    # The configuration files are read only while their Settings sub-tab is shown.
    if ($script:Controls['ConfigurationSettingsTab'] -and -not $script:Controls.ConfigurationSettingsTab.IsSelected) { return $false }
    return $null -ne $script:Controls.MainTabs.SelectedItem -and [string]$script:Controls.MainTabs.SelectedItem.Header -eq 'Configuration'
}

function Update-GuiConfigurationView {
    # UI-thread display of the last read: masking and filtering only, no file access, nothing logged.
    $view = $script:ConfigurationView
    $show = [datetime]::UtcNow -lt $script:ConfigurationShowUntil
    $script:Controls.ConfigShowSensitiveButton.Content = if ($show) { 'Hide ({0:0} s)' -f ($script:ConfigurationShowUntil - [datetime]::UtcNow).TotalSeconds } else { 'Show' }
    if ($null -eq $view) { return }
    $filter = $script:Controls.ConfigFilterBox.Text.Trim()
    foreach ($pair in @(@('Global', 'ConfigGlobal'), @('Tenant', 'ConfigTenant'), @('Mail', 'ConfigMail'))) {
        $file = @($view.Files | Where-Object Kind -eq $pair[0])[0]
        $grid = $script:Controls[$pair[1] + 'Grid']; $status = $script:Controls[$pair[1] + 'StatusText']; $tab = $script:Controls[$pair[1] + 'Tab']
        if ($null -eq $file) { $grid.ItemsSource = @(); continue }
        $rows = @(Select-SmartM365OrchestratorConfigRows -Rows @($file.Rows) -Filter $filter -ShowSensitive:$show -HideNotInTemplate)
        $grid.ItemsSource = $rows
        $tab.Header = '{0} ({1})' -f $pair[0], [IO.Path]::GetFileName([string]$file.Path)
        $statusText = switch ($file.Status) { 'OK' { 'Read' } 'Missing' { 'File not found' } 'Busy' { 'Transport lock held by another process; previous values not shown' } default { 'Unreadable' } }
        $status.Text = '{0}: {1}. {2} key(s), {3} shown; {4} not in template (hidden), {5} missing from file; {6} sensitive value(s) {7}.{8}' -f $statusText, $file.Path, @($file.Rows).Count, $rows.Count, $file.NotInTemplate, $file.MissingFromFile, @($file.Rows | Where-Object Sensitive).Count, $(if ($show) { 'shown' } else { 'hidden' }), $(if ($file.Error) { ' ' + $file.Error } else { '' })
    }
    $script:Controls.ConfigSummaryText.Text = 'Tenant profile {0}, read {1:HH:mm:ss}. Read-only: nothing is written; sensitive values are hidden unless Show is pressed.' -f $view.ProfileKey, $view.ReadAt
}

function Set-GuiConfiguration {
    param([AllowNull()]$View, [string]$ErrorMessage = '')
    if ($ErrorMessage) {
        $script:Controls.ConfigSummaryText.Text = 'Configuration could not be read: ' + $ErrorMessage
        Write-GuiActivity -Message 'Configuration view could not be read; see the Configuration tab.' -Level WARN -DisplayOnly
        return
    }
    if ($null -eq $View) { return }
    $script:ConfigurationView = $View
    Update-GuiConfigurationView
}

function Show-GuiConfigurationSensitive {
    if ([datetime]::UtcNow -lt $script:ConfigurationShowUntil) { $script:ConfigurationShowUntil = [datetime]::MinValue }
    else {
        $script:ConfigurationShowUntil = [datetime]::UtcNow.AddSeconds(30)
        Write-GuiActivity -Message 'Sensitive configuration values shown for 30 seconds (values are not logged).' -Level WARN
    }
    Update-GuiConfigurationView
}

function Open-GuiConfigurationFolder {
    $folder = if ($null -ne $script:ConfigurationView) { $script:ConfigurationView.ConfigFolder } else { Join-Path $script:SmartM365RootPath 'Config' }
    if ([string]$script:Controls.ConfigFileTabs.SelectedItem.Name -ne 'ConfigGlobalTab') { $folder = Join-Path $folder 'Tenants' }
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) { throw "Configuration folder not found: $folder" }
    Start-Process -FilePath 'explorer.exe' -ArgumentList ('"{0}"' -f $folder)
}

function Test-GuiDashboardTabVisible {
    return $null -ne $script:Controls.MainTabs.SelectedItem -and [string]$script:Controls.MainTabs.SelectedItem.Header -eq 'Dashboard'
}

function Set-GuiDashboard {
    param([AllowNull()]$Dashboard, [string]$ErrorMessage = '')
    if ($ErrorMessage) {
        if ($ErrorMessage -ne $script:DashboardLastError) { Write-GuiActivity -Message "Dashboard could not be built: $ErrorMessage" -Level WARN }
        $script:DashboardLastError = $ErrorMessage
        $script:Controls.DashboardBannerBorder.Background = ConvertTo-OrchestratorBrush $script:DashboardBannerColors['Critical']
        $script:Controls.DashboardBannerText.Text = 'Dashboard unavailable (previous values kept): ' + $ErrorMessage
        return
    }
    $script:DashboardLastError = ''
    if ($null -eq $Dashboard) { return }
    $script:DashboardData = $Dashboard
    Set-OrchestratorDashboardView -Controls $script:Controls -Dashboard $Dashboard
}

function Update-GuiDashboardLoading {
    # The overlay appears only when a requested dashboard read lasts more than 350 ms.
    $show = $script:DashboardLoading.Pending -and $script:DashboardLoading.Clock.ElapsedMilliseconds -ge 350 -and (Test-GuiDashboardTabVisible)
    $script:Controls.DashboardLoadingOverlay.Visibility = if ($show) { 'Visible' } else { 'Collapsed' }
    if ($show) {
        $message = [string]$script:RefreshContext.Progress.Message
        $script:Controls.DashboardLoadingText.Text = '{0}{1}{2:0} s elapsed. The window stays usable.' -f $(if ($message) { $message } else { 'Reading shared orchestrator data in the background.' }), [Environment]::NewLine, $script:DashboardLoading.Clock.Elapsed.TotalSeconds
    }
}

function Stop-GuiDashboardLoading {
    if ($null -ne $script:PendingRefreshRequest -and $script:PendingRefreshRequest.Mode -ne 'History' -and ([bool]$script:PendingRefreshRequest.ShowLoading -or $script:PendingRefreshRequest.Mode -eq 'Full')) { return }
    $script:DashboardLoading.Pending = $false
    $script:DashboardLoading.Clock.Reset()
    $script:Controls.DashboardLoadingOverlay.Visibility = 'Collapsed'
}

function Invoke-GuiDashboardTarget {
    param([string]$Target)
    if ([string]::IsNullOrWhiteSpace($Target)) { return }
    switch ($Target) {
        'Failures24h' {
            Select-GuiTab -Header 'History'
            Show-FailuresLast24Hours
        }
        'Events' {
            $script:Controls.EventsLevelCombo.SelectedIndex = 1
            $script:Controls.EventsPeriodCombo.SelectedIndex = 2
            Select-GuiTab -Header 'Events'
        }
        default { Select-GuiTab -Header $Target }
    }
}

function Select-GuiTab {
    param([Parameter(Mandatory = $true)][string]$Header)
    # Requests is a sub-tab of Operations since GUI 1.6.3.
    if ($Header -eq 'Requests') { $script:Controls.OperationsRequestsTab.IsSelected = $true; $Header = 'Operations' }
    # Servers and versions is a sub-tab of Configuration since GUI 1.6.13.
    if ($Header -eq 'Servers and versions') { $script:Controls.ConfigurationServersTab.IsSelected = $true; $Header = 'Configuration' }
    foreach ($tab in $script:Controls.MainTabs.Items) { if ([string]$tab.Header -eq $Header) { $script:Controls.MainTabs.SelectedItem = $tab; return } }
}

function Test-GuiEventsTabVisible {
    return $null -ne $script:Controls.MainTabs.SelectedItem -and [string]$script:Controls.MainTabs.SelectedItem.Header -eq 'Events'
}

function Get-GuiEventsHours {
    $item = $script:Controls.EventsPeriodCombo.SelectedItem
    if ($null -eq $item) { return 24 }
    return [int]$item.Hours
}

function Update-GuiEventsView {
    # UI-thread filtering of the last read; no file access.
    $result = $script:LogEventsResult
    if ($null -eq $result) { $script:Controls.EventsGrid.ItemsSource = @(); return }
    $server = Get-ComboText -Combo $script:Controls.EventsServerCombo
    $category = Get-ComboText -Combo $script:Controls.EventsCategoryCombo
    $levelItem = $script:Controls.EventsLevelCombo.SelectedItem
    $selection = Select-SmartM365OrchestratorLogEvents -Events @($result.Events) -Server $(if ($server -eq 'All') { '' } else { $server }) `
        -Level $(if ($null -ne $levelItem) { [string]$levelItem.Value } else { 'All' }) -Category $(if ($category -eq 'All') { '' } else { $category }) `
        -Text $script:Controls.EventsTextBox.Text.Trim() -Since $result.Since -GroupRepeats:([bool]$script:Controls.EventsGroupCheck.IsChecked)
    $script:Controls.EventsGrid.ItemsSource = @($selection.Rows)
    $sources = @($result.Sources | ForEach-Object { '{0}: {1}' -f $_.Server, $_.Status }) -join '; '
    $shown = if ($selection.Truncated) { '{0} of {1} rows shown (newest)' -f @($selection.Rows).Count, $selection.Total } else { '{0} row(s)' -f $selection.Total }
    $script:Controls.EventsSummaryText.Text = '{0} for {1} event(s) since {2:yyyy-MM-dd HH:mm}, read {3:HH:mm:ss}. Logs: {4}' -f $shown, $selection.Matching, $result.Since, $result.ReadAt, $sources
    $script:Controls.EventsDetailBox.Text = ''
}

function Set-GuiLogEvents {
    param([AllowNull()]$Result, [string]$ErrorMessage = '')
    if ($ErrorMessage) {
        if ($ErrorMessage -ne $script:LogEventsLastError) { Write-GuiActivity -Message "Resident log events could not be read: $ErrorMessage" -Level WARN }
        $script:LogEventsLastError = $ErrorMessage
        $script:Controls.EventsSummaryText.Text = 'Resident logs unreadable (previous rows kept): ' + $ErrorMessage
        return
    }
    $script:LogEventsLastError = ''
    if ($null -eq $Result) { return }
    $script:LogEventsResult = $Result
    Update-GuiEventsView
}

function Show-SelectedGuiEvent {
    $row = $script:Controls.EventsGrid.SelectedItem
    if ($null -eq $row) { $script:Controls.EventsDetailBox.Text = ''; return }
    $repeat = if ($row.Count -gt 1) { '; {0} occurrence(s) since {1}' -f $row.Count, $row.First } else { '' }
    $script:Controls.EventsDetailBox.Text = '[{0}][{1}][{2}] {3}{4}{5}{6}{5}Log: {7}' -f $row.Time, $row.Server, $row.Level, $row.Category, $repeat, [Environment]::NewLine, $row.Message, $row.Source
}

function Get-GuiRefreshWork {
    return {
        param($Progress, $InputData)
        $ErrorActionPreference = 'Stop'
        Set-StrictMode -Version 2.0
        Import-Module $InputData.ManagementModulePath -ErrorAction Stop
        Import-Module $InputData.InsightsModulePath -ErrorAction Stop
        $root = [string]$InputData.Root
        if (-not (Test-Path -LiteralPath $root -PathType Container)) { throw "Shared data folder is unavailable: $root" }
        if ($InputData.Mode -eq 'History') {
            $rows = @(Get-SmartM365OrchestratorHistory -SharedDataFolderPath $root -From $InputData.HistoryFrom -To $InputData.HistoryTo -Server $InputData.HistoryServer -JobName $InputData.HistoryJob -Status $InputData.HistoryStatus)
            if ($InputData.OnlyFailures) { $rows = @($rows | Where-Object { $_.Status -in @('Failed', 'TimedOut', 'Interrupted', 'BlockedDependencyFailed', 'BlockedDependencyTimeout', 'BlockedPrerequisite') }) }
            return [pscustomobject]@{ Mode = 'History'; History = $rows; ReadAt = Get-Date }
        }
        $snapshot = $null; $recentRuns = @(); $servers = @(); $history = @(); $versions = @(); $owners = @{}
        $scriptVersions = $null; $scriptVersionsError = ''; $swiValidity = $null; $swiValidityError = ''
        if ($InputData.Mode -eq 'Full') {
            $Progress.Message = 'Reading configuration and history...'
            $snapshot = Get-SmartM365OrchestratorConfigurationSnapshot -SharedDataFolderPath $root
            $recentRuns = @(Get-SmartM365OrchestratorRecentRuns -SharedDataFolderPath $root -Days 21)
            $servers = @(Get-SmartM365OrchestratorServerStatus -SharedDataFolderPath $root -ClusterDocument $snapshot.Cluster)
            foreach ($server in $servers) {
                foreach ($jobName in @([string]$server.AssignedJobNames -split ', ' | Where-Object { $_ })) { $owners[$jobName] = [string]$server.Server }
            }
            $history = @(Get-SmartM365OrchestratorHistory -SharedDataFolderPath $root -From $InputData.HistoryFrom -To $InputData.HistoryTo)
            $versions = @(Get-SmartM365OrchestratorConfigurationVersions -SharedDataFolderPath $root)
            # Local job scripts (published manifest and committed catalog): header read only, isolated.
            $Progress.Message = 'Reading job script versions...'
            try {
                $documents = @($snapshot.Jobs)
                if (Test-Path -LiteralPath $InputData.JobsTemplatePath -PathType Leaf) { $documents += Read-SmartM365OrchestratorJson -Path $InputData.JobsTemplatePath }
                $paths = foreach ($document in $documents) {
                    foreach ($job in @($document.Jobs)) { if ($null -ne $job.PSObject.Properties['ScriptPath']) { [string]$job.ScriptPath } }
                }
                $scriptVersions = Get-SmartM365OrchestratorScriptVersion -SmartInventoryRoot $InputData.SmartInventoryRoot -ScriptPath @($paths) -Cache $InputData.ScriptVersionCache
            }
            catch { $scriptVersionsError = $_.Exception.Message }
            # SWI validity of the latest collections (DATA-LAST, receipts, SWI contracts): read-only, isolated.
            $Progress.Message = 'Checking SWI validity of the latest collections...'
            try { $swiValidity = Get-SmartM365OrchestratorSwiValidityForRoot -SharedDataFolderPath $root -SmartM365RootPath $InputData.SmartM365RootPath -JobsDocument $snapshot.Jobs }
            catch { $swiValidityError = $_.Exception.Message }
        }
        $Progress.Message = 'Reading live operations and requests...'
        $publishedCluster = Read-SmartM365OrchestratorJson -Path (Join-Path $root 'Config/Orchestrator-Cluster.json')
        $operations = Get-SmartM365OrchestratorOperations -SharedDataFolderPath $root -ClusterDocument $publishedCluster -MailFolderPath $InputData.MailFolderPath -MailHours 24
        $requests = @(Get-SmartM365OrchestratorRecentPipelineRuns -SharedDataFolderPath $root -Count 30)
        $liveServers = @(); $publishedHashes = $null
        if ($InputData.Mode -eq 'Live') {
            # Isolated: the Servers grid and the published-change check never hide the live operations.
            try {
                $configurationPaths = Get-SmartM365OrchestratorConfigurationPaths -SharedDataFolderPath $root
                $publishedHashes = [pscustomobject]@{ JobsHash = Get-SmartM365OrchestratorFileHash $configurationPaths.JobsPath; ClusterHash = Get-SmartM365OrchestratorFileHash $configurationPaths.ClusterPath }
                $liveServers = @(Get-SmartM365OrchestratorServerStatus -SharedDataFolderPath $root -ClusterDocument $publishedCluster)
            }
            catch { $publishedHashes = $null; $liveServers = @() }
        }
        $logEvents = $null; $logEventsError = ''
        $eventsOption = $InputData.PSObject.Properties['LogEventsHours']
        if ($null -ne $eventsOption -and [int]$eventsOption.Value -gt 0) {
            # Read-only and isolated: an unreadable log never hides the live operations.
            $Progress.Message = 'Reading resident orchestrator logs...'
            try { $logEvents = Get-SmartM365OrchestratorLogEvents -LogFolderPath $InputData.MailFolderPath -Server @($publishedCluster.ExpectedOrchestratorServers) -Hours ([int]$eventsOption.Value) -Cache $InputData.LogEventsCache }
            catch { $logEventsError = $_.Exception.Message }
        }
        $dashboard = $null; $dashboardError = ''
        $dashboardOption = $InputData.PSObject.Properties['IncludeDashboard']
        if ($null -ne $dashboardOption -and [bool]$dashboardOption.Value) {
            # Isolated like the log events: a dashboard failure never hides the live operations.
            $Progress.Message = 'Building dashboard: job health, election plan and resident logs...'
            try {
                $jobsDocument = if ($null -ne $snapshot) { $snapshot.Jobs } else { Read-SmartM365OrchestratorJson -Path (Join-Path $root 'Config/Orchestrator-Jobs.json') }
                $dashboardEvents = $logEvents
                if ($null -eq $dashboardEvents -and -not [string]::IsNullOrWhiteSpace([string]$InputData.MailFolderPath)) {
                    try { $dashboardEvents = Get-SmartM365OrchestratorLogEvents -LogFolderPath $InputData.MailFolderPath -Server @($publishedCluster.ExpectedOrchestratorServers) -Hours 24 -Cache $InputData.LogEventsCache } catch { $dashboardEvents = $null }
                }
                $dashboardRuns = if ($InputData.Mode -eq 'Full') { @($recentRuns) } else { $null }
                # Live refreshes reuse the SWI rows of the last full refresh, so blocking jobs stay listed.
                $swiOption = $InputData.PSObject.Properties['SwiRows']
                $dashboardSwi = if ($null -ne $swiValidity) { @($swiValidity.Rows) } elseif ($null -ne $swiOption) { @($swiOption.Value) } else { @() }
                $dashboard = Get-SmartM365OrchestratorDashboardData -SharedDataFolderPath $root -JobsDocument $jobsDocument -Operations $operations -Runs $dashboardRuns -Requests $requests -LogEvents $dashboardEvents -SwiRows $dashboardSwi
            }
            catch { $dashboardError = $_.Exception.Message }
        }
        $configuration = $null; $configurationError = ''
        $configurationOption = $InputData.PSObject.Properties['ConfigurationProfile']
        if ($null -ne $configurationOption -and -not [string]::IsNullOrWhiteSpace([string]$configurationOption.Value)) {
            # Read-only; values never reach the progress text or any log.
            $Progress.Message = 'Reading the configuration files...'
            try {
                Import-Module $InputData.ConfigViewModulePath -ErrorAction Stop
                $configuration = Get-SmartM365OrchestratorConfigurationView -RootPath $InputData.ConfigRootPath -ProfileKey ([string]$configurationOption.Value)
            }
            catch { $configurationError = $_.Exception.Message }
        }
        return [pscustomobject]@{
            Mode = $InputData.Mode; Snapshot = $snapshot; RecentRuns = $recentRuns; Servers = $servers
            LiveServers = $liveServers; PublishedHashes = $publishedHashes
            Owners = $owners; History = $history; Versions = $versions; Operations = $operations
            Requests = $requests; LogEvents = $logEvents; LogEventsError = $logEventsError
            Dashboard = $dashboard; DashboardError = $dashboardError
            Configuration = $configuration; ConfigurationError = $configurationError
            ScriptVersions = $scriptVersions; ScriptVersionsError = $scriptVersionsError
            SwiValidity = $swiValidity; SwiValidityError = $swiValidityError; ReadAt = Get-Date
        }
    }
}

function Start-GuiRefresh {
    param([ValidateSet('Full', 'Live', 'History')][string]$Mode = 'Live', [datetime]$HistoryFrom = [datetime]::MinValue, [datetime]$HistoryTo = [datetime]::MinValue, [switch]$OnlyFailures, [switch]$ShowLoading)
    $includeDashboard = $Mode -eq 'Full' -or ($Mode -eq 'Live' -and (Test-GuiDashboardTabVisible))
    if ($includeDashboard -and ($Mode -eq 'Full' -or $ShowLoading -or $null -eq $script:DashboardData)) {
        $script:DashboardLoading.Pending = $true
        if (-not $script:DashboardLoading.Clock.IsRunning) { $script:DashboardLoading.Clock.Restart() }
    }
    if ($null -ne $script:RefreshContext.Worker) {
        if ($Mode -eq 'Full' -or $null -eq $script:PendingRefreshRequest -or $script:PendingRefreshRequest.Mode -ne 'Full') {
            $script:PendingRefreshRequest = @{ Mode = $Mode; HistoryFrom = $HistoryFrom; HistoryTo = $HistoryTo; OnlyFailures = [bool]$OnlyFailures; ShowLoading = [bool]$ShowLoading }
        }
        return
    }
    $now = Get-Date
    $from = if ($Mode -eq 'History' -and $script:Controls.HistoryFromPicker.SelectedDate) { [datetime]$script:Controls.HistoryFromPicker.SelectedDate } else { $now.AddDays(-7).Date }
    $to = if ($Mode -eq 'History' -and $script:Controls.HistoryToPicker.SelectedDate) { ([datetime]$script:Controls.HistoryToPicker.SelectedDate).Date.AddDays(1).AddTicks(-1) } else { $now }
    if ($HistoryFrom -ne [datetime]::MinValue) { $from = $HistoryFrom }
    if ($HistoryTo -ne [datetime]::MinValue) { $to = $HistoryTo }
    $server = ''; $job = ''; $status = ''
    if ($Mode -eq 'History') {
        $server = Get-ComboText -Combo $script:Controls.HistoryServerCombo
        $job = Get-ComboText -Combo $script:Controls.HistoryJobCombo
        $status = Get-ComboText -Combo $script:Controls.HistoryStatusCombo
        if ($server -eq 'All') { $server = '' }; if ($job -eq 'All') { $job = '' }; if ($status -eq 'All') { $status = '' }
    }
    $inputData = [pscustomobject]@{
        Mode = $Mode; Root = $script:SharedDataFolderPath; MailFolderPath = $script:MailFolderPath
        ManagementModulePath = (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.Management.psm1')
        InsightsModulePath = (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.Insights.psm1')
        HistoryFrom = $from; HistoryTo = $to; HistoryServer = $server; HistoryJob = $job; HistoryStatus = $status; OnlyFailures = [bool]$OnlyFailures
        LogEventsHours = $(if ($Mode -ne 'History' -and (Test-GuiEventsTabVisible)) { Get-GuiEventsHours } else { 0 }); LogEventsCache = $script:LogEventsCache
        IncludeDashboard = $includeDashboard
        ConfigurationProfile = $(if ($Mode -ne 'History' -and (Test-GuiConfigurationTabVisible)) { [string]$script:EffectiveConfig.ProfileKey } else { '' })
        ConfigRootPath = $script:SmartM365RootPath; ConfigViewModulePath = (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.ConfigView.psm1')
        SmartInventoryRoot = $script:SmartInventoryRoot; JobsTemplatePath = (Join-Path $PSScriptRoot 'Orchestrator-Jobs.json.template'); ScriptVersionCache = $script:ScriptVersionCache
        SmartM365RootPath = $script:SmartM365RootPath
        SwiRows = @($script:SwiByName.Values | ForEach-Object { [pscustomobject]@{ Name = $_.Name; Swi = $_.Swi; Blocks = [bool](Get-OrchestratorGuiPropertyValue -Object $_ -Name 'Blocks' -DefaultValue $false); Detail = $_.Detail } })
    }
    $script:RefreshContext.Input = $inputData
    $script:RefreshContext.Progress.Message = ''
    $script:RefreshContext.Clock.Restart()
    if ($Mode -eq 'Full') {
        $script:Controls.MainTabs.IsEnabled = $false
        $script:Controls.PublishButton.IsEnabled = $false
    }
    if ($Mode -eq 'History') { $script:Controls.StatusText.Text = 'Refreshing history...' }
    else { $script:Controls.LastRefreshText.Text = 'Refreshing ' + $Mode.ToLowerInvariant() + '... last successful: ' + $(if ($script:LastSuccessfulRefresh) { $script:LastSuccessfulRefresh.ToString('yyyy-MM-dd HH:mm:ss') } else { 'never' }) }
    try { [void](Start-SmartM365OrchestratorGuiWorker -Context $script:RefreshContext -Work (Get-GuiRefreshWork) -InputObject $inputData) }
    catch {
        $script:Controls.MainTabs.IsEnabled = $null -ne $script:Snapshot
        $script:Controls.LastRefreshText.Text = 'Refresh failed: ' + $_.Exception.Message
        Write-GuiException -Context 'Background refresh could not start' -ErrorRecord $_
    }
}

function Set-GuiOperations {
    param([Parameter(Mandatory = $true)]$Operations)
    Refresh-MaintenanceView -Operations $Operations
    Add-GuiJobDurationEstimates -Operations $Operations
    $script:Controls.OperationsServersGrid.ItemsSource = @($Operations.Servers)
    $script:Controls.OperationsRunningGrid.ItemsSource = @($Operations.Running)
    $script:Controls.StopRunningJobButton.IsEnabled = $false
    $script:Controls.OperationsPendingGrid.ItemsSource = @($Operations.Pending)
    $script:Controls.OperationsIncidentsGrid.ItemsSource = @($Operations.Incidents)
    $script:Controls.OperationsMailsGrid.ItemsSource = @($Operations.Mails)
    $script:Controls.ServersMetricsCanvas.Tag = if ($Operations.PSObject.Properties['Metrics']) { @($Operations.Metrics) } else { @() }
    Update-OrchestratorServerMetricsChart -Canvas $script:Controls.ServersMetricsCanvas
}

function Receive-GuiRefresh {
    $received = Receive-SmartM365OrchestratorGuiWorker -Context $script:RefreshContext
    if ($null -eq $received) { return }
    $mode = [string]$script:RefreshContext.Input.Mode
    try {
        if ($received.Error) { throw $received.Error }
        $result = @($received.Output | Select-Object -First 1)[0]
        if ($null -eq $result) { throw 'The refresh worker returned no data.' }
        if ($mode -eq 'Full') {
            $script:Snapshot = $result.Snapshot
            $script:DraftJobs = Copy-JsonDocument -Document $script:Snapshot.Jobs
            $script:DraftCluster = Copy-JsonDocument -Document $script:Snapshot.Cluster
            # Older cluster document: the servers already declared are Active (same rule as the residents).
            [void](Initialize-SmartM365OrchestratorServerStates -ClusterDocument $script:DraftCluster)
            Update-JobHealth -Runs $result.RecentRuns
            $scriptVersionsProperty = $result.PSObject.Properties['ScriptVersions']
            if ($null -ne $scriptVersionsProperty -and $scriptVersionsProperty.Value -is [hashtable]) { $script:ScriptVersions = $scriptVersionsProperty.Value }
            $scriptVersionsErrorProperty = $result.PSObject.Properties['ScriptVersionsError']
            if ($null -ne $scriptVersionsErrorProperty -and $scriptVersionsErrorProperty.Value) { Write-GuiActivity -Level WARN -Message "Job script versions unavailable: $($scriptVersionsErrorProperty.Value)" }
            Set-GuiSwiValidity -Validity $(if ($result.PSObject.Properties['SwiValidity']) { $result.SwiValidity } else { $null }) -ErrorText $(if ($result.PSObject.Properties['SwiValidityError']) { [string]$result.SwiValidityError } else { '' })
            Refresh-PlanningView -Owners $result.Owners
            Refresh-LaunchView
            $servers = @(Refresh-ServersView -Servers $result.Servers)
            $script:Controls.HistoryServerCombo.ItemsSource = @('All') + @($script:DraftCluster.ExpectedOrchestratorServers)
            $script:Controls.HistoryServerCombo.SelectedIndex = 0
            $script:Controls.HistoryJobCombo.ItemsSource = @('All') + @($script:DraftJobs.Jobs.Name | Sort-Object)
            $script:Controls.HistoryJobCombo.SelectedIndex = 0
            Set-GuiEventsServers
            $script:HistoryRows = @($result.History)
            $script:Controls.HistoryGrid.ItemsSource = $script:HistoryRows
            $script:InitialHistoryLoaded = $true
            $script:Controls.VersionsGrid.ItemsSource = @($result.Versions)
            $script:Controls.ConnectionText.Text = "Tenant: $Tenant"
            $script:Controls.StatusText.Text = 'Shared configuration loaded'
            Write-GuiActivity -Message ("Configuration loaded: {0} jobs, {1} servers in {2} ms." -f @($script:DraftJobs.Jobs).Count, $servers.Count, $script:RefreshContext.Clock.ElapsedMilliseconds) -Level SUCCESS
        }
        if ($mode -eq 'History') {
            $script:HistoryRows = @($result.History)
            $script:Controls.HistoryGrid.ItemsSource = $script:HistoryRows
        }
        else {
            Set-GuiOperations -Operations $result.Operations
            if (-not $script:CancellationBusy -and -not $script:RequestSubmitBusy) { Set-GuiRequestRows -Rows @($result.Requests) }
            if ([int]$script:RefreshContext.Input.LogEventsHours -gt 0) { Set-GuiLogEvents -Result $result.LogEvents -ErrorMessage ([string]$result.LogEventsError) }
            if ([string]$script:RefreshContext.Input.ConfigurationProfile) { Set-GuiConfiguration -View $result.Configuration -ErrorMessage ([string]$result.ConfigurationError) }
            if ([bool]$script:RefreshContext.Input.IncludeDashboard) { Set-GuiDashboard -Dashboard $result.Dashboard -ErrorMessage ([string]$result.DashboardError) }
            if ($mode -eq 'Live') {
                try { Update-GuiPublishedConfigurationState -Result $result }
                catch { Write-GuiException -Context 'Servers and configuration refresh failed' -ErrorRecord $_ }
            }
        }
        if ($mode -eq 'History') {
            $script:Controls.StatusText.Text = "History refreshed: $(@($script:HistoryRows).Count) row(s)."
        }
        else {
            $script:LastSuccessfulRefresh = [datetime]$result.ReadAt
            $script:Controls.LastRefreshText.Text = 'Refreshed: ' + $script:LastSuccessfulRefresh.ToString('yyyy-MM-dd HH:mm:ss')
        }
    }
    catch {
        $lastGood = if ($script:LastSuccessfulRefresh) { $script:LastSuccessfulRefresh.ToString('yyyy-MM-dd HH:mm:ss') } else { 'never' }
        if ($mode -ne 'History') { $script:Controls.LastRefreshText.Text = "STALE - last successful refresh: $lastGood" }
        $script:Controls.StatusText.Text = 'Shared data refresh failed: ' + $_.Exception.Message
        if ($mode -ne 'History') {
            $script:Controls.OperationsRunningGrid.ItemsSource = @()
            $script:Controls.OperationsPendingGrid.ItemsSource = @()
            $script:Controls.StopRunningJobButton.IsEnabled = $false
        }
        Write-GuiException -Context 'Background refresh failed' -ErrorRecord $_
    }
    finally {
        if ([bool]$script:RefreshContext.Input.IncludeDashboard) { Stop-GuiDashboardLoading }
        if ($mode -eq 'Full') {
            $script:Controls.MainTabs.IsEnabled = $null -ne $script:Snapshot
            $script:Controls.PublishButton.IsEnabled = $null -ne $script:Snapshot
        }
        if ($script:PendingRefreshRequest) {
            $nextRequest = $script:PendingRefreshRequest
            $script:PendingRefreshRequest = $null
            Start-GuiRefresh @nextRequest
        }
    }
}

function Refresh-AllViews {
    Start-GuiRefresh -Mode Full
}

function Refresh-AllViewsSync {
    try {
        $script:Snapshot = Get-SmartM365OrchestratorConfigurationSnapshot -SharedDataFolderPath $script:SharedDataFolderPath
        $script:DraftJobs = Copy-JsonDocument -Document $script:Snapshot.Jobs
        $script:DraftCluster = Copy-JsonDocument -Document $script:Snapshot.Cluster
        [void](Initialize-SmartM365OrchestratorServerStates -ClusterDocument $script:DraftCluster)
        try { Update-JobHealth } catch { $script:HealthByName = @{}; Write-GuiException -Context 'Job health could not be computed' -ErrorRecord $_ }
        try {
            $templatePath = Join-Path $PSScriptRoot 'Orchestrator-Jobs.json.template'
            $templatePaths = if (Test-Path -LiteralPath $templatePath -PathType Leaf) { Get-GuiJobScriptPaths -JobsDocument (Read-SmartM365OrchestratorJson -Path $templatePath) } else { @() }
            $script:ScriptVersions = Get-SmartM365OrchestratorScriptVersion -SmartInventoryRoot $script:SmartInventoryRoot -ScriptPath (@(Get-GuiJobScriptPaths -JobsDocument $script:DraftJobs) + @($templatePaths)) -Cache $script:ScriptVersionCache
        }
        catch { Write-GuiActivity -Level WARN -Message "Job script versions unavailable: $($_.Exception.Message)" }
        try { Set-GuiSwiValidity -Validity (Get-SmartM365OrchestratorSwiValidityForRoot -SharedDataFolderPath $script:SharedDataFolderPath -SmartM365RootPath $script:SmartM365RootPath -JobsDocument $script:DraftJobs) }
        catch { Set-GuiSwiValidity -Validity $null -ErrorText $_.Exception.Message }
        Refresh-PlanningView
        Refresh-LaunchView
        $servers = @(Refresh-ServersView)
        $historyNow = Get-Date
        $historyFrom = $historyNow.AddDays(-7)
        $history7 = @(Get-SmartM365OrchestratorHistory -SharedDataFolderPath $script:SharedDataFolderPath -From $historyFrom.Date -To $historyNow.Date.AddDays(1).AddTicks(-1))
        if (-not $script:InitialHistoryLoaded) {
            $script:HistoryRows = $history7
            $script:Controls.HistoryGrid.ItemsSource = $script:HistoryRows
            $script:InitialHistoryLoaded = $true
        }
        elseif ($null -ne $script:Controls.MainTabs.SelectedItem -and
            [string]$script:Controls.MainTabs.SelectedItem.Header -eq 'History') { Refresh-HistoryView }
        $operations = Refresh-OperationsView
        Refresh-RequestsView
        if ($null -ne $operations) {
            Set-GuiDashboard -Dashboard (Get-SmartM365OrchestratorDashboardData -SharedDataFolderPath $script:SharedDataFolderPath -JobsDocument $script:DraftJobs -Operations $operations -Runs $script:RecentRuns -Requests $script:RequestRows -SwiRows @($script:SwiByName.Values))
        }
        $script:Controls.VersionsGrid.ItemsSource = @(Get-SmartM365OrchestratorConfigurationVersions -SharedDataFolderPath $script:SharedDataFolderPath)
        $script:Controls.ConnectionText.Text = "Tenant: $Tenant"
        $script:Controls.LastRefreshText.Text = 'Refreshed: ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        $script:Controls.StatusText.Text = 'Shared configuration loaded'
        Write-GuiActivity -Message ("Configuration loaded: {0} jobs, {1} servers." -f @($script:DraftJobs.Jobs).Count, @($servers).Count) -Level SUCCESS
    }
    catch {
        $script:Controls.StatusText.Text = 'Refresh failed'
        Write-GuiException -Context 'Refresh failed' -ErrorRecord $_
        $script:OrchestratorDialog.Show($_.Exception.Message, 'Refresh failed', 'OK', 'Error') | Out-Null
    }
}

function Apply-SelectedJobToDraft {
    param([switch]$SkipDependentsConfirmation)
    $row = $script:Controls.PlanningGrid.SelectedItem
    if ($null -eq $row) { throw 'Select a job first.' }
    $job = @($script:DraftJobs.Jobs | Where-Object { [string]$_.Name -eq [string]$row.Name })[0]
    $enabled = [bool]$script:Controls.JobEnabledCheck.IsChecked
    if ([bool]$job.Enabled -and -not $enabled -and -not $SkipDependentsConfirmation) {
        $dependents = @(Get-SmartM365OrchestratorDependents -JobsDocument $script:DraftJobs -JobName ([string]$job.Name) -EnabledOnly)
        if ($dependents.Count) {
            $warning = "{0} enabled job(s) depend on '{1}':`n{2}`n`nOnce disabled, the dependency gate ignores it and these jobs run without waiting for it. Disable it anyway?" -f $dependents.Count, $job.Name, ($dependents -join [Environment]::NewLine)
            if ($script:OrchestratorDialog.Show($warning, 'Job used as a dependency', 'YesNo', 'Warning') -ne 'Yes') { return }
        }
    }
    $dependsOn = @($script:Controls.DependsOnBox.Text -split '[,;\r\n]' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Select-Object -Unique)
    if ([string]$job.Name -in $dependsOn) { throw 'A job cannot depend on itself.' }
    $dependencyMode = Get-ComboText -Combo $script:Controls.DependencyModeCombo
    if ([string]::IsNullOrWhiteSpace($dependencyMode)) { $dependencyMode = 'LatestOccurrence' }
    $dependencyMaxAge = 0
    if (-not [int]::TryParse(([string]$script:Controls.DependencyMaxAgeBox.Text).Trim(), [ref]$dependencyMaxAge) -or $dependencyMaxAge -lt 0) {
        if (-not [string]::IsNullOrWhiteSpace($script:Controls.DependencyMaxAgeBox.Text)) { throw 'Dependency max age must be a whole number of hours (0 = automatic).' }
        $dependencyMaxAge = 0
    }
    $jobBackup = Copy-JsonDocument -Document $job
    $scheduleType = Get-ComboText -Combo $script:Controls.ScheduleTypeCombo
    $times = @($script:Controls.TimesBox.Text -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    $days = @(Get-SelectedScheduleDays)
    if ($scheduleType -eq 'Weekly' -and $days.Count -eq 0) { throw 'Select at least one day for a Weekly schedule.' }
    if ($scheduleType -ne 'Weekly') { $days = @() }
    [object]$scheduleDays = [string[]]@()
    if ($scheduleType -eq 'Weekly') { [object]$scheduleDays = [string[]]$days }
    $mode = Get-ComboText -Combo $script:Controls.AssignmentCombo
    $pinnedServer = Get-ComboText -Combo $script:Controls.PinnedServerCombo
    $previousMode = [string](Get-OrchestratorGuiPropertyValue -Object $job -Name 'AssignmentMode' -DefaultValue 'Legacy')
    $previousAllowedServers = @(Get-OrchestratorGuiPropertyValue -Object $job -Name 'AllowedServers' -DefaultValue @())
    $newAllowedServers = @()

    if ($mode -eq 'Pinned') {
        if ([string]::IsNullOrWhiteSpace($pinnedServer)) { throw 'Select exactly one pinned server.' }
        $matchingServers = @($script:DraftCluster.ExpectedOrchestratorServers | Where-Object { [string]$_ -ieq $pinnedServer })
        if ($matchingServers.Count -ne 1) { throw "Pinned server '$pinnedServer' is not an expected Orchestrator server." }
        $pinnedServer = [string]$matchingServers[0]
        $newAllowedServers = @($pinnedServer)
    }
    elseif (-not [string]::IsNullOrWhiteSpace($pinnedServer)) {
        throw "Pinned server '$pinnedServer' cannot be used while Assignment is '$mode'. Select Assignment 'Pinned' first."
    }

    Set-JsonProperty -Object $job -Name Enabled -Value $enabled
    Set-JsonProperty -Object $job.Schedule -Name Type -Value $scheduleType
    Set-JsonProperty -Object $job.Schedule -Name Times -Value $times
    Set-JsonProperty -Object $job.Schedule -Name DaysOfWeek -Value $scheduleDays
    Set-JsonProperty -Object $job.Schedule -Name MissedRunPolicy -Value (Get-ComboText -Combo $script:Controls.MissedPolicyCombo)
    Set-JsonProperty -Object $job -Name AssignmentMode -Value $mode
    Set-JsonProperty -Object $job -Name AllowedServers -Value $newAllowedServers
    Set-JsonProperty -Object $job -Name TimeoutMinutes -Value ([int]$script:Controls.TimeoutBox.Text)
    Set-JsonProperty -Object $job -Name MaxRetries -Value ([int]$script:Controls.RetriesBox.Text)
    Set-JsonProperty -Object $job -Name RetryDelaySeconds -Value ([int]$script:Controls.RetryDelayBox.Text)
    Set-JsonProperty -Object $job -Name EstimatedDurationMinutes -Value ([double]::Parse($script:Controls.DurationBox.Text, [System.Globalization.CultureInfo]::InvariantCulture))
    Set-JsonProperty -Object $job -Name DependsOn -Value ([string[]]$dependsOn)
    if ($dependsOn.Count -gt 0 -or $job.PSObject.Properties['DependencyMode']) { Set-JsonProperty -Object $job -Name DependencyMode -Value $dependencyMode }
    if ($dependencyMaxAge -gt 0 -or $job.PSObject.Properties['DependencyMaxAgeHours']) { Set-JsonProperty -Object $job -Name DependencyMaxAgeHours -Value $dependencyMaxAge }
    $validation = Test-SmartM365OrchestratorJobsDocument -Document $script:DraftJobs
    if (-not $validation.Valid) {
        $index = [array]::IndexOf(@($script:DraftJobs.Jobs), $job)
        $jobs = @($script:DraftJobs.Jobs); $jobs[$index] = $jobBackup; $script:DraftJobs.Jobs = $jobs
        throw ("The draft was not changed:`n" + ($validation.Errors -join [Environment]::NewLine))
    }
    Refresh-PlanningView
    $script:Controls.PlanningGrid.SelectedItem = @($script:PlanningRows | Where-Object Name -eq $job.Name)[0]
    $script:Controls.StatusText.Text = "Draft changed: $($job.Name)"
    # The draft is local until it is published: say so, the status line alone is easy to miss.
    # Informational only: never block the unattended smoke test.
    if (-not $SmokeTest) { $script:OrchestratorDialog.Show(("Job '{0}' was changed in the draft.{1}{1}Nothing runs differently yet: click Publish changes to apply it to the orchestrators." -f $job.Name, [Environment]::NewLine), 'Draft changed', 'OK', 'Information') | Out-Null }
    $previousAllowedText = if ($previousAllowedServers.Count -gt 0) { $previousAllowedServers -join ',' } else { '<none>' }
    $newAllowedText = if ($newAllowedServers.Count -gt 0) { $newAllowedServers -join ',' } else { '<none>' }
    Write-GuiActivity -Message ("Job '{0}' updated in the draft. AssignmentMode={1}->{2}; AllowedServers={3}->{4}; DependsOn={5}; DependencyMode={6}; DependencyMaxAgeHours={7}" -f $job.Name, $previousMode, $mode, $previousAllowedText, $newAllowedText, ($(if ($dependsOn.Count) { $dependsOn -join ',' } else { '<none>' })), $dependencyMode, $dependencyMaxAge)
}

function Set-DraftServerState {
    # Writes ServerStates explicitly; a server without an entry would be Inactive.
    param([Parameter(Mandatory)][string]$Server, [Parameter(Mandatory)][ValidateSet('Active', 'Inactive')][string]$State, $Cluster = $script:DraftCluster)
    [void](Initialize-SmartM365OrchestratorServerStates -ClusterDocument $Cluster)
    Set-JsonProperty -Object $Cluster.ServerStates -Name $Server.ToUpperInvariant() -Value $State
}

function Add-ServerToDraft {
    # A new server is always added Inactive: activate it after checking it (Apply server settings to draft).
    param([Parameter(Mandatory)][string]$Server, $Cluster = $script:DraftCluster)
    $server = $Server.Trim().ToUpperInvariant()
    if ($server -notmatch '^[A-Z0-9._-]+$') { throw 'Enter a valid server name.' }
    if ($server -in @($Cluster.ExpectedOrchestratorServers)) { throw 'This server already exists.' }
    [void](Initialize-SmartM365OrchestratorServerStates -ClusterDocument $Cluster)
    Set-JsonProperty -Object $Cluster -Name ExpectedOrchestratorServers -Value (@($Cluster.ExpectedOrchestratorServers) + $server)
    if (-not $Cluster.PSObject.Properties['ElectionWeightsByServer'] -or $null -eq $Cluster.ElectionWeightsByServer) {
        Set-JsonProperty -Object $Cluster -Name ElectionWeightsByServer -Value ([pscustomobject]@{})
    }
    Set-JsonProperty -Object $Cluster.ElectionWeightsByServer -Name $server -Value 1.0
    Set-DraftServerState -Server $server -State 'Inactive' -Cluster $Cluster
    return $server
}

function Test-GuiDraftChanged {
    # True when the draft differs from the published configuration it was loaded from.
    if ($null -eq $script:Snapshot -or $null -eq $script:DraftJobs -or $null -eq $script:DraftCluster) { return $false }
    $publishedJobs = Copy-JsonDocument -Document $script:Snapshot.Jobs
    $publishedCluster = Copy-JsonDocument -Document $script:Snapshot.Cluster
    [void](Initialize-SmartM365OrchestratorServerStates -ClusterDocument $publishedCluster)
    $sameJobs = ($script:DraftJobs | ConvertTo-Json -Depth 100 -Compress) -ceq ($publishedJobs | ConvertTo-Json -Depth 100 -Compress)
    $sameCluster = ($script:DraftCluster | ConvertTo-Json -Depth 100 -Compress) -ceq ($publishedCluster | ConvertTo-Json -Depth 100 -Compress)
    return -not ($sameJobs -and $sameCluster)
}

function Publish-GuiServerAddition {
    # Add server publishes at once (GUI 1.6.27): the server is added Inactive to the PUBLISHED cluster,
    # cluster file only, so other unpublished draft changes stay in the draft and are not published.
    param([Parameter(Mandatory)][string]$Server, [switch]$SkipConfirmation)
    if ($null -eq $script:Snapshot) { throw 'The shared configuration is not loaded yet.' }
    $draftChanged = Test-GuiDraftChanged
    $published = Copy-JsonDocument -Document $script:Snapshot.Cluster
    $name = Add-ServerToDraft -Server $Server -Cluster $published
    if (-not $SkipConfirmation) {
        $note = if ($draftChanged) { "`n`nYour other unpublished draft changes are not published; they stay in the draft." } else { '' }
        $question = "Add server '{0}' and publish it now?`n`nIt is added Inactive: its orchestrator runs (heartbeat, monitoring, maintenance) but takes no job until you set it Active.{1}" -f $name, $note
        if ($script:OrchestratorDialog.Show($question, 'Add server', 'YesNo', 'Question') -ne 'Yes') { return $null }
    }
    $result = Publish-SmartM365OrchestratorConfiguration -SharedDataFolderPath $script:SharedDataFolderPath -JobsDocument $script:Snapshot.Jobs -ClusterDocument $published -ExpectedClusterHash $script:Snapshot.ClusterHash -ChangeSummary ("Server {0} added (Inactive) from SmartM365 Orchestrator GUI" -f $name) -ClusterOnly
    Write-GuiActivity -Message ("Server '{0}' added Inactive and published as version {1}." -f $name, $result.VersionId) -Level SUCCESS
    if ($draftChanged) {
        # Keep the draft: add the server to it and move its publication base to the new version.
        if ($name -notin @($script:DraftCluster.ExpectedOrchestratorServers)) { [void](Add-ServerToDraft -Server $name) }
        $script:Snapshot = Get-SmartM365OrchestratorConfigurationSnapshot -SharedDataFolderPath $script:SharedDataFolderPath
        [void](Refresh-ServersView)
    }
    else { Refresh-AllViews }
    return [pscustomobject]@{ Server = $name; VersionId = [string]$result.VersionId }
}

function Update-GuiPublishedConfigurationState {
    # Live refresh (every 60 s): a configuration published elsewhere reloads every view, unless the
    # draft has unpublished changes (kept; Publish changes then asks for a refresh). Otherwise the
    # Servers grid is refreshed (heartbeat, versions, CPU, RAM, drive C:).
    param([Parameter(Mandatory)]$Result)
    $hashes = Get-OrchestratorGuiPropertyValue -Object $Result -Name 'PublishedHashes' -DefaultValue $null
    if ($null -eq $script:Snapshot -or $null -eq $hashes) { return }
    $draftChanged = Test-GuiDraftChanged
    $changed = [string]$hashes.JobsHash -ne [string]$script:Snapshot.JobsHash -or [string]$hashes.ClusterHash -ne [string]$script:Snapshot.ClusterHash
    if ($changed) {
        if ($draftChanged) {
            if (-not $script:PublishedChangeWarned) { Write-GuiActivity -Level WARN -Message 'The shared configuration was published elsewhere; the draft has unpublished changes and is kept. Publish changes will ask for a refresh.' }
            $script:PublishedChangeWarned = $true
            $script:Controls.StatusText.Text = 'Shared configuration changed elsewhere; your draft has unpublished changes (Refresh discards them).'
            return
        }
        $script:PublishedChangeWarned = $false
        Write-GuiActivity -Message 'The shared configuration was published elsewhere: every view is reloaded.'
        Start-GuiRefresh -Mode Full
        return
    }
    $liveServers = @(Get-OrchestratorGuiPropertyValue -Object $Result -Name 'LiveServers' -DefaultValue @())
    if (-not $draftChanged -and $liveServers.Count -gt 0) { [void](Refresh-ServersView -Servers $liveServers) }
}

function Apply-SelectedServerToDraft {
    $row = $script:Controls.ServersGrid.SelectedItem
    if ($null -eq $row) { throw 'Select a server first.' }
    $server = ([string]$row.Server).ToUpperInvariant()
    $previousState = [string]$row.State
    $newState = if ([bool]$script:Controls.ServerActiveCheck.IsChecked) { 'Active' } else { 'Inactive' }
    $weight = [double]::Parse($script:Controls.ServerWeightBox.Text, [System.Globalization.CultureInfo]::InvariantCulture)
    if ($weight -le 0) { throw 'Election weight must be greater than zero.' }
    if (-not $script:DraftCluster.PSObject.Properties['ElectionWeightsByServer']) {
        Set-JsonProperty -Object $script:DraftCluster -Name ElectionWeightsByServer -Value ([pscustomobject]@{})
    }
    Set-JsonProperty -Object $script:DraftCluster.ElectionWeightsByServer -Name $server -Value $weight
    if (-not $script:DraftCluster.PSObject.Properties['ServerJobPolicies']) {
        Set-JsonProperty -Object $script:DraftCluster -Name ServerJobPolicies -Value ([pscustomobject]@{})
    }
    $policy = Get-ComboText -Combo $script:Controls.ServerPolicyCombo
    $policyValues = @($policy -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    if ($policyValues.Count -eq 0) {
        [void]$script:DraftCluster.ServerJobPolicies.PSObject.Properties.Remove($server)
    }
    else {
        Set-JsonProperty -Object $script:DraftCluster.ServerJobPolicies -Name $server -Value ([pscustomobject]@{ OnlyJobsRequiring = $policyValues })
    }
    Set-DraftServerState -Server $server -State $newState
    $validation = Test-SmartM365OrchestratorClusterDocument -Document $script:DraftCluster
    if (-not $validation.Valid) { throw $validation.Errors -join [Environment]::NewLine }
    Refresh-ServersView | Out-Null
    $script:Controls.StatusText.Text = "Draft changed: $server"
    $stateNote = if ($newState -eq $previousState) { '' } elseif ($newState -eq 'Active') { "{0}{0}State: Active. After publication the server can be elected and take jobs; the election plan is rebalanced automatically." -f [Environment]::NewLine } else { "{0}{0}State: Inactive. After publication the server launches no new job (running jobs finish); its elected jobs move to the active servers and jobs pinned to it wait." -f [Environment]::NewLine }
    # Informational only: never block the unattended smoke test.
    if (-not $SmokeTest) { $script:OrchestratorDialog.Show(("Server '{0}' was changed in the draft.{2}{1}{1}Nothing changes yet: click Publish changes to apply it to the orchestrators." -f $server, [Environment]::NewLine, $stateNote), 'Draft changed', 'OK', 'Information') | Out-Null }
    Write-GuiActivity -Message "Server '$server' settings updated in the draft. State=$previousState->$newState"
}

function Get-DraftValidation {
    # Run by Publish changes before anything is written (GUI 1.6.27 has no separate Validate draft button).
    $jobsValidation = Test-SmartM365OrchestratorJobsDocument -Document $script:DraftJobs
    $clusterValidation = Test-SmartM365OrchestratorClusterDocument -Document $script:DraftCluster
    $consistencyValidation = Test-SmartM365OrchestratorConfigurationConsistency -JobsDocument $script:DraftJobs -ClusterDocument $script:DraftCluster
    [pscustomobject]@{
        Errors = @(@($jobsValidation.Errors) + @($clusterValidation.Errors) + @($consistencyValidation.Errors) | Where-Object { $_ })
        Warnings = @(@($jobsValidation.Warnings) + @($clusterValidation.Warnings) + @($consistencyValidation.Warnings) | Where-Object { $_ })
    }
}

function Publish-Draft {
    $validation = Get-DraftValidation
    if ($validation.Errors.Count -gt 0) {
        $script:OrchestratorDialog.Show(("The draft is not valid; nothing was published.`n`n" + ($validation.Errors -join [Environment]::NewLine)), 'Invalid draft', 'OK', 'Error') | Out-Null
        return
    }
    $enabled = @($script:DraftJobs.Jobs | Where-Object Enabled).Count
    $confirmation = "The draft is valid. Publish the shared configuration?`n`nJobs: $(@($script:DraftJobs.Jobs).Count) ($enabled enabled)`nServers: $(@($script:DraftCluster.ExpectedOrchestratorServers).Count)`n`nEvery Orchestrator server will reload it automatically."
    if ($validation.Warnings.Count -gt 0) { $confirmation += "`n`nWarnings:`n" + ($validation.Warnings -join [Environment]::NewLine) }
    if ($script:OrchestratorDialog.Show($confirmation, 'Publish shared configuration', 'YesNo', 'Warning') -ne 'Yes') { return }
    try {
        Write-GuiActivity -Message ("Publication requested. ExpectedJobsHash={0}; ExpectedClusterHash={1}" -f $script:Snapshot.JobsHash, $script:Snapshot.ClusterHash)
        $result = Publish-SmartM365OrchestratorConfiguration `
            -SharedDataFolderPath $script:SharedDataFolderPath `
            -JobsDocument $script:DraftJobs `
            -ClusterDocument $script:DraftCluster `
            -ExpectedJobsHash $script:Snapshot.JobsHash `
            -ExpectedClusterHash $script:Snapshot.ClusterHash `
            -ChangeSummary 'Published from SmartM365 Orchestrator GUI'
        Write-GuiActivity -Message "Configuration published as version $($result.VersionId)." -Level SUCCESS
        $script:OrchestratorDialog.Show("Configuration published successfully.`nVersion: $($result.VersionId)", 'Published', 'OK', 'Information') | Out-Null
        Refresh-AllViews
    }
    catch {
        Write-GuiException -Context 'Publication failed' -ErrorRecord $_
        $script:OrchestratorDialog.Show($_.Exception.Message, 'Publication failed', 'OK', 'Error') | Out-Null
    }
}

function Request-ElectionRebalance {
    param([switch]$SkipConfirmation)

    if (-not $SkipConfirmation) {
        $confirmation = "Recalculate elected job owners now?`n`nThe active published configuration, current live capabilities, server weights, policies and duration history will be used.`nPinned and Manual jobs are not changed.`n`nUnsaved draft changes are not included; publish them first if they must affect this rebalance."
        if ($script:OrchestratorDialog.Show($confirmation, 'Rebalance elected jobs', 'YesNo', 'Warning') -ne 'Yes') { return $null }
    }
    try {
        $result = Request-SmartM365OrchestratorRebalance -SharedDataFolderPath $script:SharedDataFolderPath
        $script:Controls.StatusText.Text = "Rebalance requested: $($result.RequestId)"
        Write-GuiActivity -Message ("Election rebalance requested. RequestId={0}; RequestPath={1}" -f $result.RequestId, $result.RequestPath) -Level SUCCESS
        if (-not $SkipConfirmation) {
            $script:OrchestratorDialog.Show("Rebalance request submitted successfully.`nRequest: $($result.RequestId)`n`nA resident Orchestrator will apply it on its next tick. Refresh the GUI to see the new owners.", 'Rebalance requested', 'OK', 'Information') | Out-Null
        }
        return $result
    }
    catch {
        Write-GuiException -Context 'Election rebalance request failed' -ErrorRecord $_
        if ($SkipConfirmation) { throw }
        $script:OrchestratorDialog.Show($_.Exception.Message, 'Rebalance failed', 'OK', 'Error') | Out-Null
        return $null
    }
}

$smartM365Root = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$script:SmartM365RootPath = $smartM365Root
$tenantContextPath = Join-Path -Path $smartM365Root -ChildPath 'Config\SmartM365-TenantContext.ps1'
. $tenantContextPath
$script:EffectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
Import-Module (Join-Path $smartM365Root 'Modules/SmartM365.Core/SmartM365.Core.psd1') -MinimumVersion '1.0.87' -ErrorAction Stop
$localConfigPath = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-Inventory-Orchestrator.local.json'; $localConfigPath = Resolve-SmartM365JsonConfigurationPath -Path $localConfigPath
if (-not (Test-Path -LiteralPath $localConfigPath)) {
    Copy-Item -LiteralPath (Get-SmartM365JsonTemplateName -Path $localConfigPath) -Destination $localConfigPath -ErrorAction Stop
}
$localConfig = Get-Content -LiteralPath $localConfigPath -Raw | ConvertFrom-Json -Depth 100
if (Get-Command Sync-SmartM365JsonConfigWithTemplate -ErrorAction SilentlyContinue) {
    $localConfig = Sync-SmartM365JsonConfigWithTemplate -Config $localConfig -Path $localConfigPath
}
if ([string]::IsNullOrWhiteSpace($SharedDataFolderPath)) {
    $dataFolder = Resolve-ConfigTokens -Value (Get-ConfigValue -Config $localConfig -Name 'OrchestratorDataFolderPath' -DefaultValue (Join-Path $PSScriptRoot 'Output'))
    if ((Split-Path -Path $dataFolder -Leaf) -eq $env:COMPUTERNAME) { $dataFolder = Split-Path -Path $dataFolder -Parent }
    $SharedDataFolderPath = $dataFolder
}
$script:SharedDataFolderPath = [System.IO.Path]::GetFullPath($SharedDataFolderPath)
$startupContextMs = $script:StartupClock.ElapsedMilliseconds
$resolvedGuiLogFolderPath = if (-not [string]::IsNullOrWhiteSpace($GuiLogFolderPath)) {
    [System.IO.Path]::GetFullPath((Resolve-ConfigTokens -Value $GuiLogFolderPath))
}
else {
    $logAllRootPath = Resolve-ConfigTokens -Value (Get-ConfigValue -Config $localConfig -Name 'LogAllRootPath' -DefaultValue (Join-Path $PSScriptRoot 'Logs'))
    Join-Path -Path $logAllRootPath -ChildPath (Join-Path 'SmartM365-Orchestrator-GUI' $env:COMPUTERNAME)
}
$script:GuiLogPath = Initialize-GuiLogPath -PreferredFolderPath $resolvedGuiLogFolderPath
$startupLogMs = $script:StartupClock.ElapsedMilliseconds
$script:MailFolderPath = Join-Path -Path (Resolve-ConfigTokens -Value (Get-ConfigValue -Config $localConfig -Name 'LogAllRootPath' -DefaultValue (Join-Path $PSScriptRoot 'Logs'))) -ChildPath 'SmartM365-Orchestrator'
Write-GuiActivity -Message ("GUI session started. Version={0}; Tenant={1}; User={2}; Computer={3}; SharedDataFolderPath={4}; LogPath={5}" -f $script:AppVersion, $Tenant, [Security.Principal.WindowsIdentity]::GetCurrent().Name, $env:COMPUTERNAME, $script:SharedDataFolderPath, $script:GuiLogPath)

$bootstrapJobsPath = ConvertTo-SmartM365InventoryConfigLocation -Path (Join-Path -Path $PSScriptRoot -ChildPath 'Orchestrator-Jobs.json')
$selectedBootstrap=Get-SmartM365JsonReadPath $bootstrapJobsPath -Optional
if ($selectedBootstrap) { $bootstrapJobsPath=$selectedBootstrap } else { $bootstrapJobsPath = Join-Path -Path $PSScriptRoot -ChildPath 'Orchestrator-Jobs.json.template' }
$bootstrapCluster = [pscustomobject][ordered]@{
    SchemaVersion = 1
    ExpectedOrchestratorServers = @(Get-ConfigValue -Config $localConfig -Name 'ExpectedOrchestratorServers' -DefaultValue @())
    ElectionWeightsByServer = Get-ConfigValue -Config $localConfig -Name 'ElectionWeightsByServer' -DefaultValue ([pscustomobject]@{})
    ServerJobPolicies = Get-ConfigValue -Config $localConfig -Name 'ServerJobPolicies' -DefaultValue ([pscustomobject]@{})
    PeerMonitoringEnabled = [bool](Get-ConfigValue -Config $localConfig -Name 'PeerMonitoringEnabled' -DefaultValue $true)
    PeerJobMonitoringEnabled = [bool](Get-ConfigValue -Config $localConfig -Name 'PeerJobMonitoringEnabled' -DefaultValue $true)
    PeerMonitoringCheckIntervalSeconds = [int](Get-ConfigValue -Config $localConfig -Name 'PeerMonitoringCheckIntervalSeconds' -DefaultValue 60)
    PeerHeartbeatStaleMinutes = [int](Get-ConfigValue -Config $localConfig -Name 'PeerHeartbeatStaleMinutes' -DefaultValue 5)
    PeerMonitoringConfirmationChecks = [int](Get-ConfigValue -Config $localConfig -Name 'PeerMonitoringConfirmationChecks' -DefaultValue 2)
    PeerJobStartGraceMinutes = [int](Get-ConfigValue -Config $localConfig -Name 'PeerJobStartGraceMinutes' -DefaultValue 15)
    PeerAlertReminderMinutes = [int](Get-ConfigValue -Config $localConfig -Name 'PeerAlertReminderMinutes' -DefaultValue 240)
    PeerAlertMailRetryMinutes = [int](Get-ConfigValue -Config $localConfig -Name 'PeerAlertMailRetryMinutes' -DefaultValue 15)
    PeerRecoveryEmailEnabled = [bool](Get-ConfigValue -Config $localConfig -Name 'PeerRecoveryEmailEnabled' -DefaultValue $true)
}
Initialize-SmartM365OrchestratorCentralConfiguration -SharedDataFolderPath $script:SharedDataFolderPath -BootstrapJobsPath $bootstrapJobsPath -BootstrapClusterDocument $bootstrapCluster | Out-Null
Write-GuiActivity -Message ('Startup timing (ms): splash={0}; context={1}; log={2}; shared configuration={3}.' -f $script:SplashReadyMs, ($startupContextMs - $script:SplashReadyMs), ($startupLogMs - $startupContextMs), ($script:StartupClock.ElapsedMilliseconds - $startupLogMs))

$window = ConvertFrom-OrchestratorGuiXaml -Text $xaml
$script:Controls = @{}
foreach ($name in @(
    'MainTabs', 'HeaderLogo', 'SharedPathText', 'ConnectionText', 'LastRefreshText', 'StatusText', 'RefreshButton', 'RebalanceButton', 'RequestPrepareButton', 'PublishButton',
    'DashboardBannerBorder', 'DashboardBannerText', 'DashboardTilesList', 'DashboardActionsGroup', 'DashboardActionsGrid', 'DashboardGrid', 'DashboardTimelineCanvas', 'DashboardMetricsCanvas', 'ServersMetricsCanvas', 'DashboardLoadingOverlay', 'DashboardLoadingText',
    'PlanningGrid', 'SelectedJobText', 'JobArgumentsBox', 'JobEnabledCheck', 'ScheduleTypeCombo', 'TimesBox', 'DaysPanel',
    'MondayCheck', 'TuesdayCheck', 'WednesdayCheck', 'ThursdayCheck', 'FridayCheck', 'SaturdayCheck', 'SundayCheck', 'MissedPolicyCombo',
    'AssignmentCombo', 'PinnedServerCombo', 'TimeoutBox', 'RetriesBox', 'RetryDelayBox', 'DurationBox', 'ApplyJobButton',
    'HistoryFromPicker', 'HistoryToPicker', 'HistoryServerCombo', 'HistoryJobCombo', 'HistoryStatusCombo', 'HistoryRefreshButton',
    'ExportCsvButton', 'ExportHtmlButton', 'HistoryGrid', 'ServersGrid', 'NewServerBox', 'AddServerButton', 'RemoveServerButton',
    'SelectedServerText', 'ServerActiveCheck', 'ServerWeightBox', 'ServerPolicyCombo', 'ApplyServerButton', 'TestServerButton', 'TestServerAsServiceAccountButton', 'ServerTestHintText', 'VersionsGrid', 'RollbackButton',
    'ConfigurationTabs', 'ConfigurationSettingsTab', 'ConfigurationServersTab', 'FooterText', 'VersionText',
    'AutoRefreshCheck', 'OperationsServersGrid', 'OperationsRunningGrid', 'JobStopReasonBox', 'StopRunningJobButton', 'OperationsPendingGrid', 'OperationsIncidentsGrid', 'OperationsMailsGrid',
    'DependsOnBox', 'DependencyModeCombo', 'DependencyMaxAgeBox', 'DependentsText', 'ReadinessGrid', 'IncludeDependenciesCheck', 'RequestRunButton',
    'Failures24hButton', 'RequestsGrid', 'RequestJobsGrid', 'CancellationReasonBox', 'CancellationProgressText', 'CancelRequestButton', 'CancelAllRequestsButton', 'MaintenanceBannerText', 'MaintenanceDetailText',
    'MaintenanceReasonBox', 'EnableMaintenanceButton', 'DisableMaintenanceButton',
    'EventsServerCombo', 'EventsLevelCombo', 'EventsCategoryCombo', 'EventsPeriodCombo', 'EventsTextBox', 'EventsGroupCheck', 'EventsRefreshButton', 'EventsSummaryText', 'EventsGrid', 'EventsDetailBox',
    'ConfigFilterBox', 'ConfigShowSensitiveButton', 'ConfigRefreshButton', 'ConfigOpenFolderButton', 'ConfigSummaryText', 'ConfigFileTabs', 'ConfigGlobalTab', 'ConfigTenantTab', 'ConfigMailTab', 'ConfigGlobalStatusText', 'ConfigTenantStatusText', 'ConfigMailStatusText', 'ConfigGlobalGrid', 'ConfigTenantGrid', 'ConfigMailGrid',
    'OrchestratorStartButton', 'OrchestratorStopButton', 'OrchestratorRestartButton', 'FullCollectButton', 'SendSummaryButton', 'OperationsTabs', 'OperationsRequestsTab', 'OnPremModeText', 'JobExecutionModeCombo', 'ApplyJobExecutionModeButton', 'ApplyModeToAllOnPremButton', 'LaunchJobsGrid', 'RunJobNowButton', 'RunJobAsServiceAccountButton', 'RequestJobButton', 'LaunchIncludeDependenciesCheck', 'LaunchJobCommandBox', 'ManualLaunchesGrid', 'RunManualLaunchButton', 'ManualLaunchCommandBox'
)) {
    $script:Controls[$name] = $window.FindName($name)
}

$iconPath = Join-Path $PSScriptRoot 'WorkplaceCloudHub.ico'
if (Test-Path -LiteralPath $iconPath) { $window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]$iconPath) }
$logoPath = Join-Path $PSScriptRoot 'WorkplaceCloudHub-lockup-WPF.png'
if (Test-Path -LiteralPath $logoPath) {
    $bitmap = [System.Windows.Media.Imaging.BitmapImage]::new()
    $bitmap.BeginInit(); $bitmap.CacheOption = 'OnLoad'; $bitmap.UriSource = [Uri]$logoPath; $bitmap.EndInit(); $bitmap.Freeze()
    $script:Controls.HeaderLogo.Source = $bitmap
}
$script:Controls.SharedPathText.Text = $script:SharedDataFolderPath
$script:Controls.FooterText.Text = 'Shared changes are validated, versioned and audited. Run and stop requests are executed by the owning orchestrators.'
$script:Controls.VersionText.Text = "v$($script:AppVersion)"
$script:Controls.HistoryFromPicker.SelectedDate = (Get-Date).AddDays(-7).Date
$script:Controls.HistoryToPicker.SelectedDate = (Get-Date).Date
$script:Controls.HistoryStatusCombo.ItemsSource = @('All', 'Success', 'CompletedWithWarnings', 'Failed', 'TimedOut', 'Interrupted', 'BlockedDependencyFailed', 'BlockedDependencyTimeout', 'BlockedPrerequisite', 'Skipped', 'Cancelled', 'Retried')
$script:Controls.HistoryStatusCombo.SelectedIndex = 0
$script:Controls.EventsLevelCombo.DisplayMemberPath = 'Label'
$script:Controls.EventsLevelCombo.ItemsSource = @(
    [pscustomobject]@{ Label = 'All'; Value = 'All' }, [pscustomobject]@{ Label = 'Errors and warnings'; Value = 'ErrorsAndWarnings' },
    [pscustomobject]@{ Label = 'ERROR'; Value = 'ERROR' }, [pscustomobject]@{ Label = 'WARN'; Value = 'WARN' }, [pscustomobject]@{ Label = 'INFO'; Value = 'INFO' }
)
$script:Controls.EventsLevelCombo.SelectedIndex = 0
$script:Controls.EventsCategoryCombo.ItemsSource = @('All', 'Lifecycle', 'StopRequest', 'Heartbeat', 'StatePersistence', 'ConfigReload', 'Maintenance', 'Error', 'Warning')
$script:Controls.EventsCategoryCombo.SelectedIndex = 0
$script:Controls.EventsPeriodCombo.DisplayMemberPath = 'Label'
$script:Controls.EventsPeriodCombo.ItemsSource = @(
    [pscustomobject]@{ Label = '1 h'; Hours = 1 }, [pscustomobject]@{ Label = '6 h'; Hours = 6 }, [pscustomobject]@{ Label = '24 h'; Hours = 24 },
    [pscustomobject]@{ Label = '3 days'; Hours = 72 }, [pscustomobject]@{ Label = '7 days'; Hours = 168 }
)
$script:Controls.EventsPeriodCombo.SelectedIndex = 2
$script:Controls.EventsServerCombo.ItemsSource = @('All')
$script:Controls.EventsServerCombo.SelectedIndex = 0

$script:Controls.ScheduleTypeCombo.Add_SelectionChanged({ Update-ScheduleDaysControlState })
$script:Controls.AssignmentCombo.Add_SelectionChanged({ Update-PinnedServerControlState })
$script:Controls.PinnedServerCombo.Add_SelectionChanged({
    if ($null -ne $script:Controls.PinnedServerCombo.SelectedItem -and (Get-ComboText -Combo $script:Controls.AssignmentCombo) -ne 'Pinned') {
        Select-ComboText -Combo $script:Controls.AssignmentCombo -Text 'Pinned'
        Update-PinnedServerControlState
    }
})
$script:Controls.PlanningGrid.Add_SelectionChanged({
    $row = $script:Controls.PlanningGrid.SelectedItem
    if ($null -eq $row) { return }
    $job = @($script:DraftJobs.Jobs | Where-Object Name -eq $row.Name)[0]
    $script:Controls.SelectedJobText.Text = [string]$job.Name
    $script:Controls.JobArgumentsBox.Text = [string]$row.Arguments
    $script:Controls.JobEnabledCheck.IsChecked = [bool]$job.Enabled
    Select-ComboText -Combo $script:Controls.ScheduleTypeCombo -Text ([string]$job.Schedule.Type)
    $script:Controls.TimesBox.Text = @($job.Schedule.Times) -join ', '
    Set-SelectedScheduleDays -Days @(Get-OrchestratorGuiPropertyValue -Object $job.Schedule -Name 'DaysOfWeek' -DefaultValue @())
    Update-ScheduleDaysControlState
    Select-ComboText -Combo $script:Controls.MissedPolicyCombo -Text ([string](Get-OrchestratorGuiPropertyValue -Object $job.Schedule -Name 'MissedRunPolicy' -DefaultValue 'RunOnce'))
    $assignmentMode = [string](Get-OrchestratorGuiPropertyValue -Object $job -Name 'AssignmentMode' -DefaultValue 'Legacy')
    Select-ComboText -Combo $script:Controls.AssignmentCombo -Text $assignmentMode
    Update-PinnedServerControlState
    $allowedServers = @(Get-OrchestratorGuiPropertyValue -Object $job -Name 'AllowedServers' -DefaultValue @())
    if ($assignmentMode -eq 'Pinned' -and $allowedServers.Count -eq 1) {
        Select-ComboText -Combo $script:Controls.PinnedServerCombo -Text ([string]$allowedServers[0])
    }
    $script:Controls.TimeoutBox.Text = [string](Get-OrchestratorGuiPropertyValue -Object $job -Name 'TimeoutMinutes' -DefaultValue 240)
    $script:Controls.RetriesBox.Text = [string](Get-OrchestratorGuiPropertyValue -Object $job -Name 'MaxRetries' -DefaultValue 0)
    $script:Controls.RetryDelayBox.Text = [string](Get-OrchestratorGuiPropertyValue -Object $job -Name 'RetryDelaySeconds' -DefaultValue 300)
    $script:Controls.DurationBox.Text = ([double](Get-OrchestratorGuiPropertyValue -Object $job -Name 'EstimatedDurationMinutes' -DefaultValue 5)).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $script:Controls.DependsOnBox.Text = @(Get-OrchestratorGuiPropertyValue -Object $job -Name 'DependsOn' -DefaultValue @()) -join ', '
    Select-ComboText -Combo $script:Controls.DependencyModeCombo -Text ([string](Get-OrchestratorGuiPropertyValue -Object $job -Name 'DependencyMode' -DefaultValue 'LatestOccurrence'))
    $script:Controls.DependencyMaxAgeBox.Text = [string](Get-OrchestratorGuiPropertyValue -Object $job -Name 'DependencyMaxAgeHours' -DefaultValue 0)
    Show-SelectedJobDependencies -Job $job
})
$script:Controls.ServersGrid.Add_SelectionChanged({
    $row = $script:Controls.ServersGrid.SelectedItem
    if ($null -eq $row -or $script:ServerSelectionRestoring) { return }
    $script:Controls.SelectedServerText.Text = [string]$row.Server
    $script:Controls.ServerActiveCheck.IsChecked = ([string]$row.State -eq 'Active')
    $script:Controls.ServerWeightBox.Text = ([double]$row.Weight).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    Select-ComboText -Combo $script:Controls.ServerPolicyCombo -Text ([string]$row.Policy)
    Update-ServerTestButtons -Server ([string]$row.Server)
})
$script:Controls.ApplyJobButton.Add_Click({ try { Apply-SelectedJobToDraft } catch { Write-GuiException -Context 'Invalid job draft' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Invalid job', 'OK', 'Error') | Out-Null } })
$script:Controls.TestServerButton.Add_Click({
    try { Invoke-GuiLaunch -Kind Manual -Name $script:ServerAcceptanceLaunchName }
    catch { Write-GuiException -Context 'Server test failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Server test failed', 'OK', 'Error') | Out-Null }
})
$script:Controls.TestServerAsServiceAccountButton.Add_Click({
    try { Invoke-GuiLaunch -Kind Manual -Name $script:ServerAcceptanceLaunchName -AsServiceAccount }
    catch { Write-GuiException -Context 'Server test with service account failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Server test with service account failed', 'OK', 'Error') | Out-Null }
})
$script:Controls.ApplyServerButton.Add_Click({ try { Apply-SelectedServerToDraft } catch { Write-GuiException -Context 'Invalid server draft' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Invalid server', 'OK', 'Error') | Out-Null } })
$script:Controls.RefreshButton.Add_Click({
    if ((Test-GuiDraftChanged) -and $script:OrchestratorDialog.Show("The draft has unpublished changes.`n`nRefresh reloads the published configuration and discards them. Continue?", 'Discard draft changes', 'YesNo', 'Warning') -ne 'Yes') { return }
    Refresh-AllViews
})
$script:Controls.RebalanceButton.Add_Click({ [void](Request-ElectionRebalance) })
$script:Controls.PublishButton.Add_Click({ Publish-Draft })
$script:Controls.HistoryRefreshButton.Add_Click({ Refresh-HistoryView })
$script:Controls.MainTabs.Add_SelectionChanged({
    if ($_.OriginalSource -ne $script:Controls.MainTabs -or -not $script:InitialHistoryLoaded) { return }
    if ($null -ne $script:Controls.MainTabs.SelectedItem -and
        [string]$script:Controls.MainTabs.SelectedItem.Header -eq 'History') {
        try { Refresh-HistoryView } catch { Write-GuiException -Context 'History refresh failed' -ErrorRecord $_ }
    }
    if (Test-GuiEventsTabVisible) { Start-GuiRefresh -Mode Live }
    if (Test-GuiDashboardTabVisible) { Start-GuiRefresh -Mode Live -ShowLoading }
    if (Test-GuiConfigurationTabVisible) { Start-GuiRefresh -Mode Live }
    elseif ([datetime]::UtcNow -lt $script:ConfigurationShowUntil) { $script:ConfigurationShowUntil = [datetime]::MinValue; Update-GuiConfigurationView }
})
$script:Controls.ConfigurationTabs.Add_SelectionChanged({
    # Settings sub-tab shown again: read the configuration files, as when Configuration is opened.
    if ($_.OriginalSource -ne $script:Controls.ConfigurationTabs -or -not $script:InitialHistoryLoaded) { return }
    if (Test-GuiConfigurationTabVisible) { Start-GuiRefresh -Mode Live }
    elseif ([datetime]::UtcNow -lt $script:ConfigurationShowUntil) { $script:ConfigurationShowUntil = [datetime]::MinValue; Update-GuiConfigurationView }
})
$script:Controls.ConfigFilterBox.Add_TextChanged({ try { Update-GuiConfigurationView } catch { Write-GuiException -Context 'Configuration filter failed' -ErrorRecord $_ } })
$script:Controls.ConfigShowSensitiveButton.Add_Click({ try { Show-GuiConfigurationSensitive } catch { Write-GuiException -Context 'Configuration display failed' -ErrorRecord $_ } })
$script:Controls.ConfigRefreshButton.Add_Click({ Start-GuiRefresh -Mode Live })
$script:Controls.ConfigOpenFolderButton.Add_Click({ try { Open-GuiConfigurationFolder } catch { Write-GuiException -Context 'Open configuration folder failed' -ErrorRecord $_ } })
$script:ConfigurationTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:ConfigurationTimer.Interval = [TimeSpan]::FromSeconds(1)
$script:ConfigurationTimer.Add_Tick({
    # Re-mask when the 30-second window ends; refreshes the countdown on the button meanwhile.
    if ($script:ConfigurationShowUntil -eq [datetime]::MinValue) { return }
    if ([datetime]::UtcNow -ge $script:ConfigurationShowUntil) { $script:ConfigurationShowUntil = [datetime]::MinValue }
    Update-GuiConfigurationView
})
$script:ConfigurationTimer.Start()
$script:Controls.DashboardTimelineCanvas.Add_SizeChanged({ param($sender) Update-OrchestratorDashboardTimeline -Canvas $sender })
$script:Controls.DashboardMetricsCanvas.Add_SizeChanged({ param($sender) Update-OrchestratorServerMetricsChart -Canvas $sender })
$script:Controls.ServersMetricsCanvas.Add_SizeChanged({ param($sender) Update-OrchestratorServerMetricsChart -Canvas $sender })
$script:Controls.DashboardTilesList.AddHandler([System.Windows.UIElement]::MouseLeftButtonUpEvent, [System.Windows.Input.MouseButtonEventHandler]{
    param($sender, $eventArgs)
    try {
        $tile = $eventArgs.OriginalSource.DataContext
        if ($null -ne $tile -and $tile.PSObject.Properties['Target']) { Invoke-GuiDashboardTarget -Target ([string]$tile.Target) }
    }
    catch { Write-GuiException -Context 'Dashboard navigation failed' -ErrorRecord $_ }
})
$script:Controls.DashboardActionsGrid.Add_MouseDoubleClick({
    try { $row = $script:Controls.DashboardActionsGrid.SelectedItem; if ($null -ne $row) { Invoke-GuiDashboardTarget -Target ([string]$row.Target) } }
    catch { Write-GuiException -Context 'Dashboard navigation failed' -ErrorRecord $_ }
})
foreach ($eventsFilter in @('EventsServerCombo', 'EventsLevelCombo', 'EventsCategoryCombo')) {
    $script:Controls[$eventsFilter].Add_SelectionChanged({ try { Update-GuiEventsView } catch { Write-GuiException -Context 'Event filter failed' -ErrorRecord $_ } })
}
$script:Controls.EventsTextBox.Add_TextChanged({ try { Update-GuiEventsView } catch { Write-GuiException -Context 'Event filter failed' -ErrorRecord $_ } })
$script:Controls.EventsGroupCheck.Add_Click({ try { Update-GuiEventsView } catch { Write-GuiException -Context 'Event filter failed' -ErrorRecord $_ } })
$script:Controls.EventsPeriodCombo.Add_SelectionChanged({ if ($script:InitialHistoryLoaded -and (Test-GuiEventsTabVisible)) { Start-GuiRefresh -Mode Live } })
$script:Controls.EventsRefreshButton.Add_Click({ Start-GuiRefresh -Mode Live })
$script:Controls.EventsGrid.Add_SelectionChanged({ Show-SelectedGuiEvent })
$script:Controls.HistoryGrid.Add_MouseDoubleClick({
    $row = $script:Controls.HistoryGrid.SelectedItem
    if ($null -eq $row) { return }
    if (-not [string]::IsNullOrWhiteSpace([string]$row.LogPath) -and (Test-Path -LiteralPath $row.LogPath)) { Start-Process -FilePath $row.LogPath; return }
    $logText = if ([string]::IsNullOrWhiteSpace([string]$row.LogPath)) { '<no log path recorded>' } else { [string]$row.LogPath }
    $script:OrchestratorDialog.Show("The log of this run is not reachable from this computer.`n`n$logText`n`nThe log stays on the server that ran the job ($($row.Server)) until it is copied to LOG-ALL.", 'Log not found', 'OK', 'Information') | Out-Null
})
$script:Controls.Failures24hButton.Add_Click({ try { Show-FailuresLast24Hours } catch { Write-GuiException -Context 'Failure filter failed' -ErrorRecord $_ } })
$script:Controls.EnableMaintenanceButton.Add_Click({
    try { Set-GuiMaintenance -Enabled $true }
    catch { Write-GuiException -Context 'Enable maintenance failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Maintenance', 'OK', 'Error') | Out-Null; Refresh-MaintenanceView }
})
$script:Controls.DisableMaintenanceButton.Add_Click({
    try { Set-GuiMaintenance -Enabled $false }
    catch { Write-GuiException -Context 'Disable maintenance failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Maintenance', 'OK', 'Error') | Out-Null; Refresh-MaintenanceView }
})
$script:Controls.OperationsRunningGrid.Add_SelectionChanged({
    $row = $script:Controls.OperationsRunningGrid.SelectedItem
    $script:Controls.StopRunningJobButton.IsEnabled = $null -ne $row -and [bool]$row.StopSupported
})
$script:Controls.StopRunningJobButton.Add_Click({
    try { Request-SelectedRunningJobStop }
    catch { Write-GuiException -Context 'Job stop request failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Stop request failed', 'OK', 'Error') | Out-Null }
})
$script:Controls.LaunchJobsGrid.Add_SelectionChanged({ try { Update-SelectedLaunchPreview -Kind Job } catch { Write-GuiException -Context 'Job preview failed' -ErrorRecord $_ } })
$script:Controls.ManualLaunchesGrid.Add_SelectionChanged({ try { Update-SelectedLaunchPreview -Kind Manual } catch { Write-GuiException -Context 'Launch preview failed' -ErrorRecord $_ } })
$script:Controls.RunJobNowButton.Add_Click({
    try { Invoke-GuiLaunch -Kind Job }
    catch { Write-GuiException -Context 'Run now failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Run now failed', 'OK', 'Error') | Out-Null }
})
$script:Controls.RunJobAsServiceAccountButton.Add_Click({
    try { Invoke-GuiLaunch -Kind Job -AsServiceAccount }
    catch { Write-GuiException -Context 'Run now with service account failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Run now with service account failed', 'OK', 'Error') | Out-Null }
})
$script:Controls.RunManualLaunchButton.Add_Click({
    try { Invoke-GuiLaunch -Kind Manual }
    catch { Write-GuiException -Context 'Manual launch failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Manual launch failed', 'OK', 'Error') | Out-Null }
})
$script:Controls.RequestJobButton.Add_Click({
    try {
        $row = $script:Controls.LaunchJobsGrid.SelectedItem
        if ($null -eq $row) { throw 'Select a job first.' }
        [void](Request-SelectedJobRun -JobName ([string]$row.Name) -IncludeDependencies ([bool]$script:Controls.LaunchIncludeDependenciesCheck.IsChecked))
    }
    catch { Write-GuiException -Context 'Run request failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Run request failed', 'OK', 'Error') | Out-Null }
})
$script:Controls.ApplyJobExecutionModeButton.Add_Click({
    try {
        $row = $script:Controls.LaunchJobsGrid.SelectedItem
        if ($null -eq $row) { throw 'Select an on-premises job first.' }
        [void](Set-GuiJobExecutionMode -JobName ([string]$row.Name) -Mode (Get-ComboText -Combo $script:Controls.JobExecutionModeCombo))
    }
    catch { Write-GuiException -Context 'Execution mode change failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Execution mode', 'OK', 'Error') | Out-Null }
})
$script:Controls.ApplyModeToAllOnPremButton.Add_Click({
    try {
        $mode = Get-ComboText -Combo $script:Controls.JobExecutionModeCombo
        if ([string]::IsNullOrWhiteSpace($mode)) { $mode = '(default)' }
        $names = @($script:LaunchJobRows | Where-Object { $_.OnPremises -and $_.InManifest } | ForEach-Object { [string]$_.Name })
        if ($names.Count -eq 0) { throw 'No on-premises job in the published configuration.' }
        if ($script:OrchestratorDialog.Show(("Set the execution mode {0} in the draft for {1} on-premises job(s)?`n`n{2}`n`nPublish the changes to apply." -f $mode, $names.Count, ($names -join [Environment]::NewLine)), 'Execution mode', 'YesNo', 'Question') -ne 'Yes') { return }
        [void](Set-GuiJobExecutionMode -JobName $names -Mode $mode)
    }
    catch { Write-GuiException -Context 'Execution mode change failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Execution mode', 'OK', 'Error') | Out-Null }
})
$script:Controls.OrchestratorStartButton.Add_Click({ try { Invoke-GuiOrchestratorAction -Action Start } catch { Write-GuiException -Context 'Orchestrator action Start failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Orchestrator action failed', 'OK', 'Error') | Out-Null } })
$script:Controls.OrchestratorStopButton.Add_Click({ try { Invoke-GuiOrchestratorAction -Action Stop } catch { Write-GuiException -Context 'Orchestrator action Stop failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Orchestrator action failed', 'OK', 'Error') | Out-Null } })
$script:Controls.OrchestratorRestartButton.Add_Click({ try { Invoke-GuiOrchestratorAction -Action Restart } catch { Write-GuiException -Context 'Orchestrator action Restart failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Orchestrator action failed', 'OK', 'Error') | Out-Null } })
$script:Controls.FullCollectButton.Add_Click({ try { Invoke-GuiOrchestratorAction -Action FullCollect } catch { Write-GuiException -Context 'Orchestrator action FullCollect failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Orchestrator action failed', 'OK', 'Error') | Out-Null } })
$script:Controls.SendSummaryButton.Add_Click({ try { Invoke-GuiOrchestratorAction -Action SendSummary } catch { Write-GuiException -Context 'Orchestrator action SendSummary failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Orchestrator action failed', 'OK', 'Error') | Out-Null } })
$script:Controls.RequestPrepareButton.Add_Click({
    try { [void](Request-SelectedJobRun -JobName $script:PrepareJobName -IncludeDependencies $false) }
    catch { Write-GuiException -Context 'Prepare request failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Prepare request failed', 'OK', 'Error') | Out-Null }
})
$script:Controls.RequestRunButton.Add_Click({
    try { [void](Request-SelectedJobRun) }
    catch { Write-GuiException -Context 'Run request failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Run request failed', 'OK', 'Error') | Out-Null }
})
$script:Controls.RequestsGrid.Add_SelectionChanged({
    try { Update-SelectedPipelineRequest }
    catch { Write-GuiException -Context 'Request selection failed' -ErrorRecord $_ }
})
$script:Controls.CancelRequestButton.Add_Click({
    try { Stop-SelectedPipelineRequest }
    catch { Show-GuiCancellationError -Message $_.Exception.Message }
})
$script:Controls.CancelAllRequestsButton.Add_Click({
    try { Stop-AllPipelineRequests }
    catch { Show-GuiCancellationError -Message $_.Exception.Message }
})
$script:Controls.OperationsMailsGrid.Add_MouseDoubleClick({
    $row = $script:Controls.OperationsMailsGrid.SelectedItem
    if ($null -ne $row -and (Test-Path -LiteralPath $row.Path)) { Start-Process -FilePath $row.Path }
})
$script:Controls.AddServerButton.Add_Click({
    try {
        $added = Publish-GuiServerAddition -Server $script:Controls.NewServerBox.Text
        if ($null -eq $added) { return }
        $script:Controls.NewServerBox.Clear()
        $script:OrchestratorDialog.Show(("Server '{0}' was added Inactive and published (version {1}).{2}{2}Install its orchestrator task, check it in Operations > Live, then select it here, tick Active (takes jobs), apply and publish." -f $added.Server, $added.VersionId, [Environment]::NewLine), 'Server added', 'OK', 'Information') | Out-Null
    }
    catch { Write-GuiException -Context 'Add server failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Add server', 'OK', 'Error') | Out-Null }
})
$script:Controls.RemoveServerButton.Add_Click({
    $row = $script:Controls.ServersGrid.SelectedItem
    if ($null -eq $row) { return }
    $server = [string]$row.Server
    $pinnedJobs = @($script:DraftJobs.Jobs | Where-Object { $_.AssignmentMode -eq 'Pinned' -and $server -in @($_.AllowedServers) })
    if ($pinnedJobs.Count -gt 0) {
        $script:OrchestratorDialog.Show("Server is still used by pinned jobs: $($pinnedJobs.Name -join ', ')", 'Cannot remove server', 'OK', 'Error') | Out-Null
        return
    }
    if ($script:OrchestratorDialog.Show("Remove $server from the cluster draft?", 'Remove server', 'YesNo', 'Warning') -ne 'Yes') { return }
    $script:DraftCluster.ExpectedOrchestratorServers = @($script:DraftCluster.ExpectedOrchestratorServers | Where-Object { $_ -ine $server })
    [void]$script:DraftCluster.ElectionWeightsByServer.PSObject.Properties.Remove($server)
    [void]$script:DraftCluster.ServerJobPolicies.PSObject.Properties.Remove($server)
    if ($script:DraftCluster.PSObject.Properties['ServerStates'] -and $null -ne $script:DraftCluster.ServerStates) { [void]$script:DraftCluster.ServerStates.PSObject.Properties.Remove($server) }
    Refresh-ServersView | Out-Null
    Write-GuiActivity -Message "Server '$server' removed from the draft."
})
$script:Controls.RollbackButton.Add_Click({
    $version = $script:Controls.VersionsGrid.SelectedItem
    if ($null -eq $version) { return }
    if ($script:OrchestratorDialog.Show("Rollback to the configuration before $($version.VersionId)?`nThis creates a new audited version.", 'Rollback', 'YesNo', 'Warning') -ne 'Yes') { return }
    try {
        $result = Restore-SmartM365OrchestratorConfigurationVersion -SharedDataFolderPath $script:SharedDataFolderPath -VersionFolderPath $version.FolderPath -Snapshot Before -ExpectedJobsHash $script:Snapshot.JobsHash -ExpectedClusterHash $script:Snapshot.ClusterHash
        Write-GuiActivity -Message "Rollback published as $($result.VersionId)." -Level SUCCESS
        Refresh-AllViews
    }
    catch { Write-GuiException -Context 'Rollback failed' -ErrorRecord $_; $script:OrchestratorDialog.Show($_.Exception.Message, 'Rollback failed', 'OK', 'Error') | Out-Null }
})
$script:Controls.ExportCsvButton.Add_Click({
    $dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $dialog.Filter = 'CSV files (*.csv)|*.csv'
    $dialog.FileName = 'SmartM365-Orchestrator-History-{0}.csv' -f (Get-Date).ToString('yyyyMMdd-HHmmss')
    if ($dialog.ShowDialog()) {
        $script:HistoryRows | Select-Object -Property $script:HistoryExportColumns | Export-Csv -LiteralPath $dialog.FileName -NoTypeInformation -Encoding utf8
        Write-GuiActivity -Message "History exported to $($dialog.FileName)." -Level SUCCESS
    }
})
$script:Controls.ExportHtmlButton.Add_Click({
    $dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $dialog.Filter = 'HTML files (*.html)|*.html'
    $dialog.FileName = 'SmartM365-Orchestrator-History-{0}.html' -f (Get-Date).ToString('yyyyMMdd-HHmmss')
    if ($dialog.ShowDialog()) {
        $body = $script:HistoryRows | Select-Object -Property $script:HistoryExportColumns | ConvertTo-Html -Title 'SmartM365 Orchestrator history' -PreContent '<h1>SmartM365 Orchestrator history</h1>'
        [System.IO.File]::WriteAllText($dialog.FileName, ($body -join [Environment]::NewLine), [System.Text.UTF8Encoding]::new($false))
        Write-GuiActivity -Message "History exported to $($dialog.FileName)." -Level SUCCESS
    }
})

if ($SmokeTest) {
    Refresh-AllViewsSync
    $script:Controls.HistoryServerCombo.ItemsSource = @('All') + @($script:DraftCluster.ExpectedOrchestratorServers)
    $script:Controls.HistoryServerCombo.SelectedIndex = 0
    $script:Controls.HistoryJobCombo.ItemsSource = @('All') + @($script:DraftJobs.Jobs.Name | Sort-Object)
    $script:Controls.HistoryJobCombo.SelectedIndex = 0
}
else { Refresh-AllViews }
Write-GuiActivity -Message ('Startup initial refresh dispatched after {0} ms.' -f $script:StartupClock.ElapsedMilliseconds)
if ($SmokeTest) {
    if (@($script:LaunchJobRows).Count -eq 0 -or @($script:ManualLaunchRows).Count -eq 0) { throw 'Operations launch lists are empty.' }
    $script:Controls.ManualLaunchesGrid.SelectedItem = $script:ManualLaunchRows[0]
    if (-not $script:Controls.RunManualLaunchButton.IsEnabled -or [string]::IsNullOrWhiteSpace($script:Controls.ManualLaunchCommandBox.Text)) { throw 'Manual launch preview is not displayed.' }
    if (@($script:ManualLaunchRows | Where-Object { [string]$_.Name -eq $script:ServerAcceptanceLaunchName }).Count -ne 1) { throw 'Server acceptance test is not in Manual launches.' }
    if ($null -eq (Get-GuiLaunchPlanText -Name $script:ServerAcceptanceLaunchName).Plan) { throw 'Server acceptance test command is not resolved.' }
    Update-ServerTestButtons -Server $env:COMPUTERNAME
    if (-not $script:Controls.TestServerButton.IsEnabled) { throw 'Server test button is not enabled for this computer.' }
    Update-ServerTestButtons -Server 'SMOKE-OTHER-SERVER'
    if ($script:Controls.TestServerButton.IsEnabled) { throw 'Server test button is enabled for another computer.' }
    foreach ($action in @('Start', 'Stop', 'Restart', 'FullCollect', 'SendSummary')) { if (-not (Test-Path -LiteralPath (Get-GuiOrchestratorActionCommand -Action $action).ScriptPath)) { throw "Orchestrator action script missing: $action" } }
    $planningNames = @($script:PlanningRows | ForEach-Object { [string]$_.Name })
    $sortedPlanningNames = @($planningNames | Sort-Object)
    if (($planningNames -join "`0") -cne ($sortedPlanningNames -join "`0")) { throw 'Planning rows are not sorted by Job ascending.' }
    $jobColumn = @($script:Controls.PlanningGrid.Columns | Where-Object { [string]$_.Header -eq 'Job' })[0]
    if ($null -eq $jobColumn -or $jobColumn.SortDirection -ne [System.ComponentModel.ListSortDirection]::Ascending) { throw 'Planning Job column does not display the default ascending sort.' }

    $smokeServer = @($script:DraftCluster.ExpectedOrchestratorServers | Select-Object -First 1)[0]
    if ([string]::IsNullOrWhiteSpace([string]$smokeServer)) {
        $smokeServer = 'SMOKE-SERVER'
        Set-JsonProperty -Object $script:DraftCluster -Name ExpectedOrchestratorServers -Value @($smokeServer)
        [void](Refresh-ServersView)
    }
    $script:Controls.PlanningGrid.SelectedItem = $script:PlanningRows[0]
    $argumentsColumn = @($script:Controls.PlanningGrid.Columns | Where-Object { [string]$_.Header -eq 'Arguments' })[0]
    if ($null -eq $argumentsColumn -or [string]$argumentsColumn.Binding.Path.Path -ne 'Arguments') { throw 'Planning arguments column is not bound to the job arguments.' }
    if (-not $script:Controls.JobArgumentsBox.IsReadOnly -or [string]$script:Controls.JobArgumentsBox.Text -cne [string]$script:PlanningRows[0].Arguments) {
        throw 'Selected job arguments are not displayed read-only.'
    }
    $rowWithArguments = @($script:PlanningRows | Where-Object { [string]$_.Arguments -ne '(none)' } | Select-Object -First 1)
    if ($rowWithArguments.Count) {
        $script:Controls.PlanningGrid.SelectedItem = $rowWithArguments[0]
        if ([string]$script:Controls.JobArgumentsBox.Text -cne [string]$rowWithArguments[0].Arguments) {
            throw 'Selected job arguments were not displayed in full.'
        }
        $script:Controls.PlanningGrid.SelectedItem = $script:PlanningRows[0]
    }
    Select-ComboText -Combo $script:Controls.AssignmentCombo -Text 'Elected'
    Update-PinnedServerControlState
    if ($script:Controls.PinnedServerCombo.IsEnabled) { throw 'Pinned server selector stayed enabled for Elected assignment.' }

    Select-ComboText -Combo $script:Controls.ScheduleTypeCombo -Text 'Weekly'
    Update-ScheduleDaysControlState
    if (-not $script:Controls.DaysPanel.IsEnabled) { throw 'Weekday checkboxes did not enable for a Weekly schedule.' }
    $missingWeeklyDayRejected = $false
    try { Apply-SelectedJobToDraft } catch { $missingWeeklyDayRejected = $_.Exception.Message -like '*Select at least one day*' }
    if (-not $missingWeeklyDayRejected) { throw 'A Weekly schedule without a selected day was not rejected.' }
    Set-SelectedScheduleDays -Days @('Tuesday', 'Thursday')
    Apply-SelectedJobToDraft
    $smokeJob = @($script:DraftJobs.Jobs | Where-Object Name -eq $script:PlanningRows[0].Name)[0]
    if ([string]$smokeJob.Schedule.Type -ne 'Weekly' -or (@($smokeJob.Schedule.DaysOfWeek) -join ',') -cne 'Tuesday,Thursday') {
        throw 'Weekday checkbox selections did not round-trip to the Weekly schedule.'
    }
    Select-ComboText -Combo $script:Controls.ScheduleTypeCombo -Text 'Daily'
    Update-ScheduleDaysControlState
    if ($script:Controls.DaysPanel.IsEnabled -or @(Get-SelectedScheduleDays).Count -ne 0) { throw 'Weekday checkboxes were not disabled and cleared for a Daily schedule.' }
    Apply-SelectedJobToDraft
    $smokeJob = @($script:DraftJobs.Jobs | Where-Object Name -eq $script:PlanningRows[0].Name)[0]
    if ([string]$smokeJob.Schedule.Type -ne 'Daily' -or @($smokeJob.Schedule.DaysOfWeek).Count -ne 0) {
        throw ("Daily schedule retained weekday selections. Type={0}; Days={1}" -f [string]$smokeJob.Schedule.Type, (@($smokeJob.Schedule.DaysOfWeek) -join ','))
    }
    Write-GuiActivity -Message 'SMOKE_TEST_SCHEDULE_DAYS_OK Weekly=Tuesday,Thursday; DailyDays=0' -Level SUCCESS

    Select-ComboText -Combo $script:Controls.AssignmentCombo -Text 'Pinned'
    Update-PinnedServerControlState
    if (-not $script:Controls.PinnedServerCombo.IsEnabled) { throw 'Pinned server selector did not enable for Pinned assignment.' }
    $missingPinnedServerRejected = $false
    try { Apply-SelectedJobToDraft } catch { $missingPinnedServerRejected = $_.Exception.Message -like '*Select exactly one pinned server*' }
    if (-not $missingPinnedServerRejected) { throw 'Pinned assignment without a server was not rejected.' }
    Select-ComboText -Combo $script:Controls.PinnedServerCombo -Text ([string]$smokeServer)
    Apply-SelectedJobToDraft
    $smokeJob = @($script:DraftJobs.Jobs | Where-Object Name -eq $script:PlanningRows[0].Name)[0]
    if ([string]$smokeJob.AssignmentMode -ne 'Pinned' -or @($smokeJob.AllowedServers).Count -ne 1 -or [string]$smokeJob.AllowedServers[0] -ine [string]$smokeServer) {
        throw 'Pinned assignment smoke test did not preserve exactly one expected server.'
    }
    $dependencyTarget = @($script:PlanningRows | Where-Object { $_.Name -ne $script:PlanningRows[0].Name } | Select-Object -First 1)[0]
    $script:Controls.PlanningGrid.SelectedItem = $script:PlanningRows[0]
    $script:Controls.DependsOnBox.Text = [string]$dependencyTarget.Name
    Select-ComboText -Combo $script:Controls.DependencyModeCombo -Text 'FreshSuccess'
    $script:Controls.DependencyMaxAgeBox.Text = '30'
    Apply-SelectedJobToDraft -SkipDependentsConfirmation
    $smokeJob = @($script:DraftJobs.Jobs | Where-Object Name -eq $script:PlanningRows[0].Name)[0]
    if ((@($smokeJob.DependsOn) -join ',') -cne [string]$dependencyTarget.Name -or [string]$smokeJob.DependencyMode -ne 'FreshSuccess' -or [int]$smokeJob.DependencyMaxAgeHours -ne 30) { throw 'Dependency fields did not round-trip to the draft.' }
    if (@(Get-SmartM365OrchestratorDependents -JobsDocument $script:DraftJobs -JobName ([string]$dependencyTarget.Name)) -notcontains [string]$smokeJob.Name) { throw 'Dependents lookup did not find the new dependency.' }
    $readiness = @(Get-SmartM365OrchestratorDependencyReadiness -SharedDataFolderPath $script:SharedDataFolderPath -JobsDocument $script:DraftJobs -JobName ([string]$smokeJob.Name) -Runs $script:RecentRuns)
    if ($readiness.Count -ne 1 -or [string]$readiness[0].Dependency -ne [string]$dependencyTarget.Name) { throw 'Dependency readiness did not return the configured dependency.' }
    $script:Controls.DependsOnBox.Text = 'Unknown-Smoke-Dependency'
    $unknownDependencyRejected = $false
    try { Apply-SelectedJobToDraft -SkipDependentsConfirmation } catch { $unknownDependencyRejected = $_.Exception.Message -like '*unknown dependency*' }
    $smokeJob = @($script:DraftJobs.Jobs | Where-Object Name -eq $script:PlanningRows[0].Name)[0]
    if (-not $unknownDependencyRejected -or (@($smokeJob.DependsOn) -join ',') -cne [string]$dependencyTarget.Name) { throw 'An unknown dependency was not rejected or the draft was not restored.' }
    $script:Controls.DependsOnBox.Text = ''
    Apply-SelectedJobToDraft -SkipDependentsConfirmation
    Write-GuiActivity -Message ("SMOKE_TEST_DEPENDENCIES_OK Dependency={0}; Readiness={1}" -f $dependencyTarget.Name, $readiness[0].State) -Level SUCCESS

    $operations = Refresh-OperationsView
    if ($null -eq $operations -or @($operations.Servers).Count -ne @($script:DraftCluster.ExpectedOrchestratorServers).Count) { throw 'Operations view did not return one row per expected server.' }
    $smokeOwners = Get-ElectionOwners
    $activeRequests = @($script:RequestRows | Where-Object Status -in @('Running', 'Queued', 'Cancelling'))
    $publishedRunnable = @($script:Snapshot.Jobs.Jobs | Where-Object {
        $smokeMode = [string](Get-OrchestratorGuiPropertyValue -Object $_ -Name 'AssignmentMode' -DefaultValue 'Legacy')
        [bool]$_.Enabled -and @(Get-OrchestratorGuiPropertyValue -Object $_ -Name 'DependsOn' -DefaultValue @()).Count -eq 0 -and
        ($smokeMode -in @('Legacy', 'Pinned') -or ($smokeMode -eq 'Elected' -and $smokeOwners.ContainsKey([string]$_.Name)))
    } | Select-Object -First 1)
    if ($activeRequests.Count) {
        # A smoke test never adds work to a live queue.
        Write-GuiActivity -Message ("SMOKE_TEST_RUN_REQUEST_SKIPPED Active or queued request(s): {0}" -f (@($activeRequests | ForEach-Object BatchId) -join ', ')) -Level WARN
    }
    elseif ($publishedRunnable.Count) {
        $script:Controls.PlanningGrid.SelectedItem = @($script:PlanningRows | Where-Object Name -eq $publishedRunnable[0].Name)[0]
        $script:Controls.IncludeDependenciesCheck.IsChecked = $false
        $runRequest = Request-SelectedJobRun -SkipConfirmation
        if (-not @($script:RequestRows | Where-Object BatchId -eq $runRequest.BatchId).Count) { throw 'The run request is missing from the Requests view.' }
        # The same job again is an identical request: refused, whether queueing is available or not.
        $duplicateRejected = $false
        try { [void](Request-SelectedJobRun -SkipConfirmation) } catch { $duplicateRejected = $_.Exception.Message -like '*identical pipeline request*' -or $_.Exception.Message -like '*active pipeline request already exists*' }
        if (-not $duplicateRejected) { throw 'An identical run request was accepted while the first one is active.' }
        Write-GuiActivity -Message ("SMOKE_TEST_RUN_REQUEST_OK Job={0}; BatchId={1}" -f $publishedRunnable[0].Name, $runRequest.BatchId) -Level SUCCESS
    }
    else { Write-GuiActivity -Message 'SMOKE_TEST_RUN_REQUEST_SKIPPED No enabled runnable job without dependency in the published configuration.' -Level WARN }
    Write-GuiActivity -Message ("SMOKE_TEST_OPERATIONS_OK Servers={0}; HealthRows={1}" -f @($operations.Servers).Count, $script:HealthByName.Count) -Level SUCCESS
    Set-GuiEventsServers
    Set-GuiLogEvents -Result (Get-SmartM365OrchestratorLogEvents -LogFolderPath $script:MailFolderPath -Server @($script:DraftCluster.ExpectedOrchestratorServers) -Hours 24 -Cache $script:LogEventsCache)
    if ($null -eq $script:DashboardData -or @($script:Controls.DashboardTilesList.ItemsSource).Count -ne 9 -or @($script:Controls.DashboardGrid.ItemsSource).Count -ne @($script:DraftCluster.ExpectedOrchestratorServers).Count) { throw 'Dashboard was not built from the shared data.' }
    Write-GuiActivity -Message ("SMOKE_TEST_DASHBOARD_OK Banner={0}; Tiles={1}; Actions={2}; TimelineItems={3}" -f $script:DashboardData.Banner.Level, @($script:DashboardData.Tiles).Count, @($script:DashboardData.Actions).Count, @($script:DashboardData.Timeline.Items).Count) -Level SUCCESS
    Set-GuiConfiguration -View (Get-SmartM365OrchestratorConfigurationView -RootPath $script:SmartM365RootPath -ProfileKey ([string]$script:EffectiveConfig.ProfileKey))
    $smokeConfigRows = @(@($script:Controls.ConfigGlobalGrid.ItemsSource) + @($script:Controls.ConfigTenantGrid.ItemsSource) + @($script:Controls.ConfigMailGrid.ItemsSource))
    if (@($script:Controls.ConfigGlobalGrid.ItemsSource).Count -eq 0 -or @($smokeConfigRows | Where-Object { $_.Sensitive -and -not $_.Masked -and $_.Model -ne 'Missing from file' }).Count) { throw 'Configuration view is empty or shows a sensitive value by default.' }
    Write-GuiActivity -Message ("SMOKE_TEST_CONFIGURATION_OK Files={0}; Rows={1}; Hidden={2}" -f @($script:ConfigurationView.Files | Where-Object Status -eq 'OK').Count, $smokeConfigRows.Count, @($smokeConfigRows | Where-Object Masked).Count) -Level SUCCESS
    Write-GuiActivity -Message ("SMOKE_TEST_EVENTS_OK Events={0}; Rows={1}; Logs={2}" -f @($script:LogEventsResult.Events).Count, @($script:Controls.EventsGrid.ItemsSource).Count, (@($script:LogEventsResult.Sources | ForEach-Object { '{0}={1}' -f $_.Server, $_.Status }) -join ',')) -Level SUCCESS

    $rebalanceRequest = Request-ElectionRebalance -SkipConfirmation
    $rebalanceRequestPath = Join-Path -Path $script:SharedDataFolderPath -ChildPath 'Election\Orchestrator-RebalanceRequest.json'
    $savedRebalanceRequest = Read-SmartM365OrchestratorJson -Path $rebalanceRequestPath
    if ($null -eq $script:Controls.RebalanceButton -or [string]$script:Controls.RebalanceButton.Content -ne 'Rebalance now') { throw 'Rebalance now button is missing.' }
    if ([string]::IsNullOrWhiteSpace([string]$rebalanceRequest.RequestId) -or [string]$savedRebalanceRequest.RequestId -cne [string]$rebalanceRequest.RequestId) {
        throw 'Rebalance request smoke test did not persist the submitted request atomically.'
    }
    Write-GuiActivity -Message ("SMOKE_TEST_REBALANCE_OK RequestId={0}" -f $rebalanceRequest.RequestId) -Level SUCCESS
    Write-GuiActivity -Message ("SMOKE_TEST_PINNED_OK Job={0}; Server={1}; DefaultSort=JobAscending" -f $smokeJob.Name, $smokeServer) -Level SUCCESS
    $smokeNewServer = 'SMOKE-NEW-SERVER'
    [void](Add-ServerToDraft -Server $smokeNewServer)
    $smokeNewRow = @(Refresh-ServersView | Where-Object { [string]$_.Server -eq $smokeNewServer })
    if ($smokeNewRow.Count -ne 1 -or [string]$smokeNewRow[0].State -ne 'Inactive') { throw 'A server added in the GUI must be Inactive.' }
    $smokeActiveRows = @(Refresh-ServersView | Where-Object { [string]$_.Server -ne $smokeNewServer -and [string]$_.State -eq 'Active' })
    if (-not (Test-SmartM365OrchestratorClusterDocument -Document $script:DraftCluster).Valid) { throw 'Draft with an Inactive server is not valid.' }
    $script:DraftCluster.ExpectedOrchestratorServers = @($script:DraftCluster.ExpectedOrchestratorServers | Where-Object { $_ -ne $smokeNewServer })
    [void]$script:DraftCluster.ElectionWeightsByServer.PSObject.Properties.Remove($smokeNewServer)
    [void]$script:DraftCluster.ServerStates.PSObject.Properties.Remove($smokeNewServer)
    [void](Refresh-ServersView)
    Write-GuiActivity -Message ("SMOKE_TEST_SERVER_STATE_OK NewServer=Inactive; ActiveServers={0}" -f $smokeActiveRows.Count) -Level SUCCESS
    if ($script:SmokePublishAllowed) {
        $smokeAdded = Publish-GuiServerAddition -Server 'SMOKE-PUBLISHED-SERVER' -SkipConfirmation
        $smokePublishedCluster = Read-SmartM365OrchestratorJson -Path (Get-SmartM365OrchestratorConfigurationPaths -SharedDataFolderPath $script:SharedDataFolderPath).ClusterPath
        if ([string]$smokePublishedCluster.ServerStates.'SMOKE-PUBLISHED-SERVER' -ne 'Inactive' -or 'SMOKE-PUBLISHED-SERVER' -notin @($script:DraftCluster.ExpectedOrchestratorServers)) { throw 'Add server must publish the server Inactive and keep it in the draft.' }
        Write-GuiActivity -Message ("SMOKE_TEST_ADD_SERVER_PUBLISHED_OK Version={0}" -f $smokeAdded.VersionId) -Level SUCCESS
    }
    else { Write-GuiActivity -Message 'SMOKE_TEST_ADD_SERVER_PUBLISHED_SKIPPED No explicit -SharedDataFolderPath (never published to the real share).' -Level WARN }
    Write-GuiActivity -Message ("SMOKE_TEST_OK SmartM365 Orchestrator GUI v{0} | Jobs={1} | PlanningRows={2}" -f $script:AppVersion, @($script:DraftJobs.Jobs).Count, @($script:PlanningRows).Count) -Level SUCCESS
    $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
}

$script:CancellationTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:CancellationTimer.Interval = [TimeSpan]::FromMilliseconds(150)
$script:CancellationTimer.Add_Tick({ Receive-GuiPipelineCancellation })
$script:CancellationTimer.Start()
$script:RequestSubmitTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:RequestSubmitTimer.Interval = [TimeSpan]::FromMilliseconds(150)
$script:RequestSubmitTimer.Add_Tick({ try { Receive-GuiJobRunSubmission } catch { Write-GuiException -Context 'Run request result handling failed' -ErrorRecord $_ } })
$script:RequestSubmitTimer.Start()
$script:SendSummaryTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:SendSummaryTimer.Interval = [TimeSpan]::FromMilliseconds(500)
$script:SendSummaryTimer.Add_Tick({ try { Receive-GuiSendSummary } catch { Write-GuiException -Context 'Send summary result handling failed' -ErrorRecord $_ } })
$script:SendSummaryTimer.Start()
$script:RefreshTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:RefreshTimer.Interval = [TimeSpan]::FromMilliseconds(150)
$script:RefreshTimer.Add_Tick({ Receive-GuiRefresh; Update-GuiDashboardLoading })
$script:RefreshTimer.Start()
$script:AutoRefreshTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:AutoRefreshTimer.Interval = [TimeSpan]::FromSeconds(60)
$script:AutoRefreshTimer.Add_Tick({
    if (-not [bool]$script:Controls.AutoRefreshCheck.IsChecked) { return }
    try { Invoke-AutoRefresh } catch { Write-GuiException -Context 'Auto-refresh failed' -ErrorRecord $_ }
})
$script:AutoRefreshTimer.Start()
$window.Add_ContentRendered({
    if ($script:GuiSplash) {
        Hide-SmartM365GuiSplash -Splash $script:GuiSplash
        $window.Activate() | Out-Null
    }
    Write-GuiActivity -Message ('Startup window ready after {0} ms.' -f $script:StartupClock.ElapsedMilliseconds)
})
$window.Add_Closed({
    if ($script:AutoRefreshTimer) { $script:AutoRefreshTimer.Stop() }
    if ($script:ConfigurationTimer) { $script:ConfigurationTimer.Stop() }
    if ($script:RefreshTimer) { $script:RefreshTimer.Stop() }
    if ($script:CancellationTimer) { $script:CancellationTimer.Stop() }
    if ($script:RequestSubmitTimer) { $script:RequestSubmitTimer.Stop() }
    Write-GuiActivity -Message 'GUI session closed.'
    if ($script:GuiSplash) { Close-SmartM365GuiSplash -Splash $script:GuiSplash }
})
$window.Add_Closing({
    param($sender, $eventArgs)
    if ($script:CancellationBusy) {
        $eventArgs.Cancel = $true
        $script:OrchestratorDialog.Show('Cancellation is still being checked or published. Wait for the result before closing this window. Running collectors are not stopped.', 'Operation in progress', 'OK', 'Information') | Out-Null
    }
    elseif ($script:RequestSubmitBusy) {
        $eventArgs.Cancel = $true
        $script:OrchestratorDialog.Show('A run request is still being submitted. Wait for the result before closing this window.', 'Operation in progress', 'OK', 'Information') | Out-Null
    }
})
[void]$window.ShowDialog()
    $smartM365RunGuard.ReachedTerminal=$true
    $smartM365RunGuard.ReturnedNormally=$true
} catch {
    if ($smartM365RunGuard -and $_.Exception -isnot [Management.Automation.PipelineStoppedException] -and $_.Exception -isnot [OperationCanceledException]) { $smartM365RunGuard.ErrorRecord=$_ }
    # Started by SmartInventory.cmd with a hidden console: the error would otherwise be invisible.
    if ($env:SMARTM365_GUI_DETACHED -eq '1' -and $_.Exception -isnot [Management.Automation.PipelineStoppedException]) {
        try {
            Add-Type -AssemblyName PresentationFramework -ErrorAction Stop
            $script:OrchestratorDialog.Show(("The orchestrator GUI could not start.`n`n{0}" -f $_.Exception.Message), 'SmartM365 Orchestrator GUI', 'OK', 'Error') | Out-Null
        }
        catch { Write-Verbose ("Startup error message box unavailable: {0}" -f $_.Exception.Message) }
    }
    throw
} finally {
    if ($smartM365RunGuard) {
    Complete-SmartM365InventoryRun -State $smartM365RunGuard -Cleanup {
        foreach($name in @('AutoRefreshTimer','RefreshTimer','CancellationTimer','RequestSubmitTimer','ConfigurationTimer','SendSummaryTimer')){
            $timer=Get-Variable $name -Scope Script -ValueOnly -ErrorAction SilentlyContinue
            if($timer){Invoke-SmartM365RunCleanup $smartM365RunGuard $name { $timer.Stop() }}
        }
        foreach($name in @('RefreshContext','CancellationContext','RequestSubmitContext')){
            $context=Get-Variable $name -Scope Script -ValueOnly -ErrorAction SilentlyContinue
            if($context -and (Get-Command Close-SmartM365OrchestratorGuiWorker -ErrorAction SilentlyContinue)){
                Invoke-SmartM365RunCleanup $smartM365RunGuard $name { Close-SmartM365OrchestratorGuiWorker -Context $context }
            }
        }
        $splash=Get-Variable GuiSplash -Scope Script -ValueOnly -ErrorAction SilentlyContinue
        if($splash){Invoke-SmartM365RunCleanup $smartM365RunGuard 'splash' { Close-SmartM365GuiSplash $splash }}
        if(Get-Command Write-SmartM365CompletionBanner -ErrorAction SilentlyContinue){
            $log=Get-Variable GuiLogPath -Scope Script -ValueOnly -ErrorAction SilentlyContinue
            Write-SmartM365CompletionBanner -Status Auto -ScriptName $smartM365RunGuard.ScriptName -StartedAt $smartM365RunGuard.StartedAt -LogPath $log -WarningCount $smartM365RunGuard.CleanupWarnings.Count
        }
    }
    if($smartM365RunGuard.ExitCode -eq 130){exit 130}
    }
}
if($smartM365RunGuard.ExitCode -ne 0){exit $smartM365RunGuard.ExitCode}


# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD/ot+byaO87iKR
# 2DKOiU759FSG+yttowr4OKgevqHP0aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIB5K9ZwtyJpxTyQrcr8z1Ud2wsGgiU2EyKREJAwD+C7aMA0GCSqG
# SIb3DQEBAQUABIIBgDFu1dNPE8mL2Rudy1SHPaRPZpPfMk3AUAhtqsd78BiuF4EI
# XOu9yEuLzOr2JLSxml+mKspIq9thJPdgwtcqc0gTs6c4IMgy/GbSRgeqmyAhdF0x
# nf8DrXw/V7uAJvad0J3QvIEB06UlhOdg8huXdgp4/XYOWUfHE6dKR/uqXTnHztPr
# 9gDEcXTkXl/jHXVtp4MhFKPPUPmLFNNSuSLmYENPR0NnDwE0IaJdbIMLPRRCaNSw
# IIJOQW2ilM2alzKFY26IvZWnZrn3mF159wHs9Zg/984nbAiqx4i+6BR4whgr1Y81
# 5keagq0b7fTwApogBEgdSUn8q5MNmyw4M8fye9UynfpZoo6fa0NjXZCi0OgEtYvD
# dD/y2JrsYlpyOYQ94YydbIQ3BTI7iLBpRmGs9uhScsK40ofK8/CeoXtnzA0cWwwV
# Uilf3LfN2ZzvTc1Sy54f92bYFtJL15uRt9Ui6hWEpsUEHuWtVm/nV2hgQ58gWQwo
# QV+JcfvL03ZrIMgP/qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzAxMDIx
# NDBaMC8GCSqGSIb3DQEJBDEiBCAyCd63ow6a4XtPWD087z12OZ0pVe7407kdwUsM
# Q2sFbDANBgkqhkiG9w0BAQEFAASCAgCd/lRVWUWXeQMZkcmFujbGVDaiLhMqrdRF
# kWRBNmBhiC0fbjydIyr8jfbykR65PVr0VRHSZi6KQnu1Qnk/jIqK8rg5glpKk+Yy
# Wz0t6SPETA62GcxETBTDh1SCwIpx+IM3HvQVUURoFuu4mPJwdcvobWg2xh4gOapl
# yTAXHNMAijhFIsIGRoeuYuWSRDkKcL9hIS53hyhihyhObQgkoaV24t2Mm/HioCq/
# yDYsWV6AM5tx2JDyKPqBBhAwFRY4cWOiXEayJkMGcJxvOXpvc2ckDfy+vTzqbPVl
# 5ZbTNVCzWYGhdpC/96oLMp8/xXTXu7nbomEZzyqVsyVsWZkiDpSnqJb4dHVNJrBQ
# Tuv/vcR0SzwgUfUNar8vVMadzLMnkDGcQjGlOXIJGY6DUIONhPSGEgGddnQDaGvB
# +bdV/ZsYt+NjmvXgLUF/GyicgFbniVGnKc3jTYEHmP4vKyDWLzzeg9N2Frnthptf
# 4qkJWyhtD38aIn0Ojog5KXOSXC0yk2N7mZ9OqsHgC9lO9N4ZGnblQMxrngUeyRrN
# rcJA1p9NAmLE40okrOoGzQhvG6QEjJ526aya8fucj4s6Ha2U1uOlgbPnq4BhnDjV
# YGybeTY0GdEWVA3p2HW9jnQlq46XWx84GS8bQM77J2J7nsNWVYclz95BM4olIukj
# Q2HlZdkQbg==
# SIG # End signature block
