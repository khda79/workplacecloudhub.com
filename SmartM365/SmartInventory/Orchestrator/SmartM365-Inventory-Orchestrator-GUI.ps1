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
1.3.6
#>
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$SharedDataFolderPath = '',
    [string]$GuiLogFolderPath = '',
    [switch]$ValidateOnly,
    [switch]$SmokeTest
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:AppVersion = '1.3.6'
$script:StartupClock = [Diagnostics.Stopwatch]::StartNew()
$script:Snapshot = $null
$script:DraftJobs = $null
$script:DraftCluster = $null
$script:PlanningRows = @()
$script:HistoryRows = @()
$script:InitialHistoryLoaded = $false
$script:RecentRuns = @()
$script:HealthByName = @{}
$script:RequestRows = @()
$script:MailFolderPath = ''
$script:AutoRefreshTimer = $null
$script:CancellationTimer = $null
$script:CancellationBusy = $false
$script:CancellationUiState = @{}
$script:CancellationContext = [pscustomobject]@{
    Worker = $null; Progress = [hashtable]::Synchronized(@{ Message = '' })
    Input = $null; Clock = [Diagnostics.Stopwatch]::new()
}
$script:HistoryFailureStatuses = @('Failed', 'TimedOut', 'Interrupted')
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
        <Style TargetType="Button">
            <Setter Property="Padding" Value="13,7"/>
            <Setter Property="Margin" Value="4,0"/>
            <Setter Property="MinHeight" Value="32"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Padding" Value="6,4"/>
            <Setter Property="Margin" Value="0,3,0,8"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Padding" Value="5,3"/>
            <Setter Property="Margin" Value="0,3,0,8"/>
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
            <Setter Property="RowHeaderWidth" Value="0"/>
        </Style>
        <Style TargetType="GroupBox">
            <Setter Property="Margin" Value="0,0,0,12"/>
            <Setter Property="Padding" Value="10"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrushSoft}"/>
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
                <Button x:Name="ValidateButton" Content="Validate draft"/>
                <Button x:Name="RebalanceButton" Content="Rebalance now" Background="#FFF4CE"/>
                <Button x:Name="PublishButton" Content="Publish changes" Background="{StaticResource AccentBrush}" Foreground="White" FontWeight="SemiBold"/>
            </StackPanel>
        </Grid>

        <TabControl x:Name="MainTabs" Grid.Row="2">
            <TabItem Header="Dashboard">
                <Grid Margin="12">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="92"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <UniformGrid Columns="5" Margin="0,0,0,12">
                        <Border Background="White" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="7" Margin="4" Padding="12"><StackPanel><TextBlock Text="Jobs" Foreground="{StaticResource MutedBrush}"/><TextBlock x:Name="JobsCountText" FontSize="25" FontWeight="SemiBold"/></StackPanel></Border>
                        <Border Background="White" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="7" Margin="4" Padding="12"><StackPanel><TextBlock Text="Enabled" Foreground="{StaticResource MutedBrush}"/><TextBlock x:Name="EnabledCountText" FontSize="25" FontWeight="SemiBold"/></StackPanel></Border>
                        <Border Background="White" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="7" Margin="4" Padding="12"><StackPanel><TextBlock Text="Servers online" Foreground="{StaticResource MutedBrush}"/><TextBlock x:Name="OnlineServersText" FontSize="25" FontWeight="SemiBold"/></StackPanel></Border>
                        <Border Background="White" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="7" Margin="4" Padding="12"><StackPanel><TextBlock Text="Success (7 days)" Foreground="{StaticResource MutedBrush}"/><TextBlock x:Name="SuccessCountText" FontSize="25" FontWeight="SemiBold"/></StackPanel></Border>
                        <Border Background="White" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="7" Margin="4" Padding="12"><StackPanel><TextBlock Text="Failures (7 days)" Foreground="{StaticResource MutedBrush}"/><TextBlock x:Name="FailureCountText" FontSize="25" FontWeight="SemiBold"/></StackPanel></Border>
                    </UniformGrid>
                    <DataGrid x:Name="DashboardGrid" Grid.Row="1">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Server" Binding="{Binding Server}" Width="170"/>
                            <DataGridCheckBoxColumn Header="Online" Binding="{Binding Online}" Width="70"/>
                            <DataGridTextColumn Header="Heartbeat age (min)" Binding="{Binding HeartbeatAgeMinutes}" Width="140"/>
                            <DataGridTextColumn Header="Weight" Binding="{Binding Weight}" Width="70"/>
                            <DataGridTextColumn Header="Policy" Binding="{Binding Policy}" Width="220"/>
                            <DataGridTextColumn Header="Capabilities" Binding="{Binding Capabilities}" Width="*"/>
                            <DataGridTextColumn Header="Assigned jobs" Binding="{Binding AssignedJobs}" Width="100"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </TabItem>

            <TabItem Header="Operations">
                <Grid Margin="12">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="150"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="210"/>
                    </Grid.RowDefinitions>
                    <StackPanel Margin="0,0,0,10">
                        <TextBlock x:Name="MaintenanceDetailText" TextWrapping="Wrap" Margin="0,0,0,6"/>
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Reason" VerticalAlignment="Center" Margin="0,0,8,0"/>
                            <TextBox x:Name="MaintenanceReasonBox" Width="330" MaxLength="1000" ToolTip="Required to enable maintenance. Does not publish the configuration draft."/>
                            <Button x:Name="EnableMaintenanceButton" Content="Enable maintenance" Background="#FFF4CE" IsEnabled="False"/>
                            <Button x:Name="DisableMaintenanceButton" Content="Disable maintenance" IsEnabled="False"/>
                        </StackPanel>
                    </StackPanel>
                    <DataGrid x:Name="OperationsServersGrid" Grid.Row="1">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Server" Binding="{Binding Server}" Width="160"/>
                            <DataGridTextColumn Header="State" Binding="{Binding State}" Width="100"/>
                            <DataGridTextColumn Header="Heartbeat age (min)" Binding="{Binding HeartbeatAgeMinutes}" Width="140"/>
                            <DataGridTextColumn Header="Version" Binding="{Binding Version}" Width="80"/>
                            <DataGridTextColumn Header="Maintenance" Binding="{Binding Maintenance}" Width="100"/>
                            <DataGridTextColumn Header="Acknowledgement" Binding="{Binding MaintenanceAcknowledgement}" Width="145"/>
                            <DataGridTextColumn Header="Running" Binding="{Binding Running}" Width="70"/>
                            <DataGridTextColumn Header="Pending" Binding="{Binding Pending}" Width="70"/>
                            <DataGridTextColumn Header="Recycle in" Binding="{Binding RecycleIn}" Width="90"/>
                            <DataGridTextColumn Header="PID" Binding="{Binding Pid}" Width="70"/>
                            <DataGridTextColumn Header="State persistence" Binding="{Binding StatePersistence}" Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    <Grid Grid.Row="2" Margin="0,10,0,0">
                        <Grid.ColumnDefinitions><ColumnDefinition Width="9*"/><ColumnDefinition Width="11*"/></Grid.ColumnDefinitions>
                        <GroupBox Header="Running jobs" Margin="0,0,6,0">
                            <DataGrid x:Name="OperationsRunningGrid">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Server" Binding="{Binding Server}" Width="130"/>
                                    <DataGridTextColumn Header="Job" Binding="{Binding Job}" Width="*"/>
                                    <DataGridTextColumn Header="Started" Binding="{Binding Started}" Width="120"/>
                                    <DataGridTextColumn Header="Duration (min)" Binding="{Binding DurationMinutes}" Width="100"/>
                                </DataGrid.Columns>
                            </DataGrid>
                        </GroupBox>
                        <GroupBox Grid.Column="1" Header="Pending jobs (why they wait)" Margin="6,0,0,0">
                            <DataGrid x:Name="OperationsPendingGrid">
                                <DataGrid.Columns>
                                    <DataGridTextColumn Header="Server" Binding="{Binding Server}" Width="110"/>
                                    <DataGridTextColumn Header="Job" Binding="{Binding Job}" Width="165"/>
                                    <DataGridTextColumn Header="Reason" Binding="{Binding Reason}" Width="125"/>
                                    <DataGridTextColumn Header="Waiting (min)" Binding="{Binding WaitingMinutes}" Width="85"/>
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
                    <Grid Grid.Row="3" Margin="0,10,0,0">
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

            <TabItem Header="Planning">
                <Grid Margin="12">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="360"/>
                    </Grid.ColumnDefinitions>
                    <DataGrid x:Name="PlanningGrid" SelectionMode="Single">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Job" Binding="{Binding Name}" SortMemberPath="Name" SortDirection="Ascending" Width="250"/>
                            <DataGridCheckBoxColumn Header="Enabled" Binding="{Binding Enabled}" Width="70"/>
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
                            <DataGridTextColumn Header="Start" Binding="{Binding StartTime}" Width="145"/>
                            <DataGridTextColumn Header="Server" Binding="{Binding Server}" Width="150"/>
                            <DataGridTextColumn Header="Job" Binding="{Binding JobName}" Width="250"/>
                            <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="100"/>
                            <DataGridTextColumn Header="Duration (sec)" Binding="{Binding DurationSec}" Width="110"/>
                            <DataGridTextColumn Header="Exit" Binding="{Binding ExitCode}" Width="60"/>
                            <DataGridTextColumn Header="Retry" Binding="{Binding RetryCount}" Width="60"/>
                            <DataGridTextColumn Header="Log" Binding="{Binding LogPath}" Width="*"/>
                        </DataGrid.Columns>
                    </DataGrid>
                </Grid>
            </TabItem>

            <TabItem Header="Requests">
                <Grid Margin="12">
                    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                    <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                        <TextBlock Text="Cancellation reason" VerticalAlignment="Center" Margin="0,0,8,0"/>
                        <TextBox x:Name="CancellationReasonBox" Width="340" Margin="0,0,8,0"/>
                        <Button x:Name="CancelRequestButton" Content="Cancel remaining jobs" IsEnabled="False" ToolTip="Cancel pending jobs and retries in the selected request. Running collectors continue. All residents must support cancellation."/>
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

            <TabItem Header="Servers and versions">
                <Grid Margin="12">
                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="390"/></Grid.ColumnDefinitions>
                    <Grid>
                        <Grid.RowDefinitions><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                        <DataGrid x:Name="ServersGrid">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Server" Binding="{Binding Server}" Width="170"/>
                                <DataGridCheckBoxColumn Header="Online" Binding="{Binding Online}" Width="70"/>
                                <DataGridTextColumn Header="Weight" Binding="{Binding Weight}" Width="70"/>
                                <DataGridTextColumn Header="Policy" Binding="{Binding Policy}" Width="220"/>
                                <DataGridTextColumn Header="Capabilities" Binding="{Binding Capabilities}" Width="*"/>
                                <DataGridTextColumn Header="Jobs" Binding="{Binding AssignedJobs}" Width="60"/>
                            </DataGrid.Columns>
                        </DataGrid>
                        <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,10,0,0">
                            <TextBox x:Name="NewServerBox" Width="220" Margin="0,0,8,0"/>
                            <Button x:Name="AddServerButton" Content="Add server"/>
                            <Button x:Name="RemoveServerButton" Content="Remove selected"/>
                        </StackPanel>
                    </Grid>
                    <ScrollViewer Grid.Column="1" Margin="14,0,0,0" VerticalScrollBarVisibility="Auto">
                        <StackPanel>
                            <GroupBox Header="Selected server">
                                <StackPanel>
                                    <TextBlock x:Name="SelectedServerText" FontWeight="SemiBold" Text="Select a server"/>
                                    <TextBlock Text="Election weight" Margin="0,8,0,0"/>
                                    <TextBox x:Name="ServerWeightBox"/>
                                    <TextBlock Text="Only jobs requiring capabilities, comma separated (blank = unrestricted)"/>
                                    <ComboBox x:Name="ServerPolicyCombo" IsEditable="True">
                                        <ComboBoxItem Content=""/>
                                        <ComboBoxItem Content="ExchangeOnPrem"/>
                                        <ComboBoxItem Content="AD"/>
                                        <ComboBoxItem Content="Graph"/>
                                        <ComboBoxItem Content="EXO"/>
                                        <ComboBoxItem Content="TeamsPowerShell"/>
                                    </ComboBox>
                                    <Button x:Name="ApplyServerButton" Content="Apply server settings to draft" Background="#E5F1FB" Margin="0,6,0,0"/>
                                </StackPanel>
                            </GroupBox>
                            <GroupBox Header="Configuration versions">
                                <Grid>
                                    <Grid.RowDefinitions><RowDefinition Height="260"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
                                    <DataGrid x:Name="VersionsGrid">
                                        <DataGrid.Columns>
                                            <DataGridTextColumn Header="Version" Binding="{Binding VersionId}" Width="*"/>
                                            <DataGridTextColumn Header="Created" Binding="{Binding Created}" Width="135"/>
                                        </DataGrid.Columns>
                                    </DataGrid>
                                    <Button x:Name="RollbackButton" Grid.Row="1" Content="Rollback to state before selected version" Margin="0,8,0,0"/>
                                </Grid>
                            </GroupBox>
                        </StackPanel>
                    </ScrollViewer>
                </Grid>
            </TabItem>

            <TabItem Header="Activity">
                <Grid Margin="12">
                    <TextBox x:Name="ActivityBox" FontFamily="Consolas" FontSize="12" IsReadOnly="True" AcceptsReturn="True" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
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
Import-Module -Name (Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365.Orchestrator.Insights.psm1') -Force -ErrorAction Stop
Import-Module -Name (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.Maintenance.psm1') -ErrorAction Stop
Import-Module -Name (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.Pipeline.psm1') -ErrorAction Stop
Import-Module -Name (Join-Path $PSScriptRoot 'SmartM365.Orchestrator.GuiWorker.psm1') -ErrorAction Stop

if ($ValidateOnly) {
    $validationWindow = ConvertFrom-OrchestratorGuiXaml -Text $xaml
    foreach ($controlName in @('PlanningGrid', 'HistoryGrid', 'ServersGrid', 'VersionsGrid', 'PublishButton', 'RebalanceButton', 'ApplyJobButton', 'ApplyServerButton', 'RollbackButton', 'DaysPanel', 'MondayCheck', 'TuesdayCheck', 'WednesdayCheck', 'ThursdayCheck', 'FridayCheck', 'SaturdayCheck', 'SundayCheck',
        'AutoRefreshCheck', 'OperationsServersGrid', 'OperationsRunningGrid', 'OperationsPendingGrid', 'OperationsIncidentsGrid', 'OperationsMailsGrid',
        'DependsOnBox', 'DependencyModeCombo', 'DependencyMaxAgeBox', 'DependentsText', 'ReadinessGrid', 'IncludeDependenciesCheck', 'RequestRunButton',
        'Failures24hButton', 'RequestsGrid', 'RequestJobsGrid', 'CancellationReasonBox', 'CancellationProgressText', 'CancelRequestButton', 'CancelAllRequestsButton', 'MaintenanceBannerText', 'MaintenanceDetailText',
        'MaintenanceReasonBox', 'EnableMaintenanceButton', 'DisableMaintenanceButton')) {
        if (-not $validationWindow.FindName($controlName)) { throw "Required XAML control not found: $controlName" }
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
    return
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
        $line = '[{0}][{1}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Level, $physicalLine
        if ($script:Controls -and $script:Controls.ActivityBox) {
            $script:Controls.ActivityBox.AppendText($line + [Environment]::NewLine)
            $script:Controls.ActivityBox.ScrollToEnd()
        }
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

function Refresh-PlanningView {
    $owners = Get-ElectionOwners
    $rows = foreach ($job in @($script:DraftJobs.Jobs)) {
        $mode = if ($job.PSObject.Properties['AssignmentMode']) { [string]$job.AssignmentMode } else { 'Legacy' }
        $allowedServers = @(Get-OrchestratorGuiPropertyValue -Object $job -Name 'AllowedServers' -DefaultValue @())
        $server = switch ($mode) {
            'Pinned' { $allowedServers -join ', ' }
            'Elected' { if ($owners.ContainsKey([string]$job.Name)) { $owners[[string]$job.Name] } else { 'Not elected' } }
            'Manual' { 'Manual' }
            default { $allowedServers -join ', ' }
        }
        $health = if ($script:HealthByName.ContainsKey([string]$job.Name)) { $script:HealthByName[[string]$job.Name] } else { $null }
        [pscustomobject]@{
            Name = [string]$job.Name
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
        }
    }
    $script:PlanningRows = @($rows | Sort-Object -Property Name)
    $script:Controls.PlanningGrid.ItemsSource = $script:PlanningRows
    Set-PlanningDefaultSort
}

function Refresh-ServersView {
    $servers = @(Get-SmartM365OrchestratorServerStatus -SharedDataFolderPath $script:SharedDataFolderPath -ClusterDocument $script:DraftCluster)
    $script:Controls.DashboardGrid.ItemsSource = $servers
    $script:Controls.ServersGrid.ItemsSource = $servers
    $script:Controls.PinnedServerCombo.ItemsSource = @($script:DraftCluster.ExpectedOrchestratorServers)
    $script:Controls.OnlineServersText.Text = '{0}/{1}' -f @($servers | Where-Object Online).Count, @($servers).Count
    return $servers
}

function Refresh-HistoryView {
    $from = if ($script:Controls.HistoryFromPicker.SelectedDate) { [datetime]$script:Controls.HistoryFromPicker.SelectedDate } else { (Get-Date).AddDays(-7) }
    $to = if ($script:Controls.HistoryToPicker.SelectedDate) { ([datetime]$script:Controls.HistoryToPicker.SelectedDate).Date.AddDays(1).AddTicks(-1) } else { Get-Date }
    $server = Get-ComboText -Combo $script:Controls.HistoryServerCombo
    $job = Get-ComboText -Combo $script:Controls.HistoryJobCombo
    $status = Get-ComboText -Combo $script:Controls.HistoryStatusCombo
    if ($server -eq 'All') { $server = '' }
    if ($job -eq 'All') { $job = '' }
    if ($status -eq 'All') { $status = '' }
    $script:HistoryRows = @(Get-SmartM365OrchestratorHistory -SharedDataFolderPath $script:SharedDataFolderPath -From $from -To $to -Server $server -JobName $job -Status $status)
    $script:Controls.HistoryGrid.ItemsSource = $script:HistoryRows
    Write-GuiActivity -Message ("History refreshed: {0} run(s)." -f $script:HistoryRows.Count)
}

function Update-JobHealth {
    $script:RecentRuns = @(Get-SmartM365OrchestratorRecentRuns -SharedDataFolderPath $script:SharedDataFolderPath -Days 21)
    $script:HealthByName = @{}
    foreach ($health in @(Get-SmartM365OrchestratorJobHealth -JobsDocument $script:DraftJobs -Runs $script:RecentRuns)) { $script:HealthByName[[string]$health.Name] = $health }
}

function Refresh-OperationsView {
    try {
        $publishedCluster = Read-SmartM365OrchestratorJson -Path (Join-Path $script:SharedDataFolderPath 'Config/Orchestrator-Cluster.json')
        $operations = Get-SmartM365OrchestratorOperations -SharedDataFolderPath $script:SharedDataFolderPath -ClusterDocument $publishedCluster -MailFolderPath $script:MailFolderPath -MailHours 24
        Refresh-MaintenanceView -Operations $operations
        $script:Controls.OperationsServersGrid.ItemsSource = @($operations.Servers)
        $script:Controls.OperationsRunningGrid.ItemsSource = @($operations.Running)
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
        $script:MaintenanceViewState = $control
        $script:Controls.EnableMaintenanceButton.IsEnabled = (-not $control.Enabled -and $applied)
        $script:Controls.DisableMaintenanceButton.IsEnabled = $control.Enabled
        $stateText = if ($control.Enabled) { if ($applied) { 'ACTIVE' } else { 'ACTIVATION REQUESTED' } }
            elseif ($applied) { 'INACTIVE' } else { 'ACKNOWLEDGEMENT PENDING' }
        $script:Controls.MaintenanceBannerText.Text = "Maintenance: $stateText | revision $($control.Revision) | scheduled launches only; manual Pipeline requests remain available"
        $running = ($servers | Measure-Object Running -Sum).Sum
        $stamp = if ($control.ChangedAtUtc) { ([datetimeoffset]::Parse($control.ChangedAtUtc)).LocalDateTime.ToString('yyyy-MM-dd HH:mm:ss') } else { 'not changed' }
        $script:Controls.MaintenanceDetailText.Text = "Changed: $stamp | By: $($control.ChangedBy) | Reason: $($control.Reason) | Running jobs: $running`n" + (($servers | ForEach-Object { "$($_.Server): $($_.Status)" }) -join ' | ')
    }
    catch {
        $script:MaintenanceViewState = $null
        $script:Controls.EnableMaintenanceButton.IsEnabled = $false
        $script:Controls.DisableMaintenanceButton.IsEnabled = $false
        $script:Controls.MaintenanceBannerText.Text = 'Maintenance: CONTROL UNAVAILABLE - do not assume scheduling is paused or resumed'
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
    $message = "$action shared maintenance for this tenant?`n`nRunning jobs remain supervised. Manual Pipeline requests keep their dependency and concurrency checks.`nResume skips suspended automatic occurrences and retries.`n`nReason: $reason`n`nConfiguration draft changes are NOT published."
    if ([System.Windows.MessageBox]::Show($message, "$action maintenance", 'YesNo', 'Warning') -ne 'Yes') { return }
    $result = Set-SmartM365OrchestratorMaintenance -SharedDataFolderPath $script:SharedDataFolderPath -Enabled $Enabled -ExpectedRevision $script:MaintenanceViewState.Revision -Reason $reason
    Write-GuiActivity -Message ("Maintenance transition published: revision={0}; enabled={1}; reason={2}. Awaiting server acknowledgement." -f $result.Revision, $result.Enabled, $result.Reason)
    [void](Refresh-OperationsView)
}

function Update-SelectedPipelineRequest {
    $row = $script:Controls.RequestsGrid.SelectedItem
    # Do not assign an if-expression: its pipeline unwraps a singleton array.
    $jobRows = @()
    if ($null -ne $row) { $jobRows = @($row.JobRows) }
    $script:Controls.RequestJobsGrid.ItemsSource = $jobRows
    $script:Controls.CancelRequestButton.IsEnabled = -not $script:CancellationBusy -and $null -ne $row -and $row.Status -in @('Running', 'Cancelling')
}

function Set-GuiRequestRows {
    param([AllowEmptyCollection()][object[]]$Rows)
    $selectedBatch = if ($null -ne $script:Controls.RequestsGrid.SelectedItem) { [string]$script:Controls.RequestsGrid.SelectedItem.BatchId } else { '' }
    $script:RequestRows = @($Rows)
    $script:Controls.RequestsGrid.ItemsSource = $script:RequestRows
    $selectedRow = @($script:RequestRows | Where-Object { [string]$_.BatchId -eq $selectedBatch })
    if ($selectedRow.Count) { $script:Controls.RequestsGrid.SelectedItem = $selectedRow[0] }
    Update-SelectedPipelineRequest
    $script:Controls.CancelAllRequestsButton.IsEnabled = -not $script:CancellationBusy -and @($script:RequestRows | Where-Object Status -in @('Running', 'Cancelling')).Count -gt 0
}

function Refresh-RequestsView {
    try {
        if ($script:CancellationBusy) { return }
        Set-GuiRequestRows -Rows @(Get-SmartM365OrchestratorRecentPipelineRuns -SharedDataFolderPath $script:SharedDataFolderPath -Count 30)
    }
    catch { Write-GuiException -Context 'Requests refresh failed' -ErrorRecord $_ }
}

function Confirm-PipelineCancellation {
    param([string]$Message, [string]$Title)
    return [System.Windows.MessageBox]::Show($Message, $Title, 'YesNo', 'Warning') -eq 'Yes'
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
                foreach ($run in $runs) {
                    $Progress.Message = "Checking batch $($run.BatchId)..."
                    $null = Stop-SmartM365OrchestratorPipelineRequest -SharedDataFolderPath $InputData.Root -BatchId $run.BatchId -Tenant $InputData.Tenant -Reason $InputData.Reason -ValidateOnly
                }
                return [pscustomobject]@{ Phase = 'Validate'; BatchIds = @($runs | ForEach-Object BatchId) }
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
    $names = @('RefreshButton','ValidateButton','RebalanceButton','PublishButton','HistoryRefreshButton',
        'RequestRunButton','EnableMaintenanceButton','DisableMaintenanceButton','RollbackButton',
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
        $script:Controls.CancelAllRequestsButton.IsEnabled = @($script:RequestRows | Where-Object Status -in @('Running', 'Cancelling')).Count -gt 0
    }
}

function Start-GuiPipelineCancellation {
    param([switch]$All, [string]$BatchId = '')
    if ($script:CancellationBusy) { return }
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
    if ($null -eq $row -or $row.Status -notin @('Running', 'Cancelling')) { throw 'Select an active request.' }
    Start-GuiPipelineCancellation -BatchId $row.BatchId
}

function Stop-AllPipelineRequests {
    Start-GuiPipelineCancellation -All
}

function Show-GuiCancellationError {
    param([string]$Message)
    Write-GuiActivity -Message $Message -Level ERROR -DisplayOnly
    [System.Windows.MessageBox]::Show($Message, 'Cancellation failed', 'OK', 'Error') | Out-Null
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
    if ($script:CancellationBusy) { return }
    [void](Refresh-OperationsView)
    Refresh-RequestsView
    $script:Controls.LastRefreshText.Text = 'Live refresh: ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
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
    param([switch]$SkipConfirmation)

    $row = $script:Controls.PlanningGrid.SelectedItem
    if ($null -eq $row) { throw 'Select a job first.' }
    $jobName = [string]$row.Name
    $includeDependencies = [bool]$script:Controls.IncludeDependenciesCheck.IsChecked
    # Fast pre-check; New-SmartM365OrchestratorPipelineRequest repeats it under the submission lock.
    Refresh-RequestsView
    $activeRequest = @($script:RequestRows | Where-Object Status -in @('Running', 'Cancelling') | Select-Object -First 1)
    if ($activeRequest.Count) { throw "An active pipeline request already exists: $($activeRequest[0].BatchId) (created $($activeRequest[0].Created), $($activeRequest[0].Pending) job(s) not finished). See the Requests tab." }
    if (-not $SkipConfirmation) {
        $message = "Ask the orchestrators to run '$jobName' now?"
        if ($includeDependencies) { $message += "`n`nIts enabled dependencies are included and run first." }
        else { $message += "`n`nDependencies are not run; the job starts on its elected server when its dependency rule allows it." }
        if (Test-JobDraftChanged -JobName $jobName) { $message += "`n`nThis job has unpublished draft changes. The request uses the PUBLISHED configuration." }
        $message += "`n`nOnly one request can be active at a time."
        if ([System.Windows.MessageBox]::Show($message, 'Request run', 'YesNo', 'Question') -ne 'Yes') { return $null }
    }
    $request = New-SmartM365OrchestratorJobRunRequest -SharedDataFolderPath $script:SharedDataFolderPath -JobsPath (Get-PublishedJobsPath) -JobName @($jobName) -Tenant $Tenant -IncludeDependencies:$includeDependencies
    Write-GuiActivity -Message ("Run requested. Job={0}; IncludeDependencies={1}; BatchId={2}; Jobs={3}" -f $jobName, $includeDependencies, $request.BatchId, $request.TotalCount) -Level SUCCESS
    $script:Controls.StatusText.Text = "Run requested: $($request.BatchId)"
    Refresh-RequestsView
    if (-not $SkipConfirmation) {
        [System.Windows.MessageBox]::Show("Request submitted.`nBatch: $($request.BatchId)`nJobs: $($request.TotalCount)`n`nFollow it in the Requests tab.", 'Run requested', 'OK', 'Information') | Out-Null
    }
    return $request
}

function Show-FailuresLast24Hours {
    $from = (Get-Date).AddHours(-24)
    $script:Controls.HistoryFromPicker.SelectedDate = $from.Date
    $script:Controls.HistoryToPicker.SelectedDate = (Get-Date).Date
    $script:Controls.HistoryServerCombo.SelectedIndex = 0
    $script:Controls.HistoryJobCombo.SelectedIndex = 0
    $script:Controls.HistoryStatusCombo.SelectedIndex = 0
    $script:HistoryRows = @(Get-SmartM365OrchestratorHistory -SharedDataFolderPath $script:SharedDataFolderPath -From $from -To (Get-Date) | Where-Object { $_.Status -in $script:HistoryFailureStatuses })
    $script:Controls.HistoryGrid.ItemsSource = $script:HistoryRows
    Write-GuiActivity -Message ("Failures in the last 24 h: {0} run(s)." -f $script:HistoryRows.Count)
}

function Refresh-AllViews {
    try {
        $script:Snapshot = Get-SmartM365OrchestratorConfigurationSnapshot -SharedDataFolderPath $script:SharedDataFolderPath
        $script:DraftJobs = Copy-JsonDocument -Document $script:Snapshot.Jobs
        $script:DraftCluster = Copy-JsonDocument -Document $script:Snapshot.Cluster
        try { Update-JobHealth } catch { $script:HealthByName = @{}; Write-GuiException -Context 'Job health could not be computed' -ErrorRecord $_ }
        Refresh-PlanningView
        $servers = @(Refresh-ServersView)
        $historyNow = Get-Date
        $historyFrom = $historyNow.AddDays(-7)
        $history7 = @(Get-SmartM365OrchestratorHistory -SharedDataFolderPath $script:SharedDataFolderPath -From $historyFrom.Date -To $historyNow.Date.AddDays(1).AddTicks(-1))
        if (-not $script:InitialHistoryLoaded) {
            $script:HistoryRows = $history7
            $script:Controls.HistoryGrid.ItemsSource = $script:HistoryRows
            $script:InitialHistoryLoaded = $true
        }
        $script:Controls.JobsCountText.Text = [string]@($script:DraftJobs.Jobs).Count
        $script:Controls.EnabledCountText.Text = [string]@($script:DraftJobs.Jobs | Where-Object Enabled).Count
        $script:Controls.SuccessCountText.Text = [string]@($history7 | Where-Object { $_.StartTime -ge $historyFrom -and $_.StartTime -le $historyNow -and $_.Status -eq 'Success' }).Count
        $script:Controls.FailureCountText.Text = [string]@($history7 | Where-Object { $_.StartTime -ge $historyFrom -and $_.StartTime -le $historyNow -and $_.Status -in $script:HistoryFailureStatuses }).Count
        [void](Refresh-OperationsView)
        Refresh-RequestsView
        $script:Controls.VersionsGrid.ItemsSource = @(Get-SmartM365OrchestratorConfigurationVersions -SharedDataFolderPath $script:SharedDataFolderPath)
        $script:Controls.ConnectionText.Text = "Tenant: $Tenant"
        $script:Controls.LastRefreshText.Text = 'Refreshed: ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        $script:Controls.StatusText.Text = 'Shared configuration loaded'
        Write-GuiActivity -Message ("Configuration loaded: {0} jobs, {1} servers." -f @($script:DraftJobs.Jobs).Count, @($servers).Count) -Level SUCCESS
    }
    catch {
        $script:Controls.StatusText.Text = 'Refresh failed'
        Write-GuiException -Context 'Refresh failed' -ErrorRecord $_
        [System.Windows.MessageBox]::Show($_.Exception.Message, 'Refresh failed', 'OK', 'Error') | Out-Null
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
            if ([System.Windows.MessageBox]::Show($warning, 'Job used as a dependency', 'YesNo', 'Warning') -ne 'Yes') { return }
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
    $previousAllowedText = if ($previousAllowedServers.Count -gt 0) { $previousAllowedServers -join ',' } else { '<none>' }
    $newAllowedText = if ($newAllowedServers.Count -gt 0) { $newAllowedServers -join ',' } else { '<none>' }
    Write-GuiActivity -Message ("Job '{0}' updated in the draft. AssignmentMode={1}->{2}; AllowedServers={3}->{4}; DependsOn={5}; DependencyMode={6}; DependencyMaxAgeHours={7}" -f $job.Name, $previousMode, $mode, $previousAllowedText, $newAllowedText, ($(if ($dependsOn.Count) { $dependsOn -join ',' } else { '<none>' })), $dependencyMode, $dependencyMaxAge)
}

function Apply-SelectedServerToDraft {
    $row = $script:Controls.ServersGrid.SelectedItem
    if ($null -eq $row) { throw 'Select a server first.' }
    $server = ([string]$row.Server).ToUpperInvariant()
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
    $validation = Test-SmartM365OrchestratorClusterDocument -Document $script:DraftCluster
    if (-not $validation.Valid) { throw $validation.Errors -join [Environment]::NewLine }
    Refresh-ServersView | Out-Null
    $script:Controls.StatusText.Text = "Draft changed: $server"
    Write-GuiActivity -Message "Server '$server' settings updated in the draft."
}

function Show-DraftValidation {
    $jobsValidation = Test-SmartM365OrchestratorJobsDocument -Document $script:DraftJobs
    $clusterValidation = Test-SmartM365OrchestratorClusterDocument -Document $script:DraftCluster
    $consistencyValidation = Test-SmartM365OrchestratorConfigurationConsistency -JobsDocument $script:DraftJobs -ClusterDocument $script:DraftCluster
    $errors = @($jobsValidation.Errors) + @($clusterValidation.Errors) + @($consistencyValidation.Errors)
    $warnings = @($jobsValidation.Warnings) + @($clusterValidation.Warnings) + @($consistencyValidation.Warnings)
    if ($errors.Count -gt 0) {
        [System.Windows.MessageBox]::Show(($errors -join [Environment]::NewLine), 'Invalid draft', 'OK', 'Error') | Out-Null
        return $false
    }
    $message = "Draft is valid.`nJobs: $($jobsValidation.JobCount)`nServers: $($clusterValidation.ServerCount)"
    if ($warnings.Count -gt 0) { $message += "`n`nWarnings:`n" + ($warnings -join [Environment]::NewLine) }
    [System.Windows.MessageBox]::Show($message, 'Validation', 'OK', 'Information') | Out-Null
    return $true
}

function Publish-Draft {
    if (-not (Show-DraftValidation)) { return }
    $enabled = @($script:DraftJobs.Jobs | Where-Object Enabled).Count
    $confirmation = "Publish the shared configuration?`n`nJobs: $(@($script:DraftJobs.Jobs).Count) ($enabled enabled)`nServers: $(@($script:DraftCluster.ExpectedOrchestratorServers).Count)`n`nEvery Orchestrator server will reload it automatically."
    if ([System.Windows.MessageBox]::Show($confirmation, 'Publish shared configuration', 'YesNo', 'Warning') -ne 'Yes') { return }
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
        [System.Windows.MessageBox]::Show("Configuration published successfully.`nVersion: $($result.VersionId)", 'Published', 'OK', 'Information') | Out-Null
        Refresh-AllViews
    }
    catch {
        Write-GuiException -Context 'Publication failed' -ErrorRecord $_
        [System.Windows.MessageBox]::Show($_.Exception.Message, 'Publication failed', 'OK', 'Error') | Out-Null
    }
}

function Request-ElectionRebalance {
    param([switch]$SkipConfirmation)

    if (-not $SkipConfirmation) {
        $confirmation = "Recalculate elected job owners now?`n`nThe active published configuration, current live capabilities, server weights, policies and duration history will be used.`nPinned and Manual jobs are not changed.`n`nUnsaved draft changes are not included; publish them first if they must affect this rebalance."
        if ([System.Windows.MessageBox]::Show($confirmation, 'Rebalance elected jobs', 'YesNo', 'Warning') -ne 'Yes') { return $null }
    }
    try {
        $result = Request-SmartM365OrchestratorRebalance -SharedDataFolderPath $script:SharedDataFolderPath
        $script:Controls.StatusText.Text = "Rebalance requested: $($result.RequestId)"
        Write-GuiActivity -Message ("Election rebalance requested. RequestId={0}; RequestPath={1}" -f $result.RequestId, $result.RequestPath) -Level SUCCESS
        if (-not $SkipConfirmation) {
            [System.Windows.MessageBox]::Show("Rebalance request submitted successfully.`nRequest: $($result.RequestId)`n`nA resident Orchestrator will apply it on its next tick. Refresh the GUI to see the new owners.", 'Rebalance requested', 'OK', 'Information') | Out-Null
        }
        return $result
    }
    catch {
        Write-GuiException -Context 'Election rebalance request failed' -ErrorRecord $_
        if ($SkipConfirmation) { throw }
        [System.Windows.MessageBox]::Show($_.Exception.Message, 'Rebalance failed', 'OK', 'Error') | Out-Null
        return $null
    }
}

$smartM365Root = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$tenantContextPath = Join-Path -Path $smartM365Root -ChildPath 'Config\SmartM365-TenantContext.ps1'
. $tenantContextPath
$script:EffectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$localConfigPath = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-Inventory-Orchestrator.local.json'; $localConfigPath = Resolve-SmartM365JsonConfigurationPath -Path $localConfigPath
if (-not (Test-Path -LiteralPath $localConfigPath)) {
    Copy-Item -LiteralPath ($localConfigPath + '.template') -Destination $localConfigPath -ErrorAction Stop
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

$bootstrapJobsPath = Join-Path -Path $PSScriptRoot -ChildPath 'Orchestrator-Jobs.json'
$selectedBootstrap=Get-SmartM365JsonReadPath $bootstrapJobsPath -Optional
if ($selectedBootstrap) { $bootstrapJobsPath=$selectedBootstrap } else { $bootstrapJobsPath += '.template' }
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
    'HeaderLogo', 'SharedPathText', 'ConnectionText', 'LastRefreshText', 'StatusText', 'RefreshButton', 'ValidateButton', 'RebalanceButton', 'PublishButton',
    'JobsCountText', 'EnabledCountText', 'OnlineServersText', 'SuccessCountText', 'FailureCountText', 'DashboardGrid',
    'PlanningGrid', 'SelectedJobText', 'JobEnabledCheck', 'ScheduleTypeCombo', 'TimesBox', 'DaysPanel',
    'MondayCheck', 'TuesdayCheck', 'WednesdayCheck', 'ThursdayCheck', 'FridayCheck', 'SaturdayCheck', 'SundayCheck', 'MissedPolicyCombo',
    'AssignmentCombo', 'PinnedServerCombo', 'TimeoutBox', 'RetriesBox', 'RetryDelayBox', 'DurationBox', 'ApplyJobButton',
    'HistoryFromPicker', 'HistoryToPicker', 'HistoryServerCombo', 'HistoryJobCombo', 'HistoryStatusCombo', 'HistoryRefreshButton',
    'ExportCsvButton', 'ExportHtmlButton', 'HistoryGrid', 'ServersGrid', 'NewServerBox', 'AddServerButton', 'RemoveServerButton',
    'SelectedServerText', 'ServerWeightBox', 'ServerPolicyCombo', 'ApplyServerButton', 'VersionsGrid', 'RollbackButton',
    'ActivityBox', 'FooterText', 'VersionText',
    'AutoRefreshCheck', 'OperationsServersGrid', 'OperationsRunningGrid', 'OperationsPendingGrid', 'OperationsIncidentsGrid', 'OperationsMailsGrid',
    'DependsOnBox', 'DependencyModeCombo', 'DependencyMaxAgeBox', 'DependentsText', 'ReadinessGrid', 'IncludeDependenciesCheck', 'RequestRunButton',
    'Failures24hButton', 'RequestsGrid', 'RequestJobsGrid', 'CancellationReasonBox', 'CancellationProgressText', 'CancelRequestButton', 'CancelAllRequestsButton', 'MaintenanceBannerText', 'MaintenanceDetailText',
    'MaintenanceReasonBox', 'EnableMaintenanceButton', 'DisableMaintenanceButton'
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
$script:Controls.FooterText.Text = 'Shared changes are validated, versioned and audited. Run requests are executed by the orchestrators; this GUI never starts or stops a process.'
$script:Controls.VersionText.Text = "v$($script:AppVersion)"
$script:Controls.HistoryFromPicker.SelectedDate = (Get-Date).AddDays(-7).Date
$script:Controls.HistoryToPicker.SelectedDate = (Get-Date).Date
$script:Controls.HistoryStatusCombo.ItemsSource = @('All', 'Success', 'CompletedWithWarnings', 'Failed', 'TimedOut', 'Interrupted', 'Retried')
$script:Controls.HistoryStatusCombo.SelectedIndex = 0

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
    if ($null -eq $row) { return }
    $script:Controls.SelectedServerText.Text = [string]$row.Server
    $script:Controls.ServerWeightBox.Text = ([double]$row.Weight).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    Select-ComboText -Combo $script:Controls.ServerPolicyCombo -Text ([string]$row.Policy)
})
$script:Controls.ApplyJobButton.Add_Click({ try { Apply-SelectedJobToDraft } catch { Write-GuiException -Context 'Invalid job draft' -ErrorRecord $_; [System.Windows.MessageBox]::Show($_.Exception.Message, 'Invalid job', 'OK', 'Error') | Out-Null } })
$script:Controls.ApplyServerButton.Add_Click({ try { Apply-SelectedServerToDraft } catch { Write-GuiException -Context 'Invalid server draft' -ErrorRecord $_; [System.Windows.MessageBox]::Show($_.Exception.Message, 'Invalid server', 'OK', 'Error') | Out-Null } })
$script:Controls.RefreshButton.Add_Click({ Refresh-AllViews })
$script:Controls.ValidateButton.Add_Click({ [void](Show-DraftValidation) })
$script:Controls.RebalanceButton.Add_Click({ [void](Request-ElectionRebalance) })
$script:Controls.PublishButton.Add_Click({ Publish-Draft })
$script:Controls.HistoryRefreshButton.Add_Click({ Refresh-HistoryView })
$script:Controls.HistoryGrid.Add_MouseDoubleClick({
    $row = $script:Controls.HistoryGrid.SelectedItem
    if ($null -eq $row) { return }
    if (-not [string]::IsNullOrWhiteSpace([string]$row.LogPath) -and (Test-Path -LiteralPath $row.LogPath)) { Start-Process -FilePath $row.LogPath; return }
    $logText = if ([string]::IsNullOrWhiteSpace([string]$row.LogPath)) { '<no log path recorded>' } else { [string]$row.LogPath }
    [System.Windows.MessageBox]::Show("The log of this run is not reachable from this computer.`n`n$logText`n`nThe log stays on the server that ran the job ($($row.Server)) until it is copied to LOG-ALL.", 'Log not found', 'OK', 'Information') | Out-Null
})
$script:Controls.Failures24hButton.Add_Click({ try { Show-FailuresLast24Hours } catch { Write-GuiException -Context 'Failure filter failed' -ErrorRecord $_ } })
$script:Controls.EnableMaintenanceButton.Add_Click({
    try { Set-GuiMaintenance -Enabled $true }
    catch { Write-GuiException -Context 'Enable maintenance failed' -ErrorRecord $_; [System.Windows.MessageBox]::Show($_.Exception.Message, 'Maintenance', 'OK', 'Error') | Out-Null; Refresh-MaintenanceView }
})
$script:Controls.DisableMaintenanceButton.Add_Click({
    try { Set-GuiMaintenance -Enabled $false }
    catch { Write-GuiException -Context 'Disable maintenance failed' -ErrorRecord $_; [System.Windows.MessageBox]::Show($_.Exception.Message, 'Maintenance', 'OK', 'Error') | Out-Null; Refresh-MaintenanceView }
})
$script:Controls.RequestRunButton.Add_Click({
    try { [void](Request-SelectedJobRun) }
    catch { Write-GuiException -Context 'Run request failed' -ErrorRecord $_; [System.Windows.MessageBox]::Show($_.Exception.Message, 'Run request failed', 'OK', 'Error') | Out-Null }
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
        $server = $script:Controls.NewServerBox.Text.Trim().ToUpperInvariant()
        if ($server -notmatch '^[A-Z0-9._-]+$') { throw 'Enter a valid server name.' }
        if ($server -in @($script:DraftCluster.ExpectedOrchestratorServers)) { throw 'This server already exists.' }
        $script:DraftCluster.ExpectedOrchestratorServers = @($script:DraftCluster.ExpectedOrchestratorServers) + $server
        Set-JsonProperty -Object $script:DraftCluster.ElectionWeightsByServer -Name $server -Value 1.0
        $script:Controls.NewServerBox.Clear()
        Refresh-ServersView | Out-Null
        Write-GuiActivity -Message "Server '$server' added to the draft."
    }
    catch { Write-GuiException -Context 'Add server failed' -ErrorRecord $_; [System.Windows.MessageBox]::Show($_.Exception.Message, 'Add server', 'OK', 'Error') | Out-Null }
})
$script:Controls.RemoveServerButton.Add_Click({
    $row = $script:Controls.ServersGrid.SelectedItem
    if ($null -eq $row) { return }
    $server = [string]$row.Server
    $pinnedJobs = @($script:DraftJobs.Jobs | Where-Object { $_.AssignmentMode -eq 'Pinned' -and $server -in @($_.AllowedServers) })
    if ($pinnedJobs.Count -gt 0) {
        [System.Windows.MessageBox]::Show("Server is still used by pinned jobs: $($pinnedJobs.Name -join ', ')", 'Cannot remove server', 'OK', 'Error') | Out-Null
        return
    }
    if ([System.Windows.MessageBox]::Show("Remove $server from the cluster draft?", 'Remove server', 'YesNo', 'Warning') -ne 'Yes') { return }
    $script:DraftCluster.ExpectedOrchestratorServers = @($script:DraftCluster.ExpectedOrchestratorServers | Where-Object { $_ -ine $server })
    [void]$script:DraftCluster.ElectionWeightsByServer.PSObject.Properties.Remove($server)
    [void]$script:DraftCluster.ServerJobPolicies.PSObject.Properties.Remove($server)
    Refresh-ServersView | Out-Null
    Write-GuiActivity -Message "Server '$server' removed from the draft."
})
$script:Controls.RollbackButton.Add_Click({
    $version = $script:Controls.VersionsGrid.SelectedItem
    if ($null -eq $version) { return }
    if ([System.Windows.MessageBox]::Show("Rollback to the configuration before $($version.VersionId)?`nThis creates a new audited version.", 'Rollback', 'YesNo', 'Warning') -ne 'Yes') { return }
    try {
        $result = Restore-SmartM365OrchestratorConfigurationVersion -SharedDataFolderPath $script:SharedDataFolderPath -VersionFolderPath $version.FolderPath -Snapshot Before -ExpectedJobsHash $script:Snapshot.JobsHash -ExpectedClusterHash $script:Snapshot.ClusterHash
        Write-GuiActivity -Message "Rollback published as $($result.VersionId)." -Level SUCCESS
        Refresh-AllViews
    }
    catch { Write-GuiException -Context 'Rollback failed' -ErrorRecord $_; [System.Windows.MessageBox]::Show($_.Exception.Message, 'Rollback failed', 'OK', 'Error') | Out-Null }
})
$script:Controls.ExportCsvButton.Add_Click({
    $dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $dialog.Filter = 'CSV files (*.csv)|*.csv'
    $dialog.FileName = 'SmartM365-Orchestrator-History-{0}.csv' -f (Get-Date).ToString('yyyyMMdd-HHmmss')
    if ($dialog.ShowDialog()) {
        $script:HistoryRows | Export-Csv -LiteralPath $dialog.FileName -NoTypeInformation -Encoding utf8
        Write-GuiActivity -Message "History exported to $($dialog.FileName)." -Level SUCCESS
    }
})
$script:Controls.ExportHtmlButton.Add_Click({
    $dialog = [Microsoft.Win32.SaveFileDialog]::new()
    $dialog.Filter = 'HTML files (*.html)|*.html'
    $dialog.FileName = 'SmartM365-Orchestrator-History-{0}.html' -f (Get-Date).ToString('yyyyMMdd-HHmmss')
    if ($dialog.ShowDialog()) {
        $body = $script:HistoryRows | Select-Object StartTime, Server, JobName, Status, DurationSec, ExitCode, RetryCount, LogPath | ConvertTo-Html -Title 'SmartM365 Orchestrator history' -PreContent '<h1>SmartM365 Orchestrator history</h1>'
        [System.IO.File]::WriteAllText($dialog.FileName, ($body -join [Environment]::NewLine), [System.Text.UTF8Encoding]::new($false))
        Write-GuiActivity -Message "History exported to $($dialog.FileName)." -Level SUCCESS
    }
})

Refresh-AllViews
Write-GuiActivity -Message ('Startup initial views loaded after {0} ms.' -f $script:StartupClock.ElapsedMilliseconds)
$script:Controls.HistoryServerCombo.ItemsSource = @('All') + @($script:DraftCluster.ExpectedOrchestratorServers)
$script:Controls.HistoryServerCombo.SelectedIndex = 0
$script:Controls.HistoryJobCombo.ItemsSource = @('All') + @($script:DraftJobs.Jobs.Name | Sort-Object)
$script:Controls.HistoryJobCombo.SelectedIndex = 0
if ($SmokeTest) {
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
    $activeRequests = @($script:RequestRows | Where-Object Status -in @('Running', 'Cancelling'))
    $publishedRunnable = @($script:Snapshot.Jobs.Jobs | Where-Object {
        $smokeMode = [string](Get-OrchestratorGuiPropertyValue -Object $_ -Name 'AssignmentMode' -DefaultValue 'Legacy')
        [bool]$_.Enabled -and @(Get-OrchestratorGuiPropertyValue -Object $_ -Name 'DependsOn' -DefaultValue @()).Count -eq 0 -and
        ($smokeMode -in @('Legacy', 'Pinned') -or ($smokeMode -eq 'Elected' -and $smokeOwners.ContainsKey([string]$_.Name)))
    } | Select-Object -First 1)
    if ($activeRequests.Count) {
        $blockedRejected = $false
        $script:Controls.PlanningGrid.SelectedItem = $script:PlanningRows[0]
        try { [void](Request-SelectedJobRun -SkipConfirmation) } catch { $blockedRejected = $_.Exception.Message -like '*active pipeline request already exists*' }
        if (-not $blockedRejected) { throw 'A run request was accepted while another request is active.' }
        Write-GuiActivity -Message ("SMOKE_TEST_RUN_REQUEST_BLOCKED_OK ActiveBatch={0}" -f $activeRequests[0].BatchId) -Level SUCCESS
    }
    elseif ($publishedRunnable.Count) {
        $script:Controls.PlanningGrid.SelectedItem = @($script:PlanningRows | Where-Object Name -eq $publishedRunnable[0].Name)[0]
        $script:Controls.IncludeDependenciesCheck.IsChecked = $false
        $runRequest = Request-SelectedJobRun -SkipConfirmation
        if (-not @($script:RequestRows | Where-Object BatchId -eq $runRequest.BatchId).Count) { throw 'The run request is missing from the Requests view.' }
        $duplicateRejected = $false
        try { [void](Request-SelectedJobRun -SkipConfirmation) } catch { $duplicateRejected = $_.Exception.Message -like '*active pipeline request already exists*' }
        if (-not $duplicateRejected) { throw 'A second run request was accepted while the first one is active.' }
        Write-GuiActivity -Message ("SMOKE_TEST_RUN_REQUEST_OK Job={0}; BatchId={1}" -f $publishedRunnable[0].Name, $runRequest.BatchId) -Level SUCCESS
    }
    else { Write-GuiActivity -Message 'SMOKE_TEST_RUN_REQUEST_SKIPPED No enabled runnable job without dependency in the published configuration.' -Level WARN }
    Write-GuiActivity -Message ("SMOKE_TEST_OPERATIONS_OK Servers={0}; HealthRows={1}" -f @($operations.Servers).Count, $script:HealthByName.Count) -Level SUCCESS

    $rebalanceRequest = Request-ElectionRebalance -SkipConfirmation
    $rebalanceRequestPath = Join-Path -Path $script:SharedDataFolderPath -ChildPath 'Election\Orchestrator-RebalanceRequest.json'
    $savedRebalanceRequest = Read-SmartM365OrchestratorJson -Path $rebalanceRequestPath
    if ($null -eq $script:Controls.RebalanceButton -or [string]$script:Controls.RebalanceButton.Content -ne 'Rebalance now') { throw 'Rebalance now button is missing.' }
    if ([string]::IsNullOrWhiteSpace([string]$rebalanceRequest.RequestId) -or [string]$savedRebalanceRequest.RequestId -cne [string]$rebalanceRequest.RequestId) {
        throw 'Rebalance request smoke test did not persist the submitted request atomically.'
    }
    Write-GuiActivity -Message ("SMOKE_TEST_REBALANCE_OK RequestId={0}" -f $rebalanceRequest.RequestId) -Level SUCCESS
    Write-GuiActivity -Message ("SMOKE_TEST_PINNED_OK Job={0}; Server={1}; DefaultSort=JobAscending" -f $smokeJob.Name, $smokeServer) -Level SUCCESS
    Write-GuiActivity -Message ("SMOKE_TEST_OK SmartM365 Orchestrator GUI v{0} | Jobs={1} | PlanningRows={2}" -f $script:AppVersion, @($script:DraftJobs.Jobs).Count, @($script:PlanningRows).Count) -Level SUCCESS
    return
}

$script:CancellationTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:CancellationTimer.Interval = [TimeSpan]::FromMilliseconds(150)
$script:CancellationTimer.Add_Tick({ Receive-GuiPipelineCancellation })
$script:CancellationTimer.Start()
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
    if ($script:CancellationTimer) { $script:CancellationTimer.Stop() }
    Write-GuiActivity -Message 'GUI session closed.'
    if ($script:GuiSplash) { Close-SmartM365GuiSplash -Splash $script:GuiSplash }
})
$window.Add_Closing({
    param($sender, $eventArgs)
    if ($script:CancellationBusy) {
        $eventArgs.Cancel = $true
        [System.Windows.MessageBox]::Show('Cancellation is still being checked or published. Wait for the result before closing this window. Running collectors are not stopped.', 'Operation in progress', 'OK', 'Information') | Out-Null
    }
})
[void]$window.ShowDialog()

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBgYdukmQ/SM1GV
# tauo/Ksc/z36+sLNOMPqdcmpdhp9yaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINHrFbHeQv3ACdi4Dm/3o6cY41HGW+yLG1OU/TgAQvgrMA0GCSqG
# SIb3DQEBAQUABIIBgApACCXNMj1zNG9zD9BoQX5nf/Izn8Za5xDxy240TNfNYB+0
# nWF0U3Asz5LA/AilBbKIGZU/58+0DIa7SVuSegHssZNtALqVlpadVFXxvfaaApH4
# 1Ol8F9NSSk9gv+/Jw3eyn8EY7zkacZEtMAuTW3mL787yZ2GEaPXC2W69M471Br5M
# YVC07k4rL5Atk13uPLZiB7GjtL0FVMW76SXe3vw1UZrFpDH2EiRcFgVvuewGiVIi
# 8rjcs4xgJd5Y/s0QI41aE4yOI8MxhblQlX5hUyhhIq8KAX9o7mIeCs6UL4RVZirM
# DeKxTa4o/D/YFLJ3jVtoOsL2HIZSkJiMTmTrg8iAesX45K9JRrSd+o+DHPFaccN2
# D3rHQy08oaiq66ycZKybemKI3WtjPK33AgraGyRHz2OtdrbfyEPBkuaNCE1E1gws
# bBnIB6wy9eVSDAyUrSk3k777l3ViXtAc7MbDfT1F2lA2SiYMFdEQ51fXp1qaSqfp
# X4FAYQ1bFElqxjeqQ6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDcxNzU1
# MjdaMC8GCSqGSIb3DQEJBDEiBCCzhmefdIT+dLD1LPuXA7XHIqlIHh8m4ESGMoiN
# CSfy7TANBgkqhkiG9w0BAQEFAASCAgAJe2IalxIK29qX71qF2/WBb7KZTPVtqcZy
# fS4cDlygV+XFB9nPKcJBdKT8FxS4QdHwHLk0S1RrSAOLiuGTH8yDku3T9OCFxzYk
# 2Ox+yY1tYnEnktcmeYwOUgYyGnuDDw52rEhGtd/psTVFzKj+izaFlzdE6Uw858VL
# G48h7VdLxeTtDKOYaWBq8C+fEbPJoY+YDI7dj0obug3h8qj/eIPwMSPjMEAdCQu9
# w1+w15qTOsC+i/187wi6Q877PTreRidMAfCchC/oheWcJasFyynGrbhe+a/xOHPH
# CgUJZ0htrWiX6/STcEs7f5uGVA2upS3BcDXS34bialpOeS9h6MyM+o0pZ4Mz9CyP
# OL7Zmm/ujWF3SKq/dt2v5whLP/7wmlwyOffWvP23Gu76C6c9rOf4mhlwpyaMMyzR
# dCDVlt/UJdEdwtJJWZH+gKwTxIYq+AHM8SWqclGai9ba2bagPIVVSEI5KnoOfTZ3
# tQRnzFm8+qbRiyX2bGhah1SQvkzWqjc3MB68hIsTjc8EJw6Z0jKcJPTPFv8m5m25
# 1q0kHEJljQkfm+bCIQInOUauDEJBROt2hCLuXaQBVKDrCfMsDoj7un1MTgKxZZH7
# IgjbluqKiIce/cuJB45GZz/O2Or0GGGP7DaYz8pXzFZ1m6H2S5Wo32uAQyiw4AO4
# RMOe1Vacog==
# SIG # End signature block
