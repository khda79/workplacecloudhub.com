#Requires -Version 7.0
<#
.SYNOPSIS
Central WPF management console for the SmartM365 Inventory Orchestrator.

.DESCRIPTION
Edits the shared jobs and cluster configuration with validation, optimistic
concurrency, atomic publication, version history and an audit CSV. It also
shows the all-server execution history and the current election plan.

.PARAMETER Tenant
SmartM365 tenant profile. Defaults to test.

.PARAMETER SharedDataFolderPath
Optional direct path to the shared Orchestrator data folder. When omitted,
the path is resolved from the tenant and local Orchestrator configuration.

.PARAMETER ValidateOnly
Parses the XAML, imports the management module and validates the committed
configuration templates without opening a window or touching runtime data.

.PARAMETER SmokeTest
Loads the complete WPF data model without showing the splash or main window.
Intended only for isolated tests with SharedDataFolderPath pointing to a temporary folder.

.VERSION
1.0.2
#>
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$SharedDataFolderPath = '',
    [switch]$ValidateOnly,
    [switch]$SmokeTest
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:AppVersion = '1.0.2'
$script:Snapshot = $null
$script:DraftJobs = $null
$script:DraftCluster = $null
$script:PlanningRows = @()
$script:HistoryRows = @()

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SmartM365 Orchestrator" Width="1500" Height="900"
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
            <TextBlock x:Name="StatusText" VerticalAlignment="Center" Foreground="{StaticResource MutedBrush}" Text="Ready"/>
            <StackPanel Grid.Column="1" Orientation="Horizontal">
                <Button x:Name="RefreshButton" Content="Refresh"/>
                <Button x:Name="ValidateButton" Content="Validate draft"/>
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

            <TabItem Header="Planning">
                <Grid Margin="12">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="360"/>
                    </Grid.ColumnDefinitions>
                    <DataGrid x:Name="PlanningGrid" SelectionMode="Single">
                        <DataGrid.Columns>
                            <DataGridTextColumn Header="Job" Binding="{Binding Name}" Width="250"/>
                            <DataGridCheckBoxColumn Header="Enabled" Binding="{Binding Enabled}" Width="70"/>
                            <DataGridTextColumn Header="Frequency" Binding="{Binding Frequency}" Width="90"/>
                            <DataGridTextColumn Header="Times" Binding="{Binding Times}" Width="155"/>
                            <DataGridTextColumn Header="Days" Binding="{Binding Days}" Width="150"/>
                            <DataGridTextColumn Header="Assignment" Binding="{Binding Assignment}" Width="90"/>
                            <DataGridTextColumn Header="Server / owner" Binding="{Binding Server}" Width="155"/>
                            <DataGridTextColumn Header="Next run" Binding="{Binding NextRun}" Width="145"/>
                            <DataGridTextColumn Header="Group" Binding="{Binding Group}" Width="100"/>
                        </DataGrid.Columns>
                    </DataGrid>
                    <ScrollViewer Grid.Column="1" Margin="14,0,0,0" VerticalScrollBarVisibility="Auto">
                        <StackPanel>
                            <GroupBox Header="Selected job">
                                <StackPanel>
                                    <TextBlock x:Name="SelectedJobText" FontWeight="SemiBold" FontSize="15" Text="Select a job"/>
                                    <CheckBox x:Name="JobEnabledCheck" Content="Enabled" Margin="0,10,0,8"/>
                                    <TextBlock Text="Frequency"/>
                                    <ComboBox x:Name="ScheduleTypeCombo"><ComboBoxItem Content="Daily"/><ComboBoxItem Content="Weekly"/></ComboBox>
                                    <TextBlock Text="Times (HH:mm, comma separated)"/>
                                    <TextBox x:Name="TimesBox"/>
                                    <TextBlock Text="Days (English, comma separated)"/>
                                    <TextBox x:Name="DaysBox"/>
                                    <TextBlock Text="Missed run policy"/>
                                    <ComboBox x:Name="MissedPolicyCombo"><ComboBoxItem Content="RunOnce"/><ComboBoxItem Content="Skip"/></ComboBox>
                                    <TextBlock Text="Assignment"/>
                                    <ComboBox x:Name="AssignmentCombo"><ComboBoxItem Content="Elected"/><ComboBoxItem Content="Pinned"/><ComboBoxItem Content="Manual"/><ComboBoxItem Content="Legacy"/></ComboBox>
                                    <TextBlock Text="Pinned server"/>
                                    <ComboBox x:Name="PinnedServerCombo" IsEditable="True"/>
                                    <UniformGrid Columns="2">
                                        <StackPanel Margin="0,0,6,0"><TextBlock Text="Timeout (min)"/><TextBox x:Name="TimeoutBox"/></StackPanel>
                                        <StackPanel Margin="6,0,0,0"><TextBlock Text="Retries"/><TextBox x:Name="RetriesBox"/></StackPanel>
                                    </UniformGrid>
                                    <UniformGrid Columns="2">
                                        <StackPanel Margin="0,0,6,0"><TextBlock Text="Retry delay (sec)"/><TextBox x:Name="RetryDelayBox"/></StackPanel>
                                        <StackPanel Margin="6,0,0,0"><TextBlock Text="Estimated (min)"/><TextBox x:Name="DurationBox"/></StackPanel>
                                    </UniformGrid>
                                    <Button x:Name="ApplyJobButton" Content="Apply to draft" Background="#E5F1FB" Margin="0,6,0,0"/>
                                    <TextBlock Text="Elected owners are read-only and come from the shared election plan." Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap" Margin="0,9,0,0"/>
                                </StackPanel>
                            </GroupBox>
                        </StackPanel>
                    </ScrollViewer>
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
                            <ColumnDefinition Width="150"/><ColumnDefinition Width="150"/><ColumnDefinition Width="180"/><ColumnDefinition Width="250"/><ColumnDefinition Width="150"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Grid.Column="0" Margin="0,0,8,0"><TextBlock Text="From"/><DatePicker x:Name="HistoryFromPicker"/></StackPanel>
                        <StackPanel Grid.Column="1" Margin="0,0,8,0"><TextBlock Text="To"/><DatePicker x:Name="HistoryToPicker"/></StackPanel>
                        <StackPanel Grid.Column="2" Margin="0,0,8,0"><TextBlock Text="Server"/><ComboBox x:Name="HistoryServerCombo"/></StackPanel>
                        <StackPanel Grid.Column="3" Margin="0,0,8,0"><TextBlock Text="Job"/><ComboBox x:Name="HistoryJobCombo" IsEditable="True"/></StackPanel>
                        <StackPanel Grid.Column="4" Margin="0,0,8,0"><TextBlock Text="Status"/><ComboBox x:Name="HistoryStatusCombo"/></StackPanel>
                        <Button x:Name="HistoryRefreshButton" Grid.Column="5" Content="Filter" VerticalAlignment="Bottom" Margin="4,0,4,8"/>
                        <Button x:Name="ExportCsvButton" Grid.Column="6" Content="Export CSV" VerticalAlignment="Bottom" Margin="4,0,4,8"/>
                        <Button x:Name="ExportHtmlButton" Grid.Column="7" Content="Export HTML" VerticalAlignment="Bottom" Margin="4,0,4,8"/>
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
            <Grid><TextBlock x:Name="FooterText" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center"/><TextBlock x:Name="VersionText" Text="v1.0.2" HorizontalAlignment="Right" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center"/></Grid>
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
$managementModulePath = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365.Orchestrator.Management.psm1'
Import-Module -Name $managementModulePath -Force -ErrorAction Stop

if ($ValidateOnly) {
    $validationWindow = ConvertFrom-OrchestratorGuiXaml -Text $xaml
    foreach ($controlName in @('PlanningGrid', 'HistoryGrid', 'ServersGrid', 'VersionsGrid', 'PublishButton', 'ApplyJobButton', 'ApplyServerButton', 'RollbackButton')) {
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
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')][string]$Level = 'INFO'
    )
    $line = '[{0}][{1}] {2}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Level, $Message
    if ($script:Controls -and $script:Controls.ActivityBox) {
        $script:Controls.ActivityBox.AppendText($line + [Environment]::NewLine)
        $script:Controls.ActivityBox.ScrollToEnd()
    }
    Microsoft.PowerShell.Utility\Write-Host $line
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
    if (-not (Test-Path -LiteralPath $planPath)) { return $owners }
    try {
        $plan = Read-SmartM365OrchestratorJson -Path $planPath
        foreach ($assignment in @($plan.Assignments)) { $owners[[string]$assignment.JobName] = [string]$assignment.OwnerServer }
    }
    catch { Write-GuiActivity -Message "Election plan could not be read: $($_.Exception.Message)" -Level WARN }
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
        }
    }
    $script:PlanningRows = @($rows)
    $script:Controls.PlanningGrid.ItemsSource = $script:PlanningRows
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

function Refresh-AllViews {
    try {
        $script:Snapshot = Get-SmartM365OrchestratorConfigurationSnapshot -SharedDataFolderPath $script:SharedDataFolderPath
        $script:DraftJobs = Copy-JsonDocument -Document $script:Snapshot.Jobs
        $script:DraftCluster = Copy-JsonDocument -Document $script:Snapshot.Cluster
        Refresh-PlanningView
        $servers = @(Refresh-ServersView)
        $history7 = @(Get-SmartM365OrchestratorHistory -SharedDataFolderPath $script:SharedDataFolderPath -From (Get-Date).AddDays(-7) -To (Get-Date))
        $script:Controls.JobsCountText.Text = [string]@($script:DraftJobs.Jobs).Count
        $script:Controls.EnabledCountText.Text = [string]@($script:DraftJobs.Jobs | Where-Object Enabled).Count
        $script:Controls.SuccessCountText.Text = [string]@($history7 | Where-Object Status -eq 'Success').Count
        $script:Controls.FailureCountText.Text = [string]@($history7 | Where-Object { $_.Status -in @('Failed', 'Timeout', 'Blocked') }).Count
        $script:Controls.VersionsGrid.ItemsSource = @(Get-SmartM365OrchestratorConfigurationVersions -SharedDataFolderPath $script:SharedDataFolderPath)
        $script:Controls.ConnectionText.Text = "Tenant: $Tenant"
        $script:Controls.LastRefreshText.Text = 'Refreshed: ' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        $script:Controls.StatusText.Text = 'Shared configuration loaded'
        Write-GuiActivity -Message ("Configuration loaded: {0} jobs, {1} servers." -f @($script:DraftJobs.Jobs).Count, @($servers).Count) -Level SUCCESS
    }
    catch {
        $script:Controls.StatusText.Text = 'Refresh failed'
        Write-GuiActivity -Message $_.Exception.Message -Level ERROR
        [System.Windows.MessageBox]::Show($_.Exception.Message, 'Refresh failed', 'OK', 'Error') | Out-Null
    }
}

function Apply-SelectedJobToDraft {
    $row = $script:Controls.PlanningGrid.SelectedItem
    if ($null -eq $row) { throw 'Select a job first.' }
    $job = @($script:DraftJobs.Jobs | Where-Object { [string]$_.Name -eq [string]$row.Name })[0]
    $scheduleType = Get-ComboText -Combo $script:Controls.ScheduleTypeCombo
    $times = @($script:Controls.TimesBox.Text -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    $days = @($script:Controls.DaysBox.Text -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | Sort-Object -Unique)
    $mode = Get-ComboText -Combo $script:Controls.AssignmentCombo
    $pinnedServer = Get-ComboText -Combo $script:Controls.PinnedServerCombo

    Set-JsonProperty -Object $job -Name Enabled -Value ([bool]$script:Controls.JobEnabledCheck.IsChecked)
    Set-JsonProperty -Object $job.Schedule -Name Type -Value $scheduleType
    Set-JsonProperty -Object $job.Schedule -Name Times -Value $times
    Set-JsonProperty -Object $job.Schedule -Name DaysOfWeek -Value $(if ($scheduleType -eq 'Weekly') { $days } else { @() })
    Set-JsonProperty -Object $job.Schedule -Name MissedRunPolicy -Value (Get-ComboText -Combo $script:Controls.MissedPolicyCombo)
    Set-JsonProperty -Object $job -Name AssignmentMode -Value $mode
    Set-JsonProperty -Object $job -Name AllowedServers -Value $(if ($mode -eq 'Pinned') { @($pinnedServer) } else { @() })
    Set-JsonProperty -Object $job -Name TimeoutMinutes -Value ([int]$script:Controls.TimeoutBox.Text)
    Set-JsonProperty -Object $job -Name MaxRetries -Value ([int]$script:Controls.RetriesBox.Text)
    Set-JsonProperty -Object $job -Name RetryDelaySeconds -Value ([int]$script:Controls.RetryDelayBox.Text)
    Set-JsonProperty -Object $job -Name EstimatedDurationMinutes -Value ([double]::Parse($script:Controls.DurationBox.Text, [System.Globalization.CultureInfo]::InvariantCulture))
    $validation = Test-SmartM365OrchestratorJobsDocument -Document $script:DraftJobs
    if (-not $validation.Valid) { throw $validation.Errors -join [Environment]::NewLine }
    Refresh-PlanningView
    $script:Controls.PlanningGrid.SelectedItem = @($script:PlanningRows | Where-Object Name -eq $job.Name)[0]
    $script:Controls.StatusText.Text = "Draft changed: $($job.Name)"
    Write-GuiActivity -Message "Job '$($job.Name)' updated in the draft."
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
        Write-GuiActivity -Message $_.Exception.Message -Level ERROR
        [System.Windows.MessageBox]::Show($_.Exception.Message, 'Publication failed', 'OK', 'Error') | Out-Null
    }
}

$smartM365Root = Split-Path -Path (Split-Path -Path $PSScriptRoot -Parent) -Parent
$tenantContextPath = Join-Path -Path $smartM365Root -ChildPath 'Config\SmartM365-TenantContext.ps1'
. $tenantContextPath
$script:EffectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$localConfigPath = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-Inventory-Orchestrator.local.json'
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

$bootstrapJobsPath = Join-Path -Path $PSScriptRoot -ChildPath 'Orchestrator-Jobs.json'
if (-not (Test-Path -LiteralPath $bootstrapJobsPath)) { $bootstrapJobsPath += '.template' }
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

$script:GuiSplash = $null
$splashPath = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365.GuiSplash.ps1'
if (-not $SmokeTest -and (Test-Path -LiteralPath $splashPath)) {
    . $splashPath
    $script:GuiSplash = Start-SmartM365GuiSplash `
        -ProductName 'SmartM365 Orchestrator' `
        -Subtitle 'Central planning and execution history' `
        -LogoPath (Join-Path $PSScriptRoot 'WorkplaceCloudHub-lockup-WPF.png') `
        -WindowIconPath (Join-Path $PSScriptRoot 'WorkplaceCloudHub.ico')
}

$window = ConvertFrom-OrchestratorGuiXaml -Text $xaml
$script:Controls = @{}
foreach ($name in @(
    'HeaderLogo', 'SharedPathText', 'ConnectionText', 'LastRefreshText', 'StatusText', 'RefreshButton', 'ValidateButton', 'PublishButton',
    'JobsCountText', 'EnabledCountText', 'OnlineServersText', 'SuccessCountText', 'FailureCountText', 'DashboardGrid',
    'PlanningGrid', 'SelectedJobText', 'JobEnabledCheck', 'ScheduleTypeCombo', 'TimesBox', 'DaysBox', 'MissedPolicyCombo',
    'AssignmentCombo', 'PinnedServerCombo', 'TimeoutBox', 'RetriesBox', 'RetryDelayBox', 'DurationBox', 'ApplyJobButton',
    'HistoryFromPicker', 'HistoryToPicker', 'HistoryServerCombo', 'HistoryJobCombo', 'HistoryStatusCombo', 'HistoryRefreshButton',
    'ExportCsvButton', 'ExportHtmlButton', 'HistoryGrid', 'ServersGrid', 'NewServerBox', 'AddServerButton', 'RemoveServerButton',
    'SelectedServerText', 'ServerWeightBox', 'ServerPolicyCombo', 'ApplyServerButton', 'VersionsGrid', 'RollbackButton',
    'ActivityBox', 'FooterText', 'VersionText'
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
$script:Controls.FooterText.Text = 'Shared changes are validated, versioned and audited. This GUI does not start or stop jobs.'
$script:Controls.VersionText.Text = "v$($script:AppVersion)"
$script:Controls.HistoryFromPicker.SelectedDate = (Get-Date).AddDays(-7).Date
$script:Controls.HistoryToPicker.SelectedDate = (Get-Date).Date
$script:Controls.HistoryStatusCombo.ItemsSource = @('All', 'Success', 'Failed', 'Timeout', 'Blocked', 'Running')
$script:Controls.HistoryStatusCombo.SelectedIndex = 0

$script:Controls.PlanningGrid.Add_SelectionChanged({
    $row = $script:Controls.PlanningGrid.SelectedItem
    if ($null -eq $row) { return }
    $job = @($script:DraftJobs.Jobs | Where-Object Name -eq $row.Name)[0]
    $script:Controls.SelectedJobText.Text = [string]$job.Name
    $script:Controls.JobEnabledCheck.IsChecked = [bool]$job.Enabled
    Select-ComboText -Combo $script:Controls.ScheduleTypeCombo -Text ([string]$job.Schedule.Type)
    $script:Controls.TimesBox.Text = @($job.Schedule.Times) -join ', '
    $script:Controls.DaysBox.Text = @(Get-OrchestratorGuiPropertyValue -Object $job.Schedule -Name 'DaysOfWeek' -DefaultValue @()) -join ', '
    Select-ComboText -Combo $script:Controls.MissedPolicyCombo -Text ([string](Get-OrchestratorGuiPropertyValue -Object $job.Schedule -Name 'MissedRunPolicy' -DefaultValue 'RunOnce'))
    Select-ComboText -Combo $script:Controls.AssignmentCombo -Text ([string](Get-OrchestratorGuiPropertyValue -Object $job -Name 'AssignmentMode' -DefaultValue 'Legacy'))
    $script:Controls.PinnedServerCombo.Text = @(Get-OrchestratorGuiPropertyValue -Object $job -Name 'AllowedServers' -DefaultValue @()) -join ', '
    $script:Controls.TimeoutBox.Text = [string](Get-OrchestratorGuiPropertyValue -Object $job -Name 'TimeoutMinutes' -DefaultValue 240)
    $script:Controls.RetriesBox.Text = [string](Get-OrchestratorGuiPropertyValue -Object $job -Name 'MaxRetries' -DefaultValue 0)
    $script:Controls.RetryDelayBox.Text = [string](Get-OrchestratorGuiPropertyValue -Object $job -Name 'RetryDelaySeconds' -DefaultValue 300)
    $script:Controls.DurationBox.Text = ([double](Get-OrchestratorGuiPropertyValue -Object $job -Name 'EstimatedDurationMinutes' -DefaultValue 5)).ToString([System.Globalization.CultureInfo]::InvariantCulture)
})
$script:Controls.ServersGrid.Add_SelectionChanged({
    $row = $script:Controls.ServersGrid.SelectedItem
    if ($null -eq $row) { return }
    $script:Controls.SelectedServerText.Text = [string]$row.Server
    $script:Controls.ServerWeightBox.Text = ([double]$row.Weight).ToString([System.Globalization.CultureInfo]::InvariantCulture)
    Select-ComboText -Combo $script:Controls.ServerPolicyCombo -Text ([string]$row.Policy)
})
$script:Controls.ApplyJobButton.Add_Click({ try { Apply-SelectedJobToDraft } catch { [System.Windows.MessageBox]::Show($_.Exception.Message, 'Invalid job', 'OK', 'Error') | Out-Null } })
$script:Controls.ApplyServerButton.Add_Click({ try { Apply-SelectedServerToDraft } catch { [System.Windows.MessageBox]::Show($_.Exception.Message, 'Invalid server', 'OK', 'Error') | Out-Null } })
$script:Controls.RefreshButton.Add_Click({ Refresh-AllViews })
$script:Controls.ValidateButton.Add_Click({ [void](Show-DraftValidation) })
$script:Controls.PublishButton.Add_Click({ Publish-Draft })
$script:Controls.HistoryRefreshButton.Add_Click({ Refresh-HistoryView })
$script:Controls.HistoryGrid.Add_MouseDoubleClick({
    $row = $script:Controls.HistoryGrid.SelectedItem
    if ($null -ne $row -and (Test-Path -LiteralPath $row.LogPath)) { Start-Process -FilePath $row.LogPath }
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
    catch { [System.Windows.MessageBox]::Show($_.Exception.Message, 'Add server', 'OK', 'Error') | Out-Null }
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
    catch { [System.Windows.MessageBox]::Show($_.Exception.Message, 'Rollback failed', 'OK', 'Error') | Out-Null }
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
$script:Controls.HistoryServerCombo.ItemsSource = @('All') + @($script:DraftCluster.ExpectedOrchestratorServers)
$script:Controls.HistoryServerCombo.SelectedIndex = 0
$script:Controls.HistoryJobCombo.ItemsSource = @('All') + @($script:DraftJobs.Jobs.Name | Sort-Object)
$script:Controls.HistoryJobCombo.SelectedIndex = 0
Refresh-HistoryView
if ($SmokeTest) {
    "[{0}] SMOKE_TEST_OK SmartM365 Orchestrator GUI v{1} | Jobs={2} | PlanningRows={3}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $script:AppVersion, @($script:DraftJobs.Jobs).Count, @($script:PlanningRows).Count
    return
}

$window.Add_ContentRendered({
    if ($script:GuiSplash) {
        Hide-SmartM365GuiSplash -Splash $script:GuiSplash
        $window.Activate() | Out-Null
    }
})
$window.Add_Closed({
    if ($script:GuiSplash) { Close-SmartM365GuiSplash -Splash $script:GuiSplash }
})
[void]$window.ShowDialog()

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCACJsEdX5O6cmeD
# 0b6k/0jUh9kBq1MJVwlrsz+9/oMLUKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIJu4wXReT4kcOmC8dT52sNyguABCB4t/0/PHXJuKZ8AgMA0GCSqG
# SIb3DQEBAQUABIIBgIlFRZWvZFCVyjEaRO9OuqXvb8PHQBeCFy/oRhf8L5WAtVKD
# mbsOE4XxC5M8VOJtq65eFaP5+wvisBwxXu7IM/xCHS+TtkS605/+Lxqs0a+4b72u
# MlaQEhe+bwBXQ2vF2WOB7jIfHTT9XmJdxOExTURwFJdjLfdlQFvzVhZ/LHSwmPqu
# HCGEx30r9tQqcz4C+DmUGMvrjuVnhEKQPbfSWShpTn2wGbgR+iPqwNI+S7W2rMqg
# VxVClOKN2x/w4hOtnnxRFCGjB1SjFvqyMNhbGT47pDMWJTWIr5m0jqS82dwYekEz
# rswltsk5gkFGc0upfDwxfL1tSGb8VbHAkP3IwI15q54PLlIaDEFx699uogpi/kwg
# OPl0QEAeL4oRiOZWRWC7uIAe/p3f5ZC9w3kQmioDRqTm2FnQwYqdp+Ex1LBGcj7e
# 60hm5O6wZHcXVTuLmUDz8zbULbslm8+weoO5RHqmBxdBd4dFA7MjxZn6ef4hIAFT
# kKfjFFCRlBpRVBVG5KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkyMDAw
# MTFaMC8GCSqGSIb3DQEJBDEiBCAiNHcdhKvBXJIvGkPtAenGT231irTEdqe6+bgR
# zNJV4jANBgkqhkiG9w0BAQEFAASCAgCUFb9254cQ/5fMcG7mT4T0X2MmZox2RHU+
# uuaxivDyRH+tt0FgTGr4If7KRvOVFNrVXFTV47MeDid0N/IJ3RQiVZJqSJVZc25w
# jvgNykxC3wrzu2+CXhfyGxpT504+XEyDnZu8oARdENTyM2ZBs22Emfftk3qW+Hhu
# HZ4GGb2diJKnECjI/vycIMxOPjSGmcYhYkqJIRtMco4Zy+cvfxdn8w1NT4mdIUbb
# CK+MTxtYhQRF6s9zZZ/qUSofxEwRttEN8V6Axee+gC45jBT/zwVkFYQd2zyhuJkk
# az/CjwXP8W4ukETIVAKPgc3ZUjX2cIrZTe1JszKEbCc2jbG82FEgI1D52Hew5pTH
# pdAF19C1dqjCRxBwVnID75OM5QGgI/RZi+jNFcBYbeHSB7JopVZPZA8YCmcxR0bH
# 0LbwNAbPnx/q34618ZcDCog5Zj+KGaGAkxI4XlptSBFAWqz+pXSxsnIeaBCNDDeH
# D7cB8QRy+ah+2mm4yz2OTFzNymBIUGh6Ghko6PeAthIBa/cEqpX4hLLmJKIozyVm
# C1SpvjaX5piNF/gIeCKOY8Xrro6HyocwCX2RYk1DXF+eGXHMiX6GaMVgE1XEV42B
# r+dRfWKKK/tbfaCZbqtvELzjo0mz7DGpJseR9IJ/04kA8At9xxY158Sc6qvi/H11
# q0yMUA9VZQ==
# SIG # End signature block
