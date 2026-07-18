<#
.SYNOPSIS
Interactive read-only preflight application for Exchange hybrid migration batches.

.VERSION
1.4.0
#>
#requires -Version 7.0

[CmdletBinding()]
param(
    [string]$CsvPath = '',
    [string]$ConfigPath = '',
    [switch]$NoSplash,
    [switch]$ValidateOnly
)

$ErrorActionPreference = 'Stop'
trap {
    $message = $_.Exception.Message
    try {
        Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
        [void][System.Windows.MessageBox]::Show($message, 'Smart Exchange Migration Readiness - startup error', 'OK', 'Error')
    }
    catch {}
    if ($script:SessionLogPath) {
        try { [IO.File]::AppendAllText($script:SessionLogPath, "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [ERROR] Startup failure: $message`r`n", [Text.UTF8Encoding]::new($false)) } catch {}
    }
    exit 1
}
$script:AppVersion = '1.4.0'
$script:Batch = $null
$script:Assessment = $null
$script:Export = $null
$script:CancelRequested = $false
$script:IsBusy = $false
$script:SessionLogPath = ''
$script:SessionLogError = ''
$script:MigrationEndpointsLoaded = $false
$script:MigrationEndpointSelectionConfirmed = $false
$script:UpdatingMigrationEndpointCombo = $false

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase

$readinessModulePath = Join-Path $PSScriptRoot 'SmartM365.ExchangeMigrationReadiness.psm1'
Import-Module -Name $readinessModulePath -Force -ErrorAction Stop

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Smart Exchange Migration Readiness"
        Width="1440" Height="900" MinWidth="1180" MinHeight="720"
        WindowStartupLocation="CenterScreen"
        Background="#F5F8FB"
        FontFamily="Segoe UI"
        UseLayoutRounding="True"
        SnapsToDevicePixels="True">
    <Window.Resources>
        <SolidColorBrush x:Key="AccentBrush" Color="#0078D4"/>
        <SolidColorBrush x:Key="AccentDarkBrush" Color="#005A9E"/>
        <SolidColorBrush x:Key="AccentSoftBrush" Color="#E6F4FF"/>
        <SolidColorBrush x:Key="PanelBrush" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="TextBrush" Color="#1F2937"/>
        <SolidColorBrush x:Key="MutedBrush" Color="#5F6B7A"/>
        <SolidColorBrush x:Key="BorderBrushSoft" Color="#DDE7F0"/>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Height" Value="34"/>
            <Setter Property="Padding" Value="14,0"/>
            <Setter Property="Margin" Value="0,0,8,0"/>
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrushSoft}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="ButtonBorder" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="8" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Background" Value="#F8FBFE"/>
                                <Setter TargetName="ButtonBorder" Property="BorderBrush" Value="#B9DDF7"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.45"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Height" Value="34"/>
            <Setter Property="Padding" Value="9,5"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrushSoft}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Height" Value="34"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrushSoft}"/>
            <Setter Property="Background" Value="#FFFFFF"/>
        </Style>
        <Style TargetType="TabControl">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
        </Style>
        <Style TargetType="TabItem">
            <Setter Property="Foreground" Value="{StaticResource MutedBrush}"/>
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrushSoft}"/>
            <Setter Property="Padding" Value="14,7"/>
            <Setter Property="Margin" Value="0,0,6,0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="TabBorder" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="8" Padding="{TemplateBinding Padding}">
                            <ContentPresenter ContentSource="Header" HorizontalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="{StaticResource AccentSoftBrush}"/>
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#B9DDF7"/>
                                <Setter Property="Foreground" Value="{StaticResource AccentDarkBrush}"/>
                            </Trigger>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="#F8FBFE"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrushSoft}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="GridLinesVisibility" Value="Horizontal"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#EDF2F7"/>
            <Setter Property="RowHeaderWidth" Value="0"/>
            <Setter Property="AutoGenerateColumns" Value="True"/>
            <Setter Property="CanUserAddRows" Value="False"/>
            <Setter Property="CanUserDeleteRows" Value="False"/>
            <Setter Property="IsReadOnly" Value="True"/>
            <Setter Property="AlternatingRowBackground" Value="#F8FBFE"/>
            <Setter Property="HeadersVisibility" Value="Column"/>
        </Style>
    </Window.Resources>

    <Grid Margin="18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="18">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="150"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                    <Border HorizontalAlignment="Left" CornerRadius="12" Padding="10,4" Background="{StaticResource AccentSoftBrush}">
                        <TextBlock Text="EXCHANGE HYBRID PREFLIGHT" Foreground="{StaticResource AccentDarkBrush}" FontSize="11" FontWeight="SemiBold"/>
                    </Border>
                    <TextBlock Text="Smart Exchange Migration Readiness" FontSize="25" FontWeight="SemiBold" Margin="0,9,0,3"/>
                    <TextBlock Text="Read-only GO / NO-GO assessment for Exchange Online migration batches" Foreground="{StaticResource MutedBrush}" FontSize="13"/>
                </StackPanel>
                <Border x:Name="HeaderLogoLink" Grid.Column="1" Width="132" Height="82" CornerRadius="8" Background="#FFFFFF" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" Padding="5" Cursor="Hand" ToolTip="Open WorkplaceCloudHub.com">
                    <Image x:Name="HeaderLogo" Stretch="Uniform" RenderOptions.BitmapScalingMode="HighQuality"/>
                </Border>
            </Grid>
        </Border>

        <Border Grid.Row="1" Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="12" Margin="0,12,0,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                    <Border Width="10" Height="10" CornerRadius="5" Background="{StaticResource AccentBrush}" Margin="0,0,10,0"/>
                    <TextBlock x:Name="StatusText" Text="Select a migration batch CSV." FontWeight="SemiBold" VerticalAlignment="Center"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal">
                    <TextBlock Text="Mode" VerticalAlignment="Center" Foreground="{StaticResource MutedBrush}" Margin="0,0,6,0"/>
                    <ComboBox x:Name="ModeCombo" Width="118" Margin="0,0,10,0" ToolTip="Live prefers live sources and falls back to CSV for AD, Exchange on-premises, and Entra Connect. CacheOnly uses CSV inventories only.">
                        <ComboBoxItem Content="Live" IsSelected="True"/>
                        <ComboBoxItem Content="CacheOnly"/>
                    </ComboBox>
                    <TextBlock Text="Auth" VerticalAlignment="Center" Foreground="{StaticResource MutedBrush}" Margin="0,0,6,0"/>
                    <ComboBox x:Name="AuthenticationModeCombo" Width="122" Margin="0,0,10,0" ToolTip="Interactive opens the standard sign-in windows. Application uses the configured app registration and certificate.">
                        <ComboBoxItem Content="Interactive" IsSelected="True"/>
                        <ComboBoxItem Content="Application"/>
                    </ComboBox>
                    <TextBlock Text="Endpoint" VerticalAlignment="Center" Foreground="{StaticResource MutedBrush}" Margin="0,0,6,0"/>
                    <ComboBox x:Name="MigrationEndpointCombo" Width="180" Margin="0,0,10,0" IsEditable="True" IsReadOnly="True" IsEnabled="False" ToolTip="ExchangeRemoteMove endpoints are loaded automatically after Exchange Online authentication."/>
                    <Button x:Name="RunButton" Content="Run assessment" Style="{StaticResource PrimaryButton}" IsEnabled="False"/>
                    <Button x:Name="CancelButton" Content="Cancel" IsEnabled="False"/>
                    <Button x:Name="OpenOutputButton" Content="Open output" IsEnabled="False"/>
                </StackPanel>
            </Grid>
        </Border>

        <TabControl x:Name="MainTabs" Grid.Row="2">
            <TabItem Header="Batch">
                <Grid Margin="0,12,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Border Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="14">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel>
                                <TextBlock Text="Migration batch CSV" FontWeight="SemiBold" Margin="0,0,0,6"/>
                                <TextBox x:Name="CsvPathBox"/>
                                <TextBlock x:Name="CsvMetadataText" Text="No CSV loaded." Foreground="{StaticResource MutedBrush}" Margin="0,6,0,0"/>
                            </StackPanel>
                            <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Bottom" Margin="12,0,0,22">
                                <Button x:Name="BrowseCsvButton" Content="Browse" Style="{StaticResource PrimaryButton}"/>
                            </StackPanel>
                        </Grid>
                    </Border>
                    <DataGrid x:Name="BatchGrid" Grid.Row="1" Margin="0,12,0,0"/>
                </Grid>
            </TabItem>

            <TabItem Header="Options">
                <Grid Margin="0,12,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Border Background="White" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="12" Margin="0,0,0,12">
                        <Grid>
                            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
                            <StackPanel>
                                <TextBlock Text="Readiness checks" FontWeight="SemiBold" FontSize="15"/>
                                <TextBlock x:Name="OptionsSummaryText" Text="All optional checks are enabled." Foreground="{StaticResource MutedBrush}" Margin="0,5,0,0"/>
                                <TextBlock Text="Mandatory source and CSV-integrity checks cannot be disabled. Changes apply to the next run only." Foreground="{StaticResource MutedBrush}" Margin="0,3,0,0" TextWrapping="Wrap"/>
                            </StackPanel>
                            <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                                <Button x:Name="EnableAllOptionsButton" Content="Enable all"/>
                                <Button x:Name="ResetOptionsButton" Content="Reset defaults" Margin="0"/>
                            </StackPanel>
                        </Grid>
                    </Border>
                    <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                        <ItemsControl x:Name="OptionsItemsControl">
                            <ItemsControl.ItemTemplate>
                                <DataTemplate>
                                    <Border Background="White" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="6" Padding="10" Margin="0,0,0,6">
                                        <Grid>
                                            <Grid.ColumnDefinitions><ColumnDefinition Width="35"/><ColumnDefinition Width="155"/><ColumnDefinition Width="280"/><ColumnDefinition Width="*"/><ColumnDefinition Width="210"/></Grid.ColumnDefinitions>
                                            <CheckBox IsChecked="{Binding Enabled, Mode=TwoWay, UpdateSourceTrigger=PropertyChanged}" IsEnabled="{Binding CanDisable}" VerticalAlignment="Center"/>
                                            <TextBlock Grid.Column="1" Text="{Binding Category}" VerticalAlignment="Center" Foreground="{StaticResource MutedBrush}"/>
                                            <TextBlock Grid.Column="2" Text="{Binding Name}" VerticalAlignment="Center" FontWeight="SemiBold" TextWrapping="Wrap" Margin="8,0"/>
                                            <TextBlock Grid.Column="3" Text="{Binding Description}" VerticalAlignment="Center" TextWrapping="Wrap" Margin="8,0"/>
                                            <TextBlock Grid.Column="4" Text="{Binding CheckId}" VerticalAlignment="Center" FontFamily="Consolas" Foreground="{StaticResource MutedBrush}" Margin="8,0,0,0"/>
                                        </Grid>
                                    </Border>
                                </DataTemplate>
                            </ItemsControl.ItemTemplate>
                        </ItemsControl>
                    </ScrollViewer>
                </Grid>
            </TabItem>

            <TabItem Header="Sources">
                <Grid Margin="0,12,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Border Grid.Row="0" Grid.Column="0" Background="White" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,0,6,8">
                        <StackPanel>
                            <TextBlock Text="Active Directory" FontWeight="SemiBold" FontSize="15"/>
                            <TextBlock x:Name="AdStateText" Text="Checked automatically during assessment" Foreground="{StaticResource MutedBrush}" Margin="0,5,0,0"/>
                        </StackPanel>
                    </Border>
                    <Border Grid.Row="0" Grid.Column="1" Background="White" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="6,0,0,8">
                        <StackPanel>
                            <TextBlock Text="Exchange on-premises" FontWeight="SemiBold" FontSize="15"/>
                            <TextBlock x:Name="OnPremStateText" Text="Checked automatically during assessment" Foreground="{StaticResource MutedBrush}" Margin="0,5,0,0"/>
                        </StackPanel>
                    </Border>
                    <Border Grid.Row="1" Grid.Column="0" Background="White" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,8,6,8">
                        <StackPanel>
                            <TextBlock Text="Exchange Online" FontWeight="SemiBold" FontSize="15"/>
                            <TextBlock x:Name="ExoStateText" Text="Interactive connection starts with the assessment" Foreground="{StaticResource MutedBrush}" Margin="0,5,0,0"/>
                        </StackPanel>
                    </Border>
                    <Border Grid.Row="1" Grid.Column="1" Background="White" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="6,8,0,8">
                        <StackPanel>
                            <TextBlock Text="Microsoft Graph" FontWeight="SemiBold" FontSize="15"/>
                            <TextBlock x:Name="GraphStateText" Text="Interactive connection starts with the assessment" Foreground="{StaticResource MutedBrush}" Margin="0,5,0,0"/>
                        </StackPanel>
                    </Border>
                    <Border Grid.Row="2" Grid.ColumnSpan="2" Background="White" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,8,0,0">
                        <StackPanel>
                            <TextBlock Text="Hybrid migration and Microsoft Entra Connect" FontWeight="SemiBold" FontSize="15"/>
                            <TextBlock Text="Checks the Exchange Online migration endpoint and Entra Connect synchronization health automatically during the assessment." Foreground="{StaticResource MutedBrush}" Margin="0,5,0,0" TextWrapping="Wrap"/>
                            <TextBlock x:Name="HybridStateText" Text="Pending assessment" Foreground="{StaticResource MutedBrush}" Margin="0,5,0,0" TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                </Grid>
            </TabItem>

            <TabItem x:Name="CsvSourcesTab" Header="CSV sources">
                <Grid Margin="0,12,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Border Background="White" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="12" Margin="0,0,0,12">
                        <TextBlock Text="CSV files expected by the selected mode. Used identifies the effective source; Present and Fresh describe the selected file from DATA-LAST or the tenant cache root." Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap"/>
                    </Border>
                    <DataGrid x:Name="CsvSourcesGrid" Grid.Row="1"/>
                </Grid>
            </TabItem>

            <TabItem x:Name="ResultsTab" Header="Results">
                <Grid Margin="0,12,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <UniformGrid Rows="1" Columns="4" Margin="0,0,0,12">
                        <Border Background="#ECFDF3" BorderBrush="#A7E3BC" BorderThickness="1" CornerRadius="8" Padding="12" Margin="0,0,6,0">
                            <StackPanel><TextBlock Text="GO" Foreground="#146C43"/><TextBlock x:Name="GoCountText" Text="0" FontSize="24" FontWeight="SemiBold" Foreground="#146C43"/></StackPanel>
                        </Border>
                        <Border Background="#FFF8E7" BorderBrush="#F3D48A" BorderThickness="1" CornerRadius="8" Padding="12" Margin="6,0">
                            <StackPanel><TextBlock Text="GO-WARNING" Foreground="#8A5A00"/><TextBlock x:Name="WarningCountText" Text="0" FontSize="24" FontWeight="SemiBold" Foreground="#8A5A00"/></StackPanel>
                        </Border>
                        <Border Background="#FFF0F0" BorderBrush="#F0B3B3" BorderThickness="1" CornerRadius="8" Padding="12" Margin="6,0">
                            <StackPanel><TextBlock Text="NO-GO" Foreground="#B42318"/><TextBlock x:Name="NoGoCountText" Text="0" FontSize="24" FontWeight="SemiBold" Foreground="#B42318"/></StackPanel>
                        </Border>
                        <Border Background="#F3F4F6" BorderBrush="#D1D5DB" BorderThickness="1" CornerRadius="8" Padding="12" Margin="6,0,0,0">
                            <StackPanel><TextBlock Text="UNKNOWN" Foreground="#4B5563"/><TextBlock x:Name="UnknownCountText" Text="0" FontSize="24" FontWeight="SemiBold" Foreground="#4B5563"/></StackPanel>
                        </Border>
                    </UniformGrid>
                    <DataGrid x:Name="SummaryGrid" Grid.Row="1"/>
                </Grid>
            </TabItem>

            <TabItem Header="Mailbox details">
                <Grid Margin="0,12,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Border Background="White" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="12" Margin="0,0,0,12">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="Mailbox filter" VerticalAlignment="Center" Margin="0,0,10,0"/>
                            <TextBox x:Name="MailboxFilterBox" Width="340"/>
                            <Button x:Name="ClearFilterButton" Content="Clear" Margin="8,0,0,0"/>
                        </StackPanel>
                    </Border>
                    <DataGrid x:Name="FindingsGrid" Grid.Row="1"/>
                </Grid>
            </TabItem>

            <TabItem Header="Permissions baseline">
                <Grid Margin="0,12,0,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <Border Background="White" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="12" Margin="0,0,0,12">
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <TextBlock Text="Explicit Full Access, Send As and Send on Behalf grants captured before migration." VerticalAlignment="Center"/>
                            <Button x:Name="ComparePermissionsButton" Grid.Column="1" Content="Collect EXO and compare" IsEnabled="False"/>
                        </Grid>
                    </Border>
                    <DataGrid x:Name="PermissionsGrid" Grid.Row="1"/>
                </Grid>
            </TabItem>

            <TabItem Header="Activity">
                <Grid Margin="0,12,0,0">
                    <TextBox x:Name="ActivityBox" Height="Auto" VerticalAlignment="Stretch" VerticalContentAlignment="Top" FontFamily="Consolas" FontSize="12" IsReadOnly="True" AcceptsReturn="True" TextWrapping="NoWrap" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
                </Grid>
            </TabItem>
        </TabControl>

        <Border Grid.Row="3" Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="10" Margin="0,12,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="360"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <TextBlock x:Name="FooterText" Text="Read-only mode - no tenant or directory changes" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center"/>
                <ProgressBar x:Name="RunProgress" Grid.Column="1" Height="12" Minimum="0" Maximum="100" Value="0" Margin="12,0"/>
                <TextBlock x:Name="VersionText" Grid.Column="2" Text="v1.4.0" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center"/>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

function ConvertFrom-SemrXaml {
    param([Parameter(Mandatory)][string]$Text)
    $reader = [System.Xml.XmlNodeReader]::new([xml]$Text)
    return [System.Windows.Markup.XamlReader]::Load($reader)
}

if ($ValidateOnly) {
    $validationWindow = ConvertFrom-SemrXaml -Text $xaml
    $requiredControls = @(
        'CsvPathBox', 'BrowseCsvButton', 'ModeCombo', 'AuthenticationModeCombo', 'MigrationEndpointCombo', 'RunButton', 'SummaryGrid',
        'FindingsGrid', 'PermissionsGrid', 'CsvSourcesGrid', 'OptionsItemsControl', 'ResultsTab', 'ActivityBox'
    )
    foreach ($controlName in $requiredControls) {
        if (-not $validationWindow.FindName($controlName)) {
            throw "Required XAML control not found: $controlName"
        }
    }
    $templatePath = Join-Path $PSScriptRoot 'Config\SmartM365-ExchangeMigrationReadiness.local.json.template'
    Get-Content -LiteralPath $templatePath -Raw | ConvertFrom-Json | Out-Null
    "VALIDATION_OK Smart Exchange Migration Readiness v$script:AppVersion"
    return
}

if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $ConfigPath = Join-Path $PSScriptRoot 'Config\SmartM365-ExchangeMigrationReadiness.local.json'
}
$script:Config = Get-SemrConfig -Path $ConfigPath
$script:ConfiguredDisabledChecks = if ($script:Config.Contains('DisabledChecks')) { @($script:Config.DisabledChecks) } else { @() }
$script:CheckOptions = @(Get-SemrCheckCatalog | ForEach-Object {
    [pscustomobject][ordered]@{
        Enabled = [bool]($_.Mandatory -or $script:ConfiguredDisabledChecks -notcontains $_.CheckId)
        CanDisable = [bool]$_.CanDisable
        Mandatory = [bool]$_.Mandatory
        Category = [string]$_.Category
        Name = [string]$_.Name
        Description = [string]$_.Description
        CheckId = [string]$_.CheckId
    }
})

$script:GuiSplash = $null
if (-not $NoSplash) {
    $splashPath = Join-Path $PSScriptRoot 'SmartM365.GuiSplash.ps1'
    if (Test-Path -LiteralPath $splashPath) {
        . $splashPath
        $script:GuiSplash = Start-SmartM365GuiSplash `
            -ProductName 'Smart Exchange Migration Readiness' `
            -Subtitle 'Exchange Online migration batch preflight' `
            -LogoPath (Join-Path $PSScriptRoot 'WorkplaceCloudHub-lockup-WPF.png') `
            -WindowIconPath (Join-Path $PSScriptRoot 'WorkplaceCloudHub.ico')
    }
}

$window = ConvertFrom-SemrXaml -Text $xaml
$controls = @{}
foreach ($name in @(
    'HeaderLogoLink', 'HeaderLogo', 'StatusText', 'ModeCombo', 'AuthenticationModeCombo', 'MigrationEndpointCombo', 'RunButton', 'CancelButton', 'OpenOutputButton',
    'CsvPathBox', 'CsvMetadataText', 'BrowseCsvButton', 'BatchGrid',
    'AdStateText', 'OnPremStateText', 'ExoStateText', 'GraphStateText', 'HybridStateText',
    'GoCountText', 'WarningCountText', 'NoGoCountText', 'UnknownCountText', 'SummaryGrid',
    'MailboxFilterBox', 'ClearFilterButton', 'FindingsGrid', 'PermissionsGrid', 'ComparePermissionsButton',
    'CsvSourcesGrid', 'CsvSourcesTab', 'OptionsItemsControl', 'OptionsSummaryText', 'EnableAllOptionsButton', 'ResetOptionsButton',
    'ResultsTab', 'ActivityBox', 'FooterText', 'RunProgress', 'VersionText', 'MainTabs'
)) {
    $controls[$name] = $window.FindName($name)
}

$iconPath = Join-Path $PSScriptRoot 'WorkplaceCloudHub.ico'
if (Test-Path -LiteralPath $iconPath) {
    $window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]$iconPath)
}
$logoPath = Join-Path $PSScriptRoot 'WorkplaceCloudHub-lockup-WPF.png'
if (Test-Path -LiteralPath $logoPath) {
    $bitmap = [System.Windows.Media.Imaging.BitmapImage]::new()
    $bitmap.BeginInit()
    $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $bitmap.UriSource = [Uri]$logoPath
    $bitmap.EndInit()
    $bitmap.Freeze()
    $controls.HeaderLogo.Source = $bitmap
}
$controls.VersionText.Text = "v$script:AppVersion"
$controls.OptionsItemsControl.ItemsSource = $script:CheckOptions
$controls.ModeCombo.SelectedIndex = if ([string]$script:Config.Mode -eq 'CacheOnly') { 1 } else { 0 }
$controls.AuthenticationModeCombo.SelectedIndex = if ([string]$script:Config._AuthenticationMode -eq 'Application') { 1 } else { 0 }
$configuredEndpointName = [string]$script:Config.Hybrid.MigrationEndpointName
$controls.MigrationEndpointCombo.Text = if ($configuredEndpointName) { $configuredEndpointName } else { 'Loaded after EXO connection' }
if ($CsvPath) { $controls.CsvPathBox.Text = $CsvPath }

function Write-SemrActivity {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'SUCCESS')][string]$Level = 'INFO'
    )
    $line = "{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    $controls.ActivityBox.AppendText("$line`r`n")
    $controls.ActivityBox.ScrollToEnd()
    if ($script:SessionLogPath) {
        try { [IO.File]::AppendAllText($script:SessionLogPath, "$line`r`n", [Text.UTF8Encoding]::new($false)) } catch {}
    }
}

function Update-SemrOptionsSummary {
    $disabled = @($script:CheckOptions | Where-Object { $_.CanDisable -and -not $_.Enabled })
    $enabledCount = @($script:CheckOptions | Where-Object Enabled).Count
    $controls.OptionsSummaryText.Text = "$enabledCount enabled; $($disabled.Count) optional check(s) disabled."
}

function Sync-SemrCheckOptionsToConfig {
    foreach ($option in $script:CheckOptions) { if ($option.Mandatory) { $option.Enabled = $true } }
    $script:Config['DisabledChecks'] = @($script:CheckOptions | Where-Object { $_.CanDisable -and -not $_.Enabled } | ForEach-Object CheckId)
    Update-SemrOptionsSummary
}

function Refresh-SemrOptionsDisplay {
    $controls.OptionsItemsControl.ItemsSource = $null
    $controls.OptionsItemsControl.ItemsSource = $script:CheckOptions
    Update-SemrOptionsSummary
}

function Update-SemrCsvSourcesDisplay {
    param(
        [switch]$AssessmentCompleted,
        [AllowNull()]$EntraConnect = $null
    )

    try {
        $controls.CsvSourcesGrid.ItemsSource = @(Get-SemrCsvSourceInventory -Config $script:Config -Batch $script:Batch -EntraConnect $EntraConnect -AssessmentCompleted:$AssessmentCompleted)
    }
    catch {
        Write-SemrActivity -Message "CSV source inventory could not be refreshed: $($_.Exception.Message)" -Level WARN
    }
}

function Confirm-SemrStaleCsvUsage {
    $script:Config['_AllowStaleCache'] = $false
    $sources = @(Get-SemrCsvSourceInventory -Config $script:Config -Batch $script:Batch)
    $mode = Get-SemrSelectedMode
    $staleSources = @($sources | Where-Object {
        $_.Present -and -not $_.Fresh -and $_.Status -ne 'Not used - related checks disabled' -and (
            $mode -eq 'CacheOnly' -or $_.ExpectedUse -eq 'Live fallback / CacheOnly'
        )
    })
    if ($staleSources.Count -eq 0) { return $true }

    $maximumAgeHours = [double]$script:Config.Cache.MaximumAgeHours
    $fileList = ($staleSources | ForEach-Object { "- $($_.FileName): $($_.AgeHours) hour(s) old`n  $($_.Path)" }) -join "`n"
    $modeExplanation = if ($mode -eq 'Live') { 'These files may be needed if a live on-premises source is unavailable.' } else { 'CacheOnly requires these files for the assessment.' }
    $message = @"
The following CSV file(s) exceed MaximumAgeHours ($maximumAgeHours):

$fileList

$modeExplanation

Do you want to continue and explicitly accept these stale files for this run?
"@
    $answer = [System.Windows.MessageBox]::Show($window, $message, 'Stale CSV cache detected', 'YesNo', 'Warning')
    if ($answer -eq [System.Windows.MessageBoxResult]::Yes) {
        $script:Config['_AllowStaleCache'] = $true
        Write-SemrActivity -Message "Operator accepted $($staleSources.Count) stale CSV file(s) for this run." -Level WARN
        Update-SemrCsvSourcesDisplay
        return $true
    }

    $controls.StatusText.Text = 'Assessment cancelled: stale CSV files were not accepted.'
    Write-SemrActivity -Message 'Assessment cancelled because stale CSV files were not accepted.' -Level WARN
    return $false
}

function Confirm-SemrMicrosoftGraphModule {
    try {
        $moduleState = Get-SemrMicrosoftGraphModuleState
    }
    catch {
        Show-SemrError -Title 'Microsoft Graph module check failed' -ErrorRecord $_
        return $false
    }
    if ($moduleState.Available) { return $true }

    $missingList = (@($moduleState.MissingModules) | ForEach-Object { "- $_" }) -join "`n"
    $message = @"
The following Microsoft Graph module(s) are missing from PowerShell 7:

$missingList

Do you want to install them from PowerShell Gallery for the current user?

This installation does not change the tenant and does not require administrator rights.
"@
    $answer = [System.Windows.MessageBox]::Show(
        $window,
        $message,
        'Install Microsoft Graph modules',
        [System.Windows.MessageBoxButton]::YesNo,
        [System.Windows.MessageBoxImage]::Question
    )
    if ($answer -ne [System.Windows.MessageBoxResult]::Yes) {
        $controls.StatusText.Text = 'Assessment cancelled: required Microsoft Graph modules are missing.'
        Write-SemrActivity -Message "Operator declined installation of missing Microsoft Graph modules: $(@($moduleState.MissingModules) -join ', ')." -Level WARN
        return $false
    }

    try {
        Write-SemrActivity -Message "Installing missing Microsoft Graph modules for CurrentUser: $(@($moduleState.MissingModules) -join ', ')."
        Set-SemrProcessingStatus -Message 'Installing Microsoft Graph modules for the current user; please wait...'
        $installProgress = {
            Set-SemrProcessingStatus -Message 'Installing Microsoft Graph modules for the current user; please wait...' -SkipActivity
        }
        $installedState = Install-SemrMicrosoftGraphModule -ModuleNames @($moduleState.MissingModules) -ProgressCallback $installProgress
        if (-not $installedState.Available) {
            throw "Modules are still unavailable after installation: $(@($installedState.MissingModules) -join ', ')."
        }
        Write-SemrActivity -Message 'Required Microsoft Graph modules are available in PowerShell 7.' -Level SUCCESS
        return $true
    }
    catch {
        Show-SemrError -Title 'Microsoft Graph module installation failed' -ErrorRecord $_
        return $false
    }
}

function Set-SemrProcessingStatus {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS')][string]$Level = 'INFO',
        [switch]$SkipActivity
    )

    $controls.StatusText.Text = $Message
    if (-not $SkipActivity) { Write-SemrActivity -Message $Message -Level $Level }
    Invoke-SemrDoEvent
}

function Show-SemrError {
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)]$ErrorRecord
    )
    $message = if ($ErrorRecord -is [System.Management.Automation.ErrorRecord]) {
        $ErrorRecord.Exception.Message
    }
    elseif ($ErrorRecord -is [Exception]) {
        $ErrorRecord.Message
    }
    else {
        [string]$ErrorRecord
    }
    Write-SemrActivity -Message "$Title - $message" -Level ERROR
    [System.Windows.MessageBox]::Show($window, $message, $Title, 'OK', 'Error') | Out-Null
}

function Invoke-SemrDoEvent {
    $frame = [System.Windows.Threading.DispatcherFrame]::new()
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [System.Windows.Threading.DispatcherOperationCallback]{
            param($frameObject)
            $frameObject.Continue = $false
            return $null
        },
        $frame
    ) | Out-Null
    [System.Windows.Threading.Dispatcher]::PushFrame($frame)
}

function Get-SemrSelectedMode {
    $selected = $controls.ModeCombo.SelectedItem
    if ($selected -and $selected.Content) { return [string]$selected.Content }
    return 'Live'
}

function Get-SemrSelectedAuthenticationMode {
    $selected = $controls.AuthenticationModeCombo.SelectedItem
    if ($selected -and $selected.Content) { return [string]$selected.Content }
    return 'Interactive'
}

function Show-SemrMigrationEndpointSelection {
    param(
        [Parameter(Mandatory)][object[]]$Endpoints,
        [string]$PreferredName = ''
    )

    $dialog = [System.Windows.Window]::new()
    $dialog.Title = 'Select migration endpoint'
    $dialog.Owner = $window
    $dialog.Width = 540
    $dialog.SizeToContent = [System.Windows.SizeToContent]::Height
    $dialog.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
    $dialog.ResizeMode = [System.Windows.ResizeMode]::NoResize
    $dialog.ShowInTaskbar = $false
    $dialog.Background = [System.Windows.Media.Brushes]::White

    $panel = [System.Windows.Controls.StackPanel]::new()
    $panel.Margin = [System.Windows.Thickness]::new(20)
    $title = [System.Windows.Controls.TextBlock]::new()
    $title.Text = 'Several ExchangeRemoteMove endpoints are available.'
    $title.FontSize = 16
    $title.FontWeight = [System.Windows.FontWeights]::SemiBold
    $title.Margin = [System.Windows.Thickness]::new(0,0,0,8)
    [void]$panel.Children.Add($title)
    $description = [System.Windows.Controls.TextBlock]::new()
    $description.Text = 'Select the endpoint to validate for this assessment. The choice applies only to the current application session and does not modify the JSON file.'
    $description.TextWrapping = [System.Windows.TextWrapping]::Wrap
    $description.Margin = [System.Windows.Thickness]::new(0,0,0,14)
    [void]$panel.Children.Add($description)

    $combo = [System.Windows.Controls.ComboBox]::new()
    $combo.Height = 36
    $combo.DisplayMemberPath = 'Name'
    $combo.ItemsSource = @($Endpoints)
    $preferred = @($Endpoints | Where-Object Name -EQ $PreferredName | Select-Object -First 1)
    if ($preferred.Count -eq 1) { $combo.SelectedItem = $preferred[0] }
    $combo.Margin = [System.Windows.Thickness]::new(0,0,0,18)
    [void]$panel.Children.Add($combo)

    $buttonPanel = [System.Windows.Controls.StackPanel]::new()
    $buttonPanel.Orientation = [System.Windows.Controls.Orientation]::Horizontal
    $buttonPanel.HorizontalAlignment = [System.Windows.HorizontalAlignment]::Right
    $continueButton = [System.Windows.Controls.Button]::new()
    $continueButton.Content = 'Continue'
    $continueButton.MinWidth = 100
    $continueButton.Height = 34
    $continueButton.Margin = [System.Windows.Thickness]::new(0,0,8,0)
    $continueButton.IsEnabled = $null -ne $combo.SelectedItem
    $cancelButton = [System.Windows.Controls.Button]::new()
    $cancelButton.Content = 'Cancel'
    $cancelButton.MinWidth = 90
    $cancelButton.Height = 34
    [void]$buttonPanel.Children.Add($continueButton)
    [void]$buttonPanel.Children.Add($cancelButton)
    [void]$panel.Children.Add($buttonPanel)

    $combo.Add_SelectionChanged({ $continueButton.IsEnabled = $null -ne $combo.SelectedItem })
    $continueButton.Add_Click({
        $dialog.Tag = $combo.SelectedItem
        $dialog.DialogResult = $true
        $dialog.Close()
    })
    $cancelButton.Add_Click({
        $dialog.DialogResult = $false
        $dialog.Close()
    })
    $dialog.Content = $panel
    $accepted = $dialog.ShowDialog()
    if ($accepted -and $dialog.Tag) { return $dialog.Tag }
    return $null
}

function Initialize-SemrMigrationEndpointSelection {
    Set-SemrProcessingStatus -Message 'Loading ExchangeRemoteMove migration endpoints; please wait...'
    $endpointOptions = @(Get-SemrMigrationEndpointOption)
    $currentSelection = $controls.MigrationEndpointCombo.SelectedItem
    $currentName = if ($currentSelection -and $currentSelection.PSObject.Properties['Name']) { [string]$currentSelection.Name } else { '' }
    $configuredName = [string]$script:Config.Hybrid.MigrationEndpointName
    $preferredName = if ($currentName) { $currentName } else { $configuredName }
    $preferred = @($endpointOptions | Where-Object Name -EQ $preferredName | Select-Object -First 1)

    $script:UpdatingMigrationEndpointCombo = $true
    try {
        $controls.MigrationEndpointCombo.ItemsSource = @($endpointOptions)
        $controls.MigrationEndpointCombo.DisplayMemberPath = 'Name'
        $controls.MigrationEndpointCombo.SelectedIndex = -1
        if ($preferred.Count -eq 1) { $controls.MigrationEndpointCombo.SelectedItem = $preferred[0] }
    }
    finally {
        $script:UpdatingMigrationEndpointCombo = $false
    }
    $script:MigrationEndpointsLoaded = $true

    if ($endpointOptions.Count -eq 0) {
        $script:Config.Hybrid['MigrationEndpointName'] = ''
        $controls.MigrationEndpointCombo.Text = 'No ExchangeRemoteMove endpoint found'
        $script:MigrationEndpointSelectionConfirmed = $false
        Write-SemrActivity -Message 'No ExchangeRemoteMove migration endpoint was returned by Exchange Online.' -Level WARN
        return $true
    }
    if ($endpointOptions.Count -eq 1) {
        $controls.MigrationEndpointCombo.SelectedItem = $endpointOptions[0]
        $script:Config.Hybrid['MigrationEndpointName'] = [string]$endpointOptions[0].Name
        $script:MigrationEndpointSelectionConfirmed = $true
        Write-SemrActivity -Message "Migration endpoint selected automatically: $($endpointOptions[0].Name)." -Level SUCCESS
        return $true
    }

    if ($script:MigrationEndpointSelectionConfirmed -and $preferred.Count -eq 1) {
        $script:Config.Hybrid['MigrationEndpointName'] = [string]$preferred[0].Name
        Write-SemrActivity -Message "Migration endpoint retained for this session: $($preferred[0].Name)."
        return $true
    }

    $selectedEndpoint = Show-SemrMigrationEndpointSelection -Endpoints $endpointOptions -PreferredName $preferredName
    if (-not $selectedEndpoint) {
        $controls.StatusText.Text = 'Assessment cancelled: no migration endpoint was selected.'
        Write-SemrActivity -Message 'Assessment cancelled because no migration endpoint was selected.' -Level WARN
        return $false
    }
    $script:UpdatingMigrationEndpointCombo = $true
    try { $controls.MigrationEndpointCombo.SelectedItem = $selectedEndpoint }
    finally { $script:UpdatingMigrationEndpointCombo = $false }
    $script:Config.Hybrid['MigrationEndpointName'] = [string]$selectedEndpoint.Name
    $script:MigrationEndpointSelectionConfirmed = $true
    Write-SemrActivity -Message "Migration endpoint selected for this session: $($selectedEndpoint.Name)." -Level SUCCESS
    return $true
}

function Sync-SemrConnectionDisplay {
    $mode = Get-SemrSelectedMode
    $authenticationMode = Get-SemrSelectedAuthenticationMode
    $script:Config['Mode'] = $mode
    $script:Config.Authentication['Mode'] = $authenticationMode
    $state = Get-SemrConnectionState
    $cachePath = [string]$script:Config._CacheRootPath

    if ($mode -eq 'CacheOnly') {
        $controls.AdStateText.Text = 'CacheOnly (CSV inventory)'
        $controls.OnPremStateText.Text = 'CacheOnly (CSV inventory)'
        $controls.ExoStateText.Text = 'CacheOnly (CSV inventories)'
        $controls.GraphStateText.Text = 'CacheOnly (CSV inventories)'
        foreach ($control in @($controls.AdStateText, $controls.OnPremStateText, $controls.ExoStateText, $controls.GraphStateText)) {
            $control.Foreground = '#146C43'
        }
        if (-not $script:Assessment) {
            $controls.HybridStateText.Text = 'CacheOnly: migration endpoint test skipped; Entra Connect health will be read from CSV.'
            $controls.HybridStateText.Foreground = '#5F6B7A'
        }
    }
    else {
        $controls.AdStateText.Text = if ($state.ActiveDirectory) { 'Connected (Live)' } else { 'Automatic live check; CSV fallback if unavailable' }
        $controls.OnPremStateText.Text = if ($state.OnPremisesExchange) { 'Connected (Live)' } else { 'Automatic live check; CSV fallback if unavailable' }
        $controls.ExoStateText.Text = if ($state.ExchangeOnline) { "Connected (Live, $authenticationMode)" } elseif ($authenticationMode -eq 'Application') { 'Application certificate connection starts automatically with Run assessment' } else { 'Interactive connection starts automatically with Run assessment' }
        $controls.GraphStateText.Text = if ($state.MicrosoftGraph) { "Live evidence loaded (isolated process, $authenticationMode)" } elseif ($authenticationMode -eq 'Application') { 'Application certificate collection starts automatically after Exchange Online' } else { 'Interactive collection starts automatically after Exchange Online' }
        $controls.AdStateText.Foreground = if ($state.ActiveDirectory) { '#146C43' } else { '#8A5A00' }
        $controls.OnPremStateText.Foreground = if ($state.OnPremisesExchange) { '#146C43' } else { '#8A5A00' }
        $controls.ExoStateText.Foreground = if ($state.ExchangeOnline) { '#146C43' } else { '#5F6B7A' }
        $controls.GraphStateText.Foreground = if ($state.MicrosoftGraph) { '#146C43' } else { '#5F6B7A' }
        if (-not $script:Assessment) {
            $controls.HybridStateText.Text = 'Automatic: Exchange Online migration endpoint and Entra Connect synchronization health.'
            $controls.HybridStateText.Foreground = '#5F6B7A'
        }
    }

    $controls.ModeCombo.IsEnabled = -not $script:IsBusy
    $controls.AuthenticationModeCombo.IsEnabled = (-not $script:IsBusy -and $mode -eq 'Live')
    $controls.MigrationEndpointCombo.IsEnabled = (-not $script:IsBusy -and $mode -eq 'Live' -and $script:MigrationEndpointsLoaded)
    $controls.RunButton.IsEnabled = (-not $script:IsBusy -and $null -ne $script:Batch)
    $alternativeCachePath = [string]$script:Config._AlternativeCacheRootPath
    $controls.FooterText.Text = "Read-only | Mode: $mode | Auth: $authenticationMode | Tenant: $($script:Config._TenantProfileKey) | Cache: $cachePath | Alternative: $alternativeCachePath"
}
function Switch-SemrBusyState {
    param(
        [bool]$Busy,
        [string]$Message = ''
    )
    $script:IsBusy = $Busy
    foreach ($button in @($controls.BrowseCsvButton, $controls.RunButton)) {
        $button.IsEnabled = -not $Busy
    }
    $controls.CsvPathBox.IsEnabled = -not $Busy
    $controls.CancelButton.IsEnabled = $Busy
    $window.Cursor = if ($Busy) { [System.Windows.Input.Cursors]::Wait } else { [System.Windows.Input.Cursors]::Arrow }
    Sync-SemrConnectionDisplay
    if ($Message) { $controls.StatusText.Text = $Message }
    Invoke-SemrDoEvent
}
function Resolve-SemrOutputRoot {
    $configured = [string]$script:Config.OutputRoot
    if ([string]::IsNullOrWhiteSpace($configured)) { $configured = 'Output' }
    if ([System.IO.Path]::IsPathRooted($configured)) { return $configured }
    return Join-Path $PSScriptRoot $configured
}

function Initialize-SemrSessionLog {
    try {
        $logFolder = Join-Path (Resolve-SemrOutputRoot) 'Logs'
        New-Item -Path $logFolder -ItemType Directory -Force | Out-Null
        $script:SessionLogPath = Join-Path $logFolder ("SmartM365-ExchangeMigrationReadiness-{0}.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
        $line = "{0} [INFO] Smart Exchange Migration Readiness v{1} session log created." -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $script:AppVersion
        [IO.File]::WriteAllText($script:SessionLogPath, "$line`r`n", [Text.UTF8Encoding]::new($false))
    }
    catch {
        $script:SessionLogPath = ''
        $script:SessionLogError = $_.Exception.Message
    }
}

Initialize-SemrSessionLog

function Import-SemrCsvToGui {
    try {
        $path = $controls.CsvPathBox.Text.Trim()
        $script:Batch = Import-SemrBatchCsv -Path $path
        $controls.BatchGrid.ItemsSource = @($script:Batch.Rows)
        $controls.CsvMetadataText.Text = "{0} row(s) | Encoding: {1} | Delimiter: {2} | Identity: {3}" -f @(
            $script:Batch.Rows.Count, $script:Batch.Encoding, $script:Batch.Delimiter, $script:Batch.IdentityColumn
        )
        $controls.StatusText.Text = "Batch loaded: $($script:Batch.Rows.Count) mailbox row(s)."
        $controls.RunButton.IsEnabled = $true
        Write-SemrActivity -Message "Loaded CSV '$($script:Batch.Path)' with $($script:Batch.Rows.Count) row(s)." -Level SUCCESS
        Set-SemrProcessingStatus -Message 'Checking primary and alternative CSV cache paths; please wait...' -SkipActivity
        Update-SemrCsvSourcesDisplay
        $controls.StatusText.Text = "Batch loaded: $($script:Batch.Rows.Count) mailbox row(s)."
    }
    catch {
        $script:Batch = $null
        $controls.RunButton.IsEnabled = $false
        Show-SemrError -Title 'CSV validation failed' -ErrorRecord $_
    }
}

function Sync-SemrResultCount {
    $summary = @($script:Assessment.Summary)
    $controls.GoCountText.Text = @($summary | Where-Object Decision -EQ 'GO').Count
    $controls.WarningCountText.Text = @($summary | Where-Object Decision -EQ 'GO-WARNING').Count
    $controls.NoGoCountText.Text = @($summary | Where-Object Decision -EQ 'NO-GO').Count
    $controls.UnknownCountText.Text = @($summary | Where-Object Decision -EQ 'UNKNOWN').Count
}

$controls.HeaderLogoLink.Add_MouseLeftButtonUp({
    try {
        Start-Process 'https://workplacecloudhub.com'
    }
    catch {
        Show-SemrError -Title 'Unable to open WorkplaceCloudHub' -ErrorRecord $_
    }
})

$controls.BrowseCsvButton.Add_Click({
    $dialog = [Microsoft.Win32.OpenFileDialog]::new()
    $dialog.Title = 'Select Exchange migration batch CSV'
    $dialog.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    if ($dialog.ShowDialog($window)) {
        $controls.CsvPathBox.Text = $dialog.FileName
        Import-SemrCsvToGui
    }
})

$controls.ModeCombo.Add_SelectionChanged({
    if ($script:IsBusy) { return }
    $mode = Get-SemrSelectedMode
    $script:Config['Mode'] = $mode
    $authenticationMode = Get-SemrSelectedAuthenticationMode
    $controls.HybridStateText.Text = if ($mode -eq 'CacheOnly') { 'CacheOnly: live endpoint test skipped; Entra sync health is read from CSV cache.' } else { "Live: EXO/Graph authentication mode is $authenticationMode; AD, Exchange on-premises and Entra Connect use live data when available, otherwise CSV fallback." }
    $controls.HybridStateText.Foreground = '#5F6B7A'
    $controls.StatusText.Text = "Mode selected: $mode"
    Sync-SemrConnectionDisplay
    Set-SemrProcessingStatus -Message 'Checking primary and alternative CSV cache paths; please wait...' -SkipActivity
    Update-SemrCsvSourcesDisplay
    $controls.StatusText.Text = "Mode selected: $mode"
    Write-SemrActivity -Message "Execution mode changed to $mode. Cache root: $($script:Config._CacheRootPath)"
})

$controls.AuthenticationModeCombo.Add_SelectionChanged({
    if ($script:IsBusy) { return }
    $authenticationMode = Get-SemrSelectedAuthenticationMode
    $previousMode = [string]$script:Config.Authentication.Mode
    $script:Config.Authentication['Mode'] = $authenticationMode
    if ($previousMode -ne $authenticationMode) {
        $state = Get-SemrConnectionState
        if ($state.ExchangeOnline -or $state.MicrosoftGraph) {
            Disconnect-SemrSession
            $script:MigrationEndpointsLoaded = $false
            $script:MigrationEndpointSelectionConfirmed = $false
            $script:UpdatingMigrationEndpointCombo = $true
            try {
                $controls.MigrationEndpointCombo.ItemsSource = $null
                $controls.MigrationEndpointCombo.SelectedIndex = -1
                $configuredEndpointName = [string]$script:Config.Hybrid.MigrationEndpointName
                $controls.MigrationEndpointCombo.Text = if ($configuredEndpointName) { $configuredEndpointName } else { 'Loaded after EXO connection' }
            }
            finally {
                $script:UpdatingMigrationEndpointCombo = $false
            }
            Write-SemrActivity -Message 'Existing live connections were closed because the authentication mode changed.' -Level INFO
        }
        Write-SemrActivity -Message "Cloud authentication mode changed to $authenticationMode."
    }
    $controls.StatusText.Text = "Authentication selected: $authenticationMode"
    Sync-SemrConnectionDisplay
})

$controls.MigrationEndpointCombo.Add_SelectionChanged({
    if ($script:UpdatingMigrationEndpointCombo -or -not $script:MigrationEndpointsLoaded) { return }
    $selectedEndpoint = $controls.MigrationEndpointCombo.SelectedItem
    if (-not $selectedEndpoint -or -not $selectedEndpoint.PSObject.Properties['Name']) { return }
    $script:Config.Hybrid['MigrationEndpointName'] = [string]$selectedEndpoint.Name
    $script:MigrationEndpointSelectionConfirmed = $true
    $controls.HybridStateText.Text = "Migration endpoint selected for the next assessment: $($selectedEndpoint.Name)."
    $controls.HybridStateText.Foreground = '#5F6B7A'
    Write-SemrActivity -Message "Migration endpoint changed for this application session: $($selectedEndpoint.Name)."
})

$controls.OptionsItemsControl.AddHandler(
    [System.Windows.Controls.Primitives.ToggleButton]::CheckedEvent,
    [System.Windows.RoutedEventHandler]{ param($sender,$eventArgs) Update-SemrOptionsSummary }
)
$controls.OptionsItemsControl.AddHandler(
    [System.Windows.Controls.Primitives.ToggleButton]::UncheckedEvent,
    [System.Windows.RoutedEventHandler]{ param($sender,$eventArgs) Update-SemrOptionsSummary }
)
$controls.EnableAllOptionsButton.Add_Click({
    foreach ($option in $script:CheckOptions) { $option.Enabled = $true }
    Refresh-SemrOptionsDisplay
})
$controls.ResetOptionsButton.Add_Click({
    foreach ($option in $script:CheckOptions) { $option.Enabled = [bool]($option.Mandatory -or $script:ConfiguredDisabledChecks -notcontains $option.CheckId) }
    Refresh-SemrOptionsDisplay
})

$controls.RunButton.Add_Click({
    if (-not $script:Batch) { return }
    try {
        $script:CancelRequested = $false
        Sync-SemrCheckOptionsToConfig
        Write-SemrActivity -Message "$(@($script:CheckOptions | Where-Object Enabled).Count) checks enabled; $(@($script:Config.DisabledChecks).Count) optional checks disabled for this run."
        Set-SemrProcessingStatus -Message 'Checking CSV cache freshness; please wait...'
        if (-not (Confirm-SemrStaleCsvUsage)) { return }
        Switch-SemrBusyState -Busy $true -Message 'Assessment running - please wait...'
        $controls.RunProgress.Value = 0
        Write-SemrActivity -Message "Starting read-only assessment in $($script:Config.Mode) mode for $($script:Batch.Rows.Count) mailbox row(s)."
        if ((Get-SemrSelectedMode) -eq 'Live') {
            $authenticationMode = Get-SemrSelectedAuthenticationMode
            $script:Config.Authentication['Mode'] = $authenticationMode
            if (-not (Confirm-SemrMicrosoftGraphModule)) { return }
            $connectionState = Get-SemrConnectionState
            if (-not $connectionState.ExchangeOnline) {
                $exoStatus = if ($authenticationMode -eq 'Application') { 'Connecting to Exchange Online with the configured application certificate; please wait...' } else { 'Connecting interactively to Exchange Online. Complete authentication, then please wait...' }
                Set-SemrProcessingStatus -Message $exoStatus
                $exoConfig = $script:Config.ExchangeOnline
                Connect-SemrExchangeOnline -UserPrincipalName ([string]$exoConfig.UserPrincipalName) -DisableWam ([bool]$exoConfig.DisableWam) -TenantId ([string]$script:Config._TenantId) -AuthenticationMode $authenticationMode -AppId ([string]$script:Config._AppId) -Organization ([string]$script:Config._OrgDomain) -CertificateThumbprint ([string]$script:Config._CertificateThumbprint) | Out-Null
                Write-SemrActivity -Message "Exchange Online connected using $authenticationMode authentication." -Level SUCCESS
                Sync-SemrConnectionDisplay
            }

            $endpointCheck = @($script:CheckOptions | Where-Object CheckId -EQ 'HYBRID-ENDPOINT' | Select-Object -First 1)
            if ($endpointCheck.Count -eq 1 -and $endpointCheck[0].Enabled) {
                if (-not (Initialize-SemrMigrationEndpointSelection)) { return }
            }

            $connectionState = Get-SemrConnectionState
            if (-not $connectionState.MicrosoftGraph) {
                $graphStatus = if ($authenticationMode -eq 'Application') { 'Connecting to Microsoft Graph with the configured application certificate; please wait...' } else { 'Connecting interactively to Microsoft Graph. Complete authentication, then please wait...' }
                Set-SemrProcessingStatus -Message $graphStatus
                $graphConfig = $script:Config.MicrosoftGraph
                $graphProgressAction = {
                    param($Message)
                    Set-SemrProcessingStatus -Message "$Message Please wait..." -SkipActivity
                }
                $emailAddresses = @($script:Batch.Rows | ForEach-Object { [string]$_.EmailAddress })
                Connect-SemrMicrosoftGraph -Scopes @($graphConfig.Scopes) -TenantId ([string]$script:Config._TenantId) -AuthenticationMode $authenticationMode -AppId ([string]$script:Config._AppId) -CertificateThumbprint ([string]$script:Config._CertificateThumbprint) -EmailAddresses $emailAddresses -ProgressCallback $graphProgressAction | Out-Null
                Write-SemrActivity -Message "Microsoft Graph evidence collected in an isolated PowerShell process using $authenticationMode authentication." -Level SUCCESS
                Sync-SemrConnectionDisplay
            }
        }
        $progressAction = {
            param($Current, $Total, $Mailbox, $Phase)
            $percent = if ($Total -gt 0) { [math]::Round(($Current / $Total) * 100, 0) } else { 0 }
            $controls.RunProgress.Value = $percent
            $mailboxText = if ([string]::IsNullOrWhiteSpace([string]$Mailbox)) { '' } else { " - $Mailbox" }
            $statusLine = "[$Current/$Total] $Phase$mailboxText. Please wait..."
            $controls.StatusText.Text = $statusLine
            Write-SemrActivity -Message $statusLine
            Invoke-SemrDoEvent
        }
        $cancellationCheck = {
            Invoke-SemrDoEvent
            return $script:CancelRequested
        }
        $script:Assessment = Invoke-SemrAssessment -Batch $script:Batch -Config $script:Config -ProgressCallback $progressAction -CancellationCheck $cancellationCheck
        if ($script:Assessment.SourceInitialization) {
            $sourceState = $script:Assessment.SourceInitialization
            Write-SemrActivity -Message ([string]$sourceState.ActiveDirectoryMessage) -Level $(if ($sourceState.ActiveDirectoryLive -or $script:Config.Mode -eq 'CacheOnly') { 'INFO' } else { 'WARN' })
            Write-SemrActivity -Message ([string]$sourceState.ExchangeOnPremisesMessage) -Level $(if ($sourceState.ExchangeOnPremisesLive -or $script:Config.Mode -eq 'CacheOnly') { 'INFO' } else { 'WARN' })
        }
        Write-SemrActivity -Message "Entra Connect evidence source: $($script:Assessment.EntraConnect.Source). $($script:Assessment.EntraConnect.Message)" -Level $(if ($script:Assessment.EntraConnect.Available) { 'INFO' } else { 'WARN' })
        $endpointName = if ([string]::IsNullOrWhiteSpace([string]$script:Assessment.Hybrid.EndpointName)) { 'not available' } else { [string]$script:Assessment.Hybrid.EndpointName }
        $controls.HybridStateText.Text = "Migration endpoint: $endpointName | $($script:Assessment.Hybrid.Message)`nEntra Connect: $($script:Assessment.EntraConnect.Message)"
        $controls.HybridStateText.Foreground = if ($script:Assessment.EntraConnect.Available -and ($script:Assessment.Hybrid.ConnectivitySuccess -or (Get-SemrSelectedMode) -eq 'CacheOnly')) { '#146C43' } else { '#8A5A00' }
        Set-SemrProcessingStatus -Message 'Generating CSV reports and finalizing the assessment; please wait...'
        $script:Export = Export-SemrAssessment -Assessment $script:Assessment -OutputRoot (Resolve-SemrOutputRoot)
        $controls.SummaryGrid.ItemsSource = @($script:Assessment.Summary)
        $controls.FindingsGrid.ItemsSource = @($script:Assessment.Findings)
        $controls.PermissionsGrid.ItemsSource = @($script:Assessment.PermissionsBaseline)
        $controls.CsvSourcesGrid.ItemsSource = @($script:Assessment.CsvSources)
        $controls.OpenOutputButton.IsEnabled = $true
        $controls.ComparePermissionsButton.IsEnabled = (Test-Path -LiteralPath $script:Export.PermissionsPath)
        Sync-SemrResultCount
        $controls.RunProgress.Value = 100
        $controls.StatusText.Text = if ($script:Assessment.Cancelled) { 'Assessment cancelled; partial results exported.' } else { "Assessment complete: $($script:Assessment.RunId)" }
        Write-SemrActivity -Message "Assessment exported to '$($script:Export.RunFolder)'." -Level SUCCESS
        $controls.MainTabs.SelectedItem = $controls.ResultsTab
    }
    catch {
        Update-SemrCsvSourcesDisplay -AssessmentCompleted
        Show-SemrError -Title 'Assessment failed' -ErrorRecord $_
        $controls.StatusText.Text = 'Assessment failed. Review Activity.'
    }
    finally {
        Switch-SemrBusyState -Busy $false
    }
})

$controls.CancelButton.Add_Click({
    $script:CancelRequested = $true
    $controls.StatusText.Text = 'Cancellation requested...'
    Write-SemrActivity -Message 'Cancellation requested by the operator.' -Level WARN
})

$controls.OpenOutputButton.Add_Click({
    if ($script:Export -and (Test-Path -LiteralPath $script:Export.RunFolder)) {
        Start-Process explorer.exe -ArgumentList @($script:Export.RunFolder)
    }
})

$controls.MailboxFilterBox.Add_TextChanged({
    if (-not $script:Assessment) { return }
    $filter = $controls.MailboxFilterBox.Text.Trim()
    $controls.FindingsGrid.ItemsSource = if ($filter) {
        @($script:Assessment.Findings | Where-Object { $_.EmailAddress -like "*$filter*" -or $_.CheckId -like "*$filter*" -or $_.Message -like "*$filter*" })
    }
    else {
        @($script:Assessment.Findings)
    }
})
$controls.ClearFilterButton.Add_Click({ $controls.MailboxFilterBox.Text = '' })

$controls.ComparePermissionsButton.Add_Click({
    try {
        if (-not $script:Batch -or -not $script:Export) { return }
        Switch-SemrBusyState -Busy $true -Message 'Collecting current Exchange Online permissions...'
        $progressAction = {
            param($Current, $Total, $Mailbox, $Phase)
            $controls.RunProgress.Value = if ($Total -gt 0) { [math]::Round(($Current / $Total) * 100, 0) } else { 0 }
            $controls.StatusText.Text = "[$Current/$Total] $Phase - $Mailbox"
            Invoke-SemrDoEvent
        }
        $current = @(Get-SemrPostMigrationPermission -Batch $script:Batch -ProgressCallback $progressAction)
        $currentPath = Join-Path $script:Export.RunFolder 'Permissions-Current-EXO.csv'
        if ($current.Count -gt 0) {
            $current | Export-Csv -LiteralPath $currentPath -NoTypeInformation -Encoding utf8
        }
        else {
            Set-Content -LiteralPath $currentPath -Value '"EmailAddress","PermissionType","Delegate","IsInherited","Source","CapturedAt"' -Encoding utf8
        }
        $comparison = @(Compare-SemrPermissionsBaseline -BaselinePath $script:Export.PermissionsPath -CurrentPermissions $current)
        $controls.PermissionsGrid.ItemsSource = $comparison
        $comparisonPath = Join-Path $script:Export.RunFolder 'Permissions-Comparison.csv'
        $comparison | Export-Csv -LiteralPath $comparisonPath -NoTypeInformation -Encoding utf8
        $controls.StatusText.Text = 'Post-migration permission comparison completed.'
        Write-SemrActivity -Message "Current EXO permissions exported to '$currentPath'." -Level SUCCESS
        Write-SemrActivity -Message "Permissions comparison exported to '$comparisonPath'." -Level SUCCESS
    }
    catch { Show-SemrError -Title 'Permissions comparison failed' -ErrorRecord $_ }
    finally { Switch-SemrBusyState -Busy $false }
})

$window.Add_ContentRendered({
    if ($script:GuiSplash) {
        Hide-SmartM365GuiSplash -Splash $script:GuiSplash
    }
    Sync-SemrConnectionDisplay
    Update-SemrOptionsSummary
    Set-SemrProcessingStatus -Message 'Checking primary and alternative CSV cache paths; please wait...' -SkipActivity
    Update-SemrCsvSourcesDisplay
    Write-SemrActivity -Message "Smart Exchange Migration Readiness v$script:AppVersion started in $($script:Config.Mode) mode with $($script:Config._AuthenticationMode) cloud authentication. Tenant profile: $($script:Config._TenantProfileKey). Cache root: $($script:Config._CacheRootPath). Alternative cache: $($script:Config._AlternativeCacheRootPath)"
    Write-SemrActivity -Message 'Some live operations can take several minutes and may temporarily delay UI response. Follow the current-operation status and please wait before closing the application.'
    if ($script:SessionLogPath) {
        Write-SemrActivity -Message "Session log: $script:SessionLogPath"
    }
    elseif ($script:SessionLogError) {
        Write-SemrActivity -Message "Session log could not be created: $script:SessionLogError" -Level WARN
    }
    $controls.StatusText.Text = if ($script:Batch) { "Batch loaded: $($script:Batch.Rows.Count) mailbox row(s)." } else { 'Ready. Select a migration batch CSV.' }
    foreach ($key in @($script:Config._AddedKeys)) {
        Write-SemrActivity -Message "Configuration key added from template: $key" -Level INFO
    }
    $updateHelper = Join-Path $PSScriptRoot 'SmartM365.GuiUpdateCheck.ps1'
    $updateManifest = Join-Path $PSScriptRoot 'SmartM365.GuiUpdateCheck.psd1'
    if ((Test-Path -LiteralPath $updateHelper) -and (Test-Path -LiteralPath $updateManifest)) {
        try {
            . $updateHelper
            Start-SmartM365GuiUpdateCheck -Owner $window -ManifestPath $updateManifest -AppRoot $PSScriptRoot
        }
        catch { $null = $_ }
    }
    if ($controls.CsvPathBox.Text.Trim()) {
        Import-SemrCsvToGui
    }
})

$window.Add_Closing({
    if ($script:IsBusy) {
        $script:CancelRequested = $true
    }
    Disconnect-SemrSession
})

try {
    [void]$window.ShowDialog()
}
finally {
    if ($script:SessionLogPath) {
        try { [IO.File]::AppendAllText($script:SessionLogPath, "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [INFO] Application closed.`r`n", [Text.UTF8Encoding]::new($false)) } catch {}
    }
    if ($script:GuiSplash) {
        Close-SmartM365GuiSplash -Splash $script:GuiSplash
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDhyMTlQKn5FqD9
# S6UTHV6F6F1RMXOqY49ap0D9V1Wp56CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPnvED6yGg7K0Y2QE2z29jUrji0asvYvSCb5PQnQ+lkzMA0GCSqG
# SIb3DQEBAQUABIIBgA8Ox3B+6UXtXTA8ZdTy9jsDfU1EQS+5/iPrlBJovIOVUW2L
# wiQH+SU2sAN/ENiuoeNG9asAhcJX2vrLSoN31Pp9IsJZHw2B5mCaBQDows9rsELF
# 31JwZi8bshGp+uwYd6tWm5JQFdA6eT9zJbVakc167AQXwjMW4Z86P3y3c0fS4pXZ
# sFlpQO+n0AIp67tyzg2ilZBmRecqLoJ+dWMnY+qA2oaIPYIUgA4xLiXAPyWLyw75
# SCSNmgUw/gnkpy/jNaAe46LcpEM0NJ5IQAubKdJWA5ysukkZm4VBxY5HyCGzh95e
# BKkk677poAUeXE4pX6gGB/tleH2h7SNyI9+P9vj2OC3rsvwgFCS8gSyc6B7eF8nI
# w3x6xOOxFaN3GNSpbBnqWH9/oX4H+KW9rN2AVVnJcqYB1uo7uGOe0tbi2SL+TLj3
# VwgT8u5zHbNqamHBWPeXVG0lNE4+Yy58XGJ7uoB1B7oA3Eyg3HLSY1yyhA5NKnvU
# zoywjp18MIFqDavNEqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTgxMTQ3
# NTBaMC8GCSqGSIb3DQEJBDEiBCD5WHlbjXoTMqV8mPUqba+q5Fw5S7EVz3nqg2DP
# 3tuJrTANBgkqhkiG9w0BAQEFAASCAgB/g5RJfPL0UvMkRdSbZHAr8i2MKpwmLzZn
# f3d4m4C5Pm7YS1gDgyNr55dfWx/5wL1KHVHltroQhWjJgaLrWif/fWw41Orwd7sG
# xfHPSM9hW2MiuUZAxGhR3K167fjo55pca086SitvE+orVojV3N1Rk7Y+wOczL4n7
# wSvzbJw7gXSK8ckUJTuypylKX2p8w0/RE4IPz5zg5KLF4Y1en7Ahrxg6W+ujtmwo
# eHzlj/YlobbfFQaO4C1YwBmHQXaReOohHKKvFhR93e7irfBcv6KwL3tfoYFOvpLf
# 7qFJTFJDNfLfeiCu5X16Cc8MlONc0gVsR+yPqR8j0KCc5DYe7vo5srCN4hUAa5YC
# jHIvISf9ZWBF1i80IlmP9W8Fy1HNprXSCHA4OaL0apmrbZ3ldN1aIqWNd3UJKZoq
# TcnBpDO72m3tlt9Ehn4EKe9f+O2hBy4TXONVCxOwAXpfk8Llg1Fk9ibINNpXSLny
# MYgCrfzzoJfbP8Hqv2oHvcpaJTrqGZkoaq3eiQ/5eBFTfwLenPabHi88XnpY2uuw
# pWwZsruKo4IjfaGqsVDCQNZNdXalTCc8YFaJLsKvCzXSJhLisMskRUoewyLbrO5P
# iKYlZ/UsXflB5ZIUYWXVWTPYkATHrpD8yvulUl98DVMcH7T+oBd+nIFjqXI+LS3H
# KO3RcHmfdg==
# SIG # End signature block
