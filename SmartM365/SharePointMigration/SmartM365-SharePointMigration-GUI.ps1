<#
.SYNOPSIS
    Smart SharePoint Migration dashboard GUI.

.DESCRIPTION
    WPF dashboard for SharePoint migration workflows. Discovers local migration
    folders, shows the last run status for each step, and launches inventory,
    comparison, and operation actions in a new PowerShell console window that stays open after completion.

.PARAMETER ValidateOnly
    Loads the GUI resources and exits without showing the window.

.VERSION
    1.0.11
#>

#Requires -Version 7.4

[CmdletBinding()]
param(
    [Alias('DryRun')]
    [switch]$ValidateOnly
)

$script:AppName    = 'Smart SharePoint Migration'
$script:AppVersion = '1.0.11'
$script:ScriptRoot = $PSScriptRoot
Microsoft.PowerShell.Utility\Write-Host ('{0} Script  : {1} v{2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $MyInvocation.MyCommand.Name, $script:AppVersion) -ForegroundColor Cyan

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

. (Join-Path $script:ScriptRoot 'SmartM365.GuiSplash.ps1')
. (Join-Path $script:ScriptRoot 'SmartM365-SharePointMigration-NewWizard.ps1')
. (Join-Path $script:ScriptRoot 'Scripts\Launchers\Generic\SmartM365-SharePointMigration-GuiActivity.ps1')

$updateCheckModulePath = Join-Path $script:ScriptRoot 'SmartM365.GuiUpdateCheck.ps1'
if (Test-Path -LiteralPath $updateCheckModulePath -PathType Leaf) {
    . $updateCheckModulePath
}

$script:Splash = $null
if (-not $ValidateOnly) {
    $script:Splash = Start-SmartM365GuiSplash `
        -ProductName    $script:AppName `
        -Subtitle       'SharePoint migration dashboard' `
        -MinimumDurationMs 4000
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Get-MigrationFolders {
    $root = Join-Path $script:ScriptRoot 'Migrations'
    if (-not (Test-Path -LiteralPath $root -PathType Container)) { return @() }
    $results = [System.Collections.Generic.List[object]]::new()
    foreach ($dir in (Get-ChildItem -LiteralPath $root -Directory)) {
        if ($dir.Name -like '.new-*') { continue }
        $cfgPath = Join-Path $dir.FullName 'migration.config.psd1'
        if (-not (Test-Path -LiteralPath $cfgPath -PathType Leaf)) { continue }
        try {
            $cfg = Import-PowerShellDataFile -LiteralPath $cfgPath
        } catch { continue }
        if ($cfg.Name -eq 'NewMigration') { continue }
        $results.Add([pscustomobject]@{
            Name       = $dir.Name
            Root       = $dir.FullName
            Config     = $cfg
            ConfigPath = $cfgPath
        })
    }
    return $results.ToArray()
}

function Get-LatestCsvFile {
    param([string]$Directory, [string]$Filter)
    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) { return $null }
    Get-ChildItem -LiteralPath $Directory -Filter $Filter -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '*-Errors.csv' -and $_.Name -notlike '*-Run.log' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}
function Get-CsvFileItems {
    param([string]$Directory, [string]$Filter)
    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) { return @() }
    @(Get-ChildItem -LiteralPath $Directory -Filter $Filter -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notlike '*-Errors.csv' -and $_.Name -notlike '*-Run.log' } |
        Sort-Object LastWriteTime -Descending |
        ForEach-Object {
            [pscustomobject]@{
                Display  = ('{0}  {1}' -f $_.LastWriteTime.ToString('yyyy-MM-dd HH:mm'), $_.Name)
                File     = $_
                FullName = $_.FullName
                Name     = $_.Name
            }
        })
}

function Get-LatestSubfolder {
    param([string]$Directory, [string]$Pattern)
    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) { return $null }
    Get-ChildItem -LiteralPath $Directory -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like $Pattern } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}
function Get-MigrationEndpointType {
    param($Config, [string]$Side)
    $section = if ($Side -eq 'Source') { $Config.Source } else { $Config.Target }
    $defaultType = if ($Side -eq 'Source') { 'SP2019' } else { 'SPO' }
    if ($section -and $section.ContainsKey('Type') -and -not [string]::IsNullOrWhiteSpace([string]$section.Type)) {
        $rawType = [string]$section.Type
    }
    else {
        $rawType = $defaultType
    }

    switch -Regex ($rawType.Trim().ToUpperInvariant()) {
        '^(SP2016|SHAREPOINT2016|2016)$' { return 'SP2016' }
        '^(SP2019|SHAREPOINT2019|2019)$' { return 'SP2019' }
        '^(SPO|SHAREPOINTONLINE|ONLINE)$' { return 'SPO' }
    }

    return $rawType
}

function Get-MigrationEndpointUrlText {
    param($Config, [string]$Side)
    $section = if ($Side -eq 'Source') { $Config.Source } else { $Config.Target }
    $endpointType = Get-MigrationEndpointType $Config $Side
    $keys = if ($endpointType -eq 'SPO') { @('SiteUrl', 'UrlsFile', 'WebApplicationUrl') } else { @('WebApplicationUrl', 'SiteUrl', 'UrlsFile') }
    foreach ($key in $keys) {
        if ($section -and $section.ContainsKey($key) -and -not [string]::IsNullOrWhiteSpace([string]$section[$key])) {
            return [string]$section[$key]
        }
    }

    if ($Config.ContainsKey('Comparison') -and $Config.Comparison.ContainsKey('PathMappingsFile') -and -not [string]::IsNullOrWhiteSpace([string]$Config.Comparison.PathMappingsFile)) {
        return [string]$Config.Comparison.PathMappingsFile
    }

    return ''
}

function Get-MigrationStatus {
    param($Migration)
    $cfg  = $Migration.Config
    $root = $Migration.Root
    $name = $cfg.Name

    $srcFileDir  = Join-Path $root $cfg.Output.SourceFileScans
    $tgtFileDir  = Join-Path $root $cfg.Output.TargetFileScans
    $fileCmpDir  = Join-Path $root $cfg.Output.FileComparisons
    $historyPath = if ($cfg.Output.ContainsKey('FileHistoryComparisons') -and -not [string]::IsNullOrWhiteSpace([string]$cfg.Output.FileHistoryComparisons)) {
        [string]$cfg.Output.FileHistoryComparisons
    } elseif ($cfg.Output.ContainsKey('SourceHistoryComparisons') -and -not [string]::IsNullOrWhiteSpace([string]$cfg.Output.SourceHistoryComparisons)) {
        [string]$cfg.Output.SourceHistoryComparisons
    } else {
        'comparisons\scan-history'
    }
    $histDir     = Join-Path $root $historyPath
    $srcPermDir  = Join-Path $root $cfg.Output.SourcePermissionScans
    $tgtPermDir  = Join-Path $root $cfg.Output.TargetPermissionScans
    $permCmpDir  = Join-Path $root $cfg.Output.PermissionComparisons

    [pscustomobject]@{
        SourceFileCsv         = Get-LatestCsvFile    $srcFileDir  ("{0}-FileInventory-$name-*.csv" -f (Get-MigrationEndpointType $cfg 'Source'))
        SourceFileCsvItems    = Get-CsvFileItems     $srcFileDir  ("{0}-FileInventory-$name-*.csv" -f (Get-MigrationEndpointType $cfg 'Source'))
        TargetFileCsv         = Get-LatestCsvFile    $tgtFileDir  ("{0}-FileInventory-$name-*.csv" -f (Get-MigrationEndpointType $cfg 'Target'))
        TargetFileCsvItems    = Get-CsvFileItems     $tgtFileDir  ("{0}-FileInventory-$name-*.csv" -f (Get-MigrationEndpointType $cfg 'Target'))
        FileComparisonFolder  = Get-LatestSubfolder  $fileCmpDir  "$name-*"
        HistoryFolder         = Get-LatestSubfolder  $histDir     '*-Changes-*'
        SourcePermCsv         = Get-LatestCsvFile    $srcPermDir  ("{0}-PermissionInventory-$name-*.csv" -f (Get-MigrationEndpointType $cfg 'Source'))
        TargetPermCsv         = Get-LatestCsvFile    $tgtPermDir  ("{0}-PermissionInventory-$name-*.csv" -f (Get-MigrationEndpointType $cfg 'Target'))
        PermComparisonFolder  = Get-LatestSubfolder  $permCmpDir  "$name-*"
    }
}

function Get-MigrationOperations {
    param($Migration)
    if ((Get-MigrationEndpointType $Migration.Config 'Target') -ne 'SPO') { return @() }
    $opsDir = Join-Path $Migration.Root 'launchers\interactive\operations'
    if (-not (Test-Path -LiteralPath $opsDir -PathType Container)) { return @() }
    $results = [System.Collections.Generic.List[object]]::new()
    foreach ($f in (Get-ChildItem -LiteralPath $opsDir -Filter '*.cmd' -File | Sort-Object Name)) {
        $raw = [System.IO.Path]::GetFileNameWithoutExtension($f.Name) -replace '^\d+-', '' -replace '-', ' '
        $results.Add([pscustomobject]@{
            DisplayName = $raw
            CmdFile     = $f.FullName
        })
    }
    return $results.ToArray()
}

function Format-ItemAge {
    param([System.IO.FileSystemInfo]$Item)
    if ($null -eq $Item) { return [pscustomobject]@{ Text = 'No run yet'; HasRun = $false } }

    $now = Get-Date
    $age = $now - $Item.LastWriteTime
    $dayOffset = ($now.Date - $Item.LastWriteTime.Date).Days
    $text = if ($dayOffset -eq 0 -and $age.TotalMinutes -lt 90) {
        '{0:HH:mm} today ({1:n0} min ago)' -f $Item.LastWriteTime, [int]$age.TotalMinutes
    } elseif ($dayOffset -eq 0) {
        '{0:HH:mm} today' -f $Item.LastWriteTime
    } elseif ($dayOffset -eq 1) {
        '{0:HH:mm} yesterday' -f $Item.LastWriteTime
    } else {
        '{0:yyyy-MM-dd HH:mm}' -f $Item.LastWriteTime
    }
    return [pscustomobject]@{ Text = $text; HasRun = $true }
}

function Open-InExplorer {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return }
    try { Invoke-Item -LiteralPath $Path } catch { Start-Process explorer.exe -ArgumentList ('"' + $Path + '"') }
}

# ---------------------------------------------------------------------------
# XAML
# ---------------------------------------------------------------------------

[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Smart SharePoint Migration"
    Width="1100" Height="760"
    MinWidth="820" MinHeight="580"
    WindowStartupLocation="CenterScreen"
    UseLayoutRounding="True"
    SnapsToDevicePixels="True"
    Background="#F5F8FB">

  <Window.Resources>
    <SolidColorBrush x:Key="Accent"       Color="#0078D4"/>
    <SolidColorBrush x:Key="AccentSoft"   Color="#E6F4FF"/>
    <SolidColorBrush x:Key="AccentText"   Color="#005A9E"/>
    <SolidColorBrush x:Key="TextPrimary"  Color="#1F2937"/>
    <SolidColorBrush x:Key="TextMuted"    Color="#5F6B7A"/>
    <SolidColorBrush x:Key="Border"       Color="#DDE7F0"/>
    <SolidColorBrush x:Key="Surface"      Color="#FFFFFF"/>
    <SolidColorBrush x:Key="BgPage"       Color="#F5F8FB"/>
    <SolidColorBrush x:Key="BgSecondary"  Color="#F0F4F8"/>
    <SolidColorBrush x:Key="WarnBg"       Color="#FFF8E6"/>
    <SolidColorBrush x:Key="WarnText"     Color="#8A5E00"/>

    <Style x:Key="Btn" TargetType="Button">
      <Setter Property="Height"          Value="30"/>
      <Setter Property="Padding"         Value="10,0"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="FontSize"        Value="12"/>
      <Setter Property="Background"      Value="Transparent"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="BorderBrush"     Value="#0078D4"/>
      <Setter Property="Foreground"      Value="#0078D4"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    CornerRadius="6" Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="#E6F4FF"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter Property="Background" Value="#CCE8FF"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.4"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="BtnGhost" TargetType="Button" BasedOn="{StaticResource Btn}">
      <Setter Property="BorderBrush" Value="#DDE7F0"/>
      <Setter Property="Foreground"  Value="#5F6B7A"/>
      <Style.Triggers>
        <Trigger Property="IsMouseOver" Value="True">
          <Setter Property="Background" Value="#F0F4F8"/>
        </Trigger>
      </Style.Triggers>
    </Style>

    <Style x:Key="Tab" TargetType="ToggleButton">
      <Setter Property="Height"          Value="36"/>
      <Setter Property="Padding"         Value="14,0"/>
      <Setter Property="Cursor"          Value="Hand"/>
      <Setter Property="FontSize"        Value="13"/>
      <Setter Property="Background"      Value="Transparent"/>
      <Setter Property="BorderThickness" Value="0,0,0,2"/>
      <Setter Property="BorderBrush"     Value="Transparent"/>
      <Setter Property="Foreground"      Value="#5F6B7A"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="ToggleButton">
            <Border Background="{TemplateBinding Background}"
                    BorderBrush="{TemplateBinding BorderBrush}"
                    BorderThickness="{TemplateBinding BorderThickness}"
                    Padding="{TemplateBinding Padding}">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsChecked" Value="True">
                <Setter Property="Foreground"  Value="#0078D4"/>
                <Setter Property="BorderBrush" Value="#0078D4"/>
                <Setter Property="FontWeight"  Value="Medium"/>
              </Trigger>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter Property="Background" Value="#F0F4F8"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>

    <Style x:Key="SectionLabel" TargetType="TextBlock">
      <Setter Property="FontSize"   Value="11"/>
      <Setter Property="FontWeight" Value="Medium"/>
      <Setter Property="Foreground" Value="#5F6B7A"/>
      <Setter Property="Margin"     Value="0,0,0,8"/>
    </Style>

    <Style x:Key="StepCard" TargetType="Border">
      <Setter Property="Background"       Value="White"/>
      <Setter Property="BorderBrush"      Value="#DDE7F0"/>
      <Setter Property="BorderThickness"  Value="1"/>
      <Setter Property="CornerRadius"     Value="8"/>
      <Setter Property="Padding"          Value="14,10"/>
      <Setter Property="Margin"           Value="0,0,0,8"/>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
    </Grid.RowDefinitions>

    <!-- Header -->
    <Border Grid.Row="0" Background="White" BorderBrush="#DDE7F0" BorderThickness="0,0,0,1" Padding="18,11">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
          <Border Width="38" Height="38" Background="#0078D4" CornerRadius="8" Margin="0,0,12,0">
            <TextBlock Text="SP" Foreground="White" FontSize="14" FontWeight="Medium"
                       HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <StackPanel VerticalAlignment="Center">
            <TextBlock Text="Smart SharePoint Migration" FontSize="15" FontWeight="Medium" Foreground="#1F2937"/>
            <TextBlock Text="SharePoint migration dashboard" FontSize="12" Foreground="#5F6B7A" Margin="0,1,0,0"/>
          </StackPanel>
        </StackPanel>
        <Image Grid.Column="1" x:Name="imgLogo" Height="30" HorizontalAlignment="Right"
               Margin="0,0,16,0" VerticalAlignment="Center" Stretch="Uniform"/>
        <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="Migration" FontSize="11" Foreground="#5F6B7A" VerticalAlignment="Center" Margin="0,0,7,0"/>
          <ComboBox x:Name="cmbMigration" Width="140" Height="30" FontSize="13" VerticalContentAlignment="Center"/>
          <Button x:Name="btnNewMigration" Content="+ New"    Style="{StaticResource BtnGhost}" Width="58" Margin="8,0,0,0"/>
          <Button x:Name="btnRefresh"      Content="Refresh"  Style="{StaticResource BtnGhost}" Width="62" Margin="6,0,0,0"/>
          <CheckBox x:Name="chkAutoRefresh" Content="Auto 30s" IsChecked="True" Margin="10,0,0,0"
                    VerticalAlignment="Center" FontSize="11" Foreground="#5F6B7A"/>
          <TextBlock x:Name="lblLastRefresh" Margin="8,0,0,0" VerticalAlignment="Center"
                     FontSize="10" Foreground="#5F6B7A"/>
        </StackPanel>
      </Grid>
    </Border>

    <!-- Context bar -->
    <Border Grid.Row="1" Background="White" BorderBrush="#DDE7F0" BorderThickness="0,0,0,1" Padding="18,7">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="Source  " FontSize="11" Foreground="#5F6B7A" VerticalAlignment="Center"/>
          <TextBlock x:Name="lblSourceType" Text="SP2019" FontSize="11" Foreground="#0078D4" VerticalAlignment="Center" Margin="0,0,5,0"/>
          <TextBlock x:Name="lblSourceUrl"  Text="-"     FontSize="12" Foreground="#1F2937" VerticalAlignment="Center"/>
        </StackPanel>
        <TextBlock Grid.Column="1" Text="  ->  " FontSize="12" Foreground="#DDE7F0"
                   HorizontalAlignment="Center" VerticalAlignment="Center"/>
        <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="Target  " FontSize="11" Foreground="#5F6B7A" VerticalAlignment="Center"/>
          <TextBlock x:Name="lblTargetType" Text="SPO" FontSize="11" Foreground="#0078D4" VerticalAlignment="Center" Margin="0,0,5,0"/>
          <TextBlock x:Name="lblTargetUrl"  Text="-"  FontSize="12" Foreground="#1F2937" VerticalAlignment="Center"/>
        </StackPanel>
        <StackPanel Grid.Column="3" Orientation="Horizontal" VerticalAlignment="Center" Margin="22,0,0,0">
          <TextBlock Text="Auth  " FontSize="11" Foreground="#5F6B7A" VerticalAlignment="Center"/>
          <ComboBox x:Name="cmbAuthMode" Width="115" Height="26" FontSize="12" VerticalContentAlignment="Center">
            <ComboBoxItem Content="Interactive" IsSelected="True"/>
            <ComboBoxItem Content="Device login"/>
            <ComboBoxItem Content="Certificate"/>
          </ComboBox>
        </StackPanel>
        <Button Grid.Column="4" x:Name="btnOpenConfig" Content="Config" Style="{StaticResource BtnGhost}"
                Width="60" Height="26" Margin="8,0,0,0"/>
      </Grid>
    </Border>

    <!-- Tab bar -->
    <Border Grid.Row="2" Background="White" BorderBrush="#DDE7F0" BorderThickness="0,0,0,1">
      <StackPanel Orientation="Horizontal" Margin="10,0">
        <ToggleButton x:Name="tabFiles"       Content="Files"       Style="{StaticResource Tab}" IsChecked="True"/>
        <ToggleButton x:Name="tabPermissions" Content="Permissions" Style="{StaticResource Tab}"/>
        <ToggleButton x:Name="tabOperations"  Content="Operations"  Style="{StaticResource Tab}"/>
        <ToggleButton x:Name="tabDiagnostics" Content="Migration Diagnostics" Style="{StaticResource Tab}"/>
        <ToggleButton x:Name="tabLogs"        Content="Logs"        Style="{StaticResource Tab}"/>
        <ToggleButton x:Name="tabConfig"      Content="Config"      Style="{StaticResource Tab}"/>
      </StackPanel>
    </Border>

    <!-- Tab content -->
    <ScrollViewer Grid.Row="3" VerticalScrollBarVisibility="Auto">
      <Grid>

        <!-- FILES -->
        <StackPanel x:Name="panelFiles" Margin="18,14" Visibility="Visible">

          <TextBlock Text="INVENTORY" Style="{StaticResource SectionLabel}"/>

          <Border Style="{StaticResource StepCard}">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <Border x:Name="numScanSrc" Grid.Column="0" Width="28" Height="28" CornerRadius="14"
                      Background="#E6F4FF" Margin="0,0,12,0" VerticalAlignment="Center">
                <TextBlock Text="1" FontSize="12" FontWeight="Medium" Foreground="#0078D4"
                           HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Border>
              <StackPanel Grid.Column="1" VerticalAlignment="Center">
                <TextBlock Text="Scan source files" FontSize="13" FontWeight="Medium" Foreground="#1F2937"/>
                <StackPanel Orientation="Horizontal" Margin="0,3,0,0">
                  <Border x:Name="badgeScanSrc" CornerRadius="10" Padding="6,2" Margin="0,0,8,0" Background="#E6F4FF">
                    <TextBlock x:Name="lblScanSrcAge" Text="No run yet" FontSize="11" Foreground="#005A9E"/>
                  </Border>
                  <ComboBox x:Name="cmbScanSrcFile" Width="430" Height="24" FontSize="11" DisplayMemberPath="Display" VerticalContentAlignment="Center"/>
                </StackPanel>
              </StackPanel>
              <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                <Button x:Name="btnOpenScanSrc" Content="Open" Style="{StaticResource BtnGhost}"
                        Width="58" Margin="0,0,6,0" Visibility="Collapsed"/>
                <Button x:Name="btnRunScanSrc"  Content="Run"  Style="{StaticResource Btn}" Width="55"/>
              </StackPanel>
            </Grid>
          </Border>

          <Border Style="{StaticResource StepCard}">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <Border Grid.Column="0" Width="28" Height="28" CornerRadius="14"
                      Background="#E6F4FF" Margin="0,0,12,0" VerticalAlignment="Center">
                <TextBlock Text="2" FontSize="12" FontWeight="Medium" Foreground="#0078D4"
                           HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Border>
              <StackPanel Grid.Column="1" VerticalAlignment="Center">
                <TextBlock Text="Scan target files" FontSize="13" FontWeight="Medium" Foreground="#1F2937"/>
                <StackPanel Orientation="Horizontal" Margin="0,3,0,0">
                  <Border x:Name="badgeScanTgt" CornerRadius="10" Padding="6,2" Margin="0,0,8,0" Background="#E6F4FF">
                    <TextBlock x:Name="lblScanTgtAge" Text="No run yet" FontSize="11" Foreground="#005A9E"/>
                  </Border>
                  <ComboBox x:Name="cmbScanTgtFile" Width="430" Height="24" FontSize="11" DisplayMemberPath="Display" VerticalContentAlignment="Center"/>
                </StackPanel>
              </StackPanel>
              <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                <Button x:Name="btnOpenScanTgt" Content="Open" Style="{StaticResource BtnGhost}"
                        Width="58" Margin="0,0,6,0" Visibility="Collapsed"/>
                <Button x:Name="btnRunScanTgt"  Content="Run"  Style="{StaticResource Btn}" Width="55"/>
              </StackPanel>
            </Grid>
          </Border>

          <Rectangle Height="1" Fill="#DDE7F0" Margin="0,4,0,14"/>
          <TextBlock Text="COMPARISON" Style="{StaticResource SectionLabel}"/>

          <Border Style="{StaticResource StepCard}">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <Border Grid.Column="0" Width="28" Height="28" CornerRadius="14"
                      Background="#E6F4FF" Margin="0,0,12,0" VerticalAlignment="Center">
                <TextBlock Text="3" FontSize="12" FontWeight="Medium" Foreground="#0078D4"
                           HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Border>
              <StackPanel Grid.Column="1" VerticalAlignment="Center">
                <TextBlock Text="Compare files (source vs target)" FontSize="13" FontWeight="Medium" Foreground="#1F2937"/>
                <StackPanel Orientation="Horizontal" Margin="0,3,0,0">
                  <Border x:Name="badgeCmpFiles" CornerRadius="10" Padding="6,2" Margin="0,0,8,0" Background="#E6F4FF">
                    <TextBlock x:Name="lblCmpFilesAge" Text="No run yet" FontSize="11" Foreground="#005A9E"/>
                  </Border>
                  <TextBlock x:Name="lblCmpFilesDir" Text="" FontSize="11" Foreground="#5F6B7A" VerticalAlignment="Center"/>
                </StackPanel>
              </StackPanel>
              <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                <Button x:Name="btnOpenCmpFiles" Content="Open" Style="{StaticResource BtnGhost}"
                        Width="58" Margin="0,0,6,0" Visibility="Collapsed"/>
                <Button x:Name="btnRunCmpFiles"  Content="Run"  Style="{StaticResource Btn}" Width="55"/>
              </StackPanel>
            </Grid>
          </Border>

          <Border Style="{StaticResource StepCard}">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <Border Grid.Column="0" Width="28" Height="28" CornerRadius="14"
                      Background="#FFF8E6" Margin="0,0,12,0" VerticalAlignment="Center">
                <TextBlock Text="4" FontSize="12" FontWeight="Medium" Foreground="#8A5E00"
                           HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Border>
              <StackPanel Grid.Column="1" VerticalAlignment="Center">
                <TextBlock Text="Compare scan history" FontSize="13" FontWeight="Medium" Foreground="#1F2937"/>
                <StackPanel Orientation="Horizontal" Margin="0,3,0,0">
                  <Border x:Name="badgeHistory" CornerRadius="10" Padding="6,2" Margin="0,0,8,0" Background="#F0F4F8">
                    <TextBlock x:Name="lblHistoryAge" Text="No run yet" FontSize="11" Foreground="#5F6B7A"/>
                  </Border>
                  <ComboBox x:Name="cmbHistorySide" Width="80" Height="24" FontSize="11" VerticalContentAlignment="Center" Margin="0,0,8,0">
                    <ComboBoxItem Content="Source" IsSelected="True"/>
                    <ComboBoxItem Content="Target"/>
                  </ComboBox>
                  <TextBlock Text="Previous" FontSize="11" Foreground="#5F6B7A" VerticalAlignment="Center" Margin="0,0,5,0"/>
                  <ComboBox x:Name="cmbHistoryOldFile" Width="270" Height="24" FontSize="11" DisplayMemberPath="Display" VerticalContentAlignment="Center" Margin="0,0,8,0"/>
                  <TextBlock Text="Current" FontSize="11" Foreground="#5F6B7A" VerticalAlignment="Center" Margin="0,0,5,0"/>
                  <ComboBox x:Name="cmbHistoryNewFile" Width="270" Height="24" FontSize="11" DisplayMemberPath="Display" VerticalContentAlignment="Center"/>
                </StackPanel>
              </StackPanel>
              <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                <Button x:Name="btnOpenHistory" Content="Open" Style="{StaticResource BtnGhost}"
                        Width="58" Margin="0,0,6,0" Visibility="Collapsed"/>
                <Button x:Name="btnRunHistory"  Content="Compare"  Style="{StaticResource Btn}" Width="70"/>
              </StackPanel>
            </Grid>
          </Border>

          <Button x:Name="btnGlobalFileReport" Content="Global file comparison report"
                  ToolTip="Generate HTML, Excel and CSV reports for the latest file comparisons"
                  Style="{StaticResource Btn}" HorizontalAlignment="Right" Padding="12,6" Margin="0,10,0,0"/>

        </StackPanel>

        <!-- PERMISSIONS -->
        <StackPanel x:Name="panelPermissions" Margin="18,14" Visibility="Collapsed">

          <TextBlock Text="INVENTORY" Style="{StaticResource SectionLabel}"/>

          <Border Style="{StaticResource StepCard}">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <Border Grid.Column="0" Width="28" Height="28" CornerRadius="14"
                      Background="#E6F4FF" Margin="0,0,12,0" VerticalAlignment="Center">
                <TextBlock Text="1" FontSize="12" FontWeight="Medium" Foreground="#0078D4"
                           HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Border>
              <StackPanel Grid.Column="1" VerticalAlignment="Center">
                <TextBlock Text="Scan source permissions" FontSize="13" FontWeight="Medium" Foreground="#1F2937"/>
                <StackPanel Orientation="Horizontal" Margin="0,3,0,0">
                  <Border x:Name="badgeScanSrcPerm" CornerRadius="10" Padding="6,2" Margin="0,0,8,0" Background="#E6F4FF">
                    <TextBlock x:Name="lblScanSrcPermAge" Text="No run yet" FontSize="11" Foreground="#005A9E"/>
                  </Border>
                  <TextBlock x:Name="lblScanSrcPermFile" Text="" FontSize="11" Foreground="#5F6B7A" VerticalAlignment="Center"/>
                </StackPanel>
              </StackPanel>
              <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                <Button x:Name="btnOpenScanSrcPerm" Content="Open" Style="{StaticResource BtnGhost}"
                        Width="58" Margin="0,0,6,0" Visibility="Collapsed"/>
                <Button x:Name="btnRunScanSrcPerm"  Content="Run"  Style="{StaticResource Btn}" Width="55"/>
              </StackPanel>
            </Grid>
          </Border>

          <Border Style="{StaticResource StepCard}">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <Border Grid.Column="0" Width="28" Height="28" CornerRadius="14"
                      Background="#E6F4FF" Margin="0,0,12,0" VerticalAlignment="Center">
                <TextBlock Text="2" FontSize="12" FontWeight="Medium" Foreground="#0078D4"
                           HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Border>
              <StackPanel Grid.Column="1" VerticalAlignment="Center">
                <TextBlock Text="Scan target permissions" FontSize="13" FontWeight="Medium" Foreground="#1F2937"/>
                <StackPanel Orientation="Horizontal" Margin="0,3,0,0">
                  <Border x:Name="badgeScanTgtPerm" CornerRadius="10" Padding="6,2" Margin="0,0,8,0" Background="#E6F4FF">
                    <TextBlock x:Name="lblScanTgtPermAge" Text="No run yet" FontSize="11" Foreground="#005A9E"/>
                  </Border>
                  <TextBlock x:Name="lblScanTgtPermFile" Text="" FontSize="11" Foreground="#5F6B7A" VerticalAlignment="Center"/>
                </StackPanel>
              </StackPanel>
              <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                <Button x:Name="btnOpenScanTgtPerm" Content="Open" Style="{StaticResource BtnGhost}"
                        Width="58" Margin="0,0,6,0" Visibility="Collapsed"/>
                <Button x:Name="btnRunScanTgtPerm"  Content="Run"  Style="{StaticResource Btn}" Width="55"/>
              </StackPanel>
            </Grid>
          </Border>

          <Rectangle Height="1" Fill="#DDE7F0" Margin="0,4,0,14"/>
          <TextBlock Text="COMPARISON" Style="{StaticResource SectionLabel}"/>

          <Border Style="{StaticResource StepCard}">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <Border Grid.Column="0" Width="28" Height="28" CornerRadius="14"
                      Background="#E6F4FF" Margin="0,0,12,0" VerticalAlignment="Center">
                <TextBlock Text="3" FontSize="12" FontWeight="Medium" Foreground="#0078D4"
                           HorizontalAlignment="Center" VerticalAlignment="Center"/>
              </Border>
              <StackPanel Grid.Column="1" VerticalAlignment="Center">
                <TextBlock Text="Compare permissions (source vs target)" FontSize="13" FontWeight="Medium" Foreground="#1F2937"/>
                <StackPanel Orientation="Horizontal" Margin="0,3,0,0">
                  <Border x:Name="badgeCmpPerms" CornerRadius="10" Padding="6,2" Margin="0,0,8,0" Background="#E6F4FF">
                    <TextBlock x:Name="lblCmpPermsAge" Text="No run yet" FontSize="11" Foreground="#005A9E"/>
                  </Border>
                  <TextBlock x:Name="lblCmpPermsDir" Text="" FontSize="11" Foreground="#5F6B7A" VerticalAlignment="Center"/>
                </StackPanel>
              </StackPanel>
              <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                <Button x:Name="btnOpenCmpPerms" Content="Open" Style="{StaticResource BtnGhost}"
                        Width="58" Margin="0,0,6,0" Visibility="Collapsed"/>
                <Button x:Name="btnRunCmpPerms"  Content="Run"  Style="{StaticResource Btn}" Width="55"/>
              </StackPanel>
            </Grid>
          </Border>

          <Button x:Name="btnGlobalPermissionsReport" Content="Global permissions comparison report"
                  ToolTip="Generate HTML, Excel and CSV reports for the latest permissions comparisons"
                  Style="{StaticResource Btn}" HorizontalAlignment="Right" Padding="12,6" Margin="0,10,0,0"/>

        </StackPanel>

        <!-- OPERATIONS -->
        <StackPanel x:Name="panelOperations" Margin="18,14" Visibility="Collapsed">
          <TextBlock Text="MIGRATION OPERATIONS" Style="{StaticResource SectionLabel}"/>
          <TextBlock x:Name="lblNoOps" Text="No operations found for this migration."
                     FontSize="13" Foreground="#5F6B7A" Visibility="Collapsed" Margin="4,0"/>
          <ItemsControl x:Name="listOps">
            <ItemsControl.ItemTemplate>
              <DataTemplate>
                <Border Background="White" BorderBrush="#DDE7F0" BorderThickness="1"
                        CornerRadius="8" Padding="14,10" Margin="0,0,0,8">
                  <Grid>
                    <Grid.ColumnDefinitions>
                      <ColumnDefinition Width="*"/>
                      <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <StackPanel Grid.Column="0" VerticalAlignment="Center">
                      <TextBlock Text="{Binding DisplayName}" FontSize="13" FontWeight="Medium" Foreground="#1F2937"/>
                      <TextBlock Text="{Binding CmdFile}" FontSize="11" Foreground="#5F6B7A" Margin="0,2,0,0"/>
                    </StackPanel>
                    <Button Grid.Column="1" Content="Run" Tag="{Binding CmdFile}"
                            Style="{StaticResource Btn}" Width="55" VerticalAlignment="Center"/>
                  </Grid>
                </Border>
              </DataTemplate>
            </ItemsControl.ItemTemplate>
          </ItemsControl>
        </StackPanel>

        <!-- MIGRATION DIAGNOSTICS -->
        <StackPanel x:Name="panelDiagnostics" Margin="18,14" Visibility="Collapsed">
          <TextBlock Text="SHAREGATE REPORT ANALYSIS" Style="{StaticResource SectionLabel}"/>
          <Border Style="{StaticResource StepCard}">
            <StackPanel>
              <TextBlock x:Name="lblDiagScope" Text="Analysis only: no ShareGate connection or migration action." Foreground="#5F6B7A" FontSize="12" Margin="0,0,0,7"/>
              <TextBox x:Name="txtDiagInput" Height="28" VerticalContentAlignment="Center" FontSize="12"/>
              <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
                <Button x:Name="btnDiagBrowseFile" Content="Browse file" Style="{StaticResource BtnGhost}" Width="94"/>
                <Button x:Name="btnDiagBrowseFolder" Content="Browse folder" Style="{StaticResource BtnGhost}" Width="104" Margin="6,0,0,0"/>
                <Button x:Name="btnDiagAnalyze" Content="Analyze reports" Style="{StaticResource Btn}" Width="112" Margin="16,0,0,0"/>
                <Button x:Name="btnDiagOpenReport" Content="Open HTML report" Style="{StaticResource BtnGhost}" Width="120" Margin="6,0,0,0" IsEnabled="False"/>
              </StackPanel>
              <TextBlock x:Name="lblDiagProgress" Text="Select a migration report folder or a CSV/XLSX file." Foreground="#5F6B7A" FontSize="11" Margin="0,8,0,0" TextWrapping="Wrap"/>
            </StackPanel>
          </Border>
          <Border Style="{StaticResource StepCard}">
            <StackPanel>
              <TextBlock Text="SUMMARY" Style="{StaticResource SectionLabel}"/>
              <TextBlock x:Name="lblDiagKpis" Text="No analysis yet." TextWrapping="Wrap" FontSize="13" Foreground="#1F2937"/>
              <TextBlock x:Name="lblDiagInterpretation" Text="Residual rates exclude Accepted issues. Fixed is a tracking state, not proof of a successful new migration." TextWrapping="Wrap" FontSize="11" Foreground="#5F6B7A" Margin="0,6,0,0"/>
            </StackPanel>
          </Border>
          <Border Style="{StaticResource StepCard}">
            <StackPanel>
              <TextBlock Text="ISSUE PATTERNS" Style="{StaticResource SectionLabel}"/>
              <StackPanel Orientation="Horizontal" Margin="0,0,0,7">
                <TextBlock Text="Session" VerticalAlignment="Center" Margin="0,0,5,0"/>
                <ComboBox x:Name="cmbDiagSession" Width="115" Height="27"/>
                <TextBlock Text="Status" VerticalAlignment="Center" Margin="10,0,5,0"/>
                <ComboBox x:Name="cmbDiagStatus" Width="115" Height="27">
                  <ComboBoxItem Content="All statuses" IsSelected="True"/>
                  <ComboBoxItem Content="Error"/>
                  <ComboBoxItem Content="Warning"/>
                </ComboBox>
                <TextBox x:Name="txtDiagFilter" Width="230" Height="27" Margin="10,0,0,0" VerticalContentAlignment="Center" ToolTip="Filter category or pattern text"/>
                <Button x:Name="btnDiagFilter" Content="Filter" Style="{StaticResource BtnGhost}" Width="62" Margin="6,0,0,0"/>
              </StackPanel>
              <DataGrid x:Name="gridDiagPatterns" Height="230" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single" CanUserAddRows="False" HeadersVisibility="Column" AlternatingRowBackground="#F7FAFE">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Category" Binding="{Binding Category}" Width="180"/>
                  <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="70"/>
                  <DataGridTextColumn Header="State" Binding="{Binding State}" Width="75"/>
                  <DataGridTextColumn Header="Lines" Binding="{Binding Lines}" Width="55"/>
                  <DataGridTextColumn Header="Items" Binding="{Binding Items}" Width="55"/>
                  <DataGridTextColumn Header="Pattern (double-click for rows)" Binding="{Binding Pattern}" Width="*"/>
                </DataGrid.Columns>
              </DataGrid>
              <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
                <TextBlock Text="Selected pattern state" VerticalAlignment="Center" Margin="0,0,7,0"/>
                <ComboBox x:Name="cmbDiagState" Width="105" Height="27">
                  <ComboBoxItem Content="To fix" IsSelected="True"/>
                  <ComboBoxItem Content="Accepted"/>
                  <ComboBoxItem Content="Fixed"/>
                </ComboBox>
                <Button x:Name="btnDiagSaveState" Content="Save state" Style="{StaticResource BtnGhost}" Width="82" Margin="7,0,0,0" IsEnabled="False"/>
                <Button x:Name="btnDiagHelp" Content="Open help link" Style="{StaticResource BtnGhost}" Width="100" Margin="7,0,0,0" IsEnabled="False"/>
              </StackPanel>
            </StackPanel>
          </Border>
          <Border Style="{StaticResource StepCard}">
            <StackPanel>
              <TextBlock Text="RAW REPORT ROWS" Style="{StaticResource SectionLabel}"/>
              <StackPanel Orientation="Horizontal" Margin="0,0,0,7">
                <TextBox x:Name="txtDiagRowFilter" Width="250" Height="27" VerticalContentAlignment="Center" ToolTip="Filter selected pattern rows by item, message, source or destination"/>
                <Button x:Name="btnDiagRowFilter" Content="Filter rows" Style="{StaticResource BtnGhost}" Width="82" Margin="6,0,0,0"/>
              </StackPanel>
              <DataGrid x:Name="gridDiagRows" Height="180" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single" CanUserAddRows="False" HeadersVisibility="Column">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Date" Binding="{Binding Timestamp}" Width="140"/>
                  <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="70"/>
                  <DataGridTextColumn Header="Type" Binding="{Binding ObjectType}" Width="95"/>
                  <DataGridTextColumn Header="Item" Binding="{Binding ItemName}" Width="180"/>
                  <DataGridTextColumn Header="Message" Binding="{Binding Message}" Width="*"/>
                </DataGrid.Columns>
              </DataGrid>
              <TextBox x:Name="txtDiagRaw" Height="105" Margin="0,7,0,0" IsReadOnly="True" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" FontFamily="Consolas" FontSize="11"/>
            </StackPanel>
          </Border>
        </StackPanel>

        <!-- CONFIG -->
        <Grid x:Name="panelConfig" Margin="18,14" Visibility="Collapsed" MinHeight="500">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <Border Grid.Row="0" Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="10,8" Margin="0,0,0,8">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <StackPanel Grid.Column="0" VerticalAlignment="Center">
                <TextBlock x:Name="lblConfigPath" Text="migration.config.psd1" FontSize="12" FontWeight="Medium" Foreground="#1F2937"/>
                <TextBlock x:Name="lblConfigStatus" Text="" FontSize="11" Foreground="#5F6B7A" Margin="0,2,0,0"/>
              </StackPanel>
              <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                <Button x:Name="btnReloadConfig" Content="Reload" Style="{StaticResource BtnGhost}" Width="64" Margin="0,0,6,0"/>
                <Button x:Name="btnSaveConfig" Content="Save" Style="{StaticResource Btn}" Width="58" Margin="0,0,6,0" IsEnabled="False"/>
                <Button x:Name="btnOpenConfigFolder" Content="Open folder" Style="{StaticResource BtnGhost}" Width="92"/>
              </StackPanel>
            </Grid>
          </Border>
          <Border Grid.Row="1" Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8">
            <TextBox x:Name="txtConfigContent"
                     FontFamily="Consolas,Courier New"
                     FontSize="12"
                     Foreground="#1F2937"
                     Background="White"
                     BorderThickness="0"
                     Padding="10"
                     AcceptsReturn="True"
                     AcceptsTab="True"
                     TextWrapping="NoWrap"
                     VerticalScrollBarVisibility="Auto"
                     HorizontalScrollBarVisibility="Auto"/>
          </Border>
        </Grid>

        <!-- LOGS -->
        <Grid x:Name="panelLogs" Margin="18,14" Visibility="Collapsed" MinHeight="400">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="340"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <StackPanel Grid.Column="0" Margin="0,0,12,0">
            <TextBlock Text="SHARED ACTIVITY" Style="{StaticResource SectionLabel}" Margin="0,0,0,8"/>
            <ListBox x:Name="listActivity" BorderBrush="#DDE7F0" BorderThickness="1"
                     FontSize="11" Height="210" ScrollViewer.HorizontalScrollBarVisibility="Auto" Margin="0,0,0,14"/>
            <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
              <TextBlock Text="LOG FILES" Style="{StaticResource SectionLabel}" Margin="0"/>
              <Button x:Name="btnRefreshLogs" Content="Refresh" Style="{StaticResource BtnGhost}"
                      Height="22" Padding="6,0" FontSize="11" Margin="8,0,0,0" VerticalAlignment="Center"/>
            </StackPanel>
            <ListBox x:Name="listLogFiles" BorderBrush="#DDE7F0" BorderThickness="1"
                     FontSize="11" ScrollViewer.HorizontalScrollBarVisibility="Disabled"
                     MaxHeight="520"/>
          </StackPanel>
          <Border Grid.Column="1" Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8">
            <Grid>
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <Border Grid.Row="0" BorderBrush="#DDE7F0" BorderThickness="0,0,0,1" Padding="10,7">
                <StackPanel Orientation="Horizontal">
                  <TextBlock x:Name="lblLogName" Text="Select a log file" FontSize="12"
                             FontWeight="Medium" Foreground="#5F6B7A" VerticalAlignment="Center"/>
                  <Button x:Name="btnOpenLogDir" Content="Open folder" Style="{StaticResource BtnGhost}"
                          Margin="12,0,0,0" Height="26" Padding="8,0" FontSize="11"/>
                  <Button x:Name="btnOpenRunLog" Content="Open run log" Style="{StaticResource BtnGhost}"
                          Margin="8,0,0,0" Height="26" Padding="8,0" FontSize="11" IsEnabled="False"/>
                </StackPanel>
              </Border>
              <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto">
                <TextBlock x:Name="txtLogContent" FontFamily="Consolas,Courier New" FontSize="11"
                           Foreground="#1F2937" TextWrapping="NoWrap" Padding="10" VerticalAlignment="Top"/>
              </ScrollViewer>
            </Grid>
          </Border>
        </Grid>

      </Grid>
    </ScrollViewer>
  </Grid>
</Window>
'@

if ($ValidateOnly) {
    try {
        $reader = [System.Xml.XmlNodeReader]::new($xaml)
        $null   = [System.Windows.Markup.XamlReader]::Load($reader)
        $null   = Show-SmartM365NewMigrationWizard -ProjectRoot $script:ScriptRoot -ValidateOnly
    } catch {
        Close-SmartM365GuiSplash -Splash $script:Splash
        throw "XAML validation failed: $_"
    }
    Close-SmartM365GuiSplash -Splash $script:Splash
    Write-Host "$($script:AppName) v$($script:AppVersion) GUI validation completed."
    exit 0
}

function Initialize-DiagnosticsLocalConfig {
    $configRoot = Join-Path $script:ScriptRoot 'Config'
    foreach ($name in @('sharegate-diagnostics.columns', 'sharegate-diagnostics.rules')) {
        $template = Join-Path $configRoot ($name + '.json.template')
        $runtime = Join-Path $configRoot ($name + '.json.txt')
        if (-not (Test-Path -LiteralPath $template -PathType Leaf)) {
            throw "Diagnostics configuration template is missing: $template"
        }
        if (Test-Path -LiteralPath $runtime -PathType Leaf) { continue }
        $temporary = Join-Path $configRoot ('.' + $name + '.' + [guid]::NewGuid().ToString('N') + '.tmp')
        try {
            [System.IO.File]::WriteAllBytes($temporary, [System.IO.File]::ReadAllBytes($template))
            try { [System.IO.File]::Move($temporary, $runtime) }
            catch {
                if (-not (Test-Path -LiteralPath $runtime -PathType Leaf)) { throw }
            }
        }
        finally {
            if (Test-Path -LiteralPath $temporary -PathType Leaf) {
                Remove-Item -LiteralPath $temporary -Force
            }
        }
    }
}

try { Initialize-DiagnosticsLocalConfig }
catch {
    Close-SmartM365GuiSplash -Splash $script:Splash
    [System.Windows.MessageBox]::Show("Failed to initialize local diagnostics configuration:`n$_", $script:AppName, 'OK', 'Error') | Out-Null
    exit 1
}

# ---------------------------------------------------------------------------
# Load window
# ---------------------------------------------------------------------------

try {
    $reader        = [System.Xml.XmlNodeReader]::new($xaml)
    $script:Window = [System.Windows.Markup.XamlReader]::Load($reader)
} catch {
    Close-SmartM365GuiSplash -Splash $script:Splash
    [System.Windows.MessageBox]::Show("Failed to load GUI:`n$_", $script:AppName, 'OK', 'Error')
    exit 1
}

function ctrl { param([string]$n) $script:Window.FindName($n) }

# Header
$cmbMigration    = ctrl 'cmbMigration'
$btnRefresh      = ctrl 'btnRefresh'
$chkAutoRefresh = ctrl 'chkAutoRefresh'
$lblLastRefresh = ctrl 'lblLastRefresh'
$btnNewMigration = ctrl 'btnNewMigration'
$imgLogo         = ctrl 'imgLogo'

# Context bar
$lblSourceType = ctrl 'lblSourceType'
$lblSourceUrl  = ctrl 'lblSourceUrl'
$lblTargetType = ctrl 'lblTargetType'
$lblTargetUrl  = ctrl 'lblTargetUrl'
$cmbAuthMode   = ctrl 'cmbAuthMode'
$btnOpenConfig = ctrl 'btnOpenConfig'

# Tabs
$tabFiles       = ctrl 'tabFiles'
$tabPermissions = ctrl 'tabPermissions'
$tabOperations  = ctrl 'tabOperations'
$tabDiagnostics = ctrl 'tabDiagnostics'
$tabLogs        = ctrl 'tabLogs'
$tabConfig      = ctrl 'tabConfig'

# Panels
$panelFiles       = ctrl 'panelFiles'
$panelPermissions = ctrl 'panelPermissions'
$panelOperations  = ctrl 'panelOperations'
$panelDiagnostics = ctrl 'panelDiagnostics'
$panelLogs        = ctrl 'panelLogs'
$panelConfig      = ctrl 'panelConfig'

# Files step
$badgeScanSrc  = ctrl 'badgeScanSrc'
$lblScanSrcAge = ctrl 'lblScanSrcAge'
$cmbScanSrcFile= ctrl 'cmbScanSrcFile'
$btnOpenScanSrc= ctrl 'btnOpenScanSrc'
$btnRunScanSrc = ctrl 'btnRunScanSrc'

$badgeScanTgt  = ctrl 'badgeScanTgt'
$lblScanTgtAge = ctrl 'lblScanTgtAge'
$cmbScanTgtFile= ctrl 'cmbScanTgtFile'
$btnOpenScanTgt= ctrl 'btnOpenScanTgt'
$btnRunScanTgt = ctrl 'btnRunScanTgt'

$badgeCmpFiles  = ctrl 'badgeCmpFiles'
$lblCmpFilesAge = ctrl 'lblCmpFilesAge'
$lblCmpFilesDir = ctrl 'lblCmpFilesDir'
$btnOpenCmpFiles= ctrl 'btnOpenCmpFiles'
$btnRunCmpFiles = ctrl 'btnRunCmpFiles'

$badgeHistory  = ctrl 'badgeHistory'
$lblHistoryAge = ctrl 'lblHistoryAge'
$btnOpenHistory= ctrl 'btnOpenHistory'
$btnRunHistory = ctrl 'btnRunHistory'
$cmbHistorySide = ctrl 'cmbHistorySide'
$cmbHistoryOldFile = ctrl 'cmbHistoryOldFile'
$cmbHistoryNewFile = ctrl 'cmbHistoryNewFile'

# Permissions step
$badgeScanSrcPerm  = ctrl 'badgeScanSrcPerm'
$lblScanSrcPermAge = ctrl 'lblScanSrcPermAge'
$lblScanSrcPermFile= ctrl 'lblScanSrcPermFile'
$btnOpenScanSrcPerm= ctrl 'btnOpenScanSrcPerm'
$btnRunScanSrcPerm = ctrl 'btnRunScanSrcPerm'

$badgeScanTgtPerm  = ctrl 'badgeScanTgtPerm'
$lblScanTgtPermAge = ctrl 'lblScanTgtPermAge'
$lblScanTgtPermFile= ctrl 'lblScanTgtPermFile'
$btnOpenScanTgtPerm= ctrl 'btnOpenScanTgtPerm'
$btnRunScanTgtPerm = ctrl 'btnRunScanTgtPerm'

$badgeCmpPerms  = ctrl 'badgeCmpPerms'
$lblCmpPermsAge = ctrl 'lblCmpPermsAge'
$lblCmpPermsDir = ctrl 'lblCmpPermsDir'
$btnOpenCmpPerms= ctrl 'btnOpenCmpPerms'
$btnRunCmpPerms = ctrl 'btnRunCmpPerms'

# Operations
$listOps   = ctrl 'listOps'
$lblNoOps  = ctrl 'lblNoOps'

# Migration diagnostics
$lblDiagScope = ctrl 'lblDiagScope'
$txtDiagInput = ctrl 'txtDiagInput'
$btnDiagBrowseFile = ctrl 'btnDiagBrowseFile'
$btnDiagBrowseFolder = ctrl 'btnDiagBrowseFolder'
$btnDiagAnalyze = ctrl 'btnDiagAnalyze'
$btnDiagOpenReport = ctrl 'btnDiagOpenReport'
$lblDiagProgress = ctrl 'lblDiagProgress'
$lblDiagKpis = ctrl 'lblDiagKpis'
$lblDiagInterpretation = ctrl 'lblDiagInterpretation'
$cmbDiagSession = ctrl 'cmbDiagSession'
$cmbDiagStatus = ctrl 'cmbDiagStatus'
$txtDiagFilter = ctrl 'txtDiagFilter'
$btnDiagFilter = ctrl 'btnDiagFilter'
$gridDiagPatterns = ctrl 'gridDiagPatterns'
$gridDiagRows = ctrl 'gridDiagRows'
$txtDiagRowFilter = ctrl 'txtDiagRowFilter'
$btnDiagRowFilter = ctrl 'btnDiagRowFilter'
$txtDiagRaw = ctrl 'txtDiagRaw'
$cmbDiagState = ctrl 'cmbDiagState'
$btnDiagSaveState = ctrl 'btnDiagSaveState'
$btnDiagHelp = ctrl 'btnDiagHelp'

# Logs
$listLogFiles  = ctrl 'listLogFiles'
$listActivity = ctrl 'listActivity'
$lblLogName    = ctrl 'lblLogName'
$txtLogContent = ctrl 'txtLogContent'
$btnOpenLogDir = ctrl 'btnOpenLogDir'
$btnRefreshLogs= ctrl 'btnRefreshLogs'
$btnOpenRunLog = ctrl 'btnOpenRunLog'
$btnGlobalFileReport = ctrl 'btnGlobalFileReport'
$btnGlobalPermissionsReport = ctrl 'btnGlobalPermissionsReport'

# Config
$lblConfigPath = ctrl 'lblConfigPath'
$lblConfigStatus = ctrl 'lblConfigStatus'
$txtConfigContent = ctrl 'txtConfigContent'
$btnReloadConfig = ctrl 'btnReloadConfig'
$btnSaveConfig = ctrl 'btnSaveConfig'
$btnOpenConfigFolder = ctrl 'btnOpenConfigFolder'

# ---------------------------------------------------------------------------
# State
# ---------------------------------------------------------------------------

$script:Migrations        = @()
$script:CurrentMigration  = $null
$script:CurrentStatus     = $null
$script:ConfigEditorLoading = $false
$script:ConfigEditorDirty   = $false
$script:UpdateCheckTimer     = $null
$script:AutoRefreshTimer = $null
$script:WizardOpen = $false
$script:TargetScopeMismatch = $false
$script:ConfigEditorLoadedHash = ''
$script:DiagInputPath = ''
$script:DiagSummary = $null
$script:DiagRows = @()
$script:DiagKnownSessions = @()
$script:DiagProcess = $null
$script:DiagTimer = $null
$script:DiagOutputDirectory = ''
$script:DiagActivity = ''
$script:DiagProjectRoot = ''

# ---------------------------------------------------------------------------
# Logo / icon
# ---------------------------------------------------------------------------

$logoFile = Join-Path $script:ScriptRoot 'WorkplaceCloudHub-lockup-WPF.png'
if (Test-Path -LiteralPath $logoFile -PathType Leaf) {
    try {
        $bmp = [System.Windows.Media.Imaging.BitmapImage]::new()
        $bmp.BeginInit()
        $bmp.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bmp.UriSource   = [System.Uri]::new($logoFile)
        $bmp.EndInit()
        $bmp.Freeze()
        $imgLogo.Source = $bmp
    } catch {}
}

$iconFile = Join-Path $script:ScriptRoot 'WorkplaceCloudHub.ico'
if (Test-Path -LiteralPath $iconFile -PathType Leaf) {
    try {
        $script:Window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create(
            [System.Uri]::new($iconFile))
    } catch {}
}

# ---------------------------------------------------------------------------
# Badge helpers
# ---------------------------------------------------------------------------

function Set-Badge {
    param(
        [System.Windows.Controls.Border]$Badge,
        [System.Windows.Controls.TextBlock]$Label,
        [string]$Text,
        [bool]$HasRun
    )
    $Label.Text = $Text
    if ($HasRun) {
        $Badge.Background = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(230, 244, 255))
        $Label.Foreground = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(0, 90, 158))
    } else {
        $Badge.Background = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(240, 244, 248))
        $Label.Foreground = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(95, 107, 122))
    }
}
function Set-ScanComboItems {
    param(
        [System.Windows.Controls.ComboBox]$ComboBox,
        [object[]]$Items,
        [System.IO.FileInfo]$SelectedFile
    )

    $previous = if ($ComboBox.SelectedItem) { [string]$ComboBox.SelectedItem.FullName } else { '' }
    $ComboBox.Items.Clear()
    foreach ($item in @($Items)) { [void]$ComboBox.Items.Add($item) }
    $ComboBox.IsEnabled = ($ComboBox.Items.Count -gt 0)
    if ($ComboBox.Items.Count -eq 0) {
        $ComboBox.SelectedIndex = -1
        return
    }

    $selectedIndex = 0
    if ($previous) {
        for ($i = 0; $i -lt $ComboBox.Items.Count; $i++) {
            if ([string]$ComboBox.Items[$i].FullName -eq $previous) {
                $selectedIndex = $i
                break
            }
        }
    }
    elseif ($SelectedFile) {
        for ($i = 0; $i -lt $ComboBox.Items.Count; $i++) {
            if ([string]$ComboBox.Items[$i].FullName -eq [string]$SelectedFile.FullName) {
                $selectedIndex = $i
                break
            }
        }
    }
    $ComboBox.SelectedIndex = $selectedIndex
}

function Get-SelectedScanFile {
    param([System.Windows.Controls.ComboBox]$ComboBox)
    if ($null -eq $ComboBox -or $null -eq $ComboBox.SelectedItem) { return $null }
    return $ComboBox.SelectedItem.File
}

function Update-ScanFileSelection {
    param(
        [System.Windows.Controls.ComboBox]$ComboBox,
        [System.Windows.Controls.Border]$Badge,
        [System.Windows.Controls.TextBlock]$Label,
        [System.Windows.Controls.Button]$OpenButton
    )

    $selectedFile = Get-SelectedScanFile -ComboBox $ComboBox
    $age = Format-ItemAge $selectedFile
    Set-Badge $Badge $Label $age.Text $age.HasRun
    $OpenButton.Visibility = if ($selectedFile) { 'Visible' } else { 'Collapsed' }
    if ($selectedFile) { $OpenButton.Tag = (Split-Path $selectedFile.FullName -Parent) }
}
function Get-HistorySide {
    $sel = $cmbHistorySide.SelectedItem
    if ($null -eq $sel) { return 'Source' }
    $txt = if ($sel.PSObject.Properties.Name -contains 'Content') { [string]$sel.Content } else { [string]$sel }
    if ($txt -eq 'Target') { return 'Target' }
    return 'Source'
}

function Set-HistoryComboItems {
    param(
        [System.Windows.Controls.ComboBox]$ComboBox,
        [object[]]$Items,
        [int]$DefaultIndex
    )

    $previous = if ($ComboBox.SelectedItem) { [string]$ComboBox.SelectedItem.FullName } else { '' }
    $ComboBox.Items.Clear()
    foreach ($item in @($Items)) { [void]$ComboBox.Items.Add($item) }
    $ComboBox.IsEnabled = ($ComboBox.Items.Count -gt 0)
    if ($ComboBox.Items.Count -eq 0) {
        $ComboBox.SelectedIndex = -1
        return
    }

    $previousIndex = -1
    if ($previous) {
        for ($i = 0; $i -lt $ComboBox.Items.Count; $i++) {
            if ([string]$ComboBox.Items[$i].FullName -eq $previous) { $previousIndex = $i; break }
        }
    }
    if ($previousIndex -ge 0) {
        $ComboBox.SelectedIndex = $previousIndex
    }
    elseif ($DefaultIndex -ge 0 -and $DefaultIndex -lt $ComboBox.Items.Count) {
        $ComboBox.SelectedIndex = $DefaultIndex
    }
    else {
        $ComboBox.SelectedIndex = 0
    }
}

function Update-HistoryRunState {
    $oldCsv = Get-SelectedScanFile -ComboBox $cmbHistoryOldFile
    $newCsv = Get-SelectedScanFile -ComboBox $cmbHistoryNewFile
    $btnRunHistory.IsEnabled = ($oldCsv -and $newCsv -and $oldCsv.FullName -ne $newCsv.FullName)
}

function Update-HistoryScanSelection {
    if ($null -eq $script:CurrentStatus) { return }

    $items = if ((Get-HistorySide) -eq 'Target') { @($script:CurrentStatus.TargetFileCsvItems) } else { @($script:CurrentStatus.SourceFileCsvItems) }
    Set-HistoryComboItems $cmbHistoryNewFile $items 0
    Set-HistoryComboItems $cmbHistoryOldFile $items 1
    Update-HistoryRunState
}

# ---------------------------------------------------------------------------
# Auth mode
# ---------------------------------------------------------------------------

function Get-AuthMode {
    $sel = $cmbAuthMode.SelectedItem
    if ($null -eq $sel) { return 'Interactive' }
    $txt = if ($sel.PSObject.Properties.Name -contains 'Content') { [string]$sel.Content } else { [string]$sel }
    switch ($txt) {
        'Device login' { return 'DeviceLogin' }
        'Certificate'  { return 'Certificate' }
        default        { return 'Interactive' }
    }
}

# ---------------------------------------------------------------------------
# Run action (new window)
# ---------------------------------------------------------------------------

function Get-SelectedMigrationFolderName {
    if ($null -eq $script:CurrentMigration) { return '' }
    return [string]$script:CurrentMigration.Name
}

function Invoke-MigrationAction {
    param([string]$Action, [string]$OperationPath = '')
    if ($null -eq $script:CurrentMigration) { return }

    if ($OperationPath -and $script:TargetScopeMismatch) {
        [System.Windows.MessageBox]::Show(
            'Target.SiteUrl differs from the target mapping. Align the configuration before running a site operation.',
            $script:AppName, 'OK', 'Warning') | Out-Null
        return
    }
    $launcher = Join-Path $script:ScriptRoot 'Scripts\Launchers\Generic\SmartM365-SharePointMigration-GuiRun.ps1'
    $exe      = Join-Path $PSHOME 'pwsh.exe'

    $migName = Get-SelectedMigrationFolderName
    try {
        $activity = New-SmartM365GuiActivity -ProjectRoot $script:ScriptRoot -Migration $migName `
            -Action $(if ($OperationPath) { [System.IO.Path]::GetFileNameWithoutExtension($OperationPath) } else { $Action })
    }
    catch {
        [System.Windows.MessageBox]::Show("Could not create shared activity log:`n$($_.Exception.Message)",
            $script:AppName, 'OK', 'Error') | Out-Null
        return
    }
    $args    = @('-NoExit', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$launcher`"",
                 '-ActivityPath', "`"$activity`"", '-MigrationName', "`"$migName`"")
    if ($OperationPath) { $args += @('-OperationPath', "`"$OperationPath`"") }
    else { $args += @('-Action', $Action) }

    switch (Get-AuthMode) {
        'DeviceLogin' { $args += @('-AuthMode', 'DeviceLogin') }
        'Certificate' { $args += @('-AuthMode', 'Certificate') }
    }

    if ($Action -eq 'CompareFiles') {
        $sourceCsv = Get-SelectedScanFile -ComboBox $cmbScanSrcFile
        $targetCsv = Get-SelectedScanFile -ComboBox $cmbScanTgtFile
        if ($sourceCsv) { $args += @('-SourceCsv', "`"$($sourceCsv.FullName)`"") }
        if ($targetCsv) { $args += @('-TargetCsv', "`"$($targetCsv.FullName)`"") }
    }
    elseif ($Action -eq 'CompareScanHistory') {
        $oldCsv = Get-SelectedScanFile -ComboBox $cmbHistoryOldFile
        $newCsv = Get-SelectedScanFile -ComboBox $cmbHistoryNewFile
        $args += @('-HistorySide', (Get-HistorySide))
        if ($oldCsv) { $args += @('-OldCsv', "`"$($oldCsv.FullName)`"") }
        if ($newCsv) { $args += @('-NewCsv', "`"$($newCsv.FullName)`"") }
    }

    try {
        Start-Process -FilePath $exe -ArgumentList $args -WorkingDirectory $script:ScriptRoot -ErrorAction Stop
        Refresh-ActivityList
    }
    catch {
        Write-SmartM365GuiActivityEvent -Path $activity -Status 'Failed' -ExitCode 1 `
            -Detail $_.Exception.Message
        [System.Windows.MessageBox]::Show("Could not launch action:`n$($_.Exception.Message)",
            $script:AppName, 'OK', 'Error') | Out-Null
    }
}

# ---------------------------------------------------------------------------
# Tab switching
# ---------------------------------------------------------------------------

function Switch-Tab {
    param([string]$Tab)
    $tabFiles.IsChecked       = ($Tab -eq 'Files')
    $tabPermissions.IsChecked = ($Tab -eq 'Permissions')
    $tabOperations.IsChecked  = ($Tab -eq 'Operations')
    $tabDiagnostics.IsChecked = ($Tab -eq 'Diagnostics')
    $tabLogs.IsChecked        = ($Tab -eq 'Logs')
    $tabConfig.IsChecked      = ($Tab -eq 'Config')

    $panelFiles.Visibility       = if ($Tab -eq 'Files')       { 'Visible' } else { 'Collapsed' }
    $panelPermissions.Visibility = if ($Tab -eq 'Permissions') { 'Visible' } else { 'Collapsed' }
    $panelOperations.Visibility  = if ($Tab -eq 'Operations')  { 'Visible' } else { 'Collapsed' }
    $panelDiagnostics.Visibility = if ($Tab -eq 'Diagnostics') { 'Visible' } else { 'Collapsed' }
    $panelLogs.Visibility        = if ($Tab -eq 'Logs')        { 'Visible' } else { 'Collapsed' }
    $panelConfig.Visibility      = if ($Tab -eq 'Config')      { 'Visible' } else { 'Collapsed' }
}

# ---------------------------------------------------------------------------
# Update UI
# ---------------------------------------------------------------------------

function Update-UI {
    if ($null -eq $script:CurrentMigration) { return }

    $cfg = $script:CurrentMigration.Config
    $st  = $script:CurrentStatus

    $lblSourceType.Text = Get-MigrationEndpointType $cfg 'Source'
    $sourceScope = Get-MigrationScope $script:CurrentMigration 'Source'
    $targetScope = Get-MigrationScope $script:CurrentMigration 'Target'
    $lblSourceUrl.Text  = $sourceScope.Text
    $lblSourceUrl.ToolTip = $sourceScope.Tooltip
    $lblTargetType.Text = Get-MigrationEndpointType $cfg 'Target'
    $lblTargetUrl.Text  = $targetScope.Text
    $lblTargetUrl.ToolTip = $targetScope.Tooltip
    $script:TargetScopeMismatch = $targetScope.Mismatch
    $lblTargetUrl.Foreground = if ($targetScope.Mismatch) {
        [System.Windows.Media.Brushes]::Firebrick
    } else { [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(31, 41, 55)) }
    if ($targetScope.Mismatch) { $lblTargetUrl.Text += '  [CONFIG DIFFERS]' }

    # --- Files ---
    Set-ScanComboItems $cmbScanSrcFile @($st.SourceFileCsvItems) $st.SourceFileCsv
    Update-ScanFileSelection $cmbScanSrcFile $badgeScanSrc $lblScanSrcAge $btnOpenScanSrc

    Set-ScanComboItems $cmbScanTgtFile @($st.TargetFileCsvItems) $st.TargetFileCsv
    Update-ScanFileSelection $cmbScanTgtFile $badgeScanTgt $lblScanTgtAge $btnOpenScanTgt

    $r = Format-ItemAge $st.FileComparisonFolder
    Set-Badge $badgeCmpFiles $lblCmpFilesAge $r.Text $r.HasRun
    $lblCmpFilesDir.Text    = if ($st.FileComparisonFolder) { $st.FileComparisonFolder.Name } else { '' }
    $btnOpenCmpFiles.Visibility = if ($st.FileComparisonFolder) { 'Visible' } else { 'Collapsed' }
    if ($st.FileComparisonFolder) { $btnOpenCmpFiles.Tag = $st.FileComparisonFolder.FullName }

    $r = Format-ItemAge $st.HistoryFolder
    Set-Badge $badgeHistory $lblHistoryAge $r.Text $r.HasRun
    $btnOpenHistory.Visibility = if ($st.HistoryFolder) { 'Visible' } else { 'Collapsed' }
    if ($st.HistoryFolder) { $btnOpenHistory.Tag = $st.HistoryFolder.FullName }
    Update-HistoryScanSelection

    # --- Permissions ---
    $r = Format-ItemAge $st.SourcePermCsv
    Set-Badge $badgeScanSrcPerm $lblScanSrcPermAge $r.Text $r.HasRun
    $lblScanSrcPermFile.Text = if ($st.SourcePermCsv) { $st.SourcePermCsv.Name } else { '' }
    $btnOpenScanSrcPerm.Visibility = if ($st.SourcePermCsv) { 'Visible' } else { 'Collapsed' }
    if ($st.SourcePermCsv) { $btnOpenScanSrcPerm.Tag = (Split-Path $st.SourcePermCsv.FullName -Parent) }

    $r = Format-ItemAge $st.TargetPermCsv
    Set-Badge $badgeScanTgtPerm $lblScanTgtPermAge $r.Text $r.HasRun
    $lblScanTgtPermFile.Text = if ($st.TargetPermCsv) { $st.TargetPermCsv.Name } else { '' }
    $btnOpenScanTgtPerm.Visibility = if ($st.TargetPermCsv) { 'Visible' } else { 'Collapsed' }
    if ($st.TargetPermCsv) { $btnOpenScanTgtPerm.Tag = (Split-Path $st.TargetPermCsv.FullName -Parent) }

    $r = Format-ItemAge $st.PermComparisonFolder
    Set-Badge $badgeCmpPerms $lblCmpPermsAge $r.Text $r.HasRun
    $lblCmpPermsDir.Text = if ($st.PermComparisonFolder) { $st.PermComparisonFolder.Name } else { '' }
    $btnOpenCmpPerms.Visibility = if ($st.PermComparisonFolder) { 'Visible' } else { 'Collapsed' }
    if ($st.PermComparisonFolder) { $btnOpenCmpPerms.Tag = $st.PermComparisonFolder.FullName }

    # --- Operations ---
    $ops = @(Get-MigrationOperations -Migration $script:CurrentMigration)
    $listOps.Items.Clear()
    foreach ($op in $ops) { [void]$listOps.Items.Add($op) }
    $lblNoOps.Visibility = if (@($ops).Count -eq 0) { 'Visible' } else { 'Collapsed' }

    # --- Logs ---
    Refresh-LogList
}

function Set-ConfigEditorStatus {
    param(
        [string]$Text,
        [bool]$IsError = $false
    )

    $lblConfigStatus.Text = $Text
    if ($IsError) {
        $lblConfigStatus.Foreground = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(176, 0, 32))
    }
    else {
        $lblConfigStatus.Foreground = [System.Windows.Media.SolidColorBrush]::new(
            [System.Windows.Media.Color]::FromRgb(95, 107, 122))
    }
}

function Set-ConfigEditorDirty {
    param([bool]$Dirty)

    $script:ConfigEditorDirty = $Dirty
    $btnSaveConfig.IsEnabled = $Dirty
}

function Test-ConfigEditorContent {
    param([string]$Content)

    $tempPath = [System.IO.Path]::ChangeExtension([System.IO.Path]::GetTempFileName(), '.psd1')
    try {
        $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
        [System.IO.File]::WriteAllText($tempPath, $Content, $utf8NoBom)
        $data = Import-PowerShellDataFile -LiteralPath $tempPath
        if ($null -eq $data -or -not ($data -is [hashtable])) {
            throw 'The file must contain a PowerShell data hashtable.'
        }
        foreach ($requiredKey in @('Name', 'Source', 'Target', 'Comparison', 'Output')) {
            if (-not $data.ContainsKey($requiredKey)) {
                throw "Missing required key: $requiredKey"
            }
        }
        return $data
    }
    finally {
        Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
    }
}

function Load-ConfigEditor {
    if ($null -eq $script:CurrentMigration) { return }

    $script:ConfigEditorLoading = $true
    try {
        $configPath = $script:CurrentMigration.ConfigPath
        $lblConfigPath.Text = $configPath
        if (Test-Path -LiteralPath $configPath -PathType Leaf) {
            $txtConfigContent.Text = Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop
            $script:ConfigEditorLoadedHash = (Get-FileHash -LiteralPath $configPath -Algorithm SHA256).Hash
            Set-ConfigEditorDirty $false
            Set-ConfigEditorStatus ("Loaded {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
        }
        else {
            $script:ConfigEditorLoadedHash = ''
            $txtConfigContent.Text = ''
            Set-ConfigEditorDirty $false
            Set-ConfigEditorStatus "Config file not found: $configPath" $true
        }
    }
    catch {
        $script:ConfigEditorLoadedHash = ''
        $txtConfigContent.Text = ''
        Set-ConfigEditorDirty $false
        Set-ConfigEditorStatus "Could not load config: $($_.Exception.Message)" $true
    }
    finally {
        $script:ConfigEditorLoading = $false
    }
}

function Save-ConfigEditor {
    if ($null -eq $script:CurrentMigration) { return }

    $configActivity = $null
    try {
        $content = [string]$txtConfigContent.Text
        $validatedConfig = Test-ConfigEditorContent -Content $content
        $configPath = $script:CurrentMigration.ConfigPath
        $configActivity = New-SmartM365GuiActivity -ProjectRoot $script:ScriptRoot `
            -Migration $script:CurrentMigration.Name -Action 'SaveConfig'
        if (Test-Path -LiteralPath $configPath -PathType Leaf) {
            $currentHash = (Get-FileHash -LiteralPath $configPath -Algorithm SHA256).Hash
            if ($script:ConfigEditorLoadedHash -and $currentHash -ne $script:ConfigEditorLoadedHash) {
                throw 'The configuration changed on disk since it was loaded. Review and reload it before saving.'
            }
            $backupPath = "{0}.bak-{1}-{2}" -f $configPath, (Get-Date -Format 'yyyyMMdd-HHmmss'), [guid]::NewGuid().ToString('N')
            Copy-Item -LiteralPath $configPath -Destination $backupPath -Force
        }

        $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
        [System.IO.File]::WriteAllText($configPath, $content, $utf8NoBom)
        $script:ConfigEditorLoadedHash = (Get-FileHash -LiteralPath $configPath -Algorithm SHA256).Hash
        $script:CurrentMigration.Config = $validatedConfig
        $script:CurrentStatus = Get-MigrationStatus -Migration $script:CurrentMigration
        Set-ConfigEditorDirty $false
        Update-UI
        Set-ConfigEditorStatus ("Saved {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
        Write-SmartM365GuiActivityEvent -Path $configActivity -Status 'Succeeded' -ExitCode 0 `
            -Detail 'Migration configuration saved.'
    }
    catch {
        try {
            if (-not $configActivity) {
                $configActivity = New-SmartM365GuiActivity -ProjectRoot $script:ScriptRoot `
                    -Migration $script:CurrentMigration.Name -Action 'SaveConfig'
            }
            Write-SmartM365GuiActivityEvent -Path $configActivity -Status 'Failed' -ExitCode 1 `
                -Detail $_.Exception.Message
        } catch { [void]$_.Exception }
        Set-ConfigEditorStatus "Save failed: $($_.Exception.Message)" $true
        [System.Windows.MessageBox]::Show("Config save failed:`n$($_.Exception.Message)", $script:AppName, 'OK', 'Error') | Out-Null
    }
}
function Refresh-LogList {
    $previous = if ($listLogFiles.SelectedItem) { [string]$listLogFiles.SelectedItem.FullName } else { '' }
    $listLogFiles.Items.Clear()
    if ($null -eq $script:CurrentMigration) { return }
    $logsDir = Join-Path $script:CurrentMigration.Root 'logs'
    if (-not (Test-Path -LiteralPath $logsDir -PathType Container)) { return }
    Get-ChildItem -LiteralPath $logsDir -Filter '*.log' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 60 |
        ForEach-Object {
            $display = ('{0}  {1}' -f $_.LastWriteTime.ToString('yyyy-MM-dd HH:mm'), $_.Name)
            $item = [pscustomobject]@{
                Display  = $display
                FullName = $_.FullName
                Name     = $_.Name
            }
            [void]$listLogFiles.Items.Add($item)
            if ($previous -and $item.FullName -eq $previous) { $listLogFiles.SelectedItem = $item }
        }
}

function Refresh-ActivityList {
    $previous = if ($listActivity.SelectedItem) { [string]$listActivity.SelectedItem.FullName } else { '' }
    $listActivity.Items.Clear()
    $directory = Get-SmartM365GuiActivityDirectory -ProjectRoot $script:ScriptRoot
    Get-ChildItem -LiteralPath $directory -Filter '*.log' -File -ErrorAction Stop |
        Sort-Object Name -Descending | Select-Object -First 100 | ForEach-Object {
            try {
                $item = Read-SmartM365GuiActivity -Path $_.FullName
                if ($item) {
                    [void]$listActivity.Items.Add($item)
                    if ($previous -and $item.FullName -eq $previous) { $listActivity.SelectedItem = $item }
                }
            } catch { [void]$_.Exception }
        }
}

function Refresh-GuiState {
    try {
        if (-not (Test-Path -LiteralPath (Join-Path $script:ScriptRoot 'Migrations') -PathType Container)) {
            throw 'Migration folder is unavailable.'
        }
        Load-Migrations
        Refresh-ActivityList
        $lblLastRefresh.Text = 'Updated ' + (Get-Date -Format 'HH:mm:ss')
        $lblLastRefresh.ToolTip = 'Shared activity and migration state refreshed.'
    }
    catch {
        $lblLastRefresh.Text = 'Refresh failed'
        $lblLastRefresh.ToolTip = $_.Exception.Message
    }
}

function Set-CurrentMigration {
    param($Migration)
    $sameConfig = $null -ne $script:CurrentMigration -and
        [string]::Equals($script:CurrentMigration.ConfigPath, $Migration.ConfigPath,
            [System.StringComparison]::OrdinalIgnoreCase)
    if ($script:ConfigEditorDirty -and -not $sameConfig) {
        $answer = [System.Windows.MessageBox]::Show(
            'Discard unsaved config changes before switching migrations?',
            $script:AppName, 'YesNo', 'Warning')
        if ($answer -ne [System.Windows.MessageBoxResult]::Yes) {
            if ($null -ne $script:CurrentMigration) {
                $cmbMigration.SelectedItem = $script:CurrentMigration.Name
            }
            return
        }
    }
    $script:CurrentMigration = $Migration
    if (-not $sameConfig) {
        $script:DiagInputPath = Join-Path $Migration.Root 'ShareGate\MigrationReport'
        $txtDiagInput.Text = $script:DiagInputPath
        $script:DiagSummary = $null
        $script:DiagRows = @()
        $script:DiagKnownSessions = @()
        $gridDiagPatterns.ItemsSource = $null
        $gridDiagRows.ItemsSource = $null
        $txtDiagRaw.Text = ''
        $lblDiagKpis.Text = 'No analysis yet.'
        $lblDiagProgress.Text = 'Select a migration report folder or a CSV/XLSX file.'
        $btnDiagOpenReport.IsEnabled = $false
        $cmbDiagSession.Items.Clear()
        [void]$cmbDiagSession.Items.Add('All sessions')
        $cmbDiagSession.SelectedIndex = 0
    }
    $script:CurrentStatus    = Get-MigrationStatus -Migration $Migration
    Update-UI
    if (-not $sameConfig -or -not $script:ConfigEditorDirty) {
        Load-ConfigEditor
    }
}

function Get-MigrationScope {
    param($Migration, [ValidateSet('Source', 'Target')][string]$Side)
    $cfg = $Migration.Config
    $section = if ($Side -eq 'Source') { $cfg.Source } else { $cfg.Target }
    $urls = [System.Collections.Generic.List[string]]::new()
    $urlsFile = if ($section.ContainsKey('UrlsFile')) { [string]$section.UrlsFile } else { '' }
    if ($urlsFile) {
        $path = if ([System.IO.Path]::IsPathRooted($urlsFile)) { $urlsFile } else { Join-Path $Migration.Root $urlsFile }
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            foreach ($line in [System.IO.File]::ReadAllLines($path)) {
                $value = $line.Trim()
                if ($value -and -not $value.StartsWith('#') -and -not $urls.Contains($value)) { $urls.Add($value) }
            }
        }
    }
    elseif ($cfg.Comparison.PathMappingsFile) {
        $mapping = [string]$cfg.Comparison.PathMappingsFile
        $path = if ([System.IO.Path]::IsPathRooted($mapping)) { $mapping } else { Join-Path $Migration.Root $mapping }
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            foreach ($line in [System.IO.File]::ReadAllLines($path)) {
                $value = $line.Trim()
                if (-not $value -or $value.StartsWith('#')) { continue }
                $parts = @($value -split '[\t;, ]+' | Where-Object { $_ })
                if ($parts.Count -ne 2) { continue }
                $url = if ($Side -eq 'Source') { $parts[0] } else { $parts[1] }
                if (-not $urls.Contains($url)) { $urls.Add($url) }
            }
        }
    }
    if ($urls.Count -eq 0) {
        $fallback = Get-MigrationEndpointUrlText $cfg $Side
        if ($fallback) { $urls.Add($fallback) }
    }
    $configured = if ($section.ContainsKey('SiteUrl')) { [string]$section.SiteUrl } else { '' }
    $mismatch = $false
    if ($Side -eq 'Target' -and $configured -and $urls.Count -gt 0) {
        $matched = $false
        foreach ($url in $urls) {
            if ([string]::Equals($url.TrimEnd('/'), $configured.TrimEnd('/'),
                    [System.StringComparison]::OrdinalIgnoreCase)) { $matched = $true; break }
        }
        $mismatch = -not $matched
    }
    [pscustomobject]@{
        Text = if ($urls.Count -eq 1) { $urls[0] } else { '{0} sites (hover for URLs)' -f $urls.Count }
        Tooltip = (($urls.ToArray() -join "`n") + $(if ($mismatch) { "`nConfig SiteUrl differs: $configured" } else { '' }))
        Mismatch = $mismatch
    }
}

function Load-Migrations {
    $prev = if ($cmbMigration.SelectedItem) { [string]$cmbMigration.SelectedItem } else { $null }
    $script:Migrations = @(Get-MigrationFolders)
    $cmbMigration.Items.Clear()
    foreach ($m in $script:Migrations) { [void]$cmbMigration.Items.Add($m.Name) }
    if ($script:Migrations.Count -eq 0) { return }
    $idx = if ($prev) { $cmbMigration.Items.IndexOf($prev) } else { -1 }
    $cmbMigration.SelectedIndex = if ($idx -ge 0) { $idx } else { 0 }
}

# ---------------------------------------------------------------------------
# Events
# ---------------------------------------------------------------------------

$tabFiles.Add_Click({       Switch-Tab 'Files' })
$tabPermissions.Add_Click({ Switch-Tab 'Permissions' })
$tabOperations.Add_Click({  Switch-Tab 'Operations' })
$tabDiagnostics.Add_Click({ Switch-Tab 'Diagnostics' })
$tabLogs.Add_Click({        Switch-Tab 'Logs' })
$tabConfig.Add_Click({      Switch-Tab 'Config' })

$cmbMigration.Add_SelectionChanged({
    $idx = $cmbMigration.SelectedIndex
    if ($idx -ge 0 -and $idx -lt $script:Migrations.Count) {
        Set-CurrentMigration -Migration $script:Migrations[$idx]
    }
})

$btnRefresh.Add_Click({
    Refresh-GuiState
})

$btnNewMigration.Add_Click({
    $script:WizardOpen = $true
    try {
        $activity = New-SmartM365GuiActivity -ProjectRoot $script:ScriptRoot -Migration '<new>' -Action 'NewMigration'
        $createdName = Show-SmartM365NewMigrationWizard -Owner $script:Window `
            -ProjectRoot $script:ScriptRoot -AuthMode (Get-AuthMode) -ActivityPath $activity
        if ($createdName) {
            Write-SmartM365GuiActivityEvent -Path $activity -Status 'Succeeded' `
                -Detail "Created migration $createdName after URL validation." -ExitCode 0 -Migration $createdName
            Refresh-GuiState
            $cmbMigration.SelectedItem = $createdName
        }
    }
    catch {
        if ($activity) {
            Write-SmartM365GuiActivityEvent -Path $activity -Status 'Failed' -ExitCode 1 `
                -Detail $_.Exception.Message
        }
        [System.Windows.MessageBox]::Show("Could not open new migration wizard:`n$($_.Exception.Message)",
            $script:AppName, 'OK', 'Error') | Out-Null
    }
    finally {
        $script:WizardOpen = $false
        try { Refresh-ActivityList }
        catch { $lblLastRefresh.ToolTip = $_.Exception.Message }
    }
})

$btnOpenConfig.Add_Click({
    Switch-Tab 'Config'
})

$btnReloadConfig.Add_Click({
    Load-ConfigEditor
})

$btnSaveConfig.Add_Click({
    Save-ConfigEditor
})

$btnOpenConfigFolder.Add_Click({
    if ($null -eq $script:CurrentMigration) { return }
    Open-InExplorer (Split-Path -Path $script:CurrentMigration.ConfigPath -Parent)
})

$txtConfigContent.Add_TextChanged({
    if (-not $script:ConfigEditorLoading) {
        Set-ConfigEditorDirty $true
        Set-ConfigEditorStatus 'Modified'
    }
})

# Files
$btnRunScanSrc.Add_Click({  Invoke-MigrationAction 'ScanSourceFiles' })
$btnRunScanTgt.Add_Click({  Invoke-MigrationAction 'ScanTargetFiles' })
$btnRunCmpFiles.Add_Click({ Invoke-MigrationAction 'CompareFiles' })
$btnRunHistory.Add_Click({  Invoke-MigrationAction 'CompareScanHistory' })

$cmbScanSrcFile.Add_SelectionChanged({ Update-ScanFileSelection $cmbScanSrcFile $badgeScanSrc $lblScanSrcAge $btnOpenScanSrc })
$cmbScanTgtFile.Add_SelectionChanged({ Update-ScanFileSelection $cmbScanTgtFile $badgeScanTgt $lblScanTgtAge $btnOpenScanTgt })
$cmbHistorySide.Add_SelectionChanged({ Update-HistoryScanSelection })
$cmbHistoryOldFile.Add_SelectionChanged({ Update-HistoryRunState })
$cmbHistoryNewFile.Add_SelectionChanged({ Update-HistoryRunState })
$btnOpenScanSrc.Add_Click({  Open-InExplorer ([string]$btnOpenScanSrc.Tag) })
$btnOpenScanTgt.Add_Click({  Open-InExplorer ([string]$btnOpenScanTgt.Tag) })
$btnOpenCmpFiles.Add_Click({ Open-InExplorer ([string]$btnOpenCmpFiles.Tag) })
$btnOpenHistory.Add_Click({  Open-InExplorer ([string]$btnOpenHistory.Tag) })

# Permissions
$btnRunScanSrcPerm.Add_Click({  Invoke-MigrationAction 'ScanSourcePermissions' })
$btnRunScanTgtPerm.Add_Click({  Invoke-MigrationAction 'ScanTargetPermissions' })
$btnRunCmpPerms.Add_Click({     Invoke-MigrationAction 'ComparePermissions' })

$btnOpenScanSrcPerm.Add_Click({ Open-InExplorer ([string]$btnOpenScanSrcPerm.Tag) })
$btnOpenScanTgtPerm.Add_Click({ Open-InExplorer ([string]$btnOpenScanTgtPerm.Tag) })
$btnOpenCmpPerms.Add_Click({    Open-InExplorer ([string]$btnOpenCmpPerms.Tag) })

# Operations - buttons inside DataTemplate handled via bubbled RoutedEvent
$listOps.AddHandler(
    [System.Windows.Controls.Button]::ClickEvent,
    [System.Windows.RoutedEventHandler]{
        param($s, $e)
        $btn = $e.OriginalSource
        if ($btn -is [System.Windows.Controls.Button] -and -not [string]::IsNullOrWhiteSpace([string]$btn.Tag)) {
            $cmd = [string]$btn.Tag
            if (Test-Path -LiteralPath $cmd -PathType Leaf) {
                Invoke-MigrationAction -Action 'Operation' -OperationPath $cmd
            }
        }
    }
)

# Logs
$listLogFiles.Add_SelectionChanged({
    $sel = $listLogFiles.SelectedItem
    if ($null -eq $sel) { return }
    $lblLogName.Text = $sel.Name
    try {
        $txtLogContent.Text = Get-Content -LiteralPath $sel.FullName -Raw -ErrorAction Stop
    } catch {
        $txtLogContent.Text = "Could not read log file: $_"
    }
})

$btnOpenLogDir.Add_Click({
    if ($null -eq $script:CurrentMigration) { return }
    Open-InExplorer (Join-Path $script:CurrentMigration.Root 'logs')
})

$btnRefreshLogs.Add_Click({ Refresh-LogList })

function Invoke-GlobalComparisonReport {
    param([ValidateSet('Files', 'Permissions')][string]$Kind)

    $activity = $null
    $reportLabel = if ($Kind -eq 'Files') { 'files' } else { 'permissions' }
    try {
        $action = if ($Kind -eq 'Files') { 'GlobalFileComparisonReport' } else { 'GlobalPermissionsComparisonReport' }
        $activity = New-SmartM365GuiActivity -ProjectRoot $script:ScriptRoot -Migration '<all>' -Action $action
        $python = Join-Path $script:ScriptRoot 'Tools\Python\python.exe'
        if (-not (Test-Path -LiteralPath $python -PathType Leaf)) {
            $python = (Get-Command python -ErrorAction Stop).Source
        }
        $generatorName = if ($Kind -eq 'Files') { 'build_global_report.py' } else { 'build_global_permissions_report.py' }
        $generator = Join-Path $script:ScriptRoot "Scripts\Compare\$generatorName"
        $migrationRoot = Join-Path $script:ScriptRoot 'Migrations'
        $reportDirectory = Join-Path $migrationRoot 'reports\global'
        if ($Kind -eq 'Permissions') { $reportDirectory = Join-Path $reportDirectory 'permissions' }
        $result = @(& $python $generator --migrations-root $migrationRoot --output-directory $reportDirectory 2>&1)
        if ($LASTEXITCODE -ne 0) { throw ($result -join "`n") }
        $reportPath = [string]($result | Select-Object -Last 1)
        if (-not (Test-Path -LiteralPath $reportPath -PathType Leaf)) { throw "Report was not created: $reportPath" }
        $excelPath = [System.IO.Path]::ChangeExtension($reportPath, '.xlsx')
        if (-not (Test-Path -LiteralPath $excelPath -PathType Leaf)) { throw "Excel report was not created: $excelPath" }
        Write-SmartM365GuiActivityEvent -Path $activity -Status 'Succeeded' -ExitCode 0 -Detail "Global $reportLabel comparison report: $reportPath; Excel: $excelPath"
        Open-InExplorer $reportPath
    }
    catch {
        if ($activity) { Write-SmartM365GuiActivityEvent -Path $activity -Status 'Failed' -ExitCode 1 -Detail $_.Exception.Message }
        [System.Windows.MessageBox]::Show("Could not generate global $reportLabel comparison report:`n$($_.Exception.Message)", $script:AppName, 'OK', 'Error') | Out-Null
    }
    finally { Refresh-ActivityList }
}

$btnGlobalFileReport.Add_Click({ Invoke-GlobalComparisonReport -Kind 'Files' })
$btnGlobalPermissionsReport.Add_Click({ Invoke-GlobalComparisonReport -Kind 'Permissions' })

function Refresh-DiagnosticPatterns {
    if (-not $script:DiagSummary) { return }
    $status = if ($cmbDiagStatus.SelectedItem) { [string]$cmbDiagStatus.SelectedItem.Content } else { 'All statuses' }
    $search = [string]$txtDiagFilter.Text
    $matching = @($script:DiagSummary.Patterns | Where-Object {
        ($status -eq 'All statuses' -or $_.Status -eq $status) -and
        (-not $search -or $_.Category.Contains($search, [System.StringComparison]::OrdinalIgnoreCase) -or
            $_.Pattern.Contains($search, [System.StringComparison]::OrdinalIgnoreCase))
    } | ForEach-Object { [pscustomobject]$_ })
    $gridDiagPatterns.ItemsSource = $matching
    $gridDiagRows.ItemsSource = $null
    $txtDiagRaw.Text = ''
    $btnDiagSaveState.IsEnabled = $false
    $btnDiagHelp.IsEnabled = $false
}

function Refresh-DiagnosticRows {
    $pattern = $gridDiagPatterns.SelectedItem
    if (-not $pattern) { $gridDiagRows.ItemsSource = $null; return }
    $search = [string]$txtDiagRowFilter.Text
    $gridDiagRows.ItemsSource = @($script:DiagRows | Where-Object {
        $_.PatternKey -eq $pattern.PatternKey -and
        (-not $search -or $_.ItemName.Contains($search, [System.StringComparison]::OrdinalIgnoreCase) -or
            $_.Message.Contains($search, [System.StringComparison]::OrdinalIgnoreCase) -or
            $_.SourceUrl.Contains($search, [System.StringComparison]::OrdinalIgnoreCase) -or
            $_.DestinationUrl.Contains($search, [System.StringComparison]::OrdinalIgnoreCase))
    })
    $txtDiagRaw.Text = ''
}

function Load-DiagnosticResult {
    param([string]$Directory)
    $summaryPath = Join-Path $Directory 'Summary.json.txt'
    if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf)) { throw "Analysis summary was not created: $summaryPath" }
    $script:DiagSummary = Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json -AsHashtable
    $script:DiagRows = @(Import-Csv -LiteralPath $script:DiagSummary.RowsPath)
    $selected = if ($cmbDiagSession.SelectedItem) { [string]$cmbDiagSession.SelectedItem } else { 'All sessions' }
    $script:DiagKnownSessions = @($script:DiagKnownSessions + @($script:DiagSummary.Sessions) | Sort-Object -Unique)
    $cmbDiagSession.Items.Clear()
    [void]$cmbDiagSession.Items.Add('All sessions')
    foreach ($session in $script:DiagKnownSessions) { [void]$cmbDiagSession.Items.Add([string]$session) }
    $index = $cmbDiagSession.Items.IndexOf($selected)
    $cmbDiagSession.SelectedIndex = if ($index -ge 0) { $index } else { 0 }
    $lines = $script:DiagSummary.Lines
    $lineStates = $script:DiagSummary.IssueLineState
    $itemStates = $script:DiagSummary.IssueItemState
    $lineRate = if ($null -ne $script:DiagSummary.ResidualLineRate) { '{0:N2}%' -f [double]$script:DiagSummary.ResidualLineRate } else { 'n/a' }
    $itemRate = if ($null -ne $script:DiagSummary.ResidualItemRate) { '{0:N2}%' -f [double]$script:DiagSummary.ResidualItemRate } else { 'n/a' }
    $lblDiagKpis.Text = ('Lines: {0} | Success: {1} | Error: {2} | Warning: {3} | Accepted: {4} | To fix: {5}`nDistinct keyed items: {6} | Unkeyed lines: {7} | Items to fix: {8} | Residual lines: {9} | Residual items: {10}' -f
        $lines, $script:DiagSummary.LineStatus.Success, $script:DiagSummary.LineStatus.Error,
        $script:DiagSummary.LineStatus.Warning, $lineStates.Accepted, $lineStates['To fix'],
        $script:DiagSummary.DistinctItems, $script:DiagSummary.UnkeyedRows, $itemStates['To fix'],
        $lineRate, $itemRate).Replace('`n', "`n")
    $btnDiagOpenReport.IsEnabled = (Test-Path -LiteralPath $script:DiagSummary.ReportPath -PathType Leaf)
    $lblDiagProgress.Text = "Analysis completed: $($script:DiagSummary.Sessions.Count) session(s); $($script:DiagSummary.DuplicateRowsSuppressed) duplicate rows suppressed."
    if ($script:DiagSummary.ConflictingDuplicateRows -gt 0) {
        $lblDiagProgress.Text += " $($script:DiagSummary.ConflictingDuplicateRows) conflicting duplicate rows require review."
    }
    Refresh-DiagnosticPatterns
}

function Start-DiagnosticAnalysis {
    if ($script:DiagProcess -and -not $script:DiagProcess.HasExited) { return }
    if (-not $script:CurrentMigration) { return }
    $inputPath = [string]$txtDiagInput.Text
    if (-not (Test-Path -LiteralPath $inputPath)) {
        [System.Windows.MessageBox]::Show("Report path does not exist:`n$inputPath", $script:AppName, 'OK', 'Warning') | Out-Null
        return
    }
    $wrapper = Join-Path $script:ScriptRoot 'Scripts\Diagnostics\SmartM365-SharePointMigration-Diagnostics.ps1'
    $output = Join-Path $script:CurrentMigration.Root ('ShareGate\Diagnostics\{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [guid]::NewGuid().ToString('N'))
    $activity = $null
    try {
        New-Item -ItemType Directory -Path $output -Force | Out-Null
        $activity = New-SmartM365GuiActivity -ProjectRoot $script:ScriptRoot -Migration $script:CurrentMigration.Name -Action 'MigrationDiagnostics'
        $arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$wrapper`"",
            '-ProjectRoot', "`"$($script:CurrentMigration.Root)`"", '-InputPath', "`"$inputPath`"",
            '-OutputDirectory', "`"$output`"", '-ActivityPath', "`"$activity`"")
        $session = if ($cmbDiagSession.SelectedItem) { [string]$cmbDiagSession.SelectedItem } else { 'All sessions' }
        if ($session -ne 'All sessions') { $arguments += @('-SessionId', "`"$session`"") }
        $exe = Join-Path $PSHOME 'pwsh.exe'
        $script:DiagProcess = Start-Process -FilePath $exe -ArgumentList $arguments -WorkingDirectory $script:ScriptRoot -WindowStyle Hidden -PassThru `
            -RedirectStandardOutput (Join-Path $output 'analysis.stdout.log') -RedirectStandardError (Join-Path $output 'analysis.stderr.log') -ErrorAction Stop
        $script:DiagOutputDirectory = $output
        $script:DiagActivity = $activity
        $script:DiagProjectRoot = $script:CurrentMigration.Root
        $btnDiagAnalyze.IsEnabled = $false
        $lblDiagProgress.Text = "Analyzing local report files in $inputPath ..."
        $script:DiagTimer.Start()
        Refresh-ActivityList
    }
    catch {
        if ($activity) { Write-SmartM365GuiActivityEvent -Path $activity -Status 'Failed' -ExitCode 1 -Detail $_.Exception.Message }
        $lblDiagProgress.Text = "Could not start analysis: $($_.Exception.Message)"
        $btnDiagAnalyze.IsEnabled = $true
    }
}

$btnDiagBrowseFile.Add_Click({
    $dialog = [Microsoft.Win32.OpenFileDialog]::new()
    $dialog.Filter = 'ShareGate reports (*.csv;*.xlsx)|*.csv;*.xlsx|All files (*.*)|*.*'
    if ($dialog.ShowDialog($script:Window)) { $txtDiagInput.Text = $dialog.FileName }
})
$btnDiagBrowseFolder.Add_Click({
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = [System.Windows.Forms.FolderBrowserDialog]::new()
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtDiagInput.Text = $dialog.SelectedPath }
    $dialog.Dispose()
})
$btnDiagAnalyze.Add_Click({ Start-DiagnosticAnalysis })
$btnDiagOpenReport.Add_Click({ if ($script:DiagSummary) { Open-InExplorer $script:DiagSummary.ReportPath } })
$btnDiagFilter.Add_Click({ Refresh-DiagnosticPatterns })
$btnDiagRowFilter.Add_Click({ Refresh-DiagnosticRows })
$cmbDiagStatus.Add_SelectionChanged({ if ($script:DiagSummary) { Refresh-DiagnosticPatterns } })
$cmbDiagSession.Add_SelectionChanged({
    if ($script:DiagSummary -and $cmbDiagSession.SelectedItem) { $lblDiagProgress.Text = 'Click Analyze reports to apply the selected session.' }
})
$gridDiagPatterns.Add_SelectionChanged({
    $pattern = $gridDiagPatterns.SelectedItem
    $btnDiagSaveState.IsEnabled = ($null -ne $pattern)
    $btnDiagHelp.IsEnabled = ($null -ne $pattern -and [string]$pattern.HelpLink -match '^https?://')
    if ($pattern) {
        foreach ($item in $cmbDiagState.Items) {
            if ([string]$item.Content -eq [string]$pattern.State) { $cmbDiagState.SelectedItem = $item; break }
        }
    }
})
$gridDiagPatterns.Add_MouseDoubleClick({
    Refresh-DiagnosticRows
})
$gridDiagRows.Add_SelectionChanged({
    if ($gridDiagRows.SelectedItem) { $txtDiagRaw.Text = ($gridDiagRows.SelectedItem | ConvertTo-Json -Depth 4) }
})
$btnDiagHelp.Add_Click({
    $pattern = $gridDiagPatterns.SelectedItem
    if ($pattern -and [string]$pattern.HelpLink -match '^https?://') { Start-Process -FilePath ([string]$pattern.HelpLink) }
})
$btnDiagSaveState.Add_Click({
    $pattern = $gridDiagPatterns.SelectedItem
    if (-not $pattern -or -not $script:CurrentMigration) { return }
    $state = [string]$cmbDiagState.SelectedItem.Content
    $activity = $null
    try {
        $activity = New-SmartM365GuiActivity -ProjectRoot $script:ScriptRoot -Migration $script:CurrentMigration.Name -Action 'DiagnosticPatternState'
        $wrapper = Join-Path $script:ScriptRoot 'Scripts\Diagnostics\SmartM365-SharePointMigration-Diagnostics.ps1'
        $exe = Join-Path $PSHOME 'pwsh.exe'
        $result = @(& $exe -NoProfile -ExecutionPolicy Bypass -File $wrapper -ProjectRoot $script:CurrentMigration.Root -PatternKey $pattern.PatternKey -SetPatternState $state -ExpectedState $pattern.State -ActivityPath $activity 2>&1)
        if ($LASTEXITCODE -ne 0) { throw ($result -join "`n") }
        $lblDiagProgress.Text = "Pattern state saved as $state. Reanalyzing to update KPIs."
        Start-DiagnosticAnalysis
    }
    catch {
        if ($activity) { Write-SmartM365GuiActivityEvent -Path $activity -Status 'Failed' -ExitCode 1 -Detail $_.Exception.Message }
        [System.Windows.MessageBox]::Show("Could not save pattern state:`n$($_.Exception.Message)", $script:AppName, 'OK', 'Error') | Out-Null
    }
    finally { Refresh-ActivityList }
})

$listActivity.Add_SelectionChanged({
    $sel = $listActivity.SelectedItem
    if ($null -eq $sel) { return }
    $lblLogName.Text = $sel.Display
    $btnOpenRunLog.IsEnabled = [bool]($sel.LogPath -and (Test-Path -LiteralPath $sel.LogPath -PathType Leaf))
    $btnOpenRunLog.Tag = $sel.LogPath
    try { $txtLogContent.Text = Get-Content -LiteralPath $sel.FullName -Raw -ErrorAction Stop }
    catch { $txtLogContent.Text = "Could not read activity log: $_" }
})

$btnOpenRunLog.Add_Click({
    if ($btnOpenRunLog.Tag -and (Test-Path -LiteralPath $btnOpenRunLog.Tag -PathType Leaf)) {
        Start-Process -FilePath 'notepad.exe' -ArgumentList "`"$($btnOpenRunLog.Tag)`""
    }
})

# ---------------------------------------------------------------------------
# Init and show
# ---------------------------------------------------------------------------

if (Get-Command -Name Start-SmartM365GuiUpdateCheck -ErrorAction SilentlyContinue) {
    $script:Window.Add_ContentRendered({
        try {
            $manifestPath = Join-Path $script:ScriptRoot 'SmartM365.GuiUpdateCheck.psd1'
            $script:UpdateCheckTimer = Start-SmartM365GuiUpdateCheck -Owner $script:Window -ManifestPath $manifestPath -AppRoot $script:ScriptRoot -OnStatus {
                param(
                    [string]$Message,
                    [string]$Title
                )
                $null = $Message
                $script:Window.Title = ("{0} - {1}" -f $script:AppName, $Title)
            }
        }
        catch { [void]$_.Exception }
    })
}

$script:AutoRefreshTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:AutoRefreshTimer.Interval = [TimeSpan]::FromSeconds(30)
$script:AutoRefreshTimer.Add_Tick({
    if ($chkAutoRefresh.IsChecked -and -not $script:WizardOpen) { Refresh-GuiState }
})
$chkAutoRefresh.Add_Checked({ if ($script:AutoRefreshTimer) { $script:AutoRefreshTimer.Start() } })
$chkAutoRefresh.Add_Unchecked({ if ($script:AutoRefreshTimer) { $script:AutoRefreshTimer.Stop() } })
$script:DiagTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:DiagTimer.Interval = [TimeSpan]::FromSeconds(1)
$script:DiagTimer.Add_Tick({
    if (-not $script:DiagProcess -or -not $script:DiagProcess.HasExited) { return }
    $script:DiagTimer.Stop()
    $code = $script:DiagProcess.ExitCode
    $script:DiagProcess.Dispose()
    $script:DiagProcess = $null
    $btnDiagAnalyze.IsEnabled = $true
    try {
        if ($code -ne 0) {
            $stderrPath = Join-Path $script:DiagOutputDirectory 'analysis.stderr.log'
            $errorText = if (Test-Path -LiteralPath $stderrPath -PathType Leaf) { Get-Content -LiteralPath $stderrPath -Raw } else { '' }
            throw "Analysis exited with code $code. $errorText"
        }
        if ($script:CurrentMigration -and $script:CurrentMigration.Root -eq $script:DiagProjectRoot) {
            Load-DiagnosticResult -Directory $script:DiagOutputDirectory
        }
    }
    catch {
        $lblDiagProgress.Text = "Analysis failed: $($_.Exception.Message)"
        if ($script:DiagActivity) {
            Write-SmartM365GuiActivityEvent -Path $script:DiagActivity -Status 'Failed' -ExitCode 1 -Detail $_.Exception.Message
        }
    }
    Refresh-ActivityList
})
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    $lblDiagScope.Text = 'Analysis only. CSV is available; XLSX needs the optional ImportExcel module.'
}
try {
    $script:SessionActivity = New-SmartM365GuiActivity -ProjectRoot $script:ScriptRoot `
        -Migration '<gui>' -Action 'GuiSession'
}
catch {
    Close-SmartM365GuiSplash -Splash $script:Splash
    [System.Windows.MessageBox]::Show(
        "The shared activity folder is unavailable or not writable:`n$($_.Exception.Message)",
        $script:AppName, 'OK', 'Error') | Out-Null
    exit 1
}
Refresh-GuiState
$script:AutoRefreshTimer.Start()
Close-SmartM365GuiSplash -Splash $script:Splash
try { [void]$script:Window.ShowDialog() }
finally {
    $script:AutoRefreshTimer.Stop()
    if ($script:DiagTimer) { $script:DiagTimer.Stop() }
    Write-SmartM365GuiActivityEvent -Path $script:SessionActivity -Status 'Closed' -ExitCode 0 `
        -Detail 'GUI window closed.'
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAd7wWeC+v37qvC
# YT5otXKYmkSyAWQZ+LVP8WehcqMp06CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIABiXgy9Qoxkp+X8s+eWY0d9WgeSjVF82Xnf3xgCd0lZMA0GCSqG
# SIb3DQEBAQUABIIBgDxGa/Q+TejoW4JUGSnDNobnTthdCQgvCTcrd3NSaXdQw0eh
# XDN2Xqz8J4yihDZkgG01lXZyWm/QfJ6PtEr7RcvksOEcfT6eMfmX/44svM/M6OHs
# hsjqf18PIRwuv7juziqcqtRQ90mkHjsqq3Ol9FtUnmpC3rr+TvWQ5zeCHj7hde3N
# WA3y4SJIV5CyNjPOqaaTgNs4u1JGROlLfr7XGsyAzmQcUOu7JJV22doMIY2u8ufm
# 0nJq0LOXe0Pohb/a15oogksGPCq8Cq6WuvsoXSEqoev8Cj28+3nObkzo4dq0SaIW
# pkI7/GlwNFZeI9/8V0uCSxzrsV5uclNcmuWuL+JpWVvR4TVoAw+ilod8M0TISHSN
# E8762BqSdHqvvWUdPur2PkTYuMAaufq+12qXAS8D+iwrVZvhAtnUeOv8DhHEPCdv
# v3uWCH6fDRq2hr1Ix7dfv4WLGiRSo2uoa72jTnjjldISGAlXzCUA4KthixC5w/wg
# pVEMnR3kPRUXQ4kE8aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMwMDE5
# MTJaMC8GCSqGSIb3DQEJBDEiBCA1Ioq6FBdlDOEnwP6do6/X1uK1GKtv6U6X37DN
# 6bmOBjANBgkqhkiG9w0BAQEFAASCAgBlLFRAXT+HCOC05yv+izpD0aY7qfJZ3dfo
# 1ASYQi8MjBpzF2JrNo0cIzTGjv/Z+nhVfxrK+7EiEUXUwTMCceHiPKF31Zm6IcUI
# niTzdj1zpJFFzb6ULBDHFjlBqWVvm4+u8CNhS/zLlhlNLZ1zuvygrrLG/mkrJbyW
# J7uE9g5gJSSCl2Y1jImwRMUTpNYpOiyqeBxbTI5yeOuCvRYN1TOw9KkSnkeDnoBU
# riQPp9yaOnc6fG6XloUIqDkD04uYI0WkqPMwjdyoAdXK46+7obrF3MsGkhe7/6Oc
# ymQrkJ0ZKwOl6xnhs65s6sr6lnXYyUmtYUcxCRUPF5aq+QzFZ9O8MvfAGQ9MPn+y
# tFOB3yO867N5E+PZBodUu5I2tkEoxuIBFUm6ZKEWNZ0gq3o9oaLVCGTq3YaCtryL
# ouzmgyxRkZ0ZRf4TZj4OiEr52ymVhu31HF83XeswZ0esxEFldGipCB6rDtl1JmCP
# iOJO5BBtmQXAvnWecn3pLCiRowp2njOvYL25wZWGD41pp+XCvXjzLWRx1cuw2Nvy
# c3WSHRZ335a7Jsc11m1juV0OOOP6kfzAgrSWh0HnwKQrXcYiskkdN4n7YD2clXeV
# /GWhnuM2dOEKARCNnSGZnwlt6+Kyr+ifay9vzuuKHuKEBSLZHL6UvVS1YuFrQY0L
# IzW9zERMhA==
# SIG # End signature block
