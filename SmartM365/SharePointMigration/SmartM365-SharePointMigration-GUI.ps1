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
    1.0.7
#>

#Requires -Version 5.1

[CmdletBinding()]
param(
    [switch]$ValidateOnly
)

$script:AppName    = 'Smart SharePoint Migration'
$script:AppVersion = '1.0.7'
$script:ScriptRoot = $PSScriptRoot

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

. (Join-Path $script:ScriptRoot 'SmartM365.GuiSplash.ps1')

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
        if ($dir.Name -eq '_Template') { continue }
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
    Width="980" Height="720"
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
            <ColumnDefinition Width="210"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <StackPanel Grid.Column="0" Margin="0,0,12,0">
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
    } catch {
        Close-SmartM365GuiSplash -Splash $script:Splash
        throw "XAML validation failed: $_"
    }
    Close-SmartM365GuiSplash -Splash $script:Splash
    Write-Host "$($script:AppName) v$($script:AppVersion) GUI validation completed."
    exit 0
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
$tabLogs        = ctrl 'tabLogs'
$tabConfig      = ctrl 'tabConfig'

# Panels
$panelFiles       = ctrl 'panelFiles'
$panelPermissions = ctrl 'panelPermissions'
$panelOperations  = ctrl 'panelOperations'
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

# Logs
$listLogFiles  = ctrl 'listLogFiles'
$lblLogName    = ctrl 'lblLogName'
$txtLogContent = ctrl 'txtLogContent'
$btnOpenLogDir = ctrl 'btnOpenLogDir'
$btnRefreshLogs= ctrl 'btnRefreshLogs'

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

    $ComboBox.Items.Clear()
    foreach ($item in @($Items)) { [void]$ComboBox.Items.Add($item) }
    $ComboBox.IsEnabled = ($ComboBox.Items.Count -gt 0)
    if ($ComboBox.Items.Count -eq 0) {
        $ComboBox.SelectedIndex = -1
        return
    }

    $selectedIndex = 0
    if ($SelectedFile) {
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

    $ComboBox.Items.Clear()
    foreach ($item in @($Items)) { [void]$ComboBox.Items.Add($item) }
    $ComboBox.IsEnabled = ($ComboBox.Items.Count -gt 0)
    if ($ComboBox.Items.Count -eq 0) {
        $ComboBox.SelectedIndex = -1
        return
    }

    if ($DefaultIndex -ge 0 -and $DefaultIndex -lt $ComboBox.Items.Count) {
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

function Invoke-MigrationAction {
    param([string]$Action)
    if ($null -eq $script:CurrentMigration) { return }

    $launcher = Join-Path $script:ScriptRoot 'Scripts\Launchers\Generic\SmartM365-SharePointMigration-Launcher.ps1'
    $pwsh7    = Join-Path $env:ProgramFiles 'PowerShell\7\pwsh.exe'
    $exe      = if (Test-Path -LiteralPath $pwsh7 -PathType Leaf) { $pwsh7 } else { 'powershell.exe' }

    $migName = $script:CurrentMigration.Config.Name
    $args    = @('-NoExit', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$launcher`"",
                 '-MigrationName', "`"$migName`"", '-Action', $Action)

    switch (Get-AuthMode) {
        'DeviceLogin' { $args += '-DeviceLogin' }
        'Certificate' { $args += '-UseCertificate' }
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

    Start-Process -FilePath $exe -ArgumentList $args -WorkingDirectory $script:ScriptRoot
}

# ---------------------------------------------------------------------------
# Tab switching
# ---------------------------------------------------------------------------

function Switch-Tab {
    param([string]$Tab)
    $tabFiles.IsChecked       = ($Tab -eq 'Files')
    $tabPermissions.IsChecked = ($Tab -eq 'Permissions')
    $tabOperations.IsChecked  = ($Tab -eq 'Operations')
    $tabLogs.IsChecked        = ($Tab -eq 'Logs')
    $tabConfig.IsChecked      = ($Tab -eq 'Config')

    $panelFiles.Visibility       = if ($Tab -eq 'Files')       { 'Visible' } else { 'Collapsed' }
    $panelPermissions.Visibility = if ($Tab -eq 'Permissions') { 'Visible' } else { 'Collapsed' }
    $panelOperations.Visibility  = if ($Tab -eq 'Operations')  { 'Visible' } else { 'Collapsed' }
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
    $lblSourceUrl.Text  = Get-MigrationEndpointUrlText $cfg 'Source'
    $lblTargetType.Text = Get-MigrationEndpointType $cfg 'Target'
    $lblTargetUrl.Text  = Get-MigrationEndpointUrlText $cfg 'Target'

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
    $ops = Get-MigrationOperations -Migration $script:CurrentMigration
    $listOps.Items.Clear()
    foreach ($op in $ops) { [void]$listOps.Items.Add($op) }
    $lblNoOps.Visibility = if ($ops.Count -eq 0) { 'Visible' } else { 'Collapsed' }

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
            Set-ConfigEditorDirty $false
            Set-ConfigEditorStatus ("Loaded {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
        }
        else {
            $txtConfigContent.Text = ''
            Set-ConfigEditorDirty $false
            Set-ConfigEditorStatus "Config file not found: $configPath" $true
        }
    }
    catch {
        Set-ConfigEditorStatus "Could not load config: $($_.Exception.Message)" $true
    }
    finally {
        $script:ConfigEditorLoading = $false
    }
}

function Save-ConfigEditor {
    if ($null -eq $script:CurrentMigration) { return }

    try {
        $content = [string]$txtConfigContent.Text
        $validatedConfig = Test-ConfigEditorContent -Content $content
        $configPath = $script:CurrentMigration.ConfigPath
        if (Test-Path -LiteralPath $configPath -PathType Leaf) {
            $backupPath = "{0}.bak-{1}" -f $configPath, (Get-Date -Format 'yyyyMMdd-HHmmss')
            Copy-Item -LiteralPath $configPath -Destination $backupPath -Force
        }

        $utf8NoBom = [System.Text.UTF8Encoding]::new($false)
        [System.IO.File]::WriteAllText($configPath, $content, $utf8NoBom)
        $script:CurrentMigration.Config = $validatedConfig
        $script:CurrentStatus = Get-MigrationStatus -Migration $script:CurrentMigration
        Set-ConfigEditorDirty $false
        Update-UI
        Set-ConfigEditorStatus ("Saved {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    }
    catch {
        Set-ConfigEditorStatus "Save failed: $($_.Exception.Message)" $true
        [System.Windows.MessageBox]::Show("Config save failed:`n$($_.Exception.Message)", $script:AppName, 'OK', 'Error') | Out-Null
    }
}
function Refresh-LogList {
    $listLogFiles.Items.Clear()
    if ($null -eq $script:CurrentMigration) { return }
    $logsDir = Join-Path $script:CurrentMigration.Root 'logs'
    if (-not (Test-Path -LiteralPath $logsDir -PathType Container)) { return }
    Get-ChildItem -LiteralPath $logsDir -Filter '*.log' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 60 |
        ForEach-Object {
            $display = ('{0}  {1}' -f $_.LastWriteTime.ToString('yyyy-MM-dd HH:mm'), $_.Name)
            [void]$listLogFiles.Items.Add([pscustomobject]@{
                Display  = $display
                FullName = $_.FullName
                Name     = $_.Name
            })
        }
}

function Set-CurrentMigration {
    param($Migration)
    $script:CurrentMigration = $Migration
    $script:CurrentStatus    = Get-MigrationStatus -Migration $Migration
    Update-UI
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
$tabLogs.Add_Click({        Switch-Tab 'Logs' })
$tabConfig.Add_Click({      Switch-Tab 'Config' })

$cmbMigration.Add_SelectionChanged({
    $idx = $cmbMigration.SelectedIndex
    if ($idx -ge 0 -and $idx -lt $script:Migrations.Count) {
        Set-CurrentMigration -Migration $script:Migrations[$idx]
    }
})

$btnRefresh.Add_Click({
    Load-Migrations
})

$btnNewMigration.Add_Click({
    $tpl = Join-Path $script:ScriptRoot 'Migrations\_Template'
    $msg = "Copy the _Template folder to create a new migration:`n$tpl`n`nThen edit migration.config.psd1 in the copy, and click Refresh."
    try { Open-InExplorer (Join-Path $script:ScriptRoot 'Migrations') } catch {}
    [System.Windows.MessageBox]::Show($msg, $script:AppName, 'OK', 'Information')
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
                $dir = Split-Path $cmd -Parent
                Start-Process -FilePath 'cmd.exe' -ArgumentList "/C `"$cmd`"" -WorkingDirectory $dir
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

Load-Migrations
Close-SmartM365GuiSplash -Splash $script:Splash
[void]$script:Window.ShowDialog()


# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDC6Ux2CizuCR2C
# LvnW9QeVA83TsTL9yFNcvkAeYEzTf6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICqyb1gz0cwjXE4K11ETkMLPgzWnRyxkv9GaqsNpgsooMA0GCSqG
# SIb3DQEBAQUABIIBgGT3IdaE6lDGvgONkjJrFaJdX7UBOZvUj7nTdrsRd03RUkXw
# dLEmTiJALqj0UO1MRmuLWUkGIAqc9Mj2vgHhbjqwSJ20zDVXa25z5wk0cHJFYcaq
# fYYdJmAf70fOhuCY5336ot6ur3rcUnUHPVsGj0LwuN/g1eX7TmnfRb6GEyeX/mlB
# UN6ggeck6zLp69NsGAbopc/MPi+6a63WcrLVdUJ3Xkn7DKdeLUVVEVqDXWQ38FpQ
# +5RFqtezwcRCJarBOxtUM10KBX2YjLqM0CKbqujJV/rSZUkcQ5B6/M84ib+XkVth
# eIT2lkRy33S6AHflZALWoLAUy4LjnlRo7+6XBp46GV1N9mV5nEFfBb84RyB87eBX
# Y64nES+0uCtYl8flu2+dWCTfkUYfx/uYBMcUu7KgCbN5P6W4XLmtOV0lrnxBgRzQ
# eGAKM5dQREME0eNYtCfg8HWMeNxb7mSGwrWyU4wKtGB3vIYSbXYBwqLcRKRT/UPV
# 7DpWjLLrhbkGbXYOEaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODUw
# MTlaMC8GCSqGSIb3DQEJBDEiBCCdJ+ZXPwrI0tlL9KgGfzGfsHriTkD14KDGXavS
# 1NDHCzANBgkqhkiG9w0BAQEFAASCAgAinaSls+M+A6bmkPdPZCw5ES3HAicQtw3g
# 4tsD6zQvPSmhWxzxdfQpsHKUVEqa58rc3rCyeAH/Ip8zqO/EiZmW474RG6BwiWoA
# zZnnlTATwXRhHC1USYPvRimOpjdsWJBwnIger9rat3kEgUHvcrPWm9GwJijLgGj7
# O26B5i3Elqu2iFoEbxsjDMG5OngKAFZUEOFgZSxbKAyrYHyXJHVagzVEHtZzMolo
# CdtBf+YOLEs+0dnmCc4iOAKfm6ohDk7M64dzRAn/WipsOZknShFf8+VXDEjzU3yj
# cQ5oFccPliYkF+pRXnhaAdX3qh1ZoxHL2D1dORtB2NxQtiwsL8JVJIggvXkTDVLx
# YeuS1jlPBfyhvjNhAI6akfvVCpq0/oXPENi+zbLy4BwIQBoXVeBIHsS7i3jUE6fx
# 8u/iy9+xgyFau/1+aJ8tHnb3+zmpvcSHi18CDdPP6NHgr8H++kBdxxzybZVl3ZNr
# 9f4WjLhj74k8HYoVYv9nDFbhUttQPzydUM9coKgsG16JVKeJFvjegcuTCaCXCfGy
# xbTH4UEZLGJU/sIK6glC1tJQqWi6Ec5M1D/3aW7jHQCZz6qykqYTcVDIIKyVuQaR
# ll9XhRWmTRquUsoebLHKEyG9CEiAQi1odWrMfdW/GVceELhWsu7I6qw8dFjRW5g6
# RzD3OAzVtg==
# SIG # End signature block
