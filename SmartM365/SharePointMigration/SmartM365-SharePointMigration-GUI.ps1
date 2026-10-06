<#
.SYNOPSIS
    Smart SharePoint Migration dashboard GUI.

.DESCRIPTION
    WPF dashboard for SharePoint migration workflows. Discovers local migration
    folders, shows the last run status for each step, and launches inventory,
    comparison, and operation actions in a new PowerShell console window that stays open after completion.

.PARAMETER ValidateOnly
    Loads the GUI resources and exits without showing the window.

.PARAMETER FarmToolkitRoot
    Shared UNC root of the toolkit used by generated farm commands. Defaults to
    the directory containing this GUI when launched from the shared toolkit.

.VERSION
    1.0.64
#>

#Requires -Version 7.4

[CmdletBinding()]
param(
    [Alias('DryRun')]
    [switch]$ValidateOnly,
    [string]$FarmToolkitRoot = ''
)

$script:AppName    = 'Smart SharePoint Migration'
$script:AppVersion = '1.0.64'
$script:ScriptRoot = $PSScriptRoot
$script:FarmToolkitRoot = if ($FarmToolkitRoot) { $FarmToolkitRoot } else { $PSScriptRoot }
$script:SummaryLastGoodRows = @{}
$script:SummaryLastErrorMessages = @{}
$script:NextActivityCleanupUtc = $null
$script:LastGlobalRefreshError = ''
Microsoft.PowerShell.Utility\Write-Host ('{0} Script  : {1} v{2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $MyInvocation.MyCommand.Name, $script:AppVersion) -ForegroundColor Cyan

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase

. (Join-Path $script:ScriptRoot 'SmartM365.GuiSplash.ps1')
. (Join-Path $script:ScriptRoot 'SmartM365-SharePointMigration-NewWizard.ps1')
. (Join-Path $script:ScriptRoot 'Scripts\Launchers\Generic\SmartM365-SharePointMigration-GuiActivity.ps1')
. (Join-Path $script:ScriptRoot 'Scripts\Launchers\Generic\SmartM365-SharePointMigration-BatchGui.ps1')
. (Join-Path $script:ScriptRoot 'SmartM365-SharePointMigration-Summary.ps1')
. (Join-Path $script:ScriptRoot 'Scripts\Diagnostics\SmartM365-SharePointMigration-CrossCheck.ps1')

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
        Where-Object { Test-SmartM365InventoryCsvComplete -File $_ } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}
function Get-CsvFileItems {
    param([string]$Directory, [string]$Filter)
    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) { return @() }
    @(Get-ChildItem -LiteralPath $Directory -Filter $Filter -File -Recurse -ErrorAction SilentlyContinue |
        Where-Object { Test-SmartM365InventoryCsvComplete -File $_ } |
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
    param([string]$Directory, [string]$Pattern, [switch]$UseRunTimestamp)
    if (-not (Test-Path -LiteralPath $Directory -PathType Container)) { return $null }
    Get-ChildItem -LiteralPath $Directory -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -like $Pattern } |
        Sort-Object { $stamp = if ($UseRunTimestamp) { Get-SmartM365PortfolioTimestamp $_.Name }; if ($stamp) { $stamp } else { $_.LastWriteTime } } -Descending |
        Select-Object -First 1
}
function Get-ComparisonHtmlReport {
    param([System.IO.DirectoryInfo]$Folder)
    if ($null -eq $Folder -or -not (Test-Path -LiteralPath $Folder.FullName -PathType Container)) { return $null }
    Get-ChildItem -LiteralPath $Folder.FullName -Filter '*-summary-*.html' -File -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}
function Get-LatestComparisonResultFolder {
    param([string]$Directory, [string]$MigrationName, [ValidateSet('Files','Permissions')][string]$Kind)
    $result = Get-SmartM365LatestPortfolioComparison -Directory $Directory -MigrationName $MigrationName -Kind $Kind
    if ($result) { return Get-Item -LiteralPath (Split-Path $result.Path -Parent) -ErrorAction Stop }
    return $null
}
function Get-ComparisonBadgeText {
    param(
        [System.IO.DirectoryInfo]$Folder,
        [string]$MigrationName,
        [ValidateSet('Files','Permissions')][string]$Kind
    )
    if ($null -eq $Folder) { return 'No comparison result yet' }
    $comparison = Get-SmartM365LatestPortfolioComparison -Directory $Folder.Parent.FullName `
        -MigrationName $MigrationName -Kind $Kind -SelectedFolder $Folder
    $stamp = Get-SmartM365PortfolioTimestamp $Folder.Name
    $date = if ($comparison) { $comparison.Date } elseif ($stamp) { $stamp } else { $Folder.LastWriteTime }
    $verifiedEmptyTarget = $Kind -eq 'Files' -and $comparison -and $comparison.Target -eq 0 -and
        $comparison.Summary.PSObject.Properties['TargetEmptyVerified'] -and
        [string]$comparison.Summary.TargetEmptyVerified -eq 'True'
    $rate = if ($comparison -and $comparison.Source -gt 0 -and ($comparison.Target -gt 0 -or $verifiedEmptyTarget)) {
        '{0:N2} %' -f ([double]$comparison.Matched / [double]$comparison.Source * 100)
    } else { 'Rate unavailable' }
    return ('{0} · {1}' -f (Format-RunAgeText -Date $date), $rate)
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
    $permissionHistoryPath = if ($cfg.Output.ContainsKey('PermissionHistoryComparisons') -and
        -not [string]::IsNullOrWhiteSpace([string]$cfg.Output.PermissionHistoryComparisons)) {
        [string]$cfg.Output.PermissionHistoryComparisons
    } else { 'comparisons\permission-scan-history' }
    $permissionHistoryDir = Join-Path $root $permissionHistoryPath

    $fileComparisonFolder = Get-LatestComparisonResultFolder $fileCmpDir $name Files
    $permComparisonFolder = Get-LatestComparisonResultFolder $permCmpDir $name Permissions

    [pscustomobject]@{
        SourceFileCsv         = Get-LatestCsvFile    $srcFileDir  ("{0}-FileInventory-$name-*.csv" -f (Get-MigrationEndpointType $cfg 'Source'))
        SourceFileCsvItems    = @(Get-CsvFileItems   $srcFileDir  ("{0}-FileInventory-$name-*.csv" -f (Get-MigrationEndpointType $cfg 'Source')))
        TargetFileCsv         = Get-LatestCsvFile    $tgtFileDir  ("{0}-FileInventory-$name-*.csv" -f (Get-MigrationEndpointType $cfg 'Target'))
        TargetFileCsvItems    = @(Get-CsvFileItems   $tgtFileDir  ("{0}-FileInventory-$name-*.csv" -f (Get-MigrationEndpointType $cfg 'Target')))
        FileComparisonFolder  = $fileComparisonFolder
        FileComparisonAttemptFolder = Get-LatestSubfolder $fileCmpDir "$name-*" -UseRunTimestamp
        FileComparisonReport  = Get-ComparisonHtmlReport $fileComparisonFolder
        HistoryFolder         = Get-LatestSubfolder  $histDir     '*-Changes-*'
        SourcePermCsv         = Get-LatestCsvFile    $srcPermDir  ("{0}-PermissionInventory-$name-*.csv" -f (Get-MigrationEndpointType $cfg 'Source'))
        SourcePermCsvItems    = @(Get-CsvFileItems   $srcPermDir  ("{0}-PermissionInventory-$name-*.csv" -f (Get-MigrationEndpointType $cfg 'Source')))
        TargetPermCsv         = Get-LatestCsvFile    $tgtPermDir  ("{0}-PermissionInventory-$name-*.csv" -f (Get-MigrationEndpointType $cfg 'Target'))
        TargetPermCsvItems    = @(Get-CsvFileItems   $tgtPermDir  ("{0}-PermissionInventory-$name-*.csv" -f (Get-MigrationEndpointType $cfg 'Target')))
        PermComparisonFolder  = $permComparisonFolder
        PermComparisonAttemptFolder = Get-LatestSubfolder $permCmpDir "$name-*" -UseRunTimestamp
        PermComparisonReport  = Get-ComparisonHtmlReport $permComparisonFolder
        PermissionHistoryFolder = Get-LatestSubfolder $permissionHistoryDir '*-PermissionChanges-*'
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

function Format-RunAgeText {
    param([datetime]$Date, [datetime]$Now = (Get-Date))
    $age = $Now - $Date
    $dayOffset = ($Now.Date - $Date.Date).Days
    if ($dayOffset -eq 0 -and $age.TotalMinutes -ge 0 -and $age.TotalMinutes -lt 90) {
        return ('{0:HH:mm} today ({1:n0} min ago)' -f $Date, [int]$age.TotalMinutes)
    } elseif ($dayOffset -eq 0) {
        return ('{0:HH:mm} today' -f $Date)
    } elseif ($dayOffset -eq 1) {
        return ('{0:HH:mm} yesterday' -f $Date)
    }
    return ('{0:yyyy-MM-dd HH:mm}' -f $Date)
}
function Format-ItemAge {
    param([System.IO.FileSystemInfo]$Item)
    if ($null -eq $Item) { return [pscustomobject]@{ Text = 'No run yet'; HasRun = $false } }
    return [pscustomobject]@{ Text = (Format-RunAgeText -Date $Item.LastWriteTime); HasRun = $true }
}

function Open-InExplorer {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return }
    try { Invoke-Item -LiteralPath $Path } catch { Start-Process explorer.exe -ArgumentList ('"' + $Path + '"') }
}

function New-FarmDiagnosticsWindow {
    param([System.Windows.Window]$Owner, [string]$MigrationName, [System.Windows.UIElement]$Content)
    $window = [System.Windows.Window]::new()
    $window.Title = 'Source farm diagnostics · ' + $MigrationName
    $window.Width = 1000; $window.Height = 500
    $window.MinWidth = 760; $window.MinHeight = 360
    $window.WindowStartupLocation = 'CenterOwner'
    $window.ShowInTaskbar = $false
    $window.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#F5F8FB')
    $window.FontFamily = [System.Windows.Media.FontFamily]::new('Segoe UI')
    if ($Owner) {
        if ($Owner.IsVisible) { $window.Owner = $Owner }
        $window.Icon = $Owner.Icon; $window.Resources = $Owner.Resources
    }
    $panel = [System.Windows.Controls.DockPanel]::new()
    $panel.Margin = [System.Windows.Thickness]::new(18)
    $panel.Background = $window.Background
    $heading = [System.Windows.Controls.TextBlock]::new()
    $heading.Text = 'Selected migration: ' + $MigrationName
    $heading.FontSize = 16; $heading.FontWeight = [System.Windows.FontWeights]::SemiBold
    $heading.Margin = [System.Windows.Thickness]::new(0,0,0,14)
    [System.Windows.Controls.DockPanel]::SetDock($heading, 'Top')
    [void]$panel.Children.Add($heading)
    $scroll = [System.Windows.Controls.ScrollViewer]::new()
    $Content.VerticalAlignment = 'Top'
    $scroll.VerticalScrollBarVisibility = 'Auto'; $scroll.Content = $Content
    [void]$panel.Children.Add($scroll)
    $window.Content = $panel
    return $window
}

# ---------------------------------------------------------------------------
# XAML
# ---------------------------------------------------------------------------

[xml]$xaml = @'
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Smart SharePoint Migration"
    Width="1480" Height="800"
    MinWidth="1280" MinHeight="580"
    WindowStartupLocation="CenterScreen"
    WindowState="Maximized"
    UseLayoutRounding="True"
    SnapsToDevicePixels="True"
    Background="#F5F8FB">

  <Window.Resources>
    <Style x:Key="PortfolioBadge" TargetType="Border">
      <Setter Property="Background" Value="#F0F4F8"/>
      <Setter Property="TextElement.Foreground" Value="#64748B"/>
      <Setter Property="CornerRadius" Value="5"/>
      <Setter Property="Padding" Value="4,5"/>
      <Setter Property="Margin" Value="3"/>
      <Setter Property="VerticalAlignment" Value="Center"/>
    </Style>
    <Style x:Key="PortfolioStatusBadge" TargetType="Border" BasedOn="{StaticResource PortfolioBadge}">
      <Style.Triggers>
        <DataTrigger Binding="{Binding Status}" Value="Up to date">
          <Setter Property="Background" Value="#ECFAF4"/><Setter Property="TextElement.Foreground" Value="#087F5B"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Status}" Value="Compare needed">
          <Setter Property="Background" Value="#FFF8E8"/><Setter Property="TextElement.Foreground" Value="#9A6700"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Status}" Value="Review needed">
          <Setter Property="Background" Value="#FFF8E8"/><Setter Property="TextElement.Foreground" Value="#9A6700"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Status}" Value="Scan needed">
          <Setter Property="Background" Value="#FFF0F2"/><Setter Property="TextElement.Foreground" Value="#B42335"/>
        </DataTrigger>
        <DataTrigger Binding="{Binding Status}" Value="Refresh error">
          <Setter Property="Background" Value="#FFF0F2"/><Setter Property="TextElement.Foreground" Value="#B42335"/>
        </DataTrigger>
      </Style.Triggers>
    </Style>
    <DataTemplate x:Key="PortfolioRateTemplate">
      <Border Background="{Binding Background}" CornerRadius="5" Margin="3,2" Padding="6,2">
        <StackPanel VerticalAlignment="Center">
          <TextBlock Text="{Binding Percent}" FontSize="20" FontWeight="Bold" Foreground="{Binding Color}" HorizontalAlignment="Center"/>
          <TextBlock Text="{Binding Caption}" FontSize="10" Foreground="#5F6B7A" HorizontalAlignment="Center"/>
          <ProgressBar Value="{Binding Progress}" Minimum="0" Maximum="100" Height="2" Margin="0,2,0,0" Foreground="{Binding Color}" Background="#E2E8F0" BorderThickness="0" Visibility="{Binding ProgressVisibility}"/>
          <TextBlock Text="{Binding Notice}" FontSize="10" FontWeight="SemiBold" Foreground="{Binding Color}" HorizontalAlignment="Center" Margin="0,1,0,0" Visibility="{Binding NoticeVisibility}"/>
        </StackPanel>
      </Border>
    </DataTemplate>
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

    <Style x:Key="DiagMetricTile" TargetType="Border">
      <Setter Property="Background" Value="#F5F8FB"/>
      <Setter Property="BorderBrush" Value="#DDE7F0"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="CornerRadius" Value="6"/>
      <Setter Property="Padding" Value="9,7"/>
    </Style>
    <Style x:Key="DiagMetricLabel" TargetType="TextBlock">
      <Setter Property="FontSize" Value="10"/>
      <Setter Property="Foreground" Value="#5F6B7A"/>
      <Setter Property="TextWrapping" Value="Wrap"/>
    </Style>
    <Style x:Key="DiagMetricValue" TargetType="TextBlock">
      <Setter Property="FontSize" Value="19"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Foreground" Value="#17324D"/>
    </Style>
  </Window.Resources>

  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
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
          <Image x:Name="imgLogo" Width="48" Height="38" Margin="0,0,12,0"
                 VerticalAlignment="Center" Stretch="Uniform"/>
          <StackPanel VerticalAlignment="Center">
            <TextBlock Text="Smart SharePoint Migration" FontSize="15" FontWeight="Medium" Foreground="#1F2937"/>
            <TextBlock Text="SharePoint migration dashboard" FontSize="12" Foreground="#5F6B7A" Margin="0,1,0,0"/>
          </StackPanel>
        </StackPanel>
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
        <ToggleButton x:Name="tabSummary"     Content="Overview"    Style="{StaticResource Tab}" IsChecked="True"/>
        <ToggleButton x:Name="tabFiles"       Content="Files &amp; Permissions" Style="{StaticResource Tab}"/>
        <ToggleButton x:Name="tabDiagnostics" Content="Migration Diagnostics" Style="{StaticResource Tab}"/>
        <ToggleButton x:Name="tabBatches" Content="Batch runs" Style="{StaticResource Tab}"/>
        <ToggleButton x:Name="tabOperations"  Content="Operations"  Style="{StaticResource Tab}"/>
        <ToggleButton x:Name="tabLogs"        Content="Logs"        Style="{StaticResource Tab}"/>
        <ToggleButton x:Name="tabConfig"      Content="Config"      Style="{StaticResource Tab}"/>
      </StackPanel>
    </Border>

    <!-- Tab content -->
    <ScrollViewer Grid.Row="3" VerticalScrollBarVisibility="Auto">
      <Grid>

        <!-- PORTFOLIO SUMMARY -->
        <StackPanel x:Name="panelSummary" Margin="18,14" Visibility="Visible">
          <TextBlock Text="MIGRATION OVERVIEW" Style="{StaticResource SectionLabel}"/>
          <TextBlock Text="Files: matches / source keys. Permissions: matches / source permission keys. Global: (files % + permissions %) / 2. Select a migration to open Files &amp; Permissions."
                     FontSize="12" Foreground="#5F6B7A" Margin="0,0,0,10" TextWrapping="Wrap"/>
          <Grid Margin="0,0,0,12">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="260"/>
            </Grid.ColumnDefinitions>
            <Border Grid.Column="0" Style="{StaticResource StepCard}" Margin="0,0,5,0" Background="#EFF7FF">
              <StackPanel>
                <TextBlock Text="MIGRATIONS" Style="{StaticResource SectionLabel}"/>
                <TextBlock x:Name="lblPortfolioMigrations" Text="—" FontSize="22" FontWeight="SemiBold" Foreground="#17324D"/>
                <TextBlock Text="Source: displayed migrations" FontSize="10" Foreground="#5F6B7A"/>
              </StackPanel>
            </Border>
            <Border Grid.Column="1" Style="{StaticResource StepCard}" Margin="5,0,5,0" Background="#EFF7FF">
              <StackPanel>
                <TextBlock Text="FILES COMPARED" Style="{StaticResource SectionLabel}"/>
                <TextBlock x:Name="lblPortfolioFilesCompared" Text="—" FontSize="22" FontWeight="SemiBold" Foreground="#17324D"/>
                <TextBlock Text="Source: file comparisons" FontSize="10" Foreground="#5F6B7A"/>
              </StackPanel>
            </Border>
            <Border Grid.Column="2" Style="{StaticResource StepCard}" Margin="5,0,5,0" Background="#F0FBF8">
              <StackPanel>
                <TextBlock Text="PERMISSIONS COMPARED" Style="{StaticResource SectionLabel}"/>
                <TextBlock x:Name="lblPortfolioPermissionsCompared" Text="—" FontSize="22" FontWeight="SemiBold" Foreground="#17324D"/>
                <TextBlock Text="Source: permission comparisons" FontSize="10" Foreground="#5F6B7A"/>
              </StackPanel>
            </Border>
            <Border Grid.Column="3" Style="{StaticResource StepCard}" Margin="5,0,5,0" Background="#FFF8EB">
              <StackPanel>
                <TextBlock Text="ACTIONS REQUIRED" Style="{StaticResource SectionLabel}"/>
                <TextBlock x:Name="lblPortfolioActions" Text="—" FontSize="22" FontWeight="SemiBold" Foreground="#8B5E00"/>
                <TextBlock Text="Scan, compare or review" FontSize="10" Foreground="#5F6B7A"/>
              </StackPanel>
            </Border>
            <Border Grid.Column="4" Style="{StaticResource StepCard}" Margin="5,0,0,0" Background="#FFF2F2">
              <StackPanel>
                <TextBlock Text="REFRESH ERRORS" Style="{StaticResource SectionLabel}"/>
                <TextBlock x:Name="lblPortfolioRefreshErrors" Text="—" FontSize="22" FontWeight="SemiBold" Foreground="#A4262C"/>
                <TextBlock Text="Source: overview refresh" FontSize="10" Foreground="#5F6B7A"/>
              </StackPanel>
            </Border>
            <Border Grid.Column="5" Style="{StaticResource StepCard}" Margin="10,0,0,0">
              <StackPanel>
                <TextBlock Text="GLOBAL REPORTS" Style="{StaticResource SectionLabel}"/>
                <Button x:Name="btnOverviewGlobalFileReport" Content="Global file comparison report"
                        ToolTip="Generate HTML, Excel and CSV reports for the latest file comparisons"
                        Style="{StaticResource Btn}" FontSize="11" Padding="6,0" HorizontalAlignment="Stretch" Height="28" Margin="0,2,0,5"/>
                <Button x:Name="btnOverviewGlobalPermissionsReport" Content="Global permissions comparison report"
                        ToolTip="Generate HTML, Excel and CSV reports for the latest permissions comparisons"
                        Style="{StaticResource Btn}" FontSize="11" Padding="6,0" HorizontalAlignment="Stretch" Height="28"/>
              </StackPanel>
            </Border>
          </Grid>
          <Border Style="{StaticResource StepCard}" Padding="0">
            <DataGrid x:Name="gridSummary" AutoGenerateColumns="False" IsReadOnly="True"
                      CanUserAddRows="False" CanUserDeleteRows="False" CanUserSortColumns="True"
                      SelectionMode="Single" SelectionUnit="FullRow" HeadersVisibility="Column"
                      GridLinesVisibility="None" RowHeight="68" ColumnHeaderHeight="36"
                      AlternatingRowBackground="#F5F8FB" Background="White"
                      BorderThickness="0" HorizontalScrollBarVisibility="Auto"
                      VerticalScrollBarVisibility="Disabled" EnableRowVirtualization="True"
                      FrozenColumnCount="6">
              <DataGrid.RowStyle>
                <Style TargetType="DataGridRow">
                  <Setter Property="ToolTip" Value="{Binding StatusTooltip}"/>
                </Style>
              </DataGrid.RowStyle>
              <DataGrid.Columns>
                <DataGridTextColumn Header="Migration" Binding="{Binding Migration}" Width="110">
                  <DataGridTextColumn.ElementStyle>
                    <Style TargetType="TextBlock"><Setter Property="VerticalAlignment" Value="Center"/></Style>
                  </DataGridTextColumn.ElementStyle>
                </DataGridTextColumn>
                <DataGridTemplateColumn Header="Gap (days)" Width="80" SortMemberPath="ScanGapDays">
                  <DataGridTemplateColumn.CellTemplate><DataTemplate>
                    <Border Style="{StaticResource PortfolioBadge}" Background="{Binding ScanGapVisual.Background}" ToolTip="{Binding ScanGapTooltip}">
                      <TextBlock Text="{Binding ScanGapText}" Foreground="{Binding ScanGapVisual.Color}" FontSize="12" FontWeight="SemiBold" HorizontalAlignment="Center"/>
                    </Border>
                  </DataTemplate></DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTemplateColumn Header="Status" Width="115" SortMemberPath="Status">
                  <DataGridTemplateColumn.CellTemplate><DataTemplate>
                    <Border Style="{StaticResource PortfolioStatusBadge}" ToolTip="{Binding StatusTooltip}">
                      <TextBlock Text="{Binding Status}" FontSize="11" FontWeight="SemiBold" TextWrapping="Wrap" TextAlignment="Center" HorizontalAlignment="Center"/>
                    </Border>
                  </DataTemplate></DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTemplateColumn Header="▣ Files comparison" Width="150" SortMemberPath="ComparisonRate">
                  <DataGridTemplateColumn.CellTemplate><DataTemplate>
                    <ContentControl Content="{Binding ComparisonVisual}" ContentTemplate="{StaticResource PortfolioRateTemplate}" ToolTip="{Binding ComparisonTooltip}" HorizontalContentAlignment="Stretch" VerticalContentAlignment="Stretch"/>
                  </DataTemplate></DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTemplateColumn Header="◆ Permissions comparison" Width="180" SortMemberPath="PermissionComparisonRate">
                  <DataGridTemplateColumn.CellTemplate><DataTemplate>
                    <ContentControl Content="{Binding PermissionComparisonVisual}" ContentTemplate="{StaticResource PortfolioRateTemplate}" ToolTip="{Binding PermissionComparisonTooltip}" HorizontalContentAlignment="Stretch" VerticalContentAlignment="Stretch"/>
                  </DataTemplate></DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTemplateColumn Header="◉ Global comparison" Width="150" SortMemberPath="GlobalComparisonRate">
                  <DataGridTemplateColumn.CellTemplate><DataTemplate>
                    <ContentControl Content="{Binding GlobalComparisonVisual}" ContentTemplate="{StaticResource PortfolioRateTemplate}" ToolTip="{Binding GlobalComparisonTooltip}" HorizontalContentAlignment="Stretch" VerticalContentAlignment="Stretch"/>
                  </DataTemplate></DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTextColumn Header="Source" Binding="{Binding Source}" Width="*" MinWidth="140">
                  <DataGridTextColumn.ElementStyle>
                    <Style TargetType="TextBlock">
                      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
                      <Setter Property="ToolTip" Value="{Binding SourceTooltip}"/>
                    </Style>
                  </DataGridTextColumn.ElementStyle>
                </DataGridTextColumn>
                <DataGridTextColumn Header="Source scans" Binding="{Binding SourceScansDisplay}" Width="135" SortMemberPath="SourceScansSortDate">
                  <DataGridTextColumn.ElementStyle>
                    <Style TargetType="TextBlock"><Setter Property="FontSize" Value="10.5"/><Setter Property="VerticalAlignment" Value="Center"/><Setter Property="ToolTip" Value="{Binding SourceScansTooltip}"/></Style>
                  </DataGridTextColumn.ElementStyle>
                </DataGridTextColumn>
                <DataGridTextColumn Header="Source inventory" Binding="{Binding SourceInventoryDisplay}" Width="160">
                  <DataGridTextColumn.ElementStyle>
                    <Style TargetType="TextBlock"><Setter Property="FontSize" Value="10.5"/><Setter Property="VerticalAlignment" Value="Center"/><Setter Property="TextWrapping" Value="Wrap"/><Setter Property="ToolTip" Value="{Binding SourceInventoryTooltip}"/></Style>
                  </DataGridTextColumn.ElementStyle>
                </DataGridTextColumn>
                <DataGridTextColumn Header="Destination" Binding="{Binding Destination}" Width="1.3*" MinWidth="160">
                  <DataGridTextColumn.ElementStyle>
                    <Style TargetType="TextBlock">
                      <Setter Property="TextTrimming" Value="CharacterEllipsis"/>
                      <Setter Property="ToolTip" Value="{Binding DestinationTooltip}"/>
                    </Style>
                  </DataGridTextColumn.ElementStyle>
                </DataGridTextColumn>
                <DataGridTextColumn Header="Target scans" Binding="{Binding TargetScansDisplay}" Width="135" SortMemberPath="TargetScansSortDate">
                  <DataGridTextColumn.ElementStyle>
                    <Style TargetType="TextBlock"><Setter Property="FontSize" Value="10.5"/><Setter Property="VerticalAlignment" Value="Center"/><Setter Property="ToolTip" Value="{Binding TargetScansTooltip}"/></Style>
                  </DataGridTextColumn.ElementStyle>
                </DataGridTextColumn>
                <DataGridTextColumn Header="Target inventory" Binding="{Binding TargetInventoryDisplay}" Width="160">
                  <DataGridTextColumn.ElementStyle>
                    <Style TargetType="TextBlock"><Setter Property="FontSize" Value="10.5"/><Setter Property="VerticalAlignment" Value="Center"/><Setter Property="TextWrapping" Value="Wrap"/><Setter Property="ToolTip" Value="{Binding TargetInventoryTooltip}"/></Style>
                  </DataGridTextColumn.ElementStyle>
                </DataGridTextColumn>
              </DataGrid.Columns>
            </DataGrid>
          </Border>
          <TextBlock x:Name="lblSummaryStatus" FontSize="11" Foreground="#5F6B7A" Margin="0,3,0,0"/>
        </StackPanel>

        <!-- FILES AND PERMISSIONS -->
        <Grid x:Name="panelWorkflows" Margin="18,14" Visibility="Collapsed">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="1"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <StackPanel x:Name="panelFiles" Grid.Column="0" Margin="0,0,14,0">
          <TextBlock Text="FILES" FontSize="16" FontWeight="SemiBold" Foreground="#1F2937" Margin="0,0,0,10"/>

          <TextBlock Text="INVENTORY" Style="{StaticResource SectionLabel}"/>

          <Border Style="{StaticResource StepCard}">
            <Grid>
              <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
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
                  <ComboBox x:Name="cmbScanSrcFile" Width="290" Height="24" FontSize="11" DisplayMemberPath="Display" VerticalContentAlignment="Center"/>
                </StackPanel>
              </StackPanel>
              <StackPanel Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
                <Button x:Name="btnOpenSourceSite" Content="Open source site" Style="{StaticResource Btn}" Width="145" Margin="0,0,6,0" IsEnabled="False"/>
                <Button x:Name="btnOpenScanSrc" Content="Open folder" Style="{StaticResource BtnGhost}"
                        Width="100" Margin="0,0,6,0" Visibility="Collapsed"/>
                <Button x:Name="btnRunScanSrc"  Content="Run"  Style="{StaticResource Btn}" Width="55"/>
              </StackPanel>
            </Grid>
          </Border>

          <Border Style="{StaticResource StepCard}">
            <Grid>
              <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
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
                  <ComboBox x:Name="cmbScanTgtFile" Width="290" Height="24" FontSize="11" DisplayMemberPath="Display" VerticalContentAlignment="Center"/>
                </StackPanel>
              </StackPanel>
              <StackPanel Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
                <Button x:Name="btnOpenTargetSite" Content="Open destination site" Style="{StaticResource Btn}" Width="170" Margin="0,0,6,0" IsEnabled="False"/>
                <Button x:Name="btnOpenScanTgt" Content="Open folder" Style="{StaticResource BtnGhost}"
                        Width="100" Margin="0,0,6,0" Visibility="Collapsed"/>
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
                <Grid Margin="0,3,0,0">
                  <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <Border Grid.Column="0" x:Name="badgeCmpFiles" CornerRadius="10" Padding="6,2" Margin="0,0,8,0" Background="#E6F4FF">
                    <TextBlock x:Name="lblCmpFilesAge" Text="No comparison result yet" FontSize="11" Foreground="#005A9E"/>
                  </Border>
                  <TextBlock Grid.Column="1" x:Name="lblCmpFilesDir" Text="" FontSize="11" Foreground="#5F6B7A" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
                </Grid>
                <TextBlock x:Name="lblCmpFilesAvailability" FontSize="11" Foreground="#8A5E00" Margin="0,4,8,0" TextWrapping="Wrap" Visibility="Collapsed"/>
              </StackPanel>
              <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                <Button x:Name="btnOpenCmpFiles" Content="Open" Style="{StaticResource BtnGhost}"
                        Width="58" Margin="0,0,6,0" Visibility="Collapsed"/>
                <Button x:Name="btnReportCmpFiles" Content="HTML report" Style="{StaticResource BtnGhost}"
                        Width="88" Margin="0,0,6,0" Visibility="Collapsed"/>
                <Button x:Name="btnRunCmpFiles" Content="Run" Style="{StaticResource Btn}" Width="55" IsEnabled="False" ToolTipService.ShowOnDisabled="True"/>
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
                </StackPanel>
                <StackPanel Orientation="Horizontal" Margin="0,3,0,0">
                  <TextBlock Text="Previous" FontSize="11" Foreground="#5F6B7A" VerticalAlignment="Center" Margin="0,0,5,0"/>
                  <ComboBox x:Name="cmbHistoryOldFile" Width="175" Height="24" FontSize="11" DisplayMemberPath="Display" VerticalContentAlignment="Center" Margin="0,0,8,0"/>
                  <TextBlock Text="Current" FontSize="11" Foreground="#5F6B7A" VerticalAlignment="Center" Margin="0,0,5,0"/>
                  <ComboBox x:Name="cmbHistoryNewFile" Width="175" Height="24" FontSize="11" DisplayMemberPath="Display" VerticalContentAlignment="Center"/>
                </StackPanel>
              </StackPanel>
              <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                <Button x:Name="btnOpenHistory" Content="Open" Style="{StaticResource BtnGhost}"
                        Width="58" Margin="0,0,6,0" Visibility="Collapsed"/>
                <Button x:Name="btnRunHistory" Content="Compare" Style="{StaticResource Btn}" Width="70" IsEnabled="False" ToolTipService.ShowOnDisabled="True"/>
              </StackPanel>
            </Grid>
          </Border>

          <Button x:Name="btnGlobalFileReport" Content="Global file comparison report"
                  ToolTip="Generate HTML, Excel and CSV reports for the latest file comparisons"
                  Style="{StaticResource Btn}" HorizontalAlignment="Right" Padding="12,6" Margin="0,10,0,0"/>

          </StackPanel>
          <Border Grid.Column="1" Background="#CBD8E6" Margin="0,0,0,0"/>

          <!-- PERMISSIONS -->
          <StackPanel x:Name="panelPermissions" Grid.Column="2" Margin="14,0,0,0">
          <TextBlock Text="PERMISSIONS" FontSize="16" FontWeight="SemiBold" Foreground="#1F2937" Margin="0,0,0,10"/>

          <TextBlock Text="INVENTORY" Style="{StaticResource SectionLabel}"/>

          <Border Style="{StaticResource StepCard}">
            <Grid>
              <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
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
                  <ComboBox x:Name="cmbScanSrcPermFile" Width="290" Height="24" FontSize="11" DisplayMemberPath="Display" VerticalContentAlignment="Center"/>
                </StackPanel>
              </StackPanel>
              <StackPanel Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
                <Button x:Name="btnOpenSourceSitePerm" Content="Open source site" Style="{StaticResource Btn}" Width="145" Margin="0,0,6,0" IsEnabled="False"/>
                <Button x:Name="btnOpenScanSrcPerm" Content="Open folder" Style="{StaticResource BtnGhost}"
                        Width="100" Margin="0,0,6,0" Visibility="Collapsed"/>
                <Button x:Name="btnRunScanSrcPerm"  Content="Run"  Style="{StaticResource Btn}" Width="55"/>
              </StackPanel>
            </Grid>
          </Border>

          <Border Style="{StaticResource StepCard}">
            <Grid>
              <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
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
                  <ComboBox x:Name="cmbScanTgtPermFile" Width="290" Height="24" FontSize="11" DisplayMemberPath="Display" VerticalContentAlignment="Center"/>
                </StackPanel>
              </StackPanel>
              <StackPanel Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="2" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,8,0,0">
                <Button x:Name="btnOpenTargetSitePerm" Content="Open destination site" Style="{StaticResource Btn}" Width="170" Margin="0,0,6,0" IsEnabled="False"/>
                <Button x:Name="btnOpenScanTgtPerm" Content="Open folder" Style="{StaticResource BtnGhost}"
                        Width="100" Margin="0,0,6,0" Visibility="Collapsed"/>
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
                <Grid Margin="0,3,0,0">
                  <Grid.ColumnDefinitions><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <Border Grid.Column="0" x:Name="badgeCmpPerms" CornerRadius="10" Padding="6,2" Margin="0,0,8,0" Background="#E6F4FF">
                    <TextBlock x:Name="lblCmpPermsAge" Text="No comparison result yet" FontSize="11" Foreground="#005A9E"/>
                  </Border>
                  <TextBlock Grid.Column="1" x:Name="lblCmpPermsDir" Text="" FontSize="11" Foreground="#5F6B7A" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
                </Grid>
                <TextBlock x:Name="lblCmpPermsAvailability" FontSize="11" Foreground="#8A5E00" Margin="0,4,8,0" TextWrapping="Wrap" Visibility="Collapsed"/>
              </StackPanel>
              <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                <Button x:Name="btnOpenCmpPerms" Content="Open" Style="{StaticResource BtnGhost}"
                        Width="58" Margin="0,0,6,0" Visibility="Collapsed"/>
                <Button x:Name="btnReportCmpPerms" Content="HTML report" Style="{StaticResource BtnGhost}"
                        Width="88" Margin="0,0,6,0" Visibility="Collapsed"/>
                <Button x:Name="btnRunCmpPerms" Content="Run" Style="{StaticResource Btn}" Width="55" IsEnabled="False" ToolTipService.ShowOnDisabled="True"/>
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
                <TextBlock Text="Compare permission scan history" FontSize="13" FontWeight="Medium" Foreground="#1F2937"/>
                <StackPanel Orientation="Horizontal" Margin="0,3,0,0">
                  <Border x:Name="badgePermHistory" CornerRadius="10" Padding="6,2" Margin="0,0,8,0" Background="#F0F4F8">
                    <TextBlock x:Name="lblPermHistoryAge" Text="No run yet" FontSize="11" Foreground="#5F6B7A"/>
                  </Border>
                  <ComboBox x:Name="cmbPermHistorySide" Width="80" Height="24" FontSize="11" VerticalContentAlignment="Center" Margin="0,0,8,0">
                    <ComboBoxItem Content="Source" IsSelected="True"/>
                    <ComboBoxItem Content="Target"/>
                  </ComboBox>
                </StackPanel>
                <StackPanel Orientation="Horizontal" Margin="0,3,0,0">
                  <TextBlock Text="Previous" FontSize="11" Foreground="#5F6B7A" VerticalAlignment="Center" Margin="0,0,5,0"/>
                  <ComboBox x:Name="cmbPermHistoryOldFile" Width="175" Height="24" FontSize="11" DisplayMemberPath="Display" VerticalContentAlignment="Center" Margin="0,0,8,0"/>
                  <TextBlock Text="Current" FontSize="11" Foreground="#5F6B7A" VerticalAlignment="Center" Margin="0,0,5,0"/>
                  <ComboBox x:Name="cmbPermHistoryNewFile" Width="175" Height="24" FontSize="11" DisplayMemberPath="Display" VerticalContentAlignment="Center"/>
                </StackPanel>
              </StackPanel>
              <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                <Button x:Name="btnOpenPermHistory" Content="Open" Style="{StaticResource BtnGhost}"
                        Width="58" Margin="0,0,6,0" Visibility="Collapsed"/>
                <Button x:Name="btnRunPermHistory" Content="Compare" Style="{StaticResource Btn}" Width="70" IsEnabled="False" ToolTipService.ShowOnDisabled="True"/>
              </StackPanel>
            </Grid>
          </Border>

          <Button x:Name="btnGlobalPermissionsReport" Content="Global permissions comparison report"
                  ToolTip="Generate HTML, Excel and CSV reports for the latest permissions comparisons"
                  Style="{StaticResource Btn}" HorizontalAlignment="Right" Padding="12,6" Margin="0,10,0,0"/>

          </StackPanel>
        </Grid>

        <!-- OPERATIONS -->
        <ContentControl x:Name="panelBatches" Margin="18,14" Visibility="Collapsed"/>

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
        <Grid x:Name="panelDiagnostics" Margin="18,14" Visibility="Collapsed">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*" MinWidth="560"/>
            <ColumnDefinition Width="1"/>
            <ColumnDefinition Width="*" MinWidth="560"/>
          </Grid.ColumnDefinitions>
          <StackPanel Grid.Row="0" Grid.ColumnSpan="3">
            <TextBlock x:Name="lblDiagMigration" Text="Selected migration: none" FontSize="16" FontWeight="SemiBold" Foreground="#17324D" Margin="0,0,0,10"/>
            <StackPanel x:Name="panelDiagAnalysisLoading" Visibility="Collapsed" Margin="0,0,0,14">
              <TextBlock Text="{Binding Text, ElementName=lblDiagProgress}" FontSize="15" FontWeight="SemiBold" Foreground="#0078D4" TextWrapping="Wrap"/>
              <ProgressBar Height="5" Margin="0,7,0,0" IsIndeterminate="True" Foreground="#0078D4"/>
            </StackPanel>
            <StackPanel x:Name="panelCrossCheckLoading" Visibility="Collapsed" Margin="0,0,0,14">
              <TextBlock x:Name="lblCrossCheckLoading" Text="Loading comparison reports…" FontSize="15" FontWeight="SemiBold" Foreground="#0078D4"/>
              <ProgressBar Height="5" Margin="0,7,0,0" IsIndeterminate="True" Foreground="#0078D4"/>
            </StackPanel>
          </StackPanel>
          <StackPanel Grid.Row="2" Grid.Column="0" Margin="0,0,14,14">
            <TextBlock Text="SHAREGATE REPORT ANALYSIS" Style="{StaticResource SectionLabel}"/>
            <Border Style="{StaticResource StepCard}">
              <StackPanel>
                <TextBlock x:Name="lblDiagScope" Text="Analysis only: no ShareGate connection or migration action." Foreground="#5F6B7A" FontSize="12" Margin="0,0,0,7"/>
                <TextBlock x:Name="lblDiagInputPath" Text="SharePointMigration\Migrations\…\ShareGate\MigrationReport" FontFamily="Consolas" FontSize="11" Foreground="#17324D" TextWrapping="Wrap"/>
                <TextBlock x:Name="lblDiagLatestReport" Text="Place the latest ShareGate migration report here (CSV or XLSX)." TextWrapping="Wrap" FontSize="12" Foreground="#8B5E00" Margin="0,7,0,0"/>
                <StackPanel Orientation="Horizontal" Margin="0,8,0,0">
                  <Button x:Name="btnDiagOpenFolder" Content="Open report folder" Style="{StaticResource BtnGhost}" Width="130"/>
                  <Button x:Name="btnDiagRefresh" Content="Refresh reports" Style="{StaticResource BtnGhost}" Width="110" Margin="6,0,0,0"/>
                  <Button x:Name="btnDiagAnalyze" Content="Analyze latest report" Style="{StaticResource Btn}" Width="145" Margin="12,0,0,0" IsEnabled="False"/>
                  <Button x:Name="btnDiagOpenReport" Content="Open analysis HTML" Style="{StaticResource BtnGhost}" Width="130" Margin="6,0,0,0" IsEnabled="False" ToolTip="Open the HTML generated for the latest ShareGate report"/>
                </StackPanel>
                <TextBlock x:Name="lblDiagProgress" Text="Select a migration report folder or a CSV/XLSX file." Foreground="#5F6B7A" FontSize="11" Margin="0,8,0,0" TextWrapping="Wrap"/>
              </StackPanel>
            </Border>
            <Border Style="{StaticResource StepCard}" Margin="0,3,0,0">
              <StackPanel>
                <TextBlock Text="SHAREGATE REPORT &amp; ANALYSIS STATUS" Style="{StaticResource SectionLabel}"/>
                <TextBlock Text="For the latest ShareGate migration report. File and permission comparison states are shown in Cross-check." TextWrapping="Wrap" Foreground="#5F6B7A" Margin="0,0,0,8"/>
                <Grid>
                  <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                  <Border Grid.Column="0" Style="{StaticResource DiagMetricTile}" Background="#EFF7FF" BorderBrush="#C9E3FA" Margin="0,0,4,0">
                    <StackPanel>
                      <TextBlock Text="SHAREGATE REPORT" Style="{StaticResource DiagMetricLabel}"/>
                      <TextBlock x:Name="lblDiagReportState" Text="Waiting" FontSize="15" FontWeight="SemiBold" Foreground="#17324D"/>
                      <TextBlock x:Name="lblDiagReportEvidence" Text="Select a migration" FontSize="10" Foreground="#5F6B7A" TextWrapping="Wrap"/>
                    </StackPanel>
                  </Border>
                  <Border Grid.Column="1" Style="{StaticResource DiagMetricTile}" Background="#F0FBF8" BorderBrush="#C9EADF" Margin="4,0,4,0">
                    <StackPanel>
                      <TextBlock Text="ANALYSIS" Style="{StaticResource DiagMetricLabel}"/>
                      <TextBlock x:Name="lblDiagAnalysisState" Text="Waiting" FontSize="15" FontWeight="SemiBold" Foreground="#17324D"/>
                      <TextBlock x:Name="lblDiagAnalysisEvidence" Text="No result loaded" FontSize="10" Foreground="#5F6B7A" TextWrapping="Wrap"/>
                    </StackPanel>
                  </Border>
                  <Border Grid.Column="2" Style="{StaticResource DiagMetricTile}" Background="#F5F8FB" BorderBrush="#DDE7F0" Margin="4,0,0,0">
                    <StackPanel>
                      <TextBlock Text="HTML REPORT" Style="{StaticResource DiagMetricLabel}"/>
                      <TextBlock x:Name="lblDiagHtmlState" Text="Waiting" FontSize="15" FontWeight="SemiBold" Foreground="#17324D"/>
                      <TextBlock x:Name="lblDiagHtmlEvidence" Text="No report loaded" FontSize="10" Foreground="#5F6B7A" TextWrapping="Wrap"/>
                    </StackPanel>
                  </Border>
                </Grid>
                <Border Background="#F5F8FB" CornerRadius="6" Padding="9,7" Margin="0,9,0,0">
                  <StackPanel>
                    <TextBlock Text="NEXT ACTION" Style="{StaticResource DiagMetricLabel}" FontWeight="SemiBold"/>
                    <TextBlock x:Name="lblDiagNextAction" Text="Select a migration to inspect its report." FontSize="11" Foreground="#17324D" TextWrapping="Wrap"/>
                  </StackPanel>
                </Border>
              </StackPanel>
            </Border>
          </StackPanel>
          <Border x:Name="cardDiagSummary" Grid.Row="1" Grid.ColumnSpan="3" Style="{StaticResource StepCard}" Margin="0,0,0,14">
            <StackPanel>
              <TextBlock Text="SUMMARY" Style="{StaticResource SectionLabel}"/>
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                <Border Grid.Column="0" Background="#FFF8EB" BorderBrush="#F3D8A2" BorderThickness="1" CornerRadius="6" Padding="8" Margin="0,0,5,0">
                  <DockPanel>
                    <Border DockPanel.Dock="Left" Width="30" Height="30" CornerRadius="15" Background="#FCE7BD" Margin="0,0,8,0" VerticalAlignment="Top">
                      <TextBlock Text="&#x26A0;" FontFamily="Segoe UI Symbol" FontSize="17" Foreground="#9A6200" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <StackPanel>
                      <TextBlock Text="ITEMS TO FIX" FontSize="10" FontWeight="SemiBold" Foreground="#705422"/>
                      <TextBlock x:Name="lblSummaryShareGateValue" Text="—" FontSize="20" FontWeight="SemiBold" Foreground="#1F2937"/>
                      <TextBlock x:Name="lblSummaryShareGateDetail" Text="Analyze the latest report" FontSize="10" TextWrapping="Wrap" Foreground="#5F6B7A"/>
                      <TextBlock Text="Source: ShareGate analysis" FontSize="10" Foreground="#8B5E00"/>
                    </StackPanel>
                  </DockPanel>
                </Border>
                <Border Grid.Column="1" Background="#EFF7FF" BorderBrush="#C9E3FA" BorderThickness="1" CornerRadius="6" Padding="8" Margin="5,0,5,0">
                  <DockPanel>
                    <Border DockPanel.Dock="Left" Width="30" Height="30" CornerRadius="15" Background="#DCEFFF" Margin="0,0,8,0" VerticalAlignment="Top">
                      <TextBlock Text="&#x25A3;" FontFamily="Segoe UI Symbol" FontSize="18" Foreground="#1266A3" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <StackPanel>
                      <TextBlock Text="FILES MATCH" FontSize="10" FontWeight="SemiBold" Foreground="#285477"/>
                      <TextBlock x:Name="lblSummaryFilesValue" Text="—" FontSize="20" FontWeight="SemiBold" Foreground="#1F2937"/>
                      <TextBlock x:Name="lblSummaryFilesDetail" Text="Loading comparison" FontSize="10" TextWrapping="Wrap" Foreground="#5F6B7A"/>
                      <TextBlock Text="Source: file comparison" FontSize="10" Foreground="#1266A3"/>
                    </StackPanel>
                  </DockPanel>
                </Border>
                <Border Grid.Column="2" Background="#F0FBF8" BorderBrush="#C9EADF" BorderThickness="1" CornerRadius="6" Padding="8" Margin="5,0,5,0">
                  <DockPanel>
                    <Border DockPanel.Dock="Left" Width="30" Height="30" CornerRadius="15" Background="#D8F2E8" Margin="0,0,8,0" VerticalAlignment="Top">
                      <TextBlock Text="&#x25C6;" FontFamily="Segoe UI Symbol" FontSize="17" Foreground="#167658" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <StackPanel>
                      <TextBlock Text="PERMISSIONS MATCH" FontSize="10" FontWeight="SemiBold" Foreground="#28624F"/>
                      <TextBlock x:Name="lblSummaryPermissionsValue" Text="—" FontSize="20" FontWeight="SemiBold" Foreground="#1F2937"/>
                      <TextBlock x:Name="lblSummaryPermissionsDetail" Text="Loading comparison" FontSize="10" TextWrapping="Wrap" Foreground="#5F6B7A"/>
                      <TextBlock Text="Source: permission comparison" FontSize="10" Foreground="#167658"/>
                    </StackPanel>
                  </DockPanel>
                </Border>
                <Border Grid.Column="3" Background="#F7F3FF" BorderBrush="#E0D5F5" BorderThickness="1" CornerRadius="6" Padding="8" Margin="5,0,0,0">
                  <DockPanel>
                    <Border DockPanel.Dock="Left" Width="30" Height="30" CornerRadius="15" Background="#EDE5FA" Margin="0,0,8,0" VerticalAlignment="Top">
                      <TextBlock Text="&#x21C4;" FontFamily="Segoe UI Symbol" FontSize="18" Foreground="#7353A1" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                    </Border>
                    <StackPanel>
                      <TextBlock Text="SCOPES WITH DIFFERENCES" FontSize="10" FontWeight="SemiBold" Foreground="#58457C"/>
                      <TextBlock x:Name="lblSummaryCrossCheckValue" Text="—" FontSize="20" FontWeight="SemiBold" Foreground="#1F2937"/>
                      <TextBlock x:Name="lblSummaryCrossCheckDetail" Text="Loading cross-check" FontSize="10" TextWrapping="Wrap" Foreground="#5F6B7A"/>
                      <TextBlock Text="Source: cross-check" FontSize="10" Foreground="#7353A1"/>
                    </StackPanel>
                  </DockPanel>
                </Border>
              </Grid>
              <TextBlock Text="Independent measures; percentages use different denominators." FontSize="10" Foreground="#5F6B7A" Margin="0,8,0,0"/>
              <TextBlock Text="LATEST FILE INVENTORIES" Style="{StaticResource SectionLabel}" Margin="0,12,0,6"/>
              <Grid>
                <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                <Border Grid.Column="0" Style="{StaticResource DiagMetricTile}" Background="#EFF7FF" BorderBrush="#C9E3FA" Margin="0,0,5,0">
                  <StackPanel>
                    <TextBlock Text="SOURCE" Style="{StaticResource DiagMetricLabel}" FontWeight="SemiBold"/>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                      <StackPanel Grid.Column="0"><TextBlock Text="FILES" Style="{StaticResource DiagMetricLabel}"/><TextBlock x:Name="lblDiagSourceFiles" Text="—" Style="{StaticResource DiagMetricValue}"/></StackPanel>
                      <StackPanel Grid.Column="1"><TextBlock Text="FOLDERS WITH FILES" Style="{StaticResource DiagMetricLabel}"/><TextBlock x:Name="lblDiagSourceFolders" Text="—" Style="{StaticResource DiagMetricValue}"/></StackPanel>
                      <StackPanel Grid.Column="2"><TextBlock Text="FILE VOLUME" Style="{StaticResource DiagMetricLabel}"/><TextBlock x:Name="lblDiagSourceVolume" Text="—" Style="{StaticResource DiagMetricValue}"/></StackPanel>
                    </Grid>
                    <TextBlock x:Name="lblDiagSourceInventoryEvidence" Text="Waiting for the latest source scan" FontSize="10" Foreground="#5F6B7A" TextWrapping="Wrap" Margin="0,6,0,0"/>
                  </StackPanel>
                </Border>
                <Border Grid.Column="1" Style="{StaticResource DiagMetricTile}" Background="#F0FBF8" BorderBrush="#C9EADF" Margin="5,0,0,0">
                  <StackPanel>
                    <TextBlock Text="DESTINATION" Style="{StaticResource DiagMetricLabel}" FontWeight="SemiBold"/>
                    <Grid>
                      <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                      <StackPanel Grid.Column="0"><TextBlock Text="FILES" Style="{StaticResource DiagMetricLabel}"/><TextBlock x:Name="lblDiagTargetFiles" Text="—" Style="{StaticResource DiagMetricValue}"/></StackPanel>
                      <StackPanel Grid.Column="1"><TextBlock Text="FOLDERS WITH FILES" Style="{StaticResource DiagMetricLabel}"/><TextBlock x:Name="lblDiagTargetFolders" Text="—" Style="{StaticResource DiagMetricValue}"/></StackPanel>
                      <StackPanel Grid.Column="2"><TextBlock Text="FILE VOLUME" Style="{StaticResource DiagMetricLabel}"/><TextBlock x:Name="lblDiagTargetVolume" Text="—" Style="{StaticResource DiagMetricValue}"/></StackPanel>
                    </Grid>
                    <TextBlock x:Name="lblDiagTargetInventoryEvidence" Text="Waiting for the latest destination scan" FontSize="10" Foreground="#5F6B7A" TextWrapping="Wrap" Margin="0,6,0,0"/>
                  </StackPanel>
                </Border>
              </Grid>
              <TextBlock Text="Folders are derived from file paths; empty folders are excluded. File volume sums current SizeBytes only (versions and recycle bin excluded)." FontSize="10" Foreground="#5F6B7A" TextWrapping="Wrap" Margin="0,7,0,0"/>
            </StackPanel>
          </Border>
          <Border Grid.Row="2" Grid.Column="1" Background="#DDE7F0" Margin="0,0,0,14"/>
          <StackPanel Grid.Row="2" Grid.Column="2" Margin="14,0,0,14">
            <TextBlock Text="SHAREGATE ANALYSIS DETAIL" Style="{StaticResource SectionLabel}"/>
            <Border x:Name="cardShareGateDetail" Style="{StaticResource StepCard}">
              <StackPanel>
                <TextBlock x:Name="lblDiagKpis" Text="Analyze the latest ShareGate report to calculate its indicators." TextWrapping="Wrap" FontSize="12" Foreground="#1F2937"/>
                <StackPanel x:Name="panelDiagMetrics" Visibility="Collapsed">
                  <Grid>
                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                    <Border Grid.Column="0" Style="{StaticResource DiagMetricTile}" Background="#EFF7FF" BorderBrush="#C9E3FA" Margin="0,0,4,0">
                      <StackPanel>
                        <TextBlock Text="LINES ANALYZED" Style="{StaticResource DiagMetricLabel}"/>
                        <TextBlock x:Name="lblDiagLines" Text="—" Style="{StaticResource DiagMetricValue}"/>
                      </StackPanel>
                    </Border>
                    <Border Grid.Column="1" Style="{StaticResource DiagMetricTile}" Background="#EFF7FF" BorderBrush="#C9E3FA" Margin="4,0,0,0">
                      <StackPanel>
                        <TextBlock Text="DISTINCT KEYED ITEMS" Style="{StaticResource DiagMetricLabel}"/>
                        <TextBlock x:Name="lblDiagKeyedItems" Text="—" Style="{StaticResource DiagMetricValue}"/>
                      </StackPanel>
                    </Border>
                  </Grid>
                  <TextBlock Text="LINE STATUS" Style="{StaticResource DiagMetricLabel}" Margin="0,9,0,5" FontWeight="SemiBold"/>
                  <Grid>
                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                    <Border Grid.Column="0" Style="{StaticResource DiagMetricTile}" Background="#F0FBF8" BorderBrush="#C9EADF" Margin="0,0,4,0"><StackPanel><TextBlock Text="SUCCESS" Style="{StaticResource DiagMetricLabel}"/><TextBlock x:Name="lblDiagSuccess" Text="—" Style="{StaticResource DiagMetricValue}" Foreground="#167658"/></StackPanel></Border>
                    <Border Grid.Column="1" Style="{StaticResource DiagMetricTile}" Background="#FFF2F2" BorderBrush="#F2D2D2" Margin="4,0,4,0"><StackPanel><TextBlock Text="ERROR" Style="{StaticResource DiagMetricLabel}"/><TextBlock x:Name="lblDiagError" Text="—" Style="{StaticResource DiagMetricValue}" Foreground="#A32939"/></StackPanel></Border>
                    <Border Grid.Column="2" Style="{StaticResource DiagMetricTile}" Background="#FFF8EB" BorderBrush="#F3D8A2" Margin="4,0,4,0"><StackPanel><TextBlock Text="WARNING" Style="{StaticResource DiagMetricLabel}"/><TextBlock x:Name="lblDiagWarning" Text="—" Style="{StaticResource DiagMetricValue}" Foreground="#8B5E00"/></StackPanel></Border>
                    <Border Grid.Column="3" Style="{StaticResource DiagMetricTile}" Margin="4,0,4,0"><StackPanel><TextBlock Text="ACCEPTED" Style="{StaticResource DiagMetricLabel}"/><TextBlock x:Name="lblDiagAccepted" Text="—" Style="{StaticResource DiagMetricValue}"/></StackPanel></Border>
                    <Border Grid.Column="4" Style="{StaticResource DiagMetricTile}" Background="#FFF8EB" BorderBrush="#F3D8A2" Margin="4,0,0,0"><StackPanel><TextBlock Text="TO FIX" Style="{StaticResource DiagMetricLabel}"/><TextBlock x:Name="lblDiagToFixLines" Text="—" Style="{StaticResource DiagMetricValue}" Foreground="#8B5E00"/></StackPanel></Border>
                  </Grid>
                  <TextBlock Text="REVIEW DETAIL" Style="{StaticResource DiagMetricLabel}" Margin="0,9,0,5" FontWeight="SemiBold"/>
                  <Grid>
                    <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                    <Border Grid.Column="0" Style="{StaticResource DiagMetricTile}" Margin="0,0,4,0"><StackPanel><TextBlock Text="UNKEYED LINES" Style="{StaticResource DiagMetricLabel}"/><TextBlock x:Name="lblDiagUnkeyedLines" Text="—" Style="{StaticResource DiagMetricValue}"/></StackPanel></Border>
                    <Border Grid.Column="1" Style="{StaticResource DiagMetricTile}" Background="#FFF8EB" BorderBrush="#F3D8A2" Margin="4,0,4,0"><StackPanel><TextBlock Text="ITEMS TO FIX" Style="{StaticResource DiagMetricLabel}"/><TextBlock x:Name="lblDiagItemsToFix" Text="—" Style="{StaticResource DiagMetricValue}" Foreground="#8B5E00"/></StackPanel></Border>
                    <Border Grid.Column="2" Style="{StaticResource DiagMetricTile}" Margin="4,0,4,0"><StackPanel><TextBlock Text="RESIDUAL LINES" Style="{StaticResource DiagMetricLabel}"/><TextBlock x:Name="lblDiagResidualLines" Text="—" Style="{StaticResource DiagMetricValue}"/></StackPanel></Border>
                    <Border Grid.Column="3" Style="{StaticResource DiagMetricTile}" Margin="4,0,0,0"><StackPanel><TextBlock Text="RESIDUAL ITEMS" Style="{StaticResource DiagMetricLabel}"/><TextBlock x:Name="lblDiagResidualItems" Text="—" Style="{StaticResource DiagMetricValue}"/></StackPanel></Border>
                  </Grid>
                </StackPanel>
                <TextBlock x:Name="lblDiagInterpretation" Text="File, permission and scope details are in Cross-check below. ShareGate residual rates exclude Accepted issues; Fixed is a tracking state, not proof of a successful new migration." TextWrapping="Wrap" FontSize="11" Foreground="#5F6B7A" Margin="0,9,0,0"/>
              </StackPanel>
            </Border>
          </StackPanel>
          <Border Grid.Row="3" Grid.ColumnSpan="3" Style="{StaticResource StepCard}" Margin="0,0,0,14">
            <StackPanel>
              <DockPanel LastChildFill="False" Margin="0,0,0,7">
                <TextBlock Text="CROSS-CHECK: SHAREGATE / FILES / PERMISSIONS" Style="{StaticResource SectionLabel}" DockPanel.Dock="Left"/>
                <Button x:Name="btnCrossCheckRefresh" Content="Refresh cross-check" Style="{StaticResource BtnGhost}" Width="130" DockPanel.Dock="Right"/>
              </DockPanel>
              <TextBlock Text="Three separate measures. Comparison percentages use source inventory keys; ShareGate issues use report items and lines. Scope matches are indicative, not proof for an individual item."
                         TextWrapping="Wrap" FontSize="11" Foreground="#5F6B7A" Margin="0,0,0,7"/>
              <DataGrid x:Name="gridCrossCheckEvidence" Height="110" AutoGenerateColumns="False" IsReadOnly="True" CanUserAddRows="False" HeadersVisibility="Column" AlternatingRowBackground="#F7FAFE">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Evidence" Binding="{Binding Evidence}" Width="100"/>
                  <DataGridTextColumn Header="Date" Binding="{Binding Date}" Width="145"/>
                  <DataGridTextColumn Header="Rate" Binding="{Binding Rate}" Width="80"/>
                  <DataGridTextColumn Header="Coverage" Binding="{Binding Coverage}" Width="170"/>
                  <DataGridTextColumn Header="Differences" Binding="{Binding Differences}" Width="*"/>
                  <DataGridTextColumn Header="Evidence state" Binding="{Binding State}" Width="200"/>
                </DataGrid.Columns>
              </DataGrid>
              <StackPanel Orientation="Horizontal" Margin="0,7,0,7">
                <Button x:Name="btnCrossCheckFilesReport" Content="Open files report" Style="{StaticResource BtnGhost}" Width="115" IsEnabled="False"/>
                <Button x:Name="btnCrossCheckPermissionsReport" Content="Open permissions report" Style="{StaticResource BtnGhost}" Width="145" Margin="7,0,0,0" IsEnabled="False"/>
                <TextBlock x:Name="lblCrossCheckStatus" Text="Select a migration to compare existing reports." Margin="12,4,0,0" TextWrapping="Wrap" FontSize="11" Foreground="#5F6B7A"/>
              </StackPanel>
              <TextBlock Text="DIFFERENCES BY SITE AND LIST" Style="{StaticResource SectionLabel}" Margin="0,3,0,6"/>
              <DataGrid x:Name="gridCrossCheckScopes" Height="175" AutoGenerateColumns="False" IsReadOnly="True" CanUserAddRows="False" HeadersVisibility="Column" AlternatingRowBackground="#F7FAFE">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Site (matched scope)" Binding="{Binding Site}" Width="195"/>
                  <DataGridTextColumn Header="List" Binding="{Binding List}" Width="130"/>
                  <DataGridTextColumn Header="ShareGate to fix" Binding="{Binding ShareGateToFix}" Width="105"/>
                  <DataGridTextColumn Header="Files missing" Binding="{Binding FilesMissing}" Width="85"/>
                  <DataGridTextColumn Header="Files extra" Binding="{Binding FilesExtra}" Width="75"/>
                  <DataGridTextColumn Header="File change flags" Binding="{Binding FileChangeFlags}" Width="90"/>
                  <DataGridTextColumn Header="Perms missing" Binding="{Binding PermsMissing}" Width="90"/>
                  <DataGridTextColumn Header="Disabled missing" Binding="{Binding PermsDisabled}" Width="90"/>
                  <DataGridTextColumn Header="Perms extra" Binding="{Binding PermsExtra}" Width="75"/>
                  <DataGridTextColumn Header="Perms changed" Binding="{Binding PermsChanged}" Width="95"/>
                  <DataGridTextColumn Header="Interpretation" Binding="{Binding Assessment}" Width="*"/>
                </DataGrid.Columns>
              </DataGrid>
            </StackPanel>
          </Border>
          <StackPanel Grid.Row="4" Grid.ColumnSpan="3">
          <Border x:Name="cardTransient" Style="{StaticResource StepCard}" Visibility="Collapsed">
            <StackPanel>
              <TextBlock Text="SHAREGATE 401 RETRY RESULTS" Style="{StaticResource SectionLabel}"/>
              <TextBlock x:Name="lblTransientStatus" Text="No batch result found." TextWrapping="Wrap" FontSize="12"/>
              <TextBlock x:Name="lblTransientCounts" Text="A reviewed ShareGate batch run will appear here." TextWrapping="Wrap" FontSize="12" Foreground="#5F6B7A" Margin="0,5,0,0"/>
              <StackPanel Orientation="Horizontal" Margin="0,7,0,0">
                <Button x:Name="btnTransientRefresh" Content="Refresh batch results" Style="{StaticResource BtnGhost}" Width="128"/>
                <Button x:Name="btnTransientOpen" Content="Open results CSV" Style="{StaticResource BtnGhost}" Width="112" Margin="6,0,0,0" IsEnabled="False"/>
                <Button x:Name="btnTransientOutOfBatch" Content="Open excluded CSV" Style="{StaticResource BtnGhost}" Width="135" Margin="6,0,0,0" IsEnabled="False"/>
              </StackPanel>
            </StackPanel>
          </Border>
          <Border x:Name="farmDiagnosticsHost" Visibility="Collapsed">
          <Border x:Name="cardFarmDiagnostics" Style="{StaticResource StepCard}">
            <StackPanel>
              <TextBlock Text="SOURCE FARM DIAGNOSTICS" Style="{StaticResource SectionLabel}"/>
              <TextBlock x:Name="lblFarmResult" Text="No farm diagnostic result found." TextWrapping="Wrap" FontSize="12"/>
              <TextBlock x:Name="lblFarmPeaks" Text="Analyze ShareGate reports to prepare farm windows." TextWrapping="Wrap" FontSize="11" Foreground="#5F6B7A" Margin="0,5,0,0"/>
              <StackPanel Orientation="Horizontal" Margin="0,7,0,5">
                <Button x:Name="btnFarmRefresh" Content="Refresh farm results" Style="{StaticResource BtnGhost}" Width="120"/>
                <Button x:Name="btnFarmOpenReport" Content="Open farm report" Style="{StaticResource BtnGhost}" Width="110" Margin="6,0,0,0" IsEnabled="False"/>
                <Button x:Name="btnFarmCheck" Content="Check prerequisites" Style="{StaticResource BtnGhost}" Width="132" Margin="6,0,0,0"/>
                <Button x:Name="btnFarmRun" Content="Run (read-only)" Style="{StaticResource Btn}" Width="110" Margin="6,0,0,0" IsEnabled="False"/>
              </StackPanel>
              <TextBlock x:Name="lblFarmPrerequisites" Text="Requires: farm server, elevated account with SharePoint Shell Admin rights, Windows PowerShell 5.1, SharePoint snap-in/module, shared UNC toolkit and valid access-peak CSV. Run the displayed DryRun first, then check prerequisites." TextWrapping="Wrap" FontSize="11" Foreground="#5F6B7A"/>
              <TextBox x:Name="txtFarmDryRun" IsReadOnly="True" Height="28" Margin="0,5,0,0" VerticalContentAlignment="Center" FontFamily="Consolas" FontSize="10" ToolTip="DryRun command for the farm server"/>
              <TextBox x:Name="txtFarmRun" IsReadOnly="True" Height="28" Margin="0,5,0,0" VerticalContentAlignment="Center" FontFamily="Consolas" FontSize="10" ToolTip="Real read-only diagnostic command for the farm server"/>
            </StackPanel>
          </Border>
          </Border>
          </StackPanel>
          <StackPanel x:Name="panelDiagReview" Grid.Row="5" Grid.ColumnSpan="3" IsEnabled="False">
          <TextBlock Text="ISSUE REVIEW" Style="{StaticResource SectionLabel}"/>
          <Border Style="{StaticResource StepCard}">
            <StackPanel>
              <TextBlock Text="ISSUE PATTERNS" Style="{StaticResource SectionLabel}"/>
              <StackPanel Orientation="Horizontal" Margin="0,0,0,7">
                <TextBlock Text="Session" VerticalAlignment="Center" Margin="0,0,5,0"/>
                <ComboBox x:Name="cmbDiagSession" Width="100" Height="27"/>
                <TextBlock Text="Status" VerticalAlignment="Center" Margin="10,0,5,0"/>
                <ComboBox x:Name="cmbDiagStatus" Width="100" Height="27">
                  <ComboBoxItem Content="All statuses" IsSelected="True"/>
                  <ComboBoxItem Content="Error"/>
                  <ComboBoxItem Content="Warning"/>
                </ComboBox>
                <TextBox x:Name="txtDiagFilter" Width="170" Height="27" Margin="10,0,0,0" VerticalContentAlignment="Center" ToolTip="Filter category or pattern text"/>
                <Button x:Name="btnDiagFilter" Content="Filter" Style="{StaticResource BtnGhost}" Width="62" Margin="6,0,0,0"/>
              </StackPanel>
              <DataGrid x:Name="gridDiagPatterns" Height="230" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single" CanUserAddRows="False" HeadersVisibility="Column" AlternatingRowBackground="#F7FAFE">
                <DataGrid.Columns>
                  <DataGridTextColumn Header="Category" Binding="{Binding Category}" Width="160"/>
                  <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="70"/>
                  <DataGridTextColumn Header="State" Binding="{Binding State}" Width="75"/>
                  <DataGridTextColumn Header="Lines" Binding="{Binding Lines}" Width="55"/>
                  <DataGridTextColumn Header="Items" Binding="{Binding Items}" Width="55"/>
                  <DataGridTextColumn Header="Pattern" Binding="{Binding Pattern}" Width="*"/>
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
          <Button x:Name="btnFarmWindow" Content="Source farm diagnostics…" Style="{StaticResource BtnGhost}" HorizontalAlignment="Left" Margin="0,2,0,0"/>
          </StackPanel>
        </Grid>

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
    <Border Grid.Row="4" Background="#F5F8FB" BorderBrush="#DDE7F0" BorderThickness="0,1,0,0" Padding="18,5">
      <TextBlock x:Name="lblAppVersion" HorizontalAlignment="Right" FontSize="10" Foreground="#5F6B7A"/>
    </Border>
  </Grid>
</Window>
'@

if ($ValidateOnly) {
    try {
        $reader = [System.Xml.XmlNodeReader]::new($xaml)
        $validationWindow = [System.Windows.Markup.XamlReader]::Load($reader)
        $farmHost = $validationWindow.FindName('farmDiagnosticsHost')
        $farmCard = $farmHost.Child; $farmHost.Child = $null
        $farmValidationWindow = New-FarmDiagnosticsWindow -Owner $validationWindow -MigrationName 'Validation' -Content $farmCard
        $farmValidationWindow.Content.Children[1].Content = $null
        $farmValidationWindow.Close()
        $farmHost.Child = $farmCard
        $null = Initialize-SmartM365BatchGui -Panel $validationWindow.FindName('panelBatches') -Root $script:ScriptRoot -SourceRoot $script:FarmToolkitRoot -ValidateOnly
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
    $availableWidth = [System.Windows.SystemParameters]::WorkArea.Width - 24
    $script:Window.Width = [math]::Max($script:Window.MinWidth,
        [math]::Min($script:Window.Width, $availableWidth))
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
$lblAppVersion   = ctrl 'lblAppVersion'
$lblAppVersion.Text = 'v' + $script:AppVersion

# Context bar
$lblSourceType = ctrl 'lblSourceType'
$lblSourceUrl  = ctrl 'lblSourceUrl'
$lblTargetType = ctrl 'lblTargetType'
$lblTargetUrl  = ctrl 'lblTargetUrl'
$cmbAuthMode   = ctrl 'cmbAuthMode'
$btnOpenConfig = ctrl 'btnOpenConfig'

# Tabs
$tabSummary     = ctrl 'tabSummary'
$tabFiles       = ctrl 'tabFiles'
$tabOperations  = ctrl 'tabOperations'
$tabDiagnostics = ctrl 'tabDiagnostics'
$tabBatches = ctrl 'tabBatches'
$tabLogs        = ctrl 'tabLogs'
$tabConfig      = ctrl 'tabConfig'

# Panels
$panelSummary     = ctrl 'panelSummary'
$gridSummary      = ctrl 'gridSummary'
$lblSummaryStatus = ctrl 'lblSummaryStatus'
$lblPortfolioMigrations = ctrl 'lblPortfolioMigrations'
$lblPortfolioFilesCompared = ctrl 'lblPortfolioFilesCompared'
$lblPortfolioPermissionsCompared = ctrl 'lblPortfolioPermissionsCompared'
$lblPortfolioActions = ctrl 'lblPortfolioActions'
$lblPortfolioRefreshErrors = ctrl 'lblPortfolioRefreshErrors'
$panelWorkflows   = ctrl 'panelWorkflows'
$panelOperations  = ctrl 'panelOperations'
$panelBatches = ctrl 'panelBatches'
$panelDiagnostics = ctrl 'panelDiagnostics'
$panelLogs        = ctrl 'panelLogs'
$panelConfig      = ctrl 'panelConfig'

# Files step
$badgeScanSrc  = ctrl 'badgeScanSrc'
$lblScanSrcAge = ctrl 'lblScanSrcAge'
$cmbScanSrcFile= ctrl 'cmbScanSrcFile'
$btnOpenScanSrc= ctrl 'btnOpenScanSrc'
$btnOpenSourceSite = ctrl 'btnOpenSourceSite'
$btnOpenTargetSite = ctrl 'btnOpenTargetSite'
$btnOpenSourceSitePerm = ctrl 'btnOpenSourceSitePerm'
$btnOpenTargetSitePerm = ctrl 'btnOpenTargetSitePerm'
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
$btnReportCmpFiles = ctrl 'btnReportCmpFiles'
$btnRunCmpFiles = ctrl 'btnRunCmpFiles'
$lblCmpFilesAvailability = ctrl 'lblCmpFilesAvailability'

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
$cmbScanSrcPermFile= ctrl 'cmbScanSrcPermFile'
$btnOpenScanSrcPerm= ctrl 'btnOpenScanSrcPerm'
$btnRunScanSrcPerm = ctrl 'btnRunScanSrcPerm'

$badgeScanTgtPerm  = ctrl 'badgeScanTgtPerm'
$lblScanTgtPermAge = ctrl 'lblScanTgtPermAge'
$cmbScanTgtPermFile= ctrl 'cmbScanTgtPermFile'
$btnOpenScanTgtPerm= ctrl 'btnOpenScanTgtPerm'
$btnRunScanTgtPerm = ctrl 'btnRunScanTgtPerm'

$badgeCmpPerms  = ctrl 'badgeCmpPerms'
$lblCmpPermsAge = ctrl 'lblCmpPermsAge'
$lblCmpPermsDir = ctrl 'lblCmpPermsDir'
$btnOpenCmpPerms= ctrl 'btnOpenCmpPerms'
$btnReportCmpPerms = ctrl 'btnReportCmpPerms'
$btnRunCmpPerms = ctrl 'btnRunCmpPerms'
$lblCmpPermsAvailability = ctrl 'lblCmpPermsAvailability'
$badgePermHistory = ctrl 'badgePermHistory'
$lblPermHistoryAge = ctrl 'lblPermHistoryAge'
$cmbPermHistorySide = ctrl 'cmbPermHistorySide'
$cmbPermHistoryOldFile = ctrl 'cmbPermHistoryOldFile'
$cmbPermHistoryNewFile = ctrl 'cmbPermHistoryNewFile'
$btnOpenPermHistory = ctrl 'btnOpenPermHistory'
$btnRunPermHistory = ctrl 'btnRunPermHistory'

# Operations
$listOps   = ctrl 'listOps'
$lblNoOps  = ctrl 'lblNoOps'

# Migration diagnostics
$lblDiagMigration = ctrl 'lblDiagMigration'
$lblDiagScope = ctrl 'lblDiagScope'
$lblDiagInputPath = ctrl 'lblDiagInputPath'
$lblDiagLatestReport = ctrl 'lblDiagLatestReport'
$btnDiagOpenFolder = ctrl 'btnDiagOpenFolder'
$btnDiagRefresh = ctrl 'btnDiagRefresh'
$btnDiagAnalyze = ctrl 'btnDiagAnalyze'
$btnDiagOpenReport = ctrl 'btnDiagOpenReport'
$cardDiagSummary = ctrl 'cardDiagSummary'
$cardTransient = ctrl 'cardTransient'
$panelDiagReview = ctrl 'panelDiagReview'
$lblDiagProgress = ctrl 'lblDiagProgress'
$panelDiagAnalysisLoading = ctrl 'panelDiagAnalysisLoading'
$lblDiagReportState = ctrl 'lblDiagReportState'
$lblDiagReportEvidence = ctrl 'lblDiagReportEvidence'
$lblDiagAnalysisState = ctrl 'lblDiagAnalysisState'
$lblDiagAnalysisEvidence = ctrl 'lblDiagAnalysisEvidence'
$lblDiagHtmlState = ctrl 'lblDiagHtmlState'
$lblDiagHtmlEvidence = ctrl 'lblDiagHtmlEvidence'
$lblDiagNextAction = ctrl 'lblDiagNextAction'
$lblDiagKpis = ctrl 'lblDiagKpis'
$panelDiagMetrics = ctrl 'panelDiagMetrics'
$lblDiagLines = ctrl 'lblDiagLines'
$lblDiagKeyedItems = ctrl 'lblDiagKeyedItems'
$lblDiagSuccess = ctrl 'lblDiagSuccess'
$lblDiagError = ctrl 'lblDiagError'
$lblDiagWarning = ctrl 'lblDiagWarning'
$lblDiagAccepted = ctrl 'lblDiagAccepted'
$lblDiagToFixLines = ctrl 'lblDiagToFixLines'
$lblDiagUnkeyedLines = ctrl 'lblDiagUnkeyedLines'
$lblDiagItemsToFix = ctrl 'lblDiagItemsToFix'
$lblDiagResidualLines = ctrl 'lblDiagResidualLines'
$lblDiagResidualItems = ctrl 'lblDiagResidualItems'
$lblDiagInterpretation = ctrl 'lblDiagInterpretation'
$lblSummaryShareGateValue = ctrl 'lblSummaryShareGateValue'
$lblSummaryShareGateDetail = ctrl 'lblSummaryShareGateDetail'
$lblSummaryFilesValue = ctrl 'lblSummaryFilesValue'
$lblSummaryFilesDetail = ctrl 'lblSummaryFilesDetail'
$lblSummaryPermissionsValue = ctrl 'lblSummaryPermissionsValue'
$lblSummaryPermissionsDetail = ctrl 'lblSummaryPermissionsDetail'
$lblSummaryCrossCheckValue = ctrl 'lblSummaryCrossCheckValue'
$lblSummaryCrossCheckDetail = ctrl 'lblSummaryCrossCheckDetail'
$lblDiagSourceFiles = ctrl 'lblDiagSourceFiles'
$lblDiagSourceFolders = ctrl 'lblDiagSourceFolders'
$lblDiagSourceVolume = ctrl 'lblDiagSourceVolume'
$lblDiagSourceInventoryEvidence = ctrl 'lblDiagSourceInventoryEvidence'
$lblDiagTargetFiles = ctrl 'lblDiagTargetFiles'
$lblDiagTargetFolders = ctrl 'lblDiagTargetFolders'
$lblDiagTargetVolume = ctrl 'lblDiagTargetVolume'
$lblDiagTargetInventoryEvidence = ctrl 'lblDiagTargetInventoryEvidence'
$gridCrossCheckEvidence = ctrl 'gridCrossCheckEvidence'
$gridCrossCheckScopes = ctrl 'gridCrossCheckScopes'
$lblCrossCheckStatus = ctrl 'lblCrossCheckStatus'
$panelCrossCheckLoading = ctrl 'panelCrossCheckLoading'
$lblCrossCheckLoading = ctrl 'lblCrossCheckLoading'
$btnCrossCheckRefresh = ctrl 'btnCrossCheckRefresh'
$btnCrossCheckFilesReport = ctrl 'btnCrossCheckFilesReport'
$btnCrossCheckPermissionsReport = ctrl 'btnCrossCheckPermissionsReport'
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
$lblTransientStatus = ctrl 'lblTransientStatus'
$lblTransientCounts = ctrl 'lblTransientCounts'
$btnTransientRefresh = ctrl 'btnTransientRefresh'
$btnTransientOpen = ctrl 'btnTransientOpen'
$btnTransientOutOfBatch = ctrl 'btnTransientOutOfBatch'
$lblFarmResult = ctrl 'lblFarmResult'
$lblFarmPeaks = ctrl 'lblFarmPeaks'
$btnFarmRefresh = ctrl 'btnFarmRefresh'
$btnFarmOpenReport = ctrl 'btnFarmOpenReport'
$btnFarmCheck = ctrl 'btnFarmCheck'
$btnFarmRun = ctrl 'btnFarmRun'
$btnFarmWindow = ctrl 'btnFarmWindow'
$farmDiagnosticsHost = ctrl 'farmDiagnosticsHost'
$cardFarmDiagnostics = ctrl 'cardFarmDiagnostics'
$lblFarmPrerequisites = ctrl 'lblFarmPrerequisites'
$txtFarmDryRun = ctrl 'txtFarmDryRun'
$txtFarmRun = ctrl 'txtFarmRun'

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
$btnOverviewGlobalFileReport = ctrl 'btnOverviewGlobalFileReport'
$btnOverviewGlobalPermissionsReport = ctrl 'btnOverviewGlobalPermissionsReport'

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
$script:SummaryLoading = $false
$script:WizardOpen = $false
$script:TargetScopeMismatch = $false
$script:ConfigEditorLoadedHash = ''
$script:DiagInputPath = ''
$script:DiagLatestReport = $null
$script:DiagReportSignature = ''
$script:DiagReportHash = ''
$script:DiagLoading = $false
$script:DiagAnalysisVerified = $false
$script:DiagSummary = $null
$script:CrossCheckSignature = ''
$script:CrossCheckRequestedSignature = ''
$script:CrossCheckJobSignature = ''
$script:CrossCheckJob = $null
$script:CrossCheckResult = $null
$script:DiagRows = @()
$script:DiagKnownSessions = @()
$script:DiagProcess = $null
$script:DiagTimer = $null
$script:DiagOutputDirectory = ''
$script:DiagActivity = ''
$script:DiagProjectRoot = ''
$script:DiagActiveReportKey = ''
$script:DiagAutoAttempts = @{}
$script:DiagLoadedDirectory = ''
$script:FarmReportPath = ''
$script:FarmInvocation = $null
$script:FarmPrerequisitesPassed = $false
$script:TransientResultsPath = ''
$script:TransientOutOfBatchPath = ''

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
    $previousLatest = [string]$ComboBox.Tag
    $ComboBox.Items.Clear()
    foreach ($item in @($Items)) {
        if ($null -eq $item -or -not $item.PSObject.Properties['FullName']) { continue }
        [void]$ComboBox.Items.Add($item)
    }
    $ComboBox.IsEnabled = ($ComboBox.Items.Count -gt 0)
    if ($ComboBox.Items.Count -eq 0) {
        $ComboBox.SelectedIndex = -1
        $ComboBox.Tag = ''
        return
    }

    $latest = [string]$ComboBox.Items[0].FullName
    $selectedIndex = 0
    if ($previous -and $previous -ne $previousLatest) {
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
    $ComboBox.Tag = $latest
}

function Get-SelectedScanFile {
    param([System.Windows.Controls.ComboBox]$ComboBox)
    if ($null -eq $ComboBox -or $null -eq $ComboBox.SelectedItem) { return $null }
    return $ComboBox.SelectedItem.File
}

function Test-ComparisonInventoryHasRows {
    param([System.IO.FileInfo]$File)
    $receiptPath = $File.FullName + '.manifest.json.txt'
    if (Test-Path -LiteralPath $receiptPath -PathType Leaf) {
        $receipt = Get-Content -LiteralPath $receiptPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($receipt.SchemaVersion -ne 1 -or $receipt.InventoryFile -ne $File.Name -or
            -not $receipt.PSObject.Properties['Rows'] -or $null -eq $receipt.Rows -or
            [long]$receipt.Rows -lt 0 -or -not $receipt.Sha256 -or -not $receipt.CompletedAtUtc) {
            throw 'Invalid scan receipt. Rerun this inventory scan.'
        }
        return [long]$receipt.Rows -gt 0
    }
    # Older scans have no receipt: stop after the first record, without counting the inventory.
    return @(Import-Csv -LiteralPath $File.FullName -Delimiter ';' -ErrorAction Stop | Select-Object -First 1).Count -gt 0
}

function Get-ComparisonRunState {
    param([System.IO.FileInfo]$FirstCsv, [System.IO.FileInfo]$SecondCsv, [switch]$History, [switch]$AllowEmptyTarget)
    $issues = [System.Collections.Generic.List[string]]::new()
    $labels = if ($History) { @('Previous', 'Current') } else { @('Source', 'Target') }
    $files = @($FirstCsv, $SecondCsv)
    for ($i = 0; $i -lt 2; $i++) {
        $file = $files[$i]
        if ($null -eq $file) {
            $issues.Add("$($labels[$i]) scan unavailable (missing or incomplete).")
        }
        elseif (-not (Test-Path -LiteralPath $file.FullName -PathType Leaf)) {
            $issues.Add("$($labels[$i]) scan file is no longer available.")
        }
        elseif (-not (Test-SmartM365InventoryCsvComplete -File $file)) {
            $issues.Add("$($labels[$i]) scan is incomplete: an Errors.csv file is present.")
        }
        elseif (-not $History) {
            try {
                if (-not (Test-ComparisonInventoryHasRows -File $file) -and ($i -eq 0 -or -not $AllowEmptyTarget)) {
                    $issues.Add("$($labels[$i]) inventory is empty (0 data rows). No comparison percentage can be calculated.")
                }
            }
            catch { $issues.Add("$($labels[$i]) scan cannot be verified: $($_.Exception.Message)") }
        }
    }
    if ($History -and $FirstCsv -and $SecondCsv -and
        [string]::Equals($FirstCsv.FullName, $SecondCsv.FullName, [StringComparison]::OrdinalIgnoreCase)) {
        $issues.Add('Select two different completed scans.')
    }
    $reason = $issues -join ' '
    if ($issues.Count -gt 0 -and -not $History) { $reason += ' Check the scan scope and logs; obtain a complete inventory with data before comparing.' }
    return [pscustomobject]@{ Ready = ($issues.Count -eq 0); Reason = $reason }
}

function Update-ComparisonRunState {
    foreach ($controls in @(
        @{ Source=$cmbScanSrcFile; Target=$cmbScanTgtFile; Button=$btnRunCmpFiles; Label=$lblCmpFilesAvailability; Result='FileComparisonFolder'; Attempt='FileComparisonAttemptFolder'; AllowEmptyTarget=$true },
        @{ Source=$cmbScanSrcPermFile; Target=$cmbScanTgtPermFile; Button=$btnRunCmpPerms; Label=$lblCmpPermsAvailability; Result='PermComparisonFolder'; Attempt='PermComparisonAttemptFolder'; AllowEmptyTarget=$false }
    )) {
        $state = Get-ComparisonRunState -FirstCsv (Get-SelectedScanFile $controls.Source) -SecondCsv (Get-SelectedScanFile $controls.Target) -AllowEmptyTarget:$controls.AllowEmptyTarget
        $controls.Button.IsEnabled = $state.Ready
        $controls.Button.ToolTip = if ($state.Ready) { 'Compare the selected completed source and target scans.' } else { $state.Reason }
        $resultNote = ''
        if ($script:CurrentStatus -and $script:CurrentStatus.PSObject.Properties[$controls.Attempt]) {
            $attempt = $script:CurrentStatus.($controls.Attempt)
            $result = $script:CurrentStatus.($controls.Result)
            if ($attempt -and (-not $result -or $attempt.FullName -ne $result.FullName)) {
                $stamp = Get-SmartM365PortfolioTimestamp $attempt.Name
                if (-not $stamp) { $stamp = $attempt.LastWriteTime }
                $resultNote = 'Latest attempt: {0}; no comparison result was produced.' -f (Format-RunAgeText $stamp)
                if ($result) { $resultNote += ' Showing the previous available comparison result.' }
            }
        }
        $controls.Label.Text = (@($state.Reason, $resultNote) | Where-Object { $_ }) -join ' '
        $controls.Label.Visibility = if ($controls.Label.Text) { 'Visible' } else { 'Collapsed' }
    }
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
    if (-not $selectedFile) { $age.Text = 'No complete scan' }
    Set-Badge $Badge $Label $age.Text $age.HasRun
    $OpenButton.Visibility = if ($selectedFile) { 'Visible' } else { 'Collapsed' }
    if ($selectedFile) { $OpenButton.Tag = (Split-Path $selectedFile.FullName -Parent) }
    Update-ComparisonRunState
}
function Get-HistorySide {
    param([System.Windows.Controls.ComboBox]$ComboBox = $cmbHistorySide)
    $sel = $ComboBox.SelectedItem
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
    foreach ($item in @($Items)) {
        if ($null -eq $item -or -not $item.PSObject.Properties['FullName']) { continue }
        [void]$ComboBox.Items.Add($item)
    }
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
    $state = Get-ComparisonRunState -FirstCsv $oldCsv -SecondCsv $newCsv -History
    $btnRunHistory.IsEnabled = $state.Ready
    $btnRunHistory.ToolTip = $state.Reason
}

function Update-HistoryScanSelection {
    if ($null -eq $script:CurrentStatus) { return }

    $items = if ((Get-HistorySide) -eq 'Target') { @($script:CurrentStatus.TargetFileCsvItems) } else { @($script:CurrentStatus.SourceFileCsvItems) }
    Set-HistoryComboItems $cmbHistoryNewFile $items 0
    Set-HistoryComboItems $cmbHistoryOldFile $items 1
    Update-HistoryRunState
}

function Update-PermissionHistoryRunState {
    $oldCsv = Get-SelectedScanFile -ComboBox $cmbPermHistoryOldFile
    $newCsv = Get-SelectedScanFile -ComboBox $cmbPermHistoryNewFile
    $state = Get-ComparisonRunState -FirstCsv $oldCsv -SecondCsv $newCsv -History
    $btnRunPermHistory.IsEnabled = $state.Ready
    $btnRunPermHistory.ToolTip = $state.Reason
}

function Update-PermissionHistoryScanSelection {
    if ($null -eq $script:CurrentStatus) { return }
    $items = if ((Get-HistorySide -ComboBox $cmbPermHistorySide) -eq 'Target') {
        @($script:CurrentStatus.TargetPermCsvItems)
    } else { @($script:CurrentStatus.SourcePermCsvItems) }
    $items = @($items | Sort-Object {
        $stamp = Get-SmartM365PortfolioTimestamp $_.Name
        if ($stamp) { $stamp } else { $_.File.LastWriteTime }
    } -Descending)
    Set-HistoryComboItems $cmbPermHistoryNewFile $items 0
    Set-HistoryComboItems $cmbPermHistoryOldFile $items 1
    Update-PermissionHistoryRunState
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
    if ($Action -in @('CompareFiles', 'ComparePermissions')) {
        $permissions = $Action -eq 'ComparePermissions'
        try {
            $script:CurrentStatus = Get-MigrationStatus -Migration $script:CurrentMigration
            $st = $script:CurrentStatus
            if ($permissions) {
                Set-ScanComboItems $cmbScanSrcPermFile @($st.SourcePermCsvItems) $st.SourcePermCsv
                Update-ScanFileSelection $cmbScanSrcPermFile $badgeScanSrcPerm $lblScanSrcPermAge $btnOpenScanSrcPerm
                Set-ScanComboItems $cmbScanTgtPermFile @($st.TargetPermCsvItems) $st.TargetPermCsv
                Update-ScanFileSelection $cmbScanTgtPermFile $badgeScanTgtPerm $lblScanTgtPermAge $btnOpenScanTgtPerm
            }
            else {
                Set-ScanComboItems $cmbScanSrcFile @($st.SourceFileCsvItems) $st.SourceFileCsv
                Update-ScanFileSelection $cmbScanSrcFile $badgeScanSrc $lblScanSrcAge $btnOpenScanSrc
                Set-ScanComboItems $cmbScanTgtFile @($st.TargetFileCsvItems) $st.TargetFileCsv
                Update-ScanFileSelection $cmbScanTgtFile $badgeScanTgt $lblScanTgtAge $btnOpenScanTgt
            }
        }
        catch {
            $kind = if ($permissions) { 'permission' } else { 'file' }
            [System.Windows.MessageBox]::Show("Could not refresh $kind scans before comparison:`n$($_.Exception.Message)",
                $script:AppName, 'OK', 'Error') | Out-Null
            return
        }

        $sourceCombo = if ($permissions) { $cmbScanSrcPermFile } else { $cmbScanSrcFile }
        $targetCombo = if ($permissions) { $cmbScanTgtPermFile } else { $cmbScanTgtFile }
        $state = Get-ComparisonRunState -FirstCsv (Get-SelectedScanFile $sourceCombo) -SecondCsv (Get-SelectedScanFile $targetCombo) -AllowEmptyTarget:(-not $permissions)
        Update-ComparisonRunState
        if (-not $state.Ready) {
            [System.Windows.MessageBox]::Show($state.Reason, $script:AppName, 'OK', 'Warning') | Out-Null
            return
        }

        $olderScans = [System.Collections.Generic.List[string]]::new()
        $pairs = if ($permissions) { @(
            @{ Side = 'Source'; ComboBox = $cmbScanSrcPermFile; Latest = $st.SourcePermCsv },
            @{ Side = 'Target'; ComboBox = $cmbScanTgtPermFile; Latest = $st.TargetPermCsv }
        ) } else { @(
            @{ Side = 'Source'; ComboBox = $cmbScanSrcFile; Latest = $st.SourceFileCsv },
            @{ Side = 'Target'; ComboBox = $cmbScanTgtFile; Latest = $st.TargetFileCsv }
        ) }
        foreach ($pair in $pairs) {
            $selected = Get-SelectedScanFile -ComboBox $pair.ComboBox
            if ($selected -and $pair.Latest -and
                -not [string]::Equals($selected.FullName, $pair.Latest.FullName,
                    [System.StringComparison]::OrdinalIgnoreCase)) {
                $olderScans.Add(('{0}: {1} (latest: {2})' -f $pair.Side, $selected.Name, $pair.Latest.Name))
            }
        }
        if ($olderScans.Count -gt 0) {
            $answer = [System.Windows.MessageBox]::Show(
                ("An older inventory is selected while a newer scan is available:`n{0}`n`nContinue with the selected inventory?" -f ($olderScans -join "`n")),
                $script:AppName, 'YesNo', 'Warning')
            if ($answer -ne [System.Windows.MessageBoxResult]::Yes) { return }
        }
    }
    if ($Action -in @('CompareScanHistory', 'ComparePermissionScanHistory')) {
        $oldCombo = if ($Action -eq 'ComparePermissionScanHistory') { $cmbPermHistoryOldFile } else { $cmbHistoryOldFile }
        $newCombo = if ($Action -eq 'ComparePermissionScanHistory') { $cmbPermHistoryNewFile } else { $cmbHistoryNewFile }
        $state = Get-ComparisonRunState -FirstCsv (Get-SelectedScanFile $oldCombo) -SecondCsv (Get-SelectedScanFile $newCombo) -History
        if (-not $state.Ready) {
            Update-HistoryRunState
            Update-PermissionHistoryRunState
            [System.Windows.MessageBox]::Show($state.Reason, $script:AppName, 'OK', 'Warning') | Out-Null
            return
        }
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

    if ($Action -in @('CompareFiles', 'ComparePermissions')) {
        $sourceCombo = if ($Action -eq 'ComparePermissions') { $cmbScanSrcPermFile } else { $cmbScanSrcFile }
        $targetCombo = if ($Action -eq 'ComparePermissions') { $cmbScanTgtPermFile } else { $cmbScanTgtFile }
        $sourceCsv = Get-SelectedScanFile -ComboBox $sourceCombo
        $targetCsv = Get-SelectedScanFile -ComboBox $targetCombo
        if ($sourceCsv) { $args += @('-SourceCsv', "`"$($sourceCsv.FullName)`"") }
        if ($targetCsv) { $args += @('-TargetCsv', "`"$($targetCsv.FullName)`"") }
    }
    elseif ($Action -in @('CompareScanHistory', 'ComparePermissionScanHistory')) {
        $permissionHistory = $Action -eq 'ComparePermissionScanHistory'
        $oldCombo = if ($permissionHistory) { $cmbPermHistoryOldFile } else { $cmbHistoryOldFile }
        $newCombo = if ($permissionHistory) { $cmbPermHistoryNewFile } else { $cmbHistoryNewFile }
        $sideCombo = if ($permissionHistory) { $cmbPermHistorySide } else { $cmbHistorySide }
        $oldCsv = Get-SelectedScanFile -ComboBox $oldCombo
        $newCsv = Get-SelectedScanFile -ComboBox $newCombo
        $args += @('-HistorySide', (Get-HistorySide -ComboBox $sideCombo))
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
    $tabSummary.IsChecked     = ($Tab -eq 'Summary')
    $tabFiles.IsChecked       = ($Tab -eq 'Files')
    $tabOperations.IsChecked  = ($Tab -eq 'Operations')
    $tabDiagnostics.IsChecked = ($Tab -eq 'Diagnostics')
    $tabBatches.IsChecked     = ($Tab -eq 'Batches')
    $tabLogs.IsChecked        = ($Tab -eq 'Logs')
    $tabConfig.IsChecked      = ($Tab -eq 'Config')

    $panelSummary.Visibility     = if ($Tab -eq 'Summary')     { 'Visible' } else { 'Collapsed' }
    $panelWorkflows.Visibility   = if ($Tab -eq 'Files')       { 'Visible' } else { 'Collapsed' }
    $panelOperations.Visibility  = if ($Tab -eq 'Operations')  { 'Visible' } else { 'Collapsed' }
    $panelDiagnostics.Visibility = if ($Tab -eq 'Diagnostics') { 'Visible' } else { 'Collapsed' }
    $panelBatches.Visibility     = if ($Tab -eq 'Batches')     { 'Visible' } else { 'Collapsed' }
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
    Set-MigrationSiteButton -Button $btnOpenSourceSite -Scope $sourceScope
    Set-MigrationSiteButton -Button $btnOpenTargetSite -Scope $targetScope
    Set-MigrationSiteButton -Button $btnOpenSourceSitePerm -Scope $sourceScope
    Set-MigrationSiteButton -Button $btnOpenTargetSitePerm -Scope $targetScope
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

    $fileComparisonText = Get-ComparisonBadgeText -Folder $st.FileComparisonFolder -MigrationName $script:CurrentMigration.Name -Kind Files
    Set-Badge $badgeCmpFiles $lblCmpFilesAge $fileComparisonText ($null -ne $st.FileComparisonFolder)
    $lblCmpFilesDir.Text    = if ($st.FileComparisonFolder) { $st.FileComparisonFolder.Name } else { '' }
    $btnOpenCmpFiles.Visibility = if ($st.FileComparisonFolder) { 'Visible' } else { 'Collapsed' }
    if ($st.FileComparisonFolder) { $btnOpenCmpFiles.Tag = $st.FileComparisonFolder.FullName }
    $btnReportCmpFiles.Visibility = if ($st.FileComparisonReport) { 'Visible' } else { 'Collapsed' }
    $btnReportCmpFiles.Tag = if ($st.FileComparisonReport) { $st.FileComparisonReport.FullName } else { $null }

    $r = Format-ItemAge $st.HistoryFolder
    Set-Badge $badgeHistory $lblHistoryAge $r.Text $r.HasRun
    $btnOpenHistory.Visibility = if ($st.HistoryFolder) { 'Visible' } else { 'Collapsed' }
    if ($st.HistoryFolder) { $btnOpenHistory.Tag = $st.HistoryFolder.FullName }
    Update-HistoryScanSelection

    # --- Permissions ---
    Set-ScanComboItems $cmbScanSrcPermFile @($st.SourcePermCsvItems) $st.SourcePermCsv
    Update-ScanFileSelection $cmbScanSrcPermFile $badgeScanSrcPerm $lblScanSrcPermAge $btnOpenScanSrcPerm

    Set-ScanComboItems $cmbScanTgtPermFile @($st.TargetPermCsvItems) $st.TargetPermCsv
    Update-ScanFileSelection $cmbScanTgtPermFile $badgeScanTgtPerm $lblScanTgtPermAge $btnOpenScanTgtPerm

    $permissionComparisonText = Get-ComparisonBadgeText -Folder $st.PermComparisonFolder -MigrationName $script:CurrentMigration.Name -Kind Permissions
    Set-Badge $badgeCmpPerms $lblCmpPermsAge $permissionComparisonText ($null -ne $st.PermComparisonFolder)
    $lblCmpPermsDir.Text = if ($st.PermComparisonFolder) { $st.PermComparisonFolder.Name } else { '' }
    $btnOpenCmpPerms.Visibility = if ($st.PermComparisonFolder) { 'Visible' } else { 'Collapsed' }
    if ($st.PermComparisonFolder) { $btnOpenCmpPerms.Tag = $st.PermComparisonFolder.FullName }
    $btnReportCmpPerms.Visibility = if ($st.PermComparisonReport) { 'Visible' } else { 'Collapsed' }
    $btnReportCmpPerms.Tag = if ($st.PermComparisonReport) { $st.PermComparisonReport.FullName } else { $null }

    $r = Format-ItemAge $st.PermissionHistoryFolder
    Set-Badge $badgePermHistory $lblPermHistoryAge $r.Text $r.HasRun
    $btnOpenPermHistory.Visibility = if ($st.PermissionHistoryFolder) { 'Visible' } else { 'Collapsed' }
    if ($st.PermissionHistoryFolder) { $btnOpenPermHistory.Tag = $st.PermissionHistoryFolder.FullName }
    Update-PermissionHistoryScanSelection

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

function Invoke-ActivityLogRetention {
    $now = [datetime]::UtcNow
    if ($script:NextActivityCleanupUtc -and $now -lt $script:NextActivityCleanupUtc) { return }
    $script:NextActivityCleanupUtc = $now.AddDays(1)
    $directory = Get-SmartM365GuiActivityDirectory -ProjectRoot $script:ScriptRoot
    $cutoff = $now.AddDays(-7)
    $removed = 0
    $failed = 0
    foreach ($file in @(Get-ChildItem -LiteralPath $directory -Filter '*.log' -File -ErrorAction Stop)) {
        if ($file.Name -cnotmatch '^\d{8}-\d{6}-[0-9a-f]{32}\.log$' -or $file.LastWriteTimeUtc -ge $cutoff) { continue }
        try {
            $current = Get-Item -LiteralPath $file.FullName -ErrorAction Stop
            if ($current.LastWriteTimeUtc -ge $cutoff) { continue }
            Remove-Item -LiteralPath $current.FullName -ErrorAction Stop
            $removed++
        }
        catch {
            if (Test-Path -LiteralPath $file.FullName -PathType Leaf) { $failed++ }
        }
    }
    if ($removed -gt 0 -or $failed -gt 0) {
        $activity = New-SmartM365GuiActivity -ProjectRoot $script:ScriptRoot -Migration '<gui>' -Action 'ActivityLogRetention'
        Write-SmartM365GuiActivityEvent -Path $activity -Status $(if ($failed) { 'Partial' } else { 'Succeeded' }) `
            -ExitCode $(if ($failed) { 1 } else { 0 }) -Detail "Retention=7 days; removed=$removed; failed=$failed"
    }
}

function Format-FileInventoryVolume {
    param([long]$Bytes)
    if ($Bytes -ge 1TB) { return '{0:N2} TiB' -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return '{0:N2} GiB' -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return '{0:N1} MiB' -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return '{0:N1} KiB' -f ($Bytes / 1KB) }
    return '{0:N0} B' -f $Bytes
}

function Get-FileInventoryMetricCsvHash {
    param([System.IO.FileInfo]$File)
    if (-not (Get-Variable -Name InventoryMetricHashes -Scope Script -ErrorAction SilentlyContinue)) {
        $script:InventoryMetricHashes = @{}
    }
    $key = '{0}|{1}' -f $File.Length, $File.LastWriteTimeUtc.Ticks
    $entry = $script:InventoryMetricHashes[$File.FullName]
    if ($entry -and $entry.Key -eq $key) { return $entry.Hash }
    $hash = (Get-FileHash -LiteralPath $File.FullName -Algorithm SHA256 -ErrorAction Stop).Hash
    $script:InventoryMetricHashes[$File.FullName] = @{ Key=$key; Hash=$hash }
    return $hash
}

function Get-FileInventoryMetricPresentation {
    param($Scan)
    if (-not $Scan -or -not $Scan.File) {
        return [pscustomobject]@{ Files='—'; Folders='—'; Volume='—'; Table='—'; Evidence='No file inventory'; Tooltip='No file inventory scan is available.' }
    }
    $evidence = '{0} · {1}' -f $Scan.Date.ToString('yyyy-MM-dd HH:mm'), $Scan.Provenance
    $basis = "Latest file inventory: $($Scan.File.Name)`n$evidence`nFolders are derived from file paths and exclude empty folders. Volume sums current SizeBytes; versions and recycle bin are excluded."
    $metricsPath = "$($Scan.File.FullName).metrics.json.txt"
    if (-not (Test-Path -LiteralPath $metricsPath -PathType Leaf)) {
        return [pscustomobject]@{ Files='—'; Folders='—'; Volume='—'; Table="Metrics unavailable`nRerun scan"; Evidence="Metrics unavailable · $evidence"; Tooltip="$basis`nThis scan predates inventory metrics. Rerun the source or destination file scan." }
    }
    try {
        $Scan.File.Refresh()
        $result = Get-Content -LiteralPath $metricsPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ([int]$result.SchemaVersion -ne 1 -or $result.InventoryFile -cne $Scan.File.Name -or
            [long]$result.CsvLengthBytes -ne [long]$Scan.File.Length) {
            throw 'Metrics do not match the latest inventory CSV.'
        }
        foreach ($name in @('Rows','Files','FoldersWithFiles','KnownSizeBytes','MissingSizeFiles','MissingPathRows','InvalidLibraryRows','DuplicateRows','ConflictingSizeRows')) {
            if ($null -eq $result.$name -or [long]$result.$name -lt 0) { throw "Invalid inventory metric: $name" }
        }
        $hashProperty = $result.PSObject.Properties['CsvSha256']
        if ($hashProperty) {
            if ([string]$hashProperty.Value -notmatch '^[0-9a-fA-F]{64}$' -or
                (Get-FileInventoryMetricCsvHash -File $Scan.File) -ne [string]$hashProperty.Value) {
                throw 'Inventory CSV content differs from its scan metrics SHA256.'
            }
            $evidence = $evidence.Replace('scan receipt (hash not recalculated)', 'scan receipt') + ' · metrics CSV SHA256 verified'
            $basis = $basis.Replace('scan receipt (hash not recalculated)', 'scan receipt; metrics CSV SHA256 verified')
        }
        elseif ([long]$result.CsvLastWriteTimeUtcTicks -ne $Scan.File.LastWriteTimeUtc.Ticks) {
            # Synced copies can truncate subsecond timestamps. Require the scan receipt's content hash.
            $recordedSecond = [long]$result.CsvLastWriteTimeUtcTicks - ([long]$result.CsvLastWriteTimeUtcTicks % [TimeSpan]::TicksPerSecond)
            if ($Scan.File.LastWriteTimeUtc.Ticks -ne $recordedSecond) { throw 'Metrics CSV timestamp does not match this scan.' }
            $receipt = Get-Content -LiteralPath ($Scan.File.FullName + '.manifest.json.txt') -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ([int]$receipt.SchemaVersion -ne 1 -or $receipt.InventoryFile -cne $Scan.File.Name -or
                [long]$receipt.Rows -ne [long]$result.Rows -or [string]$receipt.Sha256 -notmatch '^[0-9a-fA-F]{64}$' -or
                (Get-FileInventoryMetricCsvHash -File $Scan.File) -ne [string]$receipt.Sha256) {
                throw 'Synchronized inventory CSV cannot be verified against its scan receipt.'
            }
            $evidence = $evidence.Replace('scan receipt (hash not recalculated)', 'scan receipt') + ' · metrics receipt SHA256 verified'
            $basis = $basis.Replace('scan receipt (hash not recalculated)', 'scan receipt; metrics CSV SHA256 verified')
        }
    }
    catch {
        return [pscustomobject]@{ Files='—'; Folders='—'; Volume='—'; Table='Metrics unavailable'; Evidence="Metrics unavailable · $evidence"; Tooltip="$basis`n$($_.Exception.Message)" }
    }
    $pathsComplete = [int]$result.missingPathRows -eq 0
    $foldersComplete = $pathsComplete -and [int]$result.invalidLibraryRows -eq 0
    $volumeComplete = $pathsComplete -and [int]$result.missingSizeFiles -eq 0 -and [int]$result.conflictingSizeRows -eq 0
    $files = if ($pathsComplete) { '{0:N0}' -f [long]$result.files } else { '—' }
    $folders = if ($foldersComplete) { '{0:N0}' -f [long]$result.foldersWithFiles } else { '—' }
    $volume = if ($volumeComplete) { Format-FileInventoryVolume -Bytes ([long]$result.knownSizeBytes) } else { '—' }
    $quality = "Rows: $($result.rows); duplicate paths: $($result.duplicateRows); missing paths: $($result.missingPathRows); invalid library paths: $($result.invalidLibraryRows); missing sizes: $($result.missingSizeFiles); conflicting sizes: $($result.conflictingSizeRows)."
    return [pscustomobject]@{
        Files=$files; Folders=$folders; Volume=$volume
        Table="Files $files · folders $folders`nVolume $volume"
        Evidence=$evidence
        Tooltip="$basis`n$quality"
    }
}

function Update-PortfolioInventoryRow {
    param($Row)
    foreach ($side in @('Source','Target')) {
        $file = $Row.("${side}FileScanFile")
        $scan = if ($file) { [pscustomobject]@{
            File=$file; Date=$Row.("${side}FileScanDate"); Provenance=$Row.("${side}FileScanProvenance")
        } } else { $null }
        $value = Get-FileInventoryMetricPresentation -Scan $scan
        $Row.("${side}InventoryDisplay") = $value.Table
        $Row.("${side}InventoryTooltip") = $value.Tooltip
    }
}

function Refresh-DiagnosticInventoryMetrics {
    if (-not $script:CurrentMigration) { return }
    $migration = $script:CurrentMigration
    $cfg = $migration.Config
    foreach ($spec in @(
        @{ Side='Source'; Type=(Get-MigrationEndpointType $cfg 'Source'); Folder=$cfg.Output.SourceFileScans; Files=$lblDiagSourceFiles; Folders=$lblDiagSourceFolders; Volume=$lblDiagSourceVolume; Evidence=$lblDiagSourceInventoryEvidence },
        @{ Side='Target'; Type=(Get-MigrationEndpointType $cfg 'Target'); Folder=$cfg.Output.TargetFileScans; Files=$lblDiagTargetFiles; Folders=$lblDiagTargetFolders; Volume=$lblDiagTargetVolume; Evidence=$lblDiagTargetInventoryEvidence }
    )) {
        $directory = Join-Path $migration.Root $spec.Folder
        $scan = Get-SmartM365LatestPortfolioScan -Directory $directory -Filter ("{0}-FileInventory-{1}-*.csv" -f $spec.Type, $migration.Name)
        $value = Get-FileInventoryMetricPresentation -Scan $scan
        $spec.Files.Text = $value.Files
        $spec.Folders.Text = $value.Folders
        $spec.Volume.Text = $value.Volume
        $spec.Evidence.Text = $value.Evidence
        $spec.Evidence.ToolTip = $value.Tooltip
    }
}

function Refresh-PortfolioSummary {
    $rows = [System.Collections.Generic.List[object]]::new()
    if ($null -eq $script:SummaryLastGoodRows) { $script:SummaryLastGoodRows = @{} }
    if ($null -eq $script:SummaryLastErrorMessages) { $script:SummaryLastErrorMessages = @{} }
    $errors = [System.Collections.Generic.List[string]]::new()
    foreach ($migration in $script:Migrations) {
        try {
            $source = Get-MigrationScope $migration 'Source'
            $target = Get-MigrationScope $migration 'Target'
            $row = Get-SmartM365PortfolioRow -Migration $migration `
                -SourceScope $source.Text -SourceTooltip $source.Tooltip `
                -TargetScope $target.Text -TargetTooltip $target.Tooltip `
                -SourceType (Get-MigrationEndpointType $migration.Config 'Source') `
                -TargetType (Get-MigrationEndpointType $migration.Config 'Target')
            Update-PortfolioInventoryRow -Row $row
            $rows.Add($row)
            $script:SummaryLastGoodRows[$migration.Name] = $row
            [void]$script:SummaryLastErrorMessages.Remove($migration.Name)
        }
        catch {
            $message = [string]$_.Exception.Message
            $errors.Add(('{0}: {1}' -f $migration.Name, $message))
            if ($script:SummaryLastErrorMessages[$migration.Name] -ne $message) {
                try {
                    $activity = New-SmartM365GuiActivity -ProjectRoot $script:ScriptRoot `
                        -Migration $migration.Name -Action 'PortfolioSummaryRefresh'
                    Write-SmartM365GuiActivityEvent -Path $activity -Status 'Failed' `
                        -ExitCode 1 -Detail $message -Migration $migration.Name
                }
                catch { }
            }
            $script:SummaryLastErrorMessages[$migration.Name] = $message
            if ($script:SummaryLastGoodRows.ContainsKey($migration.Name)) {
                $cached = $script:SummaryLastGoodRows[$migration.Name] | Select-Object *
                $cached.Status = 'Refresh error'
                $cached.StatusTooltip = "Refresh failed; showing last successful values.`n$message"
                $rows.Add($cached)
            }
            else {
                $rows.Add([pscustomobject]@{
                    Migration = $migration.Name; Source = '—'; SourceTooltip = ''
                    SourceScan = '—'; SourceScanTooltip = ''
                    SourceInventoryDisplay = '—'; SourceInventoryTooltip = $message
                    SourcePermissionScan = '—'; SourceScansDisplay = "Files —`nPerms —"
                    SourceScansSortDate = [datetime]::MinValue; SourceScansTooltip = $message
                    Destination = '—'; DestinationTooltip = ''
                    TargetScan = '—'; TargetScanTooltip = ''
                    TargetInventoryDisplay = '—'; TargetInventoryTooltip = $message
                    TargetPermissionScan = '—'; TargetScansDisplay = "Files —`nPerms —"
                    TargetScansSortDate = [datetime]::MinValue; TargetScansTooltip = $message
                    ScanGapDays = $null; ScanGapText = '—'; ScanGapTooltip = $message
                    ScanGapVisual = Get-SmartM365PortfolioGapVisual -Days $null
                    ComparisonRate = $null; ComparisonPercent = '—'
                    ComparisonDate = '—'; ComparisonDisplay = '—'; ComparisonTooltip = $message
                    PermissionComparisonRate = $null; PermissionComparisonPercent = '—'
                    PermissionComparisonDate = '—'; PermissionComparisonDisplay = '—'; PermissionComparisonTooltip = $message
                    ComparisonVisual = Get-SmartM365PortfolioRateVisual -Rate $null -Notice 'Refresh error'
                    PermissionComparisonVisual = Get-SmartM365PortfolioRateVisual -Rate $null -Notice 'Refresh error'
                    GlobalComparisonRate = $null; GlobalComparisonTooltip = $message
                    GlobalComparisonVisual = Get-SmartM365PortfolioRateVisual -Rate $null -Notice 'Refresh error' -Caption 'Equal weighting'
                    Status = 'Refresh error'; StatusTooltip = $message
                })
            }
        }
    }
    $displayedRows = $rows.ToArray()
    $fileCount = @($displayedRows | Where-Object { $_.ComparisonDate -ne '—' }).Count
    $permissionCount = @($displayedRows | Where-Object { $_.PermissionComparisonDate -ne '—' }).Count
    $scanCount = @($displayedRows | Where-Object { $_.Status -eq 'Scan needed' }).Count
    $compareCount = @($displayedRows | Where-Object { $_.Status -eq 'Compare needed' }).Count
    $reviewCount = @($displayedRows | Where-Object { $_.Status -eq 'Review needed' }).Count
    $actionCount = $scanCount + $compareCount + $reviewCount
    $lblPortfolioMigrations.Text = [string]$displayedRows.Count
    $lblPortfolioFilesCompared.Text = '{0} / {1}' -f $fileCount, $displayedRows.Count
    $lblPortfolioPermissionsCompared.Text = '{0} / {1}' -f $permissionCount, $displayedRows.Count
    $lblPortfolioActions.Text = [string]$actionCount
    $lblPortfolioActions.ToolTip = 'Scan needed: {0}; Compare needed: {1}; Review needed: {2}' -f $scanCount, $compareCount, $reviewCount
    $lblPortfolioRefreshErrors.Text = [string]$errors.Count
    $lblPortfolioRefreshErrors.ToolTip = if ($errors.Count) { $errors -join "`n" } else { 'No refresh errors.' }
    $script:SummaryLoading = $true
    try { $gridSummary.ItemsSource = $displayedRows }
    finally { $script:SummaryLoading = $false }
    $lblSummaryStatus.Text = ('{0} migrations · {1} refresh errors · Updated {2}' -f
        $rows.Count, $errors.Count, (Get-Date -Format 'HH:mm:ss'))
    $lblSummaryStatus.ToolTip = if ($errors.Count) { $errors -join "`n" } else { $null }
}

function Refresh-GuiState {
    try {
        if (-not (Test-Path -LiteralPath (Join-Path $script:ScriptRoot 'Migrations') -PathType Container)) {
            throw 'Migration folder is unavailable.'
        }
        Load-Migrations
        if ($tabSummary.IsChecked) { Refresh-PortfolioSummary }
        if ($script:CurrentMigration -and $tabDiagnostics.IsChecked) { Refresh-DiagnosticInventoryMetrics }
        Refresh-TransientResults
        try { Invoke-ActivityLogRetention }
        catch {
            Microsoft.PowerShell.Utility\Write-Warning ('{0} Activity log retention failed: {1}' -f
                (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$_.Exception.Message)
        }
        Refresh-ActivityList
        $lblLastRefresh.Text = 'Updated ' + (Get-Date -Format 'HH:mm:ss')
        $lblLastRefresh.ToolTip = 'Shared activity and migration state refreshed.'
        $script:LastGlobalRefreshError = ''
    }
    catch {
        $refreshError = $_
        $message = [string]$_.Exception.Message
        $lblLastRefresh.Text = 'Refresh failed'
        $lblLastRefresh.ToolTip = $message
        if ($tabSummary.IsChecked -and $null -eq $gridSummary.ItemsSource) {
            $lblSummaryStatus.Text = "Overview refresh failed: $message"
            $lblSummaryStatus.ToolTip = $message
        }
        if ($script:LastGlobalRefreshError -ne $message) {
            try {
                $activity = New-SmartM365GuiActivity -ProjectRoot $script:ScriptRoot `
                    -Migration '<gui>' -Action 'GuiRefresh'
                Write-SmartM365GuiActivityEvent -Path $activity -Status 'Failed' `
                    -ExitCode 1 -Detail ($message + ' | ' + $refreshError.ScriptStackTrace)
            }
            catch { }
        }
        $script:LastGlobalRefreshError = $message
    }
}

function Get-DiagnosticReportDisplayPath {
    param([string]$Folder)
    if (-not $Folder) { return 'SharePointMigration\Migrations' }
    $relative = [IO.Path]::GetRelativePath($script:ScriptRoot, $Folder).Replace('/', '\')
    if ($relative -eq '..' -or $relative.StartsWith('..\', [StringComparison]::Ordinal) -or
        [IO.Path]::IsPathRooted($relative)) { return 'Report folder is outside SharePointMigration.' }
    return 'SharePointMigration\' + $relative
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
        $lblDiagInputPath.Text = Get-DiagnosticReportDisplayPath -Folder $script:DiagInputPath
        $script:DiagReportSignature = ''
        Clear-DiagnosticResult
    }
    $script:CurrentStatus    = Get-MigrationStatus -Migration $Migration
    Refresh-DiagnosticInventoryMetrics
    Refresh-DiagnosticReportState
    Refresh-TransientResults
    if (-not $sameConfig) { Refresh-FarmDiagnostics }
    Update-UI
    Refresh-DiagnosticCrossCheck
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
        Urls = @($urls.ToArray())
        Tooltip = (($urls.ToArray() -join "`n") + $(if ($mismatch) { "`nConfig SiteUrl differs: $configured" } else { '' }))
        Mismatch = $mismatch
    }
}

function Set-MigrationSiteButton {
    param($Button, $Scope)
    $valid = @($Scope.Urls | Where-Object {
        $uri = $null
        [Uri]::TryCreate([string]$_, [UriKind]::Absolute, [ref]$uri) -and $uri.Scheme -in @('http','https')
    } | Select-Object -Unique)
    $Button.Tag = $valid
    $Button.IsEnabled = $valid.Count -gt 0
    $Button.ToolTip = if ($valid.Count) { $valid -join "`n" } else { 'No valid site URL is configured for this migration.' }
}

function Open-MigrationSiteUrl {
    param([string]$Url)
    try { Start-Process -FilePath $Url -ErrorAction Stop }
    catch {
        [void][System.Windows.MessageBox]::Show("Could not open the site in your browser:`n$($_.Exception.Message)", $script:AppName, 'OK', 'Error')
    }
}

function Open-MigrationSite {
    param($Button)
    $urls = @($Button.Tag)
    if (-not $Button.IsEnabled -or $urls.Count -eq 0) { return }
    if ($urls.Count -eq 1) { Open-MigrationSiteUrl -Url $urls[0]; return }
    $menu = [System.Windows.Controls.ContextMenu]::new()
    foreach ($url in $urls) {
        $item = [System.Windows.Controls.MenuItem]::new()
        $item.Header = $url; $item.Tag = $url
        $item.Add_Click({ param($sender, $eventArgs) Open-MigrationSiteUrl -Url ([string]$sender.Tag) })
        [void]$menu.Items.Add($item)
    }
    $Button.ContextMenu = $menu
    $menu.PlacementTarget = $Button; $menu.Placement = 'Bottom'; $menu.IsOpen = $true
}

function Load-Migrations {
    $prev = if ($cmbMigration.SelectedItem) { [string]$cmbMigration.SelectedItem } else { $null }
    $script:Migrations = @(Get-MigrationFolders)
    Sync-SmartM365BatchMigrations -Names @($script:Migrations | ForEach-Object Name)
    $cmbMigration.Items.Clear()
    foreach ($m in $script:Migrations) { [void]$cmbMigration.Items.Add($m.Name) }
    if ($script:Migrations.Count -eq 0) { return }
    $idx = if ($prev) { $cmbMigration.Items.IndexOf($prev) } else { -1 }
    $cmbMigration.SelectedIndex = if ($idx -ge 0) { $idx } else { 0 }
}

# ---------------------------------------------------------------------------
# Events
# ---------------------------------------------------------------------------

$tabSummary.Add_Click({     Switch-Tab 'Summary'; Refresh-PortfolioSummary })
$tabFiles.Add_Click({       Switch-Tab 'Files' })
$tabOperations.Add_Click({  Switch-Tab 'Operations' })
$tabDiagnostics.Add_Click({ Switch-Tab 'Diagnostics'; Refresh-DiagnosticInventoryMetrics; Refresh-DiagnosticCrossCheck })
$tabBatches.Add_Click({ Switch-Tab 'Batches'; Refresh-SmartM365BatchGui -Force })
$tabLogs.Add_Click({        Switch-Tab 'Logs' })
$tabConfig.Add_Click({      Switch-Tab 'Config' })

$gridSummary.Add_SelectionChanged({
    if ($script:SummaryLoading -or -not $gridSummary.SelectedItem) { return }
    $migrationName = [string]$gridSummary.SelectedItem.Migration
    if ($cmbMigration.Items.Contains($migrationName)) {
        $cmbMigration.SelectedItem = $migrationName
        if ($script:CurrentMigration -and $script:CurrentMigration.Name -eq $migrationName) {
            Switch-Tab 'Files'
        }
    }
})

$cmbMigration.Add_SelectionChanged({
    $idx = $cmbMigration.SelectedIndex
    if ($idx -ge 0 -and $idx -lt $script:Migrations.Count) {
        $requestedMigration = $script:Migrations[$idx]
        $previousMigration = $script:CurrentMigration
        $previousStatus = $script:CurrentStatus
        try {
            Set-CurrentMigration -Migration $requestedMigration
        }
        catch {
            $selectionError = $_
            $script:CurrentMigration = $previousMigration
            $script:CurrentStatus = $previousStatus
            try {
                $activity = New-SmartM365GuiActivity -ProjectRoot $script:ScriptRoot `
                    -Migration $requestedMigration.Name -Action 'SelectMigration'
                Write-SmartM365GuiActivityEvent -Path $activity -Status 'Failed' -ExitCode 1 `
                    -Detail ($selectionError.Exception.Message + ' | ' + $selectionError.ScriptStackTrace)
            }
            catch { [void]$_.Exception }
            $lblLastRefresh.Text = 'Migration load failed'
            $lblLastRefresh.ToolTip = $selectionError.Exception.Message
            [System.Windows.MessageBox]::Show(
                "Could not load migration '$($requestedMigration.Name)':`n$($selectionError.Exception.Message)",
                $script:AppName, 'OK', 'Error') | Out-Null
            if ($previousMigration) { $cmbMigration.SelectedItem = $previousMigration.Name }
        }
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
$btnOpenSourceSite.Add_Click({ Open-MigrationSite -Button $btnOpenSourceSite })
$btnOpenTargetSite.Add_Click({ Open-MigrationSite -Button $btnOpenTargetSite })
$btnOpenSourceSitePerm.Add_Click({ Open-MigrationSite -Button $btnOpenSourceSitePerm })
$btnOpenTargetSitePerm.Add_Click({ Open-MigrationSite -Button $btnOpenTargetSitePerm })
$btnOpenScanTgt.Add_Click({  Open-InExplorer ([string]$btnOpenScanTgt.Tag) })
$btnOpenCmpFiles.Add_Click({ Open-InExplorer ([string]$btnOpenCmpFiles.Tag) })
$btnReportCmpFiles.Add_Click({ Open-InExplorer ([string]$btnReportCmpFiles.Tag) })
$btnOpenHistory.Add_Click({  Open-InExplorer ([string]$btnOpenHistory.Tag) })

# Permissions
$btnRunScanSrcPerm.Add_Click({  Invoke-MigrationAction 'ScanSourcePermissions' })
$btnRunScanTgtPerm.Add_Click({  Invoke-MigrationAction 'ScanTargetPermissions' })
$btnRunCmpPerms.Add_Click({     Invoke-MigrationAction 'ComparePermissions' })
$btnRunPermHistory.Add_Click({ Invoke-MigrationAction 'ComparePermissionScanHistory' })

$cmbScanSrcPermFile.Add_SelectionChanged({ Update-ScanFileSelection $cmbScanSrcPermFile $badgeScanSrcPerm $lblScanSrcPermAge $btnOpenScanSrcPerm })
$cmbScanTgtPermFile.Add_SelectionChanged({ Update-ScanFileSelection $cmbScanTgtPermFile $badgeScanTgtPerm $lblScanTgtPermAge $btnOpenScanTgtPerm })
$cmbPermHistorySide.Add_SelectionChanged({ Update-PermissionHistoryScanSelection })
$cmbPermHistoryOldFile.Add_SelectionChanged({ Update-PermissionHistoryRunState })
$cmbPermHistoryNewFile.Add_SelectionChanged({ Update-PermissionHistoryRunState })

$btnOpenScanSrcPerm.Add_Click({ Open-InExplorer ([string]$btnOpenScanSrcPerm.Tag) })
$btnOpenScanTgtPerm.Add_Click({ Open-InExplorer ([string]$btnOpenScanTgtPerm.Tag) })
$btnOpenCmpPerms.Add_Click({    Open-InExplorer ([string]$btnOpenCmpPerms.Tag) })
$btnReportCmpPerms.Add_Click({ Open-InExplorer ([string]$btnReportCmpPerms.Tag) })
$btnOpenPermHistory.Add_Click({ Open-InExplorer ([string]$btnOpenPermHistory.Tag) })

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
        $reportPath = [string]($result |
            ForEach-Object { ([string]$_).Trim() } |
            Where-Object { $_.EndsWith('.html', [System.StringComparison]::OrdinalIgnoreCase) } |
            Select-Object -Last 1)
        if (-not $reportPath) { throw "Generator did not return an HTML report path. Output:`n$($result -join "`n")" }
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
$btnOverviewGlobalFileReport.Add_Click({ Invoke-GlobalComparisonReport -Kind 'Files' })
$btnOverviewGlobalPermissionsReport.Add_Click({ Invoke-GlobalComparisonReport -Kind 'Permissions' })

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
    if (-not $pattern) { $gridDiagRows.ItemsSource = $null; $txtDiagRaw.Text = ''; return }
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

function Clear-DiagnosticResult {
    $script:DiagSummary = $null
    $script:DiagAnalysisVerified = $false
    $script:DiagLoadedDirectory = ''
    $script:DiagRows = @()
    $script:DiagKnownSessions = @()
    $gridDiagPatterns.ItemsSource = $null
    $gridDiagRows.ItemsSource = $null
    $txtDiagRaw.Text = ''
    $lblDiagKpis.Text = 'No analysis for the latest report yet.'
    $lblDiagKpis.Visibility = 'Visible'
    $panelDiagMetrics.Visibility = 'Collapsed'
    $lblDiagReportState.Text = 'Checking'
    $lblDiagReportEvidence.Text = 'Latest report'
    $lblDiagAnalysisState.Text = 'Checking'
    $lblDiagAnalysisEvidence.Text = 'Matching analysis'
    $lblDiagHtmlState.Text = 'Checking'
    $lblDiagHtmlEvidence.Text = 'Latest HTML'
    $lblDiagNextAction.Text = 'Checking the selected migration report.'
    $lblSummaryShareGateValue.Text = '—'
    $lblSummaryShareGateDetail.Text = 'Analyze the latest report'
    $lblSummaryShareGateDetail.ToolTip = $null
    $lblSummaryFilesValue.Text = '—'
    $lblSummaryFilesDetail.Text = 'Loading comparison'
    $lblSummaryFilesDetail.ToolTip = $null
    $lblSummaryPermissionsValue.Text = '—'
    $lblSummaryPermissionsDetail.Text = 'Loading comparison'
    $lblSummaryPermissionsDetail.ToolTip = $null
    $lblSummaryCrossCheckValue.Text = '—'
    $lblSummaryCrossCheckDetail.Text = 'Loading cross-check'
    $lblSummaryCrossCheckDetail.ToolTip = $null
    $btnDiagOpenReport.IsEnabled = $false
    $panelDiagReview.IsEnabled = $false
    $script:FarmInvocation = $null
    $script:FarmPrerequisitesPassed = $false
    $btnFarmCheck.IsEnabled = $false
    $btnFarmRun.IsEnabled = $false
    $txtFarmDryRun.Text = ''
    $txtFarmRun.Text = ''
    $script:DiagLoading = $true
    try {
        $cmbDiagSession.Items.Clear()
        [void]$cmbDiagSession.Items.Add('All sessions')
        $cmbDiagSession.SelectedIndex = 0
    }
    finally { $script:DiagLoading = $false }
}

function Update-DiagnosticReportStatus {
    param([string]$Failure = '')

    $report = $script:DiagLatestReport
    $isRunning = [bool]($script:DiagProcess -and $script:DiagActiveReportKey -eq (Get-DiagnosticReportKey))
    $panelDiagAnalysisLoading.Visibility = if ($isRunning -and -not $Failure) { 'Visible' } else { 'Collapsed' }
    $hasHtml = [bool]($script:DiagSummary -and $btnDiagOpenReport.IsEnabled)
    if (-not $script:DiagSummary) {
        $lblDiagKpis.Text = if (-not $report) { 'No ShareGate migration report is available. Deposit the latest CSV or XLSX in MigrationReport.' }
            elseif ($Failure) { 'The ShareGate analysis failed. Review the error above, then retry Analyze latest report.' }
            elseif ($isRunning) { 'Analyzing ShareGate report: ' + $report.Name + '. Indicators will appear when analysis finishes.' }
            elseif ($script:DiagProcess) { 'Report queued for automatic analysis: ' + $report.Name + '. Indicators will appear when analysis finishes.' }
            else { 'ShareGate report detected: ' + $report.Name + '. Its analysis will start automatically; indicators and HTML will appear when it finishes.' }
    }
    if (-not $report) {
        $lblDiagReportState.Text = 'Missing'
        $lblDiagReportEvidence.Text = 'CSV or XLSX required'
        $lblDiagAnalysisState.Text = 'Unavailable'
        $lblDiagAnalysisEvidence.Text = 'No report to analyze'
        $lblDiagHtmlState.Text = 'Unavailable'
        $lblDiagHtmlEvidence.Text = 'No analysis HTML'
        $lblDiagNextAction.Text = 'Place the latest ShareGate report in MigrationReport, then click Refresh reports.'
        return
    }

    $lblDiagReportState.Text = 'Detected'
    $lblDiagReportEvidence.Text = '{0} · {1}' -f $report.Extension.TrimStart('.').ToUpperInvariant(), $report.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
    $lblDiagHtmlState.Text = if ($hasHtml) { 'Available' } else { 'Unavailable' }
    $lblDiagHtmlEvidence.Text = if ($hasHtml) { 'Open analysis HTML' } else { 'No HTML for latest report' }

    if ($Failure) {
        $lblDiagAnalysisState.Text = 'Failed'
        $lblDiagAnalysisEvidence.Text = 'See analysis status above'
        $lblDiagNextAction.Text = 'Review the analysis error above, then retry Analyze latest report.'
    }
    elseif ($isRunning) {
        $lblDiagAnalysisState.Text = 'Running'
        $lblDiagAnalysisEvidence.Text = 'Processing latest report'
        $lblDiagNextAction.Text = 'Wait for the analysis to finish.'
    }
    elseif ($script:DiagProcess -and -not $script:DiagSummary) {
        $lblDiagAnalysisState.Text = 'Queued'
        $lblDiagAnalysisEvidence.Text = 'Another analysis is finishing'
        $lblDiagNextAction.Text = 'The latest report will be analyzed automatically when the current analysis finishes.'
    }
    elseif ($script:DiagSummary) {
        $lblDiagAnalysisState.Text = if ($script:DiagAnalysisVerified) { 'SHA256 verified' } else { 'Legacy match' }
        $lblDiagAnalysisEvidence.Text = if ($script:DiagAnalysisVerified) { 'Exact report content matched' } else { 'Path and timestamp only' }
        if ($script:DiagAnalysisVerified) {
            $lblDiagNextAction.Text = 'Review issue patterns and the cross-check below.'
        }
        elseif ($report.Extension -eq '.xlsx') {
            $lblDiagNextAction.Text = 'Reanalyze to verify this report by SHA256. ImportExcel installs automatically if needed.'
        }
        else {
            $lblDiagNextAction.Text = 'Reanalyze this report to verify it by SHA256, then review the findings.'
        }
    }
    elseif ($report.Extension -eq '.xlsx') {
        $lblDiagAnalysisState.Text = 'Ready to analyze'
        $lblDiagAnalysisEvidence.Text = 'XLSX support prepared automatically'
        $lblDiagNextAction.Text = 'The latest report will be analyzed automatically. ImportExcel installs automatically for the current user if needed.'
    }
    else {
        $lblDiagAnalysisState.Text = 'Not analyzed'
        $lblDiagAnalysisEvidence.Text = 'No matching analysis'
        $lblDiagNextAction.Text = 'The latest report will be analyzed automatically to create its summary and HTML report.'
    }
}

function Update-DiagnosticSummaryCrossCheck {
    param($Result)
    foreach ($spec in @(
        @{ Evidence='Files'; Value=$lblSummaryFilesValue; Detail=$lblSummaryFilesDetail; Path='FilesSummary' },
        @{ Evidence='Permissions'; Value=$lblSummaryPermissionsValue; Detail=$lblSummaryPermissionsDetail; Path='PermissionsSummary' }
    )) {
        $entry = @($Result.Evidence | Where-Object { $_.Evidence -eq $spec.Evidence } | Select-Object -First 1)
        if ($entry.Count -and $null -ne $entry[0] -and $entry[0].Rate -ne '—') {
            $spec.Value.Text = [string]$entry[0].Rate
            $spec.Detail.Text = "$($entry[0].Coverage) · $($entry[0].State)"
        }
        else {
            $spec.Value.Text = '—'
            $spec.Detail.Text = if ($entry.Count -and $null -ne $entry[0]) { [string]$entry[0].State } else { 'Comparison unavailable' }
        }
        $spec.Detail.ToolTip = [string]$Result.($spec.Path)
    }
    $hasEvidence = @($Result.Evidence | Where-Object {
        ($_.Evidence -eq 'ShareGate' -and $_.State -ne 'Analysis missing') -or
        ($_.Evidence -in @('Files','Permissions') -and $_.Rate -ne '—')
    }).Count -gt 0
    $lblSummaryCrossCheckValue.Text = if ($hasEvidence) { [string](@($Result.Scopes).Count) } else { '—' }
    $lblSummaryCrossCheckDetail.Text = if ($hasEvidence) { "Scopes with differences · $($Result.Ambiguous) ambiguous" } else { 'Evidence unavailable' }
    $lblSummaryCrossCheckDetail.ToolTip = 'Matched site and list scopes with at least one reported difference; not a count of individual files or permissions.'
}

function Get-DiagnosticCrossCheckSignature {
    if (-not $script:CurrentMigration -or -not $script:CurrentStatus) { return '' }
    $parts = [System.Collections.Generic.List[string]]::new()
    $parts.Add([string]$script:CurrentMigration.Root)
    $parts.Add([string]$script:DiagLoadedDirectory)
    $parts.Add([string]$script:DiagAnalysisVerified)
    foreach ($name in @('SourceFileCsv','TargetFileCsv','SourcePermCsv','TargetPermCsv','FileComparisonFolder','PermComparisonFolder')) {
        $item = $script:CurrentStatus.PSObject.Properties[$name].Value
        $parts.Add($(if ($item) { [string]$item.FullName } else { '' }))
        $parts.Add($(if ($item) { [string]$item.LastWriteTimeUtc.Ticks } else { '' }))
    }
    return $parts -join '|'
}

function Start-DiagnosticCrossCheckJob {
    param([string]$Signature)
    $context = [pscustomobject]@{
        Migration=$script:CurrentMigration; Status=$script:CurrentStatus
        Summary=$script:DiagSummary; Rows=@($script:DiagRows)
        Verified=[bool]$script:DiagAnalysisVerified
    }
    $summaryScript = Join-Path $script:ScriptRoot 'SmartM365-SharePointMigration-Summary.ps1'
    $crossCheckScript = Join-Path $script:ScriptRoot 'Scripts\Diagnostics\SmartM365-SharePointMigration-CrossCheck.ps1'
    $script:CrossCheckJob = Start-ThreadJob -ScriptBlock {
        param($InputContext,$SummaryScript,$CrossCheckScript)
        . $SummaryScript
        . $CrossCheckScript
        Get-SmartM365DiagnosticCrossCheck -Migration $InputContext.Migration -Status $InputContext.Status `
            -DiagnosticSummary $InputContext.Summary -DiagnosticRows @($InputContext.Rows) `
            -DiagnosticVerified ([bool]$InputContext.Verified)
    } -ArgumentList $context,$summaryScript,$crossCheckScript -ErrorAction Stop
    $script:CrossCheckJobSignature = $Signature
    $script:CrossCheckTimer.Start()
}

function Refresh-DiagnosticCrossCheck {
    param([switch]$Force)
    if (-not $tabDiagnostics.IsChecked -and -not $Force) { return }
    $signature = Get-DiagnosticCrossCheckSignature
    if (-not $Force -and $signature -and $signature -eq $script:CrossCheckSignature) {
        if ($script:CrossCheckResult) { Update-DiagnosticSummaryCrossCheck $script:CrossCheckResult }
        return
    }
    if (-not $Force -and $script:CrossCheckJob -and $signature -eq $script:CrossCheckJobSignature) { return }
    $script:CrossCheckRequestedSignature = $signature
    $gridCrossCheckEvidence.ItemsSource = $null
    $gridCrossCheckScopes.ItemsSource = $null
    $btnCrossCheckFilesReport.IsEnabled = $false
    $btnCrossCheckPermissionsReport.IsEnabled = $false
    $btnCrossCheckFilesReport.Tag = $null
    $btnCrossCheckPermissionsReport.Tag = $null
    $script:CrossCheckResult = $null
    $lblSummaryFilesValue.Text = '—'
    $lblSummaryFilesDetail.Text = 'Loading comparison'
    $lblSummaryPermissionsValue.Text = '—'
    $lblSummaryPermissionsDetail.Text = 'Loading comparison'
    $lblSummaryCrossCheckValue.Text = '—'
    $lblSummaryCrossCheckDetail.Text = 'Loading cross-check'
    if (-not $script:CurrentMigration -or -not $script:CurrentStatus) {
        $panelCrossCheckLoading.Visibility = 'Collapsed'
        $lblCrossCheckStatus.Text = 'Select a migration to compare existing reports.'
        $lblSummaryFilesDetail.Text = 'Select a migration'
        $lblSummaryPermissionsDetail.Text = 'Select a migration'
        $lblSummaryCrossCheckDetail.Text = 'Select a migration'
        return
    }
    $script:CrossCheckLoadingStarted = Get-Date
    $panelCrossCheckLoading.Visibility = 'Visible'
    $lblCrossCheckLoading.Text = "Loading comparison reports for $($script:CurrentMigration.Name)…"
    $lblCrossCheckStatus.Text = if ($script:CrossCheckJob) { 'Waiting for the previous cross-check, then loading the selected migration…' }
        else { 'Loading existing comparison reports…' }
    if ($script:CrossCheckJob) { return }
    try {
        Start-DiagnosticCrossCheckJob -Signature $signature
    }
    catch {
        $panelCrossCheckLoading.Visibility = 'Collapsed'
        $lblCrossCheckStatus.Text = 'Cross-check unavailable: ' + $_.Exception.Message
        $lblSummaryFilesDetail.Text = 'Cross-check unavailable'
        $lblSummaryPermissionsDetail.Text = 'Cross-check unavailable'
        $lblSummaryCrossCheckDetail.Text = 'Cross-check unavailable'
    }
}

function Get-CurrentDiagnosticAnalysis {
    param([System.IO.FileInfo]$Report)
    if (-not $Report -or -not $script:CurrentMigration) { return $null }
    $root = Join-Path $script:CurrentMigration.Root 'ShareGate\Diagnostics'
    if (-not (Test-Path -LiteralPath $root -PathType Container)) { return $null }
    $dirs = @(Get-ChildItem -LiteralPath $root -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending)
    foreach ($dir in $dirs) {
        $summaryPath = Join-Path $dir.FullName 'Summary.json.txt'
        if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf)) { continue }
        try {
            $summaryJson = Get-Content -LiteralPath $summaryPath -Raw -ErrorAction Stop
            $summary = $summaryJson | ConvertFrom-Json -AsHashtable
            $jsonDocument = [System.Text.Json.JsonDocument]::Parse($summaryJson)
            try { $generatedAtUtc = $jsonDocument.RootElement.GetProperty('GeneratedAtUtc').GetString() }
            finally { $jsonDocument.Dispose() }
            if ([string]$summary.Project -ne [string]$script:CurrentMigration.Name) { continue }
            if (@($summary.Inputs).Count -ne 1) { continue }
            $recordedInput = [string]$summary.Inputs[0]
            if (-not [string]::Equals($recordedInput, $Report.FullName, [StringComparison]::OrdinalIgnoreCase) -and
                -not [string]::Equals([IO.Path]::GetFileName($recordedInput), $Report.Name, [StringComparison]::OrdinalIgnoreCase)) { continue }
            if (-not (Test-Path -LiteralPath (Join-Path $dir.FullName 'MigrationDiagnostics-Report.html') -PathType Leaf) -or
                -not (Test-Path -LiteralPath (Join-Path $dir.FullName 'ClassifiedRows.csv') -PathType Leaf)) { continue }
            $evidence = @($summary['InputEvidence'])
            if ($evidence.Count -eq 1 -and $evidence[0]) {
                if ([int64]$evidence[0].Size -ne $Report.Length -or
                    -not [string]::Equals([IO.Path]::GetFileName([string]$evidence[0].Path), $Report.Name, [StringComparison]::OrdinalIgnoreCase)) { continue }
                if (-not $script:DiagReportHash) { $script:DiagReportHash = (Get-FileHash -LiteralPath $Report.FullName -Algorithm SHA256).Hash }
                if (-not [string]::Equals([string]$evidence[0].Sha256, $script:DiagReportHash, [StringComparison]::OrdinalIgnoreCase)) { continue }
                return [pscustomobject]@{ Directory=$dir.FullName; Verified=$true; Session=[string]$summary.SelectedSessionId; Generated=$generatedAtUtc }
            }
            $generated = [datetimeoffset]::MinValue
            if ([datetimeoffset]::TryParse($generatedAtUtc, [Globalization.CultureInfo]::InvariantCulture,
                    [Globalization.DateTimeStyles]::None, [ref]$generated) -and $generated.UtcDateTime -ge $Report.LastWriteTimeUtc) {
                return [pscustomobject]@{ Directory=$dir.FullName; Verified=$false; Session=''; Generated=$generatedAtUtc }
            }
        }
        catch { continue }
    }
    return $null
}

function Get-DiagnosticReportKey {
    if (-not $script:CurrentMigration -or -not $script:DiagLatestReport -or -not $script:DiagReportSignature) { return '' }
    return $script:CurrentMigration.Root + '|' + $script:DiagReportSignature
}

function Start-AutomaticDiagnosticAnalysis {
    if (-not $script:CurrentMigration -or -not $script:DiagLatestReport -or $script:DiagSummary -or $script:DiagProcess) { return }
    $key = Get-DiagnosticReportKey
    if (-not $key -or -not (Test-Path -LiteralPath $script:DiagLatestReport.FullName -PathType Leaf)) { return }
    if ($script:DiagAutoAttempts.ContainsKey($key)) {
        if ($script:DiagAutoAttempts[$key]) {
            $lblDiagProgress.Text = 'Automatic analysis failed: ' + $script:DiagAutoAttempts[$key] + ' Click Analyze latest report to retry.'
            Update-DiagnosticReportStatus -Failure $script:DiagAutoAttempts[$key]
        }
        return
    }
    # One automatic attempt per input signature in this GUI session; failures need a manual retry.
    $script:DiagAutoAttempts[$key] = ''
    Start-DiagnosticAnalysis -Automatic
}

function Refresh-DiagnosticReportState {
    param([switch]$Force, [switch]$SkipAutoAnalysis)
    if (-not $script:CurrentMigration) { return }
    $folder = Join-Path $script:CurrentMigration.Root 'ShareGate\MigrationReport'
    $script:DiagInputPath = $folder
    $lblDiagInputPath.Text = Get-DiagnosticReportDisplayPath -Folder $folder
    $lblDiagMigration.Text = 'Selected migration: ' + $script:CurrentMigration.Name
    $files = @()
    if (Test-Path -LiteralPath $folder -PathType Container) {
        $files = @(Get-ChildItem -LiteralPath $folder -File -ErrorAction SilentlyContinue | Where-Object Extension -In @('.csv', '.xlsx') | Sort-Object -Property @{ Expression='LastWriteTimeUtc'; Descending=$true }, Name)
    }
    $report = if ($files.Count) { $files[0] } else { $null }
    if ($report -and $report.Extension -eq '.xlsx') {
        $csv = @($files | Where-Object { $_.Extension -eq '.csv' -and $_.BaseName -eq $report.BaseName } | Select-Object -First 1)
        if ($csv.Count) { $report = $csv[0] }
    }
    $signature = if ($report) { '{0}|{1}|{2}' -f $report.FullName, $report.Length, $report.LastWriteTimeUtc.Ticks } else { '(empty)' }
    if (-not $Force -and $signature -eq $script:DiagReportSignature) {
        if (-not $SkipAutoAnalysis) { Start-AutomaticDiagnosticAnalysis }
        return
    }
    $script:DiagReportSignature = $signature
    $script:DiagLatestReport = $report
    $script:DiagReportHash = ''
    Clear-DiagnosticResult
    $lblDiagLatestReport.ToolTip = if ($report) { $report.FullName } else { $folder }
    $isRunning = $script:DiagProcess -and -not $script:DiagProcess.HasExited
    $btnDiagAnalyze.IsEnabled = [bool]($report -and -not $isRunning)
    if (-not $report) {
        $lblDiagLatestReport.Text = 'No ShareGate report found. Place the latest migration report (CSV or XLSX) in MigrationReport, then click Refresh reports.'
        $lblDiagProgress.Text = 'Analysis unavailable until a report is deposited.'
        Update-DiagnosticReportStatus
        return
    }
    $size = '{0:N1} MB' -f ($report.Length / 1MB)
    $lblDiagLatestReport.Text = "Latest report: $($report.Name) | $($report.LastWriteTime.ToString('yyyy-MM-dd HH:mm')) | $size | $($report.Extension.ToUpperInvariant())"
    $lblDiagLatestReport.ToolTip = $report.FullName
    $cached = Get-CurrentDiagnosticAnalysis -Report $report
    if ($cached) {
        $script:DiagAnalysisVerified = [bool]$cached.Verified
        Load-DiagnosticResult -Directory $cached.Directory
        $basis = if ($cached.Verified) { 'SHA256 verified' } else { 'legacy analysis: path and timestamp only' }
        $session = if ($cached.Session) { " | session $($cached.Session)" } else { '' }
        $lblDiagProgress.Text = "Existing analysis for latest report ($basis$session): $($cached.Generated) UTC."
    }
    elseif ($report.Extension -eq '.xlsx' -and -not $isRunning) {
        $lblDiagProgress.Text = 'Ready to analyze XLSX. ImportExcel installs automatically for the current user if needed.'
    }
    else { $lblDiagProgress.Text = 'Latest report detected. Waiting for automatic analysis to create its HTML report and summary.' }
    Update-DiagnosticReportStatus
    if (-not $SkipAutoAnalysis) { Start-AutomaticDiagnosticAnalysis }
}

function Refresh-TransientResults {
    $cardTransient.Visibility = 'Collapsed'
    $script:TransientResultsPath = ''
    $script:TransientOutOfBatchPath = ''
    $btnTransientOpen.IsEnabled = $false
    $btnTransientOutOfBatch.IsEnabled = $false
    $lblTransientStatus.Text = 'No batch result found for this migration.'
    $lblTransientCounts.Text = 'A reviewed ShareGate batch run will appear here.'
    if (-not $script:CurrentMigration) { return }
    $diagnostics = Join-Path $script:CurrentMigration.Root 'ShareGate\Diagnostics'
    if (-not (Test-Path -LiteralPath $diagnostics -PathType Container)) { return }
    $runs = @(Get-ChildItem -LiteralPath $diagnostics -Directory -Filter 'Transient-*' -ErrorAction SilentlyContinue | Sort-Object Name -Descending)
    foreach ($run in $runs) {
        $summaryPath = Join-Path $run.FullName 'Transient-Summary.json.txt'
        if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf)) { continue }
        try {
            $summary = Get-Content -LiteralPath $summaryPath -Raw -ErrorAction Stop | ConvertFrom-Json -AsHashtable
            if ($summary.SchemaVersion -ne 1) { continue }
            $total = [int]$summary.Success + [int]$summary.Skipped + [int]$summary.Error +
                [int]$summary.Warning + [int]$summary.Mixed + [int]$summary.Unreported + [int]$summary.NotAttempted
            if ($total -ne [int]$summary.PlannedItems) { throw 'Batch result counters do not equal PlannedItems.' }
            $lblTransientStatus.Text = "Latest run: $($summary.RunStatus) | session $($summary.SessionId) | $($summary.CompletedBatches)/$($summary.BatchCount) batches | $($summary.GeneratedAtUtc) UTC"
            $lblTransientCounts.Text = "Planned: $($summary.PlannedItems) | Success: $($summary.Success) | Skipped: $($summary.Skipped) | Error: $($summary.Error) | Warning: $($summary.Warning) | Mixed: $($summary.Mixed) | Unreported: $($summary.Unreported) | Not attempted: $($summary.NotAttempted) | Excluded for separate handling: $($summary.OutOfBatchLines) ($($summary.OutOfBatchSiteLines) Site, $($summary.OutOfBatchFileLines) File) | Home pages separate: $($summary.SeparatePageItems)"
            $cardTransient.Visibility = 'Visible'
            $resultsPath = Join-Path $run.FullName 'Transient-Results.csv'
            if ($resultsPath -and (Test-Path -LiteralPath $resultsPath -PathType Leaf)) {
                $resolved = (Resolve-Path -LiteralPath $resultsPath).ProviderPath
                if ($resolved.StartsWith($run.FullName.TrimEnd('\') + '\', [StringComparison]::OrdinalIgnoreCase)) {
                    $script:TransientResultsPath = $resolved
                    $btnTransientOpen.IsEnabled = $true
                }
            }
            $horsLotPath = Join-Path $run.FullName 'Transient-HorsLot.csv'
            if (Test-Path -LiteralPath $horsLotPath -PathType Leaf) {
                $script:TransientOutOfBatchPath = $horsLotPath
                $btnTransientOutOfBatch.IsEnabled = $true
            }
            return
        }
        catch {
            $lblTransientStatus.Text = "Batch summary cannot be read: $summaryPath"
            $lblTransientCounts.Text = $_.Exception.Message
            return
        }
    }
}

function Show-FarmDiagnosticsWindow {
    if (-not $script:CurrentMigration) { return }
    $farmDiagnosticsHost.Child = $null
    $farmWindow = $null
    try {
        $farmWindow = New-FarmDiagnosticsWindow -Owner $script:Window -MigrationName $script:CurrentMigration.Name -Content $cardFarmDiagnostics
        Refresh-FarmDiagnostics
        [void]$farmWindow.ShowDialog()
    }
    finally {
        if ($farmWindow) { $farmWindow.Content.Children[1].Content = $null }
        $farmDiagnosticsHost.Child = $cardFarmDiagnostics
    }
}

function Refresh-FarmDiagnostics {
    $script:FarmReportPath = ''
    $script:FarmInvocation = $null
    $script:FarmPrerequisitesPassed = $false
    $btnFarmOpenReport.IsEnabled = $false
    $btnFarmCheck.IsEnabled = $false
    $btnFarmRun.IsEnabled = $false
    $txtFarmDryRun.Text = ''
    $txtFarmRun.Text = ''
    $lblFarmPrerequisites.Text = 'Requires: farm server, elevated account with SharePoint Shell Admin rights, Windows PowerShell 5.1, SharePoint snap-in/module, shared UNC toolkit and valid access-peak CSV. Run the displayed DryRun first, then check prerequisites.'
    if (-not $script:CurrentMigration) { return }
    $project = [string]$script:CurrentMigration.Name
    $diagnostics = Join-Path $script:CurrentMigration.Root 'ShareGate\Diagnostics'
    $lblFarmResult.Text = 'No farm diagnostic result found for this project.'
    $lblFarmPeaks.Text = 'Analyze ShareGate reports to prepare farm windows.'
    if (-not (Test-Path -LiteralPath $diagnostics -PathType Container)) { return }
    $farmDirs = @(Get-ChildItem -LiteralPath $diagnostics -Directory -Filter 'Farm-*' -ErrorAction SilentlyContinue | Sort-Object Name -Descending)
    foreach ($dir in $farmDirs) {
        $summaryPath = Join-Path $dir.FullName 'Farm-Summary.json.txt'
        if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf)) { continue }
        try {
            $result = Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json -AsHashtable
            if ($result.SchemaVersion -ne 1 -or $result.Project -ne $project) { continue }
            $missing = @($result.Coverage | Where-Object { $_.Status -in @('Missing','Partial') }).Count
            $lblFarmResult.Text = "Latest farm result: $($result.Status); $missing missing or partial source(s); $($result.GeneratedAtUtc) UTC; $dir"
            if (Test-Path -LiteralPath $result.ReportPath -PathType Leaf) {
                $script:FarmReportPath = [string]$result.ReportPath
                $btnFarmOpenReport.IsEnabled = $true
            }
            break
        }
        catch { $lblFarmResult.Text = "Farm result cannot be read: $summaryPath. $($_.Exception.Message)" }
    }
    if (-not $script:DiagAnalysisVerified) {
        $lblFarmPeaks.Text = 'Analyze the latest report to record its SHA256 before running source farm diagnostics.'
        return
    }
    $analysisFiles = @()
    if ($script:DiagLoadedDirectory -and (Test-Path -LiteralPath $script:DiagLoadedDirectory -PathType Container)) {
        $analysisFiles = @(Get-ChildItem -LiteralPath $script:DiagLoadedDirectory -File -Filter 'AccessFailures-5min.csv' -ErrorAction SilentlyContinue)
    }
    if (-not $analysisFiles.Count) { return }
    $peaks = @()
    $peakFile = $null
    foreach ($file in $analysisFiles) {
        try {
            $peaks = @(Import-Csv -LiteralPath $file.FullName | Where-Object { [int]$_.Lines -gt 0 -and $_.WindowUtc -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2} UTC$' } | Sort-Object { [int]$_.Lines } -Descending)
            if ($peaks.Count) { $peakFile = $file; break }
        }
        catch { continue }
    }
    if (-not $peakFile) { $lblFarmPeaks.Text = 'No valid five-minute access peak is available.'; return }
    $peak = $peaks[0]
    $utc = ([string]$peak.WindowUtc -replace ' UTC$','').Replace(' ','T') + ':00Z'
    $relative = $peakFile.FullName.Substring($script:CurrentMigration.Root.TrimEnd('\').Length).TrimStart('\')
    $farmRoot = $script:FarmToolkitRoot
    if (-not $farmRoot.StartsWith('\\')) {
        $lblFarmPeaks.Text = 'Set -FarmToolkitRoot to the shared UNC toolkit path to generate farm commands.'
        return
    }
    $scriptPath = Join-Path $farmRoot 'Scripts\Diagnostics\SmartM365-SharePointMigration-FarmDiagnostic.ps1'
    $uncPeakPath = Join-Path (Join-Path (Join-Path $farmRoot 'Migrations') $project) $relative
    $script:FarmInvocation = [pscustomobject]@{ ScriptPath=$scriptPath; Project=$project; Around=$utc; PeaksCsv=$uncPeakPath; ToolkitRoot=$farmRoot }
    $btnFarmCheck.IsEnabled = $true
    $common = 'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "{0}" -Project "{1}" -Around "{2}" -WindowMinutes 30 -ShareGatePeaksCsv "{3}" -ToolkitRoot "{4}"' -f $scriptPath,$project,$utc,$uncPeakPath,$farmRoot
    $txtFarmDryRun.Text = $common + ' -DryRun'
    $txtFarmRun.Text = $common
    $lblFarmPeaks.Text = "ShareGate access windows: $($peaks.Count) in $($peakFile.Name). Largest: $($peak.WindowUtc), $($peak.Lines) lines. Both commands include every window through the peaks CSV."
}

function Test-FarmDiagnosticsPrerequisites {
    Refresh-DiagnosticReportState -Force
    $script:FarmPrerequisitesPassed = $false
    $btnFarmRun.IsEnabled = $false
    $issues = [System.Collections.Generic.List[string]]::new()
    $invocation = $script:FarmInvocation
    if (-not $invocation) { $issues.Add('Analyze a report with valid access-failure windows first.') }
    else {
        if (-not $invocation.ToolkitRoot.StartsWith('\\') -or -not (Test-Path -LiteralPath $invocation.ScriptPath -PathType Leaf)) {
            $issues.Add('Shared UNC toolkit or farm diagnostic script is unavailable. Set -FarmToolkitRoot.')
        }
        if (-not (Test-Path -LiteralPath $invocation.PeaksCsv -PathType Leaf)) {
            $issues.Add('The access-peak CSV is not available on the shared UNC path.')
        }
    }
    $ps51 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $ps51 -PathType Leaf)) { $issues.Add('Windows PowerShell 5.1 is unavailable.') }
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = [Security.Principal.WindowsPrincipal]::new($identity)
    if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
        $issues.Add('Start the GUI elevated on a SharePoint farm server.')
    }
    if (-not $issues.Count) {
        $check = @'
$ErrorActionPreference = 'Stop'
if ($PSVersionTable.PSVersion.ToString() -notlike '5.1.*') { throw 'Windows PowerShell 5.1 is required.' }
if (-not (Get-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue)) {
    try { Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop }
    catch { Import-Module SharePointServer -ErrorAction Stop }
}
$null = Get-SPFarm -ErrorAction Stop
$localName = $env:COMPUTERNAME
$servers = @(Get-SPServer -ErrorAction Stop)
if (-not @($servers | Where-Object { ([string]$_.Address).Split('.')[0] -eq $localName }).Count) {
    throw "This machine ($localName) is not a server in the SharePoint farm."
}
'@
        $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($check))
        try {
            $output = @(& $ps51 -NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand $encoded 2>&1)
            if ($LASTEXITCODE -ne 0) { throw (($output | ForEach-Object { [string]$_ }) -join ' ') }
        }
        catch { $issues.Add('SharePoint farm access failed in Windows PowerShell 5.1: ' + $_.Exception.Message) }
    }
    if ($issues.Count) {
        $lblFarmPrerequisites.Text = 'Run unavailable: ' + ($issues -join ' | ')
        return $false
    }
    $script:FarmPrerequisitesPassed = $true
    $btnFarmRun.IsEnabled = $true
    $lblFarmPrerequisites.Text = 'Ready: elevated farm server, SharePoint Shell access, Windows PowerShell 5.1, shared toolkit and access-peak CSV verified. Run collects read-only farm evidence.'
    return $true
}

function Load-DiagnosticResult {
    param([string]$Directory)
    $summaryPath = Join-Path $Directory 'Summary.json.txt'
    if (-not (Test-Path -LiteralPath $summaryPath -PathType Leaf)) { throw "Analysis summary was not created: $summaryPath" }
    $script:DiagSummary = Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json -AsHashtable
    $script:DiagLoadedDirectory = $Directory
    $script:DiagSummary.ReportPath = Join-Path $Directory 'MigrationDiagnostics-Report.html'
    $script:DiagSummary.RowsPath = Join-Path $Directory 'ClassifiedRows.csv'
    $sessionLabel = if (@($script:DiagSummary.Sessions).Count) { @($script:DiagSummary.Sessions) -join ', ' } else { '(none)' }
    $lblDiagLatestReport.Text += " | Session: $sessionLabel | Lines: $($script:DiagSummary.Lines)"
    $script:DiagRows = @(Import-Csv -LiteralPath $script:DiagSummary.RowsPath)
    $selected = if ($cmbDiagSession.SelectedItem) { [string]$cmbDiagSession.SelectedItem } else { 'All sessions' }
    $script:DiagKnownSessions = @($script:DiagKnownSessions + @($script:DiagSummary.Sessions) | Sort-Object -Unique)
    $script:DiagLoading = $true
    try {
        $cmbDiagSession.Items.Clear()
        [void]$cmbDiagSession.Items.Add('All sessions')
        foreach ($session in $script:DiagKnownSessions) { [void]$cmbDiagSession.Items.Add([string]$session) }
        $index = $cmbDiagSession.Items.IndexOf($selected)
        $cmbDiagSession.SelectedIndex = if ($index -ge 0) { $index } else { 0 }
    }
    finally { $script:DiagLoading = $false }
    $lines = $script:DiagSummary.Lines
    $lineStatuses = $script:DiagSummary.LineStatus
    $lineStates = $script:DiagSummary.IssueLineState
    $itemStates = $script:DiagSummary.IssueItemState
    $lineRate = if ($null -ne $script:DiagSummary.ResidualLineRate) { '{0:N2}%' -f [double]$script:DiagSummary.ResidualLineRate } else { 'n/a' }
    $itemRate = if ($null -ne $script:DiagSummary.ResidualItemRate) { '{0:N2}%' -f [double]$script:DiagSummary.ResidualItemRate } else { 'n/a' }
    $lblDiagLines.Text = '{0:N0}' -f [int]$lines
    $lblDiagKeyedItems.Text = '{0:N0}' -f [int]$script:DiagSummary.DistinctItems
    $lblDiagSuccess.Text = '{0:N0}' -f [int]$lineStatuses['Success']
    $lblDiagError.Text = '{0:N0}' -f [int]$lineStatuses['Error']
    $lblDiagWarning.Text = '{0:N0}' -f [int]$lineStatuses['Warning']
    $lblDiagAccepted.Text = '{0:N0}' -f [int]$lineStates['Accepted']
    $lblDiagToFixLines.Text = '{0:N0}' -f [int]$lineStates['To fix']
    $lblDiagUnkeyedLines.Text = '{0:N0}' -f [int]$script:DiagSummary.UnkeyedRows
    $lblDiagItemsToFix.Text = '{0:N0}' -f [int]$itemStates['To fix']
    $lblDiagResidualLines.Text = $lineRate
    $lblDiagResidualItems.Text = $itemRate
    $lblDiagKpis.Visibility = 'Collapsed'
    $panelDiagMetrics.Visibility = 'Visible'
    $lblSummaryShareGateValue.Text = [string]([int]$itemStates['To fix'])
    $verification = if ($script:DiagAnalysisVerified) { 'SHA256 verified' } else { 'legacy analysis' }
    $lblSummaryShareGateDetail.Text = "$($script:DiagSummary.DistinctItems) keyed items · $itemRate residual · $verification"
    $lblSummaryShareGateDetail.ToolTip = "ShareGate analysis: $($script:DiagSummary.ReportPath)"
    $btnDiagOpenReport.IsEnabled = (Test-Path -LiteralPath $script:DiagSummary.ReportPath -PathType Leaf)
    $panelDiagReview.IsEnabled = $true
    $lblDiagProgress.Text = "Analysis completed: $(@($script:DiagSummary.Sessions).Count) session(s); $($script:DiagSummary.DuplicateRowsSuppressed) duplicate rows suppressed."
    if ($script:DiagSummary.ConflictingDuplicateRows -gt 0) {
        $lblDiagProgress.Text += " $($script:DiagSummary.ConflictingDuplicateRows) conflicting duplicate rows require review."
    }
    Refresh-DiagnosticPatterns
    Refresh-FarmDiagnostics
    Refresh-DiagnosticCrossCheck
}

function Update-DiagnosticAnalysisProgress {
    if (-not $script:CurrentMigration -or $script:CurrentMigration.Root -ne $script:DiagProjectRoot) { return }
    if ($script:DiagActiveReportKey -ne (Get-DiagnosticReportKey)) { return }
    $path = Join-Path $script:DiagOutputDirectory 'analysis.phase.json.txt'
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { return }
    try {
        $phase = Get-Content -LiteralPath $path -Raw | ConvertFrom-Json
        if (-not $phase.Message) { return }
        $lblDiagProgress.Text = [string]$phase.Message
        $lblDiagAnalysisState.Text = if ($phase.State -eq 'InstallingImportExcel') { 'Installing module' } else { 'Running' }
        $lblDiagAnalysisEvidence.Text = if ($phase.State -eq 'InstallingImportExcel') { 'ImportExcel · CurrentUser · PSGallery' } else { 'Processing latest report' }
        $lblDiagNextAction.Text = 'Wait for preparation and report analysis to finish.'
    }
    catch { } # A phase file can briefly be incomplete while the worker writes it.
}

function Start-DiagnosticAnalysis {
    param([switch]$Automatic)
    if ($script:DiagProcess) { return }
    if (-not $script:CurrentMigration) { return }
    Refresh-DiagnosticReportState -SkipAutoAnalysis
    $report = $script:DiagLatestReport
    if (-not $report -or -not (Test-Path -LiteralPath $report.FullName -PathType Leaf)) {
        [System.Windows.MessageBox]::Show('Place the latest ShareGate CSV or XLSX report in MigrationReport first.', $script:AppName, 'OK', 'Warning') | Out-Null
        return
    }
    $inputPath = $report.FullName
    $wrapper = Join-Path $script:ScriptRoot 'Scripts\Diagnostics\SmartM365-SharePointMigration-Diagnostics.ps1'
    $output = Join-Path $script:CurrentMigration.Root ('ShareGate\Diagnostics\{0}-{1}' -f (Get-Date -Format 'yyyyMMdd-HHmmss'), [guid]::NewGuid().ToString('N'))
    $activity = $null
    $key = Get-DiagnosticReportKey
    $script:DiagAutoAttempts[$key] = ''
    try {
        New-Item -ItemType Directory -Path $output -Force | Out-Null
        $activity = New-SmartM365GuiActivity -ProjectRoot $script:ScriptRoot -Migration $script:CurrentMigration.Name -Action 'MigrationDiagnostics'
        $arguments = @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', "`"$wrapper`"",
            '-ProjectRoot', "`"$($script:CurrentMigration.Root)`"", '-InputPath', "`"$inputPath`"",
            '-OutputDirectory', "`"$output`"", '-ActivityPath', "`"$activity`"")
        $session = if (-not $Automatic -and $cmbDiagSession.SelectedItem) { [string]$cmbDiagSession.SelectedItem } else { 'All sessions' }
        if ($session -ne 'All sessions') { $arguments += @('-SessionId', "`"$session`"") }
        $exe = Join-Path $PSHOME 'pwsh.exe'
        $script:DiagProcess = Start-Process -FilePath $exe -ArgumentList $arguments -WorkingDirectory $script:ScriptRoot -WindowStyle Hidden -PassThru `
            -RedirectStandardOutput (Join-Path $output 'analysis.stdout.log') -RedirectStandardError (Join-Path $output 'analysis.stderr.log') -ErrorAction Stop
        $script:DiagOutputDirectory = $output
        $script:DiagActivity = $activity
        $script:DiagProjectRoot = $script:CurrentMigration.Root
        $script:DiagActiveReportKey = $key
        $btnDiagAnalyze.IsEnabled = $false
        $lblDiagProgress.Text = if ($report.Extension -eq '.xlsx') { 'Preparing XLSX analysis; ImportExcel will be installed automatically if needed…' } else { "Analyzing local report files in $inputPath ..." }
        Update-DiagnosticReportStatus
        $script:DiagTimer.Start()
        Refresh-ActivityList
    }
    catch {
        $script:DiagAutoAttempts[$key] = $_.Exception.Message
        if ($activity) { Write-SmartM365GuiActivityEvent -Path $activity -Status 'Failed' -ExitCode 1 -Detail $_.Exception.Message }
        $lblDiagProgress.Text = "Could not start analysis: $($_.Exception.Message)"
        $btnDiagAnalyze.IsEnabled = [bool]$script:DiagLatestReport
        Update-DiagnosticReportStatus -Failure $_.Exception.Message
    }
}

$btnDiagOpenFolder.Add_Click({
    if (-not $script:CurrentMigration) { return }
    $folder = Join-Path $script:CurrentMigration.Root 'ShareGate\MigrationReport'
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) { [void](New-Item -ItemType Directory -Path $folder -Force) }
    Open-InExplorer $folder
})
$btnDiagRefresh.Add_Click({ Refresh-DiagnosticReportState -Force; Refresh-DiagnosticCrossCheck -Force })
$btnCrossCheckRefresh.Add_Click({ Refresh-DiagnosticReportState -Force; Refresh-DiagnosticCrossCheck -Force })
$btnCrossCheckFilesReport.Add_Click({ if ($btnCrossCheckFilesReport.Tag) { Open-InExplorer $btnCrossCheckFilesReport.Tag } })
$btnCrossCheckPermissionsReport.Add_Click({ if ($btnCrossCheckPermissionsReport.Tag) { Open-InExplorer $btnCrossCheckPermissionsReport.Tag } })
$btnFarmRefresh.Add_Click({ Refresh-FarmDiagnostics })
$btnFarmWindow.Add_Click({ Show-FarmDiagnosticsWindow })
$btnFarmCheck.Add_Click({ [void](Test-FarmDiagnosticsPrerequisites) })
$btnFarmRun.Add_Click({
    if (-not (Test-FarmDiagnosticsPrerequisites)) { return }
    $invocation = $script:FarmInvocation
    $ps51 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    try {
        $arguments = @('-NoProfile','-ExecutionPolicy','Bypass','-File',('"' + $invocation.ScriptPath + '"'),
            '-Project',('"' + $invocation.Project + '"'),'-Around',('"' + $invocation.Around + '"'),
            '-WindowMinutes','30','-ShareGatePeaksCsv',('"' + $invocation.PeaksCsv + '"'),
            '-ToolkitRoot',('"' + $invocation.ToolkitRoot + '"'))
        $process = Start-Process -FilePath $ps51 -ArgumentList $arguments -PassThru -WindowStyle Normal -ErrorAction Stop
        $lblFarmPrerequisites.Text = "Read-only farm diagnostic started in Windows PowerShell 5.1 (PID $($process.Id)). Refresh farm results after completion."
    }
    catch { $lblFarmPrerequisites.Text = "Could not start farm diagnostic: $($_.Exception.Message)" }
})
$btnTransientRefresh.Add_Click({ Refresh-TransientResults })
$btnTransientOpen.Add_Click({ if ($script:TransientResultsPath) { Open-InExplorer $script:TransientResultsPath } })
$btnTransientOutOfBatch.Add_Click({ if ($script:TransientOutOfBatchPath) { Open-InExplorer $script:TransientOutOfBatchPath } })
$btnFarmOpenReport.Add_Click({ if ($script:FarmReportPath) { Open-InExplorer $script:FarmReportPath } })
$btnDiagAnalyze.Add_Click({ Start-DiagnosticAnalysis })
$btnDiagOpenReport.Add_Click({ if ($script:DiagSummary) { Open-InExplorer $script:DiagSummary.ReportPath } })
$btnDiagFilter.Add_Click({ Refresh-DiagnosticPatterns })
$btnDiagRowFilter.Add_Click({ Refresh-DiagnosticRows })
$cmbDiagStatus.Add_SelectionChanged({ if ($script:DiagSummary) { Refresh-DiagnosticPatterns } })
$cmbDiagSession.Add_SelectionChanged({
    if (-not $script:DiagLoading -and $script:DiagSummary -and $cmbDiagSession.SelectedItem) { $lblDiagProgress.Text = 'Click Analyze latest report to apply the selected session.' }
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
$script:ActivityRetentionTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:ActivityRetentionTimer.Interval = [TimeSpan]::FromHours(1)
$script:ActivityRetentionTimer.Add_Tick({
    try { Invoke-ActivityLogRetention }
    catch {
        Microsoft.PowerShell.Utility\Write-Warning ('{0} Activity log retention failed: {1}' -f
            (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$_.Exception.Message)
    }
})
function Complete-DiagnosticAnalysis {
    if (-not $script:DiagProcess -or -not $script:DiagProcess.HasExited) { return }
    $script:DiagTimer.Stop()
    $reportKey = $script:DiagActiveReportKey
    $code = $script:DiagProcess.ExitCode
    $script:DiagProcess.Dispose()
    $script:DiagProcess = $null
    $btnDiagAnalyze.IsEnabled = [bool]$script:DiagLatestReport
    try {
        if ($code -ne 0) {
            $stderrPath = Join-Path $script:DiagOutputDirectory 'analysis.stderr.log'
            $errorText = if (Test-Path -LiteralPath $stderrPath -PathType Leaf) { Get-Content -LiteralPath $stderrPath -Raw } else { '' }
            throw "Analysis exited with code $code. $errorText"
        }
        if ($script:CurrentMigration -and $script:CurrentMigration.Root -eq $script:DiagProjectRoot) {
            Refresh-DiagnosticReportState -Force -SkipAutoAnalysis
            if ($reportKey -eq (Get-DiagnosticReportKey) -and -not $script:DiagSummary) {
                throw 'Analysis finished without a matching summary and HTML report.'
            }
        }
    }
    catch {
        $script:DiagAutoAttempts[$reportKey] = $_.Exception.Message
        if ($reportKey -eq (Get-DiagnosticReportKey)) {
            $lblDiagProgress.Text = "Analysis failed: $($_.Exception.Message)"
            Update-DiagnosticReportStatus -Failure $_.Exception.Message
        }
        if ($script:DiagActivity) {
            Write-SmartM365GuiActivityEvent -Path $script:DiagActivity -Status 'Failed' -ExitCode 1 -Detail $_.Exception.Message
        }
    }
    Refresh-ActivityList
    Start-AutomaticDiagnosticAnalysis
}
$script:DiagTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:DiagTimer.Interval = [TimeSpan]::FromSeconds(1)
$script:DiagTimer.Add_Tick({
    if (-not $script:DiagProcess) { return }
    if (-not $script:DiagProcess.HasExited) { Update-DiagnosticAnalysisProgress; return }
    Complete-DiagnosticAnalysis
})
$script:CrossCheckTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:CrossCheckTimer.Interval = [TimeSpan]::FromMilliseconds(500)
$script:CrossCheckTimer.Add_Tick({
    $job = $script:CrossCheckJob
    if (-not $job) { return }
    if ($panelCrossCheckLoading.Visibility -eq 'Visible' -and $script:CrossCheckLoadingStarted) {
        $elapsed = [int]((Get-Date) - $script:CrossCheckLoadingStarted).TotalSeconds
        $lblCrossCheckLoading.Text = "Loading comparison reports for $($script:CurrentMigration.Name)… $elapsed s"
    }
    if ($job.State -in @('Running','NotStarted')) { return }
    $script:CrossCheckTimer.Stop()
    $script:CrossCheckJob = $null
    $currentSignature = Get-DiagnosticCrossCheckSignature
    $isCurrent = $script:CrossCheckJobSignature -eq $currentSignature
    try {
        if ($job.State -ne 'Completed') { throw "Background cross-check ended with state $($job.State)." }
        $result = @(Receive-Job -Job $job -ErrorAction Stop)[0]
        if (-not $result) { throw 'Background cross-check returned no result.' }
        if ($isCurrent) {
            $panelCrossCheckLoading.Visibility = 'Collapsed'
            $script:CrossCheckResult = $result
            Update-DiagnosticSummaryCrossCheck $result
            $gridCrossCheckEvidence.ItemsSource = @($result.Evidence)
            $gridCrossCheckScopes.ItemsSource = @($result.Scopes)
            $btnCrossCheckFilesReport.Tag = $result.FilesReport
            $btnCrossCheckPermissionsReport.Tag = $result.PermissionsReport
            $btnCrossCheckFilesReport.IsEnabled = [bool]$result.FilesReport
            $btnCrossCheckPermissionsReport.IsEnabled = [bool]$result.PermissionsReport
            $lblCrossCheckStatus.Text = ('{0} scopes with differences; {1} ambiguous; unmatched rows: ShareGate {2}, files {3}, permissions {4}. Only matching site and list titles are placed on one row.' -f `
                @($result.Scopes).Count, $result.Ambiguous, $result.Unmatched.ShareGate,
                $result.Unmatched.Files, $result.Unmatched.Permissions)
            $script:CrossCheckSignature = $currentSignature
        }
    }
    catch {
        if ($isCurrent) {
            $panelCrossCheckLoading.Visibility = 'Collapsed'
            $lblCrossCheckStatus.Text = 'Cross-check unavailable: ' + $_.Exception.Message
            $lblSummaryFilesDetail.Text = 'Cross-check unavailable'
            $lblSummaryPermissionsDetail.Text = 'Cross-check unavailable'
            $lblSummaryCrossCheckDetail.Text = 'Cross-check unavailable'
        }
    }
    finally { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue }
    if (-not $isCurrent -and $tabDiagnostics.IsChecked) { Refresh-DiagnosticCrossCheck -Force }
})
$lblDiagScope.Text = 'Analysis only: no ShareGate connection or migration action. XLSX support installs automatically if needed.'
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
Initialize-SmartM365BatchGui -Panel $panelBatches -Root $script:ScriptRoot -SourceRoot $script:FarmToolkitRoot
Refresh-GuiState
$script:AutoRefreshTimer.Start()
$script:ActivityRetentionTimer.Start()
Close-SmartM365GuiSplash -Splash $script:Splash
try { [void]$script:Window.ShowDialog() }
finally {
    Stop-SmartM365BatchGui
    $script:AutoRefreshTimer.Stop()
    $script:ActivityRetentionTimer.Stop()
    if ($script:DiagTimer) { $script:DiagTimer.Stop() }
    if ($script:CrossCheckTimer) { $script:CrossCheckTimer.Stop() }
    Write-SmartM365GuiActivityEvent -Path $script:SessionActivity -Status 'Closed' -ExitCode 0 `
        -Detail 'GUI window closed.'
}




# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAx5yBwa+yyRtKC
# Kr6mu94Ra+G72dHSWN07YofZiuMawaCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCiSRimbZVPRDV1hmoWsH5v
# c19Tek2NMVR/paj1102ULzANBgkqhkiG9w0BAQEFAASCAYAJoX5CKW7rSfix9I3h
# UISzWh/9J9zJaH/byowSyJ9R9s3EdnEeT5q1y3DnJCfDE5BclFtEMuv7H6CljwoW
# JayjR/6oY9nOXmjY9WvXsIiCwdVnVtt0x9CanS5bnZwv9li7yZLIXrJBN7B7j1Al
# NSQl4evoUb1O4fb/xZLC0d33zCqz/gALB/izi9+uljDOCd1X1fXlRIv68zzF9+G+
# fPYeOjrc87NWoOoAN1g1xv3YLbNMHGilcAr5hWmQPAvmIIgZ6n/lmmgDFQIPXb1d
# z7UYtwYsHMyXlu5ADzmMq7D80obUj4jVIF3Mjq4G3XPtMmDf32dSnLJlmOrpgtCw
# AhfKOETBPvUhICBo9x7pjhtkZCPwvdAVtHEM63ir+QeZewbcm7GXoQN6yqIp1PAl
# 9SMSbQcT7ZtfWJRY6Gp4NiXSQVDWC+4TfiKKc998ke33XmDR4JOzrpHEJx5u4vkf
# 5Mq6zCpxGp8flVj+B+SLOF5VxMHQWmwF4TjzCmDstujEoz4=
# SIG # End signature block
