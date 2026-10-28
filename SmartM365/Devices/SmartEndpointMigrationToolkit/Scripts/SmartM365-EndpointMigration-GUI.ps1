<#
.SYNOPSIS
    Smart Endpoint Migration Toolkit GUI (WPF).

.DESCRIPTION
    Hybrid Join, Intune Enrollment and Windows 11 Upgrade campaigns from one launcher:
    Campaign selector, Waves, New Wave, Automatic Wave (Preview distinct from Create),
    Single PC, Launch Wave, Launch All Waves (selected Campaign only, staggered), Runs and
    reports, controlled stop. Options are generated from the option catalog; the choices
    of the operator are captured for each Run and isolated per Campaign.
    Requires Windows PowerShell 5.1 started with -STA.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

param(
    [switch]$ValidateOnly,
    [switch]$Demo,
    [ValidateSet('Overview','Issues','Completed')][string]$DemoScenario = 'Overview',
    [string]$DemoView = '',
    [ValidateRange(1000,3840)][int]$DemoWindowWidth = 1500,
    [ValidateRange(650,2160)][int]$DemoWindowHeight = 950
)

if ($Demo) {
    . (Join-Path $PSScriptRoot 'SmartM365.GuiDemo.ps1')
    Start-SmartM365GuiDemo -GuiScriptPath $PSCommandPath -ProfilePath (Join-Path $PSScriptRoot 'SmartM365-EndpointMigration-Demo.ps1') -Scenario $DemoScenario -View $DemoView -Width $DemoWindowWidth -Height $DemoWindowHeight -ValidateOnly:$ValidateOnly
    return
}

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if ([string]$PSVersionTable['PSEdition'] -eq 'Core' -or [Threading.Thread]::CurrentThread.ApartmentState -ne [Threading.ApartmentState]::STA) {
    throw ('SEMT_UNSUPPORTED_GUI_HOST: the GUI runs in Windows PowerShell 5.1 started with -STA (current: {0} {1}, {2}). Use Start-SmartEndpointMigrationToolkit-GUI.cmd.' -f [string]$PSVersionTable['PSEdition'],[string]$PSVersionTable['PSVersion'],[Threading.Thread]::CurrentThread.ApartmentState)
}
$script:GuiVersion = '1.0.0'
$script:PackageRoot = Split-Path -Path $PSScriptRoot -Parent
Import-Module (Join-Path $script:PackageRoot 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force
$script:CampaignTitles = [ordered]@{ HybridJoin = 'Hybrid Join and Intune enrollment'; Windows11Upgrade = 'Windows 11 upgrade' }
$script:CategoryOrder = @('Mode','Actions','Setup','Safety','Schedule','Parallelism','Network','Inventory','Evidence')
$script:InventoryRefresh = $null
$script:GuiClosing = $false
$script:AllAdDomainsLabel = '(All domains)'
$script:AdDomains = @()
$script:AdDomainStamp = $null
$script:RunModes = [ordered]@{ 'Loop' = @('Loop',$false); 'Once' = @('Once',$false); 'Loop - ignore run guard (first cycle)' = @('Loop',$true); 'Once - ignore run guard' = @('Once',$true) }

#region Testable helpers (no UI)

function Assert-CmdLiteralValue {
    param([AllowNull()][string]$Value,[string]$Context = 'CMD value')
    if ($null -eq $Value) { return }
    if ($Value.IndexOf([char]0) -ge 0 -or $Value -match '[\r\n"%!]') {
        throw ("{0} contains a character that cannot be represented safely in a generated CMD launcher. Remove quotes, percent signs, exclamation marks or line breaks." -f $Context)
    }
}

function ConvertTo-CmdArgument {
    param([AllowNull()][string]$Value)
    if ($null -eq $Value) { return '""' }
    Assert-CmdLiteralValue -Value $Value -Context 'CMD argument'
    if ($Value -match '^[A-Za-z0-9_:\\./=-]+$') { return $Value }
    return ('"{0}"' -f $Value)
}

function ConvertTo-CmdWindowTitle {
    param([AllowNull()][string]$Value)
    $title = ([string]$Value -replace '[\r\n\t]+', ' ').Trim()
    $title = $title -replace '[&|<>^%!"]', '-'
    if ($title.Length -gt 240) { $title = $title.Substring(0, 240) }
    return $title
}

function New-GuiLaunchCommandFile {
    <#
    Writes a temporary CMD launcher: UTF-8 code page, delayed expansion disabled, literal paths
    quoted, pushd/popd, exit code preserved; an optional copy is kept as Run evidence.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$WorkingDirectory,
        [Parameter(Mandatory = $true)][string[]]$Commands,
        [string]$NamePrefix = 'SEMT-GUI',
        [string]$WindowTitle,
        [string]$EvidenceDirectory,
        [switch]$PauseWhenDone
    )

    $launcherRoot = Join-Path $env:TEMP 'SmartEndpointMigrationToolkit\GuiLaunchers'
    New-Item -ItemType Directory -Path $launcherRoot -Force | Out-Null
    $safePrefix = [regex]::Replace($NamePrefix, '[^A-Za-z0-9._-]+', '-').Trim('-._')
    if ([string]::IsNullOrWhiteSpace($safePrefix)) { $safePrefix = 'SEMT-GUI' }
    $launchPath = Join-Path $launcherRoot ('{0}_{1}.cmd' -f $safePrefix,(Get-Date -Format 'yyyyMMdd-HHmmss-fff'))
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('@echo off')
    $lines.Add('chcp 65001 >nul')
    $lines.Add('setlocal DisableDelayedExpansion')
    $lines.Add('echo GUI launcher : %~f0')
    if (-not [string]::IsNullOrWhiteSpace($WindowTitle)) { $lines.Add(('title {0}' -f (ConvertTo-CmdWindowTitle -Value $WindowTitle))) }
    $lines.Add(('pushd {0} || exit /b 1' -f (ConvertTo-CmdArgument -Value $WorkingDirectory)))
    foreach ($command in @($Commands)) { if (-not [string]::IsNullOrWhiteSpace($command)) { $lines.Add($command) } }
    $lines.Add('set "EXITCODE=%ERRORLEVEL%"')
    $lines.Add('popd')
    if ($PauseWhenDone) { $lines.Add('echo.'); $lines.Add('echo Finished with exit code %EXITCODE%.'); $lines.Add('pause') }
    $lines.Add('exit /b %EXITCODE%')
    [IO.File]::WriteAllLines($launchPath, $lines.ToArray(), [Text.UTF8Encoding]::new($false))
    if (-not [string]::IsNullOrWhiteSpace($EvidenceDirectory)) {
        New-Item -ItemType Directory -Path $EvidenceDirectory -Force | Out-Null
        Copy-Item -LiteralPath $launchPath -Destination (Join-Path $EvidenceDirectory (Split-Path $launchPath -Leaf)) -Force
    }
    return $launchPath
}

function Start-GuiLaunchCommandFile {
    param([Parameter(Mandatory = $true)][string]$LaunchCommandPath,[Parameter(Mandatory = $true)][string]$WorkingDirectory)
    Assert-CmdLiteralValue -Value $LaunchCommandPath -Context 'GUI launcher path'
    Start-Process -FilePath 'cmd.exe' -ArgumentList ('/d /s /c ""{0}""' -f $LaunchCommandPath) -WorkingDirectory $WorkingDirectory -Verb RunAs | Out-Null
}

function Get-GuiStatePath {
    param([Parameter(Mandatory = $true)][string]$Campaign)
    return (Join-Path $script:PackageRoot ("Config\GuiState.{0}.local.json" -f $Campaign))
}

function Get-GuiInitialOptionValues {
    <#
    Values shown when a Campaign is opened: campaign defaults, global configuration and SEMT_*
    environment, then the last GUI choices saved for this Campaign (invalid values are ignored).
    #>
    param([Parameter(Mandatory = $true)]$Catalog,[Parameter(Mandatory = $true)][string]$Campaign,[switch]$IgnoreSavedState)

    $global = Read-SemtGlobalConfig -PackageRoot $script:PackageRoot
    $resolved = Resolve-SemtEffectiveOptions -Catalog $Catalog -Campaign $Campaign -GlobalLayer $global.Data -EnvironmentLayer (Get-SemtEnvironmentOptionLayer -Catalog $Catalog -Campaign $Campaign)
    $values = @{}
    foreach ($entry in $resolved.GetEnumerator()) { $values[$entry.Key] = $entry.Value.Value }
    if (-not $IgnoreSavedState) {
        $state = $null
        try { $state = Read-SemtJsonFile -Path (Get-GuiStatePath -Campaign $Campaign) } catch { $state = $null }
        if ($state -and [string](Get-SemtPropertyValue -InputObject $state -Name 'Campaign' -Default '') -eq $Campaign) {
            $saved = ConvertTo-SemtOptionLayer -Layer $state
            foreach ($key in @($saved.Keys)) {
                if (-not $values.ContainsKey($key) -or -not $Catalog.ByName.ContainsKey($key)) { continue }
                try { $values[$key] = ConvertTo-SemtOptionValue -Definition $Catalog.ByName[$key] -Value $saved[$key] -Source 'GuiState' } catch { }
            }
        }
    }
    return $values
}

function Save-GuiState {
    param([Parameter(Mandatory = $true)][string]$Campaign,[Parameter(Mandatory = $true)][hashtable]$Values)
    Write-SemtJsonFileAtomic -Path (Get-GuiStatePath -Campaign $Campaign) -InputObject ([pscustomobject]@{ SchemaVersion = 1; Campaign = $Campaign; UpdatedUtc = (Get-SemtUtcNowText); Options = $Values })
}

function New-GuiSelectionFile {
    <#
    Captures the GUI choices for one Run. Empty text and path fields are not captured, so the
    Wave or global configuration still applies to them.
    #>
    param([Parameter(Mandatory = $true)]$Catalog,[Parameter(Mandatory = $true)][string]$Campaign,[Parameter(Mandatory = $true)][hashtable]$Values)

    $captured = [ordered]@{}
    foreach ($definition in @(Get-SemtCampaignOptionDefinitions -Catalog $Catalog -Campaign $Campaign | Where-Object { $_.GuiVisible })) {
        $name = [string]$definition.Name
        if (-not $Values.ContainsKey($name)) { continue }
        $value = $Values[$name]
        if ($definition.Type -in @('String','Path') -and [string]::IsNullOrWhiteSpace([string]$value)) { continue }
        $captured[$name] = ConvertTo-SemtOptionValue -Definition $definition -Value $value -Source 'Gui'
    }
    $root = Join-Path $script:PackageRoot ("Runs\{0}\_GuiSelections" -f $Campaign)
    $path = Join-Path $root ('{0}-{1}.json' -f (Get-Date -Format 'yyyyMMdd-HHmmss-fff'),([guid]::NewGuid().ToString('N').Substring(0, 6)))
    Write-SemtJsonFileAtomic -Path $path -InputObject ([pscustomobject]@{ SchemaVersion = 1; Product = 'SmartEndpointMigrationToolkit'; Campaign = $Campaign; CapturedUtc = (Get-SemtUtcNowText); CapturedBy = [Environment]::UserName; Options = $captured })
    return $path
}

function Test-UnlimitedCycleConfirmationRequired {
    param([string]$RunMode,[int]$MaxCycles)
    return ($RunMode -eq 'Loop' -and $MaxCycles -le 0)
}

function Get-GuiLaunchSummary {
    param(
        [Parameter(Mandatory = $true)]$Catalog,
        [Parameter(Mandatory = $true)][string]$Campaign,
        [Parameter(Mandatory = $true)][hashtable]$Values,
        [Parameter(Mandatory = $true)][string]$Target,
        [Parameter(Mandatory = $true)][string]$RunMode,
        [bool]$IgnoreRunGuard
    )

    $execution = if ($Values['DryRun']) { 'DRY RUN - DNS, ping, SMB 445 and administrative shares only; nothing is copied or executed' } elseif ($Values['AuditOnly']) { 'REMOTE AUDIT - the endpoint script runs diagnostics only' } else { 'REPAIR - the endpoint applies the enabled actions below' }
    $actions = @(Get-SemtCampaignOptionDefinitions -Catalog $Catalog -Campaign $Campaign | Where-Object { $_.Type -eq 'Boolean' -and $_.Category -in @('Actions','Setup') -and $Values.ContainsKey($_.Name) -and [bool]$Values[$_.Name] } | ForEach-Object { '  - {0}' -f $_.Label })
    $cycles = if ($RunMode -eq 'Once') { '1' } elseif ([int]$Values['MaxCycles'] -gt 0) { [string]$Values['MaxCycles'] } else { 'unlimited until complete or stopped' }
    $lines = @(
        ('Campaign: {0}' -f $script:CampaignTitles[$Campaign]),
        ('Target: {0}' -f $Target),
        ('Run mode: {0} (cycles: {1})' -f $RunMode,$cycles),
        ('Execution: {0}' -f $execution),
        ('Run guard: {0}' -f $(if ($IgnoreRunGuard) { 'IGNORED for the first cycle (endpoint and technician history)' } else { 'applied' })),
        ('Workers per Wave: {0}; shared cap on this technician computer: {1}' -f $Values['ThrottleLimit'],$Values['GlobalWorkerLimit']),
        'Enabled actions:'
    )
    if ($actions.Count -eq 0) { $lines += '  (none)' } else { $lines += $actions }
    return ($lines -join [Environment]::NewLine)
}

function Get-GuiRunCommand {
    param(
        [Parameter(Mandatory = $true)][string]$Campaign,
        [string]$WavePath,
        [string]$ComputerName,
        [Parameter(Mandatory = $true)][string]$RunMode,
        [bool]$IgnoreRunGuard,
        [Parameter(Mandatory = $true)][string]$SelectionPath
    )

    $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $engine = Join-Path $script:PackageRoot 'Scripts\SmartM365-Invoke-EndpointMigrationWithPsExec.ps1'
    $parts = @((ConvertTo-CmdArgument -Value $powershell),'-NoProfile','-ExecutionPolicy','Bypass','-File',(ConvertTo-CmdArgument -Value $engine),'-Campaign',$Campaign)
    if (-not [string]::IsNullOrWhiteSpace($WavePath)) { $parts += @('-WavePath',(ConvertTo-CmdArgument -Value $WavePath)) }
    else { $parts += @('-ComputerName',(ConvertTo-CmdArgument -Value $ComputerName)) }
    $parts += @('-RunMode',$RunMode,'-GuiSelectionPath',(ConvertTo-CmdArgument -Value $SelectionPath))
    if ($IgnoreRunGuard) { $parts += '-IgnoreRunGuard' }
    return ($parts -join ' ')
}

function ConvertTo-PowerShellLiteral {
    param([AllowNull()][string]$Value)
    return ("'{0}'" -f ([string]$Value).Replace("'", "''"))
}

function Get-AutomaticWaveCommandText {
    param(
        [Parameter(Mandatory = $true)][string]$Campaign,
        [Parameter(Mandatory = $true)][hashtable]$Filters,
        [Parameter(Mandatory = $true)][string]$EvidenceRoot,
        [switch]$Create
    )

    $script = Join-Path $script:PackageRoot ("Campaigns\{0}\Scripts\SmartM365-EndpointMigration-{0}-New-AutomaticWave.ps1" -f $Campaign)
    $parts = @('&',(ConvertTo-PowerShellLiteral $script),'-PackageRoot',(ConvertTo-PowerShellLiteral $script:PackageRoot),'-WaveName',(ConvertTo-PowerShellLiteral $Filters.WaveName),'-EvidenceRoot',(ConvertTo-PowerShellLiteral $EvidenceRoot))
    foreach ($pair in @(@('ComputerNamePrefix','Prefixes'),@('ComputerNameContains','Contains'))) {
        $items = @($Filters[$pair[1]] | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($items.Count -gt 0) { $parts += @(('-{0}' -f $pair[0]),('@({0})' -f (($items | ForEach-Object { ConvertTo-PowerShellLiteral $_ }) -join ','))) }
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$Filters.AdDomain)) { $parts += @('-AdDomainFilter',('@({0})' -f (ConvertTo-PowerShellLiteral ([string]$Filters.AdDomain)))) }
    if ($Filters.ExcludeStaleAd) { $parts += @('-ExcludeStaleAd','-AdLastLogonMaxAgeDays',[string][int]$Filters.LastLogonDays) }
    if ($Campaign -eq 'Windows11Upgrade') {
        $parts += @('-Source',$Filters.Source)
        if ($Filters.ExcludeIntunePresent) { $parts += '-ExcludeIntunePresent' }
    }
    if ($Create) { $parts += '-Create' }
    return ($parts -join ' ')
}

function Get-InventoryRefreshPlan {
    <#
    Two independent streams started in parallel: AD (ActiveDirectory module, the longest) and Graph
    (one delegated sign-in). HybridJoin uses the combined Graph exporter (Intune, then Entra on the
    same connection); Windows11Upgrade needs Intune only.
    #>
    param([Parameter(Mandatory = $true)][ValidateSet('HybridJoin','Windows11Upgrade')][string]$Campaign)

    $ad = [pscustomobject]@{ Label = 'AD'; Script = 'SmartM365-EndpointMigration-Export-ADDevicesCsv.ps1'; Arguments = @('-ForceRefresh') }
    $graph = if ($Campaign -eq 'HybridJoin') {
        [pscustomobject]@{ Label = 'Intune and Entra'; Script = 'SmartM365-EndpointMigration-Export-AutomaticGraphInventories.ps1'; Arguments = @() }
    }
    else {
        [pscustomobject]@{ Label = 'Intune'; Script = 'SmartM365-EndpointMigration-Export-IntuneDevicesCsv.ps1'; Arguments = @('-ForceRefresh') }
    }
    return @(
        [pscustomobject]@{ Name = 'AD'; Title = 'Active Directory'; Exports = @($ad) }
        [pscustomobject]@{ Name = 'Graph'; Title = ('Microsoft Graph ({0})' -f $graph.Label); Exports = @($graph) }
    )
}

function Get-InventoryRefreshCommandText {
    <#
    Command of one hidden Windows PowerShell started by "Refresh shared inventories": every export of
    the stream runs even when a previous one fails, all output streams go to the log shown in the GUI
    window, and the exit code is 1 when at least one export failed.
    #>
    param([Parameter(Mandatory = $true)][object[]]$Exports,[Parameter(Mandatory = $true)][string]$LogPath)

    $lines = @('$ErrorActionPreference = ''Continue''','$semtFailed = 0','. {')
    foreach ($export in $Exports) {
        $path = Join-Path $script:PackageRoot ('Scripts\' + $export.Script)
        $call = (@('&',(ConvertTo-PowerShellLiteral $path)) + @($export.Arguments)) -join ' '
        $lines += ("    try {{ {0} }} catch {{ `$semtFailed++; Write-Host ('{{0:yyyy-MM-dd HH:mm:ss}} {1} export failed: {{1}}' -f (Get-Date),`$_.Exception.Message) }}" -f $call,$export.Label)
    }
    $lines += '    Write-Host (''{0:yyyy-MM-dd HH:mm:ss} Inventory refresh finished; failed exports: {1}.'' -f (Get-Date),$semtFailed)'
    $lines += ('}} *> {0}' -f (ConvertTo-PowerShellLiteral $LogPath))
    $lines += 'exit ([int]($semtFailed -gt 0))'
    return ($lines -join [Environment]::NewLine)
}

function Get-AutomaticFilterSignature {
    param([Parameter(Mandatory = $true)][string]$Campaign,[Parameter(Mandatory = $true)][hashtable]$Filters)
    return ('{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}' -f $Campaign,$Filters.WaveName,(@($Filters.Prefixes) -join ';'),(@($Filters.Contains) -join ';'),$Filters.ExcludeStaleAd,$Filters.LastLogonDays,$Filters.Source,$Filters.ExcludeIntunePresent,$Filters.AdDomain)
}

function Get-AdInventoryDomains {
    <#
    Distinct ADDomain values of an AD inventory CSV, read with a streaming CSV parser. Empty when the
    file or the column is missing.
    #>
    param([AllowNull()][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return @() }
    Add-Type -AssemblyName Microsoft.VisualBasic
    $domains = New-Object 'System.Collections.Generic.SortedSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path, [Text.Encoding]::UTF8, $true)
    try {
        $parser.SetDelimiters(',')
        $parser.HasFieldsEnclosedInQuotes = $true
        if ($parser.EndOfData) { return @() }
        $index = [Array]::IndexOf([string[]]@($parser.ReadFields() | ForEach-Object { $_.Trim() }), 'ADDomain')
        if ($index -lt 0) { return @() }
        while (-not $parser.EndOfData) {
            $fields = $parser.ReadFields()
            if ($null -ne $fields -and $fields.Count -gt $index) {
                $value = ([string]$fields[$index]).Trim().Trim('.').ToLowerInvariant()
                if ($value) { [void]$domains.Add($value) }
            }
        }
    }
    finally { $parser.Close() }
    return @($domains)
}

function Split-GuiList {
    param([AllowNull()][string]$Text)
    return @(([string]$Text -split '[;,\r\n]+') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
}

#endregion

if (-not $ValidateOnly) {
    $splashScript = Join-Path $PSScriptRoot 'SmartM365.GuiSplash.ps1'
    if (Test-Path -LiteralPath $splashScript -PathType Leaf) { . $splashScript }
}
$updateCheckScript = Join-Path $PSScriptRoot 'SmartM365.GuiUpdateCheck.ps1'
if (Test-Path -LiteralPath $updateCheckScript -PathType Leaf) { . $updateCheckScript }

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase

$splash = $null
if (-not $ValidateOnly -and (Get-Command -Name Start-SmartM365GuiSplash -ErrorAction SilentlyContinue)) {
    try { $splash = Start-SmartM365GuiSplash -Framework Wpf -ProductName 'Smart Endpoint Migration Toolkit' -LogoPath (Join-Path $script:PackageRoot 'Assets\WorkplaceCloudHub-lockup-WPF.png') -WindowIconPath (Join-Path $script:PackageRoot 'Assets\WorkplaceCloudHub.ico') } catch { $splash = $null }
}

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Smart Endpoint Migration Toolkit v1.0.0"
        Width="1240" Height="860" MinWidth="1080" MinHeight="740"
        WindowStartupLocation="CenterScreen"
        Background="#F5F8FB" FontFamily="Segoe UI" FontSize="12"
        UseLayoutRounding="True" SnapsToDevicePixels="True">
    <Window.Resources>
        <SolidColorBrush x:Key="TextBrush" Color="#1F2937"/>
        <SolidColorBrush x:Key="MutedBrush" Color="#475569"/>
        <SolidColorBrush x:Key="BorderBrush" Color="#D8E4F0"/>
        <Style TargetType="Button">
            <Setter Property="MinHeight" Value="34"/>
            <Setter Property="Padding" Value="14,6"/>
            <Setter Property="Margin" Value="4"/>
            <Setter Property="BorderBrush" Value="#CBDDEC"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style x:Key="PrimaryButtonStyle" TargetType="Button">
            <Setter Property="MinHeight" Value="34"/>
            <Setter Property="Padding" Value="14,6"/>
            <Setter Property="Margin" Value="4"/>
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderBrush" Value="#0078D4"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="MinHeight" Value="32"/>
            <Setter Property="Margin" Value="0,4,0,10"/>
            <Setter Property="BorderBrush" Value="#CBDDEC"/>
            <Setter Property="Padding" Value="8,5"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="MinHeight" Value="32"/>
            <Setter Property="Margin" Value="0,4,0,10"/>
            <Setter Property="BorderBrush" Value="#CBDDEC"/>
        </Style>
        <Style x:Key="NumericTextBoxStyle" TargetType="TextBox">
            <Setter Property="MinHeight" Value="32"/>
            <Setter Property="Margin" Value="0"/>
            <Setter Property="BorderBrush" Value="#CBDDEC"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="TextAlignment" Value="Center"/>
        </Style>
        <Style x:Key="StepperButtonStyle" TargetType="Button">
            <Setter Property="Width" Value="32"/>
            <Setter Property="MinWidth" Value="32"/>
            <Setter Property="MinHeight" Value="32"/>
            <Setter Property="Padding" Value="0"/>
            <Setter Property="Margin" Value="0"/>
            <Setter Property="BorderBrush" Value="#CBDDEC"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="Cursor" Value="Hand"/>
        </Style>
        <Style TargetType="CheckBox">
            <Setter Property="Margin" Value="0,4,18,4"/>
            <Setter Property="VerticalAlignment" Value="Center"/>
        </Style>
        <Style TargetType="TabItem">
            <Setter Property="Foreground" Value="#475569"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#D8E4F0"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="Margin" Value="0,0,6,0"/>
            <Setter Property="MinHeight" Value="34"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="TabBorder" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="8" Padding="{TemplateBinding Padding}">
                            <ContentPresenter ContentSource="Header" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="#F8FBFE"/>
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#B9DDF7"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter Property="Foreground" Value="#005A9E"/>
                                <Setter Property="FontWeight" Value="SemiBold"/>
                                <Setter TargetName="TabBorder" Property="Background" Value="#E6F4FF"/>
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#B9DDF7"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid Margin="22">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="20" Margin="0,0,0,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="190"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                    <Border Background="#E7F3FF" BorderBrush="#B5DCFF" BorderThickness="1" CornerRadius="8" Padding="10,3" HorizontalAlignment="Left" Margin="0,0,0,12">
                        <TextBlock Text="SMARTM365 DEVICES" Foreground="#005A9E" FontWeight="SemiBold"/>
                    </Border>
                    <TextBlock x:Name="ProductVersionText" Text="v1.0.0" Foreground="{StaticResource MutedBrush}" Margin="0,0,0,4"/>
                    <TextBlock Text="Smart Endpoint Migration Toolkit" Foreground="{StaticResource TextBrush}" FontSize="28" FontWeight="SemiBold"/>
                    <TextBlock Text="Hybrid Join, Intune Enrollment &amp; Windows 11 Upgrade" Foreground="{StaticResource MutedBrush}" Margin="0,6,0,0"/>
                </StackPanel>
                <Border x:Name="HeaderLogoLink" Grid.Column="1" BorderBrush="#D8E4F0" BorderThickness="1" CornerRadius="8" Background="#F8FBFF" Padding="12" Cursor="Hand" ToolTip="Open WorkplaceCloudHub.com">
                    <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                        <Image x:Name="HeaderLogoImage" Width="128" Height="72" Stretch="Uniform" SnapsToDevicePixels="True" RenderOptions.BitmapScalingMode="HighQuality"/>
                        <TextBlock Text="workplacecloudhub.com" Foreground="#475569" FontSize="11" HorizontalAlignment="Center" Margin="0,4,0,0"/>
                    </StackPanel>
                </Border>
            </Grid>
        </Border>

        <Border Grid.Row="1" CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="14" Margin="0,0,0,12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="300"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Margin="0,0,16,0">
                    <TextBlock Text="Campaign" FontWeight="SemiBold"/>
                    <ComboBox x:Name="CampaignCombo" Margin="0,4,0,0"/>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Ellipse Width="10" Height="10" Fill="#0078D4" Margin="2,0,12,0"/>
                    <StackPanel>
                        <TextBlock x:Name="StatusTitle" Text="Ready" FontWeight="SemiBold" Foreground="#0078D4"/>
                        <TextBlock x:Name="StatusText" Text="Select a Campaign and a Wave." Foreground="{StaticResource MutedBrush}" Margin="0,3,0,0" TextWrapping="Wrap"/>
                    </StackPanel>
                </StackPanel>
                <StackPanel Grid.Column="2" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button x:Name="RefreshButton" Content="Refresh"/>
                    <Button x:Name="StopAllButton" Content="Stop running Waves" Background="#FFF4E5" Foreground="#9A3412" BorderBrush="#FDBA74" MinWidth="150"/>
                    <Button x:Name="LaunchAllButton" Content="Launch All Waves" Style="{StaticResource PrimaryButtonStyle}" MinWidth="150"/>
                </StackPanel>
            </Grid>
        </Border>

        <TabControl Grid.Row="2" x:Name="MainTabs" Background="Transparent" BorderThickness="0" Padding="0">
            <TabItem x:Name="WavesTab" Header="Waves">
                <Grid Margin="0,12,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="1.05*"/>
                        <ColumnDefinition Width="0.95*"/>
                    </Grid.ColumnDefinitions>
                    <Border Grid.Column="0" CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="18" Margin="0,0,8,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <TextBlock Text="Wave" Foreground="{StaticResource TextBrush}" FontSize="18" FontWeight="SemiBold"/>
                            <Grid Grid.Row="1" Margin="0,14,0,0">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="160"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel>
                                    <TextBlock Text="Existing Waves of the selected Campaign"/>
                                    <ComboBox x:Name="WaveCombo"/>
                                </StackPanel>
                                <Button Grid.Column="1" x:Name="LaunchWaveButton" Content="Launch Wave" Style="{StaticResource PrimaryButtonStyle}" Margin="12,22,0,10"/>
                            </Grid>
                            <UniformGrid Grid.Row="2" Columns="4" Margin="0,4,0,12">
                                <StackPanel>
                                    <TextBlock Text="Computers" FontWeight="SemiBold"/>
                                    <TextBlock x:Name="WaveDeviceCountText" Text="-" Foreground="{StaticResource MutedBrush}" Margin="0,4,0,0"/>
                                </StackPanel>
                                <StackPanel>
                                    <TextBlock Text="Source" FontWeight="SemiBold"/>
                                    <TextBlock x:Name="WaveSourceText" Text="-" Foreground="{StaticResource MutedBrush}" Margin="0,4,0,0" TextWrapping="Wrap"/>
                                </StackPanel>
                                <StackPanel>
                                    <TextBlock Text="Launchers" FontWeight="SemiBold"/>
                                    <TextBlock x:Name="WaveWrappersText" Text="-" Foreground="{StaticResource MutedBrush}" Margin="0,4,0,0" TextWrapping="Wrap"/>
                                </StackPanel>
                                <StackPanel>
                                    <TextBlock Text="Last Run" FontWeight="SemiBold"/>
                                    <TextBlock x:Name="WaveLastRunText" Text="-" Foreground="{StaticResource MutedBrush}" Margin="0,4,0,0" TextWrapping="Wrap"/>
                                </StackPanel>
                            </UniformGrid>
                            <StackPanel Grid.Row="3">
                                <TextBlock Text="Run mode"/>
                                <ComboBox x:Name="WaveModeCombo"/>
                            </StackPanel>
                            <WrapPanel Grid.Row="4" VerticalAlignment="Bottom">
                                <Button x:Name="OpenWaveFolderButton" Content="Folder"/>
                                <Button x:Name="OpenWaveComputersButton" Content="Computers.txt"/>
                                <Button x:Name="OpenWaveConfigButton" Content="Wave options"/>
                                <Button x:Name="RefreshWrappersButton" Content="Refresh launchers"/>
                            </WrapPanel>
                        </Grid>
                    </Border>
                    <Border Grid.Column="1" CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="18" Margin="8,0,0,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <TextBlock Text="Activity" Foreground="{StaticResource TextBrush}" FontSize="18" FontWeight="SemiBold"/>
                            <TextBox Grid.Row="1" x:Name="ActivityText" Margin="0,14,0,0" IsReadOnly="True" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>

            <TabItem x:Name="SinglePcTab" Header="Single PC">
                <Border CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="18" Margin="0,12,0,0">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Margin="0,0,14,0">
                            <TextBlock Text="Single PC" FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}" Margin="0,0,0,14"/>
                            <TextBlock Text="Computer name (short name or FQDN)"/>
                            <TextBox x:Name="SingleComputerText"/>
                            <TextBlock Text="Run mode"/>
                            <ComboBox x:Name="SingleModeCombo"/>
                            <TextBlock Text="A Single PC Run uses a one-computer Wave created inside its Run folder and the same engine, options and protections as a Wave." Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap" Margin="0,0,0,10"/>
                            <Button x:Name="LaunchSingleButton" Content="Launch Single PC" Style="{StaticResource PrimaryButtonStyle}" HorizontalAlignment="Left" MinWidth="170"/>
                        </StackPanel>
                        <StackPanel Grid.Column="1" Margin="14,0,0,0">
                            <TextBlock Text="Single PC Runs" FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}" Margin="0,0,0,14"/>
                            <TextBox x:Name="SingleRunFolderText" IsReadOnly="True"/>
                            <Button x:Name="OpenSingleRunFolderButton" Content="Open folder" HorizontalAlignment="Left"/>
                        </StackPanel>
                    </Grid>
                </Border>
            </TabItem>

            <TabItem x:Name="NewWaveTab" Header="New Wave">
                <Border CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="18" Margin="0,12,0,0">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <StackPanel Margin="0,0,14,0">
                            <TextBlock Text="New Wave" FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}" Margin="0,0,0,14"/>
                            <TextBlock Text="Wave name (letters, digits, dot, dash, underscore)"/>
                            <TextBox x:Name="NewWaveNameText"/>
                            <TextBlock Text="Computers (one per line, optional)"/>
                            <TextBox x:Name="NewWaveComputersText" AcceptsReturn="True" Height="220" VerticalScrollBarVisibility="Auto" TextWrapping="NoWrap"/>
                            <TextBlock Text="Create only prepares the Wave folder, Computers.txt, Wave.local.json and the Loop/Once launchers. It never launches the Wave." Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap" Margin="0,0,0,10"/>
                            <Button x:Name="CreateWaveButton" Content="Create Wave" Style="{StaticResource PrimaryButtonStyle}" HorizontalAlignment="Left" MinWidth="140"/>
                        </StackPanel>
                        <StackPanel Grid.Column="1" Margin="14,0,0,0">
                            <TextBlock Text="Created Wave" FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}" Margin="0,0,0,14"/>
                            <TextBox x:Name="NewWavePathText" IsReadOnly="True"/>
                            <Button x:Name="OpenNewWaveComputersButton" Content="Open Computers.txt" HorizontalAlignment="Left"/>
                        </StackPanel>
                    </Grid>
                </Border>
            </TabItem>

            <TabItem x:Name="AutomaticWaveTab" Header="Automatic Wave">
                <Grid Margin="0,12,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="0.95*"/>
                        <ColumnDefinition Width="1.05*"/>
                    </Grid.ColumnDefinitions>
                    <Border Grid.Column="0" CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="16" Margin="0,0,8,0">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel>
                                <TextBlock Text="Selection" FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}" Margin="0,0,0,10"/>
                                <TextBlock x:Name="AutomaticInventoryStatusText" Text="-" Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap" Margin="0,0,0,8"/>
                                <Button x:Name="AutomaticRefreshInventoryButton" Content="Refresh shared inventories" HorizontalAlignment="Left"/>
                                <TextBlock Text="Wave name" Margin="0,8,0,0"/>
                                <TextBox x:Name="AutomaticWaveNameText"/>
                                <TextBlock Text="Computer name prefixes (separated by ;)"/>
                                <TextBox x:Name="AutomaticNamePrefixText"/>
                                <TextBlock Text="Computer name contains (separated by ;)"/>
                                <TextBox x:Name="AutomaticNameContainsText"/>
                                <TextBlock Text="AD domain (from the shared AD inventory)"/>
                                <ComboBox x:Name="AutomaticAdDomainCombo"/>
                                <StackPanel x:Name="AutomaticSourcePanel">
                                    <TextBlock Text="Inventory source"/>
                                    <ComboBox x:Name="AutomaticSourceCombo"/>
                                    <CheckBox x:Name="AutomaticExcludeIntuneCheck" Content="Exclude devices present in Intune"/>
                                </StackPanel>
                                <CheckBox x:Name="AutomaticExcludeStaleAdCheck" Content="Exclude computers without a recent AD logon"/>
                                <TextBlock Text="Maximum AD logon age (days)" Margin="0,6,0,0"/>
                                <Grid Margin="0,4,0,10" Width="160" HorizontalAlignment="Left">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="32"/>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="32"/>
                                    </Grid.ColumnDefinitions>
                                    <Button x:Name="AutomaticLastLogonDaysDownButton" Grid.Column="0" Content="-" Style="{StaticResource StepperButtonStyle}"/>
                                    <TextBox x:Name="AutomaticLastLogonDaysText" Grid.Column="1" Text="45" Style="{StaticResource NumericTextBoxStyle}"/>
                                    <Button x:Name="AutomaticLastLogonDaysUpButton" Grid.Column="2" Content="+" Style="{StaticResource StepperButtonStyle}"/>
                                </Grid>
                                <WrapPanel>
                                    <Button x:Name="AutomaticPreviewButton" Content="Preview" MinWidth="110"/>
                                    <Button x:Name="AutomaticCreateButton" Content="Create Wave" Style="{StaticResource PrimaryButtonStyle}" MinWidth="130"/>
                                </WrapPanel>
                                <TextBlock Text="Preview never creates anything. Create is available only for the previewed filters and never launches the Wave." Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap" Margin="0,6,0,0"/>
                            </StackPanel>
                        </ScrollViewer>
                    </Border>
                    <Border Grid.Column="1" CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="16" Margin="8,0,0,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>
                            <TextBlock Text="Preview" FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}"/>
                            <TextBox Grid.Row="1" x:Name="AutomaticSummaryText" Margin="0,12,0,0" IsReadOnly="True" AcceptsReturn="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
                            <Grid Grid.Row="2">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBox x:Name="AutomaticEvidencePathText" IsReadOnly="True" Margin="0,8,8,0"/>
                                <Button Grid.Column="1" x:Name="AutomaticOpenEvidenceButton" Content="Evidence"/>
                            </Grid>
                        </Grid>
                    </Border>
                </Grid>
            </TabItem>

            <TabItem x:Name="RunsTab" Header="Runs">
                <Border CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="16" Margin="0,12,0,0">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <StackPanel>
                            <TextBlock Text="Runs and reports" FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}"/>
                            <TextBlock x:Name="ActiveRunsText" Text="-" Foreground="{StaticResource MutedBrush}" Margin="0,6,0,8" TextWrapping="Wrap"/>
                        </StackPanel>
                        <DataGrid Grid.Row="1" x:Name="RunsGrid" AutoGenerateColumns="False" IsReadOnly="True" SelectionMode="Single" HeadersVisibility="Column" GridLinesVisibility="Horizontal" BorderBrush="#D8E4F0">
                            <DataGrid.Columns>
                                <DataGridTextColumn Header="Wave" Binding="{Binding WaveName}" Width="160"/>
                                <DataGridTextColumn Header="Run" Binding="{Binding RunId}" Width="170"/>
                                <DataGridTextColumn Header="Status" Binding="{Binding Status}" Width="130"/>
                                <DataGridTextColumn Header="Cycles" Binding="{Binding Cycles}" Width="60"/>
                                <DataGridTextColumn Header="Started (UTC)" Binding="{Binding StartedUtc}" Width="170"/>
                                <DataGridTextColumn Header="Updated (UTC)" Binding="{Binding UpdatedUtc}" Width="170"/>
                                <DataGridTextColumn Header="Technician" Binding="{Binding Technician}" Width="*"/>
                            </DataGrid.Columns>
                        </DataGrid>
                        <WrapPanel Grid.Row="2" Margin="0,8,0,0">
                            <Button x:Name="RefreshRunsButton" Content="Refresh"/>
                            <Button x:Name="OpenRunReportButton" Content="Open report"/>
                            <Button x:Name="OpenRunFolderButton" Content="Open Run folder"/>
                            <Button x:Name="StopSelectedRunButton" Content="Controlled stop" Background="#FFF4E5" Foreground="#9A3412" BorderBrush="#FDBA74"/>
                            <Button x:Name="ForceStopSelectedRunButton" Content="Force local stop" Background="#FDECEC" Foreground="#991B1B" BorderBrush="#FCA5A5"/>
                        </WrapPanel>
                    </Grid>
                </Border>
            </TabItem>

            <TabItem x:Name="OptionsTab" Header="Options">
                <Border CornerRadius="8" BorderBrush="{StaticResource BorderBrush}" BorderThickness="1" Background="White" Padding="16" Margin="0,12,0,0">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <StackPanel>
                            <TextBlock x:Name="OptionsCampaignText" Text="Options" FontSize="18" FontWeight="SemiBold" Foreground="{StaticResource TextBrush}"/>
                            <TextBlock Text="These choices are captured for each Run of this Campaign and win over the Wave and global configuration. Empty text fields keep the Wave or global value. Hover an option for its unit, default and environment variable." Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap" Margin="0,6,0,8"/>
                        </StackPanel>
                        <ScrollViewer Grid.Row="1" VerticalScrollBarVisibility="Auto">
                            <StackPanel x:Name="OptionsHost"/>
                        </ScrollViewer>
                        <WrapPanel Grid.Row="2" Margin="0,8,0,0">
                            <Button x:Name="SaveOptionsButton" Content="Save choices"/>
                            <Button x:Name="ResetOptionsButton" Content="Reset to defaults"/>
                        </WrapPanel>
                    </Grid>
                </Border>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
'@

$reader = [System.Xml.XmlReader]::Create([System.IO.StringReader]::new($xaml))
$window = [Windows.Markup.XamlReader]::Load($reader)
$controls = @{}
foreach ($match in [regex]::Matches($xaml, 'x:Name="(?<Name>[A-Za-z][A-Za-z0-9_]*)"')) {
    $name = $match.Groups['Name'].Value
    $control = $window.FindName($name)
    if ($control) { $controls[$name] = $control }
}
$controls.ProductVersionText.Text = 'v{0}' -f $script:GuiVersion

$iconPath = Join-Path $script:PackageRoot 'Assets\WorkplaceCloudHub.ico'
if (Test-Path -LiteralPath $iconPath -PathType Leaf) { try { $window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]$iconPath) } catch { } }
$logoPath = Join-Path $script:PackageRoot 'Assets\WorkplaceCloudHub-lockup-WPF.png'
if (Test-Path -LiteralPath $logoPath -PathType Leaf) {
    try {
        $bitmap = [System.Windows.Media.Imaging.BitmapImage]::new()
        $bitmap.BeginInit(); $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad; $bitmap.UriSource = [Uri]$logoPath; $bitmap.EndInit(); $bitmap.Freeze()
        $controls.HeaderLogoImage.Source = $bitmap
    }
    catch { }
}

#region UI helpers

function Add-Status {
    param([string]$Message,[string]$Title = 'Ready')
    $controls.StatusTitle.Text = $Title
    $controls.StatusText.Text = $Message
    $controls.ActivityText.AppendText(('{0:yyyy-MM-dd HH:mm:ss}  {1}' -f (Get-Date),$Message) + [Environment]::NewLine)
    $controls.ActivityText.ScrollToEnd()
}

function Show-GuiError {
    param([string]$Message)
    [System.Windows.MessageBox]::Show($window, $Message, 'Smart Endpoint Migration Toolkit', 'OK', 'Error') | Out-Null
}

function Show-GuiQuestion {
    param([string]$Message,[string]$Title = 'Confirm')
    return ([System.Windows.MessageBox]::Show($window, $Message, $Title, 'YesNo', 'Question') -eq [System.Windows.MessageBoxResult]::Yes)
}

function Open-ShellPath {
    param([AllowNull()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { Show-GuiError ("Path not found: {0}" -f $Path); return }
    Start-Process -FilePath $Path | Out-Null
}

function Open-TextFile {
    param([AllowNull()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { Show-GuiError ("File not found: {0}" -f $Path); return }
    Start-Process -FilePath 'notepad.exe' -ArgumentList @(('"{0}"' -f $Path)) | Out-Null
}

function Get-SelectedCampaign {
    $item = $controls.CampaignCombo.SelectedItem
    if ($null -eq $item) { return 'HybridJoin' }
    return [string]$item.Tag
}

function Register-NumericStepper {
    param($TextBox,$DownButton,$UpButton,[int]$Minimum,[int]$Maximum,[int]$Step = 1)
    $state = [pscustomobject]@{ TextBox = $TextBox; Minimum = $Minimum; Maximum = $Maximum; Step = $Step }
    $handler = {
        param($eventSource,$routedEventArgs)
        $target = $eventSource.Tag
        $value = 0
        if (-not [int]::TryParse([string]$target.TextBox.Text, [ref]$value)) { $value = $target.Minimum }
        if ([string]$eventSource.Content -eq '+') { $value += $target.Step } else { $value -= $target.Step }
        if ($value -lt $target.Minimum) { $value = $target.Minimum }
        if ($value -gt $target.Maximum) { $value = $target.Maximum }
        $target.TextBox.Text = [string]$value
    }
    $DownButton.Tag = $state; $UpButton.Tag = $state
    $DownButton.Add_Click($handler); $UpButton.Add_Click($handler)
}

function Add-SemtGuiOptionControls {
    <#
    Builds one editor per GUI-visible option of the Campaign, grouped by category: check box
    for booleans, numeric stepper for integers, combo box for choices, text box for strings and
    paths. The demo profile reuses this function so demonstrations show the real options.
    #>
    param([Parameter(Mandatory = $true)]$OptionsHost,[Parameter(Mandatory = $true)]$Catalog,[Parameter(Mandatory = $true)][string]$Campaign,[hashtable]$Values = @{})

    $OptionsHost.Children.Clear()
    $editors = @{}
    $definitions = @(Get-SemtCampaignOptionDefinitions -Catalog $Catalog -Campaign $Campaign | Where-Object { $_.GuiVisible })
    $categories = @($script:CategoryOrder) + @($definitions | ForEach-Object { [string]$_.Category } | Where-Object { $script:CategoryOrder -notcontains $_ } | Select-Object -Unique)
    foreach ($category in $categories) {
        $items = @($definitions | Where-Object { [string]$_.Category -eq $category })
        if ($items.Count -eq 0) { continue }
        $group = New-Object System.Windows.Controls.GroupBox
        $group.Header = $category
        $group.Margin = [System.Windows.Thickness]::new(0,0,0,10)
        $group.Padding = [System.Windows.Thickness]::new(8)
        $panel = New-Object System.Windows.Controls.WrapPanel
        foreach ($definition in $items) {
            $name = [string]$definition.Name
            $default = Get-SemtOptionDefault -Definition $definition -Campaign $Campaign
            $value = if ($Values.ContainsKey($name)) { $Values[$name] } else { $default }
            $tip = '{0}{1}Default: {2}{3}{1}Option: {4}; environment: {5}' -f $definition.Description,[Environment]::NewLine,$default,$(if ($definition.Unit) { ' ' + $definition.Unit } else { '' }),$name,$definition.EnvVar
            $cell = New-Object System.Windows.Controls.StackPanel
            $cell.Width = 340
            $cell.Margin = [System.Windows.Thickness]::new(0,0,14,6)
            $cell.ToolTip = $tip
            $control = $null
            switch ([string]$definition.Type) {
                'Boolean' {
                    $control = New-Object System.Windows.Controls.CheckBox
                    $control.Content = [string]$definition.Label
                    $control.IsChecked = [bool]$value
                    $control.ToolTip = $tip
                    [void]$cell.Children.Add($control)
                }
                'Integer' {
                    $label = New-Object System.Windows.Controls.TextBlock
                    $label.Text = if ($definition.Unit) { '{0} ({1})' -f $definition.Label,$definition.Unit } else { [string]$definition.Label }
                    $label.TextWrapping = 'Wrap'
                    [void]$cell.Children.Add($label)
                    $grid = New-Object System.Windows.Controls.Grid
                    $grid.Width = 170; $grid.HorizontalAlignment = 'Left'; $grid.Margin = [System.Windows.Thickness]::new(0,4,0,4)
                    foreach ($width in @(32,0,32)) { $column = New-Object System.Windows.Controls.ColumnDefinition; $column.Width = if ($width -gt 0) { [System.Windows.GridLength]::new($width) } else { [System.Windows.GridLength]::new(1,'Star') }; [void]$grid.ColumnDefinitions.Add($column) }
                    $down = New-Object System.Windows.Controls.Button; $down.Content = '-'; $down.Style = $window.Resources['StepperButtonStyle']
                    $text = New-Object System.Windows.Controls.TextBox; $text.Text = [string]$value; $text.Style = $window.Resources['NumericTextBoxStyle']; [System.Windows.Controls.Grid]::SetColumn($text, 1)
                    $up = New-Object System.Windows.Controls.Button; $up.Content = '+'; $up.Style = $window.Resources['StepperButtonStyle']; [System.Windows.Controls.Grid]::SetColumn($up, 2)
                    [void]$grid.Children.Add($down); [void]$grid.Children.Add($text); [void]$grid.Children.Add($up)
                    $minimum = [int](Get-SemtPropertyValue -InputObject $definition -Name 'Minimum' -Default 0)
                    $maximum = [int](Get-SemtPropertyValue -InputObject $definition -Name 'Maximum' -Default 100000)
                    $range = $maximum - $minimum
                    $step = if ($range -ge 2000) { 10 } else { 1 }
                    Register-NumericStepper -TextBox $text -DownButton $down -UpButton $up -Minimum $minimum -Maximum $maximum -Step $step
                    [void]$cell.Children.Add($grid)
                    $control = $text
                }
                'Choice' {
                    $label = New-Object System.Windows.Controls.TextBlock; $label.Text = [string]$definition.Label; [void]$cell.Children.Add($label)
                    $control = New-Object System.Windows.Controls.ComboBox
                    foreach ($choice in @($definition.Choices)) { [void]$control.Items.Add([string]$choice) }
                    $control.SelectedItem = [string]$value
                    [void]$cell.Children.Add($control)
                }
                default {
                    $label = New-Object System.Windows.Controls.TextBlock; $label.Text = [string]$definition.Label; $label.TextWrapping = 'Wrap'; [void]$cell.Children.Add($label)
                    $control = New-Object System.Windows.Controls.TextBox
                    $control.Text = [string]$value
                    [void]$cell.Children.Add($control)
                }
            }
            $editors[$name] = [pscustomobject]@{ Definition = $definition; Control = $control }
            [void]$panel.Children.Add($cell)
        }
        $group.Content = $panel
        [void]$OptionsHost.Children.Add($group)
    }
    return $editors
}

function Get-SemtGuiOptionValues {
    param([Parameter(Mandatory = $true)][hashtable]$Editors)
    $values = @{}
    foreach ($entry in $Editors.GetEnumerator()) {
        $definition = $entry.Value.Definition
        $control = $entry.Value.Control
        $raw = switch ([string]$definition.Type) {
            'Boolean' { [bool]$control.IsChecked }
            'Choice' { [string]$control.SelectedItem }
            default { [string]$control.Text }
        }
        $values[$entry.Key] = ConvertTo-SemtOptionValue -Definition $definition -Value $raw -Source ('GUI field "{0}"' -f $definition.Label)
    }
    return $values
}

#endregion

$script:Catalog = Import-SemtOptionCatalog -PackageRoot $script:PackageRoot
$script:Editors = @{}
$script:CurrentCampaign = ''
$script:Waves = @()
$script:SelectedWave = $null
$script:LaunchQueue = New-Object System.Collections.Queue
$script:LaunchTimer = $null
$script:AutomaticPreviewSignature = ''
$script:AutomaticProcess = $null
$script:AutomaticTimer = $null

foreach ($campaign in $script:CampaignTitles.Keys) {
    $item = New-Object System.Windows.Controls.ComboBoxItem
    $item.Content = $script:CampaignTitles[$campaign]
    $item.Tag = $campaign
    [void]$controls.CampaignCombo.Items.Add($item)
}
foreach ($combo in @($controls.WaveModeCombo,$controls.SingleModeCombo)) { foreach ($mode in $script:RunModes.Keys) { [void]$combo.Items.Add($mode) } }
$controls.WaveModeCombo.SelectedIndex = 0
$controls.SingleModeCombo.SelectedIndex = 1
foreach ($source in @('Both','AD','Intune')) { [void]$controls.AutomaticSourceCombo.Items.Add($source) }
$controls.AutomaticSourceCombo.SelectedIndex = 0
Register-NumericStepper -TextBox $controls.AutomaticLastLogonDaysText -DownButton $controls.AutomaticLastLogonDaysDownButton -UpButton $controls.AutomaticLastLogonDaysUpButton -Minimum 1 -Maximum 3650

function Save-CurrentCampaignState {
    param([switch]$Quiet)
    if ($ValidateOnly -or [string]::IsNullOrWhiteSpace($script:CurrentCampaign) -or $script:Editors.Count -eq 0) { return $true }
    try {
        $values = Get-SemtGuiOptionValues -Editors $script:Editors
        Save-GuiState -Campaign $script:CurrentCampaign -Values $values
        return $true
    }
    catch {
        if (-not $Quiet) { Show-GuiError $_.Exception.Message }
        return $false
    }
}

function Update-WaveView {
    $wave = $script:SelectedWave
    if ($null -eq $wave) {
        foreach ($name in @('WaveDeviceCountText','WaveSourceText','WaveWrappersText','WaveLastRunText')) { $controls[$name].Text = '-' }
        $controls.LaunchWaveButton.IsEnabled = $false
        return
    }
    $controls.WaveDeviceCountText.Text = [string]$wave.ComputerCount
    $source = 'Manual'
    try {
        $config = Read-SemtJsonFile -Path (Join-Path $wave.Path 'Wave.local.json')
        if ($config) { $source = ('{0} {1}' -f (Get-SemtPropertyValue -InputObject $config -Name 'Source' -Default 'Manual'),(Get-SemtPropertyValue -InputObject $config -Name 'Description' -Default '')).Trim() }
    }
    catch { $source = 'Wave.local.json unreadable' }
    $controls.WaveSourceText.Text = $source
    $controls.WaveWrappersText.Text = if ($wave.WrappersReady) { 'Ready' } else { 'Missing: ' + $wave.MissingWrappers }
    $last = @(Get-SemtRuns -PackageRoot $script:PackageRoot -Campaign $script:CurrentCampaign -Last 200 | Where-Object { $_.WaveName -eq $wave.Name } | Select-Object -First 1)
    $controls.WaveLastRunText.Text = if ($last.Count -gt 0) { '{0} ({1})' -f $last[0].Status,$last[0].UpdatedUtc } else { 'Never' }
    $controls.LaunchWaveButton.IsEnabled = [bool]$wave.Launchable
}

function Update-WaveList {
    param([string]$SelectName)
    $script:Waves = @(Get-SemtWaves -PackageRoot $script:PackageRoot -Campaign $script:CurrentCampaign)
    $controls.WaveCombo.Items.Clear()
    foreach ($wave in $script:Waves) { [void]$controls.WaveCombo.Items.Add(('{0}  ({1} computer(s))' -f $wave.Name,$wave.ComputerCount)) }
    $index = 0
    if (-not [string]::IsNullOrWhiteSpace($SelectName)) { for ($i = 0; $i -lt $script:Waves.Count; $i++) { if ($script:Waves[$i].Name -eq $SelectName) { $index = $i } } }
    if ($script:Waves.Count -gt 0) { $controls.WaveCombo.SelectedIndex = $index } else { $script:SelectedWave = $null; Update-WaveView }
    $hasLaunchable = (@($script:Waves | Where-Object { $_.Launchable }).Count -gt 0)
    $controls.LaunchAllButton.IsEnabled = $hasLaunchable
    $controls.WaveCombo.IsEnabled = ($script:Waves.Count -gt 0)
}

function Update-RunsView {
    $runs = @(Get-SemtRuns -PackageRoot $script:PackageRoot -Campaign $script:CurrentCampaign -Last 100)
    $controls.RunsGrid.ItemsSource = $runs
    $active = @(Get-SemtActiveRuns -RunsRoot (Join-Path $script:PackageRoot 'Runs') -Campaign $script:CurrentCampaign | Where-Object { $_.Alive })
    $controls.ActiveRunsText.Text = if ($active.Count -eq 0) { 'No Run of this Campaign is active on this technician computer.' } else { '{0} active Run(s): {1}' -f $active.Count,((@($active) | ForEach-Object { '{0} [{1}]' -f $_.WaveName,$_.Status }) -join ', ') }
    $singleRoot = Join-Path $script:PackageRoot ("Runs\{0}\_SinglePC" -f $script:CurrentCampaign)
    $controls.SingleRunFolderText.Text = $singleRoot
}

function Update-AutomaticView {
    $isWindows11 = ($script:CurrentCampaign -eq 'Windows11Upgrade')
    $controls.AutomaticSourcePanel.Visibility = if ($isWindows11) { 'Visible' } else { 'Collapsed' }
    $kinds = if ($isWindows11) { @('AD','Intune') } else { @('AD','Intune','Entra') }
    $parts = foreach ($kind in $kinds) {
        $maxAge = if ($kind -eq 'AD') { 12 } else { 2 }
        $check = Test-SemtSharedInventoryUsable -Path (Get-SemtSharedInventoryPath -PackageRoot $script:PackageRoot -Kind $kind) -Kind $kind -MaxAgeHours $maxAge
        '{0}: {1}' -f $kind,$(if ($null -ne $check.AgeHours) { '{0} ({1} h)' -f $check.Reason,$check.AgeHours } else { $check.Reason })
    }
    $controls.AutomaticInventoryStatusText.Text = 'Shared inventories - ' + ($parts -join '; ')
    # AD domain list: read again only when the shared AD inventory file changed.
    $adPath = Get-SemtSharedInventoryPath -PackageRoot $script:PackageRoot -Kind AD
    $adFile = Get-Item -LiteralPath $adPath -ErrorAction SilentlyContinue
    $adStamp = if ($adFile) { '{0}|{1}' -f $adFile.LastWriteTimeUtc.Ticks,$adFile.Length } else { '' }
    if ($adStamp -ne $script:AdDomainStamp) {
        $script:AdDomains = @(Get-AdInventoryDomains -Path $adPath)
        $script:AdDomainStamp = $adStamp
    }
    $previousDomain = if ($controls.AutomaticAdDomainCombo.SelectedIndex -gt 0) { [string]$controls.AutomaticAdDomainCombo.SelectedItem } else { '' }
    $controls.AutomaticAdDomainCombo.Items.Clear()
    [void]$controls.AutomaticAdDomainCombo.Items.Add($script:AllAdDomainsLabel)
    foreach ($domain in @($script:AdDomains)) { [void]$controls.AutomaticAdDomainCombo.Items.Add($domain) }
    $controls.AutomaticAdDomainCombo.SelectedIndex = [Math]::Max(0, $controls.AutomaticAdDomainCombo.Items.IndexOf($previousDomain))
    if ([string]::IsNullOrWhiteSpace($controls.AutomaticWaveNameText.Text) -or $controls.AutomaticWaveNameText.Text -like 'WAVE-AUTO-*') {
        $prefix = if ($isWindows11) { 'W10' } else { 'HJ' }
        $controls.AutomaticWaveNameText.Text = 'WAVE-AUTO-{0}-{1}' -f $prefix,(Get-Date -Format 'yyyyMMdd-HHmmss')
    }
    $script:AutomaticPreviewSignature = ''
    $controls.AutomaticCreateButton.IsEnabled = $false
}

function Switch-Campaign {
    $new = Get-SelectedCampaign
    if ($new -eq $script:CurrentCampaign) { return }
    if (-not [string]::IsNullOrWhiteSpace($script:CurrentCampaign)) { [void](Save-CurrentCampaignState -Quiet) }
    $script:CurrentCampaign = $new
    $values = Get-GuiInitialOptionValues -Catalog $script:Catalog -Campaign $new
    $script:Editors = Add-SemtGuiOptionControls -OptionsHost $controls.OptionsHost -Catalog $script:Catalog -Campaign $new -Values $values
    $controls.OptionsCampaignText.Text = 'Options - {0}' -f $script:CampaignTitles[$new]
    Update-WaveList
    Update-RunsView
    Update-AutomaticView
    Add-Status -Message ('Campaign selected: {0}. Options and selections are kept separately for each Campaign.' -f $script:CampaignTitles[$new])
}

function Start-GuiRun {
    param([string]$WavePath,[string]$ComputerName,[Parameter(Mandatory = $true)][string]$ModeKey,[Parameter(Mandatory = $true)][string]$TargetText,[switch]$SkipConfirmation)

    $campaign = $script:CurrentCampaign
    $values = Get-SemtGuiOptionValues -Editors $script:Editors
    $mode = $script:RunModes[$ModeKey]
    $runMode = [string]$mode[0]
    $ignoreGuard = [bool]$mode[1]
    if (-not $SkipConfirmation) {
        $summary = Get-GuiLaunchSummary -Catalog $script:Catalog -Campaign $campaign -Values $values -Target $TargetText -RunMode $runMode -IgnoreRunGuard $ignoreGuard
        if (-not (Show-GuiQuestion -Message ($summary + [Environment]::NewLine + [Environment]::NewLine + 'Launch now in an elevated console?') -Title 'Confirm launch')) { return $false }
        if ((Test-UnlimitedCycleConfirmationRequired -RunMode $runMode -MaxCycles ([int]$values['MaxCycles'])) -and -not (Show-GuiQuestion -Message 'Loop mode with no cycle limit repeats until every computer is complete or the Run is stopped. Continue?' -Title 'Unlimited cycles')) { return $false }
    }
    Save-GuiState -Campaign $campaign -Values $values
    $selection = New-GuiSelectionFile -Catalog $script:Catalog -Campaign $campaign -Values $values
    $command = Get-GuiRunCommand -Campaign $campaign -WavePath $WavePath -ComputerName $ComputerName -RunMode $runMode -IgnoreRunGuard $ignoreGuard -SelectionPath $selection
    $title = 'SEMT {0} - {1} - {2}' -f $campaign,$TargetText,$runMode
    $launcher = New-GuiLaunchCommandFile -WorkingDirectory (Join-Path $script:PackageRoot 'Scripts') -Commands @($command) -NamePrefix ('SEMT-{0}' -f $campaign) -WindowTitle $title -EvidenceDirectory (Split-Path $selection -Parent) -PauseWhenDone
    Start-GuiLaunchCommandFile -LaunchCommandPath $launcher -WorkingDirectory (Join-Path $script:PackageRoot 'Scripts')
    Add-Status -Title 'Launched' -Message ('{0}: elevated console started. Choices captured in {1}' -f $TargetText,$selection)
    return $true
}

function Invoke-LaunchQueueTick {
    if ($script:LaunchQueue.Count -eq 0) {
        if ($script:LaunchTimer) { $script:LaunchTimer.Stop() }
        Add-Status -Message 'Launch All Waves: every queued Wave was started.'
        return
    }
    $wave = $script:LaunchQueue.Dequeue()
    try { [void](Start-GuiRun -WavePath $wave.Path -ModeKey ([string]$controls.WaveModeCombo.SelectedItem) -TargetText ("Wave '{0}'" -f $wave.Name) -SkipConfirmation) }
    catch { Add-Status -Title 'Launch failed' -Message ('{0}: {1}' -f $wave.Name,$_.Exception.Message) }
}

function Invoke-AutomaticWave {
    param([switch]$Create)

    $campaign = $script:CurrentCampaign
    $filters = @{
        WaveName = $controls.AutomaticWaveNameText.Text.Trim()
        Prefixes = @(Split-GuiList $controls.AutomaticNamePrefixText.Text)
        Contains = @(Split-GuiList $controls.AutomaticNameContainsText.Text)
        ExcludeStaleAd = [bool]$controls.AutomaticExcludeStaleAdCheck.IsChecked
        LastLogonDays = [int]$controls.AutomaticLastLogonDaysText.Text
        Source = [string]$controls.AutomaticSourceCombo.SelectedItem
        AdDomain = if ($controls.AutomaticAdDomainCombo.SelectedIndex -gt 0) { [string]$controls.AutomaticAdDomainCombo.SelectedItem } else { '' }
        ExcludeIntunePresent = [bool]$controls.AutomaticExcludeIntuneCheck.IsChecked
    }
    if (-not (Test-SemtWaveName -Name $filters.WaveName)) { throw ("'{0}' is not a valid Wave name." -f $filters.WaveName) }
    $signature = Get-AutomaticFilterSignature -Campaign $campaign -Filters $filters
    if ($Create -and $signature -ne $script:AutomaticPreviewSignature) { throw 'The filters changed since the last preview. Run Preview again before Create.' }
    if ($Create -and -not (Show-GuiQuestion -Message ("Create Wave '{0}' from the previewed selection?{1}{1}The Wave is created but NOT launched." -f $filters.WaveName,[Environment]::NewLine) -Title 'Create Wave')) { return }
    $evidence = Join-Path $script:PackageRoot ("Runs\{0}\_AutomaticWave\{1}" -f $campaign,(Get-Date -Format 'yyyyMMdd-HHmmss'))
    New-SemtDirectory -Path $evidence
    $commandText = Get-AutomaticWaveCommandText -Campaign $campaign -Filters $filters -EvidenceRoot $evidence -Create:$Create
    $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes(("`$ErrorActionPreference='Stop'; {0} *> '{1}'" -f $commandText,(Join-Path $evidence 'GuiAutomaticWave.log').Replace("'","''"))))
    $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $script:AutomaticProcess = Start-Process -FilePath $powershell -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-EncodedCommand',$encoded) -WindowStyle Hidden -PassThru
    $controls.AutomaticPreviewButton.IsEnabled = $false
    $controls.AutomaticCreateButton.IsEnabled = $false
    $controls.AutomaticEvidencePathText.Text = $evidence
    $controls.AutomaticSummaryText.Text = if ($Create) { 'Creating the Wave...' } else { 'Previewing the selection...' }
    Add-Status -Title 'Automatic Wave' -Message $controls.AutomaticSummaryText.Text
    $state = [pscustomobject]@{ Create = [bool]$Create; Evidence = $evidence; Signature = $signature; WaveName = $filters.WaveName }
    $script:AutomaticTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:AutomaticTimer.Interval = [TimeSpan]::FromSeconds(1)
    $script:AutomaticTimer.Tag = $state
    $script:AutomaticTimer.Add_Tick({
        param($eventSource,$routedEventArgs)
        if (-not $script:AutomaticProcess.HasExited) { return }
        $eventSource.Stop()
        $info = $eventSource.Tag
        $controls.AutomaticPreviewButton.IsEnabled = $true
        $summaryFile = Get-ChildItem -LiteralPath $info.Evidence -Filter 'AutomaticWaveSummary.json' -File -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        $logFile = Join-Path $info.Evidence 'GuiAutomaticWave.log'
        $logText = if (Test-Path -LiteralPath $logFile) { (Get-Content -LiteralPath $logFile -Raw -ErrorAction SilentlyContinue) } else { '' }
        if ($script:AutomaticProcess.ExitCode -ne 0 -or -not $summaryFile) {
            $controls.AutomaticSummaryText.Text = ('Automatic Wave failed (exit code {0}).{1}{2}' -f $script:AutomaticProcess.ExitCode,[Environment]::NewLine,$logText)
            Add-Status -Title 'Automatic Wave failed' -Message ('See {0}' -f $logFile)
            return
        }
        $summary = Get-Content -LiteralPath $summaryFile.FullName -Raw | ConvertFrom-Json
        $controls.AutomaticSummaryText.Text = ($summary | ConvertTo-Json -Depth 6)
        if ($info.Create) {
            Add-Status -Title 'Wave created' -Message ("Wave '{0}' created from the automatic selection. It was not launched." -f $info.WaveName)
            Update-WaveList -SelectName $info.WaveName
            Update-AutomaticView
        }
        else {
            $script:AutomaticPreviewSignature = $info.Signature
            $controls.AutomaticCreateButton.IsEnabled = $true
            Add-Status -Title 'Preview ready' -Message 'Review the selection, then Create the Wave or change the filters.'
        }
    })
    $script:AutomaticTimer.Start()
}

$inventoryRefreshXaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Refresh shared inventories" Width="1000" Height="720" MinWidth="600" MinHeight="420"
        WindowStartupLocation="CenterOwner" ShowInTaskbar="False">
    <Grid Margin="12">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <TextBlock x:Name="InventoryRefreshStateText" FontWeight="SemiBold" TextWrapping="Wrap" Margin="0,0,0,8"/>
        <TextBlock x:Name="InventoryRefreshADStateText" Grid.Row="1" TextWrapping="Wrap" Margin="0,0,0,4"/>
        <TextBox x:Name="InventoryRefreshADOutputText" Grid.Row="2" IsReadOnly="True" FontFamily="Consolas" FontSize="12"
                 Background="#0C0C0C" Foreground="#CCCCCC" Padding="6" TextWrapping="NoWrap"
                 VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
        <TextBlock x:Name="InventoryRefreshGraphStateText" Grid.Row="3" TextWrapping="Wrap" Margin="0,10,0,4"/>
        <TextBox x:Name="InventoryRefreshGraphOutputText" Grid.Row="4" IsReadOnly="True" FontFamily="Consolas" FontSize="12"
                 Background="#0C0C0C" Foreground="#CCCCCC" Padding="6" TextWrapping="NoWrap"
                 VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto"/>
        <WrapPanel Grid.Row="5" HorizontalAlignment="Right" Margin="0,8,0,0">
            <Button x:Name="InventoryRefreshOpenFolderButton" Content="Open log folder" Padding="12,4" Margin="0,0,8,0"/>
            <Button x:Name="InventoryRefreshStopButton" Content="Stop refresh" Padding="12,4" Margin="0,0,8,0"/>
            <Button x:Name="InventoryRefreshCloseButton" Content="Close" Padding="12,4" IsCancel="True"/>
        </WrapPanel>
    </Grid>
</Window>
'@

function Read-InventoryRefreshLog {
    param([Parameter(Mandatory = $true)][string]$Path,[AllowNull()][string]$Current)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return [string]$Current }
    try {
        $stream = [IO.FileStream]::new($Path, [IO.FileMode]::Open, [IO.FileAccess]::Read, [IO.FileShare]'ReadWrite, Delete')
        try { return [IO.StreamReader]::new($stream, [Text.Encoding]::UTF8, $true).ReadToEnd() } finally { $stream.Dispose() }
    }
    catch { return [string]$Current }
}

function Test-InventoryRefreshRunning {
    return ($null -ne $script:InventoryRefresh -and @($script:InventoryRefresh.Streams | Where-Object { -not $_.Process.HasExited }).Count -gt 0)
}

function Update-InventoryRefreshWindow {
    $state = $script:InventoryRefresh
    if ($null -eq $state) { return }
    foreach ($stream in $state.Streams) {
        $text = Read-InventoryRefreshLog -Path $stream.LogPath -Current $stream.Output.Text
        if ($text -ne $stream.Output.Text) { $stream.Output.Text = $text; $stream.Output.ScrollToEnd() }
        if ($stream.Finished -or -not $stream.Process.HasExited) { continue }
        $stream.Finished = $true
        $stream.ExitCode = $stream.Process.ExitCode
        $stream.StateText.Text = if ($stream.ExitCode -eq 0) { '{0}: finished at {1:HH:mm:ss}.' -f $stream.Title,(Get-Date) } else { '{0}: finished at {1:HH:mm:ss} with errors (exit code {2}).' -f $stream.Title,(Get-Date),$stream.ExitCode }
    }
    if (@($state.Streams | Where-Object { -not $_.Finished }).Count -gt 0) { return }
    $state.Timer.Stop()
    $state.Finished = $true
    $state.StopButton.IsEnabled = $false
    $failed = @($state.Streams | Where-Object { $_.ExitCode -ne 0 } | ForEach-Object { $_.Title })
    if ($failed.Count -eq 0) { $state.StateText.Text = ('Finished at {0:HH:mm:ss}. The shared inventories were refreshed.' -f (Get-Date)); Add-Status -Title 'Inventories refreshed' -Message 'Shared inventory refresh finished.' }
    else { $state.StateText.Text = ('Finished at {0:HH:mm:ss} with errors: {1}. Review the output below.' -f (Get-Date),($failed -join ', ')); Add-Status -Title 'Inventory refresh failed' -Message ('{0}. See {1}' -f ($failed -join ', '),$state.Evidence) }
    try { Update-AutomaticView } catch { }
}

function Stop-InventoryRefresh {
    if (-not (Test-InventoryRefreshRunning)) { return }
    foreach ($stream in $script:InventoryRefresh.Streams) {
        if (-not $stream.Process.HasExited) { try { Stop-Process -Id $stream.Process.Id -Force -ErrorAction Stop } catch { } }
    }
    Add-Status -Title 'Inventory refresh stopped' -Message ('Stopped by the operator. See {0}' -f $script:InventoryRefresh.Evidence)
}

function Start-InventoryRefresh {
    <#
    Runs the AD and Microsoft Graph exports in parallel, each in a hidden Windows PowerShell, and
    shows their output in a GUI window. Closing the window while the refresh runs only hides it; the
    button shows it again.
    #>
    $evidence = Join-Path $script:PackageRoot ("Runs\{0}\_InventoryRefresh\{1}" -f $script:CurrentCampaign,(Get-Date -Format 'yyyyMMdd-HHmmss'))
    New-SemtDirectory -Path $evidence
    $popup = [Windows.Markup.XamlReader]::Load([System.Xml.XmlReader]::Create([System.IO.StringReader]::new($inventoryRefreshXaml)))
    $popup.Owner = $window
    if ($window.Icon) { $popup.Icon = $window.Icon }
    $powershell = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $streams = foreach ($item in @(Get-InventoryRefreshPlan -Campaign $script:CurrentCampaign)) {
        $logPath = Join-Path $evidence ('InventoryRefresh-{0}.log' -f $item.Name)
        $encoded = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes((Get-InventoryRefreshCommandText -Exports $item.Exports -LogPath $logPath)))
        $process = Start-Process -FilePath $powershell -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-EncodedCommand',$encoded) -WindowStyle Hidden -PassThru
        $null = $process.Handle  # keeps ExitCode readable after the process exits
        $stateText = $popup.FindName(('InventoryRefresh{0}StateText' -f $item.Name))
        $stateText.Text = ('{0}: running since {1:HH:mm:ss}.' -f $item.Title,(Get-Date))
        [pscustomobject]@{ Name = $item.Name; Title = $item.Title; Process = $process; LogPath = $logPath; Finished = $false; ExitCode = $null; StateText = $stateText; Output = $popup.FindName(('InventoryRefresh{0}OutputText' -f $item.Name)) }
    }
    $script:InventoryRefresh = [pscustomobject]@{
        Streams = @($streams); Evidence = $evidence; Window = $popup; Finished = $false
        StateText = $popup.FindName('InventoryRefreshStateText'); StopButton = $popup.FindName('InventoryRefreshStopButton'); Timer = New-Object System.Windows.Threading.DispatcherTimer
    }
    $script:InventoryRefresh.StateText.Text = 'AD and Microsoft Graph exports run in parallel. AD needs the ActiveDirectory module; the Microsoft Graph sign-in window opens now.'
    $popup.FindName('InventoryRefreshOpenFolderButton').Add_Click({ Open-ShellPath $script:InventoryRefresh.Evidence })
    $script:InventoryRefresh.StopButton.Add_Click({
        if (Show-GuiQuestion -Message 'Stop the inventory refresh? The exports in progress are interrupted.' -Title 'Stop refresh') { Stop-InventoryRefresh }
    })
    $popup.FindName('InventoryRefreshCloseButton').Add_Click({ $script:InventoryRefresh.Window.Close() })
    $popup.Add_Closing({
        param($eventSource,$cancelEventArgs)
        if ($script:GuiClosing -or $script:InventoryRefresh.Finished) { return }
        $cancelEventArgs.Cancel = $true
        $eventSource.Hide()
        Add-Status -Message 'Inventory refresh continues in the background. Click "Refresh shared inventories" to show its output.'
    })
    $script:InventoryRefresh.Timer.Interval = [TimeSpan]::FromSeconds(1)
    $script:InventoryRefresh.Timer.Add_Tick({ try { Update-InventoryRefreshWindow } catch { } })
    $script:InventoryRefresh.Timer.Start()
    $popup.Show()
    Add-Status -Message ('Inventory refresh started in parallel: {0}.' -f ((@($script:InventoryRefresh.Streams) | ForEach-Object { $_.Title }) -join ', '))
}

#region Events

$controls.HeaderLogoLink.Add_MouseLeftButtonUp({ try { Start-Process 'https://workplacecloudhub.com' | Out-Null } catch { } })
$controls.CampaignCombo.Add_SelectionChanged({ try { Switch-Campaign } catch { Show-GuiError $_.Exception.Message } })
$controls.WaveCombo.Add_SelectionChanged({
    $index = $controls.WaveCombo.SelectedIndex
    $script:SelectedWave = if ($index -ge 0 -and $index -lt $script:Waves.Count) { $script:Waves[$index] } else { $null }
    try { Update-WaveView } catch { Show-GuiError $_.Exception.Message }
})
$controls.RefreshButton.Add_Click({ try { $name = if ($script:SelectedWave) { $script:SelectedWave.Name } else { '' }; Update-WaveList -SelectName $name; Update-RunsView; Update-AutomaticView; Add-Status -Message 'Refreshed.' } catch { Show-GuiError $_.Exception.Message } })
$controls.LaunchWaveButton.Add_Click({
    try {
        $wave = $script:SelectedWave
        if ($null -eq $wave) { return }
        if (@(Get-SemtActiveRuns -RunsRoot (Join-Path $script:PackageRoot 'Runs') -Campaign $script:CurrentCampaign | Where-Object { $_.Alive -and $_.WaveName -eq $wave.Name }).Count -gt 0) { Show-GuiError ("Wave '{0}' is already running." -f $wave.Name); return }
        [void](Start-GuiRun -WavePath $wave.Path -ModeKey ([string]$controls.WaveModeCombo.SelectedItem) -TargetText ("Wave '{0}' ({1} computer(s))" -f $wave.Name,$wave.ComputerCount))
    }
    catch { Show-GuiError $_.Exception.Message }
})
$controls.LaunchAllButton.Add_Click({
    try {
        $active = @(Get-SemtActiveRuns -RunsRoot (Join-Path $script:PackageRoot 'Runs') -Campaign $script:CurrentCampaign | Where-Object { $_.Alive } | ForEach-Object { $_.WaveName })
        $ready = @($script:Waves | Where-Object { $_.Launchable -and $active -notcontains $_.Name })
        if ($ready.Count -eq 0) { Show-GuiError 'No launchable Wave of this Campaign is waiting.'; return }
        $values = Get-SemtGuiOptionValues -Editors $script:Editors
        $mode = $script:RunModes[[string]$controls.WaveModeCombo.SelectedItem]
        $summary = Get-GuiLaunchSummary -Catalog $script:Catalog -Campaign $script:CurrentCampaign -Values $values -Target ('{0} Wave(s): {1}' -f $ready.Count,((@($ready) | ForEach-Object { $_.Name }) -join ', ')) -RunMode ([string]$mode[0]) -IgnoreRunGuard ([bool]$mode[1])
        if (-not (Show-GuiQuestion -Message ($summary + [Environment]::NewLine + [Environment]::NewLine + 'Launch All Waves of this Campaign only, one every 5 seconds?') -Title 'Launch All Waves')) { return }
        if ((Test-UnlimitedCycleConfirmationRequired -RunMode ([string]$mode[0]) -MaxCycles ([int]$values['MaxCycles'])) -and -not (Show-GuiQuestion -Message 'Loop mode with no cycle limit repeats until every computer is complete or the Run is stopped. Continue?' -Title 'Unlimited cycles')) { return }
        $script:LaunchQueue.Clear()
        foreach ($wave in $ready) { $script:LaunchQueue.Enqueue($wave) }
        Invoke-LaunchQueueTick
        $script:LaunchTimer = New-Object System.Windows.Threading.DispatcherTimer
        $script:LaunchTimer.Interval = [TimeSpan]::FromSeconds(5)
        $script:LaunchTimer.Add_Tick({ Invoke-LaunchQueueTick })
        $script:LaunchTimer.Start()
    }
    catch { Show-GuiError $_.Exception.Message }
})
$controls.StopAllButton.Add_Click({
    try {
        $active = @(Get-SemtActiveRuns -RunsRoot (Join-Path $script:PackageRoot 'Runs') -Campaign $script:CurrentCampaign | Where-Object { $_.Alive })
        if ($active.Count -eq 0) { Add-Status -Message 'No active Run of this Campaign.'; return }
        $force = (@($active | Where-Object { Test-Path -LiteralPath $_.SignalPath -PathType Leaf }).Count -gt 0)
        $question = if ($force) { 'A controlled stop was already requested. Force the local stop of running workers? Remote endpoints may continue and must be verified before relaunch.' } else { 'Request a controlled stop of {0} active Run(s) of this Campaign? No new computer starts; running workers finish.' -f $active.Count }
        if (-not (Show-GuiQuestion -Message $question -Title 'Stop running Waves')) { return }
        foreach ($run in $active) { [void](Request-SemtRunStop -ActiveRun $run -Force:$force) }
        Add-Status -Title 'Stop requested' -Message ('{0} Run(s): {1}' -f $active.Count,$(if ($force) { 'forced local stop' } else { 'controlled stop' }))
        Update-RunsView
    }
    catch { Show-GuiError $_.Exception.Message }
})
$controls.OpenWaveFolderButton.Add_Click({ if ($script:SelectedWave) { Open-ShellPath $script:SelectedWave.Path } })
$controls.OpenWaveComputersButton.Add_Click({ if ($script:SelectedWave) { Open-TextFile $script:SelectedWave.ComputersPath } })
$controls.OpenWaveConfigButton.Add_Click({ if ($script:SelectedWave) { Open-TextFile (Join-Path $script:SelectedWave.Path 'Wave.local.json') } })
$controls.RefreshWrappersButton.Add_Click({
    try {
        if ($null -eq $script:SelectedWave) { return }
        [void](Write-SemtWaveWrappers -WavePath $script:SelectedWave.Path)
        Update-WaveList -SelectName $script:SelectedWave.Name
        Add-Status -Message ("Launchers refreshed for Wave '{0}'." -f $script:SelectedWave.Name)
    }
    catch { Show-GuiError $_.Exception.Message }
})
$controls.LaunchSingleButton.Add_Click({
    try {
        $name = ConvertTo-SemtComputerName -Line $controls.SingleComputerText.Text
        if ([string]::IsNullOrWhiteSpace($name)) { Show-GuiError 'Enter a computer name.'; return }
        [void](Start-GuiRun -ComputerName $name -ModeKey ([string]$controls.SingleModeCombo.SelectedItem) -TargetText ("Single PC '{0}'" -f $name))
    }
    catch { Show-GuiError $_.Exception.Message }
})
$controls.OpenSingleRunFolderButton.Add_Click({ $path = $controls.SingleRunFolderText.Text; if (Test-Path -LiteralPath $path) { Open-ShellPath $path } else { Show-GuiError 'No Single PC Run yet for this Campaign.' } })
$controls.CreateWaveButton.Add_Click({
    try {
        $name = $controls.NewWaveNameText.Text.Trim()
        $computers = @(Split-GuiList $controls.NewWaveComputersText.Text)
        if (-not (Show-GuiQuestion -Message ("Create Wave '{0}' for {1} with {2} computer(s)?{3}{3}The Wave is created but NOT launched." -f $name,$script:CampaignTitles[$script:CurrentCampaign],$computers.Count,[Environment]::NewLine) -Title 'Create Wave')) { return }
        $wave = New-SemtWave -PackageRoot $script:PackageRoot -Campaign $script:CurrentCampaign -Name $name -Computers $computers -Source Manual
        $controls.NewWavePathText.Text = $wave.ComputersPath
        Update-WaveList -SelectName $wave.Name
        Add-Status -Title 'Wave created' -Message ("Wave '{0}' created ({1} computer(s)). It was not launched." -f $wave.Name,$wave.ComputerCount)
    }
    catch { Show-GuiError $_.Exception.Message }
})
$controls.OpenNewWaveComputersButton.Add_Click({ Open-TextFile $controls.NewWavePathText.Text })
foreach ($name in @('AutomaticWaveNameText','AutomaticNamePrefixText','AutomaticNameContainsText','AutomaticLastLogonDaysText')) { $controls[$name].Add_TextChanged({ if ($script:AutomaticPreviewSignature) { $script:AutomaticPreviewSignature = ''; $controls.AutomaticCreateButton.IsEnabled = $false } }) }
foreach ($name in @('AutomaticExcludeStaleAdCheck','AutomaticExcludeIntuneCheck')) { $controls[$name].Add_Click({ $script:AutomaticPreviewSignature = ''; $controls.AutomaticCreateButton.IsEnabled = $false }) }
foreach ($name in @('AutomaticSourceCombo','AutomaticAdDomainCombo')) { $controls[$name].Add_SelectionChanged({ $script:AutomaticPreviewSignature = ''; $controls.AutomaticCreateButton.IsEnabled = $false }) }
$controls.AutomaticPreviewButton.Add_Click({ try { Invoke-AutomaticWave } catch { Show-GuiError $_.Exception.Message; $controls.AutomaticPreviewButton.IsEnabled = $true } })
$controls.AutomaticCreateButton.Add_Click({ try { Invoke-AutomaticWave -Create } catch { Show-GuiError $_.Exception.Message; $controls.AutomaticPreviewButton.IsEnabled = $true } })
$controls.AutomaticOpenEvidenceButton.Add_Click({ Open-ShellPath $controls.AutomaticEvidencePathText.Text })
$controls.AutomaticRefreshInventoryButton.Add_Click({
    try {
        if ($script:InventoryRefresh -and -not $script:InventoryRefresh.Finished) { $script:InventoryRefresh.Window.Show(); [void]$script:InventoryRefresh.Window.Activate(); return }
        $kinds = if ($script:CurrentCampaign -eq 'Windows11Upgrade') { 'AD and Intune' } else { 'AD, Intune and Entra' }
        if (-not (Show-GuiQuestion -Message ('Refresh the shared {0} inventories? AD and Microsoft Graph run in parallel and their output is shown in a window of the GUI. AD needs the ActiveDirectory module; Microsoft Graph opens one interactive sign-in.' -f $kinds) -Title 'Refresh shared inventories')) { return }
        if ($script:InventoryRefresh) { $script:InventoryRefresh.Window.Close() }
        Start-InventoryRefresh
    }
    catch { Show-GuiError $_.Exception.Message }
})
$controls.RefreshRunsButton.Add_Click({ try { Update-RunsView } catch { Show-GuiError $_.Exception.Message } })
$controls.OpenRunReportButton.Add_Click({ $run = $controls.RunsGrid.SelectedItem; if ($run) { $report = [string](Get-SemtPropertyValue -InputObject $run -Name 'ReportPath' -Default ''); if ([string]::IsNullOrWhiteSpace($report)) { $report = @(Get-ChildItem -LiteralPath (Join-Path $run.RunRoot 'Reports') -Filter '*.html' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1 | ForEach-Object { $_.FullName }) -join '' }; Open-ShellPath $report } })
$controls.OpenRunFolderButton.Add_Click({ $run = $controls.RunsGrid.SelectedItem; if ($run) { Open-ShellPath ([string]$run.RunRoot) } })
foreach ($pair in @(@('StopSelectedRunButton',$false),@('ForceStopSelectedRunButton',$true))) {
    $button = $controls[$pair[0]]
    $button.Tag = [bool]$pair[1]
    $button.Add_Click({
        param($eventSource,$routedEventArgs)
        try {
            $run = $controls.RunsGrid.SelectedItem
            if ($null -eq $run) { return }
            $active = @(Get-SemtActiveRuns -RunsRoot (Join-Path $script:PackageRoot 'Runs') -Campaign $script:CurrentCampaign | Where-Object { $_.Alive -and $_.RunId -eq [string]$run.RunId })
            if ($active.Count -eq 0) { Add-Status -Message ('Run {0} is not active.' -f $run.RunId); return }
            $force = [bool]$eventSource.Tag
            $question = if ($force) { 'Force the local stop of running workers? Remote endpoints may continue and must be verified before relaunch.' } else { 'Request a controlled stop? No new computer starts; running workers finish.' }
            if (-not (Show-GuiQuestion -Message $question -Title 'Stop Run')) { return }
            Add-Status -Title 'Stop requested' -Message ('Run {0}: {1}' -f $run.RunId,(Request-SemtRunStop -ActiveRun $active[0] -Force:$force))
            Update-RunsView
        }
        catch { Show-GuiError $_.Exception.Message }
    })
}
$controls.SaveOptionsButton.Add_Click({ if (Save-CurrentCampaignState) { Add-Status -Message ('Choices saved for {0}.' -f $script:CampaignTitles[$script:CurrentCampaign]) } })
$controls.ResetOptionsButton.Add_Click({
    try {
        if (-not (Show-GuiQuestion -Message 'Reset the options of this Campaign to the global configuration and campaign defaults?' -Title 'Reset options')) { return }
        $values = Get-GuiInitialOptionValues -Catalog $script:Catalog -Campaign $script:CurrentCampaign -IgnoreSavedState
        $script:Editors = Add-SemtGuiOptionControls -OptionsHost $controls.OptionsHost -Catalog $script:Catalog -Campaign $script:CurrentCampaign -Values $values
        Save-GuiState -Campaign $script:CurrentCampaign -Values $values
        Add-Status -Message 'Options reset to defaults.'
    }
    catch { Show-GuiError $_.Exception.Message }
})
$window.Add_Closing({
    param($eventSource,$routedEventArgs)
    [void](Save-CurrentCampaignState -Quiet)
    if (Test-InventoryRefreshRunning) {
        $choice = [System.Windows.MessageBox]::Show($window, "The shared inventory refresh is still running in the background.`n`nYes: stop it, then close.`nNo: close and let it finish (its log stays in the Runs folder).`nCancel: keep the GUI open.", 'Inventory refresh', 'YesNoCancel', 'Warning')
        if ($choice -eq [System.Windows.MessageBoxResult]::Cancel) { $routedEventArgs.Cancel = $true; return }
        if ($choice -eq [System.Windows.MessageBoxResult]::Yes) { Stop-InventoryRefresh }
    }
    $script:GuiClosing = $true
    $active = @(Get-SemtActiveRuns -RunsRoot (Join-Path $script:PackageRoot 'Runs') | Where-Object { $_.Alive })
    if ($active.Count -eq 0) { return }
    $choice = [System.Windows.MessageBox]::Show($window, ("{0} Run(s) are still active.`n`nYes: request a controlled stop, then close.`nNo: close and leave them running.`nCancel: keep the GUI open." -f $active.Count), 'Active Runs', 'YesNoCancel', 'Warning')
    if ($choice -eq [System.Windows.MessageBoxResult]::Cancel) { $script:GuiClosing = $false; $routedEventArgs.Cancel = $true; return }
    if ($choice -eq [System.Windows.MessageBoxResult]::Yes) { foreach ($run in $active) { try { [void](Request-SemtRunStop -ActiveRun $run) } catch { } } }
})

#endregion

$controls.CampaignCombo.SelectedIndex = 0
if ($ValidateOnly) {
    for ($index = 0; $index -lt $controls.CampaignCombo.Items.Count; $index++) {
        $controls.CampaignCombo.SelectedIndex = $index
        if ($script:Editors.Count -eq 0) { throw ('Options were not generated for {0}.' -f $script:CurrentCampaign) }
        [void](Get-SemtGuiOptionValues -Editors $script:Editors)
        Write-Output ('ValidateOnly: GUI {0} loaded; campaign={1}; options={2}; waves={3}.' -f $script:GuiVersion,$script:CurrentCampaign,$script:Editors.Count,$script:Waves.Count)
    }
    return
}

$window.Add_ContentRendered({
    if ($splash -and (Get-Command -Name Close-SmartM365GuiSplash -ErrorAction SilentlyContinue)) { Close-SmartM365GuiSplash -Splash $splash }
    if (Get-Command -Name Start-SmartM365GuiUpdateCheck -ErrorAction SilentlyContinue) {
        try { $script:UpdateCheckTimer = Start-SmartM365GuiUpdateCheck -Owner $window -ManifestPath (Join-Path $PSScriptRoot 'SmartM365.GuiUpdateCheck.psd1') -AppRoot $script:PackageRoot -OnStatus { param([string]$Message,[string]$Title) Add-Status -Title $Title -Message $Message } } catch { }
    }
})
if (Get-Command -Name Set-SmartM365WpfWindowVisible -ErrorAction SilentlyContinue) { $window.Add_SourceInitialized({ Set-SmartM365WpfWindowVisible -Window $window }) }
[void]$window.ShowDialog()

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBxzlIZWNPX7MSU
# gsXItoY7ZC/7rx0YGtqpOO8UIJOizaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEaScfGIZ1RpBsgrLdIzpi2M3lNAlkDE2kCv2Yk++/ZfMA0GCSqG
# SIb3DQEBAQUABIIBgFqND/TDLa9MJVxS3ji+VcF0PqnOfbhIQDp6CdexhT6qnqoO
# 7F7lCo8LdiNqAN6c+Bh9R/0Dt1k6OGlH5l+aXvCIaCI+wM/tecQhDVexouugyDEc
# vlJB2UPAv+rkTOGi0GNrBPNkHIgP9bq2ace6CMwafA3Eqa/8709i5jP70fxeSIto
# ZianK7oseQw9LZnBe9B4Lz+8yixFxXcGHfgu8Stghbviow8ojUtyIKcoS9q63IyG
# SV3wNvc3atGmR7e3mzxn2udCd/10CGR8DlWAdb+zlJCuqdiPJk0GP+/Uf7rwtZaR
# OxpwfLtpRzznETeZj1NbOth9RW51QsOoz4n0ylyoSAhNblRY2Rbbo9TmTeLQw5ND
# 1Ts7rLNOYq8CD3RlJzCErP9az8JHvtbUqkSsAa/odITSgHY9ziLNpyozQ3UXSOdd
# oGHSuVhV20VjGcQsBaET4cKyzXvAaP/j4XSM/GEl+4VmL7NNCLifO06l/okLEl8C
# nDjEgq+CAqKWxP9fg6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgxNTU4
# NDZaMC8GCSqGSIb3DQEJBDEiBCCHIhP1OPmmDRp/i6HU3agY1TLV1F2No2p96fiy
# RuVzXDANBgkqhkiG9w0BAQEFAASCAgBaNeulqcM84zPChvwSVyL316djvWLOO1z+
# umX6mwOnYnTJ/ZCnwCuaqJ/vPBS1dYQP3EugTJ5AvfzJBOGqun39RdVclyLaZ5yk
# NeAmJ730KqEwuK6gcXwqCqI0vJz2YP+7db2ZbE39gFvmT/Ne5AIRgZ8tKwdLDSpa
# w5IzLFYHG3Gir4a7ZSVcmn4eZj/VSasr37F6FX99AnFKrepyOtAcYVV6mL8bpVum
# Citsu3ijkqLnEHLBds6NGMbxc3Sf8TMe9JST7vNO6bNg6aKb4r5TgHNbrnvA4l3A
# 1lwj0iWtPMAMRqvCPi8J+ENVvo65XDKw24uGURcSYgMVW+TrdrOmdmyx9MlbYoyF
# lhqxWIkdsVnxKFc/1pAj13JoNvHoCuR3VhZcPpRyTa864ZR6RWDovK050BjtJtgh
# Hlql2uLVObHLxa838/pqM0JAnedod2RJY/gGOnlqCcTsslI88Ut5VWIXuxlmYFvv
# c0A8LUvczZfM3El8K004FYcmYsUcKH/eBPPikIg6Edbrlhs1sm7oSUfGLK4L/xma
# qlOCqevI6Mr59LvgoQopf3p9H07Ty+Zj4ygjUcosjNAK7psYRR9wY1Ii/CAFnSU7
# B31Ywmrr3cH2vTeS1USxmXIXx4pk8TiwGPSqNbe3uhSVgvZtV6CAHAkb9eI6xB1X
# uPFxRZBqTA==
# SIG # End signature block
