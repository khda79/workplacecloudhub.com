<#
.SYNOPSIS
WPF launcher for SmartM365 Exchange user notification campaigns.

.DESCRIPTION
Provides a SmartM365-styled launcher for Exchange user notification campaigns.
The GUI keeps campaign logic inside the campaign scripts. It validates files,
builds the PowerShell command line, launches the selected campaign, and streams
output back to the operator.
#>
[CmdletBinding()]
param()

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

if ([System.Threading.Thread]::CurrentThread.ApartmentState -ne 'STA') {
    $pwsh = (Get-Process -Id $PID).Path
    Start-Process -FilePath $pwsh -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', $PSCommandPath) | Out-Null
    return
}

Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms

$script:Window = $null
$script:LoadingWindow = $null
$script:GuiSplash = $null
$script:OutputBox = $null
$script:CurrentProcess = $null
$script:UiReady = $false
$script:SelectedCampaignIndex = 0
$script:AutoLanguageLabel = '(Auto)'
$script:GuiScriptRoot = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $PSScriptRoot
}
elseif (-not [string]::IsNullOrWhiteSpace($PSCommandPath)) {
    Split-Path -Path $PSCommandPath -Parent
}
else {
    (Get-Location).Path
}

function Get-ExchangeUserNotificationsRoot {
    $current = $script:GuiScriptRoot
    while ($current) {
        if ((Split-Path -Path $current -Leaf) -eq 'ExchangeUserNotifications') { return $current }
        $parent = Split-Path -Path $current -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) { break }
        $current = $parent
    }
    throw 'ExchangeUserNotifications root not found.'
}

function Find-UpwardPath {
    param(
        [Parameter(Mandatory)][string]$RelativePath,
        [string]$StartPath = $script:GuiScriptRoot
    )

    $current = $StartPath
    while ($current) {
        $candidate = Join-Path -Path $current -ChildPath $RelativePath
        if (Test-Path -LiteralPath $candidate) { return $candidate }
        $parent = Split-Path -Path $current -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) { break }
        $current = $parent
    }
    return ''
}

function Show-LoadingWindow {
    param([string]$Message = 'Loading SmartM365...')

    if ($script:LoadingWindow) { return }

    $accentBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(0, 120, 212))
    $mutedBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(95, 107, 122))
    $borderBrush = [System.Windows.Media.SolidColorBrush]::new([System.Windows.Media.Color]::FromRgb(210, 224, 240))

    $panel = [System.Windows.Controls.StackPanel]::new()
    $panel.Margin = [System.Windows.Thickness]::new(22)

    $badge = [System.Windows.Controls.TextBlock]::new()
    $badge.Text = 'SMARTM365 COMMUNICATIONS'
    $badge.Foreground = $accentBrush
    $badge.FontSize = 11
    $badge.FontWeight = [System.Windows.FontWeights]::SemiBold
    [void]$panel.Children.Add($badge)

    $title = [System.Windows.Controls.TextBlock]::new()
    $title.Text = 'Exchange User Notifications'
    $title.FontSize = 20
    $title.FontWeight = [System.Windows.FontWeights]::SemiBold
    $title.Margin = [System.Windows.Thickness]::new(0, 8, 0, 4)
    [void]$panel.Children.Add($title)

    $messageBlock = [System.Windows.Controls.TextBlock]::new()
    $messageBlock.Text = $Message
    $messageBlock.Foreground = $mutedBrush
    $messageBlock.FontSize = 12
    $messageBlock.Margin = [System.Windows.Thickness]::new(0, 0, 0, 14)
    [void]$panel.Children.Add($messageBlock)

    $progress = [System.Windows.Controls.ProgressBar]::new()
    $progress.Height = 4
    $progress.IsIndeterminate = $true
    [void]$panel.Children.Add($progress)

    $border = [System.Windows.Controls.Border]::new()
    $border.Background = [System.Windows.Media.Brushes]::White
    $border.BorderBrush = $borderBrush
    $border.BorderThickness = [System.Windows.Thickness]::new(1)
    $border.CornerRadius = [System.Windows.CornerRadius]::new(8)
    $border.Child = $panel

    $script:LoadingWindow = [System.Windows.Window]::new()
    $script:LoadingWindow.Title = 'SmartM365'
    $script:LoadingWindow.Width = 390
    $script:LoadingWindow.Height = 165
    $script:LoadingWindow.WindowStartupLocation = 'CenterScreen'
    $script:LoadingWindow.ResizeMode = 'NoResize'
    $script:LoadingWindow.ShowInTaskbar = $false
    $script:LoadingWindow.Topmost = $true
    $script:LoadingWindow.Content = $border
    $loadingIconPath = Find-UpwardPath -RelativePath 'WorkplaceCloudHub.ico'
    if (-not [string]::IsNullOrWhiteSpace($loadingIconPath)) {
        try {
            $script:LoadingWindow.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]$loadingIconPath)
        }
        catch {
            Write-Verbose ("Loading window icon load failed: {0}" -f $_.Exception.Message)
        }
    }
    $script:LoadingWindow.Show()

    try {
        $script:LoadingWindow.Dispatcher.Invoke([Action]{}, [System.Windows.Threading.DispatcherPriority]::ApplicationIdle)
    }
    catch {
        Write-Verbose ("Loading window render pump failed: {0}" -f $_.Exception.Message)
    }
}

function Close-LoadingWindow {
    if (-not $script:LoadingWindow) { return }
    try { $script:LoadingWindow.Close() }
    catch { Write-Verbose ("Loading window close failed: {0}" -f $_.Exception.Message) }
    $script:LoadingWindow = $null
}

function Hide-LoadingWindow {
    if (-not $script:LoadingWindow) { return }
    try {
        $script:LoadingWindow.Topmost = $false
        $script:LoadingWindow.Hide()
    }
    catch {
        Write-Verbose ("Loading window hide failed: {0}" -f $_.Exception.Message)
    }
}

function Get-PowerShellPath {
    $command = Get-Command pwsh -ErrorAction SilentlyContinue
    if ($command -and $command.Source) { return $command.Source }
    return (Get-Process -Id $PID).Path
}

function ConvertTo-QuotedArgument {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) { return '""' }
    return '"' + ($Value -replace '"', '\"') + '"'
}

function Get-CampaignDefinition {
    param([Parameter(Mandatory)][string]$RootPath)

    @(
        [pscustomobject]@{
            Id = 'ExchangeMigration'
            Name = 'Exchange Migration'
            Badge = 'MIGRATION'
            Description = 'Mailbox move notification'
            ScriptPath = Join-Path $RootPath 'ExchangeMigration\SmartM365-ExchangeMigration-NotifyUsers.ps1'
            LocalConfigPath = Join-Path $RootPath 'Config\Campaigns\ExchangeMigration.local.json'
            TemplateConfigPath = Join-Path $RootPath 'Config\Campaigns\ExchangeMigration.local.json.template'
            TemplateFolderPath = Join-Path $RootPath 'ExchangeMigration\Templates'
            RecipientPathParameter = '-RecipientsPath'
            SupportsRecipientPath = $true
            SupportsEffectiveDate = $true
            SupportsFromList = $false
            SupportsInventory = $false
            SupportsSkipConfirmation = $false
            ExpectedTokens = @('Date','FooterLogoImgTag','Hotline','LogoImgTag','NewWebmailUrl','OldWebmailUrl','TermsPortalBlockStyle','TermsPortalUrl')
        },
        [pscustomobject]@{
            Id = 'ExchangeArchive'
            Name = 'Exchange Archive'
            Badge = 'ARCHIVE'
            Description = 'Online archive enablement notification'
            ScriptPath = Join-Path $RootPath 'ExchangeArchive\SmartM365-ExchangeArchive-NotifyUsers.ps1'
            LocalConfigPath = Join-Path $RootPath 'Config\Campaigns\ExchangeArchive.local.json'
            TemplateConfigPath = Join-Path $RootPath 'Config\Campaigns\ExchangeArchive.local.json.template'
            TemplateFolderPath = Join-Path $RootPath 'ExchangeArchive\Templates'
            RecipientPathParameter = '-RecipientsPath'
            SupportsRecipientPath = $true
            SupportsEffectiveDate = $true
            SupportsFromList = $false
            SupportsInventory = $false
            SupportsSkipConfirmation = $false
            ExpectedTokens = @('Date','FooterLogoImgTag','Hotline','LogoImgTag','MailboxUsageText')
        },
        [pscustomobject]@{
            Id = 'ExchangeMigrationMailboxSizeReduction'
            Name = 'Migration Mailbox Size Reduction'
            Badge = 'SIZE'
            Description = 'Mailbox quota reduction notification'
            ScriptPath = Join-Path $RootPath 'ExchangeMigrationMailboxSizeReduction\SmartM365-ExchangeMigrationMailboxSizeReduction-NotifyUsers.ps1'
            LocalConfigPath = Join-Path $RootPath 'Config\Campaigns\ExchangeMigrationMailboxSizeReduction.local.json'
            TemplateConfigPath = Join-Path $RootPath 'Config\Campaigns\ExchangeMigrationMailboxSizeReduction.local.json.template'
            TemplateFolderPath = Join-Path $RootPath 'ExchangeMigrationMailboxSizeReduction\Templates'
            RecipientPathParameter = '-ListCsvPath'
            SupportsRecipientPath = $true
            SupportsEffectiveDate = $false
            SupportsFromList = $true
            SupportsInventory = $false
            SupportsSkipConfirmation = $true
            ExpectedTokens = @('FooterLogoImgTag','FreeTextBlockStyle','FreeTextBlockTextHtml','Hotline','LogoImgTag','MailboxQuotaGB','MailboxSizeGB','WebmailUrl')
        }
    )
}

function Read-JsonFile {
    param([Parameter(Mandatory)][string]$Path)

    Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
}

function Initialize-GuiLocalJsonFilesFromTemplates {
    param([array]$Campaigns)

    $root = Get-ExchangeUserNotificationsRoot
    $pairs = New-Object System.Collections.ArrayList
    [void]$pairs.Add([pscustomobject]@{
        LocalPath = (Join-Path -Path $root -ChildPath 'Config\Communications.local.json')
        TemplatePath = (Join-Path -Path $root -ChildPath 'Config\Communications.local.json.template')
        Description = 'shared communication local configuration'
    })

    foreach ($campaign in $Campaigns) {
        [void]$pairs.Add([pscustomobject]@{
            LocalPath = $campaign.LocalConfigPath
            TemplatePath = $campaign.TemplateConfigPath
            Description = ("{0} campaign local configuration" -f $campaign.Name)
        })
    }

    $created = New-Object System.Collections.ArrayList
    foreach ($pair in $pairs) {
        if ([string]::IsNullOrWhiteSpace([string]$pair.LocalPath)) { continue }
        if (Test-Path -LiteralPath $pair.LocalPath) { continue }
        if (-not (Test-Path -LiteralPath $pair.TemplatePath)) {
            Write-Warning ("Local JSON not found and template is missing. Local JSON: {0}; Template: {1}" -f $pair.LocalPath, $pair.TemplatePath)
            continue
        }

        $folder = Split-Path -Path $pair.LocalPath -Parent
        if (-not [string]::IsNullOrWhiteSpace($folder) -and -not (Test-Path -LiteralPath $folder)) {
            New-Item -Path $folder -ItemType Directory -Force -ErrorAction Stop | Out-Null
        }
        Copy-Item -LiteralPath $pair.TemplatePath -Destination $pair.LocalPath -Force -ErrorAction Stop
        [void]$created.Add($pair)
    }

    if ($created.Count -eq 0) { return }

    Write-Host 'Created missing Exchange notification local JSON file(s) from template:' -ForegroundColor Yellow
    foreach ($pair in $created) {
        Write-Host ("- {0}: {1}" -f $pair.Description, $pair.LocalPath) -ForegroundColor Yellow
    }
    Write-Host 'Review the generated local JSON values; continuing with current file values.' -ForegroundColor Yellow
}
function Get-ConfigPropertyValue {
    param(
        [AllowNull()]$Config,
        [Parameter(Mandatory)][string]$Name,
        [string]$DefaultValue = ''
    )

    if ($null -eq $Config) { return $DefaultValue }
    $property = $Config.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $DefaultValue }
    return [string]$property.Value
}

function Get-CampaignConfigValue {
    param(
        [Parameter(Mandatory)]$Campaign,
        [Parameter(Mandatory)][string]$Name,
        [string]$DefaultValue = ''
    )

    $value = $DefaultValue
    if (Test-Path -LiteralPath $Campaign.TemplateConfigPath) {
        try {
            $templateConfig = Read-JsonFile -Path $Campaign.TemplateConfigPath
            $value = Get-ConfigPropertyValue -Config $templateConfig -Name $Name -DefaultValue $value
        }
        catch {
            Write-Verbose ("Unable to read template campaign config '{0}': {1}" -f $Campaign.TemplateConfigPath, $_.Exception.Message)
        }
    }
    if (Test-Path -LiteralPath $Campaign.LocalConfigPath) {
        try {
            $localConfig = Read-JsonFile -Path $Campaign.LocalConfigPath
            $value = Get-ConfigPropertyValue -Config $localConfig -Name $Name -DefaultValue $value
        }
        catch {
            Write-Verbose ("Unable to read local campaign config '{0}': {1}" -f $Campaign.LocalConfigPath, $_.Exception.Message)
        }
    }
    return $value
}

function Get-SharedConfigValue {
    param(
        [Parameter(Mandatory)][string]$Name,
        [string]$DefaultValue = ''
    )

    $root = Get-ExchangeUserNotificationsRoot
    $templatePath = Join-Path -Path $root -ChildPath 'Config\Communications.local.json.template'
    $localPath = Join-Path -Path $root -ChildPath 'Config\Communications.local.json'
    $value = $DefaultValue
    foreach ($path in @($templatePath, $localPath)) {
        if (-not (Test-Path -LiteralPath $path)) { continue }
        try {
            $config = Read-JsonFile -Path $path
            $candidate = Get-ConfigPropertyValue -Config $config -Name $Name -DefaultValue $value
            if (-not [string]::IsNullOrWhiteSpace($candidate) -and $candidate -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) {
                $value = $candidate
            }
        }
        catch {
            Write-Verbose ("Unable to read shared config '{0}': {1}" -f $path, $_.Exception.Message)
        }
    }
    return $value
}

function Get-LocalSharedConfigValue {
    param(
        [Parameter(Mandatory)][string]$Name,
        [string]$DefaultValue = ''
    )

    $root = Get-ExchangeUserNotificationsRoot
    $localPath = Join-Path -Path $root -ChildPath 'Config\Communications.local.json'
    if (-not (Test-Path -LiteralPath $localPath)) { return $DefaultValue }
    try {
        $config = Read-JsonFile -Path $localPath
        return Get-ConfigPropertyValue -Config $config -Name $Name -DefaultValue $DefaultValue
    }
    catch {
        Write-Verbose ("Unable to read shared local config '{0}': {1}" -f $localPath, $_.Exception.Message)
        return $DefaultValue
    }
}

function Get-LocalCampaignConfigValue {
    param(
        [Parameter(Mandatory)]$Campaign,
        [Parameter(Mandatory)][string]$Name,
        [string]$DefaultValue = ''
    )

    if (-not (Test-Path -LiteralPath $Campaign.LocalConfigPath)) { return $DefaultValue }
    try {
        $config = Read-JsonFile -Path $Campaign.LocalConfigPath
        return Get-ConfigPropertyValue -Config $config -Name $Name -DefaultValue $DefaultValue
    }
    catch {
        Write-Verbose ("Unable to read local campaign config '{0}': {1}" -f $Campaign.LocalConfigPath, $_.Exception.Message)
        return $DefaultValue
    }
}

function Get-CampaignConfigObjectValue {
    param(
        [Parameter(Mandatory)]$Campaign,
        [Parameter(Mandatory)][string]$Name
    )

    foreach ($path in @($Campaign.TemplateConfigPath, $Campaign.LocalConfigPath)) {
        if (-not (Test-Path -LiteralPath $path)) { continue }
        try {
            $config = Read-JsonFile -Path $path
            $property = $config.PSObject.Properties[$Name]
            if ($null -ne $property -and $null -ne $property.Value) { return $property.Value }
        }
        catch {
            Write-Verbose ("Unable to read campaign config '{0}': {1}" -f $path, $_.Exception.Message)
        }
    }
    return $null
}

function Test-TeamsUserMessageTemplateConfigured {
    param([Parameter(Mandatory)]$Campaign)

    $messageByLanguage = Get-CampaignConfigObjectValue -Campaign $Campaign -Name 'TeamsUserMessageByLanguage'
    if ($null -eq $messageByLanguage) { return $false }

    foreach ($property in @($messageByLanguage.PSObject.Properties)) {
        if ($null -ne $property.Value -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
            return $true
        }
    }
    return $false
}

function Get-TeamsUserMessageGuiState {
    param([Parameter(Mandatory)]$Campaign)

    $sharedMode = Get-LocalSharedConfigValue -Name 'TeamsUserMessageMode' -DefaultValue ''
    $campaignMode = Get-LocalCampaignConfigValue -Campaign $Campaign -Name 'TeamsUserMessageMode' -DefaultValue ''
    $mode = if (-not [string]::IsNullOrWhiteSpace($campaignMode) -and $campaignMode -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) { $campaignMode } else { $sharedMode }
    $modeEnabled = $mode -in @('Graph', 'GraphDelegated', 'Delegated')
    $templateConfigured = Test-TeamsUserMessageTemplateConfigured -Campaign $Campaign
    $available = ($modeEnabled -and $templateConfigured)

    $reason = if (-not $modeEnabled) {
        'Configure TeamsUserMessageMode = GraphDelegated in the local communications or campaign config to enable this option.'
    }
    elseif (-not $templateConfigured) {
        'TeamsUserMessageByLanguage is missing or empty in the campaign configuration.'
    }
    else {
        'Optional. Uses Microsoft Graph delegated User.Read, Chat.Create and ChatMessage.Send permissions to send a one-on-one Teams message.'
    }

    [pscustomobject]@{
        Available = $available
        Checked = $available
        ToolTip = $reason
    }
}

function ConvertTo-ConfigDisplayText {
    param([AllowNull()]$Value)

    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [bool] -or $Value -is [int] -or $Value -is [double] -or $Value -is [decimal]) { return [string]$Value }

    try {
        return (($Value | ConvertTo-Json -Compress -Depth 8) -replace '\s+', ' ')
    }
    catch {
        return [string]$Value
    }
}

function Get-CampaignConfigRows {
    param([Parameter(Mandatory)]$Campaign)

    $rows = @()
    $indexByName = @{}

    foreach ($source in @(
        @{ Label = 'Template'; Path = $Campaign.TemplateConfigPath },
        @{ Label = 'Local'; Path = $Campaign.LocalConfigPath }
    )) {
        if (-not (Test-Path -LiteralPath $source.Path)) { continue }

        try {
            $config = Read-JsonFile -Path $source.Path
            foreach ($property in @($config.PSObject.Properties)) {
                $row = [pscustomobject]@{
                    Name = $property.Name
                    Value = ConvertTo-ConfigDisplayText -Value $property.Value
                    Source = $source.Label
                }

                if ($indexByName.ContainsKey($property.Name)) {
                    $rows[$indexByName[$property.Name]] = $row
                }
                else {
                    $indexByName[$property.Name] = $rows.Count
                    $rows += $row
                }
            }
        }
        catch {
            $rows += [pscustomobject]@{
                Name = 'Configuration read error'
                Value = $_.Exception.Message
                Source = $source.Label
            }
        }
    }

    return @($rows)
}

function Show-CampaignConfigRows {
    param([Parameter(Mandatory)]$Campaign)

    if (-not $script:ConfigParameterList) { return }
    $rows = @(Get-CampaignConfigRows -Campaign $Campaign)
    $script:ConfigParameterList.ItemsSource = $null
    $script:ConfigParameterList.ItemsSource = $rows
}

function Get-CampaignDefaultRecipientPath {
    param([Parameter(Mandatory)]$Campaign)

    if ($Campaign.SupportsFromList) {
        return Get-CampaignConfigValue -Campaign $Campaign -Name 'ListCsvPath'
    }

    return Get-CampaignConfigValue -Campaign $Campaign -Name 'RecipientsPath'
}

function Get-DefaultTenantProfile {
    $defaultTenant = 'test'
    foreach ($relativePath in @('Config\SmartM365.global.local.json.template', 'Config\SmartM365.global.local.json')) {
        $path = Find-UpwardPath -RelativePath $relativePath
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        try {
            $config = Read-JsonFile -Path $path
            $candidate = Get-ConfigPropertyValue -Config $config -Name 'DefaultTenant' -DefaultValue $defaultTenant
            if (-not [string]::IsNullOrWhiteSpace($candidate) -and $candidate -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) {
                $defaultTenant = $candidate
            }
        }
        catch {
            Write-Verbose ("Unable to read default tenant profile from '{0}': {1}" -f $path, $_.Exception.Message)
        }
    }
    return $defaultTenant
}

function Get-TenantProfileNames {
    $root = Get-ExchangeUserNotificationsRoot
    $smartM365Root = Split-Path -Path (Split-Path -Path $root -Parent) -Parent
    $tenantFolder = Join-Path -Path $smartM365Root -ChildPath 'Config\Tenants'
    $names = New-Object System.Collections.Generic.List[string]

    if (Test-Path -LiteralPath $tenantFolder -PathType Container) {
        foreach ($file in (Get-ChildItem -LiteralPath $tenantFolder -Filter '*.local.json' -File | Sort-Object BaseName)) {
            [void]$names.Add(($file.Name -replace '\.local\.json$', ''))
        }
    }

    $defaultTenant = Get-DefaultTenantProfile
    if (-not [string]::IsNullOrWhiteSpace($defaultTenant) -and -not $names.Contains($defaultTenant)) {
        [void]$names.Insert(0, $defaultTenant)
    }
    if ($names.Count -eq 0) { [void]$names.Add('test') }

    return @($names | Sort-Object -Unique)
}

function Get-SelectedTenantProfile {
    if ($script:TenantProfileCombo -and $script:TenantProfileCombo.SelectedItem) {
        return [string]$script:TenantProfileCombo.SelectedItem
    }
    return Get-DefaultTenantProfile
}

function Select-ComboBoxValue {
    param(
        [Parameter(Mandatory)]$ComboBox,
        [string]$Value
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { $Value = 'Auto' }
    $index = $ComboBox.Items.IndexOf($Value)
    $ComboBox.SelectedIndex = if ($index -ge 0) { $index } else { 0 }
}

function Add-GuiConfigValue {
    param(
        [Parameter(Mandatory)]$Map,
        [Parameter(Mandatory)][string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) { return }

    try {
        $config = Read-JsonFile -Path $Path
        foreach ($property in @($config.PSObject.Properties)) {
            if ($null -ne $property.Value -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
                $Map[$property.Name] = [string]$property.Value
            }
        }
    }
    catch {
        Write-Verbose ("Unable to read GUI config values from '{0}': {1}" -f $Path, $_.Exception.Message)
    }
}

function Resolve-GuiTokenValue {
    param(
        [Parameter(Mandatory)][string]$Value,
        [Parameter(Mandatory)]$Map
    )

    $resolved = $Value
    for ($pass = 0; $pass -lt 10; $pass++) {
        $before = $resolved
        foreach ($key in @($Map.Keys)) {
            $token = '{{' + [string]$key + '}}'
            if ($resolved.Contains($token)) {
                $resolved = $resolved.Replace($token, [string]$Map[$key])
            }
        }
        if ($resolved -eq $before) { break }
    }

    return $resolved
}

function Get-TemplateLanguageFromFile {
    param(
        [Parameter(Mandatory)]$Campaign,
        [Parameter(Mandatory)]$File
    )

    $baseName = [string](Get-CampaignConfigValue -Campaign $Campaign -Name 'TemplateBaseName' -DefaultValue '')
    if ([string]::IsNullOrWhiteSpace($baseName)) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
    }

    $nameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
    if ($nameWithoutExtension -eq $baseName) { return 'default' }

    $prefix = $baseName + '.'
    if ($nameWithoutExtension.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $nameWithoutExtension.Substring($prefix.Length)
    }

    return $nameWithoutExtension
}

function Resolve-GuiTokenizedPath {
    param(
        [string]$Value,
        [AllowNull()]$Campaign
    )

    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    if ($Value -notmatch '\{\{') { return $Value }

    try {
        $globalTemplatePath = Find-UpwardPath -RelativePath 'Config\SmartM365.global.local.json.template'
        $globalLocalPath = Find-UpwardPath -RelativePath 'Config\SmartM365.global.local.json'
        $smartM365Root = if (-not [string]::IsNullOrWhiteSpace($globalLocalPath)) {
            Split-Path -Path $globalLocalPath -Parent
        }
        elseif (-not [string]::IsNullOrWhiteSpace($globalTemplatePath)) {
            Split-Path -Path $globalTemplatePath -Parent
        }
        else {
            Split-Path -Path (Get-ExchangeUserNotificationsRoot) -Parent
        }
        $selectedCampaign = if ($Campaign) { $Campaign } else { Get-SelectedCampaign }
        $campaignRootPath = if ($selectedCampaign -and -not [string]::IsNullOrWhiteSpace([string]$selectedCampaign.ScriptPath)) {
            Split-Path -Path ([string]$selectedCampaign.ScriptPath) -Parent
        }
        else {
            Get-ExchangeUserNotificationsRoot
        }

        $map = [ordered]@{
            DefaultTenant = (Get-DefaultTenantProfile)
            TenantKey = (Get-SelectedTenantProfile)
            CampaignRootPath = $campaignRootPath
            SmartM365RootPath = $smartM365Root
            WorkspaceRootPath = '{{SmartM365RootPath}}'
            DataAllRootPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-ALL'
            LatestCsvFolderPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-LAST'
            LogAllRootPath = '{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\LOG-ALL'
        }

        if (-not [string]::IsNullOrWhiteSpace($globalTemplatePath)) { Add-GuiConfigValue -Map $map -Path $globalTemplatePath }
        if (-not [string]::IsNullOrWhiteSpace($globalLocalPath)) { Add-GuiConfigValue -Map $map -Path $globalLocalPath }

        if ([string]::IsNullOrWhiteSpace([string]$map['TenantKey']) -and $map.Contains('DefaultTenant') -and -not [string]::IsNullOrWhiteSpace([string]$map['DefaultTenant'])) {
            $map['TenantKey'] = [string]$map['DefaultTenant']
        }

        $tenantConfigPath = Join-Path -Path $smartM365Root -ChildPath ('Config\Tenants\{0}.local.json' -f $map['TenantKey'])
        Add-GuiConfigValue -Map $map -Path $tenantConfigPath

        $map['SmartM365RootPath'] = $smartM365Root
        if ($map['WorkspaceRootPath'] -eq '.' -or $map['WorkspaceRootPath'] -eq '{{SmartM365RootPath}}') {
            $map['WorkspaceRootPath'] = $smartM365Root
        }

        return Resolve-GuiTokenValue -Value ([string]$Value) -Map $map
    }
    catch {
        Write-Verbose ("Unable to resolve recipient path '{0}': {1}" -f $Value, $_.Exception.Message)
        return $Value
    }
}

function Get-CampaignTemplateRootPath {
    param([Parameter(Mandatory)]$Campaign)

    $defaultPath = $Campaign.TemplateFolderPath
    $configuredPath = Get-CampaignConfigValue -Campaign $Campaign -Name 'TemplateRootPath' -DefaultValue $defaultPath
    if ([string]::IsNullOrWhiteSpace($configuredPath)) { $configuredPath = $defaultPath }
    return Resolve-GuiTokenizedPath -Value $configuredPath -Campaign $Campaign
}

function Get-CampaignTemplateRootPaths {
    param([Parameter(Mandatory)]$Campaign)

    $configuredPath = Get-CampaignTemplateRootPath -Campaign $Campaign
    return @($configuredPath, $Campaign.TemplateFolderPath) |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
        Select-Object -Unique
}

function Get-RecipientRowCount {
    param([Parameter(Mandatory)][string]$Path)

    $result = [ordered]@{
        Exists = $false
        CsvFiles = 0
        Rows = 0
        Mailboxes = 0
        Error = ''
    }

    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return [pscustomobject]$result }
    $result.Exists = $true

    $files = @(if (Test-Path -LiteralPath $Path -PathType Container) {
        Get-ChildItem -LiteralPath $Path -Filter '*.csv' -File | Sort-Object Name
    }
    else {
        Get-Item -LiteralPath $Path
    })
    $result.CsvFiles = $files.Count

    $emailColumns = @('PrimarySmtpAddress','PrimarySMTPaddress','EmailAddress','Mail','UserPrincipalName','Recipient')
    foreach ($file in $files) {
        try {
            $rows = @(Import-Csv -LiteralPath $file.FullName)
            $result.Rows += $rows.Count
            foreach ($row in $rows) {
                foreach ($column in $emailColumns) {
                    $property = $row.PSObject.Properties[$column]
                    if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
                        $result.Mailboxes++
                        break
                    }
                }
            }
        }
        catch {
            $result.Error = $_.Exception.Message
        }
    }

    return [pscustomobject]$result
}

function Get-CampaignTemplateLanguage {
    param([Parameter(Mandatory)]$Campaign)

    $languages = New-Object System.Collections.Generic.List[string]
    foreach ($templateFolderPath in (Get-CampaignTemplateRootPaths -Campaign $Campaign)) {
        if (-not (Test-Path -LiteralPath $templateFolderPath -PathType Container)) { continue }
        foreach ($file in (Get-ChildItem -LiteralPath $templateFolderPath -Filter '*.html' -File | Sort-Object Name)) {
            $name = $file.BaseName
            $language = if ($name -match '\.([a-z]{2}(?:-[A-Z]{2})?)$') { $Matches[1] } else { 'default' }
            if (-not $languages.Contains($language)) { $languages.Add($language) }
        }
    }

    return @($languages | Sort-Object)
}

function Get-FileStatusText {
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$ReadyText = 'Ready',
        [string]$MissingText = 'Missing'
    )

    if (Test-Path -LiteralPath $Path) { return $ReadyText }
    return $MissingText
}

function Test-JsonFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [switch]$Optional
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        if ($Optional) { return @{ Ok = $true; Message = 'Optional file not present.' } }
        return @{ Ok = $false; Message = "Missing: $Path" }
    }

    try {
        Read-JsonFile -Path $Path | Out-Null
        return @{ Ok = $true; Message = "JSON OK: $Path" }
    }
    catch {
        return @{ Ok = $false; Message = "JSON invalid: $Path - $($_.Exception.Message)" }
    }
}

function Test-ScriptParser {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return @("Missing: $Path") }

    try {
        $tokens = $null
        $errors = $null
        [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors) | Out-Null
        if ($errors -and $errors.Count -gt 0) {
            return @($errors | ForEach-Object { "Parser: $($_.Message)" })
        }
        return @()
    }
    catch {
        return @("Parser failed: $($_.Exception.Message)")
    }
}

function Get-TemplateToken {
    param([Parameter(Mandatory)][string]$Content)

    @([regex]::Matches($Content, '\{\{([^}]+)\}\}') | ForEach-Object { $_.Groups[1].Value } | Sort-Object -Unique)
}

function Get-CampaignEffectiveTemplateFiles {
    param([Parameter(Mandatory)]$Campaign)

    $templateFolderPaths = @(Get-CampaignTemplateRootPaths -Campaign $Campaign)
    $seenTemplateNames = @{}
    $rows = @()

    foreach ($templateFolderPath in $templateFolderPaths) {
        if (-not (Test-Path -LiteralPath $templateFolderPath -PathType Container)) { continue }
        foreach ($file in (Get-ChildItem -LiteralPath $templateFolderPath -Filter '*.html' -File | Sort-Object Name)) {
            $templateKey = $file.Name.ToLowerInvariant()
            if ($seenTemplateNames.ContainsKey($templateKey)) { continue }
            $seenTemplateNames[$templateKey] = $true
            $source = if ($templateFolderPath -eq $Campaign.TemplateFolderPath) { 'Built-in' } else { 'Custom' }
            $language = Get-TemplateLanguageFromFile -Campaign $Campaign -File $file
            $row = [pscustomobject]@{
                Language = $language
                FileName = $file.Name
                Path = $file.FullName
                Modified = $file.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
                Source = $source
            }
            $rows += $row
        }
    }

    return @($rows)
}

function Test-CampaignTemplate {
    param([Parameter(Mandatory)]$Campaign)

    $issues = New-Object System.Collections.ArrayList
    $templateFolderPaths = @(Get-CampaignTemplateRootPaths -Campaign $Campaign)
    $existingTemplateFolderPaths = @($templateFolderPaths | Where-Object { Test-Path -LiteralPath $_ -PathType Container })
    if ($existingTemplateFolderPaths.Count -eq 0) {
        [void]$issues.Add("Missing template folder: $($templateFolderPaths -join ', ')")
        return @($issues)
    }

    $expected = @($Campaign.ExpectedTokens | Sort-Object -Unique)
    $templateFiles = @(Get-CampaignEffectiveTemplateFiles -Campaign $Campaign)
    if ($templateFiles.Count -eq 0) { [void]$issues.Add("No HTML template found in $($templateFolderPaths -join ', ')") }

    foreach ($file in $templateFiles) {
        try {
            $content = Get-Content -LiteralPath $file.Path -Raw -Encoding UTF8 -ErrorAction Stop
        }
        catch {
            [void]$issues.Add(("Template unreadable: {0}; Error={1}" -f $file.Path, $_.Exception.Message))
            continue
        }

        if ([string]::IsNullOrWhiteSpace($content)) {
            [void]$issues.Add(("Template empty: {0}" -f $file.Path))
            continue
        }

        $found = @(Get-TemplateToken -Content $content)
        $missing = @($expected | Where-Object { $_ -notin $found })
        $extra = @($found | Where-Object { $_ -notin $expected })
        if ($missing.Count -gt 0 -or $extra.Count -gt 0) {
            [void]$issues.Add(("Template tokens mismatch: {0} ({1}); Missing=[{2}]; Extra=[{3}]" -f $file.FileName, $file.Source, ($missing -join ', '), ($extra -join ', ')))
        }
    }

    return @($issues)
}

function Test-Campaign {
    param(
        [Parameter(Mandatory)]$Campaign,
        [switch]$SkipTemplates
    )

    $issues = New-Object System.Collections.ArrayList
    foreach ($issue in (Test-ScriptParser -Path $Campaign.ScriptPath)) { [void]$issues.Add($issue) }

    foreach ($jsonCheck in @(
        (Test-JsonFile -Path $Campaign.TemplateConfigPath),
        (Test-JsonFile -Path $Campaign.LocalConfigPath -Optional)
    )) {
        if (-not $jsonCheck.Ok) { [void]$issues.Add($jsonCheck.Message) }
    }

    if (-not $SkipTemplates) {
        foreach ($issue in (Test-CampaignTemplate -Campaign $Campaign)) { [void]$issues.Add($issue) }
    }
    return @($issues)
}

function Get-SelectedCampaign {
    $index = $script:SelectedCampaignIndex
    if ($index -lt 0) { $index = 0 }
    if ($index -ge $script:Campaigns.Count) { $index = 0 }
    return $script:Campaigns[$index]
}

function Select-CampaignTab {
    param([Parameter(Mandatory)][int]$Index)

    if ($Index -lt 0 -or $Index -ge $script:Campaigns.Count) { return }
    $script:SelectedCampaignIndex = $Index

    $campaignTabs = @(
        $script:ExchangeMigrationCampaignTab,
        $script:ExchangeArchiveCampaignTab,
        $script:ExchangeSizeReductionCampaignTab
    )
    $targetTab = $campaignTabs[$Index]
    if ($targetTab -and $script:RunCampaignContent -and $targetTab.Content -ne $script:RunCampaignContent) {
        $currentParent = $script:RunCampaignContent.Parent
        if ($currentParent -is [System.Windows.Controls.ContentControl]) {
            $currentParent.Content = $null
        }
        $targetTab.Content = $script:RunCampaignContent
    }

    Show-CampaignView
}

function Add-OutputLine {
    param([string]$Text)

    if (-not $script:OutputBox) { return }
    $action = [Action[string]]{
        param([string]$Line)
        $script:OutputBox.AppendText(("{0}`r`n" -f $Line))
        $script:OutputBox.ScrollToEnd()
    }

    if ($script:OutputBox.Dispatcher.CheckAccess()) { $action.Invoke($Text) }
    else { [void]$script:OutputBox.Dispatcher.BeginInvoke($action, $Text) }
}

function Show-Status {
    param(
        [string]$Text,
        [string]$Color = '#5F6B7A'
    )

    if (-not $script:StatusText) { return }
    $statusTextValue = $Text
    $statusColorValue = $Color
    $action = [Action]{
        $script:StatusText.Text = $statusTextValue
        $script:StatusDot.Fill = $script:BrushConverter.ConvertFromString($statusColorValue)
    }

    if ($script:StatusText.Dispatcher.CheckAccess()) { $action.Invoke() }
    else { [void]$script:StatusText.Dispatcher.BeginInvoke($action) }
}

function Show-RunState {
    param([bool]$Running)

    $script:RunButton.IsEnabled = -not $Running
    $script:TenantProfileCombo.IsEnabled = -not $Running
    foreach ($campaignTab in @($script:ExchangeMigrationCampaignTab, $script:ExchangeArchiveCampaignTab, $script:ExchangeSizeReductionCampaignTab)) {
        if ($campaignTab) {
            $campaignTabIndex = -1
            [void][int]::TryParse([string]$campaignTab.Tag, [ref]$campaignTabIndex)
            $campaignTab.IsEnabled = (-not $Running -or $campaignTabIndex -eq $script:SelectedCampaignIndex)
        }
    }
    $script:MailSendModeCombo.IsEnabled = -not $Running
    $script:ExchangeManagementModeCombo.IsEnabled = -not $Running
    $script:EffectiveDatePicker.IsEnabled = -not $Running
    $script:TeamsUserMessageCheck.IsEnabled = ((-not $Running) -and $script:TeamsUserMessageAvailable)
    $script:DryRunRadio.IsEnabled = -not $Running
    $script:LiveRadio.IsEnabled = -not $Running
    $script:Window.Cursor = if ($Running) { [System.Windows.Input.Cursors]::Wait } else { [System.Windows.Input.Cursors]::Arrow }
}

function Get-RunMode {
    if ($script:LiveRadio.IsChecked) { return 'Live' }
    return 'DryRun'
}

function Get-CommandArgument {
    $campaign = Get-SelectedCampaign
    $commandArgs = New-Object System.Collections.Generic.List[string]

    $commandArgs.Add('-NoProfile')
    $commandArgs.Add('-ExecutionPolicy')
    $commandArgs.Add('Bypass')
    $commandArgs.Add('-File')
    $commandArgs.Add($campaign.ScriptPath)
    $tenantProfile = Get-SelectedTenantProfile
    if (-not [string]::IsNullOrWhiteSpace($tenantProfile)) {
        $commandArgs.Add('-Tenant')
        $commandArgs.Add($tenantProfile)
    }
    $pathValue = $script:RecipientPathBox.Text.Trim()
    if ($campaign.SupportsFromList) {
        $commandArgs.Add('-FromList')
        if (-not [string]::IsNullOrWhiteSpace($pathValue)) {
            $commandArgs.Add('-ListCsvPath')
            $commandArgs.Add($pathValue)
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($pathValue)) {
        $commandArgs.Add($campaign.RecipientPathParameter)
        $commandArgs.Add($pathValue)
    }

    $forceLanguage = [string]$script:ForceLanguageCombo.SelectedItem
    if ($forceLanguage -eq $script:AutoLanguageLabel) { $forceLanguage = '' }
    if (-not [string]::IsNullOrWhiteSpace($forceLanguage)) {
        $commandArgs.Add('-ForceLanguage')
        $commandArgs.Add($forceLanguage)
    }

    if ($campaign.SupportsEffectiveDate -and $script:EffectiveDatePicker -and $script:EffectiveDatePicker.SelectedDate.HasValue) {
        $effectiveDate = $script:EffectiveDatePicker.SelectedDate.Value.ToString('yyyy-MM-dd')
        if (-not [string]::IsNullOrWhiteSpace($effectiveDate)) {
            $commandArgs.Add('-ForceEffectiveDate')
            $commandArgs.Add($effectiveDate)
        }
    }

    $mailSendMode = [string]$script:MailSendModeCombo.SelectedItem
    if (-not [string]::IsNullOrWhiteSpace($mailSendMode)) {
        $commandArgs.Add('-MailSendMode')
        $commandArgs.Add($mailSendMode)
    }

    $exchangeManagementMode = [string]$script:ExchangeManagementModeCombo.SelectedItem
    if (-not [string]::IsNullOrWhiteSpace($exchangeManagementMode)) {
        $commandArgs.Add('-ExchangeManagementMode')
        $commandArgs.Add($exchangeManagementMode)
    }

    if ($script:ForceSendCheck.IsChecked) { $commandArgs.Add('-ForceSend') }
    if ($script:NoSummaryCheck.IsChecked) { $commandArgs.Add('-NoSummaryEmail') }
    $commandArgs.Add('-TeamsUserMessageMode')
    if ($script:TeamsUserMessageCheck.IsChecked) { $commandArgs.Add('GraphDelegated') } else { $commandArgs.Add('Disabled') }
    if ($campaign.SupportsSkipConfirmation -and $script:SkipConfirmationCheck.IsChecked) { $commandArgs.Add('-SkipConfirmation') }
    if ((Get-RunMode) -eq 'DryRun') { $commandArgs.Add('-WhatIf') }

    return @($commandArgs)
}

function Get-CommandPreview {
    $pwsh = Get-PowerShellPath
    $commandArgs = Get-CommandArgument
    return ((ConvertTo-QuotedArgument $pwsh) + ' ' + (($commandArgs | ForEach-Object { ConvertTo-QuotedArgument $_ }) -join ' '))
}

function Show-CommandPreview {
    if (-not $script:UiReady) { return }
    if (-not $script:CommandPreviewBox) { return }
    try { $script:CommandPreviewBox.Text = Get-CommandPreview }
    catch { $script:CommandPreviewBox.Text = $_.Exception.Message }
}

function Show-RecipientEstimate {
    if (-not $script:UiReady) { return }

    $campaign = Get-SelectedCampaign
    if ($campaign.SupportsInventory -and $script:InventoryRadio.IsChecked) {
        $script:RecipientCountText.Text = 'Mailboxes: runtime'
        $script:RecipientPathStatusText.Text = 'Inventory mode: candidates are selected at runtime from inventory and live checks.'
        return
    }

    $configuredPath = $script:RecipientPathBox.Text.Trim()
    $pathSource = 'configured'

    if ([string]::IsNullOrWhiteSpace($configuredPath)) {
        $script:RecipientCountText.Text = 'Mailboxes: -'
        $script:RecipientPathStatusText.Text = 'No recipients path configured.'
        return
    }

    $resolvedPath = Resolve-GuiTokenizedPath -Value $configuredPath
    if ($resolvedPath -match '\{\{') {
        $script:RecipientCountText.Text = 'Mailboxes: -'
        $script:RecipientPathStatusText.Text = ('Unable to resolve {0} path yet.' -f $pathSource)
        return
    }

    $count = Get-RecipientRowCount -Path $resolvedPath
    if (-not $count.Exists) {
        $script:RecipientCountText.Text = 'Mailboxes: 0'
        $script:RecipientPathStatusText.Text = ('No recipients found: {0} path does not exist.' -f $pathSource)
        return
    }
    if ($count.CsvFiles -eq 0) {
        $script:RecipientCountText.Text = 'Mailboxes: 0'
        $script:RecipientPathStatusText.Text = ('No CSV file found in {0} path.' -f $pathSource)
        return
    }
    if (-not [string]::IsNullOrWhiteSpace($count.Error)) {
        $script:RecipientCountText.Text = ('Mailboxes: {0}' -f $count.Mailboxes)
        $script:RecipientPathStatusText.Text = ('Recipient count warning: {0}' -f $count.Error)
        return
    }

    $script:RecipientCountText.Text = ('Mailboxes: {0}' -f $count.Mailboxes)
    $script:RecipientPathStatusText.Text = ('{0} CSV file(s), {1} row(s), from {2} path.' -f $count.CsvFiles, $count.Rows, $pathSource)
}

function Show-TemplateSummary {
    param([Parameter(Mandatory)]$Campaign)

    $templateFolderPaths = @(Get-CampaignTemplateRootPaths -Campaign $Campaign)
    $existingTemplateFolderPaths = @($templateFolderPaths | Where-Object { Test-Path -LiteralPath $_ -PathType Container })
    if ($existingTemplateFolderPaths.Count -eq 0) {
        $script:TemplateSummaryText.Text = 'No template folder found.'
        return
    }

    $files = @(Get-CampaignEffectiveTemplateFiles -Campaign $Campaign)
    $languages = @(Get-CampaignTemplateLanguage -Campaign $Campaign)

    $script:TemplateSummaryText.Text = ('{0} templates - {1}' -f $files.Count, ($languages -join ', '))
}

function Get-CampaignTemplateRows {
    param([Parameter(Mandatory)]$Campaign)

    return @(Get-CampaignEffectiveTemplateFiles -Campaign $Campaign)
}

function Show-TemplatePreview {
    if (-not $script:TemplatePreviewBrowser) { return }

    $selectedTemplate = $script:TemplateList.SelectedItem
    if (-not $selectedTemplate -or [string]::IsNullOrWhiteSpace([string]$selectedTemplate.Path) -or -not (Test-Path -LiteralPath $selectedTemplate.Path)) {
        $script:TemplatePreviewBrowser.NavigateToString('<html><body style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#5F6B7A;padding:18px;">Select a template to preview.</body></html>')
        return
    }

    try {
        $html = Get-Content -LiteralPath $selectedTemplate.Path -Raw -Encoding UTF8
        $script:TemplatePreviewBrowser.NavigateToString($html)
    }
    catch {
        $errorHtml = '<html><body style="font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#B42318;padding:18px;">Preview failed: ' + [System.Net.WebUtility]::HtmlEncode($_.Exception.Message) + '</body></html>'
        $script:TemplatePreviewBrowser.NavigateToString($errorHtml)
    }
}

function Show-TemplateList {
    param([Parameter(Mandatory)]$Campaign)

    if (-not $script:TemplateList) { return }

    $templates = @(Get-CampaignTemplateRows -Campaign $Campaign)
    $script:TemplateList.ItemsSource = $templates
    if ($templates.Count -gt 0) {
        $script:TemplateList.SelectedIndex = 0
    }
    else {
        Show-TemplatePreview
    }
}

function Get-SelectedTemplateCampaign {
    if ($script:TemplateCampaignCombo -and $script:TemplateCampaignCombo.SelectedIndex -ge 0 -and $script:TemplateCampaignCombo.SelectedIndex -lt $script:Campaigns.Count) {
        return $script:Campaigns[$script:TemplateCampaignCombo.SelectedIndex]
    }

    return Get-SelectedCampaign
}

function Show-TemplateCampaignView {
    if (-not $script:UiReady) { return }

    $campaign = Get-SelectedTemplateCampaign
    $templateFolderPaths = @(Get-CampaignTemplateRootPaths -Campaign $campaign)
    $script:TemplateFolderPathText.Text = if ($templateFolderPaths.Count -gt 1) {
        ('{0} (fallback: {1})' -f $templateFolderPaths[0], $templateFolderPaths[1])
    }
    else {
        $templateFolderPaths[0]
    }
    Show-TemplateSummary -Campaign $campaign
    Show-TemplateList -Campaign $campaign
}

function Set-TemplateCampaignSelection {
    param([Parameter(Mandatory)][int]$Index)

    if (-not $script:TemplateCampaignCombo) { return }
    if ($Index -lt 0 -or $Index -ge $script:Campaigns.Count) { $Index = 0 }

    if ($script:TemplateCampaignCombo.SelectedIndex -ne $Index) {
        $script:TemplateCampaignCombo.SelectedIndex = $Index
    }
    else {
        Show-TemplateCampaignView
    }
}

function Show-CampaignView {
    if (-not $script:UiReady) { return }

    $campaign = Get-SelectedCampaign
    $script:CampaignBadgeText.Text = $campaign.Badge
    $script:CampaignTitleText.Text = $campaign.Name
    $script:CampaignDescriptionText.Text = $campaign.Description
    $script:ScriptPathText.Text = Get-FileStatusText -Path $campaign.ScriptPath -ReadyText 'Ready' -MissingText 'Missing'
    $script:LocalConfigPathText.Text = Get-FileStatusText -Path $campaign.LocalConfigPath -ReadyText 'Present' -MissingText 'Optional local config not found. Template defaults will be used.'
    $script:TemplateConfigPathText.Text = Get-FileStatusText -Path $campaign.TemplateConfigPath -ReadyText 'Ready' -MissingText 'Missing'
    $script:ScriptPathText.ToolTip = $campaign.ScriptPath
    $script:LocalConfigPathText.ToolTip = $campaign.LocalConfigPath
    $script:TemplateConfigPathText.ToolTip = $campaign.TemplateConfigPath
    $script:OpenLocalConfigButton.IsEnabled = Test-Path -LiteralPath $campaign.LocalConfigPath
    Show-CampaignConfigRows -Campaign $campaign
    Set-TemplateCampaignSelection -Index $script:SelectedCampaignIndex
    $mailSendMode = Get-SharedConfigValue -Name 'MailSendMode' -DefaultValue 'Auto'
    Select-ComboBoxValue -ComboBox $script:MailSendModeCombo -Value $mailSendMode
    $script:MailSendModeCombo.ToolTip = 'Auto uses Graph when SMTP server is empty, otherwise SMTP relay. Graph and SmtpRelay force a specific mode.'
    $exchangeManagementMode = Get-CampaignConfigValue -Campaign $campaign -Name 'ExchangeManagementMode' -DefaultValue 'Auto'
    Select-ComboBoxValue -ComboBox $script:ExchangeManagementModeCombo -Value $exchangeManagementMode
    $script:ExchangeManagementModeCombo.ToolTip = 'Auto tries Exchange Online first, then Exchange 2016 fallback when enabled. ExchangeOnline and Exchange2016 force a specific mode.'
    $teamsUserMessageState = Get-TeamsUserMessageGuiState -Campaign $campaign
    $script:TeamsUserMessageAvailable = [bool]$teamsUserMessageState.Available
    $script:TeamsUserMessageCheck.IsChecked = [bool]$teamsUserMessageState.Checked
    $script:TeamsUserMessageCheck.IsEnabled = [bool]$teamsUserMessageState.Available
    $script:TeamsUserMessageCheck.ToolTip = [string]$teamsUserMessageState.ToolTip
    $script:SkipConfirmationCheck.Visibility = if ($campaign.SupportsSkipConfirmation) { 'Visible' } else { 'Collapsed' }
    $script:SourceModePanel.Visibility = 'Collapsed'
    $script:EffectiveDatePanel.Visibility = if ($campaign.SupportsEffectiveDate) { 'Visible' } else { 'Collapsed' }
    $defaultEffectiveDateText = if ($campaign.SupportsEffectiveDate) {
        Get-CampaignConfigValue -Campaign $campaign -Name 'DefaultEffectiveDate'
    }
    else {
        ''
    }
    $defaultEffectiveDate = [datetime]::MinValue
    $script:EffectiveDatePicker.SelectedDate = if (-not [string]::IsNullOrWhiteSpace($defaultEffectiveDateText) -and [datetime]::TryParseExact($defaultEffectiveDateText.Trim(), 'yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$defaultEffectiveDate)) {
        $defaultEffectiveDate
    }
    else {
        $null
    }
    $defaultRecipientPath = Get-CampaignDefaultRecipientPath -Campaign $campaign
    $displayRecipientPath = if ([string]::IsNullOrWhiteSpace($defaultRecipientPath)) {
        ''
    }
    else {
        $resolvedDefaultRecipientPath = Resolve-GuiTokenizedPath -Value $defaultRecipientPath
        if ($resolvedDefaultRecipientPath -match '\{\{') { $defaultRecipientPath } else { $resolvedDefaultRecipientPath }
    }
    $script:RecipientPathLabel.Text = 'Recipients path'
    $script:RecipientPathBox.Text = $displayRecipientPath
    $script:RecipientPathHelpText.Text = if ([string]::IsNullOrWhiteSpace($defaultRecipientPath)) {
        'Enter the recipients path to use for this campaign.'
    }
    else {
        'Loaded from the campaign JSON configuration. You can edit it before running.'
    }
    Show-RecipientEstimate

    $script:ForceLanguageCombo.Items.Clear()
    [void]$script:ForceLanguageCombo.Items.Add($script:AutoLanguageLabel)
    foreach ($language in (Get-CampaignTemplateLanguage -Campaign $campaign)) {
        if ($language -ne 'default') { [void]$script:ForceLanguageCombo.Items.Add($language) }
    }
    $script:ForceLanguageCombo.SelectedIndex = 0

    if ($campaign.SupportsInventory) {
        $script:InventoryRadio.IsChecked = $true
        $script:RecipientPathBox.IsEnabled = $false
        $script:BrowsePathButton.IsEnabled = $false
        $script:BrowseFolderButton.IsEnabled = $false
    }
    else {
        if ($campaign.SupportsFromList) { $script:FromListRadio.IsChecked = $true }
        $script:RecipientPathBox.IsEnabled = $true
        $script:BrowsePathButton.IsEnabled = $true
        $script:BrowseFolderButton.IsEnabled = $true
    }

    Show-CommandPreview
    Show-Status -Text 'Ready' -Color '#0078D4'
}

function Show-ValidationResult {
    param(
        [Parameter(Mandatory)]$Campaign,
        [switch]$TemplatesOnly,
        [switch]$SkipTemplates
    )

    $validationLabel = if ($TemplatesOnly) {
        'Template validation'
    }
    elseif ($SkipTemplates) {
        'Campaign pre-run validation'
    }
    else {
        'Validation'
    }
    $issues = if ($TemplatesOnly) {
        Test-CampaignTemplate -Campaign $Campaign
    }
    else {
        Test-Campaign -Campaign $Campaign -SkipTemplates:$SkipTemplates
    }
    if ($issues.Count -eq 0) {
        Add-OutputLine ('[{0}] {1} OK for {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $validationLabel, $Campaign.Name)
        Show-Status -Text ('{0} OK' -f $validationLabel) -Color '#107C10'
        return $true
    }

    Add-OutputLine ('[{0}] {1} issues for {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $validationLabel, $Campaign.Name)
    foreach ($issue in $issues) { Add-OutputLine ("  {0}" -f $issue) }
    Show-Status -Text ('{0} failed' -f $validationLabel) -Color '#B42318'
    return $false
}

function Open-ExistingPath {
    param([string]$Path)

    if (Test-Path -LiteralPath $Path) {
        Invoke-Item -LiteralPath $Path
        return
    }

    [System.Windows.MessageBox]::Show("Path not found:`r`n$Path", 'SmartM365', 'OK', 'Warning') | Out-Null
}

function Open-ExternalUrl {
    param([Parameter(Mandatory)][string]$Url)

    try {
        $psi = [System.Diagnostics.ProcessStartInfo]::new($Url)
        $psi.UseShellExecute = $true
        [System.Diagnostics.Process]::Start($psi) | Out-Null
    }
    catch {
        [System.Windows.MessageBox]::Show("Unable to open:`r`n$Url`r`n`r`n$($_.Exception.Message)", 'SmartM365', 'OK', 'Warning') | Out-Null
    }
}

function Select-CsvPath {
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'Select CSV file'
    $dialog.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:RecipientPathBox.Text = $dialog.FileName
        Show-CommandPreview
    }
}

function Select-FolderPath {
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = 'Select recipients folder'
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:RecipientPathBox.Text = $dialog.SelectedPath
        Show-CommandPreview
    }
}

function Complete-Run {
    param([int]$ExitCode)

    Add-OutputLine ('ExitCode={0}' -f $ExitCode)
    if ($ExitCode -eq 0) {
        Show-Status -Text 'Run completed' -Color '#107C10'
    }
    else {
        Show-Status -Text 'Run failed' -Color '#B42318'
    }
    $script:CurrentProcess = $null
    Show-RunState -Running $false
}

function Invoke-CurrentRunStop {
    if ($script:CurrentProcess -and -not $script:CurrentProcess.HasExited) {
        try {
            $script:CurrentProcess.Kill()
            Add-OutputLine 'Run stop requested.'
        }
        catch {
            Add-OutputLine ("Failed to stop run: {0}" -f $_.Exception.Message)
        }
    }
}

function Invoke-CampaignRun {
    $campaign = Get-SelectedCampaign
    Show-Status -Text 'Checking templates' -Color '#0078D4'
    if (-not (Show-ValidationResult -Campaign $campaign -TemplatesOnly)) {
        [System.Windows.MessageBox]::Show('Template validation failed. See Output tab.', 'SmartM365', 'OK', 'Warning') | Out-Null
        return
    }

    Show-Status -Text 'Checking campaign' -Color '#0078D4'
    if (-not (Show-ValidationResult -Campaign $campaign -SkipTemplates)) {
        [System.Windows.MessageBox]::Show('Campaign validation failed. See Output tab.', 'SmartM365', 'OK', 'Warning') | Out-Null
        return
    }

    $mode = Get-RunMode
    if ($mode -eq 'Live') {
        $message = "Run LIVE campaign '$($campaign.Name)'?"
        $answer = [System.Windows.MessageBox]::Show($message, 'Confirm live send', 'YesNo', 'Warning')
        if ($answer -ne [System.Windows.MessageBoxResult]::Yes) {
            Add-OutputLine 'Live run cancelled.'
            return
        }
    }

    $pwsh = Get-PowerShellPath
    $commandArgs = Get-CommandArgument
    Add-OutputLine ''
    Add-OutputLine ('[{0}] Starting {1} ({2})' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $campaign.Name, $mode)
    Show-Status -Text 'Running' -Color '#0078D4'
    Show-RunState -Running $true

    try {
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $pwsh
        foreach ($arg in $commandArgs) { [void]$psi.ArgumentList.Add($arg) }
        $psi.WorkingDirectory = Split-Path -Path $campaign.ScriptPath -Parent
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.CreateNoWindow = $true

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $psi
        $process.EnableRaisingEvents = $true
        $process.add_OutputDataReceived({
            param($processSender, $dataEvent)
            $null = $processSender
            if (-not [string]::IsNullOrWhiteSpace($dataEvent.Data)) {
                $line = $dataEvent.Data
                [void]$script:Window.Dispatcher.BeginInvoke([Action]{ Add-OutputLine $line })
            }
        })
        $process.add_ErrorDataReceived({
            param($processSender, $dataEvent)
            $null = $processSender
            if (-not [string]::IsNullOrWhiteSpace($dataEvent.Data)) {
                $line = 'STDERR: ' + $dataEvent.Data
                [void]$script:Window.Dispatcher.BeginInvoke([Action]{ Add-OutputLine $line })
            }
        })
        $process.add_Exited({
            param($processSender, $exitEvent)
            $null = $exitEvent
            $exitCode = $processSender.ExitCode
            [void]$script:Window.Dispatcher.BeginInvoke([Action]{ Complete-Run -ExitCode $exitCode })
        })

        $script:CurrentProcess = $process
        [void]$process.Start()
        $process.BeginOutputReadLine()
        $process.BeginErrorReadLine()
    }
    catch {
        Add-OutputLine ("Run failed: {0}" -f $_.Exception.Message)
        Show-Status -Text 'Run failed' -Color '#B42318'
        $script:CurrentProcess = $null
        Show-RunState -Running $false
    }
}

$splashHelperPath = Find-UpwardPath -RelativePath 'SmartM365.GuiSplash.ps1'
if (-not [string]::IsNullOrWhiteSpace($splashHelperPath)) {
    . $splashHelperPath
    $script:GuiSplash = Start-SmartM365GuiSplash -Framework Wpf -ProductName 'Exchange User Notifications' -BadgeText '' -Subtitle 'Powered by WorkplaceCloudHub.com' -ShowSiteUrl $false
}
else {
    Show-LoadingWindow -Message 'Preparing campaign launcher...'
}

$root = Get-ExchangeUserNotificationsRoot
$script:Campaigns = @(Get-CampaignDefinition -RootPath $root)
Initialize-GuiLocalJsonFilesFromTemplates -Campaigns $script:Campaigns
$script:BrushConverter = New-Object System.Windows.Media.BrushConverter

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="SmartM365 Exchange User Notifications"
        Width="1180"
        Height="780"
        MinWidth="1040"
        MinHeight="700"
        WindowStartupLocation="CenterScreen"
        Background="#F5F8FB">
    <Window.Resources>
        <SolidColorBrush x:Key="PageBrush" Color="#F5F8FB"/>
        <SolidColorBrush x:Key="CardBrush" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="AccentBrush" Color="#0078D4"/>
        <SolidColorBrush x:Key="AccentDarkBrush" Color="#005A9E"/>
        <SolidColorBrush x:Key="AccentSoftBrush" Color="#E6F4FF"/>
        <SolidColorBrush x:Key="TextBrush" Color="#1F2937"/>
        <SolidColorBrush x:Key="MutedBrush" Color="#5F6B7A"/>
        <SolidColorBrush x:Key="BorderBrushSoft" Color="#DDE7F0"/>
        <SolidColorBrush x:Key="DangerBrush" Color="#B42318"/>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
        </Style>
        <Style TargetType="Button">
            <Setter Property="Height" Value="34"/>
            <Setter Property="Padding" Value="14,0"/>
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrushSoft}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" CornerRadius="6">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Foreground" Value="{StaticResource TextBrush}"/>
            <Setter Property="BorderBrush" Value="{StaticResource BorderBrushSoft}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="Height" Value="32"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Height" Value="32"/>
        </Style>
        <Style TargetType="TabControl">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="0"/>
        </Style>
        <Style TargetType="TabItem">
            <Setter Property="FontFamily" Value="Segoe UI"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Foreground" Value="{StaticResource MutedBrush}"/>
            <Setter Property="Padding" Value="16,8"/>
            <Setter Property="Margin" Value="0,0,6,0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="TabItem">
                        <Border x:Name="TabBorder" Background="#FFFFFF" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="{TemplateBinding Padding}">
                            <ContentPresenter ContentSource="Header" HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="#F8FBFE"/>
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#B9DDF7"/>
                            </Trigger>
                            <Trigger Property="IsSelected" Value="True">
                                <Setter TargetName="TabBorder" Property="Background" Value="{StaticResource AccentSoftBrush}"/>
                                <Setter TargetName="TabBorder" Property="BorderBrush" Value="#B9DDF7"/>
                                <Setter Property="Foreground" Value="{StaticResource AccentDarkBrush}"/>
                                <Setter Property="FontWeight" Value="SemiBold"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid Margin="24">
        <Grid.RowDefinitions>
            <RowDefinition Height="136"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="38"/>
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="{StaticResource CardBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="20,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="170"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" VerticalAlignment="Center">
                    <Border HorizontalAlignment="Left" Background="{StaticResource AccentSoftBrush}" BorderBrush="#B9DDF7" BorderThickness="1" CornerRadius="12" Padding="10,4">
                        <TextBlock Text="SMARTM365 COMMUNICATIONS" Foreground="{StaticResource AccentDarkBrush}" FontSize="11" FontWeight="SemiBold"/>
                    </Border>
                    <TextBlock Text="Exchange User Notifications" FontSize="27" FontWeight="SemiBold" Margin="0,10,0,0"/>
                    <TextBlock Text="Campaign launcher for migration, archive, and mailbox size reduction notifications" FontSize="13" Foreground="{StaticResource MutedBrush}" Margin="0,5,0,0"/>
                </StackPanel>
                <Border x:Name="HeaderLogoLink" Grid.Column="1" Height="90" VerticalAlignment="Center" Background="#F8FBFE" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="10" Cursor="Hand" ToolTip="Open WorkplaceCloudHub.com">
                    <Grid>
                        <Image x:Name="HeaderLogoImage" Stretch="Uniform" RenderOptions.BitmapScalingMode="HighQuality" SnapsToDevicePixels="True" Visibility="Collapsed"/>
                        <StackPanel x:Name="HeaderLogoFallback" VerticalAlignment="Center">
                            <TextBlock Text="SM365" HorizontalAlignment="Center" FontSize="28" FontWeight="SemiBold" Foreground="{StaticResource AccentBrush}"/>
                            <TextBlock Text="Exchange" HorizontalAlignment="Center" FontSize="11" Foreground="{StaticResource MutedBrush}"/>
                        </StackPanel>
                    </Grid>
                </Border>
            </Grid>
        </Border>

        <Border Grid.Row="1" Background="{StaticResource CardBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="16,12" Margin="0,14,0,0" MinHeight="62">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <Ellipse x:Name="StatusDot" Width="10" Height="10" Fill="{StaticResource AccentBrush}" Margin="0,0,10,0"/>
                    <StackPanel>
                        <TextBlock x:Name="StatusText" Text="Ready" Foreground="{StaticResource AccentBrush}" FontWeight="SemiBold" FontSize="12"/>
                        <TextBlock Text="Dry Run is selected by default. Live sends require confirmation." Foreground="{StaticResource MutedBrush}" FontSize="12" Margin="0,3,0,0"/>
                    </StackPanel>
                </StackPanel>
                <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center">
                    <Button x:Name="RunButton" Content="Run campaign" Style="{StaticResource PrimaryButton}" MinWidth="140"/>
                </StackPanel>
            </Grid>
        </Border>

        <TabControl x:Name="MainTabs" Grid.Row="2" Margin="0,14,0,0">
            <TabItem x:Name="ExchangeMigrationCampaignTab" Header="Exchange Migration" Tag="0">
                <Grid x:Name="RunCampaignContent" Margin="0,14,0,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="1.05*"/>
                        <ColumnDefinition Width="0.95*"/>
                    </Grid.ColumnDefinitions>

                    <Border Grid.Column="0" Background="{StaticResource CardBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="18" Margin="0,0,7,0">
                        <ScrollViewer VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled">
                        <StackPanel>
                            <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="0,0,0,16">
                                <Border Background="{StaticResource AccentSoftBrush}" BorderBrush="#B9DDF7" BorderThickness="1" CornerRadius="12" Padding="10,4" Margin="0,0,10,0">
                                    <TextBlock x:Name="CampaignBadgeText" Text="MIGRATION" Foreground="{StaticResource AccentDarkBrush}" FontSize="11" FontWeight="SemiBold"/>
                                </Border>
                                <StackPanel>
                                    <TextBlock x:Name="CampaignTitleText" Text="Exchange Migration" FontSize="18" FontWeight="SemiBold"/>
                                    <TextBlock x:Name="CampaignDescriptionText" Text="Mailbox move notification" FontSize="12" Foreground="{StaticResource MutedBrush}" Margin="0,3,0,0"/>
                                </StackPanel>
                            </StackPanel>

                            <TextBlock Text="Configuration profile" FontWeight="SemiBold" Margin="0,0,0,5"/>
                            <ComboBox x:Name="TenantProfileCombo" Margin="0,0,0,10"/>

                            <StackPanel Margin="0,0,0,10">
                                <TextBlock Text="Mode" FontWeight="SemiBold" Margin="0,0,0,6"/>
                                <StackPanel Orientation="Horizontal" Margin="0,5,0,0">
                                    <RadioButton x:Name="DryRunRadio" Content="Dry Run" IsChecked="True" Margin="0,0,14,0"/>
                                    <RadioButton x:Name="LiveRadio" Content="Live"/>
                                </StackPanel>
                            </StackPanel>

                            <StackPanel x:Name="SourceModePanel" Visibility="Collapsed" Margin="0,0,0,10">
                                <TextBlock Text="Recipient source" FontWeight="SemiBold" Margin="0,0,0,6"/>
                                <StackPanel Orientation="Horizontal">
                                    <RadioButton x:Name="InventoryRadio" Content="Inventory selection" IsChecked="True" Margin="0,0,18,0"/>
                                    <RadioButton x:Name="FromListRadio" Content="CSV list"/>
                                </StackPanel>
                            </StackPanel>

                            <Grid Margin="0,0,0,5">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock x:Name="RecipientPathLabel" Text="Recipients path" FontWeight="SemiBold"/>
                                <TextBlock x:Name="RecipientCountText" Grid.Column="1" Text="Mailboxes: -" Foreground="{StaticResource AccentDarkBrush}" FontSize="12" FontWeight="SemiBold" Margin="12,0,0,0"/>
                            </Grid>
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="96"/>
                                    <ColumnDefinition Width="96"/>
                                </Grid.ColumnDefinitions>
                                <TextBox x:Name="RecipientPathBox" Grid.Column="0"/>
                                <Button x:Name="BrowsePathButton" Grid.Column="1" Content="CSV" Margin="8,0,0,0"/>
                                <Button x:Name="BrowseFolderButton" Grid.Column="2" Content="Folder" Margin="8,0,0,0"/>
                            </Grid>
                                <TextBlock x:Name="RecipientPathHelpText" Text="Loaded from the campaign JSON configuration. You can edit it before running." TextWrapping="Wrap" Foreground="{StaticResource MutedBrush}" FontSize="11" Margin="0,4,0,0"/>
                            <TextBlock x:Name="RecipientPathStatusText" Text="" TextWrapping="Wrap" Foreground="{StaticResource AccentDarkBrush}" FontSize="12" FontWeight="SemiBold" Margin="0,4,0,0"/>
                        </StackPanel>
                        </ScrollViewer>
                    </Border>

                    <Border Grid.Column="1" Background="{StaticResource CardBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="18" Margin="7,0,0,0">
                        <StackPanel>
                            <TextBlock Text="Run options" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,14"/>
                            <Grid Margin="0,0,0,12">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Grid.Column="0" Margin="0,0,8,0">
                                    <TextBlock Text="Mail send mode" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                    <ComboBox x:Name="MailSendModeCombo"/>
                                </StackPanel>
                                <StackPanel Grid.Column="1" Margin="8,0,0,0">
                                    <TextBlock Text="Exchange mode" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                    <ComboBox x:Name="ExchangeManagementModeCombo"/>
                                </StackPanel>
                            </Grid>
                            <TextBlock Text="Force language" FontWeight="SemiBold" Margin="0,0,0,5"/>
                            <ComboBox x:Name="ForceLanguageCombo" Margin="0,0,0,4"/>
                            <TextBlock Text="Auto lets the script resolve language from CSV, mailbox, AD, or domain mapping." TextWrapping="Wrap" Foreground="{StaticResource MutedBrush}" FontSize="11" Margin="0,0,0,12"/>
                            <StackPanel x:Name="EffectiveDatePanel" Visibility="Collapsed" Margin="0,0,0,12">
                                <TextBlock Text="Force effective date" FontWeight="SemiBold" Margin="0,0,0,5"/>
                                <DatePicker x:Name="EffectiveDatePicker" Language="en-US"/>
                                <TextBlock Text="Overrides CSV EffectiveDate for every recipient." TextWrapping="Wrap" Foreground="{StaticResource MutedBrush}" FontSize="11" Margin="0,4,0,0"/>
                            </StackPanel>
                            <CheckBox x:Name="ForceSendCheck" Content="Force send" Margin="0,0,0,8"/>
                            <CheckBox x:Name="TeamsUserMessageCheck" Content="Send Teams message" Margin="0,0,0,8"/>
                            <CheckBox x:Name="NoSummaryCheck" Content="No summary email" Margin="0,0,0,8"/>
                            <CheckBox x:Name="SkipConfirmationCheck" Content="Skip campaign confirmation" IsChecked="True" Margin="0,0,0,0"/>
                        </StackPanel>
                    </Border>
                </Grid>
            </TabItem>

            <TabItem x:Name="ExchangeArchiveCampaignTab" Header="Exchange Archive" Tag="1"/>

            <TabItem x:Name="ExchangeSizeReductionCampaignTab" Header="Size Reduction" Tag="2"/>

            <TabItem Header="Configuration">
                <Border Margin="0,14,0,0" Background="{StaticResource CardBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="18">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <Grid Grid.Row="0" Margin="0,0,0,16">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel>
                                <TextBlock Text="Campaign files" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,14"/>
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="1.2*"/>
                                        <ColumnDefinition Width="1.2*"/>
                                        <ColumnDefinition Width="1.2*"/>
                                    </Grid.ColumnDefinitions>
                                    <StackPanel Grid.Column="0" Margin="0,0,16,0">
                                        <TextBlock Text="Script" FontWeight="SemiBold"/>
                                        <TextBlock x:Name="ScriptPathText" TextWrapping="Wrap" Foreground="{StaticResource MutedBrush}" Margin="0,3,0,0"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="1" Margin="0,0,16,0">
                                        <TextBlock Text="Local config" FontWeight="SemiBold"/>
                                        <TextBlock x:Name="LocalConfigPathText" TextWrapping="Wrap" Foreground="{StaticResource MutedBrush}" Margin="0,3,0,0"/>
                                    </StackPanel>
                                    <StackPanel Grid.Column="2">
                                        <TextBlock Text="Template config" FontWeight="SemiBold"/>
                                        <TextBlock x:Name="TemplateConfigPathText" TextWrapping="Wrap" Foreground="{StaticResource MutedBrush}" Margin="0,3,0,0"/>
                                    </StackPanel>
                                </Grid>
                            </StackPanel>
                            <StackPanel Grid.Column="1" MinWidth="150" Margin="18,0,0,0" VerticalAlignment="Top">
                                <Button x:Name="OpenLocalConfigButton" Content="Open local config"/>
                            </StackPanel>
                        </Grid>

                        <Grid Grid.Row="1">
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <TextBlock Text="Local config parameters" FontSize="18" FontWeight="SemiBold" Margin="0,0,0,12"/>
                            <ListView x:Name="ConfigParameterList" Grid.Row="1" MinHeight="290" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1">
                                <ListView.View>
                                    <GridView>
                                        <GridViewColumn Header="Name" DisplayMemberBinding="{Binding Name}" Width="245"/>
                                        <GridViewColumn Header="Value" DisplayMemberBinding="{Binding Value}" Width="610"/>
                                        <GridViewColumn Header="Source" DisplayMemberBinding="{Binding Source}" Width="100"/>
                                    </GridView>
                                </ListView.View>
                            </ListView>
                        </Grid>
                    </Grid>
                </Border>
            </TabItem>

            <TabItem Header="Templates">
                <Border Margin="0,14,0,0" Background="{StaticResource CardBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="18">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <Grid Grid.Row="0">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="*"/>
                                <ColumnDefinition Width="Auto"/>
                            </Grid.ColumnDefinitions>
                            <StackPanel>
                                <TextBlock Text="Campaign" FontWeight="SemiBold"/>
                                <ComboBox x:Name="TemplateCampaignCombo" Width="320" HorizontalAlignment="Left" Margin="0,5,0,12"/>
                                <TextBlock Text="Template folder" FontWeight="SemiBold"/>
                                <TextBlock x:Name="TemplateFolderPathText" TextWrapping="Wrap" Foreground="{StaticResource MutedBrush}" Margin="0,3,0,8"/>
                                <TextBlock x:Name="TemplateSummaryText" TextWrapping="Wrap" Foreground="{StaticResource AccentDarkBrush}" FontWeight="SemiBold"/>
                            </StackPanel>
                            <StackPanel Grid.Column="1" Orientation="Horizontal" Margin="18,0,0,0">
                                <Button x:Name="OpenTemplateFolderButton" Content="Open folder" MinWidth="110"/>
                            </StackPanel>
                        </Grid>
                        <Grid Grid.Row="1" Margin="0,14,0,0">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="365"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <ListView x:Name="TemplateList" Grid.Column="0" MinHeight="360" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1">
                                <ListView.View>
                                    <GridView>
                                        <GridViewColumn Header="Language" DisplayMemberBinding="{Binding Language}" Width="75"/>
                                        <GridViewColumn Header="Template" DisplayMemberBinding="{Binding FileName}" Width="150"/>
                                        <GridViewColumn Header="Source" DisplayMemberBinding="{Binding Source}" Width="65"/>
                                        <GridViewColumn Header="Modified" DisplayMemberBinding="{Binding Modified}" Width="80"/>
                                    </GridView>
                                </ListView.View>
                            </ListView>
                            <Border Grid.Column="1" Margin="14,0,0,0" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" Background="#FFFFFF">
                                <WebBrowser x:Name="TemplatePreviewBrowser"/>
                            </Border>
                        </Grid>
                    </Grid>
                </Border>
            </TabItem>

            <TabItem Header="Output">
                <Border Margin="0,14,0,0" Background="{StaticResource CardBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="14">
                    <Grid>
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>
                        <StackPanel Grid.Row="0" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,0,0,10">
                            <Button x:Name="ClearOutputButton" Content="Clear output" MinWidth="120"/>
                        </StackPanel>
                        <TextBox x:Name="OutputBox" Grid.Row="1" MinHeight="320" TextWrapping="NoWrap" AcceptsReturn="True" AcceptsTab="True" HorizontalScrollBarVisibility="Auto" VerticalScrollBarVisibility="Auto" IsReadOnly="True" FontFamily="Consolas" FontSize="12"/>
                    </Grid>
                </Border>
            </TabItem>
        </TabControl>

        <Border Grid.Row="3" Background="{StaticResource CardBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="12,0" Margin="0,14,0,0">
            <Grid>
                <TextBlock Text="SmartM365 - Exchange user notification campaigns" VerticalAlignment="Center" Foreground="{StaticResource MutedBrush}" FontSize="12"/>
            </Grid>
        </Border>
    </Grid>
</Window>
'@

$reader = [System.Xml.XmlNodeReader]::new([xml]$xaml)
$script:Window = [Windows.Markup.XamlReader]::Load($reader)

$controlNames = @(
    'HeaderLogoLink','HeaderLogoImage','HeaderLogoFallback','StatusDot','StatusText','RunButton','MainTabs',
    'ExchangeMigrationCampaignTab','ExchangeArchiveCampaignTab','ExchangeSizeReductionCampaignTab','RunCampaignContent',
    'TenantProfileCombo','CampaignBadgeText','CampaignTitleText','CampaignDescriptionText','DryRunRadio','LiveRadio',
    'SourceModePanel','InventoryRadio','FromListRadio','RecipientPathLabel','RecipientPathBox','RecipientPathHelpText','RecipientPathStatusText','RecipientCountText','BrowsePathButton','BrowseFolderButton',
    'MailSendModeCombo','ExchangeManagementModeCombo','ForceLanguageCombo','EffectiveDatePanel','EffectiveDatePicker','ForceSendCheck','TeamsUserMessageCheck','NoSummaryCheck','SkipConfirmationCheck','CommandPreviewBox','ScriptPathText',
    'LocalConfigPathText','TemplateConfigPathText','OpenLocalConfigButton','ConfigParameterList',
    'TemplateCampaignCombo','TemplateFolderPathText','TemplateSummaryText','TemplateList','TemplatePreviewBrowser','OpenTemplateFolderButton','OutputBox','ClearOutputButton'
)
foreach ($name in $controlNames) {
    Set-Variable -Name $name -Scope Script -Value $script:Window.FindName($name)
}

$tenantProfiles = @(Get-TenantProfileNames)
$defaultTenantProfile = Get-DefaultTenantProfile
foreach ($tenantProfile in $tenantProfiles) { [void]$script:TenantProfileCombo.Items.Add($tenantProfile) }
if ($script:TenantProfileCombo.Items.Count -gt 0) {
    $defaultTenantIndex = $script:TenantProfileCombo.Items.IndexOf($defaultTenantProfile)
    $script:TenantProfileCombo.SelectedIndex = if ($defaultTenantIndex -ge 0) { $defaultTenantIndex } else { 0 }
}

foreach ($mode in @('Auto','Graph','SmtpRelay','Disabled')) { [void]$script:MailSendModeCombo.Items.Add($mode) }
foreach ($mode in @('Auto','ExchangeOnline','Exchange2016','Disabled')) { [void]$script:ExchangeManagementModeCombo.Items.Add($mode) }
foreach ($campaign in $script:Campaigns) { [void]$script:TemplateCampaignCombo.Items.Add($campaign.Name) }
if ($script:TemplateCampaignCombo.Items.Count -gt 0) { $script:TemplateCampaignCombo.SelectedIndex = 0 }

$script:OutputBox = $script:Window.FindName('OutputBox')

$iconPath = Find-UpwardPath -RelativePath 'WorkplaceCloudHub.ico'
if (-not [string]::IsNullOrWhiteSpace($iconPath)) {
    try {
        $script:Window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]$iconPath)
    }
    catch {
        Write-Verbose ("Window icon load failed: {0}" -f $_.Exception.Message)
    }
}

$headerLogoPath = Find-UpwardPath -RelativePath 'WorkplaceCloudHub-lockup-WPF.png'
if (-not [string]::IsNullOrWhiteSpace($headerLogoPath)) {
    try {
        $headerLogo = [System.Windows.Media.Imaging.BitmapImage]::new()
        $headerLogo.BeginInit()
        $headerLogo.UriSource = [Uri]$headerLogoPath
        $headerLogo.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $headerLogo.EndInit()
        $script:HeaderLogoImage.Source = $headerLogo
        $script:HeaderLogoImage.Visibility = 'Visible'
        $script:HeaderLogoFallback.Visibility = 'Collapsed'
    }
    catch {
        Write-Verbose ("Header logo load failed: {0}" -f $_.Exception.Message)
    }
}

$script:UiReady = $true

$script:TenantProfileCombo.Add_SelectionChanged({ Show-CommandPreview; Show-RecipientEstimate })
$script:MainTabs.Add_SelectionChanged({
    param($sender, $eventArgs)
    if ($eventArgs.OriginalSource -ne $script:MainTabs) { return }
    $selectedTab = $script:MainTabs.SelectedItem
    if ($selectedTab -is [System.Windows.Controls.TabItem] -and $selectedTab.Tag -ne $null) {
        $campaignIndex = 0
        if ([int]::TryParse([string]$selectedTab.Tag, [ref]$campaignIndex)) {
            Select-CampaignTab -Index $campaignIndex
        }
    }
})
$script:RecipientPathBox.Add_TextChanged({ Show-CommandPreview; Show-RecipientEstimate })
$script:MailSendModeCombo.Add_SelectionChanged({ Show-CommandPreview })
$script:ExchangeManagementModeCombo.Add_SelectionChanged({ Show-CommandPreview })
$script:ForceLanguageCombo.Add_SelectionChanged({ Show-CommandPreview })
$script:EffectiveDatePicker.Add_SelectedDateChanged({ Show-CommandPreview })
$script:DryRunRadio.Add_Checked({ Show-CommandPreview })
$script:LiveRadio.Add_Checked({ Show-CommandPreview })
$script:InventoryRadio.Add_Checked({
    if ($script:UiReady) {
        $script:RecipientPathBox.IsEnabled = $false
        $script:BrowsePathButton.IsEnabled = $false
        $script:BrowseFolderButton.IsEnabled = $false
        Show-CommandPreview
        Show-RecipientEstimate
    }
})
$script:FromListRadio.Add_Checked({
    if ($script:UiReady) {
        $script:RecipientPathBox.IsEnabled = $true
        $script:BrowsePathButton.IsEnabled = $true
        $script:BrowseFolderButton.IsEnabled = $false
        Show-CommandPreview
        Show-RecipientEstimate
    }
})
foreach ($checkBox in @($script:ForceSendCheck, $script:TeamsUserMessageCheck, $script:NoSummaryCheck, $script:SkipConfirmationCheck)) {
    $checkBox.Add_Checked({ Show-CommandPreview })
    $checkBox.Add_Unchecked({ Show-CommandPreview })
}

$script:BrowsePathButton.Add_Click({ Select-CsvPath })
$script:BrowseFolderButton.Add_Click({ Select-FolderPath })
$script:HeaderLogoLink.Add_MouseLeftButtonUp({ Open-ExternalUrl -Url 'https://workplacecloudhub.com' })
$script:TemplateCampaignCombo.Add_SelectionChanged({ Show-TemplateCampaignView })
$script:TemplateList.Add_SelectionChanged({ Show-TemplatePreview })
$script:RunButton.Add_Click({ Invoke-CampaignRun })
$script:OpenLocalConfigButton.Add_Click({ Open-ExistingPath -Path (Get-SelectedCampaign).LocalConfigPath })
$script:OpenTemplateFolderButton.Add_Click({ Open-ExistingPath -Path (Get-SelectedTemplateCampaign).TemplateFolderPath })
$script:ClearOutputButton.Add_Click({ $script:OutputBox.Clear() })
$script:Window.Add_Closing({
    param($windowSender, $closingEvent)
    $null = $windowSender
    if ($script:CurrentProcess -and -not $script:CurrentProcess.HasExited) {
        $answer = [System.Windows.MessageBox]::Show('A campaign is still running. Stop it and close?', 'SmartM365', 'YesNo', 'Warning')
        if ($answer -ne [System.Windows.MessageBoxResult]::Yes) {
            $closingEvent.Cancel = $true
            return
        }
        Invoke-CurrentRunStop
    }
})
$script:Window.Add_ContentRendered({
    if ($script:GuiSplash) { Hide-SmartM365GuiSplash -Splash $script:GuiSplash }
    else { Hide-LoadingWindow }
})

Show-CampaignView
Add-OutputLine 'SmartM365 Exchange User Notifications GUI ready.'
Add-OutputLine 'Dry Run is selected by default.'

[void]$script:Window.ShowDialog()
if ($script:GuiSplash) { Close-SmartM365GuiSplash -Splash $script:GuiSplash }
else { Close-LoadingWindow }
