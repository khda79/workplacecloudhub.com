#Requires -Version 5.1

<#
.SYNOPSIS
    Opens the Smart Device Reboot Manager user GUI.

.DESCRIPTION
    Shows a local WPF restart notification when a device uptime threshold is
    reached. The app provides a clear SmartM365 user interface, file-based
    reminder state, configurable postpone choices, preview mode, and structured
    local logging.

.NOTES
    Version: 1.0
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [int]$RecommendedRestartAfterDays = 5,
    [int]$RequiredRestartAfterDays = 7,
    [int]$RestartCountdownMinutes = 15,
    [int]$DefaultPostponeMinutes = 10,
    [int[]]$PostponeOptionsMinutes = @(),
    [int]$MaxPostponeCount = 2,
    [int]$ReminderCooldownHours = 20,
    [string]$WindowIconPath = 'WorkplaceCloudHub.ico',
    [string]$WindowTitle = '',
    [string]$CompanyName = '',
    [ValidateSet('auto','en','fr','de','es','nl','it','pt','pl','ar','tr','sv','da','nb','fi','ro','hu','ja','ko','zh-Hans','zh','uk')]
    [string]$DefaultLanguage = 'auto',
    [switch]$ForceLanguage,
    [string]$LanguageCatalogPath = '',
    [bool]$SplashEnabled = $true,
    [int]$SplashMinimumDurationMs = 6000,
    [string]$SplashProductName = '',
    [string]$SplashBadgeText = 'WORKPLACECLOUDHUB.COM',
    [string]$SplashSubtitle = 'Powered by WorkplaceCloudHub',
    [string]$SplashLogoPath = 'WorkplaceCloudHub-lockup-WPF.png',
    [switch]$EnableDebugLogging,
    [switch]$PreviewOnly,
    [switch]$TestRequiredRestart,
    [switch]$TestRecommendedRestart,
    [switch]$NeverForceRestart,
    [switch]$ResetRunState,
    [switch]$ValidateOnly,
    [string]$ConfigPath
)

$ErrorActionPreference = 'Stop'
$PSDefaultParameterValues['*:ErrorAction'] = 'Stop'

$script:CliBoundParameters = @{}
foreach ($key in $PSBoundParameters.Keys) {
    $script:CliBoundParameters[$key] = $true
}

$script:AppName = 'Smart Device Reboot Manager'
$script:AppFolderName = 'SmartM365\DeviceRebootManager'
$script:StartTopmost = $true
$script:RestoreTopmostAfterPostpone = $true
$script:RestoreTopmostAfterMinutes = 5
$script:PreviewOnly = [bool]$PreviewOnly
$script:PostponeCount = 0
$script:AllowClose = $true
$script:Ui = @{}

$script:UserDataRoot = Join-Path -Path $env:APPDATA -ChildPath $script:AppFolderName
$script:ProgramDataRoot = Join-Path -Path $env:ProgramData -ChildPath $script:AppFolderName
$script:LogPath = Join-Path -Path $script:UserDataRoot -ChildPath 'SmartM365-DeviceRebootManager.log'
$script:StatePath = Join-Path -Path $script:UserDataRoot -ChildPath 'SmartM365-DeviceRebootManager.state.json'
$script:PreferencesPath = Join-Path -Path $script:UserDataRoot -ChildPath 'SmartM365-DeviceRebootManager.preferences.json'

function Initialize-AppFolder {
    try {
        if (-not (Test-Path -LiteralPath $script:UserDataRoot)) {
            New-Item -ItemType Directory -Path $script:UserDataRoot -Force | Out-Null
        }
        $probePath = Join-Path -Path $script:UserDataRoot -ChildPath ('.write-test-{0}.tmp' -f ([guid]::NewGuid().ToString('N')))
        [IO.File]::WriteAllText($probePath, 'test')
        Remove-Item -LiteralPath $probePath -Force -ErrorAction SilentlyContinue
        return
    }
    catch {
        $primaryError = $_.Exception.Message
        $fallbackRoot = Join-Path -Path $PSScriptRoot -ChildPath 'Logs'
        try {
            if (-not (Test-Path -LiteralPath $fallbackRoot)) {
                New-Item -ItemType Directory -Path $fallbackRoot -Force | Out-Null
            }
        }
        catch {
            $fallbackRoot = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath $script:AppFolderName
            if (-not (Test-Path -LiteralPath $fallbackRoot)) {
                New-Item -ItemType Directory -Path $fallbackRoot -Force | Out-Null
            }
        }

        $script:UserDataRoot = $fallbackRoot
        $script:LogPath = Join-Path -Path $script:UserDataRoot -ChildPath 'SmartM365-DeviceRebootManager.log'
        $script:StatePath = Join-Path -Path $script:UserDataRoot -ChildPath 'SmartM365-DeviceRebootManager.state.json'
        $script:PreferencesPath = Join-Path -Path $script:UserDataRoot -ChildPath 'SmartM365-DeviceRebootManager.preferences.json'
        if ($EnableDebugLogging) {
            Write-Warning "Unable to use AppData for Smart Device Reboot Manager ($primaryError). Falling back to '$script:UserDataRoot'."
        }
    }
}

function Invoke-LaunchLogRotation {
    Initialize-AppFolder

    try {
        if (Test-Path -LiteralPath $script:LogPath) {
            $item = Get-Item -LiteralPath $script:LogPath
            if ($item.Length -gt 0) {
                $stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
                $archivePath = Join-Path -Path $script:UserDataRoot -ChildPath ("SmartM365-DeviceRebootManager-{0}.log" -f $stamp)
                Move-Item -LiteralPath $script:LogPath -Destination $archivePath -Force
            }
        }

        New-Item -ItemType File -Path $script:LogPath -Force | Out-Null

        $keep = 10
        Get-ChildItem -LiteralPath $script:UserDataRoot -Filter 'SmartM365-DeviceRebootManager-*.log' |
            Sort-Object LastWriteTime -Descending |
            Select-Object -Skip $keep |
            ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
    }
    catch {
        if ($EnableDebugLogging) { Write-Warning "Unable to rotate launch log: $($_.Exception.Message)" }
    }
}

function Write-AppLog {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('DEBUG','INFO','WARN','ERROR','SUCCESS')]
        [string]$Level = 'INFO'
    )

    try {
        Initialize-AppFolder
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
        $lines = @([regex]::Split(([string]$Message), '\r?\n') | ForEach-Object { '{0} [{1}] {2}' -f $timestamp, $Level, $_ })
        [IO.File]::AppendAllText($script:LogPath, (($lines -join "`r`n") + "`r`n"), $utf8NoBom)
        if ($EnableDebugLogging) { $lines | ForEach-Object { Write-Host $_ } }
        if ($script:Ui.ContainsKey('ActivityText') -and $null -ne $script:Ui.ActivityText) {
            $uiLine = '[{0}] {1}' -f (Get-Date -Format 'HH:mm:ss'), $Message
            $script:Ui.ActivityText.AppendText("$uiLine`r`n")
            $script:Ui.ActivityText.ScrollToEnd()
        }
    }
    catch {
        if ($EnableDebugLogging) { Write-Warning "Unable to write app log: $($_.Exception.Message)" }
    }
}

function Stop-AppWithError {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [int]$ExitCode = 1
    )

    try { Write-AppLog -Level ERROR -Message $Message } catch { Write-Verbose "Unable to write fatal log entry: $($_.Exception.Message)" }
    try { if ($script:Window -and $script:Window.IsVisible) { $script:Window.Close() } } catch { Write-Verbose "Unable to close GUI after fatal error: $($_.Exception.Message)" }
    [Environment]::Exit($ExitCode)
}

trap {
    $message = if ($_.Exception) { $_.Exception.ToString() } else { $_.ToString() }
    Stop-AppWithError -Message $message -ExitCode 1
}

function Invoke-Safely {
    param([Parameter(Mandatory = $true)][scriptblock]$ScriptBlock)
    try { & $ScriptBlock }
    catch {
        $message = if ($_.Exception) { $_.Exception.ToString() } else { $_.ToString() }
        Stop-AppWithError -Message ("UI handler failed: {0}" -f $message) -ExitCode 3
    }
}

function Get-FirstExistingPath {
    param([string[]]$Paths)

    foreach ($path in $Paths) {
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path)) {
            return $path
        }
    }

    return ''
}

function Resolve-AppRelativePath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    if ([IO.Path]::IsPathRooted($Path)) { return $Path }
    return Join-Path -Path $PSScriptRoot -ChildPath $Path
}

function Initialize-AppConfigFromTemplate {
    param(
        [Parameter(Mandatory = $true)][string]$ConfigPath,
        [Parameter(Mandatory = $true)][string]$TemplatePath
    )

    if (Test-Path -LiteralPath $ConfigPath) { return $false }
    if (-not (Test-Path -LiteralPath $TemplatePath)) { return $false }

    $parent = Split-Path -Path $ConfigPath -Parent
    if (-not [string]::IsNullOrWhiteSpace($parent)) {
        New-Item -Path $parent -ItemType Directory -Force | Out-Null
    }

    Copy-Item -LiteralPath $TemplatePath -Destination $ConfigPath -ErrorAction Stop
    Write-AppLog -Level WARN -Message ("Created local configuration from template: {0}" -f $ConfigPath)
    Write-Host ((@(
        'Created Device Reboot Manager local JSON from template.',
        "Local JSON: $ConfigPath",
        "Template: $TemplatePath",
        'Review the generated local JSON values; continuing with default template values unless edited before next run.'
    )) -join [Environment]::NewLine) -ForegroundColor Yellow
    return $true
}

function Read-AppConfig {
    $scriptConfig = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-DeviceRebootManager-GUI.config.json'
    $scriptTemplateConfig = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-DeviceRebootManager-GUI.config.json.template'
    $programConfig = Join-Path -Path $script:ProgramDataRoot -ChildPath 'SmartM365-DeviceRebootManager-GUI.config.json'
    $userConfig = Join-Path -Path $script:UserDataRoot -ChildPath 'SmartM365-DeviceRebootManager-GUI.config.json'

    if (-not [string]::IsNullOrWhiteSpace($ConfigPath) -and -not (Test-Path -LiteralPath $ConfigPath)) {
        Initialize-AppConfigFromTemplate -ConfigPath $ConfigPath -TemplatePath $scriptTemplateConfig | Out-Null
    }
    elseif (-not (Test-Path -LiteralPath $scriptConfig) -and -not (Test-Path -LiteralPath $programConfig) -and -not (Test-Path -LiteralPath $userConfig)) {
        Initialize-AppConfigFromTemplate -ConfigPath $scriptConfig -TemplatePath $scriptTemplateConfig | Out-Null
    }

    $effectivePath = Get-FirstExistingPath -Paths @($ConfigPath, $scriptConfig, $programConfig, $userConfig)
    if ([string]::IsNullOrWhiteSpace($effectivePath)) {
        Write-AppLog -Level WARN -Message 'No configuration file found. CLI and built-in defaults will be used.'
        return $null
    }

    Write-AppLog -Message ("Loading configuration: {0}" -f $effectivePath)
    return Get-Content -LiteralPath $effectivePath -Raw -Encoding UTF8 | ConvertFrom-Json
}
function Get-ConfigValue {
    param(
        [object]$Config,
        [Parameter(Mandatory = $true)][string]$Name,
        [object]$CurrentValue,
        [ValidateSet('int','intArray','string','bool')]
        [string]$Type = 'string'
    )

    if ($null -eq $Config) { return $CurrentValue }
    $property = $Config.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $CurrentValue }
    if ($script:CliBoundParameters.ContainsKey($Name)) { return $CurrentValue }

    switch ($Type) {
        'int' { return [int]$property.Value }
        'intArray' { return @($property.Value | ForEach-Object { [int]$_ }) }
        'bool' { return [bool]$property.Value }
        default { return [string]$property.Value }
    }
}

function Apply-AppConfig {
    param([object]$Config)

    if ($null -eq $Config) { return }

    $script:StartTopmost = Get-ConfigValue -Config $Config -Name 'StartTopmost' -CurrentValue $script:StartTopmost -Type bool
    $script:RestoreTopmostAfterPostpone = Get-ConfigValue -Config $Config -Name 'RestoreTopmostAfterPostpone' -CurrentValue $script:RestoreTopmostAfterPostpone -Type bool
    $script:RestoreTopmostAfterMinutes = Get-ConfigValue -Config $Config -Name 'RestoreTopmostAfterMinutes' -CurrentValue $script:RestoreTopmostAfterMinutes -Type int

    $script:Effective = [ordered]@{
        RecommendedRestartAfterDays = Get-ConfigValue -Config $Config -Name 'RecommendedRestartAfterDays' -CurrentValue $RecommendedRestartAfterDays -Type int
        RequiredRestartAfterDays = Get-ConfigValue -Config $Config -Name 'RequiredRestartAfterDays' -CurrentValue $RequiredRestartAfterDays -Type int
        RestartCountdownMinutes = Get-ConfigValue -Config $Config -Name 'RestartCountdownMinutes' -CurrentValue $RestartCountdownMinutes -Type int
        DefaultPostponeMinutes = Get-ConfigValue -Config $Config -Name 'DefaultPostponeMinutes' -CurrentValue $DefaultPostponeMinutes -Type int
        PostponeOptionsMinutes = Get-ConfigValue -Config $Config -Name 'PostponeOptionsMinutes' -CurrentValue $PostponeOptionsMinutes -Type intArray
        MaxPostponeCount = Get-ConfigValue -Config $Config -Name 'MaxPostponeCount' -CurrentValue $MaxPostponeCount -Type int
        ReminderCooldownHours = Get-ConfigValue -Config $Config -Name 'ReminderCooldownHours' -CurrentValue $ReminderCooldownHours -Type int
        WindowIconPath = Get-ConfigValue -Config $Config -Name 'WindowIconPath' -CurrentValue $WindowIconPath -Type string
        WindowTitle = Get-ConfigValue -Config $Config -Name 'WindowTitle' -CurrentValue $WindowTitle -Type string
        CompanyName = Get-ConfigValue -Config $Config -Name 'CompanyName' -CurrentValue $CompanyName -Type string
        DefaultLanguage = Get-ConfigValue -Config $Config -Name 'DefaultLanguage' -CurrentValue $DefaultLanguage -Type string
        ForceLanguage = Get-ConfigValue -Config $Config -Name 'ForceLanguage' -CurrentValue ([bool]$ForceLanguage) -Type bool
        LanguageCatalogPath = Get-ConfigValue -Config $Config -Name 'LanguageCatalogPath' -CurrentValue $LanguageCatalogPath -Type string
        SplashEnabled = Get-ConfigValue -Config $Config -Name 'SplashEnabled' -CurrentValue $SplashEnabled -Type bool
        SplashMinimumDurationMs = Get-ConfigValue -Config $Config -Name 'SplashMinimumDurationMs' -CurrentValue $SplashMinimumDurationMs -Type int
        SplashProductName = Get-ConfigValue -Config $Config -Name 'SplashProductName' -CurrentValue $SplashProductName -Type string
        SplashBadgeText = Get-ConfigValue -Config $Config -Name 'SplashBadgeText' -CurrentValue $SplashBadgeText -Type string
        SplashSubtitle = Get-ConfigValue -Config $Config -Name 'SplashSubtitle' -CurrentValue $SplashSubtitle -Type string
        SplashLogoPath = Get-ConfigValue -Config $Config -Name 'SplashLogoPath' -CurrentValue $SplashLogoPath -Type string
        EnableDebugLogging = Get-ConfigValue -Config $Config -Name 'EnableDebugLogging' -CurrentValue ([bool]$EnableDebugLogging) -Type bool
        PreviewOnly = Get-ConfigValue -Config $Config -Name 'PreviewOnly' -CurrentValue ([bool]$PreviewOnly) -Type bool
        TestRequiredRestart = Get-ConfigValue -Config $Config -Name 'TestRequiredRestart' -CurrentValue ([bool]$TestRequiredRestart) -Type bool
        TestRecommendedRestart = Get-ConfigValue -Config $Config -Name 'TestRecommendedRestart' -CurrentValue ([bool]$TestRecommendedRestart) -Type bool
        NeverForceRestart = Get-ConfigValue -Config $Config -Name 'NeverForceRestart' -CurrentValue ([bool]$NeverForceRestart) -Type bool
    }
}

function Initialize-EffectiveSettings {
    if ($null -eq $script:Effective) {
        $script:Effective = [ordered]@{
            RecommendedRestartAfterDays = $RecommendedRestartAfterDays
            RequiredRestartAfterDays = $RequiredRestartAfterDays
            RestartCountdownMinutes = $RestartCountdownMinutes
            DefaultPostponeMinutes = $DefaultPostponeMinutes
            PostponeOptionsMinutes = $PostponeOptionsMinutes
            MaxPostponeCount = $MaxPostponeCount
            ReminderCooldownHours = $ReminderCooldownHours
            WindowIconPath = $WindowIconPath
            WindowTitle = $WindowTitle
            CompanyName = $CompanyName
            DefaultLanguage = $DefaultLanguage
            ForceLanguage = [bool]$ForceLanguage
            LanguageCatalogPath = $LanguageCatalogPath
            SplashEnabled = $SplashEnabled
            SplashMinimumDurationMs = $SplashMinimumDurationMs
            SplashProductName = $SplashProductName
            SplashBadgeText = $SplashBadgeText
            SplashSubtitle = $SplashSubtitle
            SplashLogoPath = $SplashLogoPath
            EnableDebugLogging = [bool]$EnableDebugLogging
            PreviewOnly = [bool]$PreviewOnly
            TestRequiredRestart = [bool]$TestRequiredRestart
            TestRecommendedRestart = [bool]$TestRecommendedRestart
            NeverForceRestart = [bool]$NeverForceRestart
        }
    }

    if (-not $script:Effective.PostponeOptionsMinutes -or $script:Effective.PostponeOptionsMinutes.Count -eq 0) {
        $script:Effective.PostponeOptionsMinutes = @([int]$script:Effective.DefaultPostponeMinutes)
    }

    $script:Effective.RecommendedRestartAfterDays = [Math]::Max(0, [int]$script:Effective.RecommendedRestartAfterDays)
    $script:Effective.RequiredRestartAfterDays = [Math]::Max($script:Effective.RecommendedRestartAfterDays + 1, [int]$script:Effective.RequiredRestartAfterDays)
    $script:Effective.RestartCountdownMinutes = [Math]::Max(1, [int]$script:Effective.RestartCountdownMinutes)
    $script:Effective.MaxPostponeCount = [Math]::Max(0, [int]$script:Effective.MaxPostponeCount)
    $script:Effective.ReminderCooldownHours = [Math]::Max(1, [int]$script:Effective.ReminderCooldownHours)
    $script:Effective.SplashMinimumDurationMs = [Math]::Max(0, [int]$script:Effective.SplashMinimumDurationMs)
    $script:PreviewOnly = [bool]$script:Effective.PreviewOnly
}

function Read-ReminderState {
    if (-not (Test-Path -LiteralPath $script:StatePath)) { return $null }

    try {
        return Get-Content -LiteralPath $script:StatePath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        Write-AppLog -Level WARN -Message ("Unable to read reminder state. It will be ignored. {0}" -f $_.Exception.Message)
        return $null
    }
}

function Save-ReminderState {
    param(
        [Parameter(Mandatory = $true)][string]$Reason,
        [int]$DaysSinceReboot,
        [datetime]$LastBoot
    )

    $state = [ordered]@{
        LastReminderUtc = (Get-Date).ToUniversalTime().ToString('o')
        LastReason = $Reason
        DaysSinceReboot = $DaysSinceReboot
        LastBootUtc = $LastBoot.ToUniversalTime().ToString('o')
        ComputerName = $env:COMPUTERNAME
        UserName = $env:USERNAME
    }

    Initialize-AppFolder
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($script:StatePath, ($state | ConvertTo-Json -Depth 4), $utf8NoBom)
    Write-AppLog -Message ("Reminder state saved. Reason={0}" -f $Reason)
}

function Read-UserPreferences {
    if (-not (Test-Path -LiteralPath $script:PreferencesPath)) { return $null }

    try {
        return Get-Content -LiteralPath $script:PreferencesPath -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        Write-AppLog -Level WARN -Message ("Unable to read user preferences. They will be ignored. {0}" -f $_.Exception.Message)
        return $null
    }
}

function Save-UserLanguagePreference {
    param([Parameter(Mandatory = $true)][string]$LanguageCode)

    $preferences = [ordered]@{
        DefaultLanguage = $LanguageCode
        UpdatedUtc = (Get-Date).ToUniversalTime().ToString('o')
        UserName = $env:USERNAME
    }

    Initialize-AppFolder
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($script:PreferencesPath, ($preferences | ConvertTo-Json -Depth 4), $utf8NoBom)
    Write-AppLog -Message ("User language preference saved: {0}" -f $LanguageCode)
}

function Apply-UserLanguagePreference {
    if ($script:Effective.ForceLanguage) { return }
    if ($script:CliBoundParameters.ContainsKey('DefaultLanguage')) { return }

    $preferences = Read-UserPreferences
    if ($null -eq $preferences -or [string]::IsNullOrWhiteSpace([string]$preferences.DefaultLanguage)) { return }

    $script:Effective.DefaultLanguage = [string]$preferences.DefaultLanguage
    Write-AppLog -Message ("User language preference loaded: {0}" -f $script:Effective.DefaultLanguage)
}

function Test-ReminderCooldown {
    param([object]$State)

    if ($null -eq $State) { return $false }
    if (-not $State.LastReminderUtc) { return $false }

    try {
        $lastReminder = [datetime]::Parse($State.LastReminderUtc).ToUniversalTime()
        $ageHours = ((Get-Date).ToUniversalTime() - $lastReminder).TotalHours
        return ($ageHours -lt [double]$script:Effective.ReminderCooldownHours)
    }
    catch {
        return $false
    }
}

function Get-LastBootTime {
    try {
        $os = Get-CimInstance Win32_OperatingSystem
        if ($os.LastBootUpTime -is [datetime]) { return $os.LastBootUpTime }
        return [Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime)
    }
    catch {
        Write-AppLog -Level WARN -Message ("Last boot detection failed, falling back to max threshold. {0}" -f $_.Exception.Message)
        return (Get-Date).AddDays(-[int]$script:Effective.RequiredRestartAfterDays)
    }
}

function Get-PrimaryIPv4Address {
    try {
        $address = Get-NetIPAddress -AddressFamily IPv4 |
            Where-Object {
                $_.IPAddress -and
                $_.IPAddress -notlike '169.254.*' -and
                $_.IPAddress -ne '127.0.0.1' -and
                $_.PrefixOrigin -ne 'WellKnown'
            } |
            Sort-Object InterfaceMetric, InterfaceIndex |
            Select-Object -First 1 -ExpandProperty IPAddress

        if (-not [string]::IsNullOrWhiteSpace([string]$address)) { return [string]$address }
    }
    catch {
        Write-AppLog -Level WARN -Message ("Primary IPv4 detection failed. {0}" -f $_.Exception.Message)
    }

    return 'N/A'
}

function Get-DetectedCompanyName {
    if (-not [string]::IsNullOrWhiteSpace([string]$script:Effective.CompanyName)) {
        return [string]$script:Effective.CompanyName
    }

    $registryPaths = @(
        'HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\TenantInfo\*',
        'HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\JoinInfo\*',
        'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin\TenantInfo\*',
        'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\WorkplaceJoin\JoinInfo\*'
    )
    $candidateProperties = @('DisplayName','TenantDisplayName','TenantName','TenantDomain','IdpDomain')

    foreach ($path in $registryPaths) {
        try {
            foreach ($item in @(Get-ItemProperty -Path $path -ErrorAction SilentlyContinue)) {
                foreach ($propertyName in $candidateProperties) {
                    $property = $item.PSObject.Properties[$propertyName]
                    if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) {
                        return [string]$property.Value
                    }
                }
            }
        }
        catch {
            Write-AppLog -Level DEBUG -Message ("Company registry detection failed for {0}. {1}" -f $path, $_.Exception.Message)
        }
    }

    try {
        $tenantLine = & dsregcmd.exe /status 2>$null |
            Where-Object { $_ -match '^\s*(TenantName|WorkplaceTenantName)\s*:' } |
            Select-Object -First 1
        if ($tenantLine -match ':\s*(.+)$') {
            $tenantName = $matches[1].Trim()
            if (-not [string]::IsNullOrWhiteSpace($tenantName)) { return $tenantName }
        }
    }
    catch {
        Write-AppLog -Level DEBUG -Message ("Company dsregcmd detection failed. {0}" -f $_.Exception.Message)
    }

    try {
        $computerSystem = Get-CimInstance Win32_ComputerSystem
        if ($computerSystem.PartOfDomain -and -not [string]::IsNullOrWhiteSpace([string]$computerSystem.Domain)) {
            return [string]$computerSystem.Domain
        }
    }
    catch {
        Write-AppLog -Level DEBUG -Message ("Company domain detection failed. {0}" -f $_.Exception.Message)
    }

    return ''
}

function Test-UserSessionVisible {
    if ($script:Effective.TestRequiredRestart -or $script:Effective.TestRecommendedRestart) { return $true }

    try {
        $sessionLines = & query session $env:USERNAME 2>$null
        $sessionMatch = $sessionLines | Where-Object { $_ -match [regex]::Escape($env:USERNAME) } | Select-Object -First 1
        if ($sessionMatch -match 'Disc') {
            Write-AppLog -Level WARN -Message ("Session for {0} appears disconnected. Exiting without showing UI." -f $env:USERNAME)
            return $false
        }
    }
    catch {
        Write-AppLog -Level WARN -Message ("Session visibility check failed, continuing. {0}" -f $_.Exception.Message)
    }

    return $true
}

function Invoke-RelaunchSta {
    $exe = (Get-Process -Id $PID).Path
    $arguments = New-Object System.Collections.Generic.List[string]
    $arguments.Add('-NoProfile')
    $arguments.Add('-ExecutionPolicy')
    $arguments.Add('Bypass')
    $arguments.Add('-Sta')
    $arguments.Add('-File')
    $arguments.Add(('"{0}"' -f $PSCommandPath))

    foreach ($key in @('RecommendedRestartAfterDays','RequiredRestartAfterDays','RestartCountdownMinutes','DefaultPostponeMinutes','MaxPostponeCount','ReminderCooldownHours','DefaultLanguage','WindowIconPath','WindowTitle','CompanyName','LanguageCatalogPath')) {
        $value = $script:Effective[$key]
        if ($null -ne $value -and "$value" -ne '') {
            $arguments.Add("-$key")
            $arguments.Add(('"{0}"' -f $value))
        }
    }

    if ($script:Effective.PostponeOptionsMinutes -and $script:Effective.PostponeOptionsMinutes.Count -gt 0) {
        $arguments.Add('-PostponeOptionsMinutes')
        foreach ($value in $script:Effective.PostponeOptionsMinutes) {
            $arguments.Add([string][int]$value)
        }
    }
    if ($script:Effective.EnableDebugLogging) { $arguments.Add('-EnableDebugLogging') }
    if ($script:Effective.PreviewOnly) { $arguments.Add('-PreviewOnly') }
    if ($script:Effective.TestRequiredRestart) { $arguments.Add('-TestRequiredRestart') }
    if ($script:Effective.TestRecommendedRestart) { $arguments.Add('-TestRecommendedRestart') }
    if ($script:Effective.NeverForceRestart) { $arguments.Add('-NeverForceRestart') }
    if ($script:Effective.ForceLanguage) { $arguments.Add('-ForceLanguage') }
    if ($ConfigPath) {
        $arguments.Add('-ConfigPath')
        $arguments.Add(('"{0}"' -f $ConfigPath))
    }

    Start-Process -FilePath $exe -ArgumentList ($arguments -join ' ') -WindowStyle Hidden | Out-Null
    [Environment]::Exit(0)
}

function Invoke-SafeRestart {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param([string]$Reason = 'User action')

    if ($script:PreviewOnly) {
        Write-AppLog -Level WARN -Message ("PreviewOnly is enabled. Restart skipped. Reason={0}" -f $Reason)
        [System.Windows.MessageBox]::Show('Preview mode is enabled. No restart was triggered.', $script:AppName, 'OK', 'Information') | Out-Null
        return
    }

    if (-not $PSCmdlet.ShouldProcess($env:COMPUTERNAME, 'Restart computer')) {
        Write-AppLog -Level WARN -Message 'Restart skipped by ShouldProcess.'
        return
    }

    Write-AppLog -Level WARN -Message ("Restart requested. Reason={0}" -f $Reason)
    try {
        $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        if ($isAdmin) {
            Restart-Computer -Force
        }
        else {
            Start-Process -FilePath 'shutdown.exe' -ArgumentList '/r /t 0' -WindowStyle Hidden
        }
    }
    catch {
        Stop-AppWithError -Message ("Restart failed: {0}" -f $_.Exception.Message) -ExitCode 20
    }
}

function Import-AppStringsCatalog {
    $defaultPath = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-DeviceRebootManager-GUI.strings.psd1'
    $path = if (-not [string]::IsNullOrWhiteSpace($script:Effective.LanguageCatalogPath)) { Resolve-AppRelativePath -Path $script:Effective.LanguageCatalogPath } else { $defaultPath }

    if (-not (Test-Path -LiteralPath $path)) {
        throw "Strings resource file not found: $path"
    }

    $tokens = $null
    $parseErrors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$parseErrors)
    if ($parseErrors -and $parseErrors.Count -gt 0) {
        throw ("Strings resource file could not be parsed: {0}" -f (($parseErrors | ForEach-Object { $_.Message }) -join '; '))
    }

    $statement = $ast.EndBlock.Statements | Select-Object -First 1
    if ($null -eq $statement -or $null -eq $statement.PipelineElements -or $statement.PipelineElements.Count -eq 0) {
        throw "Strings resource file is empty: $path"
    }

    $catalog = $statement.PipelineElements[0].Expression.SafeGetValue()
    if (-not $catalog.ContainsKey('Strings') -or -not $catalog.Strings.ContainsKey('en')) {
        throw "Strings resource file is invalid or missing the English fallback: $path"
    }

    return $catalog
}

function Resolve-AppLanguage {
    param([Parameter(Mandatory = $true)][hashtable]$Catalog)

    $candidates = New-Object System.Collections.Generic.List[string]
    if ($script:Effective.DefaultLanguage -eq 'auto') {
        try {
            $culture = [Globalization.CultureInfo]::CurrentUICulture
            $candidates.Add($culture.Name)
            $candidates.Add($culture.TwoLetterISOLanguageName)
        }
        catch {
            $candidates.Add('en')
        }
    }
    else {
        $candidates.Add([string]$script:Effective.DefaultLanguage)
    }

    $aliases = if ($Catalog.ContainsKey('Aliases')) { $Catalog.Aliases } else { @{} }
    foreach ($candidate in $candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        $normalized = $candidate.Trim()
        if ($aliases.ContainsKey($normalized)) { $normalized = [string]$aliases[$normalized] }
        elseif ($aliases.ContainsKey($normalized.ToLowerInvariant())) { $normalized = [string]$aliases[$normalized.ToLowerInvariant()] }

        if ($Catalog.Strings.ContainsKey($normalized)) { return $normalized }

        $baseName = ($normalized -split '-')[0]
        if ($aliases.ContainsKey($baseName)) { $baseName = [string]$aliases[$baseName] }
        if ($Catalog.Strings.ContainsKey($baseName)) { return $baseName }
    }

    if ($Catalog.ContainsKey('DefaultLanguage')) { return [string]$Catalog.DefaultLanguage }
    return 'en'
}

function Get-AppStrings {
    $catalog = Import-AppStringsCatalog
    $script:StringsCatalog = $catalog
    $selected = Resolve-AppLanguage -Catalog $catalog
    $script:SelectedLanguage = $selected

    $english = @{}
    foreach ($key in $catalog.Strings.en.Keys) {
        $english[$key] = $catalog.Strings.en[$key]
    }

    $localized = $catalog.Strings[$selected]
    foreach ($key in $localized.Keys) {
        $english[$key] = $localized[$key]
    }

    if (-not $english.ContainsKey('FlowDirection')) {
        $english.FlowDirection = 'LeftToRight'
    }
    if (-not $english.ContainsKey('Close')) {
        $english.Close = 'Close'
    }
    if (-not $english.ContainsKey('DaysUnit')) {
        $english.DaysUnit = 'days'
    }
    if (-not $english.ContainsKey('ShowingGui')) {
        $english.ShowingGui = 'Showing Smart Device Reboot Manager GUI.'
    }
    if (-not $english.ContainsKey('ClosedGui')) {
        $english.ClosedGui = 'Smart Device Reboot Manager GUI closed.'
    }
    if (-not $english.ContainsKey('LanguageLabel')) {
        $english.LanguageLabel = 'Language'
    }
    if (-not $english.ContainsKey('LanguageAuto')) {
        $english.LanguageAuto = 'Automatic'
    }

    Write-AppLog -Message ("Language selected: {0}" -f $selected)
    return [pscustomobject]$english
}

function Add-WpfAssemblies {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName PresentationCore
    Add-Type -AssemblyName WindowsBase
}

function Import-SmartM365GuiSplash {
    $current = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
        $PSScriptRoot
    }
    else {
        Split-Path -Parent $MyInvocation.MyCommand.Path
    }

    while ($current) {
        $splashPath = Join-Path -Path $current -ChildPath 'SmartM365.GuiSplash.ps1'
        if (Test-Path -LiteralPath $splashPath) {
            return $splashPath
        }

        if ((Split-Path -Path $current -Leaf) -eq 'SmartM365') {
            return $null
        }

        $parent = Split-Path -Path $current -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) { break }
        $current = $parent
    }

    return $null
}

function Set-UiValue {
    param([string]$Name, [object]$Value)
    if ($script:Ui.ContainsKey($Name) -and $null -ne $script:Ui[$Name]) {
        $script:Ui[$Name].Text = [string]$Value
    }
}

function Update-CountdownUi {
    $remaining = [Math]::Max(0, [int]$script:SecondsRemaining)
    $minutes = [Convert]::ToInt32([Math]::Floor($remaining / 60))
    $seconds = $remaining % 60
    Set-UiValue -Name 'CountdownText' -Value ([string]::Format($script:Strings.Countdown, $minutes, $seconds))

    if ($script:Ui.ContainsKey('CountdownProgress') -and $script:Ui.CountdownProgress) {
        $elapsed = [Math]::Max(0, $script:InitialSeconds - $remaining)
        $script:Ui.CountdownProgress.Value = [Math]::Min(100, [Math]::Round(($elapsed / [double]$script:InitialSeconds) * 100, 0))
    }
}

function Set-PostponeState {
    if (-not $script:ForceMode) {
        $script:Ui.PostponeButton.IsEnabled = $false
        $script:Ui.PostponeCombo.IsEnabled = $false
        $script:Ui.PostponeButton.Content = $script:Strings.Postpone
        return
    }

    $remainingPostpones = [Math]::Max(0, [int]$script:Effective.MaxPostponeCount - [int]$script:PostponeCount)
    $script:Ui.PostponeButton.IsEnabled = ($remainingPostpones -gt 0)
    $script:Ui.PostponeCombo.IsEnabled = ($remainingPostpones -gt 0)
    $script:Ui.PostponeButton.Content = ('{0} ({1})' -f $script:Strings.Postpone, $remainingPostpones)
}

function Get-EffectiveWindowTitle {
    if ($script:Effective.Contains('WindowTitle') -and -not [string]::IsNullOrWhiteSpace([string]$script:Effective.WindowTitle)) {
        return [string]$script:Effective.WindowTitle
    }

    $baseTitle = [string]$script:Strings.WindowTitle
    if (-not [string]::IsNullOrWhiteSpace([string]$script:DetectedCompanyName)) {
        return ('{0} - {1}' -f $script:DetectedCompanyName, $baseTitle)
    }

    return $baseTitle
}

function Get-LanguageChoices {
    $fallbackNames = @{
        en = 'English'
        fr = 'French'
        de = 'German'
        es = 'Spanish'
        nl = 'Dutch'
        it = 'Italian'
        pt = 'Portuguese'
        pl = 'Polish'
        ar = 'Arabic'
        tr = 'Turkish'
        sv = 'Swedish'
        da = 'Danish'
        nb = 'Norwegian'
        fi = 'Finnish'
        ro = 'Romanian'
        hu = 'Hungarian'
        ja = 'Japanese'
        ko = 'Korean'
        'zh-Hans' = 'Chinese'
        uk = 'Ukrainian'
    }
    $catalogNames = if ($script:StringsCatalog -and $script:StringsCatalog.ContainsKey('LanguageNames')) { $script:StringsCatalog.LanguageNames } else { @{} }

    $choices = New-Object System.Collections.Generic.List[object]
    $choices.Add([pscustomobject]@{ Code = 'auto'; Display = $script:Strings.LanguageAuto })
    foreach ($code in @('en','fr','de','es','nl','it','pt','pl','ar','tr','sv','da','nb','fi','ro','hu','ja','ko','zh-Hans','uk')) {
        $display = if ($catalogNames.ContainsKey($code)) { [string]$catalogNames[$code] } else { [string]$fallbackNames[$code] }
        $choices.Add([pscustomobject]@{ Code = $code; Display = $display })
    }

    return $choices
}

function Update-LanguageSelector {
    if (-not $script:Ui.ContainsKey('LanguageCombo') -or $null -eq $script:Ui.LanguageCombo) { return }

    $script:IsUpdatingLanguageSelector = $true
    try {
        $selectedCode = [string]$script:Effective.DefaultLanguage
        if ([string]::IsNullOrWhiteSpace($selectedCode)) { $selectedCode = 'auto' }

        $script:Ui.LanguageCombo.Items.Clear()
        $index = 0
        $selectedIndex = 0
        foreach ($choice in Get-LanguageChoices) {
            [void]$script:Ui.LanguageCombo.Items.Add($choice)
            if ($choice.Code -eq $selectedCode) { $selectedIndex = $index }
            $index++
        }

        $script:Ui.LanguageCombo.DisplayMemberPath = 'Display'
        $script:Ui.LanguageCombo.SelectedIndex = $selectedIndex
        $script:Ui.LanguageCombo.IsEnabled = (-not [bool]$script:Effective.ForceLanguage)
    }
    finally {
        $script:IsUpdatingLanguageSelector = $false
    }
}

function Apply-LocalizedUi {
    if ($script:Window) {
        $script:Window.Title = Get-EffectiveWindowTitle
        if ($script:Strings.FlowDirection -eq 'RightToLeft') {
            $script:Window.FlowDirection = [System.Windows.FlowDirection]::RightToLeft
        }
        else {
            $script:Window.FlowDirection = [System.Windows.FlowDirection]::LeftToRight
        }
    }

    if ($script:ForceMode) {
        $script:Ui.StatusPillText.Text = $script:Strings.PillForced
        $script:Ui.StatusPill.Background = '#FFF1F0'
        $script:Ui.StatusPillText.Foreground = '#B42318'
        $script:Ui.TitleText.Text = $script:Strings.ForcedTitle
        $script:Ui.SummaryText.Text = $script:Strings.ForcedSummary
        $script:Ui.NextActionText.Text = $script:Strings.NextActionForced
    }
    else {
        $script:Ui.StatusPillText.Text = $script:Strings.PillRecommended
        $script:Ui.StatusPill.Background = '#E6F4FF'
        $script:Ui.StatusPillText.Foreground = '#005A9E'
        $script:Ui.TitleText.Text = $script:Strings.RecommendedTitle
        $script:Ui.SummaryText.Text = $script:Strings.RecommendedSummary
        $script:Ui.NextActionText.Text = $script:Strings.NextActionRecommended
        $script:Ui.CountdownText.Text = $script:Strings.NoCountdown
        $script:Ui.CountdownProgress.Value = 0
    }

    if ($script:PreviewOnly) {
        $script:Ui.SummaryText.Text = ('{0} ({1})' -f $script:Ui.SummaryText.Text, $script:Strings.Preview)
    }

    $script:Ui.DaysBadgeText.Text = [string]$script:DaysSinceReboot
    $script:Ui.DaysUnitText.Text = $script:Strings.DaysUnit
    $script:Ui.UptimeLabel.Text = $script:Strings.Uptime
    $script:Ui.UptimeValue.Text = ('{0} d' -f $script:DaysSinceReboot)
    $script:Ui.ThresholdLabel.Text = $script:Strings.Threshold
    $script:Ui.ThresholdValue.Text = ('{0} d' -f $script:Effective.RecommendedRestartAfterDays)
    $script:Ui.MaximumLabel.Text = $script:Strings.Maximum
    $script:Ui.MaximumValue.Text = ('{0} d' -f $script:Effective.RequiredRestartAfterDays)
    $script:Ui.LastBootLabel.Text = $script:Strings.LastBoot
    $script:Ui.LastBootValue.Text = $script:LastBoot.ToString('yyyy-MM-dd HH:mm')
    $script:Ui.ActivityTitleText.Text = $script:Strings.Activity
    $script:Ui.LanguageLabelText.Text = $script:Strings.LanguageLabel
    $script:Ui.DeviceIdentityText.Text = ('{0} ({1})' -f $env:COMPUTERNAME, $script:PrimaryIPv4Address)
    $script:Ui.RestartButton.Content = $script:Strings.RestartNow
    $script:Ui.LaterButton.Content = $script:Strings.RemindLater
    $script:Ui.CloseButton.Content = $script:Strings.Close

    Set-PostponeState
    Update-LanguageSelector
    if ($script:ForceMode) { Update-CountdownUi }
}

function New-RebootManagerWindow {
    Add-WpfAssemblies

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Smart Device Reboot Manager"
        Width="800"
        Height="600"
        MinWidth="760"
        MinHeight="560"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize"
        Background="#F4F7FB"
        FontFamily="Segoe UI">
    <Window.Resources>
        <SolidColorBrush x:Key="InkBrush" Color="#102033"/>
        <SolidColorBrush x:Key="MutedBrush" Color="#5D6B7C"/>
        <SolidColorBrush x:Key="PanelBrush" Color="#FFFFFF"/>
        <SolidColorBrush x:Key="BorderBrushSoft" Color="#D9E2EC"/>
        <SolidColorBrush x:Key="AccentBrush" Color="#0078D4"/>
        <SolidColorBrush x:Key="DangerBrush" Color="#B42318"/>
        <Style TargetType="Button">
            <Setter Property="MinWidth" Value="118"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="Margin" Value="0,0,8,0"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
            <Setter Property="BorderBrush" Value="#B9C8D7"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="ButtonBorder"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                CornerRadius="6">
                            <ContentPresenter HorizontalAlignment="Center"
                                              VerticalAlignment="Center"
                                              Margin="{TemplateBinding Padding}"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Background" Value="#EAF6FF"/>
                                <Setter TargetName="ButtonBorder" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter TargetName="ButtonBorder" Property="Background" Value="#D8ECFA"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.45"/>
                                <Setter Property="Cursor" Value="Arrow"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="PrimaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="{StaticResource AccentBrush}"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
        </Style>
        <Style x:Key="DangerButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
            <Setter Property="Background" Value="#FFF1F0"/>
            <Setter Property="Foreground" Value="{StaticResource DangerBrush}"/>
            <Setter Property="BorderBrush" Value="#FDA29B"/>
        </Style>
        <Style TargetType="ProgressBar">
            <Setter Property="Height" Value="10"/>
            <Setter Property="Foreground" Value="{StaticResource AccentBrush}"/>
            <Setter Property="Background" Value="#E5EAF0"/>
        </Style>
        <Style TargetType="ComboBox">
            <Setter Property="Height" Value="34"/>
            <Setter Property="Padding" Value="8,4"/>
            <Setter Property="Margin" Value="0,0,8,0"/>
            <Setter Property="Background" Value="#FFFFFF"/>
            <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
            <Setter Property="BorderBrush" Value="#B9C8D7"/>
            <Setter Property="BorderThickness" Value="1"/>
        </Style>
        <Style x:Key="MetricLabel" TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource MutedBrush}"/>
            <Setter Property="FontSize" Value="12"/>
        </Style>
        <Style x:Key="MetricValue" TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
            <Setter Property="FontSize" Value="20"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Margin" Value="0,3,0,0"/>
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
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <StackPanel>
                    <Border x:Name="StatusPill" HorizontalAlignment="Left" CornerRadius="12" Padding="10,4" Background="#E6F4FF">
                        <TextBlock x:Name="StatusPillText" Foreground="#005A9E" FontWeight="SemiBold" FontSize="12"/>
                    </Border>
                    <TextBlock x:Name="TitleText" Margin="0,12,0,0" FontSize="28" FontWeight="SemiBold" Foreground="{StaticResource InkBrush}" TextWrapping="Wrap"/>
                    <TextBlock x:Name="SummaryText" Margin="0,8,0,0" FontSize="14" Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap"/>
                </StackPanel>
                <Border Grid.Column="1" Width="116" Height="116" CornerRadius="58" Background="#EAF6FF" Margin="18,0,0,0">
                    <StackPanel VerticalAlignment="Center">
                        <TextBlock x:Name="DaysBadgeText" HorizontalAlignment="Center" FontSize="36" FontWeight="Bold" Foreground="{StaticResource AccentBrush}"/>
                        <TextBlock x:Name="DaysUnitText" HorizontalAlignment="Center" Foreground="{StaticResource MutedBrush}" FontSize="12"/>
                    </StackPanel>
                </Border>
            </Grid>
        </Border>

        <Border Grid.Row="1" Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,12,0,0">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="1.3*"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0">
                    <TextBlock x:Name="UptimeLabel" Style="{StaticResource MetricLabel}"/>
                    <TextBlock x:Name="UptimeValue" Style="{StaticResource MetricValue}"/>
                </StackPanel>
                <StackPanel Grid.Column="1">
                    <TextBlock x:Name="ThresholdLabel" Style="{StaticResource MetricLabel}"/>
                    <TextBlock x:Name="ThresholdValue" Style="{StaticResource MetricValue}"/>
                </StackPanel>
                <StackPanel Grid.Column="2">
                    <TextBlock x:Name="MaximumLabel" Style="{StaticResource MetricLabel}"/>
                    <TextBlock x:Name="MaximumValue" Style="{StaticResource MetricValue}"/>
                </StackPanel>
                <StackPanel Grid.Column="3">
                    <TextBlock x:Name="LastBootLabel" Style="{StaticResource MetricLabel}"/>
                    <TextBlock x:Name="LastBootValue" Style="{StaticResource MetricValue}" FontSize="15"/>
                </StackPanel>
            </Grid>
        </Border>

        <Grid Grid.Row="2" Margin="0,12,0,12">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="1.1*"/>
                <ColumnDefinition Width="0.9*"/>
            </Grid.ColumnDefinitions>
            <Border Grid.Column="0" Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,12,0">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <TextBlock x:Name="CountdownText" Grid.Row="0" FontSize="24" FontWeight="SemiBold" Foreground="{StaticResource InkBrush}" TextWrapping="Wrap"/>
                    <ProgressBar x:Name="CountdownProgress" Grid.Row="1" Margin="0,14,0,12" Minimum="0" Maximum="100"/>
                    <TextBlock x:Name="NextActionText" Grid.Row="2" Foreground="{StaticResource MutedBrush}" FontSize="14" TextWrapping="Wrap"/>
                    <Separator Grid.Row="3" Margin="0,14,0,10"/>
                    <TextBlock x:Name="DeviceIdentityText" Grid.Row="4" Foreground="{StaticResource MutedBrush}" FontSize="13" TextWrapping="Wrap"/>
                </Grid>
            </Border>
            <Border Grid.Column="1" Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="16">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                    </Grid.RowDefinitions>
                    <TextBlock x:Name="ActivityTitleText" Grid.Row="0" FontWeight="SemiBold" Foreground="{StaticResource InkBrush}" Margin="0,0,0,6"/>
                    <Grid Grid.Row="1" Margin="0,0,0,8">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock x:Name="LanguageLabelText" Grid.Column="0" Foreground="{StaticResource MutedBrush}" FontSize="12" VerticalAlignment="Center" Margin="0,0,8,0"/>
                        <ComboBox x:Name="LanguageCombo" Grid.Column="1" MinWidth="150" HorizontalAlignment="Stretch" Height="28" Margin="0"/>
                    </Grid>
                    <TextBox x:Name="ActivityText" Grid.Row="2" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" BorderBrush="#E4EAF1" Background="#FAFCFE"/>
                </Grid>
            </Border>
        </Grid>

        <Border Grid.Row="3" Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource BorderBrushSoft}" BorderThickness="1" CornerRadius="8" Padding="12">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="120"/>
                    <ColumnDefinition Width="*"/>
                    <ColumnDefinition Width="Auto"/>
                </Grid.ColumnDefinitions>
                <Button x:Name="RestartButton" Grid.Column="0" Width="158" Style="{StaticResource DangerButton}"/>
                <ComboBox x:Name="PostponeCombo" Grid.Column="1" Width="112"/>
                <Button x:Name="PostponeButton" Grid.Column="2" Width="126" HorizontalAlignment="Left"/>
                <StackPanel Grid.Column="3" Orientation="Horizontal" HorizontalAlignment="Right">
                    <Button x:Name="LaterButton" Width="160"/>
                    <Button x:Name="CloseButton" Width="118" Margin="0"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)

    foreach ($name in @(
        'StatusPill','StatusPillText','TitleText','SummaryText','DaysBadgeText','DaysUnitText',
        'UptimeLabel','UptimeValue','ThresholdLabel','ThresholdValue','MaximumLabel','MaximumValue',
        'LastBootLabel','LastBootValue','CountdownText','CountdownProgress','NextActionText',
        'DeviceIdentityText','ActivityTitleText','ActivityText','LanguageLabelText',
        'LanguageCombo','RestartButton','PostponeCombo','PostponeButton','LaterButton','CloseButton'
    )) {
        $script:Ui[$name] = $window.FindName($name)
    }

    return $window
}

Initialize-AppFolder
Invoke-LaunchLogRotation
$config = Read-AppConfig
Apply-AppConfig -Config $config
Initialize-EffectiveSettings
$script:DetectedCompanyName = Get-DetectedCompanyName
Apply-UserLanguagePreference

if ($ValidateOnly) {
    if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA) {
        throw "Validation requires an STA runspace. Start with: powershell.exe -STA -NoProfile -File `"$PSCommandPath`" -ValidateOnly"
    }

    $script:Strings = Get-AppStrings
    $script:ForceMode = $false
    $script:DaysSinceReboot = [int]$script:Effective.RecommendedRestartAfterDays
    $script:LastBoot = (Get-Date).AddDays(-[int]$script:DaysSinceReboot)
    $script:PrimaryIPv4Address = Get-PrimaryIPv4Address
    $script:InitialSeconds = [Math]::Max(60, [int]$script:Effective.RestartCountdownMinutes * 60)
    $script:SecondsRemaining = $script:InitialSeconds
    $script:Window = New-RebootManagerWindow
    Apply-LocalizedUi
    Write-AppLog -Message 'Smart Device Reboot Manager GUI validation completed.'
    Write-Output 'Smart Device Reboot Manager GUI validation completed.'
    [Environment]::Exit(0)
}

if ($ResetRunState -and (Test-Path -LiteralPath $script:StatePath)) {
    Remove-Item -LiteralPath $script:StatePath -Force
    Write-AppLog -Message 'Reminder state reset by CLI.'
}

Write-AppLog -Message ("Starting. RecommendedRestartAfterDays={0}; RequiredRestartAfterDays={1}; RestartCountdownMinutes={2}; PreviewOnly={3}; NeverForceRestart={4}" -f $script:Effective.RecommendedRestartAfterDays, $script:Effective.RequiredRestartAfterDays, $script:Effective.RestartCountdownMinutes, $script:PreviewOnly, $script:Effective.NeverForceRestart)

$lastBoot = Get-LastBootTime
$daysSinceReboot = (New-TimeSpan -Start $lastBoot -End (Get-Date)).Days
if ($script:Effective.TestRequiredRestart) {
    $daysSinceReboot = [int]$script:Effective.RequiredRestartAfterDays + 1
    $lastBoot = (Get-Date).AddDays(-$daysSinceReboot)
}
elseif ($script:Effective.TestRecommendedRestart) {
    $daysSinceReboot = [Math]::Min([int]$script:Effective.RequiredRestartAfterDays - 1, [int]$script:Effective.RecommendedRestartAfterDays + 1)
    $lastBoot = (Get-Date).AddDays(-$daysSinceReboot)
}

$script:ForceMode = (-not [bool]$script:Effective.NeverForceRestart -and $daysSinceReboot -ge [int]$script:Effective.RequiredRestartAfterDays)
Write-AppLog -Message ("LastBoot={0}; DaysSinceReboot={1}; ForceMode={2}" -f $lastBoot.ToString('s'), $daysSinceReboot, $script:ForceMode)

if ($daysSinceReboot -lt [int]$script:Effective.RecommendedRestartAfterDays) {
    Write-AppLog -Message 'Device is below the reminder threshold. Exiting.'
    [Environment]::Exit(0)
}

if (-not $script:ForceMode -and -not ($script:Effective.TestRecommendedRestart -or $script:Effective.TestRequiredRestart)) {
    $state = Read-ReminderState
    if (Test-ReminderCooldown -State $state) {
        Write-AppLog -Message ("Reminder cooldown is still active ({0} hours). Exiting." -f $script:Effective.ReminderCooldownHours)
        [Environment]::Exit(0)
    }
}

if (-not (Test-UserSessionVisible)) { [Environment]::Exit(0) }

if ([Threading.Thread]::CurrentThread.GetApartmentState() -ne [Threading.ApartmentState]::STA) {
    Invoke-RelaunchSta
}

$script:GuiSplash = $null
$splashHelperPath = if ([bool]$script:Effective.SplashEnabled) { Import-SmartM365GuiSplash } else { $null }
if ($splashHelperPath) {
    . $splashHelperPath
    $splashProductName = if ([string]::IsNullOrWhiteSpace([string]$script:Effective.SplashProductName)) { 'Device Reboot Manager' } else { [string]$script:Effective.SplashProductName }
    $script:GuiSplash = Start-SmartM365GuiSplash `
        -Framework Wpf `
        -ProductName $splashProductName `
        -MinimumDurationMs ([int]$script:Effective.SplashMinimumDurationMs) `
        -BadgeText ([string]$script:Effective.SplashBadgeText) `
        -Subtitle ([string]$script:Effective.SplashSubtitle) `
        -LogoPath ([string]$script:Effective.SplashLogoPath) `
        -WindowIconPath ([string]$script:Effective.WindowIconPath)
}

$script:DaysSinceReboot = $daysSinceReboot
$script:LastBoot = $lastBoot
$script:PrimaryIPv4Address = Get-PrimaryIPv4Address
$script:InitialSeconds = [Math]::Max(60, [int]$script:Effective.RestartCountdownMinutes * 60)
$script:SecondsRemaining = $script:InitialSeconds

$script:Strings = Get-AppStrings
$script:Window = New-RebootManagerWindow
$script:Window.Topmost = [bool]$script:StartTopmost

$resolvedIconPath = Resolve-AppRelativePath -Path $script:Effective.WindowIconPath
if (-not [string]::IsNullOrWhiteSpace($resolvedIconPath) -and (Test-Path -LiteralPath $resolvedIconPath)) {
    try { $script:Window.Icon = New-Object System.Windows.Media.Imaging.BitmapImage([Uri]$resolvedIconPath) }
    catch { Write-AppLog -Level WARN -Message ("Unable to load icon: {0}" -f $_.Exception.Message) }
}

if ($script:ForceMode) {
    $script:AllowClose = [bool]$script:PreviewOnly
}
else {
    $script:AllowClose = $true
}

foreach ($minutes in $script:Effective.PostponeOptionsMinutes) {
    [void]$script:Ui.PostponeCombo.Items.Add(('{0} min' -f [int]$minutes))
}
$script:Ui.PostponeCombo.SelectedIndex = 0

if (-not $script:ForceMode) {
    $script:Ui.LaterButton.IsEnabled = $true
}
else {
    $script:Ui.LaterButton.IsEnabled = $false
}
$script:Ui.CloseButton.IsEnabled = (-not $script:ForceMode -or $script:PreviewOnly)

Apply-LocalizedUi

$script:Ui.LanguageCombo.Add_SelectionChanged({
    Invoke-Safely {
        if ($script:IsUpdatingLanguageSelector) { return }
        if ($script:Effective.ForceLanguage) { return }
        $selected = $script:Ui.LanguageCombo.SelectedItem
        if ($null -eq $selected -or [string]::IsNullOrWhiteSpace([string]$selected.Code)) { return }

        $script:Effective.DefaultLanguage = [string]$selected.Code
        Save-UserLanguagePreference -LanguageCode $script:Effective.DefaultLanguage
        $script:Strings = Get-AppStrings
        Apply-LocalizedUi
    }
})

$script:Window.Add_Closing({
    if (-not $script:AllowClose) {
        $_.Cancel = $true
        Write-AppLog -Level WARN -Message 'Window close blocked while restart is required.'
    }
})

$script:Ui.RestartButton.Add_Click({
    Invoke-Safely {
        $choice = [System.Windows.MessageBox]::Show($script:Strings.ConfirmRestart, $script:AppName, 'OKCancel', 'Warning')
        if ($choice -ne 'OK') {
            Write-AppLog -Message 'Restart confirmation cancelled by user.'
            return
        }
        if ($script:Timer) { $script:Timer.Stop() }
        Save-ReminderState -Reason 'RestartNow' -DaysSinceReboot $daysSinceReboot -LastBoot $lastBoot
        $script:AllowClose = $true
        $script:Window.Close()
        Invoke-SafeRestart -Reason 'RestartNow'
    }
})

$script:Ui.PostponeButton.Add_Click({
    Invoke-Safely {
        if (-not $script:ForceMode) { return }
        if ($script:PostponeCount -ge [int]$script:Effective.MaxPostponeCount) { return }
        $selectedText = [string]$script:Ui.PostponeCombo.SelectedItem
        $minutes = [int](([regex]::Match($selectedText, '\d+')).Value)
        $script:SecondsRemaining += ($minutes * 60)
        $script:InitialSeconds += ($minutes * 60)
        $script:PostponeCount++
        Save-ReminderState -Reason ("Postpone{0}Minutes" -f $minutes) -DaysSinceReboot $daysSinceReboot -LastBoot $lastBoot
        Write-AppLog -Message ("Restart postponed by {0} minutes. Count={1}/{2}" -f $minutes, $script:PostponeCount, $script:Effective.MaxPostponeCount)
        Set-PostponeState
        Update-CountdownUi
        if ($script:Window.Topmost) { $script:Window.Topmost = $false }
    }
})

$script:Ui.LaterButton.Add_Click({
    Invoke-Safely {
        if ($script:ForceMode) { return }
        Save-ReminderState -Reason 'RemindLater' -DaysSinceReboot $daysSinceReboot -LastBoot $lastBoot
        Write-AppLog -Message 'User selected remind me later.'
        $script:AllowClose = $true
        $script:Window.Close()
    }
})

$script:Ui.CloseButton.Add_Click({
    Invoke-Safely {
        if ($script:ForceMode -and -not $script:PreviewOnly) {
            Write-AppLog -Level WARN -Message 'Close button ignored while restart is required.'
            return
        }
        Save-ReminderState -Reason 'Close' -DaysSinceReboot $daysSinceReboot -LastBoot $lastBoot
        $script:AllowClose = $true
        $script:Window.Close()
    }
})

if ($script:ForceMode) {
    $script:Timer = New-Object System.Windows.Threading.DispatcherTimer
    $script:Timer.Interval = [TimeSpan]::FromSeconds(1)
    $script:Timer.Add_Tick({
        Invoke-Safely {
            $script:SecondsRemaining--
            if ($script:SecondsRemaining -le 0) {
                $script:Timer.Stop()
                Save-ReminderState -Reason 'CountdownExpired' -DaysSinceReboot $daysSinceReboot -LastBoot $lastBoot
                Write-AppLog -Level WARN -Message 'Countdown expired. Restart will be triggered.'
                $script:AllowClose = $true
                $script:Window.Close()
                Invoke-SafeRestart -Reason 'CountdownExpired'
                return
            }

            if ($script:RestoreTopmostAfterPostpone) {
                $thresholdSeconds = [int]$script:RestoreTopmostAfterMinutes * 60
                if ($script:SecondsRemaining -le $thresholdSeconds -and -not $script:Window.Topmost) {
                    $script:Window.Topmost = $true
                    Write-AppLog -Message ("Topmost restored under {0} minutes." -f $script:RestoreTopmostAfterMinutes)
                }
            }

            Update-CountdownUi
        }
    })
    $script:Timer.Start()
}

$script:Window.Add_ContentRendered({
    if ($script:GuiSplash) {
        Hide-SmartM365GuiSplash -Splash $script:GuiSplash
    }
})

Write-AppLog -Message $script:Strings.ShowingGui
[void]$script:Window.ShowDialog()
if ($script:GuiSplash) {
    Close-SmartM365GuiSplash -Splash $script:GuiSplash
}
Write-AppLog -Message $script:Strings.ClosedGui
[Environment]::Exit(0)
