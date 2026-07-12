<# 
.SYNOPSIS
    Smart Endpoint Diagnostics Analyzer.

.DESCRIPTION
    PowerShell/WPF analyzer for Microsoft Intune device diagnostics ZIP files
    and local endpoint diagnostic captures.

    This is a PowerShell rewrite of the SmartLogAnalyzer logic. It does not
    depend on Python, PyInstaller, the source repository Git history, or
    SignPath release workflows.
#>

[CmdletBinding()]
param(
    [string]$ZipPath,
    [switch]$Cli,
    [switch]$CollectLocal,
    [string]$ExportHtmlPath,
    [string]$ExportAnonymizedZipPath,
    [switch]$RunAI,
    [ValidateSet('claude','openai','ollama')]
    [string]$AIProvider,
    [string]$AIModel,
    [string]$AIApiKey,
    [string]$AIOllamaUrl = 'http://localhost:11434',
    [int]$AIMaxTokens = 2048,
    [double]$AITemperature = 0.3,
    [switch]$SkipCabExtraction,
    [switch]$SkipEventLogScan,
    [switch]$ValidateOnly
)

function Test-SedaIsAdministrator {
    try {
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

function ConvertTo-SedaArgumentList {
    param([hashtable]$Parameters)
    $args = New-Object System.Collections.Generic.List[string]
    foreach ($entry in $Parameters.GetEnumerator()) {
        if ($entry.Key -eq 'ValidateOnly') { continue }
        $value = $entry.Value
        if ($value -is [System.Management.Automation.SwitchParameter]) {
            if ($value.IsPresent) { $args.Add("-$($entry.Key)") }
            continue
        }
        if ($value -is [bool]) {
            if ($value) { $args.Add("-$($entry.Key)") }
            continue
        }
        if ($null -ne $value) {
            $args.Add("-$($entry.Key)")
            $args.Add([string]$value)
        }
    }
    return [string[]]$args
}

function Start-SedaElevatedSelf {
    param([hashtable]$Parameters)
    if ($ValidateOnly -or ($Cli -and -not $CollectLocal) -or (Test-SedaIsAdministrator)) { return }
    $runner = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    if (-not (Test-Path -LiteralPath $runner -PathType Leaf)) {
        try { $runner = (Get-Process -Id $PID).Path } catch {}
    }
    if (-not $runner) {
        $runner = Join-Path $PSHOME 'powershell.exe'
    }
    $argList = New-Object System.Collections.Generic.List[string]
    $argList.Add('-NoProfile')
    $argList.Add('-ExecutionPolicy')
    $argList.Add('Bypass')
    $argList.Add('-WindowStyle')
    $argList.Add('Hidden')
    $argList.Add('-STA')
    $argList.Add('-File')
    $argList.Add($PSCommandPath)
    foreach ($arg in (ConvertTo-SedaArgumentList -Parameters $Parameters)) { $argList.Add($arg) }
    Start-Process -FilePath $runner -ArgumentList ([string[]]$argList) -Verb RunAs -WindowStyle Hidden
    [Environment]::Exit(0)
}

Start-SedaElevatedSelf -Parameters $PSBoundParameters

$script:AppName = 'Smart Endpoint Diagnostics Analyzer'
$script:AppVersion = '0.2.0'
$script:BasePath = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:LogoPath = Join-Path $script:BasePath 'WorkplaceCloudHub.ico'
$script:WorkplaceLogoPath = Join-Path $script:BasePath 'WorkplaceCloudHub-lockup-WPF.png'
$script:AIConfigPath = Join-Path $HOME '.smartloganalyzer_ai.json'
$script:LogDirectory = Join-Path $script:BasePath 'Logs'
$script:LogRetention = 10
$script:LogPath = $null
$script:AiFocusContext = ''

$script:ImeThemes = @(
    'agentexecutor',
    'appactionprocessor',
    'appworkload',
    'clientcertcheck',
    'clienthealth',
    'devicehealthmonitoring',
    'healthscripts',
    'intunemanagementextension',
    'intuneremediations',
    'notificationinfralogs',
    'sensor',
    'win32appinventory'
)

$script:ImeThemeLabels = @{
    agentexecutor = 'Agent Executor'
    appactionprocessor = 'App Action Processor'
    appworkload = 'App Workload'
    clientcertcheck = 'Client Cert Check'
    clienthealth = 'Client Health'
    devicehealthmonitoring = 'Device Health Monitoring'
    healthscripts = 'Health Scripts'
    intunemanagementextension = 'IME (Main)'
    intuneremediations = 'Remediations'
    notificationinfralogs = 'Notifications'
    sensor = 'Sensor'
    win32appinventory = 'Win32 App Inventory'
}

$script:MdmErrorCodes = @{
    '0X80180001' = 'MDM_ENROLLMENT_FAILED'
    '0X80180002' = 'MDM_ENROLLMENT_INVALID_SERVER'
    '0X80180003' = 'MDM_ENROLLMENT_TRANSPORT_ERROR'
    '0X80180004' = 'MDM_ENROLLMENT_POLICY_ERROR'
    '0X80180005' = 'MDM_ENROLLMENT_CLIENT_CERT_FAILED'
    '0X80180006' = 'MDM_ENROLLMENT_SERVER_CERT_FAILED'
    '0X80180007' = 'MDM_ENROLLMENT_SYNC_FAILED'
    '0X80180008' = 'MDM_ENROLLMENT_NOT_SUPPORTED'
    '0X80180009' = 'MDM_ENROLLMENT_BLOCKED'
    '0X8018000A' = 'MDM_ENROLLMENT_ALREADY_ENROLLED'
    '0X80180014' = 'MDM_ENROLLMENT_QUOTA_EXCEEDED'
    '0X87D10001' = 'COMPLIANCE_FAILED'
    '0X87D10002' = 'COMPLIANCE_POLICY_NOT_FOUND'
    '0X87D11001' = 'APP_INSTALL_FAILED'
    '0X87D11003' = 'APP_NOT_APPLICABLE'
    '0X87D11004' = 'APP_INSTALL_TIMEOUT'
    '0X87D11006' = 'APP_SUPERSEDED'
    '0X87D11007' = 'APP_DETECTION_FAILED'
    '0X87D1FDE8' = 'APP_INSTALL_ERROR_GENERIC'
    '0X80070002' = 'FILE_NOT_FOUND'
    '0X80070005' = 'ACCESS_DENIED'
    '0X80070057' = 'INVALID_PARAMETER'
    '0X800704CF' = 'NETWORK_NOT_AVAILABLE'
    '0X80072EE2' = 'WINHTTP_TIMEOUT'
    '0X80072EE7' = 'WINHTTP_SERVER_NOT_FOUND'
    '0X80072EFD' = 'WINHTTP_CONNECTION_ABORTED'
    '0X80072EFE' = 'WINHTTP_CONNECTION_RESET'
    '0XCAA5001C' = 'AAD_TOKEN_BROKER_FAILED'
    '0XCAA20003' = 'AAD_INVALID_GRANT'
}

function New-SedaObject {
    param([System.Collections.IDictionary]$Property)
    New-Object psobject -Property $Property
}

function Invoke-SedaLogRotation {
    try {
        if (-not (Test-Path -LiteralPath $script:LogDirectory -PathType Container)) { return }
        Get-ChildItem -LiteralPath $script:LogDirectory -Filter 'SmartEndpointDiagnosticsAnalyzer_*.log' -File -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending |
            Select-Object -Skip $script:LogRetention |
            Remove-Item -Force -ErrorAction SilentlyContinue
    } catch {}
}

function Initialize-SedaLog {
    try {
        New-Item -ItemType Directory -Path $script:LogDirectory -Force | Out-Null
        $script:LogPath = Join-Path $script:LogDirectory ('SmartEndpointDiagnosticsAnalyzer_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
        $header = @(
            ('=' * 90)
            "$script:AppName v$script:AppVersion"
            "Started: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss zzz')"
            "Host: $env:COMPUTERNAME"
            "User: $env:USERNAME"
            "PowerShell: $($PSVersionTable.PSVersion)"
            "Process: $PID"
            ('=' * 90)
        )
        [System.IO.File]::WriteAllLines($script:LogPath, [string[]]$header, [System.Text.Encoding]::UTF8)
        Invoke-SedaLogRotation
    } catch {
        $script:LogPath = $null
    }
}

function Write-SedaLog {
    param(
        [ValidateSet('DEBUG','INFO','WARN','ERROR')]
        [string]$Level = 'INFO',
        [Parameter(Mandatory)][string]$Message,
        [System.Exception]$Exception
    )
    if (-not $script:LogPath) { return }
    try {
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
        $lines = @([regex]::Split(([string]$Message), '\r?\n') | ForEach-Object { '{0} [{1}] {2}' -f $timestamp, $Level, $_ })
        Add-Content -LiteralPath $script:LogPath -Value $lines -Encoding UTF8
        if ($Exception) {
            Add-Content -LiteralPath $script:LogPath -Value ('{0} [{1}] Exception: {2}' -f $timestamp, $Level, $Exception.Message) -Encoding UTF8
            if ($Exception.StackTrace) {
                Add-Content -LiteralPath $script:LogPath -Value ('{0} [{1}] StackTrace: {2}' -f $timestamp, $Level, ($Exception.StackTrace -replace "`r?`n", ' | ')) -Encoding UTF8
            }
        }
    } catch {}
}

Initialize-SedaLog
Write-SedaLog -Level INFO -Message 'Application initialized.'

function Get-SedaLocalBannerInfo {
    $os = ''
    try {
        $osInfo = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $os = ('{0} {1}' -f $osInfo.Caption, $osInfo.BuildNumber).Trim()
    } catch {
        $os = [System.Environment]::OSVersion.VersionString
    }
    $user = if ($env:USERDOMAIN) { "$env:USERDOMAIN\$env:USERNAME" } else { $env:USERNAME }
    $accountName = $user
    $upn = ''
    $sid = ''
    $context = 'Standard'
    try {
        $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()
        if ($identity.Name) { $accountName = $identity.Name }
        if ($identity.User) { $sid = [string]$identity.User.Value }
        $principal = New-Object System.Security.Principal.WindowsPrincipal($identity)
        if ($principal.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator)) { $context = 'Administrator' }
    } catch {}
    try {
        $whoamiUpn = (& whoami.exe /upn 2>$null | Select-Object -First 1)
        if ($whoamiUpn -and $whoamiUpn -notmatch 'ERROR|Erreur|not found') { $upn = [string]$whoamiUpn }
    } catch {}
    return New-SedaObject ([ordered]@{
        ComputerName = if ($env:COMPUTERNAME) { $env:COMPUTERNAME } else { 'Unknown' }
        User = if ($user) { $user } else { 'Unknown' }
        Account = if ($accountName) { $accountName } else { 'Unknown' }
        Upn = if ($upn) { $upn } else { 'UPN unavailable' }
        Sid = if ($sid) { $sid } else { 'SID unavailable' }
        Context = $context
        OS = if ($os) { $os } else { 'Unknown OS' }
        PowerShell = "PowerShell $($PSVersionTable.PSVersion)"
    })
}

function Get-SedaTextContent {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return '' }

    $bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($bytes.Length -eq 0) { return '' }

    $encodings = New-Object System.Collections.Generic.List[System.Text.Encoding]
    if ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
        [void]$encodings.Add([System.Text.Encoding]::Unicode)
    }
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
        [void]$encodings.Add([System.Text.Encoding]::UTF8)
    }
    [void]$encodings.Add([System.Text.Encoding]::UTF8)
    [void]$encodings.Add([System.Text.Encoding]::Unicode)
    [void]$encodings.Add([System.Text.Encoding]::GetEncoding(1252))
    [void]$encodings.Add([System.Text.Encoding]::GetEncoding(850))
    [void]$encodings.Add([System.Text.Encoding]::GetEncoding(28591))

    foreach ($encoding in $encodings) {
        try {
            $text = $encoding.GetString($bytes)
            if ($text) { return $text }
        } catch {
        }
    }

    return [System.Text.Encoding]::Default.GetString($bytes)
}

function ConvertTo-SedaHashtable {
    param([object]$Object)
    $hash = [ordered]@{}
    if ($null -eq $Object) { return $hash }
    foreach ($property in $Object.PSObject.Properties) {
        $hash[$property.Name] = $property.Value
    }
    return $hash
}

function ConvertTo-SedaKeyValueText {
    param([object]$Object)
    if ($null -eq $Object) { return '' }
    if ($Object -is [System.Collections.IDictionary]) {
        return (($Object.GetEnumerator() | ForEach-Object { "$($_.Key): $($_.Value)" }) -join [Environment]::NewLine)
    }
    if ($Object.PSObject -and $Object.PSObject.Properties.Count -gt 0) {
        return (($Object.PSObject.Properties | ForEach-Object { "$($_.Name): $($_.Value)" }) -join [Environment]::NewLine)
    }
    return [string]$Object
}

function ConvertTo-SedaKeyValueRows {
    param([object]$Object)
    $rows = @()
    if ($null -eq $Object) { return $rows }
    if ($Object -is [System.Collections.IDictionary]) {
        foreach ($entry in $Object.GetEnumerator()) {
            $rows += New-SedaObject ([ordered]@{ Property = [string]$entry.Key; Value = [string]$entry.Value })
        }
        return $rows
    }
    if ($Object.PSObject -and $Object.PSObject.Properties.Count -gt 0) {
        foreach ($property in $Object.PSObject.Properties) {
            $rows += New-SedaObject ([ordered]@{ Property = [string]$property.Name; Value = [string]$property.Value })
        }
        return $rows
    }
    return @(New-SedaObject ([ordered]@{ Property = 'Value'; Value = [string]$Object }))
}

function ConvertTo-SedaGridRows {
    param([object[]]$Rows)
    if ($null -eq $Rows) { return @() }
    return @($Rows)
}

function Get-SedaImeTheme {
    param([string]$FileName)
    $name = ([System.IO.Path]::GetFileName($FileName)).ToLowerInvariant()
    foreach ($theme in $script:ImeThemes) {
        if ($name.StartsWith($theme)) { return $theme }
    }
    return $null
}

function Get-SedaEvtxType {
    param([string]$FileName)
    $name = ([System.IO.Path]::GetFileName($FileName)).ToLowerInvariant()
    if ($name -like '*setup*') { return 'evtx_setup' }
    if ($name -like '*system*') { return 'evtx_system' }
    if ($name -like '*application*') { return 'evtx_application' }
    return $null
}

function Get-SedaZipCategory {
    param([string]$RelativePath)
    $basename = [System.IO.Path]::GetFileName($RelativePath).ToLowerInvariant()
    $dirname = ([System.IO.Path]::GetDirectoryName($RelativePath) -as [string]).ToLowerInvariant()
    $ext = [System.IO.Path]::GetExtension($RelativePath).ToLowerInvariant()

    if ($dirname -like '*foldersfiles programdata_microsoft_intunemanagementextension_logs*') { return 'ime_logs' }
    if ($dirname -like '*foldersfiles windir_logs_windowsupdate_etl*') { return 'wu_etl' }
    if ($dirname -like '*windowsupdatelog*' -and $ext -eq '.log') { return 'wu_generated_log' }
    if ($dirname -like '*extended*' -and $basename -like '*win11_readiness*.json') { return 'win11_readiness' }
    if ($dirname -like '*extended*' -and $basename -like '*win11_readiness_raw*.txt') { return 'win11_readiness_raw' }
    if ($basename -like 'windowsupdate*.log' -or $basename -like '*windowsupdate.generated*.log') { return 'wu_generated_log' }
    if ($basename -like '*windowsupdate_orchestrator*') { return 'wu_registry' }
    if ($dirname -like '*diagnosticlogcsp*') { return 'etl_logs' }
    if ($dirname -like '*panther*' -or $basename -like '*setupact*') { return 'setup_logs' }
    if ($dirname -like '*cbs*' -and $basename -like '*cbs*') { return 'cbs_logs' }
    if ($dirname -like '*computername_log*') { return 'system_logs' }

    if ($basename -match '^\(\d+\) registrykey') {
        if ($basename -like '*targetversionupgradeexperienceindicators*' -or $basename -like '*win11_upgrade_compatibility_indicators*') { return 'reg_win11_upgrade_indicators' }
        if ($basename -like '*cloudmanagedupdate*') { return 'reg_cloudmanagedupdate' }
        if ($basename -like '*currentversion_uninstall*' -and $basename -notlike '*wow6432*') { return 'reg_uninstall_x64' }
        if ($basename -like '*wow6432node*' -and $basename -like '*uninstall*') { return 'reg_uninstall_x86' }
        return 'registry'
    }
    if ($basename -match '^\(\d+\) command') {
        if ($basename -like '*pnputil*') { return 'cmd_pnputil' }
        if ($basename -like '*wlan_show_profiles*') { return 'cmd_wlan_profiles' }
        if ($basename -like '*certutil*') { return 'cmd_certutil' }
        return 'command_output'
    }
    if ($basename -match '^\(\d+\) events') { return 'event_logs' }
    if ($basename -match '^\(\d+\) folderfiles') { return 'folder_files' }
    if ($basename -match '^\(\d+\) no results') { return 'collection_errors' }
    if ($basename -eq 'results.xml') { return 'results_xml' }
    if ($ext -eq '.evtx') { return 'event_logs' }
    if ($ext -eq '.etl') { return 'etl_logs' }
    if ($ext -eq '.reg') { return 'registry' }
    if ($ext -eq '.log') { return 'log_files' }
    if ($ext -eq '.cab') { return 'cab' }
    if ($ext -eq '.xml') { return 'xml' }
    if ($ext -eq '.json') { return 'json' }
    if ($ext -eq '.html') {
        if ($basename -like '*battery*') { return 'battery_report' }
        return 'html'
    }
    return 'unknown'
}

function Expand-SedaDiagnosticZip {
    param([Parameter(Mandatory)][string]$Path)
    Write-SedaLog -Level INFO -Message "Expanding diagnostic ZIP: $Path"
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "ZIP file not found: $Path"
    }

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $tempRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('seda_diag_' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null
    [System.IO.Compression.ZipFile]::ExtractToDirectory($Path, $tempRoot)

    $inventory = [ordered]@{
        all_files = @()
        ime_themes = @{}
    }

    $files = Get-ChildItem -LiteralPath $tempRoot -Recurse -File -Force
    foreach ($file in $files) {
        $relative = $file.FullName.Substring($tempRoot.Length).TrimStart('\', '/')
        $category = Get-SedaZipCategory -RelativePath $relative
        $inventory.all_files += $file.FullName
        if (-not $inventory.Contains($category)) { $inventory[$category] = @() }
        $inventory[$category] += $file.FullName

        if ($category -eq 'log_files') {
            $lower = $file.Name.ToLowerInvariant()
            if ($lower.StartsWith('autopatchclient')) {
                if (-not $inventory.Contains('autopatch_logs')) { $inventory.autopatch_logs = @() }
                $inventory.autopatch_logs += $file.FullName
            } elseif ($lower.StartsWith('wingetcom')) {
                if (-not $inventory.Contains('winget_logs')) { $inventory.winget_logs = @() }
                $inventory.winget_logs += $file.FullName
            }
        }

        if ($category -eq 'ime_logs') {
            $theme = Get-SedaImeTheme -FileName $file.Name
            if ($theme) {
                if (-not $inventory.ime_themes.Contains($theme)) { $inventory.ime_themes[$theme] = @() }
                $inventory.ime_themes[$theme] += $file.FullName
            }
        }

        if ($category -eq 'event_logs' -and $file.Extension.ToLowerInvariant() -eq '.evtx') {
            $evtxType = Get-SedaEvtxType -FileName $file.Name
            if ($evtxType) {
                if (-not $inventory.Contains($evtxType)) { $inventory[$evtxType] = @() }
                $inventory[$evtxType] += $file.FullName
            }
        }
    }

    $item = Get-Item -LiteralPath $Path
    Write-SedaLog -Level INFO -Message "ZIP expanded to $tempRoot with $(@($files).Count) files."
    return New-SedaObject @{
        ZipPath = $Path
        ZipName = $item.Name
        ZipSizeMb = [math]::Round($item.Length / 1MB, 1)
        ExtractDir = $tempRoot
        Inventory = $inventory
    }
}

function Find-SedaInventoryFile {
    param(
        [hashtable]$Inventory,
        [string]$Keyword
    )
    $keywordLower = $Keyword.ToLowerInvariant()
    foreach ($file in @($Inventory.all_files)) {
        if ([System.IO.Path]::GetFileName($file).ToLowerInvariant().Contains($keywordLower)) {
            return $file
        }
    }
    return ''
}

function ConvertFrom-SedaRegValue {
    param([string]$Value)
    if ($null -eq $Value) { return '' }
    $raw = $Value.Trim()
    if ($raw.StartsWith('"') -and $raw.EndsWith('"')) {
        if ($raw.Length -lt 2) { return '' }
        return $raw.Substring(1, $raw.Length - 2).Replace('\\"', '"').Replace('\\\\', '\')
    }
    if ($raw -match '^dword:([0-9a-fA-F]+)') {
        return [string][Convert]::ToInt64($Matches[1], 16)
    }
    if ($raw -match '^hex\(2\):(.+)') {
        try {
            $bytes = ($Matches[1] -split ',' | Where-Object { $_ }) | ForEach-Object { [Convert]::ToByte($_, 16) }
            return ([System.Text.Encoding]::Unicode.GetString([byte[]]$bytes)).Trim([char]0)
        } catch {
            return $raw
        }
    }
    return $raw
}

function ConvertFrom-SedaRegFile {
    param([Parameter(Mandatory)][string]$Path)
    $content = Get-SedaTextContent -Path $Path
    $keys = [ordered]@{}
    $current = ''
    $pending = ''
    foreach ($rawLine in ($content -split "`r?`n")) {
        $line = $rawLine.Trim()
        if (-not $line -or $line.StartsWith(';')) { continue }
        if ($pending) {
            $line = $pending + $line
            $pending = ''
        }
        if ($line.EndsWith('\')) {
            $pending = $line.Substring(0, $line.Length - 1).TrimEnd()
            continue
        }
        if ($line.StartsWith('[') -and $line.EndsWith(']')) {
            $current = $line.Substring(1, $line.Length - 2)
            if (-not $keys.Contains($current)) { $keys[$current] = [ordered]@{} }
            continue
        }
        if ($current -and $line.Contains('=')) {
            $name, $value = $line -split '=', 2
            $keyName = $name.Trim().Trim('"')
            $keys[$current][$keyName] = ConvertFrom-SedaRegValue -Value $value
        }
    }
    return $keys
}

function ConvertFrom-SedaDword {
    param([string]$Value)
    if ($Value -match 'dword:([0-9a-fA-F]+)') { return [Convert]::ToInt64($Matches[1], 16) }
    $n = 0
    if ([int]::TryParse([string]$Value, [ref]$n)) { return $n }
    return -1
}

function Get-SedaDsRegCmd {
    param([string]$Path)
    $sections = [ordered]@{}
    $issues = @()
    $deviceInfo = [ordered]@{}
    $ssoInfo = [ordered]@{}
    if (-not $Path) {
        return New-SedaObject @{ Sections = $sections; DeviceInfo = $deviceInfo; SsoInfo = $ssoInfo; CriticalIssues = $issues; RawText = '' }
    }

    $rawText = Get-SedaTextContent -Path $Path
    $current = 'Header'
    $sections[$current] = [ordered]@{}
    foreach ($line in ($rawText -split "`r?`n")) {
        if ($line -match '\|\s*(.+?)\s*\|' -and $line -notmatch '---|===') {
            $title = $Matches[1].Trim()
            if ($title.Length -gt 2) {
                $current = $title
                if (-not $sections.Contains($current)) { $sections[$current] = [ordered]@{} }
            }
            continue
        }
        if ($line -match '^\s{2,}(\S[^:]*?)\s*:\s*(.*?)\s*$') {
            if (-not $sections.Contains($current)) { $sections[$current] = [ordered]@{} }
            $sections[$current][$Matches[1].Trim()] = $Matches[2].Trim()
        }
    }

    $dev = $sections['Device State']
    $detail = $sections['Device Details']
    $tenant = $sections['Tenant Details']
    $sso = $sections['SSO State']
    $user = $sections['User State']

    foreach ($map in @(
        @{ Out = 'Device Name'; In = 'Device Name'; Source = $dev },
        @{ Out = 'AAD Joined'; In = 'AzureAdJoined'; Source = $dev },
        @{ Out = 'Domain Joined'; In = 'DomainJoined'; Source = $dev },
        @{ Out = 'Domain Name'; In = 'DomainName'; Source = $dev },
        @{ Out = 'Virtual Desktop'; In = 'Virtual Desktop'; Source = $dev },
        @{ Out = 'Device ID'; In = 'DeviceId'; Source = $detail },
        @{ Out = 'TPM Protected'; In = 'TpmProtected'; Source = $detail },
        @{ Out = 'Device Auth Status'; In = 'DeviceAuthStatus'; Source = $detail },
        @{ Out = 'Tenant Name'; In = 'TenantName'; Source = $tenant },
        @{ Out = 'Tenant ID'; In = 'TenantId'; Source = $tenant },
        @{ Out = 'MDM URL'; In = 'MdmUrl'; Source = $tenant }
    )) {
        if ($map.Source -and $map.Source[$map.In]) { $deviceInfo[$map.Out] = $map.Source[$map.In] }
    }

    foreach ($source in @($sso, $user)) {
        if ($source) {
            foreach ($key in $source.Keys) { $ssoInfo[$key] = $source[$key] }
        }
    }

    if ($sso -and ([string]$sso.AzureAdPrt).ToUpperInvariant() -eq 'NO') {
        $attempt = [string]$sso['Attempt Status']
        $serverError = [string]$sso['Server Error Code']
        $serverDescription = [string]$sso['Server Error Description']
        $httpStatus = [string]$sso['HTTP status']
        $detailText = 'AzureAdPrt=NO'
        if ($attempt) { $detailText += " | AttemptStatus: $attempt" }
        if ($serverError) { $detailText += " | ServerError: $serverError" }
        if ($httpStatus -and $httpStatus -ne '0') { $detailText += " | HTTP: $httpStatus" }
        $recommendation = ''
        if ($serverDescription -like '*AADSTS50034*') { $recommendation = 'User does not exist in the tenant. Check UPN and Entra Connect synchronization.' }
        elseif ($serverError -like '*invalid_grant*') { $recommendation = 'Invalid token. Check synchronization and Conditional Access policies.' }
        elseif ($attempt -eq '0xc0000072') { $recommendation = 'Account may be disabled or locked in Active Directory.' }
        elseif ($serverDescription -like '*AADSTS50076*') { $recommendation = 'MFA is required. The user must complete MFA authentication.' }
        $issues += New-SedaObject @{ Severity = 'ERROR'; Category = 'AAD PRT'; Title = 'Primary Refresh Token (PRT) not acquired'; Detail = $detailText; Recommendation = $recommendation; Source = 'DSRegCmd' }
    }

    $wam = if ($sso) { [string]$sso.WamDefaultSet } else { '' }
    if ($wam.ToUpperInvariant().Contains('ERROR')) {
        $issues += New-SedaObject @{ Severity = 'WARNING'; Category = 'WAM'; Title = "WAM Default Set error: $wam"; Detail = "WamDefaultSet=$wam"; Recommendation = 'WAM errors can block single sign-on authentication.'; Source = 'DSRegCmd' }
    }

    if ($dev -and ([string]$dev.AzureAdJoined).ToUpperInvariant() -eq 'NO') {
        $issues += New-SedaObject @{ Severity = 'ERROR'; Category = 'AAD Join'; Title = 'Device is not Azure AD joined'; Detail = "AzureAdJoined=NO, DomainJoined=$($dev.DomainJoined)"; Recommendation = 'The device should be Hybrid Azure AD Joined, Azure AD Joined, or Workplace Joined depending on ownership and enrollment model.'; Source = 'DSRegCmd' }
    }

    return New-SedaObject @{ Sections = $sections; DeviceInfo = $deviceInfo; SsoInfo = $ssoInfo; CriticalIssues = $issues; RawText = $rawText }
}

function Get-SedaEnrollments {
    param([string]$Path)
    $enrollments = @()
    $summary = [ordered]@{}
    if (-not $Path) { return New-SedaObject @{ Enrollments = $enrollments; Summary = $summary } }
    $keys = ConvertFrom-SedaRegFile -Path $Path
    $root = 'hkey_local_machine\software\microsoft\enrollments'
    $stateLabels = @{ '0' = 'Not enrolled'; '1' = 'Enrolled (active)'; '2' = 'Enrollment in progress'; '3' = 'Unenrollment in progress'; '4' = 'Pending renewal' }
    $typeLabels = @{ '0' = 'Unknown'; '1' = 'DeviceEnrollment (MDM Device)'; '2' = 'UserEnrollment (MDM User)'; '6' = 'ExternallyManaged'; '18' = 'Autopilot'; '30' = 'DeployAuthority' }
    foreach ($key in $keys.Keys) {
        $lower = $key.ToLowerInvariant()
        if ($lower -eq $root) { continue }
        if (-not $lower.StartsWith($root + '\')) { continue }
        $rel = $lower.Substring($root.Length).Trim('\')
        if (-not $rel -or $rel.Contains('\')) { continue }
        $values = $keys[$key]
        $state = [string](ConvertFrom-SedaDword -Value $values.EnrollmentState)
        $type = [string](ConvertFrom-SedaDword -Value $values.EnrollmentType)
        $item = New-SedaObject @{
            GUID = $rel.ToUpperInvariant()
            State = if ($stateLabels[$state]) { $stateLabels[$state] } else { "Unknown ($state)" }
            Type = if ($typeLabels[$type]) { $typeLabels[$type] } else { "Unknown ($type)" }
            ProviderID = [string]$values.ProviderID
            UPN = [string]$values.UPN
            EnrollmentURL = [string]$values.EnrollmentURL
            DiscoveryServiceFullURL = [string]$values.DiscoveryServiceFullURL
            AADResourceID = [string]$values.AADResourceID
        }
        $enrollments += $item
        if ($item.UPN) { $summary.UPN = $item.UPN }
        if ($item.EnrollmentURL) { $summary['Enrollment URL'] = $item.EnrollmentURL }
        elseif ($item.DiscoveryServiceFullURL) { $summary['Enrollment URL'] = $item.DiscoveryServiceFullURL }
    }
    if ($enrollments.Count -gt 0) { $summary['Enrollment Count'] = [string]$enrollments.Count }
    return New-SedaObject @{ Enrollments = $enrollments; Summary = $summary }
}

function Get-SedaResultsXml {
    param([string[]]$Paths)
    $items = @()
    $errors = @()
    $codes = @{
        '0' = 'SUCCESS'
        '-2147024893' = 'Success (key collected)'
        '-2147024895' = 'ERROR - Registry key not found (0x80070001)'
        '-2147418113' = 'ERROR - Unspecified failure (0x8000ffff)'
    }
    foreach ($path in @($Paths)) {
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
        try {
            [xml]$xml = Get-SedaTextContent -Path $path
            foreach ($child in $xml.DocumentElement.ChildNodes) {
                $hresult = [string]$child.HRESULT
                if (-not $hresult) { $hresult = '0' }
                $nameNode = $child.SelectSingleNode('Name')
                $name = if ($nameNode -and $nameNode.InnerText) { $nameNode.InnerText.Trim() } elseif ($child.InnerText) { $child.InnerText.Trim() } else { $child.Name }
                $ok = $hresult -in @('0', '-2147024893')
                $item = New-SedaObject @{
                    Type = $child.Name
                    Name = $name
                    HResult = $hresult
                    Status = if ($codes[$hresult]) { $codes[$hresult] } else { "Code: $hresult" }
                    Ok = $ok
                }
                $items += $item
                if (-not $ok) { $errors += $item }
            }
        } catch {
        }
    }
    return New-SedaObject @{ Items = $items; Errors = $errors }
}

function Get-SedaFirewallIssues {
    param([string]$Path)
    $profiles = [ordered]@{}
    $issues = @()
    if (-not $Path) { return New-SedaObject @{ Profiles = $profiles; Issues = $issues } }
    $current = ''
    foreach ($rawLine in (Get-SedaTextContent -Path $Path -split "`r?`n")) {
        $line = $rawLine.Replace([char]0x00A0, ' ')
        if ($line -match '^(\w+)\s+Profile\s+Settings:') {
            $current = (Get-Culture).TextInfo.ToTitleCase($Matches[1].ToLowerInvariant())
            $profiles[$current] = [ordered]@{}
            continue
        }
        if ($line -match '^Param[eè]tres\s+Profil\s+(?:de\s+)?(\w+)\s*[:\-]') {
            $raw = $Matches[1].ToLowerInvariant().Replace('é', 'e').Replace('è', 'e')
            $map = @{ domaine = 'Domain'; domain = 'Domain'; prive = 'Private'; private = 'Private'; public = 'Public' }
            $current = if ($map[$raw]) { $map[$raw] } else { $Matches[1] }
            $profiles[$current] = [ordered]@{}
            continue
        }
        if ($current -and $line.Trim() -and $line -match '^(.+?)\s{2,}(\S.*?)\s*$') {
            $key = $Matches[1].Trim()
            $value = $Matches[2].Trim()
            if ($key.Length -lt 60) { $profiles[$current][$key] = $value }
        }
    }
    foreach ($profile in $profiles.Keys) {
        $settings = $profiles[$profile]
        $state = [string]($settings.State, $settings.Etat, $settings.'État' | Where-Object { $_ } | Select-Object -First 1)
        $inbound = [string]($settings.'Inbound connections', $settings.'Strategie de pare-feu', $settings.'Stratégie de pare-feu' | Where-Object { $_ } | Select-Object -First 1)
        $stateUpper = $state.ToUpperInvariant()
        if ($stateUpper.Contains('OFF') -or $stateUpper.Contains('INACTIF') -or $stateUpper.Contains('DESACTIV')) {
            $issues += New-SedaObject @{ Severity = 'ERROR'; Category = 'Firewall'; Title = "Firewall disabled - $profile profile"; Detail = "Profile $profile`: State=$state"; Recommendation = 'Enable Windows Firewall for this profile.'; Source = 'netsh advfirewall' }
        }
        $inboundUpper = $inbound.ToUpperInvariant().Replace(' ', '')
        if ($inboundUpper.Contains('ALLOWINBOUND') -or $inboundUpper -eq 'ALLOW' -or $inboundUpper.Contains('AUTORISERENTRANT')) {
            $issues += New-SedaObject @{ Severity = 'WARNING'; Category = 'Firewall'; Title = "Inbound connections allowed - $profile"; Detail = "Profile $profile`: Inbound=$inbound"; Recommendation = 'Verify whether this firewall posture is intentional.'; Source = 'netsh advfirewall' }
        }
    }
    return New-SedaObject @{ Profiles = $profiles; Issues = $issues }
}

function Get-SedaImeLogEvents {
    param(
        [hashtable]$ImeThemes
    )
    $events = @()
    $scanned = @()
    $imeRegex = [regex]::new('<!\[LOG\[(.*?)\]LOG\]!\><time="([\d:.]+)"[^>]*date="([\d-]+)"[^>]*component="([^"]*)"[^>]*type="(\d+)"', [System.Text.RegularExpressions.RegexOptions]::Singleline)
    $hexRegex = [regex]'0[xX][0-9A-Fa-f]{4,8}'
    $errorWords = @('error', 'failed', 'failure', 'exception', 'critical')
    $warningWords = @('warning', 'warn', 'timeout', 'retry')
    $ignoreWords = @('no error', 'no failure', 'errorlevel 0', '0 error', 'success')

    foreach ($theme in @($ImeThemes.Keys)) {
        foreach ($file in @($ImeThemes[$theme])) {
            if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { continue }
            $item = Get-Item -LiteralPath $file
            if ($item.Length -gt 20MB) { continue }
            $scanned += $file
            $content = Get-SedaTextContent -Path $file
            $short = [System.IO.Path]::GetFileName($file)
            if ($content.Contains('<![LOG[')) {
                $seen = @{}
                foreach ($match in $imeRegex.Matches($content)) {
                    $message = $match.Groups[1].Value.Trim()
                    $type = $match.Groups[5].Value
                    if ($type -notin @('2', '3')) { continue }
                    $severity = if ($type -eq '3') { 'ERROR' } else { 'WARNING' }
                    $dedupe = "$short|$($message.Substring(0, [Math]::Min(80, $message.Length)))"
                    if ($seen[$dedupe]) { continue }
                    $seen[$dedupe] = $true
                    $code = ''
                    $known = ''
                    foreach ($cm in $hexRegex.Matches($message)) {
                        $candidate = $cm.Value.ToUpperInvariant()
                        if ($script:MdmErrorCodes[$candidate]) {
                            $code = $cm.Value
                            $known = $script:MdmErrorCodes[$candidate]
                            $severity = 'ERROR'
                            break
                        }
                    }
                    $events += New-SedaObject @{
                        Severity = $severity
                        Category = if ($match.Groups[4].Value) { $match.Groups[4].Value } else { $theme.ToUpperInvariant() }
                        Message = $message.Substring(0, [Math]::Min(300, $message.Length))
                        SourceFile = $short
                        Theme = $theme
                        LineNumber = 0
                        RawLine = $message.Substring(0, [Math]::Min(500, $message.Length))
                        Timestamp = "$($match.Groups[3].Value) $($match.Groups[2].Value)"
                        ErrorCode = $code
                        KnownCode = $known
                        LogFormat = 'ime'
                    }
                }
            } else {
                $lineNumber = 0
                foreach ($line in ($content -split "`r?`n")) {
                    $lineNumber++
                    $lower = $line.ToLowerInvariant()
                    if ($ignoreWords | Where-Object { $lower.Contains($_) }) { continue }
                    $severity = ''
                    if ($errorWords | Where-Object { $lower.Contains($_) }) { $severity = 'ERROR' }
                    elseif ($warningWords | Where-Object { $lower.Contains($_) }) { $severity = 'WARNING' }
                    if (-not $severity) {
                        foreach ($cm in $hexRegex.Matches($line)) {
                            if ($script:MdmErrorCodes[$cm.Value.ToUpperInvariant()]) { $severity = 'ERROR'; break }
                        }
                    }
                    if (-not $severity) { continue }
                    $code = ''
                    $known = ''
                    foreach ($cm in $hexRegex.Matches($line)) {
                        $candidate = $cm.Value.ToUpperInvariant()
                        if ($script:MdmErrorCodes[$candidate]) {
                            $code = $cm.Value
                            $known = $script:MdmErrorCodes[$candidate]
                            break
                        }
                    }
                    $trimmed = $line.Trim()
                    $events += New-SedaObject @{
                        Severity = $severity
                        Category = 'Log'
                        Message = $trimmed.Substring(0, [Math]::Min(300, $trimmed.Length))
                        SourceFile = $short
                        Theme = $theme
                        LineNumber = $lineNumber
                        RawLine = $trimmed.Substring(0, [Math]::Min(500, $trimmed.Length))
                        Timestamp = ''
                        ErrorCode = $code
                        KnownCode = $known
                        LogFormat = 'text'
                    }
                }
            }
        }
    }

    $themeCounts = [ordered]@{}
    foreach ($theme in $script:ImeThemes) {
        $themeEvents = @($events | Where-Object { $_.Theme -eq $theme })
        if ($themeEvents.Count -gt 0 -or $ImeThemes.Contains($theme)) {
            $themeCounts[$theme] = New-SedaObject @{
                Errors = @($themeEvents | Where-Object { $_.Severity -eq 'ERROR' }).Count
                Warnings = @($themeEvents | Where-Object { $_.Severity -eq 'WARNING' }).Count
            }
        }
    }

    return New-SedaObject @{
        Events = $events
        Summary = New-SedaObject @{
            ErrorCount = @($events | Where-Object { $_.Severity -eq 'ERROR' }).Count
            WarningCount = @($events | Where-Object { $_.Severity -eq 'WARNING' }).Count
            TotalEvents = $events.Count
            ScannedFiles = @($scanned | Select-Object -Unique).Count
            ThemeCounts = $themeCounts
        }
    }
}

function Get-SedaFileTimeString {
    param([UInt64]$FileTime)
    try {
        return [DateTime]::FromFileTimeUtc([int64]$FileTime).ToString('yyyy-MM-dd HH:mm:ss UTC')
    } catch {
        return ('<filetime 0x{0:x16}>' -f $FileTime)
    }
}

function Get-SedaWindowsUpdateInfo {
    param(
        [string]$OrchestratorReg,
        [string[]]$AllRegFiles,
        [string]$ReportingEvents
    )
    $info = [ordered]@{}
    $issues = @()
    $policies = @()
    $reporting = @()

    $labels = @{
        OsVersion = 'OS Version'
        BuildString = 'OS Build'
        UBR = 'Update Build Revision (UBR)'
        Preshutdown = 'Pre-shutdown Reboot Required'
        RebootRequired = 'Reboot Required'
        RebootRequiredReason = 'Reboot Required Reason'
        DisabledAutomaticRestarts = 'Automatic Restarts Disabled'
        UpdateManagerCtorFailures = 'Update Manager Init Failures'
        MoStackEnabled = 'Modern Orchestrator Enabled'
        OobeCompleteTimeStamp = 'OOBE Completion Date'
        WUfBPolicyHash = 'WUfB Policy Hash'
        PolicyReportHash = 'WUfB Policy Hash'
        PolicyReportTimestamp = 'WUfB Policy Sync Date'
        SettingsETag = 'WUfB Settings ETag'
        SettingsRefreshInterval = 'Settings Refresh Interval'
        ScanTriggerTime = 'Last Scan Trigger Time'
        PerformScanTriggerTime = 'Next Scan Trigger Time'
        InstallTriggerTime = 'Last Install Trigger Time'
        NextRefreshTime = 'Next WU Refresh Time'
        FeatureUpdatePauseEnabled = 'Feature Update Pause Enabled'
        QualityUpdatePauseEnabled = 'Quality Update Pause Enabled'
        DeferFeatureUpdatePeriodInDays = 'Feature Update Deferral (days)'
        DeferQualityUpdatePeriodInDays = 'Quality Update Deferral (days)'
        FlightInfo = 'Flight Info'
        FlightPendingCommit = 'Flight Pending Commit'
    }
    $boolValues = @('Preshutdown', 'RebootRequired', 'DisabledAutomaticRestarts', 'FeatureUpdatePauseEnabled', 'QualityUpdatePauseEnabled')

    if ($OrchestratorReg) {
        $keys = ConvertFrom-SedaRegFile -Path $OrchestratorReg
        foreach ($key in $keys.Keys) {
            if ($key.ToLowerInvariant() -notlike '*\software\microsoft\windows\currentversion\windowsupdate\orchestrator*') { continue }
            foreach ($name in $keys[$key].Keys) {
                $label = if ($labels[$name]) { $labels[$name] } else { $name }
                $value = $keys[$key][$name]
                if ($value -match '^\d+$') {
                    $n = [UInt64]$value
                    if ($boolValues -contains $name) { $value = if ($n -eq 0) { 'No' } else { 'Yes' } }
                    elseif ($name -in @('DeferFeatureUpdatePeriodInDays', 'DeferQualityUpdatePeriodInDays')) { $value = "$n day(s)" }
                    elseif ($name -eq 'UBR') { $value = ('{0}  (hex: 0x{0:x4})' -f $n) }
                }
                $info[$label] = $value
            }
        }

        if ($info['Reboot Required'] -eq 'Yes') {
            $issues += New-SedaObject @{ Severity = 'WARNING'; Title = 'Reboot required'; Detail = 'RebootRequired=Yes in WU Orchestrator registry'; Recommendation = 'Device needs a reboot to complete pending updates.'; Source = 'Windows Update' }
        }
        if ($info['Pre-shutdown Reboot Required'] -eq 'Yes') {
            $issues += New-SedaObject @{ Severity = 'WARNING'; Title = 'Pre-shutdown reboot pending'; Detail = 'Preshutdown=Yes'; Recommendation = 'Ensure device reboots to complete Windows Update installation.'; Source = 'Windows Update' }
        }
        $failures = 0
        if ([int]::TryParse([string]$info['Update Manager Init Failures'], [ref]$failures) -and $failures -gt 0) {
            $issues += New-SedaObject @{ Severity = 'ERROR'; Title = "Update Manager initialization failed $failures time(s)"; Detail = "UpdateManagerCtorFailures=$failures"; Recommendation = 'Check Windows Update service state and event logs.'; Source = 'Windows Update' }
        }
        foreach ($label in @('Feature Update Pause Enabled', 'Quality Update Pause Enabled')) {
            if ($info[$label] -eq 'Yes') {
                $issues += New-SedaObject @{ Severity = 'WARNING'; Title = "$label"; Detail = "$label=Yes"; Recommendation = 'Verify that update pause is intentional via WUfB policy.'; Source = 'Windows Update' }
            }
        }
    }

    $wuPolicyRoot = 'hkey_local_machine\software\policies\microsoft\windows\windowsupdate'
    $policyLabels = @{
        WUServer = 'WSUS Server URL'
        WUStatusServer = 'WSUS Status Server URL'
        TargetGroup = 'Target Group Name'
        TargetGroupEnabled = 'Target Group Enabled'
        DisableWindowsUpdateAccess = 'Disable WU Access for non-admins'
        DisableDualScan = 'Disable Dual Scan (force WSUS only)'
        ExcludeWUDriversInQualityUpdate = 'Exclude Drivers from Quality Updates'
        SetPolicyDrivenUpdateSourceForFeatureUpdates = 'Policy-Driven Source: Feature Updates'
        SetPolicyDrivenUpdateSourceForQualityUpdates = 'Policy-Driven Source: Quality Updates'
        SetPolicyDrivenUpdateSourceForDriverUpdates = 'Policy-Driven Source: Driver Updates'
        SetPolicyDrivenUpdateSourceForOtherUpdates = 'Policy-Driven Source: Other Updates'
        DeferFeatureUpdatesPeriodInDays = 'Feature Update Deferral (days)'
        DeferQualityUpdatesPeriodInDays = 'Quality Update Deferral (days)'
        ProductVersion = 'Product Version (upgrade target)'
        TargetReleaseVersionInfo = 'Target Release Version'
        NoAutoUpdate = 'Disable Automatic Updates'
        AUOptions = 'Automatic Updates Option'
        UseWUServer = 'Use WSUS Server (AU)'
    }
    foreach ($regFile in @($AllRegFiles)) {
        if (-not $regFile) { continue }
        $keys = ConvertFrom-SedaRegFile -Path $regFile
        foreach ($key in $keys.Keys) {
            if (-not $key.ToLowerInvariant().StartsWith($wuPolicyRoot)) { continue }
            foreach ($name in $keys[$key].Keys) {
                $displayKey = $key.Substring($key.ToLowerInvariant().IndexOf('windowsupdate'))
                $value = $keys[$key][$name]
                $policies += New-SedaObject @{
                    KeyPath = $displayKey
                    Name = $name
                    Label = if ($policyLabels[$name]) { $policyLabels[$name] } else { $name }
                    Value = $value
                    Source = [System.IO.Path]::GetFileName($regFile)
                }
            }
        }
    }

    if ($ReportingEvents) {
        $text = Get-SedaTextContent -Path $ReportingEvents
        foreach ($line in ($text -split "`r?`n")) {
            if (-not $line.Trim()) { continue }
            $fields = $line -split "`t"
            if ($fields.Count -lt 12) { continue }
            $eventString = $fields[3].Trim()
            $hresult = $fields[7].Trim()
            $status = $fields[9].Trim()
            $message = $fields[11].Trim()
            $eventName = if ($eventString -match '\[([A-Z_]+)\]') { $Matches[1] } else { $eventString }
            $level = ''
            if ($eventName -in @('AGENT_DETECTION_FAILED', 'AGENT_INSTALLING_FAILED', 'AGENT_DOWNLOAD_FAILED') -or $status -eq 'Failure') { $level = 'Error' }
            elseif ($eventName -eq 'AGENT_DOWNLOAD_CANCELED') { $level = 'Warning' }
            elseif ($hresult -and $hresult -ne '0') {
                try {
                    $hr = [Convert]::ToInt64($hresult, 16)
                    if (($hr -band 0x80000000) -ne 0) { $level = 'Error' }
                } catch {
                }
            }
            if (-not $level) { continue }
            $errorCode = ''
            if ($hresult -and $hresult -ne '0') {
                try { $errorCode = ('0x{0:X8}' -f [Convert]::ToInt64($hresult, 16)) } catch { $errorCode = $hresult }
            }
            $timestamp = ($fields[1].Trim() -replace ':\d{3}[+-]\d{4}$', '')
            $reporting += New-SedaObject @{ Level = $level; Timestamp = $timestamp; Source = $eventName; Message = $message; ErrorCode = $errorCode; File = [System.IO.Path]::GetFileName($ReportingEvents) }
        }
    }

    return New-SedaObject @{ Info = $info; Issues = $issues; Policies = $policies; ReportingEvents = $reporting }
}

function Convert-SedaWindowsUpdateEtlsToLog {
    param(
        [string[]]$Paths,
        [string]$OutputDirectory
    )
    $files = @($Paths | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 50)
    if ($files.Count -eq 0) { return '' }
    $cmd = Get-Command Get-WindowsUpdateLog -ErrorAction SilentlyContinue
    if (-not $cmd) {
        Write-SedaLog -Level WARN -Message 'Get-WindowsUpdateLog is not available on this host.'
        return ''
    }
    if (-not $OutputDirectory) { $OutputDirectory = Join-Path ([System.IO.Path]::GetTempPath()) ('seda_wulog_' + [guid]::NewGuid().ToString('N')) }
    New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
    $logPath = Join-Path $OutputDirectory ('WindowsUpdate.generated_{0:yyyyMMdd_HHmmss}.log' -f (Get-Date))
    try {
        Write-SedaLog -Level INFO -Message "Generating WindowsUpdate.log with Get-WindowsUpdateLog from $($files.Count) ETL file(s)."
        Get-WindowsUpdateLog -ETLPath $files -LogPath $logPath -ErrorAction Stop *> (Join-Path $OutputDirectory 'Get-WindowsUpdateLog.output.txt')
        if (Test-Path -LiteralPath $logPath -PathType Leaf) {
            Write-SedaLog -Level INFO -Message "Generated WindowsUpdate.log: $logPath"
            return $logPath
        }
    } catch {
        Write-SedaLog -Level WARN -Message 'Get-WindowsUpdateLog failed for ETL conversion.' -Exception $_.Exception
    }
    return ''
}

function Get-SedaWindowsUpdateLogEvents {
    param([string]$Path)
    $events = @()
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $events }
    $errKeywords = [regex]'(?i)\b(error|failed|failure|fatal|0x[0-9a-f]{8}|hr=|hresult|denied|timeout)\b'
    $warnKeywords = [regex]'(?i)\b(warn|warning|retry|defer|blocked|pending|paused)\b'
    $sourceFile = [System.IO.Path]::GetFileName($Path)
    foreach ($line in (Get-SedaTextContent -Path $Path -split "`r?`n")) {
        $trimmed = $line.Trim()
        if (-not $trimmed) { continue }
        $level = ''
        if ($trimmed -match $errKeywords) { $level = 'Error' }
        elseif ($trimmed -match $warnKeywords) { $level = 'Warning' }
        else { continue }
        $code = ''
        if ($trimmed -match '(?i)(0x[0-9a-f]{8})') { $code = $Matches[1] }
        $timestamp = ''
        if ($trimmed -match '^(\d{4}[/-]\d{2}[/-]\d{2}\s+\d{2}:\d{2}:\d{2}(?:\.\d+)?)') { $timestamp = $Matches[1] }
        $source = 'WindowsUpdate.log'
        if ($trimmed -match '^\S+\s+\S+\s+\S+\s+\S+\s+([A-Za-z0-9_.-]+)\s+') { $source = $Matches[1] }
        $events += New-SedaObject @{ Level=$level; Timestamp=$timestamp; EventId=''; Source=$source; Message=$trimmed.Substring(0, [Math]::Min(500, $trimmed.Length)); ErrorCode=$code; EtlFile=$sourceFile; LogFile=$Path }
    }
    return $events
}

function Get-SedaKeywordLogEvents {
    param(
        [string[]]$Paths,
        [string]$Area = 'Log',
        [int]$MaxLinesPerFile = 6000
    )
    $events = @()
    $errKeywords = [regex]'(?i)\b(error|failed|failure|fatal|exception|denied|timeout|0x[0-9a-f]{8})\b'
    $warnKeywords = [regex]'(?i)\b(warn|warning|retry|blocked|pending|deprecated|fallback)\b'
    foreach ($path in @($Paths)) {
        if (-not $path -or -not (Test-Path -LiteralPath $path -PathType Leaf)) { continue }
        $file = [System.IO.Path]::GetFileName($path)
        $lineNumber = 0
        foreach ($line in (Get-SedaTextContent -Path $path -split "`r?`n")) {
            $lineNumber++
            if ($lineNumber -gt $MaxLinesPerFile) { break }
            $trimmed = $line.Trim()
            if (-not $trimmed) { continue }
            $level = ''
            if ($trimmed -match $errKeywords) { $level = 'Error' }
            elseif ($trimmed -match $warnKeywords) { $level = 'Warning' }
            else { continue }
            $code = ''
            if ($trimmed -match '(?i)(0x[0-9a-f]{8})') { $code = $Matches[1] }
            $timestamp = ''
            if ($trimmed -match '(\d{4}[-/]\d{2}[-/]\d{2}[ T]\d{2}:\d{2}:\d{2}(?:\.\d+)?)') { $timestamp = $Matches[1] }
            elseif ($trimmed -match '(\d{1,2}[-/]\d{1,2}[-/]\d{4}\s+\d{1,2}:\d{2}:\d{2}(?:\.\d+)?)') { $timestamp = $Matches[1] }
            $events += New-SedaObject @{
                Level = $level
                Timestamp = $timestamp
                Area = $Area
                Source = $file
                LineNumber = $lineNumber
                ErrorCode = $code
                Message = $trimmed.Substring(0, [Math]::Min(600, $trimmed.Length))
                RawLine = $trimmed
                Path = $path
            }
        }
    }
    return $events
}

function Get-SedaWuEtlEvents {
    param(
        [string[]]$Paths,
        [string]$GeneratedLogPath,
        [string]$GeneratedOutputDirectory
    )
    $events = @()
    $files = @($Paths | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Select-Object -First 50)
    $generatedLog = ''
    if ($files.Count -gt 0) {
        $generatedLog = Convert-SedaWindowsUpdateEtlsToLog -Paths $files -OutputDirectory $GeneratedOutputDirectory
    } elseif ($GeneratedLogPath -and (Test-Path -LiteralPath $GeneratedLogPath -PathType Leaf)) {
        $generatedLog = $GeneratedLogPath
    }
    if ($generatedLog) {
        $events = @(Get-SedaWindowsUpdateLogEvents -Path $generatedLog)
        $errorCount = @($events | Where-Object { $_.Level -in @('Critical','Error') }).Count
        $warningCount = @($events | Where-Object { $_.Level -eq 'Warning' }).Count
        $statusPrefix = if ($files.Count -gt 0) { 'Generated and parsed WindowsUpdate.log with Get-WindowsUpdateLog' } else { 'Parsed existing WindowsUpdate.log because no ETL file was available' }
        return New-SedaObject @{
            Status = "$statusPrefix - $errorCount errors, $warningCount warnings, $(@($events).Count) matched events."
            Events = $events
            ErrorCount = $errorCount
            WarningCount = $warningCount
            FilesScanned = $files.Count
            Decoder = 'Get-WindowsUpdateLog'
            GeneratedLogPath = $generatedLog
        }
    }
    if ($files.Count -eq 0) {
        return New-SedaObject @{ Status = 'No Windows Update ETL file or generated WindowsUpdate.log found.'; Events = @(); ErrorCount = 0; WarningCount = 0; FilesScanned = 0; Decoder = ''; GeneratedLogPath = '' }
    }

    $errKeywords = [regex]'(?i)\b(error|failed|failure|fatal|0x[0-9a-f]{8}|hr=|hresult|denied|timeout)\b'
    $warnKeywords = [regex]'(?i)\b(warn|warning|retry|defer|blocked|pending|paused)\b'
    $decoder = 'Get-WinEvent'
    $filesScanned = 0

    foreach ($path in $files) {
        $fname = [System.IO.Path]::GetFileName($path)
        $decoded = $false
        try {
            $rawEvents = Get-WinEvent -Path $path -MaxEvents 500 -ErrorAction Stop
            foreach ($ev in @($rawEvents)) {
                $message = ([string]$ev.Message).Trim()
                $level = [string]$ev.LevelDisplayName
                if ([string]::IsNullOrWhiteSpace($level)) {
                    $level = switch ([int]$ev.Level) { 1 { 'Critical' } 2 { 'Error' } 3 { 'Warning' } 4 { 'Information' } 5 { 'Verbose' } default { 'Information' } }
                }
                if ($level -in @('Information','Verbose')) {
                    if ($message -match $errKeywords) { $level = 'Error' }
                    elseif ($message -match $warnKeywords) { $level = 'Warning' }
                }
                $code = ''
                if ($message -match '(?i)(0x[0-9a-f]{8})') { $code = $Matches[1] }
                $events += New-SedaObject @{ Level=$level; Timestamp=('{0:yyyy-MM-dd HH:mm:ss}' -f $ev.TimeCreated); EventId=[string]$ev.Id; Source=[string]$ev.ProviderName; Message=$message; ErrorCode=$code; EtlFile=$fname }
            }
            $decoded = $true
            $filesScanned++
        } catch {
            $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('seda_etl_' + [guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $tmp -Force | Out-Null
            $xmlOut = Join-Path $tmp 'out.xml'
            $stdout = Join-Path $tmp 'tracerpt.out'
            $stderr = Join-Path $tmp 'tracerpt.err'
            try {
                $process = Start-Process -FilePath 'tracerpt.exe' -ArgumentList @($path, '-of', 'XML', '-o', $xmlOut, '-y') -NoNewWindow -Wait -PassThru -RedirectStandardOutput $stdout -RedirectStandardError $stderr -ErrorAction Stop
                if ($process.ExitCode -ge 0 -and (Test-Path -LiteralPath $xmlOut -PathType Leaf)) {
                    [xml]$xml = Get-SedaTextContent -Path $xmlOut
                    foreach ($node in @($xml.SelectNodes('//Event'))) {
                        $system = $node.System
                        if (-not $system) { continue }
                        $levelValue = [int]([string]$system.Level)
                        $level = switch ($levelValue) { 1 { 'Critical' } 2 { 'Error' } 3 { 'Warning' } 4 { 'Information' } 5 { 'Verbose' } default { 'Information' } }
                        $dataText = ''
                        if ($node.EventData) {
                            $dataText = (($node.EventData.ChildNodes | ForEach-Object { [string]$_.InnerText }) -join ' ').Trim()
                        }
                        if ($level -in @('Information','Verbose')) {
                            if ($dataText -match $errKeywords) { $level = 'Error' }
                            elseif ($dataText -match $warnKeywords) { $level = 'Warning' }
                        }
                        $code = ''
                        if ($dataText -match '(?i)(0x[0-9a-f]{8})') { $code = $Matches[1] }
                        $provider = ''
                        if ($system.Provider) { $provider = [string]$system.Provider.Name; if (-not $provider) { $provider = [string]$system.Provider.Guid } }
                        $timestamp = ''
                        if ($system.TimeCreated) { $timestamp = [string]$system.TimeCreated.SystemTime }
                        if ($timestamp -match 'T') { $timestamp = $timestamp.Substring(0, [Math]::Min(19, $timestamp.Length)).Replace('T',' ') }
                        $events += New-SedaObject @{ Level=$level; Timestamp=$timestamp; EventId=[string]$system.EventID; Source=$provider; Message=$dataText; ErrorCode=$code; EtlFile=$fname }
                    }
                    $decoder = 'tracerpt'
                    $decoded = $true
                    $filesScanned++
                }
            } catch {
            } finally {
                Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
        if (-not $decoded) {
            $events += New-SedaObject @{ Level='Warning'; Timestamp=''; EventId=''; Source='ETL decoder'; Message='Unable to decode ETL file with Get-WinEvent or tracerpt.exe.'; ErrorCode=''; EtlFile=$fname }
        }
    }

    $errorCount = @($events | Where-Object { $_.Level -in @('Critical','Error') }).Count
    $warningCount = @($events | Where-Object { $_.Level -eq 'Warning' }).Count
    return New-SedaObject @{
        Status = "Scanned $filesScanned/$($files.Count) ETL files ($decoder) - $errorCount errors, $warningCount warnings, $(@($events).Count) total events."
        Events = $events
        ErrorCount = $errorCount
        WarningCount = $warningCount
        FilesScanned = $filesScanned
        Decoder = $decoder
        GeneratedLogPath = ''
    }
}

function Get-SedaInstalledApps {
    param([string[]]$RegFiles)
    $apps = @()
    foreach ($path in @($RegFiles)) {
        if (-not $path) { continue }
        $keys = ConvertFrom-SedaRegFile -Path $path
        $arch = if ($path.ToLowerInvariant() -like '*wow6432*') { 'x86' } else { 'x64' }
        foreach ($key in $keys.Keys) {
            if ($key.ToLowerInvariant() -notlike '*\uninstall\*') { continue }
            $values = $keys[$key]
            if (-not $values.DisplayName) { continue }
            $installDate = [string]$values.InstallDate
            if ($installDate -match '^(\d{4})(\d{2})(\d{2})') { $installDate = "$($Matches[1])-$($Matches[2])-$($Matches[3])" }
            $apps += New-SedaObject @{
                Name = [string]$values.DisplayName
                Version = [string]$values.DisplayVersion
                Publisher = [string]$values.Publisher
                InstallDate = $installDate
                InstallLocation = [string]$values.InstallLocation
                Arch = $arch
            }
        }
    }
    return @($apps | Sort-Object Name)
}

function Get-SedaDrivers {
    param([string]$Path)
    $drivers = @()
    if (-not $Path) { return $drivers }
    $content = Get-SedaTextContent -Path $Path
    $current = [ordered]@{}
    $flush = {
        if ($current.original_name -or $current.published_name) {
            $script:__sedaDrivers += New-SedaObject @{
                PublishedName = [string]$current.published_name
                OriginalName = [string]$current.original_name
                Provider = [string]$current.provider
                ClassName = [string]$current.class_name
                DriverVersion = [string]$current.driver_version
                Signer = [string]$current.signer
            }
        }
    }
    $script:__sedaDrivers = @()
    foreach ($raw in ($content -split "`r?`n")) {
        $line = $raw.Trim()
        if (-not $line) {
            & $flush
            $current = [ordered]@{}
            continue
        }
        if (-not $line.Contains(':')) { continue }
        $key, $value = $line -split ':', 2
        $keyLower = $key.Trim().ToLowerInvariant()
        $value = $value.Trim()
        if ($keyLower -like '*published name*' -or $keyLower -like '*nom publie*' -or $keyLower -like '*nom publié*') { $current.published_name = $value }
        elseif ($keyLower -like '*original name*' -or $keyLower -like "*nom d'origine*") { $current.original_name = $value }
        elseif ($keyLower -like '*provider name*' -or $keyLower -like '*nom du fournisseur*') { $current.provider = $value }
        elseif ($keyLower -like '*class name*' -or $keyLower -like '*nom de la classe*') { $current.class_name = $value }
        elseif ($keyLower -like '*driver version*' -or $keyLower -like '*version du pilote*') { $current.driver_version = $value }
        elseif ($keyLower -like '*signer name*' -or $keyLower -like '*nom du signataire*') { $current.signer = $value }
    }
    & $flush
    $drivers = $script:__sedaDrivers
    Remove-Variable -Name __sedaDrivers -Scope Script -ErrorAction SilentlyContinue
    return @($drivers | Sort-Object Provider, OriginalName)
}

function Get-SedaWifiProfiles {
    param([string]$Path)
    $profiles = @()
    if (-not $Path) { return $profiles }
    $content = Get-SedaTextContent -Path $Path
    $currentType = ''
    foreach ($raw in ($content -split "`r?`n")) {
        $line = $raw.Trim()
        $lower = $line.ToLowerInvariant()
        if ($lower -like '*group policy*' -or $lower -like '*strategie de groupe*' -or $lower -like '*stratégie de groupe*') { $currentType = 'GPO' }
        elseif ($lower -like '*user profile*' -or $lower -like '*profils utilisateurs*') { $currentType = 'User' }
        if ($line.Contains(':')) {
            $key, $value = $line -split ':', 2
            $key = $key.Trim().ToLowerInvariant()
            $value = $value.Trim()
            if ($value -and ($key -like '*ssidname*' -or $key -like '*ssid name*' -or $key -like '*nom du profil*' -or $key -like '*profile name*')) {
                if (-not (@($profiles | Select-Object -ExpandProperty Ssid) -contains $value)) {
                    $profiles += New-SedaObject @{ Ssid = $value; ProfileType = $currentType }
                }
            }
        }
    }
    return $profiles
}

function Get-SedaExtraSummary {
    param([hashtable]$Inventory)
    $summary = [ordered]@{}
    $ipconfigFile = Find-SedaInventoryFile -Inventory $Inventory -Keyword 'ipconfig_exe_all'
    if (-not $ipconfigFile) { $ipconfigFile = Find-SedaInventoryFile -Inventory $Inventory -Keyword 'ipconfig' }
    if ($ipconfigFile) {
        $text = Get-SedaTextContent -Path $ipconfigFile
        foreach ($line in ($text -split "`r?`n")) {
            if (-not $summary.Hostname -and $line -match '(?:Host\s+Name|Hostname|Nom\s+de\s+l.{0,3}h.{1,5}te)[^:]*:\s*(\S+)') { $summary.Hostname = $Matches[1].Trim() }
            if (-not $summary.IPAddress -and $line -match '(?:Adresse\s+IPv4|IPv4\s+Address|IPv4-Adresse)[^:]*:\s*((?:\d{1,3}\.){3}\d{1,3})') {
                $candidate = $Matches[1].Trim()
                if (-not $candidate.StartsWith('127.') -and -not $candidate.StartsWith('169.254.')) { $summary.IPAddress = $candidate }
            }
        }
    }
    $proxyFile = Find-SedaInventoryFile -Inventory $Inventory -Keyword 'winhttp_show_proxy'
    if (-not $proxyFile) { $proxyFile = Find-SedaInventoryFile -Inventory $Inventory -Keyword 'winhttp' }
    if ($proxyFile) {
        $text = Get-SedaTextContent -Path $proxyFile
        if ($text -match '(?:Acc[eè]s\s+direct|Direct\s+access)') { $summary.Proxy = 'Direct (no proxy)' }
        if ($text -match '(?:Proxy\s+Server|Serveurs?\s+proxy)[^:]*:\s*(\S+)') { $summary.Proxy = $Matches[1].Trim() }
    }
    if (-not $summary.Proxy) { $summary.Proxy = 'Unknown' }

    $logonUi = Find-SedaInventoryFile -Inventory $Inventory -Keyword 'logonui'
    if ($logonUi) {
        $keys = ConvertFrom-SedaRegFile -Path $logonUi
        foreach ($key in $keys.Keys) {
            if ($keys[$key].LastLoggedOnDisplayName) { $summary.LastUser = $keys[$key].LastLoggedOnDisplayName; break }
            if ($keys[$key].LastLoggedOnSAMUser) { $summary.LastUser = $keys[$key].LastLoggedOnSAMUser; break }
        }
    }
    if (-not $summary.LastUser) { $summary.LastUser = 'Unknown' }

    $imeReg = Find-SedaInventoryFile -Inventory $Inventory -Keyword 'intunemanagementextension'
    if ($imeReg) {
        $text = Get-SedaTextContent -Path $imeReg
        if ($text -match 'AgentVersion:([\d.]+)') { $summary.ImeVersion = $Matches[1] }
        elseif ($text -match '"Version"\s*=\s*"([\d.]+)"') { $summary.ImeVersion = $Matches[1] }
        if ($text -match '"LastSyncFeatureList"\s*=\s*"([^"]+)"') { $summary.LastSync = $Matches[1] }
    }
    if (-not $summary.ImeVersion) { $summary.ImeVersion = 'Unknown' }

    $msinfo = Find-SedaInventoryFile -Inventory $Inventory -Keyword 'msinfo32'
    if ($msinfo) {
        $text = Get-SedaTextContent -Path $msinfo
        if ($text -match '(?:OS\s+Name|Nom\s+du\s+(?:systeme|syst[eè]me|SE))[^:]*:\s*(.+)') { $summary.OSName = $Matches[1].Trim() }
        if ($text -match '(?:OS\s+Version|Version\s+du\s+(?:systeme|syst[eè]me|SE))[^:]*:\s*(.+)') { $summary.OSVersion = ($Matches[1].Trim() -replace '\s+N/A\s+', ' ') }
    }
    return $summary
}

function Get-SedaComplianceSummary {
    param(
        [object]$DsReg,
        [object]$Enrollments,
        [object]$Firewall,
        [object]$Results
    )
    $statuses = @()
    $sections = $DsReg.Sections
    $dev = $sections['Device State']
    $detail = $sections['Device Details']
    $sso = $sections['SSO State']
    $user = $sections['User State']
    $aadJoined = if ($dev) { [string]$dev.AzureAdJoined } else { 'UNKNOWN' }
    $workplaceJoined = if ($dev.WorkplaceJoined) { [string]$dev.WorkplaceJoined } elseif ($user.WorkplaceJoined) { [string]$user.WorkplaceJoined } else { '' }
    $isByod = ($aadJoined.ToUpperInvariant() -ne 'YES' -and $workplaceJoined.ToUpperInvariant() -eq 'YES')

    if ($isByod) {
        $statuses += New-SedaObject @{ Area = 'Device Join'; Status = 'COMPLIANT'; Details = 'WorkplaceJoined=YES (BYOD/Personal device)'; SourceFile = 'DSRegCmd' }
    } else {
        $statuses += New-SedaObject @{ Area = 'AAD / PRT'; Status = if ($aadJoined.ToUpperInvariant() -eq 'YES') { 'COMPLIANT' } elseif ($aadJoined.ToUpperInvariant() -eq 'NO') { 'NON_COMPLIANT' } else { 'UNKNOWN' }; Details = "AzureAdJoined=$aadJoined"; SourceFile = 'DSRegCmd' }
        $prt = if ($sso) { [string]$sso.AzureAdPrt } else { 'UNKNOWN' }
        $statuses += New-SedaObject @{ Area = 'AAD / PRT'; Status = if ($prt.ToUpperInvariant() -eq 'YES') { 'COMPLIANT' } elseif ($prt.ToUpperInvariant() -eq 'NO') { 'NON_COMPLIANT' } else { 'UNKNOWN' }; Details = "AzureAdPrt=$prt"; SourceFile = 'DSRegCmd' }
        if ($sso.WamDefaultSet) {
            $statuses += New-SedaObject @{ Area = 'AAD / PRT'; Status = if ([string]$sso.WamDefaultSet -like '*ERROR*') { 'NON_COMPLIANT' } else { 'COMPLIANT' }; Details = "WamDefaultSet=$($sso.WamDefaultSet)"; SourceFile = 'DSRegCmd' }
        }
        if ($detail.DeviceAuthStatus) {
            $statuses += New-SedaObject @{ Area = 'MDM Enrollment'; Status = if ([string]$detail.DeviceAuthStatus -like '*SUCCESS*') { 'COMPLIANT' } else { 'NON_COMPLIANT' }; Details = "DeviceAuthStatus=$($detail.DeviceAuthStatus)"; SourceFile = 'DSRegCmd' }
        }
    }
    if ($detail.TpmProtected) {
        $statuses += New-SedaObject @{ Area = 'TPM'; Status = if ([string]$detail.TpmProtected -like '*YES*') { 'COMPLIANT' } else { 'NON_COMPLIANT' }; Details = "TpmProtected=$($detail.TpmProtected)"; SourceFile = 'DSRegCmd' }
    }
    if ($sso.NgcSet) {
        $statuses += New-SedaObject @{ Area = 'Hello for Business'; Status = if ([string]$sso.NgcSet -like '*YES*') { 'COMPLIANT' } else { 'NON_COMPLIANT' }; Details = "NgcSet=$($sso.NgcSet)"; SourceFile = 'DSRegCmd' }
    }
    if ($Enrollments.Enrollments.Count -gt 0) {
        $active = @($Enrollments.Enrollments | Where-Object { $_.State -like '*active*' }).Count
        $statuses += New-SedaObject @{ Area = 'MDM Enrollment'; Status = if ($active -gt 0) { 'COMPLIANT' } else { 'NON_COMPLIANT' }; Details = "$active/$($Enrollments.Enrollments.Count) active enrollments"; SourceFile = 'Enrollments.reg' }
    }
    foreach ($profile in $Firewall.Profiles.Keys) {
        $state = [string]$Firewall.Profiles[$profile].State
        if ($state.ToUpperInvariant().Contains('ON')) {
            $statuses += New-SedaObject @{ Area = 'Firewall'; Status = 'COMPLIANT'; Details = "Profile $profile`: active"; SourceFile = 'netsh firewall' }
        } elseif ($state.ToUpperInvariant().Contains('OFF')) {
            $statuses += New-SedaObject @{ Area = 'Firewall'; Status = 'NON_COMPLIANT'; Details = "Profile $profile`: disabled"; SourceFile = 'netsh firewall' }
        }
    }
    foreach ($error in @($Results.Errors)) {
        $statuses += New-SedaObject @{ Area = 'Collection'; Status = 'UNKNOWN'; Details = "Collection failed: $($error.Name) - $($error.Status)"; SourceFile = 'results.xml' }
    }

    $unique = @()
    $seen = @{}
    foreach ($status in $statuses) {
        $key = "$($status.Area)|$($status.Status)|$($status.Details)"
        if (-not $seen[$key]) { $seen[$key] = $true; $unique += $status }
    }
    $compliant = @($unique | Where-Object { $_.Status -eq 'COMPLIANT' }).Count
    $nonCompliant = @($unique | Where-Object { $_.Status -eq 'NON_COMPLIANT' }).Count
    $pending = @($unique | Where-Object { $_.Status -eq 'PENDING' }).Count
    $unknown = @($unique | Where-Object { $_.Status -eq 'UNKNOWN' }).Count
    $overall = if ($nonCompliant -gt 0) { 'NON_COMPLIANT' } elseif ($pending -gt 0 -and $compliant -eq 0) { 'PENDING' } elseif ($compliant -gt 0) { 'COMPLIANT' } else { 'UNKNOWN' }

    return New-SedaObject @{ OverallStatus = $overall; CompliantCount = $compliant; NonCompliantCount = $nonCompliant; PendingCount = $pending; UnknownCount = $unknown; PolicyStatuses = $unique }
}

function Get-SedaWin11CompatibilityIndicators {
    param([string]$Path)
    $indicators = @()
    $blocking = @()
    $status = 'NoIndicatorsPath'
    if (-not $Path) { return New-SedaObject @{ Status = $status; Indicators = $indicators; BlockingIndicators = $blocking } }
    $keys = ConvertFrom-SedaRegFile -Path $Path
    $root = 'hkey_local_machine\software\microsoft\windows nt\currentversion\appcompatflags\targetversionupgradeexperienceindicators'
    foreach ($key in $keys.Keys) {
        $lower = $key.ToLowerInvariant()
        if ($lower -eq $root -or -not $lower.StartsWith($root + '\')) { continue }
        $rel = $key.Substring($root.Length).Trim('\')
        if (-not $rel) { continue }
        $target = ($rel -split '\\', 2)[0]
        $values = $keys[$key]
        $reasonParts = @()
        foreach ($pair in @(
            @{ Label = 'UpEx'; Name = 'UpEx' },
            @{ Label = 'GatedBlockId'; Name = 'GatedBlockId' },
            @{ Label = 'RedReason'; Name = 'RedReason' },
            @{ Label = 'SysReqIssue'; Name = 'SysReqIssue' }
        )) {
            $value = [string]$values[$pair.Name]
            $clean = $value.Trim().Trim('"')
            if ($clean -and $clean.ToLowerInvariant() -notin @('none', 'n/a', 'notapplicable', 'not applicable')) {
                $reasonParts += "$($pair.Label)=$clean"
            }
        }
        $reason = if ($reasonParts) { $reasonParts -join '; ' } else { 'No blocking indicator fields' }
        $isBlocking = ($reason -match 'UpEx=.*(Red|Blocked|Hold)' -or $values.GatedBlockId -or $values.RedReason -or $values.SysReqIssue)
        $indicator = New-SedaObject @{ TargetVersion = $target; UpEx = [string]$values.UpEx; GatedBlockId = [string]$values.GatedBlockId; RedReason = [string]$values.RedReason; SysReqIssue = [string]$values.SysReqIssue; ReasonText = $reason; IsBlocking = [bool]$isBlocking; SourceFile = $Path }
        $indicators += $indicator
        if ($indicator.IsBlocking) { $blocking += $indicator }
    }
    if ($blocking.Count -gt 0) { $status = 'BlockingConditionDetected' }
    elseif ($indicators.Count -gt 0) { $status = 'NoBlockingConditionDetected' }
    else { $status = 'NoIndicators' }
    return New-SedaObject @{ Status = $status; Indicators = $indicators; BlockingIndicators = $blocking }
}

function Get-SedaHardwareReadiness {
    param([string]$Path)
    $checks = @()
    $notes = @()
    $rawLogging = ''
    $status = 'NotAvailable'
    $returnCode = $null
    $returnResult = 'No HardwareReadiness.ps1 data'
    $returnReason = ''
    if (-not $Path -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return New-SedaObject @{ Parsed=$false; SourceFile=$Path; Status=$status; ReturnCode=$returnCode; ReturnResult=$returnResult; ReturnReason=$returnReason; Checks=$checks; Notes=$notes; RawLogging=$rawLogging }
    }
    try {
        $text = Get-SedaTextContent -Path $Path
        $jsonLine = @($text -split "`r?`n" | Where-Object { $_.Trim().StartsWith('{') -and $_ -match 'returnCode' } | Select-Object -Last 1)[0]
        if (-not $jsonLine) { $jsonLine = $text.Trim() }
        $obj = $jsonLine | ConvertFrom-Json -ErrorAction Stop
        $returnCode = $obj.returnCode
        $returnResult = [string]$obj.returnResult
        $returnReason = ([string]$obj.returnReason).Trim().Trim(',')
        $rawLogging = [string]$obj.logging
        $display = @{
            Storage = @('Disk >= 64 GB','OS disk total size')
            Memory = @('RAM >= 4 GB','Physical memory')
            TPM = @('TPM 2.0','Trusted Platform Module version')
            Processor = @('CPU Compatible','Processor family and speed')
            SecureBoot = @('Secure Boot / UEFI','UEFI Secure Boot capability')
            'i7-7820hq CPU' = @('i7-7820HQ override','Surface Studio 2 / Precision 5520 exception')
        }
        foreach ($match in [regex]::Matches($rawLogging, '([A-Za-z0-9_ ()-]+?):\s*(.*?)\.\s*(PASS|FAIL|UNDETERMINED|FAILED TO RUN)\s*;?', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
            $key = $match.Groups[1].Value.Trim()
            $detail = $match.Groups[2].Value.Trim()
            $verdict = $match.Groups[3].Value.ToUpperInvariant()
            $labels = if ($display.ContainsKey($key)) { $display[$key] } else { @($key,'') }
            $value = $detail
            $valueMatch = [regex]::Match($detail, '=([^.;{]+)')
            if ($valueMatch.Success) { $value = $valueMatch.Groups[1].Value.Trim() }
            $state = switch ($verdict) {
                'PASS' { 'PASS' }
                'FAIL' { 'FAIL' }
                'UNDETERMINED' { 'UNDETERMINED' }
                default { 'UNKNOWN' }
            }
            $checks += New-SedaObject ([ordered]@{
                Requirement = $labels[0]
                Status = $state
                Value = $value
                Detail = if ($detail.Length -gt 240) { $detail.Substring(0,240) } else { $detail }
                Source = 'HardwareReadiness.ps1'
            })
        }
        if ($checks.Count -eq 0) {
            $checks += New-SedaObject ([ordered]@{ Requirement='Microsoft Check'; Status='UNKNOWN'; Value=$returnResult; Detail=$rawLogging; Source='HardwareReadiness.ps1' })
        }
        if ($returnCode -eq 0) {
            $status = 'CAPABLE'
            $notes += 'Device meets Windows 11 hardware requirements.'
        } elseif ($returnCode -eq 1) {
            $status = 'NOT_CAPABLE'
            $notes += "Blockers: $(if ($returnReason) { $returnReason } else { 'see failed checks' })"
            $notes += 'Actions: enable TPM 2.0, Secure Boot/UEFI, provide at least 4 GB RAM, 64 GB OS disk, and a supported CPU as required.'
        } elseif ($returnCode -eq -1) {
            $status = 'UNDETERMINED'
            $notes += 'Some hardware readiness checks could not be completed.'
        } else {
            $status = 'FAILED_TO_RUN'
            $notes += 'HardwareReadiness.ps1 did not return a usable capable/not capable result.'
        }
        return New-SedaObject @{ Parsed=$true; SourceFile=$Path; Status=$status; ReturnCode=$returnCode; ReturnResult=$returnResult; ReturnReason=$returnReason; Checks=$checks; Notes=$notes; RawLogging=$rawLogging }
    } catch {
        Write-SedaLog -Level WARN -Message "Unable to parse HardwareReadiness output: $Path" -Exception $_.Exception
        return New-SedaObject @{ Parsed=$false; SourceFile=$Path; Status='ParseFailed'; ReturnCode=$returnCode; ReturnResult='Parse failed'; ReturnReason=$_.Exception.Message; Checks=$checks; Notes=@($_.Exception.Message); RawLogging=$rawLogging }
    }
}

function ConvertFrom-SedaHtmlText {
    param([string]$Html)
    if ($null -eq $Html) { return '' }
    $text = [regex]::Replace($Html, '<[^>]+>', '')
    $text = [System.Net.WebUtility]::HtmlDecode($text)
    return (($text -split '\s+') -join ' ').Trim()
}

function Get-SedaHtmlTableRows {
    param([string]$HtmlSection)
    $rows = @()
    foreach ($rowMatch in [regex]::Matches($HtmlSection, '<tr[^>]*>(.*?)</tr>', [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
        $cells = @()
        foreach ($cellMatch in [regex]::Matches($rowMatch.Groups[1].Value, '<t[dh][^>]*>(.*?)</t[dh]>', [System.Text.RegularExpressions.RegexOptions]::Singleline)) {
            $cells += ConvertFrom-SedaHtmlText $cellMatch.Groups[1].Value
        }
        if (@($cells | Where-Object { $_ }).Count -gt 0) { $rows += ,$cells }
    }
    return $rows
}

function Get-SedaMdmDiagSection {
    param(
        [string[]]$Sections,
        [string]$Keyword
    )
    foreach ($section in $Sections) {
        $titleMatch = [regex]::Match($section, 'SectionTitle[^>]*>(.*?)</(?:span|div|a)', [System.Text.RegularExpressions.RegexOptions]::Singleline)
        if ($titleMatch.Success -and (ConvertFrom-SedaHtmlText $titleMatch.Groups[1].Value).ToLowerInvariant().Contains($Keyword.ToLowerInvariant())) {
            return $section
        }
    }
    return $null
}

function Expand-SedaCabFile {
    param(
        [string]$CabPath,
        [string]$DestinationPath
    )
    $result = New-SedaObject @{ Success = $false; Status = 'No CAB file selected'; ExtractPath = ''; Files = @() }
    if (-not $CabPath -or -not (Test-Path -LiteralPath $CabPath -PathType Leaf)) {
        $result.Status = "CAB file not found: $CabPath"
        return $result
    }
    if (-not $DestinationPath) {
        $DestinationPath = Join-Path ([System.IO.Path]::GetTempPath()) ('seda_cab_' + [guid]::NewGuid().ToString('N'))
    }
    New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
    $expand = Get-Command expand.exe -ErrorAction SilentlyContinue
    if (-not $expand) {
        $result.Status = 'expand.exe not available'
        return $result
    }
    try {
        $proc = Start-Process -FilePath $expand.Source -ArgumentList @($CabPath, '-F:*', $DestinationPath) -WindowStyle Hidden -Wait -PassThru -ErrorAction Stop
        $files = @(Get-ChildItem -LiteralPath $DestinationPath -File -Recurse -Force -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName)
        $result.Success = ($proc.ExitCode -eq 0 -or $files.Count -gt 0)
        $result.ExtractPath = $DestinationPath
        $result.Files = $files
        $result.Status = if ($result.Success) { "Extracted $($files.Count) file(s) from $([System.IO.Path]::GetFileName($CabPath))" } else { "expand.exe failed with exit code $($proc.ExitCode)" }
    } catch {
        $result.Status = "CAB extraction failed: $($_.Exception.Message)"
    }
    return $result
}

function Get-SedaMdmDiagReport {
    param(
        [string[]]$CabFiles,
        [switch]$SkipExtraction
    )
    $report = New-SedaObject @{
        Parsed = $false; HtmlFile = ''; CabStatus = 'No CAB file found'; DeviceInfo = @{}; ConnectionInfo = @{}; AccountInfo = @{};
        Certificates = @(); ConfigSources = @(); ManagedPolicies = @(); Laps = @{}; BlockedGps = @(); UnmanagedAreas = @(); Issues = @()
    }
    if (-not $CabFiles -or $CabFiles.Count -eq 0) { return $report }
    if ($SkipExtraction) {
        $report.CabStatus = 'CAB extraction skipped'
        return $report
    }
    $cabResult = Expand-SedaCabFile -CabPath $CabFiles[0]
    $report.CabStatus = $cabResult.Status
    if (-not $cabResult.Success) { return $report }
    $htmlFile = @($cabResult.Files | Where-Object { [System.IO.Path]::GetFileName($_).ToLowerInvariant() -eq 'mdmdiaghtmlreport.html' } | Select-Object -First 1)[0]
    if (-not $htmlFile) {
        $report.CabStatus += ' - MDMDiagHTMLReport.html not found'
        return $report
    }
    $html = Get-SedaTextContent -Path $htmlFile
    if (-not $html) { return $report }
    $sections = @([regex]::Matches($html, '<section[^>]*>(.*?)</section>', [System.Text.RegularExpressions.RegexOptions]::Singleline) | ForEach-Object { $_.Groups[1].Value })
    $sectionMap = @{
        Device = Get-SedaMdmDiagSection -Sections $sections -Keyword 'Device Info'
        Connection = Get-SedaMdmDiagSection -Sections $sections -Keyword 'Connection Info'
        Account = Get-SedaMdmDiagSection -Sections $sections -Keyword 'Device Management Account'
        Certificates = Get-SedaMdmDiagSection -Sections $sections -Keyword 'Certificates'
        Sources = Get-SedaMdmDiagSection -Sections $sections -Keyword 'Enrolled configuration'
        Policies = Get-SedaMdmDiagSection -Sections $sections -Keyword 'Managed policies'
        Laps = Get-SedaMdmDiagSection -Sections $sections -Keyword 'LAPS'
        BlockedGps = Get-SedaMdmDiagSection -Sections $sections -Keyword 'Blocked Group'
        Unmanaged = Get-SedaMdmDiagSection -Sections $sections -Keyword 'Unmanaged'
    }
    foreach ($pair in @(@('Device','DeviceInfo'), @('Connection','ConnectionInfo'), @('Account','AccountInfo'), @('Laps','Laps'))) {
        $sec = $sectionMap[$pair[0]]
        if (-not $sec) { continue }
        $target = $pair[1]
        foreach ($row in Get-SedaHtmlTableRows -HtmlSection $sec) {
            if ($row.Count -ge 2 -and $row[0]) { $report.$target[$row[0]] = $row[1] }
        }
    }
    if ($sectionMap.Certificates) {
        foreach ($row in (Get-SedaHtmlTableRows -HtmlSection $sectionMap.Certificates | Select-Object -Skip 1)) {
            if ($row.Count -ge 2 -and $row[0] -and $row[0] -ne 'Issued to') { $report.Certificates += New-SedaObject @{ IssuedTo = $row[0]; IssuedBy = $row[1] } }
        }
    }
    if ($sectionMap.Sources) {
        foreach ($row in Get-SedaHtmlTableRows -HtmlSection $sectionMap.Sources) {
            if ($row.Count -ge 2 -and $row[0] -and $row[0] -ne 'Configuration source') { $report.ConfigSources += New-SedaObject @{ Source = $row[0]; Id = $row[1] } }
        }
    }
    if ($sectionMap.Policies) {
        foreach ($row in Get-SedaHtmlTableRows -HtmlSection $sectionMap.Policies) {
            if ($row.Count -ge 2 -and $row[0] -and $row[0] -ne 'Area') {
                $report.ManagedPolicies += New-SedaObject @{ Area = $row[0]; Policy = $row[1]; Value = if ($row.Count -gt 2) { $row[2] } else { '' } }
            }
        }
    }
    if ($sectionMap.BlockedGps) {
        foreach ($row in Get-SedaHtmlTableRows -HtmlSection $sectionMap.BlockedGps) {
            if ($row.Count -ge 2 -and $row[0] -and $row[0] -notin @('Blocked GP Entity','Target')) {
                $report.BlockedGps += New-SedaObject @{ Path = $row[0]; Name = $row[1] }
            }
        }
    }
    if ($sectionMap.Unmanaged) {
        foreach ($row in Get-SedaHtmlTableRows -HtmlSection $sectionMap.Unmanaged) {
            if ($row.Count -ge 2 -and $row[0] -and $row[0] -ne 'Area') { $report.UnmanagedAreas += New-SedaObject @{ Area = $row[0]; Policies = $row[1] } }
        }
    }
    foreach ($policy in @($report.ManagedPolicies)) {
        if ($policy.Area -eq 'DeviceGuard' -and $policy.Policy -eq 'EnableVirtualizationBasedSecurity' -and $policy.Value -eq '0') {
            $report.Issues += New-SedaObject @{ Severity='WARNING'; Area='DeviceGuard'; Title='Virtualization Based Security (VBS) disabled'; Detail='EnableVirtualizationBasedSecurity = 0'; Recommendation='Enable VBS in Intune Endpoint Security.' }
        }
        if ($policy.Area -eq 'ControlPolicyConflict' -and $policy.Policy -eq 'MDMWinsOverGP' -and $policy.Value -eq '0') {
            $report.Issues += New-SedaObject @{ Severity='WARNING'; Area='ControlPolicyConflict'; Title='Group Policy wins over MDM (MDMWinsOverGP = 0)'; Detail='GPO can override MDM policy.'; Recommendation='Set MDMWinsOverGP = 1 or audit conflicting GPO settings.' }
        }
        if ($policy.Area -eq 'DeviceHealthMonitoring' -and $policy.Policy -eq 'AllowDeviceHealthMonitoring' -and $policy.Value -eq '0') {
            $report.Issues += New-SedaObject @{ Severity='INFO'; Area='DeviceHealthMonitoring'; Title='Device Health Monitoring disabled'; Detail='AllowDeviceHealthMonitoring = 0'; Recommendation='Enable Intune Device Health Monitoring where expected.' }
        }
    }
    if ($report.BlockedGps.Count -gt 0) {
        $report.Issues += New-SedaObject @{ Severity='WARNING'; Area='GroupPolicy'; Title="$($report.BlockedGps.Count) Group Policy setting(s) blocked by MDM"; Detail='GPO settings conflict with MDM policy and are blocked.'; Recommendation='Review blocked GPO list and migrate settings to Intune.' }
    }
    if ($report.Certificates.Count -eq 0) {
        $managedBy = [string]$report.ConnectionInfo['Managed by']
        $lastSync = [string]($report.ConnectionInfo['Last sync'] + $report.ConnectionInfo['Last MDM sync'])
        $report.Issues += if ($managedBy -or $lastSync) {
            New-SedaObject @{ Severity='WARNING'; Area='Certificates'; Title='No MDM device certificates visible in this report'; Detail='No MDM certificate rows found; this may be expected for BYOD/Workplace Join.'; Recommendation='Verify enrollment in Intune portal.' }
        } else {
            New-SedaObject @{ Severity='ERROR'; Area='Certificates'; Title='No Intune MDM certificates found'; Detail='No active management certificate or connection evidence detected.'; Recommendation='Re-enroll the device or renew the MDM certificate.' }
        }
    }
    $report.Parsed = $true
    $report.HtmlFile = $htmlFile
    return $report
}

function Get-SedaBatteryReport {
    param([string]$Path)
    $info = [ordered]@{ Parsed=$false; ComputerName=''; SystemProduct=''; OSBuild=''; BIOS=''; DesignCapacity=''; FullChargeCapacity=''; CycleCount=''; BatteryName=''; BatterySerial=''; BatteryChemistry=''; LastFullCharge=''; HealthPct=0 }
    if (-not $Path) { return New-SedaObject $info }
    $html = Get-SedaTextContent -Path $Path
    if (-not $html) { return New-SedaObject $info }
    $pairs = @{}
    foreach ($row in Get-SedaHtmlTableRows -HtmlSection $html) {
        if ($row.Count -ge 2 -and $row[0] -and -not $pairs.ContainsKey($row[0].ToLowerInvariant())) { $pairs[$row[0].ToLowerInvariant()] = $row[1] }
    }
    function get-battery-value([string[]]$keys) {
        foreach ($key in $keys) {
            $lower = $key.ToLowerInvariant()
            if ($pairs.ContainsKey($lower) -and $pairs[$lower]) { return $pairs[$lower] }
        }
        return ''
    }
    $info.Parsed = $true
    $info.ComputerName = get-battery-value @('computer name')
    $info.SystemProduct = get-battery-value @('system product name','system product','platform role')
    $info.OSBuild = get-battery-value @('os build')
    $info.BIOS = get-battery-value @('system bios','bios')
    $info.DesignCapacity = get-battery-value @('design capacity')
    $info.FullChargeCapacity = get-battery-value @('full charge capacity')
    $info.CycleCount = get-battery-value @('cycle count')
    $info.BatteryName = get-battery-value @('name','battery name')
    $info.BatterySerial = get-battery-value @('serial number')
    $info.BatteryChemistry = get-battery-value @('chemistry')
    $info.LastFullCharge = get-battery-value @('last full charge')
    try {
        $design = [double](([regex]::Replace($info.DesignCapacity, '[^\d.]', '')))
        $full = [double](([regex]::Replace($info.FullChargeCapacity, '[^\d.]', '')))
        if ($design -gt 0) { $info.HealthPct = [Math]::Round(($full / $design) * 100, 1) }
    } catch {}
    return New-SedaObject $info
}

function Get-SedaFirewallProfiles {
    param([string]$Path)
    $profiles = @()
    if (-not $Path) { return $profiles }
    $current = $null
    foreach ($raw in (Get-SedaTextContent -Path $Path -split "`r?`n")) {
        $line = $raw.Trim()
        $lower = $line.ToLowerInvariant()
        if ($lower -match 'domain profile|profil de domaine|domänenprofil|perfil de dominio|profilo di dominio') {
            $current = New-SedaObject @{ Name='Domain'; State=''; FirewallPolicy=''; RemoteManagement=''; InboundNotification=''; UnicastResponse=''; LogAllowed=''; LogDropped=''; LogFileName='' }; $profiles += $current; continue
        }
        if ($lower -match 'private profile|profil privé|privates profil|perfil privado|profilo privato') {
            $current = New-SedaObject @{ Name='Private'; State=''; FirewallPolicy=''; RemoteManagement=''; InboundNotification=''; UnicastResponse=''; LogAllowed=''; LogDropped=''; LogFileName='' }; $profiles += $current; continue
        }
        if ($lower -match 'public profile|profil public|öffentliches profil|perfil público|profilo pubblico') {
            $current = New-SedaObject @{ Name='Public'; State=''; FirewallPolicy=''; RemoteManagement=''; InboundNotification=''; UnicastResponse=''; LogAllowed=''; LogDropped=''; LogFileName='' }; $profiles += $current; continue
        }
        if (-not $current -or $line -match '^-{10,}') { continue }
        if ($line -match '^(\w[\w\s]{2,}?)\s{2,}(.+)$') {
            $key = ($Matches[1].Trim().ToLowerInvariant() -replace '\s+', '')
            $val = $Matches[2].Trim()
            if ($key -in @('state','état','zustand','estado','stato')) {
                $norm = $val.ToLowerInvariant().TrimEnd('.')
                if ($norm -in @('on','actif','activé','ein','activo','attivo','enabled')) { $val = 'ON' }
                elseif ($norm -in @('off','inactif','désactivé','aus','desactivado','disattivato','disabled')) { $val = 'OFF' }
                $current.State = $val
            } elseif ($key -in @('firewallpolicy','stratégiedepare-feu','firewallrichtlinie')) { $current.FirewallPolicy = $val }
            elseif ($key -in @('remotemanagement','administrationàdistance')) { $current.RemoteManagement = $val }
            elseif ($key -eq 'inboundusernotification') { $current.InboundNotification = $val }
            elseif ($key -eq 'unicastresponsetomulticast') { $current.UnicastResponse = $val }
            elseif ($key -like '*logallowed*') { $current.LogAllowed = $val }
            elseif ($key -like '*logdropped*') { $current.LogDropped = $val }
            elseif ($key -eq 'filename' -and -not $current.LogFileName) { $current.LogFileName = $val }
        }
    }
    return $profiles
}

function Get-SedaCertificates {
    param([string[]]$Paths)
    $certs = @()
    foreach ($path in @($Paths)) {
        if (-not $path) { continue }
        $text = Get-SedaTextContent -Path $path
        if (-not $text) { continue }
        foreach ($block in [regex]::Split($text, '={5,}.*?={5,}')) {
            if (-not $block.Trim()) { continue }
            $cert = New-SedaObject @{ Store=[System.IO.Path]::GetFileName($path); Serial=''; Issuer=''; Subject=''; NotBefore=''; NotAfter=''; Thumbprint=''; DaysToExpiry=$null; Status='OK' }
            foreach ($line in ($block -split "`r?`n")) {
                $l = $line.Trim()
                if ($l -match '^Serial Number:\s*(.+)$') { $cert.Serial = $Matches[1].Trim(); continue }
                if ($l -match '^Issuer:\s*(.+)$') { $cert.Issuer = $Matches[1].Trim(); continue }
                if ($l -match '^Subject:\s*(.+)$') { $cert.Subject = $Matches[1].Trim(); continue }
                if ($l -match '^NotBefore:\s*(.+)$') { $cert.NotBefore = $Matches[1].Trim(); continue }
                if ($l -match '^NotAfter:\s*(.+)$') { $cert.NotAfter = $Matches[1].Trim(); continue }
                if ($l -match '^(?:Cert Hash|Thumbprint)[^:]*:\s*(.+)$') { $cert.Thumbprint = ($Matches[1] -replace '\s+', '').ToUpperInvariant(); continue }
            }
            if (-not $cert.Serial -and -not $cert.Subject) { continue }
            if ($cert.NotAfter) {
                foreach ($fmt in @('M/d/yyyy h:mm tt','M/d/yyyy H:mm','yyyy-MM-dd H:mm','M/d/yyyy H:mm:ss','d/M/yyyy H:mm:ss','yyyy-MM-ddTHH:mm:ss')) {
                    try {
                        $dt = [datetime]::ParseExact($cert.NotAfter.Split('.')[0].Trim(), $fmt, [Globalization.CultureInfo]::InvariantCulture)
                        $cert.DaysToExpiry = [int]($dt - (Get-Date)).TotalDays
                        if ($cert.DaysToExpiry -lt 0) { $cert.Status = 'Expired' }
                        elseif ($cert.DaysToExpiry -lt 30) { $cert.Status = 'Expiring' }
                        break
                    } catch {}
                }
            }
            $certs += $cert
        }
    }
    return $certs
}

function Get-SedaEventLogScan {
    param(
        [hashtable]$Inventory,
        [int]$MaxEvents = 2000,
        [switch]$SkipScan
    )
    $logs = @{}
    foreach ($type in @('evtx_application','evtx_setup','evtx_system')) {
        $logs[$type] = New-SedaObject @{ LogType=$type; SourceFile=''; Events=@(); TotalCount=0; CriticalCount=0; ErrorCount=0; WarningCount=0; InfoCount=0; LastStatus='No EVTX file found' }
        $path = @($Inventory[$type] | Select-Object -First 1)[0]
        if (-not $path) { continue }
        $logs[$type].SourceFile = [System.IO.Path]::GetFileName($path)
        if ($SkipScan) { $logs[$type].LastStatus = 'EVTX scan skipped'; continue }
        $wevtutil = Get-Command wevtutil.exe -ErrorAction SilentlyContinue
        if (-not $wevtutil) { $logs[$type].LastStatus = 'wevtutil.exe not available'; continue }
        try {
            $xmlText = & $wevtutil.Source qe $path /lf:true /f:XML "/c:$MaxEvents" /rd:true 2>$null
            if (-not $xmlText) { $logs[$type].LastStatus = "wevtutil returned no output for $($logs[$type].SourceFile)"; continue }
            $clean = ($xmlText -join [Environment]::NewLine) -replace '<\?xml[^?]*\?>', ''
            $clean = [regex]::Replace($clean, '[\x00-\x08\x0b\x0c\x0e-\x1f]', '')
            [xml]$wrapped = "<Events>$clean</Events>"
            foreach ($eventNode in @($wrapped.Events.Event)) {
                $system = $eventNode.System
                $levelNum = 0
                try { $levelNum = [int]$system.Level } catch {}
                $level = switch ($levelNum) { 1 {'Critical'} 2 {'Error'} 3 {'Warning'} 4 {'Information'} 5 {'Verbose'} default {'Unknown'} }
                $timestamp = ''
                try { $timestamp = ([string]$system.TimeCreated.SystemTime).Substring(0, [Math]::Min(19, ([string]$system.TimeCreated.SystemTime).Length)).Replace('T',' ') } catch {}
                $messageParts = @()
                foreach ($container in @($eventNode.EventData, $eventNode.UserData, $eventNode.RenderingInfo)) {
                    if ($container) {
                        foreach ($child in $container.ChildNodes) {
                            $txt = ([string]$child.InnerText).Trim()
                            if ($txt) { $messageParts += $txt }
                        }
                    }
                }
                $provider = ''
                try { $provider = [string]$system.Provider.Name; if (-not $provider) { $provider = [string]$system.Provider.Guid } } catch {}
                $ev = New-SedaObject @{ LevelNum=$levelNum; Level=$level; Timestamp=$timestamp; EventId=[string]$system.EventID; Provider=$provider; Channel=[string]$system.Channel; Message=(($messageParts -join ' | ').Substring(0, [Math]::Min(300, ($messageParts -join ' | ').Length))); SourceFile=$logs[$type].SourceFile }
                $logs[$type].Events += $ev
                $logs[$type].TotalCount++
                if ($levelNum -eq 1) { $logs[$type].CriticalCount++ }
                elseif ($levelNum -eq 2) { $logs[$type].ErrorCount++ }
                elseif ($levelNum -eq 3) { $logs[$type].WarningCount++ }
                else { $logs[$type].InfoCount++ }
            }
            $logs[$type].LastStatus = "$($logs[$type].SourceFile): $($logs[$type].TotalCount) events - $($logs[$type].CriticalCount) critical, $($logs[$type].ErrorCount) errors, $($logs[$type].WarningCount) warnings, $($logs[$type].InfoCount) info"
        } catch {
            $logs[$type].LastStatus = "EVTX scan failed: $($_.Exception.Message)"
        }
    }
    $summary = @()
    $events = @()
    foreach ($key in @($logs.Keys)) {
        $log = $logs[$key]
        $logName = $key -replace '^evtx_', ''
        $summary += New-SedaObject @{ Log=$logName; Critical=$log.CriticalCount; Error=$log.ErrorCount; Warning=$log.WarningCount; Information=$log.InfoCount; Other=0; Scanned=$log.LastStatus; File=$log.SourceFile }
        foreach ($event in @($log.Events)) {
            $events += New-SedaObject @{ Log=$logName; TimeCreated=$event.Timestamp; Level=$event.Level; Id=$event.EventId; Provider=$event.Provider; Message=$event.Message; File=$event.SourceFile }
        }
    }
    return New-SedaObject @{ Logs=$logs; Summary=$summary; Events=$events }
}

function Get-SedaHealthReport {
    param([object]$Analysis)
    $findings = @()
    function add-finding([string]$Category,[string]$Severity,[string]$Title,[string]$Detail,[string]$Value='',[string]$Action='') {
        $script:__sedaHealthFindings += New-SedaObject @{ Area=$Category; Category=$Category; Severity=$Severity; Title=$Title; Details=$Detail; Detail=$Detail; Value=$Value; Recommendation=$Action; Action=$Action; Source=$Category }
    }
    $script:__sedaHealthFindings = @()

    $extDir = @($Analysis.ZipInfo.AllFiles | Where-Object { $_ -match '\\extended(\\|$)' } | Select-Object -First 1)
    if (-not $extDir) {
        $extendedFile = @($Analysis.ZipInfo.AllFiles | Where-Object { $_ -match '\\extended\\' } | Select-Object -First 1)[0]
        if ($extendedFile) { $extDir = Split-Path -Parent $extendedFile }
    }
    function find-ext([string]$Name) {
        if (-not $extDir) { return '' }
        $candidate = @(Get-ChildItem -LiteralPath $extDir -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.Name.ToLowerInvariant().Contains($Name.ToLowerInvariant()) } | Select-Object -First 1)[0]
        if ($candidate) { return $candidate.FullName }
        return ''
    }
    $bitLocker = Get-SedaTextContent -Path (find-ext 'ps_bitlocker_status')
    if ($bitLocker) {
        if ($bitLocker -match 'ProtectionStatus\s*:\s*Off') { add-finding 'BitLocker' 'ERROR' 'BitLocker protection is OFF' 'Drive is not actively protected by BitLocker.' 'ProtectionStatus=Off' 'Enable BitLocker or verify Intune encryption policy.' }
        elseif ($bitLocker -match 'ProtectionStatus\s*:\s*On') { add-finding 'BitLocker' 'OK' 'BitLocker protection ON' 'Drive is actively BitLocker-protected.' 'ProtectionStatus=On' }
        if ($bitLocker -notmatch 'KeyProtector.*RecoveryPassword') { add-finding 'BitLocker' 'WARN' 'No RecoveryPassword key protector detected' 'Recovery key may not be escrowed to Entra ID / Intune.' '' 'Verify BitLocker recovery key escrow.' }
    } else { add-finding 'BitLocker' 'INFO' 'BitLocker data unavailable' 'ps_bitlocker_status not found.' }

    $defender = Get-SedaTextContent -Path (find-ext 'ps_defender_status')
    if ($defender) {
        if ($defender -match 'RealTimeProtectionEnabled\s*:\s*False') { add-finding 'Defender' 'ERROR' 'Real-time protection DISABLED' 'RealTimeProtectionEnabled=False' '' 'Re-enable via Intune Defender policy.' }
        elseif ($defender -match 'RealTimeProtectionEnabled\s*:\s*True') { add-finding 'Defender' 'OK' 'Real-time protection enabled' '' }
        if ($defender -match 'AntivirusEnabled\s*:\s*False') { add-finding 'Defender' 'ERROR' 'Antivirus disabled' 'AntivirusEnabled=False' '' 'Check Tamper Protection and Intune Antivirus policy.' }
        if ($defender -match 'AMRunningMode\s*:\s*(.+)') {
            $mode = $Matches[1].Trim()
            if ($mode -match 'passive') { add-finding 'Defender' 'WARN' 'Defender running in Passive mode' "AMRunningMode=$mode" '' 'Verify EDR configuration if using third-party AV.' }
            else { add-finding 'Defender' 'OK' "Defender mode: $mode" '' }
        }
    } else { add-finding 'Defender' 'INFO' 'Defender data unavailable' 'ps_defender_status not found.' }

    $deviceJoin = ConvertTo-SedaHashtable $Analysis.DeviceInfo
    $sso = ConvertTo-SedaHashtable $Analysis.SsoInfo
    $azureJoined = [string]$deviceJoin.AzureAdJoined
    $workplaceJoined = [string]$deviceJoin.WorkplaceJoined
    $domainJoined = [string]$deviceJoin.DomainJoined
    $authStatus = [string]$sso.DeviceAuthStatus
    if ($azureJoined -match 'YES' -or $workplaceJoined -match 'YES') {
        add-finding 'Entra Join' 'OK' 'Device has an Entra join state' "AzureAdJoined=$azureJoined; WorkplaceJoined=$workplaceJoined; DomainJoined=$domainJoined"
    } else {
        add-finding 'Entra Join' 'WARN' 'Device is not Azure AD / Workplace joined' "AzureAdJoined=$azureJoined; WorkplaceJoined=$workplaceJoined; DomainJoined=$domainJoined" '' 'Verify expected join model and Intune enrollment.'
    }
    if ($authStatus) {
        if ($authStatus -match 'FAILED|ERROR|disabled|deleted') { add-finding 'Entra Join' 'ERROR' 'Device authentication failed' $authStatus '' 'Check Entra device object and run dsregcmd /status.' }
        elseif ($authStatus -match 'SUCCESS') { add-finding 'Entra Join' 'OK' 'Device authentication successful' $authStatus }
    }

    $disk = Get-SedaTextContent -Path (find-ext 'ps_disk_usage')
    if ($disk) {
        foreach ($line in ($disk -split "`r?`n")) {
            $parts = @($line -split '\s+' | Where-Object { $_ })
            if ($parts.Count -lt 4 -or $parts[0] -notmatch '^[A-Za-z]$') { continue }
            try {
                $freePct = [double]($parts[-1] -replace ',', '.')
                $freeGb = ([double]($parts[2] -replace ',', '')) / 1GB
                $totalGb = ([double]($parts[3] -replace ',', '')) / 1GB
                if ($freePct -lt 10 -or $freeGb -lt 5) { add-finding 'Storage' 'ERROR' "Drive $($parts[0]): critically low disk space" ("{0:N1} GB free / {1:N1} GB total ({2:N1}% free)" -f $freeGb,$totalGb,$freePct) "$freePct%" 'Run cleanup or expand disk.' }
                elseif ($freePct -lt 20) { add-finding 'Storage' 'WARN' "Drive $($parts[0]): low disk space" ("{0:N1} GB free / {1:N1} GB total ({2:N1}% free)" -f $freeGb,$totalGb,$freePct) "$freePct%" }
                else { add-finding 'Storage' 'OK' "Drive $($parts[0]): $([Math]::Round($freeGb,1)) GB free ($freePct%)" '' }
            } catch {}
        }
    } else { add-finding 'Storage' 'INFO' 'Disk usage data unavailable' '' }

    $topProcesses = Get-SedaTextContent -Path (find-ext 'ps_top_processes')
    if ($topProcesses) {
        $highCpu = @()
        foreach ($line in (($topProcesses -split "`r?`n") | Select-Object -Skip 2 -First 12)) {
            $parts = @($line -split '\s+' | Where-Object { $_ })
            if ($parts.Count -ge 3) {
                try {
                    $cpu = [double]($parts[2] -replace ',', '.')
                    if ($cpu -gt 50) { $highCpu += "$($parts[0]) ($([math]::Round($cpu,0))s CPU)" }
                } catch {}
            }
        }
        if ($highCpu.Count -gt 0) { add-finding 'Performance' 'WARN' "$($highCpu.Count) process(es) with high CPU time" ($highCpu -join ', ') '' 'Investigate in Task Manager or Event Viewer.' }
        else { add-finding 'Performance' 'OK' 'No runaway CPU processes detected' '' }
    }
    $startup = Get-SedaTextContent -Path (find-ext 'ps_startup_programs')
    if ($startup) {
        $startupLines = @($startup -split "`r?`n" | Where-Object { $_.Trim() -and $_ -notmatch '^-+$' -and $_ -notmatch '^\s*Name\s+Command' -and $_ -notmatch '^ps_startup_programs' })
        if ($startupLines.Count -gt 15) { add-finding 'Performance' 'WARN' "High startup program count ($($startupLines.Count) items)" "$($startupLines.Count) startup entries detected." "$($startupLines.Count)" 'Review startup items via Task Manager.' }
        elseif ($startupLines.Count -gt 0) { add-finding 'Performance' 'OK' "Startup programs: $($startupLines.Count) items" '' }
    }
    $sysInfo = Get-SedaTextContent -Path (find-ext 'ps_system_info')
    if ($sysInfo -match 'CsTotalPhysicalMemory\s*:\s*(\d+)') {
        $ramGb = [double]$Matches[1] / 1GB
        if ($ramGb -lt 4) { add-finding 'Performance' 'ERROR' ("Insufficient RAM ({0:N1} GB)" -f $ramGb) 'Minimum for Windows 11 is 4 GB; 8 GB recommended for Intune workloads.' ("{0:N1} GB" -f $ramGb) 'Upgrade RAM or investigate memory-heavy processes.' }
        elseif ($ramGb -lt 8) { add-finding 'Performance' 'WARN' ("Low RAM ({0:N1} GB)" -f $ramGb) '8 GB recommended for smooth operation.' ("{0:N1} GB" -f $ramGb) }
        else { add-finding 'Performance' 'OK' ("RAM: {0:N1} GB" -f $ramGb) '' }
    }

    $aadText = (@($Analysis.EventLogs.Events | Where-Object { $_.Log -match 'AAD|Azure' -or $_.Provider -match 'AAD|Azure' -or $_.Message -match 'AADSTS|Microsoft\.AAD' } | Select-Object -First 200 | ForEach-Object { "$($_.Id) $($_.Provider) $($_.Message)" }) -join [Environment]::NewLine)
    $authCodes = @('70011','70016','70043','50076','50126','50132','50133','53003','65001','700082','AADSTS')
    $foundAuthCodes = @($authCodes | Where-Object { $aadText -match [regex]::Escape($_) })
    if ($aadText) {
        if ($foundAuthCodes.Count -gt 0) { add-finding 'Office Auth' 'ERROR' "$($foundAuthCodes.Count) Modern Auth error code(s) detected" ($foundAuthCodes -join ', ') '' 'Check AAD Operational log and Conditional Access policies.' }
        else { add-finding 'Office Auth' 'OK' 'No Modern Auth error codes detected in AAD events' '' }
    } else { add-finding 'Office Auth' 'INFO' 'AAD Operational event log not available' 'Enable Microsoft-Windows-AAD/Operational or run local collection.' }
    $proxyText = Get-SedaTextContent -Path (find-ext 'ps_proxy_config')
    if ($proxyText -match 'ProxyServer\s*:\s*(.+)') {
        $proxy = $Matches[1].Trim()
        if ($proxy -and $proxy -notmatch '^\(?none\)?$|direct') { add-finding 'Office Auth' 'WARN' "Proxy configured: $proxy" 'Proxies can interfere with Modern Auth token acquisition.' $proxy 'Verify proxy bypass list for Microsoft identity and Office endpoints.' }
    }

    $apps32 = @($Analysis.Applications | Where-Object { $_.Arch -eq 'x86' })
    if ($apps32.Count -gt 0) { add-finding 'Legacy Apps' 'WARN' "$($apps32.Count) 32-bit application(s) detected" (($apps32 | Select-Object -First 12 -ExpandProperty Name) -join ', ') '' 'Verify Windows 11 x64 compatibility.' }
    else { add-finding 'Legacy Apps' 'OK' 'No 32-bit apps detected' '' }
    $legacy = @($Analysis.Applications | Where-Object { $_.Name -match 'Visual C\+\+ 2005|Visual C\+\+ 2008|Visual C\+\+ 2010|\.NET Framework 2\.0|\.NET Framework 3\.0|Java 6|Java 7|Java 8|Adobe Flash|Silverlight|Python 2\.' })
    if ($legacy.Count -gt 0) { add-finding 'Legacy Apps' 'ERROR' "$($legacy.Count) legacy/EOL runtime(s) installed" (($legacy | Select-Object -First 10 -ExpandProperty Name) -join ', ') '' 'Update or remove legacy runtimes.' }

    $storeApps = Get-SedaTextContent -Path (find-ext 'ps_store_apps')
    if ($storeApps) {
        $storeLines = @($storeApps -split "`r?`n" | Where-Object { $_.Trim() -and $_ -notmatch '^-+$' -and $_ -notmatch '^\s*Name\s+' })
        $missingCritical = @()
        foreach ($pkg in @('Microsoft.WindowsStore','Microsoft.StorePurchaseApp')) {
            if ($storeApps -notmatch [regex]::Escape($pkg)) { $missingCritical += $pkg }
        }
        add-finding 'Store Apps' 'INFO' "$($storeLines.Count) Store app package line(s) found" ''
        if ($missingCritical.Count -gt 0) { add-finding 'Store Apps' 'WARN' 'Some critical Store packages may be missing/deprovisioned' ($missingCritical -join ', ') '' 'Re-provision via Intune or Microsoft Store where needed.' }
    } else { add-finding 'Store Apps' 'INFO' 'Store apps list unavailable' 'ps_store_apps not found.' }

    $badDrivers = @($Analysis.Drivers | Where-Object { $_.Provider -match 'unknown|unsigned|test certificate' -or $_.Signer -match 'unknown|unsigned|test certificate' })
    if ($badDrivers.Count -gt 0) { add-finding 'Drivers' 'WARN' "$($badDrivers.Count) potentially unsigned/unknown driver(s)" (($badDrivers | Select-Object -First 10 -ExpandProperty OriginalName) -join ', ') '' 'Verify driver signatures.' }
    elseif ($Analysis.Drivers.Count -gt 0) { add-finding 'Drivers' 'OK' "All $($Analysis.Drivers.Count) drivers have known publishers" '' }
    else { add-finding 'Drivers' 'INFO' 'No driver data available' '' }

    foreach ($profile in @($Analysis.Hardware.FirewallProfiles)) {
        if ($profile.State -match 'OFF|DISABLED') { add-finding 'Firewall' 'ERROR' "$($profile.Name) firewall profile is OFF" "State=$($profile.State)" '' 'Enable Windows Firewall profile.' }
    }
    if (@($Analysis.Hardware.FirewallProfiles).Count -gt 0 -and -not @($Analysis.Hardware.FirewallProfiles | Where-Object { $_.State -match 'OFF|DISABLED' }).Count) {
        add-finding 'Firewall' 'OK' 'Windows Firewall profiles are enabled' ''
    }
    if ($Analysis.Hardware.Battery.Parsed -and $Analysis.Hardware.Battery.HealthPct -gt 0) {
        if ($Analysis.Hardware.Battery.HealthPct -lt 50) { add-finding 'Battery' 'ERROR' "Battery health critically low ($($Analysis.Hardware.Battery.HealthPct)%)" '' "$($Analysis.Hardware.Battery.HealthPct)%" 'Plan battery replacement.' }
        elseif ($Analysis.Hardware.Battery.HealthPct -lt 80) { add-finding 'Battery' 'WARN' "Battery health degraded ($($Analysis.Hardware.Battery.HealthPct)%)" '' "$($Analysis.Hardware.Battery.HealthPct)%" }
        else { add-finding 'Battery' 'OK' "Battery health: $($Analysis.Hardware.Battery.HealthPct)%" '' }
    }
    $certProblems = @($Analysis.Hardware.Certificates | Where-Object { $_.Status -in @('Expired','Expiring') })
    foreach ($cert in $certProblems | Select-Object -First 20) {
        add-finding 'Certificates' ($(if ($cert.Status -eq 'Expired') {'ERROR'} else {'WARN'})) "$($cert.Status) certificate" $cert.Subject "$($cert.DaysToExpiry) days" 'Renew or remove stale certificate.'
    }

    $errors = @($script:__sedaHealthFindings | Where-Object { $_.Severity -eq 'ERROR' })
    $warnings = @($script:__sedaHealthFindings | Where-Object { $_.Severity -eq 'WARN' })
    $findings = $script:__sedaHealthFindings
    Remove-Variable -Name __sedaHealthFindings -Scope Script -ErrorAction SilentlyContinue
    return New-SedaObject @{ Findings=$findings; ErrorCount=$errors.Count; WarningCount=$warnings.Count; Summary=[ordered]@{ Total=$findings.Count; Errors=$errors.Count; Warnings=$warnings.Count; Status=if ($errors.Count -gt 0) { 'Critical' } elseif ($warnings.Count -gt 0) { 'Review recommended' } else { 'Healthy' } } }
}

function Get-SedaAIConfig {
    $defaults = [ordered]@{ Provider='claude'; ApiKey=''; Model='claude-haiku-4-5-20251001'; OllamaUrl='http://localhost:11434'; MaxTokens=2048; Temperature=0.3; RememberApiKey=$false }
    if (Test-Path -LiteralPath $script:AIConfigPath) {
        try {
            $json = Get-Content -LiteralPath $script:AIConfigPath -Raw | ConvertFrom-Json
            foreach ($prop in $json.PSObject.Properties) {
                $name = switch ($prop.Name) { 'provider' {'Provider'} 'api_key' {'ApiKey'} 'model' {'Model'} 'ollama_url' {'OllamaUrl'} 'max_tokens' {'MaxTokens'} 'temperature' {'Temperature'} 'remember_api_key' {'RememberApiKey'} default { $prop.Name } }
                if ($defaults.Contains($name)) { $defaults[$name] = $prop.Value }
            }
        } catch {}
    }
    return New-SedaObject $defaults
}

function Save-SedaAIConfig {
    param([object]$Config)
    $data = [ordered]@{
        provider = $Config.Provider
        api_key = if ($Config.RememberApiKey) { $Config.ApiKey } else { '' }
        model = $Config.Model
        ollama_url = $Config.OllamaUrl
        max_tokens = [int]$Config.MaxTokens
        temperature = [double]$Config.Temperature
        remember_api_key = [bool]$Config.RememberApiKey
    }
    try { $data | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $script:AIConfigPath -Encoding UTF8 } catch {}
}

function Build-SedaAIPrompt {
    param([object]$Analysis)
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add('# Intune Device Diagnostic Analysis')
    $lines.Add('')
    if (-not [string]::IsNullOrWhiteSpace($script:AiFocusContext)) {
        $lines.Add('## User-selected focus')
        $lines.Add($script:AiFocusContext)
        $lines.Add('')
    }
    $lines.Add('## Device Information')
    foreach ($pair in @('ComputerName','IPAddress','OSVersion','Proxy','LastUser','ImeVersion')) { $lines.Add("- **$pair**: $($Analysis.DeviceSummary.$pair)") }
    $lines.Add('')
    $lines.Add('## Compliance Status')
    $lines.Add("- Overall: **$($Analysis.Compliance.OverallStatus)**")
    $lines.Add("- Compliant policies: $($Analysis.Compliance.CompliantCount)")
    $lines.Add("- Non-compliant: $($Analysis.Compliance.NonCompliantCount)")
    $lines.Add("- Pending: $($Analysis.Compliance.PendingCount)")
    foreach ($policy in @($Analysis.Compliance.PolicyStatuses | Where-Object { $_.Status -in @('NON_COMPLIANT','FAILED','ERROR') } | Select-Object -First 20)) { $lines.Add("  - **[$($policy.Area)]** $($policy.Details)") }
    $lines.Add('')
    if ($Analysis.CriticalIssues.Count -gt 0) {
        $lines.Add('## Critical Issues Detected')
        foreach ($issue in @($Analysis.CriticalIssues)) { $lines.Add("- **[$($issue.Severity)]** [$($issue.Category)] $($issue.Title) - $($issue.Detail)") }
        $lines.Add('')
    }
    $imeIssues = @($Analysis.ImeEvents | Where-Object { $_.Severity -in @('ERROR','WARNING') } | Select-Object -First 50)
    if ($imeIssues.Count -gt 0) {
        $lines.Add("## IME Log Issues (showing $($imeIssues.Count))")
        foreach ($ev in $imeIssues) { $lines.Add("- [$($ev.Severity)] [$($ev.Category)] $($ev.Message) $(if ($ev.ErrorCode) { "(code: $($ev.ErrorCode))" })") }
        $lines.Add('')
    }
    if ($Analysis.WindowsUpdate.Info.Count -gt 0) {
        $lines.Add('## Windows Update Status')
        foreach ($entry in (ConvertTo-SedaHashtable $Analysis.WindowsUpdate.Info).GetEnumerator()) { $lines.Add("- **$($entry.Key)**: $($entry.Value)") }
        foreach ($ev in @($Analysis.WindowsUpdate.ReportingEvents | Where-Object { $_.ErrorCode } | Select-Object -First 15)) { $lines.Add("  - [$($ev.Source)] $($ev.ErrorCode) - $($ev.Message)") }
        foreach ($ev in @($Analysis.WindowsUpdate.EtlEvents | Where-Object { $_.Level -in @('Critical','Error','Warning') } | Select-Object -First 15)) { $lines.Add("  - [ETL/$($ev.Source)] $($ev.ErrorCode) - $($ev.Message)") }
        $lines.Add('')
    }
    $evtxErrors = @($Analysis.EventLogs.Events | Where-Object { $_.Level -in @('Critical','Error') })
    if ($evtxErrors.Count -gt 0) {
        $lines.Add("## Windows Event Log Errors ($($evtxErrors.Count) total)")
        foreach ($ev in @($evtxErrors | Select-Object -First 30)) { $lines.Add("- [$($ev.Level)] [$($ev.Log)] EventID $($ev.Id) ($($ev.Provider)) - $($ev.Message)") }
        $lines.Add('')
    }
    if ($Analysis.Health.Findings.Count -gt 0) {
        $lines.Add('## Health Findings')
        foreach ($f in @($Analysis.Health.Findings | Where-Object { $_.Severity -in @('ERROR','WARN') } | Select-Object -First 40)) { $lines.Add("- [$($f.Severity)] [$($f.Area)] $($f.Title) - $($f.Details)") }
        $lines.Add('')
    }
    $lines.Add('---')
    $lines.Add('Please analyze this data and provide your expert assessment with prioritized issues and remediation steps.')
    return ($lines -join [Environment]::NewLine)
}

function Invoke-SedaAIAnalysis {
    param(
        [object]$Analysis,
        [object]$Config
    )
    $systemPrompt = 'You are an expert Microsoft Intune administrator and Windows systems engineer. Analyze Intune device diagnostic data and provide clear, actionable guidance with an executive summary, priority issues, root cause analysis, remediation steps, and preventive recommendations. Be direct and practical. Always mention specific error codes and event IDs found.'
    if (-not $Config.ApiKey -and $Config.Provider -ne 'ollama') { throw "No API key configured for $($Config.Provider)." }
    $prompt = Build-SedaAIPrompt -Analysis $Analysis
    $headers = @{ 'Content-Type' = 'application/json' }
    if ($Config.Provider -eq 'claude') {
        $headers['x-api-key'] = $Config.ApiKey
        $headers['anthropic-version'] = '2023-06-01'
        $body = @{ model=$Config.Model; max_tokens=[int]$Config.MaxTokens; system=$systemPrompt; messages=@(@{ role='user'; content=$prompt }) } | ConvertTo-Json -Depth 8
        $result = Invoke-RestMethod -Uri 'https://api.anthropic.com/v1/messages' -Method Post -Headers $headers -Body $body -TimeoutSec 120
        return [string]$result.content[0].text
    }
    if ($Config.Provider -eq 'openai') {
        $headers['Authorization'] = "Bearer $($Config.ApiKey)"
        $body = @{ model=$Config.Model; max_tokens=[int]$Config.MaxTokens; temperature=[double]$Config.Temperature; messages=@(@{ role='system'; content=$systemPrompt }, @{ role='user'; content=$prompt }) } | ConvertTo-Json -Depth 8
        $result = Invoke-RestMethod -Uri 'https://api.openai.com/v1/chat/completions' -Method Post -Headers $headers -Body $body -TimeoutSec 120
        return [string]$result.choices[0].message.content
    }
    $body = @{ model=$Config.Model; prompt="$systemPrompt`n`n$prompt"; stream=$false; options=@{ num_predict=[int]$Config.MaxTokens; temperature=[double]$Config.Temperature } } | ConvertTo-Json -Depth 8
    $result = Invoke-RestMethod -Uri (($Config.OllamaUrl.TrimEnd('/')) + '/api/generate') -Method Post -Headers $headers -Body $body -TimeoutSec 180
    return [string]$result.response
}

function Get-SedaInsights {
    param([object]$Analysis)
    $actions = @()
    $timeline = @()
    $penalty = 0

    foreach ($issue in @($Analysis.CriticalIssues)) {
        $actions += New-SedaObject @{ Severity = $issue.Severity; Title = $issue.Title; Detail = $issue.Detail; Recommendation = $issue.Recommendation; Source = $issue.Source }
        $penalty += if ($issue.Severity -eq 'ERROR') { 15 } else { 7 }
    }
    foreach ($status in @($Analysis.Compliance.PolicyStatuses)) {
        if ($status.Status -in @('NON_COMPLIANT', 'FAILED', 'ERROR')) {
            $actions += New-SedaObject @{ Severity = $status.Status; Title = "$($status.Area) is not compliant"; Detail = $status.Details; Recommendation = 'Review matching policy and device-side evidence.'; Source = $status.SourceFile }
            $penalty += 10
        }
    }
    $imeErrors = @($Analysis.ImeEvents | Where-Object { $_.Severity -eq 'ERROR' })
    if ($imeErrors.Count -gt 0) {
        $first = $imeErrors[0]
        $actions += New-SedaObject @{ Severity = 'ERROR'; Title = "IME logs contain $($imeErrors.Count) error(s)"; Detail = $first.Message; Recommendation = 'Review IME theme, app install, detection rule, or remediation output.'; Source = $first.SourceFile }
        $penalty += [Math]::Min(20, $imeErrors.Count * 2)
    }
    foreach ($wuIssue in @($Analysis.WindowsUpdate.Issues)) {
        $actions += New-SedaObject @{ Severity = $wuIssue.Severity; Title = $wuIssue.Title; Detail = $wuIssue.Detail; Recommendation = $wuIssue.Recommendation; Source = 'Windows Update' }
        $penalty += if ($wuIssue.Severity -eq 'ERROR') { 8 } else { 5 }
    }
    foreach ($indicator in @($Analysis.Win11Compatibility.BlockingIndicators)) {
        $actions += New-SedaObject @{ Severity = 'ERROR'; Title = "Windows 11 upgrade blocked for $($indicator.TargetVersion)"; Detail = $indicator.ReasonText; Recommendation = 'Review AppCompat TargetVersionUpgradeExperienceIndicators and resolve the blocker.'; Source = 'Win11 Upgrade Experience' }
        $penalty += 12
    }
    foreach ($check in @($Analysis.Win11Compatibility.HardwareReadiness.Checks | Where-Object { $_.Status -in @('FAIL','UNDETERMINED','UNKNOWN') })) {
        $severity = if ($check.Status -eq 'FAIL') { 'ERROR' } else { 'WARNING' }
        $actions += New-SedaObject @{ Severity = $severity; Title = "Windows 11 readiness: $($check.Requirement)"; Detail = "$($check.Status) - $($check.Detail)"; Recommendation = 'Review Microsoft HardwareReadiness.ps1 result and remediate the failing hardware requirement.'; Source = 'HardwareReadiness.ps1' }
        $penalty += if ($severity -eq 'ERROR') { 10 } else { 4 }
    }
    foreach ($finding in @($Analysis.Health.Findings)) {
        if ($finding.Severity -in @('ERROR','WARN')) {
            $actions += New-SedaObject @{ Severity = $finding.Severity; Title = $finding.Title; Detail = $finding.Details; Recommendation = $finding.Recommendation; Source = $finding.Area }
            $penalty += if ($finding.Severity -eq 'ERROR') { 10 } else { 5 }
        }
    }
    foreach ($event in @($Analysis.ImeEvents | Select-Object -First 200)) {
        $timeline += New-SedaObject @{ Timestamp = $event.Timestamp; Severity = $event.Severity; Source = "IME/$($event.Theme)"; Title = if ($event.KnownCode) { $event.KnownCode } else { $event.Category }; Detail = $event.Message }
    }
    foreach ($event in @($Analysis.WindowsUpdate.ReportingEvents | Select-Object -First 100)) {
        $timeline += New-SedaObject @{ Timestamp = $event.Timestamp; Severity = $event.Level; Source = 'Windows Update'; Title = $event.Source; Detail = $event.Message }
    }
    foreach ($event in @($Analysis.WindowsUpdate.EtlEvents | Where-Object { $_.Level -in @('Critical','Error','Warning') } | Select-Object -First 100)) {
        $actions += New-SedaObject @{ Severity=$event.Level; Title='Windows Update ETL event'; Detail="$($event.ErrorCode) $($event.Message)"; Recommendation='Review Windows Update ETL timeline around this event.'; Source=$event.EtlFile }
        $timeline += New-SedaObject @{ Timestamp = $event.Timestamp; Severity = $event.Level; Source = 'Windows Update ETL'; Title = $event.Source; Detail = $event.Message }
        if ($event.Level -in @('Critical','Error')) { $penalty += 2 } else { $penalty += 1 }
    }
    foreach ($event in @($Analysis.EventLogs.Events | Where-Object { $_.Level -in @('Critical','Error','Warning') } | Select-Object -First 300)) {
        $timeline += New-SedaObject @{ Timestamp = $event.TimeCreated; Severity = $event.Level; Source = "EventLog/$($event.Log)"; Title = "Event $($event.Id)"; Detail = $event.Message }
        if ($event.Level -in @('Critical','Error')) { $penalty += 1 }
    }

    $wufbEntries = @()
    foreach ($label in @('Reboot Required','Pre-shutdown Reboot Required','Feature Update Pause Enabled','Quality Update Pause Enabled','Feature Update Deferral (days)','Quality Update Deferral (days)','Next WU Refresh Time','WUfB Policy Hash','WUfB Policy Sync Date')) {
        if ($Analysis.WindowsUpdate.Info[$label]) {
            $sev = if ($label -like '*Required*' -and $Analysis.WindowsUpdate.Info[$label] -eq 'Yes') { 'WARNING' } else { 'INFO' }
            $wufbEntries += New-SedaObject @{ Severity=$sev; Title=$label; Detail=[string]$Analysis.WindowsUpdate.Info[$label]; Recommendation=''; Source='WU Orchestrator' }
        }
    }
    foreach ($policy in @($Analysis.WindowsUpdate.Policies)) {
        $sev = 'INFO'
        if ($policy.Label -like '*Disable*' -and $policy.Value -like '*Yes*') { $sev = 'WARNING' }
        $wufbEntries += New-SedaObject @{ Severity=$sev; Title=$policy.Label; Detail=[string]$policy.Value; Recommendation=''; Source=$policy.KeyPath }
    }
    foreach ($entry in @($wufbEntries | Where-Object { $_.Severity -eq 'WARNING' })) {
        $actions += New-SedaObject @{ Severity=$entry.Severity; Title=$entry.Title; Detail=$entry.Detail; Recommendation='Review Windows Update for Business policy.'; Source=$entry.Source }
        $penalty += 4
    }

    $score = [Math]::Max(0, [Math]::Min(100, 100 - $penalty))
    $status = if ($score -ge 85) { 'Healthy' } elseif ($score -ge 70) { 'Attention' } elseif ($score -ge 50) { 'Degraded' } else { 'Critical' }
    $top = @($actions | Sort-Object @{ Expression = { switch -Regex ($_.Severity) { 'ERROR|NON_COMPLIANT|CRITICAL' { 0; break } 'WARN|WARNING' { 1; break } default { 2 } } } } | Select-Object -First 5)
    $rootCauses = @($actions | Sort-Object @{ Expression = { switch -Regex ($_.Severity) { 'ERROR|NON_COMPLIANT|CRITICAL' { 0; break } 'WARN|WARNING' { 1; break } default { 2 } } } } | Select-Object -First 12)
    $searchRows = @()
    foreach ($item in @($top)) { $searchRows += New-SedaObject @{ Area='Top action'; Severity=$item.Severity; Title=$item.Title; Detail=$item.Detail; Source=$item.Source } }
    foreach ($item in @($rootCauses)) { $searchRows += New-SedaObject @{ Area='Root cause'; Severity=$item.Severity; Title=$item.Title; Detail=$item.Detail; Source=$item.Source } }
    foreach ($item in @($wufbEntries)) { $searchRows += New-SedaObject @{ Area='WUfB'; Severity=$item.Severity; Title=$item.Title; Detail=$item.Detail; Source=$item.Source } }
    foreach ($item in @($timeline | Select-Object -First 300)) { $searchRows += New-SedaObject @{ Area='Timeline'; Severity=$item.Severity; Title=$item.Title; Detail=$item.Detail; Source=$item.Source } }
    $wufbStatus = if (@($wufbEntries | Where-Object { $_.Severity -eq 'ERROR' }).Count -gt 0) { 'Action required' } elseif (@($wufbEntries | Where-Object { $_.Severity -eq 'WARNING' }).Count -gt 0) { 'Review recommended' } elseif ($wufbEntries.Count -gt 0) { 'No blocking WUfB issue detected' } else { 'No WUfB data found' }
    return New-SedaObject @{ Score = $score; Status = $status; TopActions = $top; RootCauses = $rootCauses; Timeline = @($timeline | Sort-Object Timestamp | Select-Object -First 300); Wufb = New-SedaObject @{ Status=$wufbStatus; Entries=$wufbEntries }; SearchRows = $searchRows }
}

function Invoke-SedaAnalysis {
    param([Parameter(Mandatory)][string]$Path)
    Write-SedaLog -Level INFO -Message "Analysis started: $Path"
    $zip = Expand-SedaDiagnosticZip -Path $Path
    $inventory = $zip.Inventory
    $dsregPath = Find-SedaInventoryFile -Inventory $inventory -Keyword 'dsregcmd'
    $enrollPath = Find-SedaInventoryFile -Inventory $inventory -Keyword 'enrollments'
    if ($enrollPath -and [System.IO.Path]::GetExtension($enrollPath).ToLowerInvariant() -ne '.reg') { $enrollPath = '' }
    $firewallPath = Find-SedaInventoryFile -Inventory $inventory -Keyword 'advfirewall_show_allprofiles'
    $dsreg = Get-SedaDsRegCmd -Path $dsregPath
    $enrollments = Get-SedaEnrollments -Path $enrollPath
    $results = Get-SedaResultsXml -Paths $inventory.results_xml
    $firewall = Get-SedaFirewallIssues -Path $firewallPath
    $ime = Get-SedaImeLogEvents -ImeThemes $inventory.ime_themes
    $extra = Get-SedaExtraSummary -Inventory $inventory

    $wuReg = @($inventory.wu_registry | Select-Object -First 1)[0]
    if (-not $wuReg) { $wuReg = Find-SedaInventoryFile -Inventory $inventory -Keyword 'windowsupdate_orchestrator' }
    $allReg = @($inventory.registry) + @($inventory.wu_registry)
    $reporting = Find-SedaInventoryFile -Inventory $inventory -Keyword 'reportingevents'
    $wu = Get-SedaWindowsUpdateInfo -OrchestratorReg $wuReg -AllRegFiles $allReg -ReportingEvents $reporting
    $wuGeneratedLog = @($inventory.wu_generated_log | Select-Object -First 1)[0]
    $wuGeneratedDir = Join-Path $zip.ExtractDir 'GeneratedLogs'
    $wuEtl = Get-SedaWuEtlEvents -Paths $inventory.wu_etl -GeneratedLogPath $wuGeneratedLog -GeneratedOutputDirectory $wuGeneratedDir
    $wu | Add-Member -NotePropertyName EtlStatus -NotePropertyValue $wuEtl.Status -Force
    $wu | Add-Member -NotePropertyName EtlEvents -NotePropertyValue $wuEtl.Events -Force
    $wu | Add-Member -NotePropertyName EtlErrorCount -NotePropertyValue $wuEtl.ErrorCount -Force
    $wu | Add-Member -NotePropertyName EtlWarningCount -NotePropertyValue $wuEtl.WarningCount -Force
    $wu | Add-Member -NotePropertyName EtlFilesScanned -NotePropertyValue $wuEtl.FilesScanned -Force
    $wu | Add-Member -NotePropertyName GeneratedLogPath -NotePropertyValue $wuEtl.GeneratedLogPath -Force

    $apps = Get-SedaInstalledApps -RegFiles (@($inventory.reg_uninstall_x64) + @($inventory.reg_uninstall_x86))
    $drivers = Get-SedaDrivers -Path @($inventory.cmd_pnputil | Select-Object -First 1)[0]
    $wifi = Get-SedaWifiProfiles -Path @($inventory.cmd_wlan_profiles | Select-Object -First 1)[0]
    $battery = Get-SedaBatteryReport -Path @($inventory.battery_report | Select-Object -First 1)[0]
    $firewallProfiles = Get-SedaFirewallProfiles -Path $firewallPath
    $certificates = Get-SedaCertificates -Paths $inventory.cmd_certutil
    $mdmDiag = Get-SedaMdmDiagReport -CabFiles $inventory.cab -SkipExtraction:$SkipCabExtraction
    $eventLogs = Get-SedaEventLogScan -Inventory $inventory -SkipScan:$SkipEventLogScan
    $win11File = @($inventory.reg_win11_upgrade_indicators | Select-Object -First 1)[0]
    if (-not $win11File) { $win11File = Find-SedaInventoryFile -Inventory $inventory -Keyword 'targetversionupgradeexperienceindicators' }
    $win11 = Get-SedaWin11CompatibilityIndicators -Path $win11File
    $hardwareReadinessFile = @($inventory.win11_readiness | Select-Object -First 1)[0]
    if (-not $hardwareReadinessFile) { $hardwareReadinessFile = Find-SedaInventoryFile -Inventory $inventory -Keyword 'win11_readiness' }
    $hardwareReadiness = Get-SedaHardwareReadiness -Path $hardwareReadinessFile
    $win11 | Add-Member -NotePropertyName HardwareReadiness -NotePropertyValue $hardwareReadiness -Force
    if ($hardwareReadiness.Parsed -and $hardwareReadiness.Status) {
        $win11 | Add-Member -NotePropertyName Status -NotePropertyValue $hardwareReadiness.Status -Force
    }
    $compliance = Get-SedaComplianceSummary -DsReg $dsreg -Enrollments $enrollments -Firewall $firewall -Results $results
    $mdmDiagIssues = @($mdmDiag.Issues | ForEach-Object {
        New-SedaObject @{ Severity = $_.Severity; Category = $_.Area; Title = $_.Title; Detail = $_.Detail; Recommendation = $_.Recommendation; Source = 'MDM Diagnostics' }
    })
    $critical = @($dsreg.CriticalIssues) + @($firewall.Issues) + @($wu.Issues) + @($mdmDiagIssues)

    $deviceName = $dsreg.DeviceInfo['Device Name']
    if (-not $deviceName) { $deviceName = $extra.Hostname }
    if (-not $deviceName) { $deviceName = 'Unknown' }
    $osVersion = $extra.OSVersion
    if (-not $osVersion) { $osVersion = $extra.OSName }
    if (-not $osVersion) { $osVersion = $wu.Info['OS Version'] }
    if (-not $osVersion) { $osVersion = $wu.Info['OS Build'] }
    if (-not $osVersion) { $osVersion = 'Unknown' }

    $analysis = New-SedaObject @{
        SourceZipPath = $Path
        ExtractDir = $zip.ExtractDir
        ZipInfo = New-SedaObject @{ ZipPath = $zip.ZipPath; ZipName = $zip.ZipName; ZipSizeMb = $zip.ZipSizeMb; TotalFiles = @($inventory.all_files).Count; ExtractDir = $zip.ExtractDir; AllFiles = @($inventory.all_files) }
        Inventory = $inventory
        DeviceSummary = New-SedaObject @{ ComputerName = $deviceName; IPAddress = if ($extra.IPAddress) { $extra.IPAddress } else { 'Not found' }; OSVersion = $osVersion; Proxy = $extra.Proxy; LastUser = $extra.LastUser; ImeVersion = $extra.ImeVersion }
        DsReg = $dsreg
        DeviceInfo = $dsreg.DeviceInfo
        SsoInfo = $dsreg.SsoInfo
        Enrollments = $enrollments.Enrollments
        EnrollmentInfo = $enrollments.Summary
        ResultsXml = $results
        Firewall = $firewall
        ImeEvents = $ime.Events
        ErrorSummary = $ime.Summary
        WindowsUpdate = $wu
        Applications = $apps
        Drivers = $drivers
        WifiProfiles = $wifi
        Hardware = New-SedaObject @{ Battery = $battery; FirewallProfiles = $firewallProfiles; Certificates = $certificates }
        MdmDiagnostics = $mdmDiag
        EventLogs = $eventLogs
        Win11Compatibility = $win11
        Compliance = $compliance
        CriticalIssues = $critical
    }
    $analysis | Add-Member -MemberType NoteProperty -Name Health -Value (Get-SedaHealthReport -Analysis $analysis)
    $analysis | Add-Member -MemberType NoteProperty -Name Insights -Value (Get-SedaInsights -Analysis $analysis)
    Write-SedaLog -Level INFO -Message "Analysis completed: $($zip.ZipName); files=$($analysis.ZipInfo.TotalFiles); IME errors=$($analysis.ErrorSummary.TotalErrors); IME warnings=$($analysis.ErrorSummary.TotalWarnings); health=$($analysis.Insights.Score)."
    return $analysis
}

function Format-SedaSummaryText {
    param([object]$Analysis)
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add(('=' * 72))
    $lines.Add("  $script:AppName - Analysis Report")
    $lines.Add("  {0:yyyy-MM-dd HH:mm:ss}" -f (Get-Date))
    $lines.Add(('=' * 72))
    $lines.Add('')
    $lines.Add("  File        : $($Analysis.ZipInfo.ZipName)")
    $lines.Add("  Size        : $($Analysis.ZipInfo.ZipSizeMb) MB")
    $lines.Add("  ZIP files   : $($Analysis.ZipInfo.TotalFiles)")
    $lines.Add('')
    $lines.Add("  Computer    : $($Analysis.DeviceSummary.ComputerName)")
    $lines.Add("  IP Address  : $($Analysis.DeviceSummary.IPAddress)")
    $lines.Add("  OS Version  : $($Analysis.DeviceSummary.OSVersion)")
    $lines.Add("  Proxy       : $($Analysis.DeviceSummary.Proxy)")
    $lines.Add("  Last User   : $($Analysis.DeviceSummary.LastUser)")
    $lines.Add("  IME Version : $($Analysis.DeviceSummary.ImeVersion)")
    $lines.Add('')
    $lines.Add('-' * 72)
    $lines.Add("Overall compliance : $($Analysis.Compliance.OverallStatus)")
    $lines.Add("  Compliant        : $($Analysis.Compliance.CompliantCount)")
    $lines.Add("  Non-compliant    : $($Analysis.Compliance.NonCompliantCount)")
    $lines.Add("  Pending          : $($Analysis.Compliance.PendingCount)")
    $lines.Add("  Unknown          : $($Analysis.Compliance.UnknownCount)")
    $lines.Add('')
    $lines.Add("IME errors         : $($Analysis.ErrorSummary.ErrorCount)")
    $lines.Add("IME warnings       : $($Analysis.ErrorSummary.WarningCount)")
    $lines.Add("IME files scanned  : $($Analysis.ErrorSummary.ScannedFiles)")
    $lines.Add('')
    $lines.Add("Health score       : $($Analysis.Insights.Score)/100 ($($Analysis.Insights.Status))")
    $lines.Add('')
    if ($Analysis.CriticalIssues.Count -gt 0) {
        $lines.Add('-' * 72)
        $lines.Add('Critical issues:')
        foreach ($issue in $Analysis.CriticalIssues) {
            $lines.Add("  [$($issue.Severity)] [$($issue.Category)] $($issue.Title)")
            if ($issue.Detail) { $lines.Add("       $($issue.Detail)") }
            if ($issue.Recommendation) { $lines.Add("       => $($issue.Recommendation)") }
        }
        $lines.Add('')
    }
    if ($Analysis.Insights.TopActions.Count -gt 0) {
        $lines.Add('-' * 72)
        $lines.Add('Top actions:')
        foreach ($action in $Analysis.Insights.TopActions) {
            $lines.Add("  [$($action.Severity)] $($action.Title)")
            if ($action.Recommendation) { $lines.Add("       => $($action.Recommendation)") }
        }
        $lines.Add('')
    }
    if ($Analysis.ErrorSummary.ThemeCounts.Count -gt 0) {
        $lines.Add('-' * 72)
        $lines.Add('IME errors/warnings by theme:')
        foreach ($theme in $script:ImeThemes) {
            $count = $Analysis.ErrorSummary.ThemeCounts[$theme]
            if ($count -and ($count.Errors + $count.Warnings) -gt 0) {
                $label = if ($script:ImeThemeLabels[$theme]) { $script:ImeThemeLabels[$theme] } else { $theme }
                $lines.Add(('  {0,-34} {1,3} errors  {2,3} warnings' -f $label, $count.Errors, $count.Warnings))
            }
        }
        $lines.Add('')
    }
    return ($lines -join [Environment]::NewLine)
}

function ConvertTo-SedaTextTable {
    param(
        [object[]]$Rows,
        [string[]]$Properties
    )
    if (-not $Rows -or $Rows.Count -eq 0) { return '(no data)' }
    return ($Rows | Select-Object -Property $Properties | Format-Table -AutoSize | Out-String -Width 240).Trim()
}

function Export-SedaHtmlReport {
    param(
        [Parameter(Mandatory)][object]$Analysis,
        [Parameter(Mandatory)][string]$Path
    )
    function ConvertTo-SedaHtmlEncoded([object]$value) {
        return [System.Net.WebUtility]::HtmlEncode([string]$value)
    }

    function New-SedaHtmlBadge([string]$text) {
        $class = switch -Regex ($text) {
            'ERROR|NON_COMPLIANT|FAILED|Critical' { 'err'; break }
            'WARN|WARNING|PENDING|Attention|Degraded' { 'warn'; break }
            'OK|COMPLIANT|Healthy' { 'ok'; break }
            default { 'info' }
        }
        return "<span class='badge $class'>$(ConvertTo-SedaHtmlEncoded $text)</span>"
    }

    function New-SedaHtmlTable([object[]]$rows, [string[]]$properties) {
        if (-not $rows -or $rows.Count -eq 0) { return "<div class='alert'>No data.</div>" }
        $head = ($properties | ForEach-Object { "<th>$(ConvertTo-SedaHtmlEncoded $_)</th>" }) -join ''
        $body = foreach ($row in $rows) {
            $cells = foreach ($property in $properties) {
                $value = $row.$property
                if ($property -match 'Severity|Status|Level') { "<td>$(New-SedaHtmlBadge ([string]$value))</td>" } else { "<td>$(ConvertTo-SedaHtmlEncoded $value)</td>" }
            }
            "<tr>$($cells -join '')</tr>"
        }
        return "<table><thead><tr>$head</tr></thead><tbody>$($body -join "`n")</tbody></table>"
    }

    $css = @'
body{font-family:Segoe UI,Arial,sans-serif;background:#f4f7fb;color:#102033;margin:0}
header{background:#071527;color:#fff;padding:24px 32px;border-bottom:4px solid #3aa0ff}
header h1{margin:0;font-size:24px} header p{margin:6px 0 0;color:#b8c7dc}
main{padding:24px 32px;max-width:1240px}
section{background:#fff;border:1px solid #d9e4f2;border-radius:8px;margin:0 0 18px;padding:16px 18px;box-shadow:0 1px 4px rgba(0,0,0,.05)}
h2{font-size:18px;margin:0 0 12px;color:#1a6bb5}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px}
.kpi{background:#f8fbff;border-left:4px solid #3aa0ff;padding:12px;border-radius:6px}
.kpi strong{display:block;font-size:22px}.muted{color:#607086}
table{border-collapse:collapse;width:100%;font-size:13px}th{background:#e8f0fb;text-align:left;padding:8px;border-bottom:2px solid #d9e4f2}td{padding:7px;border-bottom:1px solid #edf2f7;vertical-align:top}
.badge{display:inline-block;color:#fff;border-radius:14px;padding:2px 8px;font-weight:700;font-size:11px}.err{background:#d64545}.warn{background:#b7791f}.ok{background:#2f9e5b}.info{background:#3a7bd5}
pre{background:#071527;color:#dfe8f4;padding:12px;border-radius:6px;overflow:auto;white-space:pre-wrap}
.alert{background:#eff6ff;border-left:4px solid #3aa0ff;padding:10px;border-radius:6px}
'@

    $html = @"
<!doctype html>
<html>
<head>
<meta charset="utf-8">
<title>$(ConvertTo-SedaHtmlEncoded $script:AppName) - $(ConvertTo-SedaHtmlEncoded $Analysis.ZipInfo.ZipName)</title>
<style>$css</style>
</head>
<body>
<header>
<h1>$(ConvertTo-SedaHtmlEncoded $script:AppName)</h1>
<p>Source: <strong>$(ConvertTo-SedaHtmlEncoded $Analysis.ZipInfo.ZipName)</strong> | Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm')</p>
</header>
<main>
<section>
<h2>Overview</h2>
<div class="grid">
<div class="kpi"><span class="muted">Health Score</span><strong>$(ConvertTo-SedaHtmlEncoded "$($Analysis.Insights.Score)/100")</strong>$(New-SedaHtmlBadge $Analysis.Insights.Status)</div>
<div class="kpi"><span class="muted">Compliance</span><strong>$(ConvertTo-SedaHtmlEncoded $Analysis.Compliance.OverallStatus)</strong></div>
<div class="kpi"><span class="muted">IME Errors</span><strong>$(ConvertTo-SedaHtmlEncoded $Analysis.ErrorSummary.ErrorCount)</strong></div>
<div class="kpi"><span class="muted">IME Warnings</span><strong>$(ConvertTo-SedaHtmlEncoded $Analysis.ErrorSummary.WarningCount)</strong></div>
<div class="kpi"><span class="muted">Files</span><strong>$(ConvertTo-SedaHtmlEncoded $Analysis.ZipInfo.TotalFiles)</strong></div>
</div>
</section>
<section>
<h2>Device</h2>
$(New-SedaHtmlTable @($Analysis.DeviceSummary) @('ComputerName','IPAddress','OSVersion','Proxy','LastUser','ImeVersion'))
</section>
<section>
<h2>Top Actions</h2>
$(New-SedaHtmlTable $Analysis.Insights.TopActions @('Severity','Title','Detail','Recommendation','Source'))
</section>
<section>
<h2>Critical Issues</h2>
$(New-SedaHtmlTable $Analysis.CriticalIssues @('Severity','Category','Title','Detail','Recommendation','Source'))
</section>
<section>
<h2>Compliance</h2>
$(New-SedaHtmlTable $Analysis.Compliance.PolicyStatuses @('Area','Status','Details','SourceFile'))
</section>
<section>
<h2>Windows Update</h2>
<pre>$(ConvertTo-SedaHtmlEncoded (ConvertTo-SedaKeyValueText $Analysis.WindowsUpdate.Info))</pre>
$(New-SedaHtmlTable $Analysis.WindowsUpdate.ReportingEvents @('Level','Timestamp','Source','ErrorCode','Message','File'))
<h3>ETL Scan</h3>
<p>$(ConvertTo-SedaHtmlEncoded $Analysis.WindowsUpdate.EtlStatus)</p>
<p>Generated log: $(ConvertTo-SedaHtmlEncoded $Analysis.WindowsUpdate.GeneratedLogPath)</p>
$(New-SedaHtmlTable (@($Analysis.WindowsUpdate.EtlEvents) | Select-Object -First 1000) @('Level','Timestamp','EventId','Source','ErrorCode','Message','EtlFile'))
</section>
<section>
<h2>IME Log Events</h2>
$(New-SedaHtmlTable (@($Analysis.ImeEvents) | Select-Object -First 1000) @('Severity','Timestamp','Theme','Category','ErrorCode','KnownCode','Message','SourceFile'))
</section>
<section>
<h2>Installed Applications</h2>
$(New-SedaHtmlTable $Analysis.Applications @('Name','Version','Publisher','InstallDate','Arch'))
</section>
<section>
<h2>Drivers</h2>
$(New-SedaHtmlTable $Analysis.Drivers @('OriginalName','Provider','ClassName','DriverVersion','Signer'))
</section>
<section>
<h2>Hardware &amp; Security</h2>
<h3>Battery</h3>
$(New-SedaHtmlTable @($Analysis.Hardware.Battery) @('ComputerName','SystemProduct','BIOS','DesignCapacity','FullChargeCapacity','HealthPct','CycleCount'))
<h3>Firewall Profiles</h3>
$(New-SedaHtmlTable $Analysis.Hardware.FirewallProfiles @('Name','State','FirewallPolicy','RemoteManagement','LogAllowed','LogDropped','LogFileName'))
<h3>Certificates</h3>
$(New-SedaHtmlTable (@($Analysis.Hardware.Certificates) | Select-Object -First 500) @('Status','Subject','Issuer','NotAfter','DaysToExpiry','Store'))
</section>
<section>
<h2>Health Findings</h2>
$(New-SedaHtmlTable $Analysis.Health.Findings @('Area','Severity','Title','Details','Recommendation','Source'))
</section>
<section>
<h2>MDM Diagnostics</h2>
<p>$(ConvertTo-SedaHtmlEncoded $Analysis.MdmDiagnostics.CabStatus)</p>
<h3>Device Info</h3>
<pre>$(ConvertTo-SedaHtmlEncoded (ConvertTo-SedaKeyValueText $Analysis.MdmDiagnostics.DeviceInfo))</pre>
<h3>Connection Info</h3>
<pre>$(ConvertTo-SedaHtmlEncoded (ConvertTo-SedaKeyValueText $Analysis.MdmDiagnostics.ConnectionInfo))</pre>
<h3>Managed Policies</h3>
$(New-SedaHtmlTable (@($Analysis.MdmDiagnostics.ManagedPolicies) | Select-Object -First 1000) @('Area','Policy','Value'))
<h3>Blocked GPOs</h3>
$(New-SedaHtmlTable $Analysis.MdmDiagnostics.BlockedGps @('Path','Name'))
<h3>MDM Diagnostic Issues</h3>
$(New-SedaHtmlTable $Analysis.MdmDiagnostics.Issues @('Severity','Area','Title','Detail','Recommendation'))
</section>
<section>
<h2>Event Logs</h2>
$(New-SedaHtmlTable $Analysis.EventLogs.Summary @('Log','Critical','Error','Warning','Information','Other','Scanned'))
$(New-SedaHtmlTable (@($Analysis.EventLogs.Events) | Where-Object { $_.Level -in @('Critical','Error','Warning') } | Select-Object -First 1000) @('Level','TimeCreated','Id','Provider','Message','File'))
</section>
<section>
<h2>Windows 11 Upgrade Indicators</h2>
$(New-SedaHtmlTable $Analysis.Win11Compatibility.Indicators @('TargetVersion','IsBlocking','ReasonText','SourceFile'))
</section>
</main>
</body>
</html>
"@

    $dir = Split-Path -Parent $Path
    if ($dir) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    [System.IO.File]::WriteAllText($Path, $html, [System.Text.Encoding]::UTF8)
    Write-SedaLog -Level INFO -Message "HTML report exported: $Path"
    return $Path
}

function Export-SedaAnonymizedZip {
    param(
        [Parameter(Mandatory)][string]$SourceZip,
        [Parameter(Mandatory)][string]$OutputZip
    )
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $textExtensions = @('.csv', '.html', '.json', '.log', '.md', '.reg', '.txt', '.xml')
    $outDir = Split-Path -Parent $OutputZip
    if ($outDir) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    if (Test-Path -LiteralPath $OutputZip) { Remove-Item -LiteralPath $OutputZip -Force }

    $src = [System.IO.Compression.ZipFile]::OpenRead($SourceZip)
    $dst = [System.IO.Compression.ZipFile]::Open($OutputZip, [System.IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($entry in $src.Entries) {
            $newEntry = $dst.CreateEntry($entry.FullName, [System.IO.Compression.CompressionLevel]::Optimal)
            $inStream = $entry.Open()
            $outStream = $newEntry.Open()
            try {
                $ext = [System.IO.Path]::GetExtension($entry.FullName).ToLowerInvariant()
                if ($textExtensions -contains $ext) {
                    $ms = New-Object System.IO.MemoryStream
                    $inStream.CopyTo($ms)
                    $bytes = $ms.ToArray()
                    $text = [System.Text.Encoding]::UTF8.GetString($bytes)
                    if (-not $text) { $text = [System.Text.Encoding]::Unicode.GetString($bytes) }
                    $text = $text -replace '\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b', '<EMAIL>'
                    $text = $text -replace '\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b', '<GUID>'
                    $text = $text -replace '\b(?:\d{1,3}\.){3}\d{1,3}\b', '<IPV4>'
                    $text = $text -replace '(?i)\b(C:\\Users\\)[^\\\r\n]+', '$1<USER>'
                    $text = $text -replace '(?i)\b(TenantId|Tenant ID|DeviceId|Device ID|SerialNumber|Serial Number)\s*[:=]\s*([^\r\n,;]+)', '$1=<REDACTED>'
                    $text = $text -replace '(?i)\b(UPN|UserPrincipalName|Email|PrimaryUser)\s*[:=]\s*([^\r\n,;]+)', '$1=<REDACTED>'
                    $outBytes = [System.Text.Encoding]::UTF8.GetBytes($text)
                    $outStream.Write($outBytes, 0, $outBytes.Length)
                } else {
                    $inStream.CopyTo($outStream)
                }
            } finally {
                $outStream.Dispose()
                $inStream.Dispose()
            }
        }
    } finally {
        $dst.Dispose()
        $src.Dispose()
    }
    Write-SedaLog -Level INFO -Message "Anonymized ZIP exported: $OutputZip"
    return $OutputZip
}

function Invoke-SedaLocalCollection {
    param([string]$ZipPath)
    Write-SedaLog -Level INFO -Message "Local collection started. Target ZIP: $ZipPath"
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('seda_collect_' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Path $tmp -Force | Out-Null
    try {
        $script:SedaCollectCounter = 0
        function next-name([string]$Kind, [string]$Name, [string]$Ext) {
            $script:SedaCollectCounter++
            return Join-Path $tmp ('({0}) {1} {2}.{3}' -f $script:SedaCollectCounter, $Kind, $Name, $Ext)
        }
        function write-collect-text([string]$Path, [string]$Text, [string]$Header = '') {
            $body = if ($Header) { "$Header`r`n$('=' * $Header.Length)`r`n$Text" } else { $Text }
            [System.IO.File]::WriteAllText($Path, [string]$body, [System.Text.Encoding]::UTF8)
        }
        function copy-collect-files([string]$Source, [string]$Destination, [int]$MaxFiles = 100, [int]$MaxAgeDays = 60, [string[]]$Extensions = @()) {
            if (-not (Test-Path -LiteralPath $Source)) { return 0 }
            New-Item -ItemType Directory -Path $Destination -Force | Out-Null
            $cutoff = (Get-Date).AddDays(-1 * $MaxAgeDays)
            $files = Get-ChildItem -LiteralPath $Source -Recurse -File -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -ge $cutoff -and ($Extensions.Count -eq 0 -or $Extensions -contains $_.Extension.ToLowerInvariant()) } |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First $MaxFiles
            foreach ($file in @($files)) {
                try { Copy-Item -LiteralPath $file.FullName -Destination (Join-Path $Destination $file.Name) -Force -ErrorAction Stop } catch {}
            }
            return @($files).Count
        }
        function run-collect-command([string]$Name, [string]$File, [string[]]$Args, [string]$Ext = 'log') {
            $dest = next-name 'Command' $Name $Ext
            try {
                $output = & $File @Args 2>&1 | Out-String -Width 4096
                write-collect-text -Path $dest -Text $output -Header "$File $($Args -join ' ')"
            } catch {
                write-collect-text -Path $dest -Text $_.Exception.Message -Header "$File $($Args -join ' ')"
            }
            return $dest
        }
        function run-collect-ps([string]$Name, [string]$Code) {
            $dest = Join-Path $extendedDest "$Name.txt"
            try {
                $output = & powershell.exe -NoProfile -ExecutionPolicy Bypass -Command $Code 2>&1 | Out-String -Width 4096
                write-collect-text -Path $dest -Text $output -Header $Name
            } catch {
                write-collect-text -Path $dest -Text $_.Exception.Message -Header $Name
            }
            return $dest
        }
        $registryKeys = @(
            @{ Name = 'MDM_Policy_Result'; Path = 'HKLM\SOFTWARE\Microsoft\PolicyManager\current\device' },
            @{ Name = 'MDM_Enrollment'; Path = 'HKLM\SOFTWARE\Microsoft\Enrollments' },
            @{ Name = 'MDM_DeviceManagement'; Path = 'HKLM\SOFTWARE\Microsoft\DeviceManageabilityCSP' },
            @{ Name = 'WindowsUpdate_Settings'; Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate' },
            @{ Name = 'WindowsUpdate_AU'; Path = 'HKLM\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU' },
            @{ Name = 'WindowsUpdate_Orchestrator'; Path = 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Orchestrator' },
            @{ Name = 'CloudManagedUpdate'; Path = 'HKLM\SOFTWARE\Microsoft\CloudManagedUpdate' },
            @{ Name = 'Win11_Upgrade_Compatibility_Indicators'; Path = 'HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\TargetVersionUpgradeExperienceIndicators' },
            @{ Name = 'CurrentVersion_Uninstall'; Path = 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall' },
            @{ Name = 'WOW6432Node_Uninstall'; Path = 'HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall' },
            @{ Name = 'IntuneManagementExtension'; Path = 'HKLM\SOFTWARE\Microsoft\IntuneManagementExtension' },
            @{ Name = 'DSRegCmd_State'; Path = 'HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\CDJ' }
        )
        foreach ($rk in $registryKeys) {
            $dest = next-name 'RegistryKey' ($rk.Name + ' export') 'reg'
            & reg.exe export $rk.Path $dest /y *> $null
        }
        $commands = @(
            @{ Name = 'windir_system32_Dsregcmd_exe_status output'; File = 'dsregcmd.exe'; Args = @('/status') },
            @{ Name = 'windir_system32_ipconfig_exe_all output'; File = 'ipconfig.exe'; Args = @('/all') },
            @{ Name = 'windir_system32_netsh_exe_advfirewall_show_allprofiles output'; File = 'netsh.exe'; Args = @('advfirewall','show','allprofiles') },
            @{ Name = 'windir_system32_netsh_exe_winhttp_show_proxy output'; File = 'netsh.exe'; Args = @('winhttp','show','proxy') },
            @{ Name = 'windir_system32_pnputil_exe_enum-drivers output'; File = 'pnputil.exe'; Args = @('/enum-drivers') },
            @{ Name = 'windir_system32_netsh_exe_wlan_show_profiles output'; File = 'netsh.exe'; Args = @('wlan','show','profiles') },
            @{ Name = 'windir_system32_certutil_exe_store_my output'; File = 'certutil.exe'; Args = @('-store','My') },
            @{ Name = 'windir_system32_certutil_exe_user_store_my output'; File = 'certutil.exe'; Args = @('-user','-store','My') },
            @{ Name = 'windir_system32_systeminfo_exe output'; File = 'systeminfo.exe'; Args = @() }
        )
        foreach ($cmd in $commands) {
            run-collect-command -Name $cmd.Name -File $cmd.File -Args @($cmd.Args) | Out-Null
        }
        $imeDest = Join-Path $tmp '(90) FoldersFiles ProgramData_Microsoft_IntuneManagementExtension_Logs'
        if (Test-Path 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs') {
            New-Item -ItemType Directory -Path $imeDest -Force | Out-Null
            Get-ChildItem 'C:\ProgramData\Microsoft\IntuneManagementExtension\Logs' -File -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 120 |
                Copy-Item -Destination $imeDest -Force -ErrorAction SilentlyContinue
        }
        $wuDest = Join-Path $tmp '(91) FoldersFiles windir_Logs_WindowsUpdate_etl'
        if (Test-Path 'C:\Windows\Logs\WindowsUpdate') {
            New-Item -ItemType Directory -Path $wuDest -Force | Out-Null
            Get-ChildItem 'C:\Windows\Logs\WindowsUpdate' -File -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 50 |
                Copy-Item -Destination $wuDest -Force -ErrorAction SilentlyContinue
        }
        $autopatchDest = Join-Path $tmp '(93) FoldersFiles ProgramData_Microsoft_AutopatchClient_Logs'
        copy-collect-files -Source 'C:\ProgramData\Microsoft\AutopatchClient\Logs' -Destination $autopatchDest -MaxFiles 80 -MaxAgeDays 90 -Extensions @('.log','.txt') | Out-Null
        $pantherDest = Join-Path $tmp '(94) FoldersFiles windir_Panther'
        copy-collect-files -Source 'C:\Windows\Panther' -Destination $pantherDest -MaxFiles 30 -MaxAgeDays 120 -Extensions @('.log','.xml','.etl') | Out-Null
        $miniDumpDest = Join-Path $tmp '(95) FoldersFiles windir_Minidump'
        copy-collect-files -Source 'C:\Windows\Minidump' -Destination $miniDumpDest -MaxFiles 10 -MaxAgeDays 120 -Extensions @('.dmp') | Out-Null
        try {
            $wingetRoot = Join-Path $env:LOCALAPPDATA 'Packages\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\LocalState\DiagOutputDir'
            $wingetDest = Join-Path $tmp '(96) FoldersFiles Winget_DiagOutputDir'
            copy-collect-files -Source $wingetRoot -Destination $wingetDest -MaxFiles 80 -MaxAgeDays 90 -Extensions @('.log','.txt') | Out-Null
        } catch {}
        $eventChannels = @(
            @{ Channel='Application'; Name='Application' },
            @{ Channel='System'; Name='System' },
            @{ Channel='Setup'; Name='Setup' },
            @{ Channel='Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin'; Name='MDM_Admin' },
            @{ Channel='Microsoft-Windows-AAD/Operational'; Name='AAD_Operational' }
        )
        foreach ($ev in $eventChannels) {
            $dest = next-name 'Events' $ev.Name 'evtx'
            try { & wevtutil.exe epl $ev.Channel $dest /ow:true *> $null } catch { write-collect-text -Path ($dest + '.error.txt') -Text $_.Exception.Message -Header "wevtutil $($ev.Channel)" }
        }
        $batteryPath = Join-Path $tmp 'battery-report.html'
        try { & powercfg.exe /batteryreport /output $batteryPath /duration 14 *> $null } catch { write-collect-text -Path (Join-Path $tmp 'battery-report.error.txt') -Text $_.Exception.Message -Header 'powercfg /batteryreport' }
        $wuLogDest = Join-Path $tmp '(92) FoldersFiles WindowsUpdateLog'
        New-Item -ItemType Directory -Path $wuLogDest -Force | Out-Null
        $wuLogPath = Join-Path $wuLogDest 'WindowsUpdate.generated.log'
        try {
            if (Get-Command Get-WindowsUpdateLog -ErrorAction SilentlyContinue) {
                Write-SedaLog -Level INFO -Message "Generating local WindowsUpdate.log for collection: $wuLogPath"
                Get-WindowsUpdateLog -LogPath $wuLogPath -ErrorAction Stop *> (Join-Path $wuLogDest 'Get-WindowsUpdateLog.output.txt')
            } else {
                [System.IO.File]::WriteAllText((Join-Path $wuLogDest 'Get-WindowsUpdateLog.unavailable.txt'), 'Get-WindowsUpdateLog is not available on this host.', [System.Text.Encoding]::UTF8)
            }
        } catch {
            Write-SedaLog -Level WARN -Message 'Local Get-WindowsUpdateLog generation failed.' -Exception $_.Exception
            [System.IO.File]::WriteAllText((Join-Path $wuLogDest 'Get-WindowsUpdateLog.error.txt'), $_.Exception.Message, [System.Text.Encoding]::UTF8)
        }
        $extendedDest = Join-Path $tmp 'extended'
        New-Item -ItemType Directory -Path $extendedDest -Force | Out-Null
        $psCommands = [ordered]@{
            ps_system_info = 'Get-ComputerInfo | Select CsName,OsName,OsVersion,OsBuildNumber,OsArchitecture,CsProcessors,CsTotalPhysicalMemory,OsLastBootUpTime,BiosBIOSVersion,BiosManufacturer,CsModel,CsManufacturer,HyperVisorPresent,OsLanguage,WindowsInstallationType | Format-List'
            ps_disk_usage = "Get-PSDrive -PSProvider FileSystem | Select Name,Used,Free,@{N='Total';E={`$_.Used+`$_.Free}},@{N='Free%';E={if(`$_.Used+`$_.Free -gt 0){[math]::Round(`$_.Free/(`$_.Used+`$_.Free)*100,1)}}} | Format-Table -AutoSize"
            ps_pending_reboot = "`$regs=@('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending','HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired','HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager','HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootInProgress'); foreach(`$r in `$regs){ `$n=`$r.Split('\')[-1]; `$e=Test-Path `$r; if(`$r -like '*Session Manager'){ `$pfr=(Get-ItemProperty `$r -EA SilentlyContinue).PendingFileRenameOperations; `$e=@(`$pfr).Count -gt 0 }; [PSCustomObject]@{Key=`$n;PendingReboot=`$e} } | Format-Table -AutoSize"
            ps_bitlocker_status = "try { Get-BitLockerVolume | Select MountPoint,VolumeStatus,ProtectionStatus,EncryptionPercentage,VolumeType,KeyProtector | Format-List } catch { 'Get-BitLockerVolume not available. Trying manage-bde:'; & manage-bde -status 2>&1 }"
            ps_defender_status = "try { Get-MpComputerStatus | Select AMRunningMode,AntivirusEnabled,AntispywareEnabled,RealTimeProtectionEnabled,OnAccessProtectionEnabled,IoavProtectionEnabled,BehaviorMonitorEnabled,AntivirusSignatureLastUpdated,AntispywareSignatureLastUpdated,FullScanEndTime,QuickScanEndTime,AMProductVersion,AMEngineVersion | Format-List } catch { 'Windows Defender info unavailable: ' + `$_.Exception.Message }"
            ps_tpm_status = "try { Get-Tpm | Select TpmPresent,TpmReady,TpmEnabled,TpmActivated,TpmOwned,ManagedAuthLevel,ManufacturerId,ManufacturerVersion,ManufacturerVersionFull20,SpecVersion | Format-List } catch { 'TPM info unavailable: ' + `$_.Exception.Message }"
            ps_secureboot = "try { `$sb = Confirm-SecureBootUEFI; [PSCustomObject]@{SecureBootEnabled=`$sb} | Format-List } catch { 'Secure Boot check: ' + `$_.Exception.Message }"
            ps_top_processes = 'Get-Process | Sort-Object CPU -Descending | Select-Object -First 40 Name,Id,CPU,WorkingSet,SessionId,StartTime,Path | Format-Table -AutoSize'
            ps_services_abnormal = "Get-Service | Where-Object { `$_.StartType -eq 'Automatic' -and `$_.Status -ne 'Running' } | Select-Object Name,DisplayName,Status,StartType | Sort-Object Name | Format-Table -AutoSize"
            ps_services_all = 'Get-Service | Select Name,DisplayName,Status,StartType | Sort-Object StartType,Status,Name | Format-Table -AutoSize'
            ps_timesync = "w32tm /query /status; '---'; w32tm /query /configuration"
            ps_user_profiles = "Get-CimInstance Win32_UserProfile | Select LocalPath,SID,LastUseTime,Special | Sort-Object LastUseTime -Descending | Format-Table -AutoSize"
            ps_hotfixes = 'Get-HotFix | Select HotFixID,InstalledOn,Description,InstalledBy | Sort-Object InstalledOn -Descending | Format-Table -AutoSize'
            ps_network_adapters = "Get-NetAdapter | Select Name,InterfaceDescription,Status,MacAddress,LinkSpeed | Format-Table -AutoSize; '---'; Get-NetIPConfiguration | Format-List"
            ps_dns_config = 'Get-DnsClientServerAddress | Where-Object AddressFamily -eq 2 | Select InterfaceAlias,ServerAddresses | Format-Table -AutoSize'
            ps_open_ports = "Get-NetTCPConnection -State Listen | Select LocalAddress,LocalPort,State,@{N='Process';E={(Get-Process -Id `$_.OwningProcess -EA SilentlyContinue).Name}} | Sort-Object LocalPort | Format-Table -AutoSize"
            ps_scheduled_tasks_abnormal = "Get-ScheduledTask | Where-Object { `$_.State -eq 'Disabled' -or (`$_.TaskPath -notlike '\Microsoft\*' -and `$_.State -eq 'Ready') } | Select TaskName,TaskPath,State | Sort-Object TaskPath | Format-Table -AutoSize"
            ps_startup_programs = 'Get-CimInstance Win32_StartupCommand | Select Name,Command,Location,User | Format-Table -AutoSize'
            ps_powershell_policy = "Get-ExecutionPolicy -List | Format-Table -AutoSize; '---Constrained Language Mode:'; `$ExecutionContext.SessionState.LanguageMode"
            ps_windows_activation = "Get-CimInstance SoftwareLicensingProduct | Where-Object { `$_.PartialProductKey -and `$_.ApplicationId -eq '55c92734-d682-4d71-983e-d6ec3f16059f' } | Select Name,LicenseStatus,PartialProductKey,@{N='Status';E={switch(`$_.LicenseStatus){1{'Licensed'};2{'OOBGrace'};3{'OOTGrace'};4{'NonGenuine'};5{'Notification'};6{'ExtendedGrace'};default{'Unknown'}}}} | Format-List"
            ps_shared_folders = 'Get-SmbShare | Select Name,Path,Description,ShareState | Format-Table -AutoSize'
            ps_recent_errors = 'Get-EventLog -LogName System -EntryType Error -Newest 50 2>$null | Select TimeGenerated,Source,EventID,Message | Format-Table -AutoSize -Wrap'
            ps_proxy_config = "netsh winhttp show proxy; '---IE/System Proxy:'; Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' | Select ProxyEnable,ProxyServer,ProxyOverride | Format-List"
            ps_store_apps = "try { Get-AppxPackage -AllUsers | Select Name,Version,Publisher,Architecture,InstallLocation,PackageUserInformation | Sort-Object Name | Format-Table -AutoSize } catch { 'Store apps unavailable: ' + `$_.Exception.Message }"
            ps_update_history = "`$Session=New-Object -ComObject Microsoft.Update.Session; `$Searcher=`$Session.CreateUpdateSearcher(); `$Count=`$Searcher.GetTotalHistoryCount(); `$History=`$Searcher.QueryHistory(0,[math]::Min(`$Count,50)); `$History | Select Date,Title,@{N='Result';E={switch(`$_.ResultCode){1{'InProgress'};2{'Succeeded'};3{'SucceededWithErrors'};4{'Failed'};5{'Aborted'};default{`$_.ResultCode}}}} | Format-Table -AutoSize"
            ps_intune_enrollment = "dsregcmd /status; '---'; Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Enrollments' -EA SilentlyContinue | ForEach-Object { Get-ItemProperty `$_.PSPath -EA SilentlyContinue } | Where-Object { `$_.EnrollmentType } | Select PSChildName,EnrollmentType,UPN,DiscoveryServiceFullURL | Format-List"
        }
        foreach ($entry in $psCommands.GetEnumerator()) {
            run-collect-ps -Name $entry.Key -Code $entry.Value | Out-Null
        }
        try {
            $mdmDiag = Join-Path $env:WINDIR 'System32\MDMDiagnosticsTool.exe'
            if (Test-Path -LiteralPath $mdmDiag -PathType Leaf) {
                $mdmOut = Join-Path $tmp 'mdmdiag'
                New-Item -ItemType Directory -Path $mdmOut -Force | Out-Null
                $proc = Start-Process -FilePath $mdmDiag -ArgumentList @('-out', $mdmOut) -WindowStyle Hidden -Wait -PassThru -ErrorAction Stop
                write-collect-text -Path (Join-Path $extendedDest 'mdmdiag_status.txt') -Text "ExitCode=$($proc.ExitCode)`r`nOutput=$mdmOut" -Header 'MDMDiagnosticsTool'
            }
        } catch {
            write-collect-text -Path (Join-Path $extendedDest 'mdmdiag_error.txt') -Text $_.Exception.Message -Header 'MDMDiagnosticsTool'
        }
        try {
            $gpHtml = Join-Path $extendedDest 'gpresult.html'
            $gpProc = Start-Process -FilePath 'gpresult.exe' -ArgumentList @('/H', $gpHtml, '/F') -WindowStyle Hidden -Wait -PassThru -ErrorAction Stop
            if (-not (Test-Path -LiteralPath $gpHtml -PathType Leaf)) {
                run-collect-command -Name 'windir_system32_gpresult_exe_z output' -File 'gpresult.exe' -Args @('/Z') | Out-Null
            } else {
                write-collect-text -Path (Join-Path $extendedDest 'gpresult_status.txt') -Text "ExitCode=$($gpProc.ExitCode)" -Header 'gpresult'
            }
        } catch {
            write-collect-text -Path (Join-Path $extendedDest 'gpresult_error.txt') -Text $_.Exception.Message -Header 'gpresult'
        }
        $readinessScript = Join-Path $script:BasePath 'HardwareReadiness.ps1'
        $readinessJson = Join-Path $extendedDest 'win11_readiness.json'
        if (Test-Path -LiteralPath $readinessScript -PathType Leaf) {
            try {
                Write-SedaLog -Level INFO -Message "Running HardwareReadiness.ps1 for local collection: $readinessScript"
                $raw = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $readinessScript 2>&1 | Out-String -Width 4096
                $jsonLine = @($raw -split "`r?`n" | Where-Object { $_.Trim().StartsWith('{') -and $_ -match 'returnCode' } | Select-Object -Last 1)[0]
                if ($jsonLine) {
                    [System.IO.File]::WriteAllText($readinessJson, $jsonLine.Trim(), [System.Text.Encoding]::UTF8)
                } else {
                    [System.IO.File]::WriteAllText((Join-Path $extendedDest 'win11_readiness_raw.txt'), $raw, [System.Text.Encoding]::UTF8)
                    [System.IO.File]::WriteAllText($readinessJson, '{"returnCode":-1,"returnResult":"UNDETERMINED","returnReason":"No JSON output from HardwareReadiness.ps1","logging":""}', [System.Text.Encoding]::UTF8)
                }
            } catch {
                Write-SedaLog -Level WARN -Message 'HardwareReadiness.ps1 execution failed.' -Exception $_.Exception
                [System.IO.File]::WriteAllText($readinessJson, ('{"returnCode":-2,"returnResult":"FAILED TO RUN","returnReason":"' + ($_.Exception.Message -replace '\\','\\' -replace '"','\"') + '","logging":""}'), [System.Text.Encoding]::UTF8)
            }
        } else {
            [System.IO.File]::WriteAllText($readinessJson, '{"returnCode":-2,"returnResult":"FAILED TO RUN","returnReason":"HardwareReadiness.ps1 not found","logging":""}', [System.Text.Encoding]::UTF8)
        }
        $resultsPath = Join-Path $tmp 'results.xml'
        $xmlLines = New-Object System.Collections.Generic.List[string]
        $xmlLines.Add('<DiagnosticsResults generated="' + [System.Security.SecurityElement]::Escape((Get-Date).ToString('s')) + '">')
        foreach ($file in @(Get-ChildItem -LiteralPath $tmp -Recurse -File -Force -ErrorAction SilentlyContinue | Where-Object { $_.FullName -ne $resultsPath } | Sort-Object FullName)) {
            $relative = $file.FullName.Substring($tmp.Length).TrimStart('\', '/')
            $hresult = if ($file.Name -like '*.error.txt' -or $file.Name -like '*.unavailable.txt') { '-2147418113' } else { '0' }
            $node = if ($hresult -eq '0') { 'CollectedFile' } else { 'CollectionError' }
            $xmlLines.Add(('  <{0}><Name>{1}</Name><HRESULT>{2}</HRESULT></{0}>' -f $node, [System.Security.SecurityElement]::Escape($relative), $hresult))
        }
        $xmlLines.Add('</DiagnosticsResults>')
        [System.IO.File]::WriteAllText($resultsPath, ($xmlLines -join [Environment]::NewLine), [System.Text.Encoding]::UTF8)
        if (-not $ZipPath) {
            $hostName = $env:COMPUTERNAME
            $ZipPath = Join-Path ([Environment]::GetFolderPath('Desktop')) ("{0}_DiagLogs_{1:yyyyMMdd_HHmmss}.zip" -f $hostName, (Get-Date))
        }
        if (Test-Path -LiteralPath $ZipPath) { Remove-Item -LiteralPath $ZipPath -Force }
        [System.IO.Compression.ZipFile]::CreateFromDirectory($tmp, $ZipPath)
        Write-SedaLog -Level INFO -Message "Local collection completed: $ZipPath"
        return $ZipPath
    } finally {
        Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
    }
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

function Start-SedaGui {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName WindowsBase
    Add-Type -AssemblyName System.Windows.Forms

    $splash = $null
    $splashHelperPath = Import-SmartM365GuiSplash
    if ($splashHelperPath) {
        . $splashHelperPath
        $splash = Start-SmartM365GuiSplash -Framework Wpf -ProductName 'Endpoint Diagnostics Analyzer'
    }

    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="$script:AppName" Width="1500" Height="920" MinWidth="1100" MinHeight="720"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResize"
        Background="#F4F8FB"
        FontFamily="Segoe UI">
  <Window.Resources>
    <SolidColorBrush x:Key="InkBrush" Color="#102033"/>
    <SolidColorBrush x:Key="MutedBrush" Color="#607086"/>
    <SolidColorBrush x:Key="PanelBrush" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="PanelSoftBrush" Color="#FAFCFE"/>
    <SolidColorBrush x:Key="PanelAltBrush" Color="#F6FAFD"/>
    <SolidColorBrush x:Key="PanelBorderBrush" Color="#D7E1EA"/>
    <SolidColorBrush x:Key="HeaderBrush" Color="#071D33"/>
    <SolidColorBrush x:Key="AccentBrush" Color="#00A9E0"/>
    <SolidColorBrush x:Key="AccentDarkBrush" Color="#0078A6"/>
    <SolidColorBrush x:Key="SoftAccentBrush" Color="#E8F7FC"/>
    <SolidColorBrush x:Key="GreenBrush" Color="#16884A"/>
    <SolidColorBrush x:Key="PurpleBrush" Color="#0078D4"/>

    <Style TargetType="Button">
      <Setter Property="Margin" Value="0,0,8,0"/>
      <Setter Property="Padding" Value="13,8"/>
      <Setter Property="MinWidth" Value="96"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="ToolTipService.ShowOnDisabled" Value="True"/>
      <Setter Property="Background" Value="{StaticResource PanelBrush}"/>
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
                <Setter TargetName="ButtonBorder" Property="Background" Value="#E8F7FC"/>
                <Setter TargetName="ButtonBorder" Property="BorderBrush" Value="{StaticResource AccentBrush}"/>
              </Trigger>
              <Trigger Property="IsPressed" Value="True">
                <Setter TargetName="ButtonBorder" Property="Background" Value="#D3EFF8"/>
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
    <Style x:Key="SecondaryButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Background" Value="{StaticResource SoftAccentBrush}"/>
      <Setter Property="Foreground" Value="{StaticResource AccentBrush}"/>
      <Setter Property="BorderBrush" Value="#A7DCEC"/>
    </Style>
    <Style x:Key="SuccessButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}">
      <Setter Property="Background" Value="{StaticResource GreenBrush}"/>
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="BorderBrush" Value="{StaticResource GreenBrush}"/>
    </Style>
    <Style x:Key="TextSurface" TargetType="TextBox">
      <Setter Property="FontFamily" Value="Consolas"/>
      <Setter Property="FontSize" Value="13"/>
      <Setter Property="IsReadOnly" Value="True"/>
      <Setter Property="AcceptsReturn" Value="True"/>
      <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
      <Setter Property="HorizontalScrollBarVisibility" Value="Auto"/>
      <Setter Property="Background" Value="{StaticResource PanelSoftBrush}"/>
      <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
      <Setter Property="BorderBrush" Value="#E4EAF1"/>
      <Setter Property="Padding" Value="12"/>
    </Style>
    <Style TargetType="TabControl">
      <Setter Property="Background" Value="{StaticResource PanelBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource PanelBorderBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="Padding" Value="0"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabControl">
            <Grid KeyboardNavigation.TabNavigation="Local">
              <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/>
                <RowDefinition Height="*"/>
              </Grid.RowDefinitions>
              <TabPanel x:Name="HeaderPanel"
                        Grid.Row="0"
                        IsItemsHost="True"
                        KeyboardNavigation.TabIndex="1"
                        Panel.ZIndex="1"
                        Margin="0,0,0,-1"
                        Background="{StaticResource PanelBrush}"/>
              <Border x:Name="ContentPanel"
                      Grid.Row="1"
                      Background="{StaticResource PanelBrush}"
                      BorderBrush="{TemplateBinding BorderBrush}"
                      BorderThickness="{TemplateBinding BorderThickness}">
                <ContentPresenter x:Name="PART_SelectedContentHost"
                                  ContentSource="SelectedContent"
                                  Margin="{TemplateBinding Padding}"
                                  SnapsToDevicePixels="{TemplateBinding SnapsToDevicePixels}"/>
              </Border>
            </Grid>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="TabItem">
      <Setter Property="Background" Value="#F6FAFD"/>
      <Setter Property="BorderBrush" Value="#C9D6E2"/>
      <Setter Property="Foreground" Value="{StaticResource MutedBrush}"/>
      <Setter Property="Padding" Value="12,7"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="MinHeight" Value="30"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="TabItem">
            <Grid SnapsToDevicePixels="True">
              <Border x:Name="TabBorder"
                      Background="{TemplateBinding Background}"
                      BorderBrush="{TemplateBinding BorderBrush}"
                      BorderThickness="1"
                      Margin="0,0,0,0">
                <ContentPresenter x:Name="ContentSite"
                                  ContentSource="Header"
                                  HorizontalAlignment="Center"
                                  VerticalAlignment="Center"
                                  Margin="{TemplateBinding Padding}"
                                  RecognizesAccessKey="True"/>
              </Border>
              <Border x:Name="SelectedLip"
                      Height="2"
                      VerticalAlignment="Bottom"
                      Background="{StaticResource AccentBrush}"
                      Visibility="Collapsed"/>
            </Grid>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="TabBorder" Property="Background" Value="#E8F7FC"/>
                <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
              </Trigger>
              <Trigger Property="IsSelected" Value="True">
                <Setter TargetName="TabBorder" Property="Background" Value="{StaticResource PanelBrush}"/>
                <Setter TargetName="TabBorder" Property="BorderBrush" Value="{StaticResource PanelBorderBrush}"/>
                <Setter TargetName="SelectedLip" Property="Visibility" Value="Visible"/>
                <Setter Property="Foreground" Value="{StaticResource PurpleBrush}"/>
                <Setter Property="Panel.ZIndex" Value="2"/>
              </Trigger>
              <Trigger Property="IsEnabled" Value="False">
                <Setter Property="Opacity" Value="0.45"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="DataGrid">
      <Setter Property="Background" Value="{StaticResource PanelSoftBrush}"/>
      <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
      <Setter Property="BorderBrush" Value="{StaticResource PanelBorderBrush}"/>
      <Setter Property="BorderThickness" Value="1"/>
      <Setter Property="GridLinesVisibility" Value="None"/>
      <Setter Property="HeadersVisibility" Value="Column"/>
      <Setter Property="RowBackground" Value="#FFFFFF"/>
      <Setter Property="AlternatingRowBackground" Value="#F6FAFD"/>
      <Setter Property="AutoGenerateColumns" Value="True"/>
      <Setter Property="IsReadOnly" Value="True"/>
      <Setter Property="CanUserAddRows" Value="False"/>
      <Setter Property="CanUserDeleteRows" Value="False"/>
      <Setter Property="SelectionMode" Value="Single"/>
      <Setter Property="FontFamily" Value="Consolas"/>
      <Setter Property="FontSize" Value="12"/>
      <Setter Property="HorizontalGridLinesBrush" Value="#E2EAF1"/>
      <Setter Property="VerticalGridLinesBrush" Value="#E2EAF1"/>
      <Setter Property="EnableRowVirtualization" Value="True"/>
      <Setter Property="EnableColumnVirtualization" Value="True"/>
      <Setter Property="VirtualizingPanel.IsVirtualizing" Value="True"/>
      <Setter Property="VirtualizingPanel.VirtualizationMode" Value="Recycling"/>
      <Setter Property="ScrollViewer.CanContentScroll" Value="True"/>
      <Setter Property="MaxColumnWidth" Value="520"/>
    </Style>
    <Style TargetType="DataGridColumnHeader">
      <Setter Property="Background" Value="#071D33"/>
      <Setter Property="Foreground" Value="#FFFFFF"/>
      <Setter Property="BorderBrush" Value="#183A5B"/>
      <Setter Property="Padding" Value="6,4"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
    </Style>
    <Style TargetType="DataGridCell">
      <Setter Property="BorderThickness" Value="0"/>
      <Setter Property="Padding" Value="4,2"/>
      <Setter Property="Foreground" Value="{StaticResource InkBrush}"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="DataGridCell">
            <Border Background="{TemplateBinding Background}" Padding="{TemplateBinding Padding}">
              <ContentPresenter VerticalAlignment="Center"/>
            </Border>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
    <Style TargetType="TextBlock" x:Key="SectionTitle">
      <Setter Property="Foreground" Value="{StaticResource AccentBrush}"/>
      <Setter Property="FontWeight" Value="SemiBold"/>
      <Setter Property="FontSize" Value="18"/>
      <Setter Property="Margin" Value="0,0,0,8"/>
    </Style>
    <Style TargetType="ProgressBar">
      <Setter Property="Foreground" Value="{StaticResource AccentBrush}"/>
      <Setter Property="Background" Value="#E5EAF0"/>
    </Style>
    <Style x:Key="OverlayProgress" TargetType="ProgressBar" BasedOn="{StaticResource {x:Type ProgressBar}}">
      <Setter Property="Height" Value="10"/>
      <Setter Property="IsIndeterminate" Value="True"/>
    </Style>
  </Window.Resources>

  <Grid Margin="16">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <Border Grid.Row="0" Background="{StaticResource HeaderBrush}" BorderBrush="#0BA6D6" BorderThickness="0,0,0,4" CornerRadius="0" Padding="26,24">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <Grid>
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
          </Grid.ColumnDefinitions>
          <Border x:Name="HeaderLogoLink" Width="56" Height="56" CornerRadius="0" BorderBrush="#24577F" BorderThickness="1" Background="#143657" Margin="0,0,16,0" VerticalAlignment="Top" Cursor="Hand" ToolTip="Open WorkplaceCloudHub.com">
            <Image x:Name="HeaderLogo" Stretch="Uniform" Margin="8" SnapsToDevicePixels="True" RenderOptions.BitmapScalingMode="HighQuality"/>
          </Border>
          <StackPanel Grid.Column="1">
          <Border HorizontalAlignment="Left" CornerRadius="12" Padding="10,4" Background="{StaticResource SoftAccentBrush}">
            <TextBlock Text="PowerShell diagnostics" Foreground="{StaticResource AccentBrush}" FontWeight="SemiBold" FontSize="12"/>
          </Border>
          <TextBlock Text="Smart Endpoint Diagnostics Analyzer" FontSize="28" FontWeight="SemiBold" Foreground="#FFFFFF" Margin="0,12,0,0"/>
          <TextBlock Text="workplacecloudhub.com - Intune and Windows endpoint diagnostics" FontSize="14" Foreground="#B8C7DC" Margin="0,6,0,0"/>
          </StackPanel>
        </Grid>
        <StackPanel Grid.Column="1" HorizontalAlignment="Right" VerticalAlignment="Top">
          <Border Background="#143657" BorderBrush="#24577F" BorderThickness="1" Padding="12,6" CornerRadius="0" HorizontalAlignment="Right">
            <TextBlock Text="v$script:AppVersion" Foreground="#D9F6FF" FontWeight="SemiBold"/>
          </Border>
          <Border Background="#143657" BorderBrush="#24577F" BorderThickness="1" Padding="12,10" CornerRadius="0" Margin="0,10,0,0" MinWidth="300">
            <StackPanel>
              <TextBlock x:Name="BannerComputer" Foreground="#FFFFFF" FontWeight="SemiBold"/>
              <TextBlock x:Name="BannerUser" Foreground="#D9F6FF" Margin="0,4,0,0" TextWrapping="NoWrap" TextTrimming="CharacterEllipsis"/>
              <TextBlock x:Name="BannerAccount" Foreground="#D9F6FF" Margin="0,4,0,0" TextWrapping="NoWrap" TextTrimming="CharacterEllipsis"/>
              <TextBlock x:Name="BannerOS" Foreground="#D9F6FF" Margin="0,4,0,0" TextWrapping="NoWrap" TextTrimming="CharacterEllipsis"/>
              <TextBlock x:Name="BannerPowerShell" Foreground="#D9F6FF" Margin="0,4,0,0"/>
            </StackPanel>
          </Border>
        </StackPanel>
      </Grid>
    </Border>

    <Border Grid.Row="1" Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource PanelBorderBrush}" BorderThickness="1" CornerRadius="0" Padding="12" Margin="0,8,0,0">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="Auto"/>
          <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal" DockPanel.Dock="Left">
          <Button x:Name="BtnAnalyzeZip" Content="Analyze Intune DiagLogs ZIP" Width="206" Style="{StaticResource PrimaryButton}"/>
          <Button x:Name="BtnAnalyzeLocal" Content="Analyze local device" Width="162" Style="{StaticResource SecondaryButton}"/>
          <Button x:Name="BtnExportHtml" Content="Export HTML Report" Width="158" Style="{StaticResource SuccessButton}" IsEnabled="False"/>
          <Button x:Name="BtnExportAnon" Content="Export anonymized ZIP" Width="172" IsEnabled="False"/>
          <Button x:Name="BtnReset" Content="Reset" Width="86" Margin="0"/>
          <Button x:Name="BtnCopySelected" Content="Copy selected" Width="112"/>
          <Button x:Name="BtnExportGridCsv" Content="Export grid CSV" Width="120"/>
          <Button x:Name="BtnOpenExtractDir" Content="Open extract dir" Width="120" IsEnabled="False"/>
          <Button x:Name="BtnOpenWuLog" Content="Open WU log" Width="105" IsEnabled="False"/>
          <Button x:Name="BtnOpenAppLog" Content="Open app log" Width="104"/>
        </StackPanel>
        <TextBlock x:Name="TxtFile" Grid.Column="1" Text="No file loaded" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="14,0,0,0" TextTrimming="CharacterEllipsis"/>
      </Grid>
    </Border>

    <TabControl x:Name="Tabs" Grid.Row="2" Margin="0">
      <TabItem x:Name="TabSummary" Header="Summary">
        <Grid Margin="8">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="145"/>
          </Grid.RowDefinitions>
          <TextBlock Text="Analysis Overview" Style="{StaticResource SectionTitle}"/>
          <UniformGrid Grid.Row="1" Columns="6" Margin="0,6,0,10">
            <Border Background="{StaticResource PanelSoftBrush}" Padding="14" Margin="0,0,10,0"><StackPanel><TextBlock Text="Computer Name" Foreground="{StaticResource MutedBrush}"/><TextBlock x:Name="SummaryComputer" Foreground="{StaticResource PurpleBrush}" FontWeight="SemiBold" FontSize="16"/></StackPanel></Border>
            <Border Background="{StaticResource PanelSoftBrush}" Padding="14" Margin="0,0,10,0"><StackPanel><TextBlock Text="IP Address" Foreground="{StaticResource MutedBrush}"/><TextBlock x:Name="SummaryIP" Foreground="{StaticResource PurpleBrush}" FontWeight="SemiBold" FontSize="16"/></StackPanel></Border>
            <Border Background="{StaticResource PanelSoftBrush}" Padding="14" Margin="0,0,10,0"><StackPanel><TextBlock Text="OS Version" Foreground="{StaticResource MutedBrush}"/><TextBlock x:Name="SummaryOS" Foreground="{StaticResource PurpleBrush}" FontWeight="SemiBold" FontSize="16"/></StackPanel></Border>
            <Border Background="{StaticResource PanelSoftBrush}" Padding="14" Margin="0,0,10,0"><StackPanel><TextBlock Text="Errors" Foreground="{StaticResource MutedBrush}"/><TextBlock x:Name="SummaryErrors" Foreground="#FF5A4F" FontWeight="Bold" FontSize="24"/></StackPanel></Border>
            <Border Background="{StaticResource PanelSoftBrush}" Padding="14" Margin="0,0,10,0"><StackPanel><TextBlock Text="Warnings" Foreground="{StaticResource MutedBrush}"/><TextBlock x:Name="SummaryWarnings" Foreground="#FFC107" FontWeight="Bold" FontSize="24"/></StackPanel></Border>
            <Border Background="{StaticResource PanelSoftBrush}" Padding="14"><StackPanel><TextBlock Text="Files Scanned" Foreground="{StaticResource MutedBrush}"/><TextBlock x:Name="SummaryFiles" Foreground="{StaticResource InkBrush}" FontWeight="Bold" FontSize="24"/></StackPanel></Border>
          </UniformGrid>
          <Grid Grid.Row="2">
            <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
            <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
            <TextBlock Text="Device Info" Foreground="{StaticResource AccentBrush}" FontWeight="SemiBold" Margin="4"/>
            <TextBlock Grid.Column="1" Text="Connection Info" Foreground="{StaticResource AccentBrush}" FontWeight="SemiBold" Margin="4"/>
            <DataGrid x:Name="SummaryDeviceGrid" Grid.Row="1" Margin="0,0,6,0"/>
            <DataGrid x:Name="SummaryConnectionGrid" Grid.Row="1" Grid.Column="1" Margin="6,0,0,0"/>
          </Grid>
          <TextBox x:Name="SummaryText" Grid.Row="3" Style="{StaticResource TextSurface}" TextWrapping="NoWrap" Margin="0,10,0,0"/>
        </Grid>
      </TabItem>
      <TabItem x:Name="TabMdm" Header="MDM Diagnostics">
        <Grid Margin="8">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
          <StackPanel Orientation="Horizontal" Margin="0,0,0,6">
            <TextBlock Text="MDM Diagnostics" Style="{StaticResource SectionTitle}" Margin="0,0,18,0"/>
            <Button x:Name="BtnExtractCab" Content="Extract CAB (expand.exe)" Width="166" Height="30" Style="{StaticResource PrimaryButton}" IsEnabled="False"/>
            <TextBlock x:Name="MdmCabStatus" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center"/>
          </StackPanel>
          <TabControl Grid.Row="1" Padding="0">
            <TabItem Header="Overview">
              <Grid Margin="8"><Grid.RowDefinitions><RowDefinition Height="180"/><RowDefinition Height="*"/></Grid.RowDefinitions><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                <DataGrid x:Name="MdmDeviceGrid" Margin="0,0,6,8"/>
                <DataGrid x:Name="MdmConnectionGrid" Grid.Column="1" Margin="6,0,0,8"/>
                <TextBox x:Name="MdmIssuesText" Grid.Row="1" Grid.ColumnSpan="2" Style="{StaticResource TextSurface}" TextWrapping="Wrap"/>
              </Grid>
            </TabItem>
            <TabItem Header="Policies">
              <Grid Margin="8">
                <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
                <StackPanel Orientation="Horizontal" Margin="0,0,0,8">
                  <TextBlock Text="Filter:" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="0,0,6,0"/>
                  <TextBox x:Name="MdmPolicyFilter" Width="260" Height="28" Margin="0,0,12,0"/>
                  <CheckBox x:Name="MdmShowInternalKnobs" Content="Show internal knobs" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center"/>
                  <TextBlock x:Name="MdmPolicyCount" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="18,0,0,0"/>
                </StackPanel>
                <DataGrid x:Name="MdmPoliciesGrid" Grid.Row="1"/>
              </Grid>
            </TabItem>
            <TabItem Header="Blocked GPs"><DataGrid x:Name="MdmBlockedGpsGrid" Margin="8"/></TabItem>
            <TabItem Header="LAPS"><DataGrid x:Name="MdmLapsGrid" Margin="8"/></TabItem>
            <TabItem Header="Sources"><DataGrid x:Name="MdmSourcesGrid" Margin="8"/></TabItem>
            <TabItem Header="CAB Files"><DataGrid x:Name="MdmCabFilesGrid" Margin="8"/></TabItem>
          </TabControl>
          <TextBox x:Name="MdmDiagText" Visibility="Collapsed"/>
        </Grid>
      </TabItem>
      <TabItem x:Name="TabWu" Header="Windows Update">
        <Grid Margin="8">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="120"/><RowDefinition Height="120"/><RowDefinition Height="*"/></Grid.RowDefinitions>
          <TextBlock Text="Windows Update" Style="{StaticResource SectionTitle}"/>
          <TextBox x:Name="WuInfoText" Grid.Row="1" Style="{StaticResource TextSurface}" TextWrapping="NoWrap" Margin="0,0,0,8"/>
          <DataGrid x:Name="WuPoliciesGrid" Grid.Row="2" Margin="0,0,0,8"/>
          <DataGrid x:Name="WuEventsGrid" Grid.Row="3"/>
          <TextBox x:Name="WuText" Visibility="Collapsed"/>
        </Grid>
      </TabItem>
      <TabItem x:Name="TabEvent" Header="Event Log">
        <Grid Margin="8">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="100"/><RowDefinition Height="*"/><RowDefinition Height="120"/></Grid.RowDefinitions>
          <TextBlock Text="Event Log" Style="{StaticResource SectionTitle}"/>
          <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
            <TextBlock Text="Log:" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="0,0,6,0"/>
            <ComboBox x:Name="EventLogSelector" Width="260" Height="28" Margin="0,0,10,0"/>
            <TextBlock Text="Show:" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="0,0,6,0"/>
            <ComboBox x:Name="EventLevelFilter" Width="120" Height="28" Margin="0,0,10,0">
              <ComboBoxItem Content="All"/>
              <ComboBoxItem Content="Errors"/>
              <ComboBoxItem Content="Warnings"/>
              <ComboBoxItem Content="Info"/>
            </ComboBox>
            <TextBlock Text="Search:" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="0,0,6,0"/>
            <TextBox x:Name="EventSearch" Width="280" Height="28" Margin="0,0,12,0"/>
            <TextBlock x:Name="EventCount" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center"/>
          </StackPanel>
          <DataGrid x:Name="EventLogSummaryGrid" Grid.Row="2" Margin="0,0,0,8"/>
          <DataGrid x:Name="EventLogGrid" Grid.Row="3" Margin="0,0,0,8"/>
          <TextBox x:Name="EventDetailText" Grid.Row="4" Style="{StaticResource TextSurface}" TextWrapping="Wrap"/>
          <TextBox x:Name="EventLogText" Visibility="Collapsed"/>
        </Grid>
      </TabItem>
      <TabItem x:Name="TabIme" Header="IME Logs">
        <Grid Margin="8">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="120"/></Grid.RowDefinitions>
          <TextBlock Text="IME Logs" Style="{StaticResource SectionTitle}"/>
          <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
            <TextBlock Text="Theme:" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="0,0,6,0"/>
            <ComboBox x:Name="ImeThemeFilter" Width="240" Height="28" Margin="0,0,10,0"/>
            <TextBlock Text="Show:" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="0,0,6,0"/>
            <ComboBox x:Name="ImeLevelFilter" Width="120" Height="28" Margin="0,0,10,0">
              <ComboBoxItem Content="All"/>
              <ComboBoxItem Content="Errors"/>
              <ComboBoxItem Content="Warnings"/>
              <ComboBoxItem Content="Info"/>
            </ComboBox>
            <TextBlock Text="Search:" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="0,0,6,0"/>
            <TextBox x:Name="ImeSearch" Width="300" Height="28" Margin="0,0,12,0"/>
            <TextBlock x:Name="ImeCount" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center"/>
          </StackPanel>
          <DataGrid x:Name="ImeGrid" Grid.Row="2" Margin="0,0,0,8"/>
          <TextBox x:Name="ImeText" Grid.Row="3" Style="{StaticResource TextSurface}" TextWrapping="Wrap"/>
        </Grid>
      </TabItem>
      <TabItem x:Name="TabApps" Header="Apps and Drivers">
        <Grid Margin="8">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
          <TextBlock Text="Apps &amp; Drivers" Style="{StaticResource SectionTitle}"/>
          <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
            <TextBlock Text="Search:" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="0,0,6,0"/>
            <TextBox x:Name="AppsSearch" Width="320" Height="28" Margin="0,0,12,0"/>
            <TextBlock x:Name="AppsCount" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center"/>
          </StackPanel>
          <TabControl Grid.Row="2" Padding="0">
            <TabItem Header="Apps"><DataGrid x:Name="AppsGrid" Margin="8"/></TabItem>
            <TabItem Header="Autopatch"><DataGrid x:Name="AutopatchGrid" Margin="8"/></TabItem>
            <TabItem Header="Collection Errors"><DataGrid x:Name="CollectionErrorsGrid" Margin="8"/></TabItem>
            <TabItem Header="Winget"><DataGrid x:Name="WingetGrid" Margin="8"/></TabItem>
            <TabItem Header="Drivers"><DataGrid x:Name="DriversGrid" Margin="8"/></TabItem>
            <TabItem Header="WiFi"><DataGrid x:Name="WifiGrid" Margin="8"/></TabItem>
          </TabControl>
          <TextBox x:Name="AppsText" Visibility="Collapsed"/>
        </Grid>
      </TabItem>
      <TabItem x:Name="TabHardware" Header="Hardware &amp; Security">
        <Grid Margin="8">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
          <TextBlock Text="Hardware &amp; Security" Style="{StaticResource SectionTitle}"/>
          <TabControl Grid.Row="1" Padding="0">
            <TabItem Header="Battery"><Grid Margin="8"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions><StackPanel Orientation="Horizontal" Margin="0,0,0,8"><TextBlock Text="Battery Health: " Foreground="{StaticResource MutedBrush}" FontWeight="SemiBold"/><TextBlock x:Name="BatteryHealthText" Foreground="#42D77D" FontWeight="Bold"/></StackPanel><DataGrid x:Name="BatteryGrid" Grid.Row="1"/></Grid></TabItem>
            <TabItem Header="Firewall"><DataGrid x:Name="FirewallGrid" Margin="8"/></TabItem>
            <TabItem Header="Certificates"><DataGrid x:Name="CertGrid" Margin="8"/></TabItem>
          </TabControl>
          <TextBox x:Name="HardwareText" Visibility="Collapsed"/>
        </Grid>
      </TabItem>
      <TabItem x:Name="TabInsights" Header="Insights">
        <Grid Margin="8">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="180"/></Grid.RowDefinitions>
          <TextBlock Text="Insights" Style="{StaticResource SectionTitle}"/>
          <DataGrid x:Name="InsightsGrid" Grid.Row="1" Margin="0,0,0,8"/>
          <TextBox x:Name="InsightsText" Grid.Row="2" Style="{StaticResource TextSurface}" TextWrapping="Wrap"/>
        </Grid>
      </TabItem>
      <TabItem x:Name="TabHealth" Header="Health">
        <Grid Margin="8">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
          <StackPanel Orientation="Horizontal">
            <TextBlock Text="Security &amp; Health" Style="{StaticResource SectionTitle}" Margin="0,0,18,0"/>
            <TextBlock x:Name="HealthBadge" Foreground="#FF5A4F" FontWeight="Bold" VerticalAlignment="Center"/>
          </StackPanel>
          <DataGrid x:Name="HealthGrid" Grid.Row="1"/>
          <TextBox x:Name="HealthText" Visibility="Collapsed"/>
        </Grid>
      </TabItem>
      <TabItem x:Name="TabDevice" Header="Device Info">
        <Grid Margin="8">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
          <TextBlock Text="Device Information" Style="{StaticResource SectionTitle}"/>
          <DataGrid x:Name="DeviceGrid" Grid.Row="1"/>
          <TextBox x:Name="DeviceText" Visibility="Collapsed"/>
          <TextBox x:Name="EnrollmentText" Visibility="Collapsed"/>
        </Grid>
      </TabItem>
      <TabItem x:Name="TabCompliance" Header="Compliance">
        <Grid Margin="8">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="130"/></Grid.RowDefinitions>
          <StackPanel Orientation="Horizontal">
            <TextBlock Text="Compliance" Style="{StaticResource SectionTitle}" Margin="0,0,18,0"/>
            <TextBlock x:Name="ComplianceBadge" Foreground="{StaticResource MutedBrush}" FontWeight="SemiBold" VerticalAlignment="Center"/>
          </StackPanel>
          <DataGrid x:Name="ComplianceGrid" Grid.Row="1" Margin="0,0,0,8"/>
          <TextBox x:Name="ComplianceText" Grid.Row="2" Style="{StaticResource TextSurface}" TextWrapping="Wrap"/>
        </Grid>
      </TabItem>
      <TabItem x:Name="TabEnrollments" Header="Enrollments">
        <Grid Margin="8">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="130"/></Grid.RowDefinitions>
          <TextBlock Text="MDM Enrollments" Style="{StaticResource SectionTitle}"/>
          <DataGrid x:Name="EnrollmentsGrid" Grid.Row="1" Margin="0,0,0,8"/>
          <TextBox x:Name="EnrollmentsText" Grid.Row="2" Style="{StaticResource TextSurface}" TextWrapping="Wrap"/>
        </Grid>
      </TabItem>
      <TabItem x:Name="TabZip" Header="ZIP Files">
        <Grid Margin="8">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="120"/></Grid.RowDefinitions>
          <TextBlock Text="ZIP Files" Style="{StaticResource SectionTitle}"/>
          <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
            <TextBlock Text="Category:" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="0,0,6,0"/>
            <ComboBox x:Name="ZipCategoryFilter" Width="220" Height="28" Margin="0,0,10,0"/>
            <TextBlock Text="Search:" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="0,0,6,0"/>
            <TextBox x:Name="ZipSearch" Width="320" Height="28" Margin="0,0,12,0"/>
            <TextBlock x:Name="ZipCount" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center"/>
          </StackPanel>
          <DataGrid x:Name="ZipFilesGrid" Grid.Row="2" Margin="0,0,0,8"/>
          <TextBox x:Name="ZipDetailText" Grid.Row="3" Style="{StaticResource TextSurface}" TextWrapping="Wrap"/>
        </Grid>
      </TabItem>
      <TabItem x:Name="TabSearch" Header="Search">
        <Grid Margin="8">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="120"/></Grid.RowDefinitions>
          <TextBlock Text="Search Corpus" Style="{StaticResource SectionTitle}"/>
          <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,0,0,8">
            <TextBlock Text="Search:" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="0,0,6,0"/>
            <TextBox x:Name="SearchFilter" Width="380" Height="28" Margin="0,0,12,0"/>
            <TextBlock x:Name="SearchCount" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center"/>
          </StackPanel>
          <DataGrid x:Name="SearchGrid" Grid.Row="2" Margin="0,0,0,8"/>
          <TextBox x:Name="SearchDetailText" Grid.Row="3" Style="{StaticResource TextSurface}" TextWrapping="Wrap"/>
          <TextBox x:Name="SearchText" Visibility="Collapsed"/>
        </Grid>
      </TabItem>
      <TabItem x:Name="TabAi" Header="AI Analysis">
        <Grid Margin="0,8,0,0">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
          </Grid.RowDefinitions>
          <Border Grid.Row="0" Background="{StaticResource PanelSoftBrush}" BorderBrush="{StaticResource PanelBorderBrush}" BorderThickness="1" CornerRadius="6" Padding="10" Margin="0,0,0,8">
            <Grid>
              <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="150"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="170"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="210"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="120"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="Auto"/>
              </Grid.ColumnDefinitions>
              <TextBlock Grid.Column="0" Text="Provider" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="0,0,6,0"/>
              <ComboBox x:Name="AiProvider" Grid.Column="1" Height="30" Margin="0,0,10,0">
                <ComboBoxItem Content="claude"/>
                <ComboBoxItem Content="openai"/>
                <ComboBoxItem Content="ollama"/>
              </ComboBox>
              <TextBlock Grid.Column="2" Text="Model" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="0,0,6,0"/>
              <TextBox x:Name="AiModel" Grid.Column="3" Height="30" Margin="0,0,10,0"/>
              <TextBlock Grid.Column="4" Text="API Key / Ollama URL" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="0,0,6,0"/>
              <TextBox x:Name="AiCredential" Grid.Column="5" Height="30" Margin="0,0,10,0"/>
              <TextBlock Grid.Column="6" Text="Tokens" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center" Margin="0,0,6,0"/>
              <Grid Grid.Column="7" Height="30" Margin="0,0,10,0">
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="28"/>
                  <ColumnDefinition Width="*"/>
                  <ColumnDefinition Width="28"/>
                </Grid.ColumnDefinitions>
                <Button x:Name="AiTokensDownButton" Grid.Column="0" Content="-" Width="28" Height="30" Padding="0" Margin="0"/>
                <TextBox x:Name="AiTokens" Grid.Column="1" Height="30" Margin="0" TextAlignment="Center"/>
                <Button x:Name="AiTokensUpButton" Grid.Column="2" Content="+" Width="28" Height="30" Padding="0" Margin="0"/>
              </Grid>
              <CheckBox x:Name="AiRemember" Grid.Column="8" Content="Remember API key" VerticalAlignment="Center" Foreground="{StaticResource MutedBrush}"/>
              <Button x:Name="BtnRunAi" Grid.Column="9" Content="Analyze" Width="96" Style="{StaticResource PrimaryButton}" Margin="8,0,0,0"/>
              <Button x:Name="BtnCopyAi" Grid.Column="10" Content="Copy" Width="76" Margin="8,0,0,0"/>
              <Button x:Name="BtnSaveAi" Grid.Column="11" Content="Save" Width="76" Margin="8,0,0,0"/>
            </Grid>
          </Border>
          <TextBox x:Name="AiText" Grid.Row="1" Style="{StaticResource TextSurface}" TextWrapping="Wrap"/>
        </Grid>
      </TabItem>
      <TabItem x:Name="TabWin11" Header="Win11 Readiness">
        <Grid Margin="8">
          <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="160"/></Grid.RowDefinitions>
          <StackPanel Orientation="Horizontal">
            <TextBlock Text="Windows 11 Readiness" Style="{StaticResource SectionTitle}" Margin="0,0,18,0"/>
            <TextBlock x:Name="Win11StatusText" Foreground="{StaticResource MutedBrush}" FontWeight="Bold" VerticalAlignment="Center"/>
          </StackPanel>
          <DataGrid x:Name="Win11Grid" Grid.Row="1" Margin="0,0,0,8"/>
          <TextBox x:Name="Win11Text" Grid.Row="2" Style="{StaticResource TextSurface}" TextWrapping="Wrap"/>
        </Grid>
      </TabItem>
    </TabControl>

    <Border Grid.Row="3" Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource PanelBorderBrush}" BorderThickness="1" CornerRadius="0" Padding="8" Margin="0,0,0,0">
      <DockPanel>
        <TextBlock x:Name="StatusText" Text="Ready - Open an Intune Device Diagnostics ZIP file" Foreground="{StaticResource MutedBrush}" VerticalAlignment="Center"/>
        <ProgressBar x:Name="Progress" Width="150" Height="10" DockPanel.Dock="Right" IsIndeterminate="True" Visibility="Collapsed"/>
      </DockPanel>
    </Border>

    <Grid x:Name="ScanOverlay" Grid.RowSpan="4" Panel.ZIndex="50" Visibility="Collapsed" Background="#99091F33">
      <Border Width="460" MinHeight="170" Background="{StaticResource PanelBrush}" BorderBrush="{StaticResource PanelBorderBrush}" BorderThickness="1" CornerRadius="0" Padding="22" HorizontalAlignment="Center" VerticalAlignment="Center">
        <StackPanel>
          <TextBlock x:Name="OverlayTitle" Text="Working..." FontSize="22" FontWeight="SemiBold" Foreground="{StaticResource InkBrush}"/>
          <TextBlock x:Name="OverlayStep" Text="Please wait." FontSize="14" Foreground="{StaticResource MutedBrush}" TextWrapping="Wrap" Margin="0,10,0,18"/>
          <ProgressBar Style="{StaticResource OverlayProgress}"/>
          <TextBlock Text="The window may pause briefly while Windows diagnostics are collected." Foreground="{StaticResource MutedBrush}" FontSize="12" Margin="0,14,0,0" TextWrapping="Wrap"/>
        </StackPanel>
      </Border>
    </Grid>
  </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    if (Test-Path -LiteralPath $script:LogoPath) {
        try { $window.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create((New-Object Uri($script:LogoPath))) } catch {}
    }

    $controls = @{}
    foreach ($name in @(
        'HeaderLogoLink','HeaderLogo','BannerComputer','BannerUser','BannerAccount','BannerOS','BannerPowerShell','ScanOverlay','OverlayTitle','OverlayStep',
        'BtnAnalyzeZip','BtnAnalyzeLocal','BtnExportHtml','BtnExportAnon','BtnReset','BtnCopySelected','BtnExportGridCsv','BtnOpenExtractDir','BtnOpenWuLog','BtnOpenAppLog','TxtFile','StatusText','Progress',
        'TabSummary','TabMdm','TabWu','TabEvent','TabIme','TabApps','TabHardware','TabInsights','TabHealth','TabDevice','TabCompliance','TabEnrollments','TabZip','TabSearch','TabAi','TabWin11',
        'SummaryText','SummaryComputer','SummaryIP','SummaryOS','SummaryErrors','SummaryWarnings','SummaryFiles','SummaryDeviceGrid','SummaryConnectionGrid',
        'InsightsText','InsightsGrid','HealthText','HealthGrid','HealthBadge','WuText','WuInfoText','WuPoliciesGrid','WuEventsGrid',
        'DeviceText','DeviceGrid','EnrollmentText','ComplianceText','ComplianceGrid','ComplianceBadge','EnrollmentsText','EnrollmentsGrid','MdmDiagText','BtnExtractCab','MdmCabStatus','MdmDeviceGrid','MdmConnectionGrid','MdmIssuesText','MdmPolicyFilter','MdmShowInternalKnobs','MdmPolicyCount','MdmPoliciesGrid','MdmBlockedGpsGrid','MdmLapsGrid','MdmSourcesGrid','MdmCabFilesGrid',
        'EventLogText','EventLogSummaryGrid','EventLogGrid','EventLogSelector','EventLevelFilter','EventSearch','EventCount','EventDetailText',
        'ImeText','ImeGrid','ImeThemeFilter','ImeLevelFilter','ImeSearch','ImeCount',
        'AppsText','AppsSearch','AppsCount','AppsGrid','AutopatchGrid','CollectionErrorsGrid','WingetGrid','DriversGrid','WifiGrid',
        'HardwareText','BatteryHealthText','BatteryGrid','FirewallGrid','CertGrid','SearchText','SearchFilter','SearchCount','SearchGrid','SearchDetailText','ZipCategoryFilter','ZipSearch','ZipCount','ZipFilesGrid','ZipDetailText',
        'AiProvider','AiModel','AiCredential','AiTokensDownButton','AiTokens','AiTokensUpButton','AiRemember','BtnRunAi','BtnCopyAi','BtnSaveAi','AiText','Win11Text','Win11StatusText','Win11Grid'
    )) {
        $controls[$name] = $window.FindName($name)
    }
    $bannerInfo = Get-SedaLocalBannerInfo
    $controls.BannerComputer.Text = "PC: $($bannerInfo.ComputerName)"
    $controls.BannerUser.Text = "Account: $($bannerInfo.Account)"
    $controls.BannerAccount.Text = "$($bannerInfo.Context) | $($bannerInfo.Upn)"
    $controls.BannerOS.Text = "OS: $($bannerInfo.OS)"
    $controls.BannerPowerShell.Text = $bannerInfo.PowerShell
    Write-SedaLog -Level INFO -Message "Banner info loaded: PC=$($bannerInfo.ComputerName); Account=$($bannerInfo.Account); UPN=$($bannerInfo.Upn); SID=$($bannerInfo.Sid); Context=$($bannerInfo.Context); OS=$($bannerInfo.OS); $($bannerInfo.PowerShell)."
    if ($controls.HeaderLogo -and (Test-Path -LiteralPath $script:WorkplaceLogoPath -PathType Leaf)) {
        try {
            $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
            $bitmap.BeginInit()
            $bitmap.UriSource = [Uri]$script:WorkplaceLogoPath
            $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
            $bitmap.EndInit()
            $bitmap.Freeze()
            $controls.HeaderLogo.Source = $bitmap
        } catch {
            Write-SedaLog -Level WARN -Message 'Unable to load WorkplaceCloudHub-lockup-WPF.png in GUI header.' -Exception $_.Exception
        }
    }

    function Open-ExternalUrl {
        param([Parameter(Mandatory)][string]$Url)

        try {
            $psi = [System.Diagnostics.ProcessStartInfo]::new($Url)
            $psi.UseShellExecute = $true
            [System.Diagnostics.Process]::Start($psi) | Out-Null
        }
        catch {
            [System.Windows.MessageBox]::Show($window, "Unable to open:`r`n$Url`r`n`r`n$($_.Exception.Message)", 'SmartM365', 'OK', 'Warning') | Out-Null
        }
    }

    if ($controls.HeaderLogoLink) {
        $controls.HeaderLogoLink.Add_MouseLeftButtonUp({ Open-ExternalUrl -Url 'https://workplacecloudhub.com' })
    }
    $script:CurrentAnalysis = $null
    $script:CurrentAiConfig = Get-SedaAIConfig
    $script:EventRowsFull = @()
    $script:ImeRowsFull = @()
    $script:AppsRowsFull = @()
    $script:MdmPoliciesFull = @()
    $script:SearchRowsFull = @()
    $script:ZipRowsFull = @()
    $script:CurrentGrid = $null

    function set-status([string]$text, [bool]$busy = $false) {
        $controls.StatusText.Text = $text
        $controls.Progress.Visibility = if ($busy) { 'Visible' } else { 'Collapsed' }
        if ($busy -and $controls.ScanOverlay.Visibility -eq 'Visible') { $controls.OverlayStep.Text = $text }
        Write-SedaLog -Level INFO -Message "GUI status: $text"
        [System.Windows.Forms.Application]::DoEvents()
    }

    function show-overlay([string]$title, [string]$step) {
        $controls.OverlayTitle.Text = $title
        $controls.OverlayStep.Text = $step
        $controls.ScanOverlay.Visibility = 'Visible'
        $controls.BtnAnalyzeZip.IsEnabled = $false
        $controls.BtnAnalyzeLocal.IsEnabled = $false
        $controls.BtnReset.IsEnabled = $false
        [System.Windows.Forms.Application]::DoEvents()
    }

    function hide-overlay {
        $controls.ScanOverlay.Visibility = 'Collapsed'
        $controls.BtnAnalyzeZip.IsEnabled = $true
        $controls.BtnAnalyzeLocal.IsEnabled = $true
        $controls.BtnReset.IsEnabled = $true
        [System.Windows.Forms.Application]::DoEvents()
    }

    function get-ai-provider {
        if ($controls.AiProvider.SelectedItem -and $controls.AiProvider.SelectedItem.Content) {
            return [string]$controls.AiProvider.SelectedItem.Content
        }
        return 'claude'
    }

    function set-ai-token-value([int]$Delta) {
        $tokens = 2048
        if (-not [int]::TryParse($controls.AiTokens.Text, [ref]$tokens)) {
            $tokens = 2048
        }

        $tokens += $Delta
        if ($tokens -lt 1) {
            $tokens = 1
        }

        $controls.AiTokens.Text = [string]$tokens
        $controls.AiTokens.CaretIndex = $controls.AiTokens.Text.Length
    }

    function set-ai-defaults([string]$provider, [bool]$forceModel = $false) {
        if (-not $provider) { $provider = 'claude' }
        $defaults = @{
            claude = 'claude-3-5-sonnet-20241022'
            openai = 'gpt-4o-mini'
            ollama = 'llama3.1'
        }
        if ($forceModel -or [string]::IsNullOrWhiteSpace($controls.AiModel.Text)) {
            $controls.AiModel.Text = $defaults[$provider]
        }
        if ($provider -eq 'ollama') {
            if ([string]::IsNullOrWhiteSpace($controls.AiCredential.Text) -or $controls.AiCredential.Text -match '^(sk-|AIza|anthropic|claude)') {
                $controls.AiCredential.Text = 'http://localhost:11434'
            }
            $controls.AiRemember.IsChecked = $false
            $controls.AiRemember.IsEnabled = $false
        } else {
            $controls.AiRemember.IsEnabled = $true
        }
    }

    function get-ai-config-from-ui {
        $provider = get-ai-provider
        $tokens = 2048
        [void][int]::TryParse($controls.AiTokens.Text, [ref]$tokens)
        $credential = [string]$controls.AiCredential.Text
        return New-SedaObject ([ordered]@{
            Provider = $provider
            Model = [string]$controls.AiModel.Text
            ApiKey = if ($provider -eq 'ollama') { '' } else { $credential }
            OllamaUrl = if ($provider -eq 'ollama' -and -not [string]::IsNullOrWhiteSpace($credential)) { $credential } else { 'http://localhost:11434' }
            MaxTokens = $tokens
            Temperature = 0.3
            RememberApiKey = [bool]$controls.AiRemember.IsChecked
        })
    }

    function set-grid([string]$Name, [object[]]$Rows) {
        if ($controls[$Name]) {
            $controls[$Name].ItemsSource = $null
            $controls[$Name].ItemsSource = @(ConvertTo-SedaGridRows -Rows $Rows)
        }
    }

    function get-selected-text([object]$Item) {
        if ($null -eq $Item) { return '' }
        $lines = New-Object System.Collections.Generic.List[string]
        foreach ($property in $Item.PSObject.Properties) {
            $lines.Add(('{0}: {1}' -f $property.Name, $property.Value))
        }
        return ($lines -join [Environment]::NewLine)
    }

    function row-tone([object]$Item) {
        if ($null -eq $Item) { return 'default' }
        $values = @(
            [string]$Item.Severity,
            [string]$Item.Level,
            [string]$Item.Status,
            [string]$Item.ReturnResult,
            [string]$Item.Title
        ) -join ' '
        if ($Item.PSObject.Properties.Name -contains 'IsBlocking' -and [bool]$Item.IsBlocking) { return 'error' }
        if ($values -match 'ERROR|CRITICAL|FAIL|FAILED|NON_COMPLIANT|EXPIRED|NOT_CAPABLE') { return 'error' }
        if ($values -match 'WARN|WARNING|UNDETERMINED|PENDING|EXPIRING|ATTENTION') { return 'warn' }
        if ($values -match 'OK|PASS|COMPLIANT|HEALTHY|CAPABLE|ON|TRUE') { return 'ok' }
        if ($values -match 'INFO|INFORMATION') { return 'info' }
        return 'default'
    }

    $brushes = @{
        error = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(255,241,240))
        warn = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(255,247,237))
        ok = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(236,253,245))
        info = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(239,246,255))
        default = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(255,255,255))
        alt = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.Color]::FromRgb(246,250,253))
    }

    function register-grid-formatting([string[]]$Names) {
        foreach ($name in $Names) {
            $grid = $controls[$name]
            if (-not $grid) { continue }
            $menu = New-Object System.Windows.Controls.ContextMenu
            $copyRow = New-Object System.Windows.Controls.MenuItem
            $copyRow.Header = 'Copy selected row'
            $copyRow.Add_Click({
                if ($grid.SelectedItem) {
                    [System.Windows.Clipboard]::SetText((get-selected-text $grid.SelectedItem))
                    set-status 'Selected row copied.' $false
                }
            }.GetNewClosure())
            [void]$menu.Items.Add($copyRow)
            $copyGrid = New-Object System.Windows.Controls.MenuItem
            $copyGrid.Header = 'Copy visible grid as CSV'
            $copyGrid.Add_Click({
                if ($grid.ItemsSource) {
                    $csv = @($grid.ItemsSource) | ConvertTo-Csv -NoTypeInformation
                    [System.Windows.Clipboard]::SetText(($csv -join [Environment]::NewLine))
                    set-status 'Visible grid copied as CSV.' $false
                }
            }.GetNewClosure())
            [void]$menu.Items.Add($copyGrid)
            [void]$menu.Items.Add((New-Object System.Windows.Controls.Separator))
            $searchWeb = New-Object System.Windows.Controls.MenuItem
            $searchWeb.Header = 'Search selected row on web'
            $searchWeb.Add_Click({
                if ($grid.SelectedItem) {
                    $query = [Uri]::EscapeDataString((get-selected-text $grid.SelectedItem))
                    Start-Process -FilePath "https://www.bing.com/search?q=$query" | Out-Null
                }
            }.GetNewClosure())
            [void]$menu.Items.Add($searchWeb)
            $askAi = New-Object System.Windows.Controls.MenuItem
            $askAi.Header = 'Use selected row as AI focus'
            $askAi.Add_Click({
                if ($grid.SelectedItem) {
                    $script:AiFocusContext = get-selected-text $grid.SelectedItem
                    $controls.AiText.Text = 'AI focus set from selected row:' + [Environment]::NewLine + [Environment]::NewLine + $script:AiFocusContext + [Environment]::NewLine + [Environment]::NewLine + 'Click Analyze to include this focus in the AI prompt.'
                    $controls.TabAi.IsSelected = $true
                    set-status 'AI focus set from selected row.' $false
                }
            }.GetNewClosure())
            [void]$menu.Items.Add($askAi)
            $grid.ContextMenu = $menu
            $grid.Add_PreviewMouseRightButtonDown({
                param($sender, $eventArgs)
                $script:CurrentGrid = $sender
                $source = $eventArgs.OriginalSource
                while ($source -and -not ($source -is [System.Windows.Controls.DataGridRow])) {
                    try { $source = [System.Windows.Media.VisualTreeHelper]::GetParent($source) } catch { $source = $null }
                }
                if ($source -and $source.Item) {
                    $sender.SelectedItem = $source.Item
                    $source.IsSelected = $true
                }
            })
            $grid.Add_LoadingRow({
                param($sender, $eventArgs)
                $tone = row-tone $eventArgs.Row.Item
                if ($tone -eq 'default' -and ($eventArgs.Row.GetIndex() % 2) -eq 1) { $tone = 'alt' }
                $eventArgs.Row.Background = $brushes[$tone]
                $eventArgs.Row.Foreground = [System.Windows.Media.Brushes]::Black
            })
            $grid.Add_GotFocus({
                param($sender, $eventArgs)
                $script:CurrentGrid = $sender
            })
            $grid.Add_SelectionChanged({
                param($sender, $eventArgs)
                if ($sender.SelectedItem) { $script:CurrentGrid = $sender }
            })
            $grid.Add_AutoGeneratingColumn({
                param($sender, $eventArgs)
                if ($eventArgs.PropertyName -eq 'RawLine') {
                    $eventArgs.Cancel = $true
                    return
                }
                if ($eventArgs.Column) {
                    $eventArgs.Column.MinWidth = 70
                    $eventArgs.Column.MaxWidth = 520
                    if ($eventArgs.PropertyName -in @('Message','Detail','Recommendation','Path')) {
                        $eventArgs.Column.Width = New-Object System.Windows.Controls.DataGridLength -ArgumentList 1, ([System.Windows.Controls.DataGridLengthUnitType]::Star)
                    }
                }
            })
        }
    }

    function test-row-match([object]$Item, [string]$Text) {
        if ([string]::IsNullOrWhiteSpace($Text)) { return $true }
        $needle = $Text.ToLowerInvariant()
        foreach ($property in $Item.PSObject.Properties) {
            if(([string]$property.Value).ToLowerInvariant().Contains($needle)) { return $true }
        }
        return $false
    }

    function get-combo-content([object]$ComboBox, [string]$Default = 'All') {
        if (-not $ComboBox -or -not $ComboBox.SelectedItem) { return $Default }
        if ($ComboBox.SelectedItem.PSObject.Properties.Name -contains 'Content') { return [string]$ComboBox.SelectedItem.Content }
        return [string]$ComboBox.SelectedItem
    }

    function update-event-grid {
        $rows = @($script:EventRowsFull)
        $log = get-combo-content $controls.EventLogSelector
        $level = get-combo-content $controls.EventLevelFilter
        $search = [string]$controls.EventSearch.Text
        if ($log -and $log -ne 'All') { $rows = @($rows | Where-Object { $_.Log -eq $log }) }
        if ($level -eq 'Errors') { $rows = @($rows | Where-Object { $_.Level -in @('Critical','Error') }) }
        elseif ($level -eq 'Warnings') { $rows = @($rows | Where-Object { $_.Level -eq 'Warning' }) }
        elseif ($level -eq 'Info') { $rows = @($rows | Where-Object { $_.Level -in @('Information','Info') }) }
        if (-not [string]::IsNullOrWhiteSpace($search)) { $rows = @($rows | Where-Object { test-row-match $_ $search }) }
        set-grid 'EventLogGrid' (@($rows | Select-Object -First 2000))
        $controls.EventCount.Text = "$(@($rows).Count)/$(@($script:EventRowsFull).Count) events"
        $controls.TabEvent.Header = "Event Log ($(@($rows).Count))"
    }

    function update-ime-grid {
        $rows = @($script:ImeRowsFull)
        $theme = get-combo-content $controls.ImeThemeFilter
        $level = get-combo-content $controls.ImeLevelFilter
        $search = [string]$controls.ImeSearch.Text
        if ($theme -and $theme -ne 'All') { $rows = @($rows | Where-Object { $_.ThemeLabel -eq $theme -or $_.Theme -eq $theme }) }
        if ($level -eq 'Errors') { $rows = @($rows | Where-Object { $_.Severity -eq 'ERROR' }) }
        elseif ($level -eq 'Warnings') { $rows = @($rows | Where-Object { $_.Severity -eq 'WARNING' }) }
        elseif ($level -eq 'Info') { $rows = @($rows | Where-Object { $_.Severity -in @('INFO','INFORMATION') }) }
        if (-not [string]::IsNullOrWhiteSpace($search)) { $rows = @($rows | Where-Object { test-row-match $_ $search }) }
        set-grid 'ImeGrid' (@($rows | Select-Object -First 2000))
        $controls.ImeCount.Text = "$(@($rows).Count)/$(@($script:ImeRowsFull).Count) events"
        $controls.TabIme.Header = "IME Logs ($(@($rows).Count))"
    }

    function update-apps-grid {
        $rows = @($script:AppsRowsFull)
        $search = [string]$controls.AppsSearch.Text
        if (-not [string]::IsNullOrWhiteSpace($search)) { $rows = @($rows | Where-Object { test-row-match $_ $search }) }
        set-grid 'AppsGrid' $rows
        $controls.AppsCount.Text = "$(@($rows).Count)/$(@($script:AppsRowsFull).Count) apps"
        $controls.TabApps.Header = "Apps and Drivers ($(@($script:AppsRowsFull).Count))"
    }

    function update-mdm-policy-grid {
        $rows = @($script:MdmPoliciesFull)
        $search = [string]$controls.MdmPolicyFilter.Text
        if (-not [bool]$controls.MdmShowInternalKnobs.IsChecked) {
            $rows = @($rows | Where-Object { $_.Area -notmatch 'Knob|Internal|Provider|Metadata|ADMX' -and $_.Policy -notmatch 'Knob|Internal|Provider|Metadata|ADMX' })
        }
        if (-not [string]::IsNullOrWhiteSpace($search)) { $rows = @($rows | Where-Object { test-row-match $_ $search }) }
        set-grid 'MdmPoliciesGrid' $rows
        $controls.MdmPolicyCount.Text = "$(@($rows).Count)/$(@($script:MdmPoliciesFull).Count) policies"
    }

    function update-search-grid {
        $rows = @($script:SearchRowsFull)
        $search = [string]$controls.SearchFilter.Text
        if (-not [string]::IsNullOrWhiteSpace($search)) { $rows = @($rows | Where-Object { test-row-match $_ $search }) }
        set-grid 'SearchGrid' (@($rows | Select-Object -First 2000))
        $controls.SearchCount.Text = "$(@($rows).Count)/$(@($script:SearchRowsFull).Count) rows"
        $controls.TabSearch.Header = "Search ($(@($rows).Count))"
    }

    function update-zip-grid {
        $rows = @($script:ZipRowsFull)
        $category = get-combo-content $controls.ZipCategoryFilter
        $search = [string]$controls.ZipSearch.Text
        if ($category -and $category -ne 'All') { $rows = @($rows | Where-Object { $_.Category -eq $category }) }
        if (-not [string]::IsNullOrWhiteSpace($search)) { $rows = @($rows | Where-Object { test-row-match $_ $search }) }
        set-grid 'ZipFilesGrid' (@($rows | Select-Object -First 4000))
        $controls.ZipCount.Text = "$(@($rows).Count)/$(@($script:ZipRowsFull).Count) files"
        $controls.TabZip.Header = "ZIP Files ($(@($rows).Count))"
    }

    function get-selected-grid {
        if ($script:CurrentGrid -and $script:CurrentGrid.ItemsSource) { return $script:CurrentGrid }
        foreach ($name in @(
            'EventLogGrid','ImeGrid','AppsGrid','AutopatchGrid','CollectionErrorsGrid','WingetGrid','DriversGrid','WifiGrid','MdmPoliciesGrid','MdmBlockedGpsGrid','MdmSourcesGrid','MdmCabFilesGrid',
            'WuEventsGrid','WuPoliciesGrid','HealthGrid','InsightsGrid','DeviceGrid','ComplianceGrid','EnrollmentsGrid','BatteryGrid','FirewallGrid','CertGrid','SearchGrid','ZipFilesGrid','Win11Grid'
        )) {
            $grid = $controls[$name]
            if ($grid -and $grid.SelectedItem) { return $grid }
        }
        return $null
    }

    function export-grid-csv([object]$Grid, [string]$Path) {
        if (-not $Grid -or -not $Grid.ItemsSource) { throw 'No grid data selected.' }
        $rows = @($Grid.ItemsSource)
        if ($rows.Count -eq 0) { throw 'Selected grid is empty.' }
        $rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
    }

    function open-path([string]$Path) {
        if ([string]::IsNullOrWhiteSpace($Path)) { return }
        if (Test-Path -LiteralPath $Path) {
            Start-Process -FilePath $Path | Out-Null
        }
    }

    function render-analysis([object]$analysis) {
        $controls.SummaryComputer.Text = [string]$analysis.DeviceSummary.ComputerName
        $controls.SummaryIP.Text = [string]$analysis.DeviceSummary.IPAddress
        $controls.SummaryOS.Text = [string]$analysis.DeviceSummary.OSVersion
        $controls.SummaryErrors.Text = [string](@($analysis.ImeEvents | Where-Object { $_.Severity -eq 'ERROR' }).Count + @($analysis.EventLogs.Events | Where-Object { $_.Level -in @('Critical','Error') }).Count + [int]$analysis.WindowsUpdate.EtlErrorCount)
        $controls.SummaryWarnings.Text = [string](@($analysis.ImeEvents | Where-Object { $_.Severity -eq 'WARNING' }).Count + @($analysis.EventLogs.Events | Where-Object { $_.Level -eq 'Warning' }).Count + [int]$analysis.WindowsUpdate.EtlWarningCount)
        $controls.SummaryFiles.Text = [string]$analysis.ZipInfo.TotalFiles
        $controls.SummaryText.Text = Format-SedaSummaryText -Analysis $analysis
        set-grid 'SummaryDeviceGrid' (ConvertTo-SedaKeyValueRows $analysis.DeviceInfo)
        set-grid 'SummaryConnectionGrid' (ConvertTo-SedaKeyValueRows $analysis.MdmDiagnostics.ConnectionInfo)

        set-grid 'InsightsGrid' $analysis.Insights.TopActions
        $controls.TabInsights.Header = "Insights ($(@($analysis.Insights.TopActions).Count))"
        $controls.InsightsText.Text = 'Timeline:' + [Environment]::NewLine + (ConvertTo-SedaTextTable -Rows $analysis.Insights.Timeline -Properties @('Timestamp','Severity','Source','Title','Detail'))
        set-grid 'HealthGrid' $analysis.Health.Findings
        $controls.TabHealth.Header = "Health ($($analysis.Health.ErrorCount)/$($analysis.Health.WarningCount))"
        $controls.HealthBadge.Text = "$($analysis.Health.ErrorCount) error(s) | $($analysis.Health.WarningCount) warning(s)"
        $controls.HealthText.Text = 'Health summary:' + [Environment]::NewLine + (ConvertTo-SedaKeyValueText $analysis.Health.Summary)
        set-grid 'ComplianceGrid' $analysis.Compliance.PolicyStatuses
        $controls.TabCompliance.Header = "Compliance ($($analysis.Compliance.OverallStatus))"
        $controls.ComplianceBadge.Text = "OK=$($analysis.Compliance.CompliantCount) | KO=$($analysis.Compliance.NonCompliantCount) | Pending=$($analysis.Compliance.PendingCount) | Unknown=$($analysis.Compliance.UnknownCount)"
        $controls.ComplianceText.Text = 'Compliance summary:' + [Environment]::NewLine + (ConvertTo-SedaKeyValueText $analysis.Compliance)
        set-grid 'EnrollmentsGrid' $analysis.Enrollments
        $controls.TabEnrollments.Header = "Enrollments ($(@($analysis.Enrollments).Count))"
        $controls.EnrollmentsText.Text = 'Enrollment summary:' + [Environment]::NewLine + (ConvertTo-SedaKeyValueText $analysis.EnrollmentInfo)

        $controls.WuInfoText.Text = "Windows Update Orchestrator:`r`n" + (ConvertTo-SedaKeyValueText $analysis.WindowsUpdate.Info) + [Environment]::NewLine + [Environment]::NewLine + 'Generated log: ' + [string]$analysis.WindowsUpdate.GeneratedLogPath + [Environment]::NewLine + [string]$analysis.WindowsUpdate.EtlStatus
        set-grid 'WuPoliciesGrid' $analysis.WindowsUpdate.Policies
        set-grid 'WuEventsGrid' (@($analysis.WindowsUpdate.EtlEvents) + @($analysis.WindowsUpdate.ReportingEvents) | Select-Object -First 2000)
        $controls.TabWu.Header = "Windows Update ($($analysis.WindowsUpdate.EtlErrorCount)/$($analysis.WindowsUpdate.EtlWarningCount))"
        $controls.WuText.Text = $controls.WuInfoText.Text

        $deviceRows = @()
        $deviceRows += ConvertTo-SedaKeyValueRows $analysis.DeviceSummary
        $deviceRows += ConvertTo-SedaKeyValueRows $analysis.DeviceInfo
        $deviceRows += ConvertTo-SedaKeyValueRows $analysis.SsoInfo
        $deviceRows += ConvertTo-SedaKeyValueRows $analysis.EnrollmentInfo
        $deviceRows += ConvertTo-SedaKeyValueRows $analysis.ZipInfo
        set-grid 'DeviceGrid' $deviceRows
        $controls.TabDevice.Header = "Device Info"
        $controls.DeviceText.Text = 'Device summary:' + [Environment]::NewLine + (ConvertTo-SedaKeyValueText $analysis.DeviceSummary)
        $controls.EnrollmentText.Text = 'Enrollment summary:' + [Environment]::NewLine + (ConvertTo-SedaKeyValueText $analysis.EnrollmentInfo)

        $controls.MdmCabStatus.Text = [string]$analysis.MdmDiagnostics.CabStatus
        $controls.BtnExtractCab.IsEnabled = @($analysis.Inventory.cab).Count -gt 0
        set-grid 'MdmDeviceGrid' (ConvertTo-SedaKeyValueRows $analysis.MdmDiagnostics.DeviceInfo)
        set-grid 'MdmConnectionGrid' (ConvertTo-SedaKeyValueRows $analysis.MdmDiagnostics.ConnectionInfo)
        $script:MdmPoliciesFull = @($analysis.MdmDiagnostics.ManagedPolicies)
        update-mdm-policy-grid
        set-grid 'MdmBlockedGpsGrid' $analysis.MdmDiagnostics.BlockedGps
        set-grid 'MdmLapsGrid' (ConvertTo-SedaKeyValueRows $analysis.MdmDiagnostics.Laps)
        set-grid 'MdmSourcesGrid' $analysis.MdmDiagnostics.ConfigSources
        set-grid 'MdmCabFilesGrid' (@($analysis.Inventory.cab) | ForEach-Object { New-SedaObject @{ File=[System.IO.Path]::GetFileName($_); Path=$_ } })
        $controls.MdmIssuesText.Text = (ConvertTo-SedaTextTable -Rows $analysis.MdmDiagnostics.Issues -Properties @('Severity','Area','Title','Detail','Recommendation'))
        $controls.MdmDiagText.Text = $controls.MdmIssuesText.Text
        $controls.TabMdm.Header = "MDM Diagnostics ($(@($analysis.MdmDiagnostics.Issues).Count))"

        set-grid 'EventLogSummaryGrid' $analysis.EventLogs.Summary
        $script:EventRowsFull = @($analysis.EventLogs.Events)
        $controls.EventLogSelector.Items.Clear()
        [void]$controls.EventLogSelector.Items.Add('All')
        foreach ($logName in @($script:EventRowsFull | Select-Object -ExpandProperty Log -Unique | Sort-Object)) { [void]$controls.EventLogSelector.Items.Add([string]$logName) }
        $controls.EventLogSelector.SelectedIndex = 0
        if ($controls.EventLevelFilter.SelectedIndex -lt 0) { $controls.EventLevelFilter.SelectedIndex = 0 }
        update-event-grid
        $controls.EventLogText.Text = 'Event log summary:' + [Environment]::NewLine + (ConvertTo-SedaTextTable -Rows $analysis.EventLogs.Summary -Properties @('Log','Critical','Error','Warning','Information','Other','Scanned'))

        $script:ImeRowsFull = @($analysis.ImeEvents | ForEach-Object {
            $label = if ($script:ImeThemeLabels.ContainsKey($_.Theme)) { $script:ImeThemeLabels[$_.Theme] } else { $_.Theme }
            $_ | Add-Member -NotePropertyName ThemeLabel -NotePropertyValue $label -Force
            $_
        })
        $controls.ImeThemeFilter.Items.Clear()
        [void]$controls.ImeThemeFilter.Items.Add('All')
        foreach ($themeLabel in @($script:ImeRowsFull | Select-Object -ExpandProperty ThemeLabel -Unique | Sort-Object)) { if ($themeLabel) { [void]$controls.ImeThemeFilter.Items.Add([string]$themeLabel) } }
        $controls.ImeThemeFilter.SelectedIndex = 0
        if ($controls.ImeLevelFilter.SelectedIndex -lt 0) { $controls.ImeLevelFilter.SelectedIndex = 0 }
        update-ime-grid
        $controls.ImeText.Text = 'Select an IME row to see the full details.'
        $script:AppsRowsFull = @($analysis.Applications)
        update-apps-grid
        set-grid 'AutopatchGrid' (@($analysis.Inventory.autopatch_logs) | ForEach-Object {
            Get-SedaKeywordLogEvents -Paths @($_) -Area 'Autopatch'
        })
        set-grid 'CollectionErrorsGrid' $analysis.ResultsXml.Errors
        set-grid 'WingetGrid' (@($analysis.Inventory.winget_logs) | ForEach-Object {
            Get-SedaKeywordLogEvents -Paths @($_) -Area 'Winget'
        })
        set-grid 'DriversGrid' $analysis.Drivers
        set-grid 'WifiGrid' $analysis.WifiProfiles
        $controls.AppsText.Text = "Apps: $(@($analysis.Applications).Count); Drivers: $(@($analysis.Drivers).Count); WiFi profiles: $(@($analysis.WifiProfiles).Count)"

        $controls.BatteryHealthText.Text = if ($analysis.Hardware.Battery.HealthPct -gt 0) { "$($analysis.Hardware.Battery.HealthPct)%" } else { 'N/A' }
        set-grid 'BatteryGrid' (ConvertTo-SedaKeyValueRows $analysis.Hardware.Battery)
        set-grid 'FirewallGrid' $analysis.Hardware.FirewallProfiles
        set-grid 'CertGrid' $analysis.Hardware.Certificates
        $controls.TabHardware.Header = "Hardware & Security"
        $controls.HardwareText.Text = 'Battery:' + [Environment]::NewLine + (ConvertTo-SedaKeyValueText $analysis.Hardware.Battery)

        $script:SearchRowsFull = @($analysis.Insights.SearchRows)
        update-search-grid
        $controls.SearchText.Text = 'Search corpus:' + [Environment]::NewLine + (ConvertTo-SedaTextTable -Rows $analysis.Insights.SearchRows -Properties @('Area','Severity','Title','Detail','Source'))
        $script:ZipRowsFull = @($analysis.ZipInfo.AllFiles | ForEach-Object {
            $item = Get-Item -LiteralPath $_ -ErrorAction SilentlyContinue
            New-SedaObject @{ File=[System.IO.Path]::GetFileName($_); Category=(Get-SedaZipCategory -RelativePath $_); SizeKB=if($item){[math]::Round($item.Length/1KB,1)}else{0}; Path=$_ }
        })
        $controls.ZipCategoryFilter.Items.Clear()
        [void]$controls.ZipCategoryFilter.Items.Add('All')
        foreach ($category in @($script:ZipRowsFull | Select-Object -ExpandProperty Category -Unique | Sort-Object)) { if ($category) { [void]$controls.ZipCategoryFilter.Items.Add([string]$category) } }
        $controls.ZipCategoryFilter.SelectedIndex = 0
        update-zip-grid
        $controls.AiText.Text = 'AI analysis is ready. Select a provider and click Analyze.'
        $controls.Win11StatusText.Text = "Status: $($analysis.Win11Compatibility.Status)"
        set-grid 'Win11Grid' (@($analysis.Win11Compatibility.HardwareReadiness.Checks) + @($analysis.Win11Compatibility.Indicators))
        $controls.TabWin11.Header = "Win11 Readiness ($($analysis.Win11Compatibility.Status))"
        $controls.Win11Text.Text = "HardwareReadiness.ps1: $($analysis.Win11Compatibility.HardwareReadiness.ReturnResult) (returnCode=$($analysis.Win11Compatibility.HardwareReadiness.ReturnCode))" + [Environment]::NewLine + [Environment]::NewLine + (($analysis.Win11Compatibility.HardwareReadiness.Notes) -join [Environment]::NewLine) + [Environment]::NewLine + [Environment]::NewLine + 'AppCompat indicators:' + [Environment]::NewLine + (ConvertTo-SedaTextTable -Rows $analysis.Win11Compatibility.Indicators -Properties @('TargetVersion','IsBlocking','ReasonText','SourceFile'))
        $controls.BtnExportHtml.IsEnabled = $true
        $controls.BtnExportAnon.IsEnabled = $true
        $controls.BtnOpenExtractDir.IsEnabled = [bool]$analysis.ExtractDir
        $controls.BtnOpenWuLog.IsEnabled = [bool]$analysis.WindowsUpdate.GeneratedLogPath
        $controls.TxtFile.Text = "  $($analysis.ZipInfo.ZipName)"
    }

    foreach ($item in $controls.AiProvider.Items) {
        if ([string]$item.Content -eq $script:CurrentAiConfig.Provider) { $controls.AiProvider.SelectedItem = $item }
    }
    if (-not $controls.AiProvider.SelectedItem -and $controls.AiProvider.Items.Count -gt 0) { $controls.AiProvider.SelectedIndex = 0 }
    $controls.AiModel.Text = [string]$script:CurrentAiConfig.Model
    $controls.AiCredential.Text = if ($script:CurrentAiConfig.Provider -eq 'ollama') { [string]$script:CurrentAiConfig.OllamaUrl } elseif ($script:CurrentAiConfig.RememberApiKey) { [string]$script:CurrentAiConfig.ApiKey } else { '' }
    $controls.AiTokens.Text = [string]$script:CurrentAiConfig.MaxTokens
    $controls.AiTokensDownButton.Add_Click({ set-ai-token-value -256 })
    $controls.AiTokensUpButton.Add_Click({ set-ai-token-value 256 })
    $controls.AiTokens.Add_PreviewTextInput({
        param($sender, $eventArgs)
        if ($eventArgs.Text -notmatch '^[0-9]+$') {
            $eventArgs.Handled = $true
        }
    })
    $controls.AiRemember.IsChecked = [bool]$script:CurrentAiConfig.RememberApiKey
    set-ai-defaults (get-ai-provider) $false
    $controls.AiText.Text = 'Load diagnostics, then run AI analysis with Claude, OpenAI, or Ollama.'

    $controls.AiProvider.Add_SelectionChanged({
        set-ai-defaults (get-ai-provider) $true
    })

    register-grid-formatting @(
        'InsightsGrid','HealthGrid','WuPoliciesGrid','WuEventsGrid','MdmPoliciesGrid','MdmBlockedGpsGrid','EventLogSummaryGrid','EventLogGrid',
        'ImeGrid','AppsGrid','AutopatchGrid','CollectionErrorsGrid','WingetGrid','DriversGrid','WifiGrid','BatteryGrid','FirewallGrid','CertGrid','ComplianceGrid','EnrollmentsGrid','SearchGrid','ZipFilesGrid','Win11Grid'
    )

    $controls.EventLogSelector.Add_SelectionChanged({ update-event-grid })
    $controls.EventLevelFilter.Add_SelectionChanged({ update-event-grid })
    $controls.EventSearch.Add_TextChanged({ update-event-grid })
    $controls.ImeThemeFilter.Add_SelectionChanged({ update-ime-grid })
    $controls.ImeLevelFilter.Add_SelectionChanged({ update-ime-grid })
    $controls.ImeSearch.Add_TextChanged({ update-ime-grid })
    $controls.AppsSearch.Add_TextChanged({ update-apps-grid })
    $controls.MdmPolicyFilter.Add_TextChanged({ update-mdm-policy-grid })
    $controls.MdmShowInternalKnobs.Add_Click({ update-mdm-policy-grid })
    $controls.SearchFilter.Add_TextChanged({ update-search-grid })
    $controls.ZipCategoryFilter.Add_SelectionChanged({ update-zip-grid })
    $controls.ZipSearch.Add_TextChanged({ update-zip-grid })

    $controls.EventLogGrid.Add_SelectionChanged({
        $controls.EventDetailText.Text = get-selected-text $controls.EventLogGrid.SelectedItem
    })
    $controls.ImeGrid.Add_SelectionChanged({
        $controls.ImeText.Text = get-selected-text $controls.ImeGrid.SelectedItem
    })
    $controls.AppsGrid.Add_SelectionChanged({
        $controls.AppsText.Text = get-selected-text $controls.AppsGrid.SelectedItem
    })
    $controls.HealthGrid.Add_SelectionChanged({
        $controls.HealthText.Text = get-selected-text $controls.HealthGrid.SelectedItem
    })
    $controls.WuEventsGrid.Add_SelectionChanged({
        $controls.WuText.Text = get-selected-text $controls.WuEventsGrid.SelectedItem
    })
    $controls.Win11Grid.Add_SelectionChanged({
        $controls.Win11Text.Text = get-selected-text $controls.Win11Grid.SelectedItem
    })
    foreach ($pair in @(
        @('MdmPoliciesGrid','MdmDiagText'), @('MdmBlockedGpsGrid','MdmDiagText'), @('MdmSourcesGrid','MdmDiagText'), @('MdmCabFilesGrid','MdmDiagText'),
        @('DriversGrid','AppsText'), @('WifiGrid','AppsText'), @('AutopatchGrid','AppsText'), @('CollectionErrorsGrid','AppsText'), @('WingetGrid','AppsText'),
        @('BatteryGrid','HardwareText'), @('FirewallGrid','HardwareText'), @('CertGrid','HardwareText'),
        @('ComplianceGrid','ComplianceText'), @('EnrollmentsGrid','EnrollmentsText'),
        @('SearchGrid','SearchDetailText'), @('ZipFilesGrid','ZipDetailText'), @('DeviceGrid','DeviceText'), @('InsightsGrid','InsightsText')
    )) {
        $gridName = $pair[0]
        $textName = $pair[1]
        $controls[$gridName].Add_SelectionChanged({
            param($sender, $eventArgs)
            $controls[$textName].Text = get-selected-text $sender.SelectedItem
        }.GetNewClosure())
    }

    $controls.BtnCopySelected.Add_Click({
        $grid = get-selected-grid
        if (-not $grid -or -not $grid.SelectedItem) { return }
        [System.Windows.Clipboard]::SetText((get-selected-text $grid.SelectedItem))
        set-status 'Selected row copied.' $false
    })

    $controls.BtnExportGridCsv.Add_Click({
        $grid = get-selected-grid
        if (-not $grid) { return }
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Title = 'Export current grid as CSV'
        $dialog.Filter = 'CSV files (*.csv)|*.csv|All files (*.*)|*.*'
        $dialog.FileName = 'SmartEndpointDiagnostics_grid.csv'
        if ($dialog.ShowDialog() -ne $true) { return }
        try {
            export-grid-csv -Grid $grid -Path $dialog.FileName
            set-status 'Grid exported to CSV.' $false
        } catch {
            [System.Windows.MessageBox]::Show($_.Exception.Message, 'CSV export', 'OK', 'Error') | Out-Null
        }
    })

    $controls.BtnOpenExtractDir.Add_Click({
        if ($script:CurrentAnalysis) { open-path $script:CurrentAnalysis.ExtractDir }
    })

    $controls.BtnOpenWuLog.Add_Click({
        if ($script:CurrentAnalysis) { open-path $script:CurrentAnalysis.WindowsUpdate.GeneratedLogPath }
    })

    $controls.BtnOpenAppLog.Add_Click({
        open-path $script:LogPath
    })

    $controls.BtnExtractCab.Add_Click({
        if (-not $script:CurrentAnalysis -or -not @($script:CurrentAnalysis.Inventory.cab).Count) { return }
        $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
        $dialog.Description = 'Select a folder for extracted MDM diagnostics CAB files'
        if ($dialog.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        try {
            show-overlay 'Extracting MDM diagnostics CAB' 'Running expand.exe...'
            $result = Expand-SedaCabFile -CabPath @($script:CurrentAnalysis.Inventory.cab)[0] -DestinationPath $dialog.SelectedPath
            hide-overlay
            $controls.MdmCabStatus.Text = $result.Status
            [System.Windows.MessageBox]::Show($result.Status, 'CAB extraction', 'OK', 'Information') | Out-Null
        } catch {
            hide-overlay
            [System.Windows.MessageBox]::Show($_.Exception.Message, 'CAB extraction error', 'OK', 'Error') | Out-Null
        }
    })

    $controls.BtnAnalyzeZip.Add_Click({
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
        $dialog.Title = 'Select Intune Device Diagnostics ZIP'
        $dialog.Filter = 'ZIP files (*.zip)|*.zip|All files (*.*)|*.*'
        if ($dialog.ShowDialog() -ne $true) { return }
        try {
            show-overlay 'Analyzing diagnostics ZIP' 'Expanding and parsing diagnostic files...'
            set-status 'Analyzing ZIP...' $true
            $script:CurrentAnalysis = Invoke-SedaAnalysis -Path $dialog.FileName
            set-status 'Rendering analysis results...' $true
            render-analysis $script:CurrentAnalysis
            set-status 'Analysis complete.' $false
        } catch {
            Write-SedaLog -Level ERROR -Message 'Analyze ZIP failed.' -Exception $_.Exception
            set-status 'Analysis failed.' $false
            hide-overlay
            [System.Windows.MessageBox]::Show($_.Exception.Message, 'Analysis Error', 'OK', 'Error') | Out-Null
        } finally {
            hide-overlay
        }
    })

    $controls.BtnAnalyzeLocal.Add_Click({
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Title = 'Save collected diagnostics ZIP as'
        $dialog.Filter = 'ZIP files (*.zip)|*.zip|All files (*.*)|*.*'
        $dialog.FileName = ('{0}_DiagLogs_{1:yyyyMMdd_HHmmss}.zip' -f $env:COMPUTERNAME, (Get-Date))
        if ($dialog.ShowDialog() -ne $true) { return }
        try {
            show-overlay 'Collecting local diagnostics' 'Collecting local endpoint data...'
            set-status 'Collecting local endpoint diagnostics...' $true
            $zip = Invoke-SedaLocalCollection -ZipPath $dialog.FileName
            set-status 'Analyzing collected ZIP...' $true
            $script:CurrentAnalysis = Invoke-SedaAnalysis -Path $zip
            set-status 'Rendering analysis results...' $true
            render-analysis $script:CurrentAnalysis
            set-status 'Local collection and analysis complete.' $false
        } catch {
            Write-SedaLog -Level ERROR -Message 'Local collection failed.' -Exception $_.Exception
            set-status 'Local collection failed.' $false
            hide-overlay
            [System.Windows.MessageBox]::Show($_.Exception.Message, 'Collection Error', 'OK', 'Error') | Out-Null
        } finally {
            hide-overlay
        }
    })

    $controls.BtnExportHtml.Add_Click({
        if (-not $script:CurrentAnalysis) { return }
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Title = 'Save HTML report'
        $dialog.Filter = 'HTML files (*.html)|*.html|All files (*.*)|*.*'
        $dialog.FileName = ([System.IO.Path]::GetFileNameWithoutExtension($script:CurrentAnalysis.ZipInfo.ZipName) + '_report.html')
        if ($dialog.ShowDialog() -ne $true) { return }
        try {
            show-overlay 'Exporting report' 'Writing standalone HTML report...'
            Export-SedaHtmlReport -Analysis $script:CurrentAnalysis -Path $dialog.FileName | Out-Null
            hide-overlay
            [System.Windows.MessageBox]::Show("Report saved:`n$($dialog.FileName)", 'Report exported', 'OK', 'Information') | Out-Null
        } catch {
            Write-SedaLog -Level ERROR -Message 'HTML export failed.' -Exception $_.Exception
            hide-overlay
            [System.Windows.MessageBox]::Show($_.Exception.Message, 'Export Error', 'OK', 'Error') | Out-Null
        } finally {
            hide-overlay
        }
    })

    $controls.BtnExportAnon.Add_Click({
        if (-not $script:CurrentAnalysis) { return }
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Title = 'Save anonymized ZIP'
        $dialog.Filter = 'ZIP files (*.zip)|*.zip|All files (*.*)|*.*'
        $dialog.FileName = ([System.IO.Path]::GetFileNameWithoutExtension($script:CurrentAnalysis.ZipInfo.ZipName) + '_anonymized.zip')
        if ($dialog.ShowDialog() -ne $true) { return }
        try {
            show-overlay 'Exporting anonymized ZIP' 'Redacting common identifiers and writing ZIP...'
            Export-SedaAnonymizedZip -SourceZip $script:CurrentAnalysis.SourceZipPath -OutputZip $dialog.FileName | Out-Null
            hide-overlay
            [System.Windows.MessageBox]::Show("Anonymized ZIP saved:`n$($dialog.FileName)`n`nReview before sharing.", 'Anonymized ZIP exported', 'OK', 'Information') | Out-Null
        } catch {
            Write-SedaLog -Level ERROR -Message 'Anonymized ZIP export failed.' -Exception $_.Exception
            hide-overlay
            [System.Windows.MessageBox]::Show($_.Exception.Message, 'Export Error', 'OK', 'Error') | Out-Null
        } finally {
            hide-overlay
        }
    })

    $controls.BtnRunAi.Add_Click({
        if (-not $script:CurrentAnalysis) { return }
        try {
            $config = get-ai-config-from-ui
            if ($config.Provider -ne 'ollama' -and [string]::IsNullOrWhiteSpace($config.ApiKey)) {
                [System.Windows.MessageBox]::Show('Enter an API key or switch to Ollama.', 'AI configuration', 'OK', 'Warning') | Out-Null
                return
            }
            Save-SedaAIConfig -Config $config
            $controls.BtnRunAi.IsEnabled = $false
            show-overlay 'Running AI analysis' "Sending diagnostic context to $($config.Provider)..."
            set-status 'Running AI analysis...' $true
            $result = Invoke-SedaAIAnalysis -Analysis $script:CurrentAnalysis -Config $config
            $controls.AiText.Text = [string]$result
            set-status 'AI analysis complete.' $false
        } catch {
            Write-SedaLog -Level ERROR -Message 'AI analysis failed.' -Exception $_.Exception
            set-status 'AI analysis failed.' $false
            hide-overlay
            [System.Windows.MessageBox]::Show($_.Exception.Message, 'AI Error', 'OK', 'Error') | Out-Null
        } finally {
            $controls.BtnRunAi.IsEnabled = $true
            hide-overlay
        }
    })

    $controls.BtnCopyAi.Add_Click({
        if (-not [string]::IsNullOrWhiteSpace($controls.AiText.Text)) {
            [System.Windows.Clipboard]::SetText($controls.AiText.Text)
            set-status 'AI analysis copied to clipboard.' $false
        }
    })

    $controls.BtnSaveAi.Add_Click({
        if ([string]::IsNullOrWhiteSpace($controls.AiText.Text)) { return }
        $dialog = New-Object Microsoft.Win32.SaveFileDialog
        $dialog.Title = 'Save AI analysis'
        $dialog.Filter = 'Text files (*.txt)|*.txt|Markdown files (*.md)|*.md|All files (*.*)|*.*'
        $dialog.FileName = 'SmartEndpointDiagnostics_AI_Analysis.txt'
        if ($dialog.ShowDialog() -ne $true) { return }
        try {
            Set-Content -LiteralPath $dialog.FileName -Value $controls.AiText.Text -Encoding UTF8
            set-status 'AI analysis saved.' $false
        } catch {
            [System.Windows.MessageBox]::Show($_.Exception.Message, 'Save Error', 'OK', 'Error') | Out-Null
        }
    })

    $controls.BtnReset.Add_Click({
        $script:CurrentAnalysis = $null
        $script:EventRowsFull = @()
        $script:ImeRowsFull = @()
        $script:AppsRowsFull = @()
        $script:MdmPoliciesFull = @()
        $script:SearchRowsFull = @()
        $script:ZipRowsFull = @()
        $script:CurrentGrid = $null
        $script:AiFocusContext = ''
        foreach ($name in @('SummaryText','InsightsText','HealthText','WuText','DeviceText','EnrollmentText','ComplianceText','EnrollmentsText','MdmDiagText','EventLogText','EventDetailText','ImeText','AppsText','HardwareText','SearchText','SearchDetailText','ZipDetailText','AiText','Win11Text')) { if ($controls[$name]) { $controls[$name].Text = '' } }
        foreach ($name in @('SummaryDeviceGrid','SummaryConnectionGrid','InsightsGrid','HealthGrid','WuPoliciesGrid','WuEventsGrid','DeviceGrid','ComplianceGrid','EnrollmentsGrid','MdmDeviceGrid','MdmConnectionGrid','MdmPoliciesGrid','MdmBlockedGpsGrid','MdmLapsGrid','MdmSourcesGrid','MdmCabFilesGrid','EventLogSummaryGrid','EventLogGrid','ImeGrid','AppsGrid','AutopatchGrid','CollectionErrorsGrid','WingetGrid','DriversGrid','WifiGrid','BatteryGrid','FirewallGrid','CertGrid','SearchGrid','ZipFilesGrid','Win11Grid')) {
            if ($controls[$name]) { $controls[$name].ItemsSource = $null }
        }
        foreach ($name in @('SummaryComputer','SummaryIP','SummaryOS','SummaryErrors','SummaryWarnings','SummaryFiles','MdmCabStatus','MdmPolicyCount','EventCount','ImeCount','AppsCount','SearchCount','ZipCount','BatteryHealthText','HealthBadge','ComplianceBadge','Win11StatusText')) {
            if ($controls[$name]) { $controls[$name].Text = '' }
        }
        foreach ($name in @('EventSearch','ImeSearch','AppsSearch','MdmPolicyFilter','SearchFilter','ZipSearch')) { if ($controls[$name]) { $controls[$name].Text = '' } }
        foreach ($name in @('EventLogSelector','ImeThemeFilter','ZipCategoryFilter')) { if ($controls[$name]) { $controls[$name].Items.Clear() } }
        foreach ($name in @('EventLevelFilter','ImeLevelFilter')) { if ($controls[$name]) { $controls[$name].SelectedIndex = 0 } }
        $controls.BtnExtractCab.IsEnabled = $false
        $controls.BtnOpenExtractDir.IsEnabled = $false
        $controls.BtnOpenWuLog.IsEnabled = $false
        $controls.TabSummary.Header = 'Summary'
        $controls.TabMdm.Header = 'MDM Diagnostics'
        $controls.TabWu.Header = 'Windows Update'
        $controls.TabEvent.Header = 'Event Log'
        $controls.TabIme.Header = 'IME Logs'
        $controls.TabApps.Header = 'Apps and Drivers'
        $controls.TabHardware.Header = 'Hardware & Security'
        $controls.TabInsights.Header = 'Insights'
        $controls.TabHealth.Header = 'Health'
        $controls.TabDevice.Header = 'Device Info'
        $controls.TabCompliance.Header = 'Compliance'
        $controls.TabEnrollments.Header = 'Enrollments'
        $controls.TabZip.Header = 'ZIP Files'
        $controls.TabSearch.Header = 'Search'
        $controls.TabWin11.Header = 'Win11 Readiness'
        $controls.TxtFile.Text = 'No file loaded'
        $controls.BtnExportHtml.IsEnabled = $false
        $controls.BtnExportAnon.IsEnabled = $false
        set-status 'Ready - Open an Intune Device Diagnostics ZIP file' $false
    })

    $controls.SummaryText.Text = "Open an Intune Device Diagnostics ZIP and click Analyze.`r`n`r`nThis PowerShell edition reads DSRegCmd, MDM enrollments, results.xml, IME logs, Windows Update, apps, drivers, WiFi and Windows 11 upgrade indicators."
    $controls.AiText.Text = 'Load diagnostics, then run AI analysis with Claude, OpenAI, or Ollama.'
    $window.Add_ContentRendered({
        if ($splash) {
            Hide-SmartM365GuiSplash -Splash $splash
        }
    })
    [void]$window.ShowDialog()
    if ($splash) {
        Close-SmartM365GuiSplash -Splash $splash
    }
}

if ($ValidateOnly) {
    Write-SedaLog -Level INFO -Message 'ValidateOnly completed.'
    Write-Output "$script:AppName validation completed."
    exit 0
}

if ($CollectLocal) {
    try {
        $createdZip = Invoke-SedaLocalCollection -ZipPath $ZipPath
    } catch {
        Write-SedaLog -Level ERROR -Message 'CLI local collection failed.' -Exception $_.Exception
        throw
    }
    if ($Cli) { Write-Output $createdZip }
    elseif ($createdZip) { $ZipPath = $createdZip }
}

if ($Cli) {
    if (-not $ZipPath) { throw 'Use -ZipPath with -Cli, or use -CollectLocal.' }
    try {
        $analysis = Invoke-SedaAnalysis -Path $ZipPath
        if ($ExportHtmlPath) { Export-SedaHtmlReport -Analysis $analysis -Path $ExportHtmlPath | Out-Null }
        if ($ExportAnonymizedZipPath) { Export-SedaAnonymizedZip -SourceZip $ZipPath -OutputZip $ExportAnonymizedZipPath | Out-Null }
        if ($RunAI) {
            $config = Get-SedaAIConfig
            if ($AIProvider) { $config.Provider = $AIProvider }
            if ($AIModel) { $config.Model = $AIModel }
            if ($AIApiKey) { $config.ApiKey = $AIApiKey }
            if ($AIOllamaUrl) { $config.OllamaUrl = $AIOllamaUrl }
            if ($AIMaxTokens -gt 0) { $config.MaxTokens = $AIMaxTokens }
            if ($AITemperature -ge 0) { $config.Temperature = $AITemperature }
            Write-SedaLog -Level INFO -Message "CLI AI analysis started. Provider=$($config.Provider); Model=$($config.Model)."
            Write-Output ''
            Write-Output 'AI Analysis'
            Write-Output '==========='
            Write-Output (Invoke-SedaAIAnalysis -Analysis $analysis -Config $config)
            Write-Output ''
            Write-SedaLog -Level INFO -Message 'CLI AI analysis completed.'
        }
    } catch {
        Write-SedaLog -Level ERROR -Message 'CLI analysis failed.' -Exception $_.Exception
        throw
    }
    Format-SedaSummaryText -Analysis $analysis
    exit 0
}

Write-SedaLog -Level INFO -Message 'Starting GUI.'
Start-SedaGui

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDYUPYkuSSTEvrb
# rZ72D2qRumN6LVRcwLLp3H5iKN5vlaCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCdLaS43Glx1hxqSSRz
# uIYzg4AjdMH8houZdc27biE3tjANBgkqhkiG9w0BAQEFAASCAYCyOuRFpgnP6Rh2
# R9OjmbkXtLzw/q+bVh39NwjuySO+6+NWvL4sJa35aGKMkh/BvFTU2bRGEwxkRSA0
# 6q0JkQTpkFH+B9WM5BSULvskWGKHcY58OvJsuFThRliwurIgLtJZbroHRHsytnqZ
# C4AbRyVsCedgSfV0zOzdH/L8NAOoZU6/gnzLJOp4HlJbEp4Q1GbJmSrBjDXdAOwM
# fpgomkY9W2KTcxHlkUxDSjhv8qcASlDcmP7oT4fLyyTil6vJrWW8p8ADKPnASJdY
# lA9tlnrAr/TNWarAON623RAZ8YiAObavMeWRWH50e4pxztCbOXLv2QlU3CYdRJSd
# 30FUsnn/QcOpCutz3vzgFf/wbF0s77vghyVUV4HU6C4AvcEnMGCrrTtXYpSGQ8Wi
# O3gTRoqbsPoJlqKzqVZYs8giL468xnYmngRUXu8y7nL0EnMHvtpoDmn/sKxSQiYV
# iQqFR9m1RTN5fs9zIHgIRK4HVEXqdrpJZ4wyUjPjdRfSNhrpOV4=
# SIG # End signature block
