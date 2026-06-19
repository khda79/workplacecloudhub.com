<#
.SYNOPSIS
Audits, applies, and restores a controlled thin-client workspace shell on Windows.

.DESCRIPTION
Smart ThinClient Shell prepares a Windows endpoint for Citrix, Azure Virtual Desktop,
web-only, or hybrid workspace access. Audit and Preview are read-only. Apply and
Restore are guarded by administrator checks, configuration gates, confirmation
phrases, rollback files, logs, and JSON/CSV evidence.
#>

[CmdletBinding()]
param(
    [switch]$Gui,
    [switch]$Cli,

    [ValidateSet('Audit', 'Preview', 'Apply', 'Restore')]
    [string]$Action = 'Audit',

    [ValidateSet('Auto', 'Citrix', 'AVD', 'WebOnly', 'Hybrid')]
    [string]$Profile = 'Auto',

    [ValidateSet('Auto', 'ExistingUser', 'DedicatedUser')]
    [string]$TargetUserMode = 'Auto',

    [string]$TargetUserName,
    [securestring]$DedicatedUserPassword,
    [string]$ConfigPath,
    [string]$OutputRoot,
    [string]$RollbackPath,
    [string]$ConfirmApply,
    [string]$ConfirmRestore,
    [switch]$NoSplash,
    [switch]$ValidateOnly
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:AppName = 'Smart ThinClient Shell'
$script:ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:RunId = Get-Date -Format 'yyyyMMdd-HHmmssfff'
$script:OutputRootOverride = $OutputRoot
$script:LogPath = $null
$script:OutputRoot = $null
$script:EvidenceJsonPath = $null
$script:EvidenceCsvPath = $null
$script:RollbackRoot = $null
$script:CliMode = [bool]$Cli

function Write-SmartThinClientLog {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $prefix = '{0} [{1}] ' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level
    foreach ($line in ($Message -split "(`r`n|`n|`r)")) {
        $entry = $prefix + $line
        if ($script:LogPath) {
            Add-Content -LiteralPath $script:LogPath -Value $entry -Encoding UTF8
        }
        if ($script:CliMode -or $Level -eq 'ERROR') {
            Write-Host $entry
        }
    }
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function ConvertTo-PlainText {
    param([securestring]$SecureString)
    if ($null -eq $SecureString) { return '' }
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        if ($bstr -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}

function Read-JsonFile {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    $raw = Get-Content -Raw -LiteralPath $Path
    if ([string]::IsNullOrWhiteSpace($raw)) { return $null }
    return $raw | ConvertFrom-Json
}

function Convert-ObjectToHashtable {
    param($InputObject)

    $table = [ordered]@{}
    if ($null -eq $InputObject) { return $table }
    foreach ($property in $InputObject.PSObject.Properties) {
        $table[$property.Name] = $property.Value
    }
    return $table
}

function Get-ConfigValue {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$Name,
        $Default = $null
    )

    if ($Config.Contains($Name)) {
        $value = $Config[$Name]
        if ($null -ne $value -and -not ([string]$value -eq '__USE_GLOBAL__')) {
            return $value
        }
    }
    return $Default
}

function Copy-Config {
    param([Parameter(Mandatory = $true)]$Config)

    $copy = [ordered]@{}
    foreach ($key in $Config.Keys) { $copy[$key] = $Config[$key] }
    return $copy
}

function Get-ProfileConfig {
    param([Parameter(Mandatory = $true)][string]$ProfileName)

    if ($ProfileName -eq 'Auto') { return $null }
    $profileRoot = Join-Path $script:ScriptRoot 'Profiles'
    $localPath = Join-Path $profileRoot ('{0}.profile.json' -f $ProfileName)
    $templatePath = Join-Path $profileRoot ('{0}.profile.json.template' -f $ProfileName)
    $profileConfig = Read-JsonFile -Path $localPath
    if ($null -eq $profileConfig) { $profileConfig = Read-JsonFile -Path $templatePath }
    return $profileConfig
}

function Merge-ProfileConfig {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$EffectiveProfile
    )

    $merged = Copy-Config -Config $Config
    $profileConfig = Get-ProfileConfig -ProfileName $EffectiveProfile
    if ($null -eq $profileConfig) { return $merged }

    foreach ($property in $profileConfig.PSObject.Properties) {
        if ($null -ne $property.Value -and ([string]$property.Value).Length -gt 0) {
            $merged[$property.Name] = $property.Value
        }
    }

    return $merged
}

function Resolve-Config {
    $defaults = [ordered]@{
        DefaultProfile = 'Auto'
        OutputRoot = 'C:\ProgramData\SmartThinClient\Shell'
        LogRetentionCount = 10
        ShellMode = 'Launcher'
        TargetUserMode = 'ExistingUser'
        TargetUserName = ''
        CreateDedicatedLocalUser = $false
        AllowApply = $false
        AllowRestore = $false
        RequireConfirmationPhrase = $true
        ApplyConfirmationPhrase = 'APPLY SMARTTHINCLIENT'
        RestoreConfirmationPhrase = 'RESTORE WINDOWS SHELL'
        CitrixWorkspacePath = ''
        AvdClientPath = ''
        WindowsAppPackageName = 'MicrosoftCorporationII.Windows365'
        WindowsAppLaunchUri = ''
        WebOnlyUrl = ''
        PreferredWorkspaceUrl = ''
        FallbackWebUrl = ''
        BrowserPath = ''
        BrowserKioskMode = $false
        AutoLaunchMode = 'RunKey'
        EnableShellLimitations = $false
        ShellRestrictionLevel = 'None'
        EnableAssignedAccess = $false
        AssignedAccessAppUserModelId = ''
        EnableShellLauncher = $false
        ShellLauncherDefaultAction = 0
        HybridSelectionAtStartup = $true
        HybridDefaultProvider = 'Citrix'
        WriteCsvEvidence = $true
        RollbackRetentionCount = 20
        LogoPath = 'Assets\WorkplaceCloudHub-lockup-WPF.png'
        DefaultLanguage = 'auto'
        ForceLanguage = ''
    }

    $templatePath = Join-Path $script:ScriptRoot 'SmartThinClient-Shell.config.template.json'
    $defaultConfigPath = Join-Path $script:ScriptRoot 'SmartThinClient-Shell.config.json'
    $effectivePath = $ConfigPath
    if ([string]::IsNullOrWhiteSpace($effectivePath)) { $effectivePath = $defaultConfigPath }

    $config = [ordered]@{}
    foreach ($key in $defaults.Keys) { $config[$key] = $defaults[$key] }

    foreach ($source in @($templatePath, $effectivePath)) {
        $sourceConfig = Read-JsonFile -Path $source
        if ($sourceConfig) {
            foreach ($property in $sourceConfig.PSObject.Properties) {
                $config[$property.Name] = $property.Value
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($script:OutputRootOverride)) { $config['OutputRoot'] = $script:OutputRootOverride }
    if ($TargetUserMode -ne 'Auto') { $config['TargetUserMode'] = $TargetUserMode }
    if (-not [string]::IsNullOrWhiteSpace($TargetUserName)) { $config['TargetUserName'] = $TargetUserName }
    return $config
}

function Initialize-Output {
    param([Parameter(Mandatory = $true)]$Config)

    $root = [string](Get-ConfigValue -Config $Config -Name 'OutputRoot' -Default 'C:\ProgramData\SmartThinClient\Shell')
    try {
        New-Item -Path $root -ItemType Directory -Force | Out-Null
        $probe = Join-Path $root ('.write-test-{0}.tmp' -f [guid]::NewGuid().ToString('N'))
        Set-Content -LiteralPath $probe -Value 'test' -Encoding ASCII
        Remove-Item -LiteralPath $probe -Force
    }
    catch {
        $root = Join-Path $env:LOCALAPPDATA 'SmartThinClient\Shell'
        New-Item -Path $root -ItemType Directory -Force | Out-Null
    }

    $logRoot = Join-Path $root 'Logs'
    $outputRoot = Join-Path $root 'Output'
    $rollbackRoot = Join-Path $root 'Rollback'
    New-Item -Path $logRoot, $outputRoot, $rollbackRoot -ItemType Directory -Force | Out-Null

    $script:OutputRoot = $root
    $script:RollbackRoot = $rollbackRoot
    $script:LogPath = Join-Path $logRoot ('SmartThinClient-Shell_{0}.log' -f $script:RunId)
    $script:EvidenceJsonPath = Join-Path $outputRoot ('SmartThinClient_Shell_{0}_{1}_{2}.json' -f $Action, $env:COMPUTERNAME, $script:RunId)
    $script:EvidenceCsvPath = Join-Path $outputRoot ('SmartThinClient_Shell_{0}_{1}_{2}.csv' -f $Action, $env:COMPUTERNAME, $script:RunId)
}

function Join-OptionalPath {
    param([string]$Root, [Parameter(Mandatory = $true)][string]$ChildPath)
    if ([string]::IsNullOrWhiteSpace($Root)) { return '' }
    return (Join-Path $Root $ChildPath)
}

function Find-FirstExistingPath {
    param([string[]]$Paths)
    foreach ($path in $Paths) {
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path -LiteralPath $path)) { return $path }
    }
    return ''
}

function Get-RegistryValueSnapshot {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $exists = Test-Path -LiteralPath $Path
    $value = $null
    $hasValue = $false
    $valueKind = ''
    if ($exists) {
        try {
            $item = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction Stop
            $value = $item.$Name
            $hasValue = $true
            $registryKey = Get-Item -LiteralPath $Path -ErrorAction Stop
            $valueKind = [string]$registryKey.GetValueKind($Name)
        }
        catch {
            $hasValue = $false
        }
    }

    [pscustomobject]@{
        Path = $Path
        Name = $Name
        KeyExists = $exists
        ValueExists = $hasValue
        ValueKind = $valueKind
        Value = $value
    }
}

function Restore-RegistryValue {
    param([Parameter(Mandatory = $true)]$Snapshot)

    if ($Snapshot.ValueExists) {
        if (-not (Test-Path -LiteralPath $Snapshot.Path)) {
            New-Item -Path $Snapshot.Path -Force | Out-Null
        }
        $kind = [string]$Snapshot.ValueKind
        if ([string]::IsNullOrWhiteSpace($kind)) { $kind = 'String' }
        New-ItemProperty -LiteralPath $Snapshot.Path -Name $Snapshot.Name -Value $Snapshot.Value -PropertyType $kind -Force | Out-Null
    }
    else {
        if (Test-Path -LiteralPath $Snapshot.Path) {
            Remove-ItemProperty -LiteralPath $Snapshot.Path -Name $Snapshot.Name -ErrorAction SilentlyContinue
        }
    }
}

function Set-StringRegistryValue {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Value
    )

    if (-not (Test-Path -LiteralPath $Path)) { New-Item -Path $Path -Force | Out-Null }
    New-ItemProperty -LiteralPath $Path -Name $Name -Value $Value -PropertyType String -Force | Out-Null
}

function Set-DwordRegistryValue {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Name,
        [int]$Value
    )

    if (-not (Test-Path -LiteralPath $Path)) { New-Item -Path $Path -Force | Out-Null }
    New-ItemProperty -LiteralPath $Path -Name $Name -Value $Value -PropertyType DWord -Force | Out-Null
}

function Get-CitrixWorkspaceSignal {
    param([Parameter(Mandatory = $true)]$Config)

    $candidatePaths = @(
        [string](Get-ConfigValue -Config $Config -Name 'CitrixWorkspacePath' -Default ''),
        (Join-OptionalPath -Root ${env:ProgramFiles(x86)} -ChildPath 'Citrix\ICA Client\SelfServicePlugin\SelfService.exe'),
        (Join-OptionalPath -Root ${env:ProgramFiles(x86)} -ChildPath 'Citrix\ICA Client\CDViewer.exe'),
        (Join-OptionalPath -Root ${env:ProgramFiles(x86)} -ChildPath 'Citrix\ICA Client\wfica32.exe'),
        (Join-OptionalPath -Root $env:ProgramFiles -ChildPath 'Citrix\ICA Client\SelfServicePlugin\SelfService.exe'),
        (Join-OptionalPath -Root $env:ProgramFiles -ChildPath 'Citrix\ICA Client\CDViewer.exe'),
        (Join-OptionalPath -Root $env:ProgramFiles -ChildPath 'Citrix\ICA Client\wfica32.exe')
    )

    $path = Find-FirstExistingPath -Paths $candidatePaths
    [pscustomobject]@{
        Installed = -not [string]::IsNullOrWhiteSpace($path)
        Path = $path
    }
}

function Get-AvdClientSignal {
    param([Parameter(Mandatory = $true)]$Config)

    $candidatePaths = @(
        [string](Get-ConfigValue -Config $Config -Name 'AvdClientPath' -Default ''),
        (Join-OptionalPath -Root ${env:ProgramFiles} -ChildPath 'Remote Desktop\msrdcw.exe'),
        (Join-OptionalPath -Root ${env:ProgramFiles(x86)} -ChildPath 'Remote Desktop\msrdcw.exe')
    )

    $path = Find-FirstExistingPath -Paths $candidatePaths
    $packageName = [string](Get-ConfigValue -Config $Config -Name 'WindowsAppPackageName' -Default 'MicrosoftCorporationII.Windows365')
    $windowsAppPackage = $null
    if (-not [string]::IsNullOrWhiteSpace($packageName)) {
        try { $windowsAppPackage = Get-AppxPackage -Name $packageName -ErrorAction Stop | Select-Object -First 1 }
        catch { $windowsAppPackage = $null }
    }

    [pscustomobject]@{
        Installed = (-not [string]::IsNullOrWhiteSpace($path)) -or ($null -ne $windowsAppPackage)
        ClassicClientPath = $path
        WindowsAppPackage = if ($windowsAppPackage) { $windowsAppPackage.Name } else { '' }
        WindowsAppVersion = if ($windowsAppPackage) { [string]$windowsAppPackage.Version } else { '' }
    }
}

function Get-BrowserSignal {
    param([Parameter(Mandatory = $true)]$Config)

    $candidatePaths = @(
        [string](Get-ConfigValue -Config $Config -Name 'BrowserPath' -Default ''),
        (Join-OptionalPath -Root ${env:ProgramFiles(x86)} -ChildPath 'Microsoft\Edge\Application\msedge.exe'),
        (Join-OptionalPath -Root ${env:ProgramFiles} -ChildPath 'Microsoft\Edge\Application\msedge.exe'),
        (Join-OptionalPath -Root ${env:ProgramFiles} -ChildPath 'Google\Chrome\Application\chrome.exe'),
        (Join-OptionalPath -Root ${env:ProgramFiles(x86)} -ChildPath 'Google\Chrome\Application\chrome.exe')
    )

    $path = Find-FirstExistingPath -Paths $candidatePaths
    [pscustomobject]@{
        Installed = -not [string]::IsNullOrWhiteSpace($path)
        Path = $path
    }
}

function Get-KioskCapability {
    $editionId = ''
    try { $editionId = (Get-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name EditionID -ErrorAction Stop).EditionID }
    catch { $editionId = '' }

    $assignedAccessCmdlet = [bool](Get-Command Set-AssignedAccess -ErrorAction SilentlyContinue)
    $shellLauncherClass = $false
    try {
        Get-CimClass -Namespace 'root\standardcimv2\embedded' -ClassName 'WESL_UserSetting' -ErrorAction Stop | Out-Null
        $shellLauncherClass = $true
    }
    catch {
        $shellLauncherClass = $false
    }

    [pscustomobject]@{
        EditionId = $editionId
        AssignedAccessCmdletAvailable = $assignedAccessCmdlet
        ShellLauncherClassAvailable = $shellLauncherClass
        SupportsProgressiveLockdown = $assignedAccessCmdlet -or $shellLauncherClass
    }
}

function Resolve-EffectiveProfile {
    param(
        [Parameter(Mandatory = $true)][string]$RequestedProfile,
        [Parameter(Mandatory = $true)]$Config
    )

    if ($RequestedProfile -ne 'Auto') { return $RequestedProfile }
    $defaultProfile = [string](Get-ConfigValue -Config $Config -Name 'DefaultProfile' -Default 'Auto')
    if ($defaultProfile -and $defaultProfile -ne 'Auto') { return $defaultProfile }
    return 'Hybrid'
}

function Get-TargetUserState {
    param([Parameter(Mandatory = $true)]$Config)

    $mode = [string](Get-ConfigValue -Config $Config -Name 'TargetUserMode' -Default 'ExistingUser')
    if ($TargetUserMode -ne 'Auto') { $mode = $TargetUserMode }
    $userName = [string](Get-ConfigValue -Config $Config -Name 'TargetUserName' -Default '')
    if (-not [string]::IsNullOrWhiteSpace($TargetUserName)) { $userName = $TargetUserName }
    if ([string]::IsNullOrWhiteSpace($userName)) { $userName = [Environment]::UserName }

    $localUser = $null
    if (Get-Command Get-LocalUser -ErrorAction SilentlyContinue) {
        try { $localUser = Get-LocalUser -Name $userName -ErrorAction Stop }
        catch { $localUser = $null }
    }

    [pscustomobject]@{
        Mode = $mode
        UserName = $userName
        LocalUserExists = $null -ne $localUser
        Enabled = if ($localUser) { [bool]$localUser.Enabled } else { $false }
        Sid = if ($localUser) { [string]$localUser.SID } else { '' }
    }
}

function New-DedicatedUserIfNeeded {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$TargetUser
    )

    if ($TargetUser.Mode -ne 'DedicatedUser') { return $false }
    if ($TargetUser.LocalUserExists) { return $false }
    if (-not [bool](Get-ConfigValue -Config $Config -Name 'CreateDedicatedLocalUser' -Default $false)) {
        throw "Dedicated user '$($TargetUser.UserName)' does not exist and CreateDedicatedLocalUser is false."
    }
    if ($null -eq $DedicatedUserPassword) {
        throw 'DedicatedUserPassword is required to create the dedicated local user. Do not store passwords in JSON.'
    }

    if (-not (Get-Command New-LocalUser -ErrorAction SilentlyContinue)) {
        throw 'New-LocalUser is not available in this PowerShell session.'
    }
    New-LocalUser -Name $TargetUser.UserName -Password $DedicatedUserPassword -FullName 'Smart ThinClient Shell' -Description 'Dedicated SmartThinClient workspace user' | Out-Null
    if (Get-Command Add-LocalGroupMember -ErrorAction SilentlyContinue) {
        Add-LocalGroupMember -Group 'Users' -Member $TargetUser.UserName -ErrorAction SilentlyContinue
    }
    Write-SmartThinClientLog -Message "Created dedicated local user '$($TargetUser.UserName)'."
    return $true
}

function Invoke-ThinClientAudit {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$RequestedProfile,
        [Parameter(Mandatory = $true)][string]$RequestedAction
    )

    $effectiveProfile = Resolve-EffectiveProfile -RequestedProfile $RequestedProfile -Config $Config
    $Config = Merge-ProfileConfig -Config $Config -EffectiveProfile $effectiveProfile
    Write-SmartThinClientLog -Message "Starting read-only audit. Profile=$effectiveProfile Action=$RequestedAction"

    $citrix = Get-CitrixWorkspaceSignal -Config $Config
    $avd = Get-AvdClientSignal -Config $Config
    $browser = Get-BrowserSignal -Config $Config
    $kiosk = Get-KioskCapability
    $targetUser = Get-TargetUserState -Config $Config
    $os = Get-CimInstance -ClassName Win32_OperatingSystem

    $result = [ordered]@{
        AppName = $script:AppName
        RunId = $script:RunId
        ComputerName = $env:COMPUTERNAME
        UserName = [Environment]::UserName
        IsAdministrator = Test-IsAdministrator
        RequestedAction = $RequestedAction
        RequestedProfile = $RequestedProfile
        EffectiveProfile = $effectiveProfile
        OSName = $os.Caption
        OSVersion = $os.Version
        OSBuildNumber = $os.BuildNumber
        EditionId = $kiosk.EditionId
        ShellMode = [string](Get-ConfigValue -Config $Config -Name 'ShellMode' -Default 'Launcher')
        TargetUserMode = $targetUser.Mode
        TargetUserName = $targetUser.UserName
        TargetUserExists = $targetUser.LocalUserExists
        TargetUserEnabled = $targetUser.Enabled
        TargetUserSid = $targetUser.Sid
        CitrixWorkspaceInstalled = $citrix.Installed
        CitrixWorkspacePath = $citrix.Path
        AvdClientInstalled = $avd.Installed
        AvdClassicClientPath = $avd.ClassicClientPath
        WindowsAppPackage = $avd.WindowsAppPackage
        WindowsAppVersion = $avd.WindowsAppVersion
        BrowserInstalled = $browser.Installed
        BrowserPath = $browser.Path
        WebOnlyUrlConfigured = -not [string]::IsNullOrWhiteSpace([string](Get-ConfigValue -Config $Config -Name 'WebOnlyUrl' -Default ''))
        AssignedAccessCmdletAvailable = $kiosk.AssignedAccessCmdletAvailable
        ShellLauncherClassAvailable = $kiosk.ShellLauncherClassAvailable
        SupportsProgressiveLockdown = $kiosk.SupportsProgressiveLockdown
        AllowApply = [bool](Get-ConfigValue -Config $Config -Name 'AllowApply' -Default $false)
        AllowRestore = [bool](Get-ConfigValue -Config $Config -Name 'AllowRestore' -Default $false)
        OutputRoot = $script:OutputRoot
        LogPath = $script:LogPath
        EvidenceJsonPath = $script:EvidenceJsonPath
        EvidenceCsvPath = $script:EvidenceCsvPath
        Status = if ($RequestedAction -eq 'Preview') { 'PreviewOnly' } else { 'AuditOnly' }
    }

    $audit = [pscustomobject]$result
    Export-Evidence -Evidence $audit
    Write-SmartThinClientLog -Message ("Audit completed. Json={0}" -f $script:EvidenceJsonPath)
    return $audit
}

function Export-Evidence {
    param([Parameter(Mandatory = $true)]$Evidence)

    $Evidence | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $script:EvidenceJsonPath -Encoding UTF8
    if (Test-Path -LiteralPath $script:EvidenceCsvPath) { Remove-Item -LiteralPath $script:EvidenceCsvPath -Force }
    $Evidence | Export-Csv -LiteralPath $script:EvidenceCsvPath -NoTypeInformation -Encoding UTF8
}

function Convert-AuditToText {
    param([Parameter(Mandatory = $true)]$Audit)

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($name in @(
        'Status', 'ComputerName', 'UserName', 'IsAdministrator', 'OSName', 'OSVersion',
        'EditionId', 'RequestedProfile', 'EffectiveProfile', 'ShellMode', 'TargetUserMode',
        'TargetUserName', 'TargetUserExists', 'CitrixWorkspaceInstalled', 'CitrixWorkspacePath',
        'AvdClientInstalled', 'AvdClassicClientPath', 'WindowsAppPackage', 'BrowserInstalled',
        'BrowserPath', 'AssignedAccessCmdletAvailable', 'ShellLauncherClassAvailable',
        'AllowApply', 'AllowRestore', 'LogPath', 'EvidenceJsonPath', 'EvidenceCsvPath'
    )) {
        $property = $Audit.PSObject.Properties[$name]
        if ($property) {
            $lines.Add(('{0}: {1}' -f $name, $property.Value))
        }
    }
    return ($lines -join [Environment]::NewLine)
}

function Get-LauncherCommandLine {
    $launcherPath = Join-Path $script:OutputRoot 'Launcher\SmartThinClient-LaunchWorkspace.ps1'
    return ('powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}"' -f $launcherPath)
}

function New-LauncherScript {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$EffectiveProfile,
        [Parameter(Mandatory = $true)]$Audit
    )

    $launcherRoot = Join-Path $script:OutputRoot 'Launcher'
    New-Item -Path $launcherRoot -ItemType Directory -Force | Out-Null
    $launcherPath = Join-Path $launcherRoot 'SmartThinClient-LaunchWorkspace.ps1'

    $citrixPath = [string]$Audit.CitrixWorkspacePath
    $avdPath = [string]$Audit.AvdClassicClientPath
    $browserPath = [string]$Audit.BrowserPath
    $webUrl = [string](Get-ConfigValue -Config $Config -Name 'WebOnlyUrl' -Default '')
    if ([string]::IsNullOrWhiteSpace($webUrl)) { $webUrl = [string](Get-ConfigValue -Config $Config -Name 'FallbackWebUrl' -Default '') }
    $windowsAppUri = [string](Get-ConfigValue -Config $Config -Name 'WindowsAppLaunchUri' -Default '')
    $browserKiosk = [bool](Get-ConfigValue -Config $Config -Name 'BrowserKioskMode' -Default $false)
    $hybridDefault = [string](Get-ConfigValue -Config $Config -Name 'HybridDefaultProvider' -Default 'Citrix')
    $hybridSelection = [bool](Get-ConfigValue -Config $Config -Name 'HybridSelectionAtStartup' -Default $true)

    $content = @"
param([ValidateSet('Citrix','AVD','WebOnly','Hybrid')][string]`$Profile = '$EffectiveProfile')
`$ErrorActionPreference = 'SilentlyContinue'

function Start-WorkspaceProcess {
    param([string]`$FilePath, [string]`$Arguments)
    if ([string]::IsNullOrWhiteSpace(`$FilePath)) { return `$false }
    if (-not (Test-Path -LiteralPath `$FilePath)) { return `$false }
    Start-Process -FilePath `$FilePath -ArgumentList `$Arguments
    return `$true
}

function Start-WebWorkspace {
    `$browser = '$browserPath'
    `$url = '$webUrl'
    if ([string]::IsNullOrWhiteSpace(`$url)) { return }
    if (-not [string]::IsNullOrWhiteSpace(`$browser) -and (Test-Path -LiteralPath `$browser)) {
        if ('$browserKiosk' -eq 'True') {
            Start-Process -FilePath `$browser -ArgumentList @('--kiosk', `$url, '--edge-kiosk-type=fullscreen')
        }
        else {
            Start-Process -FilePath `$browser -ArgumentList `$url
        }
    }
    else {
        Start-Process `$url
    }
}

function Start-CitrixWorkspace {
    if (Start-WorkspaceProcess -FilePath '$citrixPath' -Arguments '') { return }
    Start-WebWorkspace
}

function Start-AvdWorkspace {
    if (Start-WorkspaceProcess -FilePath '$avdPath' -Arguments '') { return }
    if (-not [string]::IsNullOrWhiteSpace('$windowsAppUri')) { Start-Process '$windowsAppUri'; return }
    Start-WebWorkspace
}

function Select-HybridProvider {
    if ('$hybridSelection' -ne 'True') { return '$hybridDefault' }
    Add-Type -AssemblyName PresentationFramework
    `$result = [System.Windows.MessageBox]::Show('Open Citrix? Choose No for Azure Virtual Desktop, Cancel for Web only.','Smart ThinClient Shell','YesNoCancel','Question')
    if (`$result -eq 'Yes') { return 'Citrix' }
    if (`$result -eq 'No') { return 'AVD' }
    return 'WebOnly'
}

if (`$Profile -eq 'Hybrid') { `$Profile = Select-HybridProvider }
switch (`$Profile) {
    'Citrix' { Start-CitrixWorkspace }
    'AVD' { Start-AvdWorkspace }
    'WebOnly' { Start-WebWorkspace }
    default { Start-WebWorkspace }
}
"@

    Set-Content -LiteralPath $launcherPath -Value $content -Encoding UTF8
    return $launcherPath
}

function New-RollbackState {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Audit,
        [Parameter(Mandatory = $true)][string]$EffectiveProfile
    )

    $runKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
    $explorerPolicy = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer'
    $systemPolicy = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System'
    $winlogon = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon'

    $snapshots = @(
        Get-RegistryValueSnapshot -Path $runKey -Name 'SmartThinClientShell'
        Get-RegistryValueSnapshot -Path $explorerPolicy -Name 'NoRun'
        Get-RegistryValueSnapshot -Path $explorerPolicy -Name 'NoControlPanel'
        Get-RegistryValueSnapshot -Path $explorerPolicy -Name 'NoViewContextMenu'
        Get-RegistryValueSnapshot -Path $systemPolicy -Name 'DisableCMD'
        Get-RegistryValueSnapshot -Path $winlogon -Name 'Shell'
        Get-RegistryValueSnapshot -Path $winlogon -Name 'Userinit'
    )

    [ordered]@{
        AppName = $script:AppName
        RunId = $script:RunId
        CreatedUtc = (Get-Date).ToUniversalTime().ToString('o')
        ComputerName = $env:COMPUTERNAME
        EffectiveProfile = $EffectiveProfile
        OutputRoot = $script:OutputRoot
        LauncherPath = Join-Path $script:OutputRoot 'Launcher\SmartThinClient-LaunchWorkspace.ps1'
        TargetUserName = $Audit.TargetUserName
        TargetUserSid = $Audit.TargetUserSid
        Config = $Config
        Registry = $snapshots
    }
}

function Save-RollbackState {
    param([Parameter(Mandatory = $true)]$State)

    $path = Join-Path $script:RollbackRoot ('SmartThinClient_Rollback_{0}_{1}.json' -f $env:COMPUTERNAME, $script:RunId)
    ([pscustomobject]$State) | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $path -Encoding UTF8
    Write-SmartThinClientLog -Message "Rollback state saved: $path"
    return $path
}

function Assert-ApplyAllowed {
    param([Parameter(Mandatory = $true)]$Config)

    if (-not (Test-IsAdministrator)) { throw 'Apply requires an elevated administrator PowerShell session.' }
    if (-not [bool](Get-ConfigValue -Config $Config -Name 'AllowApply' -Default $false)) { throw 'Apply is blocked because AllowApply is false in configuration.' }
    if ([bool](Get-ConfigValue -Config $Config -Name 'RequireConfirmationPhrase' -Default $true)) {
        $phrase = [string](Get-ConfigValue -Config $Config -Name 'ApplyConfirmationPhrase' -Default 'APPLY SMARTTHINCLIENT')
        if ($ConfirmApply -ne $phrase) { throw "Apply requires -ConfirmApply '$phrase'." }
    }
}

function Assert-RestoreAllowed {
    param([Parameter(Mandatory = $true)]$Config)

    if (-not (Test-IsAdministrator)) { throw 'Restore requires an elevated administrator PowerShell session.' }
    if (-not [bool](Get-ConfigValue -Config $Config -Name 'AllowRestore' -Default $false)) { throw 'Restore is blocked because AllowRestore is false in configuration.' }
    if ([bool](Get-ConfigValue -Config $Config -Name 'RequireConfirmationPhrase' -Default $true)) {
        $phrase = [string](Get-ConfigValue -Config $Config -Name 'RestoreConfirmationPhrase' -Default 'RESTORE WINDOWS SHELL')
        if ($ConfirmRestore -ne $phrase) { throw "Restore requires -ConfirmRestore '$phrase'." }
    }
}

function Apply-AutoLaunch {
    param([Parameter(Mandatory = $true)][string]$CommandLine)
    Set-StringRegistryValue -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' -Name 'SmartThinClientShell' -Value $CommandLine
    Write-SmartThinClientLog -Message 'Configured machine Run key for Smart ThinClient launcher.'
}

function Apply-ShellLimitations {
    param([Parameter(Mandatory = $true)]$Config)

    if (-not [bool](Get-ConfigValue -Config $Config -Name 'EnableShellLimitations' -Default $false)) {
        Write-SmartThinClientLog -Message 'Shell limitations disabled by configuration.'
        return
    }

    $level = [string](Get-ConfigValue -Config $Config -Name 'ShellRestrictionLevel' -Default 'Basic')
    $explorerPolicy = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer'
    $systemPolicy = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System'

    Set-DwordRegistryValue -Path $explorerPolicy -Name 'NoRun' -Value 1
    Set-DwordRegistryValue -Path $explorerPolicy -Name 'NoControlPanel' -Value 1
    if ($level -in @('Basic', 'Strict')) {
        Set-DwordRegistryValue -Path $explorerPolicy -Name 'NoViewContextMenu' -Value 1
    }
    if ($level -eq 'Strict') {
        Set-DwordRegistryValue -Path $systemPolicy -Name 'DisableCMD' -Value 1
    }
    Write-SmartThinClientLog -Message "Applied shell limitations. Level=$level"
}

function Apply-AssignedAccess {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Audit
    )

    if (-not [bool](Get-ConfigValue -Config $Config -Name 'EnableAssignedAccess' -Default $false)) { return }
    if (-not $Audit.AssignedAccessCmdletAvailable) { throw 'Assigned Access was requested but Set-AssignedAccess is not available on this OS.' }
    if ([string]::IsNullOrWhiteSpace($Audit.TargetUserName)) { throw 'Assigned Access requires TargetUserName.' }
    $aumid = [string](Get-ConfigValue -Config $Config -Name 'AssignedAccessAppUserModelId' -Default '')
    if ([string]::IsNullOrWhiteSpace($aumid)) { throw 'Assigned Access requires AssignedAccessAppUserModelId in configuration.' }

    Set-AssignedAccess -UserName $Audit.TargetUserName -AppUserModelId $aumid
    Write-SmartThinClientLog -Message "Applied Assigned Access for user '$($Audit.TargetUserName)' and app '$aumid'."
}

function Apply-ShellLauncher {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Audit,
        [Parameter(Mandatory = $true)][string]$LauncherCommand
    )

    if (-not [bool](Get-ConfigValue -Config $Config -Name 'EnableShellLauncher' -Default $false)) { return }
    if (-not $Audit.ShellLauncherClassAvailable) { throw 'Shell Launcher was requested but WESL_UserSetting is not available on this OS.' }
    if ([string]::IsNullOrWhiteSpace($Audit.TargetUserSid)) { throw 'Shell Launcher requires a target user SID.' }

    $defaultAction = [int](Get-ConfigValue -Config $Config -Name 'ShellLauncherDefaultAction' -Default 0)
    $shell = $LauncherCommand
    Invoke-CimMethod -Namespace 'root\standardcimv2\embedded' -ClassName 'WESL_UserSetting' -MethodName 'SetCustomShell' -Arguments @{
        Sid = $Audit.TargetUserSid
        Shell = $shell
        DefaultAction = $defaultAction
    } | Out-Null
    Write-SmartThinClientLog -Message "Applied Shell Launcher for SID '$($Audit.TargetUserSid)'."
}

function Invoke-ThinClientApply {
    param([Parameter(Mandatory = $true)]$Config)

    Assert-ApplyAllowed -Config $Config
    $audit = Invoke-ThinClientAudit -Config $Config -RequestedProfile $Profile -RequestedAction 'Preview'
    $effectiveProfile = [string]$audit.EffectiveProfile
    $Config = Merge-ProfileConfig -Config $Config -EffectiveProfile $effectiveProfile
    $targetUser = Get-TargetUserState -Config $Config
    [void](New-DedicatedUserIfNeeded -Config $Config -TargetUser $targetUser)
    $audit = Invoke-ThinClientAudit -Config $Config -RequestedProfile $effectiveProfile -RequestedAction 'Preview'

    $rollbackState = New-RollbackState -Config $Config -Audit $audit -EffectiveProfile $effectiveProfile
    $rollbackFile = Save-RollbackState -State $rollbackState
    $launcherPath = New-LauncherScript -Config $Config -EffectiveProfile $effectiveProfile -Audit $audit
    $launcherCommand = Get-LauncherCommandLine

    Apply-AutoLaunch -CommandLine $launcherCommand
    Apply-ShellLimitations -Config $Config
    Apply-AssignedAccess -Config $Config -Audit $audit
    Apply-ShellLauncher -Config $Config -Audit $audit -LauncherCommand $launcherCommand

    $result = [pscustomobject]([ordered]@{
        AppName = $script:AppName
        RunId = $script:RunId
        ComputerName = $env:COMPUTERNAME
        RequestedAction = 'Apply'
        EffectiveProfile = $effectiveProfile
        Status = 'Applied'
        LauncherPath = $launcherPath
        LauncherCommand = $launcherCommand
        RollbackPath = $rollbackFile
        OutputRoot = $script:OutputRoot
        LogPath = $script:LogPath
        EvidenceJsonPath = $script:EvidenceJsonPath
        EvidenceCsvPath = $script:EvidenceCsvPath
    })
    Export-Evidence -Evidence $result
    return $result
}

function Resolve-RollbackPath {
    if (-not [string]::IsNullOrWhiteSpace($RollbackPath)) { return $RollbackPath }
    $latest = Get-ChildItem -LiteralPath $script:RollbackRoot -Filter 'SmartThinClient_Rollback_*.json' -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($latest) { return $latest.FullName }
    throw 'No rollback file found. Use -RollbackPath to provide one.'
}

function Invoke-ThinClientRestore {
    param([Parameter(Mandatory = $true)]$Config)

    Assert-RestoreAllowed -Config $Config
    $path = Resolve-RollbackPath
    $rollback = Read-JsonFile -Path $path
    if ($null -eq $rollback) { throw "Rollback file is empty or invalid: $path" }

    foreach ($snapshot in @($rollback.Registry)) {
        Restore-RegistryValue -Snapshot $snapshot
    }

    $rollbackConfig = Convert-ObjectToHashtable -InputObject $rollback.Config
    if ([bool](Get-ConfigValue -Config $rollbackConfig -Name 'EnableAssignedAccess' -Default $false) -and (Get-Command Clear-AssignedAccess -ErrorAction SilentlyContinue)) {
        try { Clear-AssignedAccess | Out-Null } catch {}
    }

    if ([bool](Get-ConfigValue -Config $rollbackConfig -Name 'EnableShellLauncher' -Default $false) -and $rollback.TargetUserSid) {
        try {
            Invoke-CimMethod -Namespace 'root\standardcimv2\embedded' -ClassName 'WESL_UserSetting' -MethodName 'RemoveCustomShell' -Arguments @{ Sid = [string]$rollback.TargetUserSid } | Out-Null
        }
        catch {}
    }

    $result = [pscustomobject]([ordered]@{
        AppName = $script:AppName
        RunId = $script:RunId
        ComputerName = $env:COMPUTERNAME
        RequestedAction = 'Restore'
        Status = 'Restored'
        RollbackPath = $path
        OutputRoot = $script:OutputRoot
        LogPath = $script:LogPath
        EvidenceJsonPath = $script:EvidenceJsonPath
        EvidenceCsvPath = $script:EvidenceCsvPath
    })
    Export-Evidence -Evidence $result
    Write-SmartThinClientLog -Message "Restore completed from rollback: $path"
    return $result
}

function Show-ThinClientGui {
    param([Parameter(Mandatory = $true)]$Config)

    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    $splash = $null
    if (-not $NoSplash) {
        $splashPath = Join-Path $script:ScriptRoot 'SmartThinClient.GuiSplash.ps1'
        if (Test-Path -LiteralPath $splashPath) {
            . $splashPath
            $splashLogoPath = Join-Path $script:ScriptRoot 'Assets\WorkplaceCloudHub-lockup-WPF.png'
            $splashIconPath = Join-Path $script:ScriptRoot 'Assets\WorkplaceCloudHub.ico'
            $splash = Start-SmartThinClientGuiSplash -ProductName $script:AppName -MinimumDurationMs 1800 -LogoPath $splashLogoPath -WindowIconPath $splashIconPath
        }
    }

    [xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Smart ThinClient Shell" Height="760" Width="1120" MinHeight="650" MinWidth="940"
        Background="#F5F8FB" WindowStartupLocation="CenterScreen">
  <Grid Margin="18">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Border Grid.Row="0" Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="18">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="220"/>
        </Grid.ColumnDefinitions>
        <StackPanel>
          <Border Background="#E6F4FF" CornerRadius="8" Padding="10,4" HorizontalAlignment="Left">
            <TextBlock Text="Endpoint workspace" Foreground="#005A9E" FontWeight="SemiBold"/>
          </Border>
          <TextBlock Text="Smart ThinClient Shell" FontSize="28" FontWeight="SemiBold" Foreground="#1F2937" Margin="0,12,0,0"/>
          <TextBlock Text="Audit, apply, and restore a controlled Citrix, AVD, or web workspace shell." Foreground="#5F6B7A" FontSize="14" Margin="0,4,0,0"/>
        </StackPanel>
        <Border Grid.Column="1" Background="#F8FBFE" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="12">
          <Image Name="LogoImage" Stretch="Uniform"/>
        </Border>
      </Grid>
    </Border>
    <Border Grid.Row="1" Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,14,0,14">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock Text="Profile" Foreground="#5F6B7A" VerticalAlignment="Center" Margin="0,0,8,0"/>
          <ComboBox Name="ProfileCombo" Width="160" Height="34" SelectedIndex="0">
            <ComboBoxItem Content="Auto"/>
            <ComboBoxItem Content="Citrix"/>
            <ComboBoxItem Content="AVD"/>
            <ComboBoxItem Content="WebOnly"/>
            <ComboBoxItem Content="Hybrid"/>
          </ComboBox>
          <TextBlock Name="StatusText" Text="Ready" Foreground="#5F6B7A" VerticalAlignment="Center" Margin="18,0,0,0"/>
        </StackPanel>
        <StackPanel Grid.Column="1" Orientation="Horizontal">
          <Button Name="AuditButton" Content="Audit / Preview" Height="34" MinWidth="130" Background="#0078D4" Foreground="White" BorderBrush="#0078D4" Padding="16,0" Margin="0,0,8,0"/>
          <Button Name="ApplyButton" Content="Apply" Height="34" MinWidth="90" Padding="16,0" Margin="0,0,8,0"/>
          <Button Name="RestoreButton" Content="Restore" Height="34" MinWidth="90" Padding="16,0" Margin="0,0,8,0"/>
          <Button Name="OpenOutputButton" Content="Open output" Height="34" MinWidth="120" Padding="16,0"/>
        </StackPanel>
      </Grid>
    </Border>
    <Border Grid.Row="2" Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="12">
      <TextBox Name="OutputText" FontFamily="Consolas" FontSize="13" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" IsReadOnly="True" BorderThickness="0"/>
    </Border>
    <Border Grid.Row="3" Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="10" Margin="0,14,0,0">
      <TextBlock Text="Apply and Restore require admin rights, config gates, confirmation phrases, and rollback evidence." Foreground="#5F6B7A"/>
    </Border>
  </Grid>
</Window>
'@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $logoImage = $window.FindName('LogoImage')
    $profileCombo = $window.FindName('ProfileCombo')
    $statusText = $window.FindName('StatusText')
    $outputText = $window.FindName('OutputText')
    $auditButton = $window.FindName('AuditButton')
    $applyButton = $window.FindName('ApplyButton')
    $restoreButton = $window.FindName('RestoreButton')
    $openOutputButton = $window.FindName('OpenOutputButton')

    $logoPath = Join-Path $script:ScriptRoot ([string](Get-ConfigValue -Config $Config -Name 'LogoPath' -Default 'Assets\WorkplaceCloudHub-lockup-WPF.png'))
    if (Test-Path -LiteralPath $logoPath) {
        $bitmap = New-Object Windows.Media.Imaging.BitmapImage
        $bitmap.BeginInit()
        $bitmap.UriSource = New-Object Uri($logoPath, [UriKind]::Absolute)
        $bitmap.CacheOption = [Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bitmap.EndInit()
        $logoImage.Source = $bitmap
    }

    $auditButton.Add_Click({
        try {
            $selectedProfile = [string]$profileCombo.Text
            $statusText.Text = 'Running preview...'
            $audit = Invoke-ThinClientAudit -Config $Config -RequestedProfile $selectedProfile -RequestedAction 'Preview'
            $outputText.Text = Convert-AuditToText -Audit $audit
            $statusText.Text = 'Preview completed'
        }
        catch {
            $statusText.Text = 'Preview failed'
            $outputText.Text = $_.Exception.Message
            Write-SmartThinClientLog -Level 'ERROR' -Message $_.Exception.Message
        }
    })

    $applyButton.Add_Click({
        $outputText.Text = 'Apply is available from CLI with -Action Apply, configuration gates, and -ConfirmApply. This prevents accidental shell lockdown from the GUI.'
    })

    $restoreButton.Add_Click({
        $outputText.Text = 'Restore is available from CLI with -Action Restore, configuration gates, rollback evidence, and -ConfirmRestore.'
    })

    $openOutputButton.Add_Click({
        if (Test-Path -LiteralPath $script:OutputRoot) { Start-Process -FilePath explorer.exe -ArgumentList $script:OutputRoot }
    })

    $outputText.Text = 'Ready. Run Audit / Preview to collect read-only local signals.'
    if ($splash) { Close-SmartThinClientGuiSplash -Splash $splash }
    [void]$window.ShowDialog()
}

$config = Resolve-Config
if ($Profile -eq 'Auto') {
    $configuredProfile = [string](Get-ConfigValue -Config $config -Name 'DefaultProfile' -Default 'Auto')
    if ($configuredProfile -ne 'Auto') { $Profile = $configuredProfile }
}

if ($ValidateOnly) {
    Write-Host "$script:AppName validation completed."
    exit 0
}

Initialize-Output -Config $config

try {
    switch ($Action) {
        'Audit' {
            if ($Cli) {
                $audit = Invoke-ThinClientAudit -Config $config -RequestedProfile $Profile -RequestedAction 'Audit'
                Write-Host (Convert-AuditToText -Audit $audit)
            }
            else {
                Show-ThinClientGui -Config $config
            }
        }
        'Preview' {
            $audit = Invoke-ThinClientAudit -Config $config -RequestedProfile $Profile -RequestedAction 'Preview'
            Write-Host (Convert-AuditToText -Audit $audit)
        }
        'Apply' {
            $result = Invoke-ThinClientApply -Config $config
            Write-Host (Convert-AuditToText -Audit $result)
        }
        'Restore' {
            $result = Invoke-ThinClientRestore -Config $config
            Write-Host (Convert-AuditToText -Audit $result)
        }
    }
    exit 0
}
catch {
    Write-SmartThinClientLog -Level 'ERROR' -Message $_.Exception.Message
    Write-Error $_.Exception.Message
    exit 1
}
