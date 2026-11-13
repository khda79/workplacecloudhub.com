<#
.SYNOPSIS
Audits, applies, and restores a controlled thin-client workspace shell on Windows.

.DESCRIPTION
Smart ThinClient Shell prepares a Windows endpoint for Citrix, Azure Virtual Desktop,
web-only, or hybrid workspace access. Audit and CLI Preview are read-only. Visual
Preview can open live Windows settings on explicit request. Apply and
Restore are guarded by administrator checks, configuration gates, confirmation
phrases, rollback files, logs, and JSON/CSV evidence.
.NOTES
Version: 1.1.4
#>

[CmdletBinding()]
param(
    [switch]$Gui,
    [switch]$Cli,

    [ValidateSet('Audit', 'Preview', 'Launch', 'Apply', 'Restore')]
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
            $color = [ConsoleColor]::DarkGray
            if ($Level -eq 'WARN') { $color = [ConsoleColor]::Yellow }
            if ($Level -eq 'ERROR') { $color = [ConsoleColor]::Red }
            Write-Host $entry -ForegroundColor $color
        }
    }
}

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Initialize-ThinClientAdminTokenApi {
    if ('SmartThinClient.AdminToken' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;using System.ComponentModel;using System.Runtime.InteropServices;using System.Security.Principal;
namespace SmartThinClient {
 public static class AdminToken {
  [StructLayout(LayoutKind.Sequential)] struct SidAndAttributes {public IntPtr Sid;public uint Attributes;}
  [StructLayout(LayoutKind.Sequential)] struct TokenGroups {public uint Count;public SidAndAttributes First;}
  [DllImport("advapi32.dll",SetLastError=true)]
  [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
  [return:MarshalAs(UnmanagedType.Bool)]
  static extern bool GetTokenInformation(IntPtr token,int information,IntPtr buffer,int length,out int returned);
  public static bool IsAdministratorGroup(string sid,uint attributes) {
   return sid=="S-1-5-32-544" && (attributes & (4u|16u))!=0;
  }
  public static bool HasAdministratorSid(IntPtr token) {
   int length;GetTokenInformation(token,2,IntPtr.Zero,0,out length);
   int error=Marshal.GetLastWin32Error();
   if(error!=122 || length<=0 || length>1048576) throw new Win32Exception(error,"Cannot size current token groups.");
   IntPtr buffer=Marshal.AllocHGlobal(length);
   try {
    int returned;if(!GetTokenInformation(token,2,buffer,length,out returned)) throw new Win32Exception(Marshal.GetLastWin32Error());
    int offset=Marshal.OffsetOf(typeof(TokenGroups),"First").ToInt32(),stride=Marshal.SizeOf(typeof(SidAndAttributes));
    if(returned<offset || returned>length) throw new InvalidOperationException("Invalid token group buffer.");
    uint count=unchecked((uint)Marshal.ReadInt32(buffer));
    if(count>(returned-offset)/stride) throw new InvalidOperationException("Invalid token group count.");
    for(int i=0;i<count;i++) {
     var group=(SidAndAttributes)Marshal.PtrToStructure(IntPtr.Add(buffer,offset+i*stride),typeof(SidAndAttributes));
     if(group.Sid!=IntPtr.Zero && IsAdministratorGroup(new SecurityIdentifier(group.Sid).Value,group.Attributes)) return true;
    }
    return false;
   } finally {Marshal.FreeHGlobal(buffer);}
  }
 }
}
'@
}

function Get-ThinClientAdminTokenState {
    $identity = $null
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $elevated = (New-Object Security.Principal.WindowsPrincipal($identity)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
        $member = $elevated
        if (-not $member) {
            Initialize-ThinClientAdminTokenApi
            # WindowsIdentity.Groups omits deny-only SIDs from a filtered UAC token.
            $member = [SmartThinClient.AdminToken]::HasAdministratorSid($identity.Token)
        }
        [pscustomobject]@{IsAdministratorAccount=[bool]$member;IsElevated=[bool]$elevated}
    }
    catch { [pscustomobject]@{IsAdministratorAccount=$false;IsElevated=$false} }
    finally { if ($identity) { $identity.Dispose() } }
}

function Open-ThinClientAdminPanel {
    param([Parameter(Mandatory = $true)][string]$ScriptPath,[Parameter(Mandatory = $true)][string]$ConfigPath,[ValidateSet('Citrix','AVD','WebOnly','Hybrid','Auto')][string]$Profile='Auto')
    $role = Get-ThinClientAdminTokenState
    if (-not $role.IsAdministratorAccount) { throw 'The Admin control panel is available only to administrator accounts.' }
    if ([IO.Path]::GetFileName($ScriptPath) -ne 'SmartThinClient-Shell.ps1' -or -not (Test-Path -LiteralPath $ScriptPath -PathType Leaf) -or -not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) { throw 'The protected control-panel script or configuration is unavailable.' }
    $arguments = '-NoProfile -STA -WindowStyle Hidden -ExecutionPolicy Bypass -File "'+$ScriptPath+'" -Gui -Action Audit -Profile '+$Profile+' -ConfigPath "'+$ConfigPath+'"'
    $parameters = @{FilePath=(Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe');ArgumentList=$arguments;WindowStyle='Hidden';ErrorAction='Stop'}
    if (-not $role.IsElevated) { $parameters.Verb='RunAs' }
    Start-Process @parameters | Out-Null
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

function Get-RuntimeJsonPathFromTemplate {
    param([Parameter(Mandatory = $true)][string]$TemplatePath)

    if ($TemplatePath -match '\.template\.json$') {
        return ($TemplatePath -replace '\.template\.json$', '.json')
    }
    if ($TemplatePath -match '\.json\.template$') {
        return ($TemplatePath -replace '\.template$', '')
    }
    return $null
}

function Ensure-JsonFromTemplate {
    param(
        [Parameter(Mandatory = $true)][string]$TemplatePath,
        [string]$RuntimePath
    )

    if (-not (Test-Path -LiteralPath $TemplatePath)) { return $null }
    if ([string]::IsNullOrWhiteSpace($RuntimePath)) {
        $RuntimePath = Get-RuntimeJsonPathFromTemplate -TemplatePath $TemplatePath
    }
    if ([string]::IsNullOrWhiteSpace($RuntimePath)) { return $null }
    if (Test-Path -LiteralPath $RuntimePath) { return $RuntimePath }

    $runtimeDirectory = Split-Path -Parent $RuntimePath
    if (-not [string]::IsNullOrWhiteSpace($runtimeDirectory)) {
        New-Item -Path $runtimeDirectory -ItemType Directory -Force | Out-Null
    }
    Copy-Item -LiteralPath $TemplatePath -Destination $RuntimePath -Force:$false
    Write-SmartThinClientLog -Message ("Created local JSON from template: {0}" -f $RuntimePath)
    return $RuntimePath
}

function Merge-MissingJsonTemplateKeys {
    param(
        [Parameter(Mandatory = $true)][string]$TemplatePath,
        [Parameter(Mandatory = $true)][string]$RuntimePath
    )

    if (-not (Test-Path -LiteralPath $TemplatePath)) { return }
    if (-not (Test-Path -LiteralPath $RuntimePath)) { return }

    $template = Read-JsonFile -Path $TemplatePath
    $runtime = Read-JsonFile -Path $RuntimePath
    if ($null -eq $template -or $null -eq $runtime) { return }

    $runtimeTable = [ordered]@{}
    foreach ($property in $runtime.PSObject.Properties) {
        $runtimeTable[$property.Name] = $property.Value
    }

    $changed = $false
    foreach ($property in $template.PSObject.Properties) {
        if (-not $runtimeTable.Contains($property.Name)) {
            $runtimeTable[$property.Name] = $property.Value
            $changed = $true
        }
    }

    if ($changed) {
        ([pscustomobject]$runtimeTable) | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $RuntimePath -Encoding UTF8
        Write-SmartThinClientLog -Message ("Added missing JSON keys from template: {0}" -f $RuntimePath)
    }
}

function Initialize-LocalJsonFiles {
    param([string]$DefaultConfigPath)

    $templatePath = Join-Path $script:ScriptRoot 'SmartThinClient-Shell.config.template.json'
    [void](Ensure-JsonFromTemplate -TemplatePath $templatePath -RuntimePath $DefaultConfigPath)
    Merge-MissingJsonTemplateKeys -TemplatePath $templatePath -RuntimePath $DefaultConfigPath

    $profileRoot = Join-Path $script:ScriptRoot 'Profiles'
    if (Test-Path -LiteralPath $profileRoot) {
        foreach ($profileTemplate in Get-ChildItem -LiteralPath $profileRoot -Filter '*.json.template' -File -ErrorAction SilentlyContinue) {
            $runtimePath = Ensure-JsonFromTemplate -TemplatePath $profileTemplate.FullName
            if ($runtimePath) { Merge-MissingJsonTemplateKeys -TemplatePath $profileTemplate.FullName -RuntimePath $runtimePath }
        }
        foreach ($profileTemplate in Get-ChildItem -LiteralPath $profileRoot -Filter '*.template.json' -File -ErrorAction SilentlyContinue) {
            $runtimePath = Ensure-JsonFromTemplate -TemplatePath $profileTemplate.FullName
            if ($runtimePath) { Merge-MissingJsonTemplateKeys -TemplatePath $profileTemplate.FullName -RuntimePath $runtimePath }
        }
    }
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
        if ($null -ne $property.Value -and ([string]$property.Value).Length -gt 0 -and [string]$property.Value -ne '__USE_GLOBAL__') {
            if ($property.Name -eq 'WebShellSettingsButtons') {
                $buttons = ConvertTo-ThinClientSettingsButtons -Value (Get-ConfigValue -Config $merged -Name 'WebShellSettingsButtons' -Default $null)
                $overrides = ConvertTo-ThinClientSettingsButtons -Value $property.Value
                foreach ($key in $overrides.Keys) { $buttons[$key] = $overrides[$key] }
                $merged[$property.Name] = $buttons
                continue
            }
            $merged[$property.Name] = $property.Value
        }
    }

    return $merged
}

function Get-DemoWorkspaceUrl {
    param([ValidateSet('Citrix','AVD','WebOnly')][string]$Provider)
    switch ($Provider) {
        Citrix { return 'https://workplacecloudhub.com/demo/smartthinclient/citrix/' }
        AVD { return 'https://workplacecloudhub.com/demo/smartthinclient/avd/' }
        default { return '' }
    }
}

function Resolve-Config {
    $defaults = [ordered]@{
        DefaultProfile = 'Auto'
        OutputRoot = 'C:\ProgramData\SmartThinClient\Shell'
        LogRetentionCount = 10
        ShellMode = 'Launcher'
        EnableCustomShell = $false
        MachineRoot = (Join-Path $env:ProgramData 'SmartThinClient\Machine')
        TargetUserMode = 'ExistingUser'
        TargetUserName = ''
        CreateDedicatedLocalUser = $false
        AllowApply = $false
        AllowRestore = $false
        RequireConfirmationPhrase = $true
        ApplyConfirmationPhrase = 'APPLY'
        RestoreConfirmationPhrase = 'RESTORE'
        CitrixWorkspacePath = ''
        AvdClientPath = ''
        WindowsAppPackageName = 'MicrosoftCorporationII.Windows365'
        WindowsAppLaunchUri = ''
        PreferredAccessMode = 'Web'
        UseWebShell = $true
        CitrixWebUrl = Get-DemoWorkspaceUrl -Provider Citrix
        AvdWebUrl = Get-DemoWorkspaceUrl -Provider AVD
        WebShellPreferredProvider = 'Citrix'
        WebShellTitle = 'Smart ThinClient'
        WebShellHomeUrl = ''
        WebShellShowProviderButtons = $true
        WebShellAllowExternalBrowser = $true
        WebShellAllowPowerControls = $true
        WebShellAllowLimitedSettings = $true
        WebShellSettingsButtons = [ordered]@{Network=$true;Display=$true;Sound=$true;DateAndTime=$true;Bluetooth=$false;Printers=$false}
        WebShellAllowedSettingsPages = @('ms-settings:network-status', 'ms-settings:display', 'ms-settings:sound', 'ms-settings:dateandtime', 'ms-settings:bluetooth', 'ms-settings:printers')
        WebOnlyUrl = 'https://workplacecloudhub.com/'
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
    if (-not [string]::IsNullOrWhiteSpace($ConfigPath) -and -not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) { throw 'The supplied ConfigPath does not exist.' }
    if ([string]::IsNullOrWhiteSpace($effectivePath)) { $effectivePath = $defaultConfigPath }
    Initialize-LocalJsonFiles -DefaultConfigPath $defaultConfigPath

    $config = [ordered]@{}
    foreach ($key in $defaults.Keys) { $config[$key] = $defaults[$key] }

    foreach ($source in @($templatePath, $effectivePath)) {
        $sourceConfig = Read-JsonFile -Path $source
        if ($sourceConfig) {
            foreach ($property in $sourceConfig.PSObject.Properties) {
                $config[$property.Name] = $property.Value
            }
            if (-not $sourceConfig.PSObject.Properties['HybridDefaultProvider'] -and $sourceConfig.PSObject.Properties['WebShellPreferredProvider'] -and $sourceConfig.WebShellPreferredProvider -in @('Citrix','AVD','WebOnly')) {
                $config['HybridDefaultProvider'] = $sourceConfig.WebShellPreferredProvider
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($script:OutputRootOverride)) { $config['OutputRoot'] = $script:OutputRootOverride }
    if ($TargetUserMode -ne 'Auto') { $config['TargetUserMode'] = $TargetUserMode }
    if (-not [string]::IsNullOrWhiteSpace($TargetUserName)) { $config['TargetUserName'] = $TargetUserName }
    return $config
}

function Resolve-WritableOutputRoot {
    param([Parameter(Mandatory = $true)][string[]]$CandidateRoots)

    $seen = @{}
    foreach ($candidateRoot in $CandidateRoots) {
        if ([string]::IsNullOrWhiteSpace($candidateRoot)) { continue }
        if ($seen.ContainsKey($candidateRoot)) { continue }
        $seen[$candidateRoot] = $true

        try {
            New-Item -Path $candidateRoot -ItemType Directory -Force | Out-Null
            $probe = Join-Path $candidateRoot ('.write-test-{0}.tmp' -f [guid]::NewGuid().ToString('N'))
            Set-Content -LiteralPath $probe -Value 'test' -Encoding ASCII
            Remove-Item -LiteralPath $probe -Force
            return $candidateRoot
        }
        catch {
            Write-SmartThinClientLog -Level 'WARN' -Message ("Output root unavailable: {0}. {1}" -f $candidateRoot, $_.Exception.Message)
        }
    }

    throw 'No writable output root was found for logs and evidence.'
}

function Initialize-Output {
    param([Parameter(Mandatory = $true)]$Config)

    $root = [string](Get-ConfigValue -Config $Config -Name 'OutputRoot' -Default 'C:\ProgramData\SmartThinClient\Shell')
    $localAppDataRoot = Join-Path $env:LOCALAPPDATA 'SmartThinClient\Shell'
    $tempRoot = Join-Path ([IO.Path]::GetTempPath()) 'SmartThinClient\Shell'
    $root = Resolve-WritableOutputRoot -CandidateRoots @($root, $localAppDataRoot, $tempRoot)

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
        $registryKey = Get-Item -LiteralPath $Path -ErrorAction Stop
        if ($Name -in $registryKey.GetValueNames()) {
            $value = $registryKey.GetValue($Name, $null, [Microsoft.Win32.RegistryValueOptions]::DoNotExpandEnvironmentNames)
            $hasValue = $true
            $valueKind = [string]$registryKey.GetValueKind($Name)
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
            if ($null -ne (Get-ItemProperty -LiteralPath $Snapshot.Path).PSObject.Properties[$Snapshot.Name]) {
                Remove-ItemProperty -LiteralPath $Snapshot.Path -Name $Snapshot.Name -ErrorAction Stop
            }
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
        CustomShellPolicyAvailable = ($editionId -match '^(Professional|Enterprise|Education|IoTEnterprise)' -and (Test-Path -LiteralPath (Join-Path $env:WINDIR 'System32\gpedit.dll')))
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

function Get-ThinClientCurrentIdentity {
    $identity = $null
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $name = [string]$identity.Name
        $sid = if ($identity.User) { [string]$identity.User.Value } else { '' }
        $authority = if ($name.Contains('\')) { ($name -split '\\',2)[0] } else { '' }
        $kind = if ($authority -eq $env:COMPUTERNAME -or $authority -eq '.') { 'Local' }
            elseif ($authority -eq 'AzureAD') { 'Entra' }
            elseif ($authority -eq 'MicrosoftAccount') { 'MicrosoftAccount' }
            elseif ($authority -in @('NT AUTHORITY','NT SERVICE','BUILTIN','IIS APPPOOL','Window Manager')) { 'System' }
            elseif ($sid -like 'S-1-5-21-*' -and $authority) { 'Domain' }
            else { 'Other' }
        [pscustomobject]@{Name=$name;Sid=$sid;AccountType=$kind;Authenticated=[bool]$identity.IsAuthenticated}
    }
    catch { return $null }
    finally { if ($null -ne $identity) { $identity.Dispose() } }
}

function Resolve-ThinClientTargetIdentity {
    param([Parameter(Mandatory = $true)]$TargetUser, $CurrentIdentity)
    # Diagnostic identity only. Keep the operational local-user state/SID unchanged.
    if ($TargetUser.LocalUserExists) {
        $localName = if ([string]$TargetUser.UserName -like '*\*') { [string]$TargetUser.UserName } else { $env:COMPUTERNAME+'\'+$TargetUser.UserName }
        return [pscustomobject]@{Exists=$true;Enabled=[bool]$TargetUser.Enabled;AccountType='Local';Status=$(if($TargetUser.Enabled){'Detected'}else{'Disabled'});Name=$localName;Sid=[string]$TargetUser.Sid;Source='LocalAccountLookup'}
    }
    if ($TargetUser.Mode -eq 'DedicatedUser') {
        return [pscustomobject]@{Exists=$false;Enabled=$null;AccountType='Local';Status='NotVerified';Name=[string]$TargetUser.UserName;Sid='';Source='LocalAccountLookup'}
    }
    if ($null -eq $CurrentIdentity) { $CurrentIdentity = Get-ThinClientCurrentIdentity }
    if ($null -ne $CurrentIdentity -and $CurrentIdentity.Authenticated -and -not [string]::IsNullOrWhiteSpace($CurrentIdentity.Sid)) {
        $requested = [string]$TargetUser.UserName
        $current = [string]$CurrentIdentity.Name
        $shortName = ($current -split '\\')[-1]
        $unqualified = -not $requested.Contains('\') -and -not $requested.Contains('@')
        $same = $requested -ieq $current -or ($unqualified -and $requested -ieq $shortName)
        if ($requested.StartsWith('.\') -and $CurrentIdentity.AccountType -eq 'Local') { $same = $requested.Substring(2) -ieq $shortName }
        if ($same) { return [pscustomobject]@{Exists=$true;Enabled=$null;AccountType=[string]$CurrentIdentity.AccountType;Status='Detected';Name=$current;Sid=[string]$CurrentIdentity.Sid;Source='CurrentWindowsToken'} }
    }
    # Do not treat a different/unreachable domain identity as a nonexistent local account.
    [pscustomobject]@{Exists=$null;Enabled=$null;AccountType='Unknown';Status='NotVerified';Name=[string]$TargetUser.UserName;Sid='';Source='NotVerified'}
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
        Add-LocalGroupMember -SID 'S-1-5-32-545' -Member $TargetUser.UserName -ErrorAction Stop
    }
    Write-SmartThinClientLog -Message "Created dedicated local user '$($TargetUser.UserName)'."
    return $true
}

function Get-WorkspaceWebUrls {
    param($Config, [string]$EffectiveProfile)
    $urls = [ordered]@{}
    $keys = @{
        Citrix = @('CitrixWebUrl','WorkspaceUrl','StoreUrl')
        AVD = @('AvdWebUrl','FeedUrl')
        WebOnly = @('WebOnlyUrl','WebShellHomeUrl','FallbackWebUrl')
    }
    foreach ($provider in @('Citrix','AVD','WebOnly')) {
        $providerConfig = $Config
        if ($EffectiveProfile -eq 'Hybrid') { $providerConfig = Merge-ProfileConfig -Config $Config -EffectiveProfile $provider }
        $urls[$provider] = ''
        $demoUrl = Get-DemoWorkspaceUrl -Provider $provider
        $demoFallback = ''
        foreach ($key in $keys[$provider]) {
            $value = [string](Get-ConfigValue -Config $providerConfig -Name $key -Default '')
            # Demo defaults must not mask an existing custom legacy URL alias.
            if ($demoUrl -and $value -eq $demoUrl) { $demoFallback = $value; continue }
            if (-not [string]::IsNullOrWhiteSpace($value)) { $urls[$provider] = $value; break }
        }
        if (-not $urls[$provider]) { $urls[$provider] = $demoFallback }
    }
    return $urls
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
    $targetIdentity = Resolve-ThinClientTargetIdentity -TargetUser $targetUser
    $webUrls = Get-WorkspaceWebUrls -Config $Config -EffectiveProfile $effectiveProfile
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
        ShellScope = if((Get-ConfigValue -Config $Config -Name ShellMode -Default Launcher) -eq 'CustomShell'){'AllUsers'}else{'ConfiguredMode'}
        TargetUserMode = $targetUser.Mode
        TargetUserName = $targetUser.UserName
        TargetUserExists = $targetIdentity.Exists
        TargetUserEnabled = $targetIdentity.Enabled
        TargetUserSid = $targetUser.Sid
        TargetLocalUserExists = $targetUser.LocalUserExists
        TargetUserAccountType = $targetIdentity.AccountType
        TargetUserIdentityStatus = $targetIdentity.Status
        TargetUserResolvedName = $targetIdentity.Name
        TargetUserResolvedSid = $targetIdentity.Sid
        TargetUserIdentitySource = $targetIdentity.Source
        CitrixWorkspaceInstalled = $citrix.Installed
        CitrixWorkspacePath = $citrix.Path
        AvdClientInstalled = $avd.Installed
        AvdClassicClientPath = $avd.ClassicClientPath
        WindowsAppPackage = $avd.WindowsAppPackage
        WindowsAppVersion = $avd.WindowsAppVersion
        PreferredAccessMode = [string](Get-ConfigValue -Config $Config -Name 'PreferredAccessMode' -Default 'Web')
        UseWebShell = [bool](Get-ConfigValue -Config $Config -Name 'UseWebShell' -Default $true)
        CitrixWebUrlConfigured = -not [string]::IsNullOrWhiteSpace($webUrls.Citrix)
        AvdWebUrlConfigured = -not [string]::IsNullOrWhiteSpace($webUrls.AVD)
        WindowsAppLaunchUriConfigured = -not [string]::IsNullOrWhiteSpace([string](Get-ConfigValue -Config $Config -Name 'WindowsAppLaunchUri' -Default ''))
        BrowserInstalled = $browser.Installed
        BrowserPath = $browser.Path
        WebOnlyUrlConfigured = -not [string]::IsNullOrWhiteSpace($webUrls.WebOnly)
        AssignedAccessCmdletAvailable = $kiosk.AssignedAccessCmdletAvailable
        ShellLauncherClassAvailable = $kiosk.ShellLauncherClassAvailable
        CustomShellPolicyAvailable = [bool](Get-ResultValue -Result $kiosk -Name CustomShellPolicyAvailable -Default $false)
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
    if ($RequestedAction -eq 'Preview') { $audit | Add-Member NoteProperty PreviewPlan (Get-ThinClientPreviewPlan -Config $Config -Audit $audit) }
    Export-Evidence -Evidence $audit -Config $Config
    Write-SmartThinClientLog -Message ("Audit completed. Json={0}" -f $script:EvidenceJsonPath)
    return $audit
}

function Export-Evidence {
    param([Parameter(Mandatory = $true)]$Evidence, $Config = ([ordered]@{}))

    $Evidence | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $script:EvidenceJsonPath -Encoding UTF8
    if (-not (Get-ConfigValue -Config $Config -Name 'WriteCsvEvidence' -Default $true)) { return }
    if (Test-Path -LiteralPath $script:EvidenceCsvPath) { Remove-Item -LiteralPath $script:EvidenceCsvPath -Force }
    $Evidence | Export-Csv -LiteralPath $script:EvidenceCsvPath -NoTypeInformation -Encoding UTF8
}

function Convert-AuditToText {
    param([Parameter(Mandatory = $true)]$Audit)

    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($name in @(
        'Status', 'ComputerName', 'UserName', 'IsAdministrator', 'OSName', 'OSVersion',
        'EditionId', 'RequestedProfile', 'EffectiveProfile', 'ShellMode', 'TargetUserMode',
        'TargetUserName', 'TargetUserAccountType', 'TargetUserIdentityStatus', 'TargetUserResolvedName', 'TargetUserExists', 'CitrixWorkspaceInstalled', 'CitrixWorkspacePath',
        'AvdClientInstalled', 'AvdClassicClientPath', 'WindowsAppPackage', 'BrowserInstalled',
        'PreferredAccessMode', 'UseWebShell', 'CitrixWebUrlConfigured', 'AvdWebUrlConfigured',
        'BrowserPath', 'WebOnlyUrlConfigured', 'AssignedAccessCmdletAvailable', 'ShellLauncherClassAvailable',
        'AllowApply', 'AllowRestore', 'LogPath', 'EvidenceJsonPath', 'EvidenceCsvPath'
    )) {
        $property = $Audit.PSObject.Properties[$name]
        if ($property) {
            $lines.Add(('{0}: {1}' -f $name, $property.Value))
        }
    }
    if ($Audit.PSObject.Properties['PreviewPlan']) { $lines.Add(''); $lines.Add('Preview plan (no changes applied):'); $lines.Add(($Audit.PreviewPlan | ConvertTo-Json -Depth 5)) }
    return ($lines -join [Environment]::NewLine)
}

function Get-ResultValue {
    param(
        [Parameter(Mandatory = $true)]$Result,
        [Parameter(Mandatory = $true)][string]$Name,
        $Default = $null
    )

    $property = $Result.PSObject.Properties[$Name]
    if ($property -and $null -ne $property.Value) { return $property.Value }
    return $Default
}

function ConvertTo-ResultBoolean {
    param($Value)

    if ($null -eq $Value) { return $false }
    if ($Value -is [bool]) { return $Value }
    return ([string]$Value).Trim().ToLowerInvariant() -in @('true', '1', 'yes')
}

function Test-ResultHasText {
    param($Value)
    return -not [string]::IsNullOrWhiteSpace([string]$Value)
}

function New-CliStatus {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$State,
        [Parameter(Mandatory = $true)][string]$Detail,
        [Parameter(Mandatory = $true)][ConsoleColor]$Color
    )

    return [pscustomobject]([ordered]@{
        Name = $Name
        State = $State
        Detail = $Detail
        Color = $Color
    })
}

function Write-CliText {
    param(
        [string]$Text = '',
        [ConsoleColor]$Color = [ConsoleColor]::Gray
    )

    if ([string]::IsNullOrEmpty($Text)) {
        Write-Host ''
    }
    else {
        Write-Host $Text -ForegroundColor $Color
    }
}

function Write-CliSection {
    param([Parameter(Mandatory = $true)][string]$Title)

    Write-CliText ''
    Write-CliText $Title -Color Cyan
    Write-CliText ('-' * $Title.Length) -Color DarkCyan
}

function Get-ThinClientCliStatuses {
    param([Parameter(Mandatory = $true)]$Result)

    $statuses = New-Object System.Collections.Generic.List[object]

    $preferredAccessMode = [string](Get-ResultValue -Result $Result -Name 'PreferredAccessMode' -Default 'Web')
    $browserInstalled = ConvertTo-ResultBoolean (Get-ResultValue -Result $Result -Name 'BrowserInstalled' -Default $false)

    $citrixInstalled = ConvertTo-ResultBoolean (Get-ResultValue -Result $Result -Name 'CitrixWorkspaceInstalled' -Default $false)
    $citrixWebConfigured = ConvertTo-ResultBoolean (Get-ResultValue -Result $Result -Name 'CitrixWebUrlConfigured' -Default $false)
    $citrixPath = [string](Get-ResultValue -Result $Result -Name 'CitrixWorkspacePath' -Default '')
    if ($preferredAccessMode -eq 'Web' -and $browserInstalled -and $citrixWebConfigured) {
        $statuses.Add((New-CliStatus -Name 'Citrix Web' -State 'READY' -Detail 'Citrix web portal and browser are configured.' -Color Green))
    }
    elseif ($preferredAccessMode -eq 'Web' -and $citrixWebConfigured) {
        $statuses.Add((New-CliStatus -Name 'Citrix Web' -State 'PARTIAL' -Detail 'Citrix web portal is configured, but no supported browser was detected.' -Color Yellow))
    }
    elseif ($citrixInstalled) {
        $statuses.Add((New-CliStatus -Name 'Citrix' -State 'READY' -Detail ("Workspace found: {0}" -f $citrixPath) -Color Green))
    }
    else {
        $statuses.Add((New-CliStatus -Name 'Citrix' -State 'NOT READY' -Detail 'Citrix web URL and Citrix Workspace were not found.' -Color Red))
    }

    $avdInstalled = ConvertTo-ResultBoolean (Get-ResultValue -Result $Result -Name 'AvdClientInstalled' -Default $false)
    $avdWebConfigured = ConvertTo-ResultBoolean (Get-ResultValue -Result $Result -Name 'AvdWebUrlConfigured' -Default $false)
    $avdPath = [string](Get-ResultValue -Result $Result -Name 'AvdClassicClientPath' -Default '')
    $windowsAppPackage = Get-ResultValue -Result $Result -Name 'WindowsAppPackage' -Default ''
    if ($preferredAccessMode -eq 'Web' -and $browserInstalled -and $avdWebConfigured) {
        $statuses.Add((New-CliStatus -Name 'AVD Web' -State 'READY' -Detail 'AVD web portal and browser are configured.' -Color Green))
    }
    elseif ($preferredAccessMode -eq 'Web' -and $avdWebConfigured) {
        $statuses.Add((New-CliStatus -Name 'AVD Web' -State 'PARTIAL' -Detail 'AVD web portal is configured, but no supported browser was detected.' -Color Yellow))
    }
    elseif ($avdInstalled) {
        $avdDetail = 'AVD client is installed.'
        if (Test-ResultHasText $avdPath) { $avdDetail = "AVD client found: $avdPath" }
        $statuses.Add((New-CliStatus -Name 'AVD' -State 'READY' -Detail $avdDetail -Color Green))
    }
    elseif (Test-ResultHasText $windowsAppPackage) {
        $statuses.Add((New-CliStatus -Name 'AVD' -State 'READY' -Detail 'Windows App package was found.' -Color Green))
    }
    else {
        $statuses.Add((New-CliStatus -Name 'AVD' -State 'NOT READY' -Detail 'AVD web URL, AVD client, or Windows App was not found.' -Color Red))
    }

    $browserPath = [string](Get-ResultValue -Result $Result -Name 'BrowserPath' -Default '')
    $webUrlConfigured = ConvertTo-ResultBoolean (Get-ResultValue -Result $Result -Name 'WebOnlyUrlConfigured' -Default $false)
    if ($browserInstalled -and $webUrlConfigured) {
        $statuses.Add((New-CliStatus -Name 'WebOnly' -State 'READY' -Detail ("Browser and WebOnlyUrl are configured: {0}" -f $browserPath) -Color Green))
    }
    elseif ($browserInstalled) {
        $statuses.Add((New-CliStatus -Name 'WebOnly' -State 'PARTIAL' -Detail 'Browser found, but WebOnlyUrl is not configured.' -Color Yellow))
    }
    else {
        $statuses.Add((New-CliStatus -Name 'WebOnly' -State 'NOT READY' -Detail 'No supported browser was found.' -Color Red))
    }

    $assignedAccess = ConvertTo-ResultBoolean (Get-ResultValue -Result $Result -Name 'AssignedAccessCmdletAvailable' -Default $false)
    $shellLauncher = ConvertTo-ResultBoolean (Get-ResultValue -Result $Result -Name 'ShellLauncherClassAvailable' -Default $false)
    if ($assignedAccess -and $shellLauncher) {
        $statuses.Add((New-CliStatus -Name 'Kiosk / Lockdown' -State 'READY' -Detail 'Assigned Access and Shell Launcher are available.' -Color Green))
    }
    elseif ($assignedAccess -or $shellLauncher) {
        $detail = 'Partial kiosk capability.'
        if ($assignedAccess) { $detail = 'Assigned Access is available; Shell Launcher is not available.' }
        if ($shellLauncher) { $detail = 'Shell Launcher is available; Assigned Access cmdlet is not available.' }
        $statuses.Add((New-CliStatus -Name 'Kiosk / Lockdown' -State 'PARTIAL' -Detail $detail -Color Yellow))
    }
    else {
        $statuses.Add((New-CliStatus -Name 'Kiosk / Lockdown' -State 'LIMITED' -Detail 'Assigned Access and Shell Launcher were not detected.' -Color Yellow))
    }

    $isAdmin = ConvertTo-ResultBoolean (Get-ResultValue -Result $Result -Name 'IsAdministrator' -Default $false)
    $allowApply = ConvertTo-ResultBoolean (Get-ResultValue -Result $Result -Name 'AllowApply' -Default $false)
    if (-not $allowApply) {
        $statuses.Add((New-CliStatus -Name 'Apply' -State 'BLOCKED' -Detail 'AllowApply is false in the configuration.' -Color Yellow))
    }
    elseif (-not $isAdmin) {
        $statuses.Add((New-CliStatus -Name 'Apply' -State 'NEEDS ADMIN' -Detail 'Run PowerShell as administrator to apply thin-client mode.' -Color Yellow))
    }
    else {
        $statuses.Add((New-CliStatus -Name 'Apply' -State 'AVAILABLE' -Detail 'Requires the confirmation phrase before changing the PC.' -Color Green))
    }

    $allowRestore = ConvertTo-ResultBoolean (Get-ResultValue -Result $Result -Name 'AllowRestore' -Default $false)
    if (-not $allowRestore) {
        $statuses.Add((New-CliStatus -Name 'Restore' -State 'BLOCKED' -Detail 'AllowRestore is false in the configuration.' -Color Yellow))
    }
    elseif (-not $isAdmin) {
        $statuses.Add((New-CliStatus -Name 'Restore' -State 'NEEDS ADMIN' -Detail 'Run PowerShell as administrator to restore Windows shell settings.' -Color Yellow))
    }
    else {
        $statuses.Add((New-CliStatus -Name 'Restore' -State 'AVAILABLE' -Detail 'Uses the latest rollback file unless -RollbackPath is provided.' -Color Green))
    }

    return $statuses
}

function Get-ThinClientFinalMessage {
    param(
        [Parameter(Mandatory = $true)]$Result,
        [Parameter(Mandatory = $true)]$Statuses
    )

    $status = [string](Get-ResultValue -Result $Result -Name 'Status' -Default '')
    if ($status -eq 'Applied') {
        return (New-CliStatus -Name 'FINAL RESULT' -State 'APPLIED' -Detail 'Thin-client launcher and configured endpoint restrictions were applied.' -Color Green)
    }
    if ($status -eq 'Restored') {
        return (New-CliStatus -Name 'FINAL RESULT' -State 'RESTORED' -Detail 'Windows shell settings were restored from rollback evidence.' -Color Green)
    }
    if ($status -eq 'Launched') {
        return (New-CliStatus -Name 'FINAL RESULT' -State 'LAUNCHED' -Detail 'Temporary thin-client launcher was opened without installing or locking down Windows.' -Color Green)
    }

    $profile = [string](Get-ResultValue -Result $Result -Name 'EffectiveProfile' -Default 'Auto')
    $workspaceNames = @('Citrix Web', 'AVD Web', 'Citrix', 'AVD', 'WebOnly')
    $readyNames = @($Statuses | Where-Object { $_.Name -in $workspaceNames -and $_.State -eq 'READY' } | ForEach-Object { $_.Name })
    $partialNames = @($Statuses | Where-Object { $_.Name -in $workspaceNames -and $_.State -eq 'PARTIAL' } | ForEach-Object { $_.Name })

    if ($readyNames.Count -gt 0) {
        $detail = ('Ready workspace path: {0}. Effective profile: {1}.' -f ($readyNames -join ', '), $profile)
        if ($partialNames.Count -gt 0) { $detail += (' Partial: {0}.' -f ($partialNames -join ', ')) }
        return (New-CliStatus -Name 'FINAL RESULT' -State 'PREVIEW OK' -Detail $detail -Color Green)
    }

    if ($partialNames.Count -gt 0) {
        return (New-CliStatus -Name 'FINAL RESULT' -State 'PARTIAL' -Detail ('Only partial workspace readiness was detected: {0}.' -f ($partialNames -join ', ')) -Color Yellow)
    }

    return (New-CliStatus -Name 'FINAL RESULT' -State 'NOT READY' -Detail 'Install/configure Citrix Workspace, AVD/Windows App, or a WebOnlyUrl before applying thin-client mode.' -Color Red)
}

function Write-ThinClientCliSummary {
    param([Parameter(Mandatory = $true)]$Result)

    $requestedAction = [string](Get-ResultValue -Result $Result -Name 'RequestedAction' -Default $Action)
    $requestedProfile = [string](Get-ResultValue -Result $Result -Name 'RequestedProfile' -Default $Profile)
    $resultStatus = [string](Get-ResultValue -Result $Result -Name 'Status' -Default '')
    $effectiveProfile = [string](Get-ResultValue -Result $Result -Name 'EffectiveProfile' -Default '')
    if ([string]::IsNullOrWhiteSpace($effectiveProfile)) { $effectiveProfile = 'n/a' }

    Write-CliText ''
    Write-CliText 'Smart ThinClient Shell' -Color Cyan
    Write-CliText '======================' -Color DarkCyan
    Write-CliText ('Action: {0}    Status: {1}' -f $requestedAction, $resultStatus)
    Write-CliText ('Computer: {0}    User: {1}    Admin: {2}' -f (Get-ResultValue -Result $Result -Name 'ComputerName' -Default $env:COMPUTERNAME), (Get-ResultValue -Result $Result -Name 'UserName' -Default [Environment]::UserName), (Get-ResultValue -Result $Result -Name 'IsAdministrator' -Default 'n/a'))
    Write-CliText ('Profile: requested={0}    effective={1}' -f $requestedProfile, $effectiveProfile)

    $statuses = @()
    if ($resultStatus -in @('Applied', 'Restored', 'Launched')) {
        Write-CliSection -Title 'Operation'
        foreach ($name in @('LauncherPath', 'LauncherCommand', 'RollbackPath')) {
            $value = Get-ResultValue -Result $Result -Name $name -Default ''
            if (Test-ResultHasText $value) {
                Write-CliText ('{0}: {1}' -f $name, $value)
            }
        }
    }
    else {
        $statuses = @(Get-ThinClientCliStatuses -Result $Result)
        Write-CliSection -Title 'Readiness'
        foreach ($item in $statuses) {
            Write-Host ('{0,-17} ' -f $item.Name) -NoNewline -ForegroundColor White
            Write-Host ('{0,-11} ' -f $item.State) -NoNewline -ForegroundColor $item.Color
            Write-Host $item.Detail -ForegroundColor Gray
        }
    }

    $final = Get-ThinClientFinalMessage -Result $Result -Statuses $statuses
    Write-CliSection -Title 'Final message'
    Write-Host ('{0}: ' -f $final.Name) -NoNewline -ForegroundColor White
    Write-Host ('{0} - ' -f $final.State) -NoNewline -ForegroundColor $final.Color
    Write-Host $final.Detail -ForegroundColor Gray

    Write-CliSection -Title 'Evidence'
    Write-CliText ('JSON: {0}' -f (Get-ResultValue -Result $Result -Name 'EvidenceJsonPath' -Default 'n/a')) -Color DarkGray
    Write-CliText ('CSV : {0}' -f (Get-ResultValue -Result $Result -Name 'EvidenceCsvPath' -Default 'n/a')) -Color DarkGray
    Write-CliText ('Log : {0}' -f (Get-ResultValue -Result $Result -Name 'LogPath' -Default 'n/a')) -Color DarkGray
}

function Get-LauncherCommandLine {
    param([string]$Root=$script:OutputRoot,[switch]$Machine)
    $launcherPath = Join-Path $Root 'Launcher\SmartThinClient-LaunchWorkspace.ps1'
    if($Machine){return ('"{0}" -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "{1}"' -f (Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'),$launcherPath)}
    return ('powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "{0}"' -f $launcherPath)
}

function ConvertTo-QuotedPowerShellString {
    param([string]$Value)
    if ($null -eq $Value) { $Value = '' }
    return "'{0}'" -f ($Value -replace "'", "''")
}

function New-ThinClientUiResources {
    Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
    [xml]$resources = @'
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
  <Style TargetType="Button">
    <Setter Property="Background" Value="White"/><Setter Property="Foreground" Value="#24364B"/>
    <Setter Property="BorderBrush" Value="#CBD5E1"/><Setter Property="BorderThickness" Value="1"/>
    <Setter Property="FontFamily" Value="Segoe UI"/><Setter Property="FontSize" Value="12"/><Setter Property="FontWeight" Value="SemiBold"/>
    <Setter Property="Padding" Value="14,0"/><Setter Property="MinHeight" Value="34"/><Setter Property="Cursor" Value="Hand"/>
    <Setter Property="HorizontalContentAlignment" Value="Center"/><Setter Property="VerticalContentAlignment" Value="Center"/>
    <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="Button">
      <Border x:Name="Frame" CornerRadius="6" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="{TemplateBinding BorderThickness}" Padding="{TemplateBinding Padding}">
        <ContentPresenter HorizontalAlignment="{TemplateBinding HorizontalContentAlignment}" VerticalAlignment="{TemplateBinding VerticalContentAlignment}" RecognizesAccessKey="True"/>
      </Border>
      <ControlTemplate.Triggers>
        <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Frame" Property="Background" Value="#E6F4FF"/><Setter TargetName="Frame" Property="BorderBrush" Value="#78B5E8"/><Setter Property="Foreground" Value="#005A9E"/></Trigger>
        <Trigger Property="IsPressed" Value="True"><Setter TargetName="Frame" Property="Background" Value="#CDE7FC"/></Trigger>
        <Trigger Property="IsKeyboardFocused" Value="True"><Setter TargetName="Frame" Property="BorderBrush" Value="#0078D4"/><Setter TargetName="Frame" Property="BorderThickness" Value="2"/></Trigger>
        <Trigger Property="IsEnabled" Value="False"><Setter TargetName="Frame" Property="Background" Value="#F1F4F8"/><Setter TargetName="Frame" Property="BorderBrush" Value="#E2E8F0"/><Setter Property="Foreground" Value="#8491A3"/><Setter Property="Cursor" Value="Arrow"/></Trigger>
      </ControlTemplate.Triggers>
    </ControlTemplate></Setter.Value></Setter>
  </Style>
  <Style x:Key="AccentButton" TargetType="Button" BasedOn="{StaticResource {x:Type Button}}"><Setter Property="Background" Value="#0078D4"/><Setter Property="Foreground" Value="White"/><Setter Property="BorderBrush" Value="#0078D4"/></Style>
  <Style TargetType="ComboBoxItem">
    <Setter Property="Padding" Value="10,7"/><Setter Property="Foreground" Value="#24364B"/>
    <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ComboBoxItem">
      <Border x:Name="Row" Background="Transparent" CornerRadius="4" Padding="{TemplateBinding Padding}"><ContentPresenter/></Border>
      <ControlTemplate.Triggers><Trigger Property="IsHighlighted" Value="True"><Setter TargetName="Row" Property="Background" Value="#EFF6FF"/></Trigger><Trigger Property="IsSelected" Value="True"><Setter TargetName="Row" Property="Background" Value="#DCEEFF"/><Setter Property="Foreground" Value="#005A9E"/><Setter Property="FontWeight" Value="SemiBold"/></Trigger></ControlTemplate.Triggers>
    </ControlTemplate></Setter.Value></Setter>
  </Style>
  <Style TargetType="ComboBox">
    <Setter Property="Background" Value="White"/><Setter Property="BorderBrush" Value="#CBD5E1"/><Setter Property="Foreground" Value="#24364B"/><Setter Property="FontSize" Value="12"/><Setter Property="MinHeight" Value="34"/>
    <Setter Property="ScrollViewer.CanContentScroll" Value="True"/>
    <Setter Property="Template"><Setter.Value><ControlTemplate TargetType="ComboBox">
      <Grid>
        <Border x:Name="Frame" Background="{TemplateBinding Background}" BorderBrush="{TemplateBinding BorderBrush}" BorderThickness="1" CornerRadius="6"/>
        <ToggleButton Focusable="False" ClickMode="Press" IsChecked="{Binding IsDropDownOpen, RelativeSource={RelativeSource TemplatedParent}, Mode=TwoWay}">
          <ToggleButton.Template><ControlTemplate TargetType="ToggleButton"><Border Background="Transparent" CornerRadius="6"/></ControlTemplate></ToggleButton.Template>
        </ToggleButton>
        <ContentPresenter Margin="12,0,32,0" VerticalAlignment="Center" HorizontalAlignment="Left" IsHitTestVisible="False" Content="{TemplateBinding SelectionBoxItem}" ContentTemplate="{TemplateBinding SelectionBoxItemTemplate}"/>
        <Path Data="M 0 0 L 4 4 L 8 0" Stroke="#5F6B7A" StrokeThickness="1.5" HorizontalAlignment="Right" VerticalAlignment="Center" Margin="0,0,12,0" IsHitTestVisible="False"/>
        <Popup x:Name="PART_Popup" Placement="Bottom" IsOpen="{TemplateBinding IsDropDownOpen}" AllowsTransparency="True" Focusable="False" PopupAnimation="Fade">
          <Border MinWidth="{TemplateBinding ActualWidth}" Background="White" BorderBrush="#CBD5E1" BorderThickness="1" CornerRadius="6" Padding="4" Margin="0,4,0,0"><ScrollViewer MaxHeight="260" CanContentScroll="True"><ItemsPresenter KeyboardNavigation.DirectionalNavigation="Contained"/></ScrollViewer></Border>
        </Popup>
      </Grid>
      <ControlTemplate.Triggers><Trigger Property="IsMouseOver" Value="True"><Setter TargetName="Frame" Property="BorderBrush" Value="#78B5E8"/></Trigger><Trigger Property="IsKeyboardFocusWithin" Value="True"><Setter TargetName="Frame" Property="BorderBrush" Value="#0078D4"/></Trigger><Trigger Property="IsEnabled" Value="False"><Setter Property="Opacity" Value="0.55"/></Trigger></ControlTemplate.Triggers>
    </ControlTemplate></Setter.Value></Setter>
  </Style>
</ResourceDictionary>
'@
    return ,([Windows.Markup.XamlReader]::Load([Xml.XmlNodeReader]::new($resources)))
}

function Get-ThinClientProviderVisibility {
    param([string]$EffectiveProfile, [bool]$ShowProviderButtons = $true)
    return @{ Citrix = ($ShowProviderButtons -and $EffectiveProfile -in @('Citrix','Hybrid')); AVD = ($ShowProviderButtons -and $EffectiveProfile -in @('AVD','Hybrid')); WebOnly = ($ShowProviderButtons -and $EffectiveProfile -in @('WebOnly','Hybrid')) }
}

function ConvertTo-ThinClientSettingsButtons {
    param($Value)
    $buttons = [ordered]@{}
    if ($null -eq $Value) { return ,$buttons }
    if ($Value -is [Collections.IDictionary]) { foreach ($key in $Value.Keys) { $buttons[[string]$key] = $Value[$key] } }
    elseif ($Value -is [pscustomobject]) { foreach ($property in $Value.PSObject.Properties) { $buttons[$property.Name] = $property.Value } }
    else { throw 'WebShellSettingsButtons must be a JSON object of booleans.' }
    foreach ($key in $buttons.Keys) { if ($buttons[$key] -isnot [bool]) { throw "WebShellSettingsButtons.$key must be a JSON boolean." } }
    return ,$buttons
}

function Get-ThinClientVisibleSettingsPages {
    param([string[]]$SettingsPages, $SettingsButtons)
    $buttons = ConvertTo-ThinClientSettingsButtons -Value $SettingsButtons
    $names = @{'ms-settings:network-status'='Network';'ms-settings:display'='Display';'ms-settings:sound'='Sound';'ms-settings:dateandtime'='DateAndTime';'ms-settings:printers'='Printers';'ms-settings:bluetooth'='Bluetooth'}
    foreach ($page in @($SettingsPages | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
        $uri = $null
        if (-not [Uri]::TryCreate($page,[UriKind]::Absolute,[ref]$uri) -or $uri.Scheme -ne 'ms-settings') { throw 'Only configured ms-settings: pages can be opened from Settings.' }
        $key = if ($names.ContainsKey($page)) { $names[$page] } else { $page }
        if (-not $buttons.Contains($key) -or $buttons[$key]) { $page }
    }
}

function Confirm-ThinClientPowerAction {
    param($Owner,[ValidateSet('Restart','Shutdown')][string]$Action)
    Add-Type -AssemblyName PresentationFramework
    [xml]$xaml=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Confirm device action" Width="470" SizeToContent="Height" ResizeMode="NoResize" WindowStartupLocation="CenterOwner" Background="#F5F8FB" FontFamily="Segoe UI">
 <StackPanel Margin="18">
  <Border Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="18"><StackPanel>
   <Border Background="#FFF4D6" CornerRadius="6" Padding="8,3" HorizontalAlignment="Left"><TextBlock Text="Device action" Foreground="#8A5700" FontWeight="SemiBold"/></Border>
   <TextBlock Name="Heading" FontSize="23" FontWeight="SemiBold" Foreground="#1F2937" Margin="0,12,0,8"/>
   <TextBlock Name="Description" TextWrapping="Wrap" Foreground="#5F6B7A"/>
  </StackPanel></Border>
  <Border Background="#E6F4FF" BorderBrush="#B9DDF7" BorderThickness="1" CornerRadius="8" Padding="12" Margin="0,12,0,0"><TextBlock Text="Save your open work before continuing. Your Windows session and running applications will close." TextWrapping="Wrap" Foreground="#24364B"/></Border>
  <StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,16,0,0"><Button Name="CancelButton" Content="Cancel" IsCancel="True" IsDefault="True" MinWidth="100" Height="36" Margin="0,0,10,0"/><Button Name="ConfirmButton" MinWidth="110" Height="36"/></StackPanel>
 </StackPanel>
</Window>
'@
    $dialog=[Windows.Markup.XamlReader]::Load([Xml.XmlNodeReader]::new($xaml))
    [void]$dialog.Resources.MergedDictionaries.Add((New-ThinClientUiResources))
    if($Owner){$dialog.Owner=$Owner;$dialog.Icon=$Owner.Icon}
    $restart=$Action -eq 'Restart'
    $dialog.Title=if($restart){'Confirm restart'}else{'Confirm shutdown'}
    $dialog.FindName('Heading').Text=if($restart){'Restart this device?'}else{'Shut down this device?'}
    $dialog.FindName('Description').Text=if($restart){'Windows will restart. Sign in again to return to your workspace.'}else{'Windows will shut down. You will need to turn the device on again to return to your workspace.'}
    $dialog.FindName('CancelButton').Add_Click({$dialog.DialogResult=$false})
    $confirm=$dialog.FindName('ConfirmButton');$confirm.Content=if($restart){'Restart'}else{'Shut down'};$confirm.Style=$dialog.FindResource('AccentButton');$confirm.Add_Click({$dialog.DialogResult=$true})
    return $dialog.ShowDialog() -eq $true
}

function Invoke-ThinClientPowerAction {
    param($Owner,[ValidateSet('Restart','Shutdown')][string]$Action)
    if(-not(Confirm-ThinClientPowerAction -Owner $Owner -Action $Action)){return $false}
    $prior=Get-Variable -Name ShellExitRequested -Scope Script -ErrorAction SilentlyContinue
    $previousExit=if($prior){[bool]$prior.Value}else{$false}
    $script:ShellExitRequested=$true
    try {
        $arguments=if($Action -eq 'Restart'){'/r /t 0'}else{'/s /t 0'}
        $process=Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\shutdown.exe') -ArgumentList $arguments -WindowStyle Hidden -Wait -PassThru -ErrorAction Stop
        if($process.ExitCode -ne 0){throw ('Windows returned exit '+$process.ExitCode+'.')}
        return $true
    }catch{$script:ShellExitRequested=$previousExit;throw}
}

function New-ThinClientSettingsRequest {
    param([string]$Page)
    if($Page -notin @('ms-settings:network-status','ms-settings:display','ms-settings:sound','ms-settings:dateandtime','ms-settings:bluetooth','ms-settings:printers')){throw 'This Windows settings page is not allowed.'}
    if(-not('SmartThinClient.SettingsLaunchRequest' -as [type])){
        Add-Type -TypeDefinition @'
using System;using System.Diagnostics;using System.Threading;using System.IO;
namespace SmartThinClient {
 public sealed class SettingsLaunchRequest {
  public volatile bool Completed;
  public bool Requested;
  public string Error="";
  public static SettingsLaunchRequest Begin(string page) {
   if(Array.IndexOf(new[]{"ms-settings:network-status","ms-settings:display","ms-settings:sound","ms-settings:dateandtime","ms-settings:bluetooth","ms-settings:printers"},page)<0)throw new ArgumentException("Settings page is not allowed.");
   var request=new SettingsLaunchRequest();
   ThreadPool.QueueUserWorkItem(delegate(object unused){
    try {using(var process=Process.Start(new ProcessStartInfo(page){UseShellExecute=true,ErrorDialog=false})){request.Requested=true;}}
    catch(Exception ex){request.Error=ex.Message;try{string dir=Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),@"WCH\Smart365\Logs");Directory.CreateDirectory(dir);File.AppendAllText(Path.Combine(dir,"SmartThinClient-Settings.log"),DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")+" [ERROR] "+page+": "+ex.Message.Replace("\r"," ").Replace("\n"," ")+Environment.NewLine);}catch{}}
    finally {request.Completed=true;}
   });
   return request;
  }
 }
}
'@ -ErrorAction Stop
    }
    [SmartThinClient.SettingsLaunchRequest]::Begin($Page)
}

function Show-ThinClientLimitedSettings {
    param([string[]]$SettingsPages, $SettingsButtons, $Owner, [switch]$PreviewMode, [switch]$NativeMode, [string]$NativeSettingsHelper)
    $pages = @(Get-ThinClientVisibleSettingsPages -SettingsPages $SettingsPages -SettingsButtons $SettingsButtons)
    if($NativeMode){
        if([IO.Path]::GetFileName($NativeSettingsHelper) -ne 'SmartThinClient-DeviceSettings.ps1' -or -not(Test-Path -LiteralPath $NativeSettingsHelper -PathType Leaf)){throw 'The native device settings helper is unavailable. Reinstall the complete application files.'}
        . $NativeSettingsHelper
    }
    Add-Type -AssemblyName PresentationFramework
    $menu = [Windows.Window]::new(); $menu.Title = 'Smart ThinClient Settings'; $menu.Width = 360; $menu.SizeToContent = 'Height'; $menu.MaxHeight = 520; $menu.ResizeMode = 'NoResize'; $menu.WindowStartupLocation = 'CenterScreen'; $menu.Background = '#F5F8FB'
    if ($null -ne $Owner) { $menu.Owner = $Owner; $menu.WindowStartupLocation = 'CenterOwner'; if ($null -ne $Owner.Icon) { $menu.Icon = $Owner.Icon } }
    [void]$menu.Resources.MergedDictionaries.Add((New-ThinClientUiResources))
    $panel = [Windows.Controls.StackPanel]::new(); $panel.Margin = '18'
    $title = [Windows.Controls.TextBlock]::new(); $title.Text = 'Windows settings'; $title.FontSize = 22; $title.FontWeight = 'SemiBold'; $title.Foreground = '#1F2937'; $title.Margin = '0,0,0,8'; [void]$panel.Children.Add($title)
    $note = [Windows.Controls.TextBlock]::new(); $note.TextWrapping = 'Wrap'; $note.Foreground = '#5F6B7A'; $note.Margin = '0,0,0,14'
    $note.Text = if ($PreviewMode) { 'These are live Windows settings, including in Preview. Changes made there affect this device.' } else { 'Open a Windows settings page allowed by the shell configuration.' }; [void]$panel.Children.Add($note)
    if($NativeMode){$title.Text='Device settings';$note.Text='Built-in device controls work without Explorer. Changes are live, including in Preview. Available controls follow Windows permissions.'}
    $labels = @{ 'ms-settings:network-status'='Network'; 'ms-settings:display'='Display'; 'ms-settings:sound'='Sound'; 'ms-settings:dateandtime'='Date and time'; 'ms-settings:printers'='Printers'; 'ms-settings:bluetooth'='Bluetooth' }
    $list = [Windows.Controls.StackPanel]::new()
    if ($pages.Count -eq 0) { $empty = [Windows.Controls.TextBlock]::new(); $empty.Text = 'No settings pages are enabled by configuration.'; $empty.TextWrapping = 'Wrap'; $empty.Margin = '0,0,0,10'; [void]$list.Children.Add($empty) }
    $status = [Windows.Controls.TextBlock]::new();$status.TextWrapping='Wrap';$status.Foreground='#5F6B7A';$status.Margin='0,4,0,8'
    $pending=[pscustomobject]@{Request=$null;Started=[DateTime]::MinValue;Page=''}
    $timer=[Windows.Threading.DispatcherTimer]::new();$timer.Interval=[TimeSpan]::FromMilliseconds(200)
    $timer.Add_Tick({
        try {
            if($null -eq $pending.Request){$timer.Stop();return}
            if($pending.Request.Completed){
                $status.Text=if($pending.Request.Requested){'Open request sent for '+$labels[$pending.Page]+'.'}else{'Windows Settings could not open: '+$pending.Request.Error}
                $status.Foreground=if($pending.Request.Requested){'#167044'}else{'#B42318'}
                foreach($entry in $list.Children){if($entry -is [Windows.Controls.Button]){$entry.IsEnabled=$true}}
                $timer.Stop()
            }elseif(([DateTime]::Now-$pending.Started).TotalSeconds -ge 20){$status.Text='Windows Settings is not responding. You can close this dialog; the workspace remains available.';$status.Foreground='#B42318'}
        }catch{$status.Text='Windows Settings could not open: '+$_.Exception.Message;$status.Foreground='#B42318';$timer.Stop()}
    })
    foreach ($page in $pages) {
        $button = [Windows.Controls.Button]::new(); $button.Height = 36; $button.Margin = '0,0,0,8'; $button.HorizontalContentAlignment = 'Left'; $button.Content = if ($labels.ContainsKey($page)) { $labels[$page] } else { $page }; $button.Tag = $page
        $button.Add_Click({
            try {
                if($NativeMode){
                    $nativePages=@{'ms-settings:network-status'='Network';'ms-settings:display'='Display';'ms-settings:sound'='Sound';'ms-settings:dateandtime'='DateAndTime';'ms-settings:bluetooth'='Bluetooth';'ms-settings:printers'='Printers'}
                    Show-SmartThinClientDeviceSettings -Page $nativePages[[string]$this.Tag] -Owner $menu -PreviewMode:$PreviewMode
                    return
                }
                $pending.Page=[string]$this.Tag;$pending.Started=[DateTime]::Now;$pending.Request=New-ThinClientSettingsRequest -Page $pending.Page
                $status.Text='Opening '+$labels[$pending.Page]+'...';$status.Foreground='#005A9E'
                foreach($entry in $list.Children){if($entry -is [Windows.Controls.Button]){$entry.IsEnabled=$false}}
                $timer.Start()
            }catch{$status.Text='Windows Settings could not open: '+$_.Exception.Message;$status.Foreground='#B42318'}
        }); [void]$list.Children.Add($button)
    }
    $scroll = [Windows.Controls.ScrollViewer]::new(); $scroll.MaxHeight = 280; $scroll.VerticalScrollBarVisibility = 'Auto'; $scroll.Content = $list; [void]$panel.Children.Add($scroll)
    [void]$panel.Children.Add($status)
    $close = [Windows.Controls.Button]::new(); $close.Content = 'Close'; $close.Height = 34; $close.Margin = '0,6,0,0'; $close.Add_Click({ $menu.Close() }); [void]$panel.Children.Add($close)
    $menu.Content = $panel; try {[void]$menu.ShowDialog()}finally{$timer.Stop()}
}

function Get-ThinClientClientDownload {
    param([Parameter(Mandatory = $true)][ValidateSet('CitrixWorkspace','WindowsApp')][string]$Client)
    if ($Client -eq 'CitrixWorkspace') {
        return [pscustomobject]@{Label='Download Citrix Workspace';Url='https://www.citrix.com/downloads/workspace-app/windows/'}
    }
    [pscustomobject]@{Label='Download Windows App';Url='https://apps.microsoft.com/detail/9n1f85v9t8bn'}
}

function Get-ThinClientAuditSummary {
    param([Parameter(Mandatory = $true)]$Audit)
    if ($Audit -isnot [Collections.IDictionary]) { $Audit = Convert-ObjectToHashtable -InputObject $Audit }
    $checks = [Collections.Generic.List[object]]::new()
    $add = { param($label,$state,$detail,[string]$downloadClient='') [void]$checks.Add([pscustomobject]@{Label=$label;State=$state;Detail=$detail;DownloadClient=$downloadClient}) }
    $profile = [string](Get-ConfigValue -Config $Audit -Name EffectiveProfile -Default 'WebOnly')
    $os = [string](Get-ConfigValue -Config $Audit -Name OSName -Default '')
    $shellMode = [string](Get-ConfigValue -Config $Audit -Name ShellMode -Default 'Launcher')
    $unsupported = ($shellMode -eq 'AssignedAccess' -and -not (Get-ConfigValue -Config $Audit -Name AssignedAccessCmdletAvailable -Default $false)) -or ($shellMode -eq 'ShellLauncher' -and -not (Get-ConfigValue -Config $Audit -Name ShellLauncherClassAvailable -Default $false)) -or ($shellMode -eq 'CustomShell' -and -not (Get-ConfigValue -Config $Audit -Name CustomShellPolicyAvailable -Default $false))
    & $add 'Windows' $(if($unsupported){'Blocked'}elseif($os){'Success'}else{'Warning'}) $(if($unsupported){$shellMode+' unavailable'}elseif($os){[string](Get-ConfigValue -Config $Audit -Name EditionId -Default 'Detected')}else{'Not reported'})
    $native = -not [bool](Get-ConfigValue -Config $Audit -Name UseWebShell -Default $true) -and (Get-ConfigValue -Config $Audit -Name PreferredAccessMode -Default 'Web') -eq 'Native'
    $browser = [bool](Get-ConfigValue -Config $Audit -Name BrowserInstalled -Default $false)
    & $add 'Browser' $(if($browser){'Success'}elseif($native){'Info'}else{'Blocked'}) $(if($browser){'Detected'}elseif($native){'Native mode'}else{'Not detected'})
    $providers = if($profile -eq 'Hybrid'){@('Citrix','AVD')}else{@($profile)}
    foreach($provider in $providers) {
        $key = switch($provider){'Citrix'{'CitrixWebUrlConfigured'}'AVD'{'AvdWebUrlConfigured'}default{'WebOnlyUrlConfigured'}}
        $configured = [bool](Get-ConfigValue -Config $Audit -Name $key -Default $false)
        $fallback = $provider -ne 'WebOnly' -and [bool](Get-ConfigValue -Config $Audit -Name WebOnlyUrlConfigured -Default $false)
        if ($provider -eq 'WebOnly') {
            & $add 'WebOnly access' $(if($configured -and $browser){if($native){'Warning'}else{'Success'}}else{'Blocked'}) $(if($configured -and $browser){if($native){'Web fallback'}else{'URL configured'}}elseif($configured){'Browser missing'}else{'No destination'})
            continue
        }
        $windowsApp = [string](Get-ConfigValue -Config $Audit -Name WindowsAppPackage -Default '')
        $legacyAvd = [string](Get-ConfigValue -Config $Audit -Name AvdClassicClientPath -Default '')
        $client = if($provider -eq 'Citrix'){[bool](Get-ConfigValue -Config $Audit -Name CitrixWorkspaceInstalled -Default $false)}else{(-not [string]::IsNullOrWhiteSpace($windowsApp)) -or (-not [string]::IsNullOrWhiteSpace($legacyAvd))}
        $nativeReady = $native -and $client
        $webAvailable = $browser -and ($configured -or $fallback)
        $webState = if($nativeReady -and (-not $configured -or -not $browser)){'Info'}elseif($configured -and $browser){'Success'}elseif($fallback -and $browser){'Warning'}else{'Blocked'}
        $webDetail = if($nativeReady -and -not $configured){'Not configured (native mode)'}elseif($nativeReady -and -not $browser){'Browser missing (native mode)'}elseif($configured -and $browser){'URL configured'}elseif($fallback -and $browser){'WebOnly fallback'}elseif(($configured -or $fallback) -and -not $browser){'Browser missing'}else{'No destination'}
        & $add ($provider+' Web') $webState $webDetail
        $missingClientState = if(-not $native){'Info'}elseif($webAvailable){'Warning'}else{'Blocked'}
        if ($provider -eq 'Citrix') {
            $reported = $Audit.Contains('CitrixWorkspaceInstalled') -and $null -ne $Audit['CitrixWorkspaceInstalled']
            & $add 'Citrix Workspace' $(if($client){'Success'}else{$missingClientState}) $(if($client){'Detected'}elseif($reported){'Not detected'}else{'Not reported'}) $(if(-not $client){'CitrixWorkspace'}else{''})
        }
        else {
            $appReported = $Audit.Contains('WindowsAppPackage') -and $null -ne $Audit['WindowsAppPackage']
            $legacyReported = $Audit.Contains('AvdClassicClientPath') -and $null -ne $Audit['AvdClassicClientPath']
            & $add 'Windows App' $(if($windowsApp){'Success'}elseif($legacyAvd){'Info'}else{$missingClientState}) $(if($windowsApp){'Detected'}elseif($appReported){'Not detected'}else{'Not reported'}) $(if(-not $windowsApp){'WindowsApp'}else{''})
            & $add 'Remote Desktop (legacy)' $(if($legacyAvd){'Warning'}else{'Info'}) $(if($legacyAvd){'Detected (legacy)'}elseif($legacyReported){'Not detected'}else{'Not reported'})
        }
    }
    $identityStatus = [string](Get-ConfigValue -Config $Audit -Name TargetUserIdentityStatus -Default '')
    $identityKind = [string](Get-ConfigValue -Config $Audit -Name TargetUserAccountType -Default 'Unknown')
    $targetState = 'Warning'; $targetDetail = 'Not verified'
    if ($identityStatus -eq 'Detected') {
        $targetState = if($identityKind -eq 'System'){'Warning'}else{'Success'}
        $targetDetail = switch($identityKind){'Domain'{'Domain account detected'}'Local'{'Local account detected'}'Entra'{'Entra account detected'}'MicrosoftAccount'{'Microsoft account detected'}'System'{'System identity'}default{'Account detected'}}
    }
    elseif ($identityStatus -eq 'Disabled') { $targetDetail = 'Local account disabled' }
    elseif (-not $identityStatus -and (Get-ConfigValue -Config $Audit -Name TargetUserExists -Default $false)) { $targetState = 'Success'; $targetDetail = 'Account detected' }
    & $add $(if($shellMode -eq 'CustomShell'){'Current account'}else{'Target account'}) $targetState $targetDetail
    if ($shellMode -eq 'CustomShell') { & $add 'Shell scope' 'Info' 'All users (configured)' }
    $state = if(@($checks|Where-Object State -eq Blocked).Count){'Blocked'}elseif(@($checks|Where-Object State -eq Warning).Count){'Warning'}else{'Success'}
    $gates = 'Apply ' + $(if(Get-ConfigValue -Config $Audit -Name AllowApply -Default $false){'enabled'}else{'disabled'}) + ' | Restore ' + $(if(Get-ConfigValue -Config $Audit -Name AllowRestore -Default $false){'enabled'}else{'disabled'})
    [pscustomobject]@{State=$state;Status=$(switch($state){'Blocked'{'Blocked'}'Warning'{'Review required'}default{'Checks passed'}});Checks=@($checks.ToArray());Note=($gates + '. Native clients are optional in Web mode. Local signals only; sign-in and remote sessions are not tested by Audit.')}
}

function Set-ThinClientAuditSummary {
    param($Audit,$Panel,$Title,$Note,$Container)
    $summary = Get-ThinClientAuditSummary -Audit $Audit
    $colors = @{Success=@('#E8F5EE','#167044');Warning=@('#FFF4D6','#8A5700');Blocked=@('#FDECEC','#B42318');Info=@('#EDF2F7','#526579')}
    $Title.Text = 'Audit | '+$summary.Status; $Title.Foreground = $colors[$summary.State][1]; $Note.Text = $summary.Note
    $Panel.Children.Clear()
    foreach($check in $summary.Checks) {
        $card=[Windows.Controls.Border]::new();$card.Background=$colors[$check.State][0];$card.CornerRadius='6';$card.Padding='10,7';$card.Margin='0,0,6,6';$card.MinWidth=140
        $text=[Windows.Controls.TextBlock]::new();$text.Text=$check.Label+' | '+$check.Detail;$text.Foreground=$colors[$check.State][1];$text.FontWeight='SemiBold'
        $content=[Windows.Controls.StackPanel]::new();[void]$content.Children.Add($text)
        if ($check.DownloadClient) {
            $download=Get-ThinClientClientDownload -Client $check.DownloadClient
            $link=[Windows.Documents.Hyperlink]::new([Windows.Documents.Run]::new($download.Label));$link.Tag=$check.DownloadClient;$link.ToolTip=$download.Url
            $link.Add_Click({
                try { $target=Get-ThinClientClientDownload -Client ([string]$this.Tag);Start-Process -FilePath $target.Url -ErrorAction Stop }
                catch { [void][Windows.MessageBox]::Show(('Unable to open the official download page. '+$_.Exception.Message),'Smart ThinClient','OK','Warning') }
            })
            $linkText=[Windows.Controls.TextBlock]::new();$linkText.FontSize=11;$linkText.Margin='0,4,0,0';[void]$linkText.Inlines.Add($link);[void]$content.Children.Add($linkText)
        }
        $card.Child=$content;[void]$Panel.Children.Add($card)
    }
    $Container.Visibility='Visible'
}

function Get-ThinClientVersion {
    return '1.1.4'
}

function Get-ThinClientShellXaml {
    $layout = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Smart ThinClient Shell" WindowStartupLocation="CenterScreen"
        WindowState="Maximized" Background="#F5F8FB">
  <Grid>
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="40"/>
    </Grid.RowDefinitions>
    <Border Grid.Row="0" Background="#1F2937" Padding="12,8">
      <Grid Name="ToolbarLayout">
        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
        <WrapPanel Name="NavigationButtons" Grid.Column="0" HorizontalAlignment="Left" VerticalAlignment="Center">
          <Button Name="CitrixButton" Content="Citrix Web" Width="98" Height="34" Margin="0,2,8,2"/>
          <Button Name="AvdButton" Content="AVD Web" Width="90" Height="34" Margin="0,2,8,2"/>
          <Button Name="WebOnlyButton" Content="Web" Width="72" Height="34" Margin="0,2,8,2"/>
          <Button Name="BackButton" Content="Back" Width="78" Height="34" Margin="0,2,8,2"/>
          <Button Name="RefreshButton" Content="Refresh" Width="82" Height="34" Margin="0,2,8,2"/>
          <Button Name="ExternalButton" Content="Open browser" Width="112" Height="34" Margin="0,2,8,2"/>
        </WrapPanel>
        <TextBlock Name="TitleText" Grid.Column="1" Text="Smart ThinClient" Foreground="White" FontSize="20" FontWeight="SemiBold" VerticalAlignment="Center" HorizontalAlignment="Center" TextAlignment="Center" TextTrimming="CharacterEllipsis" MaxWidth="280" Margin="16,0"/>
        <WrapPanel Name="SessionButtons" Grid.Column="2" HorizontalAlignment="Right" VerticalAlignment="Center">
          <Button Name="SettingsButton" Content="Settings" Width="86" Height="34" Margin="0,2,8,2"/>
          <Button Name="AdminButton" Content="Admin" ToolTip="Open the administrator control panel" Visibility="Collapsed" Width="76" Height="34" Margin="0,2,8,2"/>
          <Button Name="SignOutButton" Content="Sign out" Width="86" Height="34" Margin="0,2,8,2"/>
          <Button Name="RestartButton" Content="Restart" Width="82" Height="34" Margin="0,2,8,2"/>
          <Button Name="ShutdownButton" Content="Shutdown" Width="92" Height="34" Margin="0,2,0,2"/>
        </WrapPanel>
      </Grid>
    </Border>
    <Border Grid.Row="1" Background="White" BorderBrush="#DDE7F0" BorderThickness="1,0,1,0">
      <ContentControl Name="WorkspaceBrowser"/>
    </Border>
    <Border Grid.Row="2" Background="#F8FBFE" BorderBrush="#DDE7F0" BorderThickness="1,1,1,0" Padding="10,0">
      <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <TextBlock Name="StatusText" Text="Ready" Foreground="#5F6B7A" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
        <TextBlock Name="VersionText" Grid.Column="1" Text="Smart ThinClient v__APP_VERSION__" Foreground="#64748B" FontSize="11" VerticalAlignment="Center" Margin="16,0,0,0" ToolTip="Smart ThinClient version __APP_VERSION__"/>
      </Grid>
    </Border>
  </Grid>
</Window>
'@
    return $layout.Replace('__APP_VERSION__',(Get-ThinClientVersion))
}

function New-LauncherScript {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$EffectiveProfile,
        [Parameter(Mandatory = $true)]$Audit
    )

    Assert-ConfigValues -Config $Config
    $launcherRoot = Join-Path $script:OutputRoot 'Launcher'
    New-Item -Path $launcherRoot -ItemType Directory -Force | Out-Null
    $launcherPath = Join-Path $launcherRoot 'SmartThinClient-LaunchWorkspace.ps1'

    $citrixPath = [string]$Audit.CitrixWorkspacePath
    $avdPath = [string]$Audit.AvdClassicClientPath
    $browserPath = [string]$Audit.BrowserPath
    $webUrls = Get-WorkspaceWebUrls -Config $Config -EffectiveProfile $EffectiveProfile
    $citrixWebUrl = $webUrls.Citrix
    $avdWebUrl = $webUrls.AVD
    $webUrl = $webUrls.WebOnly
    $windowsAppUri = [string](Get-ConfigValue -Config $Config -Name 'WindowsAppLaunchUri' -Default '')
    $browserKiosk = [bool](Get-ConfigValue -Config $Config -Name 'BrowserKioskMode' -Default $false)
    $useWebShell = [bool](Get-ConfigValue -Config $Config -Name 'UseWebShell' -Default $true)
    $preferredAccessMode = [string](Get-ConfigValue -Config $Config -Name 'PreferredAccessMode' -Default 'Web')
    $webShellTitle = [string](Get-ConfigValue -Config $Config -Name 'WebShellTitle' -Default 'Smart ThinClient')
    $showProviderButtons = [bool](Get-ConfigValue -Config $Config -Name 'WebShellShowProviderButtons' -Default $true)
    $allowExternalBrowser = [bool](Get-ConfigValue -Config $Config -Name 'WebShellAllowExternalBrowser' -Default $true)
    $allowPowerControls = [bool](Get-ConfigValue -Config $Config -Name 'WebShellAllowPowerControls' -Default $true)
    $allowLimitedSettings = [bool](Get-ConfigValue -Config $Config -Name 'WebShellAllowLimitedSettings' -Default $true)
    $primaryCustomShell = (Get-ConfigValue -Config $Config -Name ShellMode -Default Launcher) -eq 'CustomShell'
    $hybridDefault = [string](Get-ConfigValue -Config $Config -Name 'HybridDefaultProvider' -Default '')
    $webShellPreferredProvider = [string](Get-ConfigValue -Config $Config -Name 'WebShellPreferredProvider' -Default '')
    if ($hybridDefault -notin @('Citrix', 'AVD', 'WebOnly')) { $hybridDefault = $webShellPreferredProvider }
    if ($hybridDefault -notin @('Citrix', 'AVD', 'WebOnly')) { $hybridDefault = 'Citrix' }
    $hybridSelection = [bool](Get-ConfigValue -Config $Config -Name 'HybridSelectionAtStartup' -Default $true)
    $settingsPages = @(Get-ConfigValue -Config $Config -Name 'WebShellAllowedSettingsPages' -Default @('ms-settings:network-status', 'ms-settings:display', 'ms-settings:sound', 'ms-settings:dateandtime')) |
        ForEach-Object { [string]$_ } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

    $citrixPathLiteral = ConvertTo-QuotedPowerShellString -Value $citrixPath
    $avdPathLiteral = ConvertTo-QuotedPowerShellString -Value $avdPath
    $browserPathLiteral = ConvertTo-QuotedPowerShellString -Value $browserPath
    $citrixWebUrlLiteral = ConvertTo-QuotedPowerShellString -Value $citrixWebUrl
    $avdWebUrlLiteral = ConvertTo-QuotedPowerShellString -Value $avdWebUrl
    $webUrlLiteral = ConvertTo-QuotedPowerShellString -Value $webUrl
    $windowsAppUriLiteral = ConvertTo-QuotedPowerShellString -Value $windowsAppUri
    $webShellTitleLiteral = ConvertTo-QuotedPowerShellString -Value $webShellTitle
    $hybridDefaultLiteral = ConvertTo-QuotedPowerShellString -Value $hybridDefault
    $settingsPagesLiteral = '@({0})' -f (($settingsPages | ForEach-Object { ConvertTo-QuotedPowerShellString -Value $_ }) -join ', ')
    $settingsButtons = ConvertTo-ThinClientSettingsButtons -Value (Get-ConfigValue -Config $Config -Name 'WebShellSettingsButtons' -Default $null)
    $settingsButtonsLiteral = '@{' + (($settingsButtons.Keys | ForEach-Object { (ConvertTo-QuotedPowerShellString -Value $_) + ' = $' + ([string]$settingsButtons[$_]).ToLowerInvariant() }) -join '; ') + '}'

    $webViewHelperLiteral = ConvertTo-QuotedPowerShellString -Value (Join-Path $script:ScriptRoot 'SmartThinClient-WebView2.ps1')
    $deviceSettingsHelperLiteral = ConvertTo-QuotedPowerShellString -Value (Join-Path $script:ScriptRoot 'SmartThinClient-DeviceSettings.ps1')
    $windowIconLiteral = ConvertTo-QuotedPowerShellString -Value (Join-Path $script:ScriptRoot 'Assets\WorkplaceCloudHub.ico')
    $controlPanelScriptLiteral = ConvertTo-QuotedPowerShellString -Value (Join-Path $script:ScriptRoot 'SmartThinClient-Shell.ps1')
    $configuredPath = [string](Get-Variable -Name ConfigPath -ValueOnly -ErrorAction SilentlyContinue)
    $controlPanelConfig = if([string]::IsNullOrWhiteSpace($configuredPath)){Join-Path $script:ScriptRoot 'SmartThinClient-Shell.config.json'}else{$configuredPath}
    $controlPanelConfigLiteral = ConvertTo-QuotedPowerShellString -Value $controlPanelConfig
    $sharedUiFunctions = (@('New-ThinClientUiResources','Get-ThinClientProviderVisibility','ConvertTo-ThinClientSettingsButtons','Get-ThinClientVisibleSettingsPages','New-ThinClientSettingsRequest','Show-ThinClientLimitedSettings','Initialize-ThinClientAdminTokenApi','Get-ThinClientAdminTokenState','Open-ThinClientAdminPanel','Confirm-ThinClientPowerAction','Invoke-ThinClientPowerAction') | ForEach-Object { 'function ' + $_ + ' {' + (Get-Command -Name $_ -CommandType Function).Definition + '}' }) -join "`r`n"
    $shellXaml = Get-ThinClientShellXaml
    $content = @"
param([ValidateSet('Citrix','AVD','WebOnly','Hybrid')][string]`$Profile = '$EffectiveProfile')
`$ErrorActionPreference = 'Stop'

`$CitrixPath = $citrixPathLiteral
`$AvdPath = $avdPathLiteral
`$BrowserPath = $browserPathLiteral
`$CitrixWebUrl = $citrixWebUrlLiteral
`$AvdWebUrl = $avdWebUrlLiteral
`$WebOnlyUrl = $webUrlLiteral
`$WindowsAppUri = $windowsAppUriLiteral
`$ControlPanelScript = $controlPanelScriptLiteral
`$ControlPanelConfig = $controlPanelConfigLiteral
`$DeviceSettingsHelper = $deviceSettingsHelperLiteral
`$WebShellTitle = $webShellTitleLiteral
`$SettingsPages = $settingsPagesLiteral
`$SettingsButtons = $settingsButtonsLiteral
`$BrowserKioskMode = [bool]::Parse('$browserKiosk')
`$UseWebShell = [bool]::Parse('$useWebShell')
`$PreferredAccessMode = '$preferredAccessMode'
`$ShowProviderButtons = [bool]::Parse('$showProviderButtons')
`$AllowExternalBrowser = [bool]::Parse('$allowExternalBrowser')
`$AllowPowerControls = [bool]::Parse('$allowPowerControls')
`$PrimaryCustomShell = [bool]::Parse('$primaryCustomShell')
`$script:ShellExitRequested = `$false
`$AllowLimitedSettings = [bool]::Parse('$allowLimitedSettings')
`$HybridSelectionAtStartup = [bool]::Parse('$hybridSelection')
`$HybridDefaultProvider = $hybridDefaultLiteral

function Start-WorkspaceProcess {
    param([string]`$FilePath, [string]`$Arguments)
    if ([string]::IsNullOrWhiteSpace(`$FilePath)) { return `$false }
    if (-not (Test-Path -LiteralPath `$FilePath)) { return `$false }
    if ([string]::IsNullOrWhiteSpace(`$Arguments)) {
        Start-Process -FilePath `$FilePath -ErrorAction Stop
    }
    else {
        Start-Process -FilePath `$FilePath -ArgumentList `$Arguments -ErrorAction Stop
    }
    return `$true
}

function Show-LauncherNotice {
    param([string]`$Message)
    Add-Type -AssemblyName PresentationFramework
    [System.Windows.MessageBox]::Show(`$Message, 'Smart ThinClient Shell', 'OK', 'Information') | Out-Null
}

function Get-ProviderUrl {
    param([string]`$Provider)
    switch (`$Provider) {
        'Citrix' {
            if (-not [string]::IsNullOrWhiteSpace(`$CitrixWebUrl)) { return `$CitrixWebUrl }
            return `$WebOnlyUrl
        }
        'AVD' {
            if (-not [string]::IsNullOrWhiteSpace(`$AvdWebUrl)) { return `$AvdWebUrl }
            return `$WebOnlyUrl
        }
        default { return `$WebOnlyUrl }
    }
}

function Start-WebWorkspace {
    param([string]`$Provider = 'WebOnly')
    `$url = Get-ProviderUrl -Provider `$Provider
    if ([string]::IsNullOrWhiteSpace(`$url)) {
        Show-LauncherNotice "No web URL is configured for `$Provider. Nothing was installed or changed."
        return
    }
    if (-not [string]::IsNullOrWhiteSpace(`$BrowserPath) -and (Test-Path -LiteralPath `$BrowserPath)) {
        if (`$BrowserKioskMode) {
            Start-Process -FilePath `$BrowserPath -ArgumentList @('--kiosk', `$url, '--edge-kiosk-type=fullscreen')
        }
        else {
            Start-Process -FilePath `$BrowserPath -ArgumentList `$url
        }
    }
    else {
        Start-Process `$url
    }
}

function Start-CitrixWorkspace {
    if (`$PreferredAccessMode -eq 'Native' -and (Start-WorkspaceProcess -FilePath `$CitrixPath -Arguments '')) { return }
    Start-WebWorkspace -Provider 'Citrix'
}

function Start-AvdWorkspace {
    if (`$PreferredAccessMode -eq 'Native' -and (Start-WorkspaceProcess -FilePath `$AvdPath -Arguments '')) { return }
    if (`$PreferredAccessMode -eq 'Native' -and -not [string]::IsNullOrWhiteSpace(`$WindowsAppUri)) { Start-Process `$WindowsAppUri; return }
    Start-WebWorkspace -Provider 'AVD'
}

function Open-ExternalBrowser {
    param([string]`$Url)
    if ([string]::IsNullOrWhiteSpace(`$Url)) {
        Show-LauncherNotice 'No web URL is configured.'
        return
    }
    if (-not [string]::IsNullOrWhiteSpace(`$BrowserPath) -and (Test-Path -LiteralPath `$BrowserPath)) {
        Start-Process -FilePath `$BrowserPath -ArgumentList `$Url
    }
    else {
        Start-Process `$Url
    }
}

function Select-HybridProvider {
    if (-not `$HybridSelectionAtStartup) { return `$HybridDefaultProvider }
    Add-Type -AssemblyName PresentationFramework
    `$choice = New-Object Windows.Window
    `$choice.Title = 'Smart ThinClient - choose workspace'
    `$choice.Width = 360; `$choice.Height = 220
    `$choice.WindowStartupLocation = 'CenterScreen'
    `$panel = New-Object Windows.Controls.StackPanel
    foreach (`$provider in @('Citrix','AVD','WebOnly')) {
        `$button = New-Object Windows.Controls.Button
        `$button.Content = `$provider; `$button.Tag = `$provider
        `$button.Height = 42; `$button.Margin = '10,5,10,5'
        `$button.Add_Click({ param(`$sender,`$eventArgs) `$choice.Tag = [string]`$sender.Tag; `$choice.Close() })
        [void]`$panel.Children.Add(`$button)
    }
    `$choice.Content = `$panel
    [void]`$choice.ShowDialog()
    return [string]`$choice.Tag
}

$sharedUiFunctions
function Show-LimitedSettings {
    `$nativeSettings = `$PrimaryCustomShell -or -not @(Get-Process explorer -ErrorAction SilentlyContinue | Where-Object SessionId -eq ([Diagnostics.Process]::GetCurrentProcess().SessionId)).Count
    Show-ThinClientLimitedSettings -SettingsPages `$SettingsPages -SettingsButtons `$SettingsButtons -Owner `$window -NativeMode:`$nativeSettings -NativeSettingsHelper `$DeviceSettingsHelper
}

function Show-WebShell {
    param([string]`$InitialProvider)

    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    [xml]`$xaml = @'
$shellXaml
'@

    `$reader = New-Object System.Xml.XmlNodeReader `$xaml
    `$window = [Windows.Markup.XamlReader]::Load(`$reader)
    [void]`$window.Resources.MergedDictionaries.Add((New-ThinClientUiResources))
    if (Test-Path -LiteralPath $windowIconLiteral) { `$window.Icon = [Windows.Media.Imaging.BitmapFrame]::Create([Uri]::new($windowIconLiteral,[UriKind]::Absolute)) }
    `$browserHost = `$window.FindName('WorkspaceBrowser')
    `$browser = `$null
    `$browserError = ''
    try { . $webViewHelperLiteral; `$browser = New-SmartThinClientWebView2; `$browserHost.Content = `$browser }
    catch { `$browserError = `$_.Exception.Message; `$errorView = [Windows.Controls.TextBlock]::new();`$errorView.Text = 'The workspace browser could not initialize. Use Admin to review the configuration. '+`$browserError;`$errorView.TextWrapping='Wrap';`$errorView.Margin='24';`$browserHost.Content=`$errorView }
    `$titleText = `$window.FindName('TitleText')
    `$statusText = `$window.FindName('StatusText')
    if(`$browser -and `$browser.Tag -and `$browser.Tag.PSObject.Properties['StatusTarget']){`$browser.Tag.StatusTarget=`$statusText}
    if (`$PrimaryCustomShell) {
        `$window.WindowStyle = 'None'; `$window.ResizeMode = 'NoResize'
        `$window.Add_Closing({param(`$sender,`$eventArgs) if(-not `$script:ShellExitRequested){`$eventArgs.Cancel=`$true}})
    }
    `$citrixButton = `$window.FindName('CitrixButton')
    `$avdButton = `$window.FindName('AvdButton')
    `$webOnlyButton = `$window.FindName('WebOnlyButton')
    `$backButton = `$window.FindName('BackButton')
    `$refreshButton = `$window.FindName('RefreshButton')
    `$externalButton = `$window.FindName('ExternalButton')
    `$settingsButton = `$window.FindName('SettingsButton')
    `$adminButton = `$window.FindName('AdminButton')
    if ((Get-ThinClientAdminTokenState).IsAdministratorAccount) { `$adminButton.Visibility = 'Visible' }
    `$adminButton.Tag = [pscustomobject]@{ScriptPath=`$ControlPanelScript;ConfigPath=`$ControlPanelConfig;WorkspaceProfile=[string]`$Profile}
    `$adminButton.Add_Click({
        try { `$request=`$this.Tag;Open-ThinClientAdminPanel -ScriptPath `$request.ScriptPath -ConfigPath `$request.ConfigPath -Profile `$request.WorkspaceProfile }
        catch { Show-LauncherNotice ('Admin panel could not open: '+`$_.Exception.Message) }
    })
    `$signOutButton = `$window.FindName('SignOutButton')
    `$restartButton = `$window.FindName('RestartButton')
    `$shutdownButton = `$window.FindName('ShutdownButton')
    `$script:CurrentUrl = ''

    `$titleText.Text = `$WebShellTitle
    `$titleText.ToolTip = `$WebShellTitle
    `$window.Title = `$WebShellTitle
    `$providerVisibility = Get-ThinClientProviderVisibility -EffectiveProfile `$Profile -ShowProviderButtons `$ShowProviderButtons
    if (-not `$providerVisibility.Citrix) { `$citrixButton.Visibility = 'Collapsed' }
    if (-not `$providerVisibility.AVD) { `$avdButton.Visibility = 'Collapsed' }
    if (-not `$providerVisibility.WebOnly) { `$webOnlyButton.Visibility = 'Collapsed' }
    if (-not `$AllowExternalBrowser) { `$externalButton.Visibility = 'Collapsed' }
    if (-not `$AllowLimitedSettings) { `$settingsButton.Visibility = 'Collapsed' }
    if (-not `$AllowPowerControls) {
        `$signOutButton.Visibility = 'Collapsed'
        `$restartButton.Visibility = 'Collapsed'
        `$shutdownButton.Visibility = 'Collapsed'
    }

    `$navigate = {
        param([string]`$Provider)
        if (`$null -eq `$browser) { `$statusText.Text = 'Browser unavailable: '+`$browserError; return }
        `$script:CurrentProvider = `$Provider
        `$url = Get-ProviderUrl -Provider `$Provider
        if ([string]::IsNullOrWhiteSpace(`$url)) {
            `$statusText.Text = "No web URL configured for `$Provider."
            Show-LauncherNotice "No web URL is configured for `$Provider. Edit the local JSON configuration."
            return
        }
        `$script:CurrentUrl = `$url
        `$statusText.Text = "Opening `$Provider - `$url"
        try { `$browser.Source = [Uri]`$url }
        catch { `$statusText.Text = 'Navigation failed: '+`$_.Exception.Message }
    }

    `$citrixButton.Add_Click({ & `$navigate 'Citrix' })
    `$avdButton.Add_Click({ & `$navigate 'AVD' })
    `$webOnlyButton.Add_Click({ & `$navigate 'WebOnly' })
    `$backButton.Add_Click({try {if (`$browser -and `$browser.CanGoBack) { `$browser.GoBack() }}catch{`$statusText.Text='Back failed: '+`$_.Exception.Message}})
    `$refreshButton.Add_Click({try {if(`$browser){`$browser.Reload()}}catch{`$statusText.Text='Refresh failed: '+`$_.Exception.Message}})
    `$externalButton.Add_Click({try {Open-ExternalBrowser -Url `$script:CurrentUrl}catch{`$statusText.Text='Browser launch failed: '+`$_.Exception.Message}})
    `$settingsButton.Add_Click({try {Show-LimitedSettings}catch{`$statusText.Text='Settings failed: '+`$_.Exception.Message}})
    `$signOutButton.Add_Click({ `$script:ShellExitRequested=`$true;shutdown.exe /l })
    `$restartButton.Add_Click({
        try {[void](Invoke-ThinClientPowerAction -Owner `$window -Action Restart)}
        catch {`$statusText.Text='Restart failed: '+`$_.Exception.Message}
    })
    `$shutdownButton.Add_Click({
        try {[void](Invoke-ThinClientPowerAction -Owner `$window -Action Shutdown)}
        catch {`$statusText.Text='Shutdown failed: '+`$_.Exception.Message}
    })

    if (`$InitialProvider -eq 'Hybrid') { `$InitialProvider = `$HybridDefaultProvider }
    `$window.Add_ContentRendered({ & `$navigate `$InitialProvider })
    `$window.ShowDialog() | Out-Null
}

if (`$Profile -eq 'Hybrid' -and -not `$PrimaryCustomShell) {
    `$Profile = Select-HybridProvider
    if ([string]::IsNullOrWhiteSpace(`$Profile)) { exit 0 }
}
if (`$UseWebShell -and `$PreferredAccessMode -eq 'Web') {
    Show-WebShell -InitialProvider `$Profile
}
else {
    switch (`$Profile) {
        'Citrix' { Start-CitrixWorkspace }
        'AVD' { Start-AvdWorkspace }
        'WebOnly' { Start-WebWorkspace -Provider 'WebOnly' }
        default { Start-WebWorkspace -Provider 'WebOnly' }
    }
}
"@

    Set-Content -LiteralPath $launcherPath -Value $content -Encoding UTF8
    return $launcherPath
}

function Get-TargetUserSystemPolicyPath {
    param([Parameter(Mandatory = $true)][string]$Sid)
    $parsedSid = New-Object Security.Principal.SecurityIdentifier($Sid)
    return ('Registry::HKEY_USERS\{0}\Software\Policies\Microsoft\Windows\System' -f $parsedSid.Value)
}

function Initialize-ThinClientCustomShellApi {
    if ('SmartThinClient.LocalUserShellPolicy' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;
using Microsoft.Win32.SafeHandles;
namespace SmartThinClient {
 [ComImport, Guid("EA502723-A23D-11D1-A7D3-0000F87571E3"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
 internal interface ILocalPolicy {
  [PreserveSig] int New([MarshalAs(UnmanagedType.LPWStr)] string a,[MarshalAs(UnmanagedType.LPWStr)] string b,uint c);
  [PreserveSig] int OpenDSGPO([MarshalAs(UnmanagedType.LPWStr)] string a,uint b);
  [PreserveSig] int OpenLocalMachineGPO(uint flags);
  [PreserveSig] int OpenRemoteMachineGPO([MarshalAs(UnmanagedType.LPWStr)] string a,uint b);
  [PreserveSig] int Save([MarshalAs(UnmanagedType.Bool)] bool machine,[MarshalAs(UnmanagedType.Bool)] bool add,ref Guid extension,ref Guid editor);
  [PreserveSig] int Delete();
  [PreserveSig] int GetName([MarshalAs(UnmanagedType.LPWStr)] StringBuilder a,int b);
  [PreserveSig] int GetDisplayName([MarshalAs(UnmanagedType.LPWStr)] StringBuilder a,int b);
  [PreserveSig] int SetDisplayName([MarshalAs(UnmanagedType.LPWStr)] string a);
  [PreserveSig] int GetPath([MarshalAs(UnmanagedType.LPWStr)] StringBuilder a,int b);
  [PreserveSig] int GetDSPath(uint a,[MarshalAs(UnmanagedType.LPWStr)] StringBuilder b,int c);
  [PreserveSig] int GetFileSysPath(uint a,[MarshalAs(UnmanagedType.LPWStr)] StringBuilder b,int c);
  [PreserveSig] int GetRegistryKey(uint section,out IntPtr handle);
  [PreserveSig] int GetOptions(out uint a);
  [PreserveSig] int SetOptions(uint a,uint b);
  [PreserveSig] int GetType(out uint a);
  [PreserveSig] int GetMachineName([MarshalAs(UnmanagedType.LPWStr)] StringBuilder a,int b);
  [PreserveSig] int GetPropertySheetPages(out IntPtr a,out uint b);
 }
 public sealed class ShellPolicyState {
  public bool ValueExists;
  public string ValueKind;
  public string Value;
 }
 public static class LocalUserShellPolicy {
  const string KeyPath = @"Software\Microsoft\Windows\CurrentVersion\Policies\System";
  static void Check(int result) { if(result != 0) Marshal.ThrowExceptionForHR(result < 0 ? result : unchecked((int)0x80004005)); }
  static ShellPolicyState ReadKey(RegistryKey root) {
   using(var key=root.OpenSubKey(KeyPath,false)) {
    if(key==null || Array.FindIndex(key.GetValueNames(),n=>String.Equals(n,"Shell",StringComparison.OrdinalIgnoreCase))<0)
     return new ShellPolicyState {ValueExists=false,ValueKind="",Value=""};
    var kind=key.GetValueKind("Shell");
    if(kind!=RegistryValueKind.String && kind!=RegistryValueKind.ExpandString) throw new InvalidOperationException("Existing custom shell policy has an unsupported value type.");
    return new ShellPolicyState {ValueExists=true,ValueKind=kind.ToString(),Value=(string)key.GetValue("Shell","",RegistryValueOptions.DoNotExpandEnvironmentNames)};
   }
  }
  static bool Equal(ShellPolicyState a,ShellPolicyState b) {return a.ValueExists==b.ValueExists && (!a.ValueExists || (a.ValueKind==b.ValueKind && String.Equals(a.Value,b.Value,StringComparison.Ordinal)));}
  static ShellPolicyState Access(bool write,ShellPolicyState desired,ShellPolicyState expected) {
   object instance=null;
   try {
    instance=Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("EA502722-A23D-11D1-A7D3-0000F87571E3"),true));
    var policy=(ILocalPolicy)instance;
    Check(policy.OpenLocalMachineGPO(write ? 1u : 3u));
    uint options;Check(policy.GetOptions(out options));
    if((options&1u)!=0) throw new InvalidOperationException("Local user policy is disabled; no shell policy changed.");
    IntPtr native;Check(policy.GetRegistryKey(1,out native));
    using(var handle=new SafeRegistryHandle(native,true)) using(var root=RegistryKey.FromHandle(handle)) {
     var current=ReadKey(root);
     if(!write) return current;
     if(expected==null || !Equal(current,expected)) throw new InvalidOperationException("Custom shell policy changed after review; no policy saved.");
     if(desired.ValueExists) {
      RegistryValueKind kind;
      if(!Enum.TryParse(desired.ValueKind,out kind) || (kind!=RegistryValueKind.String && kind!=RegistryValueKind.ExpandString) || String.IsNullOrWhiteSpace(desired.Value)) throw new InvalidOperationException("Invalid shell policy state.");
      using(var key=root.CreateSubKey(KeyPath)) key.SetValue("Shell",desired.Value,kind);
     } else {using(var key=root.OpenSubKey(KeyPath,true)) {if(key!=null) key.DeleteValue("Shell",false);}}
     var extension=new Guid("35378EAC-683F-11D2-A89A-00C04FBBCFA2");
     var editor=new Guid("8FC0B734-A0E1-11D1-A7D3-0000F87571E3");
     Check(policy.Save(false,true,ref extension,ref editor));
     return ReadKey(root);
    }
   } finally {if(instance!=null && Marshal.IsComObject(instance)) Marshal.FinalReleaseComObject(instance);}
  }
  public static ShellPolicyState Read() {return Access(false,null,null);}
  public static ShellPolicyState Write(ShellPolicyState desired,ShellPolicyState expected) {return Access(true,desired,expected);}
 }
}
'@ -ErrorAction Stop
}

function Get-ThinClientCustomShellPolicy {
    Initialize-ThinClientCustomShellApi
    $state = [SmartThinClient.LocalUserShellPolicy]::Read()
    [pscustomobject]@{Scope='LocalPolicyAllUsers';ValueExists=[bool]$state.ValueExists;ValueKind=[string]$state.ValueKind;Value=[string]$state.Value}
}

function Set-ThinClientCustomShellPolicy {
    param($Desired,$Expected)
    if (-not (Test-IsAdministrator)) { throw 'Changing the all-users shell policy requires an elevated administrator.' }
    Initialize-ThinClientCustomShellApi
    $states=@(foreach($item in @($Desired,$Expected)) {
        if($item.Scope -ne 'LocalPolicyAllUsers' -or $item.ValueExists -isnot [bool] -or ($item.ValueExists -and ($item.ValueKind -notin @('String','ExpandString') -or [string]::IsNullOrWhiteSpace($item.Value)))) { throw 'Invalid custom shell policy snapshot.' }
        $native=[SmartThinClient.ShellPolicyState]::new();$native.ValueExists=$item.ValueExists;$native.ValueKind=$item.ValueKind;$native.Value=$item.Value;$native
    })
    $result=[SmartThinClient.LocalUserShellPolicy]::Write($states[0],$states[1])
    if($result.ValueExists -ne $Desired.ValueExists -or ($Desired.ValueExists -and ($result.ValueKind -ne $Desired.ValueKind -or $result.Value -cne $Desired.Value))) { throw 'Custom shell policy verification failed; retain the rollback for recovery.' }
}

function Get-ThinClientMachineRoot {
    param($Config)
    $programData=[IO.Path]::GetFullPath($env:ProgramData).TrimEnd('\')
    $path=[IO.Path]::GetFullPath([string](Get-ConfigValue -Config $Config -Name MachineRoot -Default (Join-Path $programData 'SmartThinClient\Machine'))).TrimEnd('\')
    if(-not $path.StartsWith($programData+'\',[StringComparison]::OrdinalIgnoreCase) -or $path -eq $programData){throw 'MachineRoot must be a dedicated local directory under ProgramData.'}
    for($current=$path;$current.Length -ge $programData.Length;$current=Split-Path -Path $current -Parent){if((Test-Path -LiteralPath $current) -and ((Get-Item -LiteralPath $current -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'MachineRoot cannot traverse a symbolic link or junction.'};if($current -eq $programData){break}}
    $path
}

function Initialize-ThinClientMachineRoot {
    param($Config)
    if(-not(Test-IsAdministrator)){throw 'Machine storage requires an elevated administrator.'}
    $path=Get-ThinClientMachineRoot -Config $Config
    $marker=Join-Path $path 'SmartThinClient-Machine.json'
    if((Test-Path -LiteralPath $path) -and -not(Test-Path -LiteralPath $marker -PathType Leaf)){throw 'Existing MachineRoot is not owned by SmartThinClient; choose a dedicated directory.'}
    if(Test-Path -LiteralPath $marker){
        $owner=(Get-Acl -LiteralPath $path).GetOwner([Security.Principal.SecurityIdentifier]).Value
        if($owner -notin @('S-1-5-18','S-1-5-32-544') -or ((Get-Item -LiteralPath $marker -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Existing machine storage is not administrator-owned.'}
        $ownership=Read-JsonFile -Path $marker;if($ownership.AppName -ne $script:AppName -or $ownership.Root -ine $path){throw 'MachineRoot ownership marker mismatch.'}
    }
    New-Item -ItemType Directory -Path $path -Force|Out-Null
    foreach($folder in @($path,(Join-Path $path 'Launcher'),(Join-Path $path 'Rollback'))) {
        if((Test-Path -LiteralPath $folder) -and ((Get-Item -LiteralPath $folder -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Machine storage cannot contain redirected directories.'}
        New-Item -ItemType Directory -Path $folder -Force|Out-Null
        $acl=[Security.AccessControl.DirectorySecurity]::new();$acl.SetAccessRuleProtection($true,$false)
        $acl.SetOwner([Security.Principal.SecurityIdentifier]::new('S-1-5-32-544'))
        foreach($sid in @('S-1-5-18','S-1-5-32-544')){$acl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new([Security.Principal.SecurityIdentifier]::new($sid),'FullControl','ContainerInherit,ObjectInherit','None','Allow'))}
        if($folder -ne (Join-Path $path 'Rollback')){$acl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new([Security.Principal.SecurityIdentifier]::new('S-1-5-11'),'ReadAndExecute','ContainerInherit,ObjectInherit','None','Allow'))}
        Set-Acl -LiteralPath $folder -AclObject $acl
    }
    [pscustomobject]@{AppName=$script:AppName;Root=$path;SchemaVersion=1}|ConvertTo-Json|Set-Content -LiteralPath $marker -Encoding UTF8
    $path
}

function Protect-ThinClientMachineLauncher {
    param([string]$Path)
    if((Get-Item -LiteralPath $Path -Force).Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Refusing a redirected machine launcher.'}
    $acl=[Security.AccessControl.FileSecurity]::new();$acl.SetAccessRuleProtection($true,$false);$acl.SetOwner([Security.Principal.SecurityIdentifier]::new('S-1-5-32-544'))
    foreach($sid in @('S-1-5-18','S-1-5-32-544')){$acl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new([Security.Principal.SecurityIdentifier]::new($sid),'FullControl','Allow'))}
    $acl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new([Security.Principal.SecurityIdentifier]::new('S-1-5-11'),'ReadAndExecute','Allow'))
    Set-Acl -LiteralPath $Path -AclObject $acl
}

function New-RollbackState {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)]$Audit,
        [Parameter(Mandatory = $true)][string]$EffectiveProfile
    )

    $runKey = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run'
    $explorerPolicy = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer'
    $snapshots = @(
        if ((Get-ConfigValue -Config $Config -Name 'AutoLaunchMode' -Default 'RunKey') -eq 'RunKey' -or (Get-ConfigValue -Config $Config -Name 'ShellMode' -Default 'Launcher') -eq 'CustomShell') {
            Get-RegistryValueSnapshot -Path $runKey -Name 'SmartThinClientShell'
        }
        $level = Get-ConfigValue -Config $Config -Name 'ShellRestrictionLevel' -Default 'None'
        if ((Get-ConfigValue -Config $Config -Name 'EnableShellLimitations' -Default $false) -and $level -in @('Basic', 'Strict')) {
            Get-RegistryValueSnapshot -Path $explorerPolicy -Name 'NoRun'
            Get-RegistryValueSnapshot -Path $explorerPolicy -Name 'NoControlPanel'
            Get-RegistryValueSnapshot -Path $explorerPolicy -Name 'NoViewContextMenu'
            if ($level -eq 'Strict') { Get-RegistryValueSnapshot -Path (Get-TargetUserSystemPolicyPath -Sid $Audit.TargetUserSid) -Name 'DisableCMD' }
        }
    )
    $launcherPath = Join-Path $script:OutputRoot 'Launcher\SmartThinClient-LaunchWorkspace.ps1'

    $state = [ordered]@{
        AppName = $script:AppName
        RunId = $script:RunId
        SchemaVersion = 2
        LauncherExisted = Test-Path -LiteralPath $launcherPath
        LauncherContent = if (Test-Path -LiteralPath $launcherPath) { [Convert]::ToBase64String([IO.File]::ReadAllBytes($launcherPath)) } else { '' }
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
    if((Get-ConfigValue -Config $Config -Name ShellMode -Default Launcher) -eq 'CustomShell') {
        $state.SchemaVersion=3
        $state.CustomShellPolicy=Get-ThinClientCustomShellPolicy
        $state.ManagedShellCommand=Get-LauncherCommandLine -Machine
    }
    return $state
}

function Save-RollbackState {
    param([Parameter(Mandatory = $true)]$State)

    $path = Join-Path $script:RollbackRoot ('SmartThinClient_Rollback_{0}_{1}.json' -f $env:COMPUTERNAME, $script:RunId)
    ([pscustomobject]$State) | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $path -Encoding UTF8
    Write-SmartThinClientLog -Message "Rollback state saved: $path"
    return $path
}

function Assert-ConfigValues {
    param([Parameter(Mandatory = $true)]$Config)
    $choices = @{
        DefaultProfile = @('Auto','Citrix','AVD','WebOnly','Hybrid')
        PreferredAccessMode = @('Web','Native')
        AutoLaunchMode = @('None','RunKey')
        ShellRestrictionLevel = @('None','Basic','Strict')
        ShellMode = @('Launcher','AssignedAccess','ShellLauncher','CustomShell')
        TargetUserMode = @('ExistingUser','DedicatedUser')
    }
    foreach ($key in $choices.Keys) {
        if ($Config.Contains($key) -and $Config[$key] -notin $choices[$key]) { throw "Invalid configuration value for $key." }
    }
    foreach ($key in @('AllowApply','AllowRestore','RequireConfirmationPhrase','CreateDedicatedLocalUser','UseWebShell','BrowserKioskMode','EnableShellLimitations','EnableAssignedAccess','EnableShellLauncher','EnableCustomShell','HybridSelectionAtStartup','WebShellShowProviderButtons','WebShellAllowExternalBrowser','WebShellAllowPowerControls','WebShellAllowLimitedSettings','WriteCsvEvidence')) {
        if ($Config.Contains($key) -and $Config[$key] -isnot [bool]) { throw "$key must be a JSON boolean, not a string." }
    }
    [void](ConvertTo-ThinClientSettingsButtons -Value (Get-ConfigValue -Config $Config -Name 'WebShellSettingsButtons' -Default $null))
}

function Assert-CimSuccess {
    param($Response)
    if ($null -eq $Response -or $null -eq $Response.PSObject.Properties['ReturnValue'] -or $Response.ReturnValue -ne 0) {
        throw 'Shell Launcher CIM operation did not return success; review the rollback file and log.'
    }
}

function Assert-ApplyPreflight {
    param($Config, $Audit, $TargetUser)
    Assert-ConfigValues -Config $Config
    $assigned = [bool](Get-ConfigValue -Config $Config -Name 'EnableAssignedAccess' -Default $false)
    $shell = [bool](Get-ConfigValue -Config $Config -Name 'EnableShellLauncher' -Default $false)
    $mode = Get-ConfigValue -Config $Config -Name 'ShellMode' -Default 'Launcher'
    if($mode -eq 'CustomShell') {
        if(-not(Get-ConfigValue -Config $Config -Name EnableCustomShell -Default $false)){throw 'CustomShell requires EnableCustomShell=true.'}
        if($assigned -or $shell -or (Get-ConfigValue -Config $Config -Name CreateDedicatedLocalUser -Default $false)){throw 'The all-users custom shell cannot be combined with per-user kiosk or dedicated-account changes.'}
        if(-not(Get-ResultValue -Result $Audit -Name CustomShellPolicyAvailable -Default $false)){throw 'The local Custom User Interface policy is unavailable on this Windows edition.'}
        if(-not(Get-ConfigValue -Config $Config -Name UseWebShell -Default $true) -or (Get-ConfigValue -Config $Config -Name PreferredAccessMode -Default Web) -ne 'Web'){throw 'The all-users custom shell requires the persistent embedded Web shell.'}
        if(-not(Get-ConfigValue -Config $Config -Name RequireConfirmationPhrase -Default $true)){throw 'All-users shell replacement requires an explicit confirmation phrase.'}
        if(-not(Get-ResultValue -Result $Audit -Name BrowserInstalled -Default $false)){throw 'A detected browser is required before replacing Explorer.'}
        $provider=[string](Get-ResultValue -Result $Audit -Name EffectiveProfile -Default WebOnly)
        $urlKeys=if($provider -eq 'Hybrid'){@('CitrixWebUrlConfigured','AvdWebUrlConfigured','WebOnlyUrlConfigured')}elseif($provider -eq 'Citrix'){@('CitrixWebUrlConfigured','WebOnlyUrlConfigured')}elseif($provider -eq 'AVD'){@('AvdWebUrlConfigured','WebOnlyUrlConfigured')}else{@('WebOnlyUrlConfigured')}
        if(@($urlKeys|Where-Object {Get-ResultValue -Result $Audit -Name $_ -Default $false}).Count -eq 0){throw 'A configured workspace Web URL is required before replacing Explorer.'}
        $machineRoot=Get-ThinClientMachineRoot -Config $Config
        $policy=Get-ThinClientCustomShellPolicy
        $managedCommand=Get-LauncherCommandLine -Root $machineRoot -Machine
        if($policy.ValueExists -and $policy.Value -notin @('explorer.exe',$managedCommand)){throw 'A different Custom User Interface policy is already configured; review it before replacement.'}
    }
    if ($assigned -and $shell) { throw 'Assigned Access and Shell Launcher cannot be applied together.' }
    if (($mode -eq 'AssignedAccess' -and -not $assigned) -or ($mode -eq 'ShellLauncher' -and -not $shell)) { throw 'ShellMode requires its matching Enable setting.' }
    if ($TargetUser.Mode -eq 'DedicatedUser' -and -not $TargetUser.LocalUserExists) {
        if (-not (Get-ConfigValue -Config $Config -Name 'CreateDedicatedLocalUser' -Default $false) -or $null -eq $DedicatedUserPassword) { throw 'Dedicated user creation requires CreateDedicatedLocalUser and DedicatedUserPassword.' }
        if (-not (Get-Command New-LocalUser -ErrorAction SilentlyContinue)) { throw 'New-LocalUser is unavailable.' }
    }
    if ((Get-ConfigValue -Config $Config -Name 'EnableShellLimitations' -Default $false) -and (Get-ConfigValue -Config $Config -Name 'ShellRestrictionLevel' -Default 'None') -eq 'Strict') {
        if (-not $TargetUser.LocalUserExists -or -not $TargetUser.Sid) { throw 'Strict restrictions require an existing local target user with a SID.' }
        [void](Get-TargetUserSystemPolicyPath -Sid $TargetUser.Sid)
        if (-not (Test-Path -LiteralPath ('Registry::HKEY_USERS\{0}' -f $TargetUser.Sid))) { throw 'Strict restrictions require the target user registry hive to be loaded. Sign in that user before applying.' }
    }
    if ($assigned -or $shell) {
        if (-not $TargetUser.LocalUserExists -or -not $TargetUser.Enabled -or -not $TargetUser.Sid) { throw 'Kiosk modes require an existing enabled local user; prepare and verify that user separately.' }
        $admins = @(Get-LocalGroupMember -SID 'S-1-5-32-544' -ErrorAction Stop)
        if ($TargetUser.Sid -in @($admins | ForEach-Object { [string]$_.SID })) { throw 'Kiosk target must not be a local administrator.' }
    }
    if ($assigned) {
        if (-not $Audit.AssignedAccessCmdletAvailable -or -not (Get-Command Get-AssignedAccess -ErrorAction SilentlyContinue) -or -not (Get-Command Clear-AssignedAccess -ErrorAction SilentlyContinue)) { throw 'Assigned Access inspection/apply/restore cmdlets are required.' }
        if ([string]::IsNullOrWhiteSpace([string](Get-ConfigValue -Config $Config -Name 'AssignedAccessAppUserModelId' -Default ''))) { throw 'AssignedAccessAppUserModelId is required.' }
        if (@(Get-AssignedAccess -ErrorAction Stop).Count) { throw 'Existing Assigned Access configuration detected; this tool will not overwrite it.' }
    }
    if ($shell) {
        if (-not $Audit.ShellLauncherClassAvailable) { throw 'Shell Launcher class is unavailable.' }
        $enabled = Invoke-CimMethod -Namespace 'root\standardcimv2\embedded' -ClassName 'WESL_UserSetting' -MethodName 'IsEnabled' -ErrorAction Stop
        Assert-CimSuccess -Response $enabled
        if (-not $enabled.Enabled) { throw 'Shell Launcher must already be enabled by an administrator; this tool does not enable Windows features.' }
        $existing = @(Get-CimInstance -Namespace 'root\standardcimv2\embedded' -ClassName 'WESL_UserSetting' -ErrorAction Stop | Where-Object { [string]$_.Sid -eq $TargetUser.Sid })
        if ($existing.Count) { throw 'Existing custom shell detected for this SID; this tool will not overwrite it.' }
        $defaultAction = Get-ConfigValue -Config $Config -Name 'ShellLauncherDefaultAction' -Default 0
        if ($defaultAction -notin @(0,1,2,3)) { throw 'ShellLauncherDefaultAction must be 0, 1, 2 or 3.' }
        if (-not (Get-ConfigValue -Config $Config -Name 'UseWebShell' -Default $true) -or (Get-ConfigValue -Config $Config -Name 'PreferredAccessMode' -Default 'Web') -ne 'Web') { throw 'Shell Launcher requires the persistent embedded web shell; external/native launchers exit immediately.' }
    }
}

function Assert-RollbackState {
    param($State,$Config=([ordered]@{}))
    foreach ($name in @('SchemaVersion','AppName','ComputerName','Config','Registry','OutputRoot','LauncherPath','LauncherExisted','LauncherContent','TargetUserSid')) {
        if ($null -eq $State.PSObject.Properties[$name]) { throw "Incomplete or legacy rollback: missing $name. Manual review is required." }
    }
    if ($State.SchemaVersion -notin @(2,3) -or $State.AppName -ne $script:AppName -or $State.ComputerName -ne $env:COMPUTERNAME) { throw 'Rollback schema, application or computer does not match; no state restored.' }
    $expectedRoot = if($State.SchemaVersion -eq 3){Get-ThinClientMachineRoot -Config $Config}else{$script:OutputRoot}
    $expected = Join-Path $expectedRoot 'Launcher\SmartThinClient-LaunchWorkspace.ps1'
    if ([IO.Path]::GetFullPath($State.OutputRoot) -ne [IO.Path]::GetFullPath($expectedRoot) -or [IO.Path]::GetFullPath($State.LauncherPath) -ne [IO.Path]::GetFullPath($expected)) { throw 'Rollback output/launcher path mismatch. Use the original OutputRoot or MachineRoot.' }
    if($State.SchemaVersion -eq 3) {
        if($null -eq $State.PSObject.Properties['CustomShellPolicy'] -or $null -eq $State.PSObject.Properties['ManagedShellCommand']){throw 'Incomplete custom-shell rollback.'}
        $policy=$State.CustomShellPolicy
        if($policy.Scope -ne 'LocalPolicyAllUsers' -or $policy.ValueExists -isnot [bool] -or ($policy.ValueExists -and ($policy.ValueKind -notin @('String','ExpandString') -or [string]::IsNullOrWhiteSpace($policy.Value)))){throw 'Invalid custom-shell rollback policy.'}
        if($State.ManagedShellCommand -cne (Get-LauncherCommandLine -Root $expectedRoot -Machine)){throw 'Custom-shell rollback command mismatch.'}
        if($policy.ValueExists -and $policy.Value -notin @('explorer.exe',$State.ManagedShellCommand)){throw 'Rollback contains an unrelated custom shell command.'}
        if($State.Config.ShellMode -ne 'CustomShell' -or $State.Config.EnableCustomShell -ne $true){throw 'Rollback does not describe an enabled all-users custom shell.'}
        if((Test-Path -LiteralPath $expected) -and ((Get-Item -LiteralPath $expected -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Rollback launcher path is redirected.'}
    }
    if ($State.LauncherExisted -isnot [bool]) { throw 'Invalid rollback launcher state.' }
    if ($State.LauncherExisted) { [void][Convert]::FromBase64String($State.LauncherContent) }
    $allowed = @{
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' = @('SmartThinClientShell')
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer' = @('NoRun','NoControlPanel','NoViewContextMenu')
        'HKLM:\SOFTWARE\Policies\Microsoft\Windows\System' = @('DisableCMD')
    }
    if ($State.TargetUserSid) {
        $allowed[(Get-TargetUserSystemPolicyPath -Sid $State.TargetUserSid)] = @('DisableCMD')
    }
    foreach ($snapshot in @($State.Registry)) {
        if (-not $allowed.ContainsKey([string]$snapshot.Path) -or $snapshot.Name -notin $allowed[[string]$snapshot.Path]) { throw 'Rollback contains an unrelated registry value.' }
        if ($snapshot.ValueExists -isnot [bool] -or ($snapshot.ValueExists -and $snapshot.ValueKind -notin @('String','ExpandString','Binary','DWord','MultiString','QWord','None'))) { throw 'Invalid registry snapshot.' }
        if ([string]$snapshot.Path -like 'Registry::HKEY_USERS\*' -and -not (Test-Path -LiteralPath ('Registry::HKEY_USERS\{0}' -f $State.TargetUserSid))) { throw 'Restore requires the original target user registry hive to be loaded.' }
    }
    Assert-ConfigValues -Config (Convert-ObjectToHashtable -InputObject $State.Config)
}

function Assert-ApplyAllowed {
    param([Parameter(Mandatory = $true)]$Config)

    if (-not (Test-IsAdministrator)) { throw 'Apply requires an elevated administrator PowerShell session.' }
    if (-not [bool](Get-ConfigValue -Config $Config -Name 'AllowApply' -Default $false)) { throw 'Apply is blocked because AllowApply is false in configuration.' }
    if ([bool](Get-ConfigValue -Config $Config -Name 'RequireConfirmationPhrase' -Default $true)) {
        $phrase = [string](Get-ConfigValue -Config $Config -Name 'ApplyConfirmationPhrase' -Default 'APPLY')
        if ($ConfirmApply -ne $phrase) { throw "Apply requires -ConfirmApply '$phrase'." }
    }
}

function Assert-RestoreAllowed {
    param([Parameter(Mandatory = $true)]$Config)

    if (-not (Test-IsAdministrator)) { throw 'Restore requires an elevated administrator PowerShell session.' }
    if (-not [bool](Get-ConfigValue -Config $Config -Name 'AllowRestore' -Default $false)) { throw 'Restore is blocked because AllowRestore is false in configuration.' }
    if ([bool](Get-ConfigValue -Config $Config -Name 'RequireConfirmationPhrase' -Default $true)) {
        $phrase = [string](Get-ConfigValue -Config $Config -Name 'RestoreConfirmationPhrase' -Default 'RESTORE')
        if ($ConfirmRestore -ne $phrase) { throw "Restore requires -ConfirmRestore '$phrase'." }
    }
}

function Apply-AutoLaunch {
    param([Parameter(Mandatory = $true)]$Config, [Parameter(Mandatory = $true)][string]$CommandLine)
    $mode = [string](Get-ConfigValue -Config $Config -Name 'AutoLaunchMode' -Default 'RunKey')
    if ($mode -eq 'None') { return }
    if ($mode -ne 'RunKey') { throw 'AutoLaunchMode must be None or RunKey.' }
    Set-StringRegistryValue -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' -Name 'SmartThinClientShell' -Value $CommandLine
    Write-SmartThinClientLog -Message 'Configured machine Run key for Smart ThinClient launcher.'
}

function Apply-ShellLimitations {
    param([Parameter(Mandatory = $true)]$Config, $Audit)

    if (-not [bool](Get-ConfigValue -Config $Config -Name 'EnableShellLimitations' -Default $false)) {
        Write-SmartThinClientLog -Message 'Shell limitations disabled by configuration.'
        return
    }

    $level = [string](Get-ConfigValue -Config $Config -Name 'ShellRestrictionLevel' -Default 'None')
    if ($level -eq 'None') { return }
    if ($level -notin @('Basic', 'Strict')) { throw 'Invalid ShellRestrictionLevel.' }
    $explorerPolicy = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer'

    Set-DwordRegistryValue -Path $explorerPolicy -Name 'NoRun' -Value 1
    Set-DwordRegistryValue -Path $explorerPolicy -Name 'NoControlPanel' -Value 1
    if ($level -in @('Basic', 'Strict')) {
        Set-DwordRegistryValue -Path $explorerPolicy -Name 'NoViewContextMenu' -Value 1
    }
    if ($level -eq 'Strict') {
        Set-DwordRegistryValue -Path (Get-TargetUserSystemPolicyPath -Sid $Audit.TargetUserSid) -Name 'DisableCMD' -Value 1
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
    $response = Invoke-CimMethod -Namespace 'root\standardcimv2\embedded' -ClassName 'WESL_UserSetting' -MethodName 'SetCustomShell' -Arguments @{
        Sid = $Audit.TargetUserSid
        Shell = $shell
        DefaultAction = $defaultAction
    }
    Assert-CimSuccess -Response $response
    Write-SmartThinClientLog -Message "Applied Shell Launcher for SID '$($Audit.TargetUserSid)'."
}

function Invoke-ThinClientApply {
    param([Parameter(Mandatory = $true)]$Config)

    Assert-ConfigValues -Config $Config
    Assert-ApplyAllowed -Config $Config
    $audit = Invoke-ThinClientAudit -Config $Config -RequestedProfile $Profile -RequestedAction 'Preview'
    $effectiveProfile = [string]$audit.EffectiveProfile
    $Config = Merge-ProfileConfig -Config $Config -EffectiveProfile $effectiveProfile
    $targetUser = Get-TargetUserState -Config $Config
    Assert-ApplyAllowed -Config $Config
    Assert-ApplyPreflight -Config $Config -Audit $audit -TargetUser $targetUser

    $custom = (Get-ConfigValue -Config $Config -Name ShellMode -Default Launcher) -eq 'CustomShell'
    $reportRoot=$script:OutputRoot;$oldRollbackRoot=$script:RollbackRoot
    try {
        if($custom){$script:OutputRoot=Initialize-ThinClientMachineRoot -Config $Config;$script:RollbackRoot=Join-Path $script:OutputRoot 'Rollback'}
        $launcherCandidate=Join-Path $script:OutputRoot 'Launcher\SmartThinClient-LaunchWorkspace.ps1'
        if($custom -and (Test-Path -LiteralPath $launcherCandidate) -and ((Get-Item -LiteralPath $launcherCandidate -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Refusing a redirected machine launcher.'}
        $rollbackState = New-RollbackState -Config $Config -Audit $audit -EffectiveProfile $effectiveProfile
        $rollbackFile = Save-RollbackState -State $rollbackState
        [void](New-DedicatedUserIfNeeded -Config $Config -TargetUser $targetUser)
        $audit = Invoke-ThinClientAudit -Config $Config -RequestedProfile $effectiveProfile -RequestedAction 'Preview'
        $rollbackState.TargetUserSid = $audit.TargetUserSid
        [void](Save-RollbackState -State $rollbackState)
        $launcherPath = New-LauncherScript -Config $Config -EffectiveProfile $effectiveProfile -Audit $audit
        $launcherCommand = Get-LauncherCommandLine -Machine:$custom
        if($custom) {
            Protect-ThinClientMachineLauncher -Path $launcherPath
            $desired=[pscustomobject]@{Scope='LocalPolicyAllUsers';ValueExists=$true;ValueKind='String';Value=$launcherCommand}
            Set-ThinClientCustomShellPolicy -Desired $desired -Expected $rollbackState.CustomShellPolicy
            $previousRun=@($rollbackState.Registry|Where-Object {$_.Path -eq 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' -and $_.Name -eq 'SmartThinClientShell'})
            if($previousRun.Count -and $previousRun[0].ValueExists){Remove-ItemProperty -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' -Name SmartThinClientShell -ErrorAction Stop}
        }
        else { Apply-AutoLaunch -Config $Config -CommandLine $launcherCommand }
        Apply-ShellLimitations -Config $Config -Audit $audit
        Apply-AssignedAccess -Config $Config -Audit $audit
        Apply-ShellLauncher -Config $Config -Audit $audit -LauncherCommand $launcherCommand

        $result = [pscustomobject]([ordered]@{
        AppName = $script:AppName
        RunId = $script:RunId
        ComputerName = $env:COMPUTERNAME
        RequestedAction = 'Apply'
        EffectiveProfile = $effectiveProfile
        Status = 'Applied'
        ShellMode = [string](Get-ConfigValue -Config $Config -Name ShellMode -Default Launcher)
        ShellScope = if($custom){'AllUsers'}else{'ConfiguredMode'}
        LauncherPath = $launcherPath
        LauncherCommand = $launcherCommand
        RollbackPath = $rollbackFile
        OutputRoot = $script:OutputRoot
        ReportRoot = $reportRoot
        LogPath = $script:LogPath
        EvidenceJsonPath = $script:EvidenceJsonPath
        EvidenceCsvPath = $script:EvidenceCsvPath
    })
        Export-Evidence -Evidence $result -Config $Config
        return $result
    }
    finally {$script:OutputRoot=$reportRoot;$script:RollbackRoot=$oldRollbackRoot}
}

function Invoke-ThinClientLaunchOnly {
    param([Parameter(Mandatory = $true)]$Config)

    $audit = Invoke-ThinClientAudit -Config $Config -RequestedProfile $Profile -RequestedAction 'Preview'
    $effectiveProfile = [string]$audit.EffectiveProfile
    $Config = Merge-ProfileConfig -Config $Config -EffectiveProfile $effectiveProfile
    $installedRoot = $script:OutputRoot
    try {
        $script:OutputRoot = Join-Path $installedRoot ('LaunchOnly\' + [guid]::NewGuid().ToString('N'))
        $launchConfig=Copy-Config -Config $Config
        if((Get-ConfigValue -Config $launchConfig -Name ShellMode -Default Launcher) -eq 'CustomShell'){$launchConfig.ShellMode='Launcher'}
        $launcherPath = New-LauncherScript -Config $launchConfig -EffectiveProfile $effectiveProfile -Audit $audit
    }
    finally { $script:OutputRoot = $installedRoot }
    $launcherCommand = ('powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "{0}" -Profile "{1}"' -f $launcherPath, $effectiveProfile)

    Start-Process -FilePath 'powershell.exe' -WindowStyle Hidden -ArgumentList ('-NoProfile -STA -ExecutionPolicy Bypass -File "{0}" -Profile "{1}"' -f $launcherPath, $effectiveProfile) | Out-Null

    $result = [pscustomobject]([ordered]@{
        AppName = $script:AppName
        RunId = $script:RunId
        ComputerName = $env:COMPUTERNAME
        UserName = [Environment]::UserName
        IsAdministrator = Test-IsAdministrator
        RequestedAction = 'Launch'
        RequestedProfile = $Profile
        EffectiveProfile = $effectiveProfile
        Status = 'Launched'
        LauncherPath = $launcherPath
        LauncherCommand = $launcherCommand
        OutputRoot = $script:OutputRoot
        LogPath = $script:LogPath
        EvidenceJsonPath = $script:EvidenceJsonPath
        EvidenceCsvPath = $script:EvidenceCsvPath
    })
    Export-Evidence -Evidence $result -Config $Config
    Write-SmartThinClientLog -Message "Launch-only thin client opened. Launcher=$launcherPath"
    return $result
}

function Resolve-RollbackPath {
    param($Config=([ordered]@{}))
    if (-not [string]::IsNullOrWhiteSpace($RollbackPath)) { return $RollbackPath }
    $root=if((Get-ConfigValue -Config $Config -Name ShellMode -Default Launcher) -eq 'CustomShell'){Join-Path (Get-ThinClientMachineRoot -Config $Config) 'Rollback'}else{$script:RollbackRoot}
    $latest = Get-ChildItem -LiteralPath $root -Filter 'SmartThinClient_Rollback_*.json' -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if ($latest) { return $latest.FullName }
    throw 'No rollback file found. Use -RollbackPath to provide one.'
}

function Invoke-ThinClientRestore {
    param([Parameter(Mandatory = $true)]$Config)

    Assert-RestoreAllowed -Config $Config
    $path = Resolve-RollbackPath -Config $Config
    $rollback = Read-JsonFile -Path $path
    if ($null -eq $rollback) { throw "Rollback file is empty or invalid: $path" }
    Assert-RollbackState -State $rollback -Config $Config

    if($rollback.SchemaVersion -eq 3) {
        $current=Get-ThinClientCustomShellPolicy
        $unchangedOriginal=$current.ValueExists -eq $rollback.CustomShellPolicy.ValueExists -and (-not $current.ValueExists -or ($current.ValueKind -eq $rollback.CustomShellPolicy.ValueKind -and $current.Value -ceq $rollback.CustomShellPolicy.Value))
        if($current.ValueExists -and $current.Value -ceq $rollback.ManagedShellCommand){Set-ThinClientCustomShellPolicy -Desired $rollback.CustomShellPolicy -Expected $current}
        elseif(-not $unchangedOriginal){throw 'Custom shell policy no longer matches this rollback; no state restored.'}
    }

    foreach ($snapshot in @($rollback.Registry)) {
        Restore-RegistryValue -Snapshot $snapshot
    }

    $rollbackConfig = Convert-ObjectToHashtable -InputObject $rollback.Config
    if ([bool](Get-ConfigValue -Config $rollbackConfig -Name 'EnableAssignedAccess' -Default $false)) {
        Clear-AssignedAccess -ErrorAction Stop | Out-Null
    }
    if ([bool](Get-ConfigValue -Config $rollbackConfig -Name 'EnableShellLauncher' -Default $false) -and $rollback.TargetUserSid) {
        $response = Invoke-CimMethod -Namespace 'root\standardcimv2\embedded' -ClassName 'WESL_UserSetting' -MethodName 'RemoveCustomShell' -Arguments @{ Sid = [string]$rollback.TargetUserSid } -ErrorAction Stop
        Assert-CimSuccess -Response $response
    }
    if ($rollback.LauncherExisted) {
        $parent = Split-Path $rollback.LauncherPath -Parent
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
        [IO.File]::WriteAllBytes($rollback.LauncherPath, [Convert]::FromBase64String($rollback.LauncherContent))
        if($rollback.SchemaVersion -eq 3){Protect-ThinClientMachineLauncher -Path $rollback.LauncherPath}
    }
    elseif (Test-Path -LiteralPath $rollback.LauncherPath) {
        Remove-Item -LiteralPath $rollback.LauncherPath -Force -ErrorAction Stop
    }

    $result = [pscustomobject]([ordered]@{
        AppName = $script:AppName
        RunId = $script:RunId
        ComputerName = $env:COMPUTERNAME
        RequestedAction = 'Restore'
        Status = 'Restored'
        RollbackPath = $path
        OutputRoot = if($rollback.SchemaVersion -eq 3){$rollback.OutputRoot}else{$script:OutputRoot}
        LogPath = $script:LogPath
        EvidenceJsonPath = $script:EvidenceJsonPath
        EvidenceCsvPath = $script:EvidenceCsvPath
    })
    Export-Evidence -Evidence $result -Config $Config
    Write-SmartThinClientLog -Message "Restore completed from rollback: $path"
    return $result
}

function Test-ThinClientRestoreValueEqual {
    param($Left,$Right)
    return $Left.ValueExists -eq $Right.ValueExists -and (-not $Left.ValueExists -or ($Left.ValueKind -ceq $Right.ValueKind -and $Left.Value -ceq $Right.Value))
}

function Get-ThinClientWindowsRestoreState {
    [pscustomobject][ordered]@{
        Policy=Get-ThinClientCustomShellPolicy
        Run=Get-RegistryValueSnapshot -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' -Name 'SmartThinClientShell'
        Winlogon=Get-RegistryValueSnapshot -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon' -Name 'Shell'
    }
}

function Get-ThinClientLegacyStartupRoot {
    param([string]$Command)
    if ($Command -notmatch ' -File "([^"]+)"$') { throw 'The recorded startup is not a recognized ThinClient launcher. Use advanced restore for manual review.' }
    $path=$Matches[1]
    if (-not [IO.Path]::IsPathRooted($path) -or $path.StartsWith('\\') -or [IO.Path]::GetFileName($path) -ine 'SmartThinClient-LaunchWorkspace.ps1') { throw 'The recorded launcher is not a local ThinClient file.' }
    $root=[IO.Path]::GetFullPath((Split-Path (Split-Path $path -Parent) -Parent)).TrimEnd('\')
    if ($root -eq [IO.Path]::GetPathRoot($root).TrimEnd('\') -or $Command -cne (Get-LauncherCommandLine -Root $root)) { throw 'The recorded startup command does not match the managed ThinClient launcher.' }
    return $root
}

function Get-ThinClientWindowsRestoreFingerprint {
    param($Plan)
    $value=[ordered]@{ComputerName=$Plan.ComputerName;InitialState=$Plan.InitialState;Launchers=@($Plan.Launchers);Steps=@($Plan.Steps|Select-Object Path,SHA256,OutputRoot,SchemaVersion,CreatedUtc,ExpectedPolicy,ExpectedRun)}
    $bytes=[Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $value -Depth 16 -Compress))
    $sha=[Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($sha.ComputeHash($bytes))).Replace('-','') } finally { $sha.Dispose() }
}

function Get-ThinClientWindowsRestorePlan {
    param([Parameter(Mandatory=$true)]$Config)
    if((Get-ConfigValue -Config $Config -Name ShellMode -Default Launcher) -notin @('Launcher','CustomShell') -or (Get-ConfigValue -Config $Config -Name EnableAssignedAccess -Default $false) -or (Get-ConfigValue -Config $Config -Name EnableShellLauncher -Default $false)){throw 'This kiosk configuration requires advanced restore of its reviewed backup.'}
    $initial=Get-ThinClientWindowsRestoreState
    if (-not $initial.Winlogon.ValueExists -or $initial.Winlogon.Value -ine 'explorer.exe') { throw 'Windows Explorer is not the recorded Windows logon shell. Use advanced restore for manual review; no changes were made.' }
    $machine=Get-ThinClientMachineRoot -Config $Config
    $managed=Get-LauncherCommandLine -Root $machine -Machine
    $policy=$initial.Policy;$run=$initial.Run;$limit=[long]::MaxValue
    $cache=@{};$seen=@{};$launchers=[Collections.Generic.List[object]]::new();$steps=[Collections.Generic.List[object]]::new()
    while ($true) {
        # Only our exact custom-shell command or recorded own startup can select a history root.
        if ($run.ValueExists) { if($run.ValueKind -ne 'String'){throw 'ThinClient startup has an unexpected registry type.'};[void](Get-ThinClientLegacyStartupRoot -Command $run.Value) }
        if ($policy.ValueExists -and $policy.Value -ine 'explorer.exe') {
            if ($policy.Value -cne $managed) { throw 'Another custom shell is configured. Automatic Windows desktop restore will not replace it.' }
            $root=$machine;$schema=3
        } elseif ($run.ValueExists) { $root=Get-ThinClientLegacyStartupRoot -Command $run.Value;$schema=2 }
        else { break }
        if ($steps.Count -ge 64) { throw 'Rollback history exceeds the automatic restore limit. Use advanced restore.' }
        if (-not $cache.ContainsKey($root)) {
            $folder=Join-Path $root 'Rollback';$launcher=Join-Path $root 'Launcher\SmartThinClient-LaunchWorkspace.ps1'
            foreach($p in @($root,$folder,(Split-Path $launcher -Parent),$launcher)) { if((Test-Path -LiteralPath $p) -and ((Get-Item -LiteralPath $p -Force).Attributes -band [IO.FileAttributes]::ReparsePoint)){throw 'Automatic restore refuses redirected history or launcher paths.'} }
            $exists=Test-Path -LiteralPath $launcher -PathType Leaf
            $launchers.Add([pscustomobject][ordered]@{Path=$launcher;Exists=[bool]$exists;SHA256=$(if($exists){(Get-FileHash -LiteralPath $launcher).Hash}else{''})})
            $items=@(Get-ChildItem -LiteralPath $folder -Filter ('SmartThinClient_Rollback_'+$env:COMPUTERNAME+'_*.json') -File -ErrorAction SilentlyContinue)
            $cache[$root]=@(foreach($file in $items) {
                if($file.Attributes -band [IO.FileAttributes]::ReparsePoint){throw 'Automatic restore refuses redirected rollback files.'}
                $beforeHash=(Get-FileHash -LiteralPath $file.FullName).Hash
                $state=Read-JsonFile -Path $file.FullName
                $priorRoot=$script:OutputRoot
                try {$script:OutputRoot=$root;Assert-RollbackState -State $state -Config $Config} finally {$script:OutputRoot=$priorRoot}
                if($state.SchemaVersion -eq 2 -and $state.Config.ShellMode -ne 'Launcher'){throw 'This history requires advanced kiosk restore; no changes were made.'}
                $snapshotConfig=Convert-ObjectToHashtable -InputObject $state.Config
                if((Get-ConfigValue -Config $snapshotConfig -Name EnableAssignedAccess -Default $false) -or (Get-ConfigValue -Config $snapshotConfig -Name EnableShellLauncher -Default $false)){throw 'This backup includes kiosk changes and requires advanced restore.'}
                $date=[DateTimeOffset]::Parse([string]$state.CreatedUtc,[Globalization.CultureInfo]::InvariantCulture)
                if((Get-FileHash -LiteralPath $file.FullName).Hash -cne $beforeHash){throw 'A rollback changed while preparing the review. Try again.'}
                [pscustomobject]@{Path=$file.FullName;SHA256=$beforeHash;State=$state;Ticks=$date.UtcDateTime.Ticks}
            })
        }
        $eligible=@($cache[$root]|Where-Object {$_.Ticks -lt $limit}|Sort-Object Ticks -Descending)
        if(-not $eligible.Count){throw 'A required earlier ThinClient backup is missing. Automatic restore cannot safely reach the Windows desktop; no changes were made.'}
        if($eligible.Count -gt 1 -and $eligible[0].Ticks -eq $eligible[1].Ticks){throw 'Rollback history is ambiguous. Use advanced restore; no changes were made.'}
        $item=$eligible[0];$state=$item.State
        if($seen.ContainsKey($item.Path) -or $state.SchemaVersion -ne $schema){throw 'The rollback sequence does not match the configured shell.'}
        $seen[$item.Path]=$true;$limit=$item.Ticks
        $steps.Add([pscustomobject][ordered]@{Path=$item.Path;SHA256=$item.SHA256;OutputRoot=$root;SchemaVersion=$schema;CreatedUtc=[string]$state.CreatedUtc;ExpectedPolicy=$policy;ExpectedRun=$run})
        if($schema -eq 3){$policy=$state.CustomShellPolicy}
        $startup=@($state.Registry|Where-Object {$_.Path -ieq 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Run' -and $_.Name -eq 'SmartThinClientShell'})
        if($startup.Count -gt 1){throw 'Rollback contains duplicate startup snapshots.'}
        if($startup.Count){$run=$startup[0]}
    }
    $plan=[pscustomobject][ordered]@{Purpose='WindowsDesktop';ComputerName=$env:COMPUTERNAME;InitialState=$initial;Launchers=$launchers.ToArray();Steps=$steps.ToArray();AlreadyWindows=($steps.Count -eq 0);ReviewHash=''}
    $plan.ReviewHash=Get-ThinClientWindowsRestoreFingerprint -Plan $plan
    return $plan
}

function Convert-ThinClientWindowsRestorePlanToText {
    param($Plan)
    if($Plan.AlreadyWindows){return 'The Windows desktop (Explorer) is already configured and ThinClient automatic startup is disabled. No changes are needed.'}
    $scope=if(@($Plan.Steps|Where-Object SchemaVersion -eq 3).Count){'All users on this device'}else{'ThinClient automatic startup on this device'}
    return (@('Action: Restore the Windows desktop (Explorer)',('Scope: '+$scope),'',('Previous ThinClient configurations to undo: '+$Plan.Steps.Count),'Remove the recorded ThinClient custom shell and automatic startup.','Restore the Windows restrictions recorded in those backups.','Keep user accounts, application files, configuration and backup history.','','Your current session stays open. Sign out and sign in again after completion to use Windows Explorer.','Domain policy can override local policy. This is not a Windows factory reset.') -join [Environment]::NewLine)
}

function Invoke-ThinClientWindowsRestore {
    param([Parameter(Mandatory=$true)]$Config,[Parameter(Mandatory=$true)]$Plan)
    Assert-RestoreAllowed -Config $Config
    if($Plan.Purpose -ne 'WindowsDesktop' -or $Plan.ComputerName -cne $env:COMPUTERNAME -or $Plan.ReviewHash -cne (Get-ThinClientWindowsRestoreFingerprint -Plan $Plan)){throw 'The Windows desktop restore review is invalid. No changes were made.'}
    $fresh=Get-ThinClientWindowsRestorePlan -Config $Config
    if($fresh.ReviewHash -cne $Plan.ReviewHash){throw 'The shell state or backup history changed after review. Review Restore again; no changes were made.'}
    if($fresh.AlreadyWindows){throw 'Windows desktop is already configured. No restore was run.'}
    $locks=[Collections.Generic.List[IO.FileStream]]::new();$completed=[Collections.Generic.List[object]]::new()
    $oldRoot=$script:OutputRoot;$oldRollback=$script:RollbackRoot;$oldJson=$script:EvidenceJsonPath;$oldCsv=$script:EvidenceCsvPath;$oldPath=$RollbackPath
    try {
        # Lock and hash every reviewed file before the first write. Existing single-backup engine remains authoritative.
        foreach($step in $fresh.Steps){
            $stream=[IO.File]::Open($step.Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::Read);$locks.Add($stream)
            $sha=[Security.Cryptography.SHA256]::Create();try{$hash=([BitConverter]::ToString($sha.ComputeHash($stream))).Replace('-','')}finally{$sha.Dispose()}
            if($hash -cne $step.SHA256){throw 'A rollback changed after review. No restore steps were run.'}
        }
        foreach($step in $fresh.Steps){
            $now=Get-ThinClientWindowsRestoreState
            if(-not(Test-ThinClientRestoreValueEqual $now.Policy $step.ExpectedPolicy) -or -not(Test-ThinClientRestoreValueEqual $now.Run $step.ExpectedRun) -or -not(Test-ThinClientRestoreValueEqual $now.Winlogon $fresh.InitialState.Winlogon)){throw 'Windows shell or startup changed during restore; remaining steps were stopped.'}
            $script:OutputRoot=$step.OutputRoot;$script:RollbackRoot=Join-Path $step.OutputRoot 'Rollback';$RollbackPath=$step.Path
            $index=$completed.Count+1
            $script:EvidenceJsonPath=Join-Path (Split-Path $oldJson) (([IO.Path]::GetFileNameWithoutExtension($oldJson))+'_step_'+$index+'.json')
            $script:EvidenceCsvPath=Join-Path (Split-Path $oldCsv) (([IO.Path]::GetFileNameWithoutExtension($oldCsv))+'_step_'+$index+'.csv')
            $result=Invoke-ThinClientRestore -Config $Config
            if($result.Status -ne 'Restored'){throw 'A restore step did not complete successfully.'}
            $completed.Add([pscustomobject]@{Path=$step.Path;SHA256=$step.SHA256;EvidenceJsonPath=$script:EvidenceJsonPath})
        }
        $final=Get-ThinClientWindowsRestoreState
        if(($final.Policy.ValueExists -and $final.Policy.Value -ine 'explorer.exe') -or $final.Run.ValueExists -or $final.Winlogon.Value -ine 'explorer.exe'){throw 'Windows desktop restore did not reach the reviewed final state.'}
    } catch {
        $message='Windows desktop restore stopped after '+$completed.Count+' of '+$fresh.Steps.Count+' steps. '+$_.Exception.Message
        $script:EvidenceJsonPath=$oldJson;$script:EvidenceCsvPath=$oldCsv
        Export-Evidence -Evidence ([pscustomobject]@{Status='RestoreFailed';ComputerName=$env:COMPUTERNAME;RequestedAction='Restore';Message=$message;CompletedSteps=$completed.ToArray();PlannedSteps=$fresh.Steps.Count}) -Config $Config
        throw $message
    } finally {
        foreach($stream in $locks){$stream.Dispose()}
        $script:OutputRoot=$oldRoot;$script:RollbackRoot=$oldRollback;$script:EvidenceJsonPath=$oldJson;$script:EvidenceCsvPath=$oldCsv;$RollbackPath=$oldPath
    }
    $result=[pscustomobject]@{Status='Restored';RequestedAction='Restore';ComputerName=$env:COMPUTERNAME;WindowsDesktopRestored=$true;RestoredSnapshots=$completed.ToArray();OutputRoot=$oldRoot;LogPath=$script:LogPath;EvidenceJsonPath=$oldJson;EvidenceCsvPath=$oldCsv}
    Export-Evidence -Evidence $result -Config $Config
    return $result
}

function Get-ThinClientVisualPreviewTarget {
    param($Config, [ValidateSet('Citrix','AVD','WebOnly')][string]$Provider, [string]$EffectiveProfile = 'Auto')
    $urls = Get-WorkspaceWebUrls -Config $Config -EffectiveProfile $EffectiveProfile
    $url = [string]$urls[$Provider]
    $fallback = $false
    if ([string]::IsNullOrWhiteSpace($url) -and $Provider -ne 'WebOnly') {
        $url = [string]$urls.WebOnly
        $fallback = -not [string]::IsNullOrWhiteSpace($url)
    }
    $uri = $null
    $valid = -not [string]::IsNullOrWhiteSpace($url) -and [Uri]::TryCreate($url, [UriKind]::Absolute, [ref]$uri) -and $uri.Scheme -in @('http','https') -and [string]::IsNullOrEmpty($uri.UserInfo)
    return [pscustomobject]@{ Provider = $Provider; Url = $url; IsValid = $valid; UsesWebOnlyFallback = $fallback }
}

function Show-ThinClientVisualPreview {
    param([Parameter(Mandatory = $true)]$Config, [Parameter(Mandatory = $true)]$Audit, $Owner)
    Add-Type -AssemblyName PresentationFramework,PresentationCore,WindowsBase
    $effectiveProfile = [string]$Audit.EffectiveProfile
    $Config = Merge-ProfileConfig -Config $Config -EffectiveProfile $effectiveProfile
    [xml]$xaml = Get-ThinClientShellXaml
    $window = [Windows.Markup.XamlReader]::Load([Xml.XmlNodeReader]::new($xaml))
    [void]$window.Resources.MergedDictionaries.Add((New-ThinClientUiResources))
    if ($null -ne $Owner) { $window.Owner = $Owner }
    $window.Title = 'Smart ThinClient - Visual Preview'
    $window.WindowState = 'Normal'
    $window.Width = 1080; $window.Height = 720
    $window.MinWidth = 760; $window.MinHeight = 520
    $iconPath = Join-Path $script:ScriptRoot 'Assets\WorkplaceCloudHub.ico'
    if (Test-Path -LiteralPath $iconPath) { $window.Icon = [Windows.Media.Imaging.BitmapFrame]::Create([Uri]::new($iconPath,[UriKind]::Absolute)) }
    . (Join-Path $script:ScriptRoot 'SmartThinClient-WebView2.ps1')
    $browserHost = $window.FindName('WorkspaceBrowser')
    $browser = New-SmartThinClientWebView2 -PreviewOnly
    $browserHost.Content = $browser
    $statusText = $window.FindName('StatusText')
    $window.FindName('TitleText').Text = [string](Get-ConfigValue -Config $Config -Name 'WebShellTitle' -Default 'Smart ThinClient')
    $window.FindName('TitleText').ToolTip = $window.FindName('TitleText').Text
    $grid = $window.Content
    $row = [Windows.Controls.RowDefinition]::new(); $row.Height = [Windows.GridLength]::Auto
    $grid.RowDefinitions.Insert(1,$row)
    foreach ($child in @($grid.Children)) { $oldRow = [Windows.Controls.Grid]::GetRow($child); if ($oldRow -ge 1) { [Windows.Controls.Grid]::SetRow($child,$oldRow+1) } }
    $banner = [Windows.Controls.Border]::new(); $banner.Background = '#E6F4FF'; $banner.Padding = '12,8'
    $label = [Windows.Controls.TextBlock]::new(); $label.TextWrapping = 'Wrap'
    $label.Text = 'VISUAL PREVIEW - No shell changes are applied. Settings opens live Windows settings. Power actions, native clients and external browser launches are disabled.'
    $banner.Child = $label; [Windows.Controls.Grid]::SetRow($banner,1); [void]$grid.Children.Add($banner)
    $showProviders = [bool](Get-ConfigValue -Config $Config -Name 'WebShellShowProviderButtons' -Default $true)
    $providerVisibility = Get-ThinClientProviderVisibility -EffectiveProfile $effectiveProfile -ShowProviderButtons $showProviders
    foreach ($pair in @(@('CitrixButton','Citrix'),@('AvdButton','AVD'),@('WebOnlyButton','WebOnly'))) { if (-not $providerVisibility[$pair[1]]) { $window.FindName($pair[0]).Visibility = 'Collapsed' } }
    $visibility = @{ ExternalButton = 'WebShellAllowExternalBrowser'; SignOutButton = 'WebShellAllowPowerControls'; RestartButton = 'WebShellAllowPowerControls'; ShutdownButton = 'WebShellAllowPowerControls' }
    foreach ($name in $visibility.Keys) {
        $button = $window.FindName($name)
        $button.IsEnabled = $false
        $button.ToolTip = 'Unavailable in visual preview.'
        if (-not (Get-ConfigValue -Config $Config -Name $visibility[$name] -Default $true)) { $button.Visibility = 'Collapsed' }
        # These controls deliberately have no action handlers in preview.
    }
    $settingsButton = $window.FindName('SettingsButton')
    $adminButton = $window.FindName('AdminButton')
    if ((Get-ThinClientAdminTokenState).IsAdministratorAccount) { $adminButton.Visibility = 'Visible' }
    $adminButton.Add_Click({
        try {
            $role = Get-ThinClientAdminTokenState
            if (-not $role.IsAdministratorAccount) { throw 'Administrator account required.' }
            if ($role.IsElevated -and $Owner -and $Owner.Title -eq 'Smart ThinClient Control Panel') { $window.Close();return }
            $panelConfig = if([string]::IsNullOrWhiteSpace($ConfigPath)){Join-Path $script:ScriptRoot 'SmartThinClient-Shell.config.json'}else{$ConfigPath}
            Open-ThinClientAdminPanel -ScriptPath (Join-Path $script:ScriptRoot 'SmartThinClient-Shell.ps1') -ConfigPath $panelConfig -Profile $effectiveProfile
        } catch { $statusText.Text='Admin panel could not open: '+$_.Exception.Message }
    })
    $settingsButton.Style = $window.FindResource('AccentButton')
    if (-not (Get-ConfigValue -Config $Config -Name 'WebShellAllowLimitedSettings' -Default $true)) { $settingsButton.Visibility = 'Collapsed' }
    $settingsButton.ToolTip = 'Open the configured live Windows settings pages.'
    $settingsButton.Add_Click({
        try {
            $pages = @(Get-ConfigValue -Config $Config -Name 'WebShellAllowedSettingsPages' -Default @('ms-settings:network-status','ms-settings:display','ms-settings:sound','ms-settings:dateandtime'))
            $buttons = Get-ConfigValue -Config $Config -Name 'WebShellSettingsButtons' -Default $null
            $nativeSettings=(Get-ConfigValue -Config $Config -Name ShellMode -Default Launcher) -eq 'CustomShell' -or -not @(Get-Process explorer -ErrorAction SilentlyContinue|Where-Object SessionId -eq ([Diagnostics.Process]::GetCurrentProcess().SessionId)).Count
            Show-ThinClientLimitedSettings -SettingsPages $pages -SettingsButtons $buttons -Owner $window -PreviewMode -NativeMode:$nativeSettings -NativeSettingsHelper (Join-Path $script:ScriptRoot 'SmartThinClient-DeviceSettings.ps1')
        } catch { $statusText.Text = 'Settings could not open: ' + $_.Exception.Message }
    })
    $navigate = {
        param([string]$Provider)
        $target = Get-ThinClientVisualPreviewTarget -Config $Config -Provider $Provider -EffectiveProfile $effectiveProfile
        if (-not $target.IsValid) {
            $statusText.Text = 'Preview - no valid HTTP/HTTPS URL configured for ' + $Provider
            $safeProvider = [Net.WebUtility]::HtmlEncode($Provider)
            $browser.NavigateToString('<html><body style="font-family:Segoe UI;padding:40px"><h1>'+ $safeProvider +' workspace preview</h1><p>No valid web URL is configured for this profile. The shell layout is shown; no client was started.</p></body></html>')
            return
        }
        $statusText.Text = 'Preview - ' + $Provider + $(if($target.UsesWebOnlyFallback){' (WebOnly fallback)'}else{''}) + ' - ' + $target.Url
        try { $browser.Source = [Uri]$target.Url }
        catch { $statusText.Text = 'Preview navigation failed: ' + $_.Exception.Message }
    }
    $browser.Add_CoreWebView2InitializationCompleted({
        param($sender,$eventArgs)
        if (-not $eventArgs.IsSuccess) { $statusText.Text = 'WebView2 initialization failed: ' + $eventArgs.InitializationException.Message }
    })
    $window.FindName('CitrixButton').Add_Click({ & $navigate 'Citrix' })
    $window.FindName('AvdButton').Add_Click({ & $navigate 'AVD' })
    $window.FindName('WebOnlyButton').Add_Click({ & $navigate 'WebOnly' })
    $window.FindName('BackButton').Add_Click({ if ($browser.CanGoBack) { $browser.GoBack() } })
    $window.FindName('RefreshButton').Add_Click({ if ($null -ne $browser.Source) { $browser.Reload() } })
    $initialProvider = $effectiveProfile
    if ($initialProvider -eq 'Hybrid') { $initialProvider = [string](Get-ConfigValue -Config $Config -Name 'HybridDefaultProvider' -Default 'Citrix') }
    if ($initialProvider -notin @('Citrix','AVD','WebOnly')) { $initialProvider = 'WebOnly' }
    $window.Add_ContentRendered({ & $navigate $initialProvider })
    try { [void]$window.ShowDialog() }
    finally { $browser.Dispose() }
}


function Get-ThinClientPreviewPlan {
    param($Config, $Audit)
    $custom=(Get-ConfigValue -Config $Config -Name ShellMode -Default Launcher) -eq 'CustomShell'
    $launcherRoot=if($custom){Get-ThinClientMachineRoot -Config $Config}else{$script:OutputRoot}
    return [pscustomobject][ordered]@{
        ChangesApplied = $false
        EffectiveProfile = $Audit.EffectiveProfile
        TargetUser = $Audit.TargetUserName
        ShellMode = Get-ConfigValue -Config $Config -Name 'ShellMode' -Default 'Launcher'
        ShellScope = if($custom){'All users - replace Explorer'}else{'Configured mode'}
        AutoLaunchMode = if($custom){'Custom User Interface policy (next sign-in)'}else{Get-ConfigValue -Config $Config -Name 'AutoLaunchMode' -Default 'RunKey'}
        ShellRestrictions = if (Get-ConfigValue -Config $Config -Name 'EnableShellLimitations' -Default $false) { Get-ConfigValue -Config $Config -Name 'ShellRestrictionLevel' -Default 'None' } else { 'None' }
        AssignedAccess = Get-ConfigValue -Config $Config -Name 'EnableAssignedAccess' -Default $false
        ShellLauncher = Get-ConfigValue -Config $Config -Name 'EnableShellLauncher' -Default $false
        CreateDedicatedUser = Get-ConfigValue -Config $Config -Name 'CreateDedicatedLocalUser' -Default $false
        LauncherPath = Join-Path $launcherRoot 'Launcher\SmartThinClient-LaunchWorkspace.ps1'
        RollbackDirectory = if($custom){Join-Path $launcherRoot 'Rollback'}else{$script:RollbackRoot}
        ApplyEnabled = Get-ConfigValue -Config $Config -Name 'AllowApply' -Default $false
        Note = 'Plan only. Apply separately checks administrator rights, configuration, confirmation and prerequisites, then saves rollback before changes.'
    }
}

function Assert-ThinClientGuiAccess {
    param($Config, [ValidateSet('Apply','Restore')][string]$RequestedAction)
    Assert-ConfigValues -Config $Config
    if (-not (Get-ConfigValue -Config $Config -Name ('Allow' + $RequestedAction) -Default $false)) {
        throw "$RequestedAction is blocked because Allow$RequestedAction is false in configuration. No changes were made."
    }
    if (-not (Test-IsAdministrator)) { throw "$RequestedAction requires an elevated administrator session. Reopen the control panel as administrator; no operation was started." }
}

function Invoke-ThinClientGuiAction {
    param(
        [Parameter(Mandatory = $true)]$Config,
        [ValidateSet('Audit','Preview','Apply','Restore')][string]$RequestedAction,
        [ValidateSet('Auto','Citrix','AVD','WebOnly','Hybrid')][string]$SelectedProfile,
        [string]$Confirmation = '',
        [string]$SelectedRollbackPath = '',
        [string]$ExpectedRollbackHash = '',
        $WindowsRestorePlan = $null,
        [securestring]$Password
    )
    Assert-ConfigValues -Config $Config
    $Profile = $SelectedProfile
    $ConfirmApply = $Confirmation
    $ConfirmRestore = $Confirmation
    $RollbackPath = $SelectedRollbackPath
    $DedicatedUserPassword = $Password
    try {
        if ($RequestedAction -in @('Audit','Preview')) {
            return Invoke-ThinClientAudit -Config $Config -RequestedProfile $SelectedProfile -RequestedAction $RequestedAction
        }
        Assert-ThinClientGuiAccess -Config $Config -RequestedAction $RequestedAction
        $effective = Resolve-EffectiveProfile -RequestedProfile $SelectedProfile -Config $Config
        $effectiveConfig = Merge-ProfileConfig -Config $Config -EffectiveProfile $effective
        Assert-ThinClientGuiAccess -Config $effectiveConfig -RequestedAction $RequestedAction
        $defaultPhrase = if ($RequestedAction -eq 'Apply') { 'APPLY' } else { 'RESTORE' }
        $phrase = [string](Get-ConfigValue -Config $effectiveConfig -Name ($RequestedAction + 'ConfirmationPhrase') -Default $defaultPhrase)
        if ([string]::IsNullOrWhiteSpace($phrase) -or $Confirmation -cne $phrase) { throw "$RequestedAction requires the exact confirmation phrase. No changes were made." }
        if ($RequestedAction -eq 'Apply') { return Invoke-ThinClientApply -Config $Config }
        if ($null -ne $WindowsRestorePlan) { return Invoke-ThinClientWindowsRestore -Config $effectiveConfig -Plan $WindowsRestorePlan }
        if ([string]::IsNullOrWhiteSpace($RollbackPath) -or [string]::IsNullOrWhiteSpace($ExpectedRollbackHash)) { throw 'Select and review an explicit rollback file before Restore.' }
        if ((Get-FileHash -LiteralPath $RollbackPath -Algorithm SHA256).Hash -ne $ExpectedRollbackHash) { throw 'The rollback file changed after review. Select and review it again.' }
        return Invoke-ThinClientRestore -Config $effectiveConfig
    }
    finally { $DedicatedUserPassword = $null; $Password = $null }
}

function Start-ThinClientGuiRun {
    param([ValidateSet('Audit','Preview','Apply','Restore')][string]$RequestedAction)
    $script:RunId = (Get-Date -Format 'yyyyMMdd-HHmmssfff') + '-' + [guid]::NewGuid().ToString('N').Substring(0,8)
    $folder = Split-Path -Path $script:EvidenceJsonPath -Parent
    $stem = 'SmartThinClient_Shell_{0}_{1}_{2}' -f $RequestedAction,$env:COMPUTERNAME,$script:RunId
    $script:EvidenceJsonPath = Join-Path $folder ($stem+'.json')
    $script:EvidenceCsvPath = Join-Path $folder ($stem+'.csv')
    if (-not [string]::IsNullOrWhiteSpace($script:LogPath)) { $script:LogPath = Join-Path (Split-Path -Path $script:LogPath -Parent) ('SmartThinClient-Shell_'+$script:RunId+'.log') }
}

function Convert-ThinClientApplyPlanToText {
    param($Config,$Audit)
    $plan = Get-ThinClientPreviewPlan -Config $Config -Audit $Audit
    $lines = @(
        ('Profile: '+$plan.EffectiveProfile)
        ('Shell mode: '+$plan.ShellMode)
        ('Scope: '+$plan.ShellScope)
        ('Automatic startup: '+$plan.AutoLaunchMode)
        ('Shell restrictions: '+$plan.ShellRestrictions)
        ('Create a dedicated local user: '+$plan.CreateDedicatedUser)
        ''
        ('Workspace launcher: '+$plan.LauncherPath)
        ('Rollback folder: '+$plan.RollbackDirectory)
        ''
        'A rollback is saved before changes. The current Windows session is not closed automatically.'
    )
    $lines -join [Environment]::NewLine
}

function Get-ThinClientOperationFeedback {
    param($Result,$Config)
    $status = [string](Get-ResultValue -Result $Result -Name Status -Default '')
    if ($status -notin @('Applied','Restored')) { throw 'The operation did not return a completed Apply or Restore result.' }
    $applied = $status -eq 'Applied'
    $mode = [string](Get-ConfigValue -Config $Config -Name ShellMode -Default 'Launcher')
    $auto = [string](Get-ConfigValue -Config $Config -Name AutoLaunchMode -Default 'RunKey')
    $desktop = [bool](Get-ResultValue -Result $Result -Name WindowsDesktopRestored -Default $false)
    $title = if($applied){'Apply completed'}elseif($desktop){'Windows desktop restored'}else{'Restore completed'}
    $message = if($applied){'The reviewed workspace settings were applied successfully. A rollback backup was saved before the changes.'}elseif($desktop){'The reviewed ThinClient configurations were undone. Windows Explorer is configured and ThinClient automatic startup is disabled.'}else{'The settings and launcher recorded in the selected rollback were restored successfully.'}
    $next = if(-not $applied){'Sign out and sign in again to use the restored shell settings. This window does not close your session.'}elseif($mode -eq 'CustomShell'){'ThinClient is configured to replace Explorer for all users at their next sign-in. Sign out and sign in to activate it. The current desktop is not closed automatically; domain policy can override local policy.'}elseif($mode -eq 'Launcher' -and $auto -eq 'None'){'Automatic startup is disabled. Use Preview to inspect the workspace; Apply does not open it automatically.'}elseif($mode -eq 'Launcher'){'ThinClient will start at the next Windows sign-in. This Launcher mode keeps Explorer. Use Preview to inspect the workspace now.'}else{'Sign out and sign in again to use the configured shell. This window does not close your session.'}
    $details = [Collections.Generic.List[string]]::new()
    if($desktop){$details.Add('Previous ThinClient configurations restored: '+@($Result.RestoredSnapshots).Count)}
    foreach($entry in @(@('EffectiveProfile','Profile'),@('LauncherPath','Workspace launcher'),@('RollbackPath',$(if($applied){'Rollback backup'}else{'Rollback restored'})),@('EvidenceJsonPath','Report'),@('LogPath','Log'))) {
        $value = [string](Get-ResultValue -Result $Result -Name $entry[0] -Default '')
        if ($value) { $details.Add($entry[1]+': '+$value) }
    }
    $signOutRecommended = -not $applied -or $mode -ne 'Launcher' -or $auto -ne 'None'
    if($signOutRecommended){$next += ' Choose Sign out now to activate the changes, or Later to keep working. Save your open work first.'}
    [pscustomobject]@{Title=$title;Message=$message;NextStep=$next;SignOutRecommended=[bool]$signOutRecommended;Details=($details -join [Environment]::NewLine);OutputRoot=[string](Get-ResultValue -Result $Result -Name OutputRoot -Default '')}
}

function Confirm-ThinClientSignOut {
    param($Owner)
    Add-Type -AssemblyName PresentationFramework
    [xml]$xaml=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Sign out of Windows?" Width="460" SizeToContent="Height" ResizeMode="NoResize" WindowStartupLocation="CenterOwner" Background="#F5F8FB" FontFamily="Segoe UI">
 <StackPanel Margin="18"><Border Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="16"><StackPanel><TextBlock Text="Sign out now?" FontSize="22" FontWeight="SemiBold" Foreground="#1F2937"/><TextBlock Text="Save your open work before continuing. This signs out your current Windows session. Sign in again to use the updated shell settings." TextWrapping="Wrap" Foreground="#5F6B7A" Margin="0,10,0,0"/></StackPanel></Border><StackPanel Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0"><Button Name="CancelButton" Content="Cancel" IsCancel="True" IsDefault="True" MinWidth="100" Height="36" Margin="0,0,10,0"/><Button Name="SignOutButton" Content="Sign out" MinWidth="110" Height="36"/></StackPanel></StackPanel>
</Window>
'@
    $dialog=[Windows.Markup.XamlReader]::Load([Xml.XmlNodeReader]::new($xaml))
    [void]$dialog.Resources.MergedDictionaries.Add((New-ThinClientUiResources))
    $dialog.Owner=$Owner;$dialog.Icon=$Owner.Icon
    $dialog.FindName('CancelButton').Add_Click({$dialog.DialogResult=$false})
    $signOut=$dialog.FindName('SignOutButton');$signOut.Style=$dialog.FindResource('AccentButton');$signOut.Add_Click({$dialog.DialogResult=$true})
    return $dialog.ShowDialog() -eq $true
}

function Invoke-ThinClientSignOut {
    param($Owner,[switch]$Confirmed)
    if(-not $Confirmed -and -not(Confirm-ThinClientSignOut -Owner $Owner)){return $false}
    Write-SmartThinClientLog -Message 'User confirmed sign-out of the current Windows session.'
    $process=Start-Process -FilePath (Join-Path $env:SystemRoot 'System32\shutdown.exe') -ArgumentList '/l' -WindowStyle Hidden -Wait -PassThru -ErrorAction Stop
    if($process.ExitCode -ne 0){throw ('Windows sign-out could not start (exit '+$process.ExitCode+'). Your session remains open.')}
    return $true
}

function Show-ThinClientOperationFeedback {
    param($Owner,$Feedback)
    Add-Type -AssemblyName PresentationFramework
    [xml]$xaml=@'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Workspace changes completed" Width="700" Height="480" MinWidth="580" MinHeight="400" WindowStartupLocation="CenterOwner" Background="#F5F8FB" FontFamily="Segoe UI">
 <Grid Margin="18"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
  <Border Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="16"><StackPanel>
   <Border Background="#E8F5EE" CornerRadius="6" Padding="8,3" HorizontalAlignment="Left"><TextBlock Text="Completed" Foreground="#167044" FontWeight="SemiBold"/></Border>
   <TextBlock Name="ResultTitle" FontSize="24" FontWeight="SemiBold" Foreground="#1F2937" Margin="0,8,0,7"/><TextBlock Name="ResultMessage" TextWrapping="Wrap" Foreground="#334155"/>
  </StackPanel></Border>
  <Border Grid.Row="1" Background="#E6F4FF" BorderBrush="#B9DDF7" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,12,0,12"><StackPanel><TextBlock Text="Next step" FontWeight="SemiBold" Foreground="#005A9E" Margin="0,0,0,5"/><TextBlock Name="NextStep" TextWrapping="Wrap" Foreground="#24364B"/></StackPanel></Border>
  <Border Grid.Row="2" Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="12"><TextBox Name="ResultDetails" IsReadOnly="True" BorderThickness="0" TextWrapping="Wrap" FontSize="11" VerticalScrollBarVisibility="Auto"/></Border>
  <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0"><Button Name="FilesButton" Content="Open files" MinWidth="110" Height="36" Margin="0,0,10,0"/><Button Name="DoneButton" Content="Done" IsCancel="True" IsDefault="True" MinWidth="110" Height="36"/><Button Name="SignOutButton" Content="Sign out now" Visibility="Collapsed" MinWidth="120" Height="36" Margin="10,0,0,0"/></StackPanel>
 </Grid>
</Window>
'@
    $dialog=[Windows.Markup.XamlReader]::Load([Xml.XmlNodeReader]::new($xaml))
    [void]$dialog.Resources.MergedDictionaries.Add((New-ThinClientUiResources))
    $dialog.Owner=$Owner;$dialog.Icon=$Owner.Icon;$dialog.Title=$Feedback.Title
    $dialog.FindName('ResultTitle').Text=$Feedback.Title;$dialog.FindName('ResultMessage').Text=$Feedback.Message
    $dialog.FindName('NextStep').Text=$Feedback.NextStep;$dialog.FindName('ResultDetails').Text=$Feedback.Details
    $done=$dialog.FindName('DoneButton');$done.Add_Click({$dialog.Close()})
    $signOut=$dialog.FindName('SignOutButton');$signOut.Style=$dialog.FindResource('AccentButton')
    if($Feedback.SignOutRecommended){$signOut.Visibility='Visible';$done.Content='Later'}else{$done.Style=$dialog.FindResource('AccentButton')}
    $signOut.Add_Click({
        try {if(Invoke-ThinClientSignOut -Owner $dialog -Confirmed){$dialog.Close()}}
        catch {$dialog.FindName('NextStep').Text='Sign-out failed: '+$_.Exception.Message;Write-SmartThinClientLog -Message ('Sign-out failed: '+$_.Exception.Message) -Level ERROR}
    })
    $files=$dialog.FindName('FilesButton');$files.IsEnabled=-not [string]::IsNullOrWhiteSpace($Feedback.OutputRoot)
    $files.Add_Click({if(Test-Path -LiteralPath $Feedback.OutputRoot -PathType Container){Start-Process -FilePath explorer.exe -ArgumentList ('"'+$Feedback.OutputRoot+'"')}})
    [void]$dialog.ShowDialog()
}

function Show-ThinClientConfirmation {
    param($Owner, [string]$RequestedAction, [string]$Summary, [string]$Phrase, [bool]$RequestPassword = $false, [switch]$WindowsDesktop)
    Add-Type -AssemblyName PresentationFramework
    [xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Confirm workspace change" Width="700" Height="600" MinWidth="580" MinHeight="480" WindowStartupLocation="CenterOwner" ResizeMode="CanResize" Background="#F5F8FB" FontFamily="Segoe UI">
 <Grid Margin="18"><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
 <Border Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="16" Margin="0,0,0,12"><StackPanel>
  <Border Background="#E6F4FF" CornerRadius="6" Padding="8,3" HorizontalAlignment="Left"><TextBlock Text="Review changes" Foreground="#005A9E" FontWeight="SemiBold" FontSize="11"/></Border>
  <TextBlock Name="HeadingText" FontSize="24" FontWeight="SemiBold" Foreground="#1F2937" Margin="0,8,0,6"/>
  <TextBlock Name="WarningText" TextWrapping="Wrap" Foreground="#5F6B7A"/>
 </StackPanel></Border>
 <Border Grid.Row="1" Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="14"><Grid><Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
  <TextBlock Text="Planned changes" FontWeight="SemiBold" Foreground="#1F2937" Margin="0,0,0,8"/>
  <TextBox Name="SummaryText" Grid.Row="1" IsReadOnly="True" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" BorderThickness="0" Background="White" Foreground="#334155" FontFamily="Segoe UI" FontSize="12"/>
 </Grid></Border>
 <Border Grid.Row="2" Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="14" Margin="0,12,0,0"><StackPanel>
  <TextBlock Name="PhraseLabel" TextWrapping="Wrap" FontWeight="SemiBold" Foreground="#1F2937"/>
  <Border BorderBrush="#CBD5E1" BorderThickness="1" CornerRadius="6" Padding="10,5" Margin="0,7,0,0"><TextBox Name="PhraseInput" BorderThickness="0" MinHeight="24" FontSize="15" VerticalContentAlignment="Center"/></Border>
  <TextBlock Text="The confirmation is case-sensitive." FontSize="11" Foreground="#5F6B7A" Margin="0,5,0,0"/>
  <StackPanel Name="PasswordPanel" Margin="0,12,0,0" Visibility="Collapsed"><TextBlock Text="Password for the new dedicated local user" TextWrapping="Wrap" Foreground="#1F2937"/><Border BorderBrush="#CBD5E1" BorderThickness="1" CornerRadius="6" Padding="10,5" Margin="0,6,0,0"><PasswordBox Name="UserPassword" BorderThickness="0" MinHeight="24"/></Border></StackPanel>
 </StackPanel></Border>
 <StackPanel Grid.Row="3" Orientation="Horizontal" HorizontalAlignment="Right" Margin="0,14,0,0"><Button Name="CancelButton" Content="Cancel" IsCancel="True" MinWidth="100" Height="36" Margin="0,0,10,0"/><Button Name="ConfirmButton" Content="Apply" IsDefault="True" IsEnabled="False" MinWidth="110" Height="36"/></StackPanel>
 </Grid>
</Window>
"@
    $dialog = [Windows.Markup.XamlReader]::Load([Xml.XmlNodeReader]::new($xaml))
    [void]$dialog.Resources.MergedDictionaries.Add((New-ThinClientUiResources))
    $dialog.Owner = $Owner
    $dialog.Icon = $Owner.Icon
    $dialog.Title = 'Confirm ' + $RequestedAction
    $dialog.FindName('HeadingText').Text = 'Confirm ' + $RequestedAction
    $dialog.FindName('WarningText').Text = if ($RequestedAction -eq 'Apply') { 'Apply changes local workspace settings. Review this plan; a rollback is saved before changes.' } else { 'Restore changes local workspace settings using the exact reviewed rollback file.' }
    $dialog.FindName('SummaryText').Text = $Summary
    $dialog.FindName('PhraseLabel').Text = 'Type exactly: ' + $Phrase
    $inputBox = $dialog.FindName('PhraseInput')
    $confirmButton = $dialog.FindName('ConfirmButton')
    $confirmButton.Content = $RequestedAction
    $confirmButton.Style = $dialog.FindResource('AccentButton')
    if($WindowsDesktop){
        $dialog.Title='Restaurer le bureau Windows'
        $dialog.FindName('HeadingText').Text='Restaurer le bureau Windows'
        $dialog.FindName('WarningText').Text='Return to Windows Explorer using the reviewed backup history. Your session stays open until you choose to sign out.'
        $confirmButton.Content='Restore Windows'
        $confirmButton.MinWidth=140
        $confirmButton.IsDefault=$false
        $dialog.FindName('CancelButton').IsDefault=$true
    }
    $passwordBox = $dialog.FindName('UserPassword')
    if ($RequestPassword) { $dialog.FindName('PasswordPanel').Visibility = 'Visible' }
    $inputBox.Add_TextChanged({ $confirmButton.IsEnabled = -not [string]::IsNullOrWhiteSpace($Phrase) -and $inputBox.Text -ceq $Phrase })
    $confirmButton.Add_Click({ if ($inputBox.Text -ceq $Phrase) { $dialog.DialogResult = $true } })
    try {
        if ($dialog.ShowDialog() -ne $true) { return $null }
        return [pscustomobject]@{ Phrase = $inputBox.Text; Password = if ($RequestPassword) { $passwordBox.SecurePassword } else { $null } }
    }
    finally { $passwordBox.Clear() }
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
        Title="Smart ThinClient Control Panel" Height="620" Width="960" MinHeight="540" MinWidth="720"
        Background="#F5F8FB" WindowStartupLocation="CenterScreen">
  <Grid Margin="14">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>
    <Border Grid.Row="0" Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="12">
      <Grid>
        <Grid.ColumnDefinitions>
          <ColumnDefinition Width="*"/>
          <ColumnDefinition Width="104"/>
        </Grid.ColumnDefinitions>
        <StackPanel>
          <Border Background="#E6F4FF" CornerRadius="8" Padding="10,4" HorizontalAlignment="Left">
            <TextBlock Text="Endpoint workspace" Foreground="#005A9E" FontWeight="SemiBold"/>
          </Border>
          <TextBlock Text="Smart ThinClient Control Panel" FontSize="24" FontWeight="SemiBold" Foreground="#1F2937" Margin="0,8,0,0"/>
          <TextBlock Text="Inspect, preview, apply, and restore Citrix, AVD, or web workspace settings." Foreground="#5F6B7A" FontSize="12" TextWrapping="Wrap" Margin="0,4,10,0"/>
        </StackPanel>
        <Border Grid.Column="1" Background="#F8FBFE" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="12">
          <Image Name="LogoImage" Height="78" Stretch="Uniform"/>
        </Border>
      </Grid>
    </Border>
    <Border Grid.Row="1" Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="10" Margin="0,10,0,10">
      <WrapPanel Orientation="Horizontal">
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="0,3,12,3">
          <TextBlock Text="Profile" Foreground="#5F6B7A" VerticalAlignment="Center" Margin="0,0,8,0"/>
          <ComboBox Name="ProfileCombo" Width="136" Height="34" SelectedIndex="0">
            <ComboBoxItem Content="Auto"/>
            <ComboBoxItem Content="Citrix"/>
            <ComboBoxItem Content="AVD"/>
            <ComboBoxItem Content="WebOnly"/>
            <ComboBoxItem Content="Hybrid"/>
          </ComboBox>
          <TextBlock Name="StatusText" Text="Ready" Foreground="#5F6B7A" VerticalAlignment="Center" Margin="12,0,0,0"/>
        </StackPanel>
        <WrapPanel Orientation="Horizontal" Margin="0,3,0,3">
          <Button Name="AuditButton" Content="Audit" Height="34" MinWidth="76" Padding="12,0" Margin="0,0,8,0"/>
          <Button Name="PreviewButton" Content="Preview" Height="34" MinWidth="76" Padding="12,0" Margin="0,0,8,0"/>
          <Button Name="ApplyButton" Content="Apply" Height="34" MinWidth="76" Padding="12,0" Margin="0,0,8,0"/>
          <Button Name="RestoreButton" Content="Restore" Height="34" MinWidth="76" Padding="12,0" Margin="0,0,8,0"/>
          <Button Name="OpenOutputButton" Content="Open files" ToolTip="Open reports, logs and rollback files" Height="34" MinWidth="106" Padding="12,0"/>
        </WrapPanel>
      </WrapPanel>
    </Border>
    <Border Grid.Row="2" Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="12">
      <Grid>
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
        <Border Name="AuditSummaryContainer" Visibility="Collapsed" Background="#F8FBFE" BorderBrush="#DDE7F0" BorderThickness="0,0,0,1" Padding="0,0,0,8" Margin="0,0,0,8">
          <StackPanel>
            <TextBlock Name="AuditSummaryTitle" FontSize="16" FontWeight="SemiBold" Margin="0,0,0,6"/>
            <WrapPanel Name="AuditSummaryCards"/>
            <TextBlock Name="AuditSummaryNote" Foreground="#5F6B7A" FontSize="11" TextWrapping="Wrap"/>
          </StackPanel>
        </Border>
      <TextBox Grid.Row="1" Name="OutputText" FontFamily="Consolas" FontSize="12" TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Auto" IsReadOnly="True" BorderThickness="0"/>
      </Grid>
    </Border>
    <Border Grid.Row="3" Background="White" BorderBrush="#DDE7F0" BorderThickness="1" CornerRadius="8" Padding="10" Margin="0,10,0,0">
      <Grid><Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <TextBlock Text="Preview does not apply shell changes. Settings opens live Windows settings. Apply/Restore remain guarded." TextWrapping="Wrap" FontSize="11" Foreground="#5F6B7A"/>
        <TextBlock Name="VersionText" Grid.Column="1" Foreground="#64748B" FontSize="11" VerticalAlignment="Center" Margin="16,0,0,0"/>
      </Grid>
    </Border>
  </Grid>
</Window>
'@

    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $window.FindName('VersionText').Text = 'Smart ThinClient v' + (Get-ThinClientVersion)
    [void]$window.Resources.MergedDictionaries.Add((New-ThinClientUiResources))
    $windowIconPath = Join-Path $script:ScriptRoot 'Assets\WorkplaceCloudHub.ico'
    if (Test-Path -LiteralPath $windowIconPath) {
        $window.Icon = [Windows.Media.Imaging.BitmapFrame]::Create([Uri]::new($windowIconPath, [UriKind]::Absolute))
    }
    $logoImage = $window.FindName('LogoImage')
    $profileCombo = $window.FindName('ProfileCombo')
    $statusText = $window.FindName('StatusText')
    $outputText = $window.FindName('OutputText')
    $auditButton = $window.FindName('AuditButton')
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

    $auditButton.Style = $window.FindResource('AccentButton')
    $applyButton = $window.FindName('ApplyButton')
    $restoreButton = $window.FindName('RestoreButton')
    $previewButton = $window.FindName('PreviewButton')
    foreach ($item in $profileCombo.Items) {
        if ([string]$item.Content -eq $Profile) { $profileCombo.SelectedItem = $item; break }
    }
    $actionButtons = @($auditButton, $previewButton, $applyButton, $restoreButton, $openOutputButton)
    $runAction = {
        param([string]$RequestedAction,[switch]$AdvancedRestore)
        $confirmation = $null
        foreach ($button in $actionButtons) { $button.IsEnabled = $false }
        try {
            Start-ThinClientGuiRun -RequestedAction $RequestedAction
            $selectedProfile = [string]$profileCombo.Text
            $statusText.Text = 'Running ' + $RequestedAction.ToLowerInvariant() + '...'
            $parameters = @{ Config = $Config; RequestedAction = $RequestedAction; SelectedProfile = $selectedProfile }
            if ($RequestedAction -in @('Apply','Restore')) {
                Assert-ThinClientGuiAccess -Config $Config -RequestedAction $RequestedAction
                $effectiveProfile = Resolve-EffectiveProfile -RequestedProfile $selectedProfile -Config $Config
                $effectiveConfig = Merge-ProfileConfig -Config $Config -EffectiveProfile $effectiveProfile
                Assert-ThinClientGuiAccess -Config $effectiveConfig -RequestedAction $RequestedAction
                $defaultPhrase = if ($RequestedAction -eq 'Apply') { 'APPLY' } else { 'RESTORE' }
                $phrase = [string](Get-ConfigValue -Config $effectiveConfig -Name ($RequestedAction + 'ConfirmationPhrase') -Default $defaultPhrase)
                $requestPassword = $false
                if ($RequestedAction -eq 'Apply') {
                    $preview = Invoke-ThinClientGuiAction -Config $Config -RequestedAction Preview -SelectedProfile $selectedProfile
                    $summary = Convert-ThinClientApplyPlanToText -Config $effectiveConfig -Audit $preview
                    $outputText.Text = $summary
                    $requestPassword = [bool](Get-ConfigValue -Config $effectiveConfig -Name 'CreateDedicatedLocalUser' -Default $false)
                }
                elseif($AdvancedRestore) {
                    $picker = New-Object Microsoft.Win32.OpenFileDialog
                    $picker.Title = 'Select the rollback file to review and restore'
                    $picker.Filter = 'Rollback JSON (*.json)|*.json'
                    $picker.CheckFileExists = $true
                    $picker.InitialDirectory = if((Get-ConfigValue -Config $Config -Name ShellMode -Default Launcher) -eq 'CustomShell'){Join-Path (Get-ThinClientMachineRoot -Config $Config) 'Rollback'}else{$script:RollbackRoot}
                    if(-not(Test-Path -LiteralPath $picker.InitialDirectory -PathType Container)){$picker.InitialDirectory=$script:RollbackRoot}
                    if ($picker.ShowDialog($window) -ne $true) { $statusText.Text = 'Restore cancelled'; return }
                    $parameters.SelectedRollbackPath = $picker.FileName
                    $rollback = Read-JsonFile -Path $picker.FileName
                    if ($null -eq $rollback) { throw 'The selected rollback is empty or invalid.' }
                    Assert-RollbackState -State $rollback -Config $Config
                    $parameters.ExpectedRollbackHash = (Get-FileHash -LiteralPath $picker.FileName -Algorithm SHA256).Hash
                    $summary = 'Rollback: ' + $picker.FileName + [Environment]::NewLine + 'SHA256: ' + $parameters.ExpectedRollbackHash + [Environment]::NewLine + (($rollback | Select-Object ComputerName, TargetUserSid, OutputRoot, LauncherPath, LauncherExisted, CustomShellPolicy, ManagedShellCommand, Config, Registry) | ConvertTo-Json -Depth 10)
                }
                else {
                    $parameters.WindowsRestorePlan=Get-ThinClientWindowsRestorePlan -Config $effectiveConfig
                    $summary=Convert-ThinClientWindowsRestorePlanToText -Plan $parameters.WindowsRestorePlan
                    $outputText.Text=$summary
                    if($parameters.WindowsRestorePlan.AlreadyWindows){$statusText.Text='Windows desktop already configured';return}
                }
                $confirmation = Show-ThinClientConfirmation -Owner $window -RequestedAction $RequestedAction -Summary $summary -Phrase $phrase -RequestPassword $requestPassword -WindowsDesktop:($RequestedAction -eq 'Restore' -and -not $AdvancedRestore)
                if ($null -eq $confirmation) { $statusText.Text = $RequestedAction + ' cancelled'; return }
                $parameters.Confirmation = $confirmation.Phrase
                $parameters.Password = $confirmation.Password
            }
            $result = Invoke-ThinClientGuiAction @parameters
            if ($RequestedAction -eq 'Audit') { Set-ThinClientAuditSummary -Audit $result -Panel ($window.FindName('AuditSummaryCards')) -Title ($window.FindName('AuditSummaryTitle')) -Note ($window.FindName('AuditSummaryNote')) -Container ($window.FindName('AuditSummaryContainer')) }
            if ($RequestedAction -eq 'Preview') {
                $statusText.Text = 'Visual preview open'
                Show-ThinClientVisualPreview -Config $Config -Audit $result -Owner $window
                $outputText.Text = 'Visual preview closed. Profile: ' + $result.EffectiveProfile + '. No shell changes were applied. Audit and the change plan are saved in the output folder.'
                $statusText.Text = 'Visual preview closed'
                return
            }
            if ($RequestedAction -in @('Apply','Restore')) {
                $window.FindName('AuditSummaryContainer').Visibility = 'Collapsed'
                $feedback = Get-ThinClientOperationFeedback -Result $result -Config $Config
                $outputText.Text = $feedback.Title+[Environment]::NewLine+$feedback.Message+[Environment]::NewLine+[Environment]::NewLine+$feedback.NextStep+[Environment]::NewLine+[Environment]::NewLine+$feedback.Details
                $statusText.Text = $feedback.Title
                Show-ThinClientOperationFeedback -Owner $window -Feedback $feedback
                return
            }
            $outputText.Text = Convert-AuditToText -Audit $result
            $statusText.Text = $RequestedAction + ' completed'
        }
        catch {
            $window.FindName('AuditSummaryContainer').Visibility = 'Collapsed'
            if ($RequestedAction -eq 'Audit') {
                $window.FindName('AuditSummaryTitle').Text = 'Audit | Blocked'
                $window.FindName('AuditSummaryTitle').Foreground = '#B42318'
                $window.FindName('AuditSummaryCards').Children.Clear()
                $window.FindName('AuditSummaryNote').Text = 'Audit did not complete: ' + $_.Exception.Message
                $window.FindName('AuditSummaryContainer').Visibility = 'Visible'
            }
            $configurationBlock = $RequestedAction -in @('Apply','Restore') -and $_.Exception.Message -eq "$RequestedAction is blocked because Allow$RequestedAction is false in configuration. No changes were made."
            $statusText.Text = if ($configurationBlock) { $RequestedAction + ' blocked by configuration' } else { $RequestedAction + ' did not complete' }
            $outputText.Text = $_.Exception.Message
            Write-SmartThinClientLog -Level 'ERROR' -Message $_.Exception.Message
        }
        finally {
            if ($null -ne $confirmation -and $null -ne $confirmation.Password) { $confirmation.Password.Dispose() }
            $confirmation = $null
            foreach ($button in $actionButtons) { $button.IsEnabled = $true }
        }
    }
    foreach ($button in @($auditButton, $previewButton, $applyButton, $restoreButton)) {
        $button.Tag = [string]$button.Content
        $button.Add_Click({ & $runAction -RequestedAction ([string]$this.Tag) })
    }
    $restoreButton.ToolTip='Restore the Windows desktop. Right-click for advanced restore of a specific backup.'
    $restoreButton.ContextMenu=[Windows.Controls.ContextMenu]::new()
    $advancedRestore=[Windows.Controls.MenuItem]::new();$advancedRestore.Header='Restore a specific backup (advanced)';$advancedRestore.Tag=$runAction
    $advancedRestore.Add_Click({& $this.Tag -RequestedAction Restore -AdvancedRestore})
    [void]$restoreButton.ContextMenu.Items.Add($advancedRestore)


    $openOutputButton.Add_Click({
        $folder=$script:OutputRoot
        if((Get-ConfigValue -Config $Config -Name ShellMode -Default Launcher) -eq 'CustomShell'){$machine=Get-ThinClientMachineRoot -Config $Config;if(Test-Path -LiteralPath (Join-Path $machine 'SmartThinClient-Machine.json')){$folder=$machine}}
        if (Test-Path -LiteralPath $folder) { Start-Process -FilePath explorer.exe -ArgumentList ('"'+$folder+'"') }
    })

    $profileCombo.Add_SelectionChanged({
        $window.FindName('AuditSummaryContainer').Visibility = 'Collapsed'
        $statusText.Text = 'Profile changed - run Audit'
        if ($null -ne $profileCombo.SelectedItem) { $outputText.Text = 'Profile: ' + [string]$profileCombo.SelectedItem.Content + '. Run Audit or Preview to inspect this profile.' }
    })
    $outputText.Text = 'Ready. Audit inspects the device; Preview opens the workspace shell without applying Windows settings.'
    if ($splash) { Close-SmartThinClientGuiSplash -Splash $splash }
    [void]$window.ShowDialog()
}

$config = Resolve-Config
Assert-ConfigValues -Config $config
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
                Write-ThinClientCliSummary -Result $audit
            }
            else {
                Show-ThinClientGui -Config $config
            }
        }
        'Preview' {
            if ($Gui) { Show-ThinClientGui -Config $config }
            else {
                $audit = Invoke-ThinClientAudit -Config $config -RequestedProfile $Profile -RequestedAction 'Preview'
                Write-ThinClientCliSummary -Result $audit
            }
        }
        'Launch' {
            $result = Invoke-ThinClientLaunchOnly -Config $config
            Write-ThinClientCliSummary -Result $result
        }
        'Apply' {
            $result = Invoke-ThinClientApply -Config $config
            Write-ThinClientCliSummary -Result $result
        }
        'Restore' {
            $result = Invoke-ThinClientRestore -Config $config
            Write-ThinClientCliSummary -Result $result
        }
    }
    exit 0
}
catch {
    Write-SmartThinClientLog -Level 'ERROR' -Message $_.Exception.Message
    if ($Cli) {
        Write-CliSection -Title 'Final message'
        Write-Host 'FINAL RESULT: ' -NoNewline -ForegroundColor White
        Write-Host 'FAILED - ' -NoNewline -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Gray
        if ($script:LogPath) { Write-CliText ("Log: {0}" -f $script:LogPath) -Color DarkGray }
    }
    Write-Error $_.Exception.Message
    exit 1
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCtHhF4n7HGpdvP
# wJvKZg8jICysi51aGkPKm0LslF2KA6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILfytq+EONOqcK9c+FqXChQXO20F0BqIMhOZJV9hPfOwMA0GCSqG
# SIb3DQEBAQUABIIBgA3W1bvmd3fBPhLi1vP91i+2sZ26WdN3KiD2KRbTqZIQl/F5
# x9UiA08TL1tHGHl52gYwCJtuQt2/ktPF9IQdsaam1OH5T6s3xU2uwGrGVjSwJSF8
# 2B7FpPoeWVgFA7UENy3nMuGuexL6cUMqi7VgicHDZoNTIjPXzxIZVNPeJYMv3lHk
# QIxnXuWXSsW6TpjCde0LillpQ5c5F71VIIze7B3hw6Sv36MBOxyAYPjQ6mI5KspP
# L8zAm7IPQuMpCTRPBpMVNmD8/UHaZNuEFpNw2P/P43dBncVptE2Vn1gKzAA1pfh/
# b74S7SFwXllVH5t0LRoAwJO7u29x1hO/Dp1Xr75E/WeyKRHloQoY1dH3jrZb3R6X
# WUYZbPm2/GZ5ceCMz2p8DozlThXhczWlfVpsZHYv7hzE19NpT+5KSDV3NOAeza3w
# 294lvB9RoHvPFT6uAV++ZQGlzdvgPvjMNpIX8O+ShKS+SBN/TlSV+jRbH4OBcDqH
# eHGWOVk29BYL53IO0aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTMyMTQ1
# NTdaMC8GCSqGSIb3DQEJBDEiBCB25aZ36KpWsCL9oFt4t5GWvulYKE/Rqq7rt365
# YEtp0zANBgkqhkiG9w0BAQEFAASCAgAOEWYa/wi4od9gEj3m2AHcFZnDB71qMnvj
# uxG0+IjCYCPubUMbi/QfHArY2bO+OqKRcOsNVXmOxlvCJ4L/JTLgR0DM9wfcLquq
# qWKFI4DG+sIhjtySwTcOQ65zSYUJFM/4xWxEyutgzmEGtlEkFk2iF6ZpdtxuR1TH
# LT4nFjBwD8BNq2C+KZ+kC3HnS+wfFgMQanG5Wl+SvnbDXBdTj6HHhPRZMePjRhnm
# +nB9K18QMucWADEUDcK2ETEi/7EyW3rQSINRRS8WB8RZOA/gAEX5N0BkB8pkXSTg
# 1CA93cAgI5qdn1u7d/g23NU1Mo5bku/Lfo4TlV414XEc5h0xA0z3LXLJ8EreRHM/
# kIAlk3jyySHPComm3iso5BBdepOFCynK47OMqVyXZJi2aQsKuegBTZocl/ZoJn9q
# LwTEXb/WwKZKNuOuq1cMOvG2wjYD/NbFxjvQoQkbJ6aCeBRff2RpDBzKX3mpZKcg
# Ghv3mqUxXSXQP9wiUsrmzlfpcA1ZIcdeSj8EfCIMA3jcfYUSkI3EvLjMKMYiRD+J
# 2FkNooATQNGhZOUxSKs1LRXIaXHpv0B/UeFsEGEJAr5CBzH8mwbXF/3mMXD8MXVl
# jd+zCQmpZh58B1v8hbJRCSH77LftIs8G8rPR3FxIKPXkZPclCMay2llRi9pmt2Pb
# uwiQwRH0sg==
# SIG # End signature block
