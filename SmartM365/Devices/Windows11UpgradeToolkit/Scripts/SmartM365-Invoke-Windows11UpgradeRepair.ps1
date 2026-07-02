<#
.SYNOPSIS
    Diagnoses and optionally repairs Windows 10 devices that should move to Windows 11.

.DESCRIPTION
    Autonomous endpoint-side script for the SmartM365 Windows 11 Upgrade Toolkit.
    It is designed to run locally or as SYSTEM through PsExec/LOT orchestration.

    By default the script is diagnostic-only. Corrective actions require explicit switches.
    Setup-based upgrade requires -AllowSetupUpgrade and a validated setup source/cache.

.VERSION
    0.1.26

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [switch]$IgnoreRunGuard,
    [ValidateRange(0, 168)][int]$RunGuardHours = 12,

    [switch]$AuditOnly,
    [switch]$AllowPolicyRepair,
    [switch]$AllowWUReset,
    [switch]$AllowForceUpgrade,
    [switch]$AllowSetupUpgrade,
    [switch]$DirectSetupUpgrade,
    [switch]$AllowReboot,
    [switch]$AllowSetupCompletionRebootWhenNoUser,
    [switch]$SkipVirtualMachines,
    [switch]$AllowDiskCleanup,
    [switch]$AllowAdvancedDiskCleanup,
    [switch]$AllowDismComponentCleanup,

    [string]$SetupSourcePath,
    [string]$SetupSourceMapPath,
    [ValidateSet('LocalCache','Share','Auto')]
    [string]$SetupExecutionMode = 'LocalCache',
    [string]$SetupMediaId = 'Win11',
    [string]$SetupLanguage = 'MatchSystem',
    [ValidateSet('Enable','Disable','NoDrivers','NoLCU','NoDriversNoLCU')]
    [string]$SetupDynamicUpdate = 'Disable',
    [switch]$SkipSetupMediaPreCopy,
    [string]$SetupCacheRoot = 'C:\ProgramData\SmartM365\Windows11UpgradeToolkit\SetupMedia',
    [ValidateRange(0, 100)][int]$SetupSourceCandidateLimit = 5,
    [ValidateRange(0, 10000)][int]$SetupMediaCopyIpGapMilliseconds = 0,
    [ValidateRange(0, 86400)][int]$SetupMediaCopyJitterSeconds = 0,
    [ValidateRange(0, 500)][int]$SetupSourceConcurrencyLimit = 0,
    [ValidateRange(1, 1440)][int]$SetupSourceConcurrencyLeaseMinutes = 240,
    [string]$SetupSourceConcurrencyGateRoot,
    [ValidateRange(0, 500)][int]$SetupSubnetConcurrencyLimit = 0,
    [string]$SetupSubnetPrefixLength = 'Auto',
    [ValidateRange(1, 1440)][int]$SetupSubnetConcurrencyLeaseMinutes = 60,
    [string]$SetupSubnetConcurrencyGateRoot,
    [ValidateRange(10, 200)][int]$MinimumFreeDiskGB = 32,
    [ValidateRange(0, 365)][int]$DiskCleanupTempFileMinAgeDays = 1,
    [ValidateRange(0, 365)][int]$DiskCleanupLogRetentionDays = 14,
    [ValidateRange(0, 365)][int]$DiskCleanupUpgradeFolderMinAgeDays = 14,
    [ValidateRange(0, 86400)][int]$RebootDelaySeconds = 180,
    [ValidateRange(30, 3600)][int]$SetupProcessHeartbeatSeconds = 300,
    [ValidateRange(0, 1440)][int]$SetupProcessTimeoutMinutes = 0,

    [string]$DataRoot = 'C:\ProgramData\SmartM365\Windows11UpgradeToolkit'
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:ScriptName = 'SmartM365-Invoke-Windows11UpgradeRepair'
$script:ScriptVersion = '0.1.26'
$script:RunId = Get-Date -Format 'yyyyMMdd-HHmmss'
$script:ComputerName = $env:COMPUTERNAME
$script:LogDir = Join-Path $DataRoot 'Logs'
$script:OutputDir = Join-Path $DataRoot 'Output'
$script:SetupLogDir = Join-Path $script:LogDir 'SetupUpgrade'
$script:InputDir = Join-Path $DataRoot 'Input'
$script:LogPath = Join-Path $script:LogDir ("{0}_{1}_{2}.log" -f $script:ScriptName,$script:ComputerName,$script:RunId)
$script:CsvPath = Join-Path $script:OutputDir ("SmartM365_Windows11Upgrade_{0}_{1}.csv" -f $script:ComputerName,$script:RunId)
$script:LastRunPath = Join-Path $DataRoot 'LastRun.json'
$script:SetupMediaManifestFileName = 'SmartM365-SetupMediaManifest.sha256.csv'

function New-SmartDirectory {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
    }
}

function Write-SmartLog {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [ValidateSet('INFO','WARN','ERROR')][string]$Level = 'INFO'
    )

    New-SmartDirectory -Path $script:LogDir
    foreach ($line in ($Message -split "`r?`n")) {
        $entry = "{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$line
        Add-Content -LiteralPath $script:LogPath -Value $entry -Encoding UTF8
    }
}

function ConvertTo-LongLiteralPath {
    param([Parameter(Mandatory = $true)][string]$Path)

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    if ($fullPath.StartsWith('\\?\')) { return $fullPath }
    if ($fullPath.StartsWith('\\')) { return '\\?\UNC\' + $fullPath.Substring(2) }
    return '\\?\' + $fullPath
}

function Get-SmartFileSha256 {
    param([Parameter(Mandatory = $true)][string]$Path)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    $stream = $null
    try {
        $stream = [System.IO.File]::OpenRead((ConvertTo-LongLiteralPath -Path $Path))
        $hashBytes = $sha.ComputeHash($stream)
        return (($hashBytes | ForEach-Object { $_.ToString('X2') }) -join '')
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
        if ($null -ne $sha) { $sha.Dispose() }
    }
}
function Convert-SignedExitCodeToHex {
    param([Parameter(Mandatory = $true)][int]$ExitCode)

    $bytes = [BitConverter]::GetBytes($ExitCode)
    $unsigned = [BitConverter]::ToUInt32($bytes, 0)
    return ('0x{0:X8}' -f $unsigned)
}

function Get-SetupExitCodeInfo {
    param([Parameter(Mandatory = $true)][int]$ExitCode)

    $hex = Convert-SignedExitCodeToHex -ExitCode $ExitCode
    $knownMeanings = @{
        '0x00000000' = 'Success.'
        '0x00000BC2' = 'Success; reboot required.'
        '0x8007000B' = 'Bad image format. Windows Setup could not read a required image; validate or recopy setup media.'
        '0xC1900101' = 'Windows Setup rollback, often driver or firmware related. Confirm with setup logs or SetupDiag.'
        '0xC190010E' = 'Windows Setup requires EULA acceptance. Use /EULA accept for quiet or non-interactive Windows 11 Setup.'
        '0xC1900200' = 'Compatibility failure: device does not meet Windows Setup minimum requirements.'
        '0xC1900202' = 'Compatibility failure: device does not meet Windows Setup minimum requirements for this upgrade.'
        '0xC1900204' = 'Migration choice not available or invalid.'
        '0xC1900208' = 'Compatibility blocker, usually an incompatible app or driver.'
        '0xC190020E' = 'Insufficient free disk space for Windows Setup.'
    }

    if ($knownMeanings.ContainsKey($hex)) {
        $meaning = $knownMeanings[$hex]
    }
    elseif ($hex -like '0xC190010*') {
        $meaning = 'Windows Setup C190010x failure. Exact cause must be confirmed from setupact.log, setuperr.log, or SetupDiag.'
    }
    elseif ($hex -like '0xC19002*') {
        $meaning = 'Windows Setup compatibility or prerequisite failure. Exact cause must be confirmed from setup logs or SetupDiag.'
    }
    else {
        $meaning = 'Windows Setup returned an unmapped exit code. Check setup logs or SetupDiag.'
    }

    [pscustomobject]@{
        Decimal = $ExitCode
        Hex = $hex
        Meaning = $meaning
        CopyLogsPath = $script:SetupLogDir
        PantherLog = 'C:\$WINDOWS.~BT\Sources\Panther\setupact.log'
        PantherErrorLog = 'C:\$WINDOWS.~BT\Sources\Panther\setuperr.log'
        RollbackLogFolder = 'C:\$WINDOWS.~BT\Sources\Rollback'
    }
}

function Format-SetupExitCodeInfo {
    param([Parameter(Mandatory = $true)]$Info)

    return ("ExitCode={0}; Hex={1}; Meaning={2}; CopyLogs={3}; Panther={4}; PantherErrors={5}; RollbackLogs={6}" -f $Info.Decimal,$Info.Hex,$Info.Meaning,$Info.CopyLogsPath,$Info.PantherLog,$Info.PantherErrorLog,$Info.RollbackLogFolder)
}

function Get-RegistryValue {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Name
    )

    try {
        $item = Get-ItemProperty -LiteralPath $Path -Name $Name -ErrorAction Stop
        return $item.$Name
    }
    catch {
        return $null
    }
}

function Get-OsSummary {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $build = [int]$os.BuildNumber
    $major = if ($build -ge 22000) { 'Windows11' } elseif ($build -ge 10240) { 'Windows10' } else { 'Other' }
    [pscustomobject]@{
        Caption = [string]$os.Caption
        Version = [string]$os.Version
        BuildNumber = $build
        Architecture = [string]$os.OSArchitecture
        MajorFamily = $major
        LastBootUpTime = $os.LastBootUpTime
    }
}

function Get-ComputerSystemSummary {
    $system = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction Stop
    $manufacturer = [string]$system.Manufacturer
    $model = [string]$system.Model
    $hypervisorPresent = $false
    try { $hypervisorPresent = [bool]$system.HypervisorPresent } catch { $hypervisorPresent = $false }

    $signature = ("{0} {1}" -f $manufacturer,$model)
    $virtualPatterns = @(
        'Virtual Machine',
        'VMware',
        'VirtualBox',
        'KVM',
        'QEMU',
        'Xen',
        'HVM domU',
        'Parallels',
        'BHYVE',
        'OpenStack',
        'Google Compute Engine',
        'Amazon EC2'
    )

    $matchedPattern = ''
    foreach ($pattern in $virtualPatterns) {
        if ($signature -match [regex]::Escape($pattern)) {
            $matchedPattern = $pattern
            break
        }
    }

    $isVirtual = -not [string]::IsNullOrWhiteSpace($matchedPattern)
    $evidence = if ($matchedPattern) {
        "Manufacturer=$manufacturer; Model=$model; Pattern=$matchedPattern"
    }
    else {
        "Manufacturer=$manufacturer; Model=$model; HypervisorPresent=$hypervisorPresent"
    }

    [pscustomobject]@{
        Manufacturer = $manufacturer
        Model = $model
        HypervisorPresent = $hypervisorPresent
        IsVirtualMachine = $isVirtual
        Evidence = $evidence
    }
}

function Get-SystemDriveFreeGb {
    $driveLetter = ([System.IO.Path]::GetPathRoot($env:SystemRoot)).TrimEnd('\')
    $disk = Get-CimInstance -ClassName Win32_LogicalDisk -Filter ("DeviceID='{0}'" -f $driveLetter) -ErrorAction Stop
    if ($null -eq $disk.FreeSpace) { return 0 }
    return [math]::Round(([double]$disk.FreeSpace / 1GB), 2)
}

function Test-PendingReboot {
    $sources = New-Object System.Collections.ArrayList

    if (Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') {
        [void]$sources.Add('CBS:Component Based Servicing\RebootPending')
    }
    if (Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') {
        [void]$sources.Add('WindowsUpdate:Auto Update\RebootRequired')
    }
    $pendingFileRename = Get-RegistryValue -Path 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name 'PendingFileRenameOperations'
    if ($null -ne $pendingFileRename) {
        [void]$sources.Add('SessionManager:PendingFileRenameOperations')
    }

    [pscustomobject]@{
        IsPending = ($sources.Count -gt 0)
        Source = (@($sources.ToArray()) -join '; ')
    }
}

function Get-IntuneEnrollmentSummary {
    $enrollmentRoot = 'HKLM:\SOFTWARE\Microsoft\Enrollments'
    $strong = New-Object System.Collections.ArrayList
    $weak = New-Object System.Collections.ArrayList

    if (Test-Path -LiteralPath $enrollmentRoot) {
        foreach ($key in @(Get-ChildItem -LiteralPath $enrollmentRoot -ErrorAction SilentlyContinue)) {
            $provider = Get-RegistryValue -Path $key.PSPath -Name 'ProviderID'
            $discovery = Get-RegistryValue -Path $key.PSPath -Name 'DiscoveryServiceFullURL'
            if ([string]$provider -eq 'MS DM Server') {
                [void]$strong.Add(("ProviderID=MS DM Server ({0})" -f $key.PSChildName))
            }
            elseif ($discovery -match 'manage\.microsoft\.com|enrollment\.manage\.microsoft\.com') {
                [void]$strong.Add(("DiscoveryUrl={0}" -f $discovery))
            }
        }
    }

    if (Test-Path -LiteralPath 'HKLM:\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts') {
        [void]$weak.Add('OMADM account path exists')
    }
    if (Get-Service -Name 'IntuneManagementExtension' -ErrorAction SilentlyContinue) {
        [void]$weak.Add('Intune Management Extension service exists')
    }

    [pscustomobject]@{
        IsIntuneEnrolled = ($strong.Count -gt 0)
        StrongEvidence = (@($strong.ToArray()) -join '; ')
        WeakEvidence = (@($weak.ToArray()) -join '; ')
    }
}

function Get-WindowsUpdatePolicySummary {
    $wuPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    $auPath = Join-Path $wuPath 'AU'
    $mdmUpdatePath = 'HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Update'
    $policyStatePath = 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UpdatePolicy\PolicyState'

    $issues = New-Object System.Collections.ArrayList
    $wus = Get-RegistryValue -Path $wuPath -Name 'WUServer'
    $wuStatus = Get-RegistryValue -Path $wuPath -Name 'WUStatusServer'
    $useWUServer = Get-RegistryValue -Path $auPath -Name 'UseWUServer'
    $noAutoUpdate = Get-RegistryValue -Path $auPath -Name 'NoAutoUpdate'
    $doNotConnect = Get-RegistryValue -Path $wuPath -Name 'DoNotConnectToWindowsUpdateInternetLocations'
    $targetRelease = Get-RegistryValue -Path $wuPath -Name 'TargetReleaseVersion'
    $targetReleaseInfo = Get-RegistryValue -Path $wuPath -Name 'TargetReleaseVersionInfo'
    $productVersion = Get-RegistryValue -Path $wuPath -Name 'ProductVersion'
    $disableDualScan = Get-RegistryValue -Path $wuPath -Name 'DisableDualScan'
    $isWufbConfigured = Get-RegistryValue -Path $policyStatePath -Name 'IsWUfBConfigured'

    if ($wus) { [void]$issues.Add('WUServer policy present') }
    if ($wuStatus) { [void]$issues.Add('WUStatusServer policy present') }
    if ($useWUServer -eq 1) { [void]$issues.Add('UseWUServer=1') }
    if ($noAutoUpdate -eq 1) { [void]$issues.Add('NoAutoUpdate=1') }
    if ($doNotConnect -eq 1) { [void]$issues.Add('DoNotConnectToWindowsUpdateInternetLocations=1') }
    if ($disableDualScan -eq 1) { [void]$issues.Add('DisableDualScan=1') }
    if ($targetRelease -eq 1 -and $targetReleaseInfo -and $productVersion -notmatch 'Windows 11') {
        [void]$issues.Add(("TargetReleaseVersion holds {0} {1}" -f $productVersion,$targetReleaseInfo).Trim())
    }

    [pscustomobject]@{
        HasLegacyBlocker = ($issues.Count -gt 0)
        Issues = (@($issues.ToArray()) -join '; ')
        MdmWUfBPolicyPresent = (Test-Path -LiteralPath $mdmUpdatePath)
        IsWUfBConfigured = if ($null -eq $isWufbConfigured) { 'Unknown' } else { [string]$isWufbConfigured }
        TargetReleaseVersion = if ($null -eq $targetRelease) { '' } else { [string]$targetRelease }
        TargetReleaseVersionInfo = [string]$targetReleaseInfo
        ProductVersion = [string]$productVersion
    }
}

function ConvertTo-Windows11IndicatorSignal {
    param([object]$Value)

    if ($null -eq $Value) { return '' }

    $values = if ($Value -is [System.Array]) { @($Value) } else { @($Value) }
    $signals = @(
        $values |
            ForEach-Object { ([string]$_).Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -notmatch '(?i)^(none|n/a|na|null)$' }
    )

    if ($signals.Count -eq 0) { return '' }
    return (@($signals) -join '; ')
}

function Get-Windows11IndicatorSummary {
    $path = 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\TargetVersionUpgradeExperienceIndicators'
    $ignoredHardware = @('Cpu','CpuFms','CpuVendor','CpuFamily','CpuModel','CpuSpeed','CpuCores','Tpm','TpmVersion','SecureBoot','UefiSecureBoot','Uefi','Ram','Memory','Storage','Disk','SystemDriveSize')
    $blocking = New-Object System.Collections.ArrayList
    $hardwareOnly = New-Object System.Collections.ArrayList

    if (-not (Test-Path -LiteralPath $path)) {
        return [pscustomobject]@{
            IndicatorPathPresent = $false
            BlockingCount = 0
            HardwareOnlyCount = 0
            ActionableBlocking = $false
            BlockingReasons = ''
        }
    }

    foreach ($key in @(Get-ChildItem -LiteralPath $path -ErrorAction SilentlyContinue)) {
        $target = $key.PSChildName
        if ($target -match '^(1507|1511|1607|1703|1709|1803|1809|1903|1909|19H1|19H2|2004|20H1|20H2|21H1|TH|RS)') {
            continue
        }

        $props = Get-ItemProperty -LiteralPath $key.PSPath -ErrorAction SilentlyContinue
        $upEx = if ($props.PSObject.Properties['UpEx']) { ConvertTo-Windows11IndicatorSignal -Value $props.UpEx } else { '' }
        $gated = if ($props.PSObject.Properties['GatedBlockId']) { ConvertTo-Windows11IndicatorSignal -Value $props.GatedBlockId } else { '' }
        $red = if ($props.PSObject.Properties['RedReason']) { ConvertTo-Windows11IndicatorSignal -Value $props.RedReason } else { '' }
        $sysReq = if ($props.PSObject.Properties['SysReqIssue']) { ConvertTo-Windows11IndicatorSignal -Value $props.SysReqIssue } else { '' }
        $hasBlock = ($upEx -match '(Red|Blocked|Hold)' -or -not [string]::IsNullOrWhiteSpace($gated) -or -not [string]::IsNullOrWhiteSpace($red) -or -not [string]::IsNullOrWhiteSpace($sysReq))
        if (-not $hasBlock) { continue }

        $reason = ("Target={0}; UpEx={1}; GatedBlockId={2}; RedReason={3}; SysReqIssue={4}" -f $target,$upEx,$gated,$red,$sysReq)
        $tokens = @(@($red,$sysReq -split '[,; ]+') | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
        $actionable = $false
        if ($gated) { $actionable = $true }
        elseif ($tokens.Count -eq 0 -and $upEx -match '(Red|Blocked|Hold)') { $actionable = $true }
        else {
            foreach ($token in $tokens) {
                if ($token -notin $ignoredHardware) {
                    $actionable = $true
                    break
                }
            }
        }

        if ($actionable) { [void]$blocking.Add($reason) } else { [void]$hardwareOnly.Add($reason) }
    }

    [pscustomobject]@{
        IndicatorPathPresent = $true
        BlockingCount = $blocking.Count
        HardwareOnlyCount = $hardwareOnly.Count
        ActionableBlocking = ($blocking.Count -gt 0)
        BlockingReasons = (@($blocking.ToArray()) -join ' | ')
    }
}

function Repair-WindowsUpdatePolicyBlockers {
    Write-SmartLog 'Repairing legacy Windows Update policy blockers.'
    $wuPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    $auPath = Join-Path $wuPath 'AU'
    $names = @(
        @{ Path = $wuPath; Name = 'WUServer' },
        @{ Path = $wuPath; Name = 'WUStatusServer' },
        @{ Path = $wuPath; Name = 'DoNotConnectToWindowsUpdateInternetLocations' },
        @{ Path = $wuPath; Name = 'DisableDualScan' },
        @{ Path = $wuPath; Name = 'TargetReleaseVersion' },
        @{ Path = $wuPath; Name = 'TargetReleaseVersionInfo' },
        @{ Path = $wuPath; Name = 'ProductVersion' },
        @{ Path = $auPath; Name = 'UseWUServer' },
        @{ Path = $auPath; Name = 'NoAutoUpdate' }
    )

    foreach ($entry in $names) {
        if (Test-Path -LiteralPath $entry.Path) {
            Remove-ItemProperty -LiteralPath $entry.Path -Name $entry.Name -ErrorAction SilentlyContinue
        }
    }

    gpupdate.exe /target:computer /force | Out-Null
}

function Reset-WindowsUpdateComponents {
    Write-SmartLog 'Resetting Windows Update components.'
    $services = 'bits','wuauserv','cryptsvc','dosvc'
    foreach ($service in $services) {
        try { Stop-Service -Name $service -Force -ErrorAction SilentlyContinue } catch { }
    }

    $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    foreach ($folder in @('SoftwareDistribution','System32\catroot2')) {
        $path = Join-Path $env:WINDIR $folder
        if (Test-Path -LiteralPath $path) {
            Rename-Item -LiteralPath $path -NewName ("{0}.SmartM365Backup.{1}" -f (Split-Path -Leaf $path),$stamp) -ErrorAction SilentlyContinue
        }
    }

    foreach ($service in $services) {
        try { Start-Service -Name $service -ErrorAction SilentlyContinue } catch { }
    }
}

function Invoke-AssignedUpdateInstall {
    Write-SmartLog 'Starting assigned Windows Update scan/download/install.'
    $resultText = 'NoUpdates'
    try {
        $session = New-Object -ComObject 'Microsoft.Update.Session'
        $searcher = $session.CreateUpdateSearcher()
        $result = $searcher.Search('IsInstalled=0 and IsHidden=0 and IsAssigned=1')
        if ($result.Updates.Count -gt 0) {
            $updates = New-Object -ComObject 'Microsoft.Update.UpdateColl'
            foreach ($update in $result.Updates) { [void]$updates.Add($update) }
            $downloader = $session.CreateUpdateDownloader()
            $downloader.Updates = $updates
            [void]$downloader.Download()
            $installer = $session.CreateUpdateInstaller()
            $installer.Updates = $updates
            $installResult = $installer.Install()
            $resultText = if ($installResult.RebootRequired) { 'InstalledRebootRequired' } else { 'Installed' }
        }
    }
    catch {
        Write-SmartLog ("COM Windows Update install failed: {0}" -f $_.Exception.Message) 'WARN'
    }

    $uso = Join-Path $env:SystemRoot 'System32\UsoClient.exe'
    if (Test-Path -LiteralPath $uso -PathType Leaf) {
        foreach ($arg in @('StartScan','StartDownload','StartInstall')) {
            Start-Process -FilePath $uso -ArgumentList $arg -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue
        }
        if ($resultText -eq 'NoUpdates') { $resultText = 'UsoClientTriggered' }
    }

    return $resultText
}

function Get-SystemInstallLanguageTag {
    try {
        $systemLocale = Get-WinSystemLocale -ErrorAction Stop
        if ($systemLocale -and -not [string]::IsNullOrWhiteSpace($systemLocale.Name)) {
            return [string]$systemLocale.Name
        }
    }
    catch { }

    try {
        $languageKey = 'HKLM:\SYSTEM\CurrentControlSet\Control\Nls\Language'
        $installLanguage = [string](Get-RegistryValue -Path $languageKey -Name 'InstallLanguage')
        if (-not [string]::IsNullOrWhiteSpace($installLanguage)) {
            $lcid = [Convert]::ToInt32($installLanguage, 16)
            return ([System.Globalization.CultureInfo]::GetCultureInfo($lcid)).Name
        }
    }
    catch { }

    try {
        return [System.Globalization.CultureInfo]::InstalledUICulture.Name
    }
    catch {
        return ''
    }
}

function Resolve-SetupLanguageRequirement {
    param([string]$RequestedLanguage)

    if ([string]::IsNullOrWhiteSpace($RequestedLanguage)) { return '' }
    $normalized = $RequestedLanguage.Trim()
    if ($normalized -in @('Any','None','Disabled')) { return '' }
    if ($normalized -in @('Auto','MatchSystem','System')) {
        $systemLanguage = Get-SystemInstallLanguageTag
        if ([string]::IsNullOrWhiteSpace($systemLanguage)) {
            throw 'Unable to detect the installed Windows system language required for setup media validation.'
        }
        Write-SmartLog ("Setup language requirement resolved from system language: {0}" -f $systemLanguage)
        return $systemLanguage
    }

    try {
        return ([System.Globalization.CultureInfo]::GetCultureInfo($normalized)).Name
    }
    catch {
        throw "Invalid SetupLanguage value '$RequestedLanguage'. Use MatchSystem, Any, or a culture tag such as fr-FR or en-US."
    }
}

function Get-SetupMediaLanguages {
    param([Parameter(Mandatory = $true)][string]$MediaRoot)

    $langIni = Join-Path $MediaRoot 'sources\lang.ini'
    if (-not (Test-Path -LiteralPath $langIni -PathType Leaf)) { return @() }

    $languages = New-Object System.Collections.ArrayList
    $inAvailableSection = $false
    foreach ($rawLine in @(Get-Content -LiteralPath $langIni -ErrorAction Stop)) {
        $line = ([string]$rawLine).Trim()
        if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith(';')) { continue }

        if ($line -match '^\[(.+)\]$') {
            $inAvailableSection = ($Matches[1] -ieq 'Available UI Languages')
            continue
        }

        if ($inAvailableSection -and $line -match '^([^=]+)=') {
            $language = $Matches[1].Trim()
            if (-not [string]::IsNullOrWhiteSpace($language) -and -not $languages.Contains($language)) {
                [void]$languages.Add($language)
            }
        }
    }

    return @($languages.ToArray())
}

function Resolve-SetupSourceMediaPath {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [string]$ExpectedLanguage
    )

    if (Test-Path -LiteralPath (Join-Path $SourcePath 'setup.exe') -PathType Leaf) {
        return $SourcePath
    }

    if ([string]::IsNullOrWhiteSpace($ExpectedLanguage)) {
        throw "SetupSourcePath does not contain setup.exe and language matching is disabled, so no language subfolder can be selected: $SourcePath"
    }

    foreach ($child in @(Get-ChildItem -LiteralPath $SourcePath -Directory -ErrorAction SilentlyContinue)) {
        $languages = @(Get-SetupMediaLanguages -MediaRoot $child.FullName)
        if (@($languages | Where-Object { $_ -ieq $ExpectedLanguage } | Select-Object -First 1).Count -gt 0) {
            Write-SmartLog ("Selected setup source subfolder '{0}' for language {1}." -f $child.FullName,$ExpectedLanguage)
            return $child.FullName
        }
    }

    throw ("No setup source subfolder under '{0}' contains language {1} in sources\lang.ini." -f $SourcePath,$ExpectedLanguage)
}

function Split-SetupSourcePathList {
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return @() }

    $result = New-Object System.Collections.ArrayList
    foreach ($item in @($Value -split ';')) {
        $path = ([string]$item).Trim().Trim([char]34)
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        [void]$result.Add($path)
    }

    return @($result.ToArray())
}

function Convert-IPv4ToUInt32 {
    param([Parameter(Mandatory = $true)][string]$Address)

    $parsed = [System.Net.IPAddress]::Parse($Address)
    if ($parsed.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) {
        throw "Not an IPv4 address: $Address"
    }

    $bytes = $parsed.GetAddressBytes()
    return (
        ([uint32]$bytes[0] -shl 24) -bor
        ([uint32]$bytes[1] -shl 16) -bor
        ([uint32]$bytes[2] -shl 8) -bor
        [uint32]$bytes[3]
    )
}

function Test-IPv4InCidr {
    param(
        [Parameter(Mandatory = $true)][string]$Address,
        [Parameter(Mandatory = $true)][string]$Cidr
    )

    $parts = $Cidr.Trim() -split '/', 2
    if ($parts.Count -ne 2) { return $false }

    $prefixLength = 0
    if (-not [int]::TryParse($parts[1], [ref]$prefixLength)) { return $false }
    if ($prefixLength -lt 0 -or $prefixLength -gt 32) { return $false }

    $addressInt = Convert-IPv4ToUInt32 -Address $Address
    $networkInt = Convert-IPv4ToUInt32 -Address $parts[0]
    $mask = if ($prefixLength -eq 0) { [uint32]0 } else { [uint32]([uint32]::MaxValue -shl (32 - $prefixLength)) }
    return (($addressInt -band $mask) -eq ($networkInt -band $mask))
}

function Get-LocalIPv4Addresses {
    $addresses = New-Object System.Collections.ArrayList

    try {
        foreach ($adapter in @(Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True' -ErrorAction Stop)) {
            foreach ($ip in @($adapter.IPAddress)) {
                $parsed = $null
                if ([System.Net.IPAddress]::TryParse([string]$ip, [ref]$parsed) -and $parsed.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork -and -not [System.Net.IPAddress]::IsLoopback($parsed)) {
                    [void]$addresses.Add($parsed.ToString())
                }
            }
        }
    }
    catch {
        foreach ($ip in @([System.Net.Dns]::GetHostAddresses($env:COMPUTERNAME))) {
            if ($ip.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork -and -not [System.Net.IPAddress]::IsLoopback($ip)) {
                [void]$addresses.Add($ip.ToString())
            }
        }
    }

    return @($addresses.ToArray() | Select-Object -Unique)
}

function Test-SetupSourceMapRowMatch {
    param(
        [Parameter(Mandatory = $true)]$Row,
        [string[]]$LocalIPv4Addresses = @()
    )

    $scopeType = ''
    if ($Row.PSObject.Properties['ScopeType']) {
        $scopeType = ([string]$Row.ScopeType).Trim()
    }
    if ([string]::IsNullOrWhiteSpace($scopeType) -and $Row.PSObject.Properties['Type']) {
        $scopeType = ([string]$Row.Type).Trim()
    }

    $scopeValue = ''
    if ($Row.PSObject.Properties['ScopeValue']) {
        $scopeValue = ([string]$Row.ScopeValue).Trim()
    }
    if ([string]::IsNullOrWhiteSpace($scopeValue) -and $Row.PSObject.Properties['Value']) {
        $scopeValue = ([string]$Row.Value).Trim()
    }

    switch -Regex ($scopeType) {
        '^(?i:Default|Fallback)$' { return $true }
        '^(?i:Subnet|CIDR)$' {
            foreach ($ip in @($LocalIPv4Addresses)) {
                try {
                    if (Test-IPv4InCidr -Address $ip -Cidr $scopeValue) { return $true }
                }
                catch { }
            }
            return $false
        }
        '^(?i:IPPrefix|Prefix)$' {
            foreach ($ip in @($LocalIPv4Addresses)) {
                if ($ip.StartsWith($scopeValue, [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
            }
            return $false
        }
        '^(?i:ComputerName|Hostname)$' { return ($env:COMPUTERNAME -ieq $scopeValue) }
        '^(?i:ComputerPrefix|HostnamePrefix)$' { return ($env:COMPUTERNAME.StartsWith($scopeValue, [System.StringComparison]::OrdinalIgnoreCase)) }
        default { return $false }
    }
}

function Get-SetupSourceCandidatesFromMap {
    param([AllowNull()][string]$MapPath)

    if ([string]::IsNullOrWhiteSpace($MapPath)) { return @() }

    $expandedMapPath = [Environment]::ExpandEnvironmentVariables($MapPath.Trim('"'))
    if (-not (Test-Path -LiteralPath $expandedMapPath -PathType Leaf)) {
        throw "Setup source map not found from target context: $expandedMapPath"
    }

    $mapPathToRead = $expandedMapPath
    try {
        New-SmartDirectory -Path $script:InputDir
        $cachedMapPath = Join-Path $script:InputDir 'SetupSourceMap.csv'
        Copy-Item -LiteralPath $expandedMapPath -Destination $cachedMapPath -Force -ErrorAction Stop
        $mapPathToRead = $cachedMapPath
        Write-SmartLog ("Setup source map cached locally: {0}" -f $cachedMapPath)
    }
    catch {
        Write-SmartLog ("Could not cache setup source map locally; reading original path. Detail={0}" -f $_.Exception.Message) 'WARN'
    }

    $rows = @(Import-Csv -LiteralPath $mapPathToRead -ErrorAction Stop)
    $localIps = @(Get-LocalIPv4Addresses)
    Write-SmartLog ("Setup source map loaded: Path={0}; Rows={1}; LocalIPv4={2}" -f $mapPathToRead,$rows.Count,($localIps -join ','))

    $matches = New-Object System.Collections.ArrayList
    foreach ($row in $rows) {
        $sourcePath = ''
        if ($row.PSObject.Properties['SetupSourcePath']) {
            $sourcePath = ([string]$row.SetupSourcePath).Trim()
        }
        if ([string]::IsNullOrWhiteSpace($sourcePath) -and $row.PSObject.Properties['SourcePath']) {
            $sourcePath = ([string]$row.SourcePath).Trim()
        }
        if ([string]::IsNullOrWhiteSpace($sourcePath)) { continue }

        if (Test-SetupSourceMapRowMatch -Row $row -LocalIPv4Addresses $localIps) {
            $priority = 100
            if ($row.PSObject.Properties['Priority']) {
                [void][int]::TryParse([string]$row.Priority, [ref]$priority)
            }
            [void]$matches.Add([pscustomobject]@{ Priority = $priority; SourcePath = $sourcePath })
        }
    }

    $matchedSources = @($matches.ToArray() | Sort-Object Priority, SourcePath | ForEach-Object { $_.SourcePath } | Select-Object -Unique)
    if ($SetupSourceCandidateLimit -gt 0 -and $matchedSources.Count -gt $SetupSourceCandidateLimit) {
        Write-SmartLog ("Setup source map matched {0} source(s); limiting benchmark shortlist to {1} by priority." -f $matchedSources.Count,$SetupSourceCandidateLimit)
        $matchedSources = @($matchedSources | Select-Object -First $SetupSourceCandidateLimit)
    }
    return $matchedSources
}

function Get-EffectiveSetupSourceCandidates {
    $candidates = New-Object System.Collections.ArrayList
    try {
        foreach ($source in @(Get-SetupSourceCandidatesFromMap -MapPath $SetupSourceMapPath)) {
            [void]$candidates.Add($source)
        }
    }
    catch {
        Write-SmartLog ("Setup source map could not be used: {0}" -f $_.Exception.Message) 'WARN'
    }
    foreach ($source in @(Split-SetupSourcePathList -Value $SetupSourcePath)) {
        [void]$candidates.Add($source)
    }

    return @($candidates.ToArray() | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
}

function Get-UncServerName {
    param([AllowNull()][string]$Path)

    $match = [regex]::Match([string]$Path, '^\\\\(?<server>[^\\]+)\\')
    if ($match.Success) { return $match.Groups['server'].Value }
    return ''
}

function Measure-TcpConnectionMilliseconds {
    param(
        [Parameter(Mandatory = $true)][string]$ServerName,
        [int]$Port = 445,
        [int]$TimeoutMilliseconds = 1200
    )

    $client = New-Object System.Net.Sockets.TcpClient
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $async = $client.BeginConnect($ServerName, $Port, $null, $null)
        if (-not $async.AsyncWaitHandle.WaitOne($TimeoutMilliseconds, $false)) {
            throw "TCP $ServerName`:$Port timed out after $TimeoutMilliseconds ms"
        }

        $client.EndConnect($async)
        return [int][math]::Max(1, $watch.ElapsedMilliseconds)
    }
    finally {
        $watch.Stop()
        try { $client.Close() } catch { }
    }
}

function Measure-FileReadSampleMilliseconds {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [int64]$MaxBytes = 8388608,
        [int]$BufferBytes = 1048576
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "File not found for read sample: $Path"
    }

    $buffer = New-Object byte[] $BufferBytes
    $bytesRead = 0L
    $watch = [System.Diagnostics.Stopwatch]::StartNew()
    $stream = $null
    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        while ($bytesRead -lt $MaxBytes) {
            $remaining = [int][math]::Min($buffer.Length, ($MaxBytes - $bytesRead))
            if ($remaining -le 0) { break }
            $read = $stream.Read($buffer, 0, $remaining)
            if ($read -le 0) { break }
            $bytesRead += [int64]$read
        }
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
        $watch.Stop()
    }

    return [pscustomobject]@{
        Path = $Path
        BytesRead = $bytesRead
        Milliseconds = [int][math]::Max(1, $watch.ElapsedMilliseconds)
    }
}

function Measure-SetupSourceReadSample {
    param([Parameter(Mandatory = $true)][string]$MediaRoot)

    $samples = New-Object System.Collections.ArrayList
    $langIni = Join-Path $MediaRoot 'sources\lang.ini'
    if (Test-Path -LiteralPath $langIni -PathType Leaf) {
        [void]$samples.Add((Measure-FileReadSampleMilliseconds -Path $langIni -MaxBytes 262144 -BufferBytes 65536))
    }

    $setupExe = Join-Path $MediaRoot 'setup.exe'
    [void]$samples.Add((Measure-FileReadSampleMilliseconds -Path $setupExe -MaxBytes 8388608 -BufferBytes 1048576))

    $totalBytes = 0L
    $totalMs = 0
    foreach ($sample in @($samples.ToArray())) {
        $totalBytes += [int64]$sample.BytesRead
        $totalMs += [int]$sample.Milliseconds
    }

    return [pscustomobject]@{
        BytesRead = $totalBytes
        Milliseconds = [int][math]::Max(1, $totalMs)
    }
}

function Resolve-PreferredSetupSourcePath {
    param(
        [Parameter(Mandatory = $true)][string[]]$SourcePaths,
        [string]$ExpectedLanguage
    )

    $candidates = @(
        @(
            foreach ($sourcePath in @($SourcePaths)) {
                Split-SetupSourcePathList -Value $sourcePath
            }
        ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
    )
    if ($candidates.Count -eq 0) {
        throw 'SetupSourcePath is empty.'
    }

    if ($candidates.Count -eq 1) {
        $singleSource = [Environment]::ExpandEnvironmentVariables($candidates[0])
        try {
            if (-not (Test-Path -LiteralPath $singleSource -PathType Container -ErrorAction Stop)) {
                throw "Path is not reachable as a container from target context: $singleSource"
            }
            $singleSource = Resolve-SetupSourceMediaPath -SourcePath $singleSource -ExpectedLanguage $ExpectedLanguage
            $null = Test-SetupMedia -MediaPath $singleSource -ExpectedLanguage $ExpectedLanguage
            Write-SmartLog ("Setup source candidate OK: Path={0}; Resolved={1}; Selection=SingleSource" -f $candidates[0],$singleSource)
        }
        catch {
            $accessHint = "Setup source is validated from the target SYSTEM context. For UNC sources, grant Share and NTFS Read to the target computer account '$script:ComputerName`$' or to a group such as Domain Computers."
            throw ("Setup source is not reachable or valid from this target. Path={0}; Error={1}; {2}" -f $singleSource,$_.Exception.Message,$accessHint)
        }
        $script:SelectedSetupSourcePath = [string]$singleSource
        $script:SetupSourceSelectionDetail = 'SingleSource'
        return $singleSource
    }

    Write-SmartLog ("Selecting nearest setup source from {0} candidate(s): {1}" -f $candidates.Count,($candidates -join ' | '))

    $validSources = New-Object System.Collections.ArrayList
    $errors = New-Object System.Collections.ArrayList
    foreach ($candidate in $candidates) {
        $expanded = [Environment]::ExpandEnvironmentVariables($candidate)
        $server = Get-UncServerName -Path $expanded
        $tcpMs = 0
        $watch = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            if (-not [string]::IsNullOrWhiteSpace($server)) {
                $tcpMs = Measure-TcpConnectionMilliseconds -ServerName $server
            }

            if (-not (Test-Path -LiteralPath $expanded -PathType Container)) {
                throw "Path is not reachable as a container from target context: $expanded"
            }

            $resolved = Resolve-SetupSourceMediaPath -SourcePath $expanded -ExpectedLanguage $ExpectedLanguage
            $null = Test-SetupMedia -MediaPath $resolved -ExpectedLanguage $ExpectedLanguage
            $readSample = Measure-SetupSourceReadSample -MediaRoot $resolved
            $watch.Stop()
            $scoreMs = [int][math]::Max(1, ([int]$readSample.Milliseconds + [int]$tcpMs))
            [void]$validSources.Add([pscustomobject]@{
                Candidate = $candidate
                ResolvedPath = $resolved
                Server = $server
                TcpMilliseconds = $tcpMs
                ReadMilliseconds = [int]$readSample.Milliseconds
                ReadBytes = [int64]$readSample.BytesRead
                ScoreMilliseconds = $scoreMs
            })
            Write-SmartLog ("Setup source candidate OK: Path={0}; Resolved={1}; TcpMs={2}; ReadMs={3}; ReadBytes={4}; ScoreMs={5}" -f $candidate,$resolved,$tcpMs,$readSample.Milliseconds,$readSample.BytesRead,$scoreMs)
        }
        catch {
            $watch.Stop()
            $message = "Path=$candidate; Error=$($_.Exception.Message)"
            [void]$errors.Add($message)
            Write-SmartLog ("Setup source candidate skipped: {0}" -f $message) 'WARN'
        }
    }

    if ($validSources.Count -eq 0) {
        throw ("No setup source candidate is reachable and valid from this target. Errors: {0}" -f (@($errors.ToArray()) -join ' | '))
    }

    $selected = @($validSources.ToArray() | Sort-Object ScoreMilliseconds, ReadMilliseconds, TcpMilliseconds, ResolvedPath | Select-Object -First 1)[0]
    $null = Test-SetupMedia -MediaPath $selected.ResolvedPath -ExpectedLanguage $ExpectedLanguage
    Write-SmartLog ("Selected nearest setup source: {0}; TcpMs={1}; ReadMs={2}; ReadBytes={3}; ScoreMs={4}" -f $selected.ResolvedPath,$selected.TcpMilliseconds,$selected.ReadMilliseconds,$selected.ReadBytes,$selected.ScoreMilliseconds)
    $script:SelectedSetupSourcePath = [string]$selected.ResolvedPath
    $script:SetupSourceSelectionDetail = ("Candidates={0}; Selected={1}; TcpMs={2}; ReadMs={3}; ReadBytes={4}; ScoreMs={5}" -f $candidates.Count,$selected.ResolvedPath,$selected.TcpMilliseconds,$selected.ReadMilliseconds,$selected.ReadBytes,$selected.ScoreMilliseconds)
    return [string]$selected.ResolvedPath
}

function Test-SetupExecutableSignature {
    param([Parameter(Mandatory = $true)][string]$SetupExe)

    try {
        $signature = Get-AuthenticodeSignature -LiteralPath $SetupExe -ErrorAction Stop
    }
    catch {
        throw ("setup.exe Authenticode signature could not be checked. Path={0}; Error={1}" -f $SetupExe,$_.Exception.Message)
    }

    $status = [string]$signature.Status
    $statusMessage = [string]$signature.StatusMessage
    if ($status -ne 'Valid') {
        throw ("setup.exe Authenticode signature is not valid. Path={0}; Status={1}; StatusMessage={2}" -f $SetupExe,$status,$statusMessage)
    }

    if ($null -eq $signature.SignerCertificate) {
        throw ("setup.exe Authenticode signature is valid but signer certificate is missing. Path={0}" -f $SetupExe)
    }

    $signerSubject = [string]$signature.SignerCertificate.Subject
    if ($signerSubject -notmatch '(?i)(CN|O)=Microsoft Corporation') {
        throw ("setup.exe Authenticode signer is not Microsoft Corporation. Path={0}; Signer={1}" -f $SetupExe,$signerSubject)
    }

    Write-SmartLog ("Authenticode validated setup executable: Path={0}; Signer={1}; Thumbprint={2}" -f $SetupExe,$signerSubject,$signature.SignerCertificate.Thumbprint)
}

function Get-SetupMediaIntegrityManifestPath {
    param([Parameter(Mandatory = $true)][string]$MediaRoot)

    return Join-Path $MediaRoot $script:SetupMediaManifestFileName
}

function Get-SetupMediaIntegrityManifestHash {
    param([Parameter(Mandatory = $true)][string]$MediaRoot)

    $manifestPath = Get-SetupMediaIntegrityManifestPath -MediaRoot $MediaRoot
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { return '' }
    return (Get-SmartFileSha256 -Path $manifestPath)
}

function Test-SetupMediaIntegrityManifest {
    param([Parameter(Mandatory = $true)][string]$MediaRoot)

    $manifestPath = Get-SetupMediaIntegrityManifestPath -MediaRoot $MediaRoot
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { return $false }

    $rootFull = [System.IO.Path]::GetFullPath($MediaRoot).TrimEnd('\', '/') + [System.IO.Path]::DirectorySeparatorChar
    $rows = @(Import-Csv -LiteralPath $manifestPath -ErrorAction Stop)
    if ($rows.Count -eq 0) {
        throw "Setup media integrity manifest is empty: $manifestPath"
    }

    foreach ($column in @('RelativePath','Length','SHA256')) {
        if (-not $rows[0].PSObject.Properties[$column]) {
            throw "Setup media integrity manifest is missing required column '$column': $manifestPath"
        }
    }

    $checkedFiles = 0
    $checkedBytes = 0L
    foreach ($row in $rows) {
        $relativePath = ([string]$row.RelativePath).Trim().TrimStart('\', '/')
        if ([string]::IsNullOrWhiteSpace($relativePath)) {
            throw "Setup media integrity manifest contains an empty RelativePath: $manifestPath"
        }
        if ([System.IO.Path]::IsPathRooted($relativePath) -or $relativePath -match '(^|[\\/])\.\.([\\/]|$)') {
            throw ("Setup media integrity manifest contains an unsafe RelativePath. Manifest={0}; RelativePath={1}" -f $manifestPath,$relativePath)
        }

        $filePath = Join-Path $MediaRoot $relativePath
        $fileFull = [System.IO.Path]::GetFullPath($filePath)
        if (-not $fileFull.StartsWith($rootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
            throw ("Setup media integrity manifest path escapes media root. Manifest={0}; RelativePath={1}" -f $manifestPath,$relativePath)
        }
        if (-not (Test-Path -LiteralPath $fileFull -PathType Leaf)) {
            throw ("Setup media integrity check failed. File missing. Manifest={0}; RelativePath={1}" -f $manifestPath,$relativePath)
        }

        $expectedLength = 0L
        if (-not [int64]::TryParse([string]$row.Length, [ref]$expectedLength)) {
            throw ("Setup media integrity manifest contains an invalid Length. Manifest={0}; RelativePath={1}; Length={2}" -f $manifestPath,$relativePath,$row.Length)
        }

        $item = Get-Item -LiteralPath $fileFull -ErrorAction Stop
        if ([int64]$item.Length -ne $expectedLength) {
            throw ("Setup media integrity check failed. Length mismatch. Manifest={0}; RelativePath={1}; ExpectedLength={2}; ActualLength={3}" -f $manifestPath,$relativePath,$expectedLength,$item.Length)
        }

        $expectedHash = ([string]$row.SHA256).Trim().ToUpperInvariant()
        if ($expectedHash -notmatch '^[A-F0-9]{64}$') {
            throw ("Setup media integrity manifest contains an invalid SHA256. Manifest={0}; RelativePath={1}; SHA256={2}" -f $manifestPath,$relativePath,$row.SHA256)
        }

        $actualHash = (Get-SmartFileSha256 -Path $fileFull).ToUpperInvariant()
        if ($actualHash -ne $expectedHash) {
            throw ("Setup media integrity check failed. SHA256 mismatch. Manifest={0}; RelativePath={1}; ExpectedSHA256={2}; ActualSHA256={3}" -f $manifestPath,$relativePath,$expectedHash,$actualHash)
        }

        $checkedFiles++
        $checkedBytes += [int64]$item.Length
    }

    Write-SmartLog ("Setup media integrity manifest validated: Manifest={0}; Files={1}; Bytes={2}" -f $manifestPath,$checkedFiles,$checkedBytes)
    return $true
}

function Test-SetupMedia {
    param(
        [Parameter(Mandatory = $true)][string]$MediaPath,
        [string]$ExpectedLanguage,
        [switch]$ValidateManifest
    )

    $setupExe = if ([System.IO.Path]::GetFileName($MediaPath) -ieq 'setup.exe') {
        $MediaPath
    }
    else {
        Join-Path $MediaPath 'setup.exe'
    }

    if (-not (Test-Path -LiteralPath $setupExe -PathType Leaf)) {
        throw "setup.exe not found at: $setupExe"
    }

    $setupItem = Get-Item -LiteralPath $setupExe -ErrorAction Stop
    if ($setupItem.Length -lt 64KB) {
        throw "setup.exe exists but size is unexpectedly small: $setupExe"
    }

    Test-SetupExecutableSignature -SetupExe $setupItem.FullName

    $mediaRoot = Split-Path -Parent $setupExe
    $null = Test-SetupInstallImageReadable -MediaRoot $mediaRoot
    if ($ValidateManifest) {
        $null = Test-SetupMediaIntegrityManifest -MediaRoot $mediaRoot
    }

    if (-not [string]::IsNullOrWhiteSpace($ExpectedLanguage)) {
        $mediaLanguages = @(Get-SetupMediaLanguages -MediaRoot $mediaRoot)
        if ($mediaLanguages.Count -eq 0) {
            throw "Setup media language could not be detected from sources\lang.ini. Expected=$ExpectedLanguage; MediaRoot=$mediaRoot"
        }

        $match = @($mediaLanguages | Where-Object { $_ -ieq $ExpectedLanguage } | Select-Object -First 1)
        if ($match.Count -eq 0) {
            throw ("Setup media language mismatch. Expected={0}; Available={1}; MediaRoot={2}" -f $ExpectedLanguage,($mediaLanguages -join ','),$mediaRoot)
        }

        $script:ResolvedSetupMediaLanguages = ($mediaLanguages -join ',')
    }

    return $setupItem.FullName
}
function Convert-ToSafePathSegment {
    param([Parameter(Mandatory = $true)][string]$Value)

    $safe = ($Value.Trim() -replace '[^A-Za-z0-9_.-]', '_')
    if ([string]::IsNullOrWhiteSpace($safe)) { return 'Default' }
    return $safe
}

function Resolve-SetupCachePath {
    param([string]$ExpectedLanguage)

    $mediaSegment = Convert-ToSafePathSegment -Value $SetupMediaId
    if (-not [string]::IsNullOrWhiteSpace($ExpectedLanguage)) {
        $mediaSegment = "{0}-{1}" -f $mediaSegment,(Convert-ToSafePathSegment -Value $ExpectedLanguage)
    }
    return Join-Path $SetupCacheRoot $mediaSegment
}

function Get-SetupInstallImageItem {
    param([Parameter(Mandatory = $true)][string]$MediaRoot)

    foreach ($name in @('sources\install.wim','sources\install.esd')) {
        $path = Join-Path $MediaRoot $name
        if (Test-Path -LiteralPath $path -PathType Leaf) {
            return Get-Item -LiteralPath $path -ErrorAction Stop
        }
    }
    return $null
}

function Test-SetupInstallImageReadable {
    param([Parameter(Mandatory = $true)][string]$MediaRoot)

    $installItem = Get-SetupInstallImageItem -MediaRoot $MediaRoot
    if ($null -eq $installItem) {
        throw "Windows setup media is incomplete. Missing sources\install.wim or sources\install.esd under: $MediaRoot"
    }

    if ($installItem.Length -lt 512MB) {
        throw ("Windows setup media install image is unexpectedly small. Path={0}; SizeBytes={1}" -f $installItem.FullName,$installItem.Length)
    }

    $stream = $null
    try {
        $stream = [System.IO.File]::Open($installItem.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $header = New-Object byte[] 8
        $read = $stream.Read($header, 0, $header.Length)
        if ($read -lt 8) {
            throw ("Windows setup media install image header is truncated. Path={0}; BytesRead={1}" -f $installItem.FullName,$read)
        }

        $signature = [System.Text.Encoding]::ASCII.GetString($header, 0, 5)
        if ($signature -ne 'MSWIM') {
            $hex = (($header | ForEach-Object { $_.ToString('X2') }) -join ' ')
            throw ("Windows setup media install image header is invalid. Expected MSWIM signature. Path={0}; Header={1}" -f $installItem.FullName,$hex)
        }
    }
    finally {
        if ($null -ne $stream) { $stream.Dispose() }
    }

    Test-SetupInstallImageDismReadable -ImagePath $installItem.FullName
    return $installItem
}

function Test-SetupInstallImageDismReadable {
    param([Parameter(Mandatory = $true)][string]$ImagePath)

    $dismPath = Join-Path $env:SystemRoot 'System32\dism.exe'
    if (-not (Test-Path -LiteralPath $dismPath -PathType Leaf)) {
        $dismCommand = Get-Command -Name 'dism.exe' -CommandType Application -ErrorAction SilentlyContinue
        if ($dismCommand) { $dismPath = $dismCommand.Source }
    }

    if (-not (Test-Path -LiteralPath $dismPath -PathType Leaf)) {
        throw 'dism.exe not found; cannot validate Windows setup install image readability.'
    }

    $dismOutput = @(& $dismPath /English /Get-WimInfo ("/WimFile:{0}" -f $ImagePath) 2>&1)
    $dismExitCode = [int]$LASTEXITCODE
    if ($dismExitCode -ne 0) {
        $detail = (($dismOutput | ForEach-Object { [string]$_ }) -join ' | ')
        if ($detail.Length -gt 1200) { $detail = $detail.Substring(0, 1200) + '...' }
        throw ("DISM cannot read Windows setup install image. ExitCode={0}; Path={1}; Detail={2}" -f $dismExitCode,$ImagePath,$detail)
    }

    Write-SmartLog ("DISM validated setup install image readability: {0}" -f $ImagePath)
}

function Get-SetupMediaFingerprint {
    param(
        [Parameter(Mandatory = $true)][string]$MediaRoot,
        [string]$ExpectedLanguage
    )

    $setupExe = Join-Path $MediaRoot 'setup.exe'
    $setupItem = Get-Item -LiteralPath $setupExe -ErrorAction Stop
    $installItem = Get-SetupInstallImageItem -MediaRoot $MediaRoot
    if ($null -eq $installItem) {
        throw "Windows setup media is incomplete. Missing sources\install.wim or sources\install.esd under: $MediaRoot"
    }

    $langIni = Join-Path $MediaRoot 'sources\lang.ini'
    $langHash = ''
    if (Test-Path -LiteralPath $langIni -PathType Leaf) {
        $langHash = Get-SmartFileSha256 -Path $langIni
    }

    [pscustomobject]@{
        MediaId = $SetupMediaId
        ExpectedLanguage = [string]$ExpectedLanguage
        MediaLanguages = (@(Get-SetupMediaLanguages -MediaRoot $MediaRoot) -join ',')
        SetupExeLength = [int64]$setupItem.Length
        SetupExeHash = Get-SmartFileSha256 -Path $setupItem.FullName
        InstallImageName = $installItem.Name
        InstallImageLength = [int64]$installItem.Length
        InstallImageLastWriteUtc = $installItem.LastWriteTimeUtc.ToString('o')
        LangIniHash = $langHash
        IntegrityManifestHash = Get-SetupMediaIntegrityManifestHash -MediaRoot $MediaRoot
    }
}

function Test-SetupCacheManifest {
    param(
        [Parameter(Mandatory = $true)][string]$CachePath,
        [Parameter(Mandatory = $true)]$Fingerprint
    )

    $manifestPath = Join-Path $CachePath 'SmartM365-SetupMedia.json'
    if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { return $false }

    try {
        $manifest = Get-Content -LiteralPath $manifestPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        foreach ($property in @('MediaId','ExpectedLanguage','SetupExeLength','SetupExeHash','InstallImageName','InstallImageLength','LangIniHash','IntegrityManifestHash')) {
            if ([string]$manifest.$property -ne [string]$Fingerprint.$property) { return $false }
        }
        return $true
    }
    catch {
        return $false
    }
}

function Test-SetupFingerprintMatch {
    param(
        [Parameter(Mandatory = $true)]$Left,
        [Parameter(Mandatory = $true)]$Right
    )

    foreach ($property in @('MediaId','ExpectedLanguage','SetupExeLength','SetupExeHash','InstallImageName','InstallImageLength','LangIniHash','IntegrityManifestHash')) {
        if ([string]$Left.$property -ne [string]$Right.$property) { return $false }
    }
    return $true
}

function Save-SetupCacheManifest {
    param(
        [Parameter(Mandatory = $true)][string]$CachePath,
        [Parameter(Mandatory = $true)]$Fingerprint,
        [string]$SourcePath
    )

    $manifestPath = Join-Path $CachePath 'SmartM365-SetupMedia.json'
    $Fingerprint | Add-Member -NotePropertyName SourcePath -NotePropertyValue ([string]$SourcePath) -Force
    $Fingerprint | Add-Member -NotePropertyName WrittenUtc -NotePropertyValue ((Get-Date).ToUniversalTime().ToString('o')) -Force
    $Fingerprint | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $manifestPath -Encoding UTF8 -Force
}

function Get-SetupSourceGateKey {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    $bytes = [System.Text.Encoding]::UTF8.GetBytes($SourcePath.ToUpperInvariant())
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        return (($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString('x2') }) -join '').Substring(0, 32)
    }
    finally {
        $sha.Dispose()
    }
}

function Remove-StaleSetupSourceCopyLeases {
    param([Parameter(Mandatory = $true)][string]$GatePath)

    $nowUtc = (Get-Date).ToUniversalTime()
    foreach ($lease in @(Get-ChildItem -LiteralPath $GatePath -Filter '*.json' -File -ErrorAction SilentlyContinue)) {
        $remove = $false
        try {
            $data = Get-Content -LiteralPath $lease.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $createdUtc = [datetime]::Parse([string]$data.CreatedUtc, $null, [System.Globalization.DateTimeStyles]::RoundtripKind)
            if (($nowUtc - $createdUtc).TotalMinutes -gt $SetupSourceConcurrencyLeaseMinutes) {
                $remove = $true
            }
        }
        catch {
            $remove = $true
        }

        if ($remove) {
            Remove-Item -LiteralPath $lease.FullName -Force -ErrorAction SilentlyContinue
        }
    }
}

function Acquire-SetupSourceCopyLease {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    $gateRoot = [Environment]::ExpandEnvironmentVariables($SetupSourceConcurrencyGateRoot.Trim('"'))
    $sourceKey = Get-SetupSourceGateKey -SourcePath $SourcePath
    $gatePath = Join-Path $gateRoot $sourceKey
    New-SmartDirectory -Path $gatePath

    $leaseId = [guid]::NewGuid().ToString('N')
    $leasePath = Join-Path $gatePath ("{0}_{1}.json" -f $env:COMPUTERNAME,$leaseId)
    $waitSeconds = [math]::Max(60, $SetupSourceConcurrencyLeaseMinutes * 60)
    $started = Get-Date
    do {
        Remove-StaleSetupSourceCopyLeases -GatePath $gatePath
        $activeCount = @(Get-ChildItem -LiteralPath $gatePath -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
        if ($activeCount -lt $SetupSourceConcurrencyLimit) {
            [pscustomobject]@{
                ComputerName = $env:COMPUTERNAME
                ProcessId = $PID
                SourcePath = $SourcePath
                CreatedUtc = (Get-Date).ToUniversalTime().ToString('o')
            } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $leasePath -Encoding UTF8 -Force
            Write-SmartLog ("Acquired setup source copy lease: Source={0}; ActiveBefore={1}; Limit={2}; Lease={3}" -f $SourcePath,$activeCount,$SetupSourceConcurrencyLimit,$leasePath)
            return $leasePath
        }

        Write-SmartLog ("Waiting for setup source copy lease: Source={0}; Active={1}; Limit={2}" -f $SourcePath,$activeCount,$SetupSourceConcurrencyLimit)
        Start-Sleep -Seconds (Get-Random -Minimum 10 -Maximum 31)
    }
    while (((Get-Date) - $started).TotalSeconds -lt $waitSeconds)

    throw ("Timed out waiting for setup source copy lease. Source={0}; Gate={1}; Limit={2}" -f $SourcePath,$gatePath,$SetupSourceConcurrencyLimit)
}

function Release-SetupSourceCopyLease {
    param([AllowNull()][string]$LeasePath)

    if (-not [string]::IsNullOrWhiteSpace($LeasePath)) {
        Remove-Item -LiteralPath $LeasePath -Force -ErrorAction SilentlyContinue
        Write-SmartLog ("Released setup source copy lease: {0}" -f $LeasePath)
    }
}

function Get-SetupUncHostName {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    $trimmed = [Environment]::ExpandEnvironmentVariables($SourcePath.Trim('"'))
    if ($trimmed -notmatch '^\\\\([^\\]+)\\') { return '' }
    return $matches[1]
}

function ConvertTo-IPv4Number {
    param([Parameter(Mandatory = $true)][string]$Address)

    $parts = @($Address.Split('.') | ForEach-Object { [uint32]$_ })
    if ($parts.Count -ne 4) { throw "Invalid IPv4 address: $Address" }
    return [uint32]((($parts[0] -shl 24) -bor ($parts[1] -shl 16) -bor ($parts[2] -shl 8) -bor $parts[3]) -band [uint32]::MaxValue)
}

function ConvertFrom-IPv4Number {
    param([Parameter(Mandatory = $true)][uint32]$Value)

    return ('{0}.{1}.{2}.{3}' -f (($Value -shr 24) -band 255), (($Value -shr 16) -band 255), (($Value -shr 8) -band 255), ($Value -band 255))
}

function Get-IPv4NetworkAddress {
    param(
        [Parameter(Mandatory = $true)][string]$Address,
        [ValidateRange(1, 32)][int]$PrefixLength = 24
    )

    $ipNumber = ConvertTo-IPv4Number -Address $Address
    $mask = if ($PrefixLength -eq 32) { [uint32]::MaxValue } else { [uint32](([uint64]::MaxValue -shl (32 - $PrefixLength)) -band [uint32]::MaxValue) }
    return ConvertFrom-IPv4Number -Value ([uint32]($ipNumber -band $mask))
}

function Get-LocalIPv4ForSetupSource {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    $hostName = Get-SetupUncHostName -SourcePath $SourcePath
    if ([string]::IsNullOrWhiteSpace($hostName)) { return $null }

    $selectedAddress = ''
    try {
        $remoteAddresses = @([System.Net.Dns]::GetHostAddresses($hostName) | Where-Object { $_.AddressFamily -eq [System.Net.Sockets.AddressFamily]::InterNetwork })
        foreach ($remoteAddress in $remoteAddresses) {
            $socket = New-Object System.Net.Sockets.Socket([System.Net.Sockets.AddressFamily]::InterNetwork, [System.Net.Sockets.SocketType]::Dgram, [System.Net.Sockets.ProtocolType]::Udp)
            try {
                $socket.Connect($remoteAddress, 445)
                if ($socket.LocalEndPoint -and $socket.LocalEndPoint.Address) {
                    $candidateAddress = [string]$socket.LocalEndPoint.Address
                    if ($candidateAddress -and $candidateAddress -ne '0.0.0.0' -and -not $candidateAddress.StartsWith('169.254.')) {
                        $selectedAddress = $candidateAddress
                        break
                    }
                }
            }
            catch { }
            finally { $socket.Dispose() }
        }
    }
    catch { }

    try {
        $ipConfig = @(Get-CimInstance -ClassName Win32_NetworkAdapterConfiguration -Filter 'IPEnabled=True' -ErrorAction Stop)
        foreach ($adapter in $ipConfig) {
            $addresses = @($adapter.IPAddress)
            $masks = @($adapter.IPSubnet)
            for ($index = 0; $index -lt $addresses.Count; $index++) {
                $address = [string]$addresses[$index]
                if ($address -notmatch '^\d+\.\d+\.\d+\.\d+$' -or $address.StartsWith('169.254.')) { continue }

                if ([string]::IsNullOrWhiteSpace($selectedAddress)) { $selectedAddress = $address }
                if ($address -eq $selectedAddress) {
                    $prefix = $null
                    if ($index -lt $masks.Count -and [string]$masks[$index] -match '^\d+\.\d+\.\d+\.\d+$') {
                        $prefix = Get-IPv4PrefixLengthFromMask -Mask ([string]$masks[$index])
                    }

                    return [pscustomobject]@{
                        Address = $selectedAddress
                        PrefixLength = $prefix
                    }
                }
            }
        }
    }
    catch { }

    if ([string]::IsNullOrWhiteSpace($selectedAddress)) { return $null }
    return [pscustomobject]@{
        Address = $selectedAddress
        PrefixLength = $null
    }
}

function Get-IPv4PrefixLengthFromMask {
    param([Parameter(Mandatory = $true)][string]$Mask)

    $maskNumber = ConvertTo-IPv4Number -Address $Mask
    $prefix = 0
    for ($bit = 31; $bit -ge 0; $bit--) {
        $bitValue = [uint32]([uint64]1 -shl $bit)
        if (($maskNumber -band $bitValue) -ne 0) { $prefix++ } else { break }
    }

    return $prefix
}

function Resolve-SetupSubnetPrefixLength {
    param([AllowNull()]$LocalAddressInfo)

    $value = ([string]$SetupSubnetPrefixLength).Trim()
    if ([string]::IsNullOrWhiteSpace($value) -or $value -ieq 'Auto') {
        if ($LocalAddressInfo -and $LocalAddressInfo.PSObject.Properties['PrefixLength'] -and $LocalAddressInfo.PrefixLength -ge 1 -and $LocalAddressInfo.PrefixLength -le 32) {
            return [int]$LocalAddressInfo.PrefixLength
        }

        Write-SmartLog 'Unable to detect local subnet prefix length for setup subnet copy gate; falling back to /24. Set -SetupSubnetPrefixLength to force another value.' 'WARN'
        return 24
    }

    $prefixLength = 0
    if (-not [int]::TryParse($value, [ref]$prefixLength) -or $prefixLength -lt 1 -or $prefixLength -gt 32) {
        throw "SetupSubnetPrefixLength must be Auto or an integer from 1 to 32. Value=$value"
    }

    return $prefixLength
}

function Get-SetupSubnetCopyGateInfo {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    $localAddressInfo = Get-LocalIPv4ForSetupSource -SourcePath $SourcePath
    $sourceHost = Get-SetupUncHostName -SourcePath $SourcePath
    if (-not $localAddressInfo -or [string]::IsNullOrWhiteSpace([string]$localAddressInfo.Address)) {
        Write-SmartLog ("Unable to resolve local IPv4 for setup subnet copy gate. Source={0}; using UNKNOWN subnet gate." -f $SourcePath) 'WARN'
        $prefixLength = Resolve-SetupSubnetPrefixLength -LocalAddressInfo $null
        $subnet = 'UNKNOWN'
        $key = 'UNKNOWN_SUBNET'
        $localIPv4 = ''
    }
    else {
        $localIPv4 = [string]$localAddressInfo.Address
        $prefixLength = Resolve-SetupSubnetPrefixLength -LocalAddressInfo $localAddressInfo
        $subnet = Get-IPv4NetworkAddress -Address $localIPv4 -PrefixLength $prefixLength
        $key = ("{0}_{1}" -f $subnet,$prefixLength).Replace('.', '-')
    }

    [pscustomobject]@{
        SourceHost = $sourceHost
        LocalIPv4 = $localIPv4
        Subnet = $subnet
        PrefixLength = $prefixLength
        Key = $key
        Display = if ($subnet -eq 'UNKNOWN') { 'UNKNOWN' } else { "{0}/{1}" -f $subnet,$prefixLength }
    }
}

function Remove-StaleSetupSubnetCopyLeases {
    param([Parameter(Mandatory = $true)][string]$GatePath)

    $nowUtc = (Get-Date).ToUniversalTime()
    foreach ($slot in @(Get-ChildItem -LiteralPath $GatePath -Directory -Filter 'slot-*' -ErrorAction SilentlyContinue)) {
        $remove = $false
        $leaseFile = Join-Path $slot.FullName 'lease.json'
        try {
            $data = Get-Content -LiteralPath $leaseFile -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $seenText = if ($data.PSObject.Properties['LastSeenUtc']) { [string]$data.LastSeenUtc } else { [string]$data.CreatedUtc }
            $lastSeenUtc = [datetime]::Parse($seenText, $null, [System.Globalization.DateTimeStyles]::RoundtripKind)
            if (($nowUtc - $lastSeenUtc).TotalMinutes -gt $SetupSubnetConcurrencyLeaseMinutes) { $remove = $true }
        }
        catch { $remove = $true }

        if ($remove) { Remove-Item -LiteralPath $slot.FullName -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

function Write-SetupSubnetCopyLease {
    param(
        [Parameter(Mandatory = $true)][string]$LeaseFilePath,
        [Parameter(Mandatory = $true)]$GateInfo,
        [Parameter(Mandatory = $true)][string]$SourcePath
    )

    $nowUtc = (Get-Date).ToUniversalTime().ToString('o')
    [pscustomobject]@{
        ComputerName = $env:COMPUTERNAME
        ProcessId = $PID
        RunId = $script:RunId
        SourcePath = $SourcePath
        SourceHost = [string]$GateInfo.SourceHost
        LocalIPv4 = [string]$GateInfo.LocalIPv4
        Subnet = [string]$GateInfo.Display
        CreatedUtc = $nowUtc
        LastSeenUtc = $nowUtc
    } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $LeaseFilePath -Encoding UTF8 -Force
}

function Update-SetupSubnetCopyLeaseHeartbeat {
    param([Parameter(Mandatory = $true)][string]$LeaseFilePath)

    try {
        $data = Get-Content -LiteralPath $LeaseFilePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        $data | Add-Member -NotePropertyName LastSeenUtc -NotePropertyValue ((Get-Date).ToUniversalTime().ToString('o')) -Force
        $data | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $LeaseFilePath -Encoding UTF8 -Force
    }
    catch { }
}

function Start-SetupSubnetCopyLeaseHeartbeat {
    param([Parameter(Mandatory = $true)][string]$LeaseFilePath)

    return Start-Job -ArgumentList $LeaseFilePath -ScriptBlock {
        param([string]$Path)
        while ($true) {
            try {
                $data = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                $data | Add-Member -NotePropertyName LastSeenUtc -NotePropertyValue ((Get-Date).ToUniversalTime().ToString('o')) -Force
                $data | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $Path -Encoding UTF8 -Force
            }
            catch { }
            Start-Sleep -Seconds 60
        }
    }
}

function Stop-SetupSubnetCopyLeaseHeartbeat {
    param([AllowNull()]$Job)

    if ($null -ne $Job) {
        try {
            Stop-Job -Job $Job -ErrorAction SilentlyContinue
            Wait-Job -Job $Job -Timeout 5 -ErrorAction SilentlyContinue | Out-Null
        }
        catch { }

        try {
            Remove-Job -Job $Job -Force -ErrorAction SilentlyContinue
        }
        catch {
            Remove-Job -Job $Job -ErrorAction SilentlyContinue
        }
    }
}

function Acquire-SetupSubnetCopyLease {
    param([Parameter(Mandatory = $true)][string]$SourcePath)

    $gateRoot = [Environment]::ExpandEnvironmentVariables($SetupSubnetConcurrencyGateRoot.Trim('"'))
    $gateInfo = Get-SetupSubnetCopyGateInfo -SourcePath $SourcePath
    $gatePath = Join-Path (Join-Path $gateRoot 'SetupSubnetCopy') $gateInfo.Key
    New-SmartDirectory -Path $gatePath

    $waitSeconds = [math]::Max(60, $SetupSubnetConcurrencyLeaseMinutes * 60)
    $started = Get-Date
    do {
        Remove-StaleSetupSubnetCopyLeases -GatePath $gatePath
        for ($slot = 1; $slot -le $SetupSubnetConcurrencyLimit; $slot++) {
            $slotPath = Join-Path $gatePath ("slot-{0:000}" -f $slot)
            try {
                New-Item -ItemType Directory -Path $slotPath -ErrorAction Stop | Out-Null
                $leaseFilePath = Join-Path $slotPath 'lease.json'
                Write-SetupSubnetCopyLease -LeaseFilePath $leaseFilePath -GateInfo $gateInfo -SourcePath $SourcePath
                Write-SmartLog ("Acquired setup subnet copy lease: Subnet={0}; LocalIPv4={1}; Source={2}; Limit={3}; Lease={4}" -f $gateInfo.Display,$gateInfo.LocalIPv4,$SourcePath,$SetupSubnetConcurrencyLimit,$slotPath)
                return [pscustomobject]@{ LeasePath = $slotPath; LeaseFilePath = $leaseFilePath; GateInfo = $gateInfo }
            }
            catch {
                if (-not (Test-Path -LiteralPath $slotPath -PathType Container)) { throw }
            }
        }

        $activeCount = @(Get-ChildItem -LiteralPath $gatePath -Directory -Filter 'slot-*' -ErrorAction SilentlyContinue).Count
        Write-SmartLog ("Waiting for setup subnet copy lease: Subnet={0}; Active={1}; Limit={2}; Gate={3}" -f $gateInfo.Display,$activeCount,$SetupSubnetConcurrencyLimit,$gatePath)
        Start-Sleep -Seconds (Get-Random -Minimum 10 -Maximum 31)
    }
    while (((Get-Date) - $started).TotalSeconds -lt $waitSeconds)

    throw ("Timed out waiting for setup subnet copy lease. Subnet={0}; Gate={1}; Limit={2}; LeaseMinutes={3}" -f $gateInfo.Display,$gatePath,$SetupSubnetConcurrencyLimit,$SetupSubnetConcurrencyLeaseMinutes)
}

function Release-SetupSubnetCopyLease {
    param([AllowNull()]$Lease)

    if ($null -ne $Lease -and -not [string]::IsNullOrWhiteSpace([string]$Lease.LeasePath)) {
        Remove-Item -LiteralPath ([string]$Lease.LeasePath) -Recurse -Force -ErrorAction SilentlyContinue
        Write-SmartLog ("Released setup subnet copy lease: Subnet={0}; Lease={1}" -f $Lease.GateInfo.Display,$Lease.LeasePath)
    }
}
function Test-SetupCacheReady {
    param(
        [Parameter(Mandatory = $true)][string]$CachePath,
        [string]$ExpectedLanguage
    )

    $setupExe = Test-SetupMedia -MediaPath $CachePath -ExpectedLanguage $ExpectedLanguage -ValidateManifest
    $fingerprint = Get-SetupMediaFingerprint -MediaRoot (Split-Path -Parent $setupExe) -ExpectedLanguage $ExpectedLanguage
    if (-not (Test-SetupCacheManifest -CachePath $CachePath -Fingerprint $fingerprint)) {
        Save-SetupCacheManifest -CachePath $CachePath -Fingerprint $fingerprint -SourcePath 'ExistingCache'
        Write-SmartLog ("Setup cache was valid but manifest was missing or stale. Manifest refreshed: {0}" -f $CachePath)
    }
    return $setupExe
}

function Clear-SetupCachePath {
    param(
        [Parameter(Mandatory = $true)][string]$CachePath,
        [Parameter(Mandatory = $true)][string]$Reason
    )

    if (-not (Test-Path -LiteralPath $CachePath)) { return }

    Write-SmartLog ("Clearing invalid local setup cache before recopy. CachePath={0}; Reason={1}" -f $CachePath,$Reason) 'WARN'
    Remove-Item -LiteralPath $CachePath -Recurse -Force -ErrorAction Stop
}

function Get-RobocopyExitCodeMeaning {
    param([Parameter(Mandatory = $true)][int]$ExitCode)

    $flags = New-Object System.Collections.ArrayList
    if (($ExitCode -band 1) -ne 0) { [void]$flags.Add('Copied') }
    if (($ExitCode -band 2) -ne 0) { [void]$flags.Add('Extra') }
    if (($ExitCode -band 4) -ne 0) { [void]$flags.Add('Mismatch') }
    if (($ExitCode -band 8) -ne 0) { [void]$flags.Add('Failure') }
    if (($ExitCode -band 16) -ne 0) { [void]$flags.Add('Fatal') }
    if ($flags.Count -eq 0) { [void]$flags.Add('NoChange') }

    return (@($flags.ToArray()) -join '+')
}

function Copy-SetupMediaToLocalCache {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$CachePath,
        [string]$ExpectedLanguage
    )

    New-SmartDirectory -Path $CachePath
    $robocopyLog = Join-Path $script:SetupLogDir ("Robocopy_SetupMedia_{0}.log" -f $script:RunId)
    New-SmartDirectory -Path $script:SetupLogDir
    Test-SetupMediaCopyDiskSpace -SourcePath $SourcePath -CachePath $CachePath

    if ($SetupMediaCopyJitterSeconds -gt 0) {
        $jitter = Get-Random -Minimum 0 -Maximum ($SetupMediaCopyJitterSeconds + 1)
        if ($jitter -gt 0) {
            Write-SmartLog ("Waiting {0} second(s) before setup media copy jitter." -f $jitter)
            Start-Sleep -Seconds $jitter
        }
    }

    $sourceLeasePath = ''
    if ($SetupSourceConcurrencyLimit -gt 0 -and -not [string]::IsNullOrWhiteSpace($SetupSourceConcurrencyGateRoot)) {
        $sourceLeasePath = Acquire-SetupSourceCopyLease -SourcePath $SourcePath
    }

    Write-SmartLog ("Copying setup media on target from '{0}' to local cache '{1}'." -f $SourcePath,$CachePath)
    $robocopy = Join-Path $env:SystemRoot 'System32\robocopy.exe'
    try {
        if (Test-Path -LiteralPath $robocopy -PathType Leaf) {
            $robocopyArgs = @($SourcePath, $CachePath, '/MIR', '/Z', '/R:2', '/W:5', '/NP', '/NFL', '/NDL', "/LOG+:$robocopyLog")
            if ($SetupMediaCopyIpGapMilliseconds -gt 0) {
                $robocopyHelp = ''
                try { $robocopyHelp = (& $robocopy /? 2>$null | Out-String) } catch { }
                if ($robocopyHelp -match '(?i)/IPG') {
                    $robocopyArgs += ("/IPG:{0}" -f $SetupMediaCopyIpGapMilliseconds)
                }
                else {
                    Write-SmartLog ("Robocopy does not advertise /IPG support. Continuing without per-PC copy bandwidth delay. Robocopy={0}" -f $robocopy) 'WARN'
                }
            }
            $subnetLease = $null
            $subnetHeartbeatJob = $null
            if ($SetupSubnetConcurrencyLimit -gt 0 -and -not [string]::IsNullOrWhiteSpace($SetupSubnetConcurrencyGateRoot)) {
                $subnetLease = Acquire-SetupSubnetCopyLease -SourcePath $SourcePath
                $subnetHeartbeatJob = Start-SetupSubnetCopyLeaseHeartbeat -LeaseFilePath $subnetLease.LeaseFilePath
            }

            try {
                & $robocopy @robocopyArgs | Out-Null
                $copyExit = [int]$LASTEXITCODE
            }
            finally {
                Stop-SetupSubnetCopyLeaseHeartbeat -Job $subnetHeartbeatJob
                if ($null -ne $subnetLease) {
                    Update-SetupSubnetCopyLeaseHeartbeat -LeaseFilePath $subnetLease.LeaseFilePath
                    Release-SetupSubnetCopyLease -Lease $subnetLease
                }
            }

            if ($copyExit -gt 7) {
                $robocopyMeaning = Get-RobocopyExitCodeMeaning -ExitCode $copyExit
                $failureDetail = "Robocopy setup media copy failed with exit code $copyExit ($robocopyMeaning). Log=$robocopyLog"
                try {
                    $script:SetupCacheAction = 'CopyFailedCacheCleared'
                    Clear-SetupCachePath -CachePath $CachePath -Reason $failureDetail
                }
                catch {
                    $script:SetupCacheAction = 'CopyFailedCacheClearFailed'
                    Write-SmartLog ("Failed to clear local setup cache after Robocopy failure. CachePath={0}; Error={1}" -f $CachePath,$_.Exception.Message) 'WARN'
                }
                throw $failureDetail
            }
            Write-SmartLog ("Robocopy setup media copy completed with exit code {0}. IPG={1}ms; Log={2}" -f $copyExit,$SetupMediaCopyIpGapMilliseconds,$robocopyLog)
        }
        else {
            Copy-Item -Path (Join-Path $SourcePath '*') -Destination $CachePath -Recurse -Force -ErrorAction Stop
            Write-SmartLog 'Setup media copied with Copy-Item because robocopy.exe was not found. Bandwidth limiting is unavailable with this fallback.'
        }
    }
    finally {
        if (-not [string]::IsNullOrWhiteSpace($sourceLeasePath)) {
            Release-SetupSourceCopyLease -LeasePath $sourceLeasePath
        }
    }

    $setupExe = Test-SetupMedia -MediaPath $CachePath -ExpectedLanguage $ExpectedLanguage -ValidateManifest
    $fingerprint = Get-SetupMediaFingerprint -MediaRoot (Split-Path -Parent $setupExe) -ExpectedLanguage $ExpectedLanguage
    Save-SetupCacheManifest -CachePath $CachePath -Fingerprint $fingerprint -SourcePath $SourcePath
    $script:SetupCacheAction = 'CopiedByTarget'
    return $setupExe
}

function Get-PathDriveAvailableBytes {
    param([Parameter(Mandatory = $true)][string]$Path)

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $root = [System.IO.Path]::GetPathRoot($fullPath)
    if ([string]::IsNullOrWhiteSpace($root)) {
        throw "Unable to resolve drive root for path: $Path"
    }

    $drive = New-Object System.IO.DriveInfo($root)
    return [int64]$drive.AvailableFreeSpace
}

function Test-SetupMediaCopyDiskSpace {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$CachePath
    )

    $sourceBytes = Get-PathSizeBytes -Path $SourcePath
    $cacheBytes = Get-PathSizeBytes -Path $CachePath
    $cacheDriveFreeBytes = Get-PathDriveAvailableBytes -Path $CachePath
    $minimumSystemFreeBytes = [int64]$MinimumFreeDiskGB * 1GB
    $additionalCopyBytes = [int64]([math]::Max([double]0, ([double]$sourceBytes - [double]$cacheBytes)))
    $systemDriveRoot = [System.IO.Path]::GetPathRoot((Join-Path $env:SystemDrive '\'))
    $cacheDriveRoot = [System.IO.Path]::GetPathRoot([System.IO.Path]::GetFullPath($CachePath))
    $systemFreeBytes = [int64]((Get-SystemDriveFreeGb) * 1GB)

    if ($cacheDriveRoot -ieq $systemDriveRoot) {
        $requiredCacheDriveFreeBytes = $additionalCopyBytes + $minimumSystemFreeBytes
    }
    else {
        $requiredCacheDriveFreeBytes = $additionalCopyBytes
        if ($systemFreeBytes -lt $minimumSystemFreeBytes) {
            $script:SetupCacheAction = 'CopyBlockedInsufficientDisk'
            throw ("Insufficient system disk space before setup media copy. SystemFreeGB={0}; RequiredSystemFreeGB={1}; CachePath={2}; SourcePath={3}" -f (Convert-BytesToGbText -Bytes $systemFreeBytes),$MinimumFreeDiskGB,$CachePath,$SourcePath)
        }
    }

    Write-SmartLog ("Setup media copy disk preflight: SourceGB={0}; ExistingCacheGB={1}; CacheDriveFreeGB={2}; RequiredCacheDriveFreeGB={3}; SystemFreeGB={4}; RequiredSystemFreeGB={5}; CachePath={6}" -f (Convert-BytesToGbText -Bytes $sourceBytes),(Convert-BytesToGbText -Bytes $cacheBytes),(Convert-BytesToGbText -Bytes $cacheDriveFreeBytes),(Convert-BytesToGbText -Bytes $requiredCacheDriveFreeBytes),(Convert-BytesToGbText -Bytes $systemFreeBytes),$MinimumFreeDiskGB,$CachePath)

    if ($cacheDriveFreeBytes -lt $requiredCacheDriveFreeBytes) {
        $script:SetupCacheAction = 'CopyBlockedInsufficientDisk'
        throw ("Insufficient disk space before setup media copy. CacheDriveFreeGB={0}; RequiredCacheDriveFreeGB={1}; SourceGB={2}; ExistingCacheGB={3}; RequiredSystemFreeGB={4}; CachePath={5}; SourcePath={6}" -f (Convert-BytesToGbText -Bytes $cacheDriveFreeBytes),(Convert-BytesToGbText -Bytes $requiredCacheDriveFreeBytes),(Convert-BytesToGbText -Bytes $sourceBytes),(Convert-BytesToGbText -Bytes $cacheBytes),$MinimumFreeDiskGB,$CachePath,$SourcePath)
    }
}

function Get-PathSizeBytes {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return 0L }
    $total = 0L
    try {
        foreach ($file in @(Get-ChildItem -LiteralPath $Path -Recurse -File -Force -ErrorAction SilentlyContinue)) {
            try { $total += [int64]$file.Length } catch { }
        }
    }
    catch { }
    return $total
}

function Convert-BytesToGbText {
    param([int64]$Bytes)
    return ([math]::Round(([double]$Bytes / 1GB), 2)).ToString('0.##')
}

function Remove-PathChildrenForCleanup {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [datetime]$OlderThan,
        [string[]]$ExcludeLiteralPaths = @()
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return 0 }

    $exclude = @{}
    foreach ($item in @($ExcludeLiteralPaths)) {
        if (-not [string]::IsNullOrWhiteSpace($item)) {
            try { $exclude[[System.IO.Path]::GetFullPath($item).TrimEnd('\').ToUpperInvariant()] = $true } catch { }
        }
    }

    $removed = 0
    foreach ($child in @(Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue)) {
        $childFull = ''
        try { $childFull = [System.IO.Path]::GetFullPath($child.FullName).TrimEnd('\').ToUpperInvariant() } catch { $childFull = '' }
        if ($childFull -and $exclude.ContainsKey($childFull)) { continue }
        if ($PSBoundParameters.ContainsKey('OlderThan') -and $child.LastWriteTime -gt $OlderThan) { continue }
        try {
            Remove-Item -LiteralPath $child.FullName -Recurse -Force -ErrorAction Stop
            $removed++
        }
        catch {
            Write-SmartLog ("Cleanup skipped locked item '{0}': {1}" -f $child.FullName,$_.Exception.Message) 'WARN'
        }
    }
    return $removed
}

function Invoke-CleanupArea {
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [Parameter(Mandatory = $true)][string]$Path,
        [datetime]$OlderThan,
        [string[]]$ExcludeLiteralPaths = @()
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        Write-SmartLog ("Disk cleanup area skipped, path missing: {0} ({1})" -f $Name,$Path)
        return [pscustomobject]@{ Name = $Name; Path = $Path; BeforeBytes = 0L; AfterBytes = 0L; RemovedBytes = 0L; RemovedItems = 0 }
    }

    $before = Get-PathSizeBytes -Path $Path
    $cleanupParams = @{
        Path = $Path
        ExcludeLiteralPaths = $ExcludeLiteralPaths
    }
    if ($PSBoundParameters.ContainsKey('OlderThan')) {
        $cleanupParams['OlderThan'] = $OlderThan
    }
    $removedItems = Remove-PathChildrenForCleanup @cleanupParams
    $after = Get-PathSizeBytes -Path $Path
    $removedBytes = [math]::Max(0L, $before - $after)
    Write-SmartLog ("Disk cleanup area {0}: RemovedItems={1}; FreedGB={2}; Path={3}" -f $Name,$removedItems,(Convert-BytesToGbText -Bytes $removedBytes),$Path)
    return [pscustomobject]@{ Name = $Name; Path = $Path; BeforeBytes = $before; AfterBytes = $after; RemovedBytes = $removedBytes; RemovedItems = $removedItems }
}

function Get-CurrentSetupCachePathForCleanup {
    try {
        $expectedLanguage = Resolve-SetupLanguageRequirement -RequestedLanguage $SetupLanguage
        return Resolve-SetupCachePath -ExpectedLanguage $expectedLanguage
    }
    catch {
        Write-SmartLog ("Could not resolve current setup cache path for cleanup exclusion: {0}" -f $_.Exception.Message) 'WARN'
        return ''
    }
}

function Test-WindowsUpdateBusy {
    $busyProcesses = @('setup','setuphost','setupprep','TiWorker','TrustedInstaller','MoUsoCoreWorker','UsoClient')
    foreach ($name in $busyProcesses) {
        if (Get-Process -Name $name -ErrorAction SilentlyContinue) { return $true }
    }
    return $false
}

function Test-WindowsSetupOrUpgradeBusy {
    if (Test-WindowsUpdateBusy) { return $true }

    try {
        $setup = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\Setup' -ErrorAction Stop
        foreach ($name in @('SystemSetupInProgress','UpgradeInProgress','OOBEInProgress')) {
            if ($setup.PSObject.Properties[$name] -and [int]$setup.$name -ne 0) { return $true }
        }
    }
    catch { }

    return $false
}

function Invoke-SafeDiskCleanup {
    $beforeFree = Get-SystemDriveFreeGb
    Write-SmartLog ("Starting safe disk cleanup. FreeDiskGB before={0}." -f $beforeFree)

    $results = New-Object System.Collections.ArrayList
    $tempCutoff = (Get-Date).AddDays(-1 * $DiskCleanupTempFileMinAgeDays)
    $logCutoff = (Get-Date).AddDays(-1 * $DiskCleanupLogRetentionDays)

    $tempPaths = New-Object System.Collections.ArrayList
    foreach ($path in @($env:TEMP,$env:TMP,(Join-Path $env:SystemRoot 'Temp'))) {
        if (-not [string]::IsNullOrWhiteSpace($path) -and -not @($tempPaths.ToArray()).Contains($path)) {
            [void]$tempPaths.Add($path)
        }
    }
    foreach ($path in @($tempPaths.ToArray())) {
        [void]$results.Add((Invoke-CleanupArea -Name 'Temp' -Path $path -OlderThan $tempCutoff))
    }

    $deliveryOptimizationCache = Join-Path $env:SystemRoot 'ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache'
    [void]$results.Add((Invoke-CleanupArea -Name 'DeliveryOptimizationCache' -Path $deliveryOptimizationCache -OlderThan $tempCutoff))

    $wuDownload = Join-Path $env:SystemRoot 'SoftwareDistribution\Download'
    if (Test-WindowsUpdateBusy) {
        Write-SmartLog ("Disk cleanup skipped Windows Update download cache because update/setup activity appears active: {0}" -f $wuDownload) 'WARN'
    }
    else {
        [void]$results.Add((Invoke-CleanupArea -Name 'WindowsUpdateDownloadCache' -Path $wuDownload -OlderThan $tempCutoff))
    }

    $currentCache = Get-CurrentSetupCachePathForCleanup
    [void]$results.Add((Invoke-CleanupArea -Name 'OldSmartM365SetupMedia' -Path $SetupCacheRoot -ExcludeLiteralPaths @($currentCache)))
    [void]$results.Add((Invoke-CleanupArea -Name 'OldSmartM365Logs' -Path $script:LogDir -OlderThan $logCutoff -ExcludeLiteralPaths @($script:LogPath)))

    $afterFree = Get-SystemDriveFreeGb
    $freed = [math]::Max(0, [math]::Round(($afterFree - $beforeFree), 2))
    $script:DiskCleanupAction = ("SafeCleanup; BeforeGB={0}; AfterGB={1}; FreedGB={2}" -f $beforeFree,$afterFree,$freed)
    $script:DiskCleanupFreedGB = $freed
    Write-SmartLog ("Safe disk cleanup completed. FreeDiskGB after={0}; FreedGB={1}." -f $afterFree,$freed)
    return $afterFree
}

function Invoke-OldUpgradeFolderCleanup {
    $beforeFree = Get-SystemDriveFreeGb
    if (Test-WindowsSetupOrUpgradeBusy) {
        $script:AdvancedDiskCleanupAction = 'OldUpgradeFoldersSkipped; Reason=SetupOrUpdateBusy'
        Write-SmartLog 'Advanced cleanup skipped old Windows upgrade folders because setup/update activity appears active.' 'WARN'
        return $beforeFree
    }

    $systemDrive = ([System.IO.Path]::GetPathRoot($env:SystemRoot)).TrimEnd('\')
    $cutoff = (Get-Date).AddDays(-1 * $DiskCleanupUpgradeFolderMinAgeDays)
    $targets = @(
        (Join-Path $systemDrive '$WINDOWS.~BT'),
        (Join-Path $systemDrive '$WINDOWS.~WS'),
        (Join-Path $systemDrive 'Windows.old')
    )

    $removed = New-Object System.Collections.ArrayList
    foreach ($target in $targets) {
        if (-not (Test-Path -LiteralPath $target -PathType Container)) { continue }
        $item = Get-Item -LiteralPath $target -Force -ErrorAction SilentlyContinue
        if ($null -eq $item) { continue }
        if ($item.LastWriteTime -gt $cutoff) {
            Write-SmartLog ("Advanced cleanup skipped recent upgrade folder: {0}; LastWrite={1}" -f $target,$item.LastWriteTime) 'WARN'
            continue
        }

        $beforeBytes = Get-PathSizeBytes -Path $target
        try {
            Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction Stop
            [void]$removed.Add(("{0}:{1}GB" -f $target,(Convert-BytesToGbText -Bytes $beforeBytes)))
            Write-SmartLog ("Advanced cleanup removed old upgrade folder: {0}; FreedGBApprox={1}" -f $target,(Convert-BytesToGbText -Bytes $beforeBytes))
        }
        catch {
            Write-SmartLog ("Advanced cleanup could not remove old upgrade folder '{0}': {1}" -f $target,$_.Exception.Message) 'WARN'
        }
    }

    $afterFree = Get-SystemDriveFreeGb
    $freed = [math]::Max(0, [math]::Round(($afterFree - $beforeFree), 2))
    $details = if ($removed.Count -gt 0) { (@($removed.ToArray()) -join '; ') } else { 'NoOldUpgradeFoldersRemoved' }
    $script:AdvancedDiskCleanupAction = ("OldUpgradeFolders; BeforeGB={0}; AfterGB={1}; FreedGB={2}; Detail={3}" -f $beforeFree,$afterFree,$freed,$details)
    $script:AdvancedDiskCleanupFreedGB = $freed
    Write-SmartLog ("Advanced old upgrade folder cleanup completed. FreeDiskGB after={0}; FreedGB={1}." -f $afterFree,$freed)
    return $afterFree
}

function Invoke-DismComponentCleanup {
    $beforeFree = Get-SystemDriveFreeGb
    New-SmartDirectory -Path $script:SetupLogDir
    $stdout = Join-Path $script:SetupLogDir ("DISM_StartComponentCleanup_{0}.stdout.txt" -f $script:RunId)
    $stderr = Join-Path $script:SetupLogDir ("DISM_StartComponentCleanup_{0}.stderr.txt" -f $script:RunId)
    $dism = Join-Path $env:SystemRoot 'System32\dism.exe'
    if (-not (Test-Path -LiteralPath $dism -PathType Leaf)) {
        throw "dism.exe not found: $dism"
    }

    Write-SmartLog 'Starting DISM StartComponentCleanup.'
    $process = Start-Process -FilePath $dism -ArgumentList @('/Online','/Cleanup-Image','/StartComponentCleanup') -Wait -PassThru -NoNewWindow -RedirectStandardOutput $stdout -RedirectStandardError $stderr -ErrorAction Stop
    $afterFree = Get-SystemDriveFreeGb
    $freed = [math]::Max(0, [math]::Round(($afterFree - $beforeFree), 2))
    $script:DismCleanupAction = ("StartComponentCleanup; ExitCode={0}; BeforeGB={1}; AfterGB={2}; FreedGB={3}; Stdout={4}; Stderr={5}" -f $process.ExitCode,$beforeFree,$afterFree,$freed,$stdout,$stderr)
    $script:DismCleanupFreedGB = $freed
    Write-SmartLog ("DISM StartComponentCleanup completed. ExitCode={0}; FreeDiskGB after={1}; FreedGB={2}." -f $process.ExitCode,$afterFree,$freed)
    if ($process.ExitCode -ne 0) {
        throw "DISM StartComponentCleanup failed with exit code $($process.ExitCode). Stdout=$stdout; Stderr=$stderr"
    }
    return $afterFree
}

function Resolve-SetupUpgradeExecutable {
    if (-not ($AllowSetupUpgrade -or $DirectSetupUpgrade)) { return '' }

    $expectedLanguage = Resolve-SetupLanguageRequirement -RequestedLanguage $SetupLanguage
    $script:ResolvedSetupLanguage = $expectedLanguage

    $cachePath = Resolve-SetupCachePath -ExpectedLanguage $expectedLanguage
    $script:ResolvedSetupCachePath = $cachePath
    if ($SetupExecutionMode -in @('LocalCache','Auto')) {
        try {
            $script:SetupCacheAction = 'AlreadyCached'
            $cachedSetupExe = Test-SetupCacheReady -CachePath $cachePath -ExpectedLanguage $expectedLanguage
            if (-not $SkipSetupMediaPreCopy) {
                try {
                    $setupSourceCandidates = @(Get-EffectiveSetupSourceCandidates)
                    if ($setupSourceCandidates.Count -gt 0) {
                        $candidateSource = Resolve-PreferredSetupSourcePath -SourcePaths $setupSourceCandidates -ExpectedLanguage $expectedLanguage
                        $sourceFingerprint = Get-SetupMediaFingerprint -MediaRoot $candidateSource -ExpectedLanguage $expectedLanguage
                        $cacheFingerprint = Get-SetupMediaFingerprint -MediaRoot (Split-Path -Parent $cachedSetupExe) -ExpectedLanguage $expectedLanguage
                        if (-not (Test-SetupFingerprintMatch -Left $sourceFingerprint -Right $cacheFingerprint)) {
                            Write-SmartLog ("Setup source differs from local cache. Updating cache from '{0}'." -f $candidateSource) 'WARN'
                            return Copy-SetupMediaToLocalCache -SourcePath $candidateSource -CachePath $cachePath -ExpectedLanguage $expectedLanguage
                        }
                    }
                }
                catch {
                    Write-SmartLog ("Could not compare setup source with local cache. Using valid local cache. Detail={0}" -f $_.Exception.Message) 'WARN'
                }
            }
            return $cachedSetupExe
        }
        catch {
            $script:SetupCacheAction = 'CacheInvalid'
            if ($SkipSetupMediaPreCopy) {
                throw ("Local setup cache is not ready and existing-media-only mode is enabled. CachePath={0}; Error={1}" -f $cachePath,$_.Exception.Message)
            }
            Write-SmartLog ("Local setup cache not ready: {0}" -f $_.Exception.Message) 'WARN'
            Clear-SetupCachePath -CachePath $cachePath -Reason $_.Exception.Message
            $setupSourceCandidates = @(Get-EffectiveSetupSourceCandidates)
            if ($SetupExecutionMode -eq 'LocalCache' -and $setupSourceCandidates.Count -eq 0) {
                throw ("Local setup cache is not ready and no SetupSourcePath was provided. CachePath={0}; Error={1}" -f $cachePath,$_.Exception.Message)
            }
        }
    }

    $setupSourceCandidates = @(Get-EffectiveSetupSourceCandidates)
    if ($setupSourceCandidates.Count -eq 0) {
        throw 'SetupSourcePath is required when -AllowSetupUpgrade is used and no valid local cache exists.'
    }

    $expandedSource = Resolve-PreferredSetupSourcePath -SourcePaths $setupSourceCandidates -ExpectedLanguage $expectedLanguage
    if ($SetupExecutionMode -in @('Share','Auto')) {
        try {
            return Test-SetupMedia -MediaPath $expandedSource -ExpectedLanguage $expectedLanguage -ValidateManifest
        }
        catch {
            if ($SetupExecutionMode -eq 'Share') { throw }
            Write-SmartLog ("Setup share source not directly usable from target context: {0}" -f $_.Exception.Message) 'WARN'
        }
    }

    if ($SetupExecutionMode -in @('LocalCache','Auto')) {
        return Copy-SetupMediaToLocalCache -SourcePath $expandedSource -CachePath $cachePath -ExpectedLanguage $expectedLanguage
    }

    throw 'Unable to resolve a valid setup.exe path.'
}

function Format-SetupProcessSnapshot {
    param([AllowNull()]$Process)

    if ($null -eq $Process) { return 'PID=<unknown>' }

    try { $Process.Refresh() } catch { }

    $startTime = ''
    try { $startTime = $Process.StartTime.ToString('yyyy-MM-dd HH:mm:ss') } catch { $startTime = '<unavailable>' }

    $cpuSeconds = ''
    try { $cpuSeconds = [math]::Round([double]$Process.TotalProcessorTime.TotalSeconds, 1).ToString([System.Globalization.CultureInfo]::InvariantCulture) } catch { $cpuSeconds = '<unavailable>' }

    $workingSetMb = ''
    try { $workingSetMb = [math]::Round(($Process.WorkingSet64 / 1MB), 1).ToString([System.Globalization.CultureInfo]::InvariantCulture) } catch { $workingSetMb = '<unavailable>' }

    $hasExited = ''
    try { $hasExited = [string]$Process.HasExited } catch { $hasExited = '<unavailable>' }

    return ("PID={0}; HasExited={1}; StartTime={2}; CPUSeconds={3}; WorkingSetMB={4}" -f $Process.Id,$hasExited,$startTime,$cpuSeconds,$workingSetMb)
}

function Invoke-SetupUpgrade {
    param([Parameter(Mandatory = $true)][string]$SetupExePath)

    New-SmartDirectory -Path $script:SetupLogDir
    $args = @(
        '/auto','upgrade',
        '/quiet',
        '/eula','accept',
        '/dynamicupdate',$SetupDynamicUpdate.ToLowerInvariant(),
        '/copylogs',"`"$script:SetupLogDir`""
    )

    Write-SmartLog ("Starting setup upgrade: {0} {1}" -f $SetupExePath,($args -join ' '))
    $process = Start-Process -FilePath $SetupExePath -ArgumentList $args -PassThru -ErrorAction Stop
    $startedAt = Get-Date
    $lastHeartbeatAt = $startedAt
    $heartbeatSeconds = [math]::Max(30, [int]$SetupProcessHeartbeatSeconds)
    $timeoutAt = if ($SetupProcessTimeoutMinutes -gt 0) { $startedAt.AddMinutes($SetupProcessTimeoutMinutes) } else { $null }

    Write-SmartLog ("setup.exe started. {0}; HeartbeatSeconds={1}; TimeoutMinutes={2}" -f (Format-SetupProcessSnapshot -Process $process),$heartbeatSeconds,$SetupProcessTimeoutMinutes)

    while ($true) {
        try { $process.Refresh() } catch { }

        if ($process.HasExited) { break }

        $now = Get-Date
        if ($timeoutAt -and $now -ge $timeoutAt) {
            Write-SmartLog ("setup.exe timeout reached after {0} minute(s). {1}" -f $SetupProcessTimeoutMinutes,(Format-SetupProcessSnapshot -Process $process)) 'ERROR'
            try { Stop-Process -Id $process.Id -Force -ErrorAction Stop } catch { Write-SmartLog ("Failed to stop timed-out setup.exe PID={0}: {1}" -f $process.Id,$_.Exception.Message) 'WARN' }
            throw ("setup.exe timed out after {0} minute(s)." -f $SetupProcessTimeoutMinutes)
        }

        if (($now - $lastHeartbeatAt).TotalSeconds -ge $heartbeatSeconds) {
            Write-SmartLog ("setup.exe still running. ElapsedMinutes={0}; {1}" -f ([math]::Round(($now - $startedAt).TotalMinutes, 1)),(Format-SetupProcessSnapshot -Process $process))
            $lastHeartbeatAt = $now
        }

        Start-Sleep -Seconds 5
    }

    $setupExitInfo = Get-SetupExitCodeInfo -ExitCode $process.ExitCode
    Write-SmartLog ("setup.exe exited. ElapsedMinutes={0}; {1}; {2}" -f ([math]::Round(((Get-Date) - $startedAt).TotalMinutes, 1)),(Format-SetupProcessSnapshot -Process $process),(Format-SetupExitCodeInfo -Info $setupExitInfo))
    return $process.ExitCode
}

function Get-InteractiveUserSessionSummary {
    $excludedDomains = @('NT AUTHORITY','WINDOW MANAGER','FONT DRIVER HOST')
    $excludedNames = @('SYSTEM','LOCAL SERVICE','NETWORK SERVICE')

    try {
        $sessions = @(Get-CimInstance -ClassName Win32_LogonSession -Filter 'LogonType = 2 OR LogonType = 10 OR LogonType = 11' -ErrorAction Stop)
        $users = New-Object System.Collections.ArrayList

        foreach ($session in $sessions) {
            $accounts = @(Get-CimAssociatedInstance -InputObject $session -Association Win32_LoggedOnUser -ErrorAction SilentlyContinue)
            foreach ($account in $accounts) {
                $domain = ([string]$account.Domain).Trim()
                $name = ([string]$account.Name).Trim()
                if ([string]::IsNullOrWhiteSpace($name)) { continue }
                if ($excludedDomains -contains $domain.ToUpperInvariant()) { continue }
                if ($excludedNames -contains $name.ToUpperInvariant()) { continue }
                if ($name -match '^(DWM|UMFD)-\d+$') { continue }
                if ($name.EndsWith('$')) { continue }

                $display = if ([string]::IsNullOrWhiteSpace($domain)) { $name } else { "{0}\{1}" -f $domain,$name }
                if (-not $users.Contains($display)) { [void]$users.Add($display) }
            }
        }

        $userList = (@($users.ToArray()) -join '; ')
        return [pscustomobject]@{
            DetectionSucceeded = $true
            InteractiveUserCount = $users.Count
            InteractiveUsers = $userList
            Detail = ("InteractiveUserCount={0}; InteractiveUsers={1}" -f $users.Count,$userList)
        }
    }
    catch {
        return [pscustomobject]@{
            DetectionSucceeded = $false
            InteractiveUserCount = ''
            InteractiveUsers = ''
            Detail = ("Interactive user detection failed: {0}" -f $_.Exception.Message)
        }
    }
}

function ConvertFrom-SmartUnicodeEscape {
    param([AllowNull()][string]$Text)
    if ($null -eq $Text) { return '' }
    return [regex]::Replace($Text, '\\u([0-9A-Fa-f]{4})', {
        param($match)
        [string][char]([Convert]::ToInt32($match.Groups[1].Value, 16))
    })
}

function Send-UserRebootNotification {
    param([ValidateSet('UpgradeReady', 'PendingReboot')][string]$Context = 'UpgradeReady')
    $lang = try { ((Get-WinUserLanguageList)[0].LanguageTag -split '-')[0].ToLower() } catch { (Get-UICulture).TwoLetterISOLanguageName.ToLower() }
    $messagesByContext = @{
        'PendingReboot' = @{
            'fr' = 'Un red\u00E9marrage est requis avant de poursuivre la mise \u00E0 niveau Windows 11. Veuillez enregistrer votre travail et red\u00E9marrer d\u00E8s que possible.'
            'es' = 'Se requiere un reinicio antes de continuar con la actualizaci\u00F3n de Windows 11. Guarde su trabajo y reinicie cuanto antes.'
            'de' = 'Ein Neustart ist erforderlich, bevor das Windows 11-Upgrade fortgesetzt werden kann. Bitte speichern Sie Ihre Arbeit und starten Sie den Computer neu.'
            'it' = '\u00C8 necessario riavviare prima di continuare con l''aggiornamento a Windows 11. Salvare il lavoro e riavviare il computer.'
            'nl' = 'Er is een herstart vereist voordat het Windows 11-upgrade kan worden voortgezet. Sla uw werk op en start opnieuw op.'
            'pt' = '\u00C9 necess\u00E1rio reiniciar antes de continuar com a atualiza\u00E7\u00E3o para o Windows 11. Guarde o seu trabalho e reinicie o computador.'
            'pl' = 'Wymagane jest ponowne uruchomienie przed kontynuowaniem aktualizacji systemu Windows 11. Zapisz prac\u0119 i uruchom ponownie komputer.'
        }
        'UpgradeReady' = @{
            'fr' = 'La mise \u00E0 niveau Windows 11 est pr\u00EAte. Veuillez enregistrer votre travail et red\u00E9marrer d\u00E8s que possible.'
            'es' = 'La actualizaci\u00F3n a Windows 11 est\u00E1 lista. Guarde su trabajo y reinicie cuanto antes.'
            'de' = 'Das Windows 11-Upgrade ist bereit. Bitte speichern Sie Ihre Arbeit und starten Sie den Computer neu.'
            'it' = 'L''aggiornamento a Windows 11 \u00E8 pronto. Salvare il lavoro e riavviare il computer.'
            'nl' = 'De Windows 11-upgrade is gereed. Sla uw werk op en start opnieuw op.'
            'pt' = 'A atualiza\u00E7\u00E3o para o Windows 11 est\u00E1 pronta. Guarde o seu trabalho e reinicie o computador.'
            'pl' = 'Aktualizacja Windows 11 jest gotowa. Zapisz prac\u0119 i uruchom ponownie komputer.'
        }
    }
    $defaultMessages = @{
        'PendingReboot' = "A restart is required before the Windows 11 upgrade can continue. Please save your work and restart as soon as possible."
        'UpgradeReady'  = "Windows 11 upgrade is ready. Please save your work and restart as soon as possible."
    }
    $messages = $messagesByContext[$Context]
    $messageTemplate = if ($messages.ContainsKey($lang)) { $messages[$lang] } else { $defaultMessages[$Context] }
    $message = ConvertFrom-SmartUnicodeEscape -Text $messageTemplate
    try {
        & "$env:SystemRoot\System32\msg.exe" * /time:300 $message 2>$null | Out-Null
        $script:UserRebootNotificationSent = 'True'
        $script:UserRebootNotificationLang = $lang
        $script:UserRebootNotificationMessage = $message
        Write-SmartLog ("User reboot notification sent via msg.exe (lang={0}; context={1}): {2}" -f $lang,$Context,$message)
    }
    catch {
        Write-SmartLog ("Failed to send user reboot notification: {0}" -f $_.Exception.Message) 'WARN'
    }
}

function Invoke-ControlledRebootWhenNoUser {
    param([Parameter(Mandatory = $true)][string]$Reason)

    $summary = Get-InteractiveUserSessionSummary
    $script:ControlledRebootUserCount = [string]$summary.InteractiveUserCount
    $script:ControlledRebootUsers = [string]$summary.InteractiveUsers
    $script:ControlledRebootDetail = [string]$summary.Detail

    if (-not $summary.DetectionSucceeded) {
        $script:ControlledRebootAction = 'SkippedUserDetectionFailed'
        Write-SmartLog ("Controlled reboot skipped because user detection failed: {0}" -f $summary.Detail) 'WARN'
        return 'UserDetectionFailed'
    }

    if ([int]$summary.InteractiveUserCount -gt 0) {
        $script:ControlledRebootAction = 'SkippedUserConnected'
        Write-SmartLog ("Controlled reboot skipped because interactive user(s) are connected: {0}. A pending reboot must be cleared before the Windows 11 upgrade can continue; it will retry on a later cycle once no user is connected." -f $summary.InteractiveUsers) 'WARN'
        Send-UserRebootNotification -Context PendingReboot
        return 'UserConnected'
    }

    try {
        shutdown.exe /r /t $RebootDelaySeconds /c $Reason | Out-Null
        $script:ControlledRebootAction = 'ScheduledNoUser'
        $script:ControlledRebootDetail = ("No interactive user connected. Reboot scheduled in {0} second(s). Reason={1}" -f $RebootDelaySeconds,$Reason)
        Write-SmartLog $script:ControlledRebootDetail
        return 'ScheduledNoUser'
    }
    catch {
        $script:ControlledRebootAction = 'ScheduleFailed'
        $script:ControlledRebootDetail = ("No interactive user connected, but reboot scheduling failed: {0}" -f $_.Exception.Message)
        Write-SmartLog $script:ControlledRebootDetail 'ERROR'
        return 'ScheduleFailed'
    }
}
function Invoke-SetupCompletionRebootWhenNoUser {
    if (-not $AllowSetupCompletionRebootWhenNoUser) { return 'NotRequested' }

    $summary = Get-InteractiveUserSessionSummary
    $script:SetupCompletionRebootUserCount = [string]$summary.InteractiveUserCount
    $script:SetupCompletionRebootUsers = [string]$summary.InteractiveUsers
    $script:SetupCompletionRebootDetail = [string]$summary.Detail

    if (-not $summary.DetectionSucceeded) {
        $script:SetupCompletionRebootAction = 'SkippedUserDetectionFailed'
        Write-SmartLog ("Setup completion reboot skipped because user detection failed: {0}" -f $summary.Detail) 'WARN'
        return 'UserDetectionFailed'
    }

    if ([int]$summary.InteractiveUserCount -gt 0) {
        $script:SetupCompletionRebootAction = 'SkippedUserConnected'
        Write-SmartLog ("Setup completion reboot skipped because interactive user(s) are connected: {0}" -f $summary.InteractiveUsers) 'WARN'
        Send-UserRebootNotification
        return 'UserConnected'
    }

    try {
        shutdown.exe /r /t $RebootDelaySeconds /c 'SmartM365 Windows 11 setup completion reboot - no interactive user connected' | Out-Null
        $script:SetupCompletionRebootAction = 'ScheduledNoUser'
        $script:SetupCompletionRebootDetail = ("No interactive user connected. Reboot scheduled in {0} second(s)." -f $RebootDelaySeconds)
        Write-SmartLog $script:SetupCompletionRebootDetail
        return 'ScheduledNoUser'
    }
    catch {
        $script:SetupCompletionRebootAction = 'ScheduleFailed'
        $script:SetupCompletionRebootDetail = ("No interactive user connected, but reboot scheduling failed: {0}" -f $_.Exception.Message)
        Write-SmartLog $script:SetupCompletionRebootDetail 'ERROR'
        return 'ScheduleFailed'
    }
}

function Resolve-SetupUpgradeSuccessOutcome {
    param(
        [Parameter(Mandatory = $true)][string]$Prefix,
        [Parameter(Mandatory = $true)][int]$SetupExitCode
    )

    $baseStatus = if ($SetupExitCode -eq 3010) { "${Prefix}_REBOOT_REQUIRED" } else { "${Prefix}_STARTED" }
    $baseNextAction = if ($SetupExitCode -eq 3010) { 'REBOOT_DEVICE' } else { 'MONITOR_SETUP_AND_REBOOT' }
    $suffix = ''

    if (-not $AllowSetupCompletionRebootWhenNoUser) {
        return [pscustomobject]@{ Status = $baseStatus; NextAction = $baseNextAction; ExitCode = 0; ActionResultSuffix = $suffix }
    }

    $rebootResult = Invoke-SetupCompletionRebootWhenNoUser
    switch ($rebootResult) {
        'ScheduledNoUser' {
            return [pscustomobject]@{ Status = "${Prefix}_REBOOT_SCHEDULED_NO_USER"; NextAction = 'REBOOT_SCHEDULED'; ExitCode = 0; ActionResultSuffix = 'SetupCompletionRebootScheduledNoUser' }
        }
        'UserConnected' {
            return [pscustomobject]@{ Status = "${Prefix}_REBOOT_SKIPPED_USER_CONNECTED"; NextAction = 'WAIT_USER_LOGOFF_OR_REBOOT_DEVICE'; ExitCode = 0; ActionResultSuffix = 'SetupCompletionRebootSkippedUserConnected' }
        }
        'UserDetectionFailed' {
            return [pscustomobject]@{ Status = "${Prefix}_REBOOT_SKIPPED_USER_DETECTION_FAILED"; NextAction = 'CHECK_USER_SESSIONS_BEFORE_REBOOT'; ExitCode = 0; ActionResultSuffix = 'SetupCompletionRebootSkippedUserDetectionFailed' }
        }
        'ScheduleFailed' {
            return [pscustomobject]@{ Status = "${Prefix}_REBOOT_SCHEDULE_FAILED"; NextAction = 'CHECK_REBOOT_SCHEDULE'; ExitCode = 1; ActionResultSuffix = 'SetupCompletionRebootScheduleFailed' }
        }
        default {
            return [pscustomobject]@{ Status = $baseStatus; NextAction = $baseNextAction; ExitCode = 0; ActionResultSuffix = $suffix }
        }
    }
}

function Save-RunResult {
    param([Parameter(Mandatory = $true)]$Result)

    New-SmartDirectory -Path $script:OutputDir
    $Result | Export-Csv -LiteralPath $script:CsvPath -NoTypeInformation -Encoding UTF8
    $Result | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $script:LastRunPath -Encoding UTF8
}

New-SmartDirectory -Path $script:LogDir
New-SmartDirectory -Path $script:OutputDir
$script:LocalIPv4Addresses = ''
try { $script:LocalIPv4Addresses = (@(Get-LocalIPv4Addresses) -join ',') } catch { $script:LocalIPv4Addresses = '' }
Write-SmartLog ("===== {0} v{1} started. ComputerName={2}; LocalIPv4={3}; RunId={4} =====" -f $script:ScriptName,$script:ScriptVersion,$script:ComputerName,$script:LocalIPv4Addresses,$script:RunId)

$exitCode = 3
$status = 'UNKNOWN'
$nextAction = 'CHECK_LOGS'
$detail = ''
$actionResult = ''
$setupExe = ''
$script:ResolvedSetupLanguage = ''
$script:ResolvedSetupMediaLanguages = ''
$script:ResolvedSetupCachePath = ''
$script:SelectedSetupSourcePath = ''
$script:SetupSourceSelectionDetail = ''
$script:SetupCacheAction = ''
$script:DiskCleanupAction = ''
$script:DiskCleanupFreedGB = ''
$script:AdvancedDiskCleanupAction = ''
$script:AdvancedDiskCleanupFreedGB = ''
$script:DismCleanupAction = ''
$script:DismCleanupFreedGB = ''
$script:SetupCompletionRebootAction = ''
$script:SetupCompletionRebootDetail = ''
$script:SetupCompletionRebootUserCount = ''
$script:SetupCompletionRebootUsers = ''
$script:ControlledRebootAction = ''
$script:ControlledRebootDetail = ''
$script:ControlledRebootUserCount = ''
$script:ControlledRebootUsers = ''
$script:UserRebootNotificationSent = ''
$script:UserRebootNotificationLang = ''
$script:UserRebootNotificationMessage = ''
$computerSystem = $null

try {
    if (-not $IgnoreRunGuard -and $RunGuardHours -gt 0 -and (Test-Path -LiteralPath $script:LastRunPath -PathType Leaf)) {
        $last = Get-Content -LiteralPath $script:LastRunPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($last.EndTimeUtc) {
            $lastEndUtc = [datetime]::Parse([string]$last.EndTimeUtc, $null, [System.Globalization.DateTimeStyles]::RoundtripKind)
            $age = (Get-Date).ToUniversalTime() - $lastEndUtc
            if ($age.TotalHours -lt $RunGuardHours) {
                $status = 'RUN_GUARD_ACTIVE'
                $nextAction = 'WAIT_OR_USE_IGNORE_RUN_GUARD'
                $detail = ("Last run age {0:N1} hour(s), guard {1} hour(s)." -f $age.TotalHours,$RunGuardHours)
                $exitCode = 3
                throw [System.OperationCanceledException]::new($detail)
            }
        }
    }

    $computerSystem = Get-ComputerSystemSummary
    if ($SkipVirtualMachines -and $computerSystem.IsVirtualMachine) {
        $status = 'SKIPPED_VIRTUAL_MACHINE'
        $nextAction = 'NO_ACTION_VIRTUAL_MACHINE'
        $detail = $computerSystem.Evidence
        $exitCode = 0
        throw [System.OperationCanceledException]::new($detail)
    }

    $os = Get-OsSummary
    Write-SmartLog ("Startup OS before upgrade: ComputerName={0}; LocalIPv4={1}; Caption={2}; Version={3}; Build={4}; Architecture={5}; Family={6}" -f $script:ComputerName,$script:LocalIPv4Addresses,$os.Caption,$os.Version,$os.BuildNumber,$os.Architecture,$os.MajorFamily)
    $freeGb = Get-SystemDriveFreeGb
    $pendingRebootInfo = Test-PendingReboot
    $pendingReboot = $pendingRebootInfo.IsPending
    if ($pendingReboot) {
        Write-SmartLog ("Pending reboot detected before upgrade. Source(s)={0}" -f $pendingRebootInfo.Source) 'WARN'
    }
    $intune = Get-IntuneEnrollmentSummary
    $policy = Get-WindowsUpdatePolicySummary
    $indicators = Get-Windows11IndicatorSummary

    $diskCleanupEligible = ($os.MajorFamily -eq 'Windows10' -and -not $DirectSetupUpgrade -and $intune.IsIntuneEnrolled -and -not $indicators.ActionableBlocking)
    if ($diskCleanupEligible -and $freeGb -lt $MinimumFreeDiskGB -and $AllowDiskCleanup -and -not $AuditOnly) {
        try {
            $freeGb = Invoke-SafeDiskCleanup
        }
        catch {
            $script:DiskCleanupAction = ("SafeCleanupFailed; Error={0}" -f $_.Exception.Message)
            Write-SmartLog ("Safe disk cleanup failed: {0}" -f $_.Exception.Message) 'WARN'
            $freeGb = Get-SystemDriveFreeGb
        }

        if ($freeGb -lt $MinimumFreeDiskGB -and ($AllowAdvancedDiskCleanup -or $AllowDismComponentCleanup)) {
            try {
                $freeGb = Invoke-OldUpgradeFolderCleanup
            }
            catch {
                $script:AdvancedDiskCleanupAction = ("OldUpgradeFoldersFailed; Error={0}" -f $_.Exception.Message)
                Write-SmartLog ("Advanced old upgrade folder cleanup failed: {0}" -f $_.Exception.Message) 'WARN'
                $freeGb = Get-SystemDriveFreeGb
            }
        }

        if ($freeGb -lt $MinimumFreeDiskGB -and ($AllowAdvancedDiskCleanup -or $AllowDismComponentCleanup)) {
            try {
                $freeGb = Invoke-DismComponentCleanup
            }
            catch {
                $script:DismCleanupAction = ("StartComponentCleanupFailed; Error={0}" -f $_.Exception.Message)
                Write-SmartLog ("DISM component cleanup failed: {0}" -f $_.Exception.Message) 'WARN'
                $freeGb = Get-SystemDriveFreeGb
            }
        }
    }

    if ($os.MajorFamily -eq 'Windows11') {
        $status = 'ALREADY_WINDOWS11'
        $nextAction = 'NO_ACTION_ALREADY_WINDOWS11'
        $exitCode = 0
    }
    elseif ($os.MajorFamily -ne 'Windows10') {
        $status = 'UNSUPPORTED_OS'
        $nextAction = 'CHECK_DEVICE_SCOPE'
        $detail = "OS family is $($os.MajorFamily)."
        $exitCode = 3
    }
    elseif ($DirectSetupUpgrade) {
        Write-SmartLog 'Direct setup upgrade requested. Skipping Intune enrollment, compatibility-indicator, pending-reboot, and policy-blocker gates; Windows Setup will perform final validation.' 'WARN'
        if ($freeGb -lt $MinimumFreeDiskGB) {
            $status = 'INSUFFICIENT_DISK'
            $nextAction = 'FREE_DISK_SPACE'
            $detail = ("FreeDiskGB={0}; RequiredGB={1}; DirectSetup=True; Setup media copy was not attempted." -f $freeGb,$MinimumFreeDiskGB)
            $exitCode = 3
        }
        else {
            $setupExe = Resolve-SetupUpgradeExecutable
            if ($AuditOnly) {
                $status = 'DIRECT_SETUP_UPGRADE_READY'
                $nextAction = 'RUN_WITHOUT_AUDIT_ONLY'
                $actionResult = 'DirectSetupValidatedAuditOnly'
                $exitCode = 0
            }
            else {
                $setupExitCode = Invoke-SetupUpgrade -SetupExePath $setupExe
                $setupExitInfo = Get-SetupExitCodeInfo -ExitCode $setupExitCode
                $setupExitDetail = Format-SetupExitCodeInfo -Info $setupExitInfo
                $actionResult = "DirectSetup$setupExitDetail"
                if ($setupExitCode -eq 0 -or $setupExitCode -eq 3010) {
                    $outcome = Resolve-SetupUpgradeSuccessOutcome -Prefix 'DIRECT_SETUP_UPGRADE' -SetupExitCode $setupExitCode
                    $status = $outcome.Status
                    $nextAction = $outcome.NextAction
                    $exitCode = $outcome.ExitCode
                    if (-not [string]::IsNullOrWhiteSpace([string]$outcome.ActionResultSuffix)) { $actionResult = "$actionResult;$($outcome.ActionResultSuffix)" }
                }
                else {
                    $status = 'DIRECT_SETUP_UPGRADE_FAILED'
                    $nextAction = 'CHECK_SETUP_LOGS'
                    $detail = $setupExitDetail
                    $exitCode = 1
                }
            }
        }
    }
    elseif (-not $intune.IsIntuneEnrolled) {
        $status = 'NOT_INTUNE_ENROLLED'
        $nextAction = 'FIX_INTUNE_ENROLLMENT_FIRST'
        $detail = $intune.WeakEvidence
        $exitCode = 3
    }
    elseif ($indicators.ActionableBlocking) {
        $status = 'WINDOWS11_COMPAT_BLOCKER'
        $nextAction = 'REVIEW_COMPATIBILITY_OR_SETUPDIAG'
        $detail = $indicators.BlockingReasons
        $exitCode = 3
    }
    elseif ($freeGb -lt $MinimumFreeDiskGB) {
        $status = if ($script:DiskCleanupAction -or $script:AdvancedDiskCleanupAction -or $script:DismCleanupAction) { 'INSUFFICIENT_DISK_AFTER_CLEANUP' } else { 'INSUFFICIENT_DISK' }
        $nextAction = 'FREE_DISK_SPACE'
        $detail = ("FreeDiskGB={0}; RequiredGB={1}; DiskCleanup={2}; AdvancedCleanup={3}; DismCleanup={4}" -f $freeGb,$MinimumFreeDiskGB,$script:DiskCleanupAction,$script:AdvancedDiskCleanupAction,$script:DismCleanupAction)
        $exitCode = 3
    }
    elseif ($pendingReboot) {
        $status = 'PENDING_REBOOT'
        $nextAction = if ($AllowReboot) { 'REBOOT_SCHEDULED' } else { 'REBOOT_DEVICE' }
        $detail = ("A reboot is pending and must be cleared before the Windows 11 upgrade can continue. Source(s)={0}" -f $pendingRebootInfo.Source)
        if ($AllowReboot -and -not $AuditOnly) {
            $rebootResult = Invoke-ControlledRebootWhenNoUser -Reason 'SmartM365 Windows 11 upgrade readiness reboot - no interactive user connected'
            switch ($rebootResult) {
                'ScheduledNoUser' {
                    $actionResult = 'RebootScheduledNoUser'
                    $exitCode = 0
                }
                'UserConnected' {
                    $status = 'PENDING_REBOOT_USER_CONNECTED'
                    $nextAction = 'WAIT_USER_LOGOFF_OR_REBOOT_DEVICE'
                    $actionResult = 'RebootSkippedUserConnected'
                    $exitCode = 0
                }
                'UserDetectionFailed' {
                    $status = 'PENDING_REBOOT_USER_DETECTION_FAILED'
                    $nextAction = 'CHECK_USER_SESSIONS_BEFORE_REBOOT'
                    $actionResult = 'RebootSkippedUserDetectionFailed'
                    $exitCode = 3
                }
                'ScheduleFailed' {
                    $status = 'PENDING_REBOOT_SCHEDULE_FAILED'
                    $nextAction = 'CHECK_REBOOT_SCHEDULE'
                    $actionResult = 'RebootScheduleFailed'
                    $exitCode = 1
                }
                default {
                    $exitCode = 3
                }
            }
        }
        else {
            $exitCode = 3
        }
    }
    elseif ($policy.HasLegacyBlocker) {
        $status = 'WU_POLICY_BLOCKER'
        $nextAction = if ($AllowPolicyRepair) { 'POLICY_REPAIR_ATTEMPTED' } else { 'ALLOW_POLICY_REPAIR' }
        $detail = $policy.Issues
        if ($AllowPolicyRepair -and -not $AuditOnly) {
            Repair-WindowsUpdatePolicyBlockers
            $actionResult = 'PolicyRepairCompleted'
            $exitCode = 0
        }
        else {
            $exitCode = 3
        }
    }
    elseif ($AllowSetupUpgrade) {
        $setupExe = Resolve-SetupUpgradeExecutable
        if ($AuditOnly) {
            $status = 'SETUP_UPGRADE_READY'
            $nextAction = 'RUN_WITHOUT_AUDIT_ONLY'
            $actionResult = 'SetupValidatedAuditOnly'
            $exitCode = 0
        }
        else {
            $setupExitCode = Invoke-SetupUpgrade -SetupExePath $setupExe
            $setupExitInfo = Get-SetupExitCodeInfo -ExitCode $setupExitCode
            $setupExitDetail = Format-SetupExitCodeInfo -Info $setupExitInfo
            $actionResult = "Setup$setupExitDetail"
            if ($setupExitCode -eq 0 -or $setupExitCode -eq 3010) {
                $outcome = Resolve-SetupUpgradeSuccessOutcome -Prefix 'SETUP_UPGRADE' -SetupExitCode $setupExitCode
                $status = $outcome.Status
                $nextAction = $outcome.NextAction
                $exitCode = $outcome.ExitCode
                if (-not [string]::IsNullOrWhiteSpace([string]$outcome.ActionResultSuffix)) { $actionResult = "$actionResult;$($outcome.ActionResultSuffix)" }
            }
            else {
                $status = 'SETUP_UPGRADE_FAILED'
                $nextAction = 'CHECK_SETUP_LOGS'
                $detail = $setupExitDetail
                $exitCode = 1
            }
        }
    }
    elseif ($AllowWUReset -and -not $AuditOnly) {
        Reset-WindowsUpdateComponents
        $status = 'WU_RESET_COMPLETED'
        $nextAction = 'RETRY_WINDOWS_UPDATE'
        $actionResult = 'WUResetCompleted'
        $exitCode = 0
    }
    elseif ($AllowForceUpgrade -and -not $AuditOnly) {
        $actionResult = Invoke-AssignedUpdateInstall
        $status = 'WINDOWS_UPDATE_FORCE_TRIGGERED'
        $nextAction = 'MONITOR_INTUNE_WINDOWS_UPDATE_STATUS'
        $exitCode = 0
    }
    else {
        $status = 'READY_TO_FORCE_UPGRADE'
        $nextAction = 'USE_ALLOW_FORCEUPGRADE_OR_ALLOWSETUPUPGRADE'
        $exitCode = 3
    }
}
catch [System.OperationCanceledException] {
    Write-SmartLog $_.Exception.Message 'WARN'
}
catch {
    if ($status -eq 'UNKNOWN') { $status = 'ERROR' }
    if ($nextAction -eq 'CHECK_LOGS') { $nextAction = 'CHECK_SCRIPT_LOG' }
    $detail = $_.Exception.Message
    Write-SmartLog $detail 'ERROR'
    $exitCode = if ($status -eq 'RUN_GUARD_ACTIVE') { 3 } else { 1 }
}
finally {
    $osFinal = $null
    try { $osFinal = Get-OsSummary } catch { }
    $finalFreeDisk = ''
    $finalPendingReboot = ''
    $finalPendingRebootSource = ''
    $finalIntuneEnrolled = ''
    $finalWuBlockers = ''
    $finalW11BlockingReasons = ''
    try { $finalFreeDisk = Get-SystemDriveFreeGb } catch { }
    try {
        $finalPendingRebootInfo = Test-PendingReboot
        $finalPendingReboot = $finalPendingRebootInfo.IsPending
        $finalPendingRebootSource = $finalPendingRebootInfo.Source
    }
    catch { }
    try { $finalIntuneEnrolled = (Get-IntuneEnrollmentSummary).IsIntuneEnrolled } catch { }
    try { $finalWuBlockers = (Get-WindowsUpdatePolicySummary).Issues } catch { }
    try { $finalW11BlockingReasons = (Get-Windows11IndicatorSummary).BlockingReasons } catch { }
    if ($null -eq $computerSystem) {
        try { $computerSystem = Get-ComputerSystemSummary } catch { }
    }

    $result = [pscustomobject]@{
        RunId = $script:RunId
        ComputerName = $script:ComputerName
        StartTimeUtc = ''
        EndTimeUtc = (Get-Date).ToUniversalTime().ToString('o')
        ScriptVersion = $script:ScriptVersion
        Status = $status
        NextAction = $nextAction
        Detail = $detail
        ActionResult = $actionResult
        ExitCode = $exitCode
        OSFamily = if ($osFinal) { $osFinal.MajorFamily } else { '' }
        OSCaption = if ($osFinal) { $osFinal.Caption } else { '' }
        OSVersion = if ($osFinal) { $osFinal.Version } else { '' }
        OSBuild = if ($osFinal) { $osFinal.BuildNumber } else { '' }
        FreeDiskGB = $finalFreeDisk
        PendingReboot = $finalPendingReboot
        PendingRebootSource = $finalPendingRebootSource
        IsVirtualMachine = if ($computerSystem) { $computerSystem.IsVirtualMachine } else { '' }
        VirtualMachineEvidence = if ($computerSystem) { $computerSystem.Evidence } else { '' }
        IntuneEnrolled = $finalIntuneEnrolled
        WUBlockers = $finalWuBlockers
        W11BlockingReasons = $finalW11BlockingReasons
        SetupExecutionMode = $SetupExecutionMode
        SetupMediaId = $SetupMediaId
        SetupLanguage = $SetupLanguage
        SetupDynamicUpdate = $SetupDynamicUpdate
        ResolvedSetupLanguage = $script:ResolvedSetupLanguage
        SetupMediaLanguages = $script:ResolvedSetupMediaLanguages
        SetupCachePath = $script:ResolvedSetupCachePath
        SelectedSetupSourcePath = $script:SelectedSetupSourcePath
        SetupSourceSelectionDetail = $script:SetupSourceSelectionDetail
        SetupCacheAction = $script:SetupCacheAction
        DiskCleanupAction = $script:DiskCleanupAction
        DiskCleanupFreedGB = $script:DiskCleanupFreedGB
        AdvancedDiskCleanupAction = $script:AdvancedDiskCleanupAction
        AdvancedDiskCleanupFreedGB = $script:AdvancedDiskCleanupFreedGB
        DismCleanupAction = $script:DismCleanupAction
        DismCleanupFreedGB = $script:DismCleanupFreedGB
        SetupCompletionRebootAction = $script:SetupCompletionRebootAction
        SetupCompletionRebootDetail = $script:SetupCompletionRebootDetail
        SetupCompletionRebootUserCount = $script:SetupCompletionRebootUserCount
        SetupCompletionRebootUsers = $script:SetupCompletionRebootUsers
        ControlledRebootAction = $script:ControlledRebootAction
        ControlledRebootDetail = $script:ControlledRebootDetail
        ControlledRebootUserCount = $script:ControlledRebootUserCount
        ControlledRebootUsers = $script:ControlledRebootUsers
        UserRebootNotificationSent = $script:UserRebootNotificationSent
        UserRebootNotificationLang = $script:UserRebootNotificationLang
        UserRebootNotificationMessage = $script:UserRebootNotificationMessage
        SetupExePath = $setupExe
        LogPath = $script:LogPath
        CsvPath = $script:CsvPath
    }

    Save-RunResult -Result $result
    Write-SmartLog ("Final Status={0}; ComputerName={1}; LocalIPv4={2}; NextAction={3}; ExitCode={4}" -f $status,$script:ComputerName,$script:LocalIPv4Addresses,$nextAction,$exitCode)
    if ($status -like 'PENDING_REBOOT*') {
        $rebootExplanation = switch ($status) {
            'PENDING_REBOOT_USER_CONNECTED' { 'Upgrade paused: a reboot is pending but was skipped because an interactive user is connected. The device will continue on a later cycle once it has rebooted or the user has logged off.' }
            'PENDING_REBOOT_USER_DETECTION_FAILED' { 'Upgrade paused: a reboot is pending but interactive user detection failed, so the reboot was skipped. Verify user sessions before rebooting.' }
            'PENDING_REBOOT_SCHEDULE_FAILED' { 'Upgrade paused: a reboot is pending and no user is connected, but scheduling the reboot failed. Check the reboot schedule.' }
            default { 'Upgrade paused: a reboot is pending and must be cleared before the Windows 11 upgrade can continue.' }
        }
        Write-SmartLog ("Pending reboot summary: {0} Source(s)={1}" -f $rebootExplanation,$finalPendingRebootSource)
    }
    if (-not [string]::IsNullOrWhiteSpace($detail)) {
        Write-SmartLog ("Final Detail={0}" -f $detail)
    }
    if ($status -eq 'WINDOWS11_COMPAT_BLOCKER' -and -not [string]::IsNullOrWhiteSpace($finalW11BlockingReasons) -and $finalW11BlockingReasons -ne $detail) {
        Write-SmartLog ("Final Windows11CompatibilityReasons={0}" -f $finalW11BlockingReasons)
    }
    Write-Output ("Status={0}; NextAction={1}; ExitCode={2}; Log={3}" -f $status,$nextAction,$exitCode,$script:LogPath)
}

exit $exitCode
