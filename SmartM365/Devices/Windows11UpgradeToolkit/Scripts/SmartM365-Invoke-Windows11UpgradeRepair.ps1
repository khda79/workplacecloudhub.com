<#
.SYNOPSIS
    Diagnoses and optionally repairs Windows 10 devices that should move to Windows 11.

.DESCRIPTION
    Autonomous endpoint-side script for the Smart Intune Windows 11 Upgrade Toolkit.
    It is designed to run locally or as SYSTEM through PsExec/LOT orchestration.

    By default the script is diagnostic-only. Corrective actions require explicit switches.
    Setup-based upgrade requires -AllowSetupUpgrade and a validated setup source/cache.

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Version: 0.1.0
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
    [switch]$AllowReboot,
    [switch]$SkipVirtualMachines,
    [switch]$AllowDiskCleanup,
    [switch]$AllowAdvancedDiskCleanup,
    [switch]$AllowDismComponentCleanup,

    [string]$SetupSourcePath,
    [ValidateSet('LocalCache','Share','Auto')]
    [string]$SetupExecutionMode = 'LocalCache',
    [string]$SetupMediaId = 'Win11',
    [string]$SetupLanguage = 'MatchSystem',
    [switch]$SkipSetupMediaPreCopy,
    [string]$SetupCacheRoot = 'C:\ProgramData\SmartM365\Windows11UpgradeToolkit\SetupMedia',
    [ValidateRange(10, 200)][int]$MinimumFreeDiskGB = 32,
    [ValidateRange(0, 365)][int]$DiskCleanupTempFileMinAgeDays = 1,
    [ValidateRange(0, 365)][int]$DiskCleanupLogRetentionDays = 14,
    [ValidateRange(0, 365)][int]$DiskCleanupUpgradeFolderMinAgeDays = 14,
    [ValidateRange(0, 86400)][int]$RebootDelaySeconds = 180,

    [string]$DataRoot = 'C:\ProgramData\SmartM365\Windows11UpgradeToolkit'
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$script:ScriptName = 'SmartM365-Invoke-Windows11UpgradeRepair'
$script:ScriptVersion = '0.1.0'
$script:RunId = Get-Date -Format 'yyyyMMdd-HHmmss'
$script:ComputerName = $env:COMPUTERNAME
$script:LogDir = Join-Path $DataRoot 'Logs'
$script:OutputDir = Join-Path $DataRoot 'Output'
$script:SetupLogDir = Join-Path $script:LogDir 'SetupUpgrade'
$script:LogPath = Join-Path $script:LogDir ("{0}_{1}_{2}.log" -f $script:ScriptName,$script:ComputerName,$script:RunId)
$script:CsvPath = Join-Path $script:OutputDir ("SmartM365_Windows11Upgrade_{0}_{1}.csv" -f $script:ComputerName,$script:RunId)
$script:LastRunPath = Join-Path $DataRoot 'LastRun.json'

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

    $isVirtual = (-not [string]::IsNullOrWhiteSpace($matchedPattern)) -or $hypervisorPresent
    $evidence = if ($matchedPattern) {
        "Manufacturer=$manufacturer; Model=$model; Pattern=$matchedPattern"
    }
    elseif ($hypervisorPresent) {
        "Manufacturer=$manufacturer; Model=$model; HypervisorPresent=True"
    }
    else {
        "Manufacturer=$manufacturer; Model=$model"
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
    $paths = @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired',
        'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager'
    )

    if (Test-Path -LiteralPath $paths[0]) { return $true }
    if (Test-Path -LiteralPath $paths[1]) { return $true }
    $pendingFileRename = Get-RegistryValue -Path $paths[2] -Name 'PendingFileRenameOperations'
    return ($null -ne $pendingFileRename)
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
        $upEx = [string]$props.UpEx
        $gated = [string]$props.GatedBlockId
        $red = [string]$props.RedReason
        $sysReq = [string]$props.SysReqIssue
        $hasBlock = ($upEx -match '(Red|Blocked|Hold)' -or -not [string]::IsNullOrWhiteSpace($gated) -or -not [string]::IsNullOrWhiteSpace($red) -or -not [string]::IsNullOrWhiteSpace($sysReq))
        if (-not $hasBlock) { continue }

        $reason = ("Target={0}; UpEx={1}; GatedBlockId={2}; RedReason={3}; SysReqIssue={4}" -f $target,$upEx,$gated,$red,$sysReq)
        $tokens = @($red,$sysReq -split '[,; ]+') | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ }
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

function Test-SetupMedia {
    param(
        [Parameter(Mandatory = $true)][string]$MediaPath,
        [string]$ExpectedLanguage
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

    $mediaRoot = Split-Path -Parent $setupExe
    $installWim = Join-Path $mediaRoot 'sources\install.wim'
    $installEsd = Join-Path $mediaRoot 'sources\install.esd'
    if (-not (Test-Path -LiteralPath $installWim -PathType Leaf) -and -not (Test-Path -LiteralPath $installEsd -PathType Leaf)) {
        throw "Windows setup media is incomplete. Missing sources\install.wim or sources\install.esd under: $mediaRoot"
    }

    $setupItem = Get-Item -LiteralPath $setupExe -ErrorAction Stop
    if ($setupItem.Length -lt 64KB) {
        throw "setup.exe exists but size is unexpectedly small: $setupExe"
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
        $langHash = (Get-FileHash -LiteralPath $langIni -Algorithm SHA256 -ErrorAction Stop).Hash
    }

    [pscustomobject]@{
        MediaId = $SetupMediaId
        ExpectedLanguage = [string]$ExpectedLanguage
        MediaLanguages = (@(Get-SetupMediaLanguages -MediaRoot $MediaRoot) -join ',')
        SetupExeLength = [int64]$setupItem.Length
        SetupExeHash = (Get-FileHash -LiteralPath $setupItem.FullName -Algorithm SHA256 -ErrorAction Stop).Hash
        InstallImageName = $installItem.Name
        InstallImageLength = [int64]$installItem.Length
        InstallImageLastWriteUtc = $installItem.LastWriteTimeUtc.ToString('o')
        LangIniHash = $langHash
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
        foreach ($property in @('MediaId','ExpectedLanguage','SetupExeLength','SetupExeHash','InstallImageName','InstallImageLength','LangIniHash')) {
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

    foreach ($property in @('MediaId','ExpectedLanguage','SetupExeLength','SetupExeHash','InstallImageName','InstallImageLength','LangIniHash')) {
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

function Test-SetupCacheReady {
    param(
        [Parameter(Mandatory = $true)][string]$CachePath,
        [string]$ExpectedLanguage
    )

    $setupExe = Test-SetupMedia -MediaPath $CachePath -ExpectedLanguage $ExpectedLanguage
    $fingerprint = Get-SetupMediaFingerprint -MediaRoot (Split-Path -Parent $setupExe) -ExpectedLanguage $ExpectedLanguage
    if (-not (Test-SetupCacheManifest -CachePath $CachePath -Fingerprint $fingerprint)) {
        Save-SetupCacheManifest -CachePath $CachePath -Fingerprint $fingerprint -SourcePath 'ExistingCache'
        Write-SmartLog ("Setup cache was valid but manifest was missing or stale. Manifest refreshed: {0}" -f $CachePath)
    }
    return $setupExe
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

    Write-SmartLog ("Copying setup media on target from '{0}' to local cache '{1}'." -f $SourcePath,$CachePath)
    $robocopy = Join-Path $env:SystemRoot 'System32\robocopy.exe'
    if (Test-Path -LiteralPath $robocopy -PathType Leaf) {
        & $robocopy $SourcePath $CachePath /MIR /Z /R:2 /W:5 /NP /NFL /NDL "/LOG+:$robocopyLog" | Out-Null
        $copyExit = [int]$LASTEXITCODE
        if ($copyExit -gt 7) {
            throw "Robocopy setup media copy failed with exit code $copyExit. Log=$robocopyLog"
        }
        Write-SmartLog ("Robocopy setup media copy completed with exit code {0}. Log={1}" -f $copyExit,$robocopyLog)
    }
    else {
        Copy-Item -Path (Join-Path $SourcePath '*') -Destination $CachePath -Recurse -Force -ErrorAction Stop
        Write-SmartLog 'Setup media copied with Copy-Item because robocopy.exe was not found.'
    }

    $setupExe = Test-SetupMedia -MediaPath $CachePath -ExpectedLanguage $ExpectedLanguage
    $fingerprint = Get-SetupMediaFingerprint -MediaRoot (Split-Path -Parent $setupExe) -ExpectedLanguage $ExpectedLanguage
    Save-SetupCacheManifest -CachePath $CachePath -Fingerprint $fingerprint -SourcePath $SourcePath
    $script:SetupCacheAction = 'CopiedByTarget'
    return $setupExe
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
    if (-not $AllowSetupUpgrade) { return '' }

    $expectedLanguage = Resolve-SetupLanguageRequirement -RequestedLanguage $SetupLanguage
    $script:ResolvedSetupLanguage = $expectedLanguage

    $cachePath = Resolve-SetupCachePath -ExpectedLanguage $expectedLanguage
    $script:ResolvedSetupCachePath = $cachePath
    if ($SetupExecutionMode -in @('LocalCache','Auto')) {
        try {
            $script:SetupCacheAction = 'AlreadyCached'
            $cachedSetupExe = Test-SetupCacheReady -CachePath $cachePath -ExpectedLanguage $expectedLanguage
            if (-not $SkipSetupMediaPreCopy -and -not [string]::IsNullOrWhiteSpace($SetupSourcePath)) {
                try {
                    $candidateSource = [Environment]::ExpandEnvironmentVariables($SetupSourcePath.Trim('"'))
                    if (Test-Path -LiteralPath $candidateSource -PathType Container) {
                        $candidateSource = Resolve-SetupSourceMediaPath -SourcePath $candidateSource -ExpectedLanguage $expectedLanguage
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
            if ($SetupExecutionMode -eq 'LocalCache' -and [string]::IsNullOrWhiteSpace($SetupSourcePath)) {
                throw ("Local setup cache is not ready and no SetupSourcePath was provided. CachePath={0}; Error={1}" -f $cachePath,$_.Exception.Message)
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($SetupSourcePath)) {
        throw 'SetupSourcePath is required when -AllowSetupUpgrade is used and no valid local cache exists.'
    }

    $expandedSource = [Environment]::ExpandEnvironmentVariables($SetupSourcePath.Trim('"'))
    if (Test-Path -LiteralPath $expandedSource -PathType Container) {
        $expandedSource = Resolve-SetupSourceMediaPath -SourcePath $expandedSource -ExpectedLanguage $expectedLanguage
    }
    if ($SetupExecutionMode -in @('Share','Auto')) {
        try {
            return Test-SetupMedia -MediaPath $expandedSource -ExpectedLanguage $expectedLanguage
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

function Invoke-SetupUpgrade {
    param([Parameter(Mandatory = $true)][string]$SetupExePath)

    New-SmartDirectory -Path $script:SetupLogDir
    $args = @(
        '/auto','upgrade',
        '/quiet',
        '/dynamicupdate','enable',
        '/copylogs',"`"$script:SetupLogDir`""
    )

    Write-SmartLog ("Starting setup upgrade: {0} {1}" -f $SetupExePath,($args -join ' '))
    $process = Start-Process -FilePath $SetupExePath -ArgumentList $args -Wait -PassThru -ErrorAction Stop
    Write-SmartLog ("setup.exe exited with code {0}" -f $process.ExitCode)
    return $process.ExitCode
}

function Save-RunResult {
    param([Parameter(Mandatory = $true)]$Result)

    New-SmartDirectory -Path $script:OutputDir
    $Result | Export-Csv -LiteralPath $script:CsvPath -NoTypeInformation -Encoding UTF8
    $Result | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $script:LastRunPath -Encoding UTF8
}

New-SmartDirectory -Path $script:LogDir
New-SmartDirectory -Path $script:OutputDir
Write-SmartLog ("===== {0} v{1} started. RunId={2} =====" -f $script:ScriptName,$script:ScriptVersion,$script:RunId)

$exitCode = 3
$status = 'UNKNOWN'
$nextAction = 'CHECK_LOGS'
$detail = ''
$actionResult = ''
$setupExe = ''
$script:ResolvedSetupLanguage = ''
$script:ResolvedSetupMediaLanguages = ''
$script:ResolvedSetupCachePath = ''
$script:SetupCacheAction = ''
$script:DiskCleanupAction = ''
$script:DiskCleanupFreedGB = ''
$script:AdvancedDiskCleanupAction = ''
$script:AdvancedDiskCleanupFreedGB = ''
$script:DismCleanupAction = ''
$script:DismCleanupFreedGB = ''
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
    $freeGb = Get-SystemDriveFreeGb
    $pendingReboot = Test-PendingReboot
    $intune = Get-IntuneEnrollmentSummary
    $policy = Get-WindowsUpdatePolicySummary
    $indicators = Get-Windows11IndicatorSummary

    $diskCleanupEligible = ($os.MajorFamily -eq 'Windows10' -and $intune.IsIntuneEnrolled -and -not $indicators.ActionableBlocking)
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
        if ($AllowReboot -and -not $AuditOnly) {
            shutdown.exe /r /t $RebootDelaySeconds /c 'SmartM365 Windows 11 upgrade readiness reboot' | Out-Null
            $actionResult = 'RebootScheduled'
            $exitCode = 0
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
            $actionResult = "SetupExitCode=$setupExitCode"
            if ($setupExitCode -eq 0) {
                $status = 'SETUP_UPGRADE_STARTED'
                $nextAction = 'MONITOR_SETUP_AND_REBOOT'
                $exitCode = 0
            }
            elseif ($setupExitCode -eq 3010) {
                $status = 'SETUP_UPGRADE_REBOOT_REQUIRED'
                $nextAction = 'REBOOT_DEVICE'
                $exitCode = 0
            }
            else {
                $status = 'SETUP_UPGRADE_FAILED'
                $nextAction = 'CHECK_SETUP_LOGS'
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
    $finalIntuneEnrolled = ''
    $finalWuBlockers = ''
    $finalW11BlockingReasons = ''
    try { $finalFreeDisk = Get-SystemDriveFreeGb } catch { }
    try { $finalPendingReboot = Test-PendingReboot } catch { }
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
        IsVirtualMachine = if ($computerSystem) { $computerSystem.IsVirtualMachine } else { '' }
        VirtualMachineEvidence = if ($computerSystem) { $computerSystem.Evidence } else { '' }
        IntuneEnrolled = $finalIntuneEnrolled
        WUBlockers = $finalWuBlockers
        W11BlockingReasons = $finalW11BlockingReasons
        SetupExecutionMode = $SetupExecutionMode
        SetupMediaId = $SetupMediaId
        SetupLanguage = $SetupLanguage
        ResolvedSetupLanguage = $script:ResolvedSetupLanguage
        SetupMediaLanguages = $script:ResolvedSetupMediaLanguages
        SetupCachePath = $script:ResolvedSetupCachePath
        SetupCacheAction = $script:SetupCacheAction
        DiskCleanupAction = $script:DiskCleanupAction
        DiskCleanupFreedGB = $script:DiskCleanupFreedGB
        AdvancedDiskCleanupAction = $script:AdvancedDiskCleanupAction
        AdvancedDiskCleanupFreedGB = $script:AdvancedDiskCleanupFreedGB
        DismCleanupAction = $script:DismCleanupAction
        DismCleanupFreedGB = $script:DismCleanupFreedGB
        SetupExePath = $setupExe
        LogPath = $script:LogPath
        CsvPath = $script:CsvPath
    }

    Save-RunResult -Result $result
    Write-SmartLog ("Final Status={0}; NextAction={1}; ExitCode={2}" -f $status,$nextAction,$exitCode)
    Write-Output ("Status={0}; NextAction={1}; ExitCode={2}; Log={3}" -f $status,$nextAction,$exitCode,$script:LogPath)
}

exit $exitCode
