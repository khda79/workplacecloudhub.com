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

    [string]$SetupSourcePath,
    [ValidateSet('LocalCache','Share','Auto')]
    [string]$SetupExecutionMode = 'LocalCache',
    [string]$SetupMediaId = 'Win11',
    [string]$SetupLanguage = 'MatchSystem',
    [string]$SetupCacheRoot = 'C:\ProgramData\SmartM365\IntuneWindows11UpgradeToolkit\SetupMedia',
    [ValidateRange(10, 200)][int]$MinimumFreeDiskGB = 32,
    [ValidateRange(0, 86400)][int]$RebootDelaySeconds = 180,

    [string]$DataRoot = 'C:\ProgramData\SmartM365\IntuneWindows11UpgradeToolkit'
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

function Resolve-SetupUpgradeExecutable {
    if (-not $AllowSetupUpgrade) { return '' }

    $expectedLanguage = Resolve-SetupLanguageRequirement -RequestedLanguage $SetupLanguage
    $script:ResolvedSetupLanguage = $expectedLanguage

    $cachePath = Join-Path $SetupCacheRoot $SetupMediaId
    if ($SetupExecutionMode -in @('LocalCache','Auto')) {
        try {
            return Test-SetupMedia -MediaPath $cachePath -ExpectedLanguage $expectedLanguage
        }
        catch {
            if ($SetupExecutionMode -eq 'LocalCache' -and [string]::IsNullOrWhiteSpace($SetupSourcePath)) {
                throw ("Local setup cache is not ready and no SetupSourcePath was provided. CachePath={0}; Error={1}" -f $cachePath,$_.Exception.Message)
            }
            Write-SmartLog ("Local setup cache not ready: {0}" -f $_.Exception.Message) 'WARN'
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
        New-SmartDirectory -Path $cachePath
        Write-SmartLog ("Copying setup media from '{0}' to local cache '{1}'." -f $expandedSource,$cachePath)
        Copy-Item -Path (Join-Path $expandedSource '*') -Destination $cachePath -Recurse -Force -ErrorAction Stop
        return Test-SetupMedia -MediaPath $cachePath -ExpectedLanguage $expectedLanguage
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
$computerSystem = $null

try {
    if (-not $IgnoreRunGuard -and $RunGuardHours -gt 0 -and (Test-Path -LiteralPath $script:LastRunPath -PathType Leaf)) {
        $last = Get-Content -LiteralPath $script:LastRunPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if ($last.EndTimeUtc) {
            $age = (Get-Date).ToUniversalTime() - ([datetime]$last.EndTimeUtc)
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
        $status = 'INSUFFICIENT_DISK'
        $nextAction = 'FREE_DISK_SPACE'
        $detail = ("FreeDiskGB={0}; RequiredGB={1}" -f $freeGb,$MinimumFreeDiskGB)
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
    elseif ($AllowWUReset -and -not $AuditOnly) {
        Reset-WindowsUpdateComponents
        $status = 'WU_RESET_COMPLETED'
        $nextAction = 'RETRY_WINDOWS_UPDATE'
        $actionResult = 'WUResetCompleted'
        $exitCode = 0
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
        SetupExePath = $setupExe
        LogPath = $script:LogPath
        CsvPath = $script:CsvPath
    }

    Save-RunResult -Result $result
    Write-SmartLog ("Final Status={0}; NextAction={1}; ExitCode={2}" -f $status,$nextAction,$exitCode)
    Write-Output ("Status={0}; NextAction={1}; ExitCode={2}; Log={3}" -f $status,$nextAction,$exitCode,$script:LogPath)
}

exit $exitCode
