<#
.SYNOPSIS
    Runs SmartM365-Invoke-Windows11UpgradeRepair.ps1 on remote computers using PsExec.

.DESCRIPTION
    LOT/PsExec orchestrator for Windows 10 to Windows 11 upgrade diagnostics and guarded repair.
    It copies the autonomous endpoint script to each target, lets the target validate/cache
    Windows 11 setup media when setup upgrade is enabled, starts the script as SYSTEM,
    collects evidence, and writes cycle CSV reports.

.VERSION
    0.1.67

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding(PositionalBinding = $false)]
param(
    [string]$ComputerListPath = (Join-Path (Get-Location) 'Computers.txt'),
    [string]$PsExecPath,
    [string]$LocalScriptPath,

    [switch]$AuditOnly,
    [switch]$DryRun,
    [switch]$RunOnce,
    [switch]$IgnoreRunGuard,
    [switch]$UseTechnicianRunGuardHistory,
    [switch]$IgnoreTechnicianRunGuardHistory,
    [ValidateRange(0, 168)][int]$RunGuardHours = 3,
    [switch]$AllowPolicyRepair,
    [switch]$AllowWUReset,
    [switch]$AllowForceUpgrade,
    [switch]$AllowSetupUpgrade,
    [switch]$DirectSetupUpgrade,
    [switch]$AllowReboot,
    [switch]$AllowSetupCompletionRebootWhenNoUser,
    [switch]$AllowSetupProfileRepair,
    [switch]$ScheduleRetryAfterReboot,
    [ValidateRange(1, 30)][int]$RetryAfterRebootMaxAttempts = 3,
    [ValidateRange(0, 3600)][int]$RetryAfterRebootDelaySeconds = 300,
    [ValidateRange(0, 3650)][int]$ForceRequiredRebootWhenUptimeOverDays = 7,
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
    [ValidateRange(0, 100)][int]$SetupSourceCandidateLimit = 5,
    [ValidateRange(0, 10000)][int]$SetupMediaCopyIpGapMilliseconds = 0,
    [ValidateRange(0, 86400)][int]$SetupMediaCopyJitterSeconds = 0,
    [ValidateRange(0, 500)][int]$SetupSourceConcurrencyLimit = 0,
    [ValidateRange(1, 1440)][int]$SetupSourceConcurrencyLeaseMinutes = 240,
    [string]$SetupSourceConcurrencyGateRoot,
    [ValidateRange(0, 500)][int]$SetupSubnetConcurrencyLimit = 0,
    [string]$SetupSubnetPrefixLength = 'Auto',
    [ValidateRange(1, 1440)][int]$SetupSubnetConcurrencyLeaseMinutes = 90,
    [string]$SetupSubnetConcurrencyGateRoot,

    [string]$AdInventoryCsv,
    [string]$AdRootInventoryCsv,
    [string]$AdInventoryNameColumn,
    [string]$AdDomain,
    [switch]$SkipAdInventoryRefresh,
    [string]$IntuneInventoryCsv,
    [string]$IntuneRootInventoryCsv,
    [string]$IntuneInventoryNameColumn,
    [ValidateRange(1, 999)][int]$IntuneInventoryPageSize = 999,
    [string]$IntuneTenantId,
    [switch]$SkipIntuneInventoryRefresh,

    [string]$LogRoot,
    [string]$ReportRoot,
    [string]$CentralLogRoot,
    [string]$LauncherLogRoot,
    [switch]$NoCentralLogCollection,
    [switch]$KeepCentralLogHistory,
    [ValidateSet('Standard','Full')]
    [string]$CentralLogCollectionMode = 'Standard',

    [ValidateRange(1, 200)][int]$ThrottleLimit = 10,
    [ValidateRange(0, 200)][int]$GlobalConcurrencyLimit = 15,
    [string]$GlobalConcurrencySemaphoreName = 'Local\SmartM365_Windows11UpgradeToolkit_ComputerWorkers',
    [ValidateRange(0, 1440)][int]$GlobalConcurrencyLeaseTimeoutMinutes = 0,
    [ValidateRange(0, 3600)][int]$DelayBetweenComputersSeconds = 0,
    [ValidateRange(1, 60)][int]$JobPollSeconds = 2,
    [ValidateRange(0, 1440)][int]$DelayBetweenCyclesMinutes = 10,
    [ValidateRange(0, 1000)][int]$MaxCycles = 0,
    [ValidateRange(0, 1440)][int]$CancellationDrainTimeoutMinutes = 15,
    [ValidateRange(0, 1440)][int]$PsExecTimeoutMinutes = 360,

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$UnexpectedArguments
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

if ($UnexpectedArguments -and $UnexpectedArguments.Count -gt 0) {
    throw ("Unexpected launcher argument(s): {0}. Pass PsExec with -PsExecPath <path>, not as a free argument." -f ($UnexpectedArguments -join ' '))
}

$script:LauncherVersion = '0.1.67'
$script:TechnicianRunGuardStartedNoResultHours = 4
$script:BaseDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$script:ToolkitRoot = Split-Path -Parent $script:BaseDir
if ([string]::IsNullOrWhiteSpace($LocalScriptPath)) {
    $LocalScriptPath = Join-Path $script:BaseDir 'SmartM365-Invoke-Windows11UpgradeRepair.ps1'
}
$LocalWorkerPath = Join-Path $script:BaseDir 'SmartM365-Windows11Upgrade-PsExecWorker.ps1'
$script:ExportAdScriptPath = Join-Path $script:BaseDir 'SmartM365-Windows11Upgrade-Export-ADDevicesCsv.ps1'
$script:ExportIntuneScriptPath = Join-Path $script:BaseDir 'SmartM365-Windows11Upgrade-Export-IntuneDevicesCsv.ps1'
$script:AdInventoryFreshnessHours = 12
$script:IntuneInventoryFreshnessHours = 2
$ComputerListPath = [System.IO.Path]::GetFullPath($ComputerListPath)
$script:LotRoot = Split-Path -Parent $ComputerListPath
$script:LauncherRunTimestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$defaultRunRoot = $null
$lotParent = Split-Path -Parent $script:LotRoot
if ((Split-Path -Leaf $lotParent) -eq 'Lots') {
    $toolkitRootFromLot = Split-Path -Parent $lotParent
    $lotNameFromPath = Split-Path -Leaf $script:LotRoot
    $defaultRunRoot = Join-Path (Join-Path (Join-Path $toolkitRootFromLot 'Runs') $lotNameFromPath) $script:LauncherRunTimestamp
}
if ([string]::IsNullOrWhiteSpace($defaultRunRoot)) {
    $missingRunRoots = @()
    if ([string]::IsNullOrWhiteSpace($LogRoot)) { $missingRunRoots += 'LogRoot' }
    if ([string]::IsNullOrWhiteSpace($ReportRoot)) { $missingRunRoots += 'ReportRoot' }
    if ([string]::IsNullOrWhiteSpace($CentralLogRoot)) { $missingRunRoots += 'CentralLogRoot' }
    if ([string]::IsNullOrWhiteSpace($LauncherLogRoot)) { $missingRunRoots += 'LauncherLogRoot' }
    if ($missingRunRoots.Count -gt 0) {
        throw ('ComputerListPath must be under Lots\LOT-* or these run roots must be specified explicitly: {0}' -f ($missingRunRoots -join ', '))
    }
}
if ([string]::IsNullOrWhiteSpace($LogRoot)) { $LogRoot = Join-Path $defaultRunRoot 'PsExecLogs' }
if ([string]::IsNullOrWhiteSpace($ReportRoot)) { $ReportRoot = Join-Path $defaultRunRoot 'Reports' }
if ([string]::IsNullOrWhiteSpace($CentralLogRoot)) { $CentralLogRoot = Join-Path $defaultRunRoot 'CentralLogs' }
$LogRoot = [System.IO.Path]::GetFullPath($LogRoot)
$ReportRoot = [System.IO.Path]::GetFullPath($ReportRoot)
$CentralLogRoot = [System.IO.Path]::GetFullPath($CentralLogRoot)
$runDataRoot = Split-Path -Parent $LogRoot
$script:LotAdInventoryCsv = Join-Path $runDataRoot 'DevicesAD.csv'
$script:LotIntuneInventoryCsv = Join-Path $runDataRoot 'DevicesIntune.csv'
if ([string]::IsNullOrWhiteSpace($LauncherLogRoot)) { $LauncherLogRoot = Join-Path (Split-Path -Parent $LogRoot) 'Logs' }
$script:LauncherLogRoot = [System.IO.Path]::GetFullPath($LauncherLogRoot)
$script:LauncherLotParentName = Split-Path -Leaf (Split-Path -Parent $script:LotRoot)
$script:IsSingleComputerPathLaunch = ($script:LauncherLotParentName -in @('SingleComputer','SingleComputerRuns'))
$script:LauncherLotName = Split-Path -Leaf $script:LotRoot
if ([string]::IsNullOrWhiteSpace($script:LauncherLotName)) { $script:LauncherLotName = 'UnknownLOT' }
$script:LauncherLogSafeLotName = [regex]::Replace($script:LauncherLotName, '[^A-Za-z0-9._-]', '_')
if ($script:IsSingleComputerPathLaunch) { $script:LauncherLogSafeLotName = 'SingleComputer' }
if ([string]::IsNullOrWhiteSpace($script:LauncherLogSafeLotName)) { $script:LauncherLogSafeLotName = 'UnknownLOT' }
$script:LauncherLogPath = Join-Path $script:LauncherLogRoot ("SmartM365-W11UT-Launcher_{0}_{1}.log" -f $script:LauncherLogSafeLotName,$script:LauncherRunTimestamp)
$script:LauncherLatestLogPath = Join-Path $script:LauncherLogRoot ("SmartM365-W11UT-Launcher_{0}_latest.log" -f $script:LauncherLogSafeLotName)

$AdInventoryUsesRecentRootCsv = $false
$requestedAdInventoryCsv = $AdInventoryCsv
$defaultRootAdCsv = Join-Path $script:ToolkitRoot 'DevicesAD.csv'
$defaultLotAdFullName = [System.IO.Path]::GetFullPath($script:LotAdInventoryCsv)
$requestedAdFullName = if (-not [string]::IsNullOrWhiteSpace($requestedAdInventoryCsv)) { [System.IO.Path]::GetFullPath($requestedAdInventoryCsv) } else { '' }
if ([string]::IsNullOrWhiteSpace($AdInventoryCsv)) {
    $AdInventoryCsv = $script:LotAdInventoryCsv
}
if ([string]::IsNullOrWhiteSpace($AdRootInventoryCsv) -and (Test-Path -LiteralPath $defaultRootAdCsv -PathType Leaf)) {
    $AdRootInventoryCsv = $defaultRootAdCsv
}
if (-not [string]::IsNullOrWhiteSpace($AdRootInventoryCsv)) {
    $adRootInventoryItem = Get-Item -LiteralPath $AdRootInventoryCsv -ErrorAction SilentlyContinue
    if ($adRootInventoryItem) {
        $adRootFullName = [System.IO.Path]::GetFullPath($adRootInventoryItem.FullName)
        $adRootInventoryAge = (Get-Date) - $adRootInventoryItem.LastWriteTime
        if ($adRootInventoryAge.TotalHours -le $script:AdInventoryFreshnessHours) {
            $AdRootInventoryCsv = $adRootFullName
            $requestedAdItem = if (-not [string]::IsNullOrWhiteSpace($requestedAdFullName)) { Get-Item -LiteralPath $requestedAdFullName -ErrorAction SilentlyContinue } else { $null }
            $shouldUseRootAdCsv = [string]::IsNullOrWhiteSpace($requestedAdFullName) -or (($requestedAdFullName -eq $defaultLotAdFullName) -and (-not $requestedAdItem))
            if ($shouldUseRootAdCsv) {
                $AdInventoryCsv = $adRootFullName
                $AdInventoryUsesRecentRootCsv = $true
            }
        }
    }
}
if (-not [string]::IsNullOrWhiteSpace($AdInventoryCsv)) { $AdInventoryCsv = [System.IO.Path]::GetFullPath($AdInventoryCsv) }
if (-not [string]::IsNullOrWhiteSpace($AdRootInventoryCsv)) { $AdRootInventoryCsv = [System.IO.Path]::GetFullPath($AdRootInventoryCsv) }
$EffectiveSkipAdInventoryRefresh = [bool]$SkipAdInventoryRefresh -or $AdInventoryUsesRecentRootCsv

$IntuneInventoryUsesRecentRootCsv = $false
$requestedIntuneInventoryCsv = $IntuneInventoryCsv
$defaultRootIntuneCsv = Join-Path $script:ToolkitRoot 'DevicesIntune.csv'
$defaultLotIntuneFullName = [System.IO.Path]::GetFullPath($script:LotIntuneInventoryCsv)
$requestedIntuneFullName = if (-not [string]::IsNullOrWhiteSpace($requestedIntuneInventoryCsv)) { [System.IO.Path]::GetFullPath($requestedIntuneInventoryCsv) } else { '' }
if ([string]::IsNullOrWhiteSpace($IntuneInventoryCsv)) {
    $IntuneInventoryCsv = $script:LotIntuneInventoryCsv
}
if ([string]::IsNullOrWhiteSpace($IntuneRootInventoryCsv) -and (Test-Path -LiteralPath $defaultRootIntuneCsv -PathType Leaf)) {
    $IntuneRootInventoryCsv = $defaultRootIntuneCsv
}
if (-not [string]::IsNullOrWhiteSpace($IntuneRootInventoryCsv)) {
    $intuneRootInventoryItem = Get-Item -LiteralPath $IntuneRootInventoryCsv -ErrorAction SilentlyContinue
    if ($intuneRootInventoryItem) {
        $intuneRootFullName = [System.IO.Path]::GetFullPath($intuneRootInventoryItem.FullName)
        $intuneRootInventoryAge = (Get-Date) - $intuneRootInventoryItem.LastWriteTime
        if ($intuneRootInventoryAge.TotalHours -le $script:IntuneInventoryFreshnessHours) {
            $IntuneRootInventoryCsv = $intuneRootFullName
            $requestedIntuneItem = if (-not [string]::IsNullOrWhiteSpace($requestedIntuneFullName)) { Get-Item -LiteralPath $requestedIntuneFullName -ErrorAction SilentlyContinue } else { $null }
            $shouldUseRootIntuneCsv = [string]::IsNullOrWhiteSpace($requestedIntuneFullName) -or (($requestedIntuneFullName -eq $defaultLotIntuneFullName) -and (-not $requestedIntuneItem))
            if ($shouldUseRootIntuneCsv) {
                $IntuneInventoryCsv = $intuneRootFullName
                $IntuneInventoryUsesRecentRootCsv = $true
            }
        }
    }
}
if (-not [string]::IsNullOrWhiteSpace($IntuneInventoryCsv)) { $IntuneInventoryCsv = [System.IO.Path]::GetFullPath($IntuneInventoryCsv) }
if (-not [string]::IsNullOrWhiteSpace($IntuneRootInventoryCsv)) { $IntuneRootInventoryCsv = [System.IO.Path]::GetFullPath($IntuneRootInventoryCsv) }
$EffectiveSkipIntuneInventoryRefresh = [bool]$SkipIntuneInventoryRefresh -or $IntuneInventoryUsesRecentRootCsv
$script:RemoteBaseDir = 'C:\ProgramData\SmartM365\Windows11UpgradeToolkit'
$script:RemoteScriptPath = Join-Path $script:RemoteBaseDir 'SmartM365-Invoke-Windows11UpgradeRepair.ps1'
$script:RemoteSetupCacheRoot = Join-Path $script:RemoteBaseDir 'SetupMedia'
if ($GlobalConcurrencyLeaseTimeoutMinutes -lt 1) {
    $timeoutBase = if ($PsExecTimeoutMinutes -gt 0) { $PsExecTimeoutMinutes } else { 240 }
    $GlobalConcurrencyLeaseTimeoutMinutes = [Math]::Max(30, $timeoutBase + 30)
}

function Get-TechnicianIdentityInfo {
    $identity = $null
    try { $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent() } catch { }

    $upn = ''
    try {
        $whoamiUpn = & whoami.exe /upn 2>$null
        if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($whoamiUpn)) { $upn = ([string]$whoamiUpn).Trim() }
    }
    catch { }

    $account = if ($identity -and -not [string]::IsNullOrWhiteSpace($identity.Name)) { [string]$identity.Name } else { [Environment]::UserName }
    $sid = if ($identity -and $identity.User) { [string]$identity.User.Value } else { '' }
    $authType = if ($identity) { [string]$identity.AuthenticationType } else { '' }

    return [pscustomobject]@{
        Account = $account
        UserPrincipalName = $upn
        Sid = $sid
        AuthenticationType = $authType
        UserDomain = [Environment]::UserDomainName
        UserName = [Environment]::UserName
        UserDnsDomain = [Environment]::GetEnvironmentVariable('USERDNSDOMAIN')
        ComputerName = $env:COMPUTERNAME
    }
}

$script:TechnicianIdentity = Get-TechnicianIdentityInfo

function New-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
    }
}

function Initialize-LotCancellationSupport {
    if (-not ('SmartM365LotCancellation' -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Threading;

public static class SmartM365LotCancellation
{
    private static ConsoleCancelEventHandler handler;
    private static int requestCount;

    public static int RequestCount { get { return requestCount; } }

    public static void Register()
    {
        Unregister();
        requestCount = 0;
        handler = new ConsoleCancelEventHandler(OnCancelKeyPress);
        Console.CancelKeyPress += handler;
    }

    public static void Unregister()
    {
        if (handler != null)
        {
            Console.CancelKeyPress -= handler;
            handler = null;
        }
    }

    private static void OnCancelKeyPress(object sender, ConsoleCancelEventArgs e)
    {
        e.Cancel = true;
        Interlocked.Increment(ref requestCount);
    }
}
"@
    }

    [SmartM365LotCancellation]::Register()
}

function Get-LotCancellationState {
    $requestCount = 0
    try { $requestCount = [SmartM365LotCancellation]::RequestCount } catch { }
    $requested = ($requestCount -gt 0)
    $force = ($requestCount -gt 1)
    $source = if ($requested) { 'Ctrl+C' } else { '' }

    if (-not [string]::IsNullOrWhiteSpace($script:CancellationSignalPath) -and (Test-Path -LiteralPath $script:CancellationSignalPath -PathType Leaf)) {
        $requested = $true
        $source = 'GUI_OR_FILE_SIGNAL'
        try {
            $signal = Get-Content -LiteralPath $script:CancellationSignalPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($signal.PSObject.Properties['Force'] -and [bool]$signal.Force) { $force = $true }
            if ($signal.PSObject.Properties['Source'] -and -not [string]::IsNullOrWhiteSpace([string]$signal.Source)) { $source = [string]$signal.Source }
        }
        catch { }
    }

    return [pscustomobject]@{ Requested = $requested; Force = $force; Source = $source; RequestCount = $requestCount }
}

function Wait-LotCancellationAware {
    param([ValidateRange(0, 86400)][int]$Seconds)

    $remaining = $Seconds
    while ($remaining -gt 0) {
        if ((Get-LotCancellationState).Requested) { return $false }
        $slice = [Math]::Min(2, $remaining)
        Start-Sleep -Seconds $slice
        $remaining -= $slice
    }
    return $true
}

function Set-ActiveLotRunState {
    param([Parameter(Mandatory = $true)][string]$Status,[string]$ReportPath = '')

    if ([string]::IsNullOrWhiteSpace($script:ActiveLotRunStatePath)) { return }
    [pscustomobject]@{
        Version = 1
        Toolkit = 'Windows11UpgradeToolkit'
        Status = $Status
        ProcessId = $PID
        LotName = $script:LauncherLotName
        ComputerListPath = $ComputerListPath
        SignalPath = $script:CancellationSignalPath
        ReportPath = $ReportPath
        StartedUtc = $script:CancellationRunStartedUtc.ToString('o')
        UpdatedUtc = (Get-Date).ToUniversalTime().ToString('o')
    } | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $script:ActiveLotRunStatePath -Encoding UTF8 -Force
}

function Complete-LotCancellationSupport {
    try { [SmartM365LotCancellation]::Unregister() } catch { }
    if (-not [string]::IsNullOrWhiteSpace($script:ActiveLotRunStatePath)) {
        Remove-Item -LiteralPath $script:ActiveLotRunStatePath -Force -ErrorAction SilentlyContinue
    }
    if (-not [string]::IsNullOrWhiteSpace($script:CancellationSignalPath)) {
        Remove-Item -LiteralPath $script:CancellationSignalPath -Force -ErrorAction SilentlyContinue
    }
}

function New-Windows11CancellationResult {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][int]$CycleNumber,
        [Parameter(Mandatory = $true)][string]$Status,
        [Parameter(Mandatory = $true)][string]$Detail
    )

    return [pscustomobject]@{
        Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        ComputerName = $ComputerName
        CycleNumber = $CycleNumber
        LauncherStatus = $Status
        RemoteStatus = ''
        RemoteNextAction = 'VERIFY_REMOTE_STATE_BEFORE_RELAUNCH'
        ExitCode = ''
        Detail = $Detail
        JobErrorMessage = ''
        RemoteLogsPath = ''
        PsExecLogPath = $script:LauncherLogPath
    }
}

function Get-TechnicianRunGuardHistoryPath {
    $stateRoot = Join-Path ([Environment]::GetFolderPath('CommonApplicationData')) 'SmartM365\Windows11UpgradeToolkit\LauncherState'
    return (Join-Path $stateRoot 'RunGuardHistory.json')
}

function Invoke-TechnicianRunGuardHistoryLock {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock,
        [object[]]$ArgumentList = @()
    )

    $lockStream = $null
    $lockPath = '{0}.lock' -f (Get-TechnicianRunGuardHistoryPath)
    $lockParent = Split-Path -Parent $lockPath
    $deadlineUtc = (Get-Date).ToUniversalTime().AddSeconds(60)
    New-Directory -Path $lockParent
    try {
        while ($null -eq $lockStream -and (Get-Date).ToUniversalTime() -lt $deadlineUtc) {
            try {
                $lockStream = [System.IO.File]::Open(
                    $lockPath,
                    [System.IO.FileMode]::OpenOrCreate,
                    [System.IO.FileAccess]::ReadWrite,
                    [System.IO.FileShare]::None
                )
            }
            catch [System.IO.IOException] { Start-Sleep -Milliseconds 150 }
            catch [System.UnauthorizedAccessException] { Start-Sleep -Milliseconds 150 }
        }
        if ($null -eq $lockStream) { throw 'Timed out waiting for technician run guard history file lock after 60 seconds.' }
        & $ScriptBlock @ArgumentList
    }
    finally {
        if ($lockStream) { $lockStream.Dispose() }
    }
}

function Get-TechnicianRunGuardFqdn {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [AllowNull()][hashtable]$AdInventoryMap
    )

    $name = ([string]$ComputerName).Trim().Trim([char]34).TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($name)) { return '' }
    if ($name.Contains('.')) { return $name.ToLowerInvariant() }

    $shortKey = Get-ComputerListKey -ComputerName $name
    if ($AdInventoryMap -and $AdInventoryMap.Count -gt 0 -and $AdInventoryMap.ContainsKey($shortKey)) {
        $adRow = $AdInventoryMap[$shortKey]
        if ($adRow.PSObject.Properties['DNSHostName'] -and -not [string]::IsNullOrWhiteSpace([string]$adRow.DNSHostName)) {
            return ([string]$adRow.DNSHostName).Trim().TrimEnd('.').ToLowerInvariant()
        }
    }

    return $name.ToLowerInvariant()
}

function ConvertTo-TechnicianRunGuardUtcDateTime {
    param(
        [AllowNull()]$Value,
        [Parameter(Mandatory = $true)][ref]$Result
    )

    $Result.Value = [datetime]::MinValue
    if ($null -eq $Value) { return $false }
    if ($Value -is [datetime]) {
        $Result.Value = ([datetime]$Value).ToUniversalTime()
        return $true
    }

    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsed)) {
        $Result.Value = $parsed.ToUniversalTime()
        return $true
    }
    return $false
}

function Read-TechnicianRunGuardHistory {    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [ValidateRange(0, 168)][int]$RunGuardHours
    )

    $entries = @()
    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        try {
            $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace($raw)) {
                $data = $raw | ConvertFrom-Json -ErrorAction Stop
                if ($data.PSObject.Properties['Entries']) { $entries = @($data.Entries) }
            }
        }
        catch {
            $entries = @()
        }
    }

    if ($RunGuardHours -gt 0) {
        $nowUtc = (Get-Date).ToUniversalTime()
        $fresh = New-Object System.Collections.ArrayList
        foreach ($entry in @($entries)) {
            $startedText = if ($entry.PSObject.Properties['LastStartedUtc']) { [string]$entry.LastStartedUtc } else { '' }
            $startedUtc = [datetime]::MinValue
            if (ConvertTo-TechnicianRunGuardUtcDateTime -Value $startedText -Result ([ref]$startedUtc)) {
                if (($nowUtc - $startedUtc.ToUniversalTime()).TotalHours -lt $RunGuardHours) { [void]$fresh.Add($entry) }
            }
        }
        $entries = @($fresh.ToArray())
    }

    return [pscustomobject]@{
        Version = 1
        UpdatedUtc = (Get-Date).ToUniversalTime().ToString('o')
        Entries = @($entries)
    }
}

function Save-TechnicianRunGuardHistory {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$History
    )

    $parent = Split-Path -Parent $Path
    New-Directory -Path $parent
    $History.UpdatedUtc = (Get-Date).ToUniversalTime().ToString('o')
    $json = $History | ConvertTo-Json -Depth 8
    $lastError = $null

    for ($attempt = 1; $attempt -le 10; $attempt++) {
        $tempPath = Join-Path $parent (".{0}.{1}.tmp" -f (Split-Path -Leaf $Path),[guid]::NewGuid().ToString('N'))
        try {
            Set-Content -LiteralPath $tempPath -Value $json -Encoding UTF8 -Force -ErrorAction Stop
            Move-Item -LiteralPath $tempPath -Destination $Path -Force -ErrorAction Stop
            return
        }
        catch [System.IO.IOException] {
            $lastError = $_
            Start-Sleep -Milliseconds ([math]::Min(2000, 150 * $attempt))
        }
        catch [System.UnauthorizedAccessException] {
            $lastError = $_
            Start-Sleep -Milliseconds ([math]::Min(2000, 150 * $attempt))
        }
        finally {
            if (Test-Path -LiteralPath $tempPath -PathType Leaf) {
                try { Remove-Item -LiteralPath $tempPath -Force -ErrorAction Stop } catch { }
            }
        }
    }

    if ($lastError) { throw $lastError }
    throw ("Failed to save technician run guard history: {0}" -f $Path)
}

function Get-TechnicianRunGuardFailureCategory {
    param(
        [AllowNull()][string]$LauncherStatus,
        [AllowNull()][string]$RemoteStatus,
        [AllowNull()][string]$Detail,
        [AllowNull()][string]$JobErrorMessage
    )

    $effectiveStatus = @($RemoteStatus,$LauncherStatus) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
    $status = ([string]$effectiveStatus).ToUpperInvariant()
    $evidence = @($Detail,$JobErrorMessage) -join ' '

    if ($status -in @('ADMIN_SHARE_UNREACHABLE','DRYRUN_ADMIN_SHARE_UNREACHABLE') -or $evidence -match 'FailureType=(DNS_FAILED|SMB_PORT_445_UNREACHABLE|PING_OK_ADMIN_SHARE_FAILED|ADMIN_SHARE_UNREACHABLE)') {
        return 'NetworkTransient'
    }
    if ($status -in @('DIRECT_SETUP_UPGRADE_FAILED','SETUP_UPGRADE_FAILED','SETUP_MIGRATION_PROFILE_FAILURE','SETUP_MIGRATION_PROFILE_REPAIR_FAILED','SETUP_MIGRATION_PLUGIN_FAILURE','SETUP_PROCESS_TIMEOUT','SETUP_PROCESS_MONITOR_INTERRUPTED','SETUP_MEDIA_COPY_FAILED','SETUP_MEDIA_COPY_TIMEOUT','SETUP_MEDIA_MANIFEST_VALIDATION_FAILED')) {
        return 'SetupFailure'
    }
    if ($status -in @('INSUFFICIENT_DISK','INSUFFICIENT_DISK_AFTER_CLEANUP','UNSUPPORTED_OS','NOT_INTUNE_ENROLLED','WINDOWS11_COMPAT_BLOCKER','WINDOWS11_HARDWARE_NOT_CAPABLE','WU_POLICY_BLOCKER','SETUP_SOURCE_LANGUAGE_UNAVAILABLE')) {
        return 'OperatorAction'
    }
    if ($status -in @('ERROR','JOB_ERROR','RUNSPACE_BROKEN','PSEXEC_EXIT_UNKNOWN','PSEXEC_COMMUNICATION_LOST','CENTRAL_LOG_COLLECTION_FAILED','REMOTE_LOG_COLLECTION_FAILED','REMOTE_RESULT_STALE','REMOTE_PAYLOAD_COPY_FAILED','SETUP_CACHE_LOCKED','SETUP_SUBNET_COPY_LEASE_TIMEOUT','SETUP_SOURCE_COPY_LEASE_TIMEOUT')) {
        return 'ExecutionTransient'
    }
    return ''
}

function Get-TechnicianRunGuardRetryDelay {
    param(
        [Parameter(Mandatory = $true)][string]$FailureCategory,
        [ValidateRange(1, 1000)][int]$ConsecutiveFailureCount,
        [ValidateRange(0, 168)][int]$RunGuardHours
    )

    $maximumDelay = [math]::Max(1, ($RunGuardHours * 60))
    switch ($FailureCategory) {
        'NetworkTransient' {
            $delays = @(5,15,30,60)
            return [math]::Min($maximumDelay, $delays[[math]::Min($ConsecutiveFailureCount - 1, $delays.Count - 1)])
        }
        'ExecutionTransient' {
            $delays = @(15,30,60,120)
            return [math]::Min($maximumDelay, $delays[[math]::Min($ConsecutiveFailureCount - 1, $delays.Count - 1)])
        }
        'SetupFailure' { return [math]::Min($maximumDelay, 360) }
        'OperatorAction' { return $maximumDelay }
        default { return 0 }
    }
}

function Test-TechnicianRunGuardEntryShouldBlock {
    param(
        [AllowNull()]$Entry,
        [ValidateRange(0, 168)][int]$RunGuardHours
    )

    if (-not $Entry) { return $false }
    $state = if ($Entry.PSObject.Properties['State']) { [string]$Entry.State } else { '' }
    if ($state -ne 'Result') {
        $startedUtc = [datetime]::MinValue
        if (-not (ConvertTo-TechnicianRunGuardUtcDateTime -Value $Entry.LastStartedUtc -Result ([ref]$startedUtc))) { return $false }
        $startedNoResultGuardHours = [math]::Min([double]$RunGuardHours, [double]$script:TechnicianRunGuardStartedNoResultHours)
        return (((Get-Date).ToUniversalTime() - $startedUtc.ToUniversalTime()).TotalHours -lt $startedNoResultGuardHours)
    }

    $launcherStatus = if ($Entry.PSObject.Properties['LauncherStatus']) { [string]$Entry.LauncherStatus } else { '' }
    $remoteStatus = if ($Entry.PSObject.Properties['RemoteStatus']) { [string]$Entry.RemoteStatus } else { '' }
    $detail = if ($Entry.PSObject.Properties['Detail']) { [string]$Entry.Detail } else { '' }
    $jobErrorMessage = if ($Entry.PSObject.Properties['JobErrorMessage']) { [string]$Entry.JobErrorMessage } else { '' }
    $effectiveStatus = @($remoteStatus, $launcherStatus) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
    $effectiveStatusUpper = if ([string]::IsNullOrWhiteSpace($effectiveStatus)) { '' } else { ([string]$effectiveStatus).ToUpperInvariant() }
    $combinedEvidence = @($detail, $jobErrorMessage) -join ' '

    if ([string]::IsNullOrWhiteSpace($effectiveStatusUpper)) { return $false }

    if ($Entry.PSObject.Properties['RetryAfterUtc'] -and -not [string]::IsNullOrWhiteSpace([string]$Entry.RetryAfterUtc)) {
        $retryAfterUtc = [datetime]::MinValue
        if (ConvertTo-TechnicianRunGuardUtcDateTime -Value $Entry.RetryAfterUtc -Result ([ref]$retryAfterUtc)) {
            if ((Get-Date).ToUniversalTime() -lt $retryAfterUtc.ToUniversalTime()) { return $true }
        }
    }

    if ($effectiveStatusUpper -in @(
        'ADMIN_SHARE_UNREACHABLE',
        'DRYRUN_ADMIN_SHARE_UNREACHABLE',
        'DRYRUN_READY',
        'READY_TO_FORCE_UPGRADE',
        'DIRECT_SETUP_UPGRADE_READY',
        'SETUP_UPGRADE_READY',
        'RUN_GUARD_ACTIVE',
        'INSUFFICIENT_DISK',
        'INSUFFICIENT_DISK_AFTER_CLEANUP',
        'UNSUPPORTED_OS',
        'NOT_INTUNE_ENROLLED',
        'WINDOWS11_COMPAT_BLOCKER',
        'WINDOWS11_HARDWARE_NOT_CAPABLE',
        'WU_POLICY_BLOCKER',
        'SETUP_SOURCE_LANGUAGE_UNAVAILABLE',
        'SETUP_CACHE_LOCKED',
        'SETUP_MEDIA_COPY_TIMEOUT',
        'SETUP_MEDIA_COPY_FAILED',
        'SETUP_MEDIA_MANIFEST_VALIDATION_FAILED',
        'SETUP_SUBNET_COPY_LEASE_TIMEOUT',
        'SETUP_SOURCE_COPY_LEASE_TIMEOUT',
        'SETUP_PROCESS_TIMEOUT',
        'SETUP_PROCESS_MONITOR_INTERRUPTED',
        'SETUP_MIGRATION_PROFILE_FAILURE',
        'SETUP_MIGRATION_PROFILE_REPAIR_FAILED',
        'SETUP_PROFILE_DUPLICATE_REPAIRED_REBOOT_REQUIRED',
        'SETUP_MIGRATION_PLUGIN_FAILURE',
        'DIRECT_SETUP_UPGRADE_FAILED',
        'SETUP_UPGRADE_FAILED',
        'ERROR',
        'JOB_ERROR',
        'RUNSPACE_BROKEN',
        'PSEXEC_EXIT_UNKNOWN',
        'PSEXEC_COMMUNICATION_LOST',
        'CENTRAL_LOG_COLLECTION_FAILED',
        'REMOTE_LOG_COLLECTION_FAILED',
        'REMOTE_PAYLOAD_COPY_FAILED'
    )) { return $false }

    if ($combinedEvidence -match 'FailureType=(DNS_FAILED|SMB_PORT_445_UNREACHABLE|PING_OK_ADMIN_SHARE_FAILED|ADMIN_SHARE_UNREACHABLE)') { return $false }
    if ($combinedEvidence -match 'Central log collection failed|Le chemin r.seau n.a pas .t. trouv.|network path was not found|Error communicating with PsExec service|Descripteur non valide') { return $false }

    if ($effectiveStatusUpper -in @(
        'DIRECT_SETUP_UPGRADE_STARTED',
        'DIRECT_SETUP_UPGRADE_REBOOT_REQUIRED',
        'DIRECT_SETUP_UPGRADE_REBOOT_SCHEDULED_NO_USER',
        'DIRECT_SETUP_UPGRADE_REBOOT_SKIPPED_USER_CONNECTED',
        'DIRECT_SETUP_UPGRADE_REBOOT_SKIPPED_USER_DETECTION_FAILED',
        'SETUP_UPGRADE_STARTED',
        'SETUP_UPGRADE_REBOOT_REQUIRED',
        'SETUP_UPGRADE_REBOOT_SCHEDULED_NO_USER',
        'SETUP_UPGRADE_REBOOT_SKIPPED_USER_CONNECTED',
        'SETUP_UPGRADE_REBOOT_SKIPPED_USER_DETECTION_FAILED',
        'PENDING_REBOOT',
        'PENDING_REBOOT_USER_CONNECTED',
        'PENDING_REBOOT_USER_DETECTION_FAILED',
        'PENDING_REBOOT_SCHEDULE_FAILED',
        'WINDOWS_UPDATE_FORCE_TRIGGERED',
        'WU_RESET_COMPLETED'
    )) { return $true }

    if ($effectiveStatusUpper -like '*_STARTED' -or $effectiveStatusUpper -like '*_REBOOT_REQUIRED' -or $effectiveStatusUpper -like '*_REBOOT_SCHEDULED_NO_USER') { return $true }
    return $false
}
function Get-ActiveTechnicianRunGuardEntry {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ComputerFqdn,
        [ValidateRange(0, 168)][int]$RunGuardHours
    )

    if ([string]::IsNullOrWhiteSpace($ComputerFqdn) -or $RunGuardHours -le 0) { return $null }

    $found = @{}
    Invoke-TechnicianRunGuardHistoryLock -ArgumentList @($Path, $RunGuardHours, $ComputerFqdn, $found) -ScriptBlock {
        param($LockedPath, $LockedRunGuardHours, $LockedComputerFqdn, $FoundRef)
        $history = Read-TechnicianRunGuardHistory -Path $LockedPath -RunGuardHours $LockedRunGuardHours
        Save-TechnicianRunGuardHistory -Path $LockedPath -History $history
        foreach ($entry in @($history.Entries)) {
            if ([string]$entry.ComputerFqdn -eq $LockedComputerFqdn) { $FoundRef.Value = $entry; break }
        }
    }
    if ($found.ContainsKey('Value') -and (Test-TechnicianRunGuardEntryShouldBlock -Entry $found.Value -RunGuardHours $RunGuardHours)) { return $found.Value }
    return $null
}

function Update-TechnicianRunGuardHistory {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ComputerFqdn,
        [Parameter(Mandatory = $true)][string]$InputComputerName,
        [ValidateRange(0, 168)][int]$RunGuardHours,
        [ValidateSet('Started','Result')][string]$State,
        [AllowNull()]$Result,
        [AllowNull()][string]$JobId
    )

    if ([string]::IsNullOrWhiteSpace($ComputerFqdn) -or $RunGuardHours -le 0) { return }

    Invoke-TechnicianRunGuardHistoryLock -ArgumentList @(
        $Path,
        $ComputerFqdn,
        $InputComputerName,
        $RunGuardHours,
        $State,
        $Result,
        $JobId,
        $cycle,
        $ComputerListPath,
        $script:LotRoot,
        [string]$script:TechnicianIdentity.Account,
        [string]$script:TechnicianIdentity.ComputerName,
        [string]$script:LauncherVersion
    ) -ScriptBlock {
        param(
            $LockedPath,
            $LockedComputerFqdn,
            $LockedInputComputerName,
            $LockedRunGuardHours,
            $LockedState,
            $LockedResult,
            $LockedJobId,
            $LockedCycle,
            $LockedComputerListPath,
            $LockedLotRoot,
            $LockedTechnicianAccount,
            $LockedTechnicianComputer,
            $LockedLauncherVersion
        )

        $history = Read-TechnicianRunGuardHistory -Path $LockedPath -RunGuardHours $LockedRunGuardHours
        $kept = New-Object System.Collections.ArrayList
        $previousEntry = $null
        foreach ($entry in @($history.Entries)) {
            if ([string]$entry.ComputerFqdn -eq $LockedComputerFqdn) { $previousEntry = $entry }
            else { [void]$kept.Add($entry) }
        }

        $nowUtc = (Get-Date).ToUniversalTime()
        $startedUtc = $nowUtc
        if ($LockedState -eq 'Result' -and $previousEntry -and $previousEntry.PSObject.Properties['LastStartedUtc']) {
            $previousStartedUtc = [datetime]::MinValue
            if (ConvertTo-TechnicianRunGuardUtcDateTime -Value $previousEntry.LastStartedUtc -Result ([ref]$previousStartedUtc)) { $startedUtc = $previousStartedUtc.ToUniversalTime() }
        }

        $launcherStatus = ''
        $remoteStatus = ''
        $remoteNextAction = ''
        $exitCode = ''
        $detail = ''
        $psExecLogPath = ''
        $remoteLogsPath = ''
        $failureCategory = ''
        $consecutiveFailureCount = 0
        $retryAfterUtc = ''

        if ($LockedResult) {
            if ($LockedResult.PSObject.Properties['LauncherStatus']) { $launcherStatus = [string]$LockedResult.LauncherStatus }
            if ($LockedResult.PSObject.Properties['RemoteStatus']) { $remoteStatus = [string]$LockedResult.RemoteStatus }
            if ($LockedResult.PSObject.Properties['RemoteNextAction']) { $remoteNextAction = [string]$LockedResult.RemoteNextAction }
            if ($LockedResult.PSObject.Properties['ExitCode']) { $exitCode = [string]$LockedResult.ExitCode }
            if ($LockedResult.PSObject.Properties['Detail']) { $detail = [string]$LockedResult.Detail }
            if ($LockedResult.PSObject.Properties['PsExecLogPath']) { $psExecLogPath = [string]$LockedResult.PsExecLogPath }
            if ($LockedResult.PSObject.Properties['RemoteLogsPath']) { $remoteLogsPath = [string]$LockedResult.RemoteLogsPath }
        }

        $previousFailureCategory = if ($previousEntry -and $previousEntry.PSObject.Properties['FailureCategory']) { [string]$previousEntry.FailureCategory } else { '' }
        $previousFailureCount = 0
        if ($previousEntry -and $previousEntry.PSObject.Properties['ConsecutiveFailureCount']) { [void][int]::TryParse([string]$previousEntry.ConsecutiveFailureCount, [ref]$previousFailureCount) }
        if ($LockedState -eq 'Started') {
            $failureCategory = $previousFailureCategory
            $consecutiveFailureCount = $previousFailureCount
        }
        else {
            $failureCategory = Get-TechnicianRunGuardFailureCategory -LauncherStatus $launcherStatus -RemoteStatus $remoteStatus -Detail $detail -JobErrorMessage $(if ($LockedResult -and $LockedResult.PSObject.Properties['JobErrorMessage']) { [string]$LockedResult.JobErrorMessage } else { '' })
            if (-not [string]::IsNullOrWhiteSpace($failureCategory)) {
                $consecutiveFailureCount = if ($failureCategory -eq $previousFailureCategory) { $previousFailureCount + 1 } else { 1 }
                $retryDelayMinutes = Get-TechnicianRunGuardRetryDelay -FailureCategory $failureCategory -ConsecutiveFailureCount $consecutiveFailureCount -RunGuardHours $LockedRunGuardHours
                if ($retryDelayMinutes -gt 0) { $retryAfterUtc = $nowUtc.AddMinutes($retryDelayMinutes).ToString('o') }
            }
        }

        [void]$kept.Add([pscustomobject]@{
            ComputerFqdn = $LockedComputerFqdn
            InputComputerName = $LockedInputComputerName
            LastStartedUtc = $startedUtc.ToString('o')
            LastUpdatedUtc = $nowUtc.ToString('o')
            LastResultUtc = if ($LockedState -eq 'Result') { $nowUtc.ToString('o') } else { '' }
            ExpiresUtc = $startedUtc.AddHours($LockedRunGuardHours).ToString('o')
            State = $LockedState
            LauncherStatus = $launcherStatus
            RemoteStatus = $remoteStatus
            RemoteNextAction = $remoteNextAction
            ExitCode = $exitCode
            Detail = if ($detail.Length -gt 800) { $detail.Substring(0, 800) } else { $detail }
            PsExecLogPath = $psExecLogPath
            RemoteLogsPath = $remoteLogsPath
            JobErrorMessage = if ($LockedResult -and $LockedResult.PSObject.Properties['JobErrorMessage']) { [string]$LockedResult.JobErrorMessage } else { '' }
            FailureCategory = $failureCategory
            ConsecutiveFailureCount = $consecutiveFailureCount
            RetryAfterUtc = $retryAfterUtc
            JobId = [string]$LockedJobId
            CycleNumber = $LockedCycle
            ComputerListPath = $LockedComputerListPath
            LotRoot = $LockedLotRoot
            TechnicianAccount = $LockedTechnicianAccount
            TechnicianComputer = $LockedTechnicianComputer
            LauncherVersion = $LockedLauncherVersion
            RunGuardHours = $LockedRunGuardHours
        })

        $history.Entries = @($kept.ToArray())
        Save-TechnicianRunGuardHistory -Path $LockedPath -History $history
    }
}
function New-TechnicianRunGuardSkippedResult {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][int]$CycleNumber,
        [Parameter(Mandatory = $true)]$HistoryEntry,
        [ValidateRange(0, 168)][int]$RunGuardHours
    )

    $startedUtc = [datetime]::MinValue
    [void](ConvertTo-TechnicianRunGuardUtcDateTime -Value $HistoryEntry.LastStartedUtc -Result ([ref]$startedUtc))
    $ageHours = if ($startedUtc -gt [datetime]::MinValue) { ((Get-Date).ToUniversalTime() - $startedUtc.ToUniversalTime()).TotalHours } else { 0 }
    $startedText = if ($startedUtc -gt [datetime]::MinValue) { $startedUtc.ToUniversalTime().ToString('o') } else { [string]$HistoryEntry.LastStartedUtc }
    $expiresUtc = [datetime]::MinValue
    if ($HistoryEntry.PSObject.Properties['ExpiresUtc']) { [void](ConvertTo-TechnicianRunGuardUtcDateTime -Value $HistoryEntry.ExpiresUtc -Result ([ref]$expiresUtc)) }
    $expiresText = if ($expiresUtc -gt [datetime]::MinValue) { $expiresUtc.ToUniversalTime().ToString('o') } else { [string]$HistoryEntry.ExpiresUtc }
    $historyState = if ($HistoryEntry.PSObject.Properties['State']) { [string]$HistoryEntry.State } else { '' }
    $historyJobId = if ($HistoryEntry.PSObject.Properties['JobId']) { [string]$HistoryEntry.JobId } else { '' }
    $lastStatus = @([string]$HistoryEntry.RemoteStatus, [string]$HistoryEntry.LauncherStatus) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1
    $isStartedWithoutResult = ($historyState -ne 'Result' -and [string]::IsNullOrWhiteSpace($lastStatus))
    $startedNoResultGuardHours = [math]::Min([double]$RunGuardHours, [double]$script:TechnicianRunGuardStartedNoResultHours)
    $effectiveGuardHours = if ($isStartedWithoutResult) { $startedNoResultGuardHours } else { [double]$RunGuardHours }
    $effectiveExpiresText = $expiresText
    if ($isStartedWithoutResult -and $startedUtc -gt [datetime]::MinValue) { $effectiveExpiresText = $startedUtc.ToUniversalTime().AddHours($startedNoResultGuardHours).ToString('o') }
    $retryAfterText = if ($HistoryEntry.PSObject.Properties['RetryAfterUtc']) { [string]$HistoryEntry.RetryAfterUtc } else { '' }
    $failureCategory = if ($HistoryEntry.PSObject.Properties['FailureCategory']) { [string]$HistoryEntry.FailureCategory } else { '' }
    $failureCount = if ($HistoryEntry.PSObject.Properties['ConsecutiveFailureCount']) { [string]$HistoryEntry.ConsecutiveFailureCount } else { '' }
    if (-not $isStartedWithoutResult -and -not [string]::IsNullOrWhiteSpace($retryAfterText)) { $effectiveExpiresText = $retryAfterText }
    $launcherStatus = if ($isStartedWithoutResult) { 'SKIPPED_BY_TECH_RUN_GUARD_STARTED_NO_RESULT' } else { 'SKIPPED_BY_TECH_RUN_GUARD' }

    return [pscustomobject]@{
        Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        ComputerName = $ComputerName
        CycleNumber = $CycleNumber
        LauncherStatus = $launcherStatus
        RemoteStatus = ''
        RemoteNextAction = if (-not [string]::IsNullOrWhiteSpace($retryAfterText)) { 'WAIT_RETRY_BACKOFF' } else { 'WAIT_RUN_GUARD_EXPIRY' }
        ExitCode = 0
        Detail = ("Technician run guard history skipped launch. FQDN={0}; LastStartedUtc={1}; AgeHours={2:N1}; GuardHours={3}; EffectiveGuardHours={4:N1}; RetryOrExpiryUtc={5}; HistoryState={6}; JobId={7}; LastStatus={8}; FailureCategory={9}; ConsecutiveFailureCount={10}" -f $HistoryEntry.ComputerFqdn,$startedText,$ageHours,$RunGuardHours,$effectiveGuardHours,$effectiveExpiresText,$historyState,$historyJobId,$lastStatus,$failureCategory,$failureCount)
        SetupCacheAction = ''
        DiskCleanupAction = ''
        DiskCleanupFreedGB = ''
        AdvancedDiskCleanupAction = ''
        AdvancedDiskCleanupFreedGB = ''
        DismCleanupAction = ''
        DismCleanupFreedGB = ''
        SetupCompletionRebootAction = ''
        SetupCompletionRebootDetail = ''
        SetupCompletionRebootUserCount = ''
        SetupCompletionRebootUsers = ''
        SetupProfileRepairAction = ''
        SetupProfileRepairDetail = ''
        SetupProfileRepairBlockingSid = ''
        SetupProfileRepairKeptSid = ''
        SetupProfileRepairProfilePath = ''
        SetupProfileRepairBackupPath = ''
        ControlledRebootAction = ''
        ControlledRebootDetail = ''
        ControlledRebootUserCount = ''
        ControlledRebootUsers = ''
        UserRebootNotificationSent = ''
        UserRebootNotificationLang = ''
        UserRebootNotificationMessage = ''
        RemoteLogsPath = ''
        PsExecLogPath = ''
        JobErrorMessage = ''
    }
}

function Get-ComputerList {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Computer list not found: $Path"
    }

    $seen = @{}
    $result = New-Object System.Collections.ArrayList
    foreach ($line in @(Get-Content -LiteralPath $Path -ErrorAction Stop)) {
        $name = ([string]$line).Trim().Trim([char]34)
        if ([string]::IsNullOrWhiteSpace($name) -or $name.StartsWith('#')) { continue }
        $key = $name.ToUpperInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        [void]$result.Add($name)
    }
    return @($result.ToArray())
}

function Test-ComputerListPresentWithRetry {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [ValidateRange(1, 20)][int]$Attempts = 5,
        [ValidateRange(0, 5000)][int]$DelayMilliseconds = 200
    )

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        if (Test-Path -LiteralPath $Path -PathType Leaf) { return $true }
        if ($attempt -lt $Attempts -and $DelayMilliseconds -gt 0) {
            Start-Sleep -Milliseconds $DelayMilliseconds
        }
    }
    return $false
}


function Get-ComputerListStats {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Computer list not found: $Path"
    }

    $rawNames = @(
        Get-Content -LiteralPath $Path -ErrorAction Stop |
            ForEach-Object { ([string]$_).Trim().Trim([char]34) } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.StartsWith('#') }
    )
    $duplicateGroups = @(
        $rawNames |
            Group-Object { ([string]$_).ToUpperInvariant() } |
            Where-Object { $_.Count -gt 1 } |
            Sort-Object -Property @{ Expression = 'Count'; Descending = $true },Name
    )
    $duplicateLineCount = 0
    foreach ($duplicateGroup in $duplicateGroups) { $duplicateLineCount += ([int]$duplicateGroup.Count - 1) }

    return [pscustomobject]@{
        RawLines = [int]$rawNames.Count
        Unique = [int](@($rawNames | ForEach-Object { ([string]$_).ToUpperInvariant() } | Select-Object -Unique).Count)
        DuplicateGroups = [int]$duplicateGroups.Count
        DuplicateLines = [int]$duplicateLineCount
        DuplicateSamples = [string](($duplicateGroups | Select-Object -First 10 | ForEach-Object { $_.Group[0] }) -join ', ')
    }
}

function Remove-DuplicateComputerListEntries {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Computer list not found: $Path"
    }

    $lines = @(Get-Content -LiteralPath $Path -ErrorAction Stop)
    $seen = @{}
    $output = New-Object System.Collections.Generic.List[string]
    $duplicateGroups = @{}
    $duplicateLines = 0

    foreach ($line in $lines) {
        $rawLine = [string]$line
        $name = $rawLine.Trim().Trim([char]34)
        if ([string]::IsNullOrWhiteSpace($name) -or $name.StartsWith('#')) {
            [void]$output.Add($rawLine)
            continue
        }

        $key = $name.ToUpperInvariant()
        if ($seen.ContainsKey($key)) {
            $duplicateLines++
            if (-not $duplicateGroups.ContainsKey($key)) { $duplicateGroups[$key] = $seen[$key] }
            continue
        }

        $seen[$key] = $name
        [void]$output.Add($rawLine)
    }

    if ($duplicateLines -le 0) {
        return [pscustomobject]@{ Changed = $false; DuplicateGroups = 0; DuplicateLines = 0; DuplicateSamples = ''; BackupPath = '' }
    }

    $backupPath = "{0}.dedup_{1}.bak" -f $Path,(Get-Date -Format 'yyyyMMdd-HHmmss')
    Copy-Item -LiteralPath $Path -Destination $backupPath -Force -ErrorAction Stop
    Set-Content -LiteralPath $Path -Value $output -Encoding UTF8 -Force

    $samples = @($duplicateGroups.Values | Select-Object -First 10)
    return [pscustomobject]@{
        Changed = $true
        DuplicateGroups = [int]$duplicateGroups.Count
        DuplicateLines = [int]$duplicateLines
        DuplicateSamples = [string]($samples -join ', ')
        BackupPath = [string]$backupPath
    }
}

function Get-ComputerListKey {
    param([Parameter(Mandatory = $true)][string]$ComputerName)

    return ($ComputerName.Trim().Trim([char]34).Split('.')[0]).ToUpperInvariant()
}

function Test-AlreadyWindows11CycleResult {
    param([Parameter(Mandatory = $true)][psobject]$Result)

    $launcherStatus = if ($Result.PSObject.Properties['LauncherStatus']) { [string]$Result.LauncherStatus } else { '' }
    $remoteStatus = if ($Result.PSObject.Properties['RemoteStatus']) { [string]$Result.RemoteStatus } else { '' }
    return ($launcherStatus -eq 'ALREADY_WINDOWS11' -or $remoteStatus -eq 'ALREADY_WINDOWS11')
}

function Move-AlreadyWindows11ComputersFromList {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerListPath,
        [Parameter(Mandatory = $true)][object[]]$CycleSummary
    )

    $alreadyWindows11 = @(
        $CycleSummary |
            Where-Object { $_ -and (Test-AlreadyWindows11CycleResult -Result $_) } |
            ForEach-Object { if ($_.PSObject.Properties['ComputerName']) { [string]$_.ComputerName } } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
    )

    if ($alreadyWindows11.Count -eq 0) {
        return [pscustomobject]@{ Moved = 0; AlreadyWindows11Path = ''; Detail = 'No already-Windows11 computer detected in this cycle.' }
    }

    $moveKeys = @{}
    foreach ($computer in $alreadyWindows11) {
        $key = Get-ComputerListKey -ComputerName $computer
        if (-not [string]::IsNullOrWhiteSpace($key) -and -not $moveKeys.ContainsKey($key)) { $moveKeys[$key] = $computer.Trim() }
    }

    $listLines = @(Get-Content -LiteralPath $ComputerListPath -ErrorAction Stop)
    $remainingLines = New-Object System.Collections.Generic.List[string]
    $movedFromList = New-Object System.Collections.Generic.List[string]

    foreach ($line in $listLines) {
        $trimmed = ([string]$line).Trim().Trim([char]34)
        if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed.StartsWith('#')) {
            [void]$remainingLines.Add($line)
            continue
        }

        $key = Get-ComputerListKey -ComputerName $trimmed
        if ($moveKeys.ContainsKey($key)) {
            [void]$movedFromList.Add($trimmed)
            continue
        }

        [void]$remainingLines.Add($line)
    }

    if ($movedFromList.Count -eq 0) {
        return [pscustomobject]@{ Moved = 0; AlreadyWindows11Path = ''; Detail = 'Already-Windows11 computers were detected, but none were still present in Computers.txt.' }
    }

    $computerListDir = Split-Path -Parent $ComputerListPath
    if ([string]::IsNullOrWhiteSpace($computerListDir)) { $computerListDir = '.' }
    $alreadyWindows11Path = Join-Path $computerListDir 'ComputersAlreadyW11.txt'

    $existingKeys = @{}
    if (Test-Path -LiteralPath $alreadyWindows11Path) {
        foreach ($line in @(Get-Content -LiteralPath $alreadyWindows11Path -ErrorAction SilentlyContinue)) {
            $trimmed = ([string]$line).Trim().Trim([char]34)
            if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed.StartsWith('#')) { continue }
            $key = Get-ComputerListKey -ComputerName $trimmed
            if (-not $existingKeys.ContainsKey($key)) { $existingKeys[$key] = $true }
        }
    }

    $appendLines = New-Object System.Collections.Generic.List[string]
    foreach ($computer in $movedFromList) {
        $key = Get-ComputerListKey -ComputerName $computer
        if (-not $existingKeys.ContainsKey($key)) {
            [void]$appendLines.Add($computer)
            $existingKeys[$key] = $true
        }
    }

    $tmpComputerListPath = '{0}.tmp.{1}.txt' -f $ComputerListPath,([guid]::NewGuid().ToString('N'))
    try {
        Set-Content -LiteralPath $tmpComputerListPath -Value $remainingLines -Encoding ASCII -Force
        Move-Item -LiteralPath $tmpComputerListPath -Destination $ComputerListPath -Force
    }
    finally {
        Remove-Item -LiteralPath $tmpComputerListPath -Force -ErrorAction SilentlyContinue
    }

    if ($appendLines.Count -gt 0) {
        Add-Content -LiteralPath $alreadyWindows11Path -Value $appendLines -Encoding ASCII
    }
    elseif (-not (Test-Path -LiteralPath $alreadyWindows11Path)) {
        New-Item -ItemType File -Path $alreadyWindows11Path -Force | Out-Null
    }

    return [pscustomobject]@{
        Moved = $movedFromList.Count
        AlreadyWindows11Path = $alreadyWindows11Path
        Detail = ('Moved {0} computer(s) from Computers.txt to ComputersAlreadyW11.txt.' -f $movedFromList.Count)
    }
}

function Test-BooleanLikeTrue {
    param([AllowNull()][object]$Value)

    if ($Value -eq $true) { return $true }
    $text = ([string]$Value).Trim()
    return ($text -in @('True','true','1','YES','Yes','yes','OUI','Oui','oui'))
}

function Get-AdInventoryMap {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $false)][string]$NameColumn
    )

    $map = @{}
    if ([string]::IsNullOrWhiteSpace($Path)) { return $map }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $map }

    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { return $map }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        $candidateColumns = @('ComputerName','computerName','DNSHostName','dnsHostName','Name','name')
        $first = $rows | Select-Object -First 1
        foreach ($candidate in $candidateColumns) {
            if ($first.PSObject.Properties.Name -contains $candidate) {
                $NameColumn = $candidate
                break
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        throw 'Unable to infer the AD inventory device name column. Use -AdInventoryNameColumn.'
    }

    foreach ($row in $rows) {
        $value = [string]$row.$NameColumn
        if ([string]::IsNullOrWhiteSpace($value)) { continue }

        $short = (Get-ComputerListKey -ComputerName $value)
        if ([string]::IsNullOrWhiteSpace($short)) { continue }

        $present = $true
        if ($row.PSObject.Properties.Name -contains 'ADInventoryPresent') {
            $present = Test-BooleanLikeTrue -Value $row.ADInventoryPresent
        }
        if (-not $present) { continue }

        if (-not $map.ContainsKey($short)) { $map[$short] = $row }
    }

    return $map
}

function Invoke-FullAdInventoryExport {
    param(
        [Parameter(Mandatory = $true)][string]$ExportScriptPath,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [Parameter(Mandatory = $true)][string]$ComputerListPath,
        [Parameter(Mandatory = $false)][string]$Domain
    )

    try {
        if (-not (Test-Path -LiteralPath $ExportScriptPath -PathType Leaf)) {
            throw "SmartM365-Windows11Upgrade-Export-ADDevicesCsv.ps1 not found: $ExportScriptPath"
        }

        $args = @(
            '-NoProfile',
            '-ExecutionPolicy', 'Bypass',
            '-File', $ExportScriptPath,
            '-OutputPath', $OutputPath,
            '-ComputerListPath', $ComputerListPath,
            '-ForceRefresh'
        )
        if (-not [string]::IsNullOrWhiteSpace($Domain)) {
            $args += '-Domain'
            $args += $Domain
        }

        $output = & powershell.exe @args 2>&1
        $exitCode = $LASTEXITCODE
        $output | Out-File -LiteralPath $LogPath -Encoding UTF8 -Force

        if ($exitCode -ne 0) {
            throw "SmartM365-Windows11Upgrade-Export-ADDevicesCsv.ps1 exited with code $exitCode. Log=$LogPath"
        }
        if (-not (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
            throw "AD CSV was not created: $OutputPath"
        }

        $map = Get-AdInventoryMap -Path $OutputPath -NameColumn 'ComputerName'
        return [pscustomobject]@{ Success = $true; CsvPath = $OutputPath; LogPath = $LogPath; InventoryMap = $map; Error = '' }
    }
    catch {
        return [pscustomobject]@{ Success = $false; CsvPath = $OutputPath; LogPath = $LogPath; InventoryMap = @{}; Error = $_.Exception.Message }
    }
}

function Test-AdInventoryWindows11 {
    param([Parameter(Mandatory = $true)]$AdRow)

    $operatingSystem = if ($AdRow.PSObject.Properties['OperatingSystem']) { [string]$AdRow.OperatingSystem } else { '' }
    return ($operatingSystem -match '(?i)\bWindows\s+11\b')
}

function Get-AlreadyWindows11RowsFromAdInventory {
    param(
        [Parameter(Mandatory = $true)][string[]]$ComputerNames,
        [Parameter(Mandatory = $true)][hashtable]$AdInventoryMap,
        [Parameter(Mandatory = $true)][string]$AdInventoryCsv
    )

    foreach ($computer in $ComputerNames) {
        $key = Get-ComputerListKey -ComputerName $computer
        if ([string]::IsNullOrWhiteSpace($key) -or -not $AdInventoryMap.ContainsKey($key)) { continue }

        $adRow = $AdInventoryMap[$key]
        if (-not (Test-AdInventoryWindows11 -AdRow $adRow)) { continue }

        [pscustomobject]@{
            Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            ComputerName = $computer
            CycleNumber = 0
            LauncherStatus = 'ALREADY_WINDOWS11'
            RemoteStatus = 'ALREADY_WINDOWS11'
            RemoteNextAction = 'NO_ACTION_AD_ALREADY_WINDOWS11'
            ExitCode = 0
            Detail = 'AD inventory already reports Windows 11; PsExec launch skipped.'
            JobErrorMessage = ''
            ADInventoryPresent = $true
            ADDomain = if ($adRow.PSObject.Properties['ADDomain']) { [string]$adRow.ADDomain } else { '' }
            ADEnabled = if ($adRow.PSObject.Properties['Enabled']) { [string]$adRow.Enabled } else { '' }
            ADDNSHostName = if ($adRow.PSObject.Properties['DNSHostName']) { [string]$adRow.DNSHostName } else { '' }
            ADDistinguishedName = if ($adRow.PSObject.Properties['DistinguishedName']) { [string]$adRow.DistinguishedName } else { '' }
            ADOperatingSystem = if ($adRow.PSObject.Properties['OperatingSystem']) { [string]$adRow.OperatingSystem } else { '' }
            ADOperatingSystemVersion = if ($adRow.PSObject.Properties['OperatingSystemVersion']) { [string]$adRow.OperatingSystemVersion } else { '' }
            ADLastLogonTimestampUtc = if ($adRow.PSObject.Properties['LastLogonTimestampUtc']) { [string]$adRow.LastLogonTimestampUtc } else { '' }
            ADInventoryCsv = $AdInventoryCsv
        }
    }
}
function Add-AdInventoryFieldsToResult {
    param(
        [Parameter(Mandatory = $true)]$Result,
        [AllowNull()][hashtable]$AdInventoryMap,
        [AllowNull()][string]$AdInventoryCsv
    )

    if ($null -eq $Result) { return $Result }
    if (-not $Result.PSObject.Properties['ComputerName']) { return $Result }

    $computerKey = Get-ComputerListKey -ComputerName ([string]$Result.ComputerName)
    $hasInventory = ($null -ne $AdInventoryMap -and $AdInventoryMap.Count -gt 0)
    $hasAdRow = ($hasInventory -and -not [string]::IsNullOrWhiteSpace($computerKey) -and $AdInventoryMap.ContainsKey($computerKey))

    if ($hasAdRow) {
        $adRow = $AdInventoryMap[$computerKey]
        $fields = [ordered]@{
            ADInventoryPresent = $true
            ADDomain = if ($adRow.PSObject.Properties['ADDomain']) { [string]$adRow.ADDomain } else { '' }
            ADEnabled = if ($adRow.PSObject.Properties['Enabled']) { [string]$adRow.Enabled } else { '' }
            ADDNSHostName = if ($adRow.PSObject.Properties['DNSHostName']) { [string]$adRow.DNSHostName } else { '' }
            ADDistinguishedName = if ($adRow.PSObject.Properties['DistinguishedName']) { [string]$adRow.DistinguishedName } else { '' }
            ADOperatingSystem = if ($adRow.PSObject.Properties['OperatingSystem']) { [string]$adRow.OperatingSystem } else { '' }
            ADOperatingSystemVersion = if ($adRow.PSObject.Properties['OperatingSystemVersion']) { [string]$adRow.OperatingSystemVersion } else { '' }
            ADLastLogonTimestampUtc = if ($adRow.PSObject.Properties['LastLogonTimestampUtc']) { [string]$adRow.LastLogonTimestampUtc } else { '' }
            ADInventoryCsv = $AdInventoryCsv
        }
    }
    elseif ($hasInventory) {
        $fields = [ordered]@{
            ADInventoryPresent = $false
            ADDomain = ''
            ADEnabled = ''
            ADDNSHostName = ''
            ADDistinguishedName = ''
            ADOperatingSystem = ''
            ADOperatingSystemVersion = ''
            ADLastLogonTimestampUtc = ''
            ADInventoryCsv = $AdInventoryCsv
        }
    }
    else {
        return $Result
    }

    foreach ($fieldName in $fields.Keys) {
        $Result | Add-Member -NotePropertyName $fieldName -NotePropertyValue $fields[$fieldName] -Force
    }

    return $Result
}
function Get-IntuneInventoryMap {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $false)][string]$NameColumn
    )

    $map = @{}
    if ([string]::IsNullOrWhiteSpace($Path)) { return $map }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $map }

    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { return $map }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        $candidateColumns = @('ComputerName','computerName','DeviceName','deviceName','ManagedDeviceName','managedDeviceName','Name','name')
        $first = $rows | Select-Object -First 1
        foreach ($candidate in $candidateColumns) {
            if ($first.PSObject.Properties.Name -contains $candidate) {
                $NameColumn = $candidate
                break
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        throw 'Unable to infer the Intune inventory device name column. Use -IntuneInventoryNameColumn.'
    }

    foreach ($row in $rows) {
        $value = [string]$row.$NameColumn
        if ([string]::IsNullOrWhiteSpace($value)) { continue }

        $short = (Get-ComputerListKey -ComputerName $value)
        if ([string]::IsNullOrWhiteSpace($short)) { continue }

        $present = $true
        if ($row.PSObject.Properties.Name -contains 'IntuneInventoryPresent') {
            $present = Test-BooleanLikeTrue -Value $row.IntuneInventoryPresent
        }
        if (-not $present) { continue }

        if (-not $map.ContainsKey($short)) {
            $map[$short] = $row
        }
        else {
            $current = $map[$short]
            $currentLastSync = [datetime]::MinValue
            $candidateLastSync = [datetime]::MinValue
            if ($current.PSObject.Properties['LastSyncDateTime']) { [datetime]::TryParse([string]$current.LastSyncDateTime, [ref]$currentLastSync) | Out-Null }
            if ($row.PSObject.Properties['LastSyncDateTime']) { [datetime]::TryParse([string]$row.LastSyncDateTime, [ref]$candidateLastSync) | Out-Null }
            if ($candidateLastSync -gt $currentLastSync) { $map[$short] = $row }
        }
    }

    return $map
}

function Invoke-FullIntuneInventoryExport {
    param(
        [Parameter(Mandatory = $true)][string]$ExportScriptPath,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [Parameter(Mandatory = $true)][string]$ComputerListPath,
        [Parameter(Mandatory = $false)][int]$PageSize = 999,
        [Parameter(Mandatory = $false)][string]$TenantId
    )

    try {
        if (-not (Test-Path -LiteralPath $ExportScriptPath -PathType Leaf)) {
            throw "SmartM365-Windows11Upgrade-Export-IntuneDevicesCsv.ps1 not found: $ExportScriptPath"
        }

        $args = @(
            '-NoProfile',
            '-ExecutionPolicy', 'Bypass',
            '-File', $ExportScriptPath,
            '-OutputPath', $OutputPath,
            '-ComputerListPath', $ComputerListPath,
            '-PageSize', ([string]$PageSize),
            '-ForceRefresh'
        )
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
            $args += '-TenantId'
            $args += $TenantId
        }

        $output = & powershell.exe @args 2>&1
        $exitCode = $LASTEXITCODE
        $output | Out-File -LiteralPath $LogPath -Encoding UTF8 -Force

        if ($exitCode -ne 0) {
            throw "SmartM365-Windows11Upgrade-Export-IntuneDevicesCsv.ps1 exited with code $exitCode. Log=$LogPath"
        }
        if (-not (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
            throw "Intune CSV was not created: $OutputPath"
        }

        $map = Get-IntuneInventoryMap -Path $OutputPath -NameColumn 'ComputerName'
        return [pscustomobject]@{ Success = $true; CsvPath = $OutputPath; LogPath = $LogPath; InventoryMap = $map; Error = '' }
    }
    catch {
        return [pscustomobject]@{ Success = $false; CsvPath = $OutputPath; LogPath = $LogPath; InventoryMap = @{}; Error = $_.Exception.Message }
    }
}

function Test-IntuneInventoryWindows11 {
    param([Parameter(Mandatory = $true)]$IntuneRow)

    $operatingSystem = if ($IntuneRow.PSObject.Properties['OperatingSystem']) { [string]$IntuneRow.OperatingSystem } else { '' }
    if ($operatingSystem -match '(?i)\bWindows\s+11\b') { return $true }

    $osVersion = if ($IntuneRow.PSObject.Properties['OSVersion']) { [string]$IntuneRow.OSVersion } else { '' }
    if ($osVersion -match '^(\d+)\.(\d+)\.(\d+)') {
        $build = 0
        if ([int]::TryParse($matches[3], [ref]$build) -and $build -ge 22000) { return $true }
    }

    return $false
}

function Get-AlreadyWindows11RowsFromIntuneInventory {
    param(
        [Parameter(Mandatory = $true)][string[]]$ComputerNames,
        [Parameter(Mandatory = $true)][hashtable]$IntuneInventoryMap,
        [Parameter(Mandatory = $true)][string]$IntuneInventoryCsv
    )

    foreach ($computer in $ComputerNames) {
        $key = Get-ComputerListKey -ComputerName $computer
        if ([string]::IsNullOrWhiteSpace($key) -or -not $IntuneInventoryMap.ContainsKey($key)) { continue }

        $intuneRow = $IntuneInventoryMap[$key]
        if (-not (Test-IntuneInventoryWindows11 -IntuneRow $intuneRow)) { continue }

        [pscustomobject]@{
            Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            ComputerName = $computer
            CycleNumber = 0
            LauncherStatus = 'ALREADY_WINDOWS11'
            RemoteStatus = 'ALREADY_WINDOWS11'
            RemoteNextAction = 'NO_ACTION_INTUNE_ALREADY_WINDOWS11'
            ExitCode = 0
            Detail = 'Intune inventory already reports Windows 11; PsExec launch skipped.'
            JobErrorMessage = ''
            IntuneInventoryPresent = $true
            IntuneDeviceName = if ($intuneRow.PSObject.Properties['DeviceName']) { [string]$intuneRow.DeviceName } else { '' }
            IntuneManagedDeviceName = if ($intuneRow.PSObject.Properties['ManagedDeviceName']) { [string]$intuneRow.ManagedDeviceName } else { '' }
            IntuneManagedDeviceId = if ($intuneRow.PSObject.Properties['IntuneManagedDeviceId']) { [string]$intuneRow.IntuneManagedDeviceId } else { '' }
            IntuneAzureADDeviceId = if ($intuneRow.PSObject.Properties['AzureADDeviceId']) { [string]$intuneRow.AzureADDeviceId } else { '' }
            IntuneOperatingSystem = if ($intuneRow.PSObject.Properties['OperatingSystem']) { [string]$intuneRow.OperatingSystem } else { '' }
            IntuneOSVersion = if ($intuneRow.PSObject.Properties['OSVersion']) { [string]$intuneRow.OSVersion } else { '' }
            IntuneLastSyncDateTime = if ($intuneRow.PSObject.Properties['LastSyncDateTime']) { [string]$intuneRow.LastSyncDateTime } else { '' }
            IntuneUserPrincipalName = if ($intuneRow.PSObject.Properties['UserPrincipalName']) { [string]$intuneRow.UserPrincipalName } else { '' }
            IntuneComplianceState = if ($intuneRow.PSObject.Properties['ComplianceState']) { [string]$intuneRow.ComplianceState } else { '' }
            IntuneManagementState = if ($intuneRow.PSObject.Properties['ManagementState']) { [string]$intuneRow.ManagementState } else { '' }
            IntuneInventoryCsv = $IntuneInventoryCsv
        }
    }
}

function Set-CycleNumberOnRows {
    param(
        [AllowNull()][object[]]$Rows,
        [Parameter(Mandatory = $true)][int]$CycleNumber
    )

    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $row | Add-Member -NotePropertyName CycleNumber -NotePropertyValue $CycleNumber -Force
    }

    return @($Rows)
}

function Invoke-Windows11InventoryPreCycleRefresh {
    param([Parameter(Mandatory = $true)][int]$CycleNumber)

    $movedFromIntune = 0
    $movedFromAd = 0
    $currentComputers = @(Get-ComputerList -Path $ComputerListPath)
    if ($currentComputers.Count -eq 0) {
        return [pscustomobject]@{ RemainingComputers = 0; MovedFromIntune = 0; MovedFromAd = 0 }
    }

    $script:IntuneInventoryMap = @{}
    if (-not [string]::IsNullOrWhiteSpace($IntuneInventoryCsv)) {
        try {
            $currentIntuneScopeKeys = @(Get-ComputerList -Path $ComputerListPath | ForEach-Object { Get-ComputerListKey -ComputerName $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            $intuneRefreshAgeHours = if ($script:IntuneInventoryLastRefreshUtc -gt [datetime]::MinValue) { ((Get-Date).ToUniversalTime() - $script:IntuneInventoryLastRefreshUtc).TotalHours } else { [double]::PositiveInfinity }
            $intuneScopeStillCovered = (@($currentIntuneScopeKeys | Where-Object { $script:IntuneInventoryRefreshScopeKeys -notcontains $_ }).Count -eq 0)
            $reuseInRunIntuneInventory = ($script:IntuneInventoryLastRefreshUtc -gt [datetime]::MinValue -and $intuneRefreshAgeHours -lt $script:IntuneInventoryFreshnessHours -and $intuneScopeStillCovered -and (Test-Path -LiteralPath $IntuneInventoryCsv -PathType Leaf))
            if ($DryRun) {
                Write-Host ("Cycle {0}: DryRun: skipping automatic Intune scoped refresh." -f $CycleNumber) -ForegroundColor Yellow
                if (Test-Path -LiteralPath $IntuneInventoryCsv -PathType Leaf) {
                    $script:IntuneInventoryMap = Get-IntuneInventoryMap -Path $IntuneInventoryCsv -NameColumn $IntuneInventoryNameColumn
                }
            }
            elseif ($EffectiveSkipIntuneInventoryRefresh) {
                $skipReason = if ($IntuneInventoryUsesRecentRootCsv) { "recent root cache (maximum age: $($script:IntuneInventoryFreshnessHours) hour(s))" } else { 'requested option' }
                Write-Host ("Cycle {0}: Intune inventory refresh skipped because of {1}. Using existing CSV when available: {2}" -f $CycleNumber,$skipReason,$IntuneInventoryCsv) -ForegroundColor Yellow
                if (Test-Path -LiteralPath $IntuneInventoryCsv -PathType Leaf) {
                    $script:IntuneInventoryMap = Get-IntuneInventoryMap -Path $IntuneInventoryCsv -NameColumn $IntuneInventoryNameColumn
                }
            }
            elseif ($reuseInRunIntuneInventory) {
                Write-Host ("Cycle {0}: reusing in-run Intune inventory. AgeHours={1:N2}; TTLHours={2}; ScopeStillCovered=True; CSV={3}" -f $CycleNumber,$intuneRefreshAgeHours,$script:IntuneInventoryFreshnessHours,$IntuneInventoryCsv) -ForegroundColor DarkGray
                $script:IntuneInventoryMap = Get-IntuneInventoryMap -Path $IntuneInventoryCsv -NameColumn $IntuneInventoryNameColumn
            }
            else {
                $cycleIntuneInventoryLogPath = Join-Path $ReportRoot ("DevicesIntune_Cycle{0}Refresh_{1}.log" -f $CycleNumber,(Get-Date -Format 'yyyyMMdd_HHmmss'))
                Write-Host ("Cycle {0}: refreshing Intune inventory scoped to current Computers.txt ({1} computer(s))..." -f $CycleNumber,$currentComputers.Count) -ForegroundColor Yellow
                $cycleIntuneInventory = Invoke-FullIntuneInventoryExport `
                    -ExportScriptPath $script:ExportIntuneScriptPath `
                    -OutputPath $IntuneInventoryCsv `
                    -LogPath $cycleIntuneInventoryLogPath `
                    -ComputerListPath $ComputerListPath `
                    -PageSize $IntuneInventoryPageSize `
                    -TenantId $IntuneTenantId
                if ($cycleIntuneInventory.Success) {
                    $script:IntuneInventoryMap = $cycleIntuneInventory.InventoryMap
                    $script:IntuneInventoryLastRefreshUtc = (Get-Date).ToUniversalTime()
                    $script:IntuneInventoryRefreshScopeKeys = @($currentIntuneScopeKeys)
                    Write-Host ("Cycle {0}: Intune inventory refreshed. Devices={1}; CSV={2}" -f $CycleNumber,$script:IntuneInventoryMap.Count,$cycleIntuneInventory.CsvPath) -ForegroundColor Green
                }
                else {
                    Write-Host ("WARNING: Cycle {0}: Intune inventory refresh failed: {1}" -f $CycleNumber,$cycleIntuneInventory.Error) -ForegroundColor Yellow
                    if (Test-Path -LiteralPath $IntuneInventoryCsv -PathType Leaf) {
                        Write-Host ("Cycle {0}: using existing Intune inventory CSV despite refresh failure: {1}" -f $CycleNumber,$IntuneInventoryCsv) -ForegroundColor Yellow
                        $script:IntuneInventoryMap = Get-IntuneInventoryMap -Path $IntuneInventoryCsv -NameColumn $IntuneInventoryNameColumn
                    }
                }
            }

            if ($script:IntuneInventoryMap.Count -gt 0) {
                $currentComputers = @(Get-ComputerList -Path $ComputerListPath)
                $alreadyWindows11FromIntune = @(Get-AlreadyWindows11RowsFromIntuneInventory -ComputerNames $currentComputers -IntuneInventoryMap $script:IntuneInventoryMap -IntuneInventoryCsv $IntuneInventoryCsv)
                $alreadyWindows11FromIntune = @(Set-CycleNumberOnRows -Rows $alreadyWindows11FromIntune -CycleNumber $CycleNumber)
                if ($alreadyWindows11FromIntune.Count -gt 0) {
                    $intuneAlreadyPath = Join-Path $ReportRoot ("DevicesIntune_AlreadyWindows11_cycle{0}_{1}.csv" -f $CycleNumber,(Get-Date -Format 'yyyyMMdd_HHmmss'))
                    @($alreadyWindows11FromIntune) | Export-Csv -LiteralPath $intuneAlreadyPath -NoTypeInformation -Encoding UTF8
                    if ($DryRun) {
                        Write-Host ("Cycle {0}: DryRun: Intune inventory detected {1} already-Windows11 computer(s). No Computers.txt change. CSV={2}" -f $CycleNumber,$alreadyWindows11FromIntune.Count,$intuneAlreadyPath) -ForegroundColor Yellow
                    }
                    else {
                        $preMoveResult = Move-AlreadyWindows11ComputersFromList -ComputerListPath $ComputerListPath -CycleSummary $alreadyWindows11FromIntune
                        $movedFromIntune = [int]$preMoveResult.Moved
                        if ($preMoveResult.Moved -gt 0) {
                            Write-Host ("Cycle {0}: Intune inventory moved {1} already-Windows11 computer(s) from Computers.txt to {2}. Evidence={3}" -f $CycleNumber,$preMoveResult.Moved,$preMoveResult.AlreadyWindows11Path,$intuneAlreadyPath) -ForegroundColor Green
                        }
                        else {
                            Write-Host ("Cycle {0}: Intune inventory detected already-Windows11 computer(s), but none were still present in Computers.txt. Evidence={1}" -f $CycleNumber,$intuneAlreadyPath) -ForegroundColor DarkGray
                        }
                    }
                }
                else {
                    Write-Host ("Cycle {0}: Intune inventory precheck found no Windows 11 computer in current Computers.txt. CSV={1}" -f $CycleNumber,$IntuneInventoryCsv) -ForegroundColor DarkGray
                }
            }
            else {
                Write-Host ("Cycle {0}: Intune inventory precheck skipped: no Intune rows loaded. CSV={1}" -f $CycleNumber,$IntuneInventoryCsv) -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host ("WARN: Cycle {0}: Intune inventory precheck failed: {1}" -f $CycleNumber,$_.Exception.Message) -ForegroundColor Yellow
        }
    }

    $script:AdInventoryMap = @{}
    $currentComputers = @(Get-ComputerList -Path $ComputerListPath)
    if ($currentComputers.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($AdInventoryCsv)) {
        try {
            $currentAdScopeKeys = @($currentComputers | ForEach-Object { Get-ComputerListKey -ComputerName $_ } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
            $adRefreshAgeHours = if ($script:AdInventoryLastRefreshUtc -gt [datetime]::MinValue) { ((Get-Date).ToUniversalTime() - $script:AdInventoryLastRefreshUtc).TotalHours } else { [double]::PositiveInfinity }
            $adScopeStillCovered = (@($currentAdScopeKeys | Where-Object { $script:AdInventoryRefreshScopeKeys -notcontains $_ }).Count -eq 0)
            $reuseInRunAdInventory = ($script:AdInventoryLastRefreshUtc -gt [datetime]::MinValue -and $adRefreshAgeHours -lt $script:AdInventoryFreshnessHours -and $adScopeStillCovered -and (Test-Path -LiteralPath $AdInventoryCsv -PathType Leaf))
            if ($DryRun) {
                Write-Host ("Cycle {0}: DryRun: skipping automatic AD scoped refresh." -f $CycleNumber) -ForegroundColor Yellow
                if (Test-Path -LiteralPath $AdInventoryCsv -PathType Leaf) {
                    $script:AdInventoryMap = Get-AdInventoryMap -Path $AdInventoryCsv -NameColumn $AdInventoryNameColumn
                }
            }
            elseif ($EffectiveSkipAdInventoryRefresh) {
                $skipReason = if ($AdInventoryUsesRecentRootCsv) { "recent root cache (maximum age: $($script:AdInventoryFreshnessHours) hour(s))" } else { 'requested option' }
                Write-Host ("Cycle {0}: AD inventory refresh skipped because of {1}. Using existing CSV when available: {2}" -f $CycleNumber,$skipReason,$AdInventoryCsv) -ForegroundColor Yellow
                if (Test-Path -LiteralPath $AdInventoryCsv -PathType Leaf) {
                    $script:AdInventoryMap = Get-AdInventoryMap -Path $AdInventoryCsv -NameColumn $AdInventoryNameColumn
                }
            }
            elseif ($reuseInRunAdInventory) {
                Write-Host ("Cycle {0}: reusing in-run AD inventory. AgeHours={1:N2}; TTLHours={2}; ScopeStillCovered=True; CSV={3}" -f $CycleNumber,$adRefreshAgeHours,$script:AdInventoryFreshnessHours,$AdInventoryCsv) -ForegroundColor DarkGray
                $script:AdInventoryMap = Get-AdInventoryMap -Path $AdInventoryCsv -NameColumn $AdInventoryNameColumn
            }
            else {
                $cycleAdInventoryLogPath = Join-Path $ReportRoot ("DevicesAD_Cycle{0}Refresh_{1}.log" -f $CycleNumber,(Get-Date -Format 'yyyyMMdd_HHmmss'))
                $cycleAdScope = if ([string]::IsNullOrWhiteSpace($AdDomain)) { 'Current AD forest, limited to current Computers.txt' } else { "Domain=$AdDomain, limited to current Computers.txt" }
                Write-Host ("Cycle {0}: refreshing AD inventory scoped to current Computers.txt ({1} computer(s)). Scope={2}..." -f $CycleNumber,$currentComputers.Count,$cycleAdScope) -ForegroundColor Yellow
                $cycleAdInventory = Invoke-FullAdInventoryExport `
                    -ExportScriptPath $script:ExportAdScriptPath `
                    -OutputPath $AdInventoryCsv `
                    -LogPath $cycleAdInventoryLogPath `
                    -ComputerListPath $ComputerListPath `
                    -Domain $AdDomain
                if ($cycleAdInventory.Success) {
                    $script:AdInventoryMap = $cycleAdInventory.InventoryMap
                    $script:AdInventoryLastRefreshUtc = (Get-Date).ToUniversalTime()
                    $script:AdInventoryRefreshScopeKeys = @($currentAdScopeKeys)
                    Write-Host ("Cycle {0}: AD inventory refreshed. Devices={1}; CSV={2}" -f $CycleNumber,$script:AdInventoryMap.Count,$cycleAdInventory.CsvPath) -ForegroundColor Green
                }
                else {
                    Write-Host ("WARNING: Cycle {0}: AD inventory refresh failed: {1}" -f $CycleNumber,$cycleAdInventory.Error) -ForegroundColor Yellow
                    if (Test-Path -LiteralPath $AdInventoryCsv -PathType Leaf) {
                        Write-Host ("Cycle {0}: using existing AD inventory CSV despite refresh failure: {1}" -f $CycleNumber,$AdInventoryCsv) -ForegroundColor Yellow
                        $script:AdInventoryMap = Get-AdInventoryMap -Path $AdInventoryCsv -NameColumn $AdInventoryNameColumn
                    }
                }
            }

            if ($script:AdInventoryMap.Count -gt 0) {
                $currentComputers = @(Get-ComputerList -Path $ComputerListPath)
                $alreadyWindows11FromAd = @(Get-AlreadyWindows11RowsFromAdInventory -ComputerNames $currentComputers -AdInventoryMap $script:AdInventoryMap -AdInventoryCsv $AdInventoryCsv)
                $alreadyWindows11FromAd = @(Set-CycleNumberOnRows -Rows $alreadyWindows11FromAd -CycleNumber $CycleNumber)
                if ($alreadyWindows11FromAd.Count -gt 0) {
                    $adAlreadyPath = Join-Path $ReportRoot ("DevicesAD_AlreadyWindows11_cycle{0}_{1}.csv" -f $CycleNumber,(Get-Date -Format 'yyyyMMdd_HHmmss'))
                    @($alreadyWindows11FromAd) | Export-Csv -LiteralPath $adAlreadyPath -NoTypeInformation -Encoding UTF8
                    if ($DryRun) {
                        Write-Host ("Cycle {0}: DryRun: AD inventory detected {1} already-Windows11 computer(s). No Computers.txt change. CSV={2}" -f $CycleNumber,$alreadyWindows11FromAd.Count,$adAlreadyPath) -ForegroundColor Yellow
                    }
                    else {
                        $preMoveResult = Move-AlreadyWindows11ComputersFromList -ComputerListPath $ComputerListPath -CycleSummary $alreadyWindows11FromAd
                        $movedFromAd = [int]$preMoveResult.Moved
                        if ($preMoveResult.Moved -gt 0) {
                            Write-Host ("Cycle {0}: AD inventory moved {1} already-Windows11 computer(s) from Computers.txt to {2}. Evidence={3}" -f $CycleNumber,$preMoveResult.Moved,$preMoveResult.AlreadyWindows11Path,$adAlreadyPath) -ForegroundColor Green
                        }
                        else {
                            Write-Host ("Cycle {0}: AD inventory detected already-Windows11 computer(s), but none were still present in Computers.txt. Evidence={1}" -f $CycleNumber,$adAlreadyPath) -ForegroundColor DarkGray
                        }
                    }
                }
                else {
                    Write-Host ("Cycle {0}: AD inventory precheck found no Windows 11 computer in current Computers.txt. CSV={1}" -f $CycleNumber,$AdInventoryCsv) -ForegroundColor DarkGray
                }
            }
            else {
                Write-Host ("Cycle {0}: AD inventory precheck skipped: no AD rows loaded. CSV={1}" -f $CycleNumber,$AdInventoryCsv) -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host ("WARN: Cycle {0}: AD inventory precheck failed: {1}" -f $CycleNumber,$_.Exception.Message) -ForegroundColor Yellow
        }
    }

    $currentComputers = @(Get-ComputerList -Path $ComputerListPath)
    return [pscustomobject]@{
        RemainingComputers = $currentComputers.Count
        MovedFromIntune = $movedFromIntune
        MovedFromAd = $movedFromAd
    }
}

function Add-IntuneInventoryFieldsToResult {
    param(
        [Parameter(Mandatory = $true)]$Result,
        [AllowNull()][hashtable]$IntuneInventoryMap,
        [AllowNull()][string]$IntuneInventoryCsv
    )

    if ($null -eq $Result) { return $Result }
    if (-not $Result.PSObject.Properties['ComputerName']) { return $Result }

    $computerKey = Get-ComputerListKey -ComputerName ([string]$Result.ComputerName)
    $hasInventory = ($null -ne $IntuneInventoryMap -and $IntuneInventoryMap.Count -gt 0)
    $hasIntuneRow = ($hasInventory -and -not [string]::IsNullOrWhiteSpace($computerKey) -and $IntuneInventoryMap.ContainsKey($computerKey))

    if ($hasIntuneRow) {
        $intuneRow = $IntuneInventoryMap[$computerKey]
        $fields = [ordered]@{
            IntuneInventoryPresent = $true
            IntuneDeviceName = if ($intuneRow.PSObject.Properties['DeviceName']) { [string]$intuneRow.DeviceName } else { '' }
            IntuneManagedDeviceName = if ($intuneRow.PSObject.Properties['ManagedDeviceName']) { [string]$intuneRow.ManagedDeviceName } else { '' }
            IntuneManagedDeviceId = if ($intuneRow.PSObject.Properties['IntuneManagedDeviceId']) { [string]$intuneRow.IntuneManagedDeviceId } else { '' }
            IntuneAzureADDeviceId = if ($intuneRow.PSObject.Properties['AzureADDeviceId']) { [string]$intuneRow.AzureADDeviceId } else { '' }
            IntuneOperatingSystem = if ($intuneRow.PSObject.Properties['OperatingSystem']) { [string]$intuneRow.OperatingSystem } else { '' }
            IntuneOSVersion = if ($intuneRow.PSObject.Properties['OSVersion']) { [string]$intuneRow.OSVersion } else { '' }
            IntuneLastSyncDateTime = if ($intuneRow.PSObject.Properties['LastSyncDateTime']) { [string]$intuneRow.LastSyncDateTime } else { '' }
            IntuneUserPrincipalName = if ($intuneRow.PSObject.Properties['UserPrincipalName']) { [string]$intuneRow.UserPrincipalName } else { '' }
            IntuneComplianceState = if ($intuneRow.PSObject.Properties['ComplianceState']) { [string]$intuneRow.ComplianceState } else { '' }
            IntuneManagementState = if ($intuneRow.PSObject.Properties['ManagementState']) { [string]$intuneRow.ManagementState } else { '' }
            IntuneInventoryCsv = $IntuneInventoryCsv
        }
    }
    elseif ($hasInventory) {
        $fields = [ordered]@{
            IntuneInventoryPresent = $false
            IntuneDeviceName = ''
            IntuneManagedDeviceName = ''
            IntuneManagedDeviceId = ''
            IntuneAzureADDeviceId = ''
            IntuneOperatingSystem = ''
            IntuneOSVersion = ''
            IntuneLastSyncDateTime = ''
            IntuneUserPrincipalName = ''
            IntuneComplianceState = ''
            IntuneManagementState = ''
            IntuneInventoryCsv = $IntuneInventoryCsv
        }
    }
    else {
        return $Result
    }

    foreach ($fieldName in $fields.Keys) {
        $Result | Add-Member -NotePropertyName $fieldName -NotePropertyValue $fields[$fieldName] -Force
    }

    return $Result
}

function Test-SingleComputerLaunch {
    param([Parameter(Mandatory = $true)][string]$Path)

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    if ($fullPath -match '\\SingleComputer(Runs)?\\') { return $true }

    $computers = @(Get-ComputerList -Path $fullPath)
    return ($computers.Count -eq 1)
}

$script:IsSingleComputerLaunch = Test-SingleComputerLaunch -Path $ComputerListPath
if ($script:IsSingleComputerLaunch) {
    $ThrottleLimit = 1
    $GlobalConcurrencyLimit = 0
}

function Write-LauncherLogLine {
    param([AllowNull()][string]$Line)

    foreach ($target in @($script:LauncherLogPath, $script:LauncherLatestLogPath)) {
        if ([string]::IsNullOrWhiteSpace($target)) { continue }
        try { Add-Content -LiteralPath $target -Value ([string]$Line) -Encoding UTF8 -ErrorAction Stop }
        catch { }
    }
}

function Write-Host {
    param(
        [Parameter(Position = 0)][object]$Object = '',
        [switch]$NoNewline,
        [System.ConsoleColor]$ForegroundColor,
        [System.ConsoleColor]$BackgroundColor
    )
    $ts = if ($null -ne $Object -and '' -ne [string]$Object) { "[{0}] " -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') } else { '' }
    $line = $ts + [string]$Object
    $p = @{ Object = $line }
    if ($PSBoundParameters.ContainsKey('ForegroundColor')) { $p['ForegroundColor'] = $ForegroundColor }
    if ($PSBoundParameters.ContainsKey('BackgroundColor')) { $p['BackgroundColor'] = $BackgroundColor }
    if ($PSBoundParameters.ContainsKey('NoNewline')) { $p['NoNewline'] = $NoNewline }
    Microsoft.PowerShell.Utility\Write-Host @p
    Write-LauncherLogLine -Line $line
}

function Test-SetupSourceMapSyntax {
    param([AllowNull()][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return }

    $expanded = [Environment]::ExpandEnvironmentVariables($Path.Trim('"'))
    if (-not (Test-Path -LiteralPath $expanded -PathType Leaf)) {
        Write-Host ("Setup source map preflight skipped because the operator cannot read it: {0}" -f $expanded) -ForegroundColor DarkYellow
        return
    }

    $rows = @(Import-Csv -LiteralPath $expanded -ErrorAction Stop)
    if ($rows.Count -eq 0) {
        throw "Setup source map is empty: $expanded"
    }

    $validRows = 0
    $rowNumber = 1
    foreach ($row in $rows) {
        $rowNumber++
        $scopeType = if ($row.PSObject.Properties['ScopeType']) { [string]$row.ScopeType } elseif ($row.PSObject.Properties['Type']) { [string]$row.Type } else { '' }
        $scopeValue = if ($row.PSObject.Properties['ScopeValue']) { [string]$row.ScopeValue } elseif ($row.PSObject.Properties['Value']) { [string]$row.Value } else { '' }
        $sourcePath = if ($row.PSObject.Properties['SetupSourcePath']) { [string]$row.SetupSourcePath } elseif ($row.PSObject.Properties['SourcePath']) { [string]$row.SourcePath } else { '' }

        if ([string]::IsNullOrWhiteSpace($scopeType) -or [string]::IsNullOrWhiteSpace($sourcePath)) {
            throw "Invalid setup source map row $rowNumber. ScopeType and SetupSourcePath are required."
        }

        if ($scopeType -notmatch '^(?i:Default|Fallback|Subnet|CIDR|IPPrefix|Prefix|ComputerName|Hostname|ComputerPrefix|HostnamePrefix)$') {
            throw "Invalid setup source map row $rowNumber. Unsupported ScopeType '$scopeType'."
        }

        if ($scopeType -match '^(?i:Subnet|CIDR)$' -and $scopeValue -notmatch '^\d{1,3}(\.\d{1,3}){3}/([0-9]|[12][0-9]|3[0-2])$') {
            throw "Invalid setup source map row $rowNumber. ScopeValue must be CIDR for ScopeType '$scopeType'."
        }

        if ($sourcePath -notmatch '^\\\\[^\\]+\\[^\\]+') {
            throw "Invalid setup source map row $rowNumber. SetupSourcePath must be UNC: $sourcePath"
        }

        $validRows++
    }

    Write-Host ("Setup source map preflight OK: {0}; Rows={1}" -f $expanded,$validRows) -ForegroundColor Green
}

function Resolve-PsExecPath {
    param([string]$Path)

    $candidate = [string]$Path
    if (-not [string]::IsNullOrWhiteSpace($candidate)) {
        $embedded = @([regex]::Matches($candidate, '(?i)(?:[A-Z]:\\|\\\\)[^"<>|?*]+?PsExec\.exe') | ForEach-Object { $_.Value.Trim() })
        if ($embedded.Count -gt 1) {
            throw ("Invalid PsExecPath value: {0}. Multiple PsExec paths were provided: {1}." -f $Path,($embedded -join ' | '))
        }
        if ($embedded.Count -eq 1) { $candidate = $embedded[0] }
        $candidate = $candidate.Trim('"')
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return (Get-Item -LiteralPath $candidate).FullName
        }
        throw ("Invalid PsExecPath value: {0}. Provide exactly one executable path." -f $Path)
    }

    $local = Join-Path $script:BaseDir 'PsExec.exe'
    if (Test-Path -LiteralPath $local -PathType Leaf) { return (Get-Item -LiteralPath $local).FullName }
    $command = Get-Command -Name 'PsExec.exe' -CommandType Application -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    throw ("PsExec.exe not found. Place it in {0} or add PsExec.exe to PATH." -f $script:BaseDir)
}

function Get-PsExecSecurityEvidence {
    param([AllowNull()][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return [pscustomobject]@{
            CheckStatus = 'SkippedDryRun'
            Path = ''
            FileName = ''
            LengthBytes = ''
            SHA256 = ''
            SignatureStatus = ''
            SignerSubject = ''
            SignerIssuer = ''
            ProductName = ''
            FileDescription = ''
            FileVersion = ''
            ProductVersion = ''
            IsMicrosoftSigned = $true
            IsExpectedFileName = $true
            IsCompliant = $true
            Detail = 'DryRun=True; PsExec is not required.'
        }
    }

    $item = Get-Item -LiteralPath $Path -ErrorAction Stop
    $hash = Get-FileHash -LiteralPath $item.FullName -Algorithm SHA256 -ErrorAction Stop
    $signature = Get-AuthenticodeSignature -LiteralPath $item.FullName -ErrorAction Stop
    $versionInfo = $item.VersionInfo
    $signerSubject = if ($signature.SignerCertificate) { [string]$signature.SignerCertificate.Subject } else { '' }
    $signerIssuer = if ($signature.SignerCertificate) { [string]$signature.SignerCertificate.Issuer } else { '' }
    $isMicrosoftSigned = ([string]$signature.Status -eq 'Valid' -and ($signerSubject -match '(^|,\s*)(CN|O)=Microsoft Corporation(,|$)'))
    $isExpectedFileName = ($item.Name -match '^(?i:PsExec|PsExec64)\.exe$')
    $isCompliant = ($isMicrosoftSigned -and $isExpectedFileName)
    $detail = if ($isCompliant) { 'Microsoft-signed PsExec binary accepted.' } else { 'Blocked: PsExec must be named PsExec.exe/PsExec64.exe and have a valid Microsoft Corporation Authenticode signature.' }

    [pscustomobject]@{
        CheckStatus = if ($isCompliant) { 'Compliant' } else { 'Blocked' }
        Path = [string]$item.FullName
        FileName = [string]$item.Name
        LengthBytes = [string]$item.Length
        SHA256 = [string]$hash.Hash
        SignatureStatus = [string]$signature.Status
        SignerSubject = $signerSubject
        SignerIssuer = $signerIssuer
        ProductName = [string]$versionInfo.ProductName
        FileDescription = [string]$versionInfo.FileDescription
        FileVersion = [string]$versionInfo.FileVersion
        ProductVersion = [string]$versionInfo.ProductVersion
        IsMicrosoftSigned = [bool]$isMicrosoftSigned
        IsExpectedFileName = [bool]$isExpectedFileName
        IsCompliant = [bool]$isCompliant
        Detail = $detail
    }
}

function Assert-PsExecSecurityEvidence {
    param([Parameter(Mandatory = $true)]$Evidence)

    if (-not [bool]$Evidence.IsCompliant) {
        throw ("PsExec security validation failed. Status={0}; Path={1}; SignatureStatus={2}; Signer={3}; SHA256={4}; Detail={5}" -f $Evidence.CheckStatus,$Evidence.Path,$Evidence.SignatureStatus,$Evidence.SignerSubject,$Evidence.SHA256,$Evidence.Detail)
    }
}

function ConvertTo-PsExecSecurityEvidenceRows {
    param([AllowNull()]$Evidence)

    if ($null -eq $Evidence) { return @() }
    $fields = @('CheckStatus','Path','FileName','LengthBytes','SHA256','SignatureStatus','SignerSubject','SignerIssuer','ProductName','FileDescription','FileVersion','ProductVersion','IsMicrosoftSigned','IsExpectedFileName','IsCompliant','Detail')
    return @($fields | ForEach-Object {
        [pscustomobject]@{ Field = $_; Value = if ($Evidence.PSObject.Properties[$_]) { [string]$Evidence.$_ } else { '' } }
    })
}

if (-not (Test-Path -LiteralPath $LocalScriptPath -PathType Leaf)) {
    throw "Local repair script not found: $LocalScriptPath"
}
if (-not (Test-Path -LiteralPath $LocalWorkerPath -PathType Leaf)) {
    throw "Local PsExec worker script not found: $LocalWorkerPath"
}

New-Directory -Path $LogRoot
New-Directory -Path $ReportRoot
New-Directory -Path $CentralLogRoot
New-Directory -Path $script:LauncherLogRoot
$script:CancellationRunStartedUtc = (Get-Date).ToUniversalTime()
$script:CancellationStateRoot = Join-Path $runDataRoot 'State'
New-Directory -Path $script:CancellationStateRoot
$script:CancellationSignalPath = Join-Path $script:CancellationStateRoot ("StopRequested_{0}.json" -f $PID)
$script:ActiveLotRunStatePath = Join-Path $script:CancellationStateRoot ("ActiveLotRun_{0}.json" -f $PID)
Remove-Item -LiteralPath $script:CancellationSignalPath -Force -ErrorAction SilentlyContinue
Initialize-LotCancellationSupport
Set-ActiveLotRunState -Status 'Starting'
Write-Host ("Controlled stop: first Ctrl+C stops new starts and drains active jobs for up to {0} minute(s); second Ctrl+C forces local job stop." -f $CancellationDrainTimeoutMinutes) -ForegroundColor DarkCyan

$launcherLogHeader = "[{0}] ===== SmartM365 Windows 11 Upgrade Toolkit launcher v{1} started. Lot={2}; ComputerList={3}; Technician={4}; TechComputer={5} =====" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$script:LauncherVersion,(Split-Path -Leaf $script:LotRoot),$ComputerListPath,$script:TechnicianIdentity.Account,$script:TechnicianIdentity.ComputerName
Set-Content -LiteralPath $script:LauncherLogPath -Value $launcherLogHeader -Encoding UTF8 -Force
Set-Content -LiteralPath $script:LauncherLatestLogPath -Value $launcherLogHeader -Encoding UTF8 -Force
$powerShellProcessPath = try { [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName } catch { '' }
$powerShellRuntimeText = "Edition={0}; Version={1}; Process={2}" -f $PSVersionTable.PSEdition,$PSVersionTable.PSVersion,$powerShellProcessPath
Add-Content -LiteralPath $script:LauncherLogPath -Value ("[{0}] PowerShell: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$powerShellRuntimeText) -Encoding UTF8
Add-Content -LiteralPath $script:LauncherLatestLogPath -Value ("[{0}] PowerShell: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$powerShellRuntimeText) -Encoding UTF8

$script:TechnicianRunGuardHistoryPath = Get-TechnicianRunGuardHistoryPath
$script:UseEffectiveTechnicianRunGuardHistory = ($UseTechnicianRunGuardHistory -and -not $IgnoreTechnicianRunGuardHistory -and -not $IgnoreRunGuard -and -not $DryRun -and $RunGuardHours -gt 0)
if ($script:UseEffectiveTechnicianRunGuardHistory) {
    try {
        New-Directory -Path (Split-Path -Parent $script:TechnicianRunGuardHistoryPath)
        Invoke-TechnicianRunGuardHistoryLock -ArgumentList @($script:TechnicianRunGuardHistoryPath, $RunGuardHours) -ScriptBlock {
            param($LockedPath, $LockedRunGuardHours)
            $history = Read-TechnicianRunGuardHistory -Path $LockedPath -RunGuardHours $LockedRunGuardHours
            Save-TechnicianRunGuardHistory -Path $LockedPath -History $history
        }
    }
    catch {
        Write-Host ("Technician run guard history is temporarily unavailable; launches will fail closed until access recovers. Error={0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
}
Test-SetupSourceMapSyntax -Path $SetupSourceMapPath

$resolvedPsExec = if ($DryRun) { '' } else { Resolve-PsExecPath -Path $PsExecPath }
$script:PsExecSecurityEvidence = Get-PsExecSecurityEvidence -Path $resolvedPsExec
if (-not $DryRun) { Assert-PsExecSecurityEvidence -Evidence $script:PsExecSecurityEvidence }
$script:PsExecSecurityEvidenceRows = @(ConvertTo-PsExecSecurityEvidenceRows -Evidence $script:PsExecSecurityEvidence)

$remoteArgs = New-Object System.Collections.ArrayList
if ($AuditOnly) { [void]$remoteArgs.Add('-AuditOnly') }
if ($IgnoreRunGuard) { [void]$remoteArgs.Add('-IgnoreRunGuard') }
[void]$remoteArgs.Add('-RunGuardHours'); [void]$remoteArgs.Add([string]$RunGuardHours)
if ($AllowPolicyRepair) { [void]$remoteArgs.Add('-AllowPolicyRepair') }
if ($AllowWUReset) { [void]$remoteArgs.Add('-AllowWUReset') }
if ($AllowForceUpgrade) { [void]$remoteArgs.Add('-AllowForceUpgrade') }
if ($AllowSetupUpgrade) { [void]$remoteArgs.Add('-AllowSetupUpgrade') }
if ($DirectSetupUpgrade) { [void]$remoteArgs.Add('-DirectSetupUpgrade') }
if ($AllowReboot) { [void]$remoteArgs.Add('-AllowReboot') }
if ($AllowSetupCompletionRebootWhenNoUser) { [void]$remoteArgs.Add('-AllowSetupCompletionRebootWhenNoUser') }
if ($AllowSetupProfileRepair) { [void]$remoteArgs.Add('-AllowSetupProfileRepair') }
if ($ScheduleRetryAfterReboot) { [void]$remoteArgs.Add('-ScheduleRetryAfterReboot') }
[void]$remoteArgs.Add('-RetryAfterRebootMaxAttempts'); [void]$remoteArgs.Add([string]$RetryAfterRebootMaxAttempts)
[void]$remoteArgs.Add('-RetryAfterRebootDelaySeconds'); [void]$remoteArgs.Add([string]$RetryAfterRebootDelaySeconds)
[void]$remoteArgs.Add('-ForceRequiredRebootWhenUptimeOverDays'); [void]$remoteArgs.Add([string]$ForceRequiredRebootWhenUptimeOverDays)
if ($SkipVirtualMachines) { [void]$remoteArgs.Add('-SkipVirtualMachines') }
if ($AllowDiskCleanup) { [void]$remoteArgs.Add('-AllowDiskCleanup') }
if ($AllowAdvancedDiskCleanup -or $AllowDismComponentCleanup) { [void]$remoteArgs.Add('-AllowAdvancedDiskCleanup') }
if ($SkipSetupMediaPreCopy) { [void]$remoteArgs.Add('-SkipSetupMediaPreCopy') }
[void]$remoteArgs.Add('-SetupExecutionMode'); [void]$remoteArgs.Add($SetupExecutionMode)
[void]$remoteArgs.Add('-SetupMediaId'); [void]$remoteArgs.Add($SetupMediaId)
[void]$remoteArgs.Add('-SetupLanguage'); [void]$remoteArgs.Add($SetupLanguage)
[void]$remoteArgs.Add('-SetupDynamicUpdate'); [void]$remoteArgs.Add($SetupDynamicUpdate)
[void]$remoteArgs.Add('-SetupCacheRoot'); [void]$remoteArgs.Add($script:RemoteSetupCacheRoot)
[void]$remoteArgs.Add('-SetupSourceCandidateLimit'); [void]$remoteArgs.Add([string]$SetupSourceCandidateLimit)
[void]$remoteArgs.Add('-SetupMediaCopyIpGapMilliseconds'); [void]$remoteArgs.Add([string]$SetupMediaCopyIpGapMilliseconds)
[void]$remoteArgs.Add('-SetupMediaCopyJitterSeconds'); [void]$remoteArgs.Add([string]$SetupMediaCopyJitterSeconds)
if ($SetupSourceConcurrencyLimit -gt 0) {
    [void]$remoteArgs.Add('-SetupSourceConcurrencyLimit'); [void]$remoteArgs.Add([string]$SetupSourceConcurrencyLimit)
    [void]$remoteArgs.Add('-SetupSourceConcurrencyLeaseMinutes'); [void]$remoteArgs.Add([string]$SetupSourceConcurrencyLeaseMinutes)
}
if (-not [string]::IsNullOrWhiteSpace($SetupSourceConcurrencyGateRoot)) {
    [void]$remoteArgs.Add('-SetupSourceConcurrencyGateRoot'); [void]$remoteArgs.Add($SetupSourceConcurrencyGateRoot)
}
if ($SetupSubnetConcurrencyLimit -gt 0) {
    [void]$remoteArgs.Add('-SetupSubnetConcurrencyLimit'); [void]$remoteArgs.Add([string]$SetupSubnetConcurrencyLimit)
    [void]$remoteArgs.Add('-SetupSubnetPrefixLength'); [void]$remoteArgs.Add([string]$SetupSubnetPrefixLength)
    [void]$remoteArgs.Add('-SetupSubnetConcurrencyLeaseMinutes'); [void]$remoteArgs.Add([string]$SetupSubnetConcurrencyLeaseMinutes)
}
if (-not [string]::IsNullOrWhiteSpace($SetupSubnetConcurrencyGateRoot)) {
    [void]$remoteArgs.Add('-SetupSubnetConcurrencyGateRoot'); [void]$remoteArgs.Add($SetupSubnetConcurrencyGateRoot)
}
if (-not [string]::IsNullOrWhiteSpace($SetupSourcePath)) {
    [void]$remoteArgs.Add('-SetupSourcePath'); [void]$remoteArgs.Add($SetupSourcePath)
}
if (-not [string]::IsNullOrWhiteSpace($SetupSourceMapPath)) {
    [void]$remoteArgs.Add('-SetupSourceMapPath'); [void]$remoteArgs.Add($SetupSourceMapPath)
}

$script:LauncherOptionRows = @(
    [pscustomobject]@{ Category = 'Operator'; Option = 'TechnicianAccount'; Value = [string]$script:TechnicianIdentity.Account }
    [pscustomobject]@{ Category = 'Operator'; Option = 'TechnicianUPN'; Value = [string]$script:TechnicianIdentity.UserPrincipalName }
    [pscustomobject]@{ Category = 'Operator'; Option = 'TechnicianSID'; Value = [string]$script:TechnicianIdentity.Sid }
    [pscustomobject]@{ Category = 'Operator'; Option = 'TechnicianAuthType'; Value = [string]$script:TechnicianIdentity.AuthenticationType }
    [pscustomobject]@{ Category = 'Operator'; Option = 'TechnicianComputer'; Value = [string]$script:TechnicianIdentity.ComputerName }
    [pscustomobject]@{ Category = 'Security'; Option = 'PsExecSecurityStatus'; Value = [string]$script:PsExecSecurityEvidence.CheckStatus }
    [pscustomobject]@{ Category = 'Security'; Option = 'PsExecSHA256'; Value = [string]$script:PsExecSecurityEvidence.SHA256 }
    [pscustomobject]@{ Category = 'Security'; Option = 'PsExecSignatureStatus'; Value = [string]$script:PsExecSecurityEvidence.SignatureStatus }
    [pscustomobject]@{ Category = 'Security'; Option = 'PsExecSignerSubject'; Value = [string]$script:PsExecSecurityEvidence.SignerSubject }
    [pscustomobject]@{ Category = 'Mode'; Option = 'DryRun'; Value = [string][bool]$DryRun }
    [pscustomobject]@{ Category = 'Mode'; Option = 'AuditOnly'; Value = [string][bool]$AuditOnly }
    [pscustomobject]@{ Category = 'Mode'; Option = 'RunOnce'; Value = [string][bool]$RunOnce }
    [pscustomobject]@{ Category = 'Mode'; Option = 'IgnoreRunGuard'; Value = [string][bool]$IgnoreRunGuard }
    [pscustomobject]@{ Category = 'Mode'; Option = 'RunGuardHours'; Value = [string]$RunGuardHours }
    [pscustomobject]@{ Category = 'Mode'; Option = 'UseTechnicianRunGuardHistory'; Value = [string][bool]$UseTechnicianRunGuardHistory }
    [pscustomobject]@{ Category = 'Mode'; Option = 'IgnoreTechnicianRunGuardHistory'; Value = [string][bool]$IgnoreTechnicianRunGuardHistory }
    [pscustomobject]@{ Category = 'Mode'; Option = 'EffectiveTechnicianRunGuardHistory'; Value = [string][bool]$script:UseEffectiveTechnicianRunGuardHistory }
    [pscustomobject]@{ Category = 'Mode'; Option = 'TechnicianRunGuardHistoryPath'; Value = [string]$script:TechnicianRunGuardHistoryPath }
    [pscustomobject]@{ Category = 'Actions'; Option = 'AllowPolicyRepair'; Value = [string][bool]$AllowPolicyRepair }
    [pscustomobject]@{ Category = 'Actions'; Option = 'AllowWUReset'; Value = [string][bool]$AllowWUReset }
    [pscustomobject]@{ Category = 'Actions'; Option = 'AllowForceUpgrade'; Value = [string][bool]$AllowForceUpgrade }
    [pscustomobject]@{ Category = 'Actions'; Option = 'AllowSetupUpgrade'; Value = [string][bool]$AllowSetupUpgrade }
    [pscustomobject]@{ Category = 'Actions'; Option = 'DirectSetupUpgrade'; Value = [string][bool]$DirectSetupUpgrade }
    [pscustomobject]@{ Category = 'Actions'; Option = 'AllowReboot'; Value = [string][bool]$AllowReboot }
    [pscustomobject]@{ Category = 'Actions'; Option = 'AllowSetupCompletionRebootWhenNoUser'; Value = [string][bool]$AllowSetupCompletionRebootWhenNoUser }
    [pscustomobject]@{ Category = 'Actions'; Option = 'AllowSetupProfileRepair'; Value = [string][bool]$AllowSetupProfileRepair }
    [pscustomobject]@{ Category = 'Actions'; Option = 'ScheduleRetryAfterReboot'; Value = [string][bool]$ScheduleRetryAfterReboot }
    [pscustomobject]@{ Category = 'Actions'; Option = 'RetryAfterRebootMaxAttempts'; Value = [string]$RetryAfterRebootMaxAttempts }
    [pscustomobject]@{ Category = 'Actions'; Option = 'RetryAfterRebootDelaySeconds'; Value = [string]$RetryAfterRebootDelaySeconds }
    [pscustomobject]@{ Category = 'Actions'; Option = 'ForceRequiredRebootWhenUptimeOverDays'; Value = [string]$ForceRequiredRebootWhenUptimeOverDays }
    [pscustomobject]@{ Category = 'Actions'; Option = 'SkipVirtualMachines'; Value = [string][bool]$SkipVirtualMachines }
    [pscustomobject]@{ Category = 'Actions'; Option = 'AllowDiskCleanup'; Value = [string][bool]$AllowDiskCleanup }
    [pscustomobject]@{ Category = 'Actions'; Option = 'AllowAdvancedDiskCleanup'; Value = [string][bool]($AllowAdvancedDiskCleanup -or $AllowDismComponentCleanup) }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSourcePath'; Value = [string]$SetupSourcePath }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSourceMapPath'; Value = [string]$SetupSourceMapPath }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupExecutionMode'; Value = [string]$SetupExecutionMode }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupMediaId'; Value = [string]$SetupMediaId }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupLanguage'; Value = [string]$SetupLanguage }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupDynamicUpdate'; Value = [string]$SetupDynamicUpdate }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SkipSetupMediaPreCopy'; Value = [string][bool]$SkipSetupMediaPreCopy }
    [pscustomobject]@{ Category = 'Setup'; Option = 'TargetSetupCacheRoot'; Value = [string]$script:RemoteSetupCacheRoot }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSourceCandidateLimit'; Value = [string]$SetupSourceCandidateLimit }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupMediaCopyIpGapMilliseconds'; Value = [string]$SetupMediaCopyIpGapMilliseconds }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupMediaCopyJitterSeconds'; Value = [string]$SetupMediaCopyJitterSeconds }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSourceConcurrencyLimit'; Value = [string]$SetupSourceConcurrencyLimit }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSourceConcurrencyLeaseMinutes'; Value = [string]$SetupSourceConcurrencyLeaseMinutes }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSourceConcurrencyGateRoot'; Value = [string]$SetupSourceConcurrencyGateRoot }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSubnetConcurrencyLimit'; Value = [string]$SetupSubnetConcurrencyLimit }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSubnetPrefixLength'; Value = [string]$SetupSubnetPrefixLength }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSubnetConcurrencyLeaseMinutes'; Value = [string]$SetupSubnetConcurrencyLeaseMinutes }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSubnetConcurrencyGateRoot'; Value = [string]$SetupSubnetConcurrencyGateRoot }
    [pscustomobject]@{ Category = 'AD'; Option = 'AdInventoryCsv'; Value = [string]$AdInventoryCsv }
    [pscustomobject]@{ Category = 'AD'; Option = 'AdRootInventoryCsv'; Value = [string]$AdRootInventoryCsv }
    [pscustomobject]@{ Category = 'AD'; Option = 'AdDomain'; Value = [string]$AdDomain }
    [pscustomobject]@{ Category = 'AD'; Option = 'SkipAdInventoryRefresh'; Value = [string][bool]$SkipAdInventoryRefresh }
    [pscustomobject]@{ Category = 'AD'; Option = 'EffectiveSkipAdInventoryRefresh'; Value = [string][bool]$EffectiveSkipAdInventoryRefresh }
    [pscustomobject]@{ Category = 'AD'; Option = 'UsesRecentRootInventory'; Value = [string][bool]$AdInventoryUsesRecentRootCsv }
    [pscustomobject]@{ Category = 'Intune'; Option = 'IntuneInventoryCsv'; Value = [string]$IntuneInventoryCsv }
    [pscustomobject]@{ Category = 'Intune'; Option = 'IntuneRootInventoryCsv'; Value = [string]$IntuneRootInventoryCsv }
    [pscustomobject]@{ Category = 'Intune'; Option = 'IntuneInventoryNameColumn'; Value = [string]$IntuneInventoryNameColumn }
    [pscustomobject]@{ Category = 'Intune'; Option = 'IntuneInventoryPageSize'; Value = [string]$IntuneInventoryPageSize }
    [pscustomobject]@{ Category = 'Intune'; Option = 'IntuneTenantId'; Value = [string]$IntuneTenantId }
    [pscustomobject]@{ Category = 'Intune'; Option = 'SkipIntuneInventoryRefresh'; Value = [string][bool]$SkipIntuneInventoryRefresh }
    [pscustomobject]@{ Category = 'Intune'; Option = 'EffectiveSkipIntuneInventoryRefresh'; Value = [string][bool]$EffectiveSkipIntuneInventoryRefresh }
    [pscustomobject]@{ Category = 'Intune'; Option = 'UsesRecentRootInventory'; Value = [string][bool]$IntuneInventoryUsesRecentRootCsv }
    [pscustomobject]@{ Category = 'Parallelism'; Option = 'ThrottleLimit'; Value = [string]$ThrottleLimit }
    [pscustomobject]@{ Category = 'Parallelism'; Option = 'GlobalConcurrencyLimit'; Value = [string]$GlobalConcurrencyLimit }
    [pscustomobject]@{ Category = 'Parallelism'; Option = 'GlobalConcurrencyLeaseTimeoutMinutes'; Value = [string]$GlobalConcurrencyLeaseTimeoutMinutes }
    [pscustomobject]@{ Category = 'Parallelism'; Option = 'DelayBetweenComputersSeconds'; Value = [string]$DelayBetweenComputersSeconds }
    [pscustomobject]@{ Category = 'Parallelism'; Option = 'DelayBetweenCyclesMinutes'; Value = [string]$DelayBetweenCyclesMinutes }
    [pscustomobject]@{ Category = 'Parallelism'; Option = 'MaxCycles'; Value = [string]$MaxCycles }
    [pscustomobject]@{ Category = 'Parallelism'; Option = 'PsExecTimeoutMinutes'; Value = [string]$PsExecTimeoutMinutes }
    [pscustomobject]@{ Category = 'Paths'; Option = 'ComputerListPath'; Value = [string]$ComputerListPath }
    [pscustomobject]@{ Category = 'Paths'; Option = 'LogRoot'; Value = [string]$LogRoot }
    [pscustomobject]@{ Category = 'Paths'; Option = 'ReportRoot'; Value = [string]$ReportRoot }
    [pscustomobject]@{ Category = 'Paths'; Option = 'CentralLogRoot'; Value = [string]$CentralLogRoot }
    [pscustomobject]@{ Category = 'Paths'; Option = 'LauncherLogRoot'; Value = [string]$script:LauncherLogRoot }
    [pscustomobject]@{ Category = 'Paths'; Option = 'LauncherLogPath'; Value = [string]$script:LauncherLogPath }
    [pscustomobject]@{ Category = 'Paths'; Option = 'LauncherLatestLogPath'; Value = [string]$script:LauncherLatestLogPath }
    [pscustomobject]@{ Category = 'Paths'; Option = 'NoCentralLogCollection'; Value = [string][bool]$NoCentralLogCollection }
    [pscustomobject]@{ Category = 'Paths'; Option = 'KeepCentralLogHistory'; Value = [string][bool]$KeepCentralLogHistory }
    [pscustomobject]@{ Category = 'Paths'; Option = 'CentralLogCollectionMode'; Value = [string]$CentralLogCollectionMode }
)

Write-Host "SmartM365 Windows 11 Upgrade Toolkit launcher v$script:LauncherVersion"
Write-Host "PowerShell    : $powerShellRuntimeText"
Write-Host "LOT root      : $script:LotRoot"
Write-Host "Launcher log  : $script:LauncherLogPath"
Write-Host "Computer list : $ComputerListPath"
Write-Host "PsExec        : $resolvedPsExec"
Write-Host ("PsExec sec.   : Status={0}; SHA256={1}; Signature={2}; Signer={3}; Version={4}" -f $script:PsExecSecurityEvidence.CheckStatus,$script:PsExecSecurityEvidence.SHA256,$script:PsExecSecurityEvidence.SignatureStatus,$script:PsExecSecurityEvidence.SignerSubject,$script:PsExecSecurityEvidence.FileVersion)
Write-Host "Repair script : $LocalScriptPath"
Write-Host "Worker script : $LocalWorkerPath"
Write-Host ("Technician     : Account={0}; UPN={1}; SID={2}; Auth={3}; Computer={4}" -f $script:TechnicianIdentity.Account,$script:TechnicianIdentity.UserPrincipalName,$script:TechnicianIdentity.Sid,$script:TechnicianIdentity.AuthenticationType,$script:TechnicianIdentity.ComputerName)
Write-Host "Mode          : DryRun=$DryRun; AuditOnly=$AuditOnly; RunOnce=$RunOnce; SkipVirtualMachines=$SkipVirtualMachines; DiskCleanup=$AllowDiskCleanup; AdvancedCleanup=$($AllowAdvancedDiskCleanup -or $AllowDismComponentCleanup); DirectSetup=$DirectSetupUpgrade; SetupCompletionRebootWhenNoUser=$AllowSetupCompletionRebootWhenNoUser; SetupProfileRepair=$AllowSetupProfileRepair; ForceRequiredRebootWhenUptimeOverDays=$ForceRequiredRebootWhenUptimeOverDays"
Write-Host "Setup         : AllowSetupUpgrade=$AllowSetupUpgrade; DirectSetup=$DirectSetupUpgrade; Effective=$([bool]($AllowSetupUpgrade -or $DirectSetupUpgrade)); Mode=$SetupExecutionMode; MediaId=$SetupMediaId; Language=$SetupLanguage; DynamicUpdate=$SetupDynamicUpdate; PreCopy=$(-not $SkipSetupMediaPreCopy)"
Write-Host "AD inventory  : Csv=$AdInventoryCsv; RootCsv=$AdRootInventoryCsv; Domain=$AdDomain; Refresh=$(-not $EffectiveSkipAdInventoryRefresh); RequestedSkip=$([bool]$SkipAdInventoryRefresh); RecentRoot=$AdInventoryUsesRecentRootCsv"
Write-Host "Intune invent.: Csv=$IntuneInventoryCsv; RootCsv=$IntuneRootInventoryCsv; Tenant=$IntuneTenantId; Refresh=$(-not $EffectiveSkipIntuneInventoryRefresh); RequestedSkip=$([bool]$SkipIntuneInventoryRefresh); RecentRoot=$IntuneInventoryUsesRecentRootCsv"
Write-Host "Tech run guard: Use=$script:UseEffectiveTechnicianRunGuardHistory; Requested=$UseTechnicianRunGuardHistory; Ignore=$IgnoreTechnicianRunGuardHistory; Hours=$RunGuardHours; Path=$script:TechnicianRunGuardHistoryPath"
Write-Host "Parallelism   : ThrottleLimit=$ThrottleLimit; GlobalConcurrencyLimit=$GlobalConcurrencyLimit; GlobalLeaseTimeout=$GlobalConcurrencyLeaseTimeoutMinutes minute(s)"
Write-Host "LOT/run options:"
foreach ($optionRow in @($script:LauncherOptionRows)) {
    Write-Host ("  [{0}] {1}={2}" -f $optionRow.Category,$optionRow.Option,$optionRow.Value)
}
if ($script:IsSingleComputerLaunch) { Write-Host "Single PC     : worker limits ignored for one-computer launch." }
Write-Host "Reports       : $ReportRoot"
Write-Host ""
try {
    $computerListDedupe = Remove-DuplicateComputerListEntries -Path $ComputerListPath
    if ($computerListDedupe.Changed) {
        Write-Host ("Computer list deduplicated before lot start: RemovedDuplicateLines={0}; DuplicateGroups={1}; Samples={2}; Backup={3}" -f $computerListDedupe.DuplicateLines,$computerListDedupe.DuplicateGroups,$computerListDedupe.DuplicateSamples,$computerListDedupe.BackupPath) -ForegroundColor Yellow
    }
}
catch {
    Write-Host ("WARN: failed to deduplicate Computers.txt before lot start: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
}


$reportColumns = @(
    'Timestamp',
    'ComputerName',
    'CycleNumber',
    'LauncherStatus',
    'RemoteStatus',
    'RemoteNextAction',
    'ExitCode',
    'Detail',
    'HardwareReadinessTag',
    'HardwareReadinessCode',
    'HardwareReadinessResult',
    'HardwareReadinessReason',
    'HardwareReadinessLog',
    'HardwareReadinessSource',
    'JobErrorMessage',
    'ADInventoryPresent',
    'ADDomain',
    'ADEnabled',
    'ADDNSHostName',
    'ADDistinguishedName',
    'ADOperatingSystem',
    'ADOperatingSystemVersion',
    'ADLastLogonTimestampUtc',
    'ADInventoryCsv',
    'IntuneInventoryPresent',
    'IntuneDeviceName',
    'IntuneManagedDeviceName',
    'IntuneManagedDeviceId',
    'IntuneAzureADDeviceId',
    'IntuneOperatingSystem',
    'IntuneOSVersion',
    'IntuneLastSyncDateTime',
    'IntuneUserPrincipalName',
    'IntuneComplianceState',
    'IntuneManagementState',
    'IntuneInventoryCsv',
    'SetupCacheAction',
    'SetupDynamicUpdate',
    'SelectedSetupSourcePath',
    'SetupSourceSelectionDetail',
    'DiskCleanupAction',
    'DiskCleanupFreedGB',
    'AdvancedDiskCleanupAction',
    'AdvancedDiskCleanupFreedGB',
    'DismCleanupAction',
    'DismCleanupFreedGB',
    'SetupCompletionRebootAction',
    'SetupCompletionRebootDetail',
    'SetupCompletionRebootUserCount',
    'SetupCompletionRebootUsers',
    'SetupProfileRepairAction',
    'SetupProfileRepairDetail',
    'SetupProfileRepairBlockingSid',
    'SetupProfileRepairKeptSid',
    'SetupProfileRepairProfilePath',
    'SetupProfileRepairBackupPath',
    'ControlledRebootAction',
    'ControlledRebootDetail',
    'ControlledRebootUserCount',
    'ControlledRebootUsers',
    'RetryAfterRebootAction',
    'RetryAfterRebootDetail',
    'RetryAfterRebootAttempt',
    'RetryAfterRebootMaxAttempts',
    'RetryAfterRebootTaskName',
    'UserRebootNotificationSent',
    'UserRebootNotificationLang',
    'UserRebootNotificationMessage',
    'RemoteLogsPath',
    'PsExecLogPath'
)

function ConvertTo-HtmlText {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    return [System.Security.SecurityElement]::Escape([string]$Value)
}

function ConvertTo-FileUri {
    param([AllowNull()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }

    $value = [string]$Path
    try {
        if ($value -match '^\\\\(?<Server>[^\\]+)\\(?<Share>[^\\]+)(?<Rest>.*)$') {
            $server = $Matches.Server
            $share = $Matches.Share
            $rest = ($Matches.Rest -replace '\\','/')
            return ('file://{0}/{1}{2}' -f $server,$share,$rest)
        }
        return ([System.Uri]::new([System.IO.Path]::GetFullPath($value))).AbsoluteUri
    }
    catch {
        return ''
    }
}

function New-HtmlLogLink {
    param(
        [AllowNull()][string]$Path,
        [string]$Label = 'Open'
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return '-' }
    $uri = ConvertTo-FileUri -Path $Path
    if ([string]::IsNullOrWhiteSpace($uri)) { return (ConvertTo-HtmlText $Path) }
    return ('<a href="{0}" title="{1}">{2}</a>' -f (ConvertTo-HtmlText $uri),(ConvertTo-HtmlText $Path),(ConvertTo-HtmlText $Label))
}

function Get-RemotePcLogsPath {
    param([AllowNull()][string]$ComputerName)
    if ([string]::IsNullOrWhiteSpace($ComputerName)) { return '' }

    $computer = ([string]$ComputerName).Trim()
    if ([string]::IsNullOrWhiteSpace($computer)) { return '' }
    return ('\\{0}\C$\ProgramData\SmartM365\Windows11UpgradeToolkit\Logs' -f $computer)
}

function ConvertTo-SimpleHtmlTable {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rows,
        [string[]]$Columns
    )

    if (-not $Rows -or $Rows.Count -eq 0) { return "<p class='empty'>No rows.</p>" }

    if (-not $Columns -or $Columns.Count -eq 0) {
        $columnSet = New-Object System.Collections.Generic.List[string]
        foreach ($row in $Rows) {
            foreach ($property in $row.PSObject.Properties) {
                if (-not $columnSet.Contains($property.Name)) { [void]$columnSet.Add($property.Name) }
            }
        }
        $Columns = @($columnSet)
    }

    $html = New-Object System.Collections.Generic.List[string]
    [void]$html.Add("<table>")
    [void]$html.Add("<tr>")
    foreach ($column in $Columns) { [void]$html.Add(("<th>{0}</th>" -f (ConvertTo-HtmlText $column))) }
    [void]$html.Add("</tr>")
    foreach ($row in $Rows) {
        [void]$html.Add("<tr>")
        foreach ($column in $Columns) {
            $value = ""
            $property = $row.PSObject.Properties[$column]
            if ($property) { $value = $property.Value }
            if ($column -eq 'Local log' -or $column -eq 'Collected logs' -or $column -eq 'Remote PC logs') {
                [void]$html.Add(("<td>{0}</td>" -f $value))
            }
            else {
                [void]$html.Add(("<td>{0}</td>" -f (ConvertTo-HtmlText $value)))
            }
        }
        [void]$html.Add("</tr>")
    }
    [void]$html.Add("</table>")
    return ($html -join "`r`n")
}

function Get-Windows11ReportRows {
    param([AllowEmptyCollection()][object[]]$Items)
    $normalized = foreach ($item in @($Items)) {
        $row = [ordered]@{}
        foreach ($column in $reportColumns) {
            $row[$column] = if ($null -ne $item -and $item.PSObject.Properties[$column]) { [string]$item.$column } else { '' }
        }
        [pscustomobject]$row
    }
    return @($normalized)
}

function ConvertTo-Windows11HtmlReportRow {
    param([AllowEmptyCollection()][object[]]$Items)

    $decorated = foreach ($item in @(Get-Windows11ReportRows -Items $Items)) {
        $row = [ordered]@{}
        foreach ($column in $reportColumns) { $row[$column] = [string]$item.$column }
        $row['Local log'] = New-HtmlLogLink -Path $row['PsExecLogPath']
        $row['Collected logs'] = New-HtmlLogLink -Path $row['RemoteLogsPath']
        $row['Remote PC logs'] = New-HtmlLogLink -Path (Get-RemotePcLogsPath -ComputerName $row['ComputerName'])
        [pscustomobject]$row
    }
    return @($decorated)
}


function Get-Windows11LatestRowsByComputer {
    param([AllowEmptyCollection()][object[]]$Rows)

    $latestByComputer = @{}
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $computer = if ($row.PSObject.Properties['ComputerName']) { [string]$row.ComputerName } else { '' }
        if ([string]::IsNullOrWhiteSpace($computer)) { continue }
        $key = Get-ComputerListKey -ComputerName $computer
        if ([string]::IsNullOrWhiteSpace($key)) { $key = $computer.ToUpperInvariant() }
        $latestByComputer[$key] = $row
    }

    return @($latestByComputer.GetEnumerator() | Sort-Object Name | ForEach-Object { $_.Value })
}


function Get-Windows11HtmlEffectiveStatus {
    param([AllowNull()][object]$Row)

    if ($null -eq $Row) { return '' }

    $remoteStatus = ''
    $launcherStatus = ''
    if ($Row.PSObject.Properties['RemoteStatus']) { $remoteStatus = [string]$Row.RemoteStatus }
    if ($Row.PSObject.Properties['LauncherStatus']) { $launcherStatus = [string]$Row.LauncherStatus }

    if (-not [string]::IsNullOrWhiteSpace($remoteStatus)) { return $remoteStatus }
    return $launcherStatus
}

function Get-Windows11AdminShareFailureType {
    param([AllowNull()][object]$Row)

    if ($null -eq $Row) { return 'UNKNOWN' }
    $detail = if ($Row.PSObject.Properties['Detail']) { [string]$Row.Detail } else { '' }
    $jobErrorMessage = if ($Row.PSObject.Properties['JobErrorMessage']) { [string]$Row.JobErrorMessage } else { '' }
    $evidence = @($detail, $jobErrorMessage) -join ' '
    if ($evidence -match 'FailureType=(?<FailureType>[A-Z0-9_]+)') { return $Matches.FailureType }
    if ($evidence -match 'DNS.*False|DNS_FAILED|host.*not.*found|nom.*introuvable') { return 'DNS_FAILED' }
    if ($evidence -match 'SMB_PORT_445_UNREACHABLE|Tcp445=False|port\s+445') { return 'SMB_PORT_445_UNREACHABLE' }
    if ($evidence -match 'AdminShare=False|Access is denied|Acc.s refus|administrative shares are not reachable') { return 'ADMIN_SHARE_UNREACHABLE' }
    return 'UNKNOWN'
}

function Add-Windows11AttemptCounter {
    param(
        [AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory = $true)][hashtable]$StatusCounter,
        [Parameter(Mandatory = $true)][hashtable]$NextActionCounter,
        [Parameter(Mandatory = $true)][hashtable]$AdminShareFailureCounter
    )

    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $status = Get-Windows11HtmlEffectiveStatus -Row $row
        if (-not [string]::IsNullOrWhiteSpace($status)) {
            if (-not $StatusCounter.ContainsKey($status)) { $StatusCounter[$status] = 0 }
            $StatusCounter[$status] = [int]$StatusCounter[$status] + 1
        }
        $nextAction = if ($row.PSObject.Properties['RemoteNextAction']) { [string]$row.RemoteNextAction } else { '' }
        if (-not [string]::IsNullOrWhiteSpace($nextAction)) {
            if (-not $NextActionCounter.ContainsKey($nextAction)) { $NextActionCounter[$nextAction] = 0 }
            $NextActionCounter[$nextAction] = [int]$NextActionCounter[$nextAction] + 1
        }
        if ($status -eq 'ADMIN_SHARE_UNREACHABLE') {
            $failureType = Get-Windows11AdminShareFailureType -Row $row
            if (-not $AdminShareFailureCounter.ContainsKey($failureType)) { $AdminShareFailureCounter[$failureType] = 0 }
            $AdminShareFailureCounter[$failureType] = [int]$AdminShareFailureCounter[$failureType] + 1
        }
    }
}

function ConvertTo-Windows11CounterRow {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Counter,
        [Parameter(Mandatory = $true)][string]$NameColumn
    )

    return @($Counter.GetEnumerator() | ForEach-Object {
        [pscustomobject][ordered]@{ $NameColumn = [string]$_.Key; Count = [int]$_.Value }
    } | Sort-Object Count, $NameColumn -Descending)
}

function New-Windows11CycleProgressRows {
    param(
        [Parameter(Mandatory = $true)][int]$CycleNumber,
        [Parameter(Mandatory = $true)][datetime]$CycleStart,
        [Parameter(Mandatory = $true)][int]$TotalComputers,
        [Parameter(Mandatory = $true)][int]$QueuedComputers,
        [Parameter(Mandatory = $true)][int]$CompletedComputers,
        [Parameter(Mandatory = $true)][int]$RunningComputers,
        [AllowNull()]$ComputerListStats
    )

    $remaining = [math]::Max(0, $TotalComputers - $CompletedComputers - $RunningComputers)
    $duplicateGroups = if ($ComputerListStats -and $ComputerListStats.PSObject.Properties['DuplicateGroups']) { [int]$ComputerListStats.DuplicateGroups } else { 0 }
    $duplicateLines = if ($ComputerListStats -and $ComputerListStats.PSObject.Properties['DuplicateLines']) { [int]$ComputerListStats.DuplicateLines } else { 0 }
    $duplicateSamples = if ($ComputerListStats -and $ComputerListStats.PSObject.Properties['DuplicateSamples']) { [string]$ComputerListStats.DuplicateSamples } else { '' }
    $rawLines = if ($ComputerListStats -and $ComputerListStats.PSObject.Properties['RawLines']) { [int]$ComputerListStats.RawLines } else { $TotalComputers }

    return @([pscustomobject]@{
        Cycle = $CycleNumber
        Started = $CycleStart.ToString('yyyy-MM-dd HH:mm:ss')
        ElapsedMinutes = [math]::Round(((Get-Date) - $CycleStart).TotalMinutes, 1)
        ComputerListLines = $rawLines
        TotalUnique = $TotalComputers
        Queued = $QueuedComputers
        CompletedRows = $CompletedComputers
        Running = $RunningComputers
        Remaining = $remaining
        DuplicateGroups = $duplicateGroups
        DuplicateLines = $duplicateLines
        DuplicateSamples = $duplicateSamples
    })
}

function New-Windows11RunningJobRows {
    param(
        [AllowEmptyCollection()][object[]]$RunningJobs,
        [Parameter(Mandatory = $true)][hashtable]$JobStartedAtById
    )

    $now = Get-Date
    return @($RunningJobs | ForEach-Object {
        $jobId = [string]$_.Id
        $computer = ([string]$_.Name) -replace '^W11UT_C\d+_',''
        $started = if ($JobStartedAtById.ContainsKey($jobId)) { [datetime]$JobStartedAtById[$jobId] } else { [datetime]::MinValue }
        [pscustomobject]@{
            ComputerName = $computer
            JobId = $jobId
            State = [string]$_.State
            Started = if ($started -gt [datetime]::MinValue) { $started.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
            ElapsedMinutes = if ($started -gt [datetime]::MinValue) { [math]::Round(($now - $started).TotalMinutes, 1) } else { '' }
        }
    })
}

function Export-Windows11ReportCsv {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory = $true)][string]$Path
    )

    if ($Rows -and $Rows.Count -gt 0) {
        @($Rows) | Select-Object -Property $reportColumns | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
        return
    }

    $emptyRow = [ordered]@{}
    foreach ($column in $reportColumns) { $emptyRow[$column] = '' }
    @([pscustomobject]$emptyRow) | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
    $lines = @(Get-Content -LiteralPath $Path -ErrorAction Stop)
    if ($lines.Count -gt 0) { Set-Content -LiteralPath $Path -Value $lines[0] -Encoding UTF8 }
}

$script:BrandLogoDataUri = $null
function Get-Windows11BrandLogoDataUri {
    if ($null -ne $script:BrandLogoDataUri) { return $script:BrandLogoDataUri }
    $script:BrandLogoDataUri = ''
    $candidates = @(
        (Join-Path $PSScriptRoot 'WorkplaceCloudHub-lockup-WPF.png'),
        (Join-Path (Split-Path -Parent $PSScriptRoot) 'WorkplaceCloudHub-lockup-WPF.png')
    )
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            try {
                $bytes = [System.IO.File]::ReadAllBytes($candidate)
                $script:BrandLogoDataUri = "data:image/png;base64," + [System.Convert]::ToBase64String($bytes)
                break
            }
            catch { }
        }
    }
    return $script:BrandLogoDataUri
}

function New-Windows11UpgradeCycleHtmlReport {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Summary,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][int]$CycleNumber,
        [Parameter(Mandatory = $true)][datetime]$GeneratedAt,
        [switch]$IsLive,
        [AllowEmptyCollection()][object[]]$CycleProgress = @(),
        [AllowEmptyCollection()][object[]]$RunningJobRows = @()
    )

    $rows = @($Summary | Where-Object { $null -ne $_ })
    $latestDataRows = @(Get-Windows11LatestRowsByComputer -Rows $rows)
    $latestRows = @(ConvertTo-Windows11HtmlReportRow -Items $latestDataRows)

    $separatedDetailStatuses = @('ADMIN_SHARE_UNREACHABLE', 'RUN_GUARD_ACTIVE', 'SKIPPED_BY_TECH_RUN_GUARD', 'SKIPPED_BY_TECH_RUN_GUARD_STARTED_NO_RESULT')
    $mainRows = @($latestRows | Where-Object { $separatedDetailStatuses -notcontains (Get-Windows11HtmlEffectiveStatus -Row $_) })
    $separatedDetailRows = @($latestRows | Where-Object { $separatedDetailStatuses -contains (Get-Windows11HtmlEffectiveStatus -Row $_) })

    $statusCounts = @(ConvertTo-Windows11CounterRow -Counter $script:AttemptStatusCounter -NameColumn 'Status')
    $nextActionCounts = @(ConvertTo-Windows11CounterRow -Counter $script:AttemptNextActionCounter -NameColumn 'NextAction')
    $latestEffectiveRows = foreach ($row in $latestRows) {
        $eff = Get-Windows11HtmlEffectiveStatus -Row $row
        [pscustomobject]@{ EffectiveStatus = $eff; NextAction = [string]$row.RemoteNextAction }
    }
    $latestStatusCounts = @($latestEffectiveRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.EffectiveStatus) } | Group-Object -Property EffectiveStatus | Sort-Object Count -Descending | ForEach-Object {
        [pscustomobject]@{ Status = $_.Name; Count = $_.Count }
    })
    $latestAdminShareFailureCounts = @($latestRows | Where-Object { (Get-Windows11HtmlEffectiveStatus -Row $_) -eq 'ADMIN_SHARE_UNREACHABLE' } | ForEach-Object { [pscustomobject]@{ FailureType = (Get-Windows11AdminShareFailureType -Row $_) } } | Group-Object -Property FailureType | Sort-Object Count -Descending | ForEach-Object {
        [pscustomobject]@{ FailureType = $_.Name; Count = $_.Count }
    })
    $adminShareFailureCounts = @(ConvertTo-Windows11CounterRow -Counter $script:AttemptAdminShareFailureCounter -NameColumn 'FailureType')

    $logoUri = Get-Windows11BrandLogoDataUri
    $logoHtml = if (-not [string]::IsNullOrWhiteSpace($logoUri)) { "<img class='logo' src='$logoUri' alt='WorkplaceCloudHub' />" } else { "" }
    $mode = if ($IsLive) { 'LIVE' } else { 'FINAL' }
    $lotName = ''
    $lotPath = [string]$script:LotRoot
    if (-not [string]::IsNullOrWhiteSpace($lotPath)) {
        try { $lotName = Split-Path -Leaf $lotPath } catch { $lotName = '' }
    }
    if ([string]::IsNullOrWhiteSpace($lotName)) { $lotName = 'Unknown LOT' }
    $lotNameHtml = ConvertTo-HtmlText $lotName
    $lotPathHtml = ConvertTo-HtmlText $lotPath
    $cycleValues = @(
        $rows |
            ForEach-Object {
                if ($_.PSObject.Properties['CycleNumber']) { [string]$_.CycleNumber }
            } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
    )
    $cycleLabel = if ($CycleNumber -le 0) {
        'all cycles'
    }
    elseif ($CycleNumber -gt 1 -or $cycleValues.Count -gt 1) {
        "cycles 1-$CycleNumber"
    }
    else {
        "cycle $CycleNumber"
    }
    $cycleLabelHtml = ConvertTo-HtmlText $cycleLabel

    $style = @"
<style>
body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 24px; background: #F5F8FB; color: #1F2937; }
.header-card { display: flex; align-items: center; justify-content: space-between; background: #fff; border: 1px solid #DDE7F0; border-radius: 8px; padding: 18px 22px; margin-bottom: 18px; }
.header-text .title { font-size: 22px; font-weight: 600; color: #1F2937; }
.header-text .subtitle { font-size: 13px; color: #5F6B7A; margin-top: 2px; }
.header-text .lot-name { font-size: 18px; font-weight: 700; color: #005A9E; margin-top: 8px; }
.header-text .meta { font-size: 12px; color: #5F6B7A; margin-top: 10px; }
.badge { display: inline-block; background: #E6F4FF; color: #005A9E; border: 1px solid #B9DDF7; border-radius: 10px; padding: 1px 10px; font-size: 11px; font-weight: 600; margin-left: 8px; vertical-align: middle; }
.logo { height: 46px; }
.card { background: #fff; border: 1px solid #DDE7F0; border-radius: 8px; padding: 14px 18px; margin-bottom: 16px; }
h2 { font-size: 15px; margin: 0 0 8px 0; color: #1F2937; }
table { border-collapse: collapse; width: 100%; font-size: 12px; }
th { background: #0078D4; color: #fff; text-align: left; font-weight: 600; }
th, td { border: 1px solid #DDE7F0; padding: 6px 8px; vertical-align: top; }
tr:nth-child(even) td { background: #F5F8FB; }
.empty { color: #5F6B7A; font-style: italic; }
.footer { font-size: 11px; color: #5F6B7A; margin-top: 8px; }
.footer a { color: #0078D4; text-decoration: none; }
</style>
"@

    $html = New-Object System.Collections.Generic.List[string]
    [void]$html.Add(("<html><head><meta charset='utf-8'><title>Windows 11 upgrade - {0} - {1}</title>{2}</head><body>" -f $lotNameHtml,$cycleLabelHtml,$style))
    [void]$html.Add("<div class='header-card'>")
    [void]$html.Add("<div class='header-text'>")
    [void]$html.Add(("<div class='title'>Windows 11 upgrade - {0}<span class='badge'>{1}</span></div>" -f $cycleLabelHtml,$mode))
    [void]$html.Add("<div class='subtitle'>Smart Intune Windows 11 Upgrade Toolkit</div>")
    [void]$html.Add(("<div class='lot-name' title='{1}'>LOT: {0}</div>" -f $lotNameHtml,$lotPathHtml))
    [void]$html.Add(("<div class='meta'>Generated: {0} | Completed attempt rows: {1} | Unique computers: {2} | Launcher: v{3}</div>" -f (ConvertTo-HtmlText $GeneratedAt.ToString('yyyy-MM-dd HH:mm:ss')),$script:AttemptRowCount,$latestRows.Count,(ConvertTo-HtmlText $script:LauncherVersion)))
    [void]$html.Add(("<div class='meta'>Launcher log: {0}</div>" -f (New-HtmlLogLink -Path $script:LauncherLogPath)))
    [void]$html.Add("</div>")
    [void]$html.Add($logoHtml)
    [void]$html.Add("</div>")
    $optionRows = @($script:LauncherOptionRows | ForEach-Object { $_ })
    $securityRows = @($script:PsExecSecurityEvidenceRows | ForEach-Object { $_ })
    [void]$html.Add("<div class='card'><h2>Cycle progress</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows @($CycleProgress) -Columns @('Cycle','Started','ElapsedMinutes','ComputerListLines','TotalUnique','Queued','CompletedRows','Running','Remaining','DuplicateGroups','DuplicateLines','DuplicateSamples')))
    [void]$html.Add("</div>")
    if ($RunningJobRows -and $RunningJobRows.Count -gt 0) {
        [void]$html.Add("<div class='card'><h2>Running jobs</h2>")
        [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows @($RunningJobRows) -Columns @('ComputerName','JobId','State','Started','ElapsedMinutes')))
        [void]$html.Add("</div>")
    }
    [void]$html.Add("<div class='card'><h2>Latest status by unique computer</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $latestStatusCounts -Columns @("Status", "Count")))
    [void]$html.Add("</div>")
    [void]$html.Add("<div class='card'><h2>Status summary by attempts</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $statusCounts -Columns @("Status", "Count")))
    [void]$html.Add("</div>")
    if ($latestAdminShareFailureCounts.Count -gt 0) {
        [void]$html.Add("<div class='card'><h2>Latest admin share failure by unique computer</h2>")
        [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $latestAdminShareFailureCounts -Columns @('FailureType','Count')))
        [void]$html.Add("</div>")
    }
    if ($adminShareFailureCounts.Count -gt 0) {
        [void]$html.Add("<div class='card'><h2>Admin share failure summary by attempts</h2>")
        [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $adminShareFailureCounts -Columns @('FailureType','Count')))
        [void]$html.Add("</div>")
    }
    [void]$html.Add("<div class='card'><h2>Next action summary by attempts</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $nextActionCounts -Columns @("NextAction", "Count")))
    [void]$html.Add("</div>")
    [void]$html.Add("<div class='card'><h2>Latest computer details</h2>")
    $htmlReportColumns = New-Object System.Collections.Generic.List[string]
    foreach ($column in $reportColumns) {
        [void]$htmlReportColumns.Add($column)
        if ($column -eq 'Detail') {
            [void]$htmlReportColumns.Add('Local log')
            [void]$htmlReportColumns.Add('Collected logs')
            [void]$htmlReportColumns.Add('Remote PC logs')
        }
    }
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $mainRows -Columns @($htmlReportColumns)))
    [void]$html.Add("<div class='footer'>Smart Intune Windows 11 Upgrade Toolkit - <a href='https://workplacecloudhub.com'>workplacecloudhub.com</a></div>")
    [void]$html.Add("</div>")
    [void]$html.Add("<div class='card'><h2>Latest run guard / admin share details</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $separatedDetailRows -Columns @($htmlReportColumns)))
    [void]$html.Add("<div class='footer'>Latest rows excluded from computer details: ADMIN_SHARE_UNREACHABLE, RUN_GUARD_ACTIVE, SKIPPED_BY_TECH_RUN_GUARD, and SKIPPED_BY_TECH_RUN_GUARD_STARTED_NO_RESULT. Per-cycle CSV files retain the complete attempt history.</div>")
    [void]$html.Add("</div>")
    [void]$html.Add("<div class='card'><h2>LOT/run options</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $optionRows -Columns @("Category", "Option", "Value")))
    [void]$html.Add("</div>")
    [void]$html.Add("<div class='card'><h2>Security evidence</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $securityRows -Columns @('Field', 'Value')))
    [void]$html.Add("<div class='footer'>Non-dry-run PsExec execution is blocked unless the binary is Microsoft-signed and named PsExec.exe or PsExec64.exe.</div>")
    [void]$html.Add("</div>")
    [void]$html.Add("</body></html>")

    ($html -join "`r`n") | Out-File -LiteralPath $Path -Encoding UTF8 -Force
}

$globalGateRoot = Join-Path ([System.IO.Path]::GetTempPath()) 'SmartM365\GlobalWorkerGates'
$globalGateName = ($GlobalConcurrencySemaphoreName -replace '[^A-Za-z0-9_.-]', '_')
if ($globalGateName.Length -gt 120) { $globalGateName = $globalGateName.Substring(0, 120) }
$globalGatePath = Join-Path $globalGateRoot $globalGateName
$globalGateMutexName = "Local\SmartM365_GlobalWorkerGate_$globalGateName"
$launcherInstanceId = [guid]::NewGuid().ToString('N')
if ($GlobalConcurrencyLimit -gt 0) {
    New-Directory -Path $globalGatePath
    Write-Host ("Global lease gate: Limit={0}; Path={1}; LeaseTimeout={2} minute(s)" -f $GlobalConcurrencyLimit,$globalGatePath,$GlobalConcurrencyLeaseTimeoutMinutes) -ForegroundColor Green
}
$script:globalGateMutex = $null
$script:GlobalGateMutexWaitSeconds = 300
if ($GlobalConcurrencyLimit -gt 0) {
    $script:globalGateMutex = New-Object System.Threading.Mutex($false, $globalGateMutexName)
}

function Invoke-WithGlobalGateMutex {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock,
        [object[]]$ArgumentList = @()
    )
    $sharedMutex = $script:globalGateMutex
    $mutex = if ($null -ne $sharedMutex) { $sharedMutex } else { New-Object System.Threading.Mutex($false, $globalGateMutexName) }
    $ownMutex = ($null -eq $sharedMutex)
    $acquired = $false
    try {
        try { $acquired = $mutex.WaitOne([TimeSpan]::FromSeconds($script:GlobalGateMutexWaitSeconds)) }
        catch [System.Threading.AbandonedMutexException] { $acquired = $true }
        if (-not $acquired) { throw ("Could not acquire global gate mutex within {0} seconds: {1}" -f $script:GlobalGateMutexWaitSeconds,$globalGateMutexName) }
        & $ScriptBlock @ArgumentList
    }
    finally {
        if ($acquired) { try { $mutex.ReleaseMutex() } catch { } }
        if ($ownMutex) { $mutex.Dispose() }
    }
}

function Test-GateProcessAlive {
    param([Parameter(Mandatory = $true)][int]$ProcessId)
    try { $null = Get-Process -Id $ProcessId -ErrorAction Stop; return $true } catch { return $false }
}

function Read-GlobalLeaseData {
    param([Parameter(Mandatory = $true)][string]$Path)

    $lastError = $null
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            return (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop)
        }
        catch {
            $lastError = $_
            Start-Sleep -Milliseconds ([math]::Min(1000, 100 * $attempt))
        }
    }

    if ($lastError) { throw $lastError }
    throw ("Failed to read global worker lease: {0}" -f $Path)
}

function Save-GlobalLeaseData {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$Data,
        [ValidateRange(1, 20)][int]$Depth = 4
    )

    $parent = Split-Path -Parent $Path
    New-Directory -Path $parent
    $json = $Data | ConvertTo-Json -Depth $Depth
    $lastError = $null
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        $tempPath = Join-Path $parent (".{0}.{1}.tmp" -f (Split-Path -Leaf $Path),[guid]::NewGuid().ToString('N'))
        try {
            Set-Content -LiteralPath $tempPath -Value $json -Encoding UTF8 -Force -ErrorAction Stop
            Move-Item -LiteralPath $tempPath -Destination $Path -Force -ErrorAction Stop
            return
        }
        catch {
            $lastError = $_
            Start-Sleep -Milliseconds ([math]::Min(1000, 100 * $attempt))
        }
        finally {
            if (Test-Path -LiteralPath $tempPath -PathType Leaf) {
                try { Remove-Item -LiteralPath $tempPath -Force -ErrorAction Stop } catch { }
            }
        }
    }

    if ($lastError) { throw $lastError }
    throw ("Failed to save global worker lease: {0}" -f $Path)
}

function Remove-GlobalLeaseFileUnlocked {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [ValidateRange(1, 20)][int]$Attempts = 5
    )

    for ($attempt = 1; $attempt -le $Attempts; $attempt++) {
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $true }
        try {
            Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
            if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $true }
        }
        catch [System.IO.IOException] { }
        catch [System.UnauthorizedAccessException] { }
        catch { }
        Start-Sleep -Milliseconds ([math]::Min(1000, 100 * $attempt))
    }

    return (-not (Test-Path -LiteralPath $Path -PathType Leaf))
}

function Remove-StaleGlobalLeases {
    $nowUtc = (Get-Date).ToUniversalTime()
    $removed = New-Object System.Collections.Generic.List[pscustomobject]
    foreach ($lease in @(Get-ChildItem -LiteralPath $globalGatePath -Filter '*.json' -File -ErrorAction SilentlyContinue)) {
        $remove = $false
        $reason = ''
        $computerName = ''
        $launcherPid = 0
        $workerPid = 0
        try {
            $data = Read-GlobalLeaseData -Path $lease.FullName
            $computerName = [string]$data.Computer
            $launcherPid = if ($data.PSObject.Properties['LauncherProcessId']) { [int]$data.LauncherProcessId } else { [int]$data.ProcessId }
            $workerPid = if ($data.PSObject.Properties['WorkerProcessId']) { [int]$data.WorkerProcessId } else { 0 }
            $createdUtc = [datetime]::Parse([string]$data.CreatedUtc, $null, [System.Globalization.DateTimeStyles]::RoundtripKind)
            if (-not (Test-GateProcessAlive -ProcessId $launcherPid)) {
                $remove = $true; $reason = 'LauncherProcessExited'
            } elseif ($workerPid -gt 0 -and -not (Test-GateProcessAlive -ProcessId $workerPid)) {
                $remove = $true; $reason = 'WorkerProcessExited'
            } elseif (($nowUtc - $createdUtc).TotalMinutes -gt $GlobalConcurrencyLeaseTimeoutMinutes) {
                $remove = $true; $reason = 'LeaseExpired'
            }
        } catch {
            $remove = $true; $reason = 'InvalidLease'
        }
        if ($remove) {
            if (Remove-GlobalLeaseFileUnlocked -Path $lease.FullName) {
                $removed.Add([pscustomobject]@{ Reason = $reason; Computer = $computerName; LauncherPid = $launcherPid })
            }
            else {
                Write-Host ("LEASE_RELEASE_DEFERRED: stale lease could not be removed after retries. Path={0}; Reason={1}" -f $lease.FullName,$reason) -ForegroundColor Yellow
            }
        }
    }
    if ($removed.Count -gt 0) {
        $pidParts = @($removed | Group-Object LauncherPid | Sort-Object Name | ForEach-Object {
            $reasons = @($_.Group | Group-Object Reason | ForEach-Object { if ($_.Count -gt 1) { "$($_.Name) x$($_.Count)" } else { $_.Name } }) -join '+'
            "PID=$($_.Name)[$reasons]"
        })
        $shortNames = @($removed | ForEach-Object { ($_.Computer -split '\.')[0] } | Where-Object { $_ })
        $computerPart = if ($shortNames.Count -le 4) { ': ' + ($shortNames -join ', ') } else { '' }
        Write-Host ("Removed {0} stale global worker lease(s) [{1}]{2}" -f $removed.Count,($pidParts -join '; '),$computerPart) -ForegroundColor DarkYellow
    }
}

function Update-GlobalLease {
    param(
        [AllowNull()][string]$LeasePath,
        [hashtable]$Properties = @{}
    )
    if ([string]::IsNullOrWhiteSpace($LeasePath) -or -not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return }
    Invoke-WithGlobalGateMutex -ScriptBlock {
        if (-not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return }
        $data = Read-GlobalLeaseData -Path $LeasePath
        foreach ($key in @($Properties.Keys)) {
            $data | Add-Member -NotePropertyName $key -NotePropertyValue $Properties[$key] -Force
        }
        $data | Add-Member -NotePropertyName LastUpdatedUtc -NotePropertyValue ((Get-Date).ToUniversalTime().ToString('o')) -Force
        Save-GlobalLeaseData -Path $LeasePath -Data $data -Depth 4
    }
}

function Acquire-GlobalLease {
    param(
        [Parameter(Mandatory = $true)][string]$Computer,
        [Parameter(Mandatory = $true)][int]$CycleNumber
    )
    if ($GlobalConcurrencyLimit -lt 1) { return '' }
    $waitStarted = Get-Date
    $lastWaitLog = $null
    $waitLogDelaySeconds = 300
    $waitLogIntervalSeconds = 300
    $waitWasLogged = $false
    while ($true) {
        if ((Get-LotCancellationState).Requested) { return '' }
        $leasePath = Invoke-WithGlobalGateMutex -ScriptBlock {
            Remove-StaleGlobalLeases
            $leases = @(Get-ChildItem -LiteralPath $globalGatePath -Filter '*.json' -File -ErrorAction SilentlyContinue)
            if ($leases.Count -lt $GlobalConcurrencyLimit) {
                $path = Join-Path $globalGatePath ("lease_{0}_{1}.json" -f $PID,([guid]::NewGuid().ToString('N')))
                $leaseData = [pscustomobject]@{
                    LauncherInstanceId = $launcherInstanceId
                    ProcessId = $PID
                    LauncherProcessId = $PID
                    Computer = $Computer
                    Cycle = $CycleNumber
                    JobId = ''
                    JobName = ''
                    WorkerProcessId = 0
                    WorkerProcessName = ''
                    WorkerStartedUtc = ''
                    CreatedUtc = (Get-Date).ToUniversalTime().ToString('o')
                    LastUpdatedUtc = (Get-Date).ToUniversalTime().ToString('o')
                    Host = $env:COMPUTERNAME
                }
                Save-GlobalLeaseData -Path $path -Data $leaseData -Depth 3
                return $path
            }
            return ''
        }
        if (-not [string]::IsNullOrWhiteSpace($leasePath)) {
            if ($waitWasLogged) {
                $waitMinutes = [math]::Round(((Get-Date) - $waitStarted).TotalMinutes, 1)
                Write-Host ("Global worker lease acquired for {0} after {1} minute(s)." -f $Computer,$waitMinutes) -ForegroundColor DarkCyan
            }
            return $leasePath
        }
        $now = Get-Date
        $waitSeconds = ($now - $waitStarted).TotalSeconds
        if ($waitSeconds -ge $waitLogDelaySeconds -and ($null -eq $lastWaitLog -or (($now - $lastWaitLog).TotalSeconds -ge $waitLogIntervalSeconds))) {
            $activeLeases = @(Get-ChildItem -LiteralPath $globalGatePath -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
            $waitMinutes = [math]::Round(($now - $waitStarted).TotalMinutes, 1)
            Write-Host ("Waiting for global worker lease: Computer={0}; Active={1}; Limit={2}; Wait={3} minute(s)." -f $Computer,$activeLeases,$GlobalConcurrencyLimit,$waitMinutes) -ForegroundColor DarkYellow
            $lastWaitLog = $now
            $waitWasLogged = $true
        }
        Start-Sleep -Seconds $JobPollSeconds
    }
}

function Test-SampleDnsResolution {
    param(
        [string[]]$ComputerNames,
        [int]$SampleSize = 5
    )

    $sample = @($ComputerNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First $SampleSize)
    $tested = 0
    $failed = 0
    foreach ($computer in $sample) {
        $tested++
        try {
            $addresses = [System.Net.Dns]::GetHostAddresses($computer)
            if ($null -eq $addresses -or $addresses.Count -lt 1) { $failed++ }
        }
        catch {
            $failed++
        }
    }

    [pscustomobject]@{
        Tested = $tested
        Failed = $failed
        AllFailed = ($tested -gt 0 -and $failed -eq $tested)
    }
}

function Release-GlobalLease {
    param([AllowNull()][string]$LeasePath)
    if ([string]::IsNullOrWhiteSpace($LeasePath)) { return }

    try {
        $removed = Invoke-WithGlobalGateMutex -ScriptBlock {
            Remove-GlobalLeaseFileUnlocked -Path $LeasePath
        }
        if (-not $removed) {
            Write-Host ("LEASE_RELEASE_DEFERRED: active lease could not be removed after retries. Path={0}" -f $LeasePath) -ForegroundColor Yellow
        }
        return
    }
    catch {
        Write-Host ("LEASE_RELEASE_DEFERRED: active lease cleanup failed without stopping the LOT. Path={0}; Error={1}" -f $LeasePath,$_.Exception.Message) -ForegroundColor Yellow
        return
    }
}

$script:IntuneInventoryMap = @{}
$script:AdInventoryMap = @{}
$script:IntuneInventoryLastRefreshUtc = [datetime]::MinValue
$script:AdInventoryLastRefreshUtc = [datetime]::MinValue
$script:IntuneInventoryRefreshScopeKeys = @()
$script:AdInventoryRefreshScopeKeys = @()
$cycle = 0
$mergedHtmlReportTimestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$mergedFinalHtmlPath = Join-Path $ReportRoot ("PsExec_Windows11Upgrade_Summary_{0}_{1}.html" -f $script:LauncherLogSafeLotName,$mergedHtmlReportTimestamp)
$mergedLiveHtmlPath = $mergedFinalHtmlPath
$script:LatestCycleResultByComputer = @{}
$script:AttemptStatusCounter = @{}
$script:AttemptNextActionCounter = @{}
$script:AttemptAdminShareFailureCounter = @{}
$script:AttemptRowCount = 0
$allCycleProgressRows = New-Object System.Collections.ArrayList
$script:LotStopReason = ''
Write-Host ("Merged HTML report: {0}" -f $mergedLiveHtmlPath) -ForegroundColor DarkCyan
Set-ActiveLotRunState -Status 'Running' -ReportPath $mergedLiveHtmlPath
do {
    $preCycleCancellation = Get-LotCancellationState
    if ($preCycleCancellation.Requested) {
        Write-Host ("Cancellation requested before cycle {0}; no new cycle will start." -f ($cycle + 1)) -ForegroundColor Yellow
        break
    }
    if ($cycle -gt 0 -and -not (Test-ComputerListPresentWithRetry -Path $ComputerListPath)) {
        $script:LotStopReason = 'LOT_INPUT_REMOVED'
        Write-Host ("[LOT_INPUT_REMOVED] Computers.txt disappeared after the LOT started. The launcher will preserve existing reports and stop cleanly. Path={0}" -f $ComputerListPath) -ForegroundColor Yellow
        Set-ActiveLotRunState -Status 'StoppedInputRemoved' -ReportPath $mergedLiveHtmlPath
        break
    }
    $cycle++
    try {
        $preCycleInventory = Invoke-Windows11InventoryPreCycleRefresh -CycleNumber $cycle
    }
    catch {
        if (-not (Test-ComputerListPresentWithRetry -Path $ComputerListPath)) {
            $script:LotStopReason = 'LOT_INPUT_REMOVED'
            Write-Host ("[LOT_INPUT_REMOVED] Computers.txt disappeared during cycle {0} preparation. Existing reports are preserved and no new worker will start. Path={1}" -f $cycle,$ComputerListPath) -ForegroundColor Yellow
            Set-ActiveLotRunState -Status 'StoppedInputRemoved' -ReportPath $mergedLiveHtmlPath
            break
        }
        throw
    }
    if ($preCycleInventory.MovedFromIntune -gt 0 -or $preCycleInventory.MovedFromAd -gt 0) {
        Write-Host ("Cycle {0}: inventory precheck removed {1} already-Windows11 computer(s); Remaining={2}" -f $cycle,($preCycleInventory.MovedFromIntune + $preCycleInventory.MovedFromAd),$preCycleInventory.RemainingComputers) -ForegroundColor Green
    }
    try {
        $computers = @(Get-ComputerList -Path $ComputerListPath)
        $computerListStats = Get-ComputerListStats -Path $ComputerListPath
    }
    catch {
        if (-not (Test-ComputerListPresentWithRetry -Path $ComputerListPath)) {
            $script:LotStopReason = 'LOT_INPUT_REMOVED'
            Write-Host ("[LOT_INPUT_REMOVED] Computers.txt disappeared before cycle {0} could queue workers. Existing reports are preserved. Path={1}" -f $cycle,$ComputerListPath) -ForegroundColor Yellow
            Set-ActiveLotRunState -Status 'StoppedInputRemoved' -ReportPath $mergedLiveHtmlPath
            break
        }
        throw
    }
    if ($computerListStats.DuplicateGroups -gt 0) {
        Write-Host ("Cycle {0}: Computers.txt contains {1} duplicate line(s) in {2} duplicate group(s). Duplicates ignored: {3}" -f $cycle,$computerListStats.DuplicateLines,$computerListStats.DuplicateGroups,$computerListStats.DuplicateSamples) -ForegroundColor Yellow
    }
    if ($computers.Count -eq 0) {
        Write-Host "No computers found in $ComputerListPath." -ForegroundColor Yellow
        break
    }

    $dnsCheck = Test-SampleDnsResolution -ComputerNames $computers
    if ($dnsCheck.AllFailed) {
        Write-Host ""
        Write-Host ("*** DNS WARNING: resolution failed on all {0} sampled computers. Check VPN connectivity and DNS configuration on this machine before continuing. ***" -f $dnsCheck.Tested) -ForegroundColor Red
        Write-Host ""
    }

    Write-Host ("Cycle {0}: {1} computer(s)." -f $cycle,$computers.Count) -ForegroundColor Cyan
    $results = New-Object System.Collections.ArrayList
    $runningJobs = @()
    $globalLeaseByJobId = @{}
    $techRunGuardFqdnByJobId = @{}
    $jobStartedAtById = @{}
    $nextIndex = 0
    $forcedCancelledJobIds = @{}
    $cancellationObservedAt = $null
    $lastLiveHtmlWrite = [datetime]::MinValue
    $lastLiveHtmlResultCount = -1
    $cycleStart = Get-Date
    $lastProgressLog = Get-Date
    try {
        $cycleProgress = New-Windows11CycleProgressRows -CycleNumber $cycle -CycleStart $cycleStart -TotalComputers $computers.Count -QueuedComputers $nextIndex -CompletedComputers $results.Count -RunningComputers $runningJobs.Count -ComputerListStats $computerListStats
        $mergedProgressRows = @($allCycleProgressRows.ToArray()) + @($cycleProgress)
        New-Windows11UpgradeCycleHtmlReport -Summary @($script:LatestCycleResultByComputer.Values) -Path $mergedLiveHtmlPath -CycleNumber $cycle -GeneratedAt (Get-Date) -IsLive -CycleProgress $mergedProgressRows -RunningJobRows @()
    }
    catch { Write-Host ("Cycle {0}: failed to initialize HTML report: {1}" -f $cycle,$_.Exception.Message) -ForegroundColor Yellow }
    while ($nextIndex -lt $computers.Count -or $runningJobs.Count -gt 0) {
        $cancelState = Get-LotCancellationState
        if ($cancelState.Requested) {
            if ($null -eq $cancellationObservedAt) {
                $cancellationObservedAt = Get-Date
                Set-ActiveLotRunState -Status 'StopRequested' -ReportPath $mergedLiveHtmlPath
                Write-Host ("Controlled stop requested by {0}: no additional computers will start. Active jobs will drain for up to {1} minute(s). Press Ctrl+C again to force their local stop." -f $cancelState.Source,$CancellationDrainTimeoutMinutes) -ForegroundColor Yellow
            }

            while ($nextIndex -lt $computers.Count) {
                $cancelledComputer = $computers[$nextIndex]
                $nextIndex++
                $cancelledResult = New-Windows11CancellationResult -ComputerName $cancelledComputer -CycleNumber $cycle -Status 'CANCELLED_NOT_STARTED' -Detail 'The operator requested a controlled stop before this computer was queued.'
                $cancelledResult = Add-AdInventoryFieldsToResult -Result $cancelledResult -AdInventoryMap $script:AdInventoryMap -AdInventoryCsv $AdInventoryCsv
                $cancelledResult = Add-IntuneInventoryFieldsToResult -Result $cancelledResult -IntuneInventoryMap $script:IntuneInventoryMap -IntuneInventoryCsv $IntuneInventoryCsv
                [void]$results.Add($cancelledResult)
            }

            $drainExpired = ($CancellationDrainTimeoutMinutes -eq 0 -or ((Get-Date) - $cancellationObservedAt).TotalMinutes -ge $CancellationDrainTimeoutMinutes)
            if ($runningJobs.Count -gt 0 -and ($cancelState.Force -or $drainExpired)) {
                foreach ($runningJob in @($runningJobs | Where-Object { $_.State -eq 'Running' })) {
                    $runningJobId = [string]$runningJob.Id
                    if (-not $forcedCancelledJobIds.ContainsKey($runningJobId)) {
                        $forcedCancelledJobIds[$runningJobId] = $true
                        Write-Host ("Forcing local worker stop for {0}. The remote endpoint may already be running and must be verified before relaunch." -f ($runningJob.Name -replace '^W11UT_C\d+_','')) -ForegroundColor Red
                    }
                    Stop-Job -Job $runningJob -ErrorAction SilentlyContinue
                }
            }
        }

        while ($nextIndex -lt $computers.Count -and $runningJobs.Count -lt $ThrottleLimit -and -not (Get-LotCancellationState).Requested) {
            $computer = $computers[$nextIndex]
            $nextIndex++

            $techRunGuardFqdn = Get-TechnicianRunGuardFqdn -ComputerName $computer -AdInventoryMap $script:AdInventoryMap
            if ($script:UseEffectiveTechnicianRunGuardHistory) {
                $activeTechRunGuard = $null
                try {
                    $activeTechRunGuard = Get-ActiveTechnicianRunGuardEntry -Path $script:TechnicianRunGuardHistoryPath -ComputerFqdn $techRunGuardFqdn -RunGuardHours $RunGuardHours
                }
                catch {
                    $historyError = "Technician run guard history is unavailable; this computer was not launched. Error=$($_.Exception.Message)"
                    $guardUnavailableResult = New-Windows11CancellationResult -ComputerName $computer -CycleNumber $cycle -Status 'TECH_RUN_GUARD_HISTORY_UNAVAILABLE' -Detail $historyError
                    $guardUnavailableResult = Add-AdInventoryFieldsToResult -Result $guardUnavailableResult -AdInventoryMap $script:AdInventoryMap -AdInventoryCsv $AdInventoryCsv
                    $guardUnavailableResult = Add-IntuneInventoryFieldsToResult -Result $guardUnavailableResult -IntuneInventoryMap $script:IntuneInventoryMap -IntuneInventoryCsv $IntuneInventoryCsv
                    [void]$results.Add($guardUnavailableResult)
                    Write-Host ("  [TECH_RUN_GUARD_HISTORY_UNAVAILABLE] {0}: {1}" -f $computer,$historyError) -ForegroundColor Yellow
                    continue
                }
                if ($null -ne $activeTechRunGuard) {
                    $skipResult = New-TechnicianRunGuardSkippedResult -ComputerName $computer -CycleNumber $cycle -HistoryEntry $activeTechRunGuard -RunGuardHours $RunGuardHours
                    $skipResult = Add-AdInventoryFieldsToResult -Result $skipResult -AdInventoryMap $script:AdInventoryMap -AdInventoryCsv $AdInventoryCsv
                    $skipResult = Add-IntuneInventoryFieldsToResult -Result $skipResult -IntuneInventoryMap $script:IntuneInventoryMap -IntuneInventoryCsv $IntuneInventoryCsv
                    [void]$results.Add($skipResult)
                    Write-Host ("  [{0}] {1}: {2}" -f $skipResult.LauncherStatus,$computer,$skipResult.Detail) -ForegroundColor DarkYellow
                    continue
                }
            }

            $globalLeasePath = Acquire-GlobalLease -Computer $computer -CycleNumber $cycle
            if ((Get-LotCancellationState).Requested) {
                Release-GlobalLease -LeasePath $globalLeasePath
                $nextIndex--
                break
            }


            try {
                $remoteArgsJson = ([pscustomobject]@{ Args = @($remoteArgs.ToArray()) } | ConvertTo-Json -Compress)
                $workerArgs = @(
                    $computer,
                    $cycle,
                    $remoteArgsJson,
                    $resolvedPsExec,
                    $LocalScriptPath,
                    $script:RemoteBaseDir,
                    $script:RemoteScriptPath,
                    $script:RemoteSetupCacheRoot,
                    $LogRoot,
                    $CentralLogRoot,
                    $SetupSourcePath,
                    $SetupSourceMapPath,
                    $SetupExecutionMode,
                    $SetupMediaId,
                    $SetupLanguage,
                    [bool]$AllowSetupUpgrade,
                    [bool]$SkipSetupMediaPreCopy,
                    [bool]$SkipVirtualMachines,
                    [bool]$DryRun,
                    [bool]$NoCentralLogCollection,
                    [bool]$KeepCentralLogHistory,
                    $CentralLogCollectionMode,
                    $PsExecTimeoutMinutes,
                    $globalLeasePath,
                    $globalGateMutexName
                )

                $job = Start-Job -Name ("W11UT_C{0}_{1}" -f $cycle,$computer) -FilePath $LocalWorkerPath -ArgumentList $workerArgs
                $jobStartedAtById[[string]$job.Id] = Get-Date
                if ($script:UseEffectiveTechnicianRunGuardHistory) {
                    $techRunGuardFqdnByJobId[[string]$job.Id] = $techRunGuardFqdn
                    try {
                        Update-TechnicianRunGuardHistory -Path $script:TechnicianRunGuardHistoryPath -ComputerFqdn $techRunGuardFqdn -InputComputerName $computer -RunGuardHours $RunGuardHours -State Started -Result $null -JobId ([string]$job.Id)
                    }
                    catch {
                        Write-Host ("TECH_RUN_GUARD_HISTORY_WRITE_DEFERRED: State=Started; Computer={0}; Error={1}" -f $computer,$_.Exception.Message) -ForegroundColor Yellow
                    }
                }
            }
            catch {
                Release-GlobalLease -LeasePath $globalLeasePath
                throw
            }

            if (-not [string]::IsNullOrWhiteSpace($globalLeasePath)) {
                Update-GlobalLease -LeasePath $globalLeasePath -Properties @{
                    JobId = [string]$job.Id
                    JobName = [string]$job.Name
                }
                $globalLeaseByJobId[[string]$job.Id] = $globalLeasePath
            }
            $runningJobs += $job
            Write-Host ("Queued {0} ({1}/{2}); running={3}" -f $computer,$nextIndex,$computers.Count,$runningJobs.Count) -ForegroundColor DarkCyan

            if ($DelayBetweenComputersSeconds -gt 0 -and $nextIndex -lt $computers.Count) {
                [void](Wait-LotCancellationAware -Seconds $DelayBetweenComputersSeconds)
            }
        }

        $finishedJobs = @($runningJobs | Where-Object { $_.State -ne 'Running' })
        if ($finishedJobs.Count -eq 0) {
            if (((Get-Date) - $lastProgressLog).TotalSeconds -ge 300) {
                $now = Get-Date
                $waitingNames = @($runningJobs | ForEach-Object {
                    $jobId = [string]$_.Id
                    $jobName = $_.Name -replace '^W11UT_C\d+_',''
                    if ($jobStartedAtById.ContainsKey($jobId)) {
                        "{0} ({1}m)" -f $jobName,[math]::Round(($now - $jobStartedAtById[$jobId]).TotalMinutes,1)
                    }
                    else {
                        $jobName
                    }
                })
                Write-Host ("Waiting for {0} job(s); Elapsed={1} min; Running: {2}" -f $runningJobs.Count,[math]::Round(($now - $cycleStart).TotalMinutes,1),($waitingNames -join ', '))
                $lastProgressLog = $now
            }
            if (((Get-Date) - $lastLiveHtmlWrite).TotalSeconds -ge 60) {
                try {
                    $liveRows = @(Get-Windows11ReportRows -Items @($results.ToArray()))
                    $cycleProgress = New-Windows11CycleProgressRows -CycleNumber $cycle -CycleStart $cycleStart -TotalComputers $computers.Count -QueuedComputers $nextIndex -CompletedComputers $results.Count -RunningComputers $runningJobs.Count -ComputerListStats $computerListStats
                    $runningJobRows = New-Windows11RunningJobRows -RunningJobs @($runningJobs) -JobStartedAtById $jobStartedAtById
                    $mergedLiveRows = @($script:LatestCycleResultByComputer.Values) + $liveRows
                    $mergedProgressRows = @($allCycleProgressRows.ToArray()) + @($cycleProgress)
                    New-Windows11UpgradeCycleHtmlReport -Summary $mergedLiveRows -Path $mergedLiveHtmlPath -CycleNumber $cycle -GeneratedAt (Get-Date) -IsLive -CycleProgress $mergedProgressRows -RunningJobRows $runningJobRows
                    $lastLiveHtmlWrite = Get-Date
                    $lastLiveHtmlResultCount = $results.Count
                }
                catch { Write-Host ("Cycle {0}: failed to update HTML report: {1}" -f $cycle,$_.Exception.Message) -ForegroundColor Yellow }
            }
            Start-Sleep -Seconds $JobPollSeconds
            continue
        }

        foreach ($job in $finishedJobs) {
            $received = $null
            $jobErrors = @()

            try {
                $receiveErrors = @()
                $received = @(Receive-Job -Job $job -ErrorAction SilentlyContinue -ErrorVariable receiveErrors)
                $childJobs = @($job.ChildJobs)
                $brokenRunspace = (
                    $childJobs.Count -gt 0 -and
                    $null -ne $childJobs[0].JobStateInfo.Reason -and
                    [string]$childJobs[0].JobStateInfo.Reason.Message -match '\bBroken\b'
                )
                $jobErrors = @(
                    @(
                        $receiveErrors | ForEach-Object { $_.ToString() }
                        $childJobs | ForEach-Object { $_.Error } | ForEach-Object { $_.ToString() }
                        if ($childJobs.Count -gt 0 -and $childJobs[0].JobStateInfo.Reason) {
                            $childJobs[0].JobStateInfo.Reason.Message
                        }
                        if ($brokenRunspace) {
                            'RUNSPACE_BROKEN: PowerShell job runspace entered a Broken state. This typically indicates PowerShell 5.1 instability under parallel load. Consider adding -DelayBetweenComputersSeconds 1 to reduce job cycling speed.'
                        }
                    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
                )
                if ($brokenRunspace) {
                    Write-Host ("  [RUNSPACE_BROKEN] {0}: job runspace entered Broken state - possible PowerShell 5.1 instability. Consider -DelayBetweenComputersSeconds 1." -f ($job.Name -replace '^W11UT_C\d+_','')) -ForegroundColor Magenta
                }
            }
            catch {
                $jobErrors += $_.Exception.Message
            }

            if (-not $received -or $received.Count -eq 0) {
            $forcedCancellation = $forcedCancelledJobIds.ContainsKey([string]$job.Id)
                $received = @([pscustomobject]@{
                    Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                    ComputerName = ($job.Name -replace '^W11UT_C\d+_','')
                    CycleNumber = $cycle
                    LauncherStatus = if ($forcedCancellation) { 'CANCELLED_BY_OPERATOR' } elseif (@($jobErrors | Where-Object { $_ -match 'RUNSPACE_BROKEN' }).Count -gt 0) { 'RUNSPACE_BROKEN' } else { 'JOB_ERROR' }
                    RemoteStatus = ''
                    RemoteNextAction = if ($forcedCancellation) { 'VERIFY_REMOTE_STATE_BEFORE_RELAUNCH' } else { '' }
                    ExitCode = ''
                    Detail = if ($forcedCancellation) { 'The operator forced the local worker to stop. The remote endpoint may already be running; verify target evidence before relaunch.' } else { ($jobErrors -join ' | ') }
                    SetupCacheAction = ''
                    DiskCleanupAction = ''
                    DiskCleanupFreedGB = ''
                    AdvancedDiskCleanupAction = ''
                    AdvancedDiskCleanupFreedGB = ''
                    DismCleanupAction = ''
                    DismCleanupFreedGB = ''
                    SetupCompletionRebootAction = ''
                    SetupCompletionRebootDetail = ''
                    SetupCompletionRebootUserCount = ''
                    SetupCompletionRebootUsers = ''
                    SetupProfileRepairAction = ''
                    SetupProfileRepairDetail = ''
                    SetupProfileRepairBlockingSid = ''
                    SetupProfileRepairKeptSid = ''
                    SetupProfileRepairProfilePath = ''
                    SetupProfileRepairBackupPath = ''
                    ControlledRebootAction = ''
                    ControlledRebootDetail = ''
                    ControlledRebootUserCount = ''
                    ControlledRebootUsers = ''
                    UserRebootNotificationSent = ''
                    UserRebootNotificationLang = ''
                    UserRebootNotificationMessage = ''
                    RemoteLogsPath = ''
                    PsExecLogPath = ''
                    JobErrorMessage = ($jobErrors -join ' | ')
                })
            }

            foreach ($item in @($received)) {
                if ($null -ne $item) {
                    if ($jobErrors.Count -gt 0 -and -not $item.PSObject.Properties['JobErrorMessage']) {
                        $item | Add-Member -NotePropertyName JobErrorMessage -NotePropertyValue ($jobErrors -join ' | ') -Force
                    }
                    $item = Add-AdInventoryFieldsToResult -Result $item -AdInventoryMap $script:AdInventoryMap -AdInventoryCsv $AdInventoryCsv
                    $item = Add-IntuneInventoryFieldsToResult -Result $item -IntuneInventoryMap $script:IntuneInventoryMap -IntuneInventoryCsv $IntuneInventoryCsv
                    $jobIdForLog = [string]$job.Id
                    $jobElapsedMinutes = if ($jobStartedAtById.ContainsKey($jobIdForLog)) { [math]::Round(((Get-Date) - $jobStartedAtById[$jobIdForLog]).TotalMinutes,1) } else { '' }
                    $computerForLog = if ($item.PSObject.Properties['ComputerName']) { [string]$item.ComputerName } else { ($job.Name -replace '^W11UT_C\d+_','') }
                    $launcherStatusForLog = if ($item.PSObject.Properties['LauncherStatus']) { [string]$item.LauncherStatus } else { '' }
                    $remoteStatusForLog = if ($item.PSObject.Properties['RemoteStatus']) { [string]$item.RemoteStatus } else { '' }
                    $statusForLog = if (-not [string]::IsNullOrWhiteSpace($remoteStatusForLog)) { $remoteStatusForLog } else { $launcherStatusForLog }
                    $nextActionForLog = if ($item.PSObject.Properties['RemoteNextAction']) { [string]$item.RemoteNextAction } else { '' }
                    $exitCodeForLog = if ($item.PSObject.Properties['ExitCode']) { [string]$item.ExitCode } else { '' }
                    $detailForLog = if ($item.PSObject.Properties['Detail']) { [string]$item.Detail } else { '' }
                    if ($detailForLog.Length -gt 180) { $detailForLog = $detailForLog.Substring(0,177) + '...' }
                    $completionColor = 'Red'
                    if ($exitCodeForLog -eq '0' -or $statusForLog -match 'ALREADY_WINDOWS11|SUCCESS|STARTED') { $completionColor = 'Green' }
                    elseif ($statusForLog -match 'RUN_GUARD|PENDING|INSUFFICIENT|TIMEOUT|SKIPPED|READY') { $completionColor = 'Yellow' }
                    Write-Host ("Completed {0}; Elapsed={1} min; Status={2}; NextAction={3}; ExitCode={4}; Detail={5}" -f $computerForLog,$jobElapsedMinutes,$statusForLog,$nextActionForLog,$exitCodeForLog,$detailForLog) -ForegroundColor $completionColor
                    if ($script:UseEffectiveTechnicianRunGuardHistory) {
                        $resultFqdn = if ($techRunGuardFqdnByJobId.ContainsKey([string]$job.Id)) { [string]$techRunGuardFqdnByJobId[[string]$job.Id] } else { Get-TechnicianRunGuardFqdn -ComputerName ([string]$item.ComputerName) -AdInventoryMap $script:AdInventoryMap }
                        try {
                            Update-TechnicianRunGuardHistory -Path $script:TechnicianRunGuardHistoryPath -ComputerFqdn $resultFqdn -InputComputerName ([string]$item.ComputerName) -RunGuardHours $RunGuardHours -State Result -Result $item -JobId ([string]$job.Id)
                        }
                        catch {
                            $historyWarning = "TECH_RUN_GUARD_HISTORY_WRITE_DEFERRED: State=Result; Computer=$($item.ComputerName); Error=$($_.Exception.Message)"
                            Write-Host $historyWarning -ForegroundColor Yellow
                            if ($item.PSObject.Properties['JobErrorMessage']) {
                                $existingJobError = [string]$item.JobErrorMessage
                                $item.JobErrorMessage = if ([string]::IsNullOrWhiteSpace($existingJobError)) { $historyWarning } else { "$existingJobError | $historyWarning" }
                            }
                            else {
                                $item | Add-Member -NotePropertyName JobErrorMessage -NotePropertyValue $historyWarning -Force
                            }
                        }
                    }
                    [void]$results.Add($item)
                    if (-not $DryRun -and (Test-AlreadyWindows11CycleResult -Result $item)) {
                        try {
                            $moveSingleResult = Move-AlreadyWindows11ComputersFromList -ComputerListPath $ComputerListPath -CycleSummary @($item)
                            if ($moveSingleResult.Moved -gt 0) {
                                Write-Host ("Moved already-Windows11 computer from Computers.txt to {0}: {1}" -f $moveSingleResult.AlreadyWindows11Path,$item.ComputerName) -ForegroundColor Green
                            }
                        }
                        catch {
                            Write-Host ("Failed to update ComputersAlreadyW11.txt for {0}: {1}" -f $item.ComputerName,$_.Exception.Message) -ForegroundColor Yellow
                        }
                    }
                }
            }
            if ($globalLeaseByJobId.ContainsKey([string]$job.Id)) {
                Release-GlobalLease -LeasePath $globalLeaseByJobId[[string]$job.Id]
                $globalLeaseByJobId.Remove([string]$job.Id)
            }
            if ($techRunGuardFqdnByJobId.ContainsKey([string]$job.Id)) {
                $techRunGuardFqdnByJobId.Remove([string]$job.Id)
            }
            if ($jobStartedAtById.ContainsKey([string]$job.Id)) {
                $jobStartedAtById.Remove([string]$job.Id)
            }
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }

        $currentRunningJobs = @($runningJobs | Where-Object { $_.State -eq 'Running' })
        $liveHtmlAgeSeconds = ((Get-Date) - $lastLiveHtmlWrite).TotalSeconds
        if ($liveHtmlAgeSeconds -ge 15 -and ($results.Count -ne $lastLiveHtmlResultCount -or $liveHtmlAgeSeconds -ge 60)) {
            try {
                $liveRows = @(Get-Windows11ReportRows -Items @($results.ToArray()))
                $cycleProgress = New-Windows11CycleProgressRows -CycleNumber $cycle -CycleStart $cycleStart -TotalComputers $computers.Count -QueuedComputers $nextIndex -CompletedComputers $results.Count -RunningComputers $currentRunningJobs.Count -ComputerListStats $computerListStats
                $runningJobRows = New-Windows11RunningJobRows -RunningJobs $currentRunningJobs -JobStartedAtById $jobStartedAtById
                $mergedLiveRows = @($script:LatestCycleResultByComputer.Values) + $liveRows
                $mergedProgressRows = @($allCycleProgressRows.ToArray()) + @($cycleProgress)
                New-Windows11UpgradeCycleHtmlReport -Summary $mergedLiveRows -Path $mergedLiveHtmlPath -CycleNumber $cycle -GeneratedAt (Get-Date) -IsLive -CycleProgress $mergedProgressRows -RunningJobRows $runningJobRows
                $lastLiveHtmlWrite = Get-Date
                $lastLiveHtmlResultCount = $results.Count
            }
            catch { Write-Host ("Cycle {0}: failed to update HTML report: {1}" -f $cycle,$_.Exception.Message) -ForegroundColor Yellow }
        }

        $runningJobs = $currentRunningJobs
    }

    $reportTimestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $reportPath = Join-Path $ReportRoot ("PsExec_Windows11Upgrade_Summary_{0}_cycle{1}_{2}.csv" -f $script:LauncherLogSafeLotName,$cycle,$reportTimestamp)
    $enrichedResults = @($results.ToArray() | ForEach-Object { $row = Add-AdInventoryFieldsToResult -Result $_ -AdInventoryMap $script:AdInventoryMap -AdInventoryCsv $AdInventoryCsv; Add-IntuneInventoryFieldsToResult -Result $row -IntuneInventoryMap $script:IntuneInventoryMap -IntuneInventoryCsv $IntuneInventoryCsv })
    $normalizedResults = @(Get-Windows11ReportRows -Items @($enrichedResults))
    Export-Windows11ReportCsv -Rows @($normalizedResults) -Path $reportPath
    Write-Host ("Cycle {0} report: {1}" -f $cycle,$reportPath) -ForegroundColor Green

    try {
        $finalProgress = New-Windows11CycleProgressRows -CycleNumber $cycle -CycleStart $cycleStart -TotalComputers $computers.Count -QueuedComputers $computers.Count -CompletedComputers $normalizedResults.Count -RunningComputers 0 -ComputerListStats $computerListStats
        Add-Windows11AttemptCounter -Rows $normalizedResults -StatusCounter $script:AttemptStatusCounter -NextActionCounter $script:AttemptNextActionCounter -AdminShareFailureCounter $script:AttemptAdminShareFailureCounter
        $script:AttemptRowCount += $normalizedResults.Count
        foreach ($normalizedResult in @($normalizedResults)) {
            $computerKey = Get-ComputerListKey -ComputerName ([string]$normalizedResult.ComputerName)
            if ([string]::IsNullOrWhiteSpace($computerKey)) { $computerKey = [string]$normalizedResult.ComputerName }
            $script:LatestCycleResultByComputer[$computerKey] = $normalizedResult
        }
        [void]$allCycleProgressRows.Add($finalProgress)
        New-Windows11UpgradeCycleHtmlReport -Summary @($script:LatestCycleResultByComputer.Values) -Path $mergedLiveHtmlPath -CycleNumber $cycle -GeneratedAt (Get-Date) -IsLive -CycleProgress @($allCycleProgressRows.ToArray()) -RunningJobRows @()
        Write-Host ("Merged HTML report updated through cycle {0}: {1}" -f $cycle,$mergedLiveHtmlPath) -ForegroundColor Green
    }
    catch { Write-Host ("Cycle {0}: failed to update merged HTML report: {1}" -f $cycle,$_.Exception.Message) -ForegroundColor Yellow }

    try {
        $moveAlreadyW11Result = Move-AlreadyWindows11ComputersFromList -ComputerListPath $ComputerListPath -CycleSummary @($normalizedResults)
        if ($moveAlreadyW11Result.Moved -gt 0) {
            Write-Host ("Cycle {0}: moved {1} already-Windows11 computer(s) to {2}" -f $cycle,$moveAlreadyW11Result.Moved,$moveAlreadyW11Result.AlreadyWindows11Path) -ForegroundColor Green
        }
    }
    catch {
        Write-Host ("Cycle {0}: failed to update ComputersAlreadyW11.txt: {1}" -f $cycle,$_.Exception.Message) -ForegroundColor Yellow
    }

    $sourceDistribution = @(
        $normalizedResults |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.SelectedSetupSourcePath) } |
            Group-Object -Property SelectedSetupSourcePath |
            Sort-Object Count -Descending |
            ForEach-Object {
                [pscustomobject]@{
                    SelectedSetupSourcePath = [string]$_.Name
                    ComputerCount = [int]$_.Count
                }
            }
    )
    if ($sourceDistribution.Count -gt 0) {
        $distributionPath = Join-Path $ReportRoot ("SetupSource_Distribution_{0}_cycle{1}_{2}.csv" -f $script:LauncherLogSafeLotName,$cycle,(Get-Date -Format 'yyyyMMdd-HHmmss'))
        $sourceDistribution | Export-Csv -LiteralPath $distributionPath -NoTypeInformation -Encoding UTF8
        Write-Host ("Setup source distribution: {0}" -f $distributionPath) -ForegroundColor Green
        foreach ($sourceGroup in $sourceDistribution) {
            Write-Host ("  {0}: {1}" -f $sourceGroup.SelectedSetupSourcePath,$sourceGroup.ComputerCount) -ForegroundColor DarkGreen
        }
    }
    if ((Get-LotCancellationState).Requested) { break }

    if ($RunOnce) { break }
    if ($MaxCycles -gt 0 -and $cycle -ge $MaxCycles) { break }
    if ($DelayBetweenCyclesMinutes -gt 0) {
        Write-Host ("Waiting {0} minute(s) before next cycle." -f $DelayBetweenCyclesMinutes) -ForegroundColor DarkYellow
        [void](Wait-LotCancellationAware -Seconds ($DelayBetweenCyclesMinutes * 60))
    }
}
while ($true)

try {
    if ($script:LatestCycleResultByComputer.Count -gt 0) {
        New-Windows11UpgradeCycleHtmlReport -Summary @($script:LatestCycleResultByComputer.Values) -Path $mergedFinalHtmlPath -CycleNumber $cycle -GeneratedAt (Get-Date) -CycleProgress @($allCycleProgressRows.ToArray()) -RunningJobRows @()
        Write-Host ("Merged final HTML report: {0}" -f $mergedFinalHtmlPath) -ForegroundColor Green
    }
}
catch { Write-Host ("Failed to write merged final HTML report: {0}" -f $_.Exception.Message) -ForegroundColor Yellow }

Write-Host ("SmartM365 Windows 11 Upgrade Toolkit launcher finished. Log={0}" -f $script:LauncherLogPath)

if ($null -ne $script:globalGateMutex) {
    try { $script:globalGateMutex.Dispose() } catch { }
    $script:globalGateMutex = $null
}

$finalLotRunStatus = if ($script:LotStopReason -eq 'LOT_INPUT_REMOVED') { 'StoppedInputRemoved' } else { 'Finished' }
Set-ActiveLotRunState -Status $finalLotRunStatus -ReportPath $mergedFinalHtmlPath
Complete-LotCancellationSupport

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCXsa9wZ42S/Pm+
# fReBvr30h5c00Q1XqF812FXYG4OT3KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIBIbiw5JyLNuDdpUNuP79P75NJpcdiRqN4pcOdcVbWQ1MA0GCSqG
# SIb3DQEBAQUABIIBgKvIoGOAYyjrjC2KeTdjkym3o2t+VIZevfOCTVBTYS+T6XWz
# CEA+w0vCNbr93G0EBIi8YQqhcr8s0ghgE7iv2+MkgztMAUW6gIN26mtK8GbtxVHC
# OdkRlyEOLsdBoimNZb53+/q5YLm0+DPxjoke7DmTX2qcXF+jUppaMy/9tWE90b+U
# P7qwHUf7jBeY5rqKZNdmRmlJu6fSnds9t488gXJV4QmrDvI4SD371CqoSSftbQN1
# MvHjjmRE6P1eR02h6uppfWF3gpZvNu74gArtnRhp3j8eWJDjiORcKO3fOwaDle40
# WcXt1jLXTZt26Q07cWo5U//hab//Wa1gx4t/dNAODAclgiv02Kz0PR1btjMcy7HI
# ue7vFuW2yZR3xLYSSNZPnTMgq6xBB26DHGFs0X673OFcKN18H7E4RD6RhIhk4BNd
# DPihdavt+89nJ7wsHAuwQrN7xrlWIPs7DjTtMUce8I1SnKGTPSlIIgk4qIKp5Vez
# TxpvSPkbPMuD4Kp1J6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjQwNzEx
# MzZaMC8GCSqGSIb3DQEJBDEiBCCEO9AFW7s9FAxyEzk9LgGHyFfPyw9TuMnViIx2
# D2vvUTANBgkqhkiG9w0BAQEFAASCAgAbabCUIrfPrEoSh6H7WbDQuqD9HGw8CeHi
# UC1jATIKzXfEdBVJta2LU5gBzHfs/iqSK/RDRjgEXj/SQ5KPzXOrY0oQuktmSFSe
# WKwremRsIecHLCXPl0t6o3TIcx++N8ZnbAZP9a1bj7Q6hxlebraX9uKh4fH1Q0EB
# HytmXthpynupfnP4nGi2Yqv5QCu1Lu40PqY7e1Nu8iPSrs6OZNVPT/P/YxJ56AFP
# eMQfna0xf18JZMTRZ3Fizha5iN9yNzGb6K1KxyH4BGBsq1LWjJE4HPQiSpYeBQA4
# gKvmI5Ft//ekBTz3zHPIXmKYY+pqpbp3mVFA6qdgkZR9pBbRDh+d1g5xFZmfXqBv
# kTd4SxJaFGj0zu6xNb4SoSiQ/V0yTEfGglxKQgb64Ui0cUbWCGRLLs8blFdV8v3K
# QMImGquDNDXwj8+ozV2mtaLvDth7mHGxfa0hSmWA5poQzZMkSIquHwYgQgS5Bo03
# 6Ws99LOSAIUyA/e+oniqij4C/ph5fvs1K5yUh6IkIgqLFcO2oxzsP6apSnPhjdHZ
# mDjAJfMcUZQOC+aSvbWgImkeqC97Flxeh0NRE30PhhLkI8PDd1CTrz9ZzUwyX418
# KB+S0rdINrPm+Tvh4TaF6hjIyn+xjE4lXWX90h1ArEqmCXUdJLsAxwTAbyND1Iq6
# O8trXLMXng==
# SIG # End signature block
