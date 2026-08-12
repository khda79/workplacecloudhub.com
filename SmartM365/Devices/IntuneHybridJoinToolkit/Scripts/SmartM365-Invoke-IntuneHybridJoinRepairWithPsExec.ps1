<#
.SYNOPSIS
Runs SmartM365-Invoke-IntuneHybridJoinRepair.ps1 on a list of remote computers using PsExec.

.DESCRIPTION
Reads one computer name per line from Computers.txt by default, copies the repair script to
C:\ProgramData\SmartM365\IntuneHybridJoinToolkit on each target through C$, then starts it as
LocalSystem with PsExec.

Lines starting with # in the computer list are ignored.
The launcher loops by default. It reloads Computers.txt at the beginning of each cycle.
After each cycle, computers detected as already Intune-enrolled are removed from Computers.txt
and appended to ComputersAlreadyEnrolled.txt in the same folder.

.PARAMETER ComputerListPath
Path to a text file containing one computer name per line.

.PARAMETER PsExecPath
Path to PsExec.exe. Defaults to PsExec.exe next to this launcher, then falls back to PATH.
The launcher validates that PsExec is available before starting any non-dry-run cycle.

.PARAMETER AllowDsregLeave
Passes -AllowDsregLeave to the repair script. Enabled by default in this launcher.
Use -AllowDsregLeave:$false to run the remote repair without dsregcmd /leave authorization.

.PARAMETER IgnoreRunGuard
Passes -IgnoreRunGuard to the repair script to bypass the target computer's 12-hour guard.
In loop mode, this is passed only during the first cycle unless -IgnoreRunGuardEveryCycle is set.

.PARAMETER AllowRebootWhenNoInteractiveUser
Passes -AllowRebootWhenNoInteractiveUser to the repair script.

.PARAMETER AllowRebootAfterDsregLeave
Passes -AllowRebootAfterDsregLeave to the repair script.

.PARAMETER AllowRemoveNonIntuneMdmEnrollment
Passes -AllowRemoveNonIntuneMdmEnrollment to the repair script. Use only when a non-Intune MDM should be removed before Intune enrollment.

.PARAMETER AllowRemoveStaleIntuneEnrollment
Passes -AllowRemoveStaleIntuneEnrollment to the repair script. Use only to clean stale local Intune enrollment traces.

.PARAMETER SkipVirtualMachines
Skips detected virtual machines before copy or repair.

.PARAMETER AuditOnly
Passes -AuditOnly to the repair script to collect diagnostics without repair actions.

.PARAMETER IntuneInventoryCsv
Optional local CSV export of Intune/Graph devices. Defaults to DevicesIntune.csv in the parent folder, then DevicesIntune.csv next to this script. The launcher adds an IntuneInventoryPresent column by matching computer names.

.PARAMETER EntraInventoryCsv
Optional local CSV export of Microsoft Entra devices. Defaults to DevicesEntra.csv in the parent folder, then DevicesEntra.csv next to this script. The launcher adds Entra registration columns by matching computer names.

.PARAMETER AdInventoryCsv
Optional local CSV export of Active Directory computers. Defaults to DevicesAD.csv in the parent folder, then DevicesAD.csv next to this script. LOT wrappers pass a LOT-local DevicesAD.csv for domain-specific fallback refreshes.

.PARAMETER AdRootInventoryCsv
Optional root forest-wide AD CSV. LOT wrappers pass the toolkit-root DevicesAD.csv here so the launcher can prefer it when it exists and is less than 12 hours old.

.PARAMETER AdDomain
Optional AD domain/controller used when refreshing DevicesAD.csv. When omitted, the AD export targets all domains in the current AD forest. For LOT runs, set EHJIR_AD_DOMAIN or create AdDomain.txt in the LOT folder only when a domain-specific export is required.

.PARAMETER LogRoot
Local folder where PsExec per-computer logs are written. Defaults to PsExecLogs next to this script.

.PARAMETER ReportRoot
Local folder where cycle CSV/HTML summaries are written. Defaults to Reports next to this script.

.PARAMETER StaleCleanupDelaySeconds
Seconds passed to the repair script after stale local Intune cleanup before same-run auto-enrollment. Defaults to 60.

.PARAMETER DelayBetweenComputersSeconds
Seconds to wait between two job starts. Defaults to 0 for large fleet throughput.

.PARAMETER ThrottleLimit
Maximum number of computers processed in parallel. Defaults to 25.

.PARAMETER GlobalConcurrencyLimit
Maximum number of active computer workers shared by all LOT launchers in the same Windows
session. This limits local PowerShell worker jobs and their PsExec executions across multiple
LOT windows. Defaults to 15. Use 0 to disable the cross-LOT limiter.

.PARAMETER GlobalConcurrencySemaphoreName
Shared global gate name used by GlobalConcurrencyLimit. Use the same name across LOTs that
must share one limit. The launcher stores recoverable worker leases under the temp folder.

.PARAMETER GlobalConcurrencyLeaseTimeoutMinutes
Maximum age of a shared global worker lease before it is considered stale and cleaned.
Use 0 for automatic sizing based on PsExec and delayed evidence timeouts.

.PARAMETER JobPollSeconds
Seconds to wait between checks for completed parallel jobs. Defaults to 2.

.PARAMETER DelayBetweenCyclesMinutes
Minutes to wait after one full pass over the computer list before starting the next pass. Defaults to 1.

.PARAMETER DisableLotRunMutex
Disables the per-LOT run mutex that prevents the same LOT from being launched multiple times concurrently.

.PARAMETER DisableNightPause
Disables the default night pause that prevents a new cycle from starting between 20:00 and 07:00 local time.

.PARAMETER NightPauseStartHour
Local hour when the night pause starts. Defaults to 20.

.PARAMETER NightPauseEndHour
Local hour when the night pause ends. Defaults to 7.

.PARAMETER MaxCycles
Maximum number of cycles to run. 0 means infinite. Defaults to 0.

.PARAMETER RunOnce
Runs a single cycle and exits.

.PARAMETER DryRun
Checks DNS, ping and administrative share reachability for each computer without copying or executing the repair script.

.PARAMETER IgnoreRunGuardEveryCycle
Passes -IgnoreRunGuard on every loop cycle. Use carefully.

.PARAMETER CentralLogRoot
Local central folder where remote logs are collected after each computer run.
Defaults to CentralLogs next to this launcher. By default, files are stored under one Latest folder per computer.

.PARAMETER KeepCentralLogHistory
Keeps one central log folder per computer and per cycle. Without this switch, the outcome bucket and hashed computer `Latest` folder are overwritten each run.

.PARAMETER NoCentralLogCollection
Disables collection of C:\ProgramData\SmartM365\IntuneHybridJoinToolkit from each remote computer.

.PARAMETER CentralLogCollectionMode
Standard skips individual remote evidence files larger than 5 MB. Full collects every supported evidence file.

.PARAMETER UseTechnicianRunGuardHistory
Uses the shared technician-side run history to prevent overlapping or redundant launches of the same AD FQDN across LOTs.

.PARAMETER IgnoreTechnicianRunGuardHistory
Explicitly bypasses the shared technician-side run history.

.PARAMETER TechnicianRunGuardHours
Expiration window for technician-side run history entries. Defaults to 12 hours.

.PARAMETER RebootDelaySeconds
Seconds used by the remote script when scheduling a controlled reboot. Defaults to 180 so this launcher can pull logs through C$ before reboot.

.PARAMETER IntuneRetrySleepMinutes
Minutes passed to the repair script between local Intune enrollment re-checks after auto-enrollment is triggered. Defaults to 5.

.PARAMETER IntuneRetryMaxRetries
Number of local Intune enrollment re-checks passed to the repair script after auto-enrollment is triggered. Defaults to 5.

.PARAMETER RetryAfterRebootDelaySeconds
Seconds the endpoint startup task waits before resuming repair after a controlled reboot. Defaults to 300.

.PARAMETER RetryAfterRebootMaxAttempts
Maximum endpoint startup-task resume attempts before cleanup. Defaults to 3.

.PARAMETER PsExecTimeoutMinutes
Maximum time to wait for one PsExec execution before marking the computer as PSEXEC_TIMEOUT.
Use 0 to disable the timeout. Defaults to 120.

.PARAMETER CancellationDrainTimeoutMinutes
Maximum time to let active jobs finish after the first controlled stop request. A second stop request
forces the remaining local workers to stop. Use 0 to force immediately. Defaults to 15.

.PARAMETER CommunicationLostEvidenceWaitMinutes
Maximum minutes to wait before collecting remote evidence when PsExec lost communication with
PSEXESVC after starting remote PowerShell. Defaults to 65 so long local retry loops can write
their final CSV/LastRun.json before the launcher pulls logs.

.PARAMETER CommunicationLostEvidencePollMinutes
Minutes between current-run final CSV checks during the communication-lost wait window.
Defaults to 10.

.PARAMETER SkipPostCycleIntuneInventory
Do not refresh Intune inventory at the end of each cycle. By default, the launcher runs a LOT-scoped SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1 inventory and adds post-cycle Intune columns to the report.

.PARAMETER PostCycleIntuneInventoryPageSize
Graph page size used by SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1 for automatic LOT-scoped inventory refreshes. Defaults to 999.

.VERSION
2.10.78
#>

#requires -Version 5.1

[CmdletBinding(PositionalBinding=$false)]
param(
    [string]$ComputerListPath,
    [string]$PsExecPath,
    [int]$DelayBetweenComputersSeconds = 0,
    [int]$ThrottleLimit = 25,
    [int]$GlobalConcurrencyLimit = 15,
    [string]$GlobalConcurrencySemaphoreName = "Local\SmartM365_IntuneHybridJoinToolkit_ComputerWorkers",
    [int]$GlobalConcurrencyLeaseTimeoutMinutes = 0,
    [int]$JobPollSeconds = 2,
    [int]$DelayBetweenCyclesMinutes = 1,
    [switch]$DisableLotRunMutex,
    [switch]$DisableNightPause,
    [int]$NightPauseStartHour = 20,
    [int]$NightPauseEndHour = 7,
    [int]$MaxCycles = 0,
    [switch]$AllowDsregLeave = $true,
    [switch]$IgnoreRunGuard,
    [switch]$IgnoreRunGuardEveryCycle,
    [switch]$RunOnce,
    [switch]$DryRun,
    [switch]$AllowRebootWhenNoInteractiveUser,
    [switch]$AllowRebootAfterDsregLeave,
    [switch]$AllowRemoveNonIntuneMdmEnrollment,
    [switch]$AllowRemoveStaleIntuneEnrollment,
    [switch]$SkipVirtualMachines,
    [switch]$AuditOnly,
    [string]$IntuneInventoryCsv,
    [string]$IntuneInventoryNameColumn,
    [string]$EntraInventoryCsv,
    [string]$EntraInventoryNameColumn,
    [string]$AdInventoryCsv,
    [string]$AdRootInventoryCsv,
    [string]$AdInventoryNameColumn,
    [string]$AdDomain,
    [string]$LogRoot,
    [string]$ReportRoot,
    [string]$CentralLogRoot,
    [string]$ArchiveRoot,
    [switch]$KeepCentralLogHistory,
    [ValidateSet("Standard","Full")]
    [string]$CentralLogCollectionMode = "Standard",
    [switch]$SkipPreRunArchive,
    [switch]$ContinueOnDnsPreflightFailure,
    [switch]$NoCentralLogCollection,
    [switch]$UseTechnicianRunGuardHistory,
    [switch]$IgnoreTechnicianRunGuardHistory,
    [ValidateRange(0,168)]
    [int]$TechnicianRunGuardHours = 12,
    [int]$StaleCleanupDelaySeconds = 60,
    [int]$RebootDelaySeconds = 180,
    [int]$IntuneRetrySleepMinutes = 5,
    [int]$IntuneRetryMaxRetries = 5,
    [ValidateRange(0,3600)][int]$RetryAfterRebootDelaySeconds = 300,
    [ValidateRange(1,30)][int]$RetryAfterRebootMaxAttempts = 3,
    [int]$PsExecTimeoutMinutes = 120,
    [ValidateRange(0,1440)][int]$CancellationDrainTimeoutMinutes = 15,
    [int]$CommunicationLostEvidenceWaitMinutes = 65,
    [int]$CommunicationLostEvidencePollMinutes = 10,
    [switch]$SkipPostCycleIntuneInventory,
    [int]$PostCycleIntuneInventoryPageSize = 999,
    [Parameter(ValueFromRemainingArguments=$true)]
    [string[]]$UnexpectedArguments
)

$ErrorActionPreference = "Stop"
$LauncherVersion = "2.10.78"
$AdInventoryFreshnessHours = 12

if ($UnexpectedArguments -and $UnexpectedArguments.Count -gt 0) {
    throw ("Unexpected launcher argument(s): {0}. Pass PsExec with -PsExecPath <path>, not as a free argument." -f ($UnexpectedArguments -join " "))
}

if ($ThrottleLimit -lt 1) { $ThrottleLimit = 1 }
if ($GlobalConcurrencyLimit -lt 0) { $GlobalConcurrencyLimit = 0 }
if ([string]::IsNullOrWhiteSpace($GlobalConcurrencySemaphoreName)) { $GlobalConcurrencySemaphoreName = "Local\SmartM365_IntuneHybridJoinToolkit_ComputerWorkers" }
if ($GlobalConcurrencyLeaseTimeoutMinutes -lt 0) { $GlobalConcurrencyLeaseTimeoutMinutes = 0 }
if ($JobPollSeconds -lt 1) { $JobPollSeconds = 1 }
if ($DelayBetweenComputersSeconds -lt 0) { $DelayBetweenComputersSeconds = 0 }
if ($DelayBetweenCyclesMinutes -lt 0) { $DelayBetweenCyclesMinutes = 0 }
if ($NightPauseStartHour -lt 0) { $NightPauseStartHour = 0 }
if ($NightPauseStartHour -gt 23) { $NightPauseStartHour = 23 }
if ($NightPauseEndHour -lt 0) { $NightPauseEndHour = 0 }
if ($NightPauseEndHour -gt 23) { $NightPauseEndHour = 23 }
if ($RebootDelaySeconds -lt 60) { $RebootDelaySeconds = 60 }
if ($StaleCleanupDelaySeconds -lt 0) { $StaleCleanupDelaySeconds = 0 }
if ($IntuneRetrySleepMinutes -lt 1) { $IntuneRetrySleepMinutes = 1 }
if ($IntuneRetryMaxRetries -lt 1) { $IntuneRetryMaxRetries = 1 }
if ($PsExecTimeoutMinutes -lt 0) { $PsExecTimeoutMinutes = 0 }
if ($CommunicationLostEvidenceWaitMinutes -lt 0) { $CommunicationLostEvidenceWaitMinutes = 0 }
if ($CommunicationLostEvidencePollMinutes -lt 1) { $CommunicationLostEvidencePollMinutes = 1 }
if ($PostCycleIntuneInventoryPageSize -lt 1) { $PostCycleIntuneInventoryPageSize = 1 }
if ($PostCycleIntuneInventoryPageSize -gt 999) { $PostCycleIntuneInventoryPageSize = 999 }
if ($GlobalConcurrencyLeaseTimeoutMinutes -lt 1) {
    $timeoutBase = if ($PsExecTimeoutMinutes -gt 0) { $PsExecTimeoutMinutes } else { 240 }
    $GlobalConcurrencyLeaseTimeoutMinutes = [Math]::Max(30, $timeoutBase + $CommunicationLostEvidenceWaitMinutes + 30)
}
if ($DryRun -and -not $RunOnce -and $MaxCycles -eq 0) { $MaxCycles = 1 }

$BaseDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $PSScriptRoot
}
else {
    Split-Path -Parent $MyInvocation.MyCommand.Path
}

$ScriptName = "SmartM365-Invoke-IntuneHybridJoinRepair.ps1"
$LocalScriptPath = Join-Path $BaseDir $ScriptName
$ExportIntuneScriptPath = Join-Path $BaseDir "SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1"
$ExportEntraScriptPath = Join-Path $BaseDir "SmartM365-IntuneHybridJoinRepair-Export-EntraDevicesCsv.ps1"
$ExportAdScriptPath = Join-Path $BaseDir "SmartM365-IntuneHybridJoinRepair-Export-ADDevicesCsv.ps1"

if ([string]::IsNullOrWhiteSpace($ComputerListPath)) {
    $ComputerListPath = Join-Path $BaseDir "Computers.txt"
}

$IntuneInventoryCsvWasProvided = -not [string]::IsNullOrWhiteSpace($IntuneInventoryCsv)
if (-not $IntuneInventoryCsvWasProvided) {
    $candidateIntuneCsvPaths = @()

    $parentDir = Split-Path -Parent $BaseDir
    if (-not [string]::IsNullOrWhiteSpace($parentDir)) {
        $candidateIntuneCsvPaths += (Join-Path $parentDir "DevicesIntune.csv")
    }
    $candidateIntuneCsvPaths += (Join-Path $BaseDir "DevicesIntune.csv")

    foreach ($candidateIntuneCsvPath in $candidateIntuneCsvPaths) {
        if (Test-Path -LiteralPath $candidateIntuneCsvPath) {
            $IntuneInventoryCsv = $candidateIntuneCsvPath
            break
        }
    }

    if ([string]::IsNullOrWhiteSpace($IntuneInventoryCsv) -and $candidateIntuneCsvPaths.Count -gt 0) {
        $IntuneInventoryCsv = $candidateIntuneCsvPaths[0]
    }
}

$EntraInventoryCsvWasProvided = -not [string]::IsNullOrWhiteSpace($EntraInventoryCsv)
if (-not $EntraInventoryCsvWasProvided) {
    $candidateEntraCsvPaths = @()

    $parentDir = Split-Path -Parent $BaseDir
    if (-not [string]::IsNullOrWhiteSpace($parentDir)) {
        $candidateEntraCsvPaths += (Join-Path $parentDir "DevicesEntra.csv")
    }
    $candidateEntraCsvPaths += (Join-Path $BaseDir "DevicesEntra.csv")

    foreach ($candidateEntraCsvPath in $candidateEntraCsvPaths) {
        if (Test-Path -LiteralPath $candidateEntraCsvPath) {
            $EntraInventoryCsv = $candidateEntraCsvPath
            break
        }
    }

    if ([string]::IsNullOrWhiteSpace($EntraInventoryCsv) -and $candidateEntraCsvPaths.Count -gt 0) {
        $EntraInventoryCsv = $candidateEntraCsvPaths[0]
    }
}

$AdInventoryCsvWasProvided = -not [string]::IsNullOrWhiteSpace($AdInventoryCsv)
$AdRootInventoryCsvWasProvided = -not [string]::IsNullOrWhiteSpace($AdRootInventoryCsv)
$AdInventoryUsesRecentRootCsv = $false
if ($AdRootInventoryCsvWasProvided) {
    $adRootInventoryItem = Get-Item -LiteralPath $AdRootInventoryCsv -ErrorAction SilentlyContinue
    if ($adRootInventoryItem) {
        $adRootInventoryAge = (Get-Date) - $adRootInventoryItem.LastWriteTime
        if ($adRootInventoryAge.TotalHours -le $AdInventoryFreshnessHours) {
            $AdInventoryCsv = $adRootInventoryItem.FullName
            $AdInventoryUsesRecentRootCsv = $true
            $AdInventoryCsvWasProvided = $true
        }
    }
}
if (-not $AdInventoryCsvWasProvided) {
    $candidateAdCsvPaths = @()

    $parentDir = Split-Path -Parent $BaseDir
    if (-not [string]::IsNullOrWhiteSpace($parentDir)) {
        $candidateAdCsvPaths += (Join-Path $parentDir "DevicesAD.csv")
    }
    $candidateAdCsvPaths += (Join-Path $BaseDir "DevicesAD.csv")

    foreach ($candidateAdCsvPath in $candidateAdCsvPaths) {
        if (Test-Path -LiteralPath $candidateAdCsvPath) {
            $AdInventoryCsv = $candidateAdCsvPath
            break
        }
    }

    if ([string]::IsNullOrWhiteSpace($AdInventoryCsv) -and $candidateAdCsvPaths.Count -gt 0) {
        $AdInventoryCsv = $candidateAdCsvPaths[0]
    }
}

$LogRootWasProvided = -not [string]::IsNullOrWhiteSpace($LogRoot)
$ReportRootWasProvided = -not [string]::IsNullOrWhiteSpace($ReportRoot)
$CentralLogRootWasProvided = -not [string]::IsNullOrWhiteSpace($CentralLogRoot)
$ArchiveRootWasProvided = -not [string]::IsNullOrWhiteSpace($ArchiveRoot)

if ([string]::IsNullOrWhiteSpace($PsExecPath)) {
    $PsExecPath = Join-Path $BaseDir "PsExec.exe"
}

if ([string]::IsNullOrWhiteSpace($LogRoot)) {
    $LogRoot = Join-Path $BaseDir "PsExecLogs"
}
if ([string]::IsNullOrWhiteSpace($ReportRoot)) {
    $ReportRoot = Join-Path $BaseDir "Reports"
}
$CollectRemoteLogs = -not [bool]$NoCentralLogCollection
if ([string]::IsNullOrWhiteSpace($CentralLogRoot)) {
    $CentralLogRoot = Join-Path $BaseDir "CentralLogs"
}

$ComputerListPath = [System.IO.Path]::GetFullPath($ComputerListPath)
$LotRoot = Split-Path -Parent $ComputerListPath
$launcherRunTimestamp = Get-Date -Format "yyyyMMdd-HHmmss"
$defaultRunRoot = $null
$lotParent = Split-Path -Parent $LotRoot
if ((Split-Path -Leaf $lotParent) -eq "Lots") {
    $toolkitRootFromLot = Split-Path -Parent $lotParent
    $lotNameFromPath = Split-Path -Leaf $LotRoot
    $defaultRunRoot = Join-Path (Join-Path (Join-Path $toolkitRootFromLot "Runs") $lotNameFromPath) $launcherRunTimestamp
}
if ($defaultRunRoot) {
    if (-not $LogRootWasProvided) { $LogRoot = Join-Path $defaultRunRoot "PsExecLogs" }
    if (-not $ReportRootWasProvided) { $ReportRoot = Join-Path $defaultRunRoot "Reports" }
    if (-not $CentralLogRootWasProvided) { $CentralLogRoot = Join-Path $defaultRunRoot "CentralLogs" }
    if (-not $ArchiveRootWasProvided) { $ArchiveRoot = Join-Path $defaultRunRoot "Archives" }
}

$LocalScriptPath = [System.IO.Path]::GetFullPath($LocalScriptPath)
if (-not [string]::IsNullOrWhiteSpace($IntuneInventoryCsv)) { $IntuneInventoryCsv = [System.IO.Path]::GetFullPath($IntuneInventoryCsv) }
if (-not [string]::IsNullOrWhiteSpace($EntraInventoryCsv)) { $EntraInventoryCsv = [System.IO.Path]::GetFullPath($EntraInventoryCsv) }
if (-not [string]::IsNullOrWhiteSpace($AdInventoryCsv)) { $AdInventoryCsv = [System.IO.Path]::GetFullPath($AdInventoryCsv) }
$LogRoot = [System.IO.Path]::GetFullPath($LogRoot)
$ReportRoot = [System.IO.Path]::GetFullPath($ReportRoot)
$CentralLogRoot = [System.IO.Path]::GetFullPath($CentralLogRoot)
if ([string]::IsNullOrWhiteSpace($ArchiveRoot)) {
    $ArchiveRoot = Join-Path $LotRoot "Archives"
}
else {
    $ArchiveRoot = [System.IO.Path]::GetFullPath($ArchiveRoot)
}
$script:ReportPathBase = $LotRoot

$RemoteRelativeDir = "ProgramData\SmartM365\IntuneHybridJoinToolkit"
$RemoteScriptPath = "C:\ProgramData\SmartM365\IntuneHybridJoinToolkit\$ScriptName"
$RemoteDataRelativeDir = "ProgramData\SmartM365\IntuneHybridJoinToolkit"

if (-not (Test-Path -LiteralPath $LocalScriptPath)) {
    throw "Repair script not found: $LocalScriptPath"
}

if (-not (Test-Path -LiteralPath $ComputerListPath)) {
    throw "Computer list not found: $ComputerListPath"
}

function Test-SingleComputerLaunch {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string]$RootPath
    )

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    $fullRoot = [System.IO.Path]::GetFullPath($RootPath)
    if ($fullPath -match '\\SingleComputer\\' -or $fullRoot -match '\\SingleComputer\\') { return $true }

    $computers = @(
        Get-Content -LiteralPath $fullPath -ErrorAction Stop |
            ForEach-Object { ([string]$_).Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.StartsWith('#') } |
            Select-Object -Unique
    )
    return ($computers.Count -eq 1)
}

$script:IsSingleComputerLaunch = Test-SingleComputerLaunch -Path $ComputerListPath -RootPath $LotRoot
if ($script:IsSingleComputerLaunch) {
    $ThrottleLimit = 1
    $GlobalConcurrencyLimit = 0
}

function New-LotRunMutexName {
    param([Parameter(Mandatory=$true)][string]$LotPath)

    $normalizedPath = ([System.IO.Path]::GetFullPath($LotPath)).TrimEnd('\').ToUpperInvariant()
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($normalizedPath)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = ([System.BitConverter]::ToString($sha.ComputeHash($bytes)) -replace '-', '')
    }
    finally {
        $sha.Dispose()
    }

    return "Local\SmartM365_IntuneHybridJoinToolkit_LotRun_$($hash.Substring(0, 32))"
}

function Acquire-LotRunMutex {
    param(
        [Parameter(Mandatory=$true)][string]$MutexName,
        [Parameter(Mandatory=$true)][string]$LotPath
    )

    $createdNew = $false
    $mutex = New-Object System.Threading.Mutex($true, $MutexName, [ref]$createdNew)
    if (-not $createdNew) {
        $mutex.Dispose()
        throw ("This LOT is already running in another launcher process. LOT={0}; Mutex={1}" -f $LotPath,$MutexName)
    }

    return $mutex
}

$LauncherStartupLogRoot = Join-Path (Split-Path -Parent $LogRoot) "Logs"
if (-not (Test-Path -LiteralPath $LauncherStartupLogRoot -PathType Container)) {
    New-Item -ItemType Directory -Path $LauncherStartupLogRoot -Force | Out-Null
}
$launcherLogSafeLot = (Split-Path -Leaf $LotRoot) -replace '[\/:*?"<>|]','_'
$LauncherStartupLogPath = Join-Path $LauncherStartupLogRoot ("SmartM365-IHJ-Launcher_{0}_{1}.log" -f $launcherLogSafeLot,$launcherRunTimestamp)
$LauncherLatestLogPath = Join-Path $LauncherStartupLogRoot "SmartM365-IHJ-Launcher_latest.log"
Set-Content -LiteralPath $LauncherStartupLogPath -Value "" -Encoding UTF8
Set-Content -LiteralPath $LauncherLatestLogPath -Value "" -Encoding UTF8
$script:LauncherLogPath = $LauncherStartupLogPath
$script:LauncherLatestLogPath = $LauncherLatestLogPath

function Write-LauncherStartupLine {
    param([string]$Message)

    Write-Host $Message
}

function Write-Host {
    param(
        [Parameter(Position = 0)][object]$Object = '',
        [switch]$NoNewline,
        [System.ConsoleColor]$ForegroundColor,
        [System.ConsoleColor]$BackgroundColor
    )
    $ts = if ($null -ne $Object -and '' -ne [string]$Object) { "[{0}] " -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') } else { '' }
    $p = @{ Object = ($ts + [string]$Object) }
    if ($PSBoundParameters.ContainsKey('ForegroundColor')) { $p['ForegroundColor'] = $ForegroundColor }
    if ($PSBoundParameters.ContainsKey('BackgroundColor')) { $p['BackgroundColor'] = $BackgroundColor }
    if ($PSBoundParameters.ContainsKey('NoNewline')) { $p['NoNewline'] = $NoNewline }
    Microsoft.PowerShell.Utility\Write-Host @p
    if (-not [string]::IsNullOrWhiteSpace([string]$Object)) {
        try {
            $logLine = "{0:yyyy-MM-dd HH:mm:ss} {1}" -f (Get-Date),([string]$Object)
            Add-Content -LiteralPath $script:LauncherLogPath -Value $logLine -Encoding UTF8
            Add-Content -LiteralPath $script:LauncherLatestLogPath -Value $logLine -Encoding UTF8
        }
        catch {}
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
    param([ValidateRange(0,86400)][int]$Seconds)

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
    param([Parameter(Mandatory=$true)][string]$Status,[string]$ReportPath = '')

    if ([string]::IsNullOrWhiteSpace($script:ActiveLotRunStatePath)) { return }
    [pscustomobject]@{
        Version = 1
        Toolkit = 'IntuneHybridJoinToolkit'
        Status = $Status
        ProcessId = $PID
        LotName = (Split-Path -Leaf $LotRoot)
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

function New-HybridJoinCancellationResult {
    param(
        [Parameter(Mandatory=$true)][string]$ComputerName,
        [Parameter(Mandatory=$true)][int]$CycleNumber,
        [Parameter(Mandatory=$true)][string]$Status,
        [Parameter(Mandatory=$true)][string]$Detail,
        [string]$NextAction = 'VERIFY_REMOTE_STATE_BEFORE_RELAUNCH'
    )

    $row = [ordered]@{}
    foreach ($column in @(Get-LauncherReportColumns)) { $row[$column] = '' }
    $row['LauncherVersion'] = $LauncherVersion
    $row['Cycle'] = $CycleNumber
    $row['Computer'] = $ComputerName
    $row['Timestamp'] = Get-Date
    $row['DryRun'] = [bool]$DryRun
    $row['Status'] = $Status
    $row['EffectiveStatus'] = $Status
    $row['NextAction'] = $NextAction
    $row['EffectiveNextAction'] = $NextAction
    $row['RemoteNextAction'] = $NextAction
    $row['RemoteDetail'] = $Detail
    $row['LogPath'] = $script:LauncherLogPath
    $row['ErrorMessage'] = ''
    return [pscustomobject]$row
}

function Get-LocalWorkerStartDiagnosticText {
    param([string]$ErrorMessage)

    $processPath = try { [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName } catch { '' }
    $processHandleCount = try { [System.Diagnostics.Process]::GetCurrentProcess().HandleCount } catch { -1 }
    $expectedExecutable = if ($PSVersionTable.PSEdition -eq 'Core') {
        Join-Path $PSHOME 'pwsh.exe'
    }
    else {
        Join-Path $PSHOME 'powershell.exe'
    }
    $fileExists = try { [System.IO.File]::Exists($expectedExecutable) } catch { $false }
    $testPath = try { Test-Path -LiteralPath $expectedExecutable -PathType Leaf -ErrorAction Stop } catch { $false }
    $fileAccess = try {
        $item = Get-Item -LiteralPath $expectedExecutable -ErrorAction Stop
        'Readable; Length={0}; LastWriteTimeUtc={1:o}' -f $item.Length,$item.LastWriteTimeUtc
    }
    catch {
        'Unavailable; Error={0}' -f $_.Exception.Message
    }

    return ('Error={0}; PSEdition={1}; PSVersion={2}; PSHOME={3}; ExpectedExecutable={4}; FileExists={5}; TestPath={6}; FileAccess={7}; ProcessPath={8}; ProcessHandleCount={9}' -f $ErrorMessage,$PSVersionTable.PSEdition,$PSVersionTable.PSVersion,$PSHOME,$expectedExecutable,$fileExists,$testPath,$fileAccess,$processPath,$processHandleCount)
}

function Start-LocalWorkerJobWithRetry {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$StartOperation,
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][string]$JobName,
        [ValidateRange(1,10)][int]$MaxAttempts = 3,
        [ValidateRange(0,60)][int]$RetryDelaySeconds = 3
    )

    $attemptDiagnostics = New-Object System.Collections.Generic.List[string]
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            $job = & $StartOperation
            if ($null -eq $job) { throw 'Start-Job returned no job object.' }
            if ($attempt -gt 1) {
                Write-Host ("LOCAL_WORKER_START_RECOVERED: Computer={0}; Job={1}; Attempt={2}/{3}." -f $ComputerName,$JobName,$attempt,$MaxAttempts) -ForegroundColor Green
            }
            return [pscustomobject]@{
                Succeeded = $true
                Job = $job
                Detail = ($attemptDiagnostics -join ' | ')
            }
        }
        catch {
            $diagnostic = Get-LocalWorkerStartDiagnosticText -ErrorMessage $_.Exception.Message
            $attemptDiagnostics.Add(("Attempt={0}/{1}; {2}" -f $attempt,$MaxAttempts,$diagnostic))
            Write-Host ("LOCAL_WORKER_START_RETRY: Computer={0}; Job={1}; Attempt={2}/{3}; {4}" -f $ComputerName,$JobName,$attempt,$MaxAttempts,$diagnostic) -ForegroundColor Yellow
            if ($attempt -lt $MaxAttempts -and $RetryDelaySeconds -gt 0) {
                Start-Sleep -Seconds $RetryDelaySeconds
            }
        }
    }

    return [pscustomobject]@{
        Succeeded = $false
        Job = $null
        Detail = ($attemptDiagnostics -join ' | ')
    }
}

function Get-ScriptHeaderVersionQuick {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return "missing"
    }

    try {
        $content = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
        $match = [regex]::Match($content, '(?m)^\.VERSION\s*\r?\n\s*([^\r\n]+)')
        if ($match.Success) { return $match.Groups[1].Value.Trim() }

        $scriptVersionMatch = [regex]::Match($content, '(?m)^\s*\$ScriptVersion\s*=\s*"([^"\r\n]+)"')
        if ($scriptVersionMatch.Success) { return $scriptVersionMatch.Groups[1].Value.Trim() }
    }
    catch {}

    return "unknown"
}

function Write-LauncherStartupInfo {
    $repairScriptVersion = Get-ScriptHeaderVersionQuick -Path $LocalScriptPath
    $powerShellProcessPath = try { [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName } catch { '' }
    Write-Host ""
    Write-Host "Remote repair launcher startup" -ForegroundColor Cyan
    Write-LauncherStartupLine ("Launcher ver : {0}" -f $LauncherVersion)
    Write-LauncherStartupLine ("PowerShell   : Edition={0}; Version={1}; Process={2}" -f $PSVersionTable.PSEdition,$PSVersionTable.PSVersion,$powerShellProcessPath)
    Write-LauncherStartupLine ("Started      : {0:yyyy-MM-dd HH:mm:ss}; Host={1}; User={2}; PID={3}" -f (Get-Date),$env:COMPUTERNAME,$env:USERNAME,$PID)
    Write-LauncherStartupLine ("LOT root     : {0}" -f $LotRoot)
    Write-LauncherStartupLine ("Script       : {0}" -f $LocalScriptPath)
    Write-LauncherStartupLine ("Script ver   : {0}" -f $repairScriptVersion)
    Write-LauncherStartupLine ("Computers    : {0}" -f $ComputerListPath)
    Write-LauncherStartupLine ("PsExec       : {0}" -f $PsExecPath)
    Write-LauncherStartupLine ("Dry run      : {0}" -f ([bool]$DryRun))
    Write-LauncherStartupLine ("Loop         : {0}; Delay={1} minute(s); MaxCycles={2}" -f (-not [bool]$RunOnce),$DelayBetweenCyclesMinutes,$MaxCycles)
    Write-LauncherStartupLine ("Night pause  : Enabled={0}; Window={1}:00-{2}:00 local time" -f (-not [bool]$DisableNightPause),$NightPauseStartHour,$NightPauseEndHour)
    Write-LauncherStartupLine ("Parallelism  : ThrottleLimit={0}; GlobalConcurrencyLimit={1}; GlobalLeaseTimeout={2} minute(s)" -f $ThrottleLimit,$GlobalConcurrencyLimit,$GlobalConcurrencyLeaseTimeoutMinutes)
    if ($script:IsSingleComputerLaunch) { Write-LauncherStartupLine "Single PC    : worker limits ignored for one-computer launch." }
    Write-LauncherStartupLine ("Archive      : Enabled={0}; Root={1}" -f (-not [bool]$SkipPreRunArchive),$ArchiveRoot)
    Write-LauncherStartupLine ("Logs         : PsExec={0}; Reports={1}; Central={2}" -f $LogRoot,$ReportRoot,$CentralLogRoot)
    Write-LauncherStartupLine ("Inventory    : Intune={0}; Entra={1}; AD={2}; ADRoot={3}; ADDomain={4}" -f $IntuneInventoryCsv,$EntraInventoryCsv,$AdInventoryCsv,$AdRootInventoryCsv,$AdDomain)
    Write-LauncherStartupLine ("Startup log  : {0}" -f $LauncherStartupLogPath)
    Write-Host ""
}

Write-LauncherStartupInfo
$script:LotRunMutex = $null
$script:LotRunMutexName = $null
if (-not $DisableLotRunMutex) {
    $script:LotRunMutexName = New-LotRunMutexName -LotPath $LotRoot
    $script:LotRunMutex = Acquire-LotRunMutex -MutexName $script:LotRunMutexName -LotPath $LotRoot
    Write-Host ("LOT run lock: acquired {0}" -f $script:LotRunMutexName) -ForegroundColor DarkGray
}
else {
    Write-Host "LOT run lock: disabled by -DisableLotRunMutex." -ForegroundColor Yellow
}

function Invoke-PreRunLotArchive {
    param(
        [Parameter(Mandatory=$true)][string[]]$Paths,
        [Parameter(Mandatory=$true)][string]$DestinationRoot,
        [Parameter(Mandatory=$true)][string]$Prefix
    )

    $existingPaths = @(
        foreach ($path in $Paths) {
            if ([string]::IsNullOrWhiteSpace($path)) { continue }
            if (-not (Test-Path -LiteralPath $path)) { continue }
            $children = @(Get-ChildItem -LiteralPath $path -Force -ErrorAction SilentlyContinue | Select-Object -First 1)
            if ($children.Count -gt 0) { $path }
        }
    )

    if ($existingPaths.Count -eq 0) {
        Write-Host "Pre-run archive: no existing LOT output folders to archive." -ForegroundColor DarkGray
        return $null
    }

    if (-not (Test-Path -LiteralPath $DestinationRoot)) {
        New-Item -ItemType Directory -Path $DestinationRoot -Force | Out-Null
    }

    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $zipPath = Join-Path $DestinationRoot ("{0}_{1}.zip" -f $Prefix,$timestamp)
    $skippedFiles = New-Object 'System.Collections.Generic.List[string]'
    $archivedFileCount = 0

    Write-Host ("Pre-run archive: compressing previous LOT outputs to {0}" -f $zipPath) -ForegroundColor Cyan

    $zipStream = [System.IO.File]::Open($zipPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
    try {
        $zipArchive = [System.IO.Compression.ZipArchive]::new($zipStream, [System.IO.Compression.ZipArchiveMode]::Create, $false)
        try {
            foreach ($path in $existingPaths) {
                $rootItem = Get-Item -LiteralPath $path -ErrorAction SilentlyContinue
                if (-not $rootItem) { continue }

                $rootFullName = $rootItem.FullName.TrimEnd('\', '/')
                $entryRoot = Split-Path -Leaf $rootFullName
                $files = @(Get-ChildItem -LiteralPath $rootFullName -Recurse -Force -File -ErrorAction SilentlyContinue)
                foreach ($file in $files) {
                    $relativeName = $file.FullName.Substring($rootFullName.Length).TrimStart('\', '/') -replace '\\', '/'
                    if ([string]::IsNullOrWhiteSpace($relativeName)) { continue }

                    $entryName = ('{0}/{1}' -f $entryRoot,$relativeName)
                    try {
                        [System.IO.Compression.ZipFileExtensions]::CreateEntryFromFile(
                            $zipArchive,
                            $file.FullName,
                            $entryName,
                            [System.IO.Compression.CompressionLevel]::Optimal
                        ) | Out-Null
                        $archivedFileCount++
                    }
                    catch [System.IO.FileNotFoundException] {
                        if ($skippedFiles.Count -lt 10) { [void]$skippedFiles.Add($file.FullName) }
                    }
                    catch [System.IO.DirectoryNotFoundException] {
                        if ($skippedFiles.Count -lt 10) { [void]$skippedFiles.Add($file.FullName) }
                    }
                    catch [System.IO.PathTooLongException] {
                        if ($skippedFiles.Count -lt 10) { [void]$skippedFiles.Add(("{0} ({1})" -f $file.FullName,$_.Exception.Message)) }
                    }
                    catch [System.IO.IOException] {
                        if ($skippedFiles.Count -lt 10) { [void]$skippedFiles.Add(("{0} ({1})" -f $file.FullName,$_.Exception.Message)) }
                    }
                    catch [System.UnauthorizedAccessException] {
                        if ($skippedFiles.Count -lt 10) { [void]$skippedFiles.Add(("{0} ({1})" -f $file.FullName,$_.Exception.Message)) }
                    }
                }
            }
        }
        finally {
            $zipArchive.Dispose()
        }
    }
    finally {
        $zipStream.Dispose()
    }

    foreach ($path in $existingPaths) {
        try {
            Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction Stop
        }
        catch {
            Write-Host ("Pre-run archive warning: cleanup could not fully remove {0}: {1}" -f $path,$_.Exception.Message) -ForegroundColor Yellow
            Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    if ($skippedFiles.Count -gt 0) {
        Write-Host ("Pre-run archive warning: skipped {0} volatile or unreadable file(s) while creating the archive." -f $skippedFiles.Count) -ForegroundColor Yellow
        foreach ($skippedFile in $skippedFiles) {
            Write-Host ("  skipped: {0}" -f $skippedFile) -ForegroundColor DarkYellow
        }
    }

    Write-Host ("Pre-run archive: archived {0} file(s) and cleaned {1} folder(s)." -f $archivedFileCount,$existingPaths.Count) -ForegroundColor Green
    return $zipPath
}
if (-not $SkipPreRunArchive) {
    $archivePaths = @($CentralLogRoot, $LogRoot, $ReportRoot)
    try {
        $archivePath = Invoke-PreRunLotArchive -Paths $archivePaths -DestinationRoot $ArchiveRoot -Prefix "IntuneHybridJoinToolkit_PreRun"
        if (-not [string]::IsNullOrWhiteSpace($archivePath)) {
            Write-Host ("Previous LOT outputs archived to: {0}" -f $archivePath) -ForegroundColor Green
        }
    }
    catch {
        throw ("Pre-run archive failed; existing LOT outputs were not cleaned. Error={0}" -f $_.Exception.Message)
    }
}
else {
    Write-Host "Pre-run archive skipped by -SkipPreRunArchive." -ForegroundColor Yellow
}

function Resolve-PsExecPath {
    param(
        [Parameter(Mandatory=$true)][string]$Path
    )

    $candidatePath = $Path.Trim().Trim([char]34)
    $embeddedPsExecPaths = @(
        [regex]::Matches($candidatePath, '(?i)[A-Z]:\\[^\r\n"]*?PsExec(?:64)?\.exe') |
            ForEach-Object { $_.Value.Trim().Trim([char]34) } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
    )

    if ($embeddedPsExecPaths.Count -gt 1) {
        throw ("Invalid PsExecPath value: {0}. Multiple PsExec paths were provided: {1}. Provide exactly one executable path, for example -PsExecPath ""C:\Sysinternals\PsExec.exe""." -f $Path, ($embeddedPsExecPaths -join " | "))
    }

    if ($embeddedPsExecPaths.Count -eq 1 -and $candidatePath -ne $embeddedPsExecPaths[0]) {
        $candidatePath = $embeddedPsExecPaths[0]
    }

    if (-not [string]::IsNullOrWhiteSpace($candidatePath) -and (Test-Path -LiteralPath $candidatePath -PathType Leaf)) {
        return (Get-Item -LiteralPath $candidatePath -ErrorAction Stop).FullName
    }

    if ($candidatePath -match '(?i)\.exe["'']?\s+\S') {
        throw ("Invalid PsExecPath value: {0}. Provide exactly one executable path, for example -PsExecPath ""C:\Sysinternals\PsExec.exe""." -f $Path)
    }

    if (-not [string]::IsNullOrWhiteSpace($candidatePath)) {
        $command = Get-Command -Name $candidatePath -CommandType Application -ErrorAction SilentlyContinue
        if ($command) {
            return $command.Source
        }
    }

    $pathCommand = Get-Command -Name "PsExec.exe" -CommandType Application -ErrorAction SilentlyContinue
    if ($pathCommand) {
        return $pathCommand.Source
    }

    throw ("PsExec.exe not found. Place PsExec.exe next to this launcher ({0}) or add PsExec.exe to PATH. You can also pass -PsExecPath <path>." -f $BaseDir)
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

if (-not $DryRun) {
    $PsExecPath = Resolve-PsExecPath -Path $PsExecPath
}
$psExecSecurityPath = if ($DryRun) { '' } else { $PsExecPath }
$script:PsExecSecurityEvidence = Get-PsExecSecurityEvidence -Path $psExecSecurityPath
if (-not $DryRun) { Assert-PsExecSecurityEvidence -Evidence $script:PsExecSecurityEvidence }
$script:PsExecSecurityEvidenceRows = @(ConvertTo-PsExecSecurityEvidenceRows -Evidence $script:PsExecSecurityEvidence)
Write-LauncherStartupLine ("PsExec sec. : Status={0}; SHA256={1}; Signature={2}; Signer={3}; Version={4}" -f $script:PsExecSecurityEvidence.CheckStatus,$script:PsExecSecurityEvidence.SHA256,$script:PsExecSecurityEvidence.SignatureStatus,$script:PsExecSecurityEvidence.SignerSubject,$script:PsExecSecurityEvidence.FileVersion)

if (-not (Test-Path -LiteralPath $LogRoot)) {
    New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null
}
if (-not (Test-Path -LiteralPath $ReportRoot)) {
    New-Item -ItemType Directory -Path $ReportRoot -Force | Out-Null
}
$script:CancellationRunStartedUtc = (Get-Date).ToUniversalTime()
$script:CancellationStateRoot = Join-Path (Split-Path -Parent $ReportRoot) 'State'
New-Item -ItemType Directory -Path $script:CancellationStateRoot -Force | Out-Null
$script:CancellationSignalPath = Join-Path $script:CancellationStateRoot ("StopRequested_{0}.json" -f $PID)
$script:ActiveLotRunStatePath = Join-Path $script:CancellationStateRoot ("ActiveLotRun_{0}.json" -f $PID)
Remove-Item -LiteralPath $script:CancellationSignalPath -Force -ErrorAction SilentlyContinue
Initialize-LotCancellationSupport
Set-ActiveLotRunState -Status 'Starting'
Write-Host ("Controlled stop: first Ctrl+C stops new starts and drains active jobs for up to {0} minute(s); second Ctrl+C forces local job stop." -f $CancellationDrainTimeoutMinutes) -ForegroundColor DarkCyan


if ($CollectRemoteLogs -and -not (Test-Path -LiteralPath $CentralLogRoot)) {
    New-Item -ItemType Directory -Path $CentralLogRoot -Force | Out-Null
}

function Get-IntuneInventorySet {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$false)][string]$NameColumn
    )

    $set = @{}
    if ([string]::IsNullOrWhiteSpace($Path)) { return $set }
    if (-not (Test-Path -LiteralPath $Path)) { throw "Intune inventory CSV not found: $Path" }

    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { return $set }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        $candidateColumns = @("ComputerName","computerName","DeviceName","deviceName","Name","name","managedDeviceName","ManagedDeviceName")
        $first = $rows | Select-Object -First 1
        foreach ($candidate in $candidateColumns) {
            if ($first.PSObject.Properties.Name -contains $candidate) {
                $NameColumn = $candidate
                break
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        throw "Unable to infer the Intune inventory device name column. Use -IntuneInventoryNameColumn."
    }

    foreach ($row in $rows) {
        if ($row.PSObject.Properties.Name -contains "IntuneInventoryPresent") {
            $presentValue = ([string]$row.IntuneInventoryPresent).Trim()
            if ($presentValue -notin @("True","true","1","YES","Yes","yes","OUI","Oui","oui")) {
                continue
            }
        }

        $value = [string]$row.$NameColumn
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $short = ($value.Trim().Split(".")[0]).ToUpperInvariant()
            if (-not $set.ContainsKey($short)) { $set[$short] = $true }
        }
    }
    if ($set.Count -eq 0 -and $rows.Count -gt 0) {
        $set["__SMARTM365_INVENTORY_CHECKED__"] = $true
    }

    return $set
}

function Get-EntraInventoryMap {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$false)][string]$NameColumn
    )

    $map = @{}
    if ([string]::IsNullOrWhiteSpace($Path)) { return $map }
    if (-not (Test-Path -LiteralPath $Path)) { return $map }

    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { return $map }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        $candidateColumns = @("ComputerName","computerName","DisplayName","displayName","DeviceName","deviceName","Name","name")
        $first = $rows | Select-Object -First 1
        foreach ($candidate in $candidateColumns) {
            if ($first.PSObject.Properties.Name -contains $candidate) {
                $NameColumn = $candidate
                break
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        throw "Unable to infer the Entra inventory device name column. Use -EntraInventoryNameColumn."
    }

    foreach ($row in $rows) {
        $value = [string]$row.$NameColumn
        if ([string]::IsNullOrWhiteSpace($value)) { continue }

        $short = ($value.Trim().Split(".")[0]).ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($short)) { continue }

        $present = $true
        if ($row.PSObject.Properties.Name -contains "EntraInventoryPresent") {
            $present = Test-BooleanLikeTrue -Value $row.EntraInventoryPresent
        }
        if (-not $present) { continue }

        if (-not $map.ContainsKey($short)) {
            $map[$short] = $row
        }
    }
    if ($map.Count -eq 0 -and $rows.Count -gt 0) {
        $map["__SMARTM365_INVENTORY_CHECKED__"] = [pscustomobject]@{}
    }

    return $map
}

function Get-AdInventoryMap {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$false)][string]$NameColumn
    )

    $map = @{}
    if ([string]::IsNullOrWhiteSpace($Path)) { return $map }
    if (-not (Test-Path -LiteralPath $Path)) { return $map }

    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { return $map }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        $candidateColumns = @("ComputerName","computerName","DNSHostName","dnsHostName","Name","name")
        $first = $rows | Select-Object -First 1
        foreach ($candidate in $candidateColumns) {
            if ($first.PSObject.Properties.Name -contains $candidate) {
                $NameColumn = $candidate
                break
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        throw "Unable to infer the AD inventory device name column. Use -AdInventoryNameColumn."
    }

    foreach ($row in $rows) {
        $value = [string]$row.$NameColumn
        if ([string]::IsNullOrWhiteSpace($value)) { continue }

        $short = ($value.Trim().Split(".")[0]).ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($short)) { continue }

        $present = $true
        if ($row.PSObject.Properties.Name -contains "ADInventoryPresent") {
            $present = Test-BooleanLikeTrue -Value $row.ADInventoryPresent
        }
        if (-not $present) { continue }

        if (-not $map.ContainsKey($short)) {
            $map[$short] = $row
        }
    }
    if ($map.Count -eq 0 -and $rows.Count -gt 0) {
        $map["__SMARTM365_INVENTORY_CHECKED__"] = [pscustomobject]@{}
    }

    return $map
}

function ConvertTo-PortableLotPath {
    param([Parameter(Mandatory=$false)]$Value)

    if ($null -eq $Value) { return $Value }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $text }
    if ([string]::IsNullOrWhiteSpace($script:ReportPathBase)) { return $text }

    try {
        $base = [System.IO.Path]::GetFullPath($script:ReportPathBase).TrimEnd('\')
        $candidate = if ([System.IO.Path]::IsPathRooted($text)) {
            [System.IO.Path]::GetFullPath($text).TrimEnd('\')
        }
        else {
            [System.IO.Path]::GetFullPath((Join-Path $base $text)).TrimEnd('\')
        }

        if ($candidate.Equals($base, [System.StringComparison]::OrdinalIgnoreCase)) {
            return "."
        }

        $prefix = $base + "\"
        if ($candidate.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            return ".\" + $candidate.Substring($prefix.Length)
        }
    }
    catch {
        return $text
    }

    return $text
}

function ConvertTo-PortableReportRow {
    param([Parameter(Mandatory=$true)]$Row)

    $pathColumns = @(
        "LogPath",
        "RemoteLogsPath",
        "RemoteCurrentRunLogsPath",
        "PostCycleIntuneInventoryCsv",
        "PostCycleEntraInventoryCsv",
        "PostCycleADInventoryCsv"
    )

    $copy = [ordered]@{}
    foreach ($property in $Row.PSObject.Properties) {
        $value = $property.Value
        if ($pathColumns -contains $property.Name) {
            $value = ConvertTo-PortableLotPath -Value $value
        }
        $copy[$property.Name] = $value
    }

    return [PSCustomObject]$copy
}

function Get-PortableReportRows {
    param([Parameter(Mandatory=$true)][AllowEmptyCollection()][object[]]$Rows)

    return @($Rows | ForEach-Object { ConvertTo-PortableReportRow -Row $_ })
}

function Test-BooleanLikeTrue {
    param([AllowNull()][object]$Value)

    if ($Value -eq $true) { return $true }
    $text = ([string]$Value).Trim()
    return ($text -in @("True","true","1","YES","Yes","yes","OUI","Oui","oui"))
}

function Get-NightPauseResumeTime {
    param(
        [Parameter(Mandatory=$true)][datetime]$Now,
        [Parameter(Mandatory=$true)][int]$StartHour,
        [Parameter(Mandatory=$true)][int]$EndHour
    )

    if ($StartHour -eq $EndHour) { return $null }

    $start = New-TimeSpan -Hours $StartHour
    $end = New-TimeSpan -Hours $EndHour
    $current = $Now.TimeOfDay

    if ($StartHour -lt $EndHour) {
        if ($current -ge $start -and $current -lt $end) {
            return $Now.Date.Add($end)
        }
        return $null
    }

    if ($current -ge $start) {
        return $Now.Date.AddDays(1).Add($end)
    }

    if ($current -lt $end) {
        return $Now.Date.Add($end)
    }

    return $null
}

function Wait-OutsideNightPauseWindow {
    param([Parameter(Mandatory=$true)][int]$NextCycleNumber)

    if ($DisableNightPause) { return }

    while ($true) {
        $now = Get-Date
        $resumeAt = Get-NightPauseResumeTime -Now $now -StartHour $NightPauseStartHour -EndHour $NightPauseEndHour
        if ($null -eq $resumeAt) { return }

        $secondsRemaining = [int][Math]::Ceiling(($resumeAt - $now).TotalSeconds)
        if ($secondsRemaining -lt 1) { return }

        Write-Host ("Night cycle pause active. Next cycle {0} will start after {1}. Press Ctrl+C to stop." -f $NextCycleNumber,$resumeAt.ToString("yyyy-MM-dd HH:mm:ss")) -ForegroundColor Yellow
        while ($secondsRemaining -gt 0) {
            $sleepSeconds = [Math]::Min($secondsRemaining, 300)
            [void](Wait-LotCancellationAware -Seconds $sleepSeconds)
            if ((Get-LotCancellationState).Requested) { return }
            $secondsRemaining -= $sleepSeconds
        }
    }
}

function Resolve-LauncherPowerShellPath {
    $windowsPowerShell = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
    if (Test-Path -LiteralPath $windowsPowerShell -PathType Leaf) { return $windowsPowerShell }

    foreach ($commandName in @("powershell.exe","pwsh.exe")) {
        $command = Get-Command $commandName -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($command) { return $command.Source }
    }

    throw "No supported PowerShell executable was found."
}
$script:PowerShellExecutable = Resolve-LauncherPowerShellPath
function Invoke-FullIntuneInventoryExport {
    param(
        [Parameter(Mandatory=$true)][string]$ExportScriptPath,
        [Parameter(Mandatory=$true)][string]$OutputPath,
        [Parameter(Mandatory=$true)][string]$LogPath,
        [Parameter(Mandatory=$true)][int]$PageSize,
        [string]$ComputerListPath
    )

    try {
        if (-not (Test-Path -LiteralPath $ExportScriptPath)) {
            throw "SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1 not found: $ExportScriptPath"
        }

        $args = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", $ExportScriptPath,
            "-OutputPath", $OutputPath,
            "-PageSize", ([string]$PageSize),
            "-ForceRefresh"
        )
        if (-not [string]::IsNullOrWhiteSpace($ComputerListPath)) {
            $args += @("-ComputerListPath",$ComputerListPath)
        }

        $output = & $script:PowerShellExecutable @args 2>&1
        $exitCode = $LASTEXITCODE
        $output | Out-File -LiteralPath $LogPath -Encoding UTF8 -Force

        if ($exitCode -ne 0) {
            throw "SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1 exited with code $exitCode. Log=$LogPath"
        }
        if (-not (Test-Path -LiteralPath $OutputPath)) {
            throw "Intune CSV was not created: $OutputPath"
        }

        $set = Get-IntuneInventorySet -Path $OutputPath -NameColumn "ComputerName"

        return [PSCustomObject]@{
            Success = $true
            CsvPath = $OutputPath
            LogPath = $LogPath
            InventorySet = $set
            Error = ""
        }
    }
    catch {
        return [PSCustomObject]@{
            Success = $false
            CsvPath = $OutputPath
            LogPath = $LogPath
            InventorySet = @{}
            Error = $_.Exception.Message
        }
    }
}

function Invoke-FullEntraInventoryExport {
    param(
        [Parameter(Mandatory=$true)][string]$ExportScriptPath,
        [Parameter(Mandatory=$true)][string]$OutputPath,
        [Parameter(Mandatory=$true)][string]$LogPath,
        [Parameter(Mandatory=$true)][int]$PageSize,
        [string]$ComputerListPath
    )

    try {
        if (-not (Test-Path -LiteralPath $ExportScriptPath)) {
            throw "SmartM365-IntuneHybridJoinRepair-Export-EntraDevicesCsv.ps1 not found: $ExportScriptPath"
        }

        $args = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", $ExportScriptPath,
            "-OutputPath", $OutputPath,
            "-PageSize", ([string]$PageSize),
            "-ForceRefresh"
        )
        if (-not [string]::IsNullOrWhiteSpace($ComputerListPath)) {
            $args += @("-ComputerListPath",$ComputerListPath)
        }

        $output = & $script:PowerShellExecutable @args 2>&1
        $exitCode = $LASTEXITCODE
        $output | Out-File -LiteralPath $LogPath -Encoding UTF8 -Force

        if ($exitCode -ne 0) {
            throw "SmartM365-IntuneHybridJoinRepair-Export-EntraDevicesCsv.ps1 exited with code $exitCode. Log=$LogPath"
        }
        if (-not (Test-Path -LiteralPath $OutputPath)) {
            throw "Entra CSV was not created: $OutputPath"
        }

        $map = Get-EntraInventoryMap -Path $OutputPath -NameColumn "ComputerName"

        return [PSCustomObject]@{
            Success = $true
            CsvPath = $OutputPath
            LogPath = $LogPath
            InventoryMap = $map
            Error = ""
        }
    }
    catch {
        return [PSCustomObject]@{
            Success = $false
            CsvPath = $OutputPath
            LogPath = $LogPath
            InventoryMap = @{}
            Error = $_.Exception.Message
        }
    }
}

function Invoke-FullAdInventoryExport {
    param(
        [Parameter(Mandatory=$true)][string]$ExportScriptPath,
        [Parameter(Mandatory=$true)][string]$OutputPath,
        [Parameter(Mandatory=$true)][string]$LogPath,
        [Parameter(Mandatory=$false)][string]$Domain,
        [string]$ComputerListPath
    )

    try {
        if (-not (Test-Path -LiteralPath $ExportScriptPath)) {
            throw "SmartM365-IntuneHybridJoinRepair-Export-ADDevicesCsv.ps1 not found: $ExportScriptPath"
        }

        $args = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", $ExportScriptPath,
            "-OutputPath", $OutputPath,
            "-ForceRefresh"
        )
        if (-not [string]::IsNullOrWhiteSpace($Domain)) {
            $args += "-Domain"
            $args += $Domain
        }
        if (-not [string]::IsNullOrWhiteSpace($ComputerListPath)) {
            $args += @("-ComputerListPath",$ComputerListPath)
        }

        $output = & $script:PowerShellExecutable @args 2>&1
        $exitCode = $LASTEXITCODE
        $output | Out-File -LiteralPath $LogPath -Encoding UTF8 -Force

        if ($exitCode -ne 0) {
            throw "SmartM365-IntuneHybridJoinRepair-Export-ADDevicesCsv.ps1 exited with code $exitCode. Log=$LogPath"
        }
        if (-not (Test-Path -LiteralPath $OutputPath)) {
            throw "AD CSV was not created: $OutputPath"
        }

        $map = Get-AdInventoryMap -Path $OutputPath -NameColumn "ComputerName"

        return [PSCustomObject]@{
            Success = $true
            CsvPath = $OutputPath
            LogPath = $LogPath
            InventoryMap = $map
            Error = ""
        }
    }
    catch {
        return [PSCustomObject]@{
            Success = $false
            CsvPath = $OutputPath
            LogPath = $LogPath
            InventoryMap = @{}
            Error = $_.Exception.Message
        }
    }
}

function Get-NextActionFromLauncherStatus {
    param([Parameter(Mandatory=$true)][string]$Status)

    switch ($Status) {
        "SUCCESS" { return "NO_ACTION_ALREADY_INTUNE_OR_HEALTHY" }
        "AUDIT_SUCCESS_ALREADY_INTUNE" { return "NO_ACTION_ALREADY_INTUNE" }
        "AUDIT_INTUNE_MISSING" { return "RUN_REPAIR" }
        "AUDIT_STALE_INTUNE_ENROLLMENT_LOCAL" { return "CLEAN_STALE_INTUNE_OPTIN" }
        "INTUNE_ENROLLMENT_PENDING_CONFIRMATION" { return "RECHECK_LATER_INTUNE_ENROLLMENT" }
        "ADMIN_SHARE_UNREACHABLE" { return "FIX_ADMIN_SHARE_OR_NETWORK" }
        "DNS_PREFLIGHT_ALL_SAMPLES_FAILED" { return "CHECK_VPN_DNS_BEFORE_LOT" }
        "RUN_GUARD_ACTIVE" { return "WAIT_RUN_GUARD" }
        "ENDPOINT_RUN_ACTIVE" { return "WAIT_ACTIVE_ENDPOINT_RUN" }
        "SKIPPED_BY_STATUS_BACKOFF" { return "WAIT_STATUS_BACKOFF" }
        "REBOOT_SAFETY_LIMIT_REACHED_POST_DSREG_LEAVE" { return "REVIEW_REBOOT_HISTORY_AND_HYBRID_JOIN" }
        "REBOOT_TRIGGERED_WAITING_FOR_USER_LOGON" { return "WAIT_USER_LOGON" }
        "WAITING_FOR_INTERACTIVE_USER_LOGON" { return "WAIT_USER_LOGON" }
        "INTUNE_USER_AUTOENROLL_LOCAL_INTERACTIVE_USER" { return "LOGON_WITH_DOMAIN_OR_AAD_USER" }
        "INTUNE_USER_AUTOENROLL_TASK_NOT_FOUND" { return "FIX_GPO_USER_AUTOENROLL_TASK" }
        "STALE_INTUNE_ENROLLMENT_LOCAL" { return "CLEAN_STALE_INTUNE_OPTIN" }
        "NON_INTUNE_MDM_ENROLLED" { return "CLEAN_NON_INTUNE_MDM_OPTIN" }
        "ENTRA_HYBRID_PENDING_ADJ_TRIGGERED" { return "RECHECK_ENTRA_PENDING_AFTER_ADJ" }
        "ENTRA_HYBRID_PENDING_RETRY_EXHAUSTED" { return "CHECK_AD_CONNECT_OR_DUPLICATE_ENTRA_DEVICE" }
        "ENTRA_PENDING_RESOLVED_POST_CYCLE" { return "RECHECK_INTUNE_ENROLLMENT" }
        "USER_NOT_AZUREAD" { return "CHECK_USER_AAD_OR_LOGON_CONTEXT" }
        "USER_PRT_NOT_AVAILABLE" { return "CHECK_USER_PRT" }
        "USER_PRT_REFRESH_FAILED" { return "FIX_USER_PRT_OR_RELOGIN" }
        "USER_SESSION_REMOTE" { return "LOGON_ON_CONSOLE" }
        "INTUNE_AUTOENROLL_POLICY_NOT_CONFIGURED" { return "CHECK_GPO_AUTOENROLL" }
        "KEY_SIGN_TEST_FAILED" { return "REPAIR_HYBRID_JOIN_KEY_OR_ALLOW_LEAVE" }
        "INTUNE_ENROLLMENT_CONNECTIVITY_FAILED" { return "CHECK_CONNECTIVITY" }
        "DOMAIN_CONTROLLER_UNREACHABLE" { return "FIX_DOMAIN_CONNECTIVITY_OR_VPN" }
        "REMOTE_DIRECTORY_CREATE_FAILED" { return "FIX_SCRIPT_COPY_OR_ADMIN_SHARE" }
        "REMOTE_SCRIPT_COPY_FAILED" { return "FIX_SCRIPT_COPY_OR_SECURITY" }
        "REMOTE_SCRIPT_MISSING" { return "FIX_SCRIPT_COPY_OR_SECURITY" }
        "PSEXEC_TIMEOUT" { return "CHECK_REMOTE_LOG_OR_RETRY" }
        "PSEXEC_COMMUNICATION_LOST" { return "RETRY_PSEXEC_OR_CHECK_REMOTE_SERVICE" }
        "PSEXEC_EXIT_UNKNOWN" { return "CHECK_CURRENT_RUN_REMOTE_LOG" }
        "PSEXEC_EXIT_-1073741790" { return "CHECK_EDR_OR_EXECUTION_POLICY_BLOCK" }
        default {
            if ($Status -like "ERROR*") { return "CHECK_CONNECTIVITY_OR_ADMIN_ACCESS" }
            if ($Status -like "PSEXEC_EXIT_-1073741790*") { return "CHECK_EDR_OR_EXECUTION_POLICY_BLOCK" }
            if ($Status -like "PSEXEC_EXIT*") { return "CHECK_CURRENT_RUN_REMOTE_LOG" }
            return "REVIEW_LOGS"
        }
    }
}

function Get-ScriptVersionFromFile {
    param([Parameter(Mandatory=$true)][string]$Path)

    try {
        if (-not (Test-Path -LiteralPath $Path)) { return "" }
        $match = Select-String -LiteralPath $Path -Pattern '^\s*\$ScriptVersion\s*=\s*"([^"]+)"' -ErrorAction Stop | Select-Object -First 1
        if ($match -and $match.Matches.Count -gt 0) {
            return $match.Matches[0].Groups[1].Value
        }
    }
    catch { }

    return ""
}

function Get-FileSha256 {
    param([Parameter(Mandatory=$true)][string]$Path)

    try {
        if (-not (Test-Path -LiteralPath $Path)) { return "" }
        return (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash
    }
    catch {
        return ""
    }
}

function Get-RemoteEvidenceFinalStatus {
    param(
        [Parameter(Mandatory=$true)][string]$EvidencePath,
        [Parameter(Mandatory=$false)][datetime]$Since = [datetime]::MinValue,
        [switch]$RequireCompletedRun
    )

    $completedRunId = ""
    $completedRunStatus = ""
    $completedRunExitCode = ""
    $completedRunNextAction = ""
    $completedRunDetail = ""
    $lastRunPath = Join-Path $EvidencePath "LastRun.json"
    if ($RequireCompletedRun) {
        if (-not (Test-Path -LiteralPath $lastRunPath -ErrorAction SilentlyContinue)) { return $null }
        try {
            $lastRun = Get-Content -LiteralPath $lastRunPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($null -eq $lastRun -or [string]::IsNullOrWhiteSpace([string]$lastRun.RunId) -or [string]::IsNullOrWhiteSpace([string]$lastRun.EndTime)) {
                return $null
            }
            $lastRunStart = [datetime]$lastRun.StartTime
            if ($lastRunStart -lt $Since.AddSeconds(-5)) { return $null }
            $completedRunId = ([string]$lastRun.RunId).Trim()
            $completedRunStatus = ([string]$lastRun.Status).Trim()
            $completedRunExitCode = ([string]$lastRun.ExitCode).Trim()
            $completedRunNextAction = ([string]$lastRun.NextAction).Trim()
            $completedRunDetail = ([string]$lastRun.Detail).Trim()
            if ([string]::IsNullOrWhiteSpace($completedRunStatus) -or [string]::IsNullOrWhiteSpace($completedRunExitCode)) {
                return $null
            }
        }
        catch {
            return $null
        }
    }

    $csv = Get-ChildItem -LiteralPath $EvidencePath -Recurse -File -Filter "IntuneHybridJoinToolkit_*.csv" -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $Since.AddSeconds(-5) } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if (-not $csv) {
        if ($RequireCompletedRun -and -not [string]::IsNullOrWhiteSpace($completedRunStatus)) {
            return [PSCustomObject]@{
                RunId = $completedRunId
                Status = $completedRunStatus
                ExitCode = $completedRunExitCode
                NextAction = $completedRunNextAction
                Detail = $completedRunDetail
                CsvPath = ""
                LastRunPath = $lastRunPath
                LastRunStatus = $completedRunStatus
                LastRunExitCode = $completedRunExitCode
                InteractiveUserName = ""
                InteractiveUserDomain = ""
                InteractiveUserAccountName = ""
                InteractiveUserAccountType = ""
                InteractiveSessionName = ""
                InteractiveSessionState = ""
                UserIsUserAzureAD = ""
                UserAzureAdPrt = ""
                UserSessionIsNotRemote = ""
                RetryAfterRebootAction = ""
                RetryAfterRebootDetail = ""
                RetryAfterRebootAttempt = ""
                RetryAfterRebootMaxAttempts = ""
                RetryAfterRebootTaskName = ""
            }
        }
        return $null
    }

    try {
        $row = Import-Csv -LiteralPath $csv.FullName -ErrorAction Stop | Select-Object -Last 1
        if ($null -eq $row -or [string]::IsNullOrWhiteSpace([string]$row.Status)) { return $null }

        $rowRunId = ""
        if ($row.PSObject.Properties["RunId"]) { $rowRunId = ([string]$row.RunId).Trim() }
        if ($RequireCompletedRun) {
            if ([string]::IsNullOrWhiteSpace($rowRunId) -or $rowRunId -ne $completedRunId) {
                return [PSCustomObject]@{
                    RunId = $completedRunId
                    Status = $completedRunStatus
                    ExitCode = $completedRunExitCode
                    NextAction = $completedRunNextAction
                    Detail = $completedRunDetail
                    CsvPath = ""
                    LastRunPath = $lastRunPath
                    LastRunStatus = $completedRunStatus
                    LastRunExitCode = $completedRunExitCode
                    InteractiveUserName = ""
                    InteractiveUserDomain = ""
                    InteractiveUserAccountName = ""
                    InteractiveUserAccountType = ""
                    InteractiveSessionName = ""
                    InteractiveSessionState = ""
                    UserIsUserAzureAD = ""
                    UserAzureAdPrt = ""
                    UserSessionIsNotRemote = ""
                    RetryAfterRebootAction = ""
                    RetryAfterRebootDetail = ""
                    RetryAfterRebootAttempt = ""
                    RetryAfterRebootMaxAttempts = ""
                    RetryAfterRebootTaskName = ""
                }
            }
        }

        $exitCode = ""
        if ($row.PSObject.Properties["ExitCode"]) { $exitCode = [string]$row.ExitCode }

        $nextAction = ""
        if ($row.PSObject.Properties["NextAction"]) { $nextAction = [string]$row.NextAction }
        if ([string]::IsNullOrWhiteSpace($nextAction) -and $RequireCompletedRun) { $nextAction = $completedRunNextAction }

        $detail = ""
        if ($row.PSObject.Properties["DsregStatusErrorMessage"]) { $detail = [string]$row.DsregStatusErrorMessage }
        if ([string]::IsNullOrWhiteSpace($detail) -and $row.PSObject.Properties["ErrorMessage"]) { $detail = [string]$row.ErrorMessage }
        if ([string]::IsNullOrWhiteSpace($detail) -and $RequireCompletedRun) { $detail = $completedRunDetail }

        $interactiveUserName = ""; if ($row.PSObject.Properties["InteractiveUserName"]) { $interactiveUserName = [string]$row.InteractiveUserName }
        $interactiveUserDomain = ""; if ($row.PSObject.Properties["InteractiveUserDomain"]) { $interactiveUserDomain = [string]$row.InteractiveUserDomain }
        $interactiveUserAccountName = ""; if ($row.PSObject.Properties["InteractiveUserAccountName"]) { $interactiveUserAccountName = [string]$row.InteractiveUserAccountName }
        $interactiveUserAccountType = ""; if ($row.PSObject.Properties["InteractiveUserAccountType"]) { $interactiveUserAccountType = [string]$row.InteractiveUserAccountType }
        $interactiveSessionName = ""; if ($row.PSObject.Properties["InteractiveSessionName"]) { $interactiveSessionName = [string]$row.InteractiveSessionName }
        $interactiveSessionState = ""; if ($row.PSObject.Properties["InteractiveSessionState"]) { $interactiveSessionState = [string]$row.InteractiveSessionState }
        $userIsUserAzureAD = ""; if ($row.PSObject.Properties["User_IsUserAzureAD"]) { $userIsUserAzureAD = [string]$row.User_IsUserAzureAD }
        $userAzureAdPrt = ""; if ($row.PSObject.Properties["User_AzureAdPrt"]) { $userAzureAdPrt = [string]$row.User_AzureAdPrt }
        $userSessionIsNotRemote = ""; if ($row.PSObject.Properties["User_SessionIsNotRemote"]) { $userSessionIsNotRemote = [string]$row.User_SessionIsNotRemote }
        $retryAfterRebootAction = ""; if ($row.PSObject.Properties["RetryAfterRebootAction"]) { $retryAfterRebootAction = [string]$row.RetryAfterRebootAction }
        $retryAfterRebootDetail = ""; if ($row.PSObject.Properties["RetryAfterRebootDetail"]) { $retryAfterRebootDetail = [string]$row.RetryAfterRebootDetail }
        $retryAfterRebootAttempt = ""; if ($row.PSObject.Properties["RetryAfterRebootAttempt"]) { $retryAfterRebootAttempt = [string]$row.RetryAfterRebootAttempt }
        $retryAfterRebootMaxAttempts = ""; if ($row.PSObject.Properties["RetryAfterRebootMaxAttempts"]) { $retryAfterRebootMaxAttempts = [string]$row.RetryAfterRebootMaxAttempts }
        $retryAfterRebootTaskName = ""; if ($row.PSObject.Properties["RetryAfterRebootTaskName"]) { $retryAfterRebootTaskName = [string]$row.RetryAfterRebootTaskName }

        return [PSCustomObject]@{
            RunId = $rowRunId
            Status = $(if ($RequireCompletedRun) { $completedRunStatus } else { ([string]$row.Status).Trim() })
            ExitCode = $(if ($RequireCompletedRun) { $completedRunExitCode } else { $exitCode.Trim() })
            NextAction = $nextAction.Trim()
            Detail = $detail.Trim()
            CsvPath = $csv.FullName
            LastRunPath = $(if ($RequireCompletedRun) { $lastRunPath } else { "" })
            LastRunStatus = $completedRunStatus
            LastRunExitCode = $completedRunExitCode
            InteractiveUserName = $interactiveUserName.Trim()
            InteractiveUserDomain = $interactiveUserDomain.Trim()
            InteractiveUserAccountName = $interactiveUserAccountName.Trim()
            InteractiveUserAccountType = $interactiveUserAccountType.Trim()
            InteractiveSessionName = $interactiveSessionName.Trim()
            InteractiveSessionState = $interactiveSessionState.Trim()
            UserIsUserAzureAD = $userIsUserAzureAD.Trim()
            UserAzureAdPrt = $userAzureAdPrt.Trim()
            UserSessionIsNotRemote = $userSessionIsNotRemote.Trim()
            RetryAfterRebootAction = $retryAfterRebootAction.Trim()
            RetryAfterRebootDetail = $retryAfterRebootDetail.Trim()
            RetryAfterRebootAttempt = $retryAfterRebootAttempt.Trim()
            RetryAfterRebootMaxAttempts = $retryAfterRebootMaxAttempts.Trim()
            RetryAfterRebootTaskName = $retryAfterRebootTaskName.Trim()
        }
    }
    catch {
        return [PSCustomObject]@{
            Status = ""
            ExitCode = ""
            NextAction = ""
            Detail = ("Could not parse remote evidence CSV: {0}" -f $_.Exception.Message)
            CsvPath = $csv.FullName
        }
    }
}

function Get-LauncherReportColumns {
    @(
        "LauncherVersion",
        "Cycle",
        "Computer",
        "ConnectionTarget",
        "Timestamp",
        "DryRun",
        "DnsResolved",
        "DnsAddressList",
        "AdminShareReachable",
        "RemotePayloadCopyAttempts",
        "PingReachable",
        "IsVirtualMachine",
        "VirtualMachineEvidence",
        "RemoteDirectoryCreated",
        "ScriptCopied",
        "LocalScriptVersion",
        "RemoteScriptVersion",
        "LocalScriptHash",
        "RemoteScriptHash",
        "PsExecExitCode",
        "RemoteStatus",
        "RemoteExitCode",
        "RemoteNextAction",
        "RemoteDetail",
        "RetryAfterRebootAction",
        "RetryAfterRebootDetail",
        "RetryAfterRebootAttempt",
        "RetryAfterRebootMaxAttempts",
        "RetryAfterRebootTaskName",
        "NextAction",
        "EffectiveStatus",
        "EffectiveNextAction",
        "LatestAttemptStatus",
        "LatestActionableStatus",
        "LatestActionableNextAction",
        "BackoffStatus",
        "BackoffUntilUtc",
        "BackoffCount",
        "InteractiveUserName",
        "InteractiveUserDomain",
        "InteractiveUserAccountName",
        "InteractiveUserAccountType",
        "InteractiveSessionName",
        "InteractiveSessionState",
        "UserIsUserAzureAD",
        "UserAzureAdPrt",
        "UserSessionIsNotRemote",
        "IntuneInventoryPresent",
        "EntraInventoryPresent",
        "EntraRegisteredState",
        "EntraAlternativeSecurityIdCount",
        "EntraPendingReason",
        "EntraRegistrationDateTime",
        "EntraTrustType",
        "EntraDeviceId",
        "EntraObjectId",
        "ADInventoryPresent",
        "ADDomain",
        "ADEnabled",
        "ADDNSHostName",
        "ADDistinguishedName",
        "ADOperatingSystem",
        "ADLastLogonTimestampUtc",
        "AdminShareFailureType",
        "PostCycleIntuneInventoryChecked",
        "PostCycleIntuneInventoryPresent",
        "PostCycleIntuneEnrollmentDetected",
        "PostCycleIntuneInventoryCsv",
        "PostCycleIntuneInventoryError",
        "PostCycleEntraInventoryChecked",
        "PostCycleEntraInventoryPresent",
        "PostCycleEntraRegisteredState",
        "PostCycleEntraAlternativeSecurityIdCount",
        "PostCycleEntraPendingResolved",
        "PostCycleEntraInventoryCsv",
        "PostCycleEntraInventoryError",
        "PostCycleADInventoryChecked",
        "PostCycleADInventoryPresent",
        "PostCycleADInventoryCsv",
        "PostCycleADInventoryError",
        "RemoteLogsCollected",
        "RemoteLogsPath",
        "RemoteCurrentRunLogsPath",
        "RemoteLogsError",
        "Status",
        "LogPath",
        "ErrorMessage",
        "JobErrorMessage"
    )
}

function Initialize-LiveCycleReport {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string[]]$Columns
    )

    $header = ($Columns | ForEach-Object { '"' + ($_ -replace '"','""') + '"' }) -join ","
    Set-Content -LiteralPath $Path -Value $header -Encoding UTF8 -Force
}

function Add-LiveCycleReportRow {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string[]]$Columns,
        [Parameter(Mandatory=$true)][psobject]$Row
    )

    ConvertTo-PortableReportRow -Row $Row | Select-Object $Columns | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 -Append
}

function Get-ComputerList {
    param([Parameter(Mandatory=$true)][string]$Path)

    @(Get-Content -LiteralPath $Path |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.StartsWith("#") } |
        Select-Object -Unique)
}
function Test-ComputerListPresentWithRetry {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [ValidateRange(1,20)][int]$Attempts = 5,
        [ValidateRange(0,5000)][int]$DelayMilliseconds = 200
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
    param([Parameter(Mandatory=$true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return [pscustomobject]@{ RawLines = 0; Unique = 0; DuplicateGroups = 0; DuplicateLines = 0; DuplicateSamples = '' }
    }

    $rawNames = @(
        Get-Content -LiteralPath $Path |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.StartsWith('#') }
    )
    $duplicateGroups = @(
        $rawNames |
            Group-Object { ([string]$_).ToUpperInvariant() } |
            Where-Object { $_.Count -gt 1 }
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
function Get-ComputerListKey {
    param([Parameter(Mandatory=$true)][string]$ComputerName)

    return ($ComputerName.Trim().Split(".")[0]).ToUpperInvariant()
}

function Get-PostCycleCloudRefreshRows {
    param([Parameter(Mandatory=$false)][object[]]$Rows = @())

    return @(
        $Rows | Where-Object {
            [string]$_.Status -notmatch '^(ADMIN_SHARE_UNREACHABLE|DNS_PREFLIGHT_ALL_SAMPLES_FAILED|REMOTE_DIRECTORY_CREATE_FAILED|REMOTE_SCRIPT_COPY_FAILED|REMOTE_SCRIPT_MISSING|CANCELLED_|SKIPPED_|DRYRUN_)'
        }
    )
}

function Merge-ScopedInventoryMap {
    param(
        [Parameter(Mandatory=$false)][hashtable]$ExistingMap,
        [Parameter(Mandatory=$false)][hashtable]$RefreshedMap,
        [Parameter(Mandatory=$false)][string[]]$ScopedComputers = @()
    )

    if ($null -eq $ExistingMap) { $ExistingMap = @{} }
    [void]$ExistingMap.Remove('__SMARTM365_INVENTORY_CHECKED__')

    foreach ($computer in @($ScopedComputers)) {
        if ([string]::IsNullOrWhiteSpace([string]$computer)) { continue }
        [void]$ExistingMap.Remove((Get-ComputerListKey -ComputerName $computer))
    }

    if ($null -ne $RefreshedMap) {
        foreach ($entry in $RefreshedMap.GetEnumerator()) {
            if ([string]$entry.Key -eq '__SMARTM365_INVENTORY_CHECKED__') { continue }
            $ExistingMap[[string]$entry.Key] = $entry.Value
        }
    }

    return $ExistingMap
}

function Test-AdInventoryRefreshDue {
    param(
        [Parameter(Mandatory=$false)][AllowNull()][object]$LastRefreshUtc,
        [Parameter(Mandatory=$true)][ValidateRange(1,168)][int]$FreshnessHours,
        [Parameter(Mandatory=$false)][datetime]$NowUtc = [datetime]::UtcNow
    )

    if ($null -eq $LastRefreshUtc) { return $true }
    try {
        $lastRefresh = ([datetime]$LastRefreshUtc).ToUniversalTime()
        return (($NowUtc.ToUniversalTime() - $lastRefresh).TotalHours -ge $FreshnessHours)
    }
    catch {
        return $true
    }
}

function Get-AdaptiveCycleDelaySeconds {
    param(
        [Parameter(Mandatory=$false)][object[]]$Rows = @(),
        [Parameter(Mandatory=$true)][ValidateRange(0,86400)][int]$MinimumDelaySeconds,
        [Parameter(Mandatory=$false)][datetime]$NowUtc = [datetime]::UtcNow
    )

    $effectiveRows = @($Rows)
    if ($effectiveRows.Count -eq 0) { return $MinimumDelaySeconds }

    $backoffStatuses = @('SKIPPED_BY_STATUS_BACKOFF','SKIPPED_BY_TECH_RUN_GUARD_STARTED_NO_RESULT')
    foreach ($row in $effectiveRows) {
        $status = if ($row.PSObject.Properties['EffectiveStatus'] -and -not [string]::IsNullOrWhiteSpace([string]$row.EffectiveStatus)) { [string]$row.EffectiveStatus } else { [string]$row.Status }
        if ($backoffStatuses -notcontains $status) { return $MinimumDelaySeconds }
    }

    $futureExpiries = New-Object System.Collections.ArrayList
    foreach ($row in $effectiveRows) {
        $rawExpiry = if ($row.PSObject.Properties['BackoffUntilUtc']) { [string]$row.BackoffUntilUtc } else { '' }
        if ([string]::IsNullOrWhiteSpace($rawExpiry)) { continue }
        $parsedExpiry = [datetime]::MinValue
        if ([datetime]::TryParse($rawExpiry,[ref]$parsedExpiry)) {
            $parsedExpiry = $parsedExpiry.ToUniversalTime()
            if ($parsedExpiry -gt $NowUtc.ToUniversalTime()) { [void]$futureExpiries.Add($parsedExpiry) }
        }
    }

    if ($futureExpiries.Count -eq 0) { return $MinimumDelaySeconds }
    $earliestExpiry = @($futureExpiries | Sort-Object | Select-Object -First 1)[0]
    $untilExpiry = [int][math]::Ceiling(($earliestExpiry - $NowUtc.ToUniversalTime()).TotalSeconds)
    return [math]::Max($MinimumDelaySeconds,$untilExpiry)
}

function Get-TechnicianRunGuardHistoryPath {
    $stateRoot = Join-Path ([Environment]::GetFolderPath("CommonApplicationData")) "SmartM365\IntuneHybridJoinToolkit\LauncherState"
    return (Join-Path $stateRoot "RunGuardHistory.json")
}

function Invoke-TechnicianRunGuardHistoryLock {
    param([Parameter(Mandatory=$true)][scriptblock]$ScriptBlock,[object[]]$ArgumentList=@())
    $lockStream = $null
    $lockPath = '{0}.lock' -f (Get-TechnicianRunGuardHistoryPath)
    $lockFolder = Split-Path -Parent $lockPath
    $deadlineUtc = (Get-Date).ToUniversalTime().AddSeconds(60)
    if (-not (Test-Path -LiteralPath $lockFolder -PathType Container)) { New-Item -ItemType Directory -Path $lockFolder -Force | Out-Null }
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
        if ($null -eq $lockStream) { throw "Timed out waiting for technician run guard history file lock after 60 seconds." }
        & $ScriptBlock @ArgumentList
    }
    finally {
        if ($lockStream) { $lockStream.Dispose() }
    }
}

function Get-TechnicianRunGuardFqdn {
    param([Parameter(Mandatory=$true)][string]$ComputerName,[AllowNull()][hashtable]$AdInventoryMap)
    $name = $ComputerName.Trim().Trim([char]34).TrimEnd(".")
    if ([string]::IsNullOrWhiteSpace($name)) { return "" }
    if ($name.Contains(".")) { return $name.ToLowerInvariant() }
    $key = Get-ComputerListKey -ComputerName $name
    if ($AdInventoryMap -and $AdInventoryMap.ContainsKey($key)) {
        $row = $AdInventoryMap[$key]
        if ($row.PSObject.Properties["DNSHostName"] -and -not [string]::IsNullOrWhiteSpace([string]$row.DNSHostName)) {
            return ([string]$row.DNSHostName).Trim().TrimEnd(".").ToLowerInvariant()
        }
    }
    return $name.ToLowerInvariant()
}

function Read-TechnicianRunGuardHistory {
    param([Parameter(Mandatory=$true)][string]$Path,[ValidateRange(0,168)][int]$Hours)
    $entries = @()
    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        try {
            $data = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($data.PSObject.Properties["Entries"]) { $entries = @($data.Entries) }
        }
        catch { $entries = @() }
    }
    $now = (Get-Date).ToUniversalTime()
    $retained = @($entries | Where-Object {
        try {
            $retentionText = if ($_.PSObject.Properties["RetentionUtc"]) { [string]$_.RetentionUtc } else { [string]$_.ExpiresUtc }
            ([datetime]$retentionText).ToUniversalTime() -gt $now
        }
        catch { $false }
    })
    return [pscustomobject]@{ Version=2; UpdatedUtc=$now.ToString("o"); Entries=$retained }
}

function Save-TechnicianRunGuardHistory {
    param([Parameter(Mandatory=$true)][string]$Path,[Parameter(Mandatory=$true)]$History)
    $folder = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -ItemType Directory -Path $folder -Force | Out-Null }
    $History.UpdatedUtc = (Get-Date).ToUniversalTime().ToString("o")
    $json = $History | ConvertTo-Json -Depth 8
    $lastError = $null
    for ($attempt = 1; $attempt -le 10; $attempt++) {
        $temporaryPath = Join-Path $folder (".{0}.{1}.tmp" -f (Split-Path -Leaf $Path),[guid]::NewGuid().ToString("N"))
        try {
            Set-Content -LiteralPath $temporaryPath -Value $json -Encoding UTF8 -Force -ErrorAction Stop
            Move-Item -LiteralPath $temporaryPath -Destination $Path -Force -ErrorAction Stop
            return
        }
        catch [System.IO.IOException] {
            $lastError = $_
            Start-Sleep -Milliseconds ([math]::Min(2000,150 * $attempt))
        }
        catch [System.UnauthorizedAccessException] {
            $lastError = $_
            Start-Sleep -Milliseconds ([math]::Min(2000,150 * $attempt))
        }
        finally {
            if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
                try { Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction Stop } catch { }
            }
        }
    }
    if ($lastError) { throw $lastError }
    throw ("Failed to save technician run guard history: {0}" -f $Path)
}

function Get-TechnicianRunGuardCooldownHours {
    param([string]$Status,[int]$BackoffCount=1,[int]$CycleNumber=1,[int]$DefaultHours=12)
    $normalized = ([string]$Status).ToUpperInvariant()
    if ($normalized -match "^(CANCELLED_|DRYRUN_|AUDIT_|SKIPPED_VIRTUAL_MACHINE)") { return 0.0 }
    if ($normalized -match "^(ADMIN_SHARE|DNS_|PSEXEC_|REMOTE_.*FAILED|CENTRAL_.*FAILED|JOB_ERROR|ERROR)") {
        $exponent = [math]::Min(5,[math]::Max(0,[math]::Max($BackoffCount,$CycleNumber) - 1))
        return ([math]::Min(360,15 * [math]::Pow(2,$exponent)) / 60.0)
    }
    if ($normalized -match "(USER|INTERACTIVE|PRT|LOGON|SESSION_REMOTE|LOCAL_INTERACTIVE)") { return 24.0 }
    if ($normalized -match "(RUN_GUARD|ENDPOINT_RUN_ACTIVE|PENDING_CONFIRMATION|REBOOT_TRIGGERED|WAITING_FOR_AAD_CONNECT|WAITING_POST_LEAVE)") { return 12.0 }
    if ($normalized -in @("SUCCESS","AUDIT_SUCCESS_ALREADY_INTUNE","ENROLLED_DETECTED_POST_CYCLE")) { return 168.0 }
    return [double][math]::Max(1,$DefaultHours)
}

function Test-TechnicianRunGuardEntryShouldBlock {
    param([AllowNull()]$Entry)
    if (-not $Entry) { return $false }
    try { return (([datetime]$Entry.ExpiresUtc).ToUniversalTime() -gt (Get-Date).ToUniversalTime()) } catch { return $false }
}

function Get-ActiveTechnicianRunGuardEntry {
    param([string]$Path,[string]$ComputerFqdn,[int]$Hours)
    if ([string]::IsNullOrWhiteSpace($ComputerFqdn) -or $Hours -le 0) { return $null }
    $found = @{}
    Invoke-TechnicianRunGuardHistoryLock -ArgumentList @($Path,$Hours,$ComputerFqdn,$found) -ScriptBlock {
        param($LockedPath,$LockedHours,$LockedFqdn,$FoundRef)
        $history = Read-TechnicianRunGuardHistory -Path $LockedPath -Hours $LockedHours
        Save-TechnicianRunGuardHistory -Path $LockedPath -History $history
        foreach ($entry in @($history.Entries)) {
            if ([string]$entry.ComputerFqdn -eq $LockedFqdn -and (Test-TechnicianRunGuardEntryShouldBlock $entry)) { $FoundRef.Value=$entry; break }
        }
    }
    if ($found.ContainsKey("Value")) { return $found.Value }
    return $null
}

function Update-TechnicianRunGuardHistory {
    param([string]$Path,[string]$ComputerFqdn,[string]$InputComputerName,[int]$Hours,[ValidateSet("Started","Result")][string]$State,[AllowNull()]$Result,[string]$JobId,[int]$CycleNumber)
    if ([string]::IsNullOrWhiteSpace($ComputerFqdn) -or $Hours -le 0) { return }
    Invoke-TechnicianRunGuardHistoryLock -ArgumentList @($Path,$ComputerFqdn,$InputComputerName,$Hours,$State,$Result,$JobId,$CycleNumber) -ScriptBlock {
        param($LockedPath,$LockedFqdn,$LockedName,$LockedHours,$LockedState,$LockedResult,$LockedJobId,$LockedCycle)
        $history = Read-TechnicianRunGuardHistory -Path $LockedPath -Hours $LockedHours
        $previous = $history.Entries | Where-Object { [string]$_.ComputerFqdn -eq $LockedFqdn } | Select-Object -First 1
        $kept = @($history.Entries | Where-Object { [string]$_.ComputerFqdn -ne $LockedFqdn })
        $now = (Get-Date).ToUniversalTime()
        $started = $now
        if ($LockedState -eq "Result" -and $previous -and $previous.LastStartedUtc) { try { $started=([datetime]$previous.LastStartedUtc).ToUniversalTime() } catch {} }
        $launcherStatus = if ($LockedResult -and $LockedResult.PSObject.Properties["Status"]) { [string]$LockedResult.Status } else { "STARTED_NO_RESULT" }
        $remoteStatus = if ($LockedResult -and $LockedResult.PSObject.Properties["RemoteStatus"]) { [string]$LockedResult.RemoteStatus } else { "" }
        $previousStatus = if ($previous -and $previous.PSObject.Properties["BackoffStatus"]) { [string]$previous.BackoffStatus } else { "" }
        $effectiveStatus = if ($LockedState -eq "Started" -and -not [string]::IsNullOrWhiteSpace($previousStatus)) { $previousStatus } else { @($remoteStatus,$launcherStatus) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1 }
        $previousBackoffCount = if ($previous -and $previous.PSObject.Properties["BackoffCount"]) { [int]$previous.BackoffCount } else { 0 }
        $backoffCount = if ($LockedState -eq "Started") { [math]::Max(1,$previousBackoffCount) } elseif ($previousStatus -eq $effectiveStatus) { $previousBackoffCount + 1 } else { 1 }
        $cooldownHours = if ($LockedState -eq "Started") { [double]$LockedHours } else { Get-TechnicianRunGuardCooldownHours -Status $effectiveStatus -BackoffCount $backoffCount -CycleNumber $LockedCycle -DefaultHours $LockedHours }
        $entry = [pscustomobject]@{ ComputerFqdn=$LockedFqdn; InputComputerName=$LockedName; State=$LockedState; LastStartedUtc=$started.ToString("o"); LastUpdatedUtc=$now.ToString("o"); ExpiresUtc=$now.AddHours($cooldownHours).ToString("o"); RetentionUtc=$now.AddDays(7).ToString("o"); LauncherStatus=$launcherStatus; RemoteStatus=$remoteStatus; BackoffStatus=$effectiveStatus; BackoffCount=$backoffCount; CooldownHours=[math]::Round($cooldownHours,2); NextAction=$(if($LockedResult -and $LockedResult.PSObject.Properties["NextAction"]){[string]$LockedResult.NextAction}else{""}); JobId=$LockedJobId; CycleNumber=$LockedCycle; LotRoot=$LotRoot; LauncherVersion=$LauncherVersion }
        if ($LockedState -ne "Result" -or $cooldownHours -gt 0) { $kept += $entry }
        $history.Entries = @($kept)
        Save-TechnicianRunGuardHistory -Path $LockedPath -History $history
    }
}

function New-TechnicianRunGuardSkippedResult {
    param([string]$ComputerName,[int]$CycleNumber,[Parameter(Mandatory=$true)]$HistoryEntry)
    $startedNoResult = ([string]$HistoryEntry.State -ne "Result")
    $skipStatus = if ($startedNoResult) { "SKIPPED_BY_TECH_RUN_GUARD_STARTED_NO_RESULT" } else { "SKIPPED_BY_STATUS_BACKOFF" }
    $backoffStatus = if ($HistoryEntry.PSObject.Properties["BackoffStatus"]) { [string]$HistoryEntry.BackoffStatus } else { [string]$HistoryEntry.RemoteStatus }
    $backoffCount = if ($HistoryEntry.PSObject.Properties["BackoffCount"]) { [string]$HistoryEntry.BackoffCount } else { "" }
    return [pscustomobject]@{ LauncherVersion=$LauncherVersion; Cycle=$CycleNumber; Computer=$ComputerName; Timestamp=Get-Date; DryRun=[bool]$DryRun; Status=$skipStatus; RemoteStatus=""; NextAction="WAIT_STATUS_BACKOFF"; RemoteNextAction=""; EffectiveStatus=$skipStatus; EffectiveNextAction="WAIT_STATUS_BACKOFF"; BackoffStatus=$backoffStatus; BackoffUntilUtc=[string]$HistoryEntry.ExpiresUtc; BackoffCount=$backoffCount; RemoteLogsCollected=$false; RemoteLogsPath=""; LogPath=$script:LauncherLogPath; ErrorMessage=("Technician status backoff active. FQDN={0}; PriorStatus={1}; Count={2}; Started={3}; NextEligible={4}; State={5}; JobId={6}" -f $HistoryEntry.ComputerFqdn,$backoffStatus,$backoffCount,$HistoryEntry.LastStartedUtc,$HistoryEntry.ExpiresUtc,$HistoryEntry.State,$HistoryEntry.JobId) }
}

function Test-AlreadyEnrolledCycleResult {
    param([Parameter(Mandatory=$true)][psobject]$Result)

    $status = ""
    $nextAction = ""
    if ($Result.PSObject.Properties["Status"]) { $status = ([string]$Result.Status).Trim() }
    if ($Result.PSObject.Properties["NextAction"]) { $nextAction = ([string]$Result.NextAction).Trim() }

    if ($Result.PSObject.Properties["IntuneInventoryPresent"] -and $Result.IntuneInventoryPresent -eq $true) {
        return $true
    }

    if ($Result.PSObject.Properties["PostCycleIntuneInventoryPresent"] -and $Result.PostCycleIntuneInventoryPresent -eq $true) {
        return $true
    }

    if ($status -in @("SUCCESS","AUDIT_SUCCESS_ALREADY_INTUNE")) {
        return $true
    }

    if ($nextAction -in @("NO_ACTION_ALREADY_INTUNE","NO_ACTION_ALREADY_INTUNE_OR_HEALTHY","NO_ACTION_INTUNE_PRESENT")) {
        return $true
    }

    return $false
}

function Move-AlreadyEnrolledComputersFromList {
    param(
        [Parameter(Mandatory=$true)][string]$ComputerListPath,
        [Parameter(Mandatory=$true)][object[]]$CycleSummary
    )

    $alreadyEnrolled = @(
        $CycleSummary |
            Where-Object { $_ -and (Test-AlreadyEnrolledCycleResult -Result $_) } |
            ForEach-Object { [string]$_.Computer } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
    )

    if ($alreadyEnrolled.Count -eq 0) {
        return [PSCustomObject]@{
            Moved = 0
            AlreadyEnrolledPath = ""
            Detail = "No already-enrolled computer detected in this cycle."
        }
    }

    $moveKeys = @{}
    foreach ($computer in $alreadyEnrolled) {
        $key = Get-ComputerListKey -ComputerName $computer
        if (-not [string]::IsNullOrWhiteSpace($key) -and -not $moveKeys.ContainsKey($key)) {
            $moveKeys[$key] = $computer.Trim()
        }
    }

    $listLines = @()
    if (Test-Path -LiteralPath $ComputerListPath) {
        $listLines = @(Get-Content -LiteralPath $ComputerListPath -ErrorAction Stop)
    }

    $remainingLines = New-Object System.Collections.Generic.List[string]
    $movedFromList = New-Object System.Collections.Generic.List[string]

    foreach ($line in $listLines) {
        $trimmed = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed.StartsWith("#")) {
            $remainingLines.Add($line)
            continue
        }

        $key = Get-ComputerListKey -ComputerName $trimmed
        if ($moveKeys.ContainsKey($key)) {
            $movedFromList.Add($trimmed)
            continue
        }

        $remainingLines.Add($line)
    }

    if ($movedFromList.Count -eq 0) {
        return [PSCustomObject]@{
            Moved = 0
            AlreadyEnrolledPath = ""
            Detail = "Already-enrolled computers were detected, but none were still present in Computers.txt."
        }
    }

    $computerListDir = Split-Path -Parent $ComputerListPath
    if ([string]::IsNullOrWhiteSpace($computerListDir)) { $computerListDir = "." }
    $alreadyEnrolledPath = Join-Path $computerListDir "ComputersAlreadyEnrolled.txt"

    $existingKeys = @{}
    if (Test-Path -LiteralPath $alreadyEnrolledPath) {
        foreach ($line in @(Get-Content -LiteralPath $alreadyEnrolledPath -ErrorAction SilentlyContinue)) {
            $trimmed = $line.Trim()
            if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed.StartsWith("#")) { continue }
            $key = Get-ComputerListKey -ComputerName $trimmed
            if (-not $existingKeys.ContainsKey($key)) { $existingKeys[$key] = $true }
        }
    }

    $appendLines = New-Object System.Collections.Generic.List[string]
    foreach ($computer in $movedFromList) {
        $key = Get-ComputerListKey -ComputerName $computer
        if (-not $existingKeys.ContainsKey($key)) {
            $appendLines.Add($computer)
            $existingKeys[$key] = $true
        }
    }

    $tmpComputerListPath = "{0}.tmp.{1}.txt" -f $ComputerListPath,([guid]::NewGuid().ToString("N"))
    try {
        Set-Content -LiteralPath $tmpComputerListPath -Value $remainingLines -Encoding ASCII -Force
        Move-Item -LiteralPath $tmpComputerListPath -Destination $ComputerListPath -Force
    }
    finally {
        Remove-Item -LiteralPath $tmpComputerListPath -Force -ErrorAction SilentlyContinue
    }

    if ($appendLines.Count -gt 0) {
        Add-Content -LiteralPath $alreadyEnrolledPath -Value $appendLines -Encoding ASCII
    }
    elseif (-not (Test-Path -LiteralPath $alreadyEnrolledPath)) {
        New-Item -ItemType File -Path $alreadyEnrolledPath -Force | Out-Null
    }

    return [PSCustomObject]@{
        Moved = $movedFromList.Count
        AlreadyEnrolledPath = $alreadyEnrolledPath
        Detail = ("Moved {0} computer(s) from Computers.txt to ComputersAlreadyEnrolled.txt." -f $movedFromList.Count)
    }
}

function ConvertTo-HtmlText {
    param([object]$Value)

    if ($null -eq $Value) {
        return ""
    }

    return [System.Security.SecurityElement]::Escape([string]$Value)
}

function ConvertTo-FileUri {
    param([AllowNull()][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $value = ([string]$Path).Trim()
    try {
        if ($value.StartsWith('\\')) {
            $parts = $value.TrimStart('\') -split '\\'
            if ($parts.Count -ge 2) {
                $server = [System.Uri]::EscapeDataString($parts[0])
                $share = [System.Uri]::EscapeDataString($parts[1])
                $rest = ''
                if ($parts.Count -gt 2) {
                    $rest = '/' + (($parts[2..($parts.Count - 1)] | ForEach-Object { [System.Uri]::EscapeDataString($_) }) -join '/')
                }
                return ('file://{0}/{1}{2}' -f $server,$share,$rest)
            }
        }
        return ([System.Uri]::new([System.IO.Path]::GetFullPath($value))).AbsoluteUri
    }
    catch { return '' }
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
    return ('\\{0}\C$\ProgramData\SmartM365\IntuneHybridJoinToolkit\Logs' -f ([string]$ComputerName).Trim())
}

function Get-HybridJoinHtmlEffectiveStatus {
    param([AllowNull()][object]$Row)
    if ($null -eq $Row) { return '' }
    if ($Row.PSObject.Properties['EffectiveStatus'] -and -not [string]::IsNullOrWhiteSpace([string]$Row.EffectiveStatus)) { return [string]$Row.EffectiveStatus }
    if ($Row.PSObject.Properties['Status']) { return [string]$Row.Status }
    return ''
}

function Get-HybridJoinHtmlEffectiveNextAction {
    param([AllowNull()][object]$Row)
    if ($null -eq $Row) { return '' }
    if ($Row.PSObject.Properties['EffectiveNextAction'] -and -not [string]::IsNullOrWhiteSpace([string]$Row.EffectiveNextAction)) { return [string]$Row.EffectiveNextAction }
    if ($Row.PSObject.Properties['NextAction']) { return [string]$Row.NextAction }
    return ''
}

function Get-HybridJoinAdminShareFailureType {
    param([AllowNull()][object]$Row)
    if ($null -eq $Row) { return 'UNKNOWN' }
    if ($Row.PSObject.Properties['AdminShareFailureType'] -and -not [string]::IsNullOrWhiteSpace([string]$Row.AdminShareFailureType)) { return [string]$Row.AdminShareFailureType }
    $evidence = @(
        $(if ($Row.PSObject.Properties['RemoteDetail']) { [string]$Row.RemoteDetail } else { '' }),
        $(if ($Row.PSObject.Properties['ErrorMessage']) { [string]$Row.ErrorMessage } else { '' }),
        $(if ($Row.PSObject.Properties['JobErrorMessage']) { [string]$Row.JobErrorMessage } else { '' })
    ) -join ' '
    if ($evidence -match 'DNS_FAILED') { return 'DNS_FAILED' }
    if ($evidence -match 'PING_FAILED_ADMIN_SHARE_FAILED') { return 'PING_FAILED_ADMIN_SHARE_FAILED' }
    if ($evidence -match 'PING_OK_ADMIN_SHARE_FAILED') { return 'PING_OK_ADMIN_SHARE_FAILED' }
    return 'UNKNOWN'
}

function Get-HybridJoinHtmlReportRows {
    param([AllowEmptyCollection()][object[]]$Items)

    $columns = @(Get-LauncherReportColumns)
    return @($Items | ForEach-Object {
        $portable = ConvertTo-PortableReportRow -Row $_
        $row = [ordered]@{}
        foreach ($column in $columns) { $row[$column] = if ($portable.PSObject.Properties[$column]) { [string]$portable.$column } else { '' } }
        $row['Local log'] = New-HtmlLogLink -Path $row['LogPath']
        $row['Collected logs'] = New-HtmlLogLink -Path $row['RemoteLogsPath']
        $row['Current run logs'] = New-HtmlLogLink -Path $row['RemoteCurrentRunLogsPath']
        $row['Remote PC logs'] = New-HtmlLogLink -Path (Get-RemotePcLogsPath -ComputerName $row['Computer'])
        [pscustomobject]$row
    })
}

function Get-HybridJoinLatestRowsByComputer {
    param([AllowEmptyCollection()][object[]]$Rows)

    $latestByComputer = @{}
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $computer = if ($row.PSObject.Properties['Computer']) { [string]$row.Computer } else { '' }
        if ([string]::IsNullOrWhiteSpace($computer)) { continue }
        $key = Get-ComputerListKey -ComputerName $computer
        if ([string]::IsNullOrWhiteSpace($key)) { $key = $computer.ToUpperInvariant() }
        $latestByComputer[$key] = $row
    }

    return @($latestByComputer.GetEnumerator() | Sort-Object Name | ForEach-Object { $_.Value })
}

function Get-HybridJoinLatestActionableRowsByComputer {
    param([AllowEmptyCollection()][object[]]$Rows)

    $latestByComputer = @{}
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $computer = if ($row.PSObject.Properties["Computer"]) { [string]$row.Computer } else { "" }
        if ([string]::IsNullOrWhiteSpace($computer)) { continue }
        $status = Get-HybridJoinHtmlEffectiveStatus -Row $row
        if ($status -match "^(CANCELLED_|RUN_GUARD_ACTIVE$|SKIPPED_BY_TECH_RUN_GUARD|SKIPPED_BY_STATUS_BACKOFF)") { continue }
        $key = Get-ComputerListKey -ComputerName $computer
        if ([string]::IsNullOrWhiteSpace($key)) { $key = $computer.ToUpperInvariant() }
        $latestByComputer[$key] = $row
    }
    return @($latestByComputer.GetEnumerator() | Sort-Object Name | ForEach-Object { $_.Value })
}
function Set-HybridJoinLatestStatusMetadata {
    param([AllowEmptyCollection()][object[]]$Rows)

    $latestAttempt = @{}
    $latestActionable = @{}
    foreach ($row in @($Rows)) {
        if ($null -eq $row -or -not $row.PSObject.Properties["Computer"]) { continue }
        $computer = [string]$row.Computer
        if ([string]::IsNullOrWhiteSpace($computer)) { continue }
        $key = Get-ComputerListKey -ComputerName $computer
        $status = Get-HybridJoinHtmlEffectiveStatus -Row $row
        $nextAction = Get-HybridJoinHtmlEffectiveNextAction -Row $row
        $latestAttempt[$key] = [pscustomobject]@{ Status=$status; NextAction=$nextAction }
        if ($status -notmatch "^(CANCELLED_|RUN_GUARD_ACTIVE$|SKIPPED_BY_TECH_RUN_GUARD|SKIPPED_BY_STATUS_BACKOFF)") {
            $latestActionable[$key] = [pscustomobject]@{ Status=$status; NextAction=$nextAction }
        }
    }
    foreach ($row in @($Rows)) {
        if ($null -eq $row -or -not $row.PSObject.Properties["Computer"]) { continue }
        $key = Get-ComputerListKey -ComputerName ([string]$row.Computer)
        $attempt = if ($latestAttempt.ContainsKey($key)) { $latestAttempt[$key] } else { $null }
        $actionable = if ($latestActionable.ContainsKey($key)) { $latestActionable[$key] } else { $null }
        $row | Add-Member -NotePropertyName LatestAttemptStatus -NotePropertyValue $(if($attempt){[string]$attempt.Status}else{""}) -Force
        $row | Add-Member -NotePropertyName LatestActionableStatus -NotePropertyValue $(if($actionable){[string]$actionable.Status}else{""}) -Force
        $row | Add-Member -NotePropertyName LatestActionableNextAction -NotePropertyValue $(if($actionable){[string]$actionable.NextAction}else{""}) -Force
    }
}
function New-HybridJoinCycleProgressRows {
    param(
        [Parameter(Mandatory=$true)][int]$CycleNumber,
        [Parameter(Mandatory=$true)][datetime]$CycleStart,
        [Parameter(Mandatory=$true)][int]$TotalComputers,
        [Parameter(Mandatory=$true)][int]$QueuedComputers,
        [Parameter(Mandatory=$true)][int]$CompletedComputers,
        [Parameter(Mandatory=$true)][int]$RunningComputers,
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

function New-HybridJoinRunningJobRows {
    param(
        [AllowEmptyCollection()][object[]]$RunningJobs,
        [Parameter(Mandatory=$true)][hashtable]$JobStartedAtById
    )

    $now = Get-Date
    return @($RunningJobs | ForEach-Object {
        $jobId = [string]$_.Id
        $computer = ([string]$_.Name) -replace '^EHJIR_C\d+_',''
        $started = if ($JobStartedAtById.ContainsKey($jobId)) { [datetime]$JobStartedAtById[$jobId] } else { [datetime]::MinValue }
        [pscustomobject]@{
            Computer = $computer
            JobId = $jobId
            State = [string]$_.State
            Started = if ($started -gt [datetime]::MinValue) { $started.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
            ElapsedMinutes = if ($started -gt [datetime]::MinValue) { [math]::Round(($now - $started).TotalMinutes, 1) } else { '' }
        }
    })
}

function ConvertTo-SimpleHtmlTable {
    param(
        [Parameter(Mandatory=$true)][AllowEmptyCollection()][object[]]$Rows,
        [string[]]$Columns
    )

    if (-not $Rows -or $Rows.Count -eq 0) {
        return "<p>No rows.</p>"
    }

    if (-not $Columns -or $Columns.Count -eq 0) {
        $columnSet = New-Object System.Collections.Generic.List[string]
        foreach ($row in $Rows) {
            foreach ($property in $row.PSObject.Properties) {
                if (-not $columnSet.Contains($property.Name)) {
                    [void]$columnSet.Add($property.Name)
                }
            }
        }
        $Columns = @($columnSet)
    }

    $html = New-Object System.Collections.Generic.List[string]
    [void]$html.Add("<table>")
    [void]$html.Add("<tr>")
    foreach ($column in $Columns) {
        [void]$html.Add(("<th>{0}</th>" -f (ConvertTo-HtmlText $column)))
    }
    [void]$html.Add("</tr>")

    foreach ($row in $Rows) {
        [void]$html.Add("<tr>")
        foreach ($column in $Columns) {
            $value = ""
            $property = $row.PSObject.Properties[$column]
            if ($property) {
                $value = $property.Value
            }
            if ($column -eq 'Local log' -or $column -eq 'Collected logs' -or $column -eq 'Current run logs' -or $column -eq 'Remote PC logs') {
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

function Copy-RemoteEvidenceFolder {
    param(
        [Parameter(Mandatory=$true)][string]$RemoteDataPath,
        [Parameter(Mandatory=$true)][string]$DestinationPath,
        [Parameter(Mandatory=$true)][string]$ScriptName
    )

    $copyCount = 0

    function Copy-EvidenceFile {
        param(
            [Parameter(Mandatory=$true)][string]$SourceFile,
            [Parameter(Mandatory=$true)][string]$TargetFolder
        )

    try {
        if (-not (Test-Path -LiteralPath $TargetFolder)) {
            [System.IO.Directory]::CreateDirectory($TargetFolder) | Out-Null
        }
        Copy-Item -LiteralPath $SourceFile -Destination $TargetFolder -Force -ErrorAction Stop
        return $true
    }
        catch [System.Management.Automation.ItemNotFoundException] {
            return $false
        }
        catch [System.IO.FileNotFoundException] {
            return $false
        }
    catch [System.IO.DirectoryNotFoundException] {
        return $false
    }
    catch {
        return $false
    }
    }

    foreach ($folderName in @("Logs","Output","Transcripts")) {
        $sourceFolder = Join-Path $RemoteDataPath $folderName
        if (Test-Path -LiteralPath $sourceFolder) {
            $targetFolder = Join-Path $DestinationPath $folderName
            $files = @(Get-ChildItem -LiteralPath $sourceFolder -Recurse -File -Force -ErrorAction SilentlyContinue)
            foreach ($file in $files) {
                $relativePath = $file.FullName.Substring($sourceFolder.Length).TrimStart("\")
                $relativeDir = Split-Path -Parent $relativePath
                if ([string]::IsNullOrWhiteSpace($relativeDir)) {
                    $fileTargetFolder = $targetFolder
                }
                else {
                    $fileTargetFolder = Join-Path $targetFolder $relativeDir
                }

                if (Copy-EvidenceFile -SourceFile $file.FullName -TargetFolder $fileTargetFolder) {
                    $copyCount++
                }
            }
        }
    }

    Get-ChildItem -LiteralPath $RemoteDataPath -File -Force -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -ne $ScriptName -and
            $_.Extension -in @(".csv",".log",".txt",".html",".json",".xml",".evtx")
        } |
        ForEach-Object {
            if (-not (Test-Path -LiteralPath $DestinationPath)) {
                New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
            }
            if (Copy-EvidenceFile -SourceFile $_.FullName -TargetFolder $DestinationPath) {
                $copyCount++
            }
        }

    if ($copyCount -eq 0) {
        if (Test-Path -LiteralPath $DestinationPath) {
            Remove-Item -LiteralPath $DestinationPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        throw "No remote evidence files found to collect."
    }

    return $copyCount
}

$script:BrandLogoDataUri = $null
function Get-IntuneHybridJoinBrandLogoDataUri {
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

function New-CycleHtmlReport {
    param(
        [Parameter(Mandatory=$true)][AllowEmptyCollection()][object[]]$Summary,
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][int]$CycleNumber,
        [Parameter(Mandatory=$true)][datetime]$GeneratedAt,
        [switch]$IsLive,
        [AllowEmptyCollection()][object[]]$CycleProgress = @(),
        [AllowEmptyCollection()][object[]]$RunningJobRows = @()
    )

    $rows = @(Get-HybridJoinHtmlReportRows -Items @($Summary | ForEach-Object { $_ }))
    $latestRows = @(Get-HybridJoinLatestRowsByComputer -Rows $rows)
    $latestActionableRows = @(Get-HybridJoinLatestActionableRowsByComputer -Rows $rows)
    $separatedStatuses = @('ADMIN_SHARE_UNREACHABLE','DNS_PREFLIGHT_ALL_SAMPLES_FAILED','RUN_GUARD_ACTIVE','SKIPPED_BY_TECH_RUN_GUARD','SKIPPED_BY_TECH_RUN_GUARD_STARTED_NO_RESULT','SKIPPED_BY_STATUS_BACKOFF')
    $mainRows = @($latestActionableRows | Where-Object { $separatedStatuses -notcontains (Get-HybridJoinHtmlEffectiveStatus -Row $_) })
    $separatedRows = @($latestActionableRows | Where-Object { $separatedStatuses -contains (Get-HybridJoinHtmlEffectiveStatus -Row $_) })

    $effectiveRows = @($rows | ForEach-Object { [pscustomobject]@{ Status = Get-HybridJoinHtmlEffectiveStatus -Row $_; NextAction = Get-HybridJoinHtmlEffectiveNextAction -Row $_ } })
    $statusCounts = @($effectiveRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Status) } | Group-Object -Property Status | Sort-Object Count -Descending | ForEach-Object { [pscustomobject]@{ Status=$_.Name; Count=$_.Count } })
    $nextActionCounts = @($effectiveRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.NextAction) } | Group-Object -Property NextAction | Sort-Object Count -Descending | ForEach-Object { [pscustomobject]@{ NextAction=$_.Name; Count=$_.Count } })
    $adminShareFailureCounts = @($rows | Where-Object { (Get-HybridJoinHtmlEffectiveStatus -Row $_) -eq 'ADMIN_SHARE_UNREACHABLE' } | ForEach-Object { [pscustomobject]@{ FailureType = Get-HybridJoinAdminShareFailureType -Row $_ } } | Group-Object -Property FailureType | Sort-Object Count -Descending | ForEach-Object { [pscustomobject]@{ FailureType=$_.Name; Count=$_.Count } })
    $latestEffectiveRows = @($latestRows | ForEach-Object { [pscustomobject]@{ Status = Get-HybridJoinHtmlEffectiveStatus -Row $_; NextAction = Get-HybridJoinHtmlEffectiveNextAction -Row $_ } })
    $latestActionableEffectiveRows = @($latestActionableRows | ForEach-Object { [pscustomobject]@{ Status = Get-HybridJoinHtmlEffectiveStatus -Row $_; NextAction = Get-HybridJoinHtmlEffectiveNextAction -Row $_ } })
    $latestStatusCounts = @($latestEffectiveRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Status) } | Group-Object -Property Status | Sort-Object Count -Descending | ForEach-Object { [pscustomobject]@{ Status=$_.Name; Count=$_.Count } })
    $latestActionableStatusCounts = @($latestActionableEffectiveRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.Status) } | Group-Object -Property Status | Sort-Object Count -Descending | ForEach-Object { [pscustomobject]@{ Status=$_.Name; Count=$_.Count } })
    $latestActionableNextActionCounts = @($latestActionableEffectiveRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.NextAction) } | Group-Object -Property NextAction | Sort-Object Count -Descending | ForEach-Object { [pscustomobject]@{ NextAction=$_.Name; Count=$_.Count } })
    $latestAdminShareFailureCounts = @($latestActionableRows | Where-Object { (Get-HybridJoinHtmlEffectiveStatus -Row $_) -eq 'ADMIN_SHARE_UNREACHABLE' } | ForEach-Object { [pscustomobject]@{ FailureType = Get-HybridJoinAdminShareFailureType -Row $_ } } | Group-Object -Property FailureType | Sort-Object Count -Descending | ForEach-Object { [pscustomobject]@{ FailureType=$_.Name; Count=$_.Count } })

    $computerListStats = $script:CurrentCycleComputerListStats
    if ($CycleProgress.Count -eq 0) {
        $CycleProgress = @([pscustomobject]@{
            Cycle = $CycleNumber
            Generated = $GeneratedAt.ToString('yyyy-MM-dd HH:mm:ss')
            ComputerListLines = $(if ($computerListStats -and $computerListStats.PSObject.Properties['RawLines']) { [int]$computerListStats.RawLines } else { '' })
            TotalUnique = $(if ($script:CurrentCycleTotalComputers) { [int]$script:CurrentCycleTotalComputers } else { $rows.Count })
            ReportRows = $rows.Count
            DuplicateGroups = $(if ($computerListStats -and $computerListStats.PSObject.Properties['DuplicateGroups']) { [int]$computerListStats.DuplicateGroups } else { 0 })
            DuplicateLines = $(if ($computerListStats -and $computerListStats.PSObject.Properties['DuplicateLines']) { [int]$computerListStats.DuplicateLines } else { 0 })
            DuplicateSamples = $(if ($computerListStats -and $computerListStats.PSObject.Properties['DuplicateSamples']) { [string]$computerListStats.DuplicateSamples } else { '' })
        })
    }

    $identity = $null
    try { $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent() } catch { }
    $optionRows = @(
        [pscustomobject]@{ Category='Operator'; Option='TechnicianAccount'; Value=$(if ($identity) { [string]$identity.Name } else { [Environment]::UserName }) }
        [pscustomobject]@{ Category='Operator'; Option='TechnicianSID'; Value=$(if ($identity -and $identity.User) { [string]$identity.User.Value } else { '' }) }
        [pscustomobject]@{ Category='Operator'; Option='TechnicianAuthType'; Value=$(if ($identity) { [string]$identity.AuthenticationType } else { '' }) }
        [pscustomobject]@{ Category='Operator'; Option='TechnicianComputer'; Value=$env:COMPUTERNAME }
        [pscustomobject]@{ Category='Mode'; Option='DryRun'; Value=[string][bool]$DryRun }
        [pscustomobject]@{ Category='Mode'; Option='AuditOnly'; Value=[string][bool]$AuditOnly }
        [pscustomobject]@{ Category='Mode'; Option='RunOnce'; Value=[string][bool]$RunOnce }
        [pscustomobject]@{ Category='Mode'; Option='IgnoreRunGuard'; Value=[string][bool]$IgnoreRunGuard }
        [pscustomobject]@{ Category='Mode'; Option='IgnoreRunGuardEveryCycle'; Value=[string][bool]$IgnoreRunGuardEveryCycle }
        [pscustomobject]@{ Category='Mode'; Option='UseTechnicianRunGuardHistory'; Value=[string][bool]$UseTechnicianRunGuardHistory }
        [pscustomobject]@{ Category='Mode'; Option='IgnoreTechnicianRunGuardHistory'; Value=[string][bool]$IgnoreTechnicianRunGuardHistory }
        [pscustomobject]@{ Category='Mode'; Option='EffectiveTechnicianRunGuardHistory'; Value=[string][bool]$script:UseEffectiveTechnicianRunGuardHistory }
        [pscustomobject]@{ Category='Timing'; Option='TechnicianRunGuardHours'; Value=[string]$TechnicianRunGuardHours }
        [pscustomobject]@{ Category='Mode'; Option='SkipVirtualMachines'; Value=[string][bool]$SkipVirtualMachines }
        [pscustomobject]@{ Category='Mode'; Option='DisableNightPause'; Value=[string][bool]$DisableNightPause }
        [pscustomobject]@{ Category='Actions'; Option='AllowDsregLeave'; Value=[string][bool]$AllowDsregLeave }
        [pscustomobject]@{ Category='Actions'; Option='AllowRebootWhenNoInteractiveUser'; Value=[string][bool]$AllowRebootWhenNoInteractiveUser }
        [pscustomobject]@{ Category='Actions'; Option='AllowRebootAfterDsregLeave'; Value=[string][bool]$AllowRebootAfterDsregLeave }
        [pscustomobject]@{ Category='Actions'; Option='AllowRemoveNonIntuneMdmEnrollment'; Value=[string][bool]$AllowRemoveNonIntuneMdmEnrollment }
        [pscustomobject]@{ Category='Actions'; Option='AllowRemoveStaleIntuneEnrollment'; Value=[string][bool]$AllowRemoveStaleIntuneEnrollment }
        [pscustomobject]@{ Category='Timing'; Option='RebootDelaySeconds'; Value=[string]$RebootDelaySeconds }
        [pscustomobject]@{ Category='Timing'; Option='RetryAfterRebootDelaySeconds'; Value=[string]$RetryAfterRebootDelaySeconds }
        [pscustomobject]@{ Category='Timing'; Option='RetryAfterRebootMaxAttempts'; Value=[string]$RetryAfterRebootMaxAttempts }
        [pscustomobject]@{ Category='Timing'; Option='PsExecTimeoutMinutes'; Value=[string]$PsExecTimeoutMinutes }
        [pscustomobject]@{ Category='Timing'; Option='CommunicationLostEvidenceWaitMinutes'; Value=[string]$CommunicationLostEvidenceWaitMinutes }
        [pscustomobject]@{ Category='Timing'; Option='JobPollSeconds'; Value=[string]$JobPollSeconds }
        [pscustomobject]@{ Category='Parallelism'; Option='ThrottleLimit'; Value=[string]$ThrottleLimit }
        [pscustomobject]@{ Category='Parallelism'; Option='GlobalConcurrencyLimit'; Value=[string]$GlobalConcurrencyLimit }
        [pscustomobject]@{ Category='Paths'; Option='ComputerListPath'; Value=[string]$ComputerListPath }
        [pscustomobject]@{ Category='Paths'; Option='LogRoot'; Value=[string]$LogRoot }
        [pscustomobject]@{ Category='Paths'; Option='ReportRoot'; Value=[string]$ReportRoot }
        [pscustomobject]@{ Category='Paths'; Option='CentralLogRoot'; Value=[string]$CentralLogRoot }
        [pscustomobject]@{ Category='Paths'; Option='CentralLogCollectionMode'; Value=[string]$CentralLogCollectionMode }
        [pscustomobject]@{ Category='Paths'; Option='TechnicianRunGuardHistoryPath'; Value=[string]$script:TechnicianRunGuardHistoryPath }
        [pscustomobject]@{ Category='Paths'; Option='LauncherLogPath'; Value=[string]$script:LauncherLogPath }
    )

    $logoUri = Get-IntuneHybridJoinBrandLogoDataUri
    $logoHtml = if (-not [string]::IsNullOrWhiteSpace($logoUri)) { "<img class='logo' src='$logoUri' alt='WorkplaceCloudHub' />" } else { "" }
    $mode = if ($IsLive) { 'LIVE' } else { 'FINAL' }
    $lotName = if (-not [string]::IsNullOrWhiteSpace([string]$LotRoot)) { Split-Path -Leaf $LotRoot } else { 'Unknown LOT' }
    $cycleValues = @($rows | ForEach-Object { if ($_.PSObject.Properties['Cycle']) { [string]$_.Cycle } } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    $cycleLabel = if ($CycleNumber -le 0) { 'all cycles' } elseif ($CycleNumber -gt 1 -or $cycleValues.Count -gt 1) { "cycles 1-$CycleNumber" } else { "cycle $CycleNumber" }

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
.card { background: #fff; border: 1px solid #DDE7F0; border-radius: 8px; padding: 14px 18px; margin-bottom: 16px; overflow-x: auto; }
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
    [void]$html.Add(("<html><head><meta charset='utf-8'>{0}<title>Intune Hybrid Join repair - {1}</title></head><body>" -f $style,(ConvertTo-HtmlText $cycleLabel)))
    [void]$html.Add("<div class='header-card'><div class='header-text'>")
    [void]$html.Add(("<div class='title'>Intune Hybrid Join repair - {0}<span class='badge'>{1}</span></div>" -f (ConvertTo-HtmlText $cycleLabel),$mode))
    [void]$html.Add("<div class='subtitle'>Smart Intune Hybrid Join Toolkit</div>")
    [void]$html.Add(("<div class='lot-name' title='{1}'>LOT: {0}</div>" -f (ConvertTo-HtmlText $lotName),(ConvertTo-HtmlText $LotRoot)))
    [void]$html.Add(("<div class='meta'>Generated: {0} | Report rows: {1} | Unique computers: {2} | Launcher: v{3}</div>" -f (ConvertTo-HtmlText $GeneratedAt.ToString('yyyy-MM-dd HH:mm:ss')),$rows.Count,$latestRows.Count,(ConvertTo-HtmlText $LauncherVersion)))
    [void]$html.Add(("<div class='meta'>Launcher log: {0}</div>" -f (New-HtmlLogLink -Path $script:LauncherLogPath)))
    [void]$html.Add("</div>$logoHtml</div>")

    [void]$html.Add("<div class='card'><h2>Cycle progress</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $CycleProgress -Columns @('Cycle','Started','ElapsedMinutes','ComputerListLines','TotalUnique','Queued','CompletedRows','Running','Remaining','DuplicateGroups','DuplicateLines','DuplicateSamples')))
    [void]$html.Add("</div>")
    if ($RunningJobRows -and $RunningJobRows.Count -gt 0) {
        [void]$html.Add("<div class='card'><h2>Running jobs</h2>")
        [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $RunningJobRows -Columns @('Computer','JobId','State','Started','ElapsedMinutes')))
        [void]$html.Add("</div>")
    }
    [void]$html.Add("<div class='card'><h2>Latest status by unique computer</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $latestStatusCounts -Columns @('Status','Count')))
    [void]$html.Add("</div>")
    [void]$html.Add("<div class='card'><h2>Latest actionable status by unique computer</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $latestActionableStatusCounts -Columns @('Status','Count')))
    [void]$html.Add("<div class='footer'>Cancellation, run-guard and status-backoff rows do not mask the last meaningful endpoint or connectivity result.</div></div>")
    [void]$html.Add("<div class='card'><h2>Status summary by attempts</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $statusCounts -Columns @('Status','Count')))
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
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $nextActionCounts -Columns @('NextAction','Count')))
    [void]$html.Add("</div>")
    [void]$html.Add("<div class='card'><h2>Latest actionable next action by unique computer</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $latestActionableNextActionCounts -Columns @('NextAction','Count')))
    [void]$html.Add("</div>")

    $detailColumns = @('Cycle','Timestamp','Computer','ConnectionTarget','Status','EffectiveStatus','NextAction','EffectiveNextAction','BackoffStatus','BackoffUntilUtc','BackoffCount','RemoteStatus','RemoteExitCode','PsExecExitCode','RemoteDetail','RetryAfterRebootAction','RetryAfterRebootDetail','RetryAfterRebootAttempt','RetryAfterRebootMaxAttempts','RetryAfterRebootTaskName','Local log','Collected logs','Current run logs','Remote PC logs','InteractiveUserAccountName','InteractiveUserAccountType','InteractiveSessionName','InteractiveSessionState','UserIsUserAzureAD','UserAzureAdPrt','UserSessionIsNotRemote','ErrorMessage','JobErrorMessage','IntuneInventoryPresent','EntraInventoryPresent','EntraRegisteredState','EntraAlternativeSecurityIdCount','EntraPendingReason','EntraRegistrationDateTime','EntraTrustType','EntraDeviceId','EntraObjectId','ADInventoryPresent','ADDomain','ADEnabled','ADDNSHostName','ADDistinguishedName','ADOperatingSystem','ADLastLogonTimestampUtc','PostCycleIntuneInventoryChecked','PostCycleIntuneInventoryPresent','PostCycleIntuneEnrollmentDetected','PostCycleIntuneInventoryCsv','PostCycleIntuneInventoryError','PostCycleEntraInventoryChecked','PostCycleEntraInventoryPresent','PostCycleEntraRegisteredState','PostCycleEntraAlternativeSecurityIdCount','PostCycleEntraPendingResolved','PostCycleEntraInventoryCsv','PostCycleEntraInventoryError','PostCycleADInventoryChecked','PostCycleADInventoryPresent','PostCycleADInventoryCsv','PostCycleADInventoryError','AdminShareReachable','AdminShareFailureType','PingReachable','DnsResolved','RemoteLogsCollected','RemoteLogsPath','RemoteCurrentRunLogsPath','LogPath')
    [void]$html.Add("<div class='card'><h2>Latest actionable computer details</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $mainRows -Columns $detailColumns))
    [void]$html.Add("<div class='footer'>Smart Intune Hybrid Join Toolkit - <a href='https://workplacecloudhub.com'>workplacecloudhub.com</a></div></div>")
    if ($separatedRows.Count -gt 0) {
        [void]$html.Add("<div class='card'><h2>Run guard / admin share details</h2>")
        [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $separatedRows -Columns $detailColumns))
        [void]$html.Add("<div class='footer'>Rows excluded from Computer details: connectivity failures and active target/technician run guards.</div></div>")
    }
    [void]$html.Add("<div class='card'><h2>LOT/run options</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $optionRows -Columns @('Category','Option','Value')))
    [void]$html.Add("</div>")
    [void]$html.Add("<div class='card'><h2>Security evidence</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows @($script:PsExecSecurityEvidenceRows) -Columns @('Field','Value')))
    [void]$html.Add("<div class='footer'>Non-dry-run PsExec execution is blocked unless the binary is Microsoft-signed and named PsExec.exe or PsExec64.exe.</div></div>")
    [void]$html.Add("</body></html>")
    [System.IO.File]::WriteAllText($Path, ($html -join "`r`n"), [System.Text.UTF8Encoding]::new($false))
}
$scriptArgsBase = @()
if ($AllowDsregLeave) { $scriptArgsBase += "-AllowDsregLeave" }
if ($AllowRebootWhenNoInteractiveUser) { $scriptArgsBase += "-AllowRebootWhenNoInteractiveUser" }
if ($AllowRebootAfterDsregLeave) { $scriptArgsBase += "-AllowRebootAfterDsregLeave" }
if ($AllowRemoveNonIntuneMdmEnrollment) { $scriptArgsBase += "-AllowRemoveNonIntuneMdmEnrollment" }
if ($AllowRemoveStaleIntuneEnrollment) { $scriptArgsBase += "-AllowRemoveStaleIntuneEnrollment" }
if ($SkipVirtualMachines) { $scriptArgsBase += "-SkipVirtualMachines" }
if ($AuditOnly) { $scriptArgsBase += "-AuditOnly" }
$scriptArgsBase += "-StaleCleanupDelaySeconds"
$scriptArgsBase += $StaleCleanupDelaySeconds
$scriptArgsBase += "-RebootDelaySeconds"
$scriptArgsBase += $RebootDelaySeconds
$scriptArgsBase += "-IntuneRetrySleepMinutes"
$scriptArgsBase += $IntuneRetrySleepMinutes
$scriptArgsBase += "-IntuneRetryMaxRetries"
$scriptArgsBase += $IntuneRetryMaxRetries
$scriptArgsBase += "-RetryAfterRebootDelaySeconds"
$scriptArgsBase += $RetryAfterRebootDelaySeconds
$scriptArgsBase += "-RetryAfterRebootMaxAttempts"
$scriptArgsBase += $RetryAfterRebootMaxAttempts

$IntuneInventorySet = @{}
if (-not [string]::IsNullOrWhiteSpace($IntuneInventoryCsv)) {
    $refreshInitialInventory = $false
    $initialInventoryReason = ""
    $intuneInventoryItem = Get-Item -LiteralPath $IntuneInventoryCsv -ErrorAction SilentlyContinue

    if ($null -eq $intuneInventoryItem) {
        $refreshInitialInventory = $true
        $initialInventoryReason = "missing"
    }
    else {
        $intuneInventoryAge = (Get-Date) - $intuneInventoryItem.LastWriteTime
        if ($intuneInventoryAge.TotalMinutes -gt 120) {
            $refreshInitialInventory = $true
            $initialInventoryReason = ("older than 120 minutes; LastWriteTime={0}; Age={1:N1} minute(s)" -f $intuneInventoryItem.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"), $intuneInventoryAge.TotalMinutes)
        }
    }

    if ($refreshInitialInventory -and $DryRun) {
        Write-Host ("DryRun: Intune inventory CSV is {0}; skipping automatic Graph inventory export." -f $initialInventoryReason) -ForegroundColor Yellow
    }
    elseif ($refreshInitialInventory) {
        $initialInventoryStamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $initialInventoryOutputPath = Join-Path $ReportRoot ("DevicesIntune_InitialScoped_{0}.csv" -f $initialInventoryStamp)
        $initialInventoryLogPath = Join-Path $ReportRoot ("DevicesIntune_InitialScoped_{0}.log" -f $initialInventoryStamp)
        Write-Host ("Intune inventory CSV is {0}. Running a LOT-scoped Graph inventory export before starting..." -f $initialInventoryReason) -ForegroundColor Yellow
        $initialInventory = Invoke-FullIntuneInventoryExport `
            -ExportScriptPath $ExportIntuneScriptPath `
            -OutputPath $initialInventoryOutputPath `
            -LogPath $initialInventoryLogPath `
            -PageSize $PostCycleIntuneInventoryPageSize `
            -ComputerListPath $ComputerListPath

        if ($initialInventory.Success) {
            $IntuneInventoryCsv = $initialInventory.CsvPath
            $IntuneInventorySet = $initialInventory.InventorySet
            Write-Host ("Initial Intune inventory refreshed. Devices={0}; CSV={1}" -f $IntuneInventorySet.Count,$initialInventory.CsvPath) -ForegroundColor Green
        }
        else {
            Write-Host ("WARNING: Initial Intune inventory refresh failed: {0}" -f $initialInventory.Error) -ForegroundColor Yellow
            if (Test-Path -LiteralPath $IntuneInventoryCsv) {
                Write-Host "Continuing with existing Intune CSV despite refresh failure." -ForegroundColor Yellow
                $IntuneInventorySet = Get-IntuneInventorySet -Path $IntuneInventoryCsv -NameColumn $IntuneInventoryNameColumn
            }
        }
    }
    else {
        $IntuneInventorySet = Get-IntuneInventorySet -Path $IntuneInventoryCsv -NameColumn $IntuneInventoryNameColumn
    }
}

$EntraInventoryMap = @{}
if (-not [string]::IsNullOrWhiteSpace($EntraInventoryCsv)) {
    try {
        $refreshInitialEntraInventory = $false
        $initialEntraInventoryReason = ""
        $entraInventoryItem = Get-Item -LiteralPath $EntraInventoryCsv -ErrorAction SilentlyContinue

        if ($null -eq $entraInventoryItem) {
            $refreshInitialEntraInventory = $true
            $initialEntraInventoryReason = "missing"
        }
        else {
            $entraInventoryAge = (Get-Date) - $entraInventoryItem.LastWriteTime
            if ($entraInventoryAge.TotalMinutes -gt 120) {
                $refreshInitialEntraInventory = $true
                $initialEntraInventoryReason = ("older than 120 minutes; LastWriteTime={0}; Age={1:N1} minute(s)" -f $entraInventoryItem.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"), $entraInventoryAge.TotalMinutes)
            }
        }

        if ($refreshInitialEntraInventory -and $DryRun) {
            Write-Host ("DryRun: Entra inventory CSV is {0}; skipping automatic Graph device export." -f $initialEntraInventoryReason) -ForegroundColor Yellow
        }
        elseif ($refreshInitialEntraInventory) {
            $initialEntraInventoryStamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $initialEntraInventoryOutputPath = Join-Path $ReportRoot ("DevicesEntra_InitialScoped_{0}.csv" -f $initialEntraInventoryStamp)
            $initialEntraInventoryLogPath = Join-Path $ReportRoot ("DevicesEntra_InitialScoped_{0}.log" -f $initialEntraInventoryStamp)
            Write-Host ("Entra inventory CSV is {0}. Running a LOT-scoped Graph device export before starting..." -f $initialEntraInventoryReason) -ForegroundColor Yellow
            $initialEntraInventory = Invoke-FullEntraInventoryExport `
                -ExportScriptPath $ExportEntraScriptPath `
                -OutputPath $initialEntraInventoryOutputPath `
                -LogPath $initialEntraInventoryLogPath `
                -PageSize $PostCycleIntuneInventoryPageSize `
                -ComputerListPath $ComputerListPath

            if ($initialEntraInventory.Success) {
                $EntraInventoryCsv = $initialEntraInventory.CsvPath
                $EntraInventoryMap = $initialEntraInventory.InventoryMap
                Write-Host ("Initial Entra inventory refreshed. Devices={0}; CSV={1}" -f $EntraInventoryMap.Count,$initialEntraInventory.CsvPath) -ForegroundColor Green
            }
            else {
                Write-Host ("WARNING: Initial Entra inventory refresh failed: {0}" -f $initialEntraInventory.Error) -ForegroundColor Yellow
                if (Test-Path -LiteralPath $EntraInventoryCsv) {
                    Write-Host "Continuing with existing Entra CSV despite refresh failure." -ForegroundColor Yellow
                    $EntraInventoryMap = Get-EntraInventoryMap -Path $EntraInventoryCsv -NameColumn $EntraInventoryNameColumn
                }
            }
        }
        elseif (Test-Path -LiteralPath $EntraInventoryCsv) {
            $EntraInventoryMap = Get-EntraInventoryMap -Path $EntraInventoryCsv -NameColumn $EntraInventoryNameColumn
        }
    }
    catch {
        Write-Host ("WARN: Could not load Entra inventory CSV '{0}': {1}" -f $EntraInventoryCsv,$_.Exception.Message) -ForegroundColor Yellow
        $EntraInventoryMap = @{}
    }
}

$AdInventoryMap = @{}
$script:AdInventoryLastRefreshUtc = $null
if (-not [string]::IsNullOrWhiteSpace($AdInventoryCsv)) {
    try {
        $refreshInitialAdInventory = $false
        $initialAdInventoryReason = ""
        $adInventoryItem = Get-Item -LiteralPath $AdInventoryCsv -ErrorAction SilentlyContinue

        if ($null -eq $adInventoryItem) {
            $refreshInitialAdInventory = $true
            $initialAdInventoryReason = "missing"
        }
        else {
            $adInventoryAge = (Get-Date) - $adInventoryItem.LastWriteTime
            if ($adInventoryAge.TotalHours -gt $AdInventoryFreshnessHours) {
                $refreshInitialAdInventory = $true
                $initialAdInventoryReason = ("older than {0} hour(s); LastWriteTime={1}; Age={2:N1} hour(s)" -f $AdInventoryFreshnessHours, $adInventoryItem.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"), $adInventoryAge.TotalHours)
            }
        }

        if ($AdInventoryUsesRecentRootCsv) {
            Write-Host ("AD forest inventory CSV is recent. Using root CSV in priority: {0}" -f $AdInventoryCsv) -ForegroundColor Green
            $AdInventoryMap = Get-AdInventoryMap -Path $AdInventoryCsv -NameColumn $AdInventoryNameColumn
            $script:AdInventoryLastRefreshUtc = $adInventoryItem.LastWriteTimeUtc
        }
        elseif ($refreshInitialAdInventory -and $DryRun) {
            Write-Host ("DryRun: AD inventory CSV is {0}; skipping automatic AD computer export." -f $initialAdInventoryReason) -ForegroundColor Yellow
        }
        elseif ($refreshInitialAdInventory) {
            $initialAdInventoryStamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $initialAdInventoryOutputPath = Join-Path $ReportRoot ("DevicesAD_InitialScoped_{0}.csv" -f $initialAdInventoryStamp)
            $initialAdInventoryLogPath = Join-Path $ReportRoot ("DevicesAD_InitialScoped_{0}.log" -f $initialAdInventoryStamp)
            $initialAdScope = if ([string]::IsNullOrWhiteSpace($AdDomain)) { "forest" } else { "domain '$AdDomain'" }
            Write-Host ("AD inventory CSV is {0}. Running a LOT-scoped AD computer export before starting. Scope={1}..." -f $initialAdInventoryReason,$initialAdScope) -ForegroundColor Yellow
            $initialAdInventory = Invoke-FullAdInventoryExport `
                -ExportScriptPath $ExportAdScriptPath `
                -OutputPath $initialAdInventoryOutputPath `
                -LogPath $initialAdInventoryLogPath `
                -Domain $AdDomain `
                -ComputerListPath $ComputerListPath

            if ($initialAdInventory.Success) {
                $AdInventoryCsv = $initialAdInventory.CsvPath
                $AdInventoryMap = $initialAdInventory.InventoryMap
                $script:AdInventoryLastRefreshUtc = [datetime]::UtcNow
                Write-Host ("Initial AD inventory refreshed. Devices={0}; CSV={1}" -f $AdInventoryMap.Count,$initialAdInventory.CsvPath) -ForegroundColor Green
            }
            else {
                Write-Host ("WARNING: Initial AD inventory refresh failed: {0}" -f $initialAdInventory.Error) -ForegroundColor Yellow
                if (Test-Path -LiteralPath $AdInventoryCsv) {
                    Write-Host "Continuing with existing AD CSV despite refresh failure." -ForegroundColor Yellow
                    $AdInventoryMap = Get-AdInventoryMap -Path $AdInventoryCsv -NameColumn $AdInventoryNameColumn
                    $script:AdInventoryLastRefreshUtc = (Get-Item -LiteralPath $AdInventoryCsv).LastWriteTimeUtc
                }
            }
        }
        elseif (Test-Path -LiteralPath $AdInventoryCsv) {
            $AdInventoryMap = Get-AdInventoryMap -Path $AdInventoryCsv -NameColumn $AdInventoryNameColumn
            $script:AdInventoryLastRefreshUtc = (Get-Item -LiteralPath $AdInventoryCsv).LastWriteTimeUtc
        }
    }
    catch {
        Write-Host ("WARN: Could not load AD inventory CSV '{0}': {1}" -f $AdInventoryCsv,$_.Exception.Message) -ForegroundColor Yellow
        $AdInventoryMap = @{}
    }
}

$localScriptVersionForDisplay = Get-ScriptVersionFromFile -Path $LocalScriptPath
$localScriptHashForDisplay = Get-FileSha256 -Path $LocalScriptPath

Write-Host "Remote repair launcher" -ForegroundColor Cyan
Write-Host "Launcher ver : $LauncherVersion"
Write-Host "Script       : $LocalScriptPath"
Write-Host "Script ver   : $localScriptVersionForDisplay"
Write-Host "Script hash  : $localScriptHashForDisplay"
Write-Host "Remote path  : $RemoteScriptPath"
Write-Host "Computers   : $ComputerListPath"
Write-Host "PsExec      : $PsExecPath"
Write-Host ("PsExec sec. : Status={0}; SHA256={1}; Signature={2}; Signer={3}; Version={4}" -f $script:PsExecSecurityEvidence.CheckStatus,$script:PsExecSecurityEvidence.SHA256,$script:PsExecSecurityEvidence.SignatureStatus,$script:PsExecSecurityEvidence.SignerSubject,$script:PsExecSecurityEvidence.FileVersion)
Write-Host "Script args : $($scriptArgsBase -join ' ')"
Write-Host "Dry run     : $([bool]$DryRun)"
Write-Host "Skip VMs    : $([bool]$SkipVirtualMachines)"
Write-Host "Audit only  : $([bool]$AuditOnly)"
Write-Host "Intune CSV  : $IntuneInventoryCsv"
Write-Host "Entra CSV   : $EntraInventoryCsv"
Write-Host "AD CSV      : $AdInventoryCsv"
Write-Host "AD domain   : $AdDomain"
Write-Host "AD root CSV : $AdRootInventoryCsv"
Write-Host "Ignore guard: $([bool]$IgnoreRunGuard); Every cycle: $([bool]$IgnoreRunGuardEveryCycle)"
Write-Host "Parallelism : ThrottleLimit=$ThrottleLimit; GlobalConcurrencyLimit=$GlobalConcurrencyLimit; GlobalLeaseTimeout=$GlobalConcurrencyLeaseTimeoutMinutes minute(s); JobPollSeconds=$JobPollSeconds"
if ($GlobalConcurrencyLimit -gt 0) {
    Write-Host "Global gate : $GlobalConcurrencySemaphoreName"
}
Write-Host "Start delay : $DelayBetweenComputersSeconds seconds between job starts"
Write-Host "Loop        : $(-not [bool]$RunOnce); Delay between cycles: $DelayBetweenCyclesMinutes minute(s); Max cycles: $MaxCycles"
Write-Host "Night pause : Enabled=$(-not [bool]$DisableNightPause); Window=$($NightPauseStartHour):00-$($NightPauseEndHour):00 local time"
Write-Host "Logs        : $LogRoot"
Write-Host "Reports     : $ReportRoot"
Write-Host "Central logs: Enabled=$CollectRemoteLogs; Path=$CentralLogRoot; History=$([bool]$KeepCentralLogHistory); Mode=$CentralLogCollectionMode"
Write-Host "Reboot delay: $RebootDelaySeconds seconds"
Write-Host "Stale delay : $StaleCleanupDelaySeconds seconds"
Write-Host "Intune wait : $IntuneRetryMaxRetries retry(ies) x $IntuneRetrySleepMinutes minute(s) = $($IntuneRetryMaxRetries * $IntuneRetrySleepMinutes) minute(s)"
Write-Host "Reboot retry: delay $RetryAfterRebootDelaySeconds second(s); max attempts $RetryAfterRebootMaxAttempts"
Write-Host "PsExec wait : $(if ($PsExecTimeoutMinutes -eq 0) { 'No timeout' } else { "$PsExecTimeoutMinutes minute(s) max per computer" })"
Write-Host "Lost PsExec : poll every $CommunicationLostEvidencePollMinutes minute(s), max $CommunicationLostEvidenceWaitMinutes minute(s), before delayed evidence collection"
Write-Host "Post Intune : Enabled=$(-not [bool]$SkipPostCycleIntuneInventory); Mode=LOT-scoped Graph inventory; PageSize=$PostCycleIntuneInventoryPageSize; Export=$ExportIntuneScriptPath"
Write-Host "Post Entra  : Enabled=$(-not [string]::IsNullOrWhiteSpace($EntraInventoryCsv)); Mode=LOT-scoped Graph device inventory; PageSize=$PostCycleIntuneInventoryPageSize; Export=$ExportEntraScriptPath"
Write-Host "Post AD     : Enabled=$(-not $AdInventoryUsesRecentRootCsv -and -not [string]::IsNullOrWhiteSpace($AdInventoryCsv)); Mode=LOT-scoped AD computer inventory; Export=$ExportAdScriptPath"
Write-Host ""

$globalConcurrencyGateRoot = Join-Path ([System.IO.Path]::GetTempPath()) "SmartM365\GlobalWorkerGates"
$globalConcurrencyGateName = ($GlobalConcurrencySemaphoreName -replace '[^A-Za-z0-9_.-]', '_')
if ($globalConcurrencyGateName.Length -gt 120) { $globalConcurrencyGateName = $globalConcurrencyGateName.Substring(0, 120) }
$globalConcurrencyGatePath = Join-Path $globalConcurrencyGateRoot $globalConcurrencyGateName
$globalConcurrencyMutexName = "Local\SmartM365_GlobalWorkerGate_$globalConcurrencyGateName"
$launcherInstanceId = [guid]::NewGuid().ToString("N")
if ($GlobalConcurrencyLimit -gt 0) {
    New-Item -ItemType Directory -Path $globalConcurrencyGatePath -Force -ErrorAction Stop | Out-Null
    Write-Host ("Global lease gate: Limit={0}; Path={1}; LeaseTimeout={2} minute(s)" -f $GlobalConcurrencyLimit,$globalConcurrencyGatePath,$GlobalConcurrencyLeaseTimeoutMinutes) -ForegroundColor Green
}

$script:globalConcurrencyMutex = $null
if ($GlobalConcurrencyLimit -gt 0) {
    $script:globalConcurrencyMutex = New-Object System.Threading.Mutex($false, $globalConcurrencyMutexName)
}

function Invoke-WithGlobalGateMutex {
    param(
        [Parameter(Mandatory=$true)][string]$MutexName,
        [Parameter(Mandatory=$true)][scriptblock]$ScriptBlock
    )

    $sharedMutex = $script:globalConcurrencyMutex
    $mutex = if ($sharedMutex -ne $null) { $sharedMutex } else { New-Object System.Threading.Mutex($false, $MutexName) }
    $ownMutex = ($sharedMutex -eq $null)
    $acquired = $false
    try {
        try {
            $acquired = $mutex.WaitOne(30000)
        }
        catch [System.Threading.AbandonedMutexException] {
            $acquired = $true
        }

        if (-not $acquired) {
            throw "Could not acquire global gate mutex within 30 seconds: $MutexName"
        }

        & $ScriptBlock
    }
    finally {
        if ($acquired) {
            try { $mutex.ReleaseMutex() } catch { }
        }
        if ($ownMutex) {
            $mutex.Dispose()
        }
    }
}

function Test-GlobalGateProcessAlive {
    param([Parameter(Mandatory=$true)][int]$ProcessId)

    try {
        $null = Get-Process -Id $ProcessId -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Read-GlobalLeaseData {
    param([Parameter(Mandatory=$true)][string]$Path)

    $lastError = $null
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try { return (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop) }
        catch {
            $lastError = $_
            Start-Sleep -Milliseconds ([math]::Min(1000,100 * $attempt))
        }
    }
    if ($lastError) { throw $lastError }
    throw ("Failed to read global worker lease: {0}" -f $Path)
}

function Save-GlobalLeaseData {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)]$Data,
        [ValidateRange(1,20)][int]$Depth = 4
    )

    $parent = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    $json = $Data | ConvertTo-Json -Depth $Depth
    $lastError = $null
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        $tempPath = Join-Path $parent (".{0}.{1}.tmp" -f (Split-Path -Leaf $Path),[guid]::NewGuid().ToString("N"))
        try {
            Set-Content -LiteralPath $tempPath -Value $json -Encoding UTF8 -Force -ErrorAction Stop
            Move-Item -LiteralPath $tempPath -Destination $Path -Force -ErrorAction Stop
            return
        }
        catch {
            $lastError = $_
            Start-Sleep -Milliseconds ([math]::Min(1000,100 * $attempt))
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

function Remove-GlobalWorkerLeaseFileUnlocked {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [ValidateRange(1,20)][int]$Attempts = 5
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
        Start-Sleep -Milliseconds ([math]::Min(1000,100 * $attempt))
    }

    return (-not (Test-Path -LiteralPath $Path -PathType Leaf))
}

function Remove-StaleGlobalWorkerLeases {
    param(
        [Parameter(Mandatory=$true)][string]$GatePath,
        [Parameter(Mandatory=$true)][int]$LeaseTimeoutMinutes
    )

    $nowUtc = (Get-Date).ToUniversalTime()
    $removed = New-Object System.Collections.Generic.List[pscustomobject]
    foreach ($lease in @(Get-ChildItem -LiteralPath $GatePath -Filter '*.json' -File -ErrorAction SilentlyContinue)) {
        $remove = $false
        $reason = ""
        $computerName = ""
        $launcherPid = 0
        $workerPid = 0
        try {
            $data = Read-GlobalLeaseData -Path $lease.FullName
            $computerName = [string]$data.Computer
            $launcherPid = if ($data.PSObject.Properties["LauncherProcessId"]) { [int]$data.LauncherProcessId } else { [int]$data.ProcessId }
            $workerPid = if ($data.PSObject.Properties["WorkerProcessId"]) { [int]$data.WorkerProcessId } else { 0 }
            $createdUtc = [datetime]::Parse($data.CreatedUtc, $null, [System.Globalization.DateTimeStyles]::RoundtripKind)
            if (-not (Test-GlobalGateProcessAlive -ProcessId $launcherPid)) {
                $remove = $true; $reason = "LauncherProcessExited"
            } elseif ($workerPid -gt 0 -and -not (Test-GlobalGateProcessAlive -ProcessId $workerPid)) {
                $remove = $true; $reason = "WorkerProcessExited"
            } elseif (($nowUtc - $createdUtc).TotalMinutes -gt $LeaseTimeoutMinutes) {
                $remove = $true; $reason = "LeaseExpired"
            }
        } catch {
            $remove = $true; $reason = "InvalidLease"
        }
        if ($remove) {
            if (Remove-GlobalWorkerLeaseFileUnlocked -Path $lease.FullName) {
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
        $computerPart = if ($shortNames.Count -le 4) { ": " + ($shortNames -join ", ") } else { "" }
        Write-Host ("Removed {0} stale global worker lease(s) [{1}]{2}" -f $removed.Count,($pidParts -join "; "),$computerPart) -ForegroundColor DarkYellow
    }
}

function Update-GlobalWorkerLease {
    param(
        [AllowNull()][string]$LeasePath,
        [hashtable]$Properties = @{}
    )

    if ([string]::IsNullOrWhiteSpace($LeasePath) -or -not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return }

    Invoke-WithGlobalGateMutex -MutexName $globalConcurrencyMutexName -ScriptBlock {
        if (-not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return }
        $data = Read-GlobalLeaseData -Path $LeasePath
        foreach ($key in @($Properties.Keys)) {
            $data | Add-Member -NotePropertyName $key -NotePropertyValue $Properties[$key] -Force
        }
        $data | Add-Member -NotePropertyName LastUpdatedUtc -NotePropertyValue ((Get-Date).ToUniversalTime().ToString("o")) -Force
        Save-GlobalLeaseData -Path $LeasePath -Data $data -Depth 4
    }
}

function Acquire-GlobalWorkerLease {
    param(
        [Parameter(Mandatory=$true)][string]$Computer,
        [Parameter(Mandatory=$true)][int]$CycleNumber
    )

    if ($GlobalConcurrencyLimit -lt 1) { return "" }

    $waitStarted = Get-Date
    $lastWaitLog = $null
    $waitLogDelaySeconds = 300
    $waitLogIntervalSeconds = 300
    $waitWasLogged = $false

    while ($true) {
        if ((Get-LotCancellationState).Requested) { return '' }
        $leasePath = Invoke-WithGlobalGateMutex -MutexName $globalConcurrencyMutexName -ScriptBlock {
            Remove-StaleGlobalWorkerLeases -GatePath $globalConcurrencyGatePath -LeaseTimeoutMinutes $GlobalConcurrencyLeaseTimeoutMinutes
            $leases = @(Get-ChildItem -LiteralPath $globalConcurrencyGatePath -Filter '*.json' -File -ErrorAction SilentlyContinue)
            if ($leases.Count -lt $GlobalConcurrencyLimit) {
                $path = Join-Path $globalConcurrencyGatePath ("lease_{0}_{1}.json" -f $PID,([guid]::NewGuid().ToString("N")))
                $leaseData = [PSCustomObject]@{
                    LauncherInstanceId = $launcherInstanceId
                    ProcessId = $PID
                    LauncherProcessId = $PID
                    Computer = $Computer
                    Cycle = $CycleNumber
                    JobId = ""
                    JobName = ""
                    WorkerProcessId = 0
                    WorkerProcessName = ""
                    WorkerStartedUtc = ""
                    CreatedUtc = (Get-Date).ToUniversalTime().ToString("o")
                    LastUpdatedUtc = (Get-Date).ToUniversalTime().ToString("o")
                    Host = $env:COMPUTERNAME
                }
                Save-GlobalLeaseData -Path $path -Data $leaseData -Depth 3
                return $path
            }
            return ""
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
            $activeLeases = @(Get-ChildItem -LiteralPath $globalConcurrencyGatePath -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
            $waitMinutes = [math]::Round(($now - $waitStarted).TotalMinutes, 1)
            Write-Host ("Waiting for global worker lease: Computer={0}; Active={1}; Limit={2}; Wait={3} minute(s)." -f $Computer,$activeLeases,$GlobalConcurrencyLimit,$waitMinutes) -ForegroundColor DarkYellow
            $lastWaitLog = $now
            $waitWasLogged = $true
        }
        Start-Sleep -Seconds $JobPollSeconds
    }
}

function Release-GlobalWorkerLease {
    param([AllowNull()][string]$LeasePath)

    if ([string]::IsNullOrWhiteSpace($LeasePath)) { return }

    try {
        $removed = Invoke-WithGlobalGateMutex -MutexName $globalConcurrencyMutexName -ScriptBlock {
            Remove-GlobalWorkerLeaseFileUnlocked -Path $LeasePath
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

function Stop-LocalPsExecProcessTreeFromLease {
    param([AllowNull()][string]$LeasePath)

    if ([string]::IsNullOrWhiteSpace($LeasePath) -or -not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return "No active worker lease was available." }
    try {
        $leaseData = Invoke-WithGlobalGateMutex -MutexName $globalConcurrencyMutexName -ScriptBlock {
            if (-not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return $null }
            return (Read-GlobalLeaseData -Path $LeasePath)
        }
        $psExecPid = if ($leaseData -and $leaseData.PSObject.Properties["PsExecProcessId"]) { [int]$leaseData.PsExecProcessId } else { 0 }
        if ($psExecPid -le 0) { return "Worker lease contains no PsExec process id." }
        if (-not (Test-GlobalGateProcessAlive -ProcessId $psExecPid)) { return ("PsExec process {0} had already exited." -f $psExecPid) }
        $taskKillOutput = & taskkill.exe /PID $psExecPid /T /F 2>&1
        $taskKillExitCode = $LASTEXITCODE
        return ("Local PsExec process tree stop requested. PID={0}; ExitCode={1}; Output={2}" -f $psExecPid,$taskKillExitCode,(@($taskKillOutput) -join " "))
    }
    catch { return ("Local PsExec process tree stop failed: {0}" -f $_.Exception.Message) }
}

function Get-ComputerDnsCandidates {
    param(
        [Parameter(Mandatory=$true)][string]$ComputerName,
        [AllowNull()][string]$DomainSuffix
    )

    $candidates = @()
    $shortName = ([string]$ComputerName).Trim()
    if (-not [string]::IsNullOrWhiteSpace($shortName)) {
        $candidates += $shortName
    }

    $suffix = ([string]$DomainSuffix).Trim().Trim('.')
    if ($shortName -notmatch '\.' -and -not [string]::IsNullOrWhiteSpace($suffix)) {
        $candidates += ("{0}.{1}" -f $shortName,$suffix)
    }

    return @($candidates | Select-Object -Unique)
}

function Resolve-ComputerConnectionTarget {
    param(
        [Parameter(Mandatory=$true)][string]$ComputerName,
        [AllowNull()][string]$DomainSuffix
    )

    $lastError = ""
    $candidates = @(Get-ComputerDnsCandidates -ComputerName $ComputerName -DomainSuffix $DomainSuffix)
    foreach ($candidate in $candidates) {
        try {
            $dns = [System.Net.Dns]::GetHostEntry($candidate)
            return [PSCustomObject]@{
                Computer = $ComputerName
                ConnectionTarget = $candidate
                Resolved = $true
                AddressList = (($dns.AddressList | ForEach-Object { $_.IPAddressToString }) -join ";")
                Error = ""
            }
        }
        catch {
            $lastError = $_.Exception.Message
        }
    }

    return [PSCustomObject]@{
        Computer = $ComputerName
        ConnectionTarget = $(if ($candidates.Count -gt 0) { $candidates[-1] } else { $ComputerName })
        Resolved = $false
        AddressList = ""
        Error = $lastError
    }
}
function Test-SampleDnsResolution {
    param(
        [Parameter(Mandatory=$true)][AllowEmptyCollection()][object[]]$ComputerNames,
        [AllowNull()][string]$DomainSuffix,
        [int]$SampleSize = 5
    )

    $allNames = @($ComputerNames | ForEach-Object { ([string]$_).Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $sample = if ($allNames.Count -le $SampleSize) { @($allNames) } else { @($allNames | Get-Random -Count $SampleSize) }
    $resolved = 0
    $failed = 0
    $results = @()
    foreach ($name in $sample) {
        $dnsResult = Resolve-ComputerConnectionTarget -ComputerName $name -DomainSuffix $DomainSuffix
        if ($dnsResult.Resolved) {
            $resolved++
        }
        else {
            $failed++
        }
        $results += $dnsResult
    }

    return [PSCustomObject]@{
        Tested    = $sample.Count
        Resolved  = $resolved
        Failed    = $failed
        AllFailed = ($failed -eq $sample.Count -and $sample.Count -gt 0)
        Samples   = @($results)
    }
}

function Invoke-IntuneHybridJoinRepairCycle {
    param(
        [Parameter(Mandatory=$true)][int]$CycleNumber,
        [Parameter(Mandatory=$true)][string[]]$CycleScriptArgs
    )

    $computers = @(Get-ComputerList -Path $ComputerListPath)
    $computerListStats = Get-ComputerListStats -Path $ComputerListPath
    $script:CurrentCycleTotalComputers = $computers.Count
    $script:CurrentCycleComputerListStats = $computerListStats
    if ($computerListStats.DuplicateGroups -gt 0) {
        Write-Host ("Cycle {0}: Computers.txt contains {1} duplicate line(s) in {2} duplicate group(s). Duplicates ignored: {3}" -f $CycleNumber,$computerListStats.DuplicateLines,$computerListStats.DuplicateGroups,$computerListStats.DuplicateSamples) -ForegroundColor Yellow
    }
    if (-not $computers -or $computers.Count -eq 0) {
        Write-Host "No computers found in $ComputerListPath." -ForegroundColor Yellow
        return $null
    }

    if (-not $DryRun -and $IntuneInventorySet -and $IntuneInventorySet.Count -gt 0) {
        $alreadyEnrolledFromInventory = @(
            $computers |
                Where-Object {
                    $inventoryKey = Get-ComputerListKey -ComputerName $_
                    -not [string]::IsNullOrWhiteSpace($inventoryKey) -and $IntuneInventorySet.ContainsKey($inventoryKey)
                } |
                ForEach-Object {
                    [PSCustomObject]@{
                        Computer = $_
                        IntuneInventoryPresent = $true
                        Status = "SKIPPED_ALREADY_INTUNE_IN_INVENTORY"
                        NextAction = "NO_ACTION_ALREADY_INTUNE"
                    }
                }
        )

        if ($alreadyEnrolledFromInventory.Count -gt 0) {
            try {
                $preMoveResult = Move-AlreadyEnrolledComputersFromList -ComputerListPath $ComputerListPath -CycleSummary $alreadyEnrolledFromInventory
                if ($preMoveResult.Moved -gt 0) {
                    Write-Host ("Cycle {0}: pre-filtered {1} already-enrolled computer(s) from DevicesIntune.csv to {2}. PsExec will be skipped for them." -f $CycleNumber,$preMoveResult.Moved,$preMoveResult.AlreadyEnrolledPath) -ForegroundColor Green
                    $computers = @(Get-ComputerList -Path $ComputerListPath)
                    $computerListStats = Get-ComputerListStats -Path $ComputerListPath
                    $script:CurrentCycleTotalComputers = $computers.Count
                    $script:CurrentCycleComputerListStats = $computerListStats
                    if (-not $computers -or $computers.Count -eq 0) {
                        Write-Host ("Cycle {0}: all computers were already present in Intune inventory. Nothing left to run." -f $CycleNumber) -ForegroundColor Green
                        return $null
                    }
                }
            }
            catch {
                Write-Host ("Cycle {0}: failed to pre-filter already-enrolled computers from DevicesIntune.csv: {1}" -f $CycleNumber,$_.Exception.Message) -ForegroundColor Yellow
            }
        }
    }

    $summary = New-Object System.Collections.Generic.List[object]
    $reportColumns = @(Get-LauncherReportColumns)
    $liveSummaryPath = Join-Path $ReportRoot ("PsExec_IntuneHybridJoinRepair_Live_cycle{0}_{1}.csv" -f $CycleNumber,(Get-Date -Format "yyyyMMdd_HHmmss"))
    Initialize-LiveCycleReport -Path $liveSummaryPath -Columns $reportColumns
    $cycleStart = Get-Date
    $lastLiveHtmlWrite = [datetime]::MinValue
    try {
        $cycleProgress = New-HybridJoinCycleProgressRows -CycleNumber $CycleNumber -CycleStart $cycleStart -TotalComputers $computers.Count -QueuedComputers 0 -CompletedComputers 0 -RunningComputers 0 -ComputerListStats $computerListStats
        $mergedProgressRows = @($script:AllCycleProgressRows.ToArray()) + @($cycleProgress)
        New-CycleHtmlReport -Summary @($script:AllCycleResults.ToArray()) -Path $script:MergedHtmlReportPath -CycleNumber $CycleNumber -GeneratedAt (Get-Date) -IsLive -CycleProgress $mergedProgressRows -RunningJobRows @()
        $lastLiveHtmlWrite = Get-Date
    }
    catch {
        Write-Host ("Cycle {0}: failed to initialize merged HTML report: {1}" -f $CycleNumber,$_.Exception.Message) -ForegroundColor Yellow
    }

    $dnsCheck = Test-SampleDnsResolution -ComputerNames $computers -DomainSuffix $AdDomain
    if ($dnsCheck.AllFailed) {
        Write-Host ""
        Write-Host ("*** DNS WARNING: resolution failed on all {0} sampled computers. Check VPN connectivity and DNS configuration on this machine before continuing. ***" -f $dnsCheck.Tested) -ForegroundColor Red
        Write-Host ""

        if (-not $ContinueOnDnsPreflightFailure) {
            $status = "DNS_PREFLIGHT_ALL_SAMPLES_FAILED"
            $nextAction = Get-NextActionFromLauncherStatus -Status $status
            $sampleDetails = @($dnsCheck.Samples | ForEach-Object { "{0}->{1}" -f $_.Computer,$_.ConnectionTarget }) -join "; "
            $detail = ("DNS resolution failed for all {0} sampled computer(s). Cycle stopped before queuing PsExec jobs. DomainSuffix={1}; Samples={2}. Use -ContinueOnDnsPreflightFailure only after confirming the network path is intentional." -f $dnsCheck.Tested,$AdDomain,$sampleDetails)
            $dnsSamplesByComputer = @{}
            foreach ($sample in @($dnsCheck.Samples)) {
                $sampleKey = ([string]$sample.Computer).Trim().Split('.')[0].ToUpperInvariant()
                if (-not [string]::IsNullOrWhiteSpace($sampleKey) -and -not $dnsSamplesByComputer.ContainsKey($sampleKey)) {
                    $dnsSamplesByComputer[$sampleKey] = $sample
                }
            }

            foreach ($computer in $computers) {
                $row = [ordered]@{}
                foreach ($column in $reportColumns) { $row[$column] = "" }
                $row["LauncherVersion"] = $LauncherVersion
                $row["Cycle"] = $CycleNumber
                $row["Computer"] = $computer
                $row["Timestamp"] = Get-Date
                $row["DryRun"] = [bool]$DryRun
                $computerKey = ([string]$computer).Trim().Split('.')[0].ToUpperInvariant()
                if ($dnsSamplesByComputer.ContainsKey($computerKey)) {
                    $sample = $dnsSamplesByComputer[$computerKey]
                    $row["ConnectionTarget"] = $sample.ConnectionTarget
                    $row["DnsResolved"] = [bool]$sample.Resolved
                    $row["DnsAddressList"] = $sample.AddressList
                }
                else {
                    $candidates = @(Get-ComputerDnsCandidates -ComputerName $computer -DomainSuffix $AdDomain)
                    $row["ConnectionTarget"] = if ($candidates.Count -gt 0) { $candidates[-1] } else { $computer }
                }
                $row["RemoteDetail"] = "Cycle stopped by DNS preflight before ping, ADMIN$ or PsExec. Blank reachability fields mean not tested."
                $row["Status"] = $status
                $row["EffectiveStatus"] = $status
                $row["NextAction"] = $nextAction
                $row["EffectiveNextAction"] = $nextAction
                $row["ErrorMessage"] = $detail
                $item = [PSCustomObject]$row
                $summary.Add($item)
                Add-LiveCycleReportRow -Path $liveSummaryPath -Columns $reportColumns -Row $item
            }

            $summaryRows = @($summary | ForEach-Object { $_ })
            $summaryPath = Join-Path $ReportRoot ("PsExec_IntuneHybridJoinRepair_Summary_cycle{0}_{1}.csv" -f $CycleNumber,(Get-Date -Format "yyyyMMdd_HHmmss"))
            Get-PortableReportRows -Rows $summaryRows | Select-Object $reportColumns | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding UTF8
            foreach ($summaryRow in $summaryRows) { [void]$script:AllCycleResults.Add($summaryRow) }
            $finalProgress = New-HybridJoinCycleProgressRows -CycleNumber $CycleNumber -CycleStart $cycleStart -TotalComputers $computers.Count -QueuedComputers 0 -CompletedComputers $summaryRows.Count -RunningComputers 0 -ComputerListStats $computerListStats
            [void]$script:AllCycleProgressRows.Add($finalProgress)
            New-CycleHtmlReport -Summary @($script:AllCycleResults.ToArray()) -Path $script:MergedHtmlReportPath -CycleNumber $CycleNumber -GeneratedAt (Get-Date) -IsLive -CycleProgress @($script:AllCycleProgressRows.ToArray()) -RunningJobRows @()
            Write-Host ("Cycle {0} stopped by DNS preflight. Summary: {1}" -f $CycleNumber,$summaryPath) -ForegroundColor Red
            return $summaryPath
        }
    }

    Write-Host ("Cycle {0} started. Computers={1}; Throttle={2}; Args={3}" -f $CycleNumber,$computers.Count,$ThrottleLimit,($CycleScriptArgs -join ' ')) -ForegroundColor Cyan
    Write-Host ("Cycle {0} live report: {1}" -f $CycleNumber,$liveSummaryPath) -ForegroundColor Green
    Write-Host ("Merged HTML report  : {0}" -f $script:MergedHtmlReportPath) -ForegroundColor Green

    $worker = {
        param(
            [string]$Computer,
            [string]$ConnectionTarget,
            [int]$CycleNumber,
            [string]$LocalScriptPath,
            [string]$ScriptName,
            [string]$RemoteRelativeDir,
            [string]$RemoteScriptPath,
            [string]$RemoteDataRelativeDir,
            [string]$PsExecPath,
            [string]$LogRoot,
            [string]$LauncherVersion,
            [bool]$DryRun,
            [bool]$CollectRemoteLogs,
            [bool]$SkipVirtualMachines,
            [string]$CentralLogRoot,
            [bool]$KeepCentralLogHistory,
            [string]$CentralLogCollectionMode,
            [hashtable]$IntuneInventorySet,
            [hashtable]$EntraInventoryMap,
            [hashtable]$AdInventoryMap,
            [string]$CycleScriptArgsJson,
            [int]$PsExecTimeoutMinutes,
            [int]$CommunicationLostEvidenceWaitMinutes,
            [int]$CommunicationLostEvidencePollMinutes,
            [string]$GlobalWorkerLeasePath,
            [string]$GlobalWorkerLeaseMutexName
        )

        $ErrorActionPreference = "Stop"
        $CycleScriptArgs = @()
        if (-not [string]::IsNullOrWhiteSpace($CycleScriptArgsJson)) {
            try {
                $decodedCycleScriptArgs = $CycleScriptArgsJson | ConvertFrom-Json -ErrorAction Stop
                if ($decodedCycleScriptArgs -and $decodedCycleScriptArgs.PSObject.Properties["Args"]) {
                    $CycleScriptArgs = @($decodedCycleScriptArgs.Args | ForEach-Object { [string]$_ })
                }
            }
            catch {
                throw ("Invalid cycle script argument payload: {0}" -f $_.Exception.Message)
            }
        }

        function Read-WorkerLeaseData {
            param([Parameter(Mandatory=$true)][string]$Path)
            $lastError = $null
            for ($attempt = 1; $attempt -le 5; $attempt++) {
                try { return (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop) }
                catch { $lastError = $_; Start-Sleep -Milliseconds ([math]::Min(1000,100 * $attempt)) }
            }
            if ($lastError) { throw $lastError }
        }

        function Save-WorkerLeaseData {
            param([Parameter(Mandatory=$true)][string]$Path,[Parameter(Mandatory=$true)]$Data)
            $parent = Split-Path -Parent $Path
            $json = $Data | ConvertTo-Json -Depth 4
            $lastError = $null
            for ($attempt = 1; $attempt -le 5; $attempt++) {
                $tempPath = Join-Path $parent (".{0}.{1}.tmp" -f (Split-Path -Leaf $Path),[guid]::NewGuid().ToString("N"))
                try {
                    Set-Content -LiteralPath $tempPath -Value $json -Encoding UTF8 -Force -ErrorAction Stop
                    Move-Item -LiteralPath $tempPath -Destination $Path -Force -ErrorAction Stop
                    return
                }
                catch { $lastError = $_; Start-Sleep -Milliseconds ([math]::Min(1000,100 * $attempt)) }
                finally {
                    if (Test-Path -LiteralPath $tempPath -PathType Leaf) {
                        try { Remove-Item -LiteralPath $tempPath -Force -ErrorAction Stop } catch { }
                    }
                }
            }
            if ($lastError) { throw $lastError }
        }
        function Invoke-WithWorkerLeaseMutex {
            param(
                [AllowNull()][string]$MutexName,
                [Parameter(Mandatory=$true)][scriptblock]$ScriptBlock
            )

            if ([string]::IsNullOrWhiteSpace($MutexName)) {
                & $ScriptBlock
                return
            }

            $mutex = New-Object System.Threading.Mutex($false, $MutexName)
            $acquired = $false
            try {
                try { $acquired = $mutex.WaitOne(30000) }
                catch [System.Threading.AbandonedMutexException] { $acquired = $true }
                if ($acquired) { & $ScriptBlock }
            }
            finally {
                if ($acquired) { try { $mutex.ReleaseMutex() } catch { } }
                $mutex.Dispose()
            }
        }

        function Update-WorkerLease {
            param(
                [AllowNull()][string]$LeasePath,
                [AllowNull()][string]$MutexName
            )

            if ([string]::IsNullOrWhiteSpace($LeasePath) -or -not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return }

            try {
                Invoke-WithWorkerLeaseMutex -MutexName $MutexName -ScriptBlock {
                    if (-not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return }
                    $data = Read-WorkerLeaseData -Path $LeasePath
                    $processName = ""
                    try { $processName = (Get-Process -Id $PID -ErrorAction Stop).ProcessName } catch { }
                    $data | Add-Member -NotePropertyName WorkerProcessId -NotePropertyValue $PID -Force
                    $data | Add-Member -NotePropertyName WorkerProcessName -NotePropertyValue $processName -Force
                    $data | Add-Member -NotePropertyName WorkerStartedUtc -NotePropertyValue ((Get-Date).ToUniversalTime().ToString("o")) -Force
                    $data | Add-Member -NotePropertyName LastUpdatedUtc -NotePropertyValue ((Get-Date).ToUniversalTime().ToString("o")) -Force
                    Save-WorkerLeaseData -Path $LeasePath -Data $data
                }
            }
            catch { }
        }

        Update-WorkerLease -LeasePath $GlobalWorkerLeasePath -MutexName $GlobalWorkerLeaseMutexName

        function Update-TimestampedLogFile {
            param([Parameter(Mandatory=$true)][string]$Path)

            if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
                return
            }

            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $updatedLines = [System.IO.File]::ReadAllLines($Path) | ForEach-Object {
                if ($_ -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\b') {
                    $_
                }
                elseif ($_ -match '^\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]\s*(.*)$') {
                    "{0} {1}" -f $Matches[1], $Matches[2]
                }
                else {
                    "{0} {1}" -f $timestamp, $_
                }
            }

            $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
            [System.IO.File]::WriteAllLines($Path, [string[]]$updatedLines, $utf8NoBom)
        }

        function Complete-WorkerResult {
            param(
                [Parameter(Mandatory=$true)]$Result,
                [Parameter(Mandatory=$true)][string]$Path
            )

            try { Update-TimestampedLogFile -Path $Path } catch { }
            return [PSCustomObject]$Result
        }

        function Get-ScriptVersionFromFile {
            param([Parameter(Mandatory=$true)][string]$Path)

            try {
                if (-not (Test-Path -LiteralPath $Path)) { return "" }
                $match = Select-String -LiteralPath $Path -Pattern '^\s*\$ScriptVersion\s*=\s*"([^"]+)"' -ErrorAction Stop | Select-Object -First 1
                if ($match -and $match.Matches.Count -gt 0) {
                    return $match.Matches[0].Groups[1].Value
                }
            }
            catch { }

            return ""
        }

        function Get-FileSha256 {
            param([Parameter(Mandatory=$true)][string]$Path)

            try {
                if (-not (Test-Path -LiteralPath $Path)) { return "" }
                return (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash
            }
            catch {
                return ""
            }
        }

        function Get-NextActionFromLauncherStatus {
            param([Parameter(Mandatory=$true)][string]$Status)

            switch ($Status) {
                "SUCCESS" { return "NO_ACTION_ALREADY_INTUNE_OR_HEALTHY" }
                "AUDIT_SUCCESS_ALREADY_INTUNE" { return "NO_ACTION_ALREADY_INTUNE" }
                "AUDIT_INTUNE_MISSING" { return "RUN_REPAIR" }
                "AUDIT_STALE_INTUNE_ENROLLMENT_LOCAL" { return "CLEAN_STALE_INTUNE_OPTIN" }
                "INTUNE_ENROLLMENT_PENDING_CONFIRMATION" { return "RECHECK_LATER_INTUNE_ENROLLMENT" }
                "ENDPOINT_RUN_ACTIVE" { return "WAIT_ACTIVE_ENDPOINT_RUN" }
                "SKIPPED_BY_STATUS_BACKOFF" { return "WAIT_STATUS_BACKOFF" }
                "REBOOT_SAFETY_LIMIT_REACHED_POST_DSREG_LEAVE" { return "REVIEW_REBOOT_HISTORY_AND_HYBRID_JOIN" }
                "ADMIN_SHARE_UNREACHABLE" { return "FIX_ADMIN_SHARE_OR_NETWORK" }
                "DNS_PREFLIGHT_ALL_SAMPLES_FAILED" { return "CHECK_VPN_DNS_BEFORE_LOT" }
                "RUN_GUARD_ACTIVE" { return "WAIT_RUN_GUARD" }
                "SKIPPED_VIRTUAL_MACHINE" { return "NO_ACTION_VIRTUAL_MACHINE" }
                "REBOOT_TRIGGERED_WAITING_FOR_USER_LOGON" { return "WAIT_USER_LOGON" }
                "WAITING_FOR_INTERACTIVE_USER_LOGON" { return "WAIT_USER_LOGON" }
                "INTUNE_USER_AUTOENROLL_LOCAL_INTERACTIVE_USER" { return "LOGON_WITH_DOMAIN_OR_AAD_USER" }
                "INTUNE_USER_AUTOENROLL_TASK_NOT_FOUND" { return "FIX_GPO_USER_AUTOENROLL_TASK" }
                "STALE_INTUNE_ENROLLMENT_LOCAL" { return "CLEAN_STALE_INTUNE_OPTIN" }
                "NON_INTUNE_MDM_ENROLLED" { return "CLEAN_NON_INTUNE_MDM_OPTIN" }
                "ENTRA_HYBRID_PENDING_ADJ_TRIGGERED" { return "RECHECK_ENTRA_PENDING_AFTER_ADJ" }
                "ENTRA_HYBRID_PENDING_RETRY_EXHAUSTED" { return "CHECK_AD_CONNECT_OR_DUPLICATE_ENTRA_DEVICE" }
                "ENTRA_PENDING_RESOLVED_POST_CYCLE" { return "RECHECK_INTUNE_ENROLLMENT" }
                "USER_NOT_AZUREAD" { return "CHECK_USER_AAD_OR_LOGON_CONTEXT" }
                "USER_PRT_NOT_AVAILABLE" { return "CHECK_USER_PRT" }
                "USER_PRT_REFRESH_FAILED" { return "FIX_USER_PRT_OR_RELOGIN" }
                "USER_SESSION_REMOTE" { return "LOGON_ON_CONSOLE" }
                "INTUNE_AUTOENROLL_POLICY_NOT_CONFIGURED" { return "CHECK_GPO_AUTOENROLL" }
                    "KEY_SIGN_TEST_FAILED" { return "REPAIR_HYBRID_JOIN_KEY_OR_ALLOW_LEAVE" }
                    "INTUNE_ENROLLMENT_CONNECTIVITY_FAILED" { return "CHECK_CONNECTIVITY" }
                    "DOMAIN_CONTROLLER_UNREACHABLE" { return "FIX_DOMAIN_CONNECTIVITY_OR_VPN" }
                    "PSEXEC_TIMEOUT" { return "CHECK_REMOTE_LOG_OR_RETRY" }
                    "PSEXEC_COMMUNICATION_LOST" { return "RETRY_PSEXEC_OR_CHECK_REMOTE_SERVICE" }
                    "PSEXEC_EXIT_UNKNOWN" { return "CHECK_CURRENT_RUN_REMOTE_LOG" }
        "PSEXEC_EXIT_-1073741790" { return "CHECK_EDR_OR_EXECUTION_POLICY_BLOCK" }
                    "REMOTE_DIRECTORY_CREATE_FAILED" { return "FIX_SCRIPT_COPY_OR_ADMIN_SHARE" }
                    "REMOTE_SCRIPT_COPY_FAILED" { return "FIX_SCRIPT_COPY_OR_SECURITY" }
                    "REMOTE_SCRIPT_MISSING" { return "FIX_SCRIPT_COPY_OR_SECURITY" }
                    default {
                    if ($Status -like "ERROR*") { return "CHECK_CONNECTIVITY_OR_ADMIN_ACCESS" }
            if ($Status -like "PSEXEC_EXIT_-1073741790*") { return "CHECK_EDR_OR_EXECUTION_POLICY_BLOCK" }
                    if ($Status -like "PSEXEC_EXIT*") { return "CHECK_CURRENT_RUN_REMOTE_LOG" }
                    return "REVIEW_LOGS"
                }
            }
        }

        function Copy-RemoteEvidenceFolder {
            param(
                [Parameter(Mandatory=$true)][string]$RemoteDataPath,
                [Parameter(Mandatory=$true)][string]$DestinationPath,
                [Parameter(Mandatory=$true)][string]$ScriptName,
                [ValidateSet("Standard","Full")][string]$CentralLogCollectionMode = "Standard",
                [datetime]$Since = [datetime]::MinValue
            )

            $copyCount = 0
            $maxFileBytes = 5MB
            $currentThreshold = $Since.AddSeconds(-5)

            function Copy-EvidenceFile {
                param([Parameter(Mandatory=$true)][string]$SourceFile,[Parameter(Mandatory=$true)][string]$TargetFolder)
                try {
                    if (-not (Test-Path -LiteralPath $TargetFolder)) { [System.IO.Directory]::CreateDirectory($TargetFolder) | Out-Null }
                    Copy-Item -LiteralPath $SourceFile -Destination $TargetFolder -Force -ErrorAction Stop
                    return $true
                }
                catch { return $false }
            }

            foreach ($folderName in @("Logs","Output","Transcripts","State")) {
                $sourceFolder = Join-Path $RemoteDataPath $folderName
                if (-not (Test-Path -LiteralPath $sourceFolder)) { continue }
                foreach ($file in @(Get-ChildItem -LiteralPath $sourceFolder -Recurse -File -Force -ErrorAction SilentlyContinue)) {
                    if ($CentralLogCollectionMode -eq "Standard" -and ($file.Length -gt $maxFileBytes -or ($file.LastWriteTime -lt $currentThreshold -and -not ($folderName -eq "State" -and $file.Name -in @("EndpointInstance.json","RebootSafety.json","RetryAfterReboot.json"))))) { continue }
                    $relativePath = $file.FullName.Substring($sourceFolder.Length).TrimStart("\")
                    $relativeDir = Split-Path -Parent $relativePath
                    $fileTargetFolder = if ([string]::IsNullOrWhiteSpace($relativeDir)) { Join-Path $DestinationPath $folderName } else { Join-Path (Join-Path $DestinationPath $folderName) $relativeDir }
                    if (Copy-EvidenceFile -SourceFile $file.FullName -TargetFolder $fileTargetFolder) { $copyCount++ }
                }
            }

            foreach ($file in @(Get-ChildItem -LiteralPath $RemoteDataPath -File -Force -ErrorAction SilentlyContinue)) {
                if ($file.Name -eq $ScriptName -or $file.Extension -notin @(".csv",".log",".txt",".html",".json",".xml",".evtx")) { continue }
                $isCurrentOrState = ($file.Name -eq "LastRun.json" -or $file.Name -eq "EndpointInstance.json" -or $file.Name -eq "RebootSafety.json" -or $file.LastWriteTime -ge $currentThreshold)
                if ($CentralLogCollectionMode -eq "Standard" -and ($file.Length -gt $maxFileBytes -or -not $isCurrentOrState)) { continue }
                if (Copy-EvidenceFile -SourceFile $file.FullName -TargetFolder $DestinationPath) { $copyCount++ }
            }

            if ($copyCount -eq 0) {
                Remove-Item -LiteralPath $DestinationPath -Recurse -Force -ErrorAction SilentlyContinue
                throw "No current remote evidence files found to collect."
            }
            return $copyCount
        }

        function Get-RemoteEvidenceFinalStatus {
            param(
                [Parameter(Mandatory=$true)][string]$EvidencePath,
                [Parameter(Mandatory=$false)][datetime]$Since = [datetime]::MinValue,
                [switch]$RequireCompletedRun
            )

            $completedRunId = ""
            $completedRunStatus = ""
            $completedRunExitCode = ""
            $completedRunNextAction = ""
            $completedRunDetail = ""
            $lastRunPath = Join-Path $EvidencePath "LastRun.json"
            if ($RequireCompletedRun) {
                if (-not (Test-Path -LiteralPath $lastRunPath -ErrorAction SilentlyContinue)) { return $null }
                try {
                    $lastRun = Get-Content -LiteralPath $lastRunPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                    if ($null -eq $lastRun -or [string]::IsNullOrWhiteSpace([string]$lastRun.RunId) -or [string]::IsNullOrWhiteSpace([string]$lastRun.EndTime)) {
                        return $null
                    }
                    $lastRunStart = [datetime]$lastRun.StartTime
                    if ($lastRunStart -lt $Since.AddSeconds(-5)) { return $null }
                    $completedRunId = ([string]$lastRun.RunId).Trim()
                    $completedRunStatus = ([string]$lastRun.Status).Trim()
                    $completedRunExitCode = ([string]$lastRun.ExitCode).Trim()
                    $completedRunNextAction = ([string]$lastRun.NextAction).Trim()
                    $completedRunDetail = ([string]$lastRun.Detail).Trim()
                    if ([string]::IsNullOrWhiteSpace($completedRunStatus) -or [string]::IsNullOrWhiteSpace($completedRunExitCode)) {
                        return $null
                    }
                }
                catch {
                    return $null
                }
            }

            $csv = Get-ChildItem -LiteralPath $EvidencePath -Recurse -File -Filter "IntuneHybridJoinToolkit_*.csv" -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -ge $Since.AddSeconds(-5) } |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

            if (-not $csv) {
                if ($RequireCompletedRun -and -not [string]::IsNullOrWhiteSpace($completedRunStatus)) {
                    return [PSCustomObject]@{
                        RunId = $completedRunId
                        Status = $completedRunStatus
                        ExitCode = $completedRunExitCode
                        NextAction = $completedRunNextAction
                        Detail = $completedRunDetail
                        CsvPath = ""
                        LastRunPath = $lastRunPath
                        LastRunStatus = $completedRunStatus
                        LastRunExitCode = $completedRunExitCode
                        InteractiveUserName = ""
                        InteractiveUserDomain = ""
                        InteractiveUserAccountName = ""
                        InteractiveUserAccountType = ""
                        InteractiveSessionName = ""
                        InteractiveSessionState = ""
                        UserIsUserAzureAD = ""
                        UserAzureAdPrt = ""
                        UserSessionIsNotRemote = ""
                        RetryAfterRebootAction = ""
                        RetryAfterRebootDetail = ""
                        RetryAfterRebootAttempt = ""
                        RetryAfterRebootMaxAttempts = ""
                        RetryAfterRebootTaskName = ""
                    }
                }
                return $null
            }

            try {
                $row = Import-Csv -LiteralPath $csv.FullName -ErrorAction Stop | Select-Object -Last 1
                if ($null -eq $row -or [string]::IsNullOrWhiteSpace([string]$row.Status)) { return $null }

                $rowRunId = ""
                if ($row.PSObject.Properties["RunId"]) { $rowRunId = ([string]$row.RunId).Trim() }
                if ($RequireCompletedRun) {
                    if ([string]::IsNullOrWhiteSpace($rowRunId) -or $rowRunId -ne $completedRunId) {
                        return [PSCustomObject]@{
                            RunId = $completedRunId
                            Status = $completedRunStatus
                            ExitCode = $completedRunExitCode
                            NextAction = $completedRunNextAction
                            Detail = $completedRunDetail
                            CsvPath = ""
                            LastRunPath = $lastRunPath
                            LastRunStatus = $completedRunStatus
                            LastRunExitCode = $completedRunExitCode
                            InteractiveUserName = ""
                            InteractiveUserDomain = ""
                            InteractiveUserAccountName = ""
                            InteractiveUserAccountType = ""
                            InteractiveSessionName = ""
                            InteractiveSessionState = ""
                            UserIsUserAzureAD = ""
                            UserAzureAdPrt = ""
                            UserSessionIsNotRemote = ""
                            RetryAfterRebootAction = ""
                            RetryAfterRebootDetail = ""
                            RetryAfterRebootAttempt = ""
                            RetryAfterRebootMaxAttempts = ""
                            RetryAfterRebootTaskName = ""
                        }
                    }
                }

                $exitCode = ""
                if ($row.PSObject.Properties["ExitCode"]) { $exitCode = [string]$row.ExitCode }

                $nextAction = ""
                if ($row.PSObject.Properties["NextAction"]) { $nextAction = [string]$row.NextAction }
                if ([string]::IsNullOrWhiteSpace($nextAction) -and $RequireCompletedRun) { $nextAction = $completedRunNextAction }

                $detail = ""
                if ($row.PSObject.Properties["DsregStatusErrorMessage"]) { $detail = [string]$row.DsregStatusErrorMessage }
                if ([string]::IsNullOrWhiteSpace($detail) -and $row.PSObject.Properties["ErrorMessage"]) { $detail = [string]$row.ErrorMessage }
                if ([string]::IsNullOrWhiteSpace($detail) -and $RequireCompletedRun) { $detail = $completedRunDetail }

                $interactiveUserName = ""; if ($row.PSObject.Properties["InteractiveUserName"]) { $interactiveUserName = [string]$row.InteractiveUserName }
                $interactiveUserDomain = ""; if ($row.PSObject.Properties["InteractiveUserDomain"]) { $interactiveUserDomain = [string]$row.InteractiveUserDomain }
                $interactiveUserAccountName = ""; if ($row.PSObject.Properties["InteractiveUserAccountName"]) { $interactiveUserAccountName = [string]$row.InteractiveUserAccountName }
                $interactiveUserAccountType = ""; if ($row.PSObject.Properties["InteractiveUserAccountType"]) { $interactiveUserAccountType = [string]$row.InteractiveUserAccountType }
                $interactiveSessionName = ""; if ($row.PSObject.Properties["InteractiveSessionName"]) { $interactiveSessionName = [string]$row.InteractiveSessionName }
                $interactiveSessionState = ""; if ($row.PSObject.Properties["InteractiveSessionState"]) { $interactiveSessionState = [string]$row.InteractiveSessionState }
                $userIsUserAzureAD = ""; if ($row.PSObject.Properties["User_IsUserAzureAD"]) { $userIsUserAzureAD = [string]$row.User_IsUserAzureAD }
                $userAzureAdPrt = ""; if ($row.PSObject.Properties["User_AzureAdPrt"]) { $userAzureAdPrt = [string]$row.User_AzureAdPrt }
                $userSessionIsNotRemote = ""; if ($row.PSObject.Properties["User_SessionIsNotRemote"]) { $userSessionIsNotRemote = [string]$row.User_SessionIsNotRemote }
                $retryAfterRebootAction = ""; if ($row.PSObject.Properties["RetryAfterRebootAction"]) { $retryAfterRebootAction = [string]$row.RetryAfterRebootAction }
                $retryAfterRebootDetail = ""; if ($row.PSObject.Properties["RetryAfterRebootDetail"]) { $retryAfterRebootDetail = [string]$row.RetryAfterRebootDetail }
                $retryAfterRebootAttempt = ""; if ($row.PSObject.Properties["RetryAfterRebootAttempt"]) { $retryAfterRebootAttempt = [string]$row.RetryAfterRebootAttempt }
                $retryAfterRebootMaxAttempts = ""; if ($row.PSObject.Properties["RetryAfterRebootMaxAttempts"]) { $retryAfterRebootMaxAttempts = [string]$row.RetryAfterRebootMaxAttempts }
                $retryAfterRebootTaskName = ""; if ($row.PSObject.Properties["RetryAfterRebootTaskName"]) { $retryAfterRebootTaskName = [string]$row.RetryAfterRebootTaskName }

                return [PSCustomObject]@{
                    RunId = $rowRunId
                    Status = $(if ($RequireCompletedRun) { $completedRunStatus } else { ([string]$row.Status).Trim() })
                    ExitCode = $(if ($RequireCompletedRun) { $completedRunExitCode } else { $exitCode.Trim() })
                    NextAction = $nextAction.Trim()
                    Detail = $detail.Trim()
                    CsvPath = $csv.FullName
                    LastRunPath = $(if ($RequireCompletedRun) { $lastRunPath } else { "" })
                    LastRunStatus = $completedRunStatus
                    LastRunExitCode = $completedRunExitCode
                    InteractiveUserName = $interactiveUserName.Trim()
                    InteractiveUserDomain = $interactiveUserDomain.Trim()
                    InteractiveUserAccountName = $interactiveUserAccountName.Trim()
                    InteractiveUserAccountType = $interactiveUserAccountType.Trim()
                    InteractiveSessionName = $interactiveSessionName.Trim()
                    InteractiveSessionState = $interactiveSessionState.Trim()
                    UserIsUserAzureAD = $userIsUserAzureAD.Trim()
                    UserAzureAdPrt = $userAzureAdPrt.Trim()
                    UserSessionIsNotRemote = $userSessionIsNotRemote.Trim()
                    RetryAfterRebootAction = $retryAfterRebootAction.Trim()
                    RetryAfterRebootDetail = $retryAfterRebootDetail.Trim()
                    RetryAfterRebootAttempt = $retryAfterRebootAttempt.Trim()
                    RetryAfterRebootMaxAttempts = $retryAfterRebootMaxAttempts.Trim()
                    RetryAfterRebootTaskName = $retryAfterRebootTaskName.Trim()
                }
            }
            catch {
                return [PSCustomObject]@{
                    Status = ""
                    ExitCode = ""
                    NextAction = ""
                    Detail = ("Could not parse remote evidence CSV: {0}" -f $_.Exception.Message)
                    CsvPath = $csv.FullName
                }
            }
        }

        if ([string]::IsNullOrWhiteSpace($ConnectionTarget)) { $ConnectionTarget = $Computer }

        $runId = Get-Date -Format "yyyyMMdd_HHmmss"
        $remoteRootShare = "\\$ConnectionTarget\C$"
        $remoteAdminDir = Join-Path $remoteRootShare $RemoteRelativeDir
        $remoteAdminScript = Join-Path $remoteAdminDir $ScriptName
        $remoteDataAdminDir = Join-Path $remoteRootShare $RemoteDataRelativeDir
        $centralIdentity = ([string]$ConnectionTarget).Trim().TrimEnd(".").ToLowerInvariant()
        $centralShortName = (($Computer -split "\.")[0] -replace '[^A-Za-z0-9_-]','_')
        $centralHashProvider = [System.Security.Cryptography.SHA256]::Create()
        try {
            $centralHash = ([BitConverter]::ToString($centralHashProvider.ComputeHash([Text.Encoding]::UTF8.GetBytes($centralIdentity))) -replace '-','').Substring(0,8)
        }
        finally { $centralHashProvider.Dispose() }
        $safeComputerName = "{0}-{1}" -f $centralShortName,$centralHash
        $logPath = Join-Path $LogRoot ("{0}_cycle{1}_{2}.log" -f $safeComputerName,$CycleNumber,$runId)

        $result = [ordered]@{
            LauncherVersion = $LauncherVersion
            Cycle = $CycleNumber
            Computer = $Computer
            ConnectionTarget = $ConnectionTarget
            Timestamp = Get-Date
            DryRun = $DryRun
            DnsResolved = $false
            DnsAddressList = ""
            AdminShareReachable = $false
            RemotePayloadCopyAttempts = ""
            PingReachable = $false
            IsVirtualMachine = ""
            VirtualMachineEvidence = ""
            RemoteDirectoryCreated = $false
            ScriptCopied = $false
            LocalScriptVersion = ""
            RemoteScriptVersion = ""
            LocalScriptHash = ""
            RemoteScriptHash = ""
            PsExecExitCode = ""
            RemoteStatus = ""
            RemoteExitCode = ""
            RemoteNextAction = ""
            RemoteDetail = ""
            NextAction = ""
            EffectiveStatus = ""
            EffectiveNextAction = ""
            InteractiveUserName = ""
            InteractiveUserDomain = ""
            InteractiveUserAccountName = ""
            InteractiveUserAccountType = ""
            InteractiveSessionName = ""
            InteractiveSessionState = ""
            UserIsUserAzureAD = ""
            UserAzureAdPrt = ""
            UserSessionIsNotRemote = ""
            RetryAfterRebootAction = ""
            RetryAfterRebootDetail = ""
            RetryAfterRebootAttempt = ""
            RetryAfterRebootMaxAttempts = ""
            RetryAfterRebootTaskName = ""
            IntuneInventoryPresent = ""
            EntraInventoryPresent = ""
            EntraRegisteredState = ""
            EntraAlternativeSecurityIdCount = ""
            EntraPendingReason = ""
            EntraRegistrationDateTime = ""
            EntraTrustType = ""
            EntraDeviceId = ""
            EntraObjectId = ""
            ADInventoryPresent = ""
            ADDomain = ""
            ADEnabled = ""
            ADDNSHostName = ""
            ADDistinguishedName = ""
            ADOperatingSystem = ""
            ADLastLogonTimestampUtc = ""
            AdminShareFailureType = ""
            PostCycleIntuneInventoryChecked = ""
            PostCycleIntuneInventoryPresent = ""
            PostCycleIntuneEnrollmentDetected = ""
            PostCycleIntuneInventoryCsv = ""
            PostCycleIntuneInventoryError = ""
            PostCycleEntraInventoryChecked = ""
            PostCycleEntraInventoryPresent = ""
            PostCycleEntraRegisteredState = ""
            PostCycleEntraAlternativeSecurityIdCount = ""
            PostCycleEntraPendingResolved = ""
            PostCycleEntraInventoryCsv = ""
            PostCycleEntraInventoryError = ""
            PostCycleADInventoryChecked = ""
            PostCycleADInventoryPresent = ""
            PostCycleADInventoryCsv = ""
            PostCycleADInventoryError = ""
            RemoteLogsCollected = $false
            RemoteLogsPath = ""
            RemoteCurrentRunLogsPath = ""
            RemoteLogsError = ""
            Status = "STARTED"
            LogPath = $logPath
            ErrorMessage = ""
        }

        try {
            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Starting $Computer. ConnectionTarget=$ConnectionTarget. Cycle=$CycleNumber" | Set-Content -LiteralPath $logPath -Encoding UTF8
            $inventoryKey = ($Computer.Split(".")[0]).ToUpperInvariant()
            if ($IntuneInventorySet -and $IntuneInventorySet.Count -gt 0) {
                $result.IntuneInventoryPresent = [bool]$IntuneInventorySet.ContainsKey($inventoryKey)
            }
            if ($EntraInventoryMap -and $EntraInventoryMap.Count -gt 0 -and $EntraInventoryMap.ContainsKey($inventoryKey)) {
                $entraRow = $EntraInventoryMap[$inventoryKey]
                $result.EntraInventoryPresent = $true
                if ($entraRow.PSObject.Properties["EntraRegisteredState"]) { $result.EntraRegisteredState = [string]$entraRow.EntraRegisteredState }
                if ($entraRow.PSObject.Properties["AlternativeSecurityIdCount"]) { $result.EntraAlternativeSecurityIdCount = [string]$entraRow.AlternativeSecurityIdCount }
                if ($entraRow.PSObject.Properties["EntraPendingReason"]) { $result.EntraPendingReason = [string]$entraRow.EntraPendingReason }
                if ($entraRow.PSObject.Properties["RegistrationDateTime"]) { $result.EntraRegistrationDateTime = [string]$entraRow.RegistrationDateTime }
                if ($entraRow.PSObject.Properties["TrustType"]) { $result.EntraTrustType = [string]$entraRow.TrustType }
                if ($entraRow.PSObject.Properties["DeviceId"]) { $result.EntraDeviceId = [string]$entraRow.DeviceId }
                if ($entraRow.PSObject.Properties["EntraObjectId"]) { $result.EntraObjectId = [string]$entraRow.EntraObjectId }
            }
            elseif ($EntraInventoryMap -and $EntraInventoryMap.Count -gt 0) {
                $result.EntraInventoryPresent = $false
            }
            if ($AdInventoryMap -and $AdInventoryMap.Count -gt 0 -and $AdInventoryMap.ContainsKey($inventoryKey)) {
                $adRow = $AdInventoryMap[$inventoryKey]
                $result.ADInventoryPresent = $true
                if ($adRow.PSObject.Properties["ADDomain"]) { $result.ADDomain = [string]$adRow.ADDomain }
                if ($adRow.PSObject.Properties["Enabled"]) { $result.ADEnabled = [string]$adRow.Enabled }
                if ($adRow.PSObject.Properties["DNSHostName"]) { $result.ADDNSHostName = [string]$adRow.DNSHostName }
                if ($adRow.PSObject.Properties["DistinguishedName"]) { $result.ADDistinguishedName = [string]$adRow.DistinguishedName }
                if ($adRow.PSObject.Properties["OperatingSystem"]) { $result.ADOperatingSystem = [string]$adRow.OperatingSystem }
                if ($adRow.PSObject.Properties["LastLogonTimestampUtc"]) { $result.ADLastLogonTimestampUtc = [string]$adRow.LastLogonTimestampUtc }
            }
            elseif ($AdInventoryMap -and $AdInventoryMap.Count -gt 0) {
                $result.ADInventoryPresent = $false
            }

            try {
                $dns = [System.Net.Dns]::GetHostEntry($ConnectionTarget)
                $result.DnsResolved = $true
                $result.DnsAddressList = (($dns.AddressList | ForEach-Object { $_.IPAddressToString }) -join ";")
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] DNS resolved: $($result.DnsAddressList)" | Add-Content -LiteralPath $logPath -Encoding UTF8
            }
            catch {
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] WARN: DNS resolution failed: $($_.Exception.Message)" | Add-Content -LiteralPath $logPath -Encoding UTF8
            }

            $result.PingReachable = Test-Connection -ComputerName $ConnectionTarget -Count 1 -Quiet -ErrorAction SilentlyContinue
            if (-not $result.PingReachable) {
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] WARN: Ping failed. Trying administrative shares anyway." | Add-Content -LiteralPath $logPath -Encoding UTF8
            }

            if ($SkipVirtualMachines) {
                $result.VirtualMachineEvidence = "Launcher-side VM pre-check skipped to avoid WinRM. Endpoint guard will check locally under PsExec/SYSTEM."
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $($result.VirtualMachineEvidence)" | Add-Content -LiteralPath $logPath -Encoding UTF8
            }

            $maxPayloadCopyAttempts = 3
            $payloadReady = $false
            $payloadFailureStatus = ""
            $payloadFailureDetail = ""
            for ($payloadCopyAttempt = 1; $payloadCopyAttempt -le $maxPayloadCopyAttempts; $payloadCopyAttempt++) {
                $result.RemotePayloadCopyAttempts = [string]$payloadCopyAttempt
                if ($payloadCopyAttempt -gt 1) {
                    $delaySeconds = [math]::Min(30, 10 * ($payloadCopyAttempt - 1))
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Retrying administrative-share preflight and remote script copy. Attempt=$payloadCopyAttempt/$maxPayloadCopyAttempts; DelaySeconds=$delaySeconds" | Add-Content -LiteralPath $logPath -Encoding UTF8
                    Start-Sleep -Seconds $delaySeconds
                }

                $adminShare = "\\$ConnectionTarget\ADMIN$"
                $rootShare = "\\$ConnectionTarget\C$"
                $adminShareReachable = Test-Path -LiteralPath $adminShare -ErrorAction SilentlyContinue
                $rootShareReachable = Test-Path -LiteralPath $rootShare -ErrorAction SilentlyContinue
                $result.AdminShareReachable = ($adminShareReachable -and $rootShareReachable)
                if ($result.AdminShareReachable) {
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Administrative shares reachable. ADMIN$=$adminShareReachable; C$=$rootShareReachable" | Add-Content -LiteralPath $logPath -Encoding UTF8
                    $result.AdminShareFailureType = ""
                }
                else {
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] WARN: Required administrative share is not reachable. Attempt=$payloadCopyAttempt/$maxPayloadCopyAttempts; ADMIN$=$adminShareReachable; C$=$rootShareReachable; ADMINPath=$adminShare; RootPath=$rootShare" | Add-Content -LiteralPath $logPath -Encoding UTF8
                    if (-not $result.DnsResolved) {
                        $result.AdminShareFailureType = "DNS_FAILED"
                    }
                    elseif (-not $result.PingReachable) {
                        $result.AdminShareFailureType = "PING_FAILED_ADMIN_SHARE_FAILED"
                    }
                    else {
                        $result.AdminShareFailureType = "PING_OK_ADMIN_SHARE_FAILED"
                    }
                }

                if ($DryRun) {
                    if ($result.AdminShareReachable) {
                        $result.Status = "DRYRUN_READY"
                        $result.NextAction = "READY_FOR_REPAIR"
                        $result.RemoteDetail = "DNS/Ping/administrative-share pre-check completed. No script copied or executed."
                        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] DryRun status: $($result.Status); Detail=$($result.RemoteDetail)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                        return (Complete-WorkerResult -Result $result -Path $logPath)
                    }

                    $payloadFailureStatus = "DRYRUN_ADMIN_SHARE_UNREACHABLE"
                    $payloadFailureDetail = "Required administrative share is not reachable. PsExec/copy would probably fail."
                    if ($payloadCopyAttempt -lt $maxPayloadCopyAttempts) { continue }
                    $result.Status = $payloadFailureStatus
                    $result.NextAction = "FIX_ADMIN_SHARE_OR_NETWORK"
                    $result.RemoteDetail = $payloadFailureDetail
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] DryRun status: $($result.Status); Detail=$($result.RemoteDetail)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                    return (Complete-WorkerResult -Result $result -Path $logPath)
                }

                if (-not $result.AdminShareReachable) {
                    $payloadFailureStatus = "ADMIN_SHARE_UNREACHABLE"
                    $payloadFailureDetail = ("{0}: Required administrative share is not reachable. Script copy and PsExec execution were skipped." -f $result.AdminShareFailureType)
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] WARN: $payloadFailureDetail" | Add-Content -LiteralPath $logPath -Encoding UTF8
                    continue
                }

                try {
                    if (-not (Test-Path -LiteralPath $remoteAdminDir)) {
                        New-Item -ItemType Directory -Path $remoteAdminDir -Force -ErrorAction Stop | Out-Null
                    }
                    $result.RemoteDirectoryCreated = Test-Path -LiteralPath $remoteAdminDir
                }
                catch {
                    $result.RemoteDirectoryCreated = $false
                    $payloadFailureStatus = "REMOTE_DIRECTORY_CREATE_FAILED"
                    $payloadFailureDetail = "Remote repair folder could not be created or verified: $remoteAdminDir; Error=$($_.Exception.Message)"
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] WARN: $payloadFailureDetail" | Add-Content -LiteralPath $logPath -Encoding UTF8
                    continue
                }

                if (-not $result.RemoteDirectoryCreated) {
                    $payloadFailureStatus = "REMOTE_DIRECTORY_CREATE_FAILED"
                    $payloadFailureDetail = "Remote repair folder could not be created or verified: $remoteAdminDir"
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] WARN: $payloadFailureDetail" | Add-Content -LiteralPath $logPath -Encoding UTF8
                    continue
                }
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Remote repair folder ready: $remoteAdminDir" | Add-Content -LiteralPath $logPath -Encoding UTF8

                $result.LocalScriptVersion = Get-ScriptVersionFromFile -Path $LocalScriptPath
                $result.LocalScriptHash = Get-FileSha256 -Path $LocalScriptPath
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Local script before copy: Version=$($result.LocalScriptVersion); SHA256=$($result.LocalScriptHash); Path=$LocalScriptPath" | Add-Content -LiteralPath $logPath -Encoding UTF8

                $remoteStagingScript = Join-Path $remoteAdminDir (".{0}.{1}.{2}.tmp" -f $ScriptName,$PID,[guid]::NewGuid().ToString("N"))
                try {
                    $localScriptItem = Get-Item -LiteralPath $LocalScriptPath -ErrorAction Stop
                    Copy-Item -LiteralPath $LocalScriptPath -Destination $remoteStagingScript -Force -ErrorAction Stop
                    if (-not (Test-Path -LiteralPath $remoteStagingScript)) { throw "Staged remote script is missing: $remoteStagingScript" }

                    $stagingScriptItem = Get-Item -LiteralPath $remoteStagingScript -ErrorAction Stop
                    $stagingScriptVersion = Get-ScriptVersionFromFile -Path $remoteStagingScript
                    $stagingScriptHash = Get-FileSha256 -Path $remoteStagingScript
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Remote staging script: Version=$stagingScriptVersion; SHA256=$stagingScriptHash; Bytes=$($stagingScriptItem.Length); Path=$remoteStagingScript" | Add-Content -LiteralPath $logPath -Encoding UTF8
                    if ($stagingScriptItem.Length -ne $localScriptItem.Length) { throw "Staged remote script size mismatch. LocalBytes=$($localScriptItem.Length); StagedBytes=$($stagingScriptItem.Length)" }
                    if ([string]::IsNullOrWhiteSpace($result.LocalScriptHash) -or [string]::IsNullOrWhiteSpace($stagingScriptHash) -or $stagingScriptHash -ne $result.LocalScriptHash) { throw "Staged remote script hash mismatch. LocalSHA256=$($result.LocalScriptHash); StagedSHA256=$stagingScriptHash" }
                    if ((-not [string]::IsNullOrWhiteSpace($result.LocalScriptVersion)) -and $stagingScriptVersion -ne $result.LocalScriptVersion) { throw "Staged remote script version mismatch. LocalVersion=$($result.LocalScriptVersion); StagedVersion=$stagingScriptVersion" }

                    Move-Item -LiteralPath $remoteStagingScript -Destination $remoteAdminScript -Force -ErrorAction Stop
                    $result.ScriptCopied = Test-Path -LiteralPath $remoteAdminScript
                    if (-not $result.ScriptCopied) { throw "Final remote script is missing after atomic move: $remoteAdminScript" }

                    $remoteScriptItem = Get-Item -LiteralPath $remoteAdminScript -ErrorAction Stop
                    $result.RemoteScriptVersion = Get-ScriptVersionFromFile -Path $remoteAdminScript
                    $result.RemoteScriptHash = Get-FileSha256 -Path $remoteAdminScript
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Remote script after atomic copy: Version=$($result.RemoteScriptVersion); SHA256=$($result.RemoteScriptHash); Bytes=$($remoteScriptItem.Length); Path=$remoteAdminScript" | Add-Content -LiteralPath $logPath -Encoding UTF8
                    if ($remoteScriptItem.Length -ne $localScriptItem.Length) { throw "Final remote script size mismatch. LocalBytes=$($localScriptItem.Length); RemoteBytes=$($remoteScriptItem.Length)" }
                    if ([string]::IsNullOrWhiteSpace($result.RemoteScriptHash) -or $result.RemoteScriptHash -ne $result.LocalScriptHash) { throw "Final remote script hash mismatch. LocalSHA256=$($result.LocalScriptHash); RemoteSHA256=$($result.RemoteScriptHash)" }
                    if ((-not [string]::IsNullOrWhiteSpace($result.LocalScriptVersion)) -and $result.RemoteScriptVersion -ne $result.LocalScriptVersion) { throw "Final remote script version mismatch. LocalVersion=$($result.LocalScriptVersion); RemoteVersion=$($result.RemoteScriptVersion)" }
                }
                catch {
                    $result.ScriptCopied = $false
                    $payloadFailureStatus = "REMOTE_SCRIPT_COPY_FAILED"
                    $payloadFailureDetail = "Remote script staged copy failed: $remoteAdminScript; Error=$($_.Exception.Message)"
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] WARN: $payloadFailureDetail" | Add-Content -LiteralPath $logPath -Encoding UTF8
                    continue
                }
                finally {
                    if (Test-Path -LiteralPath $remoteStagingScript) {
                        Remove-Item -LiteralPath $remoteStagingScript -Force -ErrorAction SilentlyContinue
                    }
                }

                $payloadReady = $true
                break
            }

            if (-not $payloadReady) {
                if ([string]::IsNullOrWhiteSpace($payloadFailureStatus)) { $payloadFailureStatus = "REMOTE_SCRIPT_COPY_FAILED" }
                if ([string]::IsNullOrWhiteSpace($payloadFailureDetail)) { $payloadFailureDetail = "Remote script copy failed after $maxPayloadCopyAttempts attempt(s)." }
                $result.Status = $payloadFailureStatus
                $result.NextAction = Get-NextActionFromLauncherStatus -Status $result.Status
                $result.RemoteDetail = $payloadFailureDetail
                $result.ErrorMessage = $result.RemoteDetail
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERROR: $($result.RemoteDetail)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                return (Complete-WorkerResult -Result $result -Path $logPath)
            }

            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Remote script copied and verified: Version=$($result.RemoteScriptVersion); SHA256=$($result.RemoteScriptHash); Path=$remoteAdminScript" | Add-Content -LiteralPath $logPath -Encoding UTF8

            $remoteScriptArgs = @($CycleScriptArgs)
            if ($result.EntraRegisteredState -eq "Pending") {
                $remoteScriptArgs += "-EntraHybridPending"
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Entra inventory state is Pending. Passing -EntraHybridPending to remote script. AlternativeSecurityIdCount=$($result.EntraAlternativeSecurityIdCount); Reason=$($result.EntraPendingReason)" | Add-Content -LiteralPath $logPath -Encoding UTF8
            }

            $argsList = @(
                "\\$ConnectionTarget",
                "-accepteula",
                "-nobanner",
                "-s",
                "-h",
                "powershell.exe",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                $RemoteScriptPath
            ) + $remoteScriptArgs

            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] PsExec command: $PsExecPath $($argsList -join ' ')" | Add-Content -LiteralPath $logPath -Encoding UTF8

            $stdoutPath = Join-Path $LogRoot ("{0}_cycle{1}_{2}_stdout.tmp" -f $safeComputerName,$CycleNumber,$runId)
            $stderrPath = Join-Path $LogRoot ("{0}_cycle{1}_{2}_stderr.tmp" -f $safeComputerName,$CycleNumber,$runId)
            $process = Start-Process -FilePath $PsExecPath -ArgumentList $argsList -PassThru -NoNewWindow -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
            try {
                Invoke-WithWorkerLeaseMutex -MutexName $GlobalWorkerLeaseMutexName -ScriptBlock {
                    if (Test-Path -LiteralPath $GlobalWorkerLeasePath -PathType Leaf) {
                        $leaseData = Read-WorkerLeaseData -Path $GlobalWorkerLeasePath
                        $leaseData | Add-Member -NotePropertyName PsExecProcessId -NotePropertyValue ([int]$process.Id) -Force
                        $leaseData | Add-Member -NotePropertyName PsExecStartedUtc -NotePropertyValue ((Get-Date).ToUniversalTime().ToString("o")) -Force
                        $leaseData | Add-Member -NotePropertyName LastUpdatedUtc -NotePropertyValue ((Get-Date).ToUniversalTime().ToString("o")) -Force
                        Save-WorkerLeaseData -Path $GlobalWorkerLeasePath -Data $leaseData
                    }
                }
            }
            catch { "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] WARN: PsExec PID could not be written to the worker lease: $($_.Exception.Message)" | Add-Content -LiteralPath $logPath -Encoding UTF8 }
            $psExecTimedOut = $false
            if ($PsExecTimeoutMinutes -gt 0) {
                $timeoutMs = [int64]$PsExecTimeoutMinutes * 60 * 1000
                if (-not $process.WaitForExit($timeoutMs)) {
                    $psExecTimedOut = $true
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERROR: PsExec timed out after $PsExecTimeoutMinutes minute(s). Killing local PsExec process." | Add-Content -LiteralPath $logPath -Encoding UTF8
                    try { $process.Kill() } catch { }
                    try { [void]$process.WaitForExit(5000) } catch { }
                }
            }
            else {
                [void]$process.WaitForExit()
            }
            try { $process.Refresh() } catch { }
            $exitCode = ""
            if (-not $psExecTimedOut) {
                try {
                    if ($process.HasExited) {
                        $exitCode = [string]$process.ExitCode
                    }
                }
                catch {
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] WARN: Could not read PsExec exit code: $($_.Exception.Message)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                }
            }

            $stdoutContent = @()
            if (Test-Path -LiteralPath $stdoutPath) {
                "----- PsExec STDOUT -----" | Add-Content -LiteralPath $logPath -Encoding UTF8
                $stdoutContent = @(Get-Content -LiteralPath $stdoutPath -ErrorAction SilentlyContinue)
                $stdoutContent | Add-Content -LiteralPath $logPath -Encoding UTF8
                Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue
            }
            if (Test-Path -LiteralPath $stderrPath) {
                "----- PsExec STDERR -----" | Add-Content -LiteralPath $logPath -Encoding UTF8
                $stderrContent = @(Get-Content -LiteralPath $stderrPath -ErrorAction SilentlyContinue)
                $stderrContent | Add-Content -LiteralPath $logPath -Encoding UTF8
                Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue
            }
            else {
                $stderrContent = @()
            }

            if ([string]::IsNullOrWhiteSpace($exitCode)) {
                $nativeExitLine = ($stderrContent | Where-Object { $_ -match "with error code\s+-?\d+" } | Select-Object -Last 1)
                if ($nativeExitLine -and $nativeExitLine -match "with error code\s+(?<Code>-?\d+)") {
                    $exitCode = $Matches.Code
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] PsExec exit code recovered from native STDERR: $exitCode" | Add-Content -LiteralPath $logPath -Encoding UTF8
                }
            }

            $result.PsExecExitCode = $exitCode
            $finalLine = $stdoutContent | Where-Object { $_ -match "^FINAL_STATUS=" } | Select-Object -Last 1
            if ($psExecTimedOut) {
                $result.Status = "PSEXEC_TIMEOUT"
                $result.RemoteDetail = "PsExec timed out after $PsExecTimeoutMinutes minute(s). The remote process may still need log review on the target."
                $result.ErrorMessage = $result.RemoteDetail
            }
            elseif ($finalLine -and $finalLine -match "^FINAL_STATUS=(?<Status>[^;]*);\s*EXIT_CODE=(?<ExitCode>[^;]*);\s*NEXT_ACTION=(?<NextAction>[^;]*);\s*DETAIL=(?<Detail>.*)$") {
                $result.RemoteStatus = $Matches.Status.Trim()
                $result.RemoteExitCode = $Matches.ExitCode.Trim()
                $result.RemoteNextAction = $Matches.NextAction.Trim()
                $result.RemoteDetail = $Matches.Detail.Trim()
                if (-not [string]::IsNullOrWhiteSpace($result.RemoteStatus)) {
                    $result.Status = $result.RemoteStatus
                }
                else {
                    $result.Status = if ($exitCode -eq "0") { "SUCCESS" } elseif ([string]::IsNullOrWhiteSpace($exitCode)) { "PSEXEC_EXIT_UNKNOWN" } else { "PSEXEC_EXIT_$exitCode" }
                }
            }
            else {
                $result.Status = if ($exitCode -eq "0") { "SUCCESS" } elseif ([string]::IsNullOrWhiteSpace($exitCode)) { "PSEXEC_EXIT_UNKNOWN" } else { "PSEXEC_EXIT_$exitCode" }
                $combinedNativeOutput = (($stdoutContent + $stderrContent) -join "`n")
                if ($combinedNativeOutput -match "(?i)(Error communicating with PsExec service|Descripteur non valide|handle is invalid)") {
                    $result.Status = "PSEXEC_COMMUNICATION_LOST"
                    $result.RemoteDetail = "PsExec lost communication with PSEXESVC after starting remote PowerShell. Remote evidence may exist, but no FINAL_STATUS/current-run CSV was returned to the launcher."
                    $result.ErrorMessage = $result.RemoteDetail
                }
                if ($combinedNativeOutput -match "(?i)(-File.*(does not exist|n.?existe|non esiste|no existe|n.?o existe|nie istnieje|nicht.*exist)|fichier.*sp.cifi..*introuvable|file.*specified.*not.*found|impossibile trovare il file specificato)") {
                    $result.Status = "REMOTE_SCRIPT_MISSING"
                    $result.RemoteDetail = "PowerShell on the remote computer reported that the -File script path does not exist: $RemoteScriptPath"
                    $result.ErrorMessage = $result.RemoteDetail
                }
                $fatalLine = $stdoutContent | Where-Object { $_ -match "FATAL ERROR:" } | Select-Object -Last 1
                if ($fatalLine) {
                    $result.RemoteDetail = ([string]$fatalLine).Trim()
                    $result.ErrorMessage = $result.RemoteDetail
                }
            }
            if ($result.Status -eq "PSEXEC_EXIT_-1073741790") {
                $accessDeniedDetail = "Remote PowerShell exited with 0xC0000022 (STATUS_ACCESS_DENIED) after PsExec started it. Check EDR, AppLocker, WDAC, PowerShell policy, or local execution permissions on the target."
                if ([string]::IsNullOrWhiteSpace($result.RemoteDetail)) { $result.RemoteDetail = $accessDeniedDetail }
                if ([string]::IsNullOrWhiteSpace($result.ErrorMessage)) { $result.ErrorMessage = $accessDeniedDetail }
            }
            $derivedNextAction = Get-NextActionFromLauncherStatus -Status $result.Status
            if ((-not [string]::IsNullOrWhiteSpace($result.RemoteNextAction)) -and $result.RemoteNextAction -ne "REVIEW_LOGS") {
                $result.NextAction = $result.RemoteNextAction
            }
            else {
                $result.NextAction = $derivedNextAction
            }
            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] PsExec exit code: $exitCode" | Add-Content -LiteralPath $logPath -Encoding UTF8
            if (-not [string]::IsNullOrWhiteSpace($result.RemoteStatus)) {
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Remote final status: $($result.RemoteStatus); Detail=$($result.RemoteDetail)" | Add-Content -LiteralPath $logPath -Encoding UTF8
            }

            if ($CollectRemoteLogs) {
                try {
                    $polledRemoteFinalStatus = $null
                    if ($result.Status -eq "PSEXEC_COMMUNICATION_LOST" -and $CommunicationLostEvidenceWaitMinutes -gt 0) {
                        $elapsedWaitMinutes = 0
                        $pollMinutes = [Math]::Min($CommunicationLostEvidencePollMinutes, $CommunicationLostEvidenceWaitMinutes)
                        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] PsExec communication was lost after remote start. Polling every $pollMinutes minute(s), up to $CommunicationLostEvidenceWaitMinutes minute(s), for current-run final CSV before delayed evidence collection." | Add-Content -LiteralPath $logPath -Encoding UTF8
                        while ($elapsedWaitMinutes -lt $CommunicationLostEvidenceWaitMinutes) {
                            $remoteFinalStatus = $null
                            if (Test-Path -LiteralPath $remoteDataAdminDir -ErrorAction SilentlyContinue) {
                                $remoteFinalStatus = Get-RemoteEvidenceFinalStatus -EvidencePath $remoteDataAdminDir -Since ([datetime]$result.Timestamp) -RequireCompletedRun
                            }
                            if ($null -ne $remoteFinalStatus -and -not [string]::IsNullOrWhiteSpace($remoteFinalStatus.Status)) {
                                $polledRemoteFinalStatus = $remoteFinalStatus
                                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Completed current-run evidence detected on remote computer after $elapsedWaitMinutes minute(s). RunId=$($remoteFinalStatus.RunId); Status=$($remoteFinalStatus.Status); Csv=$($remoteFinalStatus.CsvPath); LastRun=$($remoteFinalStatus.LastRunPath). Collecting evidence now." | Add-Content -LiteralPath $logPath -Encoding UTF8
                                break
                            }

                            $remainingMinutes = $CommunicationLostEvidenceWaitMinutes - $elapsedWaitMinutes
                            $sleepMinutes = [Math]::Min($pollMinutes, $remainingMinutes)
                            if ($sleepMinutes -le 0) { break }
                            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Current-run final CSV not found yet. Sleeping $sleepMinutes minute(s) before next check." | Add-Content -LiteralPath $logPath -Encoding UTF8
                            Start-Sleep -Seconds ([int]($sleepMinutes * 60))
                            $elapsedWaitMinutes += $sleepMinutes
                        }

                        if ($null -ne $polledRemoteFinalStatus -and -not [string]::IsNullOrWhiteSpace($polledRemoteFinalStatus.Status)) {
                            $result.RemoteStatus = $polledRemoteFinalStatus.Status
                            $result.RemoteExitCode = $polledRemoteFinalStatus.ExitCode
                            $result.RemoteNextAction = $polledRemoteFinalStatus.NextAction
                            $result.RemoteDetail = $polledRemoteFinalStatus.Detail
                            $result.Status = $polledRemoteFinalStatus.Status
                            if (-not [string]::IsNullOrWhiteSpace($polledRemoteFinalStatus.NextAction)) {
                                $result.NextAction = $polledRemoteFinalStatus.NextAction
                            }
                            else {
                                $result.NextAction = Get-NextActionFromLauncherStatus -Status $result.Status
                            }
                            if (-not [string]::IsNullOrWhiteSpace($polledRemoteFinalStatus.Detail)) {
                                $result.ErrorMessage = $polledRemoteFinalStatus.Detail
                            }
                            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] PsExec communication-lost result reclassified before evidence copy. RunId=$($polledRemoteFinalStatus.RunId); Status=$($result.Status); NextAction=$($result.NextAction)." | Add-Content -LiteralPath $logPath -Encoding UTF8
                        }
                    }
                    elseif ($result.Status -eq "PSEXEC_EXIT_UNKNOWN") {
                        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] PsExec returned no final status. Waiting 30 seconds before collecting remote evidence to allow the remote script to flush CSV/log/gpresult files." | Add-Content -LiteralPath $logPath -Encoding UTF8
                        Start-Sleep -Seconds 30
                    }

                    if (-not (Test-Path -LiteralPath $remoteDataAdminDir)) {
                        throw "Remote log folder not found: $remoteDataAdminDir"
                    }

                    $centralStatus = ([string]$result.Status).ToUpperInvariant()
                    $centralBucket = if ($centralStatus -in @("SUCCESS","AUDIT_SUCCESS_ALREADY_INTUNE","ENROLLED_DETECTED_POST_CYCLE")) { "Success" } elseif ($centralStatus -like "ADMIN_SHARE*") { "AdminShareFailure" } elseif ($centralStatus -like "PSEXEC_*" -or $centralStatus -like "REMOTE_*" -or $centralStatus -like "DNS_*") { "RemoteCollectionFailure" } else { "Errors" }
                    if (-not $KeepCentralLogHistory) {
                        foreach ($bucketName in @("Success","Errors","AdminShareFailure","RemoteCollectionFailure")) {
                            $staleComputerDir = Join-Path (Join-Path $CentralLogRoot $bucketName) $safeComputerName
                            Remove-Item -LiteralPath $staleComputerDir -Recurse -Force -ErrorAction SilentlyContinue
                        }
                    }
                    $centralComputerDir = Join-Path (Join-Path $CentralLogRoot $centralBucket) $safeComputerName
                    $centralRunDir = if ($KeepCentralLogHistory) { Join-Path $centralComputerDir ("cycle{0}_{1}" -f $CycleNumber,$runId) } else { Join-Path $centralComputerDir "Latest" }
                    $copiedEvidenceFiles = Copy-RemoteEvidenceFolder -RemoteDataPath $remoteDataAdminDir -DestinationPath $centralRunDir -ScriptName $ScriptName -CentralLogCollectionMode $CentralLogCollectionMode -Since ([datetime]$result.Timestamp)
                    $result.RemoteLogsCollected = $true
                    $result.RemoteLogsPath = $centralRunDir
                    $result.RemoteCurrentRunLogsPath = $centralRunDir
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Remote evidence collected to: $centralRunDir. Files=$copiedEvidenceFiles; Mode=$CentralLogCollectionMode; CurrentRunPath=$centralRunDir" | Add-Content -LiteralPath $logPath -Encoding UTF8
                    $completedEvidenceStatus = Get-RemoteEvidenceFinalStatus -EvidencePath $centralRunDir -Since ([datetime]$result.Timestamp) -RequireCompletedRun
                    if ($null -ne $completedEvidenceStatus -and -not [string]::IsNullOrWhiteSpace($completedEvidenceStatus.Status)) {
                        $result.InteractiveUserName = $completedEvidenceStatus.InteractiveUserName
                        $result.InteractiveUserDomain = $completedEvidenceStatus.InteractiveUserDomain
                        $result.InteractiveUserAccountName = $completedEvidenceStatus.InteractiveUserAccountName
                        $result.InteractiveUserAccountType = $completedEvidenceStatus.InteractiveUserAccountType
                        $result.InteractiveSessionName = $completedEvidenceStatus.InteractiveSessionName
                        $result.InteractiveSessionState = $completedEvidenceStatus.InteractiveSessionState
                        $result.UserIsUserAzureAD = $completedEvidenceStatus.UserIsUserAzureAD
                        $result.UserAzureAdPrt = $completedEvidenceStatus.UserAzureAdPrt
                        $result.UserSessionIsNotRemote = $completedEvidenceStatus.UserSessionIsNotRemote
                        $result.RetryAfterRebootAction = $completedEvidenceStatus.RetryAfterRebootAction
                        $result.RetryAfterRebootDetail = $completedEvidenceStatus.RetryAfterRebootDetail
                        $result.RetryAfterRebootAttempt = $completedEvidenceStatus.RetryAfterRebootAttempt
                        $result.RetryAfterRebootMaxAttempts = $completedEvidenceStatus.RetryAfterRebootMaxAttempts
                        $result.RetryAfterRebootTaskName = $completedEvidenceStatus.RetryAfterRebootTaskName
                        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Remote final evidence enriched report row. RunId=$($completedEvidenceStatus.RunId); User=$($completedEvidenceStatus.InteractiveUserAccountName); IsUserAzureAD=$($completedEvidenceStatus.UserIsUserAzureAD); AzureAdPrt=$($completedEvidenceStatus.UserAzureAdPrt)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                    }

                    if ($result.Status -in @("PSEXEC_EXIT_UNKNOWN","PSEXEC_COMMUNICATION_LOST")) {
                        $evidenceStatus = $completedEvidenceStatus
                        if ($null -ne $evidenceStatus -and -not [string]::IsNullOrWhiteSpace($evidenceStatus.Status)) {
                            $result.RemoteStatus = $evidenceStatus.Status
                            $result.RemoteExitCode = $evidenceStatus.ExitCode
                            $result.RemoteNextAction = $evidenceStatus.NextAction
                            $result.RemoteDetail = $evidenceStatus.Detail
                            $result.Status = $evidenceStatus.Status
                            if (-not [string]::IsNullOrWhiteSpace($evidenceStatus.NextAction)) {
                                $result.NextAction = $evidenceStatus.NextAction
                            }
                            else {
                                $result.NextAction = Get-NextActionFromLauncherStatus -Status $result.Status
                            }
                            if (-not [string]::IsNullOrWhiteSpace($evidenceStatus.Detail)) {
                                $result.ErrorMessage = $evidenceStatus.Detail
                            }
                            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] PsExec no-final-status result reclassified from completed current-run evidence. RunId=$($evidenceStatus.RunId); Status=$($result.Status); NextAction=$($result.NextAction); Csv=$($evidenceStatus.CsvPath); LastRun=$($evidenceStatus.LastRunPath)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                        }
                        else {
                            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] No current-run final CSV was found in collected evidence. Keeping Status=$($result.Status); NextAction=$($result.NextAction)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                        }
                    }
                }
                catch {
                    $result.RemoteLogsError = $_.Exception.Message
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] WARN: Remote log collection failed: $($result.RemoteLogsError)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                }
            }
        }
        catch {
            $result.Status = "ERROR"
            $result.NextAction = "CHECK_CONNECTIVITY_OR_ADMIN_ACCESS"
            $message = $_.Exception.Message
            $result.ErrorMessage = $message
            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERROR: $message" | Add-Content -LiteralPath $logPath -Encoding UTF8
        }

        Complete-WorkerResult -Result $result -Path $logPath
    }

    $runningJobs = @()
    $globalLeaseByJobId = @{}
    $jobStartedAtById = @{}
    $techRunGuardFqdnByJobId = @{}
    $forcedCancelledJobIds = @{}
    $cancellationObservedAt = $null
    $nextIndex = 0
    $completed = 0
    $lastProgressLog = Get-Date

    while ($nextIndex -lt $computers.Count -or $runningJobs.Count -gt 0) {
        $cancelState = Get-LotCancellationState
        if ($cancelState.Requested) {
            if ($null -eq $cancellationObservedAt) {
                $cancellationObservedAt = Get-Date
                Set-ActiveLotRunState -Status 'StopRequested' -ReportPath $script:MergedHtmlReportPath
                Write-Host ("Controlled stop requested by {0}: no additional computers will start. Active jobs will drain for up to {1} minute(s). Press Ctrl+C again to force their local stop." -f $cancelState.Source,$CancellationDrainTimeoutMinutes) -ForegroundColor Yellow
            }

            while ($nextIndex -lt $computers.Count) {
                $cancelledComputer = $computers[$nextIndex]
                $nextIndex++
                $cancelledResult = New-HybridJoinCancellationResult -ComputerName $cancelledComputer -CycleNumber $CycleNumber -Status 'CANCELLED_NOT_STARTED' -Detail 'The operator requested a controlled stop before this computer was queued.' -NextAction 'SAFE_TO_RELAUNCH'
                $summary.Add($cancelledResult)
                Add-LiveCycleReportRow -Path $liveSummaryPath -Columns $reportColumns -Row $cancelledResult
                $completed++
            }

            $drainExpired = ($CancellationDrainTimeoutMinutes -eq 0 -or ((Get-Date) - $cancellationObservedAt).TotalMinutes -ge $CancellationDrainTimeoutMinutes)
            if ($runningJobs.Count -gt 0 -and ($cancelState.Force -or $drainExpired)) {
                foreach ($runningJob in @($runningJobs | Where-Object { $_.State -eq 'Running' })) {
                    $runningJobId = [string]$runningJob.Id
                    if (-not $forcedCancelledJobIds.ContainsKey($runningJobId)) {
                        $forcedCancelledJobIds[$runningJobId] = $true
                        Write-Host ("Forcing local worker stop for {0}. The remote endpoint may already be running and must be verified before relaunch." -f ($runningJob.Name -replace '^EHJIR_C\d+_','')) -ForegroundColor Red
                        $leaseForForcedJob = if ($globalLeaseByJobId.ContainsKey($runningJobId)) { [string]$globalLeaseByJobId[$runningJobId] } else { "" }
                        $processTreeStopDetail = Stop-LocalPsExecProcessTreeFromLease -LeasePath $leaseForForcedJob
                        Write-Host ("Forced local process cleanup for {0}: {1}" -f ($runningJob.Name -replace '^EHJIR_C\d+_',''),$processTreeStopDetail) -ForegroundColor Red
                    }
                    Stop-Job -Job $runningJob -ErrorAction SilentlyContinue
                }
            }
        }

        while ($nextIndex -lt $computers.Count -and $runningJobs.Count -lt $ThrottleLimit -and -not (Get-LotCancellationState).Requested) {
            $computer = $computers[$nextIndex]
            $connectionTargetInfo = Resolve-ComputerConnectionTarget -ComputerName $computer -DomainSuffix $AdDomain
            $connectionTarget = $connectionTargetInfo.ConnectionTarget
            $nextIndex++

            $techRunGuardFqdn = Get-TechnicianRunGuardFqdn -ComputerName $computer -AdInventoryMap $script:AdInventoryMap
            if ($script:UseEffectiveTechnicianRunGuardHistory) {
                $activeTechRunGuard = $null
                try {
                    $activeTechRunGuard = Get-ActiveTechnicianRunGuardEntry -Path $script:TechnicianRunGuardHistoryPath -ComputerFqdn $techRunGuardFqdn -Hours $TechnicianRunGuardHours
                }
                catch {
                    $historyError = "Technician run guard history is unavailable; this computer was not launched. Error=$($_.Exception.Message)"
                    $guardUnavailableResult = New-HybridJoinCancellationResult -ComputerName $computer -CycleNumber $CycleNumber -Status 'TECH_RUN_GUARD_HISTORY_UNAVAILABLE' -Detail $historyError
                    $summary.Add($guardUnavailableResult)
                    Add-LiveCycleReportRow -Path $liveSummaryPath -Columns $reportColumns -Row $guardUnavailableResult
                    $completed++
                    Write-Host ("Skipped {0}/{1}: {2} => TECH_RUN_GUARD_HISTORY_UNAVAILABLE" -f $completed,$computers.Count,$computer) -ForegroundColor Yellow
                    continue
                }
                if ($activeTechRunGuard) {
                    $skipResult = New-TechnicianRunGuardSkippedResult -ComputerName $computer -CycleNumber $CycleNumber -HistoryEntry $activeTechRunGuard
                    $summary.Add($skipResult)
                    Add-LiveCycleReportRow -Path $liveSummaryPath -Columns $reportColumns -Row $skipResult
                    $completed++
                    Write-Host ("Skipped {0}/{1}: {2} => {3}" -f $completed,$computers.Count,$computer,$skipResult.Status) -ForegroundColor Yellow
                    continue
                }
            }

            $globalLeasePath = Acquire-GlobalWorkerLease -Computer $computer -CycleNumber $CycleNumber
            if ((Get-LotCancellationState).Requested) {
                Release-GlobalWorkerLease -LeasePath $globalLeasePath
                $nextIndex--
                break
            }


            try {
                $cycleScriptArgsJson = ([pscustomobject]@{ Args = @($CycleScriptArgs) } | ConvertTo-Json -Compress)
                $jobName = "EHJIR_C{0}_{1}" -f $CycleNumber,$computer
                $jobStart = Start-LocalWorkerJobWithRetry -ComputerName $computer -JobName $jobName -StartOperation {
                    Start-Job -Name $jobName -ScriptBlock $worker -ArgumentList @(
                    $computer,
                    $connectionTarget,
                    $CycleNumber,
                    $LocalScriptPath,
                    $ScriptName,
                    $RemoteRelativeDir,
                    $RemoteScriptPath,
                    $RemoteDataRelativeDir,
                    $PsExecPath,
                    $LogRoot,
                    $LauncherVersion,
                    [bool]$DryRun,
                    $CollectRemoteLogs,
                    [bool]$SkipVirtualMachines,
                    $CentralLogRoot,
                    [bool]$KeepCentralLogHistory,
                    $CentralLogCollectionMode,
                    $IntuneInventorySet,
                    $EntraInventoryMap,
                    $AdInventoryMap,
                    $cycleScriptArgsJson,
                    $PsExecTimeoutMinutes,
                    $CommunicationLostEvidenceWaitMinutes,
                    $CommunicationLostEvidencePollMinutes,
                    $globalLeasePath,
                    $globalConcurrencyMutexName
                )
                }
                if (-not $jobStart.Succeeded) {
                    Release-GlobalWorkerLease -LeasePath $globalLeasePath
                    $globalLeasePath = ''
                    $failureDetail = "Local worker could not be started after 3 attempts. The LOT continued with the next computer. $($jobStart.Detail)"
                    $failureResult = New-HybridJoinCancellationResult -ComputerName $computer -CycleNumber $CycleNumber -Status 'LOCAL_WORKER_START_FAILED' -Detail $failureDetail
                    $failureResult.ErrorMessage = $failureDetail
                    $summary.Add($failureResult)
                    Add-LiveCycleReportRow -Path $liveSummaryPath -Columns $reportColumns -Row $failureResult
                    $completed++
                    Write-Host ("Completed {0}/{1}: {2} => LOCAL_WORKER_START_FAILED; NextAction=VERIFY_REMOTE_STATE_BEFORE_RELAUNCH; Detail={3}" -f $completed,$computers.Count,$computer,$failureDetail) -ForegroundColor Red
                    continue
                }
                $job = $jobStart.Job
                $jobStartedAtById[[string]$job.Id] = Get-Date
                if ($script:UseEffectiveTechnicianRunGuardHistory) {
                    $techRunGuardFqdnByJobId[[string]$job.Id] = $techRunGuardFqdn
                    try {
                        Update-TechnicianRunGuardHistory -Path $script:TechnicianRunGuardHistoryPath -ComputerFqdn $techRunGuardFqdn -InputComputerName $computer -Hours $TechnicianRunGuardHours -State Started -Result $null -JobId ([string]$job.Id) -CycleNumber $CycleNumber
                    }
                    catch {
                        Write-Host ("TECH_RUN_GUARD_HISTORY_WRITE_DEFERRED: State=Started; Computer={0}; Error={1}" -f $computer,$_.Exception.Message) -ForegroundColor Yellow
                    }
                }
            }
            catch {
                Release-GlobalWorkerLease -LeasePath $globalLeasePath
                throw
            }

            if (-not [string]::IsNullOrWhiteSpace($globalLeasePath)) {
                Update-GlobalWorkerLease -LeasePath $globalLeasePath -Properties @{
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

        $finishedJobs = @($runningJobs | Where-Object { $_.State -ne "Running" })
        if ($finishedJobs.Count -eq 0) {
            if (((Get-Date) - $lastProgressLog).TotalSeconds -ge 300) {
                $now = Get-Date
                $waitingNames = @($runningJobs | ForEach-Object {
                    $jobId = [string]$_.Id
                    $jobName = $_.Name -replace '^EHJIR_C\d+_',''
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
                    $liveRows = @($summary | ForEach-Object { $_ })
                    $cycleProgress = New-HybridJoinCycleProgressRows -CycleNumber $CycleNumber -CycleStart $cycleStart -TotalComputers $computers.Count -QueuedComputers $nextIndex -CompletedComputers $completed -RunningComputers $runningJobs.Count -ComputerListStats $computerListStats
                    $runningJobRows = New-HybridJoinRunningJobRows -RunningJobs @($runningJobs) -JobStartedAtById $jobStartedAtById
                    $mergedLiveRows = @($script:AllCycleResults.ToArray()) + $liveRows
                    $mergedProgressRows = @($script:AllCycleProgressRows.ToArray()) + @($cycleProgress)
                    New-CycleHtmlReport -Summary $mergedLiveRows -Path $script:MergedHtmlReportPath -CycleNumber $CycleNumber -GeneratedAt (Get-Date) -IsLive -CycleProgress $mergedProgressRows -RunningJobRows $runningJobRows
                    $lastLiveHtmlWrite = Get-Date
                }
                catch {
                    Write-Host ("Cycle {0}: failed to update merged HTML report: {1}" -f $CycleNumber,$_.Exception.Message) -ForegroundColor Yellow
                }
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
                    $childJobs[0].JobStateInfo.Reason -ne $null -and
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
                            "RUNSPACE_BROKEN: PowerShell job runspace entered a Broken state. Engine=$($PSVersionTable.PSEdition) $($PSVersionTable.PSVersion). This indicates parallel runspace instability; consider adding -DelayBetweenComputersSeconds 1 to reduce job cycling speed."
                        }
                    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
                )
                if ($brokenRunspace) {
                    Write-Host ("  [RUNSPACE_BROKEN] {0}: job runspace entered Broken state under {1} {2}. Consider -DelayBetweenComputersSeconds 1." -f ($job.Name -replace "^EHJIR_C\d+_",""),$PSVersionTable.PSEdition,$PSVersionTable.PSVersion) -ForegroundColor Magenta
                }
            }
            catch {
                $jobErrors += $_.Exception.Message
            }

            if (-not $received) {
                $forcedCancellation = $forcedCancelledJobIds.ContainsKey([string]$job.Id)
                $forcedComputer = ($job.Name -replace "^EHJIR_C\d+_","")
                $forcedEvidence = $null
                $forcedRemoteDataPath = ""
                if ($forcedCancellation) {
                    try {
                        $forcedTarget = (Resolve-ComputerConnectionTarget -ComputerName $forcedComputer -DomainSuffix $AdDomain).ConnectionTarget
                        $forcedRemoteDataPath = "\\$forcedTarget\C$\$RemoteDataRelativeDir"
                        $forcedSince = if ($jobStartedAtById.ContainsKey([string]$job.Id)) { [datetime]$jobStartedAtById[[string]$job.Id] } else { (Get-Date).AddMinutes(-5) }
                        for ($evidenceAttempt = 1; $evidenceAttempt -le 3; $evidenceAttempt++) {
                            if (Test-Path -LiteralPath $forcedRemoteDataPath -ErrorAction SilentlyContinue) {
                                $forcedEvidence = Get-RemoteEvidenceFinalStatus -EvidencePath $forcedRemoteDataPath -Since $forcedSince -RequireCompletedRun
                            }
                            if ($forcedEvidence -and -not [string]::IsNullOrWhiteSpace([string]$forcedEvidence.Status)) { break }
                            if ($evidenceAttempt -lt 3) { Start-Sleep -Seconds 5 }
                        }
                    }
                    catch { $jobErrors += ("Forced-stop evidence recheck failed: {0}" -f $_.Exception.Message) }
                }
                $forcedEvidenceFound = ($forcedEvidence -and -not [string]::IsNullOrWhiteSpace([string]$forcedEvidence.Status))
                $forcedStatus = if ($forcedEvidenceFound) { [string]$forcedEvidence.Status } elseif ($forcedCancellation) { "CANCELLED_BY_OPERATOR_REMOTE_STATE_UNCONFIRMED" } else { "JOB_ERROR" }
                $forcedNextAction = if ($forcedEvidenceFound -and -not [string]::IsNullOrWhiteSpace([string]$forcedEvidence.NextAction)) { [string]$forcedEvidence.NextAction } elseif ($forcedCancellation) { "VERIFY_REMOTE_STATE_BEFORE_RELAUNCH" } else { "CHECK_JOB_ERROR" }
                $forcedDetail = if ($forcedEvidenceFound) { "Forced local worker stop reclassified from completed current-run endpoint evidence. " + [string]$forcedEvidence.Detail } elseif ($forcedCancellation) { "The operator forced the local PsExec process tree and worker to stop. No completed current-run endpoint evidence was found after three bounded checks." } else { "" }
                $received = [PSCustomObject]@{
                    LauncherVersion = $LauncherVersion
                    Cycle = $CycleNumber
                    Computer = $forcedComputer
                    Timestamp = Get-Date
                    DryRun = [bool]$DryRun
                    DnsResolved = $false
                    DnsAddressList = ""
                    AdminShareReachable = $false
                    RemotePayloadCopyAttempts = ""
                    PingReachable = $false
                    RemoteDirectoryCreated = $false
                    ScriptCopied = $false
                    PsExecExitCode = ""
                    RemoteStatus = if ($forcedEvidenceFound) { [string]$forcedEvidence.Status } else { "" }
                    RemoteExitCode = if ($forcedEvidenceFound) { [string]$forcedEvidence.ExitCode } else { "" }
                    RemoteNextAction = if ($forcedEvidenceFound) { [string]$forcedEvidence.NextAction } else { $forcedNextAction }
                    RemoteDetail = $forcedDetail
                    NextAction = $forcedNextAction
                    EffectiveStatus = $forcedStatus
                    EffectiveNextAction = $forcedNextAction
                    InteractiveUserName = if ($forcedEvidenceFound) { [string]$forcedEvidence.InteractiveUserName } else { "" }
                    InteractiveUserDomain = if ($forcedEvidenceFound) { [string]$forcedEvidence.InteractiveUserDomain } else { "" }
                    InteractiveUserAccountName = if ($forcedEvidenceFound) { [string]$forcedEvidence.InteractiveUserAccountName } else { "" }
                    InteractiveUserAccountType = if ($forcedEvidenceFound) { [string]$forcedEvidence.InteractiveUserAccountType } else { "" }
                    InteractiveSessionName = if ($forcedEvidenceFound) { [string]$forcedEvidence.InteractiveSessionName } else { "" }
                    InteractiveSessionState = if ($forcedEvidenceFound) { [string]$forcedEvidence.InteractiveSessionState } else { "" }
                    UserIsUserAzureAD = if ($forcedEvidenceFound) { [string]$forcedEvidence.UserIsUserAzureAD } else { "" }
                    UserAzureAdPrt = if ($forcedEvidenceFound) { [string]$forcedEvidence.UserAzureAdPrt } else { "" }
                    UserSessionIsNotRemote = if ($forcedEvidenceFound) { [string]$forcedEvidence.UserSessionIsNotRemote } else { "" }
                    RetryAfterRebootAction = if ($forcedEvidenceFound) { [string]$forcedEvidence.RetryAfterRebootAction } else { "" }
                    RetryAfterRebootDetail = if ($forcedEvidenceFound) { [string]$forcedEvidence.RetryAfterRebootDetail } else { "" }
                    RetryAfterRebootAttempt = if ($forcedEvidenceFound) { [string]$forcedEvidence.RetryAfterRebootAttempt } else { "" }
                    RetryAfterRebootMaxAttempts = if ($forcedEvidenceFound) { [string]$forcedEvidence.RetryAfterRebootMaxAttempts } else { "" }
                    RetryAfterRebootTaskName = if ($forcedEvidenceFound) { [string]$forcedEvidence.RetryAfterRebootTaskName } else { "" }
                    IntuneInventoryPresent = ""
                    EntraInventoryPresent = ""
                    EntraRegisteredState = ""
                    EntraAlternativeSecurityIdCount = ""
                    EntraPendingReason = ""
                    EntraRegistrationDateTime = ""
                    EntraTrustType = ""
                    EntraDeviceId = ""
                    EntraObjectId = ""
                    ADInventoryPresent = ""
                    ADDomain = ""
                    ADEnabled = ""
                    ADDNSHostName = ""
                    ADDistinguishedName = ""
                    ADOperatingSystem = ""
                    ADLastLogonTimestampUtc = ""
                    AdminShareFailureType = ""
                    PostCycleIntuneInventoryChecked = ""
                    PostCycleIntuneInventoryPresent = ""
                    PostCycleIntuneEnrollmentDetected = ""
                    PostCycleIntuneInventoryCsv = ""
                    PostCycleIntuneInventoryError = ""
                    PostCycleEntraInventoryChecked = ""
                    PostCycleEntraInventoryPresent = ""
                    PostCycleEntraRegisteredState = ""
                    PostCycleEntraAlternativeSecurityIdCount = ""
                    PostCycleEntraPendingResolved = ""
                    PostCycleEntraInventoryCsv = ""
                    PostCycleEntraInventoryError = ""
                    PostCycleADInventoryChecked = ""
                    PostCycleADInventoryPresent = ""
                    PostCycleADInventoryCsv = ""
                    PostCycleADInventoryError = ""
                    RemoteLogsCollected = $false
                    RemoteLogsPath = ""
                    RemoteCurrentRunLogsPath = if ($forcedEvidenceFound) { $forcedRemoteDataPath } else { "" }
                    RemoteLogsError = ""
                    Status = $forcedStatus
                    LogPath = $script:LauncherLogPath
                    ErrorMessage = if ($forcedEvidenceFound) { $forcedDetail } else { ($jobErrors -join " | ") }
                }
            }

            foreach ($item in @($received)) {
                if ($null -ne $item) {
                    if ($jobErrors.Count -gt 0 -and -not $item.PSObject.Properties["JobErrorMessage"]) {
                        $item | Add-Member -NotePropertyName JobErrorMessage -NotePropertyValue ($jobErrors -join " | ") -Force
                    }
                    $item | Add-Member -NotePropertyName EffectiveStatus -NotePropertyValue ([string]$item.Status) -Force
                    $item | Add-Member -NotePropertyName EffectiveNextAction -NotePropertyValue ([string]$item.NextAction) -Force
                    if ($script:UseEffectiveTechnicianRunGuardHistory) {
                        $jobKey = [string]$job.Id
                        $resultFqdn = if ($techRunGuardFqdnByJobId.ContainsKey($jobKey)) { [string]$techRunGuardFqdnByJobId[$jobKey] } else { Get-TechnicianRunGuardFqdn -ComputerName ([string]$item.Computer) -AdInventoryMap $script:AdInventoryMap }
                        try {
                            Update-TechnicianRunGuardHistory -Path $script:TechnicianRunGuardHistoryPath -ComputerFqdn $resultFqdn -InputComputerName ([string]$item.Computer) -Hours $TechnicianRunGuardHours -State Result -Result $item -JobId $jobKey -CycleNumber $CycleNumber
                        }
                        catch {
                            $historyWarning = "TECH_RUN_GUARD_HISTORY_WRITE_DEFERRED: State=Result; Computer=$($item.Computer); Error=$($_.Exception.Message)"
                            Write-Host $historyWarning -ForegroundColor Yellow
                            if ($item.PSObject.Properties['ErrorMessage']) {
                                $existingError = [string]$item.ErrorMessage
                                $item.ErrorMessage = if ([string]::IsNullOrWhiteSpace($existingError)) { $historyWarning } else { "$existingError | $historyWarning" }
                            }
                            else {
                                $item | Add-Member -NotePropertyName ErrorMessage -NotePropertyValue $historyWarning -Force
                            }
                        }
                    }
                    $summary.Add($item)
                    try {
                        Add-LiveCycleReportRow -Path $liveSummaryPath -Columns $reportColumns -Row $item
                    }
                    catch {
                        Write-Host ("Cycle {0}: failed to append live report row for {1}: {2}" -f $CycleNumber,$item.Computer,$_.Exception.Message) -ForegroundColor Yellow
                    }
                    if (-not $DryRun -and (Test-AlreadyEnrolledCycleResult -Result $item)) {
                        try {
                            $moveSingleResult = Move-AlreadyEnrolledComputersFromList -ComputerListPath $ComputerListPath -CycleSummary @($item)
                            if ($moveSingleResult.Moved -gt 0) {
                                Write-Host ("Moved already-enrolled computer from Computers.txt to {0}: {1}" -f $moveSingleResult.AlreadyEnrolledPath,$item.Computer) -ForegroundColor Green
                            }
                        }
                        catch {
                            Write-Host ("Cycle {0}: failed to move already-enrolled computer {1}: {2}" -f $CycleNumber,$item.Computer,$_.Exception.Message) -ForegroundColor Yellow
                        }
                    }
                    $completed++
                    $messageSuffix = ""
                    if ($item.PSObject.Properties["RemoteDetail"] -and -not [string]::IsNullOrWhiteSpace([string]$item.RemoteDetail)) {
                        $messageSuffix = " - $($item.RemoteDetail)"
                    }
                    elseif ($item.PSObject.Properties["ErrorMessage"] -and -not [string]::IsNullOrWhiteSpace([string]$item.ErrorMessage)) {
                        $messageSuffix = " - $($item.ErrorMessage)"
                    }
                    elseif ($item.PSObject.Properties["JobErrorMessage"] -and -not [string]::IsNullOrWhiteSpace([string]$item.JobErrorMessage)) {
                        $messageSuffix = " - $($item.JobErrorMessage)"
                    }

                    $actionText = ""
                    if ($item.PSObject.Properties["NextAction"] -and -not [string]::IsNullOrWhiteSpace([string]$item.NextAction)) {
                        $actionText = " | NextAction=$($item.NextAction)"
                    }

                    if ($item.Status -eq "JOB_ERROR" -or $item.Status -eq "ERROR") {
                        Write-Host ("Completed {0}/{1}: {2} => {3}{4}{5}" -f $completed,$computers.Count,$item.Computer,$item.Status,$actionText,$messageSuffix) -ForegroundColor Red
                    }
                    else {
                        Write-Host ("Completed {0}/{1}: {2} => {3}{4}{5}" -f $completed,$computers.Count,$item.Computer,$item.Status,$actionText,$messageSuffix)
                    }
                }
            }

            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            $jobIdKey = [string]$job.Id
            if ($techRunGuardFqdnByJobId.ContainsKey($jobIdKey)) {
                $techRunGuardFqdnByJobId.Remove($jobIdKey)
            }
            if ($jobStartedAtById.ContainsKey($jobIdKey)) {
                $jobStartedAtById.Remove($jobIdKey)
            }
            if ($globalLeaseByJobId.ContainsKey($jobIdKey)) {
                Release-GlobalWorkerLease -LeasePath $globalLeaseByJobId[$jobIdKey]
                $globalLeaseByJobId.Remove($jobIdKey)
            }
        }

        $finishedIds = @($finishedJobs | Select-Object -ExpandProperty Id)
        $runningJobs = @($runningJobs | Where-Object { $finishedIds -notcontains $_.Id })
        if (((Get-Date) - $lastLiveHtmlWrite).TotalSeconds -ge 60) {
            try {
                $liveRows = @($summary | ForEach-Object { $_ })
                $cycleProgress = New-HybridJoinCycleProgressRows -CycleNumber $CycleNumber -CycleStart $cycleStart -TotalComputers $computers.Count -QueuedComputers $nextIndex -CompletedComputers $completed -RunningComputers $runningJobs.Count -ComputerListStats $computerListStats
                $runningJobRows = New-HybridJoinRunningJobRows -RunningJobs @($runningJobs) -JobStartedAtById $jobStartedAtById
                $mergedLiveRows = @($script:AllCycleResults.ToArray()) + $liveRows
                $mergedProgressRows = @($script:AllCycleProgressRows.ToArray()) + @($cycleProgress)
                New-CycleHtmlReport -Summary $mergedLiveRows -Path $script:MergedHtmlReportPath -CycleNumber $CycleNumber -GeneratedAt (Get-Date) -IsLive -CycleProgress $mergedProgressRows -RunningJobRows $runningJobRows
                $lastLiveHtmlWrite = Get-Date
            }
            catch {
                Write-Host ("Cycle {0}: failed to update merged HTML report: {1}" -f $CycleNumber,$_.Exception.Message) -ForegroundColor Yellow
            }
        }
    }

    $summaryRowsForPostCycle = @($summary | ForEach-Object { $_ })
    $cloudRefreshRows = @(Get-PostCycleCloudRefreshRows -Rows $summaryRowsForPostCycle)
    $cycleCouldChangeCloudInventory = $cloudRefreshRows.Count -gt 0
    if (-not $cycleCouldChangeCloudInventory) {
        Write-Host ("Cycle {0}: post-cycle Graph refresh skipped because no result could have changed Intune or Entra inventory." -f $CycleNumber) -ForegroundColor DarkGray
    }
    $postCycleScopePath = Join-Path $ReportRoot ("Devices_Cycle{0}_InventoryScope.txt" -f $CycleNumber)
    $postCycleCloudScopePath = Join-Path $ReportRoot ("Devices_Cycle{0}_CloudRefreshScope.txt" -f $CycleNumber)
    $cloudRefreshKeys = @{}
    foreach ($cloudRefreshRow in $cloudRefreshRows) {
        $cloudRefreshKey = Get-ComputerListKey -ComputerName $cloudRefreshRow.Computer
        $cloudRefreshKeys[$cloudRefreshKey] = $true
    }
    if (-not $DryRun) {
        $computers | Set-Content -LiteralPath $postCycleScopePath -Encoding ASCII
        @($cloudRefreshRows | ForEach-Object { [string]$_.Computer } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique) | Set-Content -LiteralPath $postCycleCloudScopePath -Encoding ASCII
    }

    if (-not $DryRun -and -not $SkipPostCycleIntuneInventory -and $cycleCouldChangeCloudInventory -and -not (Get-LotCancellationState).Requested) {
        Write-Host ("Cycle {0}: refreshing LOT-scoped post-cycle Intune inventory..." -f $CycleNumber) -ForegroundColor Cyan
        $postInventoryStamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $postInventoryOutputPath = Join-Path $ReportRoot ("DevicesIntune_Cycle{0}Refresh_{1}.csv" -f $CycleNumber,$postInventoryStamp)
        $postInventoryLogPath = Join-Path $ReportRoot ("DevicesIntune_Cycle{0}Refresh_{1}.log" -f $CycleNumber,$postInventoryStamp)
        $postInventory = Invoke-FullIntuneInventoryExport `
            -ExportScriptPath $ExportIntuneScriptPath `
            -OutputPath $postInventoryOutputPath `
            -LogPath $postInventoryLogPath `
            -PageSize $PostCycleIntuneInventoryPageSize `
            -ComputerListPath $postCycleCloudScopePath

        if ($postInventory.Success) {
            $postSet = $postInventory.InventorySet
            $script:IntuneInventorySet = Merge-ScopedInventoryMap -ExistingMap $script:IntuneInventorySet -RefreshedMap $postSet -ScopedComputers @($cloudRefreshRows | ForEach-Object { [string]$_.Computer })
            $newlyDetected = 0
            foreach ($row in $summaryRowsForPostCycle) {
                $key = Get-ComputerListKey -ComputerName $row.Computer
                if (-not $cloudRefreshKeys.ContainsKey($key)) { continue }
                $postPresent = [bool]($postSet -and $postSet.ContainsKey($key))
                $preKnown = $row.PSObject.Properties["IntuneInventoryPresent"] -and -not [string]::IsNullOrWhiteSpace([string]$row.IntuneInventoryPresent)
                $prePresent = Test-BooleanLikeTrue -Value $row.IntuneInventoryPresent
                $postDetected = $preKnown -and (-not $prePresent) -and $postPresent
                if ($postDetected) { $newlyDetected++ }

                $row | Add-Member -NotePropertyName PostCycleIntuneInventoryChecked -NotePropertyValue $true -Force
                $row | Add-Member -NotePropertyName PostCycleIntuneInventoryPresent -NotePropertyValue $postPresent -Force
                $row | Add-Member -NotePropertyName PostCycleIntuneEnrollmentDetected -NotePropertyValue $postDetected -Force
                $row | Add-Member -NotePropertyName PostCycleIntuneInventoryCsv -NotePropertyValue (ConvertTo-PortableLotPath -Value $postInventory.CsvPath) -Force
                $row | Add-Member -NotePropertyName PostCycleIntuneInventoryError -NotePropertyValue "" -Force
                if ($postPresent) {
                    $row | Add-Member -NotePropertyName EffectiveStatus -NotePropertyValue "ENROLLED_DETECTED_POST_CYCLE" -Force
                    $row | Add-Member -NotePropertyName EffectiveNextAction -NotePropertyValue "NO_ACTION_ALREADY_INTUNE" -Force
                }
                else {
                    $row | Add-Member -NotePropertyName EffectiveStatus -NotePropertyValue ([string]$row.Status) -Force
                    $row | Add-Member -NotePropertyName EffectiveNextAction -NotePropertyValue ([string]$row.NextAction) -Force
                }
            }

            $postPresentCount = @($cloudRefreshRows | Where-Object { $_.PostCycleIntuneInventoryPresent -eq $true }).Count
            Write-Host ("Cycle {0}: post-cycle Intune inventory found {1}/{2}; newly detected this cycle={3}; CSV={4}" -f $CycleNumber,$postPresentCount,$cloudRefreshRows.Count,$newlyDetected,$postInventory.CsvPath) -ForegroundColor Green
        }
        else {
            foreach ($row in $summaryRowsForPostCycle) {
                $key = Get-ComputerListKey -ComputerName $row.Computer
                if (-not $cloudRefreshKeys.ContainsKey($key)) { continue }
                $row | Add-Member -NotePropertyName PostCycleIntuneInventoryChecked -NotePropertyValue $true -Force
                $row | Add-Member -NotePropertyName PostCycleIntuneInventoryPresent -NotePropertyValue "" -Force
                $row | Add-Member -NotePropertyName PostCycleIntuneEnrollmentDetected -NotePropertyValue "" -Force
                $row | Add-Member -NotePropertyName PostCycleIntuneInventoryCsv -NotePropertyValue (ConvertTo-PortableLotPath -Value $postInventory.CsvPath) -Force
                $row | Add-Member -NotePropertyName PostCycleIntuneInventoryError -NotePropertyValue $postInventory.Error -Force
                $row | Add-Member -NotePropertyName EffectiveStatus -NotePropertyValue ([string]$row.Status) -Force
                $row | Add-Member -NotePropertyName EffectiveNextAction -NotePropertyValue ([string]$row.NextAction) -Force
            }
            Write-Host ("Cycle {0}: post-cycle Intune inventory failed: {1}" -f $CycleNumber,$postInventory.Error) -ForegroundColor Yellow
        }

        try {
            Get-PortableReportRows -Rows $summaryRowsForPostCycle | Select-Object $reportColumns | Export-Csv -LiteralPath $liveSummaryPath -NoTypeInformation -Encoding UTF8
        }
        catch {
            Write-Host ("Cycle {0}: failed to rewrite live CSV with post-cycle Intune columns: {1}" -f $CycleNumber,$_.Exception.Message) -ForegroundColor Yellow
        }
    }

    if (-not $DryRun -and -not [string]::IsNullOrWhiteSpace($EntraInventoryCsv) -and $cycleCouldChangeCloudInventory -and -not (Get-LotCancellationState).Requested) {
        Write-Host ("Cycle {0}: refreshing LOT-scoped post-cycle Entra device inventory..." -f $CycleNumber) -ForegroundColor Cyan
        $postEntraInventoryStamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $postEntraInventoryOutputPath = Join-Path $ReportRoot ("DevicesEntra_Cycle{0}Refresh_{1}.csv" -f $CycleNumber,$postEntraInventoryStamp)
        $postEntraInventoryLogPath = Join-Path $ReportRoot ("DevicesEntra_Cycle{0}Refresh_{1}.log" -f $CycleNumber,$postEntraInventoryStamp)
        $postEntraInventory = Invoke-FullEntraInventoryExport `
            -ExportScriptPath $ExportEntraScriptPath `
            -OutputPath $postEntraInventoryOutputPath `
            -LogPath $postEntraInventoryLogPath `
            -PageSize $PostCycleIntuneInventoryPageSize `
            -ComputerListPath $postCycleCloudScopePath

        if ($postEntraInventory.Success) {
            $postEntraMap = $postEntraInventory.InventoryMap
            $script:EntraInventoryMap = Merge-ScopedInventoryMap -ExistingMap $script:EntraInventoryMap -RefreshedMap $postEntraMap -ScopedComputers @($cloudRefreshRows | ForEach-Object { [string]$_.Computer })
            $pendingResolved = 0

            foreach ($row in $summaryRowsForPostCycle) {
                $key = Get-ComputerListKey -ComputerName $row.Computer
                if (-not $cloudRefreshKeys.ContainsKey($key)) { continue }
                $postEntraPresent = [bool]($postEntraMap -and $postEntraMap.ContainsKey($key))
                $postEntraState = ""
                $postAltSecIdCount = ""

                if ($postEntraPresent) {
                    $postEntraRow = $postEntraMap[$key]
                    if ($postEntraRow.PSObject.Properties["EntraRegisteredState"]) { $postEntraState = [string]$postEntraRow.EntraRegisteredState }
                    if ($postEntraRow.PSObject.Properties["AlternativeSecurityIdCount"]) { $postAltSecIdCount = [string]$postEntraRow.AlternativeSecurityIdCount }
                }

                $wasPending = ($row.PSObject.Properties["EntraRegisteredState"] -and [string]$row.EntraRegisteredState -eq "Pending")
                $isPendingNow = ($postEntraState -eq "Pending")
                $resolvedThisCycle = [bool]($wasPending -and $postEntraPresent -and -not $isPendingNow)
                if ($resolvedThisCycle) { $pendingResolved++ }

                $row | Add-Member -NotePropertyName PostCycleEntraInventoryChecked -NotePropertyValue $true -Force
                $row | Add-Member -NotePropertyName PostCycleEntraInventoryPresent -NotePropertyValue $postEntraPresent -Force
                $row | Add-Member -NotePropertyName PostCycleEntraRegisteredState -NotePropertyValue $postEntraState -Force
                $row | Add-Member -NotePropertyName PostCycleEntraAlternativeSecurityIdCount -NotePropertyValue $postAltSecIdCount -Force
                $row | Add-Member -NotePropertyName PostCycleEntraPendingResolved -NotePropertyValue $resolvedThisCycle -Force
                $row | Add-Member -NotePropertyName PostCycleEntraInventoryCsv -NotePropertyValue (ConvertTo-PortableLotPath -Value $postEntraInventory.CsvPath) -Force
                $row | Add-Member -NotePropertyName PostCycleEntraInventoryError -NotePropertyValue "" -Force

                if ($resolvedThisCycle -and [string]$row.EffectiveStatus -ne "ENROLLED_DETECTED_POST_CYCLE") {
                    $row | Add-Member -NotePropertyName EffectiveStatus -NotePropertyValue "ENTRA_PENDING_RESOLVED_POST_CYCLE" -Force
                    $row | Add-Member -NotePropertyName EffectiveNextAction -NotePropertyValue "RECHECK_INTUNE_ENROLLMENT" -Force
                }
            }

            $postPendingCount = @($cloudRefreshRows | Where-Object { $_.PostCycleEntraRegisteredState -eq "Pending" }).Count
            Write-Host ("Cycle {0}: post-cycle Entra inventory pending={1}; pending resolved this cycle={2}; CSV={3}" -f $CycleNumber,$postPendingCount,$pendingResolved,$postEntraInventory.CsvPath) -ForegroundColor Green
        }
        else {
            foreach ($row in $summaryRowsForPostCycle) {
                $key = Get-ComputerListKey -ComputerName $row.Computer
                if (-not $cloudRefreshKeys.ContainsKey($key)) { continue }
                $row | Add-Member -NotePropertyName PostCycleEntraInventoryChecked -NotePropertyValue $true -Force
                $row | Add-Member -NotePropertyName PostCycleEntraInventoryPresent -NotePropertyValue "" -Force
                $row | Add-Member -NotePropertyName PostCycleEntraRegisteredState -NotePropertyValue "" -Force
                $row | Add-Member -NotePropertyName PostCycleEntraAlternativeSecurityIdCount -NotePropertyValue "" -Force
                $row | Add-Member -NotePropertyName PostCycleEntraPendingResolved -NotePropertyValue "" -Force
                $row | Add-Member -NotePropertyName PostCycleEntraInventoryCsv -NotePropertyValue (ConvertTo-PortableLotPath -Value $postEntraInventory.CsvPath) -Force
                $row | Add-Member -NotePropertyName PostCycleEntraInventoryError -NotePropertyValue $postEntraInventory.Error -Force
            }
            Write-Host ("Cycle {0}: post-cycle Entra inventory failed: {1}" -f $CycleNumber,$postEntraInventory.Error) -ForegroundColor Yellow
        }

        try {
            Get-PortableReportRows -Rows $summaryRowsForPostCycle | Select-Object $reportColumns | Export-Csv -LiteralPath $liveSummaryPath -NoTypeInformation -Encoding UTF8
        }
        catch {
            Write-Host ("Cycle {0}: failed to rewrite live CSV with post-cycle Entra columns: {1}" -f $CycleNumber,$_.Exception.Message) -ForegroundColor Yellow
        }
    }

    $postAdInventoryDue = Test-AdInventoryRefreshDue -LastRefreshUtc $script:AdInventoryLastRefreshUtc -FreshnessHours $AdInventoryFreshnessHours
    if (-not $DryRun -and -not $AdInventoryUsesRecentRootCsv -and -not [string]::IsNullOrWhiteSpace($AdInventoryCsv) -and $postAdInventoryDue -and -not (Get-LotCancellationState).Requested) {
        $postAdScope = if ([string]::IsNullOrWhiteSpace($AdDomain)) { "forest" } else { "domain '$AdDomain'" }
        Write-Host ("Cycle {0}: refreshing LOT-scoped post-cycle AD computer inventory. Scope={1}..." -f $CycleNumber,$postAdScope) -ForegroundColor Cyan
        $postAdInventoryStamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $postAdInventoryOutputPath = Join-Path $ReportRoot ("DevicesAD_Cycle{0}Refresh_{1}.csv" -f $CycleNumber,$postAdInventoryStamp)
        $postAdInventoryLogPath = Join-Path $ReportRoot ("DevicesAD_Cycle{0}Refresh_{1}.log" -f $CycleNumber,$postAdInventoryStamp)
        $postAdInventory = Invoke-FullAdInventoryExport `
            -ExportScriptPath $ExportAdScriptPath `
            -OutputPath $postAdInventoryOutputPath `
            -LogPath $postAdInventoryLogPath `
            -Domain $AdDomain `
            -ComputerListPath $postCycleScopePath

        if ($postAdInventory.Success) {
            $postAdMap = $postAdInventory.InventoryMap
            $script:AdInventoryMap = $postAdMap
            $script:AdInventoryLastRefreshUtc = [datetime]::UtcNow

            foreach ($row in $summaryRowsForPostCycle) {
                $key = Get-ComputerListKey -ComputerName $row.Computer
                $postAdPresent = [bool]($postAdMap -and $postAdMap.ContainsKey($key))

                $row | Add-Member -NotePropertyName PostCycleADInventoryChecked -NotePropertyValue $true -Force
                $row | Add-Member -NotePropertyName PostCycleADInventoryPresent -NotePropertyValue $postAdPresent -Force
                $row | Add-Member -NotePropertyName PostCycleADInventoryCsv -NotePropertyValue (ConvertTo-PortableLotPath -Value $postAdInventory.CsvPath) -Force
                $row | Add-Member -NotePropertyName PostCycleADInventoryError -NotePropertyValue "" -Force
            }

            $postAdPresentCount = @($summaryRowsForPostCycle | Where-Object { $_.PostCycleADInventoryPresent -eq $true }).Count
            Write-Host ("Cycle {0}: post-cycle AD inventory found {1}/{2}; CSV={3}" -f $CycleNumber,$postAdPresentCount,$summaryRowsForPostCycle.Count,$postAdInventory.CsvPath) -ForegroundColor Green
        }
        else {
            foreach ($row in $summaryRowsForPostCycle) {
                $row | Add-Member -NotePropertyName PostCycleADInventoryChecked -NotePropertyValue $true -Force
                $row | Add-Member -NotePropertyName PostCycleADInventoryPresent -NotePropertyValue "" -Force
                $row | Add-Member -NotePropertyName PostCycleADInventoryCsv -NotePropertyValue (ConvertTo-PortableLotPath -Value $postAdInventory.CsvPath) -Force
                $row | Add-Member -NotePropertyName PostCycleADInventoryError -NotePropertyValue $postAdInventory.Error -Force
            }
            Write-Host ("Cycle {0}: post-cycle AD inventory failed: {1}" -f $CycleNumber,$postAdInventory.Error) -ForegroundColor Yellow
        }

        try {
            Get-PortableReportRows -Rows $summaryRowsForPostCycle | Select-Object $reportColumns | Export-Csv -LiteralPath $liveSummaryPath -NoTypeInformation -Encoding UTF8
        }
        catch {
            Write-Host ("Cycle {0}: failed to rewrite live CSV with post-cycle AD columns: {1}" -f $CycleNumber,$_.Exception.Message) -ForegroundColor Yellow
        }
    }
    elseif (-not $DryRun -and -not $AdInventoryUsesRecentRootCsv -and -not [string]::IsNullOrWhiteSpace($AdInventoryCsv) -and -not $postAdInventoryDue) {
        $nextAdRefreshUtc = ([datetime]$script:AdInventoryLastRefreshUtc).ToUniversalTime().AddHours($AdInventoryFreshnessHours)
        Write-Host ("Cycle {0}: post-cycle AD inventory reuse; cache remains fresh until {1:u}." -f $CycleNumber,$nextAdRefreshUtc) -ForegroundColor DarkGray
    }


    if (-not $DryRun) {
        try {
            $moveResult = Move-AlreadyEnrolledComputersFromList -ComputerListPath $ComputerListPath -CycleSummary $summaryRowsForPostCycle
            if ($moveResult.Moved -gt 0) {
                Write-Host ("Cycle {0}: moved {1} already-enrolled computer(s) to {2}" -f $CycleNumber,$moveResult.Moved,$moveResult.AlreadyEnrolledPath) -ForegroundColor Green
            }
            else {
                Write-Host ("Cycle {0}: no already-enrolled computer moved from Computers.txt. {1}" -f $CycleNumber,$moveResult.Detail) -ForegroundColor DarkGray
            }
        }
        catch {
            Write-Host ("Cycle {0}: failed to update ComputersAlreadyEnrolled.txt: {1}" -f $CycleNumber,$_.Exception.Message) -ForegroundColor Yellow
        }
    }

    $latestMetadataRows = @($script:AllCycleResults.ToArray()) + @($summaryRowsForPostCycle)
    Set-HybridJoinLatestStatusMetadata -Rows $latestMetadataRows
    Get-PortableReportRows -Rows $summaryRowsForPostCycle | Select-Object $reportColumns | Export-Csv -LiteralPath $liveSummaryPath -NoTypeInformation -Encoding UTF8
    $summaryPath = Join-Path $ReportRoot ("PsExec_IntuneHybridJoinRepair_Summary_cycle{0}_{1}.csv" -f $CycleNumber,(Get-Date -Format "yyyyMMdd_HHmmss"))
    Get-PortableReportRows -Rows $summaryRowsForPostCycle | Select-Object $reportColumns | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding UTF8

    try {
        foreach ($summaryRow in $summaryRowsForPostCycle) { [void]$script:AllCycleResults.Add($summaryRow) }
        $finalProgress = New-HybridJoinCycleProgressRows -CycleNumber $CycleNumber -CycleStart $cycleStart -TotalComputers $computers.Count -QueuedComputers $computers.Count -CompletedComputers $summaryRowsForPostCycle.Count -RunningComputers 0 -ComputerListStats $computerListStats
        [void]$script:AllCycleProgressRows.Add($finalProgress)
        New-CycleHtmlReport -Summary @($script:AllCycleResults.ToArray()) -Path $script:MergedHtmlReportPath -CycleNumber $CycleNumber -GeneratedAt (Get-Date) -IsLive -CycleProgress @($script:AllCycleProgressRows.ToArray()) -RunningJobRows @()
        Write-Host ("Merged HTML report updated through cycle {0}: {1}" -f $CycleNumber,$script:MergedHtmlReportPath) -ForegroundColor Green
    }
    catch {
        Write-Host ("Cycle {0}: failed to update merged HTML report: {1}" -f $CycleNumber,$_.Exception.Message) -ForegroundColor Yellow
    }

    Write-Host ""
    Write-Host ("Cycle {0} status counts:" -f $CycleNumber) -ForegroundColor Cyan
    @($summary | Group-Object -Property Status | Sort-Object Count -Descending | ForEach-Object {
        "  {0,-45} {1,5}" -f $_.Name,$_.Count
    }) | ForEach-Object { Write-Host $_ }

    Write-Host ("Cycle {0} next-action counts:" -f $CycleNumber) -ForegroundColor Cyan
    @($summary | Group-Object -Property NextAction | Sort-Object Count -Descending | ForEach-Object {
        "  {0,-45} {1,5}" -f $_.Name,$_.Count
    }) | ForEach-Object { Write-Host $_ }

    if ($IntuneInventorySet -and $IntuneInventorySet.Count -gt 0) {
        $present = @($summary | Where-Object { $_.IntuneInventoryPresent -eq $true }).Count
        $absent = @($summary | Where-Object { $_.IntuneInventoryPresent -eq $false }).Count
        $notChecked = [math]::Max(0,$summary.Count - $present - $absent)
        Write-Host ("Cycle {0} Intune inventory match: Present={1}; Absent={2}; NotChecked={3}" -f $CycleNumber,$present,$absent,$notChecked) -ForegroundColor Cyan
    }

    if ($AdInventoryMap -and $AdInventoryMap.Count -gt 0) {
        $present = @($summary | Where-Object { $_.ADInventoryPresent -eq $true }).Count
        $absent = @($summary | Where-Object { $_.ADInventoryPresent -eq $false }).Count
        $notChecked = [math]::Max(0,$summary.Count - $present - $absent)
        Write-Host ("Cycle {0} AD inventory match: Present={1}; Absent={2}; NotChecked={3}" -f $CycleNumber,$present,$absent,$notChecked) -ForegroundColor Cyan
    }

    if (@($summaryRowsForPostCycle | Where-Object { $_.PSObject.Properties["PostCycleIntuneInventoryChecked"] -and $_.PostCycleIntuneInventoryChecked -eq $true }).Count -gt 0) {
        $postPresent = @($summaryRowsForPostCycle | Where-Object { $_.PostCycleIntuneInventoryPresent -eq $true }).Count
        $postAbsent = @($summaryRowsForPostCycle | Where-Object { $_.PostCycleIntuneInventoryPresent -eq $false }).Count
        $postNotChecked = [math]::Max(0,$summaryRowsForPostCycle.Count - $postPresent - $postAbsent)
        $postNew = @($summaryRowsForPostCycle | Where-Object { $_.PostCycleIntuneEnrollmentDetected -eq $true }).Count
        Write-Host ("Cycle {0} post-cycle Intune inventory: Present={1}; Absent={2}; NotChecked={3}; NewlyDetected={4}" -f $CycleNumber,$postPresent,$postAbsent,$postNotChecked,$postNew) -ForegroundColor Cyan
    }

    if (@($summaryRowsForPostCycle | Where-Object { $_.PSObject.Properties["PostCycleADInventoryChecked"] -and $_.PostCycleADInventoryChecked -eq $true }).Count -gt 0) {
        $postPresent = @($summaryRowsForPostCycle | Where-Object { $_.PostCycleADInventoryPresent -eq $true }).Count
        $postAbsent = @($summaryRowsForPostCycle | Where-Object { $_.PostCycleADInventoryPresent -eq $false }).Count
        $postNotChecked = [math]::Max(0,$summaryRowsForPostCycle.Count - $postPresent - $postAbsent)
        Write-Host ("Cycle {0} post-cycle AD inventory: Present={1}; Absent={2}; NotChecked={3}" -f $CycleNumber,$postPresent,$postAbsent,$postNotChecked) -ForegroundColor Cyan
    }

    Write-Host ("Cycle {0} merged HTML report: {1}" -f $CycleNumber,$script:MergedHtmlReportPath) -ForegroundColor Green

    $script:LastCycleSummaryRows = @($summaryRowsForPostCycle)
    Write-Host ("Cycle {0} done. Summary: {1}" -f $CycleNumber,$summaryPath) -ForegroundColor Green
    return $summaryPath
}

$script:TechnicianRunGuardHistoryPath = Get-TechnicianRunGuardHistoryPath
$script:UseEffectiveTechnicianRunGuardHistory = ($UseTechnicianRunGuardHistory -and -not $IgnoreTechnicianRunGuardHistory -and -not $IgnoreRunGuard -and -not $DryRun -and $TechnicianRunGuardHours -gt 0)
if ($script:UseEffectiveTechnicianRunGuardHistory) {
    try {
        Invoke-TechnicianRunGuardHistoryLock -ArgumentList @($script:TechnicianRunGuardHistoryPath,$TechnicianRunGuardHours) -ScriptBlock {
            param($LockedPath,$LockedHours)
            $history = Read-TechnicianRunGuardHistory -Path $LockedPath -Hours $LockedHours
            Save-TechnicianRunGuardHistory -Path $LockedPath -History $history
        }
    }
    catch {
        Write-Host ("Technician run guard history is temporarily unavailable; launches will fail closed until access recovers. Error={0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
}
Write-Host ("Technician run guard: Effective={0}; Requested={1}; Ignore={2}; Hours={3}; Path={4}" -f $script:UseEffectiveTechnicianRunGuardHistory,[bool]$UseTechnicianRunGuardHistory,[bool]$IgnoreTechnicianRunGuardHistory,$TechnicianRunGuardHours,$script:TechnicianRunGuardHistoryPath) -ForegroundColor DarkCyan
$mergedLotName = if (-not [string]::IsNullOrWhiteSpace([string]$LotRoot)) { Split-Path -Leaf $LotRoot } else { 'LOT' }
$mergedSafeLotName = [regex]::Replace([string]$mergedLotName, '[^A-Za-z0-9._-]+', '-').Trim('-._')
if ([string]::IsNullOrWhiteSpace($mergedSafeLotName)) { $mergedSafeLotName = 'LOT' }
$mergedHtmlReportTimestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$script:MergedHtmlReportPath = Join-Path $ReportRoot ("PsExec_IntuneHybridJoinRepair_Summary_{0}_{1}.html" -f $mergedSafeLotName,$mergedHtmlReportTimestamp)
$script:AllCycleResults = New-Object System.Collections.ArrayList
$script:AllCycleProgressRows = New-Object System.Collections.ArrayList
$script:LastCycleSummaryRows = @()
Write-Host ("Merged HTML report: {0}" -f $script:MergedHtmlReportPath) -ForegroundColor DarkCyan
Set-ActiveLotRunState -Status 'Running' -ReportPath $script:MergedHtmlReportPath

$cycle = 0
$script:LotStopReason = ''
do {
    $preCycleCancellation = Get-LotCancellationState
    if ($preCycleCancellation.Requested) {
        Write-Host ("Cancellation requested before cycle {0}; no new cycle will start." -f ($cycle + 1)) -ForegroundColor Yellow
        break
    }
    Wait-OutsideNightPauseWindow -NextCycleNumber ($cycle + 1)
    if (-not (Test-ComputerListPresentWithRetry -Path $ComputerListPath)) {
        $script:LotStopReason = 'LOT_INPUT_REMOVED'
        Write-Host ("[LOT_INPUT_REMOVED] Computers.txt disappeared after the LOT started. Existing reports are preserved and no new worker will start. Path={0}" -f $ComputerListPath) -ForegroundColor Yellow
        Set-ActiveLotRunState -Status 'StoppedInputRemoved' -ReportPath $script:MergedHtmlReportPath
        break
    }
    $cycle++
    $cycleArgs = @($scriptArgsBase)
    if ($IgnoreRunGuard -and ($cycle -eq 1 -or $IgnoreRunGuardEveryCycle)) {
        $cycleArgs += "-IgnoreRunGuard"
    }

    try {
        $null = Invoke-IntuneHybridJoinRepairCycle -CycleNumber $cycle -CycleScriptArgs $cycleArgs
    }
    catch {
        if (-not (Test-ComputerListPresentWithRetry -Path $ComputerListPath)) {
            $script:LotStopReason = 'LOT_INPUT_REMOVED'
            Write-Host ("[LOT_INPUT_REMOVED] Computers.txt disappeared during cycle {0}. Existing reports are preserved and the LOT will stop cleanly. Path={1}" -f $cycle,$ComputerListPath) -ForegroundColor Yellow
            Set-ActiveLotRunState -Status 'StoppedInputRemoved' -ReportPath $script:MergedHtmlReportPath
            break
        }
        throw
    }

    if ((Get-LotCancellationState).Requested) { break }
    if ($RunOnce) { break }
    if ($MaxCycles -gt 0 -and $cycle -ge $MaxCycles) { break }

    $minimumCycleDelaySeconds = $DelayBetweenCyclesMinutes * 60
    $cycleDelaySeconds = Get-AdaptiveCycleDelaySeconds -Rows $script:LastCycleSummaryRows -MinimumDelaySeconds $minimumCycleDelaySeconds
    if ($cycleDelaySeconds -gt 0) {
        if ($cycleDelaySeconds -gt $minimumCycleDelaySeconds) {
            $resumeAt = (Get-Date).AddSeconds($cycleDelaySeconds)
            Write-Host ("All devices remain under technician backoff. Waiting until {0:yyyy-MM-dd HH:mm:ss} ({1:N1} minute(s)) before the next cycle. Press Ctrl+C to stop." -f $resumeAt,($cycleDelaySeconds / 60.0)) -ForegroundColor DarkGray
        }
        else {
            Write-Host ("Waiting {0:N1} minute(s) before next cycle. Press Ctrl+C to stop." -f ($cycleDelaySeconds / 60.0)) -ForegroundColor DarkGray
        }
        [void](Wait-LotCancellationAware -Seconds $cycleDelaySeconds)
    }
} while ($true)

try {
    if ($script:AllCycleResults.Count -gt 0) {
        New-CycleHtmlReport -Summary @($script:AllCycleResults.ToArray()) -Path $script:MergedHtmlReportPath -CycleNumber $cycle -GeneratedAt (Get-Date) -CycleProgress @($script:AllCycleProgressRows.ToArray()) -RunningJobRows @()
        Write-Host ("Merged final HTML report: {0}" -f $script:MergedHtmlReportPath) -ForegroundColor Green
    }
}
catch {
    Write-Host ("Failed to write merged final HTML report: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
}

Write-Host ""
Write-Host "Launcher stopped after $cycle cycle(s)." -ForegroundColor Green

if ($script:globalConcurrencyMutex -ne $null) {
    try { $script:globalConcurrencyMutex.Dispose() } catch { }
    $script:globalConcurrencyMutex = $null
}

if ($script:LotRunMutex -ne $null) {
    try { $script:LotRunMutex.ReleaseMutex() } catch { }
    try { $script:LotRunMutex.Dispose() } catch { }
    $script:LotRunMutex = $null
}

$finalLotRunStatus = if ($script:LotStopReason -eq 'LOT_INPUT_REMOVED') { 'StoppedInputRemoved' } else { 'Finished' }
Set-ActiveLotRunState -Status $finalLotRunStatus -ReportPath $script:MergedHtmlReportPath
Complete-LotCancellationSupport

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCE0Yg0xS8W/YC+
# U29ylY/LQv6uWsMWXNQMTJf9qbwDIaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEU8mxoJr+lYkbpreGX1EUvH6KteoGH638VfZQaucbsQMA0GCSqG
# SIb3DQEBAQUABIIBgCsSrbSjFoiglBMPY+unbOgOkrRB53y5sk+Agv4hLlp1xFYH
# dc+YuJEg3KI1+9N2fIGixStUYDxkY/a+T0xA/qHYand9Qcc28fMOwX8sxSrgiSLj
# Qiicfr5RxXcrnPOOmkNRe+LZY1IdQWk56qyck7eu40W0uYJuX7wcUl52A1RPtaJ6
# WOjWLg6HynskONrtfwscz9mdWq99ZXhimgk2gRyCfXj7gms8CMR1P7BOphkENzkO
# yBR/KtVYksAxCG8NKjI8NG1KNw42Bs+dtmEo1LKFag1W6Y4v8cM34S05p4qRHYGj
# Wp/f74wTCMRtwMX/CGt/y4y9KdrBWNApVQh6uFQRQM/gFXqP0EWANPdUr5gMMP0u
# OHOi2l9Yu6pmzVY5IErvup/96W5bpfuBlsujfW/TyqqXzHSYUcnjLPFRUuQ2CNGm
# MbeCedhwn4693oIOuc4jeVzQFtUJP7W7KXiUL4pp88mgjbKIYZMHc7JHoz76J7en
# NNhf0V70Sc3kld+YOqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA4MTIxMDM5
# MDNaMC8GCSqGSIb3DQEJBDEiBCB74MYATCU1wRWx1Y5PMsBq45yXCWdgo+2iP0c6
# g+7zdTANBgkqhkiG9w0BAQEFAASCAgA39Gi6/IWm55hP6Psm3E88KdlJEzh5tKCf
# v2pt5WErOTVkFcx8V4Ts+zWhWwdrEr9QHOMyb0VSZIXkT2OKias7e7xVPuJpgV7u
# cRzHIdAxA3kHONLp5wMM9Wuy3BzQ8JC9/JLtDNXvGffXy1dhOdFtathkdmRnji/R
# sQ3up2OUF9uUaylxWc7dlIXD+h8T+JObThm0TAaxq0LAV44Z3sb9Kv+AoLsbt93x
# RAdR81f9DNy/7ydf+6DTsY1zhAWD1X/ixlbAseU/k8fv1Z8gYueBsz4APzWrcCwb
# bxPHX5DGoGrifENyUCJC+rMS83OtS3HMoL3mF/rwsT1ouWQNEBvua7l7V5xIICzs
# Nqkz1VpsFyARexCHZQPARPNRN8nNrxtbpIfTpMV9Qz64ENT3aClgbL5xd3AQurD9
# cPEmzW1HTrsMAShYAsMgA32giREFzilaY533A5p+xu3/GjNOdTo4nyZgstYSWaYo
# wjaiWmcgrsy1GIH9I/lXRcSilYxYSuYahlgg625PgyYkdcXd8bVzE2aZ98jnM/wI
# LeWKmObDXU4W79+gKK64JtPtAd5X6J70ReXZHSqQ8itjHtV/nhX3dschLkyNaciX
# 80uEfxuKJZls1Az+K+9dcHj0elIb2bSmNiqnXtIJLnc8yTwnnuXy9zMEXGVwyIT1
# wqgbuvzodg==
# SIG # End signature block
