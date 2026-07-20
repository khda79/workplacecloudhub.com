<#
.SYNOPSIS
SYSTEM-safe Entra Hybrid Join and Intune enrollment repair helper (dsregcmd /status, optional /leave, trigger Automatic-Device-Join, optional MDM auto-enrollment, bounded local retry).

.DESCRIPTION
Script name: SmartM365-Invoke-IntuneHybridJoinRepair.ps1

Runs locally on the machine (typically under LocalSystem via agent). Writes under:
C:\ProgramData\SmartM365\IntuneHybridJoinToolkit

- If not domain joined => Exit 2.
- If DC not reachable => Exit 4.
- Run dsregcmd /status as SYSTEM and parse machine/device fields only. User PRT shown in SYSTEM context is intentionally ignored.
- Create and run SmartM365 user-context scheduled tasks to capture and refresh the actual interactive user's PRT when available.
- If AzureAdJoined=YES and DeviceAuthStatus contains SUCCESS, or older dsreg output has KeySignTest=PASSED without DeviceAuthStatus, verify Intune enrollment.
- If Hybrid Join is healthy but Intune is not enrolled => verify MDM auto-enrollment GPO/registry policy and HTTPS connectivity, then trigger auto-enrollment using the configured credential type.
- If the launcher reports that the Entra hybrid object is pending, and the local device join is healthy => trigger Automatic-Device-Join and perform a bounded retry without dsregcmd /leave.
- If auto-enrollment policy uses User Credential => trigger a user-context task first, then the native EnterpriseMgmt task if available.
- If auto-enrollment policy uses Device Credential => run deviceenroller.exe /c /AutoEnrollMDM from SYSTEM.
- If user-context status is available and AzureAdPrt remains not YES after refreshprt => Exit 3.
- If Hybrid Join is not healthy and Intune is enrolled => Skip dsregcmd repair actions => Exit 3.
- If AllowDsregLeave AND DeviceAuthStatus starts with FAILED => run /leave, trigger Automatic-Device-Join,
  then wait 20 minutes and re-check up to 2 times. Exit 0 if success, else Exit 3.
- If AzureAdJoined=NO and diagnostic indicates error_missing_device => wait 20 minutes and re-check up to 2 times. Exit 0 if success, else Exit 3.

.PARAMETER AllowDsregLeave
Allows the script to run dsregcmd /leave when DeviceAuthStatus starts with FAILED.

.PARAMETER IgnoreRunGuard
Bypasses the local 12-hour run guard for troubleshooting or controlled re-runs.

.PARAMETER AllowRebootWhenNoInteractiveUser
Allows a reboot when Hybrid Join is healthy, Intune enrollment is missing, auto-enrollment requires User Credential, and no active interactive user is connected.

.PARAMETER AllowRebootAfterDsregLeave
Allows a reboot after a successful dsregcmd /leave when rejoin did not complete during the local retry window.

.PARAMETER AllowRemoveNonIntuneMdmEnrollment
Allows removal of existing non-Intune MDM enrollment registry keys and EnterpriseMgmt scheduled tasks before trying Intune auto-enrollment.

.PARAMETER AllowRemoveStaleIntuneEnrollment
Allows removal of stale local Intune enrollment traces when Windows reports an Intune discovery URL but no confirmed Intune ProviderID.

.PARAMETER SkipVirtualMachines
Skips detected virtual machines before DNS, domain, gpupdate, or repair actions.

.PARAMETER AuditOnly
Runs diagnostics and writes logs/CSV without leave, removal, auto-enrollment triggers, reboot, or retry actions.

.PARAMETER EntraHybridPending
Indicates that the admin-side Entra inventory found this device as a pending hybrid object (TrustType=ServerAd with empty AlternativeSecurityIds).

.PARAMETER StaleCleanupDelaySeconds
Seconds to wait after successful stale local Intune enrollment cleanup before attempting auto-enrollment in the same run. Defaults to 60.

.PARAMETER RebootDelaySeconds
Seconds used for controlled reboots. Defaults to 180 so the admin launcher can pull logs through ADMIN$ before the reboot occurs.

.PARAMETER IntuneRetrySleepMinutes
Minutes to wait between local Intune enrollment re-checks after auto-enrollment is triggered. Defaults to 5.

.PARAMETER IntuneRetryMaxRetries
Number of local Intune enrollment re-checks after auto-enrollment is triggered. Defaults to 5.

.PARAMETER RetryAfterRebootDelaySeconds
Seconds to wait after startup before the SYSTEM retry task resumes the repair. Defaults to 300.

.PARAMETER RetryAfterRebootMaxAttempts
Maximum number of startup-task resume attempts before the retry state is removed. Defaults to 3.

.VERSION
2.10.37

.EXITCODES
0 = Success (AzureAdJoined=YES, device auth is healthy, and Intune enrollment is present or was restored)
3 = Skipped / No-op / Waiting / Retry exhausted (Intune enrolled while Hybrid Join unhealthy, leave not applicable, join/enrollment not completed yet)
2 = Not domain-joined (stopped at start)
4 = Domain controller not reachable (stopped at start)
1 = Error / Failed
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [switch]$AllowDsregLeave,
    [switch]$IgnoreRunGuard,
    [switch]$AllowRebootWhenNoInteractiveUser,
    [switch]$AllowRebootAfterDsregLeave,
    [switch]$AllowRemoveNonIntuneMdmEnrollment,
    [switch]$AllowRemoveStaleIntuneEnrollment,
    [switch]$SkipVirtualMachines,
    [switch]$AuditOnly,
    [switch]$EntraHybridPending,
    [int]$StaleCleanupDelaySeconds = 60,
    [int]$RebootDelaySeconds = 180,
    [int]$IntuneRetrySleepMinutes = 5,
    [int]$IntuneRetryMaxRetries = 5,
    [ValidateRange(0,3600)][int]$RetryAfterRebootDelaySeconds = 300,
    [ValidateRange(1,30)][int]$RetryAfterRebootMaxAttempts = 3,
    [switch]$RetryAfterRebootTaskRun
)

$ScriptVersion = "2.10.37"
if ($RebootDelaySeconds -lt 60) { $RebootDelaySeconds = 60 }
if ($StaleCleanupDelaySeconds -lt 0) { $StaleCleanupDelaySeconds = 0 }
if ($IntuneRetrySleepMinutes -lt 1) { $IntuneRetrySleepMinutes = 1 }
if ($IntuneRetryMaxRetries -lt 1) { $IntuneRetryMaxRetries = 1 }

# ============================
# Enforce 64-bit PowerShell (anti-WOW64)
# ============================
try {
    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
        if (-not $env:EHJIR_Relaunched64) {
            $env:EHJIR_Relaunched64 = "1"

            $sysNativePS = Join-Path $env:WINDIR "SysNative\WindowsPowerShell\v1.0\powershell.exe"
            if (-not (Test-Path $sysNativePS)) {
                $sysNativePS = Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"
            }

            $argList = @(
                "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-File", "`"$($MyInvocation.MyCommand.Path)`""
            )
            if ($AllowDsregLeave) { $argList += "-AllowDsregLeave" }
            if ($IgnoreRunGuard) { $argList += "-IgnoreRunGuard" }
            if ($AllowRebootWhenNoInteractiveUser) { $argList += "-AllowRebootWhenNoInteractiveUser" }
            if ($AllowRebootAfterDsregLeave) { $argList += "-AllowRebootAfterDsregLeave" }
            if ($AllowRemoveNonIntuneMdmEnrollment) { $argList += "-AllowRemoveNonIntuneMdmEnrollment" }
            if ($AllowRemoveStaleIntuneEnrollment) { $argList += "-AllowRemoveStaleIntuneEnrollment" }
            if ($SkipVirtualMachines) { $argList += "-SkipVirtualMachines" }
            if ($AuditOnly) { $argList += "-AuditOnly" }
            if ($EntraHybridPending) { $argList += "-EntraHybridPending" }
            $argList += "-StaleCleanupDelaySeconds"
            $argList += $StaleCleanupDelaySeconds
            $argList += "-RebootDelaySeconds"
            $argList += $RebootDelaySeconds
            $argList += "-IntuneRetrySleepMinutes"
            $argList += $IntuneRetrySleepMinutes
            $argList += "-IntuneRetryMaxRetries"
            $argList += $IntuneRetryMaxRetries
            $argList += "-RetryAfterRebootDelaySeconds"
            $argList += $RetryAfterRebootDelaySeconds
            $argList += "-RetryAfterRebootMaxAttempts"
            $argList += $RetryAfterRebootMaxAttempts
            if ($RetryAfterRebootTaskRun) { $argList += "-RetryAfterRebootTaskRun" }

            $p = Start-Process -FilePath $sysNativePS -ArgumentList $argList -Wait -PassThru
            exit $p.ExitCode
        }
    }
}
catch {
    # Best effort; continue
}

# ============================
# ProgramData root
# ============================
$DataRoot      = "C:\ProgramData\SmartM365\IntuneHybridJoinToolkit"
$OutputDir     = Join-Path $DataRoot "Output"
$TranscriptDir = Join-Path $DataRoot "Transcripts"
$LogsDir       = Join-Path $DataRoot "Logs"
$StateDir      = Join-Path $DataRoot "State"

foreach ($d in @($DataRoot, $OutputDir, $TranscriptDir, $LogsDir, $StateDir)) {
    if (-not (Test-Path $d)) {
        New-Item -ItemType Directory -Path $d -Force | Out-Null
    }
}

# Run metadata
$RunId        = (Get-Date -Format "yyyyMMdd_HHmmss_fff")
$Timestamp    = Get-Date
$ComputerName = $env:COMPUTERNAME
$RunGuardHours = 12
$CleanupRetentionDays = 7
$RunGuardPath = Join-Path $DataRoot "LastRun.json"
$script:RetryAfterRebootTaskName = "SmartM365-IntuneHybridJoinToolkit-RetryAfterReboot"
$script:RetryAfterRebootStatePath = Join-Path $StateDir "RetryAfterReboot.json"
$script:RetryAfterRebootRunnerPath = Join-Path $StateDir "RetryAfterRebootRunner.ps1"
$script:RebootSafetyStatePath = Join-Path $StateDir "RebootSafety.json"
$script:EndpointInstanceStatePath = Join-Path $StateDir "EndpointInstance.json"
$script:EndpointInstanceMutexName = "Global\SmartM365_IntuneHybridJoinToolkit_Endpoint"
$script:EndpointInstanceMutex = $null
$script:EndpointInstanceMutexAcquired = $false
$script:EndpointInstanceLastHeartbeatUtc = [datetime]::MinValue
$script:EndpointScriptPath = if (-not [string]::IsNullOrWhiteSpace([string]$PSCommandPath)) { [string]$PSCommandPath } else { [string]$MyInvocation.MyCommand.Path }
$script:RetryAfterRebootAction = ""
$script:RetryAfterRebootDetail = ""
$script:RetryAfterRebootAttempt = ""
$script:RetryAfterRebootMaxAttemptsResult = ""
$script:RetryAfterRebootTaskNameResult = ""

# Run log file
$RunLogPath = Join-Path $LogsDir ("IntuneHybridJoinToolkit_{0}_{1}.log" -f $ComputerName, $RunId)


function Write-AtomicJsonFile {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)]$Data,
        [ValidateRange(1,20)][int]$Depth = 8
    )

    $parent = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) {
        New-Item -ItemType Directory -Path $parent -Force | Out-Null
    }
    $temporaryPath = Join-Path $parent (".{0}.{1}.tmp" -f (Split-Path -Leaf $Path),[guid]::NewGuid().ToString("N"))
    try {
        $Data | ConvertTo-Json -Depth $Depth | Set-Content -LiteralPath $temporaryPath -Encoding UTF8 -Force -ErrorAction Stop
        Move-Item -LiteralPath $temporaryPath -Destination $Path -Force -ErrorAction Stop
    }
    finally {
        Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
    }
}

function Update-EndpointInstanceState {
    param([string]$Status = "RUNNING",[switch]$Force)

    if (-not $script:EndpointInstanceMutexAcquired) { return }
    $nowUtc = (Get-Date).ToUniversalTime()
    if (-not $Force -and ($nowUtc - $script:EndpointInstanceLastHeartbeatUtc).TotalSeconds -lt 60) { return }
    $script:EndpointInstanceLastHeartbeatUtc = $nowUtc
    Write-AtomicJsonFile -Path $script:EndpointInstanceStatePath -Data ([pscustomobject]@{
        Version = 1
        ScriptVersion = $ScriptVersion
        RunId = $RunId
        ComputerName = $ComputerName
        ProcessId = $PID
        StartTimeUtc = $Timestamp.ToUniversalTime().ToString("o")
        HeartbeatUtc = $nowUtc.ToString("o")
        Status = $Status
        ScriptPath = $script:EndpointScriptPath
    })
}

function Write-RunLog {
    param([Parameter(Mandatory=$true)][string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $line = @([regex]::Split(([string]$Message), '\r?\n') | ForEach-Object { "{0} [{1}] {2}" -f $timestamp, $ComputerName, $_ })
    try { Add-Content -Path $RunLogPath -Value $line -Encoding UTF8 } catch { }
    try { Update-EndpointInstanceState } catch { }
}

function Update-TimestampedTranscript {
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

function Write-FinalStatusLine {
    param(
        [Parameter(Mandatory=$true)][string]$Status,
        [Parameter(Mandatory=$true)][int]$ExitCode,
        [Parameter(Mandatory=$false)][string]$Detail = "",
        [Parameter(Mandatory=$false)][string]$NextAction = ""
    )

    Write-Host ("FINAL_STATUS={0}; EXIT_CODE={1}; NEXT_ACTION={2}; DETAIL={3}" -f $Status,$ExitCode,$NextAction,$Detail)
}

function ConvertTo-CleanNativeOutput {
    param([AllowNull()][object[]]$Output)

    $lines = @()
    foreach ($item in @($Output)) {
        if ($null -eq $item) { continue }
        $text = [string]$item
        foreach ($line in ($text -split "(`r`n|`n|`r)")) {
            $trimmed = $line.TrimEnd()
            if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
            if ($trimmed -match "^\s*At .+:\d+ char:\d+") { continue }
            if ($trimmed -match "^\s*\+\s") { continue }
            if ($trimmed -match "^\s*CategoryInfo\s*:") { continue }
            if ($trimmed -match "^\s*FullyQualifiedErrorId\s*:") { continue }
            $lines += $trimmed
        }
    }

    return ($lines -join " ").Trim()
}

function Convert-DsregDateToUtc {
    param([Parameter(Mandatory=$false)][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }

    $clean = $Value.Trim()
    $clean = $clean -replace "\s+UTC$","Z"
    $clean = $clean -replace "\s+GMT$","Z"

    $styles = [System.Globalization.DateTimeStyles]::AssumeUniversal -bor [System.Globalization.DateTimeStyles]::AdjustToUniversal
    $cultures = @(
        [System.Globalization.CultureInfo]::InvariantCulture,
        [System.Globalization.CultureInfo]::CurrentCulture,
        (New-Object System.Globalization.CultureInfo("en-US")),
        (New-Object System.Globalization.CultureInfo("fr-FR"))
    )

    foreach ($culture in $cultures) {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse($clean, $culture, $styles, [ref]$parsed)) {
            return $parsed.ToUniversalTime()
        }
    }

    return $null
}

function Test-UserPrtRefreshNeeded {
    param(
        [Parameter(Mandatory=$true)][object]$ParsedUserDsreg,
        [Parameter(Mandatory=$false)][int]$ExpiryRefreshWindowHours = 4
    )

    $reason = ""

    if ($ParsedUserDsreg.AzureAdPrt -ne "YES") {
        $reason = ("AzureAdPrt is '{0}'." -f $ParsedUserDsreg.AzureAdPrt)
        return [PSCustomObject]@{ Needed=$true; Reason=$reason }
    }

    $expiryUtc = Convert-DsregDateToUtc -Value $ParsedUserDsreg.AzureAdPrtExpiryTime
    if ($null -ne $expiryUtc) {
        $refreshThresholdUtc = (Get-Date).ToUniversalTime().AddHours($ExpiryRefreshWindowHours)
        if ($expiryUtc -le $refreshThresholdUtc) {
            $reason = ("AzureAdPrtExpiryTime is {0:u}, within {1} hour(s)." -f $expiryUtc,$ExpiryRefreshWindowHours)
            return [PSCustomObject]@{ Needed=$true; Reason=$reason }
        }
    }

    $attemptStatus = [string]$ParsedUserDsreg.RefreshPrtAttemptStatus
    if (-not [string]::IsNullOrWhiteSpace($attemptStatus) -and $attemptStatus -notmatch "^(0x)?0+$") {
        $reason = ("RefreshPrtDiagnostics AttemptStatus is {0}." -f $attemptStatus)
        return [PSCustomObject]@{ Needed=$true; Reason=$reason }
    }

    $httpError = [string]$ParsedUserDsreg.RefreshPrtHttpError
    if (-not [string]::IsNullOrWhiteSpace($httpError) -and $httpError -notmatch "^(0x)?0+$") {
        $reason = ("RefreshPrtDiagnostics HttpError is {0}." -f $httpError)
        return [PSCustomObject]@{ Needed=$true; Reason=$reason }
    }

    $httpStatus = [string]$ParsedUserDsreg.RefreshPrtHttpStatus
    if (-not [string]::IsNullOrWhiteSpace($httpStatus) -and $httpStatus -ne "200") {
        $reason = ("RefreshPrtDiagnostics HttpStatus is {0}." -f $httpStatus)
        return [PSCustomObject]@{ Needed=$true; Reason=$reason }
    }

    return [PSCustomObject]@{ Needed=$false; Reason="AzureAdPrt is present and no stale/failed refresh diagnostic was detected." }
}

function Invoke-OldEvidenceCleanup {
    param(
        [Parameter(Mandatory=$true)][string[]]$Paths,
        [Parameter(Mandatory=$true)][int]$RetentionDays
    )

    $cutoff = (Get-Date).AddDays(-1 * $RetentionDays)
    $deletedCount = 0
    $failedCount = 0

    foreach ($path in $Paths) {
        if (-not (Test-Path -LiteralPath $path)) {
            continue
        }

        $oldFiles = @(
            Get-ChildItem -LiteralPath $path -File -Recurse -Force -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt $cutoff }
        )

        foreach ($file in $oldFiles) {
            try {
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                $deletedCount++
            }
            catch {
                $failedCount++
                Write-RunLog ("Old evidence cleanup failed for file '{0}': {1}" -f $file.FullName,$_.Exception.Message)
            }
        }
    }

    Write-RunLog ("Old evidence cleanup completed. RetentionDays={0}; DeletedFiles={1}; FailedFiles={2}" -f $RetentionDays,$deletedCount,$failedCount)
}

try {
    $script:EndpointInstanceMutex = New-Object System.Threading.Mutex($false,$script:EndpointInstanceMutexName)
    try { $script:EndpointInstanceMutexAcquired = $script:EndpointInstanceMutex.WaitOne(0) }
    catch [System.Threading.AbandonedMutexException] { $script:EndpointInstanceMutexAcquired = $true }
}
catch { $script:EndpointInstanceMutexAcquired = $false }

if (-not $script:EndpointInstanceMutexAcquired) {
    $activeMessage = "Another endpoint repair instance is already active. Mutex=$($script:EndpointInstanceMutexName). IgnoreRunGuard cannot bypass an active endpoint instance."
    Write-RunLog $activeMessage
    Write-Host $activeMessage -ForegroundColor Yellow
    Write-FinalStatusLine -Status "ENDPOINT_RUN_ACTIVE" -ExitCode 3 -Detail $activeMessage -NextAction "WAIT_ACTIVE_ENDPOINT_RUN"
    if ($script:EndpointInstanceMutex) { try { $script:EndpointInstanceMutex.Dispose() } catch { } }
    exit 3
}

try { Update-EndpointInstanceState -Force } catch { Write-RunLog ("Endpoint instance metadata write failed; mutex protection remains active. Error={0}" -f $_.Exception.Message) }

Write-Host "IntuneHybridJoinToolkit version $ScriptVersion"
Write-RunLog "Script start. Version=$ScriptVersion. RunId=$RunId. AllowDsregLeave=$([bool]$AllowDsregLeave). AllowRemoveNonIntuneMdmEnrollment=$([bool]$AllowRemoveNonIntuneMdmEnrollment). AllowRemoveStaleIntuneEnrollment=$([bool]$AllowRemoveStaleIntuneEnrollment). SkipVirtualMachines=$([bool]$SkipVirtualMachines). AuditOnly=$([bool]$AuditOnly). EntraHybridPending=$([bool]$EntraHybridPending). IgnoreRunGuard=$([bool]$IgnoreRunGuard). RetryAfterRebootTaskRun=$([bool]$RetryAfterRebootTaskRun)."
Invoke-OldEvidenceCleanup -Paths @($LogsDir, $OutputDir, $TranscriptDir) -RetentionDays $CleanupRetentionDays

# ============================
# 12-hour local run guard
# ============================
try {
    if ($IgnoreRunGuard -or $RetryAfterRebootTaskRun) {
        $guardBypassReason = if ($RetryAfterRebootTaskRun) { "-RetryAfterRebootTaskRun" } else { "-IgnoreRunGuard" }
        Write-RunLog "$RunGuardHours-hour run guard bypassed by $guardBypassReason."
    }
    elseif (Test-Path -LiteralPath $RunGuardPath) {
        $guardRaw = Get-Content -LiteralPath $RunGuardPath -Raw -ErrorAction Stop
        $guard = $guardRaw | ConvertFrom-Json -ErrorAction Stop
        $lastStart = [datetime]$guard.StartTime
        $hoursSinceLastRun = ((Get-Date) - $lastStart).TotalHours

        if ($hoursSinceLastRun -lt $RunGuardHours) {
            $guardMessage = "A repair run already started on this computer within the last $RunGuardHours hours. LastRunId=$($guard.RunId); LastStart=$($guard.StartTime); LastStatus=$($guard.Status). Script stopped by run guard."
            Write-RunLog $guardMessage
            Write-Host $guardMessage -ForegroundColor Yellow
            Write-FinalStatusLine -Status "RUN_GUARD_ACTIVE" -ExitCode 3 -Detail $guardMessage
            try { Update-EndpointInstanceState -Status "RUN_GUARD_ACTIVE" -Force } catch { }
            if ($script:EndpointInstanceMutexAcquired -and $script:EndpointInstanceMutex) {
                try { $script:EndpointInstanceMutex.ReleaseMutex() } catch { }
            }
            if ($script:EndpointInstanceMutex) { try { $script:EndpointInstanceMutex.Dispose() } catch { } }
            exit 3
        }
    }
}
catch {
    Write-RunLog ("Run guard read failed; continuing. Error={0}" -f $_.Exception.Message)
}

try {
        [PSCustomObject]@{
            RunId       = $RunId
            Version     = $ScriptVersion
            StartTime   = (Get-Date).ToString("o")
        ComputerName = $ComputerName
        AllowDsregLeave    = [bool]$AllowDsregLeave
        AuditOnly    = [bool]$AuditOnly
        EntraHybridPending = [bool]$EntraHybridPending
        RetryAfterRebootTaskRun = [bool]$RetryAfterRebootTaskRun
        EndpointProcessId = $PID
        EndpointInstanceStatePath = $script:EndpointInstanceStatePath
        HeartbeatUtc = (Get-Date).ToUniversalTime().ToString("o")
        Status      = "RUNNING"
    } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $RunGuardPath -Encoding UTF8 -Force
}
catch {
    Write-RunLog ("Run guard write failed; continuing. Error={0}" -f $_.Exception.Message)
}

# ============================
# Retry and user-context task policy
# ============================
$RetrySleepMinutes = 20
$RetryMaxRetries   = 2
$UserTaskWaitSeconds = 30
$UserPrtRetryWaitSeconds = 30
$GpUpdateWaitSeconds = 30
$UserTaskFolderPath = "\SmartM365\IntuneHybridJoinToolkit"
$UserStatusTaskName = "\SmartM365\IntuneHybridJoinToolkit\SmartM365-IHJ-UserDsregStatus"
$UserRefreshPrtTaskName = "\SmartM365\IntuneHybridJoinToolkit\SmartM365-IHJ-UserRefreshPrt"
$UserIntuneAutoEnrollTaskName = "\SmartM365\IntuneHybridJoinToolkit\SmartM365-IHJ-UserMdmAutoEnroll"
$UserNextLogonAutoEnrollTaskName = "\SmartM365\IntuneHybridJoinToolkit\SmartM365-IHJ-RunUserAutoEnrollAtLogon"

Write-RunLog ("Retry policy: SleepMinutes={0}; MaxRetries={1}; Mode=LocalWait" -f $RetrySleepMinutes, $RetryMaxRetries)
Write-RunLog ("Intune retry policy: SleepMinutes={0}; MaxRetries={1}; Mode=LocalWait" -f $IntuneRetrySleepMinutes, $IntuneRetryMaxRetries)
Write-RunLog ("User task policy: Folder={0}; StatusTask={1}; RefreshPrtTask={2}; IntuneAutoEnrollTask={3}; NextLogonTask={4}; WaitSeconds={5}; RefreshWaitSeconds={6}" -f $UserTaskFolderPath, $UserStatusTaskName, $UserRefreshPrtTaskName, $UserIntuneAutoEnrollTaskName, $UserNextLogonAutoEnrollTaskName, $UserTaskWaitSeconds, $UserPrtRetryWaitSeconds)
Write-RunLog "DNS policy: FlushDNS=True"
Write-RunLog ("GPUpdate policy: Force=True; WaitSeconds={0}" -f $GpUpdateWaitSeconds)
Write-RunLog ("Non-Intune MDM removal policy: AllowRemoveNonIntuneMdmEnrollment={0}" -f [bool]$AllowRemoveNonIntuneMdmEnrollment)

$LegacyNextLogonAutoEnrollCmdPath = Join-Path $DataRoot "SmartM365-IHJ-RunUserAutoEnrollAtLogon.cmd"
try {
    if (Test-Path -LiteralPath $LegacyNextLogonAutoEnrollCmdPath) {
        Remove-Item -LiteralPath $LegacyNextLogonAutoEnrollCmdPath -Force -ErrorAction Stop
        Write-RunLog ("Removed legacy next-logon helper command: {0}" -f $LegacyNextLogonAutoEnrollCmdPath)
    }
}
catch {
    Write-RunLog ("Legacy next-logon helper command cleanup failed. Path={0}; Error={1}" -f $LegacyNextLogonAutoEnrollCmdPath,$_.Exception.Message)
}

try {
    $service = New-Object -ComObject Schedule.Service
    $service.Connect()
    $folder = $service.GetFolder($UserTaskFolderPath)
    $taskLeafName = [string](@($UserNextLogonAutoEnrollTaskName.Trim("\").Split("\"))[-1])
    $task = $folder.GetTask($taskLeafName)
    $legacyTaskAction = $false
    foreach ($action in @($task.Definition.Actions)) {
        $actionPath = [string]$action.Path
        $actionArguments = [string]$action.Arguments
        if (($actionPath -match '(?i)(^|\\)cmd\.exe$') -and ($actionArguments -match '(?i)SmartM365-IHJ-RunUserAutoEnrollAtLogon\.cmd')) {
            $legacyTaskAction = $true
        }
    }
    if ($legacyTaskAction) {
        $folder.DeleteTask($taskLeafName, 0)
        Write-RunLog ("Removed legacy next-logon task because it referenced a local helper command. TaskName={0}" -f $UserNextLogonAutoEnrollTaskName)
    }
}
catch {
    $hresult = $_.Exception.HResult
    if ($hresult -notin @(-2147024894, -2147024893)) {
        Write-RunLog ("Legacy next-logon task inspection failed. TaskName={0}; Error={1}" -f $UserNextLogonAutoEnrollTaskName,$_.Exception.Message)
    }
}

# Transcript (best effort)
$TranscriptStarted = $false
$TranscriptFile = Join-Path $TranscriptDir ("Transcript_IntuneHybridJoinToolkit_{0}_{1}.txt" -f $ComputerName, (Get-Date -Format "yyyyMMdd_HHmmss"))
try {
    Start-Transcript -Path $TranscriptFile -Force | Out-Null
    $TranscriptStarted = $true
    Write-RunLog "Transcript started: $TranscriptFile"
}
catch {
    Write-RunLog ("Transcript failed to start: {0}" -f $_.Exception.Message)
}

# ============================
# Helpers
# ============================
function Get-IntuneEnrollmentDetected {
    $state = Get-MdmEnrollmentState
    return [bool]$state.IntuneEnrollmentDetected
}

function Test-RegistrySubKeyExists {
    param(
        [Parameter(Mandatory=$true)]$BaseKey,
        [Parameter(Mandatory=$true)][string]$SubKeyPath
    )

    try {
        $key = $BaseKey.OpenSubKey($SubKeyPath, $false)
        if ($key) {
            $key.Close()
            return $true
        }
    }
    catch { }

    return $false
}

function Test-EnterpriseMgmtTaskFolderExists {
    param([Parameter(Mandatory=$true)][string]$EnrollmentId)

    try {
        $service = New-Object -ComObject Schedule.Service
        $service.Connect()
        [void]$service.GetFolder("\Microsoft\Windows\EnterpriseMgmt\$EnrollmentId")
        return $true
    }
    catch { return $false }
}

function Get-MdmEnrollmentState {
    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::LocalMachine,
            [Microsoft.Win32.RegistryView]::Registry64
        )
        $enrollKey = $base.OpenSubKey("SOFTWARE\Microsoft\Enrollments", $false)
        if (-not $enrollKey) {
            return [PSCustomObject]@{
                AnyMdmEnrollmentDetected = $false
                IntuneEnrollmentDetected = $false
                NonIntuneMdmEnrollmentDetected = $false
                EnrollmentCount = 0
                IntuneEnrollmentIds = ""
                NonIntuneEnrollmentIds = ""
                UnconfirmedIntuneEnrollmentIds = ""
                ProviderIds = ""
                EnrollmentDetails = ""
                IgnoredEnrollmentDetails = ""
            }
        }

        $internalProviderIds = @("Deploy Authority", "Cloud Authority", "Local Authority")
        $entries = @()
        $ignoredEntries = @()
        $unconfirmedIntuneEntries = @()
        foreach ($subName in $enrollKey.GetSubKeyNames()) {
            try {
                $sub = $enrollKey.OpenSubKey($subName, $false)
                if (-not $sub) { continue }
                $providerId = [string]$sub.GetValue("ProviderID", "")
                $discoveryServiceFullUrl = [string]$sub.GetValue("DiscoveryServiceFullURL", "")
                $enrollmentType = [string]$sub.GetValue("EnrollmentType", "")
                $upn = [string]$sub.GetValue("UPN", "")
                $aadResourceId = [string]$sub.GetValue("AADResourceID", "")
                $hasProviderId = -not [string]::IsNullOrWhiteSpace($providerId)
                $hasDiscoveryUrl = -not [string]::IsNullOrWhiteSpace($discoveryServiceFullUrl)
                $isInternalProvider = $hasProviderId -and ($internalProviderIds -contains $providerId)
                $isIntuneProvider = ($providerId -eq "MS DM Server")
                $isIntuneDiscovery = ($discoveryServiceFullUrl -match "(?i)enrollment\.manage\.microsoft\.com")
                $isExternalProvider = $hasProviderId -and -not $isInternalProvider -and -not $isIntuneProvider
                $isExternalDiscovery = $hasDiscoveryUrl -and -not $isIntuneDiscovery

                $statusKeyPresent = Test-RegistrySubKeyExists -BaseKey $base -SubKeyPath ("SOFTWARE\Microsoft\Enrollments\Status\{0}" -f $subName)
                $omadmAccountPresent = Test-RegistrySubKeyExists -BaseKey $base -SubKeyPath ("SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\{0}" -f $subName)
                $policyProviderPresent = Test-RegistrySubKeyExists -BaseKey $base -SubKeyPath ("SOFTWARE\Microsoft\PolicyManager\Providers\{0}" -f $subName)
                $enterpriseMgmtTaskPresent = Test-EnterpriseMgmtTaskFolderExists -EnrollmentId $subName
                $evidence = @()
                if ($statusKeyPresent) { $evidence += "StatusKey" }
                if ($omadmAccountPresent) { $evidence += "OMADMAccount" }
                if ($policyProviderPresent) { $evidence += "PolicyProvider" }
                if ($enterpriseMgmtTaskPresent) { $evidence += "EnterpriseMgmtTasks" }
                $evidenceText = ($evidence -join ",")
                $evidenceCount = $evidence.Count

                $isIntuneCandidate = $isIntuneProvider -or $isIntuneDiscovery
                $isIntuneEnrollment = $isIntuneProvider
                $isNonIntuneMdmEnrollment = $isExternalProvider -or ($isExternalDiscovery -and ($evidenceCount -ge 1))
                $isMdm = $isIntuneEnrollment -or $isNonIntuneMdmEnrollment

                $entry = [PSCustomObject]@{
                    EnrollmentId = $subName
                    ProviderID = $providerId
                    DiscoveryServiceFullURL = $discoveryServiceFullUrl
                    EnrollmentType = $enrollmentType
                    UPN = $upn
                    AADResourceID = $aadResourceId
                    Evidence = $evidenceText
                    IsIntune = $isIntuneEnrollment
                }

                if ($isMdm) {
                    $entries += $entry
                }
                elseif ($isIntuneCandidate) {
                    $unconfirmedIntuneEntries += $entry
                    $ignoredEntries += $entry
                }
                elseif ($hasProviderId -or $hasDiscoveryUrl -or (-not [string]::IsNullOrWhiteSpace($enrollmentType))) {
                    $ignoredEntries += $entry
                }
            }
            catch { }
        }

        $intuneEntries = @($entries | Where-Object { $_.IsIntune })
        $nonIntuneEntries = @($entries | Where-Object { -not $_.IsIntune })
        $providerIds = @($entries | ForEach-Object { if ([string]::IsNullOrWhiteSpace($_.ProviderID)) { "<empty>" } else { $_.ProviderID } } | Select-Object -Unique)
        $details = @($entries | ForEach-Object {
            "EnrollmentId={0},ProviderID={1},DiscoveryURL={2},EnrollmentType={3},UPN={4},Evidence={5}" -f $_.EnrollmentId,$_.ProviderID,$_.DiscoveryServiceFullURL,$_.EnrollmentType,$_.UPN,$_.Evidence
        })
        $ignoredDetails = @($ignoredEntries | ForEach-Object {
            "EnrollmentId={0},ProviderID={1},DiscoveryURL={2},EnrollmentType={3},UPN={4},Evidence={5}" -f $_.EnrollmentId,$_.ProviderID,$_.DiscoveryServiceFullURL,$_.EnrollmentType,$_.UPN,$_.Evidence
        })

        return [PSCustomObject]@{
            AnyMdmEnrollmentDetected = ($entries.Count -gt 0)
            IntuneEnrollmentDetected = ($intuneEntries.Count -gt 0)
            NonIntuneMdmEnrollmentDetected = ($nonIntuneEntries.Count -gt 0)
            EnrollmentCount = $entries.Count
            IntuneEnrollmentIds = (($intuneEntries | ForEach-Object { $_.EnrollmentId }) -join ";")
            NonIntuneEnrollmentIds = (($nonIntuneEntries | ForEach-Object { $_.EnrollmentId }) -join ";")
            UnconfirmedIntuneEnrollmentIds = (($unconfirmedIntuneEntries | ForEach-Object { $_.EnrollmentId }) -join ";")
            ProviderIds = ($providerIds -join ";")
            EnrollmentDetails = ($details -join " | ")
            IgnoredEnrollmentDetails = ($ignoredDetails -join " | ")
        }
    }
    catch { throw }
}

function Remove-NonIntuneMdmEnrollment {
    param(
        [Parameter(Mandatory=$true)][string]$EnrollmentIds,
        [Parameter(Mandatory=$true)][string]$OutputDirPath,
        [Parameter(Mandatory=$true)][string]$ComputerNameValue,
        [Parameter(Mandatory=$true)][string]$RunIdValue,
        [string]$RemovalLabel = "Non-Intune MDM enrollment"
    )

    $ids = @($EnrollmentIds -split ";" | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
    $safeRemovalLabel = ($RemovalLabel.ToLowerInvariant() -replace '[^a-z0-9]+', '_').Trim('_')
    if ([string]::IsNullOrWhiteSpace($safeRemovalLabel)) { $safeRemovalLabel = "mdm_enrollment" }
    $backupDir = Join-Path $OutputDirPath ("{0}_{1}_backup_{2}" -f $ComputerNameValue,$safeRemovalLabel,$RunIdValue)
    $removedItems = New-Object System.Collections.Generic.List[string]
    $warnings = New-Object System.Collections.Generic.List[string]
    $errors = New-Object System.Collections.Generic.List[string]

    if ($ids.Count -eq 0) {
        return [PSCustomObject]@{
            Success = $false
            EnrollmentIds = ""
            BackupDir = ""
            RemovedItems = ""
            Detail = ("No enrollment id was provided for {0}." -f $RemovalLabel)
        }
    }

    if (-not (Test-Path -LiteralPath $backupDir)) {
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
    }

    foreach ($id in $ids) {
        $safeId = $id -replace '[\\/:*?"<>|{}]', '_'

        $registryPaths = @(
            "HKLM\SOFTWARE\Microsoft\Enrollments\$id",
            "HKLM\SOFTWARE\Microsoft\Enrollments\Status\$id",
            "HKLM\SOFTWARE\Microsoft\EnterpriseResourceManager\Tracked\$id",
            "HKLM\SOFTWARE\Microsoft\PolicyManager\AdmxInstalled\$id",
            "HKLM\SOFTWARE\Microsoft\PolicyManager\Providers\$id",
            "HKLM\SOFTWARE\Microsoft\Provisioning\OMADM\Accounts\$id",
            "HKLM\SOFTWARE\Microsoft\Provisioning\OMADM\Logger\$id",
            "HKLM\SOFTWARE\Microsoft\Provisioning\OMADM\Sessions\$id",
            "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\MDM\Enrollments\$id"
        )

        foreach ($regPath in $registryPaths) {
            $psPath = "Registry::$regPath"
            if (-not (Test-Path -LiteralPath $psPath)) { continue }

            $safeRegName = ($regPath -replace '[\\/:*?"<>|{} ]', '_')
            $backupFile = Join-Path $backupDir ("{0}_{1}.reg" -f $safeId,$safeRegName)
            try {
                $exportOutput = & reg.exe export $regPath $backupFile /y 2>&1
                $exportExitCode = $LASTEXITCODE
                if ($exportExitCode -ne 0) {
                    $warnings.Add(("Registry export failed before removal. Path={0}; ExitCode={1}; Output={2}" -f $regPath,$exportExitCode,(ConvertTo-CleanNativeOutput -Output $exportOutput)))
                }

                Remove-Item -LiteralPath $psPath -Recurse -Force -ErrorAction Stop
                $removedItems.Add($regPath)
            }
            catch {
                $errors.Add(("Registry removal failed. Path={0}; Error={1}" -f $regPath,$_.Exception.Message))
            }
        }

        $taskPath = "\Microsoft\Windows\EnterpriseMgmt\$id\"
        try {
            $tasks = @(Get-ScheduledTask -TaskPath $taskPath -ErrorAction SilentlyContinue)
            foreach ($task in $tasks) {
                try {
                    Unregister-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -Confirm:$false -ErrorAction Stop
                    $removedItems.Add(("ScheduledTask={0}{1}" -f $task.TaskPath,$task.TaskName))
                }
                catch {
                    $errors.Add(("Scheduled task removal failed. Task={0}{1}; Error={2}" -f $task.TaskPath,$task.TaskName,$_.Exception.Message))
                }
            }
        }
        catch {
            Write-RunLog ("Scheduled task enumeration skipped/failed. TaskPath={0}; Error={1}" -f $taskPath,$_.Exception.Message)
        }

        try {
            $schedule = New-Object -ComObject Schedule.Service
            $schedule.Connect()
            $enterpriseMgmtFolder = $schedule.GetFolder("\Microsoft\Windows\EnterpriseMgmt")
            $enterpriseMgmtFolder.DeleteFolder($id, 0)
            $removedItems.Add(("ScheduledTaskFolder=\Microsoft\Windows\EnterpriseMgmt\{0}" -f $id))
        }
        catch {
            Write-RunLog ("Scheduled task folder removal skipped/failed. Folder=\Microsoft\Windows\EnterpriseMgmt\{0}; Error={1}" -f $id,$_.Exception.Message)
        }
    }

    return [PSCustomObject]@{
        Success = ($errors.Count -eq 0)
        EnrollmentIds = ($ids -join ";")
        BackupDir = $backupDir
        RemovedItems = ($removedItems -join " | ")
        Detail = if ($errors.Count -eq 0) {
            $detail = ("{0} registry keys and EnterpriseMgmt tasks removed. Certificates were not removed." -f $RemovalLabel)
            if ($warnings.Count -gt 0) {
                $detail = "{0} Warnings: {1}" -f $detail,($warnings -join " | ")
            }
            $detail
        }
        else {
            $detail = ($errors -join " | ")
            if ($warnings.Count -gt 0) {
                $detail = "{0} Warnings: {1}" -f $detail,($warnings -join " | ")
            }
            $detail
        }
    }
}

function Write-AtomicCsvAppend {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][psobject]$RowObject,
        [Parameter(Mandatory=$true)][string]$RunIdValue
    )

    $tmpRow = "{0}.{1}.row.tmp" -f $Path, $RunIdValue
    $tmpNew = "{0}.{1}.new.tmp" -f $Path, $RunIdValue

    try {
        $RowObject | Export-Csv -Path $tmpRow -NoTypeInformation -Encoding UTF8 -Force

        if (-not (Test-Path $Path)) {
            Move-Item -Path $tmpRow -Destination $Path -Force
            return
        }

        $existing = @()
        try {
            $existing = Import-Csv -Path $Path -ErrorAction Stop
        }
        catch {
            $backup = "{0}.{1}.corrupt.bak" -f $Path, $RunIdValue
            Copy-Item -Path $Path -Destination $backup -Force
            $existing = @()
        }

        $newRow = Import-Csv -Path $tmpRow
        $combined = @()
        if ($existing) { $combined += $existing }
        if ($newRow)  { $combined += $newRow }

        $columns = @()
        foreach ($row in $combined) {
            foreach ($prop in $row.PSObject.Properties.Name) {
                if ($columns -notcontains $prop) { $columns += $prop }
            }
        }

        $normalized = foreach ($row in $combined) {
            $h = [ordered]@{}
            foreach ($column in $columns) {
                $p = $row.PSObject.Properties[$column]
                if ($p) { $h[$column] = $p.Value } else { $h[$column] = "" }
            }
            [PSCustomObject]$h
        }

        $normalized | Export-Csv -Path $tmpNew -NoTypeInformation -Encoding UTF8 -Force
        Move-Item -Path $tmpNew -Destination $Path -Force
    }
    finally {
        Remove-Item -Path $tmpRow -ErrorAction SilentlyContinue
        Remove-Item -Path $tmpNew -ErrorAction SilentlyContinue
    }
}

function Test-DomainControllerReachable {
    param([Parameter(Mandatory=$true)][string]$DomainName)

    $nltest = Join-Path $env:WINDIR "System32\nltest.exe"
    if (Test-Path $nltest) {
        try {
            $null = & $nltest /sc_query:$DomainName 2>&1
            $ec = $LASTEXITCODE
            if ($ec -eq 0) {
                return [PSCustomObject]@{ Reachable=$true; Method="nltest_sc_query"; Detail="nltest /sc_query returned 0." }
            }

            $null = & $nltest /dsgetdc:$DomainName 2>&1
            $ec2 = $LASTEXITCODE
            if ($ec2 -eq 0) {
                return [PSCustomObject]@{ Reachable=$true; Method="nltest_dsgetdc"; Detail="nltest /dsgetdc returned 0." }
            }

            return [PSCustomObject]@{ Reachable=$false; Method="nltest"; Detail=("nltest failed. sc_query={0}, dsgetdc={1}" -f $ec, $ec2) }
        }
        catch {
            return [PSCustomObject]@{ Reachable=$false; Method="nltest_exception"; Detail=$_.Exception.Message }
        }
    }

    return [PSCustomObject]@{ Reachable=$false; Method="no_method"; Detail="nltest.exe not found." }
}

function Invoke-DsregStatusToFile {
    param(
        [Parameter(Mandatory=$true)][string]$DsregcmdPath,
        [Parameter(Mandatory=$true)][string]$TempFilePath
    )

    $output = & $DsregcmdPath /status 2>&1
    $output | Out-File -FilePath $TempFilePath -Encoding UTF8 -Force
    return $true
}

function Read-DsregStatusLines {
    param([Parameter(Mandatory=$true)][string]$TempFilePath)

    $lines = Get-Content -Path $TempFilePath -ErrorAction Stop

    $isEmpty =
        (-not $lines) -or
        ($lines -is [string] -and [string]::IsNullOrWhiteSpace($lines)) -or
        ($lines -is [System.Array] -and $lines.Count -eq 0)

    if ($isEmpty) {
        Start-Sleep -Seconds 2
        $lines = Get-Content -Path $TempFilePath -ErrorAction Stop
    }

    $isEmpty2 =
        (-not $lines) -or
        ($lines -is [string] -and [string]::IsNullOrWhiteSpace($lines)) -or
        ($lines -is [System.Array] -and $lines.Count -eq 0)

    if ($isEmpty2) {
        throw "dsregcmd /status output was empty at read time (race/flush/EDR). File: $TempFilePath"
    }

    if ($lines -is [string]) { return @($lines) }
    return [string[]]$lines
}

function Parse-DsregStatus {
    param([Parameter(Mandatory=$false)][string[]]$Lines)

    # SYSTEM dsregcmd /status is used only for device/join state.
    # User PRT/WAM fields must be evaluated from Parse-DsregUserStatus.
    if (-not $Lines -or $Lines.Count -eq 0) {
        return [PSCustomObject]@{
            AzureAdJoined       = ""
            DeviceId            = ""
            TenantName          = ""
            TenantId            = ""
            DeviceAuthStatusRaw = ""
            ClientErrorCode     = ""
            ServerErrorCode     = ""
            ServerErrorSubCode  = ""
            ServerOperation     = ""
            ServerMessage       = ""
            HttpsStatus         = ""
            RequestId           = ""
            ErrorPhase          = ""
            MdmUrl              = ""
            MdmTouUrl           = ""
            MdmComplianceUrl    = ""
            KeySignTest         = ""
        }
    }

    $getValueAny = {
        param([string[]]$patterns)

        foreach ($p in $patterns) {
            $line = $Lines | Where-Object { $_ -match $p } | Select-Object -First 1
            if ($line) {
                $value = ($line -split ":",2)[1].Trim()
                while ($value.StartsWith(":")) {
                    $value = $value.TrimStart(":").Trim()
                }
                return $value
            }
        }
        return ""
    }

    [PSCustomObject]@{
        AzureAdJoined       = (& $getValueAny @("^\s*AzureAdJoined\s*:","^\s*AzureAdJoined\s*:\s*"))
        DeviceId            = (& $getValueAny @("^\s*DeviceId\s*:","^\s*DeviceId\s*:\s*"))
        TenantName          = (& $getValueAny @("^\s*TenantName\s*:","^\s*TenantName\s*:\s*"))
        TenantId            = (& $getValueAny @("^\s*TenantId\s*:","^\s*TenantId\s*:\s*"))
        DeviceAuthStatusRaw = (& $getValueAny @("^\s*DeviceAuthStatus\s*:","^\s*DeviceAuthStatus\s*:\s*"))

        ClientErrorCode     = (& $getValueAny @("^\s*Client\s+ErrorCode\s*:","^\s*ClientErrorCode\s*:"))
        ServerErrorCode     = (& $getValueAny @("^\s*Server\s+ErrorCode\s*:","^\s*ServerErrorCode\s*:"))
        ServerErrorSubCode  = (& $getValueAny @("^\s*Server\s+ErrorSubCode\s*:","^\s*ServerErrorSubCode\s*:"))

        ServerOperation     = (& $getValueAny @("^\s*Server\s+Operation\s*:","^\s*ServerOperation\s*:"))
        ServerMessage       = (& $getValueAny @("^\s*Server\s+Message\s*:","^\s*ServerMessage\s*:"))
        HttpsStatus         = (& $getValueAny @("^\s*Https\s+Status\s*:","^\s*HttpsStatus\s*:"))
        RequestId           = (& $getValueAny @("^\s*Request\s+Id\s*:","^\s*RequestId\s*:"))
        ErrorPhase          = (& $getValueAny @("^\s*Error\s+Phase\s*:","^\s*ErrorPhase\s*:"))
        MdmUrl              = (& $getValueAny @("^\s*MdmUrl\s*:","^\s*MDMUrl\s*:"))
        MdmTouUrl           = (& $getValueAny @("^\s*MdmTouUrl\s*:","^\s*MDMTouUrl\s*:"))
        MdmComplianceUrl    = (& $getValueAny @("^\s*MdmComplianceUrl\s*:","^\s*MDMComplianceUrl\s*:"))
        KeySignTest         = (& $getValueAny @("^\s*KeySignTest\s*:"))
    }
}

function Write-ProvisionalRunState {
    param(
        [Parameter(Mandatory=$true)][string]$StatusValue,
        [Parameter(Mandatory=$true)][string]$DetailValue
    )

    try {
        $nextActionValue = Get-NextActionForStatus -Status $StatusValue -IntuneEnrolled ([bool]$intuneEnrolled)
        $provisional = [PSCustomObject]@{
            RunId                   = $RunId
            Timestamp               = Get-Date
            ComputerName            = $ComputerName
            ScriptVersion           = $ScriptVersion
            NextAction              = $nextActionValue
            Status                  = $StatusValue
            ExitCode                = 3
            OsCaption               = $osCaption
            OsVersion               = $osVersion
            OsBuildNumber           = $osBuildNumber
            OsArchitecture          = $osArchitecture
            OsProductType           = $osProductType
            DsregStatusErrorMessage = $DetailValue
            Dsreg_AzureAdJoined     = $dsregAzureAdJoined
            Dsreg_DeviceId          = $dsregDeviceId
            Dsreg_TenantId          = $dsregTenantId
            DeviceAuthStatus        = $dsregDeviceAuthStatus
            Dsreg_KeySignTest       = $dsregKeySignTest
            IntuneEnrolled          = $intuneEnrolled
            AutoEnrollPolicyChecked = $autoEnrollPolicyChecked
            AutoEnrollPolicyConfigured = $autoEnrollPolicyConfigured
            AutoEnrollUseAADCredentialType = $autoEnrollUseAADCredentialType
            AutoEnrollCredentialTypeLabel = $autoEnrollCredentialTypeLabel
            IntuneAutoEnrollMode    = $intuneAutoEnrollMode
            IntuneAutoEnrollTaskName= $intuneAutoEnrollTaskName
            IntuneAutoEnrollExitCode= $intuneAutoEnrollExitCode
            IntuneAutoEnrollDetail  = $intuneAutoEnrollDetail
            IntuneRetrySleepMinutes = $IntuneRetrySleepMinutes
            IntuneRetryMaxRetries   = $IntuneRetryMaxRetries
            IntuneRetryWindowMinutes = ($IntuneRetrySleepMinutes * $IntuneRetryMaxRetries)
            InteractiveUserDetected = $interactiveUserDetected
            InteractiveUserAccountName = $interactiveUserAccountName
            InteractiveUserAccountType = $interactiveUserAccountType
            InteractiveSessionName  = $interactiveSessionName
            InteractiveSessionIsRemote = $interactiveSessionIsRemote
            User_AzureAdPrt         = $userAzureAdPrt
            User_AzureAdPrtExpiryTime = $userAzureAdPrtExpiryTime
            User_RefreshPrtAttemptStatus = $userRefreshPrtAttemptStatus
            User_RefreshPrtHttpStatus = $userRefreshPrtHttpStatus
            User_RefreshPrtHttpError = $userRefreshPrtHttpError
            User_RefreshPrtReason   = $userRefreshPrtReason
            User_PrtRefreshStillNeeded = $userPrtRefreshStillNeeded
            User_IsUserAzureAD      = $userIsUserAzureAD
            User_SessionIsNotRemote = $userSessionIsNotRemote
        }

        Write-AtomicCsvAppend -Path $logPath -RowObject $provisional -RunIdValue ("{0}_provisional" -f $RunId)
        Write-RunLog ("Provisional CSV state written. Status={0}; NextAction={1}" -f $StatusValue,$nextActionValue)
    }
    catch {
        Write-RunLog ("Provisional CSV state write failed. Status={0}; Error={1}" -f $StatusValue,$_.Exception.Message)
    }
}

function Test-DsregDeviceHealthy {
    param(
        [Parameter(Mandatory=$false)][string]$AzureAdJoined,
        [Parameter(Mandatory=$false)][string]$DeviceAuthStatus,
        [Parameter(Mandatory=$false)][string]$KeySignTest
    )

    if ($AzureAdJoined -ne "YES") { return $false }
    if ($DeviceAuthStatus -like "*SUCCESS*") { return $true }

    # Older dsregcmd output can omit DeviceAuthStatus. In that case, KeySignTest=PASSED
    # plus AzureAdJoined=YES is the best local signal that the device key is usable.
    if ([string]::IsNullOrWhiteSpace($DeviceAuthStatus) -and $KeySignTest -eq "PASSED") {
        return $true
    }

    return $false
}

function Parse-DsregUserStatus {
    param([Parameter(Mandatory=$false)][string[]]$Lines)

    if (-not $Lines -or $Lines.Count -eq 0) {
        return [PSCustomObject]@{
            AzureAdPrt          = ""
            AzureAdPrtAuthority = ""
            AzureAdPrtUpdateTime = ""
            AzureAdPrtExpiryTime = ""
            RefreshPrtAttemptStatus = ""
            RefreshPrtHttpStatus = ""
            RefreshPrtHttpError = ""
            RefreshPrtServerErrorCode = ""
            RefreshPrtServerErrorSubCode = ""
            EnterprisePrt       = ""
            OnPremTgt           = ""
            CloudTgt            = ""
            WorkplaceJoined     = ""
            WamDefaultSet       = ""
            WamDefaultAuthority = ""
            WamDefaultId        = ""
            NgcSet              = ""
            IsUserAzureAD       = ""
            SessionIsNotRemote  = ""
        }
    }

    $getValueAny = {
        param([string[]]$patterns)

        foreach ($p in $patterns) {
            $line = $Lines | Where-Object { $_ -match $p } | Select-Object -First 1
            if ($line) { return ($line -split ":",2)[1].Trim() }
        }
        return ""
    }

    $getValueAnyLast = {
        param([string[]]$patterns)

        foreach ($p in $patterns) {
            $line = $Lines | Where-Object { $_ -match $p } | Select-Object -Last 1
            if ($line) { return ($line -split ":",2)[1].Trim() }
        }
        return ""
    }

    [PSCustomObject]@{
        AzureAdPrt          = (& $getValueAny @("^\s*AzureAdPrt\s*:"))
        AzureAdPrtAuthority = (& $getValueAny @("^\s*AzureAdPrtAuthority\s*:"))
        AzureAdPrtUpdateTime = (& $getValueAny @("^\s*AzureAdPrtUpdateTime\s*:"))
        AzureAdPrtExpiryTime = (& $getValueAny @("^\s*AzureAdPrtExpiryTime\s*:"))
        RefreshPrtAttemptStatus = (& $getValueAnyLast @("^\s*Attempt Status\s*:"))
        RefreshPrtHttpStatus = (& $getValueAnyLast @("^\s*HTTP status\s*:"))
        RefreshPrtHttpError = (& $getValueAnyLast @("^\s*HTTP Error\s*:"))
        RefreshPrtServerErrorCode = (& $getValueAnyLast @("^\s*Server Error Code\s*:"))
        RefreshPrtServerErrorSubCode = (& $getValueAnyLast @("^\s*Server Error SubCode\s*:"))
        EnterprisePrt       = (& $getValueAny @("^\s*EnterprisePrt\s*:"))
        OnPremTgt           = (& $getValueAny @("^\s*OnPremTgt\s*:"))
        CloudTgt            = (& $getValueAny @("^\s*CloudTgt\s*:"))
        WorkplaceJoined     = (& $getValueAny @("^\s*WorkplaceJoined\s*:"))
        WamDefaultSet       = (& $getValueAny @("^\s*WamDefaultSet\s*:"))
        WamDefaultAuthority = (& $getValueAny @("^\s*WamDefaultAuthority\s*:"))
        WamDefaultId        = (& $getValueAny @("^\s*WamDefaultId\s*:"))
        NgcSet              = (& $getValueAny @("^\s*NgcSet\s*:"))
        IsUserAzureAD       = (& $getValueAny @("^\s*IsUserAzureAD\s*:"))
        SessionIsNotRemote  = (& $getValueAny @("^\s*SessionIsNotRemote\s*:"))
    }
}

function Get-ParsedDsregStatusSnapshot {
    param(
        [Parameter(Mandatory=$true)][string]$DsregcmdPath,
        [Parameter(Mandatory=$true)][string]$OutputDirPath,
        [Parameter(Mandatory=$true)][string]$RunIdValue,
        [Parameter(Mandatory=$true)][string]$ComputerNameValue,
        [Parameter(Mandatory=$true)][string]$PhaseLabel
    )

    $tempStatus = Join-Path $env:TEMP ("dsreg_status_{0}_{1}.txt" -f $PhaseLabel, $RunIdValue)

    Write-RunLog ("Running dsregcmd /status snapshot. Phase={0}" -f $PhaseLabel)
    $null = Invoke-DsregStatusToFile -DsregcmdPath $DsregcmdPath -TempFilePath $tempStatus
    Start-Sleep -Seconds 2

    if (-not (Test-Path $tempStatus)) {
        throw "dsregcmd /status output file not found (Phase=$PhaseLabel)."
    }

    $lines = Read-DsregStatusLines -TempFilePath $tempStatus

    $copy = Join-Path $OutputDirPath ("{0}_dsreg_status_{1}_{2}.txt" -f $ComputerNameValue, $PhaseLabel, $RunIdValue)
    Copy-Item -Path $tempStatus -Destination $copy -Force
    Remove-Item -Path $tempStatus -ErrorAction SilentlyContinue

    return (Parse-DsregStatus -Lines $lines)
}

function Invoke-SchtasksRun {
    param([Parameter(Mandatory=$true)][string]$TaskName)

    $safeTaskName = $TaskName -replace '"','\"'
    $cmdLine = 'schtasks.exe /Run /TN "{0}" 2>&1' -f $safeTaskName
    $output = & $env:ComSpec /d /c $cmdLine
    $exitCode = $LASTEXITCODE

    return [PSCustomObject]@{
        ExitCode=$exitCode
        Output=(($output | Out-String).Trim())
    }
}

function Invoke-SchtasksCreate {
    param(
        [Parameter(Mandatory=$true)][string]$TaskName,
        [Parameter(Mandatory=$true)][string]$Trigger,
        [Parameter(Mandatory=$true)][string]$RunAs,
        [Parameter(Mandatory=$true)][string]$RunLevel,
        [Parameter(Mandatory=$true)][string]$TaskRun
    )

    $arguments = @(
        "/Create",
        "/TN", $TaskName,
        "/SC", $Trigger,
        "/RU", $RunAs,
        "/RL", $RunLevel,
        "/TR", $TaskRun,
        "/F"
    )

    $tempStdOut = Join-Path $env:TEMP ("EHJIR_schtasks_create_{0}_stdout.tmp" -f ([guid]::NewGuid().ToString("N")))
    $tempStdErr = Join-Path $env:TEMP ("EHJIR_schtasks_create_{0}_stderr.tmp" -f ([guid]::NewGuid().ToString("N")))
    try {
        $process = Start-Process -FilePath "schtasks.exe" -ArgumentList $arguments -PassThru -Wait -WindowStyle Hidden -RedirectStandardOutput $tempStdOut -RedirectStandardError $tempStdErr
        $parts = @()
        if (Test-Path -LiteralPath $tempStdOut) { $parts += Get-Content -LiteralPath $tempStdOut -ErrorAction SilentlyContinue }
        if (Test-Path -LiteralPath $tempStdErr) { $parts += Get-Content -LiteralPath $tempStdErr -ErrorAction SilentlyContinue }

        return [PSCustomObject]@{
            ExitCode = $process.ExitCode
            Output = (($parts | Out-String).Trim())
        }
    }
    finally {
        Remove-Item -LiteralPath $tempStdOut -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $tempStdErr -Force -ErrorAction SilentlyContinue
    }
}

function Ensure-ScheduledTaskFolder {
    param([Parameter(Mandatory=$true)][string]$TaskPath)

    $normalizedPath = $TaskPath.Trim()
    if ([string]::IsNullOrWhiteSpace($normalizedPath) -or $normalizedPath -eq "\") {
        return $true
    }

    if (-not $normalizedPath.StartsWith("\")) {
        $normalizedPath = "\$normalizedPath"
    }
    $normalizedPath = $normalizedPath.TrimEnd("\")

    try {
        $service = New-Object -ComObject Schedule.Service
        $service.Connect()
        [void]$service.GetFolder($normalizedPath)
        return $true
    }
    catch { }

    try {
        $service = New-Object -ComObject Schedule.Service
        $service.Connect()
        $root = $service.GetFolder("\")
        $parts = @($normalizedPath.Trim("\").Split("\") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        $currentFolder = $root
        $currentPath = ""

        foreach ($part in $parts) {
            $currentPath = if ([string]::IsNullOrWhiteSpace($currentPath)) { "\$part" } else { "$currentPath\$part" }
            try {
                $currentFolder = $service.GetFolder($currentPath)
            }
            catch {
                try {
                    [void]$currentFolder.CreateFolder($part, $null)
                }
                catch {
                    $hresult = $_.Exception.HResult
                    if ($hresult -ne -2147024713) {
                        throw
                    }
                    Write-RunLog ("Scheduled task folder already exists while creating it. Continuing. TaskPath={0}" -f $currentPath)
                    if ($currentPath -eq $normalizedPath) {
                        return $true
                    }
                }
                $currentFolder = $service.GetFolder($currentPath)
            }
        }

        return $true
    }
    catch {
        Write-RunLog ("Scheduled task folder creation failed. TaskPath={0}; Error={1}" -f $normalizedPath,$_.Exception.Message)
        return $false
    }
}

function ConvertTo-SchtasksRunArguments {
    param([Parameter(Mandatory=$true)][string]$TaskName)

    if ($TaskName -match '"') {
        throw "Scheduled task name contains an unsupported double quote: $TaskName"
    }

    return ('/Run /TN "{0}"' -f $TaskName)
}

function Register-InteractiveUserScheduledTask {
    param(
        [Parameter(Mandatory=$true)][string]$TaskName,
        [Parameter(Mandatory=$true)][string]$Description,
        [Parameter(Mandatory=$true)][string]$CommandArguments
    )

    $normalizedTaskName = $TaskName.Trim()
    if (-not $normalizedTaskName.StartsWith("\")) {
        $normalizedTaskName = "\$normalizedTaskName"
    }

    $parts = @($normalizedTaskName.Trim("\").Split("\") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($parts.Count -lt 2) {
        throw "Interactive user task must include a task folder: $TaskName"
    }

    $leafName = [string]$parts[-1]
    $folderPath = "\" + (($parts | Select-Object -First ($parts.Count - 1)) -join "\")
    if (-not (Ensure-ScheduledTaskFolder -TaskPath $folderPath)) {
        throw "Unable to create or open scheduled task folder: $folderPath"
    }

    $service = New-Object -ComObject Schedule.Service
    $service.Connect()
    $folder = $service.GetFolder($folderPath)
    $definition = $service.NewTask(0)

    $definition.RegistrationInfo.Author = "SmartM365 Intune Hybrid Join Toolkit"
    $definition.RegistrationInfo.Description = $Description
    $definition.Settings.Enabled = $true
    $definition.Settings.Hidden = $false
    $definition.Settings.AllowDemandStart = $true
    $definition.Settings.StartWhenAvailable = $true
    $definition.Settings.DisallowStartIfOnBatteries = $false
    $definition.Settings.StopIfGoingOnBatteries = $false
    $definition.Settings.ExecutionTimeLimit = "PT10M"
    $definition.Settings.MultipleInstances = 2

    $definition.Principal.GroupId = "S-1-5-4"
    $definition.Principal.LogonType = 4
    $definition.Principal.RunLevel = 0

    $action = $definition.Actions.Create(0)
    $action.Path = "cmd.exe"
    $action.Arguments = $CommandArguments

    [void]$folder.RegisterTaskDefinition($leafName, $definition, 6, $null, $null, 4, $null)

    return [PSCustomObject]@{
        Success = $true
        TaskName = $normalizedTaskName
        Detail = "Registered or updated interactive-user scheduled task."
    }
}

function Ensure-UserContextHelperTasks {
    $taskDefinitions = @(
        [PSCustomObject]@{
            Name = $UserStatusTaskName
            Description = "SmartM365 Intune Hybrid Join Toolkit - capture dsregcmd /status in the active interactive user context."
            Arguments = "/d /c %windir%\System32\dsregcmd.exe /status > %windir%\Temp\SmartM365-IHJ-UserDsregStatus_%USERNAME%.txt 2>&1"
        },
        [PSCustomObject]@{
            Name = $UserRefreshPrtTaskName
            Description = "SmartM365 Intune Hybrid Join Toolkit - run dsregcmd /refreshprt in the active interactive user context."
            Arguments = "/d /c %windir%\System32\dsregcmd.exe /refreshprt > %windir%\Temp\SmartM365-IHJ-UserRefreshPrt_%USERNAME%.txt 2>&1"
        },
        [PSCustomObject]@{
            Name = $UserIntuneAutoEnrollTaskName
            Description = "SmartM365 Intune Hybrid Join Toolkit - run deviceenroller /AutoEnrollMDM in the active interactive user context."
            Arguments = "/d /c %windir%\System32\deviceenroller.exe /c /AutoEnrollMDM > %windir%\Temp\SmartM365-IHJ-UserMdmAutoEnroll_%USERNAME%.txt 2>&1"
        }
    )

    $results = @()
    foreach ($taskDefinition in $taskDefinitions) {
        try {
            $result = Register-InteractiveUserScheduledTask -TaskName $taskDefinition.Name -Description $taskDefinition.Description -CommandArguments $taskDefinition.Arguments
            $results += $result
            Write-RunLog ("User-context helper task ready. TaskName={0}; Detail={1}" -f $result.TaskName,$result.Detail)
        }
        catch {
            $results += [PSCustomObject]@{ Success=$false; TaskName=$taskDefinition.Name; Detail=$_.Exception.Message }
            Write-RunLog ("User-context helper task registration failed. TaskName={0}; Error={1}" -f $taskDefinition.Name,$_.Exception.Message)
        }
    }

    return [PSCustomObject]@{
        Success = (@($results | Where-Object { -not $_.Success }).Count -eq 0)
        Results = $results
        Detail = (($results | ForEach-Object { "{0}: {1}" -f $_.TaskName,$_.Detail }) -join " | ")
    }
}
function Start-AutoDeviceJoinTask {
    try {
        $tn = "\Microsoft\Windows\Workplace Join\Automatic-Device-Join"
        $taskResult = Invoke-SchtasksRun -TaskName $tn
        $taskExitCode = $taskResult.ExitCode

        if ($taskExitCode -eq 0) {
            return $true
        }

        Write-RunLog ("Automatic-Device-Join task failed. ExitCode={0}; Output={1}" -f $taskExitCode, $taskResult.Output)
        return $false
    }
    catch {
        Write-RunLog ("Automatic-Device-Join task threw an exception: {0}" -f $_.Exception.Message)
        return $false
    }
}

function Start-IntuneAutoEnrollment {
    $deviceEnrollerPath = Join-Path $env:WINDIR "System32\deviceenroller.exe"
    if (-not (Test-Path $deviceEnrollerPath)) {
        return [PSCustomObject]@{
            Success=$false
            ToolFound=$false
            ExitCode=""
            Detail="deviceenroller.exe not found."
        }
    }

    try {
        Write-RunLog "Running deviceenroller.exe /c /AutoEnrollMDM."
        $enrollOutput = & $deviceEnrollerPath /c /AutoEnrollMDM 2>&1
        $enrollExitCode = $LASTEXITCODE
        $detail = (($enrollOutput | Out-String).Trim())

        if ($enrollExitCode -eq 0) {
            return [PSCustomObject]@{
                Success=$true
                ToolFound=$true
                ExitCode=$enrollExitCode
                Detail=$detail
            }
        }

        return [PSCustomObject]@{
            Success=$false
            ToolFound=$true
            ExitCode=$enrollExitCode
            Detail=$detail
        }
    }
    catch {
        return [PSCustomObject]@{
            Success=$false
            ToolFound=$true
            ExitCode=""
            Detail=$_.Exception.Message
        }
    }
}

function Get-NativeMdmAutoEnrollmentTask {
    try {
        $tasks = Get-ScheduledTask -ErrorAction Stop |
            Where-Object {
                $_.TaskPath -like "\Microsoft\Windows\EnterpriseMgmt\*" -and
                (
                    $_.TaskName -like "*automatically enrolling in MDM*" -or
                    $_.TaskName -like "*Schedule created by enrollment client*"
                )
            } |
            Select-Object -First 1

        if (-not $tasks) { return $null }

        return [PSCustomObject]@{
            TaskName=$tasks.TaskName
            TaskPath=$tasks.TaskPath
            FullName=("{0}{1}" -f $tasks.TaskPath,$tasks.TaskName)
        }
    }
    catch {
        Write-RunLog ("Failed to enumerate native EnterpriseMgmt auto-enrollment tasks: {0}" -f $_.Exception.Message)
        return $null
    }
}

function Start-UserContextIntuneAutoEnrollment {
    param(
        [Parameter(Mandatory=$true)][string]$PreferredTaskName,
        [Parameter(Mandatory=$true)][int]$WaitSeconds,
        [Parameter(Mandatory=$false)][string]$ExpectedUserName
    )

    $startedAt = Get-Date
    $customResult = Invoke-OnDemandScheduledTask -TaskName $PreferredTaskName
    if ($customResult.Success) {
        if ($WaitSeconds -gt 0) { Start-Sleep -Seconds $WaitSeconds }
        $expectedFileName = ""
        if (-not [string]::IsNullOrWhiteSpace($ExpectedUserName)) {
            $expectedFileName = "SmartM365-IHJ-UserMdmAutoEnroll_{0}.txt" -f $ExpectedUserName
        }
        $sourceFile = Get-LatestUserTaskOutputFile -Filter "SmartM365-IHJ-UserMdmAutoEnroll_*.txt" -Since $startedAt -ExpectedFileName $expectedFileName

        return [PSCustomObject]@{
            Success=$true
            ToolFound=$true
            ExitCode=$customResult.ExitCode
            Detail=("Triggered user-context Intune auto-enrollment task '{0}' resolved as '{1}'. {2}" -f $PreferredTaskName,$customResult.ResolvedTaskName,$customResult.Detail)
            Mode="UserTask"
            TaskName=$customResult.ResolvedTaskName
            OutputFile=($(if ($sourceFile) { $sourceFile.FullName } else { "" }))
        }
    }

    Write-RunLog ("Preferred user-context Intune auto-enrollment task did not run successfully. TaskName={0}; ResolvedTaskName={1}; ExitCode={2}; Detail={3}" -f $PreferredTaskName,$customResult.ResolvedTaskName,$customResult.ExitCode,$customResult.Detail)

    $nativeTask = Get-NativeMdmAutoEnrollmentTask
    if ($nativeTask) {
        $nativeResult = Invoke-OnDemandScheduledTask -TaskName $nativeTask.FullName
        return [PSCustomObject]@{
            Success=$nativeResult.Success
            ToolFound=$true
            ExitCode=$nativeResult.ExitCode
            Detail=("Triggered native EnterpriseMgmt auto-enrollment task '{0}'. {1}" -f $nativeTask.FullName,$nativeResult.Detail)
            Mode="NativeEnterpriseMgmtTask"
            TaskName=$nativeTask.FullName
            OutputFile=""
        }
    }

    return [PSCustomObject]@{
        Success=$false
        ToolFound=$false
        ExitCode=$customResult.ExitCode
        Detail=("No usable user-context Intune auto-enrollment task found. Preferred task '{0}' failed or is missing, and no native EnterpriseMgmt auto-enrollment task was found. Preferred task detail: {1}" -f $PreferredTaskName,$customResult.Detail)
        Mode="UserTaskMissing"
        TaskName=$PreferredTaskName
        OutputFile=""
    }
}

function Get-AutoEnrollmentPolicyState {
    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey(
            [Microsoft.Win32.RegistryHive]::LocalMachine,
            [Microsoft.Win32.RegistryView]::Registry64
        )
        $key = $base.OpenSubKey("SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\MDM", $false)

        if (-not $key) {
            return [PSCustomObject]@{
                PolicyKeyPresent=$false
                AutoEnrollMDM=""
                UseAADCredentialType=""
                CredentialTypeLabel=""
                IsConfigured=$false
                Detail="MDM auto-enrollment policy key not found."
            }
        }

        $autoEnrollMdm = $key.GetValue("AutoEnrollMDM", $null)
        $useAadCredentialType = $key.GetValue("UseAADCredentialType", $null)
        $credentialTypeLabel = switch ([string]$useAadCredentialType) {
            "1" { "UserCredential" }
            "2" { "DeviceCredential" }
            default { "" }
        }

        $isConfigured =
            ([string]$autoEnrollMdm -eq "1") -and
            ([string]$useAadCredentialType -in @("1","2"))

        $detail = if ($isConfigured) {
            "MDM auto-enrollment policy is configured."
        }
        else {
            "MDM auto-enrollment policy is present but not enabled/configured as expected."
        }

        return [PSCustomObject]@{
            PolicyKeyPresent=$true
            AutoEnrollMDM=([string]$autoEnrollMdm)
            UseAADCredentialType=([string]$useAadCredentialType)
            CredentialTypeLabel=$credentialTypeLabel
            IsConfigured=$isConfigured
            Detail=$detail
        }
    }
    catch {
        return [PSCustomObject]@{
            PolicyKeyPresent=$false
            AutoEnrollMDM=""
            UseAADCredentialType=""
            CredentialTypeLabel=""
            IsConfigured=$false
            Detail=$_.Exception.Message
        }
    }
}

function Escape-LdapFilterValue {
    param([AllowEmptyString()][string]$Value)

    if ($null -eq $Value) { return "" }

    return ($Value `
        -replace "\\", "\5c" `
        -replace "\*", "\2a" `
        -replace "\(", "\28" `
        -replace "\)", "\29" `
        -replace ([string][char]0), "\00")
}

function Get-LocalComputerAdLocation {
    param([Parameter(Mandatory=$true)][string]$ComputerName)

    $defaultNamingContext = ""
    $distinguishedName = ""
    $isDefaultComputersContainer = $false
    $detail = ""
    $success = $false

    try {
        $rootDse = [ADSI]"LDAP://RootDSE"
        $defaultNamingContext = [string]$rootDse.defaultNamingContext

        if ([string]::IsNullOrWhiteSpace($defaultNamingContext)) {
            throw "RootDSE defaultNamingContext is empty."
        }

        $searchRoot = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$defaultNamingContext")
        $searcher = New-Object System.DirectoryServices.DirectorySearcher($searchRoot)
        $searcher.PageSize = 1
        $searcher.SearchScope = [System.DirectoryServices.SearchScope]::Subtree
        [void]$searcher.PropertiesToLoad.Add("distinguishedName")
        $escapedSam = Escape-LdapFilterValue -Value ("{0}$" -f $ComputerName)
        $searcher.Filter = "(&(objectCategory=computer)(sAMAccountName=$escapedSam))"

        $result = $searcher.FindOne()
        if ($null -eq $result) {
            $detail = "Computer object was not found in the domain naming context."
        }
        else {
            $distinguishedName = [string]$result.Properties["distinguishedname"][0]
            $defaultComputersContainerDn = "CN=Computers,$defaultNamingContext"
            $parentDn = ""
            $firstComma = $distinguishedName.IndexOf(",")
            if ($firstComma -ge 0 -and $firstComma -lt ($distinguishedName.Length - 1)) {
                $parentDn = $distinguishedName.Substring($firstComma + 1)
            }

            $isDefaultComputersContainer = ($parentDn -ieq $defaultComputersContainerDn)
            if ($isDefaultComputersContainer) {
                $detail = "Computer object is in the default AD Computers container; OU-linked auto-enrollment GPOs may not apply."
            }
            else {
                $detail = "Computer object is not in the default AD Computers container."
            }
        }

        $success = $true
    }
    catch {
        $detail = $_.Exception.Message
    }

    return [PSCustomObject]@{
        Success = $success
        DistinguishedName = $distinguishedName
        DefaultNamingContext = $defaultNamingContext
        IsDefaultComputersContainer = $isDefaultComputersContainer
        Detail = $detail
    }
}

function Get-HostNameFromUrl {
    param([AllowEmptyString()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return "" }

    try {
        $uri = [Uri]$Value
        if (-not [string]::IsNullOrWhiteSpace($uri.Host)) { return $uri.Host }
    }
    catch { }

    return $Value
}

function Test-Tcp443Connectivity {
    param(
        [Parameter(Mandatory=$true)][string]$HostName,
        [Parameter(Mandatory=$false)][int]$TimeoutMs = 5000,
        [Parameter(Mandatory=$false)][int]$MaxAttempts = 3,
        [Parameter(Mandatory=$false)][int]$RetryDelaySeconds = 2
    )

    if ([string]::IsNullOrWhiteSpace($HostName)) {
        return [PSCustomObject]@{ Host=$HostName; Success=$false; Detail="Host is empty." }
    }

    $lastDetail = ""
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        $client = $null
        try {
            $client = New-Object System.Net.Sockets.TcpClient
            $async = $client.BeginConnect($HostName, 443, $null, $null)
            $completed = $async.AsyncWaitHandle.WaitOne($TimeoutMs, $false)

            if (-not $completed) {
                try { $client.Close() } catch { }
                $lastDetail = ("TCP 443 timeout after {0} ms. Attempt {1}/{2}." -f $TimeoutMs,$attempt,$MaxAttempts)
            }
            else {
                $client.EndConnect($async)
                return [PSCustomObject]@{ Host=$HostName; Success=$true; Detail=("TCP 443 reachable. Attempt {0}/{1}." -f $attempt,$MaxAttempts) }
            }
        }
        catch {
            $socketException = $null
            $currentException = $_.Exception
            while ($currentException) {
                if ($currentException -is [System.Net.Sockets.SocketException]) {
                    $socketException = $currentException
                    break
                }
                $currentException = $currentException.InnerException
            }

            if ($socketException) {
                $socketError = [string]$socketException.SocketErrorCode
                $networkDetail = switch ($socketError) {
                    "HostNotFound" { "DNS_HOST_NOT_FOUND"; break }
                    "NoData" { "DNS_NO_DATA"; break }
                    "TryAgain" { "DNS_TEMPORARY_FAILURE"; break }
                    "TimedOut" { "TCP_443_TIMEOUT"; break }
                    "NetworkUnreachable" { "NETWORK_UNREACHABLE"; break }
                    "HostUnreachable" { "HOST_UNREACHABLE"; break }
                    "ConnectionRefused" { "TCP_443_CONNECTION_REFUSED"; break }
                    "AccessDenied" { "TCP_443_ACCESS_DENIED"; break }
                    default { "TCP_443_CONNECT_FAILED" }
                }
                $lastDetail = ("{0}. SocketError={1}; NativeErrorCode={2}; Attempt {3}/{4}." -f $networkDetail,$socketError,$socketException.NativeErrorCode,$attempt,$MaxAttempts)
            }
            else {
                $lastDetail = ("TCP_443_CONNECT_FAILED. ErrorType={0}; Attempt {1}/{2}." -f $_.Exception.GetType().FullName,$attempt,$MaxAttempts)
            }
        }
        finally {
            if ($client) { try { $client.Close() } catch { } }
        }

        if ($attempt -lt $MaxAttempts -and $RetryDelaySeconds -gt 0) {
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }

    return [PSCustomObject]@{ Host=$HostName; Success=$false; Detail=$lastDetail }
}

function Test-EnrollmentConnectivity {
    param(
        [Parameter(Mandatory=$false)][string[]]$AdditionalUrls
    )

    $hosts = @(
        "login.microsoftonline.com",
        "device.login.microsoftonline.com",
        "enterpriseregistration.windows.net",
        "enrollment.manage.microsoft.com"
    )

    foreach ($url in @($AdditionalUrls | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $hostName = Get-HostNameFromUrl -Value $url
        if (-not [string]::IsNullOrWhiteSpace($hostName) -and $hosts -notcontains $hostName) {
            $hosts += $hostName
        }
    }

    $results = foreach ($hostName in $hosts) {
        Test-Tcp443Connectivity -HostName $hostName
    }

    $failed = @($results | Where-Object { -not $_.Success })
    $failedText = ($failed | ForEach-Object { "{0} ({1})" -f $_.Host,$_.Detail }) -join "; "

    return [PSCustomObject]@{
        Success=($failed.Count -eq 0)
        TestedHosts=($hosts -join ";")
        FailedHosts=$failedText
        Results=$results
    }
}

function Resolve-ScheduledTaskFullName {
    param([Parameter(Mandatory=$true)][string]$TaskName)

    try {
        if ($TaskName.StartsWith("\")) { return $TaskName }

        $task = Get-ScheduledTask -ErrorAction Stop |
            Where-Object { $_.TaskName -eq $TaskName } |
            Select-Object -First 1

        if ($task) {
            return ("{0}{1}" -f $task.TaskPath,$task.TaskName)
        }
    }
    catch {
        Write-RunLog ("Scheduled task lookup failed for '{0}': {1}" -f $TaskName,$_.Exception.Message)
    }

    return $TaskName
}

function Invoke-OnDemandScheduledTask {
    param([Parameter(Mandatory=$true)][string]$TaskName)

    $resolvedTaskName = Resolve-ScheduledTaskFullName -TaskName $TaskName

    try {
        $taskResult = Invoke-SchtasksRun -TaskName $resolvedTaskName
        $taskExitCode = $taskResult.ExitCode
        $detail = $taskResult.Output

        return [PSCustomObject]@{
            Success=($taskExitCode -eq 0)
            ExitCode=$taskExitCode
            Detail=$detail
            ResolvedTaskName=$resolvedTaskName
        }
    }
    catch {
        return [PSCustomObject]@{
            Success=$false
            ExitCode=""
            Detail=$_.Exception.Message
            ResolvedTaskName=$resolvedTaskName
        }
    }
}

function Get-LatestUserTaskOutputFile {
    param(
        [Parameter(Mandatory=$true)][string]$Filter,
        [Parameter(Mandatory=$true)][datetime]$Since,
        [Parameter(Mandatory=$false)][string]$ExpectedFileName
    )

    $tempDir = Join-Path $env:WINDIR "Temp"
    $minTime = $Since.AddSeconds(-5)

    if (-not [string]::IsNullOrWhiteSpace($ExpectedFileName)) {
        $expectedPath = Join-Path $tempDir $ExpectedFileName
        $expectedItem = Get-Item -LiteralPath $expectedPath -ErrorAction SilentlyContinue
        if ($expectedItem -and $expectedItem.LastWriteTime -ge $minTime -and $expectedItem.Length -gt 0) {
            return $expectedItem
        }

        return $null
    }

    Get-ChildItem -LiteralPath $tempDir -Filter $Filter -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $minTime -and $_.Length -gt 0 } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
}

function Invoke-UserDsregStatusTaskSnapshot {
    param(
        [Parameter(Mandatory=$true)][string]$TaskName,
        [Parameter(Mandatory=$true)][string]$OutputDirPath,
        [Parameter(Mandatory=$true)][string]$RunIdValue,
        [Parameter(Mandatory=$true)][string]$ComputerNameValue,
        [Parameter(Mandatory=$true)][string]$PhaseLabel,
        [Parameter(Mandatory=$true)][int]$WaitSeconds,
        [Parameter(Mandatory=$false)][string]$ExpectedUserName
    )

    $startedAt = Get-Date
    Write-RunLog ("Running user-context dsreg status scheduled task. TaskName={0}; Phase={1}" -f $TaskName,$PhaseLabel)
    $taskResult = Invoke-OnDemandScheduledTask -TaskName $TaskName
    Write-RunLog ("User-context dsreg status task result. Success={0}; TaskName={1}; ResolvedTaskName={2}; ExitCode={3}; Detail={4}" -f $taskResult.Success,$TaskName,$taskResult.ResolvedTaskName,$taskResult.ExitCode,$taskResult.Detail)

    if ($WaitSeconds -gt 0) { Start-Sleep -Seconds $WaitSeconds }

    $expectedFileName = ""
    if (-not [string]::IsNullOrWhiteSpace($ExpectedUserName)) {
        $expectedFileName = "SmartM365-IHJ-UserDsregStatus_{0}.txt" -f $ExpectedUserName
    }

    $sourceFile = Get-LatestUserTaskOutputFile -Filter "SmartM365-IHJ-UserDsregStatus_*.txt" -Since $startedAt -ExpectedFileName $expectedFileName
    if (-not $sourceFile) {
        $expectedDetail = if ([string]::IsNullOrWhiteSpace($expectedFileName)) { "No recent SmartM365-IHJ-UserDsregStatus_*.txt output file found after running user status task." } else { ("Expected user output file not found or not recent after running user status task: C:\Windows\Temp\{0}" -f $expectedFileName) }
        return [PSCustomObject]@{
            Success=$false
            TaskSuccess=$taskResult.Success
            TaskExitCode=$taskResult.ExitCode
            SourceFile=""
            CopiedFile=""
            Parsed=(Parse-DsregUserStatus -Lines @())
            ErrorMessage=$expectedDetail
        }
    }

    $copy = Join-Path $OutputDirPath ("{0}_dsreg_user_status_{1}_{2}_{3}" -f $ComputerNameValue, $PhaseLabel, $RunIdValue, $sourceFile.Name)
    Copy-Item -LiteralPath $sourceFile.FullName -Destination $copy -Force
    $lines = Read-DsregStatusLines -TempFilePath $sourceFile.FullName
    $parsed = Parse-DsregUserStatus -Lines $lines

    return [PSCustomObject]@{
        Success=$true
        TaskSuccess=$taskResult.Success
        TaskExitCode=$taskResult.ExitCode
        SourceFile=$sourceFile.FullName
        CopiedFile=$copy
        Parsed=$parsed
        ErrorMessage=""
    }
}

function Invoke-UserRefreshPrtTask {
    param(
        [Parameter(Mandatory=$true)][string]$TaskName,
        [Parameter(Mandatory=$true)][int]$WaitSeconds,
        [Parameter(Mandatory=$false)][string]$ExpectedUserName
    )

    $startedAt = Get-Date
    Write-RunLog ("Running user-context refresh PRT scheduled task. TaskName={0}" -f $TaskName)
    $taskResult = Invoke-OnDemandScheduledTask -TaskName $TaskName
    Write-RunLog ("User-context refresh PRT task result. Success={0}; TaskName={1}; ResolvedTaskName={2}; ExitCode={3}; Detail={4}" -f $taskResult.Success,$TaskName,$taskResult.ResolvedTaskName,$taskResult.ExitCode,$taskResult.Detail)

    if ($WaitSeconds -gt 0) { Start-Sleep -Seconds $WaitSeconds }
    $expectedFileName = ""
    if (-not [string]::IsNullOrWhiteSpace($ExpectedUserName)) {
        $expectedFileName = "SmartM365-IHJ-UserRefreshPrt_{0}.txt" -f $ExpectedUserName
    }
    $sourceFile = Get-LatestUserTaskOutputFile -Filter "SmartM365-IHJ-UserRefreshPrt_*.txt" -Since $startedAt -ExpectedFileName $expectedFileName

    return [PSCustomObject]@{
        Success=$taskResult.Success
        ExitCode=$taskResult.ExitCode
        Detail=$taskResult.Detail
        SourceFile=($(if ($sourceFile) { $sourceFile.FullName } else { "" }))
    }
}

function Get-InteractiveUserSessionState {
    try {
        $raw = quser.exe 2>&1
        $exitCode = $LASTEXITCODE
        $lines = @($raw | ForEach-Object { [string]$_ })
        $activeStates = @(
            "Active",
            "Actif",
            "Attivo",
            "Attiva",
            "Activo",
            "Activa",
            "Aktiv",
            "Ativo",
            "Ativa",
            "Actief",
            "Aktywny",
            "Aktywna",
            "Aktywne"
        )
        $knownStates = @(
            "Active",
            "Actif",
            "Attivo",
            "Attiva",
            "Activo",
            "Activa",
            "Aktiv",
            "Ativo",
            "Ativa",
            "Actief",
            "Aktywny",
            "Aktywna",
            "Aktywne",
            "Disc",
            "Déco",
            "Deco",
            "Deconnecte",
            "Déconnecté",
            "Disconn",
            "Disconnesso",
            "Desconectado",
            "Desconectada",
            "Desligado",
            "Desligada",
            "Rozłączony",
            "Rozłączona",
            "Rozłączone",
            "Verbroken",
            "Ontkoppeld",
            "Getrennt"
        )

        if ($exitCode -ne 0 -or -not $lines -or $lines.Count -lt 2) {
            return [PSCustomObject]@{
                HasInteractiveUser=$false
                ActiveUser=""
                ActiveSessionName=""
                ActiveSessionId=""
                ActiveState=""
                Detail=(($lines | Out-String).Trim())
            }
        }

        foreach ($line in ($lines | Select-Object -Skip 1)) {
            $clean = $line.Trim()
            if ([string]::IsNullOrWhiteSpace($clean)) { continue }
            if ($clean.StartsWith(">")) { $clean = $clean.Substring(1).Trim() }

            $parts = @($clean -split "\s+")
            if ($parts.Count -lt 4) { continue }

            $stateIndex = -1
            $sessionIdIndex = -1
            for ($i = 1; $i -lt $parts.Count; $i++) {
                if ($parts[$i] -match "^\d+$") {
                    $sessionIdIndex = $i
                    if (($i + 1) -lt $parts.Count) {
                        $stateIndex = $i + 1
                    }
                    break
                }
            }

            if ($stateIndex -lt 0) {
                for ($i = 0; $i -lt $parts.Count; $i++) {
                    if ($parts[$i] -in $knownStates) {
                        $stateIndex = $i
                        break
                    }
                }
            }

            if ($stateIndex -lt 0) { continue }

            $userName = $parts[0]
            $sessionId = if ($sessionIdIndex -ge 0) { $parts[$sessionIdIndex] } elseif ($stateIndex -gt 0) { $parts[$stateIndex - 1] } else { "" }
            $sessionName = if ($sessionIdIndex -gt 1) { ($parts[1..($sessionIdIndex - 1)] -join " ") } elseif ($stateIndex -gt 2) { ($parts[1..($stateIndex - 2)] -join " ") } else { "" }
            $state = $parts[$stateIndex]

            if ($state -in $activeStates) {
                return [PSCustomObject]@{
                    HasInteractiveUser=$true
                    ActiveUser=$userName
                    ActiveSessionName=$sessionName
                    ActiveSessionId=$sessionId
                    ActiveState=$state
                    Detail=$clean
                }
            }
        }

        return [PSCustomObject]@{
            HasInteractiveUser=$false
            ActiveUser=""
            ActiveSessionName=""
            ActiveSessionId=""
            ActiveState=""
            Detail=(($lines | Out-String).Trim())
        }
    }
    catch {
        return [PSCustomObject]@{
            HasInteractiveUser=$false
            ActiveUser=""
            ActiveSessionName=""
            ActiveSessionId=""
            ActiveState=""
            Detail=$_.Exception.Message
        }
    }
}

function Resolve-InteractiveUserAccount {
    param(
        [Parameter(Mandatory=$false)][string]$SessionId,
        [Parameter(Mandatory=$false)][string]$UserName
    )

    $empty = [PSCustomObject]@{
        Resolved=$false
        Domain=""
        User=$UserName
        AccountName=$UserName
        AccountType="Unknown"
        Detail="No process owner found for the interactive session."
    }

    if ([string]::IsNullOrWhiteSpace($SessionId)) { return $empty }

    $sessionNumber = 0
    if (-not [int]::TryParse($SessionId, [ref]$sessionNumber)) { return $empty }

    try {
        $processes = @()
        $explorerProcesses = @(Get-CimInstance Win32_Process -Filter ("SessionId = {0} AND Name = 'explorer.exe'" -f $sessionNumber) -ErrorAction SilentlyContinue)
        if ($explorerProcesses.Count -gt 0) {
            $processes += $explorerProcesses
        }
        else {
            $processes += @(Get-CimInstance Win32_Process -Filter ("SessionId = {0}" -f $sessionNumber) -ErrorAction SilentlyContinue)
        }

        foreach ($process in $processes) {
            try {
                $owner = Invoke-CimMethod -InputObject $process -MethodName GetOwner -ErrorAction Stop
                if ($owner.ReturnValue -ne 0 -or [string]::IsNullOrWhiteSpace($owner.User)) { continue }

                if (-not [string]::IsNullOrWhiteSpace($UserName) -and $owner.User -ne $UserName) { continue }

                $domain = [string]$owner.Domain
                $user = [string]$owner.User
                $accountName = if ([string]::IsNullOrWhiteSpace($domain)) { $user } else { "{0}\{1}" -f $domain,$user }
                $accountType = "Domain"

                if ($domain -eq $env:COMPUTERNAME) {
                    $accountType = "Local"
                }
                elseif ($domain -eq "AzureAD") {
                    $accountType = "AzureAD"
                }
                elseif ([string]::IsNullOrWhiteSpace($domain)) {
                    $accountType = "Unknown"
                }

                return [PSCustomObject]@{
                    Resolved=$true
                    Domain=$domain
                    User=$user
                    AccountName=$accountName
                    AccountType=$accountType
                    Detail=("Owner resolved from process {0} ({1})." -f $process.Name,$process.ProcessId)
                }
            }
            catch { }
        }
    }
    catch {
        return [PSCustomObject]@{
            Resolved=$false
            Domain=""
            User=$UserName
            AccountName=$UserName
            AccountType="Unknown"
            Detail=$_.Exception.Message
        }
    }

    return $empty
}

function Start-ControlledReboot {
    param(
        [Parameter(Mandatory=$true)][string]$Reason,
        [Parameter(Mandatory=$false)][int]$DelaySeconds = 180
    )

    Write-RunLog ("Triggering controlled reboot. DelaySeconds={0}; Reason={1}" -f $DelaySeconds,$Reason)
    $message = "Entra/Intune repair requires a reboot. Reason: $Reason"
    $output = shutdown.exe /r /t $DelaySeconds /c $message /f 2>&1
    $exitCode = $LASTEXITCODE
    Write-RunLog ("shutdown.exe exit code={0}; Output={1}" -f $exitCode,(($output | Out-String).Trim()))

    return [PSCustomObject]@{
        ExitCode=$exitCode
        Output=(($output | Out-String).Trim())
    }
}

function Invoke-ProcessWithTimeout {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [Parameter(Mandatory=$true)][string[]]$Arguments,
        [Parameter(Mandatory=$true)][int]$TimeoutSeconds,
        [Parameter(Mandatory=$true)][string]$OutputFile
    )

    $tempStdOut = Join-Path $env:TEMP ("EHJIR_{0}_stdout.tmp" -f ([guid]::NewGuid().ToString("N")))
    $tempStdErr = Join-Path $env:TEMP ("EHJIR_{0}_stderr.tmp" -f ([guid]::NewGuid().ToString("N")))
    $timedOut = $false
    $exitCode = ""
    $outputText = ""

    try {
        $process = Start-Process -FilePath $FilePath -ArgumentList $Arguments -PassThru -WindowStyle Hidden -RedirectStandardOutput $tempStdOut -RedirectStandardError $tempStdErr
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            $timedOut = $true
            try { $process.Kill() } catch { }
        }
        else {
            try { $process.Refresh() } catch { }
            $exitCode = $process.ExitCode
        }

        $parts = @()
        if (Test-Path -LiteralPath $tempStdOut) { $parts += Get-Content -LiteralPath $tempStdOut -ErrorAction SilentlyContinue }
        if (Test-Path -LiteralPath $tempStdErr) { $parts += Get-Content -LiteralPath $tempStdErr -ErrorAction SilentlyContinue }
        $outputText = (($parts | Out-String).Trim())
        $outputText | Out-File -LiteralPath $OutputFile -Encoding UTF8 -Force
    }
    finally {
        Remove-Item -LiteralPath $tempStdOut -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $tempStdErr -Force -ErrorAction SilentlyContinue
    }

    return [PSCustomObject]@{
        ExitCode = $exitCode
        TimedOut = $timedOut
        Output = $outputText
        OutputFile = $OutputFile
    }
}

function Export-ComputerGpResultEvidence {
    param(
        [Parameter(Mandatory=$true)][string]$ComputerName,
        [Parameter(Mandatory=$true)][string]$RunId,
        [Parameter(Mandatory=$true)][string]$OutputDir,
        [Parameter(Mandatory=$true)][string]$Reason
    )

    $htmlFile = Join-Path $OutputDir ("{0}_gpresult_computer_{1}.html" -f $ComputerName, $RunId)
    $textFile = Join-Path $OutputDir ("{0}_gpresult_computer_{1}.txt" -f $ComputerName, $RunId)
    Write-RunLog ("Exporting computer gpresult. Reason={0}; Html={1}; Text={2}" -f $Reason,$htmlFile,$textFile)

    try {
        $htmlStdoutFile = Join-Path $OutputDir ("{0}_gpresult_computer_{1}_html_stdout.txt" -f $ComputerName, $RunId)
        $htmlResult = Invoke-ProcessWithTimeout -FilePath "gpresult.exe" -Arguments @("/scope","computer","/h",$htmlFile,"/f") -TimeoutSeconds 300 -OutputFile $htmlStdoutFile
        if ($htmlResult.TimedOut) {
            Write-RunLog ("gpresult HTML export timed out after 300 seconds. Continuing. OutputFile={0}" -f $htmlStdoutFile)
        }
        else {
            Write-RunLog ("gpresult HTML export exit code={0}; OutputFile={1}; Output={2}" -f $htmlResult.ExitCode,$htmlStdoutFile,$htmlResult.Output)
        }
    }
    catch {
        Write-RunLog ("gpresult HTML export failed: {0}" -f $_.Exception.Message)
    }

    try {
        $textResult = Invoke-ProcessWithTimeout -FilePath "gpresult.exe" -Arguments @("/scope","computer","/r") -TimeoutSeconds 300 -OutputFile $textFile
        if ($textResult.TimedOut) {
            Write-RunLog ("gpresult text export timed out after 300 seconds. Continuing. OutputFile={0}" -f $textFile)
        }
        else {
            Write-RunLog ("gpresult text export exit code={0}; OutputFile={1}" -f $textResult.ExitCode,$textFile)
        }
    }
    catch {
        Write-RunLog ("gpresult text export failed: {0}" -f $_.Exception.Message)
    }

    return [PSCustomObject]@{
        HtmlFile = $htmlFile
        TextFile = $textFile
    }
}

function Register-UserAutoEnrollAtLogonTask {
    param(
        [Parameter(Mandatory=$true)][string]$StatusTaskName,
        [Parameter(Mandatory=$true)][string]$AutoEnrollTaskName
    )

    $taskName = $UserNextLogonAutoEnrollTaskName
    $legacyHelperCmdPath = Join-Path $DataRoot "SmartM365-IHJ-RunUserAutoEnrollAtLogon.cmd"

    try {
        $helperTaskResult = Ensure-UserContextHelperTasks
        if (-not $helperTaskResult.Success) {
            Write-RunLog ("Next-logon helper prerequisites were not fully registered. Detail={0}" -f $helperTaskResult.Detail)
        }

        $legacyCleanupDetail = "Legacy helper command was not present."
        if (Test-Path -LiteralPath $legacyHelperCmdPath) {
            try {
                Remove-Item -LiteralPath $legacyHelperCmdPath -Force -ErrorAction Stop
                $legacyCleanupDetail = ("Removed legacy helper command: {0}" -f $legacyHelperCmdPath)
            }
            catch {
                $legacyCleanupDetail = ("Legacy helper command could not be removed: {0}. Error={1}" -f $legacyHelperCmdPath,$_.Exception.Message)
            }
            Write-RunLog $legacyCleanupDetail
        }

        $taskFolderReady = Ensure-ScheduledTaskFolder -TaskPath $UserTaskFolderPath
        if (-not $taskFolderReady) {
            throw "Next-logon auto-enrollment helper task folder could not be confirmed: $UserTaskFolderPath"
        }

        $normalizedTaskName = $taskName.Trim()
        if (-not $normalizedTaskName.StartsWith("\")) {
            $normalizedTaskName = "\$normalizedTaskName"
        }

        $parts = @($normalizedTaskName.Trim("\").Split("\") | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
        if ($parts.Count -lt 2) {
            throw "Next-logon task must include a task folder: $taskName"
        }

        $leafName = [string]$parts[-1]
        $folderPath = "\" + (($parts | Select-Object -First ($parts.Count - 1)) -join "\")
        $schtasksPath = Join-Path $env:WINDIR "System32\schtasks.exe"
        $statusTaskArguments = ConvertTo-SchtasksRunArguments -TaskName $StatusTaskName
        $autoEnrollTaskArguments = ConvertTo-SchtasksRunArguments -TaskName $AutoEnrollTaskName

        $service = New-Object -ComObject Schedule.Service
        $service.Connect()
        $folder = $service.GetFolder($folderPath)
        $definition = $service.NewTask(0)

        $definition.RegistrationInfo.Author = "SmartM365 Intune Hybrid Join Toolkit"
        $definition.RegistrationInfo.Description = "SmartM365 Intune Hybrid Join Toolkit - trigger user-context auto-enrollment helper tasks at next user logon."
        $definition.Settings.Enabled = $true
        $definition.Settings.Hidden = $false
        $definition.Settings.AllowDemandStart = $true
        $definition.Settings.StartWhenAvailable = $true
        $definition.Settings.DisallowStartIfOnBatteries = $false
        $definition.Settings.StopIfGoingOnBatteries = $false
        $definition.Settings.ExecutionTimeLimit = "PT10M"
        $definition.Settings.MultipleInstances = 2

        $definition.Principal.UserId = "SYSTEM"
        $definition.Principal.LogonType = 5
        $definition.Principal.RunLevel = 1

        $trigger = $definition.Triggers.Create(9)
        $trigger.Enabled = $true

        $statusAction = $definition.Actions.Create(0)
        $statusAction.Path = $schtasksPath
        $statusAction.Arguments = $statusTaskArguments

        $autoEnrollAction = $definition.Actions.Create(0)
        $autoEnrollAction.Path = $schtasksPath
        $autoEnrollAction.Arguments = $autoEnrollTaskArguments

        [void]$folder.RegisterTaskDefinition($leafName, $definition, 6, "SYSTEM", $null, 5, $null)
        Write-RunLog ("Next-logon auto-enrollment helper task registration succeeded. TaskName={0}; Action={1}; StatusArgs={2}; AutoEnrollArgs={3}; LegacyCleanup={4}" -f $normalizedTaskName,$schtasksPath,$statusTaskArguments,$autoEnrollTaskArguments,$legacyCleanupDetail)

        return [PSCustomObject]@{
            Success = $true
            TaskName = $normalizedTaskName
            ExitCode = 0
            Detail = ("Registered SYSTEM ONLOGON task with direct schtasks actions; {0}" -f $legacyCleanupDetail)
        }
    }
    catch {
        Write-RunLog ("Next-logon auto-enrollment helper task registration failed: {0}" -f $_.Exception.Message)
        return [PSCustomObject]@{
            Success = $false
            TaskName = $taskName
            ExitCode = ""
            Detail = $_.Exception.Message
        }
    }
}

function Get-PreviousRunInfo {
    param([Parameter(Mandatory=$true)][string]$Path)

    try {
        if (Test-Path -LiteralPath $Path) {
            $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
            return ($raw | ConvertFrom-Json -ErrorAction Stop)
        }
    }
    catch { }

    return $null
}

function Get-NextActionForStatus {
    param(
        [Parameter(Mandatory=$true)][string]$Status,
        [Parameter(Mandatory=$false)][bool]$IntuneEnrolled = $false
    )

    switch ($Status) {
        "SUCCESS" { return "NO_ACTION_ALREADY_INTUNE_OR_HEALTHY" }
        "AUDIT_SUCCESS_ALREADY_INTUNE" { return "NO_ACTION_ALREADY_INTUNE" }
        "AUDIT_INTUNE_MISSING" { return "RUN_REPAIR" }
        "AUDIT_STALE_INTUNE_ENROLLMENT_LOCAL" { return "CLEAN_STALE_INTUNE_OPTIN" }
        "AUDIT_HYBRID_JOIN_UNHEALTHY" { return "FIX_HYBRID_JOIN" }
        "INTUNE_ENROLLMENT_TRIGGERED" { return "RECHECK_AFTER_AUTOENROLL" }
        "INTUNE_ENROLLMENT_PENDING_CONFIRMATION" { return "RECHECK_LATER_INTUNE_ENROLLMENT" }
        "INTUNE_ENROLLMENT_RETRY_EXHAUSTED" { return "RECHECK_OR_INVESTIGATE_INTUNE_ENROLLMENT" }
        "INTUNE_AUTOENROLL_POLICY_NOT_CONFIGURED" { return "CHECK_GPO_AUTOENROLL" }
        "INTUNE_ENROLLMENT_CONNECTIVITY_FAILED" { return "CHECK_CONNECTIVITY" }
        "INTUNE_USER_AUTOENROLL_NO_INTERACTIVE_USER" { return "WAIT_USER_LOGON" }
        "REBOOT_TRIGGERED_WAITING_FOR_USER_LOGON" { return "WAIT_USER_LOGON" }
        "WAITING_FOR_INTERACTIVE_USER_LOGON" { return "WAIT_USER_LOGON" }
        "INTUNE_USER_AUTOENROLL_LOCAL_INTERACTIVE_USER" { return "LOGON_WITH_DOMAIN_OR_AAD_USER" }
        "INTUNE_USER_AUTOENROLL_TASK_NOT_FOUND" { return "FIX_GPO_USER_AUTOENROLL_TASK" }
        "USER_NOT_AZUREAD" { return "CHECK_USER_AAD_OR_LOGON_CONTEXT" }
        "USER_PRT_NOT_AVAILABLE" { return "CHECK_USER_PRT" }
        "USER_PRT_REFRESH_FAILED" { return "FIX_USER_PRT_OR_RELOGIN" }
        "USER_SESSION_REMOTE" { return "LOGON_ON_CONSOLE" }
        "STALE_INTUNE_ENROLLMENT_LOCAL" { return "CLEAN_STALE_INTUNE_OPTIN" }
        "STALE_INTUNE_ENROLLMENT_REMOVED" { return "RECHECK_NEXT_CYCLE" }
        "NON_INTUNE_MDM_ENROLLED" { return "CLEAN_NON_INTUNE_MDM_OPTIN" }
        "NON_INTUNE_MDM_REMOVED" { return "RECHECK_NEXT_CYCLE" }
        "ENTRA_HYBRID_PENDING_ADJ_TRIGGERED" { return "RECHECK_ENTRA_PENDING_AFTER_ADJ" }
        "ENTRA_HYBRID_PENDING_RETRY_EXHAUSTED" { return "CHECK_AD_CONNECT_OR_DUPLICATE_ENTRA_DEVICE" }
        "WAITING_FOR_AAD_CONNECT_LOCAL_RETRY_EXHAUSTED" { return "FIX_HYBRID_JOIN_OR_AAD_CONNECT" }
        "KEY_SIGN_TEST_FAILED" { return "REPAIR_HYBRID_JOIN_KEY_OR_ALLOW_LEAVE" }
        "LEAVE_NOT_APPLICABLE" { return "FIX_HYBRID_JOIN" }
        "RUN_GUARD_ACTIVE" { return "WAIT_RUN_GUARD" }
        "ENDPOINT_RUN_ACTIVE" { return "WAIT_ACTIVE_ENDPOINT_RUN" }
        "REBOOT_SAFETY_LIMIT_REACHED_POST_DSREG_LEAVE" { return "REVIEW_REBOOT_HISTORY_AND_HYBRID_JOIN" }
        "RETRY_AFTER_REBOOT_EXHAUSTED" { return "CHECK_REBOOT_STATE_OR_RELAUNCH_LOT" }
        "RETRY_AFTER_REBOOT_STATE_MISSING" { return "RELAUNCH_LOT" }
        "RETRY_AFTER_REBOOT_SCHEDULE_FAILED_POST_DSREG_LEAVE" { return "CHECK_SCHEDULED_TASK_AND_RELAUNCH" }
        "RETRY_AFTER_REBOOT_SCHEDULE_FAILED_WAITING_FOR_USER_LOGON" { return "CHECK_SCHEDULED_TASK_AND_RELAUNCH" }
        "REBOOT_SCHEDULE_FAILED_POST_DSREG_LEAVE" { return "CHECK_REBOOT_COMMAND_AND_RELAUNCH" }
        "REBOOT_SCHEDULE_FAILED_WAITING_FOR_USER_LOGON" { return "CHECK_REBOOT_COMMAND_AND_RELAUNCH" }
        "SKIPPED_VIRTUAL_MACHINE" { return "NO_ACTION_VIRTUAL_MACHINE" }
        "COMPUTER_SYSTEM_QUERY_FAILED" { return "FIX_WMI_CIM_OR_RETRY" }
        "DOMAIN_CONTROLLER_UNREACHABLE" { return "FIX_DOMAIN_CONNECTIVITY_OR_VPN" }
        default {
            if ($Status -like "REBOOT_TRIGGERED*") { return "WAIT_REBOOT_AND_RECHECK" }
            if ($Status -like "*CONNECTIVITY*") { return "CHECK_CONNECTIVITY" }
            if ($Status -like "*POLICY*") { return "CHECK_GPO_AUTOENROLL" }
            if ($IntuneEnrolled) { return "NO_ACTION_INTUNE_PRESENT" }
            return "REVIEW_LOGS"
        }
    }
}

function Copy-LatestMatchingFile {
    param(
        [Parameter(Mandatory=$true)][string]$Filter,
        [Parameter(Mandatory=$true)][datetime]$Since,
        [Parameter(Mandatory=$true)][string]$DestinationDirectory,
        [Parameter(Mandatory=$true)][string]$Prefix
    )

    $sourceFile = Get-LatestUserTaskOutputFile -Filter $Filter -Since $Since
    if (-not $sourceFile) { return "" }

    $destination = Join-Path $DestinationDirectory ("{0}_{1}" -f $Prefix,$sourceFile.Name)
    Copy-Item -LiteralPath $sourceFile.FullName -Destination $destination -Force
    return $destination
}

function Export-RecentEventLog {
    param(
        [Parameter(Mandatory=$true)][string]$LogName,
        [Parameter(Mandatory=$true)][datetime]$Since,
        [Parameter(Mandatory=$true)][string]$Path
    )

    try {
        $events = Get-WinEvent -FilterHashtable @{ LogName=$LogName; StartTime=$Since } -ErrorAction Stop |
            Select-Object TimeCreated,Id,LevelDisplayName,ProviderName,Message

        if ($events) {
            $events | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 -Force
            return [PSCustomObject]@{ Success=$true; Path=$Path; Count=@($events).Count; Detail="" }
        }

        return [PSCustomObject]@{ Success=$true; Path=""; Count=0; Detail="No events found." }
    }
    catch {
        Write-RunLog ("Event log export failed. LogName={0}; Error={1}" -f $LogName,$_.Exception.Message)
        return [PSCustomObject]@{ Success=$false; Path=""; Count=0; Detail=$_.Exception.Message }
    }
}

function Export-EnrollmentDiagnostics {
    param(
        [Parameter(Mandatory=$true)][datetime]$Since,
        [Parameter(Mandatory=$true)][string]$OutputDirPath,
        [Parameter(Mandatory=$true)][string]$ComputerNameValue,
        [Parameter(Mandatory=$true)][string]$RunIdValue
    )

    $prefix = "{0}_enrollment_diag_{1}" -f $ComputerNameValue,$RunIdValue
    $files = @()

    $autoEnrollCopy = Copy-LatestMatchingFile `
        -Filter "SmartM365-IHJ-UserMdmAutoEnroll_*.txt" `
        -Since $Since `
        -DestinationDirectory $OutputDirPath `
        -Prefix $prefix

    if (-not [string]::IsNullOrWhiteSpace($autoEnrollCopy)) { $files += $autoEnrollCopy }

    $logs = @(
        @{ Name="Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin"; Suffix="DeviceManagement_Admin" },
        @{ Name="Microsoft-Windows-User Device Registration/Admin"; Suffix="UserDeviceRegistration_Admin" },
        @{ Name="Microsoft-Windows-AAD/Operational"; Suffix="AAD_Operational" }
    )

    foreach ($log in $logs) {
        $safeSuffix = $log.Suffix
        $path = Join-Path $OutputDirPath ("{0}_{1}.csv" -f $prefix,$safeSuffix)
        $result = Export-RecentEventLog -LogName $log.Name -Since $Since -Path $path
        if ($result.Success -and -not [string]::IsNullOrWhiteSpace($result.Path)) {
            $files += $result.Path
        }
    }

    return ($files -join ";")
}

function Get-LocalOsBootInfo {
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $lastBoot = $os.LastBootUpTime
        if (-not ($lastBoot -is [datetime])) {
            $lastBoot = [System.Management.ManagementDateTimeConverter]::ToDateTime([string]$os.LastBootUpTime)
        }

        $uptime = (Get-Date) - $lastBoot
        return [PSCustomObject]@{
            Success        = $true
            Caption        = [string]$os.Caption
            Version        = [string]$os.Version
            BuildNumber    = [string]$os.BuildNumber
            Architecture   = [string]$os.OSArchitecture
            ProductType    = [string]$os.ProductType
            LastBootUpTime = $lastBoot.ToString("yyyy-MM-dd HH:mm:ss")
            UptimeHours    = [math]::Round($uptime.TotalHours, 2)
            UptimeDays     = [math]::Round($uptime.TotalDays, 2)
            Detail         = ""
        }
    }
    catch {
        return [PSCustomObject]@{
            Success        = $false
            Caption        = ""
            Version        = ""
            BuildNumber    = ""
            Architecture   = ""
            ProductType    = ""
            LastBootUpTime = ""
            UptimeHours    = ""
            UptimeDays     = ""
            Detail         = $_.Exception.Message
        }
    }
}

function Get-ComputerSystemSummary {
    try {
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

        return [PSCustomObject]@{
            Success = $true
            ComputerSystem = $system
            IsVirtualMachine = $isVirtual
            VirtualMachineEvidence = $evidence
            Detail = ""
        }
    }
    catch {
        return [PSCustomObject]@{
            Success = $false
            ComputerSystem = $null
            IsVirtualMachine = ""
            VirtualMachineEvidence = ""
            Detail = $_.Exception.Message
        }
    }
}

function Stop-ComputerSystemQueryFailed {
    param([Parameter(Mandatory=$true)][string]$Detail)

    $status = "COMPUTER_SYSTEM_QUERY_FAILED"
    $errorMessage = "Unable to query Win32_ComputerSystem. Script stopped before domain, gpupdate, dsreg, cleanup, or reboot actions. Detail=$Detail"
    Write-Host $errorMessage -ForegroundColor Yellow
    Write-RunLog $errorMessage

    $logEntry = [PSCustomObject]@{
        RunId=$RunId; Timestamp=$Timestamp; ComputerName=$ComputerName; AllowDsregLeave=[bool]$AllowDsregLeave
        AllowRemoveNonIntuneMdmEnrollment=[bool]$AllowRemoveNonIntuneMdmEnrollment
        AllowRemoveStaleIntuneEnrollment=[bool]$AllowRemoveStaleIntuneEnrollment
        SkipVirtualMachines=[bool]$SkipVirtualMachines
        ScriptVersion=$ScriptVersion
        IsVirtualMachine=""
        VirtualMachineEvidence=""
        LeaveAttempted=$false; LeaveExitCode=""; Status=$status; ErrorMessage=$errorMessage
        Dsreg_AzureAdJoined=""; Dsreg_DeviceId=""; Dsreg_TenantName=""; Dsreg_TenantId=""
        DeviceAuthStatus=""; DsregStatusErrorMessage=$Detail; IntuneEnrolled=$null
        ClientErrorCode=""; ServerErrorCode=""; ServerErrorSubCode=""; ServerOperation=""
        ServerMessage=""; HttpsStatus=""; RequestId=""; ErrorPhase="ComputerSystemQuery"
    }

    Write-AtomicCsvAppend -Path $logPath -RowObject $logEntry -RunIdValue $RunId
    Write-FinalStatusLine -Status $status -ExitCode 2 -Detail $errorMessage -NextAction (Get-NextActionForStatus -Status $status)
    exit 2
}

function Invoke-BoundedRetryUntilIntuneEnrollment {
    param(
        [Parameter(Mandatory=$true)][int]$SleepMinutes,
        [Parameter(Mandatory=$true)][int]$MaxRetries,
        [Parameter(Mandatory=$true)][string]$ContextLabel
    )

    $attempt = 0

    while ($attempt -lt $MaxRetries) {
        $attempt++
        Write-RunLog ("{0}: Intune retry attempt {1}/{2}. Sleeping {3} minutes before re-check." -f $ContextLabel, $attempt, $MaxRetries, $SleepMinutes)
        Start-Sleep -Seconds ($SleepMinutes * 60)

        $detected = Get-IntuneEnrollmentDetected
        Write-RunLog ("{0}: Intune enrollment detected={1}" -f $ContextLabel, $detected)

        if ($detected) {
            return [PSCustomObject]@{ Success=$true; Attempts=$attempt }
        }
    }

    return [PSCustomObject]@{ Success=$false; Attempts=$MaxRetries }
}

function Invoke-BoundedRetryUntilSuccess {
    param(
        [Parameter(Mandatory=$true)][string]$DsregcmdPath,
        [Parameter(Mandatory=$true)][string]$OutputDirPath,
        [Parameter(Mandatory=$true)][string]$RunIdValue,
        [Parameter(Mandatory=$true)][string]$ComputerNameValue,
        [Parameter(Mandatory=$true)][int]$SleepMinutes,
        [Parameter(Mandatory=$true)][int]$MaxRetries,
        [Parameter(Mandatory=$true)][string]$ContextLabel
    )

    $attempt = 0

    while ($attempt -lt $MaxRetries) {
        $attempt++
        Write-RunLog ("{0}: Retry attempt {1}/{2}. Sleeping {3} minutes before re-check." -f $ContextLabel, $attempt, $MaxRetries, $SleepMinutes)
        Start-Sleep -Seconds ($SleepMinutes * 60)

        $phase = "{0}_retry{1}" -f $ContextLabel, $attempt
        $p = Get-ParsedDsregStatusSnapshot `
            -DsregcmdPath $DsregcmdPath `
            -OutputDirPath $OutputDirPath `
            -RunIdValue $RunIdValue `
            -ComputerNameValue $ComputerNameValue `
            -PhaseLabel $phase

        Write-RunLog ("{0}: AzureAdJoined={1}; DeviceAuthStatus={2}; KeySignTest={3}; ServerErrorSubCode={4}; ErrorPhase={5}; ClientErrorCode={6}" -f $phase, $p.AzureAdJoined, $p.DeviceAuthStatusRaw, $p.KeySignTest, $p.ServerErrorSubCode, $p.ErrorPhase, $p.ClientErrorCode)

        if (Test-DsregDeviceHealthy -AzureAdJoined $p.AzureAdJoined -DeviceAuthStatus $p.DeviceAuthStatusRaw -KeySignTest $p.KeySignTest) {
            return [PSCustomObject]@{ Success=$true; Parsed=$p; Attempts=$attempt }
        }
    }

    return [PSCustomObject]@{ Success=$false; Parsed=$null; Attempts=$MaxRetries }
}

function Protect-RetryAfterRebootPathAcl {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [switch]$Directory
    )

    $grants = if ($Directory) {
        @("*S-1-5-18:(OI)(CI)(F)","*S-1-5-32-544:(OI)(CI)(F)")
    }
    else {
        @("*S-1-5-18:(F)","*S-1-5-32-544:(F)")
    }

    $output = & icacls.exe $Path /inheritance:r /grant:r $grants 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw ("icacls.exe failed for '{0}' with exit code {1}: {2}" -f $Path,$LASTEXITCODE,(@($output) -join " "))
    }
}

function Set-RetryAfterRebootStateProperty {
    param(
        [Parameter(Mandatory=$true)]$State,
        [Parameter(Mandatory=$true)][string]$Name,
        $Value
    )

    if ($State.PSObject.Properties[$Name]) {
        $State.$Name = $Value
    }
    else {
        Add-Member -InputObject $State -NotePropertyName $Name -NotePropertyValue $Value -Force
    }
}

function Read-RetryAfterRebootState {
    if (-not (Test-Path -LiteralPath $script:RetryAfterRebootStatePath -PathType Leaf)) { return $null }
    return Get-Content -LiteralPath $script:RetryAfterRebootStatePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
}

function Write-RetryAfterRebootState {
    param([Parameter(Mandatory=$true)]$State)

    if (-not (Test-Path -LiteralPath $StateDir -PathType Container)) {
        New-Item -ItemType Directory -Path $StateDir -Force | Out-Null
    }

    $temporaryPath = "{0}.{1}.tmp" -f $script:RetryAfterRebootStatePath,[guid]::NewGuid().ToString("N")
    try {
        $State | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $temporaryPath -Encoding UTF8 -Force
        Move-Item -LiteralPath $temporaryPath -Destination $script:RetryAfterRebootStatePath -Force
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath) {
            Remove-Item -LiteralPath $temporaryPath -Force -ErrorAction SilentlyContinue
        }
    }
}
function Read-RebootSafetyState {
    if (-not (Test-Path -LiteralPath $script:RebootSafetyStatePath -PathType Leaf)) {
        return [pscustomobject]@{ Version=1; UpdatedUtc=(Get-Date).ToUniversalTime().ToString("o"); Records=@() }
    }
    try {
        $state = Get-Content -LiteralPath $script:RebootSafetyStatePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if (-not $state.PSObject.Properties["Records"]) { Add-Member -InputObject $state -NotePropertyName Records -NotePropertyValue @() -Force }
        return $state
    }
    catch {
        Write-RunLog ("Reboot safety state could not be read; fail-closed for reboot authorization. Error={0}" -f $_.Exception.Message)
        return $null
    }
}

function Get-RebootSafetyDecision {
    param(
        [Parameter(Mandatory=$true)][string]$ReasonKey,
        [string]$UserContext = "",
        [ValidateRange(1,30)][int]$MaximumReboots = 3,
        [switch]$RequireUserContextChange
    )

    $state = Read-RebootSafetyState
    if ($null -eq $state) {
        return [pscustomobject]@{ Allowed=$false; Count=0; Detail="Reboot safety state is unreadable; automatic reboot denied." }
    }
    $record = $state.Records | Where-Object { [string]$_.ReasonKey -eq $ReasonKey } | Select-Object -First 1
    $count = if ($record -and $record.PSObject.Properties["Count"]) { [int]$record.Count } else { 0 }
    if ($count -ge $MaximumReboots) {
        return [pscustomobject]@{ Allowed=$false; Count=$count; Detail=("Durable reboot safety limit reached for {0}: {1}/{2}." -f $ReasonKey,$count,$MaximumReboots) }
    }
    if ($RequireUserContextChange -and $count -gt 0) {
        $previousContext = if ($record.PSObject.Properties["LastUserContext"]) { [string]$record.LastUserContext } else { "" }
        if ([string]::IsNullOrWhiteSpace($UserContext) -or $UserContext -eq $previousContext) {
            return [pscustomobject]@{ Allowed=$false; Count=$count; Detail=("A reboot was already recorded for the same or absent interactive-user context. PreviousContext='{0}'; CurrentContext='{1}'. Waiting for a user/session change." -f $previousContext,$UserContext) }
        }
    }
    return [pscustomobject]@{ Allowed=$true; Count=$count; Detail=("Reboot safety authorization granted for {0}: current count {1}/{2}." -f $ReasonKey,$count,$MaximumReboots) }
}

function Register-RebootSafetyEvent {
    param(
        [Parameter(Mandatory=$true)][string]$ReasonKey,
        [string]$UserContext = "",
        [string]$Reason = ""
    )

    $state = Read-RebootSafetyState
    if ($null -eq $state) { throw "Reboot safety state is unreadable; event cannot be recorded." }
    $records = New-Object System.Collections.Generic.List[object]
    $existing = $null
    foreach ($record in @($state.Records)) {
        if ([string]$record.ReasonKey -eq $ReasonKey) { $existing = $record } else { [void]$records.Add($record) }
    }
    $count = if ($existing -and $existing.PSObject.Properties["Count"]) { [int]$existing.Count } else { 0 }
    [void]$records.Add([pscustomobject]@{
        ReasonKey = $ReasonKey
        Count = ($count + 1)
        LastRebootUtc = (Get-Date).ToUniversalTime().ToString("o")
        LastUserContext = $UserContext
        LastReason = $Reason
        LastRunId = $RunId
    })
    $state.UpdatedUtc = (Get-Date).ToUniversalTime().ToString("o")
    $state.Records = @($records.ToArray())
    Write-AtomicJsonFile -Path $script:RebootSafetyStatePath -Data $state
    Protect-RetryAfterRebootPathAcl -Path $script:RebootSafetyStatePath
    Write-RunLog ("Durable reboot safety event recorded. ReasonKey={0}; Count={1}; UserContext={2}" -f $ReasonKey,($count + 1),$UserContext)
}

function New-RetryAfterRebootArgumentList {
    $arguments = New-Object System.Collections.Generic.List[string]

    foreach ($switchName in @(
        "AllowDsregLeave",
        "IgnoreRunGuard",
        "AllowRebootWhenNoInteractiveUser",
        "AllowRebootAfterDsregLeave",
        "AllowRemoveNonIntuneMdmEnrollment",
        "AllowRemoveStaleIntuneEnrollment",
        "SkipVirtualMachines",
        "AuditOnly",
        "EntraHybridPending"
    )) {
        $value = Get-Variable -Name $switchName -ValueOnly -ErrorAction SilentlyContinue
        if ($value) { [void]$arguments.Add("-$switchName") }
    }

    $valueParameters = [ordered]@{
        StaleCleanupDelaySeconds = $StaleCleanupDelaySeconds
        RebootDelaySeconds = $RebootDelaySeconds
        IntuneRetrySleepMinutes = $IntuneRetrySleepMinutes
        IntuneRetryMaxRetries = $IntuneRetryMaxRetries
        RetryAfterRebootDelaySeconds = $RetryAfterRebootDelaySeconds
        RetryAfterRebootMaxAttempts = $RetryAfterRebootMaxAttempts
    }

    foreach ($entry in $valueParameters.GetEnumerator()) {
        [void]$arguments.Add("-$($entry.Key)")
        [void]$arguments.Add([string]$entry.Value)
    }

    return @($arguments)
}

function Write-RetryAfterRebootRunner {
    $escapedStatePath = $script:RetryAfterRebootStatePath.Replace("'","''")
    $runnerContent = @"
`$ErrorActionPreference = 'Stop'
`$statePath = '$escapedStatePath'
`$state = Get-Content -LiteralPath `$statePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
`$delaySeconds = 0
if (`$state.PSObject.Properties['DelaySeconds']) { `$delaySeconds = [int]`$state.DelaySeconds }
if (`$delaySeconds -gt 0) { Start-Sleep -Seconds `$delaySeconds }
`$arguments = @()
if (`$state.PSObject.Properties['Arguments']) { `$arguments = @(`$state.Arguments | ForEach-Object { [string]`$_ }) }
`$arguments += '-RetryAfterRebootTaskRun'
& ([string]`$state.ScriptPath) @arguments
`$exitCode = if (`$global:LASTEXITCODE -is [int]) { [int]`$global:LASTEXITCODE } else { 0 }
exit `$exitCode
"@
    $runnerContent | Set-Content -LiteralPath $script:RetryAfterRebootRunnerPath -Encoding UTF8 -Force
}

function Register-RetryAfterRebootTask {
    param([Parameter(Mandatory=$true)][string]$Reason)

    $script:RetryAfterRebootAction = "ScheduleRequested"
    if (-not (Test-Path -LiteralPath $StateDir -PathType Container)) {
        New-Item -ItemType Directory -Path $StateDir -Force | Out-Null
    }
    Protect-RetryAfterRebootPathAcl -Path $StateDir -Directory

    $previousAttempts = 0
    try {
        $existingState = Read-RetryAfterRebootState
        if ($existingState -and $existingState.PSObject.Properties["Attempts"]) {
            $previousAttempts = [int]$existingState.Attempts
        }
    }
    catch {
        Write-RunLog ("Existing retry-after-reboot state could not be read and will be replaced. Error={0}" -f $_.Exception.Message)
    }

    $state = [PSCustomObject]@{
        TaskName = $script:RetryAfterRebootTaskName
        ScriptPath = $script:EndpointScriptPath
        Arguments = @(New-RetryAfterRebootArgumentList)
        Attempts = $previousAttempts
        MaxAttempts = $RetryAfterRebootMaxAttempts
        DelaySeconds = $RetryAfterRebootDelaySeconds
        Reason = $Reason
        ComputerName = $ComputerName
        CreatedUtc = (Get-Date).ToUniversalTime().ToString("o")
        LastScheduledUtc = (Get-Date).ToUniversalTime().ToString("o")
    }

    Write-RetryAfterRebootState -State $state
    Write-RetryAfterRebootRunner
    Protect-RetryAfterRebootPathAcl -Path $script:RetryAfterRebootStatePath
    Protect-RetryAfterRebootPathAcl -Path $script:RetryAfterRebootRunnerPath

    $taskPowerShell = Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"
    $taskCommand = '"{0}" -NoProfile -ExecutionPolicy Bypass -File "{1}"' -f $taskPowerShell,$script:RetryAfterRebootRunnerPath
    $createOutput = & schtasks.exe /Create /TN $script:RetryAfterRebootTaskName /SC ONSTART /RU SYSTEM /RL HIGHEST /TR $taskCommand /F 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw ("schtasks.exe failed with exit code {0}: {1}" -f $LASTEXITCODE,(@($createOutput) -join " "))
    }

    $script:RetryAfterRebootAction = "Scheduled"
    $script:RetryAfterRebootAttempt = [string]$previousAttempts
    $script:RetryAfterRebootMaxAttemptsResult = [string]$RetryAfterRebootMaxAttempts
    $script:RetryAfterRebootTaskNameResult = $script:RetryAfterRebootTaskName
    $script:RetryAfterRebootDetail = "Task=$($script:RetryAfterRebootTaskName); DelaySeconds=$RetryAfterRebootDelaySeconds; MaxAttempts=$RetryAfterRebootMaxAttempts; Reason=$Reason"
    Write-RunLog ("Retry-after-reboot task scheduled. {0}" -f $script:RetryAfterRebootDetail)
}

function Start-RetryAfterRebootTaskRun {
    $script:RetryAfterRebootAction = "TaskRun"
    $script:RetryAfterRebootTaskNameResult = $script:RetryAfterRebootTaskName

    $state = Read-RetryAfterRebootState
    if (-not $state) {
        $script:RetryAfterRebootDetail = "Retry-after-reboot task run started but state file is missing."
        return [PSCustomObject]@{ Exhausted=$true; StateMissing=$true; Attempts=0; MaxAttempts=$RetryAfterRebootMaxAttempts }
    }

    $attempts = 0
    if ($state.PSObject.Properties["Attempts"]) { $attempts = [int]$state.Attempts }
    $attempts++

    $maxAttempts = $RetryAfterRebootMaxAttempts
    if ($state.PSObject.Properties["MaxAttempts"]) { $maxAttempts = [int]$state.MaxAttempts }

    Set-RetryAfterRebootStateProperty -State $state -Name "Attempts" -Value $attempts
    Set-RetryAfterRebootStateProperty -State $state -Name "LastAttemptUtc" -Value ((Get-Date).ToUniversalTime().ToString("o"))
    Write-RetryAfterRebootState -State $state
    Protect-RetryAfterRebootPathAcl -Path $script:RetryAfterRebootStatePath

    $script:RetryAfterRebootAttempt = [string]$attempts
    $script:RetryAfterRebootMaxAttemptsResult = [string]$maxAttempts
    $script:RetryAfterRebootDetail = "Task=$($script:RetryAfterRebootTaskName); Attempt=$attempts; MaxAttempts=$maxAttempts"
    Write-RunLog ("Retry-after-reboot task run started. {0}" -f $script:RetryAfterRebootDetail)

    return [PSCustomObject]@{ Exhausted=($attempts -gt $maxAttempts); StateMissing=$false; Attempts=$attempts; MaxAttempts=$maxAttempts }
}

function Unregister-RetryAfterRebootTask {
    param([string]$Reason = "")

    $messages = New-Object System.Collections.Generic.List[string]
    $queryOutput = & schtasks.exe /Query /TN $script:RetryAfterRebootTaskName 2>&1
    if ($LASTEXITCODE -eq 0) {
        $deleteOutput = & schtasks.exe /Delete /TN $script:RetryAfterRebootTaskName /F 2>&1
        if ($LASTEXITCODE -ne 0) {
            [void]$messages.Add(("Task delete failed: {0}" -f (@($deleteOutput) -join " ")))
        }
    }

    foreach ($path in @($script:RetryAfterRebootRunnerPath,$script:RetryAfterRebootStatePath)) {
        try {
            if (Test-Path -LiteralPath $path) {
                Remove-Item -LiteralPath $path -Force -ErrorAction Stop
            }
        }
        catch {
            [void]$messages.Add(("Remove failed for {0}: {1}" -f $path,$_.Exception.Message))
        }
    }

    if ($messages.Count -gt 0) {
        Write-RunLog ("Retry-after-reboot cleanup completed with warnings. Reason={0}; Detail={1}" -f $Reason,($messages -join "; "))
    }
    else {
        Write-RunLog ("Retry-after-reboot task and state cleaned up. Reason={0}" -f $Reason)
    }
}

# ============================
# Main
# ============================
$ExitCode = 0
$logPath = Join-Path $DataRoot ("IntuneHybridJoinToolkit_{0}.csv" -f $ComputerName)

# dsregcmd path
$dsregcmdPath = Join-Path $env:WINDIR "System32\dsregcmd.exe"
if (-not (Test-Path $dsregcmdPath)) { $dsregcmdPath = "dsregcmd.exe" }
Write-RunLog "Using dsregcmd path: $dsregcmdPath"

# Log vars
$status = ""
$errorMessage = ""
$intuneEnrolled = $null
$anyMdmEnrollmentDetected = $false
$nonIntuneMdmEnrollmentDetected = $false
$mdmEnrollmentCount = 0
$mdmProviderIds = ""
$intuneEnrollmentIds = ""
$nonIntuneEnrollmentIds = ""
$unconfirmedIntuneEnrollmentIds = ""
$mdmEnrollmentDetails = ""
$ignoredEnrollmentDetails = ""
$staleIntuneEnrollmentDetected = $false
$staleIntuneEnrollmentIds = ""
$nonIntuneMdmRemovalAttempted = $false
$nonIntuneMdmRemovalSuccess = ""
$nonIntuneMdmRemovalBackupDir = ""
$nonIntuneMdmRemovalDetail = ""
$staleIntuneRemovalAttempted = $false
$staleIntuneRemovalSuccess = ""
$staleIntuneRemovalBackupDir = ""
$staleIntuneRemovalDetail = ""
$staleIntuneCleanupContinuedToAutoEnroll = $false
$nextAction = ""
$nextLogonTaskRegistered = $false
$nextLogonTaskName = ""
$nextLogonTaskDetail = ""
$leaveAttempted = $false
$leaveExitCode = ""
$dsregStatusErrorMessage = ""

$osCaption = ""
$osVersion = ""
$osBuildNumber = ""
$osArchitecture = ""
$osProductType = ""

$dsregAzureAdJoined = ""
$dsregDeviceId = ""
$dsregTenantName = ""
$dsregTenantId = ""
$dsregDeviceAuthStatus = ""
$dsregKeySignTest = ""

$clientErrorCode = ""
$serverErrorCode = ""
$serverErrorSubCode = ""
$serverOperation = ""
$serverMessage = ""
$httpsStatus = ""
$requestId = ""
$errorPhase = ""
$mdmUrl = ""
$mdmTouUrl = ""
$mdmComplianceUrl = ""

$autoEnrollPolicyChecked = $false
$autoEnrollPolicyKeyPresent = ""
$autoEnrollPolicyConfigured = ""
$autoEnrollMDM = ""
$autoEnrollUseAADCredentialType = ""
$autoEnrollCredentialTypeLabel = ""
$autoEnrollPolicyDetail = ""

$enrollmentConnectivityChecked = $false
$enrollmentConnectivityOk = ""
$enrollmentConnectivityHosts = ""
$enrollmentConnectivityFailures = ""
$intuneAutoEnrollMode = ""
$intuneAutoEnrollTaskName = ""
$intuneAutoEnrollExitCode = ""
$intuneAutoEnrollDetail = ""
$intuneAutoEnrollOutputFile = ""
$intuneEnrollmentDiagnosticsFiles = ""
$intuneEnrollmentDiagnosticsInitialFiles = ""

$userStatusAttempted = $false
$userStatusTaskExitCode = ""
$userStatusFile = ""
$userStatusErrorMessage = ""
$userAzureAdPrt = ""
$userAzureAdPrtAuthority = ""
$userAzureAdPrtUpdateTime = ""
$userAzureAdPrtExpiryTime = ""
$userRefreshPrtAttemptStatus = ""
$userRefreshPrtHttpStatus = ""
$userRefreshPrtHttpError = ""
$userRefreshPrtServerErrorCode = ""
$userRefreshPrtServerErrorSubCode = ""
$userRefreshPrtReason = ""
$userPrtRefreshStillNeeded = $false
$userEnterprisePrt = ""
$userOnPremTgt = ""
$userCloudTgt = ""
$userWorkplaceJoined = ""
$userWamDefaultSet = ""
$userWamDefaultAuthority = ""
$userWamDefaultId = ""
$userNgcSet = ""
$userIsUserAzureAD = ""
$userSessionIsNotRemote = ""
$userRefreshPrtAttempted = $false
$userRefreshPrtExitCode = ""
$userRefreshPrtFile = ""
$interactiveUserDetected = $false
$interactiveUserName = ""
$interactiveSessionName = ""
$interactiveSessionId = ""
$interactiveSessionState = ""
$interactiveSessionDetail = ""
$interactiveSessionIsRemote = $false
$interactiveUserDomain = ""
$interactiveUserAccountName = ""
$interactiveUserAccountType = ""
$interactiveUserIdentityResolved = $false
$interactiveUserIdentityDetail = ""
$rebootAttempted = $false
$rebootReason = ""
$rebootExitCode = ""
$flushDnsAttempted = $false
$flushDnsExitCode = ""
$flushDnsOutputFile = ""
$lastBootUpTime = ""
$uptimeHours = ""
$uptimeDays = ""
$bootInfoDetail = ""
$gpUpdateAttempted = $false
$gpUpdateExitCode = ""
$gpUpdateOutputFile = ""
$gpUpdateMdmPolicyWarning = $false
$gpUpdateMdmPolicyState = ""
$gpUpdateMdmPolicyDetail = ""
$gpResultHtmlFile = ""
$gpResultTextFile = ""
$finalGpUpdateAttempted = $false
$finalGpUpdateReason = ""
$finalGpUpdateExitCode = ""
$finalGpUpdateOutputFile = ""
$domainPreflightComplete = $false
$cs = $null
$isVirtualMachine = ""
$virtualMachineEvidence = ""
$domainName = ""
$dcTest = $null
$adComputerLocationChecked = $false
$adComputerDistinguishedName = ""
$adComputerDefaultNamingContext = ""
$adComputerInDefaultComputersContainer = ""
$adComputerLocationDetail = ""

try {
    if ($RetryAfterRebootTaskRun) {
        $retryTaskRun = Start-RetryAfterRebootTaskRun
        if ($retryTaskRun.Exhausted) {
            $status = if ($retryTaskRun.StateMissing) { "RETRY_AFTER_REBOOT_STATE_MISSING" } else { "RETRY_AFTER_REBOOT_EXHAUSTED" }
            $ExitCode = 3
            $dsregStatusErrorMessage = if ($retryTaskRun.StateMissing) {
                "Retry-after-reboot task started without its state file. The task was removed; relaunch the LOT."
            }
            else {
                "Retry-after-reboot exceeded max attempts. Attempt=$($retryTaskRun.Attempts); MaxAttempts=$($retryTaskRun.MaxAttempts); TaskName=$($script:RetryAfterRebootTaskName)."
            }
            Unregister-RetryAfterRebootTask -Reason $status
            throw [System.OperationCanceledException]::new($dsregStatusErrorMessage)
        }
    }

    Write-Host "Running in LOCAL mode on computer: $ComputerName" -ForegroundColor Cyan
    Write-RunLog "Running LOCAL mode."

    $bootInfo = Get-LocalOsBootInfo
    $osCaption = $bootInfo.Caption
    $osVersion = $bootInfo.Version
    $osBuildNumber = $bootInfo.BuildNumber
    $osArchitecture = $bootInfo.Architecture
    $osProductType = $bootInfo.ProductType
    $lastBootUpTime = $bootInfo.LastBootUpTime
    $uptimeHours = $bootInfo.UptimeHours
    $uptimeDays = $bootInfo.UptimeDays
    $bootInfoDetail = $bootInfo.Detail
    if ($bootInfo.Success) {
        Write-RunLog ("OS info: Caption={0}; Version={1}; Build={2}; Architecture={3}; ProductType={4}; LastBootUpTime={5}; UptimeHours={6}; UptimeDays={7}" -f $osCaption,$osVersion,$osBuildNumber,$osArchitecture,$osProductType,$lastBootUpTime,$uptimeHours,$uptimeDays)
    }
    else {
        Write-RunLog ("OS boot info failed: {0}" -f $bootInfoDetail)
    }

    $computerSystemSummary = Get-ComputerSystemSummary
    if ($computerSystemSummary.Success) {
        $cs = $computerSystemSummary.ComputerSystem
        $isVirtualMachine = $computerSystemSummary.IsVirtualMachine
        $virtualMachineEvidence = $computerSystemSummary.VirtualMachineEvidence
        Write-RunLog ("Computer system: IsVirtualMachine={0}; Evidence={1}" -f $isVirtualMachine,$virtualMachineEvidence)
        if ($SkipVirtualMachines -and $isVirtualMachine) {
            $status = "SKIPPED_VIRTUAL_MACHINE"
            $errorMessage = "Virtual machine skipped by -SkipVirtualMachines. $virtualMachineEvidence"
            Write-Host $errorMessage -ForegroundColor Yellow
            Write-RunLog $errorMessage

            $logEntry = [PSCustomObject]@{
                RunId=$RunId; Timestamp=$Timestamp; ComputerName=$ComputerName; AllowDsregLeave=[bool]$AllowDsregLeave
                AllowRemoveNonIntuneMdmEnrollment=[bool]$AllowRemoveNonIntuneMdmEnrollment
                AllowRemoveStaleIntuneEnrollment=[bool]$AllowRemoveStaleIntuneEnrollment
                SkipVirtualMachines=[bool]$SkipVirtualMachines
                ScriptVersion=$ScriptVersion
                IsVirtualMachine=$isVirtualMachine
                VirtualMachineEvidence=$virtualMachineEvidence
                LeaveAttempted=$false; LeaveExitCode=""; Status=$status; ErrorMessage=$errorMessage
                Dsreg_AzureAdJoined=""; Dsreg_DeviceId=""; Dsreg_TenantName=""; Dsreg_TenantId=""
                DeviceAuthStatus=""; DsregStatusErrorMessage=""; IntuneEnrolled=$null
                ClientErrorCode=""; ServerErrorCode=""; ServerErrorSubCode=""; ServerOperation=""
                ServerMessage=""; HttpsStatus=""; RequestId=""; ErrorPhase=""
            }
            Write-AtomicCsvAppend -Path $logPath -RowObject $logEntry -RunIdValue $RunId
            Write-FinalStatusLine -Status $status -ExitCode 0 -Detail $errorMessage -NextAction (Get-NextActionForStatus -Status $status)
            exit 0
        }
    }
    else {
        Write-RunLog ("Computer system summary failed: {0}" -f $computerSystemSummary.Detail)
    }

    # Flush DNS before domain/DC checks so stale name resolution does not poison DC/MDM checks.
    $flushDnsAttempted = $true
    $flushDnsOutputFile = Join-Path $OutputDir ("{0}_flushdns_{1}.txt" -f $ComputerName, $RunId)
    Write-Host "Flushing DNS cache..." -ForegroundColor Cyan
    Write-RunLog "Running ipconfig /flushdns."
    try {
        $flushDnsOutput = ipconfig.exe /flushdns 2>&1
        $flushDnsExitCode = $LASTEXITCODE
        $flushDnsOutput | Out-File -FilePath $flushDnsOutputFile -Encoding UTF8 -Force
        Write-RunLog ("ipconfig /flushdns completed. ExitCode={0}; OutputFile={1}" -f $flushDnsExitCode,$flushDnsOutputFile)
    }
    catch {
        $flushDnsExitCode = ""
        Write-RunLog ("ipconfig /flushdns failed to start: {0}" -f $_.Exception.Message)
    }

    # Domain/DC preflight must run before gpupdate and before any repair action.
    Write-Host "Checking Active Directory domain join status..." -ForegroundColor Cyan
    Write-RunLog "Checking domain join status."

    if ($null -eq $cs) {
        Stop-ComputerSystemQueryFailed -Detail $computerSystemSummary.Detail
    }
    if (-not $cs.PartOfDomain) {
        $status = "NOT_DOMAIN_JOINED"
        $errorMessage = "Device is not joined to an Active Directory domain. Script stopped before gpupdate/repair."
        Write-Host $errorMessage -ForegroundColor Yellow
        Write-RunLog $errorMessage

        $logEntry = [PSCustomObject]@{
            RunId=$RunId; Timestamp=$Timestamp; ComputerName=$ComputerName; AllowDsregLeave=[bool]$AllowDsregLeave
            ScriptVersion=$ScriptVersion
            LeaveAttempted=$false; LeaveExitCode=""; Status=$status; ErrorMessage=$errorMessage
            Dsreg_AzureAdJoined=""; Dsreg_DeviceId=""; Dsreg_TenantName=""; Dsreg_TenantId=""
            DeviceAuthStatus=""; DsregStatusErrorMessage=""; IntuneEnrolled=$null
            ClientErrorCode=""; ServerErrorCode=""; ServerErrorSubCode=""; ServerOperation=""
            ServerMessage=""; HttpsStatus=""; RequestId=""; ErrorPhase=""
        }
        Write-AtomicCsvAppend -Path $logPath -RowObject $logEntry -RunIdValue $RunId
        Write-FinalStatusLine -Status $status -ExitCode 2 -Detail $errorMessage -NextAction (Get-NextActionForStatus -Status $status)
        exit 2
    }

    $domainName = [string]$cs.Domain
    Write-RunLog ("Domain joined: YES. Domain={0}" -f $domainName)

    Write-Host "Checking domain controller accessibility..." -ForegroundColor Cyan
    Write-RunLog "Checking domain controller accessibility."

    if ([string]::IsNullOrWhiteSpace($domainName)) {
        $status = "DOMAIN_NAME_EMPTY"
        $errorMessage = "Device reports PartOfDomain=TRUE but Domain is empty. Script stopped before gpupdate/repair."
        Write-Host $errorMessage -ForegroundColor Red
        Write-RunLog $errorMessage

        $logEntry = [PSCustomObject]@{
            RunId=$RunId; Timestamp=$Timestamp; ComputerName=$ComputerName; AllowDsregLeave=[bool]$AllowDsregLeave
            ScriptVersion=$ScriptVersion
            LeaveAttempted=$false; LeaveExitCode=""; Status=$status; ErrorMessage=$errorMessage
            Dsreg_AzureAdJoined=""; Dsreg_DeviceId=""; Dsreg_TenantName=""; Dsreg_TenantId=""
            DeviceAuthStatus=""; DsregStatusErrorMessage=""; IntuneEnrolled=$null
            ClientErrorCode=""; ServerErrorCode=""; ServerErrorSubCode=""; ServerOperation=""
            ServerMessage=""; HttpsStatus=""; RequestId=""; ErrorPhase=""
        }
        Write-AtomicCsvAppend -Path $logPath -RowObject $logEntry -RunIdValue $RunId
        Write-FinalStatusLine -Status $status -ExitCode 1 -Detail $errorMessage -NextAction (Get-NextActionForStatus -Status $status)
        exit 1
    }

    $dcTest = Test-DomainControllerReachable -DomainName $domainName
    Write-RunLog ("DC reachability: Reachable={0}; Method={1}; Detail={2}" -f $dcTest.Reachable,$dcTest.Method,$dcTest.Detail)

    if (-not $dcTest.Reachable) {
        $status = "DOMAIN_CONTROLLER_UNREACHABLE"
        $errorMessage = "Domain controller is not reachable for domain '$domainName'. Method=$($dcTest.Method). Detail=$($dcTest.Detail). Diagnostic only; gpupdate and repair actions were skipped."
        $nextAction = Get-NextActionForStatus -Status $status
        Write-Host $errorMessage -ForegroundColor Yellow
        Write-RunLog $errorMessage

        $logEntry = [PSCustomObject]@{
            RunId=$RunId; Timestamp=$Timestamp; ComputerName=$ComputerName; AllowDsregLeave=[bool]$AllowDsregLeave
            ScriptVersion=$ScriptVersion
            LeaveAttempted=$false; LeaveExitCode=""; Status=$status; ErrorMessage=$errorMessage
            Dsreg_AzureAdJoined=""; Dsreg_DeviceId=""; Dsreg_TenantName=""; Dsreg_TenantId=""
            DeviceAuthStatus=""; DsregStatusErrorMessage=""; IntuneEnrolled=$null
            ClientErrorCode=""; ServerErrorCode=""; ServerErrorSubCode=""; ServerOperation=""
            ServerMessage=""; HttpsStatus=""; RequestId=""; ErrorPhase=""
        }
        Write-AtomicCsvAppend -Path $logPath -RowObject $logEntry -RunIdValue $RunId
        Write-FinalStatusLine -Status $status -ExitCode 4 -Detail $errorMessage -NextAction $nextAction
        exit 4
    }

    $adComputerLocationChecked = $true
    $adLocation = Get-LocalComputerAdLocation -ComputerName $ComputerName
    $adComputerDistinguishedName = $adLocation.DistinguishedName
    $adComputerDefaultNamingContext = $adLocation.DefaultNamingContext
    $adComputerInDefaultComputersContainer = $adLocation.IsDefaultComputersContainer
    $adComputerLocationDetail = $adLocation.Detail
    Write-RunLog ("AD computer location: Checked={0}; Success={1}; InDefaultComputersContainer={2}; DistinguishedName={3}; Detail={4}" -f $adComputerLocationChecked,$adLocation.Success,$adComputerInDefaultComputersContainer,$adComputerDistinguishedName,$adComputerLocationDetail)

    $domainPreflightComplete = $true

    # Refresh computer policy before checking local MDM/GPO state.
    $gpUpdateAttempted = $true
    $gpUpdateOutputFile = Join-Path $OutputDir ("{0}_gpupdate_{1}.txt" -f $ComputerName, $RunId)
    Write-Host "Refreshing computer Group Policy..." -ForegroundColor Cyan
    Write-RunLog "Running gpupdate /target:computer /force."

    try {
        $gpupdateOutput = gpupdate.exe /target:computer /force 2>&1
        $gpUpdateExitCode = $LASTEXITCODE
        $gpupdateOutput | Out-File -FilePath $gpUpdateOutputFile -Encoding UTF8 -Force
        Write-RunLog ("gpupdate completed. ExitCode={0}; OutputFile={1}" -f $gpUpdateExitCode,$gpUpdateOutputFile)

        $gpupdateText = (($gpupdateOutput | Out-String).Trim())
        if ($gpupdateText -match "MDM Policy") {
            $gpUpdateMdmPolicyWarning = $true
            $gpUpdateMdmPolicyState = "Warning"
            $gpUpdateMdmPolicyDetail = "gpupdate reported MDM Policy processing issue."
            $gpResult = Export-ComputerGpResultEvidence -ComputerName $ComputerName -RunId $RunId -OutputDir $OutputDir -Reason "gpupdate reported an MDM Policy warning"
            $gpResultHtmlFile = $gpResult.HtmlFile
            $gpResultTextFile = $gpResult.TextFile

            $gpPolicyDiagnosticText = $gpupdateText
            try {
                if (Test-Path -LiteralPath $gpResultHtmlFile) {
                    $gpPolicyDiagnosticText += "`r`n" + (Get-Content -LiteralPath $gpResultHtmlFile -Raw -ErrorAction Stop)
                }
            }
            catch {
                Write-RunLog ("gpresult HTML read failed for MDM Policy classification: {0}" -f $_.Exception.Message)
            }

            try {
                if (Test-Path -LiteralPath $gpResultTextFile) {
                    $gpPolicyDiagnosticText += "`r`n" + (Get-Content -LiteralPath $gpResultTextFile -Raw -ErrorAction Stop)
                }
            }
            catch {
                Write-RunLog ("gpresult text read failed for MDM Policy classification: {0}" -f $_.Exception.Message)
            }

            try {
                $gpPolicyDiagnosticText = [System.Net.WebUtility]::HtmlDecode($gpPolicyDiagnosticText)
            }
            catch {
                Write-RunLog ("HTML decode failed for MDM Policy classification: {0}" -f $_.Exception.Message)
            }

            $gpPolicyDiagnosticFolded = $gpPolicyDiagnosticText `
                -replace ([string][char]0x00E0), "a" `
                -replace ([string][char]0x00E1), "a" `
                -replace ([string][char]0x00E2), "a" `
                -replace ([string][char]0x00E3), "a" `
                -replace ([string][char]0x00E4), "a" `
                -replace ([string][char]0x00E7), "c" `
                -replace ([string][char]0x00E8), "e" `
                -replace ([string][char]0x00E9), "e" `
                -replace ([string][char]0x00EA), "e" `
                -replace ([string][char]0x00EB), "e" `
                -replace ([string][char]0x00ED), "i" `
                -replace ([string][char]0x00EE), "i" `
                -replace ([string][char]0x00EF), "i" `
                -replace ([string][char]0x00F3), "o" `
                -replace ([string][char]0x00F4), "o" `
                -replace ([string][char]0x00F5), "o" `
                -replace ([string][char]0x00F6), "o" `
                -replace ([string][char]0x00FA), "u" `
                -replace ([string][char]0x00F9), "u" `
                -replace ([string][char]0x00FB), "u" `
                -replace ([string][char]0x00FC), "u"

            if (($gpPolicyDiagnosticFolded -match "(?i)deja\s+inscrit") -or
                ($gpPolicyDiagnosticFolded -match "(?i)already\s+(enrolled|registered)") -or
                ($gpPolicyDiagnosticFolded -match "(?i)\b(ja|ya)\s+esta\s+(inscrit|inscrito|registrad)") -or
                ($gpPolicyDiagnosticFolded -match "(?i)\bgia\s+(iscritto|registrat)") -or
                ($gpPolicyDiagnosticFolded -match "(?i)bereits\s+(registriert|angemeldet|eingeschrieben)") -or
                ($gpPolicyDiagnosticFolded -match "(?i)(al\s+ingeschreven|reeds\s+ingeschreven)") -or
                ($gpPolicyDiagnosticFolded -match "(?i)juz\s+(zarejestrowan|zapisany)")) {
                $gpUpdateMdmPolicyState = "AlreadyEnrolled"
                $gpUpdateMdmPolicyDetail = "The local Windows MDM policy client reports an existing enrollment. This is a weak local signal only; it does not prove confirmed Intune enrollment."
                Write-RunLog ("gpupdate MDM Policy classification: State={0}; Detail={1}" -f $gpUpdateMdmPolicyState,$gpUpdateMdmPolicyDetail)
            }
            else {
                Write-RunLog ("gpupdate MDM Policy classification: State={0}; Detail={1}" -f $gpUpdateMdmPolicyState,$gpUpdateMdmPolicyDetail)
            }
        }
    }
    catch {
        $gpUpdateExitCode = ""
        Write-RunLog ("gpupdate failed to start: {0}" -f $_.Exception.Message)
    }

    if ($GpUpdateWaitSeconds -gt 0) {
        Write-RunLog ("Waiting {0} seconds after gpupdate." -f $GpUpdateWaitSeconds)
        Start-Sleep -Seconds $GpUpdateWaitSeconds
    }

    if ($domainPreflightComplete) {
        Write-RunLog ("Domain/DC preflight already completed before gpupdate. Domain={0}; DCReachable={1}" -f $domainName,$dcTest.Reachable)
    }
    else {
    # Domain joined?
    Write-Host "Checking Active Directory domain join status..." -ForegroundColor Cyan
    Write-RunLog "Checking domain join status."

    $computerSystemSummary = Get-ComputerSystemSummary
    if (-not $computerSystemSummary.Success) {
        Stop-ComputerSystemQueryFailed -Detail $computerSystemSummary.Detail
    }
    $cs = $computerSystemSummary.ComputerSystem
    if (-not $cs.PartOfDomain) {
        $status = "NOT_DOMAIN_JOINED"
        $errorMessage = "Device is not joined to an Active Directory domain. Script stopped at start."
        Write-Host $errorMessage -ForegroundColor Yellow
        Write-RunLog $errorMessage

        $logEntry = [PSCustomObject]@{
            RunId=$RunId; Timestamp=$Timestamp; ComputerName=$ComputerName; AllowDsregLeave=[bool]$AllowDsregLeave
            ScriptVersion=$ScriptVersion
            LeaveAttempted=$false; LeaveExitCode=""; Status=$status; ErrorMessage=$errorMessage
            Dsreg_AzureAdJoined=""; Dsreg_DeviceId=""; Dsreg_TenantName=""; Dsreg_TenantId=""
            DeviceAuthStatus=""; DsregStatusErrorMessage=""; IntuneEnrolled=$null
            ClientErrorCode=""; ServerErrorCode=""; ServerErrorSubCode=""; ServerOperation=""
            ServerMessage=""; HttpsStatus=""; RequestId=""; ErrorPhase=""
        }
        Write-AtomicCsvAppend -Path $logPath -RowObject $logEntry -RunIdValue $RunId
        Write-FinalStatusLine -Status $status -ExitCode 2 -Detail $errorMessage
        exit 2
    }

    $domainName = [string]$cs.Domain
    Write-RunLog ("Domain joined: YES. Domain={0}" -f $domainName)

    # DC reachable?
    Write-Host "Checking domain controller accessibility..." -ForegroundColor Cyan
    Write-RunLog "Checking domain controller accessibility."

    if ([string]::IsNullOrWhiteSpace($domainName)) {
        $status = "DOMAIN_NAME_EMPTY"
        $errorMessage = "Device reports PartOfDomain=TRUE but Domain is empty. Script stopped."
        Write-Host $errorMessage -ForegroundColor Red
        Write-RunLog $errorMessage

        $logEntry = [PSCustomObject]@{
            RunId=$RunId; Timestamp=$Timestamp; ComputerName=$ComputerName; AllowDsregLeave=[bool]$AllowDsregLeave
            ScriptVersion=$ScriptVersion
            LeaveAttempted=$false; LeaveExitCode=""; Status=$status; ErrorMessage=$errorMessage
            Dsreg_AzureAdJoined=""; Dsreg_DeviceId=""; Dsreg_TenantName=""; Dsreg_TenantId=""
            DeviceAuthStatus=""; DsregStatusErrorMessage=""; IntuneEnrolled=$null
            ClientErrorCode=""; ServerErrorCode=""; ServerErrorSubCode=""; ServerOperation=""
            ServerMessage=""; HttpsStatus=""; RequestId=""; ErrorPhase=""
        }
        Write-AtomicCsvAppend -Path $logPath -RowObject $logEntry -RunIdValue $RunId
        Write-FinalStatusLine -Status $status -ExitCode 1 -Detail $errorMessage
        exit 1
    }

    $dcTest = Test-DomainControllerReachable -DomainName $domainName
    Write-RunLog ("DC reachability: Reachable={0}; Method={1}; Detail={2}" -f $dcTest.Reachable,$dcTest.Method,$dcTest.Detail)

    if (-not $dcTest.Reachable) {
        $status = "DOMAIN_CONTROLLER_UNREACHABLE"
        $errorMessage = "Domain controller is not reachable for domain '$domainName'. Method=$($dcTest.Method). Detail=$($dcTest.Detail). Diagnostic only; gpupdate and repair actions were skipped."
        Write-Host $errorMessage -ForegroundColor Yellow
        Write-RunLog $errorMessage

        $logEntry = [PSCustomObject]@{
            RunId=$RunId; Timestamp=$Timestamp; ComputerName=$ComputerName; AllowDsregLeave=[bool]$AllowDsregLeave
            ScriptVersion=$ScriptVersion
            LeaveAttempted=$false; LeaveExitCode=""; Status=$status; ErrorMessage=$errorMessage
            Dsreg_AzureAdJoined=""; Dsreg_DeviceId=""; Dsreg_TenantName=""; Dsreg_TenantId=""
            DeviceAuthStatus=""; DsregStatusErrorMessage=""; IntuneEnrolled=$null
            ClientErrorCode=""; ServerErrorCode=""; ServerErrorSubCode=""; ServerOperation=""
            ServerMessage=""; HttpsStatus=""; RequestId=""; ErrorPhase=""
        }
        Write-AtomicCsvAppend -Path $logPath -RowObject $logEntry -RunIdValue $RunId
        Write-FinalStatusLine -Status $status -ExitCode 4 -Detail $errorMessage -NextAction (Get-NextActionForStatus -Status $status)
        exit 4
    }
    }

    # Initial dsregcmd /status
    Write-Host "Retrieving initial dsregcmd /status locally..." -ForegroundColor Cyan
    Write-RunLog "Running initial dsregcmd /status."
    Write-RunLog "SYSTEM dsreg status is used for device/join fields only; user PRT/WAM fields from SYSTEM context are ignored."

    $parsed = Get-ParsedDsregStatusSnapshot `
        -DsregcmdPath $dsregcmdPath `
        -OutputDirPath $OutputDir `
        -RunIdValue $RunId `
        -ComputerNameValue $ComputerName `
        -PhaseLabel "initial"

    $dsregAzureAdJoined    = $parsed.AzureAdJoined
    $dsregDeviceId         = $parsed.DeviceId
    $dsregTenantName       = $parsed.TenantName
    $dsregTenantId         = $parsed.TenantId
    $dsregDeviceAuthStatus = $parsed.DeviceAuthStatusRaw
    $dsregKeySignTest      = $parsed.KeySignTest

    $clientErrorCode    = $parsed.ClientErrorCode
    $serverErrorCode    = $parsed.ServerErrorCode
    $serverErrorSubCode = $parsed.ServerErrorSubCode
    $serverOperation    = $parsed.ServerOperation
    $serverMessage      = $parsed.ServerMessage
    $httpsStatus        = $parsed.HttpsStatus
    $requestId          = $parsed.RequestId
    $errorPhase         = $parsed.ErrorPhase
    $mdmUrl             = $parsed.MdmUrl
    $mdmTouUrl          = $parsed.MdmTouUrl
    $mdmComplianceUrl   = $parsed.MdmComplianceUrl

    Write-RunLog ("Initial dsreg: AzureAdJoined={0}; DeviceId={1}; TenantId={2}; DeviceAuthStatus={3}; KeySignTest={4}; ServerErrorSubCode={5}; ErrorPhase={6}; ClientErrorCode={7}; MdmUrl={8}" -f $dsregAzureAdJoined,$dsregDeviceId,$dsregTenantId,$dsregDeviceAuthStatus,$dsregKeySignTest,$serverErrorSubCode,$errorPhase,$clientErrorCode,$mdmUrl)

    # Intune is checked after dsregcmd /status so healthy devices still return success.
    Write-Host "Checking confirmed local Intune (MDM) enrollment..." -ForegroundColor Cyan
    Write-RunLog "Checking confirmed Intune enrollment via Registry64 and EnterpriseMgmt evidence."

    $mdmEnrollmentState = Get-MdmEnrollmentState
    $anyMdmEnrollmentDetected = $mdmEnrollmentState.AnyMdmEnrollmentDetected
    $nonIntuneMdmEnrollmentDetected = $mdmEnrollmentState.NonIntuneMdmEnrollmentDetected
    $mdmEnrollmentCount = $mdmEnrollmentState.EnrollmentCount
    $mdmProviderIds = $mdmEnrollmentState.ProviderIds
    $intuneEnrollmentIds = $mdmEnrollmentState.IntuneEnrollmentIds
    $nonIntuneEnrollmentIds = $mdmEnrollmentState.NonIntuneEnrollmentIds
    $unconfirmedIntuneEnrollmentIds = $mdmEnrollmentState.UnconfirmedIntuneEnrollmentIds
    $mdmEnrollmentDetails = $mdmEnrollmentState.EnrollmentDetails
    $ignoredEnrollmentDetails = $mdmEnrollmentState.IgnoredEnrollmentDetails
    $staleIntuneEnrollmentDetected = (-not $mdmEnrollmentState.IntuneEnrollmentDetected) -and (-not [string]::IsNullOrWhiteSpace($unconfirmedIntuneEnrollmentIds)) -and ($gpUpdateMdmPolicyState -eq "AlreadyEnrolled")
    if ($staleIntuneEnrollmentDetected) { $staleIntuneEnrollmentIds = $unconfirmedIntuneEnrollmentIds }
    Write-RunLog ("MDM enrollment registry state: AnyMdm={0}; Intune={1}; NonIntune={2}; Count={3}; ProviderIds={4}; IntuneIds={5}; NonIntuneIds={6}; UnconfirmedIntuneIds={7}; Details={8}; Ignored={9}" -f $anyMdmEnrollmentDetected,$mdmEnrollmentState.IntuneEnrollmentDetected,$nonIntuneMdmEnrollmentDetected,$mdmEnrollmentCount,$mdmProviderIds,$intuneEnrollmentIds,$nonIntuneEnrollmentIds,$unconfirmedIntuneEnrollmentIds,$mdmEnrollmentDetails,$ignoredEnrollmentDetails)
    if ($staleIntuneEnrollmentDetected) {
        Write-RunLog ("Stale local Intune enrollment suspected. UnconfirmedIntuneIds={0}; GpUpdateMdmPolicyState={1}" -f $staleIntuneEnrollmentIds,$gpUpdateMdmPolicyState)
    }

    if ($mdmEnrollmentState.IntuneEnrollmentDetected) {
        $intuneEnrolled = $true
        Write-RunLog "Confirmed local Intune enrollment detected. dsregcmd repair operations will be skipped if the device is not already healthy."
    }
    else {
        $intuneEnrolled = $false
        Write-RunLog "Confirmed local Intune enrollment not detected."
        $intuneEnrollmentDiagnosticsInitialFiles = Export-EnrollmentDiagnostics `
            -Since $Timestamp `
            -OutputDirPath $OutputDir `
            -ComputerNameValue $ComputerName `
            -RunIdValue ("{0}_initial" -f $RunId)
        Write-RunLog ("Initial enrollment diagnostics exported because Intune is missing: {0}" -f $intuneEnrollmentDiagnosticsInitialFiles)
    }

    $interactiveState = Get-InteractiveUserSessionState
    $interactiveUserDetected = $interactiveState.HasInteractiveUser
    $interactiveUserName = $interactiveState.ActiveUser
    $interactiveSessionName = $interactiveState.ActiveSessionName
    $interactiveSessionId = $interactiveState.ActiveSessionId
    $interactiveSessionState = $interactiveState.ActiveState
    $interactiveSessionDetail = $interactiveState.Detail
    $interactiveSessionIsRemote = ($interactiveSessionName -match "^(?i:rdp|rdp-tcp)")
    if ($interactiveUserDetected) {
        $interactiveIdentity = Resolve-InteractiveUserAccount -SessionId $interactiveSessionId -UserName $interactiveUserName
        $interactiveUserIdentityResolved = $interactiveIdentity.Resolved
        $interactiveUserDomain = $interactiveIdentity.Domain
        $interactiveUserAccountName = $interactiveIdentity.AccountName
        $interactiveUserAccountType = $interactiveIdentity.AccountType
        $interactiveUserIdentityDetail = $interactiveIdentity.Detail
    }

    Write-RunLog ("Interactive user detection: HasUser={0}; User={1}; Domain={2}; AccountName={3}; AccountType={4}; IdentityResolved={5}; SessionName={6}; SessionId={7}; State={8}; IsRemote={9}; Detail={10}; IdentityDetail={11}" -f $interactiveUserDetected,$interactiveUserName,$interactiveUserDomain,$interactiveUserAccountName,$interactiveUserAccountType,$interactiveUserIdentityResolved,$interactiveSessionName,$interactiveSessionId,$interactiveSessionState,$interactiveSessionIsRemote,$interactiveSessionDetail,$interactiveUserIdentityDetail)

    if ($interactiveUserDetected -and $interactiveUserAccountType -ne "Local") {
        # User-context status is collected through SmartM365 on-demand scheduled tasks.
        $helperTaskResult = Ensure-UserContextHelperTasks
        if (-not $helperTaskResult.Success) {
            Write-RunLog ("User-context helper tasks were not fully registered. Detail={0}" -f $helperTaskResult.Detail)
        }

        $userStatusAttempted = $true
        $userStatus = Invoke-UserDsregStatusTaskSnapshot `
            -TaskName $UserStatusTaskName `
            -OutputDirPath $OutputDir `
            -RunIdValue $RunId `
            -ComputerNameValue $ComputerName `
            -PhaseLabel "initial" `
            -WaitSeconds $UserTaskWaitSeconds `
            -ExpectedUserName $interactiveUserName

        $userStatusTaskExitCode = $userStatus.TaskExitCode
        $userStatusFile = $userStatus.CopiedFile
        $userStatusErrorMessage = $userStatus.ErrorMessage
        $userAzureAdPrt = $userStatus.Parsed.AzureAdPrt
        $userAzureAdPrtAuthority = $userStatus.Parsed.AzureAdPrtAuthority
        $userAzureAdPrtUpdateTime = $userStatus.Parsed.AzureAdPrtUpdateTime
        $userAzureAdPrtExpiryTime = $userStatus.Parsed.AzureAdPrtExpiryTime
        $userRefreshPrtAttemptStatus = $userStatus.Parsed.RefreshPrtAttemptStatus
        $userRefreshPrtHttpStatus = $userStatus.Parsed.RefreshPrtHttpStatus
        $userRefreshPrtHttpError = $userStatus.Parsed.RefreshPrtHttpError
        $userRefreshPrtServerErrorCode = $userStatus.Parsed.RefreshPrtServerErrorCode
        $userRefreshPrtServerErrorSubCode = $userStatus.Parsed.RefreshPrtServerErrorSubCode
        $userEnterprisePrt = $userStatus.Parsed.EnterprisePrt
        $userOnPremTgt = $userStatus.Parsed.OnPremTgt
        $userCloudTgt = $userStatus.Parsed.CloudTgt
        $userWorkplaceJoined = $userStatus.Parsed.WorkplaceJoined
        $userWamDefaultSet = $userStatus.Parsed.WamDefaultSet
        $userWamDefaultAuthority = $userStatus.Parsed.WamDefaultAuthority
        $userWamDefaultId = $userStatus.Parsed.WamDefaultId
        $userNgcSet = $userStatus.Parsed.NgcSet
        $userIsUserAzureAD = $userStatus.Parsed.IsUserAzureAD
        $userSessionIsNotRemote = $userStatus.Parsed.SessionIsNotRemote

        $userPrtRefreshDecision = Test-UserPrtRefreshNeeded -ParsedUserDsreg $userStatus.Parsed
        $userRefreshPrtReason = $userPrtRefreshDecision.Reason
        $userPrtRefreshStillNeeded = [bool]$userPrtRefreshDecision.Needed

        Write-RunLog ("User dsreg initial: Success={0}; IsUserAzureAD={1}; SessionIsNotRemote={2}; AzureAdPrt={3}; AzureAdPrtExpiryTime={4}; RefreshPrtAttemptStatus={5}; RefreshPrtHttpStatus={6}; RefreshPrtHttpError={7}; EnterprisePrt={8}; OnPremTgt={9}; CloudTgt={10}; WamDefaultSet={11}; WorkplaceJoined={12}; RefreshNeeded={13}; RefreshReason={14}; Error={15}" -f $userStatus.Success,$userIsUserAzureAD,$userSessionIsNotRemote,$userAzureAdPrt,$userAzureAdPrtExpiryTime,$userRefreshPrtAttemptStatus,$userRefreshPrtHttpStatus,$userRefreshPrtHttpError,$userEnterprisePrt,$userOnPremTgt,$userCloudTgt,$userWamDefaultSet,$userWorkplaceJoined,$userPrtRefreshDecision.Needed,$userRefreshPrtReason,$userStatusErrorMessage)

        if ($userStatus.Success -and $userPrtRefreshDecision.Needed -and -not $AuditOnly) {
            Write-RunLog ("Running user-context refresh PRT scheduled task because: {0}" -f $userRefreshPrtReason)
            $userRefreshPrtAttempted = $true
            $refreshResult = Invoke-UserRefreshPrtTask -TaskName $UserRefreshPrtTaskName -WaitSeconds $UserPrtRetryWaitSeconds -ExpectedUserName $interactiveUserName
            $userRefreshPrtExitCode = $refreshResult.ExitCode
            $userRefreshPrtFile = $refreshResult.SourceFile

            if (-not [string]::IsNullOrWhiteSpace($userRefreshPrtFile) -and (Test-Path $userRefreshPrtFile)) {
                $refreshCopy = Join-Path $OutputDir ("{0}_dsreg_user_refreshprt_{1}_{2}.txt" -f $ComputerName, $RunId, (Split-Path $userRefreshPrtFile -Leaf))
                Copy-Item -LiteralPath $userRefreshPrtFile -Destination $refreshCopy -Force
                $userRefreshPrtFile = $refreshCopy
            }

            $userStatusAfterRefresh = Invoke-UserDsregStatusTaskSnapshot `
                -TaskName $UserStatusTaskName `
                -OutputDirPath $OutputDir `
                -RunIdValue $RunId `
                -ComputerNameValue $ComputerName `
                -PhaseLabel "post_refreshprt" `
                -WaitSeconds $UserTaskWaitSeconds `
                -ExpectedUserName $interactiveUserName

            $userStatusTaskExitCode = $userStatusAfterRefresh.TaskExitCode
            $userStatusFile = $userStatusAfterRefresh.CopiedFile
            $userStatusErrorMessage = $userStatusAfterRefresh.ErrorMessage
            $userAzureAdPrt = $userStatusAfterRefresh.Parsed.AzureAdPrt
            $userAzureAdPrtAuthority = $userStatusAfterRefresh.Parsed.AzureAdPrtAuthority
            $userAzureAdPrtUpdateTime = $userStatusAfterRefresh.Parsed.AzureAdPrtUpdateTime
            $userAzureAdPrtExpiryTime = $userStatusAfterRefresh.Parsed.AzureAdPrtExpiryTime
            $userRefreshPrtAttemptStatus = $userStatusAfterRefresh.Parsed.RefreshPrtAttemptStatus
            $userRefreshPrtHttpStatus = $userStatusAfterRefresh.Parsed.RefreshPrtHttpStatus
            $userRefreshPrtHttpError = $userStatusAfterRefresh.Parsed.RefreshPrtHttpError
            $userRefreshPrtServerErrorCode = $userStatusAfterRefresh.Parsed.RefreshPrtServerErrorCode
            $userRefreshPrtServerErrorSubCode = $userStatusAfterRefresh.Parsed.RefreshPrtServerErrorSubCode
            $userEnterprisePrt = $userStatusAfterRefresh.Parsed.EnterprisePrt
            $userOnPremTgt = $userStatusAfterRefresh.Parsed.OnPremTgt
            $userCloudTgt = $userStatusAfterRefresh.Parsed.CloudTgt
            $userWorkplaceJoined = $userStatusAfterRefresh.Parsed.WorkplaceJoined
            $userWamDefaultSet = $userStatusAfterRefresh.Parsed.WamDefaultSet
            $userWamDefaultAuthority = $userStatusAfterRefresh.Parsed.WamDefaultAuthority
            $userWamDefaultId = $userStatusAfterRefresh.Parsed.WamDefaultId
            $userNgcSet = $userStatusAfterRefresh.Parsed.NgcSet
            $userIsUserAzureAD = $userStatusAfterRefresh.Parsed.IsUserAzureAD
            $userSessionIsNotRemote = $userStatusAfterRefresh.Parsed.SessionIsNotRemote

            $userPrtRefreshDecisionAfter = Test-UserPrtRefreshNeeded -ParsedUserDsreg $userStatusAfterRefresh.Parsed
            $userRefreshPrtReason = $userPrtRefreshDecisionAfter.Reason
            $userPrtRefreshStillNeeded = [bool]$userPrtRefreshDecisionAfter.Needed

            Write-RunLog ("User dsreg after refreshprt: Success={0}; IsUserAzureAD={1}; SessionIsNotRemote={2}; AzureAdPrt={3}; AzureAdPrtExpiryTime={4}; RefreshPrtAttemptStatus={5}; RefreshPrtHttpStatus={6}; RefreshPrtHttpError={7}; EnterprisePrt={8}; OnPremTgt={9}; CloudTgt={10}; WamDefaultSet={11}; WorkplaceJoined={12}; RefreshStillNeeded={13}; RefreshReason={14}; Error={15}" -f $userStatusAfterRefresh.Success,$userIsUserAzureAD,$userSessionIsNotRemote,$userAzureAdPrt,$userAzureAdPrtExpiryTime,$userRefreshPrtAttemptStatus,$userRefreshPrtHttpStatus,$userRefreshPrtHttpError,$userEnterprisePrt,$userOnPremTgt,$userCloudTgt,$userWamDefaultSet,$userWorkplaceJoined,$userPrtRefreshDecisionAfter.Needed,$userRefreshPrtReason,$userStatusErrorMessage)
        }
    }
    elseif ($interactiveUserDetected -and $interactiveUserAccountType -eq "Local") {
        Write-RunLog ("Interactive user is a local account ({0}). Skipping user-context dsreg status, refreshprt, and User Credential auto-enrollment tasks for this session." -f $interactiveUserAccountName)
    }
    else {
        Write-RunLog "No active interactive user detected. Skipping user-context dsreg status and refreshprt tasks."
    }

    $deviceJoinHealthy = Test-DsregDeviceHealthy -AzureAdJoined $dsregAzureAdJoined -DeviceAuthStatus $dsregDeviceAuthStatus -KeySignTest $dsregKeySignTest

    if ($AuditOnly) {
        if ($deviceJoinHealthy -and $intuneEnrolled) {
            $status = "AUDIT_SUCCESS_ALREADY_INTUNE"
            $ExitCode = 0
        }
        elseif ($deviceJoinHealthy -and -not $intuneEnrolled) {
            if ($staleIntuneEnrollmentDetected) { $status = "AUDIT_STALE_INTUNE_ENROLLMENT_LOCAL" }
            else { $status = "AUDIT_INTUNE_MISSING" }
            $ExitCode = 3
        }
        else {
            $status = "AUDIT_HYBRID_JOIN_UNHEALTHY"
            $ExitCode = 3
        }

        $dsregStatusErrorMessage = "AuditOnly enabled. Diagnostics collected; no repair, reboot, leave, removal, retry, refreshprt, or auto-enrollment trigger was performed."
        Write-RunLog $dsregStatusErrorMessage
    }
    elseif ($deviceJoinHealthy -and $EntraHybridPending) {
        $status = "ENTRA_HYBRID_PENDING_ADJ_TRIGGERED"
        $dsregStatusErrorMessage = "Admin-side Entra inventory reports this hybrid object as pending (TrustType=ServerAd with empty AlternativeSecurityIds). Local device join is healthy, so dsregcmd /leave is not executed; Automatic-Device-Join is triggered and rechecked."
        $ExitCode = 3

        Write-Host "Entra hybrid object is pending, but local join is healthy. Triggering Automatic-Device-Join without dsregcmd /leave..." -ForegroundColor Yellow
        Write-RunLog $dsregStatusErrorMessage

        $taskOk = Start-AutoDeviceJoinTask
        Write-RunLog ("Entra pending: Automatic-Device-Join task run attempted. Success={0}" -f $taskOk)

        $rPending = Invoke-BoundedRetryUntilSuccess `
            -DsregcmdPath $dsregcmdPath `
            -OutputDirPath $OutputDir `
            -RunIdValue $RunId `
            -ComputerNameValue $ComputerName `
            -SleepMinutes $RetrySleepMinutes `
            -MaxRetries $RetryMaxRetries `
            -ContextLabel "entra_pending"

        if ($rPending.Success) {
            $status = "SUCCESS"
            $ExitCode = 0
            Write-RunLog ("Entra pending local retry sees healthy local join after {0} attempt(s). Continuing normal Intune repair checks." -f $rPending.Attempts)

            $dsregAzureAdJoined    = $rPending.Parsed.AzureAdJoined
            $dsregDeviceId         = $rPending.Parsed.DeviceId
            $dsregTenantName       = $rPending.Parsed.TenantName
            $dsregTenantId         = $rPending.Parsed.TenantId
            $dsregDeviceAuthStatus = $rPending.Parsed.DeviceAuthStatusRaw
            $dsregKeySignTest      = $rPending.Parsed.KeySignTest

            $clientErrorCode    = $rPending.Parsed.ClientErrorCode
            $serverErrorCode    = $rPending.Parsed.ServerErrorCode
            $serverErrorSubCode = $rPending.Parsed.ServerErrorSubCode
            $serverOperation    = $rPending.Parsed.ServerOperation
            $serverMessage      = $rPending.Parsed.ServerMessage
            $httpsStatus        = $rPending.Parsed.HttpsStatus
            $requestId          = $rPending.Parsed.RequestId
            $errorPhase         = $rPending.Parsed.ErrorPhase
            $mdmUrl             = $rPending.Parsed.MdmUrl
            $mdmTouUrl          = $rPending.Parsed.MdmTouUrl
            $mdmComplianceUrl   = $rPending.Parsed.MdmComplianceUrl
        }
        else {
            $status = "ENTRA_HYBRID_PENDING_RETRY_EXHAUSTED"
            $dsregStatusErrorMessage = ("Entra hybrid object is pending and Automatic-Device-Join was triggered, but local recheck did not confirm a completed join after {0} retries with {1} minutes sleep. Re-export DevicesEntra.csv after directory sync delay, then recheck for duplicate/pending Entra device object." -f $RetryMaxRetries, $RetrySleepMinutes)
            $ExitCode = 3
            Write-RunLog $dsregStatusErrorMessage
        }
    }
    # SUCCESS short-circuit
    elseif ($deviceJoinHealthy) {
        $status = "SUCCESS"
        $ExitCode = 0
        Write-RunLog ("Device join health is OK. AzureAdJoined={0}; DeviceAuthStatus={1}; KeySignTest={2}" -f $dsregAzureAdJoined,$dsregDeviceAuthStatus,$dsregKeySignTest)
    }
    else {
        if ($intuneEnrolled) {
            $status = "INTUNE_ENROLLED"
            $dsregStatusErrorMessage = "MDM/Intune enrolled (ProviderID='MS DM Server'). dsregcmd repair actions skipped."
            $ExitCode = 3

            Write-Host "MDM/Intune detected. Skipping dsregcmd repair actions." -ForegroundColor Green
            Write-RunLog "Intune enrollment detected and device is not already healthy. Skipping dsregcmd repair operations."
        }
        else {
        # Missing device hint (AAD Connect timing / device object missing)
        $missingDeviceHint =
            ($serverErrorSubCode -eq "error_missing_device") -or
            ($serverMessage -match "is not found") -or
            ($clientErrorCode -eq "0x801c03f3") -or
            ($serverOperation -eq "DeviceRenew")

        if (($dsregAzureAdJoined -eq "NO") -and $missingDeviceHint) {
            $status = "WAITING_FOR_AAD_CONNECT_LOCAL_RETRY"
            $dsregStatusErrorMessage = "Join indicates missing device / AAD Connect timing. Performing bounded local retry."
            $ExitCode = 3

            Write-RunLog "Missing device hint detected. Starting bounded local retry loop (no scheduled tasks)."
            Write-RunLog ("Missing device hint details: ServerErrorSubCode={0}; ClientErrorCode={1}; ServerOperation={2}; ServerMessage={3}" -f $serverErrorSubCode,$clientErrorCode,$serverOperation,$serverMessage)

            $r = Invoke-BoundedRetryUntilSuccess `
                -DsregcmdPath $dsregcmdPath `
                -OutputDirPath $OutputDir `
                -RunIdValue $RunId `
                -ComputerNameValue $ComputerName `
                -SleepMinutes $RetrySleepMinutes `
                -MaxRetries $RetryMaxRetries `
                -ContextLabel "missing_device"

            if ($r.Success) {
                $status = "SUCCESS"
                $ExitCode = 0
                Write-RunLog ("Resolved during retry loop after {0} attempt(s)." -f $r.Attempts)

                $dsregAzureAdJoined    = $r.Parsed.AzureAdJoined
                $dsregDeviceId         = $r.Parsed.DeviceId
                $dsregTenantName       = $r.Parsed.TenantName
                $dsregTenantId         = $r.Parsed.TenantId
                $dsregDeviceAuthStatus = $r.Parsed.DeviceAuthStatusRaw
                $dsregKeySignTest      = $r.Parsed.KeySignTest

                $clientErrorCode    = $r.Parsed.ClientErrorCode
                $serverErrorCode    = $r.Parsed.ServerErrorCode
                $serverErrorSubCode = $r.Parsed.ServerErrorSubCode
                $serverOperation    = $r.Parsed.ServerOperation
                $serverMessage      = $r.Parsed.ServerMessage
                $httpsStatus        = $r.Parsed.HttpsStatus
                $requestId          = $r.Parsed.RequestId
                $errorPhase         = $r.Parsed.ErrorPhase
                $mdmUrl             = $r.Parsed.MdmUrl
                $mdmTouUrl          = $r.Parsed.MdmTouUrl
                $mdmComplianceUrl   = $r.Parsed.MdmComplianceUrl
            }
            else {
                $status = "WAITING_FOR_AAD_CONNECT_LOCAL_RETRY_EXHAUSTED"
                $dsregStatusErrorMessage = ("Missing device hint persists after {0} retries with {1} minutes sleep." -f $RetryMaxRetries, $RetrySleepMinutes)
                $ExitCode = 3
                Write-RunLog $dsregStatusErrorMessage
            }
        }
        else {
            # Decide /leave (guardrails)
            $deviceAuthFailed =
                ($dsregDeviceAuthStatus -match "^\s*FAILED\b")
            $keySignFailedHybrid =
                ($dsregAzureAdJoined -eq "YES") -and
                ($dsregKeySignTest -eq "FAILED")
            $leaveReason = ""
            if ($deviceAuthFailed) {
                $leaveReason = ("DeviceAuthStatus is FAILED. DeviceAuthStatus={0}" -f $dsregDeviceAuthStatus)
            }
            elseif ($keySignFailedHybrid) {
                $leaveReason = ("KeySignTest is FAILED on an AzureAdJoined device. AzureAdJoined={0}; KeySignTest={1}; DeviceAuthStatus={2}" -f $dsregAzureAdJoined,$dsregKeySignTest,$dsregDeviceAuthStatus)
            }

            if ($AllowDsregLeave -and ($deviceAuthFailed -or $keySignFailedHybrid)) {
                Write-Host "Hybrid Join device key/authentication is unhealthy. Executing dsregcmd /leave..." -ForegroundColor Yellow
                Write-RunLog ("Executing dsregcmd /leave because {0}" -f $leaveReason)

                $leaveAttempted = $true
                $leaveOut = & $dsregcmdPath /leave 2>&1
                $leaveExitCode = $LASTEXITCODE

                $leaveCopy = Join-Path $OutputDir ("{0}_dsreg_leave_{1}.txt" -f $ComputerName, $RunId)
                $leaveOut | Out-File -FilePath $leaveCopy -Encoding UTF8 -Force

                if ($leaveExitCode -ne 0) {
                    $status = "LEAVE_EXITCODE_{0}" -f $leaveExitCode
                    throw "dsregcmd /leave returned exit code $leaveExitCode."
                }

                $status = "LEAVE_SUCCESS"
                Write-RunLog "Leave succeeded (exit code 0)."

                # Trigger Automatic-Device-Join
                Write-Host "Triggering Automatic-Device-Join task..." -ForegroundColor Cyan
                $taskOk = Start-AutoDeviceJoinTask
                Write-RunLog ("Automatic-Device-Join task run attempted. Success={0}" -f $taskOk)

                # Bounded local retry after leave
                $status = "WAITING_POST_LEAVE_LOCAL_RETRY"
                $dsregStatusErrorMessage = "Post-leave: Automatic-Device-Join triggered. Performing bounded local retry."
                $ExitCode = 3

                Write-RunLog "Post-leave bounded local retry loop starting (no scheduled tasks)."

                $r2 = Invoke-BoundedRetryUntilSuccess `
                    -DsregcmdPath $dsregcmdPath `
                    -OutputDirPath $OutputDir `
                    -RunIdValue $RunId `
                    -ComputerNameValue $ComputerName `
                    -SleepMinutes $RetrySleepMinutes `
                    -MaxRetries $RetryMaxRetries `
                    -ContextLabel "post_leave"

                if ($r2.Success) {
                    $status = "SUCCESS"
                    $ExitCode = 0
                    Write-RunLog ("Resolved post-leave during retry loop after {0} attempt(s)." -f $r2.Attempts)

                    $dsregAzureAdJoined    = $r2.Parsed.AzureAdJoined
                    $dsregDeviceId         = $r2.Parsed.DeviceId
                    $dsregTenantName       = $r2.Parsed.TenantName
                    $dsregTenantId         = $r2.Parsed.TenantId
                    $dsregDeviceAuthStatus = $r2.Parsed.DeviceAuthStatusRaw
                    $dsregKeySignTest      = $r2.Parsed.KeySignTest

                    $clientErrorCode    = $r2.Parsed.ClientErrorCode
                    $serverErrorCode    = $r2.Parsed.ServerErrorCode
                    $serverErrorSubCode = $r2.Parsed.ServerErrorSubCode
                    $serverOperation    = $r2.Parsed.ServerOperation
                    $serverMessage      = $r2.Parsed.ServerMessage
                    $httpsStatus        = $r2.Parsed.HttpsStatus
                    $requestId          = $r2.Parsed.RequestId
                    $errorPhase         = $r2.Parsed.ErrorPhase
                    $mdmUrl             = $r2.Parsed.MdmUrl
                    $mdmTouUrl          = $r2.Parsed.MdmTouUrl
                    $mdmComplianceUrl   = $r2.Parsed.MdmComplianceUrl
                }
                else {
                    $status = "WAITING_POST_LEAVE_LOCAL_RETRY_EXHAUSTED"
                    $dsregStatusErrorMessage = ("Post-leave join not completed after {0} retries with {1} minutes sleep." -f $RetryMaxRetries, $RetrySleepMinutes)
                    $ExitCode = 3
                    Write-RunLog $dsregStatusErrorMessage

                    if ($AllowRebootAfterDsregLeave) {
                        $rebootReason = "Post-leave Hybrid Join retry exhausted."
                        $rebootSafetyDecision = Get-RebootSafetyDecision -ReasonKey "POST_DSREG_LEAVE" -MaximumReboots $RetryAfterRebootMaxAttempts
                        if (-not $rebootSafetyDecision.Allowed) {
                            $status = "REBOOT_SAFETY_LIMIT_REACHED_POST_DSREG_LEAVE"
                            $dsregStatusErrorMessage = $rebootSafetyDecision.Detail
                            $ExitCode = 3
                            Write-RunLog $dsregStatusErrorMessage
                        }
                        else {
                            try {
                                Register-RetryAfterRebootTask -Reason $rebootReason
                            }
                            catch {
                                $script:RetryAfterRebootAction = "ScheduleFailed"
                                $script:RetryAfterRebootTaskNameResult = $script:RetryAfterRebootTaskName
                                $script:RetryAfterRebootDetail = $_.Exception.Message
                                $status = "RETRY_AFTER_REBOOT_SCHEDULE_FAILED_POST_DSREG_LEAVE"
                                $dsregStatusErrorMessage = "Post-leave reboot was not triggered because the retry-after-reboot task could not be scheduled. Error=$($_.Exception.Message)"
                                $ExitCode = 3
                                Write-RunLog $dsregStatusErrorMessage
                            }

                            if ($script:RetryAfterRebootAction -eq "Scheduled") {
                                Register-RebootSafetyEvent -ReasonKey "POST_DSREG_LEAVE" -Reason $rebootReason
                                $rebootAttempted = $true
                                $rebootResult = Start-ControlledReboot -Reason $rebootReason -DelaySeconds $RebootDelaySeconds
                                $rebootExitCode = $rebootResult.ExitCode
                                if ($rebootExitCode -eq 0) {
                                    $status = "REBOOT_TRIGGERED_POST_DSREG_LEAVE"
                                }
                                else {
                                    $status = "REBOOT_SCHEDULE_FAILED_POST_DSREG_LEAVE"
                                    $dsregStatusErrorMessage = "shutdown.exe failed with exit code $rebootExitCode. The retry-after-reboot task will be removed."
                                    $ExitCode = 3
                                }
                            }
                        }
                    }
                }
            }
            else {
                if ($keySignFailedHybrid) {
                    $status = "KEY_SIGN_TEST_FAILED"
                    $dsregStatusErrorMessage = ("KeySignTest=FAILED but dsregcmd /leave was not executed. AllowDsregLeave={0}; AzureAdJoined={1}; DeviceAuthStatus={2}; ErrorPhase={3}; ClientErrorCode={4}; ServerErrorSubCode={5}" -f [bool]$AllowDsregLeave,$dsregAzureAdJoined,$dsregDeviceAuthStatus,$errorPhase,$clientErrorCode,$serverErrorSubCode)
                }
                elseif ($AllowDsregLeave) {
                    $status = "LEAVE_NOT_APPLICABLE"
                    $dsregStatusErrorMessage = ("dsregcmd /leave not applicable. AllowDsregLeave={0}; AzureAdJoined={1}; DeviceAuthStatus={2}; KeySignTest={3}; ErrorPhase={4}; ClientErrorCode={5}; ServerErrorSubCode={6}" -f [bool]$AllowDsregLeave,$dsregAzureAdJoined,$dsregDeviceAuthStatus,$dsregKeySignTest,$errorPhase,$clientErrorCode,$serverErrorSubCode)
                }
                else {
                    $status = "SKIPPED"
                    $dsregStatusErrorMessage = ("No dsregcmd /leave requested. AllowDsregLeave={0}; AzureAdJoined={1}; DeviceAuthStatus={2}; KeySignTest={3}; ErrorPhase={4}; ClientErrorCode={5}; ServerErrorSubCode={6}" -f [bool]$AllowDsregLeave,$dsregAzureAdJoined,$dsregDeviceAuthStatus,$dsregKeySignTest,$errorPhase,$clientErrorCode,$serverErrorSubCode)
                }
                $ExitCode = 3
                Write-RunLog ("No dsreg leave performed. Status={0}; AllowDsregLeave={1}; AzureAdJoined={2}; DeviceAuthStatus={3}; KeySignTest={4}; Detail={5}" -f $status,[bool]$AllowDsregLeave,$dsregAzureAdJoined,$dsregDeviceAuthStatus,$dsregKeySignTest,$dsregStatusErrorMessage)
            }
        }
        }
    }

    # If Hybrid Join is healthy but Intune enrollment is missing, try to restore MDM enrollment.
    if ($status -eq "SUCCESS" -and $intuneEnrolled -eq $false) {
        $staleCleanupCompleted = $false
        if ($staleIntuneEnrollmentDetected) {
            if ($AllowRemoveStaleIntuneEnrollment) {
                $staleIntuneRemovalAttempted = $true
                Write-Host "Hybrid Join is healthy, but a stale local Intune enrollment trace exists. Removing it because -AllowRemoveStaleIntuneEnrollment was specified." -ForegroundColor Yellow
                Write-RunLog ("Stale Intune enrollment removal requested. EnrollmentIds={0}" -f $staleIntuneEnrollmentIds)

                $removeStaleResult = Remove-NonIntuneMdmEnrollment `
                    -EnrollmentIds $staleIntuneEnrollmentIds `
                    -OutputDirPath $OutputDir `
                    -ComputerNameValue $ComputerName `
                    -RunIdValue $RunId `
                    -RemovalLabel "Stale local Intune enrollment trace"

                $staleIntuneRemovalSuccess = $removeStaleResult.Success
                $staleIntuneRemovalBackupDir = $removeStaleResult.BackupDir
                $staleIntuneRemovalDetail = $removeStaleResult.Detail
                Write-RunLog ("Stale Intune enrollment removal result: Success={0}; BackupDir={1}; RemovedItems={2}; Detail={3}" -f $removeStaleResult.Success,$removeStaleResult.BackupDir,$removeStaleResult.RemovedItems,$removeStaleResult.Detail)

                if ($removeStaleResult.Success) {
                    $status = "SUCCESS"
                    $staleCleanupCompleted = $true
                    $staleIntuneCleanupContinuedToAutoEnroll = $true
                    $staleIntuneEnrollmentDetected = $false
                    $staleIntuneEnrollmentIds = ""
                    $unconfirmedIntuneEnrollmentIds = ""
                    $dsregStatusErrorMessage = ("Stale local Intune enrollment trace was removed. Waiting {0} seconds before attempting Intune auto-enrollment in the same run." -f $StaleCleanupDelaySeconds)
                    Write-RunLog $dsregStatusErrorMessage
                    if ($StaleCleanupDelaySeconds -gt 0) {
                        Start-Sleep -Seconds $StaleCleanupDelaySeconds
                    }
                }
                else {
                    $status = "STALE_INTUNE_ENROLLMENT_REMOVE_FAILED"
                    $dsregStatusErrorMessage = ("Stale local Intune enrollment removal did not fully complete. BackupDir={0}; Detail={1}" -f $removeStaleResult.BackupDir,$removeStaleResult.Detail)
                }
                $ExitCode = 3
                if (-not $removeStaleResult.Success) {
                    Write-RunLog $dsregStatusErrorMessage
                }
            }
            else {
                $status = "STALE_INTUNE_ENROLLMENT_LOCAL"
                $dsregStatusErrorMessage = ("A stale local Intune enrollment trace is suspected. UnconfirmedIntuneIds={0}; gpupdate State={1}. Use -AllowRemoveStaleIntuneEnrollment only after validation." -f $staleIntuneEnrollmentIds,$gpUpdateMdmPolicyState)
                $ExitCode = 3
                Write-Host "Hybrid Join is healthy, but a stale local Intune enrollment trace may block auto-enrollment." -ForegroundColor Yellow
                Write-RunLog $dsregStatusErrorMessage
            }
        }
        if ($status -eq "SUCCESS" -and $intuneEnrolled -eq $false -and -not $staleCleanupCompleted -and $nonIntuneMdmEnrollmentDetected) {
            if ($AllowRemoveNonIntuneMdmEnrollment) {
                $nonIntuneMdmRemovalAttempted = $true
                Write-Host "Hybrid Join is healthy, but a non-Intune MDM enrollment already exists. Removing it because -AllowRemoveNonIntuneMdmEnrollment was specified." -ForegroundColor Yellow
                Write-RunLog ("Non-Intune MDM removal requested. ProviderIds={0}; EnrollmentIds={1}" -f $mdmProviderIds,$nonIntuneEnrollmentIds)

                $removeMdmResult = Remove-NonIntuneMdmEnrollment `
                    -EnrollmentIds $nonIntuneEnrollmentIds `
                    -OutputDirPath $OutputDir `
                    -ComputerNameValue $ComputerName `
                    -RunIdValue $RunId

                $nonIntuneMdmRemovalSuccess = $removeMdmResult.Success
                $nonIntuneMdmRemovalBackupDir = $removeMdmResult.BackupDir
                $nonIntuneMdmRemovalDetail = $removeMdmResult.Detail
                Write-RunLog ("Non-Intune MDM removal result: Success={0}; BackupDir={1}; RemovedItems={2}; Detail={3}" -f $removeMdmResult.Success,$removeMdmResult.BackupDir,$removeMdmResult.RemovedItems,$removeMdmResult.Detail)

                $mdmEnrollmentState = Get-MdmEnrollmentState
                $anyMdmEnrollmentDetected = $mdmEnrollmentState.AnyMdmEnrollmentDetected
                $nonIntuneMdmEnrollmentDetected = $mdmEnrollmentState.NonIntuneMdmEnrollmentDetected
                $mdmEnrollmentCount = $mdmEnrollmentState.EnrollmentCount
                $mdmProviderIds = $mdmEnrollmentState.ProviderIds
                $intuneEnrollmentIds = $mdmEnrollmentState.IntuneEnrollmentIds
                $nonIntuneEnrollmentIds = $mdmEnrollmentState.NonIntuneEnrollmentIds
                $unconfirmedIntuneEnrollmentIds = $mdmEnrollmentState.UnconfirmedIntuneEnrollmentIds
                $mdmEnrollmentDetails = $mdmEnrollmentState.EnrollmentDetails
                $ignoredEnrollmentDetails = $mdmEnrollmentState.IgnoredEnrollmentDetails
                $intuneEnrolled = $mdmEnrollmentState.IntuneEnrollmentDetected

                if ($removeMdmResult.Success -and -not $nonIntuneMdmEnrollmentDetected) {
                    $status = "NON_INTUNE_MDM_REMOVED"
                    $dsregStatusErrorMessage = ("Non-Intune MDM enrollment was removed. BackupDir={0}. Intune auto-enrollment will be attempted on the next run/cycle." -f $removeMdmResult.BackupDir)
                    $ExitCode = 3
                }
                else {
                    $status = "NON_INTUNE_MDM_REMOVE_FAILED"
                    $dsregStatusErrorMessage = ("Non-Intune MDM removal did not fully complete. BackupDir={0}; Detail={1}" -f $removeMdmResult.BackupDir,$removeMdmResult.Detail)
                    $ExitCode = 3
                }

                Write-RunLog $dsregStatusErrorMessage
            }
            else {
                $status = "NON_INTUNE_MDM_ENROLLED"
                $dsregStatusErrorMessage = ("A non-Intune MDM enrollment is already present. Intune auto-enrollment is skipped to avoid fighting an existing MDM. ProviderIds={0}; EnrollmentIds={1}" -f $mdmProviderIds,$nonIntuneEnrollmentIds)
                $ExitCode = 3
                Write-Host "Hybrid Join is healthy, but a non-Intune MDM enrollment already exists. Skipping Intune auto-enrollment." -ForegroundColor Yellow
                Write-RunLog $dsregStatusErrorMessage
            }
        }
        if ($status -eq "SUCCESS" -and $intuneEnrolled -eq $false -and -not $nonIntuneMdmEnrollmentDetected) {
        $status = "INTUNE_ENROLLMENT_TRIGGERED"
        $dsregStatusErrorMessage = "Hybrid Join is healthy, but Intune enrollment is missing. MDM auto-enrollment repair is required."
        $ExitCode = 3

        Write-Host "Hybrid Join is healthy, but Intune enrollment is missing. Checking MDM auto-enrollment path..." -ForegroundColor Yellow
        Write-RunLog $dsregStatusErrorMessage

        $autoEnrollPolicyChecked = $true
        $policyState = Get-AutoEnrollmentPolicyState
        $autoEnrollPolicyKeyPresent = $policyState.PolicyKeyPresent
        $autoEnrollPolicyConfigured = $policyState.IsConfigured
        $autoEnrollMDM = $policyState.AutoEnrollMDM
        $autoEnrollUseAADCredentialType = $policyState.UseAADCredentialType
        $autoEnrollCredentialTypeLabel = $policyState.CredentialTypeLabel
        $autoEnrollPolicyDetail = $policyState.Detail
        Write-RunLog ("Auto-enrollment policy: KeyPresent={0}; Configured={1}; AutoEnrollMDM={2}; UseAADCredentialType={3}; Detail={4}" -f $autoEnrollPolicyKeyPresent,$autoEnrollPolicyConfigured,$autoEnrollMDM,$autoEnrollUseAADCredentialType,$autoEnrollPolicyDetail)

        $enrollmentConnectivityChecked = $true
        $connectivity = Test-EnrollmentConnectivity -AdditionalUrls @($mdmUrl,$mdmTouUrl,$mdmComplianceUrl)
        $enrollmentConnectivityOk = $connectivity.Success
        $enrollmentConnectivityHosts = $connectivity.TestedHosts
        $enrollmentConnectivityFailures = $connectivity.FailedHosts
        Write-RunLog ("Enrollment connectivity: Success={0}; Hosts={1}; Failures={2}" -f $enrollmentConnectivityOk,$enrollmentConnectivityHosts,$enrollmentConnectivityFailures)

        if (-not $policyState.IsConfigured) {
            $status = "INTUNE_AUTOENROLL_POLICY_NOT_CONFIGURED"
            $dsregStatusErrorMessage = $policyState.Detail
            if ($adComputerInDefaultComputersContainer -eq $true) {
                $dsregStatusErrorMessage = ("{0} {1}" -f $dsregStatusErrorMessage,$adComputerLocationDetail).Trim()
            }
            $ExitCode = 3
            Write-RunLog $dsregStatusErrorMessage
            if ([string]::IsNullOrWhiteSpace($gpResultHtmlFile) -or [string]::IsNullOrWhiteSpace($gpResultTextFile)) {
                $gpResult = Export-ComputerGpResultEvidence -ComputerName $ComputerName -RunId $RunId -OutputDir $OutputDir -Reason "MDM auto-enrollment policy is not configured"
                $gpResultHtmlFile = $gpResult.HtmlFile
                $gpResultTextFile = $gpResult.TextFile
            }
        }
        elseif (-not $connectivity.Success) {
            $status = "INTUNE_ENROLLMENT_CONNECTIVITY_FAILED"
            $dsregStatusErrorMessage = ("Enrollment HTTPS connectivity failed: {0}" -f $connectivity.FailedHosts)
            $ExitCode = 3
            Write-RunLog $dsregStatusErrorMessage
        }
        else {
            if (
                $policyState.UseAADCredentialType -eq "1" -and
                $interactiveUserDetected -and
                $userStatusAttempted -and
                [string]::IsNullOrWhiteSpace($userStatusErrorMessage) -and
                $userPrtRefreshStillNeeded
            ) {
                if ($userIsUserAzureAD -eq "NO") {
                    $status = "USER_NOT_AZUREAD"
                    $dsregStatusErrorMessage = "User-context dsreg status shows IsUserAzureAD=NO and AzureAdPrt is not available. User Credential MDM auto-enrollment cannot run successfully for this session."
                }
                elseif ($userRefreshPrtAttempted) {
                    $status = "USER_PRT_REFRESH_FAILED"
                    $dsregStatusErrorMessage = ("User-context PRT remains unusable after refreshprt. AzureAdPrt='{0}'; AzureAdPrtExpiryTime='{1}'; RefreshPrtAttemptStatus='{2}'; RefreshPrtHttpStatus='{3}'; RefreshPrtHttpError='{4}'; Reason={5}. User Credential MDM auto-enrollment is skipped for this session." -f $userAzureAdPrt,$userAzureAdPrtExpiryTime,$userRefreshPrtAttemptStatus,$userRefreshPrtHttpStatus,$userRefreshPrtHttpError,$userRefreshPrtReason)
                }
                else {
                    $status = "USER_PRT_NOT_AVAILABLE"
                    $dsregStatusErrorMessage = ("User-context PRT is not usable. AzureAdPrt='{0}'; AzureAdPrtExpiryTime='{1}'; RefreshPrtAttemptStatus='{2}'; RefreshPrtHttpStatus='{3}'; RefreshPrtHttpError='{4}'; Reason={5}. User Credential MDM auto-enrollment cannot run successfully until PRT is available." -f $userAzureAdPrt,$userAzureAdPrtExpiryTime,$userRefreshPrtAttemptStatus,$userRefreshPrtHttpStatus,$userRefreshPrtHttpError,$userRefreshPrtReason)
                }
                $ExitCode = 3
                Write-RunLog $dsregStatusErrorMessage
            }
            elseif (
                $policyState.UseAADCredentialType -eq "1" -and
                $interactiveUserDetected -and
                (
                    $interactiveSessionIsRemote -or
                    ($userStatusAttempted -and $userSessionIsNotRemote -eq "NO")
                )
            ) {
                $status = "USER_SESSION_REMOTE"
                $dsregStatusErrorMessage = ("Auto-enrollment policy uses User Credential, but the active user session is remote. SessionName='{0}'; SessionIsNotRemote='{1}'. User Credential MDM auto-enrollment should be retried from a console/local logon." -f $interactiveSessionName,$userSessionIsNotRemote)
                $ExitCode = 3
                Write-RunLog $dsregStatusErrorMessage
            }
            elseif ($policyState.UseAADCredentialType -eq "2") {
                $enrollStart = Start-IntuneAutoEnrollment
                $intuneAutoEnrollMode = "SystemDeviceCredential"
            }
            elseif (-not $interactiveUserDetected) {
                $enrollStart = [PSCustomObject]@{
                    Success=$false
                    ToolFound=$false
                    ExitCode=""
                    Detail="Auto-enrollment policy uses User Credential, but no active interactive user session was detected."
                    Mode="NoInteractiveUser"
                    TaskName=""
                    OutputFile=""
                }
                $intuneAutoEnrollMode = $enrollStart.Mode
            }
            elseif ($interactiveUserAccountType -eq "Local") {
                $enrollStart = [PSCustomObject]@{
                    Success=$false
                    ToolFound=$false
                    ExitCode=""
                    Detail=("Auto-enrollment policy uses User Credential, but the active interactive user is a local account: {0}." -f $interactiveUserAccountName)
                    Mode="LocalInteractiveUser"
                    TaskName=""
                    OutputFile=""
                }
                $intuneAutoEnrollMode = $enrollStart.Mode
            }
            else {
                $enrollStart = Start-UserContextIntuneAutoEnrollment -PreferredTaskName $UserIntuneAutoEnrollTaskName -WaitSeconds $UserTaskWaitSeconds -ExpectedUserName $interactiveUserName
                $intuneAutoEnrollMode = $enrollStart.Mode
            }

            if ($status -notin @("USER_NOT_AZUREAD","USER_PRT_NOT_AVAILABLE","USER_PRT_REFRESH_FAILED","USER_SESSION_REMOTE")) {
                $intuneAutoEnrollTaskName = $enrollStart.TaskName
                $intuneAutoEnrollExitCode = $enrollStart.ExitCode
                $intuneAutoEnrollDetail = $enrollStart.Detail
                $intuneAutoEnrollOutputFile = $enrollStart.OutputFile

                Write-RunLog ("Intune auto-enrollment trigger result: Mode={0}; TaskName={1}; Success={2}; ToolFound={3}; ExitCode={4}; Detail={5}" -f $intuneAutoEnrollMode,$intuneAutoEnrollTaskName,$enrollStart.Success,$enrollStart.ToolFound,$enrollStart.ExitCode,$enrollStart.Detail)

                if (-not $enrollStart.ToolFound) {
                    if ($policyState.UseAADCredentialType -eq "2") { $status = "INTUNE_ENROLLMENT_TOOL_NOT_FOUND" }
                    elseif (-not $interactiveUserDetected) { $status = "INTUNE_USER_AUTOENROLL_NO_INTERACTIVE_USER" }
                    elseif ($intuneAutoEnrollMode -eq "LocalInteractiveUser") { $status = "INTUNE_USER_AUTOENROLL_LOCAL_INTERACTIVE_USER" }
                    else { $status = "INTUNE_USER_AUTOENROLL_TASK_NOT_FOUND" }
                    $dsregStatusErrorMessage = $enrollStart.Detail
                    $ExitCode = 3

                    if ($status -eq "INTUNE_USER_AUTOENROLL_NO_INTERACTIVE_USER" -and $AllowRebootWhenNoInteractiveUser) {
                        $nextLogonTask = Register-UserAutoEnrollAtLogonTask -StatusTaskName $UserStatusTaskName -AutoEnrollTaskName $UserIntuneAutoEnrollTaskName
                        $nextLogonTaskRegistered = $nextLogonTask.Success
                        $nextLogonTaskName = $nextLogonTask.TaskName
                        $nextLogonTaskDetail = $nextLogonTask.Detail

                        $userContextKey = ("{0}|{1}|{2}" -f $interactiveUserAccountName,$interactiveSessionId,$interactiveSessionState).ToUpperInvariant()
                        $rebootSafetyDecision = Get-RebootSafetyDecision -ReasonKey "WAITING_FOR_USER_LOGON" -UserContext $userContextKey -MaximumReboots $RetryAfterRebootMaxAttempts -RequireUserContextChange
                        if (-not $rebootSafetyDecision.Allowed) {
                            $status = "WAITING_FOR_INTERACTIVE_USER_LOGON"
                            $dsregStatusErrorMessage = $rebootSafetyDecision.Detail
                            Write-RunLog $dsregStatusErrorMessage
                        }
                        else {
                            $rebootReason = "Hybrid Join is healthy and Intune enrollment is missing. Auto-enrollment uses User Credential; reboot resets local state, but enrollment still requires a valid interactive user logon with PRT."
                            try {
                                Register-RetryAfterRebootTask -Reason $rebootReason
                            }
                            catch {
                                $script:RetryAfterRebootAction = "ScheduleFailed"
                                $script:RetryAfterRebootTaskNameResult = $script:RetryAfterRebootTaskName
                                $script:RetryAfterRebootDetail = $_.Exception.Message
                                $status = "RETRY_AFTER_REBOOT_SCHEDULE_FAILED_WAITING_FOR_USER_LOGON"
                                $dsregStatusErrorMessage = "Reboot was not triggered because the retry-after-reboot task could not be scheduled. Error=$($_.Exception.Message)"
                                $ExitCode = 3
                                Write-RunLog $dsregStatusErrorMessage
                            }

                            if ($script:RetryAfterRebootAction -eq "Scheduled") {
                                Register-RebootSafetyEvent -ReasonKey "WAITING_FOR_USER_LOGON" -UserContext $userContextKey -Reason $rebootReason
                                $rebootAttempted = $true
                                $rebootResult = Start-ControlledReboot -Reason $rebootReason -DelaySeconds $RebootDelaySeconds
                                $rebootExitCode = $rebootResult.ExitCode
                                if ($rebootExitCode -eq 0) {
                                    $status = "REBOOT_TRIGGERED_WAITING_FOR_USER_LOGON"
                                    $dsregStatusErrorMessage = $rebootReason
                                }
                                else {
                                    $status = "REBOOT_SCHEDULE_FAILED_WAITING_FOR_USER_LOGON"
                                    $dsregStatusErrorMessage = "shutdown.exe failed with exit code $rebootExitCode. The retry-after-reboot task will be removed."
                                    $ExitCode = 3
                                }
                            }
                        }
                    }
                }
                else {
                    Write-ProvisionalRunState `
                        -StatusValue "INTUNE_ENROLLMENT_PENDING_CONFIRMATION" `
                        -DetailValue ("Intune auto-enrollment was triggered successfully. Waiting up to {0} minute(s) for confirmed local Intune enrollment before final status." -f ($IntuneRetrySleepMinutes * $IntuneRetryMaxRetries))

                    $r3 = Invoke-BoundedRetryUntilIntuneEnrollment `
                        -SleepMinutes $IntuneRetrySleepMinutes `
                        -MaxRetries $IntuneRetryMaxRetries `
                        -ContextLabel "intune_auto_enroll"

                    if ($r3.Success) {
                        $status = "SUCCESS"
                        $dsregStatusErrorMessage = ""
                        $intuneEnrolled = $true
                        $ExitCode = 0
                        Write-RunLog ("Intune enrollment detected after auto-enrollment retry loop after {0} attempt(s)." -f $r3.Attempts)
                    }
                    else {
                        $status = "INTUNE_ENROLLMENT_PENDING_CONFIRMATION"
                        $dsregStatusErrorMessage = ("Intune auto-enrollment was triggered successfully, but confirmed local Intune enrollment was not detected after {0} retries with {1} minutes sleep. Recheck later; Intune cloud and local MDM confirmation can lag after deviceenroller." -f $IntuneRetryMaxRetries, $IntuneRetrySleepMinutes)
                        $ExitCode = 3
                        Write-RunLog $dsregStatusErrorMessage

                        $intuneEnrollmentDiagnosticsFiles = Export-EnrollmentDiagnostics `
                            -Since $Timestamp `
                            -OutputDirPath $OutputDir `
                            -ComputerNameValue $ComputerName `
                            -RunIdValue $RunId
                        Write-RunLog ("Enrollment diagnostics exported: {0}" -f $intuneEnrollmentDiagnosticsFiles)
                    }
                }
            }
        }
        }
    }

    if (
        $userStatusAttempted -and
        [string]::IsNullOrWhiteSpace($userStatusErrorMessage) -and
        $userPrtRefreshStillNeeded -and
        $status -notin @("WAITING_FOR_AAD_CONNECT_LOCAL_RETRY","WAITING_FOR_AAD_CONNECT_LOCAL_RETRY_EXHAUSTED","LEAVE_EXECUTED_WAITING_FOR_REJOIN","WAITING_POST_LEAVE_LOCAL_RETRY_EXHAUSTED","REBOOT_TRIGGERED_POST_DSREG_LEAVE","RETRY_AFTER_REBOOT_SCHEDULE_FAILED_POST_DSREG_LEAVE","REBOOT_SCHEDULE_FAILED_POST_DSREG_LEAVE","ENTRA_HYBRID_PENDING_RETRY_EXHAUSTED","USER_PRT_REFRESH_FAILED")
    ) {
        if ($userIsUserAzureAD -eq "NO") {
            $status = "USER_NOT_AZUREAD"
            $dsregStatusErrorMessage = "User-context dsreg status shows IsUserAzureAD=NO and AzureAdPrt is not available. User Credential MDM auto-enrollment cannot run successfully for this session."
        }
        elseif ($userRefreshPrtAttempted) {
            $status = "USER_PRT_REFRESH_FAILED"
            $dsregStatusErrorMessage = ("User-context PRT remains unusable after refreshprt. AzureAdPrt='{0}'; AzureAdPrtExpiryTime='{1}'; RefreshPrtAttemptStatus='{2}'; RefreshPrtHttpStatus='{3}'; RefreshPrtHttpError='{4}'; Reason={5}. User PRT must be repaired before User Credential MDM auto-enrollment can succeed." -f $userAzureAdPrt,$userAzureAdPrtExpiryTime,$userRefreshPrtAttemptStatus,$userRefreshPrtHttpStatus,$userRefreshPrtHttpError,$userRefreshPrtReason)
        }
        else {
            $status = "USER_PRT_NOT_AVAILABLE"
            $dsregStatusErrorMessage = ("User-context PRT is not usable. AzureAdPrt='{0}'; AzureAdPrtExpiryTime='{1}'; RefreshPrtAttemptStatus='{2}'; RefreshPrtHttpStatus='{3}'; RefreshPrtHttpError='{4}'; Reason={5}. User PRT is required for User Credential MDM auto-enrollment." -f $userAzureAdPrt,$userAzureAdPrtExpiryTime,$userRefreshPrtAttemptStatus,$userRefreshPrtHttpStatus,$userRefreshPrtHttpError,$userRefreshPrtReason)
        }
        $ExitCode = 3
        Write-RunLog $dsregStatusErrorMessage
    }

    $finalGpUpdateReasons = New-Object System.Collections.Generic.List[string]
    if ($leaveAttempted) {
        [void]$finalGpUpdateReasons.Add("DsregLeaveAttempted")
    }
    if ($nonIntuneMdmRemovalAttempted) {
        [void]$finalGpUpdateReasons.Add("NonIntuneMdmRemovalAttempted")
    }
    if ($gpUpdateMdmPolicyWarning -and ($gpUpdateMdmPolicyState -ne "AlreadyEnrolled" -or $intuneEnrolled -ne $true)) {
        [void]$finalGpUpdateReasons.Add("InitialGpUpdateMdmPolicyWarning")
    }

    if ($finalGpUpdateReasons.Count -gt 0) {
        $finalGpUpdateReason = ($finalGpUpdateReasons -join ";")
        if ($rebootAttempted) {
            Write-RunLog ("Final gpupdate skipped because a reboot is already scheduled. Reason={0}" -f $finalGpUpdateReason)
        }
        else {
            $finalGpUpdateAttempted = $true
            $finalGpUpdateOutputFile = Join-Path $OutputDir ("{0}_gpupdate_final_{1}.txt" -f $ComputerName, $RunId)
            Write-Host "Refreshing computer Group Policy after repair action..." -ForegroundColor Cyan
            Write-RunLog ("Running final gpupdate /target:computer /force. Reason={0}" -f $finalGpUpdateReason)
            try {
                $finalGpupdateOutput = gpupdate.exe /target:computer /force 2>&1
                $finalGpUpdateExitCode = $LASTEXITCODE
                $finalGpupdateOutput | Out-File -FilePath $finalGpUpdateOutputFile -Encoding UTF8 -Force
                Write-RunLog ("Final gpupdate completed. ExitCode={0}; OutputFile={1}" -f $finalGpUpdateExitCode,$finalGpUpdateOutputFile)
            }
            catch {
                $finalGpUpdateExitCode = ""
                Write-RunLog ("Final gpupdate failed to start: {0}" -f $_.Exception.Message)
            }
        }
    }

    $nextAction = Get-NextActionForStatus -Status $status -IntuneEnrolled ([bool]$intuneEnrolled)

    # Final log entry
    $logEntryFinal = [PSCustomObject]@{
        RunId                   = $RunId
        Timestamp               = $Timestamp
        ComputerName            = $ComputerName
        ScriptVersion           = $ScriptVersion
        AllowDsregLeave                = [bool]$AllowDsregLeave
        AllowRemoveNonIntuneMdmEnrollment = [bool]$AllowRemoveNonIntuneMdmEnrollment
        AllowRemoveStaleIntuneEnrollment = [bool]$AllowRemoveStaleIntuneEnrollment
        SkipVirtualMachines     = [bool]$SkipVirtualMachines
        AuditOnly               = [bool]$AuditOnly
        EntraHybridPending      = [bool]$EntraHybridPending
        NextAction              = $nextAction
        OsCaption               = $osCaption
        OsVersion               = $osVersion
        OsBuildNumber           = $osBuildNumber
        OsArchitecture          = $osArchitecture
        OsProductType           = $osProductType
        LastBootUpTime          = $lastBootUpTime
        UptimeHours             = $uptimeHours
        UptimeDays              = $uptimeDays
        BootInfoDetail          = $bootInfoDetail
        IsVirtualMachine        = $isVirtualMachine
        VirtualMachineEvidence  = $virtualMachineEvidence
        AdComputerLocationChecked = $adComputerLocationChecked
        AdComputerDistinguishedName = $adComputerDistinguishedName
        AdComputerDefaultNamingContext = $adComputerDefaultNamingContext
        AdComputerInDefaultComputersContainer = $adComputerInDefaultComputersContainer
        AdComputerLocationDetail = $adComputerLocationDetail
        LeaveAttempted          = $leaveAttempted
        LeaveExitCode           = $leaveExitCode
        Status                  = $status
        ErrorMessage            = $errorMessage
        Dsreg_AzureAdJoined     = $dsregAzureAdJoined
        Dsreg_DeviceId          = $dsregDeviceId
        Dsreg_TenantName        = $dsregTenantName
        Dsreg_TenantId          = $dsregTenantId
        DeviceAuthStatus        = $dsregDeviceAuthStatus
        Dsreg_KeySignTest       = $dsregKeySignTest
        DsregStatusErrorMessage = $dsregStatusErrorMessage
        IntuneEnrolled          = $intuneEnrolled
        AnyMdmEnrollmentDetected = $anyMdmEnrollmentDetected
        NonIntuneMdmEnrollmentDetected = $nonIntuneMdmEnrollmentDetected
        MdmEnrollmentCount      = $mdmEnrollmentCount
        MdmProviderIds          = $mdmProviderIds
        IntuneEnrollmentIds     = $intuneEnrollmentIds
        NonIntuneEnrollmentIds  = $nonIntuneEnrollmentIds
        UnconfirmedIntuneEnrollmentIds = $unconfirmedIntuneEnrollmentIds
        StaleIntuneEnrollmentDetected = $staleIntuneEnrollmentDetected
        StaleIntuneEnrollmentIds = $staleIntuneEnrollmentIds
        MdmEnrollmentDetails    = $mdmEnrollmentDetails
        IgnoredEnrollmentDetails = $ignoredEnrollmentDetails
        NonIntuneMdmRemovalAttempted = $nonIntuneMdmRemovalAttempted
        NonIntuneMdmRemovalSuccess = $nonIntuneMdmRemovalSuccess
        NonIntuneMdmRemovalBackupDir = $nonIntuneMdmRemovalBackupDir
        NonIntuneMdmRemovalDetail = $nonIntuneMdmRemovalDetail
        StaleIntuneRemovalAttempted = $staleIntuneRemovalAttempted
        StaleIntuneRemovalSuccess = $staleIntuneRemovalSuccess
        StaleIntuneRemovalBackupDir = $staleIntuneRemovalBackupDir
        StaleIntuneRemovalDetail = $staleIntuneRemovalDetail
        StaleIntuneCleanupContinuedToAutoEnroll = $staleIntuneCleanupContinuedToAutoEnroll
        StaleCleanupDelaySeconds = $StaleCleanupDelaySeconds
        IntuneRetrySleepMinutes = $IntuneRetrySleepMinutes
        IntuneRetryMaxRetries = $IntuneRetryMaxRetries
        IntuneRetryWindowMinutes = ($IntuneRetrySleepMinutes * $IntuneRetryMaxRetries)
        ClientErrorCode         = $clientErrorCode
        ServerErrorCode         = $serverErrorCode
        ServerErrorSubCode      = $serverErrorSubCode
        ServerOperation         = $serverOperation
        ServerMessage           = $serverMessage
        HttpsStatus             = $httpsStatus
        RequestId               = $requestId
        ErrorPhase              = $errorPhase
        MdmUrl                  = $mdmUrl
        MdmTouUrl               = $mdmTouUrl
        MdmComplianceUrl        = $mdmComplianceUrl
        AutoEnrollPolicyChecked = $autoEnrollPolicyChecked
        AutoEnrollPolicyKeyPresent = $autoEnrollPolicyKeyPresent
        AutoEnrollPolicyConfigured = $autoEnrollPolicyConfigured
        AutoEnrollMDM           = $autoEnrollMDM
        AutoEnrollUseAADCredentialType = $autoEnrollUseAADCredentialType
        AutoEnrollCredentialTypeLabel = $autoEnrollCredentialTypeLabel
        AutoEnrollPolicyDetail  = $autoEnrollPolicyDetail
        EnrollmentConnectivityChecked = $enrollmentConnectivityChecked
        EnrollmentConnectivityOk = $enrollmentConnectivityOk
        EnrollmentConnectivityHosts = $enrollmentConnectivityHosts
        EnrollmentConnectivityFailures = $enrollmentConnectivityFailures
        IntuneAutoEnrollMode    = $intuneAutoEnrollMode
        IntuneAutoEnrollTaskName= $intuneAutoEnrollTaskName
        IntuneAutoEnrollExitCode= $intuneAutoEnrollExitCode
        IntuneAutoEnrollDetail  = $intuneAutoEnrollDetail
        IntuneAutoEnrollOutputFile = $intuneAutoEnrollOutputFile
        IntuneEnrollmentDiagnosticsInitialFiles = $intuneEnrollmentDiagnosticsInitialFiles
        IntuneEnrollmentDiagnosticsFiles = $intuneEnrollmentDiagnosticsFiles
        UserStatusAttempted     = $userStatusAttempted
        UserStatusTaskExitCode  = $userStatusTaskExitCode
        UserStatusFile          = $userStatusFile
        UserStatusErrorMessage  = $userStatusErrorMessage
        User_AzureAdPrt         = $userAzureAdPrt
        User_AzureAdPrtAuthority = $userAzureAdPrtAuthority
        User_AzureAdPrtUpdateTime = $userAzureAdPrtUpdateTime
        User_AzureAdPrtExpiryTime = $userAzureAdPrtExpiryTime
        User_RefreshPrtAttemptStatus = $userRefreshPrtAttemptStatus
        User_RefreshPrtHttpStatus = $userRefreshPrtHttpStatus
        User_RefreshPrtHttpError = $userRefreshPrtHttpError
        User_RefreshPrtServerErrorCode = $userRefreshPrtServerErrorCode
        User_RefreshPrtServerErrorSubCode = $userRefreshPrtServerErrorSubCode
        User_RefreshPrtReason   = $userRefreshPrtReason
        User_PrtRefreshStillNeeded = $userPrtRefreshStillNeeded
        User_EnterprisePrt      = $userEnterprisePrt
        User_OnPremTgt          = $userOnPremTgt
        User_CloudTgt           = $userCloudTgt
        User_WorkplaceJoined    = $userWorkplaceJoined
        User_IsUserAzureAD      = $userIsUserAzureAD
        User_SessionIsNotRemote = $userSessionIsNotRemote
        User_WamDefaultSet      = $userWamDefaultSet
        User_WamDefaultAuthority= $userWamDefaultAuthority
        User_WamDefaultId       = $userWamDefaultId
        User_NgcSet             = $userNgcSet
        UserRefreshPrtAttempted = $userRefreshPrtAttempted
        UserRefreshPrtExitCode  = $userRefreshPrtExitCode
        UserRefreshPrtFile      = $userRefreshPrtFile
        FlushDnsAttempted       = $flushDnsAttempted
        FlushDnsExitCode        = $flushDnsExitCode
        FlushDnsOutputFile      = $flushDnsOutputFile
        GpUpdateAttempted       = $gpUpdateAttempted
        GpUpdateExitCode        = $gpUpdateExitCode
        GpUpdateOutputFile      = $gpUpdateOutputFile
        GpUpdateMdmPolicyWarning = $gpUpdateMdmPolicyWarning
        GpUpdateMdmPolicyState  = $gpUpdateMdmPolicyState
        GpUpdateMdmPolicyDetail = $gpUpdateMdmPolicyDetail
        GpResultHtmlFile        = $gpResultHtmlFile
        GpResultTextFile        = $gpResultTextFile
        FinalGpUpdateAttempted  = $finalGpUpdateAttempted
        FinalGpUpdateReason     = $finalGpUpdateReason
        FinalGpUpdateExitCode   = $finalGpUpdateExitCode
        FinalGpUpdateOutputFile = $finalGpUpdateOutputFile
        InteractiveUserDetected = $interactiveUserDetected
        InteractiveUserName     = $interactiveUserName
        InteractiveUserDomain   = $interactiveUserDomain
        InteractiveUserAccountName = $interactiveUserAccountName
        InteractiveUserAccountType = $interactiveUserAccountType
        InteractiveUserIdentityResolved = $interactiveUserIdentityResolved
        InteractiveUserIdentityDetail = $interactiveUserIdentityDetail
        InteractiveSessionName  = $interactiveSessionName
        InteractiveSessionId    = $interactiveSessionId
        InteractiveSessionState = $interactiveSessionState
        InteractiveSessionIsRemote = $interactiveSessionIsRemote
        InteractiveSessionDetail= $interactiveSessionDetail
        RebootAttempted         = $rebootAttempted
        RebootReason            = $rebootReason
        RebootExitCode          = $rebootExitCode
        NextLogonTaskRegistered = $nextLogonTaskRegistered
        NextLogonTaskName       = $nextLogonTaskName
        NextLogonTaskDetail     = $nextLogonTaskDetail
        RetryAfterRebootAction  = $script:RetryAfterRebootAction
        RetryAfterRebootDetail  = $script:RetryAfterRebootDetail
        RetryAfterRebootAttempt = $script:RetryAfterRebootAttempt
        RetryAfterRebootMaxAttempts = $script:RetryAfterRebootMaxAttemptsResult
        RetryAfterRebootTaskName = $script:RetryAfterRebootTaskNameResult
    }

    Write-AtomicCsvAppend -Path $logPath -RowObject $logEntryFinal -RunIdValue $RunId

    try {
        [PSCustomObject]@{
            RunId        = $RunId
            Version      = $ScriptVersion
            StartTime    = $Timestamp.ToString("o")
            EndTime      = (Get-Date).ToString("o")
            ComputerName = $ComputerName
            AllowDsregLeave     = [bool]$AllowDsregLeave
            AuditOnly    = [bool]$AuditOnly
            Status       = $status
            ExitCode     = $ExitCode
            NextAction   = $nextAction
            Detail       = $dsregStatusErrorMessage
            RetryAfterRebootAction = $script:RetryAfterRebootAction
            RetryAfterRebootDetail = $script:RetryAfterRebootDetail
            RetryAfterRebootAttempt = $script:RetryAfterRebootAttempt
            RetryAfterRebootMaxAttempts = $script:RetryAfterRebootMaxAttemptsResult
            RetryAfterRebootTaskName = $script:RetryAfterRebootTaskNameResult
        } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $RunGuardPath -Encoding UTF8 -Force
    }
    catch {
        Write-RunLog ("Run guard final update failed. Error={0}" -f $_.Exception.Message)
    }

    Write-Host "Local CSV log saved to $logPath"
    Write-Host "RunId for this execution: $RunId"
    Write-FinalStatusLine -Status $status -ExitCode $ExitCode -Detail $dsregStatusErrorMessage -NextAction $nextAction
    Write-RunLog ("Script end. Status={0}; ExitCode={1}" -f $status,$ExitCode)
}
catch {
    $controlledRetryStop = $status -in @("RETRY_AFTER_REBOOT_EXHAUSTED","RETRY_AFTER_REBOOT_STATE_MISSING")
    $errorMessage = $_.Exception.Message
    if (-not $controlledRetryStop) {
        $status = "ERROR"
        $ExitCode = 1
    }
    $nextAction = Get-NextActionForStatus -Status $status -IntuneEnrolled ([bool]$intuneEnrolled)

    if ($controlledRetryStop) {
        Write-Host $errorMessage -ForegroundColor Yellow
        Write-FinalStatusLine -Status $status -ExitCode $ExitCode -Detail $errorMessage -NextAction $nextAction
        Write-RunLog ("Controlled retry stop. Status={0}; Detail={1}" -f $status,$errorMessage)
    }
    else {
        Write-Host "FATAL ERROR: $errorMessage" -ForegroundColor Red
        Write-FinalStatusLine -Status "ERROR" -ExitCode 1 -Detail $errorMessage -NextAction $nextAction
        Write-RunLog ("FATAL ERROR: {0}" -f $errorMessage)
    }

    try {
        $logEntryFinal = [PSCustomObject]@{
            RunId                   = $RunId
            Timestamp               = $Timestamp
            ComputerName            = $ComputerName
            ScriptVersion           = $ScriptVersion
            AllowDsregLeave                = [bool]$AllowDsregLeave
            AllowRemoveNonIntuneMdmEnrollment = [bool]$AllowRemoveNonIntuneMdmEnrollment
            AllowRemoveStaleIntuneEnrollment = [bool]$AllowRemoveStaleIntuneEnrollment
            SkipVirtualMachines     = [bool]$SkipVirtualMachines
            AuditOnly               = [bool]$AuditOnly
            EntraHybridPending      = [bool]$EntraHybridPending
            NextAction              = $nextAction
            OsCaption               = $osCaption
            OsVersion               = $osVersion
            OsBuildNumber           = $osBuildNumber
            OsArchitecture          = $osArchitecture
            OsProductType           = $osProductType
            LastBootUpTime          = $lastBootUpTime
            UptimeHours             = $uptimeHours
            UptimeDays              = $uptimeDays
            BootInfoDetail          = $bootInfoDetail
            IsVirtualMachine        = $isVirtualMachine
            VirtualMachineEvidence  = $virtualMachineEvidence
            AdComputerLocationChecked = $adComputerLocationChecked
            AdComputerDistinguishedName = $adComputerDistinguishedName
            AdComputerDefaultNamingContext = $adComputerDefaultNamingContext
            AdComputerInDefaultComputersContainer = $adComputerInDefaultComputersContainer
            AdComputerLocationDetail = $adComputerLocationDetail
            LeaveAttempted          = $leaveAttempted
            LeaveExitCode           = $leaveExitCode
            Status                  = $status
            ErrorMessage            = $errorMessage
            Dsreg_AzureAdJoined     = $dsregAzureAdJoined
            Dsreg_DeviceId          = $dsregDeviceId
            Dsreg_TenantName        = $dsregTenantName
            Dsreg_TenantId          = $dsregTenantId
            DeviceAuthStatus        = $dsregDeviceAuthStatus
            Dsreg_KeySignTest       = $dsregKeySignTest
            DsregStatusErrorMessage = $dsregStatusErrorMessage
            IntuneEnrolled          = $intuneEnrolled
            AnyMdmEnrollmentDetected = $anyMdmEnrollmentDetected
            NonIntuneMdmEnrollmentDetected = $nonIntuneMdmEnrollmentDetected
            MdmEnrollmentCount      = $mdmEnrollmentCount
            MdmProviderIds          = $mdmProviderIds
            IntuneEnrollmentIds     = $intuneEnrollmentIds
            NonIntuneEnrollmentIds  = $nonIntuneEnrollmentIds
            UnconfirmedIntuneEnrollmentIds = $unconfirmedIntuneEnrollmentIds
            StaleIntuneEnrollmentDetected = $staleIntuneEnrollmentDetected
            StaleIntuneEnrollmentIds = $staleIntuneEnrollmentIds
            MdmEnrollmentDetails    = $mdmEnrollmentDetails
            IgnoredEnrollmentDetails = $ignoredEnrollmentDetails
            NonIntuneMdmRemovalAttempted = $nonIntuneMdmRemovalAttempted
            NonIntuneMdmRemovalSuccess = $nonIntuneMdmRemovalSuccess
            NonIntuneMdmRemovalBackupDir = $nonIntuneMdmRemovalBackupDir
            NonIntuneMdmRemovalDetail = $nonIntuneMdmRemovalDetail
            StaleIntuneRemovalAttempted = $staleIntuneRemovalAttempted
            StaleIntuneRemovalSuccess = $staleIntuneRemovalSuccess
            StaleIntuneRemovalBackupDir = $staleIntuneRemovalBackupDir
            StaleIntuneRemovalDetail = $staleIntuneRemovalDetail
            StaleIntuneCleanupContinuedToAutoEnroll = $staleIntuneCleanupContinuedToAutoEnroll
            StaleCleanupDelaySeconds = $StaleCleanupDelaySeconds
            IntuneRetrySleepMinutes = $IntuneRetrySleepMinutes
            IntuneRetryMaxRetries = $IntuneRetryMaxRetries
            IntuneRetryWindowMinutes = ($IntuneRetrySleepMinutes * $IntuneRetryMaxRetries)
            ClientErrorCode         = $clientErrorCode
            ServerErrorCode         = $serverErrorCode
            ServerErrorSubCode      = $serverErrorSubCode
            ServerOperation         = $serverOperation
            ServerMessage           = $serverMessage
            HttpsStatus             = $httpsStatus
            RequestId               = $requestId
            ErrorPhase              = $errorPhase
            MdmUrl                  = $mdmUrl
            MdmTouUrl               = $mdmTouUrl
            MdmComplianceUrl        = $mdmComplianceUrl
            AutoEnrollPolicyChecked = $autoEnrollPolicyChecked
            AutoEnrollPolicyKeyPresent = $autoEnrollPolicyKeyPresent
            AutoEnrollPolicyConfigured = $autoEnrollPolicyConfigured
            AutoEnrollMDM           = $autoEnrollMDM
            AutoEnrollUseAADCredentialType = $autoEnrollUseAADCredentialType
            AutoEnrollCredentialTypeLabel = $autoEnrollCredentialTypeLabel
            AutoEnrollPolicyDetail  = $autoEnrollPolicyDetail
            EnrollmentConnectivityChecked = $enrollmentConnectivityChecked
            EnrollmentConnectivityOk = $enrollmentConnectivityOk
            EnrollmentConnectivityHosts = $enrollmentConnectivityHosts
            EnrollmentConnectivityFailures = $enrollmentConnectivityFailures
            IntuneAutoEnrollMode    = $intuneAutoEnrollMode
            IntuneAutoEnrollTaskName= $intuneAutoEnrollTaskName
            IntuneAutoEnrollExitCode= $intuneAutoEnrollExitCode
            IntuneAutoEnrollDetail  = $intuneAutoEnrollDetail
            IntuneAutoEnrollOutputFile = $intuneAutoEnrollOutputFile
            IntuneEnrollmentDiagnosticsInitialFiles = $intuneEnrollmentDiagnosticsInitialFiles
            IntuneEnrollmentDiagnosticsFiles = $intuneEnrollmentDiagnosticsFiles
            UserStatusAttempted     = $userStatusAttempted
            UserStatusTaskExitCode  = $userStatusTaskExitCode
            UserStatusFile          = $userStatusFile
            UserStatusErrorMessage  = $userStatusErrorMessage
            User_AzureAdPrt         = $userAzureAdPrt
            User_AzureAdPrtAuthority = $userAzureAdPrtAuthority
            User_AzureAdPrtUpdateTime = $userAzureAdPrtUpdateTime
            User_AzureAdPrtExpiryTime = $userAzureAdPrtExpiryTime
            User_RefreshPrtAttemptStatus = $userRefreshPrtAttemptStatus
            User_RefreshPrtHttpStatus = $userRefreshPrtHttpStatus
            User_RefreshPrtHttpError = $userRefreshPrtHttpError
            User_RefreshPrtServerErrorCode = $userRefreshPrtServerErrorCode
            User_RefreshPrtServerErrorSubCode = $userRefreshPrtServerErrorSubCode
            User_RefreshPrtReason   = $userRefreshPrtReason
            User_PrtRefreshStillNeeded = $userPrtRefreshStillNeeded
            User_EnterprisePrt      = $userEnterprisePrt
            User_OnPremTgt          = $userOnPremTgt
            User_CloudTgt           = $userCloudTgt
            User_WorkplaceJoined    = $userWorkplaceJoined
            User_IsUserAzureAD      = $userIsUserAzureAD
            User_SessionIsNotRemote = $userSessionIsNotRemote
            User_WamDefaultSet      = $userWamDefaultSet
            User_WamDefaultAuthority= $userWamDefaultAuthority
            User_WamDefaultId       = $userWamDefaultId
            User_NgcSet             = $userNgcSet
            UserRefreshPrtAttempted = $userRefreshPrtAttempted
            UserRefreshPrtExitCode  = $userRefreshPrtExitCode
            UserRefreshPrtFile      = $userRefreshPrtFile
            FlushDnsAttempted       = $flushDnsAttempted
            FlushDnsExitCode        = $flushDnsExitCode
            FlushDnsOutputFile      = $flushDnsOutputFile
            GpUpdateAttempted       = $gpUpdateAttempted
            GpUpdateExitCode        = $gpUpdateExitCode
            GpUpdateOutputFile      = $gpUpdateOutputFile
            GpUpdateMdmPolicyWarning = $gpUpdateMdmPolicyWarning
            GpUpdateMdmPolicyState  = $gpUpdateMdmPolicyState
            GpUpdateMdmPolicyDetail = $gpUpdateMdmPolicyDetail
            GpResultHtmlFile        = $gpResultHtmlFile
            GpResultTextFile        = $gpResultTextFile
            FinalGpUpdateAttempted  = $finalGpUpdateAttempted
            FinalGpUpdateReason     = $finalGpUpdateReason
            FinalGpUpdateExitCode   = $finalGpUpdateExitCode
            FinalGpUpdateOutputFile = $finalGpUpdateOutputFile
            InteractiveUserDetected = $interactiveUserDetected
            InteractiveUserName     = $interactiveUserName
            InteractiveUserDomain   = $interactiveUserDomain
            InteractiveUserAccountName = $interactiveUserAccountName
            InteractiveUserAccountType = $interactiveUserAccountType
            InteractiveUserIdentityResolved = $interactiveUserIdentityResolved
            InteractiveUserIdentityDetail = $interactiveUserIdentityDetail
            InteractiveSessionName  = $interactiveSessionName
            InteractiveSessionId    = $interactiveSessionId
            InteractiveSessionState = $interactiveSessionState
            InteractiveSessionIsRemote = $interactiveSessionIsRemote
            InteractiveSessionDetail= $interactiveSessionDetail
            RebootAttempted         = $rebootAttempted
            RebootReason            = $rebootReason
            RebootExitCode          = $rebootExitCode
            NextLogonTaskRegistered = $nextLogonTaskRegistered
            NextLogonTaskName       = $nextLogonTaskName
            NextLogonTaskDetail     = $nextLogonTaskDetail
            RetryAfterRebootAction  = $script:RetryAfterRebootAction
            RetryAfterRebootDetail  = $script:RetryAfterRebootDetail
            RetryAfterRebootAttempt = $script:RetryAfterRebootAttempt
            RetryAfterRebootMaxAttempts = $script:RetryAfterRebootMaxAttemptsResult
            RetryAfterRebootTaskName = $script:RetryAfterRebootTaskNameResult
        }

        Write-AtomicCsvAppend -Path $logPath -RowObject $logEntryFinal -RunIdValue $RunId
    }
    catch { }

    try {
        [PSCustomObject]@{
            RunId        = $RunId
            Version      = $ScriptVersion
            StartTime    = $Timestamp.ToString("o")
            EndTime      = (Get-Date).ToString("o")
            ComputerName = $ComputerName
            AllowDsregLeave     = [bool]$AllowDsregLeave
            AuditOnly    = [bool]$AuditOnly
            Status       = $status
            ExitCode     = $ExitCode
            NextAction   = $nextAction
            Detail       = $(if (-not [string]::IsNullOrWhiteSpace($dsregStatusErrorMessage)) { $dsregStatusErrorMessage } else { $errorMessage })
            RetryAfterRebootAction = $script:RetryAfterRebootAction
            RetryAfterRebootDetail = $script:RetryAfterRebootDetail
            RetryAfterRebootAttempt = $script:RetryAfterRebootAttempt
            RetryAfterRebootMaxAttempts = $script:RetryAfterRebootMaxAttemptsResult
            RetryAfterRebootTaskName = $script:RetryAfterRebootTaskNameResult
            ErrorMessage = $errorMessage
        } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $RunGuardPath -Encoding UTF8 -Force
    }
    catch { }
}
finally {
    try {
        if (
            ($RetryAfterRebootTaskRun -or (Test-Path -LiteralPath $script:RetryAfterRebootStatePath -PathType Leaf)) -and
            $status -notlike "REBOOT_TRIGGERED*"
        ) {
            Unregister-RetryAfterRebootTask -Reason $status
        }
    }
    catch {
        Write-RunLog ("Retry-after-reboot final cleanup failed. Error={0}" -f $_.Exception.Message)
    }

    try {
        if ($TranscriptStarted) {
            Stop-Transcript | Out-Null
            Update-TimestampedTranscript -Path $TranscriptFile
            Write-RunLog "Transcript stopped."
        }
    }
    catch {
        Write-RunLog ("Stop-Transcript failed: {0}" -f $_.Exception.Message)
    }

    try { Update-EndpointInstanceState -Status $status -Force } catch { }
    if ($script:EndpointInstanceMutexAcquired -and $script:EndpointInstanceMutex) {
        try { $script:EndpointInstanceMutex.ReleaseMutex() } catch { }
    }
    if ($script:EndpointInstanceMutex) {
        try { $script:EndpointInstanceMutex.Dispose() } catch { }
    }
    $script:EndpointInstanceMutexAcquired = $false
}

exit $ExitCode

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA8my6ffV2Wink3
# 6gV2blIq00op9815iXAVFxon5eOKM6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAglpjf1yf8KiaGRXCLrnS2
# C+j998LeE8IAoeXQnEvFkTANBgkqhkiG9w0BAQEFAASCAYA5yLgWxbMgJGoKBf5Q
# OoMYlgDDbsHPZZzTIS//YBiKzyZtPT/YXEslVQU353k80nmVVBrX1LvDkM0bkpj9
# hVIVTLnyg1An++LgfCMAoALvXybpMInVKCfO21hcdRF44o/lvEfEePsna73YO5Yd
# pXl6osH+SrjZGPfAOrvfpvkPFawyrck8aFtB20joFtUoWPwpyuCHH5XG2DHifNGQ
# eHGj4Ob18ew30Ecj2H4bdh/bGn9ECH/wQkuxyQbkubs52H4yOn7rqOQGX7EAG+ik
# SrH1sPhC8AKR5kOd/mn0yL9NkorwYMgQgTkdB16xupNeb3kLrXeDP+IrvgVz9nmT
# NQr+W74ZBTttcoPO2MB/n1R7Oq7g6g67ifnTAUq/uSUDyZeyklPAnQO/isZbopPO
# V6kcmNlQXhwkbLw9VwRHcYf1bgiAItU7PkRrg4uV6ywLQBZ2xoTLGcn07lrJSQjM
# sg8goPy8g1hoMPg8lTY1Wrh/1e3ZiitbyT1FJp9JXj26T/w=
# SIG # End signature block
