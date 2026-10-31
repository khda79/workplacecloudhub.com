<#
.SYNOPSIS
SYSTEM-safe Entra Hybrid Join and Intune enrollment repair helper (dsregcmd /status, optional /leave, trigger Automatic-Device-Join, optional MDM auto-enrollment, bounded local retry).

.DESCRIPTION
Script name: SmartM365-Invoke-EndpointMigration-HybridJoin.ps1
Campaign: HybridJoin of the Smart Endpoint Migration Toolkit (autonomous single-file endpoint script).
Provenance: copied from IntuneHybridJoinToolkit SmartM365-Invoke-IntuneHybridJoinRepair.ps1 2.10.40 (see Docs\PROVENANCE.json.txt).

Runs locally on the machine (typically under LocalSystem through PsExec). Writes under -DataRoot,
by default C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin.
- One Smart Endpoint Migration Toolkit operation at a time per computer (shared target guard).
- Previous toolkits are inspected read-only; a real active legacy operation stops the run.
- Each attempt writes State\Results\<CorrelationId>.json.txt for the technician worker.

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
Allows removal of confirmed non-Intune MDM enrollment registry keys and EnterpriseMgmt scheduled tasks before trying Intune auto-enrollment. WMI-to-CSP bridge registrations are never removable through this option.

.PARAMETER AllowRemoveStaleIntuneEnrollment
Retained for backward compatibility. Ambiguous local Intune traces are not removed automatically; manual review is required.

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

.PARAMETER DataRoot
Campaign runtime folder. Defaults to C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin.

.PARAMETER CorrelationId
Attempt identifier supplied by the technician worker. A local run generates one.

.PARAMETER LegacyActivityPolicy
Warn (default) records previous toolkit activity and continues. Block stops when previous toolkit activity is detected.

.VERSION
1.0.0

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
    [switch]$CollectMdmDiagnostics,
    [switch]$EntraHybridPending,
    [int]$StaleCleanupDelaySeconds = 60,
    [int]$RebootDelaySeconds = 180,
    [int]$IntuneRetrySleepMinutes = 5,
    [int]$IntuneRetryMaxRetries = 5,
    [ValidateRange(0,3600)][int]$RetryAfterRebootDelaySeconds = 300,
    [ValidateRange(1,30)][int]$RetryAfterRebootMaxAttempts = 3,
    [switch]$RetryAfterRebootTaskRun,
    [string]$DataRoot = 'C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin',
    [string]$CorrelationId,
    [ValidateSet('Block','Warn')][string]$LegacyActivityPolicy = 'Warn'
)

# Endpoint scripts run in Windows PowerShell 5.1 (the worker and the Intune runner start them with
# powershell.exe). Any other engine is refused here, before any state is written.
if ([string]$PSVersionTable['PSEdition'] -eq 'Core' -or $PSVersionTable['PSVersion'] -lt [version]'5.1') {
    Write-Output ('FINAL_STATUS=UNSUPPORTED_POWERSHELL_HOST; EXIT_CODE=3; NEXT_ACTION=RUN_WITH_WINDOWS_POWERSHELL_5_1; CORRELATION_ID={0}; DETAIL={1} {2} is not supported; run this script with powershell.exe (Windows PowerShell 5.1).' -f [string]$CorrelationId,[string]$PSVersionTable['PSEdition'],[string]$PSVersionTable['PSVersion'])
    exit 3
}
$ScriptVersion = "1.0.0"
$script:SemtCampaign = "HybridJoin"
if ($RebootDelaySeconds -lt 60) { $RebootDelaySeconds = 60 }
if ($StaleCleanupDelaySeconds -lt 0) { $StaleCleanupDelaySeconds = 0 }
if ($IntuneRetrySleepMinutes -lt 1) { $IntuneRetrySleepMinutes = 1 }
if ($IntuneRetryMaxRetries -lt 1) { $IntuneRetryMaxRetries = 1 }

# ============================
# Enforce 64-bit PowerShell (anti-WOW64)
# ============================
try {
    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
        if (-not $env:SEMT_HJ_Relaunched64) {
            $env:SEMT_HJ_Relaunched64 = "1"

            $sysNativePS = Join-Path $env:WINDIR "SysNative\WindowsPowerShell\v1.0\powershell.exe"
            if (-not (Test-Path $sysNativePS)) {
                $sysNativePS = Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"
            }

            $argList = @(
                "-NoProfile",
                "-ExecutionPolicy", "RemoteSigned",
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
            if ($CollectMdmDiagnostics) { $argList += "-CollectMdmDiagnostics" }
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
            $argList += "-DataRoot"
            $argList += "`"$DataRoot`""
            if (-not [string]::IsNullOrWhiteSpace($CorrelationId)) { $argList += "-CorrelationId"; $argList += $CorrelationId }
            $argList += "-LegacyActivityPolicy"
            $argList += $LegacyActivityPolicy

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
$DataRoot      = [System.IO.Path]::GetFullPath($DataRoot)
$OutputDir     = Join-Path $DataRoot "Output"
$TranscriptDir = Join-Path $DataRoot "Transcripts"
$LogsDir       = Join-Path $DataRoot "Logs"
$StateDir      = Join-Path $DataRoot "State"

# The data folders are created after the data folder boundary below (nothing is written earlier).

# Run metadata
$RunId        = (Get-Date -Format "yyyyMMdd_HHmmss_fff")
$Timestamp    = Get-Date
$ComputerName = $env:COMPUTERNAME
$RunGuardHours = 12
$CleanupRetentionDays = 7
$RunGuardPath = Join-Path $DataRoot "LastRun.json.txt"
$script:RetryAfterRebootTaskName = "SmartM365-SEMT-HybridJoin-RetryAfterReboot"
# The retry task runs as SYSTEM: its runner, its state (script path and arguments) and the endpoint
# copy it starts live under Program Files, never under ProgramData.
$script:RetryAfterRebootProgramDir = Join-Path $env:ProgramFiles "WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin\RetryAfterReboot"
$script:RetryAfterRebootStatePath = Join-Path $script:RetryAfterRebootProgramDir "RetryAfterReboot.json.txt"
$script:RetryAfterRebootRunnerPath = Join-Path $script:RetryAfterRebootProgramDir "RetryAfterRebootRunner.ps1"
$script:RetryAfterRebootScriptPath = Join-Path $script:RetryAfterRebootProgramDir "SmartM365-Invoke-EndpointMigration-HybridJoin.ps1"
$script:RebootSafetyStatePath = Join-Path $StateDir "RebootSafety.json.txt"
$script:EndpointInstanceStatePath = Join-Path $StateDir "EndpointInstance.json.txt"
$script:EndpointInstanceMutexName = "Global\SmartM365_SEMT_HybridJoin_Endpoint"
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
$RunLogPath = Join-Path $LogsDir ("SEMT_HybridJoin_{0}_{1}.log" -f $ComputerName, $RunId)
if ([string]::IsNullOrWhiteSpace($CorrelationId)) { $CorrelationId = [guid]::NewGuid().ToString("N") }
elseif ($RetryAfterRebootTaskRun) { $CorrelationId = "{0}-R{1}" -f $CorrelationId,(Get-Date -Format "yyyyMMddHHmmss") }
$script:SemtAttemptStartUtc = (Get-Date).ToUniversalTime()
$script:SemtTargetGuardStatus = ""
$script:SemtLegacyActivityStatus = ""
$script:SemtLegacyActivityDetail = ""
$script:SemtCsvPath = Join-Path $DataRoot ("SEMT_HybridJoin_{0}.csv" -f $ComputerName)
$script:SemtTranscriptPath = ""


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

    if ([string]::IsNullOrWhiteSpace($NextAction) -and (Get-Command -Name Get-NextActionForStatus -ErrorAction SilentlyContinue)) {
        try { $NextAction = Get-NextActionForStatus -Status $Status } catch { $NextAction = "" }
    }
    try { Write-SemtHybridJoinAttemptResult -State "Completed" -Status $Status -ExitCode $ExitCode -NextAction $NextAction -Detail $Detail | Out-Null }
    catch { try { Write-RunLog ("Attempt result write failed: {0}" -f $_.Exception.Message) } catch { } }
    Write-Host ("FINAL_STATUS={0}; EXIT_CODE={1}; NEXT_ACTION={2}; CORRELATION_ID={3}; DETAIL={4}" -f $Status,$ExitCode,$NextAction,$CorrelationId,$Detail)
}

function Write-SemtHybridJoinAttemptResult {
    param(
        [Parameter(Mandatory=$true)][ValidateSet("Running","Completed")][string]$State,
        [string]$Status = "",
        [int]$ExitCode = 0,
        [string]$NextAction = "",
        [string]$Detail = ""
    )

    $data = [ordered]@{
        SchemaVersion = 2
        Product = "SmartEndpointMigrationToolkit"
        Campaign = $script:SemtCampaign
        ScriptVersion = $ScriptVersion
        RunId = $RunId
        CorrelationId = $CorrelationId
        ComputerName = $ComputerName
        State = $State
        StartTimeUtc = $script:SemtAttemptStartUtc.ToString("o")
        EndTimeUtc = $(if ($State -eq "Completed") { (Get-Date).ToUniversalTime().ToString("o") } else { "" })
        Status = $Status
        ExitCode = $(if ($State -eq "Completed") { [string]$ExitCode } else { "" })
        NextAction = $NextAction
        Detail = $Detail
        LogPath = $RunLogPath
        CsvPath = $script:SemtCsvPath
        TranscriptPath = $script:SemtTranscriptPath
        TargetGuardStatus = $script:SemtTargetGuardStatus
        LegacyActivityStatus = $script:SemtLegacyActivityStatus
        LegacyActivityDetail = $script:SemtLegacyActivityDetail
        RetryAfterRebootTaskRun = [bool]$RetryAfterRebootTaskRun
    }
    foreach ($name in @("gpUpdateTimedOut","gpUpdateStopDetail","finalGpUpdateTimedOut","finalGpUpdateStopDetail","mdmDiagnosticStatus","mdmDiagnosticZip","interactiveUserName","interactiveUserSid","mdmEnrollmentRecentEvidence","interactiveUserDomain","interactiveUserAccountName","interactiveUserAccountType","interactiveSessionName","interactiveSessionState","userIsUserAzureAD","userAzureAdPrt","userSessionIsNotRemote","intuneEnrolled","dsregAzureAdJoined","dsregDeviceAuthStatus","leaveAttempted","rebootAttempted","rebootReason")) {
        $value = Get-Variable -Name $name -Scope Script -ValueOnly -ErrorAction SilentlyContinue
        $data[($name.Substring(0,1).ToUpperInvariant() + $name.Substring(1))] = [string]$value
    }
    $data["Dsreg_DeviceId"] = [string](Get-Variable -Name dsregDeviceId -Scope Script -ValueOnly -ErrorAction SilentlyContinue)
    $data["Dsreg_TenantId"] = [string](Get-Variable -Name dsregTenantId -Scope Script -ValueOnly -ErrorAction SilentlyContinue)
    $data["RetryAfterRebootAction"] = [string]$script:RetryAfterRebootAction
    $data["RetryAfterRebootDetail"] = [string]$script:RetryAfterRebootDetail
    $data["RetryAfterRebootAttempt"] = [string]$script:RetryAfterRebootAttempt
    $data["RetryAfterRebootMaxAttempts"] = [string]$script:RetryAfterRebootMaxAttemptsResult
    $data["RetryAfterRebootTaskName"] = [string]$script:RetryAfterRebootTaskNameResult
    return (Write-SemtAttemptResult -CampaignDataRoot $DataRoot -Data $data)
}

function Stop-SemtEndpointEarly {
    param(
        [Parameter(Mandatory=$true)][string]$Status,
        [Parameter(Mandatory=$true)][int]$ExitCode,
        [Parameter(Mandatory=$true)][string]$Detail,
        [string]$NextAction = ""
    )

    try { Write-RunLog $Detail } catch { }
    Write-Host $Detail -ForegroundColor Yellow
    Write-FinalStatusLine -Status $Status -ExitCode $ExitCode -Detail $Detail -NextAction $NextAction
    try { Update-EndpointInstanceState -Status $Status -Force } catch { }
    Exit-SemtTargetOperationGuard -CampaignDataRoot $DataRoot
    if ($script:EndpointInstanceMutexAcquired -and $script:EndpointInstanceMutex) { try { $script:EndpointInstanceMutex.ReleaseMutex() } catch { } }
    if ($script:EndpointInstanceMutex) { try { $script:EndpointInstanceMutex.Dispose() } catch { } }
    exit $ExitCode
}

#region SEMT trusted program folder v1
# Scripts that run as SYSTEM (scheduled tasks, Intune runner, retry after reboot) are installed
# under %ProgramFiles%, which only administrators can write. ProgramData keeps only data (logs,
# state, results, setup media). The same region is copied unchanged into both endpoint scripts,
# the Intune installer and the Intune runner (a test checks that all copies are identical).
#
# A folder is trusted when it and every item below it are owned by SYSTEM, Administrators or
# TrustedInstaller and no other principal can write, append, delete, change permissions or take
# ownership. Permissions that cannot be read are never trusted. An existing folder that is not
# trusted is never repaired automatically: the operation stops (SEMT_PROGRAM_FOLDER_UNTRUSTED).

$script:SemtProgramTrustedSids = @('S-1-5-18','S-1-5-32-544','S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464')
$script:SemtProgramFolderSddl = 'O:BAD:PAI(A;OICI;FA;;;SY)(A;OICI;FA;;;BA)(A;OICI;0x1200a9;;;BU)'

function Get-SemtProgramRoot {
    param([Parameter(Mandatory = $true)][string]$Campaign, [string]$ProgramFilesPath = $env:ProgramFiles)
    if ([string]::IsNullOrWhiteSpace($ProgramFilesPath)) { throw 'SEMT_PROGRAM_FOLDER_UNTRUSTED: the Program Files folder is unknown.' }
    return (Join-Path $ProgramFilesPath ('WCH\SmartM365\SmartEndpointMigrationToolkit\{0}' -f $Campaign))
}

function Test-SemtTrustedProgramPath {
    # Read-only. State: Trusted, Untrusted (proven) or Unreadable (not proven either way).
    param([Parameter(Mandatory = $true)][string]$Path, [switch]$Recurse)
    # Write data/append/extended attributes, delete child, delete, WRITE_DAC, WRITE_OWNER and the
    # generic write/all bits that inherit-only entries may carry.
    $writeMask = 0x00000002 -bor 0x00000004 -bor 0x00000010 -bor 0x00000040 -bor 0x00010000 -bor 0x00040000 -bor 0x00080000 -bor 0x40000000 -bor 0x10000000
    try {
        $items = @(Get-Item -LiteralPath $Path -Force -ErrorAction Stop)
        if ($Recurse -and $items[0].PSIsContainer) { $items += @(Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction Stop) }
    }
    catch { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Program folder could not be read: {0}' -f $_.Exception.Message) } }
    foreach ($item in $items) {
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('Reparse point in the program folder: {0}' -f $item.FullName) } }
        try {
            $acl = Get-Acl -LiteralPath $item.FullName -ErrorAction Stop
            $owner = $acl.GetOwner([System.Security.Principal.SecurityIdentifier])
            $rules = @($acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))
        }
        catch { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Permissions of {0} could not be read: {1}' -f $item.FullName,$_.Exception.Message) } }
        if ($null -eq $owner) { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Owner of {0} could not be read.' -f $item.FullName) } }
        if ($script:SemtProgramTrustedSids -notcontains $owner.Value) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('Owner {0} of {1} is not SYSTEM, Administrators or TrustedInstaller.' -f $owner.Value,$item.FullName) } }
        foreach ($rule in $rules) {
            if ($rule.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow) { continue }
            if (([int64]$rule.FileSystemRights -band $writeMask) -eq 0) { continue }
            $sid = $rule.IdentityReference.Value
            if ($script:SemtProgramTrustedSids -contains $sid) { continue }
            # CREATOR OWNER only applies to items created later, by principals that can already write.
            if ($sid -eq 'S-1-3-0' -and ($rule.PropagationFlags -band [System.Security.AccessControl.PropagationFlags]::InheritOnly) -ne 0) { continue }
            return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('{0} can modify {1}.' -f $sid,$item.FullName) }
        }
    }
    return [pscustomobject]@{ Trusted = $true; State = 'Trusted'; Detail = '' }
}

function Set-SemtProgramFolderAcl {
    # Owner Administrators; protected DACL: SYSTEM and Administrators full control, Users read and
    # execute. Every other entry is removed. Requires an elevated or SYSTEM context.
    param([Parameter(Mandatory = $true)][string]$Path)
    $security = New-Object System.Security.AccessControl.DirectorySecurity
    $security.SetSecurityDescriptorSddlForm($script:SemtProgramFolderSddl)
    Set-Acl -LiteralPath $Path -AclObject $security -ErrorAction Stop
}

function Initialize-SemtTrustedProgramFolder {
    # Creates the folder chain below Program Files with a protected ACL, or verifies an existing
    # chain. Throws SEMT_PROGRAM_FOLDER_UNTRUSTED when any component is not trusted afterwards.
    param([Parameter(Mandatory = $true)][string]$Path, [string]$BasePath = $env:ProgramFiles)
    $base = [System.IO.Path]::GetFullPath($BasePath).TrimEnd('\')
    $target = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
    if (-not $target.StartsWith($base + '\', [System.StringComparison]::OrdinalIgnoreCase)) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} is not below {1}.' -f $target,$base) }
    $baseTrust = Test-SemtTrustedProgramPath -Path $base
    if (-not $baseTrust.Trusted) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $base,$baseTrust.State,$baseTrust.Detail) }
    $current = $base
    foreach ($segment in @($target.Substring($base.Length + 1).Split('\'))) {
        $current = Join-Path $current $segment
        if (-not (Test-Path -LiteralPath $current)) {
            New-Item -ItemType Directory -Path $current -ErrorAction Stop | Out-Null
            try { Set-SemtProgramFolderAcl -Path $current }
            catch {
                # A folder created here that cannot be protected is removed again.
                Remove-Item -LiteralPath $current -Force -ErrorAction SilentlyContinue
                throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: permissions of {0} could not be set: {1}' -f $current,$_.Exception.Message)
            }
        }
        elseif (-not (Test-Path -LiteralPath $current -PathType Container)) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} exists and is not a folder.' -f $current) }
        $trust = Test-SemtTrustedProgramPath -Path $current -Recurse:($current -eq $target)
        if (-not $trust.Trusted) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $current,$trust.State,$trust.Detail) }
    }
    return $target
}

function Install-SemtTrustedProgramFile {
    # Copies through a staging name inside the trusted folder, verifies the SHA256 against the
    # source, renames, then verifies the hash and the permissions of the installed file.
    param([Parameter(Mandatory = $true)][string]$Source, [Parameter(Mandatory = $true)][string]$Destination)
    $expected = (Get-FileHash -LiteralPath $Source -Algorithm SHA256 -ErrorAction Stop).Hash
    $folder = Split-Path -Path $Destination -Parent
    $staging = Join-Path $folder ('.{0}.{1}.tmp' -f (Split-Path -Path $Destination -Leaf),([guid]::NewGuid().ToString('N')))
    try {
        Copy-Item -LiteralPath $Source -Destination $staging -ErrorAction Stop
        if ((Get-FileHash -LiteralPath $staging -Algorithm SHA256).Hash -ne $expected) { throw ('SEMT_PROGRAM_FILE_HASH_MISMATCH: staged copy of {0} differs from the source.' -f $Source) }
        Move-Item -LiteralPath $staging -Destination $Destination -Force -ErrorAction Stop
    }
    finally { if (Test-Path -LiteralPath $staging) { Remove-Item -LiteralPath $staging -Force -ErrorAction SilentlyContinue } }
    $actual = (Get-FileHash -LiteralPath $Destination -Algorithm SHA256).Hash
    if ($actual -ne $expected) { throw ('SEMT_PROGRAM_FILE_HASH_MISMATCH: {0} Expected={1}; Actual={2}' -f $Destination,$expected,$actual) }
    $trust = Test-SemtTrustedProgramPath -Path $Destination
    if (-not $trust.Trusted) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $Destination,$trust.State,$trust.Detail) }
    return [pscustomobject]@{ Path = $Destination; Sha256 = $actual }
}

# ProgramData data folders (logs, state, results, setup media whose setup.exe runs as SYSTEM).
# Scope: only the SEMT components (SmartEndpointMigrationToolkit and the folders below it) are
# ever changed. The shared parents above it (WCH, WCH\SmartM365), used by other products and by
# the legacy caches reused in place, are verified read-only and never changed; when they are not
# protected, the operation stops. A read-only pass checks the whole scope before the first
# change: shared parents, SEMT components, the provenance of existing content and, when a
# permission replacement is planned, every reparse point and hard link below it.
# A permission replacement never makes existing content trusted: an item that another principal
# could modify, or any content of a SEMT component that was not protected, has no verified
# provenance; it is refused and left unchanged (SEMT_DATA_FOLDER_UNTRUSTED), never adopted.
$script:SemtDataProductFolderName = 'SmartEndpointMigrationToolkit'

function Get-SemtHardLinkCount {
    # Number of names of a file; -1 when it cannot be read.
    param([Parameter(Mandatory = $true)][string]$Path)
    $fsutil = Join-Path $env:SystemRoot 'System32\fsutil.exe'
    $output = @(& $fsutil hardlink list $Path 2>$null)
    if ($LASTEXITCODE -ne 0) { return -1 }
    return @($output | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count
}

function Assert-SemtNoLinkedItem {
    # Read-only. Before a DACL replacement propagates inheritance below Path: no reparse point
    # (never followed) and only files with a single name.
    param([Parameter(Mandatory = $true)][string]$Path)
    $pending = New-Object System.Collections.Generic.Queue[string]
    $pending.Enqueue([System.IO.Path]::GetFullPath($Path))
    while ($pending.Count -gt 0) {
        $folder = $pending.Dequeue()
        try { $children = @(Get-ChildItem -LiteralPath $folder -Force -ErrorAction Stop) }
        catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be listed: {1}; nothing was changed.' -f $folder,$_.Exception.Message) }
        foreach ($child in $children) {
            if (($child.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: reparse point {0}; nothing was changed.' -f $child.FullName) }
            if ($child.PSIsContainer) { $pending.Enqueue($child.FullName); continue }
            $links = Get-SemtHardLinkCount -Path $child.FullName
            if ($links -ne 1) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} has {1} hard link name(s); nothing was changed.' -f $child.FullName,$links) }
        }
    }
}

function Test-SemtCanonicalFolderAcl {
    # True when the folder already has exactly the protected ACL that this region sets.
    param([Parameter(Mandatory = $true)][string]$Path)
    try {
        $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
        if ($acl.GetOwner([System.Security.Principal.SecurityIdentifier]).Value -ne 'S-1-5-32-544') { return $false }
        if (-not $acl.AreAccessRulesProtected) { return $false }
        $rules = @($acl.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier]) | ForEach-Object { '{0}:{1}:{2}:{3}' -f $_.IdentityReference.Value,[int64]$_.FileSystemRights,[int]$_.InheritanceFlags,$_.AccessControlType })
        $expected = @('S-1-5-18:2032127:3:Allow','S-1-5-32-544:2032127:3:Allow','S-1-5-32-545:1179817:3:Allow')
        return (@($rules | Sort-Object) -join '|') -eq (@($expected | Sort-Object) -join '|')
    }
    catch { return $false }
}

function Test-SemtTrustedSharedParent {
    # Read-only; a shared parent is never changed. Trusted when its owner is SYSTEM,
    # Administrators or TrustedInstaller and no other principal can delete it, delete its
    # children, change its permissions or take ownership. Creating new items (other products) is
    # allowed: it cannot alter the existing SEMT folder below.
    param([Parameter(Mandatory = $true)][string]$Path)
    # Delete child, delete, WRITE_DAC, WRITE_OWNER and generic all.
    $mask = 0x00000040 -bor 0x00010000 -bor 0x00040000 -bor 0x00080000 -bor 0x10000000
    try {
        $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
        $acl = Get-Acl -LiteralPath $item.FullName -ErrorAction Stop
        $owner = $acl.GetOwner([System.Security.Principal.SecurityIdentifier])
        $rules = @($acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))
    }
    catch { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Shared folder {0} could not be read: {1}' -f $Path,$_.Exception.Message) } }
    if (-not $item.PSIsContainer) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('{0} is not a folder.' -f $item.FullName) } }
    if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('{0} is a reparse point.' -f $item.FullName) } }
    if ($null -eq $owner) { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Owner of {0} could not be read.' -f $item.FullName) } }
    if ($script:SemtProgramTrustedSids -notcontains $owner.Value) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('Owner {0} of {1} is not SYSTEM, Administrators or TrustedInstaller.' -f $owner.Value,$item.FullName) } }
    foreach ($rule in $rules) {
        if ($rule.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow) { continue }
        # Inherit-only entries apply to new children only, not to this folder.
        if (($rule.PropagationFlags -band [System.Security.AccessControl.PropagationFlags]::InheritOnly) -ne 0) { continue }
        if (([int64]$rule.FileSystemRights -band $mask) -eq 0) { continue }
        if ($script:SemtProgramTrustedSids -contains $rule.IdentityReference.Value) { continue }
        return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('{0} can delete, rename or change the permissions of {1}.' -f $rule.IdentityReference.Value,$item.FullName) }
    }
    return [pscustomobject]@{ Trusted = $true; State = 'Trusted'; Detail = '' }
}

function Assert-SemtDataFolderAnchor {
    # The data folder must be below the base inside a SmartEndpointMigrationToolkit folder, and
    # the base and its parents must not be reparse points.
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$BasePath)
    $base = [System.IO.Path]::GetFullPath($BasePath).TrimEnd('\')
    $target = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
    if (-not $target.StartsWith($base + '\', [System.StringComparison]::OrdinalIgnoreCase)) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} is not below {1}.' -f $target,$base) }
    $segments = @($target.Substring($base.Length + 1).Split('\'))
    $productIndex = -1
    for ($i = 0; $i -lt $segments.Count; $i++) { if ($segments[$i] -ieq $script:SemtDataProductFolderName) { $productIndex = $i; break } }
    if ($productIndex -lt 0 -or $productIndex -ge ($segments.Count - 1)) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} is not a campaign folder below {1}.' -f $target,$script:SemtDataProductFolderName) }
    $current = $base
    while (-not [string]::IsNullOrWhiteSpace($current)) {
        try { $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop }
        catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be read: {1}' -f $current,$_.Exception.Message) }
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} is a reparse point.' -f $current) }
        $parent = Split-Path -Path $current -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) { break }
        $current = $parent
    }
    return [pscustomobject]@{ Base = $base; Target = $target; Segments = $segments; ProductIndex = $productIndex }
}

function Get-SemtDataFolderPlan {
    # Read-only pass over the whole scope. Returns one step per component (Keep, Create or
    # Replace); throws SEMT_DATA_FOLDER_UNTRUSTED, with nothing changed, when the operation
    # cannot be done within the SEMT scope or when existing content has no verified provenance.
    param([Parameter(Mandatory = $true)]$Anchor)
    $plan = New-Object System.Collections.Generic.List[object]
    $current = $Anchor.Base
    $missing = $false
    for ($index = 0; $index -lt $Anchor.Segments.Count; $index++) {
        $current = Join-Path $current $Anchor.Segments[$index]
        $shared = $index -lt $Anchor.ProductIndex
        if ($missing -or -not (Test-Path -LiteralPath $current)) {
            $missing = $true
            $plan.Add([pscustomobject]@{ Path = $current; Shared = $shared; Action = 'Create'; WasUntrusted = $false })
            continue
        }
        try { $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop }
        catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be read: {1}; nothing was changed.' -f $current,$_.Exception.Message) }
        if (-not $item.PSIsContainer) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} exists and is not a folder; nothing was changed.' -f $current) }
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} is a reparse point; nothing was changed.' -f $current) }
        if ($shared) {
            $trust = Test-SemtTrustedSharedParent -Path $current
            if (-not $trust.Trusted) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: shared folder {0} ({1}) {2} It is used by other products and is never changed; nothing was changed.' -f $current,$trust.State,$trust.Detail) }
            $plan.Add([pscustomobject]@{ Path = $current; Shared = $true; Action = 'Keep'; WasUntrusted = $false })
            continue
        }
        $trust = Test-SemtTrustedProgramPath -Path $current
        if ($trust.State -eq 'Unreadable') { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} (Unreadable) {1}; nothing was changed.' -f $current,$trust.Detail) }
        if (-not $trust.Trusted) {
            # Content of a SEMT folder that was not protected has no verified provenance.
            try { $children = @(Get-ChildItem -LiteralPath $current -Force -ErrorAction Stop) }
            catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be listed: {1}; nothing was changed.' -f $current,$_.Exception.Message) }
            if ($children.Count -gt 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} was not protected ({1}) and is not empty; its content has no verified provenance and is not used; nothing was changed.' -f $current,$trust.Detail) }
            $plan.Add([pscustomobject]@{ Path = $current; Shared = $false; Action = 'Replace'; WasUntrusted = $true })
            continue
        }
        $action = if (Test-SemtCanonicalFolderAcl -Path $current) { 'Keep' } else { 'Replace' }
        $plan.Add([pscustomobject]@{ Path = $current; Shared = $false; Action = $action; WasUntrusted = $false })
    }
    # Existing content must already be trusted, item by item: below the data root and, when a
    # permission replacement is planned, everywhere it propagates to (other SEMT campaigns
    # included), so that the replacement never makes untrusted content look trusted.
    $firstReplace = @($plan | Where-Object { $_.Action -eq 'Replace' } | Select-Object -First 1)
    $existing = @($plan | Where-Object { $_.Action -ne 'Create' -and -not $_.Shared })
    $checkFrom = if ($firstReplace.Count -gt 0) { $firstReplace[0].Path } elseif (-not $missing) { $Anchor.Target } else { '' }
    if (-not [string]::IsNullOrWhiteSpace($checkFrom)) {
        $chain = @($existing | ForEach-Object { $_.Path })
        $start = [array]::IndexOf([string[]]$chain, [string]$checkFrom)
        for ($i = $start; $i -lt $chain.Count; $i++) {
            try { $children = @(Get-ChildItem -LiteralPath $chain[$i] -Force -ErrorAction Stop) }
            catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be listed: {1}; nothing was changed.' -f $chain[$i],$_.Exception.Message) }
            foreach ($child in $children) {
                if ($i -lt ($chain.Count - 1) -and $child.FullName -ieq $chain[$i + 1]) { continue }
                $trust = Test-SemtTrustedProgramPath -Path $child.FullName -Recurse
                if (-not $trust.Trusted) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: existing content has no verified provenance and is not used ({0}) {1}; nothing was changed.' -f $trust.State,$trust.Detail) }
            }
        }
    }
    # A DACL replacement propagates inheritance below the folder: the whole subtree of the first
    # replaced component is checked for reparse points and hard links before any change.
    if ($firstReplace.Count -gt 0) { Assert-SemtNoLinkedItem -Path $firstReplace[0].Path }
    return $plan
}

function Initialize-SemtTrustedDataFolder {
    # Protects the SEMT data folder chain below ProgramData and verifies it. Throws
    # SEMT_DATA_FOLDER_UNTRUSTED when it cannot be made trusted within the SEMT scope.
    param([Parameter(Mandatory = $true)][string]$Path, [string]$BasePath = $env:ProgramData)
    $anchor = Assert-SemtDataFolderAnchor -Path $Path -BasePath $BasePath
    $plan = Get-SemtDataFolderPlan -Anchor $anchor
    foreach ($step in $plan) {
        if ($step.Action -eq 'Create') {
            New-Item -ItemType Directory -Path $step.Path -ErrorAction Stop | Out-Null
            if ($step.Shared) {
                # Created with inherited permissions only: a shared folder is never changed.
                $trust = Test-SemtTrustedSharedParent -Path $step.Path
                if (-not $trust.Trusted) {
                    Remove-Item -LiteralPath $step.Path -Force -ErrorAction SilentlyContinue
                    throw ('SEMT_DATA_FOLDER_UNTRUSTED: shared folder {0} created here is not protected ({1}) {2}' -f $step.Path,$trust.State,$trust.Detail)
                }
                continue
            }
            try { Set-SemtProgramFolderAcl -Path $step.Path }
            catch {
                Remove-Item -LiteralPath $step.Path -Force -ErrorAction SilentlyContinue
                throw ('SEMT_DATA_FOLDER_UNTRUSTED: permissions of {0} could not be set: {1}' -f $step.Path,$_.Exception.Message)
            }
        }
        elseif ($step.Action -eq 'Replace') {
            try { Set-SemtProgramFolderAcl -Path $step.Path }
            catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: owner and permissions of {0} could not be replaced: {1}' -f $step.Path,$_.Exception.Message) }
        }
        else { continue }
        # A folder created here, or one that was not protected (empty in the read-only pass), must
        # still be empty once protected: nothing written meanwhile is adopted.
        if ($step.Action -eq 'Create' -or $step.WasUntrusted) {
            if (@(Get-ChildItem -LiteralPath $step.Path -Force -ErrorAction Stop).Count -gt 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: content appeared in {0} before it was protected; it is not used.' -f $step.Path) }
        }
        $trust = Test-SemtTrustedProgramPath -Path $step.Path
        if (-not $trust.Trusted) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $step.Path,$trust.State,$trust.Detail) }
    }
    $final = Test-SemtTrustedDataFolder -Path $anchor.Target -BasePath $anchor.Base
    if (-not $final.Trusted) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $anchor.Target,$final.State,$final.Detail) }
    return $anchor.Target
}

function Test-SemtTrustedDataFolder {
    # Read-only verification used before anything is run: shared parents protected, every SEMT
    # component trusted, and the data root trusted with everything below it.
    param([Parameter(Mandatory = $true)][string]$Path, [string]$BasePath = $env:ProgramData)
    try { $anchor = Assert-SemtDataFolderAnchor -Path $Path -BasePath $BasePath }
    catch { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = $_.Exception.Message } }
    $current = $anchor.Base
    for ($index = 0; $index -lt $anchor.Segments.Count; $index++) {
        $current = Join-Path $current $anchor.Segments[$index]
        if ($index -lt $anchor.ProductIndex) { $trust = Test-SemtTrustedSharedParent -Path $current }
        else { $trust = Test-SemtTrustedProgramPath -Path $current -Recurse:($current -eq $anchor.Target) }
        if (-not $trust.Trusted) { return $trust }
    }
    return [pscustomobject]@{ Trusted = $true; State = 'Trusted'; Detail = '' }
}
#endregion SEMT trusted program folder v1

#region SEMT shared target guards v1 (identical in every endpoint script; checked by Tests\Core)
# Target-side protections shared by every campaign of the Smart Endpoint Migration Toolkit.
# - One operation at a time on a target, whatever the campaign: named Global mutex held by the
#   running endpoint plus a diagnostic lease file. The mutex is released by Windows if the
#   process dies; the lease file only explains who holds it.
# - A campaign can leave a pending-operation marker that outlives its process (for example a
#   Windows Setup completion reboot). Other campaigns stop until the marker expires, the
#   computer has restarted, or the owning campaign clears it.
# - Previous toolkits (IntuneHybridJoinToolkit, Windows11UpgradeToolkit) are inspected
#   READ-ONLY. Installed files or old history never block; a real active operation blocks;
#   an unreadable state blocks with an explicit status.

$script:SemtTargetMutexName = 'Global\SmartM365_SEMT_TargetOperation'
$script:SemtTargetMutex = $null
$script:SemtTargetMutexAcquired = $false
$script:SemtLegacyHybridJoinRoot = 'C:\ProgramData\SmartM365\IntuneHybridJoinToolkit'
$script:SemtLegacyWindows11Root = 'C:\ProgramData\SmartM365\Windows11UpgradeToolkit'

function Get-SemtTargetProductRoot {
    param([Parameter(Mandatory = $true)][string]$CampaignDataRoot)
    return (Split-Path -Path $CampaignDataRoot -Parent)
}

function Get-SemtTargetLastBootUtc {
    try { return ((Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop).LastBootUpTime).ToUniversalTime() }
    catch { return $null }
}

function Write-SemtTargetJson {
    param([Parameter(Mandatory = $true)][string]$Path,[Parameter(Mandatory = $true)]$Data)
    $parent = Split-Path -Path $Path -Parent
    if (-not (Test-Path -LiteralPath $parent -PathType Container)) { New-Item -ItemType Directory -Path $parent -Force | Out-Null }
    $temp = Join-Path -Path $parent -ChildPath ('.{0}.{1}.tmp' -f (Split-Path -Path $Path -Leaf),[guid]::NewGuid().ToString('N'))
    try {
        $Data | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $temp -Encoding UTF8 -Force -ErrorAction Stop
        Move-Item -LiteralPath $temp -Destination $Path -Force -ErrorAction Stop
    }
    finally { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue }
}

function ConvertTo-SemtTargetJsonDateText {
    # Same dates on every engine: DateTime values created by PowerShell 7 JSON parsing become
    # round-trip text again (UTC/offset as UTC 'o', no zone kept without zone).
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [datetime]) {
        if ($Value.Kind -eq [System.DateTimeKind]::Unspecified) { return $Value.ToString('yyyy-MM-ddTHH:mm:ss.FFFFFFF', [Globalization.CultureInfo]::InvariantCulture) }
        return $Value.ToUniversalTime().ToString('o', [Globalization.CultureInfo]::InvariantCulture)
    }
    if ($Value -is [datetimeoffset]) { return $Value.UtcDateTime.ToString('o', [Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [System.Management.Automation.PSCustomObject]) {
        foreach ($property in @($Value.PSObject.Properties)) { $property.Value = ConvertTo-SemtTargetJsonDateText -Value $property.Value }
        return $Value
    }
    if ($Value -is [System.Collections.IList]) {
        for ($index = 0; $index -lt $Value.Count; $index++) { $Value[$index] = ConvertTo-SemtTargetJsonDateText -Value $Value[$index] }
        return ,$Value
    }
    return $Value
}
function Read-SemtTargetJson {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    # Keep ISO dates as text on PowerShell 7 too, so UTC round-trip comparisons stay exact.
    if ((Get-Command -Name ConvertFrom-Json).Parameters.ContainsKey('DateKind')) { return ($raw | ConvertFrom-Json -DateKind String -ErrorAction Stop) }
    return (ConvertTo-SemtTargetJsonDateText -Value ($raw | ConvertFrom-Json -ErrorAction Stop))
}

function Enter-SemtTargetOperationGuard {
    param(
        [Parameter(Mandatory = $true)][string]$Campaign,
        [Parameter(Mandatory = $true)][string]$CampaignDataRoot,
        [Parameter(Mandatory = $true)][string]$RunId,
        [Parameter(Mandatory = $true)][string]$CorrelationId
    )

    $productRoot = Get-SemtTargetProductRoot -CampaignDataRoot $CampaignDataRoot
    $leasePath = Join-Path -Path $productRoot -ChildPath 'State\TargetOperation.json.txt'
    $pendingPath = Join-Path -Path $productRoot -ChildPath 'State\TargetPendingOperation.json.txt'
    try {
        $script:SemtTargetMutex = New-Object System.Threading.Mutex($false, $script:SemtTargetMutexName)
        try { $script:SemtTargetMutexAcquired = $script:SemtTargetMutex.WaitOne(0) }
        catch [System.Threading.AbandonedMutexException] { $script:SemtTargetMutexAcquired = $true }
    }
    catch { $script:SemtTargetMutexAcquired = $false }
    if (-not $script:SemtTargetMutexAcquired) {
        $holder = ''
        try {
            $lease = Read-SemtTargetJson -Path $leasePath
            if ($lease) { $holder = ('Campaign={0}; RunId={1}; ProcessId={2}; StartedUtc={3}' -f $lease.Campaign,$lease.RunId,$lease.ProcessId,$lease.StartedUtc) }
        }
        catch { $holder = 'Lease unreadable.' }
        if ($script:SemtTargetMutex) { try { $script:SemtTargetMutex.Dispose() } catch { } ; $script:SemtTargetMutex = $null }
        return [pscustomobject]@{ Acquired = $false; Status = 'TARGET_OPERATION_ACTIVE'; NextAction = 'WAIT_ACTIVE_TARGET_OPERATION'; Detail = ("Another Smart Endpoint Migration Toolkit operation is active on this computer. {0}" -f $holder) }
    }

    try {
        $pending = Read-SemtTargetJson -Path $pendingPath
        if ($pending -and [string]$pending.Campaign -ne $Campaign) {
            $expires = [datetime]::MinValue
            $expiresOk = [datetime]::TryParse([string]$pending.ExpiresUtc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$expires)
            $bootAtWrite = [datetime]::MinValue
            $bootOk = [datetime]::TryParse([string]$pending.LastBootUtc, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$bootAtWrite)
            $currentBoot = Get-SemtTargetLastBootUtc
            $rebootedSince = ($bootOk -and $null -ne $currentBoot -and [math]::Abs(($currentBoot - $bootAtWrite.ToUniversalTime()).TotalMinutes) -gt 2 -and [bool]$pending.ClearedByReboot)
            if ($expiresOk -and (Get-Date).ToUniversalTime() -lt $expires.ToUniversalTime() -and -not $rebootedSince) {
                Exit-SemtTargetOperationGuard -CampaignDataRoot $CampaignDataRoot
                return [pscustomobject]@{ Acquired = $false; Status = 'TARGET_PENDING_OPERATION'; NextAction = 'WAIT_OTHER_CAMPAIGN_PENDING_OPERATION'; Detail = ("Campaign {0} left a pending operation on this computer: {1}. Expires={2}." -f $pending.Campaign,$pending.Reason,$pending.ExpiresUtc) }
            }
        }
    }
    catch {
        Exit-SemtTargetOperationGuard -CampaignDataRoot $CampaignDataRoot
        return [pscustomobject]@{ Acquired = $false; Status = 'TARGET_STATE_UNREADABLE'; NextAction = 'VERIFY_TARGET_STATE_FILES'; Detail = ("Pending-operation marker could not be read: {0}" -f $_.Exception.Message) }
    }

    $processStart = ''
    try { $processStart = (Get-Process -Id $PID -ErrorAction Stop).StartTime.ToUniversalTime().ToString('o') } catch { $processStart = '' }
    try {
        Write-SemtTargetJson -Path $leasePath -Data ([pscustomobject]@{ SchemaVersion = 1; Campaign = $Campaign; RunId = $RunId; CorrelationId = $CorrelationId; ProcessId = $PID; ProcessStartUtc = $processStart; StartedUtc = (Get-Date).ToUniversalTime().ToString('o') })
    }
    catch { }
    return [pscustomobject]@{ Acquired = $true; Status = 'TARGET_OPERATION_ACQUIRED'; NextAction = ''; Detail = 'Target operation guard acquired.' }
}

function Exit-SemtTargetOperationGuard {
    param([Parameter(Mandatory = $true)][string]$CampaignDataRoot)

    if ($script:SemtTargetMutexAcquired) {
        $leasePath = Join-Path -Path (Get-SemtTargetProductRoot -CampaignDataRoot $CampaignDataRoot) -ChildPath 'State\TargetOperation.json.txt'
        try {
            $lease = Read-SemtTargetJson -Path $leasePath
            if ($lease -and [int]$lease.ProcessId -eq $PID) { Remove-Item -LiteralPath $leasePath -Force -ErrorAction SilentlyContinue }
        }
        catch { }
        if ($script:SemtTargetMutex) { try { $script:SemtTargetMutex.ReleaseMutex() } catch { } }
    }
    if ($script:SemtTargetMutex) { try { $script:SemtTargetMutex.Dispose() } catch { } }
    $script:SemtTargetMutex = $null
    $script:SemtTargetMutexAcquired = $false
}

function Set-SemtTargetPendingOperation {
    param(
        [Parameter(Mandatory = $true)][string]$Campaign,
        [Parameter(Mandatory = $true)][string]$CampaignDataRoot,
        [Parameter(Mandatory = $true)][string]$Reason,
        [ValidateRange(1, 168)][int]$Hours = 48,
        [bool]$ClearedByReboot = $true
    )

    $path = Join-Path -Path (Get-SemtTargetProductRoot -CampaignDataRoot $CampaignDataRoot) -ChildPath 'State\TargetPendingOperation.json.txt'
    $boot = Get-SemtTargetLastBootUtc
    Write-SemtTargetJson -Path $path -Data ([pscustomobject]@{ SchemaVersion = 1; Campaign = $Campaign; Reason = $Reason; SinceUtc = (Get-Date).ToUniversalTime().ToString('o'); ExpiresUtc = (Get-Date).ToUniversalTime().AddHours($Hours).ToString('o'); LastBootUtc = $(if ($boot) { $boot.ToString('o') } else { '' }); ClearedByReboot = $ClearedByReboot })
}

function Clear-SemtTargetPendingOperation {
    param(
        [Parameter(Mandatory = $true)][string]$Campaign,
        [Parameter(Mandatory = $true)][string]$CampaignDataRoot
    )

    $path = Join-Path -Path (Get-SemtTargetProductRoot -CampaignDataRoot $CampaignDataRoot) -ChildPath 'State\TargetPendingOperation.json.txt'
    try {
        $pending = Read-SemtTargetJson -Path $path
        if ($pending -and [string]$pending.Campaign -eq $Campaign) { Remove-Item -LiteralPath $path -Force -ErrorAction Stop; return $true }
    }
    catch { }
    return $false
}

function Test-SemtWindowsSetupActive {
    $signals = New-Object System.Collections.Generic.List[string]
    try {
        $setupKey = Get-ItemProperty -LiteralPath 'HKLM:\SYSTEM\Setup' -ErrorAction Stop
        foreach ($name in @('SystemSetupInProgress','UpgradeInProgress')) {
            $property = $setupKey.PSObject.Properties[$name]
            if ($property -and [string]$property.Value -ne '0') { $signals.Add(('{0}={1}' -f $name,$property.Value)) }
        }
    }
    catch { }
    foreach ($processName in @('SetupHost','SetupPrep','setup')) {
        foreach ($process in @(Get-Process -Name $processName -ErrorAction SilentlyContinue)) {
            $path = ''
            try { $path = [string]$process.Path } catch { $path = '' }
            if ($processName -ne 'setup' -or $path -match '(?i)\\(SetupMedia|\$WINDOWS\.~BT|sources)\\' -or [string]::IsNullOrWhiteSpace($path)) { $signals.Add(('{0}:{1}' -f $process.ProcessName,$process.Id)) }
        }
    }
    return [pscustomobject]@{ Active = ($signals.Count -gt 0); Detail = ($signals -join '; ') }
}

function Get-SemtLegacyEndpointActivity {
    <#
    READ-ONLY inspection of the previous toolkits on this computer. It never writes, deletes,
    takes a lock or waits on a legacy mutex.
    Statuses: LEGACY_ABSENT, LEGACY_INSTALLED_IDLE and LEGACY_INTUNE_PACKAGE_ASSIGNED (not blocking;
    media owned by a legacy Intune package is reused in place), LEGACY_ENDPOINT_RUNNING, LEGACY_SETUP_IN_PROGRESS,
    LEGACY_COPY_IN_PROGRESS, LEGACY_RETRY_AFTER_REBOOT_PENDING (blocking),
    LEGACY_STATE_UNREADABLE (blocking).
    #>
    param([Parameter(Mandatory = $true)][string]$Campaign)

    $signals = New-Object System.Collections.Generic.List[string]
    $active = New-Object System.Collections.Generic.List[string]
    $installed = $false
    try {
        foreach ($root in @($script:SemtLegacyHybridJoinRoot,$script:SemtLegacyWindows11Root)) {
            if (Test-Path -LiteralPath $root -PathType Container) { $installed = $true; $signals.Add(('Folder={0}' -f $root)) }
        }

        $legacyScripts = '(?i)(SmartM365-Invoke-Windows11UpgradeRepair\.ps1|SmartM365-Invoke-IntuneHybridJoinRepair\.ps1|Run-IntuneUpgrade\.ps1|\\(IntuneHybridJoinToolkit|Windows11UpgradeToolkit)\\State\\RetryAfterRebootRunner\.ps1)'
        $legacyMedia = [regex]::Escape((Join-Path -Path $script:SemtLegacyWindows11Root -ChildPath 'SetupMedia'))
        foreach ($process in @(Get-CimInstance -ClassName Win32_Process -ErrorAction Stop)) {
            $name = [string]$process.Name
            $commandLine = [string]$process.CommandLine
            $executable = [string]$process.ExecutablePath
            if ($name -match '(?i)^(powershell|pwsh)\.exe$' -and $commandLine -match $legacyScripts) { $active.Add(('LEGACY_ENDPOINT_RUNNING|Process {0}:{1}' -f $name,$process.ProcessId)) }
            elseif ($name -match '(?i)^(setup|setuphost|setupprep)\.exe$' -and ($executable -match $legacyMedia -or $commandLine -match $legacyMedia)) { $active.Add(('LEGACY_SETUP_IN_PROGRESS|Process {0}:{1}' -f $name,$process.ProcessId)) }
            elseif ($name -match '(?i)^robocopy\.exe$' -and $commandLine -match $legacyMedia) { $active.Add(('LEGACY_COPY_IN_PROGRESS|Process {0}:{1}' -f $name,$process.ProcessId)) }
        }

        $legacyMutex = $null
        if ([System.Threading.Mutex]::TryOpenExisting('Global\SmartM365_IntuneHybridJoinToolkit_Endpoint', [ref]$legacyMutex)) {
            $active.Add('LEGACY_ENDPOINT_RUNNING|Legacy Hybrid Join endpoint mutex is held')
            try { $legacyMutex.Dispose() } catch { }
        }

        $taskNames = @(
            @{ Path = '\'; Name = 'SmartM365-Windows11UpgradeToolkit-RetryAfterReboot'; State = (Join-Path -Path $script:SemtLegacyWindows11Root -ChildPath 'State\RetryAfterReboot.json') },
            @{ Path = '\'; Name = 'SmartM365-IntuneHybridJoinToolkit-RetryAfterReboot'; State = (Join-Path -Path $script:SemtLegacyHybridJoinRoot -ChildPath 'State\RetryAfterReboot.json') },
            @{ Path = '\'; Name = 'SmartM365 Windows 11 Upgrade Toolkit - Intune'; State = '' }
        )
        foreach ($taskName in $taskNames) {
            $task = Get-ScheduledTask -TaskPath $taskName.Path -TaskName $taskName.Name -ErrorAction SilentlyContinue
            if (-not $task) { continue }
            $installed = $true
            $signals.Add(('Task={0}; State={1}' -f $taskName.Name,$task.State))
            if ([string]$task.State -eq 'Running') { $active.Add(('LEGACY_ENDPOINT_RUNNING|Task {0} is running' -f $taskName.Name)) }
            elseif (-not [string]::IsNullOrWhiteSpace($taskName.State) -and (Test-Path -LiteralPath $taskName.State -PathType Leaf)) { $active.Add(('LEGACY_RETRY_AFTER_REBOOT_PENDING|Task {0} will resume a legacy run at startup' -f $taskName.Name)) }
        }
        foreach ($task in @(Get-ScheduledTask -TaskPath '\SmartM365\IntuneHybridJoinToolkit\' -ErrorAction SilentlyContinue)) {
            $installed = $true
            if ([string]$task.State -eq 'Running') { $active.Add(('LEGACY_ENDPOINT_RUNNING|Task {0} is running' -f $task.TaskName)) }
        }

        $lockRoot = Join-Path -Path $script:SemtLegacyWindows11Root -ChildPath 'Locks'
        if (Test-Path -LiteralPath $lockRoot -PathType Container) {
            foreach ($lockFile in @(Get-ChildItem -LiteralPath $lockRoot -Filter 'lock.json' -File -Recurse -ErrorAction Stop)) {
                $lock = Read-SemtTargetJson -Path $lockFile.FullName
                $lockPid = 0
                if ($lock -and $lock.PSObject.Properties['PID']) { [void][int]::TryParse([string]$lock.PID, [ref]$lockPid) }
                if ($lockPid -gt 0 -and (Get-Process -Id $lockPid -ErrorAction SilentlyContinue)) { $active.Add(('LEGACY_COPY_IN_PROGRESS|Legacy media lock {0} held by PID {1}' -f $lockFile.Directory.Name,$lockPid)) }
            }
        }

        $packageRoot = 'SOFTWARE\SmartM365\Windows11UpgradeToolkit\IntunePackages'
        $baseKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
        try {
            $packages = $baseKey.OpenSubKey($packageRoot, $false)
            if ($packages) {
                try {
                    foreach ($packageName in @($packages.GetSubKeyNames())) {
                        $package = $packages.OpenSubKey($packageName, $false)
                        if (-not $package) { continue }
                        try {
                            $installed = $true
                            $installState = [string]$package.GetValue('InstallState', '')
                            $signals.Add(('IntunePackage={0}; InstallState={1}; Cache={2}' -f $packageName,$installState,[string]$package.GetValue('SetupCacheFolder', '')))
                            if ($installState -in @('Installed','RepairRequired')) { $active.Add(('LEGACY_INTUNE_PACKAGE_ASSIGNED|Legacy Intune Win32 package {0} is {1}' -f $packageName,$installState)) }
                        }
                        finally { $package.Dispose() }
                    }
                }
                finally { $packages.Dispose() }
            }
        }
        finally { $baseKey.Dispose() }
    }
    catch {
        return [pscustomobject]@{ Status = 'LEGACY_STATE_UNREADABLE'; Blocking = $true; NextAction = 'VERIFY_LEGACY_TOOLKIT_STATE'; Detail = ('Legacy toolkit state could not be read: {0}' -f $_.Exception.Message) }
    }

    $priority = @('LEGACY_SETUP_IN_PROGRESS','LEGACY_COPY_IN_PROGRESS','LEGACY_ENDPOINT_RUNNING','LEGACY_RETRY_AFTER_REBOOT_PENDING','LEGACY_INTUNE_PACKAGE_ASSIGNED')
    foreach ($candidate in $priority) {
        $found = @($active | Where-Object { $_.StartsWith($candidate + '|') })
        if ($found.Count -eq 0) { continue }
        $blocking = ($candidate -ne 'LEGACY_INTUNE_PACKAGE_ASSIGNED')
        $nextAction = switch ($candidate) {
            'LEGACY_SETUP_IN_PROGRESS' { 'WAIT_LEGACY_WINDOWS_SETUP' }
            'LEGACY_COPY_IN_PROGRESS' { 'WAIT_LEGACY_MEDIA_COPY' }
            'LEGACY_ENDPOINT_RUNNING' { 'WAIT_LEGACY_TOOLKIT_OPERATION' }
            'LEGACY_RETRY_AFTER_REBOOT_PENDING' { 'WAIT_OR_REMOVE_LEGACY_RETRY_TASK' }
            default { 'REVIEW_LEGACY_INTUNE_ASSIGNMENT' }
        }
        return [pscustomobject]@{ Status = $candidate; Blocking = $blocking; NextAction = $nextAction; Detail = (($found | ForEach-Object { $_.Substring($candidate.Length + 1) }) -join '; ') }
    }
    if ($installed) { return [pscustomobject]@{ Status = 'LEGACY_INSTALLED_IDLE'; Blocking = $false; NextAction = ''; Detail = ($signals -join '; ') } }
    return [pscustomobject]@{ Status = 'LEGACY_ABSENT'; Blocking = $false; NextAction = ''; Detail = '' }
}

function Write-SemtAttemptResult {
    <#
    Writes <DataRoot>\State\Results\<CorrelationId>.json.txt. The technician worker only trusts a
    result whose CorrelationId matches its own attempt. Keeps the newest 50 result files.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$CampaignDataRoot,
        [Parameter(Mandatory = $true)][System.Collections.IDictionary]$Data
    )

    $resultsRoot = Join-Path -Path $CampaignDataRoot -ChildPath 'State\Results'
    $path = Join-Path -Path $resultsRoot -ChildPath ('{0}.json.txt' -f [string]$Data['CorrelationId'])
    Write-SemtTargetJson -Path $path -Data ([pscustomobject]$Data)
    try {
        @(Get-ChildItem -LiteralPath $resultsRoot -Filter '*.json.txt' -File -ErrorAction Stop | Sort-Object LastWriteTimeUtc -Descending | Select-Object -Skip 50) | ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
    }
    catch { }
    return $path
}
#endregion SEMT shared target guards v1

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

# SEMT data folder boundary: before anything is written there (setup media included), the SEMT
# components of the data folder chain get owner Administrators and a protected DACL after a
# read-only check of the whole scope; shared parents are verified, never changed; existing content
# without verified provenance is refused. Failure stops the endpoint without writing under the
# data root.
try { [void](Initialize-SemtTrustedDataFolder -Path $DataRoot) }
catch {
    Write-Output ('FINAL_STATUS=DATA_FOLDER_UNTRUSTED; EXIT_CODE=1; NEXT_ACTION=VERIFY_DATA_FOLDER_PERMISSIONS; CORRELATION_ID={0}; DETAIL={1}' -f [string]$CorrelationId,($_.Exception.Message -replace '[\r\n]+',' '))
    exit 1
}
# The subfolders inherit the protected DACL of the data root.
foreach ($d in @($OutputDir, $TranscriptDir, $LogsDir, $StateDir)) {
    if (-not (Test-Path $d)) {
        New-Item -ItemType Directory -Path $d -Force | Out-Null
    }
}

try {
    $script:EndpointInstanceMutex = New-Object System.Threading.Mutex($false,$script:EndpointInstanceMutexName)
    try { $script:EndpointInstanceMutexAcquired = $script:EndpointInstanceMutex.WaitOne(0) }
    catch [System.Threading.AbandonedMutexException] { $script:EndpointInstanceMutexAcquired = $true }
}
catch { $script:EndpointInstanceMutexAcquired = $false }

if (-not $script:EndpointInstanceMutexAcquired) {
    $activeMessage = "Another endpoint repair instance is already active. Mutex=$($script:EndpointInstanceMutexName). IgnoreRunGuard cannot bypass an active endpoint instance."
    Stop-SemtEndpointEarly -Status "ENDPOINT_RUN_ACTIVE" -ExitCode 3 -Detail $activeMessage -NextAction "WAIT_ACTIVE_ENDPOINT_RUN"
}

try { Write-SemtHybridJoinAttemptResult -State "Running" -Status "RUNNING" | Out-Null } catch { Write-RunLog ("Attempt start marker write failed: {0}" -f $_.Exception.Message) }

$semtGuard = Enter-SemtTargetOperationGuard -Campaign $script:SemtCampaign -CampaignDataRoot $DataRoot -RunId $RunId -CorrelationId $CorrelationId
$script:SemtTargetGuardStatus = $semtGuard.Status
if (-not $semtGuard.Acquired) {
    Stop-SemtEndpointEarly -Status $semtGuard.Status -ExitCode 3 -Detail $semtGuard.Detail -NextAction $semtGuard.NextAction
}

$semtSetupActivity = Test-SemtWindowsSetupActive
if ($semtSetupActivity.Active) {
    Stop-SemtEndpointEarly -Status "TARGET_WINDOWS_SETUP_ACTIVE" -ExitCode 3 -Detail ("Windows Setup is active on this computer; Hybrid Join repair is postponed. Signals={0}" -f $semtSetupActivity.Detail) -NextAction "WAIT_WINDOWS_SETUP_COMPLETION"
}

$semtLegacy = Get-SemtLegacyEndpointActivity -Campaign $script:SemtCampaign
$script:SemtLegacyActivityStatus = $semtLegacy.Status
$script:SemtLegacyActivityDetail = $semtLegacy.Detail
if ($semtLegacy.Blocking -and $LegacyActivityPolicy -eq "Block") {
    Stop-SemtEndpointEarly -Status $semtLegacy.Status -ExitCode 3 -Detail ("Previous toolkit activity detected (read-only check): {0}" -f $semtLegacy.Detail) -NextAction $semtLegacy.NextAction
}
elseif ($semtLegacy.Status -ne "LEGACY_ABSENT") {
    Write-RunLog ("Previous toolkit inspection: Status={0}; Blocking={1}; Policy={2}; Detail={3}" -f $semtLegacy.Status,$semtLegacy.Blocking,$LegacyActivityPolicy,$semtLegacy.Detail)
}

try { Update-EndpointInstanceState -Force } catch { Write-RunLog ("Endpoint instance metadata write failed; mutex protection remains active. Error={0}" -f $_.Exception.Message) }

Write-Host "Smart Endpoint Migration Toolkit - Hybrid Join endpoint version $ScriptVersion"
Write-RunLog "Script start. Version=$ScriptVersion. RunId=$RunId. CorrelationId=$CorrelationId. DataRoot=$DataRoot. LegacyActivityPolicy=$LegacyActivityPolicy. AllowDsregLeave=$([bool]$AllowDsregLeave). AllowRemoveNonIntuneMdmEnrollment=$([bool]$AllowRemoveNonIntuneMdmEnrollment). AllowRemoveStaleIntuneEnrollment=$([bool]$AllowRemoveStaleIntuneEnrollment). SkipVirtualMachines=$([bool]$SkipVirtualMachines). AuditOnly=$([bool]$AuditOnly). EntraHybridPending=$([bool]$EntraHybridPending). IgnoreRunGuard=$([bool]$IgnoreRunGuard). RetryAfterRebootTaskRun=$([bool]$RetryAfterRebootTaskRun)."
Invoke-OldEvidenceCleanup -Paths @($LogsDir, $OutputDir, $TranscriptDir) -RetentionDays $CleanupRetentionDays

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

# A virtual machine skipped by -SkipVirtualMachines is detected before the run guard: the skip does
# nothing on the computer, so it neither applies nor writes the 12-hour guard (LastRun.json.txt). Writing
# it left Status=RUNNING and blocked the next Runs with RUN_GUARD_ACTIVE.
$script:SkipAsVirtualMachine = $false
if ($SkipVirtualMachines) {
    $earlyComputerSystem = Get-ComputerSystemSummary
    $script:SkipAsVirtualMachine = ($earlyComputerSystem.Success -and $earlyComputerSystem.IsVirtualMachine -eq $true)
}

# ============================
# 12-hour local run guard
# ============================
try {
    if ($script:SkipAsVirtualMachine) {
        Write-RunLog "$RunGuardHours-hour run guard not applied: virtual machine skipped by -SkipVirtualMachines."
    }
    elseif ($IgnoreRunGuard -or $RetryAfterRebootTaskRun) {
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
            Stop-SemtEndpointEarly -Status "RUN_GUARD_ACTIVE" -ExitCode 3 -Detail $guardMessage -NextAction "WAIT_RUN_GUARD"
        }
    }
}
catch {
    Write-RunLog ("Run guard read failed; continuing. Error={0}" -f $_.Exception.Message)
}

try {
    if (-not $script:SkipAsVirtualMachine) {
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
$GpUpdateTimeoutSeconds = 180
$UserTaskFolderPath = "\WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin"
$UserStatusTaskName = "\WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin\SEMT-HJ-UserDsregStatus"
$UserRefreshPrtTaskName = "\WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin\SEMT-HJ-UserRefreshPrt"
$UserIntuneAutoEnrollTaskName = "\WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin\SEMT-HJ-UserMdmAutoEnroll"
$UserNextLogonAutoEnrollTaskName = "\WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin\SEMT-HJ-RunUserAutoEnrollAtLogon"

Write-RunLog ("Retry policy: SleepMinutes={0}; MaxRetries={1}; Mode=LocalWait" -f $RetrySleepMinutes, $RetryMaxRetries)
Write-RunLog ("Intune retry policy: SleepMinutes={0}; MaxRetries={1}; Mode=LocalWait" -f $IntuneRetrySleepMinutes, $IntuneRetryMaxRetries)
Write-RunLog ("User task policy: Folder={0}; StatusTask={1}; RefreshPrtTask={2}; IntuneAutoEnrollTask={3}; NextLogonTask={4}; WaitSeconds={5}; RefreshWaitSeconds={6}" -f $UserTaskFolderPath, $UserStatusTaskName, $UserRefreshPrtTaskName, $UserIntuneAutoEnrollTaskName, $UserNextLogonAutoEnrollTaskName, $UserTaskWaitSeconds, $UserPrtRetryWaitSeconds)
Write-RunLog "DNS policy: FlushDNS=True"
Write-RunLog ("GPUpdate policy: Force=True; TimeoutSeconds={0}; PostWaitSeconds={1}" -f $GpUpdateTimeoutSeconds,$GpUpdateWaitSeconds)
Write-RunLog ("Non-Intune MDM removal policy: AllowRemoveNonIntuneMdmEnrollment={0}" -f [bool]$AllowRemoveNonIntuneMdmEnrollment)

# Transcript (best effort)
$TranscriptStarted = $false
$TranscriptFile = Join-Path $TranscriptDir ("Transcript_SEMT_HybridJoin_{0}_{1}.txt" -f $ComputerName, (Get-Date -Format "yyyyMMdd_HHmmss"))
$script:SemtTranscriptPath = $TranscriptFile
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

function Get-MdmEnrollmentClassification {
    param(
        [string]$ProviderID,
        [string]$DiscoveryServiceFullURL,
        [bool]$OmadmAccountPresent,
        [bool]$EnterpriseMgmtTaskPresent
    )

    $providerIdValue = ([string]$ProviderID).Trim()
    $discoveryUrlValue = ([string]$DiscoveryServiceFullURL).Trim()
    $internalProviderIds = @("Deploy Authority", "Cloud Authority", "Local Authority")
    $hasProviderId = -not [string]::IsNullOrWhiteSpace($providerIdValue)
    $hasDiscoveryUrl = -not [string]::IsNullOrWhiteSpace($discoveryUrlValue)
    $isInternalProvider = $hasProviderId -and ($internalProviderIds -contains $providerIdValue)
    $isIntuneProvider = ($providerIdValue -eq "MS DM Server")
    $isIntuneDiscovery = ($discoveryUrlValue -match "(?i)enrollment\.manage\.microsoft\.com")
    $isManagementBridge = ($providerIdValue -ieq "WMI_Bridge_SCCM_Server")
    $isExternalProvider = $hasProviderId -and -not $isInternalProvider -and -not $isIntuneProvider -and -not $isManagementBridge
    $isExternalDiscovery = $hasDiscoveryUrl -and -not $isIntuneDiscovery
    $hasStrongOmaDmEvidence = ($isExternalDiscovery -or $OmadmAccountPresent -or $EnterpriseMgmtTaskPresent)
    $isIntuneCandidate = ($isIntuneProvider -or $isIntuneDiscovery)
    $isIntuneEnrollment = $isIntuneProvider
    $isNonIntuneMdmEnrollment = (-not $isInternalProvider -and -not $isManagementBridge -and
        ($isExternalDiscovery -or ($isExternalProvider -and ($OmadmAccountPresent -or $EnterpriseMgmtTaskPresent))))

    [pscustomobject]@{
        HasProviderId = $hasProviderId
        HasDiscoveryUrl = $hasDiscoveryUrl
        IsInternalProvider = $isInternalProvider
        IsIntuneCandidate = $isIntuneCandidate
        IsIntuneEnrollment = $isIntuneEnrollment
        IsManagementBridge = $isManagementBridge
        IsExternalProvider = $isExternalProvider
        IsExternalDiscovery = $isExternalDiscovery
        HasStrongOmaDmEvidence = $hasStrongOmaDmEvidence
        IsNonIntuneMdmEnrollment = $isNonIntuneMdmEnrollment
        IsMdmEnrollment = ($isIntuneEnrollment -or $isNonIntuneMdmEnrollment)
    }
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
                ManagementBridgeDetected = $false
                ManagementBridgeEnrollmentIds = ""
                ManagementBridgeProviderIds = ""
                ManagementBridgeDetails = ""
            }
        }

        $entries = @()
        $ignoredEntries = @()
        $unconfirmedIntuneEntries = @()
        $managementBridgeEntries = @()
        foreach ($subName in $enrollKey.GetSubKeyNames()) {
            try {
                $sub = $enrollKey.OpenSubKey($subName, $false)
                if (-not $sub) { continue }
                $providerId = [string]$sub.GetValue("ProviderID", "")
                $discoveryServiceFullUrl = [string]$sub.GetValue("DiscoveryServiceFullURL", "")
                $enrollmentType = [string]$sub.GetValue("EnrollmentType", "")
                $upn = [string]$sub.GetValue("UPN", "")
                $aadResourceId = [string]$sub.GetValue("AADResourceID", "")

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

                $classification = Get-MdmEnrollmentClassification `
                    -ProviderID $providerId `
                    -DiscoveryServiceFullURL $discoveryServiceFullUrl `
                    -OmadmAccountPresent $omadmAccountPresent `
                    -EnterpriseMgmtTaskPresent $enterpriseMgmtTaskPresent

                $entry = [PSCustomObject]@{
                    EnrollmentId = $subName
                    ProviderID = $providerId
                    DiscoveryServiceFullURL = $discoveryServiceFullUrl
                    EnrollmentType = $enrollmentType
                    UPN = $upn
                    AADResourceID = $aadResourceId
                    Evidence = $evidenceText
                    IsIntune = $classification.IsIntuneEnrollment
                    IsNonIntuneMdm = $classification.IsNonIntuneMdmEnrollment
                    IsManagementBridge = $classification.IsManagementBridge
                    HasStrongOmaDmEvidence = $classification.HasStrongOmaDmEvidence
                }

                if ($classification.IsManagementBridge) {
                    $managementBridgeEntries += $entry
                    $ignoredEntries += $entry
                }
                elseif ($classification.IsMdmEnrollment) {
                    $entries += $entry
                }
                elseif ($classification.IsIntuneCandidate) {
                    $unconfirmedIntuneEntries += $entry
                    $ignoredEntries += $entry
                }
                elseif ($classification.HasProviderId -or $classification.HasDiscoveryUrl -or (-not [string]::IsNullOrWhiteSpace($enrollmentType))) {
                    $ignoredEntries += $entry
                }
            }
            catch { }
        }

        $intuneEntries = @($entries | Where-Object { $_.IsIntune })
        $nonIntuneEntries = @($entries | Where-Object { $_.IsNonIntuneMdm })
        $providerIds = @($entries | ForEach-Object { if ([string]::IsNullOrWhiteSpace($_.ProviderID)) { "<empty>" } else { $_.ProviderID } } | Select-Object -Unique)
        $details = @($entries | ForEach-Object {
            "EnrollmentId={0},ProviderID={1},DiscoveryURL={2},EnrollmentType={3},UPN={4},Evidence={5}" -f $_.EnrollmentId,$_.ProviderID,$_.DiscoveryServiceFullURL,$_.EnrollmentType,$_.UPN,$_.Evidence
        })
        $ignoredDetails = @($ignoredEntries | ForEach-Object {
            "EnrollmentId={0},ProviderID={1},DiscoveryURL={2},EnrollmentType={3},UPN={4},Evidence={5}" -f $_.EnrollmentId,$_.ProviderID,$_.DiscoveryServiceFullURL,$_.EnrollmentType,$_.UPN,$_.Evidence
        })
        $managementBridgeProviderIds = @($managementBridgeEntries | ForEach-Object { $_.ProviderID } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)
        $managementBridgeDetails = @($managementBridgeEntries | ForEach-Object {
            "EnrollmentId={0},ProviderID={1},DiscoveryURL={2},EnrollmentType={3},Evidence={4},StrongOmaDmEvidence={5}" -f $_.EnrollmentId,$_.ProviderID,$_.DiscoveryServiceFullURL,$_.EnrollmentType,$_.Evidence,$_.HasStrongOmaDmEvidence
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
            ManagementBridgeDetected = ($managementBridgeEntries.Count -gt 0)
            ManagementBridgeEnrollmentIds = (($managementBridgeEntries | ForEach-Object { $_.EnrollmentId }) -join ";")
            ManagementBridgeProviderIds = ($managementBridgeProviderIds -join ";")
            ManagementBridgeDetails = ($managementBridgeDetails -join " | ")
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

        $backedUpPaths = New-Object System.Collections.Generic.List[string]
        $backupFailed = $false
        foreach ($regPath in $registryPaths) {
            $psPath = "Registry::$regPath"
            if (-not (Test-Path -LiteralPath $psPath)) { continue }

            $safeRegName = ($regPath -replace '[\\/:*?"<>|{} ]', '_')
            $backupFile = Join-Path $backupDir ("{0}_{1}.reg" -f $safeId,$safeRegName)
            try {
                $exportOutput = & reg.exe export $regPath $backupFile /y 2>&1
                $exportExitCode = $LASTEXITCODE
                if ($exportExitCode -ne 0) {
                    throw ("Registry export failed. ExitCode={0}; Output={1}" -f $exportExitCode, (ConvertTo-CleanNativeOutput -Output $exportOutput))
                }
                if (-not (Test-Path -LiteralPath $backupFile -PathType Leaf) -or (Get-Item -LiteralPath $backupFile -ErrorAction Stop).Length -le 0) {
                    throw 'Registry export did not produce a non-empty backup.'
                }
                $backedUpPaths.Add($regPath)
            }
            catch {
                $backupFailed = $true
                $errors.Add(("Enrollment cleanup blocked before any removal. Id={0}; BackupPath={1}; Error={2}" -f $id, $backupFile, $_.Exception.Message))
            }
        }
        if ($backupFailed) { continue }
        foreach ($regPath in $backedUpPaths) {
            try {
                Remove-Item -LiteralPath ("Registry::$regPath") -Recurse -Force -ErrorAction Stop
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

    Write-RunLog ("Running dsregcmd /status snapshot. Phase={0}" -f $PhaseLabel)
    if (-not (Test-Path -LiteralPath $OutputDirPath -PathType Container)) {
        New-Item -ItemType Directory -Path $OutputDirPath -Force -ErrorAction Stop | Out-Null
    }

    $lines = @(& $DsregcmdPath /status 2>&1 | ForEach-Object { [string]$_ })
    if (@($lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -eq 0) {
        Start-Sleep -Seconds 2
        $lines = @(& $DsregcmdPath /status 2>&1 | ForEach-Object { [string]$_ })
    }
    if (@($lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count -eq 0) {
        throw "dsregcmd /status output was empty after retry (Phase=$PhaseLabel)."
    }

    $copy = Join-Path $OutputDirPath ("{0}_dsreg_status_{1}_{2}.txt" -f $ComputerNameValue, $PhaseLabel, $RunIdValue)
    $lines | Set-Content -LiteralPath $copy -Encoding UTF8 -Force -ErrorAction Stop

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

function ConvertTo-SemtCommandLineArgument {
    <#
    Quotes one argument for a native command line (CommandLineToArgvW rules: quotes escaped as \",
    backslashes doubled before a quote and at the end). Windows PowerShell 5.1 passes an argument
    that starts and ends with a quote without quoting it, so a /TR value such as
    "C:\...\powershell.exe" -File "C:\Program Files\..." was split and schtasks.exe failed with
    "Invalid argument/option".
    #>
    param([AllowEmptyString()][string]$Value)

    if (-not [string]::IsNullOrEmpty($Value) -and $Value -notmatch '[\s"]') { return $Value }
    $escaped = [regex]::Replace([string]$Value, '(\\*)"', { param($match) ($match.Groups[1].Value * 2) + '\"' })
    $escaped = [regex]::Replace($escaped, '(\\+)$', { param($match) $match.Groups[1].Value * 2 })
    return ('"{0}"' -f $escaped)
}

function Invoke-SemtSchtasks {
    <# Runs schtasks.exe with an explicit command line (see ConvertTo-SemtCommandLineArgument). #>
    param([Parameter(Mandatory = $true)][string[]]$Arguments)

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = Join-Path $env:WINDIR 'System32\schtasks.exe'
    $startInfo.Arguments = (@($Arguments | ForEach-Object { ConvertTo-SemtCommandLineArgument -Value $_ }) -join ' ')
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.CreateNoWindow = $true
    $process = [System.Diagnostics.Process]::Start($startInfo)
    $standardError = $process.StandardError.ReadToEndAsync()
    $standardOutput = $process.StandardOutput.ReadToEnd()
    $process.WaitForExit()
    return [pscustomobject]@{ ExitCode = $process.ExitCode; Output = ('{0} {1}' -f $standardOutput,$standardError.Result).Trim() }
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

    $definition.RegistrationInfo.Author = "Smart Endpoint Migration Toolkit - Hybrid Join"
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
            Description = "Smart Endpoint Migration Toolkit - Hybrid Join - capture dsregcmd /status in the active interactive user context."
            Arguments = "/d /c %windir%\System32\dsregcmd.exe /status > %windir%\Temp\SEMT-HJ-UserDsregStatus_%USERNAME%.txt 2>&1"
        },
        [PSCustomObject]@{
            Name = $UserRefreshPrtTaskName
            Description = "Smart Endpoint Migration Toolkit - Hybrid Join - run dsregcmd /refreshprt in the active interactive user context."
            Arguments = "/d /c %windir%\System32\dsregcmd.exe /refreshprt > %windir%\Temp\SEMT-HJ-UserRefreshPrt_%USERNAME%.txt 2>&1"
        },
        [PSCustomObject]@{
            Name = $UserIntuneAutoEnrollTaskName
            Description = "Smart Endpoint Migration Toolkit - Hybrid Join - run deviceenroller /AutoEnrollMDM in the active interactive user context."
            Arguments = "/d /c %windir%\System32\deviceenroller.exe /c /AutoEnrollMDM > %windir%\Temp\SEMT-HJ-UserMdmAutoEnroll_%USERNAME%.txt 2>&1"
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
            $expectedFileName = "SEMT-HJ-UserMdmAutoEnroll_{0}.txt" -f $ExpectedUserName
        }
        $sourceFile = Get-LatestUserTaskOutputFile -Filter "SEMT-HJ-UserMdmAutoEnroll_*.txt" -Since $startedAt -ExpectedFileName $expectedFileName

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
        $expectedFileName = "SEMT-HJ-UserDsregStatus_{0}.txt" -f $ExpectedUserName
    }

    $sourceFile = Get-LatestUserTaskOutputFile -Filter "SEMT-HJ-UserDsregStatus_*.txt" -Since $startedAt -ExpectedFileName $expectedFileName
    if (-not $sourceFile) {
        $expectedDetail = if ([string]::IsNullOrWhiteSpace($expectedFileName)) { "No recent SEMT-HJ-UserDsregStatus_*.txt output file found after running user status task." } else { ("Expected user output file not found or not recent after running user status task: C:\Windows\Temp\{0}" -f $expectedFileName) }
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
        $expectedFileName = "SEMT-HJ-UserRefreshPrt_{0}.txt" -f $ExpectedUserName
    }
    $sourceFile = Get-LatestUserTaskOutputFile -Filter "SEMT-HJ-UserRefreshPrt_*.txt" -Since $startedAt -ExpectedFileName $expectedFileName

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
            "D$([char]0x00E9)co",
            "Deco",
            "Deconnecte",
            "D$([char]0x00E9)connect$([char]0x00E9)",
            "Disconn",
            "Disconnesso",
            "Desconectado",
            "Desconectada",
            "Desligado",
            "Desligada",
            "Roz$([char]0x0142)$([char]0x0105)czony",
            "Roz$([char]0x0142)$([char]0x0105)czona",
            "Roz$([char]0x0142)$([char]0x0105)czone",
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
        Sid=""
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

                $ownerSid = ""
                try {
                    $sidResult = Invoke-CimMethod -InputObject $process -MethodName GetOwnerSid -ErrorAction Stop
                    if ($sidResult.ReturnValue -eq 0) { $ownerSid = [string]$sidResult.Sid }
                } catch { }
                return [PSCustomObject]@{
                    Sid=$ownerSid
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
            Sid=""
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

    $tempStdOut = Join-Path $env:TEMP ("SEMT_HJ_{0}_stdout.tmp" -f ([guid]::NewGuid().ToString("N")))
    $tempStdErr = Join-Path $env:TEMP ("SEMT_HJ_{0}_stderr.tmp" -f ([guid]::NewGuid().ToString("N")))
    $timedOut = $false
    $exitCode = ""
    $outputText = ""
    $stopDetail = ""

    try {
        $process = Start-Process -FilePath $FilePath -ArgumentList $Arguments -PassThru -WindowStyle Hidden -RedirectStandardOutput $tempStdOut -RedirectStandardError $tempStdErr
        if (-not $process.WaitForExit($TimeoutSeconds * 1000)) {
            $timedOut = $true
            try { $process.Kill() } catch { $stopDetail = $_.Exception.Message }
            try { if (-not $process.WaitForExit(5000)) { $stopDetail = 'Process did not exit after termination was requested.' } } catch { if (-not $stopDetail) { $stopDetail = $_.Exception.Message } }
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
        StopDetail = $stopDetail
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

    try {
        $helperTaskResult = Ensure-UserContextHelperTasks
        if (-not $helperTaskResult.Success) {
            Write-RunLog ("Next-logon helper prerequisites were not fully registered. Detail={0}" -f $helperTaskResult.Detail)
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

        $definition.RegistrationInfo.Author = "Smart Endpoint Migration Toolkit - Hybrid Join"
        $definition.RegistrationInfo.Description = "Smart Endpoint Migration Toolkit - Hybrid Join - trigger user-context auto-enrollment helper tasks at next user logon."
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
        Write-RunLog ("Next-logon auto-enrollment helper task registration succeeded. TaskName={0}; Action={1}; StatusArgs={2}; AutoEnrollArgs={3}" -f $normalizedTaskName,$schtasksPath,$statusTaskArguments,$autoEnrollTaskArguments)

        return [PSCustomObject]@{
            Success = $true
            TaskName = $normalizedTaskName
            ExitCode = 0
            Detail = "Registered SYSTEM ONLOGON task with direct schtasks actions."
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
        "RETRY_AFTER_REBOOT_EXHAUSTED" { return "CHECK_REBOOT_STATE_OR_RELAUNCH_WAVE" }
        "RETRY_AFTER_REBOOT_STATE_MISSING" { return "RELAUNCH_WAVE" }
        "RETRY_AFTER_REBOOT_SCHEDULE_FAILED_POST_DSREG_LEAVE" { return "CHECK_SCHEDULED_TASK_AND_RELAUNCH" }
        "RETRY_AFTER_REBOOT_SCHEDULE_FAILED_WAITING_FOR_USER_LOGON" { return "CHECK_SCHEDULED_TASK_AND_RELAUNCH" }
        "REBOOT_SCHEDULE_FAILED_POST_DSREG_LEAVE" { return "CHECK_REBOOT_COMMAND_AND_RELAUNCH" }
        "REBOOT_SCHEDULE_FAILED_WAITING_FOR_USER_LOGON" { return "CHECK_REBOOT_COMMAND_AND_RELAUNCH" }
        "SKIPPED_VIRTUAL_MACHINE" { return "NO_ACTION_VIRTUAL_MACHINE" }
        "COMPUTER_SYSTEM_QUERY_FAILED" { return "FIX_WMI_CIM_OR_RETRY" }
        "DOMAIN_CONTROLLER_UNREACHABLE" { return "FIX_DOMAIN_CONNECTIVITY_OR_VPN" }
        "GPUPDATE_TIMEOUT" { return "INVESTIGATE_GROUP_POLICY_TIMEOUT" }
        "FINAL_GPUPDATE_TIMEOUT" { return "INVESTIGATE_GROUP_POLICY_TIMEOUT" }
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
        $events = Get-WinEvent -FilterHashtable @{ LogName=$LogName; StartTime=$Since } -MaxEvents 200 -ErrorAction Stop |
            Select-Object TimeCreated,Id,RecordId,ActivityId,LevelDisplayName,ProviderName,Message

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
        -Filter "SEMT-HJ-UserMdmAutoEnroll_*.txt" `
        -Since $Since `
        -DestinationDirectory $OutputDirPath `
        -Prefix $prefix

    if (-not [string]::IsNullOrWhiteSpace($autoEnrollCopy)) { $files += $autoEnrollCopy }

    $logs = @(
        @{ Name="Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Enrollment"; Suffix="DeviceManagement_Enrollment" },
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
        DataRoot = $DataRoot
        CorrelationId = ($CorrelationId -replace '-R\d{14}$','')
        LegacyActivityPolicy = $LegacyActivityPolicy
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
    # Trusted Program Files folder (created with a protected ACL or verified), then a verified copy
    # of this endpoint script: the task never starts a script stored under ProgramData.
    [void](Initialize-SemtTrustedProgramFolder -Path $script:RetryAfterRebootProgramDir)
    if ([System.IO.Path]::GetFullPath($script:EndpointScriptPath) -ine [System.IO.Path]::GetFullPath($script:RetryAfterRebootScriptPath)) {
        [void](Install-SemtTrustedProgramFile -Source $script:EndpointScriptPath -Destination $script:RetryAfterRebootScriptPath)
    }

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
        ScriptPath = $script:RetryAfterRebootScriptPath
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
    $programTrust = Test-SemtTrustedProgramPath -Path $script:RetryAfterRebootProgramDir -Recurse
    if (-not $programTrust.Trusted) { throw ("SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} ({1}) {2}" -f $script:RetryAfterRebootProgramDir,$programTrust.State,$programTrust.Detail) }

    $taskPowerShell = Join-Path $env:WINDIR "System32\WindowsPowerShell\v1.0\powershell.exe"
    $taskCommand = '"{0}" -NoProfile -ExecutionPolicy RemoteSigned -File "{1}"' -f $taskPowerShell,$script:RetryAfterRebootRunnerPath
    $create = Invoke-SemtSchtasks -Arguments @('/Create','/TN',$script:RetryAfterRebootTaskName,'/SC','ONSTART','/RU','SYSTEM','/RL','HIGHEST','/TR',$taskCommand,'/F')
    if ($create.ExitCode -ne 0) {
        throw ("schtasks.exe failed with exit code {0}: {1}" -f $create.ExitCode,$create.Output)
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
    $programTrust = Test-SemtTrustedProgramPath -Path $script:RetryAfterRebootProgramDir -Recurse
    if (-not $programTrust.Trusted) { throw ("SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} ({1}) {2}" -f $script:RetryAfterRebootProgramDir,$programTrust.State,$programTrust.Detail) }

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

    $paths = @($script:RetryAfterRebootRunnerPath,$script:RetryAfterRebootStatePath)
    # The endpoint copy is removed only when another script runs: a task run keeps its own file.
    if ([System.IO.Path]::GetFullPath($script:EndpointScriptPath) -ine [System.IO.Path]::GetFullPath($script:RetryAfterRebootScriptPath)) { $paths += $script:RetryAfterRebootScriptPath }
    foreach ($path in $paths) {
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
$logPath = $script:SemtCsvPath

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
$managementBridgeDetected = $false
$managementBridgeEnrollmentIds = ""
$managementBridgeProviderIds = ""
$managementBridgeDetails = ""
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
$mdmEnrollmentRecentEvidence = ""
$mdmDiagnosticStatus = "NotRequested"
$mdmDiagnosticZip = ""

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
$interactiveUserSid = ""
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
$gpUpdateTimedOut = $false
$gpUpdateStopDetail = ''
$gpUpdateExitCode = ""
$gpUpdateOutputFile = ""
$gpUpdateMdmPolicyWarning = $false
$gpUpdateMdmPolicyState = ""
$gpUpdateMdmPolicyDetail = ""
$gpResultHtmlFile = ""
$gpResultTextFile = ""
$finalGpUpdateAttempted = $false
$finalGpUpdateTimedOut = $false
$finalGpUpdateStopDetail = ''
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
                "Retry-after-reboot task started without its state file. The task was removed; relaunch the Wave."
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

    if ($CollectMdmDiagnostics) {
        # An explicit option: capture does not alter MDM enrollment. The archive is
        # sensitive; only write under SEMT's protected data root.
        $mdmDiagnosticZip = Join-Path $OutputDir ("MDMDiag_{0}_{1}.zip" -f $ComputerName,$CorrelationId)
        $mdmDiagnosticOutput = Join-Path $OutputDir ("MDMDiag_{0}_{1}.txt" -f $ComputerName,$CorrelationId)
        try {
            $tool = Join-Path $env:SystemRoot 'System32\mdmdiagnosticstool.exe'
            $capture = Invoke-ProcessWithTimeout -FilePath $tool -Arguments @('-area','DeviceEnrollment','-zip',(ConvertTo-SemtCommandLineArgument -Value $mdmDiagnosticZip)) -TimeoutSeconds 300 -OutputFile $mdmDiagnosticOutput
            if ($capture.TimedOut) { $mdmDiagnosticStatus = 'Timeout'; $mdmDiagnosticZip = '' }
            elseif ($capture.ExitCode -ne 0 -or -not (Test-Path -LiteralPath $mdmDiagnosticZip -PathType Leaf)) { $mdmDiagnosticStatus = 'Failed'; $mdmDiagnosticZip = '' }
            else {
                Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
                $archive = [IO.Compression.ZipFile]::OpenRead($mdmDiagnosticZip)
                try { if ($archive.Entries.Count -eq 0) { throw 'MDM diagnostic ZIP is empty.' } }
                finally { $archive.Dispose() }
                $mdmDiagnosticStatus = 'Collected'
            }
        } catch { $mdmDiagnosticStatus = 'Error'; $mdmDiagnosticZip = ''; Write-RunLog ("MDM diagnostics error: {0}" -f $_.Exception.Message) }
        Write-RunLog ("MDM diagnostics: Status={0}; Zip={1}; Output={2}" -f $mdmDiagnosticStatus,$mdmDiagnosticZip,$mdmDiagnosticOutput)
    }

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

    if ($AuditOnly) {
        Write-RunLog "AuditOnly: skipping gpupdate; collecting existing Entra and MDM state."
    }
    else {
        # Refresh computer policy before checking local MDM/GPO state.
        $gpUpdateAttempted = $true
        $gpUpdateOutputFile = Join-Path $OutputDir ("{0}_gpupdate_{1}.txt" -f $ComputerName, $RunId)
        Write-Host "Refreshing computer Group Policy..." -ForegroundColor Cyan
        Write-RunLog "Running gpupdate /target:computer /force."

        try {
            $gpResultProcess = Invoke-ProcessWithTimeout -FilePath 'gpupdate.exe' -Arguments @('/target:computer','/force') -TimeoutSeconds $GpUpdateTimeoutSeconds -OutputFile $gpUpdateOutputFile
            $gpUpdateExitCode = $gpResultProcess.ExitCode
            $gpupdateOutput = $gpResultProcess.Output
            $gpUpdateTimedOut = $gpResultProcess.TimedOut
            $gpUpdateStopDetail = $gpResultProcess.StopDetail
            Write-RunLog ("gpupdate finished. TimedOut={0}; ExitCode={1}; OutputFile={2}" -f $gpUpdateTimedOut,$gpUpdateExitCode,$gpUpdateOutputFile)

            $gpupdateText = (($gpupdateOutput | Out-String).Trim())
            if (-not $gpUpdateTimedOut -and $gpupdateText -match "MDM Policy") {
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

        if ($gpUpdateTimedOut) {
            $warning = "Initial gpupdate exceeded {0} seconds. Termination requested; stop detail: {1}. Existing policy will be checked; enrollment checks continue. OutputFile={2}" -f $GpUpdateTimeoutSeconds,$gpUpdateStopDetail,$gpUpdateOutputFile
            Write-Host $warning -ForegroundColor Yellow
            Write-RunLog $warning
        }

        if ($GpUpdateWaitSeconds -gt 0 -and -not $gpUpdateTimedOut) {
            Write-RunLog ("Waiting {0} seconds after gpupdate." -f $GpUpdateWaitSeconds)
            Start-Sleep -Seconds $GpUpdateWaitSeconds
        }

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
    $managementBridgeDetected = $mdmEnrollmentState.ManagementBridgeDetected
    $managementBridgeEnrollmentIds = $mdmEnrollmentState.ManagementBridgeEnrollmentIds
    $managementBridgeProviderIds = $mdmEnrollmentState.ManagementBridgeProviderIds
    $managementBridgeDetails = $mdmEnrollmentState.ManagementBridgeDetails
    $staleIntuneEnrollmentDetected = (-not $mdmEnrollmentState.IntuneEnrollmentDetected) -and (-not [string]::IsNullOrWhiteSpace($unconfirmedIntuneEnrollmentIds)) -and ($gpUpdateMdmPolicyState -eq "AlreadyEnrolled")
    if ($staleIntuneEnrollmentDetected) { $staleIntuneEnrollmentIds = $unconfirmedIntuneEnrollmentIds }
    Write-RunLog ("MDM enrollment registry state: AnyMdm={0}; Intune={1}; NonIntune={2}; Count={3}; ProviderIds={4}; IntuneIds={5}; NonIntuneIds={6}; UnconfirmedIntuneIds={7}; ManagementBridge={8}; ManagementBridgeIds={9}; Details={10}; BridgeDetails={11}; Ignored={12}" -f $anyMdmEnrollmentDetected,$mdmEnrollmentState.IntuneEnrollmentDetected,$nonIntuneMdmEnrollmentDetected,$mdmEnrollmentCount,$mdmProviderIds,$intuneEnrollmentIds,$nonIntuneEnrollmentIds,$unconfirmedIntuneEnrollmentIds,$managementBridgeDetected,$managementBridgeEnrollmentIds,$mdmEnrollmentDetails,$managementBridgeDetails,$ignoredEnrollmentDetails)
    if ($managementBridgeDetected -and -not $nonIntuneMdmEnrollmentDetected) {
        Write-RunLog ("Management bridge registration detected and protected from MDM cleanup. ProviderIds={0}; EnrollmentIds={1}. It does not block Intune auto-enrollment without independent OMA-DM evidence." -f $managementBridgeProviderIds,$managementBridgeEnrollmentIds)
    }
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
        $interactiveUserSid = [string]$interactiveIdentity.Sid
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
            # Discovery URL plus 0x8018000A does not prove a stale registration. Never
            # delete these traces based solely on this ambiguous evidence, even when the
            # historical cleanup switch is enabled.
            $status = "INTUNE_ENROLLMENT_STATE_REVIEW_REQUIRED"
            $dsregStatusErrorMessage = "Unconfirmed Intune registration and already-enrolled policy signal require manual review; no cleanup was performed."
            $ExitCode = 3
            Write-RunLog $dsregStatusErrorMessage
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
                $managementBridgeDetected = $mdmEnrollmentState.ManagementBridgeDetected
                $managementBridgeEnrollmentIds = $mdmEnrollmentState.ManagementBridgeEnrollmentIds
                $managementBridgeProviderIds = $mdmEnrollmentState.ManagementBridgeProviderIds
                $managementBridgeDetails = $mdmEnrollmentState.ManagementBridgeDetails
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

    if (-not $AuditOnly -and -not $gpUpdateTimedOut -and $finalGpUpdateReasons.Count -gt 0) {
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
                $finalGpResult = Invoke-ProcessWithTimeout -FilePath 'gpupdate.exe' -Arguments @('/target:computer','/force') -TimeoutSeconds $GpUpdateTimeoutSeconds -OutputFile $finalGpUpdateOutputFile
                $finalGpUpdateExitCode = $finalGpResult.ExitCode
                $finalGpUpdateTimedOut = $finalGpResult.TimedOut
                $finalGpUpdateStopDetail = $finalGpResult.StopDetail
                Write-RunLog ("Final gpupdate finished. TimedOut={0}; ExitCode={1}; OutputFile={2}" -f $finalGpUpdateTimedOut,$finalGpUpdateExitCode,$finalGpUpdateOutputFile)
            }
            catch {
                $finalGpUpdateExitCode = ""
                Write-RunLog ("Final gpupdate failed to start: {0}" -f $_.Exception.Message)
            }
            if ($finalGpUpdateTimedOut) {
                $warning = "Final gpupdate exceeded {0} seconds after repair; enrollment status remains {1}. Termination requested; stop detail: {2}. OutputFile={3}" -f $GpUpdateTimeoutSeconds,$status,$finalGpUpdateStopDetail,$finalGpUpdateOutputFile
                Write-Host $warning -ForegroundColor Yellow
                Write-RunLog $warning
            }
        }
    }

    if (-not $AuditOnly -and $gpUpdateTimedOut -and $finalGpUpdateReasons.Count -gt 0) {
        Write-RunLog 'Final gpupdate skipped because the initial invocation timed out; enrollment result is preserved.'
    }

    # Enrollment events describe observations in a time window, not necessarily the
    # current attempt. Preserve timestamp and record id and do not infer a causal user.
    foreach ($file in @(($intuneEnrollmentDiagnosticsInitialFiles + ';' + $intuneEnrollmentDiagnosticsFiles) -split ';' | Where-Object { $_ -like '*DeviceManagement_Enrollment.csv' })) {
        try {
            $observations = @(Import-Csv -LiteralPath $file | Where-Object { $_.Id -in @('52','76') -and $_.Message -match '(?i)UserLicense|0x80180018|0x8018000A' } | Select-Object -First 1)
            foreach ($event in $observations) {
                $code = if ($event.Message -match '(?i)UserLicense|0x80180018') { 'UserLicense/0x80180018' } else { '0x8018000A' }
                $mdmEnrollmentRecentEvidence = 'Enrollment event {0}, record {1}, time {2}: {3}; correlation with this attempt not established.' -f $event.Id,$event.RecordId,$event.TimeCreated,$code
            }
        } catch { Write-RunLog ('Unable to summarize enrollment events: {0}' -f $_.Exception.Message) }
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
        ManagementBridgeDetected = $managementBridgeDetected
        ManagementBridgeEnrollmentIds = $managementBridgeEnrollmentIds
        ManagementBridgeProviderIds = $managementBridgeProviderIds
        ManagementBridgeDetails = $managementBridgeDetails
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
        MdmEnrollmentRecentEvidence = $mdmEnrollmentRecentEvidence
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
        GpUpdateTimedOut        = $gpUpdateTimedOut
        GpUpdateStopDetail      = $gpUpdateStopDetail
        FinalGpUpdateTimedOut   = $finalGpUpdateTimedOut
        FinalGpUpdateStopDetail = $finalGpUpdateStopDetail
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
        InteractiveUserSid      = $interactiveUserSid
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
    $controlledRetryStop = $status -in @("RETRY_AFTER_REBOOT_EXHAUSTED","RETRY_AFTER_REBOOT_STATE_MISSING","GPUPDATE_TIMEOUT","FINAL_GPUPDATE_TIMEOUT")
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
            ManagementBridgeDetected = $managementBridgeDetected
            ManagementBridgeEnrollmentIds = $managementBridgeEnrollmentIds
            ManagementBridgeProviderIds = $managementBridgeProviderIds
            ManagementBridgeDetails = $managementBridgeDetails
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
            MdmEnrollmentRecentEvidence = $mdmEnrollmentRecentEvidence
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
            GpUpdateTimedOut        = $gpUpdateTimedOut
        GpUpdateStopDetail      = $gpUpdateStopDetail
        FinalGpUpdateTimedOut   = $finalGpUpdateTimedOut
        FinalGpUpdateStopDetail = $finalGpUpdateStopDetail
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
            InteractiveUserSid      = $interactiveUserSid
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
    Exit-SemtTargetOperationGuard -CampaignDataRoot $DataRoot
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCusiPoQwfBZMpY
# pw0hHoFHb1OEtzdXO5oNTkEvOR6dpqCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAJxU2rYZC3dx4cubfwyPpp
# aHL6YF3QgmiPgAyHwkejljANBgkqhkiG9w0BAQEFAASCAYAUXnP/HMGojlpI2/I4
# D0o40AkeNkCCDuKuK+U+0ywYA2+OlCFzePQEGuTlpnJL5vm2fPm4CMMr6GF+4WrT
# z6hkBiOpmTFUlwtLv9R6PPw9XUQULqU32D+b75WpHb+68n3Jrnru/aBvVMrgkuPO
# qfbYt/7gI1mKHaeeWx8MNX2FSG8sqTsIxfJX83XeMMK/fyvZGAxZ3vj9WOFEq41r
# zq7P8tMeguzcdxGxXsiCBcZclwGmQhWGPy5v0gnf+eu+8oQWmaXAnv6MjVXt8ohb
# jxTMpRNX+6mRKK19nGrPPo79O3K84wczFJ4eVxUJ9MUqaMwvzYPuGuwT83x7DVQw
# FAza23RSEXuaw0qO/pT8guWWCqRaifQ3w2PBcjTBw2/XXjcxvLAxf23uYPq4t+Ou
# x53FNTn0kX2mZoqdQqr/0qxu7a8KsbZIqBc5N0j/f3XGGe2SOOau4eQOw6Xfre2N
# Qhgf95qv6Cw9wtRYo8+rZR76Nz3jAd2E1aUOpIhGGNw2SiI=
# SIG # End signature block
