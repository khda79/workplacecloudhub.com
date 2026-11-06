<#
.SYNOPSIS
    Offline tests of the Windows 11 upgrade decision for computers not enrolled in Intune.
.DESCRIPTION
    Extracts the real decision block of the Windows11Upgrade endpoint script (from the Intune
    gate comment to the end of the upgrade decision chain) and runs it with stubbed actions
    (disk cleanup, policy repair, setup, Windows Update). Checks that a computer not enrolled in
    Intune is upgraded by setup through the usual gates, and that NOT_INTUNE_ENROLLED is reported
    only when setup upgrade is disabled.
    Prints one PASS/FAIL line per assertion and "Summary: Passed=N; Failed=M". Exits 1 on failure.
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

$ErrorActionPreference = 'Stop'
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$w11 = Get-Content -LiteralPath (Join-Path $root 'Campaigns\Windows11Upgrade\Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1') -Raw
$passed = 0
$failed = 0
function Assert-Test([string]$Name, [bool]$Condition) {
    if ($Condition) { $script:passed++; Write-Host "PASS: $Name" }
    else { $script:failed++; Write-Host "FAIL: $Name" }
}

$start = $w11.IndexOf('    # Intune enrollment does not gate the setup upgrade')
$end = $w11.IndexOf("`r`n}`r`ncatch [System.OperationCanceledException]", $start)
if ($start -lt 0 -or $end -lt 0) { throw 'Windows 11 upgrade decision block not found.' }
$decision = [scriptblock]::Create($w11.Substring($start, $end - $start))

function Write-SmartLog { param($Message,$Level = 'INFO') [void]$script:logs.Add(('{0}|{1}' -f $Level,$Message)) }
function Invoke-SafeDiskCleanup { $script:calls['DiskCleanup']++; $script:DiskCleanupAction = 'SafeCleanup'; return $script:freeAfterCleanup }
function Invoke-OldUpgradeFolderCleanup { $script:calls['AdvancedCleanup']++; return $script:freeAfterCleanup }
function Invoke-DismComponentCleanup { $script:calls['DismCleanup']++; return $script:freeAfterCleanup }
function Repair-WindowsUpdatePolicyBlockers { $script:calls['PolicyRepair']++; [pscustomobject]@{ TimedOut = $false; StopDetail = ''; OutputFile = 'gpupdate.txt' } }
function Get-WindowsUpdatePolicySummary { [pscustomobject]@{ HasLegacyBlocker = $false; Issues = '' } }
function Resolve-PendingRebootOutcome { param($PendingRebootInfo,$RebootReason) [pscustomobject]@{ Status = 'PENDING_REBOOT_USER_CONNECTED'; NextAction = 'WAIT'; Detail = ''; ExitCode = 3; ActionResult = '' } }
function Resolve-SetupUpgradeExecutable { $script:calls['ResolveSetup']++; return 'C:\SetupMedia\setup.exe' }
function Invoke-SetupUpgrade { param($SetupExePath) $script:calls['Setup']++; return 0 }
function Get-SetupExitCodeInfo { param($ExitCode) [pscustomobject]@{ ExitCode = $ExitCode } }
function Format-SetupExitCodeInfo { param($Info) return ('ExitCode={0}' -f $Info.ExitCode) }
function Resolve-SetupUpgradeSuccessOutcome { param($Prefix,$SetupExitCode) [pscustomobject]@{ Status = "${Prefix}_STARTED"; NextAction = 'MONITOR_SETUP_AND_REBOOT'; ExitCode = 0; ActionResultSuffix = '' } }
function Save-SetupReturnCheckpoint { param($Status,$NextAction,$Detail,$ActionResult,$ExitCode,$SetupExePath) }
function Resolve-SetupFailureOutcome { param($SetupExitInfo) [pscustomobject]@{ Handled = $false } }
function Get-SetupFailureNextAction { param($Info) 'CHECK_SETUP_LOGS' }
function Reset-WindowsUpdateComponents { $script:calls['WUReset']++ }
function Invoke-AssignedUpdateInstall { $script:calls['ForceUpgrade']++; return 'Triggered' }

function Invoke-Decision {
    param(
        [bool]$Enrolled,
        [bool]$SetupUpgrade = $true,
        [bool]$Audit = $false,
        [int]$FreeGb = 50,
        [int]$FreeAfterCleanup = 50,
        [bool]$PolicyBlocker = $false
    )
    $script:logs = New-Object System.Collections.ArrayList
    $script:calls = @{ DiskCleanup = 0; AdvancedCleanup = 0; DismCleanup = 0; PolicyRepair = 0; ResolveSetup = 0; Setup = 0; WUReset = 0; ForceUpgrade = 0 }
    $script:freeAfterCleanup = $FreeAfterCleanup
    $script:DiskCleanupAction = ''; $script:AdvancedDiskCleanupAction = ''; $script:DismCleanupAction = ''
    $script:HardwareReadiness = [pscustomobject]@{ Result = 'CAPABLE'; Reason = ''; Logging = '' }
    $os = [pscustomobject]@{ MajorFamily = 'Windows10' }
    $hardwareReadinessBlocksUpgrade = $false
    $intune = [pscustomobject]@{ IsIntuneEnrolled = $Enrolled; StrongEvidence = ''; WeakEvidence = 'none' }
    $indicators = [pscustomobject]@{ ActionableBlocking = $false; BlockingReasons = '' }
    $policy = [pscustomobject]@{ HasLegacyBlocker = $PolicyBlocker; Issues = $(if ($PolicyBlocker) { 'TargetReleaseVersion' } else { '' }) }
    $freeGb = $FreeGb
    $MinimumFreeDiskGB = 32
    $pendingReboot = $false
    $pendingRebootInfo = [pscustomobject]@{ IsPending = $false; Source = '' }
    $AllowSetupUpgrade = $SetupUpgrade
    $DirectSetupUpgrade = $false
    $AuditOnly = $Audit
    $AllowDiskCleanup = $true
    $AllowAdvancedDiskCleanup = $false
    $AllowDismComponentCleanup = $false
    $AllowPolicyRepair = $true
    $AllowReboot = $true
    $ScheduleRetryAfterReboot = $true
    $AllowWUReset = $true
    $AllowForceUpgrade = $true
    $policyGpUpdateTimedOut = $false
    $policyRepairCompletedNormally = $false
    $status = 'UNKNOWN'; $nextAction = ''; $detail = ''; $actionResult = ''; $exitCode = 1
    . $decision
    return [pscustomobject]@{ Status = $status; NextAction = $nextAction; ExitCode = $exitCode; ActionResult = $actionResult; Calls = $script:calls.Clone(); Logs = @($script:logs) }
}

$warnPattern = '^WARN\|Device is not enrolled in Intune'

$r = Invoke-Decision -Enrolled $false
Assert-Test 'not enrolled + setup upgrade: setup.exe runs' ($r.Calls['Setup'] -eq 1 -and $r.Status -eq 'SETUP_UPGRADE_STARTED' -and $r.ExitCode -eq 0)
Assert-Test 'not enrolled + setup upgrade: no NOT_INTUNE_ENROLLED' ($r.Status -ne 'NOT_INTUNE_ENROLLED')
Assert-Test 'not enrolled + setup upgrade: warning logged once' (@($r.Logs | Where-Object { $_ -match $warnPattern }).Count -eq 1 -and @($r.Logs | Where-Object { $_ -match $warnPattern -and $_ -match 'continues through the usual gates' }).Count -eq 1)
Assert-Test 'not enrolled + setup upgrade: no Windows Update reset or forced update' ($r.Calls['WUReset'] -eq 0 -and $r.Calls['ForceUpgrade'] -eq 0)

$r = Invoke-Decision -Enrolled $true
Assert-Test 'enrolled + setup upgrade: setup.exe runs as before' ($r.Calls['Setup'] -eq 1 -and $r.Status -eq 'SETUP_UPGRADE_STARTED')
Assert-Test 'enrolled: no Intune warning' (@($r.Logs | Where-Object { $_ -match $warnPattern }).Count -eq 0)

$r = Invoke-Decision -Enrolled $false -SetupUpgrade $false
Assert-Test 'not enrolled + setup disabled: NOT_INTUNE_ENROLLED' ($r.Status -eq 'NOT_INTUNE_ENROLLED' -and $r.NextAction -eq 'FIX_INTUNE_ENROLLMENT_FIRST' -and $r.ExitCode -eq 3)
Assert-Test 'not enrolled + setup disabled: nothing runs' ($r.Calls['Setup'] -eq 0 -and $r.Calls['WUReset'] -eq 0 -and $r.Calls['ForceUpgrade'] -eq 0 -and $r.Calls['ResolveSetup'] -eq 0)
Assert-Test 'not enrolled + setup disabled: warning explains why' (@($r.Logs | Where-Object { $_ -match $warnPattern -and $_ -match 'setup upgrade is disabled' }).Count -eq 1)

$r = Invoke-Decision -Enrolled $false -SetupUpgrade $false -FreeGb 10 -FreeAfterCleanup 40
Assert-Test 'not enrolled + setup disabled: no disk cleanup (unchanged)' ($r.Calls['DiskCleanup'] -eq 0 -and $r.Status -eq 'NOT_INTUNE_ENROLLED')

$r = Invoke-Decision -Enrolled $false -FreeGb 10 -FreeAfterCleanup 40
Assert-Test 'not enrolled + low disk: safe disk cleanup runs, then setup' ($r.Calls['DiskCleanup'] -eq 1 -and $r.Calls['Setup'] -eq 1 -and $r.Status -eq 'SETUP_UPGRADE_STARTED')

$r = Invoke-Decision -Enrolled $false -FreeGb 10 -FreeAfterCleanup 20
Assert-Test 'not enrolled + disk still low after cleanup: INSUFFICIENT_DISK_AFTER_CLEANUP' ($r.Calls['DiskCleanup'] -eq 1 -and $r.Calls['Setup'] -eq 0 -and $r.Status -eq 'INSUFFICIENT_DISK_AFTER_CLEANUP')

$r = Invoke-Decision -Enrolled $false -PolicyBlocker $true
Assert-Test 'not enrolled + policy blocker: policy repair runs, setup waits for the next run' ($r.Calls['PolicyRepair'] -eq 1 -and $r.Calls['Setup'] -eq 0 -and $r.Status -eq 'WU_POLICY_BLOCKER' -and $r.ActionResult -eq 'PolicyRepairCompleted')

$r = Invoke-Decision -Enrolled $false -Audit $true
Assert-Test 'not enrolled + audit: SETUP_UPGRADE_READY without setup' ($r.Status -eq 'SETUP_UPGRADE_READY' -and $r.Calls['Setup'] -eq 0 -and $r.ActionResult -eq 'SetupValidatedAuditOnly')

$r = Invoke-Decision -Enrolled $false -Audit $true -FreeGb 10 -FreeAfterCleanup 40
Assert-Test 'not enrolled + audit + low disk: no cleanup, INSUFFICIENT_DISK' ($r.Calls['DiskCleanup'] -eq 0 -and $r.Status -eq 'INSUFFICIENT_DISK')

Write-Host "Summary: Passed=$passed; Failed=$failed"
if ($failed) { exit 1 }
