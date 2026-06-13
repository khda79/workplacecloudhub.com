# Implementation Notes

Last updated: 2026-06-13

This file summarizes the important decisions and current behavior of this repair toolkit.
It is intentionally neutralized: do not add real company names, real hostnames, real domains,
tenant IDs, user names, account names, or screenshots containing enterprise context.

## Purpose

The toolkit is used to remediate Windows devices that should be Hybrid Entra joined and enrolled
in Intune after MDM auto-enrollment policy has been applied by GPO.

The main workflow is:

1. Export a global Intune managed-device inventory to `DevicesIntune.csv`.
2. Optionally export a global Entra devices inventory to `DevicesEntra.csv`.
3. Process devices by LOT folders using PsExec.
4. Copy the latest repair script to each target device.
5. Run the repair as `SYSTEM`.
6. Collect local logs from the target device back to the admin machine.
7. Produce per-cycle CSV and HTML reports.

## Current Layout

```text
SmartM365\SmartIntuneHybridJoinToolkit\
  Export-IntuneDevicesCsv.cmd
  Export-EntraDevicesCsv.cmd
  Start-IntuneHybridJoinRepair-LotLauncher-GUI.cmd
  Update-LotCmdWrappers.cmd
  DevicesIntune.csv                 # generated, not part of the template
  DevicesEntra.csv                  # generated, not part of the template
  Scripts\
    SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1
    SmartM365-IntuneHybridJoinRepair-Export-EntraDevicesCsv.ps1
    SmartM365-IntuneHybridJoinRepair-LotLauncher-GUI.ps1
    SmartM365-Invoke-IntuneHybridJoinRepair.ps1
    SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1
    Run-IntuneHybridJoinRepairWithPsExec-Lot.cmd
    SmartM365-IntuneHybridJoinRepair-Update-LotCmdWrappers.ps1
  LOT-X\
    Computers.txt
    Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd
    Run-IntuneHybridJoinRepairWithPsExec-Once.cmd
    Run-IntuneHybridJoinRepairWithPsExec-Loop-IgnoreRunGuard.cmd
    Run-IntuneHybridJoinRepairWithPsExec-Once-IgnoreRunGuard.cmd
```

There must not be an `Export-IntuneDevicesCsv.cmd` or `Export-EntraDevicesCsv.cmd` inside LOT folders.
These CSV exports are global and must be launched from the repository root.

## Current Versions

```text
SmartM365-Invoke-IntuneHybridJoinRepair.ps1 : 2.10.30
SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1            : 2.10.41
SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1                 : 1.3.7
SmartM365-IntuneHybridJoinRepair-Export-EntraDevicesCsv.ps1                  : 1.0.1
SmartM365-IntuneHybridJoinRepair-LotLauncher-GUI.ps1                         : initial
```

## LOT Folder Principle

LOT folders should contain only:

```text
Computers.txt
Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd
Run-IntuneHybridJoinRepairWithPsExec-Once.cmd
Run-IntuneHybridJoinRepairWithPsExec-Loop-IgnoreRunGuard.cmd
Run-IntuneHybridJoinRepairWithPsExec-Once-IgnoreRunGuard.cmd
```

The CMD launchers in each LOT are intentionally tiny wrappers. They delegate to:

```text
Scripts\Run-IntuneHybridJoinRepairWithPsExec-Lot.cmd
```

This avoids updating every LOT folder when launcher logic changes.

Only the neutral `LOT-X` template is versioned in Git. Operational `LOT-*` folders contain
real computer lists and must remain local. `.gitignore` excludes all `LOT-*` folders except
the `LOT-X` template.

To refresh all LOT wrappers after creating or importing LOT folders:

```cmd
Update-LotCmdWrappers.cmd
```

Equivalent PowerShell command:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\SmartM365-IntuneHybridJoinRepair-Update-LotCmdWrappers.ps1
```

Preview only:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\SmartM365-IntuneHybridJoinRepair-Update-LotCmdWrappers.ps1 -WhatIf
```

This script also removes obsolete `Export-IntuneDevicesCsv.cmd` wrappers from LOT folders.

## LOT Launcher GUI

Use the root launcher:

```cmd
Start-IntuneHybridJoinRepair-LotLauncher-GUI.cmd
```

It calls:

```text
Scripts\SmartM365-IntuneHybridJoinRepair-LotLauncher-GUI.ps1
```

The GUI:

- asks the operator to select a computer list file (`.txt`, `.csv`, or any readable file);
- proposes a safe local `LOT-*` folder name from the selected file name;
- creates a unique `LOT-*` folder under the toolkit root;
- writes a normalized `Computers.txt` from the selected file;
- refreshes the standard LOT CMD wrappers;
- offers to launch the created LOT immediately.

For CSV files, the GUI uses `ComputerName`, `DeviceName`, `Name`, or `DisplayName` when present,
otherwise the first CSV column. Blank lines, comments starting with `#`, and duplicate names are
removed from the generated `Computers.txt`.

If the requested LOT folder already exists, the GUI creates a timestamped variant instead of
overwriting an existing operational folder.

Launch modes map to the standard wrappers:

```text
Loop                -> Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd
Once                -> Run-IntuneHybridJoinRepairWithPsExec-Once.cmd
LoopIgnoreRunGuard  -> Run-IntuneHybridJoinRepairWithPsExec-Loop-IgnoreRunGuard.cmd
OnceIgnoreRunGuard  -> Run-IntuneHybridJoinRepairWithPsExec-Once-IgnoreRunGuard.cmd
```

## Global Intune Inventory

Use only the root launcher:

```cmd
Export-IntuneDevicesCsv.cmd
```

It calls:

```text
Scripts\SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1
```

The full Intune inventory must be written to:

```text
DevicesIntune.csv
```

When `SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1` is launched directly from `Scripts`, its default output is now the parent folder, not `Scripts\DevicesIntune.csv`.

The export script:

- Uses `Microsoft.Graph.Authentication`, not the full `Microsoft.Graph` meta-module.
- Auto-installs `Microsoft.Graph.Authentication` for `CurrentUser` if missing.
- Uses Graph `/deviceManagement/managedDevices`.
- Defaults to `-PageSize 999`.
- Writes to a temporary CSV first, then replaces the final CSV only after the temporary export succeeds.
- Shows a clean error if the output path is not writable or the CSV is open in Excel.

## Intune CSV Selection By LOT Launcher

LOT repair launchers look for the inventory in this order:

1. Parent/root `DevicesIntune.csv`
2. LOT-local `DevicesIntune.csv` only as fallback

If the selected CSV is missing or older than 15 minutes, the repair launcher automatically
runs a full Graph inventory export before starting the lot. This uses
`Scripts\SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1` directly with `-ForceRefresh`, not the root CMD wrapper,
so looped LOT runs do not block on a CMD pause.

## PsExec Repair Launcher

Main launcher:

```text
Scripts\SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1
```

PsExec availability is checked before any non-dry-run cycle starts:

- direct PowerShell runs resolve `-PsExecPath`, `Scripts\PsExec.exe`, then `PATH`;
- the shared LOT CMD launcher checks `Scripts\PsExec.exe` or `PATH` before calling PowerShell;
- the LOT GUI checks the same locations before opening the elevated LOT console.

If PsExec is missing, launchers must stop with a clear remediation message instead of reaching
`Start-Process` and producing a late or ambiguous failure.

`-DryRun` must remain usable without PsExec and without automatic Graph inventory refreshes.
If inventory CSV files are missing or older than 15 minutes during dry-run, the launcher logs
the condition and skips the refresh instead of connecting to Graph.

Default LOT behavior:

- `ThrottleLimit = 10`
- Delay between cycles = 5 minutes
- Loop mode enabled by default
- `Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd` keeps loop mode enabled
- `Run-IntuneHybridJoinRepairWithPsExec-Loop-IgnoreRunGuard.cmd` keeps loop mode enabled and sets `-IgnoreRunGuard`
- `Run-IntuneHybridJoinRepairWithPsExec-Once.cmd` sets `-RunOnce` and stops after one cycle
- `Run-IntuneHybridJoinRepairWithPsExec-Once-IgnoreRunGuard.cmd` sets both `-RunOnce` and `-IgnoreRunGuard`
- `IgnoreRunGuard` is only sent on the first cycle unless explicitly forced every cycle
- Per-computer PsExec logs go to `LOT-*\PsExecLogs`
- Cycle CSV/HTML reports go to `LOT-*\Reports`
- A live CSV and live HTML report are created in `LOT-*\Reports` at the beginning of each
  cycle. The CSV is appended as each computer completes, and the live HTML is refreshed every
  10 completed computers plus at cycle end. This prevents large lots from having an empty
  `Reports` folder while a long cycle is still running.
- HTML cycle reports include `Timestamp` in the per-computer details table.
- Remote evidence collected from target devices goes to `LOT-*\CentralLogs`
- PsExec has a configurable per-computer timeout. Default LOT wrapper value is 120 minutes via
  `EHJIR_PSEXEC_TIMEOUT_MINUTES`. Status `PSEXEC_TIMEOUT` maps to `CHECK_REMOTE_LOG_OR_RETRY`.
- If PsExec starts the remote PowerShell process but loses the service channel
  (`Error communicating with PsExec service`, `Descripteur non valide`, or equivalent invalid
  handle text), the launcher uses `PSEXEC_COMMUNICATION_LOST` /
  `RETRY_PSEXEC_OR_CHECK_REMOTE_SERVICE`.
- For `PSEXEC_COMMUNICATION_LOST`, the launcher polls the remote evidence folder before central
  collection. Defaults: `CommunicationLostEvidencePollMinutes = 10`,
  `CommunicationLostEvidenceWaitMinutes = 65`. This lets remote runs that continue after PsExec
  loses the channel be collected as soon as completed current-run evidence appears, without
  waiting the full maximum window. `LastRun.json` is authoritative when it has a current `RunId`,
  non-empty `EndTime`, `Status`, and `ExitCode`. A matching current-run CSV enriches detail,
  user/session/PRT fields, even if the CSV `ExitCode` field is blank. `PSEXEC_EXIT_UNKNOWN`
  still uses the ordinary 30-second flush wait before the same completed-run evidence check.
- If PsExec returns no readable exit code, classify as `PSEXEC_EXIT_UNKNOWN` instead of
  producing a broken `PSEXEC_EXIT_` status.
- Remote log collection keeps the full latest snapshot in `CentralLogs\<Computer>\Latest`.
  When current-run files can be identified by timestamp, the launcher also creates
  `CentralLogs\<Computer>\LatestCurrentRun` and reports it in `RemoteCurrentRunLogsPath`.
- `ADMIN_SHARE_UNREACHABLE` includes `AdminShareFailureType`: `DNS_FAILED`,
  `PING_FAILED_ADMIN_SHARE_FAILED`, or `PING_OK_ADMIN_SHARE_FAILED`.
- Final reports include `EffectiveStatus` and `EffectiveNextAction`. If the post-cycle full
  Intune inventory sees the device, `EffectiveStatus` becomes `ENROLLED_DETECTED_POST_CYCLE`
  and `EffectiveNextAction` becomes `NO_ACTION_ALREADY_INTUNE`.
- If `DevicesEntra.csv` exists, final reports include Entra device inventory columns:
  `EntraInventoryPresent`, `EntraRegisteredState`, `EntraAlternativeSecurityIdCount`,
  `EntraPendingReason`, `EntraRegistrationDateTime`, `EntraTrustType`, `EntraDeviceId`,
  and `EntraObjectId`. `SmartM365-IntuneHybridJoinRepair-Export-EntraDevicesCsv.ps1` marks `EntraRegisteredState=Pending`
  only for Microsoft Entra hybrid objects where `TrustType=ServerAd` and
  `AlternativeSecurityIds` is empty, matching the Microsoft pending-device criterion.
- When a target computer is `EntraRegisteredState=Pending`, the PsExec launcher passes
  `-EntraHybridPending` to the remote script. If the local device join is healthy, the remote
  script triggers `Automatic-Device-Join` and retries without running `dsregcmd /leave`.
  If the local device join is unhealthy, the existing guarded `DeviceAuthStatus=FAILED` /
  `KeySignTest=FAILED` leave logic applies.
- The PsExec launcher refreshes `DevicesEntra.csv` automatically like `DevicesIntune.csv`:
  full Graph export before the lot when the CSV is missing or older than 15 minutes, and another
  full Graph export at the end of each cycle. Reports include post-cycle Entra columns such as
  `PostCycleEntraRegisteredState`, `PostCycleEntraAlternativeSecurityIdCount`, and
  `PostCycleEntraPendingResolved`.
- For `PSEXEC_COMMUNICATION_LOST` / no-final-status cases, completed `LastRun.json` is the
  authoritative final status when `RunId` matches the current-run CSV. The CSV is still used to
  enrich user/session/PRT columns, even when its `ExitCode` field is blank.
- When completed current-run evidence is collected from the remote CSV, launcher reports include
  user context columns: `InteractiveUserAccountName`, `InteractiveUserAccountType`,
  `InteractiveSessionName`, `InteractiveSessionState`, `UserIsUserAzureAD`, `UserAzureAdPrt`,
  and `UserSessionIsNotRemote`. This is especially useful for filtering `USER_NOT_AZUREAD` and
  `LOGON_WITH_DOMAIN_OR_AAD_USER` without opening per-device logs.
- Stale/non-Intune MDM cleanup treats registry export failures as backup warnings only. The
  cleanup is considered failed only when actual registry or scheduled-task removal fails.
- At run start, the remote script deletes evidence files older than 7 days from
  `Logs`, `Output`, and `Transcripts`. It keeps the current run log, `LastRun.json`, and the
  main per-computer CSV.
- At run start, the remote script logs Windows OS caption, version, build number, architecture,
  product type, last boot time, and uptime. These OS fields are also written to the local CSV.
- User-context PRT refresh is no longer based only on `AzureAdPrt=YES/NO`. The remote script
  parses `AzureAdPrtExpiryTime` plus refresh diagnostics (`Attempt Status`, `HTTP status`,
  `HTTP Error`) and runs `DSREGCMD-RefreshPRT-USER` when the PRT is absent, expired/near expiry,
  or the latest refresh diagnostic is failed.
- If user-context PRT is still expired or failed after `DSREGCMD-RefreshPRT-USER`, the remote
  script stops with `USER_PRT_REFRESH_FAILED` / `FIX_USER_PRT_OR_RELOGIN` instead of launching
  User Credential MDM auto-enrollment with an unusable PRT.
- The remote script records both terminal session detection and user dsreg remote-session state:
  `InteractiveSessionIsRemote` and `User_SessionIsNotRemote`. If User Credential auto-enrollment
  would run from an RDP/remote user session, it stops with `USER_SESSION_REMOTE` /
  `LOGON_ON_CONSOLE`.
- The PsExec launcher must copy the local repair script to
  `C:\ProgramData\SmartM365\IntuneHybridJoinToolkit\SmartM365-Invoke-IntuneHybridJoinRepair.ps1`
  before every remote launch and verify remote version/hash against the local script. If the
  remote copy does not match, it must stop with `REMOTE_SCRIPT_COPY_FAILED` before PsExec.
- When PsExec returns no structured final status, the launcher must not reclassify from old
  collected CSV evidence. It only trusts evidence files newer than the current job start, and it
  recovers `PSEXEC_EXIT_<code>` from native PsExec STDERR lines when the process object does not
  expose an exit code.
- MDM Policy `AlreadyEnrolled` classification handles localized gpresult/gpupdate text in
  French, English, Portuguese, Spanish, Italian, German, Dutch, and Polish.
- After each cycle, computers detected as already Intune-enrolled are removed from that LOT's
  `Computers.txt` and appended to `ComputersAlreadyEnrolled.txt` in the same folder.
  Detection uses either a positive Intune inventory match or a remote status/next-action that
  indicates the device is already enrolled. `ComputersAlreadyEnrolled.txt` is ignored by Git.
- At the beginning of each cycle, before any PsExec job is queued, the launcher also pre-filters
  computers already present in `DevicesIntune.csv`: it moves them from the LOT `Computers.txt`
  to `ComputersAlreadyEnrolled.txt` and skips PsExec for those machines.
- By default, the launcher runs a full post-cycle Intune inventory export after each cycle and
  updates `DevicesIntune.csv`. It then compares only the computers from the completed cycle
  against that full inventory and adds post-cycle Intune columns to the CSV/HTML reports.
  Use `-SkipPostCycleIntuneInventory` only when the post-cycle Graph refresh must be disabled.
- If PsExec returns an unknown/invalid exit code but remote evidence was collected, the launcher
  reads the latest collected local repair CSV and reclassifies the launcher status from the
  remote final status where possible.
- If PsExec returns no final status, the launcher waits briefly before collecting evidence so the
  remote script has time to flush CSV, logs, and gpresult files.
- Remote evidence collection copies files one by one and ignores files that disappear between
  enumeration and copy, especially stale backup `.reg` files. One missing file must not fail the
  entire CentralLogs collection.
- Native command warnings written to logs should be cleaned so PowerShell `NativeCommandError`
  stack details such as `At ... char`, `CategoryInfo`, and `FullyQualifiedErrorId` are not copied
  into operational detail fields.

Default repair flags sent by the shared LOT launcher:

```text
-AllowDsregLeave
-AllowRemoveStaleIntuneEnrollment
-AllowRebootWhenNoInteractiveUser
-AllowRebootAfterDsregLeave
-StaleCleanupDelaySeconds 60
-RebootDelaySeconds 180
-IntuneRetrySleepMinutes 5
-IntuneRetryMaxRetries 5
```

Do not add `-AllowRemoveNonIntuneMdmEnrollment` to the shared LOT launcher by default.
Non-Intune MDM cleanup remains explicit opt-in only.

Environment variables can override defaults before launching:

```cmd
set EHJIR_THROTTLE=10
set EHJIR_DELAY_BETWEEN_CYCLES_MINUTES=5
set EHJIR_INTUNE_RETRY_SLEEP_MINUTES=5
set EHJIR_INTUNE_RETRY_MAX_RETRIES=5
set EHJIR_STALE_CLEANUP_DELAY_SECONDS=60
set EHJIR_REBOOT_DELAY_SECONDS=180
set EHJIR_PSEXEC_TIMEOUT_MINUTES=120
```

## Remote Script Copy Behavior

The launcher copies the latest local repair script to each target at every run.
Use `C$` for `C:\ProgramData`; do not use `ADMIN$` for that path because `ADMIN$`
maps to `C:\Windows`, not to `C:\`.

Source:

```text
Scripts\SmartM365-Invoke-IntuneHybridJoinRepair.ps1
```

Target:

```text
C:\ProgramData\SmartM365\IntuneHybridJoinToolkit\SmartM365-Invoke-IntuneHybridJoinRepair.ps1
```

The copy uses `-Force`, then verifies:

- target folder exists
- script exists after copy
- local and remote file sizes match

If the remote PowerShell process reports that `-File` does not exist, the launcher should classify it as:

```text
REMOTE_SCRIPT_MISSING
NextAction=FIX_SCRIPT_COPY_OR_SECURITY
```

instead of a vague `PSEXEC_EXIT_-196608`.

## Remote Log Collection

Remote evidence is collected from:

```text
C:\ProgramData\SmartM365\IntuneHybridJoinToolkit
```

Central collection folders:

```text
LOT-*\CentralLogs\<Computer>\Latest
```

By default, `Latest` is overwritten each run. `-KeepCentralLogHistory` keeps per-cycle folders.

The launcher must not copy the remote repair `.ps1` into central logs. It collects evidence such as:

- `Logs`
- `Output`
- `Transcripts`
- root evidence files with extensions like `.csv`, `.log`, `.txt`, `.html`, `.json`, `.xml`, `.evtx`

Empty `Latest` folders should not remain. If no remote evidence files are found, the folder is removed and a collection warning is logged.

## Main Repair Logic

Main remote script:

```text
Scripts\SmartM365-Invoke-IntuneHybridJoinRepair.ps1
```

Important behavior:

- Compatible with Windows PowerShell 5.1.
- Uses a 12-hour run guard by default.
- `-IgnoreRunGuard` bypasses the guard.
- Checks domain join and domain controller reachability before `gpupdate`.
- Runs `gpupdate /target:computer /force` only after the domain/DC preflight succeeds.
- If the device is domain-joined but no domain controller is reachable, exits diagnostic-only
  with `DOMAIN_CONTROLLER_UNREACHABLE` and `FIX_DOMAIN_CONNECTIVITY_OR_VPN`.
- When the DC is unreachable, the script must not run `gpupdate`, `dsregcmd /leave`,
  user `refreshprt`, forced auto-enrollment, stale/non-Intune MDM cleanup, or corrective reboot.
- The script detects when the AD computer object is in the default `CN=Computers` container by
  reading the domain naming context dynamically. Do not hardcode a domain DN in this check.
  If auto-enrollment policy is missing and the object is in the default container, the status
  detail should mention that OU-linked auto-enrollment GPOs may not apply.
- SYSTEM `dsregcmd /status` is used only for device/join fields.
- `dsregcmd` values are normalized so outputs like `KeySignTest : : PASSED` are treated as `PASSED`.
- User PRT/WAM fields from SYSTEM context are ignored.
- User-context checks depend on scheduled tasks deployed by GPO.
- When an interactive user is detected, user-context output lookup targets the expected file
  name in `C:\Windows\Temp` instead of scanning all matching wildcard files. This avoids long
  delays on endpoints with overloaded Temp folders.
- No active interactive user means user-context PRT/status tasks are skipped.
- If auto-enrollment uses User Credential and there is no active user, repair may require waiting for a valid user logon.
- Controlled reboot is allowed only when the corresponding switches are passed.
- `dsregcmd /leave` is intentionally aggressive for broken Hybrid Join: when `-AllowDsregLeave`
  is present and `DeviceAuthStatus` starts with `FAILED`, the script can run `/leave` without
  requiring `AzureAdJoined`, `DeviceId`, `TenantId`, or Intune-enrollment state to be populated.
- `dsregcmd /leave` is also allowed when `-AllowDsregLeave` is present and
  `AzureAdJoined=YES` with `KeySignTest=FAILED`, even when `DeviceAuthStatus` is empty.
  Without leave permission, this condition should return `KEY_SIGN_TEST_FAILED`.
- A computer gpresult HTML/text export is generated when gpupdate reports an MDM Policy warning
  and also when MDM auto-enrollment policy is not configured, so the report folder has local GPO
  evidence even when repair cannot proceed.
- Before long Intune enrollment confirmation waits, the remote script writes a provisional CSV
  state. If PsExec loses communication during the wait, the launcher can still recover a meaningful
  `INTUNE_ENROLLMENT_PENDING_CONFIRMATION` status from CentralLogs.

User-context scheduled task names expected on targets:

```text
DSREGCMD-Status-USER
DSREGCMD-RefreshPRT-USER
DEVICE-ENROLLER-AutoEnrollMDM-USER
```

## Optional GPO Logon Repair Task

An optional scheduled task can be deployed by GPO to run the main repair at user logon.
Recommended neutral configuration:

- Trigger: at log on of any user, delayed by about 1 minute.
- Run as `NT AUTHORITY\SYSTEM`.
- Run whether the user is logged on or not.
- Run with highest privileges.
- Do not pass `-IgnoreRunGuard`; keep the 12-hour local guard.
- Do not pass `-AllowRebootWhenNoInteractiveUser`; the task is meant to run when a user logs on.
- Prefer disabling Task Scheduler's own "restart every 5 minutes" retry because the script uses
  non-zero exit code 3 for expected "action needed / wait / retry later" states, not only crashes.

Recommended action:

```text
Program/script:
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe

Arguments:
-NoProfile -ExecutionPolicy Bypass -File "C:\ProgramData\SmartM365\IntuneHybridJoinToolkit\SmartM365-Invoke-IntuneHybridJoinRepair.ps1" -AllowDsregLeave -AllowRemoveStaleIntuneEnrollment -AllowRebootAfterDsregLeave

Start in:
C:\ProgramData\SmartM365\IntuneHybridJoinToolkit
```

The script supports localized active-session detection for multiple Windows languages.

If no interactive user is present and User Credential MDM auto-enrollment is required,
the script attempts to register a next-logon helper task under:

```text
\IntuneHybridJoinToolkit\RunUserAutoEnrollAtLogon
```

The script must create the scheduled-task folder before creating this task and must capture
`schtasks.exe` output without leaking PowerShell `NativeCommandError` stack traces into logs.
Folder creation must be idempotent: if the Task Scheduler COM API says the folder already
exists, continue and retrieve it instead of failing the helper registration.
If COM cannot confirm the folder after an `already exists` response, continue to the
`schtasks /Create` registration attempt instead of skipping the next-logon helper.
The next-logon task must call a helper CMD file under the local data root instead of embedding
multiple `schtasks /Run` commands directly in `/TR`; otherwise `schtasks /Create` can parse
the nested `/Run` switches as duplicate create options.

## Intune Enrollment Detection

Do not treat `MdmUrl` alone as proof of Intune enrollment.

Confirmed local Intune enrollment requires stronger local evidence, especially:

```text
ProviderID = MS DM Server
```

Internal Windows authorities must not be treated as third-party MDM.

Stale local Intune enrollment cleanup is opt-in:

```text
-AllowRemoveStaleIntuneEnrollment
```

Non-Intune MDM cleanup is also opt-in and should remain explicit:

```text
-AllowRemoveNonIntuneMdmEnrollment
```

Stale local Intune cleanup uses the same low-level registry/task removal helper, but its log
and backup folder label must stay explicit as `Stale local Intune enrollment trace`, not
`Non-Intune MDM`, because it is an old incomplete Intune trace blocking a fresh enrollment
attempt.

## Important Statuses

Useful launcher/report statuses:

```text
SUCCESS
ADMIN_SHARE_UNREACHABLE
PSEXEC_COMMUNICATION_LOST
RUN_GUARD_ACTIVE
REMOTE_DIRECTORY_CREATE_FAILED
REMOTE_SCRIPT_COPY_FAILED
REMOTE_SCRIPT_MISSING
DOMAIN_CONTROLLER_UNREACHABLE
INTUNE_ENROLLMENT_PENDING_CONFIRMATION
INTUNE_USER_AUTOENROLL_TASK_NOT_FOUND
USER_NOT_AZUREAD
USER_PRT_NOT_AVAILABLE
INTUNE_AUTOENROLL_POLICY_NOT_CONFIGURED
INTUNE_ENROLLMENT_CONNECTIVITY_FAILED
KEY_SIGN_TEST_FAILED
```

`CHECK_CURRENT_RUN_REMOTE_LOG` means the launcher could not infer a clean final status from
PsExec output and no current-run final CSV was found. Files can still exist in `Latest`; do not
trust an old CSV there unless its timestamp matches the run. Check:

```text
LOT-*\PsExecLogs\<Computer>_cycle*.log
LOT-*\CentralLogs\<Computer>\Latest
LOT-*\CentralLogs\<Computer>\LatestCurrentRun
```

## Ctrl+C Behavior

Stopping the CMD/PowerShell launcher with `Ctrl+C` interrupts the local orchestrator.

Important effects:

- The current cycle may not write its final CSV/HTML report if interrupted before summary generation.
- PsExec processes already started may continue briefly.
- The remote script may continue on a target if it already started as `SYSTEM`.
- No rollback is performed.
- Already collected logs remain.

## Reports And Logs

Current intended separation:

```text
LOT-*\PsExecLogs     # one log per computer/cycle
LOT-*\Reports        # PsExec_IntuneHybridJoinRepair_Summary_cycle*.csv/html
LOT-*\CentralLogs    # collected remote evidence per computer
```

Do not put summary reports in `PsExecLogs`.

## Privacy / Sanitization Rule

Do not commit or add examples containing real:

- enterprise names
- domains
- hostnames
- tenant IDs
- account names
- user names
- device names
- screenshots showing enterprise context

Template examples should use neutral values only:

```text
PC-001
PC-002.example.local
contoso.onmicrosoft.com
```

## Validation Commands

PowerShell syntax check:

```powershell
$files = @(
  '.\Scripts\SmartM365-Invoke-IntuneHybridJoinRepair.ps1',
  '.\Scripts\SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1',
  '.\Scripts\SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1',
  '.\Scripts\SmartM365-IntuneHybridJoinRepair-Update-LotCmdWrappers.ps1',
  '.\Scripts\SmartM365-IntuneHybridJoinRepair-LotLauncher-GUI.ps1'
)
foreach ($file in $files) {
  $errors = $null
  [System.Management.Automation.PSParser]::Tokenize((Get-Content -LiteralPath $file -Raw), [ref]$errors) > $null
  if ($errors) { $errors | Format-List *; throw "PSParser failed: $file" }
}
```

Dry-run example:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1 `
  -DryRun `
  -RunOnce `
  -ComputerListPath .\LOT-X\Computers.txt `
  -LogRoot .\LOT-X\PsExecLogs `
  -ReportRoot .\LOT-X\Reports `
  -CentralLogRoot .\LOT-X\CentralLogs
```
