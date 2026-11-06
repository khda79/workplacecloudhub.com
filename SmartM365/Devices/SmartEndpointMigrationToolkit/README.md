# Smart Endpoint Migration Toolkit

Version 1.0.0. Part of SmartM365 Devices by [WorkplaceCloudHub](https://workplacecloudhub.com).

Smart Endpoint Migration Toolkit runs endpoint migration **Campaigns** from one technician
computer, with one engine, one GUI and one set of protections:

| Campaign | What the endpoint script does |
| --- | --- |
| `HybridJoin` | Hybrid Entra join diagnostics and repair, then Intune (MDM) enrollment |
| `Windows11Upgrade` | Windows 10 to Windows 11 in-place upgrade (readiness, setup media, Windows Setup) |
| `Prerequisites` | Creates the scheduled tasks that the two other Campaigns start, enable or disable. Disabled by default (`Prerequisites.Enabled`). |

It replaces day-to-day use of IntuneHybridJoinToolkit and Windows11UpgradeToolkit. Those two
products stay installed and usable; nothing is migrated from them (see
[Docs/LEGACY-COEXISTENCE.md](Docs/LEGACY-COEXISTENCE.md)).

## Vocabulary

| Term | Meaning |
| --- | --- |
| Campaign | The kind of migration: `HybridJoin` or `Windows11Upgrade`, plus `Prerequisites` for their scheduled tasks. Options, Waves, Runs, history and GUI choices are kept per Campaign. |
| Wave | A named list of computers of one Campaign: `Waves\<Campaign>\<Wave>\Computers.txt`, its `Wave.local.json.txt` options and its Loop/Once launchers. |
| Run | One execution of a Wave (or of a Single PC): `Runs\<Campaign>\<Wave>\<RunId>\` with logs, reports, effective options and evidence. |
| Cycle | One pass over the computers still pending in a Run. A Loop Run repeats cycles until every computer is complete, the cycle limit is reached or the Run is stopped. |

A **Single PC** Run uses a one-computer Wave created inside its Run folder
(`Runs\<Campaign>\_SinglePC\<RunId>\Wave`) and goes through the same engine and protections.

The technician may set `Storage.WavesPath` and `Storage.RunsPath` in
`Config\SmartEndpointMigrationToolkit.local.json.txt` to independent absolute local or UNC paths.
Both default to empty: Waves and Runs then remain under the toolkit directory. These paths are
the Waves and Runs directories themselves, each retaining the Campaign subdirectories. Change
them only when no Runs are active. Existing content is not moved automatically; copy it if needed
and regenerate existing Wave launchers using the GUI or `Scripts\SmartM365-EndpointMigration-Update-WaveCmdWrappers.ps1`.
SEMT now reads and writes its own JSON data as `.json.txt`. It does not read or migrate earlier
`.json` configuration, Wave options, Run history or endpoint state. Configure the new files before
starting a Run; existing settings and guard history are not inherited from the old filenames.
Details: [Docs/CAMPAIGN-WAVE-RUN.md](Docs/CAMPAIGN-WAVE-RUN.md).

## Requirements

- Technician computer: an elevated console and network access (SMB 445 and administrative shares)
  to the targets. The default transport also needs remote Task Scheduler RPC access; WinRM needs
  PowerShell Remoting on the target; PsExec needs the Microsoft-signed Sysinternals binary.
- PowerShell (see "PowerShell support" below): Windows PowerShell 5.1 for the GUI (STA), the
  launchers and every endpoint script; the engine, the worker and the tools also accept
  PowerShell 7.2 and later.
- Inventories (optional, recommended): the ActiveDirectory module for AD exports, and Microsoft
  Graph PowerShell with delegated sign-in for Intune, Entra and user license exports. The user
  export requests `User.Read.All` and stores personal data in the local, Git-excluded Inventory folder.
- Targets: Windows 10/11 with the administrative share reachable by the technician account.
  Endpoint scripts are single self-contained files and do not need the technician module.

## PowerShell support

| Component | Supported | Refused before any state is written |
| --- | --- | --- |
| GUI | Windows PowerShell 5.1 with `-STA` (the CMD launchers start it so) | any other host: `SEMT_UNSUPPORTED_GUI_HOST` |
| Engine, worker, exporters, release tools (module `SmartEndpointMigration.Core`) | Windows PowerShell 5.1; PowerShell 7.2 and later | Windows PowerShell before 5.1, PowerShell 6.x, 7.0, 7.1: `SEMT_UNSUPPORTED_POWERSHELL` (module load) |
| Endpoint scripts (targets) | Windows PowerShell 5.1 (`powershell.exe`, as started by the worker and the Intune runner) | PowerShell 7: `FINAL_STATUS=UNSUPPORTED_POWERSHELL_HOST`, exit 3 |

Every CMD launcher of the package starts `%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe`
explicitly; none selects another engine. JSON dates are read the same way on every supported engine:
`-DateKind String` when the engine has it, otherwise the dates converted by PowerShell 7 are turned
back into round-trip UTC text, so reservations, process start checks and pending-operation markers
compare the same instants as on Windows PowerShell 5.1.

Qualified by execution: Windows PowerShell 5.1.26100 and PowerShell 7.6.6. The path used by
PowerShell 7.2 to 7.4 (no `-DateKind`) was executed on 7.6.6 by forcing it; those versions were not
installed.

## Start

From the root of the extracted standalone release:

```bat
Start-SmartEndpointMigrationToolkit-GUI.cmd
```

To refresh the shared users and licenses inventory independently, run
`Export-UserLicensesCsv.cmd`. A delegated Graph sign-in identifies the tenant before an existing
cache is reused. A verified same-tenant cache is reused for 24 hours; pass `-ForceRefresh` to
refresh it immediately, or pass `-TenantId <tenant-guid>` to select the tenant explicitly. The
cache contains personal data under `Inventory\UsersLicenses.csv` and is excluded from Git.

Offline demonstration with fictional Contoso data (nothing is launched or queried):

```bat
Start-SmartEndpointMigrationToolkit-Demo.cmd
```

GUI tabs: **Waves** (select, Launch Wave, launchers; on the right, the Wave **Options** tab, which edits
Wave.local.json.txt with inheritance, then **Activity**), **Single PC**, **New Wave**,
**Automatic Wave** (Preview, then Create: Create never launches), **Runs** (reports, controlled
stop, forced local stop) and **Options** (generated from the option catalog, saved per Campaign).
**Launch All Waves** launches the launchable Waves of the selected Campaign only, one every
5 seconds, after one confirmation.

Without the GUI, every Wave folder contains:

```text
Start-Wave-Loop.cmd                 Start-Wave-Once.cmd
Start-Wave-Loop-IgnoreRunGuard.cmd  Start-Wave-Once-IgnoreRunGuard.cmd
```

and the engine can be called directly:

```powershell
.\Scripts\SmartM365-Invoke-EndpointMigrationWithPsExec.ps1 -Campaign HybridJoin -WavePath .\Waves\HybridJoin\WAVE-PILOT -RunMode Once -DryRun
.\Scripts\SmartM365-Invoke-EndpointMigrationWithPsExec.ps1 -Campaign Windows11Upgrade -ComputerName PC-0142 -RunMode Once -ValidateOnly
.\Scripts\SmartM365-Invoke-EndpointMigrationWithPsExec.ps1 -Campaign Prerequisites -ComputerName PC-0142 -RunMode Once -AuditOnly -Option @{ 'Prerequisites.Enabled' = $true }
```

## Execution modes

| Mode | Technician side | Endpoint side |
| --- | --- | --- |
| Dry run (`DryRun`) | DNS, ping, SMB 445, administrative share | Nothing is copied or executed |
| Remote audit (`AuditOnly`) | Checks, copy, selected transport | Hybrid Join skips gpupdate; no repair, MDM cleanup, upgrade or reboot. It may flush DNS; it registers no scheduled task and runs nothing in the user context. Prerequisites only reports the tasks that are present, missing or different. |
| Repair (default) | Full checks, copy, selected transport | Applies the enabled actions only |

`RemoteExecutionMode` applies to both campaigns: `ScheduledTask` (default), `WinRM`, or
`PsExec`. The first uses remote Task Scheduler RPC. `WinRM` uses PowerShell Remoting to create
and monitor a local scheduled task; both run the endpoint as SYSTEM. The script is copied to a
protected per-attempt folder under Program Files. After confirmed task completion, SEMT checks
the task identity and payload hash, then deletes the task and the temporary script and folder.
Retry-after-reboot files remain in their separate protected folder. Transient Task Scheduler RPC
failures and network loss while the task is followed or deleted (for example 0x800706BE or
0x80070035, network path not found) are retried with
back-off and logged with a counter; when the connection comes back, the real endpoint result is
used and the task is cleaned up. Only failures lasting `RemoteTaskRpcFailureWindowMinutes`
(default 30) in a row give `REMOTE_EXECUTION_UNCONFIRMED`. If the connection or wait
expires, the target can still be running; cleanup is unconfirmed and must be checked before a
retry. No automatic fallback to another transport is performed. A blocked execution can persist
across transports; the selected mode never changes the endpoint security policy.
Select the transport in the GUI Options, or pass `-Option @{ RemoteExecutionMode = 'WinRM' }`
(or `'PsExec'`) to the command-line engine.

`-ValidateOnly` resolves options, checks rules and component integrity, writes the Run
manifest and contacts no target.

For Hybrid Join, both audit and repair perform a user-license preflight before copying or
starting the endpoint. When a fresh, complete same-tenant user inventory uniquely identifies
the console user and that user has no assigned SKU, SEMT records `NoAssignedLicense` and runs
the endpoint in audit-only mode even if repair was requested. The device diagnostic remains in
the report, while repair and enrollment are blocked until licensing is resolved. The computer
stays in the Wave and is not retried during the same Loop Run.
When the device join is unhealthy (device key or device authentication failure) but a local
Intune enrollment trace exists, Hybrid Join does not run any dsregcmd repair and reports
`INTUNE_TRACE_DEVICE_JOIN_UNHEALTHY` (NextAction `REVIEW_INTUNE_TRACE_AND_DEVICE_JOIN`) for a manual
review of the Entra object, the Intune record and the local enrollments. It is never counted as
enrolled. The endpoint keeps a local trace of it (`State\IntuneTraceRepair.json.txt`: first seen,
device ID, Intune enrollment IDs, device key and authentication state), removed once the join is
healthy. With `HybridJoin.RepairUnhealthyJoinWithIntuneTrace` (default false), a later run repairs
the device when all of these hold: the trace is at least `HybridJoin.IntuneTraceRepairDelayHours`
(24) old with the same device ID and enrollment IDs, the device join is still unhealthy, and the
launcher passes `-IntuneAbsentDeviceId` because fresh Entra and Intune inventories show no Intune
record for the name or for this device ID. The local Intune enrollment is backed up (reg export
under `Output`) and removed, then the normal device join repair runs. If the removal does not
complete, the status is `INTUNE_TRACE_REMOVE_FAILED` and nothing else is changed.

After each cycle, Hybrid Join confirms results with Intune and Entra exports scoped to the computers
that could have changed. A computer that its name does not find is looked up again by its local Entra
device ID (`Inventory\Cycle<n>_CloudRefreshDeviceIds.csv`), so a device whose cloud name differs is
still confirmed; such rows have `LookupMatchedBy=DeviceId` in the exported CSV.

Missing, stale or ambiguous evidence remains `Unknown` and is not reported as an absent
license. Refresh the users and licenses inventory from the GUI when needed; an assigned SKU
alone does not confirm that its Intune service plan is enabled.

## Prerequisites Campaign (scheduled tasks)

The `HybridJoin` and `Windows11Upgrade` endpoint scripts never create, change or delete a
scheduled task definition. Every task they use is created by the `Prerequisites` endpoint, either
run as its own Campaign beforehand or started automatically by the worker when a task is missing
(see "Automatic prerequisites" below):

| Task | Used by | Created |
| --- | --- | --- |
| `\WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin\SEMT-HJ-UserDsregStatus`, `SEMT-HJ-UserRefreshPrt` | HybridJoin (started on demand in the interactive user context) | enabled, no trigger, group INTERACTIVE. Only `dsregcmd /status` goes through `cmd.exe` (its output is read back); `dsregcmd.exe /refreshprt` is started directly, without output file: its result is checked through the user status |
| `\WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin\SEMT-HJ-RunUserAutoEnrollAtLogon` | HybridJoin (enabled only while auto-enrollment waits for a user logon, disabled once Intune enrollment is confirmed) | **disabled**, SYSTEM, at logon: starts `SEMT-HJ-UserDsregStatus` and runs `deviceenroller.exe /c /AutoEnrollMDM` |
| `\SmartM365-SEMT-HybridJoin-RetryAfterReboot`, `\SmartM365-SEMT-Windows11Upgrade-RetryAfterReboot` | the Campaign of the same name (enabled when it schedules a retry after a restart, disabled afterwards) | **disabled**, SYSTEM, at startup |

The retry-after-reboot task has a fixed command: it runs
`C:\Program Files\WCH\SmartM365\SmartEndpointMigrationToolkit\<Campaign>\RetryAfterReboot\RetryAfterRebootRunner.ps1`,
which reads the script path and arguments from `RetryAfterReboot.json.txt` next to it. The
Campaign writes the runner and its state, checks that the pre-created task runs exactly this
command as SYSTEM, and only then enables it. When a task is missing, the endpoint logs it and
carries on without it: Hybrid Join skips the user PRT status and refreshprt; a retry after reboot is not scheduled
(`RetryAfterRebootAction=ScheduleFailed`, `SEMT_PREREQUISITE_TASK_MISSING`), so no reboot is
triggered for it.

Hybrid Join runs `deviceenroller.exe /c /AutoEnrollMDM` itself, from SYSTEM, as the native
EnterpriseMgmt task does (for User Credential, Windows uses the signed-in user's token; started in
the user context the tool returns `0x80070005`). Its own exit code is logged. When enrollment is
still not confirmed and an MDM `UserLicense` / `0x80180018` fault was logged after the trigger, the
status is `INTUNE_USER_LICENSE_REJECTED` (NextAction `ASSIGN_USER_INTUNE_LICENSE`). When the tool
returns `0x8018000A` (already enrolled) but the only local Intune traces are incomplete stubs left by
an interrupted enrollment (`EnrollmentType` 6, empty `ProviderID`, `fooUser@` UPN, no OMA-DM account
or tasks), the stub blocks every new attempt: the status is `INTUNE_ENROLLMENT_STUB_BLOCKING` at once,
without the confirmation wait. With `HybridJoin.RemoveBlockingIntuneEnrollmentStub` (default false)
the stub is backed up and removed and the tool runs again. The Prerequisites
Campaign removes the obsolete `SEMT-HJ-UserMdmAutoEnroll` task (reported only with `AuditOnly`).

### Automatic prerequisites (`AutoPrerequisites`, default on)

Before a `HybridJoin` or `Windows11Upgrade` endpoint runs, the worker reads the target scheduled
tasks from the technician computer (remote Task Scheduler; read only, nothing started). It runs the
Prerequisites endpoint for this Campaign only (`-HybridJoinTasks` or `-Windows11UpgradeTasks`), with
the same transport and its own correlation ID, when a task above is missing, has another action or
no trigger, or when the obsolete `SEMT-HJ-UserMdmAutoEnroll` task remains. Then the Campaign
endpoint runs as usual. The report shows `PrerequisitesAutoState`:

| State | Meaning |
| --- | --- |
| `Ready` | All tasks present; the Prerequisites endpoint was not run. |
| `Created` | The Prerequisites endpoint ran first and returned `PREREQUISITES_READY`. |
| `MissingAuditOnly` | A task is missing or outdated; with `AuditOnly` (or a user license audit) nothing is created. |
| `CheckFailed` | The tasks could not be read (access, RPC); the Prerequisites endpoint was not run. |
| `Skipped` | The Prerequisites endpoint declined (virtual machine skipped, target busy or Windows Setup active). |
| `Failed`, `Unconfirmed`, `CopyFailed`, `IntegrityMismatch`, `Error` | The Prerequisites run did not confirm the tasks; details in `PrerequisitesAutoDetail` and the worker log. |

In every case the Campaign endpoint still runs and reports any task that is still missing. A dry
run checks nothing. This does not depend on `Prerequisites.Enabled`, which only controls the
standalone Campaign; set `AutoPrerequisites` to false to keep the previous behaviour.

The Campaign is **disabled by default**. Set `"Prerequisites.Enabled": true` in
`Config\SmartEndpointMigrationToolkit.local.json.txt` (or `SEMT_PREREQUISITES_ENABLED=1`, or
`-Option @{ 'Prerequisites.Enabled' = $true }`) to use it; until then the engine refuses it
(`SEMT_CAMPAIGN_DISABLED`, nothing written) and the GUI does not list it. A Wave or a GUI choice
cannot enable it. A task that already has the expected definition is left unchanged; a task with
another definition is updated and keeps its enabled state. Computers whose tasks are ready leave
the Wave (`Computers.PrerequisitesReady.txt`). The Campaign reads no inventory and has no
Automatic Wave.

## Options

All options come from [Config/SmartEndpointMigrationToolkit.options.json.txt](Config/SmartEndpointMigrationToolkit.options.json.txt).
Precedence, highest first: explicit CLI, GUI choices captured for the Run, Wave
(`Wave.local.json.txt`), `SEMT_*` environment variables, global configuration
(`Config\SmartEndpointMigrationToolkit.local.json.txt`), campaign default. The effective value and
its source are written to `EffectiveOptions.json.txt` and shown in the report. Reference and default
decisions: [Docs/OPTIONS.md](Docs/OPTIONS.md).

Local configuration files are created from their `*.json.template` and completed automatically
when a new key appears. They are never committed.

## Protections

- One shared worker cap for both Campaigns on this technician computer plus a per-Wave limit;
  reservations with heartbeat and stale recovery; the strictest active limit wins.
- One Run per Wave at a time; run guard history per Campaign; Windows 11 copy limits per source
  and per subnet.
- On the target: one Smart Endpoint Migration Toolkit operation at a time across Campaigns, and a
  pending Windows 11 setup blocks the other Campaigns. The previous toolkits are not inspected and
  never block.
- Every attempt has a CorrelationId; only the result with the same CorrelationId is trusted.
- Component integrity is checked against the distribution manifest shipped in the package.

This is a single-technician-computer coordination. It is **not** a multi-technician
coordination: two technician computers do not see each other's reservations.

## Folders

| Folder | Content | In Git |
| --- | --- | --- |
| `Scripts` | GUI, engine, worker, exporters, shared Wave runner | yes |
| `Modules` | `SmartEndpointMigration.Core` (technician-side engine module) | yes |
| `Campaigns\<Campaign>` | Campaign module, endpoint script, templates, Automatic Wave engine; Windows 11 Intune Win32 and setup media tooling | yes |
| `Config` | Option catalog and templates; local configuration | catalog and templates only |
| `Waves\<Campaign>\<Wave>` | Operational Waves | no |
| `Runs` | Runs, reports, GUI selections, Automatic Wave evidence | no |
| `Inventory` | Shared AD / Intune / Entra inventory caches | no |
| `Docs` | Architecture, options, procedures, test report, provenance | yes |
| `Tests` | Offline qualification suites | yes |
| `Release` | Release builder and template generator | yes (`Output`, `Work` excluded) |

Technician state (worker gate, Wave locks, run guard history, update-check cache) lives in
`C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\LauncherState`. Endpoint data lives
in `C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\<Campaign>` on each target. Every
script started as SYSTEM on a target (endpoint started by PsExec, Intune runner, retry-after-reboot
task) is stored under `C:\Program Files\WCH\SmartM365\SmartEndpointMigrationToolkit\<Campaign>`,
created with owner Administrators and a protected ACL or verified; ProgramData never holds a script
that is run, and the SEMT components of its data folder chain are protected (owner
Administrators, protected DACL, verified) before anything is written there; shared parents are
verified, never changed, and existing content without verified provenance is refused (see
Docs\ARCHITECTURE.md, "Program Files and ProgramData").

## Documentation

- [Docs/ARCHITECTURE.md](Docs/ARCHITECTURE.md)
- [Docs/CAMPAIGN-WAVE-RUN.md](Docs/CAMPAIGN-WAVE-RUN.md)
- [Docs/OPTIONS.md](Docs/OPTIONS.md)
- [Docs/LEGACY-COEXISTENCE.md](Docs/LEGACY-COEXISTENCE.md)
- [Docs/Windows11-Procedure-EN.md](Docs/Windows11-Procedure-EN.md) / [Docs/Windows11-Procedure-FR.md](Docs/Windows11-Procedure-FR.md)
- [Campaigns/Windows11Upgrade/IntuneWin32/README.md](Campaigns/Windows11Upgrade/IntuneWin32/README.md) (Intune Win32 packages and assignment switch order)
- [Docs/TEST-REPORT.md](Docs/TEST-REPORT.md)
- [Docs/ACCEPTANCE-PROCEDURE.md](Docs/ACCEPTANCE-PROCEDURE.md)
- [Docs/CRITICAL-FIXES-LEGACY.md](Docs/CRITICAL-FIXES-LEGACY.md)

## Release

`Release\SmartM365-EndpointMigration-Build-Release.ps1` validates the package (preview by
default) and, with `-Build`, writes into `Release\Output`:

- `SmartEndpointMigrationToolkit-1.0.0.zip` (deterministic, contains
  `SmartEndpointMigrationToolkit.components.json.txt` with the SHA256 of every file)
- `SmartEndpointMigrationToolkit-1.0.0.sha256`
- `RELEASE-NOTES-1.0.0.md`

Release tag: `endpoint-migration-toolkit-v1.0.0` (created manually, never by the builder).
Production releases are built from signed sources; `-AllowUnsigned` builds an offline
qualification package marked UNSIGNED.

## License

See the repository `LICENSE` and `NOTICE`.
