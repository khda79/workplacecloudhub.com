# Architecture

## Components

```text
GUI (WPF, Windows PowerShell 5.1 STA)          Wave launchers (Start-Wave-*.cmd)
        |  generated elevated CMD                        |
        v                                                v
Scripts\SmartM365-Invoke-EndpointMigrationWithPsExec.ps1   (engine / launcher)
        |  imports Modules\SmartEndpointMigration.Core and Campaigns\<C>\SmartEndpointMigration.Campaign.<C>
        |  one versioned JSON contract per attempt
        v
Scripts\SmartM365-EndpointMigration-PsExecWorker.ps1      (one background job per computer)
        |  copy + hash check, PsExec, result by CorrelationId, evidence
        v
Campaigns\<C>\Scripts\SmartM365-Invoke-EndpointMigration-<C>.ps1   (single-file endpoint script on the target)
```

| Component | Role |
| --- | --- |
| Core module | Options, computer lists, concurrency (worker gate, Wave lock, cancellation), inventories, network preflight, Runs, worker contract, integrity, remote copy/PsExec, reports. Technician side only. |
| Campaign module | One profile per Campaign: endpoint script, option rules, pre-cycle and post-cycle hooks, per-computer arguments, completion rules, run guard policy, evidence rules, report columns. |
| Engine | Resolves options, checks integrity, prepares the Run, holds the Wave lock, runs cycles, launches workers under the shared cap, writes CSV/HTML reports and the Run history. |
| Worker | Executes one attempt from its contract. Refuses a contract of another `ContractVersion` or engine version. |
| Endpoint scripts | Self-contained, Windows PowerShell 5.1 only. They share the region "SEMT shared target guards v1", identical in both scripts (tested). The Windows 11 script also carries "SEMT setup media resolution v1", identical to the copy in the Intune package helper (tested). |

## Worker contract and results

- The engine writes `Runs\...\State\Contracts\<ContractId>.json` with `ContractVersion` (1),
  `EngineVersion`, `Campaign`, `RunId`, `CorrelationId`, `Computer`, endpoint, PsExec,
  evidence and flag sections. The worker validates every required field.
- The endpoint writes `<DataRoot>\State\Results\<CorrelationId>.json` when it starts
  (`State=Running`) and when it ends (`State=Completed`), and prints
  `FINAL_STATUS=...; EXIT_CODE=...; NEXT_ACTION=...; CORRELATION_ID=...; DETAIL=...`.
- The worker trusts only a result whose CorrelationId is its own. A result of another attempt is
  reported as foreign evidence and ignored. After a PsExec timeout or communication loss the
  worker waits for the completed result of the same attempt before classifying the computer.
- `LastRun.json` protections of the endpoint scripts are kept; retry-after-reboot runs use
  `<CorrelationId>-R<timestamp>`.
- Every computer row uses the common result schema (Campaign, Wave, Run, cycle, computer,
  CorrelationId, status, next action, launcher/remote status, network preflight, copy, hashes,
  PsExec, target guard, legacy activity, evidence paths, contract version) plus campaign
  extension fields.

## Integrity

- `SmartEndpointMigrationToolkit.components.json.txt` (written by the release builder inside the
  package) lists every file with its SHA256, size and version.
- The engine checks the module, campaign module, worker, engine, catalog and endpoint script
  against that manifest and their Authenticode signatures before any Run. The expected hash
  always comes from the shipped manifest, never from the caller. Mismatch blocks the Run
  (`SEMT_COMPONENT_INTEGRITY_BLOCKED`); `ComponentIntegrityMode=Warn` only reports it.
- The worker re-checks the endpoint script before copying it, and verifies the copied file's
  hash and version on the target share.
- Without a manifest (source tree) the check reports signatures only.

## Concurrency on the technician computer

| Mechanism | Scope | Storage |
| --- | --- | --- |
| Shared worker cap (`GlobalWorkerLimit`) | Both Campaigns, every session and account of this computer | `LauncherState\WorkerGate` (file lock + one lease file per worker) |
| Per-Wave limit (`ThrottleLimit`) | One Run | engine memory |
| Strictest limit | Running launchers started with different caps | the smallest active `Limit` wins |
| Stale recovery | Dead launcher/worker process (PID and start time), expired lease, missing heartbeat | removed under the gate lock, logged |
| Wave lock | One Run per Wave | `LauncherState\WaveLocks` exclusive handle |
| Run guard history | Per Campaign | `LauncherState\<Campaign>\RunGuardHistory.json` |
| Windows 11 copy gates | Per setup source and per subnet | UNC gate roots (see Windows 11 procedure) |

`LauncherState` ACLs are restricted to SYSTEM and Administrators when the engine runs elevated.
A `Global\` name alone is not relied upon: reservations are files with process identity, so a
crashed or other-session launcher is recognised.

Stopping is three separate actions:

1. **Controlled stop**: no new computer starts; running workers finish (`StopRequested_<pid>.json`).
2. **Forced local stop**: a second request stops local workers and their PsExec process tree.
3. **Remote state**: the endpoint may still be running; the report shows
   `CANCELLED_BY_OPERATOR_REMOTE_RUNNING` or `CANCELLED_BY_OPERATOR_REMOTE_STATE_UNCONFIRMED`,
   and the next Run reads the attempt result.

## Target-side guards

The shared region in both endpoint scripts provides:

- `Global\SmartM365_SEMT_TargetOperation` mutex plus a diagnostic `State\TargetOperation.json`
  lease: one toolkit operation at a time on a target, across Campaigns
  (`TARGET_OPERATION_ACTIVE`).
- `State\TargetPendingOperation.json`: written by Windows 11 when setup started or a reboot is
  pending (up to 168 h, cleared by reboot); it blocks the other Campaign
  (`TARGET_PENDING_OPERATION`). Unreadable state gives `TARGET_STATE_UNREADABLE`.
- `Test-SemtWindowsSetupActive`: Hybrid Join does not run during Windows Setup
  (`TARGET_WINDOWS_SETUP_ACTIVE`).
- Read-only legacy toolkit detection (see [LEGACY-COEXISTENCE.md](LEGACY-COEXISTENCE.md)).

## Inventories

Shared caches `Inventory\Devices{AD,Intune,Entra}.csv` are reused only when fresh (AD, Intune and Entra: 24 h) and, for Intune, when their provenance (tenant, delegated sign-in,
`AllManagedDevices`) is verified. A Run that needs fresher data exports a scoped inventory into
its own `Run\Inventory` folder; a scoped export never overwrites a shared cache.

User licensing evidence is collected separately into `Inventory\UsersLicenses.csv` using a
delegated `User.Read.All` session and paginated Graph user reads. It contains user identities
and assigned plans and is sensitive operational data. A complete cache is eligible for 24 h;
correlation requires an exact on-premises user SID and matching local tenant ID. An absent or
ambiguous SID, stale cache or unknown plan remains unknown. The MDM user scope is not inferred
from licensing and must be checked in Entra. Before a non-audit HybridJoin run, a bounded
read-only DCOM preflight compares the console SID, user-credential MDM policy, and a unique
registry tenant candidate with a fresh inventory. An unlicensed match stops the worker for
review before copying or running the endpoint. Registry tenant information can be stale: this
is a conservative review signal, not proof of the device tenant or the enrollment identity.
When the preflight fails or is ambiguous its state is Unknown and the run continues; this is
not evidence that licensing or scope is correct. A successful local enrollment is cloud confirmed
only when the local Entra device ID and tenant match exactly one Intune object. Name-only
inventory presence never archives a HybridJoin target. Scoped Intune exports retain duplicates.

For a targeted lookup after an audit, run `Scripts\SmartM365-EndpointMigration-Get-UserLicense.ps1`
with the confirmed session SID, device tenant ID, and a new output path in the Run directory. This
requires a delegated `User.Read.All` Graph session and permission to read license details. Errors,
empty or multiple SID matches and unavailable MDM scope remain unknown. The targeted output is
not a shared inventory. `HybridJoin.CollectMdmDiagnostics` explicitly requests a sensitive
DeviceEnrollment archive with a 300-second process timeout; the archive is copied only when central
evidence collection is enabled, within the configured size limit.

## Identities

| Item | Value |
| --- | --- |
| Environment variables | `SEMT_*`, `SEMT_HYBRIDJOIN_*`, `SEMT_WINDOWS11_*` |
| Endpoint mutex (Hybrid Join) | `Global\SmartM365_SEMT_HybridJoin_Endpoint` |
| Target mutex | `Global\SmartM365_SEMT_TargetOperation` |
| Retry tasks | `SmartM365-SEMT-HybridJoin-RetryAfterReboot`, `SmartM365-SEMT-Windows11Upgrade-RetryAfterReboot` |
| Intune Win32 (Windows 11) | see [IntuneWin32 README](../Campaigns/Windows11Upgrade/IntuneWin32/README.md) |
| Target data roots | `C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\<Campaign>` |
| Target program roots | `C:\Program Files\WCH\SmartM365\SmartEndpointMigrationToolkit\<Campaign>` (`RetryAfterReboot\`, `Intune\`) |

None of them is shared with IntuneHybridJoinToolkit or Windows11UpgradeToolkit.

## Program Files and ProgramData

Scripts that run as SYSTEM on a target are stored under Program Files, never under ProgramData,
because a standard user may create folders under ProgramData in advance and keep write access to
them:

| Script run as SYSTEM | Location |
| --- | --- |
| Endpoint started by PsExec | `<ProgramRoot>\<endpoint script>` (copied by the worker through the administrative share, SHA256 and version verified) |
| Intune runner, media helper, package manifest | `<ProgramRoot>\Intune\` |
| Endpoint started by the Intune runner | `<ProgramRoot>\<endpoint script>` |
| Retry-after-reboot runner, its state (script path and arguments) and the endpoint copy it starts | `<ProgramRoot>\RetryAfterReboot\` |

The region "SEMT trusted program folder v1" (identical in both endpoint scripts, the Intune
installer, the Intune runner and the Core module) creates missing folders with owner
Administrators and a protected DACL (SYSTEM and Administrators full control, Users read and
execute), verifies existing folders, and never repairs an existing folder that is not trusted:
the operation stops with `SEMT_PROGRAM_FOLDER_UNTRUSTED`. A folder is trusted when it and every
item below it are owned by SYSTEM, Administrators or TrustedInstaller and no other principal can
write, append, delete, change permissions or take ownership; unreadable permissions and reparse
points are not trusted. The runner and the retry task run verify their folder again before they
load or start anything. ProgramData keeps data only: logs, state, attempt results, setup media.

The data folder chain below ProgramData is protected as well, because the setup media cache that
it holds is run as SYSTEM and its logs and state are read by SYSTEM. Before anything is written
there:

- Scope: only the SEMT components (`SmartEndpointMigrationToolkit` and the campaign folder below
  it) are changed; they get owner Administrators and a protected DACL (SYSTEM and Administrators
  full control, Users read and execute). The shared parents `C:\ProgramData\WCH` and
  `...\WCH\SmartM365`, used by other products and by the legacy caches reused in place, are never
  changed: they are verified read-only (owner SYSTEM, Administrators or TrustedInstaller; no other
  principal may delete, delete children, change permissions or take ownership; creating items is
  allowed) and the operation stops when they are not protected.
- Read-only pass before the first change: shared parents, SEMT components, the provenance of the
  existing content and, when a DACL replacement is planned, every reparse point and hard link in
  everything it propagates to (other SEMT campaigns included). Any finding stops the operation
  with nothing changed.
- Provenance: a permission replacement never makes existing content trusted. An item that another
  principal could modify, or any content of a SEMT folder that was not protected, is refused and
  left unchanged for an administrator to review; it is never reset or adopted. A folder created
  here, or an unprotected one that was empty, must still be empty once protected.

Any failure is blocking: `SEMT_DATA_FOLDER_UNTRUSTED` for the installer (exit `1`, nothing written,
installed or registered) and `FINAL_STATUS=DATA_FOLDER_UNTRUSTED` for the endpoints (exit `1`,
nothing written under the data root). The Intune runner verifies the program folder and the data
chain (setup media included) before it writes, loads or starts anything; if the data chain is not
trusted it marks the package `RepairRequired` in the registry only, removes its task and stops.

Setup media is refused when its SHA256 manifest is missing or invalid (the endpoint checks the
result of the manifest validation), and, on Intune, when the reference manifest of the package is
required but missing: the comparison is never skipped. The manifest that comes with the media is
not, on its own, a proof of authenticity.

## PowerShell engines

The Core module refuses Windows PowerShell before 5.1 and PowerShell before 7.2 when it loads, so the
engine, the worker and the tools stop before writing any state. Endpoint scripts refuse any engine other
than Windows PowerShell 5.1 with `FINAL_STATUS=UNSUPPORTED_POWERSHELL_HOST`. The GUI requires Windows
PowerShell 5.1 with `-STA`. JSON dates: see README, "PowerShell support".
