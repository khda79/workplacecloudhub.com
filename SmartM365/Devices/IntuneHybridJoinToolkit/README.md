# Smart Intune Hybrid Join Toolkit

PowerShell toolkit to help diagnose and repair Windows devices that should be Hybrid Entra joined and enrolled in Intune after MDM auto-enrollment policy is applied.

The toolkit is optimized for batch operations. It keeps `Scripts\SmartM365-Invoke-IntuneHybridJoinRepair.ps1` autonomous so the same single PowerShell file can be copied to devices through PsExec/LOT folders, pushed with a GPO, or reused by an operator without requiring SmartM365 modules on the target computer.

For single-device support with a richer GUI, support bundle, and CLI export experience, use `../DeviceRegistrationTool/`. Both tools intentionally share the same diagnostic and guarded repair concepts for Hybrid Join and Intune enrollment; when one side changes, the other should be reviewed for synchronization.

## Layout

```text
Export-IntuneDevicesCsv.cmd
Export-EntraDevicesCsv.cmd
Export-ADDevicesCsv.cmd
Start-IntuneHybridJoinRepair-LotLauncher-GUI.cmd
Update-LotCmdWrappers.cmd
Scripts\
Lots\LOT-TEMPLATE\
```

- `Export-IntuneDevicesCsv.cmd` exports the global Intune inventory to `DevicesIntune.csv`.
- `Export-EntraDevicesCsv.cmd` exports the global Entra device inventory to `DevicesEntra.csv`.
- `Export-ADDevicesCsv.cmd` exports AD computer objects to `DevicesAD.csv`. From the toolkit root, it exports all domains in the current AD forest by default; pass `-Domain <domain>` or set `EHJIR_AD_DOMAIN` to limit the export to one domain.
- `Start-IntuneHybridJoinRepair-LotLauncher-GUI.cmd` opens a GUI to create a local `Lots\LOT-*` folder from a computer list file and optionally launch it.
- `Update-LotCmdWrappers.cmd` refreshes the small CMD wrappers in operational `Lots\LOT-*` folders.
- `Scripts\` contains the shared PowerShell scripts and shared LOT launchers.
- `Lots\LOT-TEMPLATE\` is the neutral template for LOT folders.
- Operational `Lots\LOT-*` folders contain LOT configuration and computer lists. Runtime data is written under `Runs\<LOT>\<yyyyMMdd-HHmmss>` so logs, reports, collected central logs, archives, and per-run inventory CSV files stay separate from configuration.

## Scripts

- `Scripts\SmartM365-Invoke-IntuneHybridJoinRepair.ps1`: remote repair script copied to target devices and executed as SYSTEM.
- `Scripts\SmartM365-Invoke-IntuneHybridJoinRepairWithPsExec.ps1`: local PsExec orchestrator for LOT folders.
- `Scripts\SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1`: full Intune managed-device inventory export.
- `Scripts\SmartM365-IntuneHybridJoinRepair-Export-EntraDevicesCsv.ps1`: full Entra device inventory export.
- `Scripts\SmartM365-IntuneHybridJoinRepair-Export-ADDevicesCsv.ps1`: AD computer inventory export for LOT report enrichment.
- `Scripts\SmartM365-IntuneHybridJoinRepair-Update-LotCmdWrappers.ps1`: refreshes small LOT CMD wrappers.
- Scripts\SmartM365-IntuneHybridJoinRepair-LotLauncher-GUI.ps1: GUI that launches existing LOT folders and creates manual or automatic LOT folders.
- Scripts\SmartM365-IntuneHybridJoinRepair-New-AutomaticLot.ps1: guarded AD/Intune/Entra selection and evidence engine.
- Scripts\SmartM365-IntuneHybridJoinRepair-Export-AutomaticGraphInventories.ps1: delegated full Intune and Entra refresh using one Graph session.

## When To Use This Toolkit

Use this toolkit when:

- you have a batch of computers to process from an operator workstation;
- you need PsExec-based SYSTEM execution on remote devices;
- you need a LOT folder with `Computers.txt`, live cycle reports, PsExec logs, and collected remote evidence;
- you need the repair logic as one portable script file for GPO or another file-copy deployment method.

Use Smart DeviceRegistration Tool instead when:

- the operator is working on one device locally;
- a GUI support experience is preferred;
- the action should collect a support bundle or support summary;
- a user-mode diagnostic-only workflow is required.

## Usage

Export the global Intune inventory from the repository root:

```cmd
Export-IntuneDevicesCsv.cmd
```

Run repair from a LOT folder:

```cmd
Lots\LOT-A\Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd
```

Before launching a LOT, the CMD and PowerShell launchers verify that `PsExec.exe` is available
in `Scripts\PsExec.exe`, `%WINDIR%\System32\PsExec.exe`, or in `PATH`. Direct PowerShell runs can also pass `-PsExecPath`. The shared LOT CMD and single-computer launcher prefer PowerShell 7 when available, with Windows PowerShell 5.1 as fallback. The WPF GUI itself remains on Windows PowerShell 5.1/STA.
The LOT Launcher GUI only checks `Scripts\PsExec.exe` and `%WINDIR%\System32\PsExec.exe` during window startup to keep the UI responsive;
if the local files are missing, `PATH` is checked when the operator launches the LOT.

At the beginning of a LOT run, the launcher archives previous `CentralLogs`, `PsExecLogs`, and `Reports`
folders into `Archives\IntuneHybridJoinToolkit_PreRun_<timestamp>.zip`, then removes those folders before
the new cycle starts. Set `EHJIR_SKIP_PRE_RUN_ARCHIVE=1` or pass `-SkipPreRunArchive` only when previous
outputs must stay in place.

If DNS resolution fails on every sampled computer during preflight, the LOT stops before queuing PsExec
jobs and writes `DNS_PREFLIGHT_ALL_SAMPLES_FAILED` rows. Set `EHJIR_CONTINUE_ON_DNS_PREFLIGHT_FAILURE=1`
or pass `-ContinueOnDnsPreflightFailure` only after confirming that the network path is intentional.

If user-context Hybrid Join or Intune enrollment evidence is required, the remote repair script creates or updates
on-demand scheduled tasks locally on the target device under `\SmartM365\IntuneHybridJoinToolkit`. These tasks
replace the former externally managed helper tasks and are used only to run user-context actions while the main repair
continues to run as SYSTEM through PsExec:

- `SmartM365-IHJ-UserDsregStatus`
- `SmartM365-IHJ-UserRefreshPrt`
- `SmartM365-IHJ-UserMdmAutoEnroll`
- `SmartM365-IHJ-RunUserAutoEnrollAtLogon` when a next-logon user auto-enrollment helper is required. This task
  stores direct `schtasks.exe /Run` actions in the scheduled task definition and removes the former local `.cmd` helper if present.

Task output is written under `C:\Windows\Temp\SmartM365-IHJ-*` and copied into the run evidence when available.

By default, the launcher does not start a new cycle during the local night window from 20:00 to 07:00.
This pause is checked before the first cycle and before every later cycle; it does not interrupt a cycle
already running. Set `EHJIR_DISABLE_NIGHT_PAUSE=1` or pass `-DisableNightPause` to allow cycles at night.
The window can be adjusted with `EHJIR_NIGHT_PAUSE_START_HOUR` and `EHJIR_NIGHT_PAUSE_END_HOUR`.

The PsExec launcher also prevents the same LOT folder from running twice at the same time on the same
operator session. This protects `CentralLogs`, `PsExecLogs`, `Reports`, and `Computers.txt` from
overlapping launches whether the LOT is started from the GUI, a wrapper CMD, or PowerShell. The emergency
override is `EHJIR_DISABLE_LOT_RUN_MUTEX=1` or `-DisableLotRunMutex`.

LOT runs use two concurrency controls:

- `EHJIR_THROTTLE` / `-ThrottleLimit` limits parallel computers inside one LOT.
- `EHJIR_GLOBAL_CONCURRENCY_LIMIT` / `-GlobalConcurrencyLimit` limits active computer workers shared by all LOT windows on the same operator session. The default is 15, so launching multiple LOTs does not multiply PsExec/PowerShell load indefinitely.
- `EHJIR_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES` / `-GlobalConcurrencyLeaseTimeoutMinutes` controls when an abandoned shared worker lease is considered stale and cleaned. The default `0` sizes this automatically from the PsExec and delayed-evidence timeouts. Leases record the launcher PID, job id, worker PowerShell PID, computer, and cycle; stale slots are repaired instead of bypassed.
- The LOT Launcher GUI starts `Launch all` LOT windows 5 seconds apart to avoid a local startup spike.
- A technician-side inter-LOT status backoff is enabled by default (`EHJIR_USE_TECHNICIAN_RUN_GUARD_HISTORY=1`) under `%ProgramData%\SmartM365\IntuneHybridJoinToolkit\LauncherState\RunGuardHistory.json`. The base guard is 12 hours; network/admin-share failures use a 15-minute exponential cooldown capped at 6 hours, completed endpoint states align with the 12-hour endpoint guard, and user/logon states wait 24 hours.
- Use `EHJIR_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY=1` for an explicit bypass, or adjust `EHJIR_TECHNICIAN_RUN_GUARD_HOURS`.

LOT wrappers skip detected virtual machines by default before remote copy or repair. To include VMs in a direct LOT launch:

```cmd
set EHJIR_SKIP_VIRTUAL_MACHINES=0
Lots\LOT-A\Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd
```

LOT wrappers also enable the guarded repair defaults used by the GUI:

- `EHJIR_ALLOW_DSREG_LEAVE=1`
- `EHJIR_ALLOW_REMOVE_STALE_INTUNE_ENROLLMENT=1`
- `EHJIR_ALLOW_REBOOT_WHEN_NO_INTERACTIVE_USER=1`
- `EHJIR_ALLOW_REBOOT_AFTER_DSREG_LEAVE=1`

Set any of these values to `0` before launching a LOT to disable that action.

`WMI_Bridge_SCCM_Server` is treated as a protected WMI-to-CSP management bridge marker, not as a confirmed
competing MDM enrollment. When it is the only external provider marker, it is reported through
`ManagementBridgeDetected`, `ManagementBridgeEnrollmentIds`, `ManagementBridgeProviderIds`, and
`ManagementBridgeDetails` without blocking Intune auto-enrollment. `-AllowRemoveNonIntuneMdmEnrollment`
never removes this bridge registration. Other external providers block auto-enrollment only when an external
discovery URL, OMADM account, or EnterpriseMgmt task confirms OMA-DM enrollment.

For either authorized controlled reboot, the endpoint first creates the SYSTEM startup task `SmartM365-IntuneHybridJoinToolkit-RetryAfterReboot`. Its protected state remains bounded to three startup attempts. A separate durable `State\RebootSafety.json` counter prevents new LOT runs from starting unlimited reboot chains; user-credential reboots additionally require a changed interactive user/session context before another reboot can be authorized.

Every endpoint run also owns the named mutex `Global\SmartM365_IntuneHybridJoinToolkit_Endpoint`. `-IgnoreRunGuard` never bypasses a genuinely active process. `State\EndpointInstance.json` records PID, RunId, start time and heartbeat; a dead process releases the mutex automatically and the next run replaces stale metadata.

Force a rerun that bypasses the target run guard:

```cmd
Lots\LOT-A\Run-IntuneHybridJoinRepairWithPsExec-Loop-IgnoreRunGuard.cmd
```

Refresh LOT wrappers after creating or importing operational LOT folders:

```cmd
Update-LotCmdWrappers.cmd
```

This skips the versioned `Lots\LOT-TEMPLATE` template and creates a blank `AdDomain.txt` in any operational
`Lots\LOT-*` folder where it is missing.

Use `Export-ADDevicesCsv.cmd` from the toolkit root to create a forest-wide `DevicesAD.csv`.
LOT runs pass this root CSV separately and use it in priority when it exists and is less than
12 hours old. If `AdDomain.txt` is missing or blank, the LOT refreshes the root `DevicesAD.csv`
as a forest-wide AD export. A LOT can still use a per-LOT AD domain by setting `EHJIR_AD_DOMAIN`
before launching the LOT, or by creating an `AdDomain.txt` file in that LOT folder with the domain
name on the first line. In that domain-specific fallback mode, the repair launcher writes and
refreshes a scoped `DevicesAD_InitialScoped_*.csv` or `DevicesAD_Cycle*Refresh_*.csv` under the run `Reports` folder so different LOT folders can target different AD domains without
overwriting each other's AD inventory.

Open the LOT launcher GUI:

```cmd
Start-IntuneHybridJoinRepair-LotLauncher-GUI.cmd
```

The GUI has an existing-LOT tab with a drop-down list of available operational `Lots\LOT-*` folders.
After a LOT is selected, it shows only the device count, AD scope, global worker limit, and launch mode.
It can launch the selected LOT or all launchable LOT folders; empty LOT folders or folders with
missing wrappers are skipped. When no operational LOT exists, the LOT selector and launch buttons
stay disabled. Detailed paths and CSV freshness stay in logs and CLI output. **Single PC** launches
one explicit target. **New LOT** creates an empty folder, refreshes wrappers, creates `Computers.txt`
and `AdDomain.txt`, and offers to open the computer list. **Automatic LOT** builds a guarded list
from inventory as documented below.

## Automatic Hybrid Join LOT

Open **Automatic LOT** in the WPF launcher to build a LOT from broad inventory snapshots. The
selection is intentionally conservative:

- the computer must be present and enabled in AD;
- AD must explicitly identify Windows 10 or Windows 11 client; Windows Server and unknown OS
  values are excluded;
- every matching Intune managed-device row excludes the computer;
- AD short-name collisions, multiple `ServerAd` Entra objects, and disabled `ServerAd` objects
  are excluded as ambiguous;
- the preferred `DNSHostName` is written to `Computers.txt`, with the AD short name as fallback.

Entra evidence classifies a selected device as `NEEDS_HYBRID_JOIN`, `HYBRID_JOIN_PENDING`, or
`NEEDS_INTUNE_ENROLLMENT`. Entra is optional enrichment: when it cannot be read, otherwise-safe
AD devices absent from Intune remain selectable as `ENTRA_INVENTORY_UNAVAILABLE`; the endpoint
script then relies on its guarded local diagnosis. AD and Intune are mandatory, so the preview
stops when either source is unavailable.

Optional semicolon-separated **Prefix(es)** and **Contains** filters use literal values. Values
inside each field use OR, while Prefix and Contains are combined with AND. **Exclude stale AD**
is disabled by default; when enabled, its default maximum LastLogon age is 45 days and unknown
or invalid timestamps are also excluded.

Every preview checks read-only root caches before refreshing:

- `DevicesAD.csv`: 12 hours;
- `DevicesIntune.csv`: 2 hours;
- `DevicesEntra.csv`: 2 hours.

Graph caches are accepted only when their CSV provenance contains a non-empty tenant, delegated
authentication, and the full inventory scope (`AllManagedDevices` or `AllEntraDevices`). Intune
and Entra caches must identify the same tenant. **Force inventory refresh this time** bypasses
valid caches for the next preview only. Fresh or newly generated sources are copied under
`Runs\AutomaticLotInventory\Sources-<timestamp>`; automatic refresh never overwrites root caches.

Graph refresh uses delegated interactive authentication only. The required read scopes are:

```text
DeviceManagementManagedDevices.Read.All
Device.Read.All
```

App-only Graph contexts are rejected. The combined refresh connects once with both scopes,
exports required Intune data first, then attempts optional Entra enrichment.

The **Create** button creates `Lots\LOT-AUTO-IHJ-...`, `Computers.txt`, `AdDomain.txt`, and the
standard wrappers. It never launches the LOT. Open **Existing LOT** to review and launch it.
A modal progress window remains visible during cache checks, inventory refresh, filtering,
evidence generation, and wrapper refresh.

Each preview or creation preserves:

```text
Runs\AutomaticLotInventory\<timestamp>\DevicesAD.csv
Runs\AutomaticLotInventory\<timestamp>\DevicesIntune.csv
Runs\AutomaticLotInventory\<timestamp>\DevicesEntra.csv       # when available
Runs\AutomaticLotInventory\<timestamp>\AutomaticLotSelection.csv
Runs\AutomaticLotInventory\<timestamp>\AutomaticLotExclusions.csv
Runs\AutomaticLotInventory\<timestamp>\AutomaticLotFilterExclusions.csv
Runs\AutomaticLotInventory\<timestamp>\AutomaticLotSummary.json
```
## Controlled LOT Stop

Each LOT launcher publishes an active-run control file in its run `State` folder.

- Press Ctrl+C once, or click **Stop running LOTs** once in the GUI, to stop queueing new computers. Computers that have not started are reported as `CANCELLED_NOT_STARTED`, while active jobs are allowed to finish for up to 15 minutes by default.
- Press Ctrl+C a second time, or click the GUI stop button a second time, to stop the remaining local workers. In-flight computers are reported as `CANCELLED_BY_OPERATOR`; the remote endpoint may already be continuing, so verify its logs and state before relaunching.
- Change the drain window with `EHJIR_CANCELLATION_DRAIN_TIMEOUT_MINUTES` or `-CancellationDrainTimeoutMinutes` (0 to 1440 minutes).
- If PowerShell terminates natively with `0xC0000005`, the CMD writes `PowerShellCrash_<timestamp>.txt` in the run `Logs` folder with recent Application events 1000, 1001, and 1023.
- When the GUI closes while LOTs are active, it asks whether to request a controlled stop, leave the LOTs running, or cancel the close.

Disconnecting an RDP session does not normally stop the launcher. Signing out of Windows, closing its console host, ending the PowerShell process, or shutting down the operator computer is abrupt: final CSV/HTML cleanup cannot be guaranteed, and remote endpoint work may continue. Request a controlled stop and wait for the final report before signing out.

## Repair Guardrails

The autonomous repair script is designed to avoid destructive actions unless diagnostic evidence supports them.

- `dsregcmd /leave` requires a valid base device identity: `AzureAdJoined=YES`, a `DeviceId`, and a `TenantId`.
- Broader leave behavior is controlled by explicit repair paths, such as failed device authentication or `KeySignTest=FAILED`.
- Intune enrollment checks distinguish strong Intune enrollment evidence from stale or weak local traces.
- MDM cleanup actions are opt-in and must be guarded by diagnostic state.
- The script writes local evidence so a LOT run can reclassify ambiguous PsExec exit states from collected remote CSV output when possible.
- `-SkipVirtualMachines` records `SKIPPED_VIRTUAL_MACHINE` and performs no DNS/domain/gpupdate/repair action on detected VMs.

The repair script must remain self-contained. Do not add mandatory runtime dependencies on DeviceRegistrationTool or SmartM365 modules.

## Notes

- LOT folders should only contain `Computers.txt` and the small CMD wrappers.
- Reports are written under `Runs\<LOT>\<yyyyMMdd-HHmmss>\Reports`.
- A live cycle CSV is written under `Runs\<LOT>\<yyyyMMdd-HHmmss>\Reports` as computers complete, and a final CSV remains available for every cycle.
- One HTML report named `PsExec_IntuneHybridJoinRepair_Summary_<LOT>_<timestamp>.html` is created for the complete LOT run. The same file is refreshed while jobs run and finalized when the launcher stops.
- The HTML report accumulates attempt summaries but renders detail tables only for the latest actionable row per unique computer. It distinguishes latest attempt from latest actionable status so cancellation, run-guard, and backoff rows do not mask the last meaningful result; the complete attempt history remains in cycle CSV files.
- Endpoint CSV, `LastRun.json`, live CSV, and HTML details expose the retry-after-reboot action, detail, attempt, maximum attempts, and scheduled task name.
- The LOT Launcher GUI watches the run report folder and opens the first non-empty merged HTML report automatically.
- Per-computer PsExec logs are written under `Runs\<LOT>\<yyyyMMdd-HHmmss>\PsExecLogs`. Their filenames use the computer short name plus an eight-character hash of the normalized connection target, which keeps long FQDNs distinct while avoiding Windows path-length failures.
- The complete timestamped launcher transcript and `SmartM365-IHJ-Launcher_latest.log` are written under the run `Logs` folder.
- Initial and post-cycle inventories remain LOT-scoped. Automatic post-cycle Intune and Entra refreshes are skipped only when every cycle result proves that no remote action could have changed cloud inventory. AD refresh behavior and global root inventory caches are unchanged.
- Collected remote evidence is written under `Runs\<LOT>\<yyyyMMdd-HHmmss>\CentralLogs\<Success|Errors|AdminShareFailure|RemoteCollectionFailure>\<short-name-hash>`. Standard mode keeps current-run files and required state only, skips files larger than 5 MB, removes stale outcome buckets, and uses one physical `Latest` copy for both report links. `EHJIR_CENTRAL_LOG_COLLECTION_MODE=Full` retains the complete supported endpoint evidence set.
- Previous `CentralLogs`, `PsExecLogs`, and `Reports` folders are archived under `Runs\<LOT>\<yyyyMMdd-HHmmss>\Archives` at the start of a LOT run by default.
- Already enrolled computers are moved from `Computers.txt` to `ComputersAlreadyEnrolled.txt`.
- Generated inventories, logs and reports are ignored by Git.
