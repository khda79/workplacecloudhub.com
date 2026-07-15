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
- `Scripts\SmartM365-IntuneHybridJoinRepair-LotLauncher-GUI.ps1`: GUI that launches existing LOT folders and creates new empty LOT folders ready for `Computers.txt`.

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
in `Scripts\PsExec.exe`, `%WINDIR%\System32\PsExec.exe`, or in `PATH`. Direct PowerShell runs can also pass `-PsExecPath`. The shared CMD and GUI resolve Windows PowerShell first, with PowerShell 7 as a fallback, instead of assuming `pwsh.exe` is installed.
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
- A technician-side inter-LOT guard is enabled by default (`EHJIR_USE_TECHNICIAN_RUN_GUARD_HISTORY=1`). It records active or successful launches by normalized AD FQDN under `%ProgramData%\SmartM365\IntuneHybridJoinToolkit\LauncherState\RunGuardHistory.json` and expires after 3 hours. Network, DNS, admin-share, PsExec, and collection failures remain immediately retryable.
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

For either authorized controlled reboot, the endpoint first creates the SYSTEM startup task `SmartM365-IntuneHybridJoinToolkit-RetryAfterReboot`. Its state and runner are stored under `%ProgramData%\SmartM365\IntuneHybridJoinToolkit\State`, protected for SYSTEM and local Administrators. The resumed run bypasses only the local run guard, waits 300 seconds by default, and is removed after success, a non-reboot outcome, a scheduling failure, or three exhausted startup attempts. Tune this with `EHJIR_RETRY_AFTER_REBOOT_DELAY_SECONDS` / `-RetryAfterRebootDelaySeconds` and `EHJIR_RETRY_AFTER_REBOOT_MAX_ATTEMPTS` / `-RetryAfterRebootMaxAttempts`.

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
120 minutes old. If `AdDomain.txt` is missing or blank, the LOT refreshes the root `DevicesAD.csv`
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
stay disabled. Detailed paths and CSV freshness stay in logs and CLI output. A second tab creates a new empty LOT folder from a
LOT name, refreshes wrappers, creates `Computers.txt` and `AdDomain.txt`, and offers to open
`Computers.txt` so the operator can paste one computer per line.

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
- The HTML report accumulates all cycles, distinguishes the latest status for each unique computer from all recorded attempts, and shows cycle progress, running jobs, duplicate input statistics, effective options, security evidence, and Hybrid Join inventory details.
- Endpoint CSV, `LastRun.json`, live CSV, and HTML details expose the retry-after-reboot action, detail, attempt, maximum attempts, and scheduled task name.
- The LOT Launcher GUI watches the run report folder and opens the first non-empty merged HTML report automatically.
- Per-computer PsExec logs are written under `Runs\<LOT>\<yyyyMMdd-HHmmss>\PsExecLogs`. Their filenames use the computer short name plus an eight-character hash of the normalized connection target, which keeps long FQDNs distinct while avoiding Windows path-length failures.
- The complete timestamped launcher transcript and `SmartM365-IHJ-Launcher_latest.log` are written under the run `Logs` folder.
- Automatic initial and post-cycle Intune, Entra, and AD refreshes are limited to the current LOT computer list and write scoped CSV/log evidence under the run `Reports` folder; global root inventory caches are never overwritten by a LOT run.
- Collected remote evidence is written under `Runs\<LOT>\<yyyyMMdd-HHmmss>\CentralLogs\<Success|Errors|AdminShareFailure|RemoteCollectionFailure>\<short-name-hash>`. Standard mode skips individual remote files larger than 5 MB; set `EHJIR_CENTRAL_LOG_COLLECTION_MODE=Full` to collect them.
- Previous `CentralLogs`, `PsExecLogs`, and `Reports` folders are archived under `Runs\<LOT>\<yyyyMMdd-HHmmss>\Archives` at the start of a LOT run by default.
- Already enrolled computers are moved from `Computers.txt` to `ComputersAlreadyEnrolled.txt`.
- Generated inventories, logs and reports are ignored by Git.
