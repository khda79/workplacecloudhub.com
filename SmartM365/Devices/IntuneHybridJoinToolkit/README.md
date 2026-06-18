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
LOT-X\
```

- `Export-IntuneDevicesCsv.cmd` exports the global Intune inventory to `DevicesIntune.csv`.
- `Export-EntraDevicesCsv.cmd` exports the global Entra device inventory to `DevicesEntra.csv`.
- `Export-ADDevicesCsv.cmd` exports AD computer objects to `DevicesAD.csv`. From the toolkit root, it exports all domains in the current AD forest by default; pass `-Domain <domain>` or set `EHJIR_AD_DOMAIN` to limit the export to one domain.
- `Start-IntuneHybridJoinRepair-LotLauncher-GUI.cmd` opens a GUI to create a local `LOT-*` folder from a computer list file and optionally launch it.
- `Update-LotCmdWrappers.cmd` refreshes the small CMD wrappers in operational `LOT-*` folders.
- `Scripts\` contains the shared PowerShell scripts and shared LOT launchers.
- `LOT-X\` is the neutral template for LOT folders.
- Operational `LOT-*` folders are ignored by Git because they can contain real computer lists.

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
Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd
```

Before launching a LOT, the CMD and PowerShell launchers verify that `PsExec.exe` is available
in `Scripts\PsExec.exe`, `%WINDIR%\System32\PsExec.exe`, or in `PATH`. Direct PowerShell runs can also pass `-PsExecPath`.
The LOT Launcher GUI only checks `Scripts\PsExec.exe` and `%WINDIR%\System32\PsExec.exe` during window startup to keep the UI responsive;
if the local files are missing, `PATH` is checked when the operator launches the LOT.

At the beginning of a LOT run, the launcher archives previous `CentralLogs`, `PsExecLogs`, and `Reports`
folders into `Archives\IntuneHybridJoinToolkit_PreRun_<timestamp>.zip`, then removes those folders before
the new cycle starts. Set `EHJIR_SKIP_PRE_RUN_ARCHIVE=1` or pass `-SkipPreRunArchive` only when previous
outputs must stay in place.

If DNS resolution fails on every sampled computer during preflight, the LOT stops before queuing PsExec
jobs and writes `DNS_PREFLIGHT_ALL_SAMPLES_FAILED` rows. Set `EHJIR_CONTINUE_ON_DNS_PREFLIGHT_FAILURE=1`
or pass `-ContinueOnDnsPreflightFailure` only after confirming that the network path is intentional.

By default, the launcher does not start a new cycle during the local night window from 20:00 to 07:00.
This pause is checked before the first cycle and before every later cycle; it does not interrupt a cycle
already running. Set `EHJIR_DISABLE_NIGHT_PAUSE=1` or pass `-DisableNightPause` to allow cycles at night.
The window can be adjusted with `EHJIR_NIGHT_PAUSE_START_HOUR` and `EHJIR_NIGHT_PAUSE_END_HOUR`.

LOT runs use two concurrency controls:

- `EHJIR_THROTTLE` / `-ThrottleLimit` limits parallel computers inside one LOT.
- `EHJIR_GLOBAL_CONCURRENCY_LIMIT` / `-GlobalConcurrencyLimit` limits active computer workers shared by all LOT windows on the same operator session. The default is 15, so launching multiple LOTs does not multiply PsExec/PowerShell load indefinitely.
- `EHJIR_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES` / `-GlobalConcurrencyLeaseTimeoutMinutes` controls when an abandoned shared worker lease is considered stale and cleaned. The default `0` sizes this automatically from the PsExec and delayed-evidence timeouts. Leases record the launcher PID, job id, worker PowerShell PID, computer, and cycle; stale slots are repaired instead of bypassed.
- The LOT Launcher GUI starts `Launch all` LOT windows 5 seconds apart to avoid a local startup spike.

LOT wrappers skip detected virtual machines by default before remote copy or repair. To include VMs in a direct LOT launch:

```cmd
set EHJIR_SKIP_VIRTUAL_MACHINES=0
Run-IntuneHybridJoinRepairWithPsExec-Loop.cmd
```

LOT wrappers also enable the guarded repair defaults used by the GUI:

- `EHJIR_ALLOW_DSREG_LEAVE=1`
- `EHJIR_ALLOW_REMOVE_STALE_INTUNE_ENROLLMENT=1`
- `EHJIR_ALLOW_REBOOT_WHEN_NO_INTERACTIVE_USER=1`
- `EHJIR_ALLOW_REBOOT_AFTER_DSREG_LEAVE=1`

Set any of these values to `0` before launching a LOT to disable that action.

Force a rerun that bypasses the target run guard:

```cmd
Run-IntuneHybridJoinRepairWithPsExec-Loop-IgnoreRunGuard.cmd
```

Refresh LOT wrappers after creating or importing operational LOT folders:

```cmd
Update-LotCmdWrappers.cmd
```

This skips the versioned `LOT-X` template and creates a blank `AdDomain.txt` in any operational
`LOT-*` folder where it is missing.

Use `Export-ADDevicesCsv.cmd` from the toolkit root to create a forest-wide `DevicesAD.csv`.
LOT runs pass this root CSV separately and use it in priority when it exists and is less than
120 minutes old. If `AdDomain.txt` is missing or blank, the LOT refreshes the root `DevicesAD.csv`
as a forest-wide AD export. A LOT can still use a per-LOT AD domain by setting `EHJIR_AD_DOMAIN`
before launching the LOT, or by creating an `AdDomain.txt` file in that LOT folder with the domain
name on the first line. In that domain-specific fallback mode, the repair launcher writes and
refreshes `LOT-*\DevicesAD.csv` so different LOT folders can target different AD domains without
overwriting each other's AD inventory.

Open the LOT launcher GUI:

```cmd
Start-IntuneHybridJoinRepair-LotLauncher-GUI.cmd
```

The GUI has an existing-LOT tab with a drop-down list of available operational `LOT-*` folders.
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
- Reports are written under `LOT-*\Reports`.
- A live cycle CSV is written under `LOT-*\Reports` as computers complete.
- Per-computer PsExec logs are written under `LOT-*\PsExecLogs`.
- Collected remote evidence is written under `LOT-*\CentralLogs`.
- Previous `CentralLogs`, `PsExecLogs`, and `Reports` folders are archived under `LOT-*\Archives` at the start of a LOT run by default.
- Already enrolled computers are moved from `Computers.txt` to `ComputersAlreadyEnrolled.txt`.
- Generated inventories, logs and reports are ignored by Git.
