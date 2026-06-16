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
- `Update-LotCmdWrappers.cmd` refreshes the small CMD wrappers in every `LOT-*` folder.
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

Force a rerun that bypasses the target run guard:

```cmd
Run-IntuneHybridJoinRepairWithPsExec-Loop-IgnoreRunGuard.cmd
```

Refresh LOT wrappers after creating or importing LOT folders:

```cmd
Update-LotCmdWrappers.cmd
```

This also creates a blank `AdDomain.txt` in any `LOT-*` folder where it is missing.

Use `Export-ADDevicesCsv.cmd` from the toolkit root to create a forest-wide `DevicesAD.csv`.
LOT runs pass this root CSV separately and use it in priority when it exists and is less than
60 minutes old. If `AdDomain.txt` is missing or blank, the LOT refreshes the root `DevicesAD.csv`
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
After a LOT is selected, it shows the device count, AD scope, selected AD CSV, root inventory
freshness, wrapper status, and launch mode. A second tab creates a new empty LOT folder from a
LOT name, refreshes wrappers, creates `Computers.txt` and `AdDomain.txt`, and offers to open
`Computers.txt` so the operator can paste one computer per line.

## Repair Guardrails

The autonomous repair script is designed to avoid destructive actions unless diagnostic evidence supports them.

- `dsregcmd /leave` requires a valid base device identity: `AzureAdJoined=YES`, a `DeviceId`, and a `TenantId`.
- Broader leave behavior is controlled by explicit repair paths, such as failed device authentication or `KeySignTest=FAILED`.
- Intune enrollment checks distinguish strong Intune enrollment evidence from stale or weak local traces.
- MDM cleanup actions are opt-in and must be guarded by diagnostic state.
- The script writes local evidence so a LOT run can reclassify ambiguous PsExec exit states from collected remote CSV output when possible.

The repair script must remain self-contained. Do not add mandatory runtime dependencies on DeviceRegistrationTool or SmartM365 modules.

## Notes

- LOT folders should only contain `Computers.txt` and the small CMD wrappers.
- Reports are written under `LOT-*\Reports`.
- A live cycle CSV is written under `LOT-*\Reports` as computers complete.
- Per-computer PsExec logs are written under `LOT-*\PsExecLogs`.
- Collected remote evidence is written under `LOT-*\CentralLogs`.
- Already enrolled computers are moved from `Computers.txt` to `ComputersAlreadyEnrolled.txt`.
- Generated inventories, logs and reports are ignored by Git.
