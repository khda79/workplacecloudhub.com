# Smart Intune Hybrid Join Toolkit

PowerShell toolkit to help diagnose and repair Windows devices that should be Hybrid Entra joined and enrolled in Intune after MDM auto-enrollment policy is applied.

## Layout

```text
Export-IntuneDevicesCsv.cmd
Start-IntuneHybridJoinRepair-LotLauncher-GUI.cmd
Update-LotCmdWrappers.cmd
Scripts\
LOT-X\
```

- `Export-IntuneDevicesCsv.cmd` exports the global Intune inventory to `DevicesIntune.csv`.
- `Export-EntraDevicesCsv.cmd` exports the global Entra device inventory to `DevicesEntra.csv`.
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
- `Scripts\SmartM365-IntuneHybridJoinRepair-Update-LotCmdWrappers.ps1`: refreshes small LOT CMD wrappers.
- `Scripts\SmartM365-IntuneHybridJoinRepair-LotLauncher-GUI.ps1`: GUI that creates a local LOT folder from a selected computer list file, writes `Computers.txt`, refreshes wrappers, and offers to launch the LOT.

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
in `Scripts\PsExec.exe` or in `PATH`. Direct PowerShell runs can also pass `-PsExecPath`.

Force a rerun that bypasses the target run guard:

```cmd
Run-IntuneHybridJoinRepairWithPsExec-Loop-IgnoreRunGuard.cmd
```

Refresh LOT wrappers after creating or importing LOT folders:

```cmd
Update-LotCmdWrappers.cmd
```

Create a LOT folder from a computer list with the GUI:

```cmd
Start-IntuneHybridJoinRepair-LotLauncher-GUI.cmd
```

The GUI accepts a text list or a CSV. For CSV files, it uses `ComputerName`, `DeviceName`,
`Name`, or `DisplayName` when present, otherwise the first CSV column.

## Notes

- LOT folders should only contain `Computers.txt` and the small CMD wrappers.
- Reports are written under `LOT-*\Reports`.
- A live cycle CSV is written under `LOT-*\Reports` as computers complete.
- Per-computer PsExec logs are written under `LOT-*\PsExecLogs`.
- Collected remote evidence is written under `LOT-*\CentralLogs`.
- Already enrolled computers are moved from `Computers.txt` to `ComputersAlreadyEnrolled.txt`.
- Generated inventories, logs and reports are ignored by Git.

See `IMPLEMENTATION-NOTES.md` for detailed neutral implementation decisions and operational behavior.
