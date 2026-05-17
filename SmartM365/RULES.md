# SmartM365 Rules

These rules apply to the whole repository. Folder-specific `RULES.md` files add constraints for their own area and must be read together with this file.

## Scope

These rules apply to scripts, modules, JSON configuration files, and documentation.

## Configuration

- Do not commit customer-specific or environment-specific values.
- `*.local.json` files are ignored by Git and are the only place for local operational values.
- `SmartM365.global.local.json` contains shared local values for the workstation or tenant.
- Keep a safe, committed `SmartM365.global.local.json.template` file at the repository root so users know which local values must be created.
- Scripts must fail with an explicit configuration error when required local/global values are missing and no safe default exists.
- Per-script `*.local.json` files must keep the same global keys for readability, but the value must be `__USE_GLOBAL__` when the script should inherit from `SmartM365.global.local.json`.
- A real value in a per-script `*.local.json` is an explicit script-level override.
- Configuration helpers must treat `__USE_GLOBAL__`, `USE_GLOBAL`, and an empty string as fallback markers.
- Configuration helpers must resolve tokenized values such as `{{WorkspaceRootPath}}`, `{{DataAllRootPath}}`, `{{LatestCsvFolderPath}}`, and `{{LogAllRootPath}}`.

## Path Layout

Global path roots are defined in `SmartM365.global.local.json`:

- `WorkspaceRootPath`
- `DataAllRootPath`: `{{WorkspaceRootPath}}\SMART-M365\DATA-ALL`
- `LatestCsvFolderPath`: `{{WorkspaceRootPath}}\SMART-M365\DATA-LAST`
- `LogAllRootPath`: `{{WorkspaceRootPath}}\SMART-M365\LOG-ALL`

Per-script CSV/data paths should use `{{DataAllRootPath}}` plus a functional path, for example:

- `{{DataAllRootPath}}\ActiveDirectory\Inventory`
- `{{DataAllRootPath}}\Exchange\EXO\Mailboxes`
- `{{DataAllRootPath}}\Exchange\OnPrem\Mailboxes`
- `{{DataAllRootPath}}\Intune\Devices\Inventory`
- `{{DataAllRootPath}}\Intune\WindowsUpdate\Status`
- `{{DataAllRootPath}}\M365\Users\ActiveUsers`

`LatestCsvFolderPath` is global and should remain a single shared `DATA-LAST` folder unless a script has a documented exception.

`LogAllRootPath` is reserved for a future split between data exports and centralized logs. Do not move log output there unless the script and module behavior are updated consistently.

## CSV Naming

Generated CSV file names outside `IntuneRemediation` must be readable from the shared `DATA-LAST` folder without relying on their source directory.

Use this pattern:

```text
<Domain>_<SourceOrPlatform>_<Object>[_Scope][_View].csv
```

Examples:

- `Exchange_EXO_Mailboxes_AllDomains.csv`
- `Exchange_EXO_AcceptedDomains.csv`
- `Exchange_OnPrem_Mailboxes_AllDomains.csv`
- `Intune_Devices_Inventory.csv`
- `M365_Entra_Devices.csv`
- `AD_Users_AllDomains.csv`

Keep timestamps only for historical copies, as the final suffix before `.csv`, for example `Exchange_EXO_Mailboxes_AllDomains_20260517_013000.csv`. Latest copies must stay non-timestamped.

Do not use misleading object prefixes such as `Mailboxes_` for non-mailbox data, version suffixes such as `_1.0`, or mixed separators in new CSV names.

## Script Rules

- Do not use the Windows registry to resolve script paths.
- Runtime modules live under `Modules/SmartM365.Core` and `Modules/SmartM365.SharePoint`; do not reintroduce `Modules.PS5` or `Modules.PS7` root folders.
- Windows PowerShell 5.1 compatibility helpers are grouped in `Modules/SmartM365.Core/Compatibility/WindowsPowerShell5/SmartM365-WindowsPowerShell5.psd1`; do not split them back into separate compatibility modules.
- Intune detection/remediation scripts follow their own rules in `IntuneRemediation/RULES.md`; do not add JSON config dependencies to them.
- Keep `RetentionMaxCSV` and `RetentionMaxLogs` centralized through config, with default fallback values of `30`.
- Error email recipients must resolve from `ErrorMailTo`; per-script configs should use `__USE_GLOBAL__` unless a script-specific recipient is required.
- Keep the permissions recap in `README.md` current when scripts add or change Graph scopes, Intune permissions, Exchange RBAC needs, AD access needs, or SharePoint upload behavior.
- Prefer shared preflight checks before main processing: required modules, output path write access, Graph permissions, Exchange Online RBAC, Exchange on-prem readiness, and AD read access.
- Inventory/report/export scripts should upload generated CSV files through the shared SharePoint helper when `EnableSharePointUpload` is enabled.
- Keep all PowerShell script filenames prefixed with `SmartM365-`, including Intune detection/remediation package scripts.
- Do not include version numbers in active script filenames. Keep the active version in the synopsis/header and rely on Git history for changes.
- Keep script comments concise and useful. Remove stale customer-specific comments or dead notes.

## Git Hygiene

- Keep `*.local.json` in `.gitignore`.
- Do not stage generated logs, temporary files, archives, or local exports.
- Do not keep archive folders in the repository; preserve history through Git instead.
