# Active Directory Inventory Rules

Read this file together with the repository root `RULES.md`.

## Scope

These rules apply to Active Directory inventory and reporting scripts in this folder.

## Runtime

- Use PowerShell 7 unless a script explicitly documents a compatibility exception.
- Import shared helpers from `Modules/SmartM365.Core\SmartM365.Core.psd1`.
- Use the `ActiveDirectory` module for AD reads.
- Do not use the Windows registry to resolve paths or configuration.

## Configuration And Output

- Use per-script `*.local.json` files for script-specific paths.
- Use `{{DataAllRootPath}}\ActiveDirectory\...` for CSV/data output.
- Use the global `LatestCsvFolderPath` for latest CSV copies.
- Use `RetentionMaxCSV` and `RetentionMaxLogs` from configuration, with fallback `30`.
- Use `ErrorMailTo` from configuration for error notifications.
- Upload generated CSV files through `Invoke-SmartM365SharePointCsvUpload` when SharePoint upload is enabled.

## Preflight

- Run `Invoke-SmartM365Preflight` before main processing.
- Include output path write checks.
- Include required module checks for `ActiveDirectory`.
- Include `-RequireActiveDirectoryRead` so the script stops early when the account cannot read AD.
