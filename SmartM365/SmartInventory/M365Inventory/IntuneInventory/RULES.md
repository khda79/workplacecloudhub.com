# Intune Inventory Rules

Read this file together with the repository root `RULES.md`.

## Scope

These rules apply to Intune inventory and export utilities. They do not apply to Intune detection/remediation packages under `IntuneRemediation`.

## Runtime

- Use PowerShell 7 unless a script explicitly documents an exception.
- Import shared helpers from `Modules/SmartM365.Core\SmartM365.Core.psd1`.
- Use Microsoft Graph for Intune data collection.
- Prefer app-only certificate authentication from local/global JSON.

## Permissions And Preflight

- Run preflight checks before main processing.
- Probe the Graph endpoints needed by the script, for example managed devices, compliance policies, Windows Autopilot, RBAC groups, remediation exports, or report export jobs.
- Stop early when required Graph permissions are missing.

## Configuration And Output

- Use `{{DataAllRootPath}}\Intune\...` for CSV/data output.
- Use the global `LatestCsvFolderPath` for latest CSV copies.
- Use configured retention and error mail settings.
- Upload generated CSV files through the shared SharePoint helper when enabled.
