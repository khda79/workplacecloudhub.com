# Exchange Inventory Rules

Read this file together with the repository root `RULES.md`.

## Scope

These rules apply to Exchange Online and Exchange on-premises inventory, reporting, migration, permission, proxy address, and backup protection scripts.

## Runtime

- Import shared helpers from `Modules/SmartM365.Core\SmartM365.Core.psd1`.
- Do not use the Windows registry to resolve paths or configuration.
- Use app-only certificate authentication from local/global JSON where cloud access is required.

## Exchange Online

- Use `ExchangeOnlineManagement` for Exchange Online cmdlets.
- Use Microsoft Graph only when Graph data or SharePoint upload is needed.
- Run preflight checks for Graph probes and Exchange Online RBAC probes before main processing.
- Treat `SmartM365-Create-AppRegistration.ps1` as an interactive bootstrap outside the Exchange inventory runtime baseline.
- For read-only Exchange Online inventory/reporting scripts, target `Exchange.ManageAsApp` plus `Global Reader` on the SmartM365 service principal unless a script proves it needs write-capable Exchange RBAC.
- Do not require Exchange Administrator for Exchange Online runtime scripts just because the bootstrap needs it for setup operations.

## Exchange On-Premises

- Use Exchange Management Shell or Exchange management tools when on-premises cmdlets are required.
- Run preflight checks for Exchange on-premises readiness and AD read access where relevant.

## Configuration And Output

- Use `{{DataAllRootPath}}\Exchange\EXO\...` for Exchange Online output.
- Use `{{DataAllRootPath}}\Exchange\OnPrem\...` for Exchange on-premises output.
- Use the global `LatestCsvFolderPath` for latest CSV copies.
- Use configured retention and error mail settings.
- Upload generated CSV files through the shared SharePoint helper when enabled.
