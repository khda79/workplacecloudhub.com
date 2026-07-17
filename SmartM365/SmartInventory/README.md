# Smart Inventory

Inventory scripts that can feed Power BI datasets, operational reports, CSV exports, SharePoint publishing, and other downstream consumers.

## Organization

- `ActiveDirectoryInventory/`: Active Directory inventory and reporting.
- `ExchangeInventory/`: Exchange Online and Exchange on-premises inventory.
- `M365Inventory/`: Microsoft 365 and Entra inventory.
- `M365Inventory/IntuneInventory/`: Intune inventory, Windows Update reporting, Autopilot, RBAC, applications, and remediation export utilities.
- `Launchers/Cloud/`: production launchers for cloud inventory scripts.
- `Launchers/OnPremises/`: production launchers for Active Directory and Exchange on-premises inventory scripts.
- `Launchers/Orchestrator/`: production launchers for orchestrator installation, start, stop, and restart operations.

Launcher names no longer carry a `-Prod` suffix. Test launchers are intentionally not maintained; use the PowerShell script parameters directly for targeted test execution.

## Active Directory launchers

Use the production launchers under `Launchers/OnPremises/`:

- `Start-SmartM365-ActiveDirectory-Inventory.cmd`: runs the complete Active Directory inventory for the selected tenant, including live domain collection, consolidated CSV generation, enrichment, duplicate and mail-routing analysis, reports, publishing, and weekly history.
- `Start-SmartM365-ActiveDirectory-DuplicateIdentity-Mail.cmd`: skips live Active Directory collection, reads the existing `DATA-LAST/AD_Users_AllDomains.csv`, regenerates the four duplicate and mail-routing diagnostic CSV files and the `AD_Users_IdentityAndMailRoutingIssues.xlsx` workbook, then forces the diagnostic email to be sent even if it was already sent that day.

Both launchers use tenant `prod` by default. The orchestrator can override it through `SMARTM365_ORCHESTRATOR_TENANT`. The mail-only launcher is intended for controlled reanalysis or notification retesting and does not replace a regular full inventory.
