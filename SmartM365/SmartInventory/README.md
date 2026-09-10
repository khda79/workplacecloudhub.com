# Smart Inventory

Inventory scripts that can feed Power BI datasets, operational reports, CSV exports, SharePoint publishing, and other downstream consumers.

## Audit status and distribution

SmartInventory is a collection of independently versioned scripts and shared modules;
this README does not assign a suite version or stable qualification. The current
source audit starts with shared CSV persistence and distributed orchestrator
persistence. Read [AUDIT-LOT1.md](AUDIT-LOT1.md) for the demonstrated corrections,
synthetic tests, prerequisites and remaining limits, and
[INVENTORY-MAP.md](INVENTORY-MAP.md) for collector, job and consumer mappings.

[AUDIT-LOT2.md](AUDIT-LOT2.md) covers the orchestrator timeout
supervision and recovery corrections. Its synthetic validation is limited to that
lifecycle scope and does not qualify all collectors or a live deployment.

[AUDIT-LOT3.md](AUDIT-LOT3.md) covers the re-adoption and resident
lock corrections, with synthetic evidence and explicit deployment limits.

Use a reviewed repository checkout with the required `Config/`, internal module
manifests and templates. A single copied collector is not a self-contained package.
No per-script ZIP release is required. Operational JSON, credentials, CSV exports
and logs must stay outside Git and public distribution.

The shared exporters reject conflicting source identities before writing
files. Existing identity-first CSV columns, names, delimiters and business fields
are preserved. A failed serialization cannot replace the previous valid CSV, but
this does not establish a transaction across all files or prove collection
completeness. Consumers must still check source dates, schemas and tenant identity.

## Organization

- `ActiveDirectoryInventory/`: Active Directory inventory and reporting.
- `ExchangeInventory/`: Exchange Online and Exchange on-premises inventory.
- `M365Inventory/`: Microsoft 365, Entra, Power BI, and Microsoft Fabric inventory.
- `M365Inventory/IntuneInventory/`: Intune inventory, Windows Update reporting, Autopilot, RBAC, applications, and remediation export utilities.
- `M365Inventory/IntuneInventory/EndpointAnalytics/`: standard Endpoint Analytics score, startup, app reliability, and work-from-anywhere exports without Advanced Analytics.
- `Launchers/Cloud/`: production launchers for cloud inventory scripts.
- `Launchers/OnPremises/`: production launchers for Active Directory and Exchange on-premises inventory scripts.
- `Launchers/Orchestrator/`: production launchers for orchestrator installation, start, stop, and restart operations.

Launcher names no longer carry a `-Prod` suffix. Test launchers are intentionally not maintained; use the PowerShell script parameters directly for targeted test execution.

## Active Directory launchers

Use the production launchers under `Launchers/OnPremises/`:

- `Start-SmartM365-ActiveDirectory-Inventory.cmd`: runs the complete Active Directory inventory for the selected tenant, including live domain collection, consolidated CSV generation, enrichment, duplicate and mail-routing analysis, reports, publishing, and weekly history.
- `Start-SmartM365-ActiveDirectory-DuplicateIdentity-Mail.cmd`: skips live Active Directory collection, reads the existing `DATA-LAST/AD_Users_AllDomains.csv`, regenerates the four duplicate and mail-routing diagnostic CSV files and the `AD_Users_IdentityAndMailRoutingIssues.xlsx` workbook, then forces the diagnostic email to be sent even if it was already sent that day.

Both launchers use tenant `prod` by default. The orchestrator can override it through `SMARTM365_ORCHESTRATOR_TENANT`. The mail-only launcher is intended for controlled reanalysis or notification retesting and does not replace a regular full inventory.
