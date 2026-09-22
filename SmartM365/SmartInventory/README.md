# Smart Inventory

Inventory scripts that can feed Power BI datasets, operational reports, CSV exports, SharePoint publishing, and other downstream consumers.

## Reliability and distribution

SmartInventory is a collection of independently versioned scripts and shared
modules; this README does not assign a suite version or stable qualification.
Shared CSV persistence, distributed orchestration, Graph collection and the
bounded AD/Exchange/readiness families have focused synthetic regression coverage.
Those results do not qualify Microsoft Graph, a tenant, on-premises systems or the
whole SmartInventory suite. Detailed audit maps, evidence and publication records
are intentionally retained outside the public repository.

Use a reviewed repository checkout with the required `Config/`, internal module
manifests and templates. A single copied collector is not a self-contained package.
No per-script ZIP release is required. Operational JSON, credentials, CSV exports
and logs must stay outside Git and public distribution.

The shared exporters reject conflicting source identities before writing
files. Existing identity-first CSV columns, names, delimiters and business fields
are preserved. A failed serialization cannot replace the previous valid CSV, but
this does not establish a transaction across all files or prove collection
completeness. Consumers must still check source dates, schemas and tenant identity.

Data quality takes precedence over runtime. A page limit reached by the Backup
Protected Mailboxes collector now fails the run before replacing the latest
CSV; use `-MaxItems` for a bounded test instead of treating `-MaxPages` as a
partial-export option. Windows Update reports the source and coverage of
`DaysSinceLastStatus`; the field remains blank when the export has no suitable
device-status date. A policy modification date is not substituted for it.

## Organization

- `ActiveDirectoryInventory/`: Active Directory inventory and reporting.
- `ExchangeInventory/`: Exchange Online and Exchange on-premises inventory.
- `M365Inventory/`: Microsoft 365, Entra, Power BI, and Microsoft Fabric inventory.
- `M365Inventory/IntuneInventory/`: Intune inventory, Windows Update reporting, Autopilot, RBAC, applications, and remediation export utilities.
- `M365Inventory/IntuneInventory/EndpointAnalytics/`: standard Endpoint Analytics score, startup, app reliability, and work-from-anywhere exports without Advanced Analytics.
- `M365Inventory/Security/`: Secure Score, authentication-method registration, and Conditional Access configuration evidence.
- `M365Inventory/IntuneInventory/Security/`: Defender agent and firewall-health reports exported from Intune without changing device or policy state.
- `ExchangeInventory/BackupProtection/`: Microsoft 365 Backup protection evidence for mailboxes, SharePoint sites, and OneDrive accounts.
- `M365Inventory/Licensing/SmartM365-LicensePricing-Inventory.ps1`: normalization of a governed pricing source. The default source is the versioned France public-list baseline `M365_License_Prices_Public-France.csv`, with every row marked estimated; a private local configuration can replace it with reviewed customer-contract prices.
- `Launchers/Cloud/`: production launchers for cloud inventory scripts.
- `Launchers/OnPremises/`: production launchers for Active Directory and Exchange on-premises inventory scripts.
- `Launchers/Orchestrator/`: production launchers for orchestrator installation, start, stop, and restart operations.

Launcher names no longer carry a `-Prod` suffix. Test launchers are intentionally not maintained; use the PowerShell script parameters directly for targeted test execution.

The Secure Score, authentication-method registration, Conditional Access, Intune endpoint-security, and governed public-France license-pricing collectors are enabled once daily in `Orchestrator/Orchestrator-Jobs.json.template`. Microsoft 365 Backup collectors remain scheduled but disabled until their service-app/controller prerequisite is available. `Policy.Read.All` and `SecurityEvents.Read.All` are new bootstrap permissions. Authentication registration and Microsoft 365 Backup reuse `AuditLog.Read.All` and `BackupRestore-Configuration.Read.All`. Intune endpoint-security reports reuse the temporary export-job permission `DeviceManagementManagedDevices.ReadWrite.All`.

## Active Directory launchers

Use the production launchers under `Launchers/OnPremises/`:

- `Start-SmartM365-ActiveDirectory-Inventory.cmd`: runs the complete Active Directory inventory for the selected tenant, including live domain collection, consolidated CSV generation, enrichment, duplicate and mail-routing analysis, reports, publishing, and weekly history.
- `Start-SmartM365-ActiveDirectory-DuplicateIdentity-Mail.cmd`: skips live Active Directory collection, reads the existing `DATA-LAST/AD_Users_AllDomains.csv`, regenerates the four duplicate and mail-routing diagnostic CSV files and the `AD_Users_IdentityAndMailRoutingIssues.xlsx` workbook, then forces the diagnostic email to be sent even if it was already sent that day.

Both launchers use tenant `prod` by default. The orchestrator can override it through `SMARTM365_ORCHESTRATOR_TENANT`. The mail-only launcher is intended for controlled reanalysis or notification retesting and does not replace a regular full inventory.

## Account classification governance

The authoritative account-type rules and workforce population mappings are stored in `Config/AccountClassification.psd1`. The Active Directory enrichment writes `AccountType`, `AccountPopulation`, and `AccountClassificationRuleVersion` to its enriched user export.

- `Human`: named and external-person accounts included in workforce KPIs.
- `Non-human`: service, shared-mailbox, room-mailbox, and system accounts excluded from workforce KPIs.
- `Review required`: admin, generic, unclassified, conflicting, or cloud-only accounts that must not be silently treated as human or non-human.

Change `RuleVersion` whenever a classification or population rule changes. Keep the three populations disjoint; validate the configuration with `Tests/Test-SmartM365-AccountClassification.ps1` before running a full inventory.
