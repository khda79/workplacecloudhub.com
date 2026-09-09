# Smart Workplace Dashboard — BETA 3.8.0-beta.1

An autonomous Power BI source project for Microsoft 365 and Workplace operations. It reads existing SmartInventory CSV exports directly. This beta is prepared for controlled evaluation; offline validation is not real-environment qualification.

## Install and start

Extract the entire package, retaining its folder structure. Open `pbip/SmartWorkplaceDashboard.pbip` in Power BI Desktop for Windows with PBIP/PBIR support enabled. The model uses compatibility level 1600. Microsoft still documents PBIP/PBIR preview limitations: [PBIP documentation](https://learn.microsoft.com/en-us/power-bi/developer/projects/projects-overview).

The distributed PBIP has no cached data. A refresh needs all 90 selected CSVs. It does not run collectors, install modules or connect to Graph/Exchange itself. No installer, automatic updater, PBIT or PBIX is supplied.

Configure these Power Query parameters before refresh:

| Parameter | Required value / default |
| --- | --- |
| SourceMode | LocalFolder (default) or SharePoint |
| DataRootPath | A single tenant root containing DATA-LAST and DATA-ALL; generic default C:\SmartM365\DATA |
| ExpectedTenantKey | Required; blank in the public package, deliberately blocking refresh until configured |
| SharePointSiteUrl | HTTPS site root for SharePoint mode |
| SharePointDataFolderUrl | Folder URL below that site, containing DATA-LAST and DATA-ALL |
| HistoryMonths | Integer 0–120, default 24 |
| HistoryMinRowRatio | 0–1, default 0.5 |

One model instance is for one tenant. Current files are read only from the selected DATA-LAST folder, without traversing tenant subfolders. Rows in CSVs that expose TenantKey must match ExpectedTenantKey. CSVs without that column depend on the operator selecting one coherent source root: provenance cannot be inferred from their rows. Never pool tenant exports. There is no RLS or authentication layer in this beta; Country slicers are not access controls. Restrict access to the PBIP/PBIX, source folders and any exported reports.

## Source contracts and interpretation

The versioned allowlist is `source-selection.json`; all 90 pinned headers and delimiters are in `source-schema.json`. Missing files, header/schema mismatches, duplicate current files, mismatching tenant rows, explicit/unknown partial-inventory flags, conflicting identities and ambiguous keys stop refresh with a named SmartInventory error. Exact duplicate rows may collapse at validated grains. Empty files without headers fail; header-only CSVs are valid empty tables, not evidence of successful collection.

Every source column is retained, plus six technical snapshot columns. Decimal comma and dot are supported; blank/invalid/nonfinite numbers and unknown booleans remain null; fractional counts are not rounded. Entity counts omit blank IDs. Disabled/noncompliant DAX comparisons use strict equality so unknown states are not false. E3/E5/F3 quantities and availability remain blank when required license evidence is absent or invalid. Missing sign-in history does not prove inactivity.

Current entities reconcile through exact identifiers. Base64 immutable IDs remain case-sensitive. Conflicting identifiers, ambiguous matches or missing canonical entity IDs block refresh for correction; there is no silent choice of one match. Device/user names and proxy-address lists are not join keys.

Freshness uses the oldest recognized report/export timestamp across rows. If a recognized timestamp field is partly missing/invalid, freshness is unknown. Without such a field, file modification time is a proxy and can be changed by file copying. Future timestamps stop source refresh; future entity activity is classified as unknown. The four age cards show users, devices, licenses and Exchange source ages; they are not a complete freshness audit of 90 tables and do not implement a universal freshness SLA.

Weekly history retains the newest file per prior week. Tied newest timestamps fail. Snapshots below the ceiling of current row count times HistoryMinRowRatio are omitted. This heuristic can omit legitimate smaller historical estates and cannot detect every partial collection. Snapshot series use file dates; they are observations, not an event history.

## Reports, filters and exports

The beta includes 90 visible source tables, two derived tables (DeviceDetail, UserDetail), 189 measures, 16 relationships and 19 report pages. All measures reside on M365_Users_Active. See `model-and-pages.md` for the exact pages.

Country filters propagate only through existing relationships. DeviceDetail and UserDetail use their own country fields. Tenant licensing/storage, unlinked AD daily history, server and other unlinked datasets remain global. This is a filter-coverage boundary, not tenant isolation.

Service-plan states measure assigned rights/provisioning, not application usage. SharePoint capacity is a licensing estimate with a base allocation, eligible seats and add-ons; Project/Visio estimates are separate. Review current contractual entitlements before relying on it. Exchange storage utilization uses aggregate mailbox quotas, not a tenant-wide pool. Hardware refresh/BIOS age are operational proxies, not purchase/manufacture dates. Migration failed items are deduplicated per batch; backup scope uses MailboxId/MemberId reconciliation.

No custom export command is included. Power BI CSV/Excel/PDF exports and their completeness/formatting require Desktop/service validation and may expose identifiers. Do not redistribute refreshed workbooks or exports as the public beta package.

## Build and offline validation

Node.js and PowerShell 7 for Windows are development prerequisites; tested here with Node 24.17.0 and PowerShell 7.6.5. No npm packages are required to build. From this application folder:

```powershell
node .\scripts\Build-SmartWorkplaceDashboard.js
pwsh -NoProfile -File .\scripts\Validate-SmartWorkplaceDashboard.ps1
node .\scripts\Test-SmartWorkplaceDashboard.cjs
```

The default build and validator use only public source contracts. Setting SMART_M365_DATA_ROOT explicitly opts into CSV-header inspection for a build; `-CheckSourceFiles` opts into the validator's header inspection. These modes do not verify runtime data quality. Do not point them at real data when performing a synthetic-only audit.

The public logo is included under assets. The PowerShell validator is Authenticode signed. The package manifest identifies the signer and timestamp status; signature trust depends on the receiving machine. Verify the ZIP SHA-256 before extracting. JavaScript, M, JSON and Markdown do not use PowerShell Authenticode signing.

Read `RELEASE-NOTES.md`, `KNOWN-LIMITATIONS.md` and `VALIDATION.md` before evaluation. The next required product qualification is a separately authorized Desktop refresh and visual/export review on an isolated test environment. No real tenant qualification is claimed.
