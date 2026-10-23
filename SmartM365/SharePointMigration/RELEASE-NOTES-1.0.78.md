# Smart SharePoint Migration Toolkit 1.0.78

Release 2026-10-23. Dashboard 1.0.78; generic launcher 1.0.33.
This release packages the toolkit updates since 1.0.64.

## Dashboard and reports

- A loading shield appears before selecting a migration or opening any of the
  eight tabs. Context text identifies the migration or tab. Success and failure
  both restore the cursor and clear the shield. Local reads remain synchronous.
  The new `SmartM365-SharePointMigration-Loading.ps1` helper is required.
- Migrations can be archived from the dashboard; ArchivedOverview preserves
  the portfolio presentation. Shared activity logs retain action evidence.
- Comparison history charts, stale-input context, workbook/folder shortcuts
  and full diagnostic message copying improve result review.
- Dashboard action consoles close automatically after success; failed actions
  retain their diagnostic information.

## Source complexity and automatic assessment

- Source scans evaluate workflows, forms, customizations and custom permission
  levels automatically. Assessment sidecars and detail CSVs link to source
  inventories through SHA256 hashes and record collection and qualification.
- Thirteen criteria include library file concentration, version concentration,
  customizations and permission levels. MAX aggregation avoids double counting
  library pressure and version pressure; missing evidence remains N/A.
- Custom JSLink and XSL, site-level event receivers, malformed evidence and
  list-level failures retain explicit risk or partial qualification.
- Empty event receiver collections, standard user display forms and permission
  baseline names are handled correctly. Legacy ApplyTheme rule names remain
  supported; current templates use ApplyThemeAndBorder.
- An unavailable Workflow 2013 service or filtered permission scope does not
  establish absence. Automatic scores estimate effort and require review.

## Inventory, comparison and operational safeguards

- Destination permission reads use the CSOM context owning each object,
  including pagination, inheritance and permission preloading. Missing or mixed
  contexts cannot silently establish inheritance. Failed inventories remain
  ineligible as completed comparison inputs.
- Inventory reads reduce SharePoint round trips; comparators cache path
  normalization. Native pagination and bounded retries preserve progress.
- Source read-only lock operations verify state through administration APIs.
  Site lock and page-comment actions retain their explicit execution controls.
- Generated cleanup actions use PowerShell 7 launchers, process-only execution
  policy bypass and recycle defaults. Candidate safeguards and Home.aspx
  protection apply to the documented scope; comparison does not execute cleanup.

## Installation and validation

Download the ZIP, SHA256 checksum and file manifest. Verify the ZIP hash, then
extract the complete package and start
`SmartM365/SharePointMigration/Start-SmartM365-SharePointMigration-GUI.cmd`.
Update application files with the GUI and jobs closed; preserve private
migrations, archives, mappings, local authentication/configuration and installed
Python. Existing assessment rules remain operator-owned; review template changes.

Windows and PowerShell 7.4 or later are required for the dashboard. SharePoint
Server collection requires Windows PowerShell 5.1 on an appropriate server.
Python runtimes and third-party modules are not bundled. The public signing
certificate and optional trust installer are included; trust is an explicit
local administrative choice. See README for operational prerequisites.

Release validation covers offline Python regressions, synthetic source and CSOM
fixtures, WPF loading/navigation and demo interactions, package integrity and
PowerShell signatures. The refreshed local DemoAssets use fictional data only
and are excluded from the source package together with private configuration,
scans, logs and migration workspaces. Offline checks do not qualify real tenant
or farm permissions, authentication or complete migration coverage. No live
scan, comparison, migration or cleanup is performed to prepare this release.

Executed release checks: 59 Python unit tests; 18 PowerShell test runs, including
51 source-assessment checks on each of PowerShell 7 and Windows PowerShell 5.1,
10 CSOM checks, 39 loading checks and 375 demo checks. All 99 packaged PowerShell
sources passed syntax and Authenticode validation. All checks passed offline.
