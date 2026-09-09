# SmartWorkplaceCMDB 0.3.0-beta.1

**Beta prerelease candidate, prepared locally. Not a stable release.**

This beta hardens local tenant isolation, trial collection and source evidence.
It does not certify production collection, tenant permissions or tenant-scale
performance. No production task or tenant has been modified during validation.

- Explicit tenant identity mismatches are rejected before CSV relabeling and
  report replacement. Declared CSV column order is respected and undeclared
  fields are omitted; the HTML report counts multiline records correctly.
- Direct fixture/bounded collectors and the orchestrator isolate test output.
  A new empty explicit root can be used as a dedicated trial root. Occupied
  unmarked roots or roots owned by another identity/mode get a unique test
  child folder. Explicit raw-file overrides cannot escape that isolated root.
- Every native source now records in-progress, completed or failed evidence
  beside its latest raw CSV, including coverage, identity, date, row count and
  SHA-256. A failed attempt retains the previous CSV, but normalizers refuse
  known failed/in-progress or changed snapshots, including in validation mode.
  Per-source locks prevent overlapping collection attempts from changing state.
- Empty/one-item fixture arrays work; missing or malformed Graph page value
  arrays fail instead of becoming a supposedly successful empty snapshot.
- Source health distinguishes complete, fixture, bounded, scoped, not collected,
  stale, unavailable evidence and changed CSVs. Empty completed sources are
  dated. Existing CSV contracts remain unchanged; legacy CSVs without evidence
  remain readable and are shown as unknown in the local report.
- Intune enrichment retains the oldest contributing Entra/Intune collection date
  and preserves an unknown date when either contributing date is absent.

## Installation and update

Extract the package into a new folder and verify its SHA-256 and public file
manifest before use. PowerShell scripts are signed; use a machine where the
WorkplaceCloudHub code-signing certificate is already trusted for `AllSigned`.
The package does not install certificates, dependencies or scheduled tasks.

Configure only local `Config/*.local.json` and tenant profiles as described in
the main README. Use PowerShell 7 and begin with `-ValidateOnly`. Fixture trials
and live prerequisites are different: fixture success does not prove cloud or
AD connectivity. A real collection requires an independently approved operator
action. SharePoint upload remains opt-in configuration and requires write access.

For updates, keep the previous application folder, back up local configuration
and runtime data, and extract side by side. Compare and copy reviewed local
settings without copying templates over them. Keep output roots explicit.
Preview any scheduled-task path change with the installer before separately
authorizing `-Execute`. No automatic update mechanism is included.

## Compatibility and limits

No CSV schema migration is required. The `.status.json`, `.lock` and
`.collection-root.json` files are private runtime data and are never included
in the distribution. Existing SharePoint publication uploads CSV files only;
JSON evidence must travel with its corresponding raw CSV to retain source
health on another host. A CSV-only copy is reported as unknown evidence.

Collection and curation across sources are not one transaction. A complete
source means completion of that collector's requested scope, not universal
inventory coverage. AD search-base/domain restrictions and disabled membership
collection are explicitly partial. Power BI Desktop refresh and real
Graph/Exchange/AD/SharePoint behavior remain unqualified. The separate
SmartWorkplaceDashboard sources are not included in this package.

See `AUDIT-BETA.md` for the offline validation evidence. External publication
requires approval of the concrete beta package, Git delta and site files.
