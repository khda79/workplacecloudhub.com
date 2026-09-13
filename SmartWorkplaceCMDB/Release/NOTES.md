# SmartWorkplaceCMDB 1.1.3

**Stable release. Git publication is a separate reviewed action; tenant data,
local configuration and private report artifacts are never part of this package.**

## 1.1.3 stable patch candidate — 2026-09-12

- Adds one native PowerShell transcript for every executed collector/build/report
  step in its existing `LOG-ALL\Jobs\<script-name>` directory.
- Records the exact transcript path next to the structured log path in the run
  CSV and exposes the shared job root in the orchestration result.
- Applies `StepLogRetentionDays` and `MaxStepLogsPerScript` independently to
  `.log` and `.transcript.txt` files, preserving the same age/count policy for
  both without halving the retained number of executions.
- Keeps `-ValidateOnly` strictly read-only: no transcript is started.

## 1.1.2 stable patch candidate — 2026-09-12

- Adds WorkplaceCloudHub startup and completion banners to every orchestrator
  invocation, including failed preflight validation.
- Prefixes every operational console line with local date and time, including
  child collector messages and step completion messages.
- Preserves read-only validation: `-ValidateOnly` displays the lifecycle and
  timestamps but creates no data, transcript, text log, or run CSV.
- Writes the same unprefixed lifecycle banners to collection/fixture
  orchestrator logs while retaining timestamp and severity on operational log
  lines and dedicated per-step logs.
- Fixes live Active Directory membership collection by explicitly capturing
  the LDAP connection and ranged-member helpers before the per-domain closure;
  both helpers now remain callable inside the closure's dynamic module.
- Routes unresolved cross-domain group members to the domain controller whose
  naming context matches their distinguished name instead of querying the
  group domain controller.

## 1.1.1 stable feature candidate — 2026-09-12

- Adds a timestamped orchestrator log, a structured run CSV and one dedicated
  log per executed pipeline step under the tenant `LOG-ALL` root.
- Uses separate `Orchestration\Logs`, `Orchestration\Runs` and
  `Jobs\<script-name>` directories so concurrent hosts and repeated runs do not
  overwrite one another.
- Applies the SmartInventory retention pattern with configurable age limits of
  30 days for orchestrator/step logs and 90 days for run CSVs. Count guards
  additionally retain at most 30 logs per orchestrator/script and 90 run CSVs.
- Prefixes every physical text-log line with local date/time and severity, and
  records the exact step-log path in the structured execution CSV.

## 1.1.0 stable feature candidate — 2026-09-12

- Applies the common paged Microsoft Graph request and bounded transient retry
  pattern to Entra users, completing the alignment of every Graph-native
  collector.
- Stages and contract-validates every single-table cloud source before
  transactionally promoting history and `DATA-LAST`; a failed promotion
  restores the previous completed source evidence.
- Sends an optional aggregate summary after a successful live, unbounded,
  unscoped `Full` collection, using direct Graph REST app-only mail or SMTP
  without a SmartM365 runtime dependency.
- Reports devices, users, mailboxes, and distinct Microsoft 365 F1, F3, E3, E5
  and Copilot assignments with current capacity/utilization and deltas since
  the immediately previous full collection, J-7 and J-30.
- Keeps F1 (`M365_F1`) distinct from F3 (`SPE_F1`), retains private aggregate
  snapshot history, retries transient mail/token failures, and never includes
  entity identifiers in the generated HTML.
- Skips notifications for validation, fixtures, bounded runs, scoped AD runs
  and non-Full pipelines. A notification failure preserves collected data and
  returns `CompletedWithWarnings`.

This feature is covered by offline synthetic tests. No live tenant collection,
mail delivery, SharePoint upload, or domain-reachable AD run was performed as
part of this code validation.

## 1.0.2 stable patch candidate — 2026-09-12

- Collects each discovered Active Directory domain as an explicit completion
  unit and retries only transient connectivity, ADWS, LDAP timeout, server-busy,
  and invalid-enumeration-context failures.
- Uses three configured retries with 5, 15, and 30 second delays and forces
  domain-controller rediscovery before a retry. Authorization, schema, and data
  errors fail immediately.
- Stages and validates all Active Directory CSVs before transactional promotion;
  a failed attempt restores the previous complete CSVs and source evidence.
- Adds diagnostic `TargetDomains` scoping. Scoped AD runs are clearly evidenced
  and cannot trigger SharePoint publication.
- Adds normalized organizational-unit inventory plus limited user, computer,
  and group attributes required for workplace analysis.
- Keeps domain parallelism disabled (`DomainParallelThrottleLimit=1`) until a
  live domain-reachable qualification demonstrates that it is safe.

This patch is covered by offline synthetic tests only. A complete live Active
Directory collection remains required on the domain-reachable host before live
qualification can be claimed.

## 1.0.1 stable patch candidate — 2026-09-12

- Pages every bulk Active Directory user, group and computer query explicitly,
  without imposing a result-set limit on complete runs.
- Replaces the AD Web Services `Get-ADGroupMember` bulk action with LDAP
  `member;range=` retrieval in safe 1,000-value windows.
- Preserves direct membership semantics and supplements explicit group members
  with user and computer primary-group membership.
- Resolves security principals across the complete forest inventory before
  querying individual unresolved objects, and rejects malformed or stalled LDAP
  range responses.
- Adds synthetic coverage for a group containing more than 5,000 direct members.

The patch avoids changing domain-controller ADWS limits. Runtime tenant names,
group identities, paths and collected rows remain private and are not included
in the release package.

## 1.0.0 stable candidate — 2026-09-11

- Promotes application, supported modules, collectors, contracts and report
  generators to the stable `1.0.0` release channel.
- Adds bounded retry handling for transient Microsoft Graph responses, including
  capped `Retry-After` handling.
- Adds Intune encryption state to the raw, curated and Power BI device contracts.
- Integrates the existing list-only Intune hardware collector and CI adapter
  into validation, orchestration and the package allowlist.
- Packages the common CI registry, source lineage, optional declared governance
  and organization-context contracts without inventing ownership or location.
- Adds deterministic data-quality findings for missing identity, primary user,
  country and observed license-assignment errors.
- Versions the Power BI generators and tests and consolidates the canonical
  report into 10 decision-oriented pages under the stable `SmartWorkplaceCMDB` name.
- Keeps Autopilot, Intune applications/policies, cloud membership expansion,
  effective service-plan modeling, Azure resource inventory, hybrid identity
  reconciliation, advanced lifecycle, Business Services and recursive impact
  analysis in the V2 backlog.

The release package contains code, schemas, documentation and synthetic test
fixtures only. It excludes tenant configuration, real CSVs, report caches,
PBIP/PBIX artifacts and identifiers. Live connector execution and production
qualification remain separate operational activities.

## Historical 0.3.0-beta.1 preparation

## Local correction after 0.3.0-beta.1 (BETA, not packaged or published)

The local 360 report fix restores detail rows when a selector filters an
entity label instead of its technical key. It keeps details blank without
a single selected entity, uses explicit English display captions, and
increases dropdown body space. Model query validation and Desktop visual
validation remain separate checks; no collector or source data is changed.

The optional local `--include-360` report increment adds Device 360, User 360
and Group 360 using existing raw and CMDB exports. It retains individual
license assignment paths, enrollment/source evidence and exact entity-finding
links while preserving summary counts and the original 23 measure formulas.
The nine-page model has 15 tables, 13 single-direction relationships and 39
measures. Ambiguous date text and sentinel timestamps remain explicitly
unqualified. This increment changes neither collectors nor source contracts,
and does not collect membership, ownership, hardware or activity information
that was not present in the inputs. These files are local BETA work and have
not been added to the distributed package or published. See PowerBI/README.md
for preparation, testing and the required Desktop import/navigation check.

The local data-quality normalizer now retains an unlinked `DiscoveryMailbox`
as `Information / TechnicalMailboxWithoutUser` only when exactly one matching
CMDB row and its Power BI fact agree on that type and the external directory
identifier is empty. Ordinary, ambiguous and externally identified unlinked
mailboxes remain warnings. The finding key is preserved; no tenant object or
user association is changed. The suite covers 16 synthetic scenarios.

A local native Power BI report builder prepares six English pages: overview,
devices, users, licenses, mailboxes and quality/source coverage. It copies
authorized local data to a new private project, adds 23 documented measures
and validates keys, types and tenant identity. It does not overwrite a PBIX,
collect tenant data or publish anything. Report labels, measure names and
descriptions, generated status labels and operator instructions are English;
the model culture is en-US. Source names and business data remain unchanged.
The local readability update uses supported full-number formatting for cards
and charts, explicit textbox padding and larger text, wrapping table content,
and display-only labels for missing values. Source values and all 23 measure
formulas remain unchanged; the relationship-generated blank category is
excluded only from charts and slicers whose physical rows have display labels.
Opening, refreshing and visually
validating the new report in Desktop are separate operator steps. See
`PowerBI/README.md`. These changes remain BETA and are not packaged or published.

Source-health freshness now preserves offsets and fractional seconds when
PowerShell materializes JSON timestamps as DateTime values. Equivalent `Z`,
`+00:00` and `+02:00` instants retain the same age, including the exact 48/168-hour
thresholds and five-minute future tolerance. The shared JSON reader and CSV
contracts are unchanged. This working-tree correction is not present in the
previously published 0.3.0-beta.1 archive.

The source-evidence suite covers both en-US and fr-FR cultures and one-tick
boundaries on Windows PowerShell 5.1 and PowerShell 7. These offline checks do
not qualify cloud access, AD, tenant-wide completeness or a stable release.

## 0.3.0-beta.1 baseline

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
Graph/Exchange/AD/SharePoint behavior remain environment-specific qualification
activities. The separate SmartWorkplaceDashboard sources are not included in
this package.

See `AUDIT-BETA.md` for the offline validation evidence. External publication
requires approval of the concrete beta package, Git delta and site files.
