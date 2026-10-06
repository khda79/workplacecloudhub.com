# Current-only CMDB preparation

Version: 0.3.6. Status: offline-tested migration candidate, not deployed or scheduled.

## Scope and execution boundaries

`SmartM365-CmdbEvidence-Prepare.ps1` prepares 46 flat reporting CSVs directly
from 33 canonical SmartInventory CSVs in `DATA-LAST`. It requires Python 3.10+
and uses the standard library only. No CMDB collector, compatibility Raw tree,
Intelligence prepared table, or Power BI runtime is required.

The output is the sibling `DATA-POWERBI-CMDB`, under the same configured private
data root. For this workstation, the intended output is
`C:\SmartM365\DATA\DATA-POWERBI-CMDB`. Intelligence's `DATA-POWERBI` and existing
CMDB `C:\Smart-CMDB\DATA\DATA-LAST` are not read, changed, or removed.

The PowerShell entry point selects the effective `-Tenant` profile (default
`test`). Its `.local.json.txt` template selects the Python executable and can
override `LatestCsvFolderPath`; otherwise the tenant source root is used.
Logs and transcripts use the existing Core LOG-ALL initialization/retention.
The log initialization path is resolved through Core's existing configuration
resolver before directory creation. Blank, relative or unresolved paths are
rejected instead of creating directories containing literal configuration tokens.
Core's automatic Teams log callback is suppressed in memory during this
offline invocation, and SharePoint upload is disabled. No local module or
tenant configuration is changed to disable notifications permanently.

`-ValidateOnly` validates acquisition proof and raw CSV contracts. Python does
not create prepared output, staging or a publication lock in this mode.
PowerShell still initializes its local configuration and operational logs.
Source-only validation is not full table/relationship validation.

Do not schedule the entry point or switch the report yet. Required acquisition
proof is intentionally not fabricated from existing files or their timestamps.

## Required producer completion proof

WorkplaceScope 1.0.8 retries a group membership 404 or a specifically recognized
invalid directory page token up to three complete SDK traversals (5 then 15
seconds). It restarts only that group's traversal from the first page and
discards partial pages, retaining other completed groups from the current run.
It does not reuse the rejected nextLink or retry generic HTTP 400 errors.
Retry/terminal diagnostics identify the group and attempt. A persistent token
failure, including a mixed token/404 sequence, fails the run without excluding
the group or publishing canonical CSVs. No previous-run memberships are reused.
Only after two targeted group reads return 404 (15 seconds apart) and a fresh
successful full group enumeration omits that ID may it exclude the group from
both the catalog and memberships. This corroborates absence during the run,
not the cause of deletion. Each exclusion produces a warning and its native ID
in the current receipt qualifications. It is never represented as an empty group.
A surviving/reappearing group, ambiguous result, permission error or failed
verification blocks all four canonical exports. No extra permissions, CSVs,
history or retry configuration are introduced. Normal reads make no extra calls.

The source root must contain the 17 individual current receipts declared in
`Modules/SmartM365.Core/SmartM365-CmdbSources.json.txt`. Their filename pattern is
`SmartInventory_<producer script name without .ps1>.current.json.txt`. There is
no extra aggregate source file, no scheduler change, and no receipt history.
Preparation assembles the verified set in its output `current.json.txt`.

The Python CSV reader accepts native fields up to 64 MiB by default (including
large AD membership JSON), without truncation. It restores the caller's CSV
limit after reading; schema, immutable-key, tenant, row-count and hash checks
remain mandatory. Both source validation and table generation use this reader.

SPO inventory 0.33 requires Core 1.0.67 to write UTF-8 BOM before publication.
Uploaded, weekly-snapshot and receipt bytes must agree; the other collectors'
default UTF-8 encoding is unchanged. Previously mismatched receipts require a
new producer acquisition, not manual CSV or receipt repair.

The shared helper starts a receipt before acquisition with `Status=Running`.
Completion can declare `Status=Completed` only after explicit consumer-scope
qualification, zero collection errors, and publication of every required
canonical CSV in that run. Shared receipts carry `Owner=SmartInventory-SourceReceipt`,
contract version 1.2 and `ScopeQualification=ConsumerScope`,
the four tenant identity fields, actual script/version/run, required scope,
acquisition interval, qualifications, and each file's logical row count/hash.
File lineage must match its parent producer receipt exactly. Failed scope or
missing current exports replace the receipt with a rejected result; native
CSV files are preserved. Merely finding a previous file does not qualify it.
Rejected producer proofs identify the producer and receipt filename. Incomplete
proofs also report status, partial-inventory flag and error count/type; the
completion requirements remain unchanged. Compare receipts on the collection
machine when a synchronized copy may lag behind; never manually restamp a proof.

The reader also accepts the previous CMDB-owned 1.1 format during rollout, with
its original exact-file and scope rules. New 1.2 receipts may include additional
current-run exports, but every CMDB-required file must still be present and only
those files feed preparation. `ConfiguredOutputsOnly` receipts from other
SmartInventory collectors cannot replace qualified CMDB source proof. Existing
CSV schemas and Intelligence history contracts are unchanged. A future real
producer acquisition replaces its own legacy receipt; no metadata conversion or
synthetic production proof is performed.

### Explicitly tolerated AD domain coverage

An unrestricted AD inventory may qualify with `IsPartialInventory=true` only
when every unavailable domain actually failed with a connectivity error
explicitly tolerated by `NonBlockingDomainErrors`. A shared 1.2 receipt declares
`DomainCoverage` with the expected, collected, unavailable and configured
non-blocking domain lists. Collected and unavailable domains must form an exact,
disjoint partition of the requested forest, with at least one collected domain.
Every required file repeats the same declaration. All CSVs must be published by
that current acquisition; failed-domain fragments and historical exports are
excluded. Native domain rows must match the collected-domain declaration.

The preparation returns `CoverageWarnings`, preserves the declaration in its
manifest and marks the six AD `SourceHealth` rows as `SuccessWithWarnings`, with
the unavailable domains visible. Missing domains are unknown, not zero and not a
complete forest. Non-allowlisted failures, actual collection errors, missing
required exports, hash mismatches and stale acquisitions still reject preparation.
Existing failed receipts must not be repaired manually: deploy the updated
collector and obtain a new acquisition before preparing a new snapshot.

### Licensing overview source compatibility

The 34-source contract includes `M365_Licenses_Users.csv`, required by the shared
Licensing producer registry. Its schema, tenant identity, current-run completion,
hash, logical row count and acquisition freshness are checked like the other
sources. It appears in `SourceHealth` but does not replace assignment-path or
service-plan facts and does not introduce another reporting table. This native
presentation view has no immutable row key: group display names can repeat and
one user/SKU can have multiple attribution rows. Those rows are preserved, not
deduplicated into an invented user/SKU grain. Licensing CSVs, collection behavior
and Intelligence tables/history keys are unchanged.

Each producer holds a local exclusive `.collection.lock` while collecting.
These small current lock files are not uploaded. Receipts use atomic local
replacement and the existing configured SharePoint artifact transport on
completion. No external transport is invoked by the offline test suite.

Full-scope prerequisites include unrestricted AD forest collection; local
mailboxes with `-IncludeRemoteMailboxes`, fresh overwrite and no collection
issues; all-mode fresh application relations without cache/resume reuse;
both EADevicePerformanceV2 and EADeviceScoresV2; successful device-level
readiness rather than summary-only fallback; feature and quality update
alerts; the Office365ActiveUserDetail activity report; unrestricted Teams
memberships; and all five WorkplaceScope policy families. Sites retain their
usage-report scope (OneDrive inclusion is not a claim of full site discovery).
Read-only and MAXITEMS runs do not create canonical completion evidence.
Optional workload enrichment is not silently reclassified as full evidence.

DiscoveredApps 1.30 provides the explicit non-destructive `-FreshDeviceDetails`
option for this acquisition gate. The inactive integration specification and
remaining cadence/deployment approvals are described in `CMDB-INTEGRATION.md`.
The active orchestrator manifest/template is not changed by that specification.

Every required filename must have its own file record in its producer's receipt; the example is not a
complete proof. `Rows` counts parsed logical CSV records, not physical lines.
Empty success still needs a complete header, scope and receipt. All required
sources must pass; a failed or partial producer cannot be rescued by an old CSV.

## Coherent Entra group catalog

WorkplaceScope 1.0.6 adds group display name, mail/security flags, group types,
native synchronized SID and catalog acquisition time to the existing
`M365_EntraGroupMembershipScope.csv`. The catalog and direct memberships use
one group enumeration in one producer run. No extra CSV, group enumeration,
membership call, permission or history is introduced. Original columns retain
their meaning; `CollectedAtUtc` still dates the direct membership acquisition,
while `GroupCollectedAtUtc` dates the group catalog.

CMDB group properties, AD/cloud SID links and Group 360 now use this enriched
scope rather than the separately collected licensing catalog. Catalog and member
row run IDs must agree; both acquisition dates must lie inside the WorkplaceScope
receipt interval. Runtime row IDs and receipt transaction IDs are distinct;
the existing exact file hashes bind the native CSVs to their receipt.
Membership identities, completion and exact per-group counts still fail closed.
Licensing assignment paths keep their native group IDs. An out-of-cohort
assignment remains explicitly unresolved with a quality finding, never a
fabricated empty group or a discarded license assignment.

`M365_EntraGroups_All.csv` remains unchanged and validated as independent
comparison evidence. The manifest records both catalog counts and their
directional differences in `PreparationQualifications.EntraGroupCatalogComparison`.
Differences caused by group creation/deletion between independent scans no longer
invalidate an internally coherent WorkplaceScope cohort. They do not establish
that a group is empty or that both API enumerations were simultaneous.

Contract 0.3.4 requires the enriched headers. Older WorkplaceScope exports must
be recollected; never add columns or alter producer receipts manually. The
Licensing collector, its receipt requirements, all 46 reporting table schemas,
Intelligence contracts/history keys and existing reports remain unchanged.

## Weekly application freshness

The DiscoveredApps job is already weekly (Sunday 00:00). Keep its complete
`-DeviceDetailMode All -FreshDeviceDetails` acquisition; no sampling or relation
cache reuse is introduced. The two Apps CSVs form one explicit freshness group:
acquisition age up to 168 hours is within target; strictly above 168 hours emits
warnings; strictly above 240 hours rejects the entire preparation. Age is measured
from `StartedAtUtc`, never completion, file modification, upload or preparation.
Exactly 240 hours remains acceptable, with a warning. The six-day job timeout is
not a freshness target; actual long-run duration still requires qualification.

All other sources retain the 48-hour acquisition-age and collection-span limits.
Apps are assessed within their own 240-hour span, not included in the Core span.
One producer cannot split across freshness groups. Unknown/overlapping groups,
missing sources, invalid bounds and partial/failed/running proofs fail closed.

`SourceEvidence.Freshness` records the dates, age, target, hard limit and state of
each input plus group spans and the earliest individual expiration. Warnings
are returned even by ValidateOnly, logged as WARNING by the offline wrapper,
and retained as source-level quality findings and SourceHealth evidence. They
are not a count of affected devices or proof that an unobserved device has no
applications. Current orphan-app/device joins still reject preparation.

Preparation and buffered readers share `cmdb_freshness.py`; both reject expired
evidence. The read session expires at the earliest source-specific deadline,
not the oldest timestamp plus a blanket limit. No historical fallback, new
service or permanent data copy is added. A Running Apps receipt still blocks
new preparation; the already loaded report is not refreshed or modified.

Intelligence's existing transport-age overrides are set to 240 hours for the
same two filenames in its preparation template; its other defaults, licence
price override, historyKey rules and CSV schemas are unchanged. Template merging
adds absent nested keys on the next configuration load; explicit existing local
overrides are preserved and need separate deployment review. Intelligence's
publication-age guard uses file modification time, not CMDB acquisition proof.
Its Data Trust view keeps the existing seven-day Fresh/Aging distinction; this
change does not certify transport timestamps as acquisition timestamps.

The contract version is bumped: old prepared batches cannot be loaded under the
new reader contract. Regenerate only after complete valid sources are available;
do not edit existing manifests or receipts to make old batches pass.

The 0.3.2 contract requires native AD membership identity/type columns, Entra
membership type/status, Teams member counts, and update-alert provenance/key
columns. Header absence blocks replacement; a legitimately unavailable value is
not silently converted to zero or successful evidence.

AD membership uniqueness includes the native group GUID as well as its SID,
member distinguished name and membership kind. Built-in SIDs can repeat across
domains, so distinct groups must retain their separate relationships. True
duplicate relationships still fail. Only an explicitly unresolved primary group
may omit its GUID; its native SID/RID and member identity remain required.

Autopilot grain is its native Autopilot ID, not serial number. Different IDs
sharing a serial remain separate records, including records with a blank serial.
Identical repeated IDs may be collapsed; contradictory records with one ID fail.
The producer publishes a complete header for successful zero acquisition. AD
domain and combined empty exports preserve their actual native column order,
and empty user/computer enrichment adds the same calculated columns as nonempty
exports. Remote mailbox zero exports retain the full native schema; query
failures and acquisition warnings cannot qualify them. These changes do not
prove a production zero population: the producer receipt is still required.

Intelligence's history keys, prepared contracts and generators are unchanged by
this correction lot. Current Autopilot counts can increase when previously lost
native identities are restored; qualify consumer counts and joins before deployment.

Core rejects acquisition older than 48 hours or Core intervals spanning over
48 hours; only the Apps group follows the weekly policy above. All sources reject
future acquisition over five minutes. Explicit timezones are
mandatory. These acquisition gates do not prove that a workload usage report
itself is fresh: its source report refresh date remains separate evidence.

Missing receipts, wrong hashes/counts/tenant, malformed CSVs, missing required
columns, duplicate native keys and missing immutable identities block promotion.
Producer integration qualifies declared source scope, not tenant business
health. Receipt/version/hash validation does not establish real API coverage;
production qualification on the collection host remains required.

## Publication and preservation

The preparation holds an exclusive local publication lock and writes its 46 CSVs
and `current.json.txt` into a unique transient staging directory. The manifest
records reporting identity, source acquisition intervals and receipts, source and
output hashes/row counts, contract hash/version and metric definitions.

Before and after local promotion, it checks the output schemas, unique keys,
reporting identity, relationships and hashes, and rechecks source hashes.
If validation or replacement fails, the previous owned snapshot is restored.
An unrelated output directory is never adopted or deleted. Unknown interrupted
staging/rollback artifacts require explicit recovery; they are not automatically
purged. A crash or locked-file/permission failure can require manual recovery.

The previous snapshot is held only for rollback during the replacement and
removed after success. There is no dated CMDB history, source copy, or persistent
adapter tree. If cleanup fails after a successful validated replacement, the
result is `PreparedWithCleanupWarning`; the new valid output is not undone.

Directory replacement and the lock protect the local process. They are not an
atomic cross-machine OneDrive/SharePoint transaction. The future report loader
must verify the manifest and every file after synchronization and reject mixed
batches before refreshing.

## Business semantics and remaining qualification

- Device identity uses native Entra/Intune correlation IDs. Same-name devices
  and all platforms remain separate. All Intune source candidates are retained;
  selection uses qualified sync/enrollment time, then the native ID.
- AD workstation coverage uses unique native SID to Entra device ID to Intune
  device ID. Missing/ambiguous SID matches remain explicit evidence gaps, not
  proof of an AD-only device. Unproven hybrid only-in-one-directory counts are
  blank, not inferred from unmatched rows. AD-only machines remain in the
  separate AD coverage fact; they are not silently added to the cloud dimension.
- Product identity is normalized name/publisher/platform, excluding version.
  `TopApplication.ReportedDeviceCount` is distinct managed-device IDs across
  versions. The old occurrence metric name/description in Power BI needs
  adaptation before a report switch. All-mode app/device relations and exact
  per-app count agreement are required; Top/None is rejected.
- Native mailbox identity is preserved per source, then normalized nonblank
  SMTP reconciles hosting. Online/Remote wins over local evidence. Conflicting
  SMTP identities within one source block output. Blank SMTP stays evidence,
  not a fabricated address. Hosting and recipient type are different metrics.
- Missing capacity stays blank. Compact plan states are decoded explicitly;
  unknown state codes are rejected. Assignment paths retain native states and
  direct/group routes. No license non-use conclusion follows from absent dates.
- User activity retains the existing disabled / observed sign-in age categories;
  unqualified dates remain invalid evidence. Sites and Teams retain the 90-day
  categories, anchored to acquisition rather than preparation time.
- Sites retain usage-report coverage; this is not complete tenant site discovery.
  Team child enumeration is checked against each team's member count. Unmatched
  guest/external users remain membership records, with unresolved links.
- Hardware uses `Intune_DeviceHardware_All.csv`, produced by the existing device
  collector with explicit per-device hardware GETs (batches of at most 20, bounded
  retry and SDK fallback). Every managed-device identity requires a detailed
  record; failed acquisition rejects CMDB full scope. Serial/manufacturer/model,
  storage, RAM and the actual per-device acquisition time are separate evidence.
  Missing and reported zero remain distinct unknown states. RAM is reused in the
  existing Windows CSV; existing Intelligence CSV columns remain unchanged.
  The additional all-platform GET volume must be qualified on the collection host.
  No beta hardware endpoint, CPU/BIOS invention or list-default capacity is used.
- Configured policy target counts are
  not effective device policy assignments; nested policy payloads remain in the
  hashed native source, not a full settings report.
- The existing 39 report table names remain, with additive hardware columns and
  seven new detail tables: ADUserSource, ADComputerSource, ADGroupSource,
  ADDomainSource, ADDirectoryObject, ADMembership and EntraGroupMembership.
  Native SID uniqueness on both AD and cloud sides is mandatory for links.
  AD user UPN is optional. AD manager/managed-by DNs resolve only to unique native
  directory objects; this is not a collected Entra manager/owner relationship.
  AD groups also link to Group 360 only through a unique AD/cloud native SID;
  duplicate AD SIDs do not produce two cloud links to the same entity.
  AD replicated logon FILETIME is qualified as approximate logon evidence;
  ambiguous locale timestamps remain raw and unqualified. Direct AD members
  reconcile to MembersJson; primary groups are separate SID/RID evidence.
  Unresolved/external members and non-user Entra objects remain records.
- Quality findings cover unknown/derived country, missing/orphan primary users,
  unqualified device activity, hardware gaps/repeated serials, native licensing
  errors, unresolved group members and mailbox-link gaps. Each native license
  path can retain its own finding. User/device/group links support 360 inspection.
  Derived country and technical DiscoveryMailbox gaps are Information, not
  actionable warnings. Noncompliance or account disability is business health,
  not a collection-integrity defect. Missing identities, duplicate immutable keys,
  stale/partial receipts and broken required joins reject preparation; they are
  not promoted as a supposedly healthy report with a few warnings.
- The current calendar covers the preparation year. Temporal model behavior,
  real source equivalence, joins, quality finding completeness and production volume
  still require qualification before the existing CMDB report can be replaced.
- Native tuple-derived keys differ from legacy CMDB keys. Bookmarks, joins,
  slicers, measures and all ten report pages need separate model/Desktop tests.

## Offline verification

Run the synthetic test suite from the repository using:

```powershell
python -B .\SmartM365\Tests\test_cmdb_preparation.py
pwsh -NoProfile -File .\SmartM365\Tests\Test-SmartM365CmdbProducerReceipts.ps1
pwsh -NoProfile -File .\SmartM365\Tests\Test-SmartM365CmdbSourceCompleteness.ps1
pwsh -NoProfile -File .\SmartM365\Tests\Test-SmartM365CmdbNativeEmptyExports.ps1
pwsh -NoProfile -File .\SmartM365\Tests\Test-SmartM365DeviceHardwareOffline.ps1
powershell.exe -NoProfile -File .\SmartM365\Tests\Test-SmartM365CmdbProducerReceipts.ps1
powershell.exe -NoProfile -File .\SmartM365\Tests\Test-SmartM365CmdbNativeEmptyExports.ps1
```

It creates synthetic fixtures in the system temporary folder and removes them
after each test. It does not read synchronized tenant CSVs, call APIs, execute
collectors, switch reports or publish anything. Synthetic success does not
establish production data equivalence or Graph permission completeness.

The source and output inventories are defined in
`cmdb-prepared-contract.json.txt`; do not relax its checks merely to get an
existing incomplete export through preparation.
