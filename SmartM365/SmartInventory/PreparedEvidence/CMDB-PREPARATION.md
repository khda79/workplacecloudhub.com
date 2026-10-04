# Current-only CMDB preparation

Version: 0.3.3. Status: offline-tested migration candidate, not deployed or scheduled.

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
Completion can declare `Status=Completed` only after explicit full-scope
qualification, zero collection errors, and publication of every required
canonical CSV in that run. Each receipt carries owner, contract version 1.1,
the four tenant identity fields, actual script/version/run, required scope,
acquisition interval, qualifications, and each file's logical row count/hash.
File lineage must match its parent producer receipt exactly. Failed scope or
missing current exports replace the receipt with a rejected result; native
CSV files are preserved. Merely finding a previous file does not qualify it.
Rejected producer proofs identify the producer and receipt filename. Incomplete
proofs also report status, partial-inventory flag and error count/type; the
completion requirements remain unchanged. Compare receipts on the collection
machine when a synchronized copy may lag behind; never manually restamp a proof.

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

Every required filename must have its own actual receipt; the example is not a
complete proof. `Rows` counts parsed logical CSV records, not physical lines.
Empty success still needs a complete header, scope and receipt. All required
sources must pass; a failed or partial producer cannot be rescued by an old CSV.

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

The initial contract rejects acquisition older than 48 hours, intervals spanning
over 48 hours, and future acquisition over five minutes. Explicit timezones are
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
