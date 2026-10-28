# Prepared evidence: deployment and operation

Workplace Intelligence prepares one governed batch in `DATA-POWERBI`, including
native CMDB evidence. The old `CmdbEvidence-Prepare` job is removed from the
template. Existing evidence in `DATA-POWERBI-CMDB` is not deleted by the job cutover.

## Targeted CMDB cutover

Deploy the reviewed code first. Import `Orchestrator/SmartM365.WorkplaceCutover.psm1`
and call `Get-SmartM365WorkplaceCutoverPlan -SharedDataFolderPath <orchestrator-root>`.
This preview reads the existing jobs/cluster bytes without migrating files, writing
configuration or submitting jobs. Inspect `Changes`, `JobsHash` and `ClusterHash`.

The plan removes only the `CmdbEvidence-Prepare` job and merges its dependency union into
`WorkplaceEvidence-Prepare`, including WorkplaceScope, VerifiedDomains,
WindowsAutopilot and AutopatchAlerts. It preserves the effective Workplace schedule,
arguments, freshness floor, server assignments, all collectors and cluster settings.
It does not copy the template schedule over the effective schedule.
An already disabled legacy job is removed as well. If dependencies were already
merged, the preview contains only `Job: Present -> Removed` for that legacy entry.

On the operational host, after confirming no manual preparation/transfer is running,
use `Publish-SmartM365WorkplaceCutover` with the same root, the two expected hashes,
an absolute `-BackupDirectory` outside the repository/orchestrator root, and
`-NoManualPreparationConfirmed`. The helper asks for confirmation, requires fresh
healthy resident heartbeats and no running/pending preparation, verifies a unique
SHA256 configuration backup, and uses the existing optimistic publication API.
`-WhatIf` performs no writes. A stale preview is rejected. Repeating an accepted
cutover is a no-op. No collectors, preparation or transfer are requested by it.
Residents may independently run already scheduled jobs: this is not a scheduler
pause or a distributed no-launch lock. Global maintenance is left unchanged.

After publication, verify residents have loaded the new configuration and that the
legacy job is absent before separately authorizing a unified preparation.
Execution history and configuration backups are retained; do not delete old
claims, logs or data as part of this targeted job removal.
Do not bypass a source-publication guard: inspect the failing captured producer
`.current.json.txt` and `.run.json.txt` pair, not an older synchronized receipt.

Only after server batch/transfer and private Desktop acceptance, retire the old
product, preparation/publisher, legacy contracts and data roots from active use.
Verify external backups of both prepared roots first. Port the legacy synthetic
fixtures/golden comparisons used by `test_native_evidence.py` before deleting
legacy Python modules/tests. Keep collectors, `SmartM365-CmdbReceipt.ps1` and
`SmartM365-CmdbSources.json.txt`, which remain shared native evidence dependencies.

`SmartM365-WorkplaceEvidence-Prepare.ps1` is a thin SmartM365 wrapper. The evidence generators, the
publication pipeline (`PreparedEvidencePipeline.psm1`) and the table contracts live in the
SmartWorkplaceIntelligence product. Both are deployed as part of the repository; there is no separate
package.

## 1. Deploy

Deploy the reviewed repository so that `SmartM365` and `SmartWorkplaceIntelligence` stay **siblings**,
for example `\\<server>\<share>\SmartM365` and `\\<server>\<share>\SmartWorkplaceIntelligence`.
The wrapper finds the product next to SmartM365 unless `WorkplaceIntelligenceRootPath` is set in its
runtime JSON. Deploy both folders together: a generator fix in SmartWorkplaceIntelligence has no effect
while the old product files remain.

The contracts are `config/prepared-evidence-contract.json.txt` and `config/prepared-source-contract.json.txt`.
The SmartM365 JSON transport reads the `.json.txt` name first, so a stale `.json` copy beside it is never
used; remove such leftovers instead of editing them.

Keep existing private runtime configuration (`*.local.json.txt`, with supported legacy `.local.json`), tenant profiles and account-classification rules
(private `SmartM365/SmartInventory/Config/AccountClassification.local.json.txt`, never in the repository), source data and scheduled tasks unchanged.
Normal and ValidateOnly runs download the two private classification workbooks from the configured
SharePoint location. Offline runs instead require them in the tenant data root (parent of DATA-LAST and
DATA-ALL); Offline never downloads a missing workbook or falls back to SharePoint.

## 2. Preflight

Run `SmartM365\SmartInventory\SmartInventory.cmd WorkplaceEvidence-Validate` under the
collection account. It selects `-ValidateOnly` (not Offline) for the tenant resolved from `-Tenant`,
`SMARTM365_ORCHESTRATOR_TENANT` or `DefaultTenant`, can create the missing runtime JSON
from the template and merge missing keys, and checks source presence, fingerprints and transport age. It
downloads and structurally checks the two mappings. It generates no prepared CSV, publishes no batch and
sends no notification. Execution logs/transcripts may still be uploaded in non-Offline modes; this is
separate from prepared-data publication.

Verify that the effective profile points to the intended sibling DATA-LAST and DATA-ALL; DATA-POWERBI is
derived from their parent. Do not weaken age, missing-source, tenant or empty-table checks to pass preflight.

## 3. Scheduled and requested runs

The template enables `WorkplaceEvidence-Prepare` daily with `FreshSuccess` dependencies and a
`DependencyMaxAgeHours` floor of 48 hours. The automatic dependency window may be longer for a less
frequent collector. Read the reviewed `Orchestrator/Orchestrator-Jobs.json.template` for the intended
schedule, and the published shared runtime manifest for the actual schedule. A local template edit or Git
update does not change an existing deployed manifest (see [the orchestrator guide](../../Orchestrator/README.md)). For an
out-of-schedule run, ask the orchestrators instead of starting the script by hand:

```
\\<server>\<share>\SmartM365\SmartInventory\SmartInventory.cmd WorkplaceEvidence-Prepare -Request
```

`SmartInventory.cmd WorkplaceEvidence-Prepare-Offline` remains available for a local rebuild without API
collection, upload or notification. It still writes logs, snapshots inputs, generates all CSVs and publishes
the validated local DATA-POWERBI pointer; do not overlap it with collection or another preparation run.

## 4. Checking a batch

A local publication validation failure does not advance `DATA-POWERBI/current.json.txt`. A later cloud
transfer failure can occur after local publication; check local and remote status separately. If the cloud
pointer upload was attempted but read-back failed, activation is uncertain until a verified retry.
For a new batch, check that `validation.json.txt` reports Passed with every contract file, that source
dates and coverage are plausible and that historical coverage is retained: the publication guard refuses
loss of a stable historical key compared with the previous batch, except a deliberate contract
`historyKeyVersion` increase recorded as `HistoryKeyResets`. Compare expected workforce, device,
license and mailbox counts before relying on the batch. Never post private logs or manifests publicly.

The synchronized SharePoint folder is not an atomic multi-file transaction; never assume that the pointer
arriving means every batch file has arrived. Validate transfer completion before a Power BI refresh.

## 5. Rollback

Redeploy the reviewed previous revision of both sibling folders when code rollback is approved. Code
rollback does not automatically roll back the selected data batch. Local retention keeps the current and
immediately previous valid batch payloads; older validated batches retain small audit receipts, not all
CSVs. Raw SmartInventory history is not deleted by this policy. Verify that a compatible retained batch
exists before planning data rollback or using `PreparedBatchId`; an older pinned batch is not protected
from retention. Do not delete or rewrite pointers as an undocumented rollback step.

## CMDB operation and logs

CMDB uses PowerShell 7, Python 3.10+ (standard library only), Core/TenantContext
and the effective tenant profile. Its generators and executable contract remain
next to `SmartM365-CmdbEvidence-Prepare.ps1`. See [CMDB preparation](CMDB-PREPARATION.md)
for commands, automation, receipt qualification, publication and recovery.

Core writes preparation logs/transcripts in the configured
`LOG-ALL/SmartM365-CmdbEvidence-Prepare` folder; the initialization output path
`LOG-ALL/Preparation/CMDB` is not their location.
Normal `EnableSharePointUpload` controls closed-transcript/run-log upload only.
`__USE_GLOBAL__` inherits the tenant/global setting. Batch publication requires
`-Publish` after successful generation. `-ValidateOnly`, including
`-ValidateOnly -Publish`, transfers neither logs nor data. The separate publisher
keeps its automatic log uploads/notifications disabled and writes its own logs
under `LOG-ALL/Publication/CMDB`.

Verify that source/output roots are fully resolved. Literal `{{WorkspaceRootPath}}`
or `{{ProfileKey}}` tokens in a path indicate configuration resolution failure,
not proof that a producer failed. Inspect the effective root before recollecting
or substituting files. Missing configuration keys merge without overwriting
existing private values; Git does not update runtime schedules.

Command examples use `-ExecutionPolicy Bypass`, matching the orchestrator's child
processes. This process-level setting does not override higher-priority Group
Policy or change Authenticode signatures, certificate stores or private-key
permissions. No certificate import is required merely to remove publisher prompts
in this launch mode.

## Intelligence capture, validation and diagnostics

## Manual diagnostics

The catalog entries (`SmartInventory.cmd <Name>`):
- WorkplaceEvidence-Validate: ValidateOnly, with SharePoint mapping reads.
- WorkplaceEvidence-Prepare: full preparation with SharePoint mapping reads.
- WorkplaceEvidence-Prepare-Offline: full local preparation.
- WorkplaceEvidence-Transfer: transfer the existing validated local batch only, without collection, mapping downloads or CSV generation. Optional `-ExpectedBatchId <id>` refuses a different current batch. This explicit mode enables its transfer without changing the saved EnableSharePointUpload setting; it cannot be combined with Offline or other diagnostic/repair modes.

ValidateOnly downloads/structurally checks the two workbooks, captures the required sources and checks transport age, CSV record shape and every row TenantKey on the verified copies. It aggregates validation failures with relative paths, including history. Later live-source changes do not invalidate these copies. It writes temporary source copies and small private audit files but does not generate/publish prepared CSVs or notify, nor certify collection success or business completeness. Non-Offline execution-log upload remains separate from prepared-data publication. Offline suppresses all SharePoint access and notifications: it requires the two workbooks already present in the tenant data root and never downloads missing files. A full offline run still writes logs, temporary snapshots, prepared CSVs and the validated local pointer.

The normal orchestrator job uses neither switch. Existing configured notifications apply. SharePoint mapping reads are independent of EnableSharePointUpload. EnableSharePointUpload defaults to true in the template, so normal runs transfer the validated batch; set it to false in the runtime JSON to keep the batch local. Run logs and transcripts are written to LOG-ALL and uploaded to SharePoint in every mode except Offline, independently of the batch transfer setting. Read access requires the existing app's Sites.Selected role and a read (or existing write) grant on the selected site. No permission or tenant grant is changed by this code.

### Validation and history
 
#### Manual Workforce memory diagnostic

Run `SmartInventory.cmd WorkplaceEvidence-WorkforceDiagnostic` on the collecting machine to execute only the unchanged Workforce calculations, including all available history. Do not add this diagnostic mode to the scheduled job. It uses the same tenant profile/configuration and downloads the two mapping workbooks from SharePoint unless explicitly Offline. No raw collector, DATA-POWERBI publication, Power BI operation, prepared-output upload or notification is performed. Diagnostic and normal preparation runs share the existing preparation locks.

The child process runs New-WorkforceIdentityEvidence.ps1 directly against existing raw exports (read-only, not an atomic snapshot). All five output paths are redirected to a unique private PreparedWorkRootPath/workforce-diagnostics/<run>/outputs folder; no production CSV is overwritten. There is no extra full source copy and no automatic cancellation or memory-setting change. Avoid a simultaneous normal preparation run, but leave other machine activity representative when investigating global memory pressure.

The parent process samples worker private bytes/working set and system available physical memory, committed bytes and commit limit approximately every five seconds. The worker writes opt-in stage boundaries with managed-memory measurements without forced garbage collection. Stages cover individual imports, indexes, identities, personas, exports and each historical week. The parent captures stdout/stderr and exit code even when the worker exits abnormally. A whole-machine crash can still prevent final reporting. Missing system counters are recorded as unavailable, not zero; the sampled private-byte maximum may miss peaks between samples.

Keep memory.csv, stages.ndjson, environment.json and result.json for diagnosis. Worker stdout/stderr are separate private logs and can include file paths/error details. Diagnostic outputs contain client data and must never be committed or uploaded publicly. Diagnostic artifacts are retained intentionally for inspection, not swept by normal run cleanup. An isolated success does not prove production publication or output equivalence. No column reduction, account exclusion, history truncation or business-rule optimization is introduced by this instrumentation.

`SmartWorkplaceIntelligence/config/prepared-source-contract.json.txt` defines the required current CSVs, mapping workbooks, daily statistics and historical source families. Consult the deployed contract rather than a fixed source count in documentation. Missing required sources/history fail rather than become zeros; explicitly optional sources must retain their unavailable/stale status rather than appear healthy. Current-file transport-age limit defaults to 336 hours (14 days), with 744 hours (`PreparedSourceAgeOverrides`) for license prices and for every Microsoft 365 usage report CSV of the UserActivity and Copilot collectors (`M365_Apps_*`, `M365_Copilot_UserUsage`, `M365_Email_*`, `M365_Groups_Activity`, `M365_Mailbox_Usage`, `M365_OneDrive_*`, `M365_SharePoint_*`, `M365_Teams_*`, `M365_Users_Activity`). These configurable limits are not a guarantee of business freshness or collector success.

Every source CSV must contain TenantKey and all rows must match the active tenant, including history. PreparedAllowLegacyTenantless must remain false: a global tenantless bypass is rejected. There is no exception registry. Unexpected empty output tables are rejected unless explicitly named in PreparedAllowEmptyTables.

source validation uses semicolons for Exchange_OnPrem_Servers_Inventory.csv as well as daily statistics, matching the evidence generators. Other source CSVs use commas. The same strict record-shape and tenant checks apply to both formats; malformed files are not retried with a different separator. No source CSV conversion is needed.

A compiled strict reader (`SmartWorkplaceIntelligence/scripts/PreparedCsvReader.cs`, deployed with the scripts) performs these checks for well-formed files. Blank lines, ambiguous quoting, a field-count or tenant mismatch and any read error send the file back to the original TextFieldParser check, so results and error messages are unchanged; the final validation log line reports how many files used that reference parser. The Workforce weekly history reads its selected columns with the same reader and falls back to Import-Csv in the same way. The native build and the SharePoint transfer log per-table and per-file durations (timings only, no values). Before upload and again before the root pointer, the transfer hashes the local batch CSVs four files at a time; every SHA256 is still compared individually, and any changed file stops the transfer before the pointer. Just before each upload only the file size is re-checked; the uploaded bytes are verified by the SharePoint read-back against the manifest SHA256.
#### Legacy tenant-key repair (manual only)

`SmartInventory.cmd WorkplaceEvidence-RepairTenantKeys` previews explicitly selected
weeks and expected file count. `-ApplyRepair` requires confirmed source ownership.
It never changes existing TenantKey values or DATA-LAST. Original bytes are
backed up/hash-verified; every parsed field/order/count is checked. Serialization
may change, while original timestamps remain. Partial replacement requires
inspection of the private journal and backups; no automatic overwrite/rollback.
This is operator-attested history, not reconstructed collection proof.
Do not run repair on current valid inputs or add it to scheduled jobs.

Inputs are snapshotted into PreparedWorkRootPath outside synchronized DATA (default OS temp, separated by profile). the source list is selected once, without an expensive initial global hash scan. Each file is streamed into its copy without blocking collector writes, then the copied bytes are checked against a fresh source SHA256, local-copy SHA256, size and modification timestamp. A failed check retries only that file. A newer file than the initial plan is allowed; provenance records the actual captured version, time and attempt count. New historical files arriving after selection are included on the next run. There is no final comparison to the live source tree and no claim of an atomic, collector-wide snapshot. All validators and generators use only the verified copies; source freshness, tenant, schema and historical-coverage gates remain enforced.

Generators run in separate child processes. `PreparedMaxParallelGenerators` (default 3, range 1-8) lets independent generators run concurrently; a generator starts only after the generators whose staged outputs it reads (Devices before Workforce, EndpointProtection and WindowsMigration; Workforce before Mailbox and Security; IntuneOperations before Trust), and the account-classification check still runs just before Workforce starts. In parallel mode each generator log is printed in one block when it completes; 1 restores the former sequential order with live output. A failed generator stops the others and nothing is published. Validation covers output schemas, types, declared unique keys and loss of historical coverage compared with the last accepted batch. Each History or Trend table declares its comparison key (`historyKey`) in the contract: only stable periods and dimensions (week, service, metric), never per-user values or the provisional snapshot date of the current week, which the next run legitimately replaces. A History or Trend table without a declared key stops the publication, and a refused publication names the missing key.

`M365_Apps_Activations.csv` is optional for the M365 user activity producer: when Graph keeps returning the report empty, the previous DATA-LAST CSV is kept unchanged and left out of the receipt, so its write time stays the time of its last successful acquisition. `PreparedOfficeActivationWarningAgeHours` (integer hours, default 72, range 1-8760) is the warning threshold on that write time. It is an optional setting that is deliberately absent from the template: add it by hand to the private `SmartM365-WorkplaceEvidence-Prepare.local.json` only to change the default; when it is absent or empty, 72 hours applies. It is measured on the acquisition, not on `Report Refresh Date`, because Graph usage reports normally lag two to three days. Past the threshold the Workforce generator writes a WARNING to the Prepare log (latest Report Refresh Date included) and Data Trust reports the `Microsoft 365 Apps activations` source as `Aging` with Decision Readiness `Ready`, so the Power BI Data Trust page counts it. It never blocks publication; the publication-age limit (`PreparedMaxSourceAgeHours`, 336 by default) still rejects the whole batch when the file is older than that limit; the template sets a 744-hour override for `M365_Apps_Activations.csv` and the other Microsoft 365 usage report CSVs in `PreparedSourceAgeOverrides`, added to existing private configurations at the next start, so a report kept empty by Microsoft for up to 31 days does not block publication. The threshold used is recorded in the batch provenance (`OfficeActivationFreshness`).

`M365_Email_AppUsage.csv`, `M365_OneDrive_UserActivity.csv`, `M365_Groups_Activity.csv` (UserActivity 1.34) and `M365_Copilot_UserCountTrend.csv` (CopilotUsage 1.1.0) are optional current files of the source contract, with a 744-hour limit set in the contract itself (`PreparedSourceAgeOverrides` does not apply to optional files). When one is absent, too old or empty, the capture records it in `OptionalSourceIssues`, the batch continues, the usage signals evidence shows `Not collected` instead of zeros, and Data Trust reports the source `Optional unavailable`. See `SmartWorkplaceIntelligence/docs/M365-USAGE-SIGNALS.md`.

Workforce history uses each week's own Entra and AD snapshots for population membership, so later departures do not remove those users from past weeks. Non-human exclusions use current governed classification for accounts still linked in the current AD population, and that week's classification for departed accounts. When no usable AD snapshot exists for a week, the generator explicitly falls back to the current workforce population and records the limitation. This is therefore not a fully independent historical census for every week. Preserve per-user history while User Explorer requires user-filtered trends.

Trend keys must stay stable from one run to the next. In `Executive KPI Trends`, the Exchange Online adoption and Microsoft 365 active-use points are keyed by the Monday of their source week (the latest observation of the week, DATA-LAST for the current week). A deliberate key-grain change is declared by raising the table's `historyKeyVersion` in the contract (default 1). The next publication then skips the coverage comparison for that table once, logs a warning, and records it under `HistoryKeyResets` in `batch.json.txt`. A version lower than the previous batch is rejected. `Executive KPI Trends` uses version 2. The daily AD metrics are unchanged.

Publication writes each tenant-prefixed table to a temporary file in the new batch and renames it to a name that does not exist yet; it no longer replaces a file it has just written. The staging file hash is checked before and after the rewrite. On the SMB share, the file server (antivirus, indexing) can hold a just-written file for a moment. A refused or busy rename, including the final `current.json.txt` replacement, is therefore retried 4 times over about 30 seconds. Each retry is logged as a warning. A persistent failure stops the publication with the file path, the previous pointer is unchanged, and the failed batch is audited under `failed/<batch-id>`.

### Versioned output and refresh

DATA-POWERBI/current.json.txt selects an immutable batches/<batch-id> directory containing the cohort's declared CSVs/artifacts, batch.json.txt, validation.json.txt and a matching current.json.txt pointer copy. The content remains JSON; only the transport filename changes. Only a fully validated batch advances the root pointer. Publication alone retains existing batch payloads. Qualified coupled retention, described below, runs only after successful SharePoint transfer/read-back. Failed publication payloads are removed while diagnostics remain under failed/. No raw history is deleted. Outputs are single-tenant and must not be combined as a multi-tenant model.
#### Legacy metadata conversion (manual only)

Current metadata uses `.json.txt`. Readers prefer it and reject malformed preferred
files rather than falling back to stale legacy copies. An unconverted legacy root
pointer prevents normal publication.
`SmartInventory.cmd WorkplaceEvidence-ConvertMetadata` previews validated legacy
renaming; `-ApplyMetadataConversion` preserves bytes, batch IDs and hashes,
publishing the root pointer last. Conflicting old/new names stop conversion.
Do not refresh/transfer during conversion or mass-rename unrelated files.
Current inputs need no conversion; installing code never runs it automatically.

Normal success or failure removes this run's source/staging copies and downloaded mappings, retaining capture/run audits, failure diagnostics and preparation logs. The shared JSON transport resolves these audit filenames, preferring `.json.txt`. An interrupted run's owned, marked source/staging payload is cleaned on the next run under the preparation locks. Unmarked legacy scratch folders, foreign-tenant folders and unrecognized batch folders are never automatically deleted. Cleanup refuses paths outside the declared root and symbolic links/junctions. A deletion failure is reported rather than silently presented as reclaimed disk space. Old batches without a verifiable publication receipt are retained with a warning. Small private logs/audits are not automatically purged; monitor disk space. No existing production folders are retroactively cleaned merely by installing this update.

Collectors may update raw sources after capture. Avoid overlapping prepared-output transfer/replacement with Power BI refresh: local atomic pointer replacement is not an atomic SharePoint transfer or a cross-query transaction. Verify transfer completion first. PreparedBatchId can pin a verified immutable batch; declare that batch in the qualified consumer registry before allowing retention. Publisher SHA256 checks and Power Query byte-length checks have different scope. Raw-source/staging cleanup remains local and separate from coupled batch retention.

Direct SharePoint transfer uses the shared Core uploader and downloader and requires the existing site read/write grant. When PreparedSharePointFolderPath is blank, the target is the normal SharePointTargetFolderPath after the shared CSV-to-DATA normalization, with `/DATA-POWERBI` appended. An explicit override must identify a library-relative DATA-POWERBI folder outside raw/log folders. Required configuration is checked before expensive generation. Missing destination folders are provisioned by the existing Core helper. No tenant URL or credential is hard-coded.

Transfer uses the same transfer implementation for normal enabled publication and TransferOnly recovery. It validates the current batch identity, all expected files, validation receipts and local SHA256/size first. It then uploads and downloads each CSV and the three batch metadata files, comparing SHA256 and size. The root current.json.txt is uploaded last, and also read back. For a 24-CSV batch, success means 28 files verified (24 CSVs, three batch metadata files, one root pointer). Reading back every file adds network traffic; no full duplicate batch is kept locally. At most one downloaded file is kept temporarily, under PreparedWorkRootPath/transfers/<run>/readback.tmp, removed after each check and on failure. A small private transfer.json audit remains. No historic CSV is recomputed or removed.

The local publication lock is held during transfer to block replacement/retirement by publishers sharing the same DATA-POWERBI root. It is not a distributed SharePoint lock: all writers for a tenant must use that shared root and avoid unrelated simultaneous cloud publication. Readers must wait for verified completion and synchronization; pointer-last publication is not an atomic OneDrive synchronization or a transaction across Power BI queries. An upload/read-back failure before pointer publication leaves the old root pointer untouched by this run. If pointer upload was attempted but its verification fails, activation is uncertain and the script explicitly reports failure; do not refresh until a successful retry. Retrying resends the same immutable files, with no recalculation or remote deletion.

Normal runs transfer when EnableSharePointUpload is enabled (template default true ; an existing runtime JSON keeps its local value). The scheduler entry alone does not change that setting. TransferOnly is a manual recovery mode, not a replacement preparation job. Qualify the actual remote transfer; synthetic transport tests do not establish live SharePoint success.

## Qualification boundary

### Qualified coupled batch retention

Deploy the reviewed preparation, transfer and retention components together.
The wrapper calls retention after both normal transfer and `TransferOnly`
success, passing the published batch ID and verified pointer SHA256. Offline,
ValidateOnly, failed transfers and local publication alone cannot trigger coupled
retirement. No additional scheduled task is installed. A retention exception is
logged separately and leaves successful publication and its pointer intact.

Blocked retention is also reported as a warning, with `Plan.BlockedReasons` in
the main execution log. A standalone retention command keeps its own completion
banner; the wrapper uses `-NoCompletionBanner` to emit only the overall result.
To inspect an existing incident without rerunning preparation or retention,
read the saved plan's `BlockedReasons` and rows with `State = Blocked`, plus the
saved journal. A consumer registry qualification failure requires reviewing the
complete consumer list and pinned batches. A target mismatch requires checking
the private configured target against the saved resolved target. Do not enable
Auto or bypass either check to clear a warning. Rows without verified published
chain membership remain blocked independently of those configuration checks.

`PreparedRetention` is added from the wrapper's committed template when its
runtime configuration is loaded; existing local values are preserved. Private
configuration and registry deployment are operator actions, separate from Git.

| Setting | Contract |
| --- | --- |
| `Mode` | `Preview` by default. Set private runtime to `Auto` only after qualification. |
| `ConsumersPath` | Blank selects `SmartM365-WorkplaceEvidence-Prepare.consumers.json.txt` beside the wrapper configuration. Missing/invalid registries block application. |
| `PlanMaxAgeHours` | 24 by default; accepted range 1..168 hours. This expires plans, never protections. |
| `Target` | Exact `TenantId`, `ClientId`, `SiteId`, `DriveId`, `RootId`, `BatchesId`, library-relative `RootPath` ending in `DATA-POWERBI`. No customer defaults. |
| `Qualification` | Initially `Unqualified`. Requires the evidence below; a boolean alone cannot authorize retirement. |
| `JournalRootPath` | Optional key, not templated. Required for Apply and Auto: an absolute shared folder outside `DATA` and its ancestors (for example a sibling of `DATA` on the server share). Plans and journals are written under `<JournalRootPath>/retention/<run-id>/`. Without it, Preview falls back to the local work root. May use effective global tokens, for example `{{WorkspaceRootPath}}\Tenants\<journal-folder>\{{ProfileKey}}`. Use `{{ProfileKey}}` (the tenant folder name, for example `prod`), not `{{TenantKey}}` (the organization-environment key): only the global path keys translate `{{TenantKey}}` to the folder name. The resolved path must still be absolute and outside `DATA`, and an unknown or empty token blocks the run. |
| `SharePointKeepBatches` | Optional, code default 2 (current and previous validated batch), range 2..50. |
| `ServerKeepBatches` | Optional, code default 0 = keep the full server history; otherwise 2..100000 and at least `SharePointKeepBatches`. |
| `MaxSharePointRemovalsPerRun` | Optional, code default 1 (at most one SharePoint batch folder per publication), range 1..50. |

The last four keys are deliberately absent from the template (see the fingerprint note below): add them to the
private configuration only when a value other than the code default is needed, then requalify.

Copy the neutral consumer template from
`SmartWorkplaceIntelligence/config/templates/PreparedBatchConsumers.json.template`
to the private wrapper directory. Set the tenant identity, qualified review
(`Status`, `ReviewedBy`, UTC `ReviewedUtc`, `DeclaredConsumersComplete=true`) and
`Consumers`. Every entry requires `ConsumerId`, `BatchId`, `Reason`, UTC
`ReviewedUtc`. A consumer may protect several batches; duplicate pairs are
rejected. There is no automatic expiry. Changes require a new explicit registry
review and matching qualification fingerprint. A server cannot discover all
distributed reports. Register a consumer before it starts using an older batch;
coordinate reference changes with the same publication authority.

First run the standalone command in Preview on the authoritative server. Preview
defaults to `-Verification Metadata`: it lists every file's identity and size,
reads the three receipts per batch and both root pointers, and validates
their ownership, manifest/validation consistency and declared file sets. It never
downloads or hashes payload CSVs/artifacts, including on the local server. Its
private `plan.json.txt` contains resolved target IDs, live access-evidence hash,
configuration/registry/module hashes, both pointers, protection reasons and the
per-store inventory with logical bytes and `Protected`, `Candidate` or `Blocked`
decisions. Payload SHA256 values are explicitly unmeasured, not copied from the
manifest as if they were measured. Same-size payload corruption is not detectable
by this inventory. Structurally valid pins remain visible even while the registry
is unqualified; the qualification failure still blocks all application.
Plans use schema 3: parsed receipt bodies/provenance are not embedded. Each raw
receipt's measured SHA256, byte length and identity remain in the file inventory
and snapshot fingerprint. Receipt validation still precedes projection; original
receipts are archived and verified before removal. Regenerate older plans.

The only permitted extra local files are empty, idle sidecars named exactly
`<owner>.json.transport.lock` or `<owner>.json.txt.sharepoint-transition.lock`,
where the corresponding `<owner>.json.txt` is a receipt or declared JSON artifact.
They remain in the physical inventory, counts, snapshot and pre-removal checks;
only cross-store content comparison excludes them. Active/inaccessible, nonempty,
orphaned, unknown or remote sidecars block the affected copies. Before cloud or
local retirement, local sidecars are reopened without creating them and held
against helper writes through archival and removal. No sidecar is deleted by
Preview. Protected rows expose `ValidationReason` separately from pin reasons.
An invalid policy still
produces a blocked plan when inventory access is available. Authentication,
enumeration or read-back failures can instead stop with a failure journal.
For a complete content audit, use `-Mode Preview -Verification Full`. It streams
remote payloads through one temporary file and hashes local files; expect
significant network traffic. Progress identifies each SharePoint batch. Files
are rechecked before removal. Bytes are logical payload
sizes, not a promise of disk or SharePoint quota reclamation.

Qualification requires `Status=Qualified`, `PolicyVersion=PreparedBatchRetention/v1`,
the operator and UTC review date, canonical `ServerOutputRoot`, `RegistrySHA256`,
`ModuleSHA256` of PreparedSharePointRetention.psm1, `AccessEvidenceSHA256`,
`Coordination=SharedServerPublicationLock`, documented `WriterExclusionEvidence`,
`ExclusiveCloudWritersConfirmed=true` and `RecoverabilityConfirmed=true`.
`ServerOutputRoot` may be written as `{{WorkspaceRootPath}}\Tenants\{{ProfileKey}}\DATA-POWERBI`
so that it follows the global data root; it is resolved from the effective
global configuration and then compared case-sensitively with the server
`DATA-POWERBI` root, exactly like a literal path. A tokenized value no longer
detects an unreviewed change of the global data root by itself.
Obtain hashes from the actual deployed files and live Preview, never example
values. Qualify the existing backlog too; incomplete old formats remain blocked.
Changing a published batch does not invalidate this policy qualification.
Changing the implementation, registry, target or reviewed access evidence does.

All prepared-output publishers and retention workers must use the SAME
authoritative server/SMB DATA-POWERBI root and its exclusive `.publication.lock`.
Retention holds that lock across planning and application, and read-only file
handles exclude writes/replacement of the private configuration and registry.
The publication handle is checked with a durable read/write round trip before
mutations; an unusable handle stops the operation.
The live access fingerprint includes the server root ACL and both SharePoint
root/batches permission listings. A site owner must additionally qualify inherited
group membership, other app grants, administrative access procedures and exclusion
of direct SharePoint/OneDrive writers for the entire operation. These cannot be
proven exhaustively by the API listings. If that exclusivity is not enforceable,
leave Preview enabled. Never treat a synchronized copy as this server authority.
Never use a synchronized review-workstation data copy as a retention target.
Folder eTags are additional preconditions, not a lock on descendant contents.

Every key added to `SmartM365-WorkplaceEvidence-Prepare.local.json.template` is written into the private
configuration at the next Prepare or retention start, which changes its fingerprint (`ConfigurationSHA256`):
a prior-run journal can then no longer authorize a pending local retirement. Prefer optional settings with
a code default (documented, not templated) while retention runs in Auto.

After qualification, set `Mode=Auto` privately. Each successful publication then
creates and fingerprints a fresh plan, verifies it under the common lock and
applies its eligible actions without manual per-plan approval. SharePoint keeps
immutable batch folders and switches its root pointer; batches are never
overwritten in place. Current/previous from both pointers, all consumers, and the
`SharePointKeepBatches` newest complete valid SharePoint batches (default 2) are
protected on BOTH stores; an absent SharePoint predecessor therefore cannot leave
only the current cloud copy protected. Consumer pins can keep additional
SharePoint folders. The server share keeps the long history: with the default
`ServerKeepBatches=0` every server batch is reported `Retained` and never removed;
with a positive value the newest `ServerKeepBatches` valid server batches are
protected on the server only. At most `MaxSharePointRemovalsPerRun` SharePoint
folders (default 1, oldest first) are retired per run; further candidates are
`Deferred` to the next publication, together with their server copy. The
selection uses the validated UTC timestamp in the batch ID, never file modification
times. All ties at the boundary timestamp are retained. Incomplete, failed,
schema-only or corrupt batches do not displace valid copies. Fewer than two valid
copies on either store blocks all application. The plan exposes `RetentionPolicy`
and `MinimumRetainedBatches` per store.

Metadata Preview selects this minimum provisionally from receipts, file sets and
sizes, without downloading payloads or claiming SHA256 verification. Apply/Auto
always recomputes it with Full verification. Immediately before EVERY cloud or
local removal, including partial-run recovery, the selected retained copies on
both stores are rechecked for identity, complete file sets and every payload hash.
Missing, inaccessible or changed retained evidence stops further removal; no
older replacement is silently selected during application. This adds payload
read-back traffic per removal. The existing shared-lock and writer-exclusion
qualification remains required; this minimum is not a substitute for recovery
evidence. Existing plans must be regenerated after the module changes.

Pointer divergence,
missing qualification or changed global controls block the whole application.
A lot with incomplete receipts, unexpected content or unknown chain membership
is individually blocked; other independently qualified lots can proceed. Verified
archived publication links permit traversal through previously retired batches.

The manual entry point is `SmartWorkplaceIntelligence/scripts/Invoke-PreparedSharePointRetention.ps1`.
`-Tenant <profile> -Mode Preview` creates a fast inspectable inventory, marked
`Verification=Metadata` and `ApplySupported=false`. Apply rejects this inventory
before downloading any remote file. To obtain an applicable plan, first use
`-Mode Preview -Verification Full`, review that newly generated plan, then use
`-Mode Apply -PlanPath <private-plan.json.txt> -PlanSHA256 <hash>`.
Manual Apply also requires qualified controls and fresh state; it cannot bypass
protections or use a stale plan. Plan schema 2 records the verification level;
older plans must be regenerated. Apply and configured Auto always run the complete
SHA256 inventory and pre-removal verification, regardless of `-Verification`.
`-Mode Configured` follows the private mode and
is the wrapper's automatic path. In private Preview mode it uses the fast
inventory. No synchronized workstation-copy mode is provided. The common server
lock remains held during either inventory so its two pointers are coordinated.
Installing or testing code never activates Auto. A module change invalidates its
existing qualification fingerprint; qualify the deployed implementation again.

Every change to `PreparedSharePointRetention.psm1` changes `ModuleSHA256`. Until
the operator reruns the private qualification procedure against the deployed file,
every run is blocked with `Registry or loaded retention implementation requires
qualification`. Before returning to Auto after such a change: deploy (share Pull),
keep `Mode=Preview`, set `JournalRootPath`, rerun the qualification procedure to
record the new `ModuleSHA256`, review a Full Preview plan, perform one manual
`-Mode Apply` of that plan and check its journal, and only then set `Mode=Auto`.

Before each removal, the three receipts are archived under `retired/<batch-id>`
on the server and SharePoint with byte/hash verification and remote read-back.
Different existing receipts cause an error. Cloud removal uses Graph v1.0
`DELETE /drives/{drive-id}/items/{item-id}` with `If-Match`; Microsoft documents
this API as a recycle-bin operation, not permanent deletion. No purge endpoint,
bypass-lock header or automatic restore operation is used.

The exact request is built and validated before `RemovalRequested` is journaled:
item ID characters, folder eTag, absolute `https://graph.microsoft.com/v1.0/drives/{id}/items/{id}`
URI without query or fragment, and the installed `Invoke-MgGraphRequest` must expose
`-SkipHttpErrorCheck` and `-StatusCodeVariable`. A refusal is `Blocked`, with no
request sent. The DELETE uses `-OutputType PSObject -SkipHttpErrorCheck
-StatusCodeVariable`; it never reads an `HttpResponseMessage` and never calls
`Set-MgRequestContext`. Outcomes, recorded with `RequestUri`, `HttpStatus` and
`Outcome` in the journal event:

| Result | Journal state |
| --- | --- |
| HTTP 204 and the folder is absent | `Recycled` |
| HTTP 404 and the folder is absent | `Recycled`, `Outcome=AlreadyAbsent` |
| HTTP 412 (eTag changed), other 4xx/503 refusal, or pre-send failure, with the folder still listed under its reviewed eTag | `NotRemoved`: nothing was deleted; a later publication plans it again |
| Folder absent after an unconfirmed result | `Recycled`, `Outcome=AbsentAfterError`, with a warning |
| HTTP 204 while the folder is still visible, other 2xx/5xx, timeout/connection failure, or unreadable state | `Uncertain`: operator reconciliation required |

Incident 2026-10-22: the first Auto run failed with `Invoke-MgGraphRequest: Invalid URI`
during the DELETE, before any request reached SharePoint (folder intact, journal
`Uncertain`). Cause, read in the Microsoft.Graph.Authentication 2.30.0 source and
reproduced offline with 2.30.0 by `Test-PreparedRetentionGraphTransport.ps1` (not
re-run on the server itself): an absolute-URI request leaves the SDK session
environment without an authority endpoint; the former `Set-MgRequestContext -MaxRetry 0`
before the DELETE discarded the cached HTTP client, and rebuilding the certificate
credential then raised `UriFormatException`. Both `Set-MgRequestContext` and
`HttpResponseMessage` were removed from this path; the explicit outcome handling
above applies whatever the cause. Other scripts that change the request context
after a first Graph request in the same process can meet the same SDK behavior. See
[Graph driveItem deletion](https://learn.microsoft.com/en-us/graph/api/driveitem-delete?view=graph-rest-1.0)
and [folder eTag limitations](https://learn.microsoft.com/en-us/graph/api/resources/driveitem?view=graph-rest-1.0).
Keep the existing `Sites.Selected` site write grant; permission/read-back failures
block retention, never justify silently broadening application permissions.

SharePoint recycle-bin retention is generally 93 days and can end earlier if an
authorized operator purges it; policies/holds can retain additional copies.
Qualification must include an operator recovery procedure and an actual later
recovery check before production activation. A Graph 204 and absence check prove
the item was removed from the active location, not a tested restoration or freed
storage. See [SharePoint recovery conditions](https://learn.microsoft.com/en-us/sharepoint/sharepoint-data-deletion)
and [Purview retention](https://learn.microsoft.com/en-us/purview/retention-policies-sharepoint).

Cloud retirement precedes the corresponding local payload removal. Journals
persist intent before requests, then record confirmed recycle, local retirement,
blocked actions or uncertain outcomes. Mutating requests are never blindly
retried. A fresh automatic run may finish an intact local lot only when an
unchanged, hash-verified prior plan/journal confirms that exact remote retirement;
new protections are evaluated first. An uncertain response, partially removed
local payload or conflicting history requires operator reconciliation and a fresh
plan. Preserve unresolved runs and their original receipts in a private incident
archive; remove them from the active `retention/` history only after that explicit
reconciliation, never merely to unblock Auto. Do not reuse an old plan after a partial application. Receipts, plans and
journals have no automatic purge in this change; monitor their storage.

Journals and plans are written under `JournalRootPath`, shared by every server and
account that runs retention for the same DATA root; each journal records its writer
`Machine`. Unresolved `Uncertain`/`RemovalRequested` events under that root AND under
the local work root (`PreparedWorkRootPath`, or the account's temporary folder) block
their batch. Read-back temporaries stay in a per-run folder of the local work root,
removed at the end of the run. When moving to a shared root, keep or archive the
earlier per-account journals; they are not migrated automatically.

Run synthetic tests only against a dedicated scratch directory:
`Test-PreparedSharePointRetention.ps1 -TestRoot <scratch>`.
The suite uses file-backed fake SharePoint and synthetic tenants, including
Preview, automatic/manual application, protections across stores, archived links,
stale plans, unsafe targets, lock contention, archive errors and partial recovery,
plus the retention counts, per-run limit, journal root and every DELETE outcome.
`Test-PreparedRetentionGraphTransport.ps1` (PowerShell 7, no parameter) calls the
installed `Invoke-MgGraphRequest` with the module's exact DELETE shape against a
localhost listener and a synthetic token: 204/404/412/429/500 classification,
`If-Match`, an invalid URI reported as a pre-send failure, and a refused connection
kept uncertain. It prints `NOT RUN` (exit 3) without the SDK. Use
`-GraphAuthenticationVersion 2.30.0` (or the server's version) to select an installed
SDK; with 2.30.0 it also replays the incident: the former `Set-MgRequestContext`
path fails with `Invalid URI` before sending, the current path sends the DELETE.
Run it on the server too, since its SDK version can differ from a workstation's.
Also run existing transfer, endpoint/unified transfer, metadata, configuration,
capture/retention, publication and lifecycle-cleanup suites. Offline checks do not
qualify server SMB coordination, SharePoint permissions/recovery or OneDrive
synchronization. Deployment, real qualification and activation remain separate.

Validate the deployed revision and specific batch. Source capture, business counts,
history retention, verified transfer, synchronization and Desktop/Service refresh
are separate evidence. Local or historical success does not qualify later changes.
Diagnostic outputs, tenant logs/manifests and mappings remain private.
