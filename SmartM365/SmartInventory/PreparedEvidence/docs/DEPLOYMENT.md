# Prepared evidence: deployment and operation

This guide covers two independent pipelines. CMDB prepares a flat current cohort
in `DATA-POWERBI-CMDB`; Intelligence prepares versioned/history-aware batches in
`DATA-POWERBI`. The numbered Intelligence procedure below does not apply to CMDB.

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

Run `SmartM365\SmartInventory\Launchers\Cloud\Test-SmartM365-WorkplaceEvidence-Prepare.cmd` under the
collection account. It selects `-Tenant prod -ValidateOnly` (not Offline), can create the missing runtime JSON
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
\\<server>\<share>\SmartM365\SmartInventory\LaunchersByOrchestrator\Request-WorkplaceEvidence-Prepare.cmd
```

`Start-SmartM365-WorkplaceEvidence-Prepare-Offline.cmd` remains available for a local rebuild without API
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

The Cloud launchers follow the existing prod convention:
- Test-SmartM365-WorkplaceEvidence-Prepare.cmd: ValidateOnly, with SharePoint mapping reads.
- Start-SmartM365-WorkplaceEvidence-Prepare.cmd: full preparation with SharePoint mapping reads.
- Start-SmartM365-WorkplaceEvidence-Prepare-Offline.cmd: full local preparation.
- Start-SmartM365-WorkplaceEvidence-Transfer.cmd: transfer the existing validated local batch only, without collection, mapping downloads or CSV generation. Optional `-ExpectedBatchId <id>` refuses a different current batch. This explicit mode enables its transfer without changing the saved EnableSharePointUpload setting; it cannot be combined with Offline or other diagnostic/repair modes.

ValidateOnly downloads/structurally checks the two workbooks, captures the required sources and checks transport age, CSV record shape and every row TenantKey on the verified copies. It aggregates validation failures with relative paths, including history. Later live-source changes do not invalidate these copies. It writes temporary source copies and small private audit files but does not generate/publish prepared CSVs or notify, nor certify collection success or business completeness. Non-Offline execution-log upload remains separate from prepared-data publication. Offline suppresses all SharePoint access and notifications: it requires the two workbooks already present in the tenant data root and never downloads missing files. A full offline run still writes logs, temporary snapshots, prepared CSVs and the validated local pointer.

The normal orchestrator job uses neither switch. Existing configured notifications apply. SharePoint mapping reads are independent of EnableSharePointUpload. EnableSharePointUpload defaults to true in the template, so normal runs transfer the validated batch; set it to false in the runtime JSON to keep the batch local. Run logs and transcripts are written to LOG-ALL and uploaded to SharePoint in every mode except Offline, independently of the batch transfer setting. Read access requires the existing app's Sites.Selected role and a read (or existing write) grant on the selected site. No permission or tenant grant is changed by this code.

### Validation and history
 
#### Manual Workforce memory diagnostic

Run Test-SmartM365-WorkplaceEvidence-WorkforceMemory.cmd on the collecting machine to execute only the unchanged Workforce calculations, including all available history. Do not add this diagnostic mode to the scheduled job. It uses the same prod profile/configuration and downloads the two mapping workbooks from SharePoint unless explicitly Offline. No raw collector, DATA-POWERBI publication, Power BI operation, prepared-output upload or notification is performed. Diagnostic and normal preparation runs share the existing preparation locks.

The child process runs New-WorkforceIdentityEvidence.ps1 directly against existing raw exports (read-only, not an atomic snapshot). All five output paths are redirected to a unique private PreparedWorkRootPath/workforce-diagnostics/<run>/outputs folder; no production CSV is overwritten. There is no extra full source copy and no automatic cancellation or memory-setting change. Avoid a simultaneous normal preparation run, but leave other machine activity representative when investigating global memory pressure.

The parent process samples worker private bytes/working set and system available physical memory, committed bytes and commit limit approximately every five seconds. The worker writes opt-in stage boundaries with managed-memory measurements without forced garbage collection. Stages cover individual imports, indexes, identities, personas, exports and each historical week. The parent captures stdout/stderr and exit code even when the worker exits abnormally. A whole-machine crash can still prevent final reporting. Missing system counters are recorded as unavailable, not zero; the sampled private-byte maximum may miss peaks between samples.

Keep memory.csv, stages.ndjson, environment.json and result.json for diagnosis. Worker stdout/stderr are separate private logs and can include file paths/error details. Diagnostic outputs contain client data and must never be committed or uploaded publicly. Diagnostic artifacts are retained intentionally for inspection, not swept by normal run cleanup. An isolated success does not prove production publication or output equivalence. No column reduction, account exclusion, history truncation or business-rule optimization is introduced by this instrumentation.

`SmartWorkplaceIntelligence/config/prepared-source-contract.json.txt` defines the required current CSVs, mapping workbooks, daily statistics and historical source families. Consult the deployed contract rather than a fixed source count in documentation. Missing required sources/history fail rather than become zeros; explicitly optional sources must retain their unavailable/stale status rather than appear healthy. Current-file transport-age limit defaults to 168 hours, with 744 hours for license prices. These configurable limits are not a guarantee of business freshness or collector success.

Every source CSV must contain TenantKey and all rows must match the active tenant, including history. PreparedAllowLegacyTenantless must remain false: a global tenantless bypass is rejected. There is no exception registry. Unexpected empty output tables are rejected unless explicitly named in PreparedAllowEmptyTables.

source validation uses semicolons for Exchange_OnPrem_Servers_Inventory.csv as well as daily statistics, matching the evidence generators. Other source CSVs use commas. The same strict record-shape and tenant checks apply to both formats; malformed files are not retried with a different separator. No source CSV conversion is needed.
#### Legacy tenant-key repair (manual only)

`Repair-SmartM365-WorkplaceEvidence-TenantKeys.cmd` previews explicitly selected
weeks and expected file count. `-ApplyRepair` requires confirmed source ownership.
It never changes existing TenantKey values or DATA-LAST. Original bytes are
backed up/hash-verified; every parsed field/order/count is checked. Serialization
may change, while original timestamps remain. Partial replacement requires
inspection of the private journal and backups; no automatic overwrite/rollback.
This is operator-attested history, not reconstructed collection proof.
Do not run repair on current valid inputs or add it to scheduled jobs.

Inputs are snapshotted into PreparedWorkRootPath outside synchronized DATA (default OS temp, separated by profile). the source list is selected once, without an expensive initial global hash scan. Each file is streamed into its copy without blocking collector writes, then the copied bytes are checked against a fresh source SHA256, local-copy SHA256, size and modification timestamp. A failed check retries only that file. A newer file than the initial plan is allowed; provenance records the actual captured version, time and attempt count. New historical files arriving after selection are included on the next run. There is no final comparison to the live source tree and no claim of an atomic, collector-wide snapshot. All validators and generators use only the verified copies; source freshness, tenant, schema and historical-coverage gates remain enforced.

All generators run sequentially. Validation covers output schemas, types, declared unique keys and loss of historical coverage compared with the last accepted batch. Each History or Trend table declares its comparison key (`historyKey`) in the contract: only stable periods and dimensions (week, service, metric), never per-user values or the provisional snapshot date of the current week, which the next run legitimately replaces. A History or Trend table without a declared key stops the publication, and a refused publication names the missing key.

Workforce history uses each week's own Entra and AD snapshots for population membership, so later departures do not remove those users from past weeks. Non-human exclusions use current governed classification for accounts still linked in the current AD population, and that week's classification for departed accounts. When no usable AD snapshot exists for a week, the generator explicitly falls back to the current workforce population and records the limitation. This is therefore not a fully independent historical census for every week. Preserve per-user history while User Explorer requires user-filtered trends.

Trend keys must stay stable from one run to the next. In `Executive KPI Trends`, the Exchange Online adoption and Microsoft 365 active-use points are keyed by the Monday of their source week (the latest observation of the week, DATA-LAST for the current week). A deliberate key-grain change is declared by raising the table's `historyKeyVersion` in the contract (default 1). The next publication then skips the coverage comparison for that table once, logs a warning, and records it under `HistoryKeyResets` in `batch.json.txt`. A version lower than the previous batch is rejected. `Executive KPI Trends` uses version 2. The daily AD metrics are unchanged.

Publication writes each tenant-prefixed table to a temporary file in the new batch and renames it to a name that does not exist yet; it no longer replaces a file it has just written. The staging file hash is checked before and after the rewrite. On the SMB share, the file server (antivirus, indexing) can hold a just-written file for a moment. A refused or busy rename, including the final `current.json.txt` replacement, is therefore retried 4 times over about 30 seconds. Each retry is logged as a warning. A persistent failure stops the publication with the file path, the previous pointer is unchanged, and the failed batch is audited under `failed/<batch-id>`.

### Versioned output and refresh

DATA-POWERBI/current.json.txt selects an immutable batches/<batch-id> directory containing 24 CSVs, batch.json.txt, validation.json.txt and a matching current.json.txt pointer copy. The content remains JSON; only the transport filename changes. Only a fully validated batch advances the root pointer. The current and immediately previous valid batches remain recoverable; older published CSV batches are retired only after a successful replacement. Their small manifests/validation receipts remain under retired/. Failed publication payloads are removed while diagnostics remain under failed/. No raw history is deleted. Outputs are single-tenant and must not be combined as a multi-tenant model.
#### Legacy metadata conversion (manual only)

Current metadata uses `.json.txt`. Readers prefer it and reject malformed preferred
files rather than falling back to stale legacy copies. An unconverted legacy root
pointer prevents normal publication.
`Convert-SmartM365-WorkplaceEvidence-Metadata.cmd` previews validated legacy
renaming; `-ApplyMetadataConversion` preserves bytes, batch IDs and hashes,
publishing the root pointer last. Conflicting old/new names stop conversion.
Do not refresh/transfer during conversion or mass-rename unrelated files.
Current inputs need no conversion; installing code never runs it automatically.

Normal success or failure removes this run's source/staging copies and downloaded mappings, retaining capture/run audits, failure diagnostics and preparation logs. The shared JSON transport resolves these audit filenames, preferring `.json.txt`. An interrupted run's owned, marked source/staging payload is cleaned on the next run under the preparation locks. Unmarked legacy scratch folders, foreign-tenant folders and unrecognized batch folders are never automatically deleted. Cleanup refuses paths outside the declared root and symbolic links/junctions. A deletion failure is reported rather than silently presented as reclaimed disk space. Old batches without a verifiable publication receipt are retained with a warning. Small private logs/audits are not automatically purged; monitor disk space. No existing production folders are retroactively cleaned merely by installing this update.

Collectors may update raw sources after capture. Avoid overlapping prepared-output transfer/replacement with Power BI refresh: local atomic pointer replacement is not an atomic SharePoint transfer or a cross-query transaction. Verify transfer completion first; PreparedBatchId can pin a verified immutable batch, but a pin older than the two retained local batches is not supported by this retention policy. Publisher SHA256 checks and Power Query byte-length checks have different scope. This cleanup is local; it does not delete remote SharePoint batches.

Direct SharePoint transfer uses the shared Core uploader and downloader and requires the existing site read/write grant. When PreparedSharePointFolderPath is blank, the target is the normal SharePointTargetFolderPath after the shared CSV-to-DATA normalization, with `/DATA-POWERBI` appended. An explicit override must identify a library-relative DATA-POWERBI folder outside raw/log folders. Required configuration is checked before expensive generation. Missing destination folders are provisioned by the existing Core helper. No tenant URL or credential is hard-coded.

Transfer uses the same transfer implementation for normal enabled publication and TransferOnly recovery. It validates the current batch identity, all expected files, validation receipts and local SHA256/size first. It then uploads and downloads each CSV and the three batch metadata files, comparing SHA256 and size. The root current.json.txt is uploaded last, and also read back. For a 24-CSV batch, success means 28 files verified (24 CSVs, three batch metadata files, one root pointer). Reading back every file adds network traffic; no full duplicate batch is kept locally. At most one downloaded file is kept temporarily, under PreparedWorkRootPath/transfers/<run>/readback.tmp, removed after each check and on failure. A small private transfer.json audit remains. No historic CSV is recomputed or removed.

The local publication lock is held during transfer to block replacement/retirement by publishers sharing the same DATA-POWERBI root. It is not a distributed SharePoint lock: all writers for a tenant must use that shared root and avoid unrelated simultaneous cloud publication. Readers must wait for verified completion and synchronization; pointer-last publication is not an atomic OneDrive synchronization or a transaction across Power BI queries. An upload/read-back failure before pointer publication leaves the old root pointer untouched by this run. If pointer upload was attempted but its verification fails, activation is uncertain and the script explicitly reports failure; do not refresh until a successful retry. Retrying resends the same immutable files, with no recalculation or remote deletion.

Normal runs transfer when EnableSharePointUpload is enabled (template default true ; an existing runtime JSON keeps its local value). The scheduler entry alone does not change that setting. TransferOnly is a manual recovery mode, not a replacement preparation job. Qualify the actual remote transfer; synthetic transport tests do not establish live SharePoint success.

## Qualification boundary

Validate the deployed revision and specific batch. Source capture, business counts,
history retention, verified transfer, synchronization and Desktop/Service refresh
are separate evidence. Local or historical success does not qualify later changes.
Diagnostic outputs, tenant logs/manifests and mappings remain private.
