# Workplace evidence preparation

One post-collection job builds the 24 CSVs consumed by SmartWorkplaceIntelligence, including observed weekly history. It reads existing SmartInventory exports, never collects tenant data and never controls Power BI.

## Deployment through Git

Pull the approved commit into the existing repository. Keep SmartM365 and SmartWorkplaceIntelligence as sibling folders: the entry point depends on the latter's scripts and two JSON contracts. No ZIP installation is required. Tenant profiles, runtime JSON, mappings, CSVs, logs and PBIP files are not distributed by this change.

The job uses the same explicit tenant profile as other jobs. Missing per-script configuration is created from its template; missing keys are merged without overwriting existing values. Required runtime dependencies are PowerShell 7, ImportExcel, the existing SmartM365 Core/TenantContext and SmartInventory/Config/AccountClassification.psd1.

The effective LatestCsvFolderPath and DataAllRootPath must be sibling DATA-LAST and DATA-ALL directories. Output is the sibling DATA-POWERBI folder. No developer-machine paths are hard-coded.

Normal and ValidateOnly runs download the private persona/site classification workbooks from SharePoint on every run, using the existing tenant site, library, app and certificate configuration. PreparedMappingSharePointFolderPath optionally overrides the source folder; otherwise SharePointTargetFolderPath is used, with the shared CSV-to-DATA normalization. The two standard workbook filenames are relative to that folder. No client URL or classification data is committed.

Both downloads must succeed and contain the expected worksheets before source planning begins. They are isolated under a unique scratch directory and copied into the run snapshot with hashes; no raw-root workbook is overwritten, and a failed download never silently falls back to cached data. Structural workbook checks do not substitute for later classification-rule validation.

## Hourly orchestration

WorkplaceEvidence-Prepare is enabled in the job template at 00:00 through 23:00, every day (orchestrator local time), with MissedRunPolicy=Skip. An elected owner and the WorkplaceEvidence-Prepare concurrency key prevent overlapping instances. Required collector dependencies remain configured; ordering alone does not prove successful or complete collection.

The elected owner must advertise SharedRuntime and Graph with Sites.Selected for the mapping downloads. This eligibility requirement does not grant permissions; the existing site-level grant must already allow the two reads.

Git updates the template, not an existing central runtime manifest. The active shared Orchestrator/Config/Orchestrator-Jobs.json must contain this same job; use the orchestrator management publication mechanism to preserve other jobs and record before/after versions. Its hot reload avoids restarting the scheduler. Do not replace the full runtime manifest with the template.

A run exceeding one hour must not overlap another occurrence. The 180-minute timeout is a safety limit, not a measured full-run duration. Each source is captured and verified independently, with at most three attempts for that file (one-second delay). An unstable/unreadable file still rejects the run. Full production collection-machine qualification remains outstanding.

## Manual diagnostics

The Cloud launchers follow the existing prod convention:
- Test-SmartM365-WorkplaceEvidence-Prepare.cmd: ValidateOnly, with SharePoint mapping reads.
- Start-SmartM365-WorkplaceEvidence-Prepare.cmd: full preparation with SharePoint mapping reads.
- Start-SmartM365-WorkplaceEvidence-Prepare-Offline.cmd: full local preparation.

ValidateOnly downloads/structurally checks the two workbooks, captures the required sources and checks transport age, CSV record shape and every row TenantKey on the verified copies. It aggregates validation failures with relative paths, including history. Later live-source changes do not invalidate these copies. It writes temporary source copies and small private audit files but does not generate prepared CSVs, upload or notify, nor certify collection success or business completeness. Offline suppresses all SharePoint access and notifications: it requires the two workbooks already present in the tenant data root and never downloads missing files. A full offline run still writes logs, temporary snapshots, prepared CSVs and the validated local pointer.

The normal orchestrator job uses neither switch. Existing configured notifications apply. SharePoint mapping reads are independent of EnableSharePointUpload; direct output upload remains disabled by default. Read access requires the existing app's Sites.Selected role and a read (or existing write) grant on the selected site. No permission or tenant grant is changed by this code.

## Validation and history

### Manual Workforce memory diagnostic (v0.1.7)

Run Test-SmartM365-WorkplaceEvidence-WorkforceMemory.cmd on the collecting machine to execute only the unchanged Workforce calculations, including all available history. Do not add this mode to the hourly orchestrator. It uses the same prod profile/configuration and downloads the two mapping workbooks from SharePoint unless explicitly Offline. No raw collector, DATA-POWERBI publication, Power BI operation, output upload or notification is performed. Diagnostic and normal preparation runs share the existing preparation locks.

The child process runs New-WorkforceIdentityEvidence.ps1 directly against existing raw exports (read-only, not an atomic snapshot). All five output paths are redirected to a unique private PreparedWorkRootPath/workforce-diagnostics/<run>/outputs folder; no production CSV is overwritten. There is no extra full source copy and no automatic cancellation or memory-setting change. Avoid a simultaneous normal preparation run, but leave other machine activity representative when investigating global memory pressure.

The parent process samples worker private bytes/working set and system available physical memory, committed bytes and commit limit approximately every five seconds. The worker writes opt-in stage boundaries with managed-memory measurements without forced garbage collection. Stages cover individual imports, indexes, identities, personas, exports and each historical week. The parent captures stdout/stderr and exit code even when the worker exits abnormally. A whole-machine crash can still prevent final reporting. Missing system counters are recorded as unavailable, not zero; the sampled private-byte maximum may miss peaks between samples.

Keep memory.csv, stages.ndjson, environment.json and result.json for diagnosis. Worker stdout/stderr are separate private logs and can include file paths/error details. Diagnostic outputs contain client data and must never be committed or uploaded publicly. Diagnostic artifacts are retained intentionally for inspection, not swept by normal run cleanup. An isolated success does not prove production publication or output equivalence. The previous locally qualified recovery batch reused verified history; it was not an identical full fresh run. No column reduction, account exclusion, history truncation or business-rule optimization is introduced by this instrumentation.

The contract requires 47 current CSVs, two mapping workbooks, daily AD statistics and 12 historical source families. Missing required sources/history fail rather than become zeros. Current-file transport-age limit defaults to 168 hours, with 744 hours for license prices. These configurable limits are not a guarantee of business freshness or collector success.

Every source CSV must contain TenantKey and all rows must match the active tenant, including history. PreparedAllowLegacyTenantless must remain false: the global bypass is rejected as of v0.1.4. There is no exception registry. Unexpected empty output tables are rejected unless explicitly named in PreparedAllowEmptyTables.

As of v0.1.5, source validation uses semicolons for Exchange_OnPrem_Servers_Inventory.csv as well as daily statistics, matching the evidence generators. Other source CSVs use commas. The same strict record-shape and tenant checks apply to both formats; malformed files are not retried with a different separator. No source CSV conversion is needed.

### One-time repair of reviewed historical exports

Older weekly CSVs can predate the tenant column. Only after the operator has confirmed their ownership, use Repair-SmartM365-WorkplaceEvidence-TenantKeys.cmd -RepairWeeks "YYYY-Www,YYYY-Www" -ExpectedRepairFileCount N. This is a preview by default. Add -ApplyRepair to perform the approved repair. Both week values and the exact expected count must be supplied; nothing client-specific is committed. The tenant value comes from the same effective profile as other launchers (prod for the Cloud launcher). Do not put repair arguments in the orchestrator.

The tool selects only historical families in the source contract, in the exact named weeks. It does not touch DATA-LAST, daily statistics or files outside those weeks. Existing TenantKey values are checked, never overwritten. Malformed rows, a conflicting tenant or a different nonzero candidate count stop the repair. A repeat run with no missing columns reports NoMissingTenantKey. Stop collectors/synchronization and avoid Power BI refresh during maintenance; the preparation and repair modes share a data-root lock, but other raw collectors do not use it.

Before any source replacement, all originals are copied and hash-verified under DATA-REPAIR-BACKUPS/<run>/originals, outside the folders read by preparation/Power BI. Repaired copies add only TenantKey as the first column; every original parsed field, header, record order and row count is compared exactly, including whitespace and multiline fields. CSV serialization is normalized to quoted UTF-8 BOM/CRLF, not byte-identical; original bytes remain backed up. Original LastWriteTimeUtc is restored to avoid misrepresenting old data as a new collection.

Source replacements are atomic per file, not one transaction across the whole set. A failure can leave a partially repaired set: stop and inspect the private repair.json journal and verified backups before resuming. There is no automatic deletion or rollback that could overwrite concurrent changes. The journal records paths, original/repaired SHA256, row counts, operator, time and per-file status; keep backups and journal private. The added tenant identity is operator-attested, not reconstructed historical collection evidence.

Repair mode makes no SharePoint calls, sends no notifications, does not generate prepared CSVs and never changes DATA-POWERBI/current.json. After successful repair, run the normal Test launcher; only a passing complete preflight should be followed by full Start. The hourly job never repairs sources automatically.

Inputs are snapshotted into PreparedWorkRootPath outside synchronized DATA (default OS temp, separated by profile). Since v0.1.6, the source list is selected once, without an expensive initial global hash scan. Each file is streamed into its copy without blocking collector writes, then the copied bytes are checked against a fresh source SHA256, local-copy SHA256, size and modification timestamp. A failed check retries only that file. A newer file than the initial plan is allowed; provenance records the actual captured version, time and attempt count. New historical files arriving after selection are included on the next run. There is no final comparison to the live source tree and no claim of an atomic, collector-wide snapshot. All validators and generators use only the verified copies; source freshness, tenant, schema and historical-coverage gates remain enforced.

All generators run sequentially. Validation covers output schemas, types, declared unique keys and loss of historical date/service/metric/country coverage compared with the last accepted batch. Workforce history is the current approved cohort viewed over observed snapshots, not an independently classified historic headcount census.

## Versioned output and refresh

DATA-POWERBI/current.json.txt selects an immutable batches/<batch-id> directory containing 24 CSVs, batch.json.txt, validation.json.txt and a matching current.json.txt pointer copy. The content remains JSON; only the transport filename changes. Only a fully validated batch advances the root pointer. The current and immediately previous valid batches remain recoverable; older published CSV batches are retired only after a successful replacement. Their small manifests/validation receipts remain under retired/. Failed publication payloads are removed while diagnostics remain under failed/. No raw history is deleted. Outputs are single-tenant and must not be combined as a multi-tenant model.

### Metadata filename migration (v0.1.9)

Use an administrator-approved `.json.txt` extension where synchronization excludes `.json`. Update the Power BI `GetPreparedEvidenceFile` expression from the generic `SmartWorkplaceIntelligence/scripts/GetPreparedEvidenceFile.pq` reader before migration. It prefers `.json.txt`, accepting legacy `.json` only when the preferred filename is absent. A present but unreadable or malformed new file is an error, not a reason to use stale legacy metadata. Local and SharePoint modes use the same convention. No tenant settings are included in the generic reader.

For an already-generated local batch, run `Convert-SmartM365-WorkplaceEvidence-Metadata.cmd` to preview; add `-ApplyMetadataConversion` to rename. Run this once before the next normal preparation. The launcher uses the existing prod profile and its sibling DATA-POWERBI directory, without SharePoint calls, workbook downloads or CSV recalculation. Do not refresh Power BI or transfer these files during migration; verify synchronization completes afterward.

The conversion validates tenant identity, publication receipts and CSV SHA256/size against the manifests, including retained previous batches. It renames metadata without changing bytes, batch IDs or hashes; the root pointer is renamed last. For one published batch this is four files. Publication locking prevents overlapping prepared publication. A partially completed rename can resume; conflicting old/new names stop conversion. Renames are atomic per file, not a transaction across the entire set. Older retired audit folders, failed diagnostics and unrelated JSON configuration files are not migrated. Private scratch diagnostics continue using JSON. New published/retired/failure metadata uses `.json.txt`.

Normal preparation rejects an unconverted legacy root pointer before doing expensive calculations, preventing two different current pointers from coexisting. The publisher also enforces that guard. Existing legacy batch receipts remain readable for retention. Installing this update alone does not rename production files or run preparation. Cloud publication, if enabled, now checks required configuration before calculation; metadata migration itself does not transfer the existing batch to SharePoint. Qualify the actual remote upload separately.

Normal success or failure removes this run's source/staging copies and downloaded mappings, retaining capture.json, run.json, failure diagnostics and preparation logs. An interrupted v0.1.6+ run's marked source/staging payload is cleaned on the next run under the preparation locks. Unmarked legacy scratch folders, foreign-tenant folders and unrecognized batch folders are never automatically deleted. Cleanup refuses paths outside the declared root and symbolic links/junctions. A deletion failure is reported rather than silently presented as reclaimed disk space. Old batches without a verifiable publication receipt are retained with a warning. Small private logs/audits are not automatically purged; monitor disk space. No existing production folders are retroactively cleaned merely by installing this update.

Collectors may update raw sources after capture. Avoid overlapping prepared-output transfer/replacement with Power BI refresh: local atomic pointer replacement is not an atomic SharePoint transfer or a cross-query transaction. Verify transfer completion first; PreparedBatchId can pin a verified immutable batch, but a pin older than the two retained local batches is not supported by this retention policy. Publisher SHA256 checks and Power Query byte-length checks have different scope. This cleanup is local; it does not delete remote SharePoint batches.

Optional direct SharePoint transfer uses the shared Core uploader, uploads batch files before the matching pointer and requires Sites.Selected with a site write grant. This path still needs remote-folder provisioning and read-back qualification; it has not been production-tested. The hourly activation does not enable it.

## Qualification boundary

Version 0.1.8 fixes fresh Workforce history initialization: absent, empty and single-row staging files remain arrays under strict mode. Existing rows and the obsolete-schema rejection are preserved. Synthetic regression tests exercise the actual initialization statements; a complete production run is still required.

Local prepared-file refresh and synthetic publication safety tests passed. The complete job still needs a successful post-collection run on the collecting machine, followed by log, source coverage, business-count and history checks. Do not equate a scheduler entry, Git publication or local tests with that production evidence. All output logs/manifests are private.
