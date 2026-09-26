# Workplace evidence preparation

One post-collection job builds the 24 CSVs consumed by SmartWorkplaceIntelligence, including observed weekly history. It reads existing SmartInventory exports, never collects tenant data and never controls Power BI.

## Deployment through Git

Pull the approved commit into the existing repository. Keep SmartM365 and SmartWorkplaceIntelligence as sibling folders: the entry point depends on the latter's scripts and two JSON contracts. No ZIP installation is required. Tenant profiles, runtime JSON, mappings, CSVs, logs and PBIP files are not distributed by this change.

The job uses the same explicit tenant profile as other jobs. Missing per-script configuration is created from its template; missing keys are merged without overwriting existing values. Required runtime dependencies are PowerShell 7, ImportExcel, the existing SmartM365 Core/TenantContext and SmartInventory/Config/AccountClassification.psd1.

The effective LatestCsvFolderPath and DataAllRootPath must be sibling DATA-LAST and DATA-ALL directories. Their parent also contains the two private persona/site classification workbooks. Output is the sibling DATA-POWERBI folder. No developer-machine paths are hard-coded.

## Hourly orchestration

WorkplaceEvidence-Prepare is enabled in the job template at 00:00 through 23:00, every day (orchestrator local time), with MissedRunPolicy=Skip. An elected owner and the WorkplaceEvidence-Prepare concurrency key prevent overlapping instances. Required collector dependencies remain configured; ordering alone does not prove successful or complete collection.

Git updates the template, not an existing central runtime manifest. The active shared Orchestrator/Config/Orchestrator-Jobs.json must contain this same job; use the orchestrator management publication mechanism to preserve other jobs and record before/after versions. Its hot reload avoids restarting the scheduler. Do not replace the full runtime manifest with the template.

A run exceeding one hour must not overlap another occurrence. The 180-minute timeout is a safety limit, not a measured full-run duration. Source changes during capture reject the run; they are not silently mixed. Full production collection-machine qualification remains outstanding.

## Manual diagnostics

The Cloud launchers follow the existing prod convention:
- Test-SmartM365-WorkplaceEvidence-Prepare.cmd: ValidateOnly and Offline.
- Start-SmartM365-WorkplaceEvidence-Prepare-Offline.cmd: full local preparation.

ValidateOnly checks source existence, fingerprints and transport age; it does not generate CSVs or certify row-level tenant identity, collection success or business completeness. Offline suppresses upload and notifications, but a full offline run writes logs, snapshots, prepared CSVs and the validated local pointer.

The normal orchestrator job uses neither switch. Existing configured notifications apply. Direct SharePoint upload remains disabled by default and was not enabled by this change.

## Validation and history

The contract requires 47 current CSVs, two mapping workbooks, daily AD statistics and 12 historical source families. Missing required sources/history fail rather than become zeros. Current-file transport-age limit defaults to 168 hours, with 744 hours for license prices. These configurable limits are not a guarantee of business freshness or collector success.

Every present TenantKey must match the active tenant. Legacy tenantless CSVs are refused unless their origin has been reviewed and PreparedAllowLegacyTenantless explicitly enabled. No such exception is enabled by this release. Unexpected empty tables are rejected unless explicitly named in PreparedAllowEmptyTables.

Inputs are snapshotted into PreparedWorkRootPath outside synchronized DATA (default OS temp, separated by profile). All generators run sequentially. Validation covers output schemas, types, declared unique keys and loss of historical date/service/metric/country coverage compared with the last accepted batch. Workforce history is the current approved cohort viewed over observed snapshots, not an independently classified historic headcount census.

## Versioned output and refresh

DATA-POWERBI/current.json selects an immutable batches/<batch-id> directory containing 24 CSVs, batch.json, validation.json and a matching pointer copy. Only a fully validated batch advances the root pointer. Previous batches and failed diagnostics remain recoverable; no raw history is deleted. Outputs are single-tenant and must not be combined as a multi-tenant model.

Hourly operation retains scratch snapshots and output batches. Monitor free disk space; no automatic retention or cleanup is implemented in this preview. Retention policy needs separate approval rather than silently discarding history.

Do not overlap preparation, synchronization and Power BI refresh. Local atomic pointer replacement is not an atomic SharePoint transfer or a cross-query transaction. Verify transfer completion first; PreparedBatchId can pin a verified immutable batch. Publisher SHA256 checks and Power Query byte-length checks have different scope.

Optional direct SharePoint transfer uses the shared Core uploader, uploads batch files before the matching pointer and requires Sites.Selected with a site write grant. This path still needs remote-folder provisioning and read-back qualification; it has not been production-tested. The hourly activation does not enable it.

## Qualification boundary

Local prepared-file refresh and synthetic publication safety tests passed. The complete job still needs a successful post-collection run on the collecting machine, followed by log, source coverage, business-count and history checks. Do not equate a scheduler entry, Git publication or local tests with that production evidence. All output logs/manifests are private.
