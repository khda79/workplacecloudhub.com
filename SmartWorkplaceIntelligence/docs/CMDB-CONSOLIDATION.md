# CMDB consolidation into Workplace Intelligence

## Delivery boundary

The consolidation is an integration candidate, not a production-qualified release. It does not execute collectors, preparation, SharePoint transfers, Desktop refresh or Desktop save. The former CMDB runtime and job remain available until the server and private Desktop gates below pass. Do not run both preparation jobs during cutover.

One owner now captures the inputs, runs the Estate/native/domain transformations, validates the composed output contract, publishes one immutable `DATA-POWERBI` batch and transfers that same batch. Native transforms have no standalone publisher or pointer. Native source populations are preserved where they differ from enabled-workforce or managed-device populations; they are not silently replaced by those narrower populations.

## Recovery coverage

- `config/consolidation-recovery-map.json.txt` maps all 46 native tables, 696 source columns, seven model-calculated columns, 249 measures, 80 relationship dispositions and 13 former pages.
- `config/consolidation-report-facets.json.txt` assigns complete native details to existing pages. Group license paths are also exposed on CMDB.
- The 27 private Intelligence pages are retained; CMDB is page 3, immediately after Operational Overview. Public definitions are synchronized without private data, caches or tenant parameters.
- Native Date values extend the shared Date range. Native Quarter and Month numbers map to separate numeric attributes; existing Quarter and Month labels retain their meaning.
- Original native source/calendar fields remain inspectable. Automatic date hierarchies are replaced by shared Date attributes, not copied as another autonomous model.
- The preexisting CMDB publisher fix is preserved conceptually: zero exit is not verified publication. The unified transfer returns a complete batch-specific receipt after every SHA256/size read-back and the final pointer read-back; the entry point asserts this receipt before reporting success. The original four modified CMDB files remain untouched for the later retirement gate.

## Identity and context

Entity keys include tenant, native system, entity type and native identifier. Names are display values, never identity joins. Existing canonical Estate decisions are reused through exact GUID/managed-device identifiers. AD and Entra identities remain distinguishable.

Observed direct, primary-group and nested membership edges remain separate. Unresolved endpoints remain in the link table; no missing parent or transitive membership is fabricated. User-device association is not console activity. Group-license routes are observations, not effective membership or guaranteed entitlement.

Incoming/outgoing relationship selection uses an explicit endpoint projection measure and a visual filter, with one-directional relationships. Native detail and metrics use explicit native-key projection. A synchronized single-entity selection takes precedence; otherwise existing qualified user/device inventory filters apply where a native key/path exists. Unrelated entity detail returns blank, not the whole tenant. Tenant-level source qualification remains explicitly tenant-scoped.

## Unified capture and batch

`native-evidence-contract.json.txt` retains receipt-based acquisition freshness, required schemas, identities and native key relationships. `native-output-contract.json.txt` adds 46 native outputs and four entity/link/endpoint/alias outputs to the existing Intelligence contract.

`prepared-artifact-contract.json.txt` captures the published license-report snapshot and CSVs, SharePoint on-prem current exports, migration current evidence and existing weekly histories. These artifacts are copied from the stable capture, not reread from changing live folders during publication. Artifacts have strict names, sizes and hashes and an explicit history index in the same immutable batch.

The transfer validates the complete composed contract, all tables, all artifacts and metadata before upload. It reads back and hashes every uploaded object; the root pointer is last. A failure after pointer upload is reported as unconfirmed, never success. Existing immutable pre-consolidation batches remain transferable under their original contracts.

## Reader and integrity boundary

All runtime model readers use `GetPreparedEvidenceFile`. Direct license, on-prem SharePoint and migration reads have been replaced by artifacts in the pinned batch. Static neutral reference catalogs remain independent.

`PreparedExpectedTenantKey` and `PreparedBatchId` must be set explicitly after the first unified server batch is validated. Public source defaults are blank and fail closed. The private tenant parameter is preserved; its batch parameter remains intentionally unpinned until that first validated cohort. Opening the project is not permission to refresh it.

SHA256 verification is performed by preparation, read-only batch validation and transfer. Power Query verifies the expected tenant, manifest membership, expected size, fixed batch and stable bytes; it does **not** independently recompute SHA256. Do not describe the M reader alone as cryptographic verification of every file.

## Offline evidence versus remaining gates

Offline tests compare all 46 native CSVs against the former implementation on identical synthetic capture inputs, including all source columns. Additional cases cover wrong identity, stale acquisition, changed source, duplicate/unknown identities, unresolved and nested members, direct/primary groups, manager links, strong Estate links, interruption and no independent publication.

Synthetic composed-batch tests cover publication, immutable artifacts, full transfer read-back, pointer ordering, corrupt artifacts, wrong tenant, partial receipts and preservation of the prior pointer. Existing capture, history, transfer and lifecycle tests remain required.

PBIR structural validation and MCP TMDL export/reimport are offline checks only. They do not qualify DAX execution, actual filter behavior, memory consumption, performance or Desktop rendering on the tenant batch.

Obsolete synthetic model tables have no report consumers. Their bulk retirement was blocked by the safety reviewer; 49 remain hidden, with only their required private automatic-calendar links restored to keep serialization valid. No production visual is bound to them. This outstanding cleanup must not be presented as completed.

### Reproducible offline checks

Run from the repository root in PowerShell 7. These suites create synthetic temporary fixtures; their simulated transfer callbacks never connect to SharePoint.

```powershell
pwsh -NoProfile -File SmartWorkplaceIntelligence/scripts/Test-UnifiedSourcePlan.ps1
pwsh -NoProfile -File SmartWorkplaceIntelligence/scripts/Test-PreparedArtifacts.ps1
pwsh -NoProfile -File SmartWorkplaceIntelligence/scripts/Test-UnifiedPreparedTransfer.ps1
Push-Location SmartWorkplaceIntelligence/scripts
try { python -B -m unittest test_native_evidence -q } finally { Pop-Location }
```

The targeted results are 126 source-selection checks, 35 artifact checks, 130 composed publication/transfer checks and 18 native migration/graph test cases. The full native suite passes all 151 tests (603 seconds in the final workstation run); it includes the former preparation regression suite, intentionally retained until retirement ports its synthetic fixtures and golden comparisons. The 135 existing endpoint-transfer checks pass with the composed contract changes. Other existing capture, history, lifecycle and metadata suites remain part of the regression gate; the metadata suite passes 43 checks.

`Test-ConsolidatedReport.mjs` requires the report folder and an explicit JSON model-metadata projection obtained through the modeling MCP. It checks 8,714 bindings and structural invariants on each public/private report. The final definitions reimport as 203 tables, 982 measures and 99 relationships. These counts include the 49 hidden legacy tables pending cleanup; they are not a claim of 203 governed production sources.

PBIR validation reports no structural errors. Five preexisting drillthrough pages still lack a visible Back button; offline schema reachability can add a schema-unavailable warning. These warnings and actual rendering/filter behavior must be checked during the Desktop gate, not silently waived by static success.

## Server cutover — operator gate, not executed by this change

1. Back up and hash the effective server orchestrator/configuration and both prepared roots outside the clone. The verified workstation backup does not replace this server backup.
2. Confirm no preparation is running. Update the reviewed integration commit using GitHub Desktop. Verify the scoped PowerShell signatures and Python 3.10+ availability; configure `PreparedPythonCommand` if necessary.
3. Inspect the effective planning revision. Disable the old `CmdbEvidence-Prepare` job through the orchestrator's existing validated publication mechanism, preserving the effective Workplace schedule (08:30 where currently configured), all collector schedules and global maintenance state. The template is not evidence of the effective server configuration.
4. Keep `WorkplaceEvidence-Prepare` and the union of its collector dependencies. The template adds Workplace Scope, Verified Domains, Windows Autopilot and Autopatch Alerts. Do not weaken freshness or change collector schedules to make validation pass.
5. Run the unified source validation, then the first unified preparation and transfer only at the approved operational gate. Inspect the new manifest, native tables, artifacts, qualification, hashes and SharePoint root pointer. Never substitute a successful offline fixture for this evidence.
6. Pin the private model to this exact tenant and batch. Refresh and inspect all 28 pages only with the explicit Desktop execution/save boundary agreed with the user. Keep the former saved cache until that check succeeds.
7. Only after server and Desktop acceptance: remove the former CMDB product, scripts/launchers/job references and `DATA-POWERBI-CMDB` from active local/server/SharePoint use, using the verified external backups. Preserve the useful collectors and shared receipts. Publish retirement separately from integration.

## Rollback

Until acceptance, retain both the verified external project/data backup and the previous immutable Intelligence batch. Revert the integration definitions and restore the former effective planning revision only after confirming no preparation is running. Restore private parameters and the saved project/cache from the verified backup; do not overwrite a newer unrelated report edit. No broad Temp/DATA cleanup is part of rollback.
