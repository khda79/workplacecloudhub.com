# Migration to SharePoint Online

Optional reporting over a private unified workbook, on-prem inventory receipts
and prepared SPO usage-site observations. No migration, collection, deletion or
Power BI refresh. Disabled by default. Deploy matching SmartM365 and
SmartWorkplaceIntelligence sibling folders; update the preparation wrapper,
template, pipeline, dispatcher, migration module, Python reader and contract
together. PowerShell 7 and Python 3.10+ standard library are required.

## Private configuration and acquisition

The public `SmartWorkplaceIntelligence-Migration-to-SPO.xlsx` template has no
tenant data. The reader supports schema 6.3 and reads tblMigrationSP,
tblMigrationSharedFolders, tblLegacyWebs, tblLegacyCollections and Parameters.
Keep the current columns, including the five hidden imported SharePoint fields.
Overview totals are not facts. Derived Excel formula caches are ignored.

Set these keys only in the private preparation runtime configuration:

| Setting | Meaning |
|---|---|
| PreparedSharePointMigrationEnabled | true to include private planning data |
| PreparedSharePointMigrationWorkbookFileName | Actual private workbook basename |
| PreparedSharePointMigrationBindingsFileName | Optional private registry basename ending .json.txt; required when configured |
| PreparedSharePointMigrationTargetTenantUrl | Private SPO origin used to resolve relative targets |
| PreparedSharePointMigrationReviewMaxAgeDays | Business review warning threshold; default 7 |
| PreparedSharePointMigrationReportingTimeZoneId | Explicit host TimeZoneInfo calendar; default UTC |

Preserve authentication, mapping folder, retention and source-root settings.
LatestCsvFolderPath ends in DATA-LAST and DataAllRootPath is its sibling DATA-ALL.
Their parent is the tenant data root for retained workbook and registry copies.
Normal acquisition downloads the enabled inputs from the existing SharePoint
mapping folder alongside the classification workbooks. The entire downloaded
set is validated before retained copies are replaced. Each run then captures
its own stable, fingerprinted inputs. Missing, invalid or cross-tenant inputs
never fall back silently to retained files. Offline requires explicit local
inputs; ValidateOnly checks/captures without publishing a prepared batch;
TransferOnly does not download inputs or regenerate CSVs. A full normal run can
transfer its batch according to EnableSharePointUpload, independently of report
publication or refresh.

## Identity, scope and approval

Operation keys hash tenant, project, canonical source-type code and MigrationID;
IDs may repeat between flows. One source with distinct destinations remains
distinct operations. URL identity ignores HTTP/HTTPS and default ports while
retaining meaningful ports, encoded separators and resource queries. Only
explicitly approved registry host pairs or dated exact alias bindings equate
domains. Compare exact normalized path segments, preserving singular/plural
distinctions. Native tenant/farm/collection/web keys are never merged.

A runtime registry requires the exact workbook Identity. Alias Bindings retain
dated Confirmation evidence. Separate TargetBindings contain OperationKey,
SourceReference, TargetURL, ParentMigrationID, EvidenceSHA256 and ConfirmedAtUtc.
A changed source, parent mapping or destination invalidates that match. Raw
TargetReviewStatus is never promoted automatically. Site-level parent approval
does not verify libraries, folders or contents or create migration validation.

Content and infrastructure receipts must match tenant, producer, run, payload
hashes, row counts and acquisition age, including per-file overrides. Completed
with configured outputs alone is not exhaustive coverage. Non-observation is
not deletion. Explicit DELETE/DONE remains declared completion; still-observed
sources generate review findings. Shared-folder paths are checked syntactically,
without claiming SMB existence, volume or effective permissions.

SPO usage sites provide site observations, not an exhaustive existence census.
Relative targets retain their original text. Destinations outside the configured
SPO tenant are not qualified; non-observation is not proof of migration failure.
Child scope follows native collection boundaries and the closest explicit
declaration. Partial content does not cover an entire web. Newly observed webs
without whole-web coverage appear NOTINPLAN without rewriting Excel.

## Outputs and indicators

The nine tables in `sharepoint-migration-contract.json.txt` cover current
operations, native webs/collections, two legacy projections, operation/web scope,
findings, history and source qualification. They belong to the same immutable
batch. Disabled runs emit typed empty current tables and Not configured while
retaining existing history.

Reported completion is MIGRATE/DONE divided by all MIGRATE operations. Validated
completion uses that denominator plus VALIDATED with a coherent, non-future
validation date. Remaining execution excludes DONE/CANCELLED; cancelled work is
shown separately. Other dispositions have separate action tracking. No population
means unavailable, not zero percent. Operations, webs, collections and storage
retain their own grains. Storage is counted once per collection; migrating volume
requires complete qualified scope without exclusions. Copied bytes and effective
ACLs are not inferred.

Observed status history starts at the first validated snapshot, preserves past
attributes and records scope changes/disappearance. ActualDate is a declared
completion calendar, not reconstructed observed progress. The previous selected
batch history is hash-verified and captured before the next build; changed
predecessors block promotion. Disabling the feature retains its history. Source
archives are not recurring refresh dependencies.

All names, code, docs, templates, examples and test fixtures are generic. Actual
workbooks, aliases, tenant/project IDs, URLs, paths, contacts, connection values,
cached data and screenshots stay private. Synthetic test success does not prove
production, Desktop rendering, content migration or Service behavior.

Offline checks: `test_spo_migration.py`, `Test-SharePointMigrationEvidence.ps1`,
mapping acquisition, scheduler and prepared-batch regression suites.

## Report experience

The single Migration to SharePoint Online page follows On-prem Content & Storage.
Progress, Schedule, Reconciliation and Quality are native bookmark views sharing
source type, project, business unit and phase slicers. Switching views preserves
filters. Clear slicers keeps the active view; filters in the native filter pane
have their own clear control. Schedule has separate deadline and observed-history
modes. Planned dates are grouped by Monday-starting week. Historical observations
use markers without connecting unmeasured days. ActualDate has its own declared
completion calendar. Snapshot dates never filter the current plan.

Current and historical business attributes have separate fact relationships.
Technical on-prem filtering uses explicit scope passage for operation measures;
shared-folder operations remain included with a Not applicable technical context.
The global webs-without-plan counter ignores business unit and phase because
those webs have no attributable planning row. Source/global findings also remain
visible separately from operation-level findings. Storage stays at collection
grain and is available in reconciliation tooltips with its qualification.

`migration_model_definitions.py` provides generic TOM definitions for Modeling
MCP. `Build-SharePointMigrationPage.js` generates only this page and its bookmarks;
it accepts an explicit PBIP root and a SMARTWORKPLACE_PYTHON executable. It does
not refresh, launch Desktop or publish. Additional offline checks are
`test_spo_migration_model.py` and `test_spo_migration_page.py`. Static acceptance
does not qualify DAX execution, interactive filtering or Desktop rendering;
those checks require a separately authorized refresh of a reviewed batch.
