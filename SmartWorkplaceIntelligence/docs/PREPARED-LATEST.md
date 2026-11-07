# Automatic selection of validated prepared batches

`PreparedBatchId` remains a required text parameter. `Latest` follows the
`BatchId` in `DATA-POWERBI/current.json.txt`; any other nonempty value pins an
immutable batch. Blank is an error, not an alias for Latest.

The publication contract is unchanged: the producer commits the root pointer
only after validation (`Passed=true`, `SchemaOnly=false`, no errors) and, on
SharePoint, after transfer and SHA256 read-back verification. Power BI trusts
that publication contract. No API-driven parameter switch, refresh lock,
manual approval, preparation job or retention change is introduced here.

## Reader and consistency boundary

`CreatePreparedReaders` creates one lexical batch context per evaluated
query. All readers within that query share its pointer, selected ID and
manifest. This includes licensing qualification/snapshots, mailbox joins,
weekly histories and the `Source-*` artifacts. The six derived migration
dimensions evaluate their source expressions within the same context.

The existing business readers and CSV transformations run before `Finish`
adds `Loaded Batch ID`. It is a hidden text column, not a key, unique column,
relationship, sort key or linguistic-schema field. Strict
`NativeEvidenceTypes` conversion and business deduplication never receive
the added column.

Both modes require the expected tenant, supported pointer/manifest schema,
matching manifest BatchId, unique manifest membership, unique file selection,
safe batch/file segments and the declared byte size. Metadata prefers
`.json.txt`; `.json` is considered only when the preferred file is absent,
never on authentication, read or parse failure. SHA256 remains a producer and
transfer check, not a new M computation.

Latest additionally compares the original pointer bytes with a pointer read
after each manifest/file read. A change, including whitespace only, raises
an error. Files are size-checked before that comparison; large payloads are
not buffered in their entirety. Small pointer/manifest binaries are buffered.
An explicit pin still validates the root pointer's identity but does not
follow its BatchId or fail because a newer batch was published.

SharePoint root-pointer reads use `Web.Contents` with the site parameter as
the base, `RelativePath`, and `IsRetry=true`, to bypass the HTTP response
cache. Library discovery matches the configured navigation name against the REST title
or root-folder name and requires exactly one match. Incomplete discovery fails
closed; navigation names are not assumed to be REST titles. HTTP 404 alone permits
the legacy pointer filename. Immutable batch navigation remains on
`SharePoint.Contents`. A Service connection for the Web source at the same
site may therefore need organizational-account credentials as well as the
SharePoint connection. See Microsoft's [Web.Contents reference](https://learn.microsoft.com/powerquery-m/web-contents).

These are per-read checks, not a transaction across all Power Query tables.
If publication falls between two otherwise stable query evaluations, the
refresh can contain more than one batch. A failed ordinary full Import
refresh should retain the previously committed model; this design does not
provide that guarantee for custom partial-commit refresh APIs.

## Visible diagnostic

The Data Trust KPI strip has a persistent dynamic title:

`Mixed batches | Majority: <BatchId> | Other tables: <count> | Empty tables (batch unobservable): <count>`

With one observed ID it displays `Loaded batch: <BatchId>`. With no observed
ID it displays `Loaded batch: Not observed`. Missing provenance on a
nonempty table is reported separately.

The hidden calculated `Loaded Batch Audit` table reads the already imported
columns: one row per monitored query, its BatchId and its row count. It does
not execute Power Query or recalculate CSVs. Majority means number of
nonempty tables, not number of business rows. Ties use the lexically smallest
ID. The audit has no relationships and the diagnostic ignores report filters.
Its rows identify individual affected tables when inspecting the model.

**Empty tables cannot expose their batch through a row column.** They are
counted explicitly and excluded from the majority calculation. A single
observed BatchId therefore does not prove which batch an empty table read.
Static reference tables and calculated business tables are not independent
prepared readers and are excluded. A stable subsequent full scheduled
refresh can restore one observed batch without manual intervention.

## Deployment order (not executed by these changes)

1. Back up and integrate the updated model and Data Trust visual definitions
   into the private PBIP, retaining its current explicit BatchId, private
   source/tenant parameters, editor settings and existing report work.
   Include the readers, provenance columns/partitions, calculated audit and
   banner measure together. For the SharePoint-only Service project use
   `powerquery/PreparedReadFolder.SharePoint.pq` as `PreparedReadFolder`;
   do not copy the checkout's Local/SharePoint transport into that project.
   Keep all other updated expressions unchanged.
2. Republish that integrated project to the intended existing Service model
   and report. Configure the Web/SharePoint organizational-account
   connections if prompted. Definitions must be present before switching
   the parameter; `Latest` is not compatible with the old pin-only reader.
3. Only then set `PreparedBatchId` to `Latest` in the Service **and** in the
   private PBIP, and save the private project. Keep
   `IsParameterQueryRequired=true` and the expected tenant unchanged.
   The checkout template already defaults to Latest.
4. Set the principal Service refresh to **12:30, Europe/Paris**, with an
   optional second refresh around **18:00**. Select the Paris time zone so
   daylight-saving changes are handled by the schedule. Preparation starts
   from a 07:45 schedule but can wait for FreshSuccess dependencies. Observed
   runs ranged from 53 minutes to 2 h 29 and finished as late as 11:03;
   09:30 is not the proposed schedule. 12:30 provides margin, not a guarantee
   against a longer or delayed future preparation.

An explicit batch ID remains available for a deliberate fixed snapshot.
Pinning does not itself add retention protection. Existing retained and
protected batches are unchanged.

## Validation and review scope

All automated fixtures are synthetic. PQTest executes the actual reader
expressions for pointer selection, explicit pin precedence, pointer changes
after manifests/payloads, Source-* and mailbox reads, byte-identical checks,
tenant/schema/batch/size errors, missing/ambiguous pointers, absence-only
legacy fallback, strict native types, mixed results and empty tables.
The existing licensing and mailbox differential harnesses keep comparing
business columns; the new provenance field is tested separately.

The DAX review covered 1,061 pre-existing measures, 184 calculated columns and
36 calculated tables, including expressions beyond the reference validator's
parsed subset. Targeted whole-table searches found:

- `ALL` in Windows/mailbox availability guards, mailbox exception summaries,
  licensing SKU selection and collaboration coverage. These count rows,
  filter explicit business columns or project SKU columns before comparison.
- `ALLSELECTED` in native detail gates, application ranking and cumulative
  migration dates. Grouping/ranking still names explicit business columns.
- Nine `SUMMARIZE` uses on Endpoint Protection, all explicitly grouping
  tenant and device identifiers.
- No direct `DISTINCT(table)` or `VALUES(table)` on a monitored table, no
  whole-table row comparison or set operation dependent on column arity.
- Calculated business tables project explicit columns (`Daily Domain`,
  `Country Selection`, `Workforce License Review Selection`) or use static
  catalogs/calendars. The added column does not enter those projections.

The new audit deliberately inspects the technical field. All existing
business DAX, keys and relationships remain unchanged. Run:

```text
python -B -m unittest test_prepared_latest_m test_license_report_m test_native_evidence_m test_mailbox_migration_m -v
python -B test_model_dax_references.py
```

Set `PQTEST_PATH` to the installed Microsoft Power Query SDK executable.
Also run `Test-ConsolidatedReport.mjs` with a Modeling MCP metadata export,
PBIR validation, offline TMDL loading, `git diff --check`, and the repository
ASCII/privacy checks. Offline model loading and synthetic tests do not prove
the deployed Service credentials, refresh behavior or rendered banner.
No Desktop/Service refresh or publication is part of this implementation.
