# Smart Workplace Intelligence

Power BI reporting over governed SmartM365 workplace evidence and observed
histories. Domain pages cover workforce, devices, applications, licensing,
collaboration, messaging, and source quality. Native details retain their source
identity and grain; reported configuration and status are qualified evidence,
not proof of effective on-device behavior.

## Data flow

The optional [Migration to SharePoint Online](docs/SPO-MIGRATION.md) feature
combines private SharePoint on-prem and shared-folder planning with qualified
inventory observations. Public templates contain no tenant data.

The optional [mailbox migration exception register](docs/MAILBOX-MIGRATION-EXCEPTIONS.md)
adds a private English Excel input, confirmed native identity matching, and dated
planning cards. It is disabled by default and does not rewrite historical trends.

1. SmartM365 SmartInventory owns collection and publishes tenant-scoped source
   CSVs, receipts, and histories.
2. `SmartM365-WorkplaceEvidence-Prepare.ps1` calls the preparation components in
   this project's `scripts` and validates an immutable `DATA-POWERBI` batch.
3. A normal run transfers the validated batch and pointer when
   `EnableSharePointUpload` is enabled. Mapping downloads and run-log publication
   have separate behavior; consult the effective deployment configuration before
   choosing a run mode.
4. Power BI consumes the qualified batch. Report refresh and save are separate
   operations from collection and preparation.

The [Exchange migration licensing scenario](docs/EXCHANGE-LICENSING.md) documents
identity reconciliation, retained SE exceptions and estimated E3/F3 purchases.

## Thematic report generation

`scripts/generate_thematic_reports.py` 1.2.4 creates report-only derivatives
from a saved master. It never refreshes, publishes, edits the master or copies
its semantic model, cache or local settings. Without `--write` it only prints
a JSON plan; writing requires a new output folder. Keep private output outside
the repository, with the current manual-publication bundle under
`publication/thematic-reports`; do not recreate a second current `delivery` set.

The default manifest is selected from the source's stable page IDs:

- `config/thematic-reports.json.txt` retains the eleven-theme profile for the
  existing 30-page source, including its reviewed legacy layout adjustments.
- `config/thematic-reports-pc-mobile.json.txt` supports a 32-page source with
  both Mobile pages. It inherits that source's layout without legacy positional
  adapters. A partial Mobile pair is refused before writing. `--manifest` can
  explicitly select a reviewed configuration; exact page coverage is mandatory.

| Order | PC/Mobile profile | Pages |
| --- | --- | --- |
| 01 | Overview | 2 |
| 02 | Daily Trends | 1 |
| 03 | Licensing | 2 |
| 04 | Workforce & Identity | 3 |
| 05 | PC Devices & Compliance | 4 |
| 06 | Mobile Devices & Compliance | 2 |
| 07 | Applications | 2 |
| 08 | Collaboration & SharePoint | 5 |
| 09 | Messaging & Hybrid | 3 |
| 10 | Security & Resilience | 3 |
| 11 | Migrations | 3 |
| 12 | Data Quality & Evidence | 2 |

The eleven historical theme keys and logical report IDs stay stable; only the
Mobile theme is new. The PC/Mobile profile removes inter-page controls while
keeping in-page views and filters. Its reviewed obsolete SharePoint navigator
can be removed only if both configured targets are absent; working controls
are preserved, and partial resolution or binding drift fails for review.
Every adaptation and source-preservation check is recorded in the receipt.

```powershell
python -B SmartWorkplaceIntelligence/scripts/generate_thematic_reports.py `
  --source-report '<saved-project>/SmartWorkplaceIntelligence.Report' `
  --theme licensing --output '<new-private-output>/licensing'
```

For one changed page, regenerate only its theme. Automatically validate all
generated pages, connections, IDs, resources, bookmarks, interactions, filters
and fidelity to the saved master. Visually check only new pages, generator
adaptations or suspected rendering differences; do not replay every unchanged
page of an already validated master. Document the actual test scope, use one
Desktop test instance at a time and discard test selections when closing.

Manual publication is separate. Preserve the exact existing Service connection;
never guess a tenant URL or model ID. Reconcile existing Service names/IDs
before replacing renamed reports: a local `logicalId` is not a Service report
ID, and renumbering local filenames does not rename cloud items. Source
credentials, model parameters and refresh schedules remain deployment settings.

Run the offline regression suite with
`python -B -m unittest discover -s SmartWorkplaceIntelligence/scripts -p test_thematic_reports.py`.
Fixtures contain no tenant data and do not open Desktop or connect to Service.

## Qualified batch retention

Publication alone no longer retires local batch payloads. After a successful
SharePoint transfer and verified root pointer, the wrapper invokes coupled
server/SharePoint retention. `PreparedRetention.Mode` defaults to `Preview`:
fast metadata inventories and journals are produced without downloading payloads,
archiving or removing batches. Candidates are not SHA256-qualified actions:
metadata previews cannot be applied. `-Mode Preview -Verification Full` creates
a complete content-verified plan for manual review. Apply and configured Auto
retain the full inventory and pre-removal SHA256 checks. Private
activation of `Auto`, after server qualification, applies each newly generated
and revalidated plan without a daily approval or another scheduled task.

Both stores protect the current batch, its predecessor, and every batch in the
qualified private consumer registry. SharePoint keeps immutable batch folders
behind its pointer: by default only the current and previous validated batches
(plus consumer pins), retiring at most one folder per publication. The server
share keeps its full history by default (`ServerKeepBatches=0`). Apply and Auto
require a shared `JournalRootPath` for plans and journals. Every SharePoint
removal is validated before its intent is journaled, and its HTTP outcome is
recorded explicitly (`Recycled`, `NotRemoved` or `Uncertain`). An old protected batch does not stop the
remaining chain traversal. Missing/incomplete historical metadata, unexpected
content and unproven publication-chain membership remain blocked. There is no
age-only policy, automatic expiry of consumer protections or PBIP discovery.

Local empty, idle JSON transport/SharePoint-transition locks are recognized only
when their exact owner is a declared batch JSON. They remain in the physical
inventory but are excluded from cross-store payload comparison. Active, nonempty,
unknown or remote sidecars remain blocking. Local helper writers are excluded
through archival/removal by retained lock handles. Compact schema-3 plans retain
file identities, sizes and measured raw SHA256 values instead of receipt bodies;
full receipt validation, archival and pre-removal hashing are unchanged. Older
plans must be regenerated.

The registry template is `config/templates/PreparedBatchConsumers.json.template`.
Keep its private runtime copy beside the preparation wrapper. An operator must
review all declared distributed consumers; a local PBIP scan cannot establish
completeness. The template is deliberately empty and unqualified.

See [server qualification, activation and recovery](../SmartM365/SmartInventory/PreparedEvidence/docs/DEPLOYMENT.md#qualified-coupled-batch-retention)
for the exact configuration fields, coordination conditions, plan/hash workflow,
recovery limitations and offline tests. Until activation, local payloads can grow.
Retention errors never roll back or replace a correctly published root pointer.

Deploy `SmartM365` and `SmartWorkplaceIntelligence` as matching sibling folders,
or explicitly configure `WorkplaceIntelligenceRootPath`. Preparation requires
PowerShell 7 and Python 3.10 or later.

## Source layout and privacy

- `pbip/` contains sanitized Power BI report and semantic model definitions.
- `scripts/` contains preparation, reader, and offline validation components.
- `config/` contains source/output contracts and generic classification templates.
- `docs/` records evidence boundaries and integration details.

The source PBIP contains neutral connection defaults. Tenant and immutable-batch
parameters must be explicitly configured for a private deployment. Customer CSVs,
workbooks, caches, saved model data, and local configuration must remain outside
the public source. The current model does not import generic CMDB entity, alias,
relationship, or endpoint projections.

## Daily directory and mailbox trends

`Daily Directory Trends` and `Daily Mailbox Trends` are prepared from the three
governed `DATA-ALL` daily statistics exports. Each source keeps its last reading
per day and domain. The preparation emits domain detail and an `All domains`
aggregate at day × source × metric grain. A blank source cell is `Not measured`,
not zero; a tenant metric with any unmeasured contributing domain remains blank.
Missing calendar dates and unqualified totals are never interpolated.

The capture reference is the most recent prior source day whose core metrics
were measured and whose capture was complete. A domain absent from a later
reading is not itself an unmeasured metric. The source day is `Partial capture`
only when the absent domains' reference total exceeds the configured percentage
of that prior complete day's reference total (1% by default). An absent domain
reaching seven consecutive observed source days (configurable) becomes
`Domain absent (persistent)`. This transition is recorded as a quality event;
the domain leaves the capture reference on that day, so a permanently retired
domain does not keep all subsequent days partial. A returning domain becomes
active again. No automatic rule classifies a large fall in metric values:
account cleanup can be a real fall. Consequently, a day with truncated values
in domains that are still present can escape automatic qualification. For
example, a source day may be detected only because another domain is absent,
while the present domains are also truncated.

Optional private qualifications are supplied with
`-DailyTrendsQualificationPath` to `Invoke-PreparedEvidenceBuild.ps1`. Copy the
neutral `config/daily-trends-qualification.local.json.template` to a private
`*.local.json.txt` file outside source control. `manualDays` entries require
`source`, ISO `date`, and `reason`; they mark the source day `Partial capture`
before the reference update. `persistentDomains` entries require `source`,
`domain`, ISO `effectiveDate`, and `reason`; they explicitly retire that domain
from the capture reference on or after the effective date while it remains
absent. The configurable defaults are
`-DailyTrendsPartialThresholdPercent 1` and
`-DailyTrendsPersistentAbsenceObservations 7`. Do not place actual domain names,
qualification reasons, or generated trend CSVs in the public repository.

Core qualification and full measurement are different. AD Users core covers
total, enabled, and disabled users; `Inactive90DaysUsers` is optional for the
full measurement date. Exchange core covers local, remote, total, enabled, and
disabled mailbox counts; the mailbox type counts and local size are required
only for the fully measured date. The page shows both dates. Shares are
calculated after aggregation: Windows 11 enabled / enabled AD computer
accounts, and remote mailboxes / total mailboxes. Their period changes are
shown in percentage points.

The three native trend charts use one adaptive date axis. `All` and
`12 months` show ISO Monday weeks; each point is the last core-qualified
reading in that week. A week without a qualified reading remains blank.
`90 days` and `30 days` show calendar days. With a free date range and
`All` selected, spans above 120 days show weeks; shorter spans show days.
The chart subtitle names the active grain, and tooltips give the retained
reading date plus qualified, incomplete, and partial-capture day counts.
Cards, period deltas, and the Windows 11 crossover date always use daily
evidence. A weekly point can coexist with an incomplete or partial day in
the same week; the tooltip reports those days rather than removing the
qualified weekly point.

The daily directory page counts AD user accounts of all types, including
service, shared, and technical accounts. This is a different population from
the qualified workforce on Executive Overview. Its Windows series count
enabled AD computer accounts, including accounts outside Intune management;
Executive Overview counts managed Intune devices. Titles, axes, and scope
tooltips identify these populations. The report does not hard-code deployment
counts in those explanations.

## Deploying the daily trends extension

Update the matching `SmartM365` and `SmartWorkplaceIntelligence` sibling
folders on the operational share after the reviewed changes have been pushed.
Request the preparation through
`SmartM365/SmartInventory/LaunchersByOrchestrator/Request-WorkplaceEvidence-Prepare.cmd`
on an execution machine. The resident orchestrator retains election,
dependency, and concurrency control. Do not invoke preparation against a
synchronized review copy of the data.

Before refreshing Power BI, inspect the successful run's transcript and new
immutable batch. Its manifest and validated payload must include
`DailyDirectoryTrends.csv` and `DailyMailboxTrends.csv`. On the first successful
upgrade from a prior batch without these tables, the transcript reports
`Historical comparison initialized: Daily Directory Trends` and
`Historical comparison initialized: Daily Mailbox Trends`; the manifest records
both `HistoryKeyInitializations`. These messages occur once, not on subsequent
ordinary runs. No existing table's historical coverage may regress.

Confirm that `DATA-POWERBI/current.json.txt` selects the new validated batch,
and that any configured SharePoint transfer and read-back verification
completed before relying on a synchronized destination. Set the private
PBIP's `PreparedBatchId` to that reviewed immutable batch, or use `Latest`
after deploying the [automatic reader](docs/PREPARED-LATEST.md), retaining the
private source and expected-tenant parameters. Then refresh Desktop, verify
the new page and the existing 30-day trends, and publish through the approved
deployment process.

Refreshing this model against the old batch fails the two new table queries:
`GetPreparedEvidenceFile` rejects files absent from the pinned batch manifest
with `Invalid prepared batch manifest`. Their dependent calculations cannot
be qualified, and the full refresh can fail. Cached visuals are not evidence
of a successful upgrade; changing only the root pointer does not update the
explicit `PreparedBatchId` pin.

Previously published immutable batches remain transferable after this upgrade.
Transfer recognizes a complete pre-daily-trends cohort; a single daily file, or
missing daily files in a batch whose history metadata declares these tables,
is rejected before transport.

The current production wrapper uses the default daily-trends settings
(1 percent missing-domain share, 7 persistent-absence observations, and no
manual qualification file). Exposing private overrides in its local JSON and
forwarding them through the pipeline is a separate, deferred change.

## License decisions from SmartInventory Licences

Smart Workplace Intelligence no longer computes license recovery, dormancy or
E3-to-F3 decisions. `New-LicensingEvidence.ps1` and its two tables
(`License Evidence`, `License Optimization Evidence`) are retired. The report
consumes the SmartInventory Licences report: the five `M365_Licenses_Report*.csv`
files, the snapshot, and the two optional per-SKU files
`M365_Licenses_ReportSkuCandidates.csv` and `M365_Licenses_ReportSkuCapacity.csv`
(Licences 1.49 and later), copied as `Source-*` artifacts of the batch.

Before the batch is published, the preparation checks the licensing receipt
captured with it: completed receipt of the same tenant and producer, every
report CSV recorded in `Files` with the captured SHA256 and declared
`Published`, every row carrying the snapshot ID, decision rows carrying the
receipt `RunId`, and snapshot `ReceiptRun.RunId` equal to the receipt `RunId`.
The result is recorded in the batch manifest as `LicenseReportQualification`
(state, reason, snapshot ID, receipt run ID, one entry per file).

- `Qualified`: the license tables are loaded.
- `Not qualified`: a lineage defect (for example a report produced outside the
  receipt run or a file hash that differs). The batch is still published and the
  other domains are unaffected; the license tables are loaded empty, never with
  older values. Another tenant or an invalid snapshot structure still fails the
  batch. License tables carry no historical comparison key, so this state is
  never a coverage regression.
- `Not verified`: a batch built before this qualification existed. Its values
  are shown as that batch's state.

Every license card, table and chart shows this state in its subtitle. The
License Overview SKU slicer and capacity cards use `License SKU Capacity`:
`Capacity Class` is `Governed` or `Excluded` from the versioned Licences
exclusion list, and SKU families come from the Licences `ReportProduct` column
(Licences 1.50; a 1.49 file is accepted with an empty family). The Risk Priority
slicer and the Dormant 31-90D card are removed: they were SWI-only rules.
Immutable batches built before the retirement remain transferable.

### Deployment

The published model can use scheduled Service refresh through its
SharePoint-only transport. Apply this checkout's model and report
changes to the private PBIP (keeping its private source, expected tenant and
`PreparedBatchId` parameters and its editor settings) before pointing its
`PreparedBatchId` to a batch prepared by this version. For automatic selection,
follow the [Latest deployment order](docs/PREPARED-LATEST.md): integrate
definitions, republish, then set `Latest` in both Service and the private PBIP,
then configure the schedule. The checkout default is `Latest`; a blank value
fails. No publication or refresh is implied by editing the checkout.

If the private PBIP still has the retired tables when its `PreparedBatchId`
points to such a batch, the Desktop refresh fails on `License Evidence` and
`License Optimization Evidence` with `Invalid prepared batch manifest`
(`LicenseEvidence.csv` is no longer in the batch). The updated PBIP refreshes
older batches without error and labels their license data `Not verified`.

A batch reported `Not qualified` means that the Licences report in `DATA-LAST`
was not produced by the current licensing receipt run (for example a report
written outside a collection run): run the Licences collection again before the
next preparation. The preparation transcript states the result
(`License report qualified by licensing receipt run ...` or a
`License report Not qualified` warning with its reason).

Available capacity is never shown as a negative number. When consumed units
exceed enabled units the cards and tables show `Over-assigned`; the absolute
value is in the card tooltip and in the `Over-assigned Units` table column.

## Further reading

- [Prepared evidence reconstruction](docs/prepared-evidence-reconstruction.md)
- [Native domain migration and generic layer retirement](docs/CMDB-CONSOLIDATION.md)
- [Teams and SharePoint evidence](docs/TEAMS-SHAREPOINT-INTEGRATION.md)
- [Device system evidence](docs/DEVICE-SYSTEM-EVIDENCE.md)
- [Private classification templates](config/README.md)

Offline validation, a successful server preparation, publication, and Power BI
rendering are distinct checks. Source publication alone does not qualify a
deployment or every report interaction.
