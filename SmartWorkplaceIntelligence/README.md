# Smart Workplace Intelligence

Power BI reporting over governed SmartM365 workplace evidence and observed
histories. Domain pages cover workforce, devices, applications, licensing,
collaboration, messaging, and source quality. Native details retain their source
identity and grain; reported configuration and status are qualified evidence,
not proof of effective on-device behavior.

## Data flow

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
PBIP's `PreparedBatchId` to that reviewed immutable batch while retaining the
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

## Further reading

- [Prepared evidence reconstruction](docs/prepared-evidence-reconstruction.md)
- [Native domain migration and generic layer retirement](docs/CMDB-CONSOLIDATION.md)
- [Teams and SharePoint evidence](docs/TEAMS-SHAREPOINT-INTEGRATION.md)
- [Device system evidence](docs/DEVICE-SYSTEM-EVIDENCE.md)
- [Private classification templates](config/README.md)

Offline validation, a successful server preparation, publication, and Power BI
rendering are distinct checks. Source publication alone does not qualify a
deployment or every report interaction.
