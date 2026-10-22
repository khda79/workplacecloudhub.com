# Smart Inventory

Inventory scripts that can feed Power BI datasets, operational reports, CSV exports, SharePoint publishing, and other downstream consumers.

## Reliability and distribution

SmartInventory is a collection of independently versioned scripts and shared
modules; this README does not assign a suite version or stable qualification.
Shared CSV persistence, distributed orchestration, Graph collection and the
bounded AD/Exchange/readiness families have focused synthetic regression coverage.
Those results do not qualify Microsoft Graph, a tenant, on-premises systems or the
whole SmartInventory suite. Detailed audit maps, evidence and publication records
are intentionally retained outside the public repository.

Use a reviewed repository checkout with the required `Config/`, internal module
manifests and templates. A single copied collector is not a self-contained package.
No per-script ZIP release is required. Operational JSON, credentials, CSV exports
and logs must stay outside Git and public distribution.

The shared exporters reject conflicting source identities before writing
files. Existing identity-first CSV columns, names, delimiters and business fields
are preserved. A failed serialization cannot replace the previous valid CSV, but
this does not establish a transaction across all files or prove collection
completeness. Consumers must still check source dates, schemas and tenant identity.

## Current source receipts

Every producer publishing reporting CSVs to `DATA-LAST` now writes one
`SmartInventory_<producer>.current.json.txt`, with `Owner=SmartInventory-SourceReceipt`
and contract version 1.2. This metadata is shared evidence, not exclusive ownership
of the CSV by CMDB. One receipt lists all current-run published CSVs; there is no
receipt per CSV, no new history and no folder scan that qualifies older files.
The receipt is `Running` before acquisition and becomes `Completed` only after
successful required publication and row/hash/tenant checks. Failed runs preserve
native CSVs but invalidate completion evidence. Caches, temporary files, archives
and the remediation-script run bundles are not `DATA-LAST` publications.

The 17 CMDB source producers keep their explicit consumer-scope checks. The 30
other `DATA-LAST` producers declare their required and optional outputs in
`Modules/SmartM365.Core/SmartM365-SourceReceipts.json.txt`. Additional current-run
exports are also recorded; absent optional outputs are not replaced by old CSVs.
Effective selection parameters are included where applicable. For these producers,
`ScopeQualification=ConfiguredOutputsOnly` and `IsPartialInventory=null` mean
that whole-tenant exhaustiveness is **not established**, not that it succeeded.
The SharePoint on-prem Content producer publishes observed rows and a separate
`SharePoint_OnPrem_CollectionCoverage.csv` when coverage gaps occur. Its source
receipt verifies current file identity, rows, and hashes only; the coverage CSV,
run log, and warning exit code identify an incomplete collection. The generic
configured-output receipt does not qualify farm-wide completeness.
No receipt claims `FullInventoryQualified=true`. Derived reports remain derived;
creating their receipt does not refresh their upstream inputs.

Read-only and MAXITEMS modes do not overwrite canonical receipts. Core 1.0.72
and Windows PowerShell 5 compatibility 1.0.50 are required by the newly integrated
collectors. The common evidence runtime requires Core 1.0.72. CMDB preparation
accepts strict legacy 1.1 receipts during rollout and qualified shared 1.2 receipts;
existing metadata is never restamped or repaired from CSV timestamps. Intelligence
CSV names, columns, prepared contracts, generators and history keys are unchanged.
Qualify the next real acquisitions on the collection machines after deployment;
the offline tests do not establish production freshness or completeness.

Data quality takes precedence over runtime. A page limit reached by the Backup
Protected Mailboxes collector now fails the run before replacing the latest
CSV; use `-MaxItems` for a bounded test instead of treating `-MaxPages` as a
partial-export option. Windows Update reports the source and coverage of
`DaysSinceLastStatus`; the field remains blank when the export has no suitable
device-status date. A policy modification date is not substituted for it.

## Daily capture retention

Core 1.0.92 adds receipt-qualified daily capture retention for eight producers:
M365 UserActivity, Copilot Usage, M365 WorkplaceScope, Intune Configurations,
Intune Managed Apps, Licences, License Pricing, and Windows Autopatch Alerts.
Other collectors keep their existing retention behavior.

After a successful source receipt, while its producer lock is still held, the
policy keeps the last complete published set for each of seven calendar days
(today plus six preceding days). Completion dates use Europe/Paris, including
daylight-saving transitions. A later failed run or a smaller same-day export
selection cannot displace a complete set. Completion qualifies the configured
exports; it does not establish whole-tenant collection completeness.

The policy reuses matching current-run timestamped CSVs. When a receipt member
has no matching timestamped copy, it creates a byte-verified copy under
`DailyCaptures/<RunId>/`. This matters for collectors that previously retained
only some of their exports. CSV schemas, current source receipts and weekly
snapshots are unchanged. Capture metadata records identity, completion date,
exact paths, sizes and SHA256 under `DailyCaptureRecords/<producer>/`; metadata
records are kept after their payload is retired. No archive is created.

Per-file count/age cleanup is bypassed only inside the opted-in output root.
For Usage and Copilot this root is the parent of the run folders, rather than
the current run folder. Whole-set retention runs only after receipt completion;
it verifies retained captures and every candidate before deletion. If SharePoint
upload is enabled, captures and their metadata must be uploaded and read back
with matching SHA256 first. Remote deletion uses the inspected item ID and
`If-Match` ETag; a remote refusal stops local deletion.

`WeeklyHistory`, `DATA-LAST`, `Corrections`, source receipts and manifests are
excluded. Unindexed legacy CSVs are preserved: old run dates alone are not
completion proof, and this change does not infer or repair old capture records.
Metadata corruption, linked paths, missing/changed retained files, locked files,
or uncertain remote results stop further cleanup and produce a warning. Deletion
of a multi-file set is not an atomic transaction; inspect the retained metadata
and run log after an interrupted or partially completed cleanup.
An output root containing a filesystem link disables both daily retention and
legacy cleanup with a warning, while allowing collection to continue.

Offline regression coverage is available through
`Tests/Test-SmartM365DailyCaptureRetentionOffline.ps1`. It uses temporary
synthetic data and mocked SharePoint only; deployment and real server/SharePoint
qualification remain separate.

## Organization

- `ActiveDirectoryInventory/`: Active Directory inventory and reporting.
- `ExchangeInventory/`: Exchange Online and Exchange on-premises inventory.
- `M365Inventory/`: Microsoft 365, Entra, Power BI, and Microsoft Fabric inventory.
- `SharePointInventory/OnPremises/`: independent SharePoint Server 2016/2019 infrastructure and content collectors; manual farm-server execution, no orchestrator job or automatic upload.
- `M365Inventory/IntuneInventory/`: Intune inventory, Windows Update reporting, Autopilot, RBAC, applications, and remediation export utilities.
- `M365Inventory/IntuneInventory/EndpointAnalytics/`: standard Endpoint Analytics score, startup, app reliability, and work-from-anywhere exports without Advanced Analytics.
- `M365Inventory/Security/`: Secure Score, authentication-method registration, and Conditional Access configuration evidence.
- `M365Inventory/IntuneInventory/Security/`: Defender agent and firewall-health reports exported from Intune without changing device or policy state.
- `ExchangeInventory/BackupProtection/`: Microsoft 365 Backup protection evidence for mailboxes, SharePoint sites, and OneDrive accounts.
- `M365Inventory/Licensing/SmartM365-LicensePricing-Inventory.ps1`: normalization of governed license prices (`M365_License_Prices.csv`) and optional Windows 10 ESU prices (`Windows_ESU_Prices.csv`). The versioned France public-list baseline `M365_License_Prices_Public-France.csv` (every row estimated) is merged with an optional private override; see [License pricing](#license-pricing).
- `Launchers/Cloud/`: production launchers for cloud inventory scripts.
- `Launchers/OnPremises/`: production launchers for Active Directory and Exchange on-premises inventory scripts.
- `Launchers/Orchestrator/`: production launchers for orchestrator installation, start, stop, and restart operations.
- `Launchers/Jobs/`: one `Start-<JobName>.cmd` per job, generated by `Orchestrator/New-SmartM365-OrchestratorJobLaunchers.ps1`. Each runs the job immediately on the current server with the command the orchestrator builds: the template `Arguments`, the engine of `PowerShellEdition`, `-Tenant` (`prod` unless `SMARTM365_ORCHESTRATOR_TENANT` is set), `-Connect` when the script declares it, or a call to the job's `LauncherPath` with its `Arguments`. Extra options are appended. `Test-SmartM365OrchestratorPipeline.ps1` fails when they drift from the template, and `Test-SmartM365SmartInventoryLaunchersOffline.ps1` fails when a hand-written launcher that reproduces a job misses one of its template `Arguments`, uses another engine, or, for a `LauncherPath` launcher, drops `%*` or ignores the orchestrator tenant.
- `Maintenance/Remove-SmartM365-ExpiredTimestampedCsv.ps1`: one-time cleanup of the timestamped run copies that WeeklyHistory already covers (see below).
- `Maintenance/Import-SmartM365-LegacyDailyStats.ps1`: one-time import of legacy AD and Exchange on-premises daily statistics into the SmartInventory daily CSVs (see below).
- `Maintenance/Invoke-SmartM365-ServerDataCleanup.ps1`: complete tenant DATA inventory with preview and verified server-only cleanup. It reports every file, groups decisions by area, preserves compact Discovered Apps summaries, keeps the prepared-batch history unless `-RetireLinkedPreparedBatches` is passed (then with audit receipts), retires old orphan atomic-write temporaries, and leaves unproven categories for review.
- `LaunchersByOrchestrator/`: one `Request-<JobName>.cmd` per job, plus `Request-SmartM365-OrchestratorJob.cmd`. They ask the resident orchestrators to run the job (elected server, shared claims and locks, scheduled dependency rules) and wait for the result. Generated by `Orchestrator/New-SmartM365-OrchestratorJobLaunchers.ps1`.

Choose `Launchers/` to run a collector immediately on the current server, outside the orchestrator (troubleshooting, first deployment); `Launchers/Jobs/` reproduces a job exactly as scheduled. Choose `LaunchersByOrchestrator/` for an out-of-schedule production run: it cannot overlap a scheduled run of the same job and never consumes stale dependency data. The orchestrator itself starts seven jobs through `Launchers/` (`LauncherPath` in the jobs manifest), so that folder keeps its name and path.

Launcher names no longer carry a `-Prod` suffix. Test launchers are intentionally not maintained; use the PowerShell script parameters directly for targeted test execution.

The Secure Score, authentication-method registration, Conditional Access, Intune endpoint-security, and governed public-France license-pricing collectors are enabled once daily in `Orchestrator/Orchestrator-Jobs.json.template`. Microsoft 365 Backup collectors remain scheduled but disabled until their service-app/controller prerequisite is available. `Policy.Read.All` and `SecurityEvents.Read.All` are new bootstrap permissions. Authentication registration and Microsoft 365 Backup reuse `AuditLog.Read.All` and `BackupRestore-Configuration.Read.All`. Intune endpoint-security reports reuse the temporary export-job permission `DeviceManagementManagedDevices.ReadWrite.All`.

## License pricing

`SmartM365-LicensePricing-Inventory.ps1` 1.1.1 publishes `M365_License_Prices.csv`
(header unchanged) from three local settings:

- `PricingSourceCsvPath`: the public baseline
  `M365Inventory/Licensing/M365_License_Prices_Public-France.csv` (template default).
  Prices are Microsoft France public list prices in EUR excluding VAT, per user per
  month with annual commitment, `IsEstimated=True`. `EffectiveFrom` is the date the
  price was observed and `ContractReference` the public page used.
- `PricingOverrideCsvPath` (private, empty by default): reviewed contract prices with
  the same columns. The public baseline is always read; a private row replaces every
  public period of the same `SkuPartNumber` and `Currency`, and private-only SKUs are
  added. `IsEstimated` and `PriceSource` show which rows are contract prices.
- `EsuPricingCsvPath` (private, empty by default): Windows 10 Extended Security
  Updates prices, published as `Windows_ESU_Prices.csv`, an optional output of the
  collector's source receipt. Without it, ESU stays not valued and the file is not
  published. There is no public ESU row because Microsoft publishes no EUR price.

The private ESU CSV uses these columns:

```text
ProgramKey,CoverageYear,Currency,UnitPricePerDeviceYear,EffectiveFrom,EffectiveTo,IsCumulative,PriceSource,PricingBasis,ContractReference,IsEstimated
```

`ProgramKey` is `WIN10_ESU`, `CoverageYear` is 1 to 3, `Currency` must be `EUR` (no
currency conversion), dates are `yyyy-MM-dd` and required, years must follow each
other without overlapping, and `IsCumulative` states whether a later year also
requires the earlier years.

`UnitPriceMonthly` and `UnitPricePerDeviceYear` are published with a decimal point
whatever the culture of the collection host (`37.78`, never `37,78`). Source prices,
public or private, must use the same format: a decimal comma or a group separator is
refused instead of being misread.

Migration: a local `PricingSourceCsvPath` that points to a file other than the public
baseline keeps the former behaviour (the private file replaces the whole baseline) and
logs a warning. Move those prices to `PricingOverrideCsvPath`, then point
`PricingSourceCsvPath` back to the public baseline (or remove the local value so the
template default applies). Setting both a private `PricingSourceCsvPath` and
`PricingOverrideCsvPath` is refused. The runtime `.local.json` is created from its
template when absent and missing template keys are added without changing local values.

Not valued in the public baseline (no verifiable public EUR price per user):
Exchange Online Archiving for Exchange Server, Teams Phone with calling plans, Intune
Suite, Office 365 Extra File Storage, Copilot Studio (capacity billing), Dynamics 365
attach, Team Members and Operations Activity licenses, Power Apps per app, and Project
Online Premium. Provide contract prices through `PricingOverrideCsvPath` when needed.

## Timestamped CSV copies on SharePoint

`Publish-SmartM365Csv` uploads the latest copy and the timestamped run copy
(`<prefix>_yyyyMMdd_HHmmss.csv`, used by the per-run SharePoint links of the mails).
Since SmartM365.Core 1.0.61 (PS5 compatibility module 1.0.45), once the run copy is
uploaded, the timestamped copies of the same CSV in the same SharePoint folder older
than 7 days are deleted (`-SharePointRetentionDays`, 0 disables), the same retention
as the server. Latest copies and WeeklyHistory never match. Before, SharePoint kept
every run while the server kept seven days. Since Core 1.0.62 (PS5 1.0.46) the same
function also handles timestamped `.xlsx` workbooks (same prefix and extension only).
The Teams (0.34) and SharePoint Online (0.30) collectors, which upload their run CSVs
and workbooks directly, call it after each successful upload. The Active Directory
inventory publishes fixed file names and needs no purge. AD HealthCheck is left
unchanged on purpose: without WeeklyHistory, its SharePoint run copies are its only
history beyond the seven days kept on the server.

`Maintenance/Remove-SmartM365-ExpiredTimestampedCsv.ps1` removes the copies
accumulated before. It is read-only without `-Execute` and removes a copy only when it
is older than `-RetentionDays` (7), its folder has a WeeklyHistory whose manifest for
the copy's ISO week lists `<prefix>.csv` with a non-empty week copy, and `<prefix>.csv`
exists under the DATA-LAST folder next to DATA-ALL. WeeklyHistory, Corrections, Archive,
Orchestrator and run folders are never considered; every kept file is reported with
its reason in the CSV report.

```powershell
# Preview, then delete (synchronized SharePoint copy or UNC DATA-ALL)
.\Maintenance\Remove-SmartM365-ExpiredTimestampedCsv.ps1 -DataAllRootPath "<...>\SMART-M365\DATA\DATA-ALL"
.\Maintenance\Remove-SmartM365-ExpiredTimestampedCsv.ps1 -DataAllRootPath "<...>\SMART-M365\DATA\DATA-ALL" -Execute
```

## Legacy daily statistics import

`Maintenance/Import-SmartM365-LegacyDailyStats.ps1` extends the daily history of
`AD_Users_DailyStats.csv`, `AD_Computers_DailyStats.csv` and
`Exchange_OnPrem_Mailboxes_DailyStats.csv` with the files of older collectors of the same
organization (`AD_Users_Daily_Stats.csv`, `AD_Computers_Daily_Stats.csv`,
`Exchange_Daily_Stats.csv`: same columns, no tenant columns).

- SmartInventory rows win: a legacy row is imported only for a day without any
  SmartInventory row, with the tenant columns of the existing rows.
- Non-numeric legacy values become empty. A single missing part of
  Total = Enabled + Disabled (Exchange also Total = Local + Remote) is computed only when
  that identity holds on every complete row of both files.
- Exchange rows written before collector 1.66 carry 0/1 per mailbox type instead of a
  count (their type columns do not sum to `TotalMailboxCount`). They take the type counts
  of the legacy row of the same day and domain when its totals are identical; otherwise
  the type columns are left empty.
- The import is refused for a legacy domain unknown to SmartInventory or when the totals
  of overlapping days differ by more than `-MaxMedianDifferencePercent` (5).

Run it on the server share, outside the AD and Exchange on-premises collector windows. It
is read-only without `-Execute`; with `-Execute` it backs up every replaced file, writes
DATA-ALL and the DATA-LAST copy atomically, refuses a file changed during the run, and is
idempotent. The next collector runs republish the files to SharePoint.

```powershell
.\Maintenance\Import-SmartM365-LegacyDailyStats.ps1 -DataRootPath '<server share>\DATA\Tenants\prod' -LegacyFolderPath '<legacy csv folder>'
.\Maintenance\Import-SmartM365-LegacyDailyStats.ps1 -DataRootPath '<server share>\DATA\Tenants\prod' -LegacyFolderPath '<legacy csv folder>' -Execute
```

## Server DATA cleanup

Copy the single maintenance script to a machine that can access the server DATA
root; it also works from a temporary folder outside the repository. It does not
launch collectors, access Microsoft Graph, or wait for scheduled jobs. The root
leaf must match `-Tenant`. Preview creates `all-files.csv`, `summary.csv`,
`verified-actions.csv`, `timestamped-csv.csv`, and `orphan-temporaries.csv`
outside DATA. Run the same
command with `-Execute` to remove
verified items. The script rechecks source hashes, folder state, and prepared
pointer before deletion; it never sweeps arbitrary prepared batches. It reads the
effective `Orchestrator/Orchestrator-Jobs.json` when running in the repository to
qualify removal of the retired CMDB lock. When running from a temporary folder,
pass `-EffectiveJobsPath` pointing to the deployed runtime jobs file; otherwise
the lock stays in the review report.

Atomic writes stage a short sibling file (`~sm<12 hex>.pending`, `.previous`,
`.tmp`, `.partial`, `.download`, or the target extension such as `.csv`). The
JSON transport keeps its staging and previous copies when a write fails, and no
writer resumes them, so the script retires them as orphan temporaries. A file is
eligible only when its name exactly matches a known writer pattern, including the
legacy `<target>.<32 hex>.pending|previous|tmp|partial|download`,
`<target>.tenantkey.<32 hex>.tmp`, `.smartm365-write-test-*.tmp` and
`.smartm365-preflight-*.tmp` names; when its last creation or write time is older
than `-OrphanTemporaryMinimumAgeDays` (default 7), or
`-OrphanPreviousMinimumAgeDays` (default 30) for a `.previous` copy, which can be
the only prior version of a file; and when no `*.transport.lock` in its folder is
held. `-Execute` requalifies each file and deletes it while holding every
transport lock of the folder, logs each removal with a timestamp, and records the
result in `orphan-temporaries.csv`. Weekly `upload.pending` markers, the
transport lock files, and the ambiguous legacy `<name>.<32 hex>.csv` form are
never removed by this rule.

The server share keeps the long history of prepared batches; SharePoint keeps
only the current and previous batch through its own publication retention. By
default, linked batches older than the current and previous batch of the
`DATA-POWERBI\current.json.txt` chain are never planned for removal: they appear
in `all-files.csv` as `Keep` ("Linked prepared batch kept: share retains
prepared-batch history"), and the log reports how many were found and kept.
`-RetireLinkedPreparedBatches` enables their retirement. The current and
previous batch always stay; `-RetainedLinkedPreparedBatchCount` (default 0)
keeps that many further batches along the chain, and
`-LinkedPreparedBatchMinimumAgeDays` (default 0) keeps any batch whose id
timestamp is more recent. A retired batch keeps its `current`, `batch` and
`validation` receipts under `DATA-POWERBI\retired\<batch id>`, SHA256-checked
before the folder is removed. Unlinked batches remain review items.

```powershell
.\Maintenance\Invoke-SmartM365-ServerDataCleanup.ps1 -Tenant prod -DataRootPath '<server share>\DATA\Tenants\prod'
.\Maintenance\Invoke-SmartM365-ServerDataCleanup.ps1 -Tenant prod -DataRootPath '<server share>\DATA\Tenants\prod' -Execute
# Optional: retire linked batches beyond 10 older ones and older than 90 days.
.\Maintenance\Invoke-SmartM365-ServerDataCleanup.ps1 -Tenant prod -DataRootPath '<server share>\DATA\Tenants\prod' -RetireLinkedPreparedBatches -RetainedLinkedPreparedBatchCount 10 -LinkedPreparedBatchMinimumAgeDays 90
```

For the next server cleanup pass, `Maintenance/SmartM365-ServerDataSecondPass.ps1`
is read-only. It writes separate CSV reports for SHA256-qualified M365 Usage run
duplicates, WorkplaceScope snapshots versus DATA-LAST, generated Entra removal
scripts, old DATA-LAST outputs and coordination locks, repair backup integrity,
prepared `failed`/`retired` receipts, and the CMDB lock. It preserves
`Intune_Devices_Win11Readiness.csv` and does not delete any file. Review the
reports before extending the cleanup rules. Pass the effective deployed jobs
file when it is not next to the script.

```powershell
.\Maintenance\SmartM365-ServerDataSecondPass.ps1 -Tenant prod -DataRootPath '<server share>\DATA\Tenants\prod'
```

After reviewing a second-pass report, use
`Maintenance/SmartM365-ServerDataSecondPassCleanup.ps1` for a separately
approved, file-only cleanup. Supply the original report directory, a fixed UTC
cutoff, and the approved counts for Usage duplicates, old generated Entra
removal scripts, DATA-LAST test outputs, and the transport log. Preview writes
`plan.csv` outside DATA; Execute requires that exact plan and recomputes all
candidate and retained-copy SHA256 hashes. It refuses changed or missing files
and never removes directories, distinct Usage files, WorkplaceScope history,
repair backups, prepared diagnostics, or the CMDB lock.

```powershell
$cleanup = @{
    Tenant = 'prod'
    DataRootPath = '<server DATA tenant root>'
    EvidenceDirectory = '<reviewed second-pass report directory>'
    CutoffUtc = [datetimeoffset]'<approved UTC cutoff>'
    ExpectedUsageFiles = <approved count>
    ExpectedEntraScripts = <approved count>
    ExpectedTestOutputs = <approved count>
    ExpectedTransportLogs = <approved count>
}
$preview = .\Maintenance\SmartM365-ServerDataSecondPassCleanup.ps1 @cleanup
.\Maintenance\SmartM365-ServerDataSecondPassCleanup.ps1 @cleanup -Execute -PlanPath $preview.PlanPath
```

For a separately approved residual cleanup, use
`Maintenance/SmartM365-ServerDataResidualCleanup.ps1`. It checks the effective
orchestrator jobs file under the tenant DATA root, an exclusively available
retired CMDB lock, the repair audit and every rollback-copy hash, and the
reviewed DATA-LAST file names, sizes, and timestamps. It also verifies that
each repaired output remains available, including an identical Discovered Apps
SummaryHistory copy when the old weekly detail has expired. Preview records a
hash-pinned plan outside DATA. Execute requires that plan, repeats the checks,
and removes only the approved DATA-LAST files, lock, and one repair backup.
The report retains a copy of the repair audit, a retained-output map, and
the backup directory list and execution results. Extra empty directories
inside that backup are permitted; linked directories or unlisted files stop
the run. Removing the repair backup discards its pre-repair rollback copies;
keep the external report.

```powershell
$cleanup = @{
    Tenant = 'prod'
    DataRootPath = '<server DATA tenant root>'
    EvidenceDirectory = '<reviewed second-pass report directory>'
    BackupFolderName = '<approved repair backup folder name>'
    ExpectedAuditTenantKey = '<audited tenant key>'
    ExpectedRepairFiles = <audited repair receipt count>
    ExpectedSummaryFallbacks = <approved fallback count>
    ApprovedDataLastNames = @('<reviewed obsolete output.csv>')
}
$preview = $null
$preview = .\Maintenance\SmartM365-ServerDataResidualCleanup.ps1 @cleanup
if ($null -eq $preview -or $preview.Mode -ne 'Preview') { throw 'Preview did not complete.' }
.\Maintenance\SmartM365-ServerDataResidualCleanup.ps1 @cleanup -Execute -PlanPath $preview.PlanPath
```

## Active Directory launchers

Use the production launchers under `Launchers/OnPremises/`:

- `Start-SmartM365-ActiveDirectory-Inventory.cmd`: runs the complete Active Directory inventory for the selected tenant, including live domain collection, consolidated CSV generation, enrichment, duplicate and mail-routing analysis, reports, publishing, and weekly history.
- `Start-SmartM365-ActiveDirectory-DuplicateIdentity-Mail.cmd`: skips live Active Directory collection, reads the existing `DATA-LAST/AD_Users_AllDomains.csv`, regenerates the four duplicate and mail-routing diagnostic CSV files and the `AD_Users_IdentityAndMailRoutingIssues.xlsx` workbook, then forces the diagnostic email to be sent even if it was already sent that day.

Both launchers use tenant `prod` by default. The orchestrator can override it through `SMARTM365_ORCHESTRATOR_TENANT`. The mail-only launcher is intended for controlled reanalysis or notification retesting and does not replace a regular full inventory.

## Account classification governance

The account-type rules and workforce population mappings are tenant configuration. The real rules live in the private, git-ignored `Config/AccountClassification.local.json` (stored as `AccountClassification.local.json.txt`); only the neutral `Config/AccountClassification.local.json.template` is published. Create the private file from the template and adapt it to the tenant naming conventions before the first inventory; a missing file stops the AD enrichment and the prepared workforce evidence instead of falling back to generic rules. All readers use `Common/SmartM365.AccountClassification.psd1`. The Active Directory enrichment writes `AccountType`, `AccountPopulation`, and `AccountClassificationRuleVersion` to its enriched user export.

- `Human`: named and external-person accounts included in workforce KPIs.
- `Non-human`: service, shared-mailbox, room-mailbox, and system accounts excluded from workforce KPIs.
- `Review required`: admin, generic, unclassified, conflicting, or cloud-only accounts that must not be silently treated as human or non-human.

Site and persona columns use the same private workbooks as SmartWorkplaceIntelligence: `SmartWorkplaceIntelligence-SiteClassification.xlsx` and `SmartWorkplaceIntelligence-PersonaClassification.xlsx`, read from the SharePoint DATA root on every Active Directory inventory run. They replace `EntityDirectories.xlsx` and `Personae-Keywords.xlsx`. All readers use `Common/SmartM365.WorkplaceClassification.psd1`, which applies the prepared-evidence rules: a non-human exclusion first, then the HQ site persona, then job-title keywords filtered by country, and the AD description when the title has no unique match. Equal-priority matches across personas stay `Review required`. A failed download, a missing worksheet or an invalid rule stops the inventory instead of publishing blank or wrong site and persona columns.

- `EntityLabel`, `EntitySiteType` and `EntityIntegrationType` come from the site workbook (`Site Display Name`, `Site Type`, `Integration Status`), keyed by the AD attribute named in its Settings. An unmapped code gives `NOTFOUND`, and the integration type then falls back to the domain rule.
- `JobFamily` holds the persona or review state. For a `Review required` multiple-match persona, `M365LicenseTargetPersona` keeps the current base license (E3, otherwise F3), so an uncertain classification never proposes a license change. The `Relaxed` columns now carry the same values as the main ones.

Organization-specific directory rules are also private, in the `DirectoryRules` section of the same file. `IntegrationFallback` holds the ordered domain rules that give `EntityIntegrationType` for users and computers whose site is unmapped. Each rule has `Domains`, an optional `CountryIn` or `CountryNotIn`, and a `Result`. For computers, the country is read before `ComputerCountryOuSuffix` in the distinguished name. `HeadquartersOuContains` lists the OU fragments that mark headquarters accounts and computers. `Windows11OutOfScopeDomains` lists the domains excluded from the Windows 11 migration phases. A missing section stops the enrichment. The Exchange hybrid identity issues collector reads the routing domain from the tenant `RemoteRoutingDomain` and the allowed TargetAddress domains from `AllowedTargetDomains` in its local configuration. It refuses the template values.

Change `RuleVersion` whenever a classification or population rule changes. Keep the three populations disjoint; validate the template and the private rules with `Tests/Test-SmartM365-AccountClassification.ps1` before running a full inventory.
