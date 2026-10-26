# Microsoft 365 Inventory

Microsoft 365 and Entra inventory scripts outside the Intune-specific surface.

## Organization

- `Devices/`: Entra device inventory.
- `Domains/`: verified domain inventory.
- `IntuneInventory/`: Intune inventory, Windows Update reporting, Autopilot, RBAC, applications, and remediation export utilities.
- `IntuneInventory/EndpointAnalytics/`: standard Microsoft Intune Endpoint Analytics inventory; Advanced Analytics report families are excluded.
- `Licensing/`: license and service plan inventory.
- `PowerBI/`: tenant-wide Power BI and Microsoft Fabric activity-event inventory.
- `Teams/`: Teams collaboration inventory plus Teams Phone PSTN, Operator Connect, and Direct Routing usage.
- `SharePoint/`: SharePoint Online site, storage, list/library, permission, and external sharing inventory.
- `SyncHealth/`: Azure AD Connect synchronization freshness checks and stale hybrid data alerts.
- `Usage/`: Microsoft 365 user and workload usage reports for FinOps and usage analysis.
- `Users/`: active user inventory.

## Microsoft 365 Licensing

`Licensing/SmartM365-Licences-Inventory.ps1` 1.51 publishes normalized license and
service-plan data without repeating user and product labels on every service-plan
assignment.

| Entity | Latest CSV | Notes |
| --- | --- | --- |
| User license assignments | `M365_Licenses_Users.csv` | One row per user and assigned SKU, including direct or group assignment context. |
| User service-plan states | `M365_Licenses_UserServicePlanStates.csv` | Compact fact table keyed by tenant, user, SKU, and plan. `StateCode` records whether the plan is active, disabled, pending, or in error. |
| Service-plan catalog | `M365_Licenses_ServicePlans_Catalog.csv` | Deduplicated SKU and service-plan names used to enrich the compact fact table. |
| Exchange service plans | `M365_Licenses_ServicePlans.csv` | Existing enriched Exchange/SPE subset retained for compatibility. |
| Tenant SKUs | `M365_Licenses_Tenant.csv` | Tenant subscription and capacity inventory. |
| Licensing groups | `M365_Licenses_Groups.csv` | Groups discovered as license assignment sources. |
| License report | `M365_Licenses_ReportSummary.csv`, `M365_Licenses_ReportCandidates.csv`, `M365_Licenses_ReportAssignmentIssues.csv`, `M365_Licenses_ReportGaps.csv`, `M365_Licenses_ReportSources.csv` | Recovery, E3-to-F3 review, assignment issues, gaps and source states of one snapshot; optional outputs of the licensing source receipt. |
| License SKU decisions | `M365_Licenses_ReportSkuCandidates.csv`, `M365_Licenses_ReportSkuCapacity.csv` | Per-SKU candidates (disabled accounts, Copilot without licensed activity) and per-SKU capacity of the same snapshot; optional outputs of the licensing source receipt. |
| License views by AD scope | `M365_Licenses_ReportByDomain.csv`, `M365_Licenses_ReportByCountry.csv` | Distinct F1/F3/E3/E5/Copilot users, licensed users and recovery candidates per AD domain and per AD country of the same snapshot (1.60); optional outputs of the licensing source receipt. |

Every full inventory (no `-MaxItems` or `-TopUsers`) publishes the license report,
whether or not an email is sent: `DATA-LAST/M365_Licenses_ReportSnapshot.json.txt`
and the five `M365_Licenses_Report*.csv` files. The single JSON document contains
the exact seven overview products, four-suite recovery and activity counts,
E3-to-F3 review, sections 06/07 mailbox and AD gaps, candidate details, source
states and hashes, tenant key, rule version and a unique snapshot ID. Schema 2 is
unchanged; the additive `Rendering` object carries the source provenance needed to
render the email, and the additive `ReceiptRun` object names the licensing receipt
run (`RunId`) that proves the licensing inputs. The report is built after the
canonical licensing CSVs are published and before the source receipt completes.
The licensing CSVs it reads are qualified by the active receipt's own publication
hashes (SmartM365.Core `Get-SmartM365CurrentSourcePublication`) because the receipt
is still running at that point; other producers' CSVs keep their own receipts.
The five report CSVs are recorded as optional files of
`SmartInventory_SmartM365-Licences-Inventory.current.json.txt`; the required
files and the receipt status do not depend on them. The receipt records CSV files
only: the snapshot `.json.txt` is not hashed in it. The snapshot is tied to the
five CSVs by the `SnapshotId` carried on every CSV row, and to the receipt by
`ReceiptRun.RunId`. An unavailable usage source
leaves the related indicators `N/D` and the report is still published. If an
integrity check prevents the report, the receipt still completes with the report
files declared unpublished, the previous report files are left in place, and no
email is sent. The log and the run state (`Notes` in the `.run.json.txt` file)
state that the report present in `DATA-LAST` does not belong to the current
receipt; consumers must treat a report absent from the current receipt as stale. When SharePoint upload is configured, the snapshot and the five CSVs are
uploaded after registration; an upload failure is logged, noted in the run state,
and the email is not sent.

Per-SKU license decisions are published with the same report snapshot, in two
additive files that the email, the workbook and the five report CSVs do not use.
Every row carries the report `SnapshotId`, `TenantKey` and `ReceiptRunId` (the
licensing receipt `RunId`, equal to `ReceiptRun.RunId` in the snapshot).

- `M365_Licenses_ReportSkuCandidates.csv`: one row per user and SKU.
  - `Disabled account`, for every paid SKU except the free, trial, viral and
    capacity SKUs listed in the versioned
    `Licensing/SmartM365-LicenseExclusions.json.txt` (SkuPartNumber, Reason,
    Source). The local JSON can only add exclusions through
    `LicenseDecisionAdditionalExcludedSkus`; it never removes a versioned one, so
    repository updates of the list always reach the servers. The account state
    combines Entra and AD: disabled in both (or disabled in Entra for a cloud-only
    account) is `Recommended`; an Entra/AD conflict, or a synchronized account
    missing from AD, is `Review`; an account absent from the Entra export is not
    counted. The existing suite recovery of the report still uses the Entra state.
  - Exchange guard: a SKU holding an Exchange mailbox or archive service plan
    (detected in `M365_Licenses_ServicePlans_Catalog.csv`, not by SKU name;
    `EXCHANGE_S_FOUNDATION` excluded) is `Blocked by dependency` when the mailbox is
    above `LicenseDecisionMailboxReviewSizeGB` (50 GB), has an active archive, a
    litigation hold or a retention hold, and `Review` when one of these states is
    unknown.
  - Room, equipment and scheduling mailboxes and Teams Phone resource accounts
    (`MCOEV_VIRTUALUSER` plan or a resource-account phone assignment) are never
    plain candidates: `Review`, with the reason.
  - `Copilot without activity`: no Microsoft 365 Copilot activity in the licensed
    apps during `LicenseDecisionCopilotInactivityDays` (60) days ending at the report
    refresh date. Copilot Chat, Edge or agent activity alone is `Review` with the
    reason `Copilot Chat only`. An assignment updated less than
    `LicenseDecisionCopilotAssignmentGraceDays` (30) days ago is not a candidate;
    `LastUpdatedDateTime` is a lower bound of the assignment age (group
    reprocessing refreshes it), so it can only delay a candidate. An unknown
    assignment date is `Review`. A missing, stale or anonymized Copilot report, or
    one whose period is shorter than the window, is `Not measured`: no candidate
    and no zero.
  - A suite already counted in `M365_Licenses_ReportCandidates.csv` (Recovery or E3
    to F3 review) is `Referenced` with `CountedIn`, never counted again.
  - Dependency guard (`Licensing/SmartM365-LicenseDependencies.json.txt`, dated
    Microsoft sources): a qualifying license (Microsoft 365 suite for Copilot,
    Dynamics 365 base for an attach license) whose removal would leave a kept
    dependent license without a qualifying license is `Blocked by dependency`;
    `Dependents removed together` marks a removal valued with its dependents.
- `M365_Licenses_ReportSkuCapacity.csv`: one row per tenant SKU with enabled,
  consumed and available units, assigned and disabled-account users, candidates by
  recommendation, the exclusion reason and `UsageEvidence`: `Measured` for Copilot
  when the report is qualified, `Not measured` for Dynamics 365 (active users are a
  signal only, never a candidate) and `Not applicable` otherwise.

Suites are recognized by part number and, for the EEA and no-Teams editions, by
SkuId GUID from the Microsoft product names and service plan identifiers CSV;
unattended licenses stay out. A missing or invalid exclusions or dependency file
stops the decisions only. A decision failure is logged as a warning and noted
in the run state; the report, its five CSVs and the email are not affected, and the
two decision files are then declared unpublished in the receipt. Consumer reports should read this document rather than independently
recalculating recovery or downgrade rules.
They must display the snapshot ID and date, reject mixed IDs across imported
tables and preserve JSON `null` as N/D rather than a verified zero. The
candidate lists contain user identifiers and names and require the same private
access control as the existing license-user exports.
The license overview and recovery email is sent at most once per tenant and
Europe/Paris calendar day, including runs with `-SendLicenseSummaryEmailOnly`.
It never recomputes: it renders the published snapshot after checking that the
licensing receipt is completed for the tenant, that each report CSV has the hash
recorded by that receipt, and that every report row carries the snapshot ID. The
mail footer prints that ID. The full collector completes and qualifies its source
receipt before sending the email. A receipt failure stops the email instead of
publishing unqualified recovery KPIs.
The licensing consumer accepts AD `PartialAccepted` coverage only from a completed,
error-free AD producer with `ConsumerScopeQualified=true`, matching successful
file/domain proofs and an unchanged CSV hash. All expected domains must be accounted
for, at least one domain must be collected, and every unavailable domain must be
explicitly covered by `NonBlockingDomainErrors`. Available domains in that allowlist
remain included: it is a connectivity-failure tolerance, not a permanent exclusion.
Other source readers retain their existing strict partial-receipt rejection.
Accepted partial AD evidence is `Provisional`, not an automatic `ForceAdCsvAnalysis`
bypass. Its collected/expected counts and tolerated unavailable domains remain
visible in the email and source CSV reason; the private snapshot also retains the
structured domain coverage. Full coverage is never claimed. In section 07, `Members`
retains the Entra member population. Unmatched members remain unassessed when AD
coverage is partial; `NoObservedMatch` is then JSON null / CSV N/D, not zero or
confirmed AD absence. `UnmatchedScopeUnknown` records that limitation in the private
snapshot. An AD domain is never inferred from a user's UPN suffix.
Section 06 splits UserMailbox accounts without F1/F3/E3/E5 into Entra enabled,
disabled and unqualified account states. The three counts reconcile with the
section total when the Entra source is qualified. The top recovery KPI band also
shows the separate E3-to-F3 downgrade review count; it is not a recovery total.
Its overview cards emphasize the used percentage, with used and enabled license
counts below it; an unsubscribed product shows `N/A`.
The three recovery KPI cards appear immediately below `License overview`, ahead
of the suite and other-product cards and the numbered detail sections.
Email-only mode sends the report published by the last full inventory; it does
not rebuild it. `-ForceAdCsvAnalysis` applies to a full inventory only and is
rejected in email-only mode. The former `-BypassLicenseUsersReceipt` switch is
removed: the full inventory qualifies its own licensing CSVs directly.
By default, `DATA-LAST/M365_Licenses_SummaryEmail_SendState.json.txt` records
delivery state. Set `LicenseSummaryMailStatePath` in the private local JSON to
the same shared path on every collection host when runs can send from multiple
machines. A failed preparation can be retried automatically; an interrupted
or uncertain delivery stays pending to avoid a duplicate. After checking delivery,
use `-ForceLicenseSummaryEmail` to send again that day. This switch works with
both a full inventory and email-only mode. On a full inventory, it does not
override a disabled email configuration or permit sampled inventories to send.

### Licenses by AD domain and country (1.60)

Each Entra account is attributed to one AD account of `AD_Users_AllDomains.csv`
by `ImmutableId_AD`, then by a unique UPN, the same rule as the AD activity
indicators. Its AD domain is `DomainName` (lower case) and its country is
`CountryName`, or the `Country` code when the name is empty. Cloud-only accounts,
synchronized accounts without AD match, ambiguous matches (duplicate AD UPN, or
ImmutableId and UPN pointing to different AD accounts) and accounts without an AD
country keep explicit rows (`ScopeKind=Unattributed`); a domain is never inferred
from the UPN suffix. With qualified `PartialAccepted` AD coverage, unmatched
accounts are `AD domain not collected` and the views are `Provisional`. An AD
export without `DomainName` leaves the views N/D; the tenant report is unchanged.

- The tenant email shows, before section 01, `Licenses by AD domain` and
  `Licenses by country`: distinct users per suite, Copilot, licensed users (a user
  with several suites counts once) and recovery candidates. Rows add up to the
  tenant totals; any difference leaves the views N/D, never approximate. Up to 25 countries
  are listed; above, the 20 largest plus `Other countries (n)`. Subscription units
  are tenant-wide and are never split.
- The same rows are published in `M365_Licenses_ReportByDomain.csv` and
  `M365_Licenses_ReportByCountry.csv`. The existing report CSVs keep their exact
  schemas. The snapshot gains an additive `AdScope` object: both breakdowns, the
  per-domain measures (suite counts, E3-to-F3 review, sections 06/07, AD activity,
  assignment issue counts) and the domain and country of the users listed in the
  report only, not of the whole directory. A snapshot without `AdScope` (published
  before 1.60) stays readable.

Per-domain emails are sent after the tenant email, each with its own daily state
`M365_Licenses_SummaryEmail_SendState.<domain>.json.txt` next to the tenant state.
Configure them in the private local JSON:

```json
"LicenseSummaryDomainMails": [
  { "Domain": "<ad-domain-fqdn>", "To": "<recipient@domain>", "Cc": "", "Enabled": true }
]
```

`Domain` is the AD `DomainName`, case-insensitive; `To` is required; `Enabled`
defaults to true. A duplicate or invalid entry fails email-only mode and is logged
as a warning on a full run, where the tenant email is still sent. A domain email
uses only its own `To`/`Cc` (never the global `Cc`). Its user counts, alerts,
recovery, E3-to-F3 review and sections 06/07 count only the domain accounts; units
are labeled as tenant-wide. Its workbook lists only the domain users. A configured
domain with no attributed user, or a report without AD views, receives an explicit
`N/D` notice without attachment, never an empty report. One domain failure never
stops the other domains or the tenant email. In email-only mode,
`-LicenseSummaryEmailDomain <domain>` sends only that domain (with
`-ForceLicenseSummaryEmail` to resend it the same day).

The attached recovery workbook (1.61) follows the email style. It is named
`M365_Licenses_Recovery_<tenant or AD domain>_<yyyyMMdd>.xlsx` and opens on an
`Overview` sheet: the email header (scope and snapshot), KPI tiles (recovery E5,
E3, F3/F1, E3-to-F3 review, awaiting capacity, licensing errors), the recovery by
license table that reconciles with the email, links to every sheet and a reminder
that these are indicators, not removals. The `Recovery candidates`, `E3 to F3
review`, `Awaiting licenses` and `Licensing errors` sheets are Excel tables from
A1 with readable headers, fixed widths, banded rows that survive a sort, `AD
domain` and `Country` columns when the report has its AD views (a domain workbook
keeps only `Country`), recovery reasons colored by type (disabled, no activity,
shared mailbox), a highlighted primary Intune PC, activity dates older than 90 days
greyed, data bars on mailbox and OneDrive sizes, and a landscape print setup with
repeated headers and a scope and snapshot footer.
The recovery sheet includes `Last AD logon` from the AD
`LastLogonDate` export and `Last M365 activity` from the M365 active-user
activity report. Both are Excel date values displayed as `yyyy-mm-dd`, so they
can be sorted chronologically. An empty cell means the date could not be
qualified; identified shared mailboxes have empty date cells. AD last logon is
replicated and approximate. The email also has a separate E3-to-F3 review
indicator for enabled E3 users with an observed UserMailbox below 2 GB, no
archive or hold, no Windows/Mac Apps use, OneDrive storage below 2 GB, and no
second target suite. Existing recovery candidates are excluded from that review.
Missing input remains `N/D`. The five usage reports (`M365_Users_Activity.csv`,
`M365_Mailbox_Usage.csv`, `M365_Email_Activity.csv`, `M365_Apps_Usage_180D.csv`,
`M365_OneDrive_Usage.csv`) are used only when the completed UserActivity receipt
lists them with the same SHA256 (fail-closed: a missing, running or failed receipt,
an unlisted file or a different hash leaves the source N/D with its reason), then
their tenant, report period and refresh date (14 days) are checked.
Frontline eligibility and required E3 features require a manual review before
changing a license.
The same workbook has an `E3 to F3 review` sheet listing each qualified review
candidate with mailbox size, OneDrive storage, archive and hold state, and
Windows/Mac Apps usage flags. Its row count is checked against the email KPI;
the Summary sheet also shows the E3 review candidate count. The review sheet
is present with headers when there are no qualified candidates.

Service-plan state codes are `A` (enabled and successfully provisioned), `D`
(disabled), `PA` (pending activation), `PI` (pending input), `PP` (pending
provisioning), and `E` (error). The compact CSV therefore remains directly
filterable by `TenantKey`, `UserId`, `SkuId`, and `PlanId`. Its rows are written
to a disk-backed staging file during collection, avoiding an in-memory copy of the
complete user/service-plan fact table. WeeklyHistory keeps
`M365_Licenses_UserServicePlanStates_Detailed.csv` with the expanded `IsEnabled`
and `PlanStatus` columns, streamed from the compact CSV. When both weekly names
exist, the former `M365_Licenses_UserServicePlanStates.csv` duplicate is retired
only after the detailed replacement has been validated.

The former `M365_Licenses_ServicePlans_Detailed.csv` export is disabled. On a
normal run, its current non-timestamped DATA-ALL/DATA-LAST and SharePoint copy is
retired only after both compact replacement files have been published
successfully. Timestamped and weekly-history copies are preserved. A `-MaxItems`
validation run never removes or replaces canonical files.

## Intune Discovered Applications

`IntuneInventory/Applications/SmartM365-Intune-DiscoveredApps-Inventory.ps1`
publishes a normalized discovered-application model:

| Entity | Latest CSV | Notes |
| --- | --- | --- |
| Application catalog | `Intune_DiscoveredApps_Summary.csv` | One row per discovered Windows application/version with publisher, platform, and reconciled device count. |
| Application-device relations | `Intune_DiscoveredApps_AppDeviceRelations.csv` | Compact `TenantKey`, `AppId`, and Intune managed `DeviceId` fact table. Join to the Summary on `TenantKey` plus `AppId`, and to `Intune_Devices_Inventory.csv` on `TenantKey` plus relation `DeviceId` = inventory `Device ID`. |

The compact relation retains every discovered app/device association while device
metadata remains in the standard Intune device inventory. Summary and relation
WeeklyHistory snapshots are both retained. The former
`Intune_DiscoveredApps_DeviceDetail.csv` current export is retired only after the
matching Summary and compact relation have both been published successfully in a
normal run. Its timestamped and existing WeeklyHistory copies are preserved, and a
`-MaxItems` run never removes canonical files. The first normal compact run can
reuse and project a recent legacy DeviceDetail cache, avoiding an unnecessary full
Graph refresh. Resume checkpoints include the committed partial-file length, so
uncommitted trailing rows are rolled back before collection continues. An incomplete
cache manifest is discarded and all target application statistics are rebuilt from
the compact CSV, allowing valid app relations to be reused while mismatched or
duplicate app counts are recollected from Graph. Same-week Summary and relation
snapshots are refreshed, while unchanged legacy WeeklyHistory files remain preserved
without being uploaded again.
## Graph collection completeness

The consolidated Graph audit covers all 22 production collectors in this folder.
Every collector-owned `@odata.nextLink` loop now rejects a missing `value` property
and a repeated next link. Results remain buffered until the complete page chain is
available; an exhausted page failure therefore cannot publish an earlier prefix as
a successful complete collection. SDK-backed list collectors retain their existing
`-All` paths.

Explicit collector retry wrappers use bounded attempts for HTTP 408, 409, 429,
500, 502, 503 and 504 and honor integer `Retry-After` when the response exposes it.
The Microsoft Graph SDK remains responsible for its internal retry policy where a
collector directly uses an SDK cmdlet.

Canonical `DATA-LAST` files are written with the shared atomic CSV publisher. The
large Discovered Apps relation is validated first, then copied through a temporary
file in the destination folder before replacement. SharePoint header-only exports
use the same atomic writer. Compliance policy detail is not promoted when any
policy or setting page is incomplete; its independent complete device summary may
still be published. Atomicity is per file, not across a family of CSV files.

No live Graph or tenant behavior is certified by the focused synthetic regression
coverage. Detailed audit evidence is retained outside the public repository.

Entra Devices 1.13, Licensing 1.16 and SharePoint 0.26 additionally make
header-only current publication, weekly manifests and append histories atomic.
Teams collaboration 0.27 applies the same rule to its optional append histories.
These persistence-only changes do not change source columns, joins, license
calculations or collection algorithms.

## Microsoft 365 Usage Reports

`Usage/SmartM365-M365UserActivity-Inventory.ps1` exports Microsoft Graph Reports data and publishes stable CSV files to the tenant `DATA-LAST` folder.

Without `-Reports`, the script exports `Office365ActiveUserDetail`, `M365AppUserDetailD30`, and `M365AppUserDetailD180`.

Supported report selections:

| Report | Latest CSV | Notes |
| --- | --- | --- |
| `Office365ActiveUserDetail` | `M365_Users_Activity.csv` | Normalized workload activity by user. |
| `MailboxUsageDetail` | `M365_Mailbox_Usage.csv` | Graph usage view: storage, quotas, deleted items, archive flag, last activity. Complements the EXO mailbox inventory. |
| `OneDriveUsageAccountDetail` | `M365_OneDrive_Usage.csv` | OneDrive storage, file counts, active files, owner, last activity. |
| `SharePointSiteUsageDetail` | `M365_SharePoint_SiteUsage.csv` | SharePoint site storage, file/page activity, owner/site metadata from Graph reports. |
| `SharePointActivityUserDetail` | `M365_SharePoint_UserActivity.csv` | SharePoint file viewing/editing, sync, internal/external sharing, and page visits by user. |
| `Office365ActivationUserDetail` | `M365_Apps_Activations.csv` | Microsoft 365 Apps / Office activations by user and platform; no period parameter. |
| `M365AppUserDetailD30` | `M365_Apps_Usage_30D.csv` | Microsoft 365 Apps user activity for D30. |
| `M365AppUserDetailD180` | `M365_Apps_Usage_180D.csv` | Microsoft 365 Apps user activity for D180. |
| `TeamsUserActivityUserDetail` | `M365_Teams_UserActivity.csv` | Detailed Teams user activity. |
| `TeamsDeviceUsageUserDetail` | `M365_Teams_DeviceUsage.csv` | Teams client device and operating system usage by user. |
| `EmailActivityUserDetail` | `M365_Email_Activity.csv` | Exchange email send/read/receive activity. |

Examples:

```powershell
.\Usage\SmartM365-M365UserActivity-Inventory.ps1 -Tenant prod -Period D180 -Connect
.\Usage\SmartM365-M365UserActivity-Inventory.ps1 -Tenant prod -Period D180 -Reports All -Connect
.\Usage\SmartM365-M365UserActivity-Inventory.ps1 -Tenant prod -Reports MailboxUsageDetail,OneDriveUsageAccountDetail -Connect
.\Usage\SmartM365-M365UserActivity-Inventory.ps1 -Tenant test -Period D30 -InteractiveAuth
```

Required Microsoft Graph permission: `Reports.Read.All`.

`SharePointActivityUserDetail` and `TeamsDeviceUsageUserDetail` use the same existing
`Reports.Read.All` permission. They do not require another app registration,
permission, or local configuration value.

For multi-report runs, DATA-ALL and DATA-LAST files are uploaded per report. A
`Report Refresh Date` older than three days is a warning. If a report repeats a
date already archived in another week, its current CSV is still published with
the actual refresh date, but no duplicate snapshot is added to the new week's
WeeklyHistory. The run completes with warnings. A report date that moves backward
or has invalid/mixed values still fails before any CSV publication. Eligible
WeeklyHistory files are published together after the date checks.

Since v1.19 the weekly history is written to `DATA-ALL\M365\Usage\WeeklyHistory`. Up to v1.18
the script wrote it under each run folder (`M365\Usage\<yyyyMMdd_HHmmss>\WeeklyHistory`), so no
consolidated history existed. `Usage/Merge-SmartM365-M365UsageWeeklyHistory.ps1` rebuilds it once,
keeping the earliest snapshot of each report per ISO week, and never overwrites a consolidated file:

```powershell
# Preview (read-only)
.\Usage\Merge-SmartM365-M365UsageWeeklyHistory.ps1 -UsageRootPath \\server\share\...\Tenants\prod\DATA-ALL\M365\Usage
# Consolidate, then remove the per-run WeeklyHistory folders that are fully covered
.\Usage\Merge-SmartM365-M365UsageWeeklyHistory.ps1 -UsageRootPath <same path> -Execute
.\Usage\Merge-SmartM365-M365UsageWeeklyHistory.ps1 -UsageRootPath <same path> -Execute -RemoveRunHistory
```

A per-run folder is removed only when each of its CSVs exists in the consolidated week and an
identical CSV remains in a run folder. Run it outside the 04:50 collector schedule.

For the OneDrive/SharePoint synchronized copy, add `-SynchronizedCopy -RecordedHistoryRootPath <production UNC WeeklyHistory root>`:
its folders are reparse points that the JSON transport refuses, so manifests are written as plain
files with the same bytes as on the share (v1.1). Files already consolidated without a recorded
capture time (interrupted run) recover their provenance when their SHA256 matches the planned snapshot.

Since v1.21 each run replaces the current week (`-OverwriteExisting`), so a week keeps its latest
report: this is what the prepared Workforce history reads (latest file of the week) and what the
published Power BI batches contain. `Merge-SmartM365-M365UsageWeeklyHistory.ps1 -SnapshotPolicy Latest`
(v1.2) rebuilds already consolidated weeks the same way from the timestamped CSVs left in each run
folder; it never creates a week that had no history.

Since v1.20 the collector publishes weekly history with `-UploadChangedFilesOnly`: only the CSVs
copied during the run and a changed manifest are uploaded instead of the whole week. Since SmartM365.Core
1.0.60 (PS5 compatibility module 1.0.44), this mode keeps an `upload.pending` marker in the week folder
until every upload succeeded; the next run that finds it republishes the whole week, so a failed
SharePoint upload is repaired. The automatic history of CSVs exported through the Core, Exchange
on-prem Local Mailboxes 1.49, SPO 0.29 and Teams 0.33 use the same mode.
Core 1.0.77 (PS5 compatibility module 1.0.51) and SharePoint JSON transition 1.0.2
verify each selected WeeklyHistory `manifest.json.txt` against the remote bytes.
An identical manifest is not uploaded again; a missing or different one is published.
The existing name transition and `upload.pending` recovery remain active.

`SmartM365-EXO-Mailboxes-Inventory.ps1` remains useful and is not replaced by `MailboxUsageDetail`: EXO gives mailbox object/stat/archive details from Exchange Online, while Graph Reports gives a period-based usage and quota report suitable for FinOps joins.

## Power BI And Microsoft Fabric Activity Events

`PowerBI/SmartM365-PowerBIFabricActivity-Inventory.ps1` retrieves tenant-wide activity events from the official Power BI admin Activity Events API. Requests are split into one UTC day, limited to the last 28 days, paged with `continuationUri` or `continuationToken`, and kept below 200 requests per hour.

Exports:

| Entity | Latest CSV | Notes |
| --- | --- | --- |
| Detailed events | `M365_PowerBI_Fabric_ActivityEvents.csv` | Useful activity, workload, identity classification, workspace, item, report, semantic model/dataset, dataflow, and capacity fields. Client IP, user agent, and unrelated audit properties are excluded. |
| Per-principal activity | `M365_PowerBI_Fabric_UserActivity.csv` | Aggregates recent activity, active UTC days, view/create/refresh/export counters, and distinct artifact counts by user/application identity. |

Examples:

```powershell
.\PowerBI\SmartM365-PowerBIFabricActivity-Inventory.ps1 -Tenant test -ValidateOnly
.\PowerBI\SmartM365-PowerBIFabricActivity-Inventory.ps1 -Tenant test -LookbackDays 28
.\PowerBI\SmartM365-PowerBIFabricActivity-Inventory.ps1 -Tenant test -FromDate 2026-07-01 -ToDate 2026-07-07 -InteractiveAuth
..\SmartInventory.cmd M365-PowerBIFabricActivity-Inventory
```

Interactive mode requires a Fabric Administrator and delegated `Tenant.Read.All` or `Tenant.ReadWrite.All`. App-only mode uses the configured certificate and requires the Fabric tenant setting **Service principals can access read-only admin APIs** for a security group containing the service principal. The Entra application must not have admin-consent-required Power BI permissions assigned.

The aggregate contains only principals for which an event was returned. After SmartFinOps left-joins the complete license/user population, a missing activity row is a review signal and never automatic proof of non-use.

## Microsoft 365 Copilot Usage

`Usage/SmartM365-CopilotUsage-Inventory.ps1` is a read-only collector for the official Microsoft Graph v1.0 endpoint `GET /copilot/reports/getMicrosoft365CopilotUsageUserDetail`.

It publishes `M365_Copilot_UserUsage.csv` to the tenant `DATA-LAST` folder and keeps timestamped DATA-ALL plus weekly history copies. Report version `v2` is the default because it adds active usage days, prompt counters for all apps and Copilot Chat work/web, and the Microsoft 365 app, Edge, agent, and Copilot Chat work/web last-activity dates. The script falls back to `v1` only when `v2` is unavailable or when `D30` is explicitly requested, because Microsoft documents `D30` for `v1` and `D28` for `v2`. For compatibility with tenants where the versioned route is not yet recognized, the v1 fallback omits the optional `version` function parameter and therefore uses the documented v1 default.

The API returns only users who have a Microsoft 365 Copilot license. The collector never requests or exports prompt text. Version `v1` does not expose prompt counters or active usage days, so the stable v2-oriented CSV columns remain empty after a v1 fallback.

Examples:

```powershell
.\Usage\SmartM365-CopilotUsage-Inventory.ps1 -Tenant test -ValidateOnly
.\Usage\SmartM365-CopilotUsage-Inventory.ps1 -Tenant prod -Period D180 -ReportVersion v2 -Connect
.\Usage\SmartM365-CopilotUsage-Inventory.ps1 -Tenant test -Period D7 -ReportVersion v2 -InteractiveAuth
.\Usage\SmartM365-CopilotUsage-Inventory.ps1 -Tenant test -Period D180 -ReportVersion v2 -Connect -MaxItems 10
```

`-MaxItems` produces suffixed `MAXITEMS` CSV files and does not replace the canonical Power BI file or weekly history. `-ValidateOnly` validates modules, parameters, paths, URI construction, and the empty CSV schema without calling Microsoft Graph.

Required Microsoft Graph permission: `Reports.Read.All`.

Official documentation:

- [Microsoft 365 Copilot usage user detail API](https://learn.microsoft.com/microsoft-365/copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusageuserdetail)
- [Microsoft 365 Copilot usage report](https://learn.microsoft.com/microsoft-365/admin/activity-reports/microsoft-365-copilot-usage)

## SharePoint Online Inventory

`SharePoint/SmartM365-SPO-Inventory.ps1` uses Microsoft Graph for the site/list inventory. When PnP.PowerShell is installed, it also tries the read-only `Get-PnPTenant` capacity call; this requires SharePoint `Sites.FullControl.All` on the same application. If that call is unavailable or denied, the mail and workbook show a clearly labeled license-based capacity estimate from a matching `M365_Licenses_Tenant.csv`. The tenant CSV keeps its measured-capacity fields blank and records the actual collection source. Fields not exposed by Graph are exported as `NotAvailableGraphOnly` instead of failing the run.

Main exports:

| Entity | Latest CSV | Notes |
| --- | --- | --- |
| Sites | `M365_SPO_Sites.csv` | Site URL, title, template, storage quota/usage, activity, owner signal, inactive/orphaned flags, and Graph-only availability markers. |
| Lists | `M365_SPO_Lists.csv` | Lists and libraries per site when Graph can resolve the site. Versioning and size fields are marked unavailable in Graph-only mode. |
| Permissions | `M365_SPO_Permissions.csv` | Owner rows from Graph usage data. Site collection admin enumeration is not required in default mode. |
| External sharing | `M365_SPO_ExternalSharing.csv` | Stable schema with Graph-only availability markers. Tenant-wide anonymous/external sharing link discovery is not available in least-privilege Graph-only mode. |
| Tenant capacity | `M365_SPO_Tenant.csv` | Always exported. Used storage comes from the site inventory. Licensed capacity is collected through `Get-PnPTenant` when PnP.PowerShell is installed and access is granted; `SharePointAdminUrl` is derived automatically and remains an optional override. The license estimate is mail/workbook-only and never replaces measured fields in this CSV. |

The daily SharePoint inventory email attaches one `M365_SPO_Inventory_yyyyMMdd_HHmmss.xlsx` workbook with the five CSV data worksheets plus `Summary` and `Findings`. Immediately after the status, it shows the seven-card KPI grid in the Exchange local-mailbox style: four cards on the first row and three on the second, with the last card spanning two columns. It shows inventoried usage-report sites, distinct sites requiring review, sites without a valid owner, sites above the storage threshold, inactive sites with a dated activity signal, lists above the item threshold, and measured tenant utilization or clearly labeled estimated utilization. The estimated percentage divides storage used by the inventoried sites by license-estimated capacity; the KPI also shows both GB values. Limited inventories show estimated capacity without a utilization percentage. A prioritized action table and the first five sites or lists to review follow the KPI grid, then report freshness, missing activity dates, Graph-enrichment coverage, collection mode, and run details. The email separates unknown activity from dated inactivity; the workbook's existing `Inactive sites` summary still includes sites without an activity date. CSV publication and weekly history remain unchanged. The script-local JSON key `ShowMailLinks` defaults to `false`; set it to `true` to display this run's SharePoint export links. The shared branded footer remains. The mail status is `Critical` only for measured tenant storage utilization of at least 95%, or when at least 10% of inventoried sites lack a valid owner or exceed the configured site quota threshold in a full inventory. Object-level severities remain in the CSV and workbook. Graph mail attachments over 3 MB fail explicitly until a larger-file delivery path is configured.

The SPO mail also shows `Teams-linked sites: matched / inventoried SharePoint sites (percent)` below the seven KPI cards, with separate counts for team main sites and private/shared channel sites. The workbook adds a `Teams site links` worksheet with the matched identities. This ratio uses the Teams and Channels `DATA-LAST` CSVs only when both files have site-link columns, one matching tenant/run, complete Teams scope, all required Graph site identities, and a Teams source run no later than and at most 24 hours before the SPO source run. It excludes OneDrive from the denominator, deduplicates matched SPO sites, and does not infer links from site names or `Team Channel` templates. If qualification fails, the mail states why the ratio is unavailable. The ratio does not affect alert status.

`-MailOnly` builds a `[SNAPSHOT]` mail and a new workbook from the five existing `DATA-LAST/M365_SPO_*.csv` files, without Graph inventory, CSV publication, SharePoint upload, or weekly history. It checks tenant/run consistency, file stability, and the tenant anchor row; an empty supporting CSV must have a write time within 30 minutes of the tenant CSV. Matching license and Teams CSVs are read separately for the capacity estimate and site-link ratio. The email shows the source run date and marks collection options that cannot be established from CSV as unverified. `-DryRun` prepares the workbook without sending mail. `-AlwaysSend` bypasses the daily sent marker for the normal or CSV mode while retaining the exclusive send lock.

Examples:

```powershell
.\SharePoint\SmartM365-SPO-Inventory.ps1 -Tenant prod
.\SharePoint\SmartM365-SPO-Inventory.ps1 -Tenant prod -IncludeOneDrive -InactiveDays 180
.\SharePoint\SmartM365-SPO-Inventory.ps1 -Tenant prod -MaxSites 10 -DryRun
.\SharePoint\SmartM365-SPO-Inventory.ps1 -Tenant prod -AppendHistory -AlwaysSend
.\SharePoint\SmartM365-SPO-Inventory.ps1 -Tenant prod -MailOnly
.\SharePoint\SmartM365-SPO-Inventory.ps1 -Tenant prod -MailOnly -AlwaysSend
.\SharePoint\SmartM365-SPO-Inventory.ps1 -Tenant prod
.\SharePoint\SmartM365-SPO-Inventory.ps1 -Tenant prod -UsePnPTenantCapacity
```

Required PowerShell modules for the baseline inventory: `Microsoft.Graph.Authentication` and `ImportExcel`. When `PnP.PowerShell` is installed, it is also used for the tenant-capacity attempt; `-UsePnPTenantCapacity` requires it explicitly.

Required Graph app-only permissions for the baseline inventory: `Reports.Read.All` and `Sites.Read.All`. `Directory.Read.All` is useful for owner enrichment. No SharePoint tenant administration permission is required for the default run.

When PnP.PowerShell is installed, the collector attempts tenant capacity by default. `-UsePnPTenantCapacity` explicitly requires the PnP path and `-SkipPnPTenantCapacity` skips it. This read-only call requires the SharePoint API application permission `Sites.FullControl.All` with administrator consent; the collector never grants it. Manual `SharePointAdminUrl` configuration is normally unnecessary because the collector derives it from tenant context. If the call is denied, the tenant CSV records blank measured-capacity fields with `CapacitySource=Get-PnPTenantUnauthorized`, while the Graph inventory continues. The mail and workbook then estimate capacity as 1,024 GB plus 10 GB for each enabled purchased seat in the known eligible SKU set (`SPE_E3`, `SPE_E5`, `ENTERPRISEPACK`, `ENTERPRISEPREMIUM`, `STANDARDPACK`, `SHAREPOINTSTANDARD`, `SHAREPOINTENTERPRISE`, `SPB`, `O365_BUSINESS_PREMIUM`, `O365_BUSINESS_ESSENTIALS`). F1/F3 seats and unknown or additional-storage SKUs are excluded. The license CSV must match tenant identity and have a collection time within 48 hours of the SPO run. For a complete inventory, the mail and workbook also show estimated utilization from inventoried site storage divided by the estimated capacity. Both estimates are labeled and excluded from tenant-utilization status thresholds. If the license evidence is unavailable or stale, the mail says `Estimate unavailable` rather than inventing a value.
