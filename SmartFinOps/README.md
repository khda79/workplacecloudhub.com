# SmartFinOps Workplace

SmartFinOps is a read-only decision-support project that applies FinOps principles to workplace technology. Its objective is not limited to reducing spend: it shows what each euro buys, what can be released at contract or renewal, what can be reused, and what remains a review or a signal.

SmartFinOps Workplace V2 values the license decisions already published by SmartM365 SmartInventory. It consumes `DATA-LAST` exports and their source receipts only and writes its results into the same SmartM365 data tree of the profile. Its only API call is the optional SharePoint mirror of its own files (Microsoft Graph, `Sites.Selected`, through SmartM365.Core); it reads nothing from Graph, Azure or any other API, and it never changes a license, mailbox or account.

| Script | Role |
| --- | --- |
| `SmartFinOps-Workplace-Valuation.ps1` | V2 valuation engine: qualified inputs, `FinOps_*` CSVs and their receipt |
| `SmartFinOps-Workplace-Report.ps1` | V2 executive HTML report built from the valuation outputs only |
| `Config/SmartFinOps-TenantContext.ps1` | Configuration, profile key, path tokens and console banners shared by both scripts |

PowerShell 7 is required. Both scripts run daily through the SmartInventory orchestrator (see V2 daily run below).

The former beta analyzer (1.5.0-beta.1, `SmartFinOps-Workplace-Analyze.ps1`) has been removed from this branch. It remains available as the [GitHub beta prerelease](https://github.com/khda79/workplacecloudhub.com/releases/tag/smartfinops-workplace-v1.5.0-beta.1), whose tag keeps its scripts, contracts and documentation.

## Configuration

Local configuration files are created from the committed templates and must not be committed:

- `SmartFinOps.global.local.json.template` -> `SmartFinOps.global.local.json`
- `Config/Tenants/tenant.local.json.template` -> `Config/Tenants/<profile>.local.json` (optional per-profile overrides)
- `Config/SmartFinOps-Contracts.local.json.template` -> `Config/SmartFinOps-Contracts.local.json` (contracts, see below)

At load, missing keys are added from the matching template without overwriting local values.

| Key | Meaning |
| --- | --- |
| `RepositoryRootPath` | Repository root, used to find the SmartM365 context (`{{RepositoryRootPath}}` by default). |
| `SmartM365LatestCsvFolderPath` | SmartInventory `DATA-LAST`: sources, current FinOps outputs and the report page. `__USE_SMARTM365__` resolves it through the SmartM365 context of the profile and stops on any resolution failure. |
| `SmartM365DataAllRootPath`, `SmartM365LogAllRootPath` | SmartM365 `DATA-ALL` (FinOps run history) and `LOG-ALL` (FinOps logs), resolved the same way. |
| `SmartM365TenantKey` | Organization TenantKey carried by the source receipts; `__USE_SMARTM365__` resolves it the same way. |
| `EnableSharePointUpload` | `true`, `false` or `__USE_SMARTM365__` (default: the SmartM365 setting of the profile). Site, library, folder and app-only identity always come from the SmartM365 context. |
| `Valuation` | `ContractsPath`, `FreshnessDailyHours` (48), `FreshnessWeeklyDays` (8), `ArchiveRetentionRuns` (30). |
| `Report` | `ShowbackMinimumGroupUsers` (5); daily email: `EnableDailyEmail` (true), `DailyEmailTo` and `DailyEmailFrom` (`__USE_SMARTM365__`: the SmartM365 `To` and `From` of the profile). |

`-Tenant` is the profile key. A profile file names it `ProfileKey`, which must match `-Tenant`. Since tenant context 1.3.0 a profile file written with the former `TenantKey` entry is migrated in place to `ProfileKey` before template synchronization; a file holding a different `ProfileKey` and `TenantKey` is refused. `{{TenantKey}}` remains accepted as an alias of `{{ProfileKey}}` in path values. Since tenant context 1.4.0 `WorkspaceRootPath`, `TenantDataRootPath` and the former `Output` fallback no longer exist: SmartFinOps writes only into the SmartM365 data tree of the profile. Beta-only keys (`PriceModel`, `StaleUserDays`, `StaleDeviceDays`, `MaxSourceAgeHours`, `LicenseCapacityWarningPercent`, `ReportTitle`, `PriceBaselinePath`, `RetentionMaxCSV`, `RetentionMaxLogs`) are no longer read and can be removed from local files.

## V2 valuation engine (lot V2.1)

`SmartFinOps-Workplace-Valuation.ps1` 2.2.4 is the V2 engine. PowerShell 7 only, read-only on the tenant; its only API call is the optional SharePoint mirror of its outputs.

```powershell
.\SmartFinOps-Workplace-Valuation.ps1 -Tenant prod -ValidateOnly   # qualify sources and contracts, compute, write nothing
.\SmartFinOps-Workplace-Valuation.ps1 -Tenant prod
```

`-Tenant` is the SmartM365 profile key: it selects the SmartM365 data tree of that profile. Source receipts and rows carry the organization TenantKey (OrganizationKey-EnvironmentKey), which can differ from the profile: it comes from `-SmartM365TenantKey`, else from `SmartM365TenantKey` in the local configuration, else (empty or `__USE_SMARTM365__`) from the SmartM365 tenant context of the profile, with the same explicit failure rules as DATA-LAST. The FinOps receipt (contract 1.1) records both `ProfileKey` and `TenantKey`; the FinOps CSV rows carry the TenantKey, and the report checks both. DATA-LAST comes from `-SmartM365LatestCsvFolderPath`, else from `SmartM365LatestCsvFolderPath` in the local configuration (DATA-ALL and LOG-ALL likewise, from `-SmartM365DataAllRootPath` and `-SmartM365LogAllRootPath`). When a setting is empty or `__USE_SMARTM365__`, it is resolved through the SmartM365 tenant context; if that resolution fails (context missing, error, empty path) the script stops with the reason and its position, and never falls back to a default path. A resolved path that still contains a `{{token}}` is refused, and every path handed to .NET is a plain absolute file-system path, so the scripts also run from a UNC share. Both V2 scripts print the WorkplaceCloudHub startup banner and, also on failure, the completion banner (status, duration, warnings, errors, CSV files, log), which is appended to the run log.

The engine never recomputes a decision. It values what SmartInventory already published and qualifies every input with its producer source receipt (`Status=Completed`, tenant, SHA256 and row count per file, optional files `Published=true`, no publication in progress, completed within `Valuation.FreshnessDailyHours`, default 48 h):

- the Licences report (`M365_Licenses_Report*.csv` and `M365_Licenses_ReportSnapshot.json.txt`): one `SnapshotId` on every row, snapshot `ReceiptRun.RunId` and per-SKU `ReceiptRunId` equal to the Licences receipt `RunId`, snapshot schema 2;
- the per-SKU decisions (`M365_Licenses_ReportSkuCandidates.csv`, `M365_Licenses_ReportSkuCapacity.csv`) and `M365_Licenses_Users.csv`;
- `M365_License_Prices.csv` and, when published, `Windows_ESU_Prices.csv`;
- `AD_Users_AllDomains.csv` and `M365_Users_Active.csv` for showback only.

An unqualified input is `Not measured`: no amount, never zero, never an older value. A missing price is `Not valued`. A source marked provisional in the Licences report makes `Provisional` only the levers that use it: `M365_OneDrive_Usage.csv` affects L4 (E3 to F3), `AD_Users_AllDomains.csv` the user levers L1 to L6, and any other provisional source every valued lever. `FinOps_Sources.csv` (`ProvisionalLevers`) and the receipt record which source made which lever provisional. Accepted partial AD coverage (`PartialAccepted`) lists the missing domains; their users are `Not measured` in the showback. Since 2.2.4 the `Provisional amounts` warning of the log gives one `<source> -> <levers>: <reason>` entry per provisional source: the reason of the Licences report, else the detail published in `FinOps_Sources.csv` (for example `Provisional in the Licences report`), else `source status Provisional`; never an empty reason. The CSV outputs are unchanged.

### Value classes and levers

`R` is realizable at contract or renewal, `Rc` conditionally realizable (outside the R total), `E` avoided cost or exposure, `S` a signal that is never monetized.

| Lever | Source | Class |
| --- | --- | --- |
| L1, L2, L3 | `ReportCandidates` Recovery: disabled account, eligible shared mailbox under 50 GB, enabled account without M365 activity in 90 days (own KPI) | R |
| L4 | `ReportCandidates` E3 to F3 review: E3 price minus F3 price | Rc |
| L5 | `ReportSkuCandidates` disabled account. `Recommended` is R; `Referenced` keeps no amount (already counted); `Review` with an unknown mailbox state is `Not measured`; other reviews and `Blocked by dependency` are signals | R / S |
| L6 | `ReportSkuCandidates` Copilot without activity (`Recommended` R, `Review` S); Copilot usage not measured | R / S |
| L7 | `ReportSkuCapacity` unassigned units after reserve (excluded SKUs skipped) | R |
| L8 | Users waiting for capacity (`ReportSummary`) served from the freed pool before any reduction | E |
| L9, L10 | Over-assigned SKU; Dynamics 365 active users without product telemetry | S |
| L11, L12 | Assignment issues; mailbox/AD gap metrics of the report (one row per metric, never summed) | S |
| L13 | Windows 10 ESU: no device population is published by SmartInventory yet | E, `Not measured` |

Prices come from the single EUR row of `M365_License_Prices.csv` effective on the snapshot date; amounts are an annual run-rate (monthly price x 12) with the date from which they are realizable. The SKU of a suite candidate is the SKU of its `Referenced` per-SKU row, or the only assigned SKU of that report product; otherwise the row is `Not valued`.

### Capped realization

Per SKU: pool = unassigned units - reserve + R candidates. Users waiting for capacity are served first (E). The rest is realizable up to `EnabledUnits - MinimumUnits` of the contract; units beyond the cap are `Capped`, in a fixed order (capacity, L1, L5, L2, L6, L3, then user ID).

A realized saving is reported when `EnabledUnits` drops below the level of the previous qualified run. The baseline is the first run of that level: realized units = min(candidate licenses valued R during that level and no longer assigned + baseline capacity units, capped at the baseline realizable units; the drop). The rest of the drop is `not attributed`.

### Outputs

`DATA-LAST` holds `FinOps_Sources.csv`, `FinOps_Opportunities.csv`, `FinOps_SkuValuation.csv`, `FinOps_Showback.csv` (AD domain, country, entity, department), `FinOps_Summary.csv`, `FinOps_Realization.csv` and `FinOps_Controls.csv`, each sorted by its key, plus `SmartFinOps_Workplace-Valuation.current.json.txt`. Every row carries `TenantKey`, `FinOpsRunId`, `SnapshotId` and the Licences `ReceiptRunId`; valued rows also carry the price row key and pricing receipt run. `DATA-ALL\FinOps\Workplace\<FinOpsRunId>\` keeps the same files and receipt (the run history, read for realization baselines); the run log is `LOG-ALL\SmartFinOps-Workplace-Valuation\SmartFinOps-Workplace-Valuation-<yyyy-MM-dd_HH-mm-ss>.log`, named in the receipt (`LogFile`). `Valuation.ArchiveRetentionRuns` (default 30) limits the history; see the retention rules below.

`FinOps_Summary.csv` carries the totals (`TOTAL_R`, `TOTAL_R_CONFIRMED`, `TOTAL_R_DISABLED_AND_SHARED`, `TOTAL_R_ENABLED_INACTIVE`, `TOTAL_R_COPILOT`, `TOTAL_R_CAPACITY`, `TOTAL_RC_E3_TO_F3`, `TOTAL_E_REUSE`, `TOTAL_E_ESU`), one `DETAIL|class|lever|status` row per group and one `GAP|section|metric` row per gap metric. Since 2.1.1 each valued total is also broken down, so that a report never has to add amounts: `PRICE|<total>|Estimated|Contract|Unknown` (price origin from `IsEstimated`; an E3 to F3 difference is estimated when either price is), `FORECAST|<total>|<RealizableFrom>` (date or `Renewal unknown`) and `SHOWBACK|<dimension>|<total>|<value>`; `SHOWBACK|<dimension>|USERS|<value>` gives the distinct users with at least one opportunity row for that value. E3 to F3 rows now carry the contract terms of their SKU (`ContractKey`, `RenewalDate`, `RealizableFrom`). Since 2.2.3 `TOTAL_R_CONFIRMED` (levers L1, L2, L5, L6, L7) is the R without the enabled accounts without activity, published with the same breakdowns, so that a consumer shows confirmed R and `TOTAL_R_ENABLED_INACTIVE` separately without subtracting.

Equivalence controls G1 to G7 compare the FinOps rows with the Licences counts (candidates by type and product, per-SKU recommendations, summary recovery counts, capacity counters, no double counting, showback totals and, since 2.1.1, every Summary breakdown against its total: units and amount, plus users for showback values; since 2.2.3, G7: `TOTAL_R` units and amount equal `TOTAL_R_CONFIRMED` plus `TOTAL_R_ENABLED_INACTIVE`, users being distinct per total). A failed control fails the run: the current files are kept but the receipt is `Failed`. **Every consumer must check the receipt (`Status=Completed`, `TenantKey`, SHA256 of each file) before reading a `FinOps_*` CSV.**

### Contracts

`Config/SmartFinOps-Contracts.local.json` (private, created from `Config/SmartFinOps-Contracts.local.json.template`) maps SKUs to contracts: `ContractKey`, `RenewalDate` (`yyyy-MM-dd`, empty means renewal unknown), `ReductionPolicy` (`AtRenewal`, `Anniversary`, `Monthly`, `Unknown`) and, per SKU, `MinimumUnits` and `ReserveUnits` (default `Defaults.ReserveUnitsPercent`, 0). Prices stay in the LicensePricing private override. Paid SKUs without a contract are added to `UnmappedSkus`; no date is invented.

Tests: `Tests\SmartFinOps-Workplace-Valuation.Tests.ps1` (synthetic data, isolated runtime, real SmartM365.Core receipts).

## V2 executive report (lot V2.2)

`SmartFinOps-Workplace-Report.ps1` 2.5.1 turns the valuation outputs into one self-contained HTML page for management. PowerShell 7 only, read-only. It is a pure consumer: it reads only the seven `FinOps_*` CSVs and the valuation receipt, and computes no decision and no amount; it displays, sorts and groups what the engine (2.1.4 to 2.2.x) published.

```powershell
.\SmartFinOps-Workplace-Valuation.ps1 -Tenant prod
.\SmartFinOps-Workplace-Report.ps1 -Tenant prod              # DATA-LAST\SmartFinOps_Workplace-Report.html
.\SmartFinOps-Workplace-Report.ps1 -Tenant prod -ValidateOnly # qualify and build, write nothing
```

Before reading anything the report qualifies the run: receipt `Owner=SmartFinOps-Receipt` 1.1 from engine 2.1.4 to 2.2.x, `Status=Completed`, same TenantKey and profile, exactly the seven files with their SHA256 and row counts (hashed and parsed from the same bytes), exact headers, one `FinOpsRunId`, `SnapshotId` and Licences `ReceiptRunId` on every row, and no failed equivalence control. Otherwise the page is "Report not available" with the reason and no figure; it replaces the current report so that no older figure stays visible, and the script exits with code 1. When the Licences report was not measured the report is still produced, with a status banner and every figure `Not measured`.

Layout (since 2.5.0, SmartM365 GUI charter: background `#F5F8FB`, white cards bordered `#DDE7F0`, accent `#0078D4`): a header card with the WorkplaceCloudHub logo and a status bar (overall certainty: Final, Provisional or Licences report not measured, and the number of qualified, provisional and not measured sources), then an always visible summary:

- four cards, each a published `FinOps_Summary` total with its certainty and its units and users (since 2.5.1 a total counted in license units without people, such as the reuse for users waiting for capacity, shows `Units: N` only, never `Users: 0`; the same rule applies to the Financial potential tables): R confirmed (`TOTAL_R_CONFIRMED`, engine 2.2.3 or later; an older run shows it as Not measured, never as a difference), R to confirm (`TOTAL_R_ENABLED_INACTIVE`, enabled accounts without activity, shown apart from the confirmed R), conditional E3 to F3 and cost avoidance, which are never added to R; the published `TOTAL_R` is quoted below the cards as the sum checked by control G7;
- the top 3 decisions: the first three rows of Priority decisions as published, each with its certainty badge (Final, To confirm, Conditional - outside R total for E3 to F3, Cost avoidance or the status), in the page and in the mail;
- the license units of levers L1 to L7 by status.

The detail follows in folding sections (HTML `details`, no script), all closed except Financial potential: financial potential (totals, price origin, by lever and status, by SKU), priority decisions, decisions not in the totals (Review, Blocked, Not measured, Not valued, Capped, signals, gap metrics), capacity, renewals and forecast (`RealizableFrom`, renewal unknown), realization when a baseline exists, showback (AD domain, country, entity, department, each folding separately), how to read this report (assumptions, certainty and status labels, levers) and the technical appendix (report qualification, sources, equivalence controls). Levers are named by their business label first, code in brackets. Statuses are always visible: an empty amount is shown as its status, never as zero; catalogue estimates and contract prices are labelled.

Confidentiality: aggregates only. No user name, UPN or user ID is rendered (detail stays in the CSVs); free text that quotes an e-mail address or a GUID is masked. In the showback, values with fewer than `Report.ShowbackMinimumGroupUsers` users (default 5) are combined into "Other (groups under N users)"; when that row alone would stay under the threshold, the smallest other values join it, so no value can be deduced by subtraction. The page is self-contained (inline CSS and SVG, the WorkplaceCloudHub logo embedded as a 16 KB PNG data URI derived from `WorkplaceCloudHub-lockup-WPF.png`, no script, no external resource, every value HTML-encoded) and deterministic: the same inputs give the same HTML except the generation time. The page is copied into `DATA-ALL\FinOps\Workplace\<FinOpsRunId>\` when that run is in the history; the log is `LOG-ALL\SmartFinOps-Workplace-Report\SmartFinOps-Workplace-Report-<yyyy-MM-dd_HH-mm-ss>.log`.

Client logo (since 2.5.1): the report reuses the SmartM365 mail branding of the profile, without any new key: `MailBrandingEnabled`, `MailClientName`, `MailClientLogoPath` and `MailClientLogoMaxKB`, read through the SmartM365 context of the profile with `Get-SmartM365MailBrandingConfig` and embedded with `ConvertTo-SmartM365MailLogoDataUri` (SmartM365.Core) as a data URI at the top left of the header, `MailClientName` as alternative text; the WorkplaceCloudHub logo stays on the right (below the title on a phone). Only a PNG, JPEG or GIF whose content matches its type is shown. No branding configured (no branding, `MailBrandingEnabled` false or empty `MailClientLogoPath`) gives a page without client logo and an INFO line. Fail-safe: a configured logo that is missing, too large, unreadable or of a refused type, or an unavailable SmartM365 context, gives a page without client logo and a `Client logo not shown: <reason>` warning in the log that never changes the exit code. `-ValidateOnly` does not resolve the branding, and `-NoClientLogo` never reads it (the demo generator always uses it).

Tests: `Tests\SmartFinOps-Workplace-Report.Tests.ps1` (synthetic FinOps fixtures, input contract checked against the engine output specification).

## V2 daily run (lot V2.4a)

The SmartInventory orchestrator runs both V2 scripts every day for the prod profile, as two jobs of the `FinOps` group in `SmartM365/SmartInventory/Orchestrator/Orchestrator-Jobs.json.template`. Their `ScriptPath` points outside SmartInventory (`..\..\SmartFinOps\...`): the scripts are launched directly, without a SmartInventory wrapper, and the orchestrator appends `-Tenant <profile>`.

| Job | Schedule | Depends on | Rule |
| --- | --- | --- | --- |
| `SmartFinOps-Workplace-Valuation` | Daily 05:45 | `M365-Licences-Inventory`, `M365-LicensePricing-Inventory`, `AD-Inventory`, `M365-ActiveUsers-Inventory` | `LatestOccurrence`, wait up to 1080 min |
| `SmartFinOps-Workplace-Report` | Daily 05:45 | `SmartFinOps-Workplace-Valuation` | `LatestOccurrence`, wait up to 1110 min |

- `LatestOccurrence` requires today's occurrence of every dependency, which is why the jobs start after Licences (04:45) and LicensePricing (05:30). A Licences report from the previous day is never valued as today's. When a dependency fails, the jobs wait for a successful recovery run; otherwise they end `BlockedDependencyTimeout` and FinOps does not run that day. The rule applies to all four dependencies.
- On weekdays Licences ends around 07:50 and both jobs finish before `WorkplaceEvidence-Prepare` (08:30). On Saturday Licences ends around 18:00 after the full Friday EXO inventory; the 1080-minute wait covers it (until 23:45).
- Elected, `Graph` (`Sites.Selected`, for the SharePoint mirror) and `SharedRuntime`, PowerShell 7, shared `ConcurrencyKey` `SmartFinOps-Workplace`, timeouts 90 and 30 min (deliberately below the 1380-minute daily policy, reported by its audit), one retry. Valuation has `ContinueOnError=false`: a failed valuation blocks the report.
- Outputs go to the SmartM365 data tree of the profile (see below), shared by every orchestrator server.

Exit codes of both scripts follow their completion banner, as the orchestrator expects:

| Code | Valuation | Report |
| --- | --- | --- |
| `0` | Completed, no warning | Page built from final figures |
| `3` | Completed with warnings: Licences not measured, a required source not measured, provisional amounts, disabled-account licenses not measured | Page built, but Licences not measured or figures provisional |
| `1` | Run error or failed equivalence control (receipt `Failed`) | Report not available, or run error |

Both jobs end with code 3 while a source of the Licences report is provisional. Since Licences 1.51, `M365_OneDrive_Usage.csv` is qualified by the UserActivity source receipt (or left unqualified with its reason) and no longer makes L4 provisional. ESU prices are optional and do not raise a warning. The first log line of each script (`Starting SmartFinOps Workplace ... v<version>`) is the success evidence the orchestrator looks for in the job log.

Launchers:

- direct: `Start-SmartFinOps-Workplace-Prod.cmd` (Valuation then Report, see below), or `Start-SmartFinOps-Workplace-Valuation-Prod.cmd` and `Start-SmartFinOps-Workplace-Report-Prod.cmd` separately (this folder);
- through the orchestrators: `SmartM365\SmartInventory\SmartInventory.cmd SmartFinOps-Workplace-Valuation -Request` and `SmartInventory.cmd SmartFinOps-Workplace-Report -Request` (add `-IncludeDependencies` to collect the sources first).

## Data location and SharePoint mirror (lot V2.4c)

Since engine 2.2.0, report 2.3.0 and tenant context 1.4.0 (upload context fixed in 2.2.1, 2.3.1 and 1.4.1) SmartFinOps writes into the SmartM365 data tree of the profile, like the SmartInventory collectors, and never into its code folder (except the private contracts file, which records `UnmappedSkus`):

| Content | Location |
| --- | --- |
| Seven `FinOps_*.csv`, `SmartFinOps_Workplace-Valuation.current.json.txt`, `SmartFinOps_Workplace-Report.html` | `DATA-LAST\` |
| Run history (files, receipt and report page of each run) | `DATA-ALL\FinOps\Workplace\<FinOpsRunId>\` |
| Logs | `LOG-ALL\SmartFinOps-Workplace-Valuation\` and `LOG-ALL\SmartFinOps-Workplace-Report\` |

The three roots are read from the SmartM365 context of the profile (`LatestCsvFolderPath`, `DataAllRootPath`, `LogAllRootPath`) unless overridden. A run stops when a root keeps a `{{token}}`, is not called `DATA-LAST`, `DATA-ALL` or `LOG-ALL`, does not exist (SmartFinOps never creates them), is not a sibling of the other two under the same tenant data root, or lies under the SmartM365 `Output` fallback. `FinOps_` and `SmartFinOps_` names do not collide with SmartInventory files, and the SmartWorkplaceIntelligence preparation copies only the files of its own contract.

History retention keeps the `Valuation.ArchiveRetentionRuns` most recent runs. It only ever removes a folder directly under `DATA-ALL\FinOps\Workplace\` whose name is exactly `yyyyMMddTHHmmssfffZ-xxxxxxxx` (lowercase hexadecimal) and that holds a SmartFinOps receipt of the same TenantKey and run; never the current run, never a realization baseline, never anything in `DATA-LAST` or `LOG-ALL`. Each removal is logged with its path. Logs are not purged. A run without a measured Licences report removes nothing.

With `EnableSharePointUpload` (the SmartM365 setting by default), both scripts mirror their files to the same relative paths under the SmartM365 SharePoint data root (`<target>/DATA/DATA-LAST`, `DATA/DATA-ALL`, `DATA/LOG-ALL`, as for SmartM365, a legacy `.../CSV` target being read as `.../DATA`), with the SmartM365.Core upload helper and the SmartM365 app-only identity (`Sites.Selected`). Before any upload the scripts initialize the SmartM365 context of the same profile (`Initialize-SmartM365TenantContext`), which gives the SmartM365.Core helpers the Graph identity and SharePoint target of that profile, as for a SmartM365 script. A SmartM365 context of another profile already loaded in the session is refused, the context must report the same TenantKey, and SmartFinOps never creates a SmartM365 runtime JSON: a profile without its SmartM365 global and profile files stops with an explicit error.

- the engine uploads the run history first, then a `Publishing` current receipt, the seven current CSVs, and the `Completed` receipt last; when a current file fails to upload, the completed receipt is withheld, so the SharePoint copy stays disqualified;
- a failed run uploads its `Failed` receipt; the report uploads its page (current and history) and both scripts upload their log (without the completion banner, appended locally afterwards);
- a run removed by retention is also removed from SharePoint (site recycle bin), only for the same run folder;
- an incomplete SmartM365 SharePoint configuration or any upload failure is a warning: the local files stay valid and the run ends with code 3, never as a plain success.

Runs written before 2.2.0 under `SmartFinOps\Data\<profile>` are not migrated or read; the first run in the SmartM365 data tree starts a new realization baseline.

## Daily email and combined launcher (lot V2.5)

`SmartFinOps-Workplace-Report.ps1` (since 2.4.0) sends one daily email per profile after the page is written (and uploaded):

- once per Europe/Paris day, with the same gate as the SmartM365 Licences summary: an exclusive state file `DATA-LAST\SmartFinOps_Workplace-Report_DailyEmail_SendState.json.txt`, `Pending` before sending and `Sent` after; `-ForceEmail` sends again; never in `-ValidateOnly`;
- through the SmartM365 mail transport of the profile (`Send-SmartM365Mail` after the SmartM365 context of the profile is initialized): `SendMailMode` Graph, SMTP or Both, the maintenance recipient policy, the `[SMART 365] - [<tenant>] - ` subject prefix of the profile, and the HTML copy of the mail in `LOG-ALL\SmartFinOps-Workplace-Report\`. No job declares `Mail.Send` (SmartM365 convention), so the orchestrator configuration does not change;
- recipients: `Report.DailyEmailTo` and `Report.DailyEmailFrom`, or the SmartM365 `To` and `From` of the profile with `__USE_SMARTM365__` (default); `Report.EnableDailyEmail` (default true) turns it off;
- subject `SmartFinOps Workplace - Daily valuation`, with `- provisional figures`, `- Licences report not measured` or `- report not available` when relevant;
- body (since 2.5.0): the summary of the page, from the same published values (nothing added): the four cards (R confirmed, R to confirm, conditional E3 to F3, cost avoidance) with their certainty, the published R total and the ESU status, the top 3 decisions and the source states in one line, English labels, inline styles only (the SmartM365 transport wraps it in the SmartM365 mail frame), the SharePoint link of the page when the upload succeeded, and the page attached. Before sending, the body and the attachment are checked: an e-mail address, a UPN or a GUID blocks the mail;
- Licences not measured or report not available: a status email with the reason and no amount (no attachment when the report is not available); never a zero or an older value;
- a mail problem (transport error, missing recipients, SmartM365 context unavailable, identifier found) is a warning: the page stays written and the run ends with code 3. After a failed send the day stays `Pending`, so no second mail is sent automatically that day.

`Start-SmartFinOps-Workplace-Prod.cmd` runs the valuation, then the report only when the valuation ended with 0 or 3, so that no report is built on an older valuation. Its exit code is the most severe of the two (1 > 3 > 0). The separate launchers and the orchestrator jobs are unchanged.

Tests: `Tests\SmartFinOps-Workplace-Report.Tests.ps1` (daily email with a simulated transport: one mail per day, `-ForceEmail`, Europe/Paris day, status mails, no identifier, send failure) and `Tests\SmartFinOps-Workplace-Launcher.Tests.ps1` (combined launcher with stub scripts). No test sends a real mail.

## Demo report

`Demo/SmartFinOps-Workplace-DemoReport.ps1` builds a demo of the executive report from synthetic data only, with the real engine and report:

```powershell
.\Demo\SmartFinOps-Workplace-DemoReport.ps1 -OutputPath C:\Temp\SmartFinOps-Demo
```

- `Demo/SmartFinOps-Workplace-DemoFixture.ps1` writes a fictional SmartInventory `DATA-LAST` (organization `contoso.invalid`, 240 users, four AD domains, all collected) and publishes it with complete SmartM365.Core source receipts, so the figures are Final: a baseline plateau, then the current day after a reduction, so the page also shows a realized saving. The scenario covers Valued, Not valued (USD price only), Review, Blocked, Not measured (Windows 10 ESU, Dynamics 365 usage), Capped and Reused, and every run ends with exit code 0.
- The generator copies the engine, the report, the tenant context, the templates and the tracked SmartM365.Core files into a temporary runtime, with SharePoint upload and the daily email disabled; no local configuration, profile, export or tenant is read. Exit codes 0 and 3 of the engine and the report are accepted; 1 or a "Report not available" page fails the run.
- The page gets a visible DEMO banner, then a leak scan refuses e-mail addresses, GUIDs, IP addresses, file-system and UNC paths, user, host and domain names of the computer, the private terms of the local Git guard, any domain other than `contoso.invalid`, external URLs, scripts and any image other than the WorkplaceCloudHub logo. The report runs with `-NoClientLogo` (generator 1.0.2), so a SmartM365 mail branding configured on the build computer never puts a client logo in the demo.
- Outputs: `SmartFinOps_Workplace-Report.demo.html` and `SmartFinOps_Workplace-Report.demo.manifest.json` (versions, exit codes, SHA256) in `-OutputPath`, which must lie outside this folder. They are generated files: never commit them. The WorkplaceCloudHub website rebuilds its demo from `origin/main` with this script.

Tests: `Tests\SmartFinOps-Workplace-DemoReport.Tests.ps1` (generation, manifest, DEMO banner, statuses shown, leak scanner cases, output folder guard).
