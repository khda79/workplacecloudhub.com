# SmartFinOps Workplace

SmartFinOps is a read-only decision-support project that applies FinOps principles to workplace technology. Its objective is not limited to reducing spend: it shows what each euro buys, what can be released at contract or renewal, what can be reused, and what remains a review or a signal.

SmartFinOps Workplace V2 values the license decisions already published by SmartM365 SmartInventory. It consumes `DATA-LAST` exports and their source receipts only; it does not connect to Microsoft Graph, Azure or any other API, and it never changes a license, mailbox or account.

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
| `WorkspaceRootPath`, `RepositoryRootPath` | Roots used by the path tokens (`{{SmartFinOpsRootPath}}` and `{{RepositoryRootPath}}` by default). |
| `TenantDataRootPath` | Output folder of the profile, `{{WorkspaceRootPath}}\Data\{{ProfileKey}}` by default. |
| `SmartM365LatestCsvFolderPath` | SmartInventory `DATA-LAST`; `__USE_SMARTM365__` resolves it through the SmartM365 context of the profile and stops on any resolution failure. |
| `SmartM365TenantKey` | Organization TenantKey carried by the source receipts; `__USE_SMARTM365__` resolves it the same way. |
| `Valuation` | `ContractsPath`, `FreshnessDailyHours` (48), `FreshnessWeeklyDays` (8), `ArchiveRetentionRuns` (30). |
| `Report` | `ShowbackMinimumGroupUsers` (5). |

`-Tenant` is the profile key. A profile file names it `ProfileKey`, which must match `-Tenant`. Since tenant context 1.3.0 a profile file written with the former `TenantKey` entry is migrated in place to `ProfileKey` before template synchronization; a file holding a different `ProfileKey` and `TenantKey` is refused. `{{TenantKey}}` remains accepted as an alias of `{{ProfileKey}}` in path values. Beta-only keys (`PriceModel`, `StaleUserDays`, `StaleDeviceDays`, `MaxSourceAgeHours`, `LicenseCapacityWarningPercent`, `ReportTitle`, `PriceBaselinePath`, `RetentionMaxCSV`, `RetentionMaxLogs`) are no longer read and can be removed from local files.

## V2 valuation engine (lot V2.1)

`SmartFinOps-Workplace-Valuation.ps1` 2.1.5 is the V2 engine. PowerShell 7 only, read-only, no API connection.

```powershell
.\SmartFinOps-Workplace-Valuation.ps1 -Tenant prod -ValidateOnly   # qualify sources and contracts, compute, write nothing
.\SmartFinOps-Workplace-Valuation.ps1 -Tenant prod
```

`-Tenant` is the SmartM365 profile key: it names `Data\<profile>` and its archive. Source receipts and rows carry the organization TenantKey (OrganizationKey-EnvironmentKey), which can differ from the profile: it comes from `-SmartM365TenantKey`, else from `SmartM365TenantKey` in the local configuration, else (empty or `__USE_SMARTM365__`) from the SmartM365 tenant context of the profile, with the same explicit failure rules as DATA-LAST. The FinOps receipt (contract 1.1) records both `ProfileKey` and `TenantKey`; the FinOps CSV rows carry the TenantKey, and the report checks both. DATA-LAST comes from `-SmartM365LatestCsvFolderPath`, else from `SmartM365LatestCsvFolderPath` in the local configuration. When that setting is empty or `__USE_SMARTM365__`, it is resolved through the SmartM365 tenant context; if that resolution fails (context missing, error, empty path) the script stops with the reason and its position, and never falls back to a default path. A resolved path that still contains a `{{token}}` is refused, and every path handed to .NET is a plain absolute file-system path, so the scripts also run from a UNC share. Both V2 scripts print the WorkplaceCloudHub startup banner and, also on failure, the completion banner (status, duration, warnings, errors, CSV files, log), which is appended to the run log.

The engine never recomputes a decision. It values what SmartInventory already published and qualifies every input with its producer source receipt (`Status=Completed`, tenant, SHA256 and row count per file, optional files `Published=true`, no publication in progress, completed within `Valuation.FreshnessDailyHours`, default 48 h):

- the Licences report (`M365_Licenses_Report*.csv` and `M365_Licenses_ReportSnapshot.json.txt`): one `SnapshotId` on every row, snapshot `ReceiptRun.RunId` and per-SKU `ReceiptRunId` equal to the Licences receipt `RunId`, snapshot schema 2;
- the per-SKU decisions (`M365_Licenses_ReportSkuCandidates.csv`, `M365_Licenses_ReportSkuCapacity.csv`) and `M365_Licenses_Users.csv`;
- `M365_License_Prices.csv` and, when published, `Windows_ESU_Prices.csv`;
- `AD_Users_AllDomains.csv` and `M365_Users_Active.csv` for showback only.

An unqualified input is `Not measured`: no amount, never zero, never an older value. A missing price is `Not valued`. A source marked provisional in the Licences report makes `Provisional` only the levers that use it: `M365_OneDrive_Usage.csv` affects L4 (E3 to F3), `AD_Users_AllDomains.csv` the user levers L1 to L6, and any other provisional source every valued lever. `FinOps_Sources.csv` (`ProvisionalLevers`) and the receipt record which source made which lever provisional. Accepted partial AD coverage (`PartialAccepted`) lists the missing domains; their users are `Not measured` in the showback.

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

`Data\<TenantKey>\` holds `FinOps_Sources.csv`, `FinOps_Opportunities.csv`, `FinOps_SkuValuation.csv`, `FinOps_Showback.csv` (AD domain, country, entity, department), `FinOps_Summary.csv`, `FinOps_Realization.csv` and `FinOps_Controls.csv`, each sorted by its key, plus `SmartFinOps_Workplace-Valuation.current.json.txt`. Every row carries `TenantKey`, `FinOpsRunId`, `SnapshotId` and the Licences `ReceiptRunId`; valued rows also carry the price row key and pricing receipt run. `Archive\<FinOpsRunId>\` keeps the same files and the run log. `Valuation.ArchiveRetentionRuns` (default 30) limits the archive, but a run that is the baseline of a realization is never purged.

`FinOps_Summary.csv` carries the totals (`TOTAL_R`, `TOTAL_R_DISABLED_AND_SHARED`, `TOTAL_R_ENABLED_INACTIVE`, `TOTAL_R_COPILOT`, `TOTAL_R_CAPACITY`, `TOTAL_RC_E3_TO_F3`, `TOTAL_E_REUSE`, `TOTAL_E_ESU`), one `DETAIL|class|lever|status` row per group and one `GAP|section|metric` row per gap metric. Since 2.1.1 each valued total is also broken down, so that a report never has to add amounts: `PRICE|<total>|Estimated|Contract|Unknown` (price origin from `IsEstimated`; an E3 to F3 difference is estimated when either price is), `FORECAST|<total>|<RealizableFrom>` (date or `Renewal unknown`) and `SHOWBACK|<dimension>|<total>|<value>`; `SHOWBACK|<dimension>|USERS|<value>` gives the distinct users with at least one opportunity row for that value. E3 to F3 rows now carry the contract terms of their SKU (`ContractKey`, `RenewalDate`, `RealizableFrom`).

Equivalence controls G1 to G6 compare the FinOps rows with the Licences counts (candidates by type and product, per-SKU recommendations, summary recovery counts, capacity counters, no double counting, showback totals and, since 2.1.1, every Summary breakdown against its total: units and amount, plus users for showback values). A failed control fails the run: the current files are kept but the receipt is `Failed`. **Every consumer must check the receipt (`Status=Completed`, `TenantKey`, SHA256 of each file) before reading a `FinOps_*` CSV.**

### Contracts

`Config/SmartFinOps-Contracts.local.json` (private, created from `Config/SmartFinOps-Contracts.local.json.template`) maps SKUs to contracts: `ContractKey`, `RenewalDate` (`yyyy-MM-dd`, empty means renewal unknown), `ReductionPolicy` (`AtRenewal`, `Anniversary`, `Monthly`, `Unknown`) and, per SKU, `MinimumUnits` and `ReserveUnits` (default `Defaults.ReserveUnitsPercent`, 0). Prices stay in the LicensePricing private override. Paid SKUs without a contract are added to `UnmappedSkus`; no date is invented.

Tests: `Tests\SmartFinOps-Workplace-Valuation.Tests.ps1` (synthetic data, isolated runtime, real SmartM365.Core receipts).

## V2 executive report (lot V2.2)

`SmartFinOps-Workplace-Report.ps1` 2.2.4 turns the valuation outputs into one self-contained HTML page for management. PowerShell 7 only, read-only. It is a pure consumer: it reads only the seven `FinOps_*` CSVs and the valuation receipt, and computes no decision and no amount; it displays, sorts and groups what the engine (2.1.1 or a later 2.1.x) published.

```powershell
.\SmartFinOps-Workplace-Valuation.ps1 -Tenant prod
.\SmartFinOps-Workplace-Report.ps1 -Tenant prod              # Data\<TenantKey>\SmartFinOps_Workplace-Report.html
.\SmartFinOps-Workplace-Report.ps1 -Tenant prod -ValidateOnly # qualify and build, write nothing
```

Before reading anything the report qualifies the run: receipt `Owner=SmartFinOps-Receipt` 1.1 from engine 2.1.4 or later, `Status=Completed`, same TenantKey and profile, exactly the seven files with their SHA256 and row counts (hashed and parsed from the same bytes), exact headers, one `FinOpsRunId`, `SnapshotId` and Licences `ReceiptRunId` on every row, and no failed equivalence control. Otherwise the page is "Report not available" with the reason and no figure; it replaces the current report so that no older figure stays visible, and the script exits with code 1. When the Licences report was not measured the report is still produced, with a status banner and every figure `Not measured`.

Sections follow the SmartFinOps rules: executive summary, financial potential (R with the enabled-inactive accounts shown separately, conditional E3 to F3 outside the total, cost avoidance, price origin), priority decisions, cost/quality/risk trade-offs (Review, Blocked, Not measured, Not valued, Capped, signals), capacity, renewals and forecast (`RealizableFrom`, renewal unknown), realization when a baseline exists, showback (AD domain, country, entity, department), assumptions and a folded technical appendix (report qualification, sources, controls G1 to G6). Statuses are always visible: an empty amount is shown as its status, never as zero; catalogue estimates and contract prices are labelled.

Confidentiality: aggregates only. No user name, UPN or user ID is rendered (detail stays in the CSVs); free text that quotes an e-mail address or a GUID is masked. In the showback, values with fewer than `Report.ShowbackMinimumGroupUsers` users (default 5) are combined into "Other (groups under N users)"; when that row alone would stay under the threshold, the smallest other values join it, so no value can be deduced by subtraction. The page is self-contained (inline CSS and SVG, no script, no external resource, every value HTML-encoded) and deterministic: the same inputs give the same HTML except the generation time. The log is written next to the page; both are copied to `Archive\<FinOpsRunId>\` when that run is archived.

Tests: `Tests\SmartFinOps-Workplace-Report.Tests.ps1` (synthetic FinOps fixtures, input contract checked against the engine output specification).

## V2 daily run (lot V2.4a)

The SmartInventory orchestrator runs both V2 scripts every day for the prod profile, as two jobs of the `FinOps` group in `SmartM365/SmartInventory/Orchestrator/Orchestrator-Jobs.json.template`. Their `ScriptPath` points outside SmartInventory (`..\..\SmartFinOps\...`): the scripts are launched directly, without a SmartInventory wrapper, and the orchestrator appends `-Tenant <profile>`.

| Job | Schedule | Depends on | Rule |
| --- | --- | --- | --- |
| `SmartFinOps-Workplace-Valuation` | Daily 05:45 | `M365-Licences-Inventory`, `M365-LicensePricing-Inventory`, `AD-Inventory`, `M365-ActiveUsers-Inventory` | `LatestOccurrence`, wait up to 1080 min |
| `SmartFinOps-Workplace-Report` | Daily 05:45 | `SmartFinOps-Workplace-Valuation` | `LatestOccurrence`, wait up to 1110 min |

- `LatestOccurrence` requires today's occurrence of every dependency, which is why the jobs start after Licences (04:45) and LicensePricing (05:30). A Licences report from the previous day is never valued as today's. When a dependency fails, the jobs wait for a successful recovery run; otherwise they end `BlockedDependencyTimeout` and FinOps does not run that day. The rule applies to all four dependencies.
- On weekdays Licences ends around 07:50 and both jobs finish before `WorkplaceEvidence-Prepare` (08:30). On Saturday Licences ends around 18:00 after the full Friday EXO inventory; the 1080-minute wait covers it (until 23:45).
- Elected, `SharedRuntime` only (no Graph), PowerShell 7, shared `ConcurrencyKey` `SmartFinOps-Workplace`, timeouts 90 and 30 min (deliberately below the 1380-minute daily policy, reported by its audit), one retry. Valuation has `ContinueOnError=false`: a failed valuation blocks the report.
- Outputs stay in `Data\<profile>` of the checkout the orchestrators run from; all servers must use the same checkout so the archive and realization baselines stay in one place.

Exit codes of both scripts follow their completion banner, as the orchestrator expects:

| Code | Valuation | Report |
| --- | --- | --- |
| `0` | Completed, no warning | Page built from final figures |
| `3` | Completed with warnings: Licences not measured, a required source not measured, provisional amounts, disabled-account licenses not measured | Page built, but Licences not measured or figures provisional |
| `1` | Run error or failed equivalence control (receipt `Failed`) | Report not available, or run error |

With the current Licences report, `M365_OneDrive_Usage.csv` is provisional (no source receipt), so both jobs end with code 3 until that source has a receipt. ESU prices are optional and do not raise a warning. The first log line of each script (`Starting SmartFinOps Workplace ... v<version>`) is the success evidence the orchestrator looks for in the job log.

Launchers:

- direct: `Start-SmartFinOps-Workplace-Valuation-Prod.cmd`, then `Start-SmartFinOps-Workplace-Report-Prod.cmd` (this folder);
- through the orchestrators: `SmartM365/SmartInventory/LaunchersByOrchestrator/Request-SmartFinOps-Workplace-Valuation.cmd` and `Request-SmartFinOps-Workplace-Report.cmd` (add `-IncludeDependencies` to collect the sources first).
