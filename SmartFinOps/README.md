# SmartFinOps Workplace — Beta 1.5.0-beta.1

**Published beta prerelease for controlled evaluation.** The versioned package is `SmartFinOps-Workplace-1.5.0-beta.1.zip`. The release manifest retains `channel: beta` and `prerelease: true`; any GitHub release must remain a prerelease.

Read [RELEASE-NOTES.md](RELEASE-NOTES.md), [KNOWN-LIMITATIONS.md](KNOWN-LIMITATIONS.md), and [VALIDATION.md](VALIDATION.md). Published on 2026-09-09: [GitHub beta prerelease](https://github.com/khda79/workplacecloudhub.com/releases/tag/smartfinops-workplace-v1.5.0-beta.1) and [six-language site guide](https://workplacecloudhub.com/tools/smartfinops-workplace/). The immutable approved ZIP retains its prepublication audit notes; this README and the GitHub prerelease record the completed publication.

ZIP SHA-256: `2D60B1B0EB81C91C50B222C6A18BD0AE414F78788F87B34E0988C67D329C5C02`.

[Download the beta ZIP](https://github.com/khda79/workplacecloudhub.com/releases/download/smartfinops-workplace-v1.5.0-beta.1/SmartFinOps-Workplace-1.5.0-beta.1.zip) · [SHA-256 file](https://github.com/khda79/workplacecloudhub.com/releases/download/smartfinops-workplace-v1.5.0-beta.1/SmartFinOps-Workplace-1.5.0-beta.1.zip.sha256)

## Evaluate the beta

Use Windows PowerShell 5.1 or PowerShell 7 on Windows. Extract the complete package, preserve `Config/`, and use existing exports from exactly one tenant. No Graph or Exchange module is needed. For AllSigned, trust the supplied public WorkplaceCloudHub code-signing certificate under your organization's policy.

```powershell
# Validate CSV contracts and freshness first; inspect the resulting DataQuality CSV.
.\SmartFinOps-Workplace-Analyze.ps1 -Tenant test -SmartM365LatestCsvFolderPath 'C:\Inventory\test\DATA-LAST' -ValidateOnly
# Full local analysis. CSVs and reports remain confidential evaluation outputs.
.\SmartFinOps-Workplace-Analyze.ps1 -Tenant test -SmartM365LatestCsvFolderPath 'C:\Inventory\test\DATA-LAST'
```

For tests that create only synthetic data in an isolated temporary directory:

```powershell
.\Tests\SmartFinOps-Workplace-Beta.Tests.ps1
.\Tests\SmartFinOps-Workplace-Integration.Tests.ps1
```

SmartFinOps is a read-only decision-support project that applies FinOps principles to workplace technology.

Its objective is not limited to reducing spend. SmartFinOps helps maximize the business value obtained from every euro spent by combining:

- potential savings;
- cost avoidance and capacity reuse;
- forecasting and renewal decisions;
- quality, performance, continuity and risk trade-offs.

The first scope is **SmartFinOps Workplace**. It consumes SmartM365 SmartInventory CSV exports, correlates Microsoft 365, Intune, Entra ID, Active Directory and Exchange signals, then produces an executive HTML report and supporting CSV files.

## Current scope

SmartFinOps Workplace uses SmartM365 `DATA-LAST` exports only. It does not connect directly to Microsoft Graph, Azure, Citrix or Azure Virtual Desktop.

Current source families include:

- Microsoft 365 users, usage, licenses and tenant capacity;
- Intune devices, compliance, Autopilot and Windows upgrade readiness;
- Entra devices;
- Active Directory users and computers;
- Exchange routing, mailbox and backup-protection signals.

Citrix, Azure Virtual Desktop and broad Azure cost analysis remain outside the first version.

## Main script

```powershell
.\SmartFinOps-Workplace-Analyze.ps1 -Tenant prod
```

The analysis is read-only. Files whose name contains `MAXITEMS` are excluded from imports, validation, freshness checks and findings.

## Simplified outputs

Each tenant now has one current folder and one optional archive:

```text
SmartFinOps\Data\<TenantKey>\
  SmartFinOps_Workplace_Report.html
  SmartFinOps_Workplace_Summary.csv
  SmartFinOps_Workplace_ValueOpportunities.csv
  SmartFinOps_Workplace_LicenseOptimization.csv
  SmartFinOps_Workplace_UserLicenseDecision.csv
  SmartFinOps_Workplace_LicenseCapacity.csv
  SmartFinOps_Workplace_DeviceOptimization.csv
  SmartFinOps_Workplace_ExchangeOptimization.csv
  SmartFinOps_Workplace_DataQuality.csv
  Archive\<RunId>\
    ...same run outputs...
    SmartFinOps-Workplace-Analyze.log
```

The current files are directly accessible under `Data\<TenantKey>`. `Archive\<RunId>` keeps the evidence for a specific run without the former `DATA-ALL`, `DATA-LAST`, `LOG-ALL` and `Workplace` nesting.

## Executive report

The HTML report is designed for a management audience. It shows, in this order:

1. executive summary;
2. indicative financial potential;
3. high-priority decisions;
4. cost, quality and performance trade-offs;
5. license capacity and forecasting;
6. assumptions and a collapsed technical appendix.

The HTML report contains aggregated findings and may contain source paths or environment metadata. Supporting CSVs contain identities and operational data. Treat all outputs as confidential.

All SmartFinOps Workplace user-facing report text and decision values are produced in English.

The top license-utilization banner leads with actionable, persona-led decisions:

- **High-confidence no-license candidates**: persona `None`, disabled or blocked account, and no recent M365 or technical evidence.
- **No-license candidates requiring review**: persona `None`, no recent evidence, but the account is still enabled.
- **Conditional E3-to-F3 candidates**: persona F3, no observed technical blocker, and documented Frontline eligibility still required.
- **E3-to-F3 activity and eligibility reviews**: no technical blocker is observed, but business need and Frontline eligibility remain unproven.
- **Dormant license telemetry signals**: a secondary review population split between named, generic/special, and other account types. It is never presented as a removable-license count or a financial opportunity.

The former “E3 licenses without observed E3 capability usage” card is intentionally excluded. A mailbox below 100 GB and no observed desktop activation do not establish that E3 capabilities are unnecessary. This combination remains available only as explicit diagnostic telemetry in the supporting CSV.

Recent M365 service evidence includes aggregate and detailed Exchange, Email, OneDrive, Teams, SharePoint, Microsoft 365 Apps, Copilot, and Teams Phone activity. Its 90-day window ends on the oldest available `Report Refresh Date` across loaded decision sources. Entra/AD sign-ins and recent Intune device presence use the analysis date as their separate technical evidence date.

Windows Update risk and Windows 11 eligibility are used for service-quality, cost-avoidance, and lifecycle forecasting. They are not converted to euros until an internal incident, remediation, or replacement cost model is approved. Endpoint Analytics remains a performance evidence source and is not monetized by the current model.

## E3, F3 or no-license decision matrix

`SmartFinOps_Workplace_UserLicenseDecision.csv` contains one consolidated decision per licensed user. The matrix crosses:

- SmartM365 `M365LicenseTargetPersona` only; `M365LicenseTargetPersonaRelaxed` is intentionally ignored;
- the direct or group-based assignment source of the current E3/F3 license, including the assigning group names;
- M365 and AD account state;
- aggregate and detailed Exchange, OneDrive, Teams and Email activity;
- Microsoft 365 desktop application activations;
- mailbox and OneDrive storage against the F3 2 GB limit;
- recent ownership of an active Intune-managed device.

`DecisionClass` is generated automatically for every licensed user:

- `Recommended`: strong evidence supports the recommendation;
- `Conditional`: a prerequisite such as documented Frontline eligibility remains;
- `Review`: the available evidence is insufficient or conflicting;
- `Keep`: current licensing remains the appropriate output from the available evidence.

SmartFinOps requires no separate manual input file. The detailed CSV contains the automatic class, recommendation, confidence, evidence, guardrails, indicative financial difference, base-license assignment mode, assigning groups, and the change path required to make a recommendation actionable.

`M365LicenseTargetPersona` is the starting signal for the approved base-license review paths:

- current E3 with persona F3: conditional E3-to-F3 opportunity;
- current E3 with persona `None`: possible no-license candidate;
- current F3 with persona `None`: possible no-license candidate;
- current F3 with persona E3: separate capability and under-licensing risk review.

Persona alone never changes a license. E3-to-F3 cases require recent activity, no observed Microsoft 365 desktop activation, Exchange and OneDrive storage within the 2 GB F3 limits, and documented Frontline eligibility. If Frontline eligibility is not proven, the result remains a conditional review opportunity and not an executable downgrade.

High-confidence no-license findings require persona `None`, a disabled or blocked account, and no recent M365 or technical evidence. Enabled dormant accounts remain review-only. Any recent evidence creates a conflict review even when the account is disabled. `LogonCount = 0` is not a no-license rule because the AD attribute can be unknown and is not replicated between domain controllers.

Before any removal or downgrade valuation, required identity, activity and licensing sources must be loaded and fresh. A matching user activity row is required; F3 also requires valid mailbox and OneDrive storage. Missing evidence, conflicting directory states or simultaneous E3/F3 base suites remain unquantified review cases.

No recommendation is an automatic license, mailbox or account change. Shared mailboxes and special accounts always follow a separate review path.

Generic M365 inactivity does not prove that a Dynamics 365, Project, Power BI, Power Apps or other add-on license is unused. Those findings stay review-only and are not monetized without product-specific telemetry.

## Disabled user-mailbox conversion decisions

`SmartFinOps_Workplace_ExchangeOptimization.csv` identifies licensed Exchange Online `UserMailbox` objects that are disabled in AD or Entra ID and still have configured delegates.

A **strong candidate** must be below 45 GB and have no observed archive, litigation or retention hold, service-account signal, or conflicting account state. Unknown archive or hold state excludes conversion. Mailboxes from 45 GB to below 50 GB require a capacity review. Mailboxes at or above 50 GB and cases with a guardrail failure are excluded.

The decision correlates the canonical mailbox inventory with the separate statistics, archive, permissions, M365 mailbox usage, directory account state, and consolidated license decision sources. Configured permissions prove delegation exists; they do not prove who actively accesses the mailbox or that no application dependency exists.

The recommendation is review-only. It never converts a mailbox or removes a license automatically. Any indicative license value is already included in the no-license opportunity and is not added again.
## Price baseline

`Config/SmartFinOps-Workplace-FrancePriceBaseline.json` contains an indicative France price baseline in EUR excluding VAT.

- Microsoft 365 E3 and F3 use a simple average of the public French prices with and without Teams.
- Other mapped products use the public Microsoft France price or a simple public tier average when applicable.
- Tenant-specific contract prices can override the baseline through `PriceModel.MonthlyUnitPriceBySkuPartNumber` in local configuration.

The supplied baseline remains a historical snapshot dated 2026-07-18, not a newly verified quote. Only EUR is supported; non-EUR overrides are rejected. Indicative prices are used to prioritize opportunities. They are not committed savings. A reduction becomes a realized saving only when purchased units can actually be reduced under the contract or at renewal. Reusing an existing license is reported separately as cost avoidance.

## Configuration

Create local runtime configuration files from the templates when overrides are needed:

- `SmartFinOps.global.local.json.template` -> `SmartFinOps.global.local.json`
- `Config/Tenants/tenant.local.json.template` -> `Config/Tenants/<TenantKey>.local.json`

Local `*.local.json` files must not be committed.

## Source contract

`Config/SmartFinOps-Workplace-SourceContracts.json` defines critical source names, semantic roles, required columns and raw-to-enriched lineage.

For Active Directory:

- `AD_Users_AllDomains.csv` and `AD_Computers_AllDomains.csv` are the canonical enriched KPI tables;
- `AD_Users_AllDomains_Brut.csv` and `AD_Computers_AllDomains_Brut.csv` are validation sources only; this beta does not automatically fall back to raw rows for decisions.

`-ValidateOnly` checks source existence, required columns and freshness without retaining full CSV row sets; it streams report-refresh timestamps across all rows. Freshness uses the oldest valid `Report Refresh Date`. Missing, malformed or future timestamps in that column produce `Unknown`; only a source without the column falls back to file modification time. The default threshold is 72 hours. Full analysis additionally detects duplicate canonical identities and invalid numeric capacity. Successful execution does not mean the source checks passed.

SharePoint user activity, Teams device usage, Microsoft 365 Copilot usage, and Teams Phone user usage are activity guardrails for base-license decisions when data is available. They are not monetized as separate products.

Readiness and control sources include Teams Phone assignments and detailed calls, Endpoint Analytics, Entra Connect synchronization health, consolidated Exchange permissions, group-license inventory, and optional Power BI/Fabric activity. These sources qualify evidence or prepare future rules without creating a financial opportunity by themselves.

`SmartFinOps_Workplace_DataQuality.csv` uses these source states:

- `Loaded`: the CSV exists, contains data and satisfies its required-column contract;
- `Empty`: the CSV contains a valid header but no data rows;
- `Missing`: an expected source is absent;
- `Missing optional`: an optional source is absent and does not reduce mandatory coverage;
- `Blocked`: a custom source contract explicitly declares an external collection block; the generic beta does not assume a tenant-specific authorization failure;
- `Invalid schema`: required columns are missing or a canonical identity is blank/duplicated; excluded from analysis.
- `Error`: the CSV could not be parsed.

Freshness is separate: `Fresh`, `Stale`, `Unknown`, or `NotChecked`. Incomplete assessments explicitly state that zero financial totals do not prove zero opportunity.

Readiness-only sources do not change license recommendations or financial potential. Files whose names contain `MAXITEMS` remain excluded.

## V2 valuation engine (lot V2.1)

`SmartFinOps-Workplace-Valuation.ps1` 2.1.5 is the V2 engine. It runs beside the beta analyzer, which stays unchanged until it is retired. PowerShell 7 only, read-only, no API connection.

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
