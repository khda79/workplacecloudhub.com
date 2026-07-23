# SmartFinOps

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

The report contains aggregated evidence only and does not expose raw personal data.

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

No recommendation is an automatic license, mailbox or account change. Shared mailboxes and special accounts always follow a separate review path.

Generic M365 inactivity does not prove that a Dynamics 365, Project, Power BI, Power Apps or other add-on license is unused. Those findings stay review-only and are not monetized without product-specific telemetry.

## Disabled user-mailbox conversion decisions

`SmartFinOps_Workplace_ExchangeOptimization.csv` identifies licensed Exchange Online `UserMailbox` objects that are disabled in AD or Entra ID and still have configured delegates.

A **strong candidate** must be below 45 GB and have no observed archive, litigation or retention hold, service-account signal, or conflicting account state. Mailboxes from 45 GB to below 50 GB require a capacity review. Mailboxes at or above 50 GB and cases with a guardrail failure are excluded.

The decision correlates the canonical mailbox inventory with the separate statistics, archive, permissions, M365 mailbox usage, directory account state, and consolidated license decision sources. Configured permissions prove delegation exists; they do not prove who actively accesses the mailbox or that no application dependency exists.

The recommendation is review-only. It never converts a mailbox or removes a license automatically. Any indicative license value is already included in the no-license opportunity and is not added again.
## Price baseline

`Config/SmartFinOps-Workplace-FrancePriceBaseline.json` contains an indicative France price baseline in EUR excluding VAT.

- Microsoft 365 E3 and F3 use a simple average of the public French prices with and without Teams.
- Other mapped products use the public Microsoft France price or a simple public tier average when applicable.
- Tenant-specific contract prices can override the baseline through `PriceModel.MonthlyUnitPriceBySkuPartNumber` in local configuration.

Indicative prices are used to prioritize opportunities. They are not committed savings. A reduction becomes a realized saving only when purchased units can actually be reduced under the contract or at renewal. Reusing an existing license is reported separately as cost avoidance.

## Configuration

Create local runtime configuration files from the templates when overrides are needed:

- `SmartFinOps.global.local.json.template` -> `SmartFinOps.global.local.json`
- `Config/Tenants/tenant.local.json.template` -> `Config/Tenants/<TenantKey>.local.json`

Local `*.local.json` files must not be committed.

## Source contract

`Config/SmartFinOps-Workplace-SourceContracts.json` defines critical source names, semantic roles, required columns and raw-to-enriched lineage.

For Active Directory:

- `AD_Users_AllDomains.csv` and `AD_Computers_AllDomains.csv` are the canonical enriched KPI tables;
- `AD_Users_AllDomains_Brut.csv` and `AD_Computers_AllDomains_Brut.csv` are validation or fallback sources only.

`-ValidateOnly` checks source existence, required columns and freshness without importing full CSV row sets. Freshness uses `Report Refresh Date` when the source exposes it and falls back to the file modification time otherwise. The default freshness threshold is 72 hours.

SharePoint user activity, Teams device usage, Microsoft 365 Copilot usage, and Teams Phone user usage are activity guardrails for base-license decisions when data is available. They are not monetized as separate products.

Readiness and control sources include Teams Phone assignments and detailed calls, Endpoint Analytics, Entra Connect synchronization health, consolidated Exchange permissions, group-license inventory, and optional Power BI/Fabric activity. These sources qualify evidence or prepare future rules without creating a financial opportunity by themselves.

`SmartFinOps_Workplace_DataQuality.csv` uses these source states:

- `Loaded`: the CSV exists, contains data and satisfies its required-column contract;
- `Empty`: the CSV contains a valid header but no data rows;
- `Missing`: an expected source is absent;
- `Missing optional`: an optional source is absent and does not reduce mandatory coverage;
- `Blocked`: collection is known to be unavailable because of an external authorization or API condition;
- `Invalid schema`: the CSV exists but required columns are missing.

Readiness-only sources do not change license recommendations or financial potential. Files whose names contain `MAXITEMS` remain excluded.
