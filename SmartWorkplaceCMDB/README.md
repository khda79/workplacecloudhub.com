# SmartWorkplaceCMDB

SmartWorkplaceCMDB is an autonomous Workplace configuration management database project for Microsoft workplace environments.

The project collects workplace inventory data, normalizes it into CMDB entities, prepares Power BI-ready tables, and produces local reports without depending on SmartM365, SmartFinOps, SmartAzure, SmartAzureVirtualDesktop, or SmartCitrix as its foundation.

## Initial Scope

The first scope is Microsoft cloud workplace inventory:

- Entra ID users, groups, and devices.
- Intune managed devices and compliance signals.
- Microsoft 365 license SKUs and assignments.
- Exchange Online mailboxes.
- User-to-device and source-to-entity relationships.
- Data quality, source freshness, and confidence scoring.

Active Directory, Azure Virtual Desktop, Citrix, local endpoint inventory, and external data sources are planned as later extensions.

## Design Principles

- Keep the project autonomous: own configuration, collectors, schemas, outputs, reports, and Power BI assets.
- Keep tenant values local: use `*.local.json` files and never commit tenant IDs, app IDs, secrets, certificates, logs, exports, or production identifiers.
- Keep raw collection data separate from normalized CMDB tables and Power BI-ready tables.
- Prefer read-only collection by default.
- Treat Power BI as a consumer of curated CMDB tables, not as the primary transformation engine.

## Repository Layout

```text
SmartWorkplaceCMDB/
  Config/       Local configuration templates.
  Collectors/   Source-specific collectors.
  Build/        CMDB normalization and Power BI table generation.
  Reports/      Local HTML report generation.
  Modules/      Shared PowerShell module.
  PowerBI/      Power BI model guidance, queries, measures, and templates.
  Schema/       CMDB and Power BI schema definitions.
  Data/         Runtime output, ignored by Git.
```

## Output Layout

Runtime output is tenant-scoped:

```text
SmartWorkplaceCMDB/Data/Tenants/<TenantKey>/
  DATA-ALL/
  DATA-LAST/
  LOG-ALL/
```

Power BI consumes curated tables from:

```text
SmartWorkplaceCMDB/Data/Tenants/<TenantKey>/DATA-LAST/PowerBI/
```

## Power BI Direction

The Power BI deliverable starts with a direction and executive overview, then expands into technician and detailed inventory views.

Initial overview pages should focus on:

- CMDB coverage.
- User and device counts.
- Compliance posture.
- License assignment coverage.
- Mailbox coverage.
- Source freshness.
- Data quality findings.

Detailed inventory pages will add drill-through views for users, devices, relationships, licenses, mailboxes, stale objects, duplicates, and source conflicts.

## Getting Started

Copy the configuration templates to local files, fill local tenant values, then run the build scaffold:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Build\SmartWorkplaceCMDB-Build.ps1 -Tenant Default
```

The first build creates the tenant output structure and initializes the first CMDB and Power BI table set.
