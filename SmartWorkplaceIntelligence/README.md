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

## Further reading

- [Prepared evidence reconstruction](docs/prepared-evidence-reconstruction.md)
- [Native domain migration and generic layer retirement](docs/CMDB-CONSOLIDATION.md)
- [Teams and SharePoint evidence](docs/TEAMS-SHAREPOINT-INTEGRATION.md)
- [Device system evidence](docs/DEVICE-SYSTEM-EVIDENCE.md)
- [Private classification templates](config/README.md)

Offline validation, a successful server preparation, publication, and Power BI
rendering are distinct checks. Source publication alone does not qualify a
deployment or every report interaction.
