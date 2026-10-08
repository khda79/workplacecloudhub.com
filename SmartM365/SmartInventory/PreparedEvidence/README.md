# Prepared evidence for CMDB and Workplace Intelligence

This directory prepares reporting CSVs from existing SmartInventory exports.
It does not run collectors or control Power BI. It contains two independent pipelines.

| Pipeline | Entry point | Output | History |
| --- | --- | --- | --- |
| SmartWorkplaceCMDB | `SmartM365-CmdbEvidence-Prepare.ps1` | `DATA-POWERBI-CMDB`: 46 flat CSVs and `current.json.txt` | Current snapshot only |
| SmartWorkplaceIntelligence | `SmartM365-WorkplaceEvidence-Prepare.ps1` | `DATA-POWERBI`: versioned batches selected by `current.json.txt` | Observed source history |

Both read canonical SmartInventory `DATA-LAST`. Intelligence also reads `DATA-ALL`,
private classification mappings and historical sources. Neither reads the other's
prepared output.

## Documentation

- [Deployment and operation](docs/DEPLOYMENT.md): prerequisites, configuration,
  logs, launchers, validation and recovery for both pipelines.
- [CMDB preparation](docs/CMDB-PREPARATION.md): qualification, weekly applications,
  automation, local generation and separate SharePoint publication.
- [CMDB source contract](docs/CMDB-SOURCE-CONTRACT.md): identities, grains, joins
  and metric interpretation.

CMDB's executable contract is `cmdb-prepared-contract.json.txt`.
Intelligence's contracts live in `SmartWorkplaceIntelligence/config`.
Use the deployed contracts and runtime orchestrator manifest for actual behavior.

## CMDB quick start

From the repository root, select the tenant explicitly:

```powershell
# Validate sources only; no generation or transfers.
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartM365\SmartInventory\PreparedEvidence\SmartM365-CmdbEvidence-Prepare.ps1 -Tenant test -ValidateOnly

# Prepare locally; configured log upload is independent of batch publication.
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartM365\SmartInventory\PreparedEvidence\SmartM365-CmdbEvidence-Prepare.ps1 -Tenant test

# Prepare, then publish and read back the new cohort.
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartM365\SmartInventory\PreparedEvidence\SmartM365-CmdbEvidence-Prepare.ps1 -Tenant test -Publish
```

Do not overlap preparation/publication against the same root or refresh Power BI
during transfer. Source acquisition age is not reset by preparation.

## Workplace Intelligence quick start

Cloud launchers under `SmartInventory/Launchers/Cloud` select the existing prod convention:

- `Test-SmartM365-WorkplaceEvidence-Prepare.cmd`: validate sources and mappings.
- `Start-SmartM365-WorkplaceEvidence-Prepare.cmd`: normal preparation.
- `Start-SmartM365-WorkplaceEvidence-Prepare-Offline.cmd`: local preparation
  without SharePoint access or notifications.
- `Start-SmartM365-WorkplaceEvidence-Transfer.cmd`: transfer an existing validated
  batch without collection or CSV generation.

Unlike CMDB validation, Intelligence `-ValidateOnly` reads SharePoint mappings
and can upload execution logs. See the deployment guide before running.

Deploy reviewed code without replacing private configuration, sources or data.
Template merging adds missing keys while preserving existing values. Git updates
do not activate jobs or reschedule the runtime manifest. Preparation, read-back,
synchronization and Power BI refresh are separate checks. Tenant CSVs, logs,
manifests, private configuration and data-bearing models never belong in Git.
