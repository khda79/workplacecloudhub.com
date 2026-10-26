# Unified Workplace evidence

This directory prepares reporting CSVs from existing SmartInventory exports.
It does not run collectors or control Power BI. Workplace Intelligence now includes
native CMDB evidence in the same governed DATA-POWERBI batch. The former independent
CMDB pipeline is retained only for rollback until server and Desktop acceptance.
Do not schedule or request both preparations.

| Pipeline | Entry point | Output | History |
| --- | --- | --- | --- |
| Legacy CMDB (job removed) | Historical `SmartM365-CmdbEvidence-Prepare.ps1` | `DATA-POWERBI-CMDB`: retained previous cohort | No scheduled execution |
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

## Legacy CMDB — no new activation

The former activation helper refuses execution and its dedicated request launcher
is removed. The template removes the old job, but existing shared configuration
is not overwritten by a Git update. Use the targeted cutover procedure in
[deployment](docs/DEPLOYMENT.md#targeted-cmdb-cutover) before the next preparation.

Do not overlap preparation/publication against the same root or refresh Power BI
during transfer. Source acquisition age is not reset by preparation.

## Workplace Intelligence quick start

Run these catalog entries with `SmartInventory\SmartInventory.cmd <Name>`
(the tenant comes from `-Tenant`, then `SMARTM365_ORCHESTRATOR_TENANT`, then `DefaultTenant` of `SmartM365.global.local.json`):

- `WorkplaceEvidence-Validate`: validate sources and mappings.
- `WorkplaceEvidence-Prepare`: normal preparation (the orchestrator job).
- `WorkplaceEvidence-Prepare-Offline`: local preparation
  without SharePoint access or notifications.
- `WorkplaceEvidence-Transfer`: transfer an existing validated
  batch without collection or CSV generation.

Unlike CMDB validation, Intelligence `-ValidateOnly` reads SharePoint mappings
and can upload execution logs. See the deployment guide before running.

Successful online preparation and `-TransferOnly` runs send an HTML completion
email through the existing SmartM365 sender, recipients and Graph/SMTP settings.
No additional mail configuration is required. The summary includes the tenant,
batch ID, CSV count, duration, output root, log and transfer status, without
attachments. When batch upload is enabled, the email is sent only after the batch
and current pointer have been read back and verified. When upload is disabled,
the summary explicitly reports a local batch only. Offline, validation, diagnostic,
history repair and metadata conversion modes do not send completion emails.
Mail delivery errors follow the existing error path under the
`Success email notification` phase; an already verified batch remains published.
Check that phase before rerunning preparation solely because a notification failed.

The final CSV count includes validated prepared tables published by this run,
excluding captured source/history artifacts. It remains available if a later
transfer fails. Validation and transfer-only modes generate no CSVs.
Warnings from child generators are relayed into the shared execution counter;
blocked retention adds a warning with its plan reasons. Such runs complete with
warnings while preserving an already verified batch. Embedded retention leaves
the completion banner to the preparation wrapper.

Online Intelligence runs download both classification workbooks under the tenant
DATA root (beside DATA-LAST and DATA-ALL), not under the OS temporary directory.
Both workbooks must pass structural validation before replacing the retained
`SmartWorkplaceIntelligence-PersonaClassification.xlsx` and
`SmartWorkplaceIntelligence-SiteClassification.xlsx` at that root. Existing copies
are overwritten on each successful mapping acquisition, not archived or removed
at completion. A failed download/validation stops the run; retained copies are
never used as an automatic online fallback. Offline mode remains local-only.

Replacement shares the raw-source capture lock and rolls back earlier changes if
the second replacement fails. If rollback itself fails, the run stops and retains
its recovery files with an explicit diagnostic. The preparation consumes its own
validated pair in DATA/mappings/<run-id>; only that intermediate folder is cleaned
at completion. Persistent root workbooks are never part of run cleanup.

The protected mailbox, SharePoint site, and OneDrive account exports come from
interactive Backup collectors. They are optional inputs with a 168-hour freshness
limit: missing or older files are excluded from the prepared snapshot and reported
in optional source diagnostics. Mailbox protection is then marked unknown instead
of treating unverified mailboxes as unprotected. The noninteractive BackupPolicyScope
exports remain required and must pass the normal publication-age check.

Deploy reviewed code without replacing private configuration, sources or data.
Template merging adds missing keys while preserving existing values. Git updates
do not activate jobs or reschedule the runtime manifest. Preparation, read-back,
synchronization and Power BI refresh are separate checks. Tenant CSVs, logs,
manifests, private configuration and data-bearing models never belong in Git.
