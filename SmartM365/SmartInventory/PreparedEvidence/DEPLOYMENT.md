# Prepared evidence: deployment and operation

`SmartM365-WorkplaceEvidence-Prepare.ps1` is a thin SmartM365 wrapper. The evidence generators, the
publication pipeline (`PreparedEvidencePipeline.psm1`) and the table contracts live in the
SmartWorkplaceIntelligence product. Both are deployed as part of the repository; there is no separate
package.

## 1. Deploy

Deploy the reviewed repository so that `SmartM365` and `SmartWorkplaceIntelligence` stay **siblings**,
for example `\\<server>\<share>\SmartM365` and `\\<server>\<share>\SmartWorkplaceIntelligence`.
The wrapper finds the product next to SmartM365 unless `WorkplaceIntelligenceRootPath` is set in its
runtime JSON. Deploy both folders together: a generator fix in SmartWorkplaceIntelligence has no effect
while the old product files remain.

The contracts are `config/prepared-evidence-contract.json.txt` and `config/prepared-source-contract.json.txt`.
The SmartM365 JSON transport reads the `.json.txt` name first, so a stale `.json` copy beside it is never
used; remove such leftovers instead of editing them.

Keep existing `*.local.json`, tenant profiles, account-classification rules
(`SmartM365/SmartInventory/Config/AccountClassification.psd1`), source data and scheduled tasks unchanged.
The two private classification workbooks must exist in the tenant data root (parent of DATA-LAST and DATA-ALL).

## 2. Preflight

Run `SmartM365\SmartInventory\Launchers\Cloud\Test-SmartM365-WorkplaceEvidence-Prepare.cmd` under the
collection account. It selects `-Tenant prod -ValidateOnly -Offline`, can create the missing runtime JSON
from the template and merge missing keys, and checks source presence, fingerprints and transport age. It
generates no CSV, uploads nothing and sends no notification.

Verify that the effective profile points to the intended sibling DATA-LAST and DATA-ALL; DATA-POWERBI is
derived from their parent. Do not weaken age, missing-source, tenant or empty-table checks to pass preflight.

## 3. Scheduled and requested runs

The orchestrator runs the `WorkplaceEvidence-Prepare` job daily at 06:30 with the `FreshSuccess` dependency
rule (48 hours) on its source collectors (see `SmartInventory/Orchestrator/README.md`). For an
out-of-schedule run, ask the orchestrators instead of starting the script by hand:

```
\\<server>\<share>\SmartM365\SmartInventory\LaunchersByOrchestrator\Request-WorkplaceEvidence-Prepare.cmd
```

`Start-SmartM365-WorkplaceEvidence-Prepare-Offline.cmd` remains available for a local rebuild without API
collection, upload or notification. It still writes logs, snapshots inputs, generates all CSVs and publishes
the validated local DATA-POWERBI pointer; do not overlap it with collection or another preparation run.

## 4. Checking a batch

A failed run returns nonzero and never replaces the last good pointer (`DATA-POWERBI/current.json.txt`).
For a new batch, check that `validation.json.txt` reports Passed with every contract file, that source
dates and coverage are plausible and that historical coverage is retained: the publication guard refuses
any loss of a historical key compared with the previous batch. Compare expected workforce, device,
license and mailbox counts before relying on the batch. Never post private logs or manifests publicly.

The synchronized SharePoint folder is not an atomic multi-file transaction; never assume that the pointer
arriving means every batch file has arrived. Validate transfer completion before a Power BI refresh.

## 5. Rollback

Redeploy the previous repository revision of both folders. Every DATA-POWERBI batch is preserved and the
previous good batch stays published until a new batch passes validation; no deletion is required.
