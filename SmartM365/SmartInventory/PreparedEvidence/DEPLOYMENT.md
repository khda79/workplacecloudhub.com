# Prepared evidence: collection-machine qualification

This is an additive deployment candidate, not an installer or a production-qualified release. Use the same SmartM365 installation and `prod` tenant profile as the existing Cloud launchers. No source collector needs modification. The two new CMD launchers use PowerShell 7 x64, preserve the process exit code and do not bypass execution policy.

## 1. Check before copying

Extract the package into a new temporary directory **outside** the installed repository and outside synchronized DATA. It contains only selected preparation scripts, two schema contracts, launchers, documentation and a disabled job snippet. It contains no client mappings, CSVs, PBIP, runtime JSON, certificate, tenant profile, Core module or account-classification rules.

In PowerShell 7, run the verifier from that extracted package:

```powershell
# Replace both paths with the actual extraction and existing installation parent.
$packageRoot = 'D:\Temporary\WorkplaceEvidence-Package'
$repositoryRoot = 'D:\ExistingInstallation' # contains SmartM365
& "$packageRoot\SmartWorkplaceIntelligence\scripts\Test-PreparedEvidenceDeploymentPackage.ps1" -PackageRoot $packageRoot -ExistingRepositoryRoot $repositoryRoot
```

This read-only check compares file hashes, parses packaged PowerShell, checks installed reference dependencies and ImportExcel. A dependency mismatch is a review stop, not permission to overwrite Core or customer classification rules. A matching hash is a compatibility baseline, not tenant-run qualification. The manifest is not a cryptographic signature; use an approved transfer channel. Under AllSigned, have the package approved/signed through the existing signing procedure; do not bypass the policy. The signature will change hashes, so regenerate the package after signing the sources.

Back up existing files that overlap the manifest, then copy **only** the listed SmartM365 and SmartWorkplaceIntelligence files into matching relative paths. Keep both directories as siblings, or explicitly configure WorkplaceIntelligenceRootPath. Do not replace the whole SmartM365 tree. Do not copy package-manifest.json or the job snippet over any active configuration. Keep existing *.local.json, tenant profiles, account-classification rules, source data and scheduled tasks unchanged. Recheck deployed payload hashes against the manifest before execution.

## 2. First manual preflight

Run `SmartM365\SmartInventory\Launchers\Cloud\Test-SmartM365-WorkplaceEvidence-Prepare.cmd` from the existing installation, under the same account as collection. This selects `-Tenant prod -ValidateOnly -Offline`. The script can create its missing local JSON from the template and merge missing keys. It does not generate CSVs, upload or send notifications. It checks source presence, file fingerprints and transport age; see README for its limits.

Verify the effective production profile points to the intended sibling DATA-LAST and DATA-ALL. DATA-POWERBI is derived from their parent, not hard-coded. The two private classification workbooks must already exist in that parent. Preserve all weekly history. Do not copy client mappings into the code package.

Keep EnableSharePointUpload=false. Keep legacy tenantless evidence refused unless its origin is explicitly approved. Do not weaken age, missing-source, tenant or empty-table checks merely to pass preflight. Provision scratch disk outside synchronized DATA for source snapshots, intermediate outputs and retained diagnostics. No full-run duration estimate has yet been qualified.

## 3. Full run after collection

Only after preflight passes, all required collectors have completed successfully and source coverage has been checked, run `Start-SmartM365-WorkplaceEvidence-Prepare-Offline.cmd`. Do not overlap it with collection, another preparation job, synchronization-dependent reading, or a Power BI refresh.

Offline means no API collection, cloud upload or notification. **It does write logs, snapshot inputs, generate all 24 CSVs and publish the validated local DATA-POWERBI pointer.** Existing accepted batches remain recoverable. A failed run must return nonzero and must not replace the last good pointer. Do not manually point Power BI at an incomplete batch.

Return the process exit code, start/end times, terminal summary and private log location. Check batch.json provenance is RebuiltFromSnapshot, validation.json has Passed=true and 24 files, source dates/coverage are plausible and historical coverage is retained. Compare expected workforce/device/license/mailbox counts before accepting the batch. Never post private logs or manifests publicly.

## 4. Automation and cloud remain separate gates

After the full run is qualified, review the separate WorkplaceEvidence-Prepare.job.json snippet, merge only that job into the existing orchestrator configuration, and choose its schedule after the required collectors. It is deliberately disabled. Its dependency list is ordering information, not proof that every collector succeeded.

Do not enable scheduling, direct SharePoint transfer, notifications or Service refresh during this first qualification. Validate transfer completion/read-back before a later Power BI refresh. The existing synchronized folder is not an atomic multi-file transaction; never assume that current.json arriving means every batch file has arrived. A pinned batch may be used after its transfer is verified.

Rollback: stop using the candidate launchers, keep the job disabled and restore only backed-up overlapping code files if necessary. Preserve every DATA-POWERBI batch and the existing raw history. No deletion is required to roll back code.
