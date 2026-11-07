# Optional AD monitoring

AD health checks are optional monitoring for Workplace Intelligence, not a prerequisite for CMDB or other evidence. The monitoring job and its alerts remain enabled independently.

`AD_HealthCheck.csv` is captured when available within 48 hours. The Security and Data Trust readers also check the oldest `RunDateUtc`, never a copied file's modification time. Missing, stale, empty or unusable monitoring produces an explicit unavailable observation without health counts or score. `NotMeasured` checks are excluded from measured coverage, not treated as healthy.

Skipped optional inputs are recorded in capture diagnostics and batch provenance. All other required-source checks and all accepted-source tenant checks remain strict. This does not authorize using an incompatible tenant or bypassing a required inventory.

Remove only `AD-HealthCheck` from `WorkplaceEvidence-Prepare.DependsOn` in the deployed job manifest; do not disable the monitoring job. `CmdbEvidence-Prepare` already has no such dependency.
