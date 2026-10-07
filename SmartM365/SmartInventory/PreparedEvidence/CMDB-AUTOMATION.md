# Current-only CMDB automation

`SmartM365-CmdbEvidence-Orchestrator.ps1` 1.0.0 reviews or explicitly applies
two jobs through the existing shared configuration API. It does not submit a
pipeline, launch collectors, change maintenance, or edit Intelligence output.
Default execution only validates the proposed configuration and logs its plan.
Use `-Apply` after deployment prerequisites and actual Power BI qualification.

## Proposed defaults

| Job | Default schedule | Timeout | Action |
| --- | --- | --- | --- |
| M365-WorkplaceScope-Inventory | Daily 00:05 | 720 minutes | Native groups/membership/configuration collection |
| CmdbEvidence-Prepare | Daily 07:30 | 180 minutes | Prepare 46 CSVs, then verified SharePoint publication |

These are reviewable operating defaults, not measured duration guarantees.
`-ScopeTime` and `-PreparationTime` accept local `HH:mm` times. Both use
`RunOnce` for a missed occurrence. The preparation job has two retries, 15
minutes apart. The scheduler's existing locks and concurrency keys still apply.
An existing enabled WorkplaceScope job is preserved without rescheduling it.

Both new jobs are pinned to the single explicitly supplied `-AllowedServers`
worker. Elected mode ignores that property, so it is not used as an apparent
eligibility filter. Verify PowerShell 7, Python 3.10+, required SDK modules,
certificate access, raw/private configuration paths, Graph permissions and
SharePoint `Sites.Selected` write access **under the resident account** first.
Hidden groups additionally need the collector's hidden-membership permission.
This helper validates manifest structure, cluster membership and consistency;
it does not prove those live prerequisites or install them.

## Dependency and freshness boundaries

Preparation keeps all 17 native producer dependencies. Its `FreshSuccess`
scheduler gate is 240 hours to accommodate weekly DiscoveredApps. Independently,
generation checks current source receipts and acquisition ages: core sources
48 hours; Apps maximum 240 hours, warning after 168 hours. Weekly application
relations retain their approved nonblocking coverage warnings. No acquisition
timestamp is reset just by preparation or upload.

The existing DiscoveredApps weekly schedule and arguments are not changed by
this helper. No application collector is launched by preparation. When Apps
passes its maximum age, a completed native collection is needed before a new
cohort can be published.

Manual source receipts do not create orchestrator success records. In particular,
the first scheduled preparation can wait for the newly registered WorkplaceScope
job, or another dependency with no fresh scheduler success. Inspect the existing
GUI's pending-job reasons; do not fabricate success records or automatically
recollect all sources to clear the gate. Maintenance continues to block scheduled
launches, including these jobs; manual preparation remains separately available.

## Review and activation

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartM365\SmartInventory\PreparedEvidence\SmartM365-CmdbEvidence-Orchestrator.ps1 -Tenant test -SharedDataFolderPath '<absolute authoritative Orchestrator path>' -AllowedServers '<verified PS7 worker>'
# After qualification, repeat the exact reviewed command with -Apply.
```

Only the two integration entries can be added/replaced. Conflicting existing
entries are rejected; existing native WorkplaceScope configuration is preserved.
Other jobs and cluster configuration are retained. Publication uses the snapshot
hashes to refuse concurrent configuration edits and verifies the resulting
configuration by read-back. Repeating an identical application skips publication.
This is an explicit deployment helper, not an automatically enabled repository
template. Private server names and paths belong in private runtime arguments.

## Preparation and publication

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartM365\SmartInventory\PreparedEvidence\SmartM365-CmdbEvidence-Prepare.ps1 -Tenant test -Publish
```

`-ValidateOnly` always remains nonpublishing. Without `-Publish`, generation stays
local. With it, the separate publisher verifies and reads back all 46 CSVs, then
publishes `current.json.txt` last. A transport failure fails the job, retaining
the local prepared output for investigation/retry. Flat-folder cloud replacement
is not atomic: consumers must validate every manifest-bound hash and must not
refresh during a transfer. See [publication boundaries](CMDB-PREPARATION.md).

The only prepared target is `DATA-POWERBI-CMDB`. Intelligence's `DATA-POWERBI`,
raw `DATA-LAST`, and the old CMDB collection tree are not rewritten by preparation.
There is no additional CMDB history and no automatic Power BI Desktop refresh.
The report must load all 46 tables from one verified current cohort.

Offline verification:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartM365\Tests\Test-SmartM365CmdbAutomationOffline.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartM365\Tests\Test-SmartM365CmdbIntegrationCandidate.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartM365\Tests\Test-SmartM365CmdbPreparationOffline.ps1
```

Offline success does not prove resident-account access, actual scheduling,
SharePoint throughput/synchronization, or Power BI rendering.
