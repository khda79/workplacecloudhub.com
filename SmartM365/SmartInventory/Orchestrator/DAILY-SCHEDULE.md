# Single daily inventory cycle

This schedule is a reviewed operating baseline, not a guarantee of start or
completion time. All times use the orchestrator server's local time. Existing
dependencies, retries, capability checks, elections and concurrency guards still
apply. No live collection or shared configuration publication is performed by
editing the template.

## Daily occurrences

| M365-VerifiedDomains-Inventory | Daily 00:05 |
| AD-Inventory | Daily 00:10 |
| AD-HealthCheck | Daily 05:00 |
| M365-ActiveUsers-Inventory | Daily 00:15 |
| M365-Licences-Inventory | Daily 00:15, chained after EXO Fast |
| M365-EntraDevices-Inventory | Daily 00:25 |
| M365-Teams-Inventory | Daily 03:30 |
| M365-SPO-Inventory | Daily 03:45 |
| Intune-Devices-Inventory | Daily 00:35 |
| Intune-Devices-Compliance-Inventory | Daily 01:30 |
| Intune-WindowsAutopilot-Inventory | Daily 04:00 |
| EXO-Mailboxes-Permissions | Daily 00:15, chained after EXO Fast |
| EXO-Mailboxes-Inventory-Fast | Daily 00:15, chained after AD and active users |
| EXO-AcceptedDomains-Inventory | Daily 00:10 |
| Mailboxes-PermissionsByUser-Report | Daily 08:30 |
| Exchange2016-Local-Mailboxes-Inventory | Daily 00:35, after a fresh ProxyAddresses check |
| Exchange2016-ProxyAddresses-Check | Daily 00:00 |
| Exchange2016-Infrastructure-Inventory | Daily 03:15, after the on-premises mailbox chain (same concurrency key), with service health, mailbox database paths and CSS-Exchange Health Checker |
| M365-Usage-Inventory | Daily 04:45 |
| Intune-DeviceSystem-Inventory | Daily 01:15 |
| Intune-Devices-BIOS-Inventory | Daily 00:55 |
| Intune-Devices-UpgradeEligibility | Daily 01:45 |
| Intune-RBAC-GroupMembers | Daily 04:15 |
| Intune-Remediations-Export | Daily 04:30 |
| Intune-AutopatchAlerts-Inventory | Daily 03:10 |
| Intune-WinUpdate-Status | Daily 02:00 |
| M365-CopilotUsage-Inventory | Daily 05:00 |
| M365-TeamsPhonePstnUsage-Inventory | Daily 05:15 |
| Intune-EndpointAnalytics-Inventory | Daily 02:45 |
| EXO-QuarantineMessages-Report | Daily 08:15 |
| M365-SecureScore-Inventory | Daily 02:00 |
| M365-AuthenticationMethodsRegistration-Inventory | Daily 02:10 |
| M365-ConditionalAccess-Inventory | Daily 02:20 |
| Intune-EndpointSecurityHealth-Inventory | Daily 02:30 |
| M365-LicensePricing-Inventory | Daily 05:30 |
| SmartFinOps-Workplace-Valuation | Daily 05:45 |
| SmartFinOps-Workplace-Report | Daily 05:45 |
| WorkplaceEvidence-Prepare | Daily 07:45 |
| M365-WorkplaceScope-Inventory | Daily 18:00 |
| Intune-Configurations-Inventory | Daily 04:30 |
| Intune-ManagedApps-Inventory | Daily 05:00 |
| Intune-EnrollmentFailures-Inventory | Daily 05:30 |
| M365-SyncHealth-Inventory | Daily 00:20, 02:20, 04:20, 06:20, 08:20, 10:20, 12:20, 14:20, 16:20, 18:20, 20:20, 22:20 |
| EXO-MigrationJobs-Inventory | Daily 06:00, 12:00, 18:00 |

The five central identity, device and license producers have one overnight
collection, with no afternoon repeat. Readiness runs once at 07:00 and unified
Intelligence preparation at 07:45. The mailbox chain (EXO Fast, permissions,
licences, hybrid-identity analysis) starts as soon as its sources succeed; see
Dependencies and freshness.
There is no separate CMDB preparation job. WorkplaceScope starts at 18:00 the preceding
evening because collection can take several hours. It shares the `AD-WorkplaceScope`
concurrency key with AD-Inventory: the two memory-heavy collections never run at the
same time anywhere in the cluster. WorkplaceScope normally ends long before AD starts at
00:10; its 480-minute timeout bounds how long it can hold AD back (02:00 without a
retry). AD-Inventory has a 240-minute timeout. A server restart during the evening can
interrupt WorkplaceScope; its catch-up run then delays AD by the remaining collection
time. AD and WorkplaceScope are therefore intentional exceptions of the timeout policy:
`Set-SmartM365-OrchestratorTimeoutPolicy.ps1` keeps them unchanged and the audit lists
them separately.

Endpoint security health passes `-ExportTimeoutSeconds 1200`, the value its manual
launcher has always used, instead of the 600-second script default. Each of its three
Intune report exports may therefore wait up to 20 minutes, one hour in total, well
inside its 1380-minute job timeout.

SyncHealth runs every two hours at minute 20 and skips missed monitoring slots;
migration-job inventory runs at 06:00, 12:00 and 18:00. These monitoring exceptions
do not trigger full identity/device/license recollection.

## Weekly exceptions

| Job | Occurrences |
| --- | --- |
| EXO-Mailboxes-Inventory | Saturday 06:00, after the same Saturday's Fast |
| EXO-Mailboxes-CalendarPermissions | Tuesday, Thursday 20:00 |
| Exchange2016-Mailboxes-CalendarPermissions | Monday, Wednesday, Friday 04:00 |
| Intune-DiscoveredApps-Inventory | Saturday 20:00 |
| Exchange2016-FreeBusy-Check | Sunday 05:00, Exchange hybrid free/busy check with CSS-Exchange Free/Busy Checker (same concurrency key as the other Exchange on-premises jobs) |

Daily EXO Fast refreshes unrestricted mailbox details without live statistics.
The unified Workplace job depends on that daily acquisition rather than the weekly live-stats
job. Fast can reuse the last statistics snapshot, whose own age is unchanged.
Weekly statistics are therefore not suitable for a requirement of fresh live
statistics every 48 hours. Permission-only output cannot substitute for mailbox
acquisition. Full and Fast retain the shared `EXOMailboxes` concurrency key. The full run
depends on the same Saturday's Fast (`LatestOccurrence`, 1080-minute wait), so it
never takes the key before that Fast. With sequential statistics it ran 18 to 21 hours,
ending on Sunday between about 00:00 and 03:00, usually before AD releases the Sunday
Fast; since EXO mailbox inventory 1.29 the statistics are collected by parallel workers
(`StatsParallelThrottle`, default 4) and the run is expected to end on Saturday. A longer
full run can still delay the Sunday Fast until it ends.

The four existing disabled jobs remain disabled: BackupProtectedMailboxes,
BackupPolicyScope, BackupProtectedSitesAndDrives and PowerBIFabricActivity.
The public template enables WorkplaceScope and the three Intune WorkplaceScope
collectors, and removes the former CmdbEvidence-Prepare job. Its dedicated request
launcher and activation path are retired. No private server is stored in the template.
Existing shared manifests require the targeted Workplace cutover publication:
remove the legacy job (including an already disabled entry) and merge its dependencies into Workplace without replacing
the effective schedule. Git deployment alone is not that publication.

## Dependencies and freshness

Readiness uses `FreshSuccess` with a 48-hour scheduler floor. A newer occurrence
is not required when an adequate completed success exists. This gate neither
makes CSVs contemporaneous nor bypasses the analysis scripts' own source checks.

The mailbox chain uses `LatestOccurrence`. EXO Fast, permissions and licences are
due at 00:15 and start as soon as their dependencies of the same night succeed
(Fast after AD and active users; permissions and licences after Fast). Fast waits
up to 1080 minutes for its dependencies; permissions and licences wait up to 1380
minutes because a late Fast (AD retry, or a full EXO run overrunning into Sunday)
delays them.
Hybrid-identity analysis is due at 01:00 with a 1380-minute wait: it only runs on the same night's AD, Fast, permissions,
licences, domains and Exchange 2016 outputs, and no longer reuses an older
permissions export when permissions finish late. A job that uses
`LatestOccurrence` must not be scheduled earlier than any of its dependencies,
otherwise the previous day's dependency occurrence would satisfy it. The weekly
full EXO run follows this rule: it is due at 06:00, after the Fast it depends on.

Workplace depends on every enabled job that produces a CSV it reads, including the
device local-system, migration-job and the three Intune configuration, managed-app and
enrollment-failure producers. Three producers are deliberately not dependencies: the weekly
full `EXO-Mailboxes-Inventory` (it would hold preparation until Saturday afternoon), the
optional `AD-HealthCheck` source (a dependency would make it blocking) and
`Mailboxes-PermissionsByUser-Report`, which runs after preparation and replaces one file
atomically. Template synchronization does not extend an existing `DependsOn` array: add
new dependencies to the shared jobs file explicitly.

Workplace retains the former CMDB producer dependencies. Its scheduler completion-age gate
uses `DependencyMaxAgeHours = 48` as a floor: each dependency accepts the larger of
48 hours and its automatic age (longest schedule gap plus 2 hours), so weekly Apps
are accepted up to about 170 hours. The former 240-hour CMDB gate was not carried
over. The preparation contract independently enforces acquisition limits: core
sources 48 hours; DiscoveredApps warning after 168 hours and maximum 240 hours.
Preparation does not renew acquisition dates or silently fall back to unrelated
old exports. An active receipt, an unavailable/invalid source or an expired
acquisition can still delay or reject preparation. A failed preparation does not
replace the last published cohort; consumers must display its actual age.
When WorkplaceScope is assigned as a `Pinned` producer, its freshness gate waits
while any job holds its concurrency key, so preparation also waits for a running
AD-Inventory through the shared `AD-WorkplaceScope` key.

SmartFinOps Valuation and Report use `LatestOccurrence`: Valuation needs today's
Licences, LicensePricing, AD and active-users occurrences, so it waits for Licences
(usually finished before 05:45 on weekdays, about 19:00 on Saturday) instead of valuing the previous day's
report. A dependency that fails and is not recovered within 1080 minutes ends the
occurrence `BlockedDependencyTimeout`; no FinOps run happens that day. A future
`WorkplaceEvidence-Prepare` dependency on FinOps should use its `FreshSuccess` rule
so preparation is never delayed until Saturday evening.

There is no 100-percent cross-source synchronization requirement. Cross-source
identity differences can occur between acquisitions, even with this timetable.
Their qualification policy is separate and is not weakened by a schedule change.

## Launch engine and execution mode

The schedule does not depend on how a job starts. The on-premises jobs
(`Exchange2016-Local-Mailboxes-Inventory`, `Exchange2016-ProxyAddresses-Check`,
`Exchange2016-Infrastructure-Inventory`, `Exchange2016-FreeBusy-Check`,
`Exchange2016-Mailboxes-CalendarPermissions` and the manual SharePoint on-premises
jobs) run through the catalog engine in Windows PowerShell 5.1, from a local
cache by default (`OnPremExecutionMode = LocalCache`). Switching to `Direct` or
`Auto` (see the SmartInventory README, "On-premises execution mode") removes the
cache copy but changes neither times, dependencies,
`ConcurrencyKey` nor timeouts. `AD-Inventory`, `AD-HealthCheck` and
`EXO-QuarantineMessages-Report` run through the engine in PowerShell 7, always
`Direct`. Every job log states the effective mode.

## Election fallback duration estimates

These rounded fallback values reflect local acquisition receipts and successful
orchestrator history. They are not deadlines or production performance promises.
In particular, AD's observed duration does not imply error-free domain coverage.

| Job | Fallback minutes |
| --- | --- |
| M365-Licences-Inventory | 90 |
| AD-Inventory | 120 |
| Exchange2016-Local-Mailboxes-Inventory | 150 |
| Exchange2016-ProxyAddresses-Check | 30 |
| Exchange2016-Infrastructure-Inventory | 90 |
| Exchange2016-FreeBusy-Check | 15 |
| M365-WorkplaceScope-Inventory | 360 |
| Intune-Configurations-Inventory | 30 |
| Intune-ManagedApps-Inventory | 60 |
| Intune-EnrollmentFailures-Inventory | 10 |
| Intune-DiscoveredApps-Inventory | 1500 |
| EXO-Mailboxes-Inventory-Fast | 90 |
| EXO-Mailboxes-Inventory | 1200 |
| EXO-Mailboxes-Permissions | 150 |
| EXO-Mailboxes-CalendarPermissions | 480 |

The election engine prefers its recorded successful-run median over these
fallbacks. Direct mail-only invocations do not create orchestrator job-run records
and were not used to estimate license acquisition. If mail-only execution is
later introduced as an orchestrated job, use a distinct job name so its duration
does not mix with full collection history.

## Offline verification and deployment

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartM365\Tests\Test-SmartM365OrchestratorDailyScheduleOffline.ps1
# Optionally verify the reviewed private jobs file too:
# ... -AdditionalManifestPath '<absolute local reviewed jobs file>'
# Launch catalog (scripts, parameters, execution, aliases) and engine parity:
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartM365\SmartInventory\Orchestrator\New-SmartM365-OrchestratorJobLaunchers.ps1 -Check
powershell -NoProfile -ExecutionPolicy Bypass -File .\SmartM365\Tests\Test-SmartM365InventoryLaunchOffline.ps1
```

Editing a template does not reschedule existing shared jobs: normal template
merge preserves operational values. Deploy scripts/templates only after the
approved Git publication. Apply reviewed runtime configuration through
`Publish-SmartM365OrchestratorConfiguration` with fresh snapshot hash guards and
read-back verification, not by overwriting the shared file.

The existing schedule-plan helper preserves runtime Arguments, Enabled,
execution-server settings and estimated durations. Do not use `-SyncEnabled`
to restore retired CMDB entries. The legacy activation entry point is blocked;
use the targeted Workplace cutover helper for existing shared configurations.
The removal is backed up and hash-guarded, preserves historical execution records,
and creates no collector or preparation request.

Offline tests do not qualify resident-account permissions, real scheduling,
SharePoint transfer/synchronization or Power BI refresh.
