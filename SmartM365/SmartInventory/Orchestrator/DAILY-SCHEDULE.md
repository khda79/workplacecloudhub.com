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
| M365-Licences-Inventory | Daily 00:45 |
| M365-EntraDevices-Inventory | Daily 00:25 |
| M365-Teams-Inventory | Daily 03:30 |
| M365-SPO-Inventory | Daily 03:45 |
| Intune-Devices-Inventory | Daily 00:35 |
| Intune-Devices-Compliance-Inventory | Daily 01:30 |
| Intune-WindowsAutopilot-Inventory | Daily 04:00 |
| EXO-Mailboxes-Permissions | Daily 05:00 |
| EXO-Mailboxes-Inventory-Fast | Daily 03:15 |
| EXO-AcceptedDomains-Inventory | Daily 00:10 |
| Mailboxes-PermissionsByUser-Report | Daily 08:30 |
| Exchange2016-Local-Mailboxes-Inventory | Daily 00:35 |
| Exchange2016-ProxyAddresses-Check | Daily 03:15 |
| Exchange2016-Infrastructure-Inventory | Daily 00:00 |
| M365-Usage-Inventory | Daily 04:45 |
| Intune-DeviceSystem-Inventory | Daily 01:15 |
| Intune-Devices-BIOS-Inventory | Daily 00:55 |
| Intune-Devices-UpgradeEligibility | Daily 01:45 |
| Intune-RBAC-GroupMembers | Daily 04:15 |
| Intune-Remediations-Export | Daily 04:30 |
| Intune-AutopatchAlerts-Inventory | Daily 03:10 |
| Intune-Windows11-Readiness-Issues | Daily 07:00 |
| Intune-WinUpdate-Status | Daily 02:00 |
| Exchange-HybridIdentity-Issues | Daily 08:00 |
| M365-CopilotUsage-Inventory | Daily 05:00 |
| M365-TeamsPhonePstnUsage-Inventory | Daily 05:15 |
| Intune-EndpointAnalytics-Inventory | Daily 02:45 |
| EXO-QuarantineMessages-Report | Daily 08:15 |
| M365-SecureScore-Inventory | Daily 02:00 |
| M365-AuthenticationMethodsRegistration-Inventory | Daily 02:10 |
| M365-ConditionalAccess-Inventory | Daily 02:20 |
| Intune-EndpointSecurityHealth-Inventory | Daily 02:30 |
| M365-LicensePricing-Inventory | Daily 05:30 |
| WorkplaceEvidence-Prepare | Daily 09:00 |
| M365-WorkplaceScope-Inventory | Daily 22:00 |
| CmdbEvidence-Prepare | Daily 09:30 |
| M365-SyncHealth-Inventory | Daily 00:20, 02:20, 04:20, 06:20, 08:20, 10:20, 12:20, 14:20, 16:20, 18:20, 20:20, 22:20 |
| EXO-MigrationJobs-Inventory | Daily 06:00, 12:00, 18:00 |

The five central identity, device and license producers have one overnight
collection, with no afternoon repeat. Readiness runs once at 07:00,
hybrid-identity analysis at 08:00, Intelligence preparation at 09:00 and CMDB
preparation/publication at 09:30. WorkplaceScope starts at 22:00 the preceding
evening because collection can take several hours.

SyncHealth runs every two hours at minute 20 and skips missed monitoring slots;
migration-job inventory runs at 06:00, 12:00 and 18:00. These monitoring exceptions
do not trigger full identity/device/license recollection.

## Weekly exceptions

| Job | Occurrences |
| --- | --- |
| EXO-Mailboxes-Inventory | Friday 20:00 |
| EXO-Mailboxes-CalendarPermissions | Tuesday, Thursday 20:00 |
| Exchange2016-Mailboxes-CalendarPermissions | Monday, Wednesday, Friday 04:00 |
| Intune-DiscoveredApps-Inventory | Saturday 20:00 |

Daily EXO Fast refreshes unrestricted mailbox details without live statistics.
The CMDB job depends on that daily acquisition rather than the weekly live-stats
job. Fast can reuse the last statistics snapshot, whose own age is unchanged.
Weekly statistics are therefore not suitable for a requirement of fresh live
statistics every 48 hours. Permission-only output cannot substitute for mailbox
acquisition. Full and Fast retain the shared `EXOMailboxes` concurrency key;
a still-running full inventory can delay a daily Fast occurrence.

The four existing disabled jobs remain disabled: BackupProtectedMailboxes,
BackupPolicyScope, BackupProtectedSitesAndDrives and PowerBIFabricActivity.
The public template also keeps WorkplaceScope and CmdbEvidence-Prepare disabled,
with explicit activation required and no private execution server.

## Dependencies and freshness

Readiness and hybrid-identity analysis use `FreshSuccess` with a 48-hour
scheduler floor. A newer occurrence is not required when an adequate completed
success exists. This gate neither makes CSVs contemporaneous nor bypasses the
analysis scripts' own source checks. Mailbox permissions and the permissions
report retain their existing dependency semantics.

CMDB retains all 17 producer dependencies. Its scheduler completion-age gate
allows weekly Apps (240 hours), while the preparation contract independently
enforces acquisition limits: core sources 48 hours; DiscoveredApps warning after
168 hours and maximum 240 hours. This is not a 240-hour allowance for core CSVs.
Preparation does not renew acquisition dates or silently fall back to unrelated
old exports. An active receipt, an unavailable/invalid source or an expired
acquisition can still delay or reject preparation. A failed preparation does not
replace the last published cohort; consumers must display its actual age.

There is no 100-percent cross-source synchronization requirement. Cross-source
identity differences can occur between acquisitions, even with this timetable.
Their qualification policy is separate and is not weakened by a schedule change.

## Election fallback duration estimates

These rounded fallback values reflect local acquisition receipts and successful
orchestrator history. They are not deadlines or production performance promises.
In particular, AD's observed duration does not imply error-free domain coverage.

| Job | Fallback minutes |
| --- | --- |
| M365-Licences-Inventory | 90 |
| AD-Inventory | 120 |
| Exchange2016-Local-Mailboxes-Inventory | 150 |
| M365-WorkplaceScope-Inventory | 270 |
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
```

Editing a template does not reschedule existing shared jobs: normal template
merge preserves operational values. Deploy scripts/templates only after the
approved Git publication. Apply reviewed runtime configuration through
`Publish-SmartM365OrchestratorConfiguration` with fresh snapshot hash guards and
read-back verification, not by overwriting the shared file.

The existing schedule-plan helper preserves runtime Arguments, Enabled,
execution-server settings and estimated durations. Do not use `-SyncEnabled`
to activate the disabled public CMDB entries. The explicit
[CMDB activation helper](../PreparedEvidence/docs/CMDB-PREPARATION.md#automation-and-scheduling) still requires a
qualified worker and does not launch collectors.

Offline tests do not qualify resident-account permissions, real scheduling,
SharePoint transfer/synchronization or Power BI refresh.
