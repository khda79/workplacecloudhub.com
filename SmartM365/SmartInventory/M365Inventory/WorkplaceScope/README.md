# Intune WorkplaceScope collections

`SmartM365-WorkplaceScope-Inventory.ps1` remains the source for Entra direct
membership and the five Intune policy families used by CMDB preparation. Three
independent jobs add Intune operational evidence without changing the CMDB
producer dependency set.

| Collector | Source | Canonical DATA-LAST files |
| --- | --- | --- |
| `SmartM365-IntuneConfigurations-Inventory.ps1` | Graph beta policy families, assignments and Settings Catalog setting definitions; `DeviceAssignmentStatusByConfigurationPolicyV3` and `DeviceStatusesByConfigurationProfileV3` export jobs | `Intune_ConfigurationPolicies.csv`, `Intune_ConfigurationSettings.csv`, `Intune_ConfigurationAssignments.csv`, `Intune_ConfigurationAssignmentStates.csv`, `Intune_ConfigurationDeviceStates.csv`, `Intune_ConfigurationDeviceCurrentStates.csv`, `Intune_ConfigurationStatusCoverage.csv` |
| `SmartM365-IntuneManagedApps-Inventory.ps1` | Graph v1.0 `mobileApps` and assignments; Graph beta `DeviceInstallStatusByApp` export job for each assigned app | `Intune_ManagedApps.csv`, `Intune_ManagedAppAssignments.csv`, `Intune_ManagedAppInstallSummary.csv`, `Intune_ManagedAppDeviceStates.csv` |
| `SmartM365-IntuneEnrollmentFailures-Inventory.ps1` | Graph v1.0 enrollment troubleshooting events | `Intune_EnrollmentFailures.csv` |

All collectors use tenant-specific certificate authentication, preflight, Core
CSV publication, source receipts, weekly history and configured SharePoint/Teams
actions. `-ValidateOnly` checks prerequisites without an export job or business
CSV. `-MaxItems` runs sample Graph/export queries without canonical publication.

The configuration collector exports native policy and Settings Catalog setting
JSON, including configured values. Keep these operational CSVs under the same
access controls as other Intune configuration exports. Its device status is
Intune's report state, not direct on-device proof. `ReportedApplied` requires
an explicit successful `ReportStatus`; raw numeric `PolicyStatus` remains available.
`Intune_ConfigurationDeviceStates.csv` preserves the report rows and their dates.
`Intune_ConfigurationDeviceCurrentStates.csv` keeps the latest dated row for
each policy and device; a missing date cannot establish a current state.
`Intune_ConfigurationStatusCoverage.csv` records source row counts and whether
the assignment-status report returned rows. A zero-row assignment-status report
when policy assignments exist does not qualify per-device assignment status.
An assignment is not proof of application.
The managed-app collector covers applications published in Intune, not apps
detected on devices. `DeviceInstallStatusByApp` requires an `ApplicationId`
filter. The catalog includes unassigned apps, but installation reports are
requested only for distinct apps with at least one assignment. The summary marks
unassigned apps `NotAssigned`; `NoReportRows` is reserved for assigned apps whose
report returned no rows. Neither state proves that the app is absent from a device.
A failed or mismatched app report stops publication for the run. Reports are paced
to limit Graph requests and use bounded retries for transient Graph errors,
including `ServiceUnavailable`, with `Retry-After` when provided.
The summary counts observed report rows, not unique devices or users.
`NoReportRows` means unobserved, not failed or not installed. `ReportedInstalled`
requires an explicit installed state from Intune. `ReportedDeploymentState`
uses `AppInstallState` when present; `InstallState` retains the raw report value,
which can be numeric. An absent `AppInstallState` falls back to `InstallState`.
Unmapped values remain `Unknown` and are summarized without device or user
identities in the run log. If every device row is `Unknown`, publication stops.
Enrollment failures come from
`enrollmentTroubleshootingEvent` records; their optional device identifiers are
retained only when Graph provides them. `EventDateTime` is exported as ISO 8601
UTC; a source value without a time zone stops the run before publication.
`ManagedDeviceIdentifier` preserves the Graph value, while
`HasManagedDeviceIdentifier` is false for a missing value or an all-zero GUID.
The source differs from the Intune admin-center Enrollment failures report.
To assess coverage, export that report for the same period, align its displayed
time zone with UTC, and compare event counts by time, user and failure reason.
Record unexplained differences; matching counts do not prove identical scope or
retention. Its weekly snapshots must not be summed as distinct failure events.

The jobs are enabled in `Orchestrator-Jobs.json.template` at 04:30, 05:00 and
05:30. Their duration estimates are initial scheduler values and need review
after tenant-run evidence. Template changes do not publish the live orchestrator
configuration. Tenant runs and publication remain separate operator actions.

For a direct run on the current host, use `SmartInventory.cmd` with the job
name (`Intune-Configurations-Inventory`, `Intune-ManagedApps-Inventory`,
`Intune-EnrollmentFailures-Inventory`). It passes additional script arguments such
as `-ValidateOnly` and bypasses orchestrator claims and scheduling; the tenant comes from `-Tenant`, then `SMARTM365_ORCHESTRATOR_TENANT`, then `DefaultTenant` of `SmartM365.global.local.json`.
Add `-Request` to run the job through the resident orchestrator after it is present
in its live manifest.
