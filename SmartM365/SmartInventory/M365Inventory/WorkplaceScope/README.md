# Intune WorkplaceScope collections

`SmartM365-WorkplaceScope-Inventory.ps1` remains the source for Entra direct
membership and the five Intune policy families used by CMDB preparation. Three
independent jobs add Intune operational evidence without changing the CMDB
producer dependency set.

| Collector | Source | Canonical DATA-LAST files |
| --- | --- | --- |
| `SmartM365-IntuneConfigurations-Inventory.ps1` | Graph beta policy families, assignments and Settings Catalog setting definitions; `DeviceAssignmentStatusByConfigurationPolicyV3` and `DeviceStatusesByConfigurationProfileV3` export jobs | `Intune_ConfigurationPolicies.csv`, `Intune_ConfigurationSettings.csv`, `Intune_ConfigurationAssignments.csv`, `Intune_ConfigurationAssignmentStates.csv`, `Intune_ConfigurationDeviceStates.csv` |
| `SmartM365-IntuneManagedApps-Inventory.ps1` | Graph v1.0 `mobileApps` and assignments; `OrgAppsInstallStatus` and `OrgDeviceInstallStatus` export jobs | `Intune_ManagedApps.csv`, `Intune_ManagedAppAssignments.csv`, `Intune_ManagedAppInstallSummary.csv`, `Intune_ManagedAppDeviceStates.csv` |
| `SmartM365-IntuneEnrollmentFailures-Inventory.ps1` | `DeviceEnrollmentFailures` export job | `Intune_EnrollmentFailures.csv` |

All collectors use tenant-specific certificate authentication, preflight, Core
CSV publication, source receipts, weekly history and configured SharePoint/Teams
actions. `-ValidateOnly` checks prerequisites without an export job or business
CSV. `-MaxItems` runs sample Graph/export queries without canonical publication.

The configuration collector exports native policy and Settings Catalog setting
JSON, including configured values. Keep these operational CSVs under the same
access controls as other Intune configuration exports. Its device status is
Intune's report state, not direct on-device proof. `ReportedApplied` requires
an explicit successful policy state; other or absent states remain distinct.
An assignment is not proof of application.
The managed-app collector covers applications published in Intune, not apps
detected on devices. Missing install-status rows mean unobserved, not failed or
not installed. `ReportedInstalled` requires an explicit installed state from
Intune. Enrollment failures contain no device ID; the report must not
be joined to an inventoried device by UPN or name. Its snapshots should not be
summed across weekly history as distinct failure events.

The jobs are enabled in `Orchestrator-Jobs.json.template` at 04:30, 05:00 and
05:30. Their duration estimates are initial scheduler values and need review
after tenant-run evidence. Template changes do not publish the live orchestrator
configuration. Tenant runs and publication remain separate operator actions.

For a direct run on the current host, use the three matching
`SmartInventory/Launchers/Cloud/Start-SmartM365-Intune*.cmd` launchers.
They default to tenant `prod`, pass additional script arguments such as
`-ValidateOnly`, and bypass orchestrator claims and scheduling. Use the
`LaunchersByOrchestrator/Request-Intune-*.cmd` launchers to request a run
through the resident orchestrator after the jobs are present in its live manifest.
