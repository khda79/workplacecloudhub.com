# Intune Inventory

Intune-focused inventory and export scripts, grouped under `SmartInventory/M365Inventory` because they feed the same reporting and dataset layer as the Microsoft 365 inventory exports.

## Organization

- `Applications/`: discovered app inventory from Intune.
- `Autopilot/`: Windows Autopilot device inventory.
- `Devices/`: managed device, BIOS, compliance, system, and upgrade eligibility inventories.
- `EndpointAnalytics/`: read-only standard Endpoint Analytics exports for scores, startup performance, application reliability, and work-from-anywhere reporting. Advanced Analytics is excluded.
- `RBAC/`: Intune RBAC group membership inventories.
- `WindowsUpdate/`: Windows Update and Autopatch reporting from Intune.
- `WindowsUpdate/Archive/`: superseded Intune Windows Update inventory versions kept for reference.

`Export-IntuneRemediations.ps1` remains at the folder root because it is a utility for exporting remediation packages rather than a device inventory report.

## Graph paging and partial-result safeguards

The consolidated Graph candidate updates only transport and publication safety in
Discovered Apps 1.25, Device System 2.4, BIOS 1.12, Compliance 1.16, Upgrade
Eligibility 1.20, RBAC 1.13, Remediations export 1.8, Windows Update status 1.32
and Autopatch alerts 1.16. Manual collection pages must expose `value`, repeated
`@odata.nextLink` values are rejected, and exhausted later-page failures return no
partial collection.

Compliance device summary and policy detail are separate completeness domains. A
complete device list can still refresh `Intune_Devices_Compliance.csv`; if any
policy or setting-state lookup is incomplete, the detailed policy CSV is not
promoted and its preceding `DATA-LAST` version remains in place. Discovered Apps
retains its resume/cache and physical-row gates and now promotes the validated
relation CSV atomically.

Permissions, configuration files, CSV schemas, filenames and business algorithms
are unchanged. Detailed synthetic audit evidence is retained outside the public
repository and does not establish tenant qualification.

Windows 11 Readiness Issues 1.21 uses the shared atomic publisher for detail,
summary, archives and weekly manifest. Its readiness classification and input
contract are unchanged.

## Endpoint Analytics

`EndpointAnalytics/SmartM365-EndpointAnalytics-Inventory.ps1` exports standard Endpoint Analytics reports through Microsoft Graph `deviceManagement/reports/exportJobs`. Microsoft currently requires `DeviceManagementManagedDevices.ReadWrite.All` to create the temporary export-job resource, although the collector performs no device, policy, baseline, assignment, or remediation change.

Battery Health (`BR*`), Resource Performance (`EAResourcePerf*`), Anomalies (`EAAnomaly*`), Device Timeline, and Device Query are intentionally excluded.

See `EndpointAnalytics/README.md` for report names, mappings, CSV grains, API-version rationale, and validation examples.
