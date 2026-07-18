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

## Endpoint Analytics

`EndpointAnalytics/SmartM365-EndpointAnalytics-Inventory.ps1` exports standard Endpoint Analytics reports through Microsoft Graph `deviceManagement/reports/exportJobs`. Microsoft currently requires `DeviceManagementManagedDevices.ReadWrite.All` to create the temporary export-job resource, although the collector performs no device, policy, baseline, assignment, or remediation change.

Battery Health (`BR*`), Resource Performance (`EAResourcePerf*`), Anomalies (`EAAnomaly*`), Device Timeline, and Device Query are intentionally excluded.

See `EndpointAnalytics/README.md` for report names, mappings, CSV grains, API-version rationale, and validation examples.
