# Intune Inventory

Intune-focused inventory and export scripts, grouped under `SmartInventory/M365Inventory` because they feed the same reporting and dataset layer as the Microsoft 365 inventory exports.

## Organization

- `Applications/`: discovered app inventory from Intune.
- `Autopilot/`: Windows Autopilot device inventory.
- `Devices/`: managed device, BIOS, compliance, system, and upgrade eligibility inventories.
- `RBAC/`: Intune RBAC group membership inventories.
- `WindowsUpdate/`: Windows Update and Autopatch reporting from Intune.
- `WindowsUpdate/Archive/`: superseded Intune Windows Update inventory versions kept for reference.

`Export-IntuneRemediations.ps1` remains at the folder root because it is a utility for exporting remediation packages rather than a device inventory report.
