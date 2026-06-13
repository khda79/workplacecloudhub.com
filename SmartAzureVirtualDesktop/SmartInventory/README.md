# SmartAzureVirtualDesktop Inventory

Azure Virtual Desktop inventory scripts for reporting, Power BI datasets, operational checks, and support handoff.

## Organization

- `Estate/`: host pools, workspaces, application groups, applications, desktops, session hosts, scaling plans, private endpoints, and subscription summaries.
- `Health/`: session host health, user sessions, drain mode, capacity, and operational status.
- `Diagnostics/`: Azure Monitor diagnostic settings for Azure Virtual Desktop resources.
- `Scaling/`: autoscale scaling plans, schedules, assignments, and host pools without scaling coverage.
- `FSLogix/`: candidate Azure Files profile storage and storage posture in AVD resource groups.
- `Cost/`: AVD cost review signals such as inactive hosts, unavailable hosts, autoscale gaps, and unattached disks in AVD resource groups.

For every SmartInventory `.ps1`, a `Start-<ScriptName>-Prod.cmd` and `Start-<ScriptName>-Test.cmd` launcher is stored next to the script.
