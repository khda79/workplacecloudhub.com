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
Discovered Apps 1.25, Device System 2.5, BIOS 1.12, Compliance 1.16, Upgrade
Eligibility 1.20, RBAC 1.13, Remediations export 1.9, Windows Update status 1.33
and Autopatch alerts 1.16. Manual collection pages must expose `value`, repeated
`@odata.nextLink` values are rejected, and exhausted later-page failures return no
partial collection.

Device System 2.5 deliberately uses managed-device pages of 500 with a 500 ms
inter-page pause and permits eight transient retries, each bounded by Graph's
`Retry-After` value and the existing retry cap. This favors a longer collection
over an incomplete publication. It does not change Graph permissions, CSV schemas,
export names, or inventory interpretation; the existing `DeviceManagement*` read
permissions remain sufficient.

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

## Windows Update country breakdown

The Windows Update status summary email places two tables between Windows
version distribution and Fleet OS coverage. The first covers enabled Windows
rows in `Intune_Devices_Inventory.csv`, once per managed Device ID. A device is
included only when its exact Azure AD Device ID matches an Entra DeviceId with
`AccountEnabled=True` or an AD ObjectGUID with `Enabled=True`. It shows
Windows 11, Windows 10, unknown/other OS, and totals for each primary user's
Entra `CountryOrRegion`. Country is joined through the managed device's UserId
and the active-user Object Id. It is not the device's physical location.
Missing or ambiguous user IDs, users absent from the active-user snapshot,
and blank countries are counted under `Country unknown`. There is no UPN or
device-name inference. OS family uses the Intune OS version build (Windows 11
at build 22000 or later; Windows 10 from build 10240 through 21999).

The same table splits each Intune Windows device into exclusive Feature Update
policy groups: the selected reference policy, another exported Feature Update
policy only, or neither exported policy. Reference-policy membership takes
precedence when a device occurs in both. These groups describe policy report
presence, not confirmed Windows Autopatch enrollment. The reference-policy
report can contain Device IDs absent from the current Intune inventory; those
are disclosed separately rather than added to the Intune total. Disabled
devices and those without positive enabled evidence are excluded and counted
separately. The existing Windows version distribution and Fleet OS coverage
sections retain their reference-policy scope and are not activation filtered.

The second table shows enabled AD Windows 10/11 computer objects with no exact
match from AD ObjectGUID to Intune Azure AD Device ID. The enabled rule is the
same Entra-or-AD rule, using the AD ObjectGUID as the identity. It also shows
the subset active in the last 45 days. No exact match does not prove that a
computer is unenrolled: identity differences, synchronization and snapshot
timing can prevent a join. AD country is not qualified and is therefore not
inferred for this table.

The tables read `Intune_Devices_Inventory.csv`, `M365_Entra_Devices.csv`,
`AD_Computers_AllDomains.csv`, and `M365_Users_Active.csv` from the tenant's
DATA-LAST snapshot. Their last write times must be no older than
`CountrySourceMaxAgeHours` (48 hours by
default). CSV rows use the effective `TenantKey` from the selected tenant
profile, which can differ from the `-Tenant` profile name. Missing, stale,
invalid, or non-reconciling sources display an unavailable notice for the
affected table. The CSV export and fleet severity
remain unchanged. The `OnChange` mail state includes both breakdowns when
available.

The Feature Update rows used by the email are the current Graph report rows in
memory. They are already scoped to the selected tenant and receive their
`TenantKey` column only when the CSV is published. The breakdown accepts these
untagged in-memory rows; if policy rows already have a `TenantKey`, they must
match the effective key.

## Endpoint Analytics

`EndpointAnalytics/SmartM365-EndpointAnalytics-Inventory.ps1` exports standard Endpoint Analytics reports through Microsoft Graph `deviceManagement/reports/exportJobs`. Microsoft currently requires `DeviceManagementManagedDevices.ReadWrite.All` to create the temporary export-job resource, although the collector performs no device, policy, baseline, assignment, or remediation change.

Battery Health (`BR*`), Resource Performance (`EAResourcePerf*`), Anomalies (`EAAnomaly*`), Device Timeline, and Device Query are intentionally excluded.

See `EndpointAnalytics/README.md` for report names, mappings, CSV grains, API-version rationale, and validation examples.
