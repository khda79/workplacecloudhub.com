# SmartAzureVirtualDesktop

PowerShell automation scripts for Azure Virtual Desktop inventory, health, diagnostics, scaling, FSLogix storage posture, and cost optimization reviews.

SmartAzureVirtualDesktop is separate from SmartAzure and SmartM365. SmartAzure covers the broad Azure estate; SmartAzureVirtualDesktop focuses on Azure Virtual Desktop host pools, session hosts, workspaces, application groups, autoscale, diagnostics, profile storage, and operational cost signals.

## Content

- `SmartInventory/`: Azure Virtual Desktop inventory exports for reporting, Power BI datasets, operational reviews, and support handoff.
- `Config/`: shared tenant context and local tenant profile templates.
- `Modules/`: shared helper functions used by SmartAzureVirtualDesktop scripts.

## First Scripts

- `SmartInventory/Estate/SmartAzureVirtualDesktop-AVDEstate-Inventory.ps1`: exports host pools, workspaces, application groups, applications, desktops, session hosts, scaling plans, private endpoint connections, and subscription summaries.
- `SmartInventory/Health/SmartAzureVirtualDesktop-SessionHostHealth-Inventory.ps1`: exports session host health, user sessions, and host pool capacity summaries.
- `SmartInventory/Diagnostics/SmartAzureVirtualDesktop-Diagnostics-Inventory.ps1`: exports Azure Monitor diagnostic settings for Azure Virtual Desktop resources and highlights resources without diagnostics.
- `SmartInventory/Scaling/SmartAzureVirtualDesktop-ScalingPlan-Inventory.ps1`: exports scaling plans, schedules, host pool assignments, and host pools without autoscale coverage.
- `SmartInventory/FSLogix/SmartAzureVirtualDesktop-FSLogixStorage-Inventory.ps1`: exports candidate FSLogix storage accounts and Azure Files shares in resource groups that contain Azure Virtual Desktop resources.
- `SmartInventory/Cost/SmartAzureVirtualDesktop-CostOptimization-Inventory.ps1`: exports common cost signals such as inactive session hosts, unavailable hosts, hosts without sessions, pools without autoscale, and unattached disks in AVD resource groups.

## Run

Run from the `SmartAzureVirtualDesktop` folder:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartInventory\Estate\SmartAzureVirtualDesktop-AVDEstate-Inventory.ps1 -Tenant test -Connect
```

Use device code authentication when browser sign-in is awkward:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartInventory\Health\SmartAzureVirtualDesktop-SessionHostHealth-Inventory.ps1 -Tenant prod -Connect -UseDeviceCode
```

Limit to one or more subscriptions:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartInventory\Estate\SmartAzureVirtualDesktop-AVDEstate-Inventory.ps1 -SubscriptionId "00000000-0000-0000-0000-000000000000"
```

By default, output is written under:

```text
SmartAzureVirtualDesktop\Data\Tenants\<TenantKey>\DATA-ALL\AzureVirtualDesktop\<Area>\<RunId>
SmartAzureVirtualDesktop\Data\Tenants\<TenantKey>\DATA-LAST
SmartAzureVirtualDesktop\Data\Tenants\<TenantKey>\LOG-ALL
```

If the root `Data` folder cannot be created or written, scripts fall back to `Output\Tenants\<TenantKey>\...` next to the running script.

Each script can have its own ignored local configuration file named `<ScriptName>.local.json` in the same folder as the script.

## Launchers

Every SmartInventory script has two adjacent command launchers:

- `Start-<ScriptName>-Prod.cmd`
- `Start-<ScriptName>-Test.cmd`

Launchers call PowerShell 7 when available and pass the matching `-Tenant prod` or `-Tenant test` value.
