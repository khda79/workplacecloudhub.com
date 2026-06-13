# SmartCitrix

PowerShell automation scripts for Citrix inventory and operational reporting.

SmartCitrix is separated from SmartM365, SmartAzure, and SmartAzureVirtualDesktop. It focuses on Citrix Virtual Apps and Desktops, StoreFront, Citrix Licensing, and future Citrix Cloud / Citrix DaaS inventory.

## Content

- `Citrix-OnPrem/`: read-only inventory scripts for on-premises Citrix Virtual Apps and Desktops, StoreFront, and Citrix Licensing environments.
- `CitrixCloud/`: reserved for Citrix Cloud and Citrix DaaS API-based inventory scripts.
- `Config/`: shared tenant context and local tenant profile templates.
- `Modules/`: shared helper functions used by SmartCitrix scripts.

## On-Premises Inventory

The first on-premises scripts use the local Citrix PowerShell SDK and an optional `-AdminAddress` pointing at a Delivery Controller.

- `Citrix-OnPrem/SmartInventory/CVAD/SmartCitrix-OnPrem-CVADSite-Inventory.ps1`: exports site, controller, zone, admin, scope, role, and service status data.
- `Citrix-OnPrem/SmartInventory/CVAD/SmartCitrix-OnPrem-CVADDelivery-Inventory.ps1`: exports catalogs, delivery groups, applications, application groups, access policies, entitlement policies, assignment policies, tags, and reboot schedules.
- `Citrix-OnPrem/SmartInventory/CVAD/SmartCitrix-OnPrem-CVADMachineSessionHealth-Inventory.ps1`: exports machines, sessions, desktops, unregistered machines, disconnected sessions, and health summaries.
- `Citrix-OnPrem/SmartInventory/CVAD/SmartCitrix-OnPrem-CVADPolicy-Inventory.ps1`: exports Citrix policy sets, policies, settings, filters, and test status when supported by the local SDK.
- `Citrix-OnPrem/SmartInventory/CVAD/SmartCitrix-OnPrem-CVADHostingPower-Inventory.ps1`: exports hosting connections, connection status, power time schemes, power actions, delayed actions, catalog reboot schedules, and reboot cycles.
- `Citrix-OnPrem/SmartInventory/StoreFront/SmartCitrix-OnPrem-StoreFront-Inventory.ps1`: exports StoreFront deployment, store, authentication, receiver, farm, gateway, and beacon data where supported by installed StoreFront modules.
- `Citrix-OnPrem/SmartInventory/Licensing/SmartCitrix-OnPrem-Licensing-Inventory.ps1`: exports licensing configuration from CVAD and optional Citrix licensing SDK commands.

## Run

Run from the `SmartCitrix` folder:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Citrix-OnPrem\SmartInventory\CVAD\SmartCitrix-OnPrem-CVADMachineSessionHealth-Inventory.ps1 -Tenant test -AdminAddress "ctx-ddc-01.contoso.com"
```

Use launchers for common local tenant profiles:

```cmd
.\Citrix-OnPrem\SmartInventory\CVAD\Start-SmartCitrix-OnPrem-CVADMachineSessionHealth-Inventory-Test.cmd
.\Citrix-OnPrem\SmartInventory\CVAD\Start-SmartCitrix-OnPrem-CVADMachineSessionHealth-Inventory-Prod.cmd
```

By default, output is written under:

```text
SmartCitrix\Data\Tenants\<TenantKey>\DATA-ALL\Citrix\<Area>\<RunId>
SmartCitrix\Data\Tenants\<TenantKey>\DATA-LAST
SmartCitrix\Data\Tenants\<TenantKey>\LOG-ALL
```

If the root `Data` folder cannot be created or written, scripts fall back to `Output\Tenants\<TenantKey>\...` next to the running script.

Each script can have its own ignored local configuration file named `<ScriptName>.local.json` in the same folder as the script.

## Citrix Cloud

`CitrixCloud/` is intentionally scaffolded separately. Citrix Cloud / Citrix DaaS scripts should use the Citrix Cloud API model and should not reuse on-premises Delivery Controller assumptions.
