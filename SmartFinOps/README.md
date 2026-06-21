# SmartFinOps

SmartFinOps is a read-only FinOps reporting project for workplace and cloud operations.

The first scope is **SmartFinOps Workplace**: it consumes existing SmartM365 SmartInventory CSV exports, correlates Microsoft 365, Intune, Entra, and Active Directory signals, then produces CSV outputs and a standalone HTML report.

## Current Scope

The initial Workplace scope uses SmartM365 `DATA-LAST` exports only. It does not connect to Microsoft Graph, Azure, Citrix, or Azure Virtual Desktop directly.

Initial source families:

- Microsoft 365 users and licensing.
- Intune managed devices, compliance, BIOS/system, Autopilot, Windows Update, and upgrade readiness when available.
- Entra devices.
- Active Directory users and computers.
- Exchange mailbox and backup-protection exports when available.

Citrix, Azure Virtual Desktop, and broader Azure cost signals are intentionally out of scope for the first version.

## Main Script

```powershell
.\SmartFinOps-Workplace-Analyze.ps1 -Tenant test
```

The script reads SmartM365 CSV files from:

```text
..\SmartM365\Data\Tenants\<TenantKey>\DATA-LAST
```

It writes SmartFinOps outputs to:

```text
SmartFinOps\Data\Tenants\<TenantKey>\DATA-ALL
SmartFinOps\Data\Tenants\<TenantKey>\DATA-LAST
SmartFinOps\Data\Tenants\<TenantKey>\LOG-ALL
```

## Outputs

- `SmartFinOps_Workplace_Summary.csv`
- `SmartFinOps_Workplace_LicenseOptimization.csv`
- `SmartFinOps_Workplace_DeviceOptimization.csv`
- `SmartFinOps_Workplace_DataQuality.csv`
- `SmartFinOps_Workplace_Report.html`

Historical timestamped copies are stored under `DATA-ALL`; stable latest copies are stored under `DATA-LAST`.

## Configuration

Create local runtime configuration files from the templates when you need overrides:

- `SmartFinOps.global.local.json.template` -> `SmartFinOps.global.local.json`
- `Config/Tenants/tenant.local.json.template` -> `Config/Tenants/<TenantKey>.local.json`

Local `*.local.json` files must not be committed.

The price model is optional. When monthly unit prices are not configured, SmartFinOps reports counts and leaves estimated cost fields empty instead of inventing prices.
