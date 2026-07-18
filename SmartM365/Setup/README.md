# SmartM365 Setup

This folder contains interactive setup and bootstrap scripts for SmartM365.

- `SmartM365-Create-AppRegistration.ps1`: creates or updates the SmartM365 Entra app registration, permissions, certificate metadata, Teams workspace, SharePoint target, and mail sender setup.
- `SmartM365-AppRegistration-Permissions.md`: documents the permission-by-permission rationale and the scripts that use each permission.
- `SmartM365-Set-TeamsWebhook.ps1`: stores and tests Teams Workflows / Power Automate webhook URLs in the selected local tenant profile.
- `Install-SmartM365-SmartInventoryPrerequisites.ps1`: installs PowerShell module prerequisites for all SmartM365 SmartInventory scripts and reports the RSAT Active Directory prerequisite when missing.

Run these scripts from the SmartM365 root with `.\Setup\<script-name>.ps1` so relative examples and local configuration paths remain easy to read.

The shared tenant-context helper lives in `..\Config\SmartM365-TenantContext.ps1`.

Setup logs and transcripts use the same SmartM365 local output convention as inventory scripts: `Data\Tenants\<TenantKey>\LOG-ALL\Setup` by default, with fallback to `Setup\Output\Tenants\<TenantKey>\LOG-ALL\Setup` when the root `Data` folder is not writable.

Install SmartInventory prerequisites:

```powershell
.\Setup\Install-SmartM365-SmartInventoryPrerequisites.ps1 -TrustRepository -AllowClobber
```

Audit without installing:

```powershell
.\Setup\Install-SmartM365-SmartInventoryPrerequisites.ps1 -WhatIf -SkipImportValidation
```

The installer covers the Microsoft Graph rollup module, the Graph submodules used by SmartInventory and setup scripts, `MicrosoftPowerBIMgmt.Profile` for Power BI admin API authentication, `ExchangeOnlineManagement`, and `MSAL.PS`. Organization cmdlets are covered by `Microsoft.Graph.Identity.DirectoryManagement`; there is no separate `Microsoft.Graph.Organization` module to install.
