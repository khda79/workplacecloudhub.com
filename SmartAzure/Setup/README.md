# SmartAzure Setup

This folder contains setup and bootstrap scripts for SmartAzure.

- `Install-SmartAzure-SmartInventoryPrerequisites.ps1`: installs PowerShell module prerequisites for all SmartAzure SmartInventory scripts.
- `SmartAzure-Create-AppRegistration.ps1`: creates or updates the Entra ID app registration used for shared Graph mail and SharePoint features.
- `SmartAzure-AppRegistration-Permissions.md`: documents the app permissions and the separate Azure RBAC requirements for SmartAzure inventory.
- `SmartAzure-Set-TeamsWebhook.ps1`: stores and tests Teams Workflows / Power Automate webhook URLs in the selected local tenant profile.

Run from the SmartAzure root:

```powershell
.\Setup\Install-SmartAzure-SmartInventoryPrerequisites.ps1 -TrustRepository -AllowClobber
```

Audit without installing:

```powershell
.\Setup\Install-SmartAzure-SmartInventoryPrerequisites.ps1 -WhatIf -SkipImportValidation
```

The installer covers the Az modules used by SmartAzure inventory scripts plus `Microsoft.Graph.Authentication` and `Microsoft.Graph.Applications` for shared Graph mail and app registration setup features.

Create or update the tenant app registration:

```powershell
.\Setup\SmartAzure-Create-AppRegistration.ps1 -Tenant test -TenantId contoso.onmicrosoft.com -UpdateExisting
```

The app registration bootstrap requires `Microsoft.Graph.Authentication` and `Microsoft.Graph.Applications`. When Mail.Send scoping is enabled, it also requires `ExchangeOnlineManagement`; use `-DisableMailSendScopeSetup` to skip the Exchange Online setup objects.
