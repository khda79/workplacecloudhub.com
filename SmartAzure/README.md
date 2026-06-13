# SmartAzure

PowerShell automation scripts for Azure infrastructure inventory, governance, security posture, and cost optimization.

SmartAzure is separate from SmartM365. SmartM365 covers Microsoft 365, Entra, Intune, Exchange, Active Directory, and endpoint tooling; SmartAzure focuses on Azure Resource Manager estate data such as subscriptions, resource groups, resources, tags, locks, providers, policy, RBAC, networking, backup, Defender for Cloud, and cost signals.

## Content

- `SmartInventory/`: Azure inventory exports for reporting, Power BI datasets, governance reviews, and operational checks.
- `Setup/`: prerequisite installer, Entra ID app registration bootstrap, app permission documentation, and Teams webhook configuration helper.
- `Config/`: shared tenant context and local tenant profile templates.

## Setup

Install the SmartInventory PowerShell prerequisites from the SmartAzure root:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Setup\Install-SmartAzure-SmartInventoryPrerequisites.ps1 -TrustRepository -AllowClobber
```

Create or update the tenant app registration used for shared Graph mail and SharePoint upload features:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Setup\SmartAzure-Create-AppRegistration.ps1 -Tenant test -TenantId contoso.onmicrosoft.com -UpdateExisting
```

The SmartAzure app registration intentionally stays narrower than SmartM365. It grants Microsoft Graph application permissions only for shared capabilities such as `Mail.Send` and `Sites.Selected`. Azure inventory access must be granted separately with Azure RBAC on the target management groups, subscriptions, or resource groups. See `Setup/SmartAzure-AppRegistration-Permissions.md` for the detailed permission model.

Configure Teams Workflows / Power Automate webhook URLs after creating them in Teams:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Setup\SmartAzure-Set-TeamsWebhook.ps1 -Tenant test -Channel Alerts -WebhookUrl "<Alerts workflow URL>"
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Setup\SmartAzure-Set-TeamsWebhook.ps1 -Tenant test -Channel Infos -WebhookUrl "<Infos workflow URL>"
```

## First Scripts

### Azure Estate Inventory

`SmartInventory/Governance/SmartAzure-AzureEstate-Inventory.ps1` exports the broad Azure estate baseline:

- management groups when accessible;
- subscriptions visible to the signed-in identity;
- Azure regions available per subscription;
- resource groups;
- resources with type, location, tags, SKU, plan, identity, and managed-by metadata;
- resource locks;
- resource providers.

Run from the `SmartAzure` folder:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartInventory\Governance\SmartAzure-AzureEstate-Inventory.ps1 -Connect
```

Use device code authentication when browser sign-in is awkward:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartInventory\Governance\SmartAzure-AzureEstate-Inventory.ps1 -Connect -UseDeviceCode
```

Limit to one or more subscriptions:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartInventory\Governance\SmartAzure-AzureEstate-Inventory.ps1 -SubscriptionId "00000000-0000-0000-0000-000000000000"
```

Use a local tenant profile key to isolate output:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartInventory\Governance\SmartAzure-AzureEstate-Inventory.ps1 -Tenant prod -Connect
```

When `Config\Tenants\<TenantKey>.local.json` contains `AzureTenantId`, `TenantId`, or `OrgDomain`, scripts use that value for Azure sign-in if `-TenantId` is not passed explicitly. Shared Graph and notification settings keep the same names as SmartM365, including `AppId`, `Thumb`, `Thumbprint`, mail, Teams, and SharePoint keys.

By default, output is written under the SmartAzure root `Data` folder:

```text
SmartAzure\Data\Tenants\<TenantKey>\DATA-ALL\Azure\Estate\<RunId>
```

Latest copies are also written under:

```text
SmartAzure\Data\Tenants\<TenantKey>\DATA-LAST
```

`-Tenant` controls the local output isolation key and defaults to `test`. It is separate from `-TenantId`, which is passed to Azure sign-in. If the root `Data` folder cannot be created or written, the script falls back to `Output\Tenants\<TenantKey>\...` next to the script. Use `-OutputRoot` and `-LatestOutputRoot` to override these paths for a specific run.

Each script can also have its own ignored local configuration file named `<ScriptName>.local.json` in the same folder. Values in that file override the central tenant/global configuration for that script.

### RBAC Inventory

`SmartInventory/RBAC/SmartAzure-RBAC-Inventory.ps1` exports Azure role assignments, privileged assignments, custom roles, optional classic administrators, and a subscription summary.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartInventory\RBAC\SmartAzure-RBAC-Inventory.ps1 -Connect
```

### Cost Optimization Inventory

`SmartInventory/Cost/SmartAzure-CostOptimization-Inventory.ps1` exports common review candidates: unattached managed disks, old snapshots, unused public IPs, stopped/deallocated VMs, and a summary.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartInventory\Cost\SmartAzure-CostOptimization-Inventory.ps1 -Connect
```

### Network Exposure Inventory

`SmartInventory/Network/SmartAzure-NetworkExposure-Inventory.ps1` exports public IPs, inbound NSG allow rules from Internet-like sources, load balancers, application gateways, private endpoints, and exposure summaries.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartInventory\Network\SmartAzure-NetworkExposure-Inventory.ps1 -Connect
```

### Policy Compliance Inventory

`SmartInventory/Governance/SmartAzure-PolicyCompliance-Inventory.ps1` exports Azure Policy assignments, definitions, initiatives, exemptions, policy state details, and compliance summaries.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartInventory\Governance\SmartAzure-PolicyCompliance-Inventory.ps1 -Tenant prod -Connect
```

### Defender for Cloud Inventory

`SmartInventory/Security/SmartAzure-DefenderForCloud-Inventory.ps1` exports Defender plans, secure score, secure score controls, security recommendations, security contacts, auto-provisioning settings, regulatory compliance data, and subscription summaries.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartInventory\Security\SmartAzure-DefenderForCloud-Inventory.ps1 -Tenant prod -Connect
```

### Backup Inventory

`SmartInventory/Backup/SmartAzure-Backup-Inventory.ps1` exports Recovery Services vaults, vault settings, protected items, protected VMs, unprotected VMs, and summaries.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartInventory\Backup\SmartAzure-Backup-Inventory.ps1 -Tenant prod -Connect
```

### Storage Security Inventory

`SmartInventory/Storage/SmartAzure-StorageSecurity-Inventory.ps1` exports storage account public network access, anonymous blob access, shared key, HTTPS only, minimum TLS, firewall posture, and blob service properties.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartInventory\Storage\SmartAzure-StorageSecurity-Inventory.ps1 -Tenant prod -Connect
```

### Key Vault Security Inventory

`SmartInventory/Security/SmartAzure-KeyVaultSecurity-Inventory.ps1` exports Key Vault public access, purge protection, soft delete, RBAC/access policy mode, access policies, secrets, certificates, and expiring items.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartInventory\Security\SmartAzure-KeyVaultSecurity-Inventory.ps1 -Tenant prod -Connect
```

### Advisor Inventory

`SmartInventory/Advisor/SmartAzure-Advisor-Inventory.ps1` exports Azure Advisor recommendations across all categories and summary counts by category and impact.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartInventory\Advisor\SmartAzure-Advisor-Inventory.ps1 -Tenant prod -Connect
```

## Launchers

Every SmartAzure and SmartM365 SmartInventory script has two adjacent command launchers:

- `Start-<ScriptName>-Prod.cmd`
- `Start-<ScriptName>-Test.cmd`

Launchers call PowerShell 7 when available and pass the matching `-Tenant prod` or `-Tenant test` value.
