# SmartAzure App Registration Permissions

This document explains the permissions added by `Setup/SmartAzure-Create-AppRegistration.ps1`.

SmartAzure does not reuse SmartM365 business permissions. Azure inventory scripts use Azure Resource Manager and must be authorized through Azure RBAC on the target management groups, subscriptions, or resource groups. The SmartAzure app registration is only used for common capabilities shared with SmartM365: certificate authentication, Graph mail notifications, SharePoint upload, and per-tenant local configuration.

## Principles

- Run the bootstrap separately per tenant with `-Tenant <TenantKey>`.
- `AppId`, `TenantId`, `Thumb`, and `Thumbprint` are written to `Config/Tenants/<TenantKey>.local.json`.
- `.local.json` profiles remain local and ignored by Git.
- Application permissions added to the SmartAzure app are limited to Microsoft Graph `Mail.Send` and `Sites.Selected`.
- Teams notifications use `TeamsAlertsWebhookUrl` and `TeamsInfosWebhookUrl` with Teams Workflows / Power Automate. They do not require runtime Graph Teams permissions.
- Azure rights for reading subscriptions, resources, Policy, Defender, Backup, Storage, Key Vault, and Advisor do not come from these Graph permissions. They must be granted through Azure RBAC.

## Microsoft Graph Application Permissions

| Permission | Why it is needed | Used by |
| --- | --- | --- |
| `Mail.Send` | Sends error notifications and HTML reports through Microsoft Graph when `SmtpServer` is empty. The bootstrap can create a shared mailbox, a mail-enabled security group, and an Exchange Online Application Access Policy to restrict this right to `MailSendAccessPolicyGroup`. | `Config/SmartAzure-TenantContext.ps1`; scripts that call the shared mail/notification helpers. |
| `Sites.Selected` | Uploads files to the SmartAzure SharePoint site without granting `Files.ReadWrite.All` or `Sites.ReadWrite.All`. The bootstrap then grants the app the `write` role on the target site. | `Modules/SmartAzure.SharePoint/SmartAzure.SharePoint.psm1`; scripts that enable `EnableSharePointUpload`. |

## Azure RBAC To Plan

SmartAzure scripts query Azure Resource Manager through Az modules. Depending on the exports you need, assign the service principal or execution account suitable Azure roles on the target scope:

| Need | Suggested Azure roles |
| --- | --- |
| General inventory for subscriptions, resource groups, resources, tags, locks, and providers | `Reader` |
| Policy compliance and governance | `Reader`, with Policy read access on the target scope |
| RBAC inventory | `Reader`, or a role that allows `Microsoft.Authorization/*/read` |
| Network exposure | `Reader`, with read access to Network resources |
| Defender for Cloud posture | `Security Reader`, or an equivalent role on the target subscriptions |
| Recovery Services Backup | `Reader`, `Backup Reader` when available/needed |
| Storage security | `Reader`, or `Storage Account Reader` depending on the scope and operations |
| Key Vault security | `Reader`; also review the target vault RBAC/access policy models |
| Advisor | `Reader`, with Advisor read access |

These roles are intentionally separate from the Graph bootstrap: `SmartAzure-Create-AppRegistration.ps1` does not replace Azure RBAC governance.

## Mail.Send Configuration

By default, the bootstrap reuses the SmartM365 model for Graph mail:

- creates or reuses a shared mailbox named `smartazure-reports@<domain>`;
- creates or reuses a mail-enabled security group named `SMART-AZURE-MailSend-Allowed`;
- adds the shared mailbox to the group;
- creates an Exchange Online Application Access Policy to restrict `Mail.Send`;
- writes `From`, an empty `SmtpServer`, and `MailSendAccessPolicyGroup` to the local tenant profile.

Use `-DisableMailSendScopeSetup` to create the app without configuring these Exchange Online objects. In that case, `Mail.Send` remains on the app permissions, but the Exchange Online scope is not configured by the script.

## SharePoint Configuration

By default, the bootstrap reuses the SmartM365 model for SharePoint upload:

- creates or reuses the `SMART-AZURE` Teams team;
- checks the `Alerts` and `Infos` channels;
- resolves the associated SharePoint site;
- writes `SharePointSiteHostname`, `SharePointSitePath`, `SharePointLibraryDisplayName`, and `SharePointTargetFolderPath`;
- grants the SmartAzure app the `write` role on this site with `Sites.Selected`;
- removes old broad grants `Files.ReadWrite.All` and `Sites.ReadWrite.All` if the app already has them.

Use `-DisableTeamsSetup` to skip creating/reusing the team and skip updating SharePoint configuration.

## Delegated Scopes Used By The Bootstrap

These scopes are requested from the administrator account that runs `Setup/SmartAzure-Create-AppRegistration.ps1`. They are not added as runtime permissions to the SmartAzure app.

| Scope | Why it is requested |
| --- | --- |
| `Application.ReadWrite.All` | Creates or updates the app registration and its public certificates. |
| `AppRoleAssignment.ReadWrite.All` | Grants admin consent for the requested application permissions. |
| `Directory.Read.All` | Reads tenant context, service principals, and the setup user. |
| `Group.ReadWrite.All` | Creates or reuses the Microsoft 365 group behind the `SMART-AZURE` Teams team. |
| `Channel.Create` | Creates the standard Teams channels `Alerts` and `Infos`. |
| `Channel.ReadBasic.All` | Checks whether the Teams channels already exist. |
| `Sites.FullControl.All` | Grants the SmartAzure service principal the `write` role on the target site with `Sites.Selected`. |

## Teams Notifications

The bootstrap does not create Teams Workflows URLs. The URLs must be created manually in Teams, then stored in the local tenant profile:

```powershell
.\Setup\SmartAzure-Set-TeamsWebhook.ps1 -Channel Alerts -WebhookUrl "<Alerts workflow URL>"
.\Setup\SmartAzure-Set-TeamsWebhook.ps1 -Channel Infos  -WebhookUrl "<Infos workflow URL>"
```

Webhook URLs are secrets and must not be published.
