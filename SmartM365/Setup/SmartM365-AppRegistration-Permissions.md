# SmartM365 App Registration Permissions

This document explains the permissions added by `Setup/SmartM365-Create-AppRegistration.ps1`, why they are needed, and which scripts currently use them.

## Important Notes

- The permissions below are Microsoft Graph **Application** permissions for the SmartM365 app registration.
- `Setup/SmartM365-Create-AppRegistration.ps1` runs only with interactive delegated authentication: an administrator signs in with the rights needed to create or update the app registration.
- `Setup/SmartM365-Create-AppRegistration.ps1` is a special setup script. Its interactive setup rights are not the runtime privilege baseline for inventory scripts.
- In multi-tenant mode, run the bootstrap separately per tenant with `-Tenant <TenantKey>`. App, certificate, SharePoint, mail, and Teams values remain in `Config/Tenants/<TenantKey>.local.json`, never in Git.
- App-only scripts use `https://graph.microsoft.com/.default` or `Connect-MgGraph -ClientId ...`; the effective permissions are therefore the permissions granted to the application in Entra ID.
- Scripts with `-InteractiveAuth` may request equivalent delegated scopes, but the app registration should primarily carry application permissions for unattended runs.
- `Intune/Remediation/CLI/SmartM365-Deploy-IntuneRemediation-CLI.ps1` is intentionally outside the app-only model: it uses only interactive delegated authentication and requests `DeviceManagementScripts.ReadWrite.All` to create or update Intune remediations. `Intune/Remediation/GUI/SmartM365-IntuneRemediation-GUI.ps1` follows the same interactive delegated model, uses the tenant selected during interactive sign-in, and also requests `DeviceManagementConfiguration.Read.All` to export Intune execution reports and `Group.Read.All` to enrich assignment exports with Entra group names.
- SharePoint uploads are centralized through `Modules/SmartM365.Core` and `Modules/SmartM365.SharePoint`; they can be used by multiple scripts when `EnableSharePointUpload` is enabled.
- Teams notifications use `EnableTeamsNotifications`, `TeamsAlertsWebhookUrl`, `TeamsInfosWebhookUrl`, and `Send-SmartM365TeamsNotification` with Teams Workflows / Power Automate URLs, not legacy Office 365 Connectors. They do not add Graph Teams permissions to the app registration. Do not use `Teamwork.Migrate.All` for normal operational notifications. On terminal errors, every inventory/report script should send a notification to the `Alerts` channel with diagnostic context and an AI help link built from the error message. On successful completion, each script should send a notification to the `Infos` channel with a `Result summary` field that summarizes the script result.
- Exchange user notification campaigns can optionally send a one-on-one Teams chat message to each recipient with `TeamsUserMessageMode = GraphDelegated`. This is different from operational Teams channel notifications. It uses Microsoft Graph **delegated** permissions `User.Read`, `Chat.Create`, and `ChatMessage.Send`; the visible Teams sender is the delegated account used during the campaign run. These delegated chat permissions are not app-only runtime permissions and are not covered by the standard SmartM365 application permission set.

## User Configuration For Teams Notifications

`Setup/SmartM365-Create-AppRegistration.ps1` creates or reuses the `SMART-M365` team and the `Alerts` / `Infos` channels, but it does not create Teams Workflows URLs. Those URLs belong to the user's Teams / Power Automate context and must be created manually in Microsoft Teams.

Expected procedure:

1. Open Microsoft Teams.
2. Go to the `SMART-M365` team and the `Alerts` channel.
3. Open `Workflows` from the channel menu.
4. Search for `webhook`.
5. Create a workflow such as `Send webhook alerts to a channel`, or one based on `When a Teams webhook request is received`.
6. Select the `SMART-M365` team and the `Alerts` channel, then copy the generated HTTP POST URL.
7. Repeat the same operation in the `Infos` channel.
8. Save and test the URLs with:

```powershell
.\Setup\SmartM365-Set-TeamsWebhook.ps1 -Channel Alerts -WebhookUrl "<Alerts workflow URL>"
.\Setup\SmartM365-Set-TeamsWebhook.ps1 -Channel Infos  -WebhookUrl "<Infos workflow URL>"
```

The URLs and `EnableTeamsNotifications` switch are stored only in `Config/Tenants/<TenantKey>.local.json` with the keys `TeamsAlertsWebhookUrl` and `TeamsInfosWebhookUrl`. This file is local and ignored by Git. Webhook URLs must be treated as secrets: do not publish them in a commit, issue, document, or public channel.

This approach avoids adding broad Microsoft Graph Teams permissions to the SmartM365 app. Operational notifications go through Power Automate; the only Teams permissions requested by the bootstrap are delegated setup scopes used to create/check the `Alerts` and `Infos` channels.

## Microsoft Graph Permissions

| Permission | Why it is needed | Scripts / modules using it |
| --- | --- | --- |
| `Directory.Read.All` | Broad Entra ID directory read access: organization, users, groups, licenses, domains, devices, sync attributes, and directory enrichment. Several Graph endpoints respond better, or only respond, with this permission in tenant inventories. | `SmartInventory/M365Inventory/Users/SmartM365-ActiveUsers-Inventory.ps1`; `SmartInventory/M365Inventory/Licensing/SmartM365-Licences-Inventory.ps1`; `SmartInventory/M365Inventory/Domains/SmartM365-VerifiedDomains-Inventory.ps1`; `SmartInventory/M365Inventory/Devices/SmartM365-EntraDevices-Inventory.ps1`; `SmartInventory/M365Inventory/IntuneInventory/RBAC/SmartM365-Intune-RBAC-GroupMembers.ps1`; `SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Compliance-Inventory.ps1`; `SmartInventory/ExchangeInventory/Mailboxes/SmartM365-EXO-Mailboxes-Inventory.ps1`; `SmartInventory/ExchangeInventory/CalendarPermissions/SmartM365-EXO-Mailboxes-CalPerm_Inventory.ps1`. |
| `User.Read.All` | Reads users and their properties for M365 exports and Exchange Online enrichment. | `SmartInventory/M365Inventory/Users/SmartM365-ActiveUsers-Inventory.ps1`; `SmartInventory/M365Inventory/Licensing/SmartM365-Licences-Inventory.ps1`; `SmartInventory/ExchangeInventory/Mailboxes/SmartM365-EXO-Mailboxes-Inventory.ps1`; `SmartInventory/ExchangeInventory/BackupProtection/SmartM365-M365-BackupProtectedMailboxes-Inventory.ps1`. |
| `AuditLog.Read.All` | Reads sign-in activity exposed by Graph, especially `signInActivity` on users for last sign-in columns. | `SmartInventory/M365Inventory/Users/SmartM365-ActiveUsers-Inventory.ps1`. |
| `BackupRestore-Configuration.Read.All` | Reads Microsoft 365 Backup protection units for mailbox backup protection inventory. | `SmartInventory/ExchangeInventory/BackupProtection/SmartM365-M365-BackupProtectedMailboxes-Inventory.ps1`. |
| `Reports.Read.All` | Reads Microsoft 365 usage reports, including active user, mailbox usage, OneDrive usage, SharePoint site usage, Office activations, Teams activity, and email activity reports for FinOps and adoption analysis. | `SmartInventory/M365Inventory/Usage/SmartM365-M365UserActivity-Inventory.ps1`. |
| `Device.Read.All` | Reads Entra ID device objects, especially to link Intune/Autopilot data to Entra objects and retrieve attributes such as deviceId, trustType, or approximate sign-in data. | `SmartInventory/M365Inventory/Devices/SmartM365-EntraDevices-Inventory.ps1`; `SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Inventory.ps1`; `SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Compliance-Inventory.ps1`; `SmartInventory/M365Inventory/IntuneInventory/Autopilot/SmartM365-WindowsAutopilot-Inventory.ps1`. |
| `GroupMember.Read.All` | Reads Entra ID group members, useful for RBAC exports and group-based comparisons. | `SmartInventory/M365Inventory/IntuneInventory/RBAC/SmartM365-Intune-RBAC-GroupMembers.ps1`; `SmartInventory/ExchangeInventory/BackupProtection/SmartM365-M365-BackupProtectedMailboxes-Inventory.ps1`. |
| `DeviceManagementApps.Read.All` | Reads detected apps / Intune apps through `/deviceManagement/detectedApps` and related relationships. | `SmartInventory/M365Inventory/IntuneInventory/Applications/SmartM365-Intune-DiscoveredApps-Inventory.ps1`. |
| `DeviceManagementConfiguration.Read.All` | Reads Intune configurations, policies, and reports, including compliance policies, Windows Update for Business, Feature Update / Quality Update profiles, and report export jobs. | `SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Device-System-Inventory.ps1`; `SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Compliance-Inventory.ps1`; `SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/SmartM365-WinUpdate_Status_From_Intune.ps1`; `SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Get-IntuneAutopatchAlerts.ps1`. |
| `DeviceManagementManagedDevices.Read.All` | Reads Intune managed devices and their properties: device inventory, BIOS, compliance, system data, upgrade eligibility, and Endpoint Analytics. | `SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Inventory.ps1`; `SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-BIOS-Inventory.ps1`; `SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Compliance-Inventory.ps1`; `SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Device-System-Inventory.ps1`; `SmartInventory/M365Inventory/IntuneInventory/Devices/SmartM365-Devices-UpgradeEligibility.ps1`; `SmartInventory/M365Inventory/IntuneInventory/Applications/SmartM365-Intune-DiscoveredApps-Inventory.ps1`; `SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/SmartM365-WinUpdate_Status_From_Intune.ps1`; `SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Get-IntuneAutopatchAlerts.ps1`. |
| `DeviceManagementScripts.Read.All` | Reads/exports Intune remediation scripts exposed as `deviceHealthScripts`, including details and assignments. | `SmartInventory/M365Inventory/IntuneInventory/SmartM365-Export-IntuneRemediations.ps1`. |
| `DeviceManagementServiceConfig.Read.All` | Reads Intune service configuration, especially Windows Autopilot identities. | `SmartInventory/M365Inventory/IntuneInventory/Autopilot/SmartM365-WindowsAutopilot-Inventory.ps1`. |
| `Mail.Send` | Sends error notifications and HTML reports through Microsoft Graph when `SmtpServer` is empty. The sender address is resolved from `From` in the tenant profile or script `*.local.json`. The bootstrap also creates an Exchange Online Application Access Policy to restrict this right to `MailSendAccessPolicyGroup`. | `Modules/SmartM365.Core/SmartM365.Core.psm1`; scripts that call `SendEmailHtmlReport`, `Send-SmartM365Mail`, or `SendFileListEmailReport`. |
| `Sites.Selected` | Uploads and replaces CSV files only on the SmartM365 SharePoint site. The bootstrap then grants the app the `write` role on the created/reused site, and removes old broad grants `Files.ReadWrite.All` and `Sites.ReadWrite.All` when they exist. | `Modules/SmartM365.Core/SmartM365.Core.psm1`; `Modules/SmartM365.SharePoint/SmartM365.SharePoint.psm1`; all inventory/export scripts that call SharePoint upload when `EnableSharePointUpload` is enabled. |

Exchange user notification campaigns under `Communications/ExchangeUserNotifications` also use `Mail.Send` when `MailSendMode = Graph` or when `MailSendMode = Auto` resolves to Graph because no SMTP relay is configured. If `MailSendMode = SmtpRelay`, mail delivery uses the configured relay instead of Microsoft Graph.

## Intune ReadWrite Permissions Included By Default

These permissions are added by default because one current script still requests them during interactive authentication. Treat them as transitional: run `Setup/SmartM365-Create-AppRegistration.ps1` with `-SkipBroadIntuneReadWritePermissions` once the related scripts are hardened for read-only use.

| Permission | Why it is present today | Scripts / modules using it |
| --- | --- | --- |
| `DeviceManagementApps.ReadWrite.All` | Broad permission requested by the current Autopatch script to access Intune/reporting data. Reduce it if the calls remain read-only. | `SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Get-IntuneAutopatchAlerts.ps1`. |
| `DeviceManagementConfiguration.ReadWrite.All` | Broad permission requested by the current Autopatch script for Windows Update profiles and export jobs. Replace it with `DeviceManagementConfiguration.Read.All` if no privileged write operation is needed. | `SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Get-IntuneAutopatchAlerts.ps1`. |
| `DeviceManagementManagedDevices.ReadWrite.All` | Broad permission requested by the current Autopatch script for managed device/reporting data. Replace it with `DeviceManagementManagedDevices.Read.All` if no device change is performed. | `SmartInventory/M365Inventory/IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Get-IntuneAutopatchAlerts.ps1`. |

## Interactive Delegated Scope For Intune Deployment

This scope is not a SmartM365 runtime application permission. It is requested from the administrator account that runs the deployment script in interactive mode.

| Delegated scope | Why it is needed | User script |
| --- | --- | --- |
| `DeviceManagementScripts.ReadWrite.All` | Creates, updates, and optionally assigns Intune remediation packages exposed by Microsoft Graph as `deviceHealthScripts`; the local GUI also uses it to list, duplicate, and delete Intune packages. | `Intune/Remediation/CLI/SmartM365-Deploy-IntuneRemediation-CLI.ps1`; `Intune/Remediation/GUI/SmartM365-IntuneRemediation-GUI.ps1`. |
| `DeviceManagementConfiguration.Read.All` | Exports the execution report CSV for a remediation through Intune report export jobs. | `Intune/Remediation/GUI/SmartM365-IntuneRemediation-GUI.ps1`. |
| `Group.Read.All` | Reads target Entra group names so `targetGroupName` can be added to `Assignments.json` during `Save all cloud`. | `Intune/Remediation/GUI/SmartM365-IntuneRemediation-GUI.ps1`. |

## Non-Graph Permission Added By The Script

| Permission | Why it is needed | Scripts / modules using it |
| --- | --- | --- |
| `Exchange.ManageAsApp` on the `Office 365 Exchange Online` API | Enables certificate-based app-only authentication with `Connect-ExchangeOnline`. This is not a Microsoft Graph permission. It is not sufficient alone: the EXO app-only token must also contain a supported Entra role for Exchange RBAC to be built. | Exchange Online scripts under `SmartInventory/ExchangeInventory`, for example `AcceptedDomains`, `BackupProtection`, `CalendarPermissions`, `Mailboxes`, and `Migration`, when they use app-only authentication. |

## Entra ID Role Assigned To The SmartM365 Service Principal

| Role | Why it is needed | Scripts / modules using it |
| --- | --- | --- |
| `Global Reader` | Privileged baseline for read-only Exchange Online runtime scripts. This Entra role is supported by Exchange Online PowerShell app-only auth and is normally enough for inventories and reports that only call `Get-*` cmdlets. | Exchange Online scripts under `SmartInventory/ExchangeInventory`, for example `AcceptedDomains`, `BackupProtection`, `CalendarPermissions`, `Mailboxes`, and `Migration`, as long as they remain read-only. |
| `Exchange Administrator` | Setup role, not runtime baseline. It may be required for the interactive administrator account that runs `Setup/SmartM365-Create-AppRegistration.ps1`, because the bootstrap creates/modifies a shared mailbox, a mail-enabled group, an Application Access Policy, and the service principal role assignment. | `Setup/SmartM365-Create-AppRegistration.ps1` only, or a future Exchange Online script that explicitly performs changes. |

## Exchange Online Configuration Created By The Bootstrap

| Object | Why it is needed | Scripts / modules using it |
| --- | --- | --- |
| Shared mailbox `smartm365-reports@<domain>` | Provides a dedicated sender mailbox for SmartM365 reports and notifications. The bootstrap writes this address to `From`. | `Modules/SmartM365.Core/SmartM365.Core.psm1`; scripts that send reports or errors through Graph. |
| Mail-enabled security group `SMART-M365-MailSend-Allowed` | Lists mailboxes authorized for the SmartM365 application with `Mail.Send`. The bootstrap writes its address to `MailSendAccessPolicyGroup`. | Exchange Online restriction applied to `Mail.Send`. |
| Exchange Online Application Access Policy | Restricts the SmartM365 application to mailboxes that are members of `MailSendAccessPolicyGroup` for Outlook Graph permissions such as `Mail.Send`. Without this restriction, application `Mail.Send` is tenant-wide. | Graph sends performed by `Modules/SmartM365.Core/SmartM365.Core.psm1`. |

## Permissions Used Only To Run The Bootstrap

These scopes are requested from the administrator who runs `Setup/SmartM365-Create-AppRegistration.ps1`. They are not added to the SmartM365 app as business permissions; they are used to create/modify the app registration and grant consent.

`Setup/SmartM365-Create-AppRegistration.ps1` is an interactive bootstrap script: it is intentionally more privileged than runtime inventory scripts. Roles required by this setup must not be copied as prerequisites for read-only Exchange Online scripts.

| Connection scope | Why it is requested |
| --- | --- |
| `Application.ReadWrite.All` | Creates or updates the app registration, its API permissions, and its public certificates. |
| `AppRoleAssignment.ReadWrite.All` | Grants admin consent as app role assignments. `Setup/SmartM365-Create-AppRegistration.ps1` does this by default; use `-DisableGrantAdminConsent` only to prepare the app without immediate consent. |
| `Channel.Create` | Creates the standard Teams channels `Alerts` and `Infos` in the `SMART-M365` team when the bootstrap initializes them. |
| `Channel.ReadBasic.All` | Checks whether the standard Teams channels `Alerts` and `Infos` already exist before creating them. |
| `Directory.Read.All` | Reads Microsoft Graph and Exchange Online API service principals, verifies tenant context, and finds the connected administrator user. |
| `Group.ReadWrite.All` | Creates or reuses the Microsoft 365 group behind the `SMART-M365` Teams team, then converts this group into a team. |
| `RoleManagement.ReadWrite.Directory` | Assigns the Entra `Global Reader` role to the SmartM365 service principal for Exchange Online app-only auth, and removes the old `Exchange Administrator` role from the service principal if a previous bootstrap version added it. This scope is used only by the administrator running the bootstrap and is not added to the SmartM365 app. |
| `Sites.FullControl.All` | Assigns the `write` role to the SmartM365 service principal on the target SharePoint site with `Sites.Selected`. This scope is used only by the administrator running the bootstrap and is not added to the SmartM365 app. |

## Delegated Permissions For User-Facing Teams Campaign Messages

These scopes are requested only when an operator enables per-user Teams chat messages for Exchange notification campaigns. They are delegated scopes for the signed-in sender account, not application permissions for unattended SmartM365 automation.

| Permission | Why it is needed |
| --- | --- |
| `User.Read` | Resolves the signed-in delegated sender with `/me`. |
| `Chat.Create` | Creates or reuses the one-on-one chat between the sender account and the recipient. |
| `ChatMessage.Send` | Posts the Teams message into that one-on-one chat. |

Do not add `Teamwork.Migrate.All` for normal campaign messages. That permission is for Teams migration/import scenarios, not user-facing operational communication.

## Future Review Items

- Replace Intune `ReadWrite` permissions with `Read` permissions once `SmartM365-Get-IntuneAutopatchAlerts.ps1` is confirmed read-only.
- Verify that all SharePoint uploads remain compatible with `Sites.Selected`; reintroduce `Files.ReadWrite.All` or `Sites.ReadWrite.All` only as a documented last resort.
- Review later whether the Entra `Global Reader` role can be replaced with a more granular Exchange RBAC assignment once the exact needs of EXO scripts are stable.
- If the Teams SharePoint site is not immediately available, rerun the bootstrap later: Teams site provisioning is asynchronous on the Microsoft 365 side.
