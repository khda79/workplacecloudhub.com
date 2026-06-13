# SmartM365

PowerShell automation scripts for Microsoft 365 and endpoint administration.

This repository contains reusable scripts and helper modules.

## Content

- Smart Inventory scripts under `SmartInventory/` for Active Directory, Exchange, Microsoft 365, Entra, and Intune data collection. These exports can feed Power BI datasets and other reporting or operational consumers.
- Device and endpoint tooling under `Devices/`, including reboot governance, device registration diagnostics, endpoint diagnostics, Intune remediation management, and Hybrid Join repair tooling.

## Devices

`Devices/` groups SmartM365 tools that run on, diagnose, repair, notify, or manage Windows endpoints.

- `Devices/DeviceRebootManager/`: localized PowerShell 5.1 WPF restart notification app.
- `Devices/DeviceRegistrationTool/`: PowerShell/WPF GUI and CLI mode for Intune enrollment, Hybrid Join, and Entra device registration checks.
- `Devices/EndpointDiagnosticsAnalyzer/`: PowerShell/WPF GUI for Intune Device Diagnostics ZIP files and local endpoint captures.
- `Devices/SmartIntuneHybridJoinToolkit/`: PsExec/LOT orchestration for Hybrid Entra Join and Intune enrollment repair.
- `Devices/SmartIntuneRemediation/`: Intune remediation scripts and the remediation manager grouped together.

## Setup

`Setup/` contains the interactive SmartM365 setup scripts. Use `Setup/SmartM365-Create-AppRegistration.ps1` to create or update the app registration used by SmartM365 app-only automation. The script grants admin consent by default; use `-DisableGrantAdminConsent` only when you want to add the permissions without granting consent immediately.

The setup run writes a text log and PowerShell transcript under `Data\Tenants\<TenantKey>\LOG-ALL\Setup` by default. If the root `Data` folder is not writable, it falls back to `Setup\Output\Tenants\<TenantKey>\LOG-ALL\Setup`. Use `-LogPath <folder>` to store them elsewhere.

At startup, the setup disconnects any existing Microsoft Graph and Exchange Online sessions before signing in again. This avoids silently reusing a cached account from WAM or a previous PowerShell session; use `-UseDeviceCode` when you want to explicitly choose the administrator account used for Graph consent and setup. This convention applies to every SmartM365 script that connects to Microsoft Graph or Exchange Online.

When the selected tenant profile already contains `Thumbprint` or `Thumb`, the setup reuses that local certificate on later `-UpdateExisting` runs instead of creating a new certificate.

The bootstrap also configures Graph mail sending: it connects to Exchange Online before Microsoft Graph, creates or reuses a dedicated SmartM365 sender shared mailbox, creates or reuses the `SMART-M365-MailSend-Allowed` mail-enabled security group, adds the sender mailbox to that group, creates an Exchange Online Application Access Policy that restricts `Mail.Send` to that group, and writes `From`, `SmtpServer`, and `MailSendAccessPolicyGroup` to the selected tenant profile. Use `-ExchangeAdminUserPrincipalName` if Exchange Online should sign in with a specific admin account.

Teams notifications use Teams Workflows / Power Automate webhook URLs instead of Microsoft Graph chat permissions or legacy Office 365 Connectors. The bootstrap creates or reuses two standard channels in the `SMART-M365` team: `Alerts` for script errors and `Infos` for successful completion or informational notices. The webhook URLs are still created by the user in Teams / Workflows, then stored locally with `Setup/SmartM365-Set-TeamsWebhook.ps1`.

Every SmartM365 inventory/report script should send a Teams notification when it fails, in addition to the existing error email flow, and should send an `Infos` notification when it completes without error. `Infos` cards must include a `Result summary` fact with a short result recap, ideally with the main counters, generated files, skipped items, or actions performed by the script. Error notifications must include enough detail to act without opening the host first: script name, tenant or organization when known, computer name, timestamp, failed phase or operation, exception message, inner exception details when available, log file path, transcript path when available, and output path or CSV path when relevant. The Teams card should also include an AI help link built from the error context, for example a prefilled ChatGPT or Copilot prompt asking for troubleshooting help with the exact SmartM365 script, operation, and error text.

Use `-RemoveAppRegistration -Confirm` to remove the SmartM365 app registration and its application service principals, remove related Exchange Online Application Access Policies, then clear the app-only authentication values from the selected tenant profile. This cleanup does not remove the `SMART-M365` Teams workspace, SharePoint files, sender mailbox, Mail.Send scope group, or local certificates.

The same bootstrap creates or reuses the `SMART-M365` Teams team, resolves its SharePoint site, and updates the selected tenant profile with the SharePoint upload target. Use `-DisableTeamsSetup` only when the Teams workspace is already handled separately.

Current Microsoft Graph application permissions include the read scopes used by the inventory scripts (`Directory.Read.All`, `User.Read.All`, `Device.Read.All`, `GroupMember.Read.All`, Intune read permissions, and `AuditLog.Read.All` for user `signInActivity`), plus `Sites.Selected` for SharePoint CSV upload and `Mail.Send` for Graph mail. The bootstrap grants the app `write` access only on the SmartM365 SharePoint site and removes older broad `Files.ReadWrite.All` / `Sites.ReadWrite.All` grants when found. Exchange Online app-only runtime inventory uses `Exchange.ManageAsApp` plus the supported Entra `Global Reader` role on the SmartM365 service principal; the bootstrap also removes the older `Exchange Administrator` service-principal role when found. `Setup/SmartM365-Create-AppRegistration.ps1` is separate because it is an interactive setup script and can require an Exchange Administrator account for mailbox, group, Application Access Policy, and service-principal role setup.

`Devices/SmartIntuneRemediation/IntuneRemediationManager/IntuneRemediationManager-CLI/SmartM365-Deploy-IntuneRemediation-CLI.ps1` is also intentionally interactive only. It deploys Intune remediation packages through Microsoft Graph `deviceHealthScripts` with delegated `DeviceManagementScripts.ReadWrite.All`; it does not use SmartM365 app-only certificate authentication. `Devices/SmartIntuneRemediation/IntuneRemediationManager/IntuneRemediationManager-GUI/SmartM365-IntuneRemediation-GUI.ps1` follows the same delegated-only model, uses the tenant selected during interactive sign-in, and also requests `DeviceManagementConfiguration.Read.All` to export Intune execution reports through report export jobs.

See `Setup/SmartM365-AppRegistration-Permissions.md` for the permission-by-permission rationale and the scripts that use each permission.

## Multi-Tenant Configuration

SmartM365 supports several tenants through local tenant profiles:

```text
Config\Tenants\test.local.json
Config\Tenants\prod.local.json
```

These files are ignored by Git. Use `Config\Tenants\tenant.local.json.template` as the safe committed model for new tenants.

Pass the target tenant directly when running a script. If omitted, scripts use `test`.

```powershell
.\SmartInventory\M365Inventory\Users\SmartM365-ActiveUsers-Inventory.ps1 -Tenant test
.\SmartInventory\M365Inventory\Users\SmartM365-ActiveUsers-Inventory.ps1 -Tenant prod
```

Each script loads `SmartM365.global.local.json`, overlays `Config\Tenants\<Tenant>.local.json` in memory, and does not rewrite the global JSON. Output roots include `TenantKey` so test and production exports do not overwrite each other:

```text
{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-ALL
{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\DATA-LAST
{{WorkspaceRootPath}}\Data\Tenants\{{TenantKey}}\LOG-ALL
```

By default, `WorkspaceRootPath` is the SmartM365 project root. If `Data` at the project root cannot be written, scripts fall back to an `Output` folder next to the running script.

Run `Setup/SmartM365-Create-AppRegistration.ps1` separately for each tenant that must be bootstrapped:

```powershell
.\Setup\SmartM365-Create-AppRegistration.ps1 -Tenant prod
```

The bootstrap is interactive and should keep tenant-specific app, SharePoint, mail, and Teams values in `Config\Tenants\<Tenant>.local.json`.

## Teams Notifications Setup

SmartM365 uses two Teams channels:

- `Alerts`: errors, failed runs, blocked scripts, and diagnostics that require action.
- `Infos`: successful script completion and informational notices.

Run `Setup/SmartM365-Create-AppRegistration.ps1` first. It creates or reuses the `SMART-M365` team and the `Alerts` / `Infos` channels. It does not create the Power Automate webhook URLs because Teams Workflows are created in the user's Microsoft 365 / Power Platform context.

Create the Teams Workflow URL for `Alerts`:

1. Open Microsoft Teams.
2. Go to the `SMART-M365` team, then the `Alerts` channel.
3. Open `Workflows` from the channel menu.
4. Search for `webhook`.
5. Select the workflow template named like `Send webhook alerts to a channel` or the template based on `When a Teams webhook request is received`.
6. Name it `SmartM365 Alerts Webhook`.
7. Select the `SMART-M365` team and the `Alerts` channel.
8. Create the workflow, then copy the generated HTTP POST URL.

Repeat the same steps for the `Infos` channel, with a workflow name such as `SmartM365 Infos Webhook`, and select the `Infos` channel.

Store and test the URLs from PowerShell:

```powershell
cd %SMARTM365_ROOT%

.\Setup\SmartM365-Set-TeamsWebhook.ps1 -Channel Alerts -WebhookUrl "<Alerts workflow URL>"
.\Setup\SmartM365-Set-TeamsWebhook.ps1 -Channel Infos  -WebhookUrl "<Infos workflow URL>"
```

The script writes the URLs to the selected tenant profile as `TeamsAlertsWebhookUrl` and `TeamsInfosWebhookUrl`, then sends a test card to the selected channel. The local JSON file is ignored by Git and must never be committed.

Useful options:

- `-TestOnly`: send the test card without saving the URL.
- `-SkipTest`: save the URL without sending a test card.
- `-Channel Default`: store a fallback `TeamsWebhookUrl` for older or generic notification usage.

Do not use legacy Office 365 Connector incoming webhook URLs for new SmartM365 notifications, and do not add Microsoft Graph `Teamwork.Migrate.All` for normal operational notifications. If a webhook URL is pasted in a ticket, public chat, commit, or documentation by mistake, treat it as a secret and rotate it in Teams / Power Automate.

Scripts should call `Send-SmartM365TeamsNotification` from `Modules/SmartM365.Core`:

- `-Level ERROR` goes to `Alerts` by default.
- `-Level SUCCESS`, `INFO`, and `WARNING` go to `Infos` by default.
- `-Channel Alerts` or `-Channel Infos` can be used when a script must force a destination.
- `-ResultSummary` should be provided for every `Infos` notification when the script has meaningful counters or output details. If omitted, the module adds a `Result summary` fact from the message text as a fallback.
- `-HelpUrl` should point to a prefilled AI troubleshooting prompt when reporting a detailed script error.



