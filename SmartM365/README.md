# SmartM365

PowerShell automation scripts for Microsoft 365 and endpoint administration.

This repository contains reusable scripts and helper modules.

## Content

- Microsoft 365 inventory scripts for users, licensing, domains, and Entra devices.
- Intune inventory scripts for managed devices, applications, Autopilot, RBAC, and Windows Update reporting.
- Exchange Online and Exchange on-premises inventory scripts.
- Active Directory inventory and reporting scripts.
- Intune detection and remediation packages organized by scenario.

## Azure App Registration

Use `SmartM365-Create-AppRegistration.ps1` to create or update the app registration used by SmartM365 app-only automation. The script grants admin consent by default; use `-DisableGrantAdminConsent` only when you want to add the permissions without granting consent immediately.

The setup run writes a text log and PowerShell transcript to `C:\Temp\WORKPLACE` by default. Use `-LogPath <folder>` to store them elsewhere.

At startup, the setup disconnects any existing Microsoft Graph and Exchange Online sessions before signing in again. This avoids silently reusing a cached account from WAM or a previous PowerShell session; use `-UseDeviceCode` when you want to explicitly choose the administrator account used for Graph consent and setup. This convention applies to every SmartM365 script that connects to Microsoft Graph or Exchange Online.

When `SmartM365.global.local.json` already contains `Thumbprint` or `Thumb`, the setup reuses that local certificate on later `-UpdateExisting` runs instead of creating a new certificate.

The bootstrap also configures Graph mail sending: it connects to Exchange Online before Microsoft Graph, creates or reuses a dedicated SmartM365 sender shared mailbox, creates or reuses the `SMART-M365-MailSend-Allowed` mail-enabled security group, adds the sender mailbox to that group, creates an Exchange Online Application Access Policy that restricts `Mail.Send` to that group, and writes `From`, `SmtpServer`, and `MailSendAccessPolicyGroup` to `SmartM365.global.local.json`. Use `-ExchangeAdminUserPrincipalName` if Exchange Online should sign in with a specific admin account.

Use `-RemoveAppRegistration -Confirm` to remove the SmartM365 app registration and its application service principals, remove related Exchange Online Application Access Policies, then clear the app-only authentication values from `SmartM365.global.local.json`. This cleanup does not remove the `SMART-M365` Teams workspace, SharePoint files, sender mailbox, Mail.Send scope group, or local certificates.

The same bootstrap creates or reuses the `SMART-M365` Teams team, resolves its SharePoint site, and updates `SmartM365.global.local.json` with the SharePoint upload target. Use `-DisableTeamsSetup` only when the Teams workspace is already handled separately.

Current Microsoft Graph application permissions include the read scopes used by the inventory scripts (`Directory.Read.All`, `User.Read.All`, `Device.Read.All`, `GroupMember.Read.All`, Intune read permissions, and `AuditLog.Read.All` for user `signInActivity`), plus `Sites.Selected` for SharePoint CSV upload and `Mail.Send` for Graph mail. The bootstrap grants the app `write` access only on the SmartM365 SharePoint site and removes older broad `Files.ReadWrite.All` / `Sites.ReadWrite.All` grants when found. Exchange Online app-only inventory automation also needs `Exchange.ManageAsApp` and separate Exchange RBAC assignment.

See `SmartM365-AppRegistration-Permissions.md` for the permission-by-permission rationale and the scripts that use each permission.
