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

The same bootstrap creates or reuses the `SMART-M365` Teams team, resolves its SharePoint site, and updates `SmartM365.global.local.json` with the SharePoint upload target. Use `-DisableTeamsSetup` only when the Teams workspace is already handled separately.

See `SmartM365-AppRegistration-Permissions.md` for the permission-by-permission rationale and the scripts that use each permission.
