# SmartM365 Setup

This folder contains interactive setup and bootstrap scripts for SmartM365.

- `SmartM365-Create-AppRegistration.ps1`: creates or updates the SmartM365 Entra app registration, permissions, certificate metadata, Teams workspace, SharePoint target, and mail sender setup.
- `SmartM365-AppRegistration-Permissions.md`: documents the permission-by-permission rationale and the scripts that use each permission.
- `SmartM365-Set-TeamsWebhook.ps1`: stores and tests Teams Workflows / Power Automate webhook URLs in the selected local tenant profile.

Run these scripts from the SmartM365 root with `.\Setup\<script-name>.ps1` so relative examples and local configuration paths remain easy to read.

The shared tenant-context helper lives in `..\Config\SmartM365-TenantContext.ps1`.
