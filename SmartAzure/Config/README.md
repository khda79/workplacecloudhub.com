# SmartAzure Config

Configuration helpers and local tenant profile templates for SmartAzure.

## Files

- `SmartAzure-TenantContext.ps1`: loads the SmartAzure root, local tenant profile, tokenized path configuration, script-local configuration, notifications, and Data/Output fallback behavior.
- `Tenants/tenant.local.json.template`: safe template for local tenant profile files.

Real tenant profiles must be named `Tenants/<TenantKey>.local.json` and stay ignored by Git.

## Configuration Order

SmartAzure follows the same model as SmartM365:

1. explicit script parameters;
2. script-local file named `<ScriptName>.local.json` in the same folder as the script;
3. tenant/global configuration loaded from `SmartAzure.global.local.json` and `Config/Tenants/<TenantKey>.local.json`;
4. script defaults.

Shared keys intentionally match SmartM365 where the same capability is used: `AppId`, `TenantId`, `Thumb`, `Thumbprint`, `From`, `MailSendAccessPolicyGroup`, `EnableSharePointUpload`, `SharePointSiteHostname`, `SharePointSitePath`, `SharePointLibraryDisplayName`, `SharePointTargetFolderPath`, `TeamsWebhookUrl`, `TeamsAlertsWebhookUrl`, and `TeamsInfosWebhookUrl`.

`AzureTenantId` is accepted as an Azure sign-in alias, but `TenantId` remains present for Graph app-only features such as mail and SharePoint upload.
