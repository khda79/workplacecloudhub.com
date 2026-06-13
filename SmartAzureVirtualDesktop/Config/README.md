# SmartAzureVirtualDesktop Config

Configuration helpers and local tenant profile templates for SmartAzureVirtualDesktop.

## Files

- `SmartAzureVirtualDesktop-TenantContext.ps1`: loads the project root, local tenant profile, tokenized path configuration, script-local configuration, and Data/Output fallback behavior.
- `Tenants/tenant.local.json.template`: safe template for local tenant profile files.

Real tenant profiles must be named `Tenants/<TenantKey>.local.json` and stay ignored by Git.

## Configuration Order

SmartAzureVirtualDesktop follows the same local model as SmartAzure:

1. explicit script parameters;
2. script-local file named `<ScriptName>.local.json` in the same folder as the script;
3. tenant/global configuration loaded from `SmartAzureVirtualDesktop.global.local.json` and `Config/Tenants/<TenantKey>.local.json`;
4. script defaults.

`AzureTenantId` is accepted as an Azure sign-in alias. `TenantId` remains available for shared configuration compatibility.
