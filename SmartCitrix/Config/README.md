# SmartCitrix Config

Configuration helpers and local tenant profile templates for SmartCitrix.

## Files

- `SmartCitrix-TenantContext.ps1`: loads the SmartCitrix root, local tenant profile, tokenized path configuration, script-local configuration, and Data/Output fallback behavior.
- `Tenants/tenant.local.json.template`: safe template for local tenant profile files.

Real tenant profiles must be named `Tenants/<TenantKey>.local.json` and stay ignored by Git.

## Configuration Order

SmartCitrix follows the same local model as SmartAzureVirtualDesktop:

1. explicit script parameters;
2. script-local file named `<ScriptName>.local.json` in the same folder as the script;
3. tenant/global configuration loaded from `SmartCitrix.global.local.json` and `Config/Tenants/<TenantKey>.local.json`;
4. script defaults.

`AdminAddress` can be stored locally for on-premises CVAD scripts. Citrix Cloud credentials must be kept out of committed files and should be added only to ignored local tenant profiles when Citrix Cloud scripts are implemented.
