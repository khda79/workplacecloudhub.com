# SmartM365 Setup

This folder contains interactive setup and bootstrap scripts for SmartM365.

The [SmartInventory configuration assistant](ConfigurationAssistant/README.md)
provides an English guided configuration flow. It has synthetic offline coverage;
qualification on a test machine and tenant is pending.

- `SmartM365-Create-AppRegistration.ps1`: creates or updates the SmartM365 Entra app registration, permissions, certificate metadata, Teams workspace, SharePoint target, and mail sender setup.
- `SmartM365-AppRegistration-Permissions.md`: documents the permission-by-permission rationale and the scripts that use each permission.
- `SmartM365-Set-TeamsWebhook.ps1`: stores and tests Teams Workflows / Power Automate webhook URLs in the selected local tenant profile.
- `Test-SmartM365-SendMail.ps1`: sends one test mail through `Send-SmartM365Mail` with the selected tenant profile; `-From`/`-To`/`-Cc` override the profile, `-AsScript <name>` applies that script's mail profile entry, `-WhatIf` only shows the effective values.
- `Install-SmartM365-SmartInventoryPrerequisites.ps1`: installs a PowerShell 7 module baseline and Python 3.10+ when missing, and reports the RSAT Active Directory prerequisite when missing. Other host and workload prerequisites remain separate.

Run these scripts from the SmartM365 root with `.\Setup\<script-name>.ps1` so relative examples and local configuration paths remain easy to read.

The shared tenant-context helper lives in `..\Config\SmartM365-TenantContext.ps1`.

The app-registration bootstrap logs and transcripts default to `Data\Tenants\<TenantKey>\LOG-ALL\Setup`, with fallback to `Setup\Output\Tenants\<TenantKey>\LOG-ALL\Setup`. The prerequisite installer writes to the console.

Install SmartInventory prerequisites:

```powershell
.\Setup\Install-SmartM365-SmartInventoryPrerequisites.ps1 -TrustRepository -AllowClobber
```

Audit without installing:

```powershell
.\Setup\Install-SmartM365-SmartInventoryPrerequisites.ps1 -AuditOnly
```

`-AuditOnly` checks Python availability, PowerShell 7 module discovery and Graph version alignment, and separately checks Graph Authentication and ImportExcel in the Windows PowerShell 5.1 module path. It does so without installing, importing, changing repository trust or contacting a tenant. It exits nonzero for missing Python or a missing or mismatched prerequisite. It does not qualify imports, Python packages or another account. `-WhatIf` previews installation; it is not an availability audit. When Python is absent, the installer uses winget to install Python 3.14 in the selected scope; if winget is unavailable, install Python 3.10+ separately. `-Scope AllUsers` checks the machine PATH and installs missing Windows PowerShell 5.1 modules in the shared module location. The installer validates their imports under 5.1. Restart existing service processes to pick up PATH changes, and verify the scheduled service identity separately. The default `CurrentUser` scope belongs to the installer account. `-TrustRepository` changes repository trust, and `-Force` and `-AllowClobber` default to true.

The baseline includes Graph, `ImportExcel`, Teams, Power BI Profile, Exchange Online, PnP and MSAL.PS. Organization cmdlets use `Microsoft.Graph.Identity.DirectoryManagement`, not a separate `Microsoft.Graph.Organization` module. The audit checks the highest visible version of each listed Graph submodule against Authentication; other installed submodules and actual imports are checked separately by server acceptance.

## First installation and workload matrix

Deploy the complete reviewed SmartM365 tree together. Configure the neutral global and tenant templates privately; do not copy customer values into templates. Review [configuration precedence](../SmartInventory/CONFIGURATION.md), output paths, private configuration root, account classification and organization-specific workbooks before enabling dependent work. Readers can create or enrich private files; a configuration read is not necessarily read-only. Protect credentials, certificates and logs.

| Workload / host | Additional requirements to qualify |
| --- | --- |
| Cloud inventory and orchestrator | PowerShell 7, Graph/EXO modules and approved application roles; service identity access to the certificate private key, code, tenant paths and shared runtime. |
| AD inventory | RSAT ActiveDirectory, domain connectivity and read rights; ImportExcel and configured classification/workbooks where used. |
| Exchange or SharePoint on-premises | Windows PowerShell 5.1, applicable local snap-ins, server role and read permissions; verify ImportExcel in the Windows PowerShell module path where required. |
| WorkplaceEvidence Prepare | Python 3.10+, compatible separately installed SmartWorkplaceIntelligence runtime and qualified source inputs. |
| SmartFinOps jobs | Matching separately installed SmartFinOps product and its own prerequisites. |
| Windows 11 readiness | PowerShell 7 inventory host, Windows PowerShell 5.1 on targets, explicit target list, authorized account and configured WinRM for remote targets. See [collector README](../SmartInventory/EndpointInventory/Windows11Readiness/README.md). |

The shipped orchestrator catalogue can contain enabled jobs whose external products or local configuration are not included. Before installing or starting a task, review every enabled job, schedule, dependency, server capability and mail/upload setting; disable unsupported jobs in the appropriate runtime configuration, or install and qualify their dependencies. Do not assume a template has been applied to an existing shared configuration. Review code-signing trust and execution policy, and use the documented server acceptance checks under the intended account; these checks can contact services and must be authorized separately. See the [orchestrator guide](../SmartInventory/Orchestrator/README.md).

The app bootstrap is interactive and changes tenant resources, consent and local configuration by default. Review [the permission matrix and App RBAC migration procedure](SmartM365-AppRegistration-Permissions.md), target tenant, certificate, Teams/SharePoint and scoped mail policy before use. Supply `-MailOutOfScopeTestAddress` with an existing mailbox outside the sender group. `-DisableGrantAdminConsent` does not disable the separate Exchange App RBAC assignment. `-WhatIf` still initializes local context and connects for discovery; it is not an offline audit. Updating an app merges existing permissions: excluding a role from the requested baseline does not revoke an existing grant. Verify effective permissions, site grant, sender restriction and service-account certificate access independently.
