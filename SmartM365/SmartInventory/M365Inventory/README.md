# Microsoft 365 Inventory

Microsoft 365 and Entra inventory scripts outside the Intune-specific surface.

## Organization

- `Devices/`: Entra device inventory.
- `Domains/`: verified domain inventory.
- `IntuneInventory/`: Intune inventory, Windows Update reporting, Autopilot, RBAC, applications, and remediation export utilities.
- `IntuneInventory/EndpointAnalytics/`: standard Microsoft Intune Endpoint Analytics inventory; Advanced Analytics report families are excluded.
- `Licensing/`: license and service plan inventory.
- `PowerBI/`: tenant-wide Power BI and Microsoft Fabric activity-event inventory.
- `Teams/`: Teams collaboration inventory plus Teams Phone PSTN, Operator Connect, and Direct Routing usage.
- `SharePoint/`: SharePoint Online site, storage, list/library, permission, and external sharing inventory.
- `SyncHealth/`: Azure AD Connect synchronization freshness checks and stale hybrid data alerts.
- `Usage/`: Microsoft 365 user and workload usage reports for FinOps and usage analysis.
- `Users/`: active user inventory.

## Microsoft 365 Usage Reports

`Usage/SmartM365-M365UserActivity-Inventory.ps1` exports Microsoft Graph Reports data and publishes stable CSV files to the tenant `DATA-LAST` folder.

Default behavior is intentionally unchanged: without `-Reports`, the script exports only `Office365ActiveUserDetail` as `M365_Users_Activity.csv`.

Supported report selections:

| Report | Latest CSV | Notes |
| --- | --- | --- |
| `Office365ActiveUserDetail` | `M365_Users_Activity.csv` | Normalized workload activity by user. |
| `MailboxUsageDetail` | `M365_Mailbox_Usage.csv` | Graph usage view: storage, quotas, deleted items, archive flag, last activity. Complements the EXO mailbox inventory. |
| `OneDriveUsageAccountDetail` | `M365_OneDrive_Usage.csv` | OneDrive storage, file counts, active files, owner, last activity. |
| `SharePointSiteUsageDetail` | `M365_SharePoint_SiteUsage.csv` | SharePoint site storage, file/page activity, owner/site metadata from Graph reports. |
| `SharePointActivityUserDetail` | `M365_SharePoint_UserActivity.csv` | SharePoint file viewing/editing, sync, internal/external sharing, and page visits by user. |
| `Office365ActivationUserDetail` | `M365_Apps_Activations.csv` | Microsoft 365 Apps / Office activations by user and platform; no period parameter. |
| `TeamsUserActivityUserDetail` | `M365_Teams_UserActivity.csv` | Detailed Teams user activity. |
| `TeamsDeviceUsageUserDetail` | `M365_Teams_DeviceUsage.csv` | Teams client device and operating system usage by user. |
| `EmailActivityUserDetail` | `M365_Email_Activity.csv` | Exchange email send/read/receive activity. |

Examples:

```powershell
.\Usage\SmartM365-M365UserActivity-Inventory.ps1 -Tenant prod -Period D180 -Connect
.\Usage\SmartM365-M365UserActivity-Inventory.ps1 -Tenant prod -Period D180 -Reports All -Connect
.\Usage\SmartM365-M365UserActivity-Inventory.ps1 -Tenant prod -Reports MailboxUsageDetail,OneDriveUsageAccountDetail -Connect
.\Usage\SmartM365-M365UserActivity-Inventory.ps1 -Tenant test -Period D30 -InteractiveAuth
```

Required Microsoft Graph permission: `Reports.Read.All`.

`SharePointActivityUserDetail` and `TeamsDeviceUsageUserDetail` use the same existing
`Reports.Read.All` permission. They do not require another app registration,
permission, or local configuration value.

For multi-report runs, DATA-ALL and DATA-LAST files are uploaded per report, while
the cumulative WeeklyHistory snapshot is built and uploaded once after all selected
reports complete successfully.

`SmartM365-EXO-Mailboxes-Inventory.ps1` remains useful and is not replaced by `MailboxUsageDetail`: EXO gives mailbox object/stat/archive details from Exchange Online, while Graph Reports gives a period-based usage and quota report suitable for FinOps joins.

## Power BI And Microsoft Fabric Activity Events

`PowerBI/SmartM365-PowerBIFabricActivity-Inventory.ps1` retrieves tenant-wide activity events from the official Power BI admin Activity Events API. Requests are split into one UTC day, limited to the last 28 days, paged with `continuationUri` or `continuationToken`, and kept below 200 requests per hour.

Exports:

| Entity | Latest CSV | Notes |
| --- | --- | --- |
| Detailed events | `M365_PowerBI_Fabric_ActivityEvents.csv` | Useful activity, workload, identity classification, workspace, item, report, semantic model/dataset, dataflow, and capacity fields. Client IP, user agent, and unrelated audit properties are excluded. |
| Per-principal activity | `M365_PowerBI_Fabric_UserActivity.csv` | Aggregates recent activity, active UTC days, view/create/refresh/export counters, and distinct artifact counts by user/application identity. |

Examples:

```powershell
.\PowerBI\SmartM365-PowerBIFabricActivity-Inventory.ps1 -Tenant test -ValidateOnly
.\PowerBI\SmartM365-PowerBIFabricActivity-Inventory.ps1 -Tenant test -LookbackDays 28
.\PowerBI\SmartM365-PowerBIFabricActivity-Inventory.ps1 -Tenant test -FromDate 2026-07-01 -ToDate 2026-07-07 -InteractiveAuth
..\Launchers\Cloud\Start-SmartM365-PowerBIFabricActivity-Inventory.cmd
```

Interactive mode requires a Fabric Administrator and delegated `Tenant.Read.All` or `Tenant.ReadWrite.All`. App-only mode uses the configured certificate and requires the Fabric tenant setting **Service principals can access read-only admin APIs** for a security group containing the service principal. The Entra application must not have admin-consent-required Power BI permissions assigned.

The aggregate contains only principals for which an event was returned. After SmartFinOps left-joins the complete license/user population, a missing activity row is a review signal and never automatic proof of non-use.

## Microsoft 365 Copilot Usage

`Usage/SmartM365-CopilotUsage-Inventory.ps1` is a read-only collector for the official Microsoft Graph v1.0 endpoint `GET /copilot/reports/getMicrosoft365CopilotUsageUserDetail`.

It publishes `M365_Copilot_UserUsage.csv` to the tenant `DATA-LAST` folder and keeps timestamped DATA-ALL plus weekly history copies. Report version `v2` is the default because it adds active usage days, prompt counters for all apps and Copilot Chat work/web, and the Microsoft 365 app, Edge, agent, and Copilot Chat work/web last-activity dates. The script falls back to `v1` only when `v2` is unavailable or when `D30` is explicitly requested, because Microsoft documents `D30` for `v1` and `D28` for `v2`. For compatibility with tenants where the versioned route is not yet recognized, the v1 fallback omits the optional `version` function parameter and therefore uses the documented v1 default.

The API returns only users who have a Microsoft 365 Copilot license. The collector never requests or exports prompt text. Version `v1` does not expose prompt counters or active usage days, so the stable v2-oriented CSV columns remain empty after a v1 fallback.

Examples:

```powershell
.\Usage\SmartM365-CopilotUsage-Inventory.ps1 -Tenant test -ValidateOnly
.\Usage\SmartM365-CopilotUsage-Inventory.ps1 -Tenant prod -Period D180 -ReportVersion v2 -Connect
.\Usage\SmartM365-CopilotUsage-Inventory.ps1 -Tenant test -Period D7 -ReportVersion v2 -InteractiveAuth
.\Usage\SmartM365-CopilotUsage-Inventory.ps1 -Tenant test -Period D180 -ReportVersion v2 -Connect -MaxItems 10
```

`-MaxItems` produces suffixed `MAXITEMS` CSV files and does not replace the canonical Power BI file or weekly history. `-ValidateOnly` validates modules, parameters, paths, URI construction, and the empty CSV schema without calling Microsoft Graph.

Required Microsoft Graph permission: `Reports.Read.All`.

Official documentation:

- [Microsoft 365 Copilot usage user detail API](https://learn.microsoft.com/microsoft-365/copilot/extensibility/api/admin-settings/reports/copilotreportroot-getmicrosoft365copilotusageuserdetail)
- [Microsoft 365 Copilot usage report](https://learn.microsoft.com/microsoft-365/admin/activity-reports/microsoft-365-copilot-usage)

## SharePoint Online Inventory

`SharePoint/SmartM365-SPO-Inventory.ps1` uses Microsoft Graph for the baseline site/list inventory and collects tenant storage capacity through `Get-PnPTenant` by default. The SharePoint administration URL is derived automatically from the configured SharePoint hostname, a collected site URL, or the tenant `onmicrosoft.com` name. `SharePointAdminUrl` remains an optional override for atypical tenants. Use `-SkipPnPTenantCapacity` for the least-privilege Graph-only mode. Fields not exposed by Graph are exported as `NotAvailableGraphOnly` instead of failing the run.

Main exports:

| Entity | Latest CSV | Notes |
| --- | --- | --- |
| Sites | `M365_SPO_Sites.csv` | Site URL, title, template, storage quota/usage, activity, owner signal, inactive/orphaned flags, and Graph-only availability markers. |
| Lists | `M365_SPO_Lists.csv` | Lists and libraries per site when Graph can resolve the site. Versioning and size fields are marked unavailable in Graph-only mode. |
| Permissions | `M365_SPO_Permissions.csv` | Owner rows from Graph usage data. Site collection admin enumeration is not required in default mode. |
| External sharing | `M365_SPO_ExternalSharing.csv` | Stable schema with Graph-only availability markers. Tenant-wide anonymous/external sharing link discovery is not available in least-privilege Graph-only mode. |
| Tenant capacity | `M365_SPO_Tenant.csv` | Always exported. Used storage comes from the site inventory. Licensed capacity is collected by default; `SharePointAdminUrl` is derived automatically and remains an optional override. |

Examples:

```powershell
.\SharePoint\SmartM365-SPO-Inventory.ps1 -Tenant prod
.\SharePoint\SmartM365-SPO-Inventory.ps1 -Tenant prod -IncludeOneDrive -InactiveDays 180
.\SharePoint\SmartM365-SPO-Inventory.ps1 -Tenant prod -MaxSites 10 -DryRun
.\SharePoint\SmartM365-SPO-Inventory.ps1 -Tenant prod -AppendHistory -AlwaysSend
.\SharePoint\SmartM365-SPO-Inventory.ps1 -Tenant prod -SkipPnPTenantCapacity
```

Required PowerShell modules for the default run: `Microsoft.Graph.Authentication` and `PnP.PowerShell`.

Required Graph app-only permissions for the baseline inventory: `Reports.Read.All` and `Sites.Read.All`. `Directory.Read.All` is useful for owner enrichment. The default capacity step additionally requires access to the SharePoint tenant administration service.

Tenant capacity collection is enabled by default and requires access to the SharePoint tenant administration site. Manual `SharePointAdminUrl` configuration is normally unnecessary. If the endpoint cannot be derived or the app lacks access, the tenant CSV records a warning with blank capacity fields; the baseline inventory continues. Use `-SkipPnPTenantCapacity` to disable the attempt explicitly.