# Microsoft 365 Inventory

Microsoft 365 and Entra inventory scripts outside the Intune-specific surface.

## Organization

- `Devices/`: Entra device inventory.
- `Domains/`: verified domain inventory.
- `IntuneInventory/`: Intune inventory, Windows Update reporting, Autopilot, RBAC, applications, and remediation export utilities.
- `Licensing/`: license and service plan inventory.
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
| `Office365ActivationUserDetail` | `M365_Apps_Activations.csv` | Microsoft 365 Apps / Office activations by user and platform; no period parameter. |
| `TeamsUserActivityUserDetail` | `M365_Teams_UserActivity.csv` | Detailed Teams user activity. |
| `EmailActivityUserDetail` | `M365_Email_Activity.csv` | Exchange email send/read/receive activity. |

Examples:

```powershell
.\Usage\SmartM365-M365UserActivity-Inventory.ps1 -Tenant prod -Period D180 -Connect
.\Usage\SmartM365-M365UserActivity-Inventory.ps1 -Tenant prod -Period D180 -Reports All -Connect
.\Usage\SmartM365-M365UserActivity-Inventory.ps1 -Tenant prod -Reports MailboxUsageDetail,OneDriveUsageAccountDetail -Connect
.\Usage\SmartM365-M365UserActivity-Inventory.ps1 -Tenant test -Period D30 -InteractiveAuth
```

Required Microsoft Graph permission: `Reports.Read.All`.

`SmartM365-EXO-Mailboxes-Inventory.ps1` remains useful and is not replaced by `MailboxUsageDetail`: EXO gives mailbox object/stat/archive details from Exchange Online, while Graph Reports gives a period-based usage and quota report suitable for FinOps joins.
