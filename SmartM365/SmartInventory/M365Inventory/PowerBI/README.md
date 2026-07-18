# Power BI and Microsoft Fabric Activity Inventory

`SmartM365-PowerBIFabricActivity-Inventory.ps1` is a read-only SmartInventory collector for tenant-wide Power BI and Microsoft Fabric activity events.

It calls the official Power BI admin endpoint:

```text
GET https://api.powerbi.com/v1.0/myorg/admin/activityevents
```

The collector does not calculate license costs or savings. SmartFinOps can later left-join `M365_PowerBI_Fabric_UserActivity.csv` with `M365_Licenses_Users.csv`.

## Safety and API limits

- Every API request covers one UTC calendar day only.
- The default range is the previous 28 complete UTC days.
- `-FromDate` and `-ToDate` are interpreted as UTC calendar dates and must remain within the 28-day API retention window.
- The current partial UTC day is intentionally rejected.
- Both `continuationUri` and `continuationToken` pagination are supported.
- A sliding one-hour request ledger prevents more than 200 requests in any hour.
- HTTP 429 honors `Retry-After`; transient HTTP 408, 409, and 5xx responses use bounded retries.
- `-MaxItems` stops after a bounded number of events and relies on the shared SmartM365 `MAXITEMS-<n>` suffix, so canonical `DATA-LAST` files and weekly history are not replaced.
- `-ValidateOnly` validates date windows, pagination contracts, the hourly limiter, identity classification, aggregation, and both empty CSV schemas without authenticating or calling the API.

## Authentication

The access token must target the global-cloud Power BI resource:

```text
https://analysis.windows.net/powerbi/api
```

The script validates the token audience and rejects a Microsoft Graph token.

### Interactive Fabric Administrator

Use `-InteractiveAuth` to sign in through the official `MicrosoftPowerBIMgmt.Profile` module. The signed-in account must be a Fabric Administrator. The delegated token must contain `Tenant.Read.All` or `Tenant.ReadWrite.All`.

```powershell
.\SmartM365-PowerBIFabricActivity-Inventory.ps1 -Tenant test -LookbackDays 7 -InteractiveAuth
```

### Service principal with certificate

App-only mode is the default unattended mode. It uses `AppId`, `TenantId`, and `Thumbprint` inherited from the selected Git-ignored SmartM365 tenant profile. Client secrets are not supported.

Fabric configuration is required:

1. Use an Entra application with no admin-consent-required Power BI permissions assigned in the Azure portal.
2. Create or select an Entra security group and add the service principal to it.
3. In the Fabric admin portal, open **Tenant settings** > **Admin API settings**.
4. Enable **Service principals can access read-only admin APIs** for that security group.

This admin API setting is distinct from the general developer setting that allows service principals to use Fabric APIs. The collector does not add `Tenant.Read.All` as an application permission; that scope is relevant only to delegated administrator tokens.

```powershell
.\SmartM365-PowerBIFabricActivity-Inventory.ps1 -Tenant test -LookbackDays 28
```

Microsoft documentation:

- [Get Activity Events REST API](https://learn.microsoft.com/en-us/rest/api/power-bi/admin/get-activity-events)
- [Enable service principal authentication for admin APIs](https://learn.microsoft.com/en-us/fabric/admin/enable-service-principal-admin-apis)
- [Fabric admin API tenant settings](https://learn.microsoft.com/en-us/fabric/admin/service-admin-portal-admin-api-settings)
- [Access the Power BI activity log](https://learn.microsoft.com/en-us/power-bi/guidance/admin-activity-log)

## Parameters

```powershell
.\SmartM365-PowerBIFabricActivity-Inventory.ps1 -Tenant test -ValidateOnly
.\SmartM365-PowerBIFabricActivity-Inventory.ps1 -Tenant test -LookbackDays 28
.\SmartM365-PowerBIFabricActivity-Inventory.ps1 -Tenant test -FromDate 2026-07-01 -ToDate 2026-07-07 -InteractiveAuth
.\SmartM365-PowerBIFabricActivity-Inventory.ps1 -Tenant test -LookbackDays 2 -MaxItems 100
```

Production launcher:

```powershell
..\..\Launchers\Cloud\Start-SmartM365-PowerBIFabricActivity-Inventory.cmd
```

The launcher defaults to `-Tenant prod -LookbackDays 28`, locates PowerShell 7 from Program Files or `PATH`, supports UNC execution, and forwards additional arguments to the collector.

Main parameters:

- `-Tenant test`
- `-LookbackDays 28`
- `-FromDate`
- `-ToDate`
- `-InteractiveAuth`
- `-ValidateOnly`
- `-MaxItems`

## Outputs

The canonical tenant-scoped files are published to `DATA-LAST`:

- `M365_PowerBI_Fabric_ActivityEvents.csv`
- `M365_PowerBI_Fabric_UserActivity.csv`

`TenantKey` is injected as the first column by `SmartM365.Core`. Timestamped copies are stored under `DATA-ALL\M365\PowerBI\Activity`, empty datasets retain their explicit schema, weekly history is enabled by default, and SharePoint upload follows the shared tenant configuration.

The detailed file excludes `ClientIP`, `UserAgent`, and unrelated audit properties. It contains:

```text
CreationTime, Activity, Operation, Workload, UserId, UserPrincipalName,
UserType, PrincipalType, IsSuccess, WorkspaceId, WorkspaceName, ItemName,
ReportName, DatasetName, DataflowName, CapacityId, CapacityName
```

`PrincipalType` distinguishes regular users, guests, applications, and service principals when `UserType` or other event evidence is available. Other official audit types, such as system, policy, partner technician, and agent events, remain separately classified instead of being forced into a user category.

The per-principal file aggregates by `UserId`, `UserPrincipalName`, and `PrincipalType`:

```text
LastActivityDate, ActiveDays, TotalActivityCount, ViewReportCount,
ViewDashboardCount, CreateOrPublishCount, RefreshCount, ExportCount,
DistinctWorkspaceCount, DistinctReportCount, DistinctDatasetCount,
HasRecentActivity
```

`CreateOrPublishCount` counts activity or operation names beginning with `Create`, `Publish`, `Upload`, `Import`, or `Save`. `RefreshCount` matches refresh activity names, and `ExportCount` matches export or download activity names.

Only principals with at least one returned event appear in the aggregate file, so `HasRecentActivity` is true for exported rows. After a left join from the complete license/user population, a missing activity row means only that no event was returned for the selected period. It is a review signal, never proof that a license, account, application, or service principal is unused.
