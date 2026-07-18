# Microsoft Intune Endpoint Analytics Inventory

`SmartM365-EndpointAnalytics-Inventory.ps1` is a read-only SmartInventory collector for Microsoft Intune Endpoint Analytics reports available with a standard Intune license.

It does not enable data collection, create policies or baselines, change assignments, run remediations, or calculate financial/productivity estimates.

## Licensing and excluded scope

The collector requires a valid Microsoft Intune license and existing Endpoint Analytics data. It does not require Intune Advanced Analytics, Intune Plan 2, or Intune Suite.

The executable catalogue deliberately excludes:

- every `BR*` Battery Health report;
- every `EAResourcePerf*` Resource Performance report;
- every `EAAnomaly*` report;
- Device Timeline;
- Device Query;
- Advanced score columns such as `ResourcePerfScore` and `OverviewBatteryHealthScore`.

`IsAdvancedAnalytics` is always `False` in DataQuality.

## Microsoft Graph API version and permission

Microsoft documents both v1.0 and beta `deviceManagement/reports/exportJobs` endpoints. The current Microsoft Intune report-name catalogue documents Endpoint Analytics report names only against:

```text
https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs
```

Version 1.0.4 therefore uses beta for all Endpoint Analytics export jobs. Review this choice when Microsoft publishes the report-name catalogue for v1.0.

The preflight validates the application-permission claim without issuing an unsupported `GET` against the `exportJobs` collection. The first real `POST` export job validates Endpoint Analytics API access during collection.

CSV publication enumerates generic row lists through the PowerShell pipeline before passing them to SmartM365.Core. This avoids the PowerShell `Argument types do not match` failure produced by applying an array subexpression directly to `List[object]`.

The operation is functionally read-only: the only POST creates a temporary export-job resource. Microsoft currently documents these least-privileged permission choices for that POST:

- `DeviceManagementManagedDevices.ReadWrite.All`;
- `DeviceManagementConfiguration.ReadWrite.All`;
- `DeviceManagementApps.ReadWrite.All`.

SmartM365 uses `DeviceManagementManagedDevices.ReadWrite.All` as the single required permission because the collected data is device experience data. Microsoft documents `DeviceManagementManagedDevices.Read.All` for reading an existing job, but not for creating one.

No write is sent to an Endpoint Analytics configuration, device, policy, assignment, baseline, or remediation endpoint.

## Standard reports and grains

| Report | SmartM365 output | Grain | Notes |
| --- | --- | --- | --- |
| `EADevicePerformanceV2` | DevicePerformance | device/source report | App reliability, crashes, mean time to failure. |
| `EADeviceModelPerformanceV2` | ModelPerformance | model/source report | Model app reliability and mean time to failure. |
| `EADeviceScoresV2` | DevicePerformance | device/source report | Overall, startup, app, and WFA scores. Advanced score columns are not selected. |
| `EAModelScoresV2` | ModelPerformance | model/source report | Model scores without Advanced score columns. |
| `EAStartupPerfDevicePerformanceV2` | StartupDevices | device | Boot/sign-in scores and timing. |
| `EAStartupPerfModelPerformanceV2` | StartupModels | model | Startup timing and average restart/stop-error counts. |
| `EAStartupPerfDeviceProcesses` | StartupProcesses | process | Only with `-IncludeStartupProcesses`. |
| `EAAppPerformance` | AppReliability | application | Reliability, crashes, usage, and mean time to failure. |
| `EAOSVersionsPerformance` | OSReliability | OS version | OS-version reliability. |
| `EAWFADeviceList` | WorkFromAnywhere | device/source report | Device and OS identity. |
| `EAWFAPerDevicePerformance` | WorkFromAnywhere | device/source report | Device WFA, cloud management, and Windows scores. |
| `EAWFAModelPerformance` | WorkFromAnywhere | model/source report | Model WFA, cloud management, and Windows scores. |

`WorkFromAnywhereDeviceList` is used only as a fallback alias if `EAWFADeviceList` is rejected. DataQuality records `AliasUsed`.

The API does not expose a user principal name, app version, or device last-seen date in these schemas. Those fields are not invented or enriched from broader endpoints.

## Raw-to-SmartM365 mapping

| Microsoft raw column | SmartM365 column |
| --- | --- |
| `MemaTimeGenerated`, `ProcessedDateTime`, `InsertedDate` | `ReportRefreshDate` |
| `DeviceManufacturer`, `Manufacturer` | `Manufacturer` |
| `DeviceModel`, `Model` | `Model` |
| `StartupPerformanceScore` | `StartupScore` |
| `LogonScore` | `SignInScore` |
| `CoreLogonTime` | `CoreSignInTime` |
| `BlueScreenCount`, `AverageBlueScreens` | `StopErrorCount` |
| `AverageRestarts` | `RestartCount` |
| `DeviceAppHealthScore`, `ModelAppHealthScore`, `AppHealthScore`, `OSVersionAppHealthScore` | `AppReliabilityScore` |
| `TotalAppCrashes` | `CrashCount` |
| `TotalAppUsageDuration`, `TimePerProcess` | `UsageDuration` |
| `AppFriendlyName`, `AppName`, `ProductName`, `FileDescription`, `ProcessName` | `ApplicationName` |
| `AppPublisher`, `Publisher` | `Publisher` |

Every request has an explicit `select` list; default report columns are never used.

## Canonical CSV files

- `Intune_EndpointAnalytics_DevicePerformance.csv`
- `Intune_EndpointAnalytics_ModelPerformance.csv`
- `Intune_EndpointAnalytics_StartupDevices.csv`
- `Intune_EndpointAnalytics_StartupModels.csv`
- `Intune_EndpointAnalytics_StartupProcesses.csv` when requested
- `Intune_EndpointAnalytics_AppReliability.csv`
- `Intune_EndpointAnalytics_OSReliability.csv`
- `Intune_EndpointAnalytics_WorkFromAnywhere.csv`
- `Intune_EndpointAnalytics_DataQuality.csv`

Historical files use `{{DataAllRootPath}}\Intune\EndpointAnalytics`; current files use `LatestCsvFolderPath`. SmartM365.Core injects `TenantKey` first, creates stable empty-schema CSVs, applies retention, handles weekly history and optional SharePoint upload, and adds MAXITEMS suffixes in test mode.

DataQuality contains:

```text
TenantKey, RunId, ReportName, ApiVersion, Status, RowCount,
ExportJobStatus, IsAdvancedAnalytics, RequiredPermission,
ErrorCode, ErrorMessage, CollectedAtUtc
```

## Examples

```powershell
# Static validation without Graph
pwsh -File .\SmartM365-EndpointAnalytics-Inventory.ps1 -Tenant test -ValidateOnly

# Validate report availability in the test tenant
pwsh -File .\SmartM365-EndpointAnalytics-Inventory.ps1 -Tenant test -ValidateOnly -InteractiveAuth

# App-only collection
pwsh -File .\SmartM365-EndpointAnalytics-Inventory.ps1 -Tenant test -Reports All -Connect

# Startup test with safe MAXITEMS filenames
pwsh -File .\SmartM365-EndpointAnalytics-Inventory.ps1 -Tenant test -Reports Startup -IncludeStartupProcesses -MaxItems 10 -InteractiveAuth

# Offline export-job and contract simulations
pwsh -File .\SmartM365-EndpointAnalytics-Inventory.ps1 -SelfTest -MaxItems 10
```

Do not run a production tenant collection without explicit operational approval.
