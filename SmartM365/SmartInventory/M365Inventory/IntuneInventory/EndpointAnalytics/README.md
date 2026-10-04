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

Version 1.0.6 continues to use beta for all Endpoint Analytics export jobs and adds throttling-aware outer retries. Review this choice when Microsoft publishes the report-name catalogue for v1.0.

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

Version 1.0.12 also requests `DeviceScopeIds`, `HealthStatus` and
`PartnerFeaturesBitmask` for `EADeviceScoresV2`. These documented metadata columns
are preserved only in private rejected-row diagnostics; canonical CSV schemas
are unchanged. They are not used to select a preferred row or relax the grain
guard. Advanced score columns remain excluded.

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
ErrorCode, ErrorMessage, CollectedAtUtc,
RawRowCount, ExcludedRowCount, ExcludedDeviceCount
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

## Device-grain consistency

For each device report, a complete export must contain at most one row per
`DeviceId` (case-insensitive, trimmed). Blank device identities and repeated rows,
including identical duplicates or conflicting scores, are rejected. Different
reports for the same device remain separate; model/application/process reports
do not inherit the device identity requirement.

`ReportConsistencyAttempts` defaults to 3 (range 1–3). A grain failure requests a
new export job and downloads the whole export again; the default delays are 15
then 30 seconds (`ReportConsistencyRetryDelaySeconds`, capped at 300 seconds).
Rows are never merged, averaged or chosen by score/date to repair an
ambiguous export. Permission failures and existing HTTP retries keep their
separate behavior. Alias fallback cannot hide a grain failure. Availability-only
validation does not download or qualify CSV contents.

Version 1.0.13 has one explicit exception after all fresh consistency attempts:
for `EADeviceScoresV2` only, every row whose trimmed, case-insensitive `DeviceId`
is repeated is excluded. This includes identical duplicates; neither a preferred
score nor the first/last row is kept. Invalid device identities remain blocking.
The final fresh export must have a saved private rejected-row diagnostic before
this exception can be used. Valid rows are preserved, the cleaned report is
checked again, and raw rows reconcile exactly with published plus excluded rows.

`DataQuality` declares `CollectedWithExclusions`, `DuplicateScoreRowsExcluded`,
and the raw, published, excluded-row and affected-device counts. The run ends as
`CompletedWithWarnings`. After valid CSV publication, one Core-branded warning
mail is sent through `ErrorMailTo` with affected device identities and excluded
score/health values, plus log/transcript paths. Core applies the existing mail
transport and maintenance recipient restriction; no raw diagnostic is attached.
Mail failure is not hidden: the collector fails while already published CSVs
remain preserved, and its completion receipt is not admitted as successful.

The CMDB completion proof includes count-only exclusion qualifications. This
means the requested scope was fully acquired, not that every source row was
usable. `MaxItems` and missing required reports still disqualify full scope.
The device remains in the independent AD/Entra/Intune inventory. CMDB and
Intelligence leave its excluded overall score unavailable: no zero, old score,
or stale export is substituted. Independent app/startup evidence remains usable.

For any other device report, or invalid score identities, persistent ambiguity
still fails before canonical business CSV
publication; the prior files remain but the failed source receipt must prevent
their admission as a successful fresh collection. All device-grain output groups
are checked again before the first publication call, including the mixed
device/model Work From Anywhere output. Console/log messages contain report names
and counts, not device identities or score values.

## Private trace evidence

Version 1.0.11 routes collector messages through the prefixed Core logger and
starts a transcript after runtime initialization. The Core completion banner is
captured before the transcript is closed. The closed transcript is uploaded
through the existing SharePoint helper when configured; validation does not
enable external actions.

Version 1.0.12 requires Core 1.0.69 and defers the active transcript upload in
shared completion. The collector closes and uploads it after the final banner;
other logs, source receipts and mail artifacts retain their existing upload
behavior. Terminal export errors include the failed report name and its cause,
rather than only an aggregate failure count.

For each rejected device-grain export, only the duplicate/invalid raw rows are
saved as a `.json.txt` diagnostic in `Diagnostics` under the script's `LOG-ALL`
folder. It includes the run, tenant, export-job identity, attempt, raw row numbers,
identical/different-row classification and differing column names. Unique rows
and signed download URLs are not retained. These files contain private device
evidence and must never enter public Git or `DATA-LAST`. They are trace artifacts,
not successful collector outputs, and do not authorize publication of raw
ambiguous rows. Only the explicitly qualified cleaned score report can publish.
They use `RetentionMaxLogs` with the shared cleanup helper (current-run evidence
is preserved) and the configured SharePoint upload behavior.

Trace regression tests use temporary synthetic fixtures and simulated external
actions:

```powershell
pwsh -NoProfile -File .\SmartM365\Tests\Test-SmartM365EndpointAnalyticsTraceOffline.ps1
```

Offline regression tests (mocked export jobs, downloads, waits and publishers):

```powershell
pwsh -NoProfile -File .\SmartM365\Tests\Test-SmartM365EndpointAnalyticsConsistencyOffline.ps1
```
