# Source Matrix — BETA 3.8.0-beta.1

All sources come directly from SmartInventory. `source-selection.json` is the authoritative allow-list. The model contains one visible table per selected CSV, named exactly like the CSV without `.csv`, and retains every source column.

| History mode | Tables | Grain |
| --- | --- | --- |
| Native daily series | `AD_Users_DailyStats`, `AD_Computers_DailyStats`, `Exchange_OnPrem_Mailboxes_DailyStats` | Date and domain |
| Deduplicated weekly snapshots | `M365_Users_Activity`, `M365_Licenses_Tenant`, `Intune_Devices_Inventory`, `Intune_Devices_Compliance`, `Exchange_EXO_Mailboxes_AllDomains`, `M365_Teams_Teams`, `M365_Teams_UserActivity`, `M365_SPO_Sites`, `Intune_EndpointAnalytics_DevicePerformance`, `Exchange_OnPrem_Servers_Inventory_Summary` | One retained snapshot per prior week plus current DATA-LAST |
| Current state | Remaining selected tables | Latest DATA-LAST export |

Every table receives:

- `__SnapshotDate`
- `__SnapshotDateTime`
- `__SnapshotPeriod`
- `__IsCurrent`
- `__SourceFile`
- `__SourceFolder`

Only `__SnapshotDate` on weekly-history tables is visible. The other metadata columns are hidden to keep the field list usable; no table itself is hidden.

`M365_Licenses_UserServicePlanStates` reads the compact physical columns
`TenantKey`, `UserId`, `SkuId`, `PlanId`, and `StateCode`. Power Query derives
the visible compatibility fields `IsEnabled`, `PlanStatus`, and
`IsServicePlanActive`; the last field is true only when `StateCode` is `A`.

`Intune_DiscoveredApps_AppDeviceRelations` keeps one current row per tenant,
application, and managed device. A hidden `__TenantAppKey` links it to the
unique current `Intune_DiscoveredApps_Summary` application grain. The relation
is deliberately not connected directly to the historical device inventory,
whose device key repeats across weekly snapshots.

## Quality rules

- No row is removed or deduplicated inside a current source table.
- Weekly duplicate runs are collapsed to the latest file for each prior week.
- Historical snapshots below 50% of the current row count are excluded as incomplete.
- Current KPIs explicitly filter `__IsCurrent = TRUE()` on historical tables.
- No direct fact-to-fact relationship is created without a validated shared grain.
- Empty but schema-valid source files are accepted.
- Operational detail remains traceable to the original CSV table and columns.


The runtime schema is pinned in source-schema.json. Read README.md for strict tenant/header/error behavior and KNOWN-LIMITATIONS.md for missing-data and historical completeness boundaries. Derived views now reject ambiguous mappings instead of selecting a record.
