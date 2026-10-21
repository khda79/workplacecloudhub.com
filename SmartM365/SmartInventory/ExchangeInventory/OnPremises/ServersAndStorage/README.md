# On-Premises Infrastructure And Readiness

Exchange 2016 on-premises infrastructure and migration readiness inventory for servers, compute, memory, logical disks, physical disks, optional database paths, optional service health, Exchange schema, and key Exchange configuration for Exchange SE / Exchange Online planning.

## Scripts

- `SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1`: inventories Exchange 2016 infrastructure through Exchange Management Shell plus WMI/DCOM only, then renders a full HTML readiness report. It does not use WinRM, PowerShell remoting, `Get-CimInstance`, or `New-CimSession`.
- Version 1.6.1 routes each timestamped/current CSV pair through the shared
  atomic exporter. A failed current promotion is terminating and does not replace
  the preceding valid current file. CSV schemas and readiness rules are unchanged.

## Output

By default, output is written under:

```text
{{DataAllRootPath}}\Exchange\OnPrem\ServersAndStorage\<RunId>
```

Use `-OutputRoot` to override the root folder for a specific run.

Set `ExcludedServers` in the script's local JSON to an array of server names that the collection host must not contact. The committed template uses an empty array. Configured names are skipped for per-server inventory and certificate collection; the HTML report and email list them in the per-server table with status `Excluded` and no measured values. The global summary CSV records both the discovered and collected server counts and the configured exclusions.

## Database summary in the report and email

Version 1.6.11 adds a database summary from the current run's existing
`Exchange_OnPrem_MigrationReadiness_Config.csv` rows (`MailboxDatabase`, collected
with `Get-MailboxDatabase -Status`). It shows total observed database names,
mounted, not mounted and unknown status counts, followed by every database name,
reported server, mounted status, database size and internal reusable space.
Recovery databases are included. No additional Exchange query is introduced.

- Missing fields display `N/A`; missing, unrecognized, failed or ambiguous mounted
  evidence displays `Unknown`. Duplicate setting rows are ambiguous, not summed.
- A failed database collection explicitly marks coverage incomplete. An absent
  database section produces `N/A` counters, not a claim of zero databases. Records
  without a database name cannot be associated or counted and are disclosed.
- `AvailableNewMailboxSpace` is reusable space **inside the database**, not free
  space on its disk. Values retain the units returned by Exchange.
- The reported `Server` association is not a list of DAG copies. Detailed DAG
  copy states and per-copy server inventory are not collected or inferred.
- Mailbox counts remain `N/A`: the separate mailbox export has a textual
  `Database` value and its own collection timestamps, but no common database GUID
  or shared snapshot with this source. This report does not load or silently join
  that export. Zero observed mailboxes must never be interpreted as an empty base.
- The source run ID and run date remain visible, including when regenerating a
  historical report. Regeneration uses only that run's supplied readiness rows.
  Database discovery remains organization-wide even when individual servers are
  excluded from per-server WMI collection.

Offline validation (synthetic fixtures only, no collector, authentication or mail):

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File SmartM365/Tests/Test-SmartM365ExchangeDatabaseMailOffline.ps1
```
