# On-Premises Infrastructure And Readiness

Exchange 2016 on-premises infrastructure and migration readiness inventory for servers, compute, memory, logical disks, physical disks, optional database paths, optional service health, Exchange schema, and key Exchange configuration for Exchange SE / Exchange Online planning.

## Scripts

- `SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1`: inventories Exchange 2016 infrastructure through Exchange Management Shell plus WMI/DCOM only, then renders a full HTML readiness report. It does not use WinRM, PowerShell remoting or `Get-CimInstance`; `New-CimSession` is used over DCOM only, to start the Health Checker scheduled task.
- `Install-SmartM365-Exchange-HealthCheckerScheduledTask.ps1`: installs, updates or removes the Health Checker scheduled task on the other Exchange servers (launcher: `Launchers/OnPremises/Start-SmartM365-Exchange-HealthChecker-ScheduledTask-Installer.cmd`).
- `SmartM365-Exchange-OnPrem-HealthChecker-LocalRunner.ps1`: deployed to each of those servers; runs Health Checker for the local server only.
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

## Exchange Health Checker (`-IncludeHealthChecker`)

Microsoft's [CSS-Exchange Health Checker](https://microsoft.github.io/CSS-Exchange/Diagnostics/HealthChecker/)
runs on every collected server. The script is vendored unchanged under
`ThirdParty/CSS-Exchange/HealthChecker` (version, release, SHA256 and license in its `PROVENANCE.md`).
Health Checker always runs on the server it checks and is never given `-Server`: PowerShell remoting
(WinRM) is not needed.

- Before running, the collector copies the vendored file to a local work folder and checks the copy:
  expected SHA256 and a valid Authenticode signature from Microsoft Corporation. On any mismatch
  Health Checker is not run (`RunStatus` `IntegrityFailed`, warning) and the rest of the inventory
  continues. It always runs with `-SkipVersionCheck`, so it never downloads or updates itself.
- **Collector server** (version 1.8.0): Health Checker runs in a local Windows PowerShell 5.1 process.
- **Other servers**: each one runs Health Checker through the scheduled task
  `\WCH\SmartM365 Exchange Health Checker`, which starts the local runner. The collector:
  1. checks the files under `C:\Program Files\SmartM365\HealthChecker\Bin` through the `C$` share and
     copies them again when their SHA256 differs;
  2. starts all the tasks at once over CIM/DCOM, then runs the collector server itself;
  3. waits for each server's `Output\status.json` and copies only `Output\result.json`.
- The runner checks SHA256 and the Microsoft signature again before every run, takes a lock, removes
  the previous run output, converts the Health Checker XML on the server into `result.json` (result
  lines as strings and numbers only) and writes `status.json` (state, UTC times, exit code, result
  file name and versions, no host name or secret). The task also runs daily at 01:00.
- The collector never deserializes XML produced by another server: it reads `result.json` with a
  20 MB limit and a strict schema. Only the collector server's own XML is read as XML.
- Fallback: when the on-demand start fails (server unreachable, task missing, task logon failure),
  the collector uses the last daily result if it is at most `HealthCheckerMaxResultAgeHours` old
  (default 26; `ResultSource` `Scheduled`). An older result gives `StaleResult`.
- Time limits: `HealthCheckerTimeoutMinutesPerServer` (default 30) per server and
  `HealthCheckerTotalBudgetMinutes` (default 1320, below the orchestrator job timeout) for the run.
  A server over its limit is `Timeout` locally or `TaskTimeout` remotely; a server left when the
  budget is used up is `NotRun`.
- Health Checker needs an elevated session. Without it, no server is run and every row is
  `NotElevated`.
- Raw output (XML, text and debug log) contains the full server configuration. It stays on a local
  disk: the collector work folder (`HealthCheckerWorkFolderPath`, default
  `%ProgramFiles%\SmartM365\HealthChecker\Work`, shares refused) and `Output` on each server. The
  verified copy of Health Checker runs elevated from the work folder, so the collector makes
  Administrators its owner with an administrators-only ACL, refuses it (`UnsafeWorkFolder`) when a
  parent is owned or writable by anyone else than SYSTEM, Administrators or TrustedInstaller, and
  checks the copy again immediately before starting it. Only the last
  run is kept on both sides; XML results are deleted once read. It is never published to DATA-ALL,
  DATA-LAST or SharePoint. `HealthCheckerKeepFailedRawOutputDays` is no longer used.

Run statuses in `Exchange_OnPrem_HealthChecker_Summary.csv`:

| `RunStatus` | Meaning |
| --- | --- |
| `Completed` | Result read. `ResultSource`: `Local`, `OnDemand` or `Scheduled` (fallback, see `RunMessage`). |
| `RemoteConnectionFailed` | The `C$` share or CIM/DCOM of the server is not reachable, or the result could not be copied. |
| `TaskTriggerFailed` | Task not installed, files not deployed, start refused, or task logon failure. |
| `TaskTimeout` | No result from the task within the per-server limit, or Health Checker stopped on the server. |
| `StaleResult` | The on-demand start failed and the last daily result is older than the limit. |
| `NoOutput` | Health Checker ended without XML (including Exchange Management Shell load failure). |
| `UnsafeWorkFolder` | The collector work folder or one of its parents is not restricted to administrators; nothing was run. |
| `IntegrityFailed`, `NotElevated`, `Timeout`, `LaunchFailed`, `ParseFailed`, `NotRun` | As before. |

Messages never contain a domain name, an address or a secret.

### Scheduled task installation

Run once from the collector server, elevated, in Windows PowerShell 5.1, with an account that is
local administrator of the target servers. Started from PowerShell 7, the installer restarts itself
in Windows PowerShell 5.1: there, the ScheduledTasks module would run through Windows compatibility
and could not use the CIM session. The launcher always uses Windows PowerShell 5.1. Network from the collector server: RPC 135 and dynamic RPC ports (DCOM), SMB 445.

```powershell
.\Install-SmartM365-Exchange-HealthCheckerScheduledTask.ps1 -ComputerName 'EXSRV01' -ServiceAccount 'CONTOSO\svc-exchange-hc' -StartNow
.\Install-SmartM365-Exchange-HealthCheckerScheduledTask.ps1 -ComputerName 'EXSRV02','EXSRV03' -ServiceAccount 'CONTOSO\svc-exchange-hc'
.\Install-SmartM365-Exchange-HealthCheckerScheduledTask.ps1 -ComputerName 'EXSRV01' -Uninstall
```

- Validate on one server first with `-StartNow`: the installer waits for the run and prints its state
  (for example `Completed`, or `ExchangeShellFailed` when the account cannot load the Exchange
  Management Shell).
- Files and permissions: under `C:\Program Files\SmartM365\HealthChecker`, where standard users
  cannot create files (the task runs elevated). `Bin` (Health Checker and runner; service account
  read) and `Output` (service account modify), owned by Administrators, SYSTEM and Administrators
  full control, inheritance removed. A server whose `C:\Program Files\SmartM365` is owned or
  writable by anyone else than SYSTEM, Administrators or TrustedInstaller is refused.
- Task: Windows PowerShell 5.1, daily at `-DailyAt` (default `01:00`), highest privileges, a new start
  is ignored while a run is in progress, stopped after one hour.
- Running the installer again is safe; `-WhatIf` shows the changes without making them.
- `-ComputerName` also accepts one value separated by commas or semicolons, as passed by
  `powershell -File` (for example `-ComputerName EXSRV01,EXSRV02`). Any failed server, validation error
  or failed first run makes the installation end as `FAILED` with its error count.

Minimum rights:

- **Task service account** (a dedicated domain account; SYSTEM is refused): local administrator of
  each server, member of the Exchange role group **View-Only Organization Management**, and
  **Log on as a batch job** (Task Scheduler grants it at registration unless a Group Policy
  overrides it).
- **Collector run account** (the orchestrator job account): local administrator of each server, for
  the `C$` share and CIM/DCOM.

Service account password: it is typed once in `Get-Credential` during installation and stored only
by Task Scheduler; it is never written to a file, a log or the output. When it expires or is
changed, run the installer again with the same servers and account: the task is updated in place.
Until then the collector reports `TaskTriggerFailed` with the task logon failure, and uses the last
daily result while it is recent enough.

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
powershell.exe -NoProfile -ExecutionPolicy Bypass -File SmartM365/Tests/Test-SmartM365ExchangeHealthCheckerOffline.ps1
```
