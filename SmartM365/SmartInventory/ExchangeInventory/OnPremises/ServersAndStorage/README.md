# On-Premises Infrastructure And Readiness

Exchange 2016 on-premises infrastructure and migration readiness inventory for servers, compute, memory, logical disks, physical disks, optional database paths, optional service health, Exchange schema, and key Exchange configuration for Exchange SE / Exchange Online planning.

## Scripts

- `SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1`: inventories Exchange 2016 infrastructure through Exchange Management Shell plus WMI/DCOM only, then renders a full HTML readiness report. Compute/storage inventory and Health Checker task triggering keep their existing WMI/DCOM paths. Certificate reads try the native snap-in first and retry failed targets through a constrained Exchange `/PowerShell/` session; `Get-CimInstance` and remote scriptblocks are not introduced. `New-CimSession` is still DCOM-only for Health Checker task triggering.
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

## Certificate connection mode and budgets (1.8.8)

The per-script local JSON key `ExchangeConnectionMode` defaults to `LocalFirst`. Existing configurations receive the missing key from the template without replacing local values. The key affects certificate reads only: Exchange snap-in loading, all other Exchange cmdlets, WMI/DCOM, Health Checker, its runner and installer are unchanged.

- `LocalFirst` (default since 1.8.7): each included server is read through the native snap-in cmdlet in an isolated local child process. Native successes retain their existing Services, Status, IsSelfSigned and domain rendering and are never retried. Only a failed or timed-out native read triggers a WARN with the local error and a remote retry for that server. A `CollectionPath=RemoteRetry` row and source log identify the retry without adding a CSV column. One remote session is opened lazily in a dedicated local worker and reused; a run with only native successes opens no session. Native reads continue for later servers even after a session startup/import failure or loss. Failed retries emit `Collection=Failed / CollectionStatus=ERROR` with both local and remote messages, then continue the infrastructure run. An unavailable session is not reopened. Cancellation propagates and cleans up owned workers; graceful worker shutdown closes its owned session. Remote retry metadata keeps the Unknown semantics described below.
- `Remote`: the shared lot A helper imports only `Get-ExchangeCertificate` into one owned Microsoft.Exchange/Kerberos session at the local management host FQDN (`http://<host>/PowerShell/`). The original `-Server` target is preserved for every read. The session is opened at the certificate phase, used sequentially for all included targets and closed afterwards, including on failure when the worker can respond. It is not kept idle while Health Checker runs. Collection duration is logged.
- In explicit `Remote` mode, session startup/import failure logs `ExchangeConnectionMode=Remote; certificate state=Failed; no fallback` at ERROR level and emits one `ExchangeCertificate / Collection=Failed / CollectionStatus=ERROR` row for each included target. The infrastructure run continues. If the session or proxy is lost during collection, certificate calls stop and the failed target plus all unread targets receive the same explicit failure diagnostic; already collected certificate rows remain. No native certificate fallback or automatic reconnect occurs. Other readiness phases, CSV/HTML exports, summary, SharePoint uploads and mail continue. A target-specific read failure still permits other targets to be read. No failure is represented as zero certificates.
- `LocalSnapin`: native Get-ExchangeCertificate -Server only, in an isolated local child process, with existing property rendering. Change the key in the script-local **source** configuration copied by the launcher; cache-only edits disappear on the next launch.

Session startup/import failures increment the existing error counter and log a warning, so the existing completion path produces `CompletedWithWarnings` with exit code **3**, after outputs and mail. The banner and execution summary retain the nonzero error count. The source-receipt helper does not qualify a run whose error count is nonzero: its run receipt is Failed/unqualified even though the inventory files were produced. This receipt contract is unchanged.

Since 1.8.10, a failed certificate target is logged at ERROR and increments the general error counter, matching its `CollectionStatus=ERROR` row. The run continues through other targets, exports, HTML and mail, ending `CompletedWithWarnings` (exit 3). A nonzero error count leaves the run receipt Failed/unqualified and does not promote a new successful current proof. This prevents incomplete collection from qualifying as a fresh complete source. Expected native failure followed by a successful remote retry remains WARN-only; missing remote metadata remains Unknown/WARNING. A missing config key uses LocalFirst in both code and template; an explicit Remote or LocalSnapin value is preserved.

CertificateTimeoutSecondsPerServer defaults to **120**, an integer from 1 to 3600. The budget resets for each included server and includes child startup, snap-in loading, session opening/import, reading and worker projection. In LocalFirst the native attempt receives at most half the total (60 seconds by default); a failure or timeout permits a remote retry using only the remaining total budget. Remote and LocalSnapin use the entire budget for their single path. There is no second full budget for a retry and no unlimited/zero setting. Missing configuration keys are synchronized from the template; explicit local values remain intact.

SmartM365-Exchange-OnPrem-CertificateWorker.ps1 lives beside the collector and is copied by the launcher's existing script-folder mirror. Native values are converted to simple data while Exchange types are still alive, preserving Services/Status/IsSelfSigned/arrays/dates. Worker messages are JSON data with request identity/schema checks, never executable input. Missing worker, malformed/truncated output and timeouts produce explicit Failed/ERROR diagnostics, never a known zero or partial success.

The controller terminates only its own blocked child, not any Exchange service. Cleanup waits are bounded to at most one second for graceful exit plus one second after termination per worker, outside the collection budget. A timeout of the remote worker disables all further remote retries for this run to preserve the one-session contract; subsequent native reads still run. The finally always performs owned-worker cleanup. An unresponsive worker may need forced termination, so immediate cancellation of the already submitted server operation or immediate server-side session removal is **not guaranteed**. No reconnect, quota increase or server change is introduced. Normal shutdown invokes the unchanged lot A session helper's cleanup in the child's finally.

Per-server logs include budget, duration, path, failure stage and local/remote errors. Failed rows remain in readiness exports; successful targets and all later infrastructure export/report/mail phases remain available. Existing warning/error accounting and completion behavior remain as documented above. Report-only regeneration starts no worker.

The endpoint is NoLanguage: it receives cmdlet calls and parameters only. Sorting, property projection and diagnostics run on the client. Native read permissions remain necessary for the unchanged collectors; the account must also have effective RBAC access to `Get-ExchangeCertificate -Server`. No rights are changed by this mode.

Remoting may return only X509 certificate properties. Subject, thumbprint and available dates are retained; raw date strings are preserved without inferring a timezone or interpreting a culture-dependent display date. `DnsNameList` supplies the existing `CertificateDomains` field when the Exchange property is absent. Available `ServicesStringForm` can supply `Services`. Missing `Services`, `Status` or `IsSelfSigned` are **Unknown**, with `CollectionStatus=WARNING` and a diagnostic, never inferred from subject, issuer or expiration alone. This is a material limit: remote access success does not prove service bindings or Exchange certificate validity. A known false, zero or empty collection remains distinct from a missing property. No CSV columns are changed.

Offline checks, without collecting Exchange data:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File SmartM365/Tests/Test-SmartM365ExchangeCertificatesOffline.ps1
pwsh.exe -NoProfile -File SmartM365/Tests/Test-SmartM365ExchangeCertificatesOffline.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File SmartM365/Tests/Test-SmartM365ExchangeHealthCheckerOffline.ps1
pwsh.exe -NoProfile -File SmartM365/Tests/Test-SmartM365ExchangeHealthCheckerOffline.ps1
```

After operator deployment, launch infrastructure directly under the service account. Check version 1.8.8, per-target 120-second total budget plus bounded cleanup, native metadata on successful targets, at most one lazy certificate session, RemoteRetry only for native failures, the formerly failing target's certificate count, explicit Unknown metadata diagnostics limited to retries, unchanged WMI/DCOM/HealthChecker results, publication and cleanup. Offline tests prove neither server access nor completeness. Report-only regeneration opens no Exchange session.

## Health Checker enabled by default (1.8.9)

The script-local JSON key `IncludeHealthChecker` defaults to `true`, including when absent from an older configuration. Existing template synchronization adds missing keys without replacing local values. Set the JSON boolean to `false` to disable ordinary runs. An explicit switch takes priority: `-IncludeHealthChecker` enables it with config false; `-IncludeHealthChecker:$false` disables it with config true. Text such as `"false"` is rejected. The effective choice and its source are logged before the receipt records scope.

Health Checker requires the existing elevated account and installed remote tasks. Enabling it runs the local child and triggers existing remote tasks; task definitions, runners, rights and launcher are unchanged. HTML-only regeneration reads existing results without starting Health Checker.

`HealthCheckerTimeoutMinutesPerServer` remains independent of activation. Existing local values (including 10 minutes) are preserved; the template default remains 30. Check `Per-server limit: 10 min` in the log. This limits the local child and the collector's wait; remote runners retain their own timeout and may continue after the collector stops waiting. Completing before the deadline does not exercise timeout enforcement.

Direct launcher from PowerShell:

```powershell
& '<source-root>\SmartInventory\Launchers\OnPremises\Start-SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.cmd'
# Explicit opt-out when invoking the cached script from Windows PowerShell 5.1:
& '<cached-script-folder>\SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1' -Tenant prod -IncludeHealthChecker:$false
```

The unchanged CMD uses Windows PowerShell 5.1 with -File, which cannot pass an explicit false switch. For that launcher, disable through the local JSON; the explicit false syntax above applies to a PowerShell script invocation.

Offline tests cover absent/true/false config, explicit switch precedence, invalid booleans, receipt/phase gating and report-only behavior in Windows PowerShell 5.1 and PowerShell 7.

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
- `HealthCheckerExcludedServers` (array, empty by default): servers that stay in the whole inventory
  but that Health Checker never starts, contacts or runs. Names are matched case-insensitively on
  the short name and the FQDN. Their summary row is `NotRun` with "excluded by configuration", with
  no warning, and they are not counted in `ServersTotal`. The installer does not need to be run for
  them.
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
| `TaskTriggerFailed` | Task not installed, files not deployed, start refused, task logon failure, or task not running and no fresh result after 3 minutes (last task result in hexadecimal: check local administrator rights, batch logon right and password). |
| `TaskTimeout` | The task kept running without a result within the per-server limit, or Health Checker stopped on the server. |
| `StaleResult` | The on-demand start failed and the last daily result is older than the limit. |
| `NoOutput` | Health Checker ended without XML (including Exchange Management Shell load failure). |
| `UnsafeWorkFolder` | The collector work folder or one of its parents is not restricted to administrators; nothing was run. |
| `IntegrityFailed`, `NotElevated`, `Timeout`, `LaunchFailed`, `ParseFailed`, `NotRun` | As before. |

Messages never contain a domain name, an address or a secret.

An interrupted run (Ctrl+C) is never reported as a success: the banner shows `CANCELLED` with a
warning line, no receipt or success status is written, and the exit code is 130.

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

- Before any change on a server, the installer checks that the service account is a local
  administrator there, directly or through domain groups (local group members over CIM/DCOM, the
  account's groups from Active Directory with a read-only LDAP query, no credential asked). A server
  where it is proven not to be is refused (error); when it cannot be determined (Active Directory
  unavailable, local group nested in Administrators), a warning is written.
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


## Combined A/B/C acceptance

Run one offline acceptance command from the repository root:

```powershell
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File SmartM365/Tests/Test-SmartM365ExchangeClosureOffline.ps1
```

The command runs synthetic suites on Windows PowerShell 5.1 and PowerShell 7, keeps per-suite logs and a results.json, and exits nonzero if any applicable suite fails. The existing receipt-reader suite requires PS7 and is run there only. It covers receipt alignment and rejection guards, lifecycle cleanup, deterministic mailbox HTML, modes and serialized objects, calendar backend, bounded certificates, Health Checker, cache copies and weekly publication. It never launches a collector or contacts Exchange.

After publication and source deployment, one operator acceptance block may launch Mailboxes WithADPermissions, Calendars and Infrastructure directly in sequence, without the orchestrator. Inventory-only and OnlyADPermission-only branches remain covered offline; distinct production executions of those modes would collect additional data and are not implied by this offline test. Expect qualification failures to remain explicit when source data is incomplete or a certificate server times out. No server repair, task/runner change or rights change is included.
