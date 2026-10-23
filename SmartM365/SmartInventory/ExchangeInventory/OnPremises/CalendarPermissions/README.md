# Exchange On-Premises Calendar Permissions

`SmartM365-Exchange-MailboxCalendarPermissions-Inventory.ps1` inventories Exchange 2016 mailbox calendar permissions (version 1.11).

- Runtime: Windows PowerShell 5.1 with Exchange Management Tools.
- Authentication: current Windows account, Kerberos to the local host FQDN at `/PowerShell/` (`Microsoft.Exchange`). No additional permissions are granted.
- Data permissions: Exchange recipient, folder statistics, folder permission and `Test-MAPIConnectivity` read/probe permissions, including their parameters and recipient scopes.
- Processing: one owned remote session per run; sequential cmdlet calls, without parallel workers or automatic reconnect/fallback. The session is removed on success and failure; collection and session durations are logged.
- NoLanguage compatibility: only cmdlet calls and parameters cross the endpoint. Filtering, localized folder selection and permission projections run on the client against deserialized objects. No remote scriptblocks are submitted.
- Backend preflight: one bounded `Test-MAPIConnectivity` probe per database (30 seconds by default; `BackendPreflightTimeoutSeconds=0` disables it). A SystemMailbox probe denial is inconclusive for target calendars: its diagnostic is logged and target reads continue. `AccessDenied` never marks a database or server unavailable. Proven availability failures are cached only for the affected database; other databases on that server continue.
- Output: `Exchange_OnPrem_MailboxCalendarPermissions_AllDomains.csv`, with the existing tenant columns plus `Mailbox`, `UPN`, `CalendarFolder`, `User`, `AccessRights`.
- Qualification: a successful calendar ACL read qualifies that calendar, even with only Default/Anonymous entries. `EmitNoPermRow` controls `(none)` rows; with it disabled, proven empty ACLs produce a header-only CSV. Missing/denied reads never become `(none)` or zero.
- Failure contract: when mailboxes were expected but no calendar ACL read succeeded, the run fails with its diagnostic CSV/log preserved. It does not replace the previous canonical permissions CSV, publish weekly history or apply retention. An actual zero-mailbox enumeration remains distinct.
- Diagnostics: failures include category, operation, attempts, duration, database and backend. Mixed successful/failed reads retain `CompletedWithWarnings` and their diagnostic CSV.
- Weekly history: qualified permissions and mailbox-error CSVs are published together once per run under `Exchange\OnPrem\CalendarPermissions\WeeklyHistory`.

Use `Start-SmartM365-Exchange-Local-MailboxCalendarPermissions-Inventory.cmd` directly. It copies the collector, configuration and complete Core module tree (including the lot A session helper) into a writable local cache. No orchestrator is required. A `-MaxItems` smoke run uses separate suffixed filenames and skips weekly history; full qualification requires a subsequent unrestricted run.

The script-local `SmartM365-Exchange-MailboxCalendarPermissions-Inventory.local.json.txt` (or resolved `.local.json`) controls the connection:

```json
"ExchangeConnectionMode": "Remote"
```

`Remote` is the default, automatically added from the template when absent without overwriting existing local values. Connection/import/scope failures log `state=Failed; no fallback` and stop. For rollback, set the same local key to `LocalSnapin` in the source configuration copied by the launcher, then relaunch directly. This preserves native snap-in loading and forest scope; the diagnostic/qualification fixes apply in both modes. Editing only the disposable cache does not persist across launches. An invalid mode is rejected.

Offline validation (no Exchange connection or collector execution):

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File SmartM365/Tests/Test-SmartM365ExchangeCalendarBackendOffline.ps1
pwsh.exe -NoProfile -File SmartM365/Tests/Test-SmartM365ExchangeCalendarBackendOffline.ps1
```

The tests exercise both connection modes with stubs, real PSSerializer round trips, denied/unknown/empty preflight results, database isolation, target denials, primary/all-calendar modes, real local CSV validation/publication helpers and failure retention. Session ownership/cleanup is also covered by `Test-SmartM365ExchangeOnPremSessionOffline.ps1`.

After operator deployment, test directly under the service account: check version 1.11, `Remote; state=Ready` once, forest scope, positive qualified-calendar count, localized folders/ACLs, diagnostic categories, canonical CSV and receipt hashes, duration, and session cleanup. Confirm a mailbox access denial does not produce server-wide `BackendPreviouslyUnavailable`. Only the operator runs these server checks; offline success does not establish production access or completeness.