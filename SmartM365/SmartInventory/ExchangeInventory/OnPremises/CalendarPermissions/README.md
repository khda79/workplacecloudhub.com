# Exchange On-Premises Calendar Permissions

`SmartM365-Exchange-MailboxCalendarPermissions-Inventory.ps1` inventories Exchange 2016 mailbox calendar permissions.

- Runtime: Windows PowerShell 5.1 on a host with Exchange Management Tools.
- Authentication: current Exchange Management Shell security context.
- Data permissions: read access to Exchange recipients, folder statistics, and mailbox folder permissions.
- Processing: sequential to stay compatible with the Exchange 2016 management snap-in.
- Backend resilience: one bounded `Test-MAPIConnectivity` preflight per mailbox database (30 seconds by default; set `BackendPreflightTimeoutSeconds` to `0` to disable it).
- Output: `Exchange_OnPrem_MailboxCalendarPermissions_AllDomains.csv`.
- Weekly history: the permissions and mailbox-error CSVs are published together once per run under `Exchange\OnPrem\CalendarPermissions\WeeklyHistory`.

Use the dedicated launcher because it copies the script, configuration, and SmartM365 compatibility module to a writable local cache before execution.
