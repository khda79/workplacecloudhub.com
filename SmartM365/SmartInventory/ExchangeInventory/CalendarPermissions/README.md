# Exchange Online Calendar Permissions

`SmartM365-EXO-Mailboxes-CalPerm_Inventory.ps1` inventories Exchange Online mailbox calendar permissions.

- Runtime: PowerShell 7.
- Authentication: Exchange Online app-only certificate authentication.
- Data permissions: `Exchange.ManageAsApp` plus read-only Exchange RBAC; no Microsoft Graph data scope is required.
- Default mode: primary calendar with four persistent parallel workers.
- Output: `Exchange_EXO_MailboxCalendarPermissions_AllDomains.csv`.
- Weekly history: `Exchange\EXO\CalendarPermissions\WeeklyHistory`.

Use `-TopMailboxes N` for a bounded smoke test. Use `-ParallelThrottle 1` to force sequential primary-calendar processing.
