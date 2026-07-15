# Exchange On-Premises Calendar Permissions

`SmartM365-Exchange-MailboxCalendarPermissions-Inventory.ps1` inventories Exchange 2016 mailbox calendar permissions.

- Runtime: Windows PowerShell 5.1 on a host with Exchange Management Tools.
- Authentication: current Exchange Management Shell security context.
- Data permissions: read access to Exchange recipients, folder statistics, and mailbox folder permissions.
- Processing: sequential to stay compatible with the Exchange 2016 management snap-in.
- Output: `Exchange_OnPrem_MailboxCalendarPermissions_AllDomains.csv`.
- Weekly history: `Exchange\OnPrem\CalendarPermissions\WeeklyHistory`.

Use the dedicated launcher because it copies the script, configuration, and SmartM365 compatibility module to a writable local cache before execution.
