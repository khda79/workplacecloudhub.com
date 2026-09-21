# Exchange Inventory

Exchange Online inventory and comparison scripts.

## Organization

- `AcceptedDomains/`: accepted domain export.
- `BackupProtection/`: Microsoft 365 Backup protected mailbox inventory through Microsoft Graph Backup Restore APIs.
- `CalendarPermissions/`: Exchange Online mailbox calendar permission inventory.
- `OnPremises/CalendarPermissions/`: Exchange 2016 mailbox calendar permission inventory.
- `Mailboxes/`: Exchange Online mailbox inventory.
- `Quarantine/`: read-only Exchange Online quarantined email metadata report with CSV and Excel mail delivery.
- `Migration/`: mailbox migration job inventory.
- `Permissions/`: mailbox permission report grouped by delegated user across on-premises and EXO sources.
- `OnPremises/Mailboxes/`: Exchange 2016 on-premises mailbox inventory and reporting.
- `OnPremises/ProxyAddresses/`: Exchange 2016 on-premises proxy address audit and remediation.
- `OnPremises/ServersAndStorage/`: Exchange 2016 server, compute, memory, disk, optional database path, and service health inventory.

## On-premises completion and mail limits

The Exchange infrastructure collector reports `CompletedWithWarnings` when its
own non-blocking warnings are present, so the orchestrator can retain the warning
state instead of reporting a clean success. The local mailbox inventory records a
successful `Mailbox inventory summary` send in a date marker protected by an
exclusive lock; additional runs on the same local day skip that summary only.
Failure notifications remain immediate. These controls require no new Exchange,
Graph, SMTP, or SharePoint permission and do not alter CSV contracts.
