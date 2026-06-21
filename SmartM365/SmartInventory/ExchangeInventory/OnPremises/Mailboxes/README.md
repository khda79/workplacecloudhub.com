# On-Premises Mailboxes

Exchange 2016 local mailbox inventory script with integrated reporting.

## Scripts

- `SmartM365-Exchange-Local-Mailboxes-Inventory.ps1`: detailed mailbox, permissions, statistics, archive, retention, mobile device inventory, and integrated mailbox reporting (`-GenerateReport` / `-ReportOnly`). Use `-DryRun` to validate paths, report CSV schema, and Exchange prerequisites without collecting inventory.

## Launchers

- `Start-SmartM365-Exchange-Local-Mailboxes-Inventory-Prod.cmd` / `Start-SmartM365-Exchange-Local-Mailboxes-Inventory-Test.cmd`: standard local mailbox inventory.
- `Start-SmartM365-Exchange-Local-Mailboxes-ADPermissions-Prod.cmd` / `Start-SmartM365-Exchange-Local-Mailboxes-ADPermissions-Test.cmd`: AD permission export only (`-OnlyADPermission`).
- `Start-SmartM365-Exchange-Local-Mailboxes-InventoryWithADPermissions-Prod.cmd` / `Start-SmartM365-Exchange-Local-Mailboxes-InventoryWithADPermissions-Test.cmd`: standard inventory with AD permissions (`-IncludeADPermission`).
