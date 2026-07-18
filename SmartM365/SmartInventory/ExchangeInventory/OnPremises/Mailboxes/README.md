# On-Premises Mailboxes

Exchange 2016 local mailbox inventory script with integrated reporting.

## Scripts

- `SmartM365-Exchange-Local-Mailboxes-Inventory.ps1`: detailed mailbox, permissions, statistics, archive, retention, mobile device inventory, and integrated mailbox reporting (`-GenerateReport` / `-ReportOnly`). Use `-DryRun` to validate paths, report CSV schema, and Exchange prerequisites without collecting inventory.
- Version 1.39 enriches the local-mailbox CSV for migration readiness with full prefixed proxy addresses, Exchange/archive GUIDs, LegacyExchangeDN, hold state, moderation and delivery restrictions.
- The existing `EmailAddresses` column remains unchanged for compatibility; the complete prefixed collection, including X500 values, is exported in `EmailAddressesAll`.

## Launchers

- `SmartM365\SmartInventory\Launchers\OnPremises\Start-SmartM365-Exchange-Local-Mailboxes-Inventory.cmd`: standard production mailbox inventory.
- `SmartM365\SmartInventory\Launchers\OnPremises\Start-SmartM365-Exchange-Local-Mailboxes-Inventory-OnlyADPermissions.cmd`: AD permission export only (`-OnlyADPermission`).
- `SmartM365\SmartInventory\Launchers\OnPremises\Start-SmartM365-Exchange-Local-Mailboxes-InventoryWithADPermissions.cmd`: standard inventory with AD permissions (`-IncludeADPermission`).
## Configuration

SharePoint upload is inherited from the tenant/global configuration by default. If `SmartM365-Exchange-Local-Mailboxes-Inventory.local.json` already exists on an Exchange server from an older template, make sure it contains:

```json
"EnableSharePointUpload": "__USE_GLOBAL__"
```

or set it explicitly to `true` for that script.
