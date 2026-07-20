# On-Premises Mailboxes

Exchange 2016 local mailbox inventory script with integrated reporting.

## Scripts

- `SmartM365-Exchange-Local-Mailboxes-Inventory.ps1`: detailed mailbox, permissions, statistics, archive, retention, mobile device inventory, and integrated mailbox reporting (`-GenerateReport` / `-ReportOnly`). Use `-DryRun` to validate paths, report CSV schema, and Exchange prerequisites without collecting inventory. Remote mailbox delegation is intentionally disabled by default because it performs two expensive permission queries per remote mailbox; enable it explicitly with `-IncludeRemoteMailboxDelegation`.
- Version 1.41 preserves the 1.40 schema/readiness fields and prevents the standard run from performing remote-mailbox delegation queries unless explicitly requested.
- Version 1.40 ajoute `InventorySchemaVersion`, conserve les propriétés de readiness de la v1.39 et expose l’état de collecte des gros éléments.
- La collecte des gros éléments est désactivée par défaut car elle est coûteuse. Utiliser `-CollectLargeItemStatistics` pour exécuter une estimation `Search-Mailbox` à 35 Mo et alimenter `LargeItemCount-Over-35MB`, `LargeItemCollectionStatus` et `LargeItemThresholdMB`.
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
