# SmartAzure Backup Inventory

Azure backup and Recovery Services posture inventory scripts.

## Scripts

- `SmartAzure-Backup-Inventory.ps1`: exports Recovery Services vaults, vault backup/security settings, protected backup items, Azure VMs protected by backup, Azure VMs without detected Recovery Services protection, and subscription summaries.

Outputs use the shared SmartAzure tenant context:

```text
SmartAzure\Data\Tenants\<TenantKey>\DATA-ALL\Azure\Backup\<RunId>
SmartAzure\Data\Tenants\<TenantKey>\DATA-LAST
```

If the root `Data` folder cannot be created or written, scripts fall back to `Output\Tenants\<TenantKey>\...` next to the script.
