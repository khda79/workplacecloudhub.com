# SmartAzure Storage Inventory

Azure Storage security posture inventory scripts.

## Scripts

- `SmartAzure-StorageSecurity-Inventory.ps1`: exports storage account public network access, blob anonymous access, shared key, HTTPS only, minimum TLS, firewall settings, and blob service properties.

Outputs use the shared SmartAzure tenant context:

```text
SmartAzure\Data\Tenants\<TenantKey>\DATA-ALL\Azure\Storage\<RunId>
SmartAzure\Data\Tenants\<TenantKey>\DATA-LAST
```

If the root `Data` folder cannot be created or written, scripts fall back to `Output\Tenants\<TenantKey>\...` next to the script.
