# SmartAzure Advisor Inventory

Azure Advisor recommendation inventory scripts.

## Scripts

- `SmartAzure-Advisor-Inventory.ps1`: exports Azure Advisor recommendations across all categories and subscription summaries.

Outputs use the shared SmartAzure tenant context:

```text
SmartAzure\Data\Tenants\<TenantKey>\DATA-ALL\Azure\Advisor\<RunId>
SmartAzure\Data\Tenants\<TenantKey>\DATA-LAST
```

If the root `Data` folder cannot be created or written, scripts fall back to `Output\Tenants\<TenantKey>\...` next to the script.
