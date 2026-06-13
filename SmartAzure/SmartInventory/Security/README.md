# SmartAzure Security Inventory

Security posture inventory scripts for Azure Resource Manager and Microsoft Defender for Cloud.

## Scripts

- `SmartAzure-DefenderForCloud-Inventory.ps1`: exports Defender for Cloud plans, secure score, secure score controls, security recommendations, auto-provisioning settings, security contacts, regulatory compliance standards, controls, assessments, and subscription summaries.
- `SmartAzure-KeyVaultSecurity-Inventory.ps1`: exports Key Vault public access, purge protection, soft delete, RBAC/access policy model, secrets, certificates, and expiration posture.

Outputs use the shared SmartAzure tenant context:

```text
SmartAzure\Data\Tenants\<TenantKey>\DATA-ALL\Azure\Security\<RunId>
SmartAzure\Data\Tenants\<TenantKey>\DATA-LAST
```

If the root `Data` folder cannot be created or written, scripts fall back to `Output\Tenants\<TenantKey>\...` next to the script.
