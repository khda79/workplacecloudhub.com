# SmartAzure Inventory

Azure inventory scripts for governance reviews, reporting, Power BI datasets, and operational checks.

## Organization

- `Governance/`: Azure estate, management group, subscription, resource, provider, lock, policy, and governance baseline exports.
- `RBAC/`: Azure RBAC role assignments, custom roles, privileged scopes, and orphaned principal checks.
- `Cost/`: optimization signals such as unused resources, budgets, reservations, and savings opportunities.
- `Network/`: public exposure, NSGs, public IPs, private endpoints, load balancers, application gateways, and DNS.
- `Security/`: Defender for Cloud, Key Vault posture, and security recommendations.
- `Backup/`: Recovery Services vaults, protected items, backup gaps, and recovery posture.
- `Storage/`: storage account security posture, public access, shared key, TLS, HTTPS, and firewall settings.
- `Advisor/`: Azure Advisor recommendations across all categories.

For every SmartInventory `.ps1`, a `Start-<ScriptName>-Prod.cmd` and `Start-<ScriptName>-Test.cmd` launcher is stored next to the script.
