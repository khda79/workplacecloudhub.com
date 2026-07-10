# Exchange SE AD Pre-Check

Read-only Active Directory pre-check before Exchange SE forest preparation steps such as `/PrepareSchema`, `/PrepareAD`, `/PrepareAllDomains`, or targeted `/PrepareDomain` operations.

## Script

- `SmartM365-ExchangeSE-ADPreCheck-Inventory.ps1`

The script checks forest/domain discovery, FSMO roles, privileged group memberships, domain controllers, Global Catalog availability, Exchange AD schema/configuration/domain versions, replication metadata, `repadmin`, `dcdiag`, DNS SRV records, AD sites/subnets, SYSVOL/NETLOGON, DFSR state, recent AD-related events, Schema Master backup events, and Schema Master connectivity.

## Requirements

- Run from a domain-connected machine with RSAT ActiveDirectory module.
- Use an account with enough read access across the target domains.
- Run elevated when possible.
- `repadmin.exe`, `dcdiag.exe`, and `dfsrdiag.exe` improve coverage.

## Output

By default, output is written under:

```text
{{DataAllRootPath}}\Exchange\OnPrem\ExchangeSEPreCheck\<RunId>
```

Latest CSV copies are written to `{{LatestCsvFolderPath}}`. CSV and log uploads use the shared SmartM365 SharePoint helper when enabled.

## Examples

```powershell
& 'C:\Program Files\PowerShell\7\pwsh.exe' -NoProfile -ExecutionPolicy Bypass -File 'C:\Path\To\SmartM365\SmartInventory\ExchangeInventory\OnPremises\ExchangeSEPreCheck\SmartM365-ExchangeSE-ADPreCheck-Inventory.ps1' -Tenant prod
```

```powershell
& 'C:\Program Files\PowerShell\7\pwsh.exe' -NoProfile -ExecutionPolicy Bypass -File 'C:\Path\To\SmartM365\SmartInventory\ExchangeInventory\OnPremises\ExchangeSEPreCheck\SmartM365-ExchangeSE-ADPreCheck-Inventory.ps1' -Tenant prod -SkipDcdiag -SkipEvents
```
