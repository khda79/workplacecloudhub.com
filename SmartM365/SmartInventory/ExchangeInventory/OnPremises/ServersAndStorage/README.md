# On-Premises Infrastructure And Readiness

Exchange 2016 on-premises infrastructure and migration readiness inventory for servers, compute, memory, logical disks, physical disks, optional database paths, optional service health, Exchange schema, and key Exchange configuration for Exchange SE / Exchange Online planning.

## Scripts

- `SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1`: inventories Exchange 2016 infrastructure through Exchange Management Shell plus WMI/DCOM only, then renders a full HTML readiness report. It does not use WinRM, PowerShell remoting, `Get-CimInstance`, or `New-CimSession`.

## Output

By default, output is written under:

```text
{{DataAllRootPath}}\Exchange\OnPrem\ServersAndStorage\<RunId>
```

Use `-OutputRoot` to override the root folder for a specific run.
