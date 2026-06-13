# On-Premises Servers And Storage

Exchange 2016 on-premises infrastructure inventory for server, compute, memory, logical disk, physical disk, optional database path, optional service health, and decommissioning summary exports.

## Scripts

- `SmartM365-Exchange-OnPrem-ServersAndStorage-Inventory.ps1`: inventories Exchange 2016 servers through Exchange Management Shell plus WMI/DCOM only. It does not use WinRM, PowerShell remoting, `Get-CimInstance`, or `New-CimSession`.

## Output

By default, output is written under:

```text
{{DataAllRootPath}}\Exchange\OnPrem\ServersAndStorage\<RunId>
```

Use `-OutputRoot` to override the root folder for a specific run.
