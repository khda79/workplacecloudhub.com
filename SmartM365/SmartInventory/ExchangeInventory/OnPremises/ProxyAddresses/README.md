# Proxy Addresses

Exchange 2016 proxy address audit and remediation scripts.

## Scripts

- `SmartM365-Check-ProxyAddresses-Exchange.ps1`: audits user/shared mailbox proxy addresses by OU or forest-wide scope. It is read-only by default; remediation only runs when `-AddMissingAddress` is explicitly passed.

## Launchers

- `SmartM365\SmartInventory\Launchers\OnPremises\Start-SmartM365-Check-ProxyAddresses-Exchange-ReadOnly.cmd`: production tenant audit launcher, read-only by default.
- `SmartM365\SmartInventory\Launchers\OnPremises\Start-SmartM365-Check-ProxyAddresses-Exchange-Write.cmd`: production tenant write launcher. Requires typing `WRITE` and runs with `-AddMissingAddress -SkipAllowListCsv`.