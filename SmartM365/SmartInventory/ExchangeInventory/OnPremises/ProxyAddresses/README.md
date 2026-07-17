# Proxy Addresses

Exchange 2016 proxy address audit and remediation scripts.

## Scripts

- `SmartM365-Check-ProxyAddresses-Exchange.ps1`: audits local and remote user/shared mailboxes by OU or forest-wide scope (`UserMailbox`, `SharedMailbox`, `RemoteUserMailbox`, and `RemoteSharedMailbox`). Local mailbox expected proxies use `Alias` plus the configured suffix. Remote mailboxes use their existing `RemoteRoutingAddress` first and fall back to `Alias` plus the configured suffix only when no usable routing address exists. It is read-only by default; remediation only runs when `-AddMissingAddress` is explicitly passed. Local recipients are updated with `Set-Mailbox`, remote recipients with `Set-RemoteMailbox`, and recipients with `EmailAddressPolicyEnabled=True` are always skipped. Before remediation, the script checks proposed addresses against every mail-enabled Exchange recipient in the forest, blocks addresses shared by multiple target recipients, and blocks addresses already assigned to another recipient. Each run also creates `Exchange_OnPrem_ProxyAddresses_<timestamp>.xlsx`, which groups the CSV outputs into the `Check`, `Summary`, and `Added` worksheets; the `Added` worksheet remains present with headers when no address was added.

## Excel dependency

- The `ImportExcel` PowerShell module must be installed for the SmartM365 execution account on the Exchange management server. The preflight stops with a clear error if it is missing; the script never installs modules automatically.

## Launchers

- `SmartM365\SmartInventory\Launchers\OnPremises\Start-SmartM365-Check-ProxyAddresses-Exchange-ReadOnly.cmd`: production tenant audit launcher, read-only across the entire forest (`-AllOrganizationalUnit`).
- `SmartM365\SmartInventory\Launchers\OnPremises\Start-SmartM365-Check-ProxyAddresses-Exchange-Write.cmd`: production tenant write launcher. Requires typing `WRITE` and runs across the entire forest with `-AllOrganizationalUnit -AddMissingAddress`.
