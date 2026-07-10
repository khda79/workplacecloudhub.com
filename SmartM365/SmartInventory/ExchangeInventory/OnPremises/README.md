# Exchange On-Premises Inventory

Exchange 2016 on-premises scripts that require Exchange Management Tools or the Exchange Management Shell.

## Organization

- `Mailboxes/`: local mailbox inventory and reporting.
- `ProxyAddresses/`: proxy address audit for user/shared mailboxes. Standard Test/Prod launchers are read-only; separate `Write-*` launchers are guarded for remediation.
- `ServersAndStorage/`: Exchange 2016 server, compute, memory, disk, optional database path, and service health inventory.
