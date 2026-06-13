# Citrix On-Premises Inventory

Inventory scripts for Citrix on-premises reporting, Power BI datasets, operational checks, and support handoff.

## Organization

- `CVAD/`: Citrix Virtual Apps and Desktops site, delivery, policy, machine/session health, hosting, and power inventory.
- `StoreFront/`: StoreFront deployment, store, authentication, receiver, gateway, farm, and beacon inventory.
- `Licensing/`: Citrix licensing configuration and licensing SDK inventory.

For every SmartInventory `.ps1`, a `Start-<ScriptName>-Prod.cmd` and `Start-<ScriptName>-Test.cmd` launcher is stored next to the script.
