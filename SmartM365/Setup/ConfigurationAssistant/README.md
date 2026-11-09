# SmartInventory configuration assistant

The English-language assistant prepares SmartInventory from a product copy already present on a local path or UNC share. Start `Lancer-SmartInventory-GUI.cmd` on the target Windows machine. The GUI requires PowerShell 7 and a Windows desktop session.

## Simplified setup

1. Select or prepare a tenant profile, connect to Microsoft 365 and enter the environment and CSV/log workspace.
2. Check the share and install missing machine prerequisites after confirmation.
3. Select or create a standard domain service account, and verify its access to the shares. When Exchange on-premises is selected, the assistant offers the View-Only Organization Management role group if needed.
4. Review the planned Microsoft 365 application, Teams workspace and associated SharePoint site, and configuration changes. Microsoft 365 resources are changed only after confirmation. Teams notification workflow configuration and tests are outside this assistant.
5. The assistant displays progress and applies conflict-free bootstrap settings to the profile. Conflicts or partial failures require review. It offers an encrypted PFX export and then offers to launch the existing scheduled-task installer.

The scheduled-task installer uses the selected profile and standard service account, checks the required local logon rights, and imports the application certificate into that account's profile. Passwords are entered interactively and are not stored in the profile. The installer controls whether the resulting task is enabled. gMSA accounts are not part of this assistant version.

## Logs and private data

Assistant logs are written in `Log` when available. Git tracks the empty log directory only. `Config/Certificats` retains only its ignore rules and placeholder in Git; PFX files and other contents remain private. Tenant and global local configuration, lab identities, passwords, tokens, data exports and logs must not be committed.

## Validation and qualification

`-ValidateOnly` performs passive local checks without applying setup changes. `-OfflineSmokeTest` exercises WPF navigation using disposable synthetic files and does not connect to Microsoft 365 or perform setup actions. The `Tests` directory contains focused offline fixtures.

Lab qualification has exercised profile preparation, service-account/share access, Microsoft 365 bootstrap, PFX export and the scheduled-task installer. These observations do not establish complete production qualification; they depend on the tested machine, identity and tenant. The Exchange mailbox inventory was separately exercised in the lab. Review current failures, signatures and resource inclusion before a Git release.
