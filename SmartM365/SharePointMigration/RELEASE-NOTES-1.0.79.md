# Smart SharePoint Migration Toolkit 1.0.79

Release 2026-10-27. Dashboard 1.0.79; generic launcher 1.0.34.
This release adds configurable shared storage for migration data.

## Shared storage for migrations, archives, logs and the Entra cache

- `Config/storage.json.txt`, created from `Config/storage.json.template`,
  defines `MigrationsRoot`, `ArchivedMigrationsRoot`, `LogsRoot` and
  `EntraUsersCacheRoot`. Empty values keep the previous layout inside the
  toolkit: `Migrations`, `ArchivedMigrations`, `Migrations\logs` and
  `_Local\EntraUsersCache`. An empty `LogsRoot` follows `MigrationsRoot`.
- Relative values resolve from the toolkit and environment variables are
  expanded. An unreadable storage file stops the scripts instead of silently
  falling back to the toolkit folders.
- A new shared resolver, `Scripts/Launchers/SmartM365-SharePointMigration-Storage.ps1`
  (Windows PowerShell 5.1 compatible), is used by the dashboard, the generic
  launcher, GUI runs and activity, dashboard batches, source/destination/
  comparison batches, farm diagnostics, site operation launchers, the new
  migration wizard and the template updater.
- A migration's `Comparison.EntraUsersCachePath` still overrides the shared
  Entra users cache folder.

## Launchers and template updates

- Template `.cmd` launchers 1.0.1 read `toolkit.root.txt` at the root of the
  migrations folder when migrations live outside the toolkit.
- `Update-MigrationsFromTemplate.cmd` resolves the configured migrations
  folder and now refreshes launchers in their `interactive/*` and
  `scheduled-tasks/*` subfolders; previously only top-level launchers were
  considered.
- Archiving a migration between two volumes or shares uses a verified copy
  (file count and bytes) before removing the source folder.

## Installation and validation

Download the ZIP, SHA256 checksum and file manifest. Verify the ZIP hash, then
extract the complete package and start
`SmartM365/SharePointMigration/Start-SmartM365-SharePointMigration-GUI.cmd`.
Update application files with the GUI and jobs closed; preserve private
migrations, archives, mappings, local authentication/configuration, the
runtime `Config/storage.json.txt` and installed Python.

Moving existing data to a share is an operator action: copy the folders while
nothing runs, verify the copy, set the keys in `Config/storage.json.txt`, write
`toolkit.root.txt` (toolkit UNC path, ASCII) at the new migrations root, run
`Update-MigrationsFromTemplate.cmd`, and re-point scheduled tasks that call
launchers from the former folder.

Windows and PowerShell 7.4 or later are required for the dashboard. SharePoint
Server collection requires Windows PowerShell 5.1 on an appropriate server.
Python runtimes and third-party modules are not bundled. The public signing
certificate and optional trust installer are included; trust is an explicit
local administrative choice. See README for operational prerequisites.

Release validation is offline: synthetic toolkits and migration folders,
fictional data, no tenant or farm connection. It does not qualify real
permissions, authentication or complete migration coverage. No live scan,
comparison, migration or cleanup is performed to prepare this release.
