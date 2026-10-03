# Migration Template

Use `+ New` in the dashboard to validate the site URLs and create a migration
from this template. Manual copying is also available when the wizard cannot
be used:

Example:

```powershell
Copy-Item -Recurse .\Migrations_Template .\Migrations\Spain
```

Then update:

- `migration.config.psd1`
- `migration.mapping.txt`

Set `Source.Type` and `Target.Type` to `SP2016`, `SP2019`, or `SPO`; setting both to `SPO` enables SharePoint Online to SharePoint Online migrations.
The grouped `.cmd` files under `launchers\interactive` and
`launchers\scheduled-tasks` automatically detect the migration folder name. If
the copied folder is `Migrations\Spain`, they run with `-MigrationName Spain`.
Use `launchers\interactive\files` for file inventory/comparison,
`launchers\interactive\permissions` for permission inventory/comparison, and
the matching `launchers\scheduled-tasks\files` or
`launchers\scheduled-tasks\permissions` folders for Windows Task Scheduler.
The `operations` folders provide launchers for disabling page comments and
setting a SharePoint Online site's lock state. These require an SPO target,
`Target.SiteUrl`, `Target.TenantAdminUrl`, Windows PowerShell 5.1 and the
SharePoint Online Management Shell. Interactive launchers prompt for preview
or execution and require typing `YES` before a change. Scheduled launchers
preview by default; add `-Execute` to apply a change. The site lock launcher
also requires an explicit `-LockState ReadOnly`, `Unlock`, or `NoAccess` for
scheduled use. Review the target site and requested state before scheduling.
If the repository copy is on a network share, configure scheduled tasks with
`cmd.exe /d /c "\\server\share\...\launchers\scheduled-tasks\files\<launcher>.cmd"`
and use UNC paths instead of mapped drives. The task account needs read/write
access to the share because outputs are written under the migration folder.

Run source and target file scans close together before file comparison. The
template uses `Comparison.MaxScanAgeDifferenceHours` to block stale scan pairs
and `Comparison.MaxScanAgeHours` to limit absolute scan age. Launcher scans
write a hash-verified `.csv.manifest.json.txt` receipt beside each CSV;
legacy scans use the timestamp in their filename. Permission comparison has
equivalent `PermissionMaxScanAgeDifferenceHours` and `PermissionMaxScanAgeHours` settings.
The GUI can also compare two permission scans from the same Source or Target
endpoint. `Output.PermissionHistoryComparisons` selects the report folder;
this offline comparison does not require a SharePoint or Entra connection.

`Comparison.ModifiedDateToleranceMinutes` flags matched files where the
destination is older than the source. It also normalizes source `Modified`
values as local time and target `Modified` values as UTC by default, which avoids
false positives from the common SP2019 local-time versus SPO UTC offset.

Runtime logs are written to the migration `logs` folder by default. Source and
target scan URL files are derived automatically from `migration.mapping.txt` unless
explicit `Source.UrlsFile` and `Target.UrlsFile` values are configured. Scan CSVs,
comparison CSVs, Excel files, and generated operation scripts stay in their
dedicated output folders.

Place ShareGate migration reports (`.xls`, `.xlsx`, or `.csv`) in
`ShareGate/MigrationReport`. The folder structure and its README are versioned;
the report files themselves remain local and ignored by Git.
