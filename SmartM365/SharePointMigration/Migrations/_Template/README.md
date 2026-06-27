# Migration Template

Copy this folder to create a new migration.

Example:

```powershell
Copy-Item -Recurse .\Migrations\_Template .\Migrations\Spain
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
If the repository copy is on a network share, configure scheduled tasks with
`cmd.exe /d /c "\\server\share\...\launchers\scheduled-tasks\files\<launcher>.cmd"`
and use UNC paths instead of mapped drives. The task account needs read/write
access to the share because outputs are written under the migration folder.

Run source and target file scans close together before file comparison. The
template uses `Comparison.MaxScanAgeDifferenceHours` to block stale scan pairs
and `Comparison.ModifiedDateToleranceMinutes` to flag matched files where the
destination is older than the source. It also normalizes source `Modified`
values as local time and target `Modified` values as UTC by default, which avoids
false positives from the common SP2019 local-time versus SPO UTC offset.

Runtime logs are written to the migration `logs` folder by default. Source and
target scan URL files are derived automatically from `migration.mapping.txt` unless
explicit `Source.UrlsFile` and `Target.UrlsFile` values are configured. Scan CSVs,
comparison CSVs, Excel files, and generated operation scripts stay in their
dedicated output folders.
