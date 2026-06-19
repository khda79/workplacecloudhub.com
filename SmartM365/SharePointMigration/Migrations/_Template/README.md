# Migration Template

Copy this folder to create a new migration.

Example:

```powershell
Copy-Item -Recurse .\Migrations\_Template .\Migrations\Spain
```

Then update:

- `migration.config.psd1`
- `migration.config.source.txt`
- `migration.config.target.txt`

The `.cmd` files in `launchers` automatically detect the migration folder name.
If the copied folder is `Migrations\Spain`, they run with `-MigrationName Spain`.
Use the `.cmd` files in `launchers\scheduled-tasks` for Windows Task Scheduler:
they run the same actions without pausing and return the PowerShell exit code.
If the repository copy is on a network share, configure scheduled tasks with
`cmd.exe /d /c "\\server\share\...\launchers\scheduled-tasks\<launcher>.cmd"`
and use UNC paths instead of mapped drives. The task account needs read/write
access to the share because outputs are written under the migration folder.

Run source and target file scans close together before file comparison. The
template uses `Comparison.MaxScanAgeDifferenceHours` to block stale scan pairs
and `Comparison.ModifiedDateToleranceMinutes` to flag matched files where the
destination is older than the source. It also normalizes source `Modified`
values as local time and target `Modified` values as UTC by default, which avoids
false positives from the common SP2019 local-time versus SPO UTC offset.

Runtime logs are written to the migration `logs` folder by default. Scan CSVs,
comparison CSVs, Excel files, and generated operation scripts stay in their
dedicated output folders.
