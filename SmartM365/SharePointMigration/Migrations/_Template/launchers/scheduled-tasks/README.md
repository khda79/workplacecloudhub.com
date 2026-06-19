# Scheduled Task Launchers

Use grouped launchers from Windows Task Scheduler:

- `files` for file scans and file comparisons.
- `permissions` for permission scans and permission comparisons.

Recommended Task Scheduler fields:

- Program/script: `cmd.exe`.
- Add arguments: `/d /c "\\server\share\SmartM365\SharePointMigration\Migrations\<MigrationName>\launchers\scheduled-tasks\files\01-Scan-Source-Files-Scheduled.cmd"` plus optional launcher parameters.
- Start in: leave empty, or use the full UNC path to the selected grouped folder.

The scheduled task account must have read/write access to the share because outputs are written under the migration folder. The `.cmd` files use `pushd`, so UNC paths are temporarily mapped by Windows while the task runs.
