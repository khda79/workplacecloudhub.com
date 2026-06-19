# Scheduled Task Launchers

Use these `.cmd` files from Windows Task Scheduler when a migration action must
run unattended. They are the same actions as the interactive launchers, but they
do not call `pause` and they return the PowerShell exit code to Task Scheduler.

Recommended Task Scheduler fields:

- Program/script: `cmd.exe`.
- Add arguments: `/d /c "\\server\share\SmartM365\SharePointMigration\Migrations\<MigrationName>\launchers\scheduled-tasks\<launcher>.cmd"` plus optional launcher parameters such as `-Force`, `-UseCertificate`, `-SourceCsv <path>`, or `-TargetCsv <path>`.
- Start in: leave empty, or use the full UNC path to this `scheduled-tasks`
  folder.

When the Git copy is on a network share, prefer UNC paths over mapped drives.
The scheduled task account must have read/write access to the share because the
launcher writes logs, scan CSVs, comparisons, and generated operation files back
under the migration folder. The `.cmd` files use `pushd`, so UNC paths are
temporarily mapped by Windows while the task runs.

The launcher writes the normal migration transcript logs under the configured
`logs` folder.
