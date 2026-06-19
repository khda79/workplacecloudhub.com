# Scheduled Task Launchers

Use these `.cmd` files from Windows Task Scheduler when a migration action must
run unattended. They are the same actions as the interactive launchers, but they
do not call `pause` and they return the PowerShell exit code to Task Scheduler.

Recommended Task Scheduler fields:

- Program/script: full path to the selected `.cmd` file.
- Start in: full path to this `scheduled-tasks` folder.
- Add arguments: optional launcher parameters such as `-Force`, `-UseCertificate`,
  `-SourceCsv <path>`, or `-TargetCsv <path>`.

The launcher writes the normal migration transcript logs under the configured
`logs` folder.
