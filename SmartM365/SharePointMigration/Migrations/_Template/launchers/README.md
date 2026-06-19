# Migration Launchers

Launchers are grouped by execution mode and functional area.

- `interactive\files` and `scheduled-tasks\files`: file inventory, file comparison, and source history comparison.
- `interactive\permissions` and `scheduled-tasks\permissions`: permission inventory and comparison.

The `.cmd` files kept directly under `launchers` and `launchers\scheduled-tasks` are compatibility wrappers. Prefer the grouped paths for new shortcuts and scheduled tasks.

For Windows Task Scheduler from a network share, use `cmd.exe` with `/d /c` and a full UNC path to the grouped scheduled launcher. Avoid mapped drives.
