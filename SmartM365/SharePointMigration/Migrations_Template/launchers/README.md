# Migration Launchers

Launchers are grouped by execution mode and functional area.

- `interactive\files` and `scheduled-tasks\files`: file inventory, file comparison, and source history comparison.
- `interactive\permissions` and `scheduled-tasks\permissions`: permission inventory and comparison.
- `interactive\operations` and `scheduled-tasks\operations`: optional SPO
  page-comment and site-lock administration. Interactive launchers confirm
  before changing a site; scheduled launchers require `-Execute` to change it.

For Windows Task Scheduler from a network share, use `cmd.exe` with `/d /c` and a full UNC path to the grouped scheduled launcher. Avoid mapped drives.
