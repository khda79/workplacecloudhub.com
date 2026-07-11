# SmartM365 Inventory Orchestrator

`SmartM365-Inventory-Orchestrator.ps1` (v1.2.0) is a PowerShell 7 resident scheduler that runs the SmartInventory scripts (ActiveDirectoryInventory, ExchangeInventory, M365Inventory, IntuneInventory, ...) unattended.

It is started by a single Windows Task Scheduler task (at server startup plus a daily trigger), loops with a one-minute tick, launches each job exactly at its scheduled occurrences, and exits cleanly after a configurable maximum lifetime (default 24 hours) so Task Scheduler restarts a fresh instance (memory recycling). The orchestrator recycle never interrupts a running job (see "Detached jobs and re-adoption").

## Files

| File | Purpose |
| --- | --- |
| `SmartM365-Inventory-Orchestrator.ps1` | Orchestrator script (PowerShell 7). |
| `SmartM365-Inventory-Orchestrator.local.json.template` | Safe committed template; copied to `SmartM365-Inventory-Orchestrator.local.json` at first run (the runtime `.local.json` is Git-ignored). |
| `Orchestrator-Jobs.json.template` | Safe committed jobs-manifest template (all schedules, neutral `AllowedServers`). |
| `Orchestrator-Jobs.json` | Runtime jobs manifest, auto-created from the template at first run and Git-ignored: it carries operational values (Enabled flags, schedules, real server names in `AllowedServers`). Hot reloaded on change. |
| `Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1` | Installs or removes the unattended Windows scheduled task under a dedicated service account. |
| `Start-SmartM365-Inventory-OrchestratorScheduledTask-Installer.cmd` | Interactive elevated launcher for scheduled-task installation or removal. |
| `Start-SmartM365-Inventory-Orchestrator-Prod.cmd` | Launcher: `-Tenant prod -Connect`. |
| `Start-SmartM365-Inventory-Orchestrator-Test.cmd` | Launcher: `-Tenant test -Connect`. |

Runtime files (tenant-isolated, created automatically, all Git-ignored). The orchestrator data and log folders are automatically suffixed with the local computer name (for example `{{DataAllRootPath}}\Orchestrator\SRV01`), so several servers can share the same UNC `DataAllRootPath`/`LogAllRootPath` without state, lock, CSV or log collisions:

| File | Location | Purpose |
| --- | --- | --- |
| `Orchestrator-State.json` | `{{DataAllRootPath}}\Orchestrator` | Per-job state (last occurrence, last run, running PID). Atomic writes. |
| `Orchestrator-Heartbeat.json` | `{{DataAllRootPath}}\Orchestrator` | Rewritten at every tick: timestamp, PID, running jobs. |
| `Orchestrator.lock` | `{{DataAllRootPath}}\Orchestrator` | Global lock; prevents two instances for the same tenant. Stale locks (dead PID) are recovered with a warning. |
| `Orchestrator_JobRuns_<yyyyMMdd>.csv` | `{{DataAllRootPath}}\Orchestrator\JobRuns` | Daily job-run tracking CSV (atomic writes). |
| `SmartM365-Inventory-Orchestrator_<yyyyMMdd>.log` | `{{LogAllRootPath}}\SmartM365-Inventory-Orchestrator` | Orchestrator log, daily rotation. |
| `<JobName>_<timestamp>.log` | `{{LogAllRootPath}}\SmartM365-Inventory-Orchestrator\Jobs\<JobName>` | One log per job execution (stdout + stderr of the child process). |

Because state, lock, logs and CSVs live under the tenant data roots plus a per-server subfolder, `prod` and `test` orchestrators are fully isolated, and several servers can run orchestrators against the same shared data roots. Note: each server keeps its own state, so a job allowed on several servers runs on each of them - pin every scheduled job to exactly one server through `AllowedServers` and treat the other servers as manual standby.

## Design

### Execution model

- One resident instance per tenant, bounded lifetime (`MaxLifetimeHours`, default 24), exit code 0 on recycle.
- 60-second tick: reload the manifest if it changed on disk, supervise running children, compute due occurrences, launch jobs, send the optional daily summary, rewrite the heartbeat, save state.
- Every job runs in its own detached child process (`-NoProfile -ExecutionPolicy Bypass`), which also isolates module/assembly conflicts between scripts (Graph SDK vs MSAL). The engine is `pwsh` by default; jobs with `PowerShellEdition = "WindowsPowerShell"` run in `powershell.exe` 5.1 instead (required by the Exchange on-premises scripts).

### Server allowlist

- `AllowedServers` in the orchestrator `.local.json` is the default list of computer names allowed to run jobs (empty = all servers allowed).
- Each job may carry its own `AllowedServers` list in the manifest; a non-empty job list overrides the default list for that job.
- A job whose effective allowlist does not contain the local computer name is never launched on this server - not by schedule, not by catch-up, and not by `-Force` (refused with a warning). Excluded jobs are listed once in the log at manifest load, and a dependency that is not allowed on this server never blocks its dependents.
- Typical use: run the cloud inventory jobs on the scheduler servers (default list) and pin the Exchange 2016 on-premises jobs to the Exchange server through their per-job list.

### Detached jobs and re-adoption (jobs longer than 24 hours)

- The child process redirects all of its output streams (`*>>`) into the per-run job log itself; there is no pipe to the orchestrator, so the child fully survives an orchestrator exit.
- The state file records the child PID and StartTime while a job is `Running`.
- On shutdown (lifetime reached or task stopped), the orchestrator stops launching new jobs, saves state and exits with code 0 WITHOUT waiting for or killing running jobs.
- On startup, the new instance checks each `Running` entry: if the PID still exists, is `pwsh`, and its StartTime matches the recorded one (5-second tolerance), the process is re-adopted and supervision resumes (timeout counted from the original StartTime, exit code read at the end). If the process is gone, the run is marked `Interrupted`, the retry policy applies and a notification is sent.
- The per-job overlap guard relies on the state file and re-adopted PIDs (not on process memory), so it stays effective across recycles, including for jobs longer than 24 hours that span several orchestrator lifecycles.

### Scheduling

- `Daily` with one or several `Times` per day, or `Weekly` with `DaysOfWeek` + `Times`.
- `MissedRunPolicy`:
  - `RunOnce` (default): a missed occurrence (orchestrator stopped or server off at the scheduled time) is caught up exactly once at restart. Several missed occurrences collapse into a single catch-up run (the most recent one).
  - `Skip`: missed occurrences are recorded as `Skipped` at startup and the job waits for its next occurrence.
- A brand-new job (no state yet) is initialized fast-forwarded: occurrences from before the job existed are not treated as missed. Note: re-enabling a long-disabled `RunOnce` job may trigger one catch-up run of its most recent occurrence (within the 8-day lookback window).
- Occurrences are computed in server local time.

### Concurrency, queueing, dependencies, overlap

- Global `MaxConcurrency` (default 2). Jobs due beyond the limit stay queued (their occurrence remains due, never lost) and start as soon as a slot frees. Re-adopted jobs count toward the limit.
- `DependsOn`: jobs due at the same occurrence run chained in topological order (the manifest is rejected at load time on a cycle). A dependent waits while a parent is running, due, or pending a retry. If a parent with `ContinueOnError=false` finally fails, dependents due at the same occurrence are marked `Skipped`.
- Overlap guards:
  - Global lock file: two orchestrator instances never run at the same time for the same tenant; a stale lock (dead PID) is recovered with a warning (exit code 3 when a live instance holds the lock).
  - Per job: a job still running at its next occurrence is not relaunched; the occurrence is marked `Skipped` with a warning in the log and the summary.

### Supervision, timeout, retries

- Each tick checks running children: exit code and duration are captured on exit; `TimeoutMinutes` triggers a full process-tree kill (`taskkill /T /F`). `TimeoutMinutes` may exceed 1440 for jobs longer than 24 hours.
- Retry policy per job: `MaxRetries` / `RetryDelaySeconds`. A failed run with retries left is recorded as `Retried`; the final failure sends the error email.

### Results and notifications

- Every execution is appended to the daily tracking CSV: `JobName, ScheduledTime, StartTime, EndTime, DurationSec, ExitCode, Status (Success/Failed/TimedOut/Skipped/Retried/Interrupted), RetryCount, LogPath`.
- Emails use `System.Net.Mail.MailMessage` with explicit UTF-8 sent through `SmtpClient` (never `Send-MailMessage`), `UseDefaultCredentials` or anonymous relay, IPv4 resolution of the SMTP endpoint with an optional pinned `RelayIp`, HTML bodies, no BCC by default.
  - `JobMailMode`: `Always` (email for every final job completion), `OnError` (final failures only, default), `Never` (no job emails). This is intentionally a dedicated key: the ecosystem-wide `SendMailMode` key carries the mail transport (Graph/SMTP/Both) and is not used by the orchestrator.
  - Optional daily HTML summary (`SendDailySummaryEmail` + `DailySummaryTime`) recapping all executions of the last 24 hours with color-coded statuses (inline styles, no external CSS).
  - A fatal error email is sent if the orchestrator itself crashes (also on an invalid manifest at startup), independent of `JobMailMode`.
  - When `SmtpServer`, `From` or the recipient is empty, mail is disabled with a warning.

## Jobs manifest schema (`Orchestrator-Jobs.json`)

```json
{
  "Jobs": [
    {
      "Name": "EXO-Mailboxes-Inventory",
      "ScriptPath": "ExchangeInventory\\Mailboxes\\SmartM365-EXO-Mailboxes-Inventory.ps1",
      "Arguments": "",
      "Enabled": false,
      "Group": "Exchange",
      "DependsOn": [],
      "TimeoutMinutes": 480,
      "MaxRetries": 1,
      "RetryDelaySeconds": 600,
      "ContinueOnError": false,
      "Schedule": {
        "Type": "Daily",
        "Times": ["01:00"],
        "MissedRunPolicy": "RunOnce"
      }
    }
  ]
}
```

| Field | Description |
| --- | --- |
| `Name` | Unique job name (letters, digits, `.`, `_`, `-`). Used for state, logs and CSV. |
| `ScriptPath` | Script path relative to the SmartInventory root (the parent folder of `Orchestrator`). |
| `Arguments` | Extra arguments appended verbatim to the child command line. `-Tenant <tenant>` (and `-Connect` when passed to the orchestrator) are always appended by the orchestrator; do not repeat them here. |
| `Enabled` | `false` by default; only enabled jobs are scheduled. |
| `Group` | Logical group (AD, Exchange, M365, Intune); informational. |
| `DependsOn` | List of job names that must complete first (cycle-checked at load). |
| `AllowedServers` | Computer names allowed to run this job. Empty = inherit the orchestrator `AllowedServers` default (empty default = all servers). |
| `PowerShellEdition` | `PowerShell7` (default, `pwsh`) or `WindowsPowerShell` (`powershell.exe` 5.1, required for Exchange on-premises scripts). |
| `TimeoutMinutes` | Process-tree kill after this duration (0 disables; may exceed 1440). Default 240. |
| `MaxRetries` / `RetryDelaySeconds` | Retry policy after Failed/TimedOut/Interrupted. Defaults 0 / 300. |
| `ContinueOnError` | When `false` and the job finally fails, dependents due at the same occurrence are `Skipped`. Default `true`. |
| `Schedule.Type` | `Daily` or `Weekly`. |
| `Schedule.Times` | One or several `"HH:mm"` values (multiple values cover the several-times-per-day case). |
| `Schedule.DaysOfWeek` | Weekly only: `["Sunday", ...]`. |
| `Schedule.MissedRunPolicy` | `RunOnce` (default) or `Skip`. |

The manifest is hot reloaded at every tick when its file changes; an invalid manifest is rejected with an error email and the last valid version stays in effect (no orchestrator restart needed to change the planning).

All shipped jobs are `Enabled=false` except `M365-VerifiedDomains-Inventory` (small, read-only, safe daily example). The template contains examples of each schedule type (Weekly, Daily once, Daily multi-time with `Skip`), one long job with `TimeoutMinutes` > 1440 (`Mailboxes-PermissionsByUser-Report`), and the `Exchange2016` group running with `PowerShellEdition = "WindowsPowerShell"`; pin those to the Exchange server through their `AllowedServers` list in the runtime manifest.

Full/fast pattern for reporting pipelines (Power BI): heavy inventories have a nightly full job (for example `EXO-Mailboxes-Inventory` with `-IncludeStats`, `Exchange2016-Local-Mailboxes-Inventory` with `-IncludeADPermission`) plus a midday `-Fast` job running the same script without the expensive switches. The `-Fast` job declares `DependsOn` on the full job so the two never overlap on the same CSVs, and uses `MissedRunPolicy = "Skip"` (a missed midday refresh has no value later). Before enabling a job, make sure its own runtime `.local.json` exists next to the target script (the child runs unattended; app-only auth must be configured).

## State file schema (`Orchestrator-State.json`)

```json
{
  "SchemaVersion": 1,
  "UpdatedUtc": "2026-07-11T06:00:00.000Z",
  "LastDailySummaryDate": "2026-07-11",
  "Jobs": {
    "EXO-Mailboxes-Inventory": {
      "LastScheduledOccurrence": "2026-07-11T01:00:00.0000000+02:00",
      "LastRunStart": "2026-07-11T01:00:05.0000000+02:00",
      "LastRunEnd": "2026-07-11T02:10:44.0000000+02:00",
      "LastStatus": "Success",
      "LastExitCode": 0,
      "RetryCount": 0,
      "Running": null,
      "PendingRetry": null
    }
  }
}
```

- `Running` (while a job is in progress): `Pid`, `StartTime`, `ScheduledOccurrence`, `LogPath`, `Attempt`, `TimeoutMinutes`. This is what re-adoption uses after a recycle/reboot/crash.
- `PendingRetry`: `NotBefore`, `Attempt`, `ScheduledOccurrence`.
- The file is written atomically (temp file + rename) after every mutation, so a crash never leaves a half-written state. An already executed occurrence is never relaunched; a missed occurrence follows `MissedRunPolicy`; a still-running job is re-adopted.

## Configuration (`SmartM365-Inventory-Orchestrator.local.json`)

Created from the committed template at first run. Keys follow the SmartM365 pattern: `__USE_GLOBAL__` inherits from `SmartM365.global.local.json`, and `{{DataAllRootPath}}`-style tokens are resolved through the tenant context.

Orchestrator-specific keys: `JobMailMode` (Always/OnError/Never), `SmtpPort`, `UseIntegratedAuth`, `UseSsl`, `RelayIp` (pin the SMTP endpoint IPv4), `SendDailySummaryEmail`, `DailySummaryTime`, `AllowedServers` (default server allowlist, empty = all), `MaxConcurrency`, `MaxLifetimeHours`, `TickSeconds`, `OrchestratorDataFolderPath`, `OrchestratorLogFolderPath`, `OrchestratorLogRetentionDays`, `JobLogRetentionDays`, `JobRunsCsvRetentionDays`.

## Parameters

| Parameter | Description |
| --- | --- |
| `-Tenant prod\|test` | Tenant profile key (default `test`). Also appended to every job command line. |
| `-Connect` | Passed through to every launched script. |
| `-DryRun` | Print the next-24h planning per job (plus pending catch-ups), launch nothing, exit. |
| `-Once` | Run a single tick then exit (tests). Launched children keep running detached. |
| `-Force <names>` | Launch the listed jobs immediately even if not due (still honors overlap and concurrency; bypasses schedule and dependency gate; advances the job occurrence pointer to now). |
| `-Only <names>` / `-Skip <names>` | Restrict/exclude jobs from launching. |
| `-MaxConcurrency <int>` / `-MaxLifetimeHours <int>` | Override the configured values. |
| `-JobsManifestPath <path>` / `-StatePath <path>` | Override default file locations. |

## Exit codes

| Code | Meaning |
| --- | --- |
| 0 | Normal end: lifetime recycle, `-DryRun`, `-Once`, or task stop handled cleanly. |
| 1 | Unexpected fatal error (fatal error email sent when mail is configured). |
| 2 | Configuration or jobs-manifest error at startup. |
| 3 | Another live orchestrator instance holds the lock. |

## Task Scheduler configuration

### Automated installation

The installer can be started from a standard PowerShell session and requests UAC elevation automatically. It securely prompts for the dedicated service-account password; the password is never accepted as a command-line parameter. `SYSTEM` and `LocalSystem` are explicitly refused because a privileged task must not launch repository files that could be modified by non-administrators.

For interactive setup, run or double-click the thin launcher below. The PowerShell installer itself requests UAC elevation, prompts for install/uninstall and prod/test, validates the service account, and asks whether the task should start immediately. The CMD file only invokes the PowerShell workflow.

All CMD launchers in this folder use `pushd` / `popd` so they can start from a UNC share. The installer launcher preserves the original UNC script path, tries the trusted PowerShell 7 executable under Program Files first, and falls back to Windows PowerShell only when PowerShell 7 is unavailable.

```text
.\SmartM365\SmartInventory\Orchestrator\Start-SmartM365-Inventory-OrchestratorScheduledTask-Installer.cmd
```

Running the PowerShell installer directly without parameters opens the same guided workflow. Supplying parameters keeps it suitable for repeatable administration and deployment automation.

```powershell
.\SmartM365\SmartInventory\Orchestrator\Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1 `
    -Tenant prod `
    -ServiceAccount 'CONTOSO\svc-smartm365' `
    -StartNow
```

The default registered task path is `\WCH\SmartM365 Inventory Orchestrator - <Tenant>`. Use `-TaskName` to override the task name. The installer creates the `WCH` Task Scheduler folder when needed, verifies administrator rights, PowerShell 7, the orchestrator, the jobs/config templates, and the SmartM365 tenant-context helper before registration. After successful registration, it removes an exact-name legacy copy from the Task Scheduler root; uninstall checks both `\WCH\` and the legacy root.

The installed task:

- directly runs `pwsh.exe -File SmartM365-Inventory-Orchestrator.ps1 -Tenant <prod|test> -Connect`;
- is stored in the `\WCH\` Task Scheduler folder;
- starts five minutes after server startup;
- starts daily at midnight and repeats every 30 minutes for one day;
- ignores a new start while an instance is already running;
- starts as soon as possible after a missed trigger;
- retries three times, one minute apart, after a failed task start;
- has no Task Scheduler execution time limit because the orchestrator manages its own lifetime.

To remove the task (no credential prompt):

```powershell
.\SmartM365\SmartInventory\Orchestrator\Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1 `
    -Tenant prod `
    -Uninstall
```

The service account must already have the local/domain rights and file/certificate access required by the enabled inventory jobs. Restrict write access to the repository checkout and orchestrator files to trusted administrators and the deployment process.

### Manual configuration

Create ONE task per tenant under the `\WCH\` Task Scheduler folder, running the launcher such as `Start-SmartM365-Inventory-Orchestrator-Prod.cmd`:

- General: dedicated service account with "Log on as a batch job", "Run whether user is logged on or not", "Run with highest privileges" if the inventory scripts need it. The account needs write access to the SmartM365 `Data` folders and the certificate/private key used by app-only auth.
- Triggers:
  - "At startup" (delay 5 minutes recommended).
  - Daily at a fixed time (for example 05:55). Recommended: set the daily trigger to "Repeat task every 30 minutes for a duration of 1 day". Combined with "Do not start a new instance", the repetition is ignored while the resident instance is alive and simply relaunches the orchestrator shortly after its 24h recycle exit, whatever time the previous instance started.
- Settings:
  - "Run task as soon as possible after a scheduled start is missed": enabled.
  - "If the task fails, restart every: 1 minute", up to 3 times (covers crashes; a normal recycle exits 0 and is restarted by the trigger repetition).
  - "If the running task does not end when requested, force it to stop": optional; stopping the orchestrator never kills running jobs (they are detached and re-adopted at the next start).
  - "Do not start a new instance" (the internal lock file also enforces this).
  - Disable "Stop the task if it runs longer than": the orchestrator bounds its own lifetime with `MaxLifetimeHours`.

## Testing before scheduling

One-line commands (run from the repository root; use `-Tenant prod` for production):

```text
pwsh -NoProfile -ExecutionPolicy Bypass -File "SmartM365\SmartInventory\Orchestrator\SmartM365-Inventory-Orchestrator.ps1" -Tenant test -DryRun
```

```text
pwsh -NoProfile -ExecutionPolicy Bypass -File "SmartM365\SmartInventory\Orchestrator\SmartM365-Inventory-Orchestrator.ps1" -Tenant test -Once
```

Optional: force one safe job through a single tick:

```text
pwsh -NoProfile -ExecutionPolicy Bypass -File "SmartM365\SmartInventory\Orchestrator\SmartM365-Inventory-Orchestrator.ps1" -Tenant test -Once -Force M365-VerifiedDomains-Inventory
```

## Troubleshooting

- **Is the loop alive?** Check `Orchestrator-Heartbeat.json` (timestamp must move every tick) and the daily orchestrator log.
- **Exit code 3 / "Another orchestrator instance is already running"**: a live instance holds `Orchestrator.lock`. If no `pwsh` orchestrator is actually running, the lock is stale and the next start recovers it automatically.
- **A job never starts**: check `Enabled`, the server allowlist (the startup log lists "Jobs not allowed on this server"; `-DryRun` shows the effective allowed list per job), the `-Only`/`-Skip` filters, the dependency chain (a parent pending retry blocks dependents), and the concurrency queue messages in the log.
- **Job marked `Interrupted`**: the child PID disappeared while the orchestrator was down (reboot, kill, crash). The retry policy applies; check the job log for partial output.
- **Job marked `Skipped`**: previous run still in progress at the new occurrence (overlap guard), missed occurrence with `MissedRunPolicy=Skip`, or a parent with `ContinueOnError=false` finally failed.
- **No emails**: verify `SmtpServer`, `From`, `To`/`ErrorMailTo` (global or local values), `JobMailMode`, and DNS/IPv4 reachability of the relay; pin `RelayIp` when DNS is unreliable.
- **Manifest changes ignored**: the file is reloaded only when its timestamp changes and it validates; an invalid manifest is rejected (error email) and the last valid version stays in effect.
- **Child exit code could not be read**: logged as a warning and treated as Success; check the job log to confirm the actual result.
