# SmartM365 Inventory Orchestrator

`SmartM365-Inventory-Orchestrator.ps1` (v1.5.41) is a PowerShell 7 resident scheduler that runs the SmartInventory scripts (ActiveDirectoryInventory, ExchangeInventory, M365Inventory, IntuneInventory, ...) unattended.

## Cancel remaining pipeline jobs

Orchestrator **1.5.41**, Pipeline CLI **1.2.0** and GUI **1.3.2** support cooperative
cancellation. In **Requests**, select an active batch, enter a reason and choose
**Cancel remaining jobs**. Confirm the batch before publication. This is independent
of scheduling maintenance and does not publish the GUI configuration draft.

GUI **1.3.2** also provides **Cancel All remaining Jobs**, with a single reason and
confirmation listing every active pipeline batch. It includes active requests outside
the displayed recent history, not future requests, automatic schedules or running
collector processes. Each batch is validated before confirmation and revalidated
under the shared locks before cancellation. If a batch changes or a control becomes
unavailable during publication, the operation stops and reports the error; already
published cancellations remain audited. Selecting zero, one or several job rows no
longer unwraps a singleton into an invalid WPF ItemsSource.

The selected-request and all-request cancellation paths run shared-folder reads,
readiness checks, publication, persistent cancellation logs and the final Requests
refresh in an in-process background runspace (no ThreadJob module required). WPF
only applies results and displays the confirmation. A stage/elapsed-time label stays
visible before and after confirmation; slow storage still takes time but does not
block the window. Duplicate clicks and competing GUI control operations are disabled,
and automatic refresh is deferred until cancellation finishes. New batches submitted
after confirmation are not included. Partial publication and display-refresh errors
are reported separately; published cancellations are never rolled back. Closing the
GUI is refused while the operation is active: no worker, lock or collector is forcibly
stopped to impose an artificial GUI timeout. Other ordinary GUI refreshes are unchanged.

Alternatively, use the CLI against the **live shared Orchestrator folder**, never its
SharePoint/OneDrive mirror. Replace the generic paths and batch identifier below:

```powershell
$parameters = @{
    Tenant = 'test'
    SharedDataFolderPath = '<live-shared-Orchestrator-folder>'
    JobsManifestPath = '<live-shared-Orchestrator-folder>/Config/Orchestrator-Jobs.json.txt'
    Cancel = $true
    BatchId = '<active-batch-id>'
    Reason = 'Replace remaining pipeline work with controlled direct launches'
}
./SmartM365-Inventory-Pipeline.ps1 @parameters -ValidateOnly
# Only after successful read-only validation:
./SmartM365-Inventory-Pipeline.ps1 @parameters
```

- Every published expected resident must advertise cancellation protocol 1 in a fresh,
  Running heartbeat for the same tenant. Deploy the entire lot and let residents recycle
  safely; cancellation is refused during partial upgrades, startup or stale heartbeats.
  Readiness diagnostics distinguish an unsupported protocol, startup/recycling, stale
  or future heartbeat and a tenant mismatch. Do not restart a live collector to bypass
  this guard. A synchronized heartbeat copy is not proof of live control readiness.
- Pending jobs and scheduled pipeline retries become `Cancelled`. Already reserved or
  running jobs continue under ordinary supervision and record their real result. No
  collector process is killed; an already running collector is not cancelled remotely.
- A shared launch fence serializes reservation/process admission and cancellation.
  Cached queues cannot resurrect a cancelled job. Running failures do not start another
  pipeline retry after cancellation. Automatic schedules are not cancelled: keep
  scheduling maintenance enabled if you intend to run collectors directly.
- Successful/failed existing results, native CSVs, logs, leases and source receipts are
  preserved. The request itself records UTC time, actor, workstation and reason. Repeating
  cancellation resumes a partially completed publication without rewriting its audit.
- A batch remains `Cancelling` while a started job has not finished. It becomes `Cancelled`
  after completion, or `Failed` if there is a genuine failure (counts retain both failures
  and cancellations). New requests remain blocked while any job is nonterminal.
- The cancellation CLI returns 0 when the control operation succeeds, not as evidence of
  a successful collection. A waiting `-Collect` console exits 3 for a cancelled batch.
  Closing that console with Ctrl+C only stops waiting; it does not cancel shared work.

Control/job locks are excluded from the SharePoint mirror. Never remove a live lock or
edit/delete a request to bypass cancellation. An inaccessible control fails closed; a
crash-held lock needs operator diagnosis before recovery, not automatic deletion.

Offline checks (synthetic data only; real SMB/deployment qualification remains separate):

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File ./SmartM365/Tests/Test-SmartM365OrchestratorPipelineCancellation.ps1
pwsh -NoProfile -Sta -ExecutionPolicy Bypass -File ./SmartM365/Tests/Test-SmartM365OrchestratorGuiRequestsOffline.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File ./SmartM365/Tests/Test-SmartM365OrchestratorPipeline.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File ./SmartM365/Tests/Test-SmartM365OrchestratorMaintenanceOffline.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File ./SmartM365/SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator-GUI.ps1 -ValidateOnly
```

## Shared scheduling maintenance

Orchestrator 1.5.39 and GUI 1.2.0 add tenant-wide scheduling maintenance. In **Operations**,
enter a reason and select **Enable maintenance**; use **Disable maintenance** to resume.
Both actions require confirmation and optimistic revision checks. They publish only the
maintenance transition, never the jobs/cluster draft edited elsewhere in the GUI.

The authoritative state is `Config/Orchestrator-Maintenance.json.txt` under the shared
Orchestrator folder. GUI and residents must use that same live shared folder, not a
SharePoint/OneDrive mirror. SharePoint remains an operational mirror, not a control plane.
The state survives restart and has no automatic expiry. A persistent `.guard` coordinates
transition writes with process creation and detects disappearance of initialized state;
it is excluded from synchronization with other locks. Never delete the state or guard to
resume scheduling. Transitions are recorded in `Audit/Orchestrator_Maintenance.csv` without
creating configuration snapshots or modifying manifest hashes.

| Operation | Maintenance behavior |
| --- | --- |
| Running collector | Continues; timeout, completion, leases and logs remain supervised. |
| Automatic due launch / scheduled retry | Suspended; no new process. |
| New/pending Pipeline request and its retries | Allowed, preserving dependencies, ownership, claims and concurrency. |
| Explicit `-Force` | Still explicit; original ownership/overlap checks and dependency-bypass semantics are unchanged. Prefer Pipeline. |
| Heartbeat, peer health, mail, retention, SharePoint mirror | Continue; deliberately suspended schedules are not missing-start incidents. |
| Direct collector outside the orchestrator | Launch is not controlled; shared mail restrictions still apply. |

A pending manual request may supersede a suspended scheduled retry for the same job;
the original failure stays in run history. Maintenance never changes CSV schemas, invents
successful source receipts, weakens freshness checks, or changes prepared evidence contracts.
A manual child may still wait for an external dependency whose latest scheduled occurrence
was suspended: collect that dependency explicitly (or resume planning), rather than reuse
an inadmissible old export.

On resume, the durable UTC cutoff advances scheduling cursors past suspended automatic
occurrences and discards their pending automatic retries; original failures and running
process records are preserved. This applies after restart, missed intervals and ownership
changes as well. Next future schedules run normally: no catch-up burst. Scheduling cursors
are not success evidence. No run is recorded as successful for a suspended occurrence.

Activation requires **every published expected server** to have a fresh, Running heartbeat,
maintenance protocol 1, healthy control and acknowledgement of the current revision.
Update all residents before using the feature. The GUI shows **Applied**, **Pending**,
**Offline**, **Unsupported version** or **Control unavailable** and keeps an always-visible
banner; publication is not announced as globally applied until every acknowledgement is
present. Disabling remains possible if a server went offline, but its acknowledgement is
still pending. A missing, malformed or unreadable initialized control pauses new launches
(including manual ones) without abandoning job supervision. The launch-phase read and
shared gate protect activation during a scheduler tick; failure does not bypass the gate.

Offline validation (synthetic temporary data only):

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File ./SmartM365/Tests/Test-SmartM365OrchestratorMaintenanceOffline.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File ./SmartM365/SmartInventory/Orchestrator/SmartM365-Inventory-Orchestrator-GUI.ps1 -ValidateOnly
```

These tests do not qualify production SMB atomicity, real deployment, task recycling or
live multi-server operation. Deployment and a controlled production test remain separate
approval steps.

### Maintenance mail policy (1.5.40)

Deploy Core 1.0.68, Windows PowerShell 5 compatibility 1.0.49 and orchestrator 1.5.40
together. Shared Graph and SMTP helpers read live maintenance control immediately before
transport. While active, report recipients come only from the effective tenant/global
`To`, and error recipients from its `ErrorMailTo`; Cc and Bcc are removed. These are the
same values obtained by inheriting `__USE_GLOBAL__`, not literal recipient strings.
Script-local configuration is not rewritten. Normal routing returns when control is inactive.
Explicit `-MailPurpose Error` is authoritative; legacy calls also recognize an explicit
ErrorMailTo recipient or the collector's Boolean terminal `ScriptFailed` marker. Ambiguous
legacy calls without error evidence use global `To`, never script-local recipients.

Residents propagate the live shared path to their children. Direct scripts use
`OrchestratorSharedDataFolderPath` from effective global/tenant configuration, or default
to `DataAllRootPath/Orchestrator`. For a custom shared control location, configure that key
for direct launches too. Do not use a synchronized mirror as control. Missing, malformed
or inaccessible initialized control blocks sending, including SMTP fallback; a standalone
deployment with an accessible data parent and no control/cluster/guard remains inactive.

An enable/disable transition, from either the GUI or management API, is notified by a
resident using global `To` even after disable. Subject/body include the action, revision,
UTC timestamp, actor, server and reason. The existing published transition audit supplies
pending events; a single `Config/Orchestrator-Maintenance-Mail.json.txt` delivery cursor and
separate `.guard` serialize workers without holding the scheduling/GUI gate during mail.
No-op control changes do not send mail. A failed notification leaves control unchanged and
is retried after 60 seconds. First deployment initializes at the current revision without
replaying old audit mail. No resident running means notification waits for a resident.

The cursor prevents normal cross-server/restart duplicates, but an accepted mail followed
by a crash before checkpoint persistence can be resent; transport does not offer exactly-once
delivery. Revision gaps stop notification rather than silently skipping transitions.
Already running collectors with an old module keep the old code until their next process;
deploy/recycle all residents and let existing collectors finish before live qualification.

Additional offline checks (both use synthetic data and mocked transport only):

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File ./SmartM365/Tests/Test-SmartM365MaintenanceMailOffline.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ./SmartM365/Tests/Test-SmartM365MaintenanceMailOffline.ps1
```

Version 1.5.20 mirrors the shared `Config`, `Audit`, `Election` and `PipelineRuns` operational trees to SharePoint when uploads are enabled. It creates missing SharePoint folders, preserves the `DATA-ALL\Orchestrator` hierarchy, skips locks and temporary files, uploads only changed JSON/CSV artifacts under a cluster-wide mirror lock, and removes expired mirrored concurrency leases without treating SharePoint as an operational source. GUI v1.0.6 replaces the free-text weekly-day field with Monday-to-Sunday checkboxes, disables and clears them for `Daily`, and requires at least one selected day for `Weekly`. Version 1.5.19 adds an atomic, cluster-wide `Rebalance now` request from the GUI. A resident orchestrator consumes each unique request under the planner lock, recalculates elected owners without sticky ownership, and records the applied request in the shared plan. Version 1.5.18 treats `CompletedWithWarnings` as a terminal distributed occurrence result, preserves Teams capability-probe diagnostics when cleanup is unavailable, and retries transient pipeline status-file replacement failures. Version 1.5.17 ensures that collectors requiring explicit external-action opt-in receive `-EnableConfiguredExternalActions`, including an additive migration of existing central manifests that preserves operator arguments and scheduling overrides. Version 1.5.15 standardized every orchestrator mail subject as `[SMART 365] - [tenant] - [ Orchestrator] - ...` and guaranteed that the common mail footer identifies the script and its version. Version 1.5.14 added an atomic distributed pipeline request consumed by the existing
resident orchestrators. A `Full` request selects every enabled non-manual job from
the effective shared manifest, includes dependency closure, preserves election,
claims, locks, dependencies and concurrency, and reports one shared batch status.

It is started by a single Windows Task Scheduler task (at server startup plus a daily trigger), loops with a one-minute tick, launches each job exactly at its scheduled occurrences, and exits cleanly after a configurable maximum lifetime (default 24 hours) so Task Scheduler restarts a fresh instance (memory recycling). The orchestrator recycle never interrupts a running job (see "Detached jobs and re-adoption").

## Files

| File | Purpose |
| --- | --- |
| `SmartM365-Inventory-Orchestrator.ps1` | Orchestrator script (PowerShell 7). |
| `SmartM365.Orchestrator.Distributed.psm1` | Automatic capability probes, weighted election planner and atomic occurrence claims. |
| `SmartM365.Orchestrator.Management.psm1` | Shared configuration validation, atomic publication, versions, rollback, audit and multi-server history aggregation. |
| `SmartM365.Orchestrator.Maintenance.psm1` | Shared scheduling state, revision-checked transitions, server acknowledgements and launch gate. |
| `SmartM365.Orchestrator.Pipeline.psm1` | Full/group selection plus atomic shared pipeline request and per-job batch status management. |
| `SmartM365-Inventory-Pipeline.ps1` | Read-only pipeline validation or atomic `-Collect` submission, with wait-by-default aggregation. |
| `SmartM365.Orchestrator.Insights.psm1` | Read-only GUI data layer: live operations, job health, dependency readiness, recent pipeline requests and job run requests. Testable headless. |
| `SmartM365-Inventory-Orchestrator-GUI.ps1` | Central WPF console for live operations, planning, dependencies, run requests, server assignment, election visibility, history and configuration versions. |
| `Set-SmartM365-OrchestratorTimeoutPolicy.ps1` | Preview-first migration that applies 23-hour daily-once and 6-day weekly-once timeouts to the shared configuration, then publishes through the normal versioned/audited path with `-Execute`. |
| `Set-SmartM365-OrchestratorSchedulePlan.ps1` | Preview-first publication of the template planning fields (schedule, timeout, dependencies, concurrency key, dependency rule) into the shared configuration; removes `Exchange2016-Local-Mailboxes-Fast`; keeps arguments, Enabled flags (unless `-SyncEnabled` takes them from the template) and every other operational field; publishes through the versioned/audited path with `-Execute`. |
| `SmartM365-Inventory-Orchestrator.local.json.template` | Safe committed template; copied to `SmartM365-Inventory-Orchestrator.local.json` at first run (the runtime `.local.json` is Git-ignored). |
| `Orchestrator-Jobs.json.template` | Safe committed jobs-manifest template (all schedules, neutral `AllowedServers`). |
| `Orchestrator-Cluster.json.template` | Safe committed cluster template. The shared runtime file is bootstrapped from the first server's current local cluster settings. |
| `Orchestrator-Jobs.json` | Legacy/local bootstrap manifest, Git-ignored. With central configuration enabled, its current values seed the shared manifest only when the shared manifest does not exist yet. |
| `Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1` | Installs or removes the unattended Windows scheduled task under a dedicated service account. |
| `..\..\..\Install-WorkplaceCloudHub-CodeSigningCertificate.ps1` | Installs the committed public Authenticode certificate into `LocalMachine` trust stores by default; `CurrentUser` remains available explicitly. |
| `..\Launchers\Orchestrator\Start-SmartM365-Inventory-OrchestratorScheduledTask-Installer.cmd` | Interactive elevated launcher for scheduled-task installation or removal. |
| `..\Launchers\Orchestrator\Start-SmartM365-Inventory-Orchestrator.cmd` | Production launcher: `-Tenant prod -Connect`. |
| `..\Launchers\Orchestrator\Start-SmartM365-Inventory-Full-Collect.cmd` | Submits the prod `Full` pipeline and waits for its shared completion status. |
| `..\Launchers\Orchestrator\Start-SmartM365-Inventory-Orchestrator-GUI.cmd` | Opens the production WPF management console in an STA PowerShell 7 process. |
| `..\Launchers\Orchestrator\Send-SmartM365-Inventory-Orchestrator-ExecutionSummary.cmd` | Sends the prod all-server execution-summary email with a consolidated row per job plus separate detailed tables for the last 24 hours and 7 days. |

| `..\Launchers\Orchestrator\Stop-SmartM365-Inventory-Orchestrator.cmd` | Launcher: requests a clean stop for the running prod orchestrator instance. |

| `Restart-SmartM365-Inventory-Orchestrator.ps1` | Stops the orchestrator cleanly, then starts the existing scheduled task. |
| `..\Launchers\Orchestrator\Restart-SmartM365-Inventory-Orchestrator.cmd` | Launcher: clean restart of the prod scheduled task. |


Runtime files are tenant-isolated, created automatically and Git-ignored. State, job-run CSVs and logs use a per-server suffix (for example `{{DataAllRootPath}}\Orchestrator\SRV01`) to prevent collisions. The lifecycle CSV stays one level above the server folders so it provides a single tenant-wide history across all orchestrator servers:

| File | Location | Purpose |
| --- | --- | --- |
| `Config\Orchestrator-Jobs.json` | `{{DataAllRootPath}}\Orchestrator` | Shared operational job configuration. All servers hot reload it. |
| `Config\Orchestrator-Cluster.json` | `{{DataAllRootPath}}\Orchestrator` | Shared expected-server list, election weights, server policies and peer-monitoring settings. |
| `Config\Versions\<VersionId>` | `{{DataAllRootPath}}\Orchestrator` | Immutable before/after snapshots created for every publication and rollback. |
| `Audit\Orchestrator_ConfigChanges.csv` | `{{DataAllRootPath}}\Orchestrator` | User, source server, timestamp, hashes, version ID and change summary for every publication. |
| `SmartM365-Orchestrator-GUI\<Server>\SmartM365-Orchestrator-GUI_<yyyyMMdd>.log` | `{{LogAllRootPath}}` | Persistent GUI activity and technical error details. A local `%LOCALAPPDATA%\SmartM365\Logs` fallback is used only when the central log root is unavailable. |
| `Orchestrator_Runs.csv` | `{{DataAllRootPath}}\Orchestrator` | Tenant-wide lifecycle history: one row per orchestrator process, shared across servers and retained indefinitely. |
| `Orchestrator_Runs.lock` | `{{DataAllRootPath}}\Orchestrator` | Cross-process file lock serializing lifecycle CSV updates from all servers. |
| `Orchestrator-State.json` | `{{DataAllRootPath}}\Orchestrator\<Server>` | Per-job state (last occurrence, last run, running PID). Atomic writes. |
| `Orchestrator-Heartbeat.json` | `{{DataAllRootPath}}\Orchestrator\<Server>` | Rewritten at every tick: timestamp, PID, running/pending jobs, ready capabilities and active election plan. |
| `Orchestrator-Capabilities.json` | `{{DataAllRootPath}}\Orchestrator\<Server>` | Generated automatically at startup and periodically; contains only readiness evidence and Graph application role names, never tokens or secrets. |
| `Orchestrator-ElectionPlan.json` | `{{DataAllRootPath}}\Orchestrator\Election` | Shared weighted assignment plan generated under an atomic planner lock. |
| `Orchestrator-RebalanceRequest.json` | `{{DataAllRootPath}}\Orchestrator\Election` | Latest atomic GUI rebalance request. The shared plan records its `RequestId` after application, making the request idempotent across resident servers. |
| `<Job>\<Occurrence>.json` | `{{DataAllRootPath}}\Orchestrator\Election\Claims` | Atomic cross-server occurrence claim. Prevents duplicate launches and records owner/status through retries and restarts. |
| `<ConcurrencyKey>.json` | `{{DataAllRootPath}}\Orchestrator\Election\Concurrency` | Atomic cluster-wide lease. Serializes jobs sharing a `ConcurrencyKey`, including jobs elected on different servers. |
| `<BatchId>\request.json` and `Jobs\<JobName>.json` | `{{DataAllRootPath}}\Orchestrator\PipelineRuns` | Immutable pipeline request plus atomically updated per-job statuses shared by all resident servers. |
| `Orchestrator-StopRequested.json` | `{{DataAllRootPath}}\Orchestrator\<Server>` | Temporary manual stop request written by `-Stop`; consumed and removed by the resident instance. |
| `Orchestrator.lock` | `{{DataAllRootPath}}\Orchestrator\<Server>` | Global lock; prevents two instances for the same tenant. Stale locks (dead PID) are recovered with a warning. |
| `Orchestrator_JobRuns_<yyyyMMdd>.csv` | `{{DataAllRootPath}}\Orchestrator\<Server>\JobRuns` | Daily job-run tracking CSV (atomic writes). |
| `SmartM365-Inventory-Orchestrator_<Server>_<yyyyMMdd>.log` | `{{LogAllRootPath}}\SmartM365-Orchestrator\<Server>` | Orchestrator log, daily rotation. |
| `Job-<JobName>_<Server>_<timestamp>.log` | `{{LogAllRootPath}}\SmartM365-Orchestrator\<Server>\Jobs` | One log per job execution (stdout + stderr of the child process). Legacy files from old per-job subfolders such as `Jobs\AD-HealthCheck` are migrated into this flat `Jobs` folder on startup when they are not referenced by a running job. |

Because tenant contexts resolve separate data roots, `prod` and `test` lifecycle histories remain isolated. For `AssignmentMode = "Elected"`, all servers read one shared plan and an occurrence can be claimed atomically by only one owner. Empty `AllowedServers` therefore no longer means that every server launches an elected job; the legacy behavior remains only for old manifest entries without `AssignmentMode`.

When SharePoint upload is enabled, one resident server at a time mirrors `Config`, `Audit`, `Election` and `PipelineRuns` beneath the corresponding SharePoint `DATA-ALL\Orchestrator` path at `OrchestratorSharePointUploadIntervalMinutes`. Only JSON and CSV files are managed; `.lock`, `.tmp`, `.guard` and takeover-lock files are excluded. SharePoint is a read-only operational copy for administrators and is never read back by the scheduler. Durable configuration, audit, claim and pipeline history is not deleted remotely. A mirrored `Election\Concurrency` lease is removed after its local lease disappears so SharePoint does not show an expired active lock.

### Central management GUI

Launch `Start-SmartM365-Inventory-Orchestrator-GUI.cmd` from the Orchestrator launchers folder. The GUI:

- shows an `Operations` tab per server: heartbeat state (`Online`, `Starting`, `Recycling`, `Stale`), heartbeat age, version, running jobs with their duration, pending jobs with the orchestrator reason (dependency, `ConcurrencyKey`, peer claim), time until the planned recycle, active peer-monitoring incidents and orchestrator mails of the last 24 hours (double-click opens the mail copy);
- shows enabled jobs, next occurrences, current elected owners and server health, plus per-job health columns computed from 21 days of all-server runs: `Health` (`OK`, `Stale`, `Failing`, `NoRecentSuccess`, `Disabled`, `Manual`), last status, last success, its age, the expected maximum age (longest schedule gap + 2 h + average duration, the same rule as the `FreshSuccess` gate) and the average duration;
- edits frequency, times, weekly days, missed-run policy, enabled state, retry/timeout values and assignment mode;
- edits `DependsOn`, `DependencyMode` (`LatestOccurrence` or `FreshSuccess`) and `DependencyMaxAgeHours` (0 = automatic; a value is a floor), shows why the selected job waits (per dependency: rule, state, last success, age, maximum age) and lists the jobs that depend on it; disabling a job used by enabled jobs asks for confirmation, and an invalid dependency (unknown job, cycle, self) leaves the draft unchanged;
- submits `Request run` for the selected job through the orchestrators (same contract as `SmartM365-Inventory-Pipeline.ps1 -Job <Job> -Collect -NoWait`, optionally with its enabled dependencies) on the published configuration, and follows requests with per-job status, owner server and detail in the `Requests` tab; only one request can be active at a time;
- refreshes `Operations` and `Requests` every 60 seconds while `Auto-refresh` is checked; this never reloads the configuration draft;
- edits the expected server list, election weights and operational server policies;
- filters the all-server run history by date, server, job and status (the real statuses `Success`, `CompletedWithWarnings`, `Failed`, `TimedOut`, `Interrupted`, `Retried`), with a `Failures 24 h` shortcut, CSV/HTML export and log opening (a message explains when the log is not reachable from the GUI computer);
- validates the complete jobs and cluster documents before publication;
- sorts the Planning grid by job name by default, uses Monday-to-Sunday checkboxes for `Weekly` schedules, and enables the pinned-server selector only for `Pinned` assignment, where exactly one expected server is required;
- submits `Rebalance now` as an explicit atomic request against the active published configuration; the next resident tick recalculates only `Elected` owners from live capabilities, weights, policies and duration history;
- uses an atomic cross-server lock and hash comparison to reject concurrent/stale edits;
- replaces shared JSON through a bounded Windows/SMB-compatible retry, skips unchanged files and rolls back an earlier replacement if the second configuration file fails;
- creates before/after versions and an audit row for every successful publication or rollback;
- writes timestamped activity and full technical error details to the daily GUI log under `LogAllRootPath`.

The GUI never starts or stops a process. `Request run` writes an atomic pipeline request that the resident orchestrators execute on the elected owner with their normal claims, `ConcurrencyKey` leases and dependency rules; a request is refused when an elected job has no owner in the current plan or when another request is still active. `Rebalance now` changes only the shared election plan; it does not launch a job. `Elected` ownership is displayed from the generated plan and cannot be edited directly. To choose a fixed owner, set `AssignmentMode` to `Pinned` and select exactly one expected server. `Manual` jobs remain excluded from scheduled election.

### Full and group pipeline collection

Validate the exact effective plan without creating a request or launching a collector:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File ".\SmartM365\SmartInventory\Orchestrator\SmartM365-Inventory-Pipeline.ps1" -Tenant prod -Pipeline Full -ValidateOnly
```

Submit the same plan to the resident orchestrators and wait for every selected job:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File ".\SmartM365\SmartInventory\Orchestrator\SmartM365-Inventory-Pipeline.ps1" -Tenant prod -Pipeline Full -Collect
```

To run only named jobs (script 1.1.0, orchestrator 1.5.36):

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File ".\SmartM365\SmartInventory\Orchestrator\SmartM365-Inventory-Pipeline.ps1" -Tenant prod -Job WorkplaceEvidence-Prepare -Collect
```

`-Job` accepts one or more names (also comma-separated) and is exclusive with `-Pipeline`. Dependencies are not added unless `-IncludeDependencies` is given. For a dependency outside the request, the orchestrator applies the requested job's scheduled dependency rule (`FreshSuccess` with its maximum age, otherwise the latest scheduled occurrence of an elected dependency or the local state of a legacy one): the job waits, then ends `BlockedDependencyTimeout`, and a failed dependency ends it `BlockedDependencyFailed`. `SmartInventory/LaunchersByOrchestrator/` contains one generated launcher per job for this mode. Since 1.5.36, disabled or manual dependencies are ignored by every pipeline selection, as by the scheduled gate, and listed as ignored; before, they made `-Pipeline Full` fail (WorkplaceEvidence-Prepare depends on the disabled Microsoft 365 Backup collectors).
Since 1.5.37, a job status left `Starting` or `Running` after its final shared write failed is reconciled by the server that wrote it, once that server no longer supervises the job: from the terminal shared claim of the request occurrence (elected jobs), otherwise from its own job-run history for that exact occurrence, or a pending retry becomes `RetryScheduled`; without any recorded result the status becomes `Interrupted` only after the job timeout plus `ElectionClaimGraceMinutes`. Before, such a status kept the batch active forever and every later request was refused with "An active pipeline request already exists".
`Full`, `AD`, `Exchange`, `Exchange2016`, `M365` and `Intune` are supported. Selection is driven by the current shared manifest: disabled jobs and `AssignmentMode=Manual` jobs are excluded, while required enabled dependencies are added automatically. `-Collect` rejects a second active batch, requires a current elected owner for each elected job, and waits by default; `-NoWait` returns after the atomic request is published. Pipeline occurrences do not update `LastScheduledOccurrence`, so a manual full collection never shifts the normal schedule. Exit code `0` means success, `3` means completed with warnings, and `1` means failure, rejection or wait timeout.

`CentralConfigurationEnabled` defaults to `true`. Local `.local.json` files still own machine/runtime concerns such as tenant authentication, paths, mail transport, capability probing and concurrency. The shared files own job planning and cluster policy. Supplying the CLI `-JobsManifestPath` override disables central configuration for that invocation, which keeps isolated tests and diagnostics possible.

### Orchestrator SharePoint uploads and dependency wait logging

When SharePoint upload is enabled in the tenant configuration, the orchestrator mirrors stable operational artifacts to the configured SharePoint target folder, preserving the local `DATA-ALL` and `LOG-ALL` relative paths. Job logs are uploaded when the child process reaches a terminal state. Mail HTML copies are uploaded immediately after successful mail send. The resident orchestrator log, state, heartbeat, lifecycle CSV and daily job-run CSV are uploaded periodically according to `OrchestratorSharePointUploadIntervalMinutes` and once more during graceful shutdown or recycle.

The periodic operational-folder mirror compares each file's size and last-write time with the previous mirror state. Since v1.5.30, an unchanged JSON file without a legacy/preferred JSON twin is not re-read or re-parsed; changed files, twin pairs and forced mirrors are still fully validated. While the mirror runs, the shared heartbeat is refreshed at most once per minute so peers do not declare the server stale and the election does not drop its jobs.

### Occurrence claim retention

Occurrence claims under `Election\Claims` are coordination state only: duplicate-launch protection, the 8-day catch-up and dependency lookback, and peer monitoring. Run history is kept in the JobRuns and lifecycle CSV files. Since v1.5.31, once per hour and under a shared lock, the orchestrator removes claims whose scheduled occurrence (from the file name) is older than `ElectionClaimRetentionDays` (default 14; `0` disables; enabled values below 9 are raised to 9). Only terminal claims (`Success`, `CompletedWithWarnings`, `Failed`, `TimedOut`, `Interrupted`) whose `SafeUntilUtc` has passed are removed, together with the expired-claim archives and transport journal of the same occurrence; active, still-safe or unreadable claims are kept and counted in the log. The next operational mirror then deletes the matching SharePoint copies (Graph deletion to the site recycle bin). Both steps are bounded by `ElectionClaimRetentionMaxSecondsPerRun` (default 120 s) and refresh the heartbeat, so an initial backlog is drained over several hourly runs without blocking the scheduler.

Dependency waits are stateful: the orchestrator logs the first wait for a job, logs again only when the blocking dependency list changes, and emits a compact reminder according to `DependencyWaitLogIntervalMinutes`. It no longer writes the same dependency wait message on every scheduler tick. A separate proof-of-life log line is emitted according to `OrchestratorHeartbeatLogIntervalMinutes` (default 30 minutes) while the resident loop is alive.

### Distributed peer monitoring

Every resident server independently monitors the other servers listed in `ExpectedOrchestratorServers`. When that list is empty, the global `AllowedServers` list is used. The local server is excluded from its own peer checks.

The monitor first validates access to the shared orchestrator root. A storage-access failure produces one `MonitoringUnavailable` incident instead of falsely declaring every peer down. For each peer, it then checks the shared `Orchestrator-Heartbeat.json`; a missing, invalid or older-than-threshold heartbeat is confirmed across consecutive checks before an alert is sent. Since v1.5.32, the heartbeat age is measured at read time rather than at the start of the scheduler tick, so a slow tick no longer makes a peer heartbeat look older or "in the future"; a timestamp more than two minutes ahead still reports a real clock problem.

Since v1.5.34, the heartbeat publishes a `Lifecycle`: `Starting` (lock acquired, history scan, capability probe, election and job re-adoption; refreshed between these steps), `Running` (tick loop) or `Recycling` (written after a planned exit for a runtime update or `MaxLifetimeHours`). Peers do not report a `Recycling` heartbeat as stale for `PeerRecycleGraceMinutes` (default 15, never shorter than `PeerHeartbeatStaleMinutes`) and the election keeps its assignments during that window, so a normal restart no longer sends an alert and a recovery email. `Starting` and `Recycling` peers are not audited for `JobNotStarted`. A planned recycle also skips the operational mirror in its final SharePoint upload; the next instance runs it. A manual stop or a fatal error does not publish `Recycling` and is alerted as before.

Since v1.5.35, before requesting the shared `ConcurrencyKey` of an elected job, the scheduler reads the occurrence's shared claim. An occurrence already running on another server within its `SafeUntilUtc` is logged once as handled by that server instead of `BlockedByConcurrencyKey`; a finished claim (including `CompletedWithWarnings`) is recorded as handled by the peer. An expired claim still goes through the atomic claim and its safe takeover. The startup log now separates jobs added from the template and template settings added to existing jobs.

When the heartbeat is healthy and `PeerJobMonitoringEnabled` is true, the monitor reads the peer state and compares enabled jobs assigned to that server with their latest expected occurrence. Before raising `JobNotStarted`, it checks the occurrence's shared claim; terminal claims and non-expired `Claimed`, `Running` or `RetryScheduled` claims suppress the false alert even if the current plan owner differs from the server that handled the occurrence. Running jobs, pending retries, valid dependency waits, `BlockedByServerConcurrency` capacity waits and a valid `BlockedByConcurrencyKey` lease are not reported as missing. If the peer heartbeat briefly omits that pending item, the monitor verifies the shared concurrency lease directly before alerting. An expired or invalid concurrency lease is reported separately as `ConcurrencyBlockStale`. Disabled/manual jobs are excluded. Server-health alerts and job-schedule alerts are sent separately. Each stable server/job issue has its own reminder timestamp, so a changing queue does not resend every previously reported issue. A recovery email follows consecutive healthy checks. Since v1.5.33, a due occurrence newer than the expected one also covers it, so an hourly job waiting for dependencies during the launch grace is no longer reported as `JobNotStarted`; a wait longer than `DependencyWaitTimeoutMinutes` is still reported.

The heartbeat also publishes the health of the per-server state persistence. If a temporary SMB lock survives all configured retries, the resident process stays alive and continues supervision, heartbeat and peer monitoring, but pauses new launches until the state file can be written again. Healthy peers report this as a separate `StatePersistenceUnavailable` server-health issue.

Recommended production values for the three-server deployment:

```json
"PeerMonitoringEnabled": true,
"PeerJobMonitoringEnabled": true,
"ExpectedOrchestratorServers": ["CPPV-CAPTSE-001", "CPPV-CAPTSE-002", "CPPV-EXCSRV-113"],
"PeerMonitoringCheckIntervalSeconds": 60,
"PeerHeartbeatStaleMinutes": 5,
"PeerMonitoringConfirmationChecks": 2,
"PeerJobStartGraceMinutes": 15,
"PeerRecycleGraceMinutes": 15,
"PeerAlertReminderMinutes": 240,
"PeerAlertMailRetryMinutes": 15,
"PeerRecoveryEmailEnabled": true,
"MaxConcurrencyByServer": {
  "CPPV-CAPTSE-001": 6,
  "CPPV-CAPTSE-002": 9,
  "CPPV-EXCSRV-113": 6
},
"AtomicWriteRetrySeconds": 30
```

Each healthy server sends its own peer alert. Consequently, if one server is down in a three-server deployment, both surviving servers send an independently sourced alert.

### Authenticode validation

Authenticode validation is optional at engine level, but the shipped local template enables Audit mode with the repo public certificate. It is intended to complement ACL hardening, not replace it: the code folder must still be read-only for the orchestrator service account and ordinary users.

When `AuthenticodeValidationEnabled=true`, the orchestrator checks signatures before each job launch:

- the job `.ps1` file;
- `SmartM365.Core.psd1` and `SmartM365.Core.psm1` when `AuthenticodeCheckCoreModule=true`;
- the Windows PowerShell 5 compatibility module files for `PowerShellEdition = "WindowsPowerShell"` jobs when `AuthenticodeCheckWindowsPowerShellModule=true`.

`AuthenticodeValidationMode` supports:

- `Audit`: log warnings for unsigned, invalid or untrusted files, but still launch the job;
- `Enforce`: reject the launch when validation fails and send a critical orchestrator email.

`AuthenticodeAllowedThumbprints` may stay empty to trust any valid signer trusted by Windows. When populated, only those signer thumbprints are accepted. The orchestrator never signs files; signing is a deployment step after repo updates.

The public self-signed certificate is stored in the repo root under `Certificates`; this is safe because it does not contain the private key. When `AuthenticodeInstallTrustedCertificates=true`, the orchestrator imports only configured certificate files whose thumbprint is allowed into `CurrentUser\Root` and `CurrentUser\TrustedPublisher`, so scheduled-task service accounts can validate signatures without manual certificate-store setup.

Repository signing is handled by `Sign-WorkplaceCloudHubRepositoryPowerShellScripts.ps1`. Before signing, it normalizes PowerShell files to CRLF, validates `contact@workplacecloudhub.com`, and requests a DigiCert timestamp countersignature. The local Git hook path is `.githooks`; the `pre-push` hook runs the signing helper before a push and blocks the push when signatures changed, so the updated signatures must be committed first.

For the default machine-wide trust install, run from an elevated PowerShell session:

```powershell
& "C:\Program Files\PowerShell\7\pwsh.exe" -NoProfile -ExecutionPolicy Bypass -File ".\Install-WorkplaceCloudHub-CodeSigningCertificate.ps1"
```

For an explicit user-scoped install:

```powershell
& "C:\Program Files\PowerShell\7\pwsh.exe" -NoProfile -ExecutionPolicy Bypass -File ".\Install-WorkplaceCloudHub-CodeSigningCertificate.ps1" -StoreLocation CurrentUser
```

## Design

### Execution model

- One resident instance per tenant, bounded lifetime (`MaxLifetimeHours`, default 24), exit code 0 on recycle.
- 60-second tick: reload the manifest if it changed on disk, supervise running children, compute due occurrences, launch jobs, send the optional daily summary, rewrite the heartbeat, save state.
- Every job runs in its own detached child process (`-NoProfile -ExecutionPolicy Bypass`), which also isolates module/assembly conflicts between scripts (Graph SDK vs MSAL). Before launch, optional Authenticode validation can audit or enforce signatures on the job script and SmartM365 modules. The engine is `pwsh` by default; jobs with `PowerShellEdition = "WindowsPowerShell"` run in `powershell.exe` 5.1 instead (required by the Exchange on-premises scripts).

### Distributed assignment, capabilities and election

`AssignmentMode` controls ownership:

- `Elected`: the normal multi-server mode. The server is selected automatically from live candidates that satisfy every `RequiredCapabilities` and `RequiredGraphAppRoles` value.
- `Pinned`: legacy-style fixed ownership. Exactly one explicit `AllowedServers` value is required.
- `Manual`: never launched by the schedule, catch-up or retry loop. An explicit `-Force <JobName>` may launch it manually and still honors local overlap/concurrency guards.
- Missing `AssignmentMode`: compatibility mode for older manifests; the previous effective `AllowedServers` behavior is retained.

Each resident server generates `Orchestrator-Capabilities.json` itself. The default `ReadOnly` probe mode uses isolated, time-bounded child PowerShell processes and read-only checks for `Graph`, `EXO`, `AD`, `ExchangeOnPrem`, and `TeamsPowerShell`; `SharedRuntime` verifies write access to the shared election folder. Graph readiness also records the application roles returned by the app-only connection. No module is installed and no token, certificate private key or secret is written to the capability document.

The planner builds connected `DependsOn` components so parents and children always have the same owner. Daily load is schedule frequency multiplied by the median successful duration from the shared job-run history; `EstimatedDurationMinutes` is the fallback. It assigns the largest new or explicitly rebalanced components first to the least normalized load. An existing owner is sticky only while the complete live-capable server set remains unchanged and the preceding plan contains no unassigned group. A server joining or leaving the eligible set, an incomplete plan, a legacy plan, a jobs/cluster configuration publication, or a new `Orchestrator-RebalanceRequest.json` request therefore triggers one full weighted rebalance. Once the server set is stable and the latest request is acknowledged by `AppliedRebalanceRequestId`, the one-minute plan lease refresh keeps the same `PlanId` and does not move jobs merely because duration medians changed. `ElectionWeight` is the local default capacity, while `ElectionWeightsByServer` lets the same JSON be deployed everywhere with per-server weights; a weight of `1.10` gives a server about 10% more target capacity than `1.00`.
`ServerJobPolicies` applies an operational allow policy after technical capability detection. `OnlyJobsRequiring` means that every individual job in a dependency component must explicitly require every listed capability before that server can become a candidate. The policy is also published in the server capability document, so another planner respects the server's own restriction even when additional modules or permissions are detected. For example, the following reserves one server for Exchange Server on-premises workloads (Exchange 2016, 2019 or Subscription Edition) without allowing AD-only, Graph, EXO or Teams jobs:

```json
"ServerJobPolicies": {
  "CPPV-EXCSRV-113": {
    "OnlyJobsRequiring": ["ExchangeOnPrem"]
  }
}
```

Before launching an elected occurrence, the owner must create its shared claim with `FileMode.CreateNew`. If shared storage or the plan is unavailable, the elected job stays queued (fail closed). A different server may take over only after `TimeoutMinutes + ElectionClaimGraceMinutes` has expired and the original owner's heartbeat is stale. Terminal claims are never taken over. When another owner already completed the occurrence, the current owner synchronizes that terminal claim into its local state instead of retrying every minute; peer monitoring also accepts terminal or still-safe active claims as proof that the occurrence is handled.

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

- Global `MaxConcurrency` (default 2 in code, 6 in the production local configuration), optionally overridden for the current host through `MaxConcurrencyByServer`. Jobs due beyond the limit stay queued (their occurrence remains due, never lost), are published as `BlockedByServerConcurrency`, and start as soon as a slot frees. Re-adopted jobs count toward the limit. The production template sets `CPPV-CAPTSE-001=6`, `CPPV-CAPTSE-002=9` and `CPPV-EXCSRV-113=6`; this local map must be present on all three servers.
- `ConcurrencyKey` is enforced through an atomic lease in shared storage. Jobs sharing the key cannot run simultaneously anywhere in the cluster. A waiting occurrence is published as `BlockedByConcurrencyKey`; the lease follows detached/re-adopted jobs, is atomically aligned with the active timeout deadline after every recycle, and is released when the run ends, including before a retry delay.
- `DependsOn`: jobs due at the same occurrence run chained in topological order (the manifest is rejected at load time on a cycle). A dependent waits while a parent is running, due, or pending a retry. If a parent with `ContinueOnError=false` finally fails, dependents due at the same occurrence are marked `BlockedDependencyFailed`. If a dependency wait exceeds `DependencyWaitTimeoutMinutes`, the occurrence is marked `BlockedDependencyTimeout`.
- Overlap guards:
  - Resident lock: acquisition and stale recovery use an exclusive guard handle; uncertainty refuses acquisition with exit code 3. See the re-adoption and resident lock safety section.
  - Per job: a job still running at its next occurrence is not relaunched; the occurrence is marked `Skipped` with a warning in the log and the summary.

### Supervision, timeout, retries

- Each tick checks running children: exit code and duration are captured on exit; `TimeoutMinutes` triggers a full process-tree kill (`taskkill /T /F`). `TimeoutMinutes` may exceed 1440 for jobs longer than 24 hours. The committed production manifest uses 1380 minutes for a daily schedule with one execution time and 8640 minutes for a weekly schedule with exactly one day and one execution time. To migrate an existing shared configuration safely, first run `Set-SmartM365-OrchestratorTimeoutPolicy.ps1 -SharedDataFolderPath <OrchestratorRoot>` for preview, then repeat with `-Execute` to create the audited version. Startup and manifest reload log a `Timeout policy audit` line with counts and mismatches from the active central manifest.
- Retry policy per job: `MaxRetries` / `RetryDelaySeconds`. A failed run with retries left is recorded as `Retried`; the final failure sends the error email.

### Results and notifications

- Inventory/report scripts must return exit code `0` when data collection and report generation completed technically, even when the business report contains Warning or Critical findings. Those findings stay visible in CSV/HTML/email severity. Non-zero exit codes are reserved for technical failures, so the orchestrator can apply retries and failure alerts consistently to every job.
- Every execution is appended to the daily tracking CSV: `JobName, ScheduledTime, StartTime, EndTime, DurationSec, ExitCode, Status (Success/Failed/TimedOut/Skipped/Retried/Interrupted), RetryCount, LogPath`.
- Every orchestrator process is registered in the tenant-wide `Orchestrator_Runs.csv`: run ID, tenant, local/UTC start and end times, duration, server, Windows user, PID, script/PowerShell versions, mode, connection flag, status, exit code, stop reason and sanitized error. Rows start as `Running` and are finalized as `Success`, `Failed` or `Rejected`.
- On a later start, an unfinished `Running` row for the same server is changed to `Interrupted` when its recorded PID and start time no longer match a live PowerShell process. The shared CSV is rewritten atomically under `Orchestrator_Runs.lock`; it has no automatic retention.
- Emails use the shared `SmartM365.Core` mail helper (`Send-SmartM365Mail`), so the orchestrator follows the same `SendMailMode` behavior as inventory scripts: `Graph`, `SMTP`, or `Both`. With an empty `SmtpServer`, Graph is the default. HTML branding and saved mail copies are handled by the shared mail layer.
  - `JobMailMode`: `Always` (email for every final job completion), `OnError` (final failures only, default), `Never` (no job emails). This remains a dedicated notification policy key; `SendMailMode` controls only the transport.
  - Optional daily HTML summary (`SendDailySummaryEmail` + `DailySummaryTime`) with separate color-coded execution tables for the last 24 hours and 7 days (inline styles, no external CSS).
  - A fatal error email is sent if the orchestrator itself crashes (also on an invalid manifest at startup), independent of `JobMailMode`.
  - Mail is disabled only when required values for the selected transport are missing, for example `From`, recipient, Graph app auth for Graph mode, or `SmtpServer` for SMTP mode.

## Jobs manifest schema (`Orchestrator-Jobs.json`)

```json
{
  "Jobs": [
    {
      "Name": "EXO-Mailboxes-Inventory",
      "ScriptPath": "ExchangeInventory\\Mailboxes\\SmartM365-EXO-Mailboxes-Inventory.ps1",
      "Arguments": "-IncludeStats -ForceLiveStats",
      "Enabled": false,
      "Group": "Exchange",
      "DependsOn": [],
      "TimeoutMinutes": 480,
      "MaxRetries": 1,
      "RetryDelaySeconds": 600,
      "ContinueOnError": false,
      "AssignmentMode": "Elected",
      "AllowedServers": [],
      "RequiredCapabilities": ["SharedRuntime", "EXO", "Graph"],
      "RequiredGraphAppRoles": ["User.Read.All"],
      "EstimatedDurationMinutes": 60,
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
| `ScriptPath` | Script path relative to the SmartInventory root (the parent folder of `Orchestrator`). This remains the canonical inventory script used for validation and Authenticode checks. |
| `LauncherPath` | Optional launcher path relative to the SmartInventory root. When present, the orchestrator verifies `ScriptPath` but launches the `.cmd` instead. This is intended for AD/on-prem Exchange jobs that need local cache/unblock/bootstrap behavior. `{{Tenant}}` and `{{TenantKey}}` tokens resolve to the current tenant key. Centralized launchers default to `prod` for manual use; the orchestrator passes its own tenant through an isolated child-process environment variable. |
| `Arguments` | Extra arguments appended verbatim to the child command line. For direct `ScriptPath` launches, `-Tenant <tenant>` is appended by the orchestrator. `-Connect` is appended only when the target script declares that parameter; do not repeat either argument. For `LauncherPath` launches, the launcher owns tenant/connect handling, so the orchestrator does not append them. |
| `Enabled` | `false` by default; only enabled jobs are scheduled. |
| `Group` | Logical group (AD, Exchange, M365, Intune); informational. |
| `DependsOn` | List of job names that must complete first (cycle-checked at load). |
| `ConcurrencyKey` | Cluster-wide mutual-exclusion key. Defaults to the job name; jobs with the same explicit value are serialized across all servers. |
| `AssignmentMode` | `Elected`, `Pinned` or `Manual`. Missing keeps the pre-1.4 legacy allowlist behavior. |
| `AllowedServers` | Used by `Pinned` and legacy entries. `Pinned` requires exactly one server. Ignored for `Elected`. |
| `RequiredCapabilities` | Election gates: `SharedRuntime`, `Graph`, `EXO`, `AD`, `ExchangeOnPrem`, `TeamsPowerShell`. There is intentionally no standalone Intune capability; Intune jobs use Graph. |
| `RequiredGraphAppRoles` | Exact Microsoft Graph application roles required by the job. Valid only when `Graph` is required. |
| `EstimatedDurationMinutes` | Positive fallback used by the planner when successful duration history is unavailable. Default 5. |
| `PowerShellEdition` | `PowerShell7` (default, `pwsh`) or `WindowsPowerShell` (`powershell.exe` 5.1, required for Exchange on-premises scripts). |
| `TimeoutMinutes` | Process-tree kill after this duration (0 disables; may exceed 1440). Default 240. |
| `MaxRetries` / `RetryDelaySeconds` | Retry policy after Failed/TimedOut/Interrupted. Defaults 0 / 300. |
| `ContinueOnError` | When `false` and the job finally fails, dependents due at the same occurrence are marked `BlockedDependencyFailed`. Default `true`. |
| `DependencyWaitTimeoutMinutes` | Optional per-job maximum dependency wait. `0` inherits the orchestrator default. When exceeded, the occurrence is marked `BlockedDependencyTimeout`. |
| `DependencyMode` | `LatestOccurrence` (default): each dependency's latest scheduled occurrence must have succeeded and must not still be running. `FreshSuccess` (v1.5.33): a dependency is satisfied when its most recent `Success` or `CompletedWithWarnings` claim finished within `DependencyMaxAgeHours`, even if a newer occurrence is running or failed; without a fresh success the dependent waits (stale data is never accepted). Used by `WorkplaceEvidence-Prepare`. |
| `DependencyMaxAgeHours` | `FreshSuccess` only. `0` (default) = automatic: longest gap between two scheduled occurrences of the dependency plus 2 hours (about 3 h hourly, 26 h once daily, 170 h once weekly). A positive value is a floor (v1.5.34): each dependency uses the larger of the automatic age and this value, so a weekly dependency is never made stale by it. `WorkplaceEvidence-Prepare` uses `48` and depends on `EXO-Mailboxes-Inventory-Fast` rather than the LiveStats job, which runs 15-20 hours and does not succeed every day; mailbox and archive sizes then come from the latest LiveStats statistics snapshot reused by the Fast run. Manual runs outside the scheduler do not create claims and are not counted. |
| `Schedule.Type` | `Daily` or `Weekly`. |
| `Schedule.Times` | One or several `"HH:mm"` values (multiple values cover the several-times-per-day case). |
| `Schedule.DaysOfWeek` | Weekly only: `["Sunday", ...]`. |
| `Schedule.MissedRunPolicy` | `RunOnce` (default) or `Skip`. |

The manifest is hot reloaded at every tick when its file changes; an invalid manifest is rejected with an error email and the last valid version stays in effect (no orchestrator restart needed to change the planning). At startup, jobs newly introduced by the committed template are appended to an existing runtime manifest by name. Existing job definitions are never overwritten, so local Enabled flags, schedules and other operational overrides remain authoritative.

The committed template mirrors the validated operational Enabled flags and schedules. Scheduled jobs use `AssignmentMode = "Elected"`; Exchange on-premises jobs require `AD` plus `ExchangeOnPrem`, so they naturally elect only a server that passes both probes. `M365-PowerBIFabricActivity-Inventory` remains disabled with `AssignmentMode = "Manual"` as requested.

Template schedule (2026-09-30 redesign):

- Night chain, ready before 06:30: sources from 00:00 (AD, users, Entra and Intune devices, domains, licences, Exchange 2016 mailboxes, Intune device reports), then derived reports chained by `DependsOn` with the `LatestOccurrence` rule (EXO Fast 01:00 after AD and users; EXO permissions 02:30 after EXO Fast; Windows Update 01:00 after upgrade eligibility; Windows 11 readiness 01:30; hybrid identity and permissions-by-user 05:00), security and posture collectors staggered from 03:00, usage reports from 04:50, `WorkplaceEvidence-Prepare` at 06:30 (`FreshSuccess`, 48 h) and the quarantine report at 07:30.
- `LatestOccurrence` dependencies (1.5.38): when the latest scheduled occurrence of an elected dependency has no claim because the overlap guard skipped it, the run that was in progress covers it if its claim was created before the occurrence time and succeeded after it. Before, dependents waited for the next occurrence (for example AD at 08:00 after a skipped 00:00).
- Daytime: AD, users, Entra and Intune devices at 08:00, 11:00, 14:00 and 17:00; hybrid identity and Windows 11 readiness 45 minutes later (`MissedRunPolicy = Skip`). Sync health stays hourly and migration jobs run at 06:00, 12:00 and 18:00.
- `EXO-Mailboxes-Inventory` (full, with statistics, 18 to 22 hours) runs every 48 hours: Monday, Wednesday and Friday at 20:00, timeout 36 hours. `EXO-Mailboxes-Inventory-Fast` runs at 01:00 on Sunday, Monday, Wednesday and Friday only, when no full run holds the shared `EXOMailboxes` key. `EXO-Mailboxes-CalendarPermissions` (7 to 8 hours) also runs every 48 hours, alternating with the full run: Tuesday, Thursday and Saturday at 21:00, after the previous full run has ended. `Exchange2016-Mailboxes-CalendarPermissions` (about 5 minutes) runs every 48 hours too: Monday, Wednesday and Friday at 03:45, behind the other Exchange 2016 jobs.
- The Exchange 2016 jobs share the `ExchangeOnPremServer` concurrency key, so they run one after another on the Exchange server without coupling their results. The former `Exchange2016-Local-Mailboxes-Fast` job ran the same script with the same arguments 30 minutes after the full job and was removed.
- Weekly on Sunday: Intune discovered apps (00:00), Teams, SharePoint Online and Autopilot (04:00 to 04:20).
- Timeouts follow the timeout policy (1380 minutes for one daily time, 8640 for one weekly day and time); jobs with several daily occurrences use a timeout below their shortest interval.

The full job can declare `DependsOn` on its fast prerequisite when a quick CSV must be published first; dependent jobs are marked `BlockedDependencyFailed` or `BlockedDependencyTimeout` instead of silently waiting forever. Fast jobs should use `ContinueOnError=false` when downstream jobs depend on their outputs. Before enabling a job, make sure its own runtime `.local.json` exists next to the target script (the child runs unattended; app-only auth must be configured).

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

- `Running` (while a job is in progress): `Pid`, `StartTime`, `ScheduledOccurrence`, `LogPath`, `Attempt`, `TimeoutMinutes`, `ClaimPath`, `ConcurrencyLeasePath`, `ConcurrencyLeaseId`. This is what re-adoption uses after a recycle/reboot/crash. Optional `TimeoutRequested=true` records that termination was requested; absent means false for older state records.
- `PendingRetry`: `NotBefore`, `Attempt`, `ScheduledOccurrence`.
- The file is written atomically (unique temp file + SMB-compatible forced rename) after every mutation. Rename collisions are retried with exponential backoff and jitter for `AtomicWriteRetrySeconds` (default 30). If retries are exhausted, the process remains resident and pauses new launches until persistence recovers. An already executed occurrence is never relaunched; a missed occurrence follows `MissedRunPolicy`; a still-running job is re-adopted.

## Timeout supervision and recovery

A timeout requests process-tree termination; it does not itself confirm exit.
Until the supervised process is confirmed exited, the job remains in `Running`,
continues to occupy its concurrency slot, and keeps blocking the same job and
other jobs with the same `ConcurrencyKey`. No terminal job-run row, lease release
or failure retry is produced during this pending termination. Process inspection
or wait failures also retain supervision. The kill request is retried on later
ticks when the process can be inspected.

Timeout intent is saved before termination. If state persistence is unavailable,
termination is deferred until persistence recovers. The existing concurrency
lease is refreshed beyond the current supervision time by at least
`ElectionClaimGraceMinutes` (minimum one minute), even after the original deadline.
Re-adoption uses a future lease deadline and restores timeout intent. A later
exit code zero, or an increased configured timeout, cannot convert a previously
requested timeout to success. Once exit is confirmed, existing `TimedOut`, retry,
CSV, dependency and notification rules apply. A missing process on restart with
saved timeout intent is finalized as `TimedOut`; older records retain the existing
`Interrupted` behavior.

This protection requires accessible state/lease storage and continued supervision.
It does not fence a partitioned or stopped server, prove that a launcher left no
detached descendants, or qualify actual Windows process permissions and SMB
behavior. Initial process identity uncertainty is handled by the follow-up correction
described below. Focused synthetic regression evidence is retained outside the
public repository. No new tenant permission or configuration key is needed.


## Re-adoption and resident lock safety

Version 1.5.12 distinguishes a confirmed missing/mismatching PID from an
inaccessible process identity. Failed PID lookup, unavailable process name or
start time leaves the saved Running record intact. The job still occupies its
local concurrency slot, blocks overlapping work and refreshes its existing
concurrency lease when available. Identity inspection is retried on later ticks;
no process kill, final result or retry is triggered while identity is uncertain.
The heartbeat retains the recorded PID. Once inspection succeeds, normal
supervision resumes; confirmed disappearance keeps the existing Interrupted or
previously requested TimedOut outcome.

Resident lock acquisition uses an additional Orchestrator.lock.guard file with an
exclusive file handle held for the resident lifetime. The original JSON lock
remains readable, with the same fields; its open handle prevents replacement
while owned. A contender cannot recover a stale lock while another current
instance is acquiring it. Normal exit removes the JSON payload while still
holding the guard, then closes the guard handle. The empty guard file persists
and must not be deleted as routine stale-file cleanup: its open handle, not its
presence or age, determines ownership. Process termination releases OS handles.

Malformed lock contents or an inaccessible incumbent are left unchanged and
acquisition is refused (existing exit code 3). Inspect ownership and storage
permissions before repairing corrupt state; never remove a live lock to bypass
the guard. A valid legacy lock is still checked before acquisition. The guarantee
requires consistent deployment and filesystem handle-sharing semantics; SMB
disconnects, mixed versions and lost handles are not production-qualified.
No tenant API permissions or CSV schemas change.

## Configuration (`SmartM365-Inventory-Orchestrator.local.json`)

Created from the committed template at first run. Keys follow the SmartM365 pattern: `__USE_GLOBAL__` inherits from `SmartM365.global.local.json`, and `{{DataAllRootPath}}`-style tokens are resolved through the tenant context.

Orchestrator-specific keys: `JobMailMode` (Always/OnError/Never), `SendMailMode` (Graph/SMTP/Both, inherits global by default), `SmtpPort`, `UseIntegratedAuth`, `UseSsl`, `RelayIp` (pin the SMTP endpoint IPv4), `SendDailySummaryEmail`, `DailySummaryTime`, `AllowedServers` (legacy/default allowlist), `DistributedSchedulingEnabled`, `CapabilityProbeMode`, `CapabilityProbeTimeoutSeconds`, `CapabilityRefreshMinutes`, `CapabilityMaxAgeMinutes`, `ElectionPlanRefreshSeconds`, `ElectionClaimGraceMinutes`, `ElectionHistoryDays`, `ElectionClaimRetentionDays`, `ElectionClaimRetentionMaxSecondsPerRun`, `ElectionWeight`, `ElectionWeightsByServer`, `ServerJobPolicies`, `ExchangeOnlineOrganization`, `MaxConcurrency`, `MaxConcurrencyByServer`, `MaxLifetimeHours`, `AtomicWriteRetrySeconds`, `TickSeconds`, `AutoRecycleOnRuntimeUpdate`, `RuntimeUpdateCheckIntervalSeconds`, `RuntimeUpdateStableChecks`, `RuntimeUpdateCooldownMinutes`, `MonitorCoreModuleVersion`, `SendRuntimeUpdateEmail`, `DependencyWaitLogIntervalMinutes`, `DependencyWaitTimeoutMinutes`, `OrchestratorRunsCsvLockTimeoutSeconds`, `OrchestratorHeartbeatLogIntervalMinutes`, `OrchestratorSharePointUploadIntervalMinutes`, `AuthenticodeValidationEnabled`, `AuthenticodeValidationMode`, `AuthenticodeAllowedThumbprints`, `AuthenticodeCheckCoreModule`, `AuthenticodeCheckWindowsPowerShellModule`, `OrchestratorDataFolderPath`, `OrchestratorLogFolderPath`, `OrchestratorLogRetentionDays`, `JobLogRetentionDays`, `JobRunsCsvRetentionDays`.

## Parameters

| Parameter | Description |
| --- | --- |
| `-Tenant prod\|test` | Tenant profile key (default `test`). Also appended to every job command line. |
| `-Connect` | Passed through only to direct target scripts that declare a `Connect` parameter; ignored for unsupported scripts and launcher-based jobs. |
| `-DryRun` | Print the next-24h planning per job (plus pending catch-ups), launch nothing, exit. |
| `-Once` | Run a single tick then exit (tests). Launched children keep running detached. |
| `-Force <names>` | Launch the listed jobs immediately even if not due. This is the only scheduler entry point for `Manual` jobs. It still honors ownership for elected jobs, atomic claims, overlap and concurrency; it bypasses schedule and dependency gates. |
| `-Only <names>` / `-Skip <names>` | Restrict/exclude jobs from launching. |
| `-MaxConcurrency <int>` / `-MaxLifetimeHours <int>` | Override the configured values. |
| `-JobsManifestPath <path>` / `-StatePath <path>` | Override default file locations. |
| `-Stop` | Writes a manual stop request for the current tenant and waits for the live instance to exit cleanly. Does not kill detached running jobs. |
| `-StopTimeoutSeconds <int>` | Maximum wait for `-Stop` before returning exit code 1. Default 180 seconds. |
| `-SendExecutionSummary` | Sends an all-server consolidated execution overview followed by the 24-hour and 7-day detailed tables, then exits without acquiring the resident lock or launching inventory jobs. |

## Exit codes

| Code | Meaning |
| --- | --- |
| 0 | Normal end: lifetime recycle, `-DryRun`, `-Once`, task stop, or execution summary sent successfully. |
| 1 | Unexpected fatal error, or manual execution-summary email could not be sent. |
| 2 | Configuration or jobs-manifest error at startup. |
| 3 | Resident lock held, inaccessible or not safely recoverable. |

## Task Scheduler configuration

### Automated installation

The installer can be started from a standard PowerShell session and requests UAC elevation automatically. It securely prompts for the dedicated service-account password; the password is never accepted as a command-line parameter. `SYSTEM` and `LocalSystem` are explicitly refused because a privileged task must not launch repository files that could be modified by non-administrators.

For interactive setup, run or double-click the thin launcher below. The PowerShell installer itself requests UAC elevation, prompts for install/uninstall and prod/test, validates the service account, and asks whether the task should start immediately. The CMD file only invokes the PowerShell workflow.

All CMD launchers under `..\Launchers` use `pushd` / `popd` so they can start from a UNC share. The installer launcher preserves the original UNC script path, tries the trusted PowerShell 7 executable under Program Files first, and falls back to Windows PowerShell only when PowerShell 7 is unavailable.

```text
.\SmartM365\SmartInventory\Launchers\Orchestrator\Start-SmartM365-Inventory-OrchestratorScheduledTask-Installer.cmd
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
- starts daily at midnight and repeats every five minutes for one day;
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

Create ONE task per tenant under the `\WCH\` Task Scheduler folder. Configure it to run `pwsh.exe -File SmartM365-Inventory-Orchestrator.ps1 -Tenant <prod|test> -Connect`, matching the action created by the installer:

- General: dedicated service account with "Log on as a batch job", "Run whether user is logged on or not", "Run with highest privileges" if the inventory scripts need it. The account needs write access to the SmartM365 `Data` folders and the certificate/private key used by app-only auth.
- Triggers:
  - "At startup" (delay 5 minutes recommended).
  - Daily at a fixed time (for example 05:55). Recommended: set the daily trigger to "Repeat task every 5 minutes for a duration of 1 day". Combined with "Do not start a new instance", the repetition is ignored while the resident instance is alive and simply relaunches the orchestrator shortly after its 24h recycle exit, whatever time the previous instance started.
- Settings:
  - "Run task as soon as possible after a scheduled start is missed": enabled.
  - "If the task fails, restart every: 1 minute", up to 3 times (covers crashes; a normal recycle exits 0 and is restarted by the trigger repetition).
  - "If the running task does not end when requested, force it to stop": optional; stopping the orchestrator never kills running jobs (they are detached and re-adopted at the next start).
  - "Do not start a new instance" (the internal lock file also enforces this).
  - Disable "Stop the task if it runs longer than": the orchestrator bounds its own lifetime with `MaxLifetimeHours`.

## Automatic runtime update recycle

With `AutoRecycleOnRuntimeUpdate=true`, the resident process checks its own script version and, when `MonitorCoreModuleVersion=true`, the `SmartM365.Core` manifest version. The default check interval is 60 seconds. A higher semantic version must remain byte-for-byte stable for two consecutive checks and pass PowerShell parser plus Authenticode signer validation before a recycle is accepted. When `SendRuntimeUpdateEmail=true` (default), the server that accepts the candidate sends one HTML notification with the old/new versions, validation result, server, tenant and recycle action.

A valid update stops the launch phase immediately, saves state, finalizes lifecycle tracking and SharePoint uploads, releases the lock, and exits with code 0 and `StopReason=RuntimeUpdate`. Detached inventory jobs keep running and are re-adopted by the next instance. The five-minute repeating Task Scheduler trigger starts the new version within five minutes; `IgnoreNew` prevents a concurrent instance while the resident process is healthy.

A lower version, changed content without a version increase, parser failure, or invalid/unapproved signature is logged as a warning and does not recycle the process. `RuntimeUpdateCooldownMinutes` throttles repeated warnings for the same rejected candidate. Existing scheduled tasks must be re-registered once with installer v1.1.4 or updated manually to use the five-minute repetition interval.

## Clean stop and restart

To make the scheduled task pick up a newly deployed orchestrator version, request a clean stop instead of killing `pwsh.exe` directly:

```text
.\SmartM365\SmartInventory\Launchers\Orchestrator\Stop-SmartM365-Inventory-Orchestrator.cmd
```

The launcher calls `SmartM365-Inventory-Orchestrator.ps1 -Tenant prod -Stop`. The running instance consumes `Orchestrator-StopRequested.json` on its next tick, stops launching new jobs, saves state, releases the lock, finalizes lifecycle tracking and performs the final SharePoint upload. Detached inventory jobs are not killed; the next orchestrator instance re-adopts them from state.

If no live orchestrator lock is found, the stop command removes any stale stop request, checks for orphaned orchestrator PowerShell processes, and stops the scheduled task when Task Scheduler is still running without a valid orchestrator lock. This clears stuck `Running` task instances that would otherwise refuse the next scheduled start.

To stop the resident instance and immediately start the registered scheduled task again, use:

```text
.\SmartM365\SmartInventory\Launchers\Orchestrator\Restart-SmartM365-Inventory-Orchestrator.cmd
```

The restart launcher does not start the orchestrator directly. It first calls the same clean stop workflow, verifies that the scheduled task is no longer running, then calls `Start-ScheduledTask` for `\WCH\SmartM365 Inventory Orchestrator - prod`. This keeps the restart under the registered service account, task folder, triggers and task security settings.

For the three-server production cluster, deploy the script and the same local capacity map everywhere, then restart one server at a time. Wait until its heartbeat is fresh and the startup log shows the expected `MaxConcurrency`, `atomicWriteRetrySeconds` and zero timeout-policy mismatches before restarting the next server.

## Manual execution-summary email

Run the dedicated launcher to send an immediate prod summary without stopping or duplicating the resident orchestrator:

```text
.\SmartM365\SmartInventory\Launchers\Orchestrator\Send-SmartM365-Inventory-Orchestrator-ExecutionSummary.cmd
```

The launcher calls `SmartM365-Inventory-Orchestrator.ps1 -Tenant prod -SendExecutionSummary`. This one-shot mode enumerates every `<Server>\JobRuns` folder under the tenant's shared orchestrator data root and reads the per-server `Orchestrator_JobRuns_<yyyyMMdd>.csv` files. It sends a consolidated table with one row per job observed across all servers during the last 7 days and a `Server(s)` column, then detailed 24-hour and 7-day tables with a `Server` column on every execution. Counts include retries and skipped attempts across all servers. It does not load the jobs manifest, acquire the resident lock, launch jobs, stop jobs, install Authenticode trust certificates, or change `LastDailySummaryDate`.

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

- **Is the loop alive?** Check `Orchestrator-Heartbeat.json` (timestamp must move every tick) and the daily orchestrator log. The log also emits `Heartbeat: alive; ...` every `OrchestratorHeartbeatLogIntervalMinutes` minutes by default.
- **Exit code 3**: another instance holds the resident lock, or its ownership/storage cannot be inspected safely. Confirmed stale valid payloads can be recovered; malformed or inaccessible locks are preserved for investigation. Never delete a live guard file.
- **A job never starts**: check `Enabled`, the server allowlist (the startup log lists "Jobs not allowed on this server"; `-DryRun` shows the effective allowed list per job), the `-Only`/`-Skip` filters, the dependency chain (a parent pending retry blocks dependents), and the concurrency queue messages in the log.
- **Job marked `Interrupted`**: the child PID disappeared while the orchestrator was down (reboot, kill, crash). The retry policy applies; check the job log for partial output.
- **Job marked `Skipped`**: previous run still in progress at the new occurrence (overlap guard), or missed occurrence with `MissedRunPolicy=Skip`.
- **Job marked `BlockedDependencyFailed` or `BlockedDependencyTimeout`**: a required dependency failed with `ContinueOnError=false`, or dependencies kept blocking longer than the configured wait timeout.
- **No emails**: verify `SmtpServer`, `From`, `To`/`ErrorMailTo` (global or local values), `JobMailMode`, and DNS/IPv4 reachability of the relay; pin `RelayIp` when DNS is unreliable.
- **Manifest changes ignored**: the file is reloaded only when its timestamp changes and it validates; an invalid manifest is rejected (error email) and the last valid version stays in effect.
- **Child exit code could not be read**: logged as an error and treated as Failed; check the job log and process history before retrying.
