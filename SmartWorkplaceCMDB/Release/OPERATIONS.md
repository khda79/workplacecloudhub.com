# Smart Workplace CMDB V1 — operations

## Preconditions

- Use PowerShell 7 for normal execution; Windows PowerShell 5.1 remains covered
  by compatible offline suites.
- Keep tenant IDs, client IDs, certificate thumbprints, raw/curated CSVs, logs,
  Power BI caches and exports outside Git and the release package.
- Grant only the documented read permissions and review the exact tenant,
  profile, output root, collection mode and retention rules before `-Collect`.

## Safe validation

```powershell
pwsh -NoProfile -File .\SmartWorkplaceCMDB\Orchestration\SmartWorkplaceCMDB-Orchestrator.ps1 -Tenant example -Pipeline Full -ValidateOnly -NoConfigWrite
```

Validation does not authenticate or collect. Fixture runs remain isolated. A
live run requires explicit `-Collect`; bounded runs must retain their `Bounded`
coverage label.

Before orchestration starts, the shared preflight checks the PowerShell runtime,
project path, required modules, tenant credential fields, AD protocol support,
output write access and configured script-signature policy. `Fixture` skips all
external dependency and credential checks. `Logging.ScriptSignaturePolicy`
accepts `Disabled`, `Audit` or `Enforce`; `Audit` records a warning when a
signature cannot be established, while `Enforce` blocks a live collection.

## Evidence and failure handling

Every source snapshot has a sidecar describing tenant identity, mode, coverage,
row count, SHA-256 and timestamps. Downstream processing stops on failed or
in-progress state, tenant mismatch, hash/count drift or incompatible headers.
The previous valid snapshot is retained when a new collection fails.

Graph transient responses use bounded retries. `Retry-After` is honored within
the configured cap; retry exhaustion fails the source rather than publishing a
partial snapshot.

Entra user activity requires `AuditLog.Read.All` in addition to `User.Read.All`.
Verified-domain collection requires `Directory.Read.All`; hybrid coverage uses
only exact normalized keys from completed AD and Entra source snapshots and
publishes aggregate counts rather than directory identifiers.
Microsoft 365 license collection writes a separate service-plan catalog and a
compact effective user/service-plan fact. Missing activity or plan evidence
remains explicit; it is never converted into zero usage or compliance.

Intune operational inventory requires the read-only application permissions
`DeviceManagementServiceConfig.Read.All`, `DeviceManagementApps.Read.All`, and
`DeviceManagementConfiguration.Read.All`. Autopilot devices, detected
applications, configuration policies and Windows update policies are published
as independent source snapshots. Configuration and update policy endpoints use
Microsoft Graph beta and therefore require tenant qualification before
production scheduling.

Every source stages and validates its CSV set before promotion. A complete live
`Full` run can then send the aggregate collection summary configured under the
tenant-local `Notifications` object. Validate `From`, `To`, the selected Graph
or SMTP mode, and certificate access before scheduling the run. Summary mail is
disabled for fixtures, bounded runs, scoped AD runs and individual pipelines.
Its private history under `DATA-ALL\CollectionSummary` is required for previous,
J-7 and J-30 deltas; do not place it in Git. A delivery failure is reported as
`CompletedWithWarnings` and does not roll back valid collection outputs.
If a complete live `Full` run fails, the same notification channel sends a
separate operational failure alert with the failed step and available log and
transcript paths. That alert does not create or alter business snapshots.

## Concurrency and run state

Only one run for the same tenant and pipeline may write at a time. The
orchestrator creates an atomic guard before the first step and rejects a second
overlapping invocation. Its state file records the run identifier, current
step, completed-step count, start time, latest heartbeat time and final status.
State is stored under `LOG-ALL\Orchestration\State\<tenant>`; it contains no
collected business data. A stale guard is not silently removed: review the
recorded host, process and state before recovery.

## Logs and retention

Collection and fixture executions use the same bounded logging pattern as
SmartInventory. Validation remains read-only and creates no log files.

```text
LOG-ALL/
  Orchestration/
    Logs/SmartWorkplaceCMDB-Orchestrator_<host>_<timestamp>.log
    Runs/SmartWorkplaceCMDB-Orchestrator_<host>_<timestamp>.csv
  Jobs/
    <script-name>/<script-name>_<host>_<timestamp>_<sequence>.log
    <script-name>/<script-name>_<host>_<timestamp>_<sequence>.transcript.txt
```

Every operational text-log line is timestamped and classified; lifecycle
banners remain unprefixed. Each executed script also has a native PowerShell
transcript in the same job directory. The run CSV records status, duration,
error, dedicated log path and transcript path for every step. Defaults are 30
days and 30 files for orchestrator logs, 30 days and 30 files per script for
both step logs and transcripts, and 90 days and 90 files for run CSVs. Step logs
and transcripts are counted independently. Configure these safeguards under the
tenant-local `Logging` object. A value of `0` disables the corresponding age or
count rule. Retention failures are warnings in the orchestrator log and never
invalidate collected data.

The console always shows the WorkplaceCloudHub startup and completion banners.
Those banner lines are deliberately not timestamped; every operational message
between them is prefixed with local date and time. Child collector messages use
the same console convention. `-ValidateOnly` shows the console lifecycle but
does not create a transcript or any file under `LOG-ALL`.

## Power BI

The PBIP and ReportData stay private. Apply generators only to an authorized
PBIR folder after a targeted backup. Validate PBIR, reload the exact Desktop
PID, review every page, then save in Desktop. Screenshots validate rendering,
not native click, scrolling, cross-filter or drill-through behavior.

## Package and release

`Build/SmartWorkplaceCMDB-Package.ps1` uses an explicit allowlist, rejects
prerelease metadata, requires valid Authenticode signatures for included
PowerShell files and writes a hash manifest. Building a package does not publish
it. Git, tag, release, website and tenant actions require separate authority.
