# SmartInventory audit - lot 2: timeout supervision

Audited and approved for repository publication on 2026-09-10 from commit
`63f6d93c8c606b5e13844d56d10fb3cc1d1f3cb0`. The approved scope contains five files. No operational deployment or
website publication is included.

## Demonstrated defects and corrections

| Reproduction | Previous behavior | Corrected behavior |
| --- | --- | --- |
| Mock process survives timeout kill | Complete-JobRun clears Running, releases the lease and queues a retry; same-job forced launch and shared-key sibling can start | Keep Running, concurrency slot and lease until confirmed exit; retry termination on later ticks |
| Mock Refresh/HasExited inspection throws, including after kill | Inspection exception is treated as exit | Keep supervision, refresh the lease and log uncertainty |
| Active process exceeds its deadline | Lease refresh returns early; re-adoption can align the lease to a past date | Retain the original timeout deadline, but renew lease protection to at least current supervision time plus configured grace |
| Restart or hot reload while termination remains pending | No durable indication of the earlier timeout request | Optional Running.TimeoutRequested flag survives re-adoption and prevents later exit zero or configuration changes from erasing the timeout |
| State persistence fails before termination | No durable timeout intent | Defer termination while persistence is unavailable; retain supervision and retry after recovery |

The main orchestrator changes from 1.5.10 to 1.5.11 in its header and runtime
version. No election weights, assignment algorithm, schedule, manifest, retry
count/delay, collector business algorithm or operational configuration is changed.
The distributed module remains 1.1.4, Core 1.0.48 and compatibility 1.0.36.
There is no SmartInventory suite version or stable qualification.

## Tests and evidence

`Tests/Test-SmartM365OrchestratorLifecycleOffline.ps1` parses the source and loads
only named function definitions into an isolated module. It never invokes the
orchestrator entry point, tenant loader, capability probe, real process launcher,
taskkill, scheduled-task API, mail or upload. Process methods and operational
side effects are mocks. One case uses only generated temporary JSON lease files
and the unchanged distributed lease functions to test competing writer exclusion
followed by release on confirmed exit. No customer exports are read.

The suite covers 20 cases: failed and confirmed kills, pre/post-kill probe errors,
wait errors, durable intent ordering, forced and shared-key overlap guards,
exactly-once completion, retries, legacy state, restart, timeout hot reload,
normal exit codes, original deadline retention and state-write recovery.
PowerShell 5.1 checks exercise extracted functions only: the resident orchestrator
still requires PowerShell 7. They are not runtime qualification on PowerShell 5.1.

Actual results on Windows: baseline 1.5.10 passes 8/20 and fails 12/20 on both
PowerShell 7.6.6 and Windows PowerShell 5.1.19041.6456; the signed candidate passes
20/20 under AllSigned on both engines. The existing distributed scheduling mock
suite also passes under PowerShell 7 AllSigned against the candidate orchestrator.
Both changed PowerShell files parse without errors on both engines and contain no
non-ASCII characters. PSScriptAnalyzer 1.25.0 reports zero Error findings on
PowerShell 7; the analyzer is unavailable in the Windows PowerShell 5.1 environment.
Both candidate signatures verify as Valid on this host. They use the existing
publisher certificate with no external timestamp or trust-store change. This
does not qualify raw GitHub LF signatures, target-host trust or deployment.

## Contracts, prerequisites and permissions

- Inventory DATA-LAST/DATA-ALL CSV names, columns, encoding, business identifiers
  and consumer mappings are untouched. SmartFinOps Workplace and
  SmartWorkplaceDashboard require no corresponding change for this lot.
- The state schema version remains 1. TimeoutRequested is optional and written
  only for a pending timeout; old records without it are supported. Existing CSV
  statuses and retry behavior apply after confirmed completion. Older binaries do
  not implement the new guarantee even if they can read the additional field.
- The dedicated service account still needs state/lease folder write access and
  rights to inspect/terminate its child processes. Restrict code, configuration
  and state writes to trusted operators. No new Graph/Exchange/AD permission,
  credential, secret or certificate trust change is introduced.
- This is repository distribution with existing dependencies and templates. No
  per-script ZIP, Gallery publication or suite release is proposed. Website delta
  is zero for EN/FR/IT/ES/DE/AR; the lot does not introduce a public product claim.

## Remaining limits and next lot

Confirmed exit refers to the tracked process, not proof of termination of every
detached descendant. Initial re-adoption still uses existing PID/name/start-time
matching; access-denied ambiguity and PID reuse require separate tests and review.
Cross-server storage partitions, clock skew, process crashes, stale-lock takeover
and mixed-version clusters are not qualified. Lease renewal failure is logged;
the current design cannot fence writers when shared storage or supervision is lost.
Tick delays and multiple blocking kill waits can delay renewals and heartbeats.
Global lock acquisition and terminal state/claim/CSV cross-file ordering remain
outside this bounded correction.

Collector pagination, throttling, partial-result handling, multi-file last-valid
publication, weekly history races and source freshness remain open from lot 1.
No live process/tenant/SMB/consumer refresh or scheduled task was tested.

Proposed next lot: reproduce initial re-adoption identity/access failures and
global-lock races with synthetic state, then prepare narrowly scoped lifecycle
corrections. Changes to business algorithms or data contracts require prior review.
