# SmartInventory audit - lot 3: re-adoption and resident lock

Audited and approved for repository publication on 2026-09-11 from commit
35583186952c3c2953650690281b3c0ddd0a2be2. Main orchestrator 1.5.11 -> 1.5.12.
No suite version or stable qualification. The approved repository scope contains
five files; no operational deployment or website publication is included.
Integration base: fda8fa49d28cefe95e906288a997519942042a79.

## Demonstrated corrections

| Synthetic trigger | Baseline failure | Candidate behavior |
| --- | --- | --- |
| PID lookup denied; process name/start time unavailable | No matching process returned, allowing completion and lease release | Preserve uncertainty and saved Running state, retain the local running gate, refresh existing lease and retry identity inspection |
| Access recovers after restart | State may already have been finalized | Re-adopt on a later tick and resume normal supervision |
| PID disappears after an uncertain lookup | Completion can occur before absence is confirmed | Complete only after confirmation; preserve prior TimedOut intent or existing Interrupted outcome |
| Two stale-lock recovery contenders interleave | Both can remove/create the same lock and report ownership | Hold one persistent guard handle across inspection, recovery and residency |
| Malformed lock or inaccessible incumbent | Payload is treated as stale and replaced | Leave it unchanged and refuse acquisition |
| Previous owner exits twice | Must not affect a new owner | Release owned handles once; guard serializes cleanup and acquisition |

RecoveryPending and RecordedPid are in-memory supervision metadata, not new
persisted state or CSV fields. Heartbeat RunningJobs.Pid retains the recorded
integer during uncertainty. The existing JSON lock payload and exit codes remain;
exit code 3 also covers refused/uncertain lock acquisition. The only additional
runtime file is Orchestrator.lock.guard. Its presence alone does not mean a
resident is active; it is deliberately never removed by routine lock cleanup.

## Executed tests

The new Test-SmartM365OrchestratorRecoveryOffline.ps1 parses named function
definitions without running module initialization or the orchestrator entry point.
Processes, PID lookup, lifecycle completion and external operations are mocks.
A synthetic C# object emulates denied identity properties. Temporary local files
exercise real file handles, competing acquisition and a deterministic interleaving
inside the stale-lock read. Abandoned handles are disposed to simulate release;
no real process crash or process termination was executed.

- Baseline 1.5.11: 4/17 pass and 13/17 expected failures on PS7.6.6 and
  Windows PowerShell 5.1.19041.6456.
- Final signed candidate: 17/17 under AllSigned on both engines.
- Existing lot 2 lifecycle suite: 20/20 on both engines against the candidate.
- Existing distributed scheduling mock suite: passed on PS7 under AllSigned.
- Two changed PowerShell files: zero parser errors and non-ASCII characters on
  both engines; signatures Valid on the CRLF candidate.
- PSScriptAnalyzer 1.25.0: zero Error findings on PS7. Analyzer unavailable in
  PS5.1; its absence is not reported as a successful analyzer run.

PS5.1 validates extracted functions only; the resident entry point requires PS7.
Signatures use the existing publisher without an external timestamp or trust
store changes. Raw GitHub LF signature validity and target-host trust are not
qualified by these tests.

## Contracts and distribution

No inventory CSV names, columns, encodings, business identifiers, consumer
mapping, schedule, election weights or collector business algorithms change.
SmartFinOps Workplace and SmartWorkplaceDashboard require no update for this lot.
No operational configuration is read or modified for test fixtures.

The service account retains its existing process-inspection and runtime-folder
rights. File storage must support the requested exclusive/read-sharing modes.
Runtime lock/state directories remain scoped by tenant and server. An inaccessible
PID can intentionally keep work blocked until inspection or state repair is
possible. No new Graph/Exchange/AD permission or authentication is introduced.

Distribution is repository-based with the existing dependency layout. No ZIP,
Gallery release, tag or suite promotion. Website scope EN/FR/IT/ES/DE/AR: zero
files. No scheduled task, tenant, collector, mail, upload or LinkedIn action.

## Remaining limits and next step

Tests qualify local synthetic handle behavior and explicit interleavings, not
a live multi-server SMB environment. Lost storage handles/network partitions,
mixed-version deployment and detached descendants remain unqualified. Legacy
executables do not participate in the new guard protocol. A recycled pwsh PID
can conservatively block a legacy lock; the existing process-start tolerance
and occurrence/lease takeover algorithms are unchanged. Corrupt state can require
operator investigation. No automatic deletion of uncertain lock/state is added.

Collector pagination, throttling, incomplete results, export freshness,
DATA-LAST/DATA-ALL consistency and weekly-history races remain open. SmartInventory
is not globally validated.

Next proposed lot: reproduce partial collection and last-valid export handling
for a bounded group of Graph collectors, preserving existing CSV contracts.
