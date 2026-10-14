# SmartInventory lifecycle

`SmartM365-InventoryLifecycle.ps1` is a definition-only helper shared by the
PowerShell 7 Core module and its Windows PowerShell 5.1 compatibility module.
It does not initialize a tenant, connect to a provider, or run an inventory.

Entry points start an owned run before initialization and finalize it in an outer
`finally`. The guard distinguishes normal return, terminal error and pipeline
cancellation, including cancellation during an inner cleanup block. Internal
helper returns are not entry-point completion.

- Successful runs retain their existing success/warning behavior.
- Terminal errors retain their error record and produce a nonzero exit code.
- Ctrl+C produces `Cancelled` and exit code `130`, not a success receipt. Shared
  mail, Teams, upload and retention helpers skip post-cancellation publication.
- Completion summaries are deferred until owned resources have been closed.
  Cleanup is idempotent; one cleanup failure is reported and does not skip the
  remaining resources.
- Only sessions, transcripts, jobs, runspaces and temporary paths owned by the
  current run are released. Local configuration is not changed by cancellation.

The existing guarded Devices, Licensing, M365 Usage and SharePoint on-premises
collectors keep their specialized interruption contracts. Machine-readable
endpoint detection keeps its stdout payload contract and releases COM/event
records independently.

Temporary staging is distinct from retained evidence: audit files, logs,
published batches, resume state and the two classification workbooks in `DATA`
are not a general temporary-file purge. Interrupted classification replacement
restores the previous pair; uncertain publication pointers retain their candidate
batch for inspection rather than risking deletion of a referenced batch.

Offline coverage: `Tests/Test-SmartM365InventoryLifecycleOffline.ps1` under
`SmartM365`, and `Test-PreparedLifecycleCleanup.ps1` under
`SmartWorkplaceIntelligence/scripts`. Both use isolated synthetic resources;
they do not execute collector or preparation entry points.
