# Smart SharePoint Migration Toolkit 1.0.64

Release 2026-10-08. Dashboard 1.0.64; generic launcher 1.0.28.
This release packages the published toolkit updates since 1.0.8.

## Dashboard and reports

- Maximized dashboard with prominent file and permission comparison rates,
  comparison dates, progress bars and an equal-weight global rate:
  `(files % + permissions %) / 2`. The global rate is unavailable if either
  input is unavailable; reports from older scans are marked for recalculation.
- Overview scan-gap badges: green through 12 hours, yellow above 12 hours
  through 24 hours, red above 24 hours. Status badges identify required actions.
- File inventories publish file counts, folders containing files and current
  file volume. The GUI reads these scan metrics; empty folders, previous versions
  and recycle-bin content are excluded from these metrics.
- Source and destination website buttons sit in the scan cards beside Open folder.
  Comparisons are disabled when their required scans are absent or incomplete.
- ShareGate reports are analyzed automatically in a background queue. XLSX
  support installs ImportExcel for the current user when needed. Analysis reads
  a stable, SHA256-verified local workbook copy with bounded retry.
- Missing ShareGate analysis is displayed as unavailable, not as zero issues.
  Full-width issue review and raw rows, single-click pattern selection, progress
  indicators and a dedicated Source farm diagnostics window improve review.
- Global file and permission reports include HTML summaries and Excel exports.

## Batch workflows

- Source scan, destination scan and comparison batch scripts have dedicated CMD
  launchers and a Batch runs tab with plan preview, results and log shortcuts.
- Destination batches allow up to two concurrent scans in both Interactive and
  Certificate modes. The limit applies per batch; source scans and comparisons
  remain sequential. Missing or incomplete prerequisites are reported explicitly.
- Source batches stage required scripts locally for SharePoint Server execution.
  Batch summaries link each action to its output and run log; cancellation and
  child-process failures propagate to the caller.

## Scan and evidence reliability

- Incomplete permission inventories are rejected instead of being published as
  successful comparison inputs. Verified empty destination file inventories can
  be compared; empty source inventories have no comparison rate.
- Associated SPO Members, Owners and Visitors group titles are collected
  correctly, with collection errors visible. Item error records include identity
  and path details; duration formatting and authentication diagnostics are fixed.
- SPO inheritance flags are read in bounded CSOM batches of at most 100 items,
  avoiding oversized requests while retaining paged enumeration. Discovery and
  transient page failures have bounded retries and explicit failure reporting.
- Scan manifests and comparison evidence checks preserve provenance. Shared
  console lifecycle output and portable Python helpers are included.
- Read-only farm diagnostics, ShareGate probes and guarded placement/duplicate
  review tools are included; operations retain their individual prerequisites
  and explicit execution controls. See README for their limits.

## Installation and validation

Download the ZIP, SHA256 checksum and file manifest. Verify the ZIP checksum,
extract the complete package and start
`SmartM365/SharePointMigration/Start-SmartM365-SharePointMigration-GUI.cmd`.
The dashboard requires Windows and PowerShell 7.4 or later; SharePoint Server
inventory and farm diagnostics require Windows PowerShell 5.1 on the appropriate
server. Python/PnP/Graph requirements are documented in README. Python runtimes,
third-party modules and private migration workspaces are not bundled.

The package includes the public code-signing certificate and optional trust
installer. Trust must be configured for the account/machine executing signed
scripts. The package contains templates, not customer scans, credentials or logs.
Existing private migration folders must be preserved when updating.

Release checks cover package hashes and signatures, GUI resource loading and
offline regressions for comparisons, diagnostics, inventory reliability, batch
workflows and metrics. Offline checks do not certify production authentication,
tenant/farm permissions or complete real scan coverage. No tenant scan, migration
or cleanup is executed as part of this release validation.
