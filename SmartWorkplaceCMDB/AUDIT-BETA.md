# SmartWorkplaceCMDB 0.3.0-beta.1: local audit

Date: 2026-09-09. Product status: **BETA**. This is a locally prepared beta
prerelease candidate, not a stable release or a production qualification.
External Git/release/site publication requires approval of its concrete scope.

## Demonstrated corrections

The first audit suite reproduced six failures before correction: exporter tenant
relabeling, ignored CSV column contracts, multiline report overcounting, missing
report tenant checks, newest-row freshness masking, and bounded orchestration
using canonical output paths. All six regressions now pass.

The source-evidence suite reproduced eleven further failures before correction:
eight direct collectors could overwrite canonical raw output during a trial,
empty JSON arrays were not accepted reliably, failed attempts had no evidence,
and Intune merging concealed the older contributing source date. Sixteen
source-evidence cases now cover these fixes and additional safety checks.

- Fixture and bounded collection roots carry tenant identity and mode markers.
  New empty explicit roots can be used for trials. Occupied unmarked roots or
  roots with another identity/mode use a unique test child folder instead.
  All trial child paths are pinned; external raw output overrides are refused.
  Unbounded live collection cannot reuse a root marked fixture or bounded.
- Raw source sidecars record in-progress, completed and failed attempts, coverage,
  date, row count and SHA-256. Empty completed sources remain distinguishable.
  Failure preserves old CSV bytes while preventing their use as a current
  successful snapshot. Per-source file locks reject overlapping attempts.
- Raw normalizers, including validation mode, reject failed/in-progress evidence,
  identity mismatch and changed snapshot bytes. Legacy files without evidence
  remain readable, with unknown source health in the report.
- Missing/malformed Graph value arrays fail rather than becoming an empty
  successful snapshot. Empty and one-item arrays work in offline fixtures.
- Source health and data-quality findings expose bounded, fixture, scoped,
  not-collected, stale and invalid evidence. AD scope restrictions and omitted
  memberships are explicit. Intune enrichment retains the oldest contributing
  date and an unknown date if either contribution lacks a date.

## Verified offline

- 18 PowerShell 7 suites pass. Suites with numeric summaries total 153 passing
  cases and zero failures; the scheduled-task preview/contract suite passes
  separately without a numeric summary.
- The full fixture pipeline exercises 22 steps, 20 canonical CSV contracts and
  17 source contracts. CSV schema files and consumer projects are unchanged.
- All 50 PowerShell files parse without errors under Windows PowerShell 5.1 and
  PowerShell 7. Runtime qualification is PowerShell 7 only.
- All 28 changed/added PowerShell files have valid timestamped Authenticode
  signatures in the CRLF worktree, with script bodies compared before/after
  signing. Source-evidence regressions pass under PowerShell 7 AllSigned.
- On this host, PowerShell 7 AllSigned cannot auto-load the built-in
  Microsoft.PowerShell.Security module because Security.types.ps1xml fails
  authorization. Package signature checks therefore use Windows PowerShell 5.1;
  no system module, certificate trust or execution policy was modified.
- No live Graph, Exchange, AD or SharePoint collection was run. Scheduled-task
  operations used previews and contract inspection; no production task changed.
- Detailed evidence is retained locally under the ignored tmp/cmdb-audit folder.
  Package inventory/checksums and extracted-package checks are retained in the
  separate local release review; neither runtime data nor logs enter the ZIP.

## Remaining limits

Cross-source collection and curation are not transactional. Complete evidence
means completion of the collector's requested scope, not universal inventory
coverage. Markers and checksums prevent accidental mixing/drift; they are not a
security boundary against a user who can edit local files. Legacy output roots
without markers are not universally protected by an ownership lock.

The SharePoint uploader still transfers CSVs only. Evidence sidecars must travel
with raw files to retain source health elsewhere; a CSV-only copy is unknown.
Live tenant permissions, authentication, source scale, scheduler execution and
Power BI Desktop refresh remain unqualified. SmartWorkplaceDashboard currently
selects SmartM365 source tables and is not included or certified as a consumer
of this beta. The prepared website delta only updates the CMDB catalogue entry
in six languages; unrelated translation debt is excluded.

The next action is review and explicit approval of the beta package, exact Git
file list and six site files before any first push or external publication.
