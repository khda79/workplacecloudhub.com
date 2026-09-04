# Stable 0.3.0 promotion - 2026-09-04

Stable 0.3.0 promotes preview18 at the user's explicit request, before an interactive Windows 11 validation. This release-channel choice must not be presented as Windows 11 certification.

- Normalized comparison of the complete GUI source against the installed preview18 (413,487 characters, excluding Authenticode signatures) confirms that only the version changed. Diagnostic algorithms and collection scope are unchanged.
- Collection (15 checks), correctness (28 checks), analysis regression, cache/performance, timeout and packaging suites pass under PowerShell 7.6.5 and Windows PowerShell 5.1.19041.6456. ValidateOnly passes under both. Packaging parses nineteen PowerShell files and tests controlled installation, integrity, detection and rollback/removal contracts.
- GUI version, runtime metadata, module manifest and Intune detection are 0.3.0; the prerelease value is empty. Changed PowerShell files are signed and timestamped. Bundle checksums and stable administrator commands were verified.
- Stable Gallery and Intune artifacts were built locally. CurrentUser installation reports 0.3.0, no integrity issues, no missing files and automatic updates disabled. The previous installation is retained in a sibling backup directory; user data/logs are preserved.
- No publication to Gallery, GitHub or Intune, tenant change, AI request, Git commit/push, or new live diagnostic scan was performed as part of promotion.

The preview18 user tests supply the real-capture evidence: complete reading of 18 Intune EVTX files (100,779 events) and 8 local EVTX files (69,724 native events = 67,557 logical events + 2,167 identical copies); all 202 eligible Winget files were collected. This is evidence for the tested workflows and conservation rules, not exhaustive certification of every diagnostic finding. See the preview17/18 validation notes for remaining limitations, including synchronous GUI finalization and restricted redacted-text export.
