# Smart Workplace Dashboard 3.8.0-beta.1 — BETA

Prepared 2026-09-09. GitHub release must be a prerelease, not latest/stable. No release is claimed as published by this preparation.

The canonical generator previously identified itself as 3.7.1 and the validator as 4.9.1. This preparation introduces one product version (version.json) and carries BETA through both scripts, model/report metadata, all page subtitles, documentation, ZIP name and six-language site guides. Component/schema protocol versions are not product release versions.

- Build without client exports using 90 pinned schemas and a bundled public logo. Default validation no longer probes a local data folder implicitly.
- Require one explicit tenant key/root; nonrecursive current-file selection and row TenantKey checks; reject missing/ambiguous CSVs and incompatible headers. Accept comma/semicolon CSVs.
- Reject declared/unknown partial inventory and future source timestamps. Use oldest recognized row refresh time, preserve unknown timestamps and round the history completeness threshold upward.
- Preserve Base64 immutable-ID case. Reject conflicting exact mappings, duplicate keys with different rows and entities without canonical IDs; preserve unmatched entities with usable identifiers.
- Keep invalid/nonfinite numbers and fractional counts unknown. Exclude blank IDs from entity counts and unknown booleans from disabled/noncompliant counts. Do not treat missing sign-ins as proven inactivity or future activity as recent activity.
- Preserve missing/invalid primary license evidence as blank; avoid invented SharePoint base capacity when the license source is empty or invalid.
- Correct stale handoff/page-count documentation and the website's obsolete dependency description. No other application is changed.

See VALIDATION.md for executed tests and KNOWN-LIMITATIONS.md for unqualified behavior.
