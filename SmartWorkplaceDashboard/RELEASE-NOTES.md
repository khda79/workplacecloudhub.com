# Smart Workplace Dashboard 3.8.0-beta.2 — BETA prerelease

Prepared 2026-09-10 for publication as a BETA prerelease with the documented qualification limits. The single maintained entry point is pbip/SmartWorkplaceDashboard.pbip.

- Correct unsupported card value properties in the canonical generator and all 161 generated KPI cards. Use labelDisplayUnits and integer labelPrecision, preserving data bindings, positions and labels.
- Add a build regression covering every generated card.
- Set explicit zero padding on all 204 textboxes, preserving their text and positions. Titles and the inactivity threshold render without textbox scrollbars; the PBIR validator now reports zero warnings.
- Keep migration success and Exchange storage-used rates blank when their denominators are absent or zero; also keep storage usage unknown when free space is absent. Seven query-scoped synthetic DAX scenarios cover empty, zero, negative and populated inputs. Valid rates remain unchanged.
- Replace the incompatible DAX variable Rows with _LicenseRows in seven license/capacity measures and the generator. Desktop previously reported seven semantic errors and eight dependent errors; all 189 measures now compile and execute on the synthetic fixture.
- Accept both generator metadata (1600) and Desktop August 2026 metadata (1606) in the validator. Ignore only typed internal rowNumber columns when comparing CSV schemas; retain checks for unexpected source fields.
- Exclude local `.pbi` caches and private files from build-test copies; verify that boundary with synthetic markers. Document safe recovery from an incompatible pre-existing Desktop cache.
- Retain the published beta data-quality guards. The experimental per-query function wrappers are not integrated.
- Local Desktop synthetic refresh, 12 expected KPI values and country-context checks passed. All 19 pages were captured and reviewed; the migration card no longer displays an invented 100% for an empty source. Native CSV/Excel/PDF exports and interactive slicer operation remain unqualified because the available Desktop API does not expose these operations. Remaining limits are tracked in VALIDATION.md.

## Previous published beta

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
