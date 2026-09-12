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

## Evidence and failure handling

Every source snapshot has a sidecar describing tenant identity, mode, coverage,
row count, SHA-256 and timestamps. Downstream processing stops on failed or
in-progress state, tenant mismatch, hash/count drift or incompatible headers.
The previous valid snapshot is retained when a new collection fails.

Graph transient responses use bounded retries. `Retry-After` is honored within
the configured cap; retry exhaustion fails the source rather than publishing a
partial snapshot.

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
