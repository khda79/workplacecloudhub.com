# SmartFinOps Workplace 1.5.0-beta.1 — Beta

Published on 2026-09-09 as a [GitHub beta prerelease](https://github.com/khda79/workplacecloudhub.com/releases/tag/smartfinops-workplace-v1.5.0-beta.1). The approved ZIP is immutable and retains the preparation audit notes. Every GitHub release of this candidate must be a **prerelease**, with tag `smartfinops-workplace-v1.5.0-beta.1` and `make_latest=false`.

The analyzer now excludes invalid schemas from calculations, consistently imports comma/semicolon CSVs, checks the oldest refresh timestamp across rows, and treats unknown/future timestamps as unreliable evidence. A missing source directory produces a partial report instead of a crash. Empty and singleton collections are handled explicitly.

Removal/downgrade valuation is suppressed when required evidence is incomplete or stale. Missing or malformed storage does not prove F3 compatibility. Conflicting directory state, simultaneous base suites, duplicate canonical identities, invalid capacity values and duplicated tenant-SKU rows are guarded. Capacity reuse remains separate from savings and mailbox conversion adds no second license value. Decimal CSV outputs are culture-independent.

Historical Power BI authorization failures are no longer asserted for every customer. All release surfaces retain the beta label. See [KNOWN-LIMITATIONS.md](KNOWN-LIMITATIONS.md) and [VALIDATION.md](VALIDATION.md) before evaluation.
