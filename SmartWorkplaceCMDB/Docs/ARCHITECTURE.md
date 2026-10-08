# Architecture and boundaries

## Evidence flow

SmartM365 SmartInventory publishes raw CSVs and producer receipts. SmartM365
CMDB preparation validates its contract and publishes a prepared batch.
SmartWorkplaceCMDB consumes that output instead of duplicating collection.
Its reader needs the matching monorepo preparation contract, source registry
and pure freshness policy; hash checks remain enforced.

The prepared branch is: SmartInventory -> CMDB preparation -> immutable
validated reader -> Power BI prepared tables -> measures and report pages.

A separate branch is: SmartM365 license report -> four published license CSVs
and snapshot descriptor -> License Report tables -> pages 04A, 04B and 04C.
These tables are not part of the prepared batch. Their reader checks schema
and snapshot identity, not simultaneous acquisition across both branches.

## Reader lifecycle

ValidateOnly checks and retains the batch without starting a server or
altering Power BI. Interactive mode serves the retained bytes on 127.0.0.1
with a random port, private token, pinned batch hash and no-store responses.
Changing synchronized files cannot change bytes already held by the session.

The reader stops on Enter after refresh, Ctrl+C, its bounded lifetime or the
earliest source deadline. Wrong identity, hash, schema or batch is rejected.
It never collects, uploads, processes via XMLA or saves the report.

## Public and private projects

PowerBI/Project preserves the existing 13-page PBIP/PBIR and BIM-format model:
103 saved tables, 249 measures and 80 relationships, including existing date
tables. This is not a model redesign or a conversion to TMDL.

The public project has neutral identity, source and session parameters and no
data/cache. Private localSettings, actual entity selections, exports and
credentials are excluded. The working project/configuration are under .local/.
Public updates must not overwrite the private project or unsaved user edits.

## Qualification and cadence

Perfect cross-source synchronization is neither required nor promised.
Acquisition times, producer completion and source-specific limits remain
explicit. Weekly application evidence retains unresolved relations without
inventing parents; missing context must not be presented as a verified zero.

The prepared branch contains 46 contract-defined tables. Synthetic tests,
structural checks and Git publication do not prove tenant permissions, server
execution, SharePoint synchronization or Desktop operational acceptance.

## Remaining directions, not commitments

Useful remaining ideas from the retired V2 backlog are authoritative site and
entity mappings, owned business-service catalogs, scoped Azure dependencies,
governed vendor lifecycle/EOL, richer source-backed impact graphs,
procurement-backed costs/savings, and an enterprise application catalog.
Each needs separately reviewed sources, scope and operating cost. Already
integrated applications, Autopilot, memberships and hybrid identity are not
listed again as future collector work.

## Existing report validation warnings

The generic copy retains the private report's existing three 360-page
drillthrough configurations. PBIR validation reports missing Drillthrough
creation markers and visible Back buttons on those pages (nine warnings,
no schema errors). This cleanup does not silently redesign their navigation
or claim page-by-page functional acceptance. Review navigation separately
before broad operational rollout.
