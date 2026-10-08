# CMDB operations — prepared consumer

Use PowerBI/README.md for the native refresh workflow and launcher command.
Use the entire matching repository: the consumer depends on the SmartM365
prepared contract, source registry and shared freshness evaluator outside this
product directory. A CMDB-only package is not a standalone installation.

The launcher validates the full batch before binding an ephemeral loopback
port. -ValidateOnly creates no listener and makes no semantic model changes.
Interactive execution prints private session parameters; it neither changes
Desktop parameters automatically nor starts XMLA processing. Desktop refresh,
successful save and stopping the reader are separate explicit actions.

Failure policy: wrong identity, changed batch, schema/hash/key violation,
incomplete evidence, expiry or memory budget exhaustion rejects the batch.
There is no historical fallback, collector execution or data upload.

The old CMDB collectors, scheduler installer, modules, normalizers and mail
reports have been retired. Remove their tasks on each host only after checking
Actions and active processes. Do not disable the SmartM365 orchestrator.
Old Smart-CMDB data/history can be removed from its verified owned root; never
remove SmartM365 DATA, the prepared batch, licensing snapshot or private PBIP.

## Package and publication

Build/SmartWorkplaceCMDB-Package.ps1 packages only Release/Files.json entries,
checks PowerShell signatures and emits hashes. No CSV, PBIP, tenant setting,
cache, private report or token belongs in the public archive. Version 1.2.0
records the breaking retirement; channel metadata is unchanged and
liveQualified remains false. A Git push is not a GitHub Release, Fabric
publication or proof that remote scheduled tasks/data were retired.
