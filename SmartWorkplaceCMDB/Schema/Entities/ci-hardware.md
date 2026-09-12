# CI hardware context — V1

The optional `-HardwareInputPath` adapter validates a completed Intune hardware
snapshot and its sidecar before creating `CMDB_CIDeviceHardware.csv`. It requires
the original raw source mapping and resolves `ManagedDeviceId` through existing
CI source evidence. It never matches by device name or serial number.

Each row retains tenant identity, `CI_ID`, native Intune IDs, source-reported
serial/manufacturer/model/storage values, qualification statuses and separate
source dates. Missing, reported zero and reported nonzero values remain distinct.

The adapter rejects failed/in-progress collection state, tenant/hash/count/date
mismatch, invalid storage, duplicate native IDs, foreign rows and unresolved
mappings before any destination is published. Empty completed snapshots are
valid and produce a header-only output with explicit coverage evidence.

Exact columns and statuses are defined by
`SmartWorkplaceCMDB.ci.hardware.json` version 1.0.0.
