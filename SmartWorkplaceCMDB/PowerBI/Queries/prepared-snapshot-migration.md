# Prepared snapshot consumer — current boundary

The migration now uses 46 prepared SmartInventory tables through one pinned,
immutable loopback session. The private model's table row counts were checked
against the complete prepared batch. Existing model names, measures and report
logic were preserved except for the reviewed source/type/schema adaptations.
Current refresh instructions are in ../README.md.

load_snapshot(root, expected_identity) validates the independently supplied four
identity fields, manifest ownership/status, exact contract version and SHA256,
producer registry digest, complete receipt/source lineage and acquisition
freshness. It retains each original CSV binary, validates exact headers, logical
row counts, immutable keys and parent relations, then rechecks manifest and
contract stability. Synchronization or preparation time does not reset source
age. No producer UNC paths need to be reachable by the synchronized consumer.

Only DiscoveredApps has a weekly 168-hour target with warnings up to a hard
240-hour limit. Core sources retain their 48-hour limits and acquisition spans.
Partial/failed receipt qualification is never bypassed by an age exception.

The snapshot's retained binaries are immutable even if synchronized files
change later. Every Desktop partition must use the same validated session
parameters; a separate preflight followed by File.Contents is not equivalent.
Hash consistency is not cryptographic producer authentication. In-memory CSVs
and validation key sets require additional RAM beyond the configured byte budget.

The four License Report tables consume a separate published SmartM365 licensing
snapshot and are not included in this batch's 46-table qualification. The private
13-page report was saved; its final functional page review was waived. Neither
Git publication nor offline tests establish global report/tenant qualification.

The historical duplicate CMDB collection chain and sidecar preparation have
been removed. This consumer does not recreate them, copy persistent exports,
start collectors, upload data or automatically publish to Fabric.
