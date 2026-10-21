# On-prem Messaging

The page follows Messaging & Hybrid and provides Overview, Servers & compute,
Storage, Databases, Readiness, and Collection & details views. It retains every attribute
of ten optional Exchange infrastructure exports at source-file, record and
attribute grain: servers, compute, logical/physical disks, remote access,
per-server summary, readiness, inventory summary, service health and database
paths. Missing artifacts produce explicit Not captured markers, not healthy
zeroes. Empty captured sources are distinguished from absent captures. Detail
views retain collection statuses, source file and record keys.

Sources declared in the manifest still require the existing batch, tenant,
schema, path, size and stable-pointer checks. Optional infrastructure capture
uses the declared semicolon delimiter with header/row validation and hash
read-back. Existing batches remain readable; a future prepared batch is required
to retain infrastructure files not present in the current selected batch.
Infrastructure is global: Country and AD Domain do not invent server location.
Observed readiness flags are neither licensing approval nor migration approval.

The database view uses only the selected immutable batch: MailboxDatabase
properties from readiness configuration and local mailbox Database references.
It shows names, associated servers, observed Mounted status, observed mailbox
counts, database size and reusable internal space. Sizes use explicit source
byte counts; unavailable or unparseable sizes remain blank. Reusable internal
space is not free disk capacity. Missing mailbox Database columns or sources
produce unknown counts; zero observed mailboxes does not prove an empty base.
Mailbox and infrastructure collections are not simultaneous. Unassigned mailbox
references and missing infrastructure are explicit; DAG copy health is not
captured. No cross-batch or DATA-LAST fallback is used. The hidden Loaded Batch
ID participates in the existing Loaded Batch Audit, without new relationships.

Overview charts show observed mailbox counts and sizes by database. Collection
& details charts show observed databases by mount status and associated server.
The capture-coverage KPI and tables are replaced; collection issues remain in
the KPI band and raw attributes remain available. Legacy Exchange means 2016
or earlier (version 15.1 or earlier). The legacy KPI shows count and percentage
of observed inventoried servers; percentage is unknown if a server version is
unavailable. The database KPI counts only infrastructure-observed databases.
Country/AD Domain do not filter these global observations; Dataset, Server Name
and Category slicers apply to attribute views. Databases has its own server
slicer; database charts and KPIs share that table's filtering context.


Report and model definitions require a future prepared batch containing the optional
infrastructure artifacts before populated visual values can be qualified. Structural
checks and synthetic tests do not refresh or publish a model.
