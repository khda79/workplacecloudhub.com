# Smart Workplace CMDB

CMDB is now a read-only consumer of prepared SmartInventory evidence. Version
1.2.0 retires its duplicate collectors, normalizers, collection modules, mail
reports, standalone scheduler and their launchers. Do not reinstall that chain.
The SmartM365/SmartInventory collectors and orchestrator are not removed.

## Current production boundary

SmartInventory acquisition → DATA-LAST → CmdbEvidence-Prepare →
DATA-POWERBI-CMDB (46 CSVs + current.json.txt) → ephemeral loopback reader →
native Power BI Desktop refresh.

The four licensing report tables use the separately published SmartM365
DATA-LAST licensing snapshot. They are not among the 46 prepared tables and
are not an atomic part of that batch. No claim of perfectly synchronized
collection times across producers is made.

The private migrated project has 13 report pages. Its 46 prepared-table row
counts and the scoped date conversion fixes were checked in Desktop. Final
page-by-page functional qualification was explicitly waived; it is not recorded
as passed. No Fabric deployment or scheduled unattended Desktop refresh is
included in this change. Release metadata keeps liveQualified=false.

## Refresh

Use [PowerBI/README.md](PowerBI/README.md) and the signed launcher
Launchers/Start-SmartWorkplaceCMDB-Refresh.ps1. Validation holds one complete
batch in memory before starting a loopback listener. An expired or inconsistent
batch is rejected, never replaced by historical data. This is a guided native
Desktop workflow, not a one-click model update: set the three private session
parameters and refresh in Desktop while the launcher remains open.

Python 3.11+ and the matching SmartM365 prepared contract, registry and freshness
module from this repository are required. The CMDB-only archive is not a
standalone collector distribution. No Graph, Exchange, AD, certificate-based
collection credential, pandas or installed reader service is required.

## Retirement and data

Old DATA-ALL, DATA-LAST and LOG-ALL under the independently owned Smart-CMDB
collection root are obsolete. Removing those data/history files is permanent
from this workstation; cloud recovery depends on the provider's retention.
Never target the protected SmartM365 synchronized DATA folder or the private
PBIP, model, report, current licensing snapshot or prepared batch.

Before removing old files on another host, disable only scheduled tasks whose
Actions reference the retired SmartWorkplaceCMDB orchestrator/collectors and
verify no such collector is running. The current SmartM365 orchestrator is
unrelated and must remain enabled. Local cleanup does not prove remote cleanup.

Historical PowerBI builders, schema documentation and synthetic fixtures remain
as developer tools. Do not regenerate the migrated private PBIP with them.
They are not invoked by the current refresh launcher.

## Public/private boundary

Git contains reusable code, templates, source contracts and synthetic tests.
Tenant configuration, tokens, real exports, logs, cache, PBIP/BIM and private
qualification evidence remain ignored. Native date parsing permits exact source
wall-clock values only for the eight explicitly mapped fields; the Autopilot
minimum-date sentinel is blank only in its contact field. No source CSV was
rewritten to mask conversion failures.
