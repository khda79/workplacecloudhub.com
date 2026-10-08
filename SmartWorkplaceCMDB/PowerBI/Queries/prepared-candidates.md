# Prepared model templates and read-session contract

The file names retain “candidate” for compatibility with the migration tools.
The loopback reader and typed function are installed in the privately migrated
model and all 46 prepared-table row counts were checked in Desktop. The
query-scoped DAX file and model-plan CLI are authoring references, not commands
that modify or refresh the project. Report-wide functional checks were waived.

ReadSession first validates and retains a complete batch. Every request must
use the same random token and pinned manifest digest, and every partition uses
the same three model session parameters. Wrong token/table returns 404, a wrong
batch returns 409, and a closed/expired session returns 410. Browser origins and
non-loopback hosts are denied. Protocol 0.1.2 provides exact schema, counts,
identity and source-specific freshness without producer paths. No service,
CSV backup/history, pandas, CORS or request logging is introduced.

Application product identity is the stripped/casefolded name/publisher/platform
JSON tuple encoded as lowercase UTF-8 hex, excluding version. The session adds
Application product key only to the catalog projection, never to source CSVs.
Missing publisher is not merged with an observed publisher. Distinct product
device counts are not a sum of version counts or a distinct fleet count.

Strict M types keep missing numeric evidence null. Source wall-clock dates are
opt-in for the four DimUser/Intune/update/Autopilot mappings (eight columns),
never a global locale fallback. Exact 01/01/0001 00:00:00 is null only in the
Autopilot LastContact field. Invalid dates and unexpected ancient values still
fail. Unavailable activity is not proof of inactivity.

Use ../README.md for the guided native refresh procedure. The launcher does
not write the three parameters automatically, process XMLA, save the PBIP or
claim a successful refresh. No functional-report qualification or unattended
service refresh is implied by protocol tests.
