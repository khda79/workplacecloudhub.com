# CI source evidence and governance — V1

The common registry preserves existing `Cmdb*Id` values as `CI_ID`. Optional raw
evidence creates one `CMDB_CISources.csv` row for each validated native source
identity. Correlation candidates are retained rather than silently reselected;
`SourceSystem + SourceID` is the source key.

An optional complete governance journal may declare business owner, technical
owner, support group, organization reference and lifecycle transitions. Events
require a known actor, reason, UTC effective time and valid previous value.
Allowed transitions and required fields are defined by the stable catalog.

Without a journal, owners remain blank with `OwnershipStatus=NotCollected` and
lifecycle remains `Unknown`. Primary user, device ownership and manager are not
substitutes for accountable owners. Journal replay is deterministic but does
not authenticate the declared author or detect omitted history; protect the
authoritative journal outside the repository.
