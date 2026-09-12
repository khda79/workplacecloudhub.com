# Current CSV contracts — V1

This inventory is generated from the supported schema JSON contracts as of
2026-09-11. JSON is authoritative; this document highlights entity grain and
the V1 additions that affect consumers.

## Curated and Power BI contracts

`SmartWorkplaceCMDB.tables.json` contract `0.4.0` defines the core curated and
Power BI files:

- `CMDB_Users.csv`: one Entra-derived user CI per `CmdbUserId`; includes
  `UsageLocation` and `UsageLocationStatus`.
- `CMDB_Devices.csv`: one reconciled cloud device CI per `CmdbDeviceId`;
  includes `EncryptionState` alongside ownership, compliance, management,
  primary user and sync evidence.
- `CMDB_Groups.csv`, `CMDB_Licenses.csv`, `CMDB_Mailboxes.csv`: one entity per
  stable CMDB key.
- `CMDB_UserDeviceRelationships.csv` and `CMDB_Relationships.csv`: typed,
  tenant-scoped edges with referential-integrity checks.
- `CMDB_DataQuality.csv` and `CMDB_BuildManifest.csv`: stable findings and build
  evidence.
- Power BI dimensions/facts include country fields on `DimUser`, encryption on
  `DimDevice`, and existing license, mailbox, compliance, relationship and
  quality grains.

Every tenant-scoped table begins with `TenantKey`, `OrganizationKey`,
`EnvironmentKey`, `TenantId`. `SourceID` is never a global key; source evidence
uses `SourceSystem + SourceID`.

## Raw contracts

`SmartWorkplaceCMDB.raw.tables.json` contract `0.11.0` defines AD domains,
users, groups, computers and direct memberships; Entra users, groups and
devices; Intune managed devices; Microsoft 365 subscribed SKUs and assignment
paths; and Exchange Online mailboxes.

V1 adds `UsageLocation` to Entra user observations and `IsEncrypted` to Intune
managed-device observations. Existing source identifiers and source collection
timestamps remain unchanged.

## Additive V1 contracts

- `SmartWorkplaceCMDB.hardware.tables.json` `1.0.0`: one source-reported Intune
  hardware observation per `ManagedDeviceId`.
- `SmartWorkplaceCMDB.ci.catalog.json` `1.0.0`: common User, Device, Group,
  License and Mailbox CI types, source mappings and governance policy.
- `SmartWorkplaceCMDB.ci.context.json` `1.0.0`: optional typed user/device
  context, native device evidence and reviewed organization references.
- `SmartWorkplaceCMDB.ci.hardware.json` `1.0.0`: validated CI hardware output.

Missing values remain missing or explicitly qualified. Hostname, serial number,
primary user, department and ownership are not promoted into identities,
locations or accountable owners without an authoritative contract.
