# Smart SharePoint Migration Toolkit 1.0.80

Release 2026-11-04. Dashboard 1.0.80; generic launcher 1.0.36.

## SharePoint Online source and destination

- Create a migration with SharePoint Online as either or both endpoints. The
  wizard supports separate authentication modes and private profile paths for
  source and destination. A tenant admin URL is optional for explicitly mapped
  site scans; operations requiring tenant administration still need it.
- Scan files and permissions on both sides and compare the results. Source and
  destination SPO scans use their own connection settings; existing SharePoint
  Server to SPO scenarios remain available.
- Comparison validates existing scan receipts against file hash, row count,
  schema, side and inventory kind. Older inventories without receipts remain
  available with unverified scan evidence.
- Cross-tenant permission comparison requires explicit confirmed identity and
  site mappings, distinct tenant profiles and review. Mapping does not prove
  group membership equivalence or effective access.

## Diagnostics and review

- ShareGate diagnostics improve source/destination attribution when two sites
  share an SPO host and retain missing field-value warnings for review.
- Optional in-memory connection reuse and an SPO source identity check support
  the guarded one-file qualification and three-file correction workflows.
  Broader real item-scoped batch correction remains disabled pending review.
- Compare item role assignments and group membership independently. ShareGate
  reports and matching role assignments do not prove sharing links were migrated.

## Validation and installation

The creation, four scans and two comparisons were exercised through the GUI
against two test sites in the same tenant. Offline regression tests also cover
existing scenarios and synthetic cross-tenant identity mapping. No real
cross-tenant test was performed. Hidden and system lists and libraries excluded
by the selected scan options are outside that inventory's coverage.

Download the ZIP, SHA256 checksum and manifest, verify the ZIP hash, then extract
the full package. Preserve private migrations, mappings, authentication profiles,
runtime configuration and installed Python when updating. Windows and
PowerShell 7.4 or later are required for the dashboard; SharePoint Server
collection requires Windows PowerShell 5.1 on an appropriate farm host.
Python runtimes and third-party modules are not bundled. See README for the
prerequisites, scope and limitations.
