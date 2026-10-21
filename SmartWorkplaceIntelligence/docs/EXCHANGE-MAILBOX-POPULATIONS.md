# Exchange mailbox populations and licensing view

The Licensing & reconciliation view in Migration to Exchange Online/SE separates
native mailbox counts from distinct people in the estimated licensing scenario.
It does not refresh, prepare or publish evidence as part of report authoring.

## Population definitions

- Total on-prem mailboxes: distinct nonblank native mailbox GUIDs in one batch,
  in the current report filter context. A mixed or missing batch is unknown.
- Register: remain on-prem: mailboxes with a confirmed retained register match.
  Duplicate pending hints do not add a second mailbox or override retention.
- Service accounts: remain on-prem: currently Not captured. The prepared AD user
  source lacks the governed classification and Workforce excludes non-human
  accounts. Generic accounts and managed-service object classes cannot replace
  the requested classification. Missing evidence is not a healthy zero.
- Potential EXO: at most the mailbox pool left after confirmed retention and
  pending register hints, before service-account and technical review. Pending
  exact native-ID or SMTP hints only quarantine candidates. They do not establish
  identity, a retained decision, or licence entitlement. Pending register rows,
  distinct quarantined mailboxes and confirmed mailboxes have different grains.

The approved Prepare extension preserves the raw AccountType for every AD user,
including disabled/non-human accounts and every currently unused or future type.
It does not derive or restrict this value to a service-only enumeration. Missing
fields stay blank in the prepared artifact. The reusable
ReadPreparedADAccountClassification reader preserves all values, uses unique
same-batch native keys and labels missing/unknown classification explicitly.

The complete activation model can be built with prepare_activation_model from
exchange_account_classification.py. It materializes an AD Account Classification
utility table with all raw types, adds it to Loaded Batch Audit, excludes confirmed
service accounts from EXO sizing, and unions register/service retention once per
licensed identity. Unknown or ambiguous classification is quarantined from new
cloud sizing. SE remains a known minimum; E3/F3 labels flag incomplete coverage.

This activation model is prepared separately from the working cached PBIP. Adding
an unprocessed imported table to its active measures would break the current view
until refreshed. Activate its definitions together with the next authorized
refresh after a future qualified batch supplies AccountType. The working PBIP
already contains the callable reader, but its service KPI intentionally remains
Not captured until that activation; a refresh alone does not wire the new facet.
No Prepare, data refresh, publication or commit is performed by these helpers.

## Licensing presentation

E3 and F3 estimated To buy values are larger than their supporting demand,
assigned-suite and unassigned-stock callouts. Existing approved sizing assumptions
remain visible: F3 assumes frontline eligibility, at most 2 GB and no archive,
hold or delegation; E3 covers other known user cases up to 100 GB. Assigned suites
are deducted even when service plans need repair. Only qualified same-batch tenant
stock is deducted, shared across countries and assumed allocatable, not reserved.
The service-data gap makes these preliminary estimates, not purchase orders.

Exchange SE shows Total required SE licences, expressed as a known minimum while
service accounts and other accessors are not inventoried. This is the distinct
confirmed retained user scope, not one licence for every mailbox. Observed
equivalents and the scope without an observed equivalent remain supporting
diagnostics; they are not subtracted from the total required scope. The previous
net-purchase To qualify callout is no longer displayed. Contract coverage, active
rights, user/device CAL choices, base/additive rights and special-mailbox accessors
still need assessment before procurement.

The other six migration views, bookmarks, filter-state preservation, source
partitions, batch selection, tenant checks and relationships are unchanged.
