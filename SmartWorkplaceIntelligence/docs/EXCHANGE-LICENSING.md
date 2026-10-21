# Exchange migration licensing scenario

Version 1.0.0

The Licensing & reconciliation view is part of Migration to Exchange Online/SE.
There is no separate Exchange Reconciliation page. The original six migration
views and their filter-preserving bookmarks remain available.

## Evidence and identity

The reconciliation uses unique same-batch AD ObjectGUID and producer-qualified
native SID links to the Entra account. SMTP and job titles are displayed, not
used to confirm identity or entitlement. AD and Entra job titles remain separate.
The hidden server-reference table reads the captured native mailbox artifact
through the existing tenant, manifest, size and stable-pointer guards. It is
included in Loaded Batch Audit; the derived reconciliation is not a redundant
source vote. Existing empty-table limitations still apply.

## Approved estimation assumptions

- Count distinct user identities, not mailbox rows. Multiple mailboxes use the
  greatest requirement. Conflicting retained/migration targets remain in review.
- F3 assumes frontline eligibility, at most 2 GB and no archive, hold or
  delegation dependency. Size alone does not establish frontline eligibility.
- E3 is the scenario suite for other known user cases up to 100 GB. Incomplete
  identities, unknown sizes/features and larger mailboxes remain unpriced.
- Shared, resource and system mailboxes require separate assessment.
- Active assigned suites reduce demand even when a service plan needs repair.
  Assignment errors remain in review; plan repairs are not duplicate purchases.
- Confirmed retained exceptions require the native local identity, same batch,
  Matched local status, Retained=1 and Requires Review=0. They stay in SE scope
  and are excluded from EXO E3/F3 demand. Pending register matches are quarantined
  from migration sizing; an SMTP exclusion hint never confirms identity or rights.
  A confirmed native retained decision takes priority over a duplicate pending hint.

E3/F3 estimated purchases = max(0, scenario users - observed assigned coverage -
qualified unassigned tenant stock). Stock uses same-batch governed SPE_E3 and
SPE_F1 capacity rows, one snapshot and unique SKU IDs. Missing, mixed, duplicate,
negative or unqualified stock remains unknown. Stock is shared across countries,
assumed allocatable and not reserved; recovery candidates are not free units.

SE cards distinguish confirmed retained users, observed equivalency and residual
CAL exposure. Existing CAL/SA stock, active contractual coverage and other
accessors are not inventoried: net CAL purchases remain To qualify. Review rows
are not distinct people and must not be added to the unpriced-person count.
These are partial planning estimates, not a contractual purchase order.

## Validation and operating boundaries

Synthetic DAX cases cover identity, retained routing, incomplete evidence,
assignment/plan states, deduplication, product thresholds and qualified stock.
Synthetic PQTest cases exercise the actual server-reference M reader. Static
checks validate DAX references, PBIR, filter-preserving navigation and layout.
Preparing definitions or running these tests does not refresh or publish a model.
Source parameters, collection, Prepare, retention and deployment are unchanged.

## Microsoft references

- [Exchange Online limits](https://learn.microsoft.com/en-us/office365/servicedescriptions/exchange-online-service-description/exchange-online-limits)
- [Enterprise to frontline plans](https://learn.microsoft.com/en-us/microsoft-365/frontline/switch-from-enterprise-to-frontline)
- [CAL equivalency](https://www.microsoft.com/licensing/terms/product/CALandMLEquivalencyLicenses/all)
- [Exchange licensing FAQ](https://www.microsoft.com/en-us/microsoft-365/exchange/microsoft-exchange-licensing-faq-email-for-business)
