# Mailbox migration exception register

## Input and activation

The neutral English workbook is
[`SmartWorkplaceIntelligence-MailboxMigrationExceptions.xlsx`](../config/templates/SmartWorkplaceIntelligence-MailboxMigrationExceptions.xlsx).
Upload a **private working copy**, under that exact operational filename, to the
same SharePoint folder as the PersonaClassification and SiteClassification
workbooks. Keep customer addresses, comments, GUIDs and binding files outside Git.
The template contains no customer rows. Its guide, validation lists and input-row
counters are conveniences; only preparation produces qualified mailbox counts.

The preparation wrapper adds these keys from its committed template to its
runtime `.local.json`, preserving existing local values and nested settings:
The JSON transport policy also accepts the corresponding `.json.txt` physical
name and prefers it when present; the same rule applies to the binding registry.

| Setting | Default | Meaning |
| --- | --- | --- |
| `PreparedMailboxMigrationExceptionsEnabled` | `false` | Require the third mapping only when explicitly enabled. |
| `PreparedMailboxMigrationBindingsPath` | empty | Optional private identity registry. An empty setting looks for `SmartM365-WorkplaceEvidence-Prepare.MailboxMigrationBindings.local.json` beside the wrapper's private configuration. If that default file is absent, SMTP-only decisions remain unconfirmed candidates. An explicit nonexistent path fails preparation. |

The SharePoint folder and authentication remain the existing configuration:
`PreparedMappingSharePointFolderPath`, falling back to `SharePointTargetFolderPath`.
No new cloud permissions or site are needed. Python 3.10+ is already required by
native evidence. The new producer uses the Python standard library and `-B`.
The previously deferred daily-trend qualification/absence wrapper settings are
not part of this change.

Once enabled, every normal run downloads and validates all three mappings before
replacing retained copies. A missing, corrupt or malformed third workbook cannot
fall back to an older local copy. Promotion failure rolls back the mapping set.
All calculation inputs are frozen and hashed in the existing capture lifecycle.

## Workbook contract

Sheet `Exceptions`, structured table `MailboxMigrationExceptions`, headers on
row 7 and data from row 8. Keep these nine required headers and their order:

1. Primary SMTP Address
2. Retain On-premises
3. Reason
4. Decision Owner
5. Decision Date
6. Review Date
7. Exchange GUID
8. AD Object GUID
9. Display Name

The legacy trailing `Source Reference` column is optional. The reader accepts
either `A7:I<last-row>` without it or `A7:J<last-row>` with it, and supplies an
empty source reference when absent. The current template has nine columns;
existing ten-column workbooks remain valid.
Deploy this reader before using a nine-column workbook on the preparation server.

Enter values, not formulas, in the decision table. Use `Yes` to retain and `No`
to withdraw a decision. Entirely empty rows are ignored. Dates must be actual
Excel dates or ISO `yyyy-mm-dd` text. The reader follows the actual table range,
including beyond 5,000 rows; extend Excel's optional formulas/dropdowns separately
when exceeding the template's 5,000-row assistance range. Macros and external
links are rejected. An optional source-reference cell records provenance, not identity.

## Identity and operator confirmation

All matching is scoped by the receipt-qualified tenant, organization, environment
and tenant ID. Existing native mailbox reconciliation supplies technical links.
Names, UPNs, aliases, substrings and wildcards never merge identities. Only trim
and case normalization apply to primary SMTP addresses. An invalid, unknown,
ambiguous or contradictory supplied GUID never falls back to SMTP.

An exact unique address without a GUID produces `Binding confirmation required`
and a `Candidate Identity`, rather than an automatically approved exception.
This is intentional: reusing an address must not transfer an old decision to a
different mailbox. Review the candidate against its native inventory identity.
Then either record the confirmed technical GUID in the private workbook or copy
approved bindings into a private registry based on
[`mailbox-migration-bindings.json.template`](../config/mailbox-migration-bindings.json.template).
The producer never writes back to the shared workbook or approves a registry.

The registry's `Identity` must exactly match the prepared batch. Each approved
entry in `Bindings` contains:

```json
{
  "PrimarySmtpAddress": "mailbox@example.invalid",
  "NativeIdentity": "local_mailboxes|11111111-1111-1111-1111-111111111111",
  "ConfirmedBy": "Example operator",
  "ConfirmedAtUtc": "2026-01-01T09:00:00Z"
}
```

Use actual reviewed values in the private file. The example is fictional and is
not an approval. Keep withdrawn bindings for identity continuity. A known binding
stays attached to its original native object after an address change. Reassignment
to another object becomes a review conflict. Duplicate/unconfirmed bindings or a
registry from another tenant/environment fail preparation. Registry bytes are
captured and hashed with the workbook; its source file is never modified.

## Counts and qualification

The planning inventory **N** comprises distinct detailed local UserMailbox,
SharedMailbox, RoomMailbox, EquipmentMailbox and LinkedMailbox observations. Its
identity rule matches the existing detailed migration inventory: Exchange GUID,
otherwise AD Object GUID. Missing or duplicate inventory identities fail the run.
RemoteMailbox recipients alone prove neither a local physical mailbox nor an
observed Exchange Online mailbox. Conflicting local/cloud hosting observations,
EXO-only retention decisions and out-of-scope types require review and are not
retained E. An explicit `No` on a confirmed, unique EXO identity is
`Withdrawn - verified EXO`: retained E=0 and review R=0. Confirmation requires
valid technical GUIDs or a confirmed private binding; SMTP-only candidates never
qualify. Hosting conflicts, reassigned addresses, unresolved competing remote
recipients, contradictory GUIDs and conflicting Yes/No rows remain in review.
Withdrawn rows retain their source references and documentation warnings.

- **Retained on-premises (E):** unique, conflict-free local planning identities
  with a valid `Yes`. Identical duplicate decisions count once, with all source
  row numbers and distinct reasons/owners retained.
- **Exceptions requiring review (R):** unresolved identity, decision or scope
  groups. Conflicting Yes/No rows never use first/last-wins precedence.
- **Remaining in migration scope:** N minus E, from the **same detailed snapshot**,
  types and domain scope. It is never historical aggregate minus Excel row count.
  It displays `Pending review` while an applicable unresolved group remains.

Missing reasons/owners, invalid documentation dates and overdue review dates are
documentation warnings. They do not cancel a valid matched Yes and are separate
from R. No dates, owners or comments are invented. Country is not inferred from
SMTP; domain comes only from inventory. An unlocalizable review blocks every
domain's net. A review localized to another domain does not block the selected
domain. Review totals include global unknown-domain issues.

A valid empty register gives E=0, R=0 and net=N. Missing/unverified enabled inputs
fail preparation and leave the previous prepared pointer untouched; they are
never cached fallbacks or verified zeroes. A negative net is a reconciliation
error. Old or disabled batches expose `Not configured`/`Not measured`; their
existing report data remains readable, without requiring a new batch solely to
load this optional feature. A partial new file pair fails closed.

## Report and provenance

The existing page and four internal bookmark tabs are retained. The Mailbox
migration view has two **Observed history** cards and three **Planning snapshot**
cards. Historical curves, crossover calculations and selected-period variations
remain unchanged. Period/Date filters affect history; planning cards keep their
explicit acquisition date. Domain applies to planning via explicit measure scope,
without adding Date relationships or retroactively applying decisions to history.

Domain detail keeps its 12 historical value/change fields and adds a separate
snapshot matrix plus an inspectable decision list. Unknown-domain reviews remain
visible even with a selected domain. The shortened tables scroll vertically.

Prepared outputs are `MailboxMigrationExceptions.csv` and
`MailboxMigrationExceptionSummary.csv`, governed by the optional extension
contract and the same batch validation, manifest, hash verification, atomic
pointer and transfer lifecycle. A Data Trust row and card tooltips carry source
qualification and provenance. The audit records the workbook hash, capture UTC,
local transported file modification UTC, binding hash, source-row count and
planning inventory acquisition date. It does **not** invent a SharePoint server
version or treat local file modification time as a decision date.
The planning date is the local inventory file receipt's UTC completion date;
the acquisition interval may span midnight or several domain reading dates.

Deployment is separate: review/push the code, upload the private workbook, enable
the private configuration, run preparation on the designated server, review the
new batch and any binding candidates, confirm identities, prepare the next batch,
then select the reviewed batch and refresh Desktop. Publication remains a separate
operator action. No preparation, SharePoint upload, refresh or publication is
performed by the implementation tests.
