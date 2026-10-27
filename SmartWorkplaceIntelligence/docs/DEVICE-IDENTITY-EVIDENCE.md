# Intune device identity evidence

Native contract 0.5.0 separates a persistent managed/hardware correlation conflict
from an acquisition failure. Device inventory 1.21 requires Core 1.0.109 and the
coordinated native readers. Existing completed receipts keep their prior strict
behavior. This change does not establish a cause for a Microsoft Graph discrepancy.

Before any Devices export, the producer compares every managed-device ID against
the independently acquired hardware observation. A consistent row causes no
additional request. A divergent device receives at most three pairs of independent
reads: the normal managed-device SDK observation and a targeted hardware GET with
explicit properties. Pauses before the second and third pairs are two and four
seconds. These reads do not invoke the initial hardware helper's batch retry loop.
Transport failures, 404 responses, malformed identities, invalid hardware values,
duplicate IDs and incomplete acquisition remain terminal failures. They cannot
become an identity quarantine and cannot promote a source set.

On convergence, the complete newly observed managed and hardware rows replace the
initial rows. The Windows view uses that same new managed object, including its
other attributes. Observation timestamps remain those of the independent reads.
On persistent disagreement, the published rows retain the initial observations;
the initial and all three subsequent pairs are preserved in a separate artifact.
No contested correlation identifier is copied into another observation to force
agreement.

`Intune_DeviceIdentityEvidence.csv` is required by the new producer's receipt even
when it has no rows. Registry `SemanticFiles` extends that receipt's `RequiredFiles`
without changing the legacy source set. Each evidence row contains tenant, run,
device, resolution, full observations and a SHA256 of its exact UTF-8 JSON payload.
The shared receipt binds the complete CSV hash and row count to the same run.
Older readers reject the extended required-file declaration. New readers reject
missing, extra, duplicate, foreign, incomplete, altered or inconsistently resolved
evidence. Every published conflicting row must be exactly covered by a complete
verified quarantine. An undeclared discrepancy still fails.

The capture selects the artifact from the current receipt. Prepared artifact
transport retains it as `Source-Intune_DeviceIdentityEvidence.csv`, with a checked
hash, tenant, run and semantic version in the artifact manifest. Native audit
summaries contain counters and that reference, not the embedded observations.
Generic legacy prepared-artifact contract hashes remain explicitly compatible;
retired producers and retired current exports are not restored.

A conflicting device remains a separate tenant-scoped `intune:<ManagedDeviceId>`
entity. DeviceSource, the managed-device bridge, hardware, applications and
directly observed compliance retain its technical identity and Intune attributes.
Entra/AD device links and user/country attribution are withheld for this entity.
Independent Entra and AD observations remain available under their own identities.
This can change the unified entity count when a previous merge is removed; it does
not establish a physical asset count.

Association and source selection statuses expose `Identity conflict`. Quality
findings, SourceHealth and hardware coverage contain explicit warnings and counts.
Windows/readiness/system/migration readers exclude contested joins and name
fallbacks, expose indeterminate attribution and count unreconciled observations.
Data Trust marks usable inventory with conflicts as `Review required`, with the
artifact reference and fingerprint. Dependent reconciliation percentages remain
blank instead of asserting a complete qualified denominator. Direct Intune
measurements remain distinct from facts that require a directory link.

All validation is offline with fictitious identities. It covers zero additional
reads on consistent input, convergence on attempts one through three, persistent
conflicts, acquisition errors, invalid proofs, source rechecks, source/artifact
capture, staged publication, tenant-scoped counts and Exchange/license equivalence.
Offline acceptance does not qualify real Graph responses, server deployment,
Power BI model integration, refresh, rendered filters or Power BI Service.
No private master or thematic report is modified by this implementation. Existing
model bindings consume the existing status and quality fields; any further model
authoring or installation requires its own reviewed scope. Changed PowerShell
files must be signed and checked before future distribution.
