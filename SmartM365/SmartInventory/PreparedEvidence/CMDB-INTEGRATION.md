# CMDB orchestrator integration candidate

This is preparation for a separately approved deployment, **not activation**.
`cmdb-orchestrator-integration.json.txt` is a review-only specification, not a
replacement jobs manifest. It is not referenced by the resident orchestrator,
the pipeline launcher, or automatic template merging. No active schedule,
capability, private configuration or Intelligence dependency is changed.

## Application freshness

DiscoveredApps 1.30 adds the explicit `-FreshDeviceDetails` switch. It requires
`-DeviceDetailMode All`, no MaxApps/MaxItems limit and no `-ResetResume`.
The default collector behavior is unchanged without the switch.

Fresh mode rejects even an otherwise compatible checkpoint. Before creating a
new partial export it preserves the previous checkpoint **byte for byte** through
the existing owned-checkpoint mechanism. The previous partial/completed exports
are left in place. Every app remains eligible for actual Graph acquisition;
All mode already prohibits previous relation-cache reuse. The first new atomic
checkpoint can replace the current pointer without destroying its saved bytes.
An interrupted fresh run keeps its own checkpoint but a subsequent **fresh** run
starts over again; do not silently turn resumed old relations into fresh proof.
Normal resume remains available for the original inventory use case, but its
reused evidence cannot qualify the strict CMDB receipt.

Existing per-app pagination, retries, relation-count reconciliation, paired
Summary/AppDeviceRelations publication checks and full-scope receipt gates are
not relaxed. Successful local tests do not prove live Graph completeness or a
feasible production duration. This option may increase API calls and run time.

The existing application job is weekly. CMDB receipts require acquisition no
older than 48 hours and intervals no longer than 48 hours. A daily preparation
cannot therefore be guaranteed from that weekly job; a multi-day resumed run
does not solve the problem. Before activation, qualify a fresh full run, then
approve cadence/timeouts or explicitly redesign freshness requirements. Do not
change those gates merely to pass a preparation run.

## Exchange local acquisition and recipient quality

Exchange local 1.53 qualifies full acquisition separately from recipient health.
Every discovered domain is queried at its domain DN, including the forest-root
containers (not only first-level OUs). Required queries are terminating on error.
Their projections must preserve the exact native object-GUID population; missing,
empty, repeated or substituted identities fail. The RemoteMailbox projection is
checked against its unrestricted native query as well.

Known recipient validation findings and unavailable statistics, quotas or mobile
fields remain in the native issue CSV. Five additive columns declare Severity,
CollectionImpact, BlocksCmdbQualification, ObjectGuid and NativeRecordRetained.
An explicit category/operation
allowlist distinguishes RecipientDataQuality and FieldUnavailable from blocking
acquisition failures. Unknown warnings are retained and blocking; no wildcard
exception or fabricated zero replaces missing evidence. Recipient findings are
nonblocking only when their exact native Identity/DN/GUID resolves to one returned
object; absent or ambiguous warning targets remain blocking. Complete query and
projection evidence remains mandatory even when all issues are nonblocking.

SMTP conflicts are reported without deleting or rewriting native records; they
block qualification because address-based reconciliation is ambiguous. An empty
SMTP remains a retained object-GUID record, not a duplicate blank address or a
proven EXO match. The receipt includes aggregate qualification notes; the issue
CSV keeps the detailed evidence. Existing failed receipts are never repaired or
stamped retroactively: a new acquisition is required.

The base mailbox/RemoteMailbox CSV schemas and Intelligence history keys are
unchanged. Supplemental field failures still appear as Error/N/A in the existing
exports, never as measured zero. This preparation does not activate jobs or
change the existing application's schedule, timeout or arguments.

## Proposed jobs

Reuse the existing resident orchestrator and pipeline. Add only:

- `M365-WorkplaceScope-Inventory`: native Entra direct members and five Intune
  policy families. SDK beta Groups and DeviceManagement modules are required.
  Hidden memberships additionally require Member.Read.Hidden. Its default
  configured external actions remain disabled pending transport approval.
- `CmdbEvidence-Prepare`: local-only current preparation into the sibling
  DATA-POWERBI-CMDB. It requires Python 3.10+ under the elected execution account;
  SharedRuntime capability alone is not proof that Python is installed. It does
  not collect, upload, notify, create history, or refresh/switch Power BI.

The candidate explicitly names the 17 producer dependencies. EXO uses the full
mailbox job rather than its permissions-only sibling. Dependency success alone
is not source proof: the preparer independently verifies every required receipt,
scope, identity, acquisition interval, file hash and logical row count.

Both proposed jobs remain Enabled=false. Schedule, TimeoutMinutes and
EstimatedDurationMinutes are intentionally absent: they must be qualified and
approved, not invented. ConditionalGraphAppRoles is review metadata, not an
orchestrator schema extension. Do not copy this specification directly into the
live manifest. Disabled jobs cannot be requested even through `-Job`.

The argument addition is also review-only. Before applying it, inspect existing
effective arguments/configuration, reject limits/destructive ResetResume, and
resolve any existing DeviceDetailMode instead of appending duplicate arguments.
Preserve all unrelated arguments and settings. Do not mutate the shared
Intelligence job's dependency list, generators or explicit historyKey contracts.

## Qualification gates

1. Review the exact effective central manifest and prerequisites on the execution
   hosts, with no collection. Choose proposed schedules, timeouts and transport.
2. With separate approval, deploy the reviewed source changes and candidate job
   entries to private configuration; validate planning before enabling anything.
3. With separate approval, execute actual full-scope acquisition, inspect the 17
   receipts and compare native populations/relations. Retained no-DNS AD objects
   stay in the existing Windows-workstation/unknown-OS scope; they are not proof
   of deployed or active PCs. Verify their GUID/SID and separate count.
4. Validate inputs, then prepare the candidate CMDB output locally and inspect
   all 46 tables/joins. Keep existing CMDB and Intelligence outputs unchanged.
5. Approve publication/synchronization and the Power BI model/report migration
   separately, including all ten pages and current-only manifest validation.

Git commit/push is a separate approval; no package is created for this change.

## Offline verification

Use the repository's synthetic tests:

```powershell
pwsh -NoProfile -File .\SmartM365\Tests\Test-SmartM365DiscoveredAppsFreshOffline.ps1
pwsh -NoProfile -File .\SmartM365\Tests\Test-SmartM365CmdbIntegrationCandidate.ps1
pwsh -NoProfile -File .\SmartM365\Tests\Test-SmartM365ExchangeLocalQualificationOffline.ps1
powershell.exe -NoProfile -File .\SmartM365\Tests\Test-SmartM365ExchangeLocalQualificationOffline.ps1
```

The integration test merges the candidate into the committed jobs template **in
memory only**, temporarily enabling candidates only in that synthetic copy to
exercise the actual dependency selector. It verifies producer coverage, disabled
real candidates, unchanged original jobs and Intelligence historyKey contracts.
It does not validate real permissions, schedules, SDK access or tenant CSVs.
