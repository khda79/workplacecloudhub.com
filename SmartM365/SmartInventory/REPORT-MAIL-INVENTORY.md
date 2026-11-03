# SmartInventory report mail: implementation inventory

Status: shared sender and resident worker connected locally; configuration disabled by default. Offline qualification only.
This inventory describes the repository sources, not the deployed production version.

## Shared transports

PowerShell 7 `SmartM365.Core` and the Windows PowerShell 5.1 compatibility module
both expose `SendEmailHtmlReport` and `Send-SmartM365GraphMail`; the former sends
through Graph or SMTP, including Graph-to-SMTP fallback. PowerShell 7 also has
`Send-SmartM365Mail`. `SmartM365-MailMaintenance.ps1` provides recipient policy.
The shared module is used outside SmartInventory, so a new scheduling policy must
establish its caller scope rather than affect all SmartM365 mail implicitly.

The `Auto` purpose resolves to `Report` unless the configured error recipient or
global `ScriptFailed` provides evidence of an error. Administrators-only routing
returns before that resolution, and `Resolve-SmartM365UnifiedMailRoute` does not
return the resolved purpose. Resolve classification independently of recipient
routing and never assume a subject or priority identifies an error.

## Identified call sites

| Source | Current purpose | Intended treatment / work to verify |
| --- | --- | --- |
| M365Inventory/Teams/SmartM365-Teams-Inventory.ps1 | Auto | Daily summary is deferred by the shared sender; the daily marker records durable submission. Delivery is tracked in the queue. |
| M365Inventory/SharePoint/SmartM365-SPO-Inventory.ps1 | Auto | Daily summary and MailOnly use the shared sender; their marker records durable submission. |
| M365Inventory/IntuneInventory/WindowsUpdate/SmartM365-WinUpdate_Status_From_Intune.ps1 | Auto via `Send-CoreSmartM365Mail` | Fatal error explicitly uses Error; the summary remains a report. A successful summary return means durable submission or transport acceptance. |
| M365Inventory/Licensing/SmartM365-Licences-Inventory.ps1 | Report / Error | Focused summaries explicitly use Report; fatal error explicitly uses Error. Daily/domain gates retain queued submissions with the Queued status; the queue tracks actual delivery. |
| PreparedEvidence/SmartM365-WorkplaceEvidence-Prepare.ps1 | Report / Error | Completion report is explicit; failure is explicit Error. |
| ExchangeInventory/Quarantine/SmartM365-EXO-QuarantineMessages-Report.ps1 | Auto / Error | The workbook email is a report; the global failure is Error. The report log currently says sent immediately after the helper returns. |
| ExchangeInventory/Hybrid/FreeBusy/SmartM365-Exchange-Hybrid-FreeBusy-Inventory.ps1 | Auto via imported alias | `Send-FreeBusyHtmlReport` is called for the generated Free/Busy report. No separate failure-mail call was found in this file. |
| ExchangeInventory/OnPremises/ServersAndStorage/SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1 | Auto via imported alias | `Send-ServersAndStorageHtmlReport` sends the generated report in the normal and regeneration paths. |
| ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1 | Report / Error through wrapper | Daily summary gate records submission; the queue records actual delivery. |
| SharePointInventory/OnPremises infrastructure and content collectors | Report | Summary wrappers explicitly pass Report and use daily gates; their markers record submission; the queue records delivery. |
| ActiveDirectoryInventory/SmartM365-ActiveDirectory-Inventory.ps1 | Report / Error through wrapper | Daily summary and duplicate-identity notification both use the Report default and distinct last-sent markers; the latter explicitly uses Operational delivery while preserving its Report recipient routing. Fatal errors explicitly use Error. |
| ActiveDirectoryInventory/SmartM365-ActiveDirectory-HealthCheck.ps1 | Report / Error through wrapper | Health-check report and critical failure use separate purposes. |
| M365Inventory/SyncHealth/SmartM365-AzureADConnect-SyncHealth-Inventory.ps1 | Report for WARNING, Error for ERROR | WARNING is an operational alert, not necessarily a scheduled report. Explicit Operational delivery keeps both levels immediate without changing recipients. |
| Orchestrator/SmartM365-Inventory-Orchestrator.ps1 | Defaults non-error to Report | Execution summary is a report. Peer/prerequisite recovery, runtime update, disk-space warning and recovery, stalled job, certificate, manifest, Authenticode and fatal notifications are operational and immediate. Individual job-result notifications are treated as operational and remain immediate; the consolidated execution summary is the deferred report. Maintenance transition uses Maintenance. |
| Maintenance/SmartM365-OrchestratorServerAcceptance.ps1 | Maintenance | Immediate. |
| Explicit error calls in Devices, Entra devices, Active Users, Autopilot, Discovered Apps, RBAC, Endpoint Analytics, Verified Domains, EXO mailboxes, accepted domains and migration | Error | Immediate; retain existing recipient restrictions. |

## Transport bypass audit

Static searches for `Send-MailMessage`, `Send-MgUserMail`, `/sendMail`,
`SmtpClient`, and the shared send functions found a local `Send-HtmlMail`
function in the orchestrator, but no invocation of that function in the same
file. Shared transport implementations contain Graph and SMTP calls as expected.
No active direct bypass is established by this preliminary search. Imported aliases `Send-CoreSmartM365Mail` and `CoreSendEmailHtmlReport` route through
the shared sender and are covered. This static inventory covers the checked-in
collector call sites, not arbitrary future dynamically constructed commands.

The source search includes `Send-CoreSmartM365Mail` and
`CoreSendEmailHtmlReport` in addition to the unprefixed helpers. SharePoint
on-premises collectors, AD inventory, local Exchange mailbox inventory, Teams,
SharePoint Online and Licensing each have their own daily state or gate, so queue
acceptance cannot safely be treated as an ordinary successful transport return.

Delivery classification must be separate from `MailPurpose`: changing an
operational warning from `Report` to `Error` would also change recipient routing.
For example SyncHealth WARNING and orchestrator disk-space WARNING are currently
Report for routing, yet should be delivered immediately. A safe migration keeps
the existing recipient purpose and passes an independent delivery class.

## Design constraints before implementation

- Durable queue data and attachments require protected access, atomic publication,
  integrity checks and cluster-wide coordination; no secrets belong in logs.
- A queue acceptance result is different from transport delivery. Existing
  Boolean return values, daily markers, orchestrator checkpoints, and callers
  that discard return values must be reviewed individually.
- Apply live maintenance and administrator-recipient policy at actual delivery;
  handle tenant isolation, unavailable controls and configuration changes safely.
- A transport acceptance immediately followed by a crash before recording success
  can lead to a duplicate; do not claim exactly-once delivery.
- Test both PowerShell editions, Graph, SMTP and fallback, restarts, failed
  writes, cross-server contention, attachments, DST, and invalid configuration
  with synthetic data and mocked transport before any live qualification.

## Implemented delivery contract

`ReportMailSchedule` in effective global/tenant configuration defaults to disabled.
Its fields are Enabled (Boolean), NotBefore (HH:mm), TimeZone (Windows time-zone ID),
QueuePath (defaults to `{{WorkspaceRootPath}}\Tenants\{{TenantKey}}\MAIL`, existing dedicated directory),
and TrustedSids (approved service/administrator SIDs). Missing, blank or inheritance
markers for QueuePath use this default; an explicit tenant path is preserved.
The queue must be provisioned separately with
protected NTFS permissions. Untrusted queue/file owners, linked paths and
unauthorized queue/file access rules are refused. Parent ACLs remain under the
existing share administration policy and are not an admission requirement. Local SYSTEM and Administrators are trusted;
all other required owners/accounts must be explicitly listed. Share availability and
service-account access still require server qualification before activation.

DeliveryClass is separate from MailPurpose: it affects scheduling, not recipients.
Error and Maintenance purposes are immediate. SyncHealth notifications, AD duplicate
identity notices and orchestrator operational events explicitly remain immediate.
The orchestrator execution summary is deferred. Other inventoried collector reports
pass through SendEmailHtmlReport (including imported aliases). The Windows Update
fatal notification now explicitly uses Error. This gate is scoped to SmartInventory
call stacks; other SmartM365 consumers retain their current behavior.

Core 1.0.122, compatibility manifest 1.0.71 and orchestrator 1.5.88 must be deployed
together. The resident processes up to ten eligible records per tick, only while the
option is enabled and the current local time permits delivery. A stopped resident
leaves the queue intact. Disabling the option holds existing queued mail; it does
not delete or release it. Direct collector runs use the same queue but need a resident
for subsequent dispatch. Graph/SMTP routing policy is re-evaluated at actual send.

A successful collector submission means transport acceptance OR durable queue
acceptance. Existing date-only daily markers serve as submission guards (their legacy
LastSent filenames are retained for compatibility). Licensing persists Queued when
appropriate. The queue is authoritative for actual delivery status; it clears message
content after acceptance, keeps a deduplication receipt and retries failures after five
minutes. Pending failures are retained, never silently deleted. Payloads and attachments
are snapshots, not references to files that a collector may overwrite or delete.

No exactly-once guarantee: a crash between remote acceptance and recording Sent can
repeat a message; a crash between enqueue and updating a collector's submission marker
can also repeat submission. Persistent receipts and locks prevent ordinary worker replays.
The checksum detects accidental payload damage, not malicious access by a trusted writer.

## Offline validation

Test-SmartM365DeferredMailIntegrationOffline extracts the actual PS7 and PS5 send
entrypoints and uses the actual admission/queue/worker with mocked transports and
synthetic permissions. It covers cutoff, attachment snapshots, Graph, SMTP, Graph
failure with SMTP fallback, immediate errors/operations, disabled behavior and replay.
It also checks that the real ACL verifier refuses an unprotected directory. Separate
queue tests cover contention, retry, tenant mismatch, invalid payload and batching;
schedule tests cover daylight-saving transitions. This does not qualify a real SMB
share, real Graph/SMTP delivery, or a running production orchestrator.
