# SmartInventory audit - lot 1

Audited and approved for repository publication on 2026-09-10. Scope:
shared CSV identity/persistence and the distributed orchestrator JSON writer.
This is not validation of all SmartInventory collectors or real tenant operation.
Integration uses upstream commit `a2ce29b59d2f11f5a6264ae8d4c0a16d8226e7a0`;
the approved source baselines remain unchanged. No operational deployment,
scheduled-task action, tenant collection or website change is included.

## Product baseline

The inspected checkout is `5d7341e81a4cacd014f0080406cb8d17d80b59f4`.
SmartInventory has no suite version declared in its README. Its active orchestrator
script is 1.5.10, with independently versioned modules and collectors. No suite
version, stable label or new collector version is introduced. No matching local
inventory tag was returned; the public GitHub releases page could not be retrieved,
so remote release absence is not claimed.

Approved versions follow the existing module patch conventions:

| Component | Baseline | Candidate | Change |
| --- | --- | --- | --- |
| SmartM365.Core | 1.0.47 | 1.0.48 | Identity conflict guard and terminating CSV serialization errors |
| Windows PowerShell 5 compatibility | 1.0.35 | 1.0.36 | Matching identity guard and dictionary identity support |
| Orchestrator distributed module | 1.1.3 | 1.1.4 | Propagate JSON directory/rename failures |
| Main orchestrator script | 1.5.10 | unchanged | No schedule, election or business algorithm changed |

The CSV functions retain their signatures. Existing callers do not require a new
API or a higher import floor; deploying the corrected behavior requires the new
module files and manifests together. The two shared modules are also used outside
SmartInventory: no other application's code is changed, but rollout must account
for that shared runtime impact.

## Reproduced and corrected defects

| ID | Reproduction on baseline | Correction | Contract impact |
| --- | --- | --- | --- |
| SI-01 | A row from another TenantKey, OrganizationKey, EnvironmentKey or TenantId is silently relabelled; mixed rows are homogenized when context is inferred | Reject every non-empty conflicting identity before writing; do not include identity values in the error | Valid single-context inputs keep their identity-first schema. Inconsistent inputs now fail rather than becoming mislabeled evidence |
| SI-02 | An IDictionary row carrying valid identity cannot supply missing context; dictionary conflicts are ignored | Read dictionary identity keys as well as object properties, using the existing case-insensitive identity comparison | Existing dictionary input support becomes consistent with object input; no business ID normalization |
| SI-03 | With ErrorActionPreference=Continue, a mocked Export-Csv writes a partial staging file and emits a non-terminating error; Core publishes that file | Use terminating errors for directory creation and both Export-Csv branches in the atomic writer | Last valid destination remains byte-identical on serializer failure; compatibility writer already stopped on Export-Csv errors |
| SI-04 | Hold a synthetic lease/claim file open against replacement; under Continue, the distributed writer reports an error but setters still acknowledge success | Make directory creation and final JSON rename errors terminating | Existing callers' catch paths receive persistence failure; ownership/election/retry algorithms unchanged |

No columns, CSV basenames, business calculations, schedules or operational JSON
were changed. This correction enforces the existing tenant-isolation contract.
It intentionally rejects conflicting identity input; it does not migrate or repair
client data, silently switch tenants, or deduplicate business records.

## Tests actually executed

`../Tests/Test-SmartM365InventorySharedOffline.ps1` extracts named function definitions
through the PowerShell parser into disposable modules. It never executes a
collector or tenant-context initializer. CSVs and lock files are synthetic and
created beneath a unique temporary root. SharePoint is a throwing mock; weekly
history is a no-op mock. Therefore these tests do not validate weekly history or
SharePoint behavior.

| Verification | Result |
| --- | --- |
| Baseline, same 34-case regression suite, PowerShell 7.6.6 | 15 passed, 19 failed as expected |
| Baseline, same suite, Windows PowerShell 5.1 | 15 passed, 19 failed as expected |
| Candidate regression suite, PowerShell 7.6.6 | 34/34 passed |
| Candidate regression suite, Windows PowerShell 5.1.19041.6456 | 34/34 passed |
| Existing Test-SmartM365OrchestratorDistributed.ps1, PowerShell 7 | Passed; synthetic capability/election/claim/lease/re-adoption scenarios |
| Existing Test-SmartM365OrchestratorManagement.ps1, PowerShell 7 | Passed; temporary config publication/conflict/history scenarios |
| Final signed regression suite, AllSigned, PowerShell 5.1 and 7 | 34/34 passed on each engine |
| Six changed/new PowerShell files, parsers 5.1 and 7 | Zero parse errors; zero non-ASCII characters |
| PSScriptAnalyzer 1.25.0, PowerShell 7 | Zero Error-severity findings; analyzer unavailable in the PowerShell 5.1 module path |
| Six local Authenticode signatures | Valid on this host; signed without external timestamp |
| Exact nine-file patch and original source hashes | Apply/whitespace checks passed; source hashes unchanged |

The 34 cases cover all four identity fields, mixed-context inference, dictionaries,
matching/missing identity, last-valid preservation, real local file locks,
multiline/quoted Unicode CSV values, explicit empty schemas, column order,
leading-zero string IDs, semicolon tenant-neutral exports, MAXITEMS isolation and
successful DATA-ALL/DATA-LAST byte parity. Same-host file-lock tests do not qualify
SMB failover or multi-server safety. No Desktop refresh or consumer calculation
engine was executed.

Run the new test only in a fresh offline PowerShell process, from the reviewed
checkout, using `-NoProfile -File SmartM365/Tests/Test-SmartM365InventorySharedOffline.ps1`.
An optional `-ResultPath` writes the assertion report. Follow the host's normal
signing policy; do not change global execution policy to run tests.

## Configuration, prerequisites and permissions

| Area | Source contract / prerequisite | Audit boundary |
| --- | --- | --- |
| Tenant configuration | Config/SmartM365-TenantContext.ps1 merges Config/SmartM365.global.local.json with Config/Tenants/<ProfileKey>.local.json, then per-script overrides/token resolution | Read source/templates only. Loader can create/add missing runtime JSON keys; it was not executed |
| Identity | ProfileKey selects a file; the loader validates OrganizationKey, EnvironmentKey and TenantKey=OrganizationKey-EnvironmentKey | Do not assume profile `test` is the exported TenantKey. Explicit output overrides still require isolation review |
| Outputs | Effective DataAllRootPath, LatestCsvFolderPath and LogAllRootPath; default roots isolate tenants, with script Output fallback | No configured customer path was read, tested or changed |
| Runtimes | PowerShell 7 for resident scheduler/cloud scripts as declared; Windows PowerShell 5.1 plus Exchange/AD tooling where declared | Copy full dependency folders; use manifest imports, not a loose PSM1 |
| Graph | Per-job RequiredGraphAppRoles in the manifest and per-collector SDK/preflight requirements | The map lists declared roles; it is not proof of effective least privilege or consent |
| Usage | Reports.Read.All; source ReportRefreshDate may lag collection time | No service report queried |
| Exchange Online | Exchange.ManageAsApp and supported read RBAC/Entra role; certificate/private-key access for service account | Runtime rights differ from interactive setup administration; no EXO session created |
| AD / Exchange on premises | AD module, directory read access and equipped Exchange shell/roles | No AD, CIM, remote shell or endpoint operation |
| SharePoint upload | Optional Sites.Selected with explicit target-site write grant in the repository design | Upload is separate from collection; actual consent/grant was not inspected |
| Notifications | Optional Graph Mail.Send/SMTP and Teams webhook configuration | Sensitive sender/webhook values remain local. No mail, Teams or upload call |
| Scheduling | Existing installer uses dedicated service account; code/config/state/private-key ACLs matter | No scheduled task queried, installed, changed or started |

The shared REST token cache is keyed by application, tenant and certificate in
source. This is not authentication qualification. Protect private keys, runtime
configuration, logs, CSVs, cached reports and generated child commands. The
orchestrator logs its decoded arguments; avoid secrets in job Arguments. Detailed
Graph errors and notifications can expose tenant data: source review of redaction
and error transport is still required.

Microsoft's API documentation distinguishes report requests from ordinary reads:
`POST /deviceManagement/reports/exportJobs` lists application ReadWrite permissions.
Do not remove those permissions solely because the script is called an inventory.
Assess each actual endpoint first. See [export job permissions](https://learn.microsoft.com/en-us/graph/api/intune-reporting-devicemanagementexportjob-create?view=graph-rest-1.0),
[Exchange app-only authentication](https://learn.microsoft.com/en-us/powershell/exchange/app-only-auth-powershell-v2?view=exchange-ps),
and [selected SharePoint permissions](https://learn.microsoft.com/en-us/graph/permissions-selected-overview).

## Remaining priority work

| Priority | Finding / open boundary | Required next evidence |
| --- | --- | --- |
| High | Update-RunningJobs logs a still-running process after timeout but proceeds to Complete-JobRun, which releases its lease; source inspection only in this lot | Mock failed process-tree termination and prove no overlap/retry before correcting lifecycle behavior |
| High | Atomic CSV writes are per-file. Publish-SmartM365Csv writes history then latest separately, with no common writer lock or cross-file commit | Inject failure between writes, concurrent older/newer runs and readers. Agree any new transactional publication protocol before changing the data contract |
| High | Export validity depends on collector completeness and validation rules. Absence of a rule can allow publication unless the strict global flag is enabled | Per-collector missing page/partial API response tests, including empty-success versus empty-failure |
| High | Direct Export-Csv, Copy-Item, streaming and legacy exporters remain alongside shared helpers | Audit each path and its error/last-valid behavior. No universal protection claim |
| High | Claims and leases use shared files; takeover, PID reuse, server clock skew, crash recovery and mixed-version cluster races are not exhaustively tested | Deterministic competing-writer tests before a separately approved deployment; no change to election weights or algorithm |
| High | Main success evidence can accept startup context patterns; exit zero alone does not prove complete exports | Mock failed/partial collector exit paths, required success markers and dependency/retry status propagation |
| Medium | Shared retry helper recognizes transient statuses and selected 409 bodies, but caps numeric Retry-After and does not parse HTTP-date values; MaxAttempts is not range-validated | Mock headers/statuses, batch subresponses, exhaustion and zero attempts. Establish bounded retry behavior without silently losing pages |
| Medium | Weekly history copies/manifest updates are separate from latest publication; UpdatedAt is a write time | Test concurrent same-week copies, interruption, manifest completeness, retention and last-valid recovery |
| Medium | CSV date columns differ by collector; the shared writer does not inject a universal collection timestamp | Preserve source/report timestamps; check oldest-row freshness, timezone/future/missing dates in each producer/consumer contract |
| Medium | Permissions, interactive sessions, token/audience checks and secret redaction vary across collectors | Mock wrong-tenant/wrong-audience auth, missing permission and redacted diagnostic output; no live auth |
| Medium | Duplicate and missing business identifiers have source-specific meanings | Reconcile compact licensing/app relations and legacy aliases with consumer mappings before any change in grain or algorithm |

Microsoft requires following returned next links for pagination and respecting
Retry-After; batch HTTP 200 can contain individually throttled subrequests. These
are audit criteria, not validated collector behavior in this lot. Sources:
[paging](https://learn.microsoft.com/en-us/graph/paging),
[throttling and batches](https://learn.microsoft.com/en-us/graph/throttling).

## Consumer and distribution boundary

SmartFinOps Workplace reads DATA-LAST using its source contracts; Dashboard reads
90 selected latest CSVs and weekly history using its mapping files. The corrected
writers retain the current leading TenantKey, OrganizationKey, EnvironmentKey,
TenantId columns and all business column names/order. Synthetic serialization
checks support this limited compatibility claim; neither consumer was modified.

Distribution should remain repository-based: ship the reviewed internal modules
and their manifests together with required dependency folders. Do not create a
release ZIP for every collector. No suite release or stable promotion is proposed.
Obsolete signature blocks were replaced with local signatures using the existing
publisher certificate, without an external timestamp. All six PowerShell candidate
files verify as Valid on this host, and the signed test passes under AllSigned on
both engines. This does not qualify target-host trust, downloaded GitHub LF bytes,
timestamped longevity or a real Enforce deployment. No trust store was changed.

The existing website's SmartM365 inventory cards link to the repository and make
no suite-version/stable claim. No EN/FR/IT/ES/DE/AR site change is necessary for this
internal correction lot; the proposed external site delta is zero files. A future
guide or wider product announcement must be prepared and reviewed separately.

Next required lot: reproduce and correct orchestrator lifecycle failure handling
and shared writer/weekly-history races, then audit Graph users/domains/licensing
and their pagination/partial-result paths. Continue exclusively with synthetic
data. Any business algorithm or data-contract migration requires prior agreement.
