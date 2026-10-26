# Smart Inventory

Inventory scripts that can feed Power BI datasets, operational reports, CSV exports, SharePoint publishing, and other downstream consumers.

## Reliability and distribution

SmartInventory is a collection of independently versioned scripts and shared
modules; this README does not assign a suite version or stable qualification.
Shared CSV persistence, distributed orchestration, Graph collection and the
bounded AD/Exchange/readiness families have focused synthetic regression coverage.
Those results do not qualify Microsoft Graph, a tenant, on-premises systems or the
whole SmartInventory suite. Detailed audit maps, evidence and publication records
are intentionally retained outside the public repository.

Use a reviewed repository checkout with the required `Config/`, internal module
manifests and templates. A single copied collector is not a self-contained package.
No per-script ZIP release is required. Operational JSON, credentials, CSV exports
and logs must stay outside Git and public distribution.

The shared exporters reject conflicting source identities before writing
files. Existing identity-first CSV columns, names, delimiters and business fields
are preserved. A failed serialization cannot replace the previous valid CSV, but
this does not establish a transaction across all files or prove collection
completeness. Consumers must still check source dates, schemas and tenant identity.

## Current source receipts

Every producer publishing reporting CSVs to `DATA-LAST` now writes one
`SmartInventory_<producer>.current.json.txt`, with `Owner=SmartInventory-SourceReceipt`
and contract version 1.2. This metadata is shared evidence, not exclusive ownership
of the CSV by CMDB. One receipt lists all current-run published CSVs; there is no
receipt per CSV, no new history and no folder scan that qualifies older files.
The receipt is `Running` before acquisition and becomes `Completed` only after
successful required publication and row/hash/tenant checks. Failed runs preserve
native CSVs but invalidate completion evidence. Caches, temporary files, archives
and the remediation-script run bundles are not `DATA-LAST` publications.

The 17 CMDB source producers keep their explicit consumer-scope checks. The 30
other `DATA-LAST` producers declare their required and optional outputs in
`Modules/SmartM365.Core/SmartM365-SourceReceipts.json.txt`. Additional current-run
exports are also recorded; absent optional outputs are not replaced by old CSVs.
Effective selection parameters are included where applicable. For these producers,
`ScopeQualification=ConfiguredOutputsOnly` and `IsPartialInventory=null` mean
that whole-tenant exhaustiveness is **not established**, not that it succeeded.
The SharePoint on-prem Content producer publishes observed rows and a separate
`SharePoint_OnPrem_CollectionCoverage.csv` when coverage gaps occur. Its source
receipt verifies current file identity, rows, and hashes only; the coverage CSV,
run log, and warning exit code identify an incomplete collection. The generic
configured-output receipt does not qualify farm-wide completeness.
No receipt claims `FullInventoryQualified=true`. Derived reports remain derived;
creating their receipt does not refresh their upstream inputs.

Read-only and MAXITEMS modes do not overwrite canonical receipts. Core 1.0.72
and Windows PowerShell 5 compatibility 1.0.50 are required by the newly integrated
collectors. The common evidence runtime requires Core 1.0.72. CMDB preparation
accepts strict legacy 1.1 receipts during rollout and qualified shared 1.2 receipts;
existing metadata is never restamped or repaired from CSV timestamps. Intelligence
CSV names, columns, prepared contracts, generators and history keys are unchanged.
Qualify the next real acquisitions on the collection machines after deployment;
the offline tests do not establish production freshness or completeness.

Data quality takes precedence over runtime. A page limit reached by the Backup
Protected Mailboxes collector now fails the run before replacing the latest
CSV; use `-MaxItems` for a bounded test instead of treating `-MaxPages` as a
partial-export option. Windows Update reports the source and coverage of
`DaysSinceLastStatus`; the field remains blank when the export has no suitable
device-status date. A policy modification date is not substituted for it.

## Daily capture retention

Core 1.0.92 adds receipt-qualified daily capture retention for eight producers:
M365 UserActivity, Copilot Usage, M365 WorkplaceScope, Intune Configurations,
Intune Managed Apps, Licences, License Pricing, and Windows Autopatch Alerts.
Other collectors keep their existing retention behavior.

After a successful source receipt, while its producer lock is still held, the
policy keeps the last complete published set for each of seven calendar days
(today plus six preceding days). Completion dates use Europe/Paris, including
daylight-saving transitions. A later failed run or a smaller same-day export
selection cannot displace a complete set. Completion qualifies the configured
exports; it does not establish whole-tenant collection completeness.

The policy reuses matching current-run timestamped CSVs. When a receipt member
has no matching timestamped copy, it creates a byte-verified copy under
`DailyCaptures/<RunId>/`. This matters for collectors that previously retained
only some of their exports. CSV schemas, current source receipts and weekly
snapshots are unchanged. Capture metadata records identity, completion date,
exact paths, sizes and SHA256 under `DailyCaptureRecords/<producer>/`; metadata
records are kept after their payload is retired. No archive is created.

Per-file count/age cleanup is bypassed only inside the opted-in output root.
For Usage and Copilot this root is the parent of the run folders, rather than
the current run folder. Whole-set retention runs only after receipt completion;
it verifies retained captures and every candidate before deletion. If SharePoint
upload is enabled, captures and their metadata must be uploaded and read back
with matching SHA256 first. Remote deletion uses the inspected item ID and
`If-Match` ETag; a remote refusal stops local deletion.

`WeeklyHistory`, `DATA-LAST`, `Corrections`, source receipts and manifests are
excluded. Unindexed legacy CSVs are preserved: old run dates alone are not
completion proof, and this change does not infer or repair old capture records.
Metadata corruption, linked paths, missing/changed retained files, locked files,
or uncertain remote results stop further cleanup and produce a warning. Deletion
of a multi-file set is not an atomic transaction; inspect the retained metadata
and run log after an interrupted or partially completed cleanup.
An output root containing a filesystem link disables both daily retention and
legacy cleanup with a warning, while allowing collection to continue.

Offline regression coverage is available through
`Tests/Test-SmartM365DailyCaptureRetentionOffline.ps1`. It uses temporary
synthetic data and mocked SharePoint only; deployment and real server/SharePoint
qualification remain separate.

## Organization

- `ActiveDirectoryInventory/`: Active Directory inventory and reporting.
- `ExchangeInventory/`: Exchange Online and Exchange on-premises inventory.
- `M365Inventory/`: Microsoft 365, Entra, Power BI, and Microsoft Fabric inventory.
- `SharePointInventory/OnPremises/`: independent SharePoint Server 2016/2019 infrastructure and content collectors; manual farm-server execution, no orchestrator job or automatic upload.
- `M365Inventory/IntuneInventory/`: Intune inventory, Windows Update reporting, Autopilot, RBAC, applications, and remediation export utilities.
- `M365Inventory/IntuneInventory/EndpointAnalytics/`: standard Endpoint Analytics score, startup, app reliability, and work-from-anywhere exports without Advanced Analytics.
- `M365Inventory/Security/`: Secure Score, authentication-method registration, and Conditional Access configuration evidence.
- `M365Inventory/IntuneInventory/Security/`: Defender agent and firewall-health reports exported from Intune without changing device or policy state.
- `ExchangeInventory/BackupProtection/`: Microsoft 365 Backup protection evidence for mailboxes, SharePoint sites, and OneDrive accounts.
- `M365Inventory/Licensing/SmartM365-LicensePricing-Inventory.ps1`: normalization of governed license prices (`M365_License_Prices.csv`) and optional Windows 10 ESU prices (`Windows_ESU_Prices.csv`). The versioned France public-list baseline `M365_License_Prices_Public-France.csv` (every row estimated) is merged with an optional private override; see [License pricing](#license-pricing).
- `SmartInventory.cmd`: the single entry point (see [Running jobs and manual launches](#running-jobs-and-manual-launches)).
- `Maintenance/Remove-SmartM365-ExpiredTimestampedCsv.ps1`: one-time cleanup of the timestamped run copies that WeeklyHistory already covers (see below).
- `Maintenance/Import-SmartM365-LegacyDailyStats.ps1`: one-time import of legacy AD and Exchange on-premises daily statistics into the SmartInventory daily CSVs (see below).
- `Maintenance/Invoke-SmartM365-ServerDataCleanup.ps1`: complete tenant DATA inventory with preview and verified server-only cleanup. It reports every file, groups decisions by area, preserves compact Discovered Apps summaries, keeps the prepared-batch history unless `-RetireLinkedPreparedBatches` is passed (then with audit receipts), retires old orphan atomic-write temporaries, and leaves unproven categories for review.
- `Launchers/` and `LaunchersByOrchestrator/`: one-line aliases of the former `.cmd` launchers, kept for one version (see below).

## Running jobs and manual launches

Every job and manual launch is declared in the catalog `Orchestrator/Orchestrator-Jobs.json.template`: the `Jobs` array (orchestrator jobs) and the `Launches` array (manual entries, `"Manual": true`, never scheduled or requested). `SmartInventory.cmd` starts Windows PowerShell 5.1 with `Orchestrator/Invoke-SmartM365InventoryJob.ps1`; the engine starts PowerShell 7 when the entry needs it, so it also works on Exchange servers.

```text
SmartInventory.cmd                                   menu: choose a domain, then an entry (number or name), then run,
                                                     request or show the command; A shows the advanced entries
SmartInventory.cmd -List                             prints the catalog (also without argument when the input is redirected)
SmartInventory.cmd <Name> [script options]           runs the entry now on this computer, outside the orchestrator
SmartInventory.cmd <Job>[,<Job>...] -Request [-IncludeDependencies] [-NoWait]
                                                     runs the jobs through the resident orchestrators
SmartInventory.cmd <Name> -WhatIf                    shows the resolved command without running it
SmartInventory.cmd <Name> -ExecutionMode Direct|LocalCache|Auto
```

- Tenant: `-Tenant`, then `SMARTM365_ORCHESTRATOR_TENANT`, then `DefaultTenant` of `SmartM365.global.local.json`. The former launchers always passed `prod`: check `DefaultTenant` on each execution server, or pass `-Tenant`.
- Runtime: `Runtime` of the entry (`PowerShell7` or `WindowsPowerShell`), else `PowerShellEdition`.
- Execution mode of the on-premises entries (those that declare `Execution`): `-ExecutionMode`, then `Execution.Mode` of the job (published manifest, then catalog), then `OnPremExecutionMode` of `SmartM365.global.local.json` (`LocalCache` by default). `LocalCache` copies the script folder, `Config`, `Modules\SmartM365.Core` and the declared `Include` files to the first writable folder among `%ProgramData%\SmartM365\LauncherCache\<CacheName>`, `%SystemRoot%\Temp`, `C:\Temp` and `%TEMP%` (robocopy `/MIR`, failure from exit code 8, `PrepareCache.log`, `Unblock-File`), refuses any output that resolves inside the cache, then runs the cached copy. The cache is hardened: a cache folder or parent that is a junction or link, or that is owned by another account than SYSTEM, Administrators, TrustedInstaller or the launching account, is skipped for the next candidate; the cache folder gets an explicit ACL (SYSTEM, Administrators and the launching account only) and is emptied; right before the launch every code file of the cache (`.ps1`, `.psm1`, `.psd1`, `.dll`, ...) must be byte-identical to its source. The orchestrator uses only `%ProgramData%` and its own `%TEMP%`, never the shared `%SystemRoot%\Temp` or `C:\Temp`. `Direct` runs the script in place on the share. `Auto` runs `Direct` when the share is in the Intranet or Trusted zone, `LocalCache` otherwise. Other entries always run `Direct`.
- `Interactive` entries need a user console; `RequiresConfirmation` entries (`Exchange2016-ProxyAddresses-Write`) ask the user to type `WRITE`.
- The orchestrator GUI (`Operations` tab) offers the same actions: `Jobs` (run now, request, execution mode of the on-premises jobs), `Manual launches`, and the orchestrator actions Start, Stop, Restart, Full collect and Send summary.

Choose a direct run for troubleshooting or a first deployment, and `-Request` for an out-of-schedule production run: it cannot overlap a scheduled run of the same job and never consumes stale dependency data.

### On-premises execution mode: LocalCache, Direct or Auto

The Active Directory and cloud entries always run `Direct`. The Exchange and SharePoint on-premises entries declare `Execution` and run in the mode chosen by `-ExecutionMode`, then `Execution.Mode` of the job, then `OnPremExecutionMode` of `SmartM365.global.local.json`.

- `LocalCache` (default): scripts started from the UNC share are treated as Internet-zone files, so on the Exchange and SharePoint servers unsigned scripts and modules loaded from the share can be blocked or prompt for confirmation (an execution policy set by Group Policy is not overridden by `-ExecutionPolicy Bypass`). The engine therefore copies the runtime locally, unblocks it and runs the copy. Exchange on-premises jobs keep this behavior until one of the conditions below is met.
- `Direct`: runs the script in place on the share, with no copy. Use it once the share is trusted by the servers.
- `Auto`: asks Windows for the zone of the share at each launch: `Direct` for the `Local intranet` or `Trusted sites` zone (or a local folder), `LocalCache` otherwise. A server where the trust is missing therefore falls back to the cache on its own.

Switch to `Auto` (preferred) or `Direct` when one of these is true on every server that runs on-premises jobs:

1. The share is in the Local intranet zone by Group Policy: `Computer Configuration > Administrative Templates > Windows Components > Internet Explorer > Internet Control Panel > Security Page > Site to Zone Assignment List`, entry `file://<server>` (or the server name) with value `1` (Local intranet) or `2` (Trusted sites). Apply the policy (`gpupdate /force`) and open a new session.
2. Or the scripts are signed and trusted: the published scripts and modules carry a valid WorkplaceCloudHub Authenticode signature (repository signing workflow), the public certificate is installed in `LocalMachine` `TrustedPublisher` and `Root` by `Install-WorkplaceCloudHub-CodeSigningCertificate.ps1`, and the effective execution policy (`Get-ExecutionPolicy -List`) is `AllSigned` or `RemoteSigned`. In this case use `Direct`: the zone does not change, so `Auto` would keep the cache.

Procedure, server by server:

1. Check the mode the engine would choose, without running anything:
   `SmartInventory.cmd Exchange2016-ProxyAddresses-Check -ExecutionMode Auto -WhatIf`
   The output must show `Execution mode: Direct (requested Auto, source Parameter, zone Intranet)` (or `zone Trusted`). For signed scripts, check `Get-AuthenticodeSignature` on a collector and on `Modules\SmartM365.Core\SmartM365.Core.psm1` (`Valid`).
2. Run one read-only job in place, outside the orchestrator: `SmartInventory.cmd Exchange2016-ProxyAddresses-Check -ExecutionMode Direct`. It must start without a security prompt and end with its usual result.
3. Pilot one job through the orchestrators: in the GUI, `Operations` > `Jobs`, select the job, set its mode to `Auto` (or `Direct`), `Apply to draft`, then `Publish changes`. The published `Execution.Mode` applies from its next launch, without restarting the orchestrators.
4. Switch all on-premises jobs: either set `"OnPremExecutionMode": "Auto"` (or `"Direct"`) in the shared `Config\SmartM365.global.local.json` (stored as `SmartM365.global.local.json.txt`), read at each launch, or use `Apply to all on-premises jobs` in the GUI and publish. Jobs whose mode is `(default)` follow `OnPremExecutionMode`.
5. Check the job logs: the orchestrator writes `Launch mode: CatalogEngine` and `Execution mode: Direct (requested Auto, source OnPremExecutionMode, zone Intranet)`; no `Cache folder` line follows.

Return to `LocalCache` at any time with the reverse change: `OnPremExecutionMode` back to `LocalCache`, or the job mode back to `(default)` or `LocalCache` in the GUI followed by `Publish changes`. For a single manual run, `SmartInventory.cmd <Job> -ExecutionMode LocalCache`. The next launch rebuilds the cache; nothing else needs to be restored. The former cache folders under `%ProgramData%\SmartM365\LauncherCache` can be deleted once `Direct` is in use.

For one version, the former launcher names stay as one-line aliases, for example `Launchers\Cloud\Start-SmartM365-ActiveUsers-Inventory.cmd` = `SmartInventory.cmd M365-ActiveUsers-Inventory`, `Launchers\Jobs\Start-<Job>.cmd` = `SmartInventory.cmd <Job>` and `LaunchersByOrchestrator\Request-<Job>.cmd` = `SmartInventory.cmd <Job> -Request`. The `LegacyLaunchers` field of each catalog entry gives the mapping. `Orchestrator/New-SmartM365-OrchestratorJobLaunchers.ps1` regenerates the aliases and, with `-Check`, validates the catalog (scripts, parameters by AST, execution, includes, tokens) and the aliases. `Tests/Test-SmartM365InventoryLaunchOffline.ps1` checks every catalog entry against `Tests/SmartInventory-LegacyLaunchers.psd1`, the frozen description of the 186 former launchers.

The Secure Score, authentication-method registration, Conditional Access, Intune endpoint-security, and governed public-France license-pricing collectors are enabled once daily in `Orchestrator/Orchestrator-Jobs.json.template`. Microsoft 365 Backup collectors remain scheduled but disabled until their service-app/controller prerequisite is available. `Policy.Read.All` and `SecurityEvents.Read.All` are new bootstrap permissions. Authentication registration and Microsoft 365 Backup reuse `AuditLog.Read.All` and `BackupRestore-Configuration.Read.All`. Intune endpoint-security reports reuse the temporary export-job permission `DeviceManagementManagedDevices.ReadWrite.All`.

## License pricing

`SmartM365-LicensePricing-Inventory.ps1` 1.1.1 publishes `M365_License_Prices.csv`
(header unchanged) from three local settings:

- `PricingSourceCsvPath`: the public baseline
  `M365Inventory/Licensing/M365_License_Prices_Public-France.csv` (template default).
  Prices are Microsoft France public list prices in EUR excluding VAT, per user per
  month with annual commitment, `IsEstimated=True`. `EffectiveFrom` is the date the
  price was observed and `ContractReference` the public page used.
- `PricingOverrideCsvPath` (private, empty by default): reviewed contract prices with
  the same columns. The public baseline is always read; a private row replaces every
  public period of the same `SkuPartNumber` and `Currency`, and private-only SKUs are
  added. `IsEstimated` and `PriceSource` show which rows are contract prices.
- `EsuPricingCsvPath` (private, empty by default): Windows 10 Extended Security
  Updates prices, published as `Windows_ESU_Prices.csv`, an optional output of the
  collector's source receipt. Without it, ESU stays not valued and the file is not
  published. There is no public ESU row because Microsoft publishes no EUR price.

The private ESU CSV uses these columns:

```text
ProgramKey,CoverageYear,Currency,UnitPricePerDeviceYear,EffectiveFrom,EffectiveTo,IsCumulative,PriceSource,PricingBasis,ContractReference,IsEstimated
```

`ProgramKey` is `WIN10_ESU`, `CoverageYear` is 1 to 3, `Currency` must be `EUR` (no
currency conversion), dates are `yyyy-MM-dd` and required, years must follow each
other without overlapping, and `IsCumulative` states whether a later year also
requires the earlier years.

`UnitPriceMonthly` and `UnitPricePerDeviceYear` are published with a decimal point
whatever the culture of the collection host (`37.78`, never `37,78`). Source prices,
public or private, must use the same format: a decimal comma or a group separator is
refused instead of being misread.

Migration: a local `PricingSourceCsvPath` that points to a file other than the public
baseline keeps the former behaviour (the private file replaces the whole baseline) and
logs a warning. Move those prices to `PricingOverrideCsvPath`, then point
`PricingSourceCsvPath` back to the public baseline (or remove the local value so the
template default applies). Setting both a private `PricingSourceCsvPath` and
`PricingOverrideCsvPath` is refused. The runtime `.local.json` is created from its
template when absent and missing template keys are added without changing local values.

Not valued in the public baseline (no verifiable public EUR price per user):
Exchange Online Archiving for Exchange Server, Teams Phone with calling plans, Intune
Suite, Office 365 Extra File Storage, Copilot Studio (capacity billing), Dynamics 365
attach, Team Members and Operations Activity licenses, Power Apps per app, and Project
Online Premium. Provide contract prices through `PricingOverrideCsvPath` when needed.

## Timestamped CSV copies on SharePoint

`Publish-SmartM365Csv` uploads the latest copy and the timestamped run copy
(`<prefix>_yyyyMMdd_HHmmss.csv`, used by the per-run SharePoint links of the mails).
Since SmartM365.Core 1.0.61 (PS5 compatibility module 1.0.45), once the run copy is
uploaded, the timestamped copies of the same CSV in the same SharePoint folder older
than 7 days are deleted (`-SharePointRetentionDays`, 0 disables), the same retention
as the server. Latest copies and WeeklyHistory never match. Before, SharePoint kept
every run while the server kept seven days. Since Core 1.0.62 (PS5 1.0.46) the same
function also handles timestamped `.xlsx` workbooks (same prefix and extension only).
The Teams (0.34) and SharePoint Online (0.30) collectors, which upload their run CSVs
and workbooks directly, call it after each successful upload. The Active Directory
inventory publishes fixed file names and needs no purge. AD HealthCheck is left
unchanged on purpose: without WeeklyHistory, its SharePoint run copies are its only
history beyond the seven days kept on the server.

`Maintenance/Remove-SmartM365-ExpiredTimestampedCsv.ps1` removes the copies
accumulated before. It is read-only without `-Execute` and removes a copy only when it
is older than `-RetentionDays` (7), its folder has a WeeklyHistory whose manifest for
the copy's ISO week lists `<prefix>.csv` with a non-empty week copy, and `<prefix>.csv`
exists under the DATA-LAST folder next to DATA-ALL. WeeklyHistory, Corrections, Archive,
Orchestrator and run folders are never considered; every kept file is reported with
its reason in the CSV report.

```powershell
# Preview, then delete (synchronized SharePoint copy or UNC DATA-ALL)
.\Maintenance\Remove-SmartM365-ExpiredTimestampedCsv.ps1 -DataAllRootPath "<...>\SMART-M365\DATA\DATA-ALL"
.\Maintenance\Remove-SmartM365-ExpiredTimestampedCsv.ps1 -DataAllRootPath "<...>\SMART-M365\DATA\DATA-ALL" -Execute
```

## Legacy daily statistics import

`Maintenance/Import-SmartM365-LegacyDailyStats.ps1` extends the daily history of
`AD_Users_DailyStats.csv`, `AD_Computers_DailyStats.csv` and
`Exchange_OnPrem_Mailboxes_DailyStats.csv` with the files of older collectors of the same
organization (`AD_Users_Daily_Stats.csv`, `AD_Computers_Daily_Stats.csv`,
`Exchange_Daily_Stats.csv`: same columns, no tenant columns).

- SmartInventory rows win: a legacy row is imported only for a day without any
  SmartInventory row, with the tenant columns of the existing rows.
- Non-numeric legacy values become empty. A single missing part of
  Total = Enabled + Disabled (Exchange also Total = Local + Remote) is computed only when
  that identity holds on every complete row of both files.
- Exchange rows written before collector 1.66 carry 0/1 per mailbox type instead of a
  count (their type columns do not sum to `TotalMailboxCount`). They take the type counts
  of the legacy row of the same day and domain when its totals are identical; otherwise
  the type columns are left empty.
- The import is refused for a legacy domain unknown to SmartInventory or when the totals
  of overlapping days differ by more than `-MaxMedianDifferencePercent` (5).

Run it on the server share, outside the AD and Exchange on-premises collector windows. It
is read-only without `-Execute`; with `-Execute` it backs up every replaced file, writes
DATA-ALL and the DATA-LAST copy atomically, refuses a file changed during the run, and is
idempotent. The next collector runs republish the files to SharePoint.

```powershell
.\Maintenance\Import-SmartM365-LegacyDailyStats.ps1 -DataRootPath '<server share>\DATA\Tenants\prod' -LegacyFolderPath '<legacy csv folder>'
.\Maintenance\Import-SmartM365-LegacyDailyStats.ps1 -DataRootPath '<server share>\DATA\Tenants\prod' -LegacyFolderPath '<legacy csv folder>' -Execute
```

## Server DATA cleanup

Copy the single maintenance script to a machine that can access the server DATA
root; it also works from a temporary folder outside the repository. It does not
launch collectors, access Microsoft Graph, or wait for scheduled jobs. The root
leaf must match `-Tenant`. Preview creates `all-files.csv`, `summary.csv`,
`verified-actions.csv`, `timestamped-csv.csv`, and `orphan-temporaries.csv`
outside DATA. Run the same
command with `-Execute` to remove
verified items. The script rechecks source hashes, folder state, and prepared
pointer before deletion; it never sweeps arbitrary prepared batches. It reads the
effective `Orchestrator/Orchestrator-Jobs.json` when running in the repository to
qualify removal of the retired CMDB lock. When running from a temporary folder,
pass `-EffectiveJobsPath` pointing to the deployed runtime jobs file; otherwise
the lock stays in the review report.

Atomic writes stage a short sibling file (`~sm<12 hex>.pending`, `.previous`,
`.tmp`, `.partial`, `.download`, or the target extension such as `.csv`). The
JSON transport keeps its staging and previous copies when a write fails, and no
writer resumes them, so the script retires them as orphan temporaries. A file is
eligible only when its name exactly matches a known writer pattern, including the
legacy `<target>.<32 hex>.pending|previous|tmp|partial|download`,
`<target>.tenantkey.<32 hex>.tmp`, `.smartm365-write-test-*.tmp` and
`.smartm365-preflight-*.tmp` names; when its last creation or write time is older
than `-OrphanTemporaryMinimumAgeDays` (default 7), or
`-OrphanPreviousMinimumAgeDays` (default 30) for a `.previous` copy, which can be
the only prior version of a file; and when no `*.transport.lock` in its folder is
held. `-Execute` requalifies each file and deletes it while holding every
transport lock of the folder, logs each removal with a timestamp, and records the
result in `orphan-temporaries.csv`. Weekly `upload.pending` markers, the
transport lock files, and the ambiguous legacy `<name>.<32 hex>.csv` form are
never removed by this rule.

The server share keeps the long history of prepared batches; SharePoint keeps
only the current and previous batch through its own publication retention. By
default, linked batches older than the current and previous batch of the
`DATA-POWERBI\current.json.txt` chain are never planned for removal: they appear
in `all-files.csv` as `Keep` ("Linked prepared batch kept: share retains
prepared-batch history"), and the log reports how many were found and kept.
`-RetireLinkedPreparedBatches` enables their retirement. The current and
previous batch always stay; `-RetainedLinkedPreparedBatchCount` (default 0)
keeps that many further batches along the chain, and
`-LinkedPreparedBatchMinimumAgeDays` (default 0) keeps any batch whose id
timestamp is more recent. A retired batch keeps its `current`, `batch` and
`validation` receipts under `DATA-POWERBI\retired\<batch id>`, SHA256-checked
before the folder is removed. Unlinked batches remain review items.

```powershell
.\Maintenance\Invoke-SmartM365-ServerDataCleanup.ps1 -Tenant prod -DataRootPath '<server share>\DATA\Tenants\prod'
.\Maintenance\Invoke-SmartM365-ServerDataCleanup.ps1 -Tenant prod -DataRootPath '<server share>\DATA\Tenants\prod' -Execute
# Optional: retire linked batches beyond 10 older ones and older than 90 days.
.\Maintenance\Invoke-SmartM365-ServerDataCleanup.ps1 -Tenant prod -DataRootPath '<server share>\DATA\Tenants\prod' -RetireLinkedPreparedBatches -RetainedLinkedPreparedBatchCount 10 -LinkedPreparedBatchMinimumAgeDays 90
```

For the next server cleanup pass, `Maintenance/SmartM365-ServerDataSecondPass.ps1`
is read-only. It writes separate CSV reports for SHA256-qualified M365 Usage run
duplicates, WorkplaceScope snapshots versus DATA-LAST, generated Entra removal
scripts, old DATA-LAST outputs and coordination locks, repair backup integrity,
prepared `failed`/`retired` receipts, and the CMDB lock. It preserves
`Intune_Devices_Win11Readiness.csv` and does not delete any file. Review the
reports before extending the cleanup rules. Pass the effective deployed jobs
file when it is not next to the script.

```powershell
.\Maintenance\SmartM365-ServerDataSecondPass.ps1 -Tenant prod -DataRootPath '<server share>\DATA\Tenants\prod'
```

After reviewing a second-pass report, use
`Maintenance/SmartM365-ServerDataSecondPassCleanup.ps1` for a separately
approved, file-only cleanup. Supply the original report directory, a fixed UTC
cutoff, and the approved counts for Usage duplicates, old generated Entra
removal scripts, DATA-LAST test outputs, and the transport log. Preview writes
`plan.csv` outside DATA; Execute requires that exact plan and recomputes all
candidate and retained-copy SHA256 hashes. It refuses changed or missing files
and never removes directories, distinct Usage files, WorkplaceScope history,
repair backups, prepared diagnostics, or the CMDB lock.

```powershell
$cleanup = @{
    Tenant = 'prod'
    DataRootPath = '<server DATA tenant root>'
    EvidenceDirectory = '<reviewed second-pass report directory>'
    CutoffUtc = [datetimeoffset]'<approved UTC cutoff>'
    ExpectedUsageFiles = <approved count>
    ExpectedEntraScripts = <approved count>
    ExpectedTestOutputs = <approved count>
    ExpectedTransportLogs = <approved count>
}
$preview = .\Maintenance\SmartM365-ServerDataSecondPassCleanup.ps1 @cleanup
.\Maintenance\SmartM365-ServerDataSecondPassCleanup.ps1 @cleanup -Execute -PlanPath $preview.PlanPath
```

For a separately approved residual cleanup, use
`Maintenance/SmartM365-ServerDataResidualCleanup.ps1`. It checks the effective
orchestrator jobs file under the tenant DATA root, an exclusively available
retired CMDB lock, the repair audit and every rollback-copy hash, and the
reviewed DATA-LAST file names, sizes, and timestamps. It also verifies that
each repaired output remains available, including an identical Discovered Apps
SummaryHistory copy when the old weekly detail has expired. Preview records a
hash-pinned plan outside DATA. Execute requires that plan, repeats the checks,
and removes only the approved DATA-LAST files, lock, and one repair backup.
The report retains a copy of the repair audit, a retained-output map, and
the backup directory list and execution results. Extra empty directories
inside that backup are permitted; linked directories or unlisted files stop
the run. Removing the repair backup discards its pre-repair rollback copies;
keep the external report.

```powershell
$cleanup = @{
    Tenant = 'prod'
    DataRootPath = '<server DATA tenant root>'
    EvidenceDirectory = '<reviewed second-pass report directory>'
    BackupFolderName = '<approved repair backup folder name>'
    ExpectedAuditTenantKey = '<audited tenant key>'
    ExpectedRepairFiles = <audited repair receipt count>
    ExpectedSummaryFallbacks = <approved fallback count>
    ApprovedDataLastNames = @('<reviewed obsolete output.csv>')
}
$preview = $null
$preview = .\Maintenance\SmartM365-ServerDataResidualCleanup.ps1 @cleanup
if ($null -eq $preview -or $preview.Mode -ne 'Preview') { throw 'Preview did not complete.' }
.\Maintenance\SmartM365-ServerDataResidualCleanup.ps1 @cleanup -Execute -PlanPath $preview.PlanPath
```

## Active Directory launches

Use the catalog entries with `SmartInventory.cmd`:

- `AD-Inventory`: runs the complete Active Directory inventory for the selected tenant, including live domain collection, consolidated CSV generation, enrichment, duplicate and mail-routing analysis, reports, publishing, and weekly history.
- `AD-DuplicateIdentity-Mail`: skips live Active Directory collection, reads the existing `DATA-LAST/AD_Users_AllDomains.csv`, regenerates the four duplicate and mail-routing diagnostic CSV files and the `AD_Users_IdentityAndMailRoutingIssues.xlsx` workbook, then forces the diagnostic email to be sent even if it was already sent that day.

The tenant comes from `-Tenant`, then `SMARTM365_ORCHESTRATOR_TENANT`, then `DefaultTenant`. The mail-only entry is intended for controlled reanalysis or notification retesting and does not replace a regular full inventory.

## Account classification governance

The account-type rules and workforce population mappings are tenant configuration. The real rules live in the private, git-ignored `Config/AccountClassification.local.json` (stored as `AccountClassification.local.json.txt`); only the neutral `Config/AccountClassification.local.json.template` is published. Create the private file from the template and adapt it to the tenant naming conventions before the first inventory; a missing file stops the AD enrichment and the prepared workforce evidence instead of falling back to generic rules. All readers use `Common/SmartM365.AccountClassification.psd1`. The Active Directory enrichment writes `AccountType`, `AccountPopulation`, and `AccountClassificationRuleVersion` to its enriched user export.

- `Human`: named and external-person accounts included in workforce KPIs.
- `Non-human`: service, shared-mailbox, room-mailbox, and system accounts excluded from workforce KPIs.
- `Review required`: admin, generic, unclassified, conflicting, or cloud-only accounts that must not be silently treated as human or non-human.

Site and persona columns use the same private workbooks as SmartWorkplaceIntelligence: `SmartWorkplaceIntelligence-SiteClassification.xlsx` and `SmartWorkplaceIntelligence-PersonaClassification.xlsx`, read from the SharePoint DATA root on every Active Directory inventory run. They replace `EntityDirectories.xlsx` and `Personae-Keywords.xlsx`. All readers use `Common/SmartM365.WorkplaceClassification.psd1`, which applies the prepared-evidence rules: a non-human exclusion first, then the HQ site persona, then job-title keywords filtered by country, and the AD description when the title has no unique match. Equal-priority matches across personas stay `Review required`. A failed download, a missing worksheet or an invalid rule stops the inventory instead of publishing blank or wrong site and persona columns.

- `EntityLabel`, `EntitySiteType` and `EntityIntegrationType` come from the site workbook (`Site Display Name`, `Site Type`, `Integration Status`), keyed by the AD attribute named in its Settings. An unmapped code gives `NOTFOUND`, and the integration type then falls back to the domain rule.
- `JobFamily` holds the persona or review state. For a `Review required` multiple-match persona, `M365LicenseTargetPersona` keeps the current base license (E3, otherwise F3), so an uncertain classification never proposes a license change. The `Relaxed` columns now carry the same values as the main ones.

Organization-specific directory rules are also private, in the `DirectoryRules` section of the same file. `IntegrationFallback` holds the ordered domain rules that give `EntityIntegrationType` for users and computers whose site is unmapped. Each rule has `Domains`, an optional `CountryIn` or `CountryNotIn`, and a `Result`. For computers, the country is read before `ComputerCountryOuSuffix` in the distinguished name. `HeadquartersOuContains` lists the OU fragments that mark headquarters accounts and computers. `Windows11OutOfScopeDomains` lists the domains excluded from the Windows 11 migration phases. A missing section stops the enrichment. The Exchange hybrid identity issues collector reads the routing domain from the tenant `RemoteRoutingDomain` and the allowed TargetAddress domains from `AllowedTargetDomains` in its local configuration. It refuses the template values.

Change `RuleVersion` whenever a classification or population rule changes. Keep the three populations disjoint; validate the template and the private rules with `Tests/Test-SmartM365-AccountClassification.ps1` before running a full inventory.
