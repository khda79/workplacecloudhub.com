# Smart SharePoint Migration Toolkit

Smart SharePoint Migration Toolkit helps validate SharePoint migrations between
a source and a destination. It inventories source and destination content,
compares files and permissions, exports review workbooks, and generates guarded
cleanup scripts for destination-side differences.

File comparisons include the current SharePoint document version exported from
both inventories. They also flag matched files where the destination modified
date is older than the source, so a copied file that is not at the latest source
version is visible in the comparison workbook.

The toolkit validates and reconciles evidence; it does not copy or migrate content.
The dashboard component is **1.0.8** and the generic launcher is **1.0.19**.
See [release notes](RELEASE-NOTES-1.0.8.md) for the changes and validation boundary.

## Install, Start, and Update

Version **1.0.8** is distributed through the
[GitHub release](https://github.com/khda79/workplacecloudhub.com/releases/tag/sharepoint-migration-toolkit-v1.0.8)
and the [existing source folder](https://github.com/khda79/workplacecloudhub.com/tree/main/SmartM365/SharePointMigration).
Download the release ZIP together with its `.sha256` and `.manifest.json` assets.
Verify the ZIP with `Get-FileHash -Algorithm SHA256` before extracting it, then open
`SmartM365/SharePointMigration` inside the extracted package. The root also includes
the public signing certificate, its optional trust installer, licence and notice.
No PowerShell Gallery package is published for this toolkit; do not use `Install-Module`.

Obtain the repository and keep the entire `SmartM365/SharePointMigration` folder;
copying just the GUI script omits required helpers, assets and templates. Launch
`Start-SmartM365-SharePointMigration-GUI.cmd` on Windows. The dashboard provides
Overview, Files & Permissions, Operations, Migration Diagnostics, Logs and Config tabs; it discovers configured
migration folders and displays the most recent output paths. An output timestamp
or an available Open button is not proof that the whole scan succeeded: inspect
the run log and any error CSV before accepting its results.
The dashboard requires PowerShell 7.4 or later. The CMD launcher checks this
requirement before opening the GUI, and the GUI enforces it for direct launches.
Windows PowerShell 5.1 remains in use as a child process for SharePoint Server
inventory and SPO site administration scripts that require it.
The farm diagnostic has a separate Windows PowerShell 5.1 launcher and does not
open or require the dashboard; see "Source farm diagnostics" below.

The header displays the mapped scan scope. When a migration maps several webs,
hover over the scope count to see their URLs. A target `SiteUrl` that differs
from every mapped target is flagged; the GUI blocks site operations until the
configuration is aligned. The header has an enabled-by-default 30-second
refresh checkbox. Refresh keeps unsaved configuration edits and the selected
log; turn the checkbox off to stop automatic updates.
Overview shows every configured migration, including those without a completed
comparison. It displays the mapped source and destination, with separate file
and permission scan dates in each source and target scan cell. The displayed gap
in decimal days compares the file scans. Rates and comparison dates are separate
for files and permissions. The file rate is matched
files divided by source unique keys;
the permission rate is matched permissions divided by source unique permission
keys. An empty inventory or missing comparison has no rate.
The Status column combines file and permission scan availability, comparison freshness,
findings, and scan provenance. Hover over a row for the separate file and permission
reasons. Click a row to select that migration and open Files & Permissions. The
table reads only inventory names, small scan receipts, and comparison summaries;
it does not parse the full inventories on each refresh. A scan receipt provides
its completion time, but the GUI does not recalculate its CSV hash on each
refresh. Dates from legacy CSV filenames are identified in the tooltips.
If one migration cannot be refreshed, its last successful values remain visible;
the row tooltip and shared activity log show the refresh error.
Files and Permissions share one tab, with a vertical divider between their workflows.
The file and permission scan selectors follow the newest completed CSV until an
older scan is selected manually. Each source-versus-target comparison refreshes
its selectors immediately before launch and asks for confirmation when a selected
scan is older than an available one. The launch log records the exact source and
target CSV paths. Both columns offer scan history comparison for Source or Target;
select a previous and current scan of the same kind and endpoint.

The Logs tab shows shared GUI activity across migrations as well as the
selected migration's script logs. Each GUI session, migration creation,
configuration save, and GUI-launched action has a separate activity file under
`Migrations/logs/gui-activity`. Records include the Windows user who used the
GUI, machine, action, local and UTC times, status, exit code, and the detailed
run log when available. The script's execution account may differ from the
GUI user. These records require write access to the shared migration folder;
folder permissions govern their integrity. Actions launched outside the GUI
are represented by their script logs rather than by GUI activity records.

The startup GitHub check reports component versions only. It does not install an
update. Set `SPMIG_GUI_UPDATE_CHECK=0` to disable that check. To update, close the
GUI and running jobs, back up local migrations and authentication files, then
replace application source files from a reviewed revision. Preserve local
`Migrations/<name>`, `Config/SPOAuth.local.psd1` and `Tools/Python`. Review the
template updater before applying it to existing migration folders; keep local
configuration and mapping values under operator control.

For a local resource check without displaying the GUI or contacting a tenant:

```powershell
.\SmartM365-SharePointMigration-GUI.ps1 -ValidateOnly
```

## Layout

- `Scripts/Inventory/`: source and destination file and permission inventories.
- `Scripts/Compare/`: source versus destination comparisons and file/permission
  scan history comparisons.
- `Scripts/Export/`: CSV to Excel export helpers used by comparison workflows.
- `Scripts/Generate/`: generated operation script builders for reviewed
  destination cleanup.
- `Scripts/Operations/`: guarded destination cleanup and SharePoint
  administration operations.
- `Scripts/Launchers/`: migration-aware launchers shared by each migration
  folder.
- `Migrations_Template/`: safe template used to create local migration folders.
- `Update-MigrationsFromTemplate.cmd`: refreshes launchers and folders in local migrations from the template.
- `_Local/`: private workstation helpers, excluded from Git.
- `Config/SPOAuth.sample.psd1`: placeholder model for local SharePoint Online
  authentication values.

## Local Migrations

In the dashboard, click `+ New` to enter a migration name, source and target
types, the first source/target web URL pair, any additional mapping pairs, and
the SPO target admin URL when applicable. The wizard checks every web through
an authenticated SharePoint read and verifies the SPO admin endpoint before
creating a directory. SharePoint Server checks use the current Windows account
and the `/_api/web` endpoint; SPO checks use the selected PnP authentication
mode and `Get-PnPWeb`. A missing site, denied access, unavailable module or
inconclusive response blocks creation. Checking the SPO admin endpoint also
requires tenant administration access through the selected PnP authentication.
Run the GUI where the URLs and the required authentication are available. The
wizard copies `Migrations_Template`, writes
the configuration and mapping, and selects the new migration only after all
checks succeed. For a multi-web migration, each additional mapping must have
one source URL and one target URL on its own line.

You can also create one local folder per migration manually by copying the template:

```powershell
Copy-Item -Recurse .\Migrations_Template .\Migrations\MyMigration
```

Run `Update-MigrationsFromTemplate.cmd` from the toolkit root to refresh
existing migration launchers and template-owned folders.

Then edit the copied files:

```text
Migrations\MyMigration\migration.config.psd1
Migrations\MyMigration\migration.mapping.txt
```

`Source.Type` and `Target.Type` define which inventory engine is used for each side. Supported values are `SP2016`, `SP2019`, and `SPO`; for example, `SPO` to `SPO` migrations are supported by setting both sides to `SPO`. When `Source.UrlsFile` and `Target.UrlsFile` are omitted, the launcher derives both scan URL files from `migration.mapping.txt`.

Set `Name` in the copied config to a report label other than `NewMigration`.
The GUI selects the migration **directory name**, which may differ from this
label. The CLI `-MigrationName` also takes the directory name.

Each active mapping line must contain exactly two fields: source and target.
Spaces, tabs, semicolons or commas separate the fields. Encode spaces inside
URLs/paths as `%20` (and separator characters inside names as URL escapes).
Comments must be on their own lines beginning with `#`. Empty mapping files,
missing fields and extra fields are rejected rather than silently narrowing the
scope. Review all mappings before scans or comparisons; URL overrides must use
the same intended scope.

```text
https://source.workplacecloudhub.com/SOURCE https://workplacecloudhub.sharepoint.com/sites/Example
```

Run both file inventories, inspect their logs/errors, then CompareFiles. Review
HTML/CSV results and Excel workbooks, including versions, sizes, modified dates,
missing files and duplicate keys. Repeat with permission inventories and
ComparePermissions. Permission comparison requires a fresh Entra users cache
and at least one SPO endpoint; the current launcher rejects an on-premises-only
permission comparison. This restriction does not prevent file comparisons.
The cache contains identity fields for all Entra users returned by Microsoft
Graph. By default, migrations targeting the same SPO host and configured tenant
share a cache under `_Local/EntraUsersCache`. Its CSV and `.meta.json.txt`
receipt are reused for up to 12 hours (or a shorter configured age). The receipt
records the Graph tenant, export time, user count, and CSV SHA256; mismatches
trigger a refresh. A per-cache lock serializes refreshes from multiple GUI users.
An explicit `Comparison.EntraUsersCachePath` still overrides the shared default.
If a required Graph submodule is absent, the script asks before installing only
that submodule for the current PowerShell 7 user. Older migration configs still
set to 24 hours are capped at 12 hours at runtime.
Permission scan history comparison reads two CSV inventories from the same
endpoint and requires neither SharePoint authentication nor an Entra cache.
It writes `PermissionChanges.csv`, `Summary.csv`, and a branded HTML report
under `comparisons/permission-scan-history` by default. It counts added,
removed, changed, unchanged, and ambiguous permission grants.
The branded HTML summaries are self-contained and show at most 20 objects with
differences; use their linked CSV and Excel exports for the full detail.
The comparators load the shared `report_html.py` from their own directory. They
add that directory explicitly because the bundled Portable Python runs in
isolated mode and does not add the script directory to its import path.
Executable PowerShell and Python console scripts print their name and version,
the WorkplaceCloudHub introduction, and a timestamped completion summary with
status, duration, and the run log path when one exists. A nested script does not
repeat the introduction or summary; CMD launchers delegate to their scripts.
A successful script
run means the comparison completed; inspect its report for migration differences.
Runtime outputs stay inside the local migration folder:

```text
scans\
comparisons\
operations\generated\
logs\
```

`Migrations/*` is ignored by Git except for each migration's `ShareGate`
folder structure. Put ShareGate
migration reports (`.xls`, `.xlsx`, or `.csv`) in
`Migrations/<name>/ShareGate/MigrationReport`. Only the generic README and
empty-folder marker are versioned there; the reports remain local. Do not commit
real migration configuration, inventory CSVs, workbooks, logs, generated cleanup
scripts, or local authentication files.

## Authentication

SharePoint Online launchers can read local values from:

```text
Config\SPOAuth.local.psd1
```

Start from `Config/SPOAuth.sample.psd1` and keep the real local file uncommitted.
The sample contains placeholders only.

## Requirements

- SharePoint Server Management Shell when a SharePoint Server farm is used as a
  source or destination.
- PowerShell 7.4 or later for current PnP.PowerShell; check the requirements of
  the exact installed module version. See the [PnP installation guide](https://pnp.github.io/powershell/articles/installation.html).
- Windows PowerShell 5.1 on a SharePoint server for SP2016/SP2019 snap-in scans.
  When a scan is launched from PowerShell 7, the generic launcher automatically
  re-enters Windows PowerShell for that on-premises endpoint and propagates failure.
- PnP.PowerShell for SharePoint Online inventory and permission scans.
- Windows PowerShell 5.1 and Microsoft.Online.SharePoint.PowerShell for SPO
  admin operations such as site lock state and page comment settings.
- Python 3 for comparison, Excel export, and generated-operation helpers.
- Microsoft.Graph.Users and Microsoft.Graph.Authentication for Entra cache
  refresh. The script asks before installing missing submodules for the account
  and PowerShell 7 host that runs the GUI; the shared migration folder does not
  provide them.
  Comparison/export Python helpers use the standard library; Excel
  itself is not needed to generate workbooks.

SP2016/SP2019/SPO are configured engine selectors, not a certification of every
farm, module, operating system or migration combination. Use an environment
pilot before relying on these results for acceptance.

## Permissions and Evidence Coverage

| Action | Required access and operational boundary |
| --- | --- |
| Local GUI, comparisons, exports | Read input files and write migration output folders. No SharePoint write permission is needed for local file comparisons. |
| SharePoint Server scans | SharePoint Management Shell/snap-in on the farm; an approved account with shell/database access and visibility of the selected sites and permissions. Do not assume a workstation can scan a farm remotely. |
| SPO file inventory | Your own Entra application/client ID for PnP authentication, with consent and read access to every selected web/library. |
| SPO permission inventory | Access to enumerate role assignments, groups and selected item permissions. Ordinary document read access alone is not sufficient evidence of complete permission coverage; validate access on representative sites. |
| Entra cache refresh | The interactive code requests Graph `User.Read.All` and `Directory.Read.All`. Certificate mode needs consented application access for the exported user fields and access to the certificate private key. |
| Cleanup and site administration | Separate, deliberately granted write/admin access for the selected target. Generated cleanup is SPO-oriented; it is not an on-premises migration engine. |

See [PnP authentication](https://pnp.github.io/powershell/articles/authentication.html)
for application registration and supported authentication modes. The toolkit
does not grant permissions or prove their adequacy. Keep app IDs, certificate
references, URLs and user exports local; logs and workbooks are not guaranteed
anonymous and must be reviewed before sharing.

Scope and paging settings remain explicit: descendant subsites are included by
the inventory roots; hidden/system content is excluded by default unless opted
in; permission list/library and item settings affect what is assessed. PageSize
and RowLimit are paging controls, not a promise that a partial or failed scan is
complete. Current file version labels and version counts do not validate every
historical version or byte-for-byte content integrity. Review scan errors,
filtered/excluded rows, disabled-user separation and Limited Access handling.

The launcher first checks for `Tools\Python\python.exe` and then falls back to
`python` or `py -3` from the local workstation.

To create or refresh the local portable runtime:

```powershell
.\Scripts\SmartM365-SharePointMigration-InstallPortablePython.ps1
```

Use `-Force` to replace an existing local runtime. The script downloads the
official Windows embeddable Python package, verifies the expected SHA256 hash,
and extracts it to:

```text
Tools\Python
```

`Tools\Python` is ignored by Git because it is generated local runtime content.
For double-click usage on Windows, run
`Tools\SmartM365-SharePointMigration-InstallPortablePython.cmd`.

If the workstation cannot reach `python.org`, download the embeddable package
from another machine, copy it locally, then run:

```powershell
.\Scripts\SmartM365-SharePointMigration-InstallPortablePython.ps1 -PackagePath .\Tools\python-3.13.13-embed-amd64.zip -Force
```

The `.cmd` launcher also forwards arguments, so the same offline package can be
used from a command prompt:

```cmd
Tools\SmartM365-SharePointMigration-InstallPortablePython.cmd -PackagePath .\python-3.13.13-embed-amd64.zip -Force
```

## Safety Model

Cleanup scripts are review-first. They are generated from comparison outputs,
default to dry-run behavior, and require explicit execution flags before making
changes. Keep the operational order: remove extra files first, then extra empty
folders, and review extra libraries separately.

For final copy validation, run fresh source and destination file scans close
together before `CompareFiles`. The launcher enforces
`Comparison.MaxScanAgeDifferenceHours` and uses
`Comparison.MaxScanAgeHours` (24 hours by default) to reject old scans even
when source and target are close to one another. New launcher scans publish a
`*.csv.manifest.json.txt` receipt with completion time, scope, row count, and
SHA-256. The comparison verifies that hash. For older scans without a receipt,
the timestamp embedded in the CSV filename is used so copying a CSV cannot
make it appear new; a nonstandard filename falls back to file time with a
warning. Zero-row scans with a receipt block comparison. Reports from older
zero-row scans display an inconclusive status.
File and permission HTML reports also show scan evidence and flag stale or
unverified scans for review when an override allowed comparison to continue.

The file report displays `SuccessPercent = matched files / source unique keys`
within the compared scope. Additional target files and filtered rows are
displayed separately; 100% does not certify migration completeness. In the GUI,
**Global file comparison report** in Files builds offline HTML, Excel and CSV
tables from each migration's latest file `Summary.csv` under
`Migrations/reports/global`. **Global permissions comparison report** in
Permissions does the same for the latest top-level permission `Summary.csv`
under `Migrations/reports/global/permissions`. Both HTML reports show a prominent
Excel link above the metrics.

Report branding is configured in `Config/report-branding.json.txt`, based on
`Config/report-branding.json.template`. `WorkplaceCloudHubLogoPath` names a
PNG or JPEG in the SharePointMigration root. `ClientLogoPath` names an optional
PNG or JPEG in `Config`; leave it empty to show only the WorkplaceCloudHub logo.
Newly generated HTML reports embed the images, so viewers do not need access
to the original files. The local `.json.txt` configuration and client image
are ignored by Git and must be copied privately when deploying another clone.

## Migration Diagnostics (local report analysis)

The **Migration Diagnostics** tab reads ShareGate migration report exports for
the selected migration from `ShareGate/MigrationReport`. You can also browse to
another CSV, XLSX, or report folder. This phase only reads report files; it
does not import the ShareGate module, connect to a site, precheck, or retry a
migration. The analysis runs in a child PowerShell process so the GUI remains
responsive. Shared GUI activity logs record the operator and the analysis
result. CSV is preferred when a CSV and XLSX have the same base name. DryRun
lists every file, its selection status, detected session IDs, and the installed
`ImportExcel` version. With ImportExcel available, a same-name XLSX is masked
only after its row count and row identity columns match the CSV. The remaining
cells may differ between export formats; the CSV remains the selected evidence.
XLSX-only analysis needs the optional `ImportExcel` module. Report rows from different
files are deduplicated by session and row ID; conflicting duplicates are
counted and flagged for review. A session selector can restrict a new analysis
to one session.

The tab and HTML report show both report-line KPIs and distinct keyed content
items. An item key uses source site, source list, and positive `Source ID`, so
version rows for the same item are grouped. Rows without that key are reported
separately and are not invented as distinct items. Raw success rates describe
the ShareGate export; they do not prove source-to-target completeness.
Residual rates count `To fix` issues and exclude `Accepted` issues from their
denominator. `Fixed` records an operator decision, not a newly verified
migration result. Unavailable site features default to `Accepted`; operators
can change any pattern to `To fix`, `Accepted`, or `Fixed`. Each pattern state
is saved as a separate private `.json.txt` file under
`ShareGate/DiagnosticsState`, with operator and machine identity, and is reused
for later analyses. When a report does not identify whether an access denial
occurred at the source or destination, the category remains undetermined.

Each run writes private evidence under `ShareGate/Diagnostics/<timestamp-id>`:
`MigrationDiagnostics-Report.html`, `Summary.json.txt`,
`ClassifiedRows.csv` (including original columns), `PatternSummary.csv`,
`UnknownPatterns.csv`, `Remediation-Actions.csv`,
`UserMapping-Candidates.csv`, and `DestinationMatches-Review.csv`.
The HTML report breaks manual actions down by category. Undetermined access
denials are excluded from manual actions.
`Help links` are clickable from a selected pattern in the GUI and in HTML.
The candidate mapping and ambiguous-match files are review inputs, not
ShareGate mapping files or authorization to retry. The latter concerns
destination object identity; it must not be treated as a user mapping.
`SourceIdentity` in the user candidates comes from the report's user item
title and may be a display name rather than a UPN; verify it before mapping.
`Copy options`, Microsoft 365 import status, and throttling statistics remain
in the classified rows for review. Generic EN/FR column aliases and rules are
in `Config/sharegate-diagnostics.columns.json.template` and
`Config/sharegate-diagnostics.rules.json.template`; private `.json.txt`
runtime copies are created at normal GUI startup if absent. The analyzer also
creates them when run directly and fills in missing defaults during analysis.
The rules can be edited locally without changing the GUI.

Preview which local files would be analyzed:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Scripts\Diagnostics\SmartM365-SharePointMigration-Diagnostics.ps1 -ProjectRoot .\Migrations\MyMigration -DryRun
```

Generate an analysis without using ShareGate:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Scripts\Diagnostics\SmartM365-SharePointMigration-Diagnostics.ps1 -ProjectRoot .\Migrations\MyMigration
```

### Source farm diagnostics

On a farm server, open an elevated console at the toolkit's shared UNC folder.
The regular CMD launcher performs DryRun; the `-Run.cmd` launcher performs real
read-only collection. Both start only Windows PowerShell 5.1. Pass the migration
folder name with `-Project` for unattended use, or omit it to see a numbered
project list. The list shows the latest CSV date or "CSV missing"; choose a
number or `0` to cancel. No project is selected by default, and a project
without a usable CSV cannot run. The launcher selects the newest
`AccessFailures-5min.csv` under that project's `ShareGate/Diagnostics` and
includes all valid UTC windows. If the launcher is stored outside the shared
toolkit, supply
`-ToolkitRoot` with the shared UNC root. It does not call the PowerShell 7 GUI.

```text
Start-SmartM365-SharePointMigration-FarmDiagnostic.cmd -Project MyMigration
Start-SmartM365-SharePointMigration-FarmDiagnostic-Run.cmd -Project MyMigration
```

`Scripts/Diagnostics/SmartM365-SharePointMigration-FarmDiagnostic.ps1` runs from an
elevated Windows PowerShell 5.1 console on a SharePoint farm server. It loads
the SharePoint snap-in or Subscription Edition module itself and checks farm
access. It only reads farm configuration and logs. It does not alter IIS,
SharePoint, audit policy, or the service state. Its only persistent writes are
the output files under `ShareGate/Diagnostics/Farm-<timestamp>` or the local
fallback folder.

The default output is the project's private diagnostics folder on the shared
toolkit path. Launch the script from that UNC path or supply `-ToolkitRoot` with
the shared UNC toolkit root. `-OutputPath` overrides its parent folder. If that path cannot
be written, results go to `C:\Temp\SPFarmDiag\<timestamp>` and the script prints
a `-CopyResultsFrom` command to copy the completed output to the share.
`-Project` takes the migration directory name. The script accepts
`-StartTime/-EndTime`, `-Around/-WindowMinutes`, or an
`AccessFailures-5min.csv` file through `-ShareGatePeaksCsv`. The time zone is
detected on the farm server unless `-TimeZone` provides a Windows time zone ID.

Remote Windows events use RPC through `Get-WinEvent -ComputerName`. IIS
configuration and W3C logs are read through administrative shares; remote ULS
uses the configured log directory through `Get-SPLogEvent -Directory`. WinRM is
optional only for checking the remote audit policy. IIS files are processed
line by line and filtered by UTC window while reading. The report records the
read duration, coverage gaps and a local rerun command for each inaccessible
server. The IIS robot user agent `MS Search 6.0 Robot` is counted as excluded
and does not contribute to access correlations.

Each run writes separate topology, IIS, WAS, Security, ULS, coverage,
correlation and nightly-recurrence CSVs, an HTML report, a log and an atomic
`Farm-Summary.json.txt` completion marker. A `Partial` result can contain
useful evidence but does not establish complete farm coverage. The Migration
Diagnostics tab shows the latest matching farm result and generates one-line
DryRun and real commands from the project's `AccessFailures-5min.csv` windows.
The GUI does not execute those commands on the farm. Run the DryRun command
first on a farm server, then review and run the real read-only command. Launch
the GUI from the shared UNC toolkit root or pass `-FarmToolkitRoot` with that
root to generate UNC commands.

The offline fixture test is:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\Test-SmartM365FarmDiagnostic.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\Test-SmartM365FarmDiagnosticLauncher.ps1
pwsh.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\Test-SmartM365FarmDiagnosticGui.ps1
```

## ShareGate probe (phase 2b)

`Scripts/Diagnostics/SmartM365-SharePointMigration-ShareGateProbe.ps1` runs in
Windows PowerShell 5.1. Its default DryRun reads the latest phase 2a
`ClassifiedRows.csv`, enumerates distinct access-case site/list endpoints,
detects the installed ShareGate application and module versions, and writes a
plan under `ShareGate/Diagnostics/Probe-<timestamp-id>`. It does not import
ShareGate or connect to a tenant. Use `-AnalysisDirectory` to choose a specific
phase 2a run and `-SessionId` to narrow the cases.
Run phase 2a against the same project folder first: generated `Diagnostics`
outputs are ignored by Git and do not arrive with a repository pull.

Explicit `-Run` imports the installed ShareGate module, records installed
cmdlet count and parameter sets without invoking copy commands, and calls
`Find-CopySessions` as a read-only license check. An empty local session
history is reported as such; session properties are recorded only when the
requested object exists locally. If the license check fails, site reads are
skipped. Source site checks are skipped by default; `-ProbeSource` enables
them when run on a host that can reach the source. With
`-SourceAuthMode Default`, `Connect-Site -Url` uses the current Windows user
for the on-premises source, without a supplied credential. On a GUI machine
with access to both environments, one `-Run -ProbeSource` checks both sides.
Destination checks
use only `Connect-Site -Browser` and `Get-List`, with no supplied username,
password, or saved-connection option. `Probe-Access.csv` attributes a denial
to Source or Destination only when that side returns an access-denied error
and the opposite endpoint is readable. Other cases remain Undetermined because
site/list reads cannot prove item-level access. The probe writes
`Probe-Results.csv`, `Probe-CmdletParameters.csv`, `Probe-SessionProperties.csv`,
and `Probe-Access.csv` atomically, plus `Probe.log`, all in the private
migration folder. A successful `Find-CopySessions` call supports that
PowerShell integration is licensed for Pro or Enterprise, but the probe cannot
identify the exact subscription tier. No migration or copy command is invoked.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\Diagnostics\SmartM365-SharePointMigration-ShareGateProbe.ps1 -ProjectRoot .\Migrations\MyMigration -SessionId 260930-6 -ProbeSource -SourceAuthMode Default -DryRun
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\Diagnostics\SmartM365-SharePointMigration-ShareGateProbe.ps1 -ProjectRoot .\Migrations\MyMigration -SessionId 260930-6 -ProbeSource -SourceAuthMode Default -Run
```

### Access evidence in phase 2a

Phase 2a classifies access failures from an explicit failing `WebUri` or
`URL ... was not authorized` in the original ShareGate exception trace. It
compares that URL's host with the row's source and destination hosts, and
keeps the trace excerpt in `ClassifiedRows.csv` and the HTML report. Missing,
mixed or ambiguous failing hosts stay Undetermined. The HTML and three CSVs
show access failures by five-minute UTC window, source site/list, and both
dimensions together. A time cluster is a clue about authentication, not proof
of its cause.

## ShareGate item pre-check (phase 2b-bis)

Run `Scripts/Diagnostics/SmartM365-SharePointMigration-ShareGatePrecheck.ps1`
only on the licensed GUI machine with Windows PowerShell 5.1. Re-run phase 2a
first so the latest `ClassifiedRows.csv` contains the new access attribution.
Both DryRun and Run require an explicit `-WhatIf`; the script refuses to start
without it. Only `-Run -WhatIf` imports ShareGate. It groups access cases by
source/destination list, connects with the current Windows user on-prem and
`Connect-Site -Browser` for SPO, then invokes exactly one
`Copy-Content -SourceItemId <id> -WhatIf` per distinct keyed item. It never
invokes a migration without `-WhatIf` or any other Copy cmdlet.

Each returned ShareGate pre-check result is exported with `Export-Report` to
a separate CSV in a private `ShareGate/Diagnostics/Precheck-<timestamp-id>/Reports`
folder. `Precheck-Items.csv` records one status per item and is updated
atomically as the run progresses. `Precheck-SkippedRows.csv` and `Precheck.log`
record exclusions and execution details. Statuses distinguish a source 401,
destination 401, 401 without an attributable side, another pre-check error,
no 401 observed, and undetermined results. No 401 observed does not prove a
full migration will succeed; the pre-check may omit work performed during a
real copy. The script does not apply user mappings or migration copy options.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\Diagnostics\SmartM365-SharePointMigration-ShareGatePrecheck.ps1 -ProjectRoot .\Migrations\MyMigration -SessionId 260930-6 -WhatIf -DryRun
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\Diagnostics\SmartM365-SharePointMigration-ShareGatePrecheck.ps1 -ProjectRoot .\Migrations\MyMigration -SessionId 260930-6 -WhatIf -Run
```

## Bounded ShareGate witness (phase 2b-ter)

`Scripts/Diagnostics/SmartM365-SharePointMigration-ShareGateWitness.ps1`
selects exactly five distinct source items from one phase 2a analysis: one
unsupported shortcut, one modern component link, two source 401 items from
the largest affected list, and one source 401 item from another list. The
selection is deterministic and displayed by `-DryRun`. The original warning
is a witness candidate, not a guarantee that ShareGate's pre-check will
reproduce it. Both modes require `-WhatIf`; only `-Run` imports ShareGate.
Source connections use the current Windows identity and destination
connections use `Connect-Site -Browser`.

The run invokes only five item-scoped `Copy-Content -WhatIf` calls. It writes
the complete `Format-List *` output of each CopyResult, an available session
ID, one `Export-Report` CSV per call, and an atomic result CSV under the
private `ShareGate/Diagnostics/Witness-<timestamp-id>` folder. The report
reader accepts both migration headers (`Status`, `Errors`) and actual
pre-check headers (`Result`, `Error`) through the configured column aliases.
An empty report remains undetermined until a positive witness produces rows.
Even then, no pre-check result authorizes a real migration.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\Diagnostics\SmartM365-SharePointMigration-ShareGateWitness.ps1 -ProjectRoot .\Migrations\MyMigration -AnalysisDirectory .\Migrations\MyMigration\ShareGate\Diagnostics\AnalysisFolder -SessionId 260930-6 -WhatIf -DryRun
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\Diagnostics\SmartM365-SharePointMigration-ShareGateWitness.ps1 -ProjectRoot .\Migrations\MyMigration -AnalysisDirectory .\Migrations\MyMigration\ShareGate\Diagnostics\AnalysisFolder -SessionId 260930-6 -WhatIf -Run
```

## Five-item ShareGate remediation pilot (phase 2c)

Real runs of the pilot and transient batch scripts are currently disabled.
Post-run inventory comparison showed that item-scoped copies placed files
at the destination library root while matching files already existed in the
expected subfolders. The local DryRun now derives each destination folder from
the classified report path, rejects ambiguous paths, and groups batch items by
destination folder. The prepared real path checks that the folder exists and
passes `-DestinationFolder` to ShareGate. It remains blocked until this behavior
is qualified with the installed ShareGate version on the GUI machine. Do not
treat ShareGate Success counts as proof that the intended paths were updated.

`Scripts/Diagnostics/SmartM365-SharePointMigration-ShareGatePathQualification.ps1`
is a separate one-file qualification for the ShareGate GUI machine. Its default
DryRun derives one source and destination path from a reviewed source-401 file
row without loading ShareGate. A real run requires the reviewed analysis hash,
`-Run -ConfirmQualification`, and a versioned console phrase. It reads the
on-premises source using the current Windows identity, authenticates to SPO
with `Connect-Site -Browser`, and checks that the source file, destination
folder, and original target file exist while no same-named root file exists.
It then makes exactly one `Copy-Content -SourceItemId` call with the explicit
`-DestinationFolder` and `New-CopySettings -OnContentItemExists Overwrite`.
This **overwrites one file in the selected SPO destination**; use it only on
a destination designated as a test copy. It does not write to the source.
It exports the ShareGate result and checks the expected file
and library root afterward. A missing or ambiguous report path is recorded as
inconclusive. The script never launches the five-item pilot or transient
batches, which remain blocked pending review of the qualification result.
The result is saved as `PathQualification-Result.json.txt` so OneDrive can
synchronize it with the rest of the private diagnostic evidence. For a run
created by v1.0.0, copy the existing `PathQualification-Result.json` to the
same folder as `PathQualification-Result.json.txt`; no ShareGate copy needs to
be repeated.

`Scripts/Diagnostics/SmartM365-SharePointMigration-ShareGatePathCorrectionPilot.ps1`
prepares a separate three-file correction after a successful one-file path
qualification. Its default DryRun imports no ShareGate module and makes no
site connection. It requires exactly three distinct source IDs and verifies
that all three are source-401 files in the same site and list as the qualified
item, with explicit destination subfolders. The DryRun prints each source and
destination route, the analysis, qualification-evidence, and plan SHA256
fingerprints, a maintenance-window estimate, and the versioned confirmation
phrase. The qualification fingerprint covers its `.json.txt` result, CSV
export, and log. No customer IDs or URLs are built into the script.

A future real run needs `-Run -ConfirmPathCorrection`, all three reviewed
fingerprints, and the exact interactive phrase. It runs only under Windows
PowerShell 5.1 with the same ShareGate module version as the qualification.
It reads the on-premises source with the current Windows identity and connects
to SPO using `-Browser`. For each selected ID, it verifies the existing target
file and folder, then calls item-scoped `Copy-Content` with the explicit
`-DestinationFolder` and `New-CopySettings -OnContentItemExists Overwrite`.
This overwrites up to three files in the selected SPO destination; it never
writes to the source. It rechecks the maintenance window and evidence before
every copy, exports a CSV per item, and stops if the report or destination
cannot prove the expected result. An existing wrong-root copy is recorded but
is not removed by this pilot. Result CSV, `.json.txt` summary, and actor log
are written under private `ShareGate/Diagnostics/PathCorrection-*`.

The old five-item pilot and 274-item batch remain blocked. Review the actual
GUI-machine DryRun before any real copy; that copy needs separate approval.

`Scripts/Diagnostics/SmartM365-SharePointMigration-RootDuplicateLiveAudit.ps1`
checks a private root-duplicate audit against current SPO metadata. It first
validates the audit structure and expected row count locally. With
`-LiveReadOnly`, it uses interactive PnP authentication, reads the selected
library once, and checks both paths, sizes, and versions for every audited
file. It writes a per-file CSV, summary, and actor log beside the private
audit. It has no SPO write or cleanup operation. Content hashes are not read.

`Scripts/Diagnostics/SmartM365-SharePointMigration-RootDuplicateCleanup.ps1`
requires the exact reviewed audit SHA256 and row count. Its default run only
validates the local audit. `-Run -ConfirmRecycle` connects interactively to
SPO, verifies every root copy and its expected subfolder original before any
change, then rechecks each root item immediately before sending that exact
file to the recycle bin. It records the actor, each result, and a final
original-preservation check in the private comparison folder. It does not
delete other comparison extras or permanently delete files.

The GUI retains shared `Migrations/logs/gui-activity` files for seven days.
At startup and once per day while open, it removes only old activity `.log`
files matching its generated filename pattern and records the cleanup when
files were removed or could not be removed.

`Scripts/Diagnostics/SmartM365-SharePointMigration-ShareGatePilot.ps1` is a
separate pilot plan. It starts in `-DryRun` mode and imports no ShareGate
module in that mode. It requires a matching analysis and witness run: both
positive witnesses must have warning rows and all three source 401 witnesses
must have header-only reports. It selects those three source 401 items plus
two more items distributed across the dominant affected list. The plan always
contains exactly five distinct source items and can be pinned to a reviewed
`ClassifiedRows.csv` SHA256 hash. The DryRun prints each source file path,
expected destination file path, and destination folder.

The previous real execution path required `-Run -ConfirmPilot` and typing the exact confirmation
phrase shown in the console. That path now stops before ShareGate is loaded.
It also required the analysis and witness SHA256
hashes displayed by `-DryRun`, so a shared input cannot change unnoticed
between review and execution. It uses the current Windows identity for the
on-premises source and `Connect-Site -Browser` for the destination. For each
selected ID, it invokes only `Copy-Content -SourceItemId <one ID>` with
`New-CopySettings -OnContentItemExists IncrementalUpdate` and a distinct
`-TaskName`. There is no `-WhatIf` during the real pilot. ShareGate's default
Insane mode waits for Microsoft 365 import completion before returning. The
script stops at the first failed call and writes one `Export-Report` CSV per
completed call, full CopyResult properties, an atomic result CSV, a plan CSV,
and an actor/machine log under the private `ShareGate/Diagnostics/Pilot-*`
folder. Review the reports and destination results before considering any
broader remediation.

The pilot refuses both `-DryRun` and `-Run` from 23:45 inclusive to 00:15
exclusive in the farm's time zone. `-FarmTimeZoneId` defaults to Windows
`W. Europe Standard Time` (Bern/Zurich); the gate checks at startup, after
interactive approval, and before every site/list connection and real copy.
An already-running synchronous ShareGate copy cannot be interrupted safely at
the window boundary, so schedule the five-item pilot with sufficient time
before 23:45.

The migration owner must review the DryRun selection and explicitly approve
real execution separately. No bulk 401 remediation is part of this pilot.

## Pilot review and prepared transient batches

`SmartM365-SharePointMigration-ShareGatePilotReview.ps1` reads an existing
five-item pilot. Its default `-DryRun` reads local evidence only. `-Run` uses
ShareGate read cmdlets to look up the destination item and, for a skipped
item, the source and destination Modified dates. It does not copy content.
The CSV labels each SPO item URL as verified, inferred, or unavailable; an
inferred URL must be checked in SPO. The original ShareGate export may leave
`Destination path` blank, in which case the review tries the source-relative
path in the destination list. The skipped-item date explanation is reported
only when both dates and destination item existence can be established.

`SmartM365-SharePointMigration-ShareGateTransientBatch.ps1` prepares an
item-scoped follow-up after a reviewed pilot. Default `-DryRun` prints each
source-list-and-destination-folder batch, selected source IDs, the evidence SHA256 values, the plan
SHA256, and an exact versioned confirmation phrase. It identifies elements
by site, list, and source ID. It excludes the five pilot item keys and
separately identifies any other `Home.aspx` in a site-pages list. The
DryRun displays that page's title, type, source site URL, relative path,
and destination site URL before excluding it for separate review. The
DryRun also requires `-QualificationDirectory` and `-PathCorrectionDirectory`.
It verifies that the one-file qualification and three-file correction prove
the destination paths of all four previously copied pilot items, checks their
ShareGate exports for `Success`, the exact subfolder and a finished Microsoft
365 import, and fingerprints all six correction files plus the three
qualification files. The plan fingerprint includes both placement proofs.
The script uses the slowest observed per-item time from the original and
path-correction pilots for a conservative maintenance-window estimate.
The prepared batch path still uses `IncrementalUpdate`; a completed item's
export must identify its exact planned destination path and finished import.
An unproven item or a skipped item stops the run for review before another
batch can start. `-Run` remains disabled until the placement-proof DryRun and
copy mode are reviewed separately.

The seven source access lines without an item ID (six Site and one File in
the reviewed session) are retained in `Transient-HorsLot.csv` with the
status `Hors lot - à traiter à part`; their counts appear in the console
and GUI summary. Batch size defaults to 50. `-Run` currently stops before
ShareGate is loaded. The prepared path requires `-Run -ConfirmBatch`, both
expected item counts, all six reviewed hashes, and the exact interactive
phrase. It uses the current Windows identity for the source,
`Connect-Site -Browser` for SPO, and only `Copy-Content -SourceItemId <IDs>`
with `IncrementalUpdate` and a distinct task name per batch. When enabled,
a real run refuses to start when its projected finish plus
`-MaintenanceMarginMinutes` (60 by default) reaches the next 23:45 farm
maintenance window. It repeats the estimate before each batch, checks the
active 23:45-00:15 window before each batch and copy, uses an exclusive
session lock, and stops after a batch
exceeds `-MaxErrorsPerBatch` (default 0). Each batch report and the atomic
global results CSV/summary stay under private `ShareGate/Diagnostics`.
The GUI Diagnostics tab reads the latest summary and shows Success, Skipped,
Error, Warning, Mixed, Unreported, NotAttempted, and hors-lot counts. A started batch
with no usable export is marked Unreported for manual review.

A successful offline test or DryRun does not authorize a real copy run.

The launcher uses
`Comparison.ModifiedDateToleranceMinutes` to produce `ChangedModifiedDate` and
`TargetOlderThanSource` review outputs. For SP2019 to SPO checks, the template
normalizes source `Modified` dates as local time and target `Modified` dates as
UTC before comparing them, while keeping the raw values in the diagnostic CSVs.

`Comparison.PermissionMaxScanAgeDifferenceHours` similarly guards permission
scan pairs (24 hours by default), before any Entra cache refresh; absolute age
is limited by `Comparison.PermissionMaxScanAgeHours` (48 hours by default).
Scheduled `-NonInteractive` runs fail on excessive gaps. Interactive runs require
typing `YES`; `-Force` is an explicit warned override after review.

The shipped template retains `AllowDuplicateKeysForDeleteScript = $true` for
compatibility. This permits generation despite duplicate keys; review
`DuplicateKeys.csv` and prefer setting it to `$false` when ambiguity must block
generation. Generation is not authorization to execute. Generated file cleanup
defaults to dry-run; `-Execute` permanently deletes unless `-Recycle` is also
chosen. Read every generated script and target before running it.

## Offline Validation

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Tests\Test-SharePointMigration.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File .\Tests\Test-SharePointMigration.ps1
pwsh -NoProfile -ExecutionPolicy Bypass -File .\Tests\Test-LauncherHosts.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\Test-SmartM365ShareGatePrecheck.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\Test-SmartM365ShareGateWitness.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\Test-SmartM365ShareGatePilot.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\Test-SmartM365ShareGatePathCorrectionPilot.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\Test-SmartM365SharePointTransientOffline.ps1
python -B -m unittest discover -s .\Tests -p test_comparisons.py
```

The tests use synthetic files, mocked external calls and mock inventory scripts.
They do not run migrations, scan real sites, exercise real cleanup, or certify
tenant permissions. The host-routing test needs PowerShell 7 and Windows
PowerShell 5.1. ExecutionPolicy Bypass above is limited to the test processes;
these results are not an AllSigned execution-policy certification.
