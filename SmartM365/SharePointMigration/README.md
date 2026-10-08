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
Release **1.0.64** includes dashboard **1.0.64** and generic launcher **1.0.28**.
See [release notes](RELEASE-NOTES-1.0.64.md) for the changes and validation boundary.

## Working source reliability updates

The working source after release 1.0.64 includes these review corrections:

- Overview, inventory selectors and generic launchers order completed scans by
  receipt completion time, then legacy filename time. Synchronization or copy
  times do not make an older scan the latest. Error sidecars and invalid receipts
  are excluded; selecting a receipt does not recompute its inventory hash.
- Source and destination SharePoint group membership caches are scoped to the
  site collection. Membership lookup failures produce error evidence and block
  final inventory publication. Transient failures receive bounded retries;
  denied access and other permanent failures stop immediately.
- Permission page enumeration resumes after the last fully processed item ID
  following a transient transport failure, preserving streaming and avoiding
  duplicate exports. Role-assignment collections also use bounded read retries.
- SPO discovery and inventory reuse connections only for the exact web URL
  within the same scan execution and authentication configuration.
- Optional token diagnostics retrieve the SharePoint token through the supported
  PnP command, tolerate unavailable details and never print the token.
- Receipt row counts support comma, semicolon and tab CSVs, including quoted
  fields that span multiple lines.
- Global file and permission reports retain historical comparison rates and
  expose `ComparisonFreshness`: `Uses latest scans`, `Recalculate`,
  `Scan unavailable`, or `Review scan freshness`. Scan dates prefer the recorded
  comparison completion times, then receipt times, then labelled legacy
  filenames. Current scan gap/age columns and configured age/gap limits make
  freshness distinct from the historical comparison time.

Offline regression checks cover copy-time inversions, invalid/error receipts,
site collections sharing group IDs, failed and recovered membership reads,
optional token claims, connection reuse, resumed pages with nonconsecutive IDs,
terminal retry failures, multiline CSVs and global report freshness. These checks
are separate from qualification on a SharePoint farm or SPO tenant. Existing
inventories must be renewed to benefit from the membership collection fixes.

## Install, Start, and Update

Version **1.0.64** is distributed through the
[GitHub release](https://github.com/khda79/workplacecloudhub.com/releases/tag/sharepoint-migration-toolkit-v1.0.64)
and the [existing source folder](https://github.com/khda79/workplacecloudhub.com/tree/main/SmartM365/SharePointMigration).
Download the release ZIP together with its `.sha256` and `.manifest.json` assets.
Verify the ZIP with `Get-FileHash -Algorithm SHA256` before extracting it, then open
`SmartM365/SharePointMigration` inside the extracted package. The root also includes
the public signing certificate, its optional trust installer, licence and notice.
No PowerShell Gallery package is published for this toolkit; do not use `Install-Module`.

Obtain the repository and keep the entire `SmartM365/SharePointMigration` folder;
copying just the GUI script omits required helpers, assets and templates. Launch
`Start-SmartM365-SharePointMigration-GUI.cmd` on Windows. The dashboard provides
Overview, Files & Permissions, Migration Diagnostics, Batch runs, Operations, Logs and Config tabs; it discovers configured
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
and permission scan dates in each source and target scan cell. Each date appears
green when the scan is from today, independently for files and permissions on
both sides. The displayed gap
in decimal days compares the file scans. Rates and comparison dates are separate
for files and permissions. The file rate is matched
files divided by source unique keys;
the permission rate is matched permissions divided by source unique permission
keys. An empty inventory or missing comparison has no rate.
File and permission rates appear in large, color-coded cells after Migration,
Gap (days), and Status, with comparison dates underneath and small progress bars.
The Global comparison column is the equal-weight arithmetic mean of the two
rates: `(files % + permissions %) / 2`. It has no value if either rate is unavailable.
Newer scans mark the affected rates and their global average as `Recalculate`.
Migration, Gap (days), Status, and the three rate columns stay visible during
horizontal scrolling. Compact rows and flexible URL columns fit nine migrations
on a maximized 1920 x 1080 display at standard scaling.
Gap badges use the unrounded file-scan time difference: green up to 12 hours,
yellow above 12 hours through 24 hours, and red above 24 hours; unavailable gaps
are gray. Status badges are green for `Up to date`, yellow for `Compare needed`
or `Review needed`, and red for `Scan needed` or `Refresh error`.
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
Comparison badges use `today` and `yesterday`, like scan badges, followed by
the matching percentage. Source-versus-target **Run** requires both selected
scans to exist without an `-Errors.csv` sidecar; a disabled comparison displays
the missing or incomplete scan reason. History **Compare** requires two different
completed scans on the selected side. Existing comparison reports remain accessible.
Source-versus-target comparisons also require a non-empty source inventory.
An empty destination file inventory is allowed so missing files can be reported;
an empty destination permission inventory is blocked, matching the launcher.
The GUI checks receipt row counts, or reads only the first record for older scans,
and displays an empty-inventory reason before launching a worker. Empty file
inventories remain available for scan history and inventory metrics.
The date and rate come from the latest readable comparison summary, rather than
an empty output folder created by a failed attempt. A newer attempt without a
summary is shown separately; the previous available report remains labelled as such.

In **Files & Permissions**, **Open source site**
and **Open destination site** appear in scan cards 1 and 2 for both files and
permissions, immediately before **Open folder**. They open the mapped site in the
default browser. If that side includes several URLs, a menu lets the operator
choose a site. Invalid or missing HTTP/HTTPS URLs disable the corresponding button.

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

### Batch destination scans

From the `SmartM365/SharePointMigration` directory, preview the target file and
permission scans for every configured SPO migration:

```powershell
.\Start-SmartM365-SharePointMigration-TargetScanBatch.cmd -PlanOnly
```

Remove `-PlanOnly` to run the batch. Both `Interactive` (the default) and
`Certificate` authentication run **at most two scans at a time**, with a
15-second gap between process starts. Each interactive scan opens its own
console; complete any requested sign-in in that window. It runs all target file
scans first, then all target permission scans. To select or retry only some migrations, use
`-MigrationNames SiteA,SiteB`; use `-InventoryMode FilesOnly` or
`PermissionsOnly` to select one scan type.

If the machine has a configured app certificate and the required SPO access,
use `-AuthMode Certificate` for background scans. `-MaxParallel 1` reduces
concurrency in either mode; values above two are rejected. This limit covers only
this batch; stop other GUI, scheduled, or cross-machine scans before starting it. The batch continues
after individual scan failures and returns a nonzero exit code if any scan
fails. Its global `batch.log` and `summary.csv` are in
`Migrations/logs/target-scan-batches/<batch-id>/`; each launcher also keeps its
own migration log. Certificate scans also capture stdout and stderr in the batch
directory; interactive output is shown in each scan console. Check the per-scan
error CSVs and manifests before using the inventories for comparisons.

### Batch comparisons for all migrations

Preview source-versus-target file and permission comparisons:

```powershell
.\Start-SmartM365-SharePointMigration-ComparisonBatch.cmd -PlanOnly
```

Remove `-PlanOnly` to execute all file comparisons, then all permission
comparisons, one at a time. The existing launcher chooses each migration's
latest inventories and generates its usual HTML reports and other comparison
outputs. Missing inventories, recorded scan errors, invalid manifests, or scan
ages outside the configured limits fail that comparison; the batch continues
with the others and returns a nonzero exit code if any comparison fails.

Use `-ComparisonMode FilesOnly` or `PermissionsOnly`, and
`-MigrationNames SiteA,SiteB` for a subset. `-AuthMode Certificate` selects
certificate authentication for any required Entra users cache refresh;
interactive authentication is the default. Permission comparisons can request
sign-in if that cache needs refreshing. No source or target inventory scan is
started. `-Force` explicitly overrides the launcher's scan age limits only;
review stale evidence before using it.

Global `batch.log`, per-comparison `*.console.log`, and a progressively saved
`summary.csv` are written to `Migrations/logs/comparison-batches/<batch-id>/`.
The migration launcher retains its own logs and reports. The batch suppresses
Explorer windows and fails stale comparisons without asking for confirmation;
interactive authentication remains available for Entra cache refresh.
PowerShell 7.4 or later and the comparison launcher's Python dependencies are
required. Run on the workstation or server that has access to the migration
files; SharePoint farm snap-ins are not required for comparison actions.

### Batch source scans on a SharePoint farm server

`SmartM365-SharePointMigration-SourceScanBatch.ps1` replaces the private
`Get-SP2019FileInventory-LocalWrapper.ps1` workflow. Run it on a SharePoint
Server farm machine under an account with SharePoint Shell access. It needs
Windows PowerShell 5.1, the registered SharePoint snap-in, and Python 3 (the
portable runtime in `Tools/Python` is used when available) to write scan
manifests. From the `SmartM365/SharePointMigration` directory, preview first:

```powershell
.\Start-SmartM365-SharePointMigration-SourceScanBatch.cmd -PlanOnly
```

Remove `-PlanOnly` to run source file scans for every configured SharePoint
Server migration, followed by their permission scans. Scans run sequentially
to avoid concurrent farm load. Use `-MigrationNames SiteA,SiteB`
for a subset, or `-InventoryMode FilesOnly` / `PermissionsOnly` for one kind.
If the wrapper is copied outside the project, pass `-ProjectRoot` pointing to
the shared `SmartM365/SharePointMigration` folder. It copies signed source
scripts, the manifest helper, and the portable Python runtime when present to
a unique folder under the farm server's local temp directory before scanning;
it never runs those dependencies from the share or edits the repository scripts.
A configured `Source.UrlsFile` or `Comparison.PathMappingsFile` is required
for every selected migration, preventing an unintended whole-web-application
scan.

The batch keeps `batch.log` and `summary.csv` in
`Migrations/logs/source-scan-batches/<batch-id>/`, plus each scan's own log,
CSV, and verified manifest. It continues after individual failures and exits
nonzero if any scan fails. Validate a `-PlanOnly` run and one real migration
on the farm server before replacing the old scheduled task.
Both root `.cmd` launchers forward their arguments and return the batch exit
code. With no arguments, they start the full batch; preview with `-PlanOnly`
first.

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

Source and destination file scans write `<inventory.csv>.metrics.json.txt` next
to their CSV. The GUI reads this small metadata file for file count, folders
containing files, and current file volume; it does not recalculate inventory
metrics. Empty folders, file versions, and recycle-bin content are excluded.
Older scans without metrics show an unavailable value until a new file scan is
run. New metrics include the CSV SHA256 to identify the exact scan content,
independently of synchronized modification times. Hash verification is cached
until the CSV length or local modification time changes. For earlier metrics,
an exact timestamp remains valid; a copy truncated to whole seconds also needs
a matching scan receipt, row count and verified CSV SHA256. Other mismatches
remain unavailable. The GUI verifies identity but never calculates the counts
or volume from the CSV.

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
the inventory roots. Site Assets and Site Pages are included by default in source
and destination file and permission scans, including library-only permission
scans. Other hidden/system content is excluded unless opted in; scan logs identify
each excluded list/library and its reason (hidden, system, empty for file scans,
or not a document library in library-only permission scans). Existing inventories
must be rescanned on both sides before comparing this expanded scope.
Permission list/library and item settings affect what is assessed. PageSize
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
warning. A zero-row source file scan with a receipt still blocks comparison.
An older source scan without a receipt remains inconclusive if empty. A zero-row target
file scan with a valid receipt can be compared before migration: all source
files are reported as missing and the file match rate is 0%. An unverified
empty target remains inconclusive. Permission comparisons still reject
zero-row scans.
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
PNG or JPEG in `Config`. Each newly generated HTML header displays one logo:
the client image when configured and valid, otherwise the WorkplaceCloudHub image.
Newly generated HTML reports embed the images, so viewers do not need access
to the original files. The local `.json.txt` configuration and client image
are ignored by Git and must be copied privately when deploying another clone.

## Migration Diagnostics (local report analysis)

The read-only **Cross-check: ShareGate / Files / Permissions** panel loads the
latest valid existing file and permission comparisons for the selected migration.
It shows each report's date, its own measure, the main differences and whether
its source/target scans are still the latest. ShareGate to-fix counts are shown
separately from file and permission match percentages; they have different
denominators and are never merged into a single success rate. The GUI loads
comparison detail in a background thread to keep migration selection responsive.
Its scope table groups to-fix ShareGate rows, `LibrarySummary.csv` and
`PermissionSummary.csv` by normalized destination site URL and list title when
that scope can be verified from report fields. Rows without destination scope
fall back to source scope and remain separate from destination-keyed rows. It shows
missing, extra and changed file or permission flags without claiming that a
ShareGate row and an inventory difference refer to the same individual item.
Ambiguous list titles, rows without a usable site/list scope, missing comparisons,
legacy or stale scan evidence, and newer scans are marked for review. Opening
the file or permission HTML report gives the complete comparison details.
The panel does not start scans, ShareGate actions or migration writes.

The **Migration Diagnostics** tab names the selected migration and reads the
newest ShareGate CSV or XLSX export in its `ShareGate/MigrationReport` folder.
It displays the file name, modification time and size. If the folder is empty,
it asks the operator to deposit the latest report and disables analysis and
issue review. For a same-name CSV/XLSX pair, the CSV is selected. An existing
analysis of that exact input is restored when the source SHA256 still matches;
older analyses without a fingerprint are labelled as date-based evidence.
The farm Run button requires a SHA256-verified analysis; reanalyze a legacy
report before using its access peaks.
Analysis uses local report files; it
does not import the ShareGate module, connect to a site, precheck, or retry a
migration. The analysis runs in a child PowerShell process so the GUI remains
responsive. For the selected migration, detecting a latest report without a
matching analysis starts analysis automatically, across all sessions. Only one
analysis runs at a time in the GUI; changing migrations queues the selected
report until that worker finishes. Each report path, size and modification time
is attempted once per GUI session. A failure is displayed with a manual retry
through **Analyze latest report**, rather than repeated on every refresh.
Shared GUI activity logs record the operator and the analysis result.
CSV is preferred when a CSV and XLSX have the same base name. The GUI
analyzes only the latest selected report file; the command-line folder DryRun
lists every file, its selection status, detected session IDs, and the installed
`ImportExcel` version. With ImportExcel available, a same-name XLSX is masked
only after its row count and row identity columns match the CSV. The remaining
cells may differ between export formats; the CSV remains the selected evidence.
XLSX analysis automatically installs `ImportExcel` from the official PSGallery
for the current user when missing, then imports and verifies its commands.
The GUI keeps **Analyze latest report** available for XLSX and displays the
installation phase while the child process works. Installation errors stop
analysis and show an actionable message; no partial CSV-only analysis is
silently substituted. `DryRun` never installs a module or registers a repository.
PowerShell's execution policy and permanent repository trust settings are preserved.
**ShareGate report & analysis status** describes the latest ShareGate input,
its matching analysis and generated HTML. **ShareGate analysis detail** is filled
after analysis; detecting a new CSV or XLSX does not reuse indicators from an
older input. File and permission comparison states are shown in Cross-check.
ShareGate analysis also displays an indeterminate progress bar at the top of the
tab, including ImportExcel preparation and the current analysis phase. It hides
when analysis finishes or fails, and when a different migration is selected.
Cross-check loading text and its progress bar appear at the top of the tab,
above Summary. Its **ShareGate to fix** column shows `—` when analysis is unavailable;
`0 items / 0 lines` is reserved for an existing analysis with no issues in that scope.
Report rows from different
files are deduplicated by session and row ID; conflicting duplicates are
counted and flagged for review. A session selector can restrict a manual analysis
to one session. **Issue review** and **Raw report rows** use the full tab width.
Selecting an issue pattern displays its matching raw report rows immediately.

The `ShareGate 401 retry results` card appears only when a previous reviewed
batch result exists. It summarizes historical retries; it does not start one.

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
Diagnostics tab opens **Source farm diagnostics** in a dedicated window using
the selected migration. Opening this window only displays existing evidence;
it does not start collection. The window shows the latest matching farm result and generates one-line
DryRun and real commands from the project's `AccessFailures-5min.csv` windows.
The GUI builds farm commands only from the analysis of the currently selected
latest report, so a newer report cannot silently reuse older access peaks.
Its `Run (read-only)` button remains disabled
until `Check prerequisites` confirms that the GUI is elevated on a farm server,
Windows PowerShell 5.1 can load the SharePoint snap-in or module, the current
account can read the farm, and the shared UNC toolkit and access-peak CSV are
available. It checks again immediately before launching a visible Windows
PowerShell 5.1 console. Run the displayed DryRun command first on the farm
server. Launch the GUI from the shared UNC toolkit root or pass
`-FarmToolkitRoot` with that root to generate UNC commands and enable the check.

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
When ShareGate exports local dates without a UTC offset, analysis converts
access-failure timestamps using the analysis machine's Windows time zone. If
the export came from another machine, pass its zone with
`-ReportTimeZoneId 'Romance Standard Time'` to
`SmartM365-SharePointMigration-Diagnostics.ps1`. The summary records the zone
and conversion count; `ClassifiedRows.csv` keeps both `Timestamp` and
`TimestampUtc`. Ambiguous or invalid clock-change times stop the analysis.
Re-run phase 2a after selecting the correct zone before starting farm diagnostics.

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

### Dashboard batch tab

The **Batch runs** tab follows **Migration Diagnostics**. Its scope is independent
of the migration selected in the header: all configured migrations by default,
or a selected subset. Three cards launch the existing source scan, destination
scan and comparison CMD launchers in dedicated consoles.

- **Both** is the default; **Files** and **Permissions** are also available.
- **Preview plan** passes `-PlanOnly`; it does not start a scan or comparison.
- Destination scans explicitly use `-MaxParallel 2` in Interactive and Certificate
  modes. Source scans and comparisons remain sequential.
- Source **Run batch** is disabled until **Check prerequisites** succeeds on the
  local source farm server: elevated Windows PowerShell 5.1, SharePoint snap-in
  and Shell/database access, trusted toolkit dependencies, Python availability
  and writable logs. The batch validates local staging again before collection.
  **Copy command** uses the editable source toolkit path for execution on that
  server; the dashboard does not remotely run source scans.
  The source CMD launcher copies and verifies the batch entry script locally;
  the batch then stages its inventory scripts and dependencies locally.
- **Open logs** and **Open summary.csv** show the latest batch evidence. Counts
  are actions, not migrations. A progressive CSV alone is not considered proof
  that a batch completed.

Dashboard requests and completion receipts are stored in
`Migrations/logs/batch-gui-runs/`. A unique `-BatchId` connects each run to its
batch log directory. The receipt records the child launcher's exit code even
when the console stays open. Closing the dashboard does not stop batch consoles.
The destination limit applies per batch, not across computers or other launchers.

The offline GUI test uses synthetic CMD launchers, validates failure receipts,
scope selection and prerequisite gating, and can export WPF previews:

```powershell
pwsh -NoProfile -File .\Tests\Test-SmartM365BatchGuiOffline.ps1
```

### Inventory reliability and diagnostic report snapshots

- SPO permission scans retain the actual associated member, owner and visitor
  group titles. A failed group read records an inventory error; a genuinely
  absent associated group remains empty. Names are cached per web for the run.
- Item inheritance is loaded in CSOM requests of at most 100 items within each
  enumeration page, avoiding oversized requests with the default 2,000-item pages.
  Pages are processed immediately; role assignments are read only for items with
  unique permissions. A missing-item batch falls back to individual reads to
  identify the affected IDs. Error CSVs include `ItemId` and `ItemUrl`, with an existence check for
  missing-item errors. Any recorded inventory error still prevents final CSV
  publication and manifest creation.
- Transient web, subsite and library discovery reads use at most three attempts,
  with waits of 5 and 15 seconds. Access errors are not retried. Token diagnostics
  appear once per host, account and authentication configuration during the run.
- Role assignment reads load Member and RoleDefinitionBindings together, with
  the same three-attempt limit for transient transport failures. Persistent errors
  retain principal identity, object path and inner exception details in the log
  and error CSV, and prevent final inventory publication.
- All four inventory scripts format elapsed durations without rounding hours.
- ShareGate XLSX analysis uses a verified local snapshot, with up to three
  attempts when an export is locked, changing or incomplete. Generated evidence
  keeps the original path and SHA256; a source/snapshot mismatch prevents report
  publication. Temporary workbooks are removed after analysis.
- New destination batch summaries include `RunLog`, `OutputCsv` and `ConsoleLog`.
  `OutputLog` points to the actual inventory log when available, including in
  interactive mode. Each action uses a unique launcher receipt; unavailable or
  mismatched receipts are reported without guessing a path from another run.
  The concurrency default remains two scans in either authentication mode.

Offline reliability and batch tests:

```powershell
pwsh -NoProfile -File .\Tests\Test-SmartM365InventoryReliabilityOffline.ps1
pwsh -NoProfile -File .\Tests\Test-SmartM365TargetScanBatchConcurrency.ps1
pwsh -NoProfile -File .\Tests\Test-SmartM365WorkbookSnapshotOffline.ps1
python -B -m unittest discover -s .\Tests -p test_sharegate_diagnostics.py
```

After deploying the toolkit, qualify these changes with a small successful SPO
permission scan, the previously failing missing-item scope, and a large library.
Check associated group columns, error IDs/paths, final publication, batch log
links and actual wall-clock duration. These offline tests do not establish a
production speed improvement or live tenant completeness.

### Other offline checks

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
