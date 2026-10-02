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
Files, Permissions, Operations, Migration Diagnostics, Logs and Config tabs; it discovers configured
migration folders and displays the most recent output paths. An output timestamp
or an available Open button is not proof that the whole scan succeeded: inspect
the run log and any error CSV before accepting its results.
The dashboard requires PowerShell 7.4 or later. The CMD launcher checks this
requirement before opening the GUI, and the GUI enforces it for direct launches.
Windows PowerShell 5.1 remains in use as a child process for SharePoint Server
inventory and SPO site administration scripts that require it.

The header displays the mapped scan scope. When a migration maps several webs,
hover over the scope count to see their URLs. A target `SiteUrl` that differs
from every mapped target is flagged; the GUI blocks site operations until the
configuration is aligned. The header has an enabled-by-default 30-second
refresh checkbox. Refresh keeps unsaved configuration edits and the selected
log; turn the checkbox off to stop automatic updates.

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
- `Scripts/Compare/`: source versus destination comparisons and source scan
  history comparisons.
- `Scripts/Export/`: CSV to Excel export helpers used by comparison workflows.
- `Scripts/Generate/`: generated operation script builders for reviewed
  destination cleanup.
- `Scripts/Operations/`: guarded destination cleanup and SharePoint
  administration operations.
- `Scripts/Launchers/`: migration-aware launchers shared by each migration
  folder.
- `Migrations/_Template/`: safe template used to create local migration folders.
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
wizard copies `_Template`, writes
the configuration and mapping, and selects the new migration only after all
checks succeed. For a multi-web migration, each additional mapping must have
one source URL and one target URL on its own line.

You can also create one local folder per migration manually by copying the template:

```powershell
Copy-Item -Recurse .\Migrations\_Template .\Migrations\MyMigration
```

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
The branded HTML summaries are self-contained and show at most 20 objects with
differences; use their linked CSV and Excel exports for the full detail.
The comparators load the shared `report_html.py` from their own directory. They
add that directory explicitly because the bundled Portable Python runs in
isolated mode and does not add the script directory to its import path.
The generic launcher prints a WorkplaceCloudHub introduction and a timestamped
execution summary with status, duration and run log path. A successful script
run means the comparison completed; inspect its report for migration differences.
Runtime outputs stay inside the local migration folder:

```text
scans\
comparisons\
operations\generated\
logs\
```

`Migrations/*` is ignored by Git except for `_Template`, the template update
launcher, and each migration's `ShareGate` folder structure. Put ShareGate
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
  refresh. Comparison/export Python helpers use the standard library; Excel
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
python -B -m unittest discover -s .\Tests -p test_comparisons.py
```

The tests use synthetic files, mocked external calls and mock inventory scripts.
They do not run migrations, scan real sites, exercise real cleanup, or certify
tenant permissions. The host-routing test needs PowerShell 7 and Windows
PowerShell 5.1. ExecutionPolicy Bypass above is limited to the test processes;
these results are not an AllSigned execution-policy certification.
