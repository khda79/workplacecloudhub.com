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
The dashboard component is **1.0.8** and the generic launcher is **1.0.18**.
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
Files, Permissions, Operations, Logs and Config tabs; it discovers configured
migration folders and displays the most recent output paths. An output timestamp
or an available Open button is not proof that the whole scan succeeded: inspect
the run log and any error CSV before accepting its results.

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

Create one local folder per migration by copying the template:

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
Runtime outputs stay inside the local migration folder:

```text
scans\
comparisons\
operations\generated\
logs\
```

`Migrations/*` is ignored by Git except for `_Template` and the template update
launcher. Do not commit real migration folders, inventory CSVs, workbooks, logs,
generated cleanup scripts, or local authentication files.

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
`Comparison.ModifiedDateToleranceMinutes` to produce `ChangedModifiedDate` and
`TargetOlderThanSource` review outputs. For SP2019 to SPO checks, the template
normalizes source `Modified` dates as local time and target `Modified` dates as
UTC before comparing them, while keeping the raw values in the diagnostic CSVs.

`Comparison.PermissionMaxScanAgeDifferenceHours` similarly guards permission
scan pairs (24 hours by default), before any Entra cache refresh. Both age guards
use CSV last-write timestamps, not a certified scan-start time or absolute age.
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
