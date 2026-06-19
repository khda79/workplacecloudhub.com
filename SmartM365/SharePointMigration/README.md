# Smart SharePoint Migration Toolkit

Smart SharePoint Migration Toolkit helps validate SharePoint migrations between
a source and a destination. It inventories source and destination content,
compares files and permissions, exports review workbooks, and generates guarded
cleanup scripts for destination-side differences.

File comparisons include the current SharePoint document version exported from
both inventories. They also flag matched files where the destination modified
date is older than the source, so a copied file that is not at the latest source
version is visible in the comparison workbook.

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
Migrations\MyMigration\migration.config.source.txt
Migrations\MyMigration\migration.config.target.txt
```

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
- PowerShell 7 for the generic migration launchers when available.
- PnP.PowerShell for SharePoint Online inventory and permission scans.
- Windows PowerShell 5.1 and Microsoft.Online.SharePoint.PowerShell for SPO
  admin operations such as site lock state and page comment settings.
- Python 3 for comparison, Excel export, and generated-operation helpers.

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
