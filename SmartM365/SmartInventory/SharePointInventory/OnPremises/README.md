# SharePoint Server 2016/2019 inventory

Two independent, read-only Windows PowerShell 5.1 x64 collectors run on a SharePoint farm server. They require the SharePoint Management Shell snap-in and administrator-prepared SharePoint Shell access. They never grant permissions or change SharePoint locks. They do not run through the SmartInventory orchestrator.

## Collectors

- `SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.ps1` writes `SharePoint_OnPrem_Farms.csv`, `Servers.csv`, `ServiceApplications.csv`, `WebApplications.csv`, `WebApplicationZones.csv`, and `ContentDatabases.csv` with the common `SharePoint_OnPrem_` prefix.
- `SmartM365-SharePoint-OnPrem-Content-Inventory.ps1` independently discovers web applications and content databases and writes `SharePoint_OnPrem_SiteCollections.csv`, `SiteAdministrators.csv`, `Webs.csv`, and `CollectionCoverage.csv`.
- Each script has its own local JSON template, run folder, log, receipt, current CSV publication, and weekly history. `FarmId`, `WebApplicationId`, and `ContentDatabaseId` are shared join keys. Content does not consume Infrastructure CSVs.
- `Get-SPWebApplication` selects content applications by default and excludes Central Administration. `IncludedWebApplicationUrls` and `ExcludedWebApplicationUrls` in each local JSON narrow the scope. Empty inclusion means all content applications.
- Both templates set `EnableSharePointUpload=false`. The scripts also force upload off at runtime; CSVs, receipts, logs, and history are never transferred automatically.

Every CSV includes `TenantKey`, `FarmId`, `RunId`, and `CollectedAtUtc`. Infrastructure files contain farm build and counts (`Farms`), server identity/role/status (`Servers`), service application identity/type/pool (`ServiceApplications`), application URL/pool/database count (`WebApplications`), zone URL/authentication (`WebApplicationZones`), and database identity/server/site count/read-only status (`ContentDatabases`). Content files contain collection identity, owner, size, quota, modification, template, language and lock observations (`SiteCollections`); administrator logins and owner flags (`SiteAdministrators`); web identity, template, list/library counts and permission inheritance (`Webs`); and one coverage result per attempted collection or database failure (`CollectionCoverage`). The exact column lists are declared in each collector's `$schemas` block.

## Publication and qualification

The scripts write run CSVs first. Empty datasets have headers. They promote current CSVs into tenant-specific `DATA-LAST` only when the collection has no recorded coverage or property failures. Their separate source receipts are completed only after current CSV validation. A failure leaves the prior completed receipt intact and writes a failed run state. The receipt contract proves current file identity, rows, and hashes; it does not prove farm-wide exhaustiveness by itself. Inspect `CollectionCoverage` and the run log.

For a full Content run, each database's observed site collection count must match `CurrentSiteCount`. An unavailable count or mismatch creates a `CollectionCoverage` database status and blocks publication. A database read error is recorded as `DatabaseEnumerationFailed`. Per-collection failures remain visible in both `SiteCollections.CollectionStatus` and `CollectionCoverage.Status`.

Content `-MaxItems N` processes at most N site collections across the selected databases. It writes only to a timestamped run folder under `OutputRoot\TEST`, marks the run limited, and leaves `DATA-LAST`, weekly history, and the source receipt unchanged. `-ValidateOnly` checks the PowerShell host, SharePoint Shell access, selected content applications, and databases without traversing any site collection or publishing a CSV.

Content has cooperative `-GlobalTimeoutMinutes` (default 720) and `-CollectionTimeoutMinutes` (default 60) checks between SharePoint calls. A single blocked SharePoint object-model call cannot be interrupted safely inside this process. An isolated child-process timeout is outside this lot. A timed-out or inaccessible collection is recorded in `CollectionCoverage`; later collections continue when the process is responsive. A global timeout stops subsequent collection and prevents current publication.

Lock inspection reads only `SPSite` and `SPContentDatabase` properties. `LockState` is normalized only when a direct site lock-state property is available, or when an observed read/write lock unequivocally indicates `NoAccess` or `ReadOnly`. Boolean values alone cannot safely distinguish `Unlock` from `NoAdditions`; in that case `LockState` is blank and `LockStatus=Unverified`. `IsReadOnly`, `ReadLocked`, `WriteLocked`, `LockIssue`, and `ContentDatabaseIsReadOnly` are kept separately. Missing properties remain blank and block a complete content qualification. The mapping, quota units, storage measurement, and administrator enumeration need a real read-only check on both SharePoint 2016 and 2019 before the output contract is frozen. No `Set-SPSite` or other write cmdlet is called.

## Manual commands

Run from a SharePoint farm server after the administrator has prepared the local tenant configuration and SharePoint Shell rights. Replace `<SmartM365Root>` with the deployed path. Each launcher copies this collector folder, `Config`, and `SmartM365.Core` to its own local cache before invoking Windows PowerShell 5.1. The caches are separate for Infrastructure and Content. The launchers default to `prod`, forward additional parameters, and return the collector exit code.

Copy each `*.local.json.template` to the same basename ending in `.local.json`, then review the tenant paths and optional `IncludedWebApplicationUrls` / `ExcludedWebApplicationUrls` arrays. The default upload setting is `false`; the scripts enforce that setting even if the local JSON is changed.

```powershell
& "<SmartM365Root>\SmartInventory\Launchers\OnPremises\Start-SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.cmd" -ValidateOnly
& "<SmartM365Root>\SmartInventory\Launchers\OnPremises\Start-SmartM365-SharePoint-OnPrem-Content-Inventory.cmd" -ValidateOnly
& "<SmartM365Root>\SmartInventory\Launchers\OnPremises\Start-SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.cmd"
& "<SmartM365Root>\SmartInventory\Launchers\OnPremises\Start-SmartM365-SharePoint-OnPrem-Content-Inventory.cmd" -MaxItems 2
& "<SmartM365Root>\SmartInventory\Launchers\OnPremises\Start-SmartM365-SharePoint-OnPrem-Content-Inventory.cmd"
```

Run each collector on 2016 and 2019 separately. Inspect the two `-ValidateOnly` results before limited or full runs. The `-MaxItems 2` run is not a complete-farm qualification.
