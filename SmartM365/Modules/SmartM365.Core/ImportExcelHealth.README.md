# ImportExcel startup health and missing-file repair

`Initialize-SmartM365ImportExcel` is shared by SmartM365.Core 1.0.88 and the Windows PowerShell 5.1 compatibility module 1.0.58. It checks a real import and the Excel commands in a fresh, same-edition process, then pins the verified manifest in the caller. No tenant-specific paths or versions are built into the helper.

The host check first loads the built-in `Microsoft.PowerShell.Security` manifest from the current engine and reads the effective execution policy. A host/type-data failure terminates with `SmartM365.ImportExcel.HostEnvironment` before discovery, installation or repair; it is not evidence of missing ImportExcel files. No forced module reload, type-data removal or policy fallback is used. ProxyAddresses 1.31 runs this startup check before loading the Exchange snap-in. An already contaminated Exchange session fails explicitly instead of attempting to repair its type table.

## Automatic repair boundary

- Healthy installation: no repair directory, package download or installation write.
- Incomplete installation: acquire a cross-process lock, recheck, then look for a healthy local copy with the exact same version and GUID. Other PowerShell module folders may supply a reference, but cannot silently replace the normal selected version.
- Compare SHA256 hashes first. Any existing runtime file differing from the reference blocks repair. Only missing files are restored, using exclusive creation; prior files are backed up and hash-verified.
- Without a healthy local reference, download the **exact installed version** into staging from the configured, already trusted repository. Repository trust, certificates, execution policy, ACLs and PSModulePath are not modified. A policy/access failure is not treated as module corruption.
- Verify restored hashes and isolated import. Import the checked manifest in the caller. Failed repair is fatal for required workbook inputs; only owned, unchanged restored files are rolled back. Backups, receipts and staging are retained under `%LOCALAPPDATA%/SmartM365/ModuleRepairs` for diagnosis.
- An entirely absent module is installed only when the caller explicitly uses `-InstallIfMissing`, preserving the existing CurrentUser/AllUsers installation policy. Repair never upgrades an installed version.
- AllUsers repairs need existing write permission. No automatic elevation. ProxyAddresses still requires the AllUsers **WindowsPowerShell** module folder.

## Covered entry points

Active Directory parent inventory (classification and diagnostic workbook), EXO Quarantine, Exchange on-prem ProxyAddresses, Intune RBAC, M365 Licensing, Teams, SharePoint Online, WorkplaceScope and WorkplaceEvidence preparation. The classification reader and standalone workforce evidence worker also perform a defensive check.

`-CheckOnly` verifies/loads a healthy installation but never repairs, installs or promotes it. The classification reader uses this defensive mode: any repair belongs to its parent startup check, before downloading inputs. Workbook regression tests also use check-only mode for their installed Excel dependency.

Checks run before AD/Graph/Exchange collection, classification downloads or prepared source capture. AD domain workers do not repeat workbook repair. Optional RBAC and licensing exports remain optional and warn at startup when unavailable. RBAC DryRun does not install/repair Excel. Preparation TransferOnly, metadata conversion, legacy-history repair and source-only ValidateOnly do not require Excel. Offline preparation and standalone workforce workers use `-SkipRepositoryRepair` and cannot download packages.

## Offline qualification

Run `SmartM365/Tests/Test-SmartM365-ImportExcelHealthOffline.ps1` in isolated PowerShell 7 and Windows PowerShell 5.1 processes. It uses synthetic ImportExcel folders and mocked package-management commands, never repairs an installed module, never runs collectors or preparation, and never contacts Graph, SharePoint or Power BI. Fixtures and backup receipts are retained for inspection.

Synthetic validation does not qualify the real server installation. Deployment and a real scheduled run remain separate approval gates. ImportExcel itself must be permitted by the server execution policy; this repair mechanism does not bypass AllSigned for unsigned third-party modules.
