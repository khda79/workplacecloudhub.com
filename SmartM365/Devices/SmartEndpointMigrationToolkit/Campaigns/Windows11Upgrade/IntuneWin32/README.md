# Smart Endpoint Migration Toolkit - Windows 11 Upgrade Intune Win32 packages

This folder builds Intune Win32 apps per Windows setup language for the Windows11Upgrade
campaign. The default package embeds Windows setup media inside the `.intunewin` so endpoints
download it from Intune/Microsoft CDN instead of SMB or a custom HTTPS host. A smaller
`WithCacheOnly` variant can also be built for devices that already have the setup media cache
locally.

These packages are a new product: they use their own data root, registry root, scheduled task,
package ids and app names, and never collide with the previous Windows11UpgradeToolkit Intune
package.

## Distribution and version scope

The product contains these builders and publisher scripts as source. It contains neither Windows
media nor ready-made `.intunewin` files. Build endpoint packages locally from a reviewed toolkit
version and licensed, extracted media. `SmartM365-EndpointMigration-Install.ps1` belongs to the
generated endpoint package and writes files, registry detection state and a SYSTEM scheduled task.

Unless `-PackageVersion` is supplied, the builder derives the package version from
`$script:ScriptVersion` in
`Campaigns\Windows11Upgrade\Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1`.
Keep the generated `SmartM365-EndpointMigration-Detect.ps1` with its matching `.intunewin` and
prefer inferred metadata when publishing.

## Identities

| Item | Value |
| --- | --- |
| Data root (data only) | `C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade` |
| Program root (scripts run as SYSTEM) | `%ProgramFiles%\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade`: endpoint script; `Intune\` runner, media helper and `PackageManifest.json.txt` |
| Setup media cache | `<DataRoot>\SetupMedia\<MediaId>-<Language>` (for example `SetupMedia\Win11-fr-FR`) |
| Registry detection root (64-bit view) | `HKLM\SOFTWARE\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade\IntunePackages\<PackageId>` |
| Scheduled task | `SmartM365 SEMT Windows 11 Upgrade - Intune` |
| Package id | `SmartM365-SEMT-Windows11Upgrade-<MediaId>-<Language>[-WithCacheOnly]` |
| Intune app name | `Smart Endpoint Migration Toolkit - Windows 11 Upgrade (<Language>)` |
| Cache-only app name | `Smart Endpoint Migration Toolkit - Windows 11 Upgrade (<Language>, WithCacheOnly)` |
| Install / runner logs | `<DataRoot>\Logs\Intune\SEMT-Intune-Install.log`, `SEMT-Intune-Runner.log` |
| Endpoint attempt result | `<DataRoot>\State\Results\<CorrelationId>.json.txt` |

### Program Files and ProgramData

Every script that the SYSTEM scheduled task runs (runner, helper, package manifest, endpoint) is
installed under `%ProgramFiles%`, which only administrators can write; ProgramData keeps only
data (logs, state, results, setup media). The installer creates the folder chain with owner
Administrators and a protected DACL (SYSTEM and Administrators full control, Users read and
execute), or verifies an existing chain: an existing folder that is not owned by SYSTEM,
Administrators or TrustedInstaller, that another principal can modify, or whose permissions cannot
be read is never repaired automatically and the install fails with `SEMT_PROGRAM_FOLDER_UNTRUSTED`
(exit `1`) before any media copy or task registration. Each script is copied through a staging
name and verified against the package SHA256. The runner loads nothing and starts nothing when its
own Program Files folder is not trusted (same status, exit `1`). Script copies left under
`<DataRoot>\Intune` by earlier qualification builds are removed and never run. `-Uninstall`
removes the installed program files.

The data folder chain below ProgramData is protected as well, because the setup media cache that
it holds is run as SYSTEM and its logs and state are read by SYSTEM. Before anything is written
there:

- Scope: only the SEMT components (`SmartEndpointMigrationToolkit` and the campaign folder below
  it) are changed; they get owner Administrators and a protected DACL (SYSTEM and Administrators
  full control, Users read and execute). The shared parents `C:\ProgramData\WCH` and
  `...\WCH\SmartM365`, used by other products and by the legacy caches reused in place, are never
  changed: they are verified read-only (owner SYSTEM, Administrators or TrustedInstaller; no other
  principal may delete, delete children, change permissions or take ownership; creating items is
  allowed) and the operation stops when they are not protected.
- Read-only pass before the first change: shared parents, SEMT components, the provenance of the
  existing content and, when a DACL replacement is planned, every reparse point and hard link in
  everything it propagates to (other SEMT campaigns included). Any finding stops the operation
  with nothing changed.
- Provenance: a permission replacement never makes existing content trusted. An item that another
  principal could modify, or any content of a SEMT folder that was not protected, is refused and
  left unchanged for an administrator to review; it is never reset or adopted. A folder created
  here, or an unprotected one that was empty, must still be empty once protected.

Any failure is blocking: `SEMT_DATA_FOLDER_UNTRUSTED` for the installer (exit `1`, nothing written,
installed or registered) and `FINAL_STATUS=DATA_FOLDER_UNTRUSTED` for the endpoints (exit `1`,
nothing written under the data root). The Intune runner verifies the program folder and the data
chain (setup media included) before it writes, loads or starts anything; if the data chain is not
trusted it marks the package `RepairRequired` in the registry only, removes its task and stops.

Setup media is refused when its SHA256 manifest is missing or invalid (the endpoint checks the
result of the manifest validation), and, on Intune, when the reference manifest of the package is
required but missing: the comparison is never skipped. The manifest that comes with the media is
not, on its own, a proof of authenticity.

The registry value names are unchanged: `PackageId`, `PackageVersion`, `Language`, `MediaId`,
`SetupCacheFolder`, `PackageMode`, `InstallState`, `InstalledUtc`, `RepairReason`,
`RepairRequiredUtc`, `CompletionReason`.

## Package model

Create one package per language:

```text
SmartM365-SEMT-Windows11Upgrade-Win11-fr-FR.intunewin
SmartM365-SEMT-Windows11Upgrade-Win11-de-DE.intunewin
SmartM365-SEMT-Windows11Upgrade-Win11-pl-PL.intunewin
```

Each generated package contains:

```text
SmartM365-EndpointMigration-Install.ps1
SmartM365-EndpointMigration-Run-IntuneUpgrade.ps1
SmartM365-EndpointMigration-SetupMediaIntegrity.ps1
PackageManifest.json.txt
SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1
SetupMedia\Win11-<language>\...
```

The package installs files under the data root, then registers and starts a SYSTEM scheduled
task. The scheduled task runs the endpoint script from local cache immediately, then every two
hours for a 30-day repetition window. It removes itself when Windows 11 is detected, or when
invalid media requires package repair; it is not an unlimited retry service.

Each run passes `-DataRoot`, `-SetupCacheRoot`, `-ReuseLegacySetupMedia`, the selected media folder
`-SetupMediaPath` and a new `-CorrelationId` (GUID) to the endpoint, together with the existing direct-setup, reboot and
cleanup switches (`-DirectSetupUpgrade`, `-AllowReboot`, `-AllowSetupCompletionRebootWhenNoUser`,
`-AllowSetupProfileRepair`, `-AllowDiskCleanup`, `-ForceRequiredRebootWhenUptimeOverDays 7`). No
other reboot is added. After the endpoint exits, the runner logs the attempt outcome from
`State\Results\<CorrelationId>.json.txt` when that attempt is `Completed`, and falls back to
`LastRun.json.txt` otherwise.

## Setup media selection

The installer, the runner and the endpoint use the same read-only resolution
(`SEMT setup media resolution v1`, identical in `SmartM365-EndpointMigration-SetupMediaIntegrity.ps1`
and in the endpoint script) before any copy. For the package cache folder (for example
`Win11-fr-FR`) it looks at the cache of this product
(`C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade\SetupMedia\...`) and at
the cache of the previous toolkit (`C:\ProgramData\SmartM365\Windows11UpgradeToolkit\SetupMedia\...`,
plus the path recorded in its `LastRun.json.txt`):

| Decision | Meaning | Installer | Runner / endpoint |
| --- | --- | --- | --- |
| `NewCacheValid` | Valid media in the cache of this product | no copy | used |
| `LegacyIntuneOwned` | Valid legacy media still owned by the previous Intune package | no copy; recorded as external | used in place, read-only (never moved, cleaned, re-permissioned or marked) |
| `LegacyAdoptable` | Valid, idle, unmanaged legacy media | no copy | the endpoint moves it on the same volume with full rollback (see below) |
| `InUse` | A live process or lock uses the media | exit `1618` (Intune retries), nothing changed | cycle skipped, nothing changed |
| `Indeterminate` | Registry, processes, lock, permissions or folder tree of legacy media unreadable | exit `1`, nothing changed (even with `-AllowSetupMediaDuplication`) | cycle skipped, nothing changed |
| `Untrusted` | Legacy media present but owned or modifiable by a non-administrator, or containing a reparse point | never used, permissions not changed; exit `1` `SEMT_SETUP_MEDIA_LEGACY_UNTRUSTED`, unless the WithMedia package was built with `-AllowSetupMediaDuplication` (packaged copy, logged `SEMT_SETUP_MEDIA_DUPLICATION_ALLOWED`) | endpoint not started; package marked `RepairRequired` (`SEMT_MEDIA_UNTRUSTED`) |
| `Incompatible` | Other language or other media build than the package | `WithMedia`: packaged media copied; `WithCacheOnly`: exit `1` | package marked `RepairRequired` |
| `Corrupt` / `Absent` | Invalid media / no media | same as `Incompatible` | package marked `RepairRequired` |

Legacy media is considered only when it is write-restricted: every item owned by SYSTEM,
Administrators or TrustedInstaller and modifiable by no other principal, with no reparse point
(otherwise `Untrusted`: never used, because `setup.exe` runs as SYSTEM, and its permissions are never
changed automatically). Permissions or a folder tree that cannot be read give `Indeterminate`, which
blocks; present legacy media is never reported as `Absent`. A path recorded in the
legacy `LastRun.json` is used only when it normalizes to a folder under the legacy `SetupMedia`
folder. Without a trusted reference manifest (WithCacheOnly), the media manifest itself is read
from the protected folder.

Validation is read-only: `setup.exe`, `sources\install.wim` (or `.esd`), the language in
`sources\lang.ini`, the full SHA256 manifest, and for `WithMedia` the same manifest as the packaged
media. The selected folder is recorded in the package registry state (`SetupMediaPath`,
`SetupMediaSource` = `NewCache`, `LegacyExternal`, `LegacyAdoptable` or `PackageCopy`,
`SetupMediaExternal`) and passed to the endpoint; the endpoint re-checks it and rejects a folder that
is neither the new cache nor a legacy cache candidate (`SETUP_MEDIA_PATH_REJECTED`).

Adoption of legacy media by the endpoint checks paths, the absence of the destination and of reparse
points, ownership, use and validity before any change; it records the content and the DACL of every
item and saves them with `icacls /save`; then it moves the folder on the same volume, resets the
permissions to the inherited ones and validates the media. On any failure it moves the folder back,
re-applies the recorded permissions and verifies location, sizes and permissions. The outcome is
`Moved`, `NotMoved`, `RolledBack`, `RollbackPartial` (moved back, permissions not fully restored) or
`RollbackFailed` (for example when the legacy folder was recreated meanwhile: it is never overwritten
and the media stays at the new cache). Evidence: `<DataRoot>\State\MediaAdoption\<timestamp>\`.

No second copy is made silently: when legacy media remains but cannot be used (valid but, for example,
the new cache folder already exists; or `Untrusted`), the installer fails with `SEMT_SETUP_MEDIA_DUPLICATION_BLOCKED` and
the endpoint with `SETUP_MEDIA_DUPLICATION_BLOCKED`, unless the package was built with
`-AllowSetupMediaDuplication` (logged). For `Untrusted` media the statuses are
`SEMT_SETUP_MEDIA_LEGACY_UNTRUSTED` (installer) and `SETUP_MEDIA_LEGACY_UNTRUSTED` (endpoint).

### WithMedia or WithCacheOnly

- Use **WithCacheOnly** for devices that already have reusable media: a valid cache of this product,
  or legacy media (owned by the previous Intune package, or idle and adoptable). Nothing large is
  downloaded or copied. If no reusable media exists, the installer fails with
  `SEMT_SETUP_MEDIA_NOT_REUSABLE` and nothing is installed: assign the WithMedia package instead.
- Use **WithMedia** for devices without reusable media. Intune downloads the whole package content
  (media included) to its own cache **before** the installer runs: reusing a local cache afterwards
  avoids the local copy to `C:\ProgramData`, not that download.
- A new copy (WithMedia only) happens only when the resolution is `Absent`, `Corrupt` or
  `Incompatible`, after the packaged media passed its SHA256 check, and never next to valid legacy
  media unless `-AllowSetupMediaDuplication` was used at build time.

## Prepare media

Put each language media under the campaign `SetupSource` folder, for example:

```text
Campaigns\Windows11Upgrade\SetupSource\FR-fr\setup.exe
Campaigns\Windows11Upgrade\SetupSource\FR-fr\sources\install.wim
Campaigns\Windows11Upgrade\SetupSource\FR-fr\SmartM365-SetupMediaManifest.sha256.csv
```

Generate manifests before building, from the product root:

```powershell
.\New-SetupMediaManifest.cmd
```

or directly:

```powershell
.\Campaigns\Windows11Upgrade\Scripts\SmartM365-EndpointMigration-New-SetupMediaManifest.ps1 `
  -SetupSourceRoot .\Campaigns\Windows11Upgrade\SetupSource -Force
```

## Build one language package

Download the [Microsoft Win32 Content Prep Tool](https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool) (`IntuneWinAppUtil.exe`) separately and run, from `Campaigns\Windows11Upgrade`:

```powershell
.\IntuneWin32\SmartM365-EndpointMigration-Build-Windows11IntunePackage.ps1 `
  -Language fr-FR `
  -IntuneWinAppUtilPath C:\Tools\IntuneWinAppUtil.exe
```

The default media folder is inferred from language (`fr-FR` -> `FR-fr`, `de-DE` -> `DE-de`).
Override with `-MediaFolder` when needed. The builder passes
`SmartM365-EndpointMigration-Install.ps1` as the IntuneWinAppUtil setup file and copies the result
to `IntuneWin32\Output\<PackageId>\<PackageId>.intunewin`. `-SkipIntuneWinBuild` stages the
package and its detection script without calling IntuneWinAppUtil.

## Build every SetupSource language package

```powershell
.\IntuneWin32\SmartM365-EndpointMigration-Build-AllWindows11IntunePackages.ps1 `
  -GenerateMissingManifest `
  -IntuneWinAppUtilPath C:\Tools\IntuneWinAppUtil.exe `
  -Force
```

The script detects folders named like `FR-fr`, `DE-de`, `ES-es`, converts them to package
languages like `fr-FR`, `de-DE`, `es-ES`, refreshes missing media manifests when requested, and
calls the single-language builder for each folder. Intermediate staging folders are removed after
each successful package by default; use `-KeepStaging` to inspect the staged package source. Use
`-IncludeMediaFolder FR-fr,DE-de` to limit the build.

## Build a cache-only package

Use `-WithCacheOnly` to create a lightweight package without Windows setup media. The package id
is `SmartM365-SEMT-Windows11Upgrade-Win11-fr-FR-WithCacheOnly`.

This package needs reusable media on the endpoint, in the cache of this product or in the cache of
the previous toolkit (see "Setup media selection"):

```text
C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade\SetupMedia\Win11-fr-FR\
C:\ProgramData\SmartM365\Windows11UpgradeToolkit\SetupMedia\Win11-fr-FR\
```

The publisher adds an Intune requirement rule that returns applicable only when the Windows setup
language matches and, for `WithCacheOnly`, `setup.exe` and `sources\install.wim` exist in one of the
two folders above (the installer then performs the full validation). The rule resolves the registry `InstallLanguage`
first and falls back to `InstalledUICulture`; regional `SystemLocale` values such as `fr-CH` do
not exclude a device whose Windows installation language is `fr-FR`.

## Intune app settings

Program install command:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\SmartM365-EndpointMigration-Install.ps1
```

Program uninstall command:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\SmartM365-EndpointMigration-Install.ps1 -Uninstall
```

Recommended install timeout: 180 minutes or more. When no reusable media exists, the full package
installer copies the packaged setup media from the Intune cache to `C:\ProgramData` before starting
the scheduled task; with reusable media it copies nothing.

Disk space requirement: the publisher sets Intune `Disk space required (MB)` to `40960` for full
`WithMedia` packages and leaves it unset for `WithCacheOnly` packages unless you override it with
`-MinimumFreeDiskSpaceInMB <MB>` (`0` disables it). Endpoint disk preflight checks remain
authoritative during upgrade execution.

Detection rule: use the generated
`IntuneWin32\Output\<PackageId>\SmartM365-EndpointMigration-Detect.ps1`. Detection succeeds when
the device is already Windows 11, or when the package registry state is present in the 64-bit
registry view with the expected package id, an installed state when that value exists
(`Installed` or `AlreadyWindows11`), and the expected package version or a newer one. Detection
never looks at where the media lives: an installation that uses legacy media in place
(`SetupMediaSource=LegacyExternal`) is detected as installed.

## Switching devices from the previous Windows11UpgradeToolkit app

Assignments stay manual. For the same devices, follow this order:

1. In Intune, remove the assignment of the old `Windows11UpgradeToolkit-<language>` Win32 app
   (do not assign it as "Uninstall" while an upgrade may be running).
2. Assign the new `Smart Endpoint Migration Toolkit - Windows 11 Upgrade (<language>)` app to the
   same devices.

Both apps can coexist on a device without collision because every identity differs. The old
package state is only read by the new endpoint:

- Setup media still owned by the old Intune package (its registry state under the old toolkit is
  installed or repair-required) is reused in place, read-only, by the installer, the runner and the
  endpoint. It is not moved, copied, cleaned, re-permissioned or marked, and it stays external to
  this product. Idle media no longer owned by the old package is adopted by the endpoint (same
  volume, full rollback).
- After step 1, while the old package is still installed on a device, its media keeps the
  `LegacyIntuneOwned` decision. When the old package is uninstalled and its media removed, the next
  runner cycle finds no usable media: a WithMedia package is repaired by Intune (packaged media
  copied), a WithCacheOnly package fails explicitly.
- Prefer the WithCacheOnly package for devices that already have legacy media, to avoid the
  download of the media by Intune.
- A real active legacy operation (running setup, running legacy task, legacy media copy) blocks the
  endpoint under the default `-LegacyActivityPolicy Block`; an installed but idle legacy package
  does not.
- The new app never removes the old scheduled task, registry key or media.

## Updates and uninstall scope

To distribute changed endpoint scripts or media, rebuild and validate the package, then
explicitly publish its content. `-UpdateMetadataOnly` does not upload new content; detection rules
remain unchanged unless `-UpdateDetectionRules` is added. Use the version actually deployed when
changing detection without replacing content.

`SmartM365-EndpointMigration-Install.ps1 -Uninstall` unregisters the named scheduled task and
removes this package's registry detection key. It leaves the cached toolkit files, Windows setup
media and logs on disk; it does not roll back Windows or guarantee that an already running upgrade
has stopped.

## Operational notes

- This mode does not use setup source shares or `SetupSourceGates`.
- The builder validates every file, size, and SHA256 recorded in
  `SmartM365-SetupMediaManifest.sha256.csv` against the source media, the staged package, and the
  final prep folder. A partial source, including a missing `boot\bcd`, cannot produce a package.
- The full-package installer validates the packaged media before copying it and validates the
  complete destination cache before writing `InstallState=Installed`. The `WithCacheOnly`
  installer validates the selected reusable media and fails before detection when none is usable.
- The installer and the endpoint coordinate setup cache writes with an expiring local lock under
  `<DataRoot>\Locks`. If another process is copying or clearing the same cache, the installer
  exits with `1618` so Intune retries instead of deleting a partial cache.
- The scheduled runner resolves the media before every Windows 10 endpoint launch. Only when no
  usable media remains (`Absent`, `Corrupt`, `Incompatible`) does it write
  `InstallState=RepairRequired` (`RepairReason=SEMT_MEDIA_<DECISION>: ...`) and remove its scheduled
  task; the detection rule then fails so Intune reinstalls the package. Valid external media never
  triggers a repair; media in use or in an unreadable state only skips the cycle.
- Windows 11 has priority over cache and repair state: the installer, runner, and detection rule
  report the application as `Installed` and remove the obsolete scheduled task as soon as Windows
  11 is detected.
- Upgrade status remains in `State\Results\<CorrelationId>.json.txt`, `LastRun.json.txt`, endpoint CSV
  output, and logs under the data root.
- On every scheduled run, `Logs\Intune\Endpoint_*.log` files older than 7 days are removed and at
  most 200 of them are kept; persistent logs (`SEMT-Intune-Runner.log`, `SEMT-Intune-Install.log`,
  `SEMT-Intune-Install-Robocopy.log`) are preserved. Cleanup failures never block the upgrade.

## Publish to Intune with Graph

The publishing helper creates the Win32 app and uploads the generated `.intunewin` through
Microsoft Graph beta, using the generated detection script. It refuses a package id that does not
start with `SmartM365-SEMT-Windows11Upgrade-`, and the batch publisher skips such packages.

Prerequisite on the publishing workstation:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

Preview without creating anything (`PackageId`, `PackageVersion` and `Language` are inferred from
the files next to the `.intunewin`):

```powershell
.\IntuneWin32\SmartM365-EndpointMigration-Publish-Windows11IntuneApp.ps1 `
  -IntuneWinPath .\IntuneWin32\Output\SmartM365-SEMT-Windows11Upgrade-Win11-fr-FR\SmartM365-SEMT-Windows11Upgrade-Win11-fr-FR.intunewin `
  -WhatIf
```

Publish every generated full-media package except FR:

```powershell
.\IntuneWin32\SmartM365-EndpointMigration-Publish-AllWindows11IntuneApps.ps1 -ExcludeLanguage fr-FR
```

Update only existing app metadata, without a new content version:

```powershell
.\IntuneWin32\SmartM365-EndpointMigration-Publish-AllWindows11IntuneApps.ps1 `
  -UpdateMetadataOnly `
  -MinimumFreeDiskSpaceInMB 40960
```

The batch publisher defaults to `-PackageMode WithMedia`; use `-PackageMode All` or
`-PackageMode WithCacheOnly` when needed. If upload and commit succeed but the final content-version
PATCH fails, use `-FinalizeExistingAppId <IntuneAppId> -FinalizeContentVersionId 1`.

Required Graph delegated permission:

```text
DeviceManagementApps.ReadWrite.All
```

No-Graph-assignment policy: the scripts create apps and upload content only. They never create,
change or remove Intune assignments, so a package cannot be pushed to the wrong language group or
overlap the old app by accident. Assignments, including the switch order above, stay a manual
operator decision.
