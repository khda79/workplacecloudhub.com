# SmartM365 Windows 11 Upgrade Toolkit - Intune Win32 packages

This folder builds Intune Win32 apps per Windows setup language. The default package embeds Windows setup media inside the `.intunewin` so endpoints download it from Intune/Microsoft CDN instead of SMB or a custom HTTPS host. A smaller `WithCacheOnly` variant can also be built for devices that already have the setup media cache locally.

## Package model

Create one package per language:

```text
SmartM365-Windows11UpgradeToolkit-Win11-fr-FR.intunewin
SmartM365-Windows11UpgradeToolkit-Win11-de-DE.intunewin
SmartM365-Windows11UpgradeToolkit-Win11-pl-PL.intunewin
```

Each generated package contains:

```text
Install.ps1
Run-IntuneUpgrade.ps1
PackageManifest.json
SmartM365-Invoke-Windows11UpgradeRepair.ps1
SetupMedia\Win11-<language>\...
```

The package installs files under:

```text
C:\ProgramData\SmartM365\Windows11UpgradeToolkit
```

Then it registers and starts a SYSTEM scheduled task. The scheduled task runs the endpoint script from local cache and retries periodically until the device is already Windows 11.

## Prepare media

Put each language media under the toolkit `SetupSource` folder, for example:

```text
SetupSource\FR-fr\setup.exe
SetupSource\FR-fr\sources\install.wim
SetupSource\FR-fr\SmartM365-SetupMediaManifest.sha256.csv
```

Generate manifests before building:

```powershell
.\Scripts\New-SmartM365SetupMediaManifest.ps1 -SetupSourceRoot .\SetupSource -Force
```

## Build one language package

Download Microsoft Win32 Content Prep Tool (`IntuneWinAppUtil.exe`) and run:

```powershell
.\IntuneWin32\Build-SmartM365Windows11IntunePackage.ps1 `
  -Language fr-FR `
  -IntuneWinAppUtilPath C:\Tools\IntuneWinAppUtil.exe
```

For German media:

```powershell
.\IntuneWin32\Build-SmartM365Windows11IntunePackage.ps1 `
  -Language de-DE `
  -IntuneWinAppUtilPath C:\Tools\IntuneWinAppUtil.exe
```

The default media folder is inferred from language (`fr-FR` -> `FR-fr`, `de-DE` -> `DE-de`). Override with `-MediaFolder` when needed.

## Build every SetupSource language package

To build one full media `.intunewin` package for every valid media folder under `SetupSource`:

```powershell
.\IntuneWin32\Build-AllSmartM365Windows11IntunePackages.ps1 `
  -GenerateMissingManifest `
  -IntuneWinAppUtilPath C:\Tools\IntuneWinAppUtil.exe `
  -Force
```

The script detects folders named like `FR-fr`, `DE-de`, `ES-es`, converts them to Intune package languages like `fr-FR`, `de-DE`, `es-ES`, refreshes missing media manifests when requested, and calls the single-language builder for each folder. Intermediate staging folders are removed after each successful package by default; use `-KeepStaging` only when you need to inspect the staged package source.

To limit the build to selected media folders:

```powershell
.\IntuneWin32\Build-AllSmartM365Windows11IntunePackages.ps1 `
  -IncludeMediaFolder FR-fr,DE-de `
  -GenerateMissingManifest `
  -IntuneWinAppUtilPath C:\Tools\IntuneWinAppUtil.exe `
  -Force
```

## Build a cache-only package

Use `-WithCacheOnly` to create a lightweight package without Windows setup media. The generated app name is `Windows11UpgradeToolkit-fr-FR-WithCacheOnly` and the package id is `SmartM365-Windows11UpgradeToolkit-Win11-fr-FR-WithCacheOnly`.

```powershell
.\IntuneWin32\Build-SmartM365Windows11IntunePackage.ps1 `
  -Language fr-FR `
  -WithCacheOnly `
  -IntuneWinAppUtilPath C:\Tools\IntuneWinAppUtil.exe
```

This package requires the cache to already exist on the endpoint:

```text
C:\ProgramData\SmartM365\Windows11UpgradeToolkit\SetupMedia\Win11-fr-FR\setup.exe
C:\ProgramData\SmartM365\Windows11UpgradeToolkit\SetupMedia\Win11-fr-FR\sources\install.wim
```

The publisher adds an Intune requirement rule that returns applicable only when the Windows setup language matches and, for `WithCacheOnly`, the local cache exists.

## Intune app settings

Program install command:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1
```

Program uninstall command:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install.ps1 -Uninstall
```

Recommended install timeout:

```text
180 minutes or more
```

The full package is large and the installer copies the packaged setup media from the Intune cache to `C:\ProgramData` before starting the scheduled task. The `WithCacheOnly` package is small and only validates an existing local setup cache.

Disk space requirement:

The publisher sets Intune `Disk space required (MB)` automatically to `51200` for full `WithMedia` packages. It leaves the value unset for `WithCacheOnly` packages unless you override it. Use `-MinimumFreeDiskSpaceInMB <MB>` to force a value, or `-MinimumFreeDiskSpaceInMB 0` to disable the requirement. This Intune requirement filters installation before download/install, but endpoint disk preflight checks still remain authoritative during upgrade execution.

Detection rule:

Use the generated detection script from:

```text
IntuneWin32\Output\<PackageId>\Detect.ps1
```

Recommended assignment:

Assign each language package only to devices with the matching Windows language, or to a dynamic group you control per language.

## Operational notes

- This mode does not use `SetupSource$` or `SetupSourceGates$`.
- Intune downloads the `.intunewin`; the full-package installer copies the packaged setup media to the toolkit local cache. The `WithCacheOnly` installer does not copy media and fails before detection if the local cache is missing.
- The Intune installer and the endpoint upgrade script coordinate setup cache writes with an expiring local lock under `%ProgramData%\SmartM365\Windows11UpgradeToolkit\Locks`. If another SmartM365 process is copying or clearing the same cache, the Intune installer exits with `1618` so Intune retries instead of deleting a partial cache.
- The scheduled task performs the upgrade asynchronously so Intune app installation can complete quickly.
- Upgrade status remains in `LastRun.json`, endpoint CSV output, and logs under `C:\ProgramData\SmartM365\Windows11UpgradeToolkit`.
- The package can be detected as installed even before Windows 11 is complete; use remediation/reporting to monitor final upgrade status.


## Publish to Intune with Graph

The publishing helper creates the Win32 app and uploads the generated `.intunewin` through Microsoft Graph beta. It uses the generated `Detect.ps1` as the Intune detection rule. Detection succeeds when the device is already Windows 11, or when the package registry state written by `Install.ps1` is present in the 64-bit registry view, has the expected package id, has an installed state when that value exists, and reports the expected package version or a newer one:

```text
HKLM\SOFTWARE\SmartM365\Windows11UpgradeToolkit\IntunePackages\<PackageId>\PackageId
HKLM\SOFTWARE\SmartM365\Windows11UpgradeToolkit\IntunePackages\<PackageId>\PackageVersion
HKLM\SOFTWARE\SmartM365\Windows11UpgradeToolkit\IntunePackages\<PackageId>\InstallState
```

`Install.ps1` writes this state explicitly through the 64-bit registry view. The generated `Detect.ps1` reads the same 64-bit view, writes an `OK: ...` detection result to stdout, and also treats Windows 11 as already compliant.

Prerequisite on the publishing workstation:

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

Preview the FR app payload without creating anything. `PackageId`, `PackageVersion`, and `Language` are inferred when the generated `Detect.ps1` is next to the `.intunewin` file:

```powershell
.\IntuneWin32\Publish-SmartM365Windows11IntuneApp.ps1 `
  -IntuneWinPath C:\tmp\SmartM365-W11UT-FR\Output\SmartM365-Windows11UpgradeToolkit-Win11-fr-FR.intunewin `
  -WhatIf
```

Explicit preview:

```powershell
.\IntuneWin32\Publish-SmartM365Windows11IntuneApp.ps1 `
  -IntuneWinPath C:\tmp\SmartM365-W11UT-FR\Output\SmartM365-Windows11UpgradeToolkit-Win11-fr-FR.intunewin `
  -PackageId SmartM365-Windows11UpgradeToolkit-Win11-fr-FR `
  -PackageVersion 0.1.30 `
  -Language fr-FR `
  -WhatIf
```

Publish the FR app:

```powershell
.\IntuneWin32\Publish-SmartM365Windows11IntuneApp.ps1 `
  -IntuneWinPath C:\tmp\SmartM365-W11UT-FR\Output\SmartM365-Windows11UpgradeToolkit-Win11-fr-FR.intunewin `
  -PackageId SmartM365-Windows11UpgradeToolkit-Win11-fr-FR `
  -PackageVersion 0.1.30 `
  -Language fr-FR
```

Publish every generated full-media package except FR:

```powershell
.\IntuneWin32\Publish-AllSmartM365Windows11IntuneApp.ps1 `
  -ExcludeLanguage fr-FR
```

Preview the same batch without creating or updating apps:

```powershell
.\IntuneWin32\Publish-AllSmartM365Windows11IntuneApp.ps1 `
  -ExcludeLanguage fr-FR `
  -WhatIf
```

The batch publisher defaults to `-PackageMode WithMedia` so existing `WithCacheOnly` packages under `Output` are not republished accidentally. Use `-PackageMode All` or `-PackageMode WithCacheOnly` when needed. It calls `Publish-SmartM365Windows11IntuneApp.ps1` once per selected `.intunewin` package and still creates/uploads apps only; assignments remain manual.

Required Graph delegated permission:

```text
DeviceManagementApps.ReadWrite.All
```

The script creates the app and uploads content only. Assignments remain manual so the package is not accidentally pushed as required to the wrong language group.
If upload and commit succeed but the final app content-version PATCH fails, finalize the existing app without re-uploading the package:

```powershell
.\IntuneWin32\Publish-SmartM365Windows11IntuneApp.ps1 `
  -IntuneWinPath C:\tmp\SmartM365-W11UT-FR\Output\SmartM365-Windows11UpgradeToolkit-Win11-fr-FR.intunewin `
  -FinalizeExistingAppId 6bd2540a-be73-44a1-8253-b79638f3f976 `
  -FinalizeContentVersionId 1
```
