# SmartM365 Windows 11 Upgrade Toolkit - Intune Win32 packages

This folder builds one Intune Win32 app per Windows setup language. The Windows setup media is packaged inside the `.intunewin` so endpoints download it from Intune/Microsoft CDN instead of SMB or a custom HTTPS host.

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

The package is large and the installer copies the packaged setup media from the Intune cache to `C:\ProgramData` before starting the scheduled task.

Detection rule:

Use the generated detection script from:

```text
IntuneWin32\Output\<PackageId>\Detect.ps1
```

Recommended assignment:

Assign each language package only to devices with the matching Windows language, or to a dynamic group you control per language.

## Operational notes

- This mode does not use `SetupSource$` or `SetupSourceGates$`.
- Intune downloads the `.intunewin`; the installer copies the packaged setup media to the toolkit local cache.
- The scheduled task performs the upgrade asynchronously so Intune app installation can complete quickly.
- Upgrade status remains in `LastRun.json`, endpoint CSV output, and logs under `C:\ProgramData\SmartM365\Windows11UpgradeToolkit`.
- The package can be detected as installed even before Windows 11 is complete; use remediation/reporting to monitor final upgrade status.


## Publish to Intune with Graph

The publishing helper creates the Win32 app and uploads the generated `.intunewin` through Microsoft Graph beta. It uses the registry value written by `Install.ps1` as the Intune detection rule:

```text
HKLM\SOFTWARE\SmartM365\Windows11UpgradeToolkit\IntunePackages\<PackageId>\PackageVersion
```

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
