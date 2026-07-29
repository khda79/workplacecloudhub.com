# Smart Device Reboot Manager

Local WPF user notification app for SmartM365 device restart governance.

## Download

Download the stable Windows package from the
[SmartM365 Device Reboot Manager v0.1.0 release](https://github.com/khda79/workplacecloudhub.com/releases/tag/device-reboot-manager-v0.1.0).

The GitHub release provides:

- `SmartM365-DeviceRebootManager-0.1.0.zip`: complete standalone source package.
- `SmartM365-DeviceRebootManager-0.1.0.intunewin`: clean Intune Win32 package.
- `SmartM365-DeviceRebootManager-0.1.0.sha256`: SHA-256 checksums for both packages.

The PowerShell module is also available from the
[PowerShell Gallery](https://www.powershellgallery.com/packages/SmartM365.DeviceRebootManager/0.1.0).

## Scripts

- `SmartM365-DeviceRebootManager-GUI.ps1`: improved SmartM365 GUI with status summary, safer state file, configurable postpone choices, preview mode, and cleaner logging.
- `Start-SmartM365-DeviceRebootManager-GUI.cmd`: launcher that starts the GUI in a PowerShell STA runspace with the console hidden.
- `Start-SmartM365-DeviceRebootManager-GUI-Test.cmd`: safe test launcher that simulates required mode with `PreviewOnly`.
- `SmartM365-DeviceRebootManager-GUI.config.json.template`: safe configuration model.
- `SmartM365-DeviceRebootManager-GUI.strings.psd1`: UI localization catalog.
- `Deploy/SmartM365-DeviceRebootManager-Install.ps1`: Intune Win32 install script.
- `Deploy/SmartM365-DeviceRebootManager-Uninstall.ps1`: Intune Win32 uninstall script.
- `Deploy/SmartM365-DeviceRebootManager-Detection.ps1`: Intune Win32 detection script.
- `Deploy/SmartM365-DeviceRebootManager-CreateScheduledTask.ps1`: standalone scheduled task creation script.

## Configuration

`SmartM365-DeviceRebootManager-GUI.ps1` loads runtime configuration in this order:

1. `-ConfigPath` when provided.
2. `SmartM365-DeviceRebootManager-GUI.config.json` next to the script.
3. `%ProgramData%\SmartM365\DeviceRebootManager\SmartM365-DeviceRebootManager-GUI.config.json`.
4. `%APPDATA%\SmartM365\DeviceRebootManager\SmartM365-DeviceRebootManager-GUI.config.json`.

The committed template contains safe generic defaults only. Runtime `.json` files are local operational configuration.

`DefaultLanguage` can be `auto`, `en`, `fr`, `de`, `es`, `nl`, `it`, `pt`, `pl`, `ar`, `tr`, `sv`, `da`, `nb`, `fi`, `ro`, `hu`, `ja`, `ko`, `zh-Hans`, or `uk`. The app also maps common regional cultures, for example `zh-CN` to `zh-Hans` and `no` to `nb`. Arabic uses right-to-left layout.

`WindowTitle` overrides the localized window title when set. Leave it empty to use the title from the selected language.

When `WindowTitle` is empty, the app tries to detect the company name from the device Entra join or AD domain and prefixes the localized title, for example `Company - Smart Device Reboot Manager`. Set `CompanyName` to force a specific company prefix, or leave it empty for automatic detection.

Users can change the UI language from the language selector in the GUI. The choice is stored in `%APPDATA%\SmartM365\DeviceRebootManager\SmartM365-DeviceRebootManager.preferences.json` and is reused on later launches unless `-DefaultLanguage` is passed on the command line. Set `ForceLanguage` to `true` to ignore the saved user preference and lock the GUI to `DefaultLanguage`.

`LanguageCatalogPath` can point to a custom PowerShell data file with the same structure as `SmartM365-DeviceRebootManager-GUI.strings.psd1`. Leave it empty to use the bundled catalog.

`SplashEnabled` controls the startup WorkplaceCloudHub splash. Set it to `false` to disable the splash. Keep it enabled and customize `SplashMinimumDurationMs`, `SplashProductName`, `SplashBadgeText`, `SplashSubtitle`, and `SplashLogoPath` to adapt the splash for a deployment. The splash URL remains `https://workplacecloudhub.com`.

At each launch, the current `SmartM365-DeviceRebootManager.log` is archived to a timestamped log file and a fresh log is created. Only the 10 most recent archived logs are retained.

## Intune Win32 Deployment

Build the clean allow-listed package, then use the returned `IntuneSourcePath`
(`...\SmartM365.DeviceRebootManager\0.1.0\Runtime`) as the Intune Win32 source.
Do not package the whole working folder because it can contain private runtime
configuration. The install script copies the runtime files to:

```text
C:\ProgramData\SmartM365\DeviceRebootManager
```

It also creates this scheduled task:

```text
\SmartM365\Device Reboot Manager
```

The task runs in the interactive user session, not in session 0. It starts at user logon and repeats every 240 minutes by default. The GUI script exits silently when the restart threshold is not reached, so regular execution is expected.

Recommended Intune Win32 commands:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Deploy\SmartM365-DeviceRebootManager-Install.ps1
```

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Deploy\SmartM365-DeviceRebootManager-Uninstall.ps1
```

Use `Deploy\SmartM365-DeviceRebootManager-Detection.ps1` directly as the
Intune custom detection script. Its default expected version is
`0.1.0`; update that value for every new Intune package:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Deploy\SmartM365-DeviceRebootManager-Detection.ps1
```

The Intune installer removes the optional
`\SmartM365\Device Reboot Manager Update` Gallery task, records the exact
installed version and source, and preserves the existing runtime configuration
during an in-place upgrade. The detection rule returns success only when the
expected version, metadata, GUI files and GUI task are present and the Gallery
updater is absent.

Publish each new version as a new Intune Win32 app, configure supersedence to
replace the previous version, and explicitly assign the new app. Intune then
controls rollout, retry, reporting and rollback; no local automatic-update task
is required.

The install script creates `SmartM365-DeviceRebootManager-GUI.config.json` from the committed template when no runtime config exists. To deploy a custom config, include a JSON file in the package and pass it with `-ConfigSourcePath`.

Example:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Deploy\SmartM365-DeviceRebootManager-Install.ps1 -ConfigSourcePath .\SmartM365-DeviceRebootManager-GUI.config.json -RepeatIntervalMinutes 120
```

### Interactive Intune publication

`Deploy\SmartM365-DeviceRebootManager-PublishIntune.ps1` prepares an
administrator-controlled publication with delegated interactive Microsoft Graph
authentication. Preview is the default: it validates the package, detection
version and signer without connecting to the tenant or changing Intune.

Preview the publication and proposed pilot group:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass `
    -File .\Deploy\SmartM365-DeviceRebootManager-PublishIntune.ps1 `
    -IntuneWinPath C:\Temp\SmartM365-DeviceRebootManager-0.1.0.intunewin
```

Publish the stable app, supersede the prior preview as an in-place update, reuse
the assigned security group, and assign the stable app as `Required`:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass `
    -File .\Deploy\SmartM365-DeviceRebootManager-PublishIntune.ps1 `
    -IntuneWinPath C:\Temp\SmartM365-DeviceRebootManager-0.1.0.intunewin `
    -TenantId <tenant-id> `
    -PilotGroupId <pilot-group-id> `
    -AssignPilotGroup `
    -SupersedeAppId <preview-app-id> `
    -SupersedenceType update `
    -Execute
```

The proposed group is
`GG-INTUNE-SmartM365-DeviceRebootManager-Pilot`. Use `-PilotGroupId` to
target an existing security group instead. The required delegated scopes are
`DeviceManagementApps.ReadWrite.All` and `Group.ReadWrite.All`; tenant consent
and an administrator account with matching Intune/Entra permissions are still
required.

The script creates a version-specific Intune app, embeds the signed detection
script, uploads and commits the encrypted content, optionally creates an Intune
supersedence relationship, and creates the pilot assignment individually through
Graph. For upgrades of the same product, use `-SupersedenceType update` so the
new installer performs the in-place update without first uninstalling the old
version. The script does not enable the local Gallery updater and does not
replace other app assignments.

If publication fails after the app content is committed, rerun with
`-ResumeAppId <incomplete-app-id>`. The script validates the existing app name,
publisher, package-version notes, and committed content version before it resumes
supersedence and assignment. It does not upload the package or create another app
in recovery mode.

To recreate only the scheduled task on an already installed device:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Deploy\SmartM365-DeviceRebootManager-CreateScheduledTask.ps1 -RepeatIntervalMinutes 240
```

## PowerShell Gallery stable package

The stable package `SmartM365.DeviceRebootManager` version `0.1.0` is published
on [PowerShell Gallery](https://www.powershellgallery.com/packages/SmartM365.DeviceRebootManager/0.1.0).

Build and validate the package locally:

```powershell
.\PowerShellGallery\SmartM365-Build-DeviceRebootManagerGalleryPackage.ps1 -Force
```

Preview publication without sending anything to PowerShell Gallery:

```powershell
.\PowerShellGallery\SmartM365-Publish-DeviceRebootManagerGalleryPackage.ps1 -ForceBuild
```

Public metadata includes the repository GPL-3.0 license. Publication still
requires an explicit `-Execute` and an API key supplied through
`PSGALLERY_API_KEY`; the key must never be written in a script or committed:

```powershell
$env:PSGALLERY_API_KEY = '<temporary-api-key>'
.\PowerShellGallery\SmartM365-Publish-DeviceRebootManagerGalleryPackage.ps1 `
    -ForceBuild `
    -Execute
```

An elevated Windows PowerShell session can install the stable package and
deploy the app:

```powershell
Install-Module SmartM365.DeviceRebootManager -Repository PSGallery -Scope AllUsers -Force
Import-Module SmartM365.DeviceRebootManager
Install-SmartM365DeviceRebootManager
```

Automatic updates are disabled by default. A standard installation does not
create an update task. Enable them explicitly when required:

```powershell
Install-SmartM365DeviceRebootManager -EnableAutomaticUpdate $true
```

When enabled, the deployment registers
`\SmartM365\Device Reboot Manager Update` under `SYSTEM`. Every 24 hours, it
checks the stable PowerShell Gallery channel by default, validates every
packaged PowerShell file against the pinned WorkplaceCloudHub signer, preserves
the local runtime configuration, and redeploys the runtime.

Operational commands:

```powershell
Get-SmartM365DeviceRebootManager
Update-SmartM365DeviceRebootManager
Uninstall-SmartM365DeviceRebootManager
```

### Clean VM validation bundle

Build a transferable ZIP without installing or publishing anything:

```powershell
.\PowerShellGallery\SmartM365-New-DeviceRebootManagerGalleryVmBundle.ps1 -Force
```

After extracting the ZIP on an isolated VM, preview the validation:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\SmartM365-Test-DeviceRebootManagerGalleryVm.ps1
```

Run the installation, opt-in, opt-out, configuration-preservation, and task
checks from an elevated Windows PowerShell session:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\SmartM365-Test-DeviceRebootManagerGalleryVm.ps1 -Execute -TrustSignerCertificate
```

`-TrustSignerCertificate` is intended only for the isolated pilot VM while the
package uses the current self-signed WorkplaceCloudHub certificate. Add
`-CleanupAfterTest` to remove the product and the certificate trust entries
created by the validation.

## Examples

Preview the required mode without rebooting:

```powershell
powershell.exe -STA -NoProfile -File .\Devices\DeviceRebootManager\SmartM365-DeviceRebootManager-GUI.ps1 -TestRequiredRestart -PreviewOnly
```

Preview the recommendation mode without rebooting:

```powershell
powershell.exe -STA -NoProfile -File .\Devices\DeviceRebootManager\SmartM365-DeviceRebootManager-GUI.ps1 -TestRecommendedRestart -PreviewOnly
```

Run in recommendation-only mode, with no mandatory restart state:

```powershell
powershell.exe -STA -NoProfile -File .\Devices\DeviceRebootManager\SmartM365-DeviceRebootManager-GUI.ps1 -NeverForceRestart
```

Use a specific config file:

```powershell
powershell.exe -STA -NoProfile -File .\Devices\DeviceRebootManager\SmartM365-DeviceRebootManager-GUI.ps1 -ConfigPath C:\ProgramData\SmartM365\DeviceRebootManager\SmartM365-DeviceRebootManager-GUI.config.json
```

Launch through the CMD wrapper:

```cmd
.\Devices\DeviceRebootManager\Start-SmartM365-DeviceRebootManager-GUI.cmd -TestRequiredRestart -PreviewOnly
```

Launch the safe test wrapper:

```cmd
.\Devices\DeviceRebootManager\Start-SmartM365-DeviceRebootManager-GUI-Test.cmd
```
