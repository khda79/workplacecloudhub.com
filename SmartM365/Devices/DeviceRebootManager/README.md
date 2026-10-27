# Smart Device Reboot Manager

## Current distribution 1.0.0 — build 2

The package and GUI product version are **1.0.0**. **Build 2** aligns the visible interface and report versions. The [1.0.0 release](https://github.com/khda79/workplacecloudhub.com/releases/tag/device-reboot-manager-v1.0.0) contains only the current build 2 assets; superseded releases and original build assets have been removed.

| Download | File |
| --- | --- |
| Standalone ZIP | [SmartM365-DeviceRebootManager-1.0.0-build.2.zip](https://github.com/khda79/workplacecloudhub.com/releases/download/device-reboot-manager-v1.0.0/SmartM365-DeviceRebootManager-1.0.0-build.2.zip) |
| Intune Win32 package | [SmartM365-DeviceRebootManager-1.0.0-build.2.intunewin](https://github.com/khda79/workplacecloudhub.com/releases/download/device-reboot-manager-v1.0.0/SmartM365-DeviceRebootManager-1.0.0-build.2.intunewin) |
| SHA-256 checksums | [SmartM365-DeviceRebootManager-1.0.0-build.2.sha256](https://github.com/khda79/workplacecloudhub.com/releases/download/device-reboot-manager-v1.0.0/SmartM365-DeviceRebootManager-1.0.0-build.2.sha256) |
| DemoAssets | [SmartM365-DeviceRebootManager-1.0.0-build.2-DemoAssets.zip](https://github.com/khda79/workplacecloudhub.com/releases/download/device-reboot-manager-v1.0.0/SmartM365-DeviceRebootManager-1.0.0-build.2-DemoAssets.zip) |
| Release notes | [SmartM365-DeviceRebootManager-1.0.0-build.2-RELEASE-NOTES.md](https://github.com/khda79/workplacecloudhub.com/releases/download/device-reboot-manager-v1.0.0/SmartM365-DeviceRebootManager-1.0.0-build.2-RELEASE-NOTES.md) |
| Release manifest | [SmartM365-DeviceRebootManager-1.0.0-build.2.manifest.json.txt](https://github.com/khda79/workplacecloudhub.com/releases/download/device-reboot-manager-v1.0.0/SmartM365-DeviceRebootManager-1.0.0-build.2.manifest.json.txt) |

Launch `Start-SmartM365-DeviceRebootManager-Demo.cmd` from the root of the extracted standalone release. The separate DemoAssets archive contains the current fictional GUI captures, applicable HTML reports, synthetic exports and website/presentation variants. It performs no tenant or device operation.

The signed runtime source is identified by the [build 2 source tag](https://github.com/khda79/workplacecloudhub.com/tree/device-reboot-manager-v1.0.0-build.2/SmartM365/Devices/DeviceRebootManager). GitHub publication does not publish a version to the PowerShell Gallery service.

## Next release 1.1.0 (source, not yet published)

The repository source is version **1.1.0**. It is not published yet; until then the 1.0.0 build 2 downloads above remain current. When built and published, the release will use the tag `releases/tag/device-reboot-manager-v1.1.0` and the artifacts `SmartM365-DeviceRebootManager-1.1.0.zip`, `SmartM365-DeviceRebootManager-1.1.0.intunewin` and `SmartM365-DeviceRebootManager-1.1.0.sha256`.

Changes in 1.1.0:

- Installs to `C:\Program Files\WCH\SmartDeviceRebootManager` (administrator-only) and removes the legacy, user-writable `C:\ProgramData\SmartM365\DeviceRebootManager` installation.
- Restart enforcement: no forced restart when the boot time is unreadable, forced and verified restart command, wall-clock countdown persisted per boot, locale-independent disconnected-session detection.
- Gallery channel: working automatic-update task, signed file catalog covering every packaged file, one signature rule, rollback to an installed module version.
- Release and Intune publication: the `.intunewin` is decrypted and compared with the runtime, interrupted publications can be resumed, SAS upload URLs are renewed.

Local WPF user notification app for SmartM365 device restart governance.

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

## Requirements, Privileges, and Boundaries

- The app is Windows-only and uses Windows PowerShell 5.1, WPF, and the Windows ScheduledTasks module.
- The published Intune Win32 metadata targets x64 devices and declares Windows 10 version 1607 as its minimum supported operating system.
- Direct GUI and preview launches run in the signed-in user's interactive session. The deployed GUI task also runs interactively with `RunLevel Limited`; it does not run the user interface in session 0 or with administrator rights.
- Installation, update, and removal write under `C:\Program Files\WCH\SmartDeviceRebootManager` and register or remove scheduled tasks, so they must run elevated. The Intune Win32 install and uninstall commands are designed to run as `SYSTEM`.
- A real restart remains subject to local Windows rights and policy. Use `-PreviewOnly` or the test launcher to review required and recommended states without restarting the device.
- The shipped PowerShell files are Authenticode-signed with the self-signed `workplacecloudhub.com` code-signing certificate, thumbprint `D70ECB7B00377EBFB76B304C08DFC6620584E114`. A valid file signature does not make that certificate trusted automatically on every endpoint; trust still depends on the local certificate stores and execution policy. Verify the published SHA-256 values as well as the expected signer before deployment.
- The tool provides local restart notification and governance. Intune remains responsible for application assignment, rollout, retry, reporting, and rollback when the Win32 package is used.

## Demo mode

Double-click `Start-SmartM365-DeviceRebootManager-Demo.cmd` to open the normal WPF
window with fictional Contoso data. Use the scene selector to show **Recommended**,
**Required**, or **Postponed**, and the language selector to prepare localized captures.
The countdown stays frozen. Restart, postpone and reminder actions are simulated
in memory; they do not restart the computer or write operational state or preferences.
The demo does not load runtime JSON, inspect the device, require administrator rights,
or modify an installed copy. The title and content identify the data as a demo.

```powershell
powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -File .\SmartM365-DeviceRebootManager-GUI.ps1 -Demo -DemoScenario Recommended -DefaultLanguage en
```

Use `-DemoWindowWidth` (default 1100) and `-DemoWindowHeight` (default 720) to adjust
the window for captures. Keep the local demo helper, strings, splash and branding
assets with the GUI when copying the application. Existing preview/test modes retain
their original behavior; the demo is a separate startup path.

`-Demo -ValidateOnly` builds the demo window without showing it or its splash and
returns the WPF window and controls for offline validation. Run the targeted suite
from the repository root:

```powershell
powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -File .\SmartM365\Tests\Test-SmartM365GuiDemoOffline.ps1 -Application RebootManager
```

## Configuration

`SmartM365-DeviceRebootManager-GUI.ps1` loads runtime configuration in this order:

1. `-ConfigPath` when provided.
2. `SmartM365-DeviceRebootManager-GUI.config.json` next to the script.
3. `%ProgramData%\SmartM365\DeviceRebootManager\SmartM365-DeviceRebootManager-GUI.config.json`. Standard users can create files there, so deployed installs always pass `-ConfigPath` to the installed, administrator-owned file.
4. `%APPDATA%\SmartM365\DeviceRebootManager\SmartM365-DeviceRebootManager-GUI.config.json`.

The committed template contains safe generic defaults only. Runtime `.json` files are local operational configuration.

`DefaultLanguage` can be `auto`, `en`, `fr`, `de`, `es`, `nl`, `it`, `pt`, `pl`, `ar`, `tr`, `sv`, `da`, `nb`, `fi`, `ro`, `hu`, `ja`, `ko`, `zh-Hans`, or `uk`. The app also maps common regional cultures, for example `zh-CN` to `zh-Hans` and `no` to `nb`. Arabic uses right-to-left layout.

`WindowTitle` overrides the localized window title when set. Leave it empty to use the title from the selected language (`PC restart`, `Redémarrage du PC`, ...). The title bar never shows the version; the version is shown in the window footer.

The window is a single column designed for standard users: status, days since the last restart with the allowed or recommended interval, last restart date in the selected language, the awareness banner, the countdown (required mode only), then the actions. Only the actions available in the current mode are shown: `Remind me later` in recommended mode, postponement while postponements remain in required mode, and `Close` unless closing is blocked. The footer shows the language selector, the PC name and the version. Activity is written to the log file only; it is not shown in the window.

When `WindowTitle` is empty, the app tries to detect the company name from the device Entra join or AD domain and prefixes the localized title, for example `Company - Smart Device Reboot Manager`. Only device-level sources are used (HKLM `CloudDomainJoin`, the device `TenantName` from `dsregcmd /status`, then the AD domain); a work or school account added to a user profile (HKCU `WorkplaceJoin`) is ignored. Set `CompanyName` to force a specific company prefix, or leave it empty for automatic detection.

Users can change the UI language from the language selector in the GUI. The choice is stored in `%APPDATA%\SmartM365\DeviceRebootManager\SmartM365-DeviceRebootManager.preferences.json` and is reused on later launches unless `-DefaultLanguage` is passed on the command line. Set `ForceLanguage` to `true` to ignore the saved user preference and lock the GUI to `DefaultLanguage`.

`LanguageCatalogPath` can point to a custom PowerShell data file with the same structure as `SmartM365-DeviceRebootManager-GUI.strings.psd1`. Leave it empty to use the bundled catalog.

The window always shows an awareness banner (`AwarenessTitle`, `AwarenessText`) explaining that sleep is not enough and that only a restart fully installs security updates, frees memory and fixes slowdowns. Both keys are translated in every bundled language; a custom catalog without them falls back to English.

`SplashEnabled` controls the startup WorkplaceCloudHub splash. Set it to `false` to disable the splash. Keep it enabled and customize `SplashMinimumDurationMs`, `SplashProductName`, `SplashBadgeText`, `SplashSubtitle`, and `SplashLogoPath` to adapt the splash for a deployment. The splash URL remains `https://workplacecloudhub.com`.

Missing keys are added from the committed template when the configuration is loaded. The installer persists them on upgrade; the GUI saves them only when the configuration file is writable, otherwise it applies them in memory. Boolean values accept JSON booleans or the strings `true`/`false`; any other value is rejected.

At each launch, the current `SmartM365-DeviceRebootManager.log` is archived to a timestamped log file and a fresh log is created. Only the 10 most recent archived logs are retained.

## Intune Win32 Deployment

Build the clean allow-listed package, then use the returned `IntuneSourcePath`
(`...\SmartM365.DeviceRebootManager\0.1.0\Runtime`) as the Intune Win32 source.
Do not package the whole working folder because it can contain private runtime
configuration. The install script copies the runtime files to:

```text
C:\Program Files\WCH\SmartDeviceRebootManager
```

Only administrators and `SYSTEM` can write there, so a standard user cannot
modify code that runs in other users' sessions or, for the Gallery updater, as
`SYSTEM`. Versions before 1.1.0 installed into
`C:\ProgramData\SmartM365\DeviceRebootManager`, where standard users can
create files and folders. The 1.1.0 installer and uninstaller never walk that
legacy folder, because a junction swapped in during a recursive `SYSTEM` deletion
could redirect it. They delete only the files the previous installer wrote, by
exact name, and then the folder if it is empty; they act only when the folder and
its parent are owned by `SYSTEM` or Administrators, and leave any other entry in
place. Nothing runs from that folder any more, because the scheduled task now
starts the Program Files copy. Its configuration is not migrated: deploy a
custom configuration again with `-ConfigSourcePath`. Per-user log,
reminder state and language preference stay under
`%APPDATA%\SmartM365\DeviceRebootManager`.

It also creates this scheduled task:

```text
\SmartM365\Device Reboot Manager
```

The task runs in the interactive user session, not in session 0. It starts at user logon and repeats every 240 minutes by default. The GUI script exits silently when the restart threshold is not reached, so regular execution is expected.

Restart behavior:

- When the last boot time cannot be read, the GUI exits without reminder or restart; it never assumes the required threshold.
- A forced restart uses `Restart-Computer -Force` for administrators and `shutdown.exe /r /f /t 0` for standard users, so open applications cannot block it. The command result is checked; a failure is logged, shown to the user and recorded as `RestartFailed`, which does not start the reminder cooldown.
- The forced countdown follows the wall clock. Its deadline and postpone count are stored per boot in the user's reminder state, so closing the process or signing out does not grant a new countdown or new postpones. An already expired deadline gets a 5-minute grace period. Preview and test runs never store this state.
- Disconnected sessions are detected from the Windows session state, independently of the display language.

Recommended Intune Win32 commands:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Deploy\SmartM365-DeviceRebootManager-Install.ps1
```

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Deploy\SmartM365-DeviceRebootManager-Uninstall.ps1
```

Use `Deploy\SmartM365-DeviceRebootManager-Detection.ps1` directly as the
Intune custom detection script. Its default expected version is
`1.1.0`; update that value for every new Intune package:

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
    -IntuneWinPath C:\Temp\SmartM365-DeviceRebootManager-1.0.0-build.2.intunewin
```

Publish the stable app, supersede the prior preview as an in-place update, reuse
the assigned security group, and assign the stable app as `Required`:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass `
    -File .\Deploy\SmartM365-DeviceRebootManager-PublishIntune.ps1 `
    -IntuneWinPath C:\Temp\SmartM365-DeviceRebootManager-1.0.0-build.2.intunewin `
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

If publication fails after the app is created, rerun with
`-ResumeAppId <incomplete-app-id>`. The script validates the existing app name,
publisher and package-version notes before it resumes. When the app has no
committed content version (a failure during upload), it uploads and commits the
package into that app instead of creating a duplicate; otherwise it resumes
supersedence and assignment only. The duplicate-name check points to
`-ResumeAppId` when an app with the same name already exists.

To recreate only the scheduled task on an already installed device:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Deploy\SmartM365-DeviceRebootManager-CreateScheduledTask.ps1 -RepeatIntervalMinutes 240
```

On Windows 10 2004 (build 19041) and later, the task runs `conhost.exe --headless "...\powershell.exe" -STA -NoProfile -WindowStyle Hidden ...`. The headless console host prevents the black console window that `powershell.exe -WindowStyle Hidden` alone briefly shows at every run, including runs where the GUI exits silently below the threshold. Older builds keep the direct `powershell.exe` action. The task result is the console host result, not the GUI exit code. Existing installations keep their previous action until the task is recreated (reinstall, or the command above). The `Pending-Reboot-Age-DeviceRebootManager` Intune remediation accepts both actions.

## PowerShell Gallery packaging

The GitHub build 2 runtime uses module version `1.0.0`. Gallery service publication is separate from the GitHub distribution linked above.

Build and validate the package locally:

Local release builders generate intermediate files named `SmartM365-DeviceRebootManager-1.0.0.zip`, `SmartM365-DeviceRebootManager-1.0.0.intunewin` and `SmartM365-DeviceRebootManager-1.0.0.sha256`. These are local build outputs; use the **build 2 download links above** for published packages.

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

The following optional Gallery-channel commands install the version available on that service; they do not select the GitHub build 2 package. Use the downloads above for this release. An elevated Windows PowerShell session can deploy a Gallery installation:

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
checks the stable PowerShell Gallery channel by default, validates the module
against the pinned WorkplaceCloudHub signer, preserves the local runtime
configuration, and redeploys the runtime.

Package integrity: every PowerShell file must carry a valid signature from the
pinned WorkplaceCloudHub certificate, and the signed
`SmartM365.DeviceRebootManager.Files.cat` catalog must match every packaged file
(launchers, configuration template, version file, images): a changed, missing or
extra file stops the installation. Signatures are checked with WinVerifyTrust:
a signature is accepted only when the file digest is intact and the signer is the
pinned certificate. A root that the device does not trust (the current
self-signed certificate) is accepted with a warning; any other failure, such as a
modified file, is rejected.

`Install-SmartM365DeviceRebootManager` deploys the runtime of the module version
that runs it, so importing an older installed version
(`Import-Module SmartM365.DeviceRebootManager -RequiredVersion <version>`) and
installing again rolls back. The newest installed version is selected only after
an automatic update.

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

`-TrustSignerCertificate` is optional: an untrusted self-signed root is accepted
with a warning when the signer is the pinned certificate. Use it only on the
isolated pilot VM to test with a trusted chain. Add
`-CleanupAfterTest` to remove the product and the certificate trust entries
created by the validation.

## Examples

Run these examples from the `SmartM365/Devices/DeviceRebootManager` folder or from the root of the extracted standalone release.

Preview the required mode without rebooting:

```powershell
powershell.exe -STA -NoProfile -File .\SmartM365-DeviceRebootManager-GUI.ps1 -TestRequiredRestart -PreviewOnly
```

Preview the recommendation mode without rebooting:

```powershell
powershell.exe -STA -NoProfile -File .\SmartM365-DeviceRebootManager-GUI.ps1 -TestRecommendedRestart -PreviewOnly
```

Run in recommendation-only mode, with no mandatory restart state:

```powershell
powershell.exe -STA -NoProfile -File .\SmartM365-DeviceRebootManager-GUI.ps1 -NeverForceRestart
```

Use a specific config file:

```powershell
powershell.exe -STA -NoProfile -File .\SmartM365-DeviceRebootManager-GUI.ps1 -ConfigPath "C:\Program Files\WCH\SmartDeviceRebootManager\SmartM365-DeviceRebootManager-GUI.config.json"
```

Launch through the CMD wrapper:

```cmd
.\Start-SmartM365-DeviceRebootManager-GUI.cmd -TestRequiredRestart -PreviewOnly
```

Launch the safe test wrapper:

```cmd
.\Start-SmartM365-DeviceRebootManager-GUI-Test.cmd
```


## Distribution version transition

The 1.0.0 package uses GUI component 1.0.0 and includes the complete signed demo runtime. Only the current build 2 release assets remain published. Git history and source tags preserve earlier source revisions. Local DemoAssets stay Git-ignored; reviewed demo archives are separate GitHub release assets.
