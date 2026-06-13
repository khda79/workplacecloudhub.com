# Smart Device Reboot Manager

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

At each launch, the current `SmartM365-DeviceRebootManager.log` is archived to a timestamped log file and a fresh log is created. Only the 10 most recent archived logs are retained.

## Intune Win32 Deployment

Package the whole `DeviceRebootManager` folder as an Intune Win32 app. The install script copies the runtime files to:

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

Use the detection script as a custom detection rule:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Deploy\SmartM365-DeviceRebootManager-Detection.ps1
```

The install script creates `SmartM365-DeviceRebootManager-GUI.config.json` from the committed template when no runtime config exists. To deploy a custom config, include a JSON file in the package and pass it with `-ConfigSourcePath`.

Example:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Deploy\SmartM365-DeviceRebootManager-Install.ps1 -ConfigSourcePath .\SmartM365-DeviceRebootManager-GUI.config.json -RepeatIntervalMinutes 120
```

To recreate only the scheduled task on an already installed device:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Deploy\SmartM365-DeviceRebootManager-CreateScheduledTask.ps1 -RepeatIntervalMinutes 240
```

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
