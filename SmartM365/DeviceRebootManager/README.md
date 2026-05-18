# Device Reboot Manager

Local WPF user notification app for SmartM365 device restart governance.

## Scripts

- `SmartM365-DeviceRebootManager-GUI.ps1`: improved SmartM365 GUI with status summary, safer state file, configurable postpone choices, preview mode, and cleaner logging.
- `Start-SmartM365-DeviceRebootManager-GUI.cmd`: launcher that starts the GUI in a PowerShell STA runspace with the console hidden.
- `Start-SmartM365-DeviceRebootManager-GUI-Test.cmd`: safe test launcher that simulates required mode with `PreviewOnly`.
- `SmartM365-DeviceRebootManager-GUI.config.json.template`: safe configuration model.
- `SmartM365-DeviceRebootManager-GUI.strings.psd1`: UI localization catalog.

## Configuration

`SmartM365-DeviceRebootManager-GUI.ps1` loads runtime configuration in this order:

1. `-ConfigPath` when provided.
2. `SmartM365-DeviceRebootManager-GUI.config.json` next to the script.
3. `%ProgramData%\SmartM365\DeviceRebootManager\SmartM365-DeviceRebootManager-GUI.config.json`.
4. `%APPDATA%\SmartM365\DeviceRebootManager\SmartM365-DeviceRebootManager-GUI.config.json`.

The committed template contains safe generic defaults only. Runtime `.json` files are local operational configuration.

`DefaultLanguage` can be `auto`, `en`, `fr`, `de`, `es`, `nl`, `it`, `pt`, `pl`, `ar`, `tr`, `sv`, `da`, `nb`, `fi`, `ro`, `hu`, `ja`, `ko`, `zh-Hans`, or `uk`. The app also maps common regional cultures, for example `zh-CN` to `zh-Hans` and `no` to `nb`. Arabic uses right-to-left layout.

`WindowTitle` overrides the localized window title when set. Leave it empty to use the title from the selected language.

When `WindowTitle` is empty, the app tries to detect the company name from the device Entra join or AD domain and prefixes the localized title, for example `Company - SmartM365 Device Reboot Manager`. Set `CompanyName` to force a specific company prefix, or leave it empty for automatic detection.

Users can change the UI language from the language selector in the GUI. The choice is stored in `%APPDATA%\SmartM365\DeviceRebootManager\SmartM365-DeviceRebootManager.preferences.json` and is reused on later launches unless `-DefaultLanguage` is passed on the command line.

`LanguageCatalogPath` can point to a custom PowerShell data file with the same structure as `SmartM365-DeviceRebootManager-GUI.strings.psd1`. Leave it empty to use the bundled catalog.

At each launch, the current `SmartM365-DeviceRebootManager.log` is archived to a timestamped log file and a fresh log is created. Only the 10 most recent archived logs are retained.

## Examples

Preview the required mode without rebooting:

```powershell
powershell.exe -STA -NoProfile -File .\DeviceRebootManager\SmartM365-DeviceRebootManager-GUI.ps1 -TestRequiredRestart -PreviewOnly
```

Preview the recommendation mode without rebooting:

```powershell
powershell.exe -STA -NoProfile -File .\DeviceRebootManager\SmartM365-DeviceRebootManager-GUI.ps1 -TestRecommendedRestart -PreviewOnly
```

Run in recommendation-only mode, with no mandatory restart state:

```powershell
powershell.exe -STA -NoProfile -File .\DeviceRebootManager\SmartM365-DeviceRebootManager-GUI.ps1 -NeverForceRestart
```

Use a specific config file:

```powershell
powershell.exe -STA -NoProfile -File .\DeviceRebootManager\SmartM365-DeviceRebootManager-GUI.ps1 -ConfigPath C:\ProgramData\SmartM365\DeviceRebootManager\SmartM365-DeviceRebootManager-GUI.config.json
```

Launch through the CMD wrapper:

```cmd
.\DeviceRebootManager\Start-SmartM365-DeviceRebootManager-GUI.cmd -TestRequiredRestart -PreviewOnly
```

Launch the safe test wrapper:

```cmd
.\DeviceRebootManager\Start-SmartM365-DeviceRebootManager-GUI-Test.cmd
```
