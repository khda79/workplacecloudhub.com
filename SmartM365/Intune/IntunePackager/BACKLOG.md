# Smart IntunePackager - Backlog of findings in other components

Issues found while building Smart IntunePackager. They concern existing components and were
intentionally **not** fixed as part of this work. Each item needs its own change, review and test.
Line numbers refer to the state of `main` on 2026-10-18.

## Intune Remediation GUI

File: `SmartM365/Intune/Remediation/GUI/SmartM365-IntuneRemediation-GUI.ps1`

1. **No merge of missing configuration keys.** `Initialize-GuiConfigurationFromTemplate` copies the
   template only when the local JSON is missing (around line 796). Keys later added to the template
   are never merged into an existing local file, which `SmartM365/RULES.md` requires.
2. **Log panel lines without the date.** `Write-GuiLog` formats UI and console lines with
   `HH:mm:ss` only (line 179). The log file already uses `yyyy-MM-dd HH:mm:ss`.
3. **Invalid `-Level` argument.** `Read-GuiConfiguration` calls `Write-GuiLog ... -Level WARN`
   (line 864) but `Write-GuiLog` has no `Level` parameter. The parameter binding error replaces the
   original configuration error inside the `catch` block.
4. **No STA relaunch.** The GUI throws when the thread is not STA (line 60) instead of relaunching
   itself with `-STA`.

## Shared GUI demo host

File: `SmartM365/Modules/SmartM365.UI/SmartM365.GuiDemo.ps1` (and its identical local copies)

5. **Hard-coded application name.** `Get-SmartM365DemoMarkup` replaces `$script:AppName` with
   `Smart Endpoint Diagnostics Analyzer` for every application (line 35). Any other GUI that uses
   `$script:AppName` in its XAML would show the wrong name in demo mode. Smart IntunePackager avoids
   the variable in its XAML for this reason.

## Win32 publishers (Windows 11 Upgrade Toolkit, Device Reboot Manager, Endpoint Diagnostics Analyzer)

Files:
- `SmartM365/Devices/Windows11UpgradeToolkit/IntuneWin32/Publish-SmartM365Windows11IntuneApp.ps1` (W11)
- `SmartM365/Devices/DeviceRebootManager/Deploy/SmartM365-DeviceRebootManager-PublishIntune.ps1` (DRM)
- `SmartM365/Devices/EndpointDiagnosticsAnalyzer/Deploy/SmartM365-EndpointDiagnosticsAnalyzer-PublishIntune.ps1` (EDA)

6. **Undocumented timeout property.** All three send `installCommandLineTimeoutInMinutes`
   (W11 line 715, DRM line 835, EDA line 837). The documented win32LobApp property is
   `installExperience.maxRunTimeInMinutes`; the configured timeout is probably ignored.
7. **Deprecated architecture property.** All three send `applicableArchitectures = 'x64'`
   (W11 line 696, DRM line 818, EDA line 820). Graph documents `allowedArchitectures` as its
   replacement.
8. **No Graph throttling handling.** No retry on 429/503/504 with `Retry-After` for Graph calls.
9. **App lookup without type check.** Existing apps are matched by display name without checking
   `@odata.type` (a non-Win32 app with the same name could be selected).
10. **No commit retry.** A `commitFileFailed` state ends the run without a second commit attempt.
11. **Loose Detection.xml parsing.** `Get-XmlValue` searches any node named `Name`, `Size`, etc.,
    instead of explicit `ApplicationInfo` / `EncryptionInfo` paths.
12. **W11 only - no disconnect before connect.** `Connect-MgGraph` is called without a prior
    `Disconnect-MgGraph` (lines 735, 778, 846), contrary to `SmartM365/RULES.md`, and without
    `-TenantId`.
13. **W11 only - `$matches` variable.** `Resolve-ExistingIntuneWin32App` assigns `$matches`
    (line 261), which is the automatic `$Matches` variable (PowerShell variables are case-insensitive).
14. **W11 only - StrictMode 2.0** instead of `Latest`.
15. **DRM/EDA only - no SAS renewal.** Large or slow uploads fail when the Azure Storage SAS
    expires; W11 renews it, DRM and EDA do not.

Smart IntunePackager's `SmartM365.IntunePackager.Graph.psm1` addresses items 6 to 15 in its own
engine, split into one function per step, so it can serve as the base of a future shared publisher.
