# Smart Intune Remediation Manager

PowerShell 7/WPF operator interface and source-controlled package workspace for Microsoft Intune Remediations.

The GUI reports version `1.4`. It browses local packages and Intune `deviceHealthScripts`, compares script content, runs PSScriptAnalyzer, publishes or updates packages, exports scripts and reports, and provides guarded reset and delete workflows.

## Distribution Status

Smart Intune Remediation Manager GUI version `1.4` is published as a standalone GitHub Release and remains available as repository source. It is not published as a PowerShell Gallery package and does not include an installer or automatic updater.

- Release: [Smart Intune Remediation Manager v1.4](https://github.com/khda79/workplacecloudhub.com/releases/tag/smart-intune-remediation-manager-v1.4)
- Standalone ZIP: [SmartM365-IntuneRemediationManager-1.4.zip](https://github.com/khda79/workplacecloudhub.com/releases/download/smart-intune-remediation-manager-v1.4/SmartM365-IntuneRemediationManager-1.4.zip)
- Published checksum: [SmartM365-IntuneRemediationManager-1.4.sha256](https://github.com/khda79/workplacecloudhub.com/releases/download/smart-intune-remediation-manager-v1.4/SmartM365-IntuneRemediationManager-1.4.sha256)
- Release manifest: [SmartM365-IntuneRemediationManager-1.4.manifest.json.txt](https://github.com/khda79/workplacecloudhub.com/releases/download/smart-intune-remediation-manager-v1.4/SmartM365-IntuneRemediationManager-1.4.manifest.json.txt)
- Release notes: [SmartM365-IntuneRemediationManager-1.4-RELEASE-NOTES.md](https://github.com/khda79/workplacecloudhub.com/releases/download/smart-intune-remediation-manager-v1.4/SmartM365-IntuneRemediationManager-1.4-RELEASE-NOTES.md)
- Source and documentation: [SmartM365/Intune/Remediation](https://github.com/khda79/workplacecloudhub.com/tree/main/SmartM365/Intune/Remediation)
- Product page: [Smart Intune Remediation Manager](https://workplacecloudhub.com/tools/smart-intune-remediation-manager/)
- Updates: download a newer GitHub Release when available or review repository changes and use a fresh clone or `git pull`; there is no in-app update check.

The release packages the source identified by its Git tag and published manifest. Use the published `.sha256` asset as the authoritative ZIP checksum and compare it with the downloaded archive before extraction, then verify the Authenticode status and expected signer of the PowerShell scripts before use.

## Requirements And Launch

- Windows with WPF.
- PowerShell 7 or later, started in STA mode. The GUI script declares `#Requires -Version 7.0`.
- `Microsoft.Graph.Authentication` for delegated Graph sign-in and requests.
- `PSScriptAnalyzer` for script analysis. If it is absent, the GUI installs it from PowerShell Gallery for the current user when analysis is first requested.
- `ImportExcel` is optional. The GUI asks before installing it for the current user; CSV export remains available when it is absent or installation fails.

From an extracted release folder:

```powershell
.\Start-SmartM365-IntuneRemediation-GUI.cmd
```

Or from a repository clone:

```powershell
git clone https://github.com/khda79/workplacecloudhub.com.git
cd .\workplacecloudhub.com\SmartM365\Intune\Remediation
.\Start-SmartM365-IntuneRemediation-GUI.cmd
```

Or launch the script directly:

```powershell
pwsh -STA -NoProfile -File .\GUI\SmartM365-IntuneRemediation-GUI.ps1
```

The CMD launcher prefers `%ProgramFiles%\PowerShell\7\pwsh.exe`. Its legacy `powershell.exe` fallback cannot satisfy the script's PowerShell 7 requirement, so install PowerShell 7 before using the launcher.

## Organization

- `Packages/`: detection, remediation, diagnostic, and action scripts organized by Intune scenario.
- `GUI/`: WPF operator interface for local package review, Intune cloud comparison, PSScriptAnalyzer checks, publish/update actions, export, reset history, and cloud deletion.

## Authentication Model

The GUI uses delegated interactive Microsoft Graph authentication. It does not use stored credentials, client secrets, certificates, or SmartM365 app-only authentication.

Typical delegated scopes are documented in the GUI README:

- `DeviceManagementScripts.ReadWrite.All`
- `DeviceManagementConfiguration.Read.All`
- `Group.Read.All`

The GUI uses Microsoft Graph `beta` endpoints. An active Intune tenant and the Remediations licensing and Intune RBAC rights required by Microsoft remain external prerequisites; delegated Graph consent does not replace the signed-in operator's tenant permissions.

## Security And Change Boundaries

- New and updated packages are sent with `runAsAccount = system`, 64-bit execution, default role scope tag `0`, and `enforceSignatureCheck = false`.
- Publishing a newly named selected package creates it without a second create-confirmation dialog. Existing-package updates, detection-only publishing, and publish-all have explicit review or confirmation paths.
- `Reset history` creates a timestamped replacement, copies assignments, and then deletes the original object. The replacement keeps the timestamped duplicate name.
- `Delete` permanently removes the selected cloud remediation after a warning. It does not remove local source files.
- Local and cloud exports can contain tenant metadata, group IDs or names, assignments, script content, and execution data. Keep them outside Git and handle them as operational evidence.
- Repository PowerShell scripts are Authenticode-signed by `workplacecloudhub.com` with thumbprint `D70ECB7B00377EBFB76B304C08DFC6620584E114`. The certificate is self-signed, so trust is not automatic and still depends on the local certificate stores and execution policy.

Review the target tenant, selected package, name-prefix mapping, script content, assignments, execution context, and signature policy before any cloud write. A local `-ValidateOnly` run does not validate tenant consent, Intune RBAC, licensing, Graph behavior, or endpoint execution.

## Typical Workflow

1. Edit or review a package under `Packages/`.
2. Run PSScriptAnalyzer locally or from the GUI.
3. Use the GUI to compare local package content with Intune cloud content.
4. Publish one package, publish detection-only, or publish all packages after reviewing the create/update summary.
5. Export cloud scripts, assignments, or execution reports when evidence is needed.

## Package Shape

Most recurring remediation packages use a pair of scripts:

```text
SmartM365-<Scenario>-Detection.ps1
SmartM365-<Scenario>-Remediation.ps1
```

Some folders contain `*-Action.ps1` or diagnostic scripts for one-off operator actions that should not run as recurring Intune remediations.

## Local Files

Runtime configuration, logs, local archives, downloaded cloud exports, and execution reports are ignored by Git. Keep tenant-specific values, group IDs, and operational exports out of commits.

## More Documentation

- `Packages/README.md`: package organization and scenario map.
- `GUI/README.md`: GUI launch, actions, authentication, configuration, and logs.

## Offline demo mode

Launch `Start-SmartM365-IntuneRemediation-Demo.cmd` for screenshots and presentations. The original WPF screens, tabs and options remain available with fictional Contoso data and visibly marked simulated results. Use the Scene selector (`Overview`, `Issues`, `Completed`), or pass `-DemoScenario Issues`. `-DemoView` selects a named tab where available; `-DemoWindowWidth` and `-DemoWindowHeight` control the capture size.

Business actions are simulated. Demo startup bypasses operational configuration, authentication, collection, remote commands and mail delivery. Edits remain in memory. **Export demo outputs** generates only synthetic files under a unique `%TEMP%\SmartM365-Demo\` session folder. HTML-capable tools retain their report views; notification campaigns preview every shipped language template with fictional tokens. No real tenant or device evidence is required.

Keep the application folder, its original report/template sources, local branding assets, `SmartM365.GuiDemo.ps1`, `SmartM365.GuiDemo.Reports.ps1` and the local `*-Demo.ps1` profile together when copying the application. `-Demo -ValidateOnly` builds the WPF demo without displaying windows or opening files. Run `SmartM365/Tests/Test-SmartM365GuiDemoApplicationsOffline.ps1` in an STA process for offline interaction and export checks.

## Release 1.4

Includes 30 remediation packages, the four reviewed corrections, and eight additional guarded packages documented in [Packages/README.md](Packages/README.md). The synthetic demo now shows all 30 packages, including both Pending Reboot Age variants. Offline validation covers the package behavior and demo controls; it does not qualify a tenant deployment or endpoint repair.
