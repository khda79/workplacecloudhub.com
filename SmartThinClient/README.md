# Smart ThinClient Shell

Local Windows endpoint tooling for Citrix, Azure Virtual Desktop (AVD), and web workspaces. The current distribution is the [SmartThinClient source folder](https://github.com/khda79/workplacecloudhub.com/tree/main/SmartThinClient). As checked on 2026-09-09, there is no dedicated Smart ThinClient GitHub release, installer, PowerShell Gallery module, Intune package, or automatic updater.

## Prerequisites and validation limits

Use **64-bit Windows PowerShell 5.1** on Windows, with .NET Framework/WPF and an interactive desktop for the GUI and launcher. CMD wrappers invoke `powershell.exe`; launchers use STA. PowerShell 7 regression tests cover isolated functions; they do not certify the complete application on PowerShell 7.

Audit and Preview normally do not need elevation. Apply and Restore require an elevated administrator session. Client software, browser, portal URLs, subscriptions, authentication, connectivity and endpoint policy must be prepared by the operator. No tenant, Citrix delivery group or AVD host pool is provisioned by this tool. A detected executable, package, cmdlet or CIM class is a local signal, not proof of a usable remote session or Windows edition entitlement.

The embedded shell uses the legacy WPF **WebBrowser** control, not Edge/WebView2. Modern Citrix and Windows App authentication/session compatibility has **not** been validated in this control. Set `UseWebShell` to `false` to use an installed external browser, then verify the portal in your environment. The window, provider buttons and optional browser kiosk switches are not a security boundary or an enforced URL allowlist.

References: [WPF WebBrowser](https://learn.microsoft.com/en-us/dotnet/api/system.windows.controls.webbrowser), [Windows App connection requirements](https://learn.microsoft.com/en-us/windows-app/get-started-connect-devices-desktops-apps), [Shell Launcher configuration](https://learn.microsoft.com/en-us/windows/configuration/shell-launcher/wesl-usersetting).

No live Apply/Restore, Windows logon, Citrix/AVD session, authentication, kiosk, client installation or tenant deployment was tested for this candidate. See [VALIDATION.md](VALIDATION.md) for the offline evidence and remaining lab checks.

## Installation and updates

Clone/download the repository and retain the complete `SmartThinClient/` folder, including `Assets/`, `Profiles/`, splash script and CMD wrappers. Start in a user-writable working copy for Audit/Preview. Review the source and its signature before running it:

```powershell
Get-AuthenticodeSignature .\SmartThinClient-Shell.ps1
.\SmartThinClient-Shell.ps1 -Cli -Action Preview
```

The repository uses a self-signed WorkplaceCloudHub code-signing certificate. Certificate trust is an operator decision; a valid signature does not certify functionality. CMD wrappers and locally generated launchers explicitly use process-scoped `ExecutionPolicy Bypass`. They do not change the machine execution policy. Generated launcher files contain local configuration and are not Authenticode-signed; package-signature validation does not extend to those files.

For updates, keep a separate backup of local JSON, output and rollback files, compare the new templates, and rerun Preview. There is no automatic migration or update check. Existing local profiles remain explicit overrides: set inherited fields to `"__USE_GLOBAL__"` when adopting the corrected templates. Never overwrite local workspace URLs with a blank template.

## Modes

| Mode | Behavior |
| --- | --- |
| Audit (default) | Inspects the local OS, users, clients, browser and kiosk capabilities. Without `-Cli`, opens the audit GUI. |
| Preview | Same local inspection, with `PreviewOnly` evidence. It does not simulate every Apply operation or certify readiness. |
| Launch | Generates a unique file under `LaunchOnly/<id>/Launcher/` and starts a child PowerShell process. It does not change the installed launcher or configure logon. `Launched` means the process was started, not that a remote session succeeded. |
| Apply | Requires elevation, `AllowApply: true` and the confirmation phrase. Performs preflight, saves rollback before account/configuration writes, writes a persistent launcher, then applies the explicitly configured options. |
| Restore | Requires elevation, `AllowRestore: true` and the confirmation phrase. Validates the rollback schema/computer/paths before writes, restores recorded values and the previous launcher, and surfaces failures. |
| ValidateOnly | Loads/validates main configuration without opening the GUI or applying endpoint settings. It is not a runtime, XAML or connectivity certification. |

Audit, Preview and ValidateOnly do not change Windows shell, policies or startup settings, but **do write local JSON configuration** when initializing templates. Audit/Preview also create logs and evidence. Launch may open applications, authenticate when used interactively and expose optional sign-out/restart/shutdown controls.

```powershell
.\SmartThinClient-Shell.ps1 -Cli -Action Audit -Profile Citrix
.\SmartThinClient-Shell.ps1 -Cli -Action Preview -Profile AVD
.\SmartThinClient-Shell.ps1 -Cli -Action Launch -Profile Hybrid
.\SmartThinClient-Shell.ps1 -ValidateOnly
```

Hybrid offers Citrix, AVD and WebOnly; closing the choice window cancels startup. With `HybridSelectionAtStartup: false`, it uses `HybridDefaultProvider` (or the legacy `WebShellPreferredProvider` override).

The GUI provides audit/preview and output access; Apply/Restore buttons explain the guarded CLI workflow. CMD Apply/Restore wrappers accept additional arguments; without a confirmation argument they remain blocked by default.

## Configuration

The main runtime file is `SmartThinClient-Shell.config.json`, initialized from `SmartThinClient-Shell.config.template.json`. Profile runtime files are initialized from `Profiles/*.profile.json.template`. Additional missing template keys are merged into these local files. All runtime JSON, logs and URLs stay out of Git. `-ConfigPath` selects an alternate existing main JSON; profiles are still read from this tool's `Profiles/` directory. `-OutputRoot` overrides the configured output root.

Main values override built-in defaults; non-empty profile values override the main configuration. The sentinel `"__USE_GLOBAL__"` inherits the main value. Boolean switches must be JSON `true`/`false`, never strings.

- `DefaultProfile`: Auto, Citrix, AVD, WebOnly or Hybrid. Auto chooses from detected local clients.
- `PreferredAccessMode`: Web or Native; `UseWebShell`: embedded WPF or external browser for Web mode.
- `CitrixWebUrl`, `AvdWebUrl`, `WebOnlyUrl`: local portal URLs. Legacy aliases include WorkspaceUrl/StoreUrl for Citrix, FeedUrl for AVD, and WebShellHomeUrl/FallbackWebUrl for web fallback. A subscription feed URL is not necessarily a browser portal.
- `CitrixWorkspacePath`, `AvdClientPath`, `BrowserPath`: explicit paths, otherwise detection searches common installation locations.
- Native Citrix starts the detected executable without connection arguments. Detection of CDViewer/wfica32 alone does not guarantee a standalone workspace launch.
- Native AVD first tries classic `msrdcw.exe`, then a configured `WindowsAppLaunchUri`, then web fallback. Detecting the Windows App package for the current user alone does not supply a launch URI or validate another user's installation.
- `BrowserKioskMode`: passes browser kiosk arguments when using external web launch. Validate the chosen browser; this is not Windows Assigned Access.
- `WebShellAllowExternalBrowser`, `WebShellAllowPowerControls`, `WebShellAllowLimitedSettings`, `WebShellAllowedSettingsPages`: optional UI controls. Their defaults expose sign-out/restart/shutdown and selected Windows settings; turn them off for a restricted review.
- `AllowApply`, `AllowRestore`: both false by default. `RequireConfirmationPhrase` defaults true.

`LogRetentionCount`, `RollbackRetentionCount`, `PreferredWorkspaceUrl`, `DefaultLanguage`, `ForceLanguage`, and profile metadata such as `LaunchMode`, `AllowedProviders`, `ExpectedExecutableNames` are not enforced as operational controls in this implementation. Logs/rollback/LaunchOnly files need operator-managed retention. The GUI is English.

## Apply and Restore scope

`AutoLaunchMode: RunKey` writes **HKLM** `SOFTWARE\Microsoft\Windows\CurrentVersion\Run\SmartThinClientShell`, affecting sign-in for all users. `AutoLaunchMode: None` writes no auto-launch value; it does not remove a previous installation. `TargetUserName` does not scope this machine Run key or machine policies. Use an administrator-controlled persistent output directory and verify read access for the target user; do not point a machine startup command at a directory writable by untrusted users.

`EnableShellLimitations: true` with Basic writes machine Explorer policies NoRun, NoControlPanel and NoViewContextMenu. Strict also writes DisableCMD. None writes no restrictions. These are machine registry writes with policy/OS-dependent effects, not a per-user lockdown guarantee; no live effect was tested. Group Policy/MDM may override them. The tool does not change autologon, Winlogon Shell/Userinit, or install/enable Windows features.

Assigned Access and Shell Launcher cannot be requested together. Both require an existing enabled local non-administrator user and the relevant local APIs. Kiosk preflight rejects missing prerequisites before applying the Run key or policies. Assigned Access requires a supplied `AssignedAccessAppUserModelId` and no existing Assigned Access configuration. Shell Launcher requires its feature already enabled, no custom shell for the target SID, and a persistent embedded web shell; its Windows edition/licensing and portal suitability still require a lab review. Existing kiosk configurations are deliberately not overwritten because this tool does not capture their complete prior state.

For Launcher mode, optional DedicatedUser creation requires `CreateDedicatedLocalUser: true` plus `-DedicatedUserPassword` as a SecureString. It uses the built-in Users group SID for localized Windows. Passwords are never written to JSON. A created user is **not deleted by Restore**, and no automatic logon or client provisioning is configured. Prepare kiosk users separately.

```powershell
# Review local JSON and Preview before these guarded, write-capable commands.
.\SmartThinClient-Shell.ps1 -Cli -Action Apply -Profile Hybrid -ConfirmApply "APPLY SMARTTHINCLIENT"
.\SmartThinClient-Shell.ps1 -Cli -Action Restore -RollbackPath "C:\ProgramData\SmartThinClient\Shell\Rollback\<reviewed-file>.json" -ConfirmRestore "RESTORE WINDOWS SHELL"
```

Restore accepts schema 2 rollback files from the same computer and original OutputRoot. Legacy files require manual review: they included unrelated Winlogon values and lacked launcher history. Only registry values selected for the Apply operation are captured. Restore clears newly configured Assigned Access or removes the newly configured custom shell, restores/removes the launcher as recorded, and retains evidence and any created account. It does not reset all Windows settings to factory defaults. Apply is not transactional; on partial failure use the recorded rollback. Repeated Apply operations create successive snapshots: review their order and restore newest first. Review concurrent policy changes before restoring an older snapshot.

## Local files and evidence

Default output: `C:\ProgramData\SmartThinClient\Shell`; if unwritable, fallback is `%LOCALAPPDATA%\SmartThinClient\Shell`, then `%TEMP%\SmartThinClient\Shell`. Audit/Preview emit JSON and optionally CSV under `Output/`, timestamped logs under `Logs/`; Apply saves JSON under `Rollback/` before changes. Evidence contains computer/user names and paths; rollback and generated scripts also contain local configuration/URLs. Treat them as sensitive operational data, not anonymized support bundles.

Offline regression command (no endpoint mutation):

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\Test-SmartThinClient.Regression.ps1
```
