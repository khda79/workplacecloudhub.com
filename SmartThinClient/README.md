# Smart ThinClient Shell

Local Windows endpoint tooling for Citrix, Azure Virtual Desktop (AVD), and web workspaces. Download the portable [Smart ThinClient 1.1.4 release](https://github.com/khda79/workplacecloudhub.com/releases/tag/smart-thinclient-v1.1.4), or use the [source folder](https://github.com/khda79/workplacecloudhub.com/tree/main/SmartThinClient). The release includes the pinned WebView2 SDK DLLs and their license; the Evergreen Runtime remains a prerequisite. No MSI installer, PowerShell Gallery module, Intune package or automatic updater is provided.

## Prerequisites and validation limits

Use **64-bit Windows PowerShell 5.1** on Windows, with .NET Framework/WPF and an interactive desktop for the GUI and launcher. CMD wrappers invoke `powershell.exe`; launchers use STA. PowerShell 7 regression tests cover isolated functions; they do not certify the complete application on PowerShell 7.

Audit and Preview normally do not need elevation. Apply and Restore require an elevated administrator session. Client software, browser, portal URLs, subscriptions, authentication, connectivity and endpoint policy must be prepared by the operator. No tenant, Citrix delivery group or AVD host pool is provisioned by this tool. A detected executable, package, cmdlet or CIM class is a local signal, not proof of a usable remote session or Windows edition entitlement.

The embedded shell and visual Preview use **Microsoft Edge WebView2**. A 64-bit Evergreen WebView2 Runtime must already be installed; the application does not install a runtime or change browser emulation/policy settings. Run `SmartThinClient-PrepareWebView2.ps1` from a trusted application folder to prepare the pinned, hash-verified Microsoft SDK dependencies, or supply its verified NuGet package with `-PackagePath` for offline preparation. The three application-local DLLs are verified against `SmartThinClient-WebView2.psd1`; browser data stays in the current user Local AppData. Native clients, new windows, downloads and permission requests are blocked in Preview. Citrix/AVD authentication and session compatibility require separate environment validation. The window, provider buttons and optional browser kiosk switches are not a security boundary or an enforced URL allowlist.

References: [WebView2 for WPF](https://learn.microsoft.com/en-us/microsoft-edge/webview2/get-started/wpf), [Windows App connection requirements](https://learn.microsoft.com/en-us/windows-app/get-started-connect-devices-desktops-apps), [Shell Launcher configuration](https://learn.microsoft.com/en-us/windows/configuration/shell-launcher/wesl-usersetting).

Real Citrix/AVD authentication and remote sessions, complete Assigned Access/Shell Launcher recovery cycles and physical device-setting changes are not certified by this package. See [VALIDATION.md](VALIDATION.md) for isolated regression evidence and validation boundaries; qualify the selected mode in your own environment.

## Installation and updates

Clone/download the repository and retain the complete `SmartThinClient/` folder, including `Assets/`, `Profiles/`, splash script and CMD wrappers. Start in a user-writable working copy for Audit/Preview. Review the source and its signature before running it:

```powershell
Get-AuthenticodeSignature .\SmartThinClient-Shell.ps1
.\SmartThinClient-Shell.ps1 -Cli -Action Preview
```

The repository uses a self-signed WorkplaceCloudHub code-signing certificate. Certificate trust is an operator decision; a valid signature does not certify functionality. CMD wrappers and locally generated launchers explicitly use process-scoped `ExecutionPolicy Bypass`. They do not change the machine execution policy. Generated launcher files contain local configuration and are not Authenticode-signed; package-signature validation does not extend to those files.

For signature verification, use a Git clone/checkout that honors the repository's `.gitattributes` (`*.ps1 text eol=crlf`). In the publication read-back, GitHub Raw served LF line endings and Windows reported `NotSigned`; restoring the repository's CRLF checkout produced `Valid` for the same source. A raw-file download or browser copy is therefore not equivalent to the signed Windows checkout. Do not bypass a failed signature check: verify the checkout and certificate trust first.

For updates, keep a separate backup of local JSON, output and rollback files, compare the new templates, and rerun Preview. There is no automatic migration or update check. Existing local profiles remain explicit overrides: set inherited fields to `"__USE_GLOBAL__"` when adopting the corrected templates. Never overwrite local workspace URLs with a blank template.

The **Smart ThinClient Control Panel** provides Audit, Preview, Apply, Restore and Open files in a compact, resizable window with shared modern buttons and a styled profile selector. Audit shows colored checks for Windows, browser, configured provider access and the target account, plus a global **Checks passed / Review required / Blocked** status. The target check uses the local account lookup or, for a matching existing user, the current authenticated Windows token. A domain user is shown as **Domain account detected**; an unconfirmed account stays **Not verified**. This diagnostic identity does not replace the local-account requirements for dedicated users, Strict restrictions or kiosk operations. Evidence keeps the diagnostic resolved SID separate from the operational local SID; a domain account's enabled state is not inferred. Disabled Apply/Restore gates are informational. These are local signals; configured URLs do not prove authentication or remote-session availability. Selecting another profile hides the previous summary until Audit runs again.

When a configuration gate refuses an action, its GUI status is **Apply blocked by configuration** or **Restore blocked by configuration**. Other failures retain their detailed error and are not labelled as a configuration block. Enabling `AllowApply`/`AllowRestore` does not bypass administrator, exact typed-confirmation, prerequisite or rollback checks.

Provider checks separate **Citrix Web** from **Citrix Workspace**, and **AVD Web** from **Windows App** and **Remote Desktop (legacy)**. A green Web card means a configured URL and a detected browser; native client detection has its own card. A client not detected in Web mode is informational, while a missing required native client needs review or blocks readiness when there is no Web fallback. Windows App is detected for the current Windows user; an aggregate AVD-client signal alone never proves Windows App is present. A detected legacy Remote Desktop client is a warning. [Microsoft recommends migrating to Windows App](https://learn.microsoft.com/en-us/previous-versions/remote-desktop-client/connect-windows-cloud-services?tabs=windows-msrdc-msi); RemoteApp denotes published individual applications, not a separate replacement client.

When Citrix Workspace or Windows App is not detected, its Audit card offers **Download Citrix Workspace** or **Download Windows App**. Clicking opens the official [Citrix Windows download page](https://www.citrix.com/downloads/workspace-app/windows/) or [Windows App Microsoft Store page](https://apps.microsoft.com/detail/9n1f85v9t8bn) in the default browser. Audit itself opens nothing and no installer is downloaded or run by the tool. An existing legacy Remote Desktop installation still offers Windows App when the current app is absent. Detected clients do not show the download prompt.

The GUI Preview button opens a separate visual shell using the same layout as the workspace launcher. It loads the configured HTTP/HTTPS page with Back/Refresh. Citrix, AVD and WebOnly show only their own provider button; Hybrid shows all three unless provider buttons are disabled in configuration. Settings is available in Preview when `WebShellAllowLimitedSettings` is true, with the same configured `WebShellAllowedSettingsPages` as the operational shell. These open **live Windows settings**, and changes made there affect the device. Only `ms-settings:` destinations are accepted. Power actions, external-browser launches and native clients remain disabled in Preview. Missing or invalid web URLs show an explicit unconfigured state; a configured WebOnly fallback is labelled. The technical change plan remains in Preview JSON evidence and in the Apply review dialog. `-Gui -Action Preview` opens this panel; CLI Preview remains available.

Apply first shows a fresh preview and asks for the exact configured confirmation phrase. The default GUI Restore action opens **Restaurer le bureau Windows**, with a readable summary and the exact configured restore phrase. It discovers the rollback history for the current ThinClient custom shell and follows recorded earlier startup commands until Windows Explorer is configured without ThinClient automatic startup. This also handles migration from the previous user-profile Launcher. It validates the complete sequence before writing, pins the reviewed state and file hashes, locks every reviewed backup, and preserves separate step evidence. Missing, ambiguous, redirected, foreign or changed history is refused; a partial failure reports the completed steps and never claims Windows has been restored. Cancel makes no changes. If Explorer is already configured without ThinClient startup, the panel reports that no restore is needed. A successful restore offers sign-out now or later; it does not close the current session automatically.

Right-click **Restore** for **Restore a specific backup (advanced)**, retaining the explicit file picker, hash review and single-snapshot behavior. This advanced path is also required for Assigned Access/Shell Launcher kiosk histories, unsupported custom shells or a manually reviewed recovery. CLI Restore retains its existing explicit `-RollbackPath` behavior. Restoring one backup can re-enable an earlier ThinClient configuration; it is distinct from the default full return to the Windows desktop.

Both actions preserve administrator checks, global/profile configuration gates, preflight validation and evidence. Apply saves rollback before writes. The GUI always requires a typed phrase, even if CLI phrase confirmation is optional in local configuration. When a configured dedicated user must be created, the confirmation dialog accepts its password through a PasswordBox; it is passed as a SecureString and is not logged. A non-elevated panel cannot perform Apply/Restore; there is no automatic elevation or automatic gate change.

`AllowApply` and `AllowRestore` default to `false`. The panel reports these blocks when clicked. Visual Preview remains available while they are false, without invoking Apply, Restore or the workspace Launch command, and without creating an installed launcher. Web navigation inside the preview can contact the configured site; it does not establish that authentication or remote sessions work. Start with `Start-SmartThinClient-Shell-GUI.cmd` to run without a console window; this wrapper returns after starting the GUI. Direct PowerShell shortcuts should include `-WindowStyle Hidden`. CLI wrappers keep their consoles.

## Modes

The Settings menu uses four visible buttons by default. Hide each independently in the runtime JSON:

```json
"WebShellSettingsButtons": {
  "Network": true,
  "Display": true,
  "Sound": true,
  "DateAndTime": true,
  "Bluetooth": false,
  "Printers": false
}
```

Set any value to `false` to hide its button, in both Preview and the operational shell. `WebShellAllowedSettingsPages` remains authoritative: visibility flags do not add omitted pages. Partial profile visibility overrides inherit unspecified global flags. Existing configurations without the new object keep their configured pages visible. `WebShellAllowLimitedSettings: false` hides the whole Settings entry. The dialog uses the owner's WorkplaceCloudHub icon.

Bluetooth and Printers are optional shortcuts to the live Windows Bluetooth and Printers & scanners pages. Set their flags to `true` to show them. New configurations include `ms-settings:bluetooth` and `ms-settings:printers` in `WebShellAllowedSettingsPages`, with both buttons hidden by default. For an existing configuration, also add those two URIs to its allowed list if wanted; existing custom lists are preserved. Showing a shortcut does not prove Bluetooth hardware or a printer is present, and opening these pages does not itself pair or install a device.

| Mode | Behavior |
| --- | --- |
| Audit (default) | Inspects the local OS, users, clients, browser and kiosk capabilities. Without `-Cli`, opens the audit GUI. |
| Preview | GUI: visual shell with power/external actions disabled; Settings can explicitly open configured live Windows settings pages. CLI: local inspection and change plan with `PreviewOnly` evidence. No automatic shell/policy/startup change or kiosk/session certification. |
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

Hybrid offers Citrix, AVD and WebOnly; closing the choice window cancels startup. With `HybridSelectionAtStartup: false`, it uses `HybridDefaultProvider`. This key takes precedence over the legacy `WebShellPreferredProvider` fallback. Hybrid resolves each provider's URL from its own local profile over the main/Hybrid configuration; blank or inherited profile values retain the global URL. Audit uses the same URL resolution as the launcher.

The GUI provides audit/preview and output access; Apply/Restore buttons explain the guarded CLI workflow. CMD Apply/Restore wrappers accept additional arguments; without a confirmation argument they remain blocked by default.

## Configuration

The main runtime file is `SmartThinClient-Shell.config.json`, initialized from `SmartThinClient-Shell.config.template.json`. Profile runtime files are initialized from `Profiles/*.profile.json.template`. Additional missing template keys are merged into these local files. All runtime JSON, logs and URLs stay out of Git. `-ConfigPath` selects an alternate existing main JSON; profiles are still read from this tool's `Profiles/` directory. `-OutputRoot` overrides the configured output root.

Main values override built-in defaults; non-empty profile values override the main configuration. The sentinel `"__USE_GLOBAL__"` inherits the main value. Boolean switches must be JSON `true`/`false`, never strings.

- `DefaultProfile`: Auto, Citrix, AVD, WebOnly or Hybrid. Auto chooses from detected local clients.
- `PreferredAccessMode`: Web or Native; `UseWebShell`: embedded WPF or external browser for Web mode.
- `CitrixWebUrl`, `AvdWebUrl`, `WebOnlyUrl`: local portal URLs. Legacy aliases include WorkspaceUrl/StoreUrl for Citrix, FeedUrl for AVD, and WebShellHomeUrl/FallbackWebUrl for web fallback. A subscription feed URL is not necessarily a browser portal.
- New WebOnly configurations default to `https://workplacecloudhub.com/`. Existing custom URLs and explicit blanks are preserved; this value does not open a site during Audit or CLI Preview. Clicking GUI Preview loads the configured page in its visual shell.
- New main configurations default to `https://workplacecloudhub.com/demo/smartthinclient/citrix/` and `https://workplacecloudhub.com/demo/smartthinclient/avd/`. These are demonstration destinations, not tenant sign-in URLs or proof of a working remote session. Their availability must be checked separately. Provider profile URL fields stay empty to inherit the main configuration. Custom primary URLs remain authoritative; existing custom WorkspaceUrl/StoreUrl/FeedUrl aliases take precedence over a demo default. Existing runtime values, including explicit blanks, are retained by the startup template-key merge.
- `CitrixWorkspacePath`, `AvdClientPath`, `BrowserPath`: explicit paths, otherwise detection searches common installation locations.
- Native Citrix starts the detected executable without connection arguments. Detection of CDViewer/wfica32 alone does not guarantee a standalone workspace launch.
- Native AVD first tries classic `msrdcw.exe`, then a configured `WindowsAppLaunchUri`, then web fallback. Detecting the Windows App package for the current user alone does not supply a launch URI or validate another user's installation.
- `BrowserKioskMode`: passes browser kiosk arguments when using external web launch. Validate the chosen browser; this is not Windows Assigned Access.
- `WebShellAllowExternalBrowser`, `WebShellAllowPowerControls`, `WebShellAllowLimitedSettings`, `WebShellAllowedSettingsPages`: optional UI controls. Their defaults expose sign-out/restart/shutdown and selected Windows settings; turn them off for a restricted review.
- `AllowApply`, `AllowRestore`: both false by default. `RequireConfirmationPhrase` defaults true.
- `WriteCsvEvidence`: true by default. False keeps JSON evidence and logs but skips CSV creation, replacement and removal, including an existing CSV at the same path.

`LogRetentionCount`, `RollbackRetentionCount`, `PreferredWorkspaceUrl`, `DefaultLanguage`, `ForceLanguage`, and profile metadata such as `LaunchMode`, `AllowedProviders`, `ExpectedExecutableNames` are not enforced as operational controls in this implementation. Logs/rollback/LaunchOnly files need operator-managed retention. The GUI is English.

## Apply and Restore scope

The workspace toolbar keeps provider/navigation buttons on the left, the workspace name centered, and Settings/Admin/sign-out/restart/shutdown on the right. Both button groups wrap within their own area on narrower windows. The same layout is used in Preview. Set `"WebShellTitle": "Customer workspace"` in the local main configuration (or override it in a profile) to personalize the centered name. Preview reads the effective configuration; an installed launcher uses the name saved when that launcher is generated. Long names are trimmed visually with their full text available in the tooltip.

After a successful Apply or Restore that needs another sign-in, the result dialog offers **Sign out now** and **Later**. The result reminds the user to save open work; clicking **Sign out now** requests current-session sign-out directly without another confirmation. Later remains the default. No forced application closure is requested. Audit, Preview, failed operations and a Launcher Apply with automatic startup disabled do not offer this action.

When Explorer is replaced, Settings uses the application-local `SmartThinClient-DeviceSettings.ps1` helper and modern built-in panels instead of relying on Windows Settings URI activation. It provides adapter/IP/DNS information and connection to saved Wi-Fi profiles, primary-display mode selection with a 15-second revert safeguard, default playback volume/mute, time zone and administrator-only clock adjustment, Bluetooth pairing when a radio exists, and installed-printer selection/default/properties. Network provisioning and printer installation remain administrator-managed. These controls are a focused subset of device settings, not a replacement for every Windows Settings page. RDP controls display resolution; absent audio/Bluetooth/printers are reported rather than represented as working hardware. All changes require an explicit control action and respect the current Windows token, including in Preview. No Explorer process is started by these panels.

In desktop Launcher mode, existing Windows Settings links launch in the background and report activation errors without blocking the workspace. Native WebView2 callbacks stay in managed handlers with failure reporting; browser event errors must not escape through a PowerShell COM-event callback and terminate the primary shell.

`ShellMode: CustomShell` replaces Explorer at the next sign-in for all users through the **local Group Policy Custom User Interface** setting (User Configuration / System). It requires `EnableCustomShell: true`, a persistent embedded web shell, administrator rights and typed confirmation. Windows Pro supports this policy; it does not require the separate Enterprise Shell Launcher feature. Higher-priority domain policy can override local policy. This interface is not a security boundary and does not block other ways to start programs.

CustomShell uses `MachineRoot`, a dedicated directory beneath ProgramData, for the shared launcher and rollback. SYSTEM and administrators can modify it; authenticated users can read/execute the launcher, while rollback remains administrator-only. A machine startup command must never depend on an operator's user profile. Reports and logs can remain in the operator's output directory. In this mode Apply removes its previous Run-key startup only after saving rollback and configuring the custom interface. It does not sign out or close the current desktop. Sign out and sign in after reviewing the successful result to activate it.

The shell exposes **Admin** only to administrator accounts, including accounts with a filtered UAC token. It opens the protected control panel and requests elevation when required. Standard users have no Admin entry, and the launch handler checks administrator membership again. Preview offers the same entry; it can return directly to an existing elevated control panel. The primary custom shell keeps its window open until a selected system action closes the session; its toolbar remains available if the embedded browser fails to initialize.

Apply defaults to the exact phrase `APPLY`; an explicitly configured custom phrase remains supported. Apply and Restore share the modern confirmation dialog. After success, a separate result dialog explains the outcome, next sign-in and rollback location. **Open files** opens the operational files directory. Each GUI action has a unique evidence ID so a later Audit or Preview cannot overwrite an earlier Apply result or rollback.

`AutoLaunchMode: RunKey` writes **HKLM** `SOFTWARE\Microsoft\Windows\CurrentVersion\Run\SmartThinClientShell`, affecting sign-in for all users. `AutoLaunchMode: None` writes no auto-launch value; it does not remove a previous installation. `TargetUserName` does not scope this machine Run key or machine policies. Use an administrator-controlled persistent output directory and verify read access for the target user; do not point a machine startup command at a directory writable by untrusted users.

`EnableShellLimitations: true` with Basic writes machine Explorer policies NoRun, NoControlPanel and NoViewContextMenu. Strict additionally writes `DisableCMD` under `HKEY_USERS/<TargetUserSid>/Software/Policies/Microsoft/Windows/System`. Strict requires an existing local target user whose registry hive is loaded; it refuses Apply before configuration writes otherwise. It does not load offline hives or create a new dedicated user for Strict. None writes no restrictions. The Explorer settings remain machine-wide, while the CMD restriction targets the selected user's hive. Policy/OS-dependent effects remain untested live; Group Policy/MDM may override them. The tool does not change autologon, Winlogon Shell/Userinit, or install/enable Windows features.

Assigned Access and Shell Launcher cannot be requested together. Both require an existing enabled local non-administrator user and the relevant local APIs. Kiosk preflight rejects missing prerequisites before applying the Run key or policies. Assigned Access requires a supplied `AssignedAccessAppUserModelId` and no existing Assigned Access configuration. Shell Launcher requires its feature already enabled, no custom shell for the target SID, and a persistent embedded web shell; its Windows edition/licensing and portal suitability still require a lab review. Existing kiosk configurations are deliberately not overwritten because this tool does not capture their complete prior state.

For Launcher mode, optional DedicatedUser creation requires `CreateDedicatedLocalUser: true` plus `-DedicatedUserPassword` as a SecureString. It uses the built-in Users group SID for localized Windows. Passwords are never written to JSON. A created user is **not deleted by Restore**, and no automatic logon or client provisioning is configured. Prepare kiosk users separately.

```powershell
# Review local JSON and Preview before these guarded, write-capable commands.
.\SmartThinClient-Shell.ps1 -Cli -Action Apply -Profile Hybrid -ConfirmApply "APPLY"
.\SmartThinClient-Shell.ps1 -Cli -Action Restore -RollbackPath "C:\ProgramData\SmartThinClient\Shell\Rollback\<reviewed-file>.json" -ConfirmRestore "RESTORE"
```

Restore accepts schema 2 rollback files from the same computer and original OutputRoot. Legacy files require manual review: they included unrelated Winlogon values and lacked launcher history. Only registry values selected for the Apply operation are captured. Restore clears newly configured Assigned Access or removes the newly configured custom shell, restores/removes the launcher as recorded, and retains evidence and any created account. It does not reset all Windows settings to factory defaults. Apply is not transactional; on partial failure use the recorded rollback. Repeated Apply operations create successive snapshots: review their order and restore newest first. Review concurrent policy changes before restoring an older snapshot.

## Local files and evidence

The control panel, operational shell and visual Preview show the product version at the lower right, independently of the configured customer workspace title. Restart and Shutdown use compact branded confirmations with Cancel as the default and a save-work reminder. Windows command failures remain visible in the shell status bar; these commands do not force-close applications.

CustomShell rollback uses schema 3 and includes the original local user-policy Shell value and shared launcher history. Restore checks the current managed policy before changing it and refuses to replace an unrelated newer shell. Existing schema 2 backups remain supported with their original OutputRoot; select the reviewed file explicitly when restoring an older Launcher installation. Restore restores the selected snapshot, which can contain a previous ThinClient installation; it is not a factory reset. The local policy write and a complete sign-in/Restore cycle require separate endpoint qualification.

A user-scoped CMD rollback snapshot must match the recorded target SID and its hive must be loaded before any registry restoration. Older schema 2 machine-scoped CMD snapshots remain supported.

Default output: `C:\ProgramData\SmartThinClient\Shell`; if unwritable, fallback is `%LOCALAPPDATA%\SmartThinClient\Shell`, then `%TEMP%\SmartThinClient\Shell`. Audit/Preview emit JSON and optionally CSV under `Output/`, timestamped logs under `Logs/`; Apply saves JSON under `Rollback/` before changes. Evidence contains computer/user names and paths; rollback and generated scripts also contain local configuration/URLs. Treat them as sensitive operational data, not anonymized support bundles.

Offline regression command (no endpoint mutation):

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Tests\Test-SmartThinClient.Regression.ps1
```
