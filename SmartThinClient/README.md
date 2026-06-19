# Smart ThinClient Shell

Smart ThinClient Shell is a local Windows endpoint tool that prepares a standard Windows PC to behave like a controlled thin-client workspace for Citrix or Azure Virtual Desktop access.

The default behavior is conservative: `Audit` and `Preview` are read-only. `Apply` and `Restore` are implemented but require administrator rights, local configuration gates, confirmation phrases, rollback storage, logs, and JSON/CSV evidence.

## Scope

Use this tool when the work is centered on the Windows endpoint:

- verifying whether Citrix Workspace app or Windows App / AVD client components are present;
- preparing a controlled web-first shell or launcher experience for virtual desktops;
- documenting the endpoint state before a kiosk or shell-lockdown rollout;
- keeping rollback and operator guardrails visible before any local change is made.

Citrix platform inventory stays in `SmartCitrix/`. Azure Virtual Desktop platform inventory stays in `SmartAzureVirtualDesktop/`. This project belongs at repository root because it changes or audits the Windows endpoint experience across Citrix, AVD, and web workspace scenarios.

## Files

- `SmartThinClient-Shell.ps1`: local GUI/CLI tool.
- `SmartThinClient-Shell.config.template.json`: safe configuration template.
- `Profiles/Citrix.profile.json.template`: Citrix endpoint profile template.
- `Profiles/AVD.profile.json.template`: Azure Virtual Desktop endpoint profile template.
- `Profiles/WebOnly.profile.json.template`: browser-only endpoint profile template.
- `Profiles/Hybrid.profile.json.template`: startup provider-choice profile template.
- `Start-SmartThinClient-Shell-GUI.cmd`: opens the local GUI.
- `Start-SmartThinClient-Shell-CLI-Preview.cmd`: runs a read-only CLI audit.
- `Start-SmartThinClient-Shell-LaunchOnly.cmd`: opens the thin-client launcher without installing or changing Windows shell settings.
- `Start-SmartThinClient-Shell-CLI-Apply.cmd`: guarded apply launcher.
- `Start-SmartThinClient-Shell-CLI-Restore.cmd`: guarded restore launcher.

## Modes

```powershell
.\SmartThinClient-Shell.ps1
.\SmartThinClient-Shell.ps1 -Cli
.\SmartThinClient-Shell.ps1 -Cli -Action Preview
.\SmartThinClient-Shell.ps1 -Cli -Action Launch -Profile Hybrid
.\SmartThinClient-Shell.ps1 -Cli -Profile Citrix
.\SmartThinClient-Shell.ps1 -Cli -Profile AVD
.\SmartThinClient-Shell.ps1 -Cli -Profile WebOnly
.\SmartThinClient-Shell.ps1 -Cli -Profile Hybrid
.\SmartThinClient-Shell.ps1 -ValidateOnly
```

`Audit` is the default action and is read-only. `Preview` is also read-only and emits the same local evidence with `PreviewOnly` status.

`Launch` is a run-only mode. It creates a temporary launcher from the detected Citrix, AVD, WebOnly, or Hybrid profile and opens it immediately. It does not install anything, does not configure auto-launch, and does not change Windows shell, kiosk, registry lockdown, or Assigned Access settings.

The preferred launch experience is web-first. When `PreferredAccessMode` is `Web` and `UseWebShell` is `true`, the generated launcher opens a controlled WPF thin-client shell with an embedded web area for Citrix Web, AVD Web, or WebOnly. It also exposes limited operator buttons such as refresh, back, open external browser, limited Windows settings, sign out, restart, and shutdown according to local JSON settings.

`Apply` creates rollback data first, then can:

- create the local launcher script;
- configure auto-launch through a machine Run key;
- optionally apply basic or strict shell limitations;
- optionally configure Assigned Access when the OS and app model support it;
- optionally configure Shell Launcher when the OS exposes the Shell Launcher WMI/CIM class.

`Restore` uses the latest rollback file, or the file passed through `-RollbackPath`, to restore saved registry values, remove the SmartThinClient auto-launch value, clear Assigned Access when available, and remove the Shell Launcher custom shell for the recorded target user SID.

## Configuration

Copy the template to a local JSON file when endpoint-specific defaults are needed:

```text
SmartThinClient-Shell.config.json
```

When the GUI or CLI starts, the tool creates the local JSON automatically from the template if it does not already exist. Templates stay as reference files; edit the generated runtime JSON instead.

Local runtime JSON files must stay out of Git.

Important settings:

- `DefaultProfile`: `Auto`, `Citrix`, `AVD`, `WebOnly`, or `Hybrid`.
- `OutputRoot`: default local output path.
- `ShellMode`: planned target mode, for example `Launcher`, `AssignedAccess`, or `ShellLauncher`.
- `AllowApply`: must be `true` before future apply actions can run.
- `AllowRestore`: must be `true` before future restore actions can run.
- `TargetUserMode`: `ExistingUser` or `DedicatedUser`.
- `TargetUserName`: existing or dedicated local user targeted by kiosk features.
- `CreateDedicatedLocalUser`: creates the dedicated local user only when `-DedicatedUserPassword` is supplied at runtime.
- `EnableShellLimitations`: applies optional Explorer/CMD limitations.
- `ShellRestrictionLevel`: `None`, `Basic`, or `Strict`.
- `EnableAssignedAccess`: uses `Set-AssignedAccess` when available.
- `AssignedAccessAppUserModelId`: required for Assigned Access.
- `EnableShellLauncher`: uses `WESL_UserSetting` when available.
- `PreferredAccessMode`: `Web` by default; use `Native` only when local Citrix/AVD clients should be preferred.
- `UseWebShell`: opens the controlled thin-client web shell instead of a plain external browser.
- `CitrixWebUrl`: local-only Citrix Workspace, StoreFront, or Gateway web URL.
- `AvdWebUrl`: local-only Azure Virtual Desktop or Windows App web URL.
- `WebShellAllowExternalBrowser`: allows opening the current portal in the configured browser for compatibility.
- `WebShellAllowPowerControls`: shows sign out, restart, and shutdown buttons.
- `WebShellAllowLimitedSettings`: shows a limited settings menu.
- `WebShellAllowedSettingsPages`: `ms-settings:` pages exposed through the limited settings menu.
- `HybridSelectionAtStartup`: shows a provider choice at launcher startup for `Hybrid`.

Apply example:

```powershell
.\SmartThinClient-Shell.ps1 -Cli -Action Apply -Profile Hybrid -ConfirmApply "APPLY SMARTTHINCLIENT"
```

Restore example:

```powershell
.\SmartThinClient-Shell.ps1 -Cli -Action Restore -ConfirmRestore "RESTORE WINDOWS SHELL"
```

Dedicated user creation does not store passwords in JSON:

```powershell
$password = Read-Host "Dedicated user password" -AsSecureString
.\SmartThinClient-Shell.ps1 -Cli -Action Apply -TargetUserMode DedicatedUser -TargetUserName "ThinClientUser" -DedicatedUserPassword $password -ConfirmApply "APPLY SMARTTHINCLIENT"
```

## Output

By default, logs and audit files are written under:

```text
C:\ProgramData\SmartThinClient\Shell
```

If that path is not writable, the tool falls back to:

```text
%LOCALAPPDATA%\SmartThinClient\Shell
%TEMP%\SmartThinClient\Shell
```

Every log line begins with a timestamp.

Each run emits:

- JSON evidence under `Output\`;
- CSV evidence under `Output\`;
- timestamped logs under `Logs\`;
- rollback JSON under `Rollback\` before any apply action changes local state.

## Guardrails

The tool must not remove administrator access, install clients, or store real workspace URLs in committed files.

Shell lockdown features are opt-in. Keep `EnableShellLimitations`, `EnableAssignedAccess`, and `EnableShellLauncher` disabled until the target user, OS edition, rollback file, and support process are validated.
