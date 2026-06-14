# Smart DeviceRegistration Tool

Standalone local tool for Hybrid Join and Microsoft Entra device registration diagnostics.

The tool opens the local WPF GUI by default. It has two operating modes:

- `User`: default diagnostic-only mode for an end user or helpdesk handoff.
- `Admin`: support mode that unlocks guarded repair options.

In the GUI, `Run diagnostic` only runs diagnostics. In Admin mode, `Run repair` appears as a separate action and is enabled only when a guarded repair action such as `dsregcmd /leave` or `Automatic-Device-Join` is selected.
`Refresh Azure AD PRT` is a separate user-context action that can be run without local admin rights; it executes `dsregcmd /refreshprt` and then refreshes the diagnostic view.

Use `-Cli` when an operator wants command-line mode.

## Script

- `SmartM365-DeviceRegistration-Tool.ps1`: checks local AD domain join, domain controller reachability, Intune enrollment signals, and `dsregcmd /status`.
- `SmartM365-DeviceRegistration-Tool.config.template.json`: safe configuration template for retry and output settings.
- `SmartM365-DeviceRegistration-Tool.strings.psd1`: GUI language catalog.
- `Start-SmartM365-DeviceRegistration-Tool-User.cmd`: opens the GUI in User mode.
- `Start-SmartM365-DeviceRegistration-Tool-Admin.cmd`: opens the GUI in Admin mode. The PowerShell script requests UAC elevation automatically when needed.
- `Start-SmartM365-DeviceRegistration-Tool-CLI-Export.cmd`: silently starts CLI User-mode diagnostics and automatically exports JSON output plus a support bundle.

## Configuration

The GUI does not expose runtime path and retry settings. Put local overrides in:

```text
SmartM365-DeviceRegistration-Tool.config.json
```

next to the script, or pass a custom path with `-ConfigPath`.
If `SmartM365-DeviceRegistration-Tool.config.json` is not present, the tool uses the committed template values and then falls back to built-in defaults if the template is also missing.

Use `SmartM365-DeviceRegistration-Tool.config.template.json` as the committed model:

```json
{
  "RetryCount": 0,
  "RetrySleepMinutes": 5,
  "OutputRoot": "C:\\ProgramData\\SmartM365\\DeviceRegistrationTool",
  "LogRetentionCount": 10,
  "DeviceProfile": "Auto",
  "RequireDomainConnectivity": false,
  "SupportEmail": "",
  "SupportEmailSendMode": "Draft",
  "LogoPath": "workplacecloudhub-v2.png",
  "DefaultLanguage": "auto",
  "ForceLanguage": "",
  "LanguageCatalogPath": ""
}
```

CLI parameters override JSON values when provided.

`RequireDomainConnectivity` controls whether AD domain connectivity is mandatory:

- `false`: skip AD domain controller validation. This is suitable for Intune-only devices.
- `true`: require AD domain join, a reachable domain controller, and a current AD domain user session before continuing Hybrid Join checks or repair actions. If any requirement is not met, the tool stops and does not run repair actions.

`DeviceProfile` can be:

- `Auto`: use `RequireDomainConnectivity` as configured.
- `IntuneOnly`: skip AD domain controller validation.
- `HybridJoin`: require AD domain join, domain controller reachability, and an AD domain user session.

`LogRetentionCount` controls local file rotation. The tool keeps the latest 10 files by default for each generated artifact type:

- run logs in `Logs\`
- PowerShell transcripts in `Transcripts\`
- `dsregcmd /status` snapshots in `Output\`
- `dsregcmd /refreshprt` outputs in `Output\`
- `dsregcmd /leave` outputs in `Output\`
- event log exports, support summaries, JSON results, and support bundles

Set `LogRetentionCount` to `0` to disable rotation.

`SupportEmail` is the support mailbox used when emailing support summaries and support bundles.
`SupportEmailSendMode` controls the Outlook behavior:

- `Draft`: create and display an Outlook email draft with the bundle attached.
- `Send`: send the email directly with the current Outlook profile.

If Outlook is not available, the tool creates an `.eml` draft file next to the bundle.

`LogoPath` controls the GUI window icon and header logo:

- use an absolute path for a shared/company logo;
- use a relative path to resolve it from the JSON configuration folder;
- leave the default `workplacecloudhub-v2.png` to use the WorkplaceCloudHub logo next to the script.

If the default PNG is missing, the GUI keeps the built-in text logo placeholder.

`DefaultLanguage` controls the GUI language:

- `auto`: use the current Windows UI culture.
- supported languages: `en`, `fr`, `de`, `es`, `nl`, `it`, `pt`, `pl`, `ar`, `tr`, `sv`, `da`, `nb`, `fi`, `ro`, `hu`, `ja`, `ko`, `zh-Hans`, `uk`.
- common regional cultures are mapped automatically, for example `fr-FR` to `fr`, `zh-CN` to `zh-Hans`, and `no` to `nb`.

`ForceLanguage` can force the GUI language for deployment. Leave it empty to use `DefaultLanguage`.

`LanguageCatalogPath` can point to a custom PowerShell data file with the same structure as `SmartM365-DeviceRegistration-Tool.strings.psd1`.

## Output

By default, logs and outputs are written under:

```text
C:\ProgramData\SmartM365\DeviceRegistrationTool
```

If the current user cannot write there, the tool falls back to:

```text
%LOCALAPPDATA%\SmartM365\DeviceRegistrationTool
```

The tool creates:

- `Logs\`: run logs.
- `Output\`: `dsregcmd /status` snapshots and `dsregcmd /leave` output when used.
- `Transcripts\`: PowerShell transcripts unless `-NoTranscript` is used.
- `SmartM365_DeviceRegistration_<ComputerName>.csv`: local CSV run summary.
- `SmartM365_DeviceRegistration_SupportBundle_<ComputerName>_<RunId>.zip`: optional support bundle when requested.

CLI support outputs:

```powershell
.\Devices\DeviceRegistrationTool\SmartM365-DeviceRegistration-Tool.ps1 -Cli -JsonOutput
.\Devices\DeviceRegistrationTool\SmartM365-DeviceRegistration-Tool.ps1 -Cli -SupportBundle
.\Devices\DeviceRegistrationTool\Start-SmartM365-DeviceRegistration-Tool-CLI-Export.cmd
```

The support bundle includes the current run log, CSV summary, `dsregcmd` snapshot, support summary, and recent device registration / MDM event log exports when available.
When `SupportEmail` is configured, `-SupportBundle` also prepares or sends the email according to `SupportEmailSendMode`.

## Policy Checks

The diagnostic checks local policy signals commonly written by GPO or MDM policy:

- `HKLM\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\MDM`
  - `AutoEnrollMDM`
  - `UseAADCredentialType`
- `HKLM\SOFTWARE\Policies\Microsoft\Windows\WorkplaceJoin`
  - `autoWorkplaceJoin`
- Scheduled task:
  - `\Microsoft\Windows\Workplace Join\Automatic-Device-Join`

These checks are available in User and Admin modes because they are read-only.

The diagnostic also records the current Windows identity, whether it is a local administrator, and whether the current PowerShell process is elevated. This helps support decide whether an Admin-mode repair can run locally or needs elevation.

## Intune Enrollment Status

The GUI enrollment banner is based on local Intune enrollment evidence, not only on AD domain join or `dsregcmd` health.
For example, a device can show `NOT_DOMAIN_JOINED` and still be locally enrolled in Intune.

The diagnostic separates strong and weak Intune enrollment evidence.
The GUI banner is green only when strong local enrollment evidence is present, such as `ProviderID=MS DM Server` or an enrollment discovery URL.
Weak signals such as EnterpriseMgmt tasks or the Intune Management Extension service are recorded as informational evidence and can indicate a stale or partial enrollment.

The MDM auto-enrollment policy is always checked and shown in a separate banner.
When Intune enrollment is already detected, missing or disabled auto-enrollment policy is informational only and does not create an error or warning.

- Green: `AutoEnrollMDM=1`, or Intune enrollment is already detected.
- Red: the MDM auto-enrollment policy is missing or disabled on a device that is not enrolled.

The GUI also separates the actionable summary from the full diagnostic text.
The GUI keeps `Run diagnostic` in a fixed action banner.
The diagnostic area uses tabs: `Summary` for enrollment banners, issues, warnings, informational findings, and support actions; `Diagnostic output` for the complete support detail; and `dsregcmd` for the raw `dsregcmd /status` snapshot.
Admin-only settings are kept in an `Options` tab, which is hidden in User mode.
`Open latest log folder` is available only in Admin mode.
The GUI also has actions to copy/email a support summary and create/email a support bundle after a diagnostic run.
If `SupportEmail` is configured, `Copy/email support summary` copies the summary to the clipboard and opens or sends the support email according to `SupportEmailSendMode`.

## Default GUI mode

Open the GUI:

```powershell
.\Devices\DeviceRegistrationTool\SmartM365-DeviceRegistration-Tool.ps1
```

Explicit GUI launch:

```powershell
.\Devices\DeviceRegistrationTool\SmartM365-DeviceRegistration-Tool.ps1 -Gui
```

Launch User mode with the CMD wrapper:

```cmd
.\Devices\DeviceRegistrationTool\Start-SmartM365-DeviceRegistration-Tool-User.cmd
```

Launch Admin mode:

```cmd
.\Devices\DeviceRegistrationTool\Start-SmartM365-DeviceRegistration-Tool-Admin.cmd
```

## CLI mode

Diagnostic-only run:

```powershell
.\Devices\DeviceRegistrationTool\SmartM365-DeviceRegistration-Tool.ps1 -Cli
```

Explicit user diagnostic run:

```powershell
.\Devices\DeviceRegistrationTool\SmartM365-DeviceRegistration-Tool.ps1 -Cli -Mode User
```

Trigger the Windows Automatic-Device-Join task without running `dsregcmd /leave`:

```powershell
.\Devices\DeviceRegistrationTool\SmartM365-DeviceRegistration-Tool.ps1 -Cli -Mode Admin -TriggerJoin
```

Trigger Intune MDM auto-enrollment when Hybrid Join is healthy and the local auto-enrollment policy is configured:

```powershell
.\Devices\DeviceRegistrationTool\SmartM365-DeviceRegistration-Tool.ps1 -Cli -Mode Admin -TriggerIntuneAutoEnrollment
```

Allow guarded repair for a disabled or deleted Entra device object:

```powershell
.\Devices\DeviceRegistrationTool\SmartM365-DeviceRegistration-Tool.ps1 -Cli -Mode Admin -RepairDisabledDeletedDevice
```

Allow the broader Hybrid Join repair guard used by `SmartM365-Invoke-IntuneHybridJoinRepair.ps1`:

```powershell
.\Devices\DeviceRegistrationTool\SmartM365-DeviceRegistration-Tool.ps1 -Cli -Mode Admin -AllowDsregLeave
```

Remove stale local MDM enrollment traces after diagnostics confirm they are not a live Intune enrollment:

```powershell
.\Devices\DeviceRegistrationTool\SmartM365-DeviceRegistration-Tool.ps1 -Cli -Mode Admin -AllowRemoveStaleIntuneEnrollment
.\Devices\DeviceRegistrationTool\SmartM365-DeviceRegistration-Tool.ps1 -Cli -Mode Admin -AllowRemoveNonIntuneMdmEnrollment
```

Preview the recommended action without running repair commands:

```powershell
.\Devices\DeviceRegistrationTool\SmartM365-DeviceRegistration-Tool.ps1 -Cli -Mode Admin -AuditOnly
```

Add bounded post-action retries:

```powershell
.\Devices\DeviceRegistrationTool\SmartM365-DeviceRegistration-Tool.ps1 -Cli -Mode Admin -RepairDisabledDeletedDevice -RetryCount 2 -RetrySleepMinutes 10
```

Admin repair actions require an elevated PowerShell process.
When `-Mode Admin` is used from a non-elevated process, the tool automatically asks for UAC elevation and relaunches itself.

## Repair guardrails

`SmartM365-Invoke-IntuneHybridJoinRepair.ps1` remains autonomous so a single file can still be pushed by LOT, PsExec, Intune, or GPO.
DeviceRegistrationTool carries the same local diagnostic and guarded repair logic where it fits the GUI/CLI support workflow.
When shared Hybrid Join or Intune enrollment behavior changes in either tool, synchronize the other tool in the same change.

`dsregcmd /leave` is only allowed when the device has `AzureAdJoined=YES`, `DeviceId`, and `TenantId`, and one of these guarded paths is selected:

- `-RepairDisabledDeletedDevice`: strict path for a failed `DeviceAuthStatus` that indicates the Entra device object is disabled or deleted.
- `-AllowDsregLeave`: broader Hybrid Join repair path for failed device authentication or `KeySignTest=FAILED`.

MDM cleanup is opt-in and scoped:

- `-AllowRemoveStaleIntuneEnrollment`: removes stale local Intune enrollment traces only when diagnostics did not confirm a live Intune enrollment.
- `-AllowRemoveNonIntuneMdmEnrollment`: removes non-Intune MDM enrollment traces.
- `-TriggerIntuneAutoEnrollment`: runs `deviceenroller.exe /c /AutoEnrollMDM` only after Hybrid Join is healthy and the local MDM auto-enrollment policy is enabled.
- `-AuditOnly`: reports the next action without executing repair commands.

When local Intune enrollment is detected, repair actions are skipped by default. Use `-AllowIntuneEnrolledAction` only when the operator intentionally wants to override that guardrail.

## Exit Codes

```text
0 = Healthy / success
1 = Error
2 = Not AD domain joined
3 = Attention required / no repair performed / join still pending
4 = Domain controller not reachable
```
