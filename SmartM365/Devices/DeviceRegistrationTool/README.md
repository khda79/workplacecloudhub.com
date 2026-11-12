# Smart Device Registration Tool

## User view (1.6.2)

WinDC/MMP-C linked enrollments are recognized as Microsoft companion registrations when
ProviderID is `Microsoft Device Management`, EnrollmentType is 26 (or `0x1a`) and the HTTPS
discovery endpoint is exactly `discovery.dm.microsoft.com/EnrollmentConfiguration`.
They remain visible in technical evidence as `MicrosoftLinkedEnrollment` and an informational
message, without a foreign-MDM warning. This does not prove enrollment health, synchronization
or primary Intune enrollment. Missing/mismatched evidence and unrelated MDM warnings remain
distinct; no enrollment or task is modified. See [Microsoft WinDC enrollment](https://learn.microsoft.com/en-us/windows/client-management/declared-configuration-enrollment).

User mode opens a compact, read-only view: one overall message, work-account connection,
organization connection and locally detected device management. Warnings and incomplete
evidence produce a review message even when the underlying registration status is healthy.
Unknown evidence stays unconfirmed. The existing diagnostic results and exports are unchanged.

The primary actions are **Check my device** and **Copy report for support**. The latter only
copies the existing support summary to the local clipboard; it never opens or sends email.
**Details for support** is collapsed initially and retains Summary, Diagnostic output and
Device status, including all findings and existing explicit support export/email controls.
Opening the details temporarily replaces the compact result card, giving technical tabs
the full content area; collapsing them returns to the compact view.
User mode hides the legacy Admin PRT control and all repair/options controls. A separate
**Renew my work connection** action requires confirmation, as described below. Admin mode keeps
its technical layout and existing guarded actions. User labels support English and French
and fall back to English for other configured languages.

The local check remains read-only. Optional license verification and confirmed session connection
renewal are separate actions; neither runs automatically as part of the diagnostic.

Standalone local tool for Hybrid Join and Microsoft Entra device registration diagnostics.

The tool opens the local WPF GUI by default. It has two operating modes:

- `User`: default diagnostic-only mode for an end user or helpdesk handoff.
- `Admin`: support mode that unlocks guarded repair options.

In the GUI, `Run diagnostic` only runs diagnostics. In Admin mode, `Run repair` appears as a separate action and is enabled only when a guarded repair action such as `dsregcmd /leave` or `Automatic-Device-Join` is selected.
`Refresh Azure AD PRT` is a separate user-context action that can be run without local admin rights; it executes `dsregcmd /refreshprt` and then refreshes the diagnostic view.

Use `-Cli` when an operator wants command-line mode.

## User diagnostics

The User diagnostic includes read-only capabilities adapted from Smart Endpoint Migration
Toolkit's HybridJoin endpoint. It does not invoke the endpoint, worker, prerequisite tasks,
PsExec, WinRM, Microsoft Graph, enrollment, cleanup, policy changes or automatic PRT renewal.
Existing explicit Admin and PRT-refresh actions remain separate and unchanged.

- **Summary** separates device registration, current-user SSO/PRT and local Intune evidence.
  `HEALTHY` still describes registration, not Intune enrollment or cloud compliance.
- **Diagnostic output** includes PRT update/expiry dates, WAM/Hello/CloudTgt/OnPremTgt when
  available, attributed acquire/refresh failures, DNS/TCP443 probes, local MDM classification,
  readable policies and recent accessible AAD/registration/MDM events.
- **Device status** retains the raw `dsregcmd /status` snapshot.
- An old PRT error or an error belonging to another account is informational. Only a dated
  failure for the current account in the last 24 hours is treated as current PRT evidence.
  Expiration is evaluated only when the source includes an explicit timezone.
- Local MDM evidence distinguishes Windows internal providers, the SCCM management bridge,
  explicit Intune provider, other MDM and unconfirmed/incomplete Intune candidates. No trace
  is removed. Discovery URLs alone do not confirm enrollment.
- Refused or invisible registry/task/event evidence is **Unknown / not verifiable with this
  account**, never proof that a task, policy or enrollment is absent. Positive local evidence
  is not a cloud verification. Events may concern another user or the device and are not a
  substitute for checking current license assignments with support.
- Connectivity probes use at most six hosts, with 1.5-second DNS and TCP waits per host.
  They test user-context DNS/TCP443 only; they do not validate TLS, proxy, authentication or
  SYSTEM-context access. No DNS flush or configuration change is performed.

Structured evidence is saved as `Output/UserDiagnostics_<RunId>.json` in the effective output
folder; its path and the independent state fields are included in JSON/CSV exports. Keep
this evidence private: it can contain account identifiers and event messages. User guidance
recommends network/sign-in checks or contacting support, never an automatic enrollment trigger.
The additional Summary labels use English with French context labels where available.

The candidate needs both the main script and its local
`SmartM365-DeviceRegistration.UserDiagnostics.ps1` helper. The published 1.2.1 ZIP predates
this candidate. Offline regression command:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\SmartM365\Tests\Test-SmartM365DeviceRegistrationUserDiagnostics.ps1
```

## Download

Download the standalone Windows package from the
[SmartM365 Device Registration Tool v1.0 release](https://github.com/khda79/workplacecloudhub.com/releases/tag/device-registration-tool-v1.0).

After downloading:

1. Verify the ZIP against the published SHA-256 file.
2. Extract the complete ZIP to a local folder.
3. Run `Start-SmartM365-DeviceRegistration-Tool-User.cmd` for diagnostic-only mode.
4. Use `Start-SmartM365-DeviceRegistration-Tool-Admin.cmd` only when guarded repair actions are required.

The PowerShell scripts in the package are Authenticode-signed by `workplacecloudhub.com`.
The package does not include an installer or third-party binaries.

## Script

- `SmartM365-DeviceRegistration-Tool.ps1`: checks local AD domain join, domain controller reachability, Intune enrollment signals, and `dsregcmd /status`.
- `SmartM365-DeviceRegistration-Tool.config.json.template`: safe configuration template for retry and output settings.
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

Use `SmartM365-DeviceRegistration-Tool.config.json.template` as the committed model:

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
  "LogoPath": "WorkplaceCloudHub-lockup-WPF.png",
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
- leave the default `WorkplaceCloudHub-lockup-WPF.png` to use the WorkplaceCloudHub logo next to the script.

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

From the standalone release folder, open the diagnostic-only GUI:

```cmd
Start-SmartM365-DeviceRegistration-Tool-User.cmd
```

From the `SmartM365` repository folder, open the GUI:

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

## Offline demo mode

Launch `Start-SmartM365-DeviceRegistration-Demo.cmd` for screenshots and presentations. The original WPF screens, tabs and options remain available with fictional Contoso data and visibly marked simulated results. Use the Scene selector (`Overview`, `Issues`, `Completed`), or pass `-DemoScenario Issues`. `-DemoView` selects a named tab where available; `-DemoWindowWidth` and `-DemoWindowHeight` control the capture size.

Business actions are simulated. Demo startup bypasses operational configuration, authentication, collection, remote commands and mail delivery. Edits remain in memory. **Export demo outputs** generates only synthetic files under a unique `%TEMP%\SmartM365-Demo\` session folder. HTML-capable tools retain their report views; notification campaigns preview every shipped language template with fictional tokens. No real tenant or device evidence is required.

Keep the application folder, its original report/template sources, local branding assets, `SmartM365.GuiDemo.ps1`, `SmartM365.GuiDemo.Reports.ps1` and the local `*-Demo.ps1` profile together when copying the application. `-Demo -ValidateOnly` builds the WPF demo without displaying windows or opening files. Run `SmartM365/Tests/Test-SmartM365GuiDemoApplicationsOffline.ps1` in an STA process for offline interaction and export checks.


## Local review correction (2026-10-18)

Source version: 1.2. Enrollment cleanup exports every existing registry key for an enrollment ID and verifies each backup is nonempty before removing any registry key or EnterpriseMgmt task for that ID. An export failure blocks that enrollment cleanup and returns a failure with the backup path. The autonomous Hybrid Join endpoint applies the same protection.

Offline regression command from the repository root:

```powershell
pwsh -NoProfile -STA -ExecutionPolicy Bypass -File SmartM365/Tests/Test-SmartM365AppReviewDataPathsOffline.ps1 -Case EnrollmentBackup
```

These checks use temporary files, harmless child processes or mocked operational boundaries. They do not qualify a real tenant, endpoint repair, campaign delivery or deployment. Demo assets remain local in `DemoAssets/`, ignored by Git, and must be refreshed after GUI changes. Live deployment requires separate qualification.
# Optional Microsoft 365 license check

In User mode, expand **Microsoft 365 licenses (optional)** and select **Check my licenses**. This is a separate, interactive, read-only Graph check; it never runs automatically or changes the local diagnostic result, enrollment or licenses.

The isolated worker requests delegated `User.Read`, uses a process-only context and reads only `GET /v1.0/me?$select=id,userPrincipalName,assignedLicenses`. The signed-in account must match the Microsoft work account entered in the optional panel (initially the Windows work UPN). Results identify that verified Microsoft account, which can differ from the Windows identity. No administrator role, application secret, directory-wide permission or `/licenseDetails` call is used. Tenant policies may require approval; this tool does not grant permissions or bypass them.

Results distinguish verified assigned SKUs, a verified empty assignment, and **Unverifiable**. Refused access, cancellation, missing dependency, timeout, account mismatch or incomplete data are never interpreted as no license. SKU identifiers alone do not prove a full Microsoft 365 suite, an Intune entitlement, service provisioning or Office activation. The check shows the account and UTC timestamp; it does not claim continuous freshness.

Optional dependency: official `Microsoft.Graph.Authentication` **2.41.2** (Windows PowerShell 5.1), either installed for the user or placed alongside the tool under `Modules/Microsoft.Graph.Authentication/2.41.2`. It is not downloaded or installed automatically. The isolated Windows PowerShell STA worker retains a hidden console handle for WAM. Complete the Microsoft sign-in and user-consent window manually. The requested User.Read consent does not grant the user an administrator role. SDK output and error streams are drained in memory, never saved. Closing the GUI or a three-minute timeout ends its own worker only.

Minimal results are saved separately as `Output/Licenses_<id>.json` under the current user's output directory. No tokens or raw SDK error text are saved. A separate progress sidecar records phase/time only; safe failure diagnostics may include phase, exception type, HTTP status, HRESULT and AADSTS code. The simple **Copy report for support** action includes the optional result when available and remains clipboard-only.

References: [Get own user profile](https://learn.microsoft.com/en-us/graph/api/user-get?view=graph-rest-1.0), [user assignedLicenses](https://learn.microsoft.com/en-us/graph/api/resources/user?view=graph-rest-1.0), [official authentication module](https://www.powershellgallery.com/packages/Microsoft.Graph.Authentication/2.41.2).


Assigned SKU names are resolved locally using `SmartM365-DeviceRegistration.LicenseCatalog.json`, a 619-product snapshot of Microsoft's public licensing reference (reference date 2026-08-19). Multiple assignments retain all names; unknown or newly introduced SKUs remain unidentified with their exact IDs. The catalog contains no tenant data, makes no extra Graph request and is not automatically refreshed. Product names are assignment labels, not proof that every included service is enabled or healthy.

## User session connection renewal

User mode offers **Renew my work connection** / **Renouveler ma connexion professionnelle**, after a diagnostic. Confirmation defaults to No; cancel sends no request. Accept launches only `dsregcmd /refreshprt` under the current user token, followed by the unchanged User diagnostic. No elevation, join/leave, enrollment cleanup, policy change or auto-enrollment is performed. The local diagnostic and this explicit action run in a hidden background worker so the GUI remains responsive.

A zero native exit code means a request was accepted, not proof of a repaired connection. A present PRT with a strictly newer observed update time confirms renewal; missing/unchanged/unreadable evidence stays **requested, unconfirmed**, with advice to lock/unlock and contact support. A command failure and a failed post-check are distinct. This does not certify Intune enrollment, compliance or application sign-in. The local session-action receipt and clipboard support report include safe timestamps/status and identity, never token contents. Optional Graph remains a separate user action.
