# Smart Intune Remediation GUI

WPF interface for browsing, exporting, editing, and publishing Microsoft Intune remediation scripts with delegated interactive Microsoft Graph authentication.

The script reports version `1.1` and requires PowerShell 7 or later. It is distributed as repository source; there is no dedicated release archive, PowerShell Gallery package, installer, published release checksum, or automatic updater. See the [workspace README](../README.md) for distribution, integrity, security, and known-limit details.

## Launch

From the `SmartM365` folder:

```powershell
pwsh -STA -NoProfile -File .\Intune\Remediation\GUI\SmartM365-IntuneRemediation-GUI.ps1
```

Or use the launcher, which hides the PowerShell console:

```cmd
.\Intune\Remediation\Start-SmartM365-IntuneRemediation-GUI.cmd
```

The GUI starts maximized. Click `Connect Graph` before using Intune cloud actions.

PowerShell 7 must be installed. The launcher prefers `%ProgramFiles%\PowerShell\7\pwsh.exe`; its legacy Windows PowerShell fallback cannot satisfy the GUI script's `#Requires -Version 7.0` declaration.

If browser sign-in is canceled, blocked, or never returns to PowerShell, launch from a visible PowerShell window and use device code authentication:

```powershell
pwsh -STA -NoProfile -File .\Intune\Remediation\GUI\SmartM365-IntuneRemediation-GUI.ps1 -GraphAuthMode DeviceCode
```

Keep the PowerShell window visible because Microsoft Graph prints the device code there.

## Authentication

Authentication is delegated and interactive only. The GUI does not use app-only authentication, certificates, client secrets, stored credentials, or unattended authentication.

Requested delegated Graph scopes:

- `DeviceManagementScripts.ReadWrite.All`
- `DeviceManagementConfiguration.Read.All`
- `Group.Read.All`

The GUI calls Microsoft Graph `beta` endpoints. Delegated consent does not replace the Intune RBAC rights and active Intune/Remediations licensing required in the target tenant. No tenant-side validation is performed by `-ValidateOnly`.

## Configuration

The local configuration file is `../SmartM365-IntuneRemediation-GUI.config.json` next to the root launcher. It is local-only and ignored by Git.

Use `../SmartM365-IntuneRemediation-GUI.config.template.json` as the committed template:

```json
{
  "LocalRemediationRoot": "Packages",
  "PublishSourceNamePrefix": "SmartM365-",
  "PublishTargetNamePrefix": "SmartM365-"
}
```

`LocalRemediationRoot` defaults to `Packages`. Relative paths are resolved from the `IntuneRemediation` root folder. `PublishSourceNamePrefix` and `PublishTargetNamePrefix` control the Intune display name used during publish. For example, `SmartM365-Example` can be published as `EMERIT-Example` without renaming local files.

## Main Actions

- `PSScriptAnalyzer`: analyze the selected local or cloud scripts.
- Local package grid `In Intune`: shows whether the local Detection/Remediation content already exists in Intune. Name matches with different content are shown separately as `Name only`.
- Cloud remediations are separated into `Active` and `Not deployed` tabs based on deployment status.
- `Save script`: save only the selected local package scripts.
- `Save all local`: create a timestamped ZIP archive of local packages.
- `Publish to Intune`: create or update the selected remediation.
- `Publish detection only`: publish only the selected detection script and leave Intune remediation content empty.
- `Publish ALL to Intune`: show a create/update summary, then publish every local package.
- `Save scripts as`: save copies of selected cloud scripts to a chosen folder.
- `Save all cloud`: export all cloud remediations to a timestamped ZIP, including metadata and assignments with group names.
- `Compare local/cloud`: compare the selected local package with the selected cloud remediation script content.
- `Activity`: open the activity log tab in the lower workspace when you need runtime details.
- `Reset history`: duplicate the selected remediation, copy assignments, then delete the old Intune object.
- `Delete`: delete only the selected cloud remediation after explicit warning.
- `Export execution`: download the selected remediation execution report as CSV or Excel. At startup, the GUI checks `ImportExcel`; if it is missing, it explains why it is needed and asks before installing it for the current user. If installation is declined or not possible, Excel is not offered and CSV remains available. After a successful Excel export, the GUI asks whether to open the workbook.

New and updated packages default to SYSTEM, 64-bit execution, role scope tag `0`, and disabled Intune signature enforcement. Publishing a newly named selected package does not add a second create-confirmation dialog; review the selection first. `Reset history` creates a timestamped replacement, copies assignments, and deletes the original, while `Delete` permanently removes the selected cloud object.

PSScriptAnalyzer is installed from PowerShell Gallery for the current user if it is missing when analysis is first requested. Repository scripts carry an Authenticode signature from the self-signed `workplacecloudhub.com` certificate (`D70ECB7B00377EBFB76B304C08DFC6620584E114`); local trust and execution-policy configuration remain the operator's responsibility.

## Logs

Activity is written to `Logs/SmartM365-IntuneRemediation-GUI.log` under this folder. At each launch, the previous log is archived with a timestamp in `Logs/` and the GUI keeps the 10 most recent log files.
