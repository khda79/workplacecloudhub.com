# Smart Intune Remediation Manager

WPF interface for browsing, exporting, editing, and publishing Microsoft Intune remediation scripts with delegated interactive Microsoft Graph authentication.

## Launch

From the `SmartM365` folder:

```powershell
pwsh -STA -NoProfile -File .\IntuneRemediationManager\IntuneRemediationManager-GUI\SmartM365-IntuneRemediation-GUI.ps1
```

Or use the launcher, which hides the PowerShell console:

```cmd
.\IntuneRemediationManager\IntuneRemediationManager-GUI\Start-SmartM365-IntuneRemediation-GUI.cmd
```

The GUI starts maximized. Click `Connect Graph` before using Intune cloud actions.

## Authentication

Authentication is delegated and interactive only. The GUI does not use app-only authentication, certificates, client secrets, stored credentials, or unattended authentication.

Requested delegated Graph scopes:

- `DeviceManagementScripts.ReadWrite.All`
- `DeviceManagementConfiguration.Read.All`
- `Group.Read.All`

## Configuration

The local configuration file is `SmartM365-IntuneRemediation-GUI.config.json` next to the GUI script. It is local-only and ignored by Git.

Use `SmartM365-IntuneRemediation-GUI.config.template.json` as the committed template:

```json
{
  "LocalRemediationRoot": "",
  "PublishSourceNamePrefix": "SmartM365-",
  "PublishTargetNamePrefix": "SmartM365-"
}
```

`LocalRemediationRoot` is requested on first launch when it is empty. `PublishSourceNamePrefix` and `PublishTargetNamePrefix` control the Intune display name used during publish. For example, `SmartM365-Example` can be published as `EMERIT-Example` without renaming local files.

## Main Actions

- `PSScriptAnalyzer`: analyze the selected local or cloud scripts.
- `Save script`: save only the selected local package scripts.
- `Save all local`: create a timestamped ZIP archive of local packages.
- `Publish to Intune`: create or update the selected remediation.
- `Publish detection only`: publish only the selected detection script and leave Intune remediation content empty.
- `Publish ALL to Intune`: show a create/update summary, then publish every local package.
- `Save scripts as`: save copies of selected cloud scripts to a chosen folder.
- `Save all cloud`: export all cloud remediations to a timestamped ZIP, including metadata and assignments with group names.
- `Reset history`: duplicate the selected remediation, copy assignments, then delete the old Intune object.
- `Delete`: delete only the selected cloud remediation after explicit warning.
- `Export execution CSV`: download the selected remediation execution report.

## Logs

Activity is written to `SmartM365-IntuneRemediation-GUI.log` in this folder. At each launch, the previous log is archived with a timestamp and the GUI keeps the 10 most recent log files.
