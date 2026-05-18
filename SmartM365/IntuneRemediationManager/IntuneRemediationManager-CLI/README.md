# Intune Remediation CLI

Interactive deployment utility for SmartM365 Intune remediation packages.

## Script

- `SmartM365-Deploy-IntuneRemediation-CLI.ps1`: creates or updates Intune remediation packages through Microsoft Graph `deviceHealthScripts`.
- `../IntuneRemediationManager-GUI/SmartM365-IntuneRemediation-GUI.ps1`: opens a local WPF GUI to browse Intune remediations, edit local SmartM365 detection/remediation scripts, publish local packages, reset Intune execution history, and export execution reports to CSV.

By default, the Intune Author/Publisher field is set to the interactive Microsoft Graph account used to run the deployment. Use `-Publisher` only when you need to override that value.

By default, the Intune description is built from the detection script and, when available, the remediation script. Each script is read from its `# Description:` header. Use `-Description` only when you need to override it.
When detection and remediation descriptions differ, the GUI publishes them on separate lines.

## Authentication

The script uses delegated interactive Microsoft Graph authentication only. It does not support app-only authentication, certificates, client secrets, stored credentials, or unattended execution.

Required delegated Graph permissions:

- `DeviceManagementScripts.ReadWrite.All`
- `DeviceManagementConfiguration.Read.All`
- `Group.Read.All`

The second scope is used for Intune report export jobs when downloading execution reports.

## Examples

Deploy one package folder:

```powershell
.\IntuneRemediationManager\IntuneRemediationManager-CLI\SmartM365-Deploy-IntuneRemediation-CLI.ps1 -Path .\IntuneRemediation\WindowsUpdate\Cache-Health
```

Deploy all packages under a folder and update existing Intune remediations with the same display name:

```powershell
.\IntuneRemediationManager\IntuneRemediationManager-CLI\SmartM365-Deploy-IntuneRemediation-CLI.ps1 -Path .\IntuneRemediation\WindowsUpdate -Recurse -UpdateExisting
```

Deploy a package and replace its assignments with one Entra group:

```powershell
.\IntuneRemediationManager\IntuneRemediationManager-CLI\SmartM365-Deploy-IntuneRemediation-CLI.ps1 -Path .\IntuneRemediation\WindowsUpdate\Cache-Health -AssignmentGroupId "00000000-0000-0000-0000-000000000000" -RunRemediationScript
```

Omit `-AssignmentGroupId` to create or update the remediation package without changing Intune assignments.

Open the GUI:

```powershell
pwsh -STA -NoProfile -File .\IntuneRemediationManager\IntuneRemediationManager-GUI\SmartM365-IntuneRemediation-GUI.ps1
```

Or use the launcher, which hides the PowerShell console:

```cmd
.\IntuneRemediationManager\IntuneRemediationManager-GUI\Start-SmartM365-IntuneRemediation-GUI.cmd
```

See `../IntuneRemediationManager-GUI/README.md` for GUI configuration, actions, and log behavior.



