# Branch Notes

This file tracks the current feature branch ownership for active SmartM365 work.

## `feature/smart-endpoint-diagnostics-analyzer`

Purpose: Smart Endpoint Diagnostics Analyzer PowerShell migration work only.

Notes:

- This branch ports the SmartLogAnalyzer behavior into SmartM365 as a PowerShell/WPF tool.
- Do not import the previous app Git history, GitHub release/signing workflow, PyInstaller build files, or SignPath configuration.
- Keep the app under `SmartM365/EndpointDiagnosticsAnalyzer/`.

## `feature/device-reboot-manager-gui`

Purpose: Device Reboot Manager work only.

Current expected tip:

- `771a7ad` - `Document Device Reboot Manager app`

Notes:

- This branch was cleaned after the Smart Intune Remediation Manager commit was moved to its own branch.
- It should not contain Smart Intune Remediation Manager changes.

## `feature/smart-intune-remediation-manager`

Purpose: Smart Intune Remediation Manager work only.

Current expected commits:

- `0c4d730` - `Add Smart Intune Remediation Manager`
- `e306382` - `Document Smart Intune Remediation Manager`

Notes:

- This branch contains the CLI and GUI under `SmartM365/IntuneRemediationManager/`.
- It also contains the root README update that introduces Smart Intune Remediation Manager.

## Local Worktree Reminder

There may be unrelated local modifications under `SmartM365/IntuneInventory/` and `SmartM365/IntuneRemediation/`.

Do not stage or push those files as part of either feature branch unless they are intentionally reviewed and selected.
