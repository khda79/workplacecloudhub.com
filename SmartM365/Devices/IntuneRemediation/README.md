# Smart Intune Remediation

Workspace for SmartM365 Intune remediation scripts and the tools used to manage them.

This area contains both the remediation packages themselves and the local operator tools used to publish, compare, archive, and report on Microsoft Intune remediation scripts.

## Organization

- `Packages/`: detection, remediation, diagnostic, and action scripts organized by Intune scenario.
- `GUI/`: WPF operator interface for local package review, Intune cloud comparison, PSScriptAnalyzer checks, publish/update actions, export, reset history, and cloud deletion.
- `CLI/`: delegated interactive deployment utility for creating or updating Intune remediation packages from PowerShell.

## Authentication Model

The GUI and CLI tools use delegated interactive Microsoft Graph authentication. They do not use stored credentials, client secrets, certificates, or SmartM365 app-only authentication.

Typical delegated scopes are documented in each tool README:

- `DeviceManagementScripts.ReadWrite.All`
- `DeviceManagementConfiguration.Read.All`
- `Group.Read.All`

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
- `CLI/README.md`: CLI deployment examples and options.
