# SmartM365 Intune

This folder groups SmartM365 tools and package workspaces that target Microsoft Intune rather than local device-only workflows.

## Tool Map

- `Remediation/`: Intune remediation packages with a delegated PowerShell 7/WPF Graph publishing and administration GUI.
- `IntunePackager/`: Smart IntunePackager, a standalone PowerShell 7/WPF tool that builds Win32 `.intunewin` packages with the official Microsoft Win32 Content Prep Tool and publishes them to Intune through Microsoft Graph. Assignments are made in the Intune admin center.
