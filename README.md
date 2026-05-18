# workplacecloudhub.com

Repository for [workplacecloudhub.com](https://workplacecloudhub.com/), a technical hub focused on Workplace, Cloud, and AI engineering.

WorkplaceCloudHub is built around practical guidance, automation patterns, and technical notes for modern enterprise IT platforms. The site is organized around articles, scripts, tools, and field-oriented content for engineers working on secure and reliable workplace and cloud operations.

## Content

- `SmartM365/`: validated PowerShell automation scripts and tools for Microsoft 365 and endpoint administration.

Additional project directories will be added over time.

## SmartM365 Tools

### Smart Device Reboot Manager

`SmartM365/DeviceRebootManager/` contains **Smart Device Reboot Manager**, a local WPF user notification app for SmartM365 device restart governance.

- Presents restart status and actions to end users with a localized GUI.
- Supports recommended or required restart modes, preview/test launchers, configurable postpone choices, and safer state/log handling.
- Can be launched directly with PowerShell or through bundled CMD wrappers that hide the console.

### Smart Intune Remediation Manager

`SmartM365/IntuneRemediationManager/` contains **Smart Intune Remediation Manager**, an interactive manager for Microsoft Intune remediation scripts.

- `IntuneRemediationManager-GUI/`: WPF interface to browse local remediation packages, view Intune cloud remediations, edit scripts, run PSScriptAnalyzer, publish to Intune, duplicate/reset execution history, delete selected cloud remediations, export execution CSV reports, and archive local or cloud scripts.
- `IntuneRemediationManager-CLI/`: delegated interactive CLI deployment script for Intune remediation packages.

The manager uses delegated interactive Microsoft Graph authentication. It does not use stored credentials, client secrets, certificates, or SmartM365 app-only runtime authentication.

## Focus Areas

- Modern Workplace: device management, productivity platforms, identity integration, and operational standards.
- Cloud Engineering: readiness, governance, monitoring, and platform operations.
- Automation: PowerShell scripts, remediation flows, reporting helpers, and repeatable operational patterns.
- AI for IT: troubleshooting assistance, documentation, script review, log analysis, and faster technical decisions.

## Contact

contact@workplacecloudhub.com
