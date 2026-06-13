# workplacecloudhub.com

Repository for [workplacecloudhub.com](https://workplacecloudhub.com/), a technical hub focused on Workplace, Cloud, and AI engineering.

WorkplaceCloudHub is built around practical guidance, automation patterns, and technical notes for modern enterprise IT platforms. The site is organized around articles, scripts, tools, and field-oriented content for engineers working on secure and reliable workplace and cloud operations.

## Content

- `SmartM365/`: validated PowerShell automation scripts and tools for Microsoft 365 and endpoint administration.

Additional project directories will be added over time.

## SmartM365 Tools

### Smart Endpoint Diagnostics Analyzer

`SmartM365/Devices/EndpointDiagnosticsAnalyzer/` contains **Smart Endpoint Diagnostics Analyzer**, a PowerShell/WPF endpoint diagnostics analyzer for Microsoft Intune and Windows endpoint troubleshooting.

- Analyzes Intune Device Diagnostics ZIP files and local endpoint diagnostic captures.
- Parses DSRegCmd, MDM enrollment, MDM Diagnostics CAB/HTML, IME logs, Windows Update ReportingEvents/ETL, EVTX event logs, hardware/security data, installed applications, drivers, WiFi profiles, proxy/device identity details, and Windows 11 upgrade indicators.
- Exports standalone HTML reports and anonymized diagnostic ZIP files for support sharing, with optional Claude/OpenAI/Ollama analysis.

### Smart Device Reboot Manager

`SmartM365/Devices/DeviceRebootManager/` contains **Smart Device Reboot Manager**, a local WPF user notification app for SmartM365 device restart governance.

- Presents restart status and actions to end users with a localized GUI.
- Supports recommended or required restart modes, preview/test launchers, configurable postpone choices, and safer state/log handling.
- Can be launched directly with PowerShell or through bundled CMD wrappers that hide the console.

### Smart DeviceRegistration Tool

`SmartM365/Devices/DeviceRegistrationTool/` contains **Smart DeviceRegistration Tool**, a local PowerShell/WPF tool for Intune device registration, Hybrid Join, and Entra device registration diagnostics.

- Provides user-mode diagnostics and guarded admin repair actions.
- Collects local enrollment, MDM, `dsregcmd`, policy, and support-bundle evidence.
- Supports localized GUI usage and CLI diagnostic exports for support handoff.

### Smart Inventory

`SmartM365/SmartInventory/` groups the inventory scripts that can feed Power BI datasets and other reporting or operational consumers.

- `ActiveDirectoryInventory/`: Active Directory inventory and reporting.
- `ExchangeInventory/`: Exchange Online and Exchange on-premises inventory.
- `M365Inventory/`: Microsoft 365, Entra, licensing, domain, and Intune inventory. Intune inventory now lives under `M365Inventory/IntuneInventory/`.

### Smart Intune Remediation Manager

`SmartM365/Devices/SmartIntuneRemediation/` contains the SmartM365 Intune remediation workspace.

- `IntuneRemediationScripts/`: Intune detection, remediation, diagnostic, and action scripts organized by scenario.
- `IntuneRemediationManager/`: interactive manager for Microsoft Intune remediation scripts.
- `IntuneRemediationManager/IntuneRemediationManager-GUI/`: WPF interface to browse local remediation packages, view Intune cloud remediations, edit scripts, run PSScriptAnalyzer, publish to Intune, duplicate/reset execution history, delete selected cloud remediations, export execution CSV reports, and archive local or cloud scripts.
- `IntuneRemediationManager/IntuneRemediationManager-CLI/`: delegated interactive CLI deployment script for Intune remediation packages.

The manager uses delegated interactive Microsoft Graph authentication. It does not use stored credentials, client secrets, certificates, or SmartM365 app-only runtime authentication.

### Smart Intune Hybrid Join Toolkit

`SmartM365/Devices/SmartIntuneHybridJoinToolkit/` contains a PsExec/LOT toolkit for diagnosing and repairing Hybrid Entra Join plus Intune enrollment issues.

## Focus Areas

- Modern Workplace: device management, productivity platforms, identity integration, and operational standards.
- Cloud Engineering: readiness, governance, monitoring, and platform operations.
- Automation: PowerShell scripts, remediation flows, reporting helpers, and repeatable operational patterns.
- AI for IT: troubleshooting assistance, documentation, script review, log analysis, and faster technical decisions.

## Contact

contact@workplacecloudhub.com
