# workplacecloudhub.com

Repository for [workplacecloudhub.com](https://workplacecloudhub.com/), a technical hub focused on Workplace, Cloud, and AI engineering.

WorkplaceCloudHub is built around practical guidance, automation patterns, and technical notes for modern enterprise IT platforms. The site is organized around articles, scripts, tools, and field-oriented content for engineers working on secure and reliable workplace and cloud operations.

## What This Repository Contains

This repository is a field toolkit rather than a single application. It contains independent PowerShell projects for Microsoft 365, Azure, Azure Virtual Desktop, Citrix, and Windows endpoint operations. The common goal is to help administrators collect reliable evidence, standardize exports, troubleshoot devices, and run repeatable operational actions without embedding tenant-specific values in the public repository.

Most tools are designed to run directly from their project folder. Local tenant configuration, generated CSV files, logs, support bundles, LOT folders, and runtime JSON files are intentionally ignored by Git.

## Content

- `SmartM365/`: validated PowerShell automation scripts and tools for Microsoft 365 and endpoint administration.
- `SmartAzure/`: Azure infrastructure inventory, governance, security posture, and cost optimization scripts.
- `SmartAzureVirtualDesktop/`: Azure Virtual Desktop inventory, health, diagnostics, autoscale, FSLogix storage, and cost optimization scripts.
- `SmartCitrix/`: Citrix on-premises and Citrix Cloud inventory scripts, with separate implementation areas for each platform model.

Additional project directories will be added over time.

## Getting Started

Clone the repository, then open the project README that matches the platform you want to work with:

```powershell
git clone https://github.com/khda79/workplacecloudhub.com.git
cd workplacecloudhub.com
```

Recommended entry points:

- Microsoft 365 and Windows endpoint tools: `SmartM365/README.md`
- SharePoint migration validation toolkit: `SmartM365/SharePointMigration/README.md`
- Azure inventory and governance scripts: `SmartAzure/README.md`
- Azure Virtual Desktop inventory scripts: `SmartAzureVirtualDesktop/README.md`
- Citrix inventory scripts: `SmartCitrix/README.md`

For SmartM365 tenant-based inventory scripts, start with `SmartM365/Setup/README.md` and `SmartM365/Setup/SmartM365-AppRegistration-Permissions.md`. Device tools under `SmartM365/Devices/` are mostly local endpoint tools and do not require the SmartM365 app-only setup unless their own README explicitly says so.

## Repository Principles

- Keep tenant values local: use `*.local.json` files and never commit secrets, tenant IDs, webhook URLs, certificates, exports, logs, or support bundles.
- Keep generated data out of Git: CSV exports, diagnostics, LOT folders, transcripts, and logs are runtime artifacts.
- Prefer read-only inventory by default: write-capable scripts and repair actions are documented in their own README and guarded by explicit parameters or GUI actions.
- Keep public documentation generic: examples should use placeholders and safe paths only.

## SmartAzureVirtualDesktop Tools

`SmartAzureVirtualDesktop/` contains PowerShell automation for Azure Virtual Desktop environments.

- `SmartInventory/Estate/SmartAzureVirtualDesktop-AVDEstate-Inventory.ps1`: exports host pools, workspaces, application groups, applications, desktops, session hosts, scaling plans, private endpoint connections, and estate summaries.
- `SmartInventory/Health/SmartAzureVirtualDesktop-SessionHostHealth-Inventory.ps1`: exports session host health, user sessions, and host pool capacity summaries.
- `SmartInventory/Diagnostics/SmartAzureVirtualDesktop-Diagnostics-Inventory.ps1`: exports Azure Monitor diagnostic settings for AVD resources and resources missing diagnostics.
- `SmartInventory/Scaling/SmartAzureVirtualDesktop-ScalingPlan-Inventory.ps1`: exports scaling plans, schedules, assignments, and host pool autoscale coverage.
- `SmartInventory/FSLogix/SmartAzureVirtualDesktop-FSLogixStorage-Inventory.ps1`: exports candidate FSLogix storage accounts and Azure Files shares in AVD resource groups.
- `SmartInventory/Cost/SmartAzureVirtualDesktop-CostOptimization-Inventory.ps1`: exports AVD cost review signals such as inactive session hosts, hosts without sessions, host pools without autoscale, and unattached disks in AVD resource groups.

## SmartCitrix Tools

`SmartCitrix/` contains PowerShell automation for Citrix estate inventory. The project separates on-premises Citrix Virtual Apps and Desktops scripts from future Citrix Cloud / Citrix DaaS API scripts.

- `Citrix-OnPrem/SmartInventory/CVAD/SmartCitrix-OnPrem-CVADSite-Inventory.ps1`: exports CVAD site, controller, zone, admin, and service status data.
- `Citrix-OnPrem/SmartInventory/CVAD/SmartCitrix-OnPrem-CVADDelivery-Inventory.ps1`: exports catalogs, delivery groups, published applications, access policies, tags, and reboot schedules.
- `Citrix-OnPrem/SmartInventory/CVAD/SmartCitrix-OnPrem-CVADMachineSessionHealth-Inventory.ps1`: exports machine, session, desktop, and health summary data.
- `Citrix-OnPrem/SmartInventory/CVAD/SmartCitrix-OnPrem-CVADPolicy-Inventory.ps1`: exports Citrix policy sets, policies, settings, and filters when available from the local SDK.
- `Citrix-OnPrem/SmartInventory/CVAD/SmartCitrix-OnPrem-CVADHostingPower-Inventory.ps1`: exports hosting connections, power management settings, power actions, and reboot cycles.
- `Citrix-OnPrem/SmartInventory/StoreFront/SmartCitrix-OnPrem-StoreFront-Inventory.ps1`: exports StoreFront deployment, store, authentication, receiver, farm, and gateway data from StoreFront PowerShell modules.
- `Citrix-OnPrem/SmartInventory/Licensing/SmartCitrix-OnPrem-Licensing-Inventory.ps1`: exports CVAD licensing configuration and available licensing SDK data.

## SmartAzure Tools

`SmartAzure/` contains PowerShell automation for Azure Resource Manager environments.

- `SmartInventory/Governance/SmartAzure-AzureEstate-Inventory.ps1`: exports management groups, subscriptions, regions, resource groups, resources, locks, providers, and a per-subscription summary.
- `SmartInventory/RBAC/SmartAzure-RBAC-Inventory.ps1`: exports role assignments, privileged assignments, custom roles, and RBAC summaries.
- `SmartInventory/Cost/SmartAzure-CostOptimization-Inventory.ps1`: exports unattached disks, old snapshots, unused public IPs, stopped/deallocated VMs, and cost review summaries.
- `SmartInventory/Network/SmartAzure-NetworkExposure-Inventory.ps1`: exports public IPs, Internet-sourced inbound NSG allow rules, load balancers, application gateways, private endpoints, and exposure summaries.
- `SmartInventory/Governance/SmartAzure-PolicyCompliance-Inventory.ps1`: exports Azure Policy assignments, definitions, initiatives, exemptions, policy state details, and compliance summaries.
- Planned inventory areas include Defender for Cloud, backup, storage security, and Key Vault security.

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

### Smart SharePoint Migration Toolkit

`SmartM365/SharePointMigration/` contains **Smart SharePoint Migration Toolkit**, a source-to-destination SharePoint migration validation toolkit.

- Inventories SharePoint source and destination content and permissions.
- Compares files, permissions, and source scan history with CSV and Excel outputs.
- Generates guarded destination cleanup scripts from reviewed comparison results.
- Provides a migration template, local launchers, and a portable Python bootstrap script for comparison/export helpers.

### Smart Intune Remediation

`SmartM365/Devices/IntuneRemediation/` contains the SmartM365 Intune remediation workspace.

- `Packages/`: Intune detection, remediation, diagnostic, and action scripts organized by scenario.
- `GUI/`: WPF interface to browse local remediation packages, view Intune cloud remediations, edit scripts, run PSScriptAnalyzer, publish to Intune, duplicate/reset execution history, delete selected cloud remediations, export execution CSV reports, and archive local or cloud scripts.
- `CLI/`: delegated interactive CLI deployment script for Intune remediation packages.

The GUI and CLI use delegated interactive Microsoft Graph authentication. They do not use stored credentials, client secrets, certificates, or SmartM365 app-only runtime authentication.

### Smart Intune Hybrid Join Toolkit

`SmartM365/Devices/IntuneHybridJoinToolkit/` contains a PsExec/LOT toolkit for diagnosing and repairing Hybrid Entra Join plus Intune enrollment issues.

## Focus Areas

- Modern Workplace: device management, productivity platforms, identity integration, and operational standards.
- Cloud Engineering: readiness, governance, monitoring, and platform operations.
- Automation: PowerShell scripts, remediation flows, reporting helpers, and repeatable operational patterns.
- AI for IT: troubleshooting assistance, documentation, script review, log analysis, and faster technical decisions.

## Contact

contact@workplacecloudhub.com
