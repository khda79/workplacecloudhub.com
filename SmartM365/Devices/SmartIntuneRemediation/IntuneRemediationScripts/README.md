# Intune Remediation Scripts

This folder groups SmartM365 Intune remediation, detection, diagnostic, and action scripts by functional area and scenario. It belongs to `SmartIntuneRemediation`, together with the CLI and GUI managers used to publish and manage these packages in Microsoft Intune.

## Organization

- `WindowsUpdate/`: Windows Update scan, cache, service, policy, and proxy scenarios.
- `DeliveryOptimization/`: Delivery Optimization and Content Engine health scenarios.
- `WUfB/`: Windows Update for Business identity and PolicyState binding scenarios.
- `SetupDiag-Upgrade/`: SetupDiag, Windows 10/11 upgrade diagnostics, and upgrade readiness scenarios.
- `Intune-MDM/`: Intune, MDM, Hybrid Join, Intune Management Extension, and Data Boundary scenarios.
- `Disk-Cleanup-Storage/`: disk space, cleanup candidates, and upgrade storage readiness scenarios.
- `UWP-Store-AppRepository/`: AppRepository, Microsoft Store, and Windows Update health scenarios.
- `Standalone/`: operator actions, diagnostics, and one-off scripts that should not run as recurring Intune remediations.

## Naming Convention

Active scripts live in scenario folders. File names do not include version numbers; the current script version belongs in the script header and Git history keeps the change history.

When a scenario contains a standard Intune remediation pair, the folder usually contains:

- `SmartM365-*-Detection.ps1`
- `SmartM365-*-Remediation.ps1`

When a script is not designed to run repeatedly as a remediation, use an explicit action or diagnostic naming pattern such as:

```text
SmartM365-*-Action.ps1
SmartM365-*-Diagnostic.ps1
```

## Consolidated Scenarios

- `WindowsUpdate/WindowsUpdate-Reset/`: replaces older Autopatch `0x80244007` and forced Windows Update reset scenarios.
- `WindowsUpdate/Policy-Blockers/`: consolidates WSUS/GPO, WSUS remnants, NoAutoUpdate, Policy-Blocking-Access, and WUfB configuration blockers.
- `WindowsUpdate/Service-And-Scan-Health/`: replaces older service health, service refresh, and forced scan scenarios.
- `WindowsUpdate/Cache-Health/`: replaces download failure scenarios.
- `Intune-MDM/MDM-Enrollment-Repair/`: consolidates stale device join, enrollment state, Hybrid Join, and missing MDM task scenarios. The detection script targets stale or broken local state; if it runs through Intune, basic IME reachability is already proven.
- `SetupDiag-Upgrade/Upgrade-Staging-Health/`: consolidates missing upgrade files and upgrade residue scenarios.
- `Disk-Cleanup-Storage/Upgrade-Storage-Readiness/`: consolidates free space, cleanup candidates, and forced disk cleanup scenarios.
- `DeliveryOptimization/ContentEngine-Health/`: replaces the older Delivery Optimization issue scenario.
- `UWP-Store-AppRepository/WU-Health/`: replaces the older broad endpoint repair scenario.
- `SetupDiag-Upgrade/Upgrade-Diagnostics/`: consolidates SetupDiag-required and upgrade-blocking issue scenarios.
- `Standalone/WindowsUpdateLog/`, `Standalone/Repair-DISM/`, `Standalone/Upgrade-Actions/`, `Standalone/Network-Diagnostics/`, `Standalone/WindowsPolicy/`, and `Standalone/Device-Recovery/`: keep one-off actions and manual diagnostics outside recurring remediations.

## Validation Before Publishing

Before publishing a package to Intune:

- review detection and remediation scripts together;
- run PSScriptAnalyzer when possible;
- confirm that remediation logic is idempotent;
- avoid storing tenant-specific values, group IDs, hostnames, or customer data in scripts;
- use the GUI or CLI manager to publish with delegated Graph authentication.

Generated exports, downloaded cloud scripts, reports, logs, and local archives are runtime artifacts and should remain outside Git.
