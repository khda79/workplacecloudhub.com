# SmartM365 Devices

This folder groups SmartM365 tools that run on, diagnose, repair, notify, or manage Windows endpoints.

These tools are intentionally separated from tenant-wide inventory scripts. They are used by helpdesk, endpoint engineering, Intune administrators, and field support operators when the work is centered on a Windows device or a fleet of devices.

## Tool Map

- `DeviceRebootManager/`: localized WPF restart notification app for device restart governance and Intune Win32 deployment.
- `DeviceRegistrationTool/`: local GUI and CLI diagnostics for Intune enrollment, Hybrid Join, Entra device registration, PRT refresh, support bundles, and guarded admin repair actions.
- `EndpointDiagnosticsAnalyzer/`: analyzer for Intune Device Diagnostics ZIP files and local endpoint captures, with HTML report export and optional AI-assisted review.
- `IntuneHybridJoinToolkit/`: PsExec/LOT toolkit for running the autonomous Hybrid Join and Intune enrollment repair script across batches of devices.
- `Windows11UpgradeToolkit/`: PsExec/LOT toolkit for diagnosing Windows 10 devices that should move to Windows 11 through Intune/Autopatch/WUfB and running guarded upgrade actions.

## Choosing The Right Tool

Use `DeviceRegistrationTool/` when you are working on one device locally or through a support session and need a GUI, CLI export, support bundle, or guarded admin repair.

Use `IntuneHybridJoinToolkit/` when you need to push one autonomous repair script to several computers through LOT folders and PsExec, or when a single-file script must be deployed through a mechanism such as GPO.

Use `Windows11UpgradeToolkit/` when a batch of Windows 10 devices should migrate to Windows 11 through Intune, Windows Autopatch, or Windows Update for Business, and you need local evidence plus explicit guarded actions such as Windows Update force triggers or setup-based upgrade.

Use `EndpointDiagnosticsAnalyzer/` when you already have an Intune Device Diagnostics ZIP or need to collect and analyze a local endpoint diagnostic bundle.

Use `DeviceRebootManager/` when the goal is user-facing restart governance.

## Operational Safety

Runtime logs, exports, LOT folders, support bundles, local configuration files, and local memories stay outside Git through the repository ignore rules.

Device repair actions are guarded and must be selected explicitly. User-mode diagnostics stay read-only where the tool supports separate user and administrator modes.

## Documentation

Each tool has its own README with launch commands, configuration options, outputs, and guardrails. Start with the tool README before running repair, deployment, or Graph-connected actions.
