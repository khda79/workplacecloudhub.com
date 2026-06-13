# Smart Endpoint Diagnostics Analyzer

PowerShell/WPF endpoint diagnostics analyzer for Microsoft Intune and Windows endpoint troubleshooting.

This tool is the SmartM365 PowerShell edition of the previous SmartLogAnalyzer logic. It intentionally does not reuse the source project Git history, Python packaging, PyInstaller build files, GitHub release workflow, or SignPath signing configuration.

## Purpose

Smart Endpoint Diagnostics Analyzer reads Intune Device Diagnostics ZIP files and local endpoint diagnostic captures, then turns raw diagnostic files into a support-focused view of device state, Intune enrollment, IME errors, Windows Update posture, application and driver data, Windows 11 upgrade indicators, compliance signals, and recommended actions.

## Main Script

- `SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1`: PowerShell/WPF GUI and CLI-capable analyzer.
- `Start-SmartM365-EndpointDiagnosticsAnalyzer-GUI.cmd`: console-hidden launcher for normal GUI use.
- `SmartM365-logo.ico`, `workplacecloudhub.png`, `workplacecloudhub-header.png`: GUI icon and header branding assets.

## Current Capabilities

- Analyze Intune Device Diagnostics ZIP files.
- Collect a local diagnostic ZIP and analyze it.
- Parse DSRegCmd status, device join, PRT, WAM, tenant and TPM-related fields.
- Parse MDM enrollment registry exports under `HKLM\Software\Microsoft\Enrollments`.
- Parse `results.xml` collection status.
- Parse IME logs by known Intune Management Extension theme.
- Detect known Intune, MDM, app, authentication and Windows error codes in logs.
- Parse Windows Update Orchestrator registry, WU policy registry, and `ReportingEvents.log`.
- Generate a fresh `WindowsUpdate.log` with `Get-WindowsUpdateLog` whenever Windows Update ETL files are available; ETL decoding falls back to `Get-WinEvent` / `tracerpt.exe` only if log generation is unavailable or fails.
- Parse installed applications, PnP drivers, WiFi profiles, proxy and device identity details.
- Extract and parse MDM Diagnostics CAB content and `MDMDiagHTMLReport.html`.
- Scan EVTX event logs through `wevtutil.exe`.
- Parse battery report, firewall profile output, and certificate inventory.
- Parse Windows 11 upgrade compatibility indicators from `TargetVersionUpgradeExperienceIndicators`.
- Build local health score, top actions, root causes, timeline, search corpus, WUfB view and compliance summary.
- Run optional AI analysis with Claude, OpenAI or Ollama.
- Write one timestamped log file per launch under `Logs\`, with automatic retention of the 10 newest logs.
- Export a standalone HTML report.
- Export an anonymized diagnostic ZIP for support sharing.

## Not Included From The Previous Project

The following source-project elements are intentionally excluded:

- Python source and modules.
- PyInstaller `.spec`, `requirements.txt`, `build.bat`, `build/`, `dist/`, and Python caches.
- GitHub Actions release/signing workflow.
- SignPath organization, project, policy or artifact configuration.
- Previous repository Git history, commits, tags, release notes, and project memory files.

## Usage

The GUI self-elevates at startup. If it is launched from a standard user context, Windows will show a UAC prompt and the elevated instance continues with the same parameters.

Open the GUI:

```powershell
powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -File .\Devices\EndpointDiagnosticsAnalyzer\SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1
```

Or use the launcher:

```cmd
.\Devices\EndpointDiagnosticsAnalyzer\Start-SmartM365-EndpointDiagnosticsAnalyzer-GUI.cmd
```

Analyze a ZIP from the command line:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Devices\EndpointDiagnosticsAnalyzer\SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1 -Cli -ZipPath .\Output\DiagLogs.zip
```

Analyze and export an HTML report:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Devices\EndpointDiagnosticsAnalyzer\SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1 -Cli -ZipPath .\Output\DiagLogs.zip -ExportHtmlPath .\Output\DiagLogs-report.html
```

Analyze with AI from the command line:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Devices\EndpointDiagnosticsAnalyzer\SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1 -Cli -ZipPath .\Output\DiagLogs.zip -RunAI -AIProvider ollama
```

Collect local diagnostics and create a ZIP:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Devices\EndpointDiagnosticsAnalyzer\SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1 -Cli -CollectLocal -ZipPath .\Output\LocalDiag.zip
```

## Privacy

The core analysis runs locally.

Anonymized ZIP export redacts common email addresses, GUIDs, IPv4 addresses, user profile paths, tenant IDs, device IDs, serial numbers and user principal names in text-like files. Binary files are preserved as-is. Always review an anonymized ZIP before sharing it externally.

## Branch

Initial PowerShell migration branch:

```text
feature/smart-endpoint-diagnostics-analyzer
```
