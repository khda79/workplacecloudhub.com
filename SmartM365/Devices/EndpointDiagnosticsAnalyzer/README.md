# Smart Endpoint Diagnostics Analyzer

PowerShell/WPF endpoint diagnostics analyzer for Microsoft Intune and Windows endpoint troubleshooting.

## Release status

The repository version is **0.3.0 stable**, promoted from preview18 without changing the diagnostic algorithms. It supports managed installation through PowerShell Gallery or Microsoft Intune. Promotion and local package building do not publish to either service.

Validation boundary: the preview18 local-device and Intune-ZIP workflows were verified on Windows 10, including collection coverage, complete EVTX counts and finalization. The stable build has not yet been tested interactively on Windows 11. Stable is the selected release channel, not a claim that every diagnostic rule has been certified on every supported device.

Preview18 corrects collection completeness and finalization reporting:

- Winget diagnostic logs and minidumps have no default age/count limit, like IME and Windows Update logs. File-type scope remains explicit; archive safety checks still apply. This may increase collection size and duration.
- Every folder-copy operation records eligible/copied counts, extension/age/count exclusions and actual errors. Scope exclusions are informational, unavailable sources are distinct, partial copies are warnings, and failed copies remain errors. Registry sources are checked before export; access denial is not treated as absence. The Collection Status tab retains reasons and coverage marks missing/partial evidence.
- Finalization has three labelled stages: export the complete result, load it, and render it. Each stage has its own logged duration; export also reports bytes. Loading and rendering remain synchronous in the GUI, so it can pause during these stages, but the current stage is painted first. No record/message limit was introduced to shorten this work.

Preview17 prioritizes evidence correctness over elapsed time:

- All outer ZIP entries passing the archive safety checks are extracted. EVTX reading streams all levels without a default record cap on either supported PowerShell runtime. Full messages, duplicate-source paths and the full ordered timeline are retained; display limits do not define diagnostic coverage.
- IME and Windows Update parsing streams source records. IME groups preserve application identities and occurrences; only an explicit later success for the same identifiable operation resolves an earlier failure. Unknown timestamps are not classified as recent. Informational source records remain available in their original logs even when not represented as diagnostic findings.
- Empty or failed evidence cannot produce a Healthy assessment. GUI, HTML and AI context include supported-area coverage. Entra device authentication is not MDM enrollment; internal upgrade indicators are unverified context, and unknown readiness checks do not reduce the score.
- Malformed collection/MDM XML, failed restart collection, neutral readiness fields, Wi-Fi profiles and unrelated battery HTML have explicit regression tests. Every MDM report candidate is processed independently with source status; conflicting values retain provenance.
- Installation validates signed runtime files and packaged SHA-256 hashes before staging a replacement. A previous installation is retained for rollback; detection checks integrity and Gallery updates refuse Intune-owned targets. Uninstall requires an owned, bounded target. Automatic updates remain disabled.
- The old anonymized ZIP action is now **redacted text ZIP**. It sanitizes common identifiers in text and entry names, accepts BOM-marked UTF-16, and blocks unsupported binary/oversized/undecodable entries. This is deliberately restrictive: ordinary diagnostic archives containing EVTX/CAB are not exportable through this action. It is **not a guarantee of anonymity**; manually inspect any output before sharing. Original diagnostics are unchanged.

Historical notes below describe earlier versions, not the current completeness contract. In particular, preview15 extraction/EVTX/IME caps are superseded by preview17.

Preview16 refines performance measurements and progress accuracy:

- Identity, enrollment, IME parsing/context, extended inventory and Windows Update substeps have separate timings; IME parsing also records per-file and grouping timings.
- Windows Update registry decoding selects only the policy subtree, and parsed trace results are cached by content SHA-256, schema and analyzer version. Invalid caches fall back to source parsing and cached source paths are rebound to the current extraction.
- Numbered analysis progress describes the phase currently running, not the phase that just finished.
- Event coverage separates EVTX files read, observed channels, failed/skipped files and files not processed because of the cap. Enrollment candidates are labelled as candidates, not confirmed active enrollments.

Preview15 optimizes performance and operator accuracy against the latest local-device and Intune diagnostics scans:

- Local system inventory uses bounded targeted CIM queries instead of `Get-ComputerInfo`.
- Large unused setup, CBS and system files remain visible in the ZIP inventory but are not extracted eagerly.
- IME parsing remains bounded per file and caches parsed results by ZIP SHA-256, cache schema and analyzer version.
- EVTX analysis focuses on actionable levels, removes duplicate logical events across copied channels and reports discarded duplicates.
- The timeline retains the 300 most recent entries; enrollment tabs distinguish current Intune candidates from historical registry records.
- Analysis progress reports seven measured phases and collection no longer displays an unreliable linear ETA.

Preview14 optimized findings confirmed against local-device and Intune diagnostics scans:

- IME messages repair common OEM 850/UTF-8 mojibake before classification and display.
- Operator summaries exclude .NET stack frames from the Insights grid while retaining full technical evidence in the selection panel and collapsed HTML details.
- Flighting requests that fail with a WinHTTP error but explicitly apply a default value, including the immediately preceding generic transport trace, are retained as non-actionable evidence and do not reduce the diagnostic score.
- Real Winget installation failures remain actionable with a concise title, error code and targeted recommendation.
- Generated Windows Update logs are cached by diagnostics-ZIP SHA-256, avoiding repeated conversion of the same ETL set.

Preview13 keeps the responsive preview9 architecture and corrects parser and scoring defects confirmed against a real local diagnostic capture:

- Managed and unmanaged MDM policy sections use exact normalized headings, preventing duplicate policy inventories.
- Registry decoding supports `REG_MULTI_SZ` and `REG_QWORD`; `RedReason=None` no longer creates a false Windows 11 blocker.
- IME classification uses the full source message, recognizes decimal WinHTTP errors such as 12007 and treats the normal SendHeartbeatReport flighting fallback as informational.
- Enrollment registry evidence claims a compliant MDM enrollment only when an authoritative Intune/MDM provider or endpoint is present.
- The overall score exposes a category-by-category penalty breakdown, while historical Modern Auth codes and startup counts remain informational.
- Battery and CPU readiness details are operator-readable, additional Windows Update codes are mapped, and common QWORD timestamps are decoded.
- The GUI identifies the PowerShell 7 analyzer and Windows PowerShell 5.1 collector runtimes and shows numbered progress.

- The overall diagnostic score ignores informational findings, duplicate reboot evidence, historical Event Log records and ETL-only Windows Update signatures.
- Windows Update explicit result records are grouped separately from ETL diagnostic traces, with occurrence ranges, affected items and clear evidence labels.
- Enrollment summaries contain only parsed enrollment evidence instead of generic blank device fields.
- Device-health counters and the overall diagnostic score are labelled separately in the GUI and HTML report.
- Missing BitLocker, Defender, storage or Store evidence identifies whether the source is an Intune package or a local collection.

- Local extended evidence now resolves from the containing folder, restoring BitLocker, Defender, storage, restart, proxy and performance findings.
- MDM diagnostics test every available CAB, preserve paths containing spaces, select the CAB containing `MDMDiagReport.xml`/HTML and populate device, connection and policy views.
- Modern Auth reports only concrete AADSTS codes; IME fallback, Workplace Join and user-script context reduce false actionable errors.
- Windows Update shows grouped root-cause signatures, occurrence ranges, operator-friendly meanings and recommendations instead of treating trace lines as failed installations.
- The GUI launcher and elevation path preserve PowerShell 7 when available and fall back to Windows PowerShell 5.1 only when needed.
- Connection Info now combines the active adapter, IPv4 address, gateway, DNS, network profile, proxy and Entra/Workplace state.
- Local `MDM_Enrollment` registry exports and direct `MDMDiagReport.xml`/`.html` files are parsed without requiring a CAB.
- MDM views expose report metadata, enrollment evidence, blocked and unmanaged policy rows, report files and explicit empty states.
- Windows Update log parsing correctly splits lines, groups repeated HRESULTs and separates overview, policies, update history, errors/warnings and raw details.
- The local collector writes future Windows Update history as structured JSON while the analyzer remains compatible with the earlier table format.
- Stable wrapped navigation, lighter tables, readable labels and populated empty-state rows replace reordering tabs and blank panels.
- ZIP analysis and local collection both run in isolated processes so the WPF window remains responsive.
- Numbered collection progress ends with an explicit PASS/FAIL result marker; a ZIP alone is no longer accepted as proof of success.
- Windows PowerShell 5.1 may not expose redirected child-process exit codes; collection requires both the ZIP and explicit PASS marker, analysis requires an importable CLIXML result, and any explicit non-zero exit code still fails.
- The GUI evaluates the analysis CLIXML presence before deciding whether an asynchronous analysis completed successfully.
- ZIP extraction validates entry count, expanded size, compression ratio and traversal paths, and temporary extraction is cleaned on reset, close or CLI completion.
- IME results are limited to recent actionable root causes, with expected flighting, session, fallback and assignment messages treated as informational.
- Historical Event Log records remain available in the timeline but do not reduce the bounded diagnostic score.
- The former Compliance view is explicitly a local configuration assessment; user SSO controls collected under SYSTEM are marked `NOT_EVALUATED`.
- Localized driver and certificate output, Windows Update logs, `results.xml`, dynamic EVTX channels and Windows 11 applicability are parsed more conservatively.

The last standalone public package remains [Endpoint Diagnostics Analyzer v0.2.0](https://github.com/khda79/workplacecloudhub.com/releases/tag/endpoint-diagnostics-analyzer-v0.2.0).

All PowerShell release files must be Authenticode-signed by `workplacecloudhub.com` before publication.

## Purpose

Smart Endpoint Diagnostics Analyzer reads Intune Device Diagnostics ZIP files and local endpoint diagnostic captures, then turns raw diagnostic files into a support-focused view of device state, Intune enrollment, IME errors, Windows Update posture, application and driver data, Windows 11 upgrade indicators, local configuration evidence, and recommended actions.

## Main capabilities

- Analyze Intune Device Diagnostics ZIP files.
- Collect a local diagnostic ZIP and analyze it.
- Parse DSRegCmd, MDM enrollment, Intune Management Extension, Windows Update, applications, drivers, Wi-Fi, proxy, certificates and Windows 11 compatibility information.
- Extract and parse MDM Diagnostics CAB content and EVTX event logs.
- Build a diagnostic score, prioritized actions, root causes, timeline, WUfB view and local configuration assessment.
- Run optional AI analysis with Claude, OpenAI or Ollama.
- Export a standalone HTML summary or a restricted redacted-text ZIP requiring manual privacy review.

## Installation architecture

Two managed channels use the same signed runtime:

- PowerShell Gallery: self-service installation for the current user by default.
- Intune Win32: managed all-user installation under the SYSTEM context.

No scheduled task is created. Automatic updates are disabled by default. Gallery users update only when an administrator or user explicitly runs the update command; Intune remains the authority for Intune installations.

Local collection and ZIP analysis run in isolated processes. The analysis worker prefers PowerShell 7 and also reads EVTX on Windows PowerShell 5.1, without default event caps. Final result loading/rendering can briefly pause the GUI. Native diagnostic commands have explicit timeouts; a timed-out step is logged and skipped, and the user can cancel the collection. Native command arguments are preserved, local device identity falls back to the extended system inventory, and the overview separates actionable health findings from historical Event Log and IME records.

Default locations:

- Gallery current-user runtime: `%LOCALAPPDATA%\Programs\SmartM365\EndpointDiagnosticsAnalyzer`
- Intune all-user runtime: `%ProgramFiles%\SmartM365\EndpointDiagnosticsAnalyzer`
- Per-user configuration and logs: `%LOCALAPPDATA%\SmartM365\EndpointDiagnosticsAnalyzer`
- Start menu shortcut: `SmartM365\Smart Endpoint Diagnostics Analyzer`

Uninstallation preserves per-user configuration and logs unless `-RemoveUserData` is explicitly supplied.

## PowerShell Gallery

After `SmartM365.EndpointDiagnosticsAnalyzer` 0.3.0 has been published, install the stable package with PSResourceGet:

```powershell
Install-PSResource -Name SmartM365.EndpointDiagnosticsAnalyzer -Version 0.3.0 -Repository PSGallery -TrustRepository
Import-Module SmartM365.EndpointDiagnosticsAnalyzer
Install-SmartM365EndpointDiagnosticsAnalyzer
```

PowerShellGet equivalent:

```powershell
Install-Module -Name SmartM365.EndpointDiagnosticsAnalyzer -RequiredVersion 0.3.0 -Repository PSGallery
Import-Module SmartM365.EndpointDiagnosticsAnalyzer
Install-SmartM365EndpointDiagnosticsAnalyzer
```

An update is always explicit:

```powershell
Update-SmartM365EndpointDiagnosticsAnalyzer
```

Build and preview publication locally:

```powershell
.\PowerShellGallery\SmartM365-Build-EndpointDiagnosticsAnalyzerGalleryPackage.ps1 -Force
.\PowerShellGallery\SmartM365-Publish-EndpointDiagnosticsAnalyzerGalleryPackage.ps1 -ForceBuild
```

Publication requires `-Execute`. Future prereleases additionally require `-AllowPrereleasePublication`. The API key is read from `PSGALLERY_API_KEY`; the script does not persist it.

## Microsoft Intune

Build the signed Win32 package:

```powershell
$intuneBuild = .\IntuneWin32\SmartM365-Build-EndpointDiagnosticsAnalyzerIntunePackage.ps1 `
    -IntuneWinAppUtilPath C:\tmp\SmartM365Tools\IntuneWinAppUtil.exe `
    -OutputRoot C:\tmp\SmartM365-EndpointDiagnosticsAnalyzer-IntunePackage `
    -Force
$intuneBuild.IntuneWinPath
```

Preview the Intune publication and pilot assignment without connecting to Graph:

```powershell
.\Deploy\SmartM365-EndpointDiagnosticsAnalyzer-PublishIntune.ps1 `
    -IntuneWinPath $intuneBuild.IntuneWinPath `
    -CreatePilotGroup `
    -AssignPilotGroup
```

`-AssignPilotGroup` uses the `available` intent by default. Use `-PilotAssignmentIntent required` only for a deliberate mandatory deployment. Tenant changes occur only with `-Execute`, followed by interactive Microsoft Graph authentication.

## Security and privacy

The core analysis runs locally. A remembered AI API key is encrypted with Windows DPAPI for the current user and stored in the application data directory; it is not stored in plaintext. If the legacy `%USERPROFILE%\.smartloganalyzer_ai.json` file exists, its configuration is migrated and any plaintext API key in that legacy file is cleared.

A command-line AI key is never forwarded in plaintext during UAC self-elevation. It is transferred through a temporary DPAPI-protected file restricted to the application data path and removed by the elevated process.

Redacted-text ZIP export sanitizes common identifiers in supported text-like files and entry names. Binary and otherwise unsupported content is rejected, not silently passed through. This is not guaranteed anonymization: always inspect an output before sharing. Original diagnostic archives are unchanged. Diagnostic extraction directories are temporary and are removed after CLI analysis, GUI reset or application close.

## Direct usage from the repository

Open the GUI:

```cmd
Start-SmartM365-EndpointDiagnosticsAnalyzer-GUI.cmd
```

Validate without opening the GUI:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1 -ValidateOnly
```

Analyze a ZIP and export HTML:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1 `
    -Cli `
    -ZipPath .\Output\DiagLogs.zip `
    -ExportHtmlPath .\Output\DiagLogs-report.html
```

## Release bundle

Create one checksum-protected handoff containing the Gallery package and the Intune administrator files:

```powershell
.\PowerShellGallery\SmartM365-New-EndpointDiagnosticsAnalyzerPreviewBundle.ps1 -Force
```

The default bundle path is `C:\tmp\SmartM365-EndpointDiagnosticsAnalyzer-Stable-0.3.0`. The builder keeps its historical `PreviewBundle` filename for compatibility and selects `Stable` or `Preview` from the version manifest. Creation performs no Gallery publication, Graph connection or tenant change.

## Local validation

The packaging test parses every PowerShell file, builds the Gallery package, imports the module, performs a controlled temporary installation, validates detection and then uninstalls it:

```powershell
.\Tests\SmartM365-Test-EndpointDiagnosticsAnalyzerPackaging.ps1
```
