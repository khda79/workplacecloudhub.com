# Smart Endpoint Diagnostics Analyzer

## Current distribution 1.0.0 — build 2

The package and GUI product version are **1.0.0**. **Build 2** aligns the visible interface and report versions. The [1.0.0 release](https://github.com/khda79/workplacecloudhub.com/releases/tag/endpoint-diagnostics-analyzer-v1.0.0) contains only the current build 2 assets; superseded releases and original build assets have been removed.

| Download | File |
| --- | --- |
| Standalone ZIP | [SmartM365-EndpointDiagnosticsAnalyzer-v1.0.0-build.2.zip](https://github.com/khda79/workplacecloudhub.com/releases/download/endpoint-diagnostics-analyzer-v1.0.0/SmartM365-EndpointDiagnosticsAnalyzer-v1.0.0-build.2.zip) |
| Intune Win32 package | [SmartM365-EndpointDiagnosticsAnalyzer-1.0.0-build.2.intunewin](https://github.com/khda79/workplacecloudhub.com/releases/download/endpoint-diagnostics-analyzer-v1.0.0/SmartM365-EndpointDiagnosticsAnalyzer-1.0.0-build.2.intunewin) |
| SHA-256 checksums | [SmartM365-EndpointDiagnosticsAnalyzer-v1.0.0-build.2.sha256](https://github.com/khda79/workplacecloudhub.com/releases/download/endpoint-diagnostics-analyzer-v1.0.0/SmartM365-EndpointDiagnosticsAnalyzer-v1.0.0-build.2.sha256) |
| DemoAssets | [SmartM365-EndpointDiagnosticsAnalyzer-v1.0.0-build.2-DemoAssets.zip](https://github.com/khda79/workplacecloudhub.com/releases/download/endpoint-diagnostics-analyzer-v1.0.0/SmartM365-EndpointDiagnosticsAnalyzer-v1.0.0-build.2-DemoAssets.zip) |
| Release notes | [SmartM365-EndpointDiagnosticsAnalyzer-v1.0.0-build.2-RELEASE-NOTES.md](https://github.com/khda79/workplacecloudhub.com/releases/download/endpoint-diagnostics-analyzer-v1.0.0/SmartM365-EndpointDiagnosticsAnalyzer-v1.0.0-build.2-RELEASE-NOTES.md) |
| Release manifest | [SmartM365-EndpointDiagnosticsAnalyzer-v1.0.0-build.2.manifest.json.txt](https://github.com/khda79/workplacecloudhub.com/releases/download/endpoint-diagnostics-analyzer-v1.0.0/SmartM365-EndpointDiagnosticsAnalyzer-v1.0.0-build.2.manifest.json.txt) |

Launch `GalleryPackage/SmartM365.EndpointDiagnosticsAnalyzer/1.0.0/Runtime/Start-SmartM365-EndpointDiagnosticsAnalyzer-Demo.cmd` from the root of the extracted standalone release. The separate DemoAssets archive contains the current fictional GUI captures, applicable HTML reports, synthetic exports and website/presentation variants. It performs no tenant or device operation.

The signed runtime source is identified by the [build 2 source tag](https://github.com/khda79/workplacecloudhub.com/tree/endpoint-diagnostics-analyzer-v1.0.0-build.2/SmartM365/Devices/EndpointDiagnosticsAnalyzer). GitHub publication does not publish a version to the PowerShell Gallery service.

PowerShell/WPF endpoint diagnostics analyzer for Microsoft Intune and Windows endpoint troubleshooting.

## Release status

**1.0.0 build 2 is the current GitHub distribution**, linked above. It includes the signed module and administrator packaging files; their presence does not deploy anything to Intune.

Historical validation boundary: the published release reports real local-device and Intune-ZIP workflows verified on Windows 10, including collection coverage, complete EVTX counts and finalization. The user reported successfully completing a subsequent interactive test of stable 0.3.0 on Windows 11 on September 7, 2026. Stable is the selected release channel, not a claim that every diagnostic rule has been certified on every supported device.

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

The current public bundle is **1.0.0 build 2**; use the download links above and the installation instructions below.

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

## PowerShell Gallery channel

The downloadable build 2 bundle includes module version **1.0.0**. GitHub release publication does not publish that version to the Gallery service. To install this exact build, use the standalone ZIP linked above and import its bundled module as shown below. Gallery service updates remain a separate administrator-controlled channel. Intune-owned installations must be updated through Intune.

Automatic updates are disabled and no scheduled update task is created.

## Verify and install the build 2 bundle

Download the **Standalone ZIP** and **SHA-256 checksums** from the table above into the same folder. The first checksum row is the application ZIP; the remaining rows cover the DemoAssets and IntuneWin packages.

```powershell
$zip = '.\SmartM365-EndpointDiagnosticsAnalyzer-v1.0.0-build.2.zip'
$published = ((Get-Content -LiteralPath '.\SmartM365-EndpointDiagnosticsAnalyzer-v1.0.0-build.2.sha256' -TotalCount 1) -split '\s+')[0]
$actual = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash
if ($published -notmatch '^[0-9A-Fa-f]{64}$' -or $actual -ne $published) { throw 'ZIP SHA-256 mismatch.' }
```

Extract the complete archive, open PowerShell in the extracted `SmartM365-EndpointDiagnosticsAnalyzer-v1.0.0-build.2` directory, then install the bundled module runtime:

```powershell
Import-Module .\GalleryPackage\SmartM365.EndpointDiagnosticsAnalyzer\1.0.0\SmartM365.EndpointDiagnosticsAnalyzer.psd1
Install-SmartM365EndpointDiagnosticsAnalyzer
```

Open the Start menu shortcut described above. The bundle contains `SHA256SUMS.txt`, the Gallery module and administrator packaging files. Ordinary installation does not require running publication scripts. For a later bundle update, download and verify the chosen release, then import its module and explicitly run its installation command; do not assume a bundle import registers the module for Gallery updates.

PowerShell payloads are Authenticode-signed by `workplacecloudhub.com`. Signature trust depends on the workstation's certificate trust and execution policy; this is not a claim of public-trust signing.

## Maintainer package publication

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

The core analysis runs locally. Optional external AI analysis is a separate user action; review sensitive diagnostics before sending them to an external provider. A remembered AI API key is encrypted with Windows DPAPI for the current user and stored in the application data directory; it is not stored in plaintext. If the legacy `%USERPROFILE%\.smartloganalyzer_ai.json` file exists, its configuration is migrated and any plaintext API key in that legacy file is cleared.

A command-line AI key is never forwarded in plaintext during UAC self-elevation. It is transferred through a temporary DPAPI-protected file restricted to the application data path and removed by the elevated process.

Redacted-text ZIP export sanitizes common identifiers in supported text-like files and entry names. Binary content such as EVTX/CAB and oversized or undecodable entries are rejected, not silently passed through; ordinary diagnostic archives containing these files cannot be exported with this action. This is not guaranteed anonymization: always inspect an output before sharing. Original diagnostic archives are unchanged. Diagnostic extraction directories are temporary and are removed after CLI analysis, GUI reset or application close.

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

The default bundle path is `C:\tmp\SmartM365-EndpointDiagnosticsAnalyzer-Stable-1.0.0`. The builder keeps its historical `PreviewBundle` filename for compatibility and selects `Stable` or `Preview` from the version manifest. Creation performs no Gallery publication, Graph connection or tenant change.

## Local validation

The packaging test parses every PowerShell file, builds the Gallery package, imports the module, performs a controlled temporary installation, validates detection and then uninstalls it:

```powershell
.\Tests\SmartM365-Test-EndpointDiagnosticsAnalyzerPackaging.ps1
```

## Offline demo mode

Launch `Start-SmartM365-EndpointDiagnosticsAnalyzer-Demo.cmd` for screenshots and presentations. The original WPF screens, tabs and options remain available with fictional Contoso data and visibly marked simulated results. Use the Scene selector (`Overview`, `Issues`, `Completed`), or pass `-DemoScenario Issues`. `-DemoView` selects a named tab where available; `-DemoWindowWidth` and `-DemoWindowHeight` control the capture size.

Business actions are simulated. Demo startup bypasses operational configuration, authentication, collection, remote commands and mail delivery. Edits remain in memory. **Export demo outputs** generates only synthetic files under a unique `%TEMP%\SmartM365-Demo\` session folder. HTML-capable tools retain their report views; notification campaigns preview every shipped language template with fictional tokens. No real tenant or device evidence is required.

Keep the application folder, its original report/template sources, local branding assets, `SmartM365.GuiDemo.ps1`, `SmartM365.GuiDemo.Reports.ps1` and the local `*-Demo.ps1` profile together when copying the application. `-Demo -ValidateOnly` builds the WPF demo without displaying windows or opening files. Run `SmartM365/Tests/Test-SmartM365GuiDemoApplicationsOffline.ps1` in an STA process for offline interaction and export checks.


## Distribution version transition

The 1.0.0 package uses GUI component 1.0.0 and includes the complete signed demo runtime. Only the current build 2 release assets remain published. Git history and source tags preserve earlier source revisions. Local DemoAssets stay Git-ignored; reviewed demo archives are separate GitHub release assets.
