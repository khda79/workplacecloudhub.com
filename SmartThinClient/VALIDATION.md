# Smart ThinClient validation

## Power confirmations and visible version - 2026-11-13

Version remains 1.1.4 and appears at the lower right of the control panel, operational shell and visual Preview. Customer WebShellTitle values remain independent. Restart and Shutdown share a compact branded WPF confirmation, with Cancel as the default, explicit action labels and a save-work reminder. Cancellation leaves the primary shell close guard intact. Confirmed requests capture Windows exit codes without forcing application closure; failures restore the previous close guard and appear in the shell status.

151/151 isolated regressions pass under Windows PowerShell 5.1 and PowerShell 7. Power command success, cancellation, launch failure and nonzero exit codes are mocked. Non-displayed WPF checks cover both dialog constructors, safe defaults, icons and styles; 24 operational/Preview layouts verify the version footer alongside custom titles. Generated operational button events exercise routing and visible error handling with mocked actions. No Windows restart, shutdown or sign-out is performed by these checks.

## Three-area workspace toolbar - 2026-11-13

Version remains 1.1.4. Preview and the operational shell keep provider/Back/Refresh/Open browser on the left, the configurable WebShellTitle centered, and Settings/Admin/sign-out/restart/shutdown on the right. Independent wrapping groups and equal outer columns preserve the center when controls are hidden or names differ in length. Long names keep their full tooltip. The existing WebShellTitle field is reused; no configuration key or local value changes.

142/142 isolated regressions pass under Windows PowerShell 5.1 and PowerShell 7. Non-displayed WPF checks exercise both actual window constructors at widths760/1080/1700, Citrix/Hybrid and default/long custom names:24 layouts preserve exact center, outside-edge group alignment, non-overlap and literal custom-title/tooltip values. No browser or system action is launched by these checks.

## Sign-out choice and native device settings - 2026-11-13

Version remains 1.1.4. Completed Apply/Restore results offer Sign out now and Later when another sign-in is relevant. The second confirmation defaults to Cancel, reminds the user to save work, and requests current-session sign-out without forcing application closure. Refusal/error handling and the absence of this action in Audit/Preview are tested without signing out.

140/140 isolated regression cases pass under Windows PowerShell 5.1 and PowerShell 7. Non-displayed WPF checks verify sign-out choices/defaults, native page routing, all six native panels under the standard user without Explorer, and error containment for the desktop URI launcher. The native helper compiles under Windows PowerShell 5.1; importing it performs no device initialization or writes. The real mutation methods and physical device operations remain outside this validation.

An operator-reported shell disappearance after interacting with Settings and the demo entry was accompanied by a Windows PowerShell InvalidComObject report. The shell was relaunched in the same standard-user session without changing policy. Demo entry-to-catalog navigation for Citrix and AVD succeeds in both the baseline and managed-callback test harnesses, so the original disappearance is not a deterministically reproduced portal defect. Managed WebView2 handlers and guarded toolbar events contain errors and retain the host window; they do not establish complete crash immunity.

Actual Windows Settings activation failed with 0x80040900 through process launch, WinRT and the activation manager without Explorer. A separately authorized temporary Explorer test opened all six pages; activation failed again after that Explorer stopped. The final implementation uses native ThinClient panels in CustomShell/Explorer-free sessions, preserving the configurable four default buttons and optional Bluetooth/Printers. It does not bootstrap Explorer, change Windows edition, reset packages or alter trust/security settings.

The six native panels and routing construct under the standard user with Explorer absent. They read actual adapter/display/audio/printer/radio signals and enforce the RDP resolution and administrator clock restrictions. Display close/keep/timeout/failure and volume/mute action forwarding are exercised with all device writes mocked. Actual display, volume, clock, time-zone, Wi-Fi connection, default-printer and device-pairing changes are not qualified by these checks. Missing radios and saved networks are explicit device/session limitations. These are focused device controls, not full Windows Settings parity.

## Apply experience and prepared all-users custom shell - 2026-11-13

Version remains 1.1.4. Apply defaults to the exact phrase APPLY; custom configured phrases remain supported. Modern shared Apply/Restore confirmation retains exact case-sensitive validation, owner icon and secure-password handling. Successful operations show readable results, next steps and rollback paths. Open files describes the operational directory; unique GUI action IDs preserve earlier evidence and rollback. Admin is hidden for standard users and rechecks administrator membership before launching the protected control panel, with elevation for filtered administrator tokens.

The new opt-in CustomShell mode prepares a shared, administrator-owned ProgramData launcher and schema 3 rollback, and uses the local all-users Custom User Interface policy instead of the machine Run key. Apply requires explicit enablement, a supported local policy, persistent embedded Web mode, a detected browser, configured destination, elevation and typed confirmation. Tests cover backup-before-policy ordering, policy failure preserving old startup, restore refusal for unrelated policy, malformed snapshots and persistent Hybrid startup without a cancellable provider picker. Existing schema 2 rollback remains supported.

134/134 isolated scenarios pass under Windows PowerShell 5.1 and PowerShell 7. Non-displayed endpoint WPF construction verifies short-phrase validation, modern templates/icons and readable completion feedback. Actual endpoint checks pass for administrator-token detection, four-profile read-only preflight, native local-policy reading, shared-storage ACLs and browser-initialization failure retaining the Admin toolbar. Storage checks use a unique test directory, not the configured operational MachineRoot. Watched Run/Winlogon values, policy files, original user launcher/rollback and boot state remain unchanged by these checks. Parser checks pass in both engines; production-script PSScriptAnalyzer reports no errors. The regression harness retains its known synthetic plaintext-to-SecureString fixture warning classified as an error by PSScriptAnalyzer.

No CustomShell policy write, real Apply/Restore, sign-out or next-sign-in transition has been executed for this change. Configuration and code preparation do not establish live shell replacement or end-to-end recovery. Domain policy precedence, actual all-user sign-in and a complete guarded Apply/Restore cycle still require endpoint qualification. The operator-enabled lab gates remain true; product defaults remain false.

## Official client download links - 2026-11-13

Version remains 1.1.4. Audit offers Download Citrix Workspace and Download Windows App on the corresponding undetected-client cards. Links are fixed official HTTPS pages, opened in the default browser only on click; no automatic installer download or execution occurs. Legacy Remote Desktop does not suppress the Windows App link. Detected clients and unrelated WebOnly profiles omit the prompts.

118/118 isolated regressions pass under Windows PowerShell 5.1 and PowerShell 7, plus a signed frozen-source PS5.1 run. WPF construction/event checks pass locally and on the endpoint, verifying profile-specific links, their visible labels and URL tooltips, installed-client suppression and exact fixed destinations through a simulated process launcher. No browser or installer is opened by the test. The compact panel still constructs at 960x620 with eight Hybrid cards.

Actual endpoint Audit evidence maps Citrix to its official download page, AVD to Windows App in Microsoft Store, and Hybrid to both. Seventeen installed-file hashes, five unchanged external configuration files/links, watched registry, code/configuration ACL and boot state are verified. Both operator-enabled lab gates remain true. No client installation, native launch, portal authentication or real Apply/Restore was performed.

## Distinct Web and native client checks - 2026-11-13

Version remains 1.1.4. Audit separates Citrix Web/Citrix Workspace and AVD Web/Windows App/Remote Desktop (legacy). Web URL/browser signals never imply native-client installation. Missing optional clients are informational in Web mode; required native clients need review with a Web fallback or block readiness without one. Windows App requires its own package signal, while a detected classic Remote Desktop executable is labelled legacy and warns. Unknown installation details remain Not reported. Product detection, launch routing and Apply/Restore engines are unchanged.

114/114 isolated regressions pass under Windows PowerShell 5.1 and PowerShell 7, plus a signed frozen-source PS5.1 run. Nine additional cases cover absent/present clients, current versus legacy AVD clients, both AVD clients, unknown evidence, native requirements, Web fallback and all eight Hybrid cards. Non-displayed WPF construction passes locally and on the endpoint at 960x620 with eight Hybrid cards.

Actual endpoint Audit runs for Citrix, AVD and Hybrid pass. The detected legacy Remote Desktop executable is an explicit warning; Windows App is not detected for the audited account, and Citrix Workspace is not detected. Both Web cards report only configured URL/browser signals, with the absent optional clients shown informationally. All 17 installed-file hashes and five unchanged external configuration files/links are verified; watched registry, code/configuration ACL and boot state remain unchanged by the checks. The operator-enabled lab gates remain true. No client was installed, portal opened, native session launched or real Apply/Restore executed in this correction.

## Configuration-blocked GUI status - 2026-11-13

Version remains 1.1.4. The actual GUI action handler now reports Apply blocked by configuration or Restore blocked by configuration for the corresponding disabled-gate refusal. Other errors retain the previous failure status and detailed message. No guard or mutation engine is changed; product defaults remain false.

105/105 isolated regressions pass under Windows PowerShell 5.1 and PowerShell 7, plus a signed frozen-source PS5.1 run. Actual-handler tests verify both disabled-gate labels, preserved no-change detail, re-enabled controls and the distinct administrator-refusal path when gates are enabled.

By explicit operator request, both gates are enabled in the backed-up lab runtime configuration; product defaults remain false. Endpoint verification confirms only those two configuration values changed, four runtime profiles stayed unchanged, and all 17 installed files match the manifest. Under an actual administrator token, both GUI access checks pass while blank confirmation still refuses execution before either mutation engine. Non-displayed handler fixtures verify both configuration-blocked labels and the simulated non-administrator failure status. Watched registry, code/configuration files and ACL, boot state and temporary-task cleanup are unchanged by these checks. The standard-user endpoint campaign was not rerun because that user had no open session; no real Apply/Restore was executed.

## Optional Bluetooth and Printers settings - 2026-11-13

Version remains 1.1.4. New configurations allow the Windows Bluetooth and Printers & scanners settings URIs, with independent Bluetooth/Printers visibility flags set to false. The original four buttons remain visible by default. Set both flags true to expose all six; an omitted allowed-list page is never reintroduced by a visibility flag. Existing explicit flags, custom page lists and partial profile inheritance remain authoritative. Preview and the operational shell use the same menu and live-settings notice.

103/103 isolated regressions pass under Windows PowerShell 5.1 and PowerShell 7, plus a signed frozen-source PS5.1 run. Coverage includes template/built-in default agreement, explicit enablement, independent hiding, legacy configured pages, allowed-list omissions, profile inheritance and generated-shell serialization. Non-displayed WPF construction under the endpoint's standard user verifies four default buttons, six enabled buttons in both menu modes, each optional button hidden independently, icon inheritance, the Preview live-settings notice and an empty menu. Actual button events use a simulated launcher to confirm the exact Bluetooth/Printers URI for each mode; no Windows settings page is opened and no device is paired or installed by this test.

Both shortcuts are enabled explicitly in the backed-up lab configuration. Other configuration fields and prior Settings choices are preserved. All 14 guarded CLI cases pass, and all ten Audit/Preview target cards still identify the standard domain user correctly. Seventeen file hashes, five configuration links, unchanged code/configuration ACL, watched registry values, boot state and temporary-task cleanup pass. Both Apply/Restore gates remain false. No operational Launch or successful real Apply/Restore is qualified.

## Target account identity correction - 2026-11-13

Version remains 1.1.4. Audit now distinguishes the diagnostic target identity from the operational local account. A matching authenticated Windows token confirms the current domain user without an Active Directory query; a different unconfirmed account stays unknown. The target card shows Domain account detected instead of incorrectly reporting Not found locally. Disabled local accounts remain warnings. No domain enabled-state or remote authentication claim is inferred, and the local SID used by operational guards is unchanged.

99/99 isolated regressions pass under Windows PowerShell 5.1 and PowerShell 7, including domain/qualified-name matching, ambiguous other accounts, disabled and colliding local accounts, invalid tokens and unchanged dedicated/Strict/kiosk local-user requirements. The signed frozen source also passes under Windows PowerShell 5.1.

After deployment, all 14 guarded CLI scenarios pass under the verified interactive standard domain-user token. All ten Audit/Preview results resolve that same domain identity, leave the operational local SID empty and the domain enabled state unknown, and produce a successful target card and Checks passed global summary. Seventeen installed-file hashes, five configuration links, all five unchanged external configuration files, unchanged operational guard bodies, code/configuration ACL, watched registry values, boot state and temporary-task cleanup are verified. Both mutation gates remain false. The summary is verified from actual evidence using the installed rendering-data function; no additional desktop interaction or real Apply/Restore was performed in this correction.

## Compact controls, Audit summary and Preview Settings - 2026-11-13

Version remains 1.1.4. The panel starts at 960x620 with a smaller header, wrapping action groups and shared rounded button/profile-selector styles. The shell, Preview and limited-settings dialog use the same control resources. Audit now has colored local checks and a global status; missing signals, local-target warnings and unsupported requested kiosk capabilities are distinguished. Disabled mutation gates are informational and no remote readiness is inferred. Profile changes hide the prior summary; Audit errors show a blocked status.

Settings also inherits the branded window icon and supports independent boolean Network/Display/Sound/DateAndTime visibility flags, all true by default. Configured page omissions and partial profile inheritance are preserved. Both the operational shell and Preview share the same filter; invalid flag types or non-ms-settings URI schemes are refused. An empty menu reports that no pages are enabled.

86/86 isolated regressions pass under Windows PowerShell 5.1 and PowerShell 7; the frozen signed source also passes under PS5.1. Non-displayed WPF construction verifies the compact panel, style templates, colored summary, owner-icon inheritance, four default pages, an independently hidden page and the empty-menu state. Fourteen actual CLI checks pass under the endpoint's standard user token with both mutation gates false. Initial desktop interaction verified the compact panel, summary, styled profile popup, Citrix-only Preview and Settings opening the native Display page without changing settings. The final icon/visibility refinements were checked by construction and CLI; their final endpoint visual review was not repeated after Computer Use was stopped. Installed-file/hash, configuration/link/ACL, watched-registry and boot checks pass, with temporary test tasks removed.

By explicit request, Preview Settings now opens the configured live Windows settings menu, respecting its enable flag and rejecting other URI schemes. Preview text describes this exception. Power/system actions, native clients, external browser, WebView2 download/popup/permission guards and Apply/Restore protections are retained. Provider buttons now match the effective profile in both Preview and the generated shell; only Hybrid exposes all three providers. Historical claims below that Preview Settings was disabled are superseded. Final signed-source and endpoint validation are recorded with the local campaign evidence.

## Public demo portal follow-up - 2026-11-13

With the supplied public Citrix-style and AVD-style demo URLs configured, both endpoints returned HTTP 200 from the test endpoint. Actual standard-user interactions in the installed WebView2 Preview passed each landing page, Enter demo workspace, Home/Apps/Desktops catalogues, simulated application and desktop sessions, return to catalogue and simulated sign-out. AVD re-entry and its catalogue Refresh also passed. These are fictional website sessions; no real Citrix/AVD authentication or remote session is qualified.

All 14 guarded CLI scenarios passed again with both provider URLs detected in each audit. Only the two backed-up provider URL fields changed; WebOnly remains the default and its homepage is unchanged. Both configuration gates remain false. Seventeen installed source/asset/dependency files match the deployment manifest; watched registry values, code/configuration ACL and boot state are unchanged, and temporary test tasks are removed. No executable source or version change was needed.

## Guarded application campaign - 2026-11-13

Version remains 1.1.4; this campaign changes no executable source. The existing 63 isolated regressions and 15 source-handler checks for WebView2 Preview restrictions pass under both Windows PowerShell 5.1 and PowerShell 7. Six offline SDK preparation checks pass, including pinned-package extraction, repeatability, and refusal of missing/changed libraries or a different package hash. Handler checks use fake event sources, not live permission prompts or downloads.

Under a verified standard user token, 14 actual CLI scenarios pass: Audit/Preview for Auto, Citrix, AVD, WebOnly and Hybrid; ValidateOnly; invalid-profile rejection; and Apply/Restore refusal. Ten evidence pairs have matching profiles/status, valid CSV and timestamped logs. An elevated administrator token also receives the expected false-configuration-gate refusal for both Apply and Restore. Both gates remain false throughout.

Actual desktop interactions cover shortcut startup and reopening, branding/title, GUI Audit with WebOnly/Citrix/Hybrid, profile selection, WebOnly and Hybrid visual Preview, labeled provider fallback, same-site navigation, Back, Refresh, maximization, Preview closure and parent-control recovery, and Open output with visible JSON/CSV files. Final code/configuration hashes, watched startup/policy registry values, code/configuration ACL and boot state match the baseline; temporary test tasks are removed. This campaign does not qualify successful real Apply/Restore, startup/kiosk enforcement, system actions, native clients or Citrix/AVD authentication. Those operational paths remain isolated-test coverage or untested, as appropriate.

## Modern embedded browser correction - 2026-11-12

Version remains 1.1.4. The legacy IE-based WPF control could show JavaScript errors on modern pages even though the shell opened. Embedded browsing now uses WebView2 with application-local verified SDK libraries and an existing Evergreen runtime; no IE emulation registry value, browser policy or runtime installer is applied. Preview retains disabled system/settings/external controls and blocks new windows, downloads and permission requests. Audit/CLI Preview and Apply/Restore configuration gates are unchanged. Final live rendering qualification is recorded separately from offline tests.

63/63 isolated regressions pass under Windows PowerShell 5.1 and PowerShell 7. WPF construction/event checks and pinned dependency validation pass. A controlled interactive test under a standard Windows user passed Audit and CLI Preview (exit 0), then opened the actual visual Preview and loaded the configured public WebOnly homepage with the existing WebView2 runtime. Navigation returned HTTP 200; DOM inspection and a browser-content capture confirmed rendered content and a functional `classList.add`. System/settings/external controls remained disabled. Application/configuration hashes, protected code ACL and boot state were unchanged by the tests; the temporary interactive task was removed. Both AllowApply and AllowRestore remained false. This does not qualify Citrix/AVD authentication, native clients, kiosk enforcement, or real Apply/Restore.

## Visual Preview correction - 2026-10-24

Version remains 1.1.4. GUI Preview now opens a separate visual shell, sharing XAML with the real workspace launcher; it no longer exposes only the technical JSON plan. The plan remains in CLI/evidence and the Apply confirmation. Preview does not call Launch or create an installed launcher. System actions, native clients, external browser and settings are disabled; navigation accepts HTTP/HTTPS only, reports missing URLs and labels WebOnly fallback.

63/63 offline regressions pass under PS5.1 and PS7, including target/fallback resolution, unsafe destination rejection, Hybrid profile overrides and absence of operational launch/mutation entrypoints in visual preview. WPF construction/event checks verify that the parent Preview button reaches the separate visual window, shared layout/banner, disabled controls, cancellation and Apply/Restore gate refusals. Browser navigation and visible modal interaction are not exercised by these construction checks. No real Apply/Restore was executed.

## Guarded GUI actions and Preview - 2026-10-24

Local 1.1.4 supersedes the earlier local 1.1.3 GUI simplification: Audit and Preview are separate buttons; Apply/Restore now call the existing guarded engines after administrator/configuration checks and a review dialog with exact typed confirmation. Restore requires an explicit validated rollback and rechecks its reviewed SHA256. Preview includes a read-only change plan. Defaults remain false for AllowApply and AllowRestore.

59/59 isolated regressions pass under Windows PowerShell 5.1 and PowerShell 7, including ten new GUI-dispatch/guard/rollback/preview scenarios. WPF construction and event-handler checks pass for all controls, selected profile, read-only Audit/Preview, disabled-gate refusal, exact-phrase enablement and cancellation. Apply/Restore success dispatch uses mocked engines in these tests. No real configuration-changing Apply/Restore, portal navigation, account/kiosk/startup change or interactive GUI validation is claimed.

## GUI clarity - 2026-10-24

Local version 1.1.3 removes GUI Apply/Restore buttons that only displayed CLI instructions, and describes the GUI as an audit/preview interface. The GUI CMD wrapper starts Windows PowerShell asynchronously with WindowStyle Hidden; CLI wrappers are unchanged. Guarded CLI Apply/Restore implementations are unchanged. The window title and heading are Smart ThinClient Control Panel, with the supplied WorkplaceCloudHub icon. 49/49 offline regression scenarios pass under Windows PowerShell 5.1 and 7. WPF construction checks pass for the title, icon and available controls; this is not an interactive GUI test. No operational Apply/Restore or portal connection is part of this change.

## WebOnly default - 2026-10-24

Version 1.1.2 sets WebOnlyUrl to https://workplacecloudhub.com/ in the template and built-in defaults. Existing custom URLs and explicit blanks remain authoritative. Audit/Preview do not navigate to the URL. 49/49 offline regression scenarios pass under Windows PowerShell 5.1 and PowerShell 7. They verify fresh/runtime-free defaults, existing custom URLs and explicit blanks. No Apply/Restore or portal connection was performed.

## Demo URL defaults - 2026-10-24

- Main template and built-in defaults now use the WorkplaceCloudHub Citrix and AVD demo destinations under `/demo/smartthinclient/citrix/` and `/demo/smartthinclient/avd/`. Script/test header version: 1.1.1. Provider profile templates retain blank URLs for inheritance; WebOnly is unchanged.
- Custom primary URLs are preserved. Custom legacy WorkspaceUrl/StoreUrl/FeedUrl values take precedence over a demo default, including Hybrid provider profiles. Startup continues to preserve existing runtime values, including explicit blanks. In this local working copy, only the two empty main runtime URL fields were filled; a private backup was kept and all other runtime values were verified unchanged.
- 49/49 offline scenarios pass in Windows PowerShell 5.1 and PowerShell 7, including six new scenarios for fresh defaults, template-free defaults, custom runtime/profile values, legacy aliases and explicit blanks. These use temporary fixtures and intercepted system operations, not real Apply/Restore.
- Demo-page availability, rendering, authentication and sessions were not tested. No browser, real Apply/Restore or VM operation was run. Git publication is limited to the main template, main script, regression harness, README and this validation record; the runtime configuration and local Windows 11 test sheet are excluded. The main script and regression harness receive renewed SHA-256 Authenticode signatures with timestamps, verified before commit and by committed-content read-back with Windows CRLF filters. The signed 2026-10-23 publication below remains the historical baseline.

## Local corrections - 2026-10-23

The current source corrects target-user CMD policy scope, Hybrid provider profile URLs, modern default-provider precedence, external web readiness and the CSV evidence switch. Script/test header version: 1.1.0. Validation is offline only; no live Apply/Restore or registry/account/kiosk operation was performed. Git publication is limited to the main script, regression harness, README and this validation record.

- The offline regression harness uses a private temporary source/configuration fixture seeded only from templates. It does not read or modify the operator's runtime profiles. Function declarations are loaded through the AST; registry/account/CIM writes, process launches and UI loading are intercepted.
- 43 scenarios pass under Windows PowerShell 5.1 and PowerShell 7. The 18 additional scenarios exercise provider profile URLs/global inheritance/legacy aliases, actual generated Hybrid routing with intercepted process launches, default-provider precedence, external web readiness, Strict hive prerequisites, target-SID rollback validation, loaded/unloaded hive restoration, schema 2 machine-snapshot compatibility, JSON preservation and optional CSV output, including the audit/export integration.
- Strict uses the existing target user's loaded registry hive. Missing hives are refused before Apply configuration writes or Restore registry writes. The tool does not load offline hives. Old schema 2 HKLM CMD snapshots remain supported.
- The main script and regression harness receive renewed SHA-256 Authenticode signatures with timestamps, verified on the Windows CRLF checkout before commit. Live Citrix/AVD authentication, logon, policy enforcement and kiosk behavior remain unvalidated.

## Historical candidate validation - 2026-09-09

Status: locally prepared for source review. No Smart ThinClient commit, push, release, Gallery/Intune package or OVH publication has been performed for this candidate. Public distribution remains the repository source folder.

## Demonstrated fixes

- New profile templates inherit main launch choices instead of forcing Web/embedded mode. Existing customized profile files retain precedence; use `__USE_GLOBAL__` explicitly when migrating them.
- Native process launch omits the invalid empty ArgumentList and surfaces startup errors. Hybrid offers a real provider selector; cancellation exits without opening a workspace.
- Launch writes a unique run-only launcher, preserving the file used by an installed startup command.
- AutoLaunchMode None writes no Run key. ShellRestrictionLevel None writes no policies. Invalid modes and string-valued JSON booleans are rejected.
- Rollback is saved before account creation/configuration writes. Registry snapshots preserve raw expandable values and fail on unreadable keys. They cover only the selected values, never untouched Winlogon Shell/Userinit. Previous launcher bytes are preserved.
- Apply preflight rejects incompatible kiosk options, missing prerequisites, a direct local administrator kiosk target, existing Assigned Access, and an existing custom shell for the target SID. Shell Launcher must already be enabled and use the persistent embedded shell.
- Restore validates schema 2, computer identity, output paths and the complete registry snapshot list before writes. Legacy/incomplete snapshots require manual review. CIM error codes and Assigned Access errors propagate instead of producing false success.
- Local group membership uses the built-in Users SID on localized Windows. CMD wrappers forward explicit operator arguments.

## Evidence

- `Tests/Test-SmartThinClient.Regression.ps1`: **25/25 scenarios pass in Windows PowerShell 5.1 and PowerShell 7**. The harness loads function declarations through the AST; endpoint-write primitives are mocked. It covers launch generation/routing prerequisites, escaping, inherited configuration, None options, rollback bytes/path/schema, confirmation gates, kiosk prerequisites and failure propagation, and rollback-before-account ordering.
- Two WPF XAML windows loaded under Windows PowerShell 5.1/STA without display, navigation or event execution. This verifies construction, not interactive layout or portal behavior.
- Four profile templates parse as JSON; generated launchers for Citrix, AVD, WebOnly and Hybrid parse as PowerShell.
- A real local CLI Preview completed with `PreviewOnly` evidence and Apply/Restore blocked. It wrote only local configuration, logs and evidence; readiness was partial, with no configured portal demonstrated.
- The site generator built successfully in an isolated copy. EN/FR/IT/ES/DE/AR catalogue cards use the project's manual localization engine with zero missing subject translations. The candidate changes only the Smart ThinClient article in six downloaded public pages; all surrounding page content is preserved, including canonical URLs, seven hreflang tags, GTM and Arabic RTL.

Local preparation evidence, signing/package manifests, source/public baseline hashes and exact patches are in the task's `.codex-work/thinclient-audit-20260909/` directory. They are intentionally excluded from the proposed Git file list. Local Preview evidence contains host/user information and must not be included in a public package.

## Unvalidated behavior and required lab follow-up

No real Apply/Restore, account creation, policy change, startup/logon change, Assigned Access, Shell Launcher, autologon, Citrix/AVD session, authentication or tenant deployment was executed. Do not describe this candidate as production kiosk certification or Windows 10/11 compatibility certification. PowerShell 7 testing covers isolated functions only.

The embedded WPF WebBrowser is a legacy control; current portal compatibility remains unvalidated. Prefer a separately validated external browser when needed. Shell Launcher still requires the persistent embedded process in this implementation, which limits its current suitability for modern portals. Detection of installed clients/API classes is not end-to-end readiness.

Apply writes machine-wide startup/policy values when those options are selected. An operator must verify the target user's access, directory ACLs, Windows edition/licensing, administrator recovery access and policy ownership in a disposable lab before any rollout. Existing kiosk configurations are refused rather than overwritten. Dedicated users and evidence remain after Restore. Apply is not transactional; keep the rollback file and review partial failures and successive snapshots.

Before approved publication, recheck Git HEAD/remote divergence, exact changed-file scope, package signatures/hashes, shared site generator/translation baselines and all six remote page baselines. Rebase the subject delta if another task changed any baseline. Use the existing cluster129 SFTP workflow with an exclusive publication lock and six explicit paths; do not synchronize the entire generated site. Verify public bytes, then submit the six URLs to IndexNow. An IndexNow HTTP 200 is receipt, not indexing.
