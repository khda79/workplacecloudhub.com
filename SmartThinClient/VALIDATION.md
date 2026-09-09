# Smart ThinClient candidate validation — 2026-09-09

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
