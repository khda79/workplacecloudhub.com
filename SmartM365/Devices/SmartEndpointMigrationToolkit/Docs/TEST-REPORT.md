# Test report - Smart Endpoint Migration Toolkit 1.0.0

Date: 2026-10-25. Technician computer: Windows 11 Pro 26300. Engines actually executed:
Windows PowerShell 5.1.26100.9711 (Desktop) and PowerShell 7.6.6 (Core). Source: commit 09d109ef
plus the uncommitted working tree of this delivery.

## Qualification status

| Scope | Status |
| --- | --- |
| Offline qualification (synthetic data, temporary folders, stubs only at operational boundaries) | **Done** - results below |
| Signed package (`RequireSignature=true`, pinned signer) | **Not done** - the sources are not signed yet; signing is a separate step (it must not be triggered by a commit or push from this work) |
| Several Windows sessions or accounts on one technician computer | **Not done** - no second account may be created in this scope |
| Real endpoints, PsExec, Windows Setup, reboots, AD, Microsoft Graph, Intune build/publication/assignment | **Not done** - no live access in this scope |

## Results of the final run

Each suite prints one line per assertion. Assertions are checks inside a suite run, not independent
tests: a suite run on two engines counts its assertions twice.

| Suite | Engine | Exit | Assertions passed / failed |
| --- | --- | --- | --- |
| Tests\Core\Test-SemtCore.ps1 | 5.1 | 0 | 87 / 0 |
| Tests\Core\Test-SemtCore.ps1 | 7.6.6 | 0 | 87 / 0 |
| Tests\Core\Test-SemtPowerShellDates.ps1 | 5.1 | 0 | 27 / 0 |
| Tests\Core\Test-SemtPowerShellDates.ps1 | 7.6.6 | 0 | 31 / 0 |
| Tests\GUI\Test-SemtGui.ps1 (child 5.1 STA processes) | 5.1 | 0 | 46 / 0 |
| Tests\Packaging\Test-SemtPackaging.ps1 | 5.1 | 0 | 42 / 0 |
| Tests\HybridJoin\Test-SemtHybridJoinAutomaticWave.ps1 | 5.1 | 0 | 71 / 0 |
| Tests\HybridJoin\Test-SemtHybridJoinAutomaticWave.ps1 | 7.6.6 | 0 | 71 / 0 |
| Tests\Windows11Upgrade\Test-SemtWindows11AutomaticWave.ps1 | 5.1 | 0 | 100 / 0 |
| Tests\Windows11Upgrade\Test-SemtWindows11AutomaticWave.ps1 | 7.6.6 | 0 | 100 / 0 |
| Tests\Windows11Upgrade\Test-SemtWindows11IntuneWin32.ps1 | 5.1 | 0 | 171 / 0 |
| Tests\Windows11Upgrade\Test-SemtWindows11IntuneWin32.ps1 | 7.6.6 | 0 | 171 / 0 |
| Tests\Windows11Upgrade\Test-SemtWindows11MediaResolution.ps1 | 5.1 | 0 | 106 / 0 |
| Tests\Windows11Upgrade\Test-SemtWindows11MediaResolution.ps1 | 7.6.6 | 0 | 106 / 0 |
| Tests\Core\Test-SemtProgramFolder.ps1 | 5.1 | 0 | 91 / 0 |
| Tests\Core\Test-SemtProgramFolder.ps1 | 7.6.6 | 0 | 91 / 0 |
| Tests\Windows11Upgrade\Test-SemtWindows11MediaAdoption.ps1 | 5.1 | 0 | 30 / 0 |
| Tests\Windows11Upgrade\Test-SemtWindows11MediaAdoption.ps1 | 7.6.6 | 0 | 30 / 0 |
| Release\SmartM365-EndpointMigration-Update-ConfigTemplates.ps1 -Check | 5.1 / 7.6.6 | 0 / 0 | templates in sync |
| GUI -ValidateOnly (STA) | 5.1 | 0 | both Campaigns loaded (31 and 51 options) |
| Release builder preview (-AllowUnsigned) | 5.1 | 0 | 92 files validated, 52 PowerShell files unsigned |

Total: 1458 assertions over 18 suite runs of the new product, 0 failure. These are offline
assertions on synthetic data: they do not replace end-to-end tests, and the absence of end-to-end
tests leaves a regression risk on real devices (see "Remaining limits").

The media resolution suite starts the real installer, runner and generated detection script in
child Windows PowerShell 5.1 processes in both rows; the 7.6.6 row runs its in-process parts on
PowerShell 7.

Regression of the previous toolkits (unchanged in this lot): Windows 11 `Test-SmartM365Windows11DirectSetupRebootCmd` 7/0,
`Test-SmartM365Windows11RunGuardAndCmd` passed, Windows 11 demo 184 checks, Hybrid Join demo 160
checks, the six Hybrid Join suites passed (exit 0).

`SmartM365\Devices\IntuneHybridJoinToolkit\Tests\Test-SmartM365IntuneHybridJoinAutomaticLot.ps1` of the **previous** Hybrid Join
toolkit is intermittent (it failed again in this final run, and earlier in 2 of 3 reruns, always on the same
assertion "wrapper refresh is reported to the progress popup"; it passed in earlier runs). Cause
(code reading): its engine reports the stage "Refreshing LOT command wrappers..." only while the
wrapper refresh process runs longer than 500 ms (`WaitForExit(500)` loop), and the test requires
that stage. Pre-existing, unrelated to this delivery, not critical; left unchanged (the previous
toolkits are out of scope of this lot).

## What the new suites prove

- **Scripts run as SYSTEM and data folders** (`Test-SemtProgramFolder.ps1`, plus the Intune part of
  the media resolution suite):
  - Shared region: identical in the 5 copies.
  - Trust rules:
    - real checks: Program Files is trusted; a folder of the test account is `Untrusted`; a
      missing folder is `Unreadable`; a junction is refused;
    - rule-level checks on synthetic security descriptors: owner; write, delete and
      WRITE_DAC entries; generic rights; inherit-only CREATOR OWNER; deny entries; files found
      with `-Recurse`.
  - Program Files chain: created components get owner Administrators and the protected DACL; an
    existing trusted component is left unchanged; an existing untrusted one is never repaired
    and stops the operation. Not elevated, the real owner change is refused and the created folder
    is removed again.
  - ProgramData chain (privileged operations simulated: the owner change is replaced by a real
    DACL-only replacement on temporary folders, trust states are synthetic):
    - scope: only the SEMT components (product folder, then campaign folder) are created or
      replaced; the shared parents `WCH` and `WCH\SmartM365` are never changed; sibling fixtures
      (another product, a legacy cache reused in place, a sibling with an explicit entry, all with
      inheritance active) keep owner, DACL and content after creation, after a real DACL
      replacement propagated below the product folder, and after every refusal;
    - an unprotected shared parent stops the operation with nothing changed; rule-level checks of
      the read-only shared parent rule (create allowed; delete, delete child, WRITE_DAC, modify
      or a standard owner refused; unreadable never trusted);
    - provenance: an untrusted state file, setup media with its own manifest in a folder another
      account could modify, an unprotected SEMT folder with content, and untrusted content of
      another SEMT campaign in the propagation scope are refused with zero ACL change, content and
      DACL unchanged; never reset or adopted;
    - reproduced case: protected parents, untrusted file listed before a junction: refused with
      zero ACL change (file and junction target unchanged); junction alone: refused;
    - real hard links (data root, other campaign) and real junctions in the propagation scope are
      refused before any change;
    - content written into a folder between the read-only pass and its protection (created here
      or unprotected and empty) is detected and refused;
    - blocking failure when the owner or DACL cannot be replaced; the read-only verification
      detects an untrusted item, SEMT component or shared parent;
    - real, not elevated: a shared folder created by the test account is refused and removed.
  - Both real endpoints, started with a data root outside ProgramData, print
    `FINAL_STATUS=DATA_FOLDER_UNTRUSTED` (exit 1) and write nothing.
  - Retry after reboot (Windows 11 and Hybrid Join): runner, state and endpoint copy are under
    Program Files; the task points there; nothing is under ProgramData; an untrusted folder
    blocks registration and the task run; cleanup removes the copy.
  - PsExec: the endpoint is copied under `C:\Program Files\...\<Campaign>` with the expected
    SHA256; the PsExec path points there; no script is under ProgramData; an untrusted folder
    refuses the copy; the launcher contract carries `RemoteProgramRoot` and the worker refuses a
    contract without it.
  - Intune installer (real script, shimmed boundaries):
    - scripts are installed under Program Files with the package hashes; no script stays under
      ProgramData (the old copy is removed); the task starts the runner from Program Files;
    - every created program folder and the data root get the protected ACL before any write;
    - `-Uninstall` removes the program files;
    - an untrusted Program Files folder or ProgramData chain gives exit 1 with nothing written,
      installed or registered; not elevated, the real checks refuse as well.
  - Intune runner:
    - with a trusted folder, the endpoint is started from Program Files;
    - with an untrusted, unreadable or user-owned program folder (real check), it exits 1 and
      loads, starts and writes nothing;
    - with an untrusted ProgramData chain, nothing is started or written, the package is
      marked `RepairRequired` (`SEMT_DATA_FOLDER_UNTRUSTED`) and the task is removed.
  - Mutations (each applied alone, files restored and hashes checked afterwards): shared parents
    treated as SEMT components: 10 failures (siblings and scope); existing-content provenance
    check removed: 6; link pre-check removed: 2 (junctions stay refused by the trust check);
    emptiness check after protection removed: 2. Earlier lot: a trust check that always says
    "trusted" made 12 assertions fail.

- **Media resolution (PsExec and Intune)**: identical shared region in the endpoint and the Intune
  helper; decisions for new cache, legacy Intune-owned, adoptable, absent, corrupt, other language,
  other media build, in use (live lock and live process found by the real process scan),
  indeterminate (unreadable lock or ownership), untrusted (writable by a non-administrator,
  reparse point) reported as its own decision `Untrusted` and never as `Absent`, unreadable
  permissions or folder tree reported as `Indeterminate` (real ACLs on temporary folders: a folder
  without READ_CONTROL and one without list access, restored afterwards; plus a missing, unknown or
  failing trust result), `LastRun.json` paths escaping the legacy `SetupMedia` folder (`..` and sibling
  prefix); endpoint: selected path used, no copy when reusable, Intune-owned media unchanged
  (content, dates, DACL, no marker), unknown path rejected, duplication blocked before any change
  and allowed only explicitly; untrusted legacy media: `SETUP_MEDIA_LEGACY_UNTRUSTED` without
  authorization (also in existing-media-only mode), one logged copy with
  `-AllowSetupMediaDuplication`, legacy media unchanged (content, dates, DACL); unreadable
  permissions: `SETUP_MEDIA_STATE_INDETERMINATE`, no copy even with the authorization; installer (real script, real robocopy): no copy for reusable media,
  external path recorded, packaged copy only when needed, `1618` when in use, duplication blocked,
  WithCacheOnly failure without reusable media; untrusted legacy media with the real ACL check:
  `SEMT_SETUP_MEDIA_LEGACY_UNTRUSTED` (WithMedia and WithCacheOnly), no copy by default, copy
  only for a package built with `-AllowSetupMediaDuplication`, never an ACL change on the legacy
  media; unreadable permissions (synthetic and real): `SEMT_SETUP_MEDIA_STATE_INDETERMINATE`, no
  copy even when duplication is allowed; runner (real script): `-SetupMediaPath`
  propagated, no `RepairRequired` for external media even after an endpoint failure (no repair
  loop), repair only without usable media, cycle skipped when in use, untrusted legacy media (real
  ACL check, even Intune-owned) not used (`SEMT_MEDIA_UNTRUSTED`, endpoint not started), cycle
  skipped when permissions are unreadable; generated detection: external
  media detected as installed. A mutation that restores the previous runner/installer behaviour
  makes 7 assertions fail; restoring the previous resolution region (untrusted media reported as
  `Absent`, unreadable permissions as untrusted) makes 29 assertions fail.
- **Setup media manifest**: the real endpoint `Test-SetupMedia -ValidateManifest` accepts media
  with a valid manifest and refuses media without one (the `$false` result is checked); Intune: a
  new cache without its manifest is `Corrupt`; a required package reference that is missing
  refuses the media (`Corrupt`, comparison never skipped); a replaced file with a manifest
  regenerated next to it is valid on its own manifest only and refused against the package
  reference. Mutations: endpoint result ignored again: 1 failure; missing reference skipped again:
  1 failure; the explicit missing-manifest check in the Intune validator is a second barrier (its
  removal makes no test fail, because the integrity check that follows refuses the media too).
- **Reversible adoption** (real moves, real `icacls`, distinctive explicit permissions): faults
  before the move (destination exists, owned, ownership unreadable, in use, invalid, reparse point,
  permissions cannot be saved), during the move (error before and after moving), during the
  permission reset (partially applied), during the validation after the move (after the real
  reset), and during the rollback (move back denied, permission restore denied, legacy folder
  recreated meanwhile). Each case compares paths, content hashes and DACLs before/after and the
  reported outcome (`NotMoved`, `RolledBack`, `RollbackPartial`, `RollbackFailed`).
- **Dates and engines**: UTC and offset dates, nested values, already deserialized
  DateTime/DateTimeOffset, process start comparison, active and stale reservations, pending markers;
  on 7.6.6 the path without `-DateKind` is forced (SIMULATED absence, real PowerShell 7 date
  conversion). A mutation removing the normalization makes 6 assertions fail on 7.6.6, including a
  live launcher seen as dead and an active reservation removed. Endpoint and GUI refusal executed
  on PowerShell 7 (no state written); module refusal of 7.1 executed with a simulated version table.

## Static validation

| Check | Result |
| --- | --- |
| Parser, Windows PowerShell 5.1 and PowerShell 7.6.6 | 58 PowerShell files of the delivery, 0 parse error on each engine |
| ASCII scan (scripts, modules, manifests, CMD) | 67 files, 0 non-ASCII byte |
| `SmartM365\Tests\Test-SmartM365ScriptStandards.ps1 -Scope Changed` | 0 error, 1 warning (`LoggingWithoutSharedInitialization`, Hybrid Join endpoint, inherited) |
| PSScriptAnalyzer (Error, Warning) on the files of this lot | 0 error; 9 warnings, all inherited unchanged from the copied sources (`$matches`, `$profile`, `$args` in the Windows 11 endpoint and the Intune publisher) |
| `git diff --check` (tracked files) | clean; untracked folder: no trailing whitespace and no mixed line endings |
| Git scope | only the new folder (plus the critical fixes of the previous lot); no `*.local.json`, Waves, Runs, inventories, logs, media, DemoAssets or memory files |

## Qualification package

`Release\Output\SmartEndpointMigrationToolkit-1.0.0.zip` (git-ignored), **UNSIGNED QUALIFICATION
BUILD**, not for production:

The SHA256 of the ZIP and of the distribution manifest are in
`Release\Output\SmartEndpointMigrationToolkit-1.0.0.sha256` (this report is inside the ZIP, so it cannot
carry the ZIP hash). The manifest lists 92 files.

Extracted outside the repository into a new folder whose name contains spaces and non-ASCII
characters (`Recette SEMT éè 中文 ...`): 93 files, no private or operational file, every manifest
hash verified; engine `-ValidateOnly` on 5.1 (Windows 11, dry run) and on 7.6.6 (Hybrid Join) with
17/17 components verified against the shipped manifest; GUI `-ValidateOnly` on 5.1 STA; the media
resolution suite passes from the extracted copy; no code reference to the previous product folders.

## DemoAssets

The GUI file changed (host check only). Captures and exports regenerated in a temporary copy are
byte-identical to the 2026-10-24 set (16 files), so the set and its capture date are kept; the
revalidation (GUI SHA256, 2026-10-25) is recorded in the local manifests.

## Remaining limits

- **Local privilege escalation through ProgramData (reported 2026-10-25): corrections made and
  verified offline only; not declared fixed.** The evidence is limited to offline tests on
  temporary folders with simulated privileged operations; it is not sufficient to conclude that
  the issue is fully fixed.
  - Verified offline:
    - scripts run as SYSTEM (PsExec endpoint, Intune runner, retry tasks) are under
      `C:\Program Files\WCH\SmartM365\SmartEndpointMigrationToolkit\<Campaign>`;
    - only the SEMT components of the ProgramData chain are changed; shared parents and their
      other children are left unchanged (simulated DACL change, sibling fixtures);
    - a read-only pass refuses links, reparse points and content without verified provenance
      before any change; a permission replacement no longer makes existing content trusted;
    - media without its manifest, or without a required package reference, is refused;
    - the runner verifies both folders before running anything.
  - Not verified, to qualify on lab devices:
    - the real owner change and protected DACL by SYSTEM and by the technician account through
      `C$` (needs elevation; simulated here);
    - real shared parents `C:\ProgramData\WCH` and `...\SmartM365` as found on devices: if other
      products gave Users delete or permission rights on them, SEMT stops
      (`SEMT_DATA_FOLDER_UNTRUSTED`) until an administrator corrects them; how often this happens
      is unknown;
    - a real chain pre-created by a standard user (refused with nothing changed; an
      administrator must review and remove that content);
    - the race window between the read-only pass and the DACL replacement: content written there
      is detected afterwards and refused, but a handle opened before the protection is not
      covered by these tests;
    - the duration of the link pre-check (fsutil for each file) on a large media cache when a
      replacement is planned;
    - end-to-end runs (Intune install, runner, PsExec, Windows Setup): none were run, so
      regressions on real devices remain possible.

- Not qualified live: PsExec copy and execution, endpoint repair, Windows Setup and reboots,
  AD/Graph exports, Intune packaging, publication and assignment.
- Not qualified: the signed package; several Windows sessions or accounts sharing one technician
  computer (the shared cap was exercised across processes of one account), and the
  `LauncherState` ACL hardening, which applies only when elevated.
- Coordination is per technician computer only.
- Legacy media is reused or adopted only when it is write-restricted (owned by and modifiable only
  by SYSTEM, Administrators, TrustedInstaller): legacy folders that inherited the default
  `ProgramData` permissions (Users may create files) are `Untrusted`. They are never used and their
  permissions are never changed; a second copy then requires `AllowSetupMediaDuplication`
  (otherwise `SETUP_MEDIA_LEGACY_UNTRUSTED`). How often real devices of the previous toolkit are in
  this state must be checked during acceptance, because it decides whether duplication must be
  allowed for them.
- Without a trusted reference manifest (PsExec, Intune WithCacheOnly), media is validated
  against its own manifest, which is not a proof of authenticity on its own; protection then
  relies on the write-restriction and provenance checks.
- Media validation in the tests uses synthetic media; the endpoint checks with Authenticode and
  DISM on real media (`Test-SetupMedia`) were not run.
- Intune downloads a WithMedia package (media included) before its installer runs; a reusable local
  cache only avoids the local copy. Use WithCacheOnly for devices with reusable media.
- PowerShell 7.2 to 7.4 were not installed: their path (no `-DateKind`) was executed on 7.6.6 by
  forcing it.
