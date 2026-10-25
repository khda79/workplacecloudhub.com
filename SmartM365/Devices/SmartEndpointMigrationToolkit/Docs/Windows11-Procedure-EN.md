# Windows 11 upgrade procedure (Smart Endpoint Migration Toolkit)

French version: [Windows11-Procedure-FR.md](Windows11-Procedure-FR.md). Keep both versions in
step.

## 1. Prepare the setup media

1. Extract licensed Windows 11 media, one folder per language, on a share readable by the
   **computer accounts** of the targets (the endpoint runs as SYSTEM), for example
   `\\files-paris\Windows11\24H2\FR-fr`.
2. Generate the SHA256 manifest of each media folder (from the package root):
   `New-SetupMediaManifest.cmd` (local `Campaigns\Windows11Upgrade\SetupSource`) or
   `Campaigns\Windows11Upgrade\Scripts\SmartM365-EndpointMigration-New-SetupMediaManifest.ps1 -MediaRoot <folder> -Force`.
   The manifest `SmartM365-SetupMediaManifest.sha256.csv` stays at the root of the media folder.
3. Several sites: list several UNC sources in `Windows11Upgrade.SetupSourcePath` (separated by
   `;`), or better a map CSV in `Windows11Upgrade.SetupSourceMapPath`
   (template `Campaigns\Windows11Upgrade\Templates\SetupSourceMap.csv.template`; scopes
   `Subnet`/`CIDR`, `IPPrefix`, `ComputerName`, `ComputerPrefix`, `Default`). Each target picks
   the nearest valid source.
4. Optional copy gates: `SetupSourceConcurrencyLimit` with `SetupSourceConcurrencyGateRoot`,
   and `SetupSubnetConcurrencyLimit` (default 1 copy per subnet) with
   `SetupSubnetConcurrencyGateRoot`. Gate roots are UNC folders where target computer accounts
   have Modify (see `Campaigns\Windows11Upgrade\SetupSourceGates\README.md`).

## 2. Configure

In the GUI, select **Windows 11 upgrade**, then **Options**:

| Group | Main options |
| --- | --- |
| Mode | Dry run, Remote audit only, run guard |
| Actions | Policy repair, WU reset, forced WU upgrade, setup upgrade, reboots, profile repair, disk cleanup |
| Setup | Source path or map, execution mode (`LocalCache` default, `Share`, `Auto`), media id, language (`MatchSystem`), Dynamic Update, legacy media reuse, duplication |
| Parallelism / Safety | workers per Wave, shared cap, PsExec timeout, run guard hours, minimum free disk (32 GB) |

Global defaults can also be set once in `Config\SmartEndpointMigrationToolkit.local.json` and per
Wave in `Wave.local.json` (for example a site-local source per Wave).

Reboots: `Allow controlled reboots` (default on) reboots for a pending reboot or setup completion
only when no user is connected, **except** when uptime exceeds
`ForceRequiredRebootWhenUptimeOverDays` (7 days by default; 0 = never). With DirectSetup and
`Allow controlled reboots` off, no reboot is ever requested.

## 3. Prepare the Wave

- **Automatic Wave**: refresh the shared AD and Intune inventories, choose the source
  (`Both`, `AD`, `Intune`), prefixes, AD logon age and *Exclude devices present in Intune* if
  needed, click **Preview**, review the selection and exclusions, then **Create Wave**.
- **New Wave**: paste the computer names, **Create Wave**.
- Devices already on Windows 11 in AD or Intune are skipped at each cycle and archived in
  `Computers.AlreadyWindows11.txt`; confirmed not-capable hardware goes to
  `Computers.HardwareNotCapable.txt`.

## 4. Qualify, then run

1. **Dry run** (Once): DNS, ping, SMB 445 and administrative share for every computer. Nothing
   is copied.
2. **Remote audit only** (Once): readiness, hardware, disk, media and source reachability
   without changes.
3. **Repair** (Loop): the enabled actions run; setup media is copied by the target itself from
   the selected source (not by the technician computer), validated, then Windows Setup starts.
4. Follow the **Runs** tab: live HTML report, status and next action per computer, evidence in
   `CentralLogs`.

## 5. Media left by the previous Windows 11 toolkit

Before any large copy, the same resolution is used for PsExec Runs and for the Intune package:

- valid media in the cache of this product: used;
- legacy media owned by an installed legacy Intune Win32 package: reused **in place**, read-only,
  external to this product (never moved, cleaned, re-permissioned or marked);
- idle, unmanaged, valid legacy media: **moved** on the same volume to the new cache; on any failure
  it goes back to its location with its permissions re-applied and verified (`RolledBack`), or the
  failure is reported explicitly (`RollbackPartial`, `RollbackFailed`) without losing data;
- media in use or in an unreadable state (including permissions or a folder tree that cannot be
  read): nothing is changed (`SETUP_MEDIA_IN_USE`, `SETUP_MEDIA_STATE_INDETERMINATE`);
- legacy media present but modifiable by non-administrators: never used and its permissions are
  never changed (`SETUP_MEDIA_LEGACY_UNTRUSTED`);
- otherwise (absent, invalid, other language or build): normal copy. A second copy next to legacy
  media that remains (valid or untrusted) requires `AllowSetupMediaDuplication`
  (`SETUP_MEDIA_DUPLICATION_BLOCKED` or `SETUP_MEDIA_LEGACY_UNTRUSTED` otherwise).

Scripts that run as SYSTEM (the endpoint started by PsExec, the retry-after-reboot task, the Intune
runner) are stored under `C:\Program Files\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade`;
a program folder that is not protected stops the operation (`SEMT_PROGRAM_FOLDER_UNTRUSTED`).

Evidence of every adoption: `<DataRoot>\State\MediaAdoption\`. Details:
[LEGACY-COEXISTENCE.md](LEGACY-COEXISTENCE.md).

## 6. Stop

- **Controlled stop** (Runs tab or *Stop running Waves*): no new computer starts.
- A second request **forces the local stop** of workers and PsExec.
- A target may still be running Windows Setup: check its row
  (`CANCELLED_BY_OPERATOR_REMOTE_RUNNING` / `_REMOTE_STATE_UNCONFIRMED`) before relaunching.

## 7. Intune Win32 alternative

For devices out of PsExec reach, build the per-language Intune Win32 packages
(`Campaigns\Windows11Upgrade\IntuneWin32\README.md`). Use the WithCacheOnly package for devices that already have reusable media (it downloads no media) and the WithMedia package otherwise (Intune downloads its media before the installer runs). Follow the assignment switch order in that
README when replacing an assignment of the previous toolkit package: never let both packages
target the same devices at the same time.

## 8. Common statuses

| Status | Meaning / action |
| --- | --- |
| `ALREADY_WINDOWS11` | Skipped, archived |
| `WINDOWS11_HARDWARE_NOT_CAPABLE` | Confirmed hardware ineligibility; archived |
| `SETUP_UPGRADE_STARTED`, `SETUP_UPGRADE_REBOOT_*` | Setup is running or waits for its reboot; the target blocks the other Campaign until reboot (168 h at most) |
| `SETUP_MEDIA_DUPLICATION_BLOCKED` | Legacy media present and not reusable; unassign the legacy Intune app, free the media or allow duplication |
| `SETUP_MEDIA_LEGACY_UNTRUSTED` | Legacy media present but modifiable by non-administrators (or containing a reparse point); not used and its permissions not changed. Secure or remove it (administrator decision), or allow duplication |
| `SETUP_MEDIA_STATE_INDETERMINATE` | Ownership, use, permissions or folder tree of the media could not be read; nothing changed. Check the device, then retry |
| `TARGET_OPERATION_ACTIVE` | Another toolkit operation runs on the target; next cycle |
| `LEGACY_*` | See [LEGACY-COEXISTENCE.md](LEGACY-COEXISTENCE.md) |
