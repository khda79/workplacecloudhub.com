# Coexistence with the previous toolkits

Smart Endpoint Migration Toolkit is a new product. IntuneHybridJoinToolkit and
Windows11UpgradeToolkit stay installed and usable; they receive no further functional change.
A critical defect found in them is fixed in the previous product first, then ported here
(see [CRITICAL-FIXES-LEGACY.md](CRITICAL-FIXES-LEGACY.md)).

## What is not migrated

Configurations, LOTs, `Computers.txt`, inventories, run history, run guard history, reports and
logs of the previous toolkits are **not** imported. Create new Waves (New Wave or Automatic
Wave) and new configuration from the templates.

The only data reused is the Windows 11 setup media already present on targets (below).

## Separate identities

| Item | Previous toolkits | Smart Endpoint Migration Toolkit |
| --- | --- | --- |
| Technician state | `C:\ProgramData\SmartM365\<Toolkit>\...` | `C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\LauncherState` |
| Target data | `C:\ProgramData\SmartM365\IntuneHybridJoinToolkit`, `...\Windows11UpgradeToolkit` | `C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\<Campaign>` |
| Environment variables | `EHJIR_*` (Hybrid Join), `W11UT_*` | `SEMT_*`, `SEMT_HYBRIDJOIN_*`, `SEMT_WINDOWS11_*` |
| Endpoint mutex | `Global\SmartM365_IntuneHybridJoinToolkit_Endpoint` | `Global\SmartM365_SEMT_HybridJoin_Endpoint`, `Global\SmartM365_SEMT_TargetOperation` |
| Retry-after-reboot tasks | `SmartM365-IntuneHybridJoinToolkit-RetryAfterReboot`, `SmartM365-Windows11UpgradeToolkit-RetryAfterReboot` | `SmartM365-SEMT-HybridJoin-RetryAfterReboot`, `SmartM365-SEMT-Windows11Upgrade-RetryAfterReboot` |
| Intune Win32 (Windows 11) | `HKLM\SOFTWARE\SmartM365\Windows11UpgradeToolkit\IntunePackages`, task `SmartM365 Windows 11 Upgrade Toolkit - Intune` | `HKLM\SOFTWARE\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade\IntunePackages`, task `SmartM365 SEMT Windows 11 Upgrade - Intune`, PackageId prefix `SmartM365-SEMT-Windows11Upgrade-` |

## Legacy activity detection on the target

Before acting, both endpoint scripts inspect the previous toolkits **read-only**: they never
write, delete, lock, wait on a legacy mutex, stop a process or change a task or registry key.
The mere presence of a previous toolkit does not block.

| Status | Blocking | Meaning | Next action |
| --- | --- | --- | --- |
| `LEGACY_ABSENT` | no | No trace of the previous toolkits | |
| `LEGACY_INSTALLED_IDLE` | no | Folders, tasks or packages exist; nothing is running | |
| `LEGACY_INTUNE_PACKAGE_ASSIGNED` | no | A legacy Windows 11 Intune Win32 package is installed; its media is reused in place | `REVIEW_LEGACY_INTUNE_ASSIGNMENT` |
| `LEGACY_ENDPOINT_RUNNING` | yes | A legacy endpoint script, retry runner or task is running, or the legacy Hybrid Join mutex is held | `WAIT_LEGACY_TOOLKIT_OPERATION` |
| `LEGACY_SETUP_IN_PROGRESS` | yes | Windows Setup runs from legacy media | `WAIT_LEGACY_WINDOWS_SETUP` |
| `LEGACY_COPY_IN_PROGRESS` | yes | Robocopy into legacy media, or a legacy media lock held by a live process | `WAIT_LEGACY_MEDIA_COPY` |
| `LEGACY_RETRY_AFTER_REBOOT_PENDING` | yes | A legacy retry task will resume a legacy run at startup | `WAIT_OR_REMOVE_LEGACY_RETRY_TASK` |
| `LEGACY_STATE_UNREADABLE` | yes | Legacy state could not be read: the state is indeterminate | `VERIFY_LEGACY_TOOLKIT_STATE` |

The option `LegacyActivityPolicy` (`Block` by default, or `Warn`) controls the reaction to a
blocking status; `Warn` is meant for an operator who verified the legacy state by hand. The status and its detail are in every result row (`LegacyActivityStatus`,
`LegacyActivityDetail`).

Hybrid Join also stops with `TARGET_WINDOWS_SETUP_ACTIVE` (`WAIT_WINDOWS_SETUP_COMPLETION`)
while any Windows Setup is running.

## Windows 11 setup media of the previous toolkit

The same read-only resolution is used by the endpoint (PsExec Runs), by the Intune installer and
by the Intune runner, before any copy. It looks at the cache of this product and at
`C:\ProgramData\SmartM365\Windows11UpgradeToolkit\SetupMedia\<MediaId>-<Language>` (and the path
recorded in the legacy `LastRun.json`, accepted only when it normalizes to a folder under the legacy `SetupMedia` folder).

| Decision | Situation | What happens |
| --- | --- | --- |
| `NewCacheValid` | Valid media in the cache of this product | Used; nothing copied. |
| `LegacyIntuneOwned` | Valid legacy media owned by a legacy Intune Win32 package (registry `SetupCacheFolder`, state not uninstalled) | Used in place, read-only, recorded as external (`SetupMediaSource=LegacyExternal`). Never moved, cleaned, replaced, re-permissioned or marked. |
| `LegacyAdoptable` | Valid, idle, unmanaged legacy media, same volume, no reparse point, new cache folder absent | Moved by the endpoint to the new cache with full rollback (below). |
| `InUse` | A live process runs from or copies into the media, or a live media lock exists | Nothing is changed: `SETUP_MEDIA_IN_USE` (endpoint), Intune exit `1618`, runner cycle skipped. |
| `Indeterminate` | Ownership, processes, locks, permissions or the folder tree of the legacy media could not be read | Nothing is changed, even when duplication is allowed: `SETUP_MEDIA_STATE_INDETERMINATE`, installer exit `1`, runner cycle skipped. |
| `Untrusted` | Legacy media present, but a non-administrator owns or can modify it, or it contains a reparse point | Never used, because its `setup.exe` would run as SYSTEM; its permissions are never changed. A second copy needs the duplication authorization: otherwise `SETUP_MEDIA_LEGACY_UNTRUSTED` (endpoint) or `SEMT_SETUP_MEDIA_LEGACY_UNTRUSTED` (installer); the Intune runner marks the package `RepairRequired` (`SEMT_MEDIA_UNTRUSTED`). |
| `Incompatible` | Other language, or for an Intune WithMedia package another media build | Normal copy path (source share or packaged media). |
| `Corrupt` / `Absent` | Invalid media / no media of this product and no legacy media found | Normal copy path. |

When valid legacy media remains but is not used (for example the new cache folder already exists,
another volume, Intune-owned media that differs from the selected source), a second copy is blocked:
`SETUP_MEDIA_DUPLICATION_BLOCKED` (endpoint) or `SEMT_SETUP_MEDIA_DUPLICATION_BLOCKED` (installer),
unless duplication is explicitly allowed (`Windows11Upgrade.AllowSetupMediaDuplication`, or an
Intune package built with `-AllowSetupMediaDuplication`); the allowed second copy is logged. The
free-disk check runs before any copy.

### Reversible adoption

1. Checks before any change: full source and destination paths, destination absent, same volume,
   no reparse point in the source tree or the destination parents, not owned by a legacy package,
   not in use, media valid.
2. Restore information: relative path, size and DACL of every item (`snapshot-before.json.txt`) and an
   `icacls /save` file (`acl-before.icacls`).
3. Same-volume move, then permission reset to the inheritance of the new cache, checked on every
   item (not only from the `icacls` exit code), then validation in place.
4. On a failure after the move: the folder is moved back (never over a folder recreated meanwhile),
   the recorded DACLs are re-applied item by item, and location, sizes and permissions are compared
   with the snapshot (same protection flags and the same explicit and inherited entries).

| Outcome | Meaning |
| --- | --- |
| `Moved` | Media in the new cache, validated. |
| `NotMoved` | A check failed before any change; nothing was modified. |
| `RolledBack` | Moved back and verified identical (location, sizes, permissions). |
| `RollbackPartial` | Moved back, but the verification found differences (usually permissions): reported, `SETUP_MEDIA_LEGACY_ADOPTION_FAILED`. |
| `RollbackFailed` | Not moved back (move-back error, or legacy folder recreated meanwhile). The reported path is where the media really is; nothing was deleted or overwritten. |

Evidence (steps, snapshots, saved permissions, outcome) stays in
`<DataRoot>\State\MediaAdoption\<timestamp>\`.

`Windows11Upgrade.ReuseLegacySetupMedia=false` disables discovery and adoption. Disk cleanup of this
product never touches the legacy media folder.

### Intune packages

- **WithCacheOnly** applies when reusable media exists (cache of this product or legacy media,
  owned or adoptable). It copies nothing; without reusable media it fails explicitly
  (`SEMT_SETUP_MEDIA_NOT_REUSABLE`).
- **WithMedia** copies its packaged media only when no reusable media exists. Intune downloads the
  package content, media included, before the installer runs: reusing a local cache avoids the local
  copy, not that download. Prefer WithCacheOnly for devices that already have legacy media.
- Follow the assignment switch order in
  [IntuneWin32/README.md](../Campaigns/Windows11Upgrade/IntuneWin32/README.md).
## Running both products

- Do not run the previous toolkit and this product against the same computers at the same time.
  The target-side detection blocks the new product while a legacy operation is active, but the
  previous toolkits do not know about the new product.
- Hybrid Join and Windows 11 Waves of this product can run at the same time: the target guard
  allows one operation per computer and a pending Windows 11 setup blocks Hybrid Join until reboot.
- The technician-side shared worker cap covers both Campaigns of this product only.
