# SmartM365 Windows 11 Upgrade Toolkit

PowerShell LOT/PsExec toolkit to diagnose Windows 10 devices that should migrate to Windows 11 through Intune, Windows Autopatch, or Windows Update for Business, then run guarded repair or upgrade actions from LOT folders.

The toolkit is optimized for batch operations. It keeps `Scripts\SmartM365-Invoke-Windows11UpgradeRepair.ps1` autonomous so the same single PowerShell file can be copied to devices through PsExec/LOT folders or reused by another deployment mechanism without requiring SmartM365 modules on the target computer.

## Layout

```text
Start-Windows11UpgradeRepair-LotLauncher-GUI.cmd
Update-LotCmdWrappers.cmd
Scripts\
LOT-X\
```

- `Start-Windows11UpgradeRepair-LotLauncher-GUI.cmd` opens a GUI to launch existing operational `LOT-*` folders or create a new empty LOT.
- `Update-LotCmdWrappers.cmd` refreshes the small CMD wrappers in operational `LOT-*` folders.
- `Scripts\SmartM365-Invoke-Windows11UpgradeRepair.ps1` is the autonomous endpoint-side diagnostic and guarded action script.
- `Scripts\SmartM365-Invoke-Windows11UpgradeRepairWithPsExec.ps1` is the local PsExec orchestrator for LOT folders.
- `Scripts\Run-Windows11UpgradeRepairWithPsExec-Lot.cmd` is the shared CMD launcher used by every LOT wrapper.
- `Scripts\SmartM365-Windows11Upgrade-LotLauncher-GUI.ps1` is the WinForms LOT launcher.
- `LOT-X\` is the neutral versioned template for LOT folders.

Operational `LOT-*` folders are ignored by Git because they can contain real computer names, logs, reports, and cached operational context.

## When To Use This Toolkit

Use this toolkit when:

- devices are expected to be Intune managed and upgraded to Windows 11 through Autopatch or WUfB;
- you need to process a batch of Windows 10 computers from an operator workstation;
- you need local evidence explaining why Windows 11 is not installing;
- you need PsExec-based SYSTEM execution on remote devices;
- you need an operator-controlled path for setup-based upgrade through `-AllowSetupUpgrade`.

Use Smart Intune Remediation packages instead when the target deployment mechanism should be recurring Intune remediation. Use Smart Endpoint Diagnostics Analyzer when you already have Intune Device Diagnostics ZIP files and need a support report.

## Usage

Open the LOT launcher GUI:

```cmd
Start-Windows11UpgradeRepair-LotLauncher-GUI.cmd
```

Run a LOT once:

```cmd
LOT-X\Run-Windows11UpgradeRepairWithPsExec-Once.cmd
```

Run diagnostic only from the PowerShell orchestrator:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\SmartM365-Invoke-Windows11UpgradeRepairWithPsExec.ps1 `
  -ComputerListPath .\LOT-X\Computers.txt `
  -RunOnce `
  -AuditOnly
```

## Default LOT Behavior

LOT CMD wrappers are action-ready by default, including policy repair, Windows Update reset,
force upgrade, setup upgrade, controlled reboot, safe disk cleanup, and virtual machine
skipping. Advanced disk cleanup is available but disabled by default. The wrappers also
require an explicit setup source when setup media must be copied to a target. For LOT/PsExec,
use a UNC path reachable by the target computers:

When setup upgrade is enabled, the endpoint attempts setup upgrade after readiness blockers
are cleared. Windows Update reset and assigned update force trigger remain fallback actions
for runs where setup upgrade is disabled.

```text
\\server\share\Windows11
```

The repository `SetupSource` folder is for local/direct tests or for preparing media before
publishing it to a network share. See `SetupSource\README.md`.
By default, setup language validation is enabled with `W11UT_SETUP_LANGUAGE=MatchSystem`.
The target compares its installed Windows language with `sources\lang.ini` from the setup
media before starting setup.
`SetupSourcePath` can point either to a single media root or to a parent folder containing
one media subfolder per language; in `MatchSystem` mode the toolkit selects the subfolder
whose `sources\lang.ini` matches the target language.
By default, Windows Setup Dynamic Update is disabled with
`W11UT_SETUP_DYNAMIC_UPDATE=Disable` so large LOT runs stay controlled and do not trigger
per-device online setup downloads. Supported values are `Enable`, `Disable`, `NoDrivers`,
`NoLCU`, and `NoDriversNoLCU`.

To disable an action for a direct LOT launch, set its environment variable to `0` before
running the wrapper:

```cmd
set W11UT_ALLOW_SETUP_UPGRADE=0
set W11UT_ALLOW_REBOOT=0
set W11UT_SKIP_VIRTUAL_MACHINES=0
set W11UT_ALLOW_DISK_CLEANUP=0
set W11UT_ALLOW_ADVANCED_DISK_CLEANUP=1
LOT-X\Run-Windows11UpgradeRepairWithPsExec-Once.cmd
```

For setup-based LOT runs without the GUI, copy
`Windows11UpgradeToolkit.config.template` to `Windows11UpgradeToolkit.config` at the
toolkit root, then set a target-reachable setup source:

```text
W11UT_SETUP_SOURCE=\\server\share\Windows11
W11UT_SETUP_DYNAMIC_UPDATE=Disable
```

You can still override the config for one CMD session by setting the environment variable
before launching a LOT wrapper:

```cmd
set W11UT_SETUP_SOURCE=\\server\share\Windows11
set W11UT_SETUP_DYNAMIC_UPDATE=Enable
LOT-X\Run-Windows11UpgradeRepairWithPsExec-Once.cmd
```

The launcher GUI reads the same config file to prefill its Options tab, then passes the
selected values to the LOT window it starts.

### Optional setup media SHA256 manifest

For stronger corruption detection, generate one manifest per Windows setup media folder
before publishing the media to a share. The manifest file must live at the root of the
same media folder as `setup.exe`:

```text
SmartM365-SetupMediaManifest.sha256.csv
```

Generate a manifest for one media root:

```powershell
.\Scripts\New-SmartM365SetupMediaManifest.ps1 -MediaRoot .\SetupSource\FR-fr -Force
```

Generate manifests for every direct media folder under `SetupSource`:

```powershell
.\Scripts\New-SmartM365SetupMediaManifest.ps1 -SetupSourceRoot .\SetupSource -Force
```

The root `New-SetupMediaManifest.cmd` wrapper runs the second command against the local
`SetupSource` folder when launched without arguments. If multiple language folders exist,
for example `FR-fr`, `PL-pl`, and `EN-us`, each folder gets its own manifest.

Targets do not create this manifest. They only read it. In `LocalCache` mode, the endpoint
uses the manifest hash to detect source/cache drift and validates the complete local cache
after Robocopy. In `Share` mode, the endpoint validates the manifest directly on the share
before launching Setup.

Each LOT can also have its own `Windows11UpgradeToolkit.config` file in the LOT folder.
This is useful when LOTs use different site-local setup sources. The wrapper loads values
in this order: environment variables already defined before launch, LOT config, then
toolkit root config. `Update-LotCmdWrappers.cmd` creates a missing LOT config from the
template but does not overwrite an existing one.

`W11UT_SETUP_SOURCE` can contain multiple UNC candidates separated by semicolons. In that
case, each target computer tests the candidates from its own SYSTEM context, validates the
setup media and language, reads a small sample from the media, and selects the nearest
valid source before running or copying setup media:

```text
W11UT_SETUP_SOURCE=\\srv-paris\share\Windows11-24H2;\\srv-lyon\share\Windows11-24H2;\\srv-marseille\share\Windows11-24H2
```

For large estates, prefer `W11UT_SETUP_SOURCE_MAP` instead of listing every share in
`W11UT_SETUP_SOURCE`. The map is a CSV reachable from the target computer context:

```text
W11UT_SETUP_SOURCE_MAP=\\srv-central\share\SetupSourceMap.csv
```

Supported map columns are `ScopeType`, `ScopeValue`, `SetupSourcePath`, and `Priority`.
Supported scope types avoid AD sites: `Subnet`/`CIDR`, `IPPrefix`, `ComputerName`,
`ComputerPrefix`, and `Default`.

```csv
ScopeType,ScopeValue,SetupSourcePath,Priority
Subnet,10.10.0.0/16,\\srv-paris\share\Windows11-24H2,10
IPPrefix,10.20.,\\srv-lyon\share\Windows11-24H2,20
ComputerPrefix,PC-MRS-,\\srv-marseille\share\Windows11-24H2,30
Default,*,\\srv-central\share\Windows11-24H2,100
```

Each target first filters the CSV using its own IPv4 addresses and computer name, then
benchmarks only the matching sources.

To reduce per-device media-copy bandwidth, the default `W11UT_SETUP_COPY_IPG_MS=20`
applies a moderate robocopy inter-packet gap. Larger values slow each PC copy more
aggressively; set it to `0` only when no per-device copy throttling is wanted. The default
`W11UT_SETUP_COPY_JITTER_SECONDS=300` spreads copy starts randomly over five minutes, and
`W11UT_SETUP_SOURCE_CANDIDATE_LIMIT` to cap the number of matched sources benchmarked by a
target. Optional source-level copy concurrency is available with
`W11UT_SETUP_SOURCE_CONCURRENCY_LIMIT` and a writable UNC lease folder in
`W11UT_SETUP_SOURCE_CONCURRENCY_GATE_ROOT`.

Before copying setup media to the local cache, the endpoint checks that the cache drive can
hold the media while still preserving the configured `-MinimumFreeDiskGB` on the system
drive. In default `LocalCache` mode, this effectively requires:

```text
Free space + existing cache size >= setup media size + MinimumFreeDiskGB
```

If you intentionally run a local/direct test with a local path, set
`W11UT_CONFIRM_LOCAL_SETUP_SOURCE=1` explicitly.

The PowerShell endpoint script remains guarded: it only performs actions when the
orchestrator passes the corresponding `-Allow*` switch.

## Setup Upgrade Mode

By default, setup media is copied/cached locally by each target before setup is executed:

```text
C:\ProgramData\SmartM365\Windows11UpgradeToolkit\SetupMedia\<SetupMediaId>-<Language>
```

The media validation checks:

- `setup.exe` exists and is not unexpectedly small;
- `setup.exe` has a valid Microsoft Authenticode signature;
- `sources\install.wim` or `sources\install.esd` exists;
- the install image has a valid WIM/ESD header and is readable by DISM `/Get-WimInfo`;
- `sources\lang.ini` contains the expected language when language matching is enabled;
- if `SmartM365-SetupMediaManifest.sha256.csv` is present in the media root, every listed media file matches its expected size and SHA256 hash;
- `SmartM365-SetupMedia.json` is present or refreshed after a valid cache/copy.
In LOT/PsExec mode, the operator workstation copies only the small endpoint script. The
target computer, running as SYSTEM through PsExec, validates `SetupSourcePath` and copies
the setup media into its own local cache with an incremental `robocopy` pass. This avoids
using the technician workstation as the large media copy engine and lets remote sites use
a site-local share.

Important: `SetupSourcePath` must be reachable from the target computer context. For LOT
runs, use a UNC path such as `\\site-server\share\Windows11\en-GB` or a parent folder
containing language subfolders. A local repo path on the operator workstation is only
usable for local/direct runs unless the target can also access that same path; the GUI
warns before allowing that case.

Supported execution modes:

- `LocalCache`: default. Use the target local cache and copy from `SetupSourcePath` on the target when the cache is not valid.
- `Share`: run directly from `SetupSourcePath`; the target SYSTEM context must be able to read the share.
- `Auto`: prefer local cache, then share, then target-side local cache copy.

`Use existing media only` / `-SkipSetupMediaPreCopy` means the target will not copy setup
media. In `LocalCache` mode, the cache must already be valid before the run starts.

Example LOT environment variables for an alternate setup source:

```cmd
set W11UT_ALLOW_SETUP_UPGRADE=1
set W11UT_SETUP_SOURCE=\\server\share\Windows11-24H2
set W11UT_SETUP_EXECUTION_MODE=LocalCache
set W11UT_SETUP_MEDIA_ID=Win11-24H2
set W11UT_SETUP_LANGUAGE=fr-FR
```

Then launch a LOT wrapper.

Setup language values:

- `MatchSystem`: default. The target system language, for example `fr-FR`, must exist in the setup media `sources\lang.ini`.
- Specific culture tag, for example `fr-FR`, `en-GB`, or `en-US`: validates the source before copy when possible and again on the target.
- `Any`: disables language matching and should only be used intentionally.

## Guardrails

The endpoint script is diagnostic-only unless action switches are provided. LOT wrappers
provide the guarded action switches by default; direct PowerShell calls do not.

- `-AllowPolicyRepair`: removes known legacy Windows Update policy blockers such as WSUS remnants, `UseWUServer`, `NoAutoUpdate`, and non-Windows 11 TargetReleaseVersion holds.
- `-AllowWUReset`: resets Windows Update cache and services.
- `-AllowForceUpgrade`: triggers assigned Windows Update scan/download/install through Windows Update APIs and `UsoClient`.
- `-AllowSetupUpgrade`: starts Windows setup upgrade only after setup media and readiness checks pass.
- `-DirectSetupUpgrade`: starts Windows setup directly after setup media validation/cache preparation; the script still blocks when the system drive is below `-MinimumFreeDiskGB`.
- `-AllowReboot`: permits a controlled reboot when a pending reboot blocks progress.
- `-AllowSetupCompletionRebootWhenNoUser`: after setup upgrade exits successfully or requests a reboot, schedules a reboot only when no interactive user is connected. If user detection fails, no automatic reboot is scheduled. LOT/config value: `W11UT_SETUP_REBOOT_WHEN_NO_USER=1`.
- `-SkipVirtualMachines`: skips detected virtual machines before repair, setup copy, or upgrade actions.
- `-AllowDiskCleanup`: when free disk is below `-MinimumFreeDiskGB`, removes only rebuildable caches and old SmartM365 logs before failing the device for disk space.
- `-AllowAdvancedDiskCleanup`: optional heavier cleanup. Removes old guarded Windows upgrade folders when no setup/update activity appears active, then runs `DISM /Online /Cleanup-Image /StartComponentCleanup` only after safe cleanup still leaves insufficient free disk.
- `-AuditOnly`: reports what would be done without running repair or upgrade actions.

The script blocks setup upgrade when the device is already Windows 11, is not Windows 10, lacks confirmed Intune enrollment, has actionable Windows 11 compatibility blockers, has insufficient disk space, or has a pending reboot that has not been handled.

### Setup failure 0x8007001F and duplicate profiles

When Windows Setup exits with `0x8007001F` during `Gather data`, check Panther logs for profile migration errors such as `Duplicate profile detected for S-...` or invalid user profile messages. In that case the setup media is usually valid; Windows Setup is blocked by duplicate or inconsistent profile registry entries under:

```text
HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList
```

Useful remote check:

```cmd
psexec \\COMPUTER -s powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList' | ForEach-Object { $p=Get-ItemProperty -LiteralPath $_.PSPath; [pscustomobject]@{ SID=$_.PSChildName; ProfileImagePath=$p.ProfileImagePath; State=$p.State; RefCount=$p.RefCount } } | Sort-Object ProfileImagePath,SID | Format-Table -AutoSize"
```

Safe disk cleanup covers temporary folders, Delivery Optimization cache, Windows Update
download cache only when no update/setup activity appears active, old SmartM365 setup media
except the current media cache, and old SmartM365 logs. It does not delete user profile
content, applications, or `WinSxS` manually. Advanced cleanup is intentionally opt-in
because it can take time and removes older upgrade rollback folders only when they are old
enough and no setup/update activity appears active. DISM cleanup uses the supported
Windows component cleanup command:

```cmd
DISM /Online /Cleanup-Image /StartComponentCleanup
```

The toolkit does not use `/ResetBase`.

## AD Precheck For Already-Upgraded Devices

Before each LOT run, the PsExec launcher can load or refresh `DevicesAD.csv` with `Scripts\SmartM365-Windows11Upgrade-Export-ADDevicesCsv.ps1`. If AD already reports a computer as Windows 11 through `OperatingSystem`, the launcher moves it from `Computers.txt` to `ComputersAlreadyW11.txt` before starting PsExec.

- Use `Export-ADDevicesCsv.cmd` from the toolkit root to generate a forest-wide `DevicesAD.csv` cache.
- Add `AdDomain.txt` in a LOT folder, or set `W11UT_AD_DOMAIN`, to restrict the automatic AD export to one domain.
- Automatic LOT refreshes are limited to the current `Computers.txt` and write `DevicesAD.csv` inside the LOT folder.
- The root `DevicesAD.csv` is read as a forest-wide cache only when it already exists and is recent; a stale or missing root cache is not overwritten by a LOT-scoped refresh.

## Reports And Logs

Remote runtime data is written under:

```text
C:\ProgramData\SmartM365\Windows11UpgradeToolkit
```

LOT-side output:

```text
LOT-*\PsExecLogs
LOT-*\Reports
LOT-*\CentralLogs\Success\<Computer>\Latest
LOT-*\CentralLogs\ADMIN_SHARE_UNREACHABLE\<Computer>\Latest
LOT-*\CentralLogs\Errors\<Computer>\Latest
```

The orchestrator copies remote `Logs`, `Output`, and `LastRun.json` back into the matching `CentralLogs` bucket. Successful endpoint runs go to `Success`, administrative-share failures go to `ADMIN_SHARE_UNREACHABLE`, and all other failures go to `Errors`. When central log history is enabled, `Latest` is replaced by a timestamped `CycleN_yyyyMMdd-HHmmss` folder.

## Multi-LOT Concurrency

The PsExec orchestrator runs computer workers in parallel. `-ThrottleLimit` controls parallelism inside one LOT window, while `W11UT_GLOBAL_CONCURRENCY_LIMIT` and `-GlobalConcurrencyLimit` default to `15` and share this recoverable worker gate across LOT windows:

```text
Local\SmartM365_Windows11UpgradeToolkit_ComputerWorkers
```

`W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES` / `-GlobalConcurrencyLeaseTimeoutMinutes`
controls when an abandoned shared worker lease is considered stale and cleaned. The default
`0` sizes this automatically from the PsExec timeout. Leases record the launcher PID, job id,
worker PowerShell PID, computer, and cycle; stale slots are repaired instead of bypassed.

The GUI staggers `Launch all` starts by 5 seconds to avoid a local startup spike.
The internal operator-side worker script is `Scripts\SmartM365-Windows11Upgrade-PsExecWorker.ps1`; target devices still receive only the autonomous endpoint script.
