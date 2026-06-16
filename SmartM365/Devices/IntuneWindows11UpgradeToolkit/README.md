# Smart Intune Windows 11 Upgrade Toolkit

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
force upgrade, setup upgrade, controlled reboot, and virtual machine skipping. They also
default `SetupSourcePath` to:

```text
SetupSource
```

Populate that folder with Windows 11 setup media before launching setup-based upgrade. See
`SetupSource\README.md`.
By default, setup language validation is enabled with `W11UT_SETUP_LANGUAGE=MatchSystem`.
The target compares its installed Windows language with `sources\lang.ini` from the setup
media before starting setup.
`SetupSourcePath` can point either to a single media root or to a parent folder containing
one media subfolder per language; in `MatchSystem` mode the toolkit selects the subfolder
whose `sources\lang.ini` matches the target language.

To disable an action for a direct LOT launch, set its environment variable to `0` before
running the wrapper:

```cmd
set W11UT_ALLOW_SETUP_UPGRADE=0
set W11UT_ALLOW_REBOOT=0
set W11UT_SKIP_VIRTUAL_MACHINES=0
LOT-X\Run-Windows11UpgradeRepairWithPsExec-Once.cmd
```

The PowerShell endpoint script remains guarded: it only performs actions when the
orchestrator passes the corresponding `-Allow*` switch.

## Setup Upgrade Mode

By default, setup media is copied/cached locally on each target before setup is executed:

```text
C:\ProgramData\SmartM365\IntuneWindows11UpgradeToolkit\SetupMedia\<SetupMediaId>
```

The media validation checks:

- `setup.exe` exists;
- `sources\install.wim` or `sources\install.esd` exists;
- `setup.exe` is not unexpectedly small.

Supported execution modes:

- `LocalCache`: default. Use the target local cache and pre-copy from the operator workstation when possible.
- `Share`: run directly from `SetupSourcePath`; the target SYSTEM context must be able to read the share.
- `Auto`: prefer local cache, then share, then local cache copy.

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
- `-AllowReboot`: permits a controlled reboot when a pending reboot blocks progress.
- `-SkipVirtualMachines`: skips detected virtual machines before repair, setup copy, or upgrade actions.
- `-AuditOnly`: reports what would be done without running repair or upgrade actions.

The script blocks setup upgrade when the device is already Windows 11, is not Windows 10, lacks confirmed Intune enrollment, has actionable Windows 11 compatibility blockers, has insufficient disk space, or has a pending reboot that has not been handled.

## Reports And Logs

Remote runtime data is written under:

```text
C:\ProgramData\SmartM365\IntuneWindows11UpgradeToolkit
```

LOT-side output:

```text
LOT-*\PsExecLogs
LOT-*\Reports
LOT-*\CentralLogs
```

The orchestrator copies remote `Logs`, `Output`, and `LastRun.json` back into `CentralLogs`.

## Multi-LOT Concurrency

The PsExec orchestrator runs computer workers in parallel. `-ThrottleLimit` controls parallelism inside one LOT window, while `W11UT_GLOBAL_CONCURRENCY_LIMIT` and `-GlobalConcurrencyLimit` default to `15` and share this recoverable worker gate across LOT windows:

```text
Local\SmartM365_IntuneWindows11UpgradeToolkit_ComputerWorkers
```

`W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES` / `-GlobalConcurrencyLeaseTimeoutMinutes`
controls when an abandoned shared worker lease is considered stale and cleaned. The default
`0` sizes this automatically from the PsExec timeout. Leases record the launcher PID, job id,
worker PowerShell PID, computer, and cycle; stale slots are repaired instead of bypassed.

The GUI staggers `Launch all` starts by 5 seconds to avoid a local startup spike.
The internal operator-side worker script is `Scripts\SmartM365-Windows11Upgrade-PsExecWorker.ps1`; target devices still receive only the autonomous endpoint script.
