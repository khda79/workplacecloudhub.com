# Critical fixes made in the previous toolkits

Policy: the previous toolkits receive no functional change, except a critical fix. A critical
fix is made in the previous product first, limited to the defect, covered by a regression test,
then ported (or verified) in Smart Endpoint Migration Toolkit. Each fix below states whether the
defect was **reproduced** or only **deduced from code**.

Each fix is committed on its own, separately from the new product (one commit per previous
toolkit). None of them is released or deployed yet.

## 1. Windows11UpgradeToolkit: DirectSetup forced reboots back on

| Item | Detail |
| --- | --- |
| Severity | Critical: unwanted reboot of a device in use |
| Evidence | **Reproduced** offline. With `W11UT_DIRECT_SETUP_UPGRADE=1` and `W11UT_ALLOW_REBOOT=0` (GUI "Allow reboot" unchecked, which the GUI passes as `0`), the LOT wrapper resolved `W11UT_ALLOW_REBOOT=1` and `W11UT_SETUP_REBOOT_WHEN_NO_USER=1`. The endpoint then received `-AllowReboot`, which enables pending-reboot handling and the uptime threshold reboot (7 days by default) that reboots even with a connected user. |
| Fix | `Scripts\Run-Windows11UpgradeRepairWithPsExec-Lot.cmd`: the two lines that forced both values when DirectSetup was on are replaced by a comment. Defaults are unchanged (both still default to `1` when not set). |
| Scope | 2 lines of one CMD file; README note. No GUI change, so no DemoAssets refresh. Package version unchanged (1.0.0 build 2 is published; a new build is needed to distribute the fix). |
| Test | New `Tests\Test-SmartM365Windows11DirectSetupRebootCmd.ps1`: 3 failures before the fix, 7/7 after. Existing `Test-SmartM365Windows11RunGuardAndCmd.ps1` and the Windows 11 offline demo test still pass. |
| Port | Smart Endpoint Migration Toolkit never derives reboot options from DirectSetup: `Windows11Upgrade.AllowReboot=false` is always honoured (Core test "Windows 11 DirectSetup never re-enables AllowReboot when it is off"). |

## 2. IntuneHybridJoinToolkit: Single PC launch ignored the GUI options

| Item | Detail |
| --- | --- |
| Severity | Critical: an unchecked authorization (`dsregcmd /leave`) still ran |
| Evidence | **Reproduced** offline. The GUI passes its options as `EHJIR_*` environment values. A LOT launch turns them into launcher parameters in `Run-IntuneHybridJoinRepairWithPsExec-Lot.cmd`; a Single PC launch called the PsExec launcher directly, which does not read `EHJIR_*`. The launcher default `-AllowDsregLeave = $true` applied even when the box was unchecked, and "Skip virtual machines", the cleanup and reboot authorizations, night pause, delays, retries and PsExec timeout chosen in the GUI were ignored. |
| Fix | `Scripts\SmartM365-IntuneHybridJoinRepair-LotLauncher-GUI.ps1`: new `Get-SingleComputerOptionArguments` maps `EHJIR_*` to launcher parameters with the same table and defaults as the LOT CMD (`-AllowDsregLeave:$false` when unchecked); `Start-ToolkitSingleComputer` appends them. Single PC still uses one worker. GUI version 1.20 -> 1.21 (header and `$GuiVersion`), README note. |
| Scope | One GUI function added, one line in `Start-ToolkitSingleComputer`. No visual change. |
| Test | New `Tests\Test-SmartM365IntuneHybridJoinSinglePcOptions.ps1`: 5 failures before the fix, 10/10 after. The 5 existing Hybrid Join tests and the Hybrid Join offline demo test still pass. |
| DemoAssets | Refreshed locally (`DemoAssets\2026-10-24`, GUI 1.21): native views and exports regenerated; they are byte-identical to 2026-10-23, so the report screenshot and website/presentation derivatives were carried forward and marked `carried_from` in the manifest. |
| Port | In Smart Endpoint Migration Toolkit a Single PC Run goes through the same engine, option resolution and GUI selection as a Wave (`-GuiSelectionPath` is passed for both; GUI test "Single PC command passes -ComputerName and no Wave"); only `ThrottleLimit` is forced to 1 by the `SinglePcOneWorker` rule. |

## Observed in the previous toolkits, not fixed there (not critical)

These were found while porting. They do not cause an unwanted action, so the previous toolkits
were left unchanged; Smart Endpoint Migration Toolkit handles them by design.

| Product | Observation | Evidence | In the new product |
| --- | --- | --- | --- |
| IntuneHybridJoinToolkit | A LOT launch with "Allow dsregcmd /leave" unchecked passes `-AllowDsregLeave:$false` to `-File`; Windows PowerShell 5.1 rejects that syntax, so on a technician computer without PowerShell 7 the Run fails to start (fail-closed). | Reproduced with a stub script (5.1 rejects, 7 accepts) | Options go through a JSON contract, not CMD switches |
| Windows11UpgradeToolkit | A forced stop can leave the local PsExec process running. | Deduced from code | Reservations record PsExec PID and start time; forced local stop stops the process tree |
| Windows11UpgradeToolkit | No per-LOT run mutex: the same LOT can be started twice. | Deduced from code | Exclusive Wave lock |
| Windows11UpgradeToolkit | Local setup source confirmation can be bypassed by the GUI path. | Deduced from code | `SEMT_SETUP_SOURCE_NOT_UNC` rule and explicit `ConfirmLocalSetupSource` |
| Both | Report `file://` links are not escaped (spaces, `#`). | Deduced from code | `ConvertTo-SemtFileUri` escapes them |
