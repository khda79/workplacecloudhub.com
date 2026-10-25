# Campaigns, Waves, Runs and Cycles

## Campaign

A Campaign is the kind of migration. Each Campaign has its own endpoint script, option
defaults, rules, Waves, Runs, run guard history, GUI choices and report columns.

| Campaign | Done when | Typical next actions |
| --- | --- | --- |
| `HybridJoin` | The device is Hybrid Entra joined and enrolled in Intune (also detected after the cycle from a scoped Intune/Entra inventory) | `REBOOT_DEVICE`, `WAIT_USER_LOGOFF_OR_REBOOT_DEVICE`, `FIX_INTUNE_ENROLLMENT_FIRST` |
| `Windows11Upgrade` | The device runs Windows 11, or is reported not capable | `REBOOT_AND_RETRY_SETUP`, `CHECK_SETUP_LOGS`, `VERIFY_LEGACY_MEDIA_LOCATION` |

Statuses shared by both Campaigns: `TARGET_OPERATION_ACTIVE` (`WAIT_ACTIVE_TARGET_OPERATION`),
`TARGET_PENDING_OPERATION` (`WAIT_OTHER_CAMPAIGN_PENDING_OPERATION`), the `LEGACY_*` statuses of
[LEGACY-COEXISTENCE.md](LEGACY-COEXISTENCE.md), `RUN_GUARD_ACTIVE`, and the technician-side
`DRYRUN_*`, `PSEXEC_*`, `REMOTE_*`, `COMPONENT_INTEGRITY_MISMATCH` and `CANCELLED_BY_OPERATOR*`.

Completed computers leave `Computers.txt` and are archived next to it
(`Computers.AlreadyEnrolled.txt`, `Computers.AlreadyWindows11.txt`,
`Computers.HardwareNotCapable.txt`).

## Wave

```text
Waves\<Campaign>\<Wave>\
  Computers.txt                       one computer per line (# comments allowed)
  Wave.local.json                     Wave options (created from the Campaign template)
  Start-Wave-Loop.cmd                 Loop until complete or stopped
  Start-Wave-Once.cmd                 one cycle
  Start-Wave-Loop-IgnoreRunGuard.cmd  Loop, run guards ignored on the first cycle
  Start-Wave-Once-IgnoreRunGuard.cmd  one cycle, run guards ignored
```

Wave names use letters, digits, dot, dash and underscore (80 characters at most). A Wave is
created by **New Wave** (pasted list), by **Automatic Wave** (inventory filters) or by hand.
Creating a Wave never launches it. An existing Wave is never overwritten.

The launchers are thin: they call `Scripts\Run-EndpointMigrationWave.cmd "<Wave>" Loop|Once
Default|IgnoreRunGuard`, which derives the Campaign from the parent folder and runs Windows
PowerShell 5.1 with the engine. `Update-WaveCmdWrappers.cmd` regenerates them for every Wave.

### Automatic Wave

1. Refresh the shared inventories if needed (AD; Intune and Entra with delegated sign-in).
2. **Preview** runs the Automatic Wave engine without `-Create`: it writes evidence under
   `Runs\<Campaign>\_AutomaticWave\` (selected, excluded with reason code, summary) and creates
   nothing.
3. **Create** is enabled only while the filters are those of the last preview. It runs the same
   selection with `-Create` and creates the Wave (`Source=Automatic`, the filters in its
   description and the evidence path). It never launches the Wave.

Default names: `WAVE-AUTO-HJ-<timestamp>`, `WAVE-AUTO-W10-<timestamp>`.

## Run

```text
Runs\<Campaign>\<Wave>\<RunId>\
  RunManifest.json           Campaign, Wave, engine/endpoint versions and hashes, technician, integrity
  EffectiveOptions.json      every option with its value and source
  Logs\                      SEMT-<Campaign>-Launcher_<RunId>.log, worker logs
  PsExecLogs\                one PsExec transcript per attempt
  State\                     ActiveRun_<pid>.json, StopRequested_<pid>.json, Contracts\
  Inventory\                 scoped inventory exports of this Run
  CentralLogs\               endpoint evidence collected per status bucket
  Reports\                   live and final HTML report, CSV per cycle and summary
Runs\<Campaign>\<Wave>\RunHistory.json
```

A Single PC Run lives in `Runs\<Campaign>\_SinglePC\<RunId>\` with its one-computer Wave in
`Wave\`. It always uses one worker (`SinglePcOneWorker` rule).

GUI choices are captured for each Run in `Runs\<Campaign>\_GuiSelections\<timestamp>.json`
(empty text fields are not captured, so the Wave or global value still applies). The generated
launcher CMD is kept next to it as evidence.

Older Runs of a Wave can be compressed (`RunArchiveAfterDays`): a Run folder is removed only
after its archive was re-opened and matched entry count and size.

## Cycle

Each cycle:

1. reloads `Computers.txt` and removes duplicates;
2. refreshes inventories when the Run needs them and the shared cache is not usable;
3. runs the Campaign pre-cycle hook (Hybrid Join: devices already in Intune are skipped;
   Windows 11: devices already on Windows 11 in AD/Intune are skipped);
4. resolves DNS for every computer (a DNS failure blocks the cycle when
   `DnsPreflightBlocksCycle` is on);
5. applies the run guard (back-off per computer and failure category);
6. starts one worker per computer under the per-Wave and shared caps;
7. classifies results, archives completed computers, runs the post-cycle hook, writes CSV/HTML;
8. waits `DelayBetweenCyclesMinutes` (adaptive for Hybrid Join) or the night pause.

`RunMode=Once` limits the Run to one cycle. `MaxCycles=0` with Loop means unlimited (the GUI
asks for an explicit confirmation). A dry run is limited to one cycle.

## Launch All Waves

Launches every launchable Wave of the **selected Campaign** that is not already running, one
every 5 seconds, with the same run mode and GUI choices, after one confirmation. It never
launches Waves of the other Campaign.
