# Intune Remediation Scripts

This folder groups SmartM365 Intune remediation, detection, diagnostic, and action scripts by functional area and scenario. It belongs to `IntuneRemediation`, together with the GUI manager used to publish and manage these packages in Microsoft Intune.

## Organization

- `WindowsUpdate/`: Windows Update scan, cache, service, policy, and proxy scenarios.
- `DeliveryOptimization/`: Delivery Optimization and Content Engine health scenarios.
- `WUfB/`: Windows Update for Business identity and PolicyState binding scenarios.
- `SetupDiag-Upgrade/`: SetupDiag, Windows 10/11 upgrade diagnostics, and upgrade readiness scenarios.
- `Intune-MDM/`: Intune, MDM, Hybrid Join, Intune Management Extension, and Data Boundary scenarios.
- `Disk-Cleanup-Storage/`: disk space, cleanup candidates, and upgrade storage readiness scenarios.
- `UWP-Store-AppRepository/`: AppRepository, Microsoft Store, and Windows Update health scenarios.
- `Standalone/`: operator actions, diagnostics, and one-off scripts that should not run as recurring Intune remediations.

## Naming Convention

Active scripts live in scenario folders. File names do not include version numbers; the current script version belongs in the script header and Git history keeps the change history.

When a scenario contains a standard Intune remediation pair, the folder usually contains:

- `SmartM365-*-Detection.ps1`
- `SmartM365-*-Remediation.ps1`

When a script is not designed to run repeatedly as a remediation, use an explicit action or diagnostic naming pattern such as:

```text
SmartM365-*-Action.ps1
SmartM365-*-Diagnostic.ps1
```

## Consolidated Scenarios

- `WindowsUpdate/WindowsUpdate-Reset/`: replaces older Autopatch `0x80244007` and forced Windows Update reset scenarios.
- `WindowsUpdate/Policy-Blockers/`: consolidates WSUS/GPO, WSUS remnants, NoAutoUpdate, Policy-Blocking-Access, and WUfB configuration blockers.
- `WindowsUpdate/Service-And-Scan-Health/`: replaces older service health, service refresh, and forced scan scenarios.
- `WindowsUpdate/Cache-Health/`: replaces download failure scenarios.
- `Intune-MDM/MDM-Enrollment-Repair/`: consolidates stale device join, enrollment state, Hybrid Join, and missing MDM task scenarios. The detection script targets stale or broken local state; if it runs through Intune, basic IME reachability is already proven.
- `SetupDiag-Upgrade/Upgrade-Staging-Health/`: consolidates missing upgrade files and upgrade residue scenarios.
- `Disk-Cleanup-Storage/Disk-Space-Cleanup/`: consolidates low free space, cleanup candidates, and forced disk cleanup scenarios.
- `DeliveryOptimization/ContentEngine-Health/`: replaces the older Delivery Optimization issue scenario.
- `UWP-Store-AppRepository/WU-Health/`: replaces the older broad endpoint repair scenario.
- `SetupDiag-Upgrade/Upgrade-Diagnostics/`: consolidates SetupDiag-required and upgrade-blocking issue scenarios.
- `Standalone/WindowsUpdateLog/`, `Standalone/Repair-DISM/`, `Standalone/Upgrade-Actions/`, `Standalone/Network-Diagnostics/`, `Standalone/WindowsPolicy/`, and `Standalone/Device-Recovery/`: keep one-off actions and manual diagnostics outside recurring remediations.

## Additional Endpoint Packages

Each scenario below contains a standalone detection/remediation pair. Run as
SYSTEM in **64-bit Windows PowerShell 5.1**. No SmartM365 module, tenant profile,
download or companion configuration is required. The Device Reboot Manager
variant intentionally integrates with a separately installed application.

| Folder / package | Default detection | Default remediation |
| --- | --- | --- |
| `Security-LAPS/Windows-LAPS-Health` | Windows LAPS backup policy enabled, latest local processing failed or last success older than 24 hours | Invoke LAPS policy processing; wait up to 30 seconds for a new success/failure event |
| `System-Time/Windows-Time-Sync` | Time service stopped or three measurements against the current source exceed 120 seconds | Start the enabled service if needed, request resynchronization and measure again |
| `Security-Defender/Defender-Signature-Freshness` | Defender in Normal mode with signatures older than 48 hours | Update from configured sources, then verify freshness |
| `Certificates/Network-Certificate-Expiry` | Machine personal-store client-authentication certificates with a private key expire within 14 days | Notify interactive sessions; no certificate removal |
| `BITS/BITS-Stale-Jobs` | Current execution account has Error/TransientError jobs unchanged for 24 hours | Retry those jobs asynchronously and inspect their new state |
| `Printing/Print-Spooler-Health` | Enabled spooler stopped or Error/Blocked print jobs older than two hours | Restart/start the service, then verify; retain jobs by default |
| `Reboot/Pending-Reboot-Age` | CBS or Windows Update pending-reboot signals observed for at least 72 hours | Display a 120-second session message; never restart automatically |
| `Reboot/Pending-Reboot-Age-DeviceRebootManager` | Same observation rule as the standard variant | Start the installed, signed Device Reboot Manager limited interactive task |

`Deferred` means applicability or action could not be established, **not** healthy.
Detection returns `1` only for actionable findings and `2` for technical errors;
`0` also covers not-applicable, observation and deferred states. Read the status
and reason in Intune output rather than equating every zero with verified health.
Remediation returns `1` for unresolved/technical failures. A requested asynchronous
action, notification or task start is distinguished from a verified repair.

### Operational boundaries

- LAPS reads policy and event IDs only; it never retrieves passwords or requests
  forced rotation. Event 10004 confirms local policy processing, not independent
  verification of password escrow in AD/Entra. CSP policy takes precedence over GPO.
- Time synchronization preserves configured peers and startup settings. Disabled
  service or unmeasurable/unreachable source is deferred. Native output localization
  is not used to choose a replacement NTP server.
- Defender passive modes and intentionally disabled protection are not overridden.
  A failed update or stale post-check is not reported as success.
- Certificate inspection covers **machine client-authentication certificates**;
  it does not prove which one a particular Wi-Fi/VPN profile uses. No certificate
  subjects, private keys or thumbprints are exported. Optional
  `-RequestAutoEnrollment` runs `certutil -pulse` only for domain members with
  auto-enrollment enabled. It requests policy processing; it does not claim renewal
  or implement Intune SCEP/PKCS renewal. Default behavior is notification.
- BITS examines only the current account (SYSTEM when deployed as recommended),
  without enumerating other users. Suspended, active and recent jobs are excluded.
  Cancellation is available only with an explicit `-CancelJobIds` GUID allowlist;
  ordinary scheduled remediation retries without deleting jobs.
- Print purge requires manual `-AllowPurge`; it rechecks each matching job ID and
  printer before removal. It never deletes the spool directory or all queue jobs.
  Disabled spoolers stay disabled. Printer/document/user names are not logged.

### Pending-reboot observation and Device Reboot Manager

Deploy **one** reboot variant per device. Both variants recheck the pending signals
before acting, record `FirstObservedUtc` alongside the boot time and reset the
observation when the signals clear, a different boot is observed, or no previous
reason remains. Age is a lower bound from sampled observations, not the original
Windows request time and not device uptime. The first run starts observation even
on a device that has been up for weeks. File-rename queues alone are excluded.

Observations live under
`%ProgramData%\SmartM365\IntuneRemediation\Output\<package>\observation.json`.
Corrupt or future-dated state produces an error, never an immediate reminder.
The default reminder cooldown is 24 hours; reminders do not clear the finding.
No interactive recipient means notification is deferred.

The Smart variant requires the standard deployment's task
`\SmartM365\Device Reboot Manager`, a valid signature on the installed GUI and
the expected PowerShell action. It rejects elevated or modified task actions.
It does not install the app, create a task, change the app configuration, or launch
a GUI in SYSTEM session 0. The app keeps its own uptime thresholds, cooldown,
postpone and restart policy: a task start does **not** guarantee a visible window
if those conditions are not met. Match the app policy to the intended reminder
cadence before rollout. The standard variant remains a notification-only option.

Both reminder variants emit `RebootVerified=False`; only a subsequent observation
can confirm that pending reboot signals cleared.

### Offline validation and sources

Run `SmartM365/Tests/Test-SmartM365IntuneAdditionalPackagesOffline.ps1` under
Windows PowerShell 5.1 and PowerShell 7. Tests execute the standalone sources with
synthetic OS interfaces and block unmocked operational commands. Their evidence
does not qualify real tenant, endpoint, printer or interactive-session behavior.

API references: [Windows LAPS processing](https://learn.microsoft.com/en-us/powershell/module/laps/invoke-lapspolicyprocessing),
[Windows Time tools](https://learn.microsoft.com/en-us/windows-server/networking/windows-time-service/windows-time-service-tools-and-settings),
[Defender signature updates](https://learn.microsoft.com/en-us/powershell/module/defender/update-mpsignature),
[BITS asynchronous retry](https://learn.microsoft.com/en-us/powershell/module/bitstransfer/resume-bitstransfer),
[targeted print-job removal](https://learn.microsoft.com/en-us/powershell/module/printmanagement/remove-printjob),
[session messages](https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/msg).

## Validation Before Publishing

`WindowsHello/Reset-WindowsHello` reports unavailable AD connectivity as `Deferred`
and does not trigger a key reset. Connectivity alone does not establish key corruption.
The remediation script requires `-AllowWindowsHelloReset` for an operator-approved
manual reset after diagnosis; its default invocation changes no keys or services.

`IntuneManagementExtension-Health` reads the native IME logs under
`%ProgramData%\Microsoft\IntuneManagementExtension\Logs`. SmartM365-generated logs
remain under `%ProgramData%\SmartM365\IntuneRemediation\Logs`.

The Time-Zone pair uses `SmartM365-Time-Zone-Detection.ps1` and
`SmartM365-Time-Zone-Remediation.ps1` so the manager associates them unambiguously.

Offline regressions (PowerShell 5.1 or 7, from the repository root):

```powershell
pwsh -NoProfile -File SmartM365/Tests/Test-SmartM365IntuneRemediationReviewOffline.ps1
```

Before publishing a package to Intune:

- review detection and remediation scripts together;
- run PSScriptAnalyzer when possible;
- confirm that remediation logic is idempotent;
- avoid storing tenant-specific values, group IDs, hostnames, or customer data in scripts;
- use the GUI manager to publish with delegated Graph authentication.

Generated exports, downloaded cloud scripts, reports, logs, and local archives are runtime artifacts and should remain outside Git.
