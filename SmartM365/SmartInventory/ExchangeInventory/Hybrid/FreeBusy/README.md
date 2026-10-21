# Exchange Hybrid Free/Busy

Weekly check of the Exchange hybrid free/busy (availability) configuration with Microsoft's
[CSS-Exchange Free/Busy Checker](https://microsoft.github.io/CSS-Exchange/Diagnostics/FreeBusyChecker/FreeBusyChecker/),
on the Exchange on-premises side and on the Exchange Online side, for DAuth and OAuth.

## Scripts

- `SmartM365-Exchange-Hybrid-FreeBusy-Inventory.ps1`: collector, Windows PowerShell 5.1, elevated,
  on an Exchange server. Started by the orchestrator job `Exchange2016-FreeBusy-Check` (Sunday 05:00)
  or by `Launchers/OnPremises/Start-SmartM365-Exchange-Hybrid-FreeBusy-Inventory.cmd`.
- `SmartM365-Exchange-Hybrid-FreeBusy-Runner.ps1`: started by the collector in its own process.
- `ThirdParty/CSS-Exchange/FreeBusyChecker`: the vendored Microsoft script (see its `PROVENANCE.md`).

## How it runs

1. The collector prepares a work folder that only administrators can change (default
   `%ProgramFiles%\SmartM365\FreeBusy\Work`, `FreeBusyWorkFolderPath` to change it; shares are refused).
   Its owner is Administrators, its access is limited to SYSTEM and Administrators, and every parent
   is checked; otherwise the run stops with `UnsafeWorkFolder`.
2. It copies Free/Busy Checker and the runner there, checks the copy (expected SHA256 and Microsoft
   signature) and the runner (same bytes as the repository runner, signed by the collector's
   certificate when the collector is signed), and checks them again just before the start.
3. The runner, in its own process with a time limit (`FreeBusyCheckerTimeoutMinutes`, default 30):
   - checks the ActiveDirectory module (Windows feature; when missing the status is
     `ActiveDirectoryModuleMissing` with the command an administrator runs:
     `Install-WindowsFeature RSAT-AD-PowerShell`);
   - installs ExchangeOnlineManagement 3.x for the current user from PSGallery when it is missing
     (NuGet provider and TLS 1.2 handled, no prompt; `ModuleInstallFailed` otherwise);
   - opens the Exchange Online session with prefix `EO` using the tenant app and its certificate
     (`AppId`, `Thumbprint` or `Thumb`, `OrgDomain` from the tenant profile) and checks with
     `Get-ConnectionInformation` that it is that app-only certificate session. It never signs in
     interactively: without the settings, or if the session is not the right one, Exchange Online is
     `NotConnected` and the on-premises part still runs;
   - records the Exchange Online cmdlets used by Free/Busy Checker that are not exposed to the app;
   - runs Free/Busy Checker with `-SkipVersionCheck` alone (any other parameter would let it update
     itself).
4. The collector reads the HTML report as text only (nothing is deserialized), masks values that can
   reveal names, addresses, URLs, certificates or identifiers, and publishes two CSVs to DATA-ALL and
   DATA-LAST with a source receipt. The report and transcript stay in the work folder (last run only)
   and are never published.

## Outputs

- `Exchange_Hybrid_FreeBusy_Summary.csv`: four rows per run (Exchange on-premises and Exchange Online,
  each for DAuth and OAuth) with `RunStatus`, `SideStatus` (`Completed`, `NotConnected`,
  `NotAuthorizedPartial`, `NotRun`), the result counts, the cmdlets not available to the app and the
  roles that contain them. Counts are empty when a side was not measured.
- `Exchange_Hybrid_FreeBusy_Findings.csv`: one row per result line with side, authentication,
  check, item, status (`OK`, `Warning`, `Error`, `Info`, `NotAuthorized`) and the masked value.

`RunStatus` is `Completed`, `EXONotConnected`, `EXONotAuthorized` (some Exchange Online checks not
available to the app; for example `Test-EOOAuthConnectivity` needs a role such as Mail Recipients),
`ModuleInstallFailed`, `ActiveDirectoryModuleMissing`, `NotElevated`, `IntegrityFailed`,
`UnsafeWorkFolder`, `ConfigurationInvalid`, `Timeout`, `LaunchFailed`, `RunnerFailed`, `NoOutput` or
`ParseFailed`. Anything other than `Completed` is a warning; the CSVs are always published. A test
that the Exchange Online service itself fails (for example `Test-EOOrganizationRelationship`
returning an internal error) is an `Error` finding, not a missing permission.

Offline validation (synthetic data only, Free/Busy Checker is never run):

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File SmartM365/Tests/Test-SmartM365ExchangeFreeBusyOffline.ps1
```
