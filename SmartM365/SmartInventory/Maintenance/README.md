# SmartInventory maintenance tools

## Orchestrator server acceptance

`SmartM365-OrchestratorServerAcceptance.ps1` (1.0.7, PowerShell 7.2+) qualifies a server before it
is switched from Inactive to Active in the orchestrator GUI. The orchestrator capability probe only
signs in or loads a module; this script makes a real, limited read for each capability, checks the
rows, then drops them.

What it guarantees:

- No CSV and no data file: rows stay in memory and only counts and check results are logged.
- Nothing written to the tenant: the SharePoint log upload and the Teams notifications of the log
  are disabled; a mail is sent only with `-SendTestMail`.
- The only writes are the run log under `LogAllRootPath` (folder `SmartM365-OrchestratorServerAcceptance`),
  a witness folder under `WorkspaceRootPath` (removed at the end) and a work folder under `%TEMP%`
  (removed at the end).

### Checks

| Capability | Real test | Data checks |
|---|---|---|
| RunAccount | Running identity; `-ExpectedAccount` fails another account, otherwise compared with the account of the orchestrator scheduled task | |
| Modules | PowerShell 7 and 5.1; Python 3.10+ from `PreparedPythonCommand` of the WorkplaceEvidence-Prepare configuration (WARN when missing: no orchestrator capability covers it); the modules the jobs import: Microsoft.Graph Authentication, Users, Groups, Identity.DirectoryManagement, Reports, DeviceManagement, Beta.Groups, Beta.DeviceManagement, ExchangeOnlineManagement, ImportExcel (required), MicrosoftTeams >= 4.7.1, ActiveDirectory, Adaxes (optional); every Microsoft.Graph module at the Authentication version; Exchange and SharePoint snap-ins | missing required module = FAIL |
| CodeSigning | Authenticode status of the script, Core, distributed module and orchestrator; execution policy | |
| Network | HTTPS to login, Graph, Exchange Online and the SharePoint host; Teams endpoints (teams.microsoft.com, api.interfaces.records.teams.microsoft.com); proxy used; latency | any HTTP answer = reachable; unreachable Teams endpoint = WARN |
| Clock | Offset against the Date header of the Entra sign-in endpoint | WARN 60 s, FAIL 300 s |
| Certificate | Same readiness test as the orchestrator: store, expiry, real signature with the private key by the running account | WARN under 30 days |
| SharedRuntime | DATA-ALL, DATA-LAST, LOG-ALL read/write; witness create/write/read/metadata/delete under WorkspaceRootPath; read of the code share; free space | median SMB call WARN 15 ms, FAIL 50 ms; DATA free FAIL < 10 %, WARN < 15 %; local disk FAIL < 2 GiB, WARN < 10 GiB |
| AD | Forest domains, `Get-ADUser -ResultSetSize 100` per domain (PowerShell 7 like AD-Inventory), ActiveDirectory in Windows PowerShell 5.1 | ObjectGUID, SamAccountName, UPN, WhenCreated |
| Graph | 100 users, 100 Entra devices, 100 Intune managed devices | Id, UPN, DeviceId, DeviceName, dates |
| GraphPermissions | Roles of the token against `RequiredGraphAppRoles` of the enabled catalog jobs, plus Mail.Send (Graph mail) and Sites.Selected (SharePoint upload) | missing roles = WARN with the jobs concerned |
| EXO | `Get-EXOMailbox -ResultSize 100` | Guid, PrimarySmtpAddress, ExternalDirectoryObjectId |
| ExchangeOnPrem | Snap-in in Windows PowerShell 5.1, `Get-Mailbox -ResultSize 100` | Guid, ExchangeGuid, PrimarySmtpAddress, Database |
| SharePointOnPrem | Snap-in, farm, web applications, `Get-SPSite -Limit 10` | Id, Url |
| TeamsPowerShell | Same calls as the Teams job: `Connect-MicrosoftTeams` (same TenantId), `Get-CsPhoneNumberAssignment -Top 10` (`Get-CsTenant` is not available app-only with recent modules) | TelephoneNumber, NumberType; no number = WARN |
| Adaxes | Services, domains and LDAP filter of the Adaxes collector configuration; 100 users per domain, one child process per domain limited by `-AdaxesDomainTimeoutSeconds` (default 120; the next domains are not read after a time-out; 0 or 1 user in every domain points to missing Adaxes permissions of the account; like the collector, a user whose objectClass the account cannot read is excluded by DN and the domain read again, 25 at most, reported as WARN with a count only) | ObjectGUID, SamAccountName, UPN |
| Mail | From and recipient; Graph: Mail.Send granted; SMTP: port 25 answers 220 | test mail only with `-SendTestMail` |

Each data set must return 1 to N rows, with key columns filled, values in the expected format
(GUID, SMTP address, URL, plausible dates) and no duplicate key. A domain that fails while others
succeed is WARN; every domain failing is FAIL.

The run ends with a summary in several tables: 1. server and run (server, OS, account, tenant,
PowerShell versions, code and data roots, duration); 2. capabilities (status, seconds, one-line result);
3. real extractions (capability, data set, domain or scope, rows, seconds, checks); 4. measurements
(modules and versions, code signing, network endpoints, clock and certificate, shares, latency and
disks); 5. catalog jobs this server could not run and what they miss; then the totals and a verdict
(READY, READY WITH WARNINGS or NOT READY with the capabilities to fix).

A capability whose component is not installed (module or snap-in) is SKIPPED, or FAIL when it was
named with `-Capability`. After a time-out, the detail names the last step the child process
announced (for example the Adaxes service and domain being read). The run ends with the catalog jobs this server could run (same rule as the
election) and a PASS/WARN/FAIL/SKIPPED table with durations. Exit code 1 when a capability is FAIL.

### Running it

On the server to qualify, from the code share, first with your account then under the orchestrator
service account (the certificate private key and the shares must work for that account):

```powershell
pwsh -NoProfile -File <code share>\SmartM365\SmartInventory\Maintenance\SmartM365-OrchestratorServerAcceptance.ps1 -Tenant prod
pwsh -NoProfile -File <code share>\SmartM365\SmartInventory\Maintenance\SmartM365-OrchestratorServerAcceptance.ps1 -Tenant prod -ExpectedAccount 'DOMAIN\service-account'
```

Other entry points:

- `SmartInventory.cmd Orchestrator-ServerAcceptance` (advanced manual entry of the catalog).
- Orchestrator GUI, Operations > Manual launches > `Orchestrator-ServerAcceptance`.
- Orchestrator GUI, Configuration > Servers and versions > Selected server: **Test this server**
  and **Test as service account** (runas with the account of the orchestrator scheduled task).
  The test reads from the computer it runs on, so these buttons are enabled only when the selected
  server is the computer running the GUI.

Useful parameters: `-Capability AD, ExchangeOnPrem` (subset), `-SampleSize` (default 100),
`-SmallSampleSize` (Teams and SharePoint, default 10), `-CapabilityTimeoutSeconds` (default 600),
`-AdaxesDomainTimeoutSeconds` (default 120),
`-SmbIterations`, `-SmbWarningMilliseconds`, `-SmbFailureMilliseconds`, `-SendTestMail`, `-TestMailTo`.

Offline test: `SmartM365/Tests/Test-SmartM365OrchestratorServerAcceptanceOffline.ps1`.
