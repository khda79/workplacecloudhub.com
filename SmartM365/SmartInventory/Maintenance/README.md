# SmartInventory maintenance tools

## Orchestrator server acceptance

`SmartM365-OrchestratorServerAcceptance.ps1` (1.0.2, PowerShell 7.2+) qualifies a server before it
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
| RunAccount | Running identity; `-ExpectedAccount` fails another account | |
| Modules | PowerShell 7 and 5.1; Microsoft.Graph.Authentication >= 2.37.0, ExchangeOnlineManagement >= 3.7.2, MicrosoftTeams >= 4.7.1, ImportExcel, ActiveDirectory, Adaxes; every Microsoft.Graph module at the Authentication version; Exchange and SharePoint snap-ins | |
| CodeSigning | Authenticode status of the script, Core, distributed module and orchestrator; execution policy | |
| Network | HTTPS to login, Graph, Exchange Online and the SharePoint host; proxy used; latency | any HTTP answer = reachable |
| Clock | Offset against the Date header of the Entra sign-in endpoint | WARN 60 s, FAIL 300 s |
| Certificate | Same readiness test as the orchestrator: store, expiry, real signature with the private key by the running account | WARN under 30 days |
| SharedRuntime | DATA-ALL, DATA-LAST, LOG-ALL read/write; witness create/write/read/metadata/delete under WorkspaceRootPath; read of the code share; free space | median SMB call WARN 15 ms, FAIL 50 ms; DATA free FAIL < 10 %, WARN < 15 %; local disk FAIL < 2 GiB, WARN < 10 GiB |
| AD | Forest domains, `Get-ADUser -ResultSetSize 100` per domain (PowerShell 7 like AD-Inventory), ActiveDirectory in Windows PowerShell 5.1 | ObjectGUID, SamAccountName, UPN, WhenCreated |
| Graph | 100 users, 100 Entra devices, 100 Intune managed devices | Id, UPN, DeviceId, DeviceName, dates |
| GraphPermissions | Roles of the token against `RequiredGraphAppRoles` of the enabled catalog jobs, plus Mail.Send (Graph mail) and Sites.Selected (SharePoint upload) | missing roles = WARN with the jobs concerned |
| EXO | `Get-EXOMailbox -ResultSize 100` | Guid, PrimarySmtpAddress, ExternalDirectoryObjectId |
| ExchangeOnPrem | Snap-in in Windows PowerShell 5.1, `Get-Mailbox -ResultSize 100` | Guid, ExchangeGuid, PrimarySmtpAddress, Database |
| SharePointOnPrem | Snap-in, farm, web applications, `Get-SPSite -Limit 10` | Id, Url |
| TeamsPowerShell | `Get-CsTenant` (same TenantId), `Get-CsOnlineUser -ResultSize 10` | Identity, UPN |
| Adaxes | Services, domains and LDAP filter of the Adaxes collector configuration; 100 users per domain, one child process per domain limited by `-AdaxesDomainTimeoutSeconds` (default 120; the next domains are not read after a time-out) | ObjectGUID, SamAccountName, UPN |
| Mail | From and recipient; Graph: Mail.Send granted; SMTP: port 25 answers 220 | test mail only with `-SendTestMail` |

Each data set must return 1 to N rows, with key columns filled, values in the expected format
(GUID, SMTP address, URL, plausible dates) and no duplicate key. A domain that fails while others
succeed is WARN; every domain failing is FAIL.

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
