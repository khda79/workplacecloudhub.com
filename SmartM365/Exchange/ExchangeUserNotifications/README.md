# Exchange User Notifications

User-facing Exchange communication campaigns.

## Scripts

- `ExchangeMigration/SmartM365-ExchangeMigration-NotifyUsers.ps1`: sends migration notifications from a recipients CSV or folder.
- `ExchangeArchive/SmartM365-ExchangeArchive-NotifyUsers.ps1`: sends archive mailbox activation notifications from a recipients CSV or folder.
- `ExchangeMigrationMailboxSizeReduction/SmartM365-ExchangeMigrationMailboxSizeReduction-NotifyUsers.ps1`: sends pre-migration mailbox size reduction notifications. The license CSV selects users matching `TargetSkuPartNumbers`; when it is missing or older than `LicenseCsvMaxAgeHours`, `EnableLiveLicenseLookupWhenCsvMissing` lets the script resolve licenses live from Microsoft Graph. `MailboxQuotaBySkuPartNumber` sets the quota per license, for example E3 at 100 GB and F1/F3 at 2 GB. Exchange management uses Exchange Online first by default, with Exchange 2016 snap-in fallback when enabled, to verify mailbox existence and read live mailbox size before applying the per-user quota.

## Local Module

`SmartM365.Communications.psm1` contains campaign-local helpers:

- JSON campaign configuration loading.
- CSV import with BOM removal and delimiter auto-detection.
- Template token expansion and unresolved-token checks.
- Sent registry load/save.
- Explicit `MailSendMode` support: `Graph`, `SmtpRelay`, `Auto`, or `Disabled`, with optional BCC.
- Explicit `ExchangeManagementMode` support: `Auto`, `ExchangeOnline`, `Exchange2016`, or `Disabled`.
- Summary HTML generation.

The module is intentionally local to this area and should not be promoted to `SmartM365.Core` unless other SmartM365 product areas need it.

## GUI

`SmartM365-ExchangeUserNotifications-GUI.ps1` is a common WPF launcher for the three campaigns. It validates script/config files, builds the command line, defaults to dry-run mode, and displays campaign output while keeping all business logic in the existing scripts.

Use `Start-SmartM365-ExchangeUserNotifications-GUI.cmd` to launch it with PowerShell 7.

The GUI opens maximized, retaining the Windows title bar and restore controls.

### Migration recipients from Exchange Online (GUI 1.1.0)

In **Exchange Migration**, choose **Migration batches EXO**, or open the
**EXO migration batches** tab. **Connect to EXO** starts delegated interactive
authentication in a dedicated PowerShell 7 worker. Nothing connects on GUI
startup. The worker verifies the selected profile's `TenantId` GUID and connected
account before reading migration data; an ambiguous or mismatched session fails.
It requires ExchangeOnlineManagement 3.7.2 or newer and EXO RBAC access to
`Get-MigrationBatch`, `Get-MigrationUser` and `Get-MigrationUserStatistics`.
No additional write role or migration modification is required. Mail/Teams
authentication remains separate from this read-only EXO session.

Select one or more active mailbox batches and click **Load recipients**. Completed
batches are hidden by default and can be displayed for inspection. Completed
users, outbound batches and non-mailbox migration types are not eligible for this
campaign. The table displays membership, migration status, per-user notification
date, schedule provenance and prior notification exclusions. Duplicate addresses
are merged while retaining their batch identities; conflicting dates require an
operator decision. Failed/stopped users are not selected automatically.

`StartAfter` is synchronization start, not a mailbox cutover date. `CompleteAfter`
is the earliest scheduled completion, not guaranteed completion. A verified
individual schedule takes precedence over the batch. UTC values are converted
using the display time zone selected in the GUI (default: Romance Standard Time).
An unknown zone, missing schedule or unverified individual override requires an
explicit confirmation/date. **Apply date to selected recipients** changes only
the notification date; it never changes the migration schedule. A global
**Force effective date** is not passed for this source, preserving different dates
across batches. Already notified recipients require the existing **Force send**
option; completed migrations remain excluded.

**Review selection / prepare CSV** writes only selected recipients under the
campaign's tenant-specific `RunOutputPath/Selections`. The CSV records tenant,
batch identities, status, UTC/local schedule, date source and timestamps. Review
the selection, return to **Exchange Migration**, then use **Dry Run** before
**Live**. Each run regenerates the selected CSV and checks registry eligibility.
Changing tenant, time zone or batch selection invalidates recipients. Read errors
discard the previous audience; a snapshot older than 15 minutes must be reloaded.

The campaign receives `-MigrationBatchSelection` with this reviewed CSV. It
validates every row's date, source, tenant/profile and freshness before processing,
then uses the reviewed membership instead of opening another EXO mailbox-check
session. Standard CSV/folder campaigns keep their existing behavior. Selection
files contain mailbox addresses and must remain private; no file is committed.

Offline validation (no EXO connection or message transport):

```powershell
pwsh -NoProfile -STA -File SmartM365/Tests/Test-SmartM365ExchangeMigrationSelectionOffline.ps1
pwsh -NoProfile -STA -File SmartM365/Tests/Test-SmartM365GuiDemoApplicationsOffline.ps1 -Application Notifications
```

The fictional demo includes EXO batches, multiple dates, unverified schedules,
date conflicts and completed/already notified recipients. Live tenant RBAC,
authentication and returned EXO scheduling properties still require operator
qualification. If individual scheduling properties are not exposed, the GUI
requires a human decision instead of certifying the batch date for every user.

Individual campaign `.cmd` launchers are intentionally not shipped; use the common GUI launcher so pre-run template/config checks are applied consistently.

## Configuration

Create local copies from the committed templates:

```text
Config/ExchangeUserNotifications.local.json.template
Config/Campaigns/ExchangeMigration.local.json.template
Config/Campaigns/ExchangeArchive.local.json.template
Config/Campaigns/ExchangeMigrationMailboxSizeReduction.local.json.template
```

Real files must use `.local.json` and remain ignored by Git.

`Config/ExchangeUserNotifications.local.json` controls the shared mail transport:

- `MailSendMode = Graph` sends through Microsoft Graph `/sendMail` after a standard delegated interactive sign-in.
- `MailSendMode = SmtpRelay` forces SMTP relay mode and requires `SmtpServer`.
- `MailSendMode = Auto` keeps compatibility: Graph is used when `SmtpServer` is empty, otherwise SMTP relay is used.
- `MailSendMode = Disabled` skips mail sends while keeping campaign processing/logging behavior.

For Graph sending, the signed-in account needs the Microsoft Graph delegated permission `Mail.Send`. The sending mailbox is resolved from `From` in the tenant profile or communications config. If it differs from the signed-in account, that account must also have the required Exchange Send As permission.

SMTP relay mode does not use Microsoft Graph mail permissions. It requires `SmtpServer` or `RelayIp`, `SmtpPort`, and any relay allow-listing needed by the local mail infrastructure.

User-facing Teams chat messages are optional and disabled by default:

- `TeamsUserMessageMode = Disabled` sends email only.
- `TeamsUserMessageMode = GraphDelegated` sends a one-on-one Teams chat message in addition to the email.

Teams user messages are separate from the operational Teams summary notifications. They use Microsoft Graph delegated permissions (`Chat.Create` and `ChatMessage.Send`) because standard Teams chat posting is delegated-only for normal messages. The GUI exposes this as `Send Teams message`; the checkbox stays disabled until `TeamsUserMessageMode = GraphDelegated` is explicitly configured in local communications or campaign config and the campaign has `TeamsUserMessageByLanguage` content. Dry runs preview the Teams path without posting.

The Teams sender name shown to users is the delegated Microsoft 365 account connected to Graph for the campaign run. Use a dedicated communications account if messages should appear from a generic sender rather than the operator.

Campaign JSON controls Exchange lookup/connectivity:

- `ExchangeManagementMode = Auto` tries Exchange Online first and falls back to the Exchange 2016 snap-in when `EnableExchange2016Fallback` is true.
- `ExchangeManagementMode = ExchangeOnline` forces EXO only.
- `ExchangeManagementMode = Exchange2016` forces the legacy Exchange 2016 snap-in.
- `ExchangeManagementMode = Disabled` skips live Exchange mailbox checks.
- `RequireExchangeManagement = true` makes the campaign stop when no Exchange management source is available.

Campaign templates are shipped in each campaign `Templates` folder. Do not edit those files for customer wording: Git resyncs can update them and create conflicts. For local customizations, copy the needed HTML files to a local ignored folder such as `ExchangeMigration/Templates.local`, then set the campaign `.local.json` value:

```json
"TemplateRootPath": "{{CampaignRootPath}}\\Templates.local"
```

The scripts look in `TemplateRootPath` first and fall back to the shipped `Templates` folder, so a partial local override can coexist with new templates delivered by Git.

Exchange Online and Microsoft Graph use standard interactive authentication. The tenant profile `TenantId` is used to validate or constrain the selected tenant; certificate and app-only authentication is not used by this application.

## Output

Campaign data is tenant-isolated by default:

```text
SmartM365/Data/Tenants/<TenantKey>/DATA-ALL/Exchange/ExchangeUserNotifications/<Campaign>
SmartM365/Data/Tenants/<TenantKey>/DATA-LAST
SmartM365/Data/Tenants/<TenantKey>/LOG-ALL/Exchange/ExchangeUserNotifications/<Campaign>
```

Each campaign writes a timestamped run log CSV and a local sent registry CSV.

With `PreventResendAcrossRuns` enabled, live runs take an exclusive lock on the
campaign registry before reading it. A second run using the same registry is
refused. Each confirmed email is flushed to disk and atomically recorded before
optional Teams messaging. If the checkpoint cannot be saved, the campaign stops
before sending to another recipient. The adjacent `.lock` file remains on disk;
its presence alone does not indicate an active run. The operating system releases
the lock when the process exits, including forced termination.

Email or user-facing Teams delivery failures produce an error notification and a
terminating error (nonzero process exit), after recipient processing and summary
generation. Successful emails stay recorded even if their Teams message fails.
Deliberately skipped recipients and successful dry runs remain normal completions.
Mail transport acknowledgement and the disk checkpoint are separate operations;
an interruption between them still requires checking the delivery log before a
replay.

For size reduction, both the license CSV and its live Graph fallback apply
`TargetSkuPartNumbers` before choosing a quota from `MailboxQuotaBySkuPartNumber`.

Offline regression command from the repository root:

```powershell
pwsh -NoProfile -File SmartM365/Tests/Test-SmartM365ExchangeCampaignReliabilityOffline.ps1
```

The same suite supports Windows PowerShell 5.1. It uses fictional recipients,
mocked mail/Graph/Teams boundaries and temporary files, including a child process
terminated after a checkpoint. It does not send messages or authenticate.

## Offline demo mode

Launch `Start-SmartM365-ExchangeUserNotifications-Demo.cmd` for screenshots and presentations. The original WPF screens, tabs and options remain available with fictional Contoso data and visibly marked simulated results. Use the Scene selector (`Overview`, `Issues`, `Completed`), or pass `-DemoScenario Issues`. `-DemoView` selects a named tab where available; `-DemoWindowWidth` and `-DemoWindowHeight` control the capture size.

Business actions are simulated. Demo startup bypasses operational configuration, authentication, collection, remote commands and mail delivery. Edits remain in memory. **Export demo outputs** generates only synthetic files under a unique `%TEMP%\SmartM365-Demo\` session folder. HTML-capable tools retain their report views; notification campaigns preview every shipped language template with fictional tokens. No real tenant or device evidence is required.

Keep the application folder, its original report/template sources, local branding assets, `SmartM365.GuiDemo.ps1`, `SmartM365.GuiDemo.Reports.ps1` and the local `*-Demo.ps1` profile together when copying the application. `-Demo -ValidateOnly` builds the WPF demo without displaying windows or opening files. Run `SmartM365/Tests/Test-SmartM365GuiDemoApplicationsOffline.ps1` in an STA process for offline interaction and export checks.


## Local review correction (2026-10-18)

Source version: GUI 1.0.5. Campaign stdout and stderr are read asynchronously and drained by a WPF dispatcher timer on the GUI thread. Completion waits for process exit and both streams to reach EOF; cancellation, startup failure and repeated runs clean up the monitor. No PowerShell scriptblock is attached to Process output/exit events on worker threads.

Offline regression command from the repository root:

```powershell
pwsh -NoProfile -STA -ExecutionPolicy Bypass -File SmartM365/Tests/Test-SmartM365GuiReviewRegressionsOffline.ps1 -Case NotificationProcess
```

These checks use temporary files, harmless child processes or mocked operational boundaries. They do not qualify a real tenant, endpoint repair, campaign delivery or deployment. Demo assets remain local in `DemoAssets/`, ignored by Git, and must be refreshed after GUI changes. Live deployment requires separate qualification.
