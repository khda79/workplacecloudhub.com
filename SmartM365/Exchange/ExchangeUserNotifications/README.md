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

Individual campaign `.cmd` launchers are intentionally not shipped; use the common GUI launcher so pre-run template/config checks are applied consistently.

## Configuration

Create local copies from the committed templates:

```text
Config/Communications.local.json.template
Config/Campaigns/ExchangeMigration.local.json.template
Config/Campaigns/ExchangeArchive.local.json.template
Config/Campaigns/ExchangeMigrationMailboxSizeReduction.local.json.template
```

Real files must use `.local.json` and remain ignored by Git.

`Config/Communications.local.json` controls the shared mail transport:

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
SmartM365/Data/Tenants/<TenantKey>/DATA-ALL/Communications/ExchangeUserNotifications/<Campaign>
SmartM365/Data/Tenants/<TenantKey>/DATA-LAST
SmartM365/Data/Tenants/<TenantKey>/LOG-ALL/Exchange/ExchangeUserNotifications/<Campaign>
```

Each campaign writes a timestamped run log CSV and a local sent registry CSV.
