# Exchange User Notifications

User-facing Exchange communication campaigns.

## Scripts

- `ExchangeMigration/SmartM365-ExchangeMigration-NotifyUsers.ps1`: sends migration notifications from a recipients CSV or folder.
- `ExchangeArchive/SmartM365-ExchangeArchive-NotifyUsers.ps1`: sends archive mailbox activation notifications from a recipients CSV or folder.
- `ExchangeMigrationMailboxSizeReduction/SmartM365-ExchangeMigrationMailboxSizeReduction-NotifyUsers.ps1`: sends pre-migration mailbox size reduction notifications. The license CSV selects users matching `TargetSkuPartNumbers`; when it is missing or older than `LicenseCsvMaxAgeHours`, `EnableLiveLicenseLookupWhenCsvMissing` lets the script resolve licenses live from Microsoft Graph. `MailboxQuotaBySkuPartNumber` sets the quota per license, for example E3 at 100 GB and F1/F3 at 2 GB. Exchange 2016 is mandatory and verifies mailbox existence, excludes `RemoteMailbox` objects by default, and reads the live mailbox size before applying the per-user quota.

## Local Module

`SmartM365.Communications.psm1` contains campaign-local helpers:

- JSON campaign configuration loading.
- CSV import with BOM removal and delimiter auto-detection.
- Template token expansion and unresolved-token checks.
- Sent registry load/save.
- Graph or SMTP mail sending with optional BCC.
- Summary HTML generation.

The module is intentionally local to this area and should not be promoted to `SmartM365.Core` unless other SmartM365 product areas need it.

## GUI

`GUI/SmartM365-ExchangeUserNotifications-GUI.ps1` is a common Windows Forms launcher for the three campaigns. It validates script/config files, builds the command line, defaults to dry-run mode, and displays campaign output while keeping all business logic in the existing scripts.

Use `GUI/Start-SmartM365-ExchangeUserNotifications-GUI.cmd` to launch it with PowerShell 7.

## Configuration

Create local copies from the committed templates:

```text
Config/Communications.local.json.template
Config/Campaigns/ExchangeMigration.local.json.template
Config/Campaigns/ExchangeArchive.local.json.template
Config/Campaigns/ExchangeMigrationMailboxSizeReduction.local.json.template
```

Real files must use `.local.json` and remain ignored by Git.

## Output

Campaign data is tenant-isolated by default:

```text
SmartM365/Data/Tenants/<TenantKey>/DATA-ALL/Communications/ExchangeUserNotifications/<Campaign>
SmartM365/Data/Tenants/<TenantKey>/DATA-LAST
SmartM365/Data/Tenants/<TenantKey>/LOG-ALL/Communications/ExchangeUserNotifications/<Campaign>
```

Each campaign writes a timestamped run log CSV and a local sent registry CSV.
