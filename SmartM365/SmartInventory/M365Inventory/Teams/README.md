# Microsoft Teams Inventory

`SmartM365-Teams-Inventory.ps1` inventories Microsoft Teams through Microsoft Graph app-only authentication and exports Power BI-ready CSV files for Teams, Members, Channels, and Guests.

The daily email contains SharePoint links only: the four timestamped CSV files and a timestamped `M365_Teams_Inventory_yyyyMMdd_HHmmss.xlsx` workbook with one worksheet per CSV. Files are not attached and local paths are never shown. If a SharePoint upload does not return a web URL, that file link is omitted.

Latest CSV files:

| Entity | Latest CSV |
| --- | --- |
| Teams | `M365_Teams_Teams.csv` |
| Members | `M365_Teams_Members.csv` |
| Channels | `M365_Teams_Channels.csv` |
| Guests | `M365_Teams_Guests.csv` |

## Collection completeness and publication

The Teams inventory follows every Graph `@odata.nextLink` for team groups,
owners, members and channels. A missing `value` property or a repeated next link
is treated as an incomplete response. Transient HTTP 408, 409, 429, 500, 502,
503 and 504 responses are retried up to five attempts; `Retry-After` is honored
when Graph supplies it.

Team groups, the D180 Teams activity report, team details, owners, members and
channels are mandatory for publication. If one of those collections still fails
after retries, the run fails before the four CSV files are written. Exact
SharePoint drive quota and private/shared channel-owner expansion remain optional
measurements: unavailable values stay blank or carry `NotMeasured` in their
documented detail field instead of being invented.

In version 0.27, each complete CSV is first written to its timestamped `DATA-ALL` path and then
promoted atomically to its canonical `DATA-LAST` path. A failed serialization or
promotion cannot truncate or replace the previous `DATA-LAST` file. The four
files share one `RunId`, `RunDateUtc` and filename timestamp. Append history is
also updated atomically. Weekly history and
the workbook are produced only after all four CSV publications succeed. This is
per-file atomicity, not a multi-file filesystem transaction.

Runtime dependency:

- `ImportExcel` builds the timestamped workbook. If it is missing, the script installs it automatically from `PSGallery` for the SmartM365 execution account (`CurrentUser`).

Required Microsoft Graph application permissions:

- `Team.ReadBasic.All`
- `TeamMember.Read.All`
- `Channel.ReadBasic.All`
- `Group.Read.All`
- `Reports.Read.All`
- `Sites.Read.All` (enforced by the current preflight for SharePoint drive quota lookup)

Optional permissions for richer detail:

- `ChannelMember.Read.All` for private/shared channel owner lookup when `-IncludeChannelOwners` is used.

Examples:

```powershell
.\SmartM365-Teams-Inventory.ps1 -Tenant prod
.\SmartM365-Teams-Inventory.ps1 -Tenant prod -MaxTeams 10 -DryRun
.\SmartM365-Teams-Inventory.ps1 -Tenant prod -InactiveDays 180 -AlwaysSend
.\SmartM365-Teams-Inventory.ps1 -Tenant prod -AppendHistory -IncludeChannelOwners
```

## Teams Phone PSTN usage

`SmartM365-TeamsPhonePstnUsage-Inventory.ps1` is a read-only app-only collector for Teams Phone Calling Plans, Operator Connect, and Direct Routing usage. It uses the official Microsoft Graph v1.0 functions:

- `GET /communications/callRecords/getPstnCalls`
- `GET /communications/callRecords/getDirectRoutingCalls`

The Graph functions do not support delegated authentication. The minimum required Microsoft Graph application permission is `CallRecords.Read.All`.

Latest CSV files:

| Entity | Latest CSV | Notes |
| --- | --- | --- |
| SmartFinOps user aggregate | `M365_Teams_PhoneUserUsage.csv` | Primary user-level usage file. No license price, allocation, or savings calculation is performed. |
| PSTN and Operator Connect calls | `M365_Teams_PSTNCalls.csv` | Detailed Graph rows with caller and callee numbers masked. |
| Direct Routing calls | `M365_Teams_DirectRoutingCalls.csv` | Detailed Graph rows with caller and callee numbers masked. |
| Phone assignments | `M365_Teams_PhoneAssignments.csv` | Created only when `-IncludePhoneAssignments` is requested and the supported Teams PowerShell prerequisites are available. |

Dates are normalized to UTC. Ranges longer than the Graph PSTN maximum request window are split into windows of at most 90 days. Every page in `@odata.nextLink` is followed, including result sets larger than 1,000 rows. Transient failures and throttling honor `Retry-After` when available.

Version 1.5 rejects a Graph collection response without `value` and repeated
`@odata.nextLink` values. A failure on a later page throws without returning the
earlier pages to the caller. Publication continues to use the shared atomic CSV
helper, so the existing `DATA-LAST` file is not replaced by an incomplete
collection. SmartFinOps consumes `M365_Teams_PhoneUserUsage.csv`; it does not
consume the four Teams collaboration inventory files above.

Examples:

```powershell
.\SmartM365-TeamsPhonePstnUsage-Inventory.ps1 -Tenant test -ValidateOnly
.\SmartM365-TeamsPhonePstnUsage-Inventory.ps1 -Tenant test -LookbackDays 90 -MaxItems 100
.\SmartM365-TeamsPhonePstnUsage-Inventory.ps1 -Tenant prod -FromDate '2026-04-01T00:00:00Z' -ToDate '2026-07-01T00:00:00Z'
.\SmartM365-TeamsPhonePstnUsage-Inventory.ps1 -Tenant prod -IncludePhoneAssignments
```

`-IncludePstn` and `-IncludeDirectRouting` default to true. Pass `-IncludePstn:$false` or `-IncludeDirectRouting:$false` to disable one source. `-MaxItems` produces `MAXITEMS-<n>` test filenames and does not replace the canonical DATA-LAST files or WeeklyHistory.

Phone assignment inventory is optional because it uses the officially supported `Get-CsPhoneNumberAssignment` cmdlet rather than Microsoft Graph call-record APIs. Its prerequisites are:

- MicrosoftTeams PowerShell module 4.7.1 or later for app-only authentication. When it is missing, the SmartM365 preflight installs the current PSGallery version with `Scope=CurrentUser`.
- Microsoft Graph application permission `Organization.Read.All` for Teams PowerShell app authentication.
- A supported Microsoft Entra Teams role assigned to the service principal. `Teams Telephony Administrator` is the most voice-focused documented role; `Teams Communications Administrator` and `Teams Administrator` also cover phone number inventory.

The Teams role can perform more than read operations, but this collector calls only `Get-CsPhoneNumberAssignment`. If any prerequisite is missing, the optional assignment phase fails explicitly and no assignment CSV is simulated.

The Graph PSTN endpoint does not return Telstra calling-plan detail. Microsoft also applies country-specific call-detail retention and phone-number obfuscation rules. This collector applies an additional local mask to caller, callee, and assignment numbers and never collects communication content.
