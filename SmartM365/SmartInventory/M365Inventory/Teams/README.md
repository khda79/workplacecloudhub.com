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

Runtime dependency:

- `ImportExcel` builds the timestamped workbook. If it is missing, the script installs it automatically from `PSGallery` for the SmartM365 execution account (`CurrentUser`).

Required Microsoft Graph application permissions:

- `Team.ReadBasic.All`
- `TeamMember.Read.All`
- `Channel.ReadBasic.All`
- `Group.Read.All`
- `Reports.Read.All`

Optional permissions for richer detail:

- `ChannelMember.Read.All` for private/shared channel owner lookup when `-IncludeChannelOwners` is used.
- `Sites.Read.All` may be required by Graph in some tenants for exact SharePoint drive quota lookup per team.

Examples:

```powershell
.\SmartM365-Teams-Inventory.ps1 -Tenant prod
.\SmartM365-Teams-Inventory.ps1 -Tenant prod -MaxTeams 10 -DryRun
.\SmartM365-Teams-Inventory.ps1 -Tenant prod -InactiveDays 180 -AlwaysSend
.\SmartM365-Teams-Inventory.ps1 -Tenant prod -AppendHistory -IncludeChannelOwners
```
