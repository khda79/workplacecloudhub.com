# Microsoft 365 usage signals

`scripts/New-M365UsageSignalsEvidence.ps1` (1.0.1) prepares four optional evidence families from
reports added in SmartInventory UserActivity 1.34 and CopilotUsage 1.1.0. It runs in the
`Collaboration` domain of `Invoke-PreparedEvidenceBuild.ps1` (1.8.0), after
`New-CollaborationEvidence.ps1`, and does not change the collaboration tables or their history.

| Source CSV (DATA-LAST) | Producer | Prepared file | Evidence table | Grain |
| --- | --- | --- | --- | --- |
| `M365_Email_AppUsage.csv` | UserActivity `EmailAppUsageUserDetail` | `EmailClientUsageEvidence.csv` | Email Client Usage Evidence | One row per email client |
| `M365_OneDrive_UserActivity.csv` | UserActivity `OneDriveActivityUserDetail` | `OneDriveActivityEvidence.csv` | OneDrive Activity Evidence | One row per activity metric |
| `M365_Groups_Activity.csv` | UserActivity `Office365GroupsActivityDetail` | `GroupActivityEvidence.csv` | Group Activity Evidence | One row per live Microsoft 365 group |
| `M365_Copilot_UserCountTrend.csv` | CopilotUsage trend | `CopilotDailyAdoptionEvidence.csv` | Copilot Daily Adoption Evidence | One row per report date and app |

## Optional sources

The four CSVs are `optionalCurrentFiles` in `config/prepared-source-contract.json.txt` with a
744-hour publication-age limit, the same limit as the other Microsoft 365 usage reports. The
limit of an optional file comes from the contract, not from `PreparedSourceAgeOverrides`. A file
that is absent, older than 744 hours, empty or unreadable is recorded in the capture's
`OptionalSourceIssues` and the batch continues.

Without a usable source the generator never writes zeros:

- Email client and OneDrive activity keep their stable rows (nine clients, five metrics) with empty
  counts and `Evidence Status` = `Not collected`, or `Schema mismatch` when required columns are
  missing.
- Group and Copilot daily tables publish their header only (`allowEmpty` in the evidence contract).

Data Trust (1.3.0) lists the four sources. An absent or empty optional source is
`Decision Readiness` = `Optional unavailable` and `Evidence Status` = `Not collected`, never
`Blocked`. A present source follows the usual schema and freshness rules on its report refresh date.

## Rules

- Deleted users and deleted groups are excluded. A group repeated by Graph is kept once.
- Email clients: Graph writes the client version in each column (for example `ProPlus`,
  `Outlook 2016` or `Undetermined`) and leaves unused clients empty, so any non-empty value counts
  the user. `User Share (%)` is
  the share of users with at least one client. POP3, IMAP4 and SMTP are flagged `Legacy Protocol`.
  Only aggregates are published; no per-user list.
- OneDrive: `Users` counts users with a non-zero count over the report period, `Files` sums the
  counts. `Any OneDrive activity` counts users with a last activity date.
- Groups: `Teams State` is `Teams-enabled` when the group ID is a team of `M365_Teams_Teams.csv`.
  Activity, owner and external-access signals use the collaboration object vocabulary
  (`Group without owner`, `Inactive group`, `Group with external members`), measured at the report
  refresh date. The owner principal name is not published.
- Copilot daily: an app with neither enabled nor active users in a row (for example v2-only apps
  after a v1 fallback) leaves no row; reported zeros stay. `Prompts Submitted` is only on `Any app`.
- Freshness uses the collaboration bands: `Current` up to 14 days, `Aging` up to 45, then `Stale`.

The Copilot table name deliberately avoids `Trend` and `History`: those tables are compared with
the previous batch and refused when a key disappears, while Microsoft's trend is a rolling window.

## Tests

`scripts/Test-M365UsageSignalsEvidence.ps1` (PowerShell 7, synthetic data only) covers the
contracts, source selection, absent, header-only, populated and drifted reports, Data Trust and
the evidence contract validator.

Power BI tables, measures and views are a separate lot.
