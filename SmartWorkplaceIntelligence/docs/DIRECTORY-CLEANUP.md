# Directory cleanup evidence

`scripts/New-DirectoryCleanupEvidence.ps1` (1.0.0) identifies AD user accounts, AD workstation
accounts and user mailboxes that are candidates for cleanup because none of their qualified
activity signals is more recent than one year. It runs as the `DirectoryCleanup` domain of
`Invoke-PreparedEvidenceBuild.ps1` (1.9.0), after `Workforce` (User Inventory Evidence) and
`Estate` (All Devices Evidence), with its own account classification checkpoint.

The evidence is read-only. It recommends an action; it never disables, deletes or changes an object.

| Prepared file | Evidence table | Grain |
| --- | --- | --- |
| `UserCleanupEvidence.csv` | User Cleanup Evidence | One row per AD user account |
| `ComputerCleanupEvidence.csv` | Computer Cleanup Evidence | One row per AD workstation account |
| `MailboxCleanupEvidence.csv` | Mailbox Cleanup Evidence | One row per user mailbox (a conflicting hybrid pair gives one row per side) |
| `DirectoryCleanupSourceQualification.csv` | Directory Cleanup Source Qualification | One row per source, plus the user mailbox scope |

The three object tables may be empty. The populations are distinct objects: users, computers and
mailboxes must not be added together as people.

## Scope

- All AD user accounts, enabled and disabled.
- AD computer accounts except servers. A server operating system is out of scope. An empty or
  unknown operating system keeps the computer with `Scope State` = `Undetermined`, which is Review.
- User mailboxes only (`UserMailbox` in Exchange Online and on-premises). Shared, room, equipment
  and discovery mailboxes are out of scope and counted in the source qualification.

## Sources and qualification

| Source | File | Used for |
| --- | --- | --- |
| AD users | `AD_Users_AllDomains.csv` | Account, domain, AD logon, account type |
| AD computers | `AD_Computers_AllDomains.csv` | Account, domain, AD logon, Entra sign-in of the matched device |
| Entra users | `M365_Users_Active.csv` | Sign-ins (interactive, non-interactive, successful) |
| Entra devices | `M365_Entra_Devices.csv` | Proof that one Entra device has the AD object GUID |
| Intune devices | `Intune_ManagedDevices_All.csv` | Last sync |
| Exchange Online | `Exchange_EXO_Mailboxes_AllDomains.csv` | Mailboxes, holds, last user action and logon |
| Exchange on-premises | `Exchange_OnPrem_Mailboxes_AllDomains.csv` | Mailboxes, holds, in-place holds, last logon |
| Microsoft 365 user activity | `M365_Users_Activity.csv` | Last activity in any workload |
| Mailbox usage | `M365_Mailbox_Usage.csv` | Mailbox last activity and storage |

Each source gets a status: `Observed`, `Not collected`, `Schema mismatch`, `Stale`,
`Freshness unknown` or `Partial coverage`. Freshness comes from the producer receipt
(`SmartInventory_*.current.json.txt`): completed, file record `Success`, SHA256 equal to the
captured file and not partial. The completion time must be within `SourceMaxAgeHours` (default
336) of the reference. Graph reports must also have a `Report Refresh Date` within
`ReportMaxAgeDays` (default 31). The file time alone never proves freshness. A source that is not
`Observed` makes every object that depends on it Review, never a candidate. An old optional file
kept by Prepare (`StaleUsed`) is therefore still `Stale` here.

Only ISO 8601 values with an explicit offset are read and converted to UTC. A value without a
timezone is `Unqualified timezone`. Graph report dates (`yyyy-MM-dd`) are UTC days. For computers,
the Entra sign-in is the offset-qualified `EntraApproximateLastSignInDateTime` of the AD export;
the offset-less column of `M365_Entra_Devices.csv` is not read.

## Matching

- **AD user to Entra user**:
  - `ObjectSID` = `OnPremisesSecurityIdentifier`, otherwise `ImmutableId_AD` = `OnPremisesImmutableId`.
  - Exactly one object gives `Matched`.
  - Several candidates, or a SID and an ImmutableId naming different objects, give `Ambiguous`.
  - A UPN-only match is `Weak`: its signals never count and the account is Review.
- **User Inventory classification** (Country, User Type, Workforce Persona, console activity) is
  reused only through the strongly matched Entra object id (`Classification Match Status`).
  Otherwise Workforce Persona and User Type are `Unknown` and the country comes from the AD
  attribute (`Country Basis`).
- **Account Type** always comes from AD `AccountType`. **Account Population** comes from the
  governed `SmartM365.AccountClassification` rules (`Account Classification Rule Version`). A
  `Review required` population (for example generic or admin accounts) is a secondary reason, not
  a block.
- **Graph reports** are keyed by UPN only. They are read through the UPN of the strongly matched
  Entra object:
  - one row is `Qualified` ;
  - no row is `No report row`, which proves nothing ;
  - several rows are `Ambiguous match`.
- **Workstations**: AD `ObjectGUID` = Entra `DeviceId` (hybrid join), with
  `EntraMatchCountByObjectGuid` = 1, and Intune `AzureADDeviceId` = Entra `DeviceId`.
  - `No Entra object` is established only when the AD export and a qualified Entra device
    capture both have no match.
  - Several or inconsistent matches are Review.
- **Mailboxes**:
  - The key is the mailbox GUID (`MailboxGuid` online, `ExchangeGuid` on-premises).
  - The owner is the AD user by ImmutableId (online) or by ObjectGUID (on-premises).
  - The same GUID on both sides is merged into one `Hybrid (both observed)` row only when both
    sides name the same owner. Otherwise each side is a Review row (`Conflicting hybrid identity`).
  - A missing or duplicated GUID is not published: the rows are counted as `Unkeyable Rows` and
    a warning is written.
  - An owner that is not found is Review, with the action `Review owner`.

`Not applicable` is set only by an explicit rule (no mailbox established by AD and two qualified
Exchange captures, no Entra device established as above), never by a missing match.

## Activity

Every positive signal counts, so that one activity on any source keeps an object out of the
candidates. Each signal has its own date column and a `<Signal> Quality` column:
- `Qualified` and `Clock skew` count as activity ;
- `No value`, `No report row`, `Weak match`, `Ambiguous match`, `Future date`,
  `Unqualified timezone` and `Not applicable` do not count ;
- a source status (`Stale`, `Not collected`...) replaces the quality when the source is not usable.

| Object | Signals |
| --- | --- |
| User | AD logon, Entra interactive, non-interactive and successful sign-in, EXO last user action, EXO last logon, on-premises last logon, mailbox usage, Microsoft 365 activity, console |
| Computer | AD logon, Entra sign-in, Intune sync |
| Mailbox | EXO last user action, EXO last logon, on-premises last logon, mailbox usage, owner AD logon, owner Entra sign-in, owner Microsoft 365 activity, owner console |

- `Last Activity Date` is the latest counted signal.
- On a tie, `Last Activity Source` is the first signal in the order of the table above.
- A date up to one day after the reference is set to the reference (`Clock skew`).
- A later date is `Future date`.
- The reference is `Reference UTC`, the batch clock (`-AsOfUtc`). `Evidence Date` is its UTC day.

## Buckets and decision

Buckets use calendar years computed with `AddYears`, so the one-year limit of 29 February is
28 February:

| Inactivity Bucket | Bucket Sort | Rule (R = reference) |
| --- | --- | --- |
| `< 1 year` | 1 | Last Activity Date after R - 1 year |
| `1-2 years` | 2 | after R - 2 years, up to R - 1 year |
| `2+ years` | 3 | R - 2 years or earlier |
| `No observed activity` | 4 | no counted signal |

`Cleanup State`, in order of priority:
1. `Excluded - hold`: Litigation Hold, Retention Hold or in-place holds on the mailbox (for a user,
   on its linked mailbox).
2. `Excluded - non-human`: non-human account population.
3. `Review` for an imperative reason:
   - undetermined scope ;
   - owner not found or ambiguous ;
   - weak or ambiguous match ;
   - an applicable source not observed ;
   - unknown mailbox presence or hold.
4. `Active`: activity within one year.
5. `Review` for a decisive reason:
   - a future date or a date without timezone ;
   - the AD latency margin ;
   - no observed activity without a known creation date older than one year.
6. `Candidate`.

`State Sort` is 1 for Active, 2 for Candidate, 3 for Review, 4 for Excluded - hold and 5 for
Excluded - non-human. `Decision Reason` holds the main reason. `Secondary Reasons` holds every
other finding, including recent activity hidden by a higher rule.

**AD latency.** When the latest date comes only from AD `LastLogonTimestamp` and lies within
`AdLatencyDays` (default 14) before the one-year limit, the state is Review
(`Near threshold (AD replication latency)`). At the two-year limit only a secondary reason is
added, because both buckets are candidates.

**Holds.** Only mailbox-level holds are collected. Organization retention policies (Purview) are
not. A mailbox without a known hold can be a candidate, with the secondary reason
`Organization retention policies not assessed`.

`Recommended Action`:

| Object | Candidate | Review | Otherwise |
| --- | --- | --- | --- |
| User | `Disable account` when enabled, `Delete account` when disabled | `Review evidence` | `Keep` |
| Computer | `Disable computer` when enabled, `Delete computer object` when disabled | `Review evidence` | `Keep` |
| Mailbox | `Remove mailbox after retention check` when the owner is disabled, otherwise `Review owner` | `Review owner` when the owner is not found or ambiguous, otherwise `Review evidence` | `Keep` |

## Data Trust

Data Trust (1.4.0) also lists the Exchange Online and on-premises mailbox inventories, the
all-platform Intune device inventory and mailbox usage (optional).

## Tests

`scripts/Test-DirectoryCleanupEvidence.ps1` (PowerShell 7, synthetic data only) covers:
- each state and action ;
- strong, weak and ambiguous matches ;
- holds ;
- servers and unknown operating systems ;
- shared mailboxes ;
- conflicting hybrid mailboxes and duplicated GUIDs ;
- the latency margin, future dates, clock skew and dates without timezone ;
- the 29 February boundaries ;
- stale, missing and unqualified sources ;
- the evidence contract validator.

## Power BI

The model imports User Cleanup Evidence, Computer Cleanup Evidence, Mailbox Cleanup Evidence,
and Directory Cleanup Source Qualification. The Directory Cleanup page has separate Users,
Computers, and Mailboxes views, with counts by state and inactivity bucket, candidate counts by
AD domain and country, detail tables, and source qualification. The three populations are not
summed as people. State and bucket order use the published sort columns.

User and mailbox details include `Persona Basis` and `Account Type Source`. Persona Basis
identifies whether the workforce persona came from User Inventory, AD attributes, or is unknown.
Account Type Source identifies an external source, SWI rules, or an unknown source. Account Type
values are not a fixed list; use Account Population to group accounts for cleanup decisions.

Country and domain selectors use the prepared country and AD domain values. For computers,
country inferred from an AD domain or computer-name prefix is displayed as Unknown; Country
Basis remains available in the details. The page retains the warning that organization retention
policies have not been assessed. Recommended actions are evidence for review, not automated
directory or mailbox changes.
