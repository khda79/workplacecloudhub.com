# Hybrid identity findings and device readiness gaps

Status: implemented and validated offline (synthetic tests and a local comparison
with the legacy outputs). Not yet run by the server preparation. No Power BI change
is part of this work. The two legacy SmartInventory producers stay in place,
unchanged, until the new prepared outputs have been compared on the server; they
are retired in a later, separate change.

| Legacy producer | Legacy output | Replaced by |
| --- | --- | --- |
| `SmartM365-Exchange-HybridIdentity-Issues-Inventory.ps1` | `Exchange_HybridIdentity_Issues.csv` | `HybridIdentityFindings.csv` (task `HybridIdentity`, `scripts/New-HybridIdentityEvidence.ps1`) |
| `SmartM365-Intune-Windows11-Readiness-Issues-Inventory.ps1` | `Intune_Windows11_Readiness_Issues.csv` | `DeviceReadinessGapEvidence.csv` (optional output of `DeviceInventory`, `scripts/DeviceReadinessGapEvidence.psm1`) |

## Principles

1. **No private rule copies.** A rule that SWI already evaluates elsewhere is
   reused, never re-implemented with its own lists or thresholds. Account
   population comes from the governed `AccountClassification` configuration
   through the shared module (see *Account population*); duplicate UPN uses the
   same source as Workforce `Duplicate UPN State`.
2. **Only what SWI does not already cover.** Size, archive, inactivity,
   delegation and licensing gaps already exist in Mailbox Evidence, Mailbox
   Exploration Inventory and the License Report. Windows 11 eligibility, disk,
   build, Windows Update, Secure Boot disabled, hardware ID conflicts,
   compliance, sync, removal candidates, AD without Intune and LTSC already
   exist in Device Inventory Evidence and Device Deployment Risk.
3. **Reference clock.** Dates and ages use `-AsOfUtc` when the build passes it
   (`$clockArgs`); otherwise each generator reads the current UTC time once, like
   the other generators. The same inputs and reference give the same bytes.
4. **Unknown is not zero.** A blank, `Error` or unparsable value is unknown: the
   rule is counted as *not assessed* for that object and produces no finding.
5. **Stable keys.** Rows carry every strong identifier available so the private
   model can join on its existing keys without name matching.

## Inputs

| Source | Capture | Used for |
| --- | --- | --- |
| `AD_Users_AllDomains.csv` | current | UPN, proxyAddresses, TargetAddress, `AccountType`, SID, immutable ID |
| `AD_Users_DuplicateUPN.csv` | current | HYB-009 (same source as Workforce) |
| `AD_Users_DuplicateSMTP.csv` | optional (added, 168 h) | HYB-016; absent means not assessed |
| `Exchange_OnPrem_Mailboxes_AllDomains.csv` | current | on-premises mailboxes, item count, size in MB |
| `Exchange_OnPrem_RemoteMailboxes_AllDomains.csv` | native | remote mailboxes |
| `Exchange_EXO_Mailboxes_AllDomains.csv` | current | EXO mailboxes, item count, size in GB |
| `Exchange_EXO_Mailboxes_AllDomains_Stats.csv` | current | EXO statistics (item count, size in GB) |
| `Exchange_EXO_AcceptedDomains.csv` | optional (added, 168 h) | routing domain, HYB-030, HYB-032; absent means not assessed |
| `M365_Entra_VerifiedDomains.csv` | native | HYB-035 |
| `M365_Users_Active.csv`, `M365_Licenses_Users.csv` | current | Entra presence and licenses |
| `UserInventoryEvidence.csv` | staged Workforce output | `Account Population` |
| `Intune_Devices_Inventory.csv`, `Intune_Devices_LocalSystem.csv`, `M365_Entra_Devices.csv`, `AD_Computers_AllDomains.csv` | current | device facts |

The remote routing domain is the single accepted domain matching
`*.mail.onmicrosoft.com`. If accepted domains are absent, or none or several
match, HYB-003 and HYB-021 are not assessed, even when additional domains are
configured: the configured domains extend the routing domain and are not an
exhaustive list of allowed TargetAddress domains.

Additional TargetAddress domains come from the private wrapper setting
`PreparedHybridAllowedTargetDomains` (array, empty in the template, never
committed with values). The wrapper passes it to the pipeline, the pipeline
passes one comma-separated value to the build (`-HybridAllowedTargetDomains`),
and the producer splits it.

## Account population

`Get-SmartM365AccountPopulation` (module `SmartM365.AccountClassification`
1.2.0) maps an AD `AccountType` to `Human`, `Non-human` or `Review required`
from the `Population` section; any type not listed is `Review required`.
Workforce 1.4.1 calls it through its existing `Get-ConfiguredAccountPopulation`;
its outputs are unchanged.

Workforce writes two values in `UserInventoryEvidence.csv` and never writes
`Non-human` (those accounts are excluded from it):

| Workforce `Account Population` | Classifier value used | Rule |
| --- | --- | --- |
| `Human` | `Human` | |
| `Workforce - review classification` | `Review required` | never upgraded to `Non-human` |
| (account absent from Workforce) | classifier on the AD `AccountType` | see below |

For an account absent from Workforce, the classifier is applied to every AD
account linked to the same Entra user through the Workforce strong keys (SID,
immutable ID, UPN). All linked AD accounts must agree; any disagreement gives
`Review required`, as in Workforce. A blank `AccountType` is
`Unclassified Account`, hence `Review required`. An uncertain classification is
never turned into `Non-human`.

The output records `Account Population`, `Account Population Source`
(`Workforce` or `AD AccountType`) and `Account Classification Rule Version`.

## Output 1: `HybridIdentityFindings.csv`

Grain: one row per account and rule (`Finding Key` = rule ID and AD object GUID,
unique). Task `HybridIdentity` (domain `HybridIdentity`, `-DependsOn 'Workforce'`,
classification checkpoint `BeforeHybridIdentity`), declared in
`config/prepared-evidence-contract.json.txt` with `stagingIncludesTenantKey`.

| Column | Meaning |
| --- | --- |
| `TenantKey`, `Finding Key` | tenant and unique key |
| `AD Object GUID`, `Entra Object ID`, `Exchange GUID` | strong keys; blank when not observed |
| `User Principal Name`, `Primary SMTP Address` | display and fallback keys |
| `Recipient Type`, `Mailbox Hosting` | `On-premises`, `Remote`, `Exchange Online`, `None` |
| `Rule ID`, `Legacy Issue Number` | new identifier and legacy number |
| `Rule Family`, `Severity`, `Is Blocker` | `Is Blocker` is true only for `Blocker` (legacy `1.Critical`) |
| `Observed Value`, `Recommended Action` | value that triggered the rule; neutral English action |
| `Account Population`, `Account Population Source`, `Account Classification Rule Version` | see above |
| `Rule Version`, `Evidence Date` | `HybridIdentityFindings/1.1`; reference UTC date |

With `-AuditPath`, the producer also writes input hashes, per-rule evaluated /
triggered / not-assessed counts and the routing decision (no address or domain).
The build log always shows the per-rule counts.

Scope, as in the legacy producer: AD accounts with a mailbox (on-premises,
remote or EXO) and a primary SMTP address; HYB-054 walks the EXO mailboxes.

### Retained rules

| Rule ID | Legacy | Family | Severity | Condition |
| --- | --- | --- | --- | --- |
| HYB-001 | 1 | Identity sync | Blocker | no Entra user matched by immutable ID or UPN |
| HYB-009 | 9 | Identity sync | Blocker | UPN listed in `AD_Users_DuplicateUPN.csv` |
| HYB-035 | 35 | Identity sync | Blocker | UPN domain not in verified domains |
| HYB-039 | 39 | Identity sync | Blocker | UPN contains an invalid character (see *Invalid characters*) |
| HYB-003 | 3 | Addressing | Blocker | proxyAddresses has no address in the remote routing domain |
| HYB-016 | 16 | Addressing | Blocker | primary SMTP or a proxy address listed in `AD_Users_DuplicateSMTP.csv` |
| HYB-024 | 24 | Addressing | Blocker | User/Shared/Room mailbox (local or remote) with no `SMTP:` entry |
| HYB-025 | 25 | Addressing | Blocker | same scope, more than one `SMTP:` entry |
| HYB-030 | 30 | Addressing | Blocker | primary SMTP domain not an EXO accepted domain |
| HYB-040 | 40 | Addressing | Blocker | primary SMTP local part contains an invalid character |
| HYB-026 | 26 | Addressing | High | same scope, primary SMTP absent from proxyAddresses |
| HYB-051 | 51 | Addressing | High | same scope, primary SMTP present in more than one proxy entry (case-insensitive) |
| HYB-031 | 31 | Addressing | Medium | primary SMTP domain is `onmicrosoft.com` |
| HYB-032 | 32 | Addressing | Info | user mailbox whose primary SMTP domain is an `InternalRelay` accepted domain |
| HYB-019 | 19 | Routing | Blocker | remote mailbox with empty TargetAddress |
| HYB-021 | 21 | Routing | Blocker | TargetAddress domain neither the routing domain nor an additional allowed domain; not assessed without a single routing domain |
| HYB-023 | 23 | Routing | Blocker | TargetAddress shared by two or more distinct accounts |
| HYB-020 | 20 | Routing | High | TargetAddress not of the form `local@domain.tld` |
| HYB-022 | 22 | Routing | Medium | TargetAddress absent from proxyAddresses |
| HYB-038 | 38 | Mailbox state | Blocker | on-premises `UserMailbox` and an EXO mailbox (split-brain) |
| HYB-054 | 54 | Mailbox state | Blocker | EXO mailbox whose AD account has no on-premises or remote mailbox |
| HYB-029 | 29 | Mailbox state | Medium | recipient type outside User/Shared/Room (local or remote) |
| HYB-002 | 2 | Migration capacity | High | item count greater than 999,999 |
| HYB-006 | 6 | Migration capacity | High | F3/F1 license group and mailbox size of 2000 MB or more |
| HYB-058 | 58 | License governance | Review | `Non-human` population and an E3, F3 or F1 license group |

### Routing rules 19 to 23

Validated explicitly, although rules 20 and 23 have no occurrence today. Their
legacy scope is kept: remote mailboxes (`RemoteUserMailbox`,
`RemoteSharedMailbox`, `RemoteRoomMailbox`). The `smtp:` prefix is removed in any
case before comparison, and all comparisons are case-insensitive. An empty
TargetAddress is HYB-019 only; HYB-020 to HYB-023 apply to a non-empty one.
HYB-023 counts distinct accounts (AD object GUID, else UPN), so an account
exported twice is not a duplicate.

### Thresholds, units and sources

- **HYB-002 item count.** On-premises value: `ItemCount` (else `TotalItemCount`)
  of the linked on-premises mailbox. EXO value: `ItemCount` of
  `Exchange_EXO_Mailboxes_AllDomains_Stats.csv` matched by UPN (legacy source);
  when that row or value is unknown, `ItemCount` of the linked
  `Exchange_EXO_Mailboxes_AllDomains.csv` row (same Graph statistic: identical on
  all 33,639 shared rows of the 2026-10-25 data). The rule uses the maximum of
  the known values, as the legacy maximum did; with no known value it is not
  assessed (the legacy producer read `Error` as 0).
- **HYB-006 size.** The legacy comparison `$sizeMb -ge 2000` is kept: 2000 MB,
  not 2048. On-premises `TotalItemSize-In-MB` is in MB; EXO `TotalItemSizeGB`
  (mailbox export or statistics, decimal comma accepted) is converted with
  x 1024. The maximum of the known values is compared; none known means not
  assessed. Example: 1.953125 GB = 2000 MB triggers, 1999 MB does not.
- **License groups.** The legacy grouping is unchanged (including its mapping of
  `SPE_F1` to the F3 group); shared and resource mailboxes have no license group.

### Invalid characters

Both sets are unchanged from the legacy producer.

- UPN (HYB-039): space and `#` `%` `&` `~` `+` `/` `=` `?` `*` `{` `}` `<`
  `>` `(` `)` `;` `:` `,` `[` `]` `"` `\`, the pipe and the backtick.
- Primary SMTP local part (HYB-040): `\` `/` `[` `]` `:` `;` `=` `,` `+` `*`
  `?` `<` `>`, the pipe and any whitespace.

### Differences from the legacy producer, by design

- **HYB-051** compares whole proxy entries. The legacy substring match counted a
  `SIP:` entry or a longer address containing the primary SMTP; all 131 legacy
  findings of 2026-10-25 were of that kind.
- **HYB-058** uses the governed population. The legacy hard-coded list counted
  `Generic Account` and `Admin Account` as non-personal (8,860 findings, mostly
  active staff with incomplete AD names); both are `Review required`.

### Rules not carried over

| Legacy | Reason |
| --- | --- |
| 7, 37, 48 | state or convention, not a defect |
| 18, 34 | never triggered |
| 33 | duplicate of HYB-030 |
| 41 | AD-only concern, no Exchange or Entra impact |
| 36 | a verified UPN domain does not need to be an EXO accepted domain |
| 4, 13 | already covered: mailbox and archive size |
| 8, 12, 27, 28, 47, 55, 56, 57 | already covered: License Report |
| 14, 15, 17, 53 | already covered: Mailbox Evidence delegation and inactivity |

## Output 2: `DeviceReadinessGapEvidence.csv`

Written by `New-DeviceInventoryEvidence.ps1` only when `-ReadinessGapOutputPath`
is given (the build gives it). `DeviceInventoryEvidence.csv`,
`DeviceRiskEvidenceSummary.csv`, `DeviceDirectorySummary.csv` and
`DeviceSystemEvidence.csv` are byte-identical with and without it.

Grain: one row per Intune device (`Inventory Scope` = `Intune device`), plus one
row per AD computer that no Intune device matched (`AD computer only`), because
legacy issue 032 (registration pending) concerns mostly such computers. The
output holds **facts, not verdicts**: thresholds and applicability (Windows 10,
eligibility) are evaluated downstream with Device Inventory Evidence.

| Column | Source | Legacy issue |
| --- | --- | --- |
| `TenantKey`, `Inventory Scope`, `Device Source ID`, `Device Name`, `Azure AD Device ID`, `AD Object GUID` | Intune device; AD computer by `IntuneDeviceId`, else by Entra device ID | |
| `Entra Registration Pending` | `M365_Entra_Devices.IsPending` | 032 |
| `Physical Memory (GB)` | `Intune_Devices_Inventory.PhysicalMemoryGB` (decimal comma accepted) | 018 |
| `Firmware Type (source)` | `Intune_Devices_LocalSystem.FirmwareType` | 070 |
| `Secure Boot Status (source)` | `Intune_Devices_LocalSystem.SecureBootStatus` | 070B |
| `Last Reboot (source)`, `Last Reboot Source`, `Last Reboot Time Basis`, `Last Reboot UTC`, `Days Since Last Reboot`, `Reboot Age Basis` | see *Last reboot* | 054, 055 |
| `Evidence Date` | reference UTC date | |

### Last reboot

- **Priority, as in the legacy producer:** the AD computer value
  (`LastRebootDateTime`, `LastRebootDate`, `LastBootUpTime`), then the Intune
  local system value (`Last Reboot Date`, `LastRebootDate`, `LastBootUpTime`)
  only when AD has no value. An invalid AD value does not fall back.
- **Explicit offset** (`Z` or `+hh:mm`): converted to `Last Reboot UTC`; the age
  is the number of UTC calendar days to the reference date.
- **Offset-less value** (`yyyy-MM-dd HH:mm:ss` and the other formats of Device
  System Evidence): never converted to UTC (`Last Reboot UTC` blank). The age is
  the number of calendar days between the source date and the reference UTC
  date, flagged `Source calendar days (offset unknown, +/-1 day)`.
- **Future:** an explicit value after the reference, or an offset-less date more
  than one day after the reference date (a local date can lead UTC by one day),
  gives `Reboot Age Basis` = `Future` and no age.
- **Invalid or missing:** `Last Reboot Time Basis` = `Invalid` or `Missing`, no
  age. Legacy thresholds were 10 days (054, Windows 10) and 30 days (055).

## Validation

| Test | Covers |
| --- | --- |
| `Test-HybridIdentityEvidence.ps1` | every retained rule positive and negative on contoso fixtures; rules 20/23 scope, prefix and case; distinct accounts; 999,999 and 2000 MB boundaries; unknown values; HYB-003/HYB-021 not assessed without accepted domains, without a routing domain and with two candidates, even with configured domains; optional sources not assessed; population mapping, conflict and Review required; header against the contract; determinism |
| `Test-DeviceReadinessGapEvidence.ps1` | reboot basis (offset-less, offset, future, invalid, missing); AD priority and fallback; AD-only scope; full generator with byte-identical existing outputs |
| `Test-CompareReadinessFindings.ps1` | evidence-based justification of rules 51/58; an unjustified new rule 51 finding reported as a regression; undetermined 018/055 without usable inventory, 054 without inventory, 070 without firmware; reference-date justification only with both dates, both ages recomputed from the same source value and the declared age matching (declared 29 for a computed 30 stays unexplained; identical or missing dates stay unexplained); read-only inputs |
| `Test-PreparedBuildScheduler.ps1`, `Test-PreparedClassificationCapture.ps1` | new task order, dependency, checkpoint `BeforeHybridIdentity`, changed rules, `-Only HybridIdentity`, allowed domains |
| `Test-PreparedEndpointTransfer.ps1` | batches prepared before this change (neither file) stay transferable; a batch with only one of the two files is rejected |

`scripts/Compare-ReadinessFindings.ps1` compares the legacy CSVs with the new
outputs. A difference is **justified only by evidence the comparator checks**;
everything else is reported as unexplained, including every new-only finding.

- Hybrid, per retained legacy number and account (both, legacy only, new only):
  - HYB-051 legacy only is justified when the AD proxyAddresses
    (`-AdUsersPath`) hold at most one whole entry equal to the primary SMTP while
    the legacy substring count is higher.
  - HYB-058 legacy only is justified when the AD `AccountType` was in the legacy
    list and the governed configuration (`-AccountClassificationConfigPath`) does
    not classify it as `Non-human`.
  - Without these inputs, or for any other rule, the difference stays unexplained.
- Devices: the legacy condition of each issue (032, 018, 054, 055, 070, 070B) is
  re-applied to the facts only when every input of its applicability is known:
  eligibility (and Windows generation for 054, 070, 070B) from a usable Device
  Inventory Evidence row, memory for 018, reboot age for 054/055, firmware and
  Secure Boot status for 070, Secure Boot status for 070B. Otherwise the device is
  *undetermined*. Legacy-only devices are *not in prepared device scope*,
  *undetermined*, *justified* or *unexplained*. A reboot-age difference is
  attributed to the reference dates only when all of these hold; otherwise it
  stays unexplained, and being near the threshold is never a justification:
  - both reference dates are known (`-LegacyReferenceDate`, the legacy run date,
    and the fact's `Evidence Date`) and they differ;
  - both ages are recomputed from the same `Last Reboot (source)` value;
  - the recomputed age at the evidence date equals `Days Since Last Reboot`;
  - the age at the legacy date reproduces the legacy verdict and the age at the
    evidence date gives the new one.

Local comparison on the 2026-10-25 data, with AD, classification and legacy run
date evidence: 23 retained hybrid rules identical account by account; HYB-051
(131 legacy-only) and HYB-058 (8,850 legacy-only) fully justified by evidence;
no new-only finding and no unexplained difference. The six device conditions
re-applied to the facts match the legacy verdicts device by device.

## Later (out of scope)

- Server run of the preparation and comparison report review.
- Private Power BI: replace the `Exchange_HybridIdentity_Issues.csv` read in
  `Mailbox Migration Inventory` with `HybridIdentityFindings.csv`; use the device
  facts.
- Retire both legacy producers, their outputs, archives and weekly history, and
  remove them from source selection, schemas, receipts and Data Trust.
