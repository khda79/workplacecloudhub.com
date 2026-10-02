# SmartInventory to SmartWorkplaceCMDB source contract

Status: migration candidate, not a qualified replacement for the CMDB collection pipeline.
Contract version: 0.5.0.
Executable preparation contract: 0.3.1 (`cmdb-prepared-contract.json.txt`).

## Ownership and boundaries

SmartInventory owns collection, collection fidelity, and preparation of the final
CMDB reporting CSVs. SmartWorkplaceCMDB owns its semantic model and report.
Preparation reads canonical SmartInventory `DATA-LAST`, not Intelligence's
`DATA-POWERBI`. Its final current-only output is `DATA-POWERBI-CMDB`, beside those
directories under the configured private data root. On this workstation that is
`C:\SmartM365\DATA\DATA-POWERBI-CMDB`.

There is no new CMDB history, persistent compatibility Raw tree, customer-specific
path in executable code, or copy of the full source directory. Transient staging
is permitted for validation and local fail-preserving replacement. Preparation must preserve the
last valid output on failure and must not overwrite unrelated directories.
Intelligence's existing history is outside this migration's cleanup scope.

Runtime manifests and configurations use `.json.txt`. Power BI's native JSON
files remain the documented technical exception. Customer CSVs, manifests,
logs, audits, models containing data, and local configurations remain private.

## Required native identity and grain

| Population | Source | Grain and identity | Rules |
| --- | --- | --- | --- |
| Entra users | `M365_Users_Active.csv` | Tenant + `Object Id` | The filename does not imply enabled-only. Retain disabled accounts and guests. UPN is an attribute, not identity. |
| AD users | `AD_Users_AllDomains.csv` | Tenant + `ObjectGUID` | Include users without UPN. Preserve SID, immutable ID, domain, primary group, and classification separately. Do not infer human status from enabled status. |
| AD workplace computers | `AD_Computers_AllDomains.csv` | Tenant + `ObjectGUID` | Keep explicit workplace scope. Do not claim this Windows-workstation export is all AD computer objects. |
| AD domains | `AD_Domains_AllDomains.csv` | Tenant + `DNSRoot` | Preserve domain SID, NetBIOS name, forest, DN, mode and directory roles. A failed domain query blocks the full run. |
| AD groups | `AD_Groups_AllDomains.csv` | Tenant + `ObjectGUID` | Retain direct member DNs; resolve identity from native keys. Unresolved/external members stay evidence gaps, not silent drops. |
| AD object identities | `AD_DirectoryObjects_AllDomains.csv` | Tenant + `ObjectGUID` | Minimal native GUID/SID/DN/class evidence includes no-UPN users, servers, contacts and foreign security principals. This is not a full server hardware inventory. |
| AD memberships | `AD_GroupMemberships_AllDomains.csv` | Tenant + group GUID + member DN + membership kind | Resolve direct members from MembersJson, which preserves semicolons in DNs; keep unresolved/external evidence. PrimaryGroup membership is separate from direct member-list evidence. |
| AD OUs | `AD_OUs_AllDomains.csv` | Tenant + `ObjectGUID` | Preserve DN and creation/change evidence. Names are not unique keys. |
| Entra groups | `M365_EntraGroups_All.csv` | Tenant + `GroupId` | All groups, not just licensing groups. Retain security/mail flags, types, dynamic rules and native synchronization evidence. |
| Entra memberships | `M365_EntraGroupMemberships_All.csv` and `M365_EntraGroupMembershipScope.csv` | Tenant + group ID + member ID; scope has one row per group | Direct membership, including non-user directory objects. Zero-member groups retain scope rows. Hidden membership needs explicit access. No inferred transitive membership. |
| Intune policies | `Intune_Policies_All.csv` | Tenant + policy family + policy ID | Settings Catalog, configuration, compliance, feature and quality update families. Preserve list-response native metadata; this is not a complete export of every nested setting or effective device assignment. |
| Intune assignment targets | `Intune_PolicyAssignments_All.csv` | Tenant + policy family + policy ID + assignment ID | Preserve include/exclude target types, assignment filters and native payload. A target group does not prove effective policy application to a device. |
| Entra devices | `M365_EntraDevices_All.csv` | Tenant + `ObjectId` | All platforms, independent of diagnostic OS/trust filters. `DeviceId` is the AD/Intune bridge, not the Entra object ID. |
| Intune devices | `Intune_ManagedDevices_All.csv` | Tenant + `ManagedDeviceId` | All platforms. Preserve ownership, agent, enrollment type, user ID, Entra device ID and dates. |
| Detailed hardware | `Intune_DeviceHardware_All.csv` | Tenant + `ManagedDeviceId` | Explicit per-device GET/select, all managed platforms; acquisition success, actual date, missing/reported-zero byte states and exact inventory correlation. Failed acquisition cannot qualify CMDB full scope. |
| Native SKUs | `M365_Licenses_Tenant.csv` | Tenant + `Id` (SKU GUID) | Keep enabled/warning/suspended/locked capacity, consumed units, capability and applies-to. Missing capacity is null, not zero. |
| Assignment paths | `M365_Licenses_AssignmentPaths.csv` | Tenant + `UserId` + `SkuId` + `AssignedByGroupId` | Empty group ID means direct. Distinguish Active, ActiveWithError, Disabled, Error and unknown evidence. Do not require effective license details to retain a path. |
| Service-plan catalog | `M365_Licenses_ServicePlans_Catalog.csv` | Tenant + `SkuId` + `PlanId` | Include tenant plans even if no user has the SKU. Tenant provisioning status is not an individual user's provisioning state. |
| User service-plan states | `M365_Licenses_UserServicePlanStates.csv` | Tenant + `UserId` + `SkuId` + `PlanId` | Decode compact state codes explicitly. Do not substitute the first observed user's plan status into the catalog. |
| Applications | `Intune_DiscoveredApps_Summary.csv` | Tenant + `AppId` | Native application/version inventory. Product identity excludes version but includes normalized name, publisher and platform. |
| Application installations | `Intune_DiscoveredApps_AppDeviceRelations.csv` | Tenant + `AppId` + `DeviceId` | Count distinct product/device pairs for footprint. Summed per-version DeviceCount is not distinct installed devices. |
| Upgrade readiness | `Intune_Devices_UpgradeEligibility.csv` | Tenant + `GraphId` | Different IDs with the same name remain distinct. Reject conflicting repeated IDs; identical repeats may be collapsed. Preserve capable/notCapable/upgraded/unknown. |
| Endpoint Analytics | `Intune_EndpointAnalytics_DevicePerformance.csv` | Tenant + `ReportName` + `DeviceId` | Reports have different fields; use score-bearing rows only for 0-100 score metrics. Blank is not zero. Repeated source keys block preparation, including conflicting scores; no row selection or averaging. Different report names for one device remain separate evidence. |
| Windows update state | `Intune_WindowsUpdate_Status.csv` | Tenant + `PolicyId` + `DeviceId` | Include successful/in-progress states, not only alerts. Maintain separate policy/device state and alert facts. |
| Windows update alerts | `Intune_WindowsAutopatch_Alerts_Detail.csv` | Tenant + source report + device + policy + event date + alert name | An empty valid alert set is allowed; it does not prove complete update coverage. |
| Mailboxes | EXO, local and remote mailbox exports | Native mailbox/object GUID per source; reconcile nonblank normalized SMTP secondarily | Online/Remote evidence takes precedence over duplicate local evidence. Retain and flag missing SMTP, conflicting identity and unavailable hosting evidence. |
| SharePoint sites | `M365_SPO_Sites.csv` | Tenant + SiteId, qualified by SiteIdentitySource | Prefer Graph composite site identity, retain report ID separately. Usage-report coverage is not all-tenant site discovery. Collection quality is separate from business Status. |
| Teams | `M365_Teams_Teams.csv` | Tenant + `TeamId` | Member collection availability must distinguish empty teams from missing/failed child collection. |
| Team memberships | `M365_Teams_Members.csv` | Tenant + `TeamId` + `UserId` + `Role` | Preserve guests and orphan-user evidence. Do not multiply team counts through child joins. |
| User activity | `M365_Users_Activity.csv` and user sign-in evidence | Tenant + observed native user identity / qualified UPN match | Preserve workload/report date. Missing sign-in history does not prove a user never connected or wasted a license. |

## Current source-export additions

The candidate adds full-device, full-group, domain, license-path and capacity
evidence, plus successful-sign-in evidence. Legacy diagnostic device exports
remain Windows-filtered for their existing consumers; those outputs are not the
source of full CMDB fleet counts.

New source exports use strict immutable-key completeness and uniqueness gates.
The legacy Windows inventory also retains every managed-device ID instead of
discarding same-name devices. This corrects a coverage loss: Intelligence's
device counts and joins require requalification before deploying the candidate.
A per-user license-resolution failure blocks licensing CSV publication rather
than publishing a misleading complete run. Identical readiness IDs may be
collapsed; duplicate names cannot determine which physical device is kept.

`CollectedAtUtc` in the new SDK exports is captured after the response is read.
It is acquisition evidence, not user activity time or Graph report refresh time.
It does not prove all families were collected atomically. Empty successful
datasets still require an explicit schema and producer completion evidence.

Autopilot exports preserve native IDs rather than deduplicating serial numbers.
Identical repeats of an ID can be collapsed; conflicting records block publication.
Blank serials do not erase native identity. Current counts may therefore increase
for existing consumers and must be requalified, not corrected back to old totals.

Successful zero acquisitions now retain complete headers for Autopilot, verified
domains, Entra users, native AD domain/combined exports and AD user/computer
enrichment, and remote mailboxes. Remote query errors or acquisition warnings
cannot be presented as a qualified zero. This is not a claim that every producer's
zero-population path has been qualified on the tenant.

The executable contract also requires AD membership identity/type, Entra
membership type/status, Teams member-count, and update-alert provenance/key
headers. Missing headers block preparation; blank legitimately unavailable values
remain separate from missing schema, zero population and successful collection.

The WorkplaceScope collector remains a non-scheduled candidate. Its four exports
are validated after all required Graph families succeed. The shared Core helper
now integrates current completion receipts into 17 producers for the 33 required
CMDB source CSVs. The public registry declares exact producer/file/scope ownership.
Every receipt records full tenant identity, native scope, run/version, acquisition
interval and actual logical CSV row counts/SHA-256 hashes. Running, failed,
restricted, reused and unavailable acquisitions cannot claim completed full scope.
There is no CMDB receipt history or global all-collector transaction. Preparation
assembles the 17 validated current receipts; no extra raw aggregate is required.

Applications now cover all Intune platforms. Complete current product/device
footprint requires All-mode relations and per-app count agreement; Top/None
evidence cannot supply this metric. All-mode does not reuse the previous-run
cache. A resumed run remains labelled ReusedEvidence; preparation must qualify
its acquisition interval rather than relabel it as a fresh collection. MaxApps
and MaxItems outputs cannot replace canonical complete filenames.

Intelligence's current product grain includes publisher, platform and tenant.
Legacy weekly observations retain their name-only, Windows-scoped definition
and numerical values; explicit definition/scope columns identify the boundary.
Per-version installation observations remain separate from distinct product
device counts. Its licence capacity remains null when source evidence is absent.

Mailbox exports now carry acquisition/scope and native-identity availability.
The legacy EXO ImmutableId field must not be used as an AD immutable-ID bridge;
use actual native identity and qualified OnPremisesImmutableId evidence instead.
Preparation consumes the exact existing producer headers: EXO `MailboxGuid`,
local `ObjectGUID` / `RecipientType` / `PrimarySMTPaddress`, and remote
`ObjectGuid` / `RecipientTypeDetails` / `PrimarySmtpAddress`. These headers are
not renamed in the collectors or in Intelligence's sources. Mailboxes without
SMTP retain their declared per-source native GUID; their addresses are not invented.
The normalized user-activity export uses `UserPrincipalName`, `ReportRefreshDate`,
`IsDeleted`, `AssignedProducts` and the six camel-case workload activity-date
columns. Preparation requires these headers, preserves report refresh separately
from acquisition, and includes all six workloads and tied latest dates in its
aggregate activity evidence. A missing workload column is not zero activity.
SPO and Teams acquisition/report dates and child-collection availability are
explicit. Unavailable enrichment is not a zero or a healthy tenant state.

## Final reporting contract and semantics

Preparation produces one flat, validated root for 46 reporting tables: the
existing 39 names (hardware gains additive fields), plus six native AD detail
tables and EntraGroupMembership for 360 evidence. Source lineage, producer outcome,
tenant identity, file hashes, source acquisition interval, source/report date
qualification and schema version belong in the current `.json.txt` manifest.
No partial source can be promoted merely because an older CSV exists.

Metric requirements:

- Intune coverage uses native matched device IDs, not an Entra `IsManaged` flag.
- AD Windows 7-11 workstation coverage has its own explicit denominator. Unknown
  AD ownership/country cannot be silently classified as Corporate/a country.
- Ownership and country filtering retain unknown evidence. User, device, mailbox
  and site countries have separate provenance.
- Licensing capacity is tenant-wide; selected-country assignment counts are
  user-linked. Distinguish F1, F3, E3, E5 and Copilot.
- License opportunity means reviewable evidence, not proof of non-use from a
  missing sign-in date. Use available workload/successful-sign-in evidence and
  state the observation window.
- Endpoint Analytics is a score on a 0-100 scale, not a proportion.
- Application product footprint counts distinct managed-device IDs across versions.
- Mailbox hosting and mailbox recipient type are separate measures.

## Promotion gates still required

1. Qualify the integrated producer freshness/completion verification on the
   collection host. Static, synthetic and PowerShell interoperability tests are
   not proof that all real native families, privileges and empty cases qualify.
   Do not trust old CSV timestamps or fabricate receipts for previous exports.
2. Qualify candidate SDK commands, permissions, production volumes and native
   export semantics on the collection host. Local synthetic Intelligence and
   source regression checks pass; real historical/source equivalence is pending.
3. Qualify the candidate current-only preparation job on complete native producer
   proof. The local engine now prepares 46 flat tables from 33 raw sources,
   validates the current manifest and uses fail-preserving local replacement.
   See [CMDB preparation](CMDB-PREPARATION.md) for execution boundaries, receipt
   requirements, semantic differences and remaining equivalence checks.
4. Adapt the CMDB semantic model to the single final output root. Verify all ten
   pages, slicers, joins and measures in Desktop without tenant publication.
5. Run one consolidated real collection/qualification campaign on the collection
   host. Compare source/result populations, joins, coverage and quality findings.
6. Only after acceptance, retire the CMDB collectors and obsolete configuration;
   perform approved Git publication and private-data cleanup separately.

Offline tests and successful parsing are not real tenant qualification. Do not
switch the production report or remove the working collectors before these gates.

## Native API references

- [Microsoft Graph license assignment states](https://learn.microsoft.com/en-us/graph/api/resources/licenseassignmentstate?view=graph-rest-1.0)
- [Microsoft Graph subscribed SKUs](https://learn.microsoft.com/en-us/graph/api/resources/subscribedsku?view=graph-rest-1.0)
- [Microsoft Graph device identity and synchronization properties](https://learn.microsoft.com/en-us/graph/api/resources/device?view=graph-rest-1.0)
- [Microsoft Graph direct group membership and hidden-member access](https://learn.microsoft.com/en-us/graph/api/group-list-members?view=graph-rest-1.0)
- [Microsoft Graph Settings Catalog policies](https://learn.microsoft.com/en-us/graph/api/intune-deviceconfigv2-devicemanagementconfigurationpolicy-list?view=graph-rest-beta)
