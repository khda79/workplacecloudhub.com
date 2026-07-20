# SmartWorkplaceCMDB

SmartWorkplaceCMDB is an autonomous Workplace configuration management database project for Microsoft workplace environments.

The project collects workplace inventory data, normalizes it into CMDB entities, prepares Power BI-ready tables, and produces local reports without requiring another repository project as a runtime dependency.

## Initial Scope

The first scope is Microsoft workplace inventory across cloud and Active Directory:

- Entra ID users, groups, and devices.
- Intune managed devices and compliance signals.
- Microsoft 365 license SKUs and assignments.
- Exchange Online mailboxes.
- User-to-device and source-to-entity relationships.
- Data quality, source freshness, and confidence scoring.

Active Directory users, groups, computers, domains, and direct group
memberships are included through the native read-only collector. Azure Virtual
Desktop, Citrix, local endpoint inventory, and external data sources remain
planned extensions.

## Current Status

The repository currently provides autonomous configuration, tenant identity,
CSV contracts, schema initialization, contract validation, a local report, and
Power BI-ready table definitions. Native pipelines collect and normalize
Microsoft Entra users, groups, and devices, enrich devices from Intune,
inventory Microsoft 365 licensing, correlate Exchange Online mailboxes to
Entra users, and inventory Active Directory from the same collection host.

## Design Principles

- Keep the project autonomous: own configuration, collectors, schemas, outputs, reports, and Power BI assets.
- Keep tenant values local: use `*.local.json` files and never commit tenant IDs, app IDs, secrets, certificates, logs, exports, or production identifiers.
- Keep raw collection data separate from normalized CMDB tables and Power BI-ready tables.
- Prefer read-only collection by default.
- Treat Power BI as a consumer of curated CMDB tables, not as the primary transformation engine.
- Preserve compatible CSV outputs by default and require explicit approval before schema-only reinitialization.

## Repository Layout

```text
SmartWorkplaceCMDB/
  Config/       Local configuration templates.
  Collectors/   Source-specific collectors.
  Orchestration/ Safe end-to-end pipeline coordination.
  Launchers/    Centralized PowerShell 7 source and full-pipeline launchers.
  Build/        CMDB contract initialization and validation.
  Reports/      Local HTML report generation.
  Modules/      Shared PowerShell module.
  PowerBI/      Power BI model guidance, queries, measures, and templates.
  Schema/       CMDB schema and CSV table contracts.
  Tests/        Autonomous contract and safety tests.
  Data/         Runtime output, ignored by Git.
```

## Identity Contract

`ProfileKey`, `OrganizationKey`, and `EnvironmentKey` accept lowercase letters, digits, and internal hyphens only. Each component is limited to 64 characters.

`TenantKey` is always derived as:

```text
<OrganizationKey>-<EnvironmentKey>
```

`TenantId` is optional, but when supplied it must be an Entra tenant GUID.

## Configuration

At runtime, the project loads and merges:

1. `Config/SmartWorkplaceCMDB.global.local.json`
2. `Config/Tenants/<ProfileKey>.local.json`
3. Explicit command-line parameters

Missing runtime JSON files are created from their templates during a normal build or report run. Missing template properties are added without overwriting existing local values. `-ValidateOnly` and `-NoConfigWrite` perform read-only configuration resolution.

## Output Layout

Runtime output is profile-scoped. `ProfileKey` selects the local folder, while every tenant-scoped CSV starts with `TenantKey`, `OrganizationKey`, `EnvironmentKey`, and `TenantId`.

```text
SmartWorkplaceCMDB/Data/Tenants/<ProfileKey>/
  DATA-ALL/
    Entra/Users/<yyyy>/<MM>/Entra_Users_<timestamp>.csv
    Entra/Groups/<yyyy>/<MM>/Entra_Groups_<timestamp>.csv
    Entra/Devices/<yyyy>/<MM>/Entra_Devices_<timestamp>.csv
    Intune/ManagedDevices/<yyyy>/<MM>/Intune_ManagedDevices_<timestamp>.csv
    M365/SubscribedSkus/<yyyy>/<MM>/M365_SubscribedSkus_<timestamp>.csv
    M365/UserLicenseAssignments/<yyyy>/<MM>/M365_UserLicenseAssignments_<timestamp>.csv
    ExchangeOnline/Mailboxes/<yyyy>/<MM>/ExchangeOnline_Mailboxes_<timestamp>.csv
    ActiveDirectory/<entity>/<yyyy>/<MM>/ActiveDirectory_<entity>_<timestamp>.csv
  DATA-LAST/
    Raw/Entra/Entra_Users.csv
    Raw/Entra/Entra_Groups.csv
    Raw/Entra/Entra_Devices.csv
    Raw/Intune/Intune_ManagedDevices.csv
    Raw/M365/M365_SubscribedSkus.csv
    Raw/M365/M365_UserLicenseAssignments.csv
    Raw/ExchangeOnline/ExchangeOnline_Mailboxes.csv
    Raw/ActiveDirectory/ActiveDirectory_Domains.csv
    Raw/ActiveDirectory/ActiveDirectory_Users.csv
    Raw/ActiveDirectory/ActiveDirectory_Groups.csv
    Raw/ActiveDirectory/ActiveDirectory_Computers.csv
    Raw/ActiveDirectory/ActiveDirectory_GroupMemberships.csv
    CMDB/ActiveDirectory/CMDB_ActiveDirectoryDomains.csv
    CMDB/ActiveDirectory/CMDB_ActiveDirectoryUsers.csv
    CMDB/ActiveDirectory/CMDB_ActiveDirectoryGroups.csv
    CMDB/ActiveDirectory/CMDB_ActiveDirectoryComputers.csv
    CMDB/ActiveDirectory/CMDB_ActiveDirectoryGroupMemberships.csv
    CMDB/CMDB_Users.csv
    CMDB/CMDB_Groups.csv
    CMDB/CMDB_Devices.csv
    CMDB/CMDB_Licenses.csv
    CMDB/CMDB_Mailboxes.csv
    CMDB/CMDB_UserDeviceRelationships.csv
    CMDB/CMDB_Relationships.csv
    CMDB/CMDB_DataQuality.csv
    CMDB/CMDB_BuildManifest.csv
    PowerBI/DimTenant.csv
    PowerBI/DimDate.csv
    PowerBI/DimUser.csv
    PowerBI/DimGroup.csv
    PowerBI/DimDevice.csv
    PowerBI/DimLicenseSku.csv
    PowerBI/FactDeviceCompliance.csv
    PowerBI/FactUserDeviceRelationship.csv
    PowerBI/FactUserLicense.csv
    PowerBI/FactMailbox.csv
    PowerBI/FactDataQuality.csv
  LOG-ALL/
```

Power BI consumes curated tables from:

```text
SmartWorkplaceCMDB/Data/Tenants/<ProfileKey>/DATA-LAST/PowerBI/
```

The canonical CSV definitions are stored in:

```text
SmartWorkplaceCMDB/Schema/SmartWorkplaceCMDB.tables.json
```

## Build Safety

The build creates missing schema-only CSV files and preserves existing compatible files.

If an existing CSV has an incompatible header, the build stops without replacing it. `-ForceInitialize` explicitly replaces schema tables with empty files and must only be used after the existing data has been reviewed or backed up.

Validate configuration and identity without writing:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Build\SmartWorkplaceCMDB-Build.ps1 -Tenant prod -OrganizationKey example -EnvironmentKey prod -TenantKey example-prod -ValidateOnly
```

Initialize missing contract files:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Build\SmartWorkplaceCMDB-Build.ps1 -Tenant prod -OrganizationKey example -EnvironmentKey prod -TenantKey example-prod
```

## Entra Users Collector

The native Entra users collector uses read-only Microsoft Graph app-only
certificate authentication. It requires the `Microsoft.Graph.Authentication`
PowerShell module and the `User.Read.All` application permission with
administrator consent.

Tenant-specific `TenantId`, `ClientId`, and `CertificateThumbprint` values belong
only in `Config/Tenants/<ProfileKey>.local.json`.

Validate live-collection prerequisites without connecting or writing output:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\Entra\SmartWorkplaceCMDB-EntraUsers-Collect.ps1 -Tenant prod -ValidateOnly
```

Run a bounded collection:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\Entra\SmartWorkplaceCMDB-EntraUsers-Collect.ps1 -Tenant prod -MaxItems 10
```

Normalize the latest raw snapshot:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\Entra\SmartWorkplaceCMDB-EntraUsers-Normalize.ps1 -Tenant prod
```

The initial collector intentionally leaves manager and sign-in activity fields
empty. These fields require separate Graph queries or additional permissions and
will be introduced only through a reviewed contract update.

## Entra Groups Collector

The Entra groups collector uses the same app-only certificate configuration and
requires the `Group.Read.All` application permission with administrator consent.
It collects base group identity, type, mail, and security properties.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\Entra\SmartWorkplaceCMDB-EntraGroups-Collect.ps1 -Tenant prod -MaxItems 10
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\Entra\SmartWorkplaceCMDB-EntraGroups-Normalize.ps1 -Tenant prod
```

Member and owner counts remain empty until the dedicated group relationship
collector is implemented.

## Entra Devices Collector

The Entra devices collector requires the `Device.Read.All` application
permission. It uses the Entra `deviceId` as the stable correlation key for later
Intune enrichment.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\Entra\SmartWorkplaceCMDB-EntraDevices-Collect.ps1 -Tenant prod -MaxItems 10
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\Entra\SmartWorkplaceCMDB-EntraDevices-Normalize.ps1 -Tenant prod
```

Without Intune enrichment, ownership, primary user, and last synchronization
remain empty.

## Intune Managed Devices Collector

The autonomous Intune collector requires the
`DeviceManagementManagedDevices.Read.All` Microsoft Graph application
permission with administrator consent and an active Intune tenant license.
It joins `managedDevices.azureADDeviceId` to the Entra `deviceId`.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\Intune\SmartWorkplaceCMDB-IntuneManagedDevices-Collect.ps1 -Tenant prod -MaxItems 10
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\Intune\SmartWorkplaceCMDB-IntuneDevices-Normalize.ps1 -Tenant prod
```

The enrichment retains Entra-only and Intune-only devices. When an Intune
record has no usable `azureADDeviceId`, its stable fallback source key is
`intune:<managedDeviceId>`. If multiple Intune records share one Entra device
ID, the record with the newest Intune synchronization time is selected. The
normalizer also publishes one `FactDeviceCompliance.csv` row per CMDB device
using the same stable device key. Intune dates are normalized to invariant UTC
ISO text before raw export.

## Primary User-Device Relationships

The relationship normalizer requires no tenant connection. It correlates
`CMDB_Devices.PrimaryUserId` from Intune with `CMDB_Users.SourceUserId` and
publishes `CMDB_UserDeviceRelationships.csv` plus
`FactUserDeviceRelationship.csv`.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\Intune\SmartWorkplaceCMDB-IntuneUserDeviceRelationships-Normalize.ps1 -Tenant prod
```

Only resolvable primary users are published. Missing user references are
counted as orphans and will feed the dedicated data-quality pipeline.

## General Relationship Consolidation

The autonomous consolidator requires no tenant connection. It combines validated
primary user-device, user-mailbox, and user-license links into
`CMDB_Relationships.csv`, while enforcing entity reference integrity and canonical
UTC collection dates.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\SmartWorkplaceCMDB-Relationships-Normalize.ps1 -Tenant prod
```

## Microsoft 365 Subscribed SKUs Collector

The subscribed SKUs collector uses the least-privileged
`LicenseAssignment.Read.All` Microsoft Graph application permission. It reads
commercial subscription capacity and publishes `CMDB_Licenses.csv` and
`DimLicenseSku.csv`. User-to-SKU assignments are a separate collector.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\M365\SmartWorkplaceCMDB-M365SubscribedSkus-Collect.ps1 -Tenant prod -MaxItems 10
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\M365\SmartWorkplaceCMDB-M365SubscribedSkus-Normalize.ps1 -Tenant prod
```

If multiple commercial subscriptions share a `skuId`, their unit counts are
aggregated into one stable CMDB license entity.

## Microsoft 365 User License Assignments Collector

The user license assignment collector requires the `User.Read.All` Microsoft
Graph application permission and reads `licenseAssignmentStates` for each
user. `-MaxItems` limits users, while the raw CSV retains every direct and
group-derived assignment state.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\M365\SmartWorkplaceCMDB-M365UserLicenseAssignments-Collect.ps1 -Tenant prod -MaxItems 10
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\M365\SmartWorkplaceCMDB-M365UserLicenseAssignments-Normalize.ps1 -Tenant prod
```

`FactUserLicense.csv` contains one row per user and SKU. When multiple states
exist, the normalizer selects the most critical state and then the newest. Its
`AssignedDateTime` field stores Graph `lastUpdatedDateTime`; Graph does not
provide the original license assignment creation time in this resource.

## Exchange Online Mailboxes Collector

The mailbox collector uses certificate-based app-only authentication with the
CMDB-owned application ID and certificate configured under `MicrosoftGraph`.
`ExchangeOnline.Organization` must contain the tenant's primary
`.onmicrosoft.com` domain. The application requires `Exchange.ManageAsApp`
administrator consent and read-only Exchange recipient RBAC.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\ExchangeOnline\SmartWorkplaceCMDB-ExchangeOnlineMailboxes-Collect.ps1 -Tenant prod -ValidateOnly
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\ExchangeOnline\SmartWorkplaceCMDB-ExchangeOnlineMailboxes-Collect.ps1 -Tenant prod -MaxItems 10
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\ExchangeOnline\SmartWorkplaceCMDB-ExchangeOnlineMailboxes-Normalize.ps1 -Tenant prod
```

The first mailbox contract is intentionally limited to base mailbox identity,
recipient type, primary SMTP address, plan, and archive status. The normalizer
publishes `CMDB_Mailboxes.csv` and `FactMailbox.csv`; it links a mailbox to an
existing CMDB user through `ExternalDirectoryObjectId` when a matching Entra
user is available. When Exchange does not expose an external directory object
ID, the stable mailbox key falls back to ExchangeGuid. Mailbox statistics are
reserved for a separate collector.

## Active Directory Collector

The Active Directory pipeline is autonomous and read-only. It runs on the same
Windows collection host as Entra, Intune, Microsoft 365, and Exchange Online.
That host must run PowerShell 7, reach Microsoft cloud services and a domain
controller, provide the RSAT `ActiveDirectory` module, and use an account with
directory read access. No password or credential is stored by
SmartWorkplaceCMDB.

The tenant-local configuration enables forest-wide discovery by default and
can define an optional preferred domain controller:

```json
"ActiveDirectory": {
  "Enabled": true,
  "ForestWide": true,
  "Server": "",
  "SearchBase": "",
  "IncludeGroupMemberships": true
}
```

The single collection host uses one tenant profile for every source. Its
`Output.DataRootPath` points to the shared CMDB data location.

Validate the domain host without writing output:

```powershell
.\SmartWorkplaceCMDB\Launchers\ActiveDirectory\Start-SmartWorkplaceCMDB-ActiveDirectory-Validate.cmd
```

Run the explicit collection and normalization:

```powershell
.\SmartWorkplaceCMDB\Launchers\ActiveDirectory\Start-SmartWorkplaceCMDB-ActiveDirectory-Collect.cmd
```

The collector uses `Get-ADForest` to enumerate every domain in the forest, then
selects an AD Web Services-capable domain controller for each domain and
publishes domain metadata, users, groups, computers, and direct group
memberships. `-Server` can set the preferred controller used to discover the
forest. A non-empty `-SearchBase` is supported only with `ForestWide` set to
`false`, because one distinguished name cannot safely represent every forest
domain. `-MaxItems 10` creates an isolated bounded test run when invoked through
the orchestrator. Direct group relationships are preserved without recursively
flattening nested groups.

The normalizer requires only the raw CSVs, not domain connectivity. It publishes
five source-specific tables under `DATA-LAST\CMDB\ActiveDirectory`. Cross-source
reconciliation with the canonical Entra and Intune entity tables is a separate
curation stage, so uncertain hybrid matches are never created silently.

## SharePoint Publication

SmartWorkplaceCMDB publishes CSV files autonomously through Microsoft Graph after
a successful unbounded live orchestration. It does not import or require the
SmartM365 runtime. Only files created or changed by the current run are uploaded,
and their relative `DATA-ALL`, `DATA-LAST`, and `LOG-ALL` structure is preserved.

The production target is configured in the tenant-local file:

```json
"SharePoint": {
  "Enabled": true,
  "SiteHostname": "<tenant>.sharepoint.com",
  "SitePath": "/sites/SMART-M365",
  "LibraryDisplayName": "Documents",
  "TargetFolderPath": "SMART-CMDB/DATA"
}
```

The publisher reuses `MicrosoftGraph.TenantId`, `ClientId`, and
`CertificateThumbprint`. The application needs write access to the configured
SharePoint site, preferably through `Sites.Selected` plus a write grant on that
site. Parent folders are created when missing, files up to 250 MB use direct
upload, and larger CSV files use a resumable upload session.

Publication runs automatically after `-Collect`. It is skipped for validation,
offline fixtures, failed pipelines, and every `-MaxItems` run. Use
`-DisableSharePointUpload` for an intentional live run without publication.
Upload failures preserve the collected files and return
`CompletedWithWarnings` with upload and failure counts.

After OneDrive synchronizes the SharePoint library on the analysis workstation,
the published copy is available under:

```text
C:\Users\<user>\EMEIS\SMART-M365 - Documents\SMART-CMDB\DATA
```

## Data Quality Normalization

The autonomous data-quality normalizer uses only curated local CSVs. It publishes
orphan references, unlinked mailboxes, missing or invalid collection dates,
duplicate entity keys, and freshness findings to `CMDB_DataQuality.csv` and
`FactDataQuality.csv`.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\SmartWorkplaceCMDB-DataQuality-Normalize.ps1 -Tenant prod
```

## Tenant and Date Dimensions

The autonomous dimension normalizer publishes one tenant row and a complete date
calendar covering the full years observed in curated date columns. Technical
sentinel dates before 1900 or more than ten years in the future are excluded.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Collectors\SmartWorkplaceCMDB-Dimensions-Normalize.ps1 -Tenant prod
```

## Autonomous Orchestration

The orchestrator coordinates the cloud source pipelines, the dedicated Active
Directory pipeline, and all local curation stages. Its default behavior is a
read-only prerequisite validation; live tenant
collection requires the explicit `-Collect` switch.

The default `Full` pipeline includes Entra, Active Directory, Intune, Microsoft
365 licensing, Exchange Online, and all local curation stages. The dedicated
`-Pipeline ActiveDirectory` option remains available for an isolated AD-only
validation or collection.

Read-only full validation:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Orchestration\SmartWorkplaceCMDB-Orchestrator.ps1 -Tenant prod -ValidateOnly
```

Explicit full collection and curation:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Orchestration\SmartWorkplaceCMDB-Orchestrator.ps1 -Tenant prod -Collect
```

Bounded live runs are accepted only for an individual source pipeline. When no
custom data root is supplied, their outputs are automatically isolated under
`Data/TestRuns/<Tenant>_MAXITEMS-<N>_<timestamp>`.

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Orchestration\SmartWorkplaceCMDB-Orchestrator.ps1 -Tenant prod -Collect -Pipeline EntraUsers -MaxItems 10
```

Centralized launchers resolve PowerShell 7 and forward extra command-line
arguments. `Launchers/Cloud` contains the cloud-source launchers plus the full
validation and collection launchers. `Launchers/ActiveDirectory` contains the
AD-only launchers. All of them are intended for the same collection host.

### Scheduled Orchestrator Installation

Install or update the unattended daily task through the guided launcher:

```powershell
.\SmartWorkplaceCMDB\Launchers\Orchestrator\Start-SmartWorkplaceCMDB-OrchestratorScheduledTask-Installer.cmd
```

The installer is preview-only until `-Execute` is confirmed. It registers
`\WCH\SmartWorkplaceCMDB Orchestrator - <Tenant>` under a dedicated Windows
account and runs PowerShell 7 with `-Tenant <Tenant> -Collect`. The default
schedule is daily at 02:00. `SYSTEM` and `LocalSystem` are refused because
the task identity needs its CurrentUser application certificate plus AD, UNC,
and SharePoint access. The password is requested only through `Get-Credential`.

See `Launchers/Orchestrator/README.md` for preview, installation, immediate
start, and removal commands.

## Validation

Run the autonomous test suite:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Tests\Test-SmartWorkplaceCMDB.ps1
```

The suite validates safe identity keys, path resolution, runtime JSON synchronization, the 20 CSV contracts, preservation of compatible output, and rejection of incompatible output.

Run the offline Entra users pipeline tests:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Tests\Test-SmartWorkplaceCMDB-EntraUsers.ps1
```

The Entra users tests use only fictitious fixture data and temporary output
folders. They never connect to Microsoft Graph or modify canonical CMDB data.

Run the offline Entra groups pipeline tests:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Tests\Test-SmartWorkplaceCMDB-EntraGroups.ps1
```

The group tests use only fictitious fixture data and temporary output folders.

Run the offline Entra devices pipeline tests:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Tests\Test-SmartWorkplaceCMDB-EntraDevices.ps1
```

Run the offline Intune managed devices pipeline tests:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Tests\Test-SmartWorkplaceCMDB-IntuneDevices.ps1
```

The Intune tests use only fictitious fixture data and temporary output folders.

Run the offline primary user-device relationship tests:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Tests\Test-SmartWorkplaceCMDB-UserDeviceRelationships.ps1
```

These tests cover stable relationship keys, reference integrity, orphan
handling, schema-only empty outputs, and read-only validation.

Run the offline general relationship consolidation tests:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Tests\Test-SmartWorkplaceCMDB-Relationships.ps1
```

These tests cover the three relationship types, canonical keys and dates,
reference rejection, schema-only empty output, and read-only validation.

Run the offline Microsoft 365 subscribed SKUs pipeline tests:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Tests\Test-SmartWorkplaceCMDB-M365SubscribedSkus.ps1
```

The licensing tests cover bounded collection, aggregation, schema-only empty
results, stable keys, and read-only validation.

Run the offline Microsoft 365 user license assignment tests:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Tests\Test-SmartWorkplaceCMDB-M365UserLicenseAssignments.ps1
```

These tests cover direct and group assignments, state precedence, stable
user/SKU keys, reference integrity, bounded users, and empty schema outputs.

Run the offline Exchange Online mailbox pipeline tests:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Tests\Test-SmartWorkplaceCMDB-ExchangeOnlineMailboxes.ps1
```

The mailbox tests use only fictitious fixture data and temporary output. They
cover bounded collection, stable mailbox keys, optional Entra correlation,
identity rejection, and schema-only empty outputs.

Run the offline Active Directory pipeline tests:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Tests\Test-SmartWorkplaceCMDB-ActiveDirectory.ps1
```

These tests use fictitious directory objects and temporary outputs. They verify
all ten raw and normalized AD tables, stable keys, bounded collection, the
two-step orchestrator pipeline, and both centralized Active Directory launchers without
requiring the ActiveDirectory module or domain connectivity.

Run the offline SharePoint publication contract tests:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Tests\Test-SmartWorkplaceCMDB-SharePoint.ps1
```

These tests validate path mapping, template safety, the `SMART-CMDB/DATA`
target, and the live/unbounded publication guards without connecting to Graph or
SharePoint.

Run the offline data-quality normalization tests:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Tests\Test-SmartWorkplaceCMDB-DataQuality.ps1
```

These tests cover orphan references, unlinked mailboxes, missing dates, freshness
thresholds, stable keys, contract integrity, and read-only validation.

Run the offline tenant and date dimension tests:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Tests\Test-SmartWorkplaceCMDB-Dimensions.ps1
```

These tests cover source-derived and bounded calendars, leap days, invariant month
names, tenant mapping, sentinel exclusion, and contract integrity.

Run the offline orchestrator and centralized launcher tests:

```powershell
pwsh -NoProfile -ExecutionPolicy Bypass -File .\SmartWorkplaceCMDB\Tests\Test-SmartWorkplaceCMDB-Orchestrator.ps1
```

These tests execute the complete 22-step pipeline from fictitious fixtures,
validate 20 canonical and 17 source contracts, verify audit logging and bounded isolation, and inspect
every centralized Cloud launcher.

## Power BI Direction

The SmartWorkplaceCMDB Power BI deliverable starts with a direction and executive overview, then expands into technician and detailed inventory views.

Initial overview pages should focus on:

- CMDB coverage.
- User and device counts.
- Compliance posture.
- License assignment coverage.
- Mailbox coverage.
- Source freshness.
- Data quality findings.

Detailed inventory pages will add drill-through views for users, devices, relationships, licenses, mailboxes, stale objects, duplicates, and source conflicts.
