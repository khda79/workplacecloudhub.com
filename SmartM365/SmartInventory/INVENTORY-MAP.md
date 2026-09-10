# SmartInventory source map - audit lot 1

Source-only snapshot, 2026-09-10. No customer exports or operational JSON were read. This is a navigation index, not collector qualification. CSV literals are code references (including inputs), not a complete inferred export schema. Dynamic names require per-collector review.

Local Git baseline: `5d7341e81a4cacd014f0080406cb8d17d80b59f4`. 46 tracked PowerShell files outside Orchestrator/TestRuns; this includes utilities and reports, not only collectors. 42 manifest jobs; 37 enabled in the committed template (not operational state). Dashboard selects 90 CSV names; FinOps declares 48 source groups.

## Collectors and related entries

| Source relative to SmartInventory | Header version | Adjacent committed config template | CSV writer paths seen | Literal CSV references |
| --- | --- | --- | --- | --- |
| `ActiveDirectoryInventory/SmartM365-ActiveDirectory-EnrichedColumns.ps1` | 1.2 | none at conventional adjacent path; inspect script-specific loader |  | AD_Computers_AllDomains.csv |
| `ActiveDirectoryInventory/SmartM365-ActiveDirectory-Enrichment.ps1` | 1.2 | none at conventional adjacent path; inspect script-specific loader | Export-Csv | AD_Computers_AllDomains.csv, AD_Users_AllDomains_Brut.csv, AD_Users_AllDomains.csv, Exchange_EXO_Mailboxes_AllDomains_Stats.csv, Exchange_EXO_Mailboxes_AllDomains.csv, Exchange_OnPrem_Mailboxes_AllDomains.csv, Exchange_OnPrem_RemoteMailboxes_AllDomains.csv, Intune_Devices_Inventory.csv, Intune_Devices_LocalSystem.csv, Intune_Windows11_Readiness_Issues.csv, Intune_WindowsUpdate_Status.csv, M365_Entra_Devices_HardwareIdConflicts.csv, M365_Entra_Devices.csv, M365_Licenses_Users.csv |
| `ActiveDirectoryInventory/SmartM365-ActiveDirectory-HealthCheck.ps1` | 1.0.21 | same basename .local.json.template | Copy-Item, Export-Csv | AD_HealthCheck_History.csv, AD_HealthCheck.csv |
| `ActiveDirectoryInventory/SmartM365-ActiveDirectory-Inventory.ps1` | 1.43 | same basename .local.json.template | Copy-Item, Export-Csv, Write-SmartM365CsvAtomically | AD_Computers_AllDomains_Brut.csv, AD_Computers_AllDomains.csv, AD_Computers_DailyStats.csv, AD_Contacts_AllDomains.csv, AD_Groups_AllDomains.csv, AD_Inventory_DailySummary.csv, AD_OUs_AllDomains.csv, AD_Users_AllDomains_Brut.csv, AD_Users_AllDomains.csv, AD_Users_DailyStats.csv, AD_Users_DuplicateRemoteRoutingAddress.csv, AD_Users_DuplicateSMTP.csv, AD_Users_DuplicateUPN.csv, AD_Users_RemoteRoutingIssues.csv |
| `ActiveDirectoryInventory/SmartM365-ActiveDirectory-UsersEnrichment.ps1` | 1.5 | none at conventional adjacent path; inspect script-specific loader | Export-Csv | AD_Users_AllDomains.csv, Exchange_EXO_AcceptedDomains.csv, Exchange_EXO_Mailboxes_AllDomains_Stats.csv, Exchange_EXO_Mailboxes_AllDomains.csv, Exchange_EXO_MigrationJobs.csv, Exchange_HybridIdentity_Issues.csv, Exchange_Mailboxes_AllSources_PermissionsByUser.csv, Exchange_OnPrem_Mailboxes_AllDomains.csv, Exchange_OnPrem_RemoteMailboxes_AllDomains.csv, M365_BackupPolicyScope_MailboxCoverage.csv, M365_Entra_VerifiedDomains.csv, M365_Licenses_ServicePlans.csv, M365_Licenses_Users.csv, M365_Users_Active.csv |
| `ExchangeInventory/AcceptedDomains/SmartM365-EXO-AcceptedDomains-Inventory.ps1` | 1.8 | same basename .local.json.template | Copy-Item, Export-SmartM365Csv | Exchange_EXO_AcceptedDomains.csv |
| `ExchangeInventory/BackupProtection/SmartM365-M365-BackupPolicyScope-Inventory.ps1` | 1.5 | same basename .local.json.template | Copy-Item, Export-SmartM365Csv | Exchange_EXO_Mailboxes_AllDomains.csv |
| `ExchangeInventory/BackupProtection/SmartM365-M365-BackupProtectedMailboxes-Inventory.ps1` | 1.14 | same basename .local.json.template | Copy-Item, Export-SmartM365Csv | M365_Backup_ProtectedMailboxes.csv |
| `ExchangeInventory/CalendarPermissions/SmartM365-EXO-Mailboxes-CalPerm_Inventory.ps1` | 2.2 | same basename .local.json.template | Copy-Item, ExportAndCopyCsv |  |
| `ExchangeInventory/Mailboxes/SmartM365-EXO-Mailboxes-Inventory.ps1` | 1.20 | same basename .local.json.template | Copy-Item, Export-Csv, ExportAndCopyCsv | AD_Users_AllDomains.csv, Exchange_EXO_Mailboxes_AllDomains_Archive.csv, Exchange_EXO_Mailboxes_AllDomains_Permissions.csv, Exchange_EXO_Mailboxes_AllDomains_Stats.csv, Exchange_EXO_Mailboxes_AllDomains.csv, M365_Users_Active.csv |
| `ExchangeInventory/Migration/SmartM365-EXO-MigJob-Inventory.ps1` | 1.8 | same basename .local.json.template | Copy-Item, ExportAndCopyCsv |  |
| `ExchangeInventory/Migration/SmartM365-Exchange-HybridIdentity-Issues-Inventory.ps1` | 1.15 | same basename .local.json.template | Copy-Item, Export-Csv | AD_Users_AllDomains_Brut.csv, AD_Users_AllDomains.csv, AD_Users_DuplicateSMTP.csv, AD_Users_DuplicateUPN.csv, Exchange_EXO_AcceptedDomains.csv, Exchange_EXO_Mailboxes_AllDomains_Archive.csv, Exchange_EXO_Mailboxes_AllDomains_Permissions.csv, Exchange_EXO_Mailboxes_AllDomains_Stats.csv, Exchange_EXO_Mailboxes_AllDomains.csv, Exchange_HybridIdentity_Issues_Summary.csv, Exchange_HybridIdentity_Issues.csv, Exchange_OnPrem_Mailboxes_AllDomains.csv, Exchange_OnPrem_RemoteMailboxes_AllDomains.csv, M365_Entra_VerifiedDomains.csv, M365_Licenses_ServicePlans.csv, M365_Licenses_Users.csv, M365_Users_Active.csv |
| `ExchangeInventory/OnPremises/CalendarPermissions/SmartM365-Exchange-MailboxCalendarPermissions-Inventory.ps1` | 1.4 | same basename .local.json.template | Copy-Item, ExportAndCopyCsv |  |
| `ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1` | 1.42 | same basename .local.json.template | Copy-Item, Write-SmartM365CsvAtomically | Exchange_OnPrem_Mailboxes_AllDomains.csv, Exchange_OnPrem_Mailboxes_DailyStats_Summary.csv, Exchange_OnPrem_Mailboxes_DailyStats.csv, Exchange_OnPrem_RemoteMailboxes_AllDomains.csv |
| `ExchangeInventory/OnPremises/ProxyAddresses/SmartM365-Check-ProxyAddresses-Exchange.ps1` | 1.24 | same basename .local.json.template | Copy-Item, Publish-SmartM365Csv | Exchange_OnPrem_ProxyAddresses_Added.csv, Exchange_OnPrem_ProxyAddresses_Check.csv, Exchange_OnPrem_ProxyAddresses_Summary.csv |
| `ExchangeInventory/OnPremises/ServersAndStorage/SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1` | 1.6.0 | same basename .local.json.template | Copy-Item | Exchange_OnPrem_Infrastructure_PerServerSummary.csv, Exchange_OnPrem_MailboxDatabases_Paths.csv, Exchange_OnPrem_MigrationReadiness_Config.csv, Exchange_OnPrem_Servers_Compute.csv, Exchange_OnPrem_Servers_DiskDrives.csv, Exchange_OnPrem_Servers_Inventory_Summary.csv, Exchange_OnPrem_Servers_Inventory.csv, Exchange_OnPrem_Servers_LogicalDisks.csv, Exchange_OnPrem_Servers_RemoteAccess.csv, Exchange_OnPrem_Servers_ServiceHealth.csv |
| `ExchangeInventory/Permissions/SmartM365-Mailboxes-AllSources-PermissionsByUserReport.ps1` | 1.11 | same basename .local.json.template | Copy-Item, Export-SmartM365Csv | Exchange_Mailboxes_AllSources_PermissionsByUser.csv |
| `ExchangeInventory/Quarantine/SmartM365-EXO-QuarantineMessages-Report.ps1` | 1.0.2 | same basename .local.json.template | Export-SmartM365Csv | Exchange_EXO_QuarantineMessages.csv |
| `M365Inventory/Devices/SmartM365-EntraDevices-Inventory.ps1` | 1.12 | same basename .local.json.template | Copy-Item, Export-SmartM365Csv, ExportAndCopyCsv |  |
| `M365Inventory/Domains/SmartM365-VerifiedDomains-Inventory.ps1` | 1.9 | same basename .local.json.template | Copy-Item, Export-SmartM365Csv | M365_Entra_VerifiedDomains.csv |
| `M365Inventory/IntuneInventory/Applications/SmartM365-Intune-DiscoveredApps-Inventory.ps1` | 1.24 | same basename .local.json.template | Copy-Item, Export-Csv, ExportAndCopyCsv | Intune_DiscoveredApps_DeviceDetail.csv |
| `M365Inventory/IntuneInventory/Applications/Tests/Test-SmartM365-Intune-DiscoveredApps-Cache.ps1` | 1.0 | none at conventional adjacent path; inspect script-specific loader | Export-Csv | Intune_DiscoveredApps_AppDeviceRelations.csv |
| `M365Inventory/IntuneInventory/Applications/Tests/Test-SmartM365-Intune-DiscoveredApps-Logging.ps1` | 1.0 | none at conventional adjacent path; inspect script-specific loader | Export-Csv | Intune_DiscoveredApps_AppDeviceRelations_20260720_200000.csv, Intune_DiscoveredApps_AppDeviceRelations.csv, Intune_DiscoveredApps_DeviceDetail.csv |
| `M365Inventory/IntuneInventory/Autopilot/SmartM365-WindowsAutopilot-Inventory.ps1` | 1.9 | same basename .local.json.template | Copy-Item, ExportAndCopyCsv |  |
| `M365Inventory/IntuneInventory/Devices/SmartM365-Detect-DeviceSystemInfo.ps1` | not identified by header scan | none at conventional adjacent path; inspect script-specific loader |  |  |
| `M365Inventory/IntuneInventory/Devices/SmartM365-Device-System-Inventory.ps1` | 2.3 | same basename .local.json.template | Copy-Item, ExportAndCopyCsv |  |
| `M365Inventory/IntuneInventory/Devices/SmartM365-Devices-BIOS-Inventory.ps1` | 1.11 | same basename .local.json.template | Copy-Item, ExportAndCopyCsv |  |
| `M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Compliance-Inventory.ps1` | 1.15 | same basename .local.json.template | Copy-Item, Export-SmartM365Csv, Write-SmartM365CsvAtomically | Intune_Devices_Compliance_Policies.csv, Intune_Devices_Compliance.csv |
| `M365Inventory/IntuneInventory/Devices/SmartM365-Devices-Inventory.ps1` | 1.12 | same basename .local.json.template | Copy-Item, ExportAndCopyCsv |  |
| `M365Inventory/IntuneInventory/Devices/SmartM365-Devices-UpgradeEligibility.ps1` | 1.18 | same basename .local.json.template | Copy-Item, Export-SmartM365Csv |  |
| `M365Inventory/IntuneInventory/EndpointAnalytics/SmartM365-EndpointAnalytics-Inventory.ps1` | 1.0.6 | none at conventional adjacent path; inspect script-specific loader |  | Intune_EndpointAnalytics_AppReliability.csv, Intune_EndpointAnalytics_DataQuality.csv, Intune_EndpointAnalytics_DevicePerformance.csv, Intune_EndpointAnalytics_ModelPerformance.csv, Intune_EndpointAnalytics_OSReliability.csv, Intune_EndpointAnalytics_StartupDevices.csv, Intune_EndpointAnalytics_StartupModels.csv, Intune_EndpointAnalytics_StartupProcesses.csv, Intune_EndpointAnalytics_WorkFromAnywhere.csv |
| `M365Inventory/IntuneInventory/RBAC/SmartM365-Intune-RBAC-GroupMembers.ps1` | 1.12 | same basename .local.json.template | Copy-Item, ExportAndCopyCsvFromConvert |  |
| `M365Inventory/IntuneInventory/SmartM365-Export-IntuneRemediations.ps1` | 1.7 | same basename .local.json.template | Copy-Item |  |
| `M365Inventory/IntuneInventory/WindowsUpdate/AutopatchAlerts/SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.ps1` | 1.15 | same basename .local.json.template | Copy-Item | Intune_WindowsAutopatch_Alerts_Detail.csv, Intune_WindowsAutopatch_Alerts_PolicySummary.csv, Intune_WindowsAutopatch_Alerts_Summary.csv |
| `M365Inventory/IntuneInventory/WindowsUpdate/SmartM365-Intune-Windows11-Readiness-Issues-Inventory.ps1` | 1.20 | same basename .local.json.template | Copy-Item, Export-Csv | AD_Computers_AllDomains_Brut.csv, AD_Computers_AllDomains.csv, AD_Users_AllDomains_Brut.csv, AD_Users_AllDomains.csv, Intune_Devices_BIOS.csv, Intune_Devices_Compliance.csv, Intune_Devices_Inventory.csv, Intune_Devices_LocalSystem.csv, Intune_Devices_UpgradeEligibility.csv, Intune_Windows11_Readiness_Issues_Summary.csv, Intune_Windows11_Readiness_Issues.csv, Intune_WindowsUpdate_Status.csv, M365_Entra_Devices_HardwareIdConflicts.csv, M365_Entra_Devices.csv, M365_Inventory_Device_LocalSystem.csv, M365_Inventory_Devices.csv, M365_Inventory_EntraDevices.csv, M365_Licenses_Users.csv, M365_WindowsUpdate_Status_From_Intune.csv |
| `M365Inventory/IntuneInventory/WindowsUpdate/SmartM365-WinUpdate_Status_From_Intune.ps1` | 1.31 | same basename .local.json.template | Copy-Item | Intune_WindowsUpdate_Status.csv |
| `M365Inventory/Licensing/SmartM365-Licences-Inventory.ps1` | 1.14 | same basename .local.json.template | Copy-Item, Export-Csv, ExportAndCopyCsvFromConvert | M365_Licenses_Groups.csv, M365_Licenses_ServicePlans_Catalog.csv, M365_Licenses_ServicePlans_Detailed.csv, M365_Licenses_ServicePlans.csv, M365_Licenses_Tenant.csv, M365_Licenses_Users.csv, M365_Licenses_UserServicePlanStates_Detailed.csv, M365_Licenses_UserServicePlanStates.csv |
| `M365Inventory/Licensing/Tests/Test-SmartM365-Licences-ServicePlanStreaming.ps1` | 1.0 | none at conventional adjacent path; inspect script-specific loader | Copy-Item | M365_Licenses_UserServicePlanStates_Detailed.csv, M365_Licenses_UserServicePlanStates_MAXITEMS-5.csv, M365_Licenses_UserServicePlanStates.csv |
| `M365Inventory/PowerBI/SmartM365-PowerBIFabricActivity-Inventory.ps1` | 1.0.0 | same basename .local.json.template | Export-SmartM365Csv |  |
| `M365Inventory/SharePoint/SmartM365-SPO-Inventory.ps1` | 0.24 | same basename .local.json.template | Copy-Item, Export-Csv, Export-SmartM365Csv | M365_Licenses_Tenant.csv |
| `M365Inventory/SyncHealth/SmartM365-AzureADConnect-SyncHealth-Inventory.ps1` | 1.15 | same basename .local.json.template | Copy-Item, Export-SmartM365Csv | M365_Entra_AzureADConnect_SyncHealth.csv |
| `M365Inventory/Teams/SmartM365-Teams-Inventory.ps1` | 0.25 | same basename .local.json.template | Export-Csv | M365_Teams_Channels_History.csv, M365_Teams_Channels.csv, M365_Teams_Guests_History.csv, M365_Teams_Guests.csv, M365_Teams_Members_History.csv, M365_Teams_Members.csv, M365_Teams_Teams_History.csv, M365_Teams_Teams.csv |
| `M365Inventory/Teams/SmartM365-TeamsPhonePstnUsage-Inventory.ps1` | 1.4 | same basename .local.json.template | Export-SmartM365Csv |  |
| `M365Inventory/Usage/SmartM365-CopilotUsage-Inventory.ps1` | 1.0.0 | same basename .local.json.template | Export-SmartM365Csv, Write-SmartM365CsvAtomically |  |
| `M365Inventory/Usage/SmartM365-M365UserActivity-Inventory.ps1` | 1.15 | same basename .local.json.template | Export-SmartM365Csv |  |
| `M365Inventory/Users/SmartM365-ActiveUsers-Inventory.ps1` | 1.9 | same basename .local.json.template | Copy-Item, ExportAndCopyCsvFromConvert |  |

## Orchestrator manifest topology

Template only. No runtime manifest was loaded or changed. Preserve existing schedules, algorithms, assignments and enabled flags.

| Job | Script path | Mode | Dependencies | Concurrency key | Capabilities / Graph application roles |
| --- | --- | --- | --- | --- | --- |
| M365-VerifiedDomains-Inventory | `M365Inventory\Domains\SmartM365-VerifiedDomains-Inventory.ps1` | Elected |  |  | Graph, SharedRuntime / Directory.Read.All |
| AD-Inventory | `ActiveDirectoryInventory\SmartM365-ActiveDirectory-Inventory.ps1` | Elected |  |  | AD, SharedRuntime /  |
| AD-HealthCheck | `ActiveDirectoryInventory\SmartM365-ActiveDirectory-HealthCheck.ps1` | Elected |  |  | AD, SharedRuntime /  |
| M365-ActiveUsers-Inventory | `M365Inventory\Users\SmartM365-ActiveUsers-Inventory.ps1` | Elected |  |  | Graph, SharedRuntime / User.Read.All, AuditLog.Read.All, Directory.Read.All |
| M365-Licences-Inventory | `M365Inventory\Licensing\SmartM365-Licences-Inventory.ps1` | Elected |  |  | Graph, SharedRuntime / Directory.Read.All, User.Read.All, Group.Read.All |
| M365-EntraDevices-Inventory | `M365Inventory\Devices\SmartM365-EntraDevices-Inventory.ps1` | Elected |  |  | Graph, SharedRuntime / Directory.Read.All, Device.Read.All |
| M365-Teams-Inventory | `M365Inventory\Teams\SmartM365-Teams-Inventory.ps1` | Elected |  |  | Graph, SharedRuntime / Team.ReadBasic.All, TeamMember.Read.All, Channel.ReadBasic.All, Group.Read.All, Reports.Read.All, Sites.Read.All |
| M365-SPO-Inventory | `M365Inventory\SharePoint\SmartM365-SPO-Inventory.ps1` | Elected |  |  | Graph, SharedRuntime / Reports.Read.All, Sites.Read.All, Directory.Read.All |
| Intune-Devices-Inventory | `M365Inventory\IntuneInventory\Devices\SmartM365-Devices-Inventory.ps1` | Elected |  |  | Graph, SharedRuntime / DeviceManagementManagedDevices.Read.All, Device.Read.All |
| Intune-Devices-Compliance-Inventory | `M365Inventory\IntuneInventory\Devices\SmartM365-Devices-Compliance-Inventory.ps1` | Elected |  |  | Graph, SharedRuntime / DeviceManagementManagedDevices.Read.All, DeviceManagementConfiguration.Read.All, Device.Read.All |
| Intune-WindowsAutopilot-Inventory | `M365Inventory\IntuneInventory\Autopilot\SmartM365-WindowsAutopilot-Inventory.ps1` | Elected |  |  | Graph, SharedRuntime / DeviceManagementServiceConfig.Read.All, Device.Read.All |
| EXO-Mailboxes-Inventory | `ExchangeInventory\Mailboxes\SmartM365-EXO-Mailboxes-Inventory.ps1` | Elected |  | EXOMailboxes | EXO, Graph, SharedRuntime / User.Read.All |
| EXO-Mailboxes-Permissions | `ExchangeInventory\Mailboxes\SmartM365-EXO-Mailboxes-Inventory.ps1` | Elected |  |  | EXO, SharedRuntime /  |
| EXO-Mailboxes-Inventory-Fast | `ExchangeInventory\Mailboxes\SmartM365-EXO-Mailboxes-Inventory.ps1` | Elected |  | EXOMailboxes | EXO, Graph, SharedRuntime / User.Read.All |
| EXO-Mailboxes-CalendarPermissions | `ExchangeInventory\CalendarPermissions\SmartM365-EXO-Mailboxes-CalPerm_Inventory.ps1` | Elected |  |  | EXO, SharedRuntime /  |
| Exchange2016-Mailboxes-CalendarPermissions | `ExchangeInventory\OnPremises\CalendarPermissions\SmartM365-Exchange-MailboxCalendarPermissions-Inventory.ps1` | Elected |  |  | AD, ExchangeOnPrem, SharedRuntime /  |
| EXO-AcceptedDomains-Inventory | `ExchangeInventory\AcceptedDomains\SmartM365-EXO-AcceptedDomains-Inventory.ps1` | Elected |  |  | EXO, SharedRuntime /  |
| EXO-QuarantineMessages-Report | `ExchangeInventory\Quarantine\SmartM365-EXO-QuarantineMessages-Report.ps1` | Elected |  |  | EXO, SharedRuntime /  |
| EXO-MigrationJobs-Inventory | `ExchangeInventory\Migration\SmartM365-EXO-MigJob-Inventory.ps1` | Elected |  |  | EXO, SharedRuntime /  |
| Mailboxes-PermissionsByUser-Report | `ExchangeInventory\Permissions\SmartM365-Mailboxes-AllSources-PermissionsByUserReport.ps1` | Elected |  |  | SharedRuntime /  |
| Exchange2016-Local-Mailboxes-Inventory | `ExchangeInventory\OnPremises\Mailboxes\SmartM365-Exchange-Local-Mailboxes-Inventory.ps1` | Elected |  | ExchangeLocalMailboxes | AD, ExchangeOnPrem, SharedRuntime /  |
| Exchange2016-Local-Mailboxes-Fast | `ExchangeInventory\OnPremises\Mailboxes\SmartM365-Exchange-Local-Mailboxes-Inventory.ps1` | Elected |  | ExchangeLocalMailboxes | AD, ExchangeOnPrem, SharedRuntime /  |
| Exchange2016-ProxyAddresses-Check | `ExchangeInventory\OnPremises\ProxyAddresses\SmartM365-Check-ProxyAddresses-Exchange.ps1` | Elected |  |  | AD, ExchangeOnPrem, SharedRuntime /  |
| Exchange2016-Infrastructure-Inventory | `ExchangeInventory\OnPremises\ServersAndStorage\SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1` | Elected |  |  | AD, ExchangeOnPrem, SharedRuntime /  |
| M365-SyncHealth-Inventory | `M365Inventory\SyncHealth\SmartM365-AzureADConnect-SyncHealth-Inventory.ps1` | Elected |  |  | Graph, SharedRuntime / Directory.Read.All |
| M365-Usage-Inventory | `M365Inventory\Usage\SmartM365-M365UserActivity-Inventory.ps1` | Elected |  |  | Graph, SharedRuntime / Reports.Read.All |
| Intune-DiscoveredApps-Inventory | `M365Inventory\IntuneInventory\Applications\SmartM365-Intune-DiscoveredApps-Inventory.ps1` | Elected |  |  | Graph, SharedRuntime / DeviceManagementApps.Read.All, DeviceManagementManagedDevices.Read.All |
| Intune-DeviceSystem-Inventory | `M365Inventory\IntuneInventory\Devices\SmartM365-Device-System-Inventory.ps1` | Elected |  |  | Graph, SharedRuntime / DeviceManagementManagedDevices.Read.All, DeviceManagementConfiguration.Read.All, DeviceManagementScripts.Read.All |
| Intune-Devices-BIOS-Inventory | `M365Inventory\IntuneInventory\Devices\SmartM365-Devices-BIOS-Inventory.ps1` | Elected |  |  | Graph, SharedRuntime / DeviceManagementManagedDevices.Read.All |
| Intune-Devices-UpgradeEligibility | `M365Inventory\IntuneInventory\Devices\SmartM365-Devices-UpgradeEligibility.ps1` | Elected |  |  | Graph, SharedRuntime / DeviceManagementManagedDevices.Read.All |
| Intune-RBAC-GroupMembers | `M365Inventory\IntuneInventory\RBAC\SmartM365-Intune-RBAC-GroupMembers.ps1` | Elected |  |  | Graph, SharedRuntime / Group.Read.All, GroupMember.Read.All |
| Intune-Remediations-Export | `M365Inventory\IntuneInventory\SmartM365-Export-IntuneRemediations.ps1` | Elected |  |  | Graph, SharedRuntime / DeviceManagementScripts.Read.All, Group.Read.All |
| Intune-AutopatchAlerts-Inventory | `M365Inventory\IntuneInventory\WindowsUpdate\AutopatchAlerts\SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.ps1` | Elected |  |  | Graph, SharedRuntime / DeviceManagementConfiguration.Read.All, DeviceManagementManagedDevices.Read.All, DeviceManagementApps.Read.All |
| Intune-Windows11-Readiness-Issues | `M365Inventory\IntuneInventory\WindowsUpdate\SmartM365-Intune-Windows11-Readiness-Issues-Inventory.ps1` | Elected | Intune-Devices-UpgradeEligibility |  | SharedRuntime /  |
| Intune-WinUpdate-Status | `M365Inventory\IntuneInventory\WindowsUpdate\SmartM365-WinUpdate_Status_From_Intune.ps1` | Elected |  |  | Graph, SharedRuntime / DeviceManagementConfiguration.Read.All, DeviceManagementManagedDevices.Read.All |
| Exchange-HybridIdentity-Issues | `ExchangeInventory\Migration\SmartM365-Exchange-HybridIdentity-Issues-Inventory.ps1` | Elected |  |  | SharedRuntime /  |
| M365-BackupProtectedMailboxes-Inventory | `ExchangeInventory\BackupProtection\SmartM365-M365-BackupProtectedMailboxes-Inventory.ps1` | Elected |  |  | Graph, SharedRuntime / BackupRestore-Configuration.Read.All |
| M365-BackupPolicyScope-Inventory | `ExchangeInventory\BackupProtection\SmartM365-M365-BackupPolicyScope-Inventory.ps1` | Elected | EXO-Mailboxes-Inventory-Fast |  | Graph, SharedRuntime / GroupMember.Read.All, User.Read.All |
| M365-CopilotUsage-Inventory | `M365Inventory\Usage\SmartM365-CopilotUsage-Inventory.ps1` | Elected |  |  | Graph, SharedRuntime / Reports.Read.All |
| M365-TeamsPhonePstnUsage-Inventory | `M365Inventory\Teams\SmartM365-TeamsPhonePstnUsage-Inventory.ps1` | Elected |  |  | Graph, SharedRuntime, TeamsPowerShell / CallRecords.Read.All, Organization.Read.All |
| Intune-EndpointAnalytics-Inventory | `M365Inventory\IntuneInventory\EndpointAnalytics\SmartM365-EndpointAnalytics-Inventory.ps1` | Elected |  |  | Graph, SharedRuntime / DeviceManagementManagedDevices.ReadWrite.All |
| M365-PowerBIFabricActivity-Inventory | `M365Inventory\PowerBI\SmartM365-PowerBIFabricActivity-Inventory.ps1` | Manual |  |  |  /  |

## Consumer contracts

SmartFinOps Workplace reads one tenant DATA-LAST root. Its Config/SmartFinOps-Workplace-SourceContracts.json defines aliases/required fields; Config/SmartFinOps-Workplace-DataContract.ps1 handles CSV parsing, identity, missing sources and freshness. MAXITEMS files are excluded. No analyzer code was modified or executed in this lot.

| FinOps source | Accepted CSV filenames | Required columns |
| --- | --- | --- |
| M365ActiveUsers | M365_Users_Active.csv | User principal name |
| M365UserActivity | M365_Users_Activity.csv | UserPrincipalName |
| M365MailboxUsage | M365_Mailbox_Usage.csv | User Principal Name, Last Activity Date, Storage Used (Byte), Has Archive |
| M365OneDriveUsage | M365_OneDrive_Usage.csv | Owner Principal Name, Last Activity Date, Storage Used (Byte) |
| M365SharePointUserActivity | M365_SharePoint_UserActivity.csv | User Principal Name, Last Activity Date, Viewed Or Edited File Count, Report Period |
| M365TeamsDeviceUsage | M365_Teams_DeviceUsage.csv | User Principal Name, Last Activity Date, Used Web, Used Windows, Is Licensed |
| M365CopilotUserUsage | M365_Copilot_UserUsage.csv | UserPrincipalName, LastActivityDate, ReportPeriod |
| M365TeamsPhoneUserUsage | M365_Teams_PhoneUserUsage.csv | UserPrincipalName, LastCallDate, TotalCallCount, TotalMinutes, HasPstnUsage, HasDirectRoutingUsage |
| M365TeamsPhoneAssignments | M365_Teams_PhoneAssignments.csv | AssignedPstnTargetId, TargetType, Capability, PstnAssignmentStatus, TelephoneNumberMasked |
| M365TeamsPstnCalls | M365_Teams_PSTNCalls.csv | UserPrincipalName, StartDateTimeUtc, Minutes, Direction, Charge, Currency |
| M365TeamsDirectRoutingCalls | M365_Teams_DirectRoutingCalls.csv | UserPrincipalName, StartDateTimeUtc, Minutes, SuccessfulCall, TrunkFullyQualifiedDomainName |
| M365PowerBIFabricActivityEvents | M365_PowerBI_Fabric_ActivityEvents.csv | CreationTime, Activity, PrincipalType |
| M365PowerBIFabricUserActivity | M365_PowerBI_Fabric_UserActivity.csv | UserPrincipalName, PrincipalType, LastActivityDate, TotalActivityCount |
| M365AppsActivations | M365_Apps_Activations.csv | User Principal Name, Product Type, Last Activated Date, Windows, Mac |
| M365TeamsUserActivity | M365_Teams_UserActivity.csv | User Principal Name, Last Activity Date |
| M365EmailActivity | M365_Email_Activity.csv | User Principal Name, Last Activity Date |
| M365LicenseUsers | M365_Licenses_Users.csv | User principal name, SkuPartNumber, Source, GroupsAssigningSku, GroupCountForSku, HasDirectAndGroup |
| M365LicenseGroups | M365_Licenses_Groups.csv | GroupId, GroupDisplayName, UsersCount, SkuPartNumbers |
| M365LicenseTenant | M365_Licenses_Tenant.csv | TenantSkuPartNumber, TenantPrepaidEnabled, TenantConsumedUnits |
| IntuneEndpointAnalyticsDevicePerformance | Intune_EndpointAnalytics_DevicePerformance.csv | DeviceId, DeviceName, EndpointAnalyticsScore, StartupScore, AppReliabilityScore, WorkFromAnywhereScore |
| IntuneEndpointAnalyticsModelPerformance | Intune_EndpointAnalytics_ModelPerformance.csv | Manufacturer, Model, EndpointAnalyticsScore |
| IntuneEndpointAnalyticsStartupDevices | Intune_EndpointAnalytics_StartupDevices.csv | DeviceId, DeviceName, StartupScore, CoreBootTime, CoreSignInTime |
| IntuneEndpointAnalyticsStartupModels | Intune_EndpointAnalytics_StartupModels.csv | Manufacturer, Model, StartupScore |
| IntuneEndpointAnalyticsStartupProcesses | Intune_EndpointAnalytics_StartupProcesses.csv | ApplicationName, Publisher, UsageDuration |
| IntuneEndpointAnalyticsAppReliability | Intune_EndpointAnalytics_AppReliability.csv | ApplicationName, AppReliabilityScore, CrashCount, MeanTimeToFailure |
| IntuneEndpointAnalyticsOSReliability | Intune_EndpointAnalytics_OSReliability.csv | OSVersion, AppReliabilityScore, MeanTimeToFailure |
| IntuneEndpointAnalyticsWorkFromAnywhere | Intune_EndpointAnalytics_WorkFromAnywhere.csv | DeviceId, DeviceName, WorkFromAnywhereScore, CloudManagementScore, WindowsScore |
| IntuneEndpointAnalyticsDataQuality | Intune_EndpointAnalytics_DataQuality.csv | ReportName, Status, RowCount, IsAdvancedAnalytics, ErrorCode, ErrorMessage |
| IntuneDevices | Intune_Devices_Inventory.csv | Device ID, Device name, Last check-in |
| IntuneUpgradeEligibility | Intune_Devices_UpgradeEligibility.csv, Intune_Devices_Windows11UpgradeEligibility.csv | GraphId, DeviceName, UpgradeEligibility, UpgradeEligibilityLabel |
| IntuneUpgradeEligibilitySummary | Intune_Devices_UpgradeEligibility_Summary.csv | TotalDeviceCount, UpgradeEligibleDeviceCount, UpgradeEligiblePercentage, ReportMode |
| IntuneWindowsUpdateStatus | Intune_WindowsUpdate_Status.csv | DeviceId, DeviceName, RiskBucket, BlockingReason, ActionPriority |
| IntuneCompliance | Intune_Devices_Compliance.csv | DeviceName, AzureADDeviceId, ComplianceState |
| EntraDevices | M365_Entra_Devices.csv | ObjectId, DeviceId, DisplayName |
| EntraConnectSyncHealth | M365_Entra_AzureADConnect_SyncHealth.csv | CheckName, Status, LastSyncDateTimeUtc, SyncAgeMinutes |
| ADUsersCanonical | AD_Users_AllDomains.csv | UserPrincipalName, Enabled, M365LicenseType, M365LicenseTargetPersona, IsE3PersonaMissingE3License, HasRemainingF3MailboxSizeNonCompliance, IsUnlicensedEnabledF3MailboxBlockingMigration |
| ADUsersRaw | AD_Users_AllDomains_Brut.csv | UserPrincipalName, Enabled |
| ADComputersCanonical | AD_Computers_AllDomains.csv | Name, Enabled, Windows11UpgradeEligibility, NeedsWindows11Upgrade, ExistsInIntune |
| ADComputersRaw | AD_Computers_AllDomains_Brut.csv | Name, Enabled |
| EXOMailboxes | Exchange_EXO_Mailboxes_AllDomains.csv | UserPrincipalName, PrimarySmtpAddress, RecipientTypeDetails, AccountEnabled, TotalItemSizeGB, ArchiveStatus, LitigationHoldEnabled, RetentionHoldEnabled |
| EXOMailboxStats | Exchange_EXO_Mailboxes_AllDomains_Stats.csv | UserPrincipalName, TotalItemSizeGB, Archive_TotalItemSizeGB |
| EXOMailboxArchive | Exchange_EXO_Mailboxes_AllDomains_Archive.csv | UserPrincipalName, Archive_TotalItemSizeGB |
| EXOMailboxPermissions | Exchange_EXO_Mailboxes_AllDomains_Permissions.csv | UserPrincipalName, SendAs, FullAccess, GrantSendOnBehalfTo |
| ExchangeAllSourcesPermissionsByUser | Exchange_Mailboxes_AllSources_PermissionsByUser.csv | ResolvedUser, Source, FullAccessOn, SendAsOn, SendOnBehalfOn, HasCrossPremisesPermissions |
| RemoteRoutingIssues | AD_Users_RemoteRoutingIssues.csv | IssueType, Severity, UserPrincipalName, Enabled |
| ProxyAddressControl | Exchange_OnPrem_ProxyAddresses_Check.csv | Identity, MailboxLocation, Status, ExpectedAddressConflict |
| BackupProtectedMailboxes | M365_Backup_ProtectedMailboxes.csv | ProtectionUnitId, Status, UserPrincipalName |
| BackupPolicyScope | M365_BackupPolicyScope_MailboxCoverage.csv | MemberUserPrincipalName, MailboxFoundInInventory, ExpectedPolicyScopeStatus |

SmartWorkplaceDashboard directly consumes DATA-LAST plus DATA-ALL weekly history. It is independent of SmartWorkplaceCMDB. source-selection.json selects files; column-mapping.json and derived-table-mapping.json document public schemas. TenantKey and source identifiers must retain their existing names and values. Export timestamps do not establish service report freshness.

Dashboard selected CSVs:

- `AD_Computers_AllDomains.csv`
- `AD_Computers_DailyStats.csv`
- `AD_Contacts_AllDomains.csv`
- `AD_HealthCheck.csv`
- `AD_Users_AllDomains.csv`
- `AD_Users_DailyStats.csv`
- `AD_Users_DuplicateRemoteRoutingAddress.csv`
- `AD_Users_DuplicateSMTP.csv`
- `AD_Users_DuplicateUPN.csv`
- `AD_Users_RemoteRoutingIssues.csv`
- `Exchange_EXO_AcceptedDomains.csv`
- `Exchange_EXO_MailboxCalendarPermissions_AllDomains.csv`
- `Exchange_EXO_Mailboxes_AllDomains.csv`
- `Exchange_EXO_MigrationJobs.csv`
- `Exchange_HybridIdentity_Issues.csv`
- `Exchange_Mailboxes_AllSources_PermissionsByUser.csv`
- `Exchange_OnPrem_MailboxCalendarPermissions_AllDomains.csv`
- `Exchange_OnPrem_Mailboxes_AllDomains.csv`
- `Exchange_OnPrem_Mailboxes_DailyStats.csv`
- `Exchange_OnPrem_MigrationReadiness_Config.csv`
- `Exchange_OnPrem_RemoteMailboxes_AllDomains.csv`
- `Exchange_OnPrem_Servers_Compute.csv`
- `Exchange_OnPrem_Servers_DiskDrives.csv`
- `Exchange_OnPrem_Servers_Inventory.csv`
- `Exchange_OnPrem_Servers_Inventory_Summary.csv`
- `Exchange_OnPrem_Servers_LogicalDisks.csv`
- `Exchange_OnPrem_Servers_RemoteAccess.csv`
- `Intune_AutopatchAlerts_Detail.csv`
- `Intune_Autopilot_Devices.csv`
- `Intune_Devices_BIOS.csv`
- `Intune_Devices_Compliance_Policies.csv`
- `Intune_Devices_Compliance.csv`
- `Intune_Devices_Inventory.csv`
- `Intune_Devices_LocalSystem.csv`
- `Intune_Devices_UpgradeEligibility.csv`
- `Intune_Devices_Win11Readiness.csv`
- `Intune_DiscoveredApps_AppDeviceRelations.csv`
- `Intune_DiscoveredApps_Summary.csv`
- `Intune_EndpointAnalytics_AppReliability.csv`
- `Intune_EndpointAnalytics_DataQuality.csv`
- `Intune_EndpointAnalytics_DevicePerformance.csv`
- `Intune_EndpointAnalytics_ModelPerformance.csv`
- `Intune_EndpointAnalytics_OSReliability.csv`
- `Intune_EndpointAnalytics_StartupDevices.csv`
- `Intune_EndpointAnalytics_StartupModels.csv`
- `Intune_EndpointAnalytics_StartupProcesses.csv`
- `Intune_EndpointAnalytics_WorkFromAnywhere.csv`
- `Intune_Windows11_Readiness_Issues.csv`
- `Intune_WindowsAutopatch_Alerts_Detail.csv`
- `Intune_WindowsUpdate_Status.csv`
- `M365_Apps_Activations.csv`
- `M365_Backup_ProtectedMailboxes.csv`
- `M365_BackupPolicyScope_GroupMembers.csv`
- `M365_BackupPolicyScope_GroupMembersWithoutMailbox.csv`
- `M365_BackupPolicyScope_MailboxCoverage.csv`
- `M365_Copilot_UserUsage.csv`
- `M365_Email_Activity.csv`
- `M365_Entra_AzureADConnect_SyncHealth.csv`
- `M365_Entra_Devices_HardwareIdConflicts_RegisteredPending.csv`
- `M365_Entra_Devices_HardwareIdConflicts.csv`
- `M365_Entra_Devices_RegisteredPending.csv`
- `M365_Entra_Devices_RemovalCandidates.csv`
- `M365_Entra_Devices.csv`
- `M365_Entra_VerifiedDomains.csv`
- `M365_Licenses_Groups.csv`
- `M365_Licenses_ServicePlans_Catalog.csv`
- `M365_Licenses_Tenant.csv`
- `M365_Licenses_UserServicePlanStates.csv`
- `M365_Licenses_Users.csv`
- `M365_Mailbox_Usage.csv`
- `M365_OneDrive_Usage.csv`
- `M365_SharePoint_SiteUsage.csv`
- `M365_SharePoint_UserActivity.csv`
- `M365_SPO_ExternalSharing.csv`
- `M365_SPO_Lists.csv`
- `M365_SPO_Permissions.csv`
- `M365_SPO_Sites.csv`
- `M365_SPO_Tenant.csv`
- `M365_Teams_Channels.csv`
- `M365_Teams_DeviceUsage.csv`
- `M365_Teams_DirectRoutingCalls.csv`
- `M365_Teams_Guests.csv`
- `M365_Teams_Members.csv`
- `M365_Teams_PhoneAssignments.csv`
- `M365_Teams_PhoneUserUsage.csv`
- `M365_Teams_PSTNCalls.csv`
- `M365_Teams_Teams.csv`
- `M365_Teams_UserActivity.csv`
- `M365_Users_Active.csv`
- `M365_Users_Activity.csv`

Other consumers found in source/documentation include EXO mailbox enrichment (AD/M365 users), cross-source mailbox-permissions reporting, device LOT selection in IntuneHybridJoinToolkit and Windows11UpgradeToolkit, and optional SharePoint CSV distribution. Those consumers are outside this correction lot. SmartWorkplaceCMDB has independent native collectors; do not assume it is the required intermediary for Dashboard.
