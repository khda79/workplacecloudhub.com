# Frozen reference of the 180 SmartInventory .cmd launchers removed in lot 4 (launch catalog migration).
# Generated from the legacy launchers before their removal; read by Test-SmartM365InventoryLaunchOffline.ps1
# to keep the catalog entries identical to what each launcher started. Retired readiness producers are excluded from active-launch equivalence.
# Paths are relative to SmartInventory; Tenant is the value the launcher passed when
# SMARTM365_ORCHESTRATOR_TENANT was prod (hard-coded launchers always passed prod).
@{
    Launchers = @(
    @{
        Launcher = 'Launchers\Cloud\Convert-SmartM365-WorkplaceEvidence-Metadata.cmd'
        Entry = 'WorkplaceEvidence-ConvertMetadata'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'PreparedEvidence\SmartM365-WorkplaceEvidence-Prepare.ps1'
        Arguments = @('-ConvertMetadata')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = 'PreparedEvidence'
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Repair-SmartM365-WorkplaceEvidence-TenantKeys.cmd'
        Entry = 'WorkplaceEvidence-RepairTenantKeys'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'PreparedEvidence\SmartM365-WorkplaceEvidence-Prepare.ps1'
        Arguments = @('-RepairLegacyHistory')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = 'PreparedEvidence'
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-ActiveUsers-Inventory.cmd'
        Entry = 'M365-ActiveUsers-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Users\SmartM365-ActiveUsers-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-AuthenticationMethodsRegistration-Inventory.cmd'
        Entry = 'M365-AuthenticationMethodsRegistration-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Security\AuthenticationMethods\SmartM365-AuthenticationMethodsRegistration-Inventory.ps1'
        Arguments = @('-EnableConfiguredExternalActions')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-AzureADConnect-SyncHealth-Inventory.cmd'
        Entry = 'M365-SyncHealth-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\SyncHealth\SmartM365-AzureADConnect-SyncHealth-Inventory.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-ConditionalAccess-Inventory.cmd'
        Entry = 'M365-ConditionalAccess-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Security\ConditionalAccess\SmartM365-ConditionalAccess-Inventory.ps1'
        Arguments = @('-EnableConfiguredExternalActions')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-CopilotUsage-Inventory.cmd'
        Entry = 'M365-CopilotUsage-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Usage\SmartM365-CopilotUsage-Inventory.ps1'
        Arguments = @('-Period', 'D180', '-ReportVersion', 'v2', '-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-Create-DeviceSystemRemediation.cmd'
        Entry = 'Intune-DeviceSystemRemediation-Create'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\Devices\SmartM365-Create-DeviceSystemRemediation.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-Device-System-Inventory.cmd'
        Entry = 'Intune-DeviceSystem-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\Devices\SmartM365-Device-System-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-Devices-BIOS-Inventory.cmd'
        Entry = 'Intune-Devices-BIOS-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\Devices\SmartM365-Devices-BIOS-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-Devices-Compliance-Inventory.cmd'
        Entry = 'Intune-Devices-Compliance-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\Devices\SmartM365-Devices-Compliance-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-Devices-Inventory.cmd'
        Entry = 'Intune-Devices-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\Devices\SmartM365-Devices-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-Devices-UpgradeEligibility.cmd'
        Entry = 'Intune-Devices-UpgradeEligibility'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\Devices\SmartM365-Devices-UpgradeEligibility.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-EndpointAnalytics-Inventory-IncludeStartupProcesses.cmd'
        Entry = 'Intune-EndpointAnalytics-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\EndpointAnalytics\SmartM365-EndpointAnalytics-Inventory.ps1'
        Arguments = @('-Reports', 'All', '-Connect', '-IncludeStartupProcesses')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-EndpointAnalytics-Inventory.cmd'
        Entry = 'Intune-EndpointAnalytics-Inventory-NoStartupProcesses'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\EndpointAnalytics\SmartM365-EndpointAnalytics-Inventory.ps1'
        Arguments = @('-Reports', 'All', '-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-EntraDevices-Inventory.cmd'
        Entry = 'M365-EntraDevices-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Devices\SmartM365-EntraDevices-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },

    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-EXO-AcceptedDomains-Inventory.cmd'
        Entry = 'EXO-AcceptedDomains-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\AcceptedDomains\SmartM365-EXO-AcceptedDomains-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-EXO-Mailboxes-CalPerm_Inventory.cmd'
        Entry = 'EXO-Mailboxes-CalendarPermissions'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\CalendarPermissions\SmartM365-EXO-Mailboxes-CalPerm_Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-EXO-Mailboxes-Inventory-IncludeStats.cmd'
        Entry = 'EXO-Mailboxes-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\Mailboxes\SmartM365-EXO-Mailboxes-Inventory.ps1'
        Arguments = @('-Connect', '-IncludeStats', '-ForceLiveStats')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-EXO-Mailboxes-Inventory.cmd'
        Entry = 'EXO-Mailboxes-Inventory-Fast'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\Mailboxes\SmartM365-EXO-Mailboxes-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-EXO-MigJob-Inventory.cmd'
        Entry = 'EXO-MigrationJobs-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\Migration\SmartM365-EXO-MigJob-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-EXO-QuarantineMessages-Report.cmd'
        Entry = 'EXO-QuarantineMessages-Report'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\Quarantine\SmartM365-EXO-QuarantineMessages-Report.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-Intune-DiscoveredApps-Inventory-FreshDeviceDetails.cmd'
        Entry = 'Intune-DiscoveredApps-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\Applications\SmartM365-Intune-DiscoveredApps-Inventory.ps1'
        Arguments = @('-Connect', '-DeviceDetailMode', 'All', '-FreshDeviceDetails')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-Intune-DiscoveredApps-Inventory.cmd'
        Entry = 'Intune-DiscoveredApps-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\Applications\SmartM365-Intune-DiscoveredApps-Inventory.ps1'
        Arguments = @('-Connect', '-DeviceDetailMode', 'All', '-FreshDeviceDetails')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-Intune-EndpointSecurityHealth-Inventory.cmd'
        Entry = 'Intune-EndpointSecurityHealth-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\Security\SmartM365-Intune-EndpointSecurityHealth-Inventory.ps1'
        Arguments = @('-EnableConfiguredExternalActions', '-ExportTimeoutSeconds', '1200')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-Intune-ExportRemediationScripts.cmd'
        Entry = 'Intune-Remediations-Export'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\SmartM365-Intune-ExportRemediationScripts.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-Intune-RBAC-GroupMembers.cmd'
        Entry = 'Intune-RBAC-GroupMembers'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\RBAC\SmartM365-Intune-RBAC-GroupMembers.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },

    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.cmd'
        Entry = 'Intune-AutopatchAlerts-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\WindowsUpdate\AutopatchAlerts\SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.ps1'
        Arguments = @('-IncludeFeatureUpdates', '-IncludeQualityUpdates')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-IntuneConfigurations-Inventory.cmd'
        Entry = 'Intune-Configurations-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\WorkplaceScope\SmartM365-IntuneConfigurations-Inventory.ps1'
        Arguments = @('-EnableConfiguredExternalActions')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-IntuneEnrollmentFailures-Inventory.cmd'
        Entry = 'Intune-EnrollmentFailures-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\WorkplaceScope\SmartM365-IntuneEnrollmentFailures-Inventory.ps1'
        Arguments = @('-EnableConfiguredExternalActions')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-IntuneManagedApps-Inventory.cmd'
        Entry = 'Intune-ManagedApps-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\WorkplaceScope\SmartM365-IntuneManagedApps-Inventory.ps1'
        Arguments = @('-EnableConfiguredExternalActions')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-Licences-Inventory.cmd'
        Entry = 'M365-Licences-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Licensing\SmartM365-Licences-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-LicensePricing-Inventory.cmd'
        Entry = 'M365-LicensePricing-Inventory-PublicPrices'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Licensing\SmartM365-LicensePricing-Inventory.ps1'
        Arguments = @('-PricingSourceCsvPath', '<SmartInventory>\M365Inventory\Licensing\M365_License_Prices_Public-France.csv', '-EnableConfiguredExternalActions')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-M365-BackupPolicyScope-Inventory.cmd'
        Entry = 'M365-BackupPolicyScope-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\BackupProtection\SmartM365-M365-BackupPolicyScope-Inventory.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-M365-BackupProtectedMailboxes-Inventory-Interactive.cmd'
        Entry = 'M365-BackupProtectedMailboxes-Interactive'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\BackupProtection\SmartM365-M365-BackupProtectedMailboxes-Inventory.ps1'
        Arguments = @('-InteractiveAuth')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-M365-BackupProtectedSitesAndDrives-Inventory-Interactive.cmd'
        Entry = 'M365-BackupProtectedSitesAndDrives-Interactive'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\BackupProtection\SmartM365-M365-BackupProtectedSitesAndDrives-Inventory.ps1'
        Arguments = @('-InteractiveAuth', '-EnableConfiguredExternalActions')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-M365UsageReports-Inventory.cmd'
        Entry = 'M365-Usage-Inventory-D180'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Usage\SmartM365-M365UserActivity-Inventory.ps1'
        Arguments = @('-Period', 'D180', '-Reports', 'All', '-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-M365UserActivity-Inventory.cmd'
        Entry = 'M365-UserActivity-Inventory-D180'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Usage\SmartM365-M365UserActivity-Inventory.ps1'
        Arguments = @('-Period', 'D180', '-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-Mailboxes-AllSources-PermissionsByUserReport.cmd'
        Entry = 'Mailboxes-PermissionsByUser-Report'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\Permissions\SmartM365-Mailboxes-AllSources-PermissionsByUserReport.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-PowerBIFabricActivity-Inventory.cmd'
        Entry = 'M365-PowerBIFabricActivity-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\PowerBI\SmartM365-PowerBIFabricActivity-Inventory.ps1'
        Arguments = @('-LookbackDays', '28')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-SecureScore-Inventory.cmd'
        Entry = 'M365-SecureScore-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Security\SecureScore\SmartM365-SecureScore-Inventory.ps1'
        Arguments = @('-EnableConfiguredExternalActions')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-SPO-Inventory.cmd'
        Entry = 'M365-SPO-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\SharePoint\SmartM365-SPO-Inventory.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-Teams-Inventory.cmd'
        Entry = 'M365-Teams-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Teams\SmartM365-Teams-Inventory.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-TeamsPhonePstnUsage-Inventory-IncludePhoneAssignments.cmd'
        Entry = 'M365-TeamsPhonePstnUsage-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Teams\SmartM365-TeamsPhonePstnUsage-Inventory.ps1'
        Arguments = @('-LookbackDays', '90', '-IncludePhoneAssignments')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-TeamsPhonePstnUsage-Inventory.cmd'
        Entry = 'M365-TeamsPhonePstnUsage-Inventory-NoPhoneAssignments'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Teams\SmartM365-TeamsPhonePstnUsage-Inventory.ps1'
        Arguments = @('-LookbackDays', '90')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-VerifiedDomains-Inventory.cmd'
        Entry = 'M365-VerifiedDomains-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Domains\SmartM365-VerifiedDomains-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-WindowsAutopilot-Inventory.cmd'
        Entry = 'Intune-WindowsAutopilot-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\Autopilot\SmartM365-WindowsAutopilot-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-WinUpdate_Status_From_Intune.cmd'
        Entry = 'Intune-WinUpdate-Status'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\WindowsUpdate\SmartM365-WinUpdate_Status_From_Intune.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-WorkplaceEvidence-Prepare-Offline.cmd'
        Entry = 'WorkplaceEvidence-Prepare-Offline'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'PreparedEvidence\SmartM365-WorkplaceEvidence-Prepare.ps1'
        Arguments = @('-Offline')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = 'PreparedEvidence'
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-WorkplaceEvidence-Prepare.cmd'
        Entry = 'WorkplaceEvidence-Prepare'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'PreparedEvidence\SmartM365-WorkplaceEvidence-Prepare.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = 'PreparedEvidence'
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Start-SmartM365-WorkplaceEvidence-Transfer.cmd'
        Entry = 'WorkplaceEvidence-Transfer'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'PreparedEvidence\SmartM365-WorkplaceEvidence-Prepare.ps1'
        Arguments = @('-TransferOnly')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = 'PreparedEvidence'
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Test-SmartM365-WorkplaceEvidence-Prepare.cmd'
        Entry = 'WorkplaceEvidence-Validate'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'PreparedEvidence\SmartM365-WorkplaceEvidence-Prepare.ps1'
        Arguments = @('-ValidateOnly')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = 'PreparedEvidence'
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Cloud\Test-SmartM365-WorkplaceEvidence-WorkforceMemory.cmd'
        Entry = 'WorkplaceEvidence-WorkforceDiagnostic'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'PreparedEvidence\SmartM365-WorkplaceEvidence-Prepare.ps1'
        Arguments = @('-WorkforceDiagnostic')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = 'PreparedEvidence'
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-AD-HealthCheck.cmd'
        Entry = 'AD-HealthCheck'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ActiveDirectoryInventory\SmartM365-ActiveDirectory-HealthCheck.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-AD-Inventory.cmd'
        Entry = 'AD-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ActiveDirectoryInventory\SmartM365-ActiveDirectory-Inventory.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },

    @{
        Launcher = 'Launchers\Jobs\Start-Exchange2016-FreeBusy-Check.cmd'
        Entry = 'Exchange2016-FreeBusy-Check'
        Action = 'Now'
        Runtime = 'WindowsPowerShell'
        Script = 'ExchangeInventory\Hybrid\FreeBusy\SmartM365-Exchange-Hybrid-FreeBusy-Inventory.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = 'ExchangeHybridFreeBusy'
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Exchange2016-Infrastructure-Inventory.cmd'
        Entry = 'Exchange2016-Infrastructure-Inventory'
        Action = 'Now'
        Runtime = 'WindowsPowerShell'
        Script = 'ExchangeInventory\OnPremises\ServersAndStorage\SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1'
        Arguments = @('-IncludeServicesHealth', '-IncludeMailboxDatabasePaths', '-IncludeHealthChecker')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = 'ExchangeInfrastructureAndReadiness'
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Exchange2016-Local-Mailboxes-Inventory.cmd'
        Entry = 'Exchange2016-Local-Mailboxes-Inventory'
        Action = 'Now'
        Runtime = 'WindowsPowerShell'
        Script = 'ExchangeInventory\OnPremises\Mailboxes\SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'
        Arguments = @('-IncludeADPermission')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = 'ExchangeLocalMailboxes\WithADPermissions'
        Include = @('SmartInventory\Common\SmartM365.IdentityDateEvidence.ps1')
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Exchange2016-Mailboxes-CalendarPermissions.cmd'
        Entry = 'Exchange2016-Mailboxes-CalendarPermissions'
        Action = 'Now'
        Runtime = 'WindowsPowerShell'
        Script = 'ExchangeInventory\OnPremises\CalendarPermissions\SmartM365-Exchange-MailboxCalendarPermissions-Inventory.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = 'ExchangeCalendarPermissions'
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Exchange2016-ProxyAddresses-Check.cmd'
        Entry = 'Exchange2016-ProxyAddresses-Check'
        Action = 'Now'
        Runtime = 'WindowsPowerShell'
        Script = 'ExchangeInventory\OnPremises\ProxyAddresses\SmartM365-Check-ProxyAddresses-Exchange.ps1'
        Arguments = @('-AllOrganizationalUnit')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = 'ExchangeProxyAddresses'
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-EXO-AcceptedDomains-Inventory.cmd'
        Entry = 'EXO-AcceptedDomains-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\AcceptedDomains\SmartM365-EXO-AcceptedDomains-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-EXO-Mailboxes-CalendarPermissions.cmd'
        Entry = 'EXO-Mailboxes-CalendarPermissions'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\CalendarPermissions\SmartM365-EXO-Mailboxes-CalPerm_Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-EXO-Mailboxes-Inventory-Fast.cmd'
        Entry = 'EXO-Mailboxes-Inventory-Fast'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\Mailboxes\SmartM365-EXO-Mailboxes-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-EXO-Mailboxes-Inventory.cmd'
        Entry = 'EXO-Mailboxes-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\Mailboxes\SmartM365-EXO-Mailboxes-Inventory.ps1'
        Arguments = @('-IncludeStats', '-ForceLiveStats', '-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-EXO-Mailboxes-Permissions.cmd'
        Entry = 'EXO-Mailboxes-Permissions'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\Mailboxes\SmartM365-EXO-Mailboxes-Inventory.ps1'
        Arguments = @('-PermissionsOnly', '-ParallelThrottle', '4', '-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-EXO-MigrationJobs-Inventory.cmd'
        Entry = 'EXO-MigrationJobs-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\Migration\SmartM365-EXO-MigJob-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-EXO-QuarantineMessages-Report.cmd'
        Entry = 'EXO-QuarantineMessages-Report'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\Quarantine\SmartM365-EXO-QuarantineMessages-Report.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Intune-AutopatchAlerts-Inventory.cmd'
        Entry = 'Intune-AutopatchAlerts-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\WindowsUpdate\AutopatchAlerts\SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.ps1'
        Arguments = @('-IncludeFeatureUpdates', '-IncludeQualityUpdates')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Intune-Configurations-Inventory.cmd'
        Entry = 'Intune-Configurations-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\WorkplaceScope\SmartM365-IntuneConfigurations-Inventory.ps1'
        Arguments = @('-EnableConfiguredExternalActions')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Intune-Devices-BIOS-Inventory.cmd'
        Entry = 'Intune-Devices-BIOS-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\Devices\SmartM365-Devices-BIOS-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Intune-Devices-Compliance-Inventory.cmd'
        Entry = 'Intune-Devices-Compliance-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\Devices\SmartM365-Devices-Compliance-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Intune-Devices-Inventory.cmd'
        Entry = 'Intune-Devices-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\Devices\SmartM365-Devices-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Intune-Devices-UpgradeEligibility.cmd'
        Entry = 'Intune-Devices-UpgradeEligibility'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\Devices\SmartM365-Devices-UpgradeEligibility.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Intune-DeviceSystem-Inventory.cmd'
        Entry = 'Intune-DeviceSystem-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\Devices\SmartM365-Device-System-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Intune-DiscoveredApps-Inventory.cmd'
        Entry = 'Intune-DiscoveredApps-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\Applications\SmartM365-Intune-DiscoveredApps-Inventory.ps1'
        Arguments = @('-DeviceDetailMode', 'All', '-FreshDeviceDetails', '-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Intune-EndpointAnalytics-Inventory.cmd'
        Entry = 'Intune-EndpointAnalytics-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\EndpointAnalytics\SmartM365-EndpointAnalytics-Inventory.ps1'
        Arguments = @('-Reports', 'All', '-IncludeStartupProcesses', '-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Intune-EndpointSecurityHealth-Inventory.cmd'
        Entry = 'Intune-EndpointSecurityHealth-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\Security\SmartM365-Intune-EndpointSecurityHealth-Inventory.ps1'
        Arguments = @('-EnableConfiguredExternalActions', '-ExportTimeoutSeconds', '1200')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Intune-EnrollmentFailures-Inventory.cmd'
        Entry = 'Intune-EnrollmentFailures-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\WorkplaceScope\SmartM365-IntuneEnrollmentFailures-Inventory.ps1'
        Arguments = @('-EnableConfiguredExternalActions')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Intune-ManagedApps-Inventory.cmd'
        Entry = 'Intune-ManagedApps-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\WorkplaceScope\SmartM365-IntuneManagedApps-Inventory.ps1'
        Arguments = @('-EnableConfiguredExternalActions')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Intune-RBAC-GroupMembers.cmd'
        Entry = 'Intune-RBAC-GroupMembers'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\RBAC\SmartM365-Intune-RBAC-GroupMembers.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Intune-Remediations-Export.cmd'
        Entry = 'Intune-Remediations-Export'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\SmartM365-Intune-ExportRemediationScripts.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },

    @{
        Launcher = 'Launchers\Jobs\Start-Intune-WindowsAutopilot-Inventory.cmd'
        Entry = 'Intune-WindowsAutopilot-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\Autopilot\SmartM365-WindowsAutopilot-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Intune-WinUpdate-Status.cmd'
        Entry = 'Intune-WinUpdate-Status'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\IntuneInventory\WindowsUpdate\SmartM365-WinUpdate_Status_From_Intune.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-M365-ActiveUsers-Inventory.cmd'
        Entry = 'M365-ActiveUsers-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Users\SmartM365-ActiveUsers-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-M365-AuthenticationMethodsRegistration-Inventory.cmd'
        Entry = 'M365-AuthenticationMethodsRegistration-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Security\AuthenticationMethods\SmartM365-AuthenticationMethodsRegistration-Inventory.ps1'
        Arguments = @('-EnableConfiguredExternalActions')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-M365-BackupPolicyScope-Inventory.cmd'
        Entry = 'M365-BackupPolicyScope-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\BackupProtection\SmartM365-M365-BackupPolicyScope-Inventory.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-M365-BackupProtectedMailboxes-Inventory.cmd'
        Entry = 'M365-BackupProtectedMailboxes-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\BackupProtection\SmartM365-M365-BackupProtectedMailboxes-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-M365-BackupProtectedSitesAndDrives-Inventory.cmd'
        Entry = 'M365-BackupProtectedSitesAndDrives-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\BackupProtection\SmartM365-M365-BackupProtectedSitesAndDrives-Inventory.ps1'
        Arguments = @('-EnableConfiguredExternalActions')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-M365-ConditionalAccess-Inventory.cmd'
        Entry = 'M365-ConditionalAccess-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Security\ConditionalAccess\SmartM365-ConditionalAccess-Inventory.ps1'
        Arguments = @('-EnableConfiguredExternalActions')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-M365-CopilotUsage-Inventory.cmd'
        Entry = 'M365-CopilotUsage-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Usage\SmartM365-CopilotUsage-Inventory.ps1'
        Arguments = @('-Period', 'D180', '-ReportVersion', 'v2', '-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-M365-EntraDevices-Inventory.cmd'
        Entry = 'M365-EntraDevices-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Devices\SmartM365-EntraDevices-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-M365-Licences-Inventory.cmd'
        Entry = 'M365-Licences-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Licensing\SmartM365-Licences-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-M365-LicensePricing-Inventory.cmd'
        Entry = 'M365-LicensePricing-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Licensing\SmartM365-LicensePricing-Inventory.ps1'
        Arguments = @('-EnableConfiguredExternalActions')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-M365-SecureScore-Inventory.cmd'
        Entry = 'M365-SecureScore-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Security\SecureScore\SmartM365-SecureScore-Inventory.ps1'
        Arguments = @('-EnableConfiguredExternalActions')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-M365-SPO-Inventory.cmd'
        Entry = 'M365-SPO-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\SharePoint\SmartM365-SPO-Inventory.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-M365-SyncHealth-Inventory.cmd'
        Entry = 'M365-SyncHealth-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\SyncHealth\SmartM365-AzureADConnect-SyncHealth-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-M365-Teams-Inventory.cmd'
        Entry = 'M365-Teams-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Teams\SmartM365-Teams-Inventory.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-M365-TeamsPhonePstnUsage-Inventory.cmd'
        Entry = 'M365-TeamsPhonePstnUsage-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Teams\SmartM365-TeamsPhonePstnUsage-Inventory.ps1'
        Arguments = @('-LookbackDays', '90', '-IncludePhoneAssignments')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-M365-Usage-Inventory.cmd'
        Entry = 'M365-Usage-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Usage\SmartM365-M365UserActivity-Inventory.ps1'
        Arguments = @('-Reports', 'All', '-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-M365-VerifiedDomains-Inventory.cmd'
        Entry = 'M365-VerifiedDomains-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\Domains\SmartM365-VerifiedDomains-Inventory.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-M365-WorkplaceScope-Inventory.cmd'
        Entry = 'M365-WorkplaceScope-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'M365Inventory\WorkplaceScope\SmartM365-WorkplaceScope-Inventory.ps1'
        Arguments = @('-EnableConfiguredExternalActions')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-Mailboxes-PermissionsByUser-Report.cmd'
        Entry = 'Mailboxes-PermissionsByUser-Report'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ExchangeInventory\Permissions\SmartM365-Mailboxes-AllSources-PermissionsByUserReport.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-SmartFinOps-Workplace-Report.cmd'
        Entry = 'SmartFinOps-Workplace-Report'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = '..\..\SmartFinOps\SmartFinOps-Workplace-Report.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-SmartFinOps-Workplace-Valuation.cmd'
        Entry = 'SmartFinOps-Workplace-Valuation'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = '..\..\SmartFinOps\SmartFinOps-Workplace-Valuation.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Jobs\Start-WorkplaceEvidence-Prepare.cmd'
        Entry = 'WorkplaceEvidence-Prepare'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'PreparedEvidence\SmartM365-WorkplaceEvidence-Prepare.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\OnPremises\Start-SmartM365-ActiveDirectory-DuplicateIdentity-Mail.cmd'
        Entry = 'AD-DuplicateIdentity-Mail'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ActiveDirectoryInventory\SmartM365-ActiveDirectory-Inventory.ps1'
        Arguments = @('-DuplicateAnalysisOnly', '-ForceSendDuplicateNotification')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\OnPremises\Start-SmartM365-ActiveDirectory-HealthCheck.cmd'
        Entry = 'AD-HealthCheck'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ActiveDirectoryInventory\SmartM365-ActiveDirectory-HealthCheck.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\OnPremises\Start-SmartM365-ActiveDirectory-Inventory.cmd'
        Entry = 'AD-Inventory'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'ActiveDirectoryInventory\SmartM365-ActiveDirectory-Inventory.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\OnPremises\Start-SmartM365-Check-ProxyAddresses-Exchange-ReadOnly.cmd'
        Entry = 'Exchange2016-ProxyAddresses-Check'
        Action = 'Now'
        Runtime = 'WindowsPowerShell'
        Script = 'ExchangeInventory\OnPremises\ProxyAddresses\SmartM365-Check-ProxyAddresses-Exchange.ps1'
        Arguments = @('-AllOrganizationalUnit')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = 'ExchangeProxyAddresses'
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\OnPremises\Start-SmartM365-Check-ProxyAddresses-Exchange-Write.cmd'
        Entry = 'Exchange2016-ProxyAddresses-Write'
        Action = 'Now'
        Runtime = 'WindowsPowerShell'
        Script = 'ExchangeInventory\OnPremises\ProxyAddresses\SmartM365-Check-ProxyAddresses-Exchange.ps1'
        Arguments = @('-AllOrganizationalUnit', '-AddMissingAddress')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = 'ExchangeProxyAddresses'
        Include = @()
        RequiresConfirmation = $True
    },
    @{
        Launcher = 'Launchers\OnPremises\Start-SmartM365-Exchange-HealthChecker-ScheduledTask-Installer.cmd'
        Entry = 'Exchange-HealthChecker-ScheduledTask-Installer'
        Action = 'Now'
        Runtime = 'WindowsPowerShell'
        Script = 'ExchangeInventory\OnPremises\ServersAndStorage\Install-SmartM365-Exchange-HealthCheckerScheduledTask.ps1'
        Arguments = @('-Interactive')
        Tenant = ''
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\OnPremises\Start-SmartM365-Exchange-Hybrid-FreeBusy-Inventory.cmd'
        Entry = 'Exchange2016-FreeBusy-Check'
        Action = 'Now'
        Runtime = 'WindowsPowerShell'
        Script = 'ExchangeInventory\Hybrid\FreeBusy\SmartM365-Exchange-Hybrid-FreeBusy-Inventory.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = 'ExchangeHybridFreeBusy'
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\OnPremises\Start-SmartM365-Exchange-Local-MailboxCalendarPermissions-Inventory.cmd'
        Entry = 'Exchange2016-Mailboxes-CalendarPermissions'
        Action = 'Now'
        Runtime = 'WindowsPowerShell'
        Script = 'ExchangeInventory\OnPremises\CalendarPermissions\SmartM365-Exchange-MailboxCalendarPermissions-Inventory.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = 'ExchangeCalendarPermissions'
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\OnPremises\Start-SmartM365-Exchange-Local-Mailboxes-Inventory-OnlyADPermissions.cmd'
        Entry = 'Exchange2016-Local-Mailboxes-OnlyADPermissions'
        Action = 'Now'
        Runtime = 'WindowsPowerShell'
        Script = 'ExchangeInventory\OnPremises\Mailboxes\SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'
        Arguments = @('-OnlyADPermission')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = 'ExchangeLocalMailboxes\OnlyADPermissions'
        Include = @('SmartInventory\Common\SmartM365.IdentityDateEvidence.ps1')
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\OnPremises\Start-SmartM365-Exchange-Local-Mailboxes-Inventory.cmd'
        Entry = 'Exchange2016-Local-Mailboxes-Inventory-MailboxesOnly'
        Action = 'Now'
        Runtime = 'WindowsPowerShell'
        Script = 'ExchangeInventory\OnPremises\Mailboxes\SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = 'ExchangeLocalMailboxes\Inventory'
        Include = @('SmartInventory\Common\SmartM365.IdentityDateEvidence.ps1')
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\OnPremises\Start-SmartM365-Exchange-Local-Mailboxes-InventoryWithADPermissions.cmd'
        Entry = 'Exchange2016-Local-Mailboxes-Inventory'
        Action = 'Now'
        Runtime = 'WindowsPowerShell'
        Script = 'ExchangeInventory\OnPremises\Mailboxes\SmartM365-Exchange-Local-Mailboxes-Inventory.ps1'
        Arguments = @('-IncludeADPermission')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = 'ExchangeLocalMailboxes\WithADPermissions'
        Include = @('SmartInventory\Common\SmartM365.IdentityDateEvidence.ps1')
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\OnPremises\Start-SmartM365-Exchange-MailboxCalendarPermissions-Inventory.cmd'
        Entry = 'Exchange2016-Mailboxes-CalendarPermissions'
        Action = 'Now'
        Runtime = 'WindowsPowerShell'
        Script = 'ExchangeInventory\OnPremises\CalendarPermissions\SmartM365-Exchange-MailboxCalendarPermissions-Inventory.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = 'ExchangeCalendarPermissions'
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\OnPremises\Start-SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.cmd'
        Entry = 'Exchange2016-Infrastructure-Inventory-Basic'
        Action = 'Now'
        Runtime = 'WindowsPowerShell'
        Script = 'ExchangeInventory\OnPremises\ServersAndStorage\SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = 'ExchangeInfrastructureAndReadiness'
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\OnPremises\Start-SmartM365-SharePoint-OnPrem-Content-Inventory.cmd'
        Entry = 'SharePointOnPrem-Content-Inventory'
        Action = 'Now'
        Runtime = 'WindowsPowerShell'
        Script = 'SharePointInventory\OnPremises\SmartM365-SharePoint-OnPrem-Content-Inventory.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = 'SharePointContent'
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\OnPremises\Start-SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.cmd'
        Entry = 'SharePointOnPrem-Infrastructure-Inventory'
        Action = 'Now'
        Runtime = 'WindowsPowerShell'
        Script = 'SharePointInventory\OnPremises\SmartM365-SharePoint-OnPrem-Infrastructure-Inventory.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $True
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = 'SharePointInfrastructure'
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Orchestrator\Restart-SmartM365-Inventory-Orchestrator.cmd'
        Entry = 'Orchestrator-Restart'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\Restart-SmartM365-Inventory-Orchestrator.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Orchestrator\Send-SmartM365-Inventory-Orchestrator-ExecutionSummary.cmd'
        Entry = 'Orchestrator-SendSummary'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Orchestrator.ps1'
        Arguments = @('-SendExecutionSummary')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Orchestrator\Start-SmartM365-Inventory-Full-Collect.cmd'
        Entry = 'Orchestrator-FullCollect'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Pipeline', 'Full', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Orchestrator\Start-SmartM365-Inventory-Orchestrator-GUI.cmd'
        Entry = 'Orchestrator-GUI'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Orchestrator-GUI.ps1'
        Arguments = @()
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Orchestrator\Start-SmartM365-Inventory-Orchestrator.cmd'
        Entry = 'Orchestrator-Start'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Orchestrator.ps1'
        Arguments = @('-Connect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Orchestrator\Start-SmartM365-Inventory-OrchestratorScheduledTask-Installer.cmd'
        Entry = 'Orchestrator-ScheduledTask-Installer'
        Action = 'Now'
        Runtime = 'PowerShell7OrWindowsPowerShell'
        Script = 'Orchestrator\Install-SmartM365-Inventory-OrchestratorScheduledTask.ps1'
        Arguments = @('-Interactive')
        Tenant = ''
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'Launchers\Orchestrator\Stop-SmartM365-Inventory-Orchestrator.cmd'
        Entry = 'Orchestrator-Stop'
        Action = 'Now'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Orchestrator.ps1'
        Arguments = @('-Stop', '-StopTimeoutSeconds', '180')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $False
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-AD-HealthCheck.cmd'
        Entry = 'AD-HealthCheck'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'AD-HealthCheck', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-AD-Inventory.cmd'
        Entry = 'AD-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'AD-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },

    @{
        Launcher = 'LaunchersByOrchestrator\Request-Exchange2016-FreeBusy-Check.cmd'
        Entry = 'Exchange2016-FreeBusy-Check'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Exchange2016-FreeBusy-Check', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Exchange2016-Infrastructure-Inventory.cmd'
        Entry = 'Exchange2016-Infrastructure-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Exchange2016-Infrastructure-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Exchange2016-Local-Mailboxes-Inventory.cmd'
        Entry = 'Exchange2016-Local-Mailboxes-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Exchange2016-Local-Mailboxes-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Exchange2016-Mailboxes-CalendarPermissions.cmd'
        Entry = 'Exchange2016-Mailboxes-CalendarPermissions'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Exchange2016-Mailboxes-CalendarPermissions', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Exchange2016-ProxyAddresses-Check.cmd'
        Entry = 'Exchange2016-ProxyAddresses-Check'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Exchange2016-ProxyAddresses-Check', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-EXO-AcceptedDomains-Inventory.cmd'
        Entry = 'EXO-AcceptedDomains-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'EXO-AcceptedDomains-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-EXO-Mailboxes-CalendarPermissions.cmd'
        Entry = 'EXO-Mailboxes-CalendarPermissions'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'EXO-Mailboxes-CalendarPermissions', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-EXO-Mailboxes-Inventory-Fast.cmd'
        Entry = 'EXO-Mailboxes-Inventory-Fast'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'EXO-Mailboxes-Inventory-Fast', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-EXO-Mailboxes-Inventory.cmd'
        Entry = 'EXO-Mailboxes-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'EXO-Mailboxes-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-EXO-Mailboxes-Permissions.cmd'
        Entry = 'EXO-Mailboxes-Permissions'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'EXO-Mailboxes-Permissions', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-EXO-MigrationJobs-Inventory.cmd'
        Entry = 'EXO-MigrationJobs-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'EXO-MigrationJobs-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-EXO-QuarantineMessages-Report.cmd'
        Entry = 'EXO-QuarantineMessages-Report'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'EXO-QuarantineMessages-Report', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Intune-AutopatchAlerts-Inventory.cmd'
        Entry = 'Intune-AutopatchAlerts-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Intune-AutopatchAlerts-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Intune-Configurations-Inventory.cmd'
        Entry = 'Intune-Configurations-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Intune-Configurations-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Intune-Devices-BIOS-Inventory.cmd'
        Entry = 'Intune-Devices-BIOS-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Intune-Devices-BIOS-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Intune-Devices-Compliance-Inventory.cmd'
        Entry = 'Intune-Devices-Compliance-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Intune-Devices-Compliance-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Intune-Devices-Inventory.cmd'
        Entry = 'Intune-Devices-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Intune-Devices-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Intune-Devices-UpgradeEligibility.cmd'
        Entry = 'Intune-Devices-UpgradeEligibility'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Intune-Devices-UpgradeEligibility', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Intune-DeviceSystem-Inventory.cmd'
        Entry = 'Intune-DeviceSystem-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Intune-DeviceSystem-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Intune-DiscoveredApps-Inventory.cmd'
        Entry = 'Intune-DiscoveredApps-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Intune-DiscoveredApps-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Intune-EndpointAnalytics-Inventory.cmd'
        Entry = 'Intune-EndpointAnalytics-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Intune-EndpointAnalytics-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Intune-EndpointSecurityHealth-Inventory.cmd'
        Entry = 'Intune-EndpointSecurityHealth-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Intune-EndpointSecurityHealth-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Intune-EnrollmentFailures-Inventory.cmd'
        Entry = 'Intune-EnrollmentFailures-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Intune-EnrollmentFailures-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Intune-ManagedApps-Inventory.cmd'
        Entry = 'Intune-ManagedApps-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Intune-ManagedApps-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Intune-RBAC-GroupMembers.cmd'
        Entry = 'Intune-RBAC-GroupMembers'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Intune-RBAC-GroupMembers', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Intune-Remediations-Export.cmd'
        Entry = 'Intune-Remediations-Export'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Intune-Remediations-Export', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },

    @{
        Launcher = 'LaunchersByOrchestrator\Request-Intune-WindowsAutopilot-Inventory.cmd'
        Entry = 'Intune-WindowsAutopilot-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Intune-WindowsAutopilot-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Intune-WinUpdate-Status.cmd'
        Entry = 'Intune-WinUpdate-Status'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Intune-WinUpdate-Status', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-M365-ActiveUsers-Inventory.cmd'
        Entry = 'M365-ActiveUsers-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'M365-ActiveUsers-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-M365-AuthenticationMethodsRegistration-Inventory.cmd'
        Entry = 'M365-AuthenticationMethodsRegistration-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'M365-AuthenticationMethodsRegistration-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-M365-BackupPolicyScope-Inventory.cmd'
        Entry = 'M365-BackupPolicyScope-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'M365-BackupPolicyScope-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-M365-BackupProtectedMailboxes-Inventory.cmd'
        Entry = 'M365-BackupProtectedMailboxes-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'M365-BackupProtectedMailboxes-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-M365-BackupProtectedSitesAndDrives-Inventory.cmd'
        Entry = 'M365-BackupProtectedSitesAndDrives-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'M365-BackupProtectedSitesAndDrives-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-M365-ConditionalAccess-Inventory.cmd'
        Entry = 'M365-ConditionalAccess-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'M365-ConditionalAccess-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-M365-CopilotUsage-Inventory.cmd'
        Entry = 'M365-CopilotUsage-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'M365-CopilotUsage-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-M365-EntraDevices-Inventory.cmd'
        Entry = 'M365-EntraDevices-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'M365-EntraDevices-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-M365-Licences-Inventory.cmd'
        Entry = 'M365-Licences-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'M365-Licences-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-M365-LicensePricing-Inventory.cmd'
        Entry = 'M365-LicensePricing-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'M365-LicensePricing-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-M365-SecureScore-Inventory.cmd'
        Entry = 'M365-SecureScore-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'M365-SecureScore-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-M365-SPO-Inventory.cmd'
        Entry = 'M365-SPO-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'M365-SPO-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-M365-SyncHealth-Inventory.cmd'
        Entry = 'M365-SyncHealth-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'M365-SyncHealth-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-M365-Teams-Inventory.cmd'
        Entry = 'M365-Teams-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'M365-Teams-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-M365-TeamsPhonePstnUsage-Inventory.cmd'
        Entry = 'M365-TeamsPhonePstnUsage-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'M365-TeamsPhonePstnUsage-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-M365-Usage-Inventory.cmd'
        Entry = 'M365-Usage-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'M365-Usage-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-M365-VerifiedDomains-Inventory.cmd'
        Entry = 'M365-VerifiedDomains-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'M365-VerifiedDomains-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-M365-WorkplaceScope-Inventory.cmd'
        Entry = 'M365-WorkplaceScope-Inventory'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'M365-WorkplaceScope-Inventory', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-Mailboxes-PermissionsByUser-Report.cmd'
        Entry = 'Mailboxes-PermissionsByUser-Report'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'Mailboxes-PermissionsByUser-Report', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-SmartFinOps-Workplace-Report.cmd'
        Entry = 'SmartFinOps-Workplace-Report'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'SmartFinOps-Workplace-Report', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-SmartFinOps-Workplace-Valuation.cmd'
        Entry = 'SmartFinOps-Workplace-Valuation'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'SmartFinOps-Workplace-Valuation', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-SmartM365-OrchestratorJob.cmd'
        Entry = 'AD-Inventory,AD-HealthCheck'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Collect', '-Job', 'AD-Inventory,AD-HealthCheck')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    },
    @{
        Launcher = 'LaunchersByOrchestrator\Request-WorkplaceEvidence-Prepare.cmd'
        Entry = 'WorkplaceEvidence-Prepare'
        Action = 'Request'
        Runtime = 'PowerShell7'
        Script = 'Orchestrator\SmartM365-Inventory-Pipeline.ps1'
        Arguments = @('-Job', 'WorkplaceEvidence-Prepare', '-Collect')
        Tenant = 'prod'
        TenantFollowsOrchestrator = $False
        ForwardsExtraArguments = $True
        WorkingDirectory = ''
        CacheName = ''
        Include = @()
        RequiresConfirmation = $False
    }
    )
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCuDC5fNuHXqlLK
# 1Lw2t3I0h6IN1E4ZozHnRSXHoY6dDqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
# s0Q4yPEDH+JoMA0GCSqGSIb3DQEBCwUAME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91
# ZGh1Yi5jb20wHhcNMjYwNzEzMDgyMjM1WhcNMjkwNzEzMDgzMjI5WjBOMR4wHAYD
# VQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRh
# Y3RAd29ya3BsYWNlY2xvdWRodWIuY29tMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEAse6XztERSyHn9DVqj8Rdv0qjc5owqvgAIGaYxBmfiQuoM48Fo4Xt
# 1ovi9brLUtf55G4XgthNPCoanxfCRRg30IVRxaDfdPXJzYmgsM5tXlsuNU49lE7E
# PJk3+jEOgSCt8NKzmVPKpNRG0NmK0a8wm12cceYZOZlSYE0+ZtT6wy5PQQjMUqIx
# XnGjt4H0nfgZZa7D4FyARKOVg/Xr9sUq5jIn3zszvg4jjeb4b0DKJtfbHukhWc2Y
# oVFgswxVBXCWIaBnfF/cjqMfK/CaToT2trVb4hG4qcQ31s1nR4keoRaOw/vyd6ap
# rEtCsT22N/Jx0dz7fIo1tVyvIaVcHdN9LW3chn0en0OKZ6Ke1OH9wf2prl4KA6Ww
# VzrAZrOlXTAItdK7D9kKO/HeJd4PZvO53oy1LdmMGLSz3OLB9e5q7yo8rfqi5Ka9
# KzM2CrSzz1yphn/H90wz7Q2pm4FIlWdcj86A/0kmhYg+5Wqqbg1drrPXu4nEBwWN
# /dzoGtKZKHTdAgMBAAGjgZYwgZMwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMD8GA1UdEQQ4MDaBHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIu
# Y29tghV3b3JrcGxhY2VjbG91ZGh1Yi5jb20wDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQUXIOOADQM78XfPAncirgCECedg9gwDQYJKoZIhvcNAQELBQADggGBADhZUB2R
# 5J/Jw030xodhEWeCQ0vnJRaiEsjOxuArQREKH3lCrQ3UsUVl292d6LnQUSTH/jF7
# rovEZ+JN2GQ/LCrXRaCuwCEGZKzlSEbtYWhfwDyj6GpIPq8Y4SeXyjdq4/rrI1bm
# iTK4Sq7EoBlGJuX6l2nfvx1tTioSr11FoDfllJR7EYawRj9hBFJ0gG0b2SuYZMgW
# gaDKefcnJDmOwcRNAZUII0ss8EeyANukWSkNN5ILZ+iKDpQgZxgDLPTiRguCyx45
# PI5wrVTjV/pR7IrtSIfq8UladlrSZJyyDn3NV2ATvIZ6wNxbTmPFcE0uMg/EYzwd
# Tek+CgXL3TxUKeldJM4YDWPimNBRhOPXzBDiOQIj6WNswt/KM1oDLnA00CNtciPN
# dn+dXlneMvTEUah9wyt8o8tkLpoBw+KN+Bq/K0O1qPtS7umi70l45pPiej+mwbwq
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIHkg8EsATjOiBGesutkwNToL2S4pLH2BbwAm3RBbtCj0MA0GCSqG
# SIb3DQEBAQUABIIBgKEodOhrDS5gCDIZL7EwyP2CIxFdd8CChDbhaG3oc8gjM9cv
# TdoXEMdfHNbNG2Bv0MFK6nrrTUEtlnh1BWN+V5BJF3Z8wLaFPF9X+pzm6kWcFLta
# kol92sUzX1GEdLbEKNayPgN/ypH0zJ4aLd71v48sq0SsHoeaAio3zQYix8GNJV14
# aRzqGjTh0JAAMQS84YoYvyzlRvpE8Iu0ggZBHagRmFp82cRLY1r5mGOYdt/cz6P7
# A1BqV7b55+923yc4f6+foD2QNwOlwboMAvSpnS2jXwhak8g0GwqieB+Po92g+o73
# 0LFHfuiBPKVXzWMPtQRCY36C/G/77/fjLlOdRx+s+YnGUMPg0vNzj+u4SYoSL87N
# p8KE0tu55ryKySheZR9dCkMZIj1VB67gF6BIuj6qQw6Ou5/T5yA+y9B2CzPxC8MB
# njz2yGLh3D0hYFxNgBfXSpR9pFigswH+J8lUrBVwokV8qHa96WskCafnXlg7N50K
# CZq8u4DOwmDuE2of+KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjcwOTQx
# NTBaMC8GCSqGSIb3DQEJBDEiBCCRrWryTK/oZVCxY2Y5Yqrw/A6/tcvBagSjsreR
# 80oydjANBgkqhkiG9w0BAQEFAASCAgCd3ompTE43LXiizwGeRHRywRm/66L5HDSZ
# RsIFqo7m7QPa3pBJwDJ1NQMf6kof25dju8Nxi2xQ1UKM/d3XzGWzWDPm/9GZgodb
# 4kXBKy25d77mA11eFbxwZXHhU5dMmBZ9dvNgIhzKgyqGX7dZ4bgC4FLdgAhKI5xK
# GMYk1PTqFHTU6rzc0hk21oD1zQcazvHTS80KFNjwhI2dBrvTerzf17FsVK1rEBh7
# +l4N6bHRdQAy2zG4yNS1MWhGNVxV241MPzuih8jQPNRL8pSJGJaSZdlAvYLRdjdf
# 7ij9BBgW3vOkfJpLO1NmYCASPL0JVbF6Tr8gOyOsFTA98xHu3HfcGUje8sgHqHhq
# VQltMMqSCN0YiBubmG0ATR4rqFaCVkgLMjnsuEbRClTVNwZGMWh54pIdOBuZy+o4
# FBRgfI6jjZ0pD55VBx0rJffSkZUZa1kVoi5uJEQzTTSzyB+Fs8QEEcbGLMlT6qSR
# RA35995g5v/4/8etRCwpGhuT88YfxMefhfDJ4SMiUpqiAtDvX7mPJCZnIZYiVy4D
# o9aGgoznorIzUREi9NBpWrUFS4SKQQgYK1a2315L0IsrcJ1Ugjckbj87VOlY6615
# nb27gMGo5wZnep5l+SZyYwTTZ4wpTZl385S2nf6lBL/JlHc7de1o4X/cjm4rmfL/
# edhF/g30wg==
# SIG # End signature block
