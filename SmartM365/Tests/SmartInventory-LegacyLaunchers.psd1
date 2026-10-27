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
