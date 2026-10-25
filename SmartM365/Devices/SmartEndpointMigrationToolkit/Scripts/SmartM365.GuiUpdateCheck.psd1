@{
    AppId = 'SmartEndpointMigrationToolkit'
    ProductName = 'Smart Endpoint Migration Toolkit'
    RepositoryUrl = 'https://github.com/khda79/workplacecloudhub.com'
    GitHubPathUrl = 'https://github.com/khda79/workplacecloudhub.com/tree/main/SmartM365/Devices/SmartEndpointMigrationToolkit'
    Branch = 'main'
    DisableEnvironmentVariable = 'SEMT_GUI_UPDATE_CHECK'
    TimeoutSeconds = 4
    CacheRoot = 'C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\LauncherState\GuiUpdateCheck'
    Components = @(
        @{
            Name = 'GUI'
            LocalPath = 'Scripts\SmartM365-EndpointMigration-GUI.ps1'
            RemotePath = 'SmartM365/Devices/SmartEndpointMigrationToolkit/Scripts/SmartM365-EndpointMigration-GUI.ps1'
            VersionSource = 'Header'
        }
        @{
            Name = 'Engine (PsExec orchestrator)'
            LocalPath = 'Scripts\SmartM365-Invoke-EndpointMigrationWithPsExec.ps1'
            RemotePath = 'SmartM365/Devices/SmartEndpointMigrationToolkit/Scripts/SmartM365-Invoke-EndpointMigrationWithPsExec.ps1'
            VersionSource = 'Header'
        }
        @{
            Name = 'Worker'
            LocalPath = 'Scripts\SmartM365-EndpointMigration-PsExecWorker.ps1'
            RemotePath = 'SmartM365/Devices/SmartEndpointMigrationToolkit/Scripts/SmartM365-EndpointMigration-PsExecWorker.ps1'
            VersionSource = 'Header'
        }
        @{
            Name = 'Hybrid Join endpoint'
            LocalPath = 'Campaigns\HybridJoin\Scripts\SmartM365-Invoke-EndpointMigration-HybridJoin.ps1'
            RemotePath = 'SmartM365/Devices/SmartEndpointMigrationToolkit/Campaigns/HybridJoin/Scripts/SmartM365-Invoke-EndpointMigration-HybridJoin.ps1'
            VersionSource = 'Header'
        }
        @{
            Name = 'Windows 11 endpoint'
            LocalPath = 'Campaigns\Windows11Upgrade\Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'
            RemotePath = 'SmartM365/Devices/SmartEndpointMigrationToolkit/Campaigns/Windows11Upgrade/Scripts/SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'
            VersionSource = 'Header'
        }
    )
}
