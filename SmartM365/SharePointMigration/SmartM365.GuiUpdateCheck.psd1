@{
    AppId = 'SharePointMigration'
    ProductName = 'Smart SharePoint Migration'
    RepositoryUrl = 'https://github.com/khda79/workplacecloudhub.com'
    GitHubPathUrl = 'https://github.com/khda79/workplacecloudhub.com/tree/main/SmartM365/SharePointMigration'
    Branch = 'main'
    DisableEnvironmentVariable = 'SPMIG_GUI_UPDATE_CHECK'
    TimeoutSeconds = 4
    CacheRoot = 'C:\ProgramData\SmartM365\SharePointMigration\GuiUpdateCheck'
    Components = @(
        @{
            Name = 'Dashboard GUI'
            LocalPath = 'SmartM365-SharePointMigration-GUI.ps1'
            RemotePath = 'SmartM365/SharePointMigration/SmartM365-SharePointMigration-GUI.ps1'
            VersionSource = 'Header'
        }
        @{
            Name = 'Generic launcher'
            LocalPath = 'Scripts\Launchers\Generic\SmartM365-SharePointMigration-Launcher.ps1'
            RemotePath = 'SmartM365/SharePointMigration/Scripts/Launchers/Generic/SmartM365-SharePointMigration-Launcher.ps1'
            VersionSource = 'Header'
        }
        @{
            Name = 'Source file inventory'
            LocalPath = 'Scripts\Inventory\SmartM365-SharePointSource-FileInventory.ps1'
            RemotePath = 'SmartM365/SharePointMigration/Scripts/Inventory/SmartM365-SharePointSource-FileInventory.ps1'
            VersionSource = 'Header'
        }
        @{
            Name = 'Target file inventory'
            LocalPath = 'Scripts\Inventory\SmartM365-SharePointTarget-FileInventory.ps1'
            RemotePath = 'SmartM365/SharePointMigration/Scripts/Inventory/SmartM365-SharePointTarget-FileInventory.ps1'
            VersionSource = 'Header'
        }
        @{
            Name = 'Source permission inventory'
            LocalPath = 'Scripts\Inventory\SmartM365-SharePointSource-PermissionInventory.ps1'
            RemotePath = 'SmartM365/SharePointMigration/Scripts/Inventory/SmartM365-SharePointSource-PermissionInventory.ps1'
            VersionSource = 'Header'
        }
        @{
            Name = 'Target permission inventory'
            LocalPath = 'Scripts\Inventory\SmartM365-SharePointTarget-PermissionInventory.ps1'
            RemotePath = 'SmartM365/SharePointMigration/Scripts/Inventory/SmartM365-SharePointTarget-PermissionInventory.ps1'
            VersionSource = 'Header'
        }
    )
}