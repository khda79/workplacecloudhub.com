@{
    AppId = 'Windows11UpgradeToolkit'
    ProductName = 'SmartM365 Windows 11 Upgrade Toolkit'
    RepositoryUrl = 'https://github.com/khda79/workplacecloudhub.com'
    GitHubPathUrl = 'https://github.com/khda79/workplacecloudhub.com/tree/main/SmartM365/Devices/Windows11UpgradeToolkit'
    Branch = 'main'
    DisableEnvironmentVariable = 'W11UT_GUI_UPDATE_CHECK'
    TimeoutSeconds = 4
    CacheRoot = 'C:\ProgramData\SmartM365\Windows11UpgradeToolkit\LauncherState'
    Components = @(
        @{
            Name = 'LOT launcher GUI'
            LocalPath = 'Scripts\SmartM365-Windows11Upgrade-LotLauncher-GUI.ps1'
            RemotePath = 'SmartM365/Devices/Windows11UpgradeToolkit/Scripts/SmartM365-Windows11Upgrade-LotLauncher-GUI.ps1'
            VersionSource = 'Header'
        }
        @{
            Name = 'Endpoint script'
            LocalPath = 'Scripts\SmartM365-Invoke-Windows11UpgradeRepair.ps1'
            RemotePath = 'SmartM365/Devices/Windows11UpgradeToolkit/Scripts/SmartM365-Invoke-Windows11UpgradeRepair.ps1'
            VersionSource = 'Variable'
            VersionVariable = 'script:ScriptVersion'
        }
        @{
            Name = 'PsExec launcher'
            LocalPath = 'Scripts\SmartM365-Invoke-Windows11UpgradeRepairWithPsExec.ps1'
            RemotePath = 'SmartM365/Devices/Windows11UpgradeToolkit/Scripts/SmartM365-Invoke-Windows11UpgradeRepairWithPsExec.ps1'
            VersionSource = 'Variable'
            VersionVariable = 'script:LauncherVersion'
        }
        @{
            Name = 'PsExec worker'
            LocalPath = 'Scripts\SmartM365-Windows11Upgrade-PsExecWorker.ps1'
            RemotePath = 'SmartM365/Devices/Windows11UpgradeToolkit/Scripts/SmartM365-Windows11Upgrade-PsExecWorker.ps1'
            VersionSource = 'Header'
        }
    )
}
