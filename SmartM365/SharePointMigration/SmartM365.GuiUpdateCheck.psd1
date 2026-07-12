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
# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBg+fzFJM6kpk3z
# kaUdsUxhgOXSqMK98XY5eL1hao2kD6CCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCBZM6FhAh2q8u65i2+5BKV3d19cx1bM0QydUsLJS3gnLDANBgkqhkiG9w0B
# AQEFAASCAYCqSm9UJ4LRP5eYXVQVIvloOOZZV+x286t5FpmFJ6sV9DXYQZVtdZBx
# sM5UoJ9MNHdEMmSqLnBJhyVtSZRBB3OeJEAqKrWp9IaxJhPU2NUnpwrf46DarVae
# aeeW1y3Bn5BuHo1lUMcvY4yDTE7jZmEgLAkvxb27wDnZkTKN1MCc56WIzHk9vRS8
# A+M7Iu489iiwgvsv048UpIOJbEkCwvpMQ/qV/5v2E9VZx1GtCetpHEjtojRa6yIT
# H6+9DTrAsBm9TkI0sUH0lpzq4lcaumYogU4O0UbMNCe5taL1R4UqiuqGNI+kDE+C
# EGM6jOgmLsJSkrMRxRbadAOJ2y+2yg5CJY6mdFD//fYIXwTvjj+4y2UEzl4fWIYt
# 0seWju8U539WB5WqWtNnLIpHdt3m7oxEkwGWdRMLFRc4nGtG/ZRxdi2sW0ISTo7u
# C37TG0DwH/ROHIjdx9u1cWbSnm5JxXSyCWE49pKmssS/f25jcx85uXilB28ZrYKz
# uOj8fz8ifBg=
# SIG # End signature block
