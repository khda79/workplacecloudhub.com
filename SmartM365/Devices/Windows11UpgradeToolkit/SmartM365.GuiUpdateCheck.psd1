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

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBWAlXhrrDDzr1N
# 9R+xsscN737OyWn+Z08mXYDb4CxgCaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDJ4DSWwuFNVEF+qLaahh657dJbDLhb+w2FMVMfSTMlKzANBgkqhkiG9w0B
# AQEFAASCAYAQhSVmO5/hZCxbaU6LRr68++qiDUtf4P0hN3ZGqq759221IHy0ige/
# 6G+oaDtT4ALcbxtBLE+++NE3hB6lQSB7o0KRMVsQeyMCHjisZhswGWzT60dxAkjU
# MGC4TPQY8rEAhQWIYnTEaeFcLmdBGA4Ifg10eQfk3l8mk1DPGXFesPEt1zUc7sM5
# AeKlV1CDtWBL6vhThIeeP5LOnpxtARfwTfF3iExL21Uve8djKIKi5lWTsn+g45QU
# /Gt54g/D6mxmDTEOGlDQYy8KZq8KOwwiuNtCkGA9474MdxVKs9adbF3U6qgkxs7v
# gtheWmBSvJYkQeq6LEL6CISiPreRSKemNmElPnmVsLHwCNx3PIAhYcbld48QyUB0
# 0jWg7hwR8e6kkKqp2+jPZcVF1621hC1ORn1l4mXJqrCJXzo2EAQJYdZlX1ttnFXf
# iW5lfJ6Oyj77gx4Zrmf0x1UV7jOPccWnIP+utY41H2bGEub9QFnYeye3UAh2+qG/
# oL9eZcCGXNA=
# SIG # End signature block
