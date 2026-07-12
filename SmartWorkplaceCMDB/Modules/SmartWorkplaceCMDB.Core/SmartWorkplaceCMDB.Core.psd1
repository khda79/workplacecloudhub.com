@{
    RootModule        = 'SmartWorkplaceCMDB.Core.psm1'
    ModuleVersion     = '0.1.1'
    GUID              = 'fe81d6e3-5d5b-4ec0-9c8b-02f82d9bc001'
    Author            = 'WorkplaceCloudHub'
    CompanyName       = 'WorkplaceCloudHub'
    Copyright         = '(c) WorkplaceCloudHub. All rights reserved.'
    Description       = 'Core helpers for SmartWorkplaceCMDB.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Get-SmartWorkplaceCMDBProjectRoot',
        'Read-SmartWorkplaceCMDBJsonFile',
        'Resolve-SmartWorkplaceCMDBTenantPath',
        'Initialize-SmartWorkplaceCMDBTenantFolder',
        'Export-SmartWorkplaceCMDBCsv'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}



# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBJhAwPjk9jJJK8
# Azy/37Qx+wmK3+C/G8VApRIW7qhepaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCAie+tPlt1wjmzLhlfWjY5+450o3pz/4WaHqYOwhciVRjANBgkqhkiG9w0B
# AQEFAASCAYBVEtKUg/bIak94QfwqOH9CHZnc5Ii3t7pdC2OseNGhEFFkyWTUA9kq
# Hc3lEcDToG1C3k1cODSVxMmYhjOwYaD5Q4MZNMdciAhnUrjlUGiq7FX/0AkVAWEC
# DUw+/+yq4ynw6UE+CyI1HdwbgERoiDHGs0mRxZEiqjSbenQJq9iV5xUOxVXZoXj7
# QHDAQab94GFL1j5dnE+ld1mRW+xK3QCy4msipfqlO2YrdEiKBav1Nnq7d3Kl/r4x
# 71nWa9OLqvGybELKowQVtYJ+5WxIA0rPM6LKddShKpmX2n/HA4rbQ8qwWlpT5GBP
# AuEc3VNL0o8Pw0kLAtLo1+K9qxzloWjdAUR2Rh7Re/jny+Z34K6TZav4Cu/R/GS9
# xJK2XDfuaQW6sLhHA+u20KQl5IoI4f4zZkOR6MQY3ddv+qlQ6PCgBOU2Np/iKOK/
# KHfmLTgAyeaUDTOSZ1aFjwY8j7o9rMPZudq7U+ddbRUc09ET5RXjlFx8rJIJyMNw
# ghF4nesp97o=
# SIG # End signature block
