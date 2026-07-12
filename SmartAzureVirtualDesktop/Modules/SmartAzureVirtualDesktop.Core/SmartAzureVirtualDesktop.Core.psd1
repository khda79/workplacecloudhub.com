@{
    RootModule        = 'SmartAzureVirtualDesktop.Core.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = '53b1de8d-6c19-4e28-8cd1-7f33a4ec39b7'
    Author            = 'https://github.com/khda79/workplacecloudhub.com'
    CompanyName       = 'WorkplaceCloudHub'
    Copyright         = '(c) WorkplaceCloudHub. All rights reserved.'
    Description       = 'Shared helpers for SmartAzureVirtualDesktop inventory scripts.'
    PowerShellVersion = '7.0'
    FunctionsToExport = @(
        'Set-SmartAvdCoreContext',
        'Write-SmartAvdLog',
        'Import-SmartAvdRequiredModule',
        'ConvertTo-SmartAvdCompactJson',
        'Get-SmartAvdObjectPropertyValue',
        'Get-SmartAvdNestedPropertyValue',
        'Invoke-SmartAvdArmGetPaged',
        'Invoke-SmartAvdSafeInventoryBlock',
        'Export-SmartAvdCsv',
        'Connect-SmartAvdCloudSession',
        'Get-SmartAvdResourceGroupNameFromId',
        'Get-SmartAvdResourceNameFromId',
        'Get-SmartAvdSubscriptionResources',
        'Get-SmartAvdResourceById',
        'Get-SmartAvdChildResources',
        'Get-SmartAvdAgeInDays',
        'ConvertTo-SmartAvdBoolString'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDks0yXjc4X7mwc
# e8ktRyouFyTosLHx6ag/8i7LU3mut6CCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCAZ8YqLNlGW6Ne8TJf1u/O8qtrDVIQMRDn+pIqOEW9iejANBgkqhkiG9w0B
# AQEFAASCAYBdcgNjVhmOTxZRJ04AlooPIfSDiGSfykT9nCedMJE+oFYt0euzCotF
# b4oHA28BRtSQiToF81lpa282vl0PQ67gy/UUb7mnNXxcQuhqG5XMP1BRiZ7h6zyQ
# Wq/FrPpmBR6ihK0ZWwDQV4VHmTrt8KTk39A0jOXMEez6ov9p5fxuzZeJ1p5qh6nd
# x8cLmFDbiQ9wtjqqte7pO4A296kDXZOYLqxDmBeO2KYqfduMlKGIIZWZHdaDhXRz
# MPVSipqqbMH0lk1SUitGrkCgGAIZAJvK88eQ/31P/nQNA37JljbcegX1j3zvssAv
# PIBcvLorP31l/0cpefIkY4HaawMshxxCPtUkP6Cp3NKQ6OjLR0VN335j1iOvcjRB
# agOOgVUErUjZGaJcJK3O25sMWPw9ogrQxGqBf5tHkhJnD6M2yF9CzWVxFu8mFYIZ
# BGe3VZfuphty41gs45z9JaNY9Fcx8pFoaGJfgwCIB0lV2Ei0e+N3ihXDSB1NpHzd
# cALDhQi/TAg=
# SIG # End signature block
