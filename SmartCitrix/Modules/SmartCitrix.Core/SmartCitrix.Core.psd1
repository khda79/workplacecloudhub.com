@{
    RootModule        = 'SmartCitrix.Core.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = '4efa0169-5db6-4ccf-a36c-5d7c415c794a'
    Author            = 'https://github.com/khda79/workplacecloudhub.com'
    CompanyName       = 'WorkplaceCloudHub'
    Copyright         = '(c) WorkplaceCloudHub. All rights reserved.'
    Description       = 'Shared helpers for SmartCitrix inventory scripts.'
    PowerShellVersion = '5.1'
    FunctionsToExport = @(
        'Set-SmartCitrixCoreContext',
        'Write-SmartCitrixLog',
        'Import-SmartCitrixPowerShellComponent',
        'ConvertTo-SmartCitrixCompactJson',
        'Get-SmartCitrixObjectPropertyValue',
        'Invoke-SmartCitrixSafeInventoryBlock',
        'Invoke-SmartCitrixSdkCommand',
        'ConvertTo-SmartCitrixFlatRow',
        'ConvertTo-SmartCitrixFlatRows',
        'Export-SmartCitrixCsv',
        'New-SmartCitrixSummaryRow'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBzeFga/gU+RVqN
# iw9Y9+VjpgoggNJGi47MVylQwNCfRaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCBJEClujh6jHJ0zV6094wMeGTCkszYYxvs625Txej7VLjANBgkqhkiG9w0B
# AQEFAASCAYBGjMkwKx3czP/du3M9xAajpTQDHSxsGFVTtPwB9y8yE/LJ5yJ1CSYn
# rGIuZE/cIz3N23ouhBGcoSwLzMpoBZuw06vTgfQ6Jb7sSEYlb/8xeXTvXFVJ2zWN
# 9EjgfE29MxRfKTBJW9EW5EAd7zELfQmjHrExX7SB4dxCe7BJO0WRe+Nf54PxEg1T
# /mzdgYms7c3EdewOezXZ+YvgRBbJS6RWYBzIyZ7yB2aZiMzLtL2F3cpujWsv+vsX
# S3DPGsqor5wf4tzwqqBT6OJO2w7DTWTW4b7p5tk80IWdN3qByOGeJ7aL19EVOc87
# lXSZH8haQzYu1wx4Xb47P7h1q96Sed56MDCybXNZuCXJZX5kXJV8YUuczrjEhVyg
# RKdivYf9DoaI/RnoHPrjWX9Ne0yZFC3dexyfzuIXkKlbFPEEgRi+r2/0oa1PpfLp
# zY+3aO/PM88gpVY7Jss7ovaxw1Gfcpp/NnfdmvdvCsvZuPfST1afAXkucmj89pyD
# n2B6FJjCfs0=
# SIG # End signature block
