@{
    RootModule        = 'SmartAzure.Core.psm1'
    ModuleVersion     = '1.0.0'
    GUID              = 'a51c769d-c5b8-4e55-915d-8c5af2f4cb68'
    Author            = 'Internal automation team'
    CompanyName       = 'Internal'
    Copyright         = '(c) Internal automation team. All rights reserved.'
    Description       = 'Core SmartAzure automation tools: logging, module preflight, Azure session, CSV export, and inventory helpers.'
    FunctionsToExport = @(
        'Set-SmartAzureCoreContext',
        'Write-SmartAzureLog',
        'Import-RequiredModule',
        'Test-SmartAzureFileLocked',
        'Remove-SmartAzureOldFiles',
        'ConvertTo-CompactJson',
        'Get-ObjectPropertyValue',
        'Get-SmartAzureNestedPropertyValue',
        'Invoke-SmartAzureArmGetPaged',
        'Invoke-SafeInventoryBlock',
        'Export-SmartAzureCsv',
        'Export-SmartAzureCsvFromConvert',
        'Invoke-SmartAzurePreflight',
        'Connect-SmartAzureCloudSession'
    )
    CmdletsToExport   = @()
    VariablesToExport = '*'
    AliasesToExport   = @()
    PrivateData       = @{
        PSData = @{
            Tags       = @('SmartAzure', 'Inventory', 'Azure')
            ProjectUri = 'https://github.com/khda79/workplacecloudhub.com'
        }
    }
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAOKHM3ByFsfCcN
# 5f4chbAk4X6zpPtrVX8Z0nP+keoWBqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDr9Xl3r+2svlcMfzDXDqv/uC2DLlPH6UPYU0z5j0YO/zANBgkqhkiG9w0B
# AQEFAASCAYBBwx5VITMm4Vt7HIoJVQ1yWSlmdPjO0dK+NAcXlxUNIxFfDBPGKCHA
# lhBZBh8q61FpgHHkHAHWNe1sTRIGARqZL7ykCgJsLEK76yvQdyKUtOt6S0zppJyr
# T8yivFFxGwXZwjqVArYmIInMqnojmqs/JIKDdT5h5YR94DtizDBHZv0W+8SMa1zx
# OD9PRgs5ZaVW1X+Th1q+4xqUfOOKsUqP1jsIElF9SaBxsjOvsUPbVTIgyFcu3wUZ
# KZdTN09VHb1UJaTOKBcCowW0fZ2w4J7gT1Dl5jbVMO77Y9oYFfgM6RhVC69L+xPs
# sTDSz/bK4sgYlKkiBzIId0W6494trWotbYEvP/YtZOBAwI7Dk/I+YMmeXqbg6s05
# qDZ4PJeKSk0nL9sOdOBipYUZYEvKYYGyArRx0t3AMMYygkKvRfH9ZFtq8z7tZMmW
# 9D8ab12wKq7W8MzsiZhMYTmVgCPsroVqxxNHL8juV67TaYELp6Zxd2UIXXrCHHLb
# Hx09DGNM6bw=
# SIG # End signature block
