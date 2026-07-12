@{
    Name = 'NewMigration'

    Source = @{
        # Supported values: SP2016, SP2019, SPO
        Type = 'SP2019'
        WebApplicationUrl = 'https://source-sharepoint.example.com'
        PermissionRootPath = '/SOURCE'
    }

    Target = @{
        # Supported values: SP2016, SP2019, SPO
        Type = 'SPO'
        SiteUrl = 'https://yourtenant.sharepoint.com/sites/NewMigration'
        TenantAdminUrl = 'https://yourtenant-admin.sharepoint.com'
        PrefixToRemove = '/sites/NewMigration'
        PermissionRootPath = '/sites/NewMigration'
    }

    Comparison = @{
        MaxScanAgeDifferenceHours = 12
        PermissionMaxScanAgeDifferenceHours = 24
        SizeToleranceBytes = 10240
        ModifiedDateToleranceMinutes = 0
        PathMappingsFile = 'migration.mapping.txt'
        SourceModifiedTimeZone = 'Auto'
        TargetModifiedTimeZone = 'Auto'
        ShareGateReplacementCharacter = '_'
        AllowDuplicateKeysForDeleteScript = $true
        EntraUsersCacheEnabled = $true
        EntraUsersCacheMaxAgeHours = 24
        EntraUsersCachePath = ''
    }

    Permissions = @{
        SourceDocumentLibrariesOnly = $false
        TargetDocumentLibrariesOnly = $false
        IncludeItemPermissions = $true
        ItemProgressInterval = 500
    }

    Output = @{
        SourceFileScans = 'scans\source\files'
        SourcePermissionScans = 'scans\source\permissions'
        TargetFileScans = 'scans\target\files'
        TargetPermissionScans = 'scans\target\permissions'
        FileComparisons = 'comparisons\files'
        PermissionComparisons = 'comparisons\permissions'
        SourceHistoryComparisons = 'comparisons\source-history'
        GeneratedOperations = 'operations\generated'
        Logs = 'logs'
    }
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA+ja60PZRWYJe4
# F88OfxHFSb3Vptpu3cNifMPpbCu81qCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDicALvn7Jtb7ylG6hkH5ojIC7nAtI7mm8z6jLGxZLFYDANBgkqhkiG9w0B
# AQEFAASCAYAfiylYuxePUZU5uuSXwoG+FLkOoUOpDHzLWLUmSMGrYC41mm0m/380
# Vzs1l2LO5Z7AKDFzrNBvxeaun2g/wtaoor5YvyKGOROJoeyisFxHbHHrRzy/SmAb
# fE5bODlvFnvhfz9C/ttAqHNu5HCMQ1mqQA+tgKrG9ks/oPEF4Ir5DibIoDh/n+iy
# LM2RwYsdLIoPgxFyIq3dAe7NSxHys2thUbPXignHa65noNe5khpbN/jB1PeDHzGl
# /jdCkjrGTC97y9MGSuwDkFrzGgydeWpvZA5p3Ub4agWX+0852KehNh3DsDZ7lZgR
# HVzsVwEwDjFWGV2c4D9npP/F7P9LQh7L9Y22qbVK8psp1sL+1C3v9h+KxNryz57u
# YeyEwEdhxEEpJSyFAJiVZHjlvnkL2VpyQoqKXAl82WpkrZHbPn9iHMsgDzWreB+2
# MAIqi16gSi4VjX6v601WL7dVFHQ+9t6Qj+Yz1xr2s48q+MYyk9MwNKCmigLAzEpG
# O+1/aH1xR8g=
# SIG # End signature block
