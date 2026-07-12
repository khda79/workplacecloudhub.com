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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA+ja60PZRWYJe4
# F88OfxHFSb3Vptpu3cNifMPpbCu81qCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDicALvn7Jtb7ylG6hk
# H5ojIC7nAtI7mm8z6jLGxZLFYDANBgkqhkiG9w0BAQEFAASCAYAvJx1wknl4MZ5C
# iibPBR/J8SHZR6daqp1Thp14cSuv6Fa2Kt+CXKhGX2LFnpoKsor78TptmvEZx1wZ
# 7UZxALR3w2PtcV1UgRIvhcebve36LaAmMdLGHjAVkwlqHbGtgw7PyHd53mLB/Jks
# GupqmYnkzrqfOPh62bYoGMweH0tX/ualUHIzywjvb2LR12qD4tvCmnX1lojtAvL3
# 5UTM0Apn9oFdI/NZQCRmCVMkERx54gu6enSW+o5nwb2s23M89NBOfgzQNzmXJr8z
# ClBDFnakLHENwmpTGVpc/EXw36jangCj4R+UydxfMykdo7EInxwV55dnupPxPi2L
# MjSD/ZYPtkT1kcRwPxLd72Hu2Pt+4GT0y08S28jhVgrM1UTGDLxgH9tEOIkQY7GM
# kfcwsxDGg0O2kxaTWgS6zJ0X6eIC95R9d5a2QHtkd+K5IwTv7G1gurw7vEJFoJLM
# /VD8/sN0X+4tWgr7e8SFwgE8A3t3J7/81kVLmCtu4e2hZQ+m03g=
# SIG # End signature block
