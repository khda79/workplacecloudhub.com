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
        PrefixToRemove = ''
        PermissionRootPath = '/sites/NewMigration'
    }

    Comparison = @{
        MaxScanAgeDifferenceHours = 12
        MaxScanAgeHours = 24
        PermissionMaxScanAgeDifferenceHours = 24
        PermissionMaxScanAgeHours = 48
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCow4EZSd8h1Fze
# i4Qhdb6EV0cnYL97B1ZlCsliflxEWKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
# s0Q4yPEDH+JoMA0GCSqGSIb3DQEBCwUAME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91
# ZGh1Yi5jb20wHhcNMjYwNzEzMDgyMjM1WhcNMjkwNzEzMDgzMjI5WjBOMR4wHAYD
# VQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRh
# Y3RAd29ya3BsYWNlY2xvdWRodWIuY29tMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEAse6XztERSyHn9DVqj8Rdv0qjc5owqvgAIGaYxBmfiQuoM48Fo4Xt
# 1ovi9brLUtf55G4XgthNPCoanxfCRRg30IVRxaDfdPXJzYmgsM5tXlsuNU49lE7E
# PJk3+jEOgSCt8NKzmVPKpNRG0NmK0a8wm12cceYZOZlSYE0+ZtT6wy5PQQjMUqIx
# XnGjt4H0nfgZZa7D4FyARKOVg/Xr9sUq5jIn3zszvg4jjeb4b0DKJtfbHukhWc2Y
# oVFgswxVBXCWIaBnfF/cjqMfK/CaToT2trVb4hG4qcQ31s1nR4keoRaOw/vyd6ap
# rEtCsT22N/Jx0dz7fIo1tVyvIaVcHdN9LW3chn0en0OKZ6Ke1OH9wf2prl4KA6Ww
# VzrAZrOlXTAItdK7D9kKO/HeJd4PZvO53oy1LdmMGLSz3OLB9e5q7yo8rfqi5Ka9
# KzM2CrSzz1yphn/H90wz7Q2pm4FIlWdcj86A/0kmhYg+5Wqqbg1drrPXu4nEBwWN
# /dzoGtKZKHTdAgMBAAGjgZYwgZMwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMD8GA1UdEQQ4MDaBHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIu
# Y29tghV3b3JrcGxhY2VjbG91ZGh1Yi5jb20wDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQUXIOOADQM78XfPAncirgCECedg9gwDQYJKoZIhvcNAQELBQADggGBADhZUB2R
# 5J/Jw030xodhEWeCQ0vnJRaiEsjOxuArQREKH3lCrQ3UsUVl292d6LnQUSTH/jF7
# rovEZ+JN2GQ/LCrXRaCuwCEGZKzlSEbtYWhfwDyj6GpIPq8Y4SeXyjdq4/rrI1bm
# iTK4Sq7EoBlGJuX6l2nfvx1tTioSr11FoDfllJR7EYawRj9hBFJ0gG0b2SuYZMgW
# gaDKefcnJDmOwcRNAZUII0ss8EeyANukWSkNN5ILZ+iKDpQgZxgDLPTiRguCyx45
# PI5wrVTjV/pR7IrtSIfq8UladlrSZJyyDn3NV2ATvIZ6wNxbTmPFcE0uMg/EYzwd
# Tek+CgXL3TxUKeldJM4YDWPimNBRhOPXzBDiOQIj6WNswt/KM1oDLnA00CNtciPN
# dn+dXlneMvTEUah9wyt8o8tkLpoBw+KN+Bq/K0O1qPtS7umi70l45pPiej+mwbwq
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDnC2k94N5QVwO6M0cAygrw
# An/rJwtIjP5t5ULS9VXhczANBgkqhkiG9w0BAQEFAASCAYCvUV0NJAc/ewZv5dNA
# 3b97NhEOGxhC02hqBJDgsfUnXBw4Yg/4YuY/dXh8L/ceyd205QVoN6EP2JEVCeFR
# olHUQf0ubp67evrByUYAr7qB6YiI/W9jmhFbdDy03hWOa6AcRzsW9DpqfmsBYhPQ
# vZLKqPBfOkufdNd+Rfu363jjoPTPPkpvMGX+0ckMJ3GTN+8FWcYKjibcgSVf7Lgx
# LzwkY5hvBOADUDDpv/CO5tCmuPxk91UF6I2EgjU/QbEnS5D183iBJz5amk2tPC4A
# MHgNF+S6nYs+01l+5uI13JpEHXMhvhxHQrdyohOHy3y0nStE6djWybA1YzYOUDNG
# Xf9jHMHkU9ctEE0opy7rg8pxB7Q/zVT40eDrCctu8wQYHCXXXOs7/LLwYecrVId3
# LjxOjxikxhm9bZqrzYNV2ifIXKl5Od68D5B8q2xv5eHcOtcoJp4vvoll+xzFFYXj
# DiGaEhylhlTaQIqBneQkAapSIZ4N/LRa6gnknGguse8n13s=
# SIG # End signature block
