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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBg+fzFJM6kpk3z
# kaUdsUxhgOXSqMK98XY5eL1hao2kD6CCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBZM6FhAh2q8u65i2+5
# BKV3d19cx1bM0QydUsLJS3gnLDANBgkqhkiG9w0BAQEFAASCAYCwJhcPdD8Jtgo3
# NIiCBp9PJSjgcPdzFAUyOlnoLYgT/s45qU0R6Zdv6CigbjJiAfwT6Z6QiD8oQVUY
# bI8UDEgVEaQhAGrq8nL2NYeUtbglNaestVAdlVAIsc17YMQAbgJ+XrHYfIIWAL2b
# 18D3RhkgwWeA7lz9mdCZqTuGxtPp2QTcBj7jtZq+XgxaXrwcWSRCVKTe7w7T/2wE
# BFJpIf/ONsqVWqEdlO15QUGlE7XS6FlBjMX0mzRDm/PgkZpqXDXQuV529YwmKeRA
# qgYknq1I6zDezmPjMbk1BoJvtSBkZnZn7Pi4UilLO38q8sXBMsIDrd1nfMNQ7EGy
# RoDeEy173+BiK+Rl9GngrD0zRknFv15Pj4I1q0gJPXkfA/ka7lMfcNwgZLGB8COx
# QoLIODhvzVcWenzje05Ev/kyDSi6AYwJ90B10m57RM9asFwDohplFF0vpkF6n6++
# kLutW/bPWc5sJHgyGoi3SkWKokbYNpdwmJUm4lIVFbI3IJ/yLHg=
# SIG # End signature block
