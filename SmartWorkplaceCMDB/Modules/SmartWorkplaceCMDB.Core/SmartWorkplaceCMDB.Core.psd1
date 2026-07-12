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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBJhAwPjk9jJJK8
# Azy/37Qx+wmK3+C/G8VApRIW7qhepaCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAie+tPlt1wjmzLhlfW
# jY5+450o3pz/4WaHqYOwhciVRjANBgkqhkiG9w0BAQEFAASCAYCjhyj+dCJtnYOH
# PQG3Z5FL4STPmUcJ/qCcAGUF8ytg7QvTXrrMPddn78CCP3PFF2wZ2JljZBmCE//y
# kbHuxbY6Eyz6EGrlHhZ90nu72v8RHrMgXXCHuzphuE/SVdPGkR+HZztbKaiahoqy
# i/q3fJvcDiIlCV1QFnBIuCZaXrCMTlYczYrAfCbNXcH5rRJ8NF0JlLhhLr9NXs5e
# b4om5DsLOZlHR8KHJY55Eh8fHWkX6oZcPBz2F+Lzj35sn8t81fYbyN99ByHtCffr
# at2EmbHF925P9X0WJuTgZPYZsz7ZiCkTVYhG8y6jS3kp3+REKEjQRG2lgA4Mc8WS
# NV7U7/MHSB3oIh333b7i0gnKvoRC3LlZfdiaMIMA95CU/qbJoPtAGo6HrC5Ov92G
# T6irlUha6jVJyF8TzjXhBtPtFLePNqUeFCOB/iR3tCnwfcDv57DRQrtFgf0EyFn2
# gONSDs3AyQ1QIAAglQsqwPxdR3k2TzJVYdWBiH9/BjMQdwsj2N4=
# SIG # End signature block
