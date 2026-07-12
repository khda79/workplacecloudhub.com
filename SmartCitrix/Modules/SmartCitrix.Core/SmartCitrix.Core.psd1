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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBzeFga/gU+RVqN
# iw9Y9+VjpgoggNJGi47MVylQwNCfRaCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBJEClujh6jHJ0zV609
# 4wMeGTCkszYYxvs625Txej7VLjANBgkqhkiG9w0BAQEFAASCAYA8fnKcScfo1DCO
# pGffiyYMCHsWepVdC3j2oj/aaapFHaEQ0n9tZOhw6KmObw3Bsu/RAPwZ/BTfeHu/
# 5R7DRFaV8VDzlXo7ivKz+SDCLXdRRjwhIkZ3U460Jpn8HndJYea3LYTkTcPeQj92
# IDJ9xMPyHT2k2QHfganNzCqKBVMqPqMFmhKtbGg6fGerf4wZf8go67SHJTBOMunX
# d3tEX0HMRItVHsuELv+BXBa2zm8OJo4AMW12WxbXSD0WHnZmjzR8keoLlHb4gZme
# yTn7twZYnG08m9AzKQbM95/n5edZH/GEJ8lt9dnAb+xK4OLb2SIW19G9HIbDwndK
# 9AHTzC/j4+MLbsrwuV1jH5FsUwKbCx4t4PRFeMi+15ycPHIS7Q+GA4GMxLHlsEXg
# A5a68pmh8LBNufN1zCt0Y1Td8/07Q48lsV4nPWm6Y37BZTCDcLnME6Ie09vr8s2x
# XHaJEri6qii6Cg3Wf0hM4XMW0JHH7fAmf8w4u2+otbi3R4ekpvA=
# SIG # End signature block
