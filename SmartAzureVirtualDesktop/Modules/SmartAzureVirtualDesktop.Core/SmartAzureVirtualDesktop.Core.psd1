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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDks0yXjc4X7mwc
# e8ktRyouFyTosLHx6ag/8i7LU3mut6CCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAZ8YqLNlGW6Ne8TJf1
# u/O8qtrDVIQMRDn+pIqOEW9iejANBgkqhkiG9w0BAQEFAASCAYBmyEFSPyI3RPHH
# qrVPcHf/7HqOXNKE5060EYdjSqkDvcljYpoyCR2obBO+wMLdCl5h5en1PIVT2IAd
# JSciU7zji4tjXTzOU5mh6uDaDqOG5nyC4WG86oEdemV9yNH/FBo0tAgtXbPASOTP
# sbr2GH4NXr74rCp2nLOQQaur2d2BPk2Dw4aPvvUZTuh2QUS1adHmL12KbtHtK7RF
# s8GSVYNBsGc8lmSj6JyRL1CwP8FitpQGcoj8uqx3V+9qKywiS8fx2zYaZkaG/jT5
# Ydym1utuXTrFPTn2IGV/DP2W+1EJrfP745Xm6o0m6hI7YQpy93Uh1jjCzvlkD7VE
# BVqBpBxVUndaj086Vc7+OwunHSe0Ix2kbtBsyo/4p11orlp7zb/6mcNmiZpOC81J
# EN8JpD5fciKhS4JQcdA6aJCsLjkOpLKCl6DeLw1xqD8LdTeM+Y55Z8kpFpRQPE8s
# 3gdkIN5aZfNDWBlKWeoEPbfjQvsUNjw8DOoVejOHbXE0BohEOt0=
# SIG # End signature block
