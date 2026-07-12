#
# Module manifest for SmartAzure.SharePoint
# SharePoint Online file upload via Microsoft Graph REST API
#
# Version : 1.0.0
# Author  : Internal IT
#

@{
    # Module identity
    ModuleVersion     = '1.0.2'
    GUID              = '7FDBE1D1-3761-4CF4-90FE-C3AB49134BEC'
    Author            = 'Internal IT'
    CompanyName       = 'Internal'
    Copyright         = '(c) Internal IT. All rights reserved.'
    Description       = 'SharePoint Online file upload module via Microsoft Graph REST API. Uses Invoke-RestMethod with a pre-acquired Graph Bearer token. No Microsoft.Graph SDK dependency.'

    # Requirements
    PowerShellVersion = '7.0'

    # Root module
    RootModule        = 'SmartAzure.SharePoint.psm1'

    # Exports
    FunctionsToExport = @(
        'Resolve-SmartAzureSpDriveId',
        'Invoke-SmartAzureSpFileUpload'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    # Metadata
    PrivateData = @{
        PSData = @{
            Tags        = @('SharePoint', 'Graph', 'SmartAzure', 'Upload', 'Azure')
            ProjectUri  = ''
            ReleaseNotes = @'
1.0.2
  - Invoke-SmartAzureSpFileUpload: ajout parametre ForceUpload (bool, defaut $true).
    La copie est toujours forcee par defaut; passer ForceUpload=$false pour
    retrouver le comportement skip-if-up-to-date.

1.0.1
  - Fix: Get-SpRemoteLastModified - locale-safe date parsing.
    Invoke-RestMethod auto-deserializes ISO 8601 dates to [datetime]; calling .ToString()
    without a format on fr-FR culture swaps day/month (e.g. May 13 -> Dec 5).
    Fix: if $raw is already a [datetime], use it directly instead of round-tripping via string.

1.0.0
  - Initial release.
  - Resolve-SmartAzureSpDriveId: resolves SharePoint site + document library drive ID.
  - Invoke-SmartAzureSpFileUpload: single-file upload with ForceUpload param,
    small/large dispatch (chunked resumable), timestamp preservation, Logger bridge.
  - No Microsoft.Graph SDK dependency (pure Invoke-RestMethod).
  - Compatible with any script holding a Graph Bearer token.
'@
        }
    }
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDcwg2LvSLpPfmf
# CApoF2Y/qK3Inz+BwnqqStHQdzLHNKCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCD/312Giy9jvIANrVgh
# SHzBjzTcpAQU5LxUoPzw6sQCpTANBgkqhkiG9w0BAQEFAASCAYAzpD3PYBii0YSP
# /ZkyQ7mt1AnfZwczQpJShxonYf1/Wmj6fUKykKSuq90qGRvWG9uHKMMLhUgpdtkU
# Wh1f7nGIxJqjgDxVqiLGA545azh4JVeJlIVhQNI11rwR9E2WL9Ii8FFrb8FcLrqX
# +g0CtyGZSvQqPAP6nJtIotJYUDoIlKoxrPDg8KRTUIR/RVWFe1z3FzRvH9U5FgNy
# gbVFg0zUkvFFa5puPo6qsF9S2zrMMkzFiykxDqEwEPQe/eN3Cr6VmtzO3BvinJb7
# 6D+scs4yX4pUXqnFhDW53SzkW2tMPt9VMdiUNtB1D+CADHUyYBYcX9TKCtFDbQed
# YNYxxrPlnBhUocH5Ysnek9qIQu+qCkD0Iv+Tovcsl8dW4wPGzFNA35CNIM6wnkR8
# OYa6w/TGZ9Rq/2YxvqvzWSgQMYy+hqEpFMY0Uy4gyRr8/Cz1JV6U35rzOMiRdJpA
# PVthHwYb0p6wjvLgl2Jalbl1oplEiB2XuKWprz3eoEhwQ+i2D2c=
# SIG # End signature block
