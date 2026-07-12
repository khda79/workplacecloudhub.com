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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDcwg2LvSLpPfmf
# CApoF2Y/qK3Inz+BwnqqStHQdzLHNKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCD/312Giy9jvIANrVghSHzBjzTcpAQU5LxUoPzw6sQCpTANBgkqhkiG9w0B
# AQEFAASCAYA6SgS9tQhPPePjZB4Eo0/JC2P1JxggZbx5Cx/F7RABnRn3lN4sahkC
# GE27rw/QwhHZuJFJGpCbh15Oj+j04gnpWtTOpMFXHVD4vnp4vSRJztPwfQ6wqp3C
# G5G/VJ2Z9BB/J8qX0LjDnGGJSL4EJCxQ2eFdC3DewrR9wWcj49Ctb43Sjw98phcf
# +sFHnWTzonT6eDjBfNUGxtETzpxNaDNKzPlOeAhBT++JwqPBLzZif6ozwVZZ8CQL
# b/PctZoHjDiPpt5RYRiVCRgxuj25CURyW0d4XXHRbIRGXvI62Wvtd7zpGKweX5sY
# Af1iv18hHFmkt2KhM/Rqsar3at4VjGiGqG0xqPBS0x1WGVSlITxECU5eUKDugo0e
# Ut+QGZgPP9mYaW1vwjx7jSQtNqc86JhOSnMo5bq2uElab2wJjHZP0p5tD8XOQ8yj
# WWeJCkNDlpnyvgSJmOdzS79xLjeuNFmz7oCYGl4tczIcaOVdcTIHmmI3j+XbaXbo
# XMf70T3511M=
# SIG # End signature block
