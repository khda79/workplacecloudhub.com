#
# Module manifest for SmartM365.SharePoint
# SharePoint Online file upload via Microsoft Graph REST API
#
# Version : 1.0.0
# Author  : Internal IT
#

@{
    # Module identity
    ModuleVersion     = '1.0.2'
    GUID              = '4C5661A1-8D39-4A69-ADE4-E8B6FB16209E'
    Author            = 'Internal IT'
    CompanyName       = 'Internal'
    Copyright         = '(c) Internal IT. All rights reserved.'
    Description       = 'SharePoint Online file upload module via Microsoft Graph REST API. Uses Invoke-RestMethod with a pre-acquired Graph Bearer token. No Microsoft.Graph SDK dependency.'

    # Requirements
    PowerShellVersion = '7.0'

    # Root module
    RootModule        = 'SmartM365.SharePoint.psm1'

    # Exports
    FunctionsToExport = @(
        'Resolve-SmartM365SpDriveId',
        'Invoke-SmartM365SpFileUpload'
    )
    CmdletsToExport   = @()
    VariablesToExport = @()
    AliasesToExport   = @()

    # Metadata
    PrivateData = @{
        PSData = @{
            Tags        = @('SharePoint', 'Graph', 'SmartM365', 'Upload', 'M365')
            ProjectUri  = ''
            ReleaseNotes = @'
1.0.2
  - Invoke-SmartM365SpFileUpload: ajout parametre ForceUpload (bool, defaut $true).
    La copie est toujours forcee par defaut; passer ForceUpload=$false pour
    retrouver le comportement skip-if-up-to-date.

1.0.1
  - Fix: Get-SpRemoteLastModified — locale-safe date parsing.
    Invoke-RestMethod auto-deserializes ISO 8601 dates to [datetime]; calling .ToString()
    without a format on fr-FR culture swaps day/month (e.g. May 13 -> Dec 5).
    Fix: if $raw is already a [datetime], use it directly instead of round-tripping via string.

1.0.0
  - Initial release.
  - Resolve-SmartM365SpDriveId: resolves SharePoint site + document library drive ID.
  - Invoke-SmartM365SpFileUpload: single-file upload with ForceUpload param,
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDpwL646yA4Ca0F
# 8hB0KR8iW/5oCdzg9OLrFpm49sLsSKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCB1T0u9nbdXyLOyt9oKdpqvSrp0nriP1ZvEs6w0+vPR7zANBgkqhkiG9w0B
# AQEFAASCAYBNptKafJ46L0Rw43TT9g1WxxyIIwm7YezwQcnV30f0tA5teNAN1F28
# CwhJAnhkG3v0uHSoGQ3ORVE+3lvrFGtcA00wYtmV5DKBs3Ks8UsK+3nyTm5cztl0
# WBy6kQRbcOF1HqvH9dPGpp+lOoE6g3KsaiUMnYoD5HGO1130aPMaWdxI0pCFL4Tm
# 8WmNs72epWrOz2WzCqLPb1Ovk8K5oPNv0UIYTcUL+afaJV9IaCOrua/6eQIuA21e
# 8q6RrAFeBwXa/+gOhG/LKQ2ztYLOf7CFzboLTKIvk8QQ30qQxvlEn7GEZ6IbL0rV
# UK5hDVdf3RE+LUO88y00GeOs9P8qtHeYRQcqyhc5eZi1OHzGQDOKGvinkKqNeolg
# oVET9hVIPWve4VqmX1aZY+idySkG0UzHVLWz1w1WkVgSNY5CSSwGtj6ByNq92AtR
# c6+3eg8A6bXNblCU9AG9hhs8oCHkHGO+njWc2TK+jyqGGq0CgQJOC7SwJWkElPJw
# kbgjz1KSwbk=
# SIG # End signature block
