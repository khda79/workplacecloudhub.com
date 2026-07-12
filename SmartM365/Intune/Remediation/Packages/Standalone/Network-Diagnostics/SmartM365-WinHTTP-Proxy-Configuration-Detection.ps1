# Name: SmartM365-WinHTTP-Proxy-Configuration-Detection.ps1
# Version: 1.0
# Description: Verifies whether the WinHTTP proxy configuration is readable and consistent

$ErrorActionPreference = "Stop"

try {
    # Retrieve WinHTTP proxy configuration
    $proxyOutput = netsh winhttp show proxy 2>&1

    if ($LASTEXITCODE -ne 0) {
        Write-Output ("Failed to retrieve WinHTTP proxy configuration: " + ($proxyOutput -join " "))
        exit 1
    }

    $proxyText = ($proxyOutput -join "`n")

    # Direct access is a valid configuration
    if ($proxyText -match "Direct access|Acc.s direct|Acces direct") {
        Write-Output "WinHTTP proxy configuration is valid: direct access"
        exit 0
    }

    # Explicit proxy is also a valid configuration
    if ($proxyText -match "Proxy Server|Serveur proxy") {
        Write-Output "WinHTTP proxy configuration is valid: proxy server configured"
        exit 0
    }

    # Any other output is considered inconsistent or unexpected
    Write-Output ("WinHTTP proxy configuration is inconsistent or unexpected: " + ($proxyText -replace "`r?`n", " "))
    exit 1
}
catch {
    Write-Output ("Technical script error: " + $_.Exception.Message)
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBXpTJN1GzFFYLt
# fiPhpAjKR/CNBoRAal2PLKflzNJRpKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCAbt+zGQcgGj0Q3/WKofZMlibutcHlr8TyDZQELy1LD+TANBgkqhkiG9w0B
# AQEFAASCAYCCWoONbYtPKr8uzgPSw2l8TFHmn/yIled+DoTVc8vK3cXOsOQA+OsR
# PaBzzFO2T1jklSrXg1DcIhfvH9VCvOy0h26XRQE8/VIqZpYQv0S5OCa6nxJDDCyj
# bfO2VWeUuVe/YC1922A77uLPMdpOo8e12fbkVAgkCfzLUp2j+icodsLwN/U854y8
# 5mveXb3NUjK0qtnqOkSNgM5/9m+ewLXSYVcPF5ELIdUJ4CtBRHAsPLq0hMBlvWiF
# VVEUautTkwbLk9B4JYi3OoQ2r2vlbje2T3kmgrYPHhCA7stsgJbtVORLH1rLuE3t
# J9mX0Datu42ytG/H2b0o+4IWtxgVAwjqrvNmZTzDboiwzAzChM/SmCyKPoacOHWm
# IM5DwuX7IhKkv8yGWTjTIYN3/+WJzJ1KfLWzHR8qp0WlFN9Y+enOfHiNJfS/x9tg
# dqqxtl7rU5Y5QIllDuzWrg6OlU0vSu+lWRV8AKFJgFOkI7PIOhWVPggg1j+LcR92
# t3IJ7t8rf7U=
# SIG # End signature block
