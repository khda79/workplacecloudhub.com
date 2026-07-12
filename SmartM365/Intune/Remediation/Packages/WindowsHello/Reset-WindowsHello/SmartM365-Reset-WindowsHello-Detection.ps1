
<#
.SYNOPSIS
    Version: 1.0
    Intune detection script for Windows Hello for Business with multilingual AD connectivity check.
.DESCRIPTION
    This script:
    - Detects if Windows Hello is configured (NGC folder or registry keys).
    - Checks Hybrid Join status using dsregcmd.
    - Dynamically checks AD connectivity (supports English and French outputs).
    Exit codes:
      0 = Healthy or not applicable.
      1 = Windows Hello configured but AD unreachable, or technical error.
.NOTES
    Run as SYSTEM via Intune.
#>

$ErrorActionPreference = "Stop"
$Scenario = "Reset-WindowsHello"

function Write-DetectionResult {
    param([string]$Message)
    Write-Output "$Scenario $Message"
}

# Check NGC folder
$ngcPath = "C:\\Windows\\ServiceProfiles\\LocalService\\AppData\\Local\\Microsoft\\NGC"
$ngcConfigured = ((Test-Path -LiteralPath $ngcPath) -and ((Get-ChildItem -LiteralPath $ngcPath -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1).Count -gt 0))

# Check registry for PIN presence
$pinConfigured = $false
try {
    $regPath = "HKLM:\\SOFTWARE\\Microsoft\\PassportForWork\\"  # Windows Hello for Business key
    $keys = Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue
    if ($keys) {
        $pinConfigured = $true
    }
} catch {
    $pinConfigured = $false
}

$helloConfigured = $ngcConfigured -or $pinConfigured

# Check Hybrid Join status
$dsreg = dsregcmd /status | Out-String
$domainJoined = $dsreg -match "DomainJoined.*YES"
$azureJoined = $dsreg -match "AzureAdJoined.*YES"

# Dynamically check AD connectivity (supports FR and EN output)
$adReachable = $false
try {
    $nltestResult = nltest /dsgetdc: | Out-String
    if ($nltestResult -match "DC:" -or $nltestResult -match "Domain Name" -or 
        $nltestResult -match "Nom du domaine" -or $nltestResult -match "Contrôleur de domaine") {
        $adReachable = $true
    }
} catch {
    $adReachable = $false
}
# Determine exit code
if (-not $helloConfigured) {
    Write-DetectionResult "Status=NotApplicable Reason=WindowsHelloNotConfigured"
    exit 0
}

if (-not $domainJoined) {
    Write-DetectionResult "Status=Healthy Reason=WindowsHelloConfiguredWithoutHybridJoin DomainJoined=$domainJoined AzureAdJoined=$azureJoined"
    exit 0
}

if ($adReachable) {
    Write-DetectionResult "Status=Healthy Reason=WindowsHelloConfiguredAndADReachable DomainJoined=$domainJoined AzureAdJoined=$azureJoined"
    exit 0
}

Write-DetectionResult "Status=RemediationRequired Reason=WindowsHelloConfiguredButADUnreachable DomainJoined=$domainJoined AzureAdJoined=$azureJoined"
exit 1

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB13KAY29Iljxa9
# jBqyeNhS9UphoKiKo+2fzD39GNlbEKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCAsZDFwomNX7lxsWowOc4pPbjZfIu3UT0gBLKm5hEUjOjANBgkqhkiG9w0B
# AQEFAASCAYAaAHB105fxVmQOrwPZSG9j/gAhsMzBf+ZPdCIDULgCmSZsjYeCkOWB
# vLrJeWJj+mxL7AZ2yoTDdy0+SHt3QANd6dOE0GTgdWr3hcAP6rzNmPfDGrQ6I/Ys
# 1gsRh+YTqaxxGTsy+WR50kRdE/E2U5mVO0s2ADpPhT0V00U33n9Z1KY49BqUpJQN
# oPcMsozhMyrUtFlB3ygqKlSjUoDO9IbReLskG/QBRxr+8r8pfoeE8BRVkRxlK4aL
# Oq6w4NRNAUS5Py1HJLm2vqWdTA5FzLRBnKogsZCp2N/cSdqEVzeuSMPzV3o2iPbC
# tMp0pDIJrCl62V+EDoyDlPpHwkMRDLY7avuHKdYpVz+Of5S8rmOu0KngnDc9KNW4
# rSjLQoM7HHJx/JCXhKK9lfaHhzcEni5w0u4HH11qDj5C2KXanOnJV1zbMWNhh/72
# I9LZUodfRLauWJrSyfz/k7zKYQq2YTa0iyExopgTdc6tqia18+zndMFmL6DJfSwl
# 9tQ6YiEvUT8=
# SIG # End signature block
