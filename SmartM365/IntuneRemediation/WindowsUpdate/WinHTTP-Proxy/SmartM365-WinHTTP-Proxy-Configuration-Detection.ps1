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
    if ($proxyText -match "Direct access") {
        Write-Output "WinHTTP proxy configuration is valid: direct access"
        exit 0
    }

    # Explicit proxy is also a valid configuration
    if ($proxyText -match "Proxy Server") {
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
