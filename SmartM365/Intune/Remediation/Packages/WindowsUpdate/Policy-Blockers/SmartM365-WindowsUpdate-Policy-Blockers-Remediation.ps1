# Name: SmartM365-WindowsUpdate-Policy-Blockers-Remediation.ps1
# Version: 1.0
# Description: Removes WSUS, Windows Update, and WUfB policy values that block cloud-managed update flows.

$ErrorActionPreference = "Stop"

$Scenario = "Policy-Blockers"
$LogRoot = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path $LogRoot "$Scenario-Remediation.log"
$WuPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate"
$WuAuPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"

function Write-SmartM365Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Scenario, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8
}

function Invoke-RegistryValueRemovalSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (-not (Test-Path -Path $Path)) {
        return
    }

    $item = Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue

    if ($item -and $item.PSObject.Properties.Name -contains $Name) {
        Remove-ItemProperty -Path $Path -Name $Name -Force -ErrorAction SilentlyContinue
        Write-SmartM365Log "RegistryValueRemoved=$Path\$Name"
    }
}

function Invoke-ServiceRestartSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

    if ($service) {
        Set-Service -Name $Name -StartupType Manual -ErrorAction SilentlyContinue

        if ($service.Status -eq "Running") {
            Restart-Service -Name $Name -Force -ErrorAction SilentlyContinue
        }
        else {
            Start-Service -Name $Name -ErrorAction SilentlyContinue
        }

        Write-SmartM365Log "ServiceRefreshed=$Name"
    }
}

function Invoke-UsoClient {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Action
    )

    $uso = Join-Path $env:SystemRoot "System32\UsoClient.exe"

    if (Test-Path -LiteralPath $uso) {
        Start-Process -FilePath $uso -ArgumentList $Action -WindowStyle Hidden -ErrorAction SilentlyContinue
        Write-SmartM365Log "UsoClient=$Action Status=Triggered"
    }
}

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-SmartM365Log "RemediationStarted"

    foreach ($name in @("WUServer", "WUStatusServer", "UpdateServiceUrlAlternate", "DoNotConnectToWindowsUpdateInternetLocations", "DisableWindowsUpdateAccess", "SetDisableUXWUAccess")) {
        Invoke-RegistryValueRemovalSafe -Path $WuPolicyPath -Name $name
    }

    foreach ($name in @("UseWUServer", "NoAutoUpdate", "AUOptions")) {
        Invoke-RegistryValueRemovalSafe -Path $WuAuPolicyPath -Name $name
    }

    foreach ($service in @("wuauserv", "bits", "dosvc")) {
        Invoke-ServiceRestartSafe -Name $service
    }

    Invoke-UsoClient -Action "RefreshSettings"
    Invoke-UsoClient -Action "StartScan"

    Write-SmartM365Log "RemediationCompleted"
    exit 0
}
catch {
    Write-SmartM365Log "RemediationFailed Message=$($_.Exception.Message)"
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBVCO5dYNQjNnvY
# IhMFvperzU5x5G63VyWZiaxhZD9IRKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCCjeXWMGVCUjD+GNhEeMxmA77vDjwItc1dxGKWgPIAciTANBgkqhkiG9w0B
# AQEFAASCAYAO5SfcjEvXf3gZnkOQyLJeJZDYJJqYXL03yp7WnCAe6xXAoM8PGOCJ
# M28XThnIBk+V8wxMaq0IbpR2dhaxu5hHn75im0PqiJaWjgyDTQN3QuJaz9Pd6h5k
# SSBA7avHLhBM2Pmc8moOdXr89tANj1XaXVkJ/kAh74Fm53U4BpkqIzNbT9QoIFcy
# ldJy4p/PURepn0nv1CPdB9s2Rc32mgHBPhMWgvKEqlcb9nBk8ZQER4v108aSDtIb
# bFze1nvikGB8DbSFQYrT72pIbg/gsjv17ThgBsfnHyuNZma4L6tZLt6JJIpxTWqr
# 8Fcy8mKyt5oK2Rybxtqu0BNf1+8aQ0RbDMPKXIvw5q9sGSR307rllosdwZKQiJu2
# S4z5Us9Z6oNwtsjMFeC5mXvN8jkhQJCZfj2vS2QzZlL8n2GBIjtW2jw0oA6Xr//N
# qDS2iL7sOSekIrJQ0PaCd/rrR/WZNeIR6lFfRbQHt0psX2DITUF/5lEOwFJRrVjJ
# 99foD+WcO/Y=
# SIG # End signature block
