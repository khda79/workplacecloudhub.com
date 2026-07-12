# Name: SmartM365-Upgrade-Staging-Health-Remediation.ps1
# Version: 1.0
# Description: Removes stale Windows upgrade staging folders only when no recent setup activity is detected.

$ErrorActionPreference = "Stop"

$Scenario = "Upgrade-Staging-Health"
$RecentSetupActivityHours = 6
$LogRoot = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path $LogRoot "$Scenario-Remediation.log"
$UpgradePaths = @('C:\$WINDOWS.~BT', 'C:\$WINDOWS.~WS')
$SetupIndicators = @(
    "C:\Windows\Panther\setupact.log",
    "C:\Windows\Panther\setuperr.log",
    'C:\$WINDOWS.~BT\Sources\Panther\setupact.log',
    'C:\$WINDOWS.~BT\Sources\Panther\setuperr.log'
)

function Write-SmartM365Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Scenario, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8
}

function Invoke-ServiceStopSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (Get-Service -Name $Name -ErrorAction SilentlyContinue) {
        Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue
        Write-SmartM365Log "ServiceStopRequested=$Name"
    }
}

function Invoke-ServiceStartSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if (Get-Service -Name $Name -ErrorAction SilentlyContinue) {
        Start-Service -Name $Name -ErrorAction SilentlyContinue
        Write-SmartM365Log "ServiceStartRequested=$Name"
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

    $recentSetupActivity = $false

    foreach ($indicator in $SetupIndicators) {
        if (Test-Path -LiteralPath $indicator) {
            $lastWrite = (Get-Item -LiteralPath $indicator -ErrorAction SilentlyContinue).LastWriteTime

            if ($lastWrite -and ((Get-Date) - $lastWrite).TotalHours -le $RecentSetupActivityHours) {
                $recentSetupActivity = $true
            }
        }
    }

    if ($recentSetupActivity) {
        Write-SmartM365Log "RecentSetupActivityDetected=True CleanupSkipped=True"
        exit 0
    }

    foreach ($service in @("bits", "wuauserv", "dosvc")) {
        Invoke-ServiceStopSafe -Name $service
    }

    foreach ($path in $UpgradePaths) {
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Recurse -Force -ErrorAction SilentlyContinue
            Write-SmartM365Log "UpgradeFolderRemoved=$path"
        }
        else {
            Write-SmartM365Log "UpgradeFolderNotFound=$path"
        }
    }

    foreach ($service in @("dosvc", "wuauserv", "bits")) {
        Invoke-ServiceStartSafe -Name $service
    }

    Invoke-UsoClient -Action "RefreshSettings"
    Invoke-UsoClient -Action "StartScan"
    Invoke-UsoClient -Action "StartDownload"

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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAjTCa6vyBNDIVi
# 8pDClkrJ/+9jAlYLXvnPSWE6hHCPKKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCBkwfSkH3eZc/pdqWX1ehFbCCEPBRGL/27vQHuS7vVCszANBgkqhkiG9w0B
# AQEFAASCAYBGpE3PFgopn5u4c9jEGqNvmXumr6dwdxfPoUuVu2UVwg6d7Zepju0O
# qpvNKlIVaZw4CWtE+8qgTMKwq4al6AeB+e749+I/CvzdWIdVdxQf3XsxNf0sfNf/
# ffVpNFECqZTK1PsLFOBJKqcaodkEGRKe8wczwNJIUCBVQ9CONWwWg+5I4hiaAfMx
# N1/poRcEKDCcfQIYadPY6sRa9dXE6mXKjOhc/1ylWvcB45/drDSVojt9FkzckYwm
# /x3kKTaDVkAisZDZKK4oIk2vPKhYX7qlEdD2BcsCbED7KxZMBCSkaU5yPj0bxXEv
# gbPh3KuJj17uhqKbdax8hp2v6Frc+TxzgyT8YavzxtEaTEqmBZ5krCmWbLZh+cFb
# wegH61NXVDszK6bO8VcCFBmzQZpvn2ibp4nNpS6g/tFmJfiiL5hW71UrX5rLkJkK
# y8pGq83ZKODefZ+0cwyqdQqYaYDL5fSJAlKHUOdKLrxV78H6JS2CeYRqUOSDQ2yL
# 9WZxUGlRayw=
# SIG # End signature block
