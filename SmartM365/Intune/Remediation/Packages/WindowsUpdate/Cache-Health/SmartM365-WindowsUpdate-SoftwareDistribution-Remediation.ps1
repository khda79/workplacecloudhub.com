# Name: SmartM365-WindowsUpdate-SoftwareDistribution-Remediation.ps1
# Version: 1.0
# Description: Safely resets the Windows Update download cache and triggers a new scan without forcing a reboot

$ErrorActionPreference = "Stop"

$LogRoot = Join-Path -Path $env:ProgramData -ChildPath 'SmartM365\IntuneRemediation\Logs\Cache-Health'
$LogPath = Join-Path -Path $LogRoot -ChildPath 'Remediate-WindowsUpdate-SoftwareDistribution.log'
if (-not (Test-Path -LiteralPath $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
function Write-SmartM365Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') $Message"
    $line | Out-File -FilePath $LogPath -Append -Encoding UTF8
    Write-Output $Message
}

function Invoke-ServiceStopSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Write-SmartM365Log "ServiceNotFound=$Name"
            return
        }

        if ($service.Status -ne "Stopped") {
            Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue
            Write-SmartM365Log "ServiceStopped=$Name"
        }
        else {
            Write-SmartM365Log "ServiceAlreadyStopped=$Name"
        }
    }
    catch {
        Write-SmartM365Log "ServiceStopFailed=$Name Message=$($_.Exception.Message)"
    }
}

function Invoke-ServiceStartSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        $service = Get-Service -Name $Name -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            Write-SmartM365Log "ServiceNotFound=$Name"
            return
        }

        if ($service.Status -ne "Running") {
            Start-Service -Name $Name -ErrorAction SilentlyContinue
            Write-SmartM365Log "ServiceStarted=$Name"
        }
        else {
            Write-SmartM365Log "ServiceAlreadyRunning=$Name"
        }
    }
    catch {
        Write-SmartM365Log "ServiceStartFailed=$Name Message=$($_.Exception.Message)"
    }
}

try {
    Write-SmartM365Log "===== Windows Update remediation started ====="

    $softwareDistributionPath = "C:\Windows\SoftwareDistribution"
    $downloadPath = Join-Path -Path $softwareDistributionPath -ChildPath "Download"

    # Stop Windows Update related services
    foreach ($serviceName in @("wuauserv", "bits", "usosvc")) {
        Invoke-ServiceStopSafe -Name $serviceName
    }

    Start-Sleep -Seconds 3

    # Clean only Download content instead of deleting the entire SoftwareDistribution folder
    if (Test-Path -Path $downloadPath) {
        Write-SmartM365Log "Cleanup=Start Path=$downloadPath"

        try {
            Get-ChildItem -Path $downloadPath -Force -ErrorAction SilentlyContinue | ForEach-Object {
                try {
                    Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
                }
                catch {
                    Write-SmartM365Log "CleanupItemSkipped=$($_.FullName)"
                }
            }

            Write-SmartM365Log "Cleanup=Completed Path=$downloadPath"
        }
        catch {
            Write-SmartM365Log "Cleanup=Partial Path=$downloadPath Message=$($_.Exception.Message)"
        }
    }
    else {
        Write-SmartM365Log "Cleanup=DownloadFolderNotFound"
    }

    # Restart services
    foreach ($serviceName in @("bits", "usosvc", "wuauserv")) {
        Invoke-ServiceStartSafe -Name $serviceName
    }

    Start-Sleep -Seconds 5

    # Trigger Windows Update scan
    $usoClientPath = Join-Path -Path $env:windir -ChildPath "System32\UsoClient.exe"

    if (Test-Path -Path $usoClientPath) {
        try {
            Start-Process -FilePath $usoClientPath -ArgumentList "StartScan" -WindowStyle Hidden -ErrorAction SilentlyContinue
            Write-SmartM365Log "WindowsUpdateScan=Triggered"
        }
        catch {
            Write-SmartM365Log "WindowsUpdateScan=Failed Message=$($_.Exception.Message)"
        }
    }
    else {
        Write-SmartM365Log "UsoClientNotFound"
    }

    Write-SmartM365Log "===== Windows Update remediation completed ====="
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAGJTfsr85nkJaQ
# o7QQLKJkHSQmVghYemetGUvaIuMIBqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCBfEmrWKvgvQZCM45rseb7V+qS3xgrDEilIo5ESy838vDANBgkqhkiG9w0B
# AQEFAASCAYAS36c8Zr5Z/ilxMINur1CgJvE2Ojwn+Q5lwwDVQwEcMfBDhbp5Dj5s
# +1HP9BAKS4qeLscbs303MxkEyaC8vbQNPOYBI5l0XZh6ciLRs6HDBjWMizNPQw5n
# n3/0yQihn0DkhjiQU9Dv9hMj9V5U3rkX9SEN6vu8Rlzpd9jgSg1Im6sPgA3HBHMq
# GhDFmoG756S57VoJC/Cizks2E9V0z+9BGWLZaHIdzuuykb5brC8AXThXtnkwa1z1
# fZ1x5La0msqXYTYy5OlXl2HG5RVSuE+5CbblnhcOLLdqD8AifYWJ8jrSWISKqiu+
# DV6bHLplub1Sxh3sL6VcXcIWODN1AY9foM3/YrqglMsqMH2ouzWR6RBtY3tapzly
# FXsQ0dgWface3JuRSu69tekRnHX3WOQrl0XxnsdlMVJSLLcrYnPs7Buj/m389Oqi
# IrP5olAgAZMyA+KEYH+Z88I6J7fV335telgpfzlBa7tz1OOpm83d4lOd4gQKQEOp
# 6Ea1TXCtj84=
# SIG # End signature block
