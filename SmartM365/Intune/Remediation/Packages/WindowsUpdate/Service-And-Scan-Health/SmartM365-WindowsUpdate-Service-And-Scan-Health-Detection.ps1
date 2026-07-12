<#
    Name: SmartM365-WindowsUpdate-Service-And-Scan-Health-Detection.ps1
    Version: 1.0
    Description: Consolidated detection for Windows Update service health and recent scan activity.
#>

[CmdletBinding()]
param(
    [int]$MaxHoursSinceLastWindowsUpdateEvent = 24
)

$ErrorActionPreference = "Stop"
$ScriptName = "WindowsUpdate-Service-And-Scan-Health"
$RequiredServices = @("bits", "wuauserv", "dosvc", "cryptsvc", "UsoSvc")
$WindowsUpdateLogName = "Microsoft-Windows-WindowsUpdateClient/Operational"

try {
    $issues = New-Object System.Collections.Generic.List[string]
    $healthyStates = New-Object System.Collections.Generic.List[string]

    foreach ($serviceName in $RequiredServices) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            $issues.Add("ServiceMissing=$serviceName")
            continue
        }

        $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$serviceName'" -ErrorAction SilentlyContinue

        if ($null -eq $serviceCim) {
            $issues.Add("ServiceCimUnavailable=$serviceName")
            continue
        }

        if ($serviceCim.StartMode -eq "Disabled") {
            $issues.Add("ServiceDisabled=$serviceName")
            continue
        }

        if ($service.Status -in @("StopPending", "PausePending", "Paused")) {
            $issues.Add("ServiceUnhealthy=$serviceName Status=$($service.Status)")
            continue
        }

        $healthyStates.Add("$serviceName=$($service.Status),StartupType=$($serviceCim.StartMode)")
    }

    $logInfo = Get-WinEvent -ListLog $WindowsUpdateLogName -ErrorAction Stop
    if (-not $logInfo.IsEnabled) {
        $issues.Add("WindowsUpdateEventLogDisabled")
    }
    else {
        $lastEvent = Get-WinEvent -LogName $WindowsUpdateLogName -MaxEvents 1 -ErrorAction SilentlyContinue

        if ($null -eq $lastEvent) {
            $issues.Add("WindowsUpdateEventLogEmpty")
        }
        else {
            $ageHours = (New-TimeSpan -Start $lastEvent.TimeCreated -End (Get-Date)).TotalHours

            if ($ageHours -ge $MaxHoursSinceLastWindowsUpdateEvent) {
                $issues.Add(("LastWUEventAgeHours={0:N1}" -f $ageHours))
            }
            else {
                $healthyStates.Add(("LastWUEventAgeHours={0:N1}" -f $ageHours))
            }
        }
    }

    if ($issues.Count -gt 0) {
        Write-Output ("{0}: remediation required. {1}" -f $ScriptName, ($issues -join "; "))
        exit 1
    }

    Write-Output ("{0}: healthy. {1}" -f $ScriptName, ($healthyStates -join "; "))
    exit 0
}
catch {
    Write-Output ("{0}: detection failed. Message={1}" -f $ScriptName, $_.Exception.Message)
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAvH7cgTLOFfzCJ
# TUATl6dMKJXrCX7kItE4fA5QdnarbqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCCFKmii13iFAnulvXZgyp2+hQ87Cxmf4/17Jb5ZeUbItTANBgkqhkiG9w0B
# AQEFAASCAYA3RdoFiFVzUEN9HLDHjhHeE9DUpvtZH/NqzA5M3qpIhmtnzqNPNtD1
# OnG6+0/wtxABt3MWpIikoSKxqurBTXihznafHxQS7qKFOHwe0bmC214c8L4fUvzm
# jChrKXsbrWL+9HZmdkmbvkdgpVkIQvN4/82Df2D392dwUfYMiMDhHT2fD5nWsiBx
# rB66+3qm0B3W5mkPvBFAGGKW67eah7rm83M/Sl8qNz4Misi/L/2GNbKyS2X02daX
# FV1iSc9MNsR8wDvye8wMW1/grk8w3tS9AQykR9DngdPXTgcjH9YBsy7mRfW33T0r
# ryRwUiIiJCjQ6xpFMQUr9gd1JM5Yo4dxaiRS2yiwgRmKHkSOXZum/qZ1qVh7S7wM
# ruCbg/7WzHPvEGDNp/EWqeaRBURwEgAyl/Pp6guCqDc3+TZsaPDtnXKTujVws7vj
# J1JytdSHBjlI8CGJtjj+7M7H6lb6rQkMNcZX6c1vujumMiwPFEVBGk9vcYrr5C7W
# z5YFQC8rssQ=
# SIG # End signature block
