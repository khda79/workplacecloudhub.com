# Name: SmartM365-MDM-Enrollment-Repair-Detection.ps1
# Version: 1.0
# Description: Detects stale local MDM enrollment, device registration, and EnterpriseMgmt scheduled task signals.

$ErrorActionPreference = "Stop"

$RequireHybridJoin = $false
$ScriptName = "Detect-MDM-Enrollment-Repair"
$Version = "1.0"

function Add-Issue {
    param(
        [Parameter(Mandatory = $true)]
        [System.Collections.Generic.List[string]]$Issues,

        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    if (-not $Issues.Contains($Message)) {
        $Issues.Add($Message)
    }
}

function Get-DsregValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($Text -match ("(?m)^\s*{0}\s*:\s*(.+?)\s*$" -f [regex]::Escape($Name))) {
        return $matches[1].Trim()
    }

    return $null
}

try {
    $issues = New-Object System.Collections.Generic.List[string]
    $dsregText = (& "$env:SystemRoot\System32\dsregcmd.exe" /status 2>$null) -join "`n"

    $deviceId = Get-DsregValue -Text $dsregText -Name "DeviceId"
    $azureAdJoined = Get-DsregValue -Text $dsregText -Name "AzureAdJoined"
    $domainJoined = Get-DsregValue -Text $dsregText -Name "DomainJoined"

    if ([string]::IsNullOrWhiteSpace($deviceId)) {
        Add-Issue -Issues $issues -Message "DeviceId missing from dsregcmd status."
    }

    if ($azureAdJoined -ne "YES") {
        Add-Issue -Issues $issues -Message "AzureAdJoined is not YES."
    }

    if ($RequireHybridJoin -and $domainJoined -ne "YES") {
        Add-Issue -Issues $issues -Message "DomainJoined is not YES."
    }

    $enrollmentRoot = "HKLM:\SOFTWARE\Microsoft\Enrollments"

    if (-not (Test-Path -Path $enrollmentRoot)) {
        Add-Issue -Issues $issues -Message "MDM enrollment registry root is missing."
    }
    else {
        $enrollmentKeys = Get-ChildItem -Path $enrollmentRoot -ErrorAction SilentlyContinue

        if ($null -eq $enrollmentKeys -or $enrollmentKeys.Count -eq 0) {
            Add-Issue -Issues $issues -Message "No MDM enrollment registry records found."
        }
    }

    $enterpriseTasks = Get-ScheduledTask -ErrorAction SilentlyContinue |
        Where-Object { $_.TaskPath -like "\Microsoft\Windows\EnterpriseMgmt\*" }

    if ($null -eq $enterpriseTasks -or $enterpriseTasks.Count -eq 0) {
        Add-Issue -Issues $issues -Message "EnterpriseMgmt scheduled tasks are missing."
    }
    else {
        $enabledTasks = $enterpriseTasks | Where-Object { $_.State -ne "Disabled" }

        if ($null -eq $enabledTasks -or $enabledTasks.Count -eq 0) {
            Add-Issue -Issues $issues -Message "EnterpriseMgmt scheduled tasks are present but disabled."
        }
    }

    $imeService = Get-Service -Name "IntuneManagementExtension" -ErrorAction SilentlyContinue

    if ($null -eq $imeService) {
        Add-Issue -Issues $issues -Message "Intune Management Extension service is missing."
    }

    if ($issues.Count -gt 0) {
        Write-Output ("{0} v{1}: remediation required. {2}" -f $ScriptName, $Version, ($issues -join " | "))
        exit 1
    }

    Write-Output ("{0} v{1}: MDM enrollment and device registration signals are healthy." -f $ScriptName, $Version)
    exit 0
}
catch {
    Write-Output ("{0} v{1}: detection failed: {2}" -f $ScriptName, $Version, $_.Exception.Message)
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCByhc+7nHahA6uE
# a+1KOhsSe2XZvxv0JjPqmDsGTYnPeKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCABI3SwjMLOi32Pv+jcan/j+pJdnKtHc3mlydM3EVw0TDANBgkqhkiG9w0B
# AQEFAASCAYCnomIpdPWz2bVfdWE0dxU1/TPqLo6kuxHtTmXEEEKMuaiPTj3Qm6LL
# VtlpB2jvLKulZfUAXjm1nfSM/X0+zYaD/kBxVsvXjkKaAJUrSrbJE7LvMwETz7Jm
# lOruljSl5LD60UN2tTEnhPBXyHhZ3yY783GZFxXbsLXjsFTIYxYu/JDUKSwM6cyX
# fWQItwbn1m+B909X+/0OHFY/px38iJkwq8Aihxvi8il3m+nKwZ4vba2JzWKRiora
# oSRckSKKpKHPA3QRgbaDw7v4bHYK4UKY4bSxrMIDemJssIn9TutV9YGoJn/kVOAx
# 1okZkCPrCZIbe9hl4lIU0D7xk/GWV+YrB4hhMqF2DeUOyxU5J5IOa7BhK1HYc0fc
# 4hM/e+hLo/mXo80f2/VbwBPYYg95xGXnpt/zQUNSCUL9P4GaWaQY9t2o2aBRt27t
# cajaRxl98JmprpvK2F/irpBEO3YrlzIJUMlPpYcW4n9KBfx4qDxvDR6bPuT8xjqE
# w4f8eQph3M4=
# SIG # End signature block
