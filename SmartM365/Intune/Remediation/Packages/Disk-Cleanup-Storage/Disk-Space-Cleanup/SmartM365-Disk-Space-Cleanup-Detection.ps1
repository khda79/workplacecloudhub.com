<#
.SYNOPSIS
    Detects whether the system drive is below the managed free-space threshold.

.VERSION
    1.6
#>
# Name: SmartM365-Disk-Space-Cleanup-Detection.ps1
# Version: 1.6
# Description: Detects whether the system drive is below the managed free-space threshold.

$ErrorActionPreference = "Stop"

$ScriptName = "SmartM365-Disk-Space-Cleanup-Detection"
$Version = "1.6"
$MinimumFreeSpaceGB = 50
$Windows10Only = $true

function ConvertTo-SingleLineValue {
    param([AllowNull()][object]$Value)

    if ($null -eq $Value) {
        return "Unknown"
    }

    $text = [string]$Value
    $text = $text -replace '[\r\n\t]+', ' '
    $text = $text -replace '\s{2,}', ' '
    return $text.Trim()
}

function Write-IntuneResult {
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Values)

    $parts = New-Object System.Collections.Generic.List[string]

    foreach ($key in $Values.Keys) {
        $parts.Add(("{0}={1}" -f $key, (ConvertTo-SingleLineValue -Value $Values[$key])))
    }

    Write-Output ($parts -join " ")
}

function Get-SystemDriveFreeSpaceGB {
    try {
        $systemDrive = $env:SystemDrive

        if ([string]::IsNullOrWhiteSpace($systemDrive)) {
            return $null
        }

        $drive = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$systemDrive'" -ErrorAction Stop

        if ($null -eq $drive -or $drive.DriveType -ne 3) {
            return $null
        }

        return [math]::Round(($drive.FreeSpace / 1GB), 2)
    }
    catch {
        return $null
    }
}

try {
    $operatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop

    if ($Windows10Only -and $operatingSystem.Caption -notmatch "Windows 10") {
        Write-IntuneResult -Values ([ordered]@{ Status = "NotApplicable"; Reason = "NotWindows10"; Script = $ScriptName; Version = $Version })
        exit 0
    }

    $freeSpaceGB = Get-SystemDriveFreeSpaceGB

    if ($null -eq $freeSpaceGB) {
        Write-IntuneResult -Values ([ordered]@{ Status = "RemediationRequired"; Reason = "FreeSpaceUnknown"; RequiredFreeSpaceGB = $MinimumFreeSpaceGB; Script = $ScriptName; Version = $Version })
        exit 1
    }

    if ($freeSpaceGB -ge $MinimumFreeSpaceGB) {
        Write-IntuneResult -Values ([ordered]@{ Status = "Ready"; FreeSpaceGB = $freeSpaceGB; RequiredFreeSpaceGB = $MinimumFreeSpaceGB; Script = $ScriptName; Version = $Version })
        exit 0
    }

    Write-IntuneResult -Values ([ordered]@{ Status = "RemediationRequired"; FreeSpaceGB = $freeSpaceGB; RequiredFreeSpaceGB = $MinimumFreeSpaceGB; Script = $ScriptName; Version = $Version })
    exit 1
}
catch {
    Write-IntuneResult -Values ([ordered]@{ Status = "Error"; Script = $ScriptName; Version = $Version; Message = $_.Exception.Message })
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCsymxq798Gkey+
# jxRX6PS1FFnpkW4wy3P+uyo3I3mkQqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCC/0X4WHJxaNJ8LJGVCnGE1DW6VKVpFh2U130gjoz7QkTANBgkqhkiG9w0B
# AQEFAASCAYCFZOyKJMfXmHxjKeRpISkYC4jXSOB6Lg4IfBc5PZxPZZudcNcBXUFa
# GyfjQezLAyhfCqOVnkZT4oa6UyhA7r86hsi9hAPxNKFbsVjF3D+1yz3826nZxEg5
# /oXyY6MdroHhaDyB0Cij3fFHuNtHVXBqmooTEv0xvYbcoURUvULAEY3LBEOYal5F
# z9gxovFFq2KclGOEQySutsbM1v6gmxd7YzPQJPgRCrJHBavyiAiWvCacnL5K0VkK
# suLs7zmg8VJ0ps/+CrH5+2rQmjwIhZT0m26HwsLUXyvOw4pXpfTxUlrgnat4BPbf
# 7El/xlq2lWcEpM104bYutxpS04JIoMDlGrUzM1dnVL7xLTAHrcEStVLW9AwiuRBW
# oRo8lvmcUnpOGnUvXilxdzIkTkUgLf/OPfSAkBdc0vg+vt20cxlYHbucViTy3EhI
# c89g//72HW8+Rw78w5X2ZC5fmdbl2at+QW+K63vHrKTaaeyhpn7aLcLiSXwsPsGB
# ANZSa7MDnqI=
# SIG # End signature block
