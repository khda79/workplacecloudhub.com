<#
.SYNOPSIS
    Detects device system information for the SmartM365 Intune inventory.

.DESCRIPTION
    This script is intended to be deployed as an Intune Platform Script on Windows devices.
    It collects Secure Boot status, BIOS information, firmware type, and the operating system
    last boot time. Results are written to stdout in a pipe-delimited format for retrieval
    through Microsoft Graph deviceManagementScripts/deviceRunStates.

    Output format:
    SecureBoot:<value>|BIOSVersion:<value>|BIOSDate:<value>|FirmwareType:<value>|LastBootUpTime:<value>

.NOTES
    Version: 1.2.0
    Author: https://github.com/khda79/workplacecloudhub.com
    Deploy via: Intune > Devices > Scripts > Platform scripts (Windows)
    Run as: System
    Run in 64-bit PowerShell: Yes
#>

$outputParts = @()

# Secure Boot
try {
    $secureBootEnabled = Confirm-SecureBootUEFI -ErrorAction Stop
    $outputParts += if ($secureBootEnabled) { 'SecureBoot:Enabled' } else { 'SecureBoot:Disabled' }
}
catch [System.PlatformNotSupportedException] {
    $outputParts += 'SecureBoot:NotSupported'
}
catch {
    $outputParts += 'SecureBoot:Error'
}

# SMBIOS BIOS version and release date
try {
    $bios = Get-CimInstance -ClassName Win32_BIOS -ErrorAction Stop
    $biosVersion = if ($bios.SMBIOSBIOSVersion) { $bios.SMBIOSBIOSVersion.Trim() } else { 'Unknown' }
    $biosDate = if ($bios.ReleaseDate) { $bios.ReleaseDate.ToString('yyyy-MM-dd') } else { 'Unknown' }
    $outputParts += "BIOSVersion:$biosVersion"
    $outputParts += "BIOSDate:$biosDate"
}
catch {
    $outputParts += 'BIOSVersion:Error'
    $outputParts += 'BIOSDate:Error'
}

# Firmware type
try {
    $computerInfo = Get-ComputerInfo -Property BiosFirmwareType -ErrorAction Stop
    $firmwareType = if ($computerInfo.BiosFirmwareType) { $computerInfo.BiosFirmwareType.ToString() } else { 'Unknown' }
    $outputParts += "FirmwareType:$firmwareType"
}
catch {
    $outputParts += 'FirmwareType:Error'
}

# Operating system last boot time, expressed in the device local time
try {
    $operatingSystem = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $lastBootUpTime = if ($operatingSystem.LastBootUpTime) {
        ([datetime]$operatingSystem.LastBootUpTime).ToString('yyyy-MM-dd HH:mm:ss')
    }
    else {
        'Unknown'
    }
    $outputParts += "LastBootUpTime:$lastBootUpTime"
}
catch {
    $outputParts += 'LastBootUpTime:Error'
}

Write-Output ($outputParts -join '|')
exit 0

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCLIzuSkk73dSh/
# NsGMdpK4kghs2+IyOKaLhcz+say+QKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCCdQuldZaa4jgF6zv5dOyUc18hnjGeZMtU+ICqEH5MuuDANBgkqhkiG9w0B
# AQEFAASCAYA3EwhvmMWp4ciJpRf6Y0idYow9DKuEOJOXjQzEQvqyK4z4NCDdVpst
# vf/MUQuWN9umWDvCv+6d1AVpxoY3w0XRs7dEuh/OG0mMxmNcn+LBYSFhKx3rmQli
# R7YXrkj6ivocShLvlGOC9Fkxh8eAxAM22AHYy9cTnTb5oiMzH7e9vfk68ALI075L
# NrjN5cf0KrTfIFrZh/zHkY2Bkj7FjYoVX/yOJFpOJFkg5vqqhj4x2kFuuEWRu5Do
# S53wrVnbirdXS1OrPGUXzY8O2K7XoxSDs5WoTUgyrXvi3z2uGjWqSWfzVS4pwCIy
# fa/05j3cE+zdnzRWdQc6QjYWIcbVkuWqd7ts4XLyY2mF3Q2XKrCO6h8t2qLJ3ufk
# +H+gs6p9HVWJHQPEXCA7ZHukfNzPWXxQh/9JA46xoO4R8ifFnuqgd7W9aoKU4I1z
# JEuxej2ZkdDc1IGZ3NLOOj9lpWwnv9At5p5pD9+U6qNab4dSEAxK0kRzdsRqEGvc
# iiBuLFVV6yo=
# SIG # End signature block
