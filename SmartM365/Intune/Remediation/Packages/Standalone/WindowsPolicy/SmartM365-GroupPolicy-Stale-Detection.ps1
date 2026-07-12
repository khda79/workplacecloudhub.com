# Name: SmartM365-GroupPolicy-Stale-Detection.ps1
# Version: 1.0
$ErrorActionPreference = "Stop"
$groupPolicyStatePath = "Registry::HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\State\Machine\Extension-List\{00000000-0000-0000-0000-000000000000}"
$maxAgeDays = 7

try {
    $state = Get-ItemProperty -Path $groupPolicyStatePath -ErrorAction Stop

    if ($null -eq $state.startTimeHi -or $null -eq $state.startTimeLo) {
        Write-Output "Status=GroupPolicyTimestampMissing"
        exit 1
    }

    $fileTime = ([Int64]$state.startTimeHi -shl 32) -bor [UInt32]$state.startTimeLo
    $lastGPUpdateDate = [datetime]::FromFileTime($fileTime)
    $lastGPUpdateDays = (New-TimeSpan -Start $lastGPUpdateDate -End (Get-Date)).TotalDays

    if ($lastGPUpdateDays -gt $maxAgeDays) {
        Write-Output ("Status=GroupPolicyStale AgeDays={0:N1} MaxAgeDays={1}" -f $lastGPUpdateDays, $maxAgeDays)
        exit 1
    }

    Write-Output ("Status=GroupPolicyFresh AgeDays={0:N1}" -f $lastGPUpdateDays)
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAHw4Zhyp2hPBck
# cxSXUzSV4hbf+oPPeJF2V+lG/UrmIqCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCAWBrBUAMS0dxhLHxgUBSUw5Dku8N4Z7r4of3fORPZDgTANBgkqhkiG9w0B
# AQEFAASCAYBJvHmEOPSetA3zUticNLrBmoHbr4TS5O0rashMjdOwyz3dUtrbH7J5
# RX0wYoSIM5OhEWlSqvCe2aELLVkW2zIn5KXkxFrbyMkgnWSAVOuZUzg/nIOkB/TF
# NNXHL7bC5pNfusGGAsgveEIcTygypzNFAhdhm3PjLAw3O+KXc/HTHpxlgiLYpjOX
# GfStfgWGbhodlbOdAO/GTWO01UTYpzMaDScgtk8CXd6+qyLHb2YBCuD9Db5ipi/8
# OeYY2NPUQNvQJsHm8IHzYIjNABRhaZViD0309Ch3H5pSUseE0GaVhgNE5msXAua7
# Nmy+eSLtUVe8smV4WxfZyIN3u1q7T621sQInaob72k5JxlNo0ZxrQMv43T8KVgf7
# habBJFLQJzYWtLgFD8PTW2IF66PJGJfebDNWSQnmcZsBEbfj8Kdqd7kLjoBTeGNJ
# IIN61YRgOiEcsUXA1m/QHSdPOP4nnbW1ktLvgj9SKjfhdLUfKewsSzJDOIQRQLpn
# MBqHOYGP4K4=
# SIG # End signature block
