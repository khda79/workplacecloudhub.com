# Name: SmartM365-WindowsUpdate-Cache-Health-Detection.ps1
# Version: 1.0
# Description: Detects an inconsistent, invalid, or abnormally large Windows Update cache

$ErrorActionPreference = "Stop"

try {
    $dataStoreLogPath = "C:\Windows\SoftwareDistribution\DataStore\Logs\edb.log"
    $downloadCachePath = "C:\Windows\SoftwareDistribution\Download"
    $maxDownloadCacheSizeBytes = 5GB
    $recentActivityMinutes = 60

    $issues = New-Object System.Collections.Generic.List[string]

    # Check DataStore edb.log only if it exists.
    # A zero-byte edb.log can indicate a damaged Windows Update DataStore transaction log.
    if (Test-Path -Path $dataStoreLogPath -PathType Leaf) {
        $dataStoreLog = Get-Item -Path $dataStoreLogPath -ErrorAction Stop

        if ($dataStoreLog.Length -eq 0) {
            $issues.Add("DataStore edb.log is empty")
        }
    }

    # Check Download cache size and recent activity
    if (Test-Path -Path $downloadCachePath) {
        $downloadFiles = Get-ChildItem -Path $downloadCachePath -Recurse -File -ErrorAction SilentlyContinue

        if ($null -ne $downloadFiles -and $downloadFiles.Count -gt 0) {
            $downloadCacheSizeBytes = ($downloadFiles | Measure-Object -Property Length -Sum).Sum

            if ($null -eq $downloadCacheSizeBytes) {
                $downloadCacheSizeBytes = 0
            }

            $latestWriteTime = ($downloadFiles | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1).LastWriteTime
            $hasRecentActivity = $false

            if ($latestWriteTime) {
                $hasRecentActivity = ((Get-Date) - $latestWriteTime).TotalMinutes -le $recentActivityMinutes
            }

            if ($downloadCacheSizeBytes -gt $maxDownloadCacheSizeBytes -and -not $hasRecentActivity) {
                $downloadCacheSizeGB = [math]::Round(($downloadCacheSizeBytes / 1GB), 2)
                $issues.Add("Download cache is abnormally large and appears stale ($downloadCacheSizeGB GB)")
            }
            elseif ($downloadCacheSizeBytes -gt $maxDownloadCacheSizeBytes -and $hasRecentActivity) {
                $downloadCacheSizeGB = [math]::Round(($downloadCacheSizeBytes / 1GB), 2)
                Write-Output "Download cache is large but recent activity was detected ($downloadCacheSizeGB GB)"
            }
        }
    }

    if ($issues.Count -gt 0) {
        Write-Output ("Windows Update cache health issues detected: " + ($issues -join "; "))
        exit 1
    }

    Write-Output "Windows Update cache appears healthy"
    exit 0
}
catch {
    Write-Output ("Technical script error: " + $_.Exception.Message)
    exit 1
}

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCtEL0qTSK94wqJ
# RMOqj/cFZcRbNF5IcOEme9xHyXX6b6CCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCAmaez+w90vQ8bf02CCB5gFpRZqraHe3Zq4yYo9StI72jANBgkqhkiG9w0B
# AQEFAASCAYAp+WGcAO8L/C+seb5y5JSTFFNhxTRpJ856E2XjQxHmsq9RaFIOLK+4
# izEI/UiGCjig9l/FdjkosUVjUlY3mfa5Dp1+d/B8j/4RmSvesTrbWGgbsBdzk7d6
# JJHZUo5At4uQ1o74qlvQCRcDvVizIQQk1+1weQKXA0Ik5yx46LLyX9i72Y1gD15T
# gKHxIo/f5xU/RfUmfUysJuC4jnAgrZniKLAmuuIypbI/8tgecVz3Lh2UrJgrxQoX
# 1QTKrzePCiW7ApnYsDYwWhcQ/p2fjytOj35nZs7dP+OWGYBAg9v2nBH3PZoOlOLm
# rbaSNe3u08avci1U4wITc5LB3f9MC0Dvi60JURWI+W9NHz2Hw9SaztxFFAn35NM/
# li0vY6T1PUh2BL76Ux1ShgkV3vw1+/iRgmEFWWauLIvmV1f+442mLnTD+aSoqipU
# AqBdRV+g6DsHgGZUNcoFLSqbsVBSZCcJyzAoCC2Pr8VUGaTG9Tr3LFh8IKpfSd7h
# TXR6hltOVfA=
# SIG # End signature block
