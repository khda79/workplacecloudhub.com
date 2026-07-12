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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCtEL0qTSK94wqJ
# RMOqj/cFZcRbNF5IcOEme9xHyXX6b6CCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAmaez+w90vQ8bf02CC
# B5gFpRZqraHe3Zq4yYo9StI72jANBgkqhkiG9w0BAQEFAASCAYCgdX/n1M5S66V6
# AYCQufSufgX4ZnUrkJsQiVEHqV+bO87khZWxR/TcAZjDO3cd58rZnCSDpGE6PWpS
# 11lDS2f/YmKmFA4UHrzS+uh5pAftm3fkG/aPyYK4wo8phZdawZ87pTP+G7xikpPh
# SbArPZosVJ57IF3P9VpvGyt85wTmFGj4voiaDLU3ETNka1cvToVjLc/zxyMXD2b0
# H9MjghX3iKoGorDhkAg5WKdkls1WbL+LB3wIj9V2jDZNoQ+HCfzQsDRkddcpxLRx
# cK1Voa2bi1d76ROgc8aeDVmQwkF1kWdTIriMvNHRB4dniLc0SgSBAi9F/k6fdImb
# vBbmCm9BPJ8kt1RyEHPuuHbBA0jVrgOFkGiiZ5GDs8XArkkDc3zrnQx/SzQGNyjQ
# FK1rbMJ7eiQOEYVx5zup6f86c4ROj3qMcl7glZEMIrZ5mB6WGR8g/43tfoUliran
# z8MpY7OUbEjn8TElaCaSeBnpwXF8lC5wWbz/xf+lXQ9xuwJM8P0=
# SIG # End signature block
