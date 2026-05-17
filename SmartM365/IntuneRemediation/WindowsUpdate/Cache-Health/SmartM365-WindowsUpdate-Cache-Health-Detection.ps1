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
