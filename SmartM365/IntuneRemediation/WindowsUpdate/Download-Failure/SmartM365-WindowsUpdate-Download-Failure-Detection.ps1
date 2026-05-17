# Name: SmartM365-WindowsUpdate-Download-Failure-Detection.ps1
# Version: 1.0
# Description: Detects an abnormally large or potentially stale Windows Update download cache

$ErrorActionPreference = "Stop"

try {
    $downloadCachePath = "C:\Windows\SoftwareDistribution\Download"
    $maxCacheSizeBytes = 5GB
    $recentActivityMinutes = 60

    if (-not (Test-Path -Path $downloadCachePath)) {
        Write-Output "Windows Update download cache folder was not found"
        exit 0
    }

    $cacheFiles = Get-ChildItem -Path $downloadCachePath -Recurse -File -ErrorAction SilentlyContinue

    if ($null -eq $cacheFiles -or $cacheFiles.Count -eq 0) {
        Write-Output "Windows Update download cache is empty"
        exit 0
    }

    $cacheSizeBytes = ($cacheFiles | Measure-Object -Property Length -Sum).Sum

    if ($null -eq $cacheSizeBytes) {
        $cacheSizeBytes = 0
    }

    $latestWriteTime = ($cacheFiles | Sort-Object -Property LastWriteTime -Descending | Select-Object -First 1).LastWriteTime
    $hasRecentActivity = $false

    if ($latestWriteTime) {
        $hasRecentActivity = ((Get-Date) - $latestWriteTime).TotalMinutes -le $recentActivityMinutes
    }

    if ($cacheSizeBytes -gt $maxCacheSizeBytes -and -not $hasRecentActivity) {
        $cacheSizeGB = [math]::Round(($cacheSizeBytes / 1GB), 2)
        Write-Output "Windows Update download cache is abnormally large and appears stale ($cacheSizeGB GB)"
        exit 1
    }

    if ($cacheSizeBytes -gt $maxCacheSizeBytes -and $hasRecentActivity) {
        $cacheSizeGB = [math]::Round(($cacheSizeBytes / 1GB), 2)
        Write-Output "Windows Update download cache is large ($cacheSizeGB GB) but recent activity was detected"
        exit 0
    }

    $cacheSizeGB = [math]::Round(($cacheSizeBytes / 1GB), 2)
    Write-Output "Windows Update download cache size is normal ($cacheSizeGB GB)"
    exit 0
}
catch {
    Write-Output ("Technical script error: " + $_.Exception.Message)
    exit 1
}
