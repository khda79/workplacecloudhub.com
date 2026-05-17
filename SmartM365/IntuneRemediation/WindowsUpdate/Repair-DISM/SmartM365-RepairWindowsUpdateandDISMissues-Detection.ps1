# Name: SmartM365-RepairWindowsUpdateandDISMissues-Detection.ps1
# Version: 1.0
# Detection Script for Windows Update and DISM issues

$RemediationNeeded = $false

# Check Windows Update service
$WUService = Get-Service -Name wuauserv -ErrorAction SilentlyContinue
if ($WUService.Status -ne 'Running') {
    Write-Output "Windows Update service is not running."
    $RemediationNeeded = $true
}

# Check DISM health
$DismLog = "$env:windir\Logs\DISM\dism.log"
$DismError = $false
if (Test-Path $DismLog) {
    $DismError = Select-String -Path $DismLog -Pattern "error|failed" -SimpleMatch
    if ($DismError) {
        Write-Output "DISM log contains errors."
        $RemediationNeeded = $true
    }
}

# Check CBS log for update errors
$CBSLog = "$env:windir\Logs\CBS\cbs.log"
$CBSError = $false
if (Test-Path $CBSLog) {
    $CBSError = Select-String -Path $CBSLog -Pattern "error|failed" -SimpleMatch
    if ($CBSError) {
        Write-Output "CBS log contains errors."
        $RemediationNeeded = $true
    }
}

# Check Windows Update cache
$WUCache = "$env:windir\SoftwareDistribution\Download"
if (Test-Path $WUCache) {
    $CacheFiles = Get-ChildItem -Path $WUCache -Recurse -ErrorAction SilentlyContinue
    if ($CacheFiles.Count -gt 1000) {
        Write-Output "Windows Update cache contains excessive files."
        $RemediationNeeded = $true
    }
}

# Output for Intune Remediation
if ($RemediationNeeded) {
    Write-Output "Remediation required."
    exit 1
} else {
    Write-Output "No remediation required."
    exit 0
}
