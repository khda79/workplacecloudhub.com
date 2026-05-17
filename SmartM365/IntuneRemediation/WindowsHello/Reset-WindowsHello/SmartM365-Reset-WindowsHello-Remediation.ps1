
<#
.SYNOPSIS
    Version: 1.0
    Intune remediation script for Windows Hello Key Trust issue.
.DESCRIPTION
    This script:
    - Logs actions to C:\ProgramData\SmartM365\IntuneRemediation\Logs\Remediate-WindowsHello.log.
    - Stops NgcSvc service.
    - Takes ownership of NGC folder and deletes it.
    - Restarts NgcSvc.
    - Optionally forces dsregcmd /join if Hybrid Join is broken.
    - Runs silently and exits with code 0 for success.
.NOTES
    Run as SYSTEM via Intune.
#>

$LogRoot = Join-Path -Path $env:ProgramData -ChildPath 'SmartM365\IntuneRemediation\Logs\Reset-WindowsHello'
$LogPath = Join-Path -Path $LogRoot -ChildPath 'Reset-WindowsHello-Remediation.log'
if (-not (Test-Path -LiteralPath $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
if (!(Test-Path (Split-Path $LogPath))) {
    New-Item -ItemType Directory -Path (Split-Path $LogPath) -Force
}
Start-Transcript -Path $LogPath -Append

Write-Host "[INFO] Starting Windows Hello remediation..."

# Stop NgcSvc service
Write-Host "[INFO] Stopping NgcSvc service..."
Stop-Service -Name NgcSvc -Force -ErrorAction SilentlyContinue

# NGC folder path
$ngcPath = "C:\\Windows\\ServiceProfiles\\LocalService\\AppData\\Local\\Microsoft\\NGC"

# Take ownership and delete NGC folder
if (Test-Path $ngcPath) {
    Write-Host "[INFO] Taking ownership of NGC folder..."
    takeown /F $ngcPath /A /R | Out-Null
    icacls $ngcPath /grant administrators:F /T | Out-Null
    Write-Host "[INFO] Deleting NGC folder..."
    Remove-Item $ngcPath -Recurse -Force
} else {
    Write-Host "[INFO] NGC folder not found. Nothing to delete."
}

# Restart NgcSvc service
Write-Host "[INFO] Restarting NgcSvc service..."
Start-Service -Name NgcSvc -ErrorAction SilentlyContinue

# Optional: Check Hybrid Join and re-join if needed
$dsreg = dsregcmd /status | Out-String
if ($dsreg -notmatch "AzureAdJoined.*YES" -or $dsreg -notmatch "DomainJoined.*YES") {
    Write-Host "[WARN] Device not properly joined. Attempting dsregcmd /join..."
    dsregcmd /join
}

Write-Host "[INFO] Windows Hello remediation completed successfully."
Stop-Transcript

exit 0
