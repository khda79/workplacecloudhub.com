
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

$ErrorActionPreference = "Stop"
$Scenario = "Reset-WindowsHello"
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogPath = Join-Path -Path $LogRoot -ChildPath "$Scenario-Remediation.log"

function Write-SmartM365Log {
    param([string]$Message)
    $line = "{0} [{1}] {2}" -f (Get-Date -Format "s"), $Scenario, $Message
    Write-Output $line
    Add-Content -LiteralPath $LogPath -Value $line -Encoding utf8
}

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-SmartM365Log "RemediationStarted"

    $ngcPath = "C:\\Windows\\ServiceProfiles\\LocalService\\AppData\\Local\\Microsoft\\NGC"
    $ngcService = Get-Service -Name NgcSvc -ErrorAction SilentlyContinue

    if ($ngcService -and $ngcService.Status -ne "Stopped") {
        Write-SmartM365Log "Stopping NgcSvc service."
        Stop-Service -Name NgcSvc -Force -ErrorAction SilentlyContinue
    }

    if (Test-Path -LiteralPath $ngcPath) {
        Write-SmartM365Log "Taking ownership of NGC folder."
        takeown.exe /F $ngcPath /A /R | Out-Null
        icacls.exe $ngcPath /grant administrators:F /T | Out-Null
        Write-SmartM365Log "Deleting NGC folder."
        Remove-Item -LiteralPath $ngcPath -Recurse -Force -ErrorAction Stop
    }
    else {
        Write-SmartM365Log "NGC folder not found; nothing to delete."
    }

    if ($ngcService) {
        Write-SmartM365Log "Starting NgcSvc service."
        Start-Service -Name NgcSvc -ErrorAction SilentlyContinue
    }

    $dsreg = dsregcmd /status | Out-String
    if ($dsreg -notmatch "AzureAdJoined.*YES" -or $dsreg -notmatch "DomainJoined.*YES") {
        Write-SmartM365Log "Device join state is incomplete; requesting dsregcmd /join."
        dsregcmd /join | Out-Null
    }

    Write-SmartM365Log "RemediationCompleted"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    try { Write-SmartM365Log "RemediationFailed Message=$($_.Exception.Message)" } catch { Write-Output "LogWriteFailed=$($_.Exception.Message)" }
    exit 1
}
