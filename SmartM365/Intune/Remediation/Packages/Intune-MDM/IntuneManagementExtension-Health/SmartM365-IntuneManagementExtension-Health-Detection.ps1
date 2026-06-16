# Name: SmartM365-IntuneManagementExtension-Health-Detection.ps1
# Version: 1.0
# Description: Verifies whether Intune Management Extension (IME) is installed and healthy

$ErrorActionPreference = "Stop"

try {
    # Verify service
    $serviceName = "IntuneManagementExtension"
    $service = Get-Service -Name $serviceName -ErrorAction Stop

    if ($service.Status -ne "Running") {
        Write-Output "Intune Management Extension service is stopped"
        exit 1
    }

    # Verify process
    $process = Get-Process -Name IntuneManagementExtension -ErrorAction SilentlyContinue
    if ($null -eq $process) {
        Write-Output "IME process is not running"
        exit 1
    }

    # Verify installation path
    $installPath = "C:\Program Files (x86)\Microsoft Intune Management Extension"
    if (-not (Test-Path $installPath)) {
        Write-Output "IME installation folder is missing"
        exit 1
    }

    # Verify logs folder
    $logPath = "C:\ProgramData\SmartM365\IntuneRemediation\Logs"
    if (-not (Test-Path $logPath)) {
        Write-Output "IME logs folder is missing"
        exit 1
    }

    # Verify main log file
    $agentLog = Join-Path $logPath "AgentExecutor.log"

    if (-not (Test-Path $agentLog)) {
        Write-Output "AgentExecutor.log is missing"
        exit 1
    }

    # Verify recent activity
    $lastWrite = (Get-Item $agentLog).LastWriteTime
    $delta = (Get-Date) - $lastWrite

    if ($delta.TotalHours -gt 24) {
        Write-Output "IME appears inactive (no recent activity detected)"
        exit 1
    }

    # Verify execution activity in logs
    $hasExecution = Select-String -Path $agentLog -Pattern "Executing" -Quiet

    if (-not $hasExecution) {
        Write-Output "No IME execution activity detected in logs"
        exit 1
    }

    Write-Output "Intune Management Extension is healthy"
    exit 0
}
catch {
    Write-Output ("Technical script error: " + $_.Exception.Message)
    exit 1
}
