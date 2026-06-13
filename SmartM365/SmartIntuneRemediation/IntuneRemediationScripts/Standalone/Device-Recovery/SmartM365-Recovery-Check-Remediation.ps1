# Name: SmartM365-Recovery-Check-Remediation.ps1
# Version: 1.0
$taskPath = "\Microsoft\Windows\Workplace Join"
$taskName = "Recovery-Check"

try {
    # Active la tâche
    Enable-ScheduledTask -TaskPath $taskPath -TaskName $taskName -ErrorAction Stop

    # Exécute la tâche
    Start-ScheduledTask -TaskPath $taskPath -TaskName $taskName -ErrorAction Stop
}
catch {
    Write-Output "Erreur lors de la remédiation : $($_.Exception.Message)"
    exit 1
}

exit 0

## Test
# powershell.exe -ExecutionPolicy Bypass -File .\SmartM365-Recovery-Check-Detection.ps1
# powershell.exe -ExecutionPolicy Bypass -File .\SmartM365-Recovery-Check-Remediation.ps1
