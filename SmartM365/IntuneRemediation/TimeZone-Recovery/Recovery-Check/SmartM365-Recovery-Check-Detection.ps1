# Name: SmartM365-Recovery-Check-Detection.ps1
# Version: 1.0
$taskPath = "\Microsoft\Windows\Workplace Join"
$taskName = "Recovery-Check"

try {
    $task = Get-ScheduledTask -TaskPath $taskPath -TaskName $taskName -ErrorAction Stop

    # Vérifie si la tâche est activée
    if ($task.State -ne "Disabled") {
        # Tâche présente et activée → conforme
        exit 0
    }
    else {
        # Tâche présente mais désactivée → non conforme
        exit 1
    }
}
catch {
    # Tâche absente → non conforme
    exit 1
}
