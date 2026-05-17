# Name: SmartM365-ForceWUScan-Detection.ps1
# Version: 1.0
[CmdletBinding()]
param(
    [int]$MaxHoursSinceLastWUEvent = 24
)

$logName = 'Microsoft-Windows-WindowsUpdateClient/Operational'

try {
    # Verify log is enabled and accessible
    $logInfo = Get-WinEvent -ListLog $logName -ErrorAction Stop
    if (-not $logInfo.IsEnabled) {
        Write-Warning "$logName is disabled. Remediation required."
        exit 1
    }

    # Get last Windows Update event (any)
    $lastEvent = Get-WinEvent -LogName $logName -MaxEvents 1 -ErrorAction Stop
    if (-not $lastEvent) {
        Write-Output "No Windows Update events found. Remediation required."
        exit 1
    }

    $ageHours = (New-TimeSpan -Start $lastEvent.TimeCreated -End (Get-Date)).TotalHours
    if ($ageHours -ge $MaxHoursSinceLastWUEvent) {
        Write-Output ("Last WU event is {0:N1} hours old (>{1}h). Remediation required." -f $ageHours, $MaxHoursSinceLastWUEvent)
        exit 1
    }

    Write-Output ("Last WU event is {0:N1} hours old. Device OK." -f $ageHours)
    exit 0
}
catch {
    Write-Error "Detection failed: $_"
    exit 1
}