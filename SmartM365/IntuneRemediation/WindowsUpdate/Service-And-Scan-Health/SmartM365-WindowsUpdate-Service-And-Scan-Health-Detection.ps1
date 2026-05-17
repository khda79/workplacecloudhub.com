<#
    Name: SmartM365-WindowsUpdate-Service-And-Scan-Health-Detection.ps1
    Version: 1.0
    Description: Consolidated detection for Windows Update service health and recent scan activity.
#>

[CmdletBinding()]
param(
    [int]$MaxHoursSinceLastWindowsUpdateEvent = 24
)

$ErrorActionPreference = "Stop"
$ScriptName = "WindowsUpdate-Service-And-Scan-Health"
$RequiredServices = @("bits", "wuauserv", "dosvc", "cryptsvc", "UsoSvc")
$WindowsUpdateLogName = "Microsoft-Windows-WindowsUpdateClient/Operational"

try {
    $issues = New-Object System.Collections.Generic.List[string]
    $healthyStates = New-Object System.Collections.Generic.List[string]

    foreach ($serviceName in $RequiredServices) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            $issues.Add("ServiceMissing=$serviceName")
            continue
        }

        $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$serviceName'" -ErrorAction SilentlyContinue

        if ($null -eq $serviceCim) {
            $issues.Add("ServiceCimUnavailable=$serviceName")
            continue
        }

        if ($serviceCim.StartMode -eq "Disabled") {
            $issues.Add("ServiceDisabled=$serviceName")
            continue
        }

        if ($service.Status -in @("StopPending", "PausePending", "Paused")) {
            $issues.Add("ServiceUnhealthy=$serviceName Status=$($service.Status)")
            continue
        }

        $healthyStates.Add("$serviceName=$($service.Status),StartupType=$($serviceCim.StartMode)")
    }

    $logInfo = Get-WinEvent -ListLog $WindowsUpdateLogName -ErrorAction Stop
    if (-not $logInfo.IsEnabled) {
        $issues.Add("WindowsUpdateEventLogDisabled")
    }
    else {
        $lastEvent = Get-WinEvent -LogName $WindowsUpdateLogName -MaxEvents 1 -ErrorAction SilentlyContinue

        if ($null -eq $lastEvent) {
            $issues.Add("WindowsUpdateEventLogEmpty")
        }
        else {
            $ageHours = (New-TimeSpan -Start $lastEvent.TimeCreated -End (Get-Date)).TotalHours

            if ($ageHours -ge $MaxHoursSinceLastWindowsUpdateEvent) {
                $issues.Add(("LastWUEventAgeHours={0:N1}" -f $ageHours))
            }
            else {
                $healthyStates.Add(("LastWUEventAgeHours={0:N1}" -f $ageHours))
            }
        }
    }

    if ($issues.Count -gt 0) {
        Write-Output ("{0}: remediation required. {1}" -f $ScriptName, ($issues -join "; "))
        exit 1
    }

    Write-Output ("{0}: healthy. {1}" -f $ScriptName, ($healthyStates -join "; "))
    exit 0
}
catch {
    Write-Output ("{0}: detection failed. Message={1}" -f $ScriptName, $_.Exception.Message)
    exit 1
}
