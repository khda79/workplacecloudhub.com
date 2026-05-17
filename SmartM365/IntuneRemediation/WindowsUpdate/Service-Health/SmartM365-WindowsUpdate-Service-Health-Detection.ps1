# Name: SmartM365-WindowsUpdate-Service-Health-Detection.ps1
# Version: 1.0
# Description: Verifies whether Windows Update related services are configured and in a healthy state

$ErrorActionPreference = "Stop"

try {
    $servicesToCheck = @(
        @{
            Name = "wuauserv"
            DisplayName = "Windows Update"
            Required = $true
        },
        @{
            Name = "BITS"
            DisplayName = "Background Intelligent Transfer Service"
            Required = $true
        },
        @{
            Name = "DoSvc"
            DisplayName = "Delivery Optimization"
            Required = $true
        },
        @{
            Name = "UsoSvc"
            DisplayName = "Update Orchestrator Service"
            Required = $true
        }
    )

    $issues = New-Object System.Collections.Generic.List[string]
    $healthyStates = New-Object System.Collections.Generic.List[string]

    foreach ($serviceItem in $servicesToCheck) {
        $serviceName = $serviceItem.Name
        $serviceDisplayName = $serviceItem.DisplayName

        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            if ($serviceItem.Required) {
                $issues.Add("$serviceName is missing")
            }
            continue
        }

        $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$serviceName'" -ErrorAction SilentlyContinue

        if ($null -eq $serviceCim) {
            $issues.Add("$serviceName CIM state is unavailable")
            continue
        }

        if ($serviceCim.StartMode -eq "Disabled") {
            $issues.Add("$serviceName is disabled")
            continue
        }

        if ($service.Status -in @("StopPending", "PausePending", "Paused")) {
            $issues.Add("$serviceName is in an unhealthy state ($($service.Status))")
            continue
        }

        if ($service.Status -eq "Running") {
            $healthyStates.Add("$serviceName=Running,StartupType=$($serviceCim.StartMode)")
            continue
        }

        if ($service.Status -eq "Stopped") {
            if ($serviceCim.StartMode -in @("Manual", "Auto")) {
                $healthyStates.Add("$serviceName=Stopped,StartupType=$($serviceCim.StartMode)")
                continue
            }

            $issues.Add("$serviceName is stopped with unexpected startup type ($($serviceCim.StartMode))")
            continue
        }

        $issues.Add("$serviceName is in an unexpected state ($($service.Status))")
    }

    if ($issues.Count -gt 0) {
        Write-Output ("Windows Update service health issues detected: " + ($issues -join "; "))
        exit 1
    }

    Write-Output ("Windows Update related services are healthy: " + ($healthyStates -join "; "))
    exit 0
}
catch {
    Write-Output ("Technical script error: " + $_.Exception.Message)
    exit 1
}
