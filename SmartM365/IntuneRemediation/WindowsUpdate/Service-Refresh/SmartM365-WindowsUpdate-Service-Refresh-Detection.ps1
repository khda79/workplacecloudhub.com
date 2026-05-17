# Name: SmartM365-WindowsUpdate-Service-Refresh-Detection.ps1
# Version: 1.0
# Description: Detects Windows Update service states that can block scan or connectivity health.

$ErrorActionPreference = "Stop"

$ScriptName = "Detect-WindowsUpdate-Service-Refresh"
$Version = "1.0"
$RequiredServices = @("bits", "wuauserv", "dosvc", "cryptsvc")

try {
    $issues = New-Object System.Collections.Generic.List[string]

    foreach ($serviceName in $RequiredServices) {
        $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue

        if ($null -eq $service) {
            $issues.Add("ServiceMissing=$serviceName")
            continue
        }

        $serviceCim = Get-CimInstance -ClassName Win32_Service -Filter "Name='$serviceName'" -ErrorAction SilentlyContinue

        if ($serviceCim -and $serviceCim.StartMode -eq "Disabled") {
            $issues.Add("ServiceDisabled=$serviceName")
        }
    }

    if ($issues.Count -gt 0) {
        Write-Output ("{0} v{1}: remediation required. {2}" -f $ScriptName, $Version, ($issues -join " | "))
        exit 1
    }

    Write-Output ("{0} v{1}: Windows Update services are present and not disabled." -f $ScriptName, $Version)
    exit 0
}
catch {
    Write-Output ("{0} v{1}: detection failed: {2}" -f $ScriptName, $Version, $_.Exception.Message)
    exit 1
}
