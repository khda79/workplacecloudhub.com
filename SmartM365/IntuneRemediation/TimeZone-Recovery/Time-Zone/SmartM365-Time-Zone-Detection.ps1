# Name: SmartM365-Time-Zone-Detection.ps1
# Version: 1.0
$LogRoot = Join-Path -Path $env:ProgramData -ChildPath 'SmartM365\IntuneRemediation\Logs\Time-Zone'
$LogPath = Join-Path -Path $LogRoot -ChildPath 'Time-Zone-Detection.log'
if (-not (Test-Path -LiteralPath $LogRoot)) { New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null }
Start-Transcript -Path $LogPath -Append
# Variables for registry path and property name
$registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\tzautoupdate"
$propertyName = "Start"
 
# Check the current registry value
$currentValue = Get-ItemProperty -Path $registryPath -Name $propertyName
 Write-Host "The current value is $currentValue"
# Check if automatic time zone detection is disabled
if ($currentValue.Start -ne 3) {
    # Return non-compliant status
    Write-Output "NonCompliant"
    Exit 1
} else {
    # Return compliant status
    Write-Output "Compliant"
    Exit 0
}
