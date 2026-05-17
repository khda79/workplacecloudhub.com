# Name: SmartM365-Force-AllowOSUpgrade-Remediation.ps1
# Version: 1.0
# Remediation script - Force Windows 11 Upgrade via registry
$regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\OSUpgrade"
$regName = "AllowOSUpgrade"
$regValue = 1

# Create the key if it doesn't exist
if (-not (Test-Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
}

# Set the value
New-ItemProperty -Path $regPath -Name $regName -PropertyType DWord -Value $regValue -Force | Out-Null
Write-Output "Registry key set: $regName = $regValue"