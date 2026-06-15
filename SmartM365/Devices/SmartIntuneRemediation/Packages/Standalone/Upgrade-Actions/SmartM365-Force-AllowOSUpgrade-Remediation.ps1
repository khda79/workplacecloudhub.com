# Name: SmartM365-Force-AllowOSUpgrade-Remediation.ps1
# Version: 1.0
# Remediation script - Force Windows 11 Upgrade via registry
$regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\OSUpgrade"
$regName = "AllowOSUpgrade"
$regValue = 1

try {
    if (-not (Test-Path $regPath)) {
        New-Item -Path $regPath -Force | Out-Null
    }

    New-ItemProperty -Path $regPath -Name $regName -PropertyType DWord -Value $regValue -Force | Out-Null
    Write-Output "Status=Completed RegistryValue=$regName Value=$regValue"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    exit 1
}
