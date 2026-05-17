# Name: SmartM365-Force-AllowOSUpgrade-Detection.ps1
# Version: 1.0
# Detection script - Check if AllowOSUpgrade is set correctly
$regPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\OSUpgrade"
$regName = "AllowOSUpgrade"
$expectedValue = 1

if (Test-Path $regPath) {
    try {
        $actualValue = Get-ItemProperty -Path $regPath -Name $regName -ErrorAction Stop | Select-Object -ExpandProperty $regName
        if ($actualValue -eq $expectedValue) {
            Write-Output "Compliant"
            exit 0
        } else {
            Write-Output "Non-compliant: Value is $actualValue"
            exit 1
        }
    } catch {
        Write-Output "Non-compliant: Value not found"
        exit 1
    }
} else {
    Write-Output "Non-compliant: Registry path not found"
    exit 1
}