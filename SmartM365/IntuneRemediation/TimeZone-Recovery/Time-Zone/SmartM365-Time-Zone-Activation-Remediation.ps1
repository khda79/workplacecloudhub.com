# Name: SmartM365-Time-Zone-Activation-Remediation.ps1
# Version: 1.0
Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\tzautoupdate" -Name "Start" -Value 3
