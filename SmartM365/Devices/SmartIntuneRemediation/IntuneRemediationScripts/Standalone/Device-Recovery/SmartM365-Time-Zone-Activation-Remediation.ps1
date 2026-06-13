# Name: SmartM365-Time-Zone-Activation-Remediation.ps1
# Version: 1.0
try {
    Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\tzautoupdate" -Name "Start" -Value 3 -ErrorAction Stop
    Write-Output "Status=Completed Service=tzautoupdate Start=3"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    exit 1
}
