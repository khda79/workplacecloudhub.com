# Name: SmartM365-Time-Zone-Detection.ps1
# Version: 1.0
# Variables for registry path and property name
$registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\tzautoupdate"
$propertyName = "Start"

try {
    $currentValue = Get-ItemProperty -Path $registryPath -Name $propertyName -ErrorAction Stop

    if ($currentValue.Start -ne 3) {
        Write-Output "Status=NonCompliant CurrentValue=$($currentValue.Start) ExpectedValue=3"
        exit 1
    }

    Write-Output "Status=Compliant CurrentValue=$($currentValue.Start)"
    exit 0
}
catch {
    Write-Output "Status=Error"
    Write-Output "Message=$($_.Exception.Message)"
    Exit 1
}
