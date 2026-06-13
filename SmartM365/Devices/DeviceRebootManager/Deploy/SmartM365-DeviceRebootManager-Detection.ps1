#Requires -Version 5.1

<#
.SYNOPSIS
    Intune Win32 detection script for SmartM365 Device Reboot Manager.
#>

[CmdletBinding()]
param(
    [string]$InstallPath = "$env:ProgramData\SmartM365\DeviceRebootManager",
    [string]$TaskPath = '\SmartM365\',
    [string]$TaskName = 'Device Reboot Manager'
)

$ErrorActionPreference = 'Stop'

$requiredFiles = @(
    'SmartM365-DeviceRebootManager-GUI.ps1'
    'SmartM365-DeviceRebootManager-GUI.strings.psd1'
    'SmartM365-DeviceRebootManager-GUI.config.json'
)

$missing = New-Object System.Collections.Generic.List[string]

foreach ($fileName in $requiredFiles) {
    $path = Join-Path -Path $InstallPath -ChildPath $fileName
    if (-not (Test-Path -LiteralPath $path)) {
        $missing.Add($path)
    }
}

$task = Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction SilentlyContinue
if ($null -eq $task) {
    $missing.Add(("{0}{1}" -f $TaskPath, $TaskName))
}

if ($missing.Count -gt 0) {
    Write-Output 'SmartM365 Device Reboot Manager is not detected.'
    foreach ($item in $missing) {
        Write-Output ("Missing: {0}" -f $item)
    }
    exit 1
}

Write-Output 'SmartM365 Device Reboot Manager detected.'
exit 0
