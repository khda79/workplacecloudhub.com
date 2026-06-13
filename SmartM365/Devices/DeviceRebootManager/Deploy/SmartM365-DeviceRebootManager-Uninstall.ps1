#Requires -Version 5.1

<#
.SYNOPSIS
    Uninstalls SmartM365 Device Reboot Manager.
#>

[CmdletBinding()]
param(
    [string]$InstallPath = "$env:ProgramData\SmartM365\DeviceRebootManager",
    [string]$TaskPath = '\SmartM365\',
    [string]$TaskName = 'Device Reboot Manager',
    [switch]$KeepConfig
)

$ErrorActionPreference = 'Stop'

try {
    $task = Get-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction SilentlyContinue
    if ($null -ne $task) {
        Unregister-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -Confirm:$false
    }
}
catch {
    Write-Warning ("Unable to remove scheduled task. {0}" -f $_.Exception.Message)
}

if (Test-Path -LiteralPath $InstallPath) {
    if ($KeepConfig) {
        $configPath = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager-GUI.config.json'
        Get-ChildItem -LiteralPath $InstallPath -Force |
            Where-Object { $_.FullName -ne $configPath } |
            ForEach-Object { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue }
    }
    else {
        Remove-Item -LiteralPath $InstallPath -Recurse -Force
    }
}

Write-Output 'SmartM365 Device Reboot Manager uninstalled.'
