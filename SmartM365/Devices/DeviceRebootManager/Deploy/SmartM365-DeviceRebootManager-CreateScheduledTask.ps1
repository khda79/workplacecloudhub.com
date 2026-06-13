#Requires -Version 5.1

<#
.SYNOPSIS
    Creates the SmartM365 Device Reboot Manager scheduled task.

.DESCRIPTION
    Registers a user-interactive scheduled task that starts the WPF GUI at user
    logon and then regularly while a user session is available. This script is
    intended for Intune Win32 deployments running as SYSTEM.
#>

[CmdletBinding()]
param(
    [string]$InstallPath = "$env:ProgramData\SmartM365\DeviceRebootManager",
    [string]$TaskPath = '\SmartM365\',
    [string]$TaskName = 'Device Reboot Manager',
    [int]$RepeatIntervalMinutes = 240,
    [string]$PowerShellPath = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe",
    [string]$ConfigPath = '',
    [switch]$RunOnceNow
)

$ErrorActionPreference = 'Stop'

function Quote-Argument {
    param([Parameter(Mandatory = $true)][string]$Value)
    return ('"{0}"' -f ($Value -replace '"', '\"'))
}

if ($RepeatIntervalMinutes -lt 15) {
    throw 'RepeatIntervalMinutes must be at least 15.'
}

if (-not (Test-Path -LiteralPath $PowerShellPath)) {
    $PowerShellPath = 'powershell.exe'
}

$appScriptPath = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager-GUI.ps1'
if (-not (Test-Path -LiteralPath $appScriptPath)) {
    throw "Device Reboot Manager script not found: $appScriptPath"
}

if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    $candidateConfigPath = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager-GUI.config.json'
    if (Test-Path -LiteralPath $candidateConfigPath) {
        $ConfigPath = $candidateConfigPath
    }
}

$taskArguments = @(
    '-STA'
    '-NoProfile'
    '-WindowStyle'
    'Hidden'
    '-ExecutionPolicy'
    'Bypass'
    '-File'
    (Quote-Argument -Value $appScriptPath)
)

if (-not [string]::IsNullOrWhiteSpace($ConfigPath)) {
    $taskArguments += @('-ConfigPath', (Quote-Argument -Value $ConfigPath))
}

$action = New-ScheduledTaskAction -Execute $PowerShellPath -Argument ($taskArguments -join ' ')
$logonTrigger = New-ScheduledTaskTrigger -AtLogOn
$repeatTrigger = New-ScheduledTaskTrigger -Daily -At (Get-Date)
$repeatTrigger.Repetition.Interval = New-TimeSpan -Minutes $RepeatIntervalMinutes
$repeatTrigger.Repetition.Duration = New-TimeSpan -Days 1

$principal = New-ScheduledTaskPrincipal -GroupId 'S-1-5-32-545' -LogonType Interactive -RunLevel LeastPrivilege
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 1)

Register-ScheduledTask `
    -TaskPath $TaskPath `
    -TaskName $TaskName `
    -Action $action `
    -Trigger @($logonTrigger, $repeatTrigger) `
    -Principal $principal `
    -Settings $settings `
    -Description 'Starts the SmartM365 Device Reboot Manager GUI in the interactive user session.' `
    -Force | Out-Null

if ($RunOnceNow) {
    Start-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName
}

Write-Output ("Scheduled task registered: {0}{1}" -f $TaskPath, $TaskName)
