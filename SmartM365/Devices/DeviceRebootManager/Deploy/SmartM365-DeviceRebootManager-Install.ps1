#Requires -Version 5.1

<#
.SYNOPSIS
    Installs SmartM365 Device Reboot Manager for Intune Win32 deployment.

.DESCRIPTION
    Copies runtime files to ProgramData and creates the scheduled task used to
    launch the GUI in the interactive user session.
#>

[CmdletBinding()]
param(
    [string]$InstallPath = "$env:ProgramData\SmartM365\DeviceRebootManager",
    [string]$ConfigSourcePath = '',
    [switch]$ForceConfig,
    [switch]$SkipScheduledTask,
    [string]$TaskPath = '\SmartM365\',
    [string]$TaskName = 'Device Reboot Manager',
    [int]$RepeatIntervalMinutes = 240
)

$ErrorActionPreference = 'Stop'

function Test-IsAdministrator {
    $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($identity)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdministrator)) {
    throw 'Installation must run elevated. Intune Win32 install should run as SYSTEM.'
}

$deployRoot = $PSScriptRoot
$sourceRoot = Split-Path -Path $deployRoot -Parent

$requiredFiles = @(
    'SmartM365-DeviceRebootManager-GUI.ps1'
    'SmartM365-DeviceRebootManager-GUI.strings.psd1'
    'SmartM365-DeviceRebootManager-GUI.config.json.template'
    'Start-SmartM365-DeviceRebootManager-GUI.cmd'
    'Start-SmartM365-DeviceRebootManager-GUI-Test.cmd'
)

$optionalFiles = @('SmartM365-logo.ico')

New-Item -ItemType Directory -Path $InstallPath -Force | Out-Null

foreach ($fileName in $requiredFiles) {
    $sourcePath = Join-Path -Path $sourceRoot -ChildPath $fileName
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        throw "Required package file not found: $sourcePath"
    }

    Copy-Item -LiteralPath $sourcePath -Destination (Join-Path -Path $InstallPath -ChildPath $fileName) -Force
}

foreach ($fileName in $optionalFiles) {
    $sourcePath = Join-Path -Path $sourceRoot -ChildPath $fileName
    if (Test-Path -LiteralPath $sourcePath) {
        Copy-Item -LiteralPath $sourcePath -Destination (Join-Path -Path $InstallPath -ChildPath $fileName) -Force
    }
}

$runtimeConfigPath = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager-GUI.config.json'
if (-not [string]::IsNullOrWhiteSpace($ConfigSourcePath)) {
    if (-not (Test-Path -LiteralPath $ConfigSourcePath)) {
        throw "Config source file not found: $ConfigSourcePath"
    }

    Copy-Item -LiteralPath $ConfigSourcePath -Destination $runtimeConfigPath -Force
}
elseif ($ForceConfig -or -not (Test-Path -LiteralPath $runtimeConfigPath)) {
    $templatePath = Join-Path -Path $InstallPath -ChildPath 'SmartM365-DeviceRebootManager-GUI.config.json.template'
    Copy-Item -LiteralPath $templatePath -Destination $runtimeConfigPath -Force
}

if (-not $SkipScheduledTask) {
    $taskScriptPath = Join-Path -Path $deployRoot -ChildPath 'SmartM365-DeviceRebootManager-CreateScheduledTask.ps1'
    if (-not (Test-Path -LiteralPath $taskScriptPath)) {
        throw "Scheduled task helper not found: $taskScriptPath"
    }

    & $taskScriptPath `
        -InstallPath $InstallPath `
        -TaskPath $TaskPath `
        -TaskName $TaskName `
        -RepeatIntervalMinutes $RepeatIntervalMinutes `
        -ConfigPath $runtimeConfigPath
}

Write-Output "SmartM365 Device Reboot Manager installed to: $InstallPath"
