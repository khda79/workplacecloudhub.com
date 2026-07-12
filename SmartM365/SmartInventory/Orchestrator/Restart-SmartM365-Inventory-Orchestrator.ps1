<#PSScriptInfo
.VERSION 1.0.0
.GUID 7b0b7843-4544-4a5e-9c5d-8f22b499fbf7
.AUTHOR https://github.com/khda79/workplacecloudhub.com
.COMPANYNAME WorkplaceCloudHub
.COPYRIGHT (c) WorkplaceCloudHub. All rights reserved.
.TAGS SmartM365 Orchestrator Restart ScheduledTask
.LICENSEURI https://github.com/khda79/workplacecloudhub.com
.PROJECTURI https://github.com/khda79/workplacecloudhub.com
#>
<#
.SYNOPSIS
Stops the SmartM365 Inventory Orchestrator cleanly and restarts its scheduled task.

.DESCRIPTION
This helper is intended for deployment/restart operations after updating the
repository. It first delegates the clean stop to SmartM365-Inventory-Orchestrator.ps1
with -Stop so the resident instance can save state and release its lock. It then
starts the existing Windows scheduled task so the orchestrator runs under the
registered service account and Task Scheduler settings.

.PARAMETER Tenant
Tenant key used by the orchestrator task. Defaults to prod.

.PARAMETER TaskPath
Windows Task Scheduler folder. Defaults to \WCH\.

.PARAMETER TaskName
Windows Task Scheduler task name. Defaults to "SmartM365 Inventory Orchestrator - <Tenant>".

.PARAMETER StopTimeoutSeconds
Maximum time to wait for the clean stop request.

.PARAMETER RestartDelaySeconds
Delay between the stop phase and task start.

.PARAMETER StartWaitSeconds
Maximum time to wait for the task to enter Running state after Start-ScheduledTask.
#>
[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$Tenant = 'prod',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$TaskPath = '\WCH\',

    [Parameter()]
    [string]$TaskName = '',

    [Parameter()]
    [ValidateRange(1, 3600)]
    [int]$StopTimeoutSeconds = 180,

    [Parameter()]
    [ValidateRange(0, 600)]
    [int]$RestartDelaySeconds = 5,

    [Parameter()]
    [ValidateRange(1, 600)]
    [int]$StartWaitSeconds = 60
)

$ErrorActionPreference = 'Stop'

function Write-RestartLog {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter()]
        [ValidateSet('INFO', 'SUCCESS', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $color = switch ($Level) {
        'SUCCESS' { 'Green' }
        'WARN' { 'Yellow' }
        'ERROR' { 'Red' }
        default { 'Cyan' }
    }
    Write-Host ("[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message) -ForegroundColor $color
}

function Get-CurrentPowerShellExecutable {
    if ($PSVersionTable.PSEdition -eq 'Core') {
        $processPath = (Get-Process -Id $PID -ErrorAction Stop).Path
        if (-not [string]::IsNullOrWhiteSpace($processPath) -and (Test-Path -LiteralPath $processPath -PathType Leaf)) {
            return $processPath
        }
    }

    $programFilesPwsh = Join-Path -Path $env:ProgramFiles -ChildPath 'PowerShell\7\pwsh.exe'
    if (Test-Path -LiteralPath $programFilesPwsh -PathType Leaf) {
        return $programFilesPwsh
    }

    $command = Get-Command -Name 'pwsh.exe' -ErrorAction SilentlyContinue
    if ($null -ne $command -and -not [string]::IsNullOrWhiteSpace($command.Source)) {
        return $command.Source
    }

    throw 'PowerShell 7 (pwsh.exe) was not found.'
}

function Test-ScheduledTaskRunning {
    param(
        [Parameter(Mandatory = $true)]
        $Task
    )

    if ([string]$Task.State -eq 'Running') {
        return $true
    }

    try {
        return ([int]$Task.State -eq 4)
    }
    catch {
        return $false
    }
}

function Get-OrchestratorTask {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    try {
        return Get-ScheduledTask -TaskPath $Path -TaskName $Name -ErrorAction Stop
    }
    catch {
        throw ("Scheduled task not found. TaskPath={0}; TaskName={1}; {2}" -f $Path, $Name, $_.Exception.Message)
    }
}

if ([string]::IsNullOrWhiteSpace($TaskName)) {
    $TaskName = "SmartM365 Inventory Orchestrator - $Tenant"
}

if (-not $TaskPath.EndsWith('\')) {
    $TaskPath = "$TaskPath\"
}

$orchestratorScript = Join-Path -Path $PSScriptRoot -ChildPath 'SmartM365-Inventory-Orchestrator.ps1'
if (-not (Test-Path -LiteralPath $orchestratorScript -PathType Leaf)) {
    throw "Orchestrator script not found: $orchestratorScript"
}

$pwsh = Get-CurrentPowerShellExecutable

Write-RestartLog -Message ("Restart requested for tenant {0} on {1}." -f $Tenant, $env:COMPUTERNAME)
Write-RestartLog -Message ("TaskPath={0}; TaskName={1}" -f $TaskPath, $TaskName)
Write-RestartLog -Message ("Stop command: `"{0}`" -NoProfile -ExecutionPolicy Bypass -File `"{1}`" -Tenant {2} -Stop -StopTimeoutSeconds {3}" -f $pwsh, $orchestratorScript, $Tenant, $StopTimeoutSeconds)

& $pwsh -NoProfile -ExecutionPolicy Bypass -File $orchestratorScript -Tenant $Tenant -Stop -StopTimeoutSeconds $StopTimeoutSeconds
$stopExitCode = $LASTEXITCODE
if ($stopExitCode -ne 0) {
    throw "Clean stop failed with exit code $stopExitCode. The scheduled task was not restarted."
}

if ($RestartDelaySeconds -gt 0) {
    Write-RestartLog -Message ("Waiting {0} second(s) before restart." -f $RestartDelaySeconds)
    Start-Sleep -Seconds $RestartDelaySeconds
}

$task = Get-OrchestratorTask -Path $TaskPath -Name $TaskName
if (Test-ScheduledTaskRunning -Task $task) {
    throw ("Scheduled task is still Running after clean stop. TaskPath={0}; TaskName={1}; State={2}. Restart aborted." -f $TaskPath, $TaskName, $task.State)
}

Write-RestartLog -Message ("Starting scheduled task: {0}{1}" -f $TaskPath, $TaskName)
Start-ScheduledTask -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop

$deadline = (Get-Date).AddSeconds($StartWaitSeconds)
do {
    Start-Sleep -Seconds 2
    $task = Get-OrchestratorTask -Path $TaskPath -Name $TaskName
    if (Test-ScheduledTaskRunning -Task $task) {
        Write-RestartLog -Level SUCCESS -Message ("Scheduled task restarted and is Running. TaskPath={0}; TaskName={1}" -f $TaskPath, $TaskName)
        exit 0
    }
} while ((Get-Date) -lt $deadline)

$taskInfo = $null
try {
    $taskInfo = Get-ScheduledTaskInfo -TaskPath $TaskPath -TaskName $TaskName -ErrorAction Stop
}
catch {
    Write-RestartLog -Level WARN -Message ("Unable to read scheduled task info: {0}" -f $_.Exception.Message)
}

if ($null -ne $taskInfo) {
    throw ("Scheduled task did not enter Running state within {0} second(s). State={1}; LastTaskResult={2}; LastRunTime={3}" -f $StartWaitSeconds, $task.State, $taskInfo.LastTaskResult, $taskInfo.LastRunTime)
}

throw ("Scheduled task did not enter Running state within {0} second(s). State={1}" -f $StartWaitSeconds, $task.State)

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBQ6HI+cEnvzrB0
# cQdSJJ7xXQqGRz0C7Jxl6nlxMQtYTKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCAhINIPqix1O1e28hqR27ZG7tD1CfxKhqesD9Hl7+xR8DANBgkqhkiG9w0B
# AQEFAASCAYAIr3VwOFknQ8OynQJ1JVZhuLaMiWm8UQpcTCk8jgTIz31t4cxgmQzm
# sL437T9kt01IiV2yW4QJe1gRIdmGnREx8Fezp9SOdNXZzDHBqDNSWYVoddTPNLYv
# 27ZVoVCGN9Qo4vq/EYyVFGbyZkgW+YCt6MUD+MkAg6KrcH42sbTNllkTXJO3wJw0
# yIjSjE912x4Dxi1hwNtxmi+33zLR1ZcTCLMn7ioaOMxHZsgLTdEeBEeEpYcNBwwa
# kPkSeqrDNL/DBNuSgwUoKAo0Ylv97TLyW8COhxLe6SQEdUHlqZ8IPta6Dd8FE9FJ
# 2OXRsud9vZQoqxOtUGdQ3GShvKrqox8t+wJ8ukJgpuwufqaMJ5ghm6psHag3BiY4
# FxiVakkwgb2J7E3M8iMJH0R8fEsEIxZ9Z/89UBgI7f8EDg2I233r87L/uXYW4jOk
# 6FeimgpcXoPScZywua/umyYoXGH1UfEdsW7u8deb5Nb4H38Qsy3oQclcJQNRHTaQ
# NJKt2LHkrmk=
# SIG # End signature block
