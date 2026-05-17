<#
.SYNOPSIS
    Version: 1.0
    Repairs Intune Management Extension health for IntuneManagementExtension-Health.
.DESCRIPTION
    Ensures the IME service is configured to start automatically, recreates the expected log directory, restarts the IME service, starts EnterpriseMgmt tasks, and triggers a Windows Update scan to refresh policy-driven activity.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$Scenario = 'IntuneManagementExtension-Health'
$LogRoot = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path $LogRoot "$Scenario-Remediation.log"

function Write-Log { param([string]$Message) $line = '{0} [{1}] {2}' -f (Get-Date -Format 's'), $Scenario, $Message; Write-Output $line; Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8 }
function Start-UsoClient { param([string]$Action) $uso = Join-Path $env:SystemRoot 'System32\UsoClient.exe'; if (Test-Path -LiteralPath $uso) { Start-Process -FilePath $uso -ArgumentList $Action -WindowStyle Hidden -ErrorAction SilentlyContinue; Write-Log "UsoClient $Action triggered." } }

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    New-Item -Path 'C:\ProgramData\SmartM365\IntuneRemediation\Logs' -ItemType Directory -Force | Out-Null
    Write-Log 'Remediation started.'

    $service = Get-Service -Name 'IntuneManagementExtension' -ErrorAction Stop
    Set-Service -Name 'IntuneManagementExtension' -StartupType Automatic -ErrorAction SilentlyContinue
    if ($service.Status -eq 'Running') {
        Restart-Service -Name 'IntuneManagementExtension' -Force -ErrorAction Stop
        Write-Log 'Intune Management Extension service restarted.'
    }
    else {
        Start-Service -Name 'IntuneManagementExtension' -ErrorAction Stop
        Write-Log 'Intune Management Extension service started.'
    }

    Get-ScheduledTask -TaskPath '\Microsoft\Windows\EnterpriseMgmt\*' -ErrorAction SilentlyContinue |
        Where-Object { $_.State -ne 'Disabled' -and ($_.TaskName -like '*PushLaunch*' -or $_.TaskName -like '*Schedule*') } |
        ForEach-Object {
            try { Start-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -ErrorAction Stop; Write-Log "Started scheduled task: $($_.TaskPath)$($_.TaskName)" } catch { Write-Log "Could not start scheduled task $($_.TaskName): $($_.Exception.Message)" }
        }

    Start-UsoClient -Action 'RefreshSettings'
    Start-UsoClient -Action 'StartScan'

    Write-Log 'Remediation completed.'
    exit 0
}
catch {
    Write-Log "Remediation failed: $($_.Exception.Message)"
    exit 1
}

