<#
.SYNOPSIS
    Version: 1.0
    Repairs stale Windows Update download cache for Download-Failure.
.DESCRIPTION
    Stops Windows Update related services, clears the SoftwareDistribution Download cache, restarts services, refreshes update settings, and triggers a fresh scan and download.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$Scenario = 'Download-Failure'
$LogRoot = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path $LogRoot "$Scenario-Remediation.log"
$DownloadCachePath = Join-Path $env:SystemRoot 'SoftwareDistribution\Download'
function Write-Log { param([string]$Message) $line = '{0} [{1}] {2}' -f (Get-Date -Format 's'), $Scenario, $Message; Write-Output $line; Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8 }
function Stop-ServiceSafe { param([string]$Name) if (Get-Service -Name $Name -ErrorAction SilentlyContinue) { Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue; Write-Log "Service stopped: $Name" } }
function Start-ServiceSafe { param([string]$Name) if (Get-Service -Name $Name -ErrorAction SilentlyContinue) { Start-Service -Name $Name -ErrorAction SilentlyContinue; Write-Log "Service started: $Name" } }
function Start-UsoClient { param([string]$Action) $uso = Join-Path $env:SystemRoot 'System32\UsoClient.exe'; if (Test-Path -LiteralPath $uso) { Start-Process -FilePath $uso -ArgumentList $Action -WindowStyle Hidden -ErrorAction SilentlyContinue; Write-Log "UsoClient $Action triggered." } }

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-Log 'Remediation started.'
    foreach ($service in @('bits','wuauserv','dosvc')) { Stop-ServiceSafe -Name $service }
    if (Test-Path -LiteralPath $DownloadCachePath) {
        Get-ChildItem -LiteralPath $DownloadCachePath -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Cleared Windows Update download cache: $DownloadCachePath"
    }
    foreach ($service in @('cryptsvc','dosvc','wuauserv','bits')) { Start-ServiceSafe -Name $service }
    Start-UsoClient -Action 'RefreshSettings'
    Start-UsoClient -Action 'StartScan'
    Start-UsoClient -Action 'StartDownload'
    Write-Log 'Remediation completed.'
    exit 0
}
catch {
    Write-Log "Remediation failed: $($_.Exception.Message)"
    exit 1
}

