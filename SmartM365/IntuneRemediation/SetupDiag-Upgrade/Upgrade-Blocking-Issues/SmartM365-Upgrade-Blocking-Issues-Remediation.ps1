<#
.SYNOPSIS
    Version: 1.0
    Repairs common Windows Update and feature update blockers for Upgrade-Blocking-Issues.
.DESCRIPTION
    Removes legacy WSUS policy blockers, refreshes Windows Update services, clears stale download cache, runs SetupDiag when available, and starts a fresh Windows Update scan/download cycle. Hardware, driver, or application compatibility blockers still require targeted operational remediation.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$Scenario = 'Upgrade-Blocking-Issues'
$LogRoot = Join-Path $env:ProgramData "SmartM365\IntuneRemediation\Logs\$Scenario"
$LogFile = Join-Path $LogRoot "$Scenario-Remediation.log"
$WuPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
$AuPolicyPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU'
$DownloadCachePath = Join-Path $env:SystemRoot 'SoftwareDistribution\Download'
function Write-Log { param([string]$Message) $line = '{0} [{1}] {2}' -f (Get-Date -Format 's'), $Scenario, $Message; Write-Output $line; Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8 }
function Remove-RegistryValueSafe { param([string]$Path,[string]$Name) if (Test-Path -Path $Path) { $item = Get-ItemProperty -Path $Path -ErrorAction SilentlyContinue; if ($item -and $item.PSObject.Properties.Name -contains $Name) { Remove-ItemProperty -Path $Path -Name $Name -Force -ErrorAction SilentlyContinue; Write-Log "Removed registry value: $Path\$Name" } } }
function Stop-ServiceSafe { param([string]$Name) if (Get-Service -Name $Name -ErrorAction SilentlyContinue) { Stop-Service -Name $Name -Force -ErrorAction SilentlyContinue; Write-Log "Service stopped: $Name" } }
function Start-ServiceSafe { param([string]$Name) if (Get-Service -Name $Name -ErrorAction SilentlyContinue) { Start-Service -Name $Name -ErrorAction SilentlyContinue; Write-Log "Service started: $Name" } }
function Start-UsoClient { param([string]$Action) $uso = Join-Path $env:SystemRoot 'System32\UsoClient.exe'; if (Test-Path -LiteralPath $uso) { Start-Process -FilePath $uso -ArgumentList $Action -WindowStyle Hidden -ErrorAction SilentlyContinue; Write-Log "UsoClient $Action triggered." } }

try {
    New-Item -Path $LogRoot -ItemType Directory -Force | Out-Null
    Write-Log 'Remediation started.'
    foreach ($name in @('WUServer','WUStatusServer','UpdateServiceUrlAlternate','DoNotConnectToWindowsUpdateInternetLocations','DisableWindowsUpdateAccess','SetDisableUXWUAccess')) { Remove-RegistryValueSafe -Path $WuPolicyPath -Name $name }
    foreach ($name in @('UseWUServer','NoAutoUpdate','AUOptions')) { Remove-RegistryValueSafe -Path $AuPolicyPath -Name $name }
    foreach ($service in @('bits','wuauserv','dosvc')) { Stop-ServiceSafe -Name $service }
    if (Test-Path -LiteralPath $DownloadCachePath) { Get-ChildItem -LiteralPath $DownloadCachePath -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue; Write-Log 'Windows Update download cache cleared.' }
    foreach ($service in @('cryptsvc','dosvc','wuauserv','bits')) { Start-ServiceSafe -Name $service }
    $setupDiag = Join-Path $env:ProgramFiles 'SetupDiag\SetupDiag.exe'
    if (Test-Path -LiteralPath $setupDiag) {
        $outputRoot = Join-Path $env:SystemRoot 'Logs\SetupDiag'
        New-Item -Path $outputRoot -ItemType Directory -Force | Out-Null
        $process = Start-Process -FilePath $setupDiag -ArgumentList "/Output:$outputRoot" -WindowStyle Hidden -PassThru -ErrorAction SilentlyContinue
        if ($process) { $process.WaitForExit(300000) | Out-Null; Write-Log "SetupDiag executed. ExitCode=$($process.ExitCode)" }
    }
    else { Write-Log 'SetupDiag.exe not found; analysis step skipped.' }
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

