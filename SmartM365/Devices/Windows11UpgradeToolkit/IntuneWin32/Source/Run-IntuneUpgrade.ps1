<#
.SYNOPSIS
    Runs the SmartM365 Windows 11 Upgrade Toolkit endpoint from an Intune scheduled task.
.DESCRIPTION
    Executes the local endpoint script against the packaged setup media cache and removes the scheduled task once the device is already Windows 11.
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>
[CmdletBinding()]
param(
    [string]$DataRoot = 'C:\ProgramData\SmartM365\Windows11UpgradeToolkit',
    [string]$TaskName = 'SmartM365 Windows 11 Upgrade Toolkit - Intune',
    [int]$RunGuardHours = 2
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$intuneRoot = Join-Path $DataRoot 'Intune'
$logRoot = Join-Path $DataRoot 'Logs\Intune'
$manifestPath = Join-Path $intuneRoot 'PackageManifest.json'
$endpointScript = Join-Path $DataRoot 'SmartM365-Invoke-Windows11UpgradeRepair.ps1'
$setupCacheRoot = Join-Path $DataRoot 'SetupMedia'

function New-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}

function Write-RunnerLog {
    param([string]$Message, [string]$Level = 'INFO')
    New-Directory -Path $logRoot
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
    Add-Content -LiteralPath (Join-Path $logRoot 'Run-IntuneUpgrade.log') -Value $line -Encoding UTF8
}

function Get-OsFamily {
    try {
        $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
        $build = [int]$os.BuildNumber
        if ($build -ge 22000 -or ([string]$os.Caption) -match 'Windows 11') { return 'Windows11' }
        if ([string]$os.Caption -match 'Windows 10') { return 'Windows10' }
        return 'Other'
    }
    catch { return 'Unknown' }
}

if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "Package manifest not found: $manifestPath" }
if (-not (Test-Path -LiteralPath $endpointScript -PathType Leaf)) { throw "Endpoint script not found: $endpointScript" }
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json

if ((Get-OsFamily) -eq 'Windows11') {
    Write-RunnerLog 'Device is already Windows 11. Removing scheduled task.'
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    exit 0
}

$setupProcesses = @(Get-Process -Name setup,setuphost,setupprep -ErrorAction SilentlyContinue)
if ($setupProcesses.Count -gt 0) {
    Write-RunnerLog ("Setup process already running; skipping this cycle. Processes={0}" -f (($setupProcesses | Select-Object -ExpandProperty Id) -join ','))
    exit 0
}

$cacheFolder = [string]$manifest.SetupCacheFolder
$cachePath = Join-Path $setupCacheRoot $cacheFolder
if (-not (Test-Path -LiteralPath (Join-Path $cachePath 'setup.exe') -PathType Leaf)) { throw "Local packaged setup cache is missing setup.exe: $cachePath" }

$args = @(
    '-RunGuardHours', [string]$RunGuardHours,
    '-DirectSetupUpgrade',
    '-SkipVirtualMachines',
    '-AllowReboot',
    '-AllowDiskCleanup',
    '-SetupExecutionMode', 'LocalCache',
    '-SetupMediaId', [string]$manifest.MediaId,
    '-SetupLanguage', [string]$manifest.Language,
    '-SetupDynamicUpdate', 'Disable',
    '-SetupCacheRoot', $setupCacheRoot,
    '-SetupProcessHeartbeatSeconds', '300',
    '-SetupProcessTimeoutMinutes', '0',
    '-SetupMediaCopyTimeoutMinutes', '180',
    '-DataRoot', $DataRoot
)

Write-RunnerLog ("Starting endpoint script for package {0}; Language={1}; Cache={2}" -f $manifest.PackageId,$manifest.Language,$cachePath)
& $endpointScript @args
$exit = [int]$LASTEXITCODE
Write-RunnerLog ("Endpoint script exited with code {0}." -f $exit)

if ((Get-OsFamily) -eq 'Windows11') {
    Write-RunnerLog 'Device is Windows 11 after endpoint run. Removing scheduled task.'
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch { }
}

exit 0
