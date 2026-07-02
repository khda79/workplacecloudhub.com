<#
.SYNOPSIS
    Installs the SmartM365 Windows 11 Upgrade Toolkit Intune endpoint package.
.DESCRIPTION
    Copies the packaged setup media and endpoint scripts to ProgramData, registers package detection state, and starts a SYSTEM scheduled task for asynchronous upgrade execution.
.VERSION
    1.0.0
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>
[CmdletBinding()]
param(
    [string]$DataRoot = 'C:\ProgramData\SmartM365\Windows11UpgradeToolkit',
    [string]$TaskName = 'SmartM365 Windows 11 Upgrade Toolkit - Intune',
    [switch]$Uninstall
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$packageRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$manifestPath = Join-Path $packageRoot 'PackageManifest.json'
if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "PackageManifest.json not found: $manifestPath" }
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json

$intuneRoot = Join-Path $DataRoot 'Intune'
$logRoot = Join-Path $DataRoot 'Logs\Intune'
$setupCacheRoot = Join-Path $DataRoot 'SetupMedia'
$registryRoot = 'HKLM:\SOFTWARE\SmartM365\Windows11UpgradeToolkit\IntunePackages'
$registryPath = Join-Path $registryRoot ([string]$manifest.PackageId)

function New-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}

function New-RegistryKey {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { New-Item -Path $Path -Force | Out-Null }
}

function Set-RegistryString {
    param(
        [string]$Path,
        [string]$Name,
        [string]$Value
    )
    New-ItemProperty -LiteralPath $Path -Name $Name -Value $Value -PropertyType String -Force | Out-Null
}

function Write-InstallLog {
    param([string]$Message, [string]$Level = 'INFO')
    New-Directory -Path $logRoot
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
    Add-Content -LiteralPath (Join-Path $logRoot 'Install.log') -Value $line -Encoding UTF8
}

if ($Uninstall) {
    Write-InstallLog "Uninstall requested for $($manifest.PackageId)."
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    if (Test-Path -LiteralPath $registryPath) { Remove-Item -LiteralPath $registryPath -Recurse -Force -ErrorAction SilentlyContinue }
    exit 0
}

Write-InstallLog "Installing $($manifest.PackageId) version $($manifest.PackageVersion)."
New-Directory -Path $DataRoot
New-Directory -Path $intuneRoot
New-Directory -Path $setupCacheRoot
New-Directory -Path $logRoot

Copy-Item -LiteralPath (Join-Path $packageRoot 'SmartM365-Invoke-Windows11UpgradeRepair.ps1') -Destination (Join-Path $DataRoot 'SmartM365-Invoke-Windows11UpgradeRepair.ps1') -Force
Copy-Item -LiteralPath (Join-Path $packageRoot 'Run-IntuneUpgrade.ps1') -Destination (Join-Path $intuneRoot 'Run-IntuneUpgrade.ps1') -Force
Copy-Item -LiteralPath $manifestPath -Destination (Join-Path $intuneRoot 'PackageManifest.json') -Force

$packageMediaRoot = Join-Path $packageRoot ("SetupMedia\{0}" -f $manifest.SetupCacheFolder)
$targetMediaRoot = Join-Path $setupCacheRoot ([string]$manifest.SetupCacheFolder)
if (-not (Test-Path -LiteralPath (Join-Path $packageMediaRoot 'setup.exe') -PathType Leaf)) { throw "Packaged setup.exe not found: $packageMediaRoot" }

Write-InstallLog "Copying packaged setup media to local cache: $packageMediaRoot -> $targetMediaRoot"
New-Directory -Path $targetMediaRoot
$robocopy = Join-Path $env:SystemRoot 'System32\robocopy.exe'
$installRobocopyLog = Join-Path $logRoot 'Install-Robocopy.log'
& $robocopy $packageMediaRoot $targetMediaRoot /MIR /R:2 /W:5 /NP /NFL /NDL "/LOG+:$installRobocopyLog" | Out-Null
$copyExit = [int]$LASTEXITCODE
if ($copyExit -gt 7) { throw "Robocopy install media copy failed with exit code $copyExit." }

New-RegistryKey -Path $registryRoot
New-RegistryKey -Path $registryPath
Set-RegistryString -Path $registryPath -Name PackageId -Value ([string]$manifest.PackageId)
Set-RegistryString -Path $registryPath -Name PackageVersion -Value ([string]$manifest.PackageVersion)
Set-RegistryString -Path $registryPath -Name Language -Value ([string]$manifest.Language)
Set-RegistryString -Path $registryPath -Name MediaId -Value ([string]$manifest.MediaId)
Set-RegistryString -Path $registryPath -Name SetupCacheFolder -Value ([string]$manifest.SetupCacheFolder)
Set-RegistryString -Path $registryPath -Name InstalledUtc -Value ((Get-Date).ToUniversalTime().ToString('o'))

$runner = Join-Path $intuneRoot 'Run-IntuneUpgrade.ps1'
$taskArgument = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -DataRoot "{1}" -TaskName "{2}"' -f $runner,$DataRoot,$TaskName
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $taskArgument
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(5) -RepetitionInterval (New-TimeSpan -Hours 2) -RepetitionDuration (New-TimeSpan -Days 30)
$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
Start-ScheduledTask -TaskName $TaskName

Write-InstallLog "Install completed and scheduled task started: $TaskName"
exit 0
