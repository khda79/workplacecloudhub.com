<#
.SYNOPSIS
    Installs the SmartM365 Windows 11 Upgrade Toolkit Intune endpoint package.
.DESCRIPTION
    Copies endpoint scripts to ProgramData, optionally copies packaged setup media or validates an existing cache, registers package detection state, and starts a SYSTEM scheduled task for asynchronous upgrade execution.
.VERSION
    1.0.6
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
$registrySubKeyRoot = 'SOFTWARE\SmartM365\Windows11UpgradeToolkit\IntunePackages'
$registrySubKey = "$registrySubKeyRoot\$([string]$manifest.PackageId)"
$packageMode = 'WithMedia'
if ($manifest.PSObject.Properties['PackageMode'] -and -not [string]::IsNullOrWhiteSpace([string]$manifest.PackageMode)) { $packageMode = [string]$manifest.PackageMode }
$requiresExistingSetupCache = $false
if ($manifest.PSObject.Properties['RequiresExistingSetupCache']) { $requiresExistingSetupCache = [bool]$manifest.RequiresExistingSetupCache }
$packageMediaRoot = Join-Path $packageRoot ("SetupMedia\{0}" -f $manifest.SetupCacheFolder)
$targetMediaRoot = Join-Path $setupCacheRoot ([string]$manifest.SetupCacheFolder)

function New-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}

function Open-Registry64LocalMachine {
    return [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
}

function Set-Registry64String {
    param(
        [string]$SubKey,
        [string]$Name,
        [string]$Value
    )

    $baseKey = Open-Registry64LocalMachine
    try {
        $key = $baseKey.CreateSubKey($SubKey)
        try { $key.SetValue($Name, [string]$Value, [Microsoft.Win32.RegistryValueKind]::String) }
        finally { if ($key) { $key.Dispose() } }
    }
    finally { $baseKey.Dispose() }
}

function Remove-Registry64SubKeyTree {
    param([string]$SubKey)

    $baseKey = Open-Registry64LocalMachine
    try { $baseKey.DeleteSubKeyTree($SubKey, $false) }
    finally { $baseKey.Dispose() }
}

function Write-InstallLog {
    param([string]$Message, [string]$Level = 'INFO')
    New-Directory -Path $logRoot
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
    Add-Content -LiteralPath (Join-Path $logRoot 'Install.log') -Value $line -Encoding UTF8
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

function Get-DirectorySizeBytes {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return [int64]0 }
    $sum = [int64]0
    foreach ($file in @(Get-ChildItem -LiteralPath $Path -File -Recurse -ErrorAction SilentlyContinue)) { $sum += [int64]$file.Length }
    return $sum
}

function Get-SetupCacheLockName {
    param([Parameter(Mandatory = $true)][string]$CachePath)

    $leaf = Split-Path -Path $CachePath -Leaf
    if ([string]::IsNullOrWhiteSpace($leaf)) { $leaf = 'Default' }
    $safe = $leaf -replace '[^A-Za-z0-9._-]', '_'
    return "SetupCache-$safe.lock"
}

function Test-SetupCacheLockProcessAlive {
    param([AllowNull()][object]$ProcessId)

    $pidValue = 0
    if ($null -eq $ProcessId -or -not [int]::TryParse([string]$ProcessId, [ref]$pidValue) -or $pidValue -le 0) { return $false }
    try { [void](Get-Process -Id $pidValue -ErrorAction Stop); return $true } catch { return $false }
}

function New-SetupCacheLockPayload {
    param(
        [Parameter(Mandatory = $true)][string]$CachePath,
        [Parameter(Mandatory = $true)][string]$LockPath,
        [Parameter(Mandatory = $true)][string]$OwnerToken,
        [Parameter(Mandatory = $true)][string]$Purpose,
        [Parameter(Mandatory = $true)][int]$LeaseMinutes
    )

    $nowUtc = (Get-Date).ToUniversalTime()
    $processName = ''
    try { $processName = (Get-Process -Id $PID -ErrorAction Stop).ProcessName } catch { }
    return [pscustomobject]@{
        OwnerToken = $OwnerToken
        Owner = 'SmartM365-Windows11UpgradeToolkit-IntuneInstall'
        ComputerName = $env:COMPUTERNAME
        PID = $PID
        ProcessName = $processName
        Purpose = $Purpose
        CachePath = $CachePath
        LockPath = $LockPath
        StartedUtc = $nowUtc.ToString('o')
        LastSeenUtc = $nowUtc.ToString('o')
        ExpiresUtc = $nowUtc.AddMinutes($LeaseMinutes).ToString('o')
    }
}

function Read-SetupCacheLockPayload {
    param([Parameter(Mandatory = $true)][string]$LockPath)

    $payloadPath = Join-Path $LockPath 'lock.json'
    if (-not (Test-Path -LiteralPath $payloadPath -PathType Leaf)) { return $null }
    try { return (Get-Content -LiteralPath $payloadPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop) } catch { return $null }
}

function Save-SetupCacheLockPayload {
    param(
        [Parameter(Mandatory = $true)][string]$LockPath,
        [Parameter(Mandatory = $true)][object]$Payload
    )

    $payloadPath = Join-Path $LockPath 'lock.json'
    $Payload | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $payloadPath -Encoding UTF8
}

function Format-SetupCacheLockPayload {
    param([AllowNull()][object]$Payload)

    if ($null -eq $Payload) { return 'LockPayload=<missing>' }
    return ("Owner={0}; PID={1}; Process={2}; Purpose={3}; StartedUtc={4}; LastSeenUtc={5}; ExpiresUtc={6}" -f $Payload.Owner,$Payload.PID,$Payload.ProcessName,$Payload.Purpose,$Payload.StartedUtc,$Payload.LastSeenUtc,$Payload.ExpiresUtc)
}

function Acquire-SetupCacheLock {
    param(
        [Parameter(Mandatory = $true)][string]$CachePath,
        [Parameter(Mandatory = $true)][string]$Purpose,
        [int]$LeaseMinutes = 90
    )

    $lockRoot = Join-Path $DataRoot 'Locks'
    New-Directory -Path $lockRoot
    $lockPath = Join-Path $lockRoot (Get-SetupCacheLockName -CachePath $CachePath)
    $ownerToken = [guid]::NewGuid().ToString('N')

    while ($true) {
        try {
            New-Item -ItemType Directory -Path $lockPath -ErrorAction Stop | Out-Null
            $payload = New-SetupCacheLockPayload -CachePath $CachePath -LockPath $lockPath -OwnerToken $ownerToken -Purpose $Purpose -LeaseMinutes $LeaseMinutes
            Save-SetupCacheLockPayload -LockPath $lockPath -Payload $payload
            Write-InstallLog ("Acquired setup cache lock: CachePath={0}; LeaseMinutes={1}; Lock={2}; {3}" -f $CachePath,$LeaseMinutes,$lockPath,(Format-SetupCacheLockPayload -Payload $payload))
            return [pscustomobject]@{ LockPath = $lockPath; OwnerToken = $ownerToken; CachePath = $CachePath }
        }
        catch {
            $payload = Read-SetupCacheLockPayload -LockPath $lockPath
            $nowUtc = (Get-Date).ToUniversalTime()
            if ($null -eq $payload) {
                $lockItem = $null
                try { $lockItem = Get-Item -LiteralPath $lockPath -ErrorAction Stop } catch { }
                if ($null -ne $lockItem -and $lockItem.LastWriteTimeUtc -lt $nowUtc.AddMinutes(-5)) {
                    Write-InstallLog ("Reclaiming empty stale setup cache lock: CachePath={0}; Lock={1}; LastWriteTimeUtc={2}" -f $CachePath,$lockPath,$lockItem.LastWriteTimeUtc.ToString('o')) 'WARN'
                    Remove-Item -LiteralPath $lockPath -Recurse -Force -ErrorAction SilentlyContinue
                    continue
                }
                throw ("Setup cache is locked by another SmartM365 process. CachePath={0}; Lock={1}; {2}" -f $CachePath,$lockPath,(Format-SetupCacheLockPayload -Payload $payload))
            }
            $expiresUtc = [datetime]::MinValue
            $expired = $true
            if ($null -ne $payload -and [datetime]::TryParse([string]$payload.ExpiresUtc, [ref]$expiresUtc)) { $expired = ($expiresUtc.ToUniversalTime() -lt $nowUtc) }
            $processAlive = if ($null -ne $payload) { Test-SetupCacheLockProcessAlive -ProcessId $payload.PID } else { $false }

            if ($expired -and -not $processAlive) {
                Write-InstallLog ("Reclaiming stale setup cache lock: CachePath={0}; Lock={1}; {2}" -f $CachePath,$lockPath,(Format-SetupCacheLockPayload -Payload $payload)) 'WARN'
                Remove-Item -LiteralPath $lockPath -Recurse -Force -ErrorAction SilentlyContinue
                continue
            }

            throw ("Setup cache is locked by another SmartM365 process. CachePath={0}; Lock={1}; {2}" -f $CachePath,$lockPath,(Format-SetupCacheLockPayload -Payload $payload))
        }
    }
}

function Release-SetupCacheLock {
    param([AllowNull()][object]$Lock)

    if ($null -eq $Lock) { return }
    $payload = Read-SetupCacheLockPayload -LockPath $Lock.LockPath
    if ($null -ne $payload -and [string]$payload.OwnerToken -ne [string]$Lock.OwnerToken) {
        Write-InstallLog ("Skipping setup cache lock release because owner changed. Lock={0}; {1}" -f $Lock.LockPath,(Format-SetupCacheLockPayload -Payload $payload)) 'WARN'
        return
    }
    Remove-Item -LiteralPath $Lock.LockPath -Recurse -Force -ErrorAction SilentlyContinue
    Write-InstallLog ("Released setup cache lock: CachePath={0}; Lock={1}" -f $Lock.CachePath,$Lock.LockPath)
}
function Set-PackageDetectionState {
    param([string]$InstallState = 'Installed')

    Write-InstallLog "Writing Intune detection registry state to HKLM:\$registrySubKey (64-bit registry view). InstallState=$InstallState"
    Set-Registry64String -SubKey $registrySubKey -Name PackageId -Value ([string]$manifest.PackageId)
    Set-Registry64String -SubKey $registrySubKey -Name PackageVersion -Value ([string]$manifest.PackageVersion)
    Set-Registry64String -SubKey $registrySubKey -Name Language -Value ([string]$manifest.Language)
    Set-Registry64String -SubKey $registrySubKey -Name MediaId -Value ([string]$manifest.MediaId)
    Set-Registry64String -SubKey $registrySubKey -Name SetupCacheFolder -Value ([string]$manifest.SetupCacheFolder)
    Set-Registry64String -SubKey $registrySubKey -Name PackageMode -Value $packageMode
    Set-Registry64String -SubKey $registrySubKey -Name InstallState -Value $InstallState
    Set-Registry64String -SubKey $registrySubKey -Name InstalledUtc -Value ((Get-Date).ToUniversalTime().ToString('o'))
}

function Test-SetupCacheReady {
    param([Parameter(Mandatory = $true)][string]$Path)

    $setupExe = Join-Path $Path 'setup.exe'
    $installWim = Join-Path $Path 'sources\install.wim'
    if (-not (Test-Path -LiteralPath $setupExe -PathType Leaf)) { throw "Local setup cache is missing setup.exe: $Path" }
    if (-not (Test-Path -LiteralPath $installWim -PathType Leaf)) { throw "Local setup cache is missing sources\install.wim: $Path" }
    return $true
}

if ($Uninstall) {
    Write-InstallLog "Uninstall requested for $($manifest.PackageId)."
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    Remove-Registry64SubKeyTree -SubKey $registrySubKey
    exit 0
}

Write-InstallLog "Installing $($manifest.PackageId) version $($manifest.PackageVersion)."
New-Directory -Path $DataRoot
New-Directory -Path $intuneRoot
New-Directory -Path $setupCacheRoot
New-Directory -Path $logRoot

if ((Get-OsFamily) -eq 'Windows11') {
    Write-InstallLog 'Device is already Windows 11 during package install. Writing detection state, removing scheduled task if present, and exiting success.'
    Set-PackageDetectionState -InstallState 'AlreadyWindows11'
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    exit 0
}

Copy-Item -LiteralPath (Join-Path $packageRoot 'SmartM365-Invoke-Windows11UpgradeRepair.ps1') -Destination (Join-Path $DataRoot 'SmartM365-Invoke-Windows11UpgradeRepair.ps1') -Force
Copy-Item -LiteralPath (Join-Path $packageRoot 'Run-IntuneUpgrade.ps1') -Destination (Join-Path $intuneRoot 'Run-IntuneUpgrade.ps1') -Force
Copy-Item -LiteralPath $manifestPath -Destination (Join-Path $intuneRoot 'PackageManifest.json') -Force

if ($requiresExistingSetupCache) {
    Write-InstallLog "Cache-only package mode enabled. Validating existing setup cache: $targetMediaRoot"
    [void](Test-SetupCacheReady -Path $targetMediaRoot)
}
else {
    if (-not (Test-Path -LiteralPath (Join-Path $packageMediaRoot 'setup.exe') -PathType Leaf)) { throw "Packaged setup.exe not found: $packageMediaRoot" }
    if (-not (Test-Path -LiteralPath (Join-Path $packageMediaRoot 'sources\install.wim') -PathType Leaf)) { throw "Packaged sources\install.wim not found: $packageMediaRoot" }

    $cacheLock = $null
    try {
        $cacheLock = Acquire-SetupCacheLock -CachePath $targetMediaRoot -Purpose ("IntunePackageInstall:{0}" -f $manifest.PackageId) -LeaseMinutes 90
        $targetCacheReady = $false
        try {
            [void](Test-SetupCacheReady -Path $targetMediaRoot)
            $packageBytes = Get-DirectorySizeBytes -Path $packageMediaRoot
            $targetBytes = Get-DirectorySizeBytes -Path $targetMediaRoot
            if ($packageBytes -gt 0 -and $targetBytes -ge $packageBytes) { $targetCacheReady = $true }
            else { Write-InstallLog ("Existing setup cache is present but smaller than package source. PackageBytes={0}; TargetBytes={1}; Cache={2}" -f $packageBytes,$targetBytes,$targetMediaRoot) 'WARN' }
        }
        catch {
            Write-InstallLog ("Existing setup cache is not ready and will be refreshed. Cache={0}; Reason={1}" -f $targetMediaRoot,$_.Exception.Message) 'WARN'
        }

        if ($targetCacheReady) {
            Write-InstallLog "Existing setup cache already looks ready. Skipping media recopy: $targetMediaRoot"
        }
        else {
            if (Test-Path -LiteralPath $targetMediaRoot -PathType Container) {
                Write-InstallLog "Removing incomplete setup cache before recopy: $targetMediaRoot"
                Remove-Item -LiteralPath $targetMediaRoot -Recurse -Force -ErrorAction Stop
            }
            Write-InstallLog "Copying packaged setup media to local cache: $packageMediaRoot -> $targetMediaRoot"
            New-Directory -Path $targetMediaRoot
            $robocopy = Join-Path $env:SystemRoot 'System32\robocopy.exe'
            $installRobocopyLog = Join-Path $logRoot 'Install-Robocopy.log'
            & $robocopy $packageMediaRoot $targetMediaRoot /MIR /R:2 /W:5 /NP /NFL /NDL "/LOG+:$installRobocopyLog" | Out-Null
            $copyExit = [int]$LASTEXITCODE
            Write-InstallLog "Robocopy setup media copy exit code: $copyExit; Log=$installRobocopyLog"
            if ($copyExit -gt 7) { throw "Robocopy install media copy failed with exit code $copyExit. Log=$installRobocopyLog" }
            [void](Test-SetupCacheReady -Path $targetMediaRoot)
        }
    }
    catch {
        if ($_.Exception.Message -like 'Setup cache is locked by another SmartM365 process.*') {
            Write-InstallLog ("Setup cache is locked by another SmartM365 process. Intune will retry. Detail={0}" -f $_.Exception.Message) 'WARN'
            exit 1618
        }
        throw
    }
    finally {
        Release-SetupCacheLock -Lock $cacheLock
    }
}
Set-PackageDetectionState -InstallState 'Installed'

$runner = Join-Path $intuneRoot 'Run-IntuneUpgrade.ps1'
$runnerLog = Join-Path $logRoot 'Run-IntuneUpgrade.log'
$taskArgument = '-NoProfile -ExecutionPolicy Bypass -File "{0}" -DataRoot "{1}" -TaskName "{2}"' -f $runner,$DataRoot,$TaskName
$action = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $taskArgument
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddHours(2) -RepetitionInterval (New-TimeSpan -Hours 2) -RepetitionDuration (New-TimeSpan -Days 30)
$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
Write-InstallLog "Scheduled task action: powershell.exe $taskArgument"
Write-InstallLog "Runner log expected at: $runnerLog"
Start-ScheduledTask -TaskName $TaskName

Write-InstallLog "Install completed and scheduled task started: $TaskName; next scheduled retry is every 2 hours."
exit 0
