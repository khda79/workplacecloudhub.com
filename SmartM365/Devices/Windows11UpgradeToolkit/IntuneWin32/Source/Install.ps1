<#
.SYNOPSIS
    Installs the SmartM365 Windows 11 Upgrade Toolkit Intune endpoint package.
.DESCRIPTION
    Copies endpoint scripts to ProgramData, optionally copies packaged setup media or validates an existing cache, registers package detection state, and starts a SYSTEM scheduled task for asynchronous upgrade execution.
.VERSION
    1.0.7
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

function Protect-ToolkitPathAcl {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [switch]$Directory
    )

    if (-not (Test-Path -LiteralPath $Path)) { return }
    $icacls = Join-Path $env:SystemRoot 'System32\icacls.exe'
    if (-not (Test-Path -LiteralPath $icacls -PathType Leaf)) {
        Write-InstallLog ("ACL hardening skipped because icacls.exe was not found. Path={0}" -f $Path) 'WARN'
        return
    }

    $grants = if ($Directory) {
        @('*S-1-5-18:(OI)(CI)F','*S-1-5-32-544:(OI)(CI)F','*S-1-5-32-545:(OI)(CI)RX')
    }
    else {
        @('*S-1-5-18:F','*S-1-5-32-544:F','*S-1-5-32-545:RX')
    }

    $output = & $icacls $Path '/inheritance:r' '/grant:r' $grants 2>&1
    $exitCode = [int]$LASTEXITCODE
    if ($exitCode -ne 0) {
        Write-InstallLog ("ACL hardening failed. Path={0}; ExitCode={1}; Output={2}" -f $Path,$exitCode,(($output | Out-String).Trim())) 'WARN'
        return
    }

    Write-InstallLog ("ACL hardening applied. Path={0}; Directory={1}" -f $Path,[bool]$Directory)
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
Protect-ToolkitPathAcl -Path $DataRoot -Directory
Protect-ToolkitPathAcl -Path $intuneRoot -Directory
Protect-ToolkitPathAcl -Path $setupCacheRoot -Directory
Protect-ToolkitPathAcl -Path $logRoot -Directory

if ((Get-OsFamily) -eq 'Windows11') {
    Write-InstallLog 'Device is already Windows 11 during package install. Writing detection state, removing scheduled task if present, and exiting success.'
    Set-PackageDetectionState -InstallState 'AlreadyWindows11'
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    exit 0
}

Copy-Item -LiteralPath (Join-Path $packageRoot 'SmartM365-Invoke-Windows11UpgradeRepair.ps1') -Destination (Join-Path $DataRoot 'SmartM365-Invoke-Windows11UpgradeRepair.ps1') -Force
Copy-Item -LiteralPath (Join-Path $packageRoot 'Run-IntuneUpgrade.ps1') -Destination (Join-Path $intuneRoot 'Run-IntuneUpgrade.ps1') -Force
Copy-Item -LiteralPath $manifestPath -Destination (Join-Path $intuneRoot 'PackageManifest.json') -Force
Protect-ToolkitPathAcl -Path (Join-Path $DataRoot 'SmartM365-Invoke-Windows11UpgradeRepair.ps1')
Protect-ToolkitPathAcl -Path (Join-Path $intuneRoot 'Run-IntuneUpgrade.ps1')
Protect-ToolkitPathAcl -Path (Join-Path $intuneRoot 'PackageManifest.json')

if ($requiresExistingSetupCache) {
    Write-InstallLog "Cache-only package mode enabled. Validating existing setup cache: $targetMediaRoot"
    [void](Test-SetupCacheReady -Path $targetMediaRoot)
    Protect-ToolkitPathAcl -Path $targetMediaRoot -Directory
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
        Protect-ToolkitPathAcl -Path $targetMediaRoot -Directory
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
$powerShellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$action = New-ScheduledTaskAction -Execute $powerShellExe -Argument $taskArgument
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddHours(2) -RepetitionInterval (New-TimeSpan -Hours 2) -RepetitionDuration (New-TimeSpan -Days 30)
$principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force | Out-Null
Write-InstallLog ("Scheduled task action: {0} {1}" -f $powerShellExe,$taskArgument)
Write-InstallLog "Runner log expected at: $runnerLog"
Start-ScheduledTask -TaskName $TaskName

Write-InstallLog "Install completed and scheduled task started: $TaskName; next scheduled retry is every 2 hours."
exit 0

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDBFisbsnNc3Kyl
# b1pRcvawAyVOR9VWnV2eQnXJHGKEUaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
# s0Q4yPEDH+JoMA0GCSqGSIb3DQEBCwUAME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91
# ZGh1Yi5jb20wHhcNMjYwNzEzMDgyMjM1WhcNMjkwNzEzMDgzMjI5WjBOMR4wHAYD
# VQQDDBV3b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRh
# Y3RAd29ya3BsYWNlY2xvdWRodWIuY29tMIIBojANBgkqhkiG9w0BAQEFAAOCAY8A
# MIIBigKCAYEAse6XztERSyHn9DVqj8Rdv0qjc5owqvgAIGaYxBmfiQuoM48Fo4Xt
# 1ovi9brLUtf55G4XgthNPCoanxfCRRg30IVRxaDfdPXJzYmgsM5tXlsuNU49lE7E
# PJk3+jEOgSCt8NKzmVPKpNRG0NmK0a8wm12cceYZOZlSYE0+ZtT6wy5PQQjMUqIx
# XnGjt4H0nfgZZa7D4FyARKOVg/Xr9sUq5jIn3zszvg4jjeb4b0DKJtfbHukhWc2Y
# oVFgswxVBXCWIaBnfF/cjqMfK/CaToT2trVb4hG4qcQ31s1nR4keoRaOw/vyd6ap
# rEtCsT22N/Jx0dz7fIo1tVyvIaVcHdN9LW3chn0en0OKZ6Ke1OH9wf2prl4KA6Ww
# VzrAZrOlXTAItdK7D9kKO/HeJd4PZvO53oy1LdmMGLSz3OLB9e5q7yo8rfqi5Ka9
# KzM2CrSzz1yphn/H90wz7Q2pm4FIlWdcj86A/0kmhYg+5Wqqbg1drrPXu4nEBwWN
# /dzoGtKZKHTdAgMBAAGjgZYwgZMwDgYDVR0PAQH/BAQDAgeAMBMGA1UdJQQMMAoG
# CCsGAQUFBwMDMD8GA1UdEQQ4MDaBHWNvbnRhY3RAd29ya3BsYWNlY2xvdWRodWIu
# Y29tghV3b3JrcGxhY2VjbG91ZGh1Yi5jb20wDAYDVR0TAQH/BAIwADAdBgNVHQ4E
# FgQUXIOOADQM78XfPAncirgCECedg9gwDQYJKoZIhvcNAQELBQADggGBADhZUB2R
# 5J/Jw030xodhEWeCQ0vnJRaiEsjOxuArQREKH3lCrQ3UsUVl292d6LnQUSTH/jF7
# rovEZ+JN2GQ/LCrXRaCuwCEGZKzlSEbtYWhfwDyj6GpIPq8Y4SeXyjdq4/rrI1bm
# iTK4Sq7EoBlGJuX6l2nfvx1tTioSr11FoDfllJR7EYawRj9hBFJ0gG0b2SuYZMgW
# gaDKefcnJDmOwcRNAZUII0ss8EeyANukWSkNN5ILZ+iKDpQgZxgDLPTiRguCyx45
# PI5wrVTjV/pR7IrtSIfq8UladlrSZJyyDn3NV2ATvIZ6wNxbTmPFcE0uMg/EYzwd
# Tek+CgXL3TxUKeldJM4YDWPimNBRhOPXzBDiOQIj6WNswt/KM1oDLnA00CNtciPN
# dn+dXlneMvTEUah9wyt8o8tkLpoBw+KN+Bq/K0O1qPtS7umi70l45pPiej+mwbwq
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBY0w
# ggR1oAMCAQICEA6bGI750C3n79tQ4ghAGFowDQYJKoZIhvcNAQEMBQAwZTELMAkG
# A1UEBhMCVVMxFTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRp
# Z2ljZXJ0LmNvbTEkMCIGA1UEAxMbRGlnaUNlcnQgQXNzdXJlZCBJRCBSb290IENB
# MB4XDTIyMDgwMTAwMDAwMFoXDTMxMTEwOTIzNTk1OVowYjELMAkGA1UEBhMCVVMx
# FTATBgNVBAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNv
# bTEhMB8GA1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MIICIjANBgkqhkiG
# 9w0BAQEFAAOCAg8AMIICCgKCAgEAv+aQc2jeu+RdSjwwIjBpM+zCpyUuySE98orY
# WcLhKac9WKt2ms2uexuEDcQwH/MbpDgW61bGl20dq7J58soR0uRf1gU8Ug9SH8ae
# FaV+vp+pVxZZVXKvaJNwwrK6dZlqczKU0RBEEC7fgvMHhOZ0O21x4i0MG+4g1ckg
# HWMpLc7sXk7Ik/ghYZs06wXGXuxbGrzryc/NrDRAX7F6Zu53yEioZldXn1RYjgwr
# t0+nMNlW7sp7XeOtyU9e5TXnMcvak17cjo+A2raRmECQecN4x7axxLVqGDgDEI3Y
# 1DekLgV9iPWCPhCRcKtVgkEy19sEcypukQF8IUzUvK4bA3VdeGbZOjFEmjNAvwjX
# WkmkwuapoGfdpCe8oU85tRFYF/ckXEaPZPfBaYh2mHY9WV1CdoeJl2l6SPDgohIb
# Zpp0yt5LHucOY67m1O+SkjqePdwA5EUlibaaRBkrfsCUtNJhbesz2cXfSwQAzH0c
# lcOP9yGyshG3u3/y1YxwLEFgqrFjGESVGnZifvaAsPvoZKYz0YkH4b235kOkGLim
# dwHhD5QMIR2yVCkliWzlDlJRR3S+Jqy2QXXeeqxfjT/JvNNBERJb5RBQ6zHFynIW
# IgnffEx1P2PsIV/EIFFrb7GrhotPwtZFX50g/KEexcCPorF+CiaZ9eRpL5gdLfXZ
# qbId5RsCAwEAAaOCATowggE2MA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFOzX
# 44LScV1kTN8uZz/nupiuHA9PMB8GA1UdIwQYMBaAFEXroq/0ksuCMS1Ri6enIZ3z
# bcgPMA4GA1UdDwEB/wQEAwIBhjB5BggrBgEFBQcBAQRtMGswJAYIKwYBBQUHMAGG
# GGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBDBggrBgEFBQcwAoY3aHR0cDovL2Nh
# Y2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0QXNzdXJlZElEUm9vdENBLmNydDBF
# BgNVHR8EPjA8MDqgOKA2hjRodHRwOi8vY3JsMy5kaWdpY2VydC5jb20vRGlnaUNl
# cnRBc3N1cmVkSURSb290Q0EuY3JsMBEGA1UdIAQKMAgwBgYEVR0gADANBgkqhkiG
# 9w0BAQwFAAOCAQEAcKC/Q1xV5zhfoKN0Gz22Ftf3v1cHvZqsoYcs7IVeqRq7IviH
# GmlUIu2kiHdtvRoU9BNKei8ttzjv9P+Aufih9/Jy3iS8UgPITtAq3votVs/59Pes
# MHqai7Je1M/RQ0SbQyHrlnKhSLSZy51PpwYDE3cnRNTnf+hZqPC/Lwum6fI0POz3
# A8eHqNJMQBk1RmppVLC4oVaO7KTVPeix3P0c2PR3WlxUjG/voVA9/HYJaISfb8rb
# II01YBwCA8sgsKxYoA5AY8WYIsGyWfVVa88nq2x2zm8jLfR+cWojayL/ErhULSd+
# 2DrZ8LaHlv1b0VysGMNNn3O3AamfV6peKOK5lDCCBrQwggScoAMCAQICEA3HrFcF
# /yGZLkBDIgw6SYYwDQYJKoZIhvcNAQELBQAwYjELMAkGA1UEBhMCVVMxFTATBgNV
# BAoTDERpZ2lDZXJ0IEluYzEZMBcGA1UECxMQd3d3LmRpZ2ljZXJ0LmNvbTEhMB8G
# A1UEAxMYRGlnaUNlcnQgVHJ1c3RlZCBSb290IEc0MB4XDTI1MDUwNzAwMDAwMFoX
# DTM4MDExNDIzNTk1OVowaTELMAkGA1UEBhMCVVMxFzAVBgNVBAoTDkRpZ2lDZXJ0
# LCBJbmMuMUEwPwYDVQQDEzhEaWdpQ2VydCBUcnVzdGVkIEc0IFRpbWVTdGFtcGlu
# ZyBSU0E0MDk2IFNIQTI1NiAyMDI1IENBMTCCAiIwDQYJKoZIhvcNAQEBBQADggIP
# ADCCAgoCggIBALR4MdMKmEFyvjxGwBysddujRmh0tFEXnU2tjQ2UtZmWgyxU7UNq
# EY81FzJsQqr5G7A6c+Gh/qm8Xi4aPCOo2N8S9SLrC6Kbltqn7SWCWgzbNfiR+2fk
# HUiljNOqnIVD/gG3SYDEAd4dg2dDGpeZGKe+42DFUF0mR/vtLa4+gKPsYfwEu7EE
# bkC9+0F2w4QJLVSTEG8yAR2CQWIM1iI5PHg62IVwxKSpO0XaF9DPfNBKS7Zazch8
# NF5vp7eaZ2CVNxpqumzTCNSOxm+SAWSuIr21Qomb+zzQWKhxKTVVgtmUPAW35xUU
# FREmDrMxSNlr/NsJyUXzdtFUUt4aS4CEeIY8y9IaaGBpPNXKFifinT7zL2gdFpBP
# 9qh8SdLnEut/GcalNeJQ55IuwnKCgs+nrpuQNfVmUB5KlCX3ZA4x5HHKS+rqBvKW
# xdCyQEEGcbLe1b8Aw4wJkhU1JrPsFfxW1gaou30yZ46t4Y9F20HHfIY4/6vHespY
# MQmUiote8ladjS/nJ0+k6MvqzfpzPDOy5y6gqztiT96Fv/9bH7mQyogxG9QEPHrP
# V6/7umw052AkyiLA6tQbZl1KhBtTasySkuJDpsZGKdlsjg4u70EwgWbVRSX1Wd4+
# zoFpp4Ra+MlKM2baoD6x0VR4RjSpWM8o5a6D8bpfm4CLKczsG7ZrIGNTAgMBAAGj
# ggFdMIIBWTASBgNVHRMBAf8ECDAGAQH/AgEAMB0GA1UdDgQWBBTvb1NK6eQGfHrK
# 4pBW9i/USezLTjAfBgNVHSMEGDAWgBTs1+OC0nFdZEzfLmc/57qYrhwPTzAOBgNV
# HQ8BAf8EBAMCAYYwEwYDVR0lBAwwCgYIKwYBBQUHAwgwdwYIKwYBBQUHAQEEazBp
# MCQGCCsGAQUFBzABhhhodHRwOi8vb2NzcC5kaWdpY2VydC5jb20wQQYIKwYBBQUH
# MAKGNWh0dHA6Ly9jYWNlcnRzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRS
# b290RzQuY3J0MEMGA1UdHwQ8MDowOKA2oDSGMmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0
# LmNvbS9EaWdpQ2VydFRydXN0ZWRSb290RzQuY3JsMCAGA1UdIAQZMBcwCAYGZ4EM
# AQQCMAsGCWCGSAGG/WwHATANBgkqhkiG9w0BAQsFAAOCAgEAF877FoAc/gc9EXZx
# ML2+C8i1NKZ/zdCHxYgaMH9Pw5tcBnPw6O6FTGNpoV2V4wzSUGvI9NAzaoQk97fr
# PBtIj+ZLzdp+yXdhOP4hCFATuNT+ReOPK0mCefSG+tXqGpYZ3essBS3q8nL2UwM+
# NMvEuBd/2vmdYxDCvwzJv2sRUoKEfJ+nN57mQfQXwcAEGCvRR2qKtntujB71WPYA
# gwPyWLKu6RnaID/B0ba2H3LUiwDRAXx1Neq9ydOal95CHfmTnM4I+ZI2rVQfjXQA
# 1WSjjf4J2a7jLzWGNqNX+DF0SQzHU0pTi4dBwp9nEC8EAqoxW6q17r0z0noDjs6+
# BFo+z7bKSBwZXTRNivYuve3L2oiKNqetRHdqfMTCW/NmKLJ9M+MtucVGyOxiDf06
# VXxyKkOirv6o02OoXN4bFzK0vlNMsvhlqgF2puE6FndlENSmE+9JGYxOGLS/D284
# NHNboDGcmWXfwXRy4kbu4QFhOm0xJuF2EZAOk5eCkhSxZON3rGlHqhpB/8MluDez
# ooIs8CVnrpHMiD2wL40mm53+/j7tFaxYKIqL0Q4ssd8xHZnIn/7GELH3IdvG2XlM
# 9q7WP/UwgOkw/HQtyRN62JK4S1C8uw3PdBunvAZapsiI5YKdvlarEvf8EA+8hcpS
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIN9IGVYTPkIXvwD3iwMu1dxbqKwHfKOvtxni66nmYIIfMA0GCSqG
# SIb3DQEBAQUABIIBgAuoDpDE+0IoetMrFB2YbqUZpnSRd0+WHYMTn+UmArw5zSOK
# uEnIrwPI0XrtwKMjG0fmnAyxatF56m+YwqbU9NLIoLCyUpmB+MlwqmYFmw/bjaSr
# fwQH/NYClfXJGwSU7mAdvnevSVpi7W2Vpw+1D6n+ZHh+LDoTFGGfBCkvyGU5KcBQ
# 3rbxEu3kgF3f84nK39uJzwoZxFKqcnR30ILgr3vnIOI+kswWMDa3wSBunwQTIjjv
# tsaLwWtyLbUbj5fLtDEoWrF6cna2ajAePYON9XEdCHdCDcu55d2qvGAVMSsl/gFh
# RcEh7qk5V76nkxhOT2HMRDt0tEQeIpPKLHsVYNzKUJhRU6SOzzCdJxj7WZuoqLkq
# wf117nNOD5fQmii/LW19PITtgnQfKV23418WVuovSFW2AFpn3u0juLdd3N0lpx/l
# m/UVh2UoTXbS3UworstKihS5kG7EWJjFXFJ6+Sz88D1nowKXfkxGmxjjjT26doA9
# upvpElKJY3SOaZHWeaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# NDBaMC8GCSqGSIb3DQEJBDEiBCAYuHBH0b+wvS36UXHoQnrRBkXg7GTLsYIQS07S
# e6rNBDANBgkqhkiG9w0BAQEFAASCAgBpop+lfvJiTML3QBgoNkWI9fXbbL0/kwU9
# 2NNIJ0kj9pIdfG9ndqF/wz5wk/D8AiuEgMP8aEc4bQqKmffmdyMPx9p3P+wPFPpI
# 3/aGvQ52Gu2H1TbUDClhDPsHe66rqkQdrPDmPJQ48PmCk5eXDdKPo256Z/+5NjoB
# Dmdu56NhBujSSGXnpBcI/V7/yAx9PSTeuiEnJDWqvFg14hp4/7yDPrUZ6vlGEO6t
# aCuT3+LQy1N/IhBvHsTRitQrGweXDmnHZdeBvRzM08nVywB0dJbInl+gGQSFTelC
# 0/3F7cJJFSyokxerPFI3PYbJ1Ge47Z6xIOCUW7ObX38ZGgh4MeuGp/6LnilxzoUB
# BHWW+eLR59ngg5EHnV8vcb418VKc16jBfs88R+21jnkcQUyI5fRnpjCFveIV2aNf
# p5DVlQl3+7T9DSqb6BxoR0BnJFoDNj2P464/rGWCNJmV/CsGYEovcGBsweBR2Vmk
# 6uyzYdKvo0tRI8vVZKZzETSEbW+Pk41B/2kWZlwx+w5zE4eNiSkQolDjXkjF3/yv
# sZ5wt7duGsqPnsksab92OnFXJcPxWul+G36hAc6fHWcQ3QCRGxtPv6hTdiFW66Ir
# l+uGcMrYSWAPxYBF9HkxWQjelIpfhCItuusCEt/wXGHe0JQ0tRvezZ10HzG44722
# ddHBzlwoEA==
# SIG # End signature block
