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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDBFisbsnNc3Kyl
# b1pRcvawAyVOR9VWnV2eQnXJHGKEUaCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDfSBlWEz5CF78A94sD
# LtXcW6isB3yjr7cZ4uup5mCCHzANBgkqhkiG9w0BAQEFAASCAYAyLiJJfyPONYyp
# e2/0gT6ClQNW/eYxWCXV08iUGLmZUGBGmM4iellc/Vrg/I1oMMnr/CAP1mACK323
# s+oo9rKxSsDpUOSR1qfZGHatl5wo6NTX+zQ3yknivEXgWb47gZfhkQLIWi5vCC8C
# IVaqEv513tgEW7CwR20AzOdNgc0uXiFixhA2jetI6WV9fckKp83AiQB2XdltAIo2
# 2Z3P4j9aEKgIOL7DUJPnNuS5V5eg6xO2+Y+70n64C2NVOg1Be9a9NrIY0voXeEqs
# iYkBUjJcydUf/701nVtU3NqS8jjxqHgvJpcMLOX5RJKFJUryXEE1HtSGosVZR+Z0
# 5rTm+c8TUWrIMjr0uYISt5ujFWOQc7Yv1FuJeJZ7I8xyLlM65wR8sTr7XSI9p4uI
# bpEhtajQhAFUmPB7obBfcdTle5FsWlWaCzpv4lziKi7oJ+1TJ3qis/nbHtLjGHnI
# nBmMGJok5IkIKzAExpNtSORX9yI+dwO0Q23PyuPOmBV2ZKBkYZ0=
# SIG # End signature block
