<#
.SYNOPSIS
    Runs the SmartM365 Windows 11 Upgrade Toolkit endpoint from an Intune scheduled task.
.DESCRIPTION
    Executes the local endpoint script against the packaged setup media cache and removes the scheduled task once the device is already Windows 11.
.VERSION
    1.0.11
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>
[CmdletBinding()]
param(
    [string]$DataRoot = 'C:\ProgramData\SmartM365\Windows11UpgradeToolkit',
    [string]$TaskName = 'SmartM365 Windows 11 Upgrade Toolkit - Intune',
    [int]$RunGuardHours = 2,
    [ValidateRange(1, 365)][int]$LogRetentionDays = 7,
    [ValidateRange(10, 5000)][int]$MaxEndpointLogFiles = 200
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$intuneRoot = Join-Path $DataRoot 'Intune'
$logRoot = Join-Path $DataRoot 'Logs\Intune'
$manifestPath = Join-Path $intuneRoot 'PackageManifest.json'
$integrityHelperPath = Join-Path $intuneRoot 'SmartM365-SetupMediaIntegrity.ps1'
$endpointScript = Join-Path $DataRoot 'SmartM365-Invoke-Windows11UpgradeRepair.ps1'
$setupCacheRoot = Join-Path $DataRoot 'SetupMedia'
$registrySubKeyRoot = 'SOFTWARE\SmartM365\Windows11UpgradeToolkit\IntunePackages'

function New-Directory {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}

function Invoke-IntuneLogMaintenance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [ValidateRange(1, 365)][int]$RetentionDays = 7,
        [ValidateRange(1, 5000)][int]$MaximumEndpointFiles = 200
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }

    $cutoff = (Get-Date).AddDays(-1 * $RetentionDays)
    $removedByAge = 0
    $removedByCount = 0
    $failed = 0
    $files = @(Get-ChildItem -LiteralPath $Path -Filter 'Endpoint_*.log' -File -ErrorAction Stop)

    foreach ($file in @($files | Where-Object { $_.LastWriteTime -lt $cutoff })) {
        try {
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
            $removedByAge++
        }
        catch { $failed++ }
    }

    $remaining = @(Get-ChildItem -LiteralPath $Path -Filter 'Endpoint_*.log' -File -ErrorAction Stop |
        Sort-Object LastWriteTime,Name -Descending)
    foreach ($file in @($remaining | Select-Object -Skip $MaximumEndpointFiles)) {
        try {
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
            $removedByCount++
        }
        catch { $failed++ }
    }

    $remainingCount = @(Get-ChildItem -LiteralPath $Path -Filter 'Endpoint_*.log' -File -ErrorAction Stop).Count
    return [pscustomobject]@{
        Scanned        = $files.Count
        RemovedByAge   = $removedByAge
        RemovedByCount = $removedByCount
        Failed         = $failed
        Remaining      = $remainingCount
    }
}

function Write-RunnerLog {
    param([string]$Message, [string]$Level = 'INFO')
    New-Directory -Path $logRoot
    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Level,$Message
    Add-Content -LiteralPath (Join-Path $logRoot 'Run-IntuneUpgrade.log') -Value $line -Encoding UTF8
}

function Invoke-RunnerLogMaintenance {
    param([string]$Phase)

    try {
        $result = Invoke-IntuneLogMaintenance -Path $logRoot -RetentionDays $LogRetentionDays -MaximumEndpointFiles $MaxEndpointLogFiles
        $level = if ($result.Failed -gt 0) { 'WARN' } else { 'INFO' }
        Write-RunnerLog ("Intune log maintenance completed. Phase={0}; RetentionDays={1}; MaximumEndpointFiles={2}; Scanned={3}; RemovedByAge={4}; RemovedByCount={5}; Failed={6}; Remaining={7}" -f $Phase,$LogRetentionDays,$MaxEndpointLogFiles,$result.Scanned,$result.RemovedByAge,$result.RemovedByCount,$result.Failed,$result.Remaining) $level
    }
    catch {
        Write-RunnerLog ("Intune log maintenance failed without blocking the upgrade. Phase={0}; Error={1}" -f $Phase,$_.Exception.Message) 'WARN'
    }
}

function Open-Registry64LocalMachine {
    return [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
}

function Set-Registry64String {
    param([string]$SubKey, [string]$Name, [string]$Value)

    $baseKey = Open-Registry64LocalMachine
    try {
        $key = $baseKey.CreateSubKey($SubKey)
        try { $key.SetValue($Name, [string]$Value, [Microsoft.Win32.RegistryValueKind]::String) }
        finally { if ($key) { $key.Dispose() } }
    }
    finally { $baseKey.Dispose() }
}

function Set-PackageRepairRequired {
    param(
        [Parameter(Mandatory = $true)][object]$Manifest,
        [Parameter(Mandatory = $true)][string]$Reason
    )

    $packageId = [string]$Manifest.PackageId
    if ([string]::IsNullOrWhiteSpace($packageId)) { throw 'Cannot mark RepairRequired because PackageId is empty.' }
    $subKey = "$registrySubKeyRoot\$packageId"
    $safeReason = if ($Reason.Length -gt 2048) { $Reason.Substring(0, 2048) } else { $Reason }
    Set-Registry64String -SubKey $subKey -Name PackageId -Value $packageId
    Set-Registry64String -SubKey $subKey -Name PackageVersion -Value ([string]$Manifest.PackageVersion)
    Set-Registry64String -SubKey $subKey -Name InstallState -Value 'RepairRequired'
    Set-Registry64String -SubKey $subKey -Name RepairReason -Value $safeReason
    Set-Registry64String -SubKey $subKey -Name RepairRequiredUtc -Value ((Get-Date).ToUniversalTime().ToString('o'))
    Write-RunnerLog ("Marked Intune package RepairRequired so detection triggers reinstall. PackageId={0}; Reason={1}" -f $packageId,$safeReason) 'WARN'
    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
        Write-RunnerLog ("Removed scheduled task while package repair is required. TaskName={0}" -f $TaskName) 'WARN'
    }
    catch {
        Write-RunnerLog ("Could not remove scheduled task after marking RepairRequired. TaskName={0}; Error={1}" -f $TaskName,$_.Exception.Message) 'WARN'
    }
}

function Set-PackageInstalledForWindows11 {
    param([Parameter(Mandatory = $true)][object]$Manifest)

    $packageId = [string]$Manifest.PackageId
    if ([string]::IsNullOrWhiteSpace($packageId)) { throw 'Cannot mark Installed because PackageId is empty.' }
    $subKey = "$registrySubKeyRoot\$packageId"
    Set-Registry64String -SubKey $subKey -Name PackageId -Value $packageId
    Set-Registry64String -SubKey $subKey -Name PackageVersion -Value ([string]$Manifest.PackageVersion)
    Set-Registry64String -SubKey $subKey -Name InstallState -Value 'Installed'
    Set-Registry64String -SubKey $subKey -Name CompletionReason -Value 'AlreadyWindows11'
    Set-Registry64String -SubKey $subKey -Name RepairReason -Value ''
    Set-Registry64String -SubKey $subKey -Name RepairRequiredUtc -Value ''
    Set-Registry64String -SubKey $subKey -Name InstalledUtc -Value ((Get-Date).ToUniversalTime().ToString('o'))
    Write-RunnerLog ("Device is Windows 11. Marked Intune package Installed. PackageId={0}; PackageVersion={1}" -f $packageId,$Manifest.PackageVersion)
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

trap {
    try { Write-RunnerLog ("Runner failed: {0}" -f $_.Exception.Message) 'ERROR' } catch { }
    exit 1
}

Invoke-RunnerLogMaintenance -Phase 'Startup'
Write-RunnerLog ("Runner started. DataRoot={0}; TaskName={1}; RunGuardHours={2}; LogRetentionDays={3}; MaxEndpointLogFiles={4}; User={5}; Computer={6}; PID={7}" -f $DataRoot,$TaskName,$RunGuardHours,$LogRetentionDays,$MaxEndpointLogFiles,([Security.Principal.WindowsIdentity]::GetCurrent().Name),$env:COMPUTERNAME,$PID)

$osFamily = Get-OsFamily
if ($osFamily -eq 'Windows11') {
    if (Test-Path -LiteralPath $manifestPath -PathType Leaf) {
        try {
            $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
            Set-PackageInstalledForWindows11 -Manifest $manifest
        }
        catch { Write-RunnerLog ("Windows 11 was detected, but package state could not be refreshed. Intune OS detection remains authoritative. Error={0}" -f $_.Exception.Message) 'WARN' }
    }
    else {
        Write-RunnerLog ("Windows 11 was detected before package manifest validation. Intune OS detection remains authoritative. MissingManifest={0}" -f $manifestPath) 'WARN'
    }
    Write-RunnerLog 'Device is already Windows 11. Removing scheduled task.'
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch { }
    exit 0
}

if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "Package manifest not found: $manifestPath" }
$manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
if (-not (Test-Path -LiteralPath $endpointScript -PathType Leaf)) { throw "Endpoint script not found: $endpointScript" }
$endpointParseErrors = $null
[void][System.Management.Automation.PSParser]::Tokenize((Get-Content -LiteralPath $endpointScript -Raw), [ref]$endpointParseErrors)
if ($endpointParseErrors) {
    $parseSummary = (($endpointParseErrors | Select-Object -First 5 | ForEach-Object { "Line={0}; Column={1}; Message={2}" -f $_.Token.StartLine,$_.Token.StartColumn,$_.Message }) -join ' | ')
    Write-RunnerLog ("Endpoint script parse validation failed: {0}" -f $parseSummary) 'ERROR'
    exit 1
}
Write-RunnerLog ("Endpoint script parse validation succeeded: {0}" -f $endpointScript)
$setupProcesses = @(Get-Process -Name setup,setuphost,setupprep -ErrorAction SilentlyContinue)
if ($setupProcesses.Count -gt 0) {
    Write-RunnerLog ("Setup process already running; skipping this cycle. Processes={0}" -f (($setupProcesses | Select-Object -ExpandProperty Id) -join ','))
    exit 0
}

$cacheFolder = [string]$manifest.SetupCacheFolder
$cachePath = Join-Path $setupCacheRoot $cacheFolder
try {
    if (-not (Test-Path -LiteralPath $integrityHelperPath -PathType Leaf)) { throw "Setup media integrity helper not found: $integrityHelperPath" }
    . $integrityHelperPath
    if (-not (Test-Path -LiteralPath (Join-Path $cachePath 'setup.exe') -PathType Leaf)) { throw "Local setup cache is missing setup.exe: $cachePath" }
    if (-not (Test-Path -LiteralPath (Join-Path $cachePath 'sources\install.wim') -PathType Leaf)) { throw "Local setup cache is missing sources\install.wim: $cachePath" }
    $cacheIntegrity = Test-SmartM365SetupMediaIntegrity -MediaRoot $cachePath
    Write-RunnerLog ("Setup cache integrity validated before endpoint launch. Files={0}; Bytes={1}; Root={2}" -f $cacheIntegrity.Files,$cacheIntegrity.Bytes,$cacheIntegrity.MediaRoot)
}
catch {
    $repairReason = $_.Exception.Message
    Set-PackageRepairRequired -Manifest $manifest -Reason $repairReason
    throw $repairReason
}

$args = @(
    '-RunGuardHours', [string]$RunGuardHours,
    '-DirectSetupUpgrade',
    '-SkipVirtualMachines',
    '-AllowReboot',
    '-AllowSetupCompletionRebootWhenNoUser',
    '-AllowSetupProfileRepair',
    '-AllowDiskCleanup',
    '-SetupExecutionMode', 'LocalCache',
    '-SetupMediaId', [string]$manifest.MediaId,
    '-SetupLanguage', [string]$manifest.Language,
    '-SetupDynamicUpdate', 'Disable',
    '-SetupCacheRoot', $setupCacheRoot,
    '-SetupProcessHeartbeatSeconds', '300',
    '-SetupProcessTimeoutMinutes', '0',
    '-SetupMediaCopyTimeoutMinutes', '180',
    '-ForceRequiredRebootWhenUptimeOverDays', '7',
    '-DataRoot', $DataRoot
)

Write-RunnerLog ("Starting endpoint script for package {0}; Language={1}; Cache={2}" -f $manifest.PackageId,$manifest.Language,$cachePath)
$powerShellExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$endpointProcessArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $endpointScript) + $args
$endpointStdOut = Join-Path $logRoot ("Endpoint_{0}_stdout.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
$endpointStdErr = Join-Path $logRoot ("Endpoint_{0}_stderr.log" -f (Get-Date -Format 'yyyyMMdd-HHmmss'))
$process = Start-Process -FilePath $powerShellExe -ArgumentList $endpointProcessArgs -Wait -PassThru -NoNewWindow -RedirectStandardOutput $endpointStdOut -RedirectStandardError $endpointStdErr
$exit = [int]$process.ExitCode
Invoke-RunnerLogMaintenance -Phase 'PostRun'
if ($exit -ne 0) {
    try { [void](Test-SmartM365SetupMediaIntegrity -MediaRoot $cachePath) }
    catch {
        $postRunCacheError = $_.Exception.Message
        try { Set-PackageRepairRequired -Manifest $manifest -Reason $postRunCacheError }
        catch { Write-RunnerLog ("Failed to mark package RepairRequired after endpoint cache failure. Error={0}" -f $_.Exception.Message) 'ERROR' }
    }
}
$latestEndpointLog = Get-ChildItem -LiteralPath (Join-Path $DataRoot 'Logs') -Filter 'SmartM365-Invoke-Windows11UpgradeRepair_*.log' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($latestEndpointLog) {
    Write-RunnerLog ("Endpoint script exited with code {0}. LatestEndpointLog={1}; LastWriteTime={2}; StdOut={3}; StdErr={4}" -f $exit,$latestEndpointLog.FullName,$latestEndpointLog.LastWriteTime,$endpointStdOut,$endpointStdErr)
}
else {
    Write-RunnerLog ("Endpoint script exited with code {0}. LatestEndpointLog=<not found>; StdOut={1}; StdErr={2}" -f $exit,$endpointStdOut,$endpointStdErr) 'WARN'
}

if ((Get-OsFamily) -eq 'Windows11') {
    Set-PackageInstalledForWindows11 -Manifest $manifest
    Write-RunnerLog 'Device is Windows 11 after endpoint run. Removing scheduled task.'
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch { }
}

exit 0

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBhylJZTL+2K22F
# N2mA4V2JG4QXuqRb8d58hNXLa3teQ6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAE55tqChOR4xTNyKBPWhc7+BZon9HCqxctGItdAgsejMA0GCSqG
# SIb3DQEBAQUABIIBgCTF/LGWc4AU57RMmT3A42+unpPHPTTOEQZKy9FyEMDUuf2u
# UzPb8hzpLwC9OLxfv3I5vwb4GwXP2P9NGfrRdurVk60nSdPuzYUE9Kb1Y9kAsPUY
# 5OpJMj06ZDIumoDMN23Jf7boMa3eGV9COjwagNEaY6PoNmEFZegYnxqJdWET4ZmN
# Xa+tS/q0+cPKBxNPMErCe2K+K/cpGpicHl6yUKhhkqy1pK4EpydfeN8ficsf0hcv
# DtakbNd9H4kz/k3h+IJvMQOpkJfyIxMSwg8TkKe+DnRV5VH827kRHjqkwigVyTWt
# ZJ75cSBGK7KRUiyAEmCQyJ6NC2WfknVyb1+07tZoeiQr4mGZzZhhHBAWYQ3sio6S
# UPpeat1tlbJonOkRVDZhURqXoCxHT70rk4LsOc2/ugxTyJvNe7tthnSzpVJlgReZ
# RiWbLpiEfCTZz+GPeG/YmWdOfb5jYkgvMO3SJFRRavYUa1JuEo1nofAMbYEIhxym
# fU3IofI6px49Uucfp6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODQ5
# NDBaMC8GCSqGSIb3DQEJBDEiBCBqrQrt8BCzBB60ZnImXfC63ZikA0E2gkKsU4mR
# I84sODANBgkqhkiG9w0BAQEFAASCAgBn/F/GYf0CphKEpq74PmKbgQNsj9ujLAUP
# kTqt0HzsIqVnmUXAQpEs84nc7dt17L/+XUO8pjJG+AS8RSf+vcin0gbZmJNA44iU
# 1VL9/1K0VfXwPTpGR/inpoGlRQ//2zHu6121cExyTBcAU9bahH2it4tR2FxUQJ0R
# nL4FQT1V/v07eZO894FdWvn/M3Nq3weEeATLcz3X9zDWggj3HQa+zdwyAM6uNLwM
# oUptxLvJ5GBa2sQicNewsRx6ZSewL8wjo+q5lO+XPGlJaKPOqc2A782yAx8tKFXX
# 9j8uEZ9Y0z921aRcEW7PvaUwlG0BrgmuhwuJbyHCjoPRvS1AFrZS90lwokxT6VmN
# F5ViLrUU4lbqKP5D3FNCpbtT20U4cEgjgKmWd+Ta6eKybi+LFMEjbuc/k7ARq8II
# iBwRJTfmhi007hIR3OXKqDaa3NLT5yzaZxrfx7lYLY2nB3wt50B+ddJL2s480/oe
# +FDPmcRuSnni0SEEU7RANuhE3+VapkU4GpJBkBElzfn7v4vZlnufJqvGJ8Me4RlZ
# 47jgM0lH2dlmlLfBuz2rOCD0j+tpgFeP3VAOAxuNh2d/x8V+aztb7H6KT4RB2i8m
# Iu811IcnMphb7JNq3r6L7W6+KdCGuqft9rBz/DtVJSS3v/diPw+XQXeISSonwl+D
# 64GlgX6jZw==
# SIG # End signature block
