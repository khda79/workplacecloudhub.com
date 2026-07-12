<#
.SYNOPSIS
    Runs the SmartM365 Windows 11 Upgrade Toolkit endpoint from an Intune scheduled task.
.DESCRIPTION
    Executes the local endpoint script against the packaged setup media cache and removes the scheduled task once the device is already Windows 11.
.VERSION
    1.0.8
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

trap {
    try { Write-RunnerLog ("Runner failed: {0}" -f $_.Exception.Message) 'ERROR' } catch { }
    exit 1
}

Write-RunnerLog ("Runner started. DataRoot={0}; TaskName={1}; RunGuardHours={2}; User={3}; Computer={4}; PID={5}" -f $DataRoot,$TaskName,$RunGuardHours,([Security.Principal.WindowsIdentity]::GetCurrent().Name),$env:COMPUTERNAME,$PID)

if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "Package manifest not found: $manifestPath" }
if (-not (Test-Path -LiteralPath $endpointScript -PathType Leaf)) { throw "Endpoint script not found: $endpointScript" }
$endpointParseErrors = $null
[void][System.Management.Automation.PSParser]::Tokenize((Get-Content -LiteralPath $endpointScript -Raw), [ref]$endpointParseErrors)
if ($endpointParseErrors) {
    $parseSummary = (($endpointParseErrors | Select-Object -First 5 | ForEach-Object { "Line={0}; Column={1}; Message={2}" -f $_.Token.StartLine,$_.Token.StartColumn,$_.Message }) -join ' | ')
    Write-RunnerLog ("Endpoint script parse validation failed: {0}" -f $parseSummary) 'ERROR'
    exit 1
}
Write-RunnerLog ("Endpoint script parse validation succeeded: {0}" -f $endpointScript)
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
if (-not (Test-Path -LiteralPath (Join-Path $cachePath 'setup.exe') -PathType Leaf)) { throw "Local setup cache is missing setup.exe: $cachePath" }
if (-not (Test-Path -LiteralPath (Join-Path $cachePath 'sources\install.wim') -PathType Leaf)) { throw "Local setup cache is missing sources\install.wim: $cachePath" }

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
$latestEndpointLog = Get-ChildItem -LiteralPath (Join-Path $DataRoot 'Logs') -Filter 'SmartM365-Invoke-Windows11UpgradeRepair_*.log' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
if ($latestEndpointLog) {
    Write-RunnerLog ("Endpoint script exited with code {0}. LatestEndpointLog={1}; LastWriteTime={2}; StdOut={3}; StdErr={4}" -f $exit,$latestEndpointLog.FullName,$latestEndpointLog.LastWriteTime,$endpointStdOut,$endpointStdErr)
}
else {
    Write-RunnerLog ("Endpoint script exited with code {0}. LatestEndpointLog=<not found>; StdOut={1}; StdErr={2}" -f $exit,$endpointStdOut,$endpointStdErr) 'WARN'
}

if ((Get-OsFamily) -eq 'Windows11') {
    Write-RunnerLog 'Device is Windows 11 after endpoint run. Removing scheduled task.'
    try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue } catch { }
}

exit 0

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBhylJZTL+2K22F
# N2mA4V2JG4QXuqRb8d58hNXLa3teQ6CCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCABOebagoTkeMUzcigT1oXO/gWaJ/RwqsXLRiLXQILHozANBgkqhkiG9w0B
# AQEFAASCAYBUhrJ208CcqC4ghiB5fzRfKZkpje7HfKSBKjWPi8qDDaHOXvl05aK2
# XaiRSFl5VLwFxUgGKwiQkKbcNupO0ud2X8yCM66V2z1aZkUe1Thz87CNebQ0r92w
# pMgRalw+isFDtJTVhctap23mznvM8ES945dvzBnlRc8Zp8J1G5SjOaiLVK/ShYS+
# S/WVbXx7wCQvmEp3od1Zs69eCHE6N+8+ngoMIMuD/Dgj5JAwvcOrxbw4w4HY9nrb
# UCrIBqxZ+y8QnhvZLthmqT0QEWodC53plwREK0wE+7vgRGnlpLvGmc5Y9j8jefq8
# OJSn8kaTCIRYwR57dGgfO7WuJHcO29l+c5dZSRCID9KKvQcTB2SPNpBatEwjnOfh
# XKSCcHXwp9c6/T8KN5RwM7DkhJaPEVyWZLiApDG6elGKcx6b4R1TPA40AZhcQz9I
# 9/8uo1icfLWh5SkZOo41HSZ5vIu9q7ZmXmjDkA8tiJPtD3WWGyHAS85L5jzdT8a8
# 6PN5m9/lDGc=
# SIG # End signature block
