<#
.SYNOPSIS
    Runs SmartM365-Invoke-Windows11UpgradeRepair.ps1 on remote computers using PsExec.

.DESCRIPTION
    LOT/PsExec orchestrator for Windows 10 to Windows 11 upgrade diagnostics and guarded repair.
    It copies the autonomous endpoint script to each target, lets the target validate/cache
    Windows 11 setup media when setup upgrade is enabled, starts the script as SYSTEM,
    collects evidence, and writes cycle CSV reports.

.VERSION
    0.1.6

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding(PositionalBinding = $false)]
param(
    [string]$ComputerListPath = (Join-Path (Get-Location) 'Computers.txt'),
    [string]$PsExecPath,
    [string]$LocalScriptPath,

    [switch]$AuditOnly,
    [switch]$DryRun,
    [switch]$RunOnce,
    [switch]$IgnoreRunGuard,
    [switch]$AllowPolicyRepair,
    [switch]$AllowWUReset,
    [switch]$AllowForceUpgrade,
    [switch]$AllowSetupUpgrade,
    [switch]$DirectSetupUpgrade,
    [switch]$AllowReboot,
    [switch]$AllowSetupCompletionRebootWhenNoUser,
    [switch]$SkipVirtualMachines,
    [switch]$AllowDiskCleanup,
    [switch]$AllowAdvancedDiskCleanup,
    [switch]$AllowDismComponentCleanup,

    [string]$SetupSourcePath,
    [string]$SetupSourceMapPath,
    [ValidateSet('LocalCache','Share','Auto')]
    [string]$SetupExecutionMode = 'LocalCache',
    [string]$SetupMediaId = 'Win11',
    [string]$SetupLanguage = 'MatchSystem',
    [ValidateSet('Enable','Disable','NoDrivers','NoLCU','NoDriversNoLCU')]
    [string]$SetupDynamicUpdate = 'Disable',
    [switch]$SkipSetupMediaPreCopy,
    [ValidateRange(0, 100)][int]$SetupSourceCandidateLimit = 5,
    [ValidateRange(0, 10000)][int]$SetupMediaCopyIpGapMilliseconds = 0,
    [ValidateRange(0, 86400)][int]$SetupMediaCopyJitterSeconds = 0,
    [ValidateRange(0, 500)][int]$SetupSourceConcurrencyLimit = 0,
    [ValidateRange(1, 1440)][int]$SetupSourceConcurrencyLeaseMinutes = 240,
    [string]$SetupSourceConcurrencyGateRoot,

    [string]$LogRoot,
    [string]$ReportRoot,
    [string]$CentralLogRoot,
    [switch]$NoCentralLogCollection,
    [switch]$KeepCentralLogHistory,

    [ValidateRange(1, 200)][int]$ThrottleLimit = 10,
    [ValidateRange(0, 200)][int]$GlobalConcurrencyLimit = 15,
    [string]$GlobalConcurrencySemaphoreName = 'Local\SmartM365_Windows11UpgradeToolkit_ComputerWorkers',
    [ValidateRange(0, 1440)][int]$GlobalConcurrencyLeaseTimeoutMinutes = 0,
    [ValidateRange(0, 3600)][int]$DelayBetweenComputersSeconds = 0,
    [ValidateRange(1, 60)][int]$JobPollSeconds = 2,
    [ValidateRange(0, 1440)][int]$DelayBetweenCyclesMinutes = 5,
    [ValidateRange(0, 1000)][int]$MaxCycles = 0,
    [ValidateRange(0, 1440)][int]$PsExecTimeoutMinutes = 180,

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$UnexpectedArguments
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

if ($UnexpectedArguments -and $UnexpectedArguments.Count -gt 0) {
    throw ("Unexpected launcher argument(s): {0}. Pass PsExec with -PsExecPath <path>, not as a free argument." -f ($UnexpectedArguments -join ' '))
}

$script:LauncherVersion = '0.1.6'
$script:BaseDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$script:ToolkitRoot = Split-Path -Parent $script:BaseDir
if ([string]::IsNullOrWhiteSpace($LocalScriptPath)) {
    $LocalScriptPath = Join-Path $script:BaseDir 'SmartM365-Invoke-Windows11UpgradeRepair.ps1'
}
$LocalWorkerPath = Join-Path $script:BaseDir 'SmartM365-Windows11Upgrade-PsExecWorker.ps1'
if ([string]::IsNullOrWhiteSpace($LogRoot)) { $LogRoot = Join-Path (Split-Path -Parent $ComputerListPath) 'PsExecLogs' }
if ([string]::IsNullOrWhiteSpace($ReportRoot)) { $ReportRoot = Join-Path (Split-Path -Parent $ComputerListPath) 'Reports' }
if ([string]::IsNullOrWhiteSpace($CentralLogRoot)) { $CentralLogRoot = Join-Path (Split-Path -Parent $ComputerListPath) 'CentralLogs' }

$script:RemoteBaseDir = 'C:\ProgramData\SmartM365\Windows11UpgradeToolkit'
$script:RemoteScriptPath = Join-Path $script:RemoteBaseDir 'SmartM365-Invoke-Windows11UpgradeRepair.ps1'
$script:RemoteSetupCacheRoot = Join-Path $script:RemoteBaseDir 'SetupMedia'
if ($GlobalConcurrencyLeaseTimeoutMinutes -lt 1) {
    $timeoutBase = if ($PsExecTimeoutMinutes -gt 0) { $PsExecTimeoutMinutes } else { 240 }
    $GlobalConcurrencyLeaseTimeoutMinutes = [Math]::Max(30, $timeoutBase + 30)
}

function New-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
    }
}

function Get-ComputerList {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Computer list not found: $Path"
    }

    $seen = @{}
    $result = New-Object System.Collections.ArrayList
    foreach ($line in @(Get-Content -LiteralPath $Path -ErrorAction Stop)) {
        $name = ([string]$line).Trim().Trim([char]34)
        if ([string]::IsNullOrWhiteSpace($name) -or $name.StartsWith('#')) { continue }
        $key = $name.ToUpperInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        [void]$result.Add($name)
    }
    return @($result.ToArray())
}

function Get-ComputerListKey {
    param([Parameter(Mandatory = $true)][string]$ComputerName)

    return ($ComputerName.Trim().Trim([char]34).Split('.')[0]).ToUpperInvariant()
}

function Test-AlreadyWindows11CycleResult {
    param([Parameter(Mandatory = $true)][psobject]$Result)

    $launcherStatus = if ($Result.PSObject.Properties['LauncherStatus']) { [string]$Result.LauncherStatus } else { '' }
    $remoteStatus = if ($Result.PSObject.Properties['RemoteStatus']) { [string]$Result.RemoteStatus } else { '' }
    return ($launcherStatus -eq 'ALREADY_WINDOWS11' -or $remoteStatus -eq 'ALREADY_WINDOWS11')
}

function Move-AlreadyWindows11ComputersFromList {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerListPath,
        [Parameter(Mandatory = $true)][object[]]$CycleSummary
    )

    $alreadyWindows11 = @(
        $CycleSummary |
            Where-Object { $_ -and (Test-AlreadyWindows11CycleResult -Result $_) } |
            ForEach-Object { if ($_.PSObject.Properties['ComputerName']) { [string]$_.ComputerName } } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
    )

    if ($alreadyWindows11.Count -eq 0) {
        return [pscustomobject]@{ Moved = 0; AlreadyWindows11Path = ''; Detail = 'No already-Windows11 computer detected in this cycle.' }
    }

    $moveKeys = @{}
    foreach ($computer in $alreadyWindows11) {
        $key = Get-ComputerListKey -ComputerName $computer
        if (-not [string]::IsNullOrWhiteSpace($key) -and -not $moveKeys.ContainsKey($key)) { $moveKeys[$key] = $computer.Trim() }
    }

    $listLines = @(Get-Content -LiteralPath $ComputerListPath -ErrorAction Stop)
    $remainingLines = New-Object System.Collections.Generic.List[string]
    $movedFromList = New-Object System.Collections.Generic.List[string]

    foreach ($line in $listLines) {
        $trimmed = ([string]$line).Trim().Trim([char]34)
        if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed.StartsWith('#')) {
            [void]$remainingLines.Add($line)
            continue
        }

        $key = Get-ComputerListKey -ComputerName $trimmed
        if ($moveKeys.ContainsKey($key)) {
            [void]$movedFromList.Add($trimmed)
            continue
        }

        [void]$remainingLines.Add($line)
    }

    if ($movedFromList.Count -eq 0) {
        return [pscustomobject]@{ Moved = 0; AlreadyWindows11Path = ''; Detail = 'Already-Windows11 computers were detected, but none were still present in Computers.txt.' }
    }

    $computerListDir = Split-Path -Parent $ComputerListPath
    if ([string]::IsNullOrWhiteSpace($computerListDir)) { $computerListDir = '.' }
    $alreadyWindows11Path = Join-Path $computerListDir 'ComputersAlreadyW11.txt'

    $existingKeys = @{}
    if (Test-Path -LiteralPath $alreadyWindows11Path) {
        foreach ($line in @(Get-Content -LiteralPath $alreadyWindows11Path -ErrorAction SilentlyContinue)) {
            $trimmed = ([string]$line).Trim().Trim([char]34)
            if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed.StartsWith('#')) { continue }
            $key = Get-ComputerListKey -ComputerName $trimmed
            if (-not $existingKeys.ContainsKey($key)) { $existingKeys[$key] = $true }
        }
    }

    $appendLines = New-Object System.Collections.Generic.List[string]
    foreach ($computer in $movedFromList) {
        $key = Get-ComputerListKey -ComputerName $computer
        if (-not $existingKeys.ContainsKey($key)) {
            [void]$appendLines.Add($computer)
            $existingKeys[$key] = $true
        }
    }

    $tmpComputerListPath = '{0}.tmp.{1}.txt' -f $ComputerListPath,([guid]::NewGuid().ToString('N'))
    try {
        Set-Content -LiteralPath $tmpComputerListPath -Value $remainingLines -Encoding ASCII -Force
        Move-Item -LiteralPath $tmpComputerListPath -Destination $ComputerListPath -Force
    }
    finally {
        Remove-Item -LiteralPath $tmpComputerListPath -Force -ErrorAction SilentlyContinue
    }

    if ($appendLines.Count -gt 0) {
        Add-Content -LiteralPath $alreadyWindows11Path -Value $appendLines -Encoding ASCII
    }
    elseif (-not (Test-Path -LiteralPath $alreadyWindows11Path)) {
        New-Item -ItemType File -Path $alreadyWindows11Path -Force | Out-Null
    }

    return [pscustomobject]@{
        Moved = $movedFromList.Count
        AlreadyWindows11Path = $alreadyWindows11Path
        Detail = ('Moved {0} computer(s) from Computers.txt to ComputersAlreadyW11.txt.' -f $movedFromList.Count)
    }
}

function Test-SingleComputerLaunch {
    param([Parameter(Mandatory = $true)][string]$Path)

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    if ($fullPath -match '\\SingleComputer(Runs)?\\') { return $true }

    $computers = @(Get-ComputerList -Path $fullPath)
    return ($computers.Count -eq 1)
}

$script:IsSingleComputerLaunch = Test-SingleComputerLaunch -Path $ComputerListPath
if ($script:IsSingleComputerLaunch) {
    $ThrottleLimit = 1
    $GlobalConcurrencyLimit = 0
}
function Test-SetupSourceMapSyntax {
    param([AllowNull()][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return }

    $expanded = [Environment]::ExpandEnvironmentVariables($Path.Trim('"'))
    if (-not (Test-Path -LiteralPath $expanded -PathType Leaf)) {
        Write-Host ("Setup source map preflight skipped because the operator cannot read it: {0}" -f $expanded) -ForegroundColor DarkYellow
        return
    }

    $rows = @(Import-Csv -LiteralPath $expanded -ErrorAction Stop)
    if ($rows.Count -eq 0) {
        throw "Setup source map is empty: $expanded"
    }

    $validRows = 0
    $rowNumber = 1
    foreach ($row in $rows) {
        $rowNumber++
        $scopeType = if ($row.PSObject.Properties['ScopeType']) { [string]$row.ScopeType } elseif ($row.PSObject.Properties['Type']) { [string]$row.Type } else { '' }
        $scopeValue = if ($row.PSObject.Properties['ScopeValue']) { [string]$row.ScopeValue } elseif ($row.PSObject.Properties['Value']) { [string]$row.Value } else { '' }
        $sourcePath = if ($row.PSObject.Properties['SetupSourcePath']) { [string]$row.SetupSourcePath } elseif ($row.PSObject.Properties['SourcePath']) { [string]$row.SourcePath } else { '' }

        if ([string]::IsNullOrWhiteSpace($scopeType) -or [string]::IsNullOrWhiteSpace($sourcePath)) {
            throw "Invalid setup source map row $rowNumber. ScopeType and SetupSourcePath are required."
        }

        if ($scopeType -notmatch '^(?i:Default|Fallback|Subnet|CIDR|IPPrefix|Prefix|ComputerName|Hostname|ComputerPrefix|HostnamePrefix)$') {
            throw "Invalid setup source map row $rowNumber. Unsupported ScopeType '$scopeType'."
        }

        if ($scopeType -match '^(?i:Subnet|CIDR)$' -and $scopeValue -notmatch '^\d{1,3}(\.\d{1,3}){3}/([0-9]|[12][0-9]|3[0-2])$') {
            throw "Invalid setup source map row $rowNumber. ScopeValue must be CIDR for ScopeType '$scopeType'."
        }

        if ($sourcePath -notmatch '^\\\\[^\\]+\\[^\\]+') {
            throw "Invalid setup source map row $rowNumber. SetupSourcePath must be UNC: $sourcePath"
        }

        $validRows++
    }

    Write-Host ("Setup source map preflight OK: {0}; Rows={1}" -f $expanded,$validRows) -ForegroundColor Green
}

function Resolve-PsExecPath {
    param([string]$Path)

    $candidate = [string]$Path
    if (-not [string]::IsNullOrWhiteSpace($candidate)) {
        $embedded = @([regex]::Matches($candidate, '(?i)(?:[A-Z]:\\|\\\\)[^"<>|?*]+?PsExec\.exe') | ForEach-Object { $_.Value.Trim() })
        if ($embedded.Count -gt 1) {
            throw ("Invalid PsExecPath value: {0}. Multiple PsExec paths were provided: {1}." -f $Path,($embedded -join ' | '))
        }
        if ($embedded.Count -eq 1) { $candidate = $embedded[0] }
        $candidate = $candidate.Trim('"')
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return (Get-Item -LiteralPath $candidate).FullName
        }
        throw ("Invalid PsExecPath value: {0}. Provide exactly one executable path." -f $Path)
    }

    $local = Join-Path $script:BaseDir 'PsExec.exe'
    if (Test-Path -LiteralPath $local -PathType Leaf) { return (Get-Item -LiteralPath $local).FullName }
    $command = Get-Command -Name 'PsExec.exe' -CommandType Application -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    throw ("PsExec.exe not found. Place it in {0} or add PsExec.exe to PATH." -f $script:BaseDir)
}

if (-not (Test-Path -LiteralPath $LocalScriptPath -PathType Leaf)) {
    throw "Local repair script not found: $LocalScriptPath"
}
if (-not (Test-Path -LiteralPath $LocalWorkerPath -PathType Leaf)) {
    throw "Local PsExec worker script not found: $LocalWorkerPath"
}

New-Directory -Path $LogRoot
New-Directory -Path $ReportRoot
New-Directory -Path $CentralLogRoot

Test-SetupSourceMapSyntax -Path $SetupSourceMapPath

$resolvedPsExec = if ($DryRun) { '' } else { Resolve-PsExecPath -Path $PsExecPath }

$remoteArgs = New-Object System.Collections.ArrayList
if ($AuditOnly) { [void]$remoteArgs.Add('-AuditOnly') }
if ($IgnoreRunGuard) { [void]$remoteArgs.Add('-IgnoreRunGuard') }
if ($AllowPolicyRepair) { [void]$remoteArgs.Add('-AllowPolicyRepair') }
if ($AllowWUReset) { [void]$remoteArgs.Add('-AllowWUReset') }
if ($AllowForceUpgrade) { [void]$remoteArgs.Add('-AllowForceUpgrade') }
if ($AllowSetupUpgrade) { [void]$remoteArgs.Add('-AllowSetupUpgrade') }
if ($DirectSetupUpgrade) { [void]$remoteArgs.Add('-DirectSetupUpgrade') }
if ($AllowReboot) { [void]$remoteArgs.Add('-AllowReboot') }
if ($AllowSetupCompletionRebootWhenNoUser) { [void]$remoteArgs.Add('-AllowSetupCompletionRebootWhenNoUser') }
if ($SkipVirtualMachines) { [void]$remoteArgs.Add('-SkipVirtualMachines') }
if ($AllowDiskCleanup) { [void]$remoteArgs.Add('-AllowDiskCleanup') }
if ($AllowAdvancedDiskCleanup -or $AllowDismComponentCleanup) { [void]$remoteArgs.Add('-AllowAdvancedDiskCleanup') }
if ($SkipSetupMediaPreCopy) { [void]$remoteArgs.Add('-SkipSetupMediaPreCopy') }
[void]$remoteArgs.Add('-SetupExecutionMode'); [void]$remoteArgs.Add($SetupExecutionMode)
[void]$remoteArgs.Add('-SetupMediaId'); [void]$remoteArgs.Add($SetupMediaId)
[void]$remoteArgs.Add('-SetupLanguage'); [void]$remoteArgs.Add($SetupLanguage)
[void]$remoteArgs.Add('-SetupDynamicUpdate'); [void]$remoteArgs.Add($SetupDynamicUpdate)
[void]$remoteArgs.Add('-SetupCacheRoot'); [void]$remoteArgs.Add($script:RemoteSetupCacheRoot)
[void]$remoteArgs.Add('-SetupSourceCandidateLimit'); [void]$remoteArgs.Add([string]$SetupSourceCandidateLimit)
[void]$remoteArgs.Add('-SetupMediaCopyIpGapMilliseconds'); [void]$remoteArgs.Add([string]$SetupMediaCopyIpGapMilliseconds)
[void]$remoteArgs.Add('-SetupMediaCopyJitterSeconds'); [void]$remoteArgs.Add([string]$SetupMediaCopyJitterSeconds)
if ($SetupSourceConcurrencyLimit -gt 0) {
    [void]$remoteArgs.Add('-SetupSourceConcurrencyLimit'); [void]$remoteArgs.Add([string]$SetupSourceConcurrencyLimit)
    [void]$remoteArgs.Add('-SetupSourceConcurrencyLeaseMinutes'); [void]$remoteArgs.Add([string]$SetupSourceConcurrencyLeaseMinutes)
}
if (-not [string]::IsNullOrWhiteSpace($SetupSourceConcurrencyGateRoot)) {
    [void]$remoteArgs.Add('-SetupSourceConcurrencyGateRoot'); [void]$remoteArgs.Add($SetupSourceConcurrencyGateRoot)
}
if (-not [string]::IsNullOrWhiteSpace($SetupSourcePath)) {
    [void]$remoteArgs.Add('-SetupSourcePath'); [void]$remoteArgs.Add($SetupSourcePath)
}
if (-not [string]::IsNullOrWhiteSpace($SetupSourceMapPath)) {
    [void]$remoteArgs.Add('-SetupSourceMapPath'); [void]$remoteArgs.Add($SetupSourceMapPath)
}

Write-Host "SmartM365 Windows 11 Upgrade Toolkit launcher v$script:LauncherVersion"
Write-Host "Computer list : $ComputerListPath"
Write-Host "PsExec        : $resolvedPsExec"
Write-Host "Repair script : $LocalScriptPath"
Write-Host "Worker script : $LocalWorkerPath"
Write-Host "Mode          : DryRun=$DryRun; AuditOnly=$AuditOnly; RunOnce=$RunOnce; SkipVirtualMachines=$SkipVirtualMachines; DiskCleanup=$AllowDiskCleanup; AdvancedCleanup=$($AllowAdvancedDiskCleanup -or $AllowDismComponentCleanup); DirectSetup=$DirectSetupUpgrade; SetupCompletionRebootWhenNoUser=$AllowSetupCompletionRebootWhenNoUser"
Write-Host "Setup         : Allow=$AllowSetupUpgrade; Mode=$SetupExecutionMode; MediaId=$SetupMediaId; Language=$SetupLanguage; DynamicUpdate=$SetupDynamicUpdate; PreCopy=$(-not $SkipSetupMediaPreCopy)"
Write-Host "Parallelism   : ThrottleLimit=$ThrottleLimit; GlobalConcurrencyLimit=$GlobalConcurrencyLimit; GlobalLeaseTimeout=$GlobalConcurrencyLeaseTimeoutMinutes minute(s)"
if ($script:IsSingleComputerLaunch) { Write-Host "Single PC     : worker limits ignored for one-computer launch." }
Write-Host "Reports       : $ReportRoot"
Write-Host ""

$reportColumns = @(
    'Timestamp',
    'ComputerName',
    'CycleNumber',
    'LauncherStatus',
    'RemoteStatus',
    'RemoteNextAction',
    'ExitCode',
    'Detail',
    'JobErrorMessage',
    'SetupCacheAction',
    'SetupDynamicUpdate',
    'SelectedSetupSourcePath',
    'SetupSourceSelectionDetail',
    'DiskCleanupAction',
    'DiskCleanupFreedGB',
    'AdvancedDiskCleanupAction',
    'AdvancedDiskCleanupFreedGB',
    'DismCleanupAction',
    'DismCleanupFreedGB',
    'SetupCompletionRebootAction',
    'SetupCompletionRebootDetail',
    'SetupCompletionRebootUserCount',
    'SetupCompletionRebootUsers',
    'ControlledRebootAction',
    'ControlledRebootDetail',
    'ControlledRebootUserCount',
    'ControlledRebootUsers',
    'RemoteLogsPath',
    'PsExecLogPath'
)

$globalGateRoot = Join-Path ([System.IO.Path]::GetTempPath()) 'SmartM365\GlobalWorkerGates'
$globalGateName = ($GlobalConcurrencySemaphoreName -replace '[^A-Za-z0-9_.-]', '_')
if ($globalGateName.Length -gt 120) { $globalGateName = $globalGateName.Substring(0, 120) }
$globalGatePath = Join-Path $globalGateRoot $globalGateName
$globalGateMutexName = "Local\SmartM365_GlobalWorkerGate_$globalGateName"
$launcherInstanceId = [guid]::NewGuid().ToString('N')
if ($GlobalConcurrencyLimit -gt 0) {
    New-Directory -Path $globalGatePath
    Write-Host ("Global lease gate: Limit={0}; Path={1}; LeaseTimeout={2} minute(s)" -f $GlobalConcurrencyLimit,$globalGatePath,$GlobalConcurrencyLeaseTimeoutMinutes) -ForegroundColor Green
}
$script:globalGateMutex = $null
if ($GlobalConcurrencyLimit -gt 0) {
    $script:globalGateMutex = New-Object System.Threading.Mutex($false, $globalGateMutexName)
}

function Invoke-WithGlobalGateMutex {
    param([Parameter(Mandatory = $true)][scriptblock]$ScriptBlock)
    $sharedMutex = $script:globalGateMutex
    $mutex = if ($null -ne $sharedMutex) { $sharedMutex } else { New-Object System.Threading.Mutex($false, $globalGateMutexName) }
    $ownMutex = ($null -eq $sharedMutex)
    $acquired = $false
    try {
        try { $acquired = $mutex.WaitOne(30000) }
        catch [System.Threading.AbandonedMutexException] { $acquired = $true }
        if (-not $acquired) { throw "Could not acquire global gate mutex within 30 seconds: $globalGateMutexName" }
        & $ScriptBlock
    }
    finally {
        if ($acquired) { try { $mutex.ReleaseMutex() } catch { } }
        if ($ownMutex) { $mutex.Dispose() }
    }
}

function Test-GateProcessAlive {
    param([Parameter(Mandatory = $true)][int]$ProcessId)
    try { $null = Get-Process -Id $ProcessId -ErrorAction Stop; return $true } catch { return $false }
}

function Remove-StaleGlobalLeases {
    $nowUtc = (Get-Date).ToUniversalTime()
    foreach ($lease in @(Get-ChildItem -LiteralPath $globalGatePath -Filter '*.json' -File -ErrorAction SilentlyContinue)) {
        $remove = $false
        $reason = ''
        $computerName = ''
        $launcherPid = 0
        $workerPid = 0
        $ageMinutes = 0
        try {
            $data = Get-Content -LiteralPath $lease.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $computerName = [string]$data.Computer
            $launcherPid = if ($data.PSObject.Properties['LauncherProcessId']) { [int]$data.LauncherProcessId } else { [int]$data.ProcessId }
            $workerPid = if ($data.PSObject.Properties['WorkerProcessId']) { [int]$data.WorkerProcessId } else { 0 }
            $createdUtc = [datetime]::Parse([string]$data.CreatedUtc, $null, [System.Globalization.DateTimeStyles]::RoundtripKind)
            $ageMinutes = [math]::Round(($nowUtc - $createdUtc).TotalMinutes, 1)
            if (-not (Test-GateProcessAlive -ProcessId $launcherPid)) {
                $remove = $true
                $reason = 'LauncherProcessExited'
            }
            elseif ($workerPid -gt 0 -and -not (Test-GateProcessAlive -ProcessId $workerPid)) {
                $remove = $true
                $reason = 'WorkerProcessExited'
            }
            elseif (($nowUtc - $createdUtc).TotalMinutes -gt $GlobalConcurrencyLeaseTimeoutMinutes) {
                $remove = $true
                $reason = 'LeaseExpired'
            }
        }
        catch {
            $remove = $true
            $reason = 'InvalidLease'
        }
        if ($remove) {
            Write-Host ("Removed stale global worker lease: Reason={0}; Computer={1}; LauncherPid={2}; WorkerPid={3}; AgeMinutes={4}; Path={5}" -f $reason,$computerName,$launcherPid,$workerPid,$ageMinutes,$lease.FullName) -ForegroundColor DarkYellow
            Remove-Item -LiteralPath $lease.FullName -Force -ErrorAction SilentlyContinue
        }
    }
}

function Update-GlobalLease {
    param(
        [AllowNull()][string]$LeasePath,
        [hashtable]$Properties = @{}
    )
    if ([string]::IsNullOrWhiteSpace($LeasePath) -or -not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return }
    Invoke-WithGlobalGateMutex -ScriptBlock {
        if (-not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return }
        $data = Get-Content -LiteralPath $LeasePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        foreach ($key in @($Properties.Keys)) {
            $data | Add-Member -NotePropertyName $key -NotePropertyValue $Properties[$key] -Force
        }
        $data | Add-Member -NotePropertyName LastUpdatedUtc -NotePropertyValue ((Get-Date).ToUniversalTime().ToString('o')) -Force
        $data | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $LeasePath -Encoding UTF8 -Force
    }
}

function Acquire-GlobalLease {
    param(
        [Parameter(Mandatory = $true)][string]$Computer,
        [Parameter(Mandatory = $true)][int]$CycleNumber
    )
    if ($GlobalConcurrencyLimit -lt 1) { return '' }
    $waitStarted = Get-Date
    $lastWaitLog = $null
    $waitLogDelaySeconds = 300
    $waitLogIntervalSeconds = 300
    $waitWasLogged = $false
    while ($true) {
        $leasePath = Invoke-WithGlobalGateMutex -ScriptBlock {
            Remove-StaleGlobalLeases
            $leases = @(Get-ChildItem -LiteralPath $globalGatePath -Filter '*.json' -File -ErrorAction SilentlyContinue)
            if ($leases.Count -lt $GlobalConcurrencyLimit) {
                $path = Join-Path $globalGatePath ("lease_{0}_{1}.json" -f $PID,([guid]::NewGuid().ToString('N')))
                [pscustomobject]@{
                    LauncherInstanceId = $launcherInstanceId
                    ProcessId = $PID
                    LauncherProcessId = $PID
                    Computer = $Computer
                    Cycle = $CycleNumber
                    JobId = ''
                    JobName = ''
                    WorkerProcessId = 0
                    WorkerProcessName = ''
                    WorkerStartedUtc = ''
                    CreatedUtc = (Get-Date).ToUniversalTime().ToString('o')
                    LastUpdatedUtc = (Get-Date).ToUniversalTime().ToString('o')
                    Host = $env:COMPUTERNAME
                } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $path -Encoding UTF8 -Force
                return $path
            }
            return ''
        }
        if (-not [string]::IsNullOrWhiteSpace($leasePath)) {
            if ($waitWasLogged) {
                $waitMinutes = [math]::Round(((Get-Date) - $waitStarted).TotalMinutes, 1)
                Write-Host ("Global worker lease acquired for {0} after {1} minute(s)." -f $Computer,$waitMinutes) -ForegroundColor DarkCyan
            }
            return $leasePath
        }
        $now = Get-Date
        $waitSeconds = ($now - $waitStarted).TotalSeconds
        if ($waitSeconds -ge $waitLogDelaySeconds -and ($null -eq $lastWaitLog -or (($now - $lastWaitLog).TotalSeconds -ge $waitLogIntervalSeconds))) {
            $activeLeases = @(Get-ChildItem -LiteralPath $globalGatePath -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
            $waitMinutes = [math]::Round(($now - $waitStarted).TotalMinutes, 1)
            Write-Host ("Waiting for global worker lease: Computer={0}; Active={1}; Limit={2}; Wait={3} minute(s)." -f $Computer,$activeLeases,$GlobalConcurrencyLimit,$waitMinutes) -ForegroundColor DarkYellow
            $lastWaitLog = $now
            $waitWasLogged = $true
        }
        Start-Sleep -Seconds $JobPollSeconds
    }
}

function Test-SampleDnsResolution {
    param(
        [string[]]$ComputerNames,
        [int]$SampleSize = 5
    )

    $sample = @($ComputerNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First $SampleSize)
    $tested = 0
    $failed = 0
    foreach ($computer in $sample) {
        $tested++
        try {
            $addresses = [System.Net.Dns]::GetHostAddresses($computer)
            if ($null -eq $addresses -or $addresses.Count -lt 1) { $failed++ }
        }
        catch {
            $failed++
        }
    }

    [pscustomobject]@{
        Tested = $tested
        Failed = $failed
        AllFailed = ($tested -gt 0 -and $failed -eq $tested)
    }
}

function Release-GlobalLease {
    param([AllowNull()][string]$LeasePath)
    if (-not [string]::IsNullOrWhiteSpace($LeasePath)) {
        Remove-Item -LiteralPath $LeasePath -Force -ErrorAction SilentlyContinue
    }
}

$cycle = 0
do {
    $cycle++
    $computers = @(Get-ComputerList -Path $ComputerListPath)
    if ($computers.Count -eq 0) {
        Write-Host "No computers found in $ComputerListPath." -ForegroundColor Yellow
        break
    }

    $dnsCheck = Test-SampleDnsResolution -ComputerNames $computers
    if ($dnsCheck.AllFailed) {
        Write-Host ""
        Write-Host ("*** DNS WARNING: resolution failed on all {0} sampled computers. Check VPN connectivity and DNS configuration on this machine before continuing. ***" -f $dnsCheck.Tested) -ForegroundColor Red
        Write-Host ""
    }

    Write-Host ("Cycle {0}: {1} computer(s)." -f $cycle,$computers.Count) -ForegroundColor Cyan
    $results = New-Object System.Collections.ArrayList
    $runningJobs = @()
    $globalLeaseByJobId = @{}
    $nextIndex = 0

    while ($nextIndex -lt $computers.Count -or $runningJobs.Count -gt 0) {
        while ($nextIndex -lt $computers.Count -and $runningJobs.Count -lt $ThrottleLimit) {
            $computer = $computers[$nextIndex]
            $nextIndex++

            $globalLeasePath = Acquire-GlobalLease -Computer $computer -CycleNumber $cycle

            try {
                $remoteArgsJson = ([pscustomobject]@{ Args = @($remoteArgs.ToArray()) } | ConvertTo-Json -Compress)
                $workerArgs = @(
                    $computer,
                    $cycle,
                    $remoteArgsJson,
                    $resolvedPsExec,
                    $LocalScriptPath,
                    $script:RemoteBaseDir,
                    $script:RemoteScriptPath,
                    $script:RemoteSetupCacheRoot,
                    $LogRoot,
                    $CentralLogRoot,
                    $SetupSourcePath,
                    $SetupSourceMapPath,
                    $SetupExecutionMode,
                    $SetupMediaId,
                    $SetupLanguage,
                    [bool]$AllowSetupUpgrade,
                    [bool]$SkipSetupMediaPreCopy,
                    [bool]$SkipVirtualMachines,
                    [bool]$DryRun,
                    [bool]$NoCentralLogCollection,
                    [bool]$KeepCentralLogHistory,
                    $PsExecTimeoutMinutes,
                    $globalLeasePath,
                    $globalGateMutexName
                )

                $job = Start-Job -Name ("W11UT_C{0}_{1}" -f $cycle,$computer) -FilePath $LocalWorkerPath -ArgumentList $workerArgs
            }
            catch {
                Release-GlobalLease -LeasePath $globalLeasePath
                throw
            }

            if (-not [string]::IsNullOrWhiteSpace($globalLeasePath)) {
                Update-GlobalLease -LeasePath $globalLeasePath -Properties @{
                    JobId = [string]$job.Id
                    JobName = [string]$job.Name
                }
                $globalLeaseByJobId[[string]$job.Id] = $globalLeasePath
            }
            $runningJobs += $job
            Write-Host ("Queued {0} ({1}/{2}); running={3}" -f $computer,$nextIndex,$computers.Count,$runningJobs.Count) -ForegroundColor DarkCyan

            if ($DelayBetweenComputersSeconds -gt 0 -and $nextIndex -lt $computers.Count) {
                Start-Sleep -Seconds $DelayBetweenComputersSeconds
            }
        }

        $finishedJobs = @($runningJobs | Where-Object { $_.State -ne 'Running' })
        if ($finishedJobs.Count -eq 0) {
            Start-Sleep -Seconds $JobPollSeconds
            continue
        }

        foreach ($job in $finishedJobs) {
            $received = $null
            $jobErrors = @()

            try {
                $receiveErrors = @()
                $received = @(Receive-Job -Job $job -ErrorAction SilentlyContinue -ErrorVariable receiveErrors)
                $childJobs = @($job.ChildJobs)
                $brokenRunspace = (
                    $childJobs.Count -gt 0 -and
                    $null -ne $childJobs[0].JobStateInfo.Reason -and
                    [string]$childJobs[0].JobStateInfo.Reason.Message -match '\bBroken\b'
                )
                $jobErrors = @(
                    @(
                        $receiveErrors | ForEach-Object { $_.ToString() }
                        $childJobs | ForEach-Object { $_.Error } | ForEach-Object { $_.ToString() }
                        if ($childJobs.Count -gt 0 -and $childJobs[0].JobStateInfo.Reason) {
                            $childJobs[0].JobStateInfo.Reason.Message
                        }
                        if ($brokenRunspace) {
                            'RUNSPACE_BROKEN: PowerShell job runspace entered a Broken state. This typically indicates PowerShell 5.1 instability under parallel load. Consider adding -DelayBetweenComputersSeconds 1 to reduce job cycling speed.'
                        }
                    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
                )
                if ($brokenRunspace) {
                    Write-Host ("  [RUNSPACE_BROKEN] {0}: job runspace entered Broken state - possible PowerShell 5.1 instability. Consider -DelayBetweenComputersSeconds 1." -f ($job.Name -replace '^W11UT_C\d+_','')) -ForegroundColor Magenta
                }
            }
            catch {
                $jobErrors += $_.Exception.Message
            }

            if (-not $received -or $received.Count -eq 0) {
                $received = @([pscustomobject]@{
                    Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                    ComputerName = ($job.Name -replace '^W11UT_C\d+_','')
                    CycleNumber = $cycle
                    LauncherStatus = if (@($jobErrors | Where-Object { $_ -match 'RUNSPACE_BROKEN' }).Count -gt 0) { 'RUNSPACE_BROKEN' } else { 'JOB_ERROR' }
                    RemoteStatus = ''
                    RemoteNextAction = ''
                    ExitCode = ''
                    Detail = ($jobErrors -join ' | ')
                    SetupCacheAction = ''
                    DiskCleanupAction = ''
                    DiskCleanupFreedGB = ''
                    AdvancedDiskCleanupAction = ''
                    AdvancedDiskCleanupFreedGB = ''
                    DismCleanupAction = ''
                    DismCleanupFreedGB = ''
                    SetupCompletionRebootAction = ''
                    SetupCompletionRebootDetail = ''
                    SetupCompletionRebootUserCount = ''
                    SetupCompletionRebootUsers = ''
                    ControlledRebootAction = ''
                    ControlledRebootDetail = ''
                    ControlledRebootUserCount = ''
                    ControlledRebootUsers = ''
                    RemoteLogsPath = ''
                    PsExecLogPath = ''
                    JobErrorMessage = ($jobErrors -join ' | ')
                })
            }

            foreach ($item in @($received)) {
                if ($null -ne $item) {
                    if ($jobErrors.Count -gt 0 -and -not $item.PSObject.Properties['JobErrorMessage']) {
                        $item | Add-Member -NotePropertyName JobErrorMessage -NotePropertyValue ($jobErrors -join ' | ') -Force
                    }
                    [void]$results.Add($item)
                }
            }
            if ($globalLeaseByJobId.ContainsKey([string]$job.Id)) {
                Release-GlobalLease -LeasePath $globalLeaseByJobId[[string]$job.Id]
                $globalLeaseByJobId.Remove([string]$job.Id)
            }
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }

        $runningJobs = @($runningJobs | Where-Object { $_.State -eq 'Running' })
    }

    $reportPath = Join-Path $ReportRoot ("PsExec_Windows11Upgrade_Summary_cycle{0}_{1}.csv" -f $cycle,(Get-Date -Format 'yyyyMMdd-HHmmss'))
    $normalizedResults = foreach ($item in @($results.ToArray())) {
        $row = [ordered]@{}
        foreach ($column in $reportColumns) {
            $row[$column] = if ($null -ne $item -and $item.PSObject.Properties[$column]) { [string]$item.$column } else { '' }
        }
        [pscustomobject]$row
    }
    @($normalizedResults) | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8
    Write-Host ("Cycle {0} report: {1}" -f $cycle,$reportPath) -ForegroundColor Green

    try {
        $moveAlreadyW11Result = Move-AlreadyWindows11ComputersFromList -ComputerListPath $ComputerListPath -CycleSummary @($normalizedResults)
        if ($moveAlreadyW11Result.Moved -gt 0) {
            Write-Host ("Cycle {0}: moved {1} already-Windows11 computer(s) to {2}" -f $cycle,$moveAlreadyW11Result.Moved,$moveAlreadyW11Result.AlreadyWindows11Path) -ForegroundColor Green
        }
    }
    catch {
        Write-Host ("Cycle {0}: failed to update ComputersAlreadyW11.txt: {1}" -f $cycle,$_.Exception.Message) -ForegroundColor Yellow
    }

    $sourceDistribution = @(
        $normalizedResults |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.SelectedSetupSourcePath) } |
            Group-Object -Property SelectedSetupSourcePath |
            Sort-Object Count -Descending |
            ForEach-Object {
                [pscustomobject]@{
                    SelectedSetupSourcePath = [string]$_.Name
                    ComputerCount = [int]$_.Count
                }
            }
    )
    if ($sourceDistribution.Count -gt 0) {
        $distributionPath = Join-Path $ReportRoot ("SetupSource_Distribution_cycle{0}_{1}.csv" -f $cycle,(Get-Date -Format 'yyyyMMdd-HHmmss'))
        $sourceDistribution | Export-Csv -LiteralPath $distributionPath -NoTypeInformation -Encoding UTF8
        Write-Host ("Setup source distribution: {0}" -f $distributionPath) -ForegroundColor Green
        foreach ($sourceGroup in $sourceDistribution) {
            Write-Host ("  {0}: {1}" -f $sourceGroup.SelectedSetupSourcePath,$sourceGroup.ComputerCount) -ForegroundColor DarkGreen
        }
    }

    if ($RunOnce) { break }
    if ($MaxCycles -gt 0 -and $cycle -ge $MaxCycles) { break }
    if ($DelayBetweenCyclesMinutes -gt 0) {
        Write-Host ("Waiting {0} minute(s) before next cycle." -f $DelayBetweenCyclesMinutes) -ForegroundColor DarkYellow
        Start-Sleep -Seconds ($DelayBetweenCyclesMinutes * 60)
    }
}
while ($true)

if ($null -ne $script:globalGateMutex) {
    try { $script:globalGateMutex.Dispose() } catch { }
    $script:globalGateMutex = $null
}
