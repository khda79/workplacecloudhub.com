<#
.SYNOPSIS
    Runs SmartM365-Invoke-Windows11UpgradeRepair.ps1 on remote computers using PsExec.

.DESCRIPTION
    LOT/PsExec orchestrator for Windows 10 to Windows 11 upgrade diagnostics and guarded repair.
    It copies the autonomous endpoint script to each target, optionally pre-caches Windows 11
    setup media locally on the target, starts the script as SYSTEM, collects evidence, and writes
    cycle CSV reports.

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Version: 0.1.0
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
    [switch]$AllowReboot,
    [switch]$SkipVirtualMachines,

    [string]$SetupSourcePath,
    [ValidateSet('LocalCache','Share','Auto')]
    [string]$SetupExecutionMode = 'LocalCache',
    [string]$SetupMediaId = 'Win11',
    [string]$SetupLanguage = 'MatchSystem',
    [switch]$SkipSetupMediaPreCopy,

    [string]$LogRoot,
    [string]$ReportRoot,
    [string]$CentralLogRoot,
    [switch]$NoCentralLogCollection,
    [switch]$KeepCentralLogHistory,

    [ValidateRange(1, 200)][int]$ThrottleLimit = 10,
    [ValidateRange(0, 200)][int]$GlobalConcurrencyLimit = 15,
    [string]$GlobalConcurrencySemaphoreName = 'Local\SmartM365_IntuneWindows11UpgradeToolkit_ComputerWorkers',
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

$script:LauncherVersion = '0.1.0'
$script:BaseDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$script:ToolkitRoot = Split-Path -Parent $script:BaseDir
if ([string]::IsNullOrWhiteSpace($LocalScriptPath)) {
    $LocalScriptPath = Join-Path $script:BaseDir 'SmartM365-Invoke-Windows11UpgradeRepair.ps1'
}
$LocalWorkerPath = Join-Path $script:BaseDir 'SmartM365-Windows11Upgrade-PsExecWorker.ps1'
if ([string]::IsNullOrWhiteSpace($LogRoot)) { $LogRoot = Join-Path (Split-Path -Parent $ComputerListPath) 'PsExecLogs' }
if ([string]::IsNullOrWhiteSpace($ReportRoot)) { $ReportRoot = Join-Path (Split-Path -Parent $ComputerListPath) 'Reports' }
if ([string]::IsNullOrWhiteSpace($CentralLogRoot)) { $CentralLogRoot = Join-Path (Split-Path -Parent $ComputerListPath) 'CentralLogs' }

$script:RemoteBaseDir = 'C:\ProgramData\SmartM365\IntuneWindows11UpgradeToolkit'
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

function Convert-ToAdminSharePath {
    param(
        [Parameter(Mandatory = $true)][string]$Computer,
        [Parameter(Mandatory = $true)][string]$LocalPath
    )

    $full = [System.IO.Path]::GetFullPath($LocalPath)
    if ($full -notmatch '^[A-Za-z]:\\') {
        throw "Only local drive paths can be converted to admin share paths: $LocalPath"
    }
    $drive = $full.Substring(0,1)
    $rest = $full.Substring(3)
    return "\\$Computer\$drive`$\$rest"
}

function Copy-DirectoryContent {
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$DestinationPath
    )

    New-Directory -Path $DestinationPath
    Copy-Item -Path (Join-Path $SourcePath '*') -Destination $DestinationPath -Recurse -Force -ErrorAction Stop
}

function Copy-RemotePayload {
    param(
        [Parameter(Mandatory = $true)][string]$Computer,
        [Parameter(Mandatory = $true)][string]$LogPath
    )

    $remoteBaseShare = Convert-ToAdminSharePath -Computer $Computer -LocalPath $script:RemoteBaseDir
    New-Directory -Path $remoteBaseShare
    $remoteScriptShare = Join-Path $remoteBaseShare (Split-Path -Leaf $script:RemoteScriptPath)
    Copy-Item -LiteralPath $LocalScriptPath -Destination $remoteScriptShare -Force -ErrorAction Stop

    $localHash = (Get-FileHash -LiteralPath $LocalScriptPath -Algorithm SHA256).Hash
    $remoteHash = (Get-FileHash -LiteralPath $remoteScriptShare -Algorithm SHA256).Hash
    if ($localHash -ne $remoteHash) {
        throw "Remote script copy hash mismatch."
    }

    Add-Content -LiteralPath $LogPath -Value ("[{0}] Remote script copied and verified." -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) -Encoding UTF8
}

function Copy-SetupMediaToRemoteCache {
    param(
        [Parameter(Mandatory = $true)][string]$Computer,
        [Parameter(Mandatory = $true)][string]$LogPath
    )

    if (-not $AllowSetupUpgrade) { return 'NotRequested' }
    if ($SkipSetupMediaPreCopy) { return 'Skipped' }
    if ($SetupExecutionMode -eq 'Share') { return 'ShareMode' }
    if ([string]::IsNullOrWhiteSpace($SetupSourcePath)) { return 'NoSource' }
    if (-not (Test-Path -LiteralPath $SetupSourcePath -PathType Container)) {
        throw "SetupSourcePath not reachable from operator workstation: $SetupSourcePath"
    }

    $setupExe = Join-Path $SetupSourcePath 'setup.exe'
    $installWim = Join-Path $SetupSourcePath 'sources\install.wim'
    $installEsd = Join-Path $SetupSourcePath 'sources\install.esd'
    if (-not (Test-Path -LiteralPath $setupExe -PathType Leaf)) { throw "setup.exe not found in SetupSourcePath: $SetupSourcePath" }
    if (-not (Test-Path -LiteralPath $installWim -PathType Leaf) -and -not (Test-Path -LiteralPath $installEsd -PathType Leaf)) {
        throw "Setup media missing sources\install.wim or sources\install.esd: $SetupSourcePath"
    }

    $remoteCache = Join-Path $script:RemoteSetupCacheRoot $SetupMediaId
    $remoteCacheShare = Convert-ToAdminSharePath -Computer $Computer -LocalPath $remoteCache
    $remoteSetupShare = Join-Path $remoteCacheShare 'setup.exe'
    if (Test-Path -LiteralPath $remoteSetupShare -PathType Leaf) {
        Add-Content -LiteralPath $LogPath -Value ("[{0}] Setup cache already has setup.exe: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$remoteCache) -Encoding UTF8
        return 'AlreadyCached'
    }

    Add-Content -LiteralPath $LogPath -Value ("[{0}] Copying setup media to remote cache: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$remoteCache) -Encoding UTF8
    Copy-DirectoryContent -SourcePath $SetupSourcePath -DestinationPath $remoteCacheShare
    return 'Copied'
}

function Collect-RemoteEvidence {
    param(
        [Parameter(Mandatory = $true)][string]$Computer,
        [Parameter(Mandatory = $true)][int]$CycleNumber
    )

    if ($NoCentralLogCollection) { return '' }

    $remoteBaseShare = Convert-ToAdminSharePath -Computer $Computer -LocalPath $script:RemoteBaseDir
    if (-not (Test-Path -LiteralPath $remoteBaseShare -PathType Container)) { return '' }

    $target = if ($KeepCentralLogHistory) {
        Join-Path (Join-Path $CentralLogRoot $Computer) ("Cycle{0}_{1}" -f $CycleNumber,(Get-Date -Format 'yyyyMMdd-HHmmss'))
    }
    else {
        Join-Path (Join-Path $CentralLogRoot $Computer) 'Latest'
    }

    if (Test-Path -LiteralPath $target) {
        Remove-Item -LiteralPath $target -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Directory -Path $target

    foreach ($child in @('Logs','Output','LastRun.json')) {
        $source = Join-Path $remoteBaseShare $child
        if (Test-Path -LiteralPath $source) {
            Copy-Item -LiteralPath $source -Destination $target -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    return $target
}

function Invoke-RemoteComputer {
    param(
        [Parameter(Mandatory = $true)][string]$Computer,
        [Parameter(Mandatory = $true)][int]$CycleNumber,
        [Parameter(Mandatory = $true)][string[]]$RemoteScriptArgs,
        [Parameter(Mandatory = $true)][string]$ResolvedPsExecPath
    )

    New-Directory -Path $LogRoot
    $logPath = Join-Path $LogRoot ("{0}_cycle{1}_{2}.log" -f $Computer,$CycleNumber,(Get-Date -Format 'yyyyMMdd-HHmmss'))
    $stdoutPath = "$logPath.stdout.txt"
    $stderrPath = "$logPath.stderr.txt"

    $result = [ordered]@{
        Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        ComputerName = $Computer
        CycleNumber = $CycleNumber
        LauncherStatus = 'UNKNOWN'
        RemoteStatus = ''
        RemoteNextAction = ''
        ExitCode = ''
        Detail = ''
        SetupCacheAction = ''
        RemoteLogsPath = ''
        PsExecLogPath = $logPath
    }

    try {
        Add-Content -LiteralPath $logPath -Value ("[{0}] Starting {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$Computer) -Encoding UTF8

        if ($DryRun) {
            $reachable = Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue
            $adminShare = Test-Path -LiteralPath ("\\{0}\C$" -f $Computer)
            $result.LauncherStatus = if ($reachable -and $adminShare) { 'DRYRUN_READY' } else { 'DRYRUN_UNREACHABLE' }
            $result.Detail = "Ping=$reachable; AdminShare=$adminShare"
            return [pscustomobject]$result
        }

        Copy-RemotePayload -Computer $Computer -LogPath $logPath
        $result.SetupCacheAction = Copy-SetupMediaToRemoteCache -Computer $Computer -LogPath $logPath

        $psexecArgs = @(
            "\\$Computer",
            '-accepteula',
            '-nobanner',
            '-s',
            'powershell.exe',
            '-NoProfile',
            '-ExecutionPolicy','Bypass',
            '-File',"`"$script:RemoteScriptPath`""
        ) + $RemoteScriptArgs

        Add-Content -LiteralPath $logPath -Value ("[{0}] PsExec command: {1} {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$ResolvedPsExecPath,($psexecArgs -join ' ')) -Encoding UTF8
        $process = Start-Process -FilePath $ResolvedPsExecPath -ArgumentList $psexecArgs -PassThru -NoNewWindow -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath

        if ($PsExecTimeoutMinutes -gt 0) {
            if (-not $process.WaitForExit($PsExecTimeoutMinutes * 60 * 1000)) {
                try { $process.Kill() } catch { }
                $result.LauncherStatus = 'PSEXEC_TIMEOUT'
                $result.Detail = "Timed out after $PsExecTimeoutMinutes minute(s)."
            }
        }
        else {
            $process.WaitForExit()
        }

        if ($result.LauncherStatus -eq 'UNKNOWN') {
            $result.ExitCode = $process.ExitCode
            $result.LauncherStatus = if ($process.ExitCode -eq 0) { 'SUCCESS' } else { "PSEXEC_EXIT_$($process.ExitCode)" }
        }

        Start-Sleep -Seconds 3
        $result.RemoteLogsPath = Collect-RemoteEvidence -Computer $Computer -CycleNumber $CycleNumber
        if ($result.RemoteLogsPath) {
            $lastRunPath = Join-Path $result.RemoteLogsPath 'LastRun.json'
            if (Test-Path -LiteralPath $lastRunPath -PathType Leaf) {
                $lastRun = Get-Content -LiteralPath $lastRunPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                $result.RemoteStatus = [string]$lastRun.Status
                $result.RemoteNextAction = [string]$lastRun.NextAction
                if ($result.RemoteStatus) {
                    $result.LauncherStatus = $result.RemoteStatus
                    $result.ExitCode = [string]$lastRun.ExitCode
                }
            }
        }
    }
    catch {
        $result.LauncherStatus = 'ERROR'
        $result.Detail = $_.Exception.Message
        Add-Content -LiteralPath $logPath -Value ("[{0}] ERROR {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$_.Exception.Message) -Encoding UTF8
    }

    return [pscustomobject]$result
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

$resolvedPsExec = if ($DryRun) { '' } else { Resolve-PsExecPath -Path $PsExecPath }

$remoteArgs = New-Object System.Collections.ArrayList
if ($AuditOnly) { [void]$remoteArgs.Add('-AuditOnly') }
if ($IgnoreRunGuard) { [void]$remoteArgs.Add('-IgnoreRunGuard') }
if ($AllowPolicyRepair) { [void]$remoteArgs.Add('-AllowPolicyRepair') }
if ($AllowWUReset) { [void]$remoteArgs.Add('-AllowWUReset') }
if ($AllowForceUpgrade) { [void]$remoteArgs.Add('-AllowForceUpgrade') }
if ($AllowSetupUpgrade) { [void]$remoteArgs.Add('-AllowSetupUpgrade') }
if ($AllowReboot) { [void]$remoteArgs.Add('-AllowReboot') }
if ($SkipVirtualMachines) { [void]$remoteArgs.Add('-SkipVirtualMachines') }
[void]$remoteArgs.Add('-SetupExecutionMode'); [void]$remoteArgs.Add($SetupExecutionMode)
[void]$remoteArgs.Add('-SetupMediaId'); [void]$remoteArgs.Add($SetupMediaId)
[void]$remoteArgs.Add('-SetupLanguage'); [void]$remoteArgs.Add($SetupLanguage)
[void]$remoteArgs.Add('-SetupCacheRoot'); [void]$remoteArgs.Add($script:RemoteSetupCacheRoot)
if (-not [string]::IsNullOrWhiteSpace($SetupSourcePath)) {
    [void]$remoteArgs.Add('-SetupSourcePath'); [void]$remoteArgs.Add($SetupSourcePath)
}

Write-Host "Smart Intune Windows 11 Upgrade Toolkit launcher v$script:LauncherVersion"
Write-Host "Computer list : $ComputerListPath"
Write-Host "PsExec        : $resolvedPsExec"
Write-Host "Repair script : $LocalScriptPath"
Write-Host "Worker script : $LocalWorkerPath"
Write-Host "Mode          : DryRun=$DryRun; AuditOnly=$AuditOnly; RunOnce=$RunOnce; SkipVirtualMachines=$SkipVirtualMachines"
Write-Host "Setup         : Allow=$AllowSetupUpgrade; Mode=$SetupExecutionMode; MediaId=$SetupMediaId; Language=$SetupLanguage; PreCopy=$(-not $SkipSetupMediaPreCopy)"
Write-Host "Parallelism   : ThrottleLimit=$ThrottleLimit; GlobalConcurrencyLimit=$GlobalConcurrencyLimit; GlobalLeaseTimeout=$GlobalConcurrencyLeaseTimeoutMinutes minute(s)"
Write-Host "Reports       : $ReportRoot"
Write-Host ""

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

function Invoke-WithGlobalGateMutex {
    param([Parameter(Mandatory = $true)][scriptblock]$ScriptBlock)
    $mutex = New-Object System.Threading.Mutex($false, $globalGateMutexName)
    $acquired = $false
    try {
        try { $acquired = $mutex.WaitOne(30000) }
        catch [System.Threading.AbandonedMutexException] { $acquired = $true }
        if (-not $acquired) { throw "Could not acquire global gate mutex within 30 seconds: $globalGateMutexName" }
        & $ScriptBlock
    }
    finally {
        if ($acquired) { try { $mutex.ReleaseMutex() } catch { } }
        $mutex.Dispose()
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
            $createdUtc = [datetime]$data.CreatedUtc
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
        if (-not [string]::IsNullOrWhiteSpace($leasePath)) { return $leasePath }
        Write-Host ("Waiting for global worker lease before queuing {0}. Active={1}; Limit={2}; Gate={3}" -f $Computer,(@(Get-ChildItem -LiteralPath $globalGatePath -Filter '*.json' -File -ErrorAction SilentlyContinue).Count),$GlobalConcurrencyLimit,$globalGatePath) -ForegroundColor DarkYellow
        Start-Sleep -Seconds $JobPollSeconds
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
                $workerArgs = @(
                    $computer,
                    $cycle,
                    (,([string[]]@($remoteArgs.ToArray()))),
                    $resolvedPsExec,
                    $LocalScriptPath,
                    $script:RemoteBaseDir,
                    $script:RemoteScriptPath,
                    $script:RemoteSetupCacheRoot,
                    $LogRoot,
                    $CentralLogRoot,
                    $SetupSourcePath,
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
            $jobResults = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
            foreach ($item in $jobResults) {
                [void]$results.Add($item)
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
    @($results.ToArray()) | Export-Csv -LiteralPath $reportPath -NoTypeInformation -Encoding UTF8
    Write-Host ("Cycle {0} report: {1}" -f $cycle,$reportPath) -ForegroundColor Green

    if ($RunOnce) { break }
    if ($MaxCycles -gt 0 -and $cycle -ge $MaxCycles) { break }
    if ($DelayBetweenCyclesMinutes -gt 0) {
        Write-Host ("Waiting {0} minute(s) before next cycle." -f $DelayBetweenCyclesMinutes) -ForegroundColor DarkYellow
        Start-Sleep -Seconds ($DelayBetweenCyclesMinutes * 60)
    }
}
while ($true)
