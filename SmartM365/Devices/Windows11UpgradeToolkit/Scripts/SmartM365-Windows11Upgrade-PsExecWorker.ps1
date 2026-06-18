<#
.SYNOPSIS
    Internal worker used by SmartM365-Invoke-Windows11UpgradeRepairWithPsExec.ps1.

.DESCRIPTION
    Runs one remote computer operation in a background job. This file is operator-side only;
    the target device still receives only SmartM365-Invoke-Windows11UpgradeRepair.ps1.
#>

#requires -Version 5.1

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$Computer,
    [Parameter(Mandatory = $true)][int]$CycleNumber,
    [Parameter(Mandatory = $true)][string]$RemoteScriptArgsJson,
    [string]$ResolvedPsExecPath,
    [Parameter(Mandatory = $true)][string]$LocalScriptPath,
    [Parameter(Mandatory = $true)][string]$RemoteBaseDir,
    [Parameter(Mandatory = $true)][string]$RemoteScriptPath,
    [Parameter(Mandatory = $true)][string]$RemoteSetupCacheRoot,
    [Parameter(Mandatory = $true)][string]$LogRoot,
    [Parameter(Mandatory = $true)][string]$CentralLogRoot,
    [string]$SetupSourcePath,
    [string]$SetupSourceMapPath,
    [ValidateSet('LocalCache','Share','Auto')]
    [string]$SetupExecutionMode = 'LocalCache',
    [string]$SetupMediaId = 'Win11',
    [string]$SetupLanguage = 'MatchSystem',
    [bool]$AllowSetupUpgrade = $false,
    [bool]$SkipSetupMediaPreCopy = $false,
    [bool]$SkipVirtualMachines = $false,
    [bool]$DryRun = $false,
    [bool]$NoCentralLogCollection = $false,
    [bool]$KeepCentralLogHistory = $false,
    [int]$PsExecTimeoutMinutes = 180,
    [string]$GlobalWorkerLeasePath,
    [string]$GlobalWorkerLeaseMutexName
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

$RemoteScriptArgs = @()
if (-not [string]::IsNullOrWhiteSpace($RemoteScriptArgsJson)) {
    try {
        $decodedRemoteScriptArgs = $RemoteScriptArgsJson | ConvertFrom-Json -ErrorAction Stop
        if ($decodedRemoteScriptArgs -and $decodedRemoteScriptArgs.PSObject.Properties['Args']) {
            $RemoteScriptArgs = @($decodedRemoteScriptArgs.Args | ForEach-Object { [string]$_ })
        }
    }
    catch {
        throw ("Invalid remote script argument payload: {0}" -f $_.Exception.Message)
    }
}

function Invoke-WithWorkerLeaseMutex {
    param(
        [AllowNull()][string]$MutexName,
        [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock
    )

    if ([string]::IsNullOrWhiteSpace($MutexName)) {
        & $ScriptBlock
        return
    }

    $mutex = New-Object System.Threading.Mutex($false, $MutexName)
    $acquired = $false
    try {
        try { $acquired = $mutex.WaitOne(30000) }
        catch [System.Threading.AbandonedMutexException] { $acquired = $true }
        if ($acquired) { & $ScriptBlock }
    }
    finally {
        if ($acquired) { try { $mutex.ReleaseMutex() } catch { } }
        $mutex.Dispose()
    }
}

function Update-WorkerLease {
    param(
        [AllowNull()][string]$LeasePath,
        [AllowNull()][string]$MutexName
    )

    if ([string]::IsNullOrWhiteSpace($LeasePath) -or -not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return }

    try {
        Invoke-WithWorkerLeaseMutex -MutexName $MutexName -ScriptBlock {
            if (-not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return }
            $data = Get-Content -LiteralPath $LeasePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $processName = ''
            try { $processName = (Get-Process -Id $PID -ErrorAction Stop).ProcessName } catch { }
            $data | Add-Member -NotePropertyName WorkerProcessId -NotePropertyValue $PID -Force
            $data | Add-Member -NotePropertyName WorkerProcessName -NotePropertyValue $processName -Force
            $data | Add-Member -NotePropertyName WorkerStartedUtc -NotePropertyValue ((Get-Date).ToUniversalTime().ToString('o')) -Force
            $data | Add-Member -NotePropertyName LastUpdatedUtc -NotePropertyValue ((Get-Date).ToUniversalTime().ToString('o')) -Force
            $data | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $LeasePath -Encoding UTF8 -Force
        }
    }
    catch { }
}

Update-WorkerLease -LeasePath $GlobalWorkerLeasePath -MutexName $GlobalWorkerLeaseMutexName

function New-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
    }
}

function Convert-ToAdminSharePath {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][string]$LocalPath
    )

    $full = [System.IO.Path]::GetFullPath($LocalPath)
    if ($full -notmatch '^[A-Za-z]:\\') {
        throw "Only local drive paths can be converted to admin share paths: $LocalPath"
    }
    $drive = $full.Substring(0,1)
    $rest = $full.Substring(3)
    return "\\$ComputerName\$drive`$\$rest"
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
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][string]$LogPath
    )

    $remoteBaseShare = Convert-ToAdminSharePath -ComputerName $ComputerName -LocalPath $RemoteBaseDir
    New-Directory -Path $remoteBaseShare
    $remoteScriptShare = Join-Path $remoteBaseShare (Split-Path -Leaf $RemoteScriptPath)
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
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][string]$LogPath
    )

    if (-not $AllowSetupUpgrade) { return 'NotRequested' }
    if ($SkipSetupMediaPreCopy) { return 'ExistingMediaOnly' }
    if ($SetupExecutionMode -eq 'Share') { return 'ShareMode' }
    if ([string]::IsNullOrWhiteSpace($SetupSourcePath) -and [string]::IsNullOrWhiteSpace($SetupSourceMapPath)) { return 'TargetCacheNoSource' }

    Add-Content -LiteralPath $LogPath -Value ("[{0}] Setup media copy is target-side. SourcePath/Map will be validated from the target SYSTEM context: Source={1}; Map={2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$SetupSourcePath,$SetupSourceMapPath) -Encoding UTF8
    return 'TargetSideCache'
}

function Convert-ToPsExecRemoteArgument {
    param([AllowNull()][string]$Value)

    if ($null -eq $Value) { return '""' }
    $text = [string]$Value
    if ($text.Length -eq 0) { return '""' }
    if ($text -match '^[A-Za-z0-9_:\\./@=-]+$') { return $text }
    return ('"{0}"' -f ($text -replace '"', '\"'))
}

function Collect-RemoteEvidence {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][int]$Cycle
    )

    if ($NoCentralLogCollection) { return '' }

    $remoteBaseShare = Convert-ToAdminSharePath -ComputerName $ComputerName -LocalPath $RemoteBaseDir
    if (-not (Test-Path -LiteralPath $remoteBaseShare -PathType Container)) { return '' }

    $target = if ($KeepCentralLogHistory) {
        Join-Path (Join-Path $CentralLogRoot $ComputerName) ("Cycle{0}_{1}" -f $Cycle,(Get-Date -Format 'yyyyMMdd-HHmmss'))
    }
    else {
        Join-Path (Join-Path $CentralLogRoot $ComputerName) 'Latest'
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
    SetupDynamicUpdate = ''
    SelectedSetupSourcePath = ''
    SetupSourceSelectionDetail = ''
    DiskCleanupAction = ''
    DiskCleanupFreedGB = ''
    AdvancedDiskCleanupAction = ''
    AdvancedDiskCleanupFreedGB = ''
    DismCleanupAction = ''
    DismCleanupFreedGB = ''
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

    if ($SkipVirtualMachines) {
        Add-Content -LiteralPath $logPath -Value ("[{0}] Launcher-side VM pre-check skipped to avoid WinRM. Endpoint guard will check locally under PsExec/SYSTEM." -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) -Encoding UTF8
    }

    Copy-RemotePayload -ComputerName $Computer -LogPath $logPath
    $result.SetupCacheAction = Copy-SetupMediaToRemoteCache -ComputerName $Computer -LogPath $logPath

    $remotePowerShellArgs = @($RemoteScriptArgs | ForEach-Object { Convert-ToPsExecRemoteArgument -Value $_ })
    $psexecArgs = @(
        "\\$Computer",
        '-accepteula',
        '-nobanner',
        '-s',
        'powershell.exe',
        '-NoProfile',
        '-ExecutionPolicy','Bypass',
        '-File',(Convert-ToPsExecRemoteArgument -Value $RemoteScriptPath)
    ) + $remotePowerShellArgs

    Add-Content -LiteralPath $logPath -Value ("[{0}] PsExec command: {1} {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$ResolvedPsExecPath,($psexecArgs -join ' ')) -Encoding UTF8
    $process = Start-Process -FilePath $ResolvedPsExecPath -ArgumentList $psexecArgs -PassThru -NoNewWindow -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
    $psExecTimedOut = $false
    $waitStarted = Get-Date
    $lastWaitLog = $waitStarted
    Add-Content -LiteralPath $logPath -Value ("[{0}] PsExec started. Waiting for remote script completion; heartbeat every 60 seconds." -f $waitStarted.ToString('yyyy-MM-dd HH:mm:ss')) -Encoding UTF8
    if ($PsExecTimeoutMinutes -gt 0) {
        $timeoutAt = $waitStarted.AddMinutes($PsExecTimeoutMinutes)
        while (-not $process.WaitForExit(1000)) {
            $now = Get-Date
            if ($now -ge $timeoutAt) {
                $psExecTimedOut = $true
                Add-Content -LiteralPath $logPath -Value ("[{0}] ERROR PsExec timed out after {1} minute(s). Killing local PsExec process." -f $now.ToString('yyyy-MM-dd HH:mm:ss'),$PsExecTimeoutMinutes) -Encoding UTF8
                try { $process.Kill() } catch { }
                try { [void]$process.WaitForExit(5000) } catch { }
                $result.LauncherStatus = 'PSEXEC_TIMEOUT'
                $result.Detail = "Timed out after $PsExecTimeoutMinutes minute(s)."
                break
            }
            if (($now - $lastWaitLog).TotalSeconds -ge 60) {
                $elapsedMinutes = [math]::Round(($now - $waitStarted).TotalMinutes, 1)
                Add-Content -LiteralPath $logPath -Value ("[{0}] PsExec still running for {1} minute(s). Remote script may be copying setup media or running Windows setup." -f $now.ToString('yyyy-MM-dd HH:mm:ss'),$elapsedMinutes) -Encoding UTF8
                $lastWaitLog = $now
            }
        }
    }
    else {
        while (-not $process.WaitForExit(1000)) {
            $now = Get-Date
            if (($now - $lastWaitLog).TotalSeconds -ge 60) {
                $elapsedMinutes = [math]::Round(($now - $waitStarted).TotalMinutes, 1)
                Add-Content -LiteralPath $logPath -Value ("[{0}] PsExec still running for {1} minute(s). Remote script may be copying setup media or running Windows setup." -f $now.ToString('yyyy-MM-dd HH:mm:ss'),$elapsedMinutes) -Encoding UTF8
                $lastWaitLog = $now
            }
        }
    }

    try { $process.Refresh() } catch { }
    if ($result.LauncherStatus -eq 'UNKNOWN') {
        if (-not $psExecTimedOut -and $process.HasExited) {
            $result.ExitCode = [string]$process.ExitCode
            $result.LauncherStatus = if ($process.ExitCode -eq 0) { 'SUCCESS' } else { "PSEXEC_EXIT_$($process.ExitCode)" }
        }
    }

    $stdoutContent = @()
    if (Test-Path -LiteralPath $stdoutPath) {
        Add-Content -LiteralPath $logPath -Value '----- PsExec STDOUT -----' -Encoding UTF8
        $stdoutContent = @(Get-Content -LiteralPath $stdoutPath -ErrorAction SilentlyContinue)
        $stdoutContent | Add-Content -LiteralPath $logPath -Encoding UTF8
        Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue
    }
    $stderrContent = @()
    if (Test-Path -LiteralPath $stderrPath) {
        Add-Content -LiteralPath $logPath -Value '----- PsExec STDERR -----' -Encoding UTF8
        $stderrContent = @(Get-Content -LiteralPath $stderrPath -ErrorAction SilentlyContinue)
        $stderrContent | Add-Content -LiteralPath $logPath -Encoding UTF8
        Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue
    }
    if ([string]::IsNullOrWhiteSpace($result.ExitCode)) {
        $nativeExitLine = ($stderrContent | Where-Object { $_ -match 'with error code\s+-?\d+' } | Select-Object -Last 1)
        if ($nativeExitLine -and $nativeExitLine -match 'with error code\s+(?<Code>-?\d+)') {
            $result.ExitCode = $Matches.Code
            Add-Content -LiteralPath $logPath -Value ("[{0}] PsExec exit code recovered from native STDERR: {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$result.ExitCode) -Encoding UTF8
        }
    }

    Start-Sleep -Seconds 3
    $result.RemoteLogsPath = Collect-RemoteEvidence -ComputerName $Computer -Cycle $CycleNumber
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
                if ($lastRun.PSObject.Properties['SetupCacheAction']) {
                    $result.SetupCacheAction = [string]$lastRun.SetupCacheAction
                }
                foreach ($propertyName in @('SetupDynamicUpdate','SelectedSetupSourcePath','SetupSourceSelectionDetail','DiskCleanupAction','DiskCleanupFreedGB','AdvancedDiskCleanupAction','AdvancedDiskCleanupFreedGB','DismCleanupAction','DismCleanupFreedGB')) {
                    if ($lastRun.PSObject.Properties[$propertyName]) {
                        $result[$propertyName] = [string]$lastRun.$propertyName
                    }
                }
            }
        }
}
catch {
    $result.LauncherStatus = 'ERROR'
    $result.Detail = $_.Exception.Message
    Add-Content -LiteralPath $logPath -Value ("[{0}] ERROR {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$_.Exception.Message) -Encoding UTF8
}

[pscustomobject]$result
