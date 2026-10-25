<#
.SYNOPSIS
    Technician-side concurrency: controlled stop, exclusive Wave runs and the shared worker cap.

.DESCRIPTION
    Scope of every mechanism in this file: ONE technician computer. State lives under
    %ProgramData%\WCH\SmartM365\SmartEndpointMigrationToolkit\LauncherState, restricted to
    SYSTEM and Administrators, and critical sections use exclusive file handles so they work
    across processes, Windows sessions and administrator accounts. Nothing here coordinates
    several technician computers.

    Worker reservations (leases) are JSON files recovered when the launcher or worker process
    is gone (PID plus process start time), when the worker heartbeat is stale, when the lease
    exceeds its absolute lifetime, or when the file is unreadable. When Runs declare different
    caps, the strictest cap among active reservations applies.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#region Controlled stop

function Initialize-SemtCancellation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$StateRoot,
        [Parameter(Mandatory = $true)][string]$Campaign,
        [Parameter(Mandatory = $true)][string]$WaveName,
        [Parameter(Mandatory = $true)][string]$RunId
    )

    if (-not ('SmartEndpointMigrationCancellation' -as [type])) {
        Add-Type -TypeDefinition @"
using System;
using System.Threading;
public static class SmartEndpointMigrationCancellation
{
    private static ConsoleCancelEventHandler handler;
    private static int requestCount;
    public static int RequestCount { get { return requestCount; } }
    public static void Register()
    {
        Unregister();
        requestCount = 0;
        handler = new ConsoleCancelEventHandler(OnCancelKeyPress);
        try { Console.CancelKeyPress += handler; } catch { }
    }
    public static void Unregister()
    {
        if (handler != null) { try { Console.CancelKeyPress -= handler; } catch { } handler = null; }
    }
    private static void OnCancelKeyPress(object sender, ConsoleCancelEventArgs e)
    {
        e.Cancel = true;
        Interlocked.Increment(ref requestCount);
    }
}
"@
    }
    [SmartEndpointMigrationCancellation]::Register()
    New-SemtDirectory -Path $StateRoot
    $script:SemtCancellation = [pscustomobject]@{
        StateRoot = $StateRoot
        SignalPath = Join-Path -Path $StateRoot -ChildPath ("StopRequested_{0}.json" -f $PID)
        ActiveRunPath = Join-Path -Path $StateRoot -ChildPath ("ActiveRun_{0}.json" -f $PID)
        Campaign = $Campaign
        WaveName = $WaveName
        RunId = $RunId
        StartedUtc = Get-SemtUtcNowText
    }
    Remove-Item -LiteralPath $script:SemtCancellation.SignalPath -Force -ErrorAction SilentlyContinue
    return $script:SemtCancellation
}

function Get-SemtCancellationState {
    [CmdletBinding()]
    param()

    $count = 0
    try { $count = [SmartEndpointMigrationCancellation]::RequestCount } catch { $count = 0 }
    $requested = ($count -gt 0)
    $force = ($count -gt 1)
    $source = if ($requested) { 'Ctrl+C' } else { '' }
    $context = Get-Variable -Name SemtCancellation -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    if ($context -and (Test-Path -LiteralPath $context.SignalPath -PathType Leaf)) {
        $requested = $true
        $source = 'StopSignal'
        try {
            $signal = Read-SemtJsonFile -Path $context.SignalPath
            if ($signal -and [bool](Get-SemtPropertyValue -InputObject $signal -Name 'Force' -Default $false)) { $force = $true }
            $signalSource = [string](Get-SemtPropertyValue -InputObject $signal -Name 'Source' -Default '')
            if (-not [string]::IsNullOrWhiteSpace($signalSource)) { $source = $signalSource }
        }
        catch { }
    }
    return [pscustomobject]@{ Requested = $requested; Force = $force; Source = $source; RequestCount = $count }
}

function Wait-SemtCancellationAware {
    [CmdletBinding()]
    param(
        [ValidateRange(0, 864000)][int]$Seconds,
        [scriptblock]$OnTick
    )

    $remaining = $Seconds
    while ($remaining -gt 0) {
        if ((Get-SemtCancellationState).Requested) { return $false }
        $slice = [math]::Min(2, $remaining)
        Start-Sleep -Seconds $slice
        $remaining -= $slice
        if ($OnTick) { & $OnTick }
    }
    return $true
}

function Set-SemtActiveRunState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Status,
        [string]$ReportPath = '',
        [string]$Detail = ''
    )

    $context = Get-Variable -Name SemtCancellation -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    if (-not $context) { return }
    Write-SemtJsonFileAtomic -Path $context.ActiveRunPath -InputObject ([pscustomobject]@{
        SchemaVersion = 1
        Product = $script:SemtProductName
        Campaign = $context.Campaign
        WaveName = $context.WaveName
        RunId = $context.RunId
        Status = $Status
        Detail = $Detail
        ProcessId = $PID
        ProcessStartUtc = Get-SemtProcessStartUtc -ProcessId $PID
        SignalPath = $context.SignalPath
        ReportPath = $ReportPath
        StartedUtc = $context.StartedUtc
        UpdatedUtc = Get-SemtUtcNowText
    })
}

function Complete-SemtCancellation {
    [CmdletBinding()]
    param([switch]$KeepActiveRunState)

    try { [SmartEndpointMigrationCancellation]::Unregister() } catch { }
    $context = Get-Variable -Name SemtCancellation -Scope Script -ValueOnly -ErrorAction SilentlyContinue
    if ($context) {
        if (-not $KeepActiveRunState) { Remove-Item -LiteralPath $context.ActiveRunPath -Force -ErrorAction SilentlyContinue }
        Remove-Item -LiteralPath $context.SignalPath -Force -ErrorAction SilentlyContinue
    }
}

function Get-SemtActiveRuns {
    <#
    Lists Runs currently executing under a package Runs folder (used by the GUI and by
    Launch All). A state file whose process is gone is reported as Stale.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RunsRoot,
        [string]$Campaign
    )

    $rows = New-Object System.Collections.Generic.List[object]
    if (-not (Test-Path -LiteralPath $RunsRoot -PathType Container)) { return @() }
    foreach ($file in @(Get-ChildItem -LiteralPath $RunsRoot -Recurse -Filter 'ActiveRun_*.json' -File -ErrorAction SilentlyContinue)) {
        $data = $null
        try { $data = Read-SemtJsonFile -Path $file.FullName } catch { $data = $null }
        if (-not $data) { continue }
        if (-not [string]::IsNullOrWhiteSpace($Campaign) -and [string]$data.Campaign -ne $Campaign) { continue }
        $alive = Test-SemtProcessAlive -ProcessId ([int]$data.ProcessId) -ExpectedStartUtc ([string](Get-SemtPropertyValue -InputObject $data -Name 'ProcessStartUtc' -Default ''))
        $rows.Add([pscustomobject]@{
            Campaign = [string]$data.Campaign
            WaveName = [string]$data.WaveName
            RunId = [string]$data.RunId
            Status = if ($alive) { [string]$data.Status } else { 'Stale' }
            ProcessId = [int]$data.ProcessId
            Alive = [bool]$alive
            SignalPath = [string]$data.SignalPath
            ReportPath = [string]$data.ReportPath
            StatePath = $file.FullName
            UpdatedUtc = [string]$data.UpdatedUtc
        })
    }
    return @($rows.ToArray())
}

function Request-SemtRunStop {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$ActiveRun,
        [switch]$Force,
        [string]$Source = 'GUI'
    )

    if (-not $ActiveRun.Alive) { return 'NotRunning' }
    Write-SemtJsonFileAtomic -Path $ActiveRun.SignalPath -InputObject ([pscustomobject]@{
        RequestedUtc = Get-SemtUtcNowText
        Source = $Source
        Force = [bool]$Force
        RequestedBy = [Environment]::UserName
    })
    return $(if ($Force) { 'ForceStopRequested' } else { 'StopRequested' })
}

#endregion

#region Exclusive Wave run

function Enter-SemtWaveLock {
    <#
    Prevents two Runs of the same Wave on this technician computer. The exclusive handle is
    held for the lifetime of the Run and released by the operating system if the process dies.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Campaign,
        [Parameter(Mandatory = $true)][string]$WavePath
    )

    $root = Join-Path -Path (Get-SemtLauncherStateRoot) -ChildPath 'WaveLocks'
    New-SemtDirectory -Path $root
    $key = (Get-SemtStringSha256 -Text ('{0}|{1}' -f $Campaign,[System.IO.Path]::GetFullPath($WavePath).TrimEnd('\').ToUpperInvariant())).Substring(0, 32)
    $lockPath = Join-Path -Path $root -ChildPath ("{0}.lock" -f $key)
    try {
        $stream = [System.IO.File]::Open($lockPath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::Read)
    }
    catch {
        $owner = ''
        try {
            $reader = New-Object System.IO.StreamReader([System.IO.File]::Open($lockPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite))
            try { $owner = $reader.ReadToEnd() } finally { $reader.Dispose() }
        }
        catch { $owner = '' }
        throw ("SEMT_WAVE_ALREADY_RUNNING: this Wave is already running on this technician computer. Wave={0}; Owner={1}" -f $WavePath,$owner)
    }
    $ownerText = ('Campaign={0}; Wave={1}; ProcessId={2}; Account={3}; StartedUtc={4}' -f $Campaign,$WavePath,$PID,[Environment]::UserName,(Get-SemtUtcNowText))
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($ownerText)
    $stream.SetLength(0)
    $stream.Write($bytes, 0, $bytes.Length)
    $stream.Flush()
    return [pscustomobject]@{ Path = $lockPath; Stream = $stream }
}

function Exit-SemtWaveLock {
    [CmdletBinding()]
    param([AllowNull()]$WaveLock)

    if ($WaveLock -and $WaveLock.Stream) {
        try { $WaveLock.Stream.SetLength(0) } catch { }
        try { $WaveLock.Stream.Dispose() } catch { }
    }
}

#endregion

#region Shared worker cap

function Initialize-SemtWorkerGate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][int]$Limit,
        [Parameter(Mandatory = $true)][int]$LeaseTimeoutMinutes,
        [Parameter(Mandatory = $true)][int]$HeartbeatStaleMinutes,
        [string]$Root
    )

    if ([string]::IsNullOrWhiteSpace($Root)) { $Root = Join-Path -Path (Get-SemtLauncherStateRoot) -ChildPath 'WorkerGate' }
    New-SemtDirectory -Path $Root
    $aclStatus = 'NotApplied'
    try { $aclStatus = Protect-SemtLocalDirectoryAcl -Path (Split-Path -Path $Root -Parent) } catch { $aclStatus = ('Failed: {0}' -f $_.Exception.Message) }
    return [pscustomobject]@{
        Root = $Root
        LeaseRoot = (Join-Path -Path $Root -ChildPath 'Leases')
        LockPath = (Join-Path -Path $Root -ChildPath 'WorkerGate.lock')
        Limit = [int]$Limit
        LeaseTimeoutMinutes = [int]$LeaseTimeoutMinutes
        HeartbeatStaleMinutes = [int]$HeartbeatStaleMinutes
        LauncherInstanceId = [guid]::NewGuid().ToString('N')
        LauncherProcessId = $PID
        LauncherProcessStartUtc = Get-SemtProcessStartUtc -ProcessId $PID
        AclStatus = $aclStatus
    }
}

function Get-SemtWorkerLeaseStaleReason {
    [CmdletBinding()]
    param(
        [AllowNull()]$Lease,
        [Parameter(Mandatory = $true)][int]$LeaseTimeoutMinutes,
        [Parameter(Mandatory = $true)][int]$HeartbeatStaleMinutes,
        [datetime]$NowUtc = (Get-Date).ToUniversalTime()
    )

    if ($null -eq $Lease) { return 'InvalidLease' }
    $launcherPid = [int](Get-SemtPropertyValue -InputObject $Lease -Name 'LauncherProcessId' -Default 0)
    if (-not (Test-SemtProcessAlive -ProcessId $launcherPid -ExpectedStartUtc ([string](Get-SemtPropertyValue -InputObject $Lease -Name 'LauncherProcessStartUtc' -Default '')))) { return 'LauncherProcessExited' }
    $workerPid = [int](Get-SemtPropertyValue -InputObject $Lease -Name 'WorkerProcessId' -Default 0)
    if ($workerPid -gt 0 -and -not (Test-SemtProcessAlive -ProcessId $workerPid -ExpectedStartUtc ([string](Get-SemtPropertyValue -InputObject $Lease -Name 'WorkerProcessStartUtc' -Default '')))) { return 'WorkerProcessExited' }
    $created = ConvertTo-SemtUtcDateTime -Value (Get-SemtPropertyValue -InputObject $Lease -Name 'CreatedUtc' -Default '')
    if ($null -eq $created) { return 'InvalidLease' }
    if ($LeaseTimeoutMinutes -gt 0 -and ($NowUtc - $created).TotalMinutes -gt $LeaseTimeoutMinutes) { return 'LeaseExpired' }
    $heartbeat = ConvertTo-SemtUtcDateTime -Value (Get-SemtPropertyValue -InputObject $Lease -Name 'HeartbeatUtc' -Default '')
    if ($null -eq $heartbeat) { $heartbeat = $created }
    if ($HeartbeatStaleMinutes -gt 0 -and ($NowUtc - $heartbeat).TotalMinutes -gt $HeartbeatStaleMinutes) { return 'HeartbeatStale' }
    return ''
}

function Get-SemtWorkerGateSnapshotUnlocked {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Gate,
        [switch]$RemoveStale
    )

    New-SemtDirectory -Path $Gate.LeaseRoot
    $active = New-Object System.Collections.Generic.List[object]
    $removed = New-Object System.Collections.Generic.List[object]
    foreach ($file in @(Get-ChildItem -LiteralPath $Gate.LeaseRoot -Filter 'lease_*.json' -File -ErrorAction SilentlyContinue)) {
        $lease = $null
        try { $lease = Read-SemtJsonFile -Path $file.FullName } catch { $lease = $null }
        $reason = Get-SemtWorkerLeaseStaleReason -Lease $lease -LeaseTimeoutMinutes $Gate.LeaseTimeoutMinutes -HeartbeatStaleMinutes $Gate.HeartbeatStaleMinutes
        if ([string]::IsNullOrWhiteSpace($reason)) {
            $active.Add([pscustomobject]@{ Path = $file.FullName; Lease = $lease })
            continue
        }
        if ($RemoveStale) {
            try {
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                $removed.Add([pscustomobject]@{ Path = $file.FullName; Reason = $reason; Computer = [string](Get-SemtPropertyValue -InputObject $lease -Name 'Computer' -Default ''); Campaign = [string](Get-SemtPropertyValue -InputObject $lease -Name 'Campaign' -Default '') })
            }
            catch { $active.Add([pscustomobject]@{ Path = $file.FullName; Lease = $lease }) }
        }
        else {
            $active.Add([pscustomobject]@{ Path = $file.FullName; Lease = $lease })
        }
    }
    $limits = @($active | ForEach-Object { [int](Get-SemtPropertyValue -InputObject $_.Lease -Name 'Limit' -Default 0) } | Where-Object { $_ -gt 0 })
    $strictest = if ($limits.Count -gt 0) { [int](@($limits | Sort-Object)[0]) } else { 0 }
    return [pscustomobject]@{ Active = @($active.ToArray()); Removed = @($removed.ToArray()); StrictestActiveLimit = $strictest }
}

function Get-SemtWorkerGateSnapshot {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Gate)

    return (Invoke-SemtWithFileLock -LockPath $Gate.LockPath -TimeoutSeconds 120 -ArgumentList @($Gate) -ScriptBlock {
        param($LockedGate)
        Get-SemtWorkerGateSnapshotUnlocked -Gate $LockedGate
    })
}

function Request-SemtWorkerLease {
    <#
    Waits for a slot under the shared cap and returns the lease path, or '' when the cap is
    disabled or a controlled stop was requested while waiting.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Gate,
        [Parameter(Mandatory = $true)][string]$Campaign,
        [Parameter(Mandatory = $true)][string]$RunId,
        [Parameter(Mandatory = $true)][string]$WaveName,
        [Parameter(Mandatory = $true)][string]$Computer,
        [Parameter(Mandatory = $true)][int]$CycleNumber,
        [ValidateRange(1, 60)][int]$PollSeconds = 2,
        [scriptblock]$OnWaitLog
    )

    if ($Gate.Limit -lt 1) { return '' }
    $waitStarted = Get-Date
    $lastLog = $null
    while ($true) {
        if ((Get-SemtCancellationState).Requested) { return '' }
        $attempt = Invoke-SemtWithFileLock -LockPath $Gate.LockPath -TimeoutSeconds 300 -ArgumentList @($Gate,$Campaign,$RunId,$WaveName,$Computer,$CycleNumber) -ScriptBlock {
            param($LockedGate, $LockedCampaign, $LockedRunId, $LockedWave, $LockedComputer, $LockedCycle)
            $snapshot = Get-SemtWorkerGateSnapshotUnlocked -Gate $LockedGate -RemoveStale
            $effectiveLimit = $LockedGate.Limit
            if ($snapshot.StrictestActiveLimit -gt 0 -and $snapshot.StrictestActiveLimit -lt $effectiveLimit) { $effectiveLimit = $snapshot.StrictestActiveLimit }
            if ($snapshot.Active.Count -ge $effectiveLimit) {
                return [pscustomobject]@{ LeasePath = ''; Active = $snapshot.Active.Count; EffectiveLimit = $effectiveLimit; Removed = $snapshot.Removed }
            }
            $leaseId = [guid]::NewGuid().ToString('N')
            $path = Join-Path -Path $LockedGate.LeaseRoot -ChildPath ("lease_{0}_{1}.json" -f $PID,$leaseId)
            $now = Get-SemtUtcNowText
            $sessionId = -1
            try { $sessionId = [int][System.Diagnostics.Process]::GetCurrentProcess().SessionId } catch { $sessionId = -1 }
            Write-SemtJsonFileAtomic -Path $path -Depth 4 -InputObject ([pscustomobject]@{
                SchemaVersion = 1
                LeaseId = $leaseId
                Campaign = $LockedCampaign
                RunId = $LockedRunId
                WaveName = $LockedWave
                Computer = $LockedComputer
                CycleNumber = $LockedCycle
                Limit = $LockedGate.Limit
                LauncherInstanceId = $LockedGate.LauncherInstanceId
                LauncherProcessId = $LockedGate.LauncherProcessId
                LauncherProcessStartUtc = $LockedGate.LauncherProcessStartUtc
                LauncherAccount = [Environment]::UserName
                LauncherSessionId = $sessionId
                JobId = ''
                WorkerProcessId = 0
                WorkerProcessStartUtc = ''
                PsExecProcessId = 0
                PsExecProcessStartUtc = ''
                CreatedUtc = $now
                HeartbeatUtc = $now
                Host = $env:COMPUTERNAME
            })
            return [pscustomobject]@{ LeasePath = $path; Active = $snapshot.Active.Count + 1; EffectiveLimit = $effectiveLimit; Removed = $snapshot.Removed }
        }
        if ($OnWaitLog -and @($attempt.Removed).Count -gt 0) {
            & $OnWaitLog ("Recovered {0} stale worker reservation(s): {1}" -f @($attempt.Removed).Count,((@($attempt.Removed) | ForEach-Object { '{0}/{1}={2}' -f $_.Campaign,$_.Computer,$_.Reason }) -join ', '))
        }
        if (-not [string]::IsNullOrWhiteSpace($attempt.LeasePath)) { return $attempt.LeasePath }
        $now = Get-Date
        if ($OnWaitLog -and ($now - $waitStarted).TotalSeconds -ge 300 -and ($null -eq $lastLog -or ($now - $lastLog).TotalSeconds -ge 300)) {
            & $OnWaitLog ("Waiting for a shared worker slot: Computer={0}; Active={1}; EffectiveLimit={2}; Wait={3:N1} minute(s)." -f $Computer,$attempt.Active,$attempt.EffectiveLimit,($now - $waitStarted).TotalMinutes)
            $lastLog = $now
        }
        Start-Sleep -Seconds $PollSeconds
    }
}

function Update-SemtWorkerLease {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$LeasePath,
        [string]$LockPath,
        [hashtable]$Properties = @{}
    )

    if ([string]::IsNullOrWhiteSpace($LeasePath)) { return $false }
    if ([string]::IsNullOrWhiteSpace($LockPath)) { $LockPath = Join-Path -Path (Split-Path -Path (Split-Path -Path $LeasePath -Parent) -Parent) -ChildPath 'WorkerGate.lock' }
    try {
        return (Invoke-SemtWithFileLock -LockPath $LockPath -TimeoutSeconds 60 -ArgumentList @($LeasePath,$Properties) -ScriptBlock {
            param($LockedLeasePath, $LockedProperties)
            if (-not (Test-Path -LiteralPath $LockedLeasePath -PathType Leaf)) { return $false }
            $data = Read-SemtJsonFile -Path $LockedLeasePath -Required
            foreach ($key in @($LockedProperties.Keys)) { $data | Add-Member -NotePropertyName $key -NotePropertyValue $LockedProperties[$key] -Force }
            $data | Add-Member -NotePropertyName HeartbeatUtc -NotePropertyValue (Get-SemtUtcNowText) -Force
            Write-SemtJsonFileAtomic -Path $LockedLeasePath -InputObject $data -Depth 4
            return $true
        })
    }
    catch { return $false }
}

function Remove-SemtWorkerLease {
    [CmdletBinding()]
    param(
        [AllowNull()]$Gate,
        [AllowNull()][string]$LeasePath
    )

    if ([string]::IsNullOrWhiteSpace($LeasePath)) { return $true }
    $lockPath = if ($Gate) { [string]$Gate.LockPath } else { Join-Path -Path (Split-Path -Path (Split-Path -Path $LeasePath -Parent) -Parent) -ChildPath 'WorkerGate.lock' }
    try {
        return (Invoke-SemtWithFileLock -LockPath $lockPath -TimeoutSeconds 120 -ArgumentList @($LeasePath) -ScriptBlock {
            param($LockedLeasePath)
            for ($attempt = 1; $attempt -le 5; $attempt++) {
                if (-not (Test-Path -LiteralPath $LockedLeasePath -PathType Leaf)) { return $true }
                try { Remove-Item -LiteralPath $LockedLeasePath -Force -ErrorAction Stop } catch { Start-Sleep -Milliseconds (100 * $attempt) }
            }
            return (-not (Test-Path -LiteralPath $LockedLeasePath -PathType Leaf))
        })
    }
    catch { return $false }
}

function Stop-SemtWorkerProcessTree {
    <#
    Stops the local PsExec process tree recorded in a reservation (forced controlled stop).
    The remote endpoint may already be running: callers must report the remote state as
    unconfirmed until current-run evidence proves otherwise.
    #>
    [CmdletBinding()]
    param([AllowNull()][string]$LeasePath)

    if ([string]::IsNullOrWhiteSpace($LeasePath) -or -not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return 'No active worker reservation.' }
    try {
        $lease = Read-SemtJsonFile -Path $LeasePath
        $psExecPid = [int](Get-SemtPropertyValue -InputObject $lease -Name 'PsExecProcessId' -Default 0)
        if ($psExecPid -le 0) { return 'The reservation holds no PsExec process id.' }
        if (-not (Test-SemtProcessAlive -ProcessId $psExecPid -ExpectedStartUtc ([string](Get-SemtPropertyValue -InputObject $lease -Name 'PsExecProcessStartUtc' -Default '')))) { return ("PsExec process {0} had already exited." -f $psExecPid) }
        $output = & taskkill.exe /PID $psExecPid /T /F 2>&1
        return ("Local PsExec process tree stop requested. PID={0}; ExitCode={1}; Output={2}" -f $psExecPid,$LASTEXITCODE,((@($output) -join ' ').Trim()))
    }
    catch { return ("Local PsExec process tree stop failed: {0}" -f $_.Exception.Message) }
}

#endregion
