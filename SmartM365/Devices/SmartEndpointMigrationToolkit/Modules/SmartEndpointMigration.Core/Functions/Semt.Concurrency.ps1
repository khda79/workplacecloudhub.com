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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDZZbiIL+3FkIg2
# MOCKpr9AQU25Yw5PVntK4fJi9nlpjqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEILnZ8PAHH+H4B/sanqXXapNIRE1FPpwWtqIsqNyjPfziMA0GCSqG
# SIb3DQEBAQUABIIBgHTD1Kr3dXCrMmNXfdeqqFyX/EsKCcrepAMrLYA3eaap+Dzh
# wO73e6GEpTdmVsPvx+JIK6FPs/4cP+3fnub4ruVsyQuCw5sxMrR883mD9e8sKwTR
# BwWlJxZzIS43yBvjoCLdrwBpw6O3IpmIypVyHDt68XYtzWnDmHepxkZx2ySJZFd+
# rQqZuK4gKe3S+yH3Ivs52Ak5oWIVCFxI8XkRB9y1vr7C2dDdAno0xdsjpJQkGPeh
# ise9l5RrsS9I/SkHRSAWHhLwKP5eXSewovA4OmQgLuOqbVt2S1aO+GEI/dhsqYJc
# CfNMlJNJxdKpTAEGiROOwDW4JuAgknztrqNlT6bOJxrG4XpIR++R2E8i3jqr2iYc
# c93Qd47RPW3Ri2BfjDDdZlysQFMSzACLIkgAuaYomLH5wDI8lechBT8iMJ10YRek
# AGsvW1eqK17SKQtZ4jc2ooPYngbg4K5Ve7cmA4SvlWXcvD0tOT+YiUUnlE3Ok5l0
# Zft+enht+Nzz5sktDqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMzE2
# NTJaMC8GCSqGSIb3DQEJBDEiBCB8cA0Zu++erVhoYsDr2JahMf0PSAhThIg+ugDT
# 29jtrTANBgkqhkiG9w0BAQEFAASCAgAuEDOzkpfslKJD8i83Z924MIA4k7Dz/nrs
# RPHm/JoUJBS3jmH+mG+Px+dANMWjlS7rMcGT5hGZScw7fV619ZglD2+kzqHfEX0e
# 6ovfzB1XGWTvWrW29Vgy0abLz3NjegsUdRMGIP84l6t6Rh35PU1gt/GgdvAEEHWA
# 7FNxRgo6lH4uMaLku6SQYZa+BmT5hDAYPUpzkciEtyq4UXG2YcPKqq9MeQ2FGbe1
# NnxtTUle87n4Q+MmxTGE+KkovD3r7vHgR/+prBakbNB29QkhHKj7+vSJ2k8aw5zy
# IHZ3NYewFbrQh919lY7ubdDp80mYQoc3oIgFCW9u9f0R6rT6XAvJoguoykxU+88u
# A6kdrLClRmiLsHmJSukGV/QUCLOKTMk25BSXdRYMNY50pe0oqQa4cp66v2SGACFm
# Yr5tK8SxyOj2S4pDPEXNer5myW8rKIKeyg/Jr60Vki47avFH0y66T+oiUDBCSb0c
# EHktNP4jvn19TCPr54zVI3VX3Hcfrwu6rcXtTJF1ScknDnUOEOb5aJzCBuSJxHUK
# 516fnaE/UFRX86iId5u1e78nUK+yYbhyB3ZkFMazG+x4fo5h3SLoYafklzTjZVj0
# h69tRqvPDN5+tsEhe1zsnnfGzp8arVsCtyNH3BigMB6xar4gfNGJGKqrJnYMT5w2
# gP2golsChw==
# SIG # End signature block
