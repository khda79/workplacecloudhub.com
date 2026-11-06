<#
.SYNOPSIS
    Worker-side remote operations: payload copy, PsExec execution, result correlation, evidence.

.DESCRIPTION
    Each endpoint attempt receives a CorrelationId. The endpoint writes
    <DataRoot>\State\Results\<CorrelationId>.json.txt when it starts (State=Running) and when it
    ends (State=Completed). The worker only trusts a result whose CorrelationId matches the
    attempt: older LastRun.json.txt files, retry-after-reboot results of previous attempts and
    evidence of other Runs can never be attributed to the current attempt.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

function Get-SemtComputerPathKey {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$ComputerName)

    $short = Get-SemtSafeName -Value ($ComputerName.Split('.')[0]) -Fallback 'Computer' -MaxLength 40
    $hash = (Get-SemtStringSha256 -Text $ComputerName.Trim().ToLowerInvariant()).Substring(0, 8).ToLowerInvariant()
    return ('{0}-{1}' -f $short,$hash)
}

function Protect-SemtRemotePathAcl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [switch]$Directory
    )

    $icacls = Join-Path -Path $env:SystemRoot -ChildPath 'System32\icacls.exe'
    if (-not (Test-Path -LiteralPath $icacls -PathType Leaf)) { return 'IcaclsNotFound' }
    $grants = if ($Directory) { @('*S-1-5-18:(OI)(CI)F','*S-1-5-32-544:(OI)(CI)F','*S-1-5-32-545:(OI)(CI)RX') } else { @('*S-1-5-18:F','*S-1-5-32-544:F','*S-1-5-32-545:RX') }
    $output = & $icacls $Path '/inheritance:r' '/grant:r' $grants 2>&1
    if ($LASTEXITCODE -ne 0) { throw ("ACL hardening failed. Path={0}; ExitCode={1}; Output={2}" -f $Path,$LASTEXITCODE,(($output | Out-String).Trim())) }
    return 'Protected'
}

function Copy-SemtRemotePayload {
    <#
    Copies the endpoint script through a staging name, verifies size, SHA256 and version, then
    renames it atomically. The expected SHA256 is supplied by the worker after its own
    verification against the distribution manifest or the launcher contract.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ConnectionTarget,
        [Parameter(Mandatory = $true)][string]$LocalScriptPath,
        [Parameter(Mandatory = $true)][string]$ExpectedSha256,
        [Parameter(Mandatory = $true)][string]$RemoteDataRoot,
        [Parameter(Mandatory = $true)][string]$RemoteProductRoot,
        [Parameter(Mandatory = $true)][string]$RemoteProgramRoot,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [string]$RemoteProgramFilesRoot = (Get-SemtTargetProgramFilesRoot)
    )

    # Data (results, logs, state) stays under ProgramData. The script that PsExec starts as SYSTEM
    # is stored under Program Files: the folder chain is created with a protected ACL (owner
    # Administrators) or verified, and any untrusted component stops the copy.
    $remoteProductShare = ConvertTo-SemtAdminSharePath -ComputerName $ConnectionTarget -LocalPath $RemoteProductRoot
    $remoteDataShare = ConvertTo-SemtAdminSharePath -ComputerName $ConnectionTarget -LocalPath $RemoteDataRoot
    $remoteProgramShare = ConvertTo-SemtAdminSharePath -ComputerName $ConnectionTarget -LocalPath $RemoteProgramRoot
    $remoteProgramFilesShare = ConvertTo-SemtAdminSharePath -ComputerName $ConnectionTarget -LocalPath $RemoteProgramFilesRoot
    New-SemtDirectory -Path $remoteProductShare
    [void](Protect-SemtRemotePathAcl -Path $remoteProductShare -Directory)
    New-SemtDirectory -Path $remoteDataShare
    [void](Initialize-SemtTrustedProgramFolder -Path $remoteProgramShare -BasePath $remoteProgramFilesShare)
    $scriptName = Split-Path -Path $LocalScriptPath -Leaf
    $finalPath = Join-Path -Path $remoteProgramShare -ChildPath $scriptName
    $stagingPath = Join-Path -Path $remoteProgramShare -ChildPath ('.{0}.{1}.tmp' -f $scriptName,([guid]::NewGuid().ToString('N')))
    $localItem = Get-Item -LiteralPath $LocalScriptPath -ErrorAction Stop
    $localVersion = Get-SemtScriptVersion -Path $LocalScriptPath
    try {
        Copy-Item -LiteralPath $LocalScriptPath -Destination $stagingPath -Force -ErrorAction Stop
        $stagingSha = Get-SemtFileSha256 -Path $stagingPath
        if ((Get-Item -LiteralPath $stagingPath).Length -ne $localItem.Length) { throw 'Staged endpoint script size mismatch.' }
        if ($stagingSha -ne $ExpectedSha256) { throw ("Staged endpoint script hash mismatch. Expected={0}; Staged={1}" -f $ExpectedSha256,$stagingSha) }
        Move-Item -LiteralPath $stagingPath -Destination $finalPath -Force -ErrorAction Stop
        $programTrust = Test-SemtTrustedProgramPath -Path $remoteProgramShare -Recurse
        if (-not $programTrust.Trusted) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $remoteProgramShare,$programTrust.State,$programTrust.Detail) }
        $remoteSha = Get-SemtFileSha256 -Path $finalPath
        if ($remoteSha -ne $ExpectedSha256) { throw ("Final endpoint script hash mismatch. Expected={0}; Remote={1}" -f $ExpectedSha256,$remoteSha) }
        $remoteVersion = Get-SemtScriptVersion -Path $finalPath
        if ($remoteVersion -ne $localVersion) { throw ("Final endpoint script version mismatch. Local={0}; Remote={1}" -f $localVersion,$remoteVersion) }
        Write-SemtLogLine -Path $LogPath -Message ("Endpoint script copied and verified: Version={0}; SHA256={1}; Path={2}" -f $remoteVersion,$remoteSha,$finalPath)
        return [pscustomobject]@{ LocalVersion = $localVersion; RemoteSha256 = $remoteSha; RemoteSharePath = $finalPath; RemoteDataShare = $remoteDataShare; RemoteScriptPath = (Join-Path -Path $RemoteProgramRoot -ChildPath $scriptName) }
    }
    finally {
        if (Test-Path -LiteralPath $stagingPath -PathType Leaf) { Remove-Item -LiteralPath $stagingPath -Force -ErrorAction SilentlyContinue }
    }
}

function Remove-SemtTemporaryPayload {
    [CmdletBinding()]
    param([Parameter(Mandatory=$true)][string]$Path,[Parameter(Mandatory=$true)][string]$ExpectedSha256)
    $parent = Split-Path -Path $Path -Parent
    $trust = Test-SemtTrustedProgramPath -Path $parent -Recurse
    if (-not $trust.Trusted) { throw ('Temporary payload cleanup refused: {0}' -f $trust.Detail) }
    $contents = @(Get-ChildItem -LiteralPath $parent -Force -ErrorAction Stop)
    if ($contents.Count -ne 1 -or $contents[0].FullName -ine $Path) { throw 'Temporary payload folder contains unexpected items; cleanup refused.' }
    if ((Get-SemtFileSha256 -Path $Path) -ne $ExpectedSha256) { throw 'Temporary endpoint hash changed; cleanup refused.' }
    Remove-Item -LiteralPath $Path -Force -ErrorAction Stop
    Remove-Item -LiteralPath $parent -Force -ErrorAction Stop
}

function Invoke-SemtSystemTaskEndpoint {
    <# ScheduledTask uses remote COM; WinRM uses the same scheduler locally on the target.
       Never stop or delete a task after a lost connection or an unconfirmed start.
       Transient RPC failures while following or cleaning the task (call failed, server
       unavailable, timeouts, disconnected object, network path lost) are retried with back-off and a new connection.
       The task is reported unconfirmed only when they last RpcFailureWindowMinutes in a row
       (0 = the first failure is final). Registration and start are never retried.
       StartTimeoutMinutes (0 = off): a task not seen running within this time (for example
       Queued, State 2, with no process) is removed by Remove-SemtUnstartedSystemTask only when
       nothing ran; the result then has NotStarted. If it started meanwhile, it is followed as usual.
       While the task has not started, its state is logged every 5 minutes. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][ValidateSet('ScheduledTask','WinRM')][string]$Mode,
        [Parameter(Mandatory=$true)][string]$ConnectionTarget,
        [Parameter(Mandatory=$true)][ValidatePattern('^[a-fA-F0-9]{32}$')][string]$CorrelationId,
        [Parameter(Mandatory=$true)][string]$RemoteScriptPath,
        [Parameter(Mandatory=$true)][string]$RemoteSharePath,
        [Parameter(Mandatory=$true)][string]$ExpectedSha256,
        [Parameter(Mandatory=$true)][AllowEmptyCollection()][string[]]$Arguments,
        [Parameter(Mandatory=$true)][string]$LogPath,
        [int]$TimeoutMinutes = 120,
        [AllowNull()][string]$LeasePath,
        [ValidateRange(0,1440)][int]$RpcFailureWindowMinutes = 30,
        [AllowNull()][string]$RemoteDataShare,
        [ValidateRange(0,1440)][int]$StartTimeoutMinutes = 0
    )
    $taskName = 'WCH-SEMT-' + $CorrelationId
    $actionArguments = (@('-NoProfile','-ExecutionPolicy','RemoteSigned','-File',$RemoteScriptPath) + $Arguments | ForEach-Object { ConvertTo-SemtWindowsCommandLineArgument -Value ([string]$_) }) -join ' '
    $request = @{
        Target = $ConnectionTarget; Name = $taskName; Arguments = $actionArguments
        Script = $RemoteScriptPath; ShareScript = $RemoteSharePath; Sha256 = $ExpectedSha256
        Mode = $Mode; RpcFailureWindowMinutes = $RpcFailureWindowMinutes
        StartTimeoutMinutes = $StartTimeoutMinutes; FollowOnly = $false
    }
    $operation = {
        param($r)
        $ErrorActionPreference = 'Stop'
        # RPC failures and network loss (bad network path, name deleted, network name not found).
        $transientCodes = @('800706BA','800706BB','800706BE','800706BF','80010108','800705B4','80070079','8007274C','80070035','80070033','80070040','80070043','800704CF')
        $windowMinutes = 30
        if ($r.ContainsKey('RpcFailureWindowMinutes') -and $null -ne $r.RpcFailureWindowMinutes) { $windowMinutes = [int]$r.RpcFailureWindowMinutes }
        $rpc = @{ Folder = $null; Failures = 0; Total = 0; FirstUtc = $null }
        $connect = {
            $connected = New-Object -ComObject 'Schedule.Service'
            if ($r.Mode -eq 'WinRM') { $connected.Connect() } else { $connected.Connect($r.Target) }
            $rpc.Folder = $connected.GetFolder('\')
        }
        $getCode = {
            param($exception)
            for ($current = $exception; $null -ne $current; $current = $current.InnerException) {
                # An empty task reply after a dropped RPC connection: retried like a transient failure.
                if ([string]$current.Message -like 'SEMT_TASK_REPLY_INCOMPLETE*') { return 'INCOMPLETE' }
                $code = '{0:X8}' -f $current.HResult
                if ($transientCodes -contains $code) { return $code }
                if ([string]$current.Message -match '(?i)0x(?<Code>[0-9A-F]{8})' -and $transientCodes -contains $Matches.Code.ToUpperInvariant()) { return $Matches.Code.ToUpperInvariant() }
            }
            return ''
        }
        $invokeRpc = {
            param([string]$Step,[scriptblock]$Call)
            while ($true) {
                try {
                    if ($null -eq $rpc.Folder) { & $connect }
                    $value = & $Call
                    if ($rpc.Failures -gt 0) {
                        Write-Warning ('SEMT_RPC_RECOVERED: Step={0}; ConsecutiveFailures={1}; TotalFailures={2}; FailingForMinutes={3:N1}' -f $Step,$rpc.Failures,$rpc.Total,((Get-Date).ToUniversalTime() - $rpc.FirstUtc).TotalMinutes)
                        $rpc.Failures = 0
                        $rpc.FirstUtc = $null
                    }
                    return $value
                }
                catch {
                    $code = & $getCode $_.Exception
                    if (-not $code) { throw }
                    $now = (Get-Date).ToUniversalTime()
                    if ($rpc.Failures -eq 0) { $rpc.FirstUtc = $now }
                    $rpc.Failures = $rpc.Failures + 1
                    $rpc.Total = $rpc.Total + 1
                    $failingMinutes = ($now - $rpc.FirstUtc).TotalMinutes
                    if ($failingMinutes -ge $windowMinutes) {
                        throw ('SEMT_RPC_UNAVAILABLE: Step={0}; Code=0x{1}; ConsecutiveFailures={2}; TotalFailures={3}; FailingForMinutes={4:N1}; WindowMinutes={5}; Error={6}' -f $Step,$code,$rpc.Failures,$rpc.Total,$failingMinutes,$windowMinutes,$_.Exception.Message)
                    }
                    $delay = [int][math]::Min(60, 5 * [math]::Pow(2, [math]::Min(4, $rpc.Failures - 1)))
                    Write-Warning ('SEMT_RPC_TRANSIENT: Step={0}; Code=0x{1}; ConsecutiveFailures={2}; TotalFailures={3}; FailingForMinutes={4:N1}; WindowMinutes={5}; RetryInSeconds={6}; Error={7}' -f $Step,$code,$rpc.Failures,$rpc.Total,$failingMinutes,$windowMinutes,$delay,$_.Exception.Message)
                    Start-Sleep -Seconds $delay
                    $rpc.Folder = $null
                }
            }
        }
        $startLimit = 0
        if ($r.ContainsKey('StartTimeoutMinutes') -and $null -ne $r.StartTimeoutMinutes) { $startLimit = [int]$r.StartTimeoutMinutes }
        $followOnly = ($r.ContainsKey('FollowOnly') -and [bool]$r.FollowOnly)
        $service = New-Object -ComObject 'Schedule.Service'
        if ($r.Mode -eq 'WinRM') { $service.Connect() } else { $service.Connect($r.Target) }
        $folder = $service.GetFolder('\')
        $rpc.Folder = $folder
        # FollowOnly: the task was registered and started by an earlier job of this attempt.
        if (-not $followOnly) {
        $definition = $service.NewTask(0)
        $definition.RegistrationInfo.Description = 'SEMT temporary SYSTEM execution ' + $r.Name
        $definition.Principal.UserId = 'SYSTEM'
        $definition.Principal.LogonType = 5
        $definition.Principal.RunLevel = 1
        $definition.Settings.AllowDemandStart = $true
        $definition.Settings.DisallowStartIfOnBatteries = $false
        $definition.Settings.StopIfGoingOnBatteries = $false
        $definition.Settings.ExecutionTimeLimit = 'PT0S'
        $action = $definition.Actions.Create(0)
        $action.Path = '%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe'
        $action.Arguments = $r.Arguments
        # TASK_CREATE (2): a collision fails instead of replacing an existing task.
        $task = $folder.RegisterTaskDefinition($r.Name,$definition,2,'SYSTEM',$null,5,$null)
        [void]$task.Run($null)
        }
        $waitStarted = Get-Date
        $nextStateLog = 5
        $seenRunning = $false
        do {
            Start-Sleep -Seconds 3
            $poll = & $invokeRpc 'GetTaskState' {
                $polled = $rpc.Folder.GetTask($r.Name)
                $polledState = if ($null -ne $polled) { $polled.State } else { $null }
                $polledLastRun = if ($null -ne $polled) { $polled.LastRunTime } else { $null }
                if ($null -eq $polledState -or $null -eq $polledLastRun) { throw ('SEMT_TASK_REPLY_INCOMPLETE: State={0}; LastRunTime={1}' -f $polledState,$polledLastRun) }
                [pscustomobject]@{ State = [int]$polledState; LastRun = [datetime]$polledLastRun }
            }
            # TASK_STATE_RUNNING = 4. Until the task is seen running, its state is reported every
            # 5 minutes and, after the start limit, the job hands the decision back to the worker.
            if ($poll.State -eq 4) { $seenRunning = $true }
            $waited = ((Get-Date) - $waitStarted).TotalMinutes
            if (-not $seenRunning -and $waited -ge $nextStateLog) {
                Write-Warning ('SEMT_TASK_STATE: State={0}; LastRunTime={1:o}; WaitingMinutes={2:N0}' -f $poll.State,$poll.LastRun,$waited)
                $nextStateLog = $nextStateLog + 5
            }
            if ($startLimit -gt 0 -and -not $seenRunning -and $poll.State -ne 4 -and $waited -ge $startLimit -and (($poll.State -ne 1 -and $poll.State -ne 3) -or $poll.LastRun.Year -lt 2000)) {
                return [pscustomobject]@{ NotStarted = $true; State = $poll.State; LastRunTime = $poll.LastRun.ToString('o'); WaitedMinutes = [math]::Round($waited,1); Cleaned = $false; ExitCode = ''; RpcTransientFailures = $rpc.Total }
            }
        } while (($poll.State -ne 1 -and $poll.State -ne 3) -or $poll.LastRun.Year -lt 2000)
        # Re-read the definition and refuse cleanup if the identity has changed.
        $actual = & $invokeRpc 'ReadTaskDefinition' {
            $read = $rpc.Folder.GetTask($r.Name)
            $readDefinition = $read.Definition
            $count = [int]$readDefinition.Actions.Count
            $first = if ($count -ge 1) { $readDefinition.Actions.Item(1) } else { $null }
            [pscustomobject]@{
                ExitCode = [string]$read.LastTaskResult; ActionCount = $count
                ActionPath = $(if ($first) { [string]$first.Path } else { '' }); ActionArguments = $(if ($first) { [string]$first.Arguments } else { '' })
                LogonType = [int]$readDefinition.Principal.LogonType; UserId = [string]$readDefinition.Principal.UserId
            }
        }
        $exitCode = $actual.ExitCode
        $principalSid = ''
        try {
            $principal = $actual.UserId
            if ($principal -match '^S-\d-') {
                $principalSid = ([System.Security.Principal.SecurityIdentifier]::new($principal)).Value
            } else {
                $principalSid = ([System.Security.Principal.NTAccount]::new($principal)).Translate([System.Security.Principal.SecurityIdentifier]).Value
            }
        } catch { $principalSid = '' }
        if ($actual.ActionCount -ne 1 -or $actual.ActionPath -ine '%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe' -or $actual.ActionArguments -cne $r.Arguments -or $actual.LogonType -ne 5 -or $principalSid -ne 'S-1-5-18') {
            throw 'Task identity changed; cleanup refused.'
        }
        $scriptPath = if ($r.Mode -eq 'WinRM') { $r.Script } else { $r.ShareScript }
        $parent = Split-Path -Path $scriptPath -Parent
        $parentItem = Get-Item -LiteralPath $parent -Force -ErrorAction Stop
        $scriptItem = Get-Item -LiteralPath $scriptPath -Force -ErrorAction Stop
        if (($parentItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -or ($scriptItem.Attributes -band [IO.FileAttributes]::ReparsePoint)) { throw 'Reparse point found; cleanup refused.' }
        if ((Get-FileHash -LiteralPath $scriptPath -Algorithm SHA256 -ErrorAction Stop).Hash -ne $r.Sha256) { throw 'Endpoint hash changed; cleanup refused.' }
        $contents = @(Get-ChildItem -LiteralPath $parent -Force -ErrorAction Stop)
        if ($contents.Count -ne 1 -or $contents[0].FullName -ine $scriptItem.FullName) { throw 'Temporary folder contains unexpected items; cleanup refused.' }
        [void](& $invokeRpc 'DeleteTask' {
            if ($rpc.Failures -gt 0) {
                # A delete retried after a lost reply may already have been applied.
                try { [void]$rpc.Folder.GetTask($r.Name) }
                catch {
                    for ($missing = $_.Exception; $null -ne $missing; $missing = $missing.InnerException) { if (('{0:X8}' -f $missing.HResult) -eq '80070002') { return 'AlreadyDeleted' } }
                    throw
                }
            }
            $rpc.Folder.DeleteTask($r.Name,0)
            'Deleted'
        })
        Remove-Item -LiteralPath $scriptPath -Force -ErrorAction Stop
        Remove-Item -LiteralPath $parent -Force -ErrorAction Stop
        [pscustomobject]@{ ExitCode = $exitCode; Cleaned = $true; RpcTransientFailures = $rpc.Total }
    }
    $startJob = {
        param($jobRequest)
        Start-Job -ArgumentList $jobRequest,$operation.ToString() -ScriptBlock {
            param($r,$source)
            $block = [scriptblock]::Create($source)
            if ($r.Mode -eq 'WinRM') { Invoke-Command -ComputerName $r.Target -ScriptBlock $block -ArgumentList $r -ErrorAction Stop }
            else { & $block $r }
        }
    }
    $job = & $startJob $request
    $started = Get-Date
    $timedOut = $false
    # The job reports each transient RPC failure and each recovery as a warning record; they are
    # copied into the worker log as they arrive, and the attempt result is re-read on recovery.
    $relay = @{ Seen = 0; Transient = 0 }
    $relayWarnings = {
        $records = @()
        if ($null -ne $job -and $job.PSObject.Properties['ChildJobs']) {
            foreach ($child in @($job.ChildJobs)) { if ($null -ne $child -and $child.PSObject.Properties['Warning'] -and $null -ne $child.Warning) { $records += @($child.Warning) } }
        }
        for ($index = $relay.Seen; $index -lt $records.Count; $index++) {
            $message = [string]$records[$index].Message
            if ($message -like 'SEMT_RPC_TRANSIENT*') {
                $relay.Transient = $relay.Transient + 1
                Write-SemtLogLine -Path $LogPath -Level WARN -Message ("{0} task RPC failure #{1} (retrying, task {2} kept): {3}" -f $Mode,$relay.Transient,$taskName,$message)
            }
            elseif ($message -like 'SEMT_RPC_RECOVERED*') {
                Write-SemtLogLine -Path $LogPath -Message ("{0} task connection recovered after {1} RPC failure(s): {2}" -f $Mode,$relay.Transient,$message)
                if (-not [string]::IsNullOrWhiteSpace($RemoteDataShare)) {
                    $recovered = Read-SemtRemoteAttemptResult -RemoteDataShare $RemoteDataShare -CorrelationId $CorrelationId
                    Write-SemtLogLine -Path $LogPath -Message ("Attempt result after RPC recovery: {0}" -f $recovered.Detail)
                }
            }
            elseif ($message -like 'SEMT_TASK_STATE*') { Write-SemtLogLine -Path $LogPath -Level WARN -Message ("{0} task {1} not started yet: {2}" -f $Mode,$taskName,$message) }
            else { Write-SemtLogLine -Path $LogPath -Level WARN -Message ("{0} task warning: {1}" -f $Mode,$message) }
        }
        $relay.Seen = $records.Count
    }
    try {
        while ($true) {
            while (-not (Wait-Job -Job $job -Timeout 5)) {
                [void](Update-SemtWorkerLease -LeasePath $LeasePath)
                & $relayWarnings
                if ($TimeoutMinutes -gt 0 -and ((Get-Date) - $started).TotalMinutes -ge $TimeoutMinutes) { $timedOut = $true; break }
            }
            & $relayWarnings
            if ($timedOut) {
                Write-SemtLogLine -Path $LogPath -Level WARN -Message ("{0} timed out; remote execution and cleanup are unconfirmed. Task={1}; RpcFailures={2}" -f $Mode,$taskName,$relay.Transient)
                return [pscustomobject]@{ Completed=$false; TimedOut=$true; ExitCode=''; Detail='Remote task may still be running; verify before retrying.'; RpcTransientFailures=$relay.Transient; RpcUnavailable=$false }
            }
            $output = @(Receive-Job -Job $job -ErrorAction Stop -WarningAction SilentlyContinue)
            if ($job.State -eq 'Completed' -and $output.Count -eq 1 -and $output[0].PSObject.Properties['NotStarted'] -and [bool]$output[0].NotStarted) {
                Write-SemtLogLine -Path $LogPath -Level WARN -Message ("{0} task {1} not started after {2} minute(s) (State={3}; LastRunTime={4}; limit RemoteTaskStartTimeoutMinutes={5}). Checking whether anything ran before removing it." -f $Mode,$taskName,$output[0].WaitedMinutes,$output[0].State,$output[0].LastRunTime,$StartTimeoutMinutes)
                $cleanup = Remove-SemtUnstartedSystemTask -Mode $Mode -ConnectionTarget $ConnectionTarget -CorrelationId $CorrelationId -RemoteDataShare $RemoteDataShare -RemoteSharePath $RemoteSharePath -ExpectedSha256 $ExpectedSha256 -RemoteScriptPath $RemoteScriptPath
                Write-SemtLogLine -Path $LogPath -Level $(if ($cleanup.Outcome -eq 'Deleted') { 'INFO' } else { 'WARN' }) -Message ("Task start cleanup {0}: {1}" -f $cleanup.Outcome,$cleanup.Detail)
                if ($cleanup.Outcome -ne 'Started') {
                    $detail = ('Task {0} did not start within {1} minute(s) (State={2}; LastRunTime={3}). Task cleanup {4}: {5}' -f $taskName,$StartTimeoutMinutes,$output[0].State,$output[0].LastRunTime,$cleanup.Outcome,$cleanup.Detail)
                    return [pscustomobject]@{ Completed=$false; TimedOut=$false; NotStarted=$true; StartCleanup=$cleanup.Outcome; ExitCode=''; Detail=$detail; RpcTransientFailures=$relay.Transient; RpcUnavailable=$false }
                }
                # It started in the meantime: follow it to the end like any started task.
                Write-SemtLogLine -Path $LogPath -Message ("{0} task {1} started meanwhile; following it." -f $Mode,$taskName)
                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
                $follow = $request.Clone()
                $follow.FollowOnly = $true
                $follow.StartTimeoutMinutes = 0
                $relay.Seen = 0
                $job = & $startJob $follow
                continue
            }
            if ($job.State -ne 'Completed' -or $output.Count -ne 1 -or -not $output[0].Cleaned) { throw 'Task execution or cleanup was not confirmed.' }
            Write-SemtLogLine -Path $LogPath -Message ("{0} task finished and temporary payload was removed. Task={1}; ExitCode={2}; RpcFailures={3}" -f $Mode,$taskName,$output[0].ExitCode,$relay.Transient)
            return [pscustomobject]@{ Completed=$true; TimedOut=$false; ExitCode=[string]$output[0].ExitCode; Detail='Task and temporary endpoint script removed.'; RpcTransientFailures=$relay.Transient; RpcUnavailable=$false }
        }
    }
    catch {
        $detail = $_.Exception.Message
        try { & $relayWarnings } catch { }
        $unavailable = $detail -like '*SEMT_RPC_UNAVAILABLE*'
        if ($unavailable) { $detail = ('RPC failures lasted {0} minute(s) or more in a row (window RemoteTaskRpcFailureWindowMinutes); task {1} was kept. {2}' -f $RpcFailureWindowMinutes,$taskName,$detail) }
        Write-SemtLogLine -Path $LogPath -Level WARN -Message ("{0} task state or cleanup unconfirmed: {1}" -f $Mode,$detail)
        return [pscustomobject]@{ Completed=$false; TimedOut=$false; ExitCode=''; Detail=$detail; RpcTransientFailures=$relay.Transient; RpcUnavailable=$unavailable }
    }
    finally { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue }
}

function Remove-SemtUnstartedSystemTask {
    <# Deletes the temporary SEMT task of an attempt only when nothing ran: the task is not running
       (State 4), it has not finished a run (Ready with a LastRunTime), and the target holds no
       SEMT trace of the attempt (attempt result, target operation lease with this correlation ID,
       HybridJoin endpoint instance started from this attempt's script). A Queued task (State 2)
       may carry a LastRunTime although no process was created: LastRunTime alone never counts.
       The task is disabled, read again (a new launch or a running state keeps it), deleted, and
       the traces are read once more; the payload is removed only when none appeared. A process
       created after the last read keeps running (deleting a task never stops its process).
       Never throws. Outcome: Deleted, Started, Absent or Failed. #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)][ValidateSet('ScheduledTask','WinRM')][string]$Mode,
        [Parameter(Mandatory=$true)][string]$ConnectionTarget,
        [Parameter(Mandatory=$true)][ValidatePattern('^[a-fA-F0-9]{32}$')][string]$CorrelationId,
        [AllowNull()][string]$RemoteDataShare,
        [AllowNull()][string]$RemoteSharePath,
        [AllowNull()][string]$ExpectedSha256,
        [AllowNull()][string]$RemoteScriptPath,
        [ValidateRange(10,600)][int]$TimeoutSeconds = 120
    )
    $taskName = 'WCH-SEMT-' + $CorrelationId
    $outcome = { param([string]$Value,[string]$Detail) [pscustomobject]@{ Outcome = $Value; TaskName = $taskName; Detail = $Detail } }
    # Returns a description of the first SEMT trace of this attempt on the target, '' when none.
    # An unreadable lease or instance file counts as a trace: the task is then kept.
    $readTrace = {
        if ([string]::IsNullOrWhiteSpace($RemoteDataShare)) { return '' }
        $attempt = Read-SemtRemoteAttemptResult -RemoteDataShare $RemoteDataShare -CorrelationId $CorrelationId
        if ($attempt.Exists -and $attempt.State -ne 'Foreign') { return ('the attempt wrote a result ({0})' -f $attempt.Detail) }
        $leasePath = Join-Path -Path (Split-Path -Path $RemoteDataShare -Parent) -ChildPath 'State\TargetOperation.json.txt'
        if (Test-Path -LiteralPath $leasePath -PathType Leaf) {
            try { $lease = Read-SemtJsonFile -Path $leasePath -Required } catch { return ('the target operation lease is unreadable: {0}' -f $_.Exception.Message) }
            if ([string](Get-SemtPropertyValue -InputObject $lease -Name 'CorrelationId' -Default '') -eq $CorrelationId) { return 'the attempt holds the target operation lease' }
        }
        if (-not [string]::IsNullOrWhiteSpace($RemoteScriptPath)) {
            $instancePath = Join-Path -Path $RemoteDataShare -ChildPath 'State\EndpointInstance.json.txt'
            if (Test-Path -LiteralPath $instancePath -PathType Leaf) {
                try { $instance = Read-SemtJsonFile -Path $instancePath -Required } catch { return ('the endpoint instance file is unreadable: {0}' -f $_.Exception.Message) }
                if ([string](Get-SemtPropertyValue -InputObject $instance -Name 'ScriptPath' -Default '') -ieq $RemoteScriptPath) { return 'an endpoint instance runs from this attempt script' }
            }
        }
        return ''
    }
    try {
        $trace = & $readTrace
        if ($trace) { return (& $outcome 'Started' ('Before cleanup, {0}; task {1} kept.' -f $trace,$taskName)) }
        $request = @{ Target = $ConnectionTarget; Name = $taskName; Mode = $Mode; Description = 'SEMT temporary SYSTEM execution ' + $taskName }
        $operation = {
            param($r)
            $ErrorActionPreference = 'Stop'
            $service = New-Object -ComObject 'Schedule.Service'
            if ($r.Mode -eq 'WinRM') { $service.Connect() } else { $service.Connect($r.Target) }
            $folder = $service.GetFolder('\')
            $read = {
                try { $task = $folder.GetTask($r.Name) }
                catch {
                    for ($e = $_.Exception; $null -ne $e; $e = $e.InnerException) { if (('{0:X8}' -f $e.HResult) -eq '80070002') { return $null } }
                    throw
                }
                $state = $task.State
                $lastRun = $task.LastRunTime
                if ($null -eq $state -or $null -eq $lastRun) { throw ('SEMT_TASK_REPLY_INCOMPLETE: State={0}; LastRunTime={1}' -f $state,$lastRun) }
                [pscustomobject]@{ Task = $task; State = [int]$state; LastRun = [datetime]$lastRun; Description = [string]$task.Definition.RegistrationInfo.Description }
            }
            $first = & $read
            if ($null -eq $first) { return [pscustomobject]@{ Outcome = 'Absent'; Detail = 'Task not found on the target.' } }
            if ($first.Description -cne $r.Description) { return [pscustomobject]@{ Outcome = 'Failed'; Detail = 'Task identity differs; deletion refused.' } }
            # TASK_STATE_RUNNING = 4. A task that is not Queued (2) and has a LastRunTime after 2000
            # finished a run. A Queued task may carry a LastRunTime without any process.
            if ($first.State -eq 4) { return [pscustomobject]@{ Outcome = 'Started'; Detail = ('Task running (State=4; LastRunTime={0:o}).' -f $first.LastRun) } }
            if ($first.State -ne 2 -and $first.LastRun.Year -ge 2000) { return [pscustomobject]@{ Outcome = 'Started'; Detail = ('Task ran (State={0}; LastRunTime={1:o}).' -f $first.State,$first.LastRun) } }
            $first.Task.Enabled = $false
            $second = & $read
            if ($null -eq $second) { return [pscustomobject]@{ Outcome = 'Absent'; Detail = 'Task disappeared while being disabled.' } }
            if ($second.State -eq 4 -or $second.LastRun -ne $first.LastRun) { return [pscustomobject]@{ Outcome = 'Started'; Detail = ('Task started while being disabled (State={0}; LastRunTime={1:o}); it keeps running.' -f $second.State,$second.LastRun) } }
            $folder.DeleteTask($r.Name,0)
            $before = if ($first.State -eq 2) { ('Queued with LastRunTime={0:o} and no process' -f $first.LastRun) } else { ('State={0}, never run' -f $first.State) }
            [pscustomobject]@{ Outcome = 'Deleted'; Detail = ('Task never started ({0}); disabled then deleted.' -f $before) }
        }
        $job = Start-Job -ArgumentList $request,$operation.ToString() -ScriptBlock {
            param($r,$source)
            $block = [scriptblock]::Create($source)
            if ($r.Mode -eq 'WinRM') { Invoke-Command -ComputerName $r.Target -ScriptBlock $block -ArgumentList $r -ErrorAction Stop }
            else { & $block $r }
        }
        try {
            if (-not (Wait-Job -Job $job -Timeout $TimeoutSeconds)) { return (& $outcome 'Failed' ('No reply from the target Task Scheduler within {0} s; task {1} kept.' -f $TimeoutSeconds,$taskName)) }
            $reply = @(Receive-Job -Job $job -ErrorAction Stop)
            if ($reply.Count -ne 1) { return (& $outcome 'Failed' ('Unexpected Task Scheduler reply; task {0} kept.' -f $taskName)) }
        }
        finally { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue }
        $result = & $outcome ([string]$reply[0].Outcome) ('{0} Task={1}' -f [string]$reply[0].Detail,$taskName)
        if ($result.Outcome -eq 'Deleted') {
            $late = & $readTrace
            if ($late) {
                # The endpoint started after the last read: it keeps running and reports normally.
                $result.Outcome = 'Started'
                $result.Detail = ('{0}; started at the last moment ({1}); temporary payload kept.' -f $result.Detail,$late)
            }
            elseif (-not [string]::IsNullOrWhiteSpace($RemoteSharePath) -and -not [string]::IsNullOrWhiteSpace($ExpectedSha256)) {
                try { Remove-SemtTemporaryPayload -Path $RemoteSharePath -ExpectedSha256 $ExpectedSha256; $result.Detail = $result.Detail + '; temporary payload removed.' }
                catch { $result.Detail = ('{0}; temporary payload kept: {1}' -f $result.Detail,$_.Exception.Message) }
            }
        }
        return $result
    }
    catch { return (& $outcome 'Failed' ('Cleanup of task {0} failed; task kept. {1}' -f $taskName,$_.Exception.Message)) }
}

function Invoke-SemtPsExecEndpoint {
    <#
    Starts PsExec, records the PsExec PID in the worker reservation, refreshes the reservation
    heartbeat every 60 seconds and enforces the timeout. A timeout stops only the local PsExec
    process; the remote endpoint state stays unconfirmed until its result is read.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$PsExecPath,
        [Parameter(Mandatory = $true)][string]$ConnectionTarget,
        [Parameter(Mandatory = $true)][string]$RemoteScriptPath,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$Arguments,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [int]$TimeoutMinutes = 0,
        [bool]$UseElevatedToken = $false,
        [AllowNull()][string]$LeasePath
    )

    $stdoutPath = '{0}.stdout.txt' -f $LogPath
    $stderrPath = '{0}.stderr.txt' -f $LogPath
    $psexecArguments = @(('\\{0}' -f $ConnectionTarget),'-accepteula','-nobanner','-s')
    if ($UseElevatedToken) { $psexecArguments += '-h' }
    # RemoteSigned, not Bypass: the endpoint script is a local signed file on the target, so it runs
    # the same way, without the Bypass switch that endpoint protection products watch closely.
    $psexecArguments += @('powershell.exe','-NoProfile','-ExecutionPolicy','RemoteSigned','-File',(ConvertTo-SemtPsExecArgument -Value $RemoteScriptPath))
    $psexecArguments += @($Arguments | ForEach-Object { ConvertTo-SemtPsExecArgument -Value $_ })
    Write-SemtLogLine -Path $LogPath -Message ("PsExec command: {0} {1}" -f $PsExecPath,($psexecArguments -join ' '))
    $process = Start-Process -FilePath $PsExecPath -ArgumentList $psexecArguments -PassThru -NoNewWindow -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
    # Windows PowerShell 5.1: without an open handle, ExitCode stays empty after WaitForExit(timeout).
    $null = $process.Handle
    $started = Get-Date
    [void](Update-SemtWorkerLease -LeasePath $LeasePath -Properties @{ PsExecProcessId = [int]$process.Id; PsExecProcessStartUtc = (Get-SemtProcessStartUtc -ProcessId $process.Id) })
    $timedOut = $false
    $lastHeartbeat = $started
    while (-not $process.WaitForExit(1000)) {
        $now = Get-Date
        if ($TimeoutMinutes -gt 0 -and ($now - $started).TotalMinutes -ge $TimeoutMinutes) {
            $timedOut = $true
            Write-SemtLogLine -Path $LogPath -Level ERROR -Message ("PsExec timed out after {0} minute(s). Stopping the local PsExec process; the remote endpoint state is unconfirmed." -f $TimeoutMinutes)
            try { & taskkill.exe /PID $process.Id /T /F 2>&1 | Out-Null } catch { }
            try { [void]$process.WaitForExit(5000) } catch { }
            break
        }
        if (($now - $lastHeartbeat).TotalSeconds -ge 60) {
            [void](Update-SemtWorkerLease -LeasePath $LeasePath)
            Write-SemtLogLine -Path $LogPath -Message ("PsExec still running for {0:N1} minute(s)." -f ($now - $started).TotalMinutes)
            $lastHeartbeat = $now
        }
    }
    try { $process.Refresh() } catch { }
    $exitCode = ''
    if (-not $timedOut) { try { if ($process.HasExited -and $null -ne $process.ExitCode) { $exitCode = [string]$process.ExitCode } } catch { $exitCode = '' } }
    $stdout = @()
    $stderr = @()
    if (Test-Path -LiteralPath $stdoutPath -PathType Leaf) { $stdout = @(Get-Content -LiteralPath $stdoutPath -ErrorAction SilentlyContinue) }
    if (Test-Path -LiteralPath $stderrPath -PathType Leaf) { $stderr = @(Get-Content -LiteralPath $stderrPath -ErrorAction SilentlyContinue) }
    if ([string]::IsNullOrWhiteSpace($exitCode)) {
        $nativeLine = $stderr | Where-Object { $_ -match 'with error code\s+-?\d+' } | Select-Object -Last 1
        if ($nativeLine -and $nativeLine -match 'with error code\s+(?<Code>-?\d+)') {
            $exitCode = $Matches.Code
            Write-SemtLogLine -Path $LogPath -Message ("PsExec exit code recovered from native STDERR: {0}" -f $exitCode)
        }
    }
    Add-Content -LiteralPath $LogPath -Value '----- PsExec STDOUT -----' -Encoding UTF8
    if ($stdout.Count -gt 0) { Add-Content -LiteralPath $LogPath -Value $stdout -Encoding UTF8 }
    Add-Content -LiteralPath $LogPath -Value '----- PsExec STDERR -----' -Encoding UTF8
    if ($stderr.Count -gt 0) { Add-Content -LiteralPath $LogPath -Value $stderr -Encoding UTF8 }
    return [pscustomobject]@{
        ExitCode = $exitCode
        TimedOut = $timedOut
        Stdout = $stdout
        Stderr = $stderr
        StdoutPath = $stdoutPath
        StderrPath = $stderrPath
        FinalStatus = ConvertFrom-SemtFinalStatusLine -Lines $stdout
        CommunicationFailureDetail = Get-SemtPsExecCommunicationFailureDetail -Lines $stderr
        ElapsedMinutes = [math]::Round(((Get-Date) - $started).TotalMinutes, 1)
    }
}

#region Automatic prerequisites

function Get-SemtPrerequisiteTaskExpectations {
    <#
    Read-only summary of the scheduled tasks that the Prerequisites endpoint creates for a campaign
    (see Get-PrerequisiteTaskDefinitions in the Prerequisites endpoint; Test-SemtPrerequisites keeps
    both lists consistent). Each action is compared as 'Path|Arguments' against ActionPatterns, in
    order. Obsolete tasks must be absent: the Prerequisites endpoint removes them.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][ValidateSet('HybridJoin','Windows11Upgrade')][string]$Campaign)

    $hybridJoinPath = '\WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin\'
    $expectations = New-Object System.Collections.Generic.List[object]
    if ($Campaign -eq 'HybridJoin') {
        $expectations.Add([pscustomobject]@{ Kind = 'Required'; TaskPath = $hybridJoinPath; TaskName = 'SEMT-HJ-UserDsregStatus'; MinimumTriggers = 0; ActionPatterns = @('(?i)^(.*\\)?cmd\.exe\|.*\\dsregcmd\.exe /status > .*SEMT-HJ-UserDsregStatus_%USERNAME%\.txt 2>&1$') })
        $expectations.Add([pscustomobject]@{ Kind = 'Required'; TaskPath = $hybridJoinPath; TaskName = 'SEMT-HJ-UserRefreshPrt'; MinimumTriggers = 0; ActionPatterns = @('(?i)^(.*\\)?dsregcmd\.exe\|/refreshprt$') })
        $expectations.Add([pscustomobject]@{ Kind = 'Required'; TaskPath = $hybridJoinPath; TaskName = 'SEMT-HJ-RunUserAutoEnrollAtLogon'; MinimumTriggers = 1; ActionPatterns = @('(?i)^(.*\\)?schtasks\.exe\|/Run /TN "\\WCH\\SmartM365\\SmartEndpointMigrationToolkit\\HybridJoin\\SEMT-HJ-UserDsregStatus"$','(?i)^(.*\\)?deviceenroller\.exe\|/c /AutoEnrollMDM$') })
        $expectations.Add([pscustomobject]@{ Kind = 'Obsolete'; TaskPath = $hybridJoinPath; TaskName = 'SEMT-HJ-UserMdmAutoEnroll'; MinimumTriggers = 0; ActionPatterns = @() })
    }
    $runner = [regex]::Escape(('\WCH\SmartM365\SmartEndpointMigrationToolkit\{0}\RetryAfterReboot\RetryAfterRebootRunner.ps1' -f $Campaign))
    $expectations.Add([pscustomobject]@{ Kind = 'Required'; TaskPath = '\'; TaskName = ('SmartM365-SEMT-{0}-RetryAfterReboot' -f $Campaign); MinimumTriggers = 1; ActionPatterns = @(('(?i)^(.*\\)?powershell\.exe\|-NoProfile -ExecutionPolicy RemoteSigned -File ".*{0}"$' -f $runner)) })
    return @($expectations.ToArray())
}

function Test-SemtPrerequisiteTaskReadState {
    <#
    Compares one task read from the target with its expectation. Returns Present, Missing,
    Outdated, Obsolete (an obsolete task still exists), Absent (an obsolete task is gone) or
    Unknown (the task could not be read).
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)]$Expectation,[AllowNull()]$Read)

    $readError = [string](Get-SemtPropertyValue -InputObject $Read -Name 'Error' -Default '')
    if ($null -eq $Read -or -not [string]::IsNullOrWhiteSpace($readError)) { return 'Unknown' }
    $exists = [bool](Get-SemtPropertyValue -InputObject $Read -Name 'Exists' -Default $false)
    if ($Expectation.Kind -eq 'Obsolete') { if ($exists) { return 'Obsolete' } else { return 'Absent' } }
    if (-not $exists) { return 'Missing' }
    $actions = @(Get-SemtPropertyValue -InputObject $Read -Name 'Actions' -Default @() | ForEach-Object { [string]$_ })
    $patterns = @($Expectation.ActionPatterns)
    if ($actions.Count -ne $patterns.Count) { return 'Outdated' }
    for ($index = 0; $index -lt $patterns.Count; $index++) {
        if ($actions[$index] -notmatch $patterns[$index]) { return 'Outdated' }
    }
    if ([int](Get-SemtPropertyValue -InputObject $Read -Name 'TriggerCount' -Default 0) -lt [int]$Expectation.MinimumTriggers) { return 'Outdated' }
    return 'Present'
}

function Get-SemtPrerequisiteTaskSummary {
    <#
    Decision from the per-task states: NeedsRun when a task is proven missing, outdated or
    obsolete; CheckFailed when nothing is proven missing but a task could not be read; Ready
    otherwise.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Tasks)

    $states = @($Tasks | ForEach-Object { [string]$_.State })
    $decision = 'Ready'
    if (@($states | Where-Object { $_ -in @('Missing','Outdated','Obsolete') }).Count -gt 0) { $decision = 'NeedsRun' }
    elseif (@($states | Where-Object { $_ -eq 'Unknown' }).Count -gt 0) { $decision = 'CheckFailed' }
    $detail = (@($Tasks | ForEach-Object { '{0}={1}' -f $_.TaskName,$_.State }) -join '; ')
    return [pscustomobject]@{ Decision = $decision; Detail = $detail; Tasks = @($Tasks) }
}

function Get-SemtRemotePrerequisiteTaskState {
    <#
    Read-only check, from the technician computer, of the scheduled tasks that a campaign needs on
    the target. ScheduledTask and PsExec modes read the remote Task Scheduler (RPC); WinRM reads it
    on the target. Nothing is created, changed, started or deleted.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ConnectionTarget,
        [Parameter(Mandatory = $true)][ValidateSet('HybridJoin','Windows11Upgrade')][string]$Campaign,
        [ValidateSet('ScheduledTask','WinRM','PsExec')][string]$Mode = 'ScheduledTask',
        [ValidateRange(10, 1800)][int]$TimeoutSeconds = 120
    )

    $expectations = @(Get-SemtPrerequisiteTaskExpectations -Campaign $Campaign)
    $request = @{
        Target = $ConnectionTarget
        Mode = $(if ($Mode -eq 'WinRM') { 'WinRM' } else { 'ScheduledTask' })
        Tasks = @($expectations | ForEach-Object { @{ TaskPath = [string]$_.TaskPath; TaskName = [string]$_.TaskName } })
    }
    $operation = {
        param($r)
        $ErrorActionPreference = 'Stop'
        $service = New-Object -ComObject 'Schedule.Service'
        if ($r.Mode -eq 'WinRM') { $service.Connect() } else { $service.Connect($r.Target) }
        foreach ($item in @($r.Tasks)) {
            $row = [ordered]@{ TaskPath = [string]$item.TaskPath; TaskName = [string]$item.TaskName; Exists = $false; Actions = @(); TriggerCount = 0; Error = '' }
            try {
                $folderPath = if ([string]$item.TaskPath -eq '\') { '\' } else { ([string]$item.TaskPath).TrimEnd('\') }
                $task = $service.GetFolder($folderPath).GetTask([string]$item.TaskName)
                $definition = $task.Definition
                $actions = @()
                for ($index = 1; $index -le [int]$definition.Actions.Count; $index++) {
                    $action = $definition.Actions.Item($index)
                    if ([int]$action.Type -eq 0) { $actions += ('{0}|{1}' -f ([string]$action.Path).Trim(),([string]$action.Arguments).Trim()) }
                    else { $actions += ('ActionType{0}|' -f [int]$action.Type) }
                }
                $row.Exists = $true
                $row.Actions = $actions
                $row.TriggerCount = [int]$definition.Triggers.Count
            }
            catch {
                $notFound = $false
                for ($current = $_.Exception; $null -ne $current; $current = $current.InnerException) {
                    if (('{0:X8}' -f $current.HResult) -in @('80070002','80070003')) { $notFound = $true }
                }
                if (-not $notFound) { $row.Error = [string]$_.Exception.Message }
            }
            [pscustomobject]$row
        }
    }
    $job = Start-Job -ArgumentList $request,$operation.ToString() -ScriptBlock {
        param($r,$source)
        $block = [scriptblock]::Create($source)
        if ($r.Mode -eq 'WinRM') { Invoke-Command -ComputerName $r.Target -ScriptBlock $block -ArgumentList $r -ErrorAction Stop }
        else { & $block $r }
    }
    $reads = @()
    $failure = ''
    try {
        if (-not (Wait-Job -Job $job -Timeout $TimeoutSeconds)) { $failure = ('Task Scheduler read timed out after {0} second(s).' -f $TimeoutSeconds) }
        else {
            try { $reads = @(Receive-Job -Job $job -ErrorAction Stop) }
            catch { $failure = ('Task Scheduler read failed: {0}' -f $_.Exception.Message) }
            if (-not $failure -and $job.State -ne 'Completed') { $failure = ('Task Scheduler read job ended in state {0}.' -f $job.State) }
        }
    }
    finally { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue }

    $tasks = foreach ($expectation in $expectations) {
        $read = $null
        if (-not $failure) { $read = @($reads | Where-Object { [string]$_.TaskPath -eq [string]$expectation.TaskPath -and [string]$_.TaskName -eq [string]$expectation.TaskName }) | Select-Object -First 1 }
        $state = Test-SemtPrerequisiteTaskReadState -Expectation $expectation -Read $read
        $errorText = if ($failure) { $failure } else { [string](Get-SemtPropertyValue -InputObject $read -Name 'Error' -Default '') }
        [pscustomobject]@{ TaskPath = $expectation.TaskPath; TaskName = $expectation.TaskName; Kind = $expectation.Kind; State = $state; Error = $errorText }
    }
    $summary = Get-SemtPrerequisiteTaskSummary -Tasks @($tasks)
    $detail = $summary.Detail
    if ($failure) { $detail = '{0} {1}' -f $failure,$detail }
    else {
        $firstError = @($tasks | Where-Object { $_.State -eq 'Unknown' -and $_.Error }) | Select-Object -First 1
        if ($firstError) { $detail = '{0} (read error on {1}: {2})' -f $detail,$firstError.TaskName,$firstError.Error }
    }
    return [pscustomobject]@{ Decision = $summary.Decision; Detail = $detail; Tasks = @($tasks) }
}

#endregion

function ConvertFrom-SemtFinalStatusLine {
    [CmdletBinding()]
    param([AllowEmptyCollection()][string[]]$Lines)

    $line = @($Lines) | Where-Object { $_ -match '^FINAL_STATUS=' } | Select-Object -Last 1
    if (-not $line) { return $null }
    if ($line -match '^FINAL_STATUS=(?<Status>[^;]*);\s*EXIT_CODE=(?<ExitCode>[^;]*);\s*NEXT_ACTION=(?<NextAction>[^;]*);\s*(?:CORRELATION_ID=(?<CorrelationId>[^;]*);\s*)?DETAIL=(?<Detail>.*)$') {
        # The CorrelationId group is absent from lines written by older endpoint versions.
        $correlation = if ($Matches.ContainsKey('CorrelationId')) { ([string]$Matches['CorrelationId']).Trim() } else { '' }
        return [pscustomobject]@{ Status = $Matches.Status.Trim(); ExitCode = $Matches.ExitCode.Trim(); NextAction = $Matches.NextAction.Trim(); CorrelationId = $correlation; Detail = $Matches.Detail.Trim() }
    }
    return $null
}

function Read-SemtRemoteAttemptResult {
    <#
    Reads <DataRoot>\State\Results\<CorrelationId>.json.txt from the target. Only a document whose
    CorrelationId matches is returned; anything else is reported as foreign evidence.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RemoteDataShare,
        [Parameter(Mandatory = $true)][string]$CorrelationId
    )

    $path = Join-Path -Path $RemoteDataShare -ChildPath ("State\Results\{0}.json.txt" -f $CorrelationId)
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        # Say why: a missing data folder means the endpoint wrote nothing visible (not started,
        # access denied or blocked on the target), not only that the attempt has not finished.
        $reason = if (-not (Test-Path -LiteralPath $RemoteDataShare -PathType Container)) { 'the endpoint data folder is not visible through the administrative share' }
        elseif (-not (Test-Path -LiteralPath (Join-Path -Path $RemoteDataShare -ChildPath 'State\Results') -PathType Container)) { 'the endpoint data folder has no State\Results folder' }
        else { 'no result file for this attempt' }
        return [pscustomobject]@{ Exists = $false; State = ''; Completed = $false; Result = $null; Path = $path; Detail = ('No result for this attempt on the target ({0}).' -f $reason) }
    }
    try {
        $document = Read-SemtJsonFile -Path $path -Required
        $documentCorrelation = [string](Get-SemtPropertyValue -InputObject $document -Name 'CorrelationId' -Default '')
        if ($documentCorrelation -ne $CorrelationId) {
            return [pscustomobject]@{ Exists = $true; State = 'Foreign'; Completed = $false; Result = $null; Path = $path; Detail = ('Foreign result ignored: CorrelationId={0}' -f $documentCorrelation) }
        }
        $state = [string](Get-SemtPropertyValue -InputObject $document -Name 'State' -Default '')
        return [pscustomobject]@{ Exists = $true; State = $state; Completed = ($state -eq 'Completed'); Result = $document; Path = $path; Detail = ('Attempt result State={0}' -f $state) }
    }
    catch {
        return [pscustomobject]@{ Exists = $true; State = 'Unreadable'; Completed = $false; Result = $null; Path = $path; Detail = ('Attempt result unreadable: {0}' -f $_.Exception.Message) }
    }
}

function Wait-SemtRemoteAttemptCompletion {
    <#
    Polls the target for the completed result of this attempt after a PsExec communication
    loss or timeout, refreshing the worker reservation heartbeat while waiting.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RemoteDataShare,
        [Parameter(Mandatory = $true)][string]$CorrelationId,
        [Parameter(Mandatory = $true)][int]$WaitMinutes,
        [Parameter(Mandatory = $true)][int]$PollMinutes,
        [AllowNull()][string]$LeasePath,
        [string]$LogPath
    )

    $deadline = (Get-Date).AddMinutes($WaitMinutes)
    $last = $null
    while ($true) {
        $last = Read-SemtRemoteAttemptResult -RemoteDataShare $RemoteDataShare -CorrelationId $CorrelationId
        if ($last.Completed) { return $last }
        if ((Get-Date) -ge $deadline) { return $last }
        if ($LogPath) { Write-SemtLogLine -Path $LogPath -Message ("Completed attempt result not found yet ({0}). Next check in {1} minute(s)." -f $last.Detail,$PollMinutes) }
        $sleepUntil = (Get-Date).AddMinutes([math]::Max(1, $PollMinutes))
        if ($sleepUntil -gt $deadline) { $sleepUntil = $deadline }
        while ((Get-Date) -lt $sleepUntil) {
            Start-Sleep -Seconds 30
            [void](Update-SemtWorkerLease -LeasePath $LeasePath)
        }
    }
}

function Copy-SemtEvidenceFile {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$SourcePath,
        [Parameter(Mandatory = $true)][string]$TargetRoot,
        [Parameter(Mandatory = $true)][string]$RelativePath,
        [long]$MaxBytes = 5MB,
        [string]$Mode = 'Standard'
    )

    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) { return $false }
    try {
        $item = Get-Item -LiteralPath $SourcePath -ErrorAction Stop
        if ($Mode -ne 'Full' -and $item.Length -gt $MaxBytes) {
            Add-Content -LiteralPath (Join-Path -Path $TargetRoot -ChildPath 'CentralLogCollection.skipped.txt') -Value ('{0} Skipped large file: {1}; SizeMB={2:N2}; Source={3}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$RelativePath,($item.Length / 1MB),$SourcePath) -Encoding UTF8
            return $false
        }
        # External diagnostic tools may still emit .json. Preserve their content while
        # keeping the technician's evidence tree compatible with synchronised storage.
        if ($RelativePath -match '(?i)\.json$') {
            $RelativePath += '.txt'
            $mapped = Join-Path -Path $TargetRoot -ChildPath $RelativePath
            if ((Test-Path -LiteralPath ($SourcePath + '.txt') -PathType Leaf) -or (Test-Path -LiteralPath $mapped)) {
                Add-Content -LiteralPath (Join-Path -Path $TargetRoot -ChildPath 'CentralLogCollection.skipped.txt') -Value ('{0} Skipped JSON evidence with conflicting .json.txt name: {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$SourcePath) -Encoding UTF8
                return $false
            }
        }
        $destination = Join-Path -Path $TargetRoot -ChildPath $RelativePath
        New-SemtDirectory -Path (Split-Path -Path $destination -Parent)
        Copy-Item -LiteralPath $SourcePath -Destination $destination -Force -ErrorAction Stop
        return $true
    }
    catch {
        Add-Content -LiteralPath (Join-Path -Path $TargetRoot -ChildPath 'CentralLogCollection.skipped.txt') -Value ('{0} Copy failed: {1}; Error={2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$SourcePath,$_.Exception.Message) -Encoding UTF8
        return $false
    }
}

function ConvertTo-SemtRemoteSharePathFromLocal {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RemoteDataShare,
        [Parameter(Mandatory = $true)][string]$RemoteDataRoot,
        [AllowNull()][string]$EndpointLocalPath
    )

    if ([string]::IsNullOrWhiteSpace($EndpointLocalPath)) { return '' }
    try {
        $base = [System.IO.Path]::GetFullPath($RemoteDataRoot).TrimEnd('\')
        $local = [System.IO.Path]::GetFullPath($EndpointLocalPath)
        if (-not $local.StartsWith($base + '\', [System.StringComparison]::OrdinalIgnoreCase)) { return '' }
        return (Join-Path -Path $RemoteDataShare -ChildPath $local.Substring($base.Length).TrimStart('\'))
    }
    catch { return '' }
}

function New-SemtCentralLogTarget {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$CentralLogRoot,
        [Parameter(Mandatory = $true)][string]$ComputerPathKey,
        [Parameter(Mandatory = $true)][string]$Bucket,
        [Parameter(Mandatory = $true)][string[]]$BucketList,
        [Parameter(Mandatory = $true)][int]$CycleNumber,
        [bool]$KeepHistory = $false
    )

    if ($BucketList -notcontains $Bucket) { $Bucket = 'Errors' }
    if ($KeepHistory) {
        return (Join-Path -Path (Join-Path -Path (Join-Path -Path $CentralLogRoot -ChildPath $Bucket) -ChildPath $ComputerPathKey) -ChildPath ('Cycle{0}_{1}' -f $CycleNumber,(Get-Date -Format 'yyyyMMdd-HHmmss')))
    }
    foreach ($other in $BucketList) {
        $latest = Join-Path -Path (Join-Path -Path (Join-Path -Path $CentralLogRoot -ChildPath $other) -ChildPath $ComputerPathKey) -ChildPath 'Latest'
        if (Test-Path -LiteralPath $latest) { Remove-Item -LiteralPath $latest -Recurse -Force -ErrorAction SilentlyContinue }
    }
    return (Join-Path -Path (Join-Path -Path (Join-Path -Path $CentralLogRoot -ChildPath $Bucket) -ChildPath $ComputerPathKey) -ChildPath 'Latest')
}

function Copy-SemtCentralEvidence {
    <#
    Standard mode copies only current-attempt evidence with short destination names: the
    attempt result, LastRun.json.txt, the endpoint log and CSV named by the attempt result, and
    the campaign evidence rules (files newer than the attempt start). Full mode mirrors the
    endpoint Logs, Output, State and Transcripts folders.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$RemoteDataShare,
        [Parameter(Mandatory = $true)][string]$RemoteDataRoot,
        [Parameter(Mandatory = $true)][string]$TargetPath,
        [Parameter(Mandatory = $true)][string]$CorrelationId,
        [AllowNull()]$AttemptResult,
        [Parameter(Mandatory = $true)][datetime]$AttemptStartedUtc,
        [ValidateSet('Standard','Full')][string]$Mode = 'Standard',
        [object[]]$EvidenceRules = @()
    )

    if (Test-Path -LiteralPath $TargetPath) { Remove-Item -LiteralPath $TargetPath -Recurse -Force -ErrorAction SilentlyContinue }
    New-SemtDirectory -Path $TargetPath
    if (-not (Test-Path -LiteralPath $RemoteDataShare -PathType Container)) {
        Add-Content -LiteralPath (Join-Path -Path $TargetPath -ChildPath 'CentralLogCollection.txt') -Value ('{0} No endpoint evidence for CorrelationId={1}: the endpoint data folder {2} is not visible through the administrative share (not created, access denied, or blocked on the target). Check the target locally and the EDR console.' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$CorrelationId,$RemoteDataShare) -Encoding UTF8
        return 0
    }
    $count = 0
    if ($Mode -eq 'Full') {
        foreach ($child in @('Logs','Output','State','Transcripts','LastRun.json.txt')) {
            $source = Join-Path -Path $RemoteDataShare -ChildPath $child
            if (Test-Path -LiteralPath $source -PathType Leaf) {
                if (Copy-SemtEvidenceFile -SourcePath $source -TargetRoot $TargetPath -RelativePath $child -Mode Full) { $count++ }
            } elseif (Test-Path -LiteralPath $source -PathType Container) {
                foreach ($file in @(Get-ChildItem -LiteralPath $source -File -Recurse -ErrorAction SilentlyContinue)) {
                    $relative = Join-Path -Path $child -ChildPath ($file.FullName.Substring($source.TrimEnd('\').Length).TrimStart('\'))
                    if (Copy-SemtEvidenceFile -SourcePath $file.FullName -TargetRoot $TargetPath -RelativePath $relative -Mode Full) { $count++ }
                }
            }
        }
        return $count
    }
    Add-Content -LiteralPath (Join-Path -Path $TargetPath -ChildPath 'CentralLogCollection.txt') -Value ('{0} Standard collection for CorrelationId={1}; full logs remain on the target under {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$CorrelationId,$RemoteDataShare) -Encoding UTF8
    if (Copy-SemtEvidenceFile -SourcePath (Join-Path -Path $RemoteDataShare -ChildPath ("State\Results\{0}.json.txt" -f $CorrelationId)) -TargetRoot $TargetPath -RelativePath 'AttemptResult.json.txt' -MaxBytes 2MB) { $count++ }
    if (Copy-SemtEvidenceFile -SourcePath (Join-Path -Path $RemoteDataShare -ChildPath 'LastRun.json.txt') -TargetRoot $TargetPath -RelativePath 'LastRun.json.txt' -MaxBytes 1MB) { $count++ }
    foreach ($named in @(@{ Property = 'LogPath'; Name = 'Endpoint.log'; Max = 10MB },@{ Property = 'CsvPath'; Name = 'Result.csv'; Max = 2MB },@{ Property = 'TranscriptPath'; Name = 'Transcript.txt'; Max = 10MB })) {
        $localPath = [string](Get-SemtPropertyValue -InputObject $AttemptResult -Name $named.Property -Default '')
        $sharePath = ConvertTo-SemtRemoteSharePathFromLocal -RemoteDataShare $RemoteDataShare -RemoteDataRoot $RemoteDataRoot -EndpointLocalPath $localPath
        if (-not [string]::IsNullOrWhiteSpace($sharePath) -and (Copy-SemtEvidenceFile -SourcePath $sharePath -TargetRoot $TargetPath -RelativePath $named.Name -MaxBytes $named.Max)) { $count++ }
    }
    $minimumUtc = $AttemptStartedUtc.AddMinutes(-5)
    foreach ($rule in @($EvidenceRules)) {
        $folder = Join-Path -Path $RemoteDataShare -ChildPath ([string]$rule.Folder)
        if (-not (Test-Path -LiteralPath $folder -PathType Container)) { continue }
        $files = @(Get-ChildItem -LiteralPath $folder -File -Recurse:([bool](Get-SemtPropertyValue -InputObject $rule -Name 'Recurse' -Default $false)) -Filter ([string](Get-SemtPropertyValue -InputObject $rule -Name 'Filter' -Default '*')) -ErrorAction SilentlyContinue)
        if ([bool](Get-SemtPropertyValue -InputObject $rule -Name 'CurrentAttemptOnly' -Default $true)) { $files = @($files | Where-Object { $_.LastWriteTimeUtc -ge $minimumUtc }) }
        $names = @(Get-SemtPropertyValue -InputObject $rule -Name 'Names' -Default @())
        if ($names.Count -gt 0) { $files = @($files | Where-Object { $names -contains $_.Name }) }
        $files = @($files | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First ([int](Get-SemtPropertyValue -InputObject $rule -Name 'MaxFiles' -Default 20)))
        foreach ($file in $files) {
            $relative = Join-Path -Path ([string]$rule.TargetFolder) -ChildPath $file.Name
            if (Copy-SemtEvidenceFile -SourcePath $file.FullName -TargetRoot $TargetPath -RelativePath $relative -MaxBytes ([long](Get-SemtPropertyValue -InputObject $rule -Name 'MaxBytes' -Default 5MB))) { $count++ }
        }
    }
    return $count
}

function Publish-SemtLauncherEvidence {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$TargetPath,
        [Parameter(Mandatory = $true)][string]$WorkerLogPath
    )

    New-SemtDirectory -Path $TargetPath
    foreach ($pair in @(@($WorkerLogPath,'Worker.log'),@(('{0}.stdout.txt' -f $WorkerLogPath),'PsExec.stdout.txt'),@(('{0}.stderr.txt' -f $WorkerLogPath),'PsExec.stderr.txt'))) {
        if (Test-Path -LiteralPath $pair[0] -PathType Leaf) { Copy-Item -LiteralPath $pair[0] -Destination (Join-Path -Path $TargetPath -ChildPath $pair[1]) -Force -ErrorAction SilentlyContinue }
    }
    return $TargetPath
}

#region SEMT trusted program folder v1
# Scripts that run as SYSTEM (scheduled tasks, Intune runner, retry after reboot) are installed
# under %ProgramFiles%, which only administrators can write. ProgramData keeps only data (logs,
# state, results, setup media). The same region is copied unchanged into both endpoint scripts,
# the Intune installer and the Intune runner (a test checks that all copies are identical).
#
# A folder is trusted when it and every item below it are owned by SYSTEM, Administrators or
# TrustedInstaller and no other principal can write, append, delete, change permissions or take
# ownership. Permissions that cannot be read are never trusted. An existing folder that is not
# trusted is never repaired automatically: the operation stops (SEMT_PROGRAM_FOLDER_UNTRUSTED).

$script:SemtProgramTrustedSids = @('S-1-5-18','S-1-5-32-544','S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464')
$script:SemtProgramFolderSddl = 'O:BAD:PAI(A;OICI;FA;;;SY)(A;OICI;FA;;;BA)(A;OICI;0x1200a9;;;BU)'

function Get-SemtProgramRoot {
    param([Parameter(Mandatory = $true)][string]$Campaign, [string]$ProgramFilesPath = $env:ProgramFiles)
    if ([string]::IsNullOrWhiteSpace($ProgramFilesPath)) { throw 'SEMT_PROGRAM_FOLDER_UNTRUSTED: the Program Files folder is unknown.' }
    return (Join-Path $ProgramFilesPath ('WCH\SmartM365\SmartEndpointMigrationToolkit\{0}' -f $Campaign))
}

function Test-SemtTrustedProgramPath {
    # Read-only. State: Trusted, Untrusted (proven) or Unreadable (not proven either way).
    param([Parameter(Mandatory = $true)][string]$Path, [switch]$Recurse)
    # Write data/append/extended attributes, delete child, delete, WRITE_DAC, WRITE_OWNER and the
    # generic write/all bits that inherit-only entries may carry.
    $writeMask = 0x00000002 -bor 0x00000004 -bor 0x00000010 -bor 0x00000040 -bor 0x00010000 -bor 0x00040000 -bor 0x00080000 -bor 0x40000000 -bor 0x10000000
    try {
        $items = @(Get-Item -LiteralPath $Path -Force -ErrorAction Stop)
        if ($Recurse -and $items[0].PSIsContainer) { $items += @(Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction Stop) }
    }
    catch { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Program folder could not be read: {0}' -f $_.Exception.Message) } }
    foreach ($item in $items) {
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('Reparse point in the program folder: {0}' -f $item.FullName) } }
        try {
            $acl = Get-Acl -LiteralPath $item.FullName -ErrorAction Stop
            $owner = $acl.GetOwner([System.Security.Principal.SecurityIdentifier])
            $rules = @($acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))
        }
        catch { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Permissions of {0} could not be read: {1}' -f $item.FullName,$_.Exception.Message) } }
        if ($null -eq $owner) { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Owner of {0} could not be read.' -f $item.FullName) } }
        if ($script:SemtProgramTrustedSids -notcontains $owner.Value) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('Owner {0} of {1} is not SYSTEM, Administrators or TrustedInstaller.' -f $owner.Value,$item.FullName) } }
        foreach ($rule in $rules) {
            if ($rule.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow) { continue }
            if (([int64]$rule.FileSystemRights -band $writeMask) -eq 0) { continue }
            $sid = $rule.IdentityReference.Value
            if ($script:SemtProgramTrustedSids -contains $sid) { continue }
            # CREATOR OWNER only applies to items created later, by principals that can already write.
            if ($sid -eq 'S-1-3-0' -and ($rule.PropagationFlags -band [System.Security.AccessControl.PropagationFlags]::InheritOnly) -ne 0) { continue }
            return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('{0} can modify {1}.' -f $sid,$item.FullName) }
        }
    }
    return [pscustomobject]@{ Trusted = $true; State = 'Trusted'; Detail = '' }
}

function Set-SemtProgramFolderAcl {
    # Owner Administrators; protected DACL: SYSTEM and Administrators full control, Users read and
    # execute. Every other entry is removed. Requires an elevated or SYSTEM context.
    param([Parameter(Mandatory = $true)][string]$Path)
    $security = New-Object System.Security.AccessControl.DirectorySecurity
    $security.SetSecurityDescriptorSddlForm($script:SemtProgramFolderSddl)
    Set-Acl -LiteralPath $Path -AclObject $security -ErrorAction Stop
}

function Initialize-SemtTrustedProgramFolder {
    # Creates the folder chain below Program Files with a protected ACL, or verifies an existing
    # chain. Throws SEMT_PROGRAM_FOLDER_UNTRUSTED when any component is not trusted afterwards.
    param([Parameter(Mandatory = $true)][string]$Path, [string]$BasePath = $env:ProgramFiles)
    $base = [System.IO.Path]::GetFullPath($BasePath).TrimEnd('\')
    $target = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
    if (-not $target.StartsWith($base + '\', [System.StringComparison]::OrdinalIgnoreCase)) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} is not below {1}.' -f $target,$base) }
    $baseTrust = Test-SemtTrustedProgramPath -Path $base
    if (-not $baseTrust.Trusted) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $base,$baseTrust.State,$baseTrust.Detail) }
    $current = $base
    foreach ($segment in @($target.Substring($base.Length + 1).Split('\'))) {
        $current = Join-Path $current $segment
        if (-not (Test-Path -LiteralPath $current)) {
            New-Item -ItemType Directory -Path $current -ErrorAction Stop | Out-Null
            try { Set-SemtProgramFolderAcl -Path $current }
            catch {
                # A folder created here that cannot be protected is removed again.
                Remove-Item -LiteralPath $current -Force -ErrorAction SilentlyContinue
                throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: permissions of {0} could not be set: {1}' -f $current,$_.Exception.Message)
            }
        }
        elseif (-not (Test-Path -LiteralPath $current -PathType Container)) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} exists and is not a folder.' -f $current) }
        $trust = Test-SemtTrustedProgramPath -Path $current -Recurse:($current -eq $target)
        if (-not $trust.Trusted) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $current,$trust.State,$trust.Detail) }
    }
    return $target
}

function Install-SemtTrustedProgramFile {
    # Copies through a staging name inside the trusted folder, verifies the SHA256 against the
    # source, renames, then verifies the hash and the permissions of the installed file.
    param([Parameter(Mandatory = $true)][string]$Source, [Parameter(Mandatory = $true)][string]$Destination)
    $expected = (Get-FileHash -LiteralPath $Source -Algorithm SHA256 -ErrorAction Stop).Hash
    $folder = Split-Path -Path $Destination -Parent
    $staging = Join-Path $folder ('.{0}.{1}.tmp' -f (Split-Path -Path $Destination -Leaf),([guid]::NewGuid().ToString('N')))
    try {
        Copy-Item -LiteralPath $Source -Destination $staging -ErrorAction Stop
        if ((Get-FileHash -LiteralPath $staging -Algorithm SHA256).Hash -ne $expected) { throw ('SEMT_PROGRAM_FILE_HASH_MISMATCH: staged copy of {0} differs from the source.' -f $Source) }
        Move-Item -LiteralPath $staging -Destination $Destination -Force -ErrorAction Stop
    }
    finally { if (Test-Path -LiteralPath $staging) { Remove-Item -LiteralPath $staging -Force -ErrorAction SilentlyContinue } }
    $actual = (Get-FileHash -LiteralPath $Destination -Algorithm SHA256).Hash
    if ($actual -ne $expected) { throw ('SEMT_PROGRAM_FILE_HASH_MISMATCH: {0} Expected={1}; Actual={2}' -f $Destination,$expected,$actual) }
    $trust = Test-SemtTrustedProgramPath -Path $Destination
    if (-not $trust.Trusted) { throw ('SEMT_PROGRAM_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $Destination,$trust.State,$trust.Detail) }
    return [pscustomobject]@{ Path = $Destination; Sha256 = $actual }
}

# ProgramData data folders (logs, state, results, setup media whose setup.exe runs as SYSTEM).
# Scope: only the SEMT components (SmartEndpointMigrationToolkit and the folders below it) are
# ever changed. The shared parents above it (WCH, WCH\SmartM365), used by other products and by
# the legacy caches reused in place, are verified read-only and never changed; when they are not
# protected, the operation stops. A read-only pass checks the whole scope before the first
# change: shared parents, SEMT components, the provenance of existing content and, when a
# permission replacement is planned, every reparse point and hard link below it.
# A permission replacement never makes existing content trusted: an item that another principal
# could modify, or any content of a SEMT component that was not protected, has no verified
# provenance; it is refused and left unchanged (SEMT_DATA_FOLDER_UNTRUSTED), never adopted.
$script:SemtDataProductFolderName = 'SmartEndpointMigrationToolkit'

function Get-SemtHardLinkCount {
    # Number of names of a file; -1 when it cannot be read.
    param([Parameter(Mandatory = $true)][string]$Path)
    $fsutil = Join-Path $env:SystemRoot 'System32\fsutil.exe'
    $output = @(& $fsutil hardlink list $Path 2>$null)
    if ($LASTEXITCODE -ne 0) { return -1 }
    return @($output | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count
}

function Assert-SemtNoLinkedItem {
    # Read-only. Before a DACL replacement propagates inheritance below Path: no reparse point
    # (never followed) and only files with a single name.
    param([Parameter(Mandatory = $true)][string]$Path)
    $pending = New-Object System.Collections.Generic.Queue[string]
    $pending.Enqueue([System.IO.Path]::GetFullPath($Path))
    while ($pending.Count -gt 0) {
        $folder = $pending.Dequeue()
        try { $children = @(Get-ChildItem -LiteralPath $folder -Force -ErrorAction Stop) }
        catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be listed: {1}; nothing was changed.' -f $folder,$_.Exception.Message) }
        foreach ($child in $children) {
            if (($child.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: reparse point {0}; nothing was changed.' -f $child.FullName) }
            if ($child.PSIsContainer) { $pending.Enqueue($child.FullName); continue }
            $links = Get-SemtHardLinkCount -Path $child.FullName
            if ($links -ne 1) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} has {1} hard link name(s); nothing was changed.' -f $child.FullName,$links) }
        }
    }
}

function Test-SemtCanonicalFolderAcl {
    # True when the folder already has exactly the protected ACL that this region sets.
    param([Parameter(Mandatory = $true)][string]$Path)
    try {
        $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
        if ($acl.GetOwner([System.Security.Principal.SecurityIdentifier]).Value -ne 'S-1-5-32-544') { return $false }
        if (-not $acl.AreAccessRulesProtected) { return $false }
        $rules = @($acl.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier]) | ForEach-Object { '{0}:{1}:{2}:{3}' -f $_.IdentityReference.Value,[int64]$_.FileSystemRights,[int]$_.InheritanceFlags,$_.AccessControlType })
        $expected = @('S-1-5-18:2032127:3:Allow','S-1-5-32-544:2032127:3:Allow','S-1-5-32-545:1179817:3:Allow')
        return (@($rules | Sort-Object) -join '|') -eq (@($expected | Sort-Object) -join '|')
    }
    catch { return $false }
}

function Test-SemtTrustedSharedParent {
    # Read-only; a shared parent is never changed. Trusted when its owner is SYSTEM,
    # Administrators or TrustedInstaller and no other principal can delete it, delete its
    # children, change its permissions or take ownership. Creating new items (other products) is
    # allowed: it cannot alter the existing SEMT folder below.
    param([Parameter(Mandatory = $true)][string]$Path)
    # Delete child, delete, WRITE_DAC, WRITE_OWNER and generic all.
    $mask = 0x00000040 -bor 0x00010000 -bor 0x00040000 -bor 0x00080000 -bor 0x10000000
    try {
        $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
        $acl = Get-Acl -LiteralPath $item.FullName -ErrorAction Stop
        $owner = $acl.GetOwner([System.Security.Principal.SecurityIdentifier])
        $rules = @($acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]))
    }
    catch { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Shared folder {0} could not be read: {1}' -f $Path,$_.Exception.Message) } }
    if (-not $item.PSIsContainer) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('{0} is not a folder.' -f $item.FullName) } }
    if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('{0} is a reparse point.' -f $item.FullName) } }
    if ($null -eq $owner) { return [pscustomobject]@{ Trusted = $false; State = 'Unreadable'; Detail = ('Owner of {0} could not be read.' -f $item.FullName) } }
    if ($script:SemtProgramTrustedSids -notcontains $owner.Value) { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('Owner {0} of {1} is not SYSTEM, Administrators or TrustedInstaller.' -f $owner.Value,$item.FullName) } }
    foreach ($rule in $rules) {
        if ($rule.AccessControlType -ne [System.Security.AccessControl.AccessControlType]::Allow) { continue }
        # Inherit-only entries apply to new children only, not to this folder.
        if (($rule.PropagationFlags -band [System.Security.AccessControl.PropagationFlags]::InheritOnly) -ne 0) { continue }
        if (([int64]$rule.FileSystemRights -band $mask) -eq 0) { continue }
        if ($script:SemtProgramTrustedSids -contains $rule.IdentityReference.Value) { continue }
        return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = ('{0} can delete, rename or change the permissions of {1}.' -f $rule.IdentityReference.Value,$item.FullName) }
    }
    return [pscustomobject]@{ Trusted = $true; State = 'Trusted'; Detail = '' }
}

function Assert-SemtDataFolderAnchor {
    # The data folder must be below the base inside a SmartEndpointMigrationToolkit folder, and
    # the base and its parents must not be reparse points.
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$BasePath)
    $base = [System.IO.Path]::GetFullPath($BasePath).TrimEnd('\')
    $target = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
    if (-not $target.StartsWith($base + '\', [System.StringComparison]::OrdinalIgnoreCase)) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} is not below {1}.' -f $target,$base) }
    $segments = @($target.Substring($base.Length + 1).Split('\'))
    $productIndex = -1
    for ($i = 0; $i -lt $segments.Count; $i++) { if ($segments[$i] -ieq $script:SemtDataProductFolderName) { $productIndex = $i; break } }
    if ($productIndex -lt 0 -or $productIndex -ge ($segments.Count - 1)) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} is not a campaign folder below {1}.' -f $target,$script:SemtDataProductFolderName) }
    $current = $base
    while (-not [string]::IsNullOrWhiteSpace($current)) {
        try { $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop }
        catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be read: {1}' -f $current,$_.Exception.Message) }
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} is a reparse point.' -f $current) }
        $parent = Split-Path -Path $current -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $current) { break }
        $current = $parent
    }
    return [pscustomobject]@{ Base = $base; Target = $target; Segments = $segments; ProductIndex = $productIndex }
}

function Get-SemtDataFolderPlan {
    # Read-only pass over the whole scope. Returns one step per component (Keep, Create or
    # Replace); throws SEMT_DATA_FOLDER_UNTRUSTED, with nothing changed, when the operation
    # cannot be done within the SEMT scope or when existing content has no verified provenance.
    param([Parameter(Mandatory = $true)]$Anchor)
    $plan = New-Object System.Collections.Generic.List[object]
    $current = $Anchor.Base
    $missing = $false
    for ($index = 0; $index -lt $Anchor.Segments.Count; $index++) {
        $current = Join-Path $current $Anchor.Segments[$index]
        $shared = $index -lt $Anchor.ProductIndex
        if ($missing -or -not (Test-Path -LiteralPath $current)) {
            $missing = $true
            $plan.Add([pscustomobject]@{ Path = $current; Shared = $shared; Action = 'Create'; WasUntrusted = $false })
            continue
        }
        try { $item = Get-Item -LiteralPath $current -Force -ErrorAction Stop }
        catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be read: {1}; nothing was changed.' -f $current,$_.Exception.Message) }
        if (-not $item.PSIsContainer) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} exists and is not a folder; nothing was changed.' -f $current) }
        if (($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} is a reparse point; nothing was changed.' -f $current) }
        if ($shared) {
            $trust = Test-SemtTrustedSharedParent -Path $current
            if (-not $trust.Trusted) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: shared folder {0} ({1}) {2} It is used by other products and is never changed; nothing was changed.' -f $current,$trust.State,$trust.Detail) }
            $plan.Add([pscustomobject]@{ Path = $current; Shared = $true; Action = 'Keep'; WasUntrusted = $false })
            continue
        }
        $trust = Test-SemtTrustedProgramPath -Path $current
        if ($trust.State -eq 'Unreadable') { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} (Unreadable) {1}; nothing was changed.' -f $current,$trust.Detail) }
        if (-not $trust.Trusted) {
            # Content of a SEMT folder that was not protected has no verified provenance.
            try { $children = @(Get-ChildItem -LiteralPath $current -Force -ErrorAction Stop) }
            catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be listed: {1}; nothing was changed.' -f $current,$_.Exception.Message) }
            if ($children.Count -gt 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} was not protected ({1}) and is not empty; its content has no verified provenance and is not used; nothing was changed.' -f $current,$trust.Detail) }
            $plan.Add([pscustomobject]@{ Path = $current; Shared = $false; Action = 'Replace'; WasUntrusted = $true })
            continue
        }
        $action = if (Test-SemtCanonicalFolderAcl -Path $current) { 'Keep' } else { 'Replace' }
        $plan.Add([pscustomobject]@{ Path = $current; Shared = $false; Action = $action; WasUntrusted = $false })
    }
    # Existing content must already be trusted, item by item: below the data root and, when a
    # permission replacement is planned, everywhere it propagates to (other SEMT campaigns
    # included), so that the replacement never makes untrusted content look trusted.
    $firstReplace = @($plan | Where-Object { $_.Action -eq 'Replace' } | Select-Object -First 1)
    $existing = @($plan | Where-Object { $_.Action -ne 'Create' -and -not $_.Shared })
    $checkFrom = if ($firstReplace.Count -gt 0) { $firstReplace[0].Path } elseif (-not $missing) { $Anchor.Target } else { '' }
    if (-not [string]::IsNullOrWhiteSpace($checkFrom)) {
        $chain = @($existing | ForEach-Object { $_.Path })
        $start = [array]::IndexOf([string[]]$chain, [string]$checkFrom)
        for ($i = $start; $i -lt $chain.Count; $i++) {
            try { $children = @(Get-ChildItem -LiteralPath $chain[$i] -Force -ErrorAction Stop) }
            catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} could not be listed: {1}; nothing was changed.' -f $chain[$i],$_.Exception.Message) }
            foreach ($child in $children) {
                if ($i -lt ($chain.Count - 1) -and $child.FullName -ieq $chain[$i + 1]) { continue }
                $trust = Test-SemtTrustedProgramPath -Path $child.FullName -Recurse
                if (-not $trust.Trusted) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: existing content has no verified provenance and is not used ({0}) {1}; nothing was changed.' -f $trust.State,$trust.Detail) }
            }
        }
    }
    # A DACL replacement propagates inheritance below the folder: the whole subtree of the first
    # replaced component is checked for reparse points and hard links before any change.
    if ($firstReplace.Count -gt 0) { Assert-SemtNoLinkedItem -Path $firstReplace[0].Path }
    return $plan
}

function Initialize-SemtTrustedDataFolder {
    # Protects the SEMT data folder chain below ProgramData and verifies it. Throws
    # SEMT_DATA_FOLDER_UNTRUSTED when it cannot be made trusted within the SEMT scope.
    param([Parameter(Mandatory = $true)][string]$Path, [string]$BasePath = $env:ProgramData)
    $anchor = Assert-SemtDataFolderAnchor -Path $Path -BasePath $BasePath
    $plan = Get-SemtDataFolderPlan -Anchor $anchor
    foreach ($step in $plan) {
        if ($step.Action -eq 'Create') {
            New-Item -ItemType Directory -Path $step.Path -ErrorAction Stop | Out-Null
            if ($step.Shared) {
                # Created with inherited permissions only: a shared folder is never changed.
                $trust = Test-SemtTrustedSharedParent -Path $step.Path
                if (-not $trust.Trusted) {
                    Remove-Item -LiteralPath $step.Path -Force -ErrorAction SilentlyContinue
                    throw ('SEMT_DATA_FOLDER_UNTRUSTED: shared folder {0} created here is not protected ({1}) {2}' -f $step.Path,$trust.State,$trust.Detail)
                }
                continue
            }
            try { Set-SemtProgramFolderAcl -Path $step.Path }
            catch {
                Remove-Item -LiteralPath $step.Path -Force -ErrorAction SilentlyContinue
                throw ('SEMT_DATA_FOLDER_UNTRUSTED: permissions of {0} could not be set: {1}' -f $step.Path,$_.Exception.Message)
            }
        }
        elseif ($step.Action -eq 'Replace') {
            try { Set-SemtProgramFolderAcl -Path $step.Path }
            catch { throw ('SEMT_DATA_FOLDER_UNTRUSTED: owner and permissions of {0} could not be replaced: {1}' -f $step.Path,$_.Exception.Message) }
        }
        else { continue }
        # A folder created here, or one that was not protected (empty in the read-only pass), must
        # still be empty once protected: nothing written meanwhile is adopted.
        if ($step.Action -eq 'Create' -or $step.WasUntrusted) {
            if (@(Get-ChildItem -LiteralPath $step.Path -Force -ErrorAction Stop).Count -gt 0) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: content appeared in {0} before it was protected; it is not used.' -f $step.Path) }
        }
        $trust = Test-SemtTrustedProgramPath -Path $step.Path
        if (-not $trust.Trusted) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $step.Path,$trust.State,$trust.Detail) }
    }
    $final = Test-SemtTrustedDataFolder -Path $anchor.Target -BasePath $anchor.Base
    if (-not $final.Trusted) { throw ('SEMT_DATA_FOLDER_UNTRUSTED: {0} ({1}) {2}' -f $anchor.Target,$final.State,$final.Detail) }
    return $anchor.Target
}

function Test-SemtTrustedDataFolder {
    # Read-only verification used before anything is run: shared parents protected, every SEMT
    # component trusted, and the data root trusted with everything below it.
    param([Parameter(Mandatory = $true)][string]$Path, [string]$BasePath = $env:ProgramData)
    try { $anchor = Assert-SemtDataFolderAnchor -Path $Path -BasePath $BasePath }
    catch { return [pscustomobject]@{ Trusted = $false; State = 'Untrusted'; Detail = $_.Exception.Message } }
    $current = $anchor.Base
    for ($index = 0; $index -lt $anchor.Segments.Count; $index++) {
        $current = Join-Path $current $anchor.Segments[$index]
        if ($index -lt $anchor.ProductIndex) { $trust = Test-SemtTrustedSharedParent -Path $current }
        else { $trust = Test-SemtTrustedProgramPath -Path $current -Recurse:($current -eq $anchor.Target) }
        if (-not $trust.Trusted) { return $trust }
    }
    return [pscustomobject]@{ Trusted = $true; State = 'Trusted'; Detail = '' }
}
#endregion SEMT trusted program folder v1

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAfP6z6R7UU9QNE
# BaiDV7+6OoBo4HFmz1wUVnya8uopYaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGcfwM+fGAkVMhljfv9+9+GK9zktP+MNqB/YkzaFj3FbMA0GCSqG
# SIb3DQEBAQUABIIBgJG/EjF6g+rplhllAXZzQ9T0Ak4sle5nvZtopYWIker+HRye
# myTFAfqsDDL9M420oPQWIt/gb1eKMFtaQLkJWHYzLAKqsJ8s5sWPF9evFlj/GdyP
# a8Z/MhXjwtwxNzwmzkYSVuEvLgwqXWjFMJWBOzCNiZTxrmIkSiQ0IOYDie+qL1pQ
# HFt0+LzukAdnh/jdYC7nJ7bviSKIyvdB3lH4f/3DKhyQNq3XZUMM+KncIlAuCwsj
# 8LUWLBJ3iWpB4Ci82fJOBzT6dhIC/lRvRJRmKWmWjjGPbgTt+tghOJBXlpGyc2JY
# VBJusb37bR4HaPwGDJmLgmwI1d2mWUauKsKUTxnYvOrHXuA7gJYtID8n+Lu6/OT0
# LwYwjAcT4W4JLrkZmwbrKOIhW//SsUhfjyZ4k6OeVFYE62x2iY6G2IyhFs4GN6J7
# /JpuZosJWRiD+ArV9pVScuyxEx5E3T7mX89tWscKHFbsvfXs+qr7E9UA8aalM7Ql
# psnr1vWyT55KSlHAT6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDYyMTI5
# MTZaMC8GCSqGSIb3DQEJBDEiBCA7jx2cgdFwHILk3NVtqp2E3EL3DJGoaxKD6S7D
# x1CB9jANBgkqhkiG9w0BAQEFAASCAgB6p6N7/UmsLSUpk04lmyVmtg6IsqMYFwtb
# KYDBX/2wDXVggBhMCf9Q8zfv/SoszxCM5BOGrN3qTcOvRB4PivBnfEF04qEsWVcv
# aihxVAtNbJLwHRX4khHpO3bYKU+3rKNO93BvI0zMUGsbO7T9i9Bn1spk/mpPDIj/
# j2YKBcU6CmRCvg27w9F4ijY/fufZoW+HMTqF4iG+YxodtkTKXmxb6R8Prxt45qGt
# gz9p3G4FcStRPAsP+sPRuAIEuAopr6XWPjb4sqSp0ZMA18Pw9zT5tGNJzeQq7mSg
# ARdgpKBOyG5TfZlaUZyjwwKzrbvwb9bwrzKblv/lqJAFmkSnlRuBe9/cwmdRouGS
# 0oy6ATICF7YY/YV1SOw5Hy9FtHHh/JOwwUid7zyO9f+S6FN0ygcDgzx5HS1k+WaF
# vdD1SsEa1Iahm1qwhe/+Rlhg7tv+I2IB4kiChhV+1ZXGiyjIUhRHMZbrxs5ra/nX
# zHqBJSEg9JgleR0pqo+Cc1BX6RajycXDKTtrkUFsaMMEtBG0V8R93MX8NZzy3xu+
# +ggJvciHpd4K1/ux1Pn0WOafqf3xvwcKg1PMf9cTaLqWwvI17+6IqO5jOPv+EF8B
# 3zsDW+AnPwjtx2OwGprvrN9C0tBfvaq5GSVwl/zH4X93560z5lj5qIa8R7T1QR8X
# M9jV5/9F/w==
# SIG # End signature block
