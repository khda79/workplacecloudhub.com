<#
.SYNOPSIS
    Internal worker of the Smart Endpoint Migration Toolkit: processes one computer of one cycle.

.DESCRIPTION
    Started as a background job by SmartM365-Invoke-EndpointMigrationWithPsExec.ps1 with the path
    of a versioned JSON contract. The worker:
    - refuses a contract of another version, engine or package;
    - verifies the endpoint script hash itself (distribution manifest when present, launcher
      contract otherwise) before copying it;
    - copies the endpoint script through an atomic staging name and verifies it on the target;
    - runs it as SYSTEM through PsExec with per-argument quoting;
    - trusts only the attempt result whose CorrelationId matches this attempt;
    - collects current-attempt evidence into the Run CentralLogs folder.
    The target receives only the autonomous endpoint script.

.VERSION
    1.0.1

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param([Parameter(Mandatory = $true)][string]$ContractPath)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:WorkerVersion = '1.0.1'

$packageRoot = Split-Path -Path $PSScriptRoot -Parent
Import-Module (Join-Path -Path $packageRoot -ChildPath 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force -ErrorAction Stop

$contract = Read-SemtWorkerContract -Path $ContractPath
if ([System.IO.Path]::GetFullPath([string]$contract.PackageRoot).TrimEnd('\') -ine [System.IO.Path]::GetFullPath($packageRoot).TrimEnd('\')) {
    throw ("SEMT_WORKER_PACKAGE_MISMATCH: contract package {0} differs from worker package {1}." -f $contract.PackageRoot,$packageRoot)
}
$campaign = [string]$contract.Campaign
Import-Module (Join-Path -Path $packageRoot -ChildPath ("Campaigns\{0}\SmartEndpointMigration.Campaign.{0}.psd1" -f $campaign)) -MinimumVersion 1.0.0 -Force -ErrorAction Stop
$campaignProfile = & ("Get-Semt{0}Profile" -f $campaign)

$computer = [string]$contract.Computer
$target = [string]$contract.ConnectionTarget
$cycle = [int]$contract.CycleNumber
$correlationId = [string]$contract.CorrelationId
$remoteMode = [string](Get-SemtPropertyValue -InputObject $contract.Flags -Name 'RemoteExecutionMode' -Default 'PsExec')
if ($remoteMode -notin @('ScheduledTask','WinRM','PsExec')) { throw 'SEMT_CONTRACT_INVALID: unknown remote execution mode.' }
$leasePath = [string](Get-SemtPropertyValue -InputObject $contract.Gate -Name 'LeasePath' -Default '')
$workerStartedUtc = (Get-Date).ToUniversalTime()
$pathKey = Get-SemtComputerPathKey -ComputerName $computer
$logPath = Join-Path -Path ([string]$contract.Evidence.LogRoot) -ChildPath ('{0}_C{1}_{2}.log' -f $pathKey,$cycle,(Get-Date -Format 'yyyyMMdd-HHmmss'))
New-SemtDirectory -Path ([string]$contract.Evidence.LogRoot)

[void](Update-SemtWorkerLease -LeasePath $leasePath -Properties @{ WorkerProcessId = $PID; WorkerProcessStartUtc = (Get-SemtProcessStartUtc -ProcessId $PID); CorrelationId = $correlationId })

$result = New-SemtResult -Campaign $campaign -ComputerName $computer -CycleNumber $cycle -WaveName ([string]$contract.WaveName) -RunId ([string]$contract.RunId) -LauncherStatus 'UNKNOWN' -ExtensionFields @($campaignProfile.ExtensionFields)
$result.ConnectionTarget = $target
$result.RemoteExecutionMode = $remoteMode
$result.CorrelationId = $correlationId
$result.PsExecLogPath = $logPath

function Set-WorkerStatus {
    param([string]$Status,[string]$Detail,[string]$NextAction = '')
    $result.LauncherStatus = $Status
    if (-not [string]::IsNullOrWhiteSpace($Detail)) { $result.Detail = $Detail }
    if (-not [string]::IsNullOrWhiteSpace($NextAction)) { $result.NextAction = $NextAction }
}

function Get-WorkerEndpointArguments {
    param([AllowEmptyCollection()][string[]]$Arguments, [AllowNull()]$LicensePreflight)
    $resolved = @($Arguments)
    if ($null -ne $LicensePreflight -and [bool]$LicensePreflight.Block) {
        if (@($resolved | Where-Object { $_ -match '^(?i)-AuditOnly:' }).Count -gt 0) {
            throw 'SEMT_LICENSE_AUDIT_ARGUMENT_INVALID: conflicting AuditOnly argument.'
        }
        if ($resolved -notcontains '-AuditOnly') { $resolved += '-AuditOnly' }
    }
    return $resolved
}

function Invoke-WorkerPrerequisitesEndpoint {
    # Copies and runs the Prerequisites endpoint with the same transport and its own correlation ID.
    param([Parameter(Mandatory = $true)]$Auto)
    $relative = [string]$Auto.RelativePath
    $check = Test-SemtComponentIntegrity -PackageRoot $packageRoot -RelativePaths @($relative) -Mode ([string]$contract.Flags.ComponentIntegrityMode)
    $sha = $check.Files[0].Sha256
    if ($check.Blocking -or $sha -ne [string]$Auto.ExpectedSha256) {
        return [pscustomobject]@{ State = 'IntegrityMismatch'; Detail = ('Prerequisites endpoint verification failed before copy. Launcher={0}; Worker={1}; {2}' -f $Auto.ExpectedSha256,$sha,$check.Summary) }
    }
    $local = Join-Path -Path $packageRoot -ChildPath $relative
    $prerequisitesCorrelation = [guid]::NewGuid().ToString('N')
    $dataRoot = [string]$Auto.RemoteDataRoot
    $programRoot = Join-Path (Join-Path ([string]$Auto.RemoteProgramRoot) 'Attempts') $prerequisitesCorrelation
    try { $payload = Copy-SemtRemotePayload -ConnectionTarget $target -LocalScriptPath $local -ExpectedSha256 $sha -RemoteDataRoot $dataRoot -RemoteProductRoot (Split-Path -Path $dataRoot -Parent) -RemoteProgramRoot $programRoot -LogPath $logPath }
    catch { return [pscustomobject]@{ State = 'CopyFailed'; Detail = ('Prerequisites endpoint copy failed: {0}' -f $_.Exception.Message) } }
    $prerequisitesArguments = @($Auto.Arguments | ForEach-Object { [string]$_ }) + @('-CorrelationId',$prerequisitesCorrelation)
    # The Prerequisites endpoint only reads and registers tasks: 30 minutes is a generous bound.
    $timeout = [int]$contract.PsExec.TimeoutMinutes
    if ($timeout -le 0 -or $timeout -gt 30) { $timeout = 30 }
    Write-SemtLogLine -Path $logPath -Message ("Automatic prerequisites: running the Prerequisites endpoint {0}. CorrelationId={1}; Arguments={2}" -f $Auto.ExpectedVersion,$prerequisitesCorrelation,($prerequisitesArguments -join ' '))
    $transportDetail = ''
    if ($remoteMode -eq 'PsExec') {
        $run = Invoke-SemtPsExecEndpoint -PsExecPath ([string]$contract.PsExec.Path) -ConnectionTarget $target -RemoteScriptPath ([string]$payload.RemoteScriptPath) -Arguments $prerequisitesArguments -LogPath ('{0}.prerequisites' -f $logPath) -TimeoutMinutes $timeout -UseElevatedToken $false -LeasePath $leasePath
        $transportDetail = ('PsExec exit code {0}' -f $run.ExitCode)
        $done = Read-SemtRemoteAttemptResult -RemoteDataShare ([string]$payload.RemoteDataShare) -CorrelationId $prerequisitesCorrelation
        if ($done.Completed -and -not $run.TimedOut) {
            try { Remove-SemtTemporaryPayload -Path ([string]$payload.RemoteSharePath) -ExpectedSha256 $sha }
            catch { Write-SemtLogLine -Path $logPath -Level WARN -Message ('Temporary Prerequisites payload cleanup failed: {0}' -f $_.Exception.Message) }
        }
    }
    else {
        $run = Invoke-SemtSystemTaskEndpoint -Mode $remoteMode -ConnectionTarget $target -CorrelationId $prerequisitesCorrelation -RemoteScriptPath ([string]$payload.RemoteScriptPath) -RemoteSharePath ([string]$payload.RemoteSharePath) -ExpectedSha256 $sha -Arguments $prerequisitesArguments -LogPath $logPath -TimeoutMinutes $timeout -LeasePath $leasePath -RpcFailureWindowMinutes ([int](Get-SemtPropertyValue -InputObject $contract.PsExec -Name 'RpcFailureWindowMinutes' -Default 30)) -RemoteDataShare ([string]$payload.RemoteDataShare)
        $transportDetail = ('{0} completed={1}; exit code {2}' -f $remoteMode,$run.Completed,$run.ExitCode)
        $done = Read-SemtRemoteAttemptResult -RemoteDataShare ([string]$payload.RemoteDataShare) -CorrelationId $prerequisitesCorrelation
    }
    if (-not $done.Completed) {
        $cleanupDetail = 'Task cleanup not applicable (PsExec): no scheduled task was registered.'
        if ($remoteMode -eq 'PsExec') {
            Write-SemtLogLine -Path $logPath -Level WARN -Message ('Automatic prerequisites unconfirmed. {0}' -f $cleanupDetail)
        }
        else {
            # A task that never started is removed so that it cannot start later and hold the target.
            $cleanup = Remove-SemtUnstartedSystemTask -Mode $remoteMode -ConnectionTarget $target -CorrelationId $prerequisitesCorrelation -RemoteDataShare ([string]$payload.RemoteDataShare) -RemoteSharePath ([string]$payload.RemoteSharePath) -ExpectedSha256 $sha
            $cleanupDetail = ('Task cleanup {0}: {1}' -f $cleanup.Outcome,$cleanup.Detail)
            Write-SemtLogLine -Path $logPath -Level $(if ($cleanup.Outcome -in @('Deleted','Absent')) { 'INFO' } else { 'WARN' }) -Message ('Automatic prerequisites unconfirmed. {0}' -f $cleanupDetail)
        }
        return [pscustomobject]@{ State = 'Unconfirmed'; Detail = ('Prerequisites endpoint result not found ({0}). {1} {2}' -f $transportDetail,$done.Detail,$cleanupDetail) }
    }
    $status = [string](Get-SemtPropertyValue -InputObject $done.Result -Name 'Status' -Default '')
    $detail = [string](Get-SemtPropertyValue -InputObject $done.Result -Name 'Detail' -Default '')
    # SKIPPED_* (virtual machine) and TARGET_* (computer busy) are refusals, not failures.
    $state = if ($status -eq 'PREREQUISITES_READY') { 'Created' } elseif ($status -match '^(SKIPPED_|TARGET_)') { 'Skipped' } else { 'Failed' }
    return [pscustomobject]@{ State = $state; Detail = ('Prerequisites endpoint {0}: {1} {2}' -f $prerequisitesCorrelation,$status,$detail) }
}

function Invoke-WorkerAutoPrerequisites {
    <#
    HybridJoin and Windows11Upgrade only: read-only check of the scheduled tasks the campaign needs.
    When one is missing or outdated (or an obsolete one remains), the Prerequisites endpoint runs
    first, unless the attempt is audit only (reported, nothing created). The campaign endpoint
    always runs afterwards and reports any task that is still missing.
    #>
    param([AllowEmptyCollection()][string[]]$EndpointArguments)
    $auto = Get-SemtPropertyValue -InputObject $contract -Name 'Prerequisites' -Default $null
    if ($null -eq $auto -or -not [bool](Get-SemtPropertyValue -InputObject $auto -Name 'Enabled' -Default $false) -or $campaign -notin @('HybridJoin','Windows11Upgrade')) { return }
    $remote = Get-SemtRemotePrerequisiteTaskState -ConnectionTarget $target -Campaign $campaign -Mode $remoteMode
    Write-SemtLogLine -Path $logPath -Message ("Automatic prerequisites check: {0}; {1}" -f $remote.Decision,$remote.Detail)
    $state = $remote.Decision
    $detail = $remote.Detail
    if ($remote.Decision -eq 'NeedsRun') {
        if (@($EndpointArguments) -contains '-AuditOnly') {
            $state = 'MissingAuditOnly'
            $detail = ('Audit only: the Prerequisites endpoint was not run. {0}' -f $remote.Detail)
        }
        else {
            $run = Invoke-WorkerPrerequisitesEndpoint -Auto $auto
            $state = $run.State
            $detail = ('{0} Before: {1}' -f $run.Detail,$remote.Detail)
        }
    }
    $level = if ($state -in @('Ready','Created','MissingAuditOnly')) { 'INFO' } else { 'WARN' }
    Write-SemtLogLine -Path $logPath -Level $level -Message ("Automatic prerequisites: State={0}; {1}" -f $state,$detail)
    $result['PrerequisitesAutoState'] = $state
    $result['PrerequisitesAutoDetail'] = $detail
}

function Get-WorkerNextAction {
    param([string]$Status)
    $campaignAction = & $campaignProfile.Hooks.NextAction $Status
    if (-not [string]::IsNullOrWhiteSpace($campaignAction)) { return $campaignAction }
    switch -Regex ($Status) {
        '^DRYRUN_READY$' { return 'READY_FOR_REPAIR' }
        'ADMIN_SHARE_UNREACHABLE$' { return 'FIX_ADMIN_SHARE_OR_NETWORK' }
        '^REMOTE_PAYLOAD_COPY_FAILED$' { return 'FIX_SCRIPT_COPY_OR_ADMIN_SHARE' }
        '^REMOTE_SCRIPT_MISSING$' { return 'FIX_SCRIPT_COPY_OR_SECURITY' }
        '^PSEXEC_TIMEOUT_REMOTE_RUNNING$' { return 'WAIT_REMOTE_ENDPOINT_COMPLETION' }
        '^PSEXEC_TIMEOUT$' { return 'CHECK_TARGET_ENDPOINT_LOG_OR_RETRY' }
        '^PSEXEC_COMMUNICATION_LOST$' { return 'VERIFY_REMOTE_STATE_BEFORE_RELAUNCH' }
        '^PSEXEC_EXIT_-1073741790$' { return 'CHECK_EDR_OR_EXECUTION_POLICY_BLOCK' }
        '^PSEXEC_EXIT' { return 'CHECK_CURRENT_RUN_REMOTE_LOG' }
        '^REMOTE_RESULT_INCOMPLETE$' { return 'CHECK_TARGET_ENDPOINT_LOG' }
        '^REMOTE_LOG_COLLECTION_FAILED$' { return 'CHECK_TARGET_ENDPOINT_LOG' }
        '^COMPONENT_INTEGRITY_MISMATCH$' { return 'REINSTALL_SIGNED_PACKAGE' }
        '^ERROR$' { return 'CHECK_CONNECTIVITY_OR_ADMIN_ACCESS' }
        default { return 'REVIEW_LOGS' }
    }
}

function Set-ResultFromAttempt {
    param($Attempt)
    $result.RemoteStatus = [string](Get-SemtPropertyValue -InputObject $Attempt -Name 'Status' -Default '')
    $result.RemoteNextAction = [string](Get-SemtPropertyValue -InputObject $Attempt -Name 'NextAction' -Default '')
    $result.ExitCode = [string](Get-SemtPropertyValue -InputObject $Attempt -Name 'ExitCode' -Default '')
    $result.Detail = [string](Get-SemtPropertyValue -InputObject $Attempt -Name 'Detail' -Default '')
    $result.TargetGuardStatus = [string](Get-SemtPropertyValue -InputObject $Attempt -Name 'TargetGuardStatus' -Default '')
    foreach ($field in @($campaignProfile.ProjectedFields)) { $result[$field] = [string](Get-SemtPropertyValue -InputObject $Attempt -Name $field -Default '') }
    $result.LauncherStatus = 'ENDPOINT_COMPLETED'
}

try {
    Write-SemtLogLine -Path $logPath -Message ("Worker {0} started. Campaign={1}; Computer={2}; Target={3}; Cycle={4}; CorrelationId={5}; Contract={6}" -f $script:WorkerVersion,$campaign,$computer,$target,$cycle,$correlationId,$ContractPath)

    $endpointRelative = [string]$contract.Endpoint.RelativePath
    $integrity = Test-SemtComponentIntegrity -PackageRoot $packageRoot -RelativePaths @($endpointRelative) -Mode ([string]$contract.Flags.ComponentIntegrityMode)
    $endpointLocal = Join-Path -Path $packageRoot -ChildPath $endpointRelative
    $endpointSha = $integrity.Files[0].Sha256
    $result.LocalScriptSha256 = $endpointSha
    $result.LocalScriptVersion = Get-SemtScriptVersion -Path $endpointLocal
    if ($integrity.Blocking -or $endpointSha -ne [string]$contract.Endpoint.ExpectedSha256) {
        throw ("SEMT_COMPONENT_INTEGRITY_MISMATCH: endpoint script verification failed before copy. Launcher={0}; Worker={1}; {2}; {3}" -f $contract.Endpoint.ExpectedSha256,$endpointSha,$integrity.Summary,$integrity.Files[0].Detail)
    }

    $access = Test-SemtRemoteAdminAccess -ConnectionTarget $target -LogPath $logPath
    $result.DnsResolved = [string]$access.DnsResolved
    $result.DnsAddressList = [string]$access.DnsAddressList
    $result.PingReachable = [string]$access.PingReachable
    $result.SmbPort445Reachable = [string]$access.SmbPort445Reachable
    $result.AdminShareReachable = [string]$access.AdminShareReachable
    $result.AdminShareFailureType = [string]$access.FailureType

    $licensePreflight = $null
    if ($campaign -eq 'HybridJoin' -and -not [bool]$contract.Flags.DryRun -and $access.AdminShareReachable) {
        $licensePreflight = Invoke-SemtUserLicensePreflight -ComputerName $target -InventoryPath (Get-SemtSharedInventoryPath -PackageRoot $packageRoot -Kind UserLicenses)
        $result['UserLicensePreflightState'] = $licensePreflight.State
        $result['UserLicensePreflightDetail'] = $licensePreflight.Detail
        if ($licensePreflight.Block) {
            $result['UserLicenseUpn'] = [string](Get-SemtPropertyValue -InputObject $licensePreflight -Name 'UserPrincipalName' -Default '')
            $result['UserLicenseInventoryUtc'] = [string](Get-SemtPropertyValue -InputObject $licensePreflight -Name 'CollectedUtc' -Default '')
            $result['UserLicenseAuditOnly'] = 'True'
            Write-SemtLogLine -Path $logPath -Message 'User has no assigned license: endpoint will run in audit-only mode; repair and enrollment are blocked.'
        }
        Write-SemtLogLine -Path $logPath -Message ("User license preflight: {0}; {1}" -f $licensePreflight.State,$licensePreflight.Detail)
    }
    if ([bool]$contract.Flags.DryRun) {
        if ($access.AdminShareReachable) { Set-WorkerStatus -Status 'DRYRUN_READY' -Detail ('Dry run: DNS, ping, SMB 445 and administrative shares checked. Nothing was copied or executed. {0}' -f $access.Detail) }
        else { Set-WorkerStatus -Status 'DRYRUN_ADMIN_SHARE_UNREACHABLE' -Detail $access.Detail }
    }
    else {
        $remoteDataRoot = [string]$contract.Endpoint.RemoteDataRoot
        $remoteProductRoot = Split-Path -Path $remoteDataRoot -Parent
        # The endpoint started by PsExec as SYSTEM is copied under Program Files on the target.
        $remoteProgramRoot = [string](Get-SemtPropertyValue -InputObject $contract.Endpoint -Name 'RemoteProgramRoot' -Default '')
        if ([string]::IsNullOrWhiteSpace($remoteProgramRoot)) { throw 'SEMT_CONTRACT_INVALID: Endpoint.RemoteProgramRoot is missing; the endpoint is never started from ProgramData.' }
        if ($correlationId -notmatch '^[a-fA-F0-9]{32}$') { throw 'SEMT_CONTRACT_INVALID: invalid correlation ID.' }
        $remoteProgramRoot = Join-Path (Join-Path $remoteProgramRoot 'Attempts') $correlationId
        if ($access.AdminShareReachable) {
            try { Invoke-WorkerAutoPrerequisites -EndpointArguments @(Get-WorkerEndpointArguments -Arguments @($contract.Endpoint.Arguments | ForEach-Object { [string]$_ }) -LicensePreflight $licensePreflight) }
            catch {
                $result['PrerequisitesAutoState'] = 'Error'
                $result['PrerequisitesAutoDetail'] = $_.Exception.Message
                Write-SemtLogLine -Path $logPath -Level WARN -Message ("Automatic prerequisites failed; the campaign endpoint still runs: {0}" -f $_.Exception.Message)
            }
            [void](Update-SemtWorkerLease -LeasePath $leasePath)
        }
        # An unconfirmed Prerequisites run may still start later on the target and hold its operation
        # lock: the Campaign endpoint started now would then stop at TARGET_OPERATION_ACTIVE after a long
        # wait. Its task was deleted only if it never started (see Remove-SemtUnstartedSystemTask);
        # in every case this attempt stops here and the next cycle or Run tries again.
        $prerequisitesUnconfirmed = ([string]$result['PrerequisitesAutoState'] -eq 'Unconfirmed')
        $copied = $null
        $copyError = ''
        for ($attempt = 1; $attempt -le 3 -and -not $prerequisitesUnconfirmed; $attempt++) {
            $result.RemotePayloadCopyAttempts = [string]$attempt
            if ($attempt -gt 1) {
                Start-Sleep -Seconds ([math]::Min(30, 10 * ($attempt - 1)))
                [void](Update-SemtWorkerLease -LeasePath $leasePath)
                $access = Test-SemtRemoteAdminAccess -ConnectionTarget $target -LogPath $logPath
                $result.AdminShareReachable = [string]$access.AdminShareReachable
                $result.AdminShareFailureType = [string]$access.FailureType
            }
            if (-not $access.AdminShareReachable) { $copyError = ('{0}: required administrative shares are not reachable. {1}' -f $access.FailureType,$access.Detail); continue }
            try {
                $copied = Copy-SemtRemotePayload -ConnectionTarget $target -LocalScriptPath $endpointLocal -ExpectedSha256 $endpointSha -RemoteDataRoot $remoteDataRoot -RemoteProductRoot $remoteProductRoot -RemoteProgramRoot $remoteProgramRoot -LogPath $logPath
                break
            }
            catch {
                $copyError = ('Endpoint script copy failed after a successful SMB preflight. {0}; Error={1}' -f $access.Detail,$_.Exception.Message)
                Write-SemtLogLine -Path $logPath -Level WARN -Message ("Payload copy attempt {0}/3 failed: {1}" -f $attempt,$copyError)
            }
        }
        if ($prerequisitesUnconfirmed) {
            Set-WorkerStatus -Status 'REMOTE_PREREQUISITES_UNCONFIRMED' -Detail ('The automatic Prerequisites run is unconfirmed and may still start on the target; the {0} endpoint was not started. {1}' -f $campaign,[string]$result['PrerequisitesAutoDetail']) -NextAction 'VERIFY_REMOTE_STATE_BEFORE_RELAUNCH'
            Write-SemtLogLine -Path $logPath -Level WARN -Message ('Campaign endpoint not started: the automatic Prerequisites run is unconfirmed.')
        }
        elseif ($null -eq $copied) {
            if (-not $access.AdminShareReachable) { Set-WorkerStatus -Status 'ADMIN_SHARE_UNREACHABLE' -Detail $copyError }
            else { Set-WorkerStatus -Status 'REMOTE_PAYLOAD_COPY_FAILED' -Detail $copyError }
        }
        else {
            $result.RemoteScriptSha256 = $copied.RemoteSha256
            $remoteDataShare = $copied.RemoteDataShare
            $arguments = @(Get-WorkerEndpointArguments -Arguments @($contract.Endpoint.Arguments | ForEach-Object { [string]$_ }) -LicensePreflight $licensePreflight)
            if ($remoteMode -eq 'PsExec') {
            $psexec = Invoke-SemtPsExecEndpoint -PsExecPath ([string]$contract.PsExec.Path) -ConnectionTarget $target -RemoteScriptPath ([string]$copied.RemoteScriptPath) -Arguments $arguments -LogPath $logPath -TimeoutMinutes ([int]$contract.PsExec.TimeoutMinutes) -UseElevatedToken ([bool]$contract.PsExec.UseElevatedToken) -LeasePath $leasePath
            $result.PsExecExitCode = [string]$psexec.ExitCode
            Start-Sleep -Seconds 3
            $attemptResult = Read-SemtRemoteAttemptResult -RemoteDataShare $remoteDataShare -CorrelationId $correlationId
            $waitMinutes = [int](Get-SemtPropertyValue -InputObject $contract.PsExec -Name 'CommunicationLostWaitMinutes' -Default 0)
            if (-not $attemptResult.Completed -and $waitMinutes -gt 0 -and ($psexec.TimedOut -or -not [string]::IsNullOrWhiteSpace($psexec.CommunicationFailureDetail) -or [string]::IsNullOrWhiteSpace($psexec.ExitCode))) {
                Write-SemtLogLine -Path $logPath -Message ("PsExec ended without a completed result. Waiting up to {0} minute(s) for the current attempt result." -f $waitMinutes)
                $attemptResult = Wait-SemtRemoteAttemptCompletion -RemoteDataShare $remoteDataShare -CorrelationId $correlationId -WaitMinutes $waitMinutes -PollMinutes ([int](Get-SemtPropertyValue -InputObject $contract.PsExec -Name 'CommunicationLostPollMinutes' -Default 10)) -LeasePath $leasePath -LogPath $logPath
            }
            Write-SemtLogLine -Path $logPath -Message ("Attempt result: {0}" -f $attemptResult.Detail)
            if ($attemptResult.Completed) {
                Set-ResultFromAttempt -Attempt $attemptResult.Result
            }
            elseif ($psexec.FinalStatus -and $psexec.FinalStatus.CorrelationId -eq $correlationId) {
                $result.RemoteStatus = $psexec.FinalStatus.Status
                $result.RemoteNextAction = $psexec.FinalStatus.NextAction
                $result.ExitCode = $psexec.FinalStatus.ExitCode
                $result.Detail = ('{0} (from the PsExec final status line; attempt result file unavailable)' -f $psexec.FinalStatus.Detail)
                $result.LauncherStatus = 'ENDPOINT_COMPLETED'
            }
            else {
                $recovery = & $campaignProfile.Hooks.WorkerRecovery $result $remoteDataShare $contract
                if ($recovery) {
                    $result.RemoteStatus = [string]$recovery.Status
                    $result.RemoteNextAction = [string]$recovery.NextAction
                    $result.ExitCode = [string]$recovery.ExitCode
                    $result.Detail = [string]$recovery.Detail
                    $result.LauncherStatus = 'ENDPOINT_STATE_RECOVERED'
                }
                elseif ($psexec.TimedOut) {
                    if ($attemptResult.State -eq 'Running') { Set-WorkerStatus -Status 'PSEXEC_TIMEOUT_REMOTE_RUNNING' -Detail ('PsExec timed out after {0} minute(s); the endpoint attempt is still marked Running on the target.' -f $contract.PsExec.TimeoutMinutes) }
                    else { Set-WorkerStatus -Status 'PSEXEC_TIMEOUT' -Detail ('PsExec timed out after {0} minute(s); the remote state is unconfirmed.' -f $contract.PsExec.TimeoutMinutes) }
                }
                elseif (-not [string]::IsNullOrWhiteSpace($psexec.CommunicationFailureDetail)) {
                    Set-WorkerStatus -Status 'PSEXEC_COMMUNICATION_LOST' -Detail ('{0} Attempt={1}' -f $psexec.CommunicationFailureDetail,$attemptResult.Detail)
                }
                else {
                    $native = (@($psexec.Stdout) + @($psexec.Stderr)) -join "`n"
                    if ($native -match '(?i)(-File.*(does not exist|n.?existe|non esiste|no existe|nie istnieje|nicht.*exist)|fichier.*sp.cifi..*introuvable|file.*specified.*not.*found)') { Set-WorkerStatus -Status 'REMOTE_SCRIPT_MISSING' -Detail 'Remote PowerShell reported that the endpoint script path does not exist.' }
                    elseif ($attemptResult.State -eq 'Running') { Set-WorkerStatus -Status 'REMOTE_RESULT_INCOMPLETE' -Detail ('PsExec exited (code {0}) while the endpoint attempt is still marked Running on the target.' -f $psexec.ExitCode) }
                    elseif ([string]::IsNullOrWhiteSpace($psexec.ExitCode)) { Set-WorkerStatus -Status 'PSEXEC_EXIT_UNKNOWN' -Detail 'PsExec exited without a native exit code and no attempt result was found.' }
                    else {
                        Set-WorkerStatus -Status ('PSEXEC_EXIT_{0}' -f $psexec.ExitCode) -Detail ('PsExec exit code {0}; no attempt result for this correlation. {1}' -f $psexec.ExitCode,$attemptResult.Detail)
                        $result.ExitCode = [string]$psexec.ExitCode
                    }
                }
            }
            if ($attemptResult.Completed -and -not $psexec.TimedOut -and [string]::IsNullOrWhiteSpace($psexec.CommunicationFailureDetail) -and -not [string]::IsNullOrWhiteSpace($psexec.ExitCode)) {
                try { Remove-SemtTemporaryPayload -Path ([string]$copied.RemoteSharePath) -ExpectedSha256 $endpointSha; Write-SemtLogLine -Path $logPath -Message 'Temporary PsExec endpoint payload removed.' }
                catch { $result.JobErrorMessage = ('Temporary payload cleanup failed: {0}' -f $_.Exception.Message); Write-SemtLogLine -Path $logPath -Level WARN -Message $result.JobErrorMessage }
            }
            }
            else {
                $transport = Invoke-SemtSystemTaskEndpoint -Mode $remoteMode -ConnectionTarget $target -CorrelationId $correlationId -RemoteScriptPath ([string]$copied.RemoteScriptPath) -RemoteSharePath ([string]$copied.RemoteSharePath) -ExpectedSha256 $endpointSha -Arguments $arguments -LogPath $logPath -TimeoutMinutes ([int]$contract.PsExec.TimeoutMinutes) -LeasePath $leasePath -RpcFailureWindowMinutes ([int](Get-SemtPropertyValue -InputObject $contract.PsExec -Name 'RpcFailureWindowMinutes' -Default 30)) -RemoteDataShare $remoteDataShare
                $attemptResult = Read-SemtRemoteAttemptResult -RemoteDataShare $remoteDataShare -CorrelationId $correlationId
                if ($attemptResult.Completed) {
                    Set-ResultFromAttempt -Attempt $attemptResult.Result
                    if (-not $transport.Completed) { $result.JobErrorMessage = ('Remote cleanup unconfirmed: {0}' -f $transport.Detail) }
                } elseif ($transport.Completed -and $transport.ExitCode -in @('-1073741790','3221225506')) {
                    Set-WorkerStatus -Status 'REMOTE_EXECUTION_ACCESS_DENIED' -Detail ("{0} task returned 0xC0000022 (access denied), with no completed result for this attempt." -f $remoteMode) -NextAction 'CHECK_TARGET_EXECUTION_ACCESS'
                    $result.ExitCode = $transport.ExitCode
                } elseif ($transport.Completed) {
                    Set-WorkerStatus -Status 'REMOTE_RESULT_INCOMPLETE' -Detail ("{0} task finished with exit code {1}, but no completed result was found. {2}" -f $remoteMode,$transport.ExitCode,$attemptResult.Detail) -NextAction 'CHECK_TARGET_ENDPOINT_LOG'
                    $result.ExitCode = $transport.ExitCode
                } else {
                    Set-WorkerStatus -Status 'REMOTE_EXECUTION_UNCONFIRMED' -Detail ("{0}: {1}; {2}" -f $remoteMode,$transport.Detail,$attemptResult.Detail) -NextAction 'VERIFY_REMOTE_STATE_BEFORE_RELAUNCH'
                }
            }
            [void](Complete-SemtResultStatus -Result $result)
            if ($null -ne $licensePreflight -and $licensePreflight.Block -and [string]$result.RemoteStatus -like 'AUDIT_*') {
                $result.NextAction = 'ASSIGN_INTUNE_LICENSE_THEN_RETRY'
            }
            if ([string]::IsNullOrWhiteSpace([string]$result.NextAction)) { $result.NextAction = Get-WorkerNextAction -Status ([string]$result.Status) }

            if ([bool]$contract.Evidence.Enabled) {
                $bucket = & $campaignProfile.Hooks.CentralLogBucket ([pscustomobject]$result)
                $targetPath = New-SemtCentralLogTarget -CentralLogRoot ([string]$contract.Evidence.CentralLogRoot) -ComputerPathKey $pathKey -Bucket $bucket -BucketList @($campaignProfile.CentralLogBuckets) -CycleNumber $cycle -KeepHistory ([bool]$contract.Evidence.KeepHistory)
                try {
                    $copiedFiles = Copy-SemtCentralEvidence -RemoteDataShare $remoteDataShare -RemoteDataRoot $remoteDataRoot -TargetPath $targetPath -CorrelationId $correlationId -AttemptResult $attemptResult.Result -AttemptStartedUtc $workerStartedUtc -Mode ([string]$contract.Evidence.Mode) -EvidenceRules @($campaignProfile.EvidenceRules)
                    $result.RemoteLogsPath = $targetPath
                    Write-SemtLogLine -Path $logPath -Message ("Evidence collected: Bucket={0}; Files={1}; Path={2}" -f $bucket,$copiedFiles,$targetPath)
                }
                catch {
                    $result.JobErrorMessage = ('Evidence collection failed: {0}' -f $_.Exception.Message)
                    if ([string]::IsNullOrWhiteSpace([string]$result.RemoteStatus)) { Set-WorkerStatus -Status 'REMOTE_LOG_COLLECTION_FAILED' -Detail $result.JobErrorMessage }
                }
                [void](Publish-SemtLauncherEvidence -TargetPath $targetPath -WorkerLogPath $logPath)
            }
        }
    }
}
catch {
    $message = $_.Exception.Message
    if ($message -like 'SEMT_COMPONENT_INTEGRITY_MISMATCH*') { Set-WorkerStatus -Status 'COMPONENT_INTEGRITY_MISMATCH' -Detail $message }
    elseif ([string]::IsNullOrWhiteSpace([string]$result.RemoteStatus)) { Set-WorkerStatus -Status 'ERROR' -Detail $message }
    else { $result.JobErrorMessage = $message }
    Write-SemtLogLine -Path $logPath -Level ERROR -Message $message
}

[void](Complete-SemtResultStatus -Result $result)
if ([string]::IsNullOrWhiteSpace([string]$result.NextAction)) { $result.NextAction = Get-WorkerNextAction -Status ([string]$result.Status) }
if ([bool]$contract.Evidence.Enabled -and [string]::IsNullOrWhiteSpace([string]$result.RemoteLogsPath)) {
    try {
        $fallbackBucket = & $campaignProfile.Hooks.CentralLogBucket ([pscustomobject]$result)
        $fallbackTarget = New-SemtCentralLogTarget -CentralLogRoot ([string]$contract.Evidence.CentralLogRoot) -ComputerPathKey $pathKey -Bucket $fallbackBucket -BucketList @($campaignProfile.CentralLogBuckets) -CycleNumber $cycle -KeepHistory ([bool]$contract.Evidence.KeepHistory)
        $result.RemoteLogsPath = Publish-SemtLauncherEvidence -TargetPath $fallbackTarget -WorkerLogPath $logPath
    }
    catch { }
}
Write-SemtLogLine -Path $logPath -Message ("Worker completed. Status={0}; NextAction={1}; ExitCode={2}" -f $result.Status,$result.NextAction,$result.ExitCode)
[pscustomobject]$result

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCx6/fzVVowbBAB
# /pHY+Fiag7zlmR2lnHFvd9hCvaMcLKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINABgZJBMfZfxXslQ+3i7foQ63xRtK7wWquynLoYTgvKMA0GCSqG
# SIb3DQEBAQUABIIBgEhW/Xh7ky+IbH4jYt69my2bEglMf1M0WJOIba9aG66UbULN
# gykxCzE8EMXDAMg33YUz1BqJhywnqDIzJSLgnINjZ6ay8ew5QyJ/mDdTZxaeZGP0
# U9Z5uNTakVoGNPLazXdSvWgz2GbIEgfpAhOcGn2qy+OZ2K8zvZpqCGTSsXDibLNm
# dqCKbYnrAvKIL3wJHvk3t1/QhpQQ2unad5NOYiIdoIh0m5L6RLd7AGuxCYRCF04B
# hV1OOsa/Fx+wLIJUx9oP9hsWlvgBzvWXO4Y0LsHqXJ5oUPW9lw6Egu5YC3DFvLB2
# FG9M7zoimuvxp7R+mki6O4w1gvHdhZlbM7iwwIh5u3nqNS3RukPtRbEH0n6gigs3
# Shx1rKCQ3Wc808LbzdOt79aOoZlqDtL/jK/PrVGJ76Krg64STn87YC/KV/yMvs5M
# W3ZA3h2H1FBz+9UoTcRby4pqdANBMgVybiXIrra1SaTHfo1Ahl72VlZ+55csdP2S
# EMAYnxNJtozXwqdOTKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDYxMjMx
# MjdaMC8GCSqGSIb3DQEJBDEiBCA4xXm2ovi/qed4ZacKg5TAfLIQIs7eTbrNuLT4
# EzKc7zANBgkqhkiG9w0BAQEFAASCAgBM57Ww3feyCCnA8K9h4uexfse5GZB43Nqv
# vZfoAN+W0jGQbeAAXrFxw1snpG0hxuQt2N4rfX0nUHJ0Sq3kIdaPgsYcsTe0Avmv
# lLM35fj+qFLxok5BqG7LQmw1FGb7KZYKMcVc6A47a77zwVmSqEk3DcXjNYIIGOwy
# j59VHdFbMYHMbaI+c1TkLXNPKdSlcqxCAOcFt/6ozsF2d57qGdw4HS2p99886Omb
# 9Oo6HBHVlv3y5001+NgQRpUoHufXUYRcMyEYAebUCA6Z9DN3eppaZ2XwXEbVLcc1
# 3BFBRq1pA7MJMlFzj2jjI5lidVxaW5JQ4p65hX4+fsmnwGiODRD4THpw1XYy65Km
# UsPCbQJbm9hjEiVOrmXsxfdrPTmm+LXMDECJk+F53KtvbSqNVFqrMobvT5VwtLKR
# 0eRiyZz9kGKqsJtRbry6PvscfA3DJU12KlAHIUFbc70YeUnXsOiyC4azo2gVPtoI
# 7JwK1gxteA0PeA94MYXvogcUyqlhNs8HpAhvP8AGlls9iTz7zfYsHZP2RCOThEhG
# EoNAEns1DrOx1vx/xuGYn+tQa92II2hZFeRwu7/j0n6KepRa3R7g1U882danHkgz
# AftuGxm9iawemoopKpYEvo+9KOL8z4SCWlmn5b8zJG5halQHTCJC+qEpqrl+I+xi
# F9Dnte8hsQ==
# SIG # End signature block
