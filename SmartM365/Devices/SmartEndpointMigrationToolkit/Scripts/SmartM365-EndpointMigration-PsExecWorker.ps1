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
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding()]
param([Parameter(Mandatory = $true)][string]$ContractPath)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$script:WorkerVersion = '1.0.0'

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
    $result.LegacyActivityStatus = [string](Get-SemtPropertyValue -InputObject $Attempt -Name 'LegacyActivityStatus' -Default '')
    $result.LegacyActivityDetail = [string](Get-SemtPropertyValue -InputObject $Attempt -Name 'LegacyActivityDetail' -Default '')
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
        $licensePreflight = Invoke-SemtUserLicensePreflight -ComputerName $target -InventoryPath (Join-Path $packageRoot 'Inventory\UsersLicenses.csv')
        $result['UserLicensePreflightState'] = $licensePreflight.State
        $result['UserLicensePreflightDetail'] = $licensePreflight.Detail
        if ($licensePreflight.Block) {
            $result['UserLicenseUpn'] = [string](Get-SemtPropertyValue -InputObject $licensePreflight -Name 'UserPrincipalName' -Default '')
            $result['UserLicenseInventoryUtc'] = [string](Get-SemtPropertyValue -InputObject $licensePreflight -Name 'CollectedUtc' -Default '')
        }
        Write-SemtLogLine -Path $logPath -Message ("User license preflight: {0}; {1}" -f $licensePreflight.State,$licensePreflight.Detail)
    }
    if ([bool]$contract.Flags.DryRun) {
        if ($access.AdminShareReachable) { Set-WorkerStatus -Status 'DRYRUN_READY' -Detail ('Dry run: DNS, ping, SMB 445 and administrative shares checked. Nothing was copied or executed. {0}' -f $access.Detail) }
        else { Set-WorkerStatus -Status 'DRYRUN_ADMIN_SHARE_UNREACHABLE' -Detail $access.Detail }
    }
    elseif ($null -ne $licensePreflight -and $licensePreflight.Block) {
        Set-WorkerStatus -Status 'USER_LICENSE_REVIEW_REQUIRED' -NextAction 'ASSIGN_INTUNE_LICENSE_THEN_RETRY' -Detail $licensePreflight.Detail
    }
    else {
        $remoteDataRoot = [string]$contract.Endpoint.RemoteDataRoot
        $remoteProductRoot = Split-Path -Path $remoteDataRoot -Parent
        # The endpoint started by PsExec as SYSTEM is copied under Program Files on the target.
        $remoteProgramRoot = [string](Get-SemtPropertyValue -InputObject $contract.Endpoint -Name 'RemoteProgramRoot' -Default '')
        if ([string]::IsNullOrWhiteSpace($remoteProgramRoot)) { throw 'SEMT_CONTRACT_INVALID: Endpoint.RemoteProgramRoot is missing; the endpoint is never started from ProgramData.' }
        if ($correlationId -notmatch '^[a-fA-F0-9]{32}$') { throw 'SEMT_CONTRACT_INVALID: invalid correlation ID.' }
        $remoteProgramRoot = Join-Path (Join-Path $remoteProgramRoot 'Attempts') $correlationId
        $copied = $null
        $copyError = ''
        for ($attempt = 1; $attempt -le 3; $attempt++) {
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
        if ($null -eq $copied) {
            if (-not $access.AdminShareReachable) { Set-WorkerStatus -Status 'ADMIN_SHARE_UNREACHABLE' -Detail $copyError }
            else { Set-WorkerStatus -Status 'REMOTE_PAYLOAD_COPY_FAILED' -Detail $copyError }
        }
        else {
            $result.RemoteScriptSha256 = $copied.RemoteSha256
            $remoteDataShare = $copied.RemoteDataShare
            $arguments = @($contract.Endpoint.Arguments | ForEach-Object { [string]$_ })
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
                $transport = Invoke-SemtSystemTaskEndpoint -Mode $remoteMode -ConnectionTarget $target -CorrelationId $correlationId -RemoteScriptPath ([string]$copied.RemoteScriptPath) -RemoteSharePath ([string]$copied.RemoteSharePath) -ExpectedSha256 $endpointSha -Arguments $arguments -LogPath $logPath -TimeoutMinutes ([int]$contract.PsExec.TimeoutMinutes) -LeasePath $leasePath
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC0cZCmiyjVsw7z
# YxScemVqWXvyGaP/A5/7lABn6+sjXKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDgvEo46V6eriohM1rpTX1q
# bV0WfyVsQ27l4VxV0kaO9DANBgkqhkiG9w0BAQEFAASCAYCTXUIA1JPqIfXTzolf
# FbyAWT/DVO1Ri8MEmUgAg1PGZz163Y75rBYDJzAbftoNAqEk6ELjCfgbudK86W1N
# DE7m0GEwhXyeb8YOmJjs/ptskDr9SWWk2VnlNZR2e1tdCQi4b1m1rIyvD5K5azjj
# WJrsGc8hfczXIzj+gog35bRVVGjjzM72RGexSVCCr5FWkdlOpd6Yon5TNwvaWAm3
# GIoi7gNr+tofIveJHGl0xPSVuGM6BdbJTPbheqAUQ6bWcU4NcXQHAbgTRYVxMqRp
# Pzk2rdwPBpK74Yi25iHIIKNNxIqCYTc9VGVR4XvZ+LOKDxAA+NtNYH4uKffJfA+3
# /PrUl9Wi3XHH659V9hmgvul0ar7hxWUrkMxd80WyMltUxtT1hI86FiYtfpEqF9sF
# NrJm+DGvU4rtk+i1K5UXlkiqhwYWJe3uyz6yPL7zsCFsaJ31INi5lMsuResmIfO5
# 2MNLWBn4Ba/Cnud1Aden6y7oXrhdpNmqRxdnHxm86plfChM=
# SIG # End signature block
