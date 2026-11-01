<#
.SYNOPSIS
    Prerequisites campaign profile of the Smart Endpoint Migration Toolkit.

.DESCRIPTION
    The Prerequisites Campaign creates on targets the scheduled tasks that the HybridJoin and
    Windows11Upgrade endpoint scripts only start, enable or disable:
    - HybridJoin: three on-demand user-context tasks (dsregcmd /status, dsregcmd /refreshprt,
      deviceenroller /AutoEnrollMDM), the next-logon task (created disabled) and the SYSTEM
      retry-after-reboot task (created disabled);
    - Windows11Upgrade: the SYSTEM retry-after-reboot task (created disabled).
    The Campaign is disabled by default (option Prerequisites.Enabled). It reads no inventory.
    Computers whose tasks are ready leave Computers.txt.
    The module requires SmartEndpointMigration.Core to be imported first.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

Set-StrictMode -Version 2.0

$script:PrerequisitesExtensionFields = @('TasksExpected','TasksPresent','TasksCreated','TasksUpdated','TasksMissing','TasksFailed','TasksDetail')
$script:PrerequisitesReadyStatuses = @('PREREQUISITES_READY','AUDIT_PREREQUISITES_PRESENT')

function Get-PrerequisitesNextAction {
    param([AllowNull()][string]$Status)

    switch ([string]$Status) {
        'PREREQUISITES_READY' { return 'NO_ACTION_PREREQUISITES_READY' }
        'AUDIT_PREREQUISITES_PRESENT' { return 'NO_ACTION_PREREQUISITES_READY' }
        'AUDIT_PREREQUISITES_MISSING' { return 'RUN_PREREQUISITES_WITHOUT_AUDIT' }
        'PREREQUISITES_TASK_FAILED' { return 'CHECK_EDR_OR_TASK_SCHEDULER' }
        'PREREQUISITES_NOTHING_SELECTED' { return 'SELECT_PREREQUISITES_TASKS' }
        'SKIPPED_VIRTUAL_MACHINE' { return 'NO_ACTION_VIRTUAL_MACHINE' }
        'PSEXEC_EXIT_-1073741790' { return 'CHECK_EDR_OR_EXECUTION_POLICY_BLOCK' }
        default { return '' }
    }
}

function Test-PrerequisitesReady {
    param([Parameter(Mandatory = $true)]$Row)
    # Only a repair run that confirmed every requested task leaves the Wave (never an audit).
    return ('PREREQUISITES_READY' -in @([string](Get-SemtPropertyValue -InputObject $Row -Name 'Status' -Default ''),[string](Get-SemtPropertyValue -InputObject $Row -Name 'RemoteStatus' -Default '')))
}

function Get-PrerequisitesCentralLogBucket {
    param([Parameter(Mandatory = $true)]$Result)

    $status = ([string](Get-SemtPropertyValue -InputObject $Result -Name 'Status' -Default '')).ToUpperInvariant()
    if ($status -in $script:PrerequisitesReadyStatuses) { return 'Success' }
    if ($status -like '*ADMIN_SHARE_UNREACHABLE') { return 'AdminShareUnreachable' }
    if ($status -like 'PSEXEC_TIMEOUT*') { return 'PsExecTimeout' }
    if ($status -eq 'PSEXEC_COMMUNICATION_LOST') { return 'PsExecCommunicationLost' }
    if ($status -like 'SKIPPED_*') { return 'Skipped' }
    if ($status -like 'TARGET_*') { return 'TargetBusy' }
    if ($status -like 'PSEXEC_*' -or $status -like 'REMOTE_*' -or $status -like 'DNS_*') { return 'RemoteCollectionFailed' }
    return 'Errors'
}

function Get-PrerequisitesCooldownHours {
    param([string]$Status,[int]$Hours)

    $normalized = ([string]$Status).ToUpperInvariant()
    if ($normalized -in $script:PrerequisitesReadyStatuses) { return [double][math]::Max(1, $Hours) }
    if ($normalized -match '^(CANCELLED_|DRYRUN_|AUDIT_|SKIPPED_VIRTUAL_MACHINE)') { return 0.0 }
    if ($normalized -match '^(ADMIN_SHARE|DNS_|PSEXEC_|REMOTE_|TARGET_|JOB_ERROR)') { return 0.25 }
    return 1.0
}

function Test-PrerequisitesRunGuardEntryRetained {
    param($Entry)

    $retention = ConvertTo-SemtUtcDateTime -Value (Get-SemtPropertyValue -InputObject $Entry -Name 'RetentionUtc' -Default '')
    return ($null -ne $retention -and $retention -gt (Get-Date).ToUniversalTime())
}

function Get-PrerequisitesRunGuardDecision {
    param($Entry,[System.Collections.IDictionary]$Options)

    $none = [pscustomobject]@{ Block = $false; BypassEndpointGuard = $false }
    if ($null -eq $Entry) { return $none }
    $expires = ConvertTo-SemtUtcDateTime -Value (Get-SemtPropertyValue -InputObject $Entry -Name 'ExpiresUtc' -Default '')
    if ($null -eq $expires -or $expires -le (Get-Date).ToUniversalTime()) { return $none }
    $startedNoResult = ([string](Get-SemtPropertyValue -InputObject $Entry -Name 'State' -Default '') -ne 'Result')
    return [pscustomobject]@{
        Block = $true
        BypassEndpointGuard = $false
        Status = if ($startedNoResult) { 'SKIPPED_BY_TECH_RUN_GUARD_STARTED_NO_RESULT' } else { 'SKIPPED_BY_TECH_RUN_GUARD' }
        NextAction = 'WAIT_RUN_GUARD_EXPIRY'
        Detail = ('Technician run guard. FQDN={0}; LastStatus={1}; LastStartedUtc={2}; NextEligibleUtc={3}; State={4}' -f $Entry.ComputerFqdn,(Get-SemtPropertyValue -InputObject $Entry -Name 'Status' -Default ''),$Entry.LastStartedUtc,$Entry.ExpiresUtc,(Get-SemtPropertyValue -InputObject $Entry -Name 'State' -Default ''))
        Fields = @{}
    }
}

function New-PrerequisitesRunGuardEntry {
    param($Previous,[string]$ComputerFqdn,[string]$InputComputerName,[ValidateSet('Started','Result')][string]$State,$Result,[string]$JobId,[int]$CycleNumber,[int]$Hours,[hashtable]$Meta)

    $now = (Get-Date).ToUniversalTime()
    $started = $now
    if ($State -eq 'Result' -and $Previous) {
        $previousStarted = ConvertTo-SemtUtcDateTime -Value (Get-SemtPropertyValue -InputObject $Previous -Name 'LastStartedUtc' -Default '')
        if ($previousStarted) { $started = $previousStarted }
    }
    $status = if ($Result) { @([string](Get-SemtPropertyValue -InputObject $Result -Name 'RemoteStatus' -Default ''),[string](Get-SemtPropertyValue -InputObject $Result -Name 'Status' -Default '')) | Where-Object { $_ } | Select-Object -First 1 } else { 'STARTED_NO_RESULT' }
    $cooldown = if ($State -eq 'Started') { [double][math]::Max(1, $Hours) } else { Get-PrerequisitesCooldownHours -Status ([string]$status) -Hours $Hours }
    if ($State -eq 'Result' -and $cooldown -le 0) { return $null }
    return [pscustomobject]@{
        ComputerFqdn = $ComputerFqdn
        InputComputerName = $InputComputerName
        State = $State
        LastStartedUtc = $started.ToString('o')
        LastUpdatedUtc = $now.ToString('o')
        ExpiresUtc = $now.AddHours($cooldown).ToString('o')
        RetentionUtc = $now.AddDays(7).ToString('o')
        Status = [string]$status
        NextAction = if ($Result) { [string](Get-SemtPropertyValue -InputObject $Result -Name 'NextAction' -Default '') } else { '' }
        CooldownHours = [math]::Round($cooldown, 2)
        JobId = $JobId
        CycleNumber = $CycleNumber
        WaveName = [string]$Meta.WaveName
        RunId = [string]$Meta.RunId
        TechnicianAccount = [string]$Meta.TechnicianAccount
        LauncherVersion = [string]$Meta.LauncherVersion
    }
}

function Get-SemtPrerequisitesProfile {
    [CmdletBinding()]
    param()

    $coreColumns = @('Timestamp','CycleNumber','ComputerName','ConnectionTarget','RemoteExecutionMode','Status','NextAction','LauncherStatus','RemoteStatus','RemoteNextAction','ExitCode','Detail','DnsResolved','DnsAddressList','PingReachable','SmbPort445Reachable','AdminShareReachable','AdminShareFailureType','RemotePayloadCopyAttempts','LocalScriptVersion','LocalScriptSha256','RemoteScriptSha256','PsExecExitCode','TargetGuardStatus')
    $tail = @('RemoteLogsPath','PsExecLogPath','JobErrorMessage','CorrelationId','RunId','WaveName','Campaign','WorkerContractVersion')
    $columns = @($coreColumns) + @($script:PrerequisitesExtensionFields) + $tail
    return [pscustomobject]@{
        Campaign = 'Prerequisites'
        Title = 'Prerequisites (scheduled tasks)'
        ProfileVersion = '1.0.0'
        EndpointScriptRelativePath = 'Campaigns\Prerequisites\Scripts\SmartM365-Invoke-EndpointMigration-Prerequisites.ps1'
        PsExecUseElevatedToken = $false
        InventoryKinds = @()
        InventoryRefreshOptions = @{}
        ExtensionFields = @($script:PrerequisitesExtensionFields)
        ProjectedFields = @($script:PrerequisitesExtensionFields)
        ReportColumns = @($columns | Select-Object -Unique)
        SeparatedStatuses = @('ADMIN_SHARE_UNREACHABLE','SKIPPED_BY_TECH_RUN_GUARD','SKIPPED_BY_TECH_RUN_GUARD_STARTED_NO_RESULT','DNS_PREFLIGHT_ALL_SAMPLES_FAILED')
        CentralLogBuckets = @('Success','Skipped','AdminShareUnreachable','RemoteCollectionFailed','PsExecTimeout','PsExecCommunicationLost','TargetBusy','Errors')
        EvidenceRules = @()
        RemoteLogsPathTemplate = '\\{0}\C$\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\Prerequisites\Logs'
        CommunicationLostWaitOption = ''
        CommunicationLostPollOption = ''
        CompletedRules = @(
            [pscustomobject]@{ Label = 'prerequisites-ready'; ArchiveFileName = 'Computers.PrerequisitesReady.txt'; Test = ${function:Test-PrerequisitesReady} }
        )
        RunGuardSchemaVersion = 1
        Hooks = [pscustomobject]@{
            ResolveOptionRules = { param($Options) @() }
            PreCycle = { param($Context,$Computers) @() }
            PerComputerArguments = { param($Context,$Computer) @() }
            PostCycle = { param($Context,$Rows) }
            NextAction = ${function:Get-PrerequisitesNextAction}
            CentralLogBucket = ${function:Get-PrerequisitesCentralLogBucket}
            WorkerRecovery = { param($Result,$RemoteDataShare,$Contract) $null }
            RunGuardPrune = ${function:Test-PrerequisitesRunGuardEntryRetained}
            RunGuardDecision = ${function:Get-PrerequisitesRunGuardDecision}
            RunGuardEntry = ${function:New-PrerequisitesRunGuardEntry}
            AdaptiveDelaySeconds = { param($Rows,$MinimumDelaySeconds) $MinimumDelaySeconds }
        }
    }
}
