<#
.SYNOPSIS
    HybridJoin campaign profile of the Smart Endpoint Migration Toolkit.

.DESCRIPTION
    Business rules ported from IntuneHybridJoinToolkit (orchestrator 2.10.78):
    - computers already present in Intune are skipped and moved to Computers.AlreadyEnrolled.txt;
    - an Entra hybrid object in Pending state passes -EntraHybridPending to the endpoint;
    - after each cycle, Intune and Entra inventories scoped to changed computers confirm
      enrollment (ENROLLED_DETECTED_POST_CYCLE) or a resolved pending object;
    - technician run guard with status back-off (history schema 2) and adaptive cycle delay.
    The module requires SmartEndpointMigration.Core to be imported first.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

Set-StrictMode -Version 2.0

$script:HybridJoinExtensionFields = @(
    'InteractiveUserName','InteractiveUserDomain','InteractiveUserAccountName','InteractiveUserAccountType',
    'InteractiveSessionName','InteractiveSessionState','UserIsUserAzureAD','UserAzureAdPrt','UserSessionIsNotRemote',
    'IntuneEnrolled','DsregAzureAdJoined','DsregDeviceAuthStatus','LeaveAttempted','RebootAttempted','RebootReason',
    'RetryAfterRebootAction','RetryAfterRebootDetail','RetryAfterRebootAttempt','RetryAfterRebootMaxAttempts','RetryAfterRebootTaskName'
)
$script:HybridJoinEngineFields = @(
    'StatusBeforePostCycle','BackoffStatus','BackoffUntilUtc','BackoffCount',
    'PostCycleIntuneInventoryChecked','PostCycleIntuneInventoryPresent','PostCycleIntuneEnrollmentDetected','PostCycleIntuneInventoryCsv','PostCycleIntuneInventoryError',
    'PostCycleEntraInventoryChecked','PostCycleEntraInventoryPresent','PostCycleEntraRegisteredState','PostCycleEntraAlternativeSecurityIdCount','PostCycleEntraPendingResolved','PostCycleEntraInventoryCsv','PostCycleEntraInventoryError',
    'PostCycleADInventoryChecked','PostCycleADInventoryPresent','PostCycleADInventoryCsv','PostCycleADInventoryError'
)

function Get-HybridJoinNextAction {
    param([AllowNull()][string]$Status)

    switch ([string]$Status) {
        'SUCCESS' { return 'NO_ACTION_ALREADY_INTUNE_OR_HEALTHY' }
        'AUDIT_SUCCESS_ALREADY_INTUNE' { return 'NO_ACTION_ALREADY_INTUNE' }
        'SKIPPED_ALREADY_INTUNE_IN_INVENTORY' { return 'NO_ACTION_ALREADY_INTUNE' }
        'ENROLLED_DETECTED_POST_CYCLE' { return 'NO_ACTION_ALREADY_INTUNE' }
        'ENTRA_PENDING_RESOLVED_POST_CYCLE' { return 'RECHECK_INTUNE_ENROLLMENT' }
        'ENDPOINT_RUN_ACTIVE' { return 'WAIT_ACTIVE_ENDPOINT_RUN' }
        'SKIPPED_BY_STATUS_BACKOFF' { return 'WAIT_STATUS_BACKOFF' }
        'RUN_GUARD_ACTIVE' { return 'WAIT_RUN_GUARD' }
        'SKIPPED_VIRTUAL_MACHINE' { return 'NO_ACTION_VIRTUAL_MACHINE' }
        'PSEXEC_EXIT_-1073741790' { return 'CHECK_EDR_OR_EXECUTION_POLICY_BLOCK' }
        default { return '' }
    }
}

function Test-HybridJoinAlreadyEnrolled {
    param([Parameter(Mandatory = $true)]$Row)

    if (Test-SemtBooleanLikeTrue -Value (Get-SemtPropertyValue -InputObject $Row -Name 'IntuneInventoryPresent' -Default '')) { return $true }
    if (Test-SemtBooleanLikeTrue -Value (Get-SemtPropertyValue -InputObject $Row -Name 'PostCycleIntuneInventoryPresent' -Default '')) { return $true }
    $status = [string](Get-SemtPropertyValue -InputObject $Row -Name 'Status' -Default '')
    $remoteStatus = [string](Get-SemtPropertyValue -InputObject $Row -Name 'RemoteStatus' -Default '')
    if ($status -in @('SUCCESS','AUDIT_SUCCESS_ALREADY_INTUNE','ENROLLED_DETECTED_POST_CYCLE','SKIPPED_ALREADY_INTUNE_IN_INVENTORY') -or $remoteStatus -in @('SUCCESS','AUDIT_SUCCESS_ALREADY_INTUNE')) { return $true }
    $nextAction = [string](Get-SemtPropertyValue -InputObject $Row -Name 'NextAction' -Default '')
    return ($nextAction -in @('NO_ACTION_ALREADY_INTUNE','NO_ACTION_ALREADY_INTUNE_OR_HEALTHY','NO_ACTION_INTUNE_PRESENT'))
}

function Get-HybridJoinCentralLogBucket {
    param([Parameter(Mandatory = $true)]$Result)

    $status = ([string](Get-SemtPropertyValue -InputObject $Result -Name 'Status' -Default '')).ToUpperInvariant()
    if ($status -in @('SUCCESS','AUDIT_SUCCESS_ALREADY_INTUNE','ENROLLED_DETECTED_POST_CYCLE')) { return 'Success' }
    if ($status -like '*ADMIN_SHARE_UNREACHABLE') { return 'AdminShareUnreachable' }
    if ($status -like 'PSEXEC_TIMEOUT*') { return 'PsExecTimeout' }
    if ($status -eq 'PSEXEC_COMMUNICATION_LOST') { return 'PsExecCommunicationLost' }
    if ($status -like 'TARGET_*' -or $status -like 'LEGACY_*' -or $status -eq 'ENDPOINT_RUN_ACTIVE') { return 'TargetBusy' }
    if ($status -like 'PSEXEC_*' -or $status -like 'REMOTE_*' -or $status -like 'DNS_*') { return 'RemoteCollectionFailed' }
    return 'Errors'
}

function Get-HybridJoinCooldownHours {
    param([string]$Status,[int]$BackoffCount = 1,[int]$CycleNumber = 1,[int]$DefaultHours = 12)

    $normalized = ([string]$Status).ToUpperInvariant()
    if ($normalized -match '^(CANCELLED_|DRYRUN_|AUDIT_|SKIPPED_VIRTUAL_MACHINE)') { return 0.0 }
    if ($normalized -match '^(ADMIN_SHARE|DNS_|PSEXEC_|REMOTE_.*FAILED|REMOTE_RESULT|CENTRAL_.*FAILED|JOB_ERROR|ERROR|TARGET_|LEGACY_)') {
        $exponent = [math]::Min(5, [math]::Max(0, [math]::Max($BackoffCount, $CycleNumber) - 1))
        return ([math]::Min(360, 15 * [math]::Pow(2, $exponent)) / 60.0)
    }
    if ($normalized -match '(USER|INTERACTIVE|PRT|LOGON|SESSION_REMOTE|LOCAL_INTERACTIVE)') { return 24.0 }
    if ($normalized -match '(RUN_GUARD|ENDPOINT_RUN_ACTIVE|PENDING_CONFIRMATION|REBOOT_TRIGGERED|WAITING_FOR_AAD_CONNECT|WAITING_POST_LEAVE)') { return 12.0 }
    if ($normalized -in @('SUCCESS','AUDIT_SUCCESS_ALREADY_INTUNE','ENROLLED_DETECTED_POST_CYCLE')) { return 168.0 }
    return [double][math]::Max(1, $DefaultHours)
}

function Test-HybridJoinRunGuardEntryRetained {
    param($Entry)

    $retention = ConvertTo-SemtUtcDateTime -Value (Get-SemtPropertyValue -InputObject $Entry -Name 'RetentionUtc' -Default (Get-SemtPropertyValue -InputObject $Entry -Name 'ExpiresUtc' -Default ''))
    return ($null -ne $retention -and $retention -gt (Get-Date).ToUniversalTime())
}

function Get-HybridJoinRunGuardDecision {
    param($Entry,[System.Collections.IDictionary]$Options)

    if ($null -eq $Entry) { return [pscustomobject]@{ Block = $false; BypassEndpointGuard = $false } }
    $expires = ConvertTo-SemtUtcDateTime -Value (Get-SemtPropertyValue -InputObject $Entry -Name 'ExpiresUtc' -Default '')
    if ($null -eq $expires -or $expires -le (Get-Date).ToUniversalTime()) { return [pscustomobject]@{ Block = $false; BypassEndpointGuard = $false } }
    $startedNoResult = ([string](Get-SemtPropertyValue -InputObject $Entry -Name 'State' -Default '') -ne 'Result')
    $backoffStatus = [string](Get-SemtPropertyValue -InputObject $Entry -Name 'BackoffStatus' -Default '')
    return [pscustomobject]@{
        Block = $true
        BypassEndpointGuard = $false
        Status = if ($startedNoResult) { 'SKIPPED_BY_TECH_RUN_GUARD_STARTED_NO_RESULT' } else { 'SKIPPED_BY_STATUS_BACKOFF' }
        NextAction = 'WAIT_STATUS_BACKOFF'
        Detail = ('Technician status back-off active. FQDN={0}; PriorStatus={1}; Count={2}; Started={3}; NextEligibleUtc={4}; State={5}' -f $Entry.ComputerFqdn,$backoffStatus,(Get-SemtPropertyValue -InputObject $Entry -Name 'BackoffCount' -Default ''),$Entry.LastStartedUtc,$Entry.ExpiresUtc,(Get-SemtPropertyValue -InputObject $Entry -Name 'State' -Default ''))
        Fields = @{ BackoffStatus = $backoffStatus; BackoffUntilUtc = [string]$Entry.ExpiresUtc; BackoffCount = [string](Get-SemtPropertyValue -InputObject $Entry -Name 'BackoffCount' -Default '') }
    }
}

function New-HybridJoinRunGuardEntry {
    param($Previous,[string]$ComputerFqdn,[string]$InputComputerName,[ValidateSet('Started','Result')][string]$State,$Result,[string]$JobId,[int]$CycleNumber,[int]$Hours,[hashtable]$Meta)

    $now = (Get-Date).ToUniversalTime()
    $started = $now
    if ($State -eq 'Result' -and $Previous) {
        $previousStarted = ConvertTo-SemtUtcDateTime -Value (Get-SemtPropertyValue -InputObject $Previous -Name 'LastStartedUtc' -Default '')
        if ($previousStarted) { $started = $previousStarted }
    }
    $launcherStatus = if ($Result) { [string](Get-SemtPropertyValue -InputObject $Result -Name 'Status' -Default '') } else { 'STARTED_NO_RESULT' }
    $remoteStatus = if ($Result) { [string](Get-SemtPropertyValue -InputObject $Result -Name 'RemoteStatus' -Default '') } else { '' }
    $previousStatus = if ($Previous) { [string](Get-SemtPropertyValue -InputObject $Previous -Name 'BackoffStatus' -Default '') } else { '' }
    $effective = if ($State -eq 'Started' -and -not [string]::IsNullOrWhiteSpace($previousStatus)) { $previousStatus } else { @($remoteStatus,$launcherStatus) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1 }
    $previousCount = 0
    if ($Previous) { [void][int]::TryParse([string](Get-SemtPropertyValue -InputObject $Previous -Name 'BackoffCount' -Default '0'), [ref]$previousCount) }
    $count = if ($State -eq 'Started') { [math]::Max(1, $previousCount) } elseif ($previousStatus -eq $effective) { $previousCount + 1 } else { 1 }
    $cooldown = if ($State -eq 'Started') { [double]$Hours } else { Get-HybridJoinCooldownHours -Status $effective -BackoffCount $count -CycleNumber $CycleNumber -DefaultHours $Hours }
    if ($State -eq 'Result' -and $cooldown -le 0) { return $null }
    return [pscustomobject]@{
        ComputerFqdn = $ComputerFqdn
        InputComputerName = $InputComputerName
        State = $State
        LastStartedUtc = $started.ToString('o')
        LastUpdatedUtc = $now.ToString('o')
        ExpiresUtc = $now.AddHours($cooldown).ToString('o')
        RetentionUtc = $now.AddDays(7).ToString('o')
        LauncherStatus = $launcherStatus
        RemoteStatus = $remoteStatus
        BackoffStatus = $effective
        BackoffCount = $count
        CooldownHours = [math]::Round($cooldown, 2)
        NextAction = if ($Result) { [string](Get-SemtPropertyValue -InputObject $Result -Name 'NextAction' -Default '') } else { '' }
        JobId = $JobId
        CycleNumber = $CycleNumber
        WaveName = [string]$Meta.WaveName
        RunId = [string]$Meta.RunId
        TechnicianAccount = [string]$Meta.TechnicianAccount
        LauncherVersion = [string]$Meta.LauncherVersion
    }
}

function Get-HybridJoinAdaptiveDelaySeconds {
    param([object[]]$Rows,[int]$MinimumDelaySeconds)

    $effectiveRows = @($Rows)
    if ($effectiveRows.Count -eq 0) { return $MinimumDelaySeconds }
    foreach ($row in $effectiveRows) {
        if ([string](Get-SemtPropertyValue -InputObject $row -Name 'Status' -Default '') -notin @('SKIPPED_BY_STATUS_BACKOFF','SKIPPED_BY_TECH_RUN_GUARD_STARTED_NO_RESULT')) { return $MinimumDelaySeconds }
    }
    $now = (Get-Date).ToUniversalTime()
    $expiries = @($effectiveRows | ForEach-Object { ConvertTo-SemtUtcDateTime -Value (Get-SemtPropertyValue -InputObject $_ -Name 'BackoffUntilUtc' -Default '') } | Where-Object { $null -ne $_ -and $_ -gt $now } | Sort-Object)
    if ($expiries.Count -eq 0) { return $MinimumDelaySeconds }
    return [math]::Max($MinimumDelaySeconds, [int][math]::Ceiling(($expiries[0] - $now).TotalSeconds))
}

function Invoke-HybridJoinPreCycle {
    <#
    Computers already present in the Intune inventory are skipped before PsExec. In dry run
    they are reported but stay in Computers.txt.
    #>
    param([Parameter(Mandatory = $true)]$Context,[string[]]$Computers)

    $intune = $Context.Inventory['Intune']
    $rows = New-Object System.Collections.Generic.List[object]
    if ($null -eq $intune -or $null -eq $intune.Map -or $intune.Map.Count -eq 0) { return @() }
    foreach ($computer in @($Computers)) {
        $key = Get-SemtComputerKey -ComputerName $computer
        if ([string]::IsNullOrWhiteSpace($key) -or -not $intune.Map.ContainsKey($key)) { continue }
        $row = New-SemtResult -Campaign 'HybridJoin' -ComputerName $computer -CycleNumber $Context.CycleNumber -WaveName $Context.RunContext.WaveName -RunId $Context.RunContext.RunId -LauncherStatus 'SKIPPED_ALREADY_INTUNE_IN_INVENTORY' -NextAction 'NO_ACTION_ALREADY_INTUNE' -Detail ('Intune inventory already reports this device ({0}); PsExec launch skipped.' -f $intune.Source) -ExtensionFields (@($script:HybridJoinExtensionFields) + @($script:HybridJoinEngineFields))
        [void](Complete-SemtResultStatus -Result $row)
        $rows.Add([pscustomobject]$row)
    }
    return @($rows.ToArray())
}

function Get-HybridJoinPerComputerArguments {
    param([Parameter(Mandatory = $true)]$Context,[Parameter(Mandatory = $true)][string]$Computer)

    $entra = $Context.Inventory['Entra']
    if ($null -eq $entra -or $null -eq $entra.Map) { return @() }
    $key = Get-SemtComputerKey -ComputerName $Computer
    if ($entra.Map.ContainsKey($key) -and [string](Get-SemtPropertyValue -InputObject $entra.Map[$key] -Name 'EntraRegisteredState' -Default '') -eq 'Pending') { return @('-EntraHybridPending') }
    return @()
}

function Invoke-HybridJoinPostCycle {
    <#
    Confirms results with Intune and Entra inventories scoped to computers whose state could
    have changed, then refreshes the AD inventory of the cycle when it is due.
    #>
    param([Parameter(Mandatory = $true)]$Context,[object[]]$Rows)

    $allRows = @($Rows | Where-Object { $null -ne $_ })
    if ($Context.DryRun -or $allRows.Count -eq 0) { return }
    $cloudRows = @($allRows | Where-Object { [string]$_.Status -notmatch '^(ADMIN_SHARE_UNREACHABLE|DNS_PREFLIGHT_ALL_SAMPLES_FAILED|REMOTE_PAYLOAD_COPY_FAILED|REMOTE_SCRIPT_MISSING|CANCELLED_|SKIPPED_|DRYRUN_|TARGET_|LEGACY_)' })
    $log = $Context.Log
    if ($cloudRows.Count -eq 0) { & $log ('Cycle {0}: post-cycle Graph confirmation skipped: no result could have changed Intune or Entra.' -f $Context.CycleNumber) 'DarkGray'; return }
    if (-not (Get-SemtOption -Options $Context.Options -Name 'HybridJoin.PostCycleCloudInventoryRefresh')) { & $log ('Cycle {0}: post-cycle Graph confirmation disabled by option.' -f $Context.CycleNumber) 'DarkGray'; return }
    if ((Get-SemtCancellationState).Requested) { return }
    $scopePath = Join-Path -Path $Context.RunContext.InventoryRoot -ChildPath ('Cycle{0}_CloudRefreshScope.txt' -f $Context.CycleNumber)
    Write-SemtTextLinesAtomic -Path $scopePath -Lines @($cloudRows | ForEach-Object { [string]$_.ComputerName } | Select-Object -Unique)
    $scoped = @($cloudRows | ForEach-Object { [string]$_.ComputerName })
    $keys = @{}
    foreach ($row in $cloudRows) { $keys[(Get-SemtComputerKey -ComputerName ([string]$row.ComputerName))] = $true }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'

    $intuneExport = Invoke-SemtInventoryExport -Kind Intune -ExportScriptPath $Context.ExportScripts.Intune -OutputPath (Join-Path $Context.RunContext.InventoryRoot ("DevicesIntune_Cycle{0}PostCycle_{1}.csv" -f $Context.CycleNumber,$stamp)) -LogPath (Join-Path $Context.RunContext.InventoryRoot ("DevicesIntune_Cycle{0}PostCycle_{1}.log" -f $Context.CycleNumber,$stamp)) -ComputerListPath $scopePath -TenantId (Get-SemtOption -Options $Context.Options -Name 'IntuneTenantId') -PageSize (Get-SemtOption -Options $Context.Options -Name 'IntuneInventoryPageSize') -SharedInventoryPath $Context.Inventory['Intune'].SharedPath
    foreach ($row in $cloudRows) {
        Set-SemtResultField -Result $row -Name 'PostCycleIntuneInventoryChecked' -Value $true
        Set-SemtResultField -Result $row -Name 'PostCycleIntuneInventoryCsv' -Value $intuneExport.CsvPath
        if (-not $intuneExport.Success) {
            Set-SemtResultField -Result $row -Name 'PostCycleIntuneInventoryError' -Value $intuneExport.Error
            continue
        }
        $key = Get-SemtComputerKey -ComputerName ([string]$row.ComputerName)
        $present = $intuneExport.Map.ContainsKey($key)
        $prePresent = Test-SemtBooleanLikeTrue -Value (Get-SemtPropertyValue -InputObject $row -Name 'IntuneInventoryPresent' -Default '')
        Set-SemtResultField -Result $row -Name 'PostCycleIntuneInventoryPresent' -Value $present
        Set-SemtResultField -Result $row -Name 'PostCycleIntuneEnrollmentDetected' -Value ([bool]($present -and -not $prePresent))
        if ($present) {
            Set-SemtResultField -Result $row -Name 'StatusBeforePostCycle' -Value ([string]$row.Status)
            Set-SemtResultField -Result $row -Name 'Status' -Value 'ENROLLED_DETECTED_POST_CYCLE'
            Set-SemtResultField -Result $row -Name 'NextAction' -Value 'NO_ACTION_ALREADY_INTUNE'
        }
    }
    if ($intuneExport.Success) {
        $Context.Inventory['Intune'].Map = Merge-SemtScopedInventoryMap -ExistingMap $Context.Inventory['Intune'].Map -RefreshedMap $intuneExport.Map -ScopedComputers $scoped
        & $log ('Cycle {0}: post-cycle Intune confirmation: {1}/{2} present; CSV={3}' -f $Context.CycleNumber,@($cloudRows | Where-Object { $_.PostCycleIntuneInventoryPresent -eq $true }).Count,$cloudRows.Count,$intuneExport.CsvPath) 'Green'
    }
    else { & $log ('Cycle {0}: post-cycle Intune confirmation failed: {1}' -f $Context.CycleNumber,$intuneExport.Error) 'Yellow' }

    if ((Get-SemtCancellationState).Requested) { return }
    $entraExport = Invoke-SemtInventoryExport -Kind Entra -ExportScriptPath $Context.ExportScripts.Entra -OutputPath (Join-Path $Context.RunContext.InventoryRoot ("DevicesEntra_Cycle{0}PostCycle_{1}.csv" -f $Context.CycleNumber,$stamp)) -LogPath (Join-Path $Context.RunContext.InventoryRoot ("DevicesEntra_Cycle{0}PostCycle_{1}.log" -f $Context.CycleNumber,$stamp)) -ComputerListPath $scopePath -TenantId (Get-SemtOption -Options $Context.Options -Name 'IntuneTenantId') -PageSize (Get-SemtOption -Options $Context.Options -Name 'IntuneInventoryPageSize') -SharedInventoryPath $Context.Inventory['Entra'].SharedPath
    foreach ($row in $cloudRows) {
        Set-SemtResultField -Result $row -Name 'PostCycleEntraInventoryChecked' -Value $true
        Set-SemtResultField -Result $row -Name 'PostCycleEntraInventoryCsv' -Value $entraExport.CsvPath
        if (-not $entraExport.Success) { Set-SemtResultField -Result $row -Name 'PostCycleEntraInventoryError' -Value $entraExport.Error; continue }
        $key = Get-SemtComputerKey -ComputerName ([string]$row.ComputerName)
        $present = $entraExport.Map.ContainsKey($key)
        $state = if ($present) { [string](Get-SemtPropertyValue -InputObject $entraExport.Map[$key] -Name 'EntraRegisteredState' -Default '') } else { '' }
        $wasPending = ([string](Get-SemtPropertyValue -InputObject $row -Name 'EntraRegisteredState' -Default '') -eq 'Pending')
        $resolved = [bool]($wasPending -and $present -and $state -ne 'Pending')
        Set-SemtResultField -Result $row -Name 'PostCycleEntraInventoryPresent' -Value $present
        Set-SemtResultField -Result $row -Name 'PostCycleEntraRegisteredState' -Value $state
        Set-SemtResultField -Result $row -Name 'PostCycleEntraAlternativeSecurityIdCount' -Value $(if ($present) { [string](Get-SemtPropertyValue -InputObject $entraExport.Map[$key] -Name 'AlternativeSecurityIdCount' -Default '') } else { '' })
        Set-SemtResultField -Result $row -Name 'PostCycleEntraPendingResolved' -Value $resolved
        if ($resolved -and [string]$row.Status -ne 'ENROLLED_DETECTED_POST_CYCLE') {
            Set-SemtResultField -Result $row -Name 'StatusBeforePostCycle' -Value ([string]$row.Status)
            Set-SemtResultField -Result $row -Name 'Status' -Value 'ENTRA_PENDING_RESOLVED_POST_CYCLE'
            Set-SemtResultField -Result $row -Name 'NextAction' -Value 'RECHECK_INTUNE_ENROLLMENT'
        }
    }
    if ($entraExport.Success) {
        $Context.Inventory['Entra'].Map = Merge-SemtScopedInventoryMap -ExistingMap $Context.Inventory['Entra'].Map -RefreshedMap $entraExport.Map -ScopedComputers $scoped
        & $log ('Cycle {0}: post-cycle Entra confirmation: pending resolved={1}; CSV={2}' -f $Context.CycleNumber,@($cloudRows | Where-Object { $_.PostCycleEntraPendingResolved -eq $true }).Count,$entraExport.CsvPath) 'Green'
    }
    else { & $log ('Cycle {0}: post-cycle Entra confirmation failed: {1}' -f $Context.CycleNumber,$entraExport.Error) 'Yellow' }
}

function Get-SemtHybridJoinProfile {
    [CmdletBinding()]
    param()

    $coreColumns = @('Timestamp','CycleNumber','ComputerName','ConnectionTarget','Status','NextAction','LauncherStatus','RemoteStatus','RemoteNextAction','ExitCode','Detail','StatusBeforePostCycle','BackoffStatus','BackoffUntilUtc','BackoffCount','DnsResolved','DnsAddressList','PingReachable','SmbPort445Reachable','AdminShareReachable','AdminShareFailureType','RemotePayloadCopyAttempts','LocalScriptVersion','LocalScriptSha256','RemoteScriptSha256','PsExecExitCode','TargetGuardStatus','LegacyActivityStatus','LegacyActivityDetail')
    $tail = @('RemoteLogsPath','PsExecLogPath','JobErrorMessage','CorrelationId','RunId','WaveName','Campaign','WorkerContractVersion')
    $columns = @($coreColumns) + @($script:HybridJoinExtensionFields) + @(Get-SemtInventoryFieldNames -Kind AD) + @(Get-SemtInventoryFieldNames -Kind Intune) + @(Get-SemtInventoryFieldNames -Kind Entra) + @($script:HybridJoinEngineFields | Where-Object { $coreColumns -notcontains $_ }) + $tail
    return [pscustomobject]@{
        Campaign = 'HybridJoin'
        Title = 'Hybrid Join and Intune enrollment'
        ProfileVersion = '1.0.0'
        EndpointScriptRelativePath = 'Campaigns\HybridJoin\Scripts\SmartM365-Invoke-EndpointMigration-HybridJoin.ps1'
        PsExecUseElevatedToken = $true
        InventoryKinds = @('AD','Intune','Entra')
        InventoryRefreshOptions = @{ AD = ''; Intune = ''; Entra = '' }
        ExtensionFields = @($script:HybridJoinExtensionFields) + @($script:HybridJoinEngineFields)
        ProjectedFields = @($script:HybridJoinExtensionFields)
        ReportColumns = @($columns | Select-Object -Unique)
        SeparatedStatuses = @('ADMIN_SHARE_UNREACHABLE','RUN_GUARD_ACTIVE','SKIPPED_BY_STATUS_BACKOFF','SKIPPED_BY_TECH_RUN_GUARD_STARTED_NO_RESULT','SKIPPED_ALREADY_INTUNE_IN_INVENTORY','DNS_PREFLIGHT_ALL_SAMPLES_FAILED')
        CentralLogBuckets = @('Success','AdminShareUnreachable','RemoteCollectionFailed','PsExecTimeout','PsExecCommunicationLost','TargetBusy','Errors')
        EvidenceRules = @(
            [pscustomobject]@{ Folder = 'Output'; Filter = '*'; Recurse = $true; CurrentAttemptOnly = $true; MaxFiles = 40; MaxBytes = 5MB; TargetFolder = 'Output' },
            [pscustomobject]@{ Folder = 'State'; Filter = '*.json'; Recurse = $false; CurrentAttemptOnly = $false; Names = @('EndpointInstance.json','RebootSafety.json','RetryAfterReboot.json'); MaxFiles = 5; MaxBytes = 1MB; TargetFolder = 'State' }
        )
        RemoteLogsPathTemplate = '\\{0}\C$\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\HybridJoin\Logs'
        CommunicationLostWaitOption = 'HybridJoin.CommunicationLostEvidenceWaitMinutes'
        CommunicationLostPollOption = 'HybridJoin.CommunicationLostEvidencePollMinutes'
        CompletedRules = @(
            [pscustomobject]@{ Label = 'already-enrolled'; ArchiveFileName = 'Computers.AlreadyEnrolled.txt'; Test = ${function:Test-HybridJoinAlreadyEnrolled} }
        )
        RunGuardSchemaVersion = 2
        Hooks = [pscustomobject]@{
            ResolveOptionRules = { param($Options) @() }
            PreCycle = ${function:Invoke-HybridJoinPreCycle}
            PerComputerArguments = ${function:Get-HybridJoinPerComputerArguments}
            PostCycle = ${function:Invoke-HybridJoinPostCycle}
            NextAction = ${function:Get-HybridJoinNextAction}
            CentralLogBucket = ${function:Get-HybridJoinCentralLogBucket}
            WorkerRecovery = { param($Result,$RemoteDataShare,$Contract) $null }
            RunGuardPrune = ${function:Test-HybridJoinRunGuardEntryRetained}
            RunGuardDecision = ${function:Get-HybridJoinRunGuardDecision}
            RunGuardEntry = ${function:New-HybridJoinRunGuardEntry}
            AdaptiveDelaySeconds = ${function:Get-HybridJoinAdaptiveDelaySeconds}
        }
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAtQCsk5bqejKsw
# jMCl8DpbemMuUOxe0blPB0VFmADdv6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIL00B5EBHzB/wgRzdzR6Li2WhT2qdQNkSOL6ZkzdjHzYMA0GCSqG
# SIb3DQEBAQUABIIBgAlX1yqpgtz6HK6lwvwxIzrCLqwc7gVr28KGsUDl9dq+yVq7
# 6n+sYPDx9hlw3Rot2JBk+fAyuqZ1uCb4bJjpyvsHd6Ux5UTktO3reoCXb1FAOtzC
# oe/H0gNH4ihRIDJYW826GME95tT6HPpYpUOuH6P/Dev5HEHggNLa4rJNoddg/uIf
# YDw71SaDy4F+pV3CLBytNvMgNjH67par7Hbsxo+uKweYMrrEVgJAFz8TMncqTXja
# AIqXb0G3kBQr6cdvd9b5TGYOuc2znbe0Kpbjf/VpmQw6Hs4HZWwv3B4t459V8sPd
# 7irLqd1WiNry+KYFW1zidmR1AHEhJvUXC1ZCYueZXoYpfMfbtwY+Vh+1NrlyjUwE
# idSXzxwPvocl2Rrh+8YrHD8YS1xhqq15+Y1sdK6pLX3sKVH0cLaURY7Rrgyk5J6P
# Qr2dphvmE/9vh/0tjDe3vZlGyqS/LCjgb79zDh/vwdvNuk32LEgKcyhY5HwrTJ4i
# WaE7c36kodj/ZjmqAKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMzE2
# NDdaMC8GCSqGSIb3DQEJBDEiBCBRuNtSQjAywkj0XJTbgC2nY2y+55VXahRBjj1X
# RfFJ1jANBgkqhkiG9w0BAQEFAASCAgAi95yF8FMd4sOEAtYqyVgQonHz3TWLLCVO
# bfHJW8eVRaeHhgVyJM9Ny49eYy1KnhP349NTssikEd4e9gerMGKO3fgd+qMvRR+1
# Gjn+BoFyC5DrCPD9Aby1JOlcdZFEWrCkCvjgCZiACVgGGLHewToO9b1eIF9RKwvc
# RYj0L6Mqqx+D5mXGqHXdy9dQtBT/rgzmQKYkgbidONqyqPyBPiWxXlZ6Pp3XZZmQ
# 1lL4nwc4K/ZEhN7c8s7apaz9McNCNEMHRvyI63kknY/h9zHZE90mas/JswV1fJ7g
# TqDtfKArgcT6/xXfU7GLNw6mUdcTOmlLB3QR3AYysNIwwoLIerWBhn3S9GqBLIPZ
# hcYuO1sxa+/I1anglZuAfGZshwptTpFAsIpK/wWfg9W1PTMqPacEf79YtwsBiXJT
# kzRBjbwLCC4D6UX6P82Q5yuvmFUDudjy7Icefxl2uncrcH5DXYnwydgxQPkNJYpq
# jcJDOxxb4Fr/+mrw2iKi22Fy5PGjlQURmhMrX8F2LNJsUD8Ntyzxeqz9pJQ0hjvp
# TS48ZIjsrA2vb/uB81P+iISjQ2mQvC73K48FHjdDPGLQpkSqHSeQmNdsq8shWSdV
# 8d4Yv5YHiFZDipb4QZ/1oi2O4GNTHtOS6SZf/l+oZgxgPmHEuNv22mP45okNfAil
# oL7oODX+Uw==
# SIG # End signature block
