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
