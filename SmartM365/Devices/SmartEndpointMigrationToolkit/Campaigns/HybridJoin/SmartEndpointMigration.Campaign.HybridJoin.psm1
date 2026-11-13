<#
.SYNOPSIS
    HybridJoin campaign profile of the Smart Endpoint Migration Toolkit.

.DESCRIPTION
    Business rules ported from IntuneHybridJoinToolkit (orchestrator 2.10.78):
    - name-only Intune matches never skip inspection; confirmed identities can be archived;
    - an Entra hybrid object in Pending state passes -EntraHybridPending to the endpoint;
    - after each cycle, Intune and Entra inventories scoped to changed computers confirm
      enrollment (ENROLLED_DETECTED_POST_CYCLE) or a resolved pending object;
    - a local Intune trace with an unhealthy device join (INTUNE_TRACE_DEVICE_JOIN_UNHEALTHY) is
      reported for manual review, never as enrolled; with RepairUnhealthyJoinWithIntuneTrace, the
      launcher passes -IntuneAbsentDeviceId when fresh inventories prove the device ID absent from
      Intune, and the endpoint repairs once its local trace is old enough;
    - technician run guard with status back-off (history schema 2) and adaptive cycle delay.
    The module requires SmartEndpointMigration.Core to be imported first.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

Set-StrictMode -Version 2.0

$script:HybridJoinExtensionFields = @(
    'InteractiveUserName','InteractiveUserSid','InteractiveUserDomain','InteractiveUserAccountName','InteractiveUserAccountType',
    'InteractiveSessionName','InteractiveSessionState','UserIsUserAzureAD','UserAzureAdPrt','UserSessionIsNotRemote',
    'GpUpdateTimedOut','GpUpdateStopDetail','FinalGpUpdateTimedOut','FinalGpUpdateStopDetail','MdmDiagnosticStatus','MdmDiagnosticZip','MdmEnrollmentRecentEvidence','IntuneEnrolled','Dsreg_DeviceId','Dsreg_TenantId','DsregAzureAdJoined','DsregDeviceAuthStatus','LeaveAttempted','RebootAttempted','RebootReason',
    'RetryAfterRebootAction','RetryAfterRebootDetail','RetryAfterRebootAttempt','RetryAfterRebootMaxAttempts','RetryAfterRebootTaskName'
)
$script:HybridJoinEngineFields = @(
    'StatusBeforePostCycle','BackoffStatus','BackoffUntilUtc','BackoffCount',
    'IntuneConfirmationState','IntuneConfirmationDetail','IntuneConfirmationSource','UserLicensePreflightState','UserLicensePreflightDetail','UserLicenseAuditOnly','UserLicenseState','UserLicenseDetail','UserLicenseUpn','UserLicenseInventoryUtc','MdmUserScopeState','PostCycleIntuneInventoryChecked','PostCycleIntuneInventoryPresent','PostCycleIntuneEnrollmentDetected','PostCycleIntuneInventoryCsv','PostCycleIntuneInventoryError',
    'PostCycleEntraInventoryChecked','PostCycleEntraInventoryPresent','PostCycleEntraRegisteredState','PostCycleEntraAlternativeSecurityIdCount','PostCycleEntraPendingResolved','PostCycleEntraInventoryCsv','PostCycleEntraInventoryError',
    'PostCycleADInventoryChecked','PostCycleADInventoryPresent','PostCycleADInventoryCsv','PostCycleADInventoryError',
    'PrerequisitesAutoState','PrerequisitesAutoDetail'
)

function Get-HybridJoinNextAction {
    param([AllowNull()][string]$Status)

    switch ([string]$Status) {
        'SUCCESS' { return 'NO_ACTION_ALREADY_INTUNE_OR_HEALTHY' }
        'AUDIT_SUCCESS_ALREADY_INTUNE' { return 'NO_ACTION_ALREADY_INTUNE' }
        'SKIPPED_ALREADY_INTUNE_IN_INVENTORY' { return 'NO_ACTION_ALREADY_INTUNE' }
        'ENROLLED_DETECTED_POST_CYCLE' { return 'NO_ACTION_ALREADY_INTUNE' }
        'ENROLLED_CONFIRMED_IN_CYCLE' { return 'NO_ACTION_ALREADY_INTUNE' }
        'ENROLLED_CONFIRMED_BY_CLOUD_IDENTITY' { return 'NO_ACTION_ALREADY_INTUNE' }
        'ENTRA_PENDING_RESOLVED_POST_CYCLE' { return 'RECHECK_INTUNE_ENROLLMENT' }
        'INTUNE_TRACE_DEVICE_JOIN_UNHEALTHY' { return 'REVIEW_INTUNE_TRACE_AND_DEVICE_JOIN' }
        'INTUNE_TRACE_REMOVE_FAILED' { return 'REVIEW_INTUNE_TRACE_AND_DEVICE_JOIN' }
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
    # Archival requires the same device identity in a successful cloud result: local identity
    # (post-cycle or in-cycle confirmation) or AD/Entra/Intune identity chain (unreachable target).
    if ([string](Get-SemtPropertyValue -InputObject $Row -Name 'UserLicensePreflightState' -Default '') -eq 'NoAssignedLicense') { return $false }
    $state = [string](Get-SemtPropertyValue -InputObject $Row -Name 'IntuneConfirmationState' -Default '')
    switch ([string](Get-SemtPropertyValue -InputObject $Row -Name 'Status' -Default '')) {
        'ENROLLED_DETECTED_POST_CYCLE' { return ($state -eq 'Confirmed') }
        'ENROLLED_CONFIRMED_IN_CYCLE' { return ($state -eq 'Confirmed') }
        'ENROLLED_CONFIRMED_BY_CLOUD_IDENTITY' { return ($state -eq 'CloudIdentityConfirmed') }
        default { return $false }
    }
}

function Get-HybridJoinIntuneConfirmation {
    param([Parameter(Mandatory = $true)]$Row,[object[]]$CloudRows)
    $localId = [guid]::Empty
    $localRaw = [string](Get-SemtPropertyValue -InputObject $Row -Name 'Dsreg_DeviceId' -Default '')
    if (-not [guid]::TryParse($localRaw, [ref]$localId) -or $localId -eq [guid]::Empty) {
        return [pscustomobject]@{ State='Unknown'; Detail='Local Entra device ID unavailable'; Present=$null }
    }
    $localTenant = [guid]::Empty
    $tenantRaw = [string](Get-SemtPropertyValue -InputObject $Row -Name 'Dsreg_TenantId' -Default '')
    if (-not [guid]::TryParse($tenantRaw, [ref]$localTenant) -or $localTenant -eq [guid]::Empty) {
        return [pscustomobject]@{ State='Unknown'; Detail='Local tenant ID unavailable'; Present=$null }
    }
    $scopeTenants = @($CloudRows | ForEach-Object { [string](Get-SemtPropertyValue -InputObject $_ -Name 'InventoryTenantId' -Default '') } | Select-Object -Unique)
    if ($scopeTenants.Count -ne 1 -or $scopeTenants[0] -ine $localTenant.ToString()) {
        return [pscustomobject]@{ State='Unknown'; Detail='Scoped Intune inventory tenant unavailable or mismatched'; Present=$null }
    }
    $matches = @($CloudRows | Where-Object {
        $candidate = [guid]::Empty
        $tenant = [guid]::Empty
        [guid]::TryParse([string](Get-SemtPropertyValue -InputObject $_ -Name 'AzureADDeviceId' -Default ''), [ref]$candidate) -and $candidate -eq $localId -and
        [guid]::TryParse([string](Get-SemtPropertyValue -InputObject $_ -Name 'InventoryTenantId' -Default ''), [ref]$tenant) -and $tenant -eq $localTenant
    })
    if ($matches.Count -eq 0) { return [pscustomobject]@{ State='NotFound'; Detail='No matching Entra device ID in scoped Intune inventory'; Present=$false } }
    if ($matches.Count -ne 1) { return [pscustomobject]@{ State='Ambiguous'; Detail='Multiple Intune objects match the local Entra device ID'; Present=$null } }
    if (-not (Test-SemtBooleanLikeTrue -Value (Get-SemtPropertyValue -InputObject $Row -Name 'IntuneEnrolled' -Default ''))) {
        return [pscustomobject]@{ State='CloudOnly'; Detail='Intune identity matched; local enrollment not confirmed'; Present=$true }
    }
    return [pscustomobject]@{ State='Confirmed'; Detail='Local enrollment and one matching Intune Entra device ID in the scoped export'; Present=$true }
}

function Get-HybridJoinCentralLogBucket {
    param([Parameter(Mandatory = $true)]$Result)

    $status = ([string](Get-SemtPropertyValue -InputObject $Result -Name 'Status' -Default '')).ToUpperInvariant()
    if ($status -in @('SUCCESS','AUDIT_SUCCESS_ALREADY_INTUNE','ENROLLED_DETECTED_POST_CYCLE','ENROLLED_CONFIRMED_IN_CYCLE','ENROLLED_CONFIRMED_BY_CLOUD_IDENTITY')) { return 'Success' }
    if ($status -like '*ADMIN_SHARE_UNREACHABLE') { return 'AdminShareUnreachable' }
    if ($status -like 'PSEXEC_TIMEOUT*') { return 'PsExecTimeout' }
    if ($status -eq 'PSEXEC_COMMUNICATION_LOST') { return 'PsExecCommunicationLost' }
    if ($status -like 'SKIPPED_*') { return 'Skipped' }
    if ($status -like 'TARGET_*' -or $status -like 'LEGACY_*' -or $status -in @('ENDPOINT_RUN_ACTIVE','RUN_GUARD_ACTIVE')) { return 'TargetBusy' }
    if ($status -like 'PSEXEC_*' -or $status -like 'REMOTE_*' -or $status -like 'DNS_*') { return 'RemoteCollectionFailed' }
    return 'Errors'
}

function Get-HybridJoinCooldownHours {
    param([string]$Status,[int]$BackoffCount = 1,[int]$CycleNumber = 1,[int]$DefaultHours = 12)

    $normalized = ([string]$Status).ToUpperInvariant()
    if ($normalized -match '^(CANCELLED_|DRYRUN_|AUDIT_|SKIPPED_VIRTUAL_MACHINE)') { return 0.0 }
    if ($normalized -match '^(ADMIN_SHARE|DNS_|PSEXEC_|REMOTE_.*FAILED|REMOTE_RESULT|REMOTE_TASK_NOT_STARTED|REMOTE_TASK_NOT_REGISTERED|CENTRAL_.*FAILED|JOB_ERROR|ERROR|TARGET_|LEGACY_)') {
        $exponent = [math]::Min(5, [math]::Max(0, [math]::Max($BackoffCount, $CycleNumber) - 1))
        return ([math]::Min(360, 15 * [math]::Pow(2, $exponent)) / 60.0)
    }
    if ($normalized -match '(USER|INTERACTIVE|PRT|LOGON|SESSION_REMOTE|LOCAL_INTERACTIVE)') { return 24.0 }
    if ($normalized -match '(RUN_GUARD|ENDPOINT_RUN_ACTIVE|PENDING_CONFIRMATION|REBOOT_TRIGGERED|WAITING_FOR_AAD_CONNECT|WAITING_POST_LEAVE)') { return 12.0 }
    if ($normalized -in @('ENROLLED_DETECTED_POST_CYCLE','ENROLLED_CONFIRMED_IN_CYCLE','ENROLLED_CONFIRMED_BY_CLOUD_IDENTITY')) { return 168.0 }
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
    param([Parameter(Mandatory = $true)]$Context,[string[]]$Computers)
    # Name-only inventory matches must never suppress inspection of a target.
    return @()
}

function Get-HybridJoinIntuneAbsenceProof {
    <#
    Entra device ID of a computer proven absent from Intune, or '' when the proof is incomplete:
    fresh Entra and Intune inventories, an Entra object for the name with a device ID, no Intune
    record for the name and no Intune record with that device ID.
    #>
    param([Parameter(Mandatory = $true)]$Context,[Parameter(Mandatory = $true)][string]$Computer)

    $entra = $Context.Inventory['Entra']
    $intune = $Context.Inventory['Intune']
    if ($null -eq $entra -or $null -eq $intune -or $null -eq $entra.Map -or $null -eq $intune.Map) { return '' }
    $freshSources = @('Shared','RunScopedRefresh')
    if ([string]$entra.Source -notin $freshSources -or [string]$intune.Source -notin $freshSources) { return '' }
    $key = Get-SemtComputerKey -ComputerName $Computer
    if (-not $entra.Map.ContainsKey($key) -or $intune.Map.ContainsKey($key)) { return '' }
    $deviceId = [guid]::Empty
    if (-not [guid]::TryParse([string](Get-SemtPropertyValue -InputObject $entra.Map[$key] -Name 'DeviceId' -Default ''), [ref]$deviceId) -or $deviceId -eq [guid]::Empty) { return '' }
    foreach ($row in @($intune.Map.Values)) {
        $candidate = [guid]::Empty
        if ([guid]::TryParse([string](Get-SemtPropertyValue -InputObject $row -Name 'AzureADDeviceId' -Default ''), [ref]$candidate) -and $candidate -eq $deviceId) { return '' }
    }
    return $deviceId.ToString()
}

function Get-HybridJoinPerComputerArguments {
    param([Parameter(Mandatory = $true)]$Context,[Parameter(Mandatory = $true)][string]$Computer)

    $arguments = New-Object System.Collections.Generic.List[string]
    $entra = $Context.Inventory['Entra']
    if ($null -eq $entra -or $null -eq $entra.Map) { return @() }
    $key = Get-SemtComputerKey -ComputerName $Computer
    if ($entra.Map.ContainsKey($key) -and [string](Get-SemtPropertyValue -InputObject $entra.Map[$key] -Name 'EntraRegisteredState' -Default '') -eq 'Pending') { $arguments.Add('-EntraHybridPending') }
    $repairOption = 'HybridJoin.RepairUnhealthyJoinWithIntuneTrace'
    if ($null -ne $Context.Options -and $Context.Options.Contains($repairOption) -and [bool](Get-SemtOption -Options $Context.Options -Name $repairOption)) {
        $absentDeviceId = Get-HybridJoinIntuneAbsenceProof -Context $Context -Computer $Computer
        if ($absentDeviceId) { $arguments.Add('-IntuneAbsentDeviceId'); $arguments.Add($absentDeviceId) }
    }
    return @($arguments.ToArray())
}

$script:HybridJoinGraphSignInRequiredDetail = 'Graph sign-in required: the Run has no reusable Microsoft Graph sign-in; refresh the Intune and Entra inventories in the GUI'

function Test-HybridJoinGraphSignInMissing {
    <#
    True when an inventory exporter failed before any Graph account was connected: its log
    exists but never printed the "Account :" line. The hidden Run process cannot show the
    interactive prompt, so the next exporter would fail the same way.
    #>
    param($Export)

    if ($null -eq $Export -or $Export.Success) { return $false }
    $logPath = [string](Get-SemtPropertyValue -InputObject $Export -Name 'LogPath' -Default '')
    if ([string]::IsNullOrWhiteSpace($logPath) -or -not (Test-Path -LiteralPath $logPath -PathType Leaf)) { return $false }
    return -not (Select-String -LiteralPath $logPath -Pattern '\bAccount\s+:' -Quiet)
}

function Invoke-HybridJoinPostCycle {
    <#
    Confirms results with Intune and Entra inventories scoped to computers whose state could
    have changed, then refreshes the AD inventory of the cycle when it is due.
    #>
    param([Parameter(Mandatory = $true)]$Context,[object[]]$Rows)

    $allRows = @($Rows | Where-Object { $null -ne $_ })
    $usersPath = Join-Path $Context.PackageRoot 'Inventory\UsersLicenses.csv'
    $userRows = @()
    $userInventoryValid = $false
    if (Test-Path -LiteralPath $usersPath -PathType Leaf) {
        try {
            $age = ([datetime]::UtcNow - (Get-Item -LiteralPath $usersPath).LastWriteTimeUtc).TotalHours
            if ($age -ge 0 -and $age -lt 24) {
                $userRows = @(Import-Csv -LiteralPath $usersPath)
                $userInventoryValid = ($userRows.Count -gt 0)
            }
        } catch { $userInventoryValid = $false }
    }
    foreach ($row in $allRows) {
        $confirmedInCycle = ([string]$row.Status -in $script:HybridJoinInCycleStatuses)
        Set-SemtResultField -Result $row -Name 'UserLicenseState' -Value 'Unknown'
        Set-SemtResultField -Result $row -Name 'UserLicenseDetail' -Value 'Fresh same-tenant inventory or unique user SID unavailable'
        Set-SemtResultField -Result $row -Name 'MdmUserScopeState' -Value 'Unknown: verify MDM user scope in Entra'
        if ([string](Get-SemtPropertyValue -InputObject $row -Name 'UserLicensePreflightState' -Default '') -eq 'NoAssignedLicense') {
            Set-SemtResultField -Result $row -Name 'UserLicenseState' -Value 'NoAssignedLicense'
            Set-SemtResultField -Result $row -Name 'UserLicenseDetail' -Value ([string](Get-SemtPropertyValue -InputObject $row -Name 'UserLicensePreflightDetail' -Default 'No assigned user license'))
        }
        $sid = [string](Get-SemtPropertyValue -InputObject $row -Name 'InteractiveUserSid' -Default '')
        $tenant = [string](Get-SemtPropertyValue -InputObject $row -Name 'Dsreg_TenantId' -Default '')
        if ($userInventoryValid -and $sid -and $tenant) {
            $matches = @($userRows | Where-Object { $_.OnPremisesSecurityIdentifier -eq $sid -and $_.InventoryTenantId -eq $tenant -and $_.InventoryComplete -eq 'True' -and $_.InventoryScope -eq 'AllUsers' })
            if ($matches.Count -eq 1) {
                $u = $matches[0]
                Set-SemtResultField -Result $row -Name 'UserLicenseUpn' -Value $u.UserPrincipalName
                Set-SemtResultField -Result $row -Name 'UserLicenseInventoryUtc' -Value $u.CollectedUtc
                # INTUNE_A and INTUNE_EDU are documented service plan IDs. Other plans
                # remain unknown rather than being classified as unlicensed.
                $plans = @($u.AssignedPlans -split ';')
                $enabled = @($plans | Where-Object { $_ -match '^(c1ec4a95-1f05-45b3-a911-aa3fa01094f5|da24caf9-af8e-485c-b7c8-e73336da2693):Enabled$' })
                if ($enabled.Count -gt 0) { $state='Enabled'; $detail='Known Intune user plan enabled' }
                elseif (-not $u.AssignedSkuIds) { $state='NoAssignedLicense'; $detail='No user SKU in completed inventory; validate current license and MDM scope' }
                else { $state='Unknown'; $detail='No known enabled Intune plan; inspect assigned plans and license details' }
                Set-SemtResultField -Result $row -Name 'UserLicenseState' -Value $state
                Set-SemtResultField -Result $row -Name 'UserLicenseDetail' -Value $detail
            }
        }
        # Confirmed and archived during the cycle: the post-cycle confirmation keeps that proof.
        if ($confirmedInCycle) { continue }
        Set-SemtResultField -Result $row -Name 'IntuneConfirmationState' -Value 'Unknown'
        Set-SemtResultField -Result $row -Name 'IntuneConfirmationDetail' -Value 'Cloud identity confirmation not completed'
        if ([string]$row.Status -in @('SUCCESS','AUDIT_SUCCESS_ALREADY_INTUNE') -and [string](Get-SemtPropertyValue -InputObject $row -Name 'UserLicensePreflightState' -Default '') -ne 'NoAssignedLicense') {
            Set-SemtResultField -Result $row -Name 'StatusBeforePostCycle' -Value ([string]$row.Status)
            Set-SemtResultField -Result $row -Name 'Status' -Value 'INTUNE_CLOUD_CONFIRMATION_PENDING'
            Set-SemtResultField -Result $row -Name 'NextAction' -Value 'VERIFY_INTUNE_IDENTITY'
        }
    }
    if ($Context.DryRun -or $allRows.Count -eq 0) { return }
    $cloudRows = @($allRows | Where-Object { [string](Get-SemtPropertyValue -InputObject $_ -Name 'UserLicensePreflightState' -Default '') -ne 'NoAssignedLicense' -and [string]$_.Status -notin $script:HybridJoinInCycleStatuses -and [string]$_.Status -notmatch '^(USER_LICENSE_REVIEW_REQUIRED|ADMIN_SHARE_UNREACHABLE|DNS_PREFLIGHT_ALL_SAMPLES_FAILED|REMOTE_PAYLOAD_COPY_FAILED|REMOTE_SCRIPT_MISSING|CANCELLED_|SKIPPED_|DRYRUN_|TARGET_|LEGACY_)' })
    $log = $Context.Log
    if ($cloudRows.Count -eq 0) { & $log ('Cycle {0}: post-cycle Graph confirmation skipped: no result could have changed Intune or Entra.' -f $Context.CycleNumber) 'DarkGray'; return }
    if (-not (Get-SemtOption -Options $Context.Options -Name 'HybridJoin.PostCycleCloudInventoryRefresh')) { & $log ('Cycle {0}: post-cycle Graph confirmation disabled by option.' -f $Context.CycleNumber) 'DarkGray'; return }
    if ((Get-SemtCancellationState).Requested) { return }
    $scopePath = Join-Path -Path $Context.RunContext.InventoryRoot -ChildPath ('Cycle{0}_CloudRefreshScope.txt' -f $Context.CycleNumber)
    Write-SemtTextLinesAtomic -Path $scopePath -Lines @($cloudRows | ForEach-Object { [string]$_.ComputerName } | Select-Object -Unique)
    $scoped = @($cloudRows | ForEach-Object { [string]$_.ComputerName })
    # Local Entra device IDs: a device whose cloud name differs from the computer name is still found.
    $deviceIdLines = New-Object System.Collections.Generic.List[string]
    $deviceIdLines.Add('ComputerName,DeviceId')
    foreach ($row in $cloudRows) {
        $localId = [guid]::Empty
        if ([guid]::TryParse([string](Get-SemtPropertyValue -InputObject $row -Name 'Dsreg_DeviceId' -Default ''), [ref]$localId) -and $localId -ne [guid]::Empty) {
            $deviceIdLines.Add(('{0},{1}' -f [string]$row.ComputerName,$localId))
        }
    }
    $deviceIdPath = ''
    if ($deviceIdLines.Count -gt 1) {
        $deviceIdPath = Join-Path -Path $Context.RunContext.InventoryRoot -ChildPath ('Cycle{0}_CloudRefreshDeviceIds.csv' -f $Context.CycleNumber)
        Write-SemtTextLinesAtomic -Path $deviceIdPath -Lines @($deviceIdLines.ToArray())
    }
    $keys = @{}
    foreach ($row in $cloudRows) { $keys[(Get-SemtComputerKey -ComputerName ([string]$row.ComputerName))] = $true }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'

    $intuneWatch = [System.Diagnostics.Stopwatch]::StartNew()
    $intuneExport =Invoke-SemtInventoryExport -Kind Intune -ExportScriptPath $Context.ExportScripts.Intune -OutputPath (Join-Path $Context.RunContext.InventoryRoot ("DevicesIntune_Cycle{0}PostCycle_{1}.csv" -f $Context.CycleNumber,$stamp)) -LogPath (Join-Path $Context.RunContext.InventoryRoot ("DevicesIntune_Cycle{0}PostCycle_{1}.log" -f $Context.CycleNumber,$stamp)) -ComputerListPath $scopePath -TenantId (Get-SemtOption -Options $Context.Options -Name 'IntuneTenantId') -PageSize (Get-SemtOption -Options $Context.Options -Name 'IntuneInventoryPageSize') -SharedInventoryPath $Context.Inventory['Intune'].SharedPath -DeviceIdListPath $deviceIdPath
    $intuneSeconds = [int][math]::Round($intuneWatch.Elapsed.TotalSeconds)
    $intuneRows = @()
    if ($intuneExport.Success) { $intuneRows = @(Import-Csv -LiteralPath $intuneExport.CsvPath) }
    $signInRequired = Test-HybridJoinGraphSignInMissing -Export $intuneExport
    $signInDetail = '{0} (the Intune export waited {1} s for the hidden sign-in prompt)' -f $script:HybridJoinGraphSignInRequiredDetail,$intuneSeconds
    foreach ($row in $cloudRows) {
        Set-SemtResultField -Result $row -Name 'PostCycleIntuneInventoryChecked' -Value $true
        Set-SemtResultField -Result $row -Name 'PostCycleIntuneInventoryCsv' -Value $intuneExport.CsvPath
        if (-not $intuneExport.Success) {
            Set-SemtResultField -Result $row -Name 'PostCycleIntuneInventoryError' -Value $intuneExport.Error
            Set-SemtResultField -Result $row -Name 'IntuneConfirmationDetail' -Value $(if ($signInRequired) { $signInDetail } else { 'Intune inventory query failed' })
            continue
        }
        $confirmation = Get-HybridJoinIntuneConfirmation -Row $row -CloudRows $intuneRows
        Set-SemtResultField -Result $row -Name 'IntuneConfirmationState' -Value $confirmation.State
        Set-SemtResultField -Result $row -Name 'IntuneConfirmationDetail' -Value $confirmation.Detail
        if ($null -ne $confirmation.Present) { Set-SemtResultField -Result $row -Name 'PostCycleIntuneInventoryPresent' -Value $confirmation.Present }
        Set-SemtResultField -Result $row -Name 'PostCycleIntuneEnrollmentDetected' -Value ($confirmation.State -eq 'Confirmed')
        if ($confirmation.State -eq 'Confirmed' -and [string]$row.Status -eq 'INTUNE_CLOUD_CONFIRMATION_PENDING') {
            if ([string]$row.Status -ne 'INTUNE_CLOUD_CONFIRMATION_PENDING') { Set-SemtResultField -Result $row -Name 'StatusBeforePostCycle' -Value ([string]$row.Status) }
            Set-SemtResultField -Result $row -Name 'Status' -Value 'ENROLLED_DETECTED_POST_CYCLE'
            Set-SemtResultField -Result $row -Name 'NextAction' -Value 'NO_ACTION_ALREADY_INTUNE'
        }
    }
    if ($intuneExport.Success) {
        $Context.Inventory['Intune'].Map = Merge-SemtScopedInventoryMap -ExistingMap $Context.Inventory['Intune'].Map -RefreshedMap $intuneExport.Map -ScopedComputers $scoped
        & $log ('Cycle {0}: post-cycle Intune confirmation: {1}/{2} present; CSV={3}' -f $Context.CycleNumber,@($cloudRows | Where-Object { $_.PostCycleIntuneInventoryPresent -eq $true }).Count,$cloudRows.Count,$intuneExport.CsvPath) 'Green'
    }
    else { & $log ('Cycle {0}: post-cycle Intune confirmation failed: {1}' -f $Context.CycleNumber,$intuneExport.Error) 'Yellow' }
    if ($signInRequired) {
        # The Entra export would wait for the same hidden sign-in prompt: it is not started.
        foreach ($row in $cloudRows) { Set-SemtResultField -Result $row -Name 'PostCycleEntraInventoryError' -Value 'Not started: Graph sign-in required' }
        & $log ('Cycle {0}: GRAPH_SIGN_IN_REQUIRED: the Run has no reusable Microsoft Graph sign-in (the Intune export waited {2} s for the hidden sign-in prompt), so {1} result(s) stay INTUNE_CLOUD_CONFIRMATION_PENDING and the Entra confirmation was not started. Refresh the Intune and Entra inventories in the GUI, then regenerate the Preview.' -f $Context.CycleNumber,$cloudRows.Count,$intuneSeconds) 'Yellow'
        return
    }

    if ((Get-SemtCancellationState).Requested) { return }
    $entraExport = Invoke-SemtInventoryExport -Kind Entra -ExportScriptPath $Context.ExportScripts.Entra -OutputPath (Join-Path $Context.RunContext.InventoryRoot ("DevicesEntra_Cycle{0}PostCycle_{1}.csv" -f $Context.CycleNumber,$stamp)) -LogPath (Join-Path $Context.RunContext.InventoryRoot ("DevicesEntra_Cycle{0}PostCycle_{1}.log" -f $Context.CycleNumber,$stamp)) -ComputerListPath $scopePath -TenantId (Get-SemtOption -Options $Context.Options -Name 'IntuneTenantId') -PageSize (Get-SemtOption -Options $Context.Options -Name 'IntuneInventoryPageSize') -SharedInventoryPath $Context.Inventory['Entra'].SharedPath -DeviceIdListPath $deviceIdPath
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
        if ($resolved -and [string]$row.Status -in @('INTUNE_CLOUD_CONFIRMATION_PENDING','ENTRA_HYBRID_PENDING')) {
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

#region In-cycle Intune confirmation and archival

# Results confirmed during the cycle are archived at once (Wave and Campaign lists) and are left
# untouched by the post-cycle confirmation.
$script:HybridJoinInCycleStatuses = @('ENROLLED_CONFIRMED_IN_CYCLE','ENROLLED_CONFIRMED_BY_CLOUD_IDENTITY')
$script:HybridJoinFreshInventorySources = @('Shared','RunScopedRefresh')
$script:HybridJoinIndexCache = @{}
$script:HybridJoinInterim = $null
$script:HybridJoinInterimTimeoutMinutes = 20

function Get-HybridJoinContextOption {
    param($Context,[Parameter(Mandatory = $true)][string]$Name,$Default)
    if ($null -ne $Context.Options -and $Context.Options.Contains($Name)) { return (Get-SemtOption -Options $Context.Options -Name $Name) }
    return $Default
}

function New-HybridJoinDeviceIdIndex {
    <#
    Index of inventory rows by Entra device ID (Intune: AzureADDeviceId; Entra: DeviceId). TenantId
    is set only when every indexed row comes from the same tenant.
    #>
    param([object[]]$Rows,[Parameter(Mandatory = $true)][ValidateSet('Intune','Entra')][string]$Kind)
    $idColumn = if ($Kind -eq 'Intune') { 'AzureADDeviceId' } else { 'DeviceId' }
    $presence = '{0}InventoryPresent' -f $Kind
    $byId = @{}
    $tenants = @{}
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        if (-not (Test-SemtBooleanLikeTrue -Value (Get-SemtPropertyValue -InputObject $row -Name $presence -Default 'True'))) { continue }
        $id = [guid]::Empty
        if (-not [guid]::TryParse([string](Get-SemtPropertyValue -InputObject $row -Name $idColumn -Default ''), [ref]$id) -or $id -eq [guid]::Empty) { continue }
        $tenants[([string](Get-SemtPropertyValue -InputObject $row -Name 'InventoryTenantId' -Default '')).Trim().ToLowerInvariant()] = $true
        $key = $id.ToString()
        if (-not $byId.ContainsKey($key)) { $byId[$key] = New-Object System.Collections.Generic.List[object] }
        $byId[$key].Add($row)
    }
    $tenantList = @($tenants.Keys)
    return [pscustomobject]@{ Kind = $Kind; ById = $byId; TenantId = $(if ($tenantList.Count -eq 1) { [string]$tenantList[0] } else { '' }) }
}

function Get-HybridJoinInventoryIndex {
    <# Device ID index of the cycle inventory, built once per CSV; $null unless the inventory is fresh. #>
    param([Parameter(Mandatory = $true)]$Context,[Parameter(Mandatory = $true)][ValidateSet('Intune','Entra')][string]$Kind)
    if ($null -eq $Context.Inventory -or -not $Context.Inventory.ContainsKey($Kind)) { return $null }
    $state = $Context.Inventory[$Kind]
    if ($null -eq $state -or [string]$state.Source -notin $script:HybridJoinFreshInventorySources) { return $null }
    $csv = [string]$state.CsvPath
    if ([string]::IsNullOrWhiteSpace($csv) -or -not (Test-Path -LiteralPath $csv -PathType Leaf)) { return $null }
    $cacheKey = '{0}|{1}|{2}' -f $Kind,$csv,(Get-Item -LiteralPath $csv).LastWriteTimeUtc.Ticks
    if (-not $script:HybridJoinIndexCache.ContainsKey($cacheKey)) {
        foreach ($old in @($script:HybridJoinIndexCache.Keys | Where-Object { $_ -like ('{0}|*' -f $Kind) })) { $script:HybridJoinIndexCache.Remove($old) }
        $script:HybridJoinIndexCache[$cacheKey] = New-HybridJoinDeviceIdIndex -Rows @(Import-Csv -LiteralPath $csv) -Kind $Kind
    }
    return $script:HybridJoinIndexCache[$cacheKey]
}

function Test-HybridJoinRecentSync {
    param($IntuneRow,[int]$MaxSyncAgeDays,[datetime]$NowUtc)
    $sync = ConvertTo-SemtUtcDateTime -Value (Get-SemtPropertyValue -InputObject $IntuneRow -Name 'LastSyncDateTime' -Default '')
    return ($null -ne $sync -and ($NowUtc - $sync).TotalDays -le $MaxSyncAgeDays)
}

function Get-HybridJoinIndexedIntuneConfirmation {
    <#
    Same proof as the post-cycle confirmation (local enrollment and exactly one Intune object with
    the local Entra device ID in the same tenant), plus an Intune sync within MaxSyncAgeDays.
    #>
    param([Parameter(Mandatory = $true)]$Row,$Index,[int]$MaxSyncAgeDays = 30,[datetime]$NowUtc = (Get-Date).ToUniversalTime())
    if ($null -eq $Index) { return [pscustomobject]@{ State = 'Unknown'; Detail = 'Fresh Intune inventory unavailable' } }
    $localId = [guid]::Empty
    if (-not [guid]::TryParse([string](Get-SemtPropertyValue -InputObject $Row -Name 'Dsreg_DeviceId' -Default ''), [ref]$localId) -or $localId -eq [guid]::Empty) { return [pscustomobject]@{ State = 'Unknown'; Detail = 'Local Entra device ID unavailable' } }
    $localTenant = [guid]::Empty
    if (-not [guid]::TryParse([string](Get-SemtPropertyValue -InputObject $Row -Name 'Dsreg_TenantId' -Default ''), [ref]$localTenant) -or $localTenant -eq [guid]::Empty) { return [pscustomobject]@{ State = 'Unknown'; Detail = 'Local tenant ID unavailable' } }
    if ([string]::IsNullOrWhiteSpace($Index.TenantId) -or $Index.TenantId -ne $localTenant.ToString()) { return [pscustomobject]@{ State = 'Unknown'; Detail = 'Intune inventory tenant unavailable or mismatched' } }
    $key = $localId.ToString()
    if (-not $Index.ById.ContainsKey($key)) { return [pscustomobject]@{ State = 'NotFound'; Detail = 'No matching Entra device ID in the Intune inventory' } }
    $found = $Index.ById[$key]
    if ($found.Count -ne 1) { return [pscustomobject]@{ State = 'Ambiguous'; Detail = 'Multiple Intune objects match the local Entra device ID' } }
    if (-not (Test-HybridJoinRecentSync -IntuneRow $found[0] -MaxSyncAgeDays $MaxSyncAgeDays -NowUtc $NowUtc)) { return [pscustomobject]@{ State = 'SyncTooOld'; Detail = ('Intune object found, but its last sync is older than {0} day(s)' -f $MaxSyncAgeDays) } }
    if (-not (Test-SemtBooleanLikeTrue -Value (Get-SemtPropertyValue -InputObject $Row -Name 'IntuneEnrolled' -Default ''))) { return [pscustomobject]@{ State = 'CloudOnly'; Detail = 'Intune identity matched; local enrollment not confirmed' } }
    return [pscustomobject]@{ State = 'Confirmed'; Detail = 'Local enrollment and one matching Intune Entra device ID with a recent sync' }
}

function Get-HybridJoinCloudIdentityProof {
    <#
    Enrollment proof for a computer the Run could not reach, without relying on names in the
    cloud: AD ObjectGUID = DeviceId of one enabled hybrid (ServerAd) Entra object = AzureADDeviceId
    of exactly one Intune MDM object of the same tenant, synced within MaxSyncAgeDays. The computer
    name is only used to read its own AD object.
    #>
    param([Parameter(Mandatory = $true)]$Context,[Parameter(Mandatory = $true)]$Row,[int]$MaxSyncAgeDays = 30,[datetime]$NowUtc = (Get-Date).ToUniversalTime())
    $no = { param([string]$Why) [pscustomobject]@{ Confirmed = $false; DeviceId = ''; Detail = $Why } }
    $ad = if ($null -ne $Context.Inventory -and $Context.Inventory.ContainsKey('AD')) { $Context.Inventory['AD'] } else { $null }
    if ($null -eq $ad -or [string]$ad.Source -notin $script:HybridJoinFreshInventorySources -or $null -eq $ad.Map) { return (& $no 'Fresh AD inventory unavailable') }
    $key = Get-SemtComputerKey -ComputerName ([string]$Row.ComputerName)
    if (-not $ad.Map.ContainsKey($key)) { return (& $no 'Computer not found in the AD inventory') }
    $guid = [guid]::Empty
    if (-not [guid]::TryParse([string](Get-SemtPropertyValue -InputObject $ad.Map[$key] -Name 'ObjectGUID' -Default ''), [ref]$guid) -or $guid -eq [guid]::Empty) { return (& $no 'AD ObjectGUID unavailable') }
    $entraIndex = Get-HybridJoinInventoryIndex -Context $Context -Kind Entra
    $intuneIndex = Get-HybridJoinInventoryIndex -Context $Context -Kind Intune
    if ($null -eq $entraIndex -or $null -eq $intuneIndex) { return (& $no 'Fresh Entra or Intune inventory unavailable') }
    $id = $guid.ToString()
    if (-not $entraIndex.ById.ContainsKey($id) -or $entraIndex.ById[$id].Count -ne 1) { return (& $no 'No single Entra device with the AD ObjectGUID as device ID') }
    $entra = $entraIndex.ById[$id][0]
    if ([string](Get-SemtPropertyValue -InputObject $entra -Name 'TrustType' -Default '') -ne 'ServerAd') { return (& $no 'Entra device is not hybrid joined (TrustType is not ServerAd)') }
    if (-not (Test-SemtBooleanLikeTrue -Value (Get-SemtPropertyValue -InputObject $entra -Name 'AccountEnabled' -Default ''))) { return (& $no 'Entra device is disabled') }
    if ([string]::IsNullOrWhiteSpace($entraIndex.TenantId) -or $entraIndex.TenantId -ne $intuneIndex.TenantId) { return (& $no 'Entra and Intune inventory tenants unavailable or mismatched') }
    if (-not $intuneIndex.ById.ContainsKey($id) -or $intuneIndex.ById[$id].Count -ne 1) { return (& $no 'No single Intune object with the Entra device ID') }
    $intune = $intuneIndex.ById[$id][0]
    if ([string](Get-SemtPropertyValue -InputObject $intune -Name 'ManagementAgent' -Default '') -notmatch 'mdm') { return (& $no 'Intune object is not MDM managed') }
    if (-not (Test-HybridJoinRecentSync -IntuneRow $intune -MaxSyncAgeDays $MaxSyncAgeDays -NowUtc $NowUtc)) { return (& $no ('Intune last sync older than {0} day(s)' -f $MaxSyncAgeDays)) }
    return [pscustomobject]@{ Confirmed = $true; DeviceId = $id; Detail = ('AD ObjectGUID = hybrid Entra device ID = one Intune MDM object synced within {0} day(s); DeviceId={1}; LastSync={2}' -f $MaxSyncAgeDays,$id,[string](Get-SemtPropertyValue -InputObject $intune -Name 'LastSyncDateTime' -Default '')) }
}

function Set-HybridJoinInCycleConfirmed {
    param([Parameter(Mandatory = $true)]$Row,[Parameter(Mandatory = $true)][string]$Status,[Parameter(Mandatory = $true)][string]$ConfirmationState,[Parameter(Mandatory = $true)][string]$Source,[string]$Detail)
    Set-SemtResultField -Result $Row -Name 'StatusBeforePostCycle' -Value ([string]$Row.Status)
    Set-SemtResultField -Result $Row -Name 'IntuneConfirmationState' -Value $ConfirmationState
    Set-SemtResultField -Result $Row -Name 'IntuneConfirmationDetail' -Value $Detail
    Set-SemtResultField -Result $Row -Name 'IntuneConfirmationSource' -Value $Source
    Set-SemtResultField -Result $Row -Name 'Status' -Value $Status
    Set-SemtResultField -Result $Row -Name 'NextAction' -Value 'NO_ACTION_ALREADY_INTUNE'
}

function Get-HybridJoinInterimState {
    param([Parameter(Mandatory = $true)]$Context)
    $runKey = [string](Get-SemtPropertyValue -InputObject $Context.RunContext -Name 'InventoryRoot' -Default '')
    if ($null -eq $script:HybridJoinInterim -or $script:HybridJoinInterim.RunKey -ne $runKey) {
        $script:HybridJoinInterim = [pscustomobject]@{ RunKey = $runKey; Pending = @{}; Job = $null; JobKeys = @(); JobStartedUtc = $null; LastStartUtc = (Get-Date).ToUniversalTime(); Disabled = $false; Sequence = 0 }
    }
    return $script:HybridJoinInterim
}

function Invoke-HybridJoinCompletedRow {
    <#
    Runs for each result as soon as its worker is collected, before the run guard update and the
    archive rules. A confirmed result gets an in-cycle status that the archive rule accepts.
    A local success not yet visible in the cycle inventory waits for the interim confirmation.
    #>
    param([Parameter(Mandatory = $true)]$Context,$Row)
    if ($null -eq $Row -or $Context.DryRun) { return }
    if ([string](Get-SemtPropertyValue -InputObject $Row -Name 'UserLicensePreflightState' -Default '') -eq 'NoAssignedLicense') { return }
    $status = [string]$Row.Status
    $maxAge = [int](Get-HybridJoinContextOption -Context $Context -Name 'HybridJoin.CloudConfirmationMaxSyncAgeDays' -Default 30)
    if ($status -in @('SUCCESS','AUDIT_SUCCESS_ALREADY_INTUNE')) {
        if (-not (Test-SemtBooleanLikeTrue -Value (Get-SemtPropertyValue -InputObject $Row -Name 'IntuneEnrolled' -Default ''))) { return }
        $confirmation = Get-HybridJoinIndexedIntuneConfirmation -Row $Row -Index (Get-HybridJoinInventoryIndex -Context $Context -Kind Intune) -MaxSyncAgeDays $maxAge
        if ($confirmation.State -eq 'Confirmed') {
            Set-HybridJoinInCycleConfirmed -Row $Row -Status 'ENROLLED_CONFIRMED_IN_CYCLE' -ConfirmationState 'Confirmed' -Source 'CycleInventory' -Detail $confirmation.Detail
            return
        }
        $state = Get-HybridJoinInterimState -Context $Context
        if (-not $state.Disabled) { $state.Pending[(Get-SemtComputerKey -ComputerName ([string]$Row.ComputerName))] = $Row }
        return
    }
    if ($status -like '*ADMIN_SHARE_UNREACHABLE' -and [bool](Get-HybridJoinContextOption -Context $Context -Name 'HybridJoin.ArchiveUnreachableWithCloudIdentity' -Default $true)) {
        $proof = Get-HybridJoinCloudIdentityProof -Context $Context -Row $Row -MaxSyncAgeDays $maxAge
        if ($proof.Confirmed) { Set-HybridJoinInCycleConfirmed -Row $Row -Status 'ENROLLED_CONFIRMED_BY_CLOUD_IDENTITY' -ConfirmationState 'CloudIdentityConfirmed' -Source 'CloudIdentityChain' -Detail $proof.Detail }
    }
}

function Complete-HybridJoinInterimExport {
    param([Parameter(Mandatory = $true)]$Context,[Parameter(Mandatory = $true)]$State,$Export)
    $log = $Context.Log
    $confirmed = New-Object System.Collections.Generic.List[object]
    if ($null -eq $Export -or -not [bool](Get-SemtPropertyValue -InputObject $Export -Name 'Success' -Default $false)) {
        if (Test-HybridJoinGraphSignInMissing -Export $Export) {
            $State.Disabled = $true
            $State.Pending.Clear()
            & $log ('WARNING: interim Intune confirmation disabled for the rest of the Run: {0}. Results are still confirmed after each cycle.' -f $script:HybridJoinGraphSignInRequiredDetail) 'Yellow'
        }
        else { & $log ('WARNING: interim Intune confirmation failed; next attempt at the next interval: {0}' -f [string](Get-SemtPropertyValue -InputObject $Export -Name 'Error' -Default 'no export result')) 'Yellow' }
        return @()
    }
    $maxAge = [int](Get-HybridJoinContextOption -Context $Context -Name 'HybridJoin.CloudConfirmationMaxSyncAgeDays' -Default 30)
    $index = New-HybridJoinDeviceIdIndex -Rows @(Import-Csv -LiteralPath ([string]$Export.CsvPath)) -Kind Intune
    foreach ($key in @($State.JobKeys)) {
        if (-not $State.Pending.ContainsKey($key)) { continue }
        $row = $State.Pending[$key]
        $confirmation = Get-HybridJoinIndexedIntuneConfirmation -Row $row -Index $index -MaxSyncAgeDays $maxAge
        if ($confirmation.State -ne 'Confirmed') { continue }
        Set-HybridJoinInCycleConfirmed -Row $row -Status 'ENROLLED_CONFIRMED_IN_CYCLE' -ConfirmationState 'Confirmed' -Source 'InterimInventory' -Detail $confirmation.Detail
        $State.Pending.Remove($key)
        $confirmed.Add($row)
    }
    & $log ('Cycle {0}: interim Intune confirmation: {1}/{2} confirmed; CSV={3}' -f $Context.CycleNumber,$confirmed.Count,@($State.JobKeys).Count,[string]$Export.CsvPath) 'Green'
    return $confirmed.ToArray()
}

function Invoke-HybridJoinInterimConfirmation {
    <#
    Interim Intune confirmation of local successes, every HybridJoin.InterimCloudConfirmationMinutes.
    The scoped export runs in a background job and never blocks worker slots or dispatch; it is
    bounded in time, skipped when nothing is pending, and disabled for the Run when Graph sign-in
    is missing. Poll returns the rows newly confirmed; Stop discards the job and the pending set.
    #>
    param([Parameter(Mandatory = $true)]$Context,[ValidateSet('Poll','Stop')][string]$Mode = 'Poll')
    $state = Get-HybridJoinInterimState -Context $Context
    if ($Mode -eq 'Stop') {
        if ($null -ne $state.Job) { Stop-Job -Job $state.Job -ErrorAction SilentlyContinue; Remove-Job -Job $state.Job -Force -ErrorAction SilentlyContinue; $state.Job = $null }
        $state.Pending.Clear()
        $state.LastStartUtc = (Get-Date).ToUniversalTime()
        return @()
    }
    $interval = [int](Get-HybridJoinContextOption -Context $Context -Name 'HybridJoin.InterimCloudConfirmationMinutes' -Default 60)
    if ($state.Disabled -or $interval -le 0 -or $Context.DryRun) { return @() }
    $now = (Get-Date).ToUniversalTime()
    if ($null -ne $state.Job) {
        if ($state.Job.State -in @('Running','NotStarted')) {
            if (($now - $state.JobStartedUtc).TotalMinutes -ge $script:HybridJoinInterimTimeoutMinutes) {
                Stop-Job -Job $state.Job -ErrorAction SilentlyContinue
                Remove-Job -Job $state.Job -Force -ErrorAction SilentlyContinue
                $state.Job = $null
                & $Context.Log ('WARNING: interim Intune confirmation stopped after {0} minute(s); next attempt at the next interval.' -f $script:HybridJoinInterimTimeoutMinutes) 'Yellow'
            }
            return @()
        }
        $output = @(Receive-Job -Job $state.Job -ErrorAction SilentlyContinue)
        Remove-Job -Job $state.Job -Force -ErrorAction SilentlyContinue
        $state.Job = $null
        $export = if ($output.Count -gt 0) { $output[$output.Count - 1] } else { $null }
        return @(Complete-HybridJoinInterimExport -Context $Context -State $state -Export $export)
    }
    if ($state.Pending.Count -eq 0 -or ($now - $state.LastStartUtc).TotalMinutes -lt $interval) { return @() }
    $keys = @($state.Pending.Keys)
    $state.Sequence++
    $stamp = '{0}_{1:D3}' -f (Get-Date -Format 'yyyyMMdd_HHmmss'),$state.Sequence
    $root = [string]$Context.RunContext.InventoryRoot
    $scopePath = Join-Path -Path $root -ChildPath ('Cycle{0}_InterimScope_{1}.txt' -f $Context.CycleNumber,$stamp)
    Write-SemtTextLinesAtomic -Path $scopePath -Lines @($keys | ForEach-Object { [string]$state.Pending[$_].ComputerName })
    $idLines = New-Object System.Collections.Generic.List[string]
    $idLines.Add('ComputerName,DeviceId')
    foreach ($key in $keys) { $idLines.Add(('{0},{1}' -f [string]$state.Pending[$key].ComputerName,[string](Get-SemtPropertyValue -InputObject $state.Pending[$key] -Name 'Dsreg_DeviceId' -Default ''))) }
    $idPath = Join-Path -Path $root -ChildPath ('Cycle{0}_InterimDeviceIds_{1}.csv' -f $Context.CycleNumber,$stamp)
    Write-SemtTextLinesAtomic -Path $idPath -Lines @($idLines.ToArray())
    $exportArguments = @{
        Kind = 'Intune'
        ExportScriptPath = [string]$Context.ExportScripts.Intune
        OutputPath = (Join-Path -Path $root -ChildPath ('DevicesIntune_Cycle{0}Interim_{1}.csv' -f $Context.CycleNumber,$stamp))
        LogPath = (Join-Path -Path $root -ChildPath ('DevicesIntune_Cycle{0}Interim_{1}.log' -f $Context.CycleNumber,$stamp))
        ComputerListPath = $scopePath
        DeviceIdListPath = $idPath
        TenantId = [string](Get-HybridJoinContextOption -Context $Context -Name 'IntuneTenantId' -Default '')
        PageSize = [int](Get-HybridJoinContextOption -Context $Context -Name 'IntuneInventoryPageSize' -Default 999)
        SharedInventoryPath = [string]$Context.Inventory['Intune'].SharedPath
    }
    $modulePath = Join-Path -Path ([string]$Context.PackageRoot) -ChildPath 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1'
    $state.Job = Start-Job -ArgumentList $modulePath,$exportArguments -ScriptBlock {
        param($ModulePath,$ExportArguments)
        Import-Module $ModulePath -MinimumVersion 1.0.0 -Force
        Invoke-SemtInventoryExport @ExportArguments
    }
    $state.JobKeys = $keys
    $state.JobStartedUtc = $now
    $state.LastStartUtc = $now
    & $Context.Log ('Cycle {0}: interim Intune confirmation started in the background for {1} pending result(s).' -f $Context.CycleNumber,$keys.Count) 'DarkGray'
    return @()
}

#endregion

function Get-SemtHybridJoinProfile {
    [CmdletBinding()]
    param()

    $coreColumns = @('Timestamp','CycleNumber','ComputerName','ConnectionTarget','RemoteExecutionMode','Status','NextAction','LauncherStatus','RemoteStatus','RemoteNextAction','ExitCode','Detail','StatusBeforePostCycle','BackoffStatus','BackoffUntilUtc','BackoffCount','DnsResolved','DnsAddressList','PingReachable','SmbPort445Reachable','AdminShareReachable','AdminShareFailureType','RemotePayloadCopyAttempts','LocalScriptVersion','LocalScriptSha256','RemoteScriptSha256','PsExecExitCode','TargetGuardStatus')
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
        CentralLogBuckets = @('Success','Skipped','AdminShareUnreachable','RemoteCollectionFailed','PsExecTimeout','PsExecCommunicationLost','TargetBusy','Errors')
        EvidenceRules = @(
            [pscustomobject]@{ Folder = 'Output'; Filter = '*'; Recurse = $true; CurrentAttemptOnly = $true; MaxFiles = 40; MaxBytes = 5MB; TargetFolder = 'Output' },
            [pscustomobject]@{ Folder = 'Output'; Filter = 'MDMDiag_*.zip'; Recurse = $false; CurrentAttemptOnly = $true; MaxFiles = 1; MaxBytes = 100MB; TargetFolder = 'Output' },
            [pscustomobject]@{ Folder = 'State'; Filter = '*.json.txt'; Recurse = $false; CurrentAttemptOnly = $false; Names = @('EndpointInstance.json.txt','RebootSafety.json.txt','RetryAfterReboot.json.txt','IntuneTraceRepair.json.txt'); MaxFiles = 5; MaxBytes = 1MB; TargetFolder = 'State' }
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
            CompletedRow = ${function:Invoke-HybridJoinCompletedRow}
            InterimConfirmation = ${function:Invoke-HybridJoinInterimConfirmation}
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
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBKPUBDO3jZYTwO
# 0tUS5ScZFgh2yscXnv+yH0DSdxIWu6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIN0aLbT2UK4jned3AHwwGTw9/JKYkJFxxAT1F4CWFDI3MA0GCSqG
# SIb3DQEBAQUABIIBgIrpq7+X+de1y0SS+Xie7uKkfCewQ+Mm0Bsp1vptTGLFerHM
# YLlR+oHjKjfkiZD9uYfZ3zFELS1ntlwHSkR8isZK7KiU/MFi8e/sjHaafGgCVs9v
# uZygiKorl2e16IMGxFNyy5hpPiM9/G0HW4SfZABVhJaQn3Vi+djeTRyF7souGCFH
# XuRJYAosEgEw2HWu1prIdBJO5IzduaP3aqYtDeDtu0ik2TB/izRuCE8nA43owvQF
# ye9MZ3jjzzDII2yOVOuQJE7SB71OFTIGB1m3zd18Ez005dHkmvlMKAO6QQVKFtZx
# oAItv1uNUB+w33yU+fBTISoILNwzodmXJQBSsmnG8wkFQ8LDIq3zCLtEHv3+FFRo
# AZDW3/juGNpc1WlROPMnWIpU7IxULUQKUHXkrSMmXV3ab1pl50uZtaJt4TPgGPjS
# U0UnW+lfPtK3G+8yC7kXo1S4zF4PiuGjP67MfOxzECOVpt4xytBXpEzvhaex8kr/
# u7K4iuC2Qte7JUnxmaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTMwOTQ1
# MzhaMC8GCSqGSIb3DQEJBDEiBCArVuf/CwT3pfkRRimcOpXkNO7g9/LcV/U8x3j8
# qkzTzzANBgkqhkiG9w0BAQEFAASCAgCljFJ/6qW+fRm9U4pY9rQDnuck6yMQcC8p
# PT4gXwj5blIw1MN4wUthmYa8Sp5SCG7nlm89ytTndBtDEkIR8hQFXx7/NEmJ+Dt2
# +cCT/eoQmsNbsuIIG9ljz+3avGa29BZTVtmSB63+6FqD1DtIgtDpwRnHXulrLVmp
# cOuzl5ns1duBANE15dt2oE9R487WO41pIvZSLpEKwB6vDGmdOIX1ft0ee/cFDo9g
# sTgASCh0tP5TwUPxJ+HsYyeiWJQPN1LvnZ77kbD2DpmhDvREVR8tgQB+2KwkRozE
# e5guFiBFP/l/oh/nIdQC8ijXAMwu7OmzJ1IRDYW/rpvs1PwMZEC7LWlm03rLZE+C
# PziqtDiyw/wMNIyV6aRjVHCLtsHm1Z6uOJz7qK00GJmcbliMc/GuSa3OTVXoI1X5
# 8f7igxHT+KZ8HGyHR72dAoHj5LohExcbaKP6+b8jWT4SG2eQTlVoLJv0KZ055hlm
# sFMKGTqXUP1rUsygCdHqF65/nZRQQ+zDueIw4/tK8C2lfmG2Q6V2lR+oKl51yrB6
# iz0ncDCK0JJHnsL/bNzsZGEKG4P6KVVMOyq9QYI9wWdsHnxUio6DWsTScr/A0ULY
# 6ThstJtjCe4o3kEnsC3/9xbedGWa3b+Q89uQRwO3eFeSBigRM/HLgThg2E5Wx2/5
# fS7+TIjeuA==
# SIG # End signature block
