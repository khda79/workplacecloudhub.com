<#
.SYNOPSIS
    HybridJoin campaign profile of the Smart Endpoint Migration Toolkit.

.DESCRIPTION
    Business rules ported from IntuneHybridJoinToolkit (orchestrator 2.10.78):
    - name-only Intune matches never skip inspection; confirmed identities can be archived;
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
    'InteractiveUserName','InteractiveUserSid','InteractiveUserDomain','InteractiveUserAccountName','InteractiveUserAccountType',
    'InteractiveSessionName','InteractiveSessionState','UserIsUserAzureAD','UserAzureAdPrt','UserSessionIsNotRemote',
    'GpUpdateTimedOut','GpUpdateStopDetail','FinalGpUpdateTimedOut','FinalGpUpdateStopDetail','MdmDiagnosticStatus','MdmDiagnosticZip','MdmEnrollmentRecentEvidence','IntuneEnrolled','Dsreg_DeviceId','Dsreg_TenantId','DsregAzureAdJoined','DsregDeviceAuthStatus','LeaveAttempted','RebootAttempted','RebootReason',
    'RetryAfterRebootAction','RetryAfterRebootDetail','RetryAfterRebootAttempt','RetryAfterRebootMaxAttempts','RetryAfterRebootTaskName'
)
$script:HybridJoinEngineFields = @(
    'StatusBeforePostCycle','BackoffStatus','BackoffUntilUtc','BackoffCount',
    'IntuneConfirmationState','IntuneConfirmationDetail','UserLicensePreflightState','UserLicensePreflightDetail','UserLicenseState','UserLicenseDetail','UserLicenseUpn','UserLicenseInventoryUtc','MdmUserScopeState','PostCycleIntuneInventoryChecked','PostCycleIntuneInventoryPresent','PostCycleIntuneEnrollmentDetected','PostCycleIntuneInventoryCsv','PostCycleIntuneInventoryError',
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
    # Archival requires the same local device identity in a successful cloud result.
    return ([string](Get-SemtPropertyValue -InputObject $Row -Name 'IntuneConfirmationState' -Default '') -eq 'Confirmed' -and [string](Get-SemtPropertyValue -InputObject $Row -Name 'Status' -Default '') -eq 'ENROLLED_DETECTED_POST_CYCLE')
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
    if ($status -in @('SUCCESS','AUDIT_SUCCESS_ALREADY_INTUNE','ENROLLED_DETECTED_POST_CYCLE')) { return 'Success' }
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
    if ($normalized -match '^(ADMIN_SHARE|DNS_|PSEXEC_|REMOTE_.*FAILED|REMOTE_RESULT|CENTRAL_.*FAILED|JOB_ERROR|ERROR|TARGET_|LEGACY_)') {
        $exponent = [math]::Min(5, [math]::Max(0, [math]::Max($BackoffCount, $CycleNumber) - 1))
        return ([math]::Min(360, 15 * [math]::Pow(2, $exponent)) / 60.0)
    }
    if ($normalized -match '(USER|INTERACTIVE|PRT|LOGON|SESSION_REMOTE|LOCAL_INTERACTIVE)') { return 24.0 }
    if ($normalized -match '(RUN_GUARD|ENDPOINT_RUN_ACTIVE|PENDING_CONFIRMATION|REBOOT_TRIGGERED|WAITING_FOR_AAD_CONNECT|WAITING_POST_LEAVE)') { return 12.0 }
    if ($normalized -eq 'ENROLLED_DETECTED_POST_CYCLE') { return 168.0 }
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
        Set-SemtResultField -Result $row -Name 'UserLicenseState' -Value 'Unknown'
        Set-SemtResultField -Result $row -Name 'UserLicenseDetail' -Value 'Fresh same-tenant inventory or unique user SID unavailable'
        Set-SemtResultField -Result $row -Name 'MdmUserScopeState' -Value 'Unknown: verify MDM user scope in Entra'
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
        Set-SemtResultField -Result $row -Name 'IntuneConfirmationState' -Value 'Unknown'
        Set-SemtResultField -Result $row -Name 'IntuneConfirmationDetail' -Value 'Cloud identity confirmation not completed'
        if ([string]$row.Status -in @('SUCCESS','AUDIT_SUCCESS_ALREADY_INTUNE')) {
            Set-SemtResultField -Result $row -Name 'StatusBeforePostCycle' -Value ([string]$row.Status)
            Set-SemtResultField -Result $row -Name 'Status' -Value 'INTUNE_CLOUD_CONFIRMATION_PENDING'
            Set-SemtResultField -Result $row -Name 'NextAction' -Value 'VERIFY_INTUNE_IDENTITY'
        }
    }
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
    $intuneRows = @()
    if ($intuneExport.Success) { $intuneRows = @(Import-Csv -LiteralPath $intuneExport.CsvPath) }
    foreach ($row in $cloudRows) {
        Set-SemtResultField -Result $row -Name 'PostCycleIntuneInventoryChecked' -Value $true
        Set-SemtResultField -Result $row -Name 'PostCycleIntuneInventoryCsv' -Value $intuneExport.CsvPath
        if (-not $intuneExport.Success) {
            Set-SemtResultField -Result $row -Name 'PostCycleIntuneInventoryError' -Value $intuneExport.Error
            Set-SemtResultField -Result $row -Name 'IntuneConfirmationDetail' -Value 'Intune inventory query failed'
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
        CentralLogBuckets = @('Success','Skipped','AdminShareUnreachable','RemoteCollectionFailed','PsExecTimeout','PsExecCommunicationLost','TargetBusy','Errors')
        EvidenceRules = @(
            [pscustomobject]@{ Folder = 'Output'; Filter = '*'; Recurse = $true; CurrentAttemptOnly = $true; MaxFiles = 40; MaxBytes = 5MB; TargetFolder = 'Output' },
            [pscustomobject]@{ Folder = 'Output'; Filter = 'MDMDiag_*.zip'; Recurse = $false; CurrentAttemptOnly = $true; MaxFiles = 1; MaxBytes = 100MB; TargetFolder = 'Output' },
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
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC/3xuKXp545lxW
# WoGkMgYmIJuNEG/U+QbW83z9INaCZKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCASGFXlEsi9iL68IXoUS0Kx
# kGH77e16YeWhBLH4JNs8mDANBgkqhkiG9w0BAQEFAASCAYAS+95lnPSBk8n8USET
# eDl2a0YrK24bcnZvU9/JqRr57ew64gISK52Wd1/FC3aYBDpz/Ie9n/w4WJEApjdt
# Nhmo8NMkNqQu3UCZqMt9bCQUV8BmGoFyaMmjdwl4y2grPuqkX5ThHOyqs+uSpWig
# rpzkEya7arwJJYHgpS5KmbCRdcvfqillg8nODVo4k6HjAUVPdaAIFBydR67NQqd4
# FhMnCrcB9QD1kHB29lkBZJIev8infe45qDd6Y7oZ5eJl3zxRbKF5ekCKYJGtfpmr
# Tgj6iZyBUHUaJpv6dRHneOcZt31WH+rhAwFina4vEz2edfy/tn9FCs3wNvC+xu6X
# AO8HNLhSIRbxYu7/Ut+X+vVW1eWKwU9Qgjq5FmEkQyOnsAb+cGkHqqAhA8ht3oQS
# Y1L4HNw1mxQfqLB+xalqeTYG5K2Uaj1ZVmGmOxJzd8EbqEf6i0jeWEfkLHICiJb2
# 2CB27peENGtBiUA4WU8yf3GLHfgLhABzTV1nF75vVAilYrs=
# SIG # End signature block
