<#
.SYNOPSIS
    Windows11Upgrade campaign profile of the Smart Endpoint Migration Toolkit.

.DESCRIPTION
    Business rules ported from Windows11UpgradeToolkit (orchestrator 1.0.0, worker 0.1.31):
    - computers reported as Windows 11 by AD or Intune are moved before PsExec;
    - Windows 11 and hardware-not-capable results leave Computers.txt;
    - setup source rules (UNC required unless explicitly confirmed, map syntax preflight);
    - technician run guard with failure categories and adaptive retry delays (schema 1);
    - setup monitoring recovery from the endpoint log when PsExec ends before the result;
    - setup source distribution report after each cycle.
    DirectSetupUpgrade never enables reboots: they always require Windows11Upgrade.AllowReboot.
    The module requires SmartEndpointMigration.Core to be imported first.

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

Set-StrictMode -Version 2.0

$script:Windows11ExtensionFields = @(
    'OSCaption','OSBuild','FreeDiskGB','PendingReboot','PendingRebootSource',
    'PolicyGpUpdateTimedOut','PolicyGpUpdateStopDetail','PolicyGpUpdateOutputFile',
    'HardwareReadinessTag','HardwareReadinessCode','HardwareReadinessResult','HardwareReadinessReason','HardwareReadinessLog','HardwareReadinessSource',
    'SetupCacheAction','SetupMediaReuseDecision','SetupDynamicUpdate','SelectedSetupSourcePath','SetupSourceSelectionDetail',
    'DiskCleanupAction','DiskCleanupFreedGB','AdvancedDiskCleanupAction','AdvancedDiskCleanupFreedGB','DismCleanupAction','DismCleanupFreedGB',
    'SetupCompletionRebootAction','SetupCompletionRebootDetail','SetupCompletionRebootUserCount','SetupCompletionRebootUsers',
    'RequiredRebootByUptimeAction','RequiredRebootByUptimeDetail',
    'SetupProfileRepairAction','SetupProfileRepairDetail','SetupProfileRepairBlockingSid','SetupProfileRepairKeptSid','SetupProfileRepairProfilePath','SetupProfileRepairBackupPath',
    'ControlledRebootAction','ControlledRebootDetail','ControlledRebootUserCount','ControlledRebootUsers',
    'RetryAfterRebootAction','RetryAfterRebootDetail','RetryAfterRebootAttempt','RetryAfterRebootMaxAttempts','RetryAfterRebootTaskName',
    'UserRebootNotificationSent','UserRebootNotificationLang','UserRebootNotificationMessage'
)
$script:Windows11EngineFields = @('RetryAfterUtc','FailureCategory','ConsecutiveFailureCount')
$script:Windows11StartedNoResultHours = 4

function Get-Windows11NextAction {
    param([AllowNull()][string]$Status)

    switch ([string]$Status) {
        'ALREADY_WINDOWS11' { return 'NO_ACTION_ALREADY_WINDOWS11' }
        'WINDOWS11_HARDWARE_NOT_CAPABLE' { return 'REVIEW_HARDWARE_READINESS' }
        'SETUP_PROCESS_MONITOR_INTERRUPTED' { return 'CHECK_SETUP_PROCESS_OR_OS_STATUS' }
        default { return '' }
    }
}

function Test-Windows11AlreadyUpgraded {
    param([Parameter(Mandatory = $true)]$Row)
    return ('ALREADY_WINDOWS11' -in @([string](Get-SemtPropertyValue -InputObject $Row -Name 'Status' -Default ''),[string](Get-SemtPropertyValue -InputObject $Row -Name 'RemoteStatus' -Default '')))
}

function Test-Windows11HardwareNotCapable {
    param([Parameter(Mandatory = $true)]$Row)
    return ('WINDOWS11_HARDWARE_NOT_CAPABLE' -in @([string](Get-SemtPropertyValue -InputObject $Row -Name 'Status' -Default ''),[string](Get-SemtPropertyValue -InputObject $Row -Name 'RemoteStatus' -Default '')))
}

function Test-Windows11SetupSourceMapSyntax {
    param([AllowNull()][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return 'NotConfigured' }
    $expanded = [Environment]::ExpandEnvironmentVariables($Path.Trim('"'))
    if (-not (Test-Path -LiteralPath $expanded -PathType Leaf)) { return ('SkippedOperatorCannotRead:{0}' -f $expanded) }
    $rows = @(Import-Csv -LiteralPath $expanded -ErrorAction Stop)
    if ($rows.Count -eq 0) { throw "SEMT_SETUP_SOURCE_MAP_INVALID: setup source map is empty: $expanded" }
    $number = 1
    foreach ($row in $rows) {
        $number++
        $scopeType = [string](Get-SemtPropertyValue -InputObject $row -Name 'ScopeType' -Default (Get-SemtPropertyValue -InputObject $row -Name 'Type' -Default ''))
        $scopeValue = [string](Get-SemtPropertyValue -InputObject $row -Name 'ScopeValue' -Default (Get-SemtPropertyValue -InputObject $row -Name 'Value' -Default ''))
        $source = [string](Get-SemtPropertyValue -InputObject $row -Name 'SetupSourcePath' -Default (Get-SemtPropertyValue -InputObject $row -Name 'SourcePath' -Default ''))
        if ([string]::IsNullOrWhiteSpace($scopeType) -or [string]::IsNullOrWhiteSpace($source)) { throw "SEMT_SETUP_SOURCE_MAP_INVALID: row $number requires ScopeType and SetupSourcePath." }
        if ($scopeType -notmatch '^(?i:Default|Fallback|Subnet|CIDR|IPPrefix|Prefix|ComputerName|Hostname|ComputerPrefix|HostnamePrefix)$') { throw "SEMT_SETUP_SOURCE_MAP_INVALID: row $number has an unsupported ScopeType '$scopeType'." }
        if ($scopeType -match '^(?i:Subnet|CIDR)$' -and $scopeValue -notmatch '^\d{1,3}(\.\d{1,3}){3}/([0-9]|[12][0-9]|3[0-2])$') { throw "SEMT_SETUP_SOURCE_MAP_INVALID: row $number requires a CIDR ScopeValue." }
        if ($source -notmatch '^\\\\[^\\]+\\[^\\]+') { throw "SEMT_SETUP_SOURCE_MAP_INVALID: row $number SetupSourcePath must be UNC: $source" }
    }
    return ('Valid:{0} row(s)' -f $rows.Count)
}

function Resolve-Windows11OptionRules {
    <#
    Returns warnings. Blocking configuration errors throw before any computer is contacted.
    #>
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Options)

    $warnings = New-Object System.Collections.Generic.List[string]
    $setup = (Get-SemtOption -Options $Options -Name 'Windows11Upgrade.AllowSetupUpgrade') -or (Get-SemtOption -Options $Options -Name 'Windows11Upgrade.DirectSetupUpgrade')
    $source = [string](Get-SemtOption -Options $Options -Name 'Windows11Upgrade.SetupSourcePath')
    $map = [string](Get-SemtOption -Options $Options -Name 'Windows11Upgrade.SetupSourceMapPath')
    $existingOnly = [bool](Get-SemtOption -Options $Options -Name 'Windows11Upgrade.SkipSetupMediaPreCopy')
    $dryRun = [bool](Get-SemtOption -Options $Options -Name 'DryRun')
    if ($setup -and -not $existingOnly -and [string]::IsNullOrWhiteSpace($source) -and [string]::IsNullOrWhiteSpace($map)) {
        $message = 'SEMT_SETUP_SOURCE_REQUIRED: setup upgrade is enabled but no setup source or source map is configured. Set Windows11Upgrade.SetupSourcePath or SetupSourceMapPath, or enable "Use existing media only" (legacy media is reused when present).'
        if ($dryRun) { $warnings.Add($message) } else { throw $message }
    }
    if (-not [string]::IsNullOrWhiteSpace($source) -and $source -notmatch '^\\\\[^\\]+\\[^\\]+' -and -not (Get-SemtOption -Options $Options -Name 'Windows11Upgrade.ConfirmLocalSetupSource')) {
        throw ("SEMT_SETUP_SOURCE_NOT_UNC: '{0}' is not a UNC path reachable by targets. Use a UNC share or set Windows11Upgrade.ConfirmLocalSetupSource=true for a test source." -f $source)
    }
    $mapStatus = Test-Windows11SetupSourceMapSyntax -Path $map
    if ($mapStatus -like 'SkippedOperatorCannotRead:*') { $warnings.Add(('Setup source map preflight skipped because this account cannot read it: {0}' -f $mapStatus.Substring(26))) }
    if ((Get-SemtOption -Options $Options -Name 'Windows11Upgrade.DirectSetupUpgrade') -and -not (Get-SemtOption -Options $Options -Name 'Windows11Upgrade.AllowReboot')) {
        $warnings.Add('DirectSetupUpgrade is enabled with AllowReboot disabled: Windows Setup can run but no reboot will be triggered; completion waits for a user or maintenance reboot.')
    }
    return @($warnings.ToArray())
}

function Invoke-Windows11PreCycle {
    param([Parameter(Mandatory = $true)]$Context,[string[]]$Computers)

    $rows = New-Object System.Collections.Generic.List[object]
    $seen = @{}
    foreach ($kind in @('Intune','AD')) {
        $state = $Context.Inventory[$kind]
        if ($null -eq $state -or $null -eq $state.Map -or $state.Map.Count -eq 0) { continue }
        foreach ($computer in @($Computers)) {
            $key = Get-SemtComputerKey -ComputerName $computer
            if ([string]::IsNullOrWhiteSpace($key) -or $seen.ContainsKey($key) -or -not $state.Map.ContainsKey($key)) { continue }
            $row = $state.Map[$key]
            $os = [string](Get-SemtPropertyValue -InputObject $row -Name 'OperatingSystem' -Default '')
            $isWindows11 = ($os -match '(?i)\bWindows\s+11\b')
            if (-not $isWindows11 -and $kind -eq 'Intune') {
                $version = [string](Get-SemtPropertyValue -InputObject $row -Name 'OSVersion' -Default '')
                if ($version -match '^(\d+)\.(\d+)\.(\d+)') { $isWindows11 = ([int]$Matches[3] -ge 22000) }
            }
            if (-not $isWindows11) { continue }
            $seen[$key] = $true
            $result = New-SemtResult -Campaign 'Windows11Upgrade' -ComputerName $computer -CycleNumber $Context.CycleNumber -WaveName $Context.RunContext.WaveName -RunId $Context.RunContext.RunId -LauncherStatus 'ALREADY_WINDOWS11' -NextAction ('NO_ACTION_{0}_ALREADY_WINDOWS11' -f $kind.ToUpperInvariant()) -Detail ('{0} inventory already reports Windows 11 ({1}); PsExec launch skipped.' -f $kind,$state.Source) -ExtensionFields (@($script:Windows11ExtensionFields) + @($script:Windows11EngineFields))
            [void](Complete-SemtResultStatus -Result $result)
            $rows.Add([pscustomobject]$result)
        }
    }
    return @($rows.ToArray())
}

function Get-Windows11CentralLogBucket {
    param([Parameter(Mandatory = $true)]$Result)

    $launcher = [string](Get-SemtPropertyValue -InputObject $Result -Name 'LauncherStatus' -Default '')
    if ($launcher -like '*ADMIN_SHARE_UNREACHABLE') { return 'AdminShareUnreachable' }
    if ($launcher -eq 'PSEXEC_COMMUNICATION_LOST') { return 'PsExecCommunicationLost' }
    if ($launcher -in @('REMOTE_LOG_COLLECTION_FAILED','REMOTE_RESULT_INCOMPLETE','REMOTE_RESULT_STALE')) { return 'RemoteCollectionFailed' }
    foreach ($name in @('RemoteStatus','Status','LauncherStatus')) {
        $value = [string](Get-SemtPropertyValue -InputObject $Result -Name $name -Default '')
        if ($value -like 'INSUFFICIENT_DISK*') { return 'InsufficientDisk' }
        if ($value -in @('WINDOWS11_COMPAT_BLOCKER','WINDOWS11_HARDWARE_NOT_CAPABLE')) { return 'Compatibility' }
        if ($value -like 'PSEXEC_TIMEOUT*') { return 'PsExecTimeout' }
        if ($value -like 'SKIPPED_*') { return 'Skipped' }
        if ($value -like 'TARGET_*' -or $value -like 'LEGACY_*' -or $value -eq 'RUN_GUARD_ACTIVE') { return 'TargetBusy' }
        if ($value -eq 'SETUP_PROCESS_TIMEOUT') { return 'SetupProcessTimeout' }
        if ($value -eq 'SETUP_MEDIA_COPY_TIMEOUT') { return 'SetupMediaCopyTimeout' }
        if ($value -eq 'SETUP_MEDIA_COPY_FAILED') { return 'SetupMediaCopyFailure' }
        if ($value -eq 'SETUP_MEDIA_MANIFEST_VALIDATION_FAILED') { return 'SetupMediaManifestFailure' }
        if ($value -in @('SETUP_MEDIA_DUPLICATION_BLOCKED','SETUP_MEDIA_LEGACY_ADOPTION_FAILED','SETUP_MEDIA_LEGACY_UNTRUSTED','SETUP_MEDIA_IN_USE','SETUP_MEDIA_STATE_INDETERMINATE','SETUP_MEDIA_PATH_REJECTED')) { return 'SetupMediaReuse' }
        if ($value -eq 'SETUP_PROCESS_MONITOR_INTERRUPTED') { return 'SetupProcessInterrupted' }
        if ($value -eq 'SETUP_SOURCE_LANGUAGE_UNAVAILABLE') { return 'SetupSourceLanguageUnavailable' }
        if ($value -eq 'SETUP_PROFILE_DUPLICATE_REPAIRED_REBOOT_REQUIRED') { return 'SetupMigrationProfileRepaired' }
        if ($value -in @('SETUP_MIGRATION_PROFILE_FAILURE','SETUP_MIGRATION_PROFILE_REPAIR_FAILED')) { return 'SetupMigrationProfileFailure' }
        if ($value -eq 'SETUP_MIGRATION_PLUGIN_FAILURE') { return 'SetupMigrationPluginFailure' }
        if ($value -in @('SETUP_SUBNET_COPY_LEASE_TIMEOUT','SETUP_SOURCE_COPY_LEASE_TIMEOUT')) { return 'SetupCopyLeaseTimeout' }
    }
    $detail = '{0} {1}' -f (Get-SemtPropertyValue -InputObject $Result -Name 'Detail' -Default ''),(Get-SemtPropertyValue -InputObject $Result -Name 'JobErrorMessage' -Default '')
    if ($detail -match '0xC1900200|0xC1900202|0xC1900208|Compatibility failure|Compatibility blocker') { return 'Compatibility' }
    if ($detail -match '0x8007001F|Duplicate profile detected') { return 'SetupMigrationProfileFailure' }
    if ($detail -match '0x8007007F|CscMig\.dll|WSManMigrationPlugin\.dll|RasMigPlugin\.dll') { return 'SetupMigrationPluginFailure' }
    $exitCode = 0
    if ([int]::TryParse([string](Get-SemtPropertyValue -InputObject $Result -Name 'ExitCode' -Default ''), [ref]$exitCode) -and $exitCode -eq 0) { return 'Success' }
    return 'Errors'
}

function Invoke-Windows11WorkerRecovery {
    <#
    When PsExec ends before a completed attempt result, the endpoint log tells whether
    setup.exe was started and still monitored (SETUP_PROCESS_MONITOR_INTERRUPTED).
    #>
    param($Result,[string]$RemoteDataShare,$Contract)

    $logsShare = Join-Path -Path $RemoteDataShare -ChildPath 'Logs'
    if (-not (Test-Path -LiteralPath $logsShare -PathType Container)) { return $null }
    $since = (ConvertTo-SemtUtcDateTime -Value $Contract.CreatedUtc)
    $latest = Get-ChildItem -LiteralPath $logsShare -Filter 'SEMT_Windows11Upgrade_*.log' -File -ErrorAction SilentlyContinue | Where-Object { $null -eq $since -or $_.LastWriteTimeUtc -ge $since.AddMinutes(-2) } | Sort-Object LastWriteTimeUtc -Descending | Select-Object -First 1
    if (-not $latest) { return $null }
    $tail = (@(Get-Content -LiteralPath $latest.FullName -Tail 150 -ErrorAction SilentlyContinue) -join "`n")
    if ($tail -notmatch [regex]::Escape([string]$Contract.CorrelationId) -and $tail -notmatch 'setup\.exe') { return $null }
    $hasFinal = ($tail -match 'Final Status=')
    $started = ($tail -match 'Starting setup upgrade:|setup\.exe started\.')
    $stillRunning = ($tail -match 'setup\.exe still running\.')
    $exited = ($tail -match 'setup\.exe exited\.')
    if (-not $hasFinal -and $started -and ($stillRunning -or -not $exited)) {
        return [pscustomobject]@{ Status = 'SETUP_PROCESS_MONITOR_INTERRUPTED'; NextAction = 'CHECK_SETUP_PROCESS_OR_OS_STATUS'; ExitCode = '3'; Detail = ('PsExec ended before the endpoint result, but the endpoint log shows setup.exe started and still monitored. RemoteLog={0}' -f $latest.FullName) }
    }
    return $null
}

function Get-Windows11FailureCategory {
    param($Result)

    $status = (@([string](Get-SemtPropertyValue -InputObject $Result -Name 'RemoteStatus' -Default ''),[string](Get-SemtPropertyValue -InputObject $Result -Name 'Status' -Default ''),[string](Get-SemtPropertyValue -InputObject $Result -Name 'LauncherStatus' -Default '')) | Where-Object { $_ } | Select-Object -First 1)
    $status = ([string]$status).ToUpperInvariant()
    $evidence = '{0} {1}' -f (Get-SemtPropertyValue -InputObject $Result -Name 'Detail' -Default ''),(Get-SemtPropertyValue -InputObject $Result -Name 'JobErrorMessage' -Default '')
    if ($status -like '*ADMIN_SHARE_UNREACHABLE' -or $evidence -match 'FailureType=(DNS_FAILED|SMB_PORT_445_UNREACHABLE|PING_OK_ADMIN_SHARE_FAILED|PING_FAILED_ADMIN_SHARE_FAILED)') { return 'NetworkTransient' }
    if ($status -in @('DIRECT_SETUP_UPGRADE_FAILED','SETUP_UPGRADE_FAILED','SETUP_MIGRATION_PROFILE_FAILURE','SETUP_MIGRATION_PROFILE_REPAIR_FAILED','SETUP_MIGRATION_PLUGIN_FAILURE','SETUP_PROCESS_TIMEOUT','SETUP_PROCESS_MONITOR_INTERRUPTED','SETUP_MEDIA_COPY_FAILED','SETUP_MEDIA_COPY_TIMEOUT','SETUP_MEDIA_MANIFEST_VALIDATION_FAILED')) { return 'SetupFailure' }
    if ($status -in @('INSUFFICIENT_DISK','INSUFFICIENT_DISK_AFTER_CLEANUP','UNSUPPORTED_OS','NOT_INTUNE_ENROLLED','WINDOWS11_COMPAT_BLOCKER','WINDOWS11_HARDWARE_NOT_CAPABLE','WU_POLICY_BLOCKER','SETUP_SOURCE_LANGUAGE_UNAVAILABLE','SETUP_MEDIA_DUPLICATION_BLOCKED','SETUP_MEDIA_LEGACY_ADOPTION_FAILED','SETUP_MEDIA_LEGACY_UNTRUSTED','SETUP_MEDIA_STATE_INDETERMINATE','SETUP_MEDIA_PATH_REJECTED','DATA_FOLDER_UNTRUSTED','PROGRAM_FOLDER_UNTRUSTED')) { return 'OperatorAction' }
    if ($status -like 'TARGET_*' -or $status -like 'LEGACY_*') { return 'ExecutionTransient' }
    if ($status -in @('ERROR','JOB_ERROR','RUNSPACE_BROKEN','PSEXEC_EXIT_UNKNOWN','PSEXEC_COMMUNICATION_LOST','PSEXEC_TIMEOUT','PSEXEC_TIMEOUT_REMOTE_RUNNING','REMOTE_LOG_COLLECTION_FAILED','REMOTE_RESULT_INCOMPLETE','REMOTE_PAYLOAD_COPY_FAILED','SETUP_CACHE_LOCKED','SETUP_MEDIA_IN_USE','SETUP_SUBNET_COPY_LEASE_TIMEOUT','SETUP_SOURCE_COPY_LEASE_TIMEOUT') -or $status -like 'PSEXEC_EXIT_*') { return 'ExecutionTransient' }
    return ''
}

function Get-Windows11RetryDelayMinutes {
    param([string]$Category,[int]$Count,[int]$Hours)

    $maximum = [math]::Max(1, $Hours * 60)
    switch ($Category) {
        'NetworkTransient' { $delays = @(5,15,30,60); return [math]::Min($maximum, $delays[[math]::Min($Count - 1, $delays.Count - 1)]) }
        'ExecutionTransient' { $delays = @(15,30,60,120); return [math]::Min($maximum, $delays[[math]::Min($Count - 1, $delays.Count - 1)]) }
        'SetupFailure' { return [math]::Min($maximum, 360) }
        'OperatorAction' { return $maximum }
        default { return 0 }
    }
}

function Test-Windows11RunGuardEntryRetained {
    param($Entry)

    $hours = [int](Get-SemtPropertyValue -InputObject $Entry -Name 'RunGuardHours' -Default 3)
    $started = ConvertTo-SemtUtcDateTime -Value (Get-SemtPropertyValue -InputObject $Entry -Name 'LastStartedUtc' -Default '')
    if ($null -eq $started) { return $false }
    $retry = ConvertTo-SemtUtcDateTime -Value (Get-SemtPropertyValue -InputObject $Entry -Name 'RetryAfterUtc' -Default '')
    $now = (Get-Date).ToUniversalTime()
    return (($now - $started).TotalHours -lt [math]::Max(1, $hours) -or ($null -ne $retry -and $retry -gt $now))
}

function Get-Windows11RunGuardDecision {
    param($Entry,[System.Collections.IDictionary]$Options)

    $none = [pscustomobject]@{ Block = $false; BypassEndpointGuard = $false }
    if ($null -eq $Entry) { return $none }
    $hours = [int](Get-SemtOption -Options $Options -Name 'TechnicianRunGuardHours')
    $now = (Get-Date).ToUniversalTime()
    $state = [string](Get-SemtPropertyValue -InputObject $Entry -Name 'State' -Default '')
    $started = ConvertTo-SemtUtcDateTime -Value (Get-SemtPropertyValue -InputObject $Entry -Name 'LastStartedUtc' -Default '')
    $retry = ConvertTo-SemtUtcDateTime -Value (Get-SemtPropertyValue -InputObject $Entry -Name 'RetryAfterUtc' -Default '')
    $status = (@([string](Get-SemtPropertyValue -InputObject $Entry -Name 'RemoteStatus' -Default ''),[string](Get-SemtPropertyValue -InputObject $Entry -Name 'Status' -Default '')) | Where-Object { $_ } | Select-Object -First 1)
    $status = ([string]$status).ToUpperInvariant()
    $detail = ('Technician run guard. FQDN={0}; LastStartedUtc={1}; State={2}; LastStatus={3}; FailureCategory={4}; ConsecutiveFailureCount={5}; RetryAfterUtc={6}' -f $Entry.ComputerFqdn,$Entry.LastStartedUtc,$state,$status,(Get-SemtPropertyValue -InputObject $Entry -Name 'FailureCategory' -Default ''),(Get-SemtPropertyValue -InputObject $Entry -Name 'ConsecutiveFailureCount' -Default ''),(Get-SemtPropertyValue -InputObject $Entry -Name 'RetryAfterUtc' -Default ''))
    if ($state -ne 'Result') {
        if ($null -ne $started -and ($now - $started).TotalHours -lt [math]::Min([double]$hours, [double]$script:Windows11StartedNoResultHours)) {
            return [pscustomobject]@{ Block = $true; BypassEndpointGuard = $false; Status = 'SKIPPED_BY_TECH_RUN_GUARD_STARTED_NO_RESULT'; NextAction = 'WAIT_RUN_GUARD_EXPIRY'; Detail = $detail; Fields = @{ RetryAfterUtc = ''; FailureCategory = ''; ConsecutiveFailureCount = '' } }
        }
        return $none
    }
    if ($null -ne $retry -and $now -lt $retry) {
        return [pscustomobject]@{ Block = $true; BypassEndpointGuard = $false; Status = 'SKIPPED_BY_TECH_RUN_GUARD'; NextAction = 'WAIT_RETRY_BACKOFF'; Detail = $detail; Fields = @{ RetryAfterUtc = [string]$Entry.RetryAfterUtc; FailureCategory = [string](Get-SemtPropertyValue -InputObject $Entry -Name 'FailureCategory' -Default ''); ConsecutiveFailureCount = [string](Get-SemtPropertyValue -InputObject $Entry -Name 'ConsecutiveFailureCount' -Default '') } }
    }
    $category = [string](Get-SemtPropertyValue -InputObject $Entry -Name 'FailureCategory' -Default '')
    if ($null -ne $retry -and $now -ge $retry) {
        $setupFailures = @('DIRECT_SETUP_UPGRADE_FAILED','SETUP_UPGRADE_FAILED','SETUP_MIGRATION_PROFILE_FAILURE','SETUP_MIGRATION_PROFILE_REPAIR_FAILED','SETUP_MIGRATION_PLUGIN_FAILURE','SETUP_PROCESS_TIMEOUT','SETUP_PROCESS_MONITOR_INTERRUPTED','SETUP_MEDIA_COPY_FAILED','SETUP_MEDIA_COPY_TIMEOUT','SETUP_MEDIA_MANIFEST_VALIDATION_FAILED')
        $endpointTransient = @('ERROR','SETUP_CACHE_LOCKED','SETUP_SUBNET_COPY_LEASE_TIMEOUT','SETUP_SOURCE_COPY_LEASE_TIMEOUT')
        if (($category -eq 'SetupFailure' -and $status -in $setupFailures) -or ($category -eq 'ExecutionTransient' -and $status -in $endpointTransient)) { return [pscustomobject]@{ Block = $false; BypassEndpointGuard = $true; Detail = $detail } }
        return $none
    }
    $inProgress = @('DIRECT_SETUP_UPGRADE_STARTED','DIRECT_SETUP_UPGRADE_REBOOT_REQUIRED','DIRECT_SETUP_UPGRADE_REBOOT_SCHEDULED_NO_USER','DIRECT_SETUP_UPGRADE_REBOOT_SKIPPED_USER_CONNECTED','DIRECT_SETUP_UPGRADE_REBOOT_SKIPPED_USER_DETECTION_FAILED','SETUP_UPGRADE_STARTED','SETUP_UPGRADE_REBOOT_REQUIRED','SETUP_UPGRADE_REBOOT_SCHEDULED_NO_USER','SETUP_UPGRADE_REBOOT_SKIPPED_USER_CONNECTED','SETUP_UPGRADE_REBOOT_SKIPPED_USER_DETECTION_FAILED','PENDING_REBOOT','PENDING_REBOOT_USER_CONNECTED','PENDING_REBOOT_USER_DETECTION_FAILED','PENDING_REBOOT_SCHEDULE_FAILED','WINDOWS_UPDATE_FORCE_TRIGGERED','WU_RESET_COMPLETED')
    $withinWindow = ($null -ne $started -and ($now - $started).TotalHours -lt $hours)
    if ($withinWindow -and ($status -in $inProgress -or $status -like '*_STARTED' -or $status -like '*_REBOOT_REQUIRED' -or $status -like '*_REBOOT_SCHEDULED_NO_USER')) {
        return [pscustomobject]@{ Block = $true; BypassEndpointGuard = $false; Status = 'SKIPPED_BY_TECH_RUN_GUARD'; NextAction = 'WAIT_RUN_GUARD_EXPIRY'; Detail = $detail; Fields = @{ RetryAfterUtc = ''; FailureCategory = $category; ConsecutiveFailureCount = [string](Get-SemtPropertyValue -InputObject $Entry -Name 'ConsecutiveFailureCount' -Default '') } }
    }
    return $none
}

function New-Windows11RunGuardEntry {
    param($Previous,[string]$ComputerFqdn,[string]$InputComputerName,[ValidateSet('Started','Result')][string]$State,$Result,[string]$JobId,[int]$CycleNumber,[int]$Hours,[hashtable]$Meta)

    $now = (Get-Date).ToUniversalTime()
    $started = $now
    if ($State -eq 'Result' -and $Previous) {
        $previousStarted = ConvertTo-SemtUtcDateTime -Value (Get-SemtPropertyValue -InputObject $Previous -Name 'LastStartedUtc' -Default '')
        if ($previousStarted) { $started = $previousStarted }
    }
    $previousCategory = if ($Previous) { [string](Get-SemtPropertyValue -InputObject $Previous -Name 'FailureCategory' -Default '') } else { '' }
    $previousCount = 0
    if ($Previous) { [void][int]::TryParse([string](Get-SemtPropertyValue -InputObject $Previous -Name 'ConsecutiveFailureCount' -Default '0'), [ref]$previousCount) }
    $category = $previousCategory
    $count = $previousCount
    $retryAfter = ''
    if ($State -eq 'Result') {
        $category = Get-Windows11FailureCategory -Result $Result
        $count = if ([string]::IsNullOrWhiteSpace($category)) { 0 } elseif ($category -eq $previousCategory) { $previousCount + 1 } else { 1 }
        if (-not [string]::IsNullOrWhiteSpace($category)) {
            $delay = Get-Windows11RetryDelayMinutes -Category $category -Count $count -Hours $Hours
            if ($delay -gt 0) { $retryAfter = $now.AddMinutes($delay).ToString('o') }
        }
    }
    $detail = if ($Result) { [string](Get-SemtPropertyValue -InputObject $Result -Name 'Detail' -Default '') } else { '' }
    if ($detail.Length -gt 800) { $detail = $detail.Substring(0, 800) }
    return [pscustomobject]@{
        ComputerFqdn = $ComputerFqdn
        InputComputerName = $InputComputerName
        State = $State
        LastStartedUtc = $started.ToString('o')
        LastUpdatedUtc = $now.ToString('o')
        LastResultUtc = if ($State -eq 'Result') { $now.ToString('o') } else { '' }
        Status = if ($Result) { [string](Get-SemtPropertyValue -InputObject $Result -Name 'Status' -Default '') } else { '' }
        LauncherStatus = if ($Result) { [string](Get-SemtPropertyValue -InputObject $Result -Name 'LauncherStatus' -Default '') } else { '' }
        RemoteStatus = if ($Result) { [string](Get-SemtPropertyValue -InputObject $Result -Name 'RemoteStatus' -Default '') } else { '' }
        NextAction = if ($Result) { [string](Get-SemtPropertyValue -InputObject $Result -Name 'NextAction' -Default '') } else { '' }
        Detail = $detail
        FailureCategory = $category
        ConsecutiveFailureCount = $count
        RetryAfterUtc = $retryAfter
        RunGuardHours = $Hours
        JobId = $JobId
        CycleNumber = $CycleNumber
        WaveName = [string]$Meta.WaveName
        RunId = [string]$Meta.RunId
        TechnicianAccount = [string]$Meta.TechnicianAccount
        LauncherVersion = [string]$Meta.LauncherVersion
    }
}

function Invoke-Windows11PostCycle {
    param([Parameter(Mandatory = $true)]$Context,[object[]]$Rows)

    $distribution = @($Rows | Where-Object { $_ -and -not [string]::IsNullOrWhiteSpace([string](Get-SemtPropertyValue -InputObject $_ -Name 'SelectedSetupSourcePath' -Default '')) } | Group-Object -Property SelectedSetupSourcePath | Sort-Object Count -Descending | ForEach-Object { [pscustomobject]@{ SelectedSetupSourcePath = [string]$_.Name; ComputerCount = [int]$_.Count } })
    if ($distribution.Count -eq 0) { return }
    $path = Join-Path -Path $Context.RunContext.ReportRoot -ChildPath ('SetupSource_Distribution_cycle{0}_{1}.csv' -f $Context.CycleNumber,(Get-Date -Format 'yyyyMMdd-HHmmss'))
    $distribution | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding UTF8
    & $Context.Log ('Cycle {0}: setup source distribution: {1}' -f $Context.CycleNumber,$path) 'Green'
}

function Get-SemtWindows11UpgradeProfile {
    [CmdletBinding()]
    param()

    $coreColumns = @('Timestamp','CycleNumber','ComputerName','ConnectionTarget','Status','NextAction','LauncherStatus','RemoteStatus','RemoteNextAction','ExitCode','Detail','RetryAfterUtc','FailureCategory','ConsecutiveFailureCount','DnsResolved','DnsAddressList','PingReachable','SmbPort445Reachable','AdminShareReachable','AdminShareFailureType','RemotePayloadCopyAttempts','LocalScriptVersion','LocalScriptSha256','RemoteScriptSha256','PsExecExitCode','TargetGuardStatus','LegacyActivityStatus','LegacyActivityDetail')
    $tail = @('RemoteLogsPath','PsExecLogPath','JobErrorMessage','CorrelationId','RunId','WaveName','Campaign','WorkerContractVersion')
    $columns = @($coreColumns) + @($script:Windows11ExtensionFields) + @(Get-SemtInventoryFieldNames -Kind AD) + @(Get-SemtInventoryFieldNames -Kind Intune) + $tail
    return [pscustomobject]@{
        Campaign = 'Windows11Upgrade'
        Title = 'Windows 11 upgrade'
        ProfileVersion = '1.0.0'
        EndpointScriptRelativePath = 'Campaigns\Windows11Upgrade\Scripts\SmartM365-Invoke-EndpointMigration-Windows11Upgrade.ps1'
        PsExecUseElevatedToken = $false
        InventoryKinds = @('AD','Intune')
        InventoryRefreshOptions = @{ AD = 'Windows11Upgrade.PreCycleAdInventoryRefresh'; Intune = 'Windows11Upgrade.PreCycleIntuneInventoryRefresh' }
        ExtensionFields = @($script:Windows11ExtensionFields) + @($script:Windows11EngineFields)
        ProjectedFields = @($script:Windows11ExtensionFields)
        ReportColumns = @($columns | Select-Object -Unique)
        SeparatedStatuses = @('ADMIN_SHARE_UNREACHABLE','RUN_GUARD_ACTIVE','SKIPPED_BY_TECH_RUN_GUARD','SKIPPED_BY_TECH_RUN_GUARD_STARTED_NO_RESULT','DNS_PREFLIGHT_ALL_SAMPLES_FAILED')
        CentralLogBuckets = @('Success','Skipped','AdminShareUnreachable','InsufficientDisk','Compatibility','SetupSourceLanguageUnavailable','SetupMigrationProfileFailure','SetupMigrationProfileRepaired','SetupMigrationPluginFailure','SetupCopyLeaseTimeout','SetupMediaCopyTimeout','SetupMediaCopyFailure','SetupMediaManifestFailure','SetupMediaReuse','SetupProcessTimeout','SetupProcessInterrupted','PsExecTimeout','PsExecCommunicationLost','RemoteCollectionFailed','TargetBusy','Errors')
        EvidenceRules = @(
            [pscustomobject]@{ Folder = 'Logs\SetupUpgrade'; Filter = '*'; Recurse = $false; CurrentAttemptOnly = $true; MaxFiles = 12; MaxBytes = 5MB; TargetFolder = 'Setup' }
        )
        RemoteLogsPathTemplate = '\\{0}\C$\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade\Logs'
        CommunicationLostWaitOption = ''
        CommunicationLostPollOption = ''
        CompletedRules = @(
            [pscustomobject]@{ Label = 'already-Windows11'; ArchiveFileName = 'Computers.AlreadyWindows11.txt'; Test = ${function:Test-Windows11AlreadyUpgraded} },
            [pscustomobject]@{ Label = 'hardware-not-capable'; ArchiveFileName = 'Computers.HardwareNotCapable.txt'; Test = ${function:Test-Windows11HardwareNotCapable} }
        )
        RunGuardSchemaVersion = 1
        Hooks = [pscustomobject]@{
            ResolveOptionRules = ${function:Resolve-Windows11OptionRules}
            PreCycle = ${function:Invoke-Windows11PreCycle}
            PerComputerArguments = { param($Context,$Computer) @() }
            PostCycle = ${function:Invoke-Windows11PostCycle}
            NextAction = ${function:Get-Windows11NextAction}
            CentralLogBucket = ${function:Get-Windows11CentralLogBucket}
            WorkerRecovery = ${function:Invoke-Windows11WorkerRecovery}
            RunGuardPrune = ${function:Test-Windows11RunGuardEntryRetained}
            RunGuardDecision = ${function:Get-Windows11RunGuardDecision}
            RunGuardEntry = ${function:New-Windows11RunGuardEntry}
            AdaptiveDelaySeconds = { param($Rows,$MinimumDelaySeconds) $MinimumDelaySeconds }
        }
    }
}

# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC/T2wAUsO9mAXg
# 63fe/xeZMHUds5v7DGL4lOeRy13euKCCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDlSmO0S8A6FLSRF9mwNDYj
# uXp7YI+rY+pLSy6LypIxtTANBgkqhkiG9w0BAQEFAASCAYCnCvdh4PqLB5Yn3YOL
# SOZSifqUsv8GGWe3k18wr503ZnkzubfzJRM4gU84dF0/hKru6/UqSHrVMBNwqaVy
# CSN7pog4XUvkiUbxZ9mQBI9KPxt/lp9TAzb/YbHFe7IlFb6hZD+8S/GoMnY6HzO1
# 6KGNDt1Ao5uoyUatR3zFcoiEF68zTI/MaIZIao52FgYyolfr9ndJ+YhMdH3hwHW2
# nann0RKau6/hKOgc5akxmf/wa+R58fI3UmSm1/KDuCcijNaH1cw78HQ/ci6GpW81
# mMqpKOdlYxCJUYrqN1FDNwlT6ozAtrBNBIzN9HqtEwGmjnqNSFSekOdoBUd4Rr7A
# VhBPISF7IdZ5GPTRRUzVCNzE/i6dCQFtrq9zc+/X47yWlLJoqhw54L1lFSZ+rq/J
# l8k2ffLSdToazXRHEXrX2vHc3kHuGRvxnIdC0WQwB4vL/+kGFCk2Do75QbOow8jY
# Z6mrW4VzEB8uXXTkl5gTLfsHyWVkGZqk/Sp4AaT/StpSgs4=
# SIG # End signature block
