<#
.SYNOPSIS
    Converts the saved options of the previous Hybrid Join and Windows 11 toolkits into the local
    configuration of Smart Endpoint Migration Toolkit.

.DESCRIPTION
    One-time, operator-started conversion. It only reads the previous toolkits:
      <LegacyDevicesRoot>\IntuneHybridJoinToolkit\IntuneHybridJoinToolkit.config   (EHJIR_*)
      <LegacyDevicesRoot>\Windows11UpgradeToolkit\Windows11UpgradeToolkit.config   (W11UT_*)
      persistent EHJIR_* / W11UT_* environment variables (Machine, then User)
    LOT folders, Computers.txt, inventories, run history, reports and logs are not converted;
    LOT-level values are only counted in the report.

    Without -Apply nothing is written: the conversion plan is shown (preview).
    With -Apply it writes only:
      Config\SmartEndpointMigrationToolkit.local.json   options valid for both Campaigns
      Config\GuiState.<Campaign>.local.json            common options whose value differs per Campaign
      Config\LegacyConfigImport.<timestamp>.local.json report of the conversion
    An existing file is copied to <file>.<timestamp>.bak before it is changed. An explicit value
    already present in a SEMT file is kept unless -ReplaceExistingValues is used.

    Values equal to the previous GUI default are not carried over (the SEMT default applies)
    unless -IncludeLegacyDefaults is used. Launch-time choices (dry run, audit only, ignore run
    guard) are never carried over.

.PARAMETER LegacyDevicesRoot
    Folder that contains the IntuneHybridJoinToolkit and/or Windows11UpgradeToolkit folders.

.PARAMETER PackageRoot
    Root of Smart Endpoint Migration Toolkit. Default: the parent of this script folder.

.PARAMETER Apply
    Writes the planned configuration. Without it the script only previews.

.PARAMETER IncludeLegacyDefaults
    Also carries values equal to the default of the previous GUI.

.PARAMETER ReplaceExistingValues
    Replaces explicit values already present in the SEMT files (after backup).

.PARAMETER LegacyEnvironmentScope
    Environment scopes read for EHJIR_* / W11UT_* variables (later scopes win). Default: Machine, User.

.PARAMETER IgnoreLegacyEnvironment
    Does not read EHJIR_* / W11UT_* environment variables.

.PARAMETER ReportPath
    Also writes the JSON report to this path (preview included).

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\SmartM365-EndpointMigration-Import-LegacyConfig.ps1 -LegacyDevicesRoot 'D:\Tools\SmartM365\Devices'

.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Scripts\SmartM365-EndpointMigration-Import-LegacyConfig.ps1 -LegacyDevicesRoot 'D:\Tools\SmartM365\Devices' -Apply

.VERSION
    1.0.0

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)][string]$LegacyDevicesRoot,
    [string]$PackageRoot,
    [switch]$Apply,
    [switch]$IncludeLegacyDefaults,
    [switch]$ReplaceExistingValues,
    [ValidateSet('Process','Machine','User')][string[]]$LegacyEnvironmentScope = @('Machine','User'),
    [switch]$IgnoreLegacyEnvironment,
    [string]$ReportPath,
    [switch]$PassThru
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($PackageRoot)) { $PackageRoot = Split-Path -Path $PSScriptRoot -Parent }
$PackageRoot = (Resolve-Path -LiteralPath $PackageRoot).ProviderPath
Import-Module (Join-Path -Path $PackageRoot -ChildPath 'Modules\SmartEndpointMigration.Core\SmartEndpointMigration.Core.psd1') -MinimumVersion 1.0.0 -Force
$catalog = Import-SemtOptionCatalog -PackageRoot $PackageRoot
$scriptVersion = '1.0.0'
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$campaigns = @('HybridJoin','Windows11Upgrade')
$warnings = New-Object 'System.Collections.Generic.List[string]'

#region Legacy key map

function Get-LegacyKeyMapEntry {
    param(
        [Parameter(Mandatory = $true)][string]$Key,
        [string]$Option = '',
        [string]$LegacyDefault = '',
        [ValidateSet('Map','LaunchChoice','NoEquivalent','Runtime')][string]$Kind = 'Map',
        [switch]$Invert,
        [string]$Note = ''
    )
    return [pscustomobject]@{ Key = $Key; Option = $Option; LegacyDefault = $LegacyDefault; Kind = $Kind; Invert = [bool]$Invert; Note = $Note }
}

# Legacy defaults are those of the previous GUI, which saves every option to its .config file.
$hybridJoinMap = @(
    Get-LegacyKeyMapEntry -Key 'EHJIR_THROTTLE' -Option 'ThrottleLimit' -LegacyDefault '10'
    Get-LegacyKeyMapEntry -Key 'EHJIR_GLOBAL_CONCURRENCY_LIMIT' -Option 'GlobalWorkerLimit' -LegacyDefault '15' -Note 'SEMT has one shared cap for both Campaigns on the technician computer.'
    Get-LegacyKeyMapEntry -Key 'EHJIR_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES' -Option 'GlobalWorkerLeaseTimeoutMinutes' -LegacyDefault '0'
    Get-LegacyKeyMapEntry -Key 'EHJIR_DELAY_BETWEEN_CYCLES_MINUTES' -Option 'DelayBetweenCyclesMinutes' -LegacyDefault '5'
    Get-LegacyKeyMapEntry -Key 'EHJIR_DISABLE_NIGHT_PAUSE' -Option 'NightPauseEnabled' -LegacyDefault '0' -Invert
    Get-LegacyKeyMapEntry -Key 'EHJIR_NIGHT_PAUSE_START_HOUR' -Option 'NightPauseStartHour' -LegacyDefault '20'
    Get-LegacyKeyMapEntry -Key 'EHJIR_NIGHT_PAUSE_END_HOUR' -Option 'NightPauseEndHour' -LegacyDefault '7'
    Get-LegacyKeyMapEntry -Key 'EHJIR_INTUNE_RETRY_SLEEP_MINUTES' -Option 'HybridJoin.IntuneRetrySleepMinutes' -LegacyDefault '3' -Note 'Previous GUI default 3, previous CMD default 5, SEMT default 5 (minimum 1).'
    Get-LegacyKeyMapEntry -Key 'EHJIR_INTUNE_RETRY_MAX_RETRIES' -Option 'HybridJoin.IntuneRetryMaxRetries' -LegacyDefault '5'
    Get-LegacyKeyMapEntry -Key 'EHJIR_RETRY_AFTER_REBOOT_DELAY_SECONDS' -Option 'RetryAfterRebootDelaySeconds' -LegacyDefault '300'
    Get-LegacyKeyMapEntry -Key 'EHJIR_RETRY_AFTER_REBOOT_MAX_ATTEMPTS' -Option 'RetryAfterRebootMaxAttempts' -LegacyDefault '3'
    Get-LegacyKeyMapEntry -Key 'EHJIR_STALE_CLEANUP_DELAY_SECONDS' -Option 'HybridJoin.StaleCleanupDelaySeconds' -LegacyDefault '30' -Note 'Previous GUI default 30, previous CMD default 60, SEMT default 60.'
    Get-LegacyKeyMapEntry -Key 'EHJIR_REBOOT_DELAY_SECONDS' -Option 'HybridJoin.RebootDelaySeconds' -LegacyDefault '300' -Note 'Previous CMD default 180, GUI and SEMT default 300 (SEMT minimum 60).'
    Get-LegacyKeyMapEntry -Key 'EHJIR_PSEXEC_TIMEOUT_MINUTES' -Option 'PsExecTimeoutMinutes' -LegacyDefault '60' -Note 'Previous GUI default 60, previous CMD default 120, SEMT Hybrid Join default 120.'
    Get-LegacyKeyMapEntry -Key 'EHJIR_CANCELLATION_DRAIN_TIMEOUT_MINUTES' -Option 'CancellationDrainTimeoutMinutes' -LegacyDefault '15'
    Get-LegacyKeyMapEntry -Key 'EHJIR_ALLOW_DSREG_LEAVE' -Option 'HybridJoin.AllowDsregLeave' -LegacyDefault '1'
    Get-LegacyKeyMapEntry -Key 'EHJIR_ALLOW_REMOVE_STALE_INTUNE_ENROLLMENT' -Option 'HybridJoin.AllowRemoveStaleIntuneEnrollment' -LegacyDefault '1'
    Get-LegacyKeyMapEntry -Key 'EHJIR_ALLOW_REBOOT_WHEN_NO_INTERACTIVE_USER' -Option 'HybridJoin.AllowRebootWhenNoInteractiveUser' -LegacyDefault '1'
    Get-LegacyKeyMapEntry -Key 'EHJIR_ALLOW_REBOOT_AFTER_DSREG_LEAVE' -Option 'HybridJoin.AllowRebootAfterDsregLeave' -LegacyDefault '0' -Note 'Previous CMD default was on (reboot even with a connected user); SEMT default is off.'
    Get-LegacyKeyMapEntry -Key 'EHJIR_SKIP_VIRTUAL_MACHINES' -Option 'SkipVirtualMachines' -LegacyDefault '1'
    Get-LegacyKeyMapEntry -Key 'EHJIR_AD_DOMAIN' -Option 'AdDomain'
    Get-LegacyKeyMapEntry -Key 'EHJIR_INTUNE_TENANT_ID' -Option 'IntuneTenantId'
    Get-LegacyKeyMapEntry -Key 'EHJIR_GUI_REMOVE_NON_INTUNE_MDM' -Option 'HybridJoin.AllowRemoveNonIntuneMdmEnrollment' -LegacyDefault '0'
    Get-LegacyKeyMapEntry -Key 'EHJIR_GUI_KEEP_CENTRAL_LOG_HISTORY' -Option 'KeepCentralLogHistory' -LegacyDefault '0'
    Get-LegacyKeyMapEntry -Key 'EHJIR_GUI_NO_CENTRAL_LOG_COLLECTION' -Option 'CentralLogCollection' -LegacyDefault '0' -Invert
    Get-LegacyKeyMapEntry -Key 'EHJIR_GUI_SKIP_POST_CYCLE_INTUNE_INVENTORY' -Option 'HybridJoin.PostCycleCloudInventoryRefresh' -LegacyDefault '0' -Invert
    Get-LegacyKeyMapEntry -Key 'EHJIR_GUI_MAX_CYCLES' -Option 'MaxCycles' -LegacyDefault '0'
    Get-LegacyKeyMapEntry -Key 'EHJIR_USE_TECHNICIAN_RUN_GUARD_HISTORY' -Option 'UseTechnicianRunGuard' -LegacyDefault '1'
    Get-LegacyKeyMapEntry -Key 'EHJIR_TECHNICIAN_RUN_GUARD_HOURS' -Option 'TechnicianRunGuardHours' -LegacyDefault '12'
    Get-LegacyKeyMapEntry -Key 'EHJIR_CENTRAL_LOG_COLLECTION_MODE' -Option 'CentralLogCollectionMode' -LegacyDefault 'Standard'
    Get-LegacyKeyMapEntry -Key 'EHJIR_CONTINUE_ON_DNS_PREFLIGHT_FAILURE' -Option 'DnsPreflightBlocksCycle' -LegacyDefault '0' -Invert
    Get-LegacyKeyMapEntry -Key 'EHJIR_GUI_DRY_RUN' -Option 'DryRun' -Kind LaunchChoice
    Get-LegacyKeyMapEntry -Key 'EHJIR_GUI_AUDIT_ONLY' -Option 'AuditOnly' -Kind LaunchChoice
    Get-LegacyKeyMapEntry -Key 'EHJIR_GUI_IGNORE_RUN_GUARD_EVERY_CYCLE' -Option 'IgnoreEndpointRunGuardEveryCycle' -Kind LaunchChoice
    Get-LegacyKeyMapEntry -Key 'EHJIR_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY' -Kind LaunchChoice -Note 'SEMT: IgnoreEndpointRunGuard also bypasses the technician history for one Run.'
    Get-LegacyKeyMapEntry -Key 'EHJIR_IGNORE_RUN_GUARD' -Option 'IgnoreEndpointRunGuard' -Kind LaunchChoice
    Get-LegacyKeyMapEntry -Key 'EHJIR_RUN_ONCE' -Kind LaunchChoice -Note 'SEMT: Run mode Once or Loop is chosen at launch.'
    Get-LegacyKeyMapEntry -Key 'EHJIR_SKIP_PRE_RUN_ARCHIVE' -Kind NoEquivalent -Note 'SEMT archives older Runs only when RunArchiveAfterDays is above 0 (default 0 = never).'
    Get-LegacyKeyMapEntry -Key 'EHJIR_DISABLE_LOT_RUN_MUTEX' -Kind NoEquivalent -Note 'SEMT always allows one Run per Wave.'
    Get-LegacyKeyMapEntry -Key 'EHJIR_LOT_DIR' -Kind Runtime
    Get-LegacyKeyMapEntry -Key 'EHJIR_RUN_DIR' -Kind Runtime
)

$windows11Map = @(
    Get-LegacyKeyMapEntry -Key 'W11UT_SETUP_SOURCE' -Option 'Windows11Upgrade.SetupSourcePath' -Note 'Several candidates separated by ; stay supported.'
    Get-LegacyKeyMapEntry -Key 'W11UT_SETUP_SOURCE_MAP' -Option 'Windows11Upgrade.SetupSourceMapPath'
    Get-LegacyKeyMapEntry -Key 'W11UT_SETUP_SOURCE_CANDIDATE_LIMIT' -Option 'Windows11Upgrade.SetupSourceCandidateLimit' -LegacyDefault '5'
    Get-LegacyKeyMapEntry -Key 'W11UT_SETUP_EXECUTION_MODE' -Option 'Windows11Upgrade.SetupExecutionMode' -LegacyDefault 'LocalCache'
    Get-LegacyKeyMapEntry -Key 'W11UT_SETUP_MEDIA_ID' -Option 'Windows11Upgrade.SetupMediaId' -LegacyDefault 'Win11' -Note 'SEMT cache: C:\ProgramData\WCH\SmartM365\SmartEndpointMigrationToolkit\Windows11Upgrade\SetupMedia\<MediaId>-<Language>; legacy media is reused or moved by the endpoint.'
    Get-LegacyKeyMapEntry -Key 'W11UT_SETUP_LANGUAGE' -Option 'Windows11Upgrade.SetupLanguage' -LegacyDefault 'MatchSystem'
    Get-LegacyKeyMapEntry -Key 'W11UT_SETUP_DYNAMIC_UPDATE' -Option 'Windows11Upgrade.SetupDynamicUpdate' -LegacyDefault 'Disable'
    Get-LegacyKeyMapEntry -Key 'W11UT_SETUP_COPY_IPG_MS' -Option 'Windows11Upgrade.SetupMediaCopyIpGapMilliseconds' -LegacyDefault '20'
    Get-LegacyKeyMapEntry -Key 'W11UT_SETUP_COPY_JITTER_SECONDS' -Option 'Windows11Upgrade.SetupMediaCopyJitterSeconds' -LegacyDefault '300'
    Get-LegacyKeyMapEntry -Key 'W11UT_SETUP_SOURCE_CONCURRENCY_LIMIT' -Option 'Windows11Upgrade.SetupSourceConcurrencyLimit' -LegacyDefault '0'
    Get-LegacyKeyMapEntry -Key 'W11UT_SETUP_SOURCE_CONCURRENCY_LEASE_MINUTES' -Option 'Windows11Upgrade.SetupSourceConcurrencyLeaseMinutes' -LegacyDefault '240'
    Get-LegacyKeyMapEntry -Key 'W11UT_SETUP_SOURCE_CONCURRENCY_GATE_ROOT' -Option 'Windows11Upgrade.SetupSourceConcurrencyGateRoot'
    Get-LegacyKeyMapEntry -Key 'W11UT_SETUP_SUBNET_CONCURRENCY_LIMIT' -Option 'Windows11Upgrade.SetupSubnetConcurrencyLimit' -LegacyDefault '1'
    Get-LegacyKeyMapEntry -Key 'W11UT_SETUP_SUBNET_PREFIX_LENGTH' -Option 'Windows11Upgrade.SetupSubnetPrefixLength' -LegacyDefault 'Auto'
    Get-LegacyKeyMapEntry -Key 'W11UT_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES' -Option 'Windows11Upgrade.SetupSubnetConcurrencyLeaseMinutes' -LegacyDefault '90'
    Get-LegacyKeyMapEntry -Key 'W11UT_SETUP_SUBNET_CONCURRENCY_GATE_ROOT' -Option 'Windows11Upgrade.SetupSubnetConcurrencyGateRoot'
    Get-LegacyKeyMapEntry -Key 'W11UT_CONFIRM_LOCAL_SETUP_SOURCE' -Option 'Windows11Upgrade.ConfirmLocalSetupSource' -LegacyDefault '0'
    Get-LegacyKeyMapEntry -Key 'W11UT_AD_DOMAIN' -Option 'AdDomain'
    Get-LegacyKeyMapEntry -Key 'W11UT_INTUNE_TENANT_ID' -Option 'IntuneTenantId'
    Get-LegacyKeyMapEntry -Key 'W11UT_INTUNE_INVENTORY_PAGE_SIZE' -Option 'IntuneInventoryPageSize' -LegacyDefault '999'
    Get-LegacyKeyMapEntry -Key 'W11UT_SKIP_INTUNE_INVENTORY_REFRESH' -Option 'Windows11Upgrade.PreCycleIntuneInventoryRefresh' -LegacyDefault '1' -Invert
    Get-LegacyKeyMapEntry -Key 'W11UT_ALLOW_POLICY_REPAIR' -Option 'Windows11Upgrade.AllowPolicyRepair' -LegacyDefault '1'
    Get-LegacyKeyMapEntry -Key 'W11UT_ALLOW_WU_RESET' -Option 'Windows11Upgrade.AllowWUReset' -LegacyDefault '1'
    Get-LegacyKeyMapEntry -Key 'W11UT_ALLOW_FORCE_UPGRADE' -Option 'Windows11Upgrade.AllowForceUpgrade' -LegacyDefault '1'
    Get-LegacyKeyMapEntry -Key 'W11UT_ALLOW_SETUP_UPGRADE' -Option 'Windows11Upgrade.AllowSetupUpgrade' -LegacyDefault '1'
    Get-LegacyKeyMapEntry -Key 'W11UT_DIRECT_SETUP_UPGRADE' -Option 'Windows11Upgrade.DirectSetupUpgrade' -LegacyDefault '0' -Note 'SEMT: direct setup never enables reboots; AllowReboot=false is always honoured.'
    Get-LegacyKeyMapEntry -Key 'W11UT_ALLOW_REBOOT' -Option 'Windows11Upgrade.AllowReboot' -LegacyDefault '1'
    Get-LegacyKeyMapEntry -Key 'W11UT_SCHEDULE_RETRY_AFTER_REBOOT' -Option 'Windows11Upgrade.ScheduleRetryAfterReboot' -LegacyDefault '1'
    Get-LegacyKeyMapEntry -Key 'W11UT_RETRY_AFTER_REBOOT_MAX_ATTEMPTS' -Option 'RetryAfterRebootMaxAttempts' -LegacyDefault '3'
    Get-LegacyKeyMapEntry -Key 'W11UT_RETRY_AFTER_REBOOT_DELAY_SECONDS' -Option 'RetryAfterRebootDelaySeconds' -LegacyDefault '300'
    Get-LegacyKeyMapEntry -Key 'W11UT_FORCE_REQUIRED_REBOOT_WHEN_UPTIME_OVER_DAYS' -Option 'Windows11Upgrade.ForceRequiredRebootWhenUptimeOverDays' -LegacyDefault '7'
    Get-LegacyKeyMapEntry -Key 'W11UT_SETUP_REBOOT_WHEN_NO_USER' -Option 'Windows11Upgrade.AllowSetupCompletionRebootWhenNoUser' -LegacyDefault '1'
    Get-LegacyKeyMapEntry -Key 'W11UT_ALLOW_SETUP_PROFILE_REPAIR' -Option 'Windows11Upgrade.AllowSetupProfileRepair' -LegacyDefault '1'
    Get-LegacyKeyMapEntry -Key 'W11UT_SKIP_VIRTUAL_MACHINES' -Option 'SkipVirtualMachines' -LegacyDefault '1'
    Get-LegacyKeyMapEntry -Key 'W11UT_ALLOW_DISK_CLEANUP' -Option 'Windows11Upgrade.AllowDiskCleanup' -LegacyDefault '1'
    Get-LegacyKeyMapEntry -Key 'W11UT_ALLOW_ADVANCED_DISK_CLEANUP' -Option 'Windows11Upgrade.AllowAdvancedDiskCleanup' -LegacyDefault '0'
    Get-LegacyKeyMapEntry -Key 'W11UT_ALLOW_DISM_COMPONENT_CLEANUP' -Option 'Windows11Upgrade.AllowAdvancedDiskCleanup' -LegacyDefault '0' -Note 'Old alias: 1 enables advanced disk cleanup.'
    Get-LegacyKeyMapEntry -Key 'W11UT_SKIP_SETUP_MEDIA_PRECOPY' -Option 'Windows11Upgrade.SkipSetupMediaPreCopy' -LegacyDefault '0'
    Get-LegacyKeyMapEntry -Key 'W11UT_THROTTLE' -Option 'ThrottleLimit' -LegacyDefault '10'
    Get-LegacyKeyMapEntry -Key 'W11UT_GLOBAL_CONCURRENCY_LIMIT' -Option 'GlobalWorkerLimit' -LegacyDefault '15' -Note 'SEMT has one shared cap for both Campaigns on the technician computer.'
    Get-LegacyKeyMapEntry -Key 'W11UT_GLOBAL_CONCURRENCY_LEASE_TIMEOUT_MINUTES' -Option 'GlobalWorkerLeaseTimeoutMinutes' -LegacyDefault '0'
    Get-LegacyKeyMapEntry -Key 'W11UT_DELAY_BETWEEN_COMPUTERS_SECONDS' -Option 'DelayBetweenComputersSeconds' -LegacyDefault '0'
    Get-LegacyKeyMapEntry -Key 'W11UT_DELAY_BETWEEN_CYCLES_MINUTES' -Option 'DelayBetweenCyclesMinutes' -LegacyDefault '10'
    Get-LegacyKeyMapEntry -Key 'W11UT_PSEXEC_TIMEOUT_MINUTES' -Option 'PsExecTimeoutMinutes' -LegacyDefault '360'
    Get-LegacyKeyMapEntry -Key 'W11UT_CANCELLATION_DRAIN_TIMEOUT_MINUTES' -Option 'CancellationDrainTimeoutMinutes' -LegacyDefault '15'
    Get-LegacyKeyMapEntry -Key 'W11UT_CENTRAL_LOG_COLLECTION_MODE' -Option 'CentralLogCollectionMode' -LegacyDefault 'Standard'
    Get-LegacyKeyMapEntry -Key 'W11UT_GUI_KEEP_CENTRAL_LOG_HISTORY' -Option 'KeepCentralLogHistory' -LegacyDefault '0'
    Get-LegacyKeyMapEntry -Key 'W11UT_GUI_NO_CENTRAL_LOG_COLLECTION' -Option 'CentralLogCollection' -LegacyDefault '0' -Invert
    Get-LegacyKeyMapEntry -Key 'W11UT_GUI_MAX_CYCLES' -Option 'MaxCycles' -LegacyDefault '0'
    Get-LegacyKeyMapEntry -Key 'W11UT_USE_TECHNICIAN_RUN_GUARD_HISTORY' -Option 'UseTechnicianRunGuard' -LegacyDefault '1'
    Get-LegacyKeyMapEntry -Key 'W11UT_RUN_GUARD_HOURS' -Option 'TechnicianRunGuardHours' -LegacyDefault '3' -Note 'Technician history window. The endpoint window is Windows11Upgrade.EndpointRunGuardHours (default 3, not in the previous configuration).'
    Get-LegacyKeyMapEntry -Key 'W11UT_AUDIT_ONLY' -Option 'AuditOnly' -Kind LaunchChoice
    Get-LegacyKeyMapEntry -Key 'W11UT_GUI_DRY_RUN' -Option 'DryRun' -Kind LaunchChoice
    Get-LegacyKeyMapEntry -Key 'W11UT_IGNORE_TECHNICIAN_RUN_GUARD_HISTORY' -Kind LaunchChoice -Note 'SEMT: IgnoreEndpointRunGuard also bypasses the technician history for one Run.'
    Get-LegacyKeyMapEntry -Key 'W11UT_IGNORE_RUN_GUARD' -Option 'IgnoreEndpointRunGuard' -Kind LaunchChoice
    Get-LegacyKeyMapEntry -Key 'W11UT_RUN_ONCE' -Kind LaunchChoice -Note 'SEMT: Run mode Once or Loop is chosen at launch.'
    Get-LegacyKeyMapEntry -Key 'W11UT_ALLOW_SETUP_SOURCE_VALIDATION' -Kind Runtime
    Get-LegacyKeyMapEntry -Key 'W11UT_LOT_DIR' -Kind Runtime
    Get-LegacyKeyMapEntry -Key 'W11UT_RUN_DIR' -Kind Runtime
)

$legacySources = @(
    [pscustomobject]@{ Campaign = 'HybridJoin'; Prefix = 'EHJIR_'; FolderName = 'IntuneHybridJoinToolkit'; ConfigName = 'IntuneHybridJoinToolkit.config'; Map = $hybridJoinMap }
    [pscustomobject]@{ Campaign = 'Windows11Upgrade'; Prefix = 'W11UT_'; FolderName = 'Windows11UpgradeToolkit'; ConfigName = 'Windows11UpgradeToolkit.config'; Map = $windows11Map }
)

#endregion

#region Helpers

function Read-LegacyConfigFile {
    <# Same syntax as the previous GUIs: NAME=value, # and ; comments, optional quotes; last value wins. #>
    param([Parameter(Mandatory = $true)][string]$Path,[Parameter(Mandatory = $true)][string]$Prefix)

    $values = [ordered]@{}
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $values }
    foreach ($rawLine in [System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::UTF8)) {
        $line = ([string]$rawLine).Trim()
        if ([string]::IsNullOrWhiteSpace($line) -or $line.StartsWith('#') -or $line.StartsWith(';')) { continue }
        $match = [regex]::Match($line, '^(?<name>[A-Za-z_][A-Za-z0-9_]*)=(?<value>.*)$')
        if (-not $match.Success) { continue }
        $name = $match.Groups['name'].Value.Trim().ToUpperInvariant()
        if (-not $name.StartsWith($Prefix)) { continue }
        $value = $match.Groups['value'].Value.Trim()
        if ($value.Length -ge 2 -and $value.StartsWith('"') -and $value.EndsWith('"')) { $value = $value.Substring(1, $value.Length - 2) }
        if ($values.Contains($name)) { $script:warnings.Add(("{0} is defined more than once in {1}; the last value is used, as in the previous GUI." -f $name,$Path)) }
        $values[$name] = $value
    }
    return $values
}

function Read-LegacyEnvironment {
    param([Parameter(Mandatory = $true)][string]$Prefix,[string[]]$Scopes)

    $values = [ordered]@{}
    foreach ($scope in @($Scopes)) {
        $variables = [Environment]::GetEnvironmentVariables($scope)
        foreach ($name in @($variables.Keys)) {
            $upper = ([string]$name).ToUpperInvariant()
            if (-not $upper.StartsWith($Prefix)) { continue }
            $values[$upper] = [pscustomobject]@{ Value = ([string]$variables[$name]).Trim(); Source = "Environment:$scope" }
        }
    }
    return $values
}

function ConvertFrom-LegacyBoolean {
    param([string]$Text)
    $value = ([string]$Text).Trim()
    if ($value -in @('1','true','yes','on')) { return $true }
    if ($value -in @('0','false','no','off')) { return $false }
    throw ("expects 1 or 0; received '{0}'" -f $value)
}

function Test-SameValue {
    param([AllowNull()]$Left,[AllowNull()]$Right)
    if ($null -eq $Left -or $null -eq $Right) { return ($null -eq $Left -and $null -eq $Right) }
    return ([string]$Left -eq [string]$Right)
}

function Format-ReportValue {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [bool]) { if ($Value) { return 'true' } return 'false' }
    return [string]$Value
}

function Get-ExplicitOptionValue {
    <# Returns the normalized explicit value of an option in an Options dictionary, or $null when it inherits. #>
    param([Parameter(Mandatory = $true)][System.Collections.IDictionary]$Options,[Parameter(Mandatory = $true)][string]$Name,[string]$FileLabel)

    if (-not $Options.Contains($Name)) { return $null }
    $raw = $Options[$Name]
    if (Test-SemtInheritValue -Value $raw) { return $null }
    try { return (ConvertTo-SemtOptionValue -Definition $catalog.ByName[$Name] -Value $raw -Source $FileLabel) }
    catch {
        $script:warnings.Add(("{0}: existing value of {1} is invalid ({2}); it is treated as different." -f $FileLabel,$Name,$_.Exception.Message))
        return [string]$raw
    }
}

function Test-PathUnderFolder {
    param([string]$Path,[string]$Folder)
    if ([string]::IsNullOrWhiteSpace($Path) -or [string]::IsNullOrWhiteSpace($Folder)) { return $false }
    foreach ($item in @($Path -split ';')) {
        $candidate = $item.Trim().TrimEnd('\')
        if ($candidate -eq $Folder.TrimEnd('\') -or $candidate.StartsWith($Folder.TrimEnd('\') + '\', [System.StringComparison]::OrdinalIgnoreCase)) { return $true }
    }
    return $false
}

function Read-SemtTargetFile {
    <# Reads a SEMT JSON file without writing it; missing keys of the template are added in memory. #>
    param([Parameter(Mandatory = $true)][string]$Path,[string]$TemplatePath,[Parameter(Mandatory = $true)][System.Collections.IDictionary]$NewDocument)

    $exists = Test-Path -LiteralPath $Path -PathType Leaf
    if ($exists) { $document = ConvertTo-SemtHashtable -InputObject (Read-SemtJsonFile -Path $Path -Required) }
    else { $document = $NewDocument }
    if ($null -eq $document) { $document = [ordered]@{} }
    if (-not [string]::IsNullOrWhiteSpace($TemplatePath)) {
        $template = ConvertTo-SemtHashtable -InputObject (Read-SemtJsonFile -Path $TemplatePath -Required)
        $added = New-Object 'System.Collections.Generic.List[string]'
        Add-SemtMissingTemplateKeys -Target $document -Template $template -AddedKeys $added
    }
    if (-not $document.Contains('Options') -or -not ($document['Options'] -is [System.Collections.IDictionary])) { $document['Options'] = [ordered]@{} }
    return [pscustomobject]@{ Path = $Path; Exists = $exists; Document = $document; Changed = $false }
}

#endregion

#region Read the previous toolkits (read-only)

if (-not (Test-Path -LiteralPath $LegacyDevicesRoot -PathType Container)) { throw "Legacy devices folder not found: $LegacyDevicesRoot" }
$legacyRoot = (Resolve-Path -LiteralPath $LegacyDevicesRoot).ProviderPath
$rows = New-Object 'System.Collections.Generic.List[object]'
$sourceSummaries = New-Object 'System.Collections.Generic.List[object]'
$legacyWindows11Folder = Join-Path -Path $legacyRoot -ChildPath 'Windows11UpgradeToolkit'

foreach ($legacy in $legacySources) {
    $folder = Join-Path -Path $legacyRoot -ChildPath $legacy.FolderName
    $configPath = Join-Path -Path $folder -ChildPath $legacy.ConfigName
    $folderExists = Test-Path -LiteralPath $folder -PathType Container
    $configValues = Read-LegacyConfigFile -Path $configPath -Prefix $legacy.Prefix
    $environmentValues = [ordered]@{}
    if (-not $IgnoreLegacyEnvironment) { $environmentValues = Read-LegacyEnvironment -Prefix $legacy.Prefix -Scopes $LegacyEnvironmentScope }

    # LOT-level values are not converted (decision: LOTs and logs are not migrated); count them only.
    $lotCount = 0; $lotOverrideCount = 0
    $lotsRoot = Join-Path -Path $folder -ChildPath 'Lots'
    if (Test-Path -LiteralPath $lotsRoot -PathType Container) {
        foreach ($lot in @(Get-ChildItem -LiteralPath $lotsRoot -Directory -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne 'LOT-TEMPLATE' })) {
            $lotCount++
            $lotConfig = Join-Path -Path $lot.FullName -ChildPath $legacy.ConfigName
            $lotDomain = Join-Path -Path $lot.FullName -ChildPath 'AdDomain.txt'
            $hasOverride = (Test-Path -LiteralPath $lotConfig -PathType Leaf) -and @((Read-LegacyConfigFile -Path $lotConfig -Prefix $legacy.Prefix).Values | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) }).Count -gt 0
            if (-not $hasOverride -and (Test-Path -LiteralPath $lotDomain -PathType Leaf)) {
                $hasOverride = @([System.IO.File]::ReadAllLines($lotDomain) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.Trim().StartsWith('#') }).Count -gt 0
            }
            if ($hasOverride) { $lotOverrideCount++ }
        }
    }
    if ($lotOverrideCount -gt 0) {
        $warnings.Add(("{0}: {1} LOT folder(s) carry their own values (LOT config or AdDomain.txt). LOTs are not converted: set these values on the new Waves if still needed." -f $legacy.FolderName,$lotOverrideCount))
    }

    $sourceSummaries.Add([pscustomobject]@{
        Campaign = $legacy.Campaign
        Folder = $folder
        FolderFound = $folderExists
        ConfigPath = $configPath
        ConfigFound = (Test-Path -LiteralPath $configPath -PathType Leaf)
        ConfigValueCount = $configValues.Count
        EnvironmentValueCount = $environmentValues.Count
        LotCount = $lotCount
        LotWithOwnValuesCount = $lotOverrideCount
    })

    $merged = [ordered]@{}
    foreach ($key in @($configValues.Keys)) { $merged[$key] = [pscustomobject]@{ Value = [string]$configValues[$key]; Source = 'Config' } }
    foreach ($key in @($environmentValues.Keys)) { $merged[$key] = $environmentValues[$key] }

    $mapByKey = @{}
    foreach ($entry in $legacy.Map) { $mapByKey[$entry.Key] = $entry }

    foreach ($key in @($merged.Keys)) {
        $item = $merged[$key]
        $row = [pscustomobject]@{
            Campaign = $legacy.Campaign; LegacyKey = $key; LegacySource = $item.Source; LegacyValue = $item.Value
            SemtOption = ''; SemtValue = $null; SemtDefault = $null; Decision = ''; Target = ''; Note = ''
        }
        if (-not $mapByKey.ContainsKey($key)) {
            $row.Decision = 'Unknown'; $row.Note = 'Not a known setting of the previous toolkit; not converted.'
            $rows.Add($row); continue
        }
        $map = $mapByKey[$key]
        $row.SemtOption = $map.Option
        $row.Note = $map.Note
        if ($map.Kind -ne 'Map') {
            $row.Decision = $map.Kind
            if ($map.Kind -eq 'LaunchChoice' -and $item.Value -notin @('','0','false')) {
                $warnings.Add(("{0}={1} ({2}) is a launch-time choice and is not carried over; choose it in the GUI when launching." -f $key,$item.Value,$item.Source))
            }
            $rows.Add($row); continue
        }

        $definition = $catalog.ByName[$map.Option]
        $row.SemtDefault = Get-SemtOptionDefault -Definition $definition -Campaign $legacy.Campaign
        if ([string]::IsNullOrWhiteSpace($item.Value)) { $row.Decision = 'Empty'; $rows.Add($row); continue }
        if ($key -eq 'W11UT_SETUP_SOURCE' -and $item.Value -eq '\\server\share\Windows11') { $row.Decision = 'Placeholder'; $row.Note = 'GUI placeholder, not a real source.'; $rows.Add($row); continue }
        if ($key -eq 'W11UT_ALLOW_DISM_COMPONENT_CLEANUP') {
            # Old alias: only 1 has an effect (it forces advanced cleanup on).
            try { $aliasOn = ConvertFrom-LegacyBoolean -Text $item.Value } catch { $aliasOn = $false }
            if (-not $aliasOn) { $row.Decision = 'Runtime'; $row.Note = 'Old alias off: no effect.'; $rows.Add($row); continue }
        }
        try {
            $raw = $item.Value
            if ($map.Invert) { $raw = -not (ConvertFrom-LegacyBoolean -Text $item.Value) }
            $row.SemtValue = ConvertTo-SemtOptionValue -Definition $definition -Value $raw -Source $key
        }
        catch {
            $row.Decision = 'Invalid'; $row.Note = $_.Exception.Message
            $warnings.Add(("{0}={1} cannot be converted: {2}" -f $key,$item.Value,$_.Exception.Message))
            $rows.Add($row); continue
        }
        if ($item.Source -eq 'Config' -and -not $IncludeLegacyDefaults -and $map.LegacyDefault -ne '' -and $item.Value -eq $map.LegacyDefault) {
            $row.Decision = 'LegacyDefault'
        }
        elseif (Test-SameValue -Left $row.SemtValue -Right (ConvertTo-SemtOptionValue -Definition $definition -Value $row.SemtDefault -Source 'Default')) {
            $row.Decision = 'SameAsSemtDefault'
        }
        else { $row.Decision = 'Candidate' }
        $rows.Add($row)
    }
}

if (@($sourceSummaries | Where-Object { $_.FolderFound }).Count -eq 0) {
    throw ("Neither IntuneHybridJoinToolkit nor Windows11UpgradeToolkit was found under {0}." -f $legacyRoot)
}

# The old alias wins over W11UT_ALLOW_ADVANCED_DISK_CLEANUP, as in the previous CMD wrapper.
$aliasRows = @($rows | Where-Object { $_.LegacyKey -eq 'W11UT_ALLOW_DISM_COMPONENT_CLEANUP' -and $_.Decision -in @('Candidate','SameAsSemtDefault','LegacyDefault') })
if ($aliasRows.Count -gt 0) {
    foreach ($row in @($rows | Where-Object { $_.LegacyKey -eq 'W11UT_ALLOW_ADVANCED_DISK_CLEANUP' })) { $row.Decision = 'Superseded'; $row.Note = 'W11UT_ALLOW_DISM_COMPONENT_CLEANUP=1 forces advanced cleanup on.' }
}

#endregion

#region Plan: global configuration or per-Campaign GUI state

$configRoot = Join-Path -Path $PackageRoot -ChildPath 'Config'
$globalFile = Read-SemtTargetFile -Path (Join-Path $configRoot 'SmartEndpointMigrationToolkit.local.json') -TemplatePath (Join-Path $configRoot 'SmartEndpointMigrationToolkit.local.json.template') -NewDocument ([ordered]@{})
$guiFiles = @{}
foreach ($campaign in $campaigns) {
    $guiFiles[$campaign] = Read-SemtTargetFile -Path (Join-Path $configRoot ("GuiState.{0}.local.json" -f $campaign)) -NewDocument ([ordered]@{ SchemaVersion = 1; Campaign = $campaign; UpdatedUtc = ''; Options = [ordered]@{} })
}

$plan = New-Object 'System.Collections.Generic.List[object]'
$optionNames = @($rows | Where-Object { $_.Decision -eq 'Candidate' } | ForEach-Object { $_.SemtOption } | Sort-Object -Unique)
foreach ($optionName in $optionNames) {
    $definition = $catalog.ByName[$optionName]
    $candidates = @($rows | Where-Object { $_.Decision -eq 'Candidate' -and $_.SemtOption -eq $optionName })
    if ([string]$definition.Scope -ne 'Common') {
        # Campaign options are ignored by the other Campaign: the global file is safe.
        $row = $candidates[-1]
        $plan.Add([pscustomobject]@{ Target = 'Global'; Option = $optionName; Value = $row.SemtValue; Rows = @($candidates) })
        continue
    }
    $campaignsWithValue = @($candidates | ForEach-Object { $_.Campaign } | Sort-Object -Unique)
    if ($campaignsWithValue.Count -eq 2) {
        $hj = @($candidates | Where-Object { $_.Campaign -eq 'HybridJoin' })[-1]
        $w11 = @($candidates | Where-Object { $_.Campaign -eq 'Windows11Upgrade' })[-1]
        if (Test-SameValue -Left $hj.SemtValue -Right $w11.SemtValue) {
            $plan.Add([pscustomobject]@{ Target = 'Global'; Option = $optionName; Value = $hj.SemtValue; Rows = @($candidates) })
        }
        else {
            foreach ($row in @($hj,$w11)) { $plan.Add([pscustomobject]@{ Target = "GuiState.$($row.Campaign)"; Option = $optionName; Value = $row.SemtValue; Rows = @($row) }) }
        }
        continue
    }
    # One Campaign carries a value. The other keeps its own previous value (even a previous default)
    # or, when it has none, its SEMT default: only a campaign-specific SEMT default must be protected.
    $row = $candidates[-1]
    $other = @($campaigns | Where-Object { $_ -ne $row.Campaign })[0]
    $otherRows = @($rows | Where-Object { $_.Campaign -eq $other -and $_.SemtOption -eq $optionName -and $null -ne $_.SemtValue })
    $hasCampaignDefaults = $null -ne (Get-SemtPropertyValue -InputObject $definition -Name 'CampaignDefaults' -Default $null)
    $otherDefault = ConvertTo-SemtOptionValue -Definition $definition -Value (Get-SemtOptionDefault -Definition $definition -Campaign $other) -Source 'Default'
    $sharedIsSafe = if ($otherRows.Count -gt 0) { Test-SameValue -Left $otherRows[-1].SemtValue -Right $row.SemtValue } else { (-not $hasCampaignDefaults) -or (Test-SameValue -Left $otherDefault -Right $row.SemtValue) }
    if ($sharedIsSafe) {
        $plan.Add([pscustomobject]@{ Target = 'Global'; Option = $optionName; Value = $row.SemtValue; Rows = @($candidates) })
    }
    else {
        $plan.Add([pscustomobject]@{ Target = "GuiState.$($row.Campaign)"; Option = $optionName; Value = $row.SemtValue; Rows = @($row) })
    }
}

foreach ($entry in $plan) {
    if ($entry.Target -eq 'Global') { $file = $globalFile; $label = 'SmartEndpointMigrationToolkit.local.json' }
    else { $campaign = $entry.Target.Substring('GuiState.'.Length); $file = $guiFiles[$campaign]; $label = "GuiState.$campaign.local.json" }
    $options = $file.Document['Options']
    $existing = Get-ExplicitOptionValue -Options $options -Name $entry.Option -FileLabel $label
    $decision = if ($entry.Target -eq 'Global') { 'ToGlobal' } else { 'ToGuiState' }
    $note = ''
    if ($null -ne $existing -and (Test-SameValue -Left $existing -Right $entry.Value)) { $decision = 'AlreadySet' }
    elseif ($null -ne $existing -and -not $ReplaceExistingValues) {
        $decision = 'KeptExisting'
        $note = ("{0} already has {1}; kept (use -ReplaceExistingValues to replace)." -f $label,(Format-ReportValue $existing))
        $warnings.Add(("{0}: {1}" -f $entry.Option,$note))
    }
    else {
        $options[$entry.Option] = $entry.Value
        $file.Changed = $true
        if ($null -ne $existing) { $note = ("replaces {0}" -f (Format-ReportValue $existing)) }
    }
    foreach ($row in @($entry.Rows)) {
        $row.Decision = $decision
        $row.Target = $label
        if ($note) { $row.Note = (@($row.Note,$note) | Where-Object { $_ }) -join ' ' }
    }
}

# A saved GUI state wins over the global file in the GUI: report the values that would hide a global one.
foreach ($entry in @($plan | Where-Object { $_.Target -eq 'Global' })) {
    foreach ($campaign in $campaigns) {
        if ($catalog.ByName[$entry.Option].Scope -notin @('Common',$campaign)) { continue }
        $state = $guiFiles[$campaign]
        if (-not $state.Exists) { continue }
        $saved = Get-ExplicitOptionValue -Options $state.Document['Options'] -Name $entry.Option -FileLabel ("GuiState.{0}.local.json" -f $campaign)
        if ($null -ne $saved -and -not (Test-SameValue -Left $saved -Right $entry.Value)) {
            $warnings.Add(("{0}: the existing GuiState.{1}.local.json shows {2} in the GUI instead of the converted global value {3}." -f $entry.Option,$campaign,(Format-ReportValue $saved),(Format-ReportValue $entry.Value)))
        }
    }
}

if (@($plan | Where-Object { $_.Target -like 'GuiState.*' }).Count -gt 0) {
    $warnings.Add('Values written to GuiState.<Campaign>.local.json apply to GUI launches only (as the previous GUI options did). Start-Wave-*.cmd launchers use the global file and Wave.local.json: set these values on the Wave when you launch from CMD.')
}

#endregion

#region Semantic checks on the planned Windows 11 setup values

$w11Effective = $null
try {
    foreach ($campaign in $campaigns) {
        $resolved = Resolve-SemtEffectiveOptions -Catalog $catalog -Campaign $campaign -GlobalLayer $globalFile.Document -GuiLayer $guiFiles[$campaign].Document
        if ($campaign -eq 'Windows11Upgrade') { $w11Effective = $resolved }
    }
}
catch { throw ("The planned configuration is rejected by SEMT, nothing written: {0}" -f $_.Exception.Message) }

if ($null -ne $w11Effective) {
    $setupSource = [string]$w11Effective['Windows11Upgrade.SetupSourcePath'].Value
    $setupMap = [string]$w11Effective['Windows11Upgrade.SetupSourceMapPath'].Value
    $confirmLocal = [bool]$w11Effective['Windows11Upgrade.ConfirmLocalSetupSource'].Value
    foreach ($pair in @(@('Windows11Upgrade.SetupSourcePath',$setupSource),@('Windows11Upgrade.SetupSourceMapPath',$setupMap),@('Windows11Upgrade.SetupSourceConcurrencyGateRoot',[string]$w11Effective['Windows11Upgrade.SetupSourceConcurrencyGateRoot'].Value),@('Windows11Upgrade.SetupSubnetConcurrencyGateRoot',[string]$w11Effective['Windows11Upgrade.SetupSubnetConcurrencyGateRoot'].Value))) {
        $name = $pair[0]; $value = $pair[1]
        if ([string]::IsNullOrWhiteSpace($value)) { continue }
        if (Test-PathUnderFolder -Path $value -Folder $legacyWindows11Folder) {
            $warnings.Add(("{0} points into the previous Windows 11 toolkit folder ({1}). Point it to a UNC share, or to the SEMT folder for a local test." -f $name,$value))
        }
        foreach ($item in @($value -split ';')) {
            if (-not [string]::IsNullOrWhiteSpace($item) -and -not $item.Trim().StartsWith('\\') -and -not ($name -like '*SetupSource*Path' -and $confirmLocal)) {
                $warnings.Add(("{0} is not a UNC path ({1}). Targets run as SYSTEM and must read it themselves; a Wave launch is refused unless Windows11Upgrade.ConfirmLocalSetupSource is on (local test only)." -f $name,$item.Trim()))
            }
        }
    }
    if ([bool]$w11Effective['Windows11Upgrade.DirectSetupUpgrade'].Value -and -not [bool]$w11Effective['Windows11Upgrade.AllowReboot'].Value) {
        $warnings.Add('DirectSetupUpgrade is on and AllowReboot is off: the previous CMD wrapper forced reboots in this case; SEMT does not reboot.')
    }
}

#endregion

#region Output, write and report

$report = [ordered]@{
    Product = 'SmartEndpointMigrationToolkit'
    Tool = 'Import-LegacyConfig'
    ToolVersion = $scriptVersion
    GeneratedUtc = Get-SemtUtcNowText
    Mode = $(if ($Apply) { 'Apply' } else { 'Preview' })
    LegacyDevicesRoot = $legacyRoot
    PackageRoot = $PackageRoot
    IncludeLegacyDefaults = [bool]$IncludeLegacyDefaults
    ReplaceExistingValues = [bool]$ReplaceExistingValues
    Sources = $sourceSummaries.ToArray()
    Rows = @($rows.ToArray() | ForEach-Object { [ordered]@{ Campaign = $_.Campaign; LegacyKey = $_.LegacyKey; LegacySource = $_.LegacySource; LegacyValue = $_.LegacyValue; SemtOption = $_.SemtOption; SemtValue = (Format-ReportValue $_.SemtValue); SemtDefault = (Format-ReportValue $_.SemtDefault); Decision = $_.Decision; Target = $_.Target; Note = $_.Note } })
    Writes = @()
    Warnings = $warnings.ToArray()
}

Write-Host ''
Write-Host ("Smart Endpoint Migration Toolkit - legacy configuration import {0} ({1})" -f $scriptVersion,$report.Mode)
Write-Host ("Previous toolkits : {0}" -f $legacyRoot)
Write-Host ("SEMT package      : {0}" -f $PackageRoot)
foreach ($summary in $sourceSummaries) {
    Write-Host ("  {0,-17} folder={1} config={2} values={3} environment={4} LOTs={5} (not converted)" -f $summary.Campaign,$summary.FolderFound,$summary.ConfigFound,$summary.ConfigValueCount,$summary.EnvironmentValueCount,$summary.LotCount)
}
$order = @{ ToGlobal = 0; ToGuiState = 1; KeptExisting = 2; AlreadySet = 3; Invalid = 4; LaunchChoice = 5; NoEquivalent = 6; Unknown = 7; SameAsSemtDefault = 8; LegacyDefault = 9; Superseded = 10; Empty = 11; Placeholder = 12; Runtime = 13 }
$table = $rows.ToArray() | Sort-Object @{ Expression = { $order[$_.Decision] } },Campaign,LegacyKey | Select-Object Campaign,LegacyKey,LegacyValue,SemtOption,@{ Name = 'SemtValue'; Expression = { Format-ReportValue $_.SemtValue } },Decision,Target
Write-Host (($table | Format-Table -AutoSize | Out-String -Width 400).TrimEnd())
Write-Host ''
Write-Host ('Decisions: ToGlobal/ToGuiState = written by -Apply; SameAsSemtDefault/LegacyDefault = SEMT default kept (use -IncludeLegacyDefaults to carry previous GUI defaults); KeptExisting = SEMT value kept.')
foreach ($warning in $warnings) { Write-Warning $warning }

$filesToWrite = @(@($globalFile) + @($campaigns | ForEach-Object { $guiFiles[$_] }) | Where-Object { $_.Changed })
if (-not $Apply) {
    Write-Host ''
    if ($filesToWrite.Count -eq 0) { Write-Host 'Preview: nothing to write.' }
    else {
        Write-Host 'Preview: nothing written. -Apply would write:'
        foreach ($file in $filesToWrite) { Write-Host ("  {0}{1}" -f $file.Path,$(if ($file.Exists) { ' (backup first)' } else { ' (new)' })) }
    }
}
else {
    $writes = New-Object 'System.Collections.Generic.List[object]'
    foreach ($file in $filesToWrite) {
        if (-not $PSCmdlet.ShouldProcess($file.Path, 'Write converted SEMT configuration')) { continue }
        $backup = ''
        if ($file.Exists) {
            $backup = '{0}.{1}.bak' -f $file.Path,$stamp
            $index = 1
            while (Test-Path -LiteralPath $backup) { $backup = '{0}.{1}-{2}.bak' -f $file.Path,$stamp,$index; $index++ }
            Copy-Item -LiteralPath $file.Path -Destination $backup -ErrorAction Stop
        }
        if ($file.Document.Contains('UpdatedUtc')) { $file.Document['UpdatedUtc'] = Get-SemtUtcNowText }
        Write-SemtJsonFileAtomic -Path $file.Path -InputObject $file.Document
        $written = ConvertTo-SemtHashtable -InputObject (Read-SemtJsonFile -Path $file.Path -Required)
        foreach ($key in @($file.Document['Options'].Keys)) {
            if (-not (Test-SameValue -Left $written['Options'][$key] -Right $file.Document['Options'][$key])) { throw ("Verification failed after writing {0}: option {1}." -f $file.Path,$key) }
        }
        $writes.Add([ordered]@{ Path = $file.Path; Backup = $backup; Created = (-not $file.Exists) })
        Write-Host ("Written: {0}{1}" -f $file.Path,$(if ($backup) { " (backup: $backup)" } else { '' }))
    }
    $report.Writes = $writes.ToArray()
    if ($filesToWrite.Count -eq 0) { Write-Host 'Nothing to write: the SEMT configuration already matches.' }
}

$reportTargets = New-Object 'System.Collections.Generic.List[string]'
if ($Apply -and -not $WhatIfPreference) { $reportTargets.Add((Join-Path $configRoot ("LegacyConfigImport.{0}.local.json" -f $stamp))) }
if (-not [string]::IsNullOrWhiteSpace($ReportPath)) { $reportTargets.Add($ReportPath) }
foreach ($target in $reportTargets) {
    Write-SemtJsonFileAtomic -Path $target -InputObject $report
    Write-Host ("Report: {0}" -f $target)
}

if ($PassThru) { [pscustomobject]$report }

#endregion

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAFyH7rKfvrdDXH
# FsHDfGZAS2smVOmxZP7gFtkL0tKDIKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIkl2V8jzMatGUYPFc15yrz5d9+l9v8UVdtZiSb6udgsMA0GCSqG
# SIb3DQEBAQUABIIBgFx1K81jdS7QymgLSpZq02XQPHNcirkwXBxq6G0AryEYTMJJ
# 0VKEF4prK33xmsnQ4g3dfBWZW3f3gMlFicrkoP9EdvjxUxJu58YCPC+W53lWZ0+Q
# MedhD7c3Gq0M22KOyAABttcvCFABFP0+O5mY+hcyYZ6JKyqXj6ioxYpLAlcIdU8m
# wehoPYqPtj4F1RBTPWqJNA+wTD2o9YkRoXYQxUo95r2ebyOAZ0GgmaHVVccGtvkh
# yDTj1Hw6TeHLBCVZuzxWNIWOzc7+X6jaUiL8K8OLwIe/UJDgkLAI4wXhXuxrLv7G
# DmD4THrCLMnNbilUaCDGwNLDEAeuriff/SwXvrX0LJr6UJAFnr24WWZMZ3VzMyAD
# 7cNuIpI+7aiovKADAB4K3PEEUZIY1WR4uik5G416aej6RYMaQ/StPEpKeydptdaZ
# +32ZF8sPtY6OGvkEVE+DhF9SQvp7NmRD4ZkpgEqlN6NHes7f6SB6YZw2HUH6o4cr
# MrdM4mnpGZ2qrSa++KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgxNDM5
# MDhaMC8GCSqGSIb3DQEJBDEiBCCLQlEATcSIV/+KewITr6bxhPNACZtjWxHL5zSC
# e9i3ETANBgkqhkiG9w0BAQEFAASCAgCnGjEzcYdf918DrL1heXODkxGfW0V982R4
# C0ZcWxdf4pdks/NECjo5l1W/6Bxsmkg65MUv0ZoX4idsrdr7FEa7kQKeBcrMfShE
# 83oehodrtZEYN/h+8uZSj4TXqSH2E2+xWNkAoenyFr45Z01X4dSR0uMdqjUrucoW
# ajdZ10FgBn11ilH8btJXZF4Cq7SxN9G9PAeF3cTRv0DPU5AGQutR3/+WeRrdu6ik
# tKdI3BmUCwbxKqrm7k95mbCZeiUJdk755zCwA9u3lmXWTh1lr6PBrPHudY0VdcWR
# orixGfrN5UHD6gIrLcqQf5GX61+WzZ6k5PFwZsg22x9kKEiVHQlLXL3Lna6euXsD
# rzbhtMRTfK/jqG+J7x79RoL1KhBmdG4SV6t8sPUjkvqpIYcl2ZGOK1iP5W72bBXI
# 7bIuv/Yey/KS/ZosMJmWuJaepxdTsDGps5IUfT92OD4A9hUfs9MVHoaVi0tWc26Z
# 9NmviAmD9qwCmYpltT3qEsHFowsPm7kmXA8VRJidFUPmPipc2HBcBzyCaTigB4Zs
# hnFSkdrfW4AduGKq1ws8VrVoRzT+p5Y4EjKJReDzbvpQvtddx/7V8hRHiFf9+4Cu
# jMxWeiqfNrwWqekiXfhg8ZIuj5kmtx3KtKRVv1I2AmAISlLCUSlFnQZlgddWOeaZ
# mkj7g5lcUg==
# SIG # End signature block
