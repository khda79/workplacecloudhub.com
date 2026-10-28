# Options

Reference: [Config/SmartEndpointMigrationToolkit.options.json.txt](../Config/SmartEndpointMigrationToolkit.options.json.txt).
The table at the end is generated from the catalog; the catalog wins if they differ.

## Precedence

Highest first:

1. Explicit CLI (`-Option @{ Name = Value }`, `-DryRun`, `-AuditOnly`, `-IgnoreRunGuard`)
2. GUI choices captured for the Run (`Runs\<Campaign>\_GuiSelections\*.json`, Campaign checked)
3. Wave (`Waves\<Campaign>\<Wave>\Wave.local.json`)
4. `SEMT_*` environment variables (`SEMT_*` common, `SEMT_HYBRIDJOIN_*`, `SEMT_WINDOWS11_*`)
5. Global configuration (`Config\SmartEndpointMigrationToolkit.local.json`)
6. Campaign default (catalog `CampaignDefaults`, else `Default`)

Then rules: `DryRunSingleCycle` (dry run with unlimited cycles runs one cycle), `RunModeOnce`
(Once = one cycle), `SinglePcOneWorker` (Single PC uses one worker).

In the global, Wave and environment layers, these values mean "inherit from the lower layer":
empty, `null`, `__INHERIT__`, `__USE_GLOBAL__`, `USE_GLOBAL`, `__USE_DEFAULT__`. Values are
validated against type, range and choices; an invalid or unknown option stops the Run with the
layer that supplied it. An option of the other Campaign is ignored.

Every Run records `EffectiveOptions.json` (value and source of each option) and the report shows
the same table. The GUI keeps its last choices per Campaign in `Config\GuiState.<Campaign>.local.json`.

## Default decisions (D11)

Defaults were taken from the real behaviour of the two previous toolkits, then reviewed:

| Option | Default | Why |
| --- | --- | --- |
| `HybridJoin.AllowRebootAfterDsregLeave` | false | The previous Hybrid Join toolkit rebooted after `dsregcmd /leave` even with a user connected. A reboot now requires an explicit choice. |
| `HybridJoin.RebootDelaySeconds` | 300 | Same as before; gives a connected user time to save. |
| `HybridJoin.StaleCleanupDelaySeconds` | 60 | Same as before. |
| `HybridJoin.IntuneRetrySleepMinutes` | 5 | Same as before. |
| `ThrottleLimit` (Hybrid Join) | 10 | Same as before. |
| `DelayBetweenCyclesMinutes` | HJ 5, W11 10 | Same as before for each product. |
| `NightPauseEnabled` | HJ on, W11 off | Hybrid Join paused at night before; Windows 11 never did (upgrades are often scheduled at night). |
| `PsExecTimeoutMinutes` | HJ 120, W11 360 | Same as before; Windows 11 needs longer for the media copy and setup. |
| `TechnicianRunGuardHours` | HJ 12, W11 3 | Same as before for each product; Windows 11 also keeps its endpoint run guard (`Windows11Upgrade.EndpointRunGuardHours` 3). |
| `Windows11Upgrade.DirectSetupUpgrade` | false | Direct setup never implies a reboot: with `AllowReboot` off the endpoint receives no `-AllowReboot` (the previous CMD wrapper forced it, see [CRITICAL-FIXES-LEGACY.md](CRITICAL-FIXES-LEGACY.md)). |
| `Windows11Upgrade.ReuseLegacySetupMedia` | true | Existing media of the previous toolkit is reused or moved before any large copy. |
| `Windows11Upgrade.AllowSetupMediaDuplication` | false | No silent second copy of the setup media. |
| `Windows11Upgrade.AllowReboot` | true | Same as before: pending-reboot handling and setup completion reboots when no interactive user is connected. It also enables `ForceRequiredRebootWhenUptimeOverDays` (7 days), which reboots even with a connected user; set that option to 0 to never force. Turning `AllowReboot` off is always honoured, including with DirectSetup. |
| `GlobalWorkerLimit` | 15 | Shared by both Campaigns on the technician computer; 0 disables the shared cap. |
## Option reference

Generated from catalog SHA256 `A587713B499D1D85BB1B3A3352BB1BDCD41D8CD26CD83A0AD88ED7ADB0ECA331`.

### Common options

| Option | Type | Default | Range / choices | Env var | Category | GUI |
| --- | --- | --- | --- | --- | --- | --- |
| `DryRun` | Boolean | False |  | `SEMT_DRY_RUN` | Mode | yes |
| `AuditOnly` | Boolean | False |  | `SEMT_AUDIT_ONLY` | Mode | yes |
| `IgnoreEndpointRunGuard` | Boolean | False |  | `SEMT_IGNORE_ENDPOINT_RUN_GUARD` | Mode | yes |
| `IgnoreEndpointRunGuardEveryCycle` | Boolean | False |  | `SEMT_IGNORE_ENDPOINT_RUN_GUARD_EVERY_CYCLE` | Mode | no |
| `MaxCycles` | Integer | 0 cycles | 0..1000 | `SEMT_MAX_CYCLES` | Schedule | yes |
| `DelayBetweenCyclesMinutes` | Integer | HJ: 5; W11: 10 minutes | 0..1440 | `SEMT_DELAY_BETWEEN_CYCLES_MINUTES` | Schedule | yes |
| `DelayBetweenComputersSeconds` | Integer | 0 seconds | 0..3600 | `SEMT_DELAY_BETWEEN_COMPUTERS_SECONDS` | Parallelism | yes |
| `NightPauseEnabled` | Boolean | HJ: True; W11: False |  | `SEMT_NIGHT_PAUSE_ENABLED` | Schedule | yes |
| `NightPauseStartHour` | Integer | 20 hour (local time) | 0..23 | `SEMT_NIGHT_PAUSE_START_HOUR` | Schedule | yes |
| `NightPauseEndHour` | Integer | 7 hour (local time) | 0..23 | `SEMT_NIGHT_PAUSE_END_HOUR` | Schedule | yes |
| `ThrottleLimit` | Integer | 10 workers | 1..200 | `SEMT_THROTTLE_LIMIT` | Parallelism | yes |
| `GlobalWorkerLimit` | Integer | 15 workers | 0..200 | `SEMT_GLOBAL_WORKER_LIMIT` | Parallelism | yes |
| `GlobalWorkerLeaseTimeoutMinutes` | Integer | 0 minutes | 0..1440 | `SEMT_GLOBAL_WORKER_LEASE_TIMEOUT_MINUTES` | Parallelism | no |
| `WorkerHeartbeatStaleMinutes` | Integer | 10 minutes | 2..240 | `SEMT_WORKER_HEARTBEAT_STALE_MINUTES` | Parallelism | no |
| `PsExecTimeoutMinutes` | Integer | HJ: 120; W11: 360 minutes | 0..1440 | `SEMT_PSEXEC_TIMEOUT_MINUTES` | Parallelism | yes |
| `CancellationDrainTimeoutMinutes` | Integer | 15 minutes | 0..1440 | `SEMT_CANCELLATION_DRAIN_TIMEOUT_MINUTES` | Parallelism | no |
| `DnsPreflightBlocksCycle` | Boolean | True |  | `SEMT_DNS_PREFLIGHT_BLOCKS_CYCLE` | Network | yes |
| `DnsPreflightSampleSize` | Integer | 5 computers | 1..50 | `SEMT_DNS_PREFLIGHT_SAMPLE_SIZE` | Network | no |
| `AdDomain` | String | (empty) |  | `SEMT_AD_DOMAIN` | Network | yes |
| `UseTechnicianRunGuard` | Boolean | True |  | `SEMT_USE_TECHNICIAN_RUN_GUARD` | Safety | yes |
| `TechnicianRunGuardHours` | Integer | HJ: 12; W11: 3 hours | 0..168 | `SEMT_TECHNICIAN_RUN_GUARD_HOURS` | Safety | yes |
| `SkipVirtualMachines` | Boolean | True |  | `SEMT_SKIP_VIRTUAL_MACHINES` | Safety | yes |
| `RetryAfterRebootDelaySeconds` | Integer | 300 seconds | 0..3600 | `SEMT_RETRY_AFTER_REBOOT_DELAY_SECONDS` | Safety | no |
| `RetryAfterRebootMaxAttempts` | Integer | 3 attempts | 1..30 | `SEMT_RETRY_AFTER_REBOOT_MAX_ATTEMPTS` | Safety | no |
| `LegacyActivityPolicy` | Choice | Block | Block, Warn | `SEMT_LEGACY_ACTIVITY_POLICY` | Safety | no |
| `CentralLogCollection` | Boolean | True |  | `SEMT_CENTRAL_LOG_COLLECTION` | Evidence | yes |
| `CentralLogCollectionMode` | Choice | Standard | Standard, Full | `SEMT_CENTRAL_LOG_COLLECTION_MODE` | Evidence | yes |
| `KeepCentralLogHistory` | Boolean | False |  | `SEMT_KEEP_CENTRAL_LOG_HISTORY` | Evidence | yes |
| `RunArchiveAfterDays` | Integer | 0 days | 0..3650 | `SEMT_RUN_ARCHIVE_AFTER_DAYS` | Evidence | no |
| `ComponentIntegrityMode` | Choice | Auto | Auto, Enforce, Warn | `SEMT_COMPONENT_INTEGRITY_MODE` | Safety | no |
| `IntuneTenantId` | String | (empty) |  | `SEMT_INTUNE_TENANT_ID` | Inventory | yes |
| `IntuneInventoryPageSize` | Integer | 999 objects | 1..999 | `SEMT_INTUNE_INVENTORY_PAGE_SIZE` | Inventory | no |
| `AdInventoryFreshnessHours` | Integer | 12 hours | 1..168 | `SEMT_AD_INVENTORY_FRESHNESS_HOURS` | Inventory | no |
| `CloudInventoryFreshnessHours` | Integer | 2 hours | 1..48 | `SEMT_CLOUD_INVENTORY_FRESHNESS_HOURS` | Inventory | no |

### HybridJoin options

| Option | Type | Default | Range / choices | Env var | Category | GUI |
| --- | --- | --- | --- | --- | --- | --- |
| `HybridJoin.AllowDsregLeave` | Boolean | True |  | `SEMT_HYBRIDJOIN_ALLOW_DSREG_LEAVE` | Actions | yes |
| `HybridJoin.AllowRemoveStaleIntuneEnrollment` | Boolean | True |  | `SEMT_HYBRIDJOIN_ALLOW_REMOVE_STALE_INTUNE_ENROLLMENT` | Actions | yes |
| `HybridJoin.AllowRemoveNonIntuneMdmEnrollment` | Boolean | False |  | `SEMT_HYBRIDJOIN_ALLOW_REMOVE_NON_INTUNE_MDM_ENROLLMENT` | Actions | yes |
| `HybridJoin.AllowRebootWhenNoInteractiveUser` | Boolean | True |  | `SEMT_HYBRIDJOIN_ALLOW_REBOOT_WHEN_NO_INTERACTIVE_USER` | Actions | yes |
| `HybridJoin.AllowRebootAfterDsregLeave` | Boolean | False |  | `SEMT_HYBRIDJOIN_ALLOW_REBOOT_AFTER_DSREG_LEAVE` | Actions | yes |
| `HybridJoin.StaleCleanupDelaySeconds` | Integer | 60 seconds | 0..3600 | `SEMT_HYBRIDJOIN_STALE_CLEANUP_DELAY_SECONDS` | Actions | yes |
| `HybridJoin.RebootDelaySeconds` | Integer | 300 seconds | 60..3600 | `SEMT_HYBRIDJOIN_REBOOT_DELAY_SECONDS` | Actions | yes |
| `HybridJoin.IntuneRetrySleepMinutes` | Integer | 5 minutes | 1..60 | `SEMT_HYBRIDJOIN_INTUNE_RETRY_SLEEP_MINUTES` | Actions | yes |
| `HybridJoin.IntuneRetryMaxRetries` | Integer | 5 checks | 1..20 | `SEMT_HYBRIDJOIN_INTUNE_RETRY_MAX_RETRIES` | Actions | yes |
| `HybridJoin.CommunicationLostEvidenceWaitMinutes` | Integer | 65 minutes | 0..600 | `SEMT_HYBRIDJOIN_COMMUNICATION_LOST_EVIDENCE_WAIT_MINUTES` | Evidence | no |
| `HybridJoin.CommunicationLostEvidencePollMinutes` | Integer | 10 minutes | 1..120 | `SEMT_HYBRIDJOIN_COMMUNICATION_LOST_EVIDENCE_POLL_MINUTES` | Evidence | no |
| `HybridJoin.PostCycleCloudInventoryRefresh` | Boolean | True |  | `SEMT_HYBRIDJOIN_POST_CYCLE_CLOUD_INVENTORY_REFRESH` | Inventory | yes |

### Windows11Upgrade options

| Option | Type | Default | Range / choices | Env var | Category | GUI |
| --- | --- | --- | --- | --- | --- | --- |
| `Windows11Upgrade.AllowPolicyRepair` | Boolean | True |  | `SEMT_WINDOWS11_ALLOW_POLICY_REPAIR` | Actions | yes |
| `Windows11Upgrade.AllowWUReset` | Boolean | True |  | `SEMT_WINDOWS11_ALLOW_WU_RESET` | Actions | yes |
| `Windows11Upgrade.AllowForceUpgrade` | Boolean | True |  | `SEMT_WINDOWS11_ALLOW_FORCE_UPGRADE` | Actions | yes |
| `Windows11Upgrade.AllowSetupUpgrade` | Boolean | True |  | `SEMT_WINDOWS11_ALLOW_SETUP_UPGRADE` | Setup | yes |
| `Windows11Upgrade.DirectSetupUpgrade` | Boolean | False |  | `SEMT_WINDOWS11_DIRECT_SETUP_UPGRADE` | Setup | yes |
| `Windows11Upgrade.AllowReboot` | Boolean | True |  | `SEMT_WINDOWS11_ALLOW_REBOOT` | Actions | yes |
| `Windows11Upgrade.AllowSetupCompletionRebootWhenNoUser` | Boolean | True |  | `SEMT_WINDOWS11_SETUP_COMPLETION_REBOOT_WHEN_NO_USER` | Actions | yes |
| `Windows11Upgrade.ForceRequiredRebootWhenUptimeOverDays` | Integer | 7 days (0 = never) | 0..3650 | `SEMT_WINDOWS11_FORCE_REQUIRED_REBOOT_WHEN_UPTIME_OVER_DAYS` | Actions | yes |
| `Windows11Upgrade.AllowSetupProfileRepair` | Boolean | True |  | `SEMT_WINDOWS11_ALLOW_SETUP_PROFILE_REPAIR` | Actions | yes |
| `Windows11Upgrade.ScheduleRetryAfterReboot` | Boolean | True |  | `SEMT_WINDOWS11_SCHEDULE_RETRY_AFTER_REBOOT` | Actions | yes |
| `Windows11Upgrade.EndpointRunGuardHours` | Integer | 3 hours | 0..168 | `SEMT_WINDOWS11_ENDPOINT_RUN_GUARD_HOURS` | Safety | no |
| `Windows11Upgrade.AllowDiskCleanup` | Boolean | True |  | `SEMT_WINDOWS11_ALLOW_DISK_CLEANUP` | Actions | yes |
| `Windows11Upgrade.AllowAdvancedDiskCleanup` | Boolean | False |  | `SEMT_WINDOWS11_ALLOW_ADVANCED_DISK_CLEANUP` | Actions | yes |
| `Windows11Upgrade.MinimumFreeDiskGB` | Integer | 32 GB | 10..200 | `SEMT_WINDOWS11_MINIMUM_FREE_DISK_GB` | Setup | no |
| `Windows11Upgrade.SetupSourcePath` | Path | (empty) |  | `SEMT_WINDOWS11_SETUP_SOURCE` | Setup | yes |
| `Windows11Upgrade.SetupSourceMapPath` | Path | (empty) |  | `SEMT_WINDOWS11_SETUP_SOURCE_MAP` | Setup | yes |
| `Windows11Upgrade.ConfirmLocalSetupSource` | Boolean | False |  | `SEMT_WINDOWS11_CONFIRM_LOCAL_SETUP_SOURCE` | Setup | yes |
| `Windows11Upgrade.SetupExecutionMode` | Choice | LocalCache | LocalCache, Share, Auto | `SEMT_WINDOWS11_SETUP_EXECUTION_MODE` | Setup | yes |
| `Windows11Upgrade.SetupMediaId` | String | Win11 |  | `SEMT_WINDOWS11_SETUP_MEDIA_ID` | Setup | yes |
| `Windows11Upgrade.SetupLanguage` | String | MatchSystem |  | `SEMT_WINDOWS11_SETUP_LANGUAGE` | Setup | yes |
| `Windows11Upgrade.SetupDynamicUpdate` | Choice | Disable | Enable, Disable, NoDrivers, NoLCU, NoDriversNoLCU | `SEMT_WINDOWS11_SETUP_DYNAMIC_UPDATE` | Setup | yes |
| `Windows11Upgrade.SkipSetupMediaPreCopy` | Boolean | False |  | `SEMT_WINDOWS11_SKIP_SETUP_MEDIA_PRECOPY` | Setup | yes |
| `Windows11Upgrade.ReuseLegacySetupMedia` | Boolean | True |  | `SEMT_WINDOWS11_REUSE_LEGACY_SETUP_MEDIA` | Setup | yes |
| `Windows11Upgrade.AllowSetupMediaDuplication` | Boolean | False |  | `SEMT_WINDOWS11_ALLOW_SETUP_MEDIA_DUPLICATION` | Setup | yes |
| `Windows11Upgrade.SetupSourceCandidateLimit` | Integer | 5 sources | 0..100 | `SEMT_WINDOWS11_SETUP_SOURCE_CANDIDATE_LIMIT` | Setup | no |
| `Windows11Upgrade.SetupMediaCopyIpGapMilliseconds` | Integer | 20 milliseconds | 0..10000 | `SEMT_WINDOWS11_SETUP_COPY_IPG_MS` | Setup | yes |
| `Windows11Upgrade.SetupMediaCopyJitterSeconds` | Integer | 300 seconds | 0..86400 | `SEMT_WINDOWS11_SETUP_COPY_JITTER_SECONDS` | Setup | yes |
| `Windows11Upgrade.SetupMediaCopyTimeoutMinutes` | Integer | 180 minutes | 0..1440 | `SEMT_WINDOWS11_SETUP_COPY_TIMEOUT_MINUTES` | Setup | no |
| `Windows11Upgrade.SetupSourceConcurrencyLimit` | Integer | 0 copies | 0..500 | `SEMT_WINDOWS11_SETUP_SOURCE_CONCURRENCY_LIMIT` | Setup | yes |
| `Windows11Upgrade.SetupSourceConcurrencyLeaseMinutes` | Integer | 240 minutes | 1..1440 | `SEMT_WINDOWS11_SETUP_SOURCE_CONCURRENCY_LEASE_MINUTES` | Setup | no |
| `Windows11Upgrade.SetupSourceConcurrencyGateRoot` | Path | (empty) |  | `SEMT_WINDOWS11_SETUP_SOURCE_CONCURRENCY_GATE_ROOT` | Setup | yes |
| `Windows11Upgrade.SetupSubnetConcurrencyLimit` | Integer | 1 copies | 0..500 | `SEMT_WINDOWS11_SETUP_SUBNET_CONCURRENCY_LIMIT` | Setup | yes |
| `Windows11Upgrade.SetupSubnetPrefixLength` | String | Auto |  | `SEMT_WINDOWS11_SETUP_SUBNET_PREFIX_LENGTH` | Setup | no |
| `Windows11Upgrade.SetupSubnetConcurrencyLeaseMinutes` | Integer | 90 minutes | 1..1440 | `SEMT_WINDOWS11_SETUP_SUBNET_CONCURRENCY_LEASE_MINUTES` | Setup | no |
| `Windows11Upgrade.SetupSubnetConcurrencyGateRoot` | Path | (empty) |  | `SEMT_WINDOWS11_SETUP_SUBNET_CONCURRENCY_GATE_ROOT` | Setup | yes |
| `Windows11Upgrade.PreCycleAdInventoryRefresh` | Boolean | True |  | `SEMT_WINDOWS11_PRE_CYCLE_AD_INVENTORY_REFRESH` | Inventory | yes |
| `Windows11Upgrade.PreCycleIntuneInventoryRefresh` | Boolean | False |  | `SEMT_WINDOWS11_PRE_CYCLE_INTUNE_INVENTORY_REFRESH` | Inventory | yes |
