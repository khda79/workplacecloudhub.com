<#
.SYNOPSIS
    Runs SmartM365-Invoke-Windows11UpgradeRepair.ps1 on remote computers using PsExec.

.DESCRIPTION
    LOT/PsExec orchestrator for Windows 10 to Windows 11 upgrade diagnostics and guarded repair.
    It copies the autonomous endpoint script to each target, lets the target validate/cache
    Windows 11 setup media when setup upgrade is enabled, starts the script as SYSTEM,
    collects evidence, and writes cycle CSV reports.

.VERSION
    0.1.39

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
#>

#requires -Version 5.1

[CmdletBinding(PositionalBinding = $false)]
param(
    [string]$ComputerListPath = (Join-Path (Get-Location) 'Computers.txt'),
    [string]$PsExecPath,
    [string]$LocalScriptPath,

    [switch]$AuditOnly,
    [switch]$DryRun,
    [switch]$RunOnce,
    [switch]$IgnoreRunGuard,
    [switch]$UseTechnicianRunGuardHistory,
    [switch]$IgnoreTechnicianRunGuardHistory,
    [ValidateRange(0, 168)][int]$RunGuardHours = 12,
    [switch]$AllowPolicyRepair,
    [switch]$AllowWUReset,
    [switch]$AllowForceUpgrade,
    [switch]$AllowSetupUpgrade,
    [switch]$DirectSetupUpgrade,
    [switch]$AllowReboot,
    [switch]$AllowSetupCompletionRebootWhenNoUser,
    [switch]$AllowSetupProfileRepair,
    [switch]$SkipVirtualMachines,
    [switch]$AllowDiskCleanup,
    [switch]$AllowAdvancedDiskCleanup,
    [switch]$AllowDismComponentCleanup,

    [string]$SetupSourcePath,
    [string]$SetupSourceMapPath,
    [ValidateSet('LocalCache','Share','Auto')]
    [string]$SetupExecutionMode = 'LocalCache',
    [string]$SetupMediaId = 'Win11',
    [string]$SetupLanguage = 'MatchSystem',
    [ValidateSet('Enable','Disable','NoDrivers','NoLCU','NoDriversNoLCU')]
    [string]$SetupDynamicUpdate = 'Disable',
    [switch]$SkipSetupMediaPreCopy,
    [ValidateRange(0, 100)][int]$SetupSourceCandidateLimit = 5,
    [ValidateRange(0, 10000)][int]$SetupMediaCopyIpGapMilliseconds = 0,
    [ValidateRange(0, 86400)][int]$SetupMediaCopyJitterSeconds = 0,
    [ValidateRange(0, 500)][int]$SetupSourceConcurrencyLimit = 0,
    [ValidateRange(1, 1440)][int]$SetupSourceConcurrencyLeaseMinutes = 240,
    [string]$SetupSourceConcurrencyGateRoot,
    [ValidateRange(0, 500)][int]$SetupSubnetConcurrencyLimit = 0,
    [string]$SetupSubnetPrefixLength = 'Auto',
    [ValidateRange(1, 1440)][int]$SetupSubnetConcurrencyLeaseMinutes = 90,
    [string]$SetupSubnetConcurrencyGateRoot,

    [string]$AdInventoryCsv,
    [string]$AdRootInventoryCsv,
    [string]$AdInventoryNameColumn,
    [string]$AdDomain,
    [switch]$SkipAdInventoryRefresh,
    [string]$IntuneInventoryCsv,
    [string]$IntuneRootInventoryCsv,
    [string]$IntuneInventoryNameColumn,
    [ValidateRange(1, 999)][int]$IntuneInventoryPageSize = 999,
    [string]$IntuneTenantId,
    [switch]$SkipIntuneInventoryRefresh,

    [string]$LogRoot,
    [string]$ReportRoot,
    [string]$CentralLogRoot,
    [switch]$NoCentralLogCollection,
    [switch]$KeepCentralLogHistory,
    [ValidateSet('Standard','Full')]
    [string]$CentralLogCollectionMode = 'Standard',

    [ValidateRange(1, 200)][int]$ThrottleLimit = 10,
    [ValidateRange(0, 200)][int]$GlobalConcurrencyLimit = 15,
    [string]$GlobalConcurrencySemaphoreName = 'Local\SmartM365_Windows11UpgradeToolkit_ComputerWorkers',
    [ValidateRange(0, 1440)][int]$GlobalConcurrencyLeaseTimeoutMinutes = 0,
    [ValidateRange(0, 3600)][int]$DelayBetweenComputersSeconds = 0,
    [ValidateRange(1, 60)][int]$JobPollSeconds = 2,
    [ValidateRange(0, 1440)][int]$DelayBetweenCyclesMinutes = 5,
    [ValidateRange(0, 1000)][int]$MaxCycles = 0,
    [ValidateRange(0, 1440)][int]$PsExecTimeoutMinutes = 360,

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$UnexpectedArguments
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

if ($UnexpectedArguments -and $UnexpectedArguments.Count -gt 0) {
    throw ("Unexpected launcher argument(s): {0}. Pass PsExec with -PsExecPath <path>, not as a free argument." -f ($UnexpectedArguments -join ' '))
}

$script:LauncherVersion = '0.1.39'
$script:BaseDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$script:ToolkitRoot = Split-Path -Parent $script:BaseDir
if ([string]::IsNullOrWhiteSpace($LocalScriptPath)) {
    $LocalScriptPath = Join-Path $script:BaseDir 'SmartM365-Invoke-Windows11UpgradeRepair.ps1'
}
$LocalWorkerPath = Join-Path $script:BaseDir 'SmartM365-Windows11Upgrade-PsExecWorker.ps1'
$script:ExportAdScriptPath = Join-Path $script:BaseDir 'SmartM365-Windows11Upgrade-Export-ADDevicesCsv.ps1'
$script:ExportIntuneScriptPath = Join-Path $script:BaseDir 'SmartM365-Windows11Upgrade-Export-IntuneDevicesCsv.ps1'
$script:AdInventoryFreshnessHours = 12
$script:IntuneInventoryFreshnessHours = 2
$ComputerListPath = [System.IO.Path]::GetFullPath($ComputerListPath)
$script:LotRoot = Split-Path -Parent $ComputerListPath
$script:LotAdInventoryCsv = Join-Path $script:LotRoot 'DevicesAD.csv'
$script:LotIntuneInventoryCsv = Join-Path $script:LotRoot 'DevicesIntune.csv'
if ([string]::IsNullOrWhiteSpace($LogRoot)) { $LogRoot = Join-Path (Split-Path -Parent $ComputerListPath) 'PsExecLogs' }
if ([string]::IsNullOrWhiteSpace($ReportRoot)) { $ReportRoot = Join-Path (Split-Path -Parent $ComputerListPath) 'Reports' }
if ([string]::IsNullOrWhiteSpace($CentralLogRoot)) { $CentralLogRoot = Join-Path (Split-Path -Parent $ComputerListPath) 'CentralLogs' }
$LogRoot = [System.IO.Path]::GetFullPath($LogRoot)
$ReportRoot = [System.IO.Path]::GetFullPath($ReportRoot)
$CentralLogRoot = [System.IO.Path]::GetFullPath($CentralLogRoot)
$script:LauncherRunTimestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$script:LauncherLogRoot = Join-Path $script:LotRoot 'Logs'
$script:LauncherLotName = Split-Path -Leaf $script:LotRoot
if ([string]::IsNullOrWhiteSpace($script:LauncherLotName)) { $script:LauncherLotName = 'UnknownLOT' }
$script:LauncherLogSafeLotName = [regex]::Replace($script:LauncherLotName, '[^A-Za-z0-9._-]', '_')
if ([string]::IsNullOrWhiteSpace($script:LauncherLogSafeLotName)) { $script:LauncherLogSafeLotName = 'UnknownLOT' }
$script:LauncherLogPath = Join-Path $script:LauncherLogRoot ("SmartM365-W11UT-Launcher_{0}_{1}.log" -f $script:LauncherLogSafeLotName,$script:LauncherRunTimestamp)
$script:LauncherLatestLogPath = Join-Path $script:LauncherLogRoot ("SmartM365-W11UT-Launcher_{0}_latest.log" -f $script:LauncherLogSafeLotName)

$AdInventoryUsesRecentRootCsv = $false
$effectiveRootAdInventoryCsv = $AdRootInventoryCsv
if ([string]::IsNullOrWhiteSpace($effectiveRootAdInventoryCsv)) {
    $defaultRootAdCsv = Join-Path $script:ToolkitRoot 'DevicesAD.csv'
    if (Test-Path -LiteralPath $defaultRootAdCsv -PathType Leaf) {
        $effectiveRootAdInventoryCsv = $defaultRootAdCsv
    }
}
if (-not [string]::IsNullOrWhiteSpace($effectiveRootAdInventoryCsv)) {
    $adRootInventoryItem = Get-Item -LiteralPath $effectiveRootAdInventoryCsv -ErrorAction SilentlyContinue
    if ($adRootInventoryItem) {
        $adRootInventoryAge = (Get-Date) - $adRootInventoryItem.LastWriteTime
        if ($adRootInventoryAge.TotalHours -le $script:AdInventoryFreshnessHours) {
            $AdInventoryCsv = $adRootInventoryItem.FullName
            $AdRootInventoryCsv = $adRootInventoryItem.FullName
            $AdInventoryUsesRecentRootCsv = $true
        }
    }
}
if ([string]::IsNullOrWhiteSpace($AdInventoryCsv)) {
    $AdInventoryCsv = $script:LotAdInventoryCsv
}
if (-not [string]::IsNullOrWhiteSpace($AdInventoryCsv)) { $AdInventoryCsv = [System.IO.Path]::GetFullPath($AdInventoryCsv) }
if (-not [string]::IsNullOrWhiteSpace($AdRootInventoryCsv)) { $AdRootInventoryCsv = [System.IO.Path]::GetFullPath($AdRootInventoryCsv) }

$IntuneInventoryUsesRecentRootCsv = $false
$requestedIntuneInventoryCsv = $IntuneInventoryCsv
$defaultRootIntuneCsv = Join-Path $script:ToolkitRoot 'DevicesIntune.csv'
$defaultLotIntuneFullName = [System.IO.Path]::GetFullPath($script:LotIntuneInventoryCsv)
$requestedIntuneFullName = if (-not [string]::IsNullOrWhiteSpace($requestedIntuneInventoryCsv)) { [System.IO.Path]::GetFullPath($requestedIntuneInventoryCsv) } else { '' }

if ([string]::IsNullOrWhiteSpace($IntuneRootInventoryCsv) -and (Test-Path -LiteralPath $defaultRootIntuneCsv -PathType Leaf)) {
    $IntuneRootInventoryCsv = $defaultRootIntuneCsv
}
if (-not [string]::IsNullOrWhiteSpace($IntuneRootInventoryCsv)) {
    $intuneRootInventoryItem = Get-Item -LiteralPath $IntuneRootInventoryCsv -ErrorAction SilentlyContinue
    if ($intuneRootInventoryItem) {
        $intuneRootFullName = [System.IO.Path]::GetFullPath($intuneRootInventoryItem.FullName)
        $requestedIntuneItem = if (-not [string]::IsNullOrWhiteSpace($requestedIntuneFullName)) { Get-Item -LiteralPath $requestedIntuneFullName -ErrorAction SilentlyContinue } else { $null }
        $shouldPreferRootIntuneCsv = [string]::IsNullOrWhiteSpace($requestedIntuneFullName) -or ($requestedIntuneFullName -eq $defaultLotIntuneFullName) -or (-not $requestedIntuneItem)
        $intuneRootInventoryAge = (Get-Date) - $intuneRootInventoryItem.LastWriteTime
        if ($shouldPreferRootIntuneCsv -and $intuneRootInventoryAge.TotalHours -le $script:IntuneInventoryFreshnessHours) {
            $IntuneInventoryCsv = $intuneRootFullName
            $IntuneRootInventoryCsv = $intuneRootFullName
            $IntuneInventoryUsesRecentRootCsv = $true
        }
    }
}
if ([string]::IsNullOrWhiteSpace($IntuneInventoryCsv)) {
    $IntuneInventoryCsv = $script:LotIntuneInventoryCsv
}
if (-not [string]::IsNullOrWhiteSpace($IntuneInventoryCsv)) { $IntuneInventoryCsv = [System.IO.Path]::GetFullPath($IntuneInventoryCsv) }
if (-not [string]::IsNullOrWhiteSpace($IntuneRootInventoryCsv)) { $IntuneRootInventoryCsv = [System.IO.Path]::GetFullPath($IntuneRootInventoryCsv) }

$script:RemoteBaseDir = 'C:\ProgramData\SmartM365\Windows11UpgradeToolkit'
$script:RemoteScriptPath = Join-Path $script:RemoteBaseDir 'SmartM365-Invoke-Windows11UpgradeRepair.ps1'
$script:RemoteSetupCacheRoot = Join-Path $script:RemoteBaseDir 'SetupMedia'
if ($GlobalConcurrencyLeaseTimeoutMinutes -lt 1) {
    $timeoutBase = if ($PsExecTimeoutMinutes -gt 0) { $PsExecTimeoutMinutes } else { 240 }
    $GlobalConcurrencyLeaseTimeoutMinutes = [Math]::Max(30, $timeoutBase + 30)
}

function Get-TechnicianIdentityInfo {
    $identity = $null
    try { $identity = [System.Security.Principal.WindowsIdentity]::GetCurrent() } catch { }

    $upn = ''
    try {
        $whoamiUpn = & whoami.exe /upn 2>$null
        if ($LASTEXITCODE -eq 0 -and -not [string]::IsNullOrWhiteSpace($whoamiUpn)) { $upn = ([string]$whoamiUpn).Trim() }
    }
    catch { }

    $account = if ($identity -and -not [string]::IsNullOrWhiteSpace($identity.Name)) { [string]$identity.Name } else { [Environment]::UserName }
    $sid = if ($identity -and $identity.User) { [string]$identity.User.Value } else { '' }
    $authType = if ($identity) { [string]$identity.AuthenticationType } else { '' }

    return [pscustomobject]@{
        Account = $account
        UserPrincipalName = $upn
        Sid = $sid
        AuthenticationType = $authType
        UserDomain = [Environment]::UserDomainName
        UserName = [Environment]::UserName
        UserDnsDomain = [Environment]::GetEnvironmentVariable('USERDNSDOMAIN')
        ComputerName = $env:COMPUTERNAME
    }
}

$script:TechnicianIdentity = Get-TechnicianIdentityInfo

function New-Directory {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
    }
}

function Get-TechnicianRunGuardHistoryPath {
    $stateRoot = Join-Path ([Environment]::GetFolderPath('CommonApplicationData')) 'SmartM365\Windows11UpgradeToolkit\LauncherState'
    return (Join-Path $stateRoot 'RunGuardHistory.json')
}

function Invoke-TechnicianRunGuardHistoryLock {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock,
        [object[]]$ArgumentList = @()
    )

    $mutex = $null
    $acquired = $false
    try {
        $mutex = New-Object System.Threading.Mutex($false, 'Local\SmartM365_Windows11UpgradeToolkit_TechnicianRunGuardHistory')
        $acquired = $mutex.WaitOne([TimeSpan]::FromSeconds(15))
        if (-not $acquired) { throw 'Timed out waiting for technician run guard history lock.' }
        & $ScriptBlock @ArgumentList
    }
    finally {
        if ($acquired -and $mutex) { try { $mutex.ReleaseMutex() } catch { } }
        if ($mutex) { $mutex.Dispose() }
    }
}

function Get-TechnicianRunGuardFqdn {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [AllowNull()][hashtable]$AdInventoryMap
    )

    $name = ([string]$ComputerName).Trim().Trim([char]34).TrimEnd('.')
    if ([string]::IsNullOrWhiteSpace($name)) { return '' }
    if ($name.Contains('.')) { return $name.ToLowerInvariant() }

    $shortKey = Get-ComputerListKey -ComputerName $name
    if ($AdInventoryMap -and $AdInventoryMap.Count -gt 0 -and $AdInventoryMap.ContainsKey($shortKey)) {
        $adRow = $AdInventoryMap[$shortKey]
        if ($adRow.PSObject.Properties['DNSHostName'] -and -not [string]::IsNullOrWhiteSpace([string]$adRow.DNSHostName)) {
            return ([string]$adRow.DNSHostName).Trim().TrimEnd('.').ToLowerInvariant()
        }
    }

    return $name.ToLowerInvariant()
}

function ConvertTo-TechnicianRunGuardUtcDateTime {
    param(
        [AllowNull()]$Value,
        [Parameter(Mandatory = $true)][ref]$Result
    )

    $Result.Value = [datetime]::MinValue
    if ($null -eq $Value) { return $false }
    if ($Value -is [datetime]) {
        $Result.Value = ([datetime]$Value).ToUniversalTime()
        return $true
    }

    $parsed = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref]$parsed)) {
        $Result.Value = $parsed.ToUniversalTime()
        return $true
    }
    return $false
}

function Read-TechnicianRunGuardHistory {    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [ValidateRange(0, 168)][int]$RunGuardHours
    )

    $entries = @()
    if (Test-Path -LiteralPath $Path -PathType Leaf) {
        try {
            $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
            if (-not [string]::IsNullOrWhiteSpace($raw)) {
                $data = $raw | ConvertFrom-Json -ErrorAction Stop
                if ($data.PSObject.Properties['Entries']) { $entries = @($data.Entries) }
            }
        }
        catch {
            $entries = @()
        }
    }

    if ($RunGuardHours -gt 0) {
        $nowUtc = (Get-Date).ToUniversalTime()
        $fresh = New-Object System.Collections.ArrayList
        foreach ($entry in @($entries)) {
            $startedText = if ($entry.PSObject.Properties['LastStartedUtc']) { [string]$entry.LastStartedUtc } else { '' }
            $startedUtc = [datetime]::MinValue
            if (ConvertTo-TechnicianRunGuardUtcDateTime -Value $startedText -Result ([ref]$startedUtc)) {
                if (($nowUtc - $startedUtc.ToUniversalTime()).TotalHours -lt $RunGuardHours) { [void]$fresh.Add($entry) }
            }
        }
        $entries = @($fresh.ToArray())
    }

    return [pscustomobject]@{
        Version = 1
        UpdatedUtc = (Get-Date).ToUniversalTime().ToString('o')
        Entries = @($entries)
    }
}

function Save-TechnicianRunGuardHistory {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)]$History
    )

    New-Directory -Path (Split-Path -Parent $Path)
    $History.UpdatedUtc = (Get-Date).ToUniversalTime().ToString('o')
    $History | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Test-TechnicianRunGuardEntryShouldBlock {
    param([AllowNull()]$Entry)

    if (-not $Entry) { return $false }
    $state = if ($Entry.PSObject.Properties['State']) { [string]$Entry.State } else { '' }
    if ($state -ne 'Result') { return $true }

    $launcherStatus = if ($Entry.PSObject.Properties['LauncherStatus']) { [string]$Entry.LauncherStatus } else { '' }
    $remoteStatus = if ($Entry.PSObject.Properties['RemoteStatus']) { [string]$Entry.RemoteStatus } else { '' }
    $detail = if ($Entry.PSObject.Properties['Detail']) { [string]$Entry.Detail } else { '' }
    $jobErrorMessage = if ($Entry.PSObject.Properties['JobErrorMessage']) { [string]$Entry.JobErrorMessage } else { '' }
    $effectiveStatus = @($remoteStatus, $launcherStatus) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1

    if ($effectiveStatus -in @(
        'ADMIN_SHARE_UNREACHABLE',
        'DRYRUN_ADMIN_SHARE_UNREACHABLE',
        'PSEXEC_EXIT_UNKNOWN',
        'PSEXEC_COMMUNICATION_LOST',
        'CENTRAL_LOG_COLLECTION_FAILED',
        'REMOTE_LOG_COLLECTION_FAILED'
    )) { return $false }

    $combinedEvidence = @($detail, $jobErrorMessage) -join ' '
    if ($combinedEvidence -match 'FailureType=(DNS_FAILED|SMB_PORT_445_UNREACHABLE|PING_OK_ADMIN_SHARE_FAILED|ADMIN_SHARE_UNREACHABLE)') { return $false }
    if ($combinedEvidence -match 'Central log collection failed|Le chemin r.seau n.a pas .t. trouv.|network path was not found|Error communicating with PsExec service|Descripteur non valide') { return $false }

    return $true
}
function Get-ActiveTechnicianRunGuardEntry {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ComputerFqdn,
        [ValidateRange(0, 168)][int]$RunGuardHours
    )

    if ([string]::IsNullOrWhiteSpace($ComputerFqdn) -or $RunGuardHours -le 0) { return $null }

    $found = @{}
    Invoke-TechnicianRunGuardHistoryLock -ArgumentList @($Path, $RunGuardHours, $ComputerFqdn, $found) -ScriptBlock {
        param($LockedPath, $LockedRunGuardHours, $LockedComputerFqdn, $FoundRef)
        $history = Read-TechnicianRunGuardHistory -Path $LockedPath -RunGuardHours $LockedRunGuardHours
        Save-TechnicianRunGuardHistory -Path $LockedPath -History $history
        foreach ($entry in @($history.Entries)) {
            if ([string]$entry.ComputerFqdn -eq $LockedComputerFqdn) { $FoundRef.Value = $entry; break }
        }
    }
    if ($found.ContainsKey('Value') -and (Test-TechnicianRunGuardEntryShouldBlock -Entry $found.Value)) { return $found.Value }
    return $null
}

function Update-TechnicianRunGuardHistory {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$ComputerFqdn,
        [Parameter(Mandatory = $true)][string]$InputComputerName,
        [ValidateRange(0, 168)][int]$RunGuardHours,
        [ValidateSet('Started','Result')][string]$State,
        [AllowNull()]$Result,
        [AllowNull()][string]$JobId
    )

    if ([string]::IsNullOrWhiteSpace($ComputerFqdn) -or $RunGuardHours -le 0) { return }

    Invoke-TechnicianRunGuardHistoryLock -ArgumentList @(
        $Path,
        $ComputerFqdn,
        $InputComputerName,
        $RunGuardHours,
        $State,
        $Result,
        $JobId,
        $cycle,
        $ComputerListPath,
        $script:LotRoot,
        [string]$script:TechnicianIdentity.Account,
        [string]$script:TechnicianIdentity.ComputerName,
        [string]$script:LauncherVersion
    ) -ScriptBlock {
        param(
            $LockedPath,
            $LockedComputerFqdn,
            $LockedInputComputerName,
            $LockedRunGuardHours,
            $LockedState,
            $LockedResult,
            $LockedJobId,
            $LockedCycle,
            $LockedComputerListPath,
            $LockedLotRoot,
            $LockedTechnicianAccount,
            $LockedTechnicianComputer,
            $LockedLauncherVersion
        )

        $history = Read-TechnicianRunGuardHistory -Path $LockedPath -RunGuardHours $LockedRunGuardHours
        $kept = New-Object System.Collections.ArrayList
        $previousEntry = $null
        foreach ($entry in @($history.Entries)) {
            if ([string]$entry.ComputerFqdn -eq $LockedComputerFqdn) { $previousEntry = $entry }
            else { [void]$kept.Add($entry) }
        }

        $nowUtc = (Get-Date).ToUniversalTime()
        $startedUtc = $nowUtc
        if ($LockedState -eq 'Result' -and $previousEntry -and $previousEntry.PSObject.Properties['LastStartedUtc']) {
            $previousStartedUtc = [datetime]::MinValue
            if (ConvertTo-TechnicianRunGuardUtcDateTime -Value $previousEntry.LastStartedUtc -Result ([ref]$previousStartedUtc)) { $startedUtc = $previousStartedUtc.ToUniversalTime() }
        }

        $launcherStatus = ''
        $remoteStatus = ''
        $remoteNextAction = ''
        $exitCode = ''
        $detail = ''
        $psExecLogPath = ''
        $remoteLogsPath = ''

        if ($LockedResult) {
            if ($LockedResult.PSObject.Properties['LauncherStatus']) { $launcherStatus = [string]$LockedResult.LauncherStatus }
            if ($LockedResult.PSObject.Properties['RemoteStatus']) { $remoteStatus = [string]$LockedResult.RemoteStatus }
            if ($LockedResult.PSObject.Properties['RemoteNextAction']) { $remoteNextAction = [string]$LockedResult.RemoteNextAction }
            if ($LockedResult.PSObject.Properties['ExitCode']) { $exitCode = [string]$LockedResult.ExitCode }
            if ($LockedResult.PSObject.Properties['Detail']) { $detail = [string]$LockedResult.Detail }
            if ($LockedResult.PSObject.Properties['PsExecLogPath']) { $psExecLogPath = [string]$LockedResult.PsExecLogPath }
            if ($LockedResult.PSObject.Properties['RemoteLogsPath']) { $remoteLogsPath = [string]$LockedResult.RemoteLogsPath }
        }

        [void]$kept.Add([pscustomobject]@{
            ComputerFqdn = $LockedComputerFqdn
            InputComputerName = $LockedInputComputerName
            LastStartedUtc = $startedUtc.ToString('o')
            LastUpdatedUtc = $nowUtc.ToString('o')
            LastResultUtc = if ($LockedState -eq 'Result') { $nowUtc.ToString('o') } else { '' }
            ExpiresUtc = $startedUtc.AddHours($LockedRunGuardHours).ToString('o')
            State = $LockedState
            LauncherStatus = $launcherStatus
            RemoteStatus = $remoteStatus
            RemoteNextAction = $remoteNextAction
            ExitCode = $exitCode
            Detail = if ($detail.Length -gt 800) { $detail.Substring(0, 800) } else { $detail }
            PsExecLogPath = $psExecLogPath
            RemoteLogsPath = $remoteLogsPath
            JobErrorMessage = if ($LockedResult -and $LockedResult.PSObject.Properties['JobErrorMessage']) { [string]$LockedResult.JobErrorMessage } else { '' }
            JobId = [string]$LockedJobId
            CycleNumber = $LockedCycle
            ComputerListPath = $LockedComputerListPath
            LotRoot = $LockedLotRoot
            TechnicianAccount = $LockedTechnicianAccount
            TechnicianComputer = $LockedTechnicianComputer
            LauncherVersion = $LockedLauncherVersion
            RunGuardHours = $LockedRunGuardHours
        })

        $history.Entries = @($kept.ToArray())
        Save-TechnicianRunGuardHistory -Path $LockedPath -History $history
    }
}
function New-TechnicianRunGuardSkippedResult {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerName,
        [Parameter(Mandatory = $true)][int]$CycleNumber,
        [Parameter(Mandatory = $true)]$HistoryEntry,
        [ValidateRange(0, 168)][int]$RunGuardHours
    )

    $startedUtc = [datetime]::MinValue
    [void](ConvertTo-TechnicianRunGuardUtcDateTime -Value $HistoryEntry.LastStartedUtc -Result ([ref]$startedUtc))
    $ageHours = if ($startedUtc -gt [datetime]::MinValue) { ((Get-Date).ToUniversalTime() - $startedUtc.ToUniversalTime()).TotalHours } else { 0 }
    $startedText = if ($startedUtc -gt [datetime]::MinValue) { $startedUtc.ToUniversalTime().ToString('o') } else { [string]$HistoryEntry.LastStartedUtc }
    $expiresUtc = [datetime]::MinValue
    if ($HistoryEntry.PSObject.Properties['ExpiresUtc']) { [void](ConvertTo-TechnicianRunGuardUtcDateTime -Value $HistoryEntry.ExpiresUtc -Result ([ref]$expiresUtc)) }
    $expiresText = if ($expiresUtc -gt [datetime]::MinValue) { $expiresUtc.ToUniversalTime().ToString('o') } else { [string]$HistoryEntry.ExpiresUtc }
    $lastStatus = @([string]$HistoryEntry.RemoteStatus, [string]$HistoryEntry.LauncherStatus) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1

    return [pscustomobject]@{
        Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        ComputerName = $ComputerName
        CycleNumber = $CycleNumber
        LauncherStatus = 'SKIPPED_BY_TECH_RUN_GUARD'
        RemoteStatus = ''
        RemoteNextAction = 'WAIT_RUN_GUARD_EXPIRY'
        ExitCode = 0
        Detail = ("Technician run guard history skipped launch. FQDN={0}; LastStartedUtc={1}; AgeHours={2:N1}; GuardHours={3}; ExpiresUtc={4}; LastStatus={5}" -f $HistoryEntry.ComputerFqdn,$startedText,$ageHours,$RunGuardHours,$expiresText,$lastStatus)
        SetupCacheAction = ''
        DiskCleanupAction = ''
        DiskCleanupFreedGB = ''
        AdvancedDiskCleanupAction = ''
        AdvancedDiskCleanupFreedGB = ''
        DismCleanupAction = ''
        DismCleanupFreedGB = ''
        SetupCompletionRebootAction = ''
        SetupCompletionRebootDetail = ''
        SetupCompletionRebootUserCount = ''
        SetupCompletionRebootUsers = ''
        SetupProfileRepairAction = ''
        SetupProfileRepairDetail = ''
        SetupProfileRepairBlockingSid = ''
        SetupProfileRepairKeptSid = ''
        SetupProfileRepairProfilePath = ''
        SetupProfileRepairBackupPath = ''
        ControlledRebootAction = ''
        ControlledRebootDetail = ''
        ControlledRebootUserCount = ''
        ControlledRebootUsers = ''
        UserRebootNotificationSent = ''
        UserRebootNotificationLang = ''
        UserRebootNotificationMessage = ''
        RemoteLogsPath = ''
        PsExecLogPath = ''
        JobErrorMessage = ''
    }
}

function Get-ComputerList {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Computer list not found: $Path"
    }

    $seen = @{}
    $result = New-Object System.Collections.ArrayList
    foreach ($line in @(Get-Content -LiteralPath $Path -ErrorAction Stop)) {
        $name = ([string]$line).Trim().Trim([char]34)
        if ([string]::IsNullOrWhiteSpace($name) -or $name.StartsWith('#')) { continue }
        $key = $name.ToUpperInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        [void]$result.Add($name)
    }
    return @($result.ToArray())
}

function Get-ComputerListStats {
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Computer list not found: $Path"
    }

    $rawNames = @(
        Get-Content -LiteralPath $Path -ErrorAction Stop |
            ForEach-Object { ([string]$_).Trim().Trim([char]34) } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.StartsWith('#') }
    )
    $duplicateGroups = @(
        $rawNames |
            Group-Object { ([string]$_).ToUpperInvariant() } |
            Where-Object { $_.Count -gt 1 } |
            Sort-Object -Property @{ Expression = 'Count'; Descending = $true },Name
    )
    $duplicateLineCount = 0
    foreach ($duplicateGroup in $duplicateGroups) { $duplicateLineCount += ([int]$duplicateGroup.Count - 1) }

    return [pscustomobject]@{
        RawLines = [int]$rawNames.Count
        Unique = [int](@($rawNames | ForEach-Object { ([string]$_).ToUpperInvariant() } | Select-Object -Unique).Count)
        DuplicateGroups = [int]$duplicateGroups.Count
        DuplicateLines = [int]$duplicateLineCount
        DuplicateSamples = [string](($duplicateGroups | Select-Object -First 10 | ForEach-Object { $_.Group[0] }) -join ', ')
    }
}

function Get-ComputerListKey {
    param([Parameter(Mandatory = $true)][string]$ComputerName)

    return ($ComputerName.Trim().Trim([char]34).Split('.')[0]).ToUpperInvariant()
}

function Test-AlreadyWindows11CycleResult {
    param([Parameter(Mandatory = $true)][psobject]$Result)

    $launcherStatus = if ($Result.PSObject.Properties['LauncherStatus']) { [string]$Result.LauncherStatus } else { '' }
    $remoteStatus = if ($Result.PSObject.Properties['RemoteStatus']) { [string]$Result.RemoteStatus } else { '' }
    return ($launcherStatus -eq 'ALREADY_WINDOWS11' -or $remoteStatus -eq 'ALREADY_WINDOWS11')
}

function Move-AlreadyWindows11ComputersFromList {
    param(
        [Parameter(Mandatory = $true)][string]$ComputerListPath,
        [Parameter(Mandatory = $true)][object[]]$CycleSummary
    )

    $alreadyWindows11 = @(
        $CycleSummary |
            Where-Object { $_ -and (Test-AlreadyWindows11CycleResult -Result $_) } |
            ForEach-Object { if ($_.PSObject.Properties['ComputerName']) { [string]$_.ComputerName } } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
    )

    if ($alreadyWindows11.Count -eq 0) {
        return [pscustomobject]@{ Moved = 0; AlreadyWindows11Path = ''; Detail = 'No already-Windows11 computer detected in this cycle.' }
    }

    $moveKeys = @{}
    foreach ($computer in $alreadyWindows11) {
        $key = Get-ComputerListKey -ComputerName $computer
        if (-not [string]::IsNullOrWhiteSpace($key) -and -not $moveKeys.ContainsKey($key)) { $moveKeys[$key] = $computer.Trim() }
    }

    $listLines = @(Get-Content -LiteralPath $ComputerListPath -ErrorAction Stop)
    $remainingLines = New-Object System.Collections.Generic.List[string]
    $movedFromList = New-Object System.Collections.Generic.List[string]

    foreach ($line in $listLines) {
        $trimmed = ([string]$line).Trim().Trim([char]34)
        if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed.StartsWith('#')) {
            [void]$remainingLines.Add($line)
            continue
        }

        $key = Get-ComputerListKey -ComputerName $trimmed
        if ($moveKeys.ContainsKey($key)) {
            [void]$movedFromList.Add($trimmed)
            continue
        }

        [void]$remainingLines.Add($line)
    }

    if ($movedFromList.Count -eq 0) {
        return [pscustomobject]@{ Moved = 0; AlreadyWindows11Path = ''; Detail = 'Already-Windows11 computers were detected, but none were still present in Computers.txt.' }
    }

    $computerListDir = Split-Path -Parent $ComputerListPath
    if ([string]::IsNullOrWhiteSpace($computerListDir)) { $computerListDir = '.' }
    $alreadyWindows11Path = Join-Path $computerListDir 'ComputersAlreadyW11.txt'

    $existingKeys = @{}
    if (Test-Path -LiteralPath $alreadyWindows11Path) {
        foreach ($line in @(Get-Content -LiteralPath $alreadyWindows11Path -ErrorAction SilentlyContinue)) {
            $trimmed = ([string]$line).Trim().Trim([char]34)
            if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed.StartsWith('#')) { continue }
            $key = Get-ComputerListKey -ComputerName $trimmed
            if (-not $existingKeys.ContainsKey($key)) { $existingKeys[$key] = $true }
        }
    }

    $appendLines = New-Object System.Collections.Generic.List[string]
    foreach ($computer in $movedFromList) {
        $key = Get-ComputerListKey -ComputerName $computer
        if (-not $existingKeys.ContainsKey($key)) {
            [void]$appendLines.Add($computer)
            $existingKeys[$key] = $true
        }
    }

    $tmpComputerListPath = '{0}.tmp.{1}.txt' -f $ComputerListPath,([guid]::NewGuid().ToString('N'))
    try {
        Set-Content -LiteralPath $tmpComputerListPath -Value $remainingLines -Encoding ASCII -Force
        Move-Item -LiteralPath $tmpComputerListPath -Destination $ComputerListPath -Force
    }
    finally {
        Remove-Item -LiteralPath $tmpComputerListPath -Force -ErrorAction SilentlyContinue
    }

    if ($appendLines.Count -gt 0) {
        Add-Content -LiteralPath $alreadyWindows11Path -Value $appendLines -Encoding ASCII
    }
    elseif (-not (Test-Path -LiteralPath $alreadyWindows11Path)) {
        New-Item -ItemType File -Path $alreadyWindows11Path -Force | Out-Null
    }

    return [pscustomobject]@{
        Moved = $movedFromList.Count
        AlreadyWindows11Path = $alreadyWindows11Path
        Detail = ('Moved {0} computer(s) from Computers.txt to ComputersAlreadyW11.txt.' -f $movedFromList.Count)
    }
}

function Test-BooleanLikeTrue {
    param([AllowNull()][object]$Value)

    if ($Value -eq $true) { return $true }
    $text = ([string]$Value).Trim()
    return ($text -in @('True','true','1','YES','Yes','yes','OUI','Oui','oui'))
}

function Get-AdInventoryMap {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $false)][string]$NameColumn
    )

    $map = @{}
    if ([string]::IsNullOrWhiteSpace($Path)) { return $map }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $map }

    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { return $map }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        $candidateColumns = @('ComputerName','computerName','DNSHostName','dnsHostName','Name','name')
        $first = $rows | Select-Object -First 1
        foreach ($candidate in $candidateColumns) {
            if ($first.PSObject.Properties.Name -contains $candidate) {
                $NameColumn = $candidate
                break
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        throw 'Unable to infer the AD inventory device name column. Use -AdInventoryNameColumn.'
    }

    foreach ($row in $rows) {
        $value = [string]$row.$NameColumn
        if ([string]::IsNullOrWhiteSpace($value)) { continue }

        $short = (Get-ComputerListKey -ComputerName $value)
        if ([string]::IsNullOrWhiteSpace($short)) { continue }

        $present = $true
        if ($row.PSObject.Properties.Name -contains 'ADInventoryPresent') {
            $present = Test-BooleanLikeTrue -Value $row.ADInventoryPresent
        }
        if (-not $present) { continue }

        if (-not $map.ContainsKey($short)) { $map[$short] = $row }
    }

    return $map
}

function Invoke-FullAdInventoryExport {
    param(
        [Parameter(Mandatory = $true)][string]$ExportScriptPath,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [Parameter(Mandatory = $true)][string]$ComputerListPath,
        [Parameter(Mandatory = $false)][string]$Domain
    )

    try {
        if (-not (Test-Path -LiteralPath $ExportScriptPath -PathType Leaf)) {
            throw "SmartM365-Windows11Upgrade-Export-ADDevicesCsv.ps1 not found: $ExportScriptPath"
        }

        $args = @(
            '-NoProfile',
            '-ExecutionPolicy', 'Bypass',
            '-File', $ExportScriptPath,
            '-OutputPath', $OutputPath,
            '-ComputerListPath', $ComputerListPath,
            '-ForceRefresh'
        )
        if (-not [string]::IsNullOrWhiteSpace($Domain)) {
            $args += '-Domain'
            $args += $Domain
        }

        $output = & powershell.exe @args 2>&1
        $exitCode = $LASTEXITCODE
        $output | Out-File -LiteralPath $LogPath -Encoding UTF8 -Force

        if ($exitCode -ne 0) {
            throw "SmartM365-Windows11Upgrade-Export-ADDevicesCsv.ps1 exited with code $exitCode. Log=$LogPath"
        }
        if (-not (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
            throw "AD CSV was not created: $OutputPath"
        }

        $map = Get-AdInventoryMap -Path $OutputPath -NameColumn 'ComputerName'
        return [pscustomobject]@{ Success = $true; CsvPath = $OutputPath; LogPath = $LogPath; InventoryMap = $map; Error = '' }
    }
    catch {
        return [pscustomobject]@{ Success = $false; CsvPath = $OutputPath; LogPath = $LogPath; InventoryMap = @{}; Error = $_.Exception.Message }
    }
}

function Test-AdInventoryWindows11 {
    param([Parameter(Mandatory = $true)]$AdRow)

    $operatingSystem = if ($AdRow.PSObject.Properties['OperatingSystem']) { [string]$AdRow.OperatingSystem } else { '' }
    return ($operatingSystem -match '(?i)\bWindows\s+11\b')
}

function Get-AlreadyWindows11RowsFromAdInventory {
    param(
        [Parameter(Mandatory = $true)][string[]]$ComputerNames,
        [Parameter(Mandatory = $true)][hashtable]$AdInventoryMap,
        [Parameter(Mandatory = $true)][string]$AdInventoryCsv
    )

    foreach ($computer in $ComputerNames) {
        $key = Get-ComputerListKey -ComputerName $computer
        if ([string]::IsNullOrWhiteSpace($key) -or -not $AdInventoryMap.ContainsKey($key)) { continue }

        $adRow = $AdInventoryMap[$key]
        if (-not (Test-AdInventoryWindows11 -AdRow $adRow)) { continue }

        [pscustomobject]@{
            Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            ComputerName = $computer
            CycleNumber = 0
            LauncherStatus = 'ALREADY_WINDOWS11'
            RemoteStatus = 'ALREADY_WINDOWS11'
            RemoteNextAction = 'NO_ACTION_AD_ALREADY_WINDOWS11'
            ExitCode = 0
            Detail = 'AD inventory already reports Windows 11; PsExec launch skipped.'
            JobErrorMessage = ''
            ADInventoryPresent = $true
            ADDomain = if ($adRow.PSObject.Properties['ADDomain']) { [string]$adRow.ADDomain } else { '' }
            ADEnabled = if ($adRow.PSObject.Properties['Enabled']) { [string]$adRow.Enabled } else { '' }
            ADDNSHostName = if ($adRow.PSObject.Properties['DNSHostName']) { [string]$adRow.DNSHostName } else { '' }
            ADDistinguishedName = if ($adRow.PSObject.Properties['DistinguishedName']) { [string]$adRow.DistinguishedName } else { '' }
            ADOperatingSystem = if ($adRow.PSObject.Properties['OperatingSystem']) { [string]$adRow.OperatingSystem } else { '' }
            ADOperatingSystemVersion = if ($adRow.PSObject.Properties['OperatingSystemVersion']) { [string]$adRow.OperatingSystemVersion } else { '' }
            ADLastLogonTimestampUtc = if ($adRow.PSObject.Properties['LastLogonTimestampUtc']) { [string]$adRow.LastLogonTimestampUtc } else { '' }
            ADInventoryCsv = $AdInventoryCsv
        }
    }
}
function Add-AdInventoryFieldsToResult {
    param(
        [Parameter(Mandatory = $true)]$Result,
        [AllowNull()][hashtable]$AdInventoryMap,
        [AllowNull()][string]$AdInventoryCsv
    )

    if ($null -eq $Result) { return $Result }
    if (-not $Result.PSObject.Properties['ComputerName']) { return $Result }

    $computerKey = Get-ComputerListKey -ComputerName ([string]$Result.ComputerName)
    $hasInventory = ($null -ne $AdInventoryMap -and $AdInventoryMap.Count -gt 0)
    $hasAdRow = ($hasInventory -and -not [string]::IsNullOrWhiteSpace($computerKey) -and $AdInventoryMap.ContainsKey($computerKey))

    if ($hasAdRow) {
        $adRow = $AdInventoryMap[$computerKey]
        $fields = [ordered]@{
            ADInventoryPresent = $true
            ADDomain = if ($adRow.PSObject.Properties['ADDomain']) { [string]$adRow.ADDomain } else { '' }
            ADEnabled = if ($adRow.PSObject.Properties['Enabled']) { [string]$adRow.Enabled } else { '' }
            ADDNSHostName = if ($adRow.PSObject.Properties['DNSHostName']) { [string]$adRow.DNSHostName } else { '' }
            ADDistinguishedName = if ($adRow.PSObject.Properties['DistinguishedName']) { [string]$adRow.DistinguishedName } else { '' }
            ADOperatingSystem = if ($adRow.PSObject.Properties['OperatingSystem']) { [string]$adRow.OperatingSystem } else { '' }
            ADOperatingSystemVersion = if ($adRow.PSObject.Properties['OperatingSystemVersion']) { [string]$adRow.OperatingSystemVersion } else { '' }
            ADLastLogonTimestampUtc = if ($adRow.PSObject.Properties['LastLogonTimestampUtc']) { [string]$adRow.LastLogonTimestampUtc } else { '' }
            ADInventoryCsv = $AdInventoryCsv
        }
    }
    elseif ($hasInventory) {
        $fields = [ordered]@{
            ADInventoryPresent = $false
            ADDomain = ''
            ADEnabled = ''
            ADDNSHostName = ''
            ADDistinguishedName = ''
            ADOperatingSystem = ''
            ADOperatingSystemVersion = ''
            ADLastLogonTimestampUtc = ''
            ADInventoryCsv = $AdInventoryCsv
        }
    }
    else {
        return $Result
    }

    foreach ($fieldName in $fields.Keys) {
        $Result | Add-Member -NotePropertyName $fieldName -NotePropertyValue $fields[$fieldName] -Force
    }

    return $Result
}
function Get-IntuneInventoryMap {
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $false)][string]$NameColumn
    )

    $map = @{}
    if ([string]::IsNullOrWhiteSpace($Path)) { return $map }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $map }

    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { return $map }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        $candidateColumns = @('ComputerName','computerName','DeviceName','deviceName','ManagedDeviceName','managedDeviceName','Name','name')
        $first = $rows | Select-Object -First 1
        foreach ($candidate in $candidateColumns) {
            if ($first.PSObject.Properties.Name -contains $candidate) {
                $NameColumn = $candidate
                break
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        throw 'Unable to infer the Intune inventory device name column. Use -IntuneInventoryNameColumn.'
    }

    foreach ($row in $rows) {
        $value = [string]$row.$NameColumn
        if ([string]::IsNullOrWhiteSpace($value)) { continue }

        $short = (Get-ComputerListKey -ComputerName $value)
        if ([string]::IsNullOrWhiteSpace($short)) { continue }

        $present = $true
        if ($row.PSObject.Properties.Name -contains 'IntuneInventoryPresent') {
            $present = Test-BooleanLikeTrue -Value $row.IntuneInventoryPresent
        }
        if (-not $present) { continue }

        if (-not $map.ContainsKey($short)) {
            $map[$short] = $row
        }
        else {
            $current = $map[$short]
            $currentLastSync = [datetime]::MinValue
            $candidateLastSync = [datetime]::MinValue
            if ($current.PSObject.Properties['LastSyncDateTime']) { [datetime]::TryParse([string]$current.LastSyncDateTime, [ref]$currentLastSync) | Out-Null }
            if ($row.PSObject.Properties['LastSyncDateTime']) { [datetime]::TryParse([string]$row.LastSyncDateTime, [ref]$candidateLastSync) | Out-Null }
            if ($candidateLastSync -gt $currentLastSync) { $map[$short] = $row }
        }
    }

    return $map
}

function Invoke-FullIntuneInventoryExport {
    param(
        [Parameter(Mandatory = $true)][string]$ExportScriptPath,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][string]$LogPath,
        [Parameter(Mandatory = $true)][string]$ComputerListPath,
        [Parameter(Mandatory = $false)][int]$PageSize = 999,
        [Parameter(Mandatory = $false)][string]$TenantId
    )

    try {
        if (-not (Test-Path -LiteralPath $ExportScriptPath -PathType Leaf)) {
            throw "SmartM365-Windows11Upgrade-Export-IntuneDevicesCsv.ps1 not found: $ExportScriptPath"
        }

        $args = @(
            '-NoProfile',
            '-ExecutionPolicy', 'Bypass',
            '-File', $ExportScriptPath,
            '-OutputPath', $OutputPath,
            '-ComputerListPath', $ComputerListPath,
            '-PageSize', ([string]$PageSize),
            '-ForceRefresh'
        )
        if (-not [string]::IsNullOrWhiteSpace($TenantId)) {
            $args += '-TenantId'
            $args += $TenantId
        }

        $output = & powershell.exe @args 2>&1
        $exitCode = $LASTEXITCODE
        $output | Out-File -LiteralPath $LogPath -Encoding UTF8 -Force

        if ($exitCode -ne 0) {
            throw "SmartM365-Windows11Upgrade-Export-IntuneDevicesCsv.ps1 exited with code $exitCode. Log=$LogPath"
        }
        if (-not (Test-Path -LiteralPath $OutputPath -PathType Leaf)) {
            throw "Intune CSV was not created: $OutputPath"
        }

        $map = Get-IntuneInventoryMap -Path $OutputPath -NameColumn 'ComputerName'
        return [pscustomobject]@{ Success = $true; CsvPath = $OutputPath; LogPath = $LogPath; InventoryMap = $map; Error = '' }
    }
    catch {
        return [pscustomobject]@{ Success = $false; CsvPath = $OutputPath; LogPath = $LogPath; InventoryMap = @{}; Error = $_.Exception.Message }
    }
}

function Test-IntuneInventoryWindows11 {
    param([Parameter(Mandatory = $true)]$IntuneRow)

    $operatingSystem = if ($IntuneRow.PSObject.Properties['OperatingSystem']) { [string]$IntuneRow.OperatingSystem } else { '' }
    if ($operatingSystem -match '(?i)\bWindows\s+11\b') { return $true }

    $osVersion = if ($IntuneRow.PSObject.Properties['OSVersion']) { [string]$IntuneRow.OSVersion } else { '' }
    if ($osVersion -match '^(\d+)\.(\d+)\.(\d+)') {
        $build = 0
        if ([int]::TryParse($matches[3], [ref]$build) -and $build -ge 22000) { return $true }
    }

    return $false
}

function Get-AlreadyWindows11RowsFromIntuneInventory {
    param(
        [Parameter(Mandatory = $true)][string[]]$ComputerNames,
        [Parameter(Mandatory = $true)][hashtable]$IntuneInventoryMap,
        [Parameter(Mandatory = $true)][string]$IntuneInventoryCsv
    )

    foreach ($computer in $ComputerNames) {
        $key = Get-ComputerListKey -ComputerName $computer
        if ([string]::IsNullOrWhiteSpace($key) -or -not $IntuneInventoryMap.ContainsKey($key)) { continue }

        $intuneRow = $IntuneInventoryMap[$key]
        if (-not (Test-IntuneInventoryWindows11 -IntuneRow $intuneRow)) { continue }

        [pscustomobject]@{
            Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            ComputerName = $computer
            CycleNumber = 0
            LauncherStatus = 'ALREADY_WINDOWS11'
            RemoteStatus = 'ALREADY_WINDOWS11'
            RemoteNextAction = 'NO_ACTION_INTUNE_ALREADY_WINDOWS11'
            ExitCode = 0
            Detail = 'Intune inventory already reports Windows 11; PsExec launch skipped.'
            JobErrorMessage = ''
            IntuneInventoryPresent = $true
            IntuneDeviceName = if ($intuneRow.PSObject.Properties['DeviceName']) { [string]$intuneRow.DeviceName } else { '' }
            IntuneManagedDeviceName = if ($intuneRow.PSObject.Properties['ManagedDeviceName']) { [string]$intuneRow.ManagedDeviceName } else { '' }
            IntuneManagedDeviceId = if ($intuneRow.PSObject.Properties['IntuneManagedDeviceId']) { [string]$intuneRow.IntuneManagedDeviceId } else { '' }
            IntuneAzureADDeviceId = if ($intuneRow.PSObject.Properties['AzureADDeviceId']) { [string]$intuneRow.AzureADDeviceId } else { '' }
            IntuneOperatingSystem = if ($intuneRow.PSObject.Properties['OperatingSystem']) { [string]$intuneRow.OperatingSystem } else { '' }
            IntuneOSVersion = if ($intuneRow.PSObject.Properties['OSVersion']) { [string]$intuneRow.OSVersion } else { '' }
            IntuneLastSyncDateTime = if ($intuneRow.PSObject.Properties['LastSyncDateTime']) { [string]$intuneRow.LastSyncDateTime } else { '' }
            IntuneUserPrincipalName = if ($intuneRow.PSObject.Properties['UserPrincipalName']) { [string]$intuneRow.UserPrincipalName } else { '' }
            IntuneComplianceState = if ($intuneRow.PSObject.Properties['ComplianceState']) { [string]$intuneRow.ComplianceState } else { '' }
            IntuneManagementState = if ($intuneRow.PSObject.Properties['ManagementState']) { [string]$intuneRow.ManagementState } else { '' }
            IntuneInventoryCsv = $IntuneInventoryCsv
        }
    }
}

function Add-IntuneInventoryFieldsToResult {
    param(
        [Parameter(Mandatory = $true)]$Result,
        [AllowNull()][hashtable]$IntuneInventoryMap,
        [AllowNull()][string]$IntuneInventoryCsv
    )

    if ($null -eq $Result) { return $Result }
    if (-not $Result.PSObject.Properties['ComputerName']) { return $Result }

    $computerKey = Get-ComputerListKey -ComputerName ([string]$Result.ComputerName)
    $hasInventory = ($null -ne $IntuneInventoryMap -and $IntuneInventoryMap.Count -gt 0)
    $hasIntuneRow = ($hasInventory -and -not [string]::IsNullOrWhiteSpace($computerKey) -and $IntuneInventoryMap.ContainsKey($computerKey))

    if ($hasIntuneRow) {
        $intuneRow = $IntuneInventoryMap[$computerKey]
        $fields = [ordered]@{
            IntuneInventoryPresent = $true
            IntuneDeviceName = if ($intuneRow.PSObject.Properties['DeviceName']) { [string]$intuneRow.DeviceName } else { '' }
            IntuneManagedDeviceName = if ($intuneRow.PSObject.Properties['ManagedDeviceName']) { [string]$intuneRow.ManagedDeviceName } else { '' }
            IntuneManagedDeviceId = if ($intuneRow.PSObject.Properties['IntuneManagedDeviceId']) { [string]$intuneRow.IntuneManagedDeviceId } else { '' }
            IntuneAzureADDeviceId = if ($intuneRow.PSObject.Properties['AzureADDeviceId']) { [string]$intuneRow.AzureADDeviceId } else { '' }
            IntuneOperatingSystem = if ($intuneRow.PSObject.Properties['OperatingSystem']) { [string]$intuneRow.OperatingSystem } else { '' }
            IntuneOSVersion = if ($intuneRow.PSObject.Properties['OSVersion']) { [string]$intuneRow.OSVersion } else { '' }
            IntuneLastSyncDateTime = if ($intuneRow.PSObject.Properties['LastSyncDateTime']) { [string]$intuneRow.LastSyncDateTime } else { '' }
            IntuneUserPrincipalName = if ($intuneRow.PSObject.Properties['UserPrincipalName']) { [string]$intuneRow.UserPrincipalName } else { '' }
            IntuneComplianceState = if ($intuneRow.PSObject.Properties['ComplianceState']) { [string]$intuneRow.ComplianceState } else { '' }
            IntuneManagementState = if ($intuneRow.PSObject.Properties['ManagementState']) { [string]$intuneRow.ManagementState } else { '' }
            IntuneInventoryCsv = $IntuneInventoryCsv
        }
    }
    elseif ($hasInventory) {
        $fields = [ordered]@{
            IntuneInventoryPresent = $false
            IntuneDeviceName = ''
            IntuneManagedDeviceName = ''
            IntuneManagedDeviceId = ''
            IntuneAzureADDeviceId = ''
            IntuneOperatingSystem = ''
            IntuneOSVersion = ''
            IntuneLastSyncDateTime = ''
            IntuneUserPrincipalName = ''
            IntuneComplianceState = ''
            IntuneManagementState = ''
            IntuneInventoryCsv = $IntuneInventoryCsv
        }
    }
    else {
        return $Result
    }

    foreach ($fieldName in $fields.Keys) {
        $Result | Add-Member -NotePropertyName $fieldName -NotePropertyValue $fields[$fieldName] -Force
    }

    return $Result
}

function Test-SingleComputerLaunch {
    param([Parameter(Mandatory = $true)][string]$Path)

    $fullPath = [System.IO.Path]::GetFullPath($Path)
    if ($fullPath -match '\\SingleComputer(Runs)?\\') { return $true }

    $computers = @(Get-ComputerList -Path $fullPath)
    return ($computers.Count -eq 1)
}

$script:IsSingleComputerLaunch = Test-SingleComputerLaunch -Path $ComputerListPath
if ($script:IsSingleComputerLaunch) {
    $ThrottleLimit = 1
    $GlobalConcurrencyLimit = 0
}

function Write-LauncherLogLine {
    param([AllowNull()][string]$Line)

    foreach ($target in @($script:LauncherLogPath, $script:LauncherLatestLogPath)) {
        if ([string]::IsNullOrWhiteSpace($target)) { continue }
        try { Add-Content -LiteralPath $target -Value ([string]$Line) -Encoding UTF8 -ErrorAction Stop }
        catch { }
    }
}

function Write-Host {
    param(
        [Parameter(Position = 0)][object]$Object = '',
        [switch]$NoNewline,
        [System.ConsoleColor]$ForegroundColor,
        [System.ConsoleColor]$BackgroundColor
    )
    $ts = if ($null -ne $Object -and '' -ne [string]$Object) { "[{0}] " -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss') } else { '' }
    $line = $ts + [string]$Object
    $p = @{ Object = $line }
    if ($PSBoundParameters.ContainsKey('ForegroundColor')) { $p['ForegroundColor'] = $ForegroundColor }
    if ($PSBoundParameters.ContainsKey('BackgroundColor')) { $p['BackgroundColor'] = $BackgroundColor }
    if ($PSBoundParameters.ContainsKey('NoNewline')) { $p['NoNewline'] = $NoNewline }
    Microsoft.PowerShell.Utility\Write-Host @p
    Write-LauncherLogLine -Line $line
}

function Test-SetupSourceMapSyntax {
    param([AllowNull()][string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return }

    $expanded = [Environment]::ExpandEnvironmentVariables($Path.Trim('"'))
    if (-not (Test-Path -LiteralPath $expanded -PathType Leaf)) {
        Write-Host ("Setup source map preflight skipped because the operator cannot read it: {0}" -f $expanded) -ForegroundColor DarkYellow
        return
    }

    $rows = @(Import-Csv -LiteralPath $expanded -ErrorAction Stop)
    if ($rows.Count -eq 0) {
        throw "Setup source map is empty: $expanded"
    }

    $validRows = 0
    $rowNumber = 1
    foreach ($row in $rows) {
        $rowNumber++
        $scopeType = if ($row.PSObject.Properties['ScopeType']) { [string]$row.ScopeType } elseif ($row.PSObject.Properties['Type']) { [string]$row.Type } else { '' }
        $scopeValue = if ($row.PSObject.Properties['ScopeValue']) { [string]$row.ScopeValue } elseif ($row.PSObject.Properties['Value']) { [string]$row.Value } else { '' }
        $sourcePath = if ($row.PSObject.Properties['SetupSourcePath']) { [string]$row.SetupSourcePath } elseif ($row.PSObject.Properties['SourcePath']) { [string]$row.SourcePath } else { '' }

        if ([string]::IsNullOrWhiteSpace($scopeType) -or [string]::IsNullOrWhiteSpace($sourcePath)) {
            throw "Invalid setup source map row $rowNumber. ScopeType and SetupSourcePath are required."
        }

        if ($scopeType -notmatch '^(?i:Default|Fallback|Subnet|CIDR|IPPrefix|Prefix|ComputerName|Hostname|ComputerPrefix|HostnamePrefix)$') {
            throw "Invalid setup source map row $rowNumber. Unsupported ScopeType '$scopeType'."
        }

        if ($scopeType -match '^(?i:Subnet|CIDR)$' -and $scopeValue -notmatch '^\d{1,3}(\.\d{1,3}){3}/([0-9]|[12][0-9]|3[0-2])$') {
            throw "Invalid setup source map row $rowNumber. ScopeValue must be CIDR for ScopeType '$scopeType'."
        }

        if ($sourcePath -notmatch '^\\\\[^\\]+\\[^\\]+') {
            throw "Invalid setup source map row $rowNumber. SetupSourcePath must be UNC: $sourcePath"
        }

        $validRows++
    }

    Write-Host ("Setup source map preflight OK: {0}; Rows={1}" -f $expanded,$validRows) -ForegroundColor Green
}

function Resolve-PsExecPath {
    param([string]$Path)

    $candidate = [string]$Path
    if (-not [string]::IsNullOrWhiteSpace($candidate)) {
        $embedded = @([regex]::Matches($candidate, '(?i)(?:[A-Z]:\\|\\\\)[^"<>|?*]+?PsExec\.exe') | ForEach-Object { $_.Value.Trim() })
        if ($embedded.Count -gt 1) {
            throw ("Invalid PsExecPath value: {0}. Multiple PsExec paths were provided: {1}." -f $Path,($embedded -join ' | '))
        }
        if ($embedded.Count -eq 1) { $candidate = $embedded[0] }
        $candidate = $candidate.Trim('"')
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return (Get-Item -LiteralPath $candidate).FullName
        }
        throw ("Invalid PsExecPath value: {0}. Provide exactly one executable path." -f $Path)
    }

    $local = Join-Path $script:BaseDir 'PsExec.exe'
    if (Test-Path -LiteralPath $local -PathType Leaf) { return (Get-Item -LiteralPath $local).FullName }
    $command = Get-Command -Name 'PsExec.exe' -CommandType Application -ErrorAction SilentlyContinue
    if ($command) { return $command.Source }
    throw ("PsExec.exe not found. Place it in {0} or add PsExec.exe to PATH." -f $script:BaseDir)
}

if (-not (Test-Path -LiteralPath $LocalScriptPath -PathType Leaf)) {
    throw "Local repair script not found: $LocalScriptPath"
}
if (-not (Test-Path -LiteralPath $LocalWorkerPath -PathType Leaf)) {
    throw "Local PsExec worker script not found: $LocalWorkerPath"
}

New-Directory -Path $LogRoot
New-Directory -Path $ReportRoot
New-Directory -Path $CentralLogRoot
New-Directory -Path $script:LauncherLogRoot
$launcherLogHeader = "[{0}] ===== SmartM365 Windows 11 Upgrade Toolkit launcher v{1} started. Lot={2}; ComputerList={3}; Technician={4}; TechComputer={5} =====" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$script:LauncherVersion,(Split-Path -Leaf $script:LotRoot),$ComputerListPath,$script:TechnicianIdentity.Account,$script:TechnicianIdentity.ComputerName
Set-Content -LiteralPath $script:LauncherLogPath -Value $launcherLogHeader -Encoding UTF8 -Force
Set-Content -LiteralPath $script:LauncherLatestLogPath -Value $launcherLogHeader -Encoding UTF8 -Force

$script:TechnicianRunGuardHistoryPath = Get-TechnicianRunGuardHistoryPath
$script:UseEffectiveTechnicianRunGuardHistory = ($UseTechnicianRunGuardHistory -and -not $IgnoreTechnicianRunGuardHistory -and -not $IgnoreRunGuard -and -not $DryRun -and $RunGuardHours -gt 0)
if ($script:UseEffectiveTechnicianRunGuardHistory) {
    try {
        New-Directory -Path (Split-Path -Parent $script:TechnicianRunGuardHistoryPath)
        Invoke-TechnicianRunGuardHistoryLock -ArgumentList @($script:TechnicianRunGuardHistoryPath, $RunGuardHours) -ScriptBlock {
            param($LockedPath, $LockedRunGuardHours)
            $history = Read-TechnicianRunGuardHistory -Path $LockedPath -RunGuardHours $LockedRunGuardHours
            Save-TechnicianRunGuardHistory -Path $LockedPath -History $history
        }
    }
    catch {
        Write-Host ("Technician run guard history disabled: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
        $script:UseEffectiveTechnicianRunGuardHistory = $false
    }
}
Test-SetupSourceMapSyntax -Path $SetupSourceMapPath

$resolvedPsExec = if ($DryRun) { '' } else { Resolve-PsExecPath -Path $PsExecPath }

$remoteArgs = New-Object System.Collections.ArrayList
if ($AuditOnly) { [void]$remoteArgs.Add('-AuditOnly') }
if ($IgnoreRunGuard) { [void]$remoteArgs.Add('-IgnoreRunGuard') }
[void]$remoteArgs.Add('-RunGuardHours'); [void]$remoteArgs.Add([string]$RunGuardHours)
if ($AllowPolicyRepair) { [void]$remoteArgs.Add('-AllowPolicyRepair') }
if ($AllowWUReset) { [void]$remoteArgs.Add('-AllowWUReset') }
if ($AllowForceUpgrade) { [void]$remoteArgs.Add('-AllowForceUpgrade') }
if ($AllowSetupUpgrade) { [void]$remoteArgs.Add('-AllowSetupUpgrade') }
if ($DirectSetupUpgrade) { [void]$remoteArgs.Add('-DirectSetupUpgrade') }
if ($AllowReboot) { [void]$remoteArgs.Add('-AllowReboot') }
if ($AllowSetupCompletionRebootWhenNoUser) { [void]$remoteArgs.Add('-AllowSetupCompletionRebootWhenNoUser') }
if ($AllowSetupProfileRepair) { [void]$remoteArgs.Add('-AllowSetupProfileRepair') }
if ($SkipVirtualMachines) { [void]$remoteArgs.Add('-SkipVirtualMachines') }
if ($AllowDiskCleanup) { [void]$remoteArgs.Add('-AllowDiskCleanup') }
if ($AllowAdvancedDiskCleanup -or $AllowDismComponentCleanup) { [void]$remoteArgs.Add('-AllowAdvancedDiskCleanup') }
if ($SkipSetupMediaPreCopy) { [void]$remoteArgs.Add('-SkipSetupMediaPreCopy') }
[void]$remoteArgs.Add('-SetupExecutionMode'); [void]$remoteArgs.Add($SetupExecutionMode)
[void]$remoteArgs.Add('-SetupMediaId'); [void]$remoteArgs.Add($SetupMediaId)
[void]$remoteArgs.Add('-SetupLanguage'); [void]$remoteArgs.Add($SetupLanguage)
[void]$remoteArgs.Add('-SetupDynamicUpdate'); [void]$remoteArgs.Add($SetupDynamicUpdate)
[void]$remoteArgs.Add('-SetupCacheRoot'); [void]$remoteArgs.Add($script:RemoteSetupCacheRoot)
[void]$remoteArgs.Add('-SetupSourceCandidateLimit'); [void]$remoteArgs.Add([string]$SetupSourceCandidateLimit)
[void]$remoteArgs.Add('-SetupMediaCopyIpGapMilliseconds'); [void]$remoteArgs.Add([string]$SetupMediaCopyIpGapMilliseconds)
[void]$remoteArgs.Add('-SetupMediaCopyJitterSeconds'); [void]$remoteArgs.Add([string]$SetupMediaCopyJitterSeconds)
if ($SetupSourceConcurrencyLimit -gt 0) {
    [void]$remoteArgs.Add('-SetupSourceConcurrencyLimit'); [void]$remoteArgs.Add([string]$SetupSourceConcurrencyLimit)
    [void]$remoteArgs.Add('-SetupSourceConcurrencyLeaseMinutes'); [void]$remoteArgs.Add([string]$SetupSourceConcurrencyLeaseMinutes)
}
if (-not [string]::IsNullOrWhiteSpace($SetupSourceConcurrencyGateRoot)) {
    [void]$remoteArgs.Add('-SetupSourceConcurrencyGateRoot'); [void]$remoteArgs.Add($SetupSourceConcurrencyGateRoot)
}
if ($SetupSubnetConcurrencyLimit -gt 0) {
    [void]$remoteArgs.Add('-SetupSubnetConcurrencyLimit'); [void]$remoteArgs.Add([string]$SetupSubnetConcurrencyLimit)
    [void]$remoteArgs.Add('-SetupSubnetPrefixLength'); [void]$remoteArgs.Add([string]$SetupSubnetPrefixLength)
    [void]$remoteArgs.Add('-SetupSubnetConcurrencyLeaseMinutes'); [void]$remoteArgs.Add([string]$SetupSubnetConcurrencyLeaseMinutes)
}
if (-not [string]::IsNullOrWhiteSpace($SetupSubnetConcurrencyGateRoot)) {
    [void]$remoteArgs.Add('-SetupSubnetConcurrencyGateRoot'); [void]$remoteArgs.Add($SetupSubnetConcurrencyGateRoot)
}
if (-not [string]::IsNullOrWhiteSpace($SetupSourcePath)) {
    [void]$remoteArgs.Add('-SetupSourcePath'); [void]$remoteArgs.Add($SetupSourcePath)
}
if (-not [string]::IsNullOrWhiteSpace($SetupSourceMapPath)) {
    [void]$remoteArgs.Add('-SetupSourceMapPath'); [void]$remoteArgs.Add($SetupSourceMapPath)
}

$script:LauncherOptionRows = @(
    [pscustomobject]@{ Category = 'Operator'; Option = 'TechnicianAccount'; Value = [string]$script:TechnicianIdentity.Account }
    [pscustomobject]@{ Category = 'Operator'; Option = 'TechnicianUPN'; Value = [string]$script:TechnicianIdentity.UserPrincipalName }
    [pscustomobject]@{ Category = 'Operator'; Option = 'TechnicianSID'; Value = [string]$script:TechnicianIdentity.Sid }
    [pscustomobject]@{ Category = 'Operator'; Option = 'TechnicianAuthType'; Value = [string]$script:TechnicianIdentity.AuthenticationType }
    [pscustomobject]@{ Category = 'Operator'; Option = 'TechnicianComputer'; Value = [string]$script:TechnicianIdentity.ComputerName }
    [pscustomobject]@{ Category = 'Mode'; Option = 'DryRun'; Value = [string][bool]$DryRun }
    [pscustomobject]@{ Category = 'Mode'; Option = 'AuditOnly'; Value = [string][bool]$AuditOnly }
    [pscustomobject]@{ Category = 'Mode'; Option = 'RunOnce'; Value = [string][bool]$RunOnce }
    [pscustomobject]@{ Category = 'Mode'; Option = 'IgnoreRunGuard'; Value = [string][bool]$IgnoreRunGuard }
    [pscustomobject]@{ Category = 'Mode'; Option = 'RunGuardHours'; Value = [string]$RunGuardHours }
    [pscustomobject]@{ Category = 'Mode'; Option = 'UseTechnicianRunGuardHistory'; Value = [string][bool]$UseTechnicianRunGuardHistory }
    [pscustomobject]@{ Category = 'Mode'; Option = 'IgnoreTechnicianRunGuardHistory'; Value = [string][bool]$IgnoreTechnicianRunGuardHistory }
    [pscustomobject]@{ Category = 'Mode'; Option = 'EffectiveTechnicianRunGuardHistory'; Value = [string][bool]$script:UseEffectiveTechnicianRunGuardHistory }
    [pscustomobject]@{ Category = 'Mode'; Option = 'TechnicianRunGuardHistoryPath'; Value = [string]$script:TechnicianRunGuardHistoryPath }
    [pscustomobject]@{ Category = 'Actions'; Option = 'AllowPolicyRepair'; Value = [string][bool]$AllowPolicyRepair }
    [pscustomobject]@{ Category = 'Actions'; Option = 'AllowWUReset'; Value = [string][bool]$AllowWUReset }
    [pscustomobject]@{ Category = 'Actions'; Option = 'AllowForceUpgrade'; Value = [string][bool]$AllowForceUpgrade }
    [pscustomobject]@{ Category = 'Actions'; Option = 'AllowSetupUpgrade'; Value = [string][bool]$AllowSetupUpgrade }
    [pscustomobject]@{ Category = 'Actions'; Option = 'DirectSetupUpgrade'; Value = [string][bool]$DirectSetupUpgrade }
    [pscustomobject]@{ Category = 'Actions'; Option = 'AllowReboot'; Value = [string][bool]$AllowReboot }
    [pscustomobject]@{ Category = 'Actions'; Option = 'AllowSetupCompletionRebootWhenNoUser'; Value = [string][bool]$AllowSetupCompletionRebootWhenNoUser }
    [pscustomobject]@{ Category = 'Actions'; Option = 'AllowSetupProfileRepair'; Value = [string][bool]$AllowSetupProfileRepair }
    [pscustomobject]@{ Category = 'Actions'; Option = 'SkipVirtualMachines'; Value = [string][bool]$SkipVirtualMachines }
    [pscustomobject]@{ Category = 'Actions'; Option = 'AllowDiskCleanup'; Value = [string][bool]$AllowDiskCleanup }
    [pscustomobject]@{ Category = 'Actions'; Option = 'AllowAdvancedDiskCleanup'; Value = [string][bool]($AllowAdvancedDiskCleanup -or $AllowDismComponentCleanup) }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSourcePath'; Value = [string]$SetupSourcePath }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSourceMapPath'; Value = [string]$SetupSourceMapPath }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupExecutionMode'; Value = [string]$SetupExecutionMode }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupMediaId'; Value = [string]$SetupMediaId }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupLanguage'; Value = [string]$SetupLanguage }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupDynamicUpdate'; Value = [string]$SetupDynamicUpdate }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SkipSetupMediaPreCopy'; Value = [string][bool]$SkipSetupMediaPreCopy }
    [pscustomobject]@{ Category = 'Setup'; Option = 'TargetSetupCacheRoot'; Value = [string]$script:RemoteSetupCacheRoot }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSourceCandidateLimit'; Value = [string]$SetupSourceCandidateLimit }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupMediaCopyIpGapMilliseconds'; Value = [string]$SetupMediaCopyIpGapMilliseconds }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupMediaCopyJitterSeconds'; Value = [string]$SetupMediaCopyJitterSeconds }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSourceConcurrencyLimit'; Value = [string]$SetupSourceConcurrencyLimit }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSourceConcurrencyLeaseMinutes'; Value = [string]$SetupSourceConcurrencyLeaseMinutes }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSourceConcurrencyGateRoot'; Value = [string]$SetupSourceConcurrencyGateRoot }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSubnetConcurrencyLimit'; Value = [string]$SetupSubnetConcurrencyLimit }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSubnetPrefixLength'; Value = [string]$SetupSubnetPrefixLength }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSubnetConcurrencyLeaseMinutes'; Value = [string]$SetupSubnetConcurrencyLeaseMinutes }
    [pscustomobject]@{ Category = 'Setup'; Option = 'SetupSubnetConcurrencyGateRoot'; Value = [string]$SetupSubnetConcurrencyGateRoot }
    [pscustomobject]@{ Category = 'AD'; Option = 'AdInventoryCsv'; Value = [string]$AdInventoryCsv }
    [pscustomobject]@{ Category = 'AD'; Option = 'AdRootInventoryCsv'; Value = [string]$AdRootInventoryCsv }
    [pscustomobject]@{ Category = 'AD'; Option = 'AdDomain'; Value = [string]$AdDomain }
    [pscustomobject]@{ Category = 'AD'; Option = 'SkipAdInventoryRefresh'; Value = [string][bool]$SkipAdInventoryRefresh }
    [pscustomobject]@{ Category = 'Intune'; Option = 'IntuneInventoryCsv'; Value = [string]$IntuneInventoryCsv }
    [pscustomobject]@{ Category = 'Intune'; Option = 'IntuneRootInventoryCsv'; Value = [string]$IntuneRootInventoryCsv }
    [pscustomobject]@{ Category = 'Intune'; Option = 'IntuneInventoryNameColumn'; Value = [string]$IntuneInventoryNameColumn }
    [pscustomobject]@{ Category = 'Intune'; Option = 'IntuneInventoryPageSize'; Value = [string]$IntuneInventoryPageSize }
    [pscustomobject]@{ Category = 'Intune'; Option = 'IntuneTenantId'; Value = [string]$IntuneTenantId }
    [pscustomobject]@{ Category = 'Intune'; Option = 'SkipIntuneInventoryRefresh'; Value = [string][bool]$SkipIntuneInventoryRefresh }
    [pscustomobject]@{ Category = 'Parallelism'; Option = 'ThrottleLimit'; Value = [string]$ThrottleLimit }
    [pscustomobject]@{ Category = 'Parallelism'; Option = 'GlobalConcurrencyLimit'; Value = [string]$GlobalConcurrencyLimit }
    [pscustomobject]@{ Category = 'Parallelism'; Option = 'GlobalConcurrencyLeaseTimeoutMinutes'; Value = [string]$GlobalConcurrencyLeaseTimeoutMinutes }
    [pscustomobject]@{ Category = 'Parallelism'; Option = 'DelayBetweenComputersSeconds'; Value = [string]$DelayBetweenComputersSeconds }
    [pscustomobject]@{ Category = 'Parallelism'; Option = 'DelayBetweenCyclesMinutes'; Value = [string]$DelayBetweenCyclesMinutes }
    [pscustomobject]@{ Category = 'Parallelism'; Option = 'MaxCycles'; Value = [string]$MaxCycles }
    [pscustomobject]@{ Category = 'Parallelism'; Option = 'PsExecTimeoutMinutes'; Value = [string]$PsExecTimeoutMinutes }
    [pscustomobject]@{ Category = 'Paths'; Option = 'ComputerListPath'; Value = [string]$ComputerListPath }
    [pscustomobject]@{ Category = 'Paths'; Option = 'LogRoot'; Value = [string]$LogRoot }
    [pscustomobject]@{ Category = 'Paths'; Option = 'ReportRoot'; Value = [string]$ReportRoot }
    [pscustomobject]@{ Category = 'Paths'; Option = 'CentralLogRoot'; Value = [string]$CentralLogRoot }
    [pscustomobject]@{ Category = 'Paths'; Option = 'LauncherLogRoot'; Value = [string]$script:LauncherLogRoot }
    [pscustomobject]@{ Category = 'Paths'; Option = 'LauncherLogPath'; Value = [string]$script:LauncherLogPath }
    [pscustomobject]@{ Category = 'Paths'; Option = 'LauncherLatestLogPath'; Value = [string]$script:LauncherLatestLogPath }
    [pscustomobject]@{ Category = 'Paths'; Option = 'NoCentralLogCollection'; Value = [string][bool]$NoCentralLogCollection }
    [pscustomobject]@{ Category = 'Paths'; Option = 'KeepCentralLogHistory'; Value = [string][bool]$KeepCentralLogHistory }
    [pscustomobject]@{ Category = 'Paths'; Option = 'CentralLogCollectionMode'; Value = [string]$CentralLogCollectionMode }
)

Write-Host "SmartM365 Windows 11 Upgrade Toolkit launcher v$script:LauncherVersion"
Write-Host "LOT root      : $script:LotRoot"
Write-Host "Launcher log  : $script:LauncherLogPath"
Write-Host "Computer list : $ComputerListPath"
Write-Host "PsExec        : $resolvedPsExec"
Write-Host "Repair script : $LocalScriptPath"
Write-Host "Worker script : $LocalWorkerPath"
Write-Host ("Technician     : Account={0}; UPN={1}; SID={2}; Auth={3}; Computer={4}" -f $script:TechnicianIdentity.Account,$script:TechnicianIdentity.UserPrincipalName,$script:TechnicianIdentity.Sid,$script:TechnicianIdentity.AuthenticationType,$script:TechnicianIdentity.ComputerName)
Write-Host "Mode          : DryRun=$DryRun; AuditOnly=$AuditOnly; RunOnce=$RunOnce; SkipVirtualMachines=$SkipVirtualMachines; DiskCleanup=$AllowDiskCleanup; AdvancedCleanup=$($AllowAdvancedDiskCleanup -or $AllowDismComponentCleanup); DirectSetup=$DirectSetupUpgrade; SetupCompletionRebootWhenNoUser=$AllowSetupCompletionRebootWhenNoUser; SetupProfileRepair=$AllowSetupProfileRepair"
Write-Host "Setup         : AllowSetupUpgrade=$AllowSetupUpgrade; DirectSetup=$DirectSetupUpgrade; Effective=$([bool]($AllowSetupUpgrade -or $DirectSetupUpgrade)); Mode=$SetupExecutionMode; MediaId=$SetupMediaId; Language=$SetupLanguage; DynamicUpdate=$SetupDynamicUpdate; PreCopy=$(-not $SkipSetupMediaPreCopy)"
Write-Host "AD inventory  : Csv=$AdInventoryCsv; RootCsv=$AdRootInventoryCsv; Domain=$AdDomain; Refresh=$(-not $SkipAdInventoryRefresh); RecentRoot=$AdInventoryUsesRecentRootCsv"
Write-Host "Intune invent.: Csv=$IntuneInventoryCsv; RootCsv=$IntuneRootInventoryCsv; Tenant=$IntuneTenantId; Refresh=$(-not $SkipIntuneInventoryRefresh); RecentRoot=$IntuneInventoryUsesRecentRootCsv"
Write-Host "Tech run guard: Use=$script:UseEffectiveTechnicianRunGuardHistory; Requested=$UseTechnicianRunGuardHistory; Ignore=$IgnoreTechnicianRunGuardHistory; Hours=$RunGuardHours; Path=$script:TechnicianRunGuardHistoryPath"
Write-Host "Parallelism   : ThrottleLimit=$ThrottleLimit; GlobalConcurrencyLimit=$GlobalConcurrencyLimit; GlobalLeaseTimeout=$GlobalConcurrencyLeaseTimeoutMinutes minute(s)"
Write-Host "LOT/run options:"
foreach ($optionRow in @($script:LauncherOptionRows)) {
    Write-Host ("  [{0}] {1}={2}" -f $optionRow.Category,$optionRow.Option,$optionRow.Value)
}
if ($script:IsSingleComputerLaunch) { Write-Host "Single PC     : worker limits ignored for one-computer launch." }
Write-Host "Reports       : $ReportRoot"
Write-Host ""

$reportColumns = @(
    'Timestamp',
    'ComputerName',
    'CycleNumber',
    'LauncherStatus',
    'RemoteStatus',
    'RemoteNextAction',
    'ExitCode',
    'Detail',
    'JobErrorMessage',
    'ADInventoryPresent',
    'ADDomain',
    'ADEnabled',
    'ADDNSHostName',
    'ADDistinguishedName',
    'ADOperatingSystem',
    'ADOperatingSystemVersion',
    'ADLastLogonTimestampUtc',
    'ADInventoryCsv',
    'IntuneInventoryPresent',
    'IntuneDeviceName',
    'IntuneManagedDeviceName',
    'IntuneManagedDeviceId',
    'IntuneAzureADDeviceId',
    'IntuneOperatingSystem',
    'IntuneOSVersion',
    'IntuneLastSyncDateTime',
    'IntuneUserPrincipalName',
    'IntuneComplianceState',
    'IntuneManagementState',
    'IntuneInventoryCsv',
    'SetupCacheAction',
    'SetupDynamicUpdate',
    'SelectedSetupSourcePath',
    'SetupSourceSelectionDetail',
    'DiskCleanupAction',
    'DiskCleanupFreedGB',
    'AdvancedDiskCleanupAction',
    'AdvancedDiskCleanupFreedGB',
    'DismCleanupAction',
    'DismCleanupFreedGB',
    'SetupCompletionRebootAction',
    'SetupCompletionRebootDetail',
    'SetupCompletionRebootUserCount',
    'SetupCompletionRebootUsers',
    'SetupProfileRepairAction',
    'SetupProfileRepairDetail',
    'SetupProfileRepairBlockingSid',
    'SetupProfileRepairKeptSid',
    'SetupProfileRepairProfilePath',
    'SetupProfileRepairBackupPath',
    'ControlledRebootAction',
    'ControlledRebootDetail',
    'ControlledRebootUserCount',
    'ControlledRebootUsers',
    'UserRebootNotificationSent',
    'UserRebootNotificationLang',
    'UserRebootNotificationMessage',
    'RemoteLogsPath',
    'PsExecLogPath'
)

function ConvertTo-HtmlText {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    return [System.Security.SecurityElement]::Escape([string]$Value)
}

function ConvertTo-FileUri {
    param([AllowNull()][string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }

    $value = [string]$Path
    try {
        if ($value -match '^\\\\(?<Server>[^\\]+)\\(?<Share>[^\\]+)(?<Rest>.*)$') {
            $server = $Matches.Server
            $share = $Matches.Share
            $rest = ($Matches.Rest -replace '\\','/')
            return ('file://{0}/{1}{2}' -f $server,$share,$rest)
        }
        return ([System.Uri]::new([System.IO.Path]::GetFullPath($value))).AbsoluteUri
    }
    catch {
        return ''
    }
}

function New-HtmlLogLink {
    param(
        [AllowNull()][string]$Path,
        [string]$Label = 'Open'
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return '-' }
    $uri = ConvertTo-FileUri -Path $Path
    if ([string]::IsNullOrWhiteSpace($uri)) { return (ConvertTo-HtmlText $Path) }
    return ('<a href="{0}" title="{1}">{2}</a>' -f (ConvertTo-HtmlText $uri),(ConvertTo-HtmlText $Path),(ConvertTo-HtmlText $Label))
}

function Get-RemotePcLogsPath {
    param([AllowNull()][string]$ComputerName)
    if ([string]::IsNullOrWhiteSpace($ComputerName)) { return '' }

    $computer = ([string]$ComputerName).Trim()
    if ([string]::IsNullOrWhiteSpace($computer)) { return '' }
    return ('\\{0}\C$\ProgramData\SmartM365\Windows11UpgradeToolkit\Logs' -f $computer)
}

function ConvertTo-SimpleHtmlTable {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rows,
        [string[]]$Columns
    )

    if (-not $Rows -or $Rows.Count -eq 0) { return "<p class='empty'>No rows.</p>" }

    if (-not $Columns -or $Columns.Count -eq 0) {
        $columnSet = New-Object System.Collections.Generic.List[string]
        foreach ($row in $Rows) {
            foreach ($property in $row.PSObject.Properties) {
                if (-not $columnSet.Contains($property.Name)) { [void]$columnSet.Add($property.Name) }
            }
        }
        $Columns = @($columnSet)
    }

    $html = New-Object System.Collections.Generic.List[string]
    [void]$html.Add("<table>")
    [void]$html.Add("<tr>")
    foreach ($column in $Columns) { [void]$html.Add(("<th>{0}</th>" -f (ConvertTo-HtmlText $column))) }
    [void]$html.Add("</tr>")
    foreach ($row in $Rows) {
        [void]$html.Add("<tr>")
        foreach ($column in $Columns) {
            $value = ""
            $property = $row.PSObject.Properties[$column]
            if ($property) { $value = $property.Value }
            if ($column -eq 'Local log' -or $column -eq 'Collected logs' -or $column -eq 'Remote PC logs') {
                [void]$html.Add(("<td>{0}</td>" -f $value))
            }
            else {
                [void]$html.Add(("<td>{0}</td>" -f (ConvertTo-HtmlText $value)))
            }
        }
        [void]$html.Add("</tr>")
    }
    [void]$html.Add("</table>")
    return ($html -join "`r`n")
}

function Get-Windows11ReportRows {
    param([AllowEmptyCollection()][object[]]$Items)
    $normalized = foreach ($item in @($Items)) {
        $row = [ordered]@{}
        foreach ($column in $reportColumns) {
            $row[$column] = if ($null -ne $item -and $item.PSObject.Properties[$column]) { [string]$item.$column } else { '' }
        }
        $row['Local log'] = New-HtmlLogLink -Path $row['PsExecLogPath']
        $row['Collected logs'] = New-HtmlLogLink -Path $row['RemoteLogsPath']
        $row['Remote PC logs'] = New-HtmlLogLink -Path (Get-RemotePcLogsPath -ComputerName $row['ComputerName'])
        [pscustomobject]$row
    }
    return @($normalized)
}


function Get-Windows11HtmlEffectiveStatus {
    param([AllowNull()][object]$Row)

    if ($null -eq $Row) { return '' }

    $remoteStatus = ''
    $launcherStatus = ''
    if ($Row.PSObject.Properties['RemoteStatus']) { $remoteStatus = [string]$Row.RemoteStatus }
    if ($Row.PSObject.Properties['LauncherStatus']) { $launcherStatus = [string]$Row.LauncherStatus }

    if (-not [string]::IsNullOrWhiteSpace($remoteStatus)) { return $remoteStatus }
    return $launcherStatus
}

function Get-Windows11AdminShareFailureType {
    param([AllowNull()][object]$Row)

    if ($null -eq $Row) { return 'UNKNOWN' }
    $detail = if ($Row.PSObject.Properties['Detail']) { [string]$Row.Detail } else { '' }
    $jobErrorMessage = if ($Row.PSObject.Properties['JobErrorMessage']) { [string]$Row.JobErrorMessage } else { '' }
    $evidence = @($detail, $jobErrorMessage) -join ' '
    if ($evidence -match 'FailureType=(?<FailureType>[A-Z0-9_]+)') { return $Matches.FailureType }
    if ($evidence -match 'DNS.*False|DNS_FAILED|host.*not.*found|nom.*introuvable') { return 'DNS_FAILED' }
    if ($evidence -match 'SMB_PORT_445_UNREACHABLE|Tcp445=False|port\s+445') { return 'SMB_PORT_445_UNREACHABLE' }
    if ($evidence -match 'AdminShare=False|Access is denied|Acc.s refus|administrative shares are not reachable') { return 'ADMIN_SHARE_UNREACHABLE' }
    return 'UNKNOWN'
}

function New-Windows11CycleProgressRows {
    param(
        [Parameter(Mandatory = $true)][int]$CycleNumber,
        [Parameter(Mandatory = $true)][datetime]$CycleStart,
        [Parameter(Mandatory = $true)][int]$TotalComputers,
        [Parameter(Mandatory = $true)][int]$QueuedComputers,
        [Parameter(Mandatory = $true)][int]$CompletedComputers,
        [Parameter(Mandatory = $true)][int]$RunningComputers,
        [AllowNull()]$ComputerListStats
    )

    $remaining = [math]::Max(0, $TotalComputers - $CompletedComputers - $RunningComputers)
    $duplicateGroups = if ($ComputerListStats -and $ComputerListStats.PSObject.Properties['DuplicateGroups']) { [int]$ComputerListStats.DuplicateGroups } else { 0 }
    $duplicateLines = if ($ComputerListStats -and $ComputerListStats.PSObject.Properties['DuplicateLines']) { [int]$ComputerListStats.DuplicateLines } else { 0 }
    $duplicateSamples = if ($ComputerListStats -and $ComputerListStats.PSObject.Properties['DuplicateSamples']) { [string]$ComputerListStats.DuplicateSamples } else { '' }
    $rawLines = if ($ComputerListStats -and $ComputerListStats.PSObject.Properties['RawLines']) { [int]$ComputerListStats.RawLines } else { $TotalComputers }

    return @([pscustomobject]@{
        Cycle = $CycleNumber
        Started = $CycleStart.ToString('yyyy-MM-dd HH:mm:ss')
        ElapsedMinutes = [math]::Round(((Get-Date) - $CycleStart).TotalMinutes, 1)
        ComputerListLines = $rawLines
        TotalUnique = $TotalComputers
        Queued = $QueuedComputers
        CompletedRows = $CompletedComputers
        Running = $RunningComputers
        Remaining = $remaining
        DuplicateGroups = $duplicateGroups
        DuplicateLines = $duplicateLines
        DuplicateSamples = $duplicateSamples
    })
}

function New-Windows11RunningJobRows {
    param(
        [AllowEmptyCollection()][object[]]$RunningJobs,
        [Parameter(Mandatory = $true)][hashtable]$JobStartedAtById
    )

    $now = Get-Date
    return @($RunningJobs | ForEach-Object {
        $jobId = [string]$_.Id
        $computer = ([string]$_.Name) -replace '^W11UT_C\d+_',''
        $started = if ($JobStartedAtById.ContainsKey($jobId)) { [datetime]$JobStartedAtById[$jobId] } else { [datetime]::MinValue }
        [pscustomobject]@{
            ComputerName = $computer
            JobId = $jobId
            State = [string]$_.State
            Started = if ($started -gt [datetime]::MinValue) { $started.ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
            ElapsedMinutes = if ($started -gt [datetime]::MinValue) { [math]::Round(($now - $started).TotalMinutes, 1) } else { '' }
        }
    })
}

function Export-Windows11ReportCsv {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory = $true)][string]$Path
    )

    if ($Rows -and $Rows.Count -gt 0) {
        @($Rows) | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
        return
    }

    $emptyRow = [ordered]@{}
    foreach ($column in $reportColumns) { $emptyRow[$column] = '' }
    $emptyRow['Local log'] = ''
    $emptyRow['Collected logs'] = ''
    $emptyRow['Remote PC logs'] = ''
    @([pscustomobject]$emptyRow) | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
    $lines = @(Get-Content -LiteralPath $Path -ErrorAction Stop)
    if ($lines.Count -gt 0) { Set-Content -LiteralPath $Path -Value $lines[0] -Encoding UTF8 }
}

$script:BrandLogoDataUri = $null
function Get-Windows11BrandLogoDataUri {
    if ($null -ne $script:BrandLogoDataUri) { return $script:BrandLogoDataUri }
    $script:BrandLogoDataUri = ''
    $candidates = @(
        (Join-Path $PSScriptRoot 'WorkplaceCloudHub-lockup-WPF.png'),
        (Join-Path (Split-Path -Parent $PSScriptRoot) 'WorkplaceCloudHub-lockup-WPF.png')
    )
    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            try {
                $bytes = [System.IO.File]::ReadAllBytes($candidate)
                $script:BrandLogoDataUri = "data:image/png;base64," + [System.Convert]::ToBase64String($bytes)
                break
            }
            catch { }
        }
    }
    return $script:BrandLogoDataUri
}

function New-Windows11UpgradeCycleHtmlReport {
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Summary,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][int]$CycleNumber,
        [Parameter(Mandatory = $true)][datetime]$GeneratedAt,
        [switch]$IsLive,
        [AllowEmptyCollection()][object[]]$CycleProgress = @(),
        [AllowEmptyCollection()][object[]]$RunningJobRows = @()
    )

    $rows = @(Get-Windows11ReportRows -Items @($Summary | ForEach-Object { $_ }))

    $separatedDetailStatuses = @('ADMIN_SHARE_UNREACHABLE', 'RUN_GUARD_ACTIVE', 'SKIPPED_BY_TECH_RUN_GUARD')
    $mainRows = @($rows | Where-Object { $separatedDetailStatuses -notcontains (Get-Windows11HtmlEffectiveStatus -Row $_) })
    $separatedDetailRows = @($rows | Where-Object { $separatedDetailStatuses -contains (Get-Windows11HtmlEffectiveStatus -Row $_) })

    $effectiveRows = foreach ($row in $rows) {
        $eff = Get-Windows11HtmlEffectiveStatus -Row $row
        [pscustomobject]@{ EffectiveStatus = $eff; NextAction = [string]$row.RemoteNextAction }
    }
    $statusCounts = @($effectiveRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.EffectiveStatus) } | Group-Object -Property EffectiveStatus | Sort-Object Count -Descending | ForEach-Object {
        [pscustomobject]@{ Status = $_.Name; Count = $_.Count }
    })
    $nextActionCounts = @($effectiveRows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.NextAction) } | Group-Object -Property NextAction | Sort-Object Count -Descending | ForEach-Object {
        [pscustomobject]@{ NextAction = $_.Name; Count = $_.Count }
    })
    $adminShareFailureCounts = @($rows | Where-Object { (Get-Windows11HtmlEffectiveStatus -Row $_) -eq 'ADMIN_SHARE_UNREACHABLE' } | ForEach-Object { [pscustomobject]@{ FailureType = (Get-Windows11AdminShareFailureType -Row $_) } } | Group-Object -Property FailureType | Sort-Object Count -Descending | ForEach-Object {
        [pscustomobject]@{ FailureType = $_.Name; Count = $_.Count }
    })

    $logoUri = Get-Windows11BrandLogoDataUri
    $logoHtml = if (-not [string]::IsNullOrWhiteSpace($logoUri)) { "<img class='logo' src='$logoUri' alt='WorkplaceCloudHub' />" } else { "" }
    $mode = if ($IsLive) { 'LIVE' } else { 'FINAL' }
    $lotName = ''
    $lotPath = [string]$script:LotRoot
    if (-not [string]::IsNullOrWhiteSpace($lotPath)) {
        try { $lotName = Split-Path -Leaf $lotPath } catch { $lotName = '' }
    }
    if ([string]::IsNullOrWhiteSpace($lotName)) { $lotName = 'Unknown LOT' }
    $lotNameHtml = ConvertTo-HtmlText $lotName
    $lotPathHtml = ConvertTo-HtmlText $lotPath

    $style = @"
<style>
body { font-family: 'Segoe UI', Arial, sans-serif; margin: 0; padding: 24px; background: #F5F8FB; color: #1F2937; }
.header-card { display: flex; align-items: center; justify-content: space-between; background: #fff; border: 1px solid #DDE7F0; border-radius: 8px; padding: 18px 22px; margin-bottom: 18px; }
.header-text .title { font-size: 22px; font-weight: 600; color: #1F2937; }
.header-text .subtitle { font-size: 13px; color: #5F6B7A; margin-top: 2px; }
.header-text .lot-name { font-size: 18px; font-weight: 700; color: #005A9E; margin-top: 8px; }
.header-text .meta { font-size: 12px; color: #5F6B7A; margin-top: 10px; }
.badge { display: inline-block; background: #E6F4FF; color: #005A9E; border: 1px solid #B9DDF7; border-radius: 10px; padding: 1px 10px; font-size: 11px; font-weight: 600; margin-left: 8px; vertical-align: middle; }
.logo { height: 46px; }
.card { background: #fff; border: 1px solid #DDE7F0; border-radius: 8px; padding: 14px 18px; margin-bottom: 16px; }
h2 { font-size: 15px; margin: 0 0 8px 0; color: #1F2937; }
table { border-collapse: collapse; width: 100%; font-size: 12px; }
th { background: #0078D4; color: #fff; text-align: left; font-weight: 600; }
th, td { border: 1px solid #DDE7F0; padding: 6px 8px; vertical-align: top; }
tr:nth-child(even) td { background: #F5F8FB; }
.empty { color: #5F6B7A; font-style: italic; }
.footer { font-size: 11px; color: #5F6B7A; margin-top: 8px; }
.footer a { color: #0078D4; text-decoration: none; }
</style>
"@

    $html = New-Object System.Collections.Generic.List[string]
    [void]$html.Add(("<html><head><meta charset='utf-8'><title>Windows 11 upgrade - {0} - cycle {1}</title>{2}</head><body>" -f $lotNameHtml,$CycleNumber,$style))
    [void]$html.Add("<div class='header-card'>")
    [void]$html.Add("<div class='header-text'>")
    [void]$html.Add(("<div class='title'>Windows 11 upgrade - cycle {0}<span class='badge'>{1}</span></div>" -f $CycleNumber,$mode))
    [void]$html.Add("<div class='subtitle'>Smart Intune Windows 11 Upgrade Toolkit</div>")
    [void]$html.Add(("<div class='lot-name' title='{1}'>LOT: {0}</div>" -f $lotNameHtml,$lotPathHtml))
    [void]$html.Add(("<div class='meta'>Generated: {0} | Report rows: {1} | Launcher: v{2}</div>" -f (ConvertTo-HtmlText $GeneratedAt.ToString('yyyy-MM-dd HH:mm:ss')),$rows.Count,(ConvertTo-HtmlText $script:LauncherVersion)))
    [void]$html.Add(("<div class='meta'>Launcher log: {0}</div>" -f (New-HtmlLogLink -Path $script:LauncherLogPath)))
    [void]$html.Add("</div>")
    [void]$html.Add($logoHtml)
    [void]$html.Add("</div>")
    $optionRows = @($script:LauncherOptionRows | ForEach-Object { $_ })
    [void]$html.Add("<div class='card'><h2>LOT/run options</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $optionRows -Columns @("Category", "Option", "Value")))
    [void]$html.Add("</div>")
    [void]$html.Add("<div class='card'><h2>Cycle progress</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows @($CycleProgress) -Columns @('Cycle','Started','ElapsedMinutes','ComputerListLines','TotalUnique','Queued','CompletedRows','Running','Remaining','DuplicateGroups','DuplicateLines','DuplicateSamples')))
    [void]$html.Add("</div>")
    if ($RunningJobRows -and $RunningJobRows.Count -gt 0) {
        [void]$html.Add("<div class='card'><h2>Running jobs</h2>")
        [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows @($RunningJobRows) -Columns @('ComputerName','JobId','State','Started','ElapsedMinutes')))
        [void]$html.Add("</div>")
    }
    [void]$html.Add("<div class='card'><h2>Status summary</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $statusCounts -Columns @("Status", "Count")))
    [void]$html.Add("</div>")
    if ($adminShareFailureCounts.Count -gt 0) {
        [void]$html.Add("<div class='card'><h2>Admin share failure summary</h2>")
        [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $adminShareFailureCounts -Columns @('FailureType','Count')))
        [void]$html.Add("</div>")
    }
    [void]$html.Add("<div class='card'><h2>Next action summary</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $nextActionCounts -Columns @("NextAction", "Count")))
    [void]$html.Add("</div>")
    [void]$html.Add("<div class='card'><h2>Computer details</h2>")
    $htmlReportColumns = New-Object System.Collections.Generic.List[string]
    foreach ($column in $reportColumns) {
        [void]$htmlReportColumns.Add($column)
        if ($column -eq 'Detail') {
            [void]$htmlReportColumns.Add('Local log')
            [void]$htmlReportColumns.Add('Collected logs')
            [void]$htmlReportColumns.Add('Remote PC logs')
        }
    }
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $mainRows -Columns @($htmlReportColumns)))
    [void]$html.Add("<div class='footer'>Smart Intune Windows 11 Upgrade Toolkit - <a href='https://workplacecloudhub.com'>workplacecloudhub.com</a></div>")
    [void]$html.Add("</div>")
    [void]$html.Add("<div class='card'><h2>Run guard / admin share details</h2>")
    [void]$html.Add((ConvertTo-SimpleHtmlTable -Rows $separatedDetailRows -Columns @($htmlReportColumns)))
    [void]$html.Add("<div class='footer'>Rows excluded from Computer details: ADMIN_SHARE_UNREACHABLE, RUN_GUARD_ACTIVE, and SKIPPED_BY_TECH_RUN_GUARD.</div>")
    [void]$html.Add("</div>")
    [void]$html.Add("</body></html>")

    ($html -join "`r`n") | Out-File -LiteralPath $Path -Encoding UTF8 -Force
}

$globalGateRoot = Join-Path ([System.IO.Path]::GetTempPath()) 'SmartM365\GlobalWorkerGates'
$globalGateName = ($GlobalConcurrencySemaphoreName -replace '[^A-Za-z0-9_.-]', '_')
if ($globalGateName.Length -gt 120) { $globalGateName = $globalGateName.Substring(0, 120) }
$globalGatePath = Join-Path $globalGateRoot $globalGateName
$globalGateMutexName = "Local\SmartM365_GlobalWorkerGate_$globalGateName"
$launcherInstanceId = [guid]::NewGuid().ToString('N')
if ($GlobalConcurrencyLimit -gt 0) {
    New-Directory -Path $globalGatePath
    Write-Host ("Global lease gate: Limit={0}; Path={1}; LeaseTimeout={2} minute(s)" -f $GlobalConcurrencyLimit,$globalGatePath,$GlobalConcurrencyLeaseTimeoutMinutes) -ForegroundColor Green
}
$script:globalGateMutex = $null
if ($GlobalConcurrencyLimit -gt 0) {
    $script:globalGateMutex = New-Object System.Threading.Mutex($false, $globalGateMutexName)
}

function Invoke-WithGlobalGateMutex {
    param(
        [Parameter(Mandatory = $true)][scriptblock]$ScriptBlock,
        [object[]]$ArgumentList = @()
    )
    $sharedMutex = $script:globalGateMutex
    $mutex = if ($null -ne $sharedMutex) { $sharedMutex } else { New-Object System.Threading.Mutex($false, $globalGateMutexName) }
    $ownMutex = ($null -eq $sharedMutex)
    $acquired = $false
    try {
        try { $acquired = $mutex.WaitOne(30000) }
        catch [System.Threading.AbandonedMutexException] { $acquired = $true }
        if (-not $acquired) { throw "Could not acquire global gate mutex within 30 seconds: $globalGateMutexName" }
        & $ScriptBlock @ArgumentList
    }
    finally {
        if ($acquired) { try { $mutex.ReleaseMutex() } catch { } }
        if ($ownMutex) { $mutex.Dispose() }
    }
}

function Test-GateProcessAlive {
    param([Parameter(Mandatory = $true)][int]$ProcessId)
    try { $null = Get-Process -Id $ProcessId -ErrorAction Stop; return $true } catch { return $false }
}

function Remove-StaleGlobalLeases {
    $nowUtc = (Get-Date).ToUniversalTime()
    $removed = New-Object System.Collections.Generic.List[pscustomobject]
    foreach ($lease in @(Get-ChildItem -LiteralPath $globalGatePath -Filter '*.json' -File -ErrorAction SilentlyContinue)) {
        $remove = $false
        $reason = ''
        $computerName = ''
        $launcherPid = 0
        $workerPid = 0
        try {
            $data = Get-Content -LiteralPath $lease.FullName -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            $computerName = [string]$data.Computer
            $launcherPid = if ($data.PSObject.Properties['LauncherProcessId']) { [int]$data.LauncherProcessId } else { [int]$data.ProcessId }
            $workerPid = if ($data.PSObject.Properties['WorkerProcessId']) { [int]$data.WorkerProcessId } else { 0 }
            $createdUtc = [datetime]::Parse([string]$data.CreatedUtc, $null, [System.Globalization.DateTimeStyles]::RoundtripKind)
            if (-not (Test-GateProcessAlive -ProcessId $launcherPid)) {
                $remove = $true; $reason = 'LauncherProcessExited'
            } elseif ($workerPid -gt 0 -and -not (Test-GateProcessAlive -ProcessId $workerPid)) {
                $remove = $true; $reason = 'WorkerProcessExited'
            } elseif (($nowUtc - $createdUtc).TotalMinutes -gt $GlobalConcurrencyLeaseTimeoutMinutes) {
                $remove = $true; $reason = 'LeaseExpired'
            }
        } catch {
            $remove = $true; $reason = 'InvalidLease'
        }
        if ($remove) {
            $removed.Add([pscustomobject]@{ Reason = $reason; Computer = $computerName; LauncherPid = $launcherPid })
            Remove-Item -LiteralPath $lease.FullName -Force -ErrorAction SilentlyContinue
        }
    }
    if ($removed.Count -gt 0) {
        $pidParts = @($removed | Group-Object LauncherPid | Sort-Object Name | ForEach-Object {
            $reasons = @($_.Group | Group-Object Reason | ForEach-Object { if ($_.Count -gt 1) { "$($_.Name) x$($_.Count)" } else { $_.Name } }) -join '+'
            "PID=$($_.Name)[$reasons]"
        })
        $shortNames = @($removed | ForEach-Object { ($_.Computer -split '\.')[0] } | Where-Object { $_ })
        $computerPart = if ($shortNames.Count -le 4) { ': ' + ($shortNames -join ', ') } else { '' }
        Write-Host ("Removed {0} stale global worker lease(s) [{1}]{2}" -f $removed.Count,($pidParts -join '; '),$computerPart) -ForegroundColor DarkYellow
    }
}

function Update-GlobalLease {
    param(
        [AllowNull()][string]$LeasePath,
        [hashtable]$Properties = @{}
    )
    if ([string]::IsNullOrWhiteSpace($LeasePath) -or -not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return }
    Invoke-WithGlobalGateMutex -ScriptBlock {
        if (-not (Test-Path -LiteralPath $LeasePath -PathType Leaf)) { return }
        $data = Get-Content -LiteralPath $LeasePath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        foreach ($key in @($Properties.Keys)) {
            $data | Add-Member -NotePropertyName $key -NotePropertyValue $Properties[$key] -Force
        }
        $data | Add-Member -NotePropertyName LastUpdatedUtc -NotePropertyValue ((Get-Date).ToUniversalTime().ToString('o')) -Force
        $data | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $LeasePath -Encoding UTF8 -Force
    }
}

function Acquire-GlobalLease {
    param(
        [Parameter(Mandatory = $true)][string]$Computer,
        [Parameter(Mandatory = $true)][int]$CycleNumber
    )
    if ($GlobalConcurrencyLimit -lt 1) { return '' }
    $waitStarted = Get-Date
    $lastWaitLog = $null
    $waitLogDelaySeconds = 300
    $waitLogIntervalSeconds = 300
    $waitWasLogged = $false
    while ($true) {
        $leasePath = Invoke-WithGlobalGateMutex -ScriptBlock {
            Remove-StaleGlobalLeases
            $leases = @(Get-ChildItem -LiteralPath $globalGatePath -Filter '*.json' -File -ErrorAction SilentlyContinue)
            if ($leases.Count -lt $GlobalConcurrencyLimit) {
                $path = Join-Path $globalGatePath ("lease_{0}_{1}.json" -f $PID,([guid]::NewGuid().ToString('N')))
                [pscustomobject]@{
                    LauncherInstanceId = $launcherInstanceId
                    ProcessId = $PID
                    LauncherProcessId = $PID
                    Computer = $Computer
                    Cycle = $CycleNumber
                    JobId = ''
                    JobName = ''
                    WorkerProcessId = 0
                    WorkerProcessName = ''
                    WorkerStartedUtc = ''
                    CreatedUtc = (Get-Date).ToUniversalTime().ToString('o')
                    LastUpdatedUtc = (Get-Date).ToUniversalTime().ToString('o')
                    Host = $env:COMPUTERNAME
                } | ConvertTo-Json -Depth 3 | Set-Content -LiteralPath $path -Encoding UTF8 -Force
                return $path
            }
            return ''
        }
        if (-not [string]::IsNullOrWhiteSpace($leasePath)) {
            if ($waitWasLogged) {
                $waitMinutes = [math]::Round(((Get-Date) - $waitStarted).TotalMinutes, 1)
                Write-Host ("Global worker lease acquired for {0} after {1} minute(s)." -f $Computer,$waitMinutes) -ForegroundColor DarkCyan
            }
            return $leasePath
        }
        $now = Get-Date
        $waitSeconds = ($now - $waitStarted).TotalSeconds
        if ($waitSeconds -ge $waitLogDelaySeconds -and ($null -eq $lastWaitLog -or (($now - $lastWaitLog).TotalSeconds -ge $waitLogIntervalSeconds))) {
            $activeLeases = @(Get-ChildItem -LiteralPath $globalGatePath -Filter '*.json' -File -ErrorAction SilentlyContinue).Count
            $waitMinutes = [math]::Round(($now - $waitStarted).TotalMinutes, 1)
            Write-Host ("Waiting for global worker lease: Computer={0}; Active={1}; Limit={2}; Wait={3} minute(s)." -f $Computer,$activeLeases,$GlobalConcurrencyLimit,$waitMinutes) -ForegroundColor DarkYellow
            $lastWaitLog = $now
            $waitWasLogged = $true
        }
        Start-Sleep -Seconds $JobPollSeconds
    }
}

function Test-SampleDnsResolution {
    param(
        [string[]]$ComputerNames,
        [int]$SampleSize = 5
    )

    $sample = @($ComputerNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First $SampleSize)
    $tested = 0
    $failed = 0
    foreach ($computer in $sample) {
        $tested++
        try {
            $addresses = [System.Net.Dns]::GetHostAddresses($computer)
            if ($null -eq $addresses -or $addresses.Count -lt 1) { $failed++ }
        }
        catch {
            $failed++
        }
    }

    [pscustomobject]@{
        Tested = $tested
        Failed = $failed
        AllFailed = ($tested -gt 0 -and $failed -eq $tested)
    }
}

function Release-GlobalLease {
    param([AllowNull()][string]$LeasePath)
    if (-not [string]::IsNullOrWhiteSpace($LeasePath)) {
        Remove-Item -LiteralPath $LeasePath -Force -ErrorAction SilentlyContinue
    }
}

$script:IntuneInventoryMap = @{}
if (-not [string]::IsNullOrWhiteSpace($IntuneInventoryCsv)) {
    try {
        $refreshInitialIntuneInventory = $false
        $initialIntuneInventoryReason = ''
        $intuneInventoryItem = Get-Item -LiteralPath $IntuneInventoryCsv -ErrorAction SilentlyContinue
        if (-not $intuneInventoryItem) {
            $refreshInitialIntuneInventory = $true
            $initialIntuneInventoryReason = 'missing'
        }
        elseif (-not $IntuneInventoryUsesRecentRootCsv) {
            $intuneInventoryAge = (Get-Date) - $intuneInventoryItem.LastWriteTime
            if ($intuneInventoryAge.TotalHours -gt $script:IntuneInventoryFreshnessHours) {
                $refreshInitialIntuneInventory = $true
                $initialIntuneInventoryReason = ('older than {0} hour(s); LastWriteTime={1}; Age={2:N1} hour(s)' -f $script:IntuneInventoryFreshnessHours,$intuneInventoryItem.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'),$intuneInventoryAge.TotalHours)
            }
        }

        if ($IntuneInventoryUsesRecentRootCsv) {
            Write-Host ("Intune inventory CSV is recent. Using root CSV in priority: {0}" -f $IntuneInventoryCsv) -ForegroundColor Green
            $script:IntuneInventoryMap = Get-IntuneInventoryMap -Path $IntuneInventoryCsv -NameColumn $IntuneInventoryNameColumn
        }
        elseif ($refreshInitialIntuneInventory -and $DryRun) {
            Write-Host ("DryRun: Intune inventory CSV is {0}; skipping automatic Intune managed device export." -f $initialIntuneInventoryReason) -ForegroundColor Yellow
            if (Test-Path -LiteralPath $IntuneInventoryCsv -PathType Leaf) {
                $script:IntuneInventoryMap = Get-IntuneInventoryMap -Path $IntuneInventoryCsv -NameColumn $IntuneInventoryNameColumn
            }
        }
        elseif ($refreshInitialIntuneInventory -and -not $SkipIntuneInventoryRefresh) {
            $initialIntuneInventoryLogPath = Join-Path $ReportRoot ("DevicesIntune_InitialRefresh_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
            Write-Host ("Intune inventory CSV is {0}. Running Intune managed device export before starting the lot..." -f $initialIntuneInventoryReason) -ForegroundColor Yellow
            $initialIntuneInventory = Invoke-FullIntuneInventoryExport `
                -ExportScriptPath $script:ExportIntuneScriptPath `
                -OutputPath $IntuneInventoryCsv `
                -LogPath $initialIntuneInventoryLogPath `
                -ComputerListPath $ComputerListPath `
                -PageSize $IntuneInventoryPageSize `
                -TenantId $IntuneTenantId
            if ($initialIntuneInventory.Success) {
                $script:IntuneInventoryMap = $initialIntuneInventory.InventoryMap
                Write-Host ("Initial Intune inventory refreshed. Devices={0}; CSV={1}" -f $script:IntuneInventoryMap.Count,$initialIntuneInventory.CsvPath) -ForegroundColor Green
            }
            else {
                Write-Host ("WARNING: Initial Intune inventory refresh failed: {0}" -f $initialIntuneInventory.Error) -ForegroundColor Yellow
                if (Test-Path -LiteralPath $IntuneInventoryCsv -PathType Leaf) {
                    Write-Host ("Using existing Intune inventory CSV despite refresh failure: {0}" -f $IntuneInventoryCsv) -ForegroundColor Yellow
                    $script:IntuneInventoryMap = Get-IntuneInventoryMap -Path $IntuneInventoryCsv -NameColumn $IntuneInventoryNameColumn
                }
            }
        }
        elseif (Test-Path -LiteralPath $IntuneInventoryCsv -PathType Leaf) {
            $script:IntuneInventoryMap = Get-IntuneInventoryMap -Path $IntuneInventoryCsv -NameColumn $IntuneInventoryNameColumn
        }

        if ($script:IntuneInventoryMap.Count -gt 0) {
            $currentComputers = @(Get-ComputerList -Path $ComputerListPath)
            $alreadyWindows11FromIntune = @(Get-AlreadyWindows11RowsFromIntuneInventory -ComputerNames $currentComputers -IntuneInventoryMap $script:IntuneInventoryMap -IntuneInventoryCsv $IntuneInventoryCsv)
            if ($alreadyWindows11FromIntune.Count -gt 0) {
                $intuneAlreadyPath = Join-Path $ReportRoot ("DevicesIntune_AlreadyWindows11_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
                @($alreadyWindows11FromIntune) | Export-Csv -LiteralPath $intuneAlreadyPath -NoTypeInformation -Encoding UTF8
                if ($DryRun) {
                    Write-Host ("DryRun: Intune inventory detected {0} already-Windows11 computer(s). No Computers.txt change. CSV={1}" -f $alreadyWindows11FromIntune.Count,$intuneAlreadyPath) -ForegroundColor Yellow
                }
                else {
                    $preMoveResult = Move-AlreadyWindows11ComputersFromList -ComputerListPath $ComputerListPath -CycleSummary $alreadyWindows11FromIntune
                    if ($preMoveResult.Moved -gt 0) {
                        Write-Host ("Intune inventory moved {0} already-Windows11 computer(s) from Computers.txt to {1}. Evidence={2}" -f $preMoveResult.Moved,$preMoveResult.AlreadyWindows11Path,$intuneAlreadyPath) -ForegroundColor Green
                    }
                    else {
                        Write-Host ("Intune inventory detected already-Windows11 computer(s), but none were still present in Computers.txt. Evidence={0}" -f $intuneAlreadyPath) -ForegroundColor DarkGray
                    }
                }
            }
            else {
                Write-Host ("Intune inventory precheck found no Windows 11 computer in current Computers.txt. CSV={0}" -f $IntuneInventoryCsv) -ForegroundColor DarkGray
            }
        }
        else {
            Write-Host ("Intune inventory precheck skipped: no Intune rows loaded. CSV={0}" -f $IntuneInventoryCsv) -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host ("WARN: Intune inventory precheck failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
}

$script:AdInventoryMap = @{}
if (-not [string]::IsNullOrWhiteSpace($AdInventoryCsv)) {
    try {
        $refreshInitialAdInventory = $false
        $initialAdInventoryReason = ''
        $adInventoryItem = Get-Item -LiteralPath $AdInventoryCsv -ErrorAction SilentlyContinue
        if (-not $adInventoryItem) {
            $refreshInitialAdInventory = $true
            $initialAdInventoryReason = 'missing'
        }
        elseif (-not $AdInventoryUsesRecentRootCsv) {
            $adInventoryAge = (Get-Date) - $adInventoryItem.LastWriteTime
            if ($adInventoryAge.TotalHours -gt $script:AdInventoryFreshnessHours) {
                $refreshInitialAdInventory = $true
                $initialAdInventoryReason = ('older than {0} hour(s); LastWriteTime={1}; Age={2:N1} hour(s)' -f $script:AdInventoryFreshnessHours,$adInventoryItem.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss'),$adInventoryAge.TotalHours)
            }
        }

        if ($AdInventoryUsesRecentRootCsv) {
            Write-Host ("AD forest inventory CSV is recent. Using root CSV in priority: {0}" -f $AdInventoryCsv) -ForegroundColor Green
            $script:AdInventoryMap = Get-AdInventoryMap -Path $AdInventoryCsv -NameColumn $AdInventoryNameColumn
        }
        elseif ($refreshInitialAdInventory -and $DryRun) {
            Write-Host ("DryRun: AD inventory CSV is {0}; skipping automatic AD computer export." -f $initialAdInventoryReason) -ForegroundColor Yellow
            if (Test-Path -LiteralPath $AdInventoryCsv -PathType Leaf) {
                $script:AdInventoryMap = Get-AdInventoryMap -Path $AdInventoryCsv -NameColumn $AdInventoryNameColumn
            }
        }
        elseif ($refreshInitialAdInventory -and -not $SkipAdInventoryRefresh) {
            $initialAdInventoryLogPath = Join-Path $ReportRoot ("DevicesAD_InitialRefresh_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
            $initialAdScope = if ([string]::IsNullOrWhiteSpace($AdDomain)) { 'Current AD forest, limited to Computers.txt' } else { "Domain=$AdDomain, limited to Computers.txt" }
            Write-Host ("AD inventory CSV is {0}. Running AD computer export before starting the lot. Scope={1}..." -f $initialAdInventoryReason,$initialAdScope) -ForegroundColor Yellow
            $initialAdInventory = Invoke-FullAdInventoryExport `
                -ExportScriptPath $script:ExportAdScriptPath `
                -OutputPath $AdInventoryCsv `
                -LogPath $initialAdInventoryLogPath `
                -ComputerListPath $ComputerListPath `
                -Domain $AdDomain
            if ($initialAdInventory.Success) {
                $script:AdInventoryMap = $initialAdInventory.InventoryMap
                Write-Host ("Initial AD inventory refreshed. Devices={0}; CSV={1}" -f $script:AdInventoryMap.Count,$initialAdInventory.CsvPath) -ForegroundColor Green
            }
            else {
                Write-Host ("WARNING: Initial AD inventory refresh failed: {0}" -f $initialAdInventory.Error) -ForegroundColor Yellow
                if (Test-Path -LiteralPath $AdInventoryCsv -PathType Leaf) {
                    Write-Host ("Using existing AD inventory CSV despite refresh failure: {0}" -f $AdInventoryCsv) -ForegroundColor Yellow
                    $script:AdInventoryMap = Get-AdInventoryMap -Path $AdInventoryCsv -NameColumn $AdInventoryNameColumn
                }
            }
        }
        elseif (Test-Path -LiteralPath $AdInventoryCsv -PathType Leaf) {
            $script:AdInventoryMap = Get-AdInventoryMap -Path $AdInventoryCsv -NameColumn $AdInventoryNameColumn
        }

        if ($script:AdInventoryMap.Count -gt 0) {
            $currentComputers = @(Get-ComputerList -Path $ComputerListPath)
            $alreadyWindows11FromAd = @(Get-AlreadyWindows11RowsFromAdInventory -ComputerNames $currentComputers -AdInventoryMap $script:AdInventoryMap -AdInventoryCsv $AdInventoryCsv)
            if ($alreadyWindows11FromAd.Count -gt 0) {
                $adAlreadyPath = Join-Path $ReportRoot ("DevicesAD_AlreadyWindows11_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
                @($alreadyWindows11FromAd) | Export-Csv -LiteralPath $adAlreadyPath -NoTypeInformation -Encoding UTF8
                if ($DryRun) {
                    Write-Host ("DryRun: AD inventory detected {0} already-Windows11 computer(s). No Computers.txt change. CSV={1}" -f $alreadyWindows11FromAd.Count,$adAlreadyPath) -ForegroundColor Yellow
                }
                else {
                    $preMoveResult = Move-AlreadyWindows11ComputersFromList -ComputerListPath $ComputerListPath -CycleSummary $alreadyWindows11FromAd
                    if ($preMoveResult.Moved -gt 0) {
                        Write-Host ("AD inventory moved {0} already-Windows11 computer(s) from Computers.txt to {1}. Evidence={2}" -f $preMoveResult.Moved,$preMoveResult.AlreadyWindows11Path,$adAlreadyPath) -ForegroundColor Green
                    }
                    else {
                        Write-Host ("AD inventory detected already-Windows11 computer(s), but none were still present in Computers.txt. Evidence={0}" -f $adAlreadyPath) -ForegroundColor DarkGray
                    }
                }
            }
            else {
                Write-Host ("AD inventory precheck found no Windows 11 computer in current Computers.txt. CSV={0}" -f $AdInventoryCsv) -ForegroundColor DarkGray
            }
        }
        else {
            Write-Host ("AD inventory precheck skipped: no AD rows loaded. CSV={0}" -f $AdInventoryCsv) -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host ("WARN: AD inventory precheck failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
}
$cycle = 0
do {
    $cycle++
    $computers = @(Get-ComputerList -Path $ComputerListPath)
    $computerListStats = Get-ComputerListStats -Path $ComputerListPath
    if ($computerListStats.DuplicateGroups -gt 0) {
        Write-Host ("Cycle {0}: Computers.txt contains {1} duplicate line(s) in {2} duplicate group(s). Duplicates ignored: {3}" -f $cycle,$computerListStats.DuplicateLines,$computerListStats.DuplicateGroups,$computerListStats.DuplicateSamples) -ForegroundColor Yellow
    }
    if ($computers.Count -eq 0) {
        Write-Host "No computers found in $ComputerListPath." -ForegroundColor Yellow
        break
    }

    $dnsCheck = Test-SampleDnsResolution -ComputerNames $computers
    if ($dnsCheck.AllFailed) {
        Write-Host ""
        Write-Host ("*** DNS WARNING: resolution failed on all {0} sampled computers. Check VPN connectivity and DNS configuration on this machine before continuing. ***" -f $dnsCheck.Tested) -ForegroundColor Red
        Write-Host ""
    }

    Write-Host ("Cycle {0}: {1} computer(s)." -f $cycle,$computers.Count) -ForegroundColor Cyan
    $results = New-Object System.Collections.ArrayList
    $runningJobs = @()
    $globalLeaseByJobId = @{}
    $techRunGuardFqdnByJobId = @{}
    $jobStartedAtById = @{}
    $nextIndex = 0

    $liveHtmlPath = Join-Path $ReportRoot ("PsExec_Windows11Upgrade_Summary_{0}_cycle{1}_live.html" -f $script:LauncherLogSafeLotName,$cycle)
    $liveCsvPath = Join-Path $ReportRoot ("PsExec_Windows11Upgrade_Summary_{0}_cycle{1}_live.csv" -f $script:LauncherLogSafeLotName,$cycle)
    $lastLiveHtmlWrite = [datetime]::MinValue
    $cycleStart = Get-Date
    $lastProgressLog = Get-Date
    try {
        $cycleProgress = New-Windows11CycleProgressRows -CycleNumber $cycle -CycleStart $cycleStart -TotalComputers $computers.Count -QueuedComputers $nextIndex -CompletedComputers $results.Count -RunningComputers $runningJobs.Count -ComputerListStats $computerListStats
        Export-Windows11ReportCsv -Rows @() -Path $liveCsvPath
        New-Windows11UpgradeCycleHtmlReport -Summary @() -Path $liveHtmlPath -CycleNumber $cycle -GeneratedAt (Get-Date) -IsLive -CycleProgress $cycleProgress -RunningJobRows @()
    }
    catch { Write-Host ("Cycle {0}: failed to initialize live report: {1}" -f $cycle,$_.Exception.Message) -ForegroundColor Yellow }
    while ($nextIndex -lt $computers.Count -or $runningJobs.Count -gt 0) {
        while ($nextIndex -lt $computers.Count -and $runningJobs.Count -lt $ThrottleLimit) {
            $computer = $computers[$nextIndex]
            $nextIndex++

            $techRunGuardFqdn = Get-TechnicianRunGuardFqdn -ComputerName $computer -AdInventoryMap $script:AdInventoryMap
            if ($script:UseEffectiveTechnicianRunGuardHistory) {
                $activeTechRunGuard = Get-ActiveTechnicianRunGuardEntry -Path $script:TechnicianRunGuardHistoryPath -ComputerFqdn $techRunGuardFqdn -RunGuardHours $RunGuardHours
                if ($null -ne $activeTechRunGuard) {
                    $skipResult = New-TechnicianRunGuardSkippedResult -ComputerName $computer -CycleNumber $cycle -HistoryEntry $activeTechRunGuard -RunGuardHours $RunGuardHours
                    $skipResult = Add-AdInventoryFieldsToResult -Result $skipResult -AdInventoryMap $script:AdInventoryMap -AdInventoryCsv $AdInventoryCsv
                    $skipResult = Add-IntuneInventoryFieldsToResult -Result $skipResult -IntuneInventoryMap $script:IntuneInventoryMap -IntuneInventoryCsv $IntuneInventoryCsv
                    [void]$results.Add($skipResult)
                    Write-Host ("  [SKIPPED_BY_TECH_RUN_GUARD] {0}: {1}" -f $computer,$skipResult.Detail) -ForegroundColor DarkYellow
                    continue
                }
            }

            $globalLeasePath = Acquire-GlobalLease -Computer $computer -CycleNumber $cycle

            try {
                $remoteArgsJson = ([pscustomobject]@{ Args = @($remoteArgs.ToArray()) } | ConvertTo-Json -Compress)
                $workerArgs = @(
                    $computer,
                    $cycle,
                    $remoteArgsJson,
                    $resolvedPsExec,
                    $LocalScriptPath,
                    $script:RemoteBaseDir,
                    $script:RemoteScriptPath,
                    $script:RemoteSetupCacheRoot,
                    $LogRoot,
                    $CentralLogRoot,
                    $SetupSourcePath,
                    $SetupSourceMapPath,
                    $SetupExecutionMode,
                    $SetupMediaId,
                    $SetupLanguage,
                    [bool]$AllowSetupUpgrade,
                    [bool]$SkipSetupMediaPreCopy,
                    [bool]$SkipVirtualMachines,
                    [bool]$DryRun,
                    [bool]$NoCentralLogCollection,
                    [bool]$KeepCentralLogHistory,
                    $CentralLogCollectionMode,
                    $PsExecTimeoutMinutes,
                    $globalLeasePath,
                    $globalGateMutexName
                )

                $job = Start-Job -Name ("W11UT_C{0}_{1}" -f $cycle,$computer) -FilePath $LocalWorkerPath -ArgumentList $workerArgs
                $jobStartedAtById[[string]$job.Id] = Get-Date
                if ($script:UseEffectiveTechnicianRunGuardHistory) {
                    $techRunGuardFqdnByJobId[[string]$job.Id] = $techRunGuardFqdn
                    Update-TechnicianRunGuardHistory -Path $script:TechnicianRunGuardHistoryPath -ComputerFqdn $techRunGuardFqdn -InputComputerName $computer -RunGuardHours $RunGuardHours -State Started -Result $null -JobId ([string]$job.Id)
                }
            }
            catch {
                Release-GlobalLease -LeasePath $globalLeasePath
                throw
            }

            if (-not [string]::IsNullOrWhiteSpace($globalLeasePath)) {
                Update-GlobalLease -LeasePath $globalLeasePath -Properties @{
                    JobId = [string]$job.Id
                    JobName = [string]$job.Name
                }
                $globalLeaseByJobId[[string]$job.Id] = $globalLeasePath
            }
            $runningJobs += $job
            Write-Host ("Queued {0} ({1}/{2}); running={3}" -f $computer,$nextIndex,$computers.Count,$runningJobs.Count) -ForegroundColor DarkCyan

            if ($DelayBetweenComputersSeconds -gt 0 -and $nextIndex -lt $computers.Count) {
                Start-Sleep -Seconds $DelayBetweenComputersSeconds
            }
        }

        $finishedJobs = @($runningJobs | Where-Object { $_.State -ne 'Running' })
        if ($finishedJobs.Count -eq 0) {
            if (((Get-Date) - $lastProgressLog).TotalSeconds -ge 300) {
                $now = Get-Date
                $waitingNames = @($runningJobs | ForEach-Object {
                    $jobId = [string]$_.Id
                    $jobName = $_.Name -replace '^W11UT_C\d+_',''
                    if ($jobStartedAtById.ContainsKey($jobId)) {
                        "{0} ({1}m)" -f $jobName,[math]::Round(($now - $jobStartedAtById[$jobId]).TotalMinutes,1)
                    }
                    else {
                        $jobName
                    }
                })
                Write-Host ("Waiting for {0} job(s); Elapsed={1} min; Running: {2}" -f $runningJobs.Count,[math]::Round(($now - $cycleStart).TotalMinutes,1),($waitingNames -join ', '))
                $lastProgressLog = $now
            }
            if (((Get-Date) - $lastLiveHtmlWrite).TotalSeconds -ge 60) {
                try {
                    $liveRows = @(Get-Windows11ReportRows -Items @($results.ToArray()))
                    $cycleProgress = New-Windows11CycleProgressRows -CycleNumber $cycle -CycleStart $cycleStart -TotalComputers $computers.Count -QueuedComputers $nextIndex -CompletedComputers $results.Count -RunningComputers $runningJobs.Count -ComputerListStats $computerListStats
                    $runningJobRows = New-Windows11RunningJobRows -RunningJobs @($runningJobs) -JobStartedAtById $jobStartedAtById
                    Export-Windows11ReportCsv -Rows $liveRows -Path $liveCsvPath
                    New-Windows11UpgradeCycleHtmlReport -Summary $liveRows -Path $liveHtmlPath -CycleNumber $cycle -GeneratedAt (Get-Date) -IsLive -CycleProgress $cycleProgress -RunningJobRows $runningJobRows
                    $lastLiveHtmlWrite = Get-Date
                }
                catch { Write-Host ("Cycle {0}: failed to update live report: {1}" -f $cycle,$_.Exception.Message) -ForegroundColor Yellow }
            }
            Start-Sleep -Seconds $JobPollSeconds
            continue
        }

        foreach ($job in $finishedJobs) {
            $received = $null
            $jobErrors = @()

            try {
                $receiveErrors = @()
                $received = @(Receive-Job -Job $job -ErrorAction SilentlyContinue -ErrorVariable receiveErrors)
                $childJobs = @($job.ChildJobs)
                $brokenRunspace = (
                    $childJobs.Count -gt 0 -and
                    $null -ne $childJobs[0].JobStateInfo.Reason -and
                    [string]$childJobs[0].JobStateInfo.Reason.Message -match '\bBroken\b'
                )
                $jobErrors = @(
                    @(
                        $receiveErrors | ForEach-Object { $_.ToString() }
                        $childJobs | ForEach-Object { $_.Error } | ForEach-Object { $_.ToString() }
                        if ($childJobs.Count -gt 0 -and $childJobs[0].JobStateInfo.Reason) {
                            $childJobs[0].JobStateInfo.Reason.Message
                        }
                        if ($brokenRunspace) {
                            'RUNSPACE_BROKEN: PowerShell job runspace entered a Broken state. This typically indicates PowerShell 5.1 instability under parallel load. Consider adding -DelayBetweenComputersSeconds 1 to reduce job cycling speed.'
                        }
                    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
                )
                if ($brokenRunspace) {
                    Write-Host ("  [RUNSPACE_BROKEN] {0}: job runspace entered Broken state - possible PowerShell 5.1 instability. Consider -DelayBetweenComputersSeconds 1." -f ($job.Name -replace '^W11UT_C\d+_','')) -ForegroundColor Magenta
                }
            }
            catch {
                $jobErrors += $_.Exception.Message
            }

            if (-not $received -or $received.Count -eq 0) {
                $received = @([pscustomobject]@{
                    Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                    ComputerName = ($job.Name -replace '^W11UT_C\d+_','')
                    CycleNumber = $cycle
                    LauncherStatus = if (@($jobErrors | Where-Object { $_ -match 'RUNSPACE_BROKEN' }).Count -gt 0) { 'RUNSPACE_BROKEN' } else { 'JOB_ERROR' }
                    RemoteStatus = ''
                    RemoteNextAction = ''
                    ExitCode = ''
                    Detail = ($jobErrors -join ' | ')
                    SetupCacheAction = ''
                    DiskCleanupAction = ''
                    DiskCleanupFreedGB = ''
                    AdvancedDiskCleanupAction = ''
                    AdvancedDiskCleanupFreedGB = ''
                    DismCleanupAction = ''
                    DismCleanupFreedGB = ''
                    SetupCompletionRebootAction = ''
                    SetupCompletionRebootDetail = ''
                    SetupCompletionRebootUserCount = ''
                    SetupCompletionRebootUsers = ''
                    SetupProfileRepairAction = ''
                    SetupProfileRepairDetail = ''
                    SetupProfileRepairBlockingSid = ''
                    SetupProfileRepairKeptSid = ''
                    SetupProfileRepairProfilePath = ''
                    SetupProfileRepairBackupPath = ''
                    ControlledRebootAction = ''
                    ControlledRebootDetail = ''
                    ControlledRebootUserCount = ''
                    ControlledRebootUsers = ''
                    UserRebootNotificationSent = ''
                    UserRebootNotificationLang = ''
                    UserRebootNotificationMessage = ''
                    RemoteLogsPath = ''
                    PsExecLogPath = ''
                    JobErrorMessage = ($jobErrors -join ' | ')
                })
            }

            foreach ($item in @($received)) {
                if ($null -ne $item) {
                    if ($jobErrors.Count -gt 0 -and -not $item.PSObject.Properties['JobErrorMessage']) {
                        $item | Add-Member -NotePropertyName JobErrorMessage -NotePropertyValue ($jobErrors -join ' | ') -Force
                    }
                    $item = Add-AdInventoryFieldsToResult -Result $item -AdInventoryMap $script:AdInventoryMap -AdInventoryCsv $AdInventoryCsv
                    $item = Add-IntuneInventoryFieldsToResult -Result $item -IntuneInventoryMap $script:IntuneInventoryMap -IntuneInventoryCsv $IntuneInventoryCsv
                    $jobIdForLog = [string]$job.Id
                    $jobElapsedMinutes = if ($jobStartedAtById.ContainsKey($jobIdForLog)) { [math]::Round(((Get-Date) - $jobStartedAtById[$jobIdForLog]).TotalMinutes,1) } else { '' }
                    $computerForLog = if ($item.PSObject.Properties['ComputerName']) { [string]$item.ComputerName } else { ($job.Name -replace '^W11UT_C\d+_','') }
                    $launcherStatusForLog = if ($item.PSObject.Properties['LauncherStatus']) { [string]$item.LauncherStatus } else { '' }
                    $remoteStatusForLog = if ($item.PSObject.Properties['RemoteStatus']) { [string]$item.RemoteStatus } else { '' }
                    $statusForLog = if (-not [string]::IsNullOrWhiteSpace($remoteStatusForLog)) { $remoteStatusForLog } else { $launcherStatusForLog }
                    $nextActionForLog = if ($item.PSObject.Properties['RemoteNextAction']) { [string]$item.RemoteNextAction } else { '' }
                    $exitCodeForLog = if ($item.PSObject.Properties['ExitCode']) { [string]$item.ExitCode } else { '' }
                    $detailForLog = if ($item.PSObject.Properties['Detail']) { [string]$item.Detail } else { '' }
                    if ($detailForLog.Length -gt 180) { $detailForLog = $detailForLog.Substring(0,177) + '...' }
                    $completionColor = 'Red'
                    if ($exitCodeForLog -eq '0' -or $statusForLog -match 'ALREADY_WINDOWS11|SUCCESS|STARTED') { $completionColor = 'Green' }
                    elseif ($statusForLog -match 'RUN_GUARD|PENDING|INSUFFICIENT|TIMEOUT|SKIPPED|READY') { $completionColor = 'Yellow' }
                    Write-Host ("Completed {0}; Elapsed={1} min; Status={2}; NextAction={3}; ExitCode={4}; Detail={5}" -f $computerForLog,$jobElapsedMinutes,$statusForLog,$nextActionForLog,$exitCodeForLog,$detailForLog) -ForegroundColor $completionColor
                    if ($script:UseEffectiveTechnicianRunGuardHistory) {
                        $resultFqdn = if ($techRunGuardFqdnByJobId.ContainsKey([string]$job.Id)) { [string]$techRunGuardFqdnByJobId[[string]$job.Id] } else { Get-TechnicianRunGuardFqdn -ComputerName ([string]$item.ComputerName) -AdInventoryMap $script:AdInventoryMap }
                        Update-TechnicianRunGuardHistory -Path $script:TechnicianRunGuardHistoryPath -ComputerFqdn $resultFqdn -InputComputerName ([string]$item.ComputerName) -RunGuardHours $RunGuardHours -State Result -Result $item -JobId ([string]$job.Id)
                    }
                    [void]$results.Add($item)
                    if (-not $DryRun -and (Test-AlreadyWindows11CycleResult -Result $item)) {
                        try {
                            $moveSingleResult = Move-AlreadyWindows11ComputersFromList -ComputerListPath $ComputerListPath -CycleSummary @($item)
                            if ($moveSingleResult.Moved -gt 0) {
                                Write-Host ("Moved already-Windows11 computer from Computers.txt to {0}: {1}" -f $moveSingleResult.AlreadyWindows11Path,$item.ComputerName) -ForegroundColor Green
                            }
                        }
                        catch {
                            Write-Host ("Failed to update ComputersAlreadyW11.txt for {0}: {1}" -f $item.ComputerName,$_.Exception.Message) -ForegroundColor Yellow
                        }
                    }
                }
            }
            if ($globalLeaseByJobId.ContainsKey([string]$job.Id)) {
                Release-GlobalLease -LeasePath $globalLeaseByJobId[[string]$job.Id]
                $globalLeaseByJobId.Remove([string]$job.Id)
            }
            if ($techRunGuardFqdnByJobId.ContainsKey([string]$job.Id)) {
                $techRunGuardFqdnByJobId.Remove([string]$job.Id)
            }
            if ($jobStartedAtById.ContainsKey([string]$job.Id)) {
                $jobStartedAtById.Remove([string]$job.Id)
            }
            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }

        $currentRunningJobs = @($runningJobs | Where-Object { $_.State -eq 'Running' })
        if (((Get-Date) - $lastLiveHtmlWrite).TotalSeconds -ge 3) {
            try {
                $liveRows = @(Get-Windows11ReportRows -Items @($results.ToArray()))
                $cycleProgress = New-Windows11CycleProgressRows -CycleNumber $cycle -CycleStart $cycleStart -TotalComputers $computers.Count -QueuedComputers $nextIndex -CompletedComputers $results.Count -RunningComputers $currentRunningJobs.Count -ComputerListStats $computerListStats
                $runningJobRows = New-Windows11RunningJobRows -RunningJobs $currentRunningJobs -JobStartedAtById $jobStartedAtById
                Export-Windows11ReportCsv -Rows $liveRows -Path $liveCsvPath
                New-Windows11UpgradeCycleHtmlReport -Summary $liveRows -Path $liveHtmlPath -CycleNumber $cycle -GeneratedAt (Get-Date) -IsLive -CycleProgress $cycleProgress -RunningJobRows $runningJobRows
                $lastLiveHtmlWrite = Get-Date
            }
            catch { Write-Host ("Cycle {0}: failed to update live report: {1}" -f $cycle,$_.Exception.Message) -ForegroundColor Yellow }
        }

        $runningJobs = $currentRunningJobs
    }

    $reportTimestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $reportPath = Join-Path $ReportRoot ("PsExec_Windows11Upgrade_Summary_{0}_cycle{1}_{2}.csv" -f $script:LauncherLogSafeLotName,$cycle,$reportTimestamp)
    $enrichedResults = @($results.ToArray() | ForEach-Object { $row = Add-AdInventoryFieldsToResult -Result $_ -AdInventoryMap $script:AdInventoryMap -AdInventoryCsv $AdInventoryCsv; Add-IntuneInventoryFieldsToResult -Result $row -IntuneInventoryMap $script:IntuneInventoryMap -IntuneInventoryCsv $IntuneInventoryCsv })
    $normalizedResults = @(Get-Windows11ReportRows -Items @($enrichedResults))
    Export-Windows11ReportCsv -Rows @($normalizedResults) -Path $reportPath
    Write-Host ("Cycle {0} report: {1}" -f $cycle,$reportPath) -ForegroundColor Green

    $htmlReportPath = Join-Path $ReportRoot ("PsExec_Windows11Upgrade_Summary_{0}_cycle{1}_{2}.html" -f $script:LauncherLogSafeLotName,$cycle,$reportTimestamp)
    try {
        $finalProgress = New-Windows11CycleProgressRows -CycleNumber $cycle -CycleStart $cycleStart -TotalComputers $computers.Count -QueuedComputers $computers.Count -CompletedComputers $normalizedResults.Count -RunningComputers 0 -ComputerListStats $computerListStats
        New-Windows11UpgradeCycleHtmlReport -Summary @($normalizedResults) -Path $htmlReportPath -CycleNumber $cycle -GeneratedAt (Get-Date) -CycleProgress $finalProgress -RunningJobRows @()
        if (Test-Path -LiteralPath $liveHtmlPath -PathType Leaf) {
            Remove-Item -LiteralPath $liveHtmlPath -Force -ErrorAction Stop
            Write-Host ("Cycle {0}: removed live HTML report after final report creation: {1}" -f $cycle,$liveHtmlPath) -ForegroundColor DarkGray
        }
        if (Test-Path -LiteralPath $liveCsvPath -PathType Leaf) {
            Remove-Item -LiteralPath $liveCsvPath -Force -ErrorAction Stop
            Write-Host ("Cycle {0}: removed live CSV report after final report creation: {1}" -f $cycle,$liveCsvPath) -ForegroundColor DarkGray
        }
        Write-Host ("Cycle {0} HTML report: {1}" -f $cycle,$htmlReportPath) -ForegroundColor Green
    }
    catch { Write-Host ("Cycle {0}: failed to write HTML report: {1}" -f $cycle,$_.Exception.Message) -ForegroundColor Yellow }

    try {
        $moveAlreadyW11Result = Move-AlreadyWindows11ComputersFromList -ComputerListPath $ComputerListPath -CycleSummary @($normalizedResults)
        if ($moveAlreadyW11Result.Moved -gt 0) {
            Write-Host ("Cycle {0}: moved {1} already-Windows11 computer(s) to {2}" -f $cycle,$moveAlreadyW11Result.Moved,$moveAlreadyW11Result.AlreadyWindows11Path) -ForegroundColor Green
        }
    }
    catch {
        Write-Host ("Cycle {0}: failed to update ComputersAlreadyW11.txt: {1}" -f $cycle,$_.Exception.Message) -ForegroundColor Yellow
    }

    $sourceDistribution = @(
        $normalizedResults |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.SelectedSetupSourcePath) } |
            Group-Object -Property SelectedSetupSourcePath |
            Sort-Object Count -Descending |
            ForEach-Object {
                [pscustomobject]@{
                    SelectedSetupSourcePath = [string]$_.Name
                    ComputerCount = [int]$_.Count
                }
            }
    )
    if ($sourceDistribution.Count -gt 0) {
        $distributionPath = Join-Path $ReportRoot ("SetupSource_Distribution_{0}_cycle{1}_{2}.csv" -f $script:LauncherLogSafeLotName,$cycle,(Get-Date -Format 'yyyyMMdd-HHmmss'))
        $sourceDistribution | Export-Csv -LiteralPath $distributionPath -NoTypeInformation -Encoding UTF8
        Write-Host ("Setup source distribution: {0}" -f $distributionPath) -ForegroundColor Green
        foreach ($sourceGroup in $sourceDistribution) {
            Write-Host ("  {0}: {1}" -f $sourceGroup.SelectedSetupSourcePath,$sourceGroup.ComputerCount) -ForegroundColor DarkGreen
        }
    }

    if ($RunOnce) { break }
    if ($MaxCycles -gt 0 -and $cycle -ge $MaxCycles) { break }
    if ($DelayBetweenCyclesMinutes -gt 0) {
        Write-Host ("Waiting {0} minute(s) before next cycle." -f $DelayBetweenCyclesMinutes) -ForegroundColor DarkYellow
        Start-Sleep -Seconds ($DelayBetweenCyclesMinutes * 60)
    }
}
while ($true)

Write-Host ("SmartM365 Windows 11 Upgrade Toolkit launcher finished. Log={0}" -f $script:LauncherLogPath)

if ($null -ne $script:globalGateMutex) {
    try { $script:globalGateMutex.Dispose() } catch { }
    $script:globalGateMutex = $null
}
