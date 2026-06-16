<#
.SYNOPSIS
Runs SmartM365-Invoke-IntuneHybridJoinRepair.ps1 on a list of remote computers using PsExec.

.DESCRIPTION
Reads one computer name per line from Computers.txt by default, copies the repair script to
C:\ProgramData\SmartM365\IntuneHybridJoinToolkit on each target through C$, then starts it as
LocalSystem with PsExec.

Lines starting with # in the computer list are ignored.
The launcher loops by default. It reloads Computers.txt at the beginning of each cycle.
After each cycle, computers detected as already Intune-enrolled are removed from Computers.txt
and appended to ComputersAlreadyEnrolled.txt in the same folder.

.PARAMETER ComputerListPath
Path to a text file containing one computer name per line.

.PARAMETER PsExecPath
Path to PsExec.exe. Defaults to PsExec.exe next to this launcher, then falls back to PATH.
The launcher validates that PsExec is available before starting any non-dry-run cycle.

.PARAMETER AllowDsregLeave
Passes -AllowDsregLeave to the repair script. Enabled by default in this launcher.
Use -AllowDsregLeave:$false to run the remote repair without dsregcmd /leave authorization.

.PARAMETER IgnoreRunGuard
Passes -IgnoreRunGuard to the repair script to bypass the target computer's 12-hour guard.
In loop mode, this is passed only during the first cycle unless -IgnoreRunGuardEveryCycle is set.

.PARAMETER AllowRebootWhenNoInteractiveUser
Passes -AllowRebootWhenNoInteractiveUser to the repair script.

.PARAMETER AllowRebootAfterDsregLeave
Passes -AllowRebootAfterDsregLeave to the repair script.

.PARAMETER AllowRemoveNonIntuneMdmEnrollment
Passes -AllowRemoveNonIntuneMdmEnrollment to the repair script. Use only when a non-Intune MDM should be removed before Intune enrollment.

.PARAMETER AllowRemoveStaleIntuneEnrollment
Passes -AllowRemoveStaleIntuneEnrollment to the repair script. Use only to clean stale local Intune enrollment traces.

.PARAMETER AuditOnly
Passes -AuditOnly to the repair script to collect diagnostics without repair actions.

.PARAMETER IntuneInventoryCsv
Optional local CSV export of Intune/Graph devices. Defaults to DevicesIntune.csv in the parent folder, then DevicesIntune.csv next to this script. The launcher adds an IntuneInventoryPresent column by matching computer names.

.PARAMETER EntraInventoryCsv
Optional local CSV export of Microsoft Entra devices. Defaults to DevicesEntra.csv in the parent folder, then DevicesEntra.csv next to this script. The launcher adds Entra registration columns by matching computer names.

.PARAMETER AdInventoryCsv
Optional local CSV export of Active Directory computers. Defaults to DevicesAD.csv in the parent folder, then DevicesAD.csv next to this script. LOT wrappers pass a LOT-local DevicesAD.csv for domain-specific fallback refreshes.

.PARAMETER AdRootInventoryCsv
Optional root forest-wide AD CSV. LOT wrappers pass the toolkit-root DevicesAD.csv here so the launcher can prefer it when it exists and is less than 60 minutes old.

.PARAMETER AdDomain
Optional AD domain/controller used when refreshing DevicesAD.csv. When omitted, the AD export targets all domains in the current AD forest. For LOT runs, set EHJIR_AD_DOMAIN or create AdDomain.txt in the LOT folder only when a domain-specific export is required.

.PARAMETER LogRoot
Local folder where PsExec per-computer logs are written. Defaults to PsExecLogs next to this script.

.PARAMETER ReportRoot
Local folder where cycle CSV/HTML summaries are written. Defaults to Reports next to this script.

.PARAMETER StaleCleanupDelaySeconds
Seconds passed to the repair script after stale local Intune cleanup before same-run auto-enrollment. Defaults to 60.

.PARAMETER DelayBetweenComputersSeconds
Seconds to wait between two job starts. Defaults to 0 for large fleet throughput.

.PARAMETER ThrottleLimit
Maximum number of computers processed in parallel. Defaults to 25.

.PARAMETER JobPollSeconds
Seconds to wait between checks for completed parallel jobs. Defaults to 2.

.PARAMETER DelayBetweenCyclesMinutes
Minutes to wait after one full pass over the computer list before starting the next pass. Defaults to 1.

.PARAMETER MaxCycles
Maximum number of cycles to run. 0 means infinite. Defaults to 0.

.PARAMETER RunOnce
Runs a single cycle and exits.

.PARAMETER DryRun
Checks DNS, ping and administrative share reachability for each computer without copying or executing the repair script.

.PARAMETER IgnoreRunGuardEveryCycle
Passes -IgnoreRunGuard on every loop cycle. Use carefully.

.PARAMETER CentralLogRoot
Local central folder where remote logs are collected after each computer run.
Defaults to CentralLogs next to this launcher. By default, files are stored under one Latest folder per computer.

.PARAMETER KeepCentralLogHistory
Keeps one central log folder per computer and per cycle. Without this switch, CentralLogs\<Computer>\Latest is overwritten each run.

.PARAMETER NoCentralLogCollection
Disables collection of C:\ProgramData\SmartM365\IntuneHybridJoinToolkit from each remote computer.

.PARAMETER RebootDelaySeconds
Seconds used by the remote script when scheduling a controlled reboot. Defaults to 180 so this launcher can pull logs through C$ before reboot.

.PARAMETER IntuneRetrySleepMinutes
Minutes passed to the repair script between local Intune enrollment re-checks after auto-enrollment is triggered. Defaults to 5.

.PARAMETER IntuneRetryMaxRetries
Number of local Intune enrollment re-checks passed to the repair script after auto-enrollment is triggered. Defaults to 5.

.PARAMETER PsExecTimeoutMinutes
Maximum time to wait for one PsExec execution before marking the computer as PSEXEC_TIMEOUT.
Use 0 to disable the timeout. Defaults to 120.

.PARAMETER CommunicationLostEvidenceWaitMinutes
Maximum minutes to wait before collecting remote evidence when PsExec lost communication with
PSEXESVC after starting remote PowerShell. Defaults to 65 so long local retry loops can write
their final CSV/LastRun.json before the launcher pulls logs.

.PARAMETER CommunicationLostEvidencePollMinutes
Minutes between current-run final CSV checks during the communication-lost wait window.
Defaults to 10.

.PARAMETER SkipPostCycleIntuneInventory
Do not refresh Intune inventory at the end of each cycle. By default, the launcher runs a full SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1 inventory and adds post-cycle Intune columns to the report.

.PARAMETER PostCycleIntuneInventoryPageSize
Graph page size used by SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1 for automatic full inventory refreshes. Defaults to 999.
#>

#requires -Version 5.1

[CmdletBinding(PositionalBinding=$false)]
param(
    [string]$ComputerListPath,
    [string]$PsExecPath,
    [int]$DelayBetweenComputersSeconds = 0,
    [int]$ThrottleLimit = 25,
    [int]$JobPollSeconds = 2,
    [int]$DelayBetweenCyclesMinutes = 1,
    [int]$MaxCycles = 0,
    [switch]$AllowDsregLeave = $true,
    [switch]$IgnoreRunGuard,
    [switch]$IgnoreRunGuardEveryCycle,
    [switch]$RunOnce,
    [switch]$DryRun,
    [switch]$AllowRebootWhenNoInteractiveUser,
    [switch]$AllowRebootAfterDsregLeave,
    [switch]$AllowRemoveNonIntuneMdmEnrollment,
    [switch]$AllowRemoveStaleIntuneEnrollment,
    [switch]$AuditOnly,
    [string]$IntuneInventoryCsv,
    [string]$IntuneInventoryNameColumn,
    [string]$EntraInventoryCsv,
    [string]$EntraInventoryNameColumn,
    [string]$AdInventoryCsv,
    [string]$AdRootInventoryCsv,
    [string]$AdInventoryNameColumn,
    [string]$AdDomain,
    [string]$LogRoot,
    [string]$ReportRoot,
    [string]$CentralLogRoot,
    [switch]$KeepCentralLogHistory,
    [switch]$NoCentralLogCollection,
    [int]$StaleCleanupDelaySeconds = 60,
    [int]$RebootDelaySeconds = 180,
    [int]$IntuneRetrySleepMinutes = 5,
    [int]$IntuneRetryMaxRetries = 5,
    [int]$PsExecTimeoutMinutes = 120,
    [int]$CommunicationLostEvidenceWaitMinutes = 65,
    [int]$CommunicationLostEvidencePollMinutes = 10,
    [switch]$SkipPostCycleIntuneInventory,
    [int]$PostCycleIntuneInventoryPageSize = 999,
    [Parameter(ValueFromRemainingArguments=$true)]
    [string[]]$UnexpectedArguments
)

$ErrorActionPreference = "Stop"
$LauncherVersion = "2.10.45"

if ($UnexpectedArguments -and $UnexpectedArguments.Count -gt 0) {
    throw ("Unexpected launcher argument(s): {0}. Pass PsExec with -PsExecPath <path>, not as a free argument." -f ($UnexpectedArguments -join " "))
}

if ($ThrottleLimit -lt 1) { $ThrottleLimit = 1 }
if ($JobPollSeconds -lt 1) { $JobPollSeconds = 1 }
if ($DelayBetweenComputersSeconds -lt 0) { $DelayBetweenComputersSeconds = 0 }
if ($DelayBetweenCyclesMinutes -lt 0) { $DelayBetweenCyclesMinutes = 0 }
if ($RebootDelaySeconds -lt 60) { $RebootDelaySeconds = 60 }
if ($StaleCleanupDelaySeconds -lt 0) { $StaleCleanupDelaySeconds = 0 }
if ($IntuneRetrySleepMinutes -lt 1) { $IntuneRetrySleepMinutes = 1 }
if ($IntuneRetryMaxRetries -lt 1) { $IntuneRetryMaxRetries = 1 }
if ($PsExecTimeoutMinutes -lt 0) { $PsExecTimeoutMinutes = 0 }
if ($CommunicationLostEvidenceWaitMinutes -lt 0) { $CommunicationLostEvidenceWaitMinutes = 0 }
if ($CommunicationLostEvidencePollMinutes -lt 1) { $CommunicationLostEvidencePollMinutes = 1 }
if ($PostCycleIntuneInventoryPageSize -lt 1) { $PostCycleIntuneInventoryPageSize = 1 }
if ($PostCycleIntuneInventoryPageSize -gt 999) { $PostCycleIntuneInventoryPageSize = 999 }
if ($DryRun -and -not $RunOnce -and $MaxCycles -eq 0) { $MaxCycles = 1 }

$BaseDir = if (-not [string]::IsNullOrWhiteSpace($PSScriptRoot)) {
    $PSScriptRoot
}
else {
    Split-Path -Parent $MyInvocation.MyCommand.Path
}

$ScriptName = "SmartM365-Invoke-IntuneHybridJoinRepair.ps1"
$LocalScriptPath = Join-Path $BaseDir $ScriptName
$ExportIntuneScriptPath = Join-Path $BaseDir "SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1"
$ExportEntraScriptPath = Join-Path $BaseDir "SmartM365-IntuneHybridJoinRepair-Export-EntraDevicesCsv.ps1"
$ExportAdScriptPath = Join-Path $BaseDir "SmartM365-IntuneHybridJoinRepair-Export-ADDevicesCsv.ps1"

if ([string]::IsNullOrWhiteSpace($ComputerListPath)) {
    $ComputerListPath = Join-Path $BaseDir "Computers.txt"
}

$IntuneInventoryCsvWasProvided = -not [string]::IsNullOrWhiteSpace($IntuneInventoryCsv)
if (-not $IntuneInventoryCsvWasProvided) {
    $candidateIntuneCsvPaths = @()

    $parentDir = Split-Path -Parent $BaseDir
    if (-not [string]::IsNullOrWhiteSpace($parentDir)) {
        $candidateIntuneCsvPaths += (Join-Path $parentDir "DevicesIntune.csv")
    }
    $candidateIntuneCsvPaths += (Join-Path $BaseDir "DevicesIntune.csv")

    foreach ($candidateIntuneCsvPath in $candidateIntuneCsvPaths) {
        if (Test-Path -LiteralPath $candidateIntuneCsvPath) {
            $IntuneInventoryCsv = $candidateIntuneCsvPath
            break
        }
    }

    if ([string]::IsNullOrWhiteSpace($IntuneInventoryCsv) -and $candidateIntuneCsvPaths.Count -gt 0) {
        $IntuneInventoryCsv = $candidateIntuneCsvPaths[0]
    }
}

$EntraInventoryCsvWasProvided = -not [string]::IsNullOrWhiteSpace($EntraInventoryCsv)
if (-not $EntraInventoryCsvWasProvided) {
    $candidateEntraCsvPaths = @()

    $parentDir = Split-Path -Parent $BaseDir
    if (-not [string]::IsNullOrWhiteSpace($parentDir)) {
        $candidateEntraCsvPaths += (Join-Path $parentDir "DevicesEntra.csv")
    }
    $candidateEntraCsvPaths += (Join-Path $BaseDir "DevicesEntra.csv")

    foreach ($candidateEntraCsvPath in $candidateEntraCsvPaths) {
        if (Test-Path -LiteralPath $candidateEntraCsvPath) {
            $EntraInventoryCsv = $candidateEntraCsvPath
            break
        }
    }

    if ([string]::IsNullOrWhiteSpace($EntraInventoryCsv) -and $candidateEntraCsvPaths.Count -gt 0) {
        $EntraInventoryCsv = $candidateEntraCsvPaths[0]
    }
}

$AdInventoryCsvWasProvided = -not [string]::IsNullOrWhiteSpace($AdInventoryCsv)
$AdRootInventoryCsvWasProvided = -not [string]::IsNullOrWhiteSpace($AdRootInventoryCsv)
$AdInventoryUsesRecentRootCsv = $false
if ($AdRootInventoryCsvWasProvided) {
    $adRootInventoryItem = Get-Item -LiteralPath $AdRootInventoryCsv -ErrorAction SilentlyContinue
    if ($adRootInventoryItem) {
        $adRootInventoryAge = (Get-Date) - $adRootInventoryItem.LastWriteTime
        if ($adRootInventoryAge.TotalMinutes -le 60) {
            $AdInventoryCsv = $adRootInventoryItem.FullName
            $AdInventoryUsesRecentRootCsv = $true
            $AdInventoryCsvWasProvided = $true
        }
    }
}
if (-not $AdInventoryCsvWasProvided) {
    $candidateAdCsvPaths = @()

    $parentDir = Split-Path -Parent $BaseDir
    if (-not [string]::IsNullOrWhiteSpace($parentDir)) {
        $candidateAdCsvPaths += (Join-Path $parentDir "DevicesAD.csv")
    }
    $candidateAdCsvPaths += (Join-Path $BaseDir "DevicesAD.csv")

    foreach ($candidateAdCsvPath in $candidateAdCsvPaths) {
        if (Test-Path -LiteralPath $candidateAdCsvPath) {
            $AdInventoryCsv = $candidateAdCsvPath
            break
        }
    }

    if ([string]::IsNullOrWhiteSpace($AdInventoryCsv) -and $candidateAdCsvPaths.Count -gt 0) {
        $AdInventoryCsv = $candidateAdCsvPaths[0]
    }
}

if ([string]::IsNullOrWhiteSpace($PsExecPath)) {
    $PsExecPath = Join-Path $BaseDir "PsExec.exe"
}

if ([string]::IsNullOrWhiteSpace($LogRoot)) {
    $LogRoot = Join-Path $BaseDir "PsExecLogs"
}
if ([string]::IsNullOrWhiteSpace($ReportRoot)) {
    $ReportRoot = Join-Path $BaseDir "Reports"
}
$CollectRemoteLogs = -not [bool]$NoCentralLogCollection
if ([string]::IsNullOrWhiteSpace($CentralLogRoot)) {
    $CentralLogRoot = Join-Path $BaseDir "CentralLogs"
}

$ComputerListPath = [System.IO.Path]::GetFullPath($ComputerListPath)
$LocalScriptPath = [System.IO.Path]::GetFullPath($LocalScriptPath)
if (-not [string]::IsNullOrWhiteSpace($IntuneInventoryCsv)) { $IntuneInventoryCsv = [System.IO.Path]::GetFullPath($IntuneInventoryCsv) }
if (-not [string]::IsNullOrWhiteSpace($EntraInventoryCsv)) { $EntraInventoryCsv = [System.IO.Path]::GetFullPath($EntraInventoryCsv) }
if (-not [string]::IsNullOrWhiteSpace($AdInventoryCsv)) { $AdInventoryCsv = [System.IO.Path]::GetFullPath($AdInventoryCsv) }
$LogRoot = [System.IO.Path]::GetFullPath($LogRoot)
$ReportRoot = [System.IO.Path]::GetFullPath($ReportRoot)
$CentralLogRoot = [System.IO.Path]::GetFullPath($CentralLogRoot)

$RemoteRelativeDir = "ProgramData\SmartM365\IntuneHybridJoinToolkit"
$RemoteScriptPath = "C:\ProgramData\SmartM365\IntuneHybridJoinToolkit\$ScriptName"
$RemoteDataRelativeDir = "ProgramData\SmartM365\IntuneHybridJoinToolkit"

if (-not (Test-Path -LiteralPath $LocalScriptPath)) {
    throw "Repair script not found: $LocalScriptPath"
}

if (-not (Test-Path -LiteralPath $ComputerListPath)) {
    throw "Computer list not found: $ComputerListPath"
}

function Resolve-PsExecPath {
    param(
        [Parameter(Mandatory=$true)][string]$Path
    )

    $candidatePath = $Path.Trim().Trim([char]34)
    $embeddedPsExecPaths = @(
        [regex]::Matches($candidatePath, '(?i)[A-Z]:\\[^\r\n"]*?PsExec(?:64)?\.exe') |
            ForEach-Object { $_.Value.Trim().Trim([char]34) } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
    )

    if ($embeddedPsExecPaths.Count -gt 1) {
        foreach ($embeddedPath in $embeddedPsExecPaths) {
            if (Test-Path -LiteralPath $embeddedPath -PathType Leaf) {
                Write-Warning ("PsExecPath contained multiple executable paths. Using '{0}' from: {1}" -f $embeddedPath, ($embeddedPsExecPaths -join " | "))
                return (Get-Item -LiteralPath $embeddedPath -ErrorAction Stop).FullName
            }
        }

        throw ("Invalid PsExecPath value: {0}. Multiple PsExec paths were provided, but none exists as a file. Provide exactly one executable path, for example -PsExecPath ""C:\Sysinternals\PsExec.exe""." -f $Path)
    }

    if ($embeddedPsExecPaths.Count -eq 1 -and $candidatePath -ne $embeddedPsExecPaths[0]) {
        $candidatePath = $embeddedPsExecPaths[0]
    }

    if (-not [string]::IsNullOrWhiteSpace($candidatePath) -and (Test-Path -LiteralPath $candidatePath -PathType Leaf)) {
        return (Get-Item -LiteralPath $candidatePath -ErrorAction Stop).FullName
    }

    if ($candidatePath -match '(?i)\.exe["'']?\s+\S') {
        throw ("Invalid PsExecPath value: {0}. Provide exactly one executable path, for example -PsExecPath ""C:\Sysinternals\PsExec.exe""." -f $Path)
    }

    if (-not [string]::IsNullOrWhiteSpace($candidatePath)) {
        $command = Get-Command -Name $candidatePath -CommandType Application -ErrorAction SilentlyContinue
        if ($command) {
            return $command.Source
        }
    }

    $pathCommand = Get-Command -Name "PsExec.exe" -CommandType Application -ErrorAction SilentlyContinue
    if ($pathCommand) {
        return $pathCommand.Source
    }

    throw ("PsExec.exe not found. Place PsExec.exe next to this launcher ({0}) or add PsExec.exe to PATH. You can also pass -PsExecPath <path>." -f $BaseDir)
}

if (-not $DryRun) {
    $PsExecPath = Resolve-PsExecPath -Path $PsExecPath
}

if (-not (Test-Path -LiteralPath $LogRoot)) {
    New-Item -ItemType Directory -Path $LogRoot -Force | Out-Null
}
if (-not (Test-Path -LiteralPath $ReportRoot)) {
    New-Item -ItemType Directory -Path $ReportRoot -Force | Out-Null
}

if ($CollectRemoteLogs -and -not (Test-Path -LiteralPath $CentralLogRoot)) {
    New-Item -ItemType Directory -Path $CentralLogRoot -Force | Out-Null
}

function Get-IntuneInventorySet {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$false)][string]$NameColumn
    )

    $set = @{}
    if ([string]::IsNullOrWhiteSpace($Path)) { return $set }
    if (-not (Test-Path -LiteralPath $Path)) { throw "Intune inventory CSV not found: $Path" }

    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { return $set }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        $candidateColumns = @("ComputerName","computerName","DeviceName","deviceName","Name","name","managedDeviceName","ManagedDeviceName")
        $first = $rows | Select-Object -First 1
        foreach ($candidate in $candidateColumns) {
            if ($first.PSObject.Properties.Name -contains $candidate) {
                $NameColumn = $candidate
                break
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        throw "Unable to infer the Intune inventory device name column. Use -IntuneInventoryNameColumn."
    }

    foreach ($row in $rows) {
        if ($row.PSObject.Properties.Name -contains "IntuneInventoryPresent") {
            $presentValue = ([string]$row.IntuneInventoryPresent).Trim()
            if ($presentValue -notin @("True","true","1","YES","Yes","yes","OUI","Oui","oui")) {
                continue
            }
        }

        $value = [string]$row.$NameColumn
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            $short = ($value.Trim().Split(".")[0]).ToUpperInvariant()
            if (-not $set.ContainsKey($short)) { $set[$short] = $true }
        }
    }

    return $set
}

function Get-EntraInventoryMap {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$false)][string]$NameColumn
    )

    $map = @{}
    if ([string]::IsNullOrWhiteSpace($Path)) { return $map }
    if (-not (Test-Path -LiteralPath $Path)) { return $map }

    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { return $map }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        $candidateColumns = @("ComputerName","computerName","DisplayName","displayName","DeviceName","deviceName","Name","name")
        $first = $rows | Select-Object -First 1
        foreach ($candidate in $candidateColumns) {
            if ($first.PSObject.Properties.Name -contains $candidate) {
                $NameColumn = $candidate
                break
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        throw "Unable to infer the Entra inventory device name column. Use -EntraInventoryNameColumn."
    }

    foreach ($row in $rows) {
        $value = [string]$row.$NameColumn
        if ([string]::IsNullOrWhiteSpace($value)) { continue }

        $short = ($value.Trim().Split(".")[0]).ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($short)) { continue }

        $present = $true
        if ($row.PSObject.Properties.Name -contains "EntraInventoryPresent") {
            $present = Test-BooleanLikeTrue -Value $row.EntraInventoryPresent
        }
        if (-not $present) { continue }

        if (-not $map.ContainsKey($short)) {
            $map[$short] = $row
        }
    }

    return $map
}

function Get-AdInventoryMap {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$false)][string]$NameColumn
    )

    $map = @{}
    if ([string]::IsNullOrWhiteSpace($Path)) { return $map }
    if (-not (Test-Path -LiteralPath $Path)) { return $map }

    $rows = @(Import-Csv -LiteralPath $Path)
    if ($rows.Count -eq 0) { return $map }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        $candidateColumns = @("ComputerName","computerName","DNSHostName","dnsHostName","Name","name")
        $first = $rows | Select-Object -First 1
        foreach ($candidate in $candidateColumns) {
            if ($first.PSObject.Properties.Name -contains $candidate) {
                $NameColumn = $candidate
                break
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($NameColumn)) {
        throw "Unable to infer the AD inventory device name column. Use -AdInventoryNameColumn."
    }

    foreach ($row in $rows) {
        $value = [string]$row.$NameColumn
        if ([string]::IsNullOrWhiteSpace($value)) { continue }

        $short = ($value.Trim().Split(".")[0]).ToUpperInvariant()
        if ([string]::IsNullOrWhiteSpace($short)) { continue }

        $present = $true
        if ($row.PSObject.Properties.Name -contains "ADInventoryPresent") {
            $present = Test-BooleanLikeTrue -Value $row.ADInventoryPresent
        }
        if (-not $present) { continue }

        if (-not $map.ContainsKey($short)) {
            $map[$short] = $row
        }
    }

    return $map
}

function Test-BooleanLikeTrue {
    param([AllowNull()][object]$Value)

    if ($Value -eq $true) { return $true }
    $text = ([string]$Value).Trim()
    return ($text -in @("True","true","1","YES","Yes","yes","OUI","Oui","oui"))
}

function Invoke-FullIntuneInventoryExport {
    param(
        [Parameter(Mandatory=$true)][string]$ExportScriptPath,
        [Parameter(Mandatory=$true)][string]$OutputPath,
        [Parameter(Mandatory=$true)][string]$LogPath,
        [Parameter(Mandatory=$true)][int]$PageSize
    )

    try {
        if (-not (Test-Path -LiteralPath $ExportScriptPath)) {
            throw "SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1 not found: $ExportScriptPath"
        }

        $args = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", $ExportScriptPath,
            "-OutputPath", $OutputPath,
            "-PageSize", ([string]$PageSize),
            "-ForceRefresh"
        )

        $output = & powershell.exe @args 2>&1
        $exitCode = $LASTEXITCODE
        $output | Out-File -LiteralPath $LogPath -Encoding UTF8 -Force

        if ($exitCode -ne 0) {
            throw "SmartM365-IntuneHybridJoinRepair-Export-IntuneDevicesCsv.ps1 exited with code $exitCode. Log=$LogPath"
        }
        if (-not (Test-Path -LiteralPath $OutputPath)) {
            throw "Intune CSV was not created: $OutputPath"
        }

        $set = Get-IntuneInventorySet -Path $OutputPath -NameColumn "ComputerName"

        return [PSCustomObject]@{
            Success = $true
            CsvPath = $OutputPath
            LogPath = $LogPath
            InventorySet = $set
            Error = ""
        }
    }
    catch {
        return [PSCustomObject]@{
            Success = $false
            CsvPath = $OutputPath
            LogPath = $LogPath
            InventorySet = @{}
            Error = $_.Exception.Message
        }
    }
}

function Invoke-FullEntraInventoryExport {
    param(
        [Parameter(Mandatory=$true)][string]$ExportScriptPath,
        [Parameter(Mandatory=$true)][string]$OutputPath,
        [Parameter(Mandatory=$true)][string]$LogPath,
        [Parameter(Mandatory=$true)][int]$PageSize
    )

    try {
        if (-not (Test-Path -LiteralPath $ExportScriptPath)) {
            throw "SmartM365-IntuneHybridJoinRepair-Export-EntraDevicesCsv.ps1 not found: $ExportScriptPath"
        }

        $args = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", $ExportScriptPath,
            "-OutputPath", $OutputPath,
            "-PageSize", ([string]$PageSize),
            "-ForceRefresh"
        )

        $output = & powershell.exe @args 2>&1
        $exitCode = $LASTEXITCODE
        $output | Out-File -LiteralPath $LogPath -Encoding UTF8 -Force

        if ($exitCode -ne 0) {
            throw "SmartM365-IntuneHybridJoinRepair-Export-EntraDevicesCsv.ps1 exited with code $exitCode. Log=$LogPath"
        }
        if (-not (Test-Path -LiteralPath $OutputPath)) {
            throw "Entra CSV was not created: $OutputPath"
        }

        $map = Get-EntraInventoryMap -Path $OutputPath -NameColumn "ComputerName"

        return [PSCustomObject]@{
            Success = $true
            CsvPath = $OutputPath
            LogPath = $LogPath
            InventoryMap = $map
            Error = ""
        }
    }
    catch {
        return [PSCustomObject]@{
            Success = $false
            CsvPath = $OutputPath
            LogPath = $LogPath
            InventoryMap = @{}
            Error = $_.Exception.Message
        }
    }
}

function Invoke-FullAdInventoryExport {
    param(
        [Parameter(Mandatory=$true)][string]$ExportScriptPath,
        [Parameter(Mandatory=$true)][string]$OutputPath,
        [Parameter(Mandatory=$true)][string]$LogPath,
        [Parameter(Mandatory=$false)][string]$Domain
    )

    try {
        if (-not (Test-Path -LiteralPath $ExportScriptPath)) {
            throw "SmartM365-IntuneHybridJoinRepair-Export-ADDevicesCsv.ps1 not found: $ExportScriptPath"
        }

        $args = @(
            "-NoProfile",
            "-ExecutionPolicy", "Bypass",
            "-File", $ExportScriptPath,
            "-OutputPath", $OutputPath,
            "-ForceRefresh"
        )
        if (-not [string]::IsNullOrWhiteSpace($Domain)) {
            $args += "-Domain"
            $args += $Domain
        }

        $output = & powershell.exe @args 2>&1
        $exitCode = $LASTEXITCODE
        $output | Out-File -LiteralPath $LogPath -Encoding UTF8 -Force

        if ($exitCode -ne 0) {
            throw "SmartM365-IntuneHybridJoinRepair-Export-ADDevicesCsv.ps1 exited with code $exitCode. Log=$LogPath"
        }
        if (-not (Test-Path -LiteralPath $OutputPath)) {
            throw "AD CSV was not created: $OutputPath"
        }

        $map = Get-AdInventoryMap -Path $OutputPath -NameColumn "ComputerName"

        return [PSCustomObject]@{
            Success = $true
            CsvPath = $OutputPath
            LogPath = $LogPath
            InventoryMap = $map
            Error = ""
        }
    }
    catch {
        return [PSCustomObject]@{
            Success = $false
            CsvPath = $OutputPath
            LogPath = $LogPath
            InventoryMap = @{}
            Error = $_.Exception.Message
        }
    }
}

function Get-NextActionFromLauncherStatus {
    param([Parameter(Mandatory=$true)][string]$Status)

    switch ($Status) {
        "SUCCESS" { return "NO_ACTION_ALREADY_INTUNE_OR_HEALTHY" }
        "AUDIT_SUCCESS_ALREADY_INTUNE" { return "NO_ACTION_ALREADY_INTUNE" }
        "AUDIT_INTUNE_MISSING" { return "RUN_REPAIR" }
        "AUDIT_STALE_INTUNE_ENROLLMENT_LOCAL" { return "CLEAN_STALE_INTUNE_OPTIN" }
        "INTUNE_ENROLLMENT_PENDING_CONFIRMATION" { return "RECHECK_LATER_INTUNE_ENROLLMENT" }
        "ADMIN_SHARE_UNREACHABLE" { return "FIX_ADMIN_SHARE_OR_NETWORK" }
        "RUN_GUARD_ACTIVE" { return "WAIT_RUN_GUARD" }
        "REBOOT_TRIGGERED_WAITING_FOR_USER_LOGON" { return "WAIT_USER_LOGON" }
        "WAITING_FOR_INTERACTIVE_USER_LOGON" { return "WAIT_USER_LOGON" }
        "INTUNE_USER_AUTOENROLL_LOCAL_INTERACTIVE_USER" { return "LOGON_WITH_DOMAIN_OR_AAD_USER" }
        "INTUNE_USER_AUTOENROLL_TASK_NOT_FOUND" { return "FIX_GPO_USER_AUTOENROLL_TASK" }
        "STALE_INTUNE_ENROLLMENT_LOCAL" { return "CLEAN_STALE_INTUNE_OPTIN" }
        "NON_INTUNE_MDM_ENROLLED" { return "CLEAN_NON_INTUNE_MDM_OPTIN" }
        "ENTRA_HYBRID_PENDING_ADJ_TRIGGERED" { return "RECHECK_ENTRA_PENDING_AFTER_ADJ" }
        "ENTRA_HYBRID_PENDING_RETRY_EXHAUSTED" { return "CHECK_AD_CONNECT_OR_DUPLICATE_ENTRA_DEVICE" }
        "ENTRA_PENDING_RESOLVED_POST_CYCLE" { return "RECHECK_INTUNE_ENROLLMENT" }
        "USER_NOT_AZUREAD" { return "CHECK_USER_AAD_OR_LOGON_CONTEXT" }
        "USER_PRT_NOT_AVAILABLE" { return "CHECK_USER_PRT" }
        "USER_PRT_REFRESH_FAILED" { return "FIX_USER_PRT_OR_RELOGIN" }
        "USER_SESSION_REMOTE" { return "LOGON_ON_CONSOLE" }
        "INTUNE_AUTOENROLL_POLICY_NOT_CONFIGURED" { return "CHECK_GPO_AUTOENROLL" }
        "KEY_SIGN_TEST_FAILED" { return "REPAIR_HYBRID_JOIN_KEY_OR_ALLOW_LEAVE" }
        "INTUNE_ENROLLMENT_CONNECTIVITY_FAILED" { return "CHECK_CONNECTIVITY" }
        "DOMAIN_CONTROLLER_UNREACHABLE" { return "FIX_DOMAIN_CONNECTIVITY_OR_VPN" }
        "REMOTE_DIRECTORY_CREATE_FAILED" { return "FIX_SCRIPT_COPY_OR_ADMIN_SHARE" }
        "REMOTE_SCRIPT_COPY_FAILED" { return "FIX_SCRIPT_COPY_OR_SECURITY" }
        "REMOTE_SCRIPT_MISSING" { return "FIX_SCRIPT_COPY_OR_SECURITY" }
        "PSEXEC_TIMEOUT" { return "CHECK_REMOTE_LOG_OR_RETRY" }
        "PSEXEC_COMMUNICATION_LOST" { return "RETRY_PSEXEC_OR_CHECK_REMOTE_SERVICE" }
        "PSEXEC_EXIT_UNKNOWN" { return "CHECK_CURRENT_RUN_REMOTE_LOG" }
        default {
            if ($Status -like "ERROR*") { return "CHECK_CONNECTIVITY_OR_ADMIN_ACCESS" }
            if ($Status -like "PSEXEC_EXIT*") { return "CHECK_CURRENT_RUN_REMOTE_LOG" }
            return "REVIEW_LOGS"
        }
    }
}

function Get-ScriptVersionFromFile {
    param([Parameter(Mandatory=$true)][string]$Path)

    try {
        if (-not (Test-Path -LiteralPath $Path)) { return "" }
        $match = Select-String -LiteralPath $Path -Pattern '^\s*\$ScriptVersion\s*=\s*"([^"]+)"' -ErrorAction Stop | Select-Object -First 1
        if ($match -and $match.Matches.Count -gt 0) {
            return $match.Matches[0].Groups[1].Value
        }
    }
    catch { }

    return ""
}

function Get-FileSha256 {
    param([Parameter(Mandatory=$true)][string]$Path)

    try {
        if (-not (Test-Path -LiteralPath $Path)) { return "" }
        return (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash
    }
    catch {
        return ""
    }
}

function Get-RemoteEvidenceFinalStatus {
    param(
        [Parameter(Mandatory=$true)][string]$EvidencePath,
        [Parameter(Mandatory=$false)][datetime]$Since = [datetime]::MinValue,
        [switch]$RequireCompletedRun
    )

    $completedRunId = ""
    $completedRunStatus = ""
    $completedRunExitCode = ""
    $completedRunNextAction = ""
    $completedRunDetail = ""
    $lastRunPath = Join-Path $EvidencePath "LastRun.json"
    if ($RequireCompletedRun) {
        if (-not (Test-Path -LiteralPath $lastRunPath -ErrorAction SilentlyContinue)) { return $null }
        try {
            $lastRun = Get-Content -LiteralPath $lastRunPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            if ($null -eq $lastRun -or [string]::IsNullOrWhiteSpace([string]$lastRun.RunId) -or [string]::IsNullOrWhiteSpace([string]$lastRun.EndTime)) {
                return $null
            }
            $lastRunStart = [datetime]$lastRun.StartTime
            if ($lastRunStart -lt $Since.AddSeconds(-5)) { return $null }
            $completedRunId = ([string]$lastRun.RunId).Trim()
            $completedRunStatus = ([string]$lastRun.Status).Trim()
            $completedRunExitCode = ([string]$lastRun.ExitCode).Trim()
            $completedRunNextAction = ([string]$lastRun.NextAction).Trim()
            $completedRunDetail = ([string]$lastRun.Detail).Trim()
            if ([string]::IsNullOrWhiteSpace($completedRunStatus) -or [string]::IsNullOrWhiteSpace($completedRunExitCode)) {
                return $null
            }
        }
        catch {
            return $null
        }
    }

    $csv = Get-ChildItem -LiteralPath $EvidencePath -Recurse -File -Filter "IntuneHybridJoinToolkit_*.csv" -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $Since.AddSeconds(-5) } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if (-not $csv) {
        if ($RequireCompletedRun -and -not [string]::IsNullOrWhiteSpace($completedRunStatus)) {
            return [PSCustomObject]@{
                RunId = $completedRunId
                Status = $completedRunStatus
                ExitCode = $completedRunExitCode
                NextAction = $completedRunNextAction
                Detail = $completedRunDetail
                CsvPath = ""
                LastRunPath = $lastRunPath
                LastRunStatus = $completedRunStatus
                LastRunExitCode = $completedRunExitCode
                InteractiveUserName = ""
                InteractiveUserDomain = ""
                InteractiveUserAccountName = ""
                InteractiveUserAccountType = ""
                InteractiveSessionName = ""
                InteractiveSessionState = ""
                UserIsUserAzureAD = ""
                UserAzureAdPrt = ""
                UserSessionIsNotRemote = ""
            }
        }
        return $null
    }

    try {
        $row = Import-Csv -LiteralPath $csv.FullName -ErrorAction Stop | Select-Object -Last 1
        if ($null -eq $row -or [string]::IsNullOrWhiteSpace([string]$row.Status)) { return $null }

        $rowRunId = ""
        if ($row.PSObject.Properties["RunId"]) { $rowRunId = ([string]$row.RunId).Trim() }
        if ($RequireCompletedRun) {
            if ([string]::IsNullOrWhiteSpace($rowRunId) -or $rowRunId -ne $completedRunId) { return $null }
        }

        $exitCode = ""
        if ($row.PSObject.Properties["ExitCode"]) { $exitCode = [string]$row.ExitCode }

        $nextAction = ""
        if ($row.PSObject.Properties["NextAction"]) { $nextAction = [string]$row.NextAction }
        if ([string]::IsNullOrWhiteSpace($nextAction) -and $RequireCompletedRun) { $nextAction = $completedRunNextAction }

        $detail = ""
        if ($row.PSObject.Properties["DsregStatusErrorMessage"]) { $detail = [string]$row.DsregStatusErrorMessage }
        if ([string]::IsNullOrWhiteSpace($detail) -and $row.PSObject.Properties["ErrorMessage"]) { $detail = [string]$row.ErrorMessage }
        if ([string]::IsNullOrWhiteSpace($detail) -and $RequireCompletedRun) { $detail = $completedRunDetail }

        $interactiveUserName = ""; if ($row.PSObject.Properties["InteractiveUserName"]) { $interactiveUserName = [string]$row.InteractiveUserName }
        $interactiveUserDomain = ""; if ($row.PSObject.Properties["InteractiveUserDomain"]) { $interactiveUserDomain = [string]$row.InteractiveUserDomain }
        $interactiveUserAccountName = ""; if ($row.PSObject.Properties["InteractiveUserAccountName"]) { $interactiveUserAccountName = [string]$row.InteractiveUserAccountName }
        $interactiveUserAccountType = ""; if ($row.PSObject.Properties["InteractiveUserAccountType"]) { $interactiveUserAccountType = [string]$row.InteractiveUserAccountType }
        $interactiveSessionName = ""; if ($row.PSObject.Properties["InteractiveSessionName"]) { $interactiveSessionName = [string]$row.InteractiveSessionName }
        $interactiveSessionState = ""; if ($row.PSObject.Properties["InteractiveSessionState"]) { $interactiveSessionState = [string]$row.InteractiveSessionState }
        $userIsUserAzureAD = ""; if ($row.PSObject.Properties["User_IsUserAzureAD"]) { $userIsUserAzureAD = [string]$row.User_IsUserAzureAD }
        $userAzureAdPrt = ""; if ($row.PSObject.Properties["User_AzureAdPrt"]) { $userAzureAdPrt = [string]$row.User_AzureAdPrt }
        $userSessionIsNotRemote = ""; if ($row.PSObject.Properties["User_SessionIsNotRemote"]) { $userSessionIsNotRemote = [string]$row.User_SessionIsNotRemote }

        return [PSCustomObject]@{
            RunId = $rowRunId
            Status = $(if ($RequireCompletedRun) { $completedRunStatus } else { ([string]$row.Status).Trim() })
            ExitCode = $(if ($RequireCompletedRun) { $completedRunExitCode } else { $exitCode.Trim() })
            NextAction = $nextAction.Trim()
            Detail = $detail.Trim()
            CsvPath = $csv.FullName
            LastRunPath = $(if ($RequireCompletedRun) { $lastRunPath } else { "" })
            LastRunStatus = $completedRunStatus
            LastRunExitCode = $completedRunExitCode
            InteractiveUserName = $interactiveUserName.Trim()
            InteractiveUserDomain = $interactiveUserDomain.Trim()
            InteractiveUserAccountName = $interactiveUserAccountName.Trim()
            InteractiveUserAccountType = $interactiveUserAccountType.Trim()
            InteractiveSessionName = $interactiveSessionName.Trim()
            InteractiveSessionState = $interactiveSessionState.Trim()
            UserIsUserAzureAD = $userIsUserAzureAD.Trim()
            UserAzureAdPrt = $userAzureAdPrt.Trim()
            UserSessionIsNotRemote = $userSessionIsNotRemote.Trim()
        }
    }
    catch {
        return [PSCustomObject]@{
            Status = ""
            ExitCode = ""
            NextAction = ""
            Detail = ("Could not parse remote evidence CSV: {0}" -f $_.Exception.Message)
            CsvPath = $csv.FullName
        }
    }
}

function Get-LauncherReportColumns {
    @(
        "LauncherVersion",
        "Cycle",
        "Computer",
        "Timestamp",
        "DryRun",
        "DnsResolved",
        "DnsAddressList",
        "AdminShareReachable",
        "PingReachable",
        "RemoteDirectoryCreated",
        "ScriptCopied",
        "LocalScriptVersion",
        "RemoteScriptVersion",
        "LocalScriptHash",
        "RemoteScriptHash",
        "PsExecExitCode",
        "RemoteStatus",
        "RemoteExitCode",
        "RemoteNextAction",
        "RemoteDetail",
        "NextAction",
        "EffectiveStatus",
        "EffectiveNextAction",
        "InteractiveUserName",
        "InteractiveUserDomain",
        "InteractiveUserAccountName",
        "InteractiveUserAccountType",
        "InteractiveSessionName",
        "InteractiveSessionState",
        "UserIsUserAzureAD",
        "UserAzureAdPrt",
        "UserSessionIsNotRemote",
        "IntuneInventoryPresent",
        "EntraInventoryPresent",
        "EntraRegisteredState",
        "EntraAlternativeSecurityIdCount",
        "EntraPendingReason",
        "EntraRegistrationDateTime",
        "EntraTrustType",
        "EntraDeviceId",
        "EntraObjectId",
        "ADInventoryPresent",
        "ADDomain",
        "ADEnabled",
        "ADDNSHostName",
        "ADDistinguishedName",
        "ADOperatingSystem",
        "ADLastLogonTimestampUtc",
        "AdminShareFailureType",
        "PostCycleIntuneInventoryChecked",
        "PostCycleIntuneInventoryPresent",
        "PostCycleIntuneEnrollmentDetected",
        "PostCycleIntuneInventoryCsv",
        "PostCycleIntuneInventoryError",
        "PostCycleEntraInventoryChecked",
        "PostCycleEntraInventoryPresent",
        "PostCycleEntraRegisteredState",
        "PostCycleEntraAlternativeSecurityIdCount",
        "PostCycleEntraPendingResolved",
        "PostCycleEntraInventoryCsv",
        "PostCycleEntraInventoryError",
        "PostCycleADInventoryChecked",
        "PostCycleADInventoryPresent",
        "PostCycleADInventoryCsv",
        "PostCycleADInventoryError",
        "RemoteLogsCollected",
        "RemoteLogsPath",
        "RemoteCurrentRunLogsPath",
        "RemoteLogsError",
        "Status",
        "LogPath",
        "ErrorMessage",
        "JobErrorMessage"
    )
}

function Initialize-LiveCycleReport {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string[]]$Columns
    )

    $header = ($Columns | ForEach-Object { '"' + ($_ -replace '"','""') + '"' }) -join ","
    Set-Content -LiteralPath $Path -Value $header -Encoding UTF8 -Force
}

function Add-LiveCycleReportRow {
    param(
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][string[]]$Columns,
        [Parameter(Mandatory=$true)][psobject]$Row
    )

    $Row | Select-Object $Columns | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 -Append
}

function Get-ComputerList {
    param([Parameter(Mandatory=$true)][string]$Path)

    @(Get-Content -LiteralPath $Path |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.StartsWith("#") } |
        Select-Object -Unique)
}

function Get-ComputerListKey {
    param([Parameter(Mandatory=$true)][string]$ComputerName)

    return ($ComputerName.Trim().Split(".")[0]).ToUpperInvariant()
}

function Test-AlreadyEnrolledCycleResult {
    param([Parameter(Mandatory=$true)][psobject]$Result)

    $status = ""
    $nextAction = ""
    if ($Result.PSObject.Properties["Status"]) { $status = ([string]$Result.Status).Trim() }
    if ($Result.PSObject.Properties["NextAction"]) { $nextAction = ([string]$Result.NextAction).Trim() }

    if ($Result.PSObject.Properties["IntuneInventoryPresent"] -and $Result.IntuneInventoryPresent -eq $true) {
        return $true
    }

    if ($Result.PSObject.Properties["PostCycleIntuneInventoryPresent"] -and $Result.PostCycleIntuneInventoryPresent -eq $true) {
        return $true
    }

    if ($status -in @("SUCCESS","AUDIT_SUCCESS_ALREADY_INTUNE")) {
        return $true
    }

    if ($nextAction -in @("NO_ACTION_ALREADY_INTUNE","NO_ACTION_ALREADY_INTUNE_OR_HEALTHY","NO_ACTION_INTUNE_PRESENT")) {
        return $true
    }

    return $false
}

function Move-AlreadyEnrolledComputersFromList {
    param(
        [Parameter(Mandatory=$true)][string]$ComputerListPath,
        [Parameter(Mandatory=$true)][object[]]$CycleSummary
    )

    $alreadyEnrolled = @(
        $CycleSummary |
            Where-Object { $_ -and (Test-AlreadyEnrolledCycleResult -Result $_) } |
            ForEach-Object { [string]$_.Computer } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
    )

    if ($alreadyEnrolled.Count -eq 0) {
        return [PSCustomObject]@{
            Moved = 0
            AlreadyEnrolledPath = ""
            Detail = "No already-enrolled computer detected in this cycle."
        }
    }

    $moveKeys = @{}
    foreach ($computer in $alreadyEnrolled) {
        $key = Get-ComputerListKey -ComputerName $computer
        if (-not [string]::IsNullOrWhiteSpace($key) -and -not $moveKeys.ContainsKey($key)) {
            $moveKeys[$key] = $computer.Trim()
        }
    }

    $listLines = @()
    if (Test-Path -LiteralPath $ComputerListPath) {
        $listLines = @(Get-Content -LiteralPath $ComputerListPath -ErrorAction Stop)
    }

    $remainingLines = New-Object System.Collections.Generic.List[string]
    $movedFromList = New-Object System.Collections.Generic.List[string]

    foreach ($line in $listLines) {
        $trimmed = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed.StartsWith("#")) {
            $remainingLines.Add($line)
            continue
        }

        $key = Get-ComputerListKey -ComputerName $trimmed
        if ($moveKeys.ContainsKey($key)) {
            $movedFromList.Add($trimmed)
            continue
        }

        $remainingLines.Add($line)
    }

    if ($movedFromList.Count -eq 0) {
        return [PSCustomObject]@{
            Moved = 0
            AlreadyEnrolledPath = ""
            Detail = "Already-enrolled computers were detected, but none were still present in Computers.txt."
        }
    }

    $computerListDir = Split-Path -Parent $ComputerListPath
    if ([string]::IsNullOrWhiteSpace($computerListDir)) { $computerListDir = "." }
    $alreadyEnrolledPath = Join-Path $computerListDir "ComputersAlreadyEnrolled.txt"

    $existingKeys = @{}
    if (Test-Path -LiteralPath $alreadyEnrolledPath) {
        foreach ($line in @(Get-Content -LiteralPath $alreadyEnrolledPath -ErrorAction SilentlyContinue)) {
            $trimmed = $line.Trim()
            if ([string]::IsNullOrWhiteSpace($trimmed) -or $trimmed.StartsWith("#")) { continue }
            $key = Get-ComputerListKey -ComputerName $trimmed
            if (-not $existingKeys.ContainsKey($key)) { $existingKeys[$key] = $true }
        }
    }

    $appendLines = New-Object System.Collections.Generic.List[string]
    foreach ($computer in $movedFromList) {
        $key = Get-ComputerListKey -ComputerName $computer
        if (-not $existingKeys.ContainsKey($key)) {
            $appendLines.Add($computer)
            $existingKeys[$key] = $true
        }
    }

    $tmpComputerListPath = "{0}.tmp.{1}.txt" -f $ComputerListPath,([guid]::NewGuid().ToString("N"))
    try {
        Set-Content -LiteralPath $tmpComputerListPath -Value $remainingLines -Encoding ASCII -Force
        Move-Item -LiteralPath $tmpComputerListPath -Destination $ComputerListPath -Force
    }
    finally {
        Remove-Item -LiteralPath $tmpComputerListPath -Force -ErrorAction SilentlyContinue
    }

    if ($appendLines.Count -gt 0) {
        Add-Content -LiteralPath $alreadyEnrolledPath -Value $appendLines -Encoding ASCII
    }
    elseif (-not (Test-Path -LiteralPath $alreadyEnrolledPath)) {
        New-Item -ItemType File -Path $alreadyEnrolledPath -Force | Out-Null
    }

    return [PSCustomObject]@{
        Moved = $movedFromList.Count
        AlreadyEnrolledPath = $alreadyEnrolledPath
        Detail = ("Moved {0} computer(s) from Computers.txt to ComputersAlreadyEnrolled.txt." -f $movedFromList.Count)
    }
}

function ConvertTo-HtmlText {
    param([object]$Value)

    if ($null -eq $Value) {
        return ""
    }

    return [System.Security.SecurityElement]::Escape([string]$Value)
}

function ConvertTo-SimpleHtmlTable {
    param(
        [Parameter(Mandatory=$true)][AllowEmptyCollection()][object[]]$Rows,
        [string[]]$Columns
    )

    if (-not $Rows -or $Rows.Count -eq 0) {
        return "<p>No rows.</p>"
    }

    if (-not $Columns -or $Columns.Count -eq 0) {
        $columnSet = New-Object System.Collections.Generic.List[string]
        foreach ($row in $Rows) {
            foreach ($property in $row.PSObject.Properties) {
                if (-not $columnSet.Contains($property.Name)) {
                    [void]$columnSet.Add($property.Name)
                }
            }
        }
        $Columns = @($columnSet)
    }

    $html = New-Object System.Collections.Generic.List[string]
    [void]$html.Add("<table>")
    [void]$html.Add("<tr>")
    foreach ($column in $Columns) {
        [void]$html.Add(("<th>{0}</th>" -f (ConvertTo-HtmlText $column)))
    }
    [void]$html.Add("</tr>")

    foreach ($row in $Rows) {
        [void]$html.Add("<tr>")
        foreach ($column in $Columns) {
            $value = ""
            $property = $row.PSObject.Properties[$column]
            if ($property) {
                $value = $property.Value
            }
            [void]$html.Add(("<td>{0}</td>" -f (ConvertTo-HtmlText $value)))
        }
        [void]$html.Add("</tr>")
    }

    [void]$html.Add("</table>")
    return ($html -join "`r`n")
}

function Copy-RemoteEvidenceFolder {
    param(
        [Parameter(Mandatory=$true)][string]$RemoteDataPath,
        [Parameter(Mandatory=$true)][string]$DestinationPath,
        [Parameter(Mandatory=$true)][string]$ScriptName
    )

    $copyCount = 0

    function Copy-EvidenceFile {
        param(
            [Parameter(Mandatory=$true)][string]$SourceFile,
            [Parameter(Mandatory=$true)][string]$TargetFolder
        )

        if (-not (Test-Path -LiteralPath $TargetFolder)) {
            New-Item -ItemType Directory -Path $TargetFolder -Force | Out-Null
        }

        try {
            Copy-Item -LiteralPath $SourceFile -Destination $TargetFolder -Force -ErrorAction Stop
            return $true
        }
        catch [System.Management.Automation.ItemNotFoundException] {
            return $false
        }
        catch [System.IO.FileNotFoundException] {
            return $false
        }
        catch [System.IO.DirectoryNotFoundException] {
            return $false
        }
    }

    foreach ($folderName in @("Logs","Output","Transcripts")) {
        $sourceFolder = Join-Path $RemoteDataPath $folderName
        if (Test-Path -LiteralPath $sourceFolder) {
            $targetFolder = Join-Path $DestinationPath $folderName
            $files = @(Get-ChildItem -LiteralPath $sourceFolder -Recurse -File -Force -ErrorAction SilentlyContinue)
            foreach ($file in $files) {
                $relativePath = $file.FullName.Substring($sourceFolder.Length).TrimStart("\")
                $relativeDir = Split-Path -Parent $relativePath
                if ([string]::IsNullOrWhiteSpace($relativeDir)) {
                    $fileTargetFolder = $targetFolder
                }
                else {
                    $fileTargetFolder = Join-Path $targetFolder $relativeDir
                }

                if (Copy-EvidenceFile -SourceFile $file.FullName -TargetFolder $fileTargetFolder) {
                    $copyCount++
                }
            }
        }
    }

    Get-ChildItem -LiteralPath $RemoteDataPath -File -Force -ErrorAction SilentlyContinue |
        Where-Object {
            $_.Name -ne $ScriptName -and
            $_.Extension -in @(".csv",".log",".txt",".html",".json",".xml",".evtx")
        } |
        ForEach-Object {
            if (-not (Test-Path -LiteralPath $DestinationPath)) {
                New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
            }
            if (Copy-EvidenceFile -SourceFile $_.FullName -TargetFolder $DestinationPath) {
                $copyCount++
            }
        }

    if ($copyCount -eq 0) {
        if (Test-Path -LiteralPath $DestinationPath) {
            Remove-Item -LiteralPath $DestinationPath -Recurse -Force -ErrorAction SilentlyContinue
        }
        throw "No remote evidence files found to collect."
    }

    return $copyCount
}

function New-CycleHtmlReport {
    param(
        [Parameter(Mandatory=$true)][AllowEmptyCollection()][object[]]$Summary,
        [Parameter(Mandatory=$true)][string]$Path,
        [Parameter(Mandatory=$true)][int]$CycleNumber,
        [Parameter(Mandatory=$true)][datetime]$GeneratedAt
    )

    $rows = @($Summary | ForEach-Object { $_ })

    $statusCounts = @($rows | Group-Object -Property Status | Sort-Object Count -Descending | ForEach-Object {
        [PSCustomObject]@{ Status=$_.Name; Count=$_.Count }
    })
    $nextActionCounts = @($rows | Group-Object -Property NextAction | Sort-Object Count -Descending | ForEach-Object {
        [PSCustomObject]@{ NextAction=$_.Name; Count=$_.Count }
    })

    $style = @"
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 24px; color: #202020; }
h1 { font-size: 22px; margin-bottom: 4px; }
h2 { font-size: 16px; margin-top: 24px; }
table { border-collapse: collapse; width: 100%; margin-top: 10px; font-size: 12px; }
th { background: #1f2937; color: #fff; text-align: left; }
th, td { border: 1px solid #d0d7de; padding: 6px 8px; vertical-align: top; }
tr:nth-child(even) { background: #f6f8fa; }
.meta { color: #555; margin-bottom: 16px; }
</style>
"@

    $html = @()
    $html += "<html><head><meta charset='utf-8'>$style<title>Intune Hybrid Join repair cycle $CycleNumber</title></head><body>"
    $html += "<h1>Remote Intune Hybrid Join repair cycle $CycleNumber</h1>"
    $html += "<div class='meta'>Generated: $($GeneratedAt.ToString('yyyy-MM-dd HH:mm:ss')) | Computers: $($rows.Count) | Launcher: $LauncherVersion</div>"
    $html += "<h2>Status summary</h2>"
    $html += ConvertTo-SimpleHtmlTable -Rows $statusCounts -Columns @("Status","Count")
    $html += "<h2>Next action summary</h2>"
    $html += ConvertTo-SimpleHtmlTable -Rows $nextActionCounts -Columns @("NextAction","Count")
    $html += "<h2>Computer details</h2>"
    $detailColumns = @(
        "Cycle",
        "Timestamp",
        "Computer",
        "Status",
        "EffectiveStatus",
        "NextAction",
        "EffectiveNextAction",
        "RemoteStatus",
        "RemoteExitCode",
        "PsExecExitCode",
        "RemoteDetail",
        "InteractiveUserAccountName",
        "InteractiveUserAccountType",
        "InteractiveSessionName",
        "InteractiveSessionState",
        "UserIsUserAzureAD",
        "UserAzureAdPrt",
        "UserSessionIsNotRemote",
        "ErrorMessage",
        "IntuneInventoryPresent",
        "EntraInventoryPresent",
        "EntraRegisteredState",
        "EntraAlternativeSecurityIdCount",
        "EntraPendingReason",
        "EntraRegistrationDateTime",
        "EntraTrustType",
        "EntraDeviceId",
        "EntraObjectId",
        "ADInventoryPresent",
        "ADDomain",
        "ADEnabled",
        "ADDNSHostName",
        "ADDistinguishedName",
        "ADOperatingSystem",
        "ADLastLogonTimestampUtc",
        "PostCycleIntuneInventoryChecked",
        "PostCycleIntuneInventoryPresent",
        "PostCycleIntuneEnrollmentDetected",
        "PostCycleIntuneInventoryCsv",
        "PostCycleIntuneInventoryError",
        "PostCycleEntraInventoryChecked",
        "PostCycleEntraInventoryPresent",
        "PostCycleEntraRegisteredState",
        "PostCycleEntraAlternativeSecurityIdCount",
        "PostCycleEntraPendingResolved",
        "PostCycleEntraInventoryCsv",
        "PostCycleEntraInventoryError",
        "PostCycleADInventoryChecked",
        "PostCycleADInventoryPresent",
        "PostCycleADInventoryCsv",
        "PostCycleADInventoryError",
        "AdminShareReachable",
        "AdminShareFailureType",
        "PingReachable",
        "DnsResolved",
        "RemoteLogsCollected",
        "RemoteLogsPath",
        "RemoteCurrentRunLogsPath",
        "LogPath"
    )
    $html += ConvertTo-SimpleHtmlTable -Rows $rows -Columns $detailColumns
    $html += "</body></html>"
    $html -join "`r`n" | Out-File -LiteralPath $Path -Encoding UTF8 -Force
}

$scriptArgsBase = @()
if ($AllowDsregLeave) { $scriptArgsBase += "-AllowDsregLeave" }
if ($AllowRebootWhenNoInteractiveUser) { $scriptArgsBase += "-AllowRebootWhenNoInteractiveUser" }
if ($AllowRebootAfterDsregLeave) { $scriptArgsBase += "-AllowRebootAfterDsregLeave" }
if ($AllowRemoveNonIntuneMdmEnrollment) { $scriptArgsBase += "-AllowRemoveNonIntuneMdmEnrollment" }
if ($AllowRemoveStaleIntuneEnrollment) { $scriptArgsBase += "-AllowRemoveStaleIntuneEnrollment" }
if ($AuditOnly) { $scriptArgsBase += "-AuditOnly" }
$scriptArgsBase += "-StaleCleanupDelaySeconds"
$scriptArgsBase += $StaleCleanupDelaySeconds
$scriptArgsBase += "-RebootDelaySeconds"
$scriptArgsBase += $RebootDelaySeconds
$scriptArgsBase += "-IntuneRetrySleepMinutes"
$scriptArgsBase += $IntuneRetrySleepMinutes
$scriptArgsBase += "-IntuneRetryMaxRetries"
$scriptArgsBase += $IntuneRetryMaxRetries

$IntuneInventorySet = @{}
if (-not [string]::IsNullOrWhiteSpace($IntuneInventoryCsv)) {
    $refreshInitialInventory = $false
    $initialInventoryReason = ""
    $intuneInventoryItem = Get-Item -LiteralPath $IntuneInventoryCsv -ErrorAction SilentlyContinue

    if ($null -eq $intuneInventoryItem) {
        $refreshInitialInventory = $true
        $initialInventoryReason = "missing"
    }
    else {
        $intuneInventoryAge = (Get-Date) - $intuneInventoryItem.LastWriteTime
        if ($intuneInventoryAge.TotalMinutes -gt 60) {
            $refreshInitialInventory = $true
            $initialInventoryReason = ("older than 60 minutes; LastWriteTime={0}; Age={1:N1} minute(s)" -f $intuneInventoryItem.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"), $intuneInventoryAge.TotalMinutes)
        }
    }

    if ($refreshInitialInventory -and $DryRun) {
        Write-Host ("DryRun: Intune inventory CSV is {0}; skipping automatic Graph inventory export." -f $initialInventoryReason) -ForegroundColor Yellow
    }
    elseif ($refreshInitialInventory) {
        $initialInventoryLogPath = Join-Path $ReportRoot ("DevicesIntune_InitialRefresh_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
        Write-Host ("Intune inventory CSV is {0}. Running full Graph inventory export before starting the lot..." -f $initialInventoryReason) -ForegroundColor Yellow
        $initialInventory = Invoke-FullIntuneInventoryExport `
            -ExportScriptPath $ExportIntuneScriptPath `
            -OutputPath $IntuneInventoryCsv `
            -LogPath $initialInventoryLogPath `
            -PageSize $PostCycleIntuneInventoryPageSize

        if ($initialInventory.Success) {
            $IntuneInventorySet = $initialInventory.InventorySet
            Write-Host ("Initial Intune inventory refreshed. Devices={0}; CSV={1}" -f $IntuneInventorySet.Count,$initialInventory.CsvPath) -ForegroundColor Green
        }
        else {
            Write-Host ("WARNING: Initial Intune inventory refresh failed: {0}" -f $initialInventory.Error) -ForegroundColor Yellow
            if (Test-Path -LiteralPath $IntuneInventoryCsv) {
                Write-Host "Continuing with existing Intune CSV despite refresh failure." -ForegroundColor Yellow
                $IntuneInventorySet = Get-IntuneInventorySet -Path $IntuneInventoryCsv -NameColumn $IntuneInventoryNameColumn
            }
        }
    }
    else {
        $IntuneInventorySet = Get-IntuneInventorySet -Path $IntuneInventoryCsv -NameColumn $IntuneInventoryNameColumn
    }
}

$EntraInventoryMap = @{}
if (-not [string]::IsNullOrWhiteSpace($EntraInventoryCsv)) {
    try {
        $refreshInitialEntraInventory = $false
        $initialEntraInventoryReason = ""
        $entraInventoryItem = Get-Item -LiteralPath $EntraInventoryCsv -ErrorAction SilentlyContinue

        if ($null -eq $entraInventoryItem) {
            $refreshInitialEntraInventory = $true
            $initialEntraInventoryReason = "missing"
        }
        else {
            $entraInventoryAge = (Get-Date) - $entraInventoryItem.LastWriteTime
            if ($entraInventoryAge.TotalMinutes -gt 60) {
                $refreshInitialEntraInventory = $true
                $initialEntraInventoryReason = ("older than 60 minutes; LastWriteTime={0}; Age={1:N1} minute(s)" -f $entraInventoryItem.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"), $entraInventoryAge.TotalMinutes)
            }
        }

        if ($refreshInitialEntraInventory -and $DryRun) {
            Write-Host ("DryRun: Entra inventory CSV is {0}; skipping automatic Graph device export." -f $initialEntraInventoryReason) -ForegroundColor Yellow
        }
        elseif ($refreshInitialEntraInventory) {
            $initialEntraInventoryLogPath = Join-Path $ReportRoot ("DevicesEntra_InitialRefresh_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
            Write-Host ("Entra inventory CSV is {0}. Running full Graph device export before starting the lot..." -f $initialEntraInventoryReason) -ForegroundColor Yellow
            $initialEntraInventory = Invoke-FullEntraInventoryExport `
                -ExportScriptPath $ExportEntraScriptPath `
                -OutputPath $EntraInventoryCsv `
                -LogPath $initialEntraInventoryLogPath `
                -PageSize $PostCycleIntuneInventoryPageSize

            if ($initialEntraInventory.Success) {
                $EntraInventoryMap = $initialEntraInventory.InventoryMap
                Write-Host ("Initial Entra inventory refreshed. Devices={0}; CSV={1}" -f $EntraInventoryMap.Count,$initialEntraInventory.CsvPath) -ForegroundColor Green
            }
            else {
                Write-Host ("WARNING: Initial Entra inventory refresh failed: {0}" -f $initialEntraInventory.Error) -ForegroundColor Yellow
                if (Test-Path -LiteralPath $EntraInventoryCsv) {
                    Write-Host "Continuing with existing Entra CSV despite refresh failure." -ForegroundColor Yellow
                    $EntraInventoryMap = Get-EntraInventoryMap -Path $EntraInventoryCsv -NameColumn $EntraInventoryNameColumn
                }
            }
        }
        elseif (Test-Path -LiteralPath $EntraInventoryCsv) {
            $EntraInventoryMap = Get-EntraInventoryMap -Path $EntraInventoryCsv -NameColumn $EntraInventoryNameColumn
        }
    }
    catch {
        Write-Host ("WARN: Could not load Entra inventory CSV '{0}': {1}" -f $EntraInventoryCsv,$_.Exception.Message) -ForegroundColor Yellow
        $EntraInventoryMap = @{}
    }
}

$AdInventoryMap = @{}
if (-not [string]::IsNullOrWhiteSpace($AdInventoryCsv)) {
    try {
        $refreshInitialAdInventory = $false
        $initialAdInventoryReason = ""
        $adInventoryItem = Get-Item -LiteralPath $AdInventoryCsv -ErrorAction SilentlyContinue

        if ($null -eq $adInventoryItem) {
            $refreshInitialAdInventory = $true
            $initialAdInventoryReason = "missing"
        }
        else {
            $adInventoryAge = (Get-Date) - $adInventoryItem.LastWriteTime
            if ($adInventoryAge.TotalMinutes -gt 60) {
                $refreshInitialAdInventory = $true
                $initialAdInventoryReason = ("older than 60 minutes; LastWriteTime={0}; Age={1:N1} minute(s)" -f $adInventoryItem.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss"), $adInventoryAge.TotalMinutes)
            }
        }

        if ($AdInventoryUsesRecentRootCsv) {
            Write-Host ("AD forest inventory CSV is recent. Using root CSV in priority: {0}" -f $AdInventoryCsv) -ForegroundColor Green
            $AdInventoryMap = Get-AdInventoryMap -Path $AdInventoryCsv -NameColumn $AdInventoryNameColumn
        }
        elseif ($refreshInitialAdInventory -and $DryRun) {
            Write-Host ("DryRun: AD inventory CSV is {0}; skipping automatic AD computer export." -f $initialAdInventoryReason) -ForegroundColor Yellow
        }
        elseif ($refreshInitialAdInventory) {
            $initialAdInventoryLogPath = Join-Path $ReportRoot ("DevicesAD_InitialRefresh_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss"))
            $initialAdScope = if ([string]::IsNullOrWhiteSpace($AdDomain)) { "forest" } else { "domain '$AdDomain'" }
            Write-Host ("AD inventory CSV is {0}. Running AD computer export before starting the lot. Scope={1}..." -f $initialAdInventoryReason,$initialAdScope) -ForegroundColor Yellow
            $initialAdInventory = Invoke-FullAdInventoryExport `
                -ExportScriptPath $ExportAdScriptPath `
                -OutputPath $AdInventoryCsv `
                -LogPath $initialAdInventoryLogPath `
                -Domain $AdDomain

            if ($initialAdInventory.Success) {
                $AdInventoryMap = $initialAdInventory.InventoryMap
                Write-Host ("Initial AD inventory refreshed. Devices={0}; CSV={1}" -f $AdInventoryMap.Count,$initialAdInventory.CsvPath) -ForegroundColor Green
            }
            else {
                Write-Host ("WARNING: Initial AD inventory refresh failed: {0}" -f $initialAdInventory.Error) -ForegroundColor Yellow
                if (Test-Path -LiteralPath $AdInventoryCsv) {
                    Write-Host "Continuing with existing AD CSV despite refresh failure." -ForegroundColor Yellow
                    $AdInventoryMap = Get-AdInventoryMap -Path $AdInventoryCsv -NameColumn $AdInventoryNameColumn
                }
            }
        }
        elseif (Test-Path -LiteralPath $AdInventoryCsv) {
            $AdInventoryMap = Get-AdInventoryMap -Path $AdInventoryCsv -NameColumn $AdInventoryNameColumn
        }
    }
    catch {
        Write-Host ("WARN: Could not load AD inventory CSV '{0}': {1}" -f $AdInventoryCsv,$_.Exception.Message) -ForegroundColor Yellow
        $AdInventoryMap = @{}
    }
}

$localScriptVersionForDisplay = Get-ScriptVersionFromFile -Path $LocalScriptPath
$localScriptHashForDisplay = Get-FileSha256 -Path $LocalScriptPath

Write-Host "Remote repair launcher" -ForegroundColor Cyan
Write-Host "Launcher ver : $LauncherVersion"
Write-Host "Script       : $LocalScriptPath"
Write-Host "Script ver   : $localScriptVersionForDisplay"
Write-Host "Script hash  : $localScriptHashForDisplay"
Write-Host "Remote path  : $RemoteScriptPath"
Write-Host "Computers   : $ComputerListPath"
Write-Host "PsExec      : $PsExecPath"
Write-Host "Script args : $($scriptArgsBase -join ' ')"
Write-Host "Dry run     : $([bool]$DryRun)"
Write-Host "Audit only  : $([bool]$AuditOnly)"
Write-Host "Intune CSV  : $IntuneInventoryCsv"
Write-Host "Entra CSV   : $EntraInventoryCsv"
Write-Host "AD CSV      : $AdInventoryCsv"
Write-Host "AD domain   : $AdDomain"
Write-Host "AD root CSV : $AdRootInventoryCsv"
Write-Host "Ignore guard: $([bool]$IgnoreRunGuard); Every cycle: $([bool]$IgnoreRunGuardEveryCycle)"
Write-Host "Parallelism : ThrottleLimit=$ThrottleLimit; JobPollSeconds=$JobPollSeconds"
Write-Host "Start delay : $DelayBetweenComputersSeconds seconds between job starts"
Write-Host "Loop        : $(-not [bool]$RunOnce); Delay between cycles: $DelayBetweenCyclesMinutes minute(s); Max cycles: $MaxCycles"
Write-Host "Logs        : $LogRoot"
Write-Host "Reports     : $ReportRoot"
Write-Host "Central logs: Enabled=$CollectRemoteLogs; Path=$CentralLogRoot; History=$([bool]$KeepCentralLogHistory)"
Write-Host "Reboot delay: $RebootDelaySeconds seconds"
Write-Host "Stale delay : $StaleCleanupDelaySeconds seconds"
Write-Host "Intune wait : $IntuneRetryMaxRetries retry(ies) x $IntuneRetrySleepMinutes minute(s) = $($IntuneRetryMaxRetries * $IntuneRetrySleepMinutes) minute(s)"
Write-Host "PsExec wait : $(if ($PsExecTimeoutMinutes -eq 0) { 'No timeout' } else { "$PsExecTimeoutMinutes minute(s) max per computer" })"
Write-Host "Lost PsExec : poll every $CommunicationLostEvidencePollMinutes minute(s), max $CommunicationLostEvidenceWaitMinutes minute(s), before delayed evidence collection"
Write-Host "Post Intune : Enabled=$(-not [bool]$SkipPostCycleIntuneInventory); Mode=Full Graph inventory; PageSize=$PostCycleIntuneInventoryPageSize; Export=$ExportIntuneScriptPath"
Write-Host "Post Entra  : Enabled=$(-not [string]::IsNullOrWhiteSpace($EntraInventoryCsv)); Mode=Full Graph device inventory; PageSize=$PostCycleIntuneInventoryPageSize; Export=$ExportEntraScriptPath"
Write-Host "Post AD     : Enabled=$(-not $AdInventoryUsesRecentRootCsv -and -not [string]::IsNullOrWhiteSpace($AdInventoryCsv)); Mode=Full AD computer inventory; Export=$ExportAdScriptPath"
Write-Host ""

function Invoke-IntuneHybridJoinRepairCycle {
    param(
        [Parameter(Mandatory=$true)][int]$CycleNumber,
        [Parameter(Mandatory=$true)][string[]]$CycleScriptArgs
    )

    $computers = @(Get-ComputerList -Path $ComputerListPath)
    if (-not $computers -or $computers.Count -eq 0) {
        Write-Host "No computers found in $ComputerListPath." -ForegroundColor Yellow
        return $null
    }

    if (-not $DryRun -and $IntuneInventorySet -and $IntuneInventorySet.Count -gt 0) {
        $alreadyEnrolledFromInventory = @(
            $computers |
                Where-Object {
                    $inventoryKey = Get-ComputerListKey -ComputerName $_
                    -not [string]::IsNullOrWhiteSpace($inventoryKey) -and $IntuneInventorySet.ContainsKey($inventoryKey)
                } |
                ForEach-Object {
                    [PSCustomObject]@{
                        Computer = $_
                        IntuneInventoryPresent = $true
                        Status = "SKIPPED_ALREADY_INTUNE_IN_INVENTORY"
                        NextAction = "NO_ACTION_ALREADY_INTUNE"
                    }
                }
        )

        if ($alreadyEnrolledFromInventory.Count -gt 0) {
            try {
                $preMoveResult = Move-AlreadyEnrolledComputersFromList -ComputerListPath $ComputerListPath -CycleSummary $alreadyEnrolledFromInventory
                if ($preMoveResult.Moved -gt 0) {
                    Write-Host ("Cycle {0}: pre-filtered {1} already-enrolled computer(s) from DevicesIntune.csv to {2}. PsExec will be skipped for them." -f $CycleNumber,$preMoveResult.Moved,$preMoveResult.AlreadyEnrolledPath) -ForegroundColor Green
                    $computers = @(Get-ComputerList -Path $ComputerListPath)
                    if (-not $computers -or $computers.Count -eq 0) {
                        Write-Host ("Cycle {0}: all computers were already present in Intune inventory. Nothing left to run." -f $CycleNumber) -ForegroundColor Green
                        return $null
                    }
                }
            }
            catch {
                Write-Host ("Cycle {0}: failed to pre-filter already-enrolled computers from DevicesIntune.csv: {1}" -f $CycleNumber,$_.Exception.Message) -ForegroundColor Yellow
            }
        }
    }

    $summary = New-Object System.Collections.Generic.List[object]
    $reportColumns = @(Get-LauncherReportColumns)
    $liveSummaryPath = Join-Path $ReportRoot ("PsExec_IntuneHybridJoinRepair_Live_cycle{0}_{1}.csv" -f $CycleNumber,(Get-Date -Format "yyyyMMdd_HHmmss"))
    $liveHtmlPath = [System.IO.Path]::ChangeExtension($liveSummaryPath, ".html")
    Initialize-LiveCycleReport -Path $liveSummaryPath -Columns $reportColumns
    New-CycleHtmlReport -Summary @() -Path $liveHtmlPath -CycleNumber $CycleNumber -GeneratedAt (Get-Date)
    Write-Host ("Cycle {0} started. Computers={1}; Throttle={2}; Args={3}" -f $CycleNumber,$computers.Count,$ThrottleLimit,($CycleScriptArgs -join ' ')) -ForegroundColor Cyan
    Write-Host ("Cycle {0} live report: {1}" -f $CycleNumber,$liveSummaryPath) -ForegroundColor Green
    Write-Host ("Cycle {0} live HTML  : {1}" -f $CycleNumber,$liveHtmlPath) -ForegroundColor Green

    $worker = {
        param(
            [string]$Computer,
            [int]$CycleNumber,
            [string]$LocalScriptPath,
            [string]$ScriptName,
            [string]$RemoteRelativeDir,
            [string]$RemoteScriptPath,
            [string]$RemoteDataRelativeDir,
            [string]$PsExecPath,
            [string]$LogRoot,
            [string]$LauncherVersion,
            [bool]$DryRun,
            [bool]$CollectRemoteLogs,
            [string]$CentralLogRoot,
            [bool]$KeepCentralLogHistory,
            [hashtable]$IntuneInventorySet,
            [hashtable]$EntraInventoryMap,
            [hashtable]$AdInventoryMap,
            [string[]]$CycleScriptArgs,
            [int]$PsExecTimeoutMinutes,
            [int]$CommunicationLostEvidenceWaitMinutes,
            [int]$CommunicationLostEvidencePollMinutes
        )

        $ErrorActionPreference = "Stop"

        function Update-TimestampedLogFile {
            param([Parameter(Mandatory=$true)][string]$Path)

            if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
                return
            }

            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $updatedLines = [System.IO.File]::ReadAllLines($Path) | ForEach-Object {
                if ($_ -match '^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\b') {
                    $_
                }
                elseif ($_ -match '^\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]\s*(.*)$') {
                    "{0} {1}" -f $Matches[1], $Matches[2]
                }
                else {
                    "{0} {1}" -f $timestamp, $_
                }
            }

            $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
            [System.IO.File]::WriteAllLines($Path, [string[]]$updatedLines, $utf8NoBom)
        }

        function Complete-WorkerResult {
            param(
                [Parameter(Mandatory=$true)]$Result,
                [Parameter(Mandatory=$true)][string]$Path
            )

            try { Update-TimestampedLogFile -Path $Path } catch { }
            return [PSCustomObject]$Result
        }

        function Get-ScriptVersionFromFile {
            param([Parameter(Mandatory=$true)][string]$Path)

            try {
                if (-not (Test-Path -LiteralPath $Path)) { return "" }
                $match = Select-String -LiteralPath $Path -Pattern '^\s*\$ScriptVersion\s*=\s*"([^"]+)"' -ErrorAction Stop | Select-Object -First 1
                if ($match -and $match.Matches.Count -gt 0) {
                    return $match.Matches[0].Groups[1].Value
                }
            }
            catch { }

            return ""
        }

        function Get-FileSha256 {
            param([Parameter(Mandatory=$true)][string]$Path)

            try {
                if (-not (Test-Path -LiteralPath $Path)) { return "" }
                return (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash
            }
            catch {
                return ""
            }
        }

        function Get-NextActionFromLauncherStatus {
            param([Parameter(Mandatory=$true)][string]$Status)

            switch ($Status) {
                "SUCCESS" { return "NO_ACTION_ALREADY_INTUNE_OR_HEALTHY" }
                "AUDIT_SUCCESS_ALREADY_INTUNE" { return "NO_ACTION_ALREADY_INTUNE" }
                "AUDIT_INTUNE_MISSING" { return "RUN_REPAIR" }
                "AUDIT_STALE_INTUNE_ENROLLMENT_LOCAL" { return "CLEAN_STALE_INTUNE_OPTIN" }
                "INTUNE_ENROLLMENT_PENDING_CONFIRMATION" { return "RECHECK_LATER_INTUNE_ENROLLMENT" }
                "ADMIN_SHARE_UNREACHABLE" { return "FIX_ADMIN_SHARE_OR_NETWORK" }
                "RUN_GUARD_ACTIVE" { return "WAIT_RUN_GUARD" }
                "REBOOT_TRIGGERED_WAITING_FOR_USER_LOGON" { return "WAIT_USER_LOGON" }
                "WAITING_FOR_INTERACTIVE_USER_LOGON" { return "WAIT_USER_LOGON" }
                "INTUNE_USER_AUTOENROLL_LOCAL_INTERACTIVE_USER" { return "LOGON_WITH_DOMAIN_OR_AAD_USER" }
                "INTUNE_USER_AUTOENROLL_TASK_NOT_FOUND" { return "FIX_GPO_USER_AUTOENROLL_TASK" }
                "STALE_INTUNE_ENROLLMENT_LOCAL" { return "CLEAN_STALE_INTUNE_OPTIN" }
                "NON_INTUNE_MDM_ENROLLED" { return "CLEAN_NON_INTUNE_MDM_OPTIN" }
                "ENTRA_HYBRID_PENDING_ADJ_TRIGGERED" { return "RECHECK_ENTRA_PENDING_AFTER_ADJ" }
                "ENTRA_HYBRID_PENDING_RETRY_EXHAUSTED" { return "CHECK_AD_CONNECT_OR_DUPLICATE_ENTRA_DEVICE" }
                "ENTRA_PENDING_RESOLVED_POST_CYCLE" { return "RECHECK_INTUNE_ENROLLMENT" }
                "USER_NOT_AZUREAD" { return "CHECK_USER_AAD_OR_LOGON_CONTEXT" }
                "USER_PRT_NOT_AVAILABLE" { return "CHECK_USER_PRT" }
                "USER_PRT_REFRESH_FAILED" { return "FIX_USER_PRT_OR_RELOGIN" }
                "USER_SESSION_REMOTE" { return "LOGON_ON_CONSOLE" }
                "INTUNE_AUTOENROLL_POLICY_NOT_CONFIGURED" { return "CHECK_GPO_AUTOENROLL" }
                    "KEY_SIGN_TEST_FAILED" { return "REPAIR_HYBRID_JOIN_KEY_OR_ALLOW_LEAVE" }
                    "INTUNE_ENROLLMENT_CONNECTIVITY_FAILED" { return "CHECK_CONNECTIVITY" }
                    "DOMAIN_CONTROLLER_UNREACHABLE" { return "FIX_DOMAIN_CONNECTIVITY_OR_VPN" }
                    "PSEXEC_TIMEOUT" { return "CHECK_REMOTE_LOG_OR_RETRY" }
                    "PSEXEC_COMMUNICATION_LOST" { return "RETRY_PSEXEC_OR_CHECK_REMOTE_SERVICE" }
                    "PSEXEC_EXIT_UNKNOWN" { return "CHECK_CURRENT_RUN_REMOTE_LOG" }
                    "REMOTE_DIRECTORY_CREATE_FAILED" { return "FIX_SCRIPT_COPY_OR_ADMIN_SHARE" }
                    "REMOTE_SCRIPT_COPY_FAILED" { return "FIX_SCRIPT_COPY_OR_SECURITY" }
                    "REMOTE_SCRIPT_MISSING" { return "FIX_SCRIPT_COPY_OR_SECURITY" }
                    default {
                    if ($Status -like "ERROR*") { return "CHECK_CONNECTIVITY_OR_ADMIN_ACCESS" }
                    if ($Status -like "PSEXEC_EXIT*") { return "CHECK_CURRENT_RUN_REMOTE_LOG" }
                    return "REVIEW_LOGS"
                }
            }
        }

        function Copy-RemoteEvidenceFolder {
            param(
                [Parameter(Mandatory=$true)][string]$RemoteDataPath,
                [Parameter(Mandatory=$true)][string]$DestinationPath,
                [Parameter(Mandatory=$true)][string]$ScriptName
            )

            $copyCount = 0

            function Copy-EvidenceFile {
                param(
                    [Parameter(Mandatory=$true)][string]$SourceFile,
                    [Parameter(Mandatory=$true)][string]$TargetFolder
                )

                if (-not (Test-Path -LiteralPath $TargetFolder)) {
                    New-Item -ItemType Directory -Path $TargetFolder -Force | Out-Null
                }

                try {
                    Copy-Item -LiteralPath $SourceFile -Destination $TargetFolder -Force -ErrorAction Stop
                    return $true
                }
                catch [System.Management.Automation.ItemNotFoundException] {
                    return $false
                }
                catch [System.IO.FileNotFoundException] {
                    return $false
                }
                catch [System.IO.DirectoryNotFoundException] {
                    return $false
                }
            }

            foreach ($folderName in @("Logs","Output","Transcripts")) {
                $sourceFolder = Join-Path $RemoteDataPath $folderName
                if (Test-Path -LiteralPath $sourceFolder) {
                    $targetFolder = Join-Path $DestinationPath $folderName
                    $files = @(Get-ChildItem -LiteralPath $sourceFolder -Recurse -File -Force -ErrorAction SilentlyContinue)
                    foreach ($file in $files) {
                        $relativePath = $file.FullName.Substring($sourceFolder.Length).TrimStart("\")
                        $relativeDir = Split-Path -Parent $relativePath
                        if ([string]::IsNullOrWhiteSpace($relativeDir)) {
                            $fileTargetFolder = $targetFolder
                        }
                        else {
                            $fileTargetFolder = Join-Path $targetFolder $relativeDir
                        }

                        if (Copy-EvidenceFile -SourceFile $file.FullName -TargetFolder $fileTargetFolder) {
                            $copyCount++
                        }
                    }
                }
            }

            Get-ChildItem -LiteralPath $RemoteDataPath -File -Force -ErrorAction SilentlyContinue |
                Where-Object {
                    $_.Name -ne $ScriptName -and
                    $_.Extension -in @(".csv",".log",".txt",".html",".json",".xml",".evtx")
                } |
                ForEach-Object {
                    if (-not (Test-Path -LiteralPath $DestinationPath)) {
                        New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null
                    }
                    if (Copy-EvidenceFile -SourceFile $_.FullName -TargetFolder $DestinationPath) {
                        $copyCount++
                    }
                }

            if ($copyCount -eq 0) {
                if (Test-Path -LiteralPath $DestinationPath) {
                    Remove-Item -LiteralPath $DestinationPath -Recurse -Force -ErrorAction SilentlyContinue
                }
                throw "No remote evidence files found to collect."
            }

            return $copyCount
        }

        function Get-RemoteEvidenceFinalStatus {
            param(
                [Parameter(Mandatory=$true)][string]$EvidencePath,
                [Parameter(Mandatory=$false)][datetime]$Since = [datetime]::MinValue,
                [switch]$RequireCompletedRun
            )

            $completedRunId = ""
            $completedRunStatus = ""
            $completedRunExitCode = ""
            $completedRunNextAction = ""
            $completedRunDetail = ""
            $lastRunPath = Join-Path $EvidencePath "LastRun.json"
            if ($RequireCompletedRun) {
                if (-not (Test-Path -LiteralPath $lastRunPath -ErrorAction SilentlyContinue)) { return $null }
                try {
                    $lastRun = Get-Content -LiteralPath $lastRunPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                    if ($null -eq $lastRun -or [string]::IsNullOrWhiteSpace([string]$lastRun.RunId) -or [string]::IsNullOrWhiteSpace([string]$lastRun.EndTime)) {
                        return $null
                    }
                    $lastRunStart = [datetime]$lastRun.StartTime
                    if ($lastRunStart -lt $Since.AddSeconds(-5)) { return $null }
                    $completedRunId = ([string]$lastRun.RunId).Trim()
                    $completedRunStatus = ([string]$lastRun.Status).Trim()
                    $completedRunExitCode = ([string]$lastRun.ExitCode).Trim()
                    $completedRunNextAction = ([string]$lastRun.NextAction).Trim()
                    $completedRunDetail = ([string]$lastRun.Detail).Trim()
                    if ([string]::IsNullOrWhiteSpace($completedRunStatus) -or [string]::IsNullOrWhiteSpace($completedRunExitCode)) {
                        return $null
                    }
                }
                catch {
                    return $null
                }
            }

            $csv = Get-ChildItem -LiteralPath $EvidencePath -Recurse -File -Filter "IntuneHybridJoinToolkit_*.csv" -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -ge $Since.AddSeconds(-5) } |
                Sort-Object LastWriteTime -Descending |
                Select-Object -First 1

            if (-not $csv) {
                if ($RequireCompletedRun -and -not [string]::IsNullOrWhiteSpace($completedRunStatus)) {
                    return [PSCustomObject]@{
                        RunId = $completedRunId
                        Status = $completedRunStatus
                        ExitCode = $completedRunExitCode
                        NextAction = $completedRunNextAction
                        Detail = $completedRunDetail
                        CsvPath = ""
                        LastRunPath = $lastRunPath
                        LastRunStatus = $completedRunStatus
                        LastRunExitCode = $completedRunExitCode
                        InteractiveUserName = ""
                        InteractiveUserDomain = ""
                        InteractiveUserAccountName = ""
                        InteractiveUserAccountType = ""
                        InteractiveSessionName = ""
                        InteractiveSessionState = ""
                        UserIsUserAzureAD = ""
                        UserAzureAdPrt = ""
                        UserSessionIsNotRemote = ""
                    }
                }
                return $null
            }

            try {
                $row = Import-Csv -LiteralPath $csv.FullName -ErrorAction Stop | Select-Object -Last 1
                if ($null -eq $row -or [string]::IsNullOrWhiteSpace([string]$row.Status)) { return $null }

                $rowRunId = ""
                if ($row.PSObject.Properties["RunId"]) { $rowRunId = ([string]$row.RunId).Trim() }
                if ($RequireCompletedRun) {
                    if ([string]::IsNullOrWhiteSpace($rowRunId) -or $rowRunId -ne $completedRunId) { return $null }
                }

                $exitCode = ""
                if ($row.PSObject.Properties["ExitCode"]) { $exitCode = [string]$row.ExitCode }

                $nextAction = ""
                if ($row.PSObject.Properties["NextAction"]) { $nextAction = [string]$row.NextAction }
                if ([string]::IsNullOrWhiteSpace($nextAction) -and $RequireCompletedRun) { $nextAction = $completedRunNextAction }

                $detail = ""
                if ($row.PSObject.Properties["DsregStatusErrorMessage"]) { $detail = [string]$row.DsregStatusErrorMessage }
                if ([string]::IsNullOrWhiteSpace($detail) -and $row.PSObject.Properties["ErrorMessage"]) { $detail = [string]$row.ErrorMessage }
                if ([string]::IsNullOrWhiteSpace($detail) -and $RequireCompletedRun) { $detail = $completedRunDetail }

                $interactiveUserName = ""; if ($row.PSObject.Properties["InteractiveUserName"]) { $interactiveUserName = [string]$row.InteractiveUserName }
                $interactiveUserDomain = ""; if ($row.PSObject.Properties["InteractiveUserDomain"]) { $interactiveUserDomain = [string]$row.InteractiveUserDomain }
                $interactiveUserAccountName = ""; if ($row.PSObject.Properties["InteractiveUserAccountName"]) { $interactiveUserAccountName = [string]$row.InteractiveUserAccountName }
                $interactiveUserAccountType = ""; if ($row.PSObject.Properties["InteractiveUserAccountType"]) { $interactiveUserAccountType = [string]$row.InteractiveUserAccountType }
                $interactiveSessionName = ""; if ($row.PSObject.Properties["InteractiveSessionName"]) { $interactiveSessionName = [string]$row.InteractiveSessionName }
                $interactiveSessionState = ""; if ($row.PSObject.Properties["InteractiveSessionState"]) { $interactiveSessionState = [string]$row.InteractiveSessionState }
                $userIsUserAzureAD = ""; if ($row.PSObject.Properties["User_IsUserAzureAD"]) { $userIsUserAzureAD = [string]$row.User_IsUserAzureAD }
                $userAzureAdPrt = ""; if ($row.PSObject.Properties["User_AzureAdPrt"]) { $userAzureAdPrt = [string]$row.User_AzureAdPrt }
                $userSessionIsNotRemote = ""; if ($row.PSObject.Properties["User_SessionIsNotRemote"]) { $userSessionIsNotRemote = [string]$row.User_SessionIsNotRemote }

                return [PSCustomObject]@{
                    RunId = $rowRunId
                    Status = $(if ($RequireCompletedRun) { $completedRunStatus } else { ([string]$row.Status).Trim() })
                    ExitCode = $(if ($RequireCompletedRun) { $completedRunExitCode } else { $exitCode.Trim() })
                    NextAction = $nextAction.Trim()
                    Detail = $detail.Trim()
                    CsvPath = $csv.FullName
                    LastRunPath = $(if ($RequireCompletedRun) { $lastRunPath } else { "" })
                    LastRunStatus = $completedRunStatus
                    LastRunExitCode = $completedRunExitCode
                    InteractiveUserName = $interactiveUserName.Trim()
                    InteractiveUserDomain = $interactiveUserDomain.Trim()
                    InteractiveUserAccountName = $interactiveUserAccountName.Trim()
                    InteractiveUserAccountType = $interactiveUserAccountType.Trim()
                    InteractiveSessionName = $interactiveSessionName.Trim()
                    InteractiveSessionState = $interactiveSessionState.Trim()
                    UserIsUserAzureAD = $userIsUserAzureAD.Trim()
                    UserAzureAdPrt = $userAzureAdPrt.Trim()
                    UserSessionIsNotRemote = $userSessionIsNotRemote.Trim()
                }
            }
            catch {
                return [PSCustomObject]@{
                    Status = ""
                    ExitCode = ""
                    NextAction = ""
                    Detail = ("Could not parse remote evidence CSV: {0}" -f $_.Exception.Message)
                    CsvPath = $csv.FullName
                }
            }
        }

        $runId = Get-Date -Format "yyyyMMdd_HHmmss"
        $logPath = Join-Path $LogRoot ("{0}_cycle{1}_{2}.log" -f $Computer, $CycleNumber, $runId)
        $remoteRootShare = "\\$Computer\C$"
        $remoteAdminDir = Join-Path $remoteRootShare $RemoteRelativeDir
        $remoteAdminScript = Join-Path $remoteAdminDir $ScriptName
        $remoteDataAdminDir = Join-Path $remoteRootShare $RemoteDataRelativeDir
        $safeComputerName = $Computer -replace '[\\/:*?"<>|]', '_'
        $centralComputerDir = Join-Path $CentralLogRoot $safeComputerName
        $centralRunDir = if ($KeepCentralLogHistory) {
            Join-Path $centralComputerDir ("cycle{0}_{1}" -f $CycleNumber,$runId)
        }
        else {
            Join-Path $centralComputerDir "Latest"
        }

        $result = [ordered]@{
            LauncherVersion = $LauncherVersion
            Cycle = $CycleNumber
            Computer = $Computer
            Timestamp = Get-Date
            DryRun = $DryRun
            DnsResolved = $false
            DnsAddressList = ""
            AdminShareReachable = $false
            PingReachable = $false
            RemoteDirectoryCreated = $false
            ScriptCopied = $false
            LocalScriptVersion = ""
            RemoteScriptVersion = ""
            LocalScriptHash = ""
            RemoteScriptHash = ""
            PsExecExitCode = ""
            RemoteStatus = ""
            RemoteExitCode = ""
            RemoteNextAction = ""
            RemoteDetail = ""
            NextAction = ""
            EffectiveStatus = ""
            EffectiveNextAction = ""
            InteractiveUserName = ""
            InteractiveUserDomain = ""
            InteractiveUserAccountName = ""
            InteractiveUserAccountType = ""
            InteractiveSessionName = ""
            InteractiveSessionState = ""
            UserIsUserAzureAD = ""
            UserAzureAdPrt = ""
            UserSessionIsNotRemote = ""
            IntuneInventoryPresent = ""
            EntraInventoryPresent = ""
            EntraRegisteredState = ""
            EntraAlternativeSecurityIdCount = ""
            EntraPendingReason = ""
            EntraRegistrationDateTime = ""
            EntraTrustType = ""
            EntraDeviceId = ""
            EntraObjectId = ""
            ADInventoryPresent = ""
            ADDomain = ""
            ADEnabled = ""
            ADDNSHostName = ""
            ADDistinguishedName = ""
            ADOperatingSystem = ""
            ADLastLogonTimestampUtc = ""
            AdminShareFailureType = ""
            PostCycleIntuneInventoryChecked = ""
            PostCycleIntuneInventoryPresent = ""
            PostCycleIntuneEnrollmentDetected = ""
            PostCycleIntuneInventoryCsv = ""
            PostCycleIntuneInventoryError = ""
            PostCycleEntraInventoryChecked = ""
            PostCycleEntraInventoryPresent = ""
            PostCycleEntraRegisteredState = ""
            PostCycleEntraAlternativeSecurityIdCount = ""
            PostCycleEntraPendingResolved = ""
            PostCycleEntraInventoryCsv = ""
            PostCycleEntraInventoryError = ""
            PostCycleADInventoryChecked = ""
            PostCycleADInventoryPresent = ""
            PostCycleADInventoryCsv = ""
            PostCycleADInventoryError = ""
            RemoteLogsCollected = $false
            RemoteLogsPath = ""
            RemoteCurrentRunLogsPath = ""
            RemoteLogsError = ""
            Status = "STARTED"
            LogPath = $logPath
            ErrorMessage = ""
        }

        try {
            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Starting $Computer. Cycle=$CycleNumber" | Set-Content -LiteralPath $logPath -Encoding UTF8
            $inventoryKey = ($Computer.Split(".")[0]).ToUpperInvariant()
            if ($IntuneInventorySet -and $IntuneInventorySet.Count -gt 0) {
                $result.IntuneInventoryPresent = [bool]$IntuneInventorySet.ContainsKey($inventoryKey)
            }
            if ($EntraInventoryMap -and $EntraInventoryMap.Count -gt 0 -and $EntraInventoryMap.ContainsKey($inventoryKey)) {
                $entraRow = $EntraInventoryMap[$inventoryKey]
                $result.EntraInventoryPresent = $true
                if ($entraRow.PSObject.Properties["EntraRegisteredState"]) { $result.EntraRegisteredState = [string]$entraRow.EntraRegisteredState }
                if ($entraRow.PSObject.Properties["AlternativeSecurityIdCount"]) { $result.EntraAlternativeSecurityIdCount = [string]$entraRow.AlternativeSecurityIdCount }
                if ($entraRow.PSObject.Properties["EntraPendingReason"]) { $result.EntraPendingReason = [string]$entraRow.EntraPendingReason }
                if ($entraRow.PSObject.Properties["RegistrationDateTime"]) { $result.EntraRegistrationDateTime = [string]$entraRow.RegistrationDateTime }
                if ($entraRow.PSObject.Properties["TrustType"]) { $result.EntraTrustType = [string]$entraRow.TrustType }
                if ($entraRow.PSObject.Properties["DeviceId"]) { $result.EntraDeviceId = [string]$entraRow.DeviceId }
                if ($entraRow.PSObject.Properties["EntraObjectId"]) { $result.EntraObjectId = [string]$entraRow.EntraObjectId }
            }
            elseif ($EntraInventoryMap -and $EntraInventoryMap.Count -gt 0) {
                $result.EntraInventoryPresent = $false
            }
            if ($AdInventoryMap -and $AdInventoryMap.Count -gt 0 -and $AdInventoryMap.ContainsKey($inventoryKey)) {
                $adRow = $AdInventoryMap[$inventoryKey]
                $result.ADInventoryPresent = $true
                if ($adRow.PSObject.Properties["ADDomain"]) { $result.ADDomain = [string]$adRow.ADDomain }
                if ($adRow.PSObject.Properties["Enabled"]) { $result.ADEnabled = [string]$adRow.Enabled }
                if ($adRow.PSObject.Properties["DNSHostName"]) { $result.ADDNSHostName = [string]$adRow.DNSHostName }
                if ($adRow.PSObject.Properties["DistinguishedName"]) { $result.ADDistinguishedName = [string]$adRow.DistinguishedName }
                if ($adRow.PSObject.Properties["OperatingSystem"]) { $result.ADOperatingSystem = [string]$adRow.OperatingSystem }
                if ($adRow.PSObject.Properties["LastLogonTimestampUtc"]) { $result.ADLastLogonTimestampUtc = [string]$adRow.LastLogonTimestampUtc }
            }
            elseif ($AdInventoryMap -and $AdInventoryMap.Count -gt 0) {
                $result.ADInventoryPresent = $false
            }

            try {
                $dns = [System.Net.Dns]::GetHostEntry($Computer)
                $result.DnsResolved = $true
                $result.DnsAddressList = (($dns.AddressList | ForEach-Object { $_.IPAddressToString }) -join ";")
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] DNS resolved: $($result.DnsAddressList)" | Add-Content -LiteralPath $logPath -Encoding UTF8
            }
            catch {
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] WARN: DNS resolution failed: $($_.Exception.Message)" | Add-Content -LiteralPath $logPath -Encoding UTF8
            }

            $result.PingReachable = Test-Connection -ComputerName $Computer -Count 1 -Quiet -ErrorAction SilentlyContinue
            if (-not $result.PingReachable) {
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] WARN: Ping failed. Trying administrative shares anyway." | Add-Content -LiteralPath $logPath -Encoding UTF8
            }

            $adminShare = "\\$Computer\ADMIN$"
            $rootShare = "\\$Computer\C$"
            $adminShareReachable = Test-Path -LiteralPath $adminShare -ErrorAction SilentlyContinue
            $rootShareReachable = Test-Path -LiteralPath $rootShare -ErrorAction SilentlyContinue
            $result.AdminShareReachable = ($adminShareReachable -and $rootShareReachable)
            if ($result.AdminShareReachable) {
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Administrative shares reachable. ADMIN$=$adminShareReachable; C$=$rootShareReachable" | Add-Content -LiteralPath $logPath -Encoding UTF8
            }
            else {
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] WARN: Required administrative share is not reachable. ADMIN$=$adminShareReachable; C$=$rootShareReachable; ADMINPath=$adminShare; RootPath=$rootShare" | Add-Content -LiteralPath $logPath -Encoding UTF8
            }
            if (-not $result.AdminShareReachable) {
                if (-not $result.DnsResolved) {
                    $result.AdminShareFailureType = "DNS_FAILED"
                }
                elseif (-not $result.PingReachable) {
                    $result.AdminShareFailureType = "PING_FAILED_ADMIN_SHARE_FAILED"
                }
                else {
                    $result.AdminShareFailureType = "PING_OK_ADMIN_SHARE_FAILED"
                }
            }

            if ($DryRun) {
                if ($result.AdminShareReachable) {
                    $result.Status = "DRYRUN_READY"
                    $result.NextAction = "READY_FOR_REPAIR"
                    $result.RemoteDetail = "DNS/Ping/administrative-share pre-check completed. No script copied or executed."
                }
                else {
                    $result.Status = "DRYRUN_ADMIN_SHARE_UNREACHABLE"
                    $result.NextAction = "FIX_ADMIN_SHARE_OR_NETWORK"
                    $result.RemoteDetail = "Required administrative share is not reachable. PsExec/copy would probably fail."
                }
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] DryRun status: $($result.Status); Detail=$($result.RemoteDetail)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                return (Complete-WorkerResult -Result $result -Path $logPath)
            }

            if (-not $result.AdminShareReachable) {
                $result.Status = "ADMIN_SHARE_UNREACHABLE"
                $result.NextAction = "FIX_ADMIN_SHARE_OR_NETWORK"
                $result.RemoteDetail = ("{0}: Required administrative share is not reachable. Script copy and PsExec execution were skipped." -f $result.AdminShareFailureType)
                $result.ErrorMessage = $result.RemoteDetail
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Skipping repair: $($result.RemoteDetail)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                return (Complete-WorkerResult -Result $result -Path $logPath)
            }

            if (-not (Test-Path -LiteralPath $remoteAdminDir)) {
                New-Item -ItemType Directory -Path $remoteAdminDir -Force -ErrorAction Stop | Out-Null
            }
            $result.RemoteDirectoryCreated = Test-Path -LiteralPath $remoteAdminDir
            if (-not $result.RemoteDirectoryCreated) {
                $result.Status = "REMOTE_DIRECTORY_CREATE_FAILED"
                $result.NextAction = Get-NextActionFromLauncherStatus -Status $result.Status
                $result.RemoteDetail = "Remote repair folder could not be created or verified: $remoteAdminDir"
                $result.ErrorMessage = $result.RemoteDetail
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERROR: $($result.RemoteDetail)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                return (Complete-WorkerResult -Result $result -Path $logPath)
            }
            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Remote repair folder ready: $remoteAdminDir" | Add-Content -LiteralPath $logPath -Encoding UTF8

            $result.LocalScriptVersion = Get-ScriptVersionFromFile -Path $LocalScriptPath
            $result.LocalScriptHash = Get-FileSha256 -Path $LocalScriptPath
            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Local script before copy: Version=$($result.LocalScriptVersion); SHA256=$($result.LocalScriptHash); Path=$LocalScriptPath" | Add-Content -LiteralPath $logPath -Encoding UTF8

            Copy-Item -LiteralPath $LocalScriptPath -Destination $remoteAdminScript -Force -ErrorAction Stop
            $result.ScriptCopied = Test-Path -LiteralPath $remoteAdminScript
            if ($result.ScriptCopied) {
                $localScriptItem = Get-Item -LiteralPath $LocalScriptPath -ErrorAction Stop
                $remoteScriptItem = Get-Item -LiteralPath $remoteAdminScript -ErrorAction Stop
                $result.RemoteScriptVersion = Get-ScriptVersionFromFile -Path $remoteAdminScript
                $result.RemoteScriptHash = Get-FileSha256 -Path $remoteAdminScript
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Remote script after copy: Version=$($result.RemoteScriptVersion); SHA256=$($result.RemoteScriptHash); Bytes=$($remoteScriptItem.Length); Path=$remoteAdminScript" | Add-Content -LiteralPath $logPath -Encoding UTF8
                if ($remoteScriptItem.Length -ne $localScriptItem.Length) {
                    $result.ScriptCopied = $false
                    $result.Status = "REMOTE_SCRIPT_COPY_FAILED"
                    $result.NextAction = Get-NextActionFromLauncherStatus -Status $result.Status
                    $result.RemoteDetail = "Remote script copy size mismatch. LocalBytes=$($localScriptItem.Length); RemoteBytes=$($remoteScriptItem.Length); RemotePath=$remoteAdminScript"
                    $result.ErrorMessage = $result.RemoteDetail
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERROR: $($result.RemoteDetail)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                    return (Complete-WorkerResult -Result $result -Path $logPath)
                }
                if ([string]::IsNullOrWhiteSpace($result.LocalScriptHash) -or [string]::IsNullOrWhiteSpace($result.RemoteScriptHash) -or $result.RemoteScriptHash -ne $result.LocalScriptHash) {
                    $result.ScriptCopied = $false
                    $result.Status = "REMOTE_SCRIPT_COPY_FAILED"
                    $result.NextAction = Get-NextActionFromLauncherStatus -Status $result.Status
                    $result.RemoteDetail = "Remote script copy hash mismatch. LocalVersion=$($result.LocalScriptVersion); RemoteVersion=$($result.RemoteScriptVersion); LocalSHA256=$($result.LocalScriptHash); RemoteSHA256=$($result.RemoteScriptHash); RemotePath=$remoteAdminScript"
                    $result.ErrorMessage = $result.RemoteDetail
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERROR: $($result.RemoteDetail)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                    return (Complete-WorkerResult -Result $result -Path $logPath)
                }
                if ((-not [string]::IsNullOrWhiteSpace($result.LocalScriptVersion)) -and $result.RemoteScriptVersion -ne $result.LocalScriptVersion) {
                    $result.ScriptCopied = $false
                    $result.Status = "REMOTE_SCRIPT_COPY_FAILED"
                    $result.NextAction = Get-NextActionFromLauncherStatus -Status $result.Status
                    $result.RemoteDetail = "Remote script copy version mismatch. LocalVersion=$($result.LocalScriptVersion); RemoteVersion=$($result.RemoteScriptVersion); RemotePath=$remoteAdminScript"
                    $result.ErrorMessage = $result.RemoteDetail
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERROR: $($result.RemoteDetail)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                    return (Complete-WorkerResult -Result $result -Path $logPath)
                }
            }
            else {
                $result.Status = "REMOTE_SCRIPT_COPY_FAILED"
                $result.NextAction = Get-NextActionFromLauncherStatus -Status $result.Status
                $result.RemoteDetail = "Remote script copy could not be verified: $remoteAdminScript"
                $result.ErrorMessage = $result.RemoteDetail
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERROR: $($result.RemoteDetail)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                return (Complete-WorkerResult -Result $result -Path $logPath)
            }
            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Remote script copied and verified: Version=$($result.RemoteScriptVersion); SHA256=$($result.RemoteScriptHash); Path=$remoteAdminScript" | Add-Content -LiteralPath $logPath -Encoding UTF8

            $remoteScriptArgs = @($CycleScriptArgs)
            if ($result.EntraRegisteredState -eq "Pending") {
                $remoteScriptArgs += "-EntraHybridPending"
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Entra inventory state is Pending. Passing -EntraHybridPending to remote script. AlternativeSecurityIdCount=$($result.EntraAlternativeSecurityIdCount); Reason=$($result.EntraPendingReason)" | Add-Content -LiteralPath $logPath -Encoding UTF8
            }

            $argsList = @(
                "\\$Computer",
                "-accepteula",
                "-nobanner",
                "-s",
                "-h",
                "powershell.exe",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-File",
                $RemoteScriptPath
            ) + $remoteScriptArgs

            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] PsExec command: $PsExecPath $($argsList -join ' ')" | Add-Content -LiteralPath $logPath -Encoding UTF8

            $stdoutPath = Join-Path $LogRoot ("{0}_cycle{1}_{2}_stdout.tmp" -f $Computer,$CycleNumber,$runId)
            $stderrPath = Join-Path $LogRoot ("{0}_cycle{1}_{2}_stderr.tmp" -f $Computer,$CycleNumber,$runId)
            $process = Start-Process -FilePath $PsExecPath -ArgumentList $argsList -PassThru -NoNewWindow -RedirectStandardOutput $stdoutPath -RedirectStandardError $stderrPath
            $psExecTimedOut = $false
            if ($PsExecTimeoutMinutes -gt 0) {
                $timeoutMs = [int64]$PsExecTimeoutMinutes * 60 * 1000
                if (-not $process.WaitForExit($timeoutMs)) {
                    $psExecTimedOut = $true
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERROR: PsExec timed out after $PsExecTimeoutMinutes minute(s). Killing local PsExec process." | Add-Content -LiteralPath $logPath -Encoding UTF8
                    try { $process.Kill() } catch { }
                    try { [void]$process.WaitForExit(5000) } catch { }
                }
            }
            else {
                [void]$process.WaitForExit()
            }
            try { $process.Refresh() } catch { }
            $exitCode = ""
            if (-not $psExecTimedOut) {
                try {
                    if ($process.HasExited) {
                        $exitCode = [string]$process.ExitCode
                    }
                }
                catch {
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] WARN: Could not read PsExec exit code: $($_.Exception.Message)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                }
            }

            $stdoutContent = @()
            if (Test-Path -LiteralPath $stdoutPath) {
                "----- PsExec STDOUT -----" | Add-Content -LiteralPath $logPath -Encoding UTF8
                $stdoutContent = @(Get-Content -LiteralPath $stdoutPath -ErrorAction SilentlyContinue)
                $stdoutContent | Add-Content -LiteralPath $logPath -Encoding UTF8
                Remove-Item -LiteralPath $stdoutPath -Force -ErrorAction SilentlyContinue
            }
            if (Test-Path -LiteralPath $stderrPath) {
                "----- PsExec STDERR -----" | Add-Content -LiteralPath $logPath -Encoding UTF8
                $stderrContent = @(Get-Content -LiteralPath $stderrPath -ErrorAction SilentlyContinue)
                $stderrContent | Add-Content -LiteralPath $logPath -Encoding UTF8
                Remove-Item -LiteralPath $stderrPath -Force -ErrorAction SilentlyContinue
            }
            else {
                $stderrContent = @()
            }

            if ([string]::IsNullOrWhiteSpace($exitCode)) {
                $nativeExitLine = ($stderrContent | Where-Object { $_ -match "with error code\s+-?\d+" } | Select-Object -Last 1)
                if ($nativeExitLine -and $nativeExitLine -match "with error code\s+(?<Code>-?\d+)") {
                    $exitCode = $Matches.Code
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] PsExec exit code recovered from native STDERR: $exitCode" | Add-Content -LiteralPath $logPath -Encoding UTF8
                }
            }

            $result.PsExecExitCode = $exitCode
            $finalLine = $stdoutContent | Where-Object { $_ -match "^FINAL_STATUS=" } | Select-Object -Last 1
            if ($psExecTimedOut) {
                $result.Status = "PSEXEC_TIMEOUT"
                $result.RemoteDetail = "PsExec timed out after $PsExecTimeoutMinutes minute(s). The remote process may still need log review on the target."
                $result.ErrorMessage = $result.RemoteDetail
            }
            elseif ($finalLine -and $finalLine -match "^FINAL_STATUS=(?<Status>[^;]*);\s*EXIT_CODE=(?<ExitCode>[^;]*);\s*NEXT_ACTION=(?<NextAction>[^;]*);\s*DETAIL=(?<Detail>.*)$") {
                $result.RemoteStatus = $Matches.Status.Trim()
                $result.RemoteExitCode = $Matches.ExitCode.Trim()
                $result.RemoteNextAction = $Matches.NextAction.Trim()
                $result.RemoteDetail = $Matches.Detail.Trim()
                if (-not [string]::IsNullOrWhiteSpace($result.RemoteStatus)) {
                    $result.Status = $result.RemoteStatus
                }
                else {
                    $result.Status = if ($exitCode -eq "0") { "SUCCESS" } elseif ([string]::IsNullOrWhiteSpace($exitCode)) { "PSEXEC_EXIT_UNKNOWN" } else { "PSEXEC_EXIT_$exitCode" }
                }
            }
            else {
                $result.Status = if ($exitCode -eq "0") { "SUCCESS" } elseif ([string]::IsNullOrWhiteSpace($exitCode)) { "PSEXEC_EXIT_UNKNOWN" } else { "PSEXEC_EXIT_$exitCode" }
                $combinedNativeOutput = (($stdoutContent + $stderrContent) -join "`n")
                if ($combinedNativeOutput -match "(?i)(Error communicating with PsExec service|Descripteur non valide|handle is invalid)") {
                    $result.Status = "PSEXEC_COMMUNICATION_LOST"
                    $result.RemoteDetail = "PsExec lost communication with PSEXESVC after starting remote PowerShell. Remote evidence may exist, but no FINAL_STATUS/current-run CSV was returned to the launcher."
                    $result.ErrorMessage = $result.RemoteDetail
                }
                if ($combinedNativeOutput -match "(?i)(-File.*(does not exist|n.?existe|non esiste|no existe|n.?o existe|nie istnieje|nicht.*exist)|fichier.*sp.cifi..*introuvable|file.*specified.*not.*found|impossibile trovare il file specificato)") {
                    $result.Status = "REMOTE_SCRIPT_MISSING"
                    $result.RemoteDetail = "PowerShell on the remote computer reported that the -File script path does not exist: $RemoteScriptPath"
                    $result.ErrorMessage = $result.RemoteDetail
                }
                $fatalLine = $stdoutContent | Where-Object { $_ -match "FATAL ERROR:" } | Select-Object -Last 1
                if ($fatalLine) {
                    $result.RemoteDetail = ([string]$fatalLine).Trim()
                    $result.ErrorMessage = $result.RemoteDetail
                }
            }
            $derivedNextAction = Get-NextActionFromLauncherStatus -Status $result.Status
            if ((-not [string]::IsNullOrWhiteSpace($result.RemoteNextAction)) -and $result.RemoteNextAction -ne "REVIEW_LOGS") {
                $result.NextAction = $result.RemoteNextAction
            }
            else {
                $result.NextAction = $derivedNextAction
            }
            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] PsExec exit code: $exitCode" | Add-Content -LiteralPath $logPath -Encoding UTF8
            if (-not [string]::IsNullOrWhiteSpace($result.RemoteStatus)) {
                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Remote final status: $($result.RemoteStatus); Detail=$($result.RemoteDetail)" | Add-Content -LiteralPath $logPath -Encoding UTF8
            }

            if ($CollectRemoteLogs) {
                try {
                    if ($result.Status -eq "PSEXEC_COMMUNICATION_LOST" -and $CommunicationLostEvidenceWaitMinutes -gt 0) {
                        $elapsedWaitMinutes = 0
                        $pollMinutes = [Math]::Min($CommunicationLostEvidencePollMinutes, $CommunicationLostEvidenceWaitMinutes)
                        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] PsExec communication was lost after remote start. Polling every $pollMinutes minute(s), up to $CommunicationLostEvidenceWaitMinutes minute(s), for current-run final CSV before delayed evidence collection." | Add-Content -LiteralPath $logPath -Encoding UTF8
                        while ($elapsedWaitMinutes -lt $CommunicationLostEvidenceWaitMinutes) {
                            $remoteFinalStatus = $null
                            if (Test-Path -LiteralPath $remoteDataAdminDir -ErrorAction SilentlyContinue) {
                                $remoteFinalStatus = Get-RemoteEvidenceFinalStatus -EvidencePath $remoteDataAdminDir -Since ([datetime]$result.Timestamp) -RequireCompletedRun
                            }
                            if ($null -ne $remoteFinalStatus -and -not [string]::IsNullOrWhiteSpace($remoteFinalStatus.Status)) {
                                "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Completed current-run evidence detected on remote computer after $elapsedWaitMinutes minute(s). RunId=$($remoteFinalStatus.RunId); Status=$($remoteFinalStatus.Status); Csv=$($remoteFinalStatus.CsvPath); LastRun=$($remoteFinalStatus.LastRunPath). Collecting evidence now." | Add-Content -LiteralPath $logPath -Encoding UTF8
                                break
                            }

                            $remainingMinutes = $CommunicationLostEvidenceWaitMinutes - $elapsedWaitMinutes
                            $sleepMinutes = [Math]::Min($pollMinutes, $remainingMinutes)
                            if ($sleepMinutes -le 0) { break }
                            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Current-run final CSV not found yet. Sleeping $sleepMinutes minute(s) before next check." | Add-Content -LiteralPath $logPath -Encoding UTF8
                            Start-Sleep -Seconds ([int]($sleepMinutes * 60))
                            $elapsedWaitMinutes += $sleepMinutes
                        }
                    }
                    elseif ($result.Status -eq "PSEXEC_EXIT_UNKNOWN") {
                        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] PsExec returned no final status. Waiting 30 seconds before collecting remote evidence to allow the remote script to flush CSV/log/gpresult files." | Add-Content -LiteralPath $logPath -Encoding UTF8
                        Start-Sleep -Seconds 30
                    }

                    if (-not (Test-Path -LiteralPath $remoteDataAdminDir)) {
                        throw "Remote log folder not found: $remoteDataAdminDir"
                    }

                    if ((-not $KeepCentralLogHistory) -and (Test-Path -LiteralPath $centralRunDir)) {
                        Remove-Item -LiteralPath $centralRunDir -Recurse -Force -ErrorAction Stop
                    }

                    $copiedEvidenceFiles = Copy-RemoteEvidenceFolder -RemoteDataPath $remoteDataAdminDir -DestinationPath $centralRunDir -ScriptName $ScriptName
                    $result.RemoteLogsCollected = $true
                    $result.RemoteLogsPath = $centralRunDir
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Remote evidence collected to: $centralRunDir. Files=$copiedEvidenceFiles" | Add-Content -LiteralPath $logPath -Encoding UTF8

                    $currentRunDir = Join-Path $centralComputerDir "LatestCurrentRun"
                    Remove-Item -LiteralPath $currentRunDir -Recurse -Force -ErrorAction SilentlyContinue
                    $currentRunFiles = @(Get-ChildItem -LiteralPath $centralRunDir -Recurse -File -Force -ErrorAction SilentlyContinue |
                        Where-Object { $_.LastWriteTime -ge ([datetime]$result.Timestamp).AddSeconds(-5) })
                    foreach ($file in $currentRunFiles) {
                        $relativePath = $file.FullName.Substring($centralRunDir.Length).TrimStart("\")
                        $targetPath = Join-Path $currentRunDir $relativePath
                        $targetFolder = Split-Path -Parent $targetPath
                        if (-not (Test-Path -LiteralPath $targetFolder)) {
                            New-Item -ItemType Directory -Path $targetFolder -Force | Out-Null
                        }
                        Copy-Item -LiteralPath $file.FullName -Destination $targetPath -Force -ErrorAction SilentlyContinue
                    }
                    if ($currentRunFiles.Count -gt 0) {
                        $result.RemoteCurrentRunLogsPath = $currentRunDir
                        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Current-run remote evidence isolated to: $currentRunDir. Files=$($currentRunFiles.Count)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                    }

                    $completedEvidenceStatus = Get-RemoteEvidenceFinalStatus -EvidencePath $centralRunDir -Since ([datetime]$result.Timestamp) -RequireCompletedRun
                    if ($null -ne $completedEvidenceStatus -and -not [string]::IsNullOrWhiteSpace($completedEvidenceStatus.Status)) {
                        $result.InteractiveUserName = $completedEvidenceStatus.InteractiveUserName
                        $result.InteractiveUserDomain = $completedEvidenceStatus.InteractiveUserDomain
                        $result.InteractiveUserAccountName = $completedEvidenceStatus.InteractiveUserAccountName
                        $result.InteractiveUserAccountType = $completedEvidenceStatus.InteractiveUserAccountType
                        $result.InteractiveSessionName = $completedEvidenceStatus.InteractiveSessionName
                        $result.InteractiveSessionState = $completedEvidenceStatus.InteractiveSessionState
                        $result.UserIsUserAzureAD = $completedEvidenceStatus.UserIsUserAzureAD
                        $result.UserAzureAdPrt = $completedEvidenceStatus.UserAzureAdPrt
                        $result.UserSessionIsNotRemote = $completedEvidenceStatus.UserSessionIsNotRemote
                        "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Remote final evidence enriched report row. RunId=$($completedEvidenceStatus.RunId); User=$($completedEvidenceStatus.InteractiveUserAccountName); IsUserAzureAD=$($completedEvidenceStatus.UserIsUserAzureAD); AzureAdPrt=$($completedEvidenceStatus.UserAzureAdPrt)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                    }

                    if ($result.Status -in @("PSEXEC_EXIT_UNKNOWN","PSEXEC_COMMUNICATION_LOST")) {
                        $evidenceStatus = $completedEvidenceStatus
                        if ($null -ne $evidenceStatus -and -not [string]::IsNullOrWhiteSpace($evidenceStatus.Status)) {
                            $result.RemoteStatus = $evidenceStatus.Status
                            $result.RemoteExitCode = $evidenceStatus.ExitCode
                            $result.RemoteNextAction = $evidenceStatus.NextAction
                            $result.RemoteDetail = $evidenceStatus.Detail
                            $result.Status = $evidenceStatus.Status
                            if (-not [string]::IsNullOrWhiteSpace($evidenceStatus.NextAction)) {
                                $result.NextAction = $evidenceStatus.NextAction
                            }
                            else {
                                $result.NextAction = Get-NextActionFromLauncherStatus -Status $result.Status
                            }
                            if (-not [string]::IsNullOrWhiteSpace($evidenceStatus.Detail)) {
                                $result.ErrorMessage = $evidenceStatus.Detail
                            }
                            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] PsExec no-final-status result reclassified from completed current-run evidence. RunId=$($evidenceStatus.RunId); Status=$($result.Status); NextAction=$($result.NextAction); Csv=$($evidenceStatus.CsvPath); LastRun=$($evidenceStatus.LastRunPath)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                        }
                        else {
                            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] No current-run final CSV was found in collected evidence. Keeping Status=$($result.Status); NextAction=$($result.NextAction)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                        }
                    }
                }
                catch {
                    $result.RemoteLogsError = $_.Exception.Message
                    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] WARN: Remote log collection failed: $($result.RemoteLogsError)" | Add-Content -LiteralPath $logPath -Encoding UTF8
                }
            }
        }
        catch {
            $result.Status = "ERROR"
            $result.NextAction = "CHECK_CONNECTIVITY_OR_ADMIN_ACCESS"
            $message = $_.Exception.Message
            $result.ErrorMessage = $message
            "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] ERROR: $message" | Add-Content -LiteralPath $logPath -Encoding UTF8
        }

        Complete-WorkerResult -Result $result -Path $logPath
    }

    $runningJobs = @()
    $nextIndex = 0
    $completed = 0

    while ($nextIndex -lt $computers.Count -or $runningJobs.Count -gt 0) {
        while ($nextIndex -lt $computers.Count -and $runningJobs.Count -lt $ThrottleLimit) {
            $computer = $computers[$nextIndex]
            $nextIndex++

            $job = Start-Job -Name ("EHJIR_C{0}_{1}" -f $CycleNumber,$computer) -ScriptBlock $worker -ArgumentList @(
                $computer,
                $CycleNumber,
                $LocalScriptPath,
                $ScriptName,
                $RemoteRelativeDir,
                $RemoteScriptPath,
                $RemoteDataRelativeDir,
                $PsExecPath,
                $LogRoot,
                $LauncherVersion,
                [bool]$DryRun,
                $CollectRemoteLogs,
                $CentralLogRoot,
                [bool]$KeepCentralLogHistory,
                $IntuneInventorySet,
                $EntraInventoryMap,
                $AdInventoryMap,
                $CycleScriptArgs,
                $PsExecTimeoutMinutes,
                $CommunicationLostEvidenceWaitMinutes,
                $CommunicationLostEvidencePollMinutes
            )
            $runningJobs += $job
            Write-Host ("Queued {0} ({1}/{2}); running={3}" -f $computer,$nextIndex,$computers.Count,$runningJobs.Count) -ForegroundColor DarkCyan

            if ($DelayBetweenComputersSeconds -gt 0 -and $nextIndex -lt $computers.Count) {
                Start-Sleep -Seconds $DelayBetweenComputersSeconds
            }
        }

        $finishedJobs = @($runningJobs | Where-Object { $_.State -ne "Running" })
        if ($finishedJobs.Count -eq 0) {
            Start-Sleep -Seconds $JobPollSeconds
            continue
        }

        foreach ($job in $finishedJobs) {
            $received = $null
            $jobErrors = @()

            try {
                $receiveErrors = @()
                $received = Receive-Job -Job $job -ErrorAction SilentlyContinue -ErrorVariable receiveErrors
                $jobErrors = @(
                    $receiveErrors | ForEach-Object { $_.ToString() }
                    $job.ChildJobs | ForEach-Object { $_.Error } | ForEach-Object { $_.ToString() }
                    if ($job.ChildJobs.Count -gt 0 -and $job.ChildJobs[0].JobStateInfo.Reason) {
                        $job.ChildJobs[0].JobStateInfo.Reason.Message
                    }
                ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
            }
            catch {
                $jobErrors += $_.Exception.Message
            }

            if (-not $received) {
                $received = [PSCustomObject]@{
                    LauncherVersion = $LauncherVersion
                    Cycle = $CycleNumber
                    Computer = ($job.Name -replace "^EHJIR_C\d+_","")
                    Timestamp = Get-Date
                    DryRun = [bool]$DryRun
                    DnsResolved = $false
                    DnsAddressList = ""
                    AdminShareReachable = $false
                    PingReachable = $false
                    RemoteDirectoryCreated = $false
                    ScriptCopied = $false
                    PsExecExitCode = ""
                    RemoteStatus = ""
                    RemoteExitCode = ""
                    RemoteNextAction = ""
                    RemoteDetail = ""
                    NextAction = "CHECK_JOB_ERROR"
                    EffectiveStatus = ""
                    EffectiveNextAction = ""
                    InteractiveUserName = ""
                    InteractiveUserDomain = ""
                    InteractiveUserAccountName = ""
                    InteractiveUserAccountType = ""
                    InteractiveSessionName = ""
                    InteractiveSessionState = ""
                    UserIsUserAzureAD = ""
                    UserAzureAdPrt = ""
                    UserSessionIsNotRemote = ""
                    IntuneInventoryPresent = ""
                    EntraInventoryPresent = ""
                    EntraRegisteredState = ""
                    EntraAlternativeSecurityIdCount = ""
                    EntraPendingReason = ""
                    EntraRegistrationDateTime = ""
                    EntraTrustType = ""
                    EntraDeviceId = ""
                    EntraObjectId = ""
                    ADInventoryPresent = ""
                    ADDomain = ""
                    ADEnabled = ""
                    ADDNSHostName = ""
                    ADDistinguishedName = ""
                    ADOperatingSystem = ""
                    ADLastLogonTimestampUtc = ""
                    AdminShareFailureType = ""
                    PostCycleIntuneInventoryChecked = ""
                    PostCycleIntuneInventoryPresent = ""
                    PostCycleIntuneEnrollmentDetected = ""
                    PostCycleIntuneInventoryCsv = ""
                    PostCycleIntuneInventoryError = ""
                    PostCycleEntraInventoryChecked = ""
                    PostCycleEntraInventoryPresent = ""
                    PostCycleEntraRegisteredState = ""
                    PostCycleEntraAlternativeSecurityIdCount = ""
                    PostCycleEntraPendingResolved = ""
                    PostCycleEntraInventoryCsv = ""
                    PostCycleEntraInventoryError = ""
                    PostCycleADInventoryChecked = ""
                    PostCycleADInventoryPresent = ""
                    PostCycleADInventoryCsv = ""
                    PostCycleADInventoryError = ""
                    RemoteLogsCollected = $false
                    RemoteLogsPath = ""
                    RemoteCurrentRunLogsPath = ""
                    RemoteLogsError = ""
                    Status = "JOB_ERROR"
                    LogPath = ""
                    ErrorMessage = ($jobErrors -join " | ")
                }
            }

            foreach ($item in @($received)) {
                if ($null -ne $item) {
                    if ($jobErrors.Count -gt 0 -and -not $item.PSObject.Properties["JobErrorMessage"]) {
                        $item | Add-Member -NotePropertyName JobErrorMessage -NotePropertyValue ($jobErrors -join " | ") -Force
                    }
                    $item | Add-Member -NotePropertyName EffectiveStatus -NotePropertyValue ([string]$item.Status) -Force
                    $item | Add-Member -NotePropertyName EffectiveNextAction -NotePropertyValue ([string]$item.NextAction) -Force
                    $summary.Add($item)
                    try {
                        Add-LiveCycleReportRow -Path $liveSummaryPath -Columns $reportColumns -Row $item
                    }
                    catch {
                        Write-Host ("Cycle {0}: failed to append live report row for {1}: {2}" -f $CycleNumber,$item.Computer,$_.Exception.Message) -ForegroundColor Yellow
                    }
                    if (-not $DryRun -and (Test-AlreadyEnrolledCycleResult -Result $item)) {
                        try {
                            $moveSingleResult = Move-AlreadyEnrolledComputersFromList -ComputerListPath $ComputerListPath -CycleSummary @($item)
                            if ($moveSingleResult.Moved -gt 0) {
                                Write-Host ("Moved already-enrolled computer from Computers.txt to {0}: {1}" -f $moveSingleResult.AlreadyEnrolledPath,$item.Computer) -ForegroundColor Green
                            }
                        }
                        catch {
                            Write-Host ("Cycle {0}: failed to move already-enrolled computer {1}: {2}" -f $CycleNumber,$item.Computer,$_.Exception.Message) -ForegroundColor Yellow
                        }
                    }
                    $completed++
                    if (($completed % 10) -eq 0) {
                        try {
                            $liveRows = @($summary | ForEach-Object { $_ })
                            New-CycleHtmlReport -Summary $liveRows -Path $liveHtmlPath -CycleNumber $CycleNumber -GeneratedAt (Get-Date)
                        }
                        catch {
                            Write-Host ("Cycle {0}: failed to update live HTML report: {1}" -f $CycleNumber,$_.Exception.Message) -ForegroundColor Yellow
                        }
                    }
                    $messageSuffix = ""
                    if ($item.PSObject.Properties["RemoteDetail"] -and -not [string]::IsNullOrWhiteSpace([string]$item.RemoteDetail)) {
                        $messageSuffix = " - $($item.RemoteDetail)"
                    }
                    elseif ($item.PSObject.Properties["ErrorMessage"] -and -not [string]::IsNullOrWhiteSpace([string]$item.ErrorMessage)) {
                        $messageSuffix = " - $($item.ErrorMessage)"
                    }
                    elseif ($item.PSObject.Properties["JobErrorMessage"] -and -not [string]::IsNullOrWhiteSpace([string]$item.JobErrorMessage)) {
                        $messageSuffix = " - $($item.JobErrorMessage)"
                    }

                    $actionText = ""
                    if ($item.PSObject.Properties["NextAction"] -and -not [string]::IsNullOrWhiteSpace([string]$item.NextAction)) {
                        $actionText = " | NextAction=$($item.NextAction)"
                    }

                    if ($item.Status -eq "JOB_ERROR" -or $item.Status -eq "ERROR") {
                        Write-Host ("Completed {0}/{1}: {2} => {3}{4}{5}" -f $completed,$computers.Count,$item.Computer,$item.Status,$actionText,$messageSuffix) -ForegroundColor Red
                    }
                    else {
                        Write-Host ("Completed {0}/{1}: {2} => {3}{4}{5}" -f $completed,$computers.Count,$item.Computer,$item.Status,$actionText,$messageSuffix)
                    }
                }
            }

            Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
        }

        $finishedIds = @($finishedJobs | Select-Object -ExpandProperty Id)
        $runningJobs = @($runningJobs | Where-Object { $finishedIds -notcontains $_.Id })
    }

    $summaryRowsForPostCycle = @($summary | ForEach-Object { $_ })

    if (-not $DryRun -and -not $SkipPostCycleIntuneInventory) {
        Write-Host ("Cycle {0}: refreshing full post-cycle Intune inventory..." -f $CycleNumber) -ForegroundColor Cyan
        $postInventoryLogPath = Join-Path $ReportRoot ("DevicesIntune_PostCycle_cycle{0}_{1}.log" -f $CycleNumber,(Get-Date -Format "yyyyMMdd_HHmmss"))
        $postInventory = Invoke-FullIntuneInventoryExport `
            -ExportScriptPath $ExportIntuneScriptPath `
            -OutputPath $IntuneInventoryCsv `
            -LogPath $postInventoryLogPath `
            -PageSize $PostCycleIntuneInventoryPageSize

        if ($postInventory.Success) {
            $postSet = $postInventory.InventorySet
            $script:IntuneInventorySet = $postSet
            $newlyDetected = 0
            foreach ($row in $summaryRowsForPostCycle) {
                $key = Get-ComputerListKey -ComputerName $row.Computer
                $postPresent = [bool]($postSet -and $postSet.ContainsKey($key))
                $preKnown = $row.PSObject.Properties["IntuneInventoryPresent"] -and -not [string]::IsNullOrWhiteSpace([string]$row.IntuneInventoryPresent)
                $prePresent = Test-BooleanLikeTrue -Value $row.IntuneInventoryPresent
                $postDetected = $preKnown -and (-not $prePresent) -and $postPresent
                if ($postDetected) { $newlyDetected++ }

                $row | Add-Member -NotePropertyName PostCycleIntuneInventoryChecked -NotePropertyValue $true -Force
                $row | Add-Member -NotePropertyName PostCycleIntuneInventoryPresent -NotePropertyValue $postPresent -Force
                $row | Add-Member -NotePropertyName PostCycleIntuneEnrollmentDetected -NotePropertyValue $postDetected -Force
                $row | Add-Member -NotePropertyName PostCycleIntuneInventoryCsv -NotePropertyValue $postInventory.CsvPath -Force
                $row | Add-Member -NotePropertyName PostCycleIntuneInventoryError -NotePropertyValue "" -Force
                if ($postPresent) {
                    $row | Add-Member -NotePropertyName EffectiveStatus -NotePropertyValue "ENROLLED_DETECTED_POST_CYCLE" -Force
                    $row | Add-Member -NotePropertyName EffectiveNextAction -NotePropertyValue "NO_ACTION_ALREADY_INTUNE" -Force
                }
                else {
                    $row | Add-Member -NotePropertyName EffectiveStatus -NotePropertyValue ([string]$row.Status) -Force
                    $row | Add-Member -NotePropertyName EffectiveNextAction -NotePropertyValue ([string]$row.NextAction) -Force
                }
            }

            $postPresentCount = @($summaryRowsForPostCycle | Where-Object { $_.PostCycleIntuneInventoryPresent -eq $true }).Count
            Write-Host ("Cycle {0}: post-cycle Intune inventory found {1}/{2}; newly detected this cycle={3}; CSV={4}" -f $CycleNumber,$postPresentCount,$summaryRowsForPostCycle.Count,$newlyDetected,$postInventory.CsvPath) -ForegroundColor Green
        }
        else {
            foreach ($row in $summaryRowsForPostCycle) {
                $row | Add-Member -NotePropertyName PostCycleIntuneInventoryChecked -NotePropertyValue $true -Force
                $row | Add-Member -NotePropertyName PostCycleIntuneInventoryPresent -NotePropertyValue "" -Force
                $row | Add-Member -NotePropertyName PostCycleIntuneEnrollmentDetected -NotePropertyValue "" -Force
                $row | Add-Member -NotePropertyName PostCycleIntuneInventoryCsv -NotePropertyValue $postInventory.CsvPath -Force
                $row | Add-Member -NotePropertyName PostCycleIntuneInventoryError -NotePropertyValue $postInventory.Error -Force
                $row | Add-Member -NotePropertyName EffectiveStatus -NotePropertyValue ([string]$row.Status) -Force
                $row | Add-Member -NotePropertyName EffectiveNextAction -NotePropertyValue ([string]$row.NextAction) -Force
            }
            Write-Host ("Cycle {0}: post-cycle Intune inventory failed: {1}" -f $CycleNumber,$postInventory.Error) -ForegroundColor Yellow
        }

        try {
            $summaryRowsForPostCycle | Select-Object $reportColumns | Export-Csv -LiteralPath $liveSummaryPath -NoTypeInformation -Encoding UTF8
        }
        catch {
            Write-Host ("Cycle {0}: failed to rewrite live CSV with post-cycle Intune columns: {1}" -f $CycleNumber,$_.Exception.Message) -ForegroundColor Yellow
        }
    }

    if (-not $DryRun -and -not [string]::IsNullOrWhiteSpace($EntraInventoryCsv)) {
        Write-Host ("Cycle {0}: refreshing full post-cycle Entra device inventory..." -f $CycleNumber) -ForegroundColor Cyan
        $postEntraInventoryLogPath = Join-Path $ReportRoot ("DevicesEntra_PostCycle_cycle{0}_{1}.log" -f $CycleNumber,(Get-Date -Format "yyyyMMdd_HHmmss"))
        $postEntraInventory = Invoke-FullEntraInventoryExport `
            -ExportScriptPath $ExportEntraScriptPath `
            -OutputPath $EntraInventoryCsv `
            -LogPath $postEntraInventoryLogPath `
            -PageSize $PostCycleIntuneInventoryPageSize

        if ($postEntraInventory.Success) {
            $postEntraMap = $postEntraInventory.InventoryMap
            $script:EntraInventoryMap = $postEntraMap
            $pendingResolved = 0

            foreach ($row in $summaryRowsForPostCycle) {
                $key = Get-ComputerListKey -ComputerName $row.Computer
                $postEntraPresent = [bool]($postEntraMap -and $postEntraMap.ContainsKey($key))
                $postEntraState = ""
                $postAltSecIdCount = ""

                if ($postEntraPresent) {
                    $postEntraRow = $postEntraMap[$key]
                    if ($postEntraRow.PSObject.Properties["EntraRegisteredState"]) { $postEntraState = [string]$postEntraRow.EntraRegisteredState }
                    if ($postEntraRow.PSObject.Properties["AlternativeSecurityIdCount"]) { $postAltSecIdCount = [string]$postEntraRow.AlternativeSecurityIdCount }
                }

                $wasPending = ($row.PSObject.Properties["EntraRegisteredState"] -and [string]$row.EntraRegisteredState -eq "Pending")
                $isPendingNow = ($postEntraState -eq "Pending")
                $resolvedThisCycle = [bool]($wasPending -and $postEntraPresent -and -not $isPendingNow)
                if ($resolvedThisCycle) { $pendingResolved++ }

                $row | Add-Member -NotePropertyName PostCycleEntraInventoryChecked -NotePropertyValue $true -Force
                $row | Add-Member -NotePropertyName PostCycleEntraInventoryPresent -NotePropertyValue $postEntraPresent -Force
                $row | Add-Member -NotePropertyName PostCycleEntraRegisteredState -NotePropertyValue $postEntraState -Force
                $row | Add-Member -NotePropertyName PostCycleEntraAlternativeSecurityIdCount -NotePropertyValue $postAltSecIdCount -Force
                $row | Add-Member -NotePropertyName PostCycleEntraPendingResolved -NotePropertyValue $resolvedThisCycle -Force
                $row | Add-Member -NotePropertyName PostCycleEntraInventoryCsv -NotePropertyValue $postEntraInventory.CsvPath -Force
                $row | Add-Member -NotePropertyName PostCycleEntraInventoryError -NotePropertyValue "" -Force

                if ($resolvedThisCycle -and [string]$row.EffectiveStatus -ne "ENROLLED_DETECTED_POST_CYCLE") {
                    $row | Add-Member -NotePropertyName EffectiveStatus -NotePropertyValue "ENTRA_PENDING_RESOLVED_POST_CYCLE" -Force
                    $row | Add-Member -NotePropertyName EffectiveNextAction -NotePropertyValue "RECHECK_INTUNE_ENROLLMENT" -Force
                }
            }

            $postPendingCount = @($summaryRowsForPostCycle | Where-Object { $_.PostCycleEntraRegisteredState -eq "Pending" }).Count
            Write-Host ("Cycle {0}: post-cycle Entra inventory pending={1}; pending resolved this cycle={2}; CSV={3}" -f $CycleNumber,$postPendingCount,$pendingResolved,$postEntraInventory.CsvPath) -ForegroundColor Green
        }
        else {
            foreach ($row in $summaryRowsForPostCycle) {
                $row | Add-Member -NotePropertyName PostCycleEntraInventoryChecked -NotePropertyValue $true -Force
                $row | Add-Member -NotePropertyName PostCycleEntraInventoryPresent -NotePropertyValue "" -Force
                $row | Add-Member -NotePropertyName PostCycleEntraRegisteredState -NotePropertyValue "" -Force
                $row | Add-Member -NotePropertyName PostCycleEntraAlternativeSecurityIdCount -NotePropertyValue "" -Force
                $row | Add-Member -NotePropertyName PostCycleEntraPendingResolved -NotePropertyValue "" -Force
                $row | Add-Member -NotePropertyName PostCycleEntraInventoryCsv -NotePropertyValue $postEntraInventory.CsvPath -Force
                $row | Add-Member -NotePropertyName PostCycleEntraInventoryError -NotePropertyValue $postEntraInventory.Error -Force
            }
            Write-Host ("Cycle {0}: post-cycle Entra inventory failed: {1}" -f $CycleNumber,$postEntraInventory.Error) -ForegroundColor Yellow
        }

        try {
            $summaryRowsForPostCycle | Select-Object $reportColumns | Export-Csv -LiteralPath $liveSummaryPath -NoTypeInformation -Encoding UTF8
        }
        catch {
            Write-Host ("Cycle {0}: failed to rewrite live CSV with post-cycle Entra columns: {1}" -f $CycleNumber,$_.Exception.Message) -ForegroundColor Yellow
        }
    }

    if (-not $DryRun -and -not $AdInventoryUsesRecentRootCsv -and -not [string]::IsNullOrWhiteSpace($AdInventoryCsv)) {
        $postAdScope = if ([string]::IsNullOrWhiteSpace($AdDomain)) { "forest" } else { "domain '$AdDomain'" }
        Write-Host ("Cycle {0}: refreshing full post-cycle AD computer inventory. Scope={1}..." -f $CycleNumber,$postAdScope) -ForegroundColor Cyan
        $postAdInventoryLogPath = Join-Path $ReportRoot ("DevicesAD_PostCycle_cycle{0}_{1}.log" -f $CycleNumber,(Get-Date -Format "yyyyMMdd_HHmmss"))
        $postAdInventory = Invoke-FullAdInventoryExport `
            -ExportScriptPath $ExportAdScriptPath `
            -OutputPath $AdInventoryCsv `
            -LogPath $postAdInventoryLogPath `
            -Domain $AdDomain

        if ($postAdInventory.Success) {
            $postAdMap = $postAdInventory.InventoryMap
            $script:AdInventoryMap = $postAdMap

            foreach ($row in $summaryRowsForPostCycle) {
                $key = Get-ComputerListKey -ComputerName $row.Computer
                $postAdPresent = [bool]($postAdMap -and $postAdMap.ContainsKey($key))

                $row | Add-Member -NotePropertyName PostCycleADInventoryChecked -NotePropertyValue $true -Force
                $row | Add-Member -NotePropertyName PostCycleADInventoryPresent -NotePropertyValue $postAdPresent -Force
                $row | Add-Member -NotePropertyName PostCycleADInventoryCsv -NotePropertyValue $postAdInventory.CsvPath -Force
                $row | Add-Member -NotePropertyName PostCycleADInventoryError -NotePropertyValue "" -Force
            }

            $postAdPresentCount = @($summaryRowsForPostCycle | Where-Object { $_.PostCycleADInventoryPresent -eq $true }).Count
            Write-Host ("Cycle {0}: post-cycle AD inventory found {1}/{2}; CSV={3}" -f $CycleNumber,$postAdPresentCount,$summaryRowsForPostCycle.Count,$postAdInventory.CsvPath) -ForegroundColor Green
        }
        else {
            foreach ($row in $summaryRowsForPostCycle) {
                $row | Add-Member -NotePropertyName PostCycleADInventoryChecked -NotePropertyValue $true -Force
                $row | Add-Member -NotePropertyName PostCycleADInventoryPresent -NotePropertyValue "" -Force
                $row | Add-Member -NotePropertyName PostCycleADInventoryCsv -NotePropertyValue $postAdInventory.CsvPath -Force
                $row | Add-Member -NotePropertyName PostCycleADInventoryError -NotePropertyValue $postAdInventory.Error -Force
            }
            Write-Host ("Cycle {0}: post-cycle AD inventory failed: {1}" -f $CycleNumber,$postAdInventory.Error) -ForegroundColor Yellow
        }

        try {
            $summaryRowsForPostCycle | Select-Object $reportColumns | Export-Csv -LiteralPath $liveSummaryPath -NoTypeInformation -Encoding UTF8
        }
        catch {
            Write-Host ("Cycle {0}: failed to rewrite live CSV with post-cycle AD columns: {1}" -f $CycleNumber,$_.Exception.Message) -ForegroundColor Yellow
        }
    }

    try {
        New-CycleHtmlReport -Summary $summaryRowsForPostCycle -Path $liveHtmlPath -CycleNumber $CycleNumber -GeneratedAt (Get-Date)
    }
    catch {
        Write-Host ("Cycle {0}: failed to write final live HTML report: {1}" -f $CycleNumber,$_.Exception.Message) -ForegroundColor Yellow
    }

    if (-not $DryRun) {
        try {
            $moveResult = Move-AlreadyEnrolledComputersFromList -ComputerListPath $ComputerListPath -CycleSummary $summaryRowsForPostCycle
            if ($moveResult.Moved -gt 0) {
                Write-Host ("Cycle {0}: moved {1} already-enrolled computer(s) to {2}" -f $CycleNumber,$moveResult.Moved,$moveResult.AlreadyEnrolledPath) -ForegroundColor Green
            }
            else {
                Write-Host ("Cycle {0}: no already-enrolled computer moved from Computers.txt. {1}" -f $CycleNumber,$moveResult.Detail) -ForegroundColor DarkGray
            }
        }
        catch {
            Write-Host ("Cycle {0}: failed to update ComputersAlreadyEnrolled.txt: {1}" -f $CycleNumber,$_.Exception.Message) -ForegroundColor Yellow
        }
    }

    $summaryPath = Join-Path $ReportRoot ("PsExec_IntuneHybridJoinRepair_Summary_cycle{0}_{1}.csv" -f $CycleNumber,(Get-Date -Format "yyyyMMdd_HHmmss"))
    $summaryRowsForPostCycle | Export-Csv -LiteralPath $summaryPath -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host ("Cycle {0} status counts:" -f $CycleNumber) -ForegroundColor Cyan
    @($summary | Group-Object -Property Status | Sort-Object Count -Descending | ForEach-Object {
        "  {0,-45} {1,5}" -f $_.Name,$_.Count
    }) | ForEach-Object { Write-Host $_ }

    Write-Host ("Cycle {0} next-action counts:" -f $CycleNumber) -ForegroundColor Cyan
    @($summary | Group-Object -Property NextAction | Sort-Object Count -Descending | ForEach-Object {
        "  {0,-45} {1,5}" -f $_.Name,$_.Count
    }) | ForEach-Object { Write-Host $_ }

    if ($IntuneInventorySet -and $IntuneInventorySet.Count -gt 0) {
        $present = @($summary | Where-Object { $_.IntuneInventoryPresent -eq $true }).Count
        $absent = @($summary | Where-Object { $_.IntuneInventoryPresent -eq $false }).Count
        Write-Host ("Cycle {0} Intune inventory match: Present={1}; Absent={2}" -f $CycleNumber,$present,$absent) -ForegroundColor Cyan
    }

    if ($AdInventoryMap -and $AdInventoryMap.Count -gt 0) {
        $present = @($summary | Where-Object { $_.ADInventoryPresent -eq $true }).Count
        $absent = @($summary | Where-Object { $_.ADInventoryPresent -eq $false }).Count
        Write-Host ("Cycle {0} AD inventory match: Present={1}; Absent={2}" -f $CycleNumber,$present,$absent) -ForegroundColor Cyan
    }

    if (@($summaryRowsForPostCycle | Where-Object { $_.PSObject.Properties["PostCycleIntuneInventoryChecked"] -and $_.PostCycleIntuneInventoryChecked -eq $true }).Count -gt 0) {
        $postPresent = @($summaryRowsForPostCycle | Where-Object { $_.PostCycleIntuneInventoryPresent -eq $true }).Count
        $postAbsent = @($summaryRowsForPostCycle | Where-Object { $_.PostCycleIntuneInventoryPresent -eq $false }).Count
        $postNew = @($summaryRowsForPostCycle | Where-Object { $_.PostCycleIntuneEnrollmentDetected -eq $true }).Count
        Write-Host ("Cycle {0} post-cycle Intune inventory: Present={1}; Absent={2}; NewlyDetected={3}" -f $CycleNumber,$postPresent,$postAbsent,$postNew) -ForegroundColor Cyan
    }

    if (@($summaryRowsForPostCycle | Where-Object { $_.PSObject.Properties["PostCycleADInventoryChecked"] -and $_.PostCycleADInventoryChecked -eq $true }).Count -gt 0) {
        $postPresent = @($summaryRowsForPostCycle | Where-Object { $_.PostCycleADInventoryPresent -eq $true }).Count
        $postAbsent = @($summaryRowsForPostCycle | Where-Object { $_.PostCycleADInventoryPresent -eq $false }).Count
        Write-Host ("Cycle {0} post-cycle AD inventory: Present={1}; Absent={2}" -f $CycleNumber,$postPresent,$postAbsent) -ForegroundColor Cyan
    }

    $htmlPath = [System.IO.Path]::ChangeExtension($summaryPath, ".html")
    try {
        New-CycleHtmlReport -Summary $summaryRowsForPostCycle -Path $htmlPath -CycleNumber $CycleNumber -GeneratedAt (Get-Date)
        Write-Host ("Cycle {0} HTML report: {1}" -f $CycleNumber,$htmlPath) -ForegroundColor Green
    }
    catch {
        Write-Host ("Cycle {0} HTML report failed: {1}: {2}" -f $CycleNumber,$_.Exception.GetType().FullName,$_.Exception.Message) -ForegroundColor Yellow
        if ($_.ScriptStackTrace) {
            Write-Host $_.ScriptStackTrace -ForegroundColor DarkYellow
        }
    }
    Write-Host ("Cycle {0} done. Summary: {1}" -f $CycleNumber,$summaryPath) -ForegroundColor Green
    return $summaryPath
}

$cycle = 0
do {
    $cycle++
    $cycleArgs = @($scriptArgsBase)
    if ($IgnoreRunGuard -and ($cycle -eq 1 -or $IgnoreRunGuardEveryCycle)) {
        $cycleArgs += "-IgnoreRunGuard"
    }

    $null = Invoke-IntuneHybridJoinRepairCycle -CycleNumber $cycle -CycleScriptArgs $cycleArgs

    if ($RunOnce) { break }
    if ($MaxCycles -gt 0 -and $cycle -ge $MaxCycles) { break }

    if ($DelayBetweenCyclesMinutes -gt 0) {
        Write-Host ("Waiting {0} minute(s) before next cycle. Press Ctrl+C to stop." -f $DelayBetweenCyclesMinutes) -ForegroundColor DarkGray
        Start-Sleep -Seconds ($DelayBetweenCyclesMinutes * 60)
    }
} while ($true)

Write-Host ""
Write-Host "Launcher stopped after $cycle cycle(s)." -ForegroundColor Green
