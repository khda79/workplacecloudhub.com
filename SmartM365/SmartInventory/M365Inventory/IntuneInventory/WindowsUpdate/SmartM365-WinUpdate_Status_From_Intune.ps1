<#
.SYNOPSIS
Windows Update (Feature Update) device status from Intune with Win11 Readiness enrichment
and full remediation columns. Single CSV output for Power BI consumption.

.DESCRIPTION
- Discovers Feature Update policies via Graph beta: deviceManagement/windowsFeatureUpdateProfiles
- For each policy: exports Intune report via Graph beta reports/exportJobs (required filter: PolicyId eq '<id>')
- Downloads ZIP, extracts CSV, consolidates into one dataset
- Optional enrichment with Win11 Readiness CSV (join by NormalizedDeviceName derived from DeviceName)
- Missing OSVersion fallback through Intune managed-device ID to Entra device ID
- Computed columns added on all rows:
    RiskBucket, BlockingReason, OSVersion, OSVersionSource, DaysSinceLastStatus,
    ActionPriority, ActionCode, ActionDescription, ActionOwner
- Single output CSV (DATA-ALL + DATA-LAST + Archive + SharePoint):
    Intune_WindowsUpdate_Status.csv
- All summaries, blocking splits, action plans are derived in Power BI from this single table
- Token auto-refresh: Graph token refreshed automatically if age > 55 min
- SharePoint upload: latest CSV copied to SharePoint Online via SmartM365.Core when enabled
- Emails:
    - Fatal error HTML email (on failure)
    - Optional action-oriented HTML summary by unique device and policy

.REQUIREMENTS
- PowerShell 7+
- MSAL.PS module
- App-only certificate authentication for Graph
- SmartM365.Core module (Modules\SmartM365.Core\SmartM365.Core.psd1)
- Graph app permission: Sites.Selected with site-level write grant (for SharePoint upload)

PARAMETERS
  -DryRun                    : Run all API calls and processing; skip file writes and emails
  -MaxPolicies               : Limit number of policies processed (0 = all; useful for testing)
  -FeatureUpdatePolicyId     : Optional Intune Feature Update policy id to process
  -EnableReadinessEnrichment : Toggle Win11 Readiness CSV enrichment (default: $true)
  -SummaryEmailMode          : Always | OnChange | Never (default: Always)
  -EnableSummaryEmail        : Toggle success email (default: $true)
  -EnableErrorEmail          : Toggle error email (default: $true)
  -RiskTopN                  : Number of action-required devices shown in email (default: 10)

VERSION
  1.58
.VERSION
1.61
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Minimum application permissions: DeviceManagementConfiguration.Read.All, DeviceManagementManagedDevices.Read.All
#>

param(
    [string]$Tenant = 'test',
    [switch]$DryRun,
    [int]$MaxPolicies = 0,
    [string]$FeatureUpdatePolicyId = '',
    [bool]$EnableReadinessEnrichment = $true,
    [ValidateSet("Always","OnChange","Never")][string]$SummaryEmailMode = "Always",
    [bool]$EnableSummaryEmail = $true,
    [bool]$EnableErrorEmail = $true,
    [int]$RiskTopN = 10,
    [ValidateRange(1,3650)][int]$DetailedArchiveRetentionDays = 7,
    [int]$MaxItems = 0
)
# Local lifecycle only; no collection or publication is performed by this guard.
$smartM365LifecycleRoot=$PSScriptRoot
while ($smartM365LifecycleRoot -and -not (Test-Path -LiteralPath (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1') -PathType Leaf)) {
    $smartM365LifecycleParent=Split-Path $smartM365LifecycleRoot -Parent
    if($smartM365LifecycleParent -eq $smartM365LifecycleRoot){break}
    $smartM365LifecycleRoot=$smartM365LifecycleParent
}
if(-not $smartM365LifecycleRoot){throw 'SmartM365 lifecycle support is missing; update the complete runtime.'}
. (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1')
. (Join-Path $smartM365LifecycleRoot 'SmartInventory/Common/SmartM365.ReportRenderers.ps1')
. (Join-Path $smartM365LifecycleRoot 'SmartInventory/Common/SmartM365.ReportInputs.ps1')
$smartM365RunGuard=Start-SmartM365InventoryRun -ScriptPath $PSCommandPath -RequireReturn
try {
if ($PSBoundParameters.ContainsKey('MaxItems') -and $MaxItems -gt 0) {
    $global:SmartM365MaxItems = [int]$MaxItems
    $global:SmartM365TestMaxItems = [int]$MaxItems
    $global:SmartM365IsMaxItemsRun = $true
    foreach ($smartM365LimitName in @('TopUsers','TopMailboxes','MaxDevices','MaxSites','MaxTeams','MaxApps','MaxPolicies','Limit','MaxPages')) {
        $smartM365LimitVariable = Get-Variable -Name $smartM365LimitName -Scope Script -ErrorAction SilentlyContinue
        if ($smartM365LimitVariable -and -not $PSBoundParameters.ContainsKey($smartM365LimitName) -and $null -ne $smartM365LimitVariable.Value) {
            Set-Variable -Name $smartM365LimitName -Value ([int]$MaxItems) -Scope Script
        }
    }
}
$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidates = @(
            (Join-Path -Path $d -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $d -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )
        foreach ($p in $candidates) {
            if (Test-Path -LiteralPath $p) { return $p }
        }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $tenantContextPath
$script:SmartM365GlobalConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$CsvTenantKey = ([string]$global:SmartM365TenantKey).Trim()
if ([string]::IsNullOrWhiteSpace($CsvTenantKey)) { throw 'The effective tenant context has no TenantKey for CSV joins.' }

# ==========================================================
# Version
# ==========================================================
$ScriptVersion = "1.61"

function Get-WinUpdateSourceReceiptUploadPaths {
    param([AllowNull()][AllowEmptyString()][string]$ReceiptPath)
    if ([string]::IsNullOrWhiteSpace($ReceiptPath)) { return }
    if ($ReceiptPath -match '\.run\.json\.txt$') { return $ReceiptPath }
    if ($ReceiptPath -notmatch '\.current\.json\.txt$') { throw 'Unexpected source receipt path.' }
    $runPath = $ReceiptPath -replace '\.current\.json\.txt$', '.run.json.txt'
    if (-not (Test-Path -LiteralPath $runPath -PathType Leaf)) { throw 'Source run state is missing; the current receipt will not be mirrored alone.' }
    $current = [IO.File]::ReadAllText($ReceiptPath) | ConvertFrom-Json
    $run = [IO.File]::ReadAllText($runPath) | ConvertFrom-Json
    if (-not $current.RunId -or $current.RunId -cne $run.RunId -or
        $current.Owner -cne 'SmartInventory-SourceReceipt' -or $run.Owner -cne 'SmartInventory-SourceRun' -or
        $run.Status -cne 'Completed' -or $run.UnqualifiedPublication -isnot [bool] -or $run.UnqualifiedPublication) {
        throw 'Source current receipt and completed run state do not match; neither will be mirrored.'
    }
    return @($ReceiptPath, $runPath)
}

# ==========================================================
# App-only authentication parameters
# ==========================================================
function Get-ScriptLocalConfig {
    param()
    Read-SmartM365ScriptConfiguration -ScriptPath $PSCommandPath
}

function Resolve-SmartM365ConfigValue {
    param([AllowNull()]$Value)
    Resolve-SmartM365ConfigurationTokens -Value $Value -Config $script:SmartM365GlobalConfig
}
function Test-SmartM365UseGlobalConfigValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -isnot [string]) { return $false }

    $text = $Value.Trim()
    return ($text -in @('__USE_GLOBAL__', 'USE_GLOBAL', '**USE_GLOBAL**', '**USE\_GLOBAL**'))
}

function Get-ScriptLocalConfigValue {
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$Name, [AllowNull()]$DefaultValue = $null)
    Get-SmartM365ConfigurationValue -Name $Name -LocalConfig $Config -EffectiveConfig $script:SmartM365GlobalConfig -DefaultValue $DefaultValue -EffectiveFallback Inherit
}

$ScriptLocalConfig = Get-ScriptLocalConfig


$global:RetentionMaxLogs = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxLogs' -DefaultValue 30)
if (-not $PSBoundParameters.ContainsKey('DetailedArchiveRetentionDays')) {
    $DetailedArchiveRetentionDays = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DetailedArchiveRetentionDays' -DefaultValue 7)
}

$global:EnableSharePointUpload = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointTargetFolderPath' -DefaultValue ''
$AppId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AppId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$TenantId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$Thumb = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumb' -DefaultValue '0000000000000000000000000000000000000000'
$OrgDomain = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue 'contoso.onmicrosoft.com'
$LogAllRootPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LogAllRootPath' -DefaultValue ''
if ([string]::IsNullOrWhiteSpace($FeatureUpdatePolicyId)) {
    $FeatureUpdatePolicyId = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'FeatureUpdatePolicyId' -DefaultValue '')
}

# ==========================================================
# Report selection
# ==========================================================
$ReportName       = "FeatureUpdateDeviceState"
$LocalizationType = "LocalizedValuesAsAdditionalColumn"
$ReportSelect     = $null

# ==========================================================
# Last status date column (configurable)
# Column name in the Intune FeatureUpdateDeviceState export used to compute DaysSinceLastStatus.
# Common names: LastUpdateStatusTime, LastUpdateStatusDateTime, PolicyLastModifiedTime
# Set to $null or "" to disable DaysSinceLastStatus computation.
# ==========================================================
$LastStatusDateColumn = "LastUpdateStatusTime"

# ==========================================================
# Win11 Readiness enrichment input
# ==========================================================
$ReadinessCsvPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ReadinessCsvPath' -DefaultValue ""
$ReadinessSummaryCsvPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ReadinessSummaryCsvPath' -DefaultValue ""
if ([string]::IsNullOrWhiteSpace($ReadinessSummaryCsvPath) -and -not [string]::IsNullOrWhiteSpace($ReadinessCsvPath)) {
    $ReadinessSummaryCsvPath = $ReadinessCsvPath -replace '_UpgradeEligibility\.csv$', '_UpgradeEligibility_Summary.csv'
}

# ==========================================================
# Output paths (UNC)
# ==========================================================
$ScriptCsvLogFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ScriptCsvLogFolderPath' -DefaultValue ""

$CsvName  = "Intune_WindowsUpdate_Status.csv"
$CsvFinal = Join-Path $ScriptCsvLogFolderPath $CsvName
$CsvTemp  = Join-Path $ScriptCsvLogFolderPath "$CsvName.tmp"

$ArchivePath = Join-Path $ScriptCsvLogFolderPath "Archive"
$LogsPath    = if ([string]::IsNullOrWhiteSpace($LogAllRootPath)) {
    Join-Path $ScriptCsvLogFolderPath "Logs"
} else {
    Join-Path $LogAllRootPath "SmartM365-WinUpdate-Status-Inventory"
}
$WorkPath    = Join-Path $ScriptCsvLogFolderPath "Work"

$LatestCsvFolderPath  = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ""
$CsvLastFinal  = Join-Path $LatestCsvFolderPath $CsvName
$CsvLastTemp   = Join-Path $LatestCsvFolderPath "$CsvName.tmp"
$IntuneDeviceInventoryCsvPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'IntuneDeviceInventoryCsvPath' -DefaultValue ""
$EntraDeviceInventoryCsvPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EntraDeviceInventoryCsvPath' -DefaultValue ""
$ActiveUsersCsvPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ActiveUsersCsvPath' -DefaultValue ""
$AdComputersCsvPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AdComputersCsvPath' -DefaultValue ""
$CountrySourceMaxAgeHours = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'CountrySourceMaxAgeHours' -DefaultValue 48)
if ($CountrySourceMaxAgeHours -lt 1) { throw 'CountrySourceMaxAgeHours must be at least 1.' }
$AdComputerNameCountryPrefixes = @{}
$countryPrefixProperty = $script:SmartM365GlobalConfig.PSObject.Properties['AdComputerNameCountryPrefixes']
if ($null -ne $countryPrefixProperty -and $null -ne $countryPrefixProperty.Value) {
    $prefixConfiguration = $countryPrefixProperty.Value
    if ($prefixConfiguration -is [System.Collections.IDictionary]) {
        $prefixEntries = @($prefixConfiguration.GetEnumerator() | ForEach-Object { [pscustomobject]@{ Key = $_.Key; Value = $_.Value } })
    } elseif ($prefixConfiguration -is [pscustomobject]) {
        $prefixEntries = @($prefixConfiguration.PSObject.Properties | ForEach-Object { [pscustomobject]@{ Key = $_.Name; Value = $_.Value } })
    } else {
        throw 'AdComputerNameCountryPrefixes must be a JSON object in the selected tenant profile.'
    }
    foreach ($entry in $prefixEntries) {
        $prefix = ([string]$entry.Key).Trim().ToUpperInvariant()
        $country = ([string]$entry.Value).Trim()
        if ($prefix -cnotmatch '^[A-Z]{2}$' -or [string]::IsNullOrWhiteSpace($country) -or $AdComputerNameCountryPrefixes.ContainsKey($prefix)) {
            throw "Invalid or duplicate AD computer country prefix '$prefix' in the selected tenant profile."
        }
        $AdComputerNameCountryPrefixes[$prefix] = $country
    }
}
if ([string]::IsNullOrWhiteSpace($IntuneDeviceInventoryCsvPath) -and -not [string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) {
    $IntuneDeviceInventoryCsvPath = Join-Path $LatestCsvFolderPath 'Intune_Devices_Inventory.csv'
}
if ([string]::IsNullOrWhiteSpace($EntraDeviceInventoryCsvPath) -and -not [string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) {
    $EntraDeviceInventoryCsvPath = Join-Path $LatestCsvFolderPath 'M365_Entra_Devices.csv'
}
if ([string]::IsNullOrWhiteSpace($ActiveUsersCsvPath) -and -not [string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) {
    $ActiveUsersCsvPath = Join-Path $LatestCsvFolderPath 'M365_Users_Active.csv'
}
if ([string]::IsNullOrWhiteSpace($AdComputersCsvPath) -and -not [string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) {
    $AdComputersCsvPath = Join-Path $LatestCsvFolderPath 'AD_Computers_AllDomains.csv'
}

# ==========================================================
# SharePoint upload
# ==========================================================
$EnableSharePointUpload = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableSharePointUpload' -DefaultValue $false)
$SP_SiteHostname        = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSiteHostname' -DefaultValue "contoso.sharepoint.com"
$SP_SitePath            = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSitePath' -DefaultValue "/sites/workplace-data"
$SP_LibraryDisplayName  = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointLibraryDisplayName' -DefaultValue "Documents"
$SP_TargetFolderPath    = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointTargetFolderPath' -DefaultValue "DATA/CSV"
[long]$SP_ChunkSize     = 10MB

# ==========================================================
# Work folder cleanup
# ==========================================================
$EnableWorkCleanup = $true
$WorkKeepZips      = 10
$WorkKeepExtracts  = 10

# ==========================================================
# Risk summary tuning
# ==========================================================
$RiskIncludeUnknown = $false

# ==========================================================
# Graph mail configuration
# ==========================================================
$From        = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'From' -DefaultValue ""
$To          = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'To' -DefaultValue ""
$Cc          = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Cc' -DefaultValue ""
$ErrorMailTo = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ErrorMailTo' -DefaultValue ""

$SummaryStatePath = Join-Path $ScriptCsvLogFolderPath "Intune_WindowsUpdate_Status.lastsummary.txt"

# ==========================================================
# Console rendering options
# ==========================================================
$ConsolePretty    = $true
$ConsoleShowRunId = $false

# ==========================================================
# Run metadata
# ==========================================================
$ScriptName = "SmartM365-WinUpdate-Status-Inventory"
$RunId      = [guid]::NewGuid().ToString()
$RunStamp   = Get-Date -Format "yyyyMMdd_HHmmss"
$LogPath    = Join-Path $LogsPath "$ScriptName-$RunStamp.log"
$global:LogTextFile = $LogPath

# ==========================================================
# Core helpers
# ==========================================================
function Import-SmartM365CorePreflight {
    if (Get-Command Invoke-CoreSmartM365Preflight -ErrorAction SilentlyContinue) { return }

        $scriptRootValue = Get-Variable -Name ScriptRoot -ValueOnly -ErrorAction SilentlyContinue
    $searchRoot = if ($PSScriptRoot) { $PSScriptRoot } elseif ($scriptRootValue) { $scriptRootValue } elseif ($PSCommandPath) { Split-Path -Path $PSCommandPath -Parent } else { (Get-Location).Path }
    while ($searchRoot) {
        $modulePath = Join-Path -Path $searchRoot -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
        if (Test-Path -LiteralPath $modulePath) {
            Import-Module -Name $modulePath -MinimumVersion '1.0.130' -Prefix Core -ErrorAction Stop
            return
        }

        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }

    throw 'SmartM365.Core module was not found. Preflight checks cannot run.'
}

function Assert-PS7 {
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        throw "PowerShell 7+ is required. Current version: $($PSVersionTable.PSVersion)"
    }
}

function Ensure-Directory {
    param([Parameter(Mandatory=$true)][string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR")][string]$Level = "INFO",
        [Parameter()][string]$Stage = ""
    )

    $normalizedLevel = $Level.ToUpperInvariant()
    if ($normalizedLevel -eq "WARN") {
        $global:SmartM365WarningCount = [int]$global:SmartM365WarningCount + 1
    }
    elseif ($normalizedLevel -eq "ERROR") {
        $global:SmartM365ErrorCount = [int]$global:SmartM365ErrorCount + 1
    }

    $ts = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $fileLine = @([regex]::Split(([string]$Message), '\r?\n') | ForEach-Object { "$ts [$Level][$ScriptName][$RunId] $_" })
    Add-Content -Path $LogPath -Value $fileLine -Encoding UTF8

    if (-not $ConsolePretty) { $fileLine | ForEach-Object { Write-Host $_ }; return }

    $prefix    = if ([string]::IsNullOrWhiteSpace($Stage)) { "" } else { "[$Stage] " }
    $rid       = if ($ConsoleShowRunId) { " ($RunId)" } else { "" }
    $consoleTs = (Get-Date).ToString("HH:mm:ss")
    $line      = "{0} {1}{2}{3}" -f $consoleTs, $prefix, $Message, $rid

    switch ($Level) {
        "INFO"  { Write-Host $line -ForegroundColor Gray }
        "WARN"  { Write-Host $line -ForegroundColor Yellow }
        "ERROR" { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }
}

function Cleanup-TempFiles {
    foreach ($p in @($CsvTemp, $CsvLastTemp)) {
        if ($p -and (Test-Path $p)) {
            try { Remove-Item -LiteralPath $p -Force -ErrorAction Stop } catch { Write-Log "Temporary CSV cleanup failed: $($_.Exception.Message)" 'WARN' 'CLEANUP' }
        }
    }
}

function Prune-Files {
    param(
        [Parameter(Mandatory=$true)][string]$Folder,
        [Parameter(Mandatory=$true)][string]$Filter,
        [Parameter(Mandatory=$true)][int]$Keep
    )

    if (-not (Test-Path $Folder)) { return }
    $files = Get-ChildItem -Path $Folder -File -Filter $Filter -ErrorAction SilentlyContinue |
             Sort-Object -Property LastWriteTime -Descending
    if (-not $files) { return }

    $toRemove = $files | Select-Object -Skip $Keep
    foreach ($f in $toRemove) {
        try {
            Remove-Item -Path $f.FullName -Force -ErrorAction Stop
            Write-Log "Pruned old file: $($f.FullName)" "INFO" "PRUNE"
        }
        catch {
            Write-Log "Failed to prune file: $($f.FullName). Error: $($_.Exception.Message)" "WARN" "PRUNE"
        }
    }
}

function Remove-DetailedArchiveFilesOlderThan {
    param(
        [Parameter(Mandatory=$true)][string]$Folder,
        [Parameter(Mandatory=$true)][string]$BaseNameWithoutExt,
        [ValidateRange(1,3650)][int]$RetentionDays = 7,
        [datetime]$ReferenceTime = (Get-Date)
    )

    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { return }
    $cutoff = $ReferenceTime.AddDays(-$RetentionDays)
    $namePattern = '^' + [regex]::Escape($BaseNameWithoutExt) + '_(?<stamp>\d{8}[-_]\d{6})\.csv$'
    $removedCount = 0
    $removedBytes = [int64]0

    foreach ($file in @(Get-ChildItem -LiteralPath $Folder -File -ErrorAction SilentlyContinue)) {
        if ($file.Name -notmatch $namePattern) { continue }
        $stampText = $Matches['stamp']
        $stamp = [datetime]::MinValue
        $format = if ($stampText.Contains('-')) { 'yyyyMMdd-HHmmss' } else { 'yyyyMMdd_HHmmss' }
        if (-not [datetime]::TryParseExact($stampText, $format, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeLocal, [ref]$stamp)) { continue }
        if ($stamp -ge $cutoff) { continue }

        try {
            $removedBytes += $file.Length
            Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
            $removedCount++
        }
        catch {
            Write-Log "Detailed archive retention could not remove '$($file.FullName)': $($_.Exception.Message)" "WARN" "PRUNE"
        }
    }

    if ($removedCount -gt 0) {
        Write-Log "Detailed archive retention removed $removedCount file(s), $removedBytes byte(s), older than $($cutoff.ToString('o')) from $Folder." "INFO" "PRUNE"
    }
}

# ==========================================================
# Work cleanup
# ==========================================================
function Prune-WorkZips {
    param([Parameter(Mandatory=$true)][string]$WorkPath,[Parameter(Mandatory=$true)][int]$Keep)

    if (-not (Test-Path $WorkPath)) { return }
    $zips = Get-ChildItem -Path $WorkPath -File -Filter "*.zip" -ErrorAction SilentlyContinue |
            Sort-Object -Property LastWriteTime -Descending
    if (-not $zips) { return }

    $toRemove = $zips | Select-Object -Skip $Keep
    foreach ($z in $toRemove) {
        try {
            Remove-Item -Path $z.FullName -Force -ErrorAction Stop
            Write-Log "Work cleanup: removed ZIP $($z.FullName)" "INFO" "CLEANUP"
        }
        catch {
            Write-Log "Work cleanup: failed to remove ZIP $($z.FullName). Error: $($_.Exception.Message)" "WARN" "CLEANUP"
        }
    }
}

function Prune-WorkExtractFolders {
    param([Parameter(Mandatory=$true)][string]$WorkPath,[Parameter(Mandatory=$true)][int]$Keep)

    if (-not (Test-Path $WorkPath)) { return }
    $folders = Get-ChildItem -Path $WorkPath -Directory -ErrorAction SilentlyContinue |
               Where-Object { $_.Name -like "Extract_*" } |
               Sort-Object -Property LastWriteTime -Descending
    if (-not $folders) { return }

    $toRemove = $folders | Select-Object -Skip $Keep
    foreach ($f in $toRemove) {
        try {
            Remove-Item -Path $f.FullName -Recurse -Force -ErrorAction Stop
            Write-Log "Work cleanup: removed folder $($f.FullName)" "INFO" "CLEANUP"
        }
        catch {
            Write-Log "Work cleanup: failed to remove folder $($f.FullName). Error: $($_.Exception.Message)" "WARN" "CLEANUP"
        }
    }
}

function Cleanup-WorkFolder {
    param([Parameter(Mandatory=$true)][string]$WorkPath,[Parameter(Mandatory=$true)][int]$KeepZips,[Parameter(Mandatory=$true)][int]$KeepExtracts)

    if (-not $EnableWorkCleanup) { Write-Log "Work cleanup disabled." "INFO" "CLEANUP"; return }
    Prune-WorkZips -WorkPath $WorkPath -Keep $KeepZips
    Prune-WorkExtractFolders -WorkPath $WorkPath -Keep $KeepExtracts
}

# ==========================================================
function Send-FatalErrorEmail {
    param([Parameter(Mandatory=$true)][string]$Subject,[Parameter(Mandatory=$true)][string]$HtmlBody)

    if (-not $EnableErrorEmail) { return }
    if ($DryRun) { Write-Log "DryRun: skipping fatal error email." "INFO" "DRYRUN"; return }
    if ([string]::IsNullOrWhiteSpace($ErrorMailTo)) { return }

    try {
        Import-SmartM365CorePreflight
        $mailParams = @{ MailPurpose = 'Error'; To = $ErrorMailTo; Subject = $Subject; BodyHtml = $HtmlBody }
        if (-not [string]::IsNullOrWhiteSpace($From)) { $mailParams['From'] = $From }
        Send-CoreSmartM365Mail @mailParams
        Write-Log "Fatal error email sent." "INFO" "MAIL"
    }
    catch {
        Write-Log "Failed to send fatal error email: $($_.Exception.Message)" "WARN" "MAIL"
    }
}

function Send-SummaryEmail {
    param([Parameter(Mandatory=$true)][string]$Subject,[Parameter(Mandatory=$true)][string]$HtmlBody)

    if (-not $EnableSummaryEmail) { return $false }
    if ($DryRun) { Write-Log "DryRun: skipping summary email." "INFO" "DRYRUN"; return $false }
    $summaryTo = $To
    if ([string]::IsNullOrWhiteSpace($summaryTo)) {
        Write-Log "Summary email skipped because To is not configured in local or global configuration. Inventory outputs remain valid." "WARN" "MAIL"
        return $false
    }

    try {
        Import-SmartM365CorePreflight
        $mailParams = @{ To = $summaryTo; Subject = $Subject; BodyHtml = $HtmlBody; ReportCadenceKey = 'Intune.WindowsUpdate' }
        if (-not [string]::IsNullOrWhiteSpace($From)) { $mailParams['From'] = $From }
        if (-not [string]::IsNullOrWhiteSpace($Cc)) { $mailParams['Cc'] = $Cc }
        Send-CoreSmartM365Mail @mailParams
        if ((Get-Variable SmartM365LastMailDelivery -Scope Global -ValueOnly -ErrorAction SilentlyContinue) -in @('Skipped','Busy')) { return $false }
        Write-Log "Summary email sent." "INFO" "MAIL"
        return $true
    }
    catch {
        Write-Log "Failed to send summary email: $($_.Exception.Message)" "ERROR" "MAIL"
        throw
    }
}

function Should-SendSummaryEmail {
    param([Parameter(Mandatory=$true)][string]$CurrentState,[Parameter(Mandatory=$true)][string]$Mode,[Parameter(Mandatory=$true)][string]$StatePath)

    if (-not $EnableSummaryEmail) { return $false }
    switch ($Mode) {
        "Never"    { return $false }
        "Always"   { return $true }
        "OnChange" {
            $previous = if (Test-Path -LiteralPath $StatePath) { [string](Get-Content -LiteralPath $StatePath -Raw -ErrorAction SilentlyContinue) } else { '' }
            return (-not [string]::Equals($previous.Trim(),$CurrentState,[StringComparison]::Ordinal))
        }
        default { return $false }
    }
}

function Save-SummaryState {
    param([Parameter(Mandatory=$true)][string]$State,[Parameter(Mandatory=$true)][string]$StatePath)
    if ($DryRun) { return }
    try { Set-Content -LiteralPath $StatePath -Value $State -Encoding UTF8 } catch { }
}

# ==========================================================
# Auth helpers
# ==========================================================
function Get-AppCertificate {
    param([Parameter(Mandatory=$true)][string]$Thumbprint)

    $thumb = $Thumbprint.Replace(" ", "").ToUpper()
    $cert = Get-Item -Path "Cert:\CurrentUser\My\$thumb" -ErrorAction SilentlyContinue
    if (-not $cert) { $cert = Get-Item -Path "Cert:\LocalMachine\My\$thumb" -ErrorAction SilentlyContinue }
    if (-not $cert) { throw "Certificate not found for thumbprint: $thumb" }
    if (-not $cert.HasPrivateKey) { throw "Certificate has no private key: $thumb" }
    return $cert
}

function Ensure-MsalPsModule {
    param(
        [string]$ModuleName = "MSAL.PS",
        [ValidateSet("CurrentUser","AllUsers")]
        [string]$InstallScope = "AllUsers",
        [string]$Repository = "PSGallery"
    )

    $availableModule = Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1
    if ($null -eq $availableModule) {
        $installCommand = "Install-Module $ModuleName -Scope $InstallScope -Repository $Repository -Force -AllowClobber"
        Write-Log "$ModuleName module is not installed. Attempting automatic installation: $installCommand" "WARN" "AUTH"

        if (-not (Get-Command -Name Install-Module -ErrorAction SilentlyContinue)) {
            throw "$ModuleName module is not installed and Install-Module is unavailable. Install PowerShellGet, then run in an elevated PowerShell session: $installCommand"
        }

        try {
            Install-Module -Name $ModuleName -Scope $InstallScope -Repository $Repository -Force -AllowClobber -ErrorAction Stop
            $availableModule = Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1
            if ($null -eq $availableModule) {
                throw "Install-Module completed but $ModuleName was not found in PSModulePath."
            }
            Write-Log "$ModuleName module installed: Version=$($availableModule.Version); Path=$($availableModule.Path)" "INFO" "AUTH"
        }
        catch {
            throw "$ModuleName module is not installed and automatic installation failed. Run PowerShell as administrator and execute: $installCommand. Error: $($_.Exception.Message)"
        }
    }

    try {
        Import-Module -Name $ModuleName -ErrorAction Stop
        Write-Log "$ModuleName module loaded: Version=$($availableModule.Version); Path=$($availableModule.Path)" "INFO" "AUTH"
    }
    catch {
        throw "$ModuleName module was found but could not be imported. Error: $($_.Exception.Message)"
    }
}

function Get-GraphTokenWithCert {
    param(
        [Parameter(Mandatory=$true)][string]$TenantId,
        [Parameter(Mandatory=$true)][string]$ClientId,
        [Parameter(Mandatory=$true)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    Ensure-MsalPsModule -InstallScope "AllUsers" -Repository "PSGallery"
    $token = Get-MsalToken -TenantId $TenantId -ClientId $ClientId -ClientCertificate $Certificate -Scopes "https://graph.microsoft.com/.default"
    if (-not $token -or [string]::IsNullOrWhiteSpace($token.AccessToken)) { throw "Failed to acquire access token using MSAL.PS." }
    return $token.AccessToken
}
# ==========================================================
# Graph token auto-refresh
# ==========================================================
function Refresh-GraphTokenIfNeeded {
    param(
        [Parameter(Mandatory=$true)][string]$TenantId,
        [Parameter(Mandatory=$true)][string]$ClientId,
        [Parameter(Mandatory=$true)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate,
        [Parameter(Mandatory=$true)][ref]$Headers,
        [Parameter()][int]$RefreshThresholdMinutes = 55
    )

    if ($script:TokenAcquiredAt -eq [datetime]::MinValue) { return }

    $elapsed = (Get-Date) - $script:TokenAcquiredAt
    if ($elapsed.TotalMinutes -lt $RefreshThresholdMinutes) { return }

    Write-Log "Graph token age $([int]$elapsed.TotalMinutes) min >= threshold $RefreshThresholdMinutes min. Refreshing..." "INFO" "AUTH"
    try {
        $newToken = Get-GraphTokenWithCert -TenantId $TenantId -ClientId $ClientId -Certificate $Certificate
        $Headers.Value = @{ Authorization = "Bearer $newToken" }
        $script:TokenAcquiredAt = Get-Date
        Write-Log "Graph token refreshed successfully at $($script:TokenAcquiredAt.ToString('HH:mm:ss'))." "INFO" "AUTH"
    }
    catch {
        Write-Log "Failed to refresh Graph token: $($_.Exception.Message). Continuing with existing token." "WARN" "AUTH"
    }
}

# ==========================================================
# Graph helpers
# ==========================================================
function Try-GetGraphErrorMessage {
    param([Parameter(Mandatory=$true)]$ErrorRecord)

    $bodyText = $null
    try {
        if ($ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.GetResponseStream()) {
            $sr = New-Object System.IO.StreamReader($ErrorRecord.Exception.Response.GetResponseStream())
            $bodyText = $sr.ReadToEnd()
        }
    } catch { }
    if ([string]::IsNullOrWhiteSpace($bodyText)) { return $null }
    try {
        $obj = $bodyText | ConvertFrom-Json -ErrorAction Stop
        if ($obj.error -and $obj.error.message) { return [string]$obj.error.message }
    } catch { }
    return $bodyText
}

function Get-GraphHttpStatusCode {
    param([Parameter(Mandatory=$true)]$ErrorRecord)

    return Get-CoreSmartM365GraphErrorStatusCode -ErrorRecord $ErrorRecord
}

function Get-GraphRetryDelaySeconds {
    param(
        [Parameter(Mandatory=$true)]$ErrorRecord,
        [Parameter(Mandatory=$true)][int]$Attempt
    )

    # Core supplies Retry-After parsing and jitter. Keep the existing polling
    # floor (15/30/60/120/240/300s); equal jitter must not shorten it.
    $wait = Get-CoreSmartM365RetryDelay -Attempt $Attempt -ErrorRecord $ErrorRecord -BaseSeconds 30 -MaximumSeconds 300
    $pollingFloor = [math]::Min(300, 15 * [math]::Pow(2, [math]::Min(20, [math]::Max(0, $Attempt - 1))))
    return [int][math]::Max($pollingFloor, $wait.Seconds)
}

function Test-GraphTransientError {
    param([Parameter(Mandatory=$true)]$ErrorRecord)

    return (Test-CoreSmartM365GraphRetryableError -ErrorRecord $ErrorRecord).Retry
}

function Invoke-GraphRestMethodWithRetry {
    param(
        [Parameter(Mandatory=$true)][ValidateSet('GET','POST')][string]$Method,
        [Parameter(Mandatory=$true)][string]$Uri,
        [Parameter(Mandatory=$true)][hashtable]$Headers,
        [Parameter()][object]$Body,
        [Parameter()][string]$ContentType,
        [Parameter()][int]$MaxAttempts = 8,
        [Parameter()][string]$Operation = 'Graph request'
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            $params = @{ Method = $Method; Uri = $Uri; Headers = $Headers }
            if ($null -ne $Body) { $params['Body'] = $Body }
            if (-not [string]::IsNullOrWhiteSpace($ContentType)) { $params['ContentType'] = $ContentType }
            return Invoke-RestMethod @params
        }
        catch {
            if ($attempt -ge $MaxAttempts -or -not (Test-GraphTransientError -ErrorRecord $_)) { throw }
            $statusCode = Get-GraphHttpStatusCode -ErrorRecord $_
            $delay = Get-GraphRetryDelaySeconds -ErrorRecord $_ -Attempt $attempt
            Write-Log "$Operation transient Graph failure. Status=$statusCode; attempt $attempt/$MaxAttempts; waiting ${delay}s." "WARN" "GRAPH"
            Start-Sleep -Seconds $delay
        }
    }
}

function Invoke-GraphGetAllPages {
    param([Parameter(Mandatory=$true)][string]$Uri,[Parameter(Mandatory=$true)][hashtable]$Headers)

    $all = New-Object System.Collections.Generic.List[object]
    $visited = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $next = $Uri
    $pageNumber = 0
    while ($next) {
        if (-not $visited.Add([string]$next)) { throw 'Windows Update Graph pagination returned a repeated @odata.nextLink; collection is incomplete.' }
        $pageNumber++
        $resp = Invoke-GraphRestMethodWithRetry -Method GET -Uri $next -Headers $Headers -Operation "Graph paging GET"
        if ($null -eq $resp -or $null -eq $resp.PSObject.Properties['value']) { throw "Windows Update Graph page $pageNumber returned an invalid collection response without a value property." }
        if ($resp.value) { $resp.value | ForEach-Object { $all.Add($_) | Out-Null } }
        $next = $resp.'@odata.nextLink'
    }
    return $all
}
function Resolve-WinUpdateMissingDeviceNames {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][hashtable]$Headers
    )

    $missingRows = @($Rows | Where-Object { [string]::IsNullOrWhiteSpace([string]$_.DeviceName) })
    if ($missingRows.Count -eq 0) { return }

    Write-Log "Graph report contains $($missingRows.Count) row(s) with an empty DeviceName; resolving unique managed devices by DeviceId." "INFO" "DATA"
    $resolvedRows = 0
    $fallbackRows = 0
    $missingByDeviceId = @($missingRows | Group-Object -Property { [string]$_.DeviceId })

    foreach ($deviceGroup in $missingByDeviceId) {
        $deviceId = [string]$deviceGroup.Name
        $resolvedName = ''
        $resolvedUpn = ''
        $lookupError = ''

        if (-not [string]::IsNullOrWhiteSpace($deviceId)) {
            try {
                $escapedDeviceId = [uri]::EscapeDataString($deviceId)
                $managedDevice = Invoke-GraphRestMethodWithRetry `
                    -Method GET `
                    -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$escapedDeviceId`?`$select=id,deviceName,userPrincipalName" `
                    -Headers $Headers `
                    -Operation "Resolve managed device name"
                $resolvedName = [string]$managedDevice.deviceName
                $resolvedUpn = [string]$managedDevice.userPrincipalName
            }
            catch {
                $lookupError = $_.Exception.Message
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($resolvedName)) {
            foreach ($row in $deviceGroup.Group) {
                $row.DeviceName = $resolvedName
                if ($row.PSObject.Properties['UPN'] -and [string]::IsNullOrWhiteSpace([string]$row.UPN) -and -not [string]::IsNullOrWhiteSpace($resolvedUpn)) {
                    $row.UPN = $resolvedUpn
                }
                $resolvedRows++
            }
            continue
        }

        foreach ($row in $deviceGroup.Group) {
            if (-not [string]::IsNullOrWhiteSpace($deviceId)) {
                $row.DeviceName = "UnknownDevice-$deviceId"
            }
            $fallbackRows++
        }
        $failureDetail = if ([string]::IsNullOrWhiteSpace($lookupError)) { 'Graph returned no deviceName.' } else { $lookupError }
        Write-Log "Managed device name could not be resolved for DeviceId '$deviceId'; using a deterministic fallback where possible. Error: $failureDetail" "WARN" "DATA"
    }

    Write-Log "Missing DeviceName resolution completed: ResolvedRows=$resolvedRows FallbackRows=$fallbackRows." "INFO" "DATA"
}

function Assert-WinUpdateOutputQuality {
    [CmdletBinding()]
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows)

    if ($Rows.Count -eq 0) { throw 'Windows Update output quality gate failed: the canonical dataset is empty.' }

    $blankPolicyId = 0
    $blankDeviceId = 0
    $blankDeviceName = 0
    $duplicatePolicyDevice = 0
    $keys = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($row in $Rows) {
        $policyId = [string]$row.PolicyId
        $deviceId = [string]$row.DeviceId
        $deviceName = [string]$row.DeviceName
        if ([string]::IsNullOrWhiteSpace($policyId)) { $blankPolicyId++ }
        if ([string]::IsNullOrWhiteSpace($deviceId)) { $blankDeviceId++ }
        if ([string]::IsNullOrWhiteSpace($deviceName)) { $blankDeviceName++ }
        if (-not [string]::IsNullOrWhiteSpace($policyId) -and -not [string]::IsNullOrWhiteSpace($deviceId)) {
            if (-not $keys.Add("$($policyId.Trim())|$($deviceId.Trim())")) { $duplicatePolicyDevice++ }
        }
    }

    if ($blankPolicyId -gt 0 -or $blankDeviceId -gt 0 -or $blankDeviceName -gt 0 -or $duplicatePolicyDevice -gt 0) {
        throw "Windows Update output quality gate failed: BlankPolicyId=$blankPolicyId BlankDeviceId=$blankDeviceId BlankDeviceName=$blankDeviceName DuplicatePolicyDeviceRows=$duplicatePolicyDevice."
    }

    Write-Log "Output quality gate passed: Rows=$($Rows.Count) BlankPolicyId=0 BlankDeviceId=0 BlankDeviceName=0 DuplicatePolicyDeviceRows=0." "INFO" "DATA"
}

# ==========================================================
# Feature Update policies discovery
# ==========================================================
function Get-FeatureUpdatePolicies {
    param([Parameter(Mandatory=$true)][hashtable]$Headers)

    $uri = "https://graph.microsoft.com/beta/deviceManagement/windowsFeatureUpdateProfiles?`$select=id,displayName"
    Write-Log "Discovering Feature Update policies..." "INFO" "GRAPH"
    $items = Invoke-GraphGetAllPages -Uri $uri -Headers $Headers
    return $items | ForEach-Object { [pscustomobject]@{ PolicyId=$_.id; PolicyName=$_.displayName } }
}

# ==========================================================
# Intune Reports exportJobs helpers
# ==========================================================
function Start-IntuneExportJobRaw {
    param(
        [Parameter(Mandatory=$true)][string]$ReportName,
        [Parameter(Mandatory=$true)][hashtable]$Headers,
        [Parameter(Mandatory=$true)][string]$Filter,
        [Parameter()][string[]]$Select = $null,
        [Parameter()][string]$LocalizationType = $null
    )

    $uri     = "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs"
    $payload = @{ reportName=$ReportName; format="csv"; filter=$Filter }
    if ($Select -and $Select.Count -gt 0) { $payload.select = $Select }
    if (-not [string]::IsNullOrWhiteSpace($LocalizationType)) { $payload.localizationType = $LocalizationType }

    $selectCount = if ($Select) { $Select.Count } else { 0 }
    Write-Log "Starting export job: reportName=$ReportName filter=$Filter selectCount=$selectCount" "INFO" "GRAPH"

    $resp = Invoke-GraphRestMethodWithRetry -Method POST -Uri $uri -Headers $Headers -Body ($payload | ConvertTo-Json -Depth 10) -ContentType "application/json" -Operation "Start Intune export job"
    if (-not $resp -or [string]::IsNullOrWhiteSpace($resp.id)) { throw "exportJobs POST returned no job id." }
    return $resp.id
}

function Get-IntuneExportJob {
    param([Parameter(Mandatory=$true)][string]$JobId,[Parameter(Mandatory=$true)][hashtable]$Headers,[Parameter()][int]$MaxAttempts = 8)
    return Invoke-GraphRestMethodWithRetry -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs('$JobId')" -Headers $Headers -MaxAttempts $MaxAttempts -Operation "Get Intune export job"
}

function Wait-IntuneExportJobComplete {
    param(
        [Parameter(Mandatory=$true)][string]$JobId,
        [Parameter(Mandatory=$true)][hashtable]$Headers,
        [Parameter()][int]$TimeoutSeconds = 900,
        [Parameter()][int]$PollSeconds = 5,
        [Parameter()][int]$PollMaxAttempts = 8
    )

    $start = Get-Date
    while ($true) {
        $job    = Get-IntuneExportJob -JobId $JobId -Headers $Headers -MaxAttempts $PollMaxAttempts
        $status = "$($job.status)".ToLowerInvariant()
        Write-Log "Export job status: $status" "INFO" "GRAPH"

        if ($status -eq "completed" -or $status -eq "complete") {
            if ([string]::IsNullOrWhiteSpace($job.url)) { throw "Export job completed but download url is empty." }
            return $job
        }
        if ($status -eq "failed") {
            $errJson = $null; try { $errJson = ($job.error | ConvertTo-Json -Depth 10) } catch { }
            throw "Export job failed. Error=$errJson"
        }
        if (((Get-Date) - $start).TotalSeconds -gt $TimeoutSeconds) { throw "Export job timeout after $TimeoutSeconds seconds." }
        Start-Sleep -Seconds $PollSeconds
    }
}

function Download-ExportZip {
    param(
        [Parameter(Mandatory=$true)][string]$DownloadUrl,
        [Parameter(Mandatory=$true)][string]$DestinationZipPath,
        [Parameter()][int]$MaxAttempts = 6
    )

    Write-Log "Downloading export ZIP..." "INFO" "DATA"
    Invoke-CoreSmartM365ExportFileDownload -Url $DownloadUrl -OutFile $DestinationZipPath -MaxAttempts $MaxAttempts -Operation 'Windows Update export ZIP download' | Out-Null
}

function Extract-FirstCsvFromZip {
    param([Parameter(Mandatory=$true)][string]$ZipPath,[Parameter(Mandatory=$true)][string]$ExtractFolder)

    Ensure-Directory -Path $ExtractFolder
    Get-ChildItem -Path $ExtractFolder -Force -ErrorAction SilentlyContinue | ForEach-Object {
        try { Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction Stop } catch { }
    }
    Expand-Archive -Path $ZipPath -DestinationPath $ExtractFolder -Force
    $csv = Get-ChildItem -Path $ExtractFolder -Recurse -File -Filter "*.csv" -ErrorAction SilentlyContinue | Select-Object -First 1
    if (-not $csv) { throw "No CSV found inside ZIP: $ZipPath" }
    return $csv.FullName
}

# ==========================================================
# Atomic copy helpers
# ==========================================================
function Copy-FileAtomic {
    param([Parameter(Mandatory=$true)][string]$SourcePath,[Parameter(Mandatory=$true)][string]$DestinationFinal,[Parameter(Mandatory=$true)][string]$DestinationTemp)

    if (-not (Test-Path $SourcePath)) { throw "Source file not found: $SourcePath" }
    Ensure-Directory -Path (Split-Path -Parent $DestinationFinal)
    if (Test-Path $DestinationTemp) { Remove-Item -Path $DestinationTemp -Force -ErrorAction SilentlyContinue }
    Copy-Item -Path $SourcePath -Destination $DestinationTemp -Force
    Move-Item -Path $DestinationTemp -Destination $DestinationFinal -Force
}

function Write-ArchiveCopyAndPrune {
    param(
        [Parameter(Mandatory=$true)][string]$SourceCsv,
        [Parameter(Mandatory=$true)][string]$ArchiveFolder,
        [Parameter(Mandatory=$true)][string]$BaseNameWithoutExt,
        [Parameter(Mandatory=$true)][string]$RunStamp,
        [ValidateRange(1,3650)][int]$RetentionDays = 7
    )

    Ensure-Directory -Path $ArchiveFolder
    $archiveFinal = Join-Path $ArchiveFolder ("{0}_{1}.csv"     -f $BaseNameWithoutExt, $RunStamp)
    $archiveTemp  = Join-Path $ArchiveFolder ("{0}_{1}.csv.tmp" -f $BaseNameWithoutExt, $RunStamp)
    Copy-FileAtomic -SourcePath $SourceCsv -DestinationFinal $archiveFinal -DestinationTemp $archiveTemp
    Write-Log "Archive CSV created: $archiveFinal" "INFO" "DATA"
    Remove-DetailedArchiveFilesOlderThan -Folder $ArchiveFolder -BaseNameWithoutExt $BaseNameWithoutExt -RetentionDays $RetentionDays
}

# ==========================================================
# HTML helpers
# ==========================================================




# ==========================================================
# Normalization + Readiness lookup
# ==========================================================
function Normalize-DeviceName {
    param([Parameter(Mandatory=$false)][string]$DeviceName)
    if ([string]::IsNullOrWhiteSpace($DeviceName)) { return $null }
    $n = $DeviceName.Trim().ToLowerInvariant()
    if ($n -match "^([^\.]+)\.") { $n = $Matches[1] }
    $n = -join ($n.ToCharArray() | Where-Object { $_ -match "[a-z0-9\-]" })
    return $n
}

function Build-ReadinessLookup {
    param([Parameter(Mandatory=$true)][object[]]$ReadinessRows)

    $lookup = @{}
    foreach ($r in $ReadinessRows) {
        $key = $r.NormalizedDeviceName
        if ([string]::IsNullOrWhiteSpace($key)) { $key = Normalize-DeviceName -DeviceName $r.DeviceName }
        if ([string]::IsNullOrWhiteSpace($key)) { continue }
        if (-not $lookup.ContainsKey($key)) { $lookup[$key] = $r }
    }
    return $lookup
}

function Add-WinUpdateEntraOsVersionFallback {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$IntuneDeviceRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$EntraDeviceRows
    )

    $intuneBridge = @{}
    $ambiguousManagedIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($intuneRow in $IntuneDeviceRows) {
        $managedDeviceId = [string]$intuneRow.'Device ID'
        $azureAdDeviceId = [string]$intuneRow.'Azure AD Device ID'
        if ([string]::IsNullOrWhiteSpace($managedDeviceId) -or [string]::IsNullOrWhiteSpace($azureAdDeviceId)) { continue }
        $managedKey = $managedDeviceId.Trim().ToLowerInvariant()
        $azureAdKey = $azureAdDeviceId.Trim().ToLowerInvariant()
        if ($ambiguousManagedIds.Contains($managedKey)) { continue }
        if ($intuneBridge.ContainsKey($managedKey) -and $intuneBridge[$managedKey] -ne $azureAdKey) {
            [void]$intuneBridge.Remove($managedKey)
            [void]$ambiguousManagedIds.Add($managedKey)
            continue
        }
        $intuneBridge[$managedKey] = $azureAdKey
    }

    $entraVersionByDeviceId = @{}
    $entraBuildByDeviceId = @{}
    foreach ($entraRow in $EntraDeviceRows) {
        $entraDeviceId = [string]$entraRow.DeviceId
        $entraOsVersion = [string]$entraRow.OperatingSystemVersion
        if ([string]::IsNullOrWhiteSpace($entraDeviceId) -or [string]::IsNullOrWhiteSpace($entraOsVersion)) { continue }
        $entraBuild = Get-WinUpdateOsBuild -Value $entraOsVersion
        if ($null -eq $entraBuild) { continue }
        $entraKey = $entraDeviceId.Trim().ToLowerInvariant()
        if (-not $entraBuildByDeviceId.ContainsKey($entraKey) -or $entraBuild -gt $entraBuildByDeviceId[$entraKey]) {
            $entraBuildByDeviceId[$entraKey] = $entraBuild
            $entraVersionByDeviceId[$entraKey] = $entraOsVersion.Trim()
        }
    }

    $output = New-Object System.Collections.Generic.List[object]
    $existingVersionRows = 0
    $entraVersionRows = 0
    $unavailableRows = 0
    foreach ($row in $Rows) {
        $o = [ordered]@{}
        foreach ($property in $row.PSObject.Properties) { $o[$property.Name] = $property.Value }

        $currentOsVersion = [string]$row.OSVersion
        if (-not [string]::IsNullOrWhiteSpace($currentOsVersion)) {
            $existingVersionRows++
            $o['OSVersionSource'] = if ([string]$row.ReadinessMatch -eq 'Matched') { 'Readiness' } else { 'Existing' }
        }
        else {
            $managedDeviceId = [string]$row.DeviceId
            $entraOsVersion = ''
            if (-not [string]::IsNullOrWhiteSpace($managedDeviceId)) {
                $managedKey = $managedDeviceId.Trim().ToLowerInvariant()
                if ($intuneBridge.ContainsKey($managedKey)) {
                    $entraKey = [string]$intuneBridge[$managedKey]
                    if ($entraVersionByDeviceId.ContainsKey($entraKey)) {
                        $entraOsVersion = [string]$entraVersionByDeviceId[$entraKey]
                    }
                }
            }
            if (-not [string]::IsNullOrWhiteSpace($entraOsVersion)) {
                $o['OSVersion'] = $entraOsVersion
                $o['OSVersionSource'] = 'Entra'
                $entraVersionRows++
            }
            else {
                $o['OSVersionSource'] = 'Unavailable'
                $unavailableRows++
            }
        }
        $output.Add([pscustomobject]$o) | Out-Null
    }

    [pscustomobject][ordered]@{
        Rows = $output.ToArray()
        ExistingVersionRows = $existingVersionRows
        EntraVersionRows = $entraVersionRows
        UnavailableRows = $unavailableRows
        IntuneBridgeKeys = $intuneBridge.Count
        AmbiguousIntuneBridgeKeys = $ambiguousManagedIds.Count
        EntraVersionKeys = $entraVersionByDeviceId.Count
    }
}

# ==========================================================
# Risk + Blocking + Remediation
# ==========================================================
function Get-RiskBucket {
    param(
        [Parameter(Mandatory=$false)][string]$AggregateState,
        [Parameter(Mandatory=$true)][bool]$IncludeUnknown
    )

    if ([string]::IsNullOrWhiteSpace($AggregateState)) { return "Unknown" }

    switch ($AggregateState) {
        "Failed"         { return "High" }
        "Error"          { return "High" }
        "Blocked"        { return "High" }
        "InProgress"     { return "Medium" }
        "Completed"      { return "Low" }
        "Success"        { return "Low" }
        "Compliant"      { return "Low" }
        "UpToDate"       { return "Low" }
        "NotApplicable"  { return "Unknown" }
        "Unknown"        { return $(if ($IncludeUnknown) { "Medium" } else { "Unknown" }) }
        default {
            $s = $AggregateState.ToLowerInvariant()
            if ($s -match "fail|error|block|rollback|uninstall")       { return "High" }
            if ($s -match "progress|install|offer|pending|reboot|wait|download") { return "Medium" }
            if ($s -match "complete|success|compliant|up")             { return "Low" }
            if ($s -match "unknown") { return $(if ($IncludeUnknown) { "Medium" } else { "Unknown" }) }
            return "Unknown"
        }
    }
}

function Get-BlockingReason {
    param(
        [string]$AggregateState,
        [string]$LatestAlertMessageLoc
    )

    $state = $(if ($null -eq $AggregateState) { "" } else { $AggregateState }).ToLowerInvariant()
    $alert = "$LatestAlertMessageLoc"

    if ($state -match "rollback|uninstall") {
        return "HardFailure_Rollback"
    }
    if ($state -match "fail|error|blocked") {
        if ($alert -match "safeguard|hold")                { return "HardFailure_SafeguardHold" }
        if ($alert -match "disk|storage|space")            { return "HardFailure_DiskSpace" }
        if ($alert -match "compatib|driver|firmware|bios") { return "HardFailure_Compatibility" }
        return "HardFailure_Generic"
    }
    if ($alert -and $alert -ne "0" -and $alert -notmatch "^not applicable$")  { return "UpdateAlert" }
    if ($state -match "inprogress|install|offering|pending|waiting|download")  { return "DeploymentInProgress" }
    if ($state -match "cancelled|notapplicable|unknown")                        { return "NotApplicableOrUnknown" }
    return "CompliantOrCompleted"
}

function Get-RemediationAction {
    param(
        [Parameter(Mandatory=$true)][string]$BlockingReason,
        [Parameter(Mandatory=$false)][string]$AggregateState,
        [Parameter(Mandatory=$false)][string]$UpgradeEligibilityLabel
    )

    if ($BlockingReason -eq "CompliantOrCompleted") {
        return @{ Priority=""; ActionCode=""; ActionDescription=""; Owner="" }
    }

    if ($UpgradeEligibilityLabel -eq "Eligible" -and $BlockingReason -like "HardFailure*") {
        return @{ Priority=0; ActionCode="ELIGIBLE_DEVICE_BLOCKED"; ActionDescription="Windows 11 eligible device blocked during update. Immediate investigation required."; Owner="Workplace" }
    }

    if ($UpgradeEligibilityLabel -in @("NotEligible","Unknown","NotApplicableOrUnknown") -and
        $BlockingReason -in @("DeploymentInProgress","UpdateAlert")) {
        return @{ Priority=2; ActionCode="FIX_READINESS_FIRST"; ActionDescription="Device not eligible for Windows 11. Fix readiness before continuing update rollout."; Owner="Workplace" }
    }

    switch ($BlockingReason) {
        "HardFailure_DiskSpace"      { return @{ Priority=1; ActionCode="FREE_DISK_SPACE";      ActionDescription="Free disk space (cleanup, remove large files/apps) and retry update."; Owner="Workplace" } }
        "HardFailure_Generic"        { return @{ Priority=1; ActionCode="CHECK_ERROR";           ActionDescription="Investigate update failure (logs, compatibility, disk, servicing stack) and remediate."; Owner="Workplace" } }
        "HardFailure_Rollback"       { return @{ Priority=1; ActionCode="INVESTIGATE_ROLLBACK";  ActionDescription="Investigate the feature update rollback or uninstall, remediate the root cause, and retry when safe."; Owner="Workplace" } }
        "HardFailure_Compatibility"  { return @{ Priority=2; ActionCode="COMPATIBILITY_REVIEW";  ActionDescription="Compatibility issue suspected (driver/firmware/app). Review and remediate before retry."; Owner="Workplace" } }
        "HardFailure_SafeguardHold"  { return @{ Priority=2; ActionCode="SAFEGUARD_HOLD";        ActionDescription="Safeguard hold detected. Validate known issue and consider mitigation or wait for Microsoft resolution."; Owner="Workplace" } }
        "UpdateAlert"                { return @{ Priority=3; ActionCode="HANDLE_ALERT";          ActionDescription="Review Intune alert details and apply recommended fix."; Owner="Workplace" } }
        "DeploymentInProgress"       { return @{ Priority=4; ActionCode="WAIT_MONITOR";          ActionDescription="Deployment in progress. Monitor status, verify device online and update service health."; Owner="N1" } }
        "NotApplicableOrUnknown"     { return @{ Priority=5; ActionCode="VERIFY_SCOPE";          ActionDescription="Verify policy scope/eligibility and device prerequisites."; Owner="N2" } }
        default                      { return @{ Priority=9; ActionCode="REVIEW";                ActionDescription="Manual review required."; Owner="Workplace" } }
    }
}

# ==========================================================
# Operational summary helpers
# ==========================================================
function Get-WinUpdateRowValue {
    param([AllowNull()][object]$Row,[Parameter(Mandatory)][string]$Name)
    if ($null -eq $Row) { return $null }
    $property = $Row.PSObject.Properties[$Name]
    if ($property) { return $property.Value }
    return $null
}

function Get-WinUpdateOperationalState {
    param([Parameter(Mandatory)][object]$Row)
    $aggregateState = [string](Get-WinUpdateRowValue -Row $Row -Name 'AggregateState')
    $currentStatus = [string](Get-WinUpdateRowValue -Row $Row -Name 'CurrentDeviceUpdateStatus_loc')
    if ("$aggregateState|$currentStatus" -match '(?i)rollback|uninstall') { return 'ActionRequired' }
    $blockingReason = [string](Get-WinUpdateRowValue -Row $Row -Name 'BlockingReason')
    if ($blockingReason -like 'HardFailure*' -or $blockingReason -eq 'UpdateAlert') { return 'ActionRequired' }
    if ($blockingReason -eq 'DeploymentInProgress') { return 'InProgress' }
    if ($blockingReason -eq 'NotApplicableOrUnknown') { return 'Unknown' }
    if ($blockingReason -eq 'CompliantOrCompleted') { return 'Completed' }

    switch ([string](Get-WinUpdateRowValue -Row $Row -Name 'RiskBucket')) {
        'High'   { return 'ActionRequired' }
        'Medium' { return 'InProgress' }
        'Low'    { return 'Completed' }
        default  { return 'Unknown' }
    }
}

function Get-WinUpdateOsBuild {
    param([AllowNull()][object]$Value)

    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $null }
    if ($text -match '^\s*\d+\.\d+\.(?<Build>\d+)(?:\.|\s|$)') {
        return [int]$matches.Build
    }
    return $null
}

function Get-WinUpdatePolicyTarget {
    param([AllowNull()][string]$PolicyName)

    $name = [string]$PolicyName
    if ([string]::IsNullOrWhiteSpace($name)) { return $null }

    $windows11Builds = @{
        '26H1' = 28000
        '25H2' = 26200
        '24H2' = 26100
        '23H2' = 22631
        '22H2' = 22621
        '21H2' = 22000
    }
    if ($name -match '(?i)\bWindows\s*11\s*(?<Version>26H1|25H2|24H2|23H2|22H2|21H2)\b') {
        $version = $matches.Version.ToUpperInvariant()
        return [pscustomobject][ordered]@{
            Label = "Windows 11 $version"
            ShortLabel = $version
            Build = [int]$windows11Builds[$version]
        }
    }

    $windows10Builds = @{
        '22H2' = 19045
        '21H2' = 19044
    }
    if ($name -match '(?i)\bWindows\s*10\s*(?<Version>22H2|21H2)\b') {
        $version = $matches.Version.ToUpperInvariant()
        return [pscustomobject][ordered]@{
            Label = "Windows 10 $version"
            ShortLabel = $version
            Build = [int]$windows10Builds[$version]
        }
    }

    return $null
}

function Get-WinUpdateOsCoverageSummary {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory)][int]$MinimumBuild
    )

    $buildByDevice = @{}
    $rowIndex = 0
    foreach ($row in $Rows) {
        $rowIndex++
        $deviceId = [string](Get-WinUpdateRowValue -Row $row -Name 'DeviceId')
        $deviceName = [string](Get-WinUpdateRowValue -Row $row -Name 'DeviceName')
        $normalizedName = Normalize-DeviceName -DeviceName $deviceName
        $key = if (-not [string]::IsNullOrWhiteSpace($deviceId)) {
            "id:$($deviceId.ToLowerInvariant())"
        } elseif ($normalizedName) {
            "name:$normalizedName"
        } else {
            "row:$rowIndex"
        }

        $build = Get-WinUpdateOsBuild -Value (Get-WinUpdateRowValue -Row $row -Name 'OSVersion')
        if (-not $buildByDevice.ContainsKey($key)) {
            $buildByDevice[$key] = $build
        } elseif ($null -ne $build -and ($null -eq $buildByDevice[$key] -or $build -gt $buildByDevice[$key])) {
            $buildByDevice[$key] = $build
        }
    }

    $total = $buildByDevice.Count
    $known = @($buildByDevice.Values | Where-Object { $null -ne $_ })
    $covered = @($known | Where-Object { $_ -ge $MinimumBuild }).Count
    $knownCount = $known.Count
    $unknown = $total - $knownCount
    $below = $knownCount - $covered
    $windows11 = @($known | Where-Object { $_ -ge 22000 }).Count
    $windows10 = @($known | Where-Object { $_ -ge 10240 -and $_ -lt 22000 }).Count
    $unknownOrOther = $total - $windows11 - $windows10

    [pscustomobject][ordered]@{
        TotalDevices = $total
        KnownOsVersion = $knownCount
        Covered = $covered
        BelowTarget = $below
        UnknownOsVersion = $unknown
        Windows11 = $windows11
        Windows10 = $windows10
        UnknownOrOther = $unknownOrOther
        Windows11Pct = if ($total -gt 0) { [math]::Round(100.0 * $windows11 / $total,2) } else { 0.0 }
        Windows10Pct = if ($total -gt 0) { [math]::Round(100.0 * $windows10 / $total,2) } else { 0.0 }
        UnknownOrOtherPct = if ($total -gt 0) { [math]::Round(100.0 * $unknownOrOther / $total,2) } else { 0.0 }
        CoveragePct = if ($total -gt 0) { [math]::Round(100.0 * $covered / $total,2) } else { 0.0 }
        KnownCoveragePct = if ($knownCount -gt 0) { [math]::Round(100.0 * $covered / $knownCount,2) } else { 0.0 }
    }
}

function Get-WinUpdateFreshCsvItem {
    param([Parameter(Mandatory)][string]$Path, [Parameter(Mandatory)][int]$MaxAgeHours)

    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Source CSV is missing: $Path" }
    $item = Get-Item -LiteralPath $Path -ErrorAction Stop
    $ageHours = ((Get-Date).ToUniversalTime() - $item.LastWriteTimeUtc).TotalHours
    if ($ageHours -gt $MaxAgeHours -or $ageHours -lt -1) {
        throw "Source CSV is outside the $MaxAgeHours-hour freshness window: $Path"
    }
    $item
}

function New-WinUpdateActivationIndex {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$EntraDeviceRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$AdComputerRows,
        [Parameter(Mandatory)][string]$TenantKey
    )

    $entraById = @{}
    foreach ($device in $EntraDeviceRows) {
        if ([string]$device.TenantKey -ne $TenantKey) { continue }
        $id = ([string]$device.DeviceId).Trim().Trim('{}').ToLowerInvariant()
        if (-not $id) { continue }
        if ($entraById.ContainsKey($id)) { throw 'Entra device inventory contains a duplicate DeviceId.' }
        $entraById[$id] = ([string]$device.AccountEnabled).Trim()
    }
    if ($entraById.Count -eq 0) { throw 'No Entra DeviceIds match the current tenant.' }

    $adById = @{}
    foreach ($computer in $AdComputerRows) {
        if ([string]$computer.TenantKey -ne $TenantKey) { continue }
        $id = ([string]$computer.ObjectGUID).Trim().Trim('{}').ToLowerInvariant()
        if (-not $id) { continue }
        if ($adById.ContainsKey($id)) { throw 'AD computer inventory contains a duplicate ObjectGUID.' }
        $adById[$id] = ([string]$computer.Enabled).Trim()
    }
    if ($adById.Count -eq 0) { throw 'No AD ObjectGUIDs match the current tenant.' }
    [pscustomobject]@{ EntraById = $entraById; AdById = $adById }
}

function Get-WinUpdateActivationState {
    param([AllowEmptyString()][string]$DeviceGuid, [Parameter(Mandatory)][object]$Index)

    $id = ([string]$DeviceGuid).Trim().Trim('{}').ToLowerInvariant()
    if (-not $id) { return 'Unqualified' }
    $entra = $Index.EntraById[$id]
    $ad = $Index.AdById[$id]
    if ($entra -eq 'True' -or $ad -eq 'True') { return 'Enabled' }
    if ($Index.EntraById.ContainsKey($id) -and $Index.AdById.ContainsKey($id) -and $entra -eq 'False' -and $ad -eq 'False') {
        return 'DisabledBoth'
    }
    return 'Unqualified'
}

function Get-WinUpdateIntuneFleetCountrySummary {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$IntuneDeviceRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$ActiveUserRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$PolicyRows,
        [Parameter(Mandatory)][object]$ActivationIndex,
        [Parameter(Mandatory)][AllowEmptyString()][string]$ReferencePolicyId,
        [Parameter(Mandatory)][string]$TenantKey
    )

    $countryByUserId = @{}
    $ambiguousUserIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($user in $ActiveUserRows) {
        if ([string]$user.TenantKey -ne $TenantKey) { continue }
        $id = ([string]$user.'Object Id').Trim().ToLowerInvariant()
        if (-not $id) { continue }
        if ($countryByUserId.ContainsKey($id)) { [void]$ambiguousUserIds.Add($id); continue }
        $countryByUserId[$id] = ([string]$user.CountryOrRegion).Trim()
    }
    if ($countryByUserId.Count -eq 0) { throw 'No active-user rows match the current tenant.' }

    $referenceIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $otherPolicyIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $PolicyRows) {
        # Graph report rows are tenant-scoped in memory; the CSV publisher adds TenantKey later.
        if ($row.PSObject.Properties['TenantKey'] -and [string]$row.TenantKey -ne $TenantKey) { continue }
        $id = ([string]$row.DeviceId).Trim().ToLowerInvariant()
        if (-not $id) { continue }
        if ($ReferencePolicyId -and [string]$row.PolicyId -eq $ReferencePolicyId) { [void]$referenceIds.Add($id) }
        else { [void]$otherPolicyIds.Add($id) }
    }

    $groups = @{}
    $seenDeviceIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $nonWindowsRows = 0
    $excludedDisabledBoth = 0
    $excludedUnqualified = 0
    $excludedReferencePolicy = 0
    foreach ($device in $IntuneDeviceRows) {
        if ([string]$device.TenantKey -ne $TenantKey) { continue }
        if ([string]$device.OS -ne 'Windows') { $nonWindowsRows++; continue }
        $id = ([string]$device.'Device ID').Trim().ToLowerInvariant()
        if (-not $id -or -not $seenDeviceIds.Add($id)) { throw 'Intune Windows inventory contains a blank or duplicate managed Device ID.' }
        $activationState = Get-WinUpdateActivationState -DeviceGuid ([string]$device.'Azure AD Device ID') -Index $ActivationIndex
        if ($activationState -ne 'Enabled') {
            if ($activationState -eq 'DisabledBoth') { $excludedDisabledBoth++ } else { $excludedUnqualified++ }
            if ($referenceIds.Contains($id)) { $excludedReferencePolicy++ }
            continue
        }
        $country = ''
        $userId = ([string]$device.UserId).Trim().ToLowerInvariant()
        if ($userId -and $countryByUserId.ContainsKey($userId) -and -not $ambiguousUserIds.Contains($userId)) {
            $country = [string]$countryByUserId[$userId]
        }
        $key = $country.ToLowerInvariant()
        if (-not $groups.ContainsKey($key)) {
            $groups[$key] = [pscustomobject]@{ Country = $country; Windows11 = 0; Windows10 = 0; UnknownOrOther = 0; Total = 0; ReferencePolicy = 0; OtherPolicyOnly = 0; NeitherPolicy = 0 }
        }
        $group = $groups[$key]
        $build = Get-WinUpdateOsBuild -Value $device.'OS version'
        if ($null -ne $build -and $build -ge 22000) { $group.Windows11++ }
        elseif ($null -ne $build -and $build -ge 10240) { $group.Windows10++ }
        else { $group.UnknownOrOther++ }
        if ($referenceIds.Contains($id)) { $group.ReferencePolicy++ }
        elseif ($otherPolicyIds.Contains($id)) { $group.OtherPolicyOnly++ }
        else { $group.NeitherPolicy++ }
        $group.Total++
    }
    if ($seenDeviceIds.Count -eq 0) { throw 'No Intune Windows devices match the current tenant.' }

    $known = @($groups.Values | Where-Object { $_.Country } | Sort-Object @{Expression='Total';Descending=$true},Country)
    $unknown = if ($groups.ContainsKey('')) { $groups[''] } else { [pscustomobject]@{ Country = ''; Windows11 = 0; Windows10 = 0; UnknownOrOther = 0; Total = 0; ReferencePolicy = 0; OtherPolicyOnly = 0; NeitherPolicy = 0 } }
    $total = [pscustomobject]@{
        Country = 'Total'
        Windows11 = [int](($groups.Values | Measure-Object -Property Windows11 -Sum).Sum)
        Windows10 = [int](($groups.Values | Measure-Object -Property Windows10 -Sum).Sum)
        UnknownOrOther = [int](($groups.Values | Measure-Object -Property UnknownOrOther -Sum).Sum)
        Total = $seenDeviceIds.Count - $excludedDisabledBoth - $excludedUnqualified
        ReferencePolicy = [int](($groups.Values | Measure-Object -Property ReferencePolicy -Sum).Sum)
        OtherPolicyOnly = [int](($groups.Values | Measure-Object -Property OtherPolicyOnly -Sum).Sum)
        NeitherPolicy = [int](($groups.Values | Measure-Object -Property NeitherPolicy -Sum).Sum)
    }
    $reportOnlyReferenceIds = 0
    foreach ($id in $referenceIds) { if (-not $seenDeviceIds.Contains($id)) { $reportOnlyReferenceIds++ } }
    [pscustomobject]@{
        Known = $known; Unknown = $unknown; Total = $total
        AllIntuneWindows = $seenDeviceIds.Count; ExcludedDisabledBoth = $excludedDisabledBoth
        ExcludedUnqualified = $excludedUnqualified; ExcludedReferencePolicy = $excludedReferencePolicy
        ReportOnlyReferenceIds = $reportOnlyReferenceIds; NonWindowsRows = $nonWindowsRows
    }
}

function Get-WinUpdateAdWithoutIntuneSummary {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$AdComputerRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$IntuneDeviceRows,
        [Parameter(Mandatory)][object]$ActivationIndex,
        [Parameter(Mandatory)][string]$TenantKey,
        [Parameter(Mandatory)][AllowEmptyCollection()][System.Collections.IDictionary]$CountryPrefixMap
    )

    $intuneAadCounts = @{}
    foreach ($device in $IntuneDeviceRows) {
        if ([string]$device.TenantKey -ne $TenantKey) { continue }
        $id = ([string]$device.'Azure AD Device ID').Trim().Trim('{}').ToLowerInvariant()
        if ($id) {
            if (-not $intuneAadCounts.ContainsKey($id)) { $intuneAadCounts[$id] = 0 }
            $intuneAadCounts[$id]++
        }
    }
    if ($intuneAadCounts.Count -eq 0) { throw 'No Intune Azure AD Device IDs match the current tenant.' }

    $seenAdIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $all = [pscustomobject]@{ Windows11 = 0; Windows10 = 0; Total = 0 }
    $active = [pscustomobject]@{ Windows11 = 0; Windows10 = 0; Total = 0 }
    $countryGroups = @{}
    $ambiguous = 0
    $unmatchedTotal = 0
    $excludedDisabledBoth = 0
    $excludedUnqualified = 0
    foreach ($ad in $AdComputerRows) {
        if ([string]$ad.TenantKey -ne $TenantKey) { continue }
        $os = [string]$ad.OperatingSystemShortName
        if ($os -notin @('Windows 10','Windows 11')) { continue }
        $id = ([string]$ad.ObjectGUID).Trim().Trim('{}').ToLowerInvariant()
        if (-not $id -or -not $seenAdIds.Add($id)) { throw 'AD Windows computer inventory contains a blank or duplicate ObjectGUID.' }
        if ($intuneAadCounts.ContainsKey($id)) {
            if ($intuneAadCounts[$id] -gt 1) { $ambiguous++ }
            continue
        }
        $unmatchedTotal++
        $activationState = Get-WinUpdateActivationState -DeviceGuid $id -Index $ActivationIndex
        if ($activationState -ne 'Enabled') {
            if ($activationState -eq 'DisabledBoth') { $excludedDisabledBoth++ } else { $excludedUnqualified++ }
            continue
        }
        if ($os -eq 'Windows 11') { $all.Windows11++ } else { $all.Windows10++ }
        $all.Total++
        $computerName = ([string]$ad.Name).Trim()
        $prefix = if ($computerName.Length -ge 2) { $computerName.Substring(0, 2).ToUpperInvariant() } else { '' }
        $country = if ($CountryPrefixMap.Contains($prefix)) { [string]$CountryPrefixMap[$prefix] } else { '' }
        $countryKey = $country.ToLowerInvariant()
        if (-not $countryGroups.ContainsKey($countryKey)) {
            $countryGroups[$countryKey] = [pscustomobject]@{ Country = $country; Windows11 = 0; Windows10 = 0; Total = 0 }
        }
        $countryGroup = $countryGroups[$countryKey]
        if ($os -eq 'Windows 11') { $countryGroup.Windows11++ } else { $countryGroup.Windows10++ }
        $countryGroup.Total++
        if ([string]$ad.IsActiveInLast45Days -eq 'True') {
            if ($os -eq 'Windows 11') { $active.Windows11++ } else { $active.Windows10++ }
            $active.Total++
        }
    }
    if ($seenAdIds.Count -eq 0) { throw 'No AD Windows 10 or Windows 11 computers match the current tenant.' }
    $knownCountries = @($countryGroups.Values | Where-Object { $_.Country } | Sort-Object @{Expression='Total';Descending=$true},Country)
    $unknownCountry = if ($countryGroups.ContainsKey('')) { $countryGroups[''] } else { [pscustomobject]@{ Country = ''; Windows11 = 0; Windows10 = 0; Total = 0 } }
    if ([int](($countryGroups.Values | Measure-Object -Property Windows11 -Sum).Sum) -ne $all.Windows11 -or
        [int](($countryGroups.Values | Measure-Object -Property Windows10 -Sum).Sum) -ne $all.Windows10) {
        throw 'Unmatched AD country OS totals do not reconcile.'
    }
    [pscustomobject]@{
        All = $all; EnabledActive45 = $active; Known = $knownCountries; Unknown = $unknownCountry; AmbiguousIntuneIds = $ambiguous
        AdWindowsClients = $seenAdIds.Count; UnmatchedTotal = $unmatchedTotal
        ExcludedDisabledBoth = $excludedDisabledBoth; ExcludedUnqualified = $excludedUnqualified
    }
}

function Get-WinUpdateEnabledWindowsCombinedSummary {
    param(
        [Parameter(Mandatory)][object]$IntuneSummary,
        [Parameter(Mandatory)][object]$AdWithoutIntuneSummary
    )

    $intune = $IntuneSummary.Total
    $ad = $AdWithoutIntuneSummary.All
    if (($intune.Windows11 + $intune.Windows10 + $intune.UnknownOrOther) -ne $intune.Total -or
        ($ad.Windows11 + $ad.Windows10) -ne $ad.Total) {
        throw 'Enabled Intune and unmatched AD OS totals do not reconcile.'
    }
    $countryGroups = @{}
    foreach ($source in @('Intune','AD')) {
        $sourceRows = if ($source -eq 'Intune') { @($IntuneSummary.Known) + @($IntuneSummary.Unknown) } else { @($AdWithoutIntuneSummary.Known) + @($AdWithoutIntuneSummary.Unknown) }
        foreach ($sourceRow in $sourceRows) {
            $country = ([string]$sourceRow.Country).Trim()
            $countryKey = $country.ToLowerInvariant()
            if (-not $countryGroups.ContainsKey($countryKey)) {
                $countryGroups[$countryKey] = [pscustomobject]@{ Country = $country; Windows11 = 0; Windows10 = 0; UnknownOrOther = 0; Total = 0; Intune = 0; AdWithoutIntune = 0; ReferencePolicy = 0; OtherPolicyOnly = 0; NeitherPolicy = 0 }
            }
            $row = $countryGroups[$countryKey]
            $row.Windows11 += [int]$sourceRow.Windows11
            $row.Windows10 += [int]$sourceRow.Windows10
            $row.Total += [int]$sourceRow.Total
            if ($source -eq 'Intune') {
                $row.UnknownOrOther += [int]$sourceRow.UnknownOrOther
                $row.Intune += [int]$sourceRow.Total
                $row.ReferencePolicy += [int]$sourceRow.ReferencePolicy
                $row.OtherPolicyOnly += [int]$sourceRow.OtherPolicyOnly
                $row.NeitherPolicy += [int]$sourceRow.NeitherPolicy
            } else {
                $row.AdWithoutIntune += [int]$sourceRow.Total
            }
        }
    }
    $knownCountries = @($countryGroups.Values | Where-Object { $_.Country } | Sort-Object @{Expression='Total';Descending=$true},Country)
    $unknownCountry = if ($countryGroups.ContainsKey('')) { $countryGroups[''] } else { [pscustomobject]@{ Country = ''; Windows11 = 0; Windows10 = 0; UnknownOrOther = 0; Total = 0; Intune = 0; AdWithoutIntune = 0; ReferencePolicy = 0; OtherPolicyOnly = 0; NeitherPolicy = 0 } }
    $countryTotal = [int](($countryGroups.Values | Measure-Object -Property Total -Sum).Sum)
    if ($countryTotal -ne ($intune.Total + $ad.Total)) { throw 'Combined country rows do not reconcile with source totals.' }
    foreach ($row in $countryGroups.Values) {
        if (($row.Windows11 + $row.Windows10 + $row.UnknownOrOther) -ne $row.Total -or
            ($row.Intune + $row.AdWithoutIntune) -ne $row.Total -or
            ($row.ReferencePolicy + $row.OtherPolicyOnly + $row.NeitherPolicy) -ne $row.Intune) {
            throw 'A combined country row does not reconcile by OS or source.'
        }
    }
    [pscustomobject]@{
        Windows11 = $intune.Windows11 + $ad.Windows11
        Windows10 = $intune.Windows10 + $ad.Windows10
        UnknownOrOther = $intune.UnknownOrOther
        Total = $intune.Total + $ad.Total
        Known = $knownCountries; Unknown = $unknownCountry
    }
}



function Get-WinUpdateProgressPhase {
    param([Parameter(Mandatory)][object]$Row)

    if ((Get-WinUpdateOperationalState -Row $Row) -ne 'InProgress') { return '' }

    $status = [string](Get-WinUpdateRowValue -Row $Row -Name 'CurrentDeviceUpdateStatus_loc')
    if ([string]::IsNullOrWhiteSpace($status)) {
        $status = [string](Get-WinUpdateRowValue -Row $Row -Name 'AggregateState')
    }
    $normalized = $status.ToLowerInvariant()

    if ($normalized -match 'offer') { return 'Offering' }
    if ($normalized -match 'installing|download|restart') { return 'Installing' }
    if ($normalized -match 'pending|hold|wait|defer') { return 'Pending' }
    return 'Other'
}

function Get-WinUpdatePolicySummary {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows)

    $result = foreach ($group in ($Rows | Group-Object -Property PolicyId)) {
        $policyRows = @($group.Group)
        if ($policyRows.Count -eq 0) { continue }

        $policyCompleted = 0
        $policyInProgress = 0
        $policyAction = 0
        $policyUnknown = 0
        $policyOffering = 0
        $policyInstalling = 0
        $policyPending = 0
        $policyOtherInProgress = 0

        foreach ($policyRow in $policyRows) {
            $policyOperationalState = Get-WinUpdateOperationalState -Row $policyRow
            switch ($policyOperationalState) {
                'Completed'      { $policyCompleted++; break }
                'InProgress'     {
                    $policyInProgress++
                    switch (Get-WinUpdateProgressPhase -Row $policyRow) {
                        'Offering'   { $policyOffering++; break }
                        'Installing' { $policyInstalling++; break }
                        'Pending'    { $policyPending++; break }
                        default      { $policyOtherInProgress++ }
                    }
                    break
                }
                'ActionRequired' { $policyAction++; break }
                default          { $policyUnknown++ }
            }
        }

        $policyTotal = $policyRows.Count
        $policyName = [string]$policyRows[0].PolicyName
        $target = Get-WinUpdatePolicyTarget -PolicyName $policyName
        $coverage = if ($null -ne $target) {
            Get-WinUpdateOsCoverageSummary -Rows $policyRows -MinimumBuild $target.Build
        } else {
            $null
        }
        $policyStatus = if ($policyAction -gt 0) { 'CRITICAL' } elseif ($policyInProgress -gt 0 -or $policyUnknown -gt 0) { 'WARNING' } else { 'OK' }

        [pscustomobject][ordered]@{
            PolicyId = [string]$group.Name
            PolicyName = $policyName
            Status = $policyStatus
            Devices = $policyTotal
            Completed = $policyCompleted
            InProgress = $policyInProgress
            Offering = $policyOffering
            Installing = $policyInstalling
            Pending = $policyPending
            OtherInProgress = $policyOtherInProgress
            ActionRequired = $policyAction
            Unknown = $policyUnknown
            PolicyCompletionPct = if ($policyTotal -gt 0) { [math]::Round(100.0 * $policyCompleted / $policyTotal,2) } else { 0.0 }
            TargetLabel = if ($null -ne $target) { [string]$target.Label } else { '' }
            TargetShortLabel = if ($null -ne $target) { [string]$target.ShortLabel } else { '' }
            TargetBuild = if ($null -ne $target) { [int]$target.Build } else { $null }
            OsCovered = if ($null -ne $coverage) { [int]$coverage.Covered } else { $null }
            OsBelowTarget = if ($null -ne $coverage) { [int]$coverage.BelowTarget } else { $null }
            OsVersionUnknown = if ($null -ne $coverage) { [int]$coverage.UnknownOsVersion } else { $null }
            Windows11 = if ($null -ne $coverage) { [int]$coverage.Windows11 } else { $null }
            Windows10 = if ($null -ne $coverage) { [int]$coverage.Windows10 } else { $null }
            UnknownOrOther = if ($null -ne $coverage) { [int]$coverage.UnknownOrOther } else { $null }
            Windows11Pct = if ($null -ne $coverage) { [double]$coverage.Windows11Pct } else { $null }
            Windows10Pct = if ($null -ne $coverage) { [double]$coverage.Windows10Pct } else { $null }
            UnknownOrOtherPct = if ($null -ne $coverage) { [double]$coverage.UnknownOrOtherPct } else { $null }
            OsCoveragePct = if ($null -ne $coverage) { [double]$coverage.CoveragePct } else { $null }
            KnownOsCoveragePct = if ($null -ne $coverage) { [double]$coverage.KnownCoveragePct } else { $null }
        }
    }

    @($result)
}

function Select-WinUpdatePrimaryCoveragePolicy {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$PolicySummary)

    @($PolicySummary | Where-Object { $null -ne $_.TargetBuild } |
        Sort-Object @{Expression='Devices';Descending=$true},@{Expression='TargetBuild';Descending=$true},PolicyName |
        Select-Object -First 1)[0]
}

function Get-WinUpdateOperationalRank {
    param([Parameter(Mandatory)][string]$State)
    switch ($State) {
        'ActionRequired' { return 0 }
        'InProgress'     { return 1 }
        'Unknown'        { return 2 }
        'Completed'      { return 3 }
        default          { return 9 }
    }
}

function Get-WinUpdatePriorityRank {
    param([AllowNull()][object]$Value)
    $priority = 999
    if ($null -ne $Value) { [void][int]::TryParse([string]$Value,[ref]$priority) }
    return $priority
}

function Get-WinUpdateReportHealth {
    param(
        [Parameter(Mandatory)][ValidateRange(0,[int]::MaxValue)][int]$TotalUniqueDevices,
        [Parameter(Mandatory)][ValidateRange(0,[int]::MaxValue)][int]$ActionRequiredCount,
        [Parameter(Mandatory)][ValidateRange(0,[int]::MaxValue)][int]$Priority0Count,
        [Parameter(Mandatory)][bool]$CoverageAvailable,
        [Parameter(Mandatory)][ValidateRange(0.0,100.0)][double]$KnownOsCoveragePct,
        [Parameter(Mandatory)][ValidateRange(0,[int]::MaxValue)][int]$FleetDevices,
        [Parameter(Mandatory)][ValidateRange(0,[int]::MaxValue)][int]$OsVersionUnknownCount
    )

    $actionRequiredRatePct = if ($TotalUniqueDevices -gt 0) {
        [math]::Round(100.0 * $ActionRequiredCount / $TotalUniqueDevices,2)
    } else { 0.0 }
    $osVersionUnknownRatePct = if ($FleetDevices -gt 0) {
        [math]::Round(100.0 * $OsVersionUnknownCount / $FleetDevices,2)
    } else { 0.0 }
    $priority0Threshold = [int][math]::Max(10,[math]::Ceiling(0.01 * $TotalUniqueDevices))

    $critical = (
        $actionRequiredRatePct -ge 1.0 -or
        $Priority0Count -ge $priority0Threshold -or
        ($CoverageAvailable -and $KnownOsCoveragePct -lt 90.0) -or
        ($CoverageAvailable -and $osVersionUnknownRatePct -ge 15.0)
    )
    $warning = (
        $ActionRequiredCount -gt 0 -or
        -not $CoverageAvailable -or
        $KnownOsCoveragePct -lt 95.0 -or
        $osVersionUnknownRatePct -ge 5.0
    )
    $status = if ($critical) { 'CRITICAL' } elseif ($warning) { 'WARNING' } else { 'OK' }

    [pscustomobject][ordered]@{
        Status = $status
        ActionRequiredRatePct = $actionRequiredRatePct
        Priority0Count = $Priority0Count
        Priority0Threshold = $priority0Threshold
        KnownOsCoveragePct = $KnownOsCoveragePct
        OsVersionUnknownRatePct = $osVersionUnknownRatePct
    }
}

function Get-WinUpdateDeviceSummaryRows {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows)

    $selectedByKey = @{}
    $keyOrder = New-Object System.Collections.Generic.List[string]
    $rowIndex = 0
    foreach ($row in $Rows) {
        $rowIndex++
        $deviceId = [string]$row.DeviceId
        $deviceName = [string]$row.DeviceName
        $normalizedName = Normalize-DeviceName -DeviceName $deviceName
        $key = if (-not [string]::IsNullOrWhiteSpace($deviceId)) { "id:$($deviceId.ToLowerInvariant())" } elseif ($normalizedName) { "name:$normalizedName" } else { "row:$rowIndex" }

        $operationalState = Get-WinUpdateOperationalState -Row $row
        $operationalRank = Get-WinUpdateOperationalRank -State $operationalState
        $priorityRank = 999
        if ($null -ne $row.ActionPriority) { [void][int]::TryParse([string]$row.ActionPriority,[ref]$priorityRank) }

        if (-not $selectedByKey.ContainsKey($key)) {
            $selectedByKey[$key] = [pscustomobject]@{
                Row = $row
                OperationalState = $operationalState
                OperationalRank = $operationalRank
                PriorityRank = $priorityRank
            }
            $keyOrder.Add($key) | Out-Null
            continue
        }

        $current = $selectedByKey[$key]
        if ($operationalRank -lt $current.OperationalRank -or ($operationalRank -eq $current.OperationalRank -and $priorityRank -lt $current.PriorityRank)) {
            $selectedByKey[$key] = [pscustomobject]@{
                Row = $row
                OperationalState = $operationalState
                OperationalRank = $operationalRank
                PriorityRank = $priorityRank
            }
        }
    }

    foreach ($key in $keyOrder) {
        $entry = $selectedByKey[$key]
        $selected = $entry.Row
        [pscustomobject][ordered]@{
            DeviceKey = $key
            DeviceId = [string]$selected.DeviceId
            DeviceName = [string]$selected.DeviceName
            PolicyId = [string]$selected.PolicyId
            PolicyName = [string]$selected.PolicyName
            OperationalState = $entry.OperationalState
            AggregateState = [string]$selected.AggregateState
            CurrentDeviceUpdateStatus = [string]$selected.CurrentDeviceUpdateStatus
            CurrentDeviceUpdateStatus_loc = [string]$selected.CurrentDeviceUpdateStatus_loc
            ProgressPhase = Get-WinUpdateProgressPhase -Row $selected
            OSVersion = [string]$selected.OSVersion
            OSVersionSource = [string]$selected.OSVersionSource
            LatestAlertMessage = [string]$selected.LatestAlertMessage
            BlockingReason = [string]$selected.BlockingReason
            UpgradeEligibilityLabel = [string]$selected.UpgradeEligibilityLabel
            ReadinessMatch = [string]$selected.ReadinessMatch
            ActionPriority = $entry.PriorityRank
            ActionCode = [string]$selected.ActionCode
            ActionDescription = [string]$selected.ActionDescription
            ActionOwner = [string]$selected.ActionOwner
        }
    }
}
# ==========================================================
function Select-WinUpdateCanonicalRows {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows)

    $selectedByKey = @{}
    $keyOrder = New-Object System.Collections.Generic.List[string]
    $rowIndex = 0
    foreach ($row in $Rows) {
        $rowIndex++
        $policyId = [string]$row.PolicyId
        $deviceId = [string]$row.DeviceId

        if (-not [string]::IsNullOrWhiteSpace($policyId)) {
            $policyKey = "id:$($policyId.ToLowerInvariant())"
        } else {
            $policyName = [string]$row.PolicyName
            $policyKey = if (-not [string]::IsNullOrWhiteSpace($policyName)) { "name:$($policyName.Trim().ToLowerInvariant())" } else { 'none' }
        }

        if (-not [string]::IsNullOrWhiteSpace($deviceId)) {
            $deviceKey = "id:$($deviceId.ToLowerInvariant())"
        } else {
            $normalizedName = Normalize-DeviceName -DeviceName ([string]$row.DeviceName)
            $deviceKey = if ($normalizedName) { "name:$normalizedName" } else { "row:$rowIndex" }
        }

        $key = "policy:$policyKey|device:$deviceKey"
        $blockingReason = [string]$row.BlockingReason
        $operationalRank = if ($blockingReason -like 'HardFailure*' -or $blockingReason -eq 'UpdateAlert') {
            0
        } elseif ($blockingReason -eq 'DeploymentInProgress') {
            1
        } elseif ($blockingReason -eq 'NotApplicableOrUnknown') {
            2
        } elseif ($blockingReason -eq 'CompliantOrCompleted') {
            3
        } else {
            switch ([string]$row.RiskBucket) {
                'High'   { 0; break }
                'Medium' { 1; break }
                'Low'    { 3; break }
                default  { 2 }
            }
        }
        $priorityRank = 999
        if ($null -ne $row.ActionPriority) { [void][int]::TryParse([string]$row.ActionPriority, [ref]$priorityRank) }

        if (-not $selectedByKey.ContainsKey($key)) {
            $selectedByKey[$key] = [pscustomobject]@{
                Row = $row
                OperationalRank = $operationalRank
                PriorityRank = $priorityRank
            }
            $keyOrder.Add($key) | Out-Null
            continue
        }

        $current = $selectedByKey[$key]
        if ($operationalRank -lt $current.OperationalRank -or ($operationalRank -eq $current.OperationalRank -and $priorityRank -lt $current.PriorityRank)) {
            $selectedByKey[$key] = [pscustomobject]@{
                Row = $row
                OperationalRank = $operationalRank
                PriorityRank = $priorityRank
            }
        }
    }

    foreach ($key in $keyOrder) {
        $selectedByKey[$key].Row
    }
}# Main
# ==========================================================
$ErrorActionPreference  = "Stop"
$script:TokenAcquiredAt = [datetime]::MinValue
$scriptStart = Get-Date
$global:SmartM365WarningCount = 0
$global:SmartM365ErrorCount = 0

$script:CompletionStatus = 'Auto'
try {
    Assert-PS7
    Import-SmartM365CorePreflight
    . (Join-Path (Get-Module SmartM365.Core).ModuleBase 'SmartM365-SharePointUploadPolicy.ps1')
    Initialize-SmartM365CollectorUploadContext -LocalConfig $ScriptLocalConfig -GlobalConfig $global:SmartM365GlobalConfig -DataAllRoot $ScriptCsvLogFolderPath -DataLastRoot $LatestCsvFolderPath -LogAllRoot $LogsPath -ExternalActionsAllowed (-not $DryRun)
    $EnableSharePointUpload=[bool]$global:EnableSharePointUpload
    $dataAllUploadEnabled=Test-SmartM365CollectorUploadEnabled -Category DataAll -Policies $global:SmartM365CollectorUploadPolicies -RunEnabled $EnableSharePointUpload
    $dataLastUploadEnabled=Test-SmartM365CollectorUploadEnabled -Category DataLast -Policies $global:SmartM365CollectorUploadPolicies -RunEnabled $EnableSharePointUpload
    Start-CoreSmartM365SourceReceipt -ScriptPath $PSCommandPath -SourceRootPath $LatestCsvFolderPath -ReadOnly:$DryRun -ScopeParameters @{MaxPolicies=$MaxPolicies;FeatureUpdatePolicyId=$FeatureUpdatePolicyId}

    Ensure-Directory -Path $ScriptCsvLogFolderPath
    Ensure-Directory -Path $ArchivePath
    Ensure-Directory -Path $LogsPath
    Ensure-Directory -Path $WorkPath
    Ensure-Directory -Path $LatestCsvFolderPath

    Prune-Files -Folder $LogsPath -Filter "$ScriptName-*.log" -Keep $global:RetentionMaxLogs

    # ----------------------------------------------------------
    # Load SmartM365.Core for optional SharePoint upload
    # ----------------------------------------------------------
    if ($EnableSharePointUpload) {
        Import-SmartM365CorePreflight
        Write-Log "SmartM365.Core module loaded for SharePoint upload." "INFO" "SP"
    }

    if ($ConsolePretty) {
        Write-Host ("{0}  v{1}  {2}" -f $ScriptName, $ScriptVersion, (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) -ForegroundColor White
        Write-Host ("Host: {0}   RunId: {1}" -f $env:COMPUTERNAME, $RunId) -ForegroundColor DarkGray
        if ($DryRun) { Write-Host "*** DRY-RUN MODE: no files written, no emails sent ***" -ForegroundColor Cyan }
        Write-Host ("=" * 80) -ForegroundColor DarkGray
    }

    Write-Log "Starting export. Version=$ScriptVersion DryRun=$DryRun MaxPolicies=$MaxPolicies EnableReadinessEnrichment=$EnableReadinessEnrichment" "INFO"

    $cert = Get-AppCertificate -Thumbprint $Thumb
    Write-Log "Certificate loaded: Subject=$($cert.Subject)" "INFO" "AUTH"

    $token = Get-GraphTokenWithCert -TenantId $TenantId -ClientId $AppId -Certificate $cert
    $script:TokenAcquiredAt = Get-Date
    Write-Log "Access token acquired at $($script:TokenAcquiredAt.ToString('HH:mm:ss'))." "INFO" "AUTH"

    $headers = @{ Authorization = "Bearer $token" }
    Import-SmartM365CorePreflight
    Invoke-CoreSmartM365Preflight -ScriptName $ScriptName -RequiredModules @('MSAL.PS') -OutputPaths @($ScriptCsvLogFolderPath, $LatestCsvFolderPath, $ArchivePath, $LogsPath, $WorkPath) -GraphAccessToken $token -RequiredGraphApplicationPermissions @('DeviceManagementConfiguration.Read.All','DeviceManagementManagedDevices.Read.All') -GraphProbeUris @(
        'https://graph.microsoft.com/beta/deviceManagement/windowsFeatureUpdateProfiles?$top=1'
    ) | Out-Null


    # ----------------------------------------------------------
    # Discover Feature Update policies
    # ----------------------------------------------------------
    $policies = Get-FeatureUpdatePolicies -Headers $headers
    if (-not $policies -or $policies.Count -eq 0) {
        throw "No Feature Update policies found (windowsFeatureUpdateProfiles)."
    }
    Write-Log "Feature Update policies discovered: $($policies.Count)" "INFO" "GRAPH"
    if (-not [string]::IsNullOrWhiteSpace($FeatureUpdatePolicyId)) {
        $requestedPolicyId = $FeatureUpdatePolicyId.Trim()
        $matchingPolicies = @($policies | Where-Object { $_.PolicyId -eq $requestedPolicyId })
        if ($matchingPolicies.Count -eq 0) {
            throw "Configured FeatureUpdatePolicyId was not found in windowsFeatureUpdateProfiles: $requestedPolicyId"
        }
        Write-Log "FeatureUpdatePolicyId configured: processing only policy $requestedPolicyId." "INFO" "GRAPH"
        $policies = $matchingPolicies
    }

    if ($MaxPolicies -gt 0 -and $policies.Count -gt $MaxPolicies) {
        Write-Log "MaxPolicies=${MaxPolicies}: limiting to first $MaxPolicies policies out of $($policies.Count)." "WARN" "GRAPH"
        $policies = $policies | Select-Object -First $MaxPolicies
    }

    $rawRows    = New-Object System.Collections.Generic.List[object]
    $exportTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")

    # ----------------------------------------------------------
    # Per-policy export loop
    # ----------------------------------------------------------
    $index = 0
    foreach ($p in $policies) {
        $index++
        $policyId   = $p.PolicyId
        $policyName = $p.PolicyName

        Write-Log "Processing policy [$index/$($policies.Count)]: $policyName ($policyId)" "INFO" "GRAPH"

        $filter  = "(PolicyId eq '$policyId')"
        $zipPath = Join-Path $WorkPath ("{0}_{1}_{2}.zip" -f $ReportName, $RunStamp, $policyId)

        # A single export job can stay unreachable (repeated 503 on its status
        # url) or produce a url that cannot be downloaded while the service is
        # healthy again for new jobs, so recreate the job once before failing.
        $maxExportJobAttempts = 2
        for ($exportJobAttempt = 1; $exportJobAttempt -le $maxExportJobAttempts; $exportJobAttempt++) {
            Refresh-GraphTokenIfNeeded -TenantId $TenantId -ClientId $AppId -Certificate $cert -Headers ([ref]$headers)
            $jobId = $null

            try {
                $jobId = Start-IntuneExportJobRaw -ReportName $ReportName -Headers $headers -Filter $filter -Select $ReportSelect -LocalizationType $LocalizationType
            }
            catch {
                $msg = Try-GetGraphErrorMessage -ErrorRecord $_
                if ($msg) { Write-Log "Graph POST error for policy ${policyId}: $msg" "WARN" "GRAPH" }
                if ($ReportSelect -and $ReportSelect.Count -gt 0) {
                    Write-Log "Retrying policy export without select." "WARN" "GRAPH"
                    $jobId = Start-IntuneExportJobRaw -ReportName $ReportName -Headers $headers -Filter $filter -Select $null -LocalizationType $LocalizationType
                }
                else { throw }
            }

            Write-Log "Export job created: $jobId" "INFO" "GRAPH"
            try {
                $job = Wait-IntuneExportJobComplete -JobId $jobId -Headers $headers -TimeoutSeconds 900 -PollSeconds 5 -PollMaxAttempts 5
                Download-ExportZip -DownloadUrl $job.url -DestinationZipPath $zipPath
                break
            }
            catch {
                if ($exportJobAttempt -ge $maxExportJobAttempts) { throw }
                Write-Log "Export job $jobId for policy $policyId could not be completed: $($_.Exception.Message) Creating a new export job (attempt $($exportJobAttempt + 1)/$maxExportJobAttempts)." "WARN" "GRAPH"
            }
        }
        Write-Log "ZIP downloaded: $zipPath" "INFO" "DATA"

        $extractFolder = Join-Path $WorkPath ("Extract_{0}_{1}" -f $RunStamp, $policyId)
        $csvExtracted  = Extract-FirstCsvFromZip -ZipPath $zipPath -ExtractFolder $extractFolder
        Write-Log "CSV extracted: $csvExtracted" "INFO" "DATA"

        $rows = Import-Csv -Path $csvExtracted
        foreach ($r in $rows) {
            $o = [ordered]@{}
            foreach ($prop in $r.PSObject.Properties) { $o[$prop.Name] = $prop.Value }
            $o["PolicyId"]       = $policyId
            $o["PolicyName"]     = $policyName
            $o["OrgDomain"]      = $OrgDomain
            $o["ExportDateTime"] = $exportTime
            $o["RunId"]          = $RunId
            $rawRows.Add([pscustomobject]$o) | Out-Null
        }
    }

    Write-Log "Total consolidated rows (raw): $($rawRows.Count)" "INFO" "DATA"
    Refresh-GraphTokenIfNeeded -TenantId $TenantId -ClientId $AppId -Certificate $cert -Headers ([ref]$headers)
    Resolve-WinUpdateMissingDeviceNames -Rows $rawRows.ToArray() -Headers $headers

    # ----------------------------------------------------------
    # Readiness enrichment
    # ----------------------------------------------------------
    $enrichedRows             = $rawRows
    $readinessTotal           = 0
    $readinessMatched         = 0
    $readinessNotMatched      = 0
    $readinessMatchPct        = 0.0
    $readinessEligibilityDist = ""
    $readinessMode            = $(if ($EnableReadinessEnrichment) { "Pending" } else { "Disabled" })
    $readinessSummaryText     = ""

    if ($EnableReadinessEnrichment) {
        $resolvedReadinessCsvPath = $null
        $resolvedReadinessSummaryCsvPath = $null
        if (-not [string]::IsNullOrWhiteSpace($ReadinessCsvPath)) {
            $resolvedReadinessCsvPath = Resolve-CoreSmartM365CsvPathWithSharePointFallback -Path $ReadinessCsvPath -Description 'Device-level Readiness CSV'
        }
        if (-not [string]::IsNullOrWhiteSpace($ReadinessSummaryCsvPath)) {
            $resolvedReadinessSummaryCsvPath = Resolve-CoreSmartM365CsvPathWithSharePointFallback -Path $ReadinessSummaryCsvPath -Description 'Readiness summary CSV'
        }

        if ([string]::IsNullOrWhiteSpace($resolvedReadinessCsvPath)) {
            if (-not [string]::IsNullOrWhiteSpace($resolvedReadinessSummaryCsvPath) -and (Test-Path $resolvedReadinessSummaryCsvPath)) {
                Write-Log "Device-level Readiness CSV not found; using summary context only: $resolvedReadinessSummaryCsvPath" "INFO" "READINESS"
                $readinessMode = "SummaryOnly"
                $summaryRow = @(Import-Csv -Path $resolvedReadinessSummaryCsvPath | Select-Object -First 1)
                if ($summaryRow.Count -gt 0) {
                    $s = $summaryRow[0]
                    $readinessSummaryText = "TotalDeviceCount={0}; UpgradeEligibleDeviceCount={1}; UpgradeEligiblePercentage={2}; ProcessorFamilyCheckFailedPercentage={3}; TPMCheckFailedPercentage={4}; SecureBootCheckFailedPercentage={5}" -f $s.TotalDeviceCount,$s.UpgradeEligibleDeviceCount,$s.UpgradeEligiblePercentage,$s.ProcessorFamilyCheckFailedPercentage,$s.TPMCheckFailedPercentage,$s.SecureBootCheckFailedPercentage
                    Write-Log "Readiness summary: $readinessSummaryText" "INFO" "READINESS"
                }

                $tmp = New-Object System.Collections.Generic.List[object]
                foreach ($row in $rawRows) {
                    $o = [ordered]@{}
                    foreach ($pp in $row.PSObject.Properties) { $o[$pp.Name] = $pp.Value }
                    $o["NormalizedDeviceName"]    = Normalize-DeviceName -DeviceName $row.DeviceName
                    $o["ReadinessMatch"]          = "SummaryOnly"
                    $o["UpgradeEligibility"]      = ""
                    $o["UpgradeEligibilityLabel"] = ""
                    $o["AzureAdJoinType"]         = ""
                    $o["ReadinessGraphId"]        = ""
                    $o["ReadinessExportDateTime"] = ""
                    $o["ReadinessRunId"]          = ""
                    $o["OSVersion"]               = ""
                    $tmp.Add([pscustomobject]$o) | Out-Null
                }
                $enrichedRows = $tmp
            }
            else {
                $readinessMode = "Missing"
                Write-Log "Readiness CSV not found locally or in SharePoint, enrichment skipped: $ReadinessCsvPath" "WARN" "READINESS"
            }
        }
        else {
            $readinessMode = "Joined"
            Write-Log "Loading Readiness CSV: $resolvedReadinessCsvPath" "INFO" "READINESS"
            $rd = Import-Csv -Path $resolvedReadinessCsvPath
            $readinessTotal = @($rd).Count
            Write-Log "Readiness rows loaded: $readinessTotal" "INFO" "READINESS"

            $lookup = Build-ReadinessLookup -ReadinessRows $rd
            Write-Log "Readiness lookup keys: $($lookup.Count)" "INFO" "READINESS"

            $tmp = New-Object System.Collections.Generic.List[object]
            foreach ($row in $rawRows) {
                $norm  = Normalize-DeviceName -DeviceName $row.DeviceName
                $rdRow = $null
                if ($norm -and $lookup.ContainsKey($norm)) { $rdRow = $lookup[$norm] }

                if ($rdRow) { $readinessMatched++ } else { $readinessNotMatched++ }

                $o = [ordered]@{}
                foreach ($pp in $row.PSObject.Properties) { $o[$pp.Name] = $pp.Value }

                $o["NormalizedDeviceName"]    = $norm
                $o["ReadinessMatch"]          = $(if ($rdRow) { "Matched"   } else { "NotMatched" })
                $o["UpgradeEligibility"]      = $(if ($rdRow) { $rdRow.UpgradeEligibility      } else { "" })
                $o["UpgradeEligibilityLabel"] = $(if ($rdRow) { $rdRow.UpgradeEligibilityLabel } else { "" })
                $o["AzureAdJoinType"]         = $(if ($rdRow) { $rdRow.AzureAdJoinType         } else { "" })
                $o["ReadinessGraphId"]        = $(if ($rdRow) { $rdRow.GraphId                 } else { "" })
                $o["ReadinessExportDateTime"] = $(if ($rdRow) { $rdRow.ExportDateTime          } else { "" })
                $o["ReadinessRunId"]          = $(if ($rdRow) { $rdRow.RunId                   } else { "" })
                $o["OSVersion"]               = $(if ($rdRow) { $rdRow.OSVersion               } else { "" })

                $tmp.Add([pscustomobject]$o) | Out-Null
            }

            $enrichedRows = $tmp
            $totalJoin    = $readinessMatched + $readinessNotMatched
            if ($totalJoin -gt 0) { $readinessMatchPct = [math]::Round((100.0 * $readinessMatched / $totalJoin), 2) }

            $eligGroups = $enrichedRows | Where-Object { $_.ReadinessMatch -eq "Matched" } |
                Group-Object UpgradeEligibilityLabel | Sort-Object Count -Descending
            $readinessEligibilityDist = ($eligGroups | ForEach-Object { "{0}={1}" -f $_.Name, $_.Count }) -join "; "

            Write-Log "Readiness join: Matched=$readinessMatched NotMatched=$readinessNotMatched MatchPct=$readinessMatchPct" "INFO" "READINESS"
        }
    }

    # ----------------------------------------------------------
    # OS version fallback: Windows Update -> Intune ID bridge -> Entra
    # ----------------------------------------------------------
    $intuneDeviceInventoryRows = @()
    $entraDeviceInventoryRows = @()
    $intuneDeviceInventoryAvailable = -not [string]::IsNullOrWhiteSpace($IntuneDeviceInventoryCsvPath) -and (Test-Path -LiteralPath $IntuneDeviceInventoryCsvPath)
    $entraDeviceInventoryAvailable = -not [string]::IsNullOrWhiteSpace($EntraDeviceInventoryCsvPath) -and (Test-Path -LiteralPath $EntraDeviceInventoryCsvPath)
    if ($intuneDeviceInventoryAvailable -and $entraDeviceInventoryAvailable) {
        $intuneDeviceInventoryRows = @(Import-Csv -LiteralPath $IntuneDeviceInventoryCsvPath)
        $entraDeviceInventoryRows = @(Import-Csv -LiteralPath $EntraDeviceInventoryCsvPath)
        Write-Log "Loading OS version fallback sources: IntuneDevices=$($intuneDeviceInventoryRows.Count) EntraDevices=$($entraDeviceInventoryRows.Count)." "INFO" "ENRICH"
    }
    else {
        Write-Log "OS version fallback sources unavailable: IntuneInventoryAvailable=$intuneDeviceInventoryAvailable EntraInventoryAvailable=$entraDeviceInventoryAvailable. Existing OSVersion values will be preserved." "WARN" "ENRICH"
    }
    $entraOsFallback = Add-WinUpdateEntraOsVersionFallback `
        -Rows $enrichedRows.ToArray() `
        -IntuneDeviceRows $intuneDeviceInventoryRows `
        -EntraDeviceRows $entraDeviceInventoryRows
    $enrichedRows = @($entraOsFallback.Rows)
    Write-Log "OS version fallback completed: ExistingRows=$($entraOsFallback.ExistingVersionRows) EntraRows=$($entraOsFallback.EntraVersionRows) UnavailableRows=$($entraOsFallback.UnavailableRows) IntuneBridgeKeys=$($entraOsFallback.IntuneBridgeKeys) AmbiguousIntuneBridgeKeys=$($entraOsFallback.AmbiguousIntuneBridgeKeys) EntraVersionKeys=$($entraOsFallback.EntraVersionKeys)." "INFO" "ENRICH"

    # ----------------------------------------------------------
    # Computed columns pass:
    #   RiskBucket, BlockingReason, DaysSinceLastStatus,
    #   ActionPriority, ActionCode, ActionDescription, ActionOwner
    # ----------------------------------------------------------
    $effectiveLastStatusDateColumn = $LastStatusDateColumn
    if ($enrichedRows.Count -gt 0 -and -not [string]::IsNullOrWhiteSpace($effectiveLastStatusDateColumn)) {
        $availableHeaders = @($enrichedRows[0].PSObject.Properties.Name)
        if ($availableHeaders -notcontains $effectiveLastStatusDateColumn) {
            Write-Log "Configured LastStatusDateColumn '$effectiveLastStatusDateColumn' is absent from the Graph report." "WARN" "ENRICH"
            $effectiveLastStatusDateColumn = @(
                'LastUpdateStatusTime',
                'LastUpdateStatusDateTime'
            ) | Where-Object { $availableHeaders -contains $_ } | Select-Object -First 1
            if ([string]::IsNullOrWhiteSpace($effectiveLastStatusDateColumn)) {
                Write-Log "No supported last-status date column is available; DaysSinceLastStatus will remain empty by design." "WARN" "ENRICH"
            }
            else {
                Write-Log "DaysSinceLastStatus will use detected column '$effectiveLastStatusDateColumn'." "INFO" "ENRICH"
            }
        }
    }
    Write-Log "Computing derived columns (RiskBucket, BlockingReason, DaysSinceLastStatus, Action*)..." "INFO" "ENRICH"

    $computedRows = New-Object System.Collections.Generic.List[object]
    $statusAgePopulatedRows = 0
    foreach ($row in $enrichedRows) {
        $o = [ordered]@{}
        foreach ($pp in $row.PSObject.Properties) { $o[$pp.Name] = $pp.Value }

        $riskBucket = Get-RiskBucket -AggregateState "$($row.AggregateState)" -IncludeUnknown $RiskIncludeUnknown
        $o["RiskBucket"] = $riskBucket

        $blockingReason = Get-BlockingReason `
            -AggregateState        "$($row.AggregateState)" `
            -LatestAlertMessageLoc "$($row.LatestAlertMessage_loc)"
        $o["BlockingReason"] = $blockingReason

        $daysSince = ""
        if (-not [string]::IsNullOrWhiteSpace($effectiveLastStatusDateColumn)) {
            $lastStatusRaw = "$($row.$effectiveLastStatusDateColumn)"
            if (-not [string]::IsNullOrWhiteSpace($lastStatusRaw)) {
                $parsedDate = [datetime]::MinValue
                if ([datetime]::TryParse($lastStatusRaw, [ref]$parsedDate) -and $parsedDate -ne [datetime]::MinValue) {
                    $daysSince = [int]((Get-Date) - $parsedDate).TotalDays
                    $statusAgePopulatedRows++
                }
            }
        }
        $o["DaysSinceLastStatus"] = $daysSince

        $action = Get-RemediationAction `
            -BlockingReason          $blockingReason `
            -AggregateState          "$($row.AggregateState)" `
            -UpgradeEligibilityLabel "$($row.UpgradeEligibilityLabel)"

        $o["BlockingReasonSort"] = switch ($blockingReason) {
            "HardFailure_DiskSpace"      { 1 }
            "HardFailure_Generic"        { 2 }
            "HardFailure_Compatibility"  { 3 }
            "HardFailure_SafeguardHold"  { 4 }
            "UpdateAlert"                { 5 }
            "DeploymentInProgress"       { 6 }
            "NotApplicableOrUnknown"     { 7 }
            default                      { 99 }
        }

        $o["ActionPriority"]    = $action.Priority
        $o["ActionCode"]        = $action.ActionCode
        $o["ActionDescription"] = $action.ActionDescription
        $o["ActionOwner"]       = $action.Owner

        $computedRows.Add([pscustomobject]$o) | Out-Null
    }
    $computedRowCount = $computedRows.Count
    $statusAgeCoveragePct = if ($computedRowCount -gt 0) { [math]::Round(100.0 * $statusAgePopulatedRows / $computedRowCount, 2) } else { 0 }
    $statusAgeSourceColumn = if ([string]::IsNullOrWhiteSpace($effectiveLastStatusDateColumn)) { 'Unavailable' } else { $effectiveLastStatusDateColumn }
    Write-Log "DaysSinceLastStatus coverage: $statusAgePopulatedRows/$computedRowCount row(s), $statusAgeCoveragePct%; source column: $statusAgeSourceColumn." "INFO" "ENRICH"
    $enrichedRows = @(Select-WinUpdateCanonicalRows -Rows $computedRows.ToArray())
    $duplicateRowsRemoved = $computedRowCount - $enrichedRows.Count
    if ($duplicateRowsRemoved -gt 0) {
        Write-Log ("Removed {0} duplicate policy/device row(s) using operational severity and action priority; canonical rows: {1}." -f $duplicateRowsRemoved, $enrichedRows.Count) "WARN" "DATA"
    }

    $count = $enrichedRows.Count
    Write-Log "Total consolidated canonical rows (enriched + computed): $count" "INFO" "DATA"
    Assert-WinUpdateOutputQuality -Rows $enrichedRows

    # ----------------------------------------------------------
    # CSV export - DATA-ALL (atomic), DATA-LAST (atomic), Archive
    # ----------------------------------------------------------
    $spUploadStatus = "Disabled"
    $spUploadError  = ""
    $spDestPath     = ""
    $spUploadUrl    = ""

    if ($DryRun) {
        Write-Log "DryRun: skipping write of $CsvFinal." "INFO" "DRYRUN"
    }
    else {
        Publish-CoreSmartM365Csv -Data $enrichedRows -TimestampedPath $CsvFinal -LatestPath $CsvLastFinal -NoSharePointUpload | Out-Null
        Write-Log "CSV exported: $CsvFinal" "INFO" "DATA"
        Write-Log "DATA-LAST updated: $CsvLastFinal" "INFO" "DATA"

        $archiveBase = [System.IO.Path]::GetFileNameWithoutExtension($CsvName)
        Write-ArchiveCopyAndPrune -SourceCsv $CsvFinal -ArchiveFolder $ArchivePath -BaseNameWithoutExt $archiveBase -RunStamp $RunStamp -RetentionDays $DetailedArchiveRetentionDays

        # ----------------------------------------------------------
        # SharePoint upload (non-blocking, independent historical and latest CSVs)
        # ----------------------------------------------------------
        if ($dataAllUploadEnabled) {
            try {
                $thumbprint = if ($global:Thumbprint) { $global:Thumbprint } else { $Thumb }
                $historyUpload = Invoke-CoreSmartM365SharePointCsvUpload `
                    -LocalFilePath $CsvFinal -Enabled $EnableSharePointUpload `
                    -SiteHostname $SP_SiteHostname -SitePath $SP_SitePath `
                    -LibraryDisplayName $SP_LibraryDisplayName -TargetFolderPath $SP_TargetFolderPath `
                    -AppId $AppId -TenantId $TenantId -Thumbprint $thumbprint
                if ($historyUpload) { Write-Log 'DATA-ALL SharePoint upload completed.' 'INFO' 'SP' }
                else { Write-Log 'DATA-ALL SharePoint upload returned no receipt.' 'WARN' 'SP' }
            }
            catch { Write-Log "DATA-ALL SharePoint upload failed (non-blocking): $($_.Exception.Message)" 'WARN' 'SP' }
        }
        if ($dataLastUploadEnabled) {
            try {
                Write-Log "SharePoint upload starting: $SP_SiteHostname$SP_SitePath / $SP_LibraryDisplayName / $SP_TargetFolderPath" "INFO" "SP"

                $thumbprint = if ($global:Thumbprint) { $global:Thumbprint } else { $Thumb }
                $spUploadRecord = Invoke-CoreSmartM365SharePointCsvUpload `
                    -LocalFilePath $CsvLastFinal `
                    -Enabled $EnableSharePointUpload `
                    -SiteHostname $SP_SiteHostname `
                    -SitePath $SP_SitePath `
                    -LibraryDisplayName $SP_LibraryDisplayName `
                    -TargetFolderPath $SP_TargetFolderPath `
                    -AppId $AppId `
                    -TenantId $TenantId `
                    -Thumbprint $thumbprint

                $spUploadStatus = if ($spUploadRecord -and ($spUploadRecord.WebUrl -or $spUploadRecord.ItemId)) { "Success" } else { "Unavailable" }
                $spUploadUrl    = if ($spUploadRecord) { [string]$spUploadRecord.WebUrl } else { '' }
                $spDestPath     = (($SP_TargetFolderPath.TrimEnd('/','\')) + "/" + [System.IO.Path]::GetFileName($CsvLastFinal)).TrimStart('/','\')
                Write-Log "SharePoint upload status: $spUploadStatus; destination=$spDestPath" "INFO" "SP"
            }
            catch {
                $spUploadStatus = "Error"
                $spUploadError  = $_.Exception.Message
                Write-Log "SharePoint upload failed (non-blocking): $spUploadError" "WARN" "SP"
            }
        }
    }

    if ($DryRun -and $EnableSharePointUpload) {
        $spUploadStatus = "Skipped (DryRun)"
        Write-Log "DryRun: skipping SharePoint upload." "INFO" "DRYRUN"
    }

    # ----------------------------------------------------------
    # Work cleanup
    # ----------------------------------------------------------
    Cleanup-WorkFolder -WorkPath $WorkPath -KeepZips $WorkKeepZips -KeepExtracts $WorkKeepExtracts

    # ----------------------------------------------------------
    # Operational summary by unique device
    # ----------------------------------------------------------
    $deviceSummaryRows = @(Get-WinUpdateDeviceSummaryRows -Rows $enrichedRows)
    $totalUniqueDevices = $deviceSummaryRows.Count
    $completedCount = @($deviceSummaryRows | Where-Object OperationalState -eq 'Completed').Count
    $inProgressCount = @($deviceSummaryRows | Where-Object OperationalState -eq 'InProgress').Count
    $actionRequiredCount = @($deviceSummaryRows | Where-Object OperationalState -eq 'ActionRequired').Count
    $priority0Count = @($deviceSummaryRows | Where-Object { $_.OperationalState -eq 'ActionRequired' -and $_.ActionPriority -eq 0 }).Count
    $unknownCount = @($deviceSummaryRows | Where-Object OperationalState -eq 'Unknown').Count
    $completionPct = if ($totalUniqueDevices -gt 0) { [math]::Round((100.0 * $completedCount / $totalUniqueDevices),2) } else { 0.0 }
    $allPolicySummary = @(Get-WinUpdatePolicySummary -Rows $enrichedRows |
        Sort-Object @{Expression='ActionRequired';Descending=$true},@{Expression='InProgress';Descending=$true},PolicyName)
    $primaryCoveragePolicy = Select-WinUpdatePrimaryCoveragePolicy -PolicySummary $allPolicySummary

    $offeringCount = @($deviceSummaryRows | Where-Object ProgressPhase -eq 'Offering').Count
    $installingCount = @($deviceSummaryRows | Where-Object ProgressPhase -eq 'Installing').Count
    $pendingCount = @($deviceSummaryRows | Where-Object ProgressPhase -eq 'Pending').Count
    $otherInProgressCount = @($deviceSummaryRows | Where-Object ProgressPhase -eq 'Other').Count

    $coverageAvailable = $null -ne $primaryCoveragePolicy
    if ($coverageAvailable) {
        $fleetPolicyName = [string]$primaryCoveragePolicy.PolicyName
        $fleetTargetLabel = [string]$primaryCoveragePolicy.TargetLabel
        $fleetTargetShortLabel = [string]$primaryCoveragePolicy.TargetShortLabel
        $fleetDevices = [int]$primaryCoveragePolicy.Devices
        $fleetOsCovered = [int]$primaryCoveragePolicy.OsCovered
        $fleetOsBelowTarget = [int]$primaryCoveragePolicy.OsBelowTarget
        $fleetOsVersionUnknown = [int]$primaryCoveragePolicy.OsVersionUnknown
        $fleetWindows11 = [int]$primaryCoveragePolicy.Windows11
        $fleetWindows10 = [int]$primaryCoveragePolicy.Windows10
        $fleetUnknownOrOther = [int]$primaryCoveragePolicy.UnknownOrOther
        $fleetWindows11Pct = [double]$primaryCoveragePolicy.Windows11Pct
        $fleetWindows10Pct = [double]$primaryCoveragePolicy.Windows10Pct
        $fleetUnknownOrOtherPct = [double]$primaryCoveragePolicy.UnknownOrOtherPct
        $fleetOsCoveragePct = [double]$primaryCoveragePolicy.OsCoveragePct
        $fleetKnownOsCoveragePct = [double]$primaryCoveragePolicy.KnownOsCoveragePct
        $fleetPolicyCompleted = [int]$primaryCoveragePolicy.Completed
        $fleetPolicyCompletionPct = [double]$primaryCoveragePolicy.PolicyCompletionPct
    } else {
        $fleetPolicyName = ''
        $fleetTargetLabel = ''
        $fleetTargetShortLabel = ''
        $fleetDevices = 0
        $fleetOsCovered = 0
        $fleetOsBelowTarget = 0
        $fleetOsVersionUnknown = 0
        $fleetWindows11 = 0
        $fleetWindows10 = 0
        $fleetUnknownOrOther = 0
        $fleetWindows11Pct = 0.0
        $fleetWindows10Pct = 0.0
        $fleetUnknownOrOtherPct = 0.0
        $fleetOsCoveragePct = 0.0
        $fleetKnownOsCoveragePct = 0.0
        $fleetPolicyCompleted = 0
        $fleetPolicyCompletionPct = 0.0
    }
    $reportHealth = Get-WinUpdateReportHealth `
        -TotalUniqueDevices $totalUniqueDevices `
        -ActionRequiredCount $actionRequiredCount `
        -Priority0Count $priority0Count `
        -CoverageAvailable $coverageAvailable `
        -KnownOsCoveragePct $fleetKnownOsCoveragePct `
        -FleetDevices $fleetDevices `
        -OsVersionUnknownCount $fleetOsVersionUnknown
    $reportStatus = $reportHealth.Status
    $summaryState = @(
        $totalUniqueDevices,$completedCount,$inProgressCount,$actionRequiredCount,$priority0Count,$unknownCount,
        $offeringCount,$installingCount,$pendingCount,$otherInProgressCount,
        $fleetTargetShortLabel,$fleetDevices,$fleetOsCovered,$fleetOsBelowTarget,$fleetOsVersionUnknown,
        $fleetWindows11,$fleetWindows10,$fleetUnknownOrOther,$fleetWindows11Pct,$fleetWindows10Pct,$fleetUnknownOrOtherPct,
        $fleetOsCoveragePct,$fleetPolicyCompleted,$fleetPolicyCompletionPct
    ) -join '|'

    $intuneFleetSummary = $null
    $adWithoutIntuneSummary = $null
    $combinedFleetSummary = $null
    $activationIndex = $null
    $intuneSourceText = ''
    $adSourceText = ''
    if ($EnableSummaryEmail -and $SummaryEmailMode -ne 'Never') {
        $intuneSourceItem = $null
        try {
            $intuneSourceItem = Get-WinUpdateFreshCsvItem -Path $IntuneDeviceInventoryCsvPath -MaxAgeHours $CountrySourceMaxAgeHours
            $entraSourceItem = Get-WinUpdateFreshCsvItem -Path $EntraDeviceInventoryCsvPath -MaxAgeHours $CountrySourceMaxAgeHours
            $adSourceItem = Get-WinUpdateFreshCsvItem -Path $AdComputersCsvPath -MaxAgeHours $CountrySourceMaxAgeHours
            if ($intuneDeviceInventoryRows.Count -eq 0) {
                $intuneDeviceInventoryRows = @(Import-Csv -LiteralPath $IntuneDeviceInventoryCsvPath -ErrorAction Stop)
            }
            if ($entraDeviceInventoryRows.Count -eq 0) {
                $entraDeviceInventoryRows = @(Import-Csv -LiteralPath $EntraDeviceInventoryCsvPath -ErrorAction Stop)
            }
            $adComputerRows = @(Import-Csv -LiteralPath $AdComputersCsvPath -ErrorAction Stop)
            if ($intuneDeviceInventoryRows.Count -eq 0 -or $entraDeviceInventoryRows.Count -eq 0 -or $adComputerRows.Count -eq 0) {
                throw 'An Intune, Entra, or AD activation source CSV has no data rows.'
            }
            foreach ($required in @('TenantKey','Device ID','OS','OS version','UserId','Azure AD Device ID')) {
                if ($intuneDeviceInventoryRows[0].PSObject.Properties.Name -notcontains $required) { throw "Intune device CSV lacks '$required'." }
            }
            foreach ($required in @('TenantKey','DeviceId','AccountEnabled')) {
                if ($entraDeviceInventoryRows[0].PSObject.Properties.Name -notcontains $required) { throw "Entra device CSV lacks '$required'." }
            }
            foreach ($required in @('TenantKey','ObjectGUID','Name','OperatingSystemShortName','Enabled','IsActiveInLast45Days')) {
                if ($adComputerRows[0].PSObject.Properties.Name -notcontains $required) { throw "AD computer CSV lacks '$required'." }
            }
            $activationIndex = New-WinUpdateActivationIndex -EntraDeviceRows $entraDeviceInventoryRows -AdComputerRows $adComputerRows -TenantKey $CsvTenantKey
        }
        catch {
            $activationIndex = $null
            Write-Log "Directory activation comparison unavailable: $($_.Exception.Message)" "WARN" "KPI"
        }
        if ($activationIndex) {
            try {
                $userSourceItem = Get-WinUpdateFreshCsvItem -Path $ActiveUsersCsvPath -MaxAgeHours $CountrySourceMaxAgeHours
                $activeUserRows = @(Import-Csv -LiteralPath $ActiveUsersCsvPath -ErrorAction Stop)
                if ($activeUserRows.Count -eq 0) { throw 'Active-user CSV has no data rows.' }
                foreach ($required in @('TenantKey','Object Id','CountryOrRegion')) {
                    if ($activeUserRows[0].PSObject.Properties.Name -notcontains $required) { throw "Active-user CSV lacks '$required'." }
                }
                $referencePolicyId = if ($coverageAvailable) { [string]$primaryCoveragePolicy.PolicyId } else { '' }
                $intuneFleetSummary = Get-WinUpdateIntuneFleetCountrySummary -IntuneDeviceRows $intuneDeviceInventoryRows -ActiveUserRows $activeUserRows -PolicyRows $enrichedRows -ActivationIndex $activationIndex -ReferencePolicyId $referencePolicyId -TenantKey $CsvTenantKey
                $intuneTotal = $intuneFleetSummary.Total
                if (($intuneTotal.Windows11 + $intuneTotal.Windows10 + $intuneTotal.UnknownOrOther) -ne $intuneTotal.Total -or
                    ($intuneTotal.ReferencePolicy + $intuneTotal.OtherPolicyOnly + $intuneTotal.NeitherPolicy) -ne $intuneTotal.Total -or
                    ($intuneTotal.Total + $intuneFleetSummary.ExcludedDisabledBoth + $intuneFleetSummary.ExcludedUnqualified) -ne $intuneFleetSummary.AllIntuneWindows -or
                    ($coverageAvailable -and ($intuneTotal.ReferencePolicy + $intuneFleetSummary.ExcludedReferencePolicy + $intuneFleetSummary.ReportOnlyReferenceIds) -ne $fleetDevices)) {
                    throw "Intune fleet totals do not reconcile: Enabled=$($intuneTotal.Total) All=$($intuneFleetSummary.AllIntuneWindows) DisabledBoth=$($intuneFleetSummary.ExcludedDisabledBoth) Unqualified=$($intuneFleetSummary.ExcludedUnqualified) OS=$($intuneTotal.Windows11 + $intuneTotal.Windows10 + $intuneTotal.UnknownOrOther) Policy=$($intuneTotal.ReferencePolicy + $intuneTotal.OtherPolicyOnly + $intuneTotal.NeitherPolicy) ReferenceEnabled=$($intuneTotal.ReferencePolicy) ReferenceExcluded=$($intuneFleetSummary.ExcludedReferencePolicy) ReferenceReportOnly=$($intuneFleetSummary.ReportOnlyReferenceIds) ReferenceReport=$fleetDevices."
                }
                $intuneSourceText = 'Intune devices: {0}; Entra devices: {1}; AD computers: {2}; active users: {3}' -f $intuneSourceItem.LastWriteTime.ToString('yyyy-MM-dd HH:mm'),$entraSourceItem.LastWriteTime.ToString('yyyy-MM-dd HH:mm'),$adSourceItem.LastWriteTime.ToString('yyyy-MM-dd HH:mm'),$userSourceItem.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
                Write-Log "Enabled Intune Windows breakdown: Devices=$($intuneTotal.Total) Windows11=$($intuneTotal.Windows11) Windows10=$($intuneTotal.Windows10) OsUnknown=$($intuneTotal.UnknownOrOther) ReferencePolicy=$($intuneTotal.ReferencePolicy) OtherPolicyOnly=$($intuneTotal.OtherPolicyOnly) NeitherPolicy=$($intuneTotal.NeitherPolicy) ExcludedDisabledBoth=$($intuneFleetSummary.ExcludedDisabledBoth) ExcludedUnqualified=$($intuneFleetSummary.ExcludedUnqualified) CountryUnknown=$($intuneFleetSummary.Unknown.Total) ReportOnlyReferenceIds=$($intuneFleetSummary.ReportOnlyReferenceIds)." "INFO" "KPI"
            }
            catch {
                $intuneFleetSummary = $null
                Write-Log "Enabled Intune fleet breakdown unavailable: $($_.Exception.Message)" "WARN" "KPI"
            }
            try {
                $adWithoutIntuneSummary = Get-WinUpdateAdWithoutIntuneSummary -AdComputerRows $adComputerRows -IntuneDeviceRows $intuneDeviceInventoryRows -ActivationIndex $activationIndex -TenantKey $CsvTenantKey -CountryPrefixMap $AdComputerNameCountryPrefixes
                if (($adWithoutIntuneSummary.All.Total + $adWithoutIntuneSummary.ExcludedDisabledBoth + $adWithoutIntuneSummary.ExcludedUnqualified) -ne $adWithoutIntuneSummary.UnmatchedTotal -or
                    $adWithoutIntuneSummary.EnabledActive45.Total -gt $adWithoutIntuneSummary.All.Total) {
                    throw 'Enabled AD/Intune comparison totals do not reconcile.'
                }
                $adSourceText = 'Intune devices: {0}; Entra devices: {1}; AD computers: {2}' -f $intuneSourceItem.LastWriteTime.ToString('yyyy-MM-dd HH:mm'),$entraSourceItem.LastWriteTime.ToString('yyyy-MM-dd HH:mm'),$adSourceItem.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
                Write-Log "Enabled AD computers without exact Intune ID match: Windows11=$($adWithoutIntuneSummary.All.Windows11) Windows10=$($adWithoutIntuneSummary.All.Windows10) Total=$($adWithoutIntuneSummary.All.Total) CountryUnknown=$($adWithoutIntuneSummary.Unknown.Total) EnabledActive45=$($adWithoutIntuneSummary.EnabledActive45.Total) ExcludedDisabledBoth=$($adWithoutIntuneSummary.ExcludedDisabledBoth) ExcludedUnqualified=$($adWithoutIntuneSummary.ExcludedUnqualified) AmbiguousIntuneIds=$($adWithoutIntuneSummary.AmbiguousIntuneIds)." "INFO" "KPI"
            }
            catch {
                $adWithoutIntuneSummary = $null
                Write-Log "Enabled AD/Intune comparison unavailable: $($_.Exception.Message)" "WARN" "KPI"
            }
        }
        if ($intuneFleetSummary -and $adWithoutIntuneSummary) {
            try {
                $combinedFleetSummary = Get-WinUpdateEnabledWindowsCombinedSummary -IntuneSummary $intuneFleetSummary -AdWithoutIntuneSummary $adWithoutIntuneSummary
                Write-Log "Enabled Windows exact-ID combined scope: Intune=$($intuneFleetSummary.Total.Total) AdWithoutIntune=$($adWithoutIntuneSummary.All.Total) Windows11=$($combinedFleetSummary.Windows11) Windows10=$($combinedFleetSummary.Windows10) UnknownOrOther=$($combinedFleetSummary.UnknownOrOther) Total=$($combinedFleetSummary.Total)." "INFO" "KPI"
            }
            catch {
                $combinedFleetSummary = $null
                Write-Log "Enabled Windows exact-ID combined scope unavailable: $($_.Exception.Message)" "WARN" "KPI"
            }
        }
        $fleetState = if ($intuneFleetSummary) {
            @(@($intuneFleetSummary.Known) + @($intuneFleetSummary.Unknown) | ForEach-Object {
                [ordered]@{ Country = $_.Country; Windows11 = $_.Windows11; Windows10 = $_.Windows10; UnknownOrOther = $_.UnknownOrOther; ReferencePolicy = $_.ReferencePolicy; OtherPolicyOnly = $_.OtherPolicyOnly; NeitherPolicy = $_.NeitherPolicy }
            }) | ConvertTo-Json -Compress -Depth 3
        } else { 'Unavailable' }
        $adState = if ($adWithoutIntuneSummary) {
            '{0},{1},{2},{3},{4},{5}' -f $adWithoutIntuneSummary.All.Windows11,$adWithoutIntuneSummary.All.Windows10,$adWithoutIntuneSummary.All.Total,$adWithoutIntuneSummary.EnabledActive45.Windows11,$adWithoutIntuneSummary.EnabledActive45.Windows10,$adWithoutIntuneSummary.EnabledActive45.Total
        } else { 'Unavailable' }
        $activationState = if ($intuneFleetSummary) { "$($intuneFleetSummary.ExcludedDisabledBoth),$($intuneFleetSummary.ExcludedUnqualified),$($intuneFleetSummary.ExcludedReferencePolicy),$($intuneFleetSummary.ReportOnlyReferenceIds)" } else { 'Unavailable' }
        $adExclusionState = if ($adWithoutIntuneSummary) { "$($adWithoutIntuneSummary.ExcludedDisabledBoth),$($adWithoutIntuneSummary.ExcludedUnqualified)" } else { 'Unavailable' }
        $combinedState = if ($combinedFleetSummary) {
            @(@($combinedFleetSummary.Known) + @($combinedFleetSummary.Unknown) | ForEach-Object {
                [ordered]@{ Country = $_.Country; Windows11 = $_.Windows11; Windows10 = $_.Windows10; UnknownOrOther = $_.UnknownOrOther; Intune = $_.Intune; AdWithoutIntune = $_.AdWithoutIntune }
            }) | ConvertTo-Json -Compress -Depth 3
        } else { 'Unavailable' }
        $summaryState = "$summaryState|IntuneFleet=$fleetState|Activation=$activationState|AdWithoutIntune=$adState|AdExcluded=$adExclusionState|CombinedFleet=$combinedState"
    }

    Write-Log "Operational policy-state summary: Devices=$totalUniqueDevices Completed=$completedCount InProgress=$inProgressCount Offering=$offeringCount Installing=$installingCount Pending=$pendingCount OtherInProgress=$otherInProgressCount ActionRequired=$actionRequiredCount ActionRequiredRate=$($reportHealth.ActionRequiredRatePct)% Priority0=$priority0Count Priority0CriticalThreshold=$($reportHealth.Priority0Threshold) Unknown=$unknownCount PolicyCompletion=$completionPct% Status=$reportStatus" "INFO" "KPI"
    if ($coverageAvailable) {
        Write-Log "Fleet OS coverage: ReferencePolicy='$fleetPolicyName' Target='$fleetTargetLabel' Devices=$fleetDevices Windows11=$fleetWindows11 Windows10=$fleetWindows10 UnknownOrOther=$fleetUnknownOrOther Covered=$fleetOsCovered BelowTarget=$fleetOsBelowTarget OsVersionUnknown=$fleetOsVersionUnknown Coverage=$fleetOsCoveragePct% KnownOsCoverage=$fleetKnownOsCoveragePct% IntunePolicyCompleted=$fleetPolicyCompleted IntunePolicyCompletion=$fleetPolicyCompletionPct%" "INFO" "KPI"
    } else {
        Write-Log "Fleet OS coverage unavailable: no supported Windows version was detected in a policy name." "WARN" "KPI"
    }

    $durationSeconds = [int](New-TimeSpan -Start $scriptStart -End (Get-Date)).TotalSeconds
    $policyCount = $policies.Count
    $actionLimit = [math]::Max(1,$RiskTopN)
    $actionRows = @($deviceSummaryRows | Where-Object OperationalState -eq 'ActionRequired' | Sort-Object ActionPriority,PolicyName,DeviceName | Select-Object -First $actionLimit `
        ActionPriority,DeviceName,PolicyName,AggregateState,CurrentDeviceUpdateStatus,LatestAlertMessage,UpgradeEligibilityLabel,ActionCode,ActionOwner,ActionDescription)
    $summaryMatched = @($deviceSummaryRows | Where-Object ReadinessMatch -eq 'Matched').Count
    $summaryNotMatched = @($deviceSummaryRows | Where-Object ReadinessMatch -eq 'NotMatched').Count
    $summaryEntraOsVersion = @($deviceSummaryRows | Where-Object OSVersionSource -eq 'Entra').Count
    $summaryUnavailableOsVersion = @($deviceSummaryRows | Where-Object OSVersionSource -eq 'Unavailable').Count
    $summaryReadinessPct = if (($summaryMatched+$summaryNotMatched) -gt 0) { [math]::Round(100.0*$summaryMatched/($summaryMatched+$summaryNotMatched),2) } else { 0 }
    $mailInputs = @{
        actionRequiredCount=$actionRequiredCount; adSourceText=$adSourceText; adWithoutIntuneSummary=$adWithoutIntuneSummary; allPolicySummary=$allPolicySummary
        combinedFleetSummary=$combinedFleetSummary; completedCount=$completedCount; completionPct=$completionPct; computedRowCount=$computedRowCount
        count=$count; coverageAvailable=$coverageAvailable; CsvLastFinal=$CsvLastFinal; durationSeconds=$durationSeconds
        EnableReadinessEnrichment=$EnableReadinessEnrichment; fleetDevices=$fleetDevices; fleetKnownOsCoveragePct=$fleetKnownOsCoveragePct; fleetOsBelowTarget=$fleetOsBelowTarget
        fleetOsCoveragePct=$fleetOsCoveragePct; fleetOsCovered=$fleetOsCovered; fleetOsVersionUnknown=$fleetOsVersionUnknown; fleetPolicyCompletionPct=$fleetPolicyCompletionPct
        fleetPolicyName=$fleetPolicyName; fleetTargetLabel=$fleetTargetLabel; fleetTargetShortLabel=$fleetTargetShortLabel; inProgressCount=$inProgressCount
        installingCount=$installingCount; intuneFleetSummary=$intuneFleetSummary; intuneSourceText=$intuneSourceText; offeringCount=$offeringCount
        OrgDomain=$OrgDomain; otherInProgressCount=$otherInProgressCount; pendingCount=$pendingCount; policyCount=$policyCount
        primaryCoveragePolicy=$primaryCoveragePolicy; priority0Count=$priority0Count; readinessMode=$readinessMode; reportHealth=$reportHealth
        reportStatus=$reportStatus; RunId=$RunId; spUploadUrl=$spUploadUrl; statusAgeCoveragePct=$statusAgeCoveragePct
        statusAgePopulatedRows=$statusAgePopulatedRows; statusAgeSourceColumn=$statusAgeSourceColumn; totalUniqueDevices=$totalUniqueDevices; unknownCount=$unknownCount
        actionRows=$actionRows; actionLimit=$actionLimit; summaryMatched=$summaryMatched; summaryNotMatched=$summaryNotMatched
        summaryEntraOsVersion=$summaryEntraOsVersion; summaryUnavailableOsVersion=$summaryUnavailableOsVersion; summaryReadinessPct=$summaryReadinessPct
    }
    if (-not $DryRun) {
        Save-SmartM365ReportInputs -Producer 'SmartM365-WinUpdate_Status_From_Intune' -LatestFolder $LatestCsvFolderPath -TenantKey $global:SmartM365TenantKey -Files $CsvName -InputData $mailInputs
    }
    $should = Should-SendSummaryEmail -CurrentState $summaryState -Mode $SummaryEmailMode -StatePath $SummaryStatePath
    if ($should) {
        $reportMail = New-SmartM365WindowsUpdateMailReport @mailInputs
        if (Send-SummaryEmail -Subject $reportMail.Subject -HtmlBody $reportMail.BodyHtml) {
            Save-SummaryState -State $summaryState -StatePath $SummaryStatePath
        }
    }
    else {
        Write-Log "Summary email skipped (Mode=$SummaryEmailMode)." "INFO" "MAIL"
    }

    Write-Log "Completed successfully. Version=$ScriptVersion DryRun=$DryRun Rows=$count UniqueDevices=$totalUniqueDevices CoverageTarget='$fleetTargetLabel' TargetedDevices=$fleetDevices OsCovered=$fleetOsCovered OsCoverage=$fleetOsCoveragePct% LastStatusDateSource=$statusAgeSourceColumn LastStatusAgeCoverage=$statusAgePopulatedRows/$computedRowCount ($statusAgeCoveragePct%) IntunePolicyCompleted=$fleetPolicyCompleted IntunePolicyCompletion=$fleetPolicyCompletionPct% ActionRequired=$actionRequiredCount" "INFO"

    $smartM365RunGuard.ReachedTerminal=$true
}
catch {
    if ($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]) { throw }
    $smartM365RunGuard.ErrorRecord=$_

    $script:CompletionStatus = 'Failed'
    $err = $_
    Write-Log "Fatal error: $($err.Exception.Message)" "ERROR"

    $duration = New-TimeSpan -Start $scriptStart -End (Get-Date)
    $subject  = "[SmartM365] WinUpdate Feature Update status - FAILED - $ScriptName"
    $html = @"
<html><body style="font-family:Segoe UI,Arial;">
<h2 style="margin:0 0 10px 0;color:#b00020;">Windows Update Feature Update status - FAILED</h2>
<ul>
<li><b>Script</b>: $(Html-Encode $ScriptName) v$(Html-Encode $ScriptVersion)</li>
<li><b>RunId</b>: $(Html-Encode $RunId)</li>
<li><b>Host</b>: $(Html-Encode $env:COMPUTERNAME)</li>
<li><b>Execution time</b>: $([int]$duration.TotalSeconds) seconds</li>
</ul>
<h3 style="margin:16px 0 8px 0;">Error</h3>
<pre style="white-space:pre-wrap;background:#f6f6f6;border:1px solid #ddd;padding:10px;">$(Html-Encode ($err | Out-String))</pre>
</body></html>
"@

    Send-FatalErrorEmail -Subject $subject -HtmlBody $html
    throw
}
finally {
    $smartM365RunGuard.InCleanup=$true

    Cleanup-TempFiles
    $sourceReceipt = $null
    if(Get-Command Complete-CoreSmartM365SourceReceipt -ErrorAction SilentlyContinue){
        $sourceReceipt = Complete-CoreSmartM365SourceReceipt -Status $(if($script:CompletionStatus -eq 'Auto'){'Success'}else{$script:CompletionStatus}) -ErrorCount $global:SmartM365ErrorCount
    }
    Write-SmartM365CompletionBanner -Status $script:CompletionStatus -ScriptName $ScriptName -StartedAt $scriptStart -LogPath $LogPath

    if ($EnableSharePointUpload -and (Get-Command Invoke-CoreSmartM365SharePointCsvUpload -ErrorAction SilentlyContinue)) {
        $runArtifacts = @($LogPath)
        if($sourceReceipt){
            try { $runArtifacts += @(Get-WinUpdateSourceReceiptUploadPaths -ReceiptPath $sourceReceipt) }
            catch { Write-Log "Source receipt upload skipped: $($_.Exception.Message)" "WARN" "SP" }
        }
        if ($global:SmartM365MailHtmlFiles) {
            $runArtifacts += @($global:SmartM365MailHtmlFiles)
        }

        foreach ($runArtifact in @($runArtifacts | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) -and (Test-Path -LiteralPath $_ -PathType Leaf) } | Sort-Object -Unique)) {
            try {
                Invoke-CoreSmartM365SharePointCsvUpload `
                    -LocalFilePath $runArtifact `
                    -Enabled $true `
                    -SiteHostname $SP_SiteHostname `
                    -SitePath $SP_SitePath `
                    -LibraryDisplayName $SP_LibraryDisplayName `
                    -TargetFolderPath $SP_TargetFolderPath `
                    -AppId $AppId `
                    -TenantId $TenantId `
                    -Thumbprint $Thumb | Out-Null
            }
            catch {
                Write-Log "Run artifact SharePoint upload failed for '$runArtifact': $($_.Exception.Message)" "WARN" "SP"
            }
        }
    }
}
    $smartM365RunGuard.ReachedTerminal=$true
    $smartM365RunGuard.ReturnedNormally=$true
} catch {
    if ($_.Exception -isnot [Management.Automation.PipelineStoppedException] -and $_.Exception -isnot [OperationCanceledException]) { $smartM365RunGuard.ErrorRecord=$_ }
    throw
} finally {
    Complete-SmartM365InventoryRun -State $smartM365RunGuard
    if($smartM365RunGuard.ExitCode -eq 130){exit 130}
}
if($smartM365RunGuard.ExitCode -ne 0){exit $smartM365RunGuard.ExitCode}

# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCGR8ZpUdx2gco6
# Q8IXSF13IqlRJ4ygoz+MRYsSXsk/5aCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgZeWsp0DbRQQP
# FEtvqlq3TfISbnsRfWm5bhXOYOGwnvgwDQYJKoZIhvcNAQEBBQAEggGAMVh40LDK
# c1y9jaVGqqlxDHYDhBaP1T97VFMgEVGppirW98lDkTkzPCHp1mSlU3Iv6cQp1Nad
# 0zn4J5+gQtwe1dMR/Rnykbz3b3SlsqS0q6wVir9WMAu16AMRsaI5IQhgm061LZ60
# UNBtQCWP8JxyWaYTenrMC6AidPxwSF9Qc44P1y0i808jhQBsE1oT008lNz7190Vg
# hVGB9ZyWu6cexXk4fPR6EFll0jGYnt6nZcJwWuaGtxHUrjyvLazGUmONVr7paWSF
# sxXTPSbM56uYOdckj28gmDTjrkKXKagw6GfQi06TiTZcrVvTZx4iFibQWPS8lEn/
# Ald+WEp8ChJ1JA8m4YDUM2BUnoh2woUPP2dy8IFlvf7bMzryeXT21Vk7hXYP3GN4
# otHzFkyGwmchPksCN37QvwPSBfI2AB1R2Sh7WkzsOt9l2YJfRMuV06njT6jz/NMJ
# H1HpCcpLRfuMGCEqJy+lty6+SPjPYnkhgCGHh+B2kPSk3i2S8QDfQpQ2oYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTE5MDkxNDU3WjA/BgkqhkiG9w0BCQQxMgQwuSM+eLnEWX31qnmtRfgBF+4U
# j945KyGTfBWWb/LXSPu3IFxJtMJMgm9kIxWJF5shMA0GCSqGSIb3DQEBAQUABIIC
# AErjmj79X0fcFD6gjXBkThFRfuFLCzHle8nrb8eeDv4qjdEx76umuyNsS4ERYiL1
# RtaUKIGmKYdVYk5dZDw13738tpvqzXdo64dak2DEW4h1CMSR3Stb10Awj709bstb
# mQuZ8DpHwL/LoF9k5Ufo60zhw9HtymbDGg3bEhddjF+SH8UTvf4/vXzN7bqei1Jh
# xrRW7Kx+d57jT2NLuxtpiUfBfsTd+uVkVj/WUMLv5wE81+QnXqI1zvd5qXTWsSG4
# S10glWIV1XZbgIaMsOvCVn21RkldsYBOd1Z+UQG1LUnx811kjCdlphhC8Hyog+Y7
# kVZ813pyxs9TUyXEOgpCXjGL0EphpiSERAv3ITAC6sYjiNlBl4BAtjiZgTGnpipJ
# KlwOXfHIv1uj4IfFRlrHlzhGRY6i6ExqvD5k0eVK72QUVtWKOKohTShGCiLXV2BS
# FUAy1RJJXpkae6oIgQa6ypXRwQaXYlPsCcOUFt3C7keczUBYgMcz71VNO3eqRabo
# B74yS4/K2giXnzVEXsjscyFAKNKwDeqoyxGt6UFxDBOSac3b21Zy74bCyePIZFcP
# Ise5PUoxoRBV/2Y8hCl27V2y3JcJefIkxmA6fy0kBI2DhqCBwgooM4TMFgZzXJij
# qlmqmWwAhKrC7zoRdwCzaQWeGaJ7aZ1Mq0QDOHEw2Qbn
# SIG # End signature block
