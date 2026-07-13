<#
.SYNOPSIS
Windows Update (Feature Update) device status from Intune with Win11 Readiness enrichment
and full remediation columns. Single CSV output for Power BI consumption.

.DESCRIPTION
- Discovers Feature Update policies via Graph beta: deviceManagement/windowsFeatureUpdateProfiles
- For each policy: exports Intune report via Graph beta reports/exportJobs (required filter: PolicyId eq '<id>')
- Downloads ZIP, extracts CSV, consolidates into one dataset
- Optional enrichment with Win11 Readiness CSV (join by NormalizedDeviceName derived from DeviceName)
- Computed columns added on all rows:
    RiskBucket, BlockingReason, OSVersion, DaysSinceLastStatus,
    ActionPriority, ActionCode, ActionDescription, ActionOwner
- Single output CSV (DATA-ALL + DATA-LAST + Archive + SharePoint):
    Intune_WindowsUpdate_Status.csv
- All summaries, blocking splits, action plans are derived in Power BI from this single table
- Token auto-refresh: Graph token refreshed automatically if age > 55 min
- SharePoint upload: latest CSV copied to SharePoint Online via SmartM365.Core when enabled
- Emails:
    - Fatal error HTML email (on failure)
    - Optional success/summary HTML email with KPI tables

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
  -RiskTopN                  : Number of top-risk policies shown in email (default: 10)

VERSION
  1.5
.VERSION
1.14

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
    [int]$MaxItems = 0
)
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

# ==========================================================
# Version
# ==========================================================
$ScriptVersion = "1.14"

# ==========================================================
# App-only authentication parameters
# ==========================================================
function Get-ScriptLocalConfig {
    [CmdletBinding()]
    param()

    $configPath = Join-Path -Path $PSScriptRoot -ChildPath ("{0}.local.json" -f [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))
    if (-not (Test-Path -LiteralPath $configPath)) {
        $templatePath = '{0}.template' -f $configPath
        if (Get-Command Initialize-SmartM365LocalJsonFromTemplate -ErrorAction SilentlyContinue) {
            Initialize-SmartM365LocalJsonFromTemplate -Path $configPath -TemplatePath $templatePath -ConfigDescription 'script local configuration' | Out-Null
        }
        else {
            if (-not (Test-Path -LiteralPath $templatePath)) {
                $message = @(
                    "Local configuration file not found: $configPath",
                    "Template to copy is also missing: $templatePath",
                    'Create the .local.json file from a safe template, then run the script again.'
                ) -join [Environment]::NewLine
                throw $message
            }

            Copy-Item -LiteralPath $templatePath -Destination $configPath -ErrorAction Stop
            Write-Host ("Created script local configuration from template: {0}" -f $configPath) -ForegroundColor Yellow
            Write-Host 'Review the generated local JSON values; continuing with current file values.' -ForegroundColor Yellow
        }
    }

    try {
        $config = Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        if (Get-Command Sync-SmartM365JsonConfigWithTemplate -ErrorAction SilentlyContinue) {
            return (Sync-SmartM365JsonConfigWithTemplate -Config $config -Path $configPath)
        }
        return $config
    }
    catch {
        throw ("Failed to read local configuration '{0}': {1}" -f $configPath, $_.Exception.Message)
    }
}

function Resolve-SmartM365ConfigValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) {
        return $Value
    }

    if ($Value -notmatch '\{\{[^}]+\}\}') {
        return $Value
    }

    if ($null -eq $script:SmartM365GlobalConfig) {
        $script:SmartM365GlobalConfig = [pscustomobject]@{}
            $scriptRootValue = Get-Variable -Name ScriptRoot -ValueOnly -ErrorAction SilentlyContinue
    $searchRoot = if ($PSScriptRoot) { $PSScriptRoot } elseif ($scriptRootValue) { $scriptRootValue } elseif ($PSCommandPath) { Split-Path -Path $PSCommandPath -Parent } else { (Get-Location).Path }
        while ($searchRoot) {
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json'
            if (Test-Path -LiteralPath $globalConfigPath) {
                try {
                    $script:SmartM365GlobalConfig = Get-Content -LiteralPath $globalConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                }
                catch {
                    throw ("Failed to read global local configuration '{0}': {1}" -f $globalConfigPath, $_.Exception.Message)
                }
                break
            }
            $parent = Split-Path -Path $searchRoot -Parent
            if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
            $searchRoot = $parent
        }
    }

    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $matches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($matches.Count -eq 0) { break }

        $changed = $false
        foreach ($match in $matches) {
            $tokenName = $match.Groups['Name'].Value
            $tokenProperty = $script:SmartM365GlobalConfig.PSObject.Properties[$tokenName]
            if ($null -eq $tokenProperty -or $null -eq $tokenProperty.Value) { continue }

            $tokenValue = Resolve-SmartM365ConfigValue -Value $tokenProperty.Value
            if ($null -eq $tokenValue) { continue }

            $resolved = $resolved.Replace($match.Value, [string]$tokenValue)
            $changed = $true
        }

        if (-not $changed) { break }
    }

    return $resolved
}
function Test-SmartM365UseGlobalConfigValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -isnot [string]) { return $false }

    $text = $Value.Trim()
    return ($text -in @('__USE_GLOBAL__', 'USE_GLOBAL', '**USE_GLOBAL**', '**USE\_GLOBAL**'))
}

function Get-ScriptLocalConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$Name,
        $DefaultValue
    )

    $property = $Config.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        if ($property.Value -is [string]) {
            $localValue = $property.Value.Trim()
            if ($localValue -and -not (Test-SmartM365UseGlobalConfigValue -Value $localValue)) {
                return Resolve-SmartM365ConfigValue -Value $property.Value
            }
        }
        else {
            return Resolve-SmartM365ConfigValue -Value $property.Value
        }
    }


    if ($null -eq $script:SmartM365GlobalConfig) {
        $script:SmartM365GlobalConfig = [pscustomobject]@{}
        $searchRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Path $PSCommandPath -Parent }
        while ($searchRoot) {
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json'
            if (Test-Path -LiteralPath $globalConfigPath) {
                try {
                    $script:SmartM365GlobalConfig = Get-Content -LiteralPath $globalConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                }
                catch {
                    throw ("Failed to read global local configuration '{0}': {1}" -f $globalConfigPath, $_.Exception.Message)
                }
                break
            }
            $parent = Split-Path -Path $searchRoot -Parent
            if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
            $searchRoot = $parent
        }
    }

    $globalProperty = $script:SmartM365GlobalConfig.PSObject.Properties[$Name]
    if ($null -ne $globalProperty -and $null -ne $globalProperty.Value) {
        if ($globalProperty.Value -is [string]) {
            $globalValue = $globalProperty.Value.Trim()
            if ([string]::IsNullOrWhiteSpace($globalValue) -or (Test-SmartM365UseGlobalConfigValue -Value $globalValue)) {
                return $DefaultValue
            }
        }
        return Resolve-SmartM365ConfigValue -Value $globalProperty.Value
    }
    return $DefaultValue
}

$ScriptLocalConfig = Get-ScriptLocalConfig


$global:RetentionMaxCSV = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxLogs' -DefaultValue 30)

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

$SummaryStatePath = Join-Path $ScriptCsvLogFolderPath "Intune_WindowsUpdate_Status.lastcount.txt"

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
            Import-Module -Name $modulePath -MinimumVersion '1.0.24' -Prefix Core -ErrorAction Stop
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
            try { Remove-Item -Path $p -Force -ErrorAction Stop } catch { }
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
        $mailParams = @{ To = $ErrorMailTo; Subject = $Subject; BodyHtml = $HtmlBody }
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

    if (-not $EnableSummaryEmail) { return }
    if ($DryRun) { Write-Log "DryRun: skipping summary email." "INFO" "DRYRUN"; return }
    $summaryTo = $To
    if ([string]::IsNullOrWhiteSpace($summaryTo)) { throw "Summary email requires To in local or global configuration." }

    try {
        Import-SmartM365CorePreflight
        $mailParams = @{ To = $summaryTo; Subject = $Subject; BodyHtml = $HtmlBody }
        if (-not [string]::IsNullOrWhiteSpace($From)) { $mailParams['From'] = $From }
        if (-not [string]::IsNullOrWhiteSpace($Cc)) { $mailParams['Cc'] = $Cc }
        Send-CoreSmartM365Mail @mailParams
        Write-Log "Summary email sent." "INFO" "MAIL"
    }
    catch {
        Write-Log "Failed to send summary email: $($_.Exception.Message)" "ERROR" "MAIL"
        throw
    }
}

function Should-SendSummaryEmail {
    param([Parameter(Mandatory=$true)][int]$CurrentCount,[Parameter(Mandatory=$true)][string]$Mode,[Parameter(Mandatory=$true)][string]$StatePath)

    if (-not $EnableSummaryEmail) { return $false }
    switch ($Mode) {
        "Never"    { return $false }
        "Always"   { return $true }
        "OnChange" {
            $previous = $null
            if (Test-Path $StatePath) { $previous = (Get-Content -Path $StatePath -ErrorAction SilentlyContinue | Select-Object -First 1) }
            if ([string]::IsNullOrWhiteSpace($previous)) { return $true }
            $prevInt = 0
            if (-not [int]::TryParse($previous, [ref]$prevInt)) { return $true }
            return ($prevInt -ne $CurrentCount)
        }
        default { return $false }
    }
}

function Save-SummaryState {
    param([Parameter(Mandatory=$true)][int]$Count,[Parameter(Mandatory=$true)][string]$StatePath)
    if ($DryRun) { return }
    try { Set-Content -Path $StatePath -Value $Count -Encoding UTF8 } catch { }
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

function Get-GraphTokenWithCert {
    param(
        [Parameter(Mandatory=$true)][string]$TenantId,
        [Parameter(Mandatory=$true)][string]$ClientId,
        [Parameter(Mandatory=$true)][System.Security.Cryptography.X509Certificates.X509Certificate2]$Certificate
    )

    if (-not (Get-Module -ListAvailable -Name "MSAL.PS")) {
        throw "MSAL.PS module is not installed. Install it with: Install-Module MSAL.PS -Scope AllUsers"
    }
    Import-Module MSAL.PS -ErrorAction Stop
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

    try {
        if ($ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.StatusCode) {
            return [int]$ErrorRecord.Exception.Response.StatusCode
        }
    } catch { }
    return $null
}

function Get-GraphRetryDelaySeconds {
    param(
        [Parameter(Mandatory=$true)]$ErrorRecord,
        [Parameter(Mandatory=$true)][int]$Attempt
    )

    try {
        $retryAfter = $ErrorRecord.Exception.Response.Headers.RetryAfter
        if ($retryAfter) {
            if ($retryAfter.Delta) { return [math]::Max(1, [int][math]::Ceiling($retryAfter.Delta.TotalSeconds)) }
            if ($retryAfter.Date) {
                $seconds = [int][math]::Ceiling(($retryAfter.Date.UtcDateTime - [datetime]::UtcNow).TotalSeconds)
                if ($seconds -gt 0) { return $seconds }
            }
        }
    } catch { }

    return [math]::Min(90, (10 * $Attempt) + (Get-Random -Minimum 1 -Maximum 5))
}

function Test-GraphTransientError {
    param([Parameter(Mandatory=$true)]$ErrorRecord)

    $statusCode = Get-GraphHttpStatusCode -ErrorRecord $ErrorRecord
    return ($statusCode -eq 429 -or ($statusCode -ge 500 -and $statusCode -le 599))
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
    $next = $Uri
    while ($next) {
        $resp = Invoke-GraphRestMethodWithRetry -Method GET -Uri $next -Headers $Headers -Operation "Graph paging GET"
        if ($resp.value) { $resp.value | ForEach-Object { $all.Add($_) | Out-Null } }
        $next = $resp.'@odata.nextLink'
    }
    return $all
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
    param([Parameter(Mandatory=$true)][string]$JobId,[Parameter(Mandatory=$true)][hashtable]$Headers)
    return Invoke-GraphRestMethodWithRetry -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs('$JobId')" -Headers $Headers -Operation "Get Intune export job"
}

function Wait-IntuneExportJobComplete {
    param(
        [Parameter(Mandatory=$true)][string]$JobId,
        [Parameter(Mandatory=$true)][hashtable]$Headers,
        [Parameter()][int]$TimeoutSeconds = 900,
        [Parameter()][int]$PollSeconds = 5
    )

    $start = Get-Date
    while ($true) {
        $job    = Get-IntuneExportJob -JobId $JobId -Headers $Headers
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
    param([Parameter(Mandatory=$true)][string]$DownloadUrl,[Parameter(Mandatory=$true)][string]$DestinationZipPath,[Parameter(Mandatory=$true)][hashtable]$Headers)

    if (Test-Path $DestinationZipPath) { Remove-Item -Path $DestinationZipPath -Force -ErrorAction SilentlyContinue }
    Write-Log "Downloading export ZIP..." "INFO" "DATA"
    try { Invoke-WebRequest -Uri $DownloadUrl -OutFile $DestinationZipPath -UseBasicParsing }
    catch {
        Write-Log "Download without headers failed, retrying with Authorization header." "WARN" "DATA"
        Invoke-WebRequest -Uri $DownloadUrl -OutFile $DestinationZipPath -UseBasicParsing -Headers $Headers
    }
    if (-not (Test-Path $DestinationZipPath)) { throw "ZIP download failed: $DestinationZipPath not found." }
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
        [Parameter(Mandatory=$true)][int]$Keep
    )

    Ensure-Directory -Path $ArchiveFolder
    $archiveFinal = Join-Path $ArchiveFolder ("{0}_{1}.csv"     -f $BaseNameWithoutExt, $RunStamp)
    $archiveTemp  = Join-Path $ArchiveFolder ("{0}_{1}.csv.tmp" -f $BaseNameWithoutExt, $RunStamp)
    Copy-FileAtomic -SourcePath $SourceCsv -DestinationFinal $archiveFinal -DestinationTemp $archiveTemp
    Write-Log "Archive CSV created: $archiveFinal" "INFO" "DATA"
    Prune-Files -Folder $ArchiveFolder -Filter ("{0}_*.csv" -f $BaseNameWithoutExt) -Keep $Keep
}

# ==========================================================
# HTML helpers
# ==========================================================
function Html-Encode {
    param([Parameter(Mandatory=$false)][object]$Value)
    if ($null -eq $Value) { return "" }
    $s = [string]$Value
    $s = $s.Replace("&","&amp;").Replace("<","&lt;").Replace(">","&gt;").Replace('"',"&quot;").Replace("'","&#39;")
    return $s
}

function Convert-ObjectsToHtmlTable {
    param(
        [Parameter(Mandatory=$true)][AllowNull()][AllowEmptyCollection()][object[]]$Rows,
        [Parameter(Mandatory=$true)][string[]]$Columns,
        [Parameter(Mandatory=$true)][string]$Title
    )

    if ($null -eq $Rows) { $Rows = @() }
    if (-not $Rows -or $Rows.Count -eq 0) {
        return "<h3 style='margin:16px 0 8px 0;'>$(Html-Encode $Title)</h3><div style='font-family:Segoe UI,Arial;font-size:12px;color:#666;'>No data.</div>"
    }

    $th  = ($Columns | ForEach-Object { "<th style='padding:6px;border:1px solid #ddd;background:#f5f5f5;text-align:left;'>$(Html-Encode $_)</th>" }) -join ""
    $trs = ($Rows | ForEach-Object {
        $row = $_
        $tds = ($Columns | ForEach-Object { "<td style='padding:6px;border:1px solid #ddd;'>$(Html-Encode $row.$_)</td>" }) -join ""
        "<tr>$tds</tr>"
    }) -join "`n"

    return @"
<h3 style="margin:16px 0 8px 0;">$(Html-Encode $Title)</h3>
<table style="border-collapse:collapse;font-family:Segoe UI,Arial;font-size:12px;">
<tr>$th</tr>
$trs
</table>
"@
}

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
            if ($s -match "fail|error|block")                          { return "High" }
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

    if ($UpgradeEligibilityLabel -and $UpgradeEligibilityLabel -ne "Eligible" -and
        $BlockingReason -in @("DeploymentInProgress","UpdateAlert")) {
        return @{ Priority=2; ActionCode="FIX_READINESS_FIRST"; ActionDescription="Device not eligible for Windows 11. Fix readiness before continuing update rollout."; Owner="Workplace" }
    }

    switch ($BlockingReason) {
        "HardFailure_DiskSpace"      { return @{ Priority=1; ActionCode="FREE_DISK_SPACE";      ActionDescription="Free disk space (cleanup, remove large files/apps) and retry update."; Owner="Workplace" } }
        "HardFailure_Generic"        { return @{ Priority=1; ActionCode="CHECK_ERROR";           ActionDescription="Investigate update failure (logs, compatibility, disk, servicing stack) and remediate."; Owner="Workplace" } }
        "HardFailure_Compatibility"  { return @{ Priority=2; ActionCode="COMPATIBILITY_REVIEW";  ActionDescription="Compatibility issue suspected (driver/firmware/app). Review and remediate before retry."; Owner="Workplace" } }
        "HardFailure_SafeguardHold"  { return @{ Priority=2; ActionCode="SAFEGUARD_HOLD";        ActionDescription="Safeguard hold detected. Validate known issue and consider mitigation or wait for Microsoft resolution."; Owner="Workplace" } }
        "UpdateAlert"                { return @{ Priority=3; ActionCode="HANDLE_ALERT";          ActionDescription="Review Intune alert details and apply recommended fix."; Owner="Workplace" } }
        "DeploymentInProgress"       { return @{ Priority=4; ActionCode="WAIT_MONITOR";          ActionDescription="Deployment in progress. Monitor status, verify device online and update service health."; Owner="N1" } }
        "NotApplicableOrUnknown"     { return @{ Priority=5; ActionCode="VERIFY_SCOPE";          ActionDescription="Verify policy scope/eligibility and device prerequisites."; Owner="N2" } }
        default                      { return @{ Priority=9; ActionCode="REVIEW";                ActionDescription="Manual review required."; Owner="Workplace" } }
    }
}

# ==========================================================
# Main
# ==========================================================
$ErrorActionPreference  = "Stop"
$script:TokenAcquiredAt = [datetime]::MinValue
$scriptStart = Get-Date

try {
    Assert-PS7

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
        Write-Host ""
        Write-Host ("=" * 78) -ForegroundColor DarkGray
        Write-Host ("{0}  v{1}  {2}" -f $ScriptName, $ScriptVersion, (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) -ForegroundColor White
        Write-Host ("Host: {0}   RunId: {1}" -f $env:COMPUTERNAME, $RunId) -ForegroundColor DarkGray
        if ($DryRun) { Write-Host "*** DRY-RUN MODE: no files written, no emails sent ***" -ForegroundColor Cyan }
        Write-Host ("=" * 78) -ForegroundColor DarkGray
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

        Refresh-GraphTokenIfNeeded -TenantId $TenantId -ClientId $AppId -Certificate $cert -Headers ([ref]$headers)

        Write-Log "Processing policy [$index/$($policies.Count)]: $policyName ($policyId)" "INFO" "GRAPH"

        $filter = "(PolicyId eq '$policyId')"
        $jobId  = $null

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
        $job = Wait-IntuneExportJobComplete -JobId $jobId -Headers $headers -TimeoutSeconds 900 -PollSeconds 5

        $zipPath = Join-Path $WorkPath ("{0}_{1}_{2}.zip" -f $ReportName, $RunStamp, $policyId)
        Download-ExportZip -DownloadUrl $job.url -DestinationZipPath $zipPath -Headers $headers
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
    # Computed columns pass:
    #   RiskBucket, BlockingReason, DaysSinceLastStatus,
    #   ActionPriority, ActionCode, ActionDescription, ActionOwner
    # ----------------------------------------------------------
    Write-Log "Computing derived columns (RiskBucket, BlockingReason, DaysSinceLastStatus, Action*)..." "INFO" "ENRICH"

    $computedRows = New-Object System.Collections.Generic.List[object]
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
        if (-not [string]::IsNullOrWhiteSpace($LastStatusDateColumn)) {
            $lastStatusRaw = "$($row.$LastStatusDateColumn)"
            if (-not [string]::IsNullOrWhiteSpace($lastStatusRaw)) {
                $parsedDate = [datetime]::MinValue
                if ([datetime]::TryParse($lastStatusRaw, [ref]$parsedDate) -and $parsedDate -ne [datetime]::MinValue) {
                    $daysSince = [int]((Get-Date) - $parsedDate).TotalDays
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
    $enrichedRows = $computedRows

    $count = $enrichedRows.Count
    Write-Log "Total consolidated rows (enriched + computed): $count" "INFO" "DATA"

    # ----------------------------------------------------------
    # CSV export - DATA-ALL (atomic), DATA-LAST (atomic), Archive
    # ----------------------------------------------------------
    $spUploadStatus = "Disabled"
    $spUploadError  = ""
    $spDestPath     = ""

    if ($DryRun) {
        Write-Log "DryRun: skipping write of $CsvFinal." "INFO" "DRYRUN"
    }
    else {
        Publish-CoreSmartM365Csv -Data $enrichedRows -TimestampedPath $CsvFinal -LatestPath $CsvLastFinal -NoSharePointUpload | Out-Null
        Write-Log "CSV exported: $CsvFinal" "INFO" "DATA"
        Write-Log "DATA-LAST updated: $CsvLastFinal" "INFO" "DATA"

        $archiveBase = [System.IO.Path]::GetFileNameWithoutExtension($CsvName)
        Write-ArchiveCopyAndPrune -SourceCsv $CsvFinal -ArchiveFolder $ArchivePath -BaseNameWithoutExt $archiveBase -RunStamp $RunStamp -Keep $global:RetentionMaxCSV

        # ----------------------------------------------------------
        # SharePoint upload (non-blocking, latest CSV)
        # ----------------------------------------------------------
        if ($EnableSharePointUpload) {
            try {
                Write-Log "SharePoint upload starting: $SP_SiteHostname$SP_SitePath / $SP_LibraryDisplayName / $SP_TargetFolderPath" "INFO" "SP"

                $thumbprint = if ($global:Thumbprint) { $global:Thumbprint } else { $Thumb }
                Invoke-CoreSmartM365SharePointCsvUpload `
                    -LocalFilePath $CsvLastFinal `
                    -Enabled $EnableSharePointUpload `
                    -SiteHostname $SP_SiteHostname `
                    -SitePath $SP_SitePath `
                    -LibraryDisplayName $SP_LibraryDisplayName `
                    -TargetFolderPath $SP_TargetFolderPath `
                    -AppId $AppId `
                    -TenantId $TenantId `
                    -Thumbprint $thumbprint

                $spUploadStatus = "Requested"
                $spDestPath     = (($SP_TargetFolderPath.TrimEnd('/','\')) + "/" + [System.IO.Path]::GetFileName($CsvLastFinal)).TrimStart('/','\')
                Write-Log "SharePoint upload requested: $spDestPath" "INFO" "SP"
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
    # KPI calculations (all from $enrichedRows)
    # ----------------------------------------------------------
    $deviceKeyCol = "DeviceId"

    $totalUniqueDevices = 0
    if ($count -gt 0) {
        $totalUniqueDevices = ($enrichedRows | Select-Object -ExpandProperty $deviceKeyCol -ErrorAction SilentlyContinue |
                               Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique).Count
    }

    $globalLowCount      = 0
    $globalHighMedCount  = 0
    $globalCompletionPct = 0.0
    $globalAtRiskPct     = 0.0

    if ($count -gt 0) {
        $globalLowCount      = ($enrichedRows | Where-Object { $_.RiskBucket -eq "Low" }).Count
        $globalHighMedCount  = ($enrichedRows | Where-Object { $_.RiskBucket -in @("High","Medium") }).Count
        $globalCompletionPct = [math]::Round((100.0 * $globalLowCount     / $count), 2)
        $globalAtRiskPct     = [math]::Round((100.0 * $globalHighMedCount / $count), 2)
    }

    $blockingCount    = ($enrichedRows | Where-Object { $_.BlockingReason -ne "CompliantOrCompleted" }).Count
    $actionableCount  = ($enrichedRows | Where-Object { $_.BlockingReason -like "HardFailure*" -or $_.BlockingReason -eq "UpdateAlert" }).Count
    $monitoringCount  = ($enrichedRows | Where-Object { $_.BlockingReason -eq "DeploymentInProgress" }).Count

    $eligibleCompleted     = ($enrichedRows | Where-Object { $_.UpgradeEligibilityLabel -eq "Eligible" -and $_.RiskBucket -eq "Low" }).Count
    $eligibleAtRisk        = ($enrichedRows | Where-Object { $_.UpgradeEligibilityLabel -eq "Eligible" -and $_.RiskBucket -in @("High","Medium") }).Count
    $notEligibleInProgress = ($enrichedRows | Where-Object { $_.UpgradeEligibilityLabel -and $_.UpgradeEligibilityLabel -ne "Eligible" -and $_.RiskBucket -eq "Medium" }).Count
    $notMatchedBlocking    = ($enrichedRows | Where-Object { $_.ReadinessMatch -eq "NotMatched" -and $_.BlockingReason -ne "CompliantOrCompleted" }).Count
    $eligibleBlockedHard   = ($enrichedRows | Where-Object { $_.UpgradeEligibilityLabel -eq "Eligible" -and $_.BlockingReason -like "HardFailure*" }).Count

    $eligibleHardFailureStatus = if ($eligibleBlockedHard -eq 0) { "Good" } else { "Critical" }

    Write-Log "KPIs: Rows=$count UniqueDevices=$totalUniqueDevices Completion=$globalCompletionPct% AtRisk=$globalAtRiskPct% Blocking=$blockingCount Actionable=$actionableCount Monitoring=$monitoringCount EligibleHardFail=$eligibleBlockedHard" "INFO" "KPI"

    # ----------------------------------------------------------
    # Success email
    # ----------------------------------------------------------
    $should = Should-SendSummaryEmail -CurrentCount $count -Mode $SummaryEmailMode -StatePath $SummaryStatePath
    if ($should) {
        $duration = New-TimeSpan -Start $scriptStart -End (Get-Date)

        # KPI table
        $kpiRows = @(
            [pscustomobject]@{ Category="Total Rows";                                                           Count=$count }
            [pscustomobject]@{ Category="Global Unique Devices";                                                Count=$totalUniqueDevices }
            [pscustomobject]@{ Category="Global Completion Rate (%)";                                           Count=$globalCompletionPct }
            [pscustomobject]@{ Category="Global At-Risk Rate (%)";                                              Count=$globalAtRiskPct }
            [pscustomobject]@{ Category="Blocking (All)";                                                       Count=$blockingCount }
            [pscustomobject]@{ Category="Blocking (Actionable)";                                                Count=$actionableCount }
            [pscustomobject]@{ Category="Blocking (Monitoring)";                                                Count=$monitoringCount }
            [pscustomobject]@{ Category="Eligible + Completed/LowRisk";                                         Count=$eligibleCompleted }
            [pscustomobject]@{ Category="Eligible + AtRisk (High/Medium)";                                      Count=$eligibleAtRisk }
            [pscustomobject]@{ Category="NotEligible + InProgress/Medium";                                      Count=$notEligibleInProgress }
            [pscustomobject]@{ Category="Readiness NotMatched + Blocking";                                      Count=$notMatchedBlocking }
            [pscustomobject]@{ Category=("Eligible + HardFailure (Critical) - " + $eligibleHardFailureStatus); Count=$eligibleBlockedHard }
        )
        $kpiTable = Convert-ObjectsToHtmlTable -Rows $kpiRows -Columns @("Category","Count") `
            -Title "Windows 11 Rollout Readiness vs Update Reality"

        # Policy summary (computed inline for email)
        $policyGroups  = $enrichedRows | Group-Object -Property PolicyId
        $policySummary = foreach ($g in $policyGroups) {
            $pPolicyId = $g.Name
            $pName     = ($g.Group | Select-Object -First 1).PolicyName
            $uniqueDevices = ($g.Group | Select-Object -ExpandProperty $deviceKeyCol -ErrorAction SilentlyContinue |
                             Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique).Count
            $coveragePct = 0.0
            if ($totalUniqueDevices -gt 0) { $coveragePct = [math]::Round((100.0 * $uniqueDevices / $totalUniqueDevices), 2) }
            $high = ($g.Group | Where-Object { $_.RiskBucket -eq "High" }).Count
            $med  = ($g.Group | Where-Object { $_.RiskBucket -eq "Medium" }).Count
            $low  = ($g.Group | Where-Object { $_.RiskBucket -eq "Low" }).Count
            $atRisk    = $high + $med
            $atRiskPct = 0.0
            if ($g.Count -gt 0) { $atRiskPct = [math]::Round((100.0 * $atRisk / $g.Count), 2) }
            [pscustomobject]@{
                PolicyName=    $pName
                PolicyId=      $pPolicyId
                Rows=          $g.Count
                UniqueDevices= $uniqueDevices
                CoveragePct=   $coveragePct
                AtRiskRows=    $atRisk
                AtRiskPct=     $atRiskPct
                HighRisk=      $high
                MediumRisk=    $med
                LowRows=       $low
            }
        }
        $policySummarySorted = $policySummary | Sort-Object -Property Rows -Descending

        $policySummaryTable = Convert-ObjectsToHtmlTable `
            -Rows ($policySummarySorted | Select-Object PolicyName,PolicyId,Rows,UniqueDevices,CoveragePct,AtRiskRows,AtRiskPct,HighRisk,MediumRisk,LowRows) `
            -Columns @("PolicyName","PolicyId","Rows","UniqueDevices","CoveragePct","AtRiskRows","AtRiskPct","HighRisk","MediumRisk","LowRows") `
            -Title "Summary by Policy (Coverage + Risk)"

        $topRiskPolicies = $policySummary |
            Sort-Object @{Expression="AtRiskRows";Descending=$true},@{Expression="AtRiskPct";Descending=$true} |
            Select-Object -First $RiskTopN

        $topRiskTable = Convert-ObjectsToHtmlTable `
            -Rows ($topRiskPolicies | Select-Object PolicyName,PolicyId,AtRiskRows,AtRiskPct,HighRisk,MediumRisk,CoveragePct) `
            -Columns @("PolicyName","PolicyId","AtRiskRows","AtRiskPct","HighRisk","MediumRisk","CoveragePct") `
            -Title ("Top {0} Policies at Risk" -f $RiskTopN)

        # Top eligible blocked devices
        $topEligibleBlockedTable = Convert-ObjectsToHtmlTable `
            -Rows @($enrichedRows | Where-Object { $_.UpgradeEligibilityLabel -eq "Eligible" -and $_.ActionCode -eq "ELIGIBLE_DEVICE_BLOCKED" } |
                Select-Object -First 20 ActionPriority,PolicyName,DeviceName,AggregateState,CurrentDeviceUpdateStatus,LatestAlertMessage,ActionCode,ActionOwner) `
            -Columns @("ActionPriority","PolicyName","DeviceName","AggregateState","CurrentDeviceUpdateStatus","LatestAlertMessage","ActionCode","ActionOwner") `
            -Title "Top Eligible Devices Blocked (Critical)"

        # Top actionable items
        $topActionsTable = Convert-ObjectsToHtmlTable `
            -Rows @($enrichedRows | Where-Object { $_.BlockingReason -like "HardFailure*" -or $_.BlockingReason -eq "UpdateAlert" } |
                Sort-Object ActionPriority,PolicyName,DeviceName | Select-Object -First 20 `
                ActionPriority,ActionOwner,ActionCode,PolicyName,DeviceName,AggregateState,BlockingReason,UpgradeEligibilityLabel) `
            -Columns @("ActionPriority","ActionOwner","ActionCode","PolicyName","DeviceName","AggregateState","BlockingReason","UpgradeEligibilityLabel") `
            -Title "Top Actionable Items"

        $subject = "[SmartM365] WinUpdate Feature Update status - SUCCESS - Rows=$count Completion=$globalCompletionPct% AtRisk=$globalAtRiskPct% Actionable=$actionableCount Monitoring=$monitoringCount EligibleHardFail=$eligibleBlockedHard"

        $spStatusColor = if ($spUploadStatus -eq "Error") { "color:#b00020;" } elseif ($spUploadStatus -eq "Success") { "color:#007700;" } else { "" }

        $html = @"
<html><body style="font-family:Segoe UI,Arial;">
<h2 style="margin:0 0 10px 0;">Windows Update Feature Update status - SUCCESS</h2>
<ul>
<li><b>Script</b>: $(Html-Encode $ScriptName) v$(Html-Encode $ScriptVersion)</li>
<li><b>RunId</b>: $(Html-Encode $RunId)</li>
<li><b>Host</b>: $(Html-Encode $env:COMPUTERNAME)</li>
<li><b>ReportName</b>: $(Html-Encode $ReportName)</li>
<li><b>Policies</b>: $($policies.Count)</li>
<li><b>Rows</b>: $count</li>
<li><b>Global unique devices</b>: $totalUniqueDevices</li>
<li><b>Global Completion Rate</b>: $globalCompletionPct% ($globalLowCount Low-risk rows)</li>
<li><b>Global At-Risk Rate</b>: $globalAtRiskPct% ($globalHighMedCount High/Medium rows)</li>
<li><b>Blocking (All / Actionable / Monitoring)</b>: $blockingCount / $actionableCount / $monitoringCount</li>
<li><b>Eligible + HardFailure</b>: $eligibleBlockedHard</li>
<li><b>Execution time</b>: $([int]$duration.TotalSeconds) seconds</li>
<li><b>ExportDateTime</b>: $(Html-Encode $exportTime)</li>
$(if ($DryRun) { "<li><b style='color:#b00020;'>DRY-RUN MODE: no files written</b></li>" })
</ul>

<h3 style="margin:16px 0 8px 0;">Readiness Enrichment</h3>
<ul>
<li><b>Enabled</b>: $(Html-Encode $EnableReadinessEnrichment)</li>
<li><b>Mode</b>: $(Html-Encode $readinessMode)</li>
<li><b>Readiness rows</b>: $readinessTotal  |  <b>Matched</b>: $readinessMatched  |  <b>NotMatched</b>: $readinessNotMatched  |  <b>MatchPct</b>: $readinessMatchPct%</li>
<li><b>Eligibility distribution</b>: $(Html-Encode $readinessEligibilityDist)</li>
<li><b>Summary</b>: $(Html-Encode $readinessSummaryText)</li>
</ul>

<h3 style="margin:16px 0 8px 0;">Output CSV</h3>
<ul>
<li><b>DATA-ALL</b>: $(Html-Encode $CsvFinal)</li>
<li><b>DATA-LAST</b>: $(Html-Encode $CsvLastFinal)</li>
</ul>

<h3 style="margin:16px 0 8px 0;">SharePoint Upload</h3>
<ul>
<li><b>Status</b>: <span style="$spStatusColor">$(Html-Encode $spUploadStatus)</span></li>
<li><b>Target</b>: $(Html-Encode "$SP_SiteHostname$SP_SitePath / $SP_LibraryDisplayName / $SP_TargetFolderPath")</li>
$(if ($spDestPath)     { "<li><b>Destination</b>: $(Html-Encode $spDestPath)</li>" })
$(if ($spUploadError)  { "<li><b style='color:#b00020;'>Error</b>: $(Html-Encode $spUploadError)</li>" })
</ul>

$kpiTable
$topEligibleBlockedTable
$topActionsTable
$policySummaryTable
$topRiskTable

</body></html>
"@

        Send-SummaryEmail -Subject $subject -HtmlBody $html
        Save-SummaryState -Count $count -StatePath $SummaryStatePath
    }
    else {
        Write-Log "Summary email skipped (Mode=$SummaryEmailMode)." "INFO" "MAIL"
    }

    Write-Log "Completed successfully. Version=$ScriptVersion DryRun=$DryRun Rows=$count Completion=$globalCompletionPct% AtRisk=$globalAtRiskPct%" "INFO"
}
catch {
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
    Cleanup-TempFiles
}


# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD2CM/TIMIZsoD6
# TcPxpCh7uUb9+w0itb9h3TGUe0VHnKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCCovBxzjfc6yeMCDZrIUn7gEhhOrmsHuToBCMpYg/9xhDANBgkqhkiG9w0B
# AQEFAASCAYBQBgMSgaW6+i47v69lN2WxYhnDgAQx5IxlcDabOkSbPFaqMUJ5VoE+
# 72wMOEA3yitD0AhqhyVF2VPKi+Rq4/NxixE13sk66DfYlqa8hYbKnOkrjRPUMZb4
# P0xE14VozAdEk8FLh2cnqGLiFy54AfHsj2dNQ9pmFCsu+IcawSvrG2QcwldR/wr2
# yf2i+odGPw+OlNVQT302Hi+SBMIIeQtCJytlrFEQEEKkn8fguOyUA2SLgSMoauDu
# 1w5dOO4iS8Bk5Y4BV3eD1wavGN+i++lAQslTRxI1yrh2dZ/aCmXqDf6CMjWy4spp
# LPaqWWbK4txIhfuAv8bczDJ3BC9MRqF5UBrxMfW4aRww1ZZ9vGYns+cu82oHwnh/
# wt0ZdoSzPzlOje91EM8ocnGf9jV1uP56//HIBWzI0C/IYIo9tZfqr+5OXn/y5jor
# KZEK5SK2uXTtRB7bGjFAkCFhugwO3T7PLvo9i57nzEHOfqdiYQltX9Usl2PhmkQU
# LpvIsyVr5+Y=
# SIG # End signature block
