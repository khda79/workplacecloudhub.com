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
  1.19
.VERSION
1.20

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
$ScriptVersion = "1.20"

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

    if (-not $EnableSummaryEmail) { return $false }
    if ($DryRun) { Write-Log "DryRun: skipping summary email." "INFO" "DRYRUN"; return $false }
    $summaryTo = $To
    if ([string]::IsNullOrWhiteSpace($summaryTo)) {
        Write-Log "Summary email skipped because To is not configured in local or global configuration. Inventory outputs remain valid." "WARN" "MAIL"
        return $false
    }

    try {
        Import-SmartM365CorePreflight
        $mailParams = @{ To = $summaryTo; Subject = $Subject; BodyHtml = $HtmlBody }
        if (-not [string]::IsNullOrWhiteSpace($From)) { $mailParams['From'] = $From }
        if (-not [string]::IsNullOrWhiteSpace($Cc)) { $mailParams['Cc'] = $Cc }
        Send-CoreSmartM365Mail @mailParams
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

    if ($UpgradeEligibilityLabel -in @("NotEligible","Unknown","NotApplicableOrUnknown") -and
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

function Get-WinUpdateDeviceSummaryRows {
    param([Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Rows)

    $groups = @{}
    $rowIndex = 0
    foreach ($row in $Rows) {
        $rowIndex++
        $deviceId = [string](Get-WinUpdateRowValue -Row $row -Name 'DeviceId')
        $deviceName = [string](Get-WinUpdateRowValue -Row $row -Name 'DeviceName')
        $normalizedName = Normalize-DeviceName -DeviceName $deviceName
        $key = if (-not [string]::IsNullOrWhiteSpace($deviceId)) { "id:$($deviceId.ToLowerInvariant())" } elseif ($normalizedName) { "name:$normalizedName" } else { "row:$rowIndex" }
        if (-not $groups.ContainsKey($key)) { $groups[$key] = New-Object System.Collections.Generic.List[object] }
        $groups[$key].Add($row) | Out-Null
    }

    foreach ($key in $groups.Keys) {
        $selected = @($groups[$key] | Sort-Object `
            @{Expression={ Get-WinUpdateOperationalRank -State (Get-WinUpdateOperationalState -Row $_) };Ascending=$true}, `
            @{Expression={ Get-WinUpdatePriorityRank -Value (Get-WinUpdateRowValue -Row $_ -Name 'ActionPriority') };Ascending=$true} | Select-Object -First 1)[0]
        [pscustomobject][ordered]@{
            DeviceKey = $key
            DeviceId = [string](Get-WinUpdateRowValue -Row $selected -Name 'DeviceId')
            DeviceName = [string](Get-WinUpdateRowValue -Row $selected -Name 'DeviceName')
            PolicyId = [string](Get-WinUpdateRowValue -Row $selected -Name 'PolicyId')
            PolicyName = [string](Get-WinUpdateRowValue -Row $selected -Name 'PolicyName')
            OperationalState = Get-WinUpdateOperationalState -Row $selected
            AggregateState = [string](Get-WinUpdateRowValue -Row $selected -Name 'AggregateState')
            CurrentDeviceUpdateStatus = [string](Get-WinUpdateRowValue -Row $selected -Name 'CurrentDeviceUpdateStatus')
            LatestAlertMessage = [string](Get-WinUpdateRowValue -Row $selected -Name 'LatestAlertMessage')
            BlockingReason = [string](Get-WinUpdateRowValue -Row $selected -Name 'BlockingReason')
            UpgradeEligibilityLabel = [string](Get-WinUpdateRowValue -Row $selected -Name 'UpgradeEligibilityLabel')
            ReadinessMatch = [string](Get-WinUpdateRowValue -Row $selected -Name 'ReadinessMatch')
            ActionPriority = Get-WinUpdatePriorityRank -Value (Get-WinUpdateRowValue -Row $selected -Name 'ActionPriority')
            ActionCode = [string](Get-WinUpdateRowValue -Row $selected -Name 'ActionCode')
            ActionDescription = [string](Get-WinUpdateRowValue -Row $selected -Name 'ActionDescription')
            ActionOwner = [string](Get-WinUpdateRowValue -Row $selected -Name 'ActionOwner')
        }
    }
}
# ==========================================================
# Main
# ==========================================================
$ErrorActionPreference  = "Stop"
$script:TokenAcquiredAt = [datetime]::MinValue
$scriptStart = Get-Date

$script:CompletionStatus = 'Auto'
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
    $spUploadUrl    = ""

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
    $unknownCount = @($deviceSummaryRows | Where-Object OperationalState -eq 'Unknown').Count
    $completionPct = if ($totalUniqueDevices -gt 0) { [math]::Round((100.0 * $completedCount / $totalUniqueDevices),2) } else { 0.0 }
    $reportStatus = if ($actionRequiredCount -gt 0) { 'CRITICAL' } elseif ($inProgressCount -gt 0 -or $unknownCount -gt 0) { 'WARNING' } else { 'OK' }
    $summaryState = "$totalUniqueDevices|$completedCount|$inProgressCount|$actionRequiredCount|$unknownCount"

    Write-Log "Operational summary: Devices=$totalUniqueDevices Completed=$completedCount InProgress=$inProgressCount ActionRequired=$actionRequiredCount Unknown=$unknownCount Completion=$completionPct% Status=$reportStatus" "INFO" "KPI"

    $should = Should-SendSummaryEmail -CurrentState $summaryState -Mode $SummaryEmailMode -StatePath $SummaryStatePath
    if ($should) {
        $duration = New-TimeSpan -Start $scriptStart -End (Get-Date)
        $policySummary = foreach ($group in ($enrichedRows | Group-Object -Property PolicyId)) {
            $policyDevices = @(Get-WinUpdateDeviceSummaryRows -Rows $group.Group)
            $policyTotal = $policyDevices.Count
            $policyCompleted = @($policyDevices | Where-Object OperationalState -eq 'Completed').Count
            $policyInProgress = @($policyDevices | Where-Object OperationalState -eq 'InProgress').Count
            $policyAction = @($policyDevices | Where-Object OperationalState -eq 'ActionRequired').Count
            $policyUnknown = @($policyDevices | Where-Object OperationalState -eq 'Unknown').Count
            $policyStatus = if ($policyAction -gt 0) { 'CRITICAL' } elseif ($policyInProgress -gt 0 -or $policyUnknown -gt 0) { 'WARNING' } else { 'OK' }
            [pscustomobject][ordered]@{
                PolicyId = [string]$group.Name
                PolicyName = [string]@($group.Group)[0].PolicyName
                Status = $policyStatus
                Devices = $policyTotal
                Completed = $policyCompleted
                InProgress = $policyInProgress
                ActionRequired = $policyAction
                Unknown = $policyUnknown
                CompletionPct = if ($policyTotal) { [math]::Round(100.0*$policyCompleted/$policyTotal,2) } else { 0 }
            }
        }
        $policySummary = @($policySummary | Sort-Object @{Expression='ActionRequired';Descending=$true},@{Expression='InProgress';Descending=$true},PolicyName)
        $policyStrips = foreach ($policy in $policySummary) {
            $policyStatusColor = switch ($policy.Status) { 'CRITICAL' {'#b91c1c'} 'WARNING' {'#b45309'} default {'#15803d'} }
            $policyStatusBackground = switch ($policy.Status) { 'CRITICAL' {'#fee2e2'} 'WARNING' {'#fef3c7'} default {'#dcfce7'} }
            $encodedPolicyName = Html-Encode $policy.PolicyName
            $encodedPolicyId = Html-Encode $policy.PolicyId
            @"
<table role="presentation" style="width:100%;border-collapse:collapse;margin:0 0 14px 0;border:1px solid #dbe4ee;background:#ffffff;">
<tr><td colspan="6" style="padding:12px 14px;border-bottom:1px solid #dbe4ee;background:#f8fafc;">
<span style="display:inline-block;margin-right:8px;padding:4px 9px;border-radius:999px;background:$policyStatusBackground;color:$policyStatusColor;font-size:11px;font-weight:700;">$($policy.Status)</span>
<strong style="font-size:14px;color:#0f172a;">$encodedPolicyName</strong>
<div style="margin-top:4px;font-size:11px;color:#64748b;">Policy ID: $encodedPolicyId</div>
</td></tr>
<tr>
<td style="width:16.66%;padding:11px 9px;border-right:1px solid #e2e8f0;"><div style="font-size:10px;color:#64748b;">DEVICES</div><div style="font-size:20px;font-weight:700;color:#334155;">$($policy.Devices)</div></td>
<td style="width:16.66%;padding:11px 9px;border-right:1px solid #e2e8f0;background:#f0fdf4;"><div style="font-size:10px;color:#166534;">COMPLETED</div><div style="font-size:20px;font-weight:700;color:#334155;">$($policy.Completed)</div></td>
<td style="width:16.66%;padding:11px 9px;border-right:1px solid #e2e8f0;background:#eff6ff;"><div style="font-size:10px;color:#1d4ed8;">IN PROGRESS</div><div style="font-size:20px;font-weight:700;color:#334155;">$($policy.InProgress)</div></td>
<td style="width:16.66%;padding:11px 9px;border-right:1px solid #e2e8f0;background:#fef2f2;"><div style="font-size:10px;color:#b91c1c;">ACTION REQUIRED</div><div style="font-size:20px;font-weight:700;color:#334155;">$($policy.ActionRequired)</div></td>
<td style="width:16.66%;padding:11px 9px;border-right:1px solid #e2e8f0;background:#fffbeb;"><div style="font-size:10px;color:#92400e;">UNKNOWN</div><div style="font-size:20px;font-weight:700;color:#334155;">$($policy.Unknown)</div></td>
<td style="width:16.66%;padding:11px 9px;"><div style="font-size:10px;color:#64748b;">COMPLETION</div><div style="font-size:20px;font-weight:700;color:#334155;">$($policy.CompletionPct)%</div></td>
</tr></table>
"@
        }
        $policyStripsHtml = $policyStrips -join "`n"

        $actionLimit = [math]::Max(1,$RiskTopN)
        $actionRows = @($deviceSummaryRows | Where-Object OperationalState -eq 'ActionRequired' | Sort-Object ActionPriority,PolicyName,DeviceName | Select-Object -First $actionLimit `
            ActionPriority,DeviceName,PolicyName,AggregateState,CurrentDeviceUpdateStatus,LatestAlertMessage,UpgradeEligibilityLabel,ActionCode,ActionOwner,ActionDescription)
        $actionTable = Convert-ObjectsToHtmlTable -Rows $actionRows -Columns @('ActionPriority','DeviceName','PolicyName','AggregateState','CurrentDeviceUpdateStatus','LatestAlertMessage','UpgradeEligibilityLabel','ActionCode','ActionOwner','ActionDescription') -Title ("Top {0} devices requiring action" -f $actionLimit)

        $summaryMatched = @($deviceSummaryRows | Where-Object ReadinessMatch -eq 'Matched').Count
        $summaryNotMatched = @($deviceSummaryRows | Where-Object ReadinessMatch -eq 'NotMatched').Count
        $summaryReadinessPct = if (($summaryMatched+$summaryNotMatched) -gt 0) { [math]::Round(100.0*$summaryMatched/($summaryMatched+$summaryNotMatched),2) } else { 0 }
        $readinessRows = @(
            [pscustomobject]@{ Metric='Enrichment enabled'; Value=$EnableReadinessEnrichment }
            [pscustomobject]@{ Metric='Mode'; Value=$readinessMode }
            [pscustomobject]@{ Metric='Matched devices'; Value=$summaryMatched }
            [pscustomobject]@{ Metric='Unmatched devices'; Value=$summaryNotMatched }
            [pscustomobject]@{ Metric='Match percentage'; Value="$summaryReadinessPct%" }
        )
        $readinessTable = Convert-ObjectsToHtmlTable -Rows $readinessRows -Columns @('Metric','Value') -Title 'Readiness data quality'

        $statusColor = switch ($reportStatus) { 'CRITICAL' {'#b91c1c'} 'WARNING' {'#b45309'} default {'#15803d'} }
        $statusBackground = switch ($reportStatus) { 'CRITICAL' {'#fee2e2'} 'WARNING' {'#fef3c7'} default {'#dcfce7'} }
        $subject = "[$reportStatus] WinUpdate Feature Update - All policies - $completedCount/$totalUniqueDevices unique devices completed - $actionRequiredCount action(s)"
        $fileLinkHtml = if (-not [string]::IsNullOrWhiteSpace($spUploadUrl)) {
            $encodedUrl = Html-Encode $spUploadUrl
            $encodedName = Html-Encode ([IO.Path]::GetFileName($CsvLastFinal))
            "<div style='margin-top:18px;padding:14px;border:1px solid #dbe4ee;background:#f8fafc;'><h3 style='margin:0 0 8px;'>Export</h3><a href='$encodedUrl' style='color:#075985;text-decoration:underline;'>$encodedName</a></div>"
        } else { '' }

        $html = @"
<div style="font-family:Segoe UI,Arial;color:#1f2937;">
<h2 style="margin:0 0 6px 0;font-size:20px;color:#0f172a;">Global - all Feature Update policies</h2>
<p style="margin:0 0 14px 0;font-size:12px;color:#64748b;">Tenant: $(Html-Encode $OrgDomain) | Policies: $($policies.Count) | Source rows: $count | Duration: $([int]$duration.TotalSeconds) seconds | RunId: $(Html-Encode $RunId)</p>
<div style="display:inline-block;padding:6px 12px;border-radius:999px;background:$statusBackground;color:$statusColor;font-weight:700;">$reportStatus</div>
<p style="margin:14px 0 18px;">Global scope: each device is counted once across all policies collected during this run. When a device appears in several policies, its most severe state is retained.</p>
<table role="presentation" style="width:100%;border-collapse:separate;border-spacing:8px;"><tr>
<td style="padding:14px;background:#f8fafc;border:1px solid #e2e8f0;"><div style="font-size:11px;color:#64748b;">UNIQUE DEVICES</div><div style="font-size:24px;font-weight:700;">$totalUniqueDevices</div></td>
<td style="padding:14px;background:#dcfce7;border:1px solid #bbf7d0;"><div style="font-size:11px;color:#166534;">COMPLETED</div><div style="font-size:24px;font-weight:700;">$completedCount</div></td>
<td style="padding:14px;background:#dbeafe;border:1px solid #bfdbfe;"><div style="font-size:11px;color:#1d4ed8;">IN PROGRESS</div><div style="font-size:24px;font-weight:700;">$inProgressCount</div></td>
<td style="padding:14px;background:#fee2e2;border:1px solid #fecaca;"><div style="font-size:11px;color:#b91c1c;">ACTION REQUIRED</div><div style="font-size:24px;font-weight:700;">$actionRequiredCount</div></td>
<td style="padding:14px;background:#fef3c7;border:1px solid #fde68a;"><div style="font-size:11px;color:#92400e;">UNKNOWN</div><div style="font-size:24px;font-weight:700;">$unknownCount</div></td>
<td style="padding:14px;background:#f8fafc;border:1px solid #e2e8f0;"><div style="font-size:11px;color:#64748b;">COMPLETION</div><div style="font-size:24px;font-weight:700;">$completionPct%</div></td>
</tr></table>
<h2 style="margin:24px 0 6px 0;font-size:18px;color:#0f172a;">Feature Update policies</h2>
<p style="margin:0 0 14px 0;font-size:12px;color:#64748b;">Each device is counted once within each policy. A device targeted by several policies appears in every relevant policy strip, so policy totals are not additive and may exceed the global unique-device total.</p>
$policyStripsHtml
$actionTable
$readinessTable
$fileLinkHtml
</div>
"@

        if (Send-SummaryEmail -Subject $subject -HtmlBody $html) {
            Save-SummaryState -State $summaryState -StatePath $SummaryStatePath
        }
    }
    else {
        Write-Log "Summary email skipped (Mode=$SummaryEmailMode)." "INFO" "MAIL"
    }

    Write-Log "Completed successfully. Version=$ScriptVersion DryRun=$DryRun Rows=$count Devices=$totalUniqueDevices Completed=$completedCount InProgress=$inProgressCount ActionRequired=$actionRequiredCount Unknown=$unknownCount Completion=$completionPct%" "INFO"
}
catch {
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
    Cleanup-TempFiles
    Write-SmartM365CompletionBanner -Status $script:CompletionStatus -ScriptName $ScriptName -StartedAt $scriptStart
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDuX1azaP+V0az4
# orlSV1PXtZ9glCm9cpeKp4IqoWdzKqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAKgO8YS43xBYLRxHan
# lXRoMA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjUwNjA0MDAwMDAwWhcN
# MzYwOTAzMjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNSAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA0EasLRLGntDqrmBWsytXum9R/4ZwCgHfyjfMGUIwYzKomd8U1nH7C8Dr0cVM
# F3BsfAFI54um8+dnxk36+jx0Tb+k+87H9WPxNyFPJIDZHhAqlUPt281mHrBbZHqR
# K71Em3/hCGC5KyyneqiZ7syvFXJ9A72wzHpkBaMUNg7MOLxI6E9RaUueHTQKWXym
# OtRwJXcrcTTPPT2V1D/+cFllESviH8YjoPFvZSjKs3SKO1QNUdFd2adw44wDcKgH
# +JRJE5Qg0NP3yiSyi5MxgU6cehGHr7zou1znOM8odbkqoK+lJ25LCHBSai25CFyD
# 23DZgPfDrJJJK77epTwMP6eKA0kWa3osAe8fcpK40uhktzUd/Yk0xUvhDU6lvJuk
# x7jphx40DQt82yepyekl4i0r8OEps/FNO4ahfvAk12hE5FVs9HVVWcO5J4dVmVzi
# x4A77p3awLbr89A90/nWGjXMGn7FQhmSlIUDy9Z2hSgctaepZTd0ILIUbWuhKuAe
# NIeWrzHKYueMJtItnj2Q+aTyLLKLM0MheP/9w6CtjuuVHJOVoIJ/DtpJRE7Ce7vM
# RHoRon4CWIvuiNN1Lk9Y+xZ66lazs2kKFSTnnkrT3pXWETTJkhd76CIDBbTRofOs
# NyEhzZtCGmnQigpFHti58CSmvEyJcAlDVcKacJ+A9/z7eacCAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFOQ7/PIx7f391/ORcWMZUEPPYYzoMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAGUqrfEcJwS5rmBB
# 7NEIRJ5jQHIh+OT2Ik/bNYulCrVvhREafBYF0RkP2AGr181o2YWPoSHz9iZEN/FP
# sLSTwVQWo2H62yGBvg7ouCODwrx6ULj6hYKqdT8wv2UV+Kbz/3ImZlJ7YXwBD9R0
# oU62PtgxOao872bOySCILdBghQ/ZLcdC8cbUUO75ZSpbh1oipOhcUT8lD8QAGB9l
# ctZTTOJM3pHfKBAEcxQFoHlt2s9sXoxFizTeHihsQyfFg5fxUFEp7W42fNBVN4ue
# LaceRf9Cq9ec1v5iQMWTFQa0xNqItH3CPFTG7aEQJmmrJTV3Qhtfparz+BW60OiM
# EgV5GWoBy4RVPRwqxv7Mk0Sy4QHs7v9y69NBqycz0BZwhB9WOfOu/CIJnzkQTwtS
# SpGGhLdjnQ4eBpjtP+XB3pQCtv4E5UCSDag6+iX8MmB10nfldPF9SVD7weCC3yXZ
# i/uuhqdwkgVxuiMFzGVFwYbQsiGnoa9F5AaAyBjFBtXVLcKtapnMG3VH3EmAp/js
# J3FVF3+d1SVDTmjFjLbNFZUWMXuZyvgLfgyPehwJVxwC+UpX2MSey2ueIu9THFVk
# T+um1vshETaWyQo8gmBto/m3acaP9QsuLj3FNwFlTxq25+T4QwX9xa6ILs84ZPvm
# povq90K8eWyG2N01c4IhSOxqt81nMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEINBuN6y+DEC4/13HE7lTwEGmL1CGAXGhg7B93FliNublMA0GCSqG
# SIb3DQEBAQUABIIBgGYs+xCwOjkRldP5U5gGu5QR93VbhpmwaU8UHyjA2vtd3THM
# gvbJrgL9uIrbeqVXe4C2Ab5vQI4VDBEjuRATfhXrmF2gzlOAwU+GWcjxy9WKqrC5
# br5gjI33QzNk/urnAue7BGIA/Qh3YZoWt6JD1sJpOh14Hj0ggCRiiXofZT/YQJMa
# MjevbK3dzfF1bX6xxU6Ian309SlT2gpapRqqyq3S4snrVtmLI/JM5I2/vCfAcmiO
# dxeTHpgDKjcC6aSF6lF1vt2//nwebXNC8Cca7bKH/a1kCAT6EKfaSj72HJpsdSMn
# bDHAR+H3UtiiUEUUTdJUydbJVcY+YcPDEnEtw44WC84IRJNJsCg2XlS2TqjSTyjv
# x5/ld1QWRf139Kv+G2J2mwFJaVSpcEBAMIZhNJDlgkmXj364lPigBLPsZ0SM9eWo
# 8ZSlQRnuwX1cS4AQavZaEc5vt093JDhF8CYMh7zyX+L+FlQ/KFel3T0ji4W8oFK0
# gqHfIAnSvIdbj9WcPKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTQwMTEy
# NTdaMC8GCSqGSIb3DQEJBDEiBCA6JEyv/HFQxPfHSRNAqCPckYn1YPBucaC/TFkY
# f2DQ8zANBgkqhkiG9w0BAQEFAASCAgCDpsM0Wvnfi05EXjascC7SHOBC3jhr+Qs4
# csPQRAqf53u8SBcJiSsawLyymKk1QRLVanyTenejg1PpND1NZH6Zm3qqK7xyEmV0
# EbUWTu+kg1RgcRx4l2Iz7qvr3c90BurIMFy2OZW9xniMFZnuoPhOppeuT1Lu/j2c
# 4S9iVmhEaQXG0X9CDophnyxr84bezm67U87Si1NK3YYOXJlaBIuHdB7PniSegtcr
# ILn3lfGZEYi8HQAGd6YMTArh0oABxoZEy0DpfWm4D204VSr4ZlWUMIl3Q/4xqZPN
# FSzf0ytGlyUcEBa2q3OOSJkB8W3mc9j4hQeMuFiJxIi+qez9nzHiqDbDjDhf5IVK
# ubFP+vxTQP7oP+SPBIDDigZ6GQMBL7G/03dCBpDKCpfnJotTPg6QETHalDKIGaBA
# rw7rlRXoiKLsNIdBbm41CTuK0GnDisYs2a0iISE39dMFadGFPX6SJHgytPqjr9CF
# fj8jZ4IsUgTWdrGgarrYPAOhVVRuA7nwPi3wv02HVM0UFAZTczmT8MgfSKJefzcr
# rrcJhiXAQfeXNnF4dBdmUyydgFap1Tr5H/pPvc6+OY/M5Dna3r648o2ohZ/wMfvF
# qt0Mg91wPd/zKlYcFy/91yk2bIdYJPdyYyZMS0AyBj5ufRX+8HWOX9AmHYOze6F4
# 8B+dpWbb9w==
# SIG # End signature block
