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
- SharePoint upload: CSV copied to SharePoint Online via SmartM365.SharePoint module (same Graph token)
- Emails:
    - Fatal error HTML email (on failure)
    - Optional success/summary HTML email with KPI tables

REQUIREMENTS
- PowerShell 7+
- MSAL.PS module
- App-only certificate authentication for Graph
- SmartM365.SharePoint module (Modules\SmartM365.SharePoint\SmartM365.SharePoint.psd1)
- Graph app permission: Sites.Selected with site-level write grant (for SharePoint upload)

PARAMETERS
  -DryRun                    : Run all API calls and processing; skip file writes and emails
  -MaxPolicies               : Limit number of policies processed (0 = all; useful for testing)
  -EnableReadinessEnrichment : Toggle Win11 Readiness CSV enrichment (default: $true)
  -SummaryEmailMode          : Always | OnChange | Never (default: Always)
  -EnableSummaryEmail        : Toggle success email (default: $true)
  -EnableErrorEmail          : Toggle error email (default: $true)
  -RiskTopN                  : Number of top-risk policies shown in email (default: 10)

VERSION
  1.4.0
.NOTES
    Author: https://github.com/khda79/M365
#>

param(
    [string]$Tenant = 'test',
[switch]$DryRun,
    [int]$MaxPolicies = 0,
    [bool]$EnableReadinessEnrichment = $true,
    [ValidateSet("Always","OnChange","Never")][string]$SummaryEmailMode = "Always",
    [bool]$EnableSummaryEmail = $true,
    [bool]$EnableErrorEmail = $true,
    [int]$RiskTopN = 10
)
$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $p = Join-Path -Path $d -ChildPath 'SmartM365-TenantContext.ps1'
        if (Test-Path -LiteralPath $p) { return $p }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $tenantContextPath
Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot | Out-Null

# ==========================================================
# Version
# ==========================================================
$ScriptVersion = "1.0"

# ==========================================================
# App-only authentication parameters
# ==========================================================
function Get-ScriptLocalConfig {
    [CmdletBinding()]
    param()

    $configPath = Join-Path -Path $PSScriptRoot -ChildPath ("{0}.local.json" -f [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))
    if (-not (Test-Path -LiteralPath $configPath)) {
        return [pscustomobject]@{}
    }

    try {
        return Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
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
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'SmartM365.global.local.json'
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
            if ($localValue -and $localValue -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) {
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
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'SmartM365.global.local.json'
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
        if ($globalProperty.Value -is [string] -and [string]::IsNullOrWhiteSpace($globalProperty.Value)) {
            return $DefaultValue
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
    Join-Path $LogAllRootPath "WinUpdate_Status_From_Intune"
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
$From              = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'From' -DefaultValue ""

$ErrorMailTo   = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ErrorMailTo' -DefaultValue ""

$SummaryStatePath = Join-Path $ScriptCsvLogFolderPath "Intune_WindowsUpdate_Status.lastcount.txt"

# ==========================================================
# Console rendering options
# ==========================================================
$ConsolePretty    = $true
$ConsoleShowRunId = $false

# ==========================================================
# Run metadata
# ==========================================================
$ScriptName = "WinUpdate_Status_From_Intune"
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
            Import-Module $modulePath -Prefix Core -ErrorAction Stop
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
    $fileLine = "[$ts][$Level][$ScriptName][$RunId] $Message"
    Add-Content -Path $LogPath -Value $fileLine -Encoding UTF8

    if (-not $ConsolePretty) { Write-Host $fileLine; return }

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
        Send-CoreSmartM365Mail -From $From -To $ErrorMailTo -Subject $Subject -BodyHtml $HtmlBody
        Write-Log "Fatal error email sent via Graph." "INFO" "MAIL"
    }
    catch {
        Write-Log "Failed to send fatal error email: $($_.Exception.Message)" "WARN" "MAIL"
    }
}

function Send-SummaryEmail {
    param([Parameter(Mandatory=$true)][string]$Subject,[Parameter(Mandatory=$true)][string]$HtmlBody)

    if (-not $EnableSummaryEmail) { return }
    if ($DryRun) { Write-Log "DryRun: skipping summary email." "INFO" "DRYRUN"; return }
    if ([string]::IsNullOrWhiteSpace($ErrorMailTo)) { return }

    try {
        Import-SmartM365CorePreflight
        Send-CoreSmartM365Mail -From $From -To $ErrorMailTo -Subject $Subject -BodyHtml $HtmlBody
        Write-Log "Summary email sent via Graph." "INFO" "MAIL"
    }
    catch {
        Write-Log "Failed to send summary email: $($_.Exception.Message)" "WARN" "MAIL"
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

function Invoke-GraphGetAllPages {
    param([Parameter(Mandatory=$true)][string]$Uri,[Parameter(Mandatory=$true)][hashtable]$Headers)

    $all = New-Object System.Collections.Generic.List[object]
    $next = $Uri
    while ($next) {
        $resp = Invoke-RestMethod -Method GET -Uri $next -Headers $Headers
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

    $resp = Invoke-RestMethod -Method POST -Uri $uri -Headers $Headers -Body ($payload | ConvertTo-Json -Depth 10) -ContentType "application/json"
    if (-not $resp -or [string]::IsNullOrWhiteSpace($resp.id)) { throw "exportJobs POST returned no job id." }
    return $resp.id
}

function Get-IntuneExportJob {
    param([Parameter(Mandatory=$true)][string]$JobId,[Parameter(Mandatory=$true)][hashtable]$Headers)
    return Invoke-RestMethod -Method GET -Uri "https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs('$JobId')" -Headers $Headers
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

    $state = ($AggregateState ?? "").ToLowerInvariant()
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
    # Load SmartM365.SharePoint module
    # ----------------------------------------------------------
    if ($EnableSharePointUpload) {
        $spModulePath = & { $d = $PSScriptRoot; while ($d) { $p = Join-Path $d 'Modules\SmartM365.SharePoint\SmartM365.SharePoint.psd1'; if (Test-Path -LiteralPath $p) { return $p }; $parent = Split-Path -Path $d -Parent; if ($parent -eq $d) { break }; $d = $parent }; throw 'SmartM365.SharePoint module not found.' }
        Write-Log "Loading SmartM365.SharePoint module: $spModulePath" "INFO" "SP"
        Import-Module $spModulePath -ErrorAction Stop
        Write-Log "SmartM365.SharePoint module loaded." "INFO" "SP"
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
    Invoke-CoreSmartM365Preflight -ScriptName $ScriptName -RequiredModules @('MSAL.PS') -OutputPaths @($ScriptCsvLogFolderPath, $LatestCsvFolderPath, $ArchivePath, $LogsPath, $WorkPath) -GraphAccessToken $token -GraphProbeUris @(
        'https://graph.microsoft.com/beta/deviceManagement/windowsFeatureUpdateProfiles?$top=1',
        'https://graph.microsoft.com/beta/deviceManagement/reports/exportJobs?$top=1'
    ) | Out-Null


    # ----------------------------------------------------------
    # Discover Feature Update policies
    # ----------------------------------------------------------
    $policies = Get-FeatureUpdatePolicies -Headers $headers
    if (-not $policies -or $policies.Count -eq 0) {
        throw "No Feature Update policies found (windowsFeatureUpdateProfiles)."
    }
    Write-Log "Feature Update policies discovered: $($policies.Count)" "INFO" "GRAPH"

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

    if ($EnableReadinessEnrichment) {
        if (-not (Test-Path $ReadinessCsvPath)) {
            Write-Log "Readiness CSV not found, enrichment skipped: $ReadinessCsvPath" "WARN" "READINESS"
        }
        else {
            Write-Log "Loading Readiness CSV: $ReadinessCsvPath" "INFO" "READINESS"
            $rd = Import-Csv -Path $ReadinessCsvPath
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
    # CSV export — DATA-ALL (atomic), DATA-LAST (atomic), Archive
    # ----------------------------------------------------------
    $spUploadStatus = "Disabled"
    $spUploadError  = ""
    $spDestPath     = ""

    if ($DryRun) {
        Write-Log "DryRun: skipping write of $CsvFinal." "INFO" "DRYRUN"
    }
    else {
        if (Test-Path $CsvTemp) { Remove-Item -Path $CsvTemp -Force -ErrorAction SilentlyContinue }
        $enrichedRows | Export-Csv -Path $CsvTemp -NoTypeInformation -Encoding UTF8
        Move-Item -Path $CsvTemp -Destination $CsvFinal -Force
        Write-Log "CSV exported: $CsvFinal" "INFO" "DATA"

        Copy-FileAtomic -SourcePath $CsvFinal -DestinationFinal $CsvLastFinal -DestinationTemp $CsvLastTemp
        Write-Log "DATA-LAST updated: $CsvLastFinal" "INFO" "DATA"

        $archiveBase = [System.IO.Path]::GetFileNameWithoutExtension($CsvName)
        Write-ArchiveCopyAndPrune -SourceCsv $CsvFinal -ArchiveFolder $ArchivePath -BaseNameWithoutExt $archiveBase -RunStamp $RunStamp -Keep $global:RetentionMaxCSV

        # ----------------------------------------------------------
        # SharePoint upload (non-blocking — uses same Graph token)
        # ----------------------------------------------------------
        if ($EnableSharePointUpload) {
            try {
                Refresh-GraphTokenIfNeeded -TenantId $TenantId -ClientId $AppId -Certificate $cert -Headers ([ref]$headers)
                Write-Log "SharePoint upload starting: $SP_SiteHostname$SP_SitePath / $SP_LibraryDisplayName / $SP_TargetFolderPath" "INFO" "SP"

                $spLogger  = { param($m, $l, $s) Write-Log $m $l $s }
                $spDriveId = Resolve-SmartM365SpDriveId -Headers $headers -SiteHostname $SP_SiteHostname -SitePath $SP_SitePath -LibraryDisplayName $SP_LibraryDisplayName -Logger $spLogger
                $spResult  = Invoke-SmartM365SpFileUpload -Headers $headers -LocalFilePath $CsvFinal -TargetFolderPath $SP_TargetFolderPath -DriveId $spDriveId -ChunkSize $SP_ChunkSize -Logger $spLogger

                $spUploadStatus = $spResult.Status
                $spDestPath     = $spResult.DestPath
                Write-Log "SharePoint upload: $spUploadStatus ($($spResult.DurationMs) ms) → $spDestPath" "INFO" "SP"

                if ($spResult.Status -eq "Error") {
                    $spUploadError = $spResult.Error
                    Write-Log "SharePoint upload error (non-blocking): $spUploadError" "WARN" "SP"
                }
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
<li><b>Readiness rows</b>: $readinessTotal  |  <b>Matched</b>: $readinessMatched  |  <b>NotMatched</b>: $readinessNotMatched  |  <b>MatchPct</b>: $readinessMatchPct%</li>
<li><b>Eligibility distribution</b>: $(Html-Encode $readinessEligibilityDist)</li>
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

        Write-Log "Summary email disabled; error emails only." "INFO" "SMTP"
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



