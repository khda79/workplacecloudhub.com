<#
.SYNOPSIS
    Generates a CSV report of Endpoint Analytics Work From Anywhere devices
    with Windows upgrade eligibility, normalized readiness join keys, and hardware check flags.

.DESCRIPTION
    Uses Microsoft Graph (Endpoint Analytics) to export Windows upgrade readiness summary data
    and exports a CSV containing:
      - DeviceName, Manufacturer, Model, OSVersion
      - NormalizedDeviceName, GraphId, AzureAdJoinType, UpgradeEligibility, UpgradeEligibilityLabel
      - RamCheckFailed, StorageCheckFailed
      - ProcessorCoreCountCheckFailed, ProcessorSpeedCheckFailed
      - TPMCheckFailed, SecureBootCheckFailed
      - ProcessorFamilyCheckFailed, Processor64BitCheckFailed
      - OSCheckFailed

    The script:
      - Is aligned with the shared framework (SmartM365.Core.psd1)
      - Uses centralized logging (console + log file)
      - Uses app-only certificate authentication by default
      - Supports interactive authentication with -InteractiveAuth (for troubleshooting)
      - Sends an error email with the log attached in case of global failure (if Send-SmartM365Mail is available)

.PARAMETER OutputPath
    Optional output directory for CSV and log file. If not specified, ScriptCsvLogFolderPath from local JSON is used.

.PARAMETER Connect
    Forces a (re)connection to Microsoft Graph (disconnects any existing Graph session first).

.PARAMETER InteractiveAuth
    Uses interactive authentication instead of app-only authentication when connecting to Microsoft Graph.

.PARAMETER Filter
    Optional OData filter applied only when -AttemptDeviceDetails is used with the legacy device-level route.
    Example:
        "contains(deviceName, 'LAPTOP')" or "upgradeEligibility ne 'none'".

.PARAMETER MaxDevices
    Optional local row limit after Graph retrieval. 0 means all rows.

.EXAMPLE
    .\Devices-UpgradeEligibility.ps1

.EXAMPLE
    .\Devices-UpgradeEligibility.ps1 -OutputPath "C:\Reports" -Connect
.VERSION
1.11


.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication.
    Minimum Graph application permissions: DeviceManagementManagedDevices.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Version : 1.9
    Requires:
      - PowerShell 7+
      - Microsoft.Graph module (Graph SDK)
      - SmartM365.Core.psd1 in the same folder or in a Modules subfolder
    Minimum application permissions: DeviceManagementManagedDevices.Read.All
#>

[CmdletBinding()]
param (
    [string]$Tenant = 'test',

    [Parameter(Mandatory = $false)]
    [string]$OutputPath,

    [Parameter(Mandatory = $false)]
    [switch]$Connect,

    [Parameter(Mandatory = $false)]
    [switch]$InteractiveAuth,

    [Parameter(Mandatory = $false)]
    [string]$Filter,

    [Parameter(Mandatory = $false)]
    [int]$MaxDevices = 0,

    [Parameter(Mandatory = $false)]
    [switch]$AttemptDeviceDetails,
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
Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot | Out-Null

#region Global and safety settings

$ErrorActionPreference = "Stop"
$ScriptVersion = "1.11"
$TaskName = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion"

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    exit 1
}

# ==========================================================
# App-only authentication parameters (same app as other inventory scripts)
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
        $searchRoot = if ($PSScriptRoot) { $PSScriptRoot } elseif ($ScriptRoot) { $ScriptRoot } elseif ($PSCommandPath) { Split-Path -Path $PSCommandPath -Parent } else { (Get-Location).Path }
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
$LatestCsvFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''
$ConfiguredScriptCsvLogFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ScriptCsvLogFolderPath' -DefaultValue ''
$AppId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AppId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$TenantId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$Thumb = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumb' -DefaultValue '0000000000000000000000000000000000000000'
$OrgDomain = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue 'contoso.onmicrosoft.com'
$LogAllRootPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LogAllRootPath' -DefaultValue ''

#endregion Global and safety settings

#region Resolve paths and load SmartM365.Core

try {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

    if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
        if (-not (Test-Path -LiteralPath $OutputPath)) {
            New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        }
        $ScriptCsvLogFolderPath = (Resolve-Path -LiteralPath $OutputPath).Path
    }
    elseif (-not [string]::IsNullOrWhiteSpace($ConfiguredScriptCsvLogFolderPath)) {
        if (-not (Test-Path -LiteralPath $ConfiguredScriptCsvLogFolderPath)) {
            New-Item -Path $ConfiguredScriptCsvLogFolderPath -ItemType Directory -Force | Out-Null
        }
        $ScriptCsvLogFolderPath = (Resolve-Path -LiteralPath $ConfiguredScriptCsvLogFolderPath).Path
    }
    else {
        $ScriptCsvLogFolderPath = $scriptDir
    }

    $coreModulePath = & {
        $d = $scriptDir
        while ($d) {
            $candidate = Join-Path $d "Modules\SmartM365.Core\SmartM365.Core.psd1"
            if (Test-Path -LiteralPath $candidate) { return $candidate }
            $parent = Split-Path -Path $d -Parent
            if ($parent -eq $d) { break }
            $d = $parent
        }
    }

    if (-not $coreModulePath) {
        Write-Host "SmartM365.Core.psd1 module not found in repository Modules\SmartM365.Core." -ForegroundColor Red
        exit 1
    }

    Import-Module -Name $coreModulePath -MinimumVersion '1.0.24' -Force -ErrorAction Stop

} catch {
    Write-Host "Failed to initialize paths or load SmartM365.Core.psd1. $_" -ForegroundColor Red
    exit 1
}

#endregion Resolve paths and load SmartM365.Core

#region Logging helpers

# If SmartM365.Core does not define $global:LogTextFile yet, create one for this script
if (-not $global:LogTextFile) {
    $logFileName = "Devices-UpgradeEligibility_{0:yyyyMMdd_HHmmss}.log" -f (Get-Date)
    $logRoot = if ([string]::IsNullOrWhiteSpace($LogAllRootPath)) {
        $ScriptCsvLogFolderPath
    } else {
        Join-Path -Path $LogAllRootPath -ChildPath "Devices-UpgradeEligibility"
    }
    if (-not (Test-Path -LiteralPath $logRoot)) {
        New-Item -Path $logRoot -ItemType Directory -Force | Out-Null
    }
    $global:LogTextFile = Join-Path -Path $logRoot -ChildPath $logFileName
    Set-SmartM365CoreContext -RunId $runId -RunOutputRoot $ScriptCsvLogFolderPath -LatestOutputRoot $LatestCsvFolderPath -LogPath $global:LogTextFile
}

function WriteLogSmartM365 {
    param(
        [string]$Message,
        [string]$Level = ""
    )

    try {
        WriteLog -Message $Message -Level $Level
    }
    catch {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        @([regex]::Split(([string]$Message), '\r?\n')) | ForEach-Object {
            Write-Host ("{0} [{1}] {2}" -f $timestamp, $Level, $_) -ForegroundColor Cyan
        }
    }

    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        if ([string]::IsNullOrWhiteSpace($Level)) {
            $Level = "INFO"
        }

        $logEntry = @([regex]::Split(([string]$Message), '\r?\n') | ForEach-Object { "{0} [{1}] {2}" -f $timestamp, $Level, $_ })
        Add-Content -Path $global:LogTextFile -Value $logEntry
    }
    catch {
        Write-Host "Failed to write to log file $global:LogTextFile. $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

#endregion Logging helpers

#region Microsoft Graph connection

function Ensure-GraphConnection {
    param(
        [switch]$ForceReconnect,
        [switch]$UseInteractiveAuth
    )

    if ($ForceReconnect) {
        WriteLogSmartM365 -Message "Connect switch specified: forcing reconnection to Microsoft Graph..." -Level "INFO"

        try {
            if (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue) {
                Disconnect-MgGraph -ErrorAction SilentlyContinue
            }
        }
        catch {
            WriteLogSmartM365 -Message ("Disconnect-MgGraph failed: {0}" -f $_.Exception.Message) -Level "WARNING"
        }
    }

    $graphContext = $null
    if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
        try {
            $graphContext = Get-MgContext -ErrorAction SilentlyContinue
        }
        catch {
            $graphContext = $null
        }
    }

    if ($graphContext -and -not $ForceReconnect) {
        WriteLogSmartM365 -Message "Existing Microsoft Graph session detected. Reusing current connection." -Level "INFO"
        return
    }

    if ($UseInteractiveAuth) {
        WriteLogSmartM365 -Message "Connecting to Microsoft Graph using interactive authentication..." -Level "INFO"
        Connect-MgGraph -Scopes "DeviceManagement.Read.All","DeviceManagementManagedDevices.Read.All"
    }
    else {
        WriteLogSmartM365 -Message "Connecting to Microsoft Graph using app-only certificate authentication..." -Level "INFO"
        Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumb -NoWelcome
    }

    WriteLogSmartM365 -Message "Connected to Microsoft Graph successfully." -Level "SUCCESS"
}

#endregion Microsoft Graph connection

function Normalize-DeviceName {
    param([Parameter(Mandatory = $false)][string]$DeviceName)

    if ([string]::IsNullOrWhiteSpace($DeviceName)) { return $null }
    $normalized = $DeviceName.Trim().ToLowerInvariant()
    if ($normalized -match '^([^\.]+)\.') { $normalized = $Matches[1] }
    $normalized = -join ($normalized.ToCharArray() | Where-Object { $_ -match '[a-z0-9\-]' })
    return $normalized
}

function Get-UpgradeEligibilityLabel {
    param([Parameter(Mandatory = $false)][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return 'Unknown' }

    switch -Regex ($Value.Trim()) {
        '^(0|eligible|capable|ready)$' { return 'Eligible' }
        '^(1|notEligible|notCapable|notReady)$' { return 'NotEligible' }
        '^(2|unknown|undetermined|notApplicable)$' { return 'NotApplicableOrUnknown' }
        default { return $Value }
    }
}

function Export-HardwareReadinessSummary {
    param(
        [Parameter(Mandatory = $true)][string]$BaseUri,
        [Parameter(Mandatory = $false)][string]$Reason = ''
    )

    $endpoint = "$BaseUri/userExperienceAnalyticsWorkFromAnywhereHardwareReadinessMetric"
    WriteLogSmartM365 -Message ("Exporting Work From Anywhere hardware readiness summary: {0}" -f $endpoint) -Level "INFO"
    $summary = Invoke-MgGraphRequest -Method GET -Uri $endpoint

    $devicesSummaryEndpoint = "$BaseUri/userExperienceAnalyticsSummarizeWorkFromAnywhereDevices()"
    $devicesSummary = $null
    try {
        WriteLogSmartM365 -Message ("Exporting Work From Anywhere device scope summary: {0}" -f $devicesSummaryEndpoint) -Level "INFO"
        $devicesSummary = Invoke-MgGraphRequest -Method GET -Uri $devicesSummaryEndpoint
    }
    catch {
        WriteLogSmartM365 -Message ("Optional Work From Anywhere device scope summary unavailable: {0}" -f $_.Exception.Message) -Level "INFO"
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $summaryReportName = "Intune_Devices_UpgradeEligibility_Summary"
    $summaryCsvPath = Join-Path -Path $ScriptCsvLogFolderPath -ChildPath ("{0}_{1}.csv" -f $summaryReportName, $timestamp)
    $latestSummaryCsvPath = if ([string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) { $null } else { Join-Path -Path $LatestCsvFolderPath -ChildPath "$summaryReportName.csv" }
    $eligiblePercent = if ([int]$summary.totalDeviceCount -gt 0) { [math]::Round(([double]$summary.upgradeEligibleDeviceCount / [double]$summary.totalDeviceCount) * 100, 2) } else { 0 }

    $row = [pscustomobject]@{
        ReportType                                = 'HardwareReadinessSummary'
        ReportMode                                = 'Summary'
        GraphId                                   = [string]$summary.id
        TotalDeviceCount                          = [int]$summary.totalDeviceCount
        UpgradeEligibleDeviceCount                = [int]$summary.upgradeEligibleDeviceCount
        UpgradeEligiblePercentage                 = $eligiblePercent
        OSCheckFailedPercentage                   = [double]$summary.osCheckFailedPercentage
        Processor64BitCheckFailedPercentage       = [double]$summary.processor64BitCheckFailedPercentage
        ProcessorCoreCountCheckFailedPercentage   = [double]$summary.processorCoreCountCheckFailedPercentage
        ProcessorFamilyCheckFailedPercentage      = [double]$summary.processorFamilyCheckFailedPercentage
        ProcessorSpeedCheckFailedPercentage       = [double]$summary.processorSpeedCheckFailedPercentage
        RamCheckFailedPercentage                  = [double]$summary.ramCheckFailedPercentage
        SecureBootCheckFailedPercentage           = [double]$summary.secureBootCheckFailedPercentage
        StorageCheckFailedPercentage              = [double]$summary.storageCheckFailedPercentage
        TPMCheckFailedPercentage                  = [double]$summary.tpmCheckFailedPercentage
        WfaTotalDevices                           = [int]$devicesSummary.totalDevices
        WfaIntuneDevices                          = [int]$devicesSummary.intuneDevices
        WfaWindows10Devices                       = [int]$devicesSummary.windows10Devices
        WfaUnsupportedOSVersionDevices            = [int]$devicesSummary.unsupportedOSversionDevices
        WfaDevicesNotAutopilotRegistered          = [int]$devicesSummary.devicesNotAutopilotRegistered
        WfaDevicesWithoutAutopilotProfileAssigned = [int]$devicesSummary.devicesWithoutAutopilotProfileAssigned
        WfaDevicesWithoutCloudIdentity            = [int]$devicesSummary.devicesWithoutCloudIdentity
        WfaCoManagedDevices                       = [int]$devicesSummary.coManagedDevices
        WfaTenantAttachDevices                    = [int]$devicesSummary.tenantAttachDevices
        WfaWindows10DevicesWithoutTenantAttach    = [int]$devicesSummary.windows10DevicesWithoutTenantAttach
        SourceGraphEndpoint                       = $endpoint
        DeviceScopeSummaryGraphEndpoint           = $devicesSummaryEndpoint
        FallbackReason                            = $Reason
        ExportDateTime                            = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        RunId                                     = $runId
    }

    if ($latestSummaryCsvPath) {
        Export-SmartM365Csv -Data @($row) -TimestampedPath $summaryCsvPath -LatestPath $latestSummaryCsvPath | Out-Null
        WriteLogSmartM365 -Message ("Latest summary CSV updated: {0}" -f $latestSummaryCsvPath) -Level "SUCCESS"
    }
    else {
        Export-SmartM365Csv -Data @($row) -TimestampedPath $summaryCsvPath | Out-Null
    }

    WriteLogSmartM365 -Message ("Hardware readiness summary exported: {0}; TotalDeviceCount={1}; UpgradeEligibleDeviceCount={2}; UpgradeEligiblePercentage={3}" -f $summaryCsvPath, $row.TotalDeviceCount, $row.UpgradeEligibleDeviceCount, $row.UpgradeEligiblePercentage) -Level "SUCCESS"
    Write-Host "Hardware readiness summary exported: $summaryCsvPath" -ForegroundColor Green
}
#region Main logic

try {
    WriteLogSmartM365 -Message "===== Devices Upgrade Eligibility inventory started =====" -Level "INFO"
    WriteLogSmartM365 -Message ("Output directory: {0}" -f $ScriptCsvLogFolderPath) -Level "INFO"

    # Ensure Graph is connected
    Ensure-GraphConnection -ForceReconnect:$Connect -UseInteractiveAuth:$InteractiveAuth
    $preflightOutputPaths = @($ScriptCsvLogFolderPath)
    if (-not [string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) {
        $preflightOutputPaths += $LatestCsvFolderPath
    }
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths $preflightOutputPaths -RequiredGraphApplicationPermissions @('DeviceManagementManagedDevices.Read.All') -GraphProbeUris @('https://graph.microsoft.com/beta/deviceManagement/userExperienceAnalyticsWorkFromAnywhereHardwareReadinessMetric') | Out-Null

    # ---------------- Work From Anywhere / Upgrade Eligibility (Endpoint Analytics) ----------------

    WriteLogSmartM365 -Message "Querying Endpoint Analytics Work From Anywhere (upgrade eligibility) from Microsoft Graph..." -Level "INFO"

    $wfaSelectProps = @(
        "id",
        "deviceName",
        "manufacturer",
        "model",
        "osVersion",
        "azureAdJoinType",
        "upgradeEligibility",
        "ramCheckFailed",
        "storageCheckFailed",
        "processorCoreCountCheckFailed",
        "processorSpeedCheckFailed",
        "tpmCheckFailed",
        "secureBootCheckFailed",
        "processorFamilyCheckFailed",
        "processor64BitCheckFailed",
        "osCheckFailed"
    )

    # Use beta endpoint for Endpoint Analytics WFA because v1.0 route may not be available
    $graphBase = "https://graph.microsoft.com/beta"
    $baseUri   = "$graphBase/deviceManagement"
    if (-not $AttemptDeviceDetails) {
        if ($MaxItems -gt 0) {
            throw "-MaxItems is not supported in summary-first mode because the supported Graph endpoints return tenant-level counters, not device rows. Use -AttemptDeviceDetails with -MaxItems when the device-level Graph route is available."
        }
        if (-not [string]::IsNullOrWhiteSpace($Filter)) {
            WriteLogSmartM365 -Message "Filter is ignored in summary mode. Use -AttemptDeviceDetails to apply device-level filtering if the Graph route becomes available." -Level "INFO"
        }
        if ($MaxDevices -gt 0) {
            WriteLogSmartM365 -Message "MaxDevices is ignored in summary mode. Use -AttemptDeviceDetails to apply a device-level row limit if the Graph route becomes available." -Level "INFO"
        }

        Export-HardwareReadinessSummary -BaseUri $baseUri -Reason 'Summary-first default. The device-level Work From Anywhere metricDevices route is currently unavailable in this tenant/API.'
        WriteLogSmartM365 -Message "===== Devices Upgrade Eligibility inventory completed successfully =====" -Level "INFO"
        exit 0
    }

    WriteLogSmartM365 -Message "AttemptDeviceDetails specified: trying legacy Work From Anywhere metricDevices route." -Level "INFO"

    # Step 1: get Work From Anywhere metrics (usually one item)
    $metrics = $null
    try {
        $metricsResponse = Invoke-MgGraphRequest -Method GET -Uri "$baseUri/userExperienceAnalyticsWorkFromAnywhereMetrics"
        $metrics         = $metricsResponse.value
    }
    catch {
        $errText = $_.Exception.Message
        WriteLogSmartM365 -Message ("Device-level Work From Anywhere metrics route unavailable; exporting supported summary instead: {0}" -f $errText) -Level "INFO"
        Export-HardwareReadinessSummary -BaseUri $baseUri -Reason ("Device-level Work From Anywhere metrics route unavailable: {0}" -f $errText)
        WriteLogSmartM365 -Message "===== Devices Upgrade Eligibility inventory completed successfully =====" -Level "INFO"
        exit 0
    }
    if (-not $metrics -or $metrics.Count -eq 0) {
        WriteLogSmartM365 -Message "No userExperienceAnalyticsWorkFromAnywhereMetrics found; exporting supported hardware readiness summary instead." -Level "INFO"
        Export-HardwareReadinessSummary -BaseUri $baseUri -Reason 'No device-level Work From Anywhere metrics returned.'
        WriteLogSmartM365 -Message "===== Devices Upgrade Eligibility inventory completed successfully =====" -Level "INFO"
        exit 0
    }
    WriteLogSmartM365 -Message ("Number of Work From Anywhere metrics objects returned: {0}" -f $metrics.Count) -Level "DEBUG"

    $allDevices = @()

    foreach ($metric in $metrics) {
        $metricId = $metric.id
        WriteLogSmartM365 -Message ("Processing Work From Anywhere metric id {0}..." -f $metricId) -Level "DEBUG"

        $relative = "$baseUri/userExperienceAnalyticsWorkFromAnywhereMetrics/$metricId/metricDevices"

        $queryParts = @()

        if (-not [string]::IsNullOrWhiteSpace($Filter)) {
            WriteLogSmartM365 -Message ("Applying OData filter on metricDevices: {0}" -f $Filter) -Level "INFO"
            $queryParts += "`$filter=$Filter"
        }

        $queryParts += "`$select=$($wfaSelectProps -join ',')"

        $queryString = $queryParts -join "&"
        $nextLink    = "$relative`?$queryString"

        while ($nextLink) {
            WriteLogSmartM365 -Message ("Calling: {0}" -f $nextLink) -Level "DEBUG"
            $page = Invoke-MgGraphRequest -Method GET -Uri $nextLink

            if ($page.value) {
                $allDevices += $page.value
            }

            if ($page.'@odata.nextLink') {
                $nextLink = $page.'@odata.nextLink'
            }
            else {
                $nextLink = $null
            }
        }
    }

    if (-not $allDevices -or $allDevices.Count -eq 0) {
        WriteLogSmartM365 -Message "No devices returned from Work From Anywhere Endpoint Analytics." -Level "WARNING"
        Write-Host "No devices found in Work From Anywhere metrics. Verify that Endpoint Analytics is populated." -ForegroundColor Yellow
        WriteLogSmartM365 -Message "===== Devices Upgrade Eligibility inventory finished (no devices) =====" -Level "INFO"
        exit 1
    }

    WriteLogSmartM365 -Message ("Retrieved {0} devices from Work From Anywhere metrics." -f $allDevices.Count) -Level "INFO"

    # Project devices into a clean object with guaranteed column order
    if ($MaxDevices -gt 0 -and $allDevices.Count -gt $MaxDevices) {
        WriteLogSmartM365 -Message ("MaxDevices={0}: limiting local report rows from {1} to {0}." -f $MaxDevices, $allDevices.Count) -Level "WARNING"
        $allDevices = @($allDevices | Select-Object -First $MaxDevices)
    }

    $exportDateTime = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $projectedRows = @($allDevices | ForEach-Object {
        $deviceName = [string]$_.deviceName
        [pscustomobject]@{
            DeviceName                    = $deviceName
            NormalizedDeviceName          = Normalize-DeviceName -DeviceName $deviceName
            GraphId                       = [string]$_.id
            Manufacturer                  = [string]$_.manufacturer
            Model                         = [string]$_.model
            OSVersion                     = [string]$_.osVersion
            AzureAdJoinType               = [string]$_.azureAdJoinType
            UpgradeEligibility            = [string]$_.upgradeEligibility
            UpgradeEligibilityLabel       = Get-UpgradeEligibilityLabel -Value ([string]$_.upgradeEligibility)
            RamCheckFailed                = $_.ramCheckFailed
            StorageCheckFailed            = $_.storageCheckFailed
            ProcessorCoreCountCheckFailed = $_.processorCoreCountCheckFailed
            ProcessorSpeedCheckFailed     = $_.processorSpeedCheckFailed
            TPMCheckFailed                = $_.tpmCheckFailed
            SecureBootCheckFailed         = $_.secureBootCheckFailed
            ProcessorFamilyCheckFailed    = $_.processorFamilyCheckFailed
            Processor64BitCheckFailed     = $_.processor64BitCheckFailed
            OSCheckFailed                 = $_.osCheckFailed
            ExportDateTime                = $exportDateTime
            RunId                         = $runId
        }
    })

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $reportName = "Intune_Devices_UpgradeEligibility"
    $duplicateReportName = "Intune_Devices_UpgradeEligibility_DuplicatesByName"
    $csvFileName = "{0}_{1}.csv" -f $reportName, $timestamp
    $csvPath = Join-Path -Path $ScriptCsvLogFolderPath -ChildPath $csvFileName
    $latestCsvPath = if ([string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) { $null } else { Join-Path -Path $LatestCsvFolderPath -ChildPath "$reportName.csv" }
    $duplicateCsvPath = Join-Path -Path $ScriptCsvLogFolderPath -ChildPath ("{0}_{1}.csv" -f $duplicateReportName, $timestamp)
    $latestDuplicateCsvPath = if ([string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) { $null } else { Join-Path -Path $LatestCsvFolderPath -ChildPath "$duplicateReportName.csv" }

    $duplicateGroups = @(
        $projectedRows |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_.NormalizedDeviceName) } |
            Group-Object -Property NormalizedDeviceName |
            Where-Object { $_.Count -gt 1 }
    )

    if ($duplicateGroups.Count -gt 0) {
        $duplicateRows = foreach ($group in $duplicateGroups) {
            foreach ($row in $group.Group) {
                $row | Select-Object @{Name='DuplicateKey';Expression={$group.Name}}, @{Name='DuplicateCount';Expression={$group.Count}}, *
            }
        }

        if ($latestDuplicateCsvPath) {
            Export-SmartM365Csv -Data @($duplicateRows) -TimestampedPath $duplicateCsvPath -LatestPath $latestDuplicateCsvPath | Out-Null
        }
        else {
            Export-SmartM365Csv -Data @($duplicateRows) -TimestampedPath $duplicateCsvPath | Out-Null
        }
        WriteLogSmartM365 -Message ("Duplicate readiness audit exported: {0}; duplicate groups: {1}" -f $duplicateCsvPath, $duplicateGroups.Count) -Level "WARNING"
    }

    $seenNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $reportData = @(
        $projectedRows |
            Sort-Object -Property NormalizedDeviceName, DeviceName |
            Where-Object {
                if ([string]::IsNullOrWhiteSpace($_.NormalizedDeviceName)) { $true }
                else { $seenNames.Add($_.NormalizedDeviceName) }
            }
    )

    WriteLogSmartM365 -Message ("Readiness rows after de-duplication: {0}; duplicates removed: {1}" -f $reportData.Count, ($projectedRows.Count - $reportData.Count)) -Level "INFO"
    WriteLogSmartM365 -Message ("Exporting report to CSV: {0}" -f $csvPath) -Level "INFO"

    if ($latestCsvPath) {
        Export-SmartM365Csv -Data @($reportData) -TimestampedPath $csvPath -LatestPath $latestCsvPath | Out-Null
        WriteLogSmartM365 -Message ("Latest CSV updated: {0}" -f $latestCsvPath) -Level "SUCCESS"
    }
    else {
        Export-SmartM365Csv -Data @($reportData) -TimestampedPath $csvPath | Out-Null
    }

    WriteLogSmartM365 -Message ("Report generated successfully: {0}" -f $csvPath) -Level "SUCCESS"
    Write-Host "Report generated successfully: $csvPath" -ForegroundColor Green

    WriteLogSmartM365 -Message "===== Devices Upgrade Eligibility inventory completed successfully =====" -Level "INFO"
}
catch {
    $globalError = $_
    WriteLogSmartM365 -Message ("Global error in Devices Upgrade Eligibility inventory: {0}" -f $globalError) -Level "ERROR"
    Write-Host "A global error occurred. Check the log file for details." -ForegroundColor Red

    try {
        if (Get-Command -Name Send-SmartM365Mail -ErrorAction SilentlyContinue) {

            $title = "Devices upgrade eligibility - ERROR"
            $msg = @"
An error occurred in script $($MyInvocation.MyCommand.Name) on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss").

Error message:
$($globalError.Exception.Message)

See attached log file for details:
$($global:LogTextFile)
"@

            $bodyHtml = $msg
            if (Get-Command -Name NewSimpleEmailBody -ErrorAction SilentlyContinue) {
                $bodyHtml = NewSimpleEmailBody -Title $title -Message $msg
            }

            $attachments = @()
            if ($global:LogTextFile -and (Test-Path $global:LogTextFile)) {
                $attachments = @($global:LogTextFile)
            }

            Send-SmartM365Mail -Subject $title -BodyHtml $bodyHtml -Attachments $attachments
            WriteLogSmartM365 -Message "Error notification email sent via Send-SmartM365Mail." -Level "INFO"
        }
        else {
            WriteLogSmartM365 -Message "Send-SmartM365Mail is not available in this environment; skipping error email notification." -Level "WARNING"
        }
    }
    catch {
        Write-Host "Failed to send error notification email: $($_.Exception.Message)" -ForegroundColor Yellow
        WriteLogSmartM365 -Message ("Failed to send error notification email: {0}" -f $_.Exception.Message) -Level "WARNING"
    }

    exit 1
}

#endregion Main logic
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBGqfOAD0NGez5w
# FQcbAFa620nyuYfdZJ+j8EuLPz9h4aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPLuXAadlv49trkkFmW74GJUWfNmGuIp6u3c/TUKAmO2MA0GCSqG
# SIb3DQEBAQUABIIBgHFzaOL+f7OFv8OflMhUFMJNgBpCSZAWTDQHm1ZuwAcAiniQ
# x1l9sSrVNICySyCnjzc6wTbHX3soE1adeWqz3zE7SaZqzg7npFZcmUF+g2MYCv9a
# 8I5GxTWpNy0jqgj8q0aISlqLPXmV9ksolG/thMBcL4nzNmvGRBZgg42RbmQaggcK
# j44eY2UYYDljzG22lWVRb6lhBNYCDfc6ev2Sdz31sV4EF9IZduaplECA7Iui3VZ9
# 3Ko8jfoXtC0MSMmSi3TwdtrKuGquuc1wnjX28tzcyDglpcDIAzBJcLWZS+rctrcI
# nHzF1gl3G2b0qXq1p7iL6LZOJhSC+mGaiht9vL2e2Z1AuqYiHIL7ht9bEW/ZPwgP
# ckfB1Q9lmGYFN1eg6IKGZ1e1O6yjpQ3OD4bjAXtcq8rmLKnbd/Auj1lslfky2p3A
# m2NcezWoNDU2dKaXskx1l7ctI0NfHN8fOWgDDlHT0k05MqK+nv34rf3k/3cq0rQd
# /S1HPncQz+RedoHKPqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODUw
# MzJaMC8GCSqGSIb3DQEJBDEiBCA4Kw19vpnozLmOOfvoQNQjThdgqhHA8N8j2+9b
# OJV/zDANBgkqhkiG9w0BAQEFAASCAgA0pThX2ZNvUnvCHuI8ArVvAuWQvO8MYUAt
# U7P2cHB5JT2hbcmM24UBVO/SY8PXB/pp6hkNOCnriib4Ak7W7rfSKwMbIgUiMy+P
# ZQlPMU26TTCFab7EnSYqTkIxaWPFzPSPB50Hf0NPvn+dFBU/8eVEFHsQOSOGgdDx
# p0iJ9F+b07GOn9MMfyimm0DE2T50gU7NvuTWX73BWbSZh5xvZjRBo859urLgHp12
# sEmSsG7B1OOcUWbq5/mvqBkoZiwshBSSQRlGaqbnk7toP8r/1xKtOYMsyFwP9zbZ
# AQLDDjdg7vzHebw0fKSkStnwGJpjS5sUWzwWsxakHgUkcpYq+evANrUhWngp5npp
# SEGG0+Dgl2yfkKGpbNAwGmnaBmWiAwDc9v2+tQs/6RZTkxlTDwgPpOp4QzRO4OZg
# KZH1IN5RlBu1n7vnzTOdrvPd0nJUFLdF9zohMSL7L+BEWEUH5uFn8CiexUxGIpf6
# ue5KbBcBG0lHvnpMnclqgcB3AqAmgxqKbRjN8bRTOUadXtVrL4Omj19npYwjpRJU
# 0NC1w0o3I/Sb6ejUIorKxqiYcKwM2txz1Z1l6nfkB5r8Vzeeed0b0EbX7Vmcii3R
# VJKHpigMJP1CCQ7TJd+PgEvXA1c5ZRrf0M+J4U2P+I6Eksk6ArLYOWG3fWxjgERX
# 0MIrfZ/GSQ==
# SIG # End signature block
