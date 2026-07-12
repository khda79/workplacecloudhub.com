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
      - Is aligned with the shared framework (SmartM365.Core.psm1)
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
1.8


.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication.
    Minimum Graph application permissions: DeviceManagementManagedDevices.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Version : 1.8
    Requires:
      - PowerShell 7+
      - Microsoft.Graph module (Graph SDK)
      - SmartM365.Core.psm1 in the same folder or in a Modules subfolder
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
$ScriptVersion = "1.8"
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
            $candidate = Join-Path $d "Modules\SmartM365.Core\SmartM365.Core.psm1"
            if (Test-Path -LiteralPath $candidate) { return $candidate }
            $parent = Split-Path -Path $d -Parent
            if ($parent -eq $d) { break }
            $d = $parent
        }
    }

    if (-not $coreModulePath) {
        Write-Host "SmartM365.Core.psm1 module not found in repository Modules\SmartM365.Core." -ForegroundColor Red
        exit 1
    }

    Import-Module -Name $coreModulePath -Force -ErrorAction Stop

} catch {
    Write-Host "Failed to initialize paths or load SmartM365.Core.psm1. $_" -ForegroundColor Red
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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCe0McdHwMU2dLL
# PCZwq83hxUHi3fT0tRyVGoZOsKpCxaCCBEgwggREMIICrKADAgECAhBxu0EivlCF
# tUbJPfe/Va5qMA0GCSqGSIb3DQEBCwUAMDoxODA2BgNVBAMML1NtYXJ0TTM2NSBP
# cmNoZXN0cmF0b3IgQ29kZSBTaWduaW5nIFNlbGYtU2lnbmVkMB4XDTI2MDcxMTIz
# MTc1MloXDTI5MDcxMTIzMjc1MVowOjE4MDYGA1UEAwwvU21hcnRNMzY1IE9yY2hl
# c3RyYXRvciBDb2RlIFNpZ25pbmcgU2VsZi1TaWduZWQwggGiMA0GCSqGSIb3DQEB
# AQUAA4IBjwAwggGKAoIBgQC4A+QoBzUXkXXMoVrptgMss1BNRwJhNcYop9CKHvJY
# QnBLkhSI10Z7EBCZsDSAfICechL0e7Lrwaz8/sTRQeITCKMRzxFe9Oq1CxZfRUh0
# U1T/m8+9q/OR0C6hCSZ9LvpiZExBSmQsQlXyl8smfFK2+gecLOQUPFD7gcpM03gv
# 6OkX/bLpBQZs52K3RnH+YKje0L6W985qxn1M5nDmC4rc2U90k4evzMMPOjTX7jZA
# PHOT3g6ByPWI2SNowO1ptXheS4KGjbx3IH+4+r4UwIPc32hauiAfjXr63inQdkII
# 7tYVI5GBiJB20Gzujm5KuHU9qVXMvAAk7WR9DBGdH4Pq5Or3WD58KV2Mazx0SWhV
# A4ikEEENTbaWIaFEYgWR2PAtPv7rt/p5ZK05fP7Nt/TfSHzBFQsKS4wFchiWQTVj
# kdAPuzsipnwiJyOSmQ7FppnuuhUxEq9ZkOigDLett9ZoY5oNcASOnpCWnxnWx/aq
# xDuJOnKBOGRly1KFUQ+OABUCAwEAAaNGMEQwDgYDVR0PAQH/BAQDAgeAMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBQkjQccxcT1k6xhYBW0XHlelX6nFjAN
# BgkqhkiG9w0BAQsFAAOCAYEAk3bN0vTJBIFnyLm4zxarRLfr6uEl9Y2Xk4P16AxG
# DDLN+Zd7T+oblgAIz4/0EHPJ3DsonLsjOnZBOp5iJr1nSxBy9Cs6K1T6k2mtSr93
# mOT2MSNDlLOFhk37U46yFDJHfX4rQLTmltOoUpeU7V7Cr5EnWJ4xbdmexZUx5vz+
# qeqqe86VxT00Npb5OXINvs8+gH85J+x4HWmrTDzruME1JLkX388g3AQvVd5Xf0YY
# 2InRPQ7Y0jrzccH6OSz14DHSnzN5pKzVzvv9aFDuZ+gCkbC8ZIr890I8WXxbYskX
# 8bTTP0Sa8Jhw22OCOwzDhFxxqivhbqHRybgQ6KdSoDxS51WHp3saGlWfwmFyWkIe
# L5eEpdz8r2vpTbaJVZnVT/SxpYobgZIn3zbss0JFiltcgguIoc+fNbMEUoqnEARQ
# dD4+fIPF32CUclDI6JpugYJLSuvJt6gy4k78A1jQaYTbdZ6Twt+Pup+3ocnWmeyV
# umYxx47CZmI93XUw5yflFPRUMYICgDCCAnwCAQEwTjA6MTgwNgYDVQQDDC9TbWFy
# dE0zNjUgT3JjaGVzdHJhdG9yIENvZGUgU2lnbmluZyBTZWxmLVNpZ25lZAIQcbtB
# Ir5QhbVGyT33v1WuajANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQow
# CKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcC
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAR0RWYwvQ6SdMdyGjb
# 7Lf0WJIwVrnTzrrImnAhgeOStDANBgkqhkiG9w0BAQEFAASCAYBs2ZMt3PD/Micj
# mNnmW+3boY9/ms44APURUMZlOoWLeNWns/HY9XKrlWBCFXASRSJ9fYfgTAxtCauE
# bakwKB4Wk/Y/jRgc9iuQvuZ71dJbZe7UB2QSdmmVq56zx1lhkoSwPkuM4NHhHxt0
# GpTlz4JI+56OM8iBOSpfHrsFets52X02Ve+1C+RMqdYpsauBPjY90VEMAO+Q42ya
# TvGItutENFXaOsqyTU86HOvXtNIJFqGau+jbENB/+m3k80Ml8s7L3TITG0Ph9u4T
# SVVp26XC1wBUzAw+VdwewgPf4Zgv9d+Av9J4rwA/D2ZODHOPtQOvVo1G+RuL5EdH
# fjgFU/7W8ZMnvM0jkpIXETrWq6pm9qkUoMCEq8SwPWeaZ5YJ1s/whb/dLLMnaTns
# SnGSt6VQnII9Yk9AcmEfo+kF+3F16vq5rjSDZXAJUbLADqs+74xZU1K891t0+NWf
# wp4TE7NBH7BxtkjUe6gmCsPVgzWzCdiudPDA9zP8qFfJ1NGnEAc=
# SIG # End signature block
