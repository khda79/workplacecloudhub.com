<#
.SYNOPSIS
    Generates a CSV report of Endpoint Analytics Work From Anywhere devices
    with Windows upgrade eligibility and hardware check flags.

.DESCRIPTION
    Uses Microsoft Graph (Endpoint Analytics) to query userExperienceAnalyticsWorkFromAnywhereDevice
    and exports a CSV containing:
      - DeviceName, Manufacturer, Model, OSVersion
      - UpgradeEligibility
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
    Optional output directory for CSV and log file. If not specified, the script directory is used.

.PARAMETER Connect
    Forces a (re)connection to Microsoft Graph (disconnects any existing Graph session first).

.PARAMETER InteractiveAuth
    Uses interactive authentication instead of app-only authentication when connecting to Microsoft Graph.

.PARAMETER Filter
    Optional OData filter applied on userExperienceAnalyticsWorkFromAnywhereDevice (server-side).
    Example:
        "contains(deviceName, 'LAPTOP')" or "upgradeEligibility ne 'none'".

.EXAMPLE
    .\Devices-UpgradeEligibility.ps1

.EXAMPLE
    .\Devices-UpgradeEligibility.ps1 -OutputPath "C:\Reports" -Connect
.NOTES
    Author: https://github.com/khda79/M365
    Version: 1.0
    Requires:
      - PowerShell 7+
      - Microsoft.Graph module (Graph SDK)
      - SmartM365.Core.psm1 in the same folder or in a Modules subfolder
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
    [string]$Filter
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

#region Global and safety settings

$ErrorActionPreference = "Stop"

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
        $searchRoot = if ($PSScriptRoot) { $PSScriptRoot } elseif ($ScriptRoot) { $ScriptRoot } elseif ($PSCommandPath) { Split-Path -Path $PSCommandPath -Parent } else { (Get-Location).Path }
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
$LatestCsvFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''
$AppId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AppId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$TenantId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$Thumb = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumb' -DefaultValue '0000000000000000000000000000000000000000'
$OrgDomain = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue 'contoso.onmicrosoft.com'
$LogAllRootPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LogAllRootPath' -DefaultValue ''

#endregion Global and safety settings

#region Resolve paths and load SmartM365.Core

try {
    $scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $ScriptCsvLogFolderPath = $scriptDir
    }
    else {
        if (-not (Test-Path -LiteralPath $OutputPath)) {
            New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
        }
        $ScriptCsvLogFolderPath = (Resolve-Path -LiteralPath $OutputPath).Path
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
        Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumb
    }

    WriteLogSmartM365 -Message "Connected to Microsoft Graph successfully." -Level "SUCCESS"
}

#endregion Microsoft Graph connection

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
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths $preflightOutputPaths -GraphProbeUris @('https://graph.microsoft.com/beta/deviceManagement/userExperienceAnalyticsWorkFromAnywhereMetrics') | Out-Null

    # ---------------- Work From Anywhere / Upgrade Eligibility (Endpoint Analytics) ----------------

    WriteLogSmartM365 -Message "Querying Endpoint Analytics Work From Anywhere (upgrade eligibility) from Microsoft Graph..." -Level "INFO"

    $wfaSelectProps = @(
        "deviceName",
        "manufacturer",
        "model",
        "osVersion",
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

    # Step 1: get Work From Anywhere metrics (usually one item)
    $metrics = $null
    try {
        $metricsResponse = Invoke-MgGraphRequest -Method GET -Uri "$baseUri/userExperienceAnalyticsWorkFromAnywhereMetrics"
        $metrics         = $metricsResponse.value
    }
    catch {
        $errText = $_.Exception.Message
        WriteLogSmartM365 -Message ("Call to userExperienceAnalyticsWorkFromAnywhereMetrics failed: {0}" -f $errText) -Level "ERROR"

        if ($errText -like "*No OData route exists that match template*") {
            WriteLogSmartM365 -Message "Endpoint Analytics 'Work From Anywhere' API route is not available in this tenant. Check that Endpoint Analytics / Work from anywhere is enabled in Intune." -Level "WARNING"
        }

        Write-Host "Work From Anywhere Graph API is not available in this tenant. Cannot generate upgrade eligibility report via Graph." -ForegroundColor Yellow
        WriteLogSmartM365 -Message "===== Devices Upgrade Eligibility inventory finished (WFA API not available) =====" -Level "INFO"
        exit 1
    }

    if (-not $metrics -or $metrics.Count -eq 0) {
        WriteLogSmartM365 -Message "No userExperienceAnalyticsWorkFromAnywhereMetrics found. Make sure Endpoint Analytics / Work From Anywhere is enabled." -Level "WARNING"
        Write-Host "No Work From Anywhere metrics found. Check Endpoint Analytics configuration." -ForegroundColor Yellow
        WriteLogSmartM365 -Message "===== Devices Upgrade Eligibility inventory finished (no metrics) =====" -Level "INFO"
        exit 1
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
    $reportData = $allDevices | Select-Object `
        @{ Name = "DeviceName";                     Expression = { $_.deviceName } }, `
        @{ Name = "Manufacturer";                   Expression = { $_.manufacturer } }, `
        @{ Name = "Model";                          Expression = { $_.model } }, `
        @{ Name = "OSVersion";                      Expression = { $_.osVersion } }, `
        @{ Name = "UpgradeEligibility";             Expression = { $_.upgradeEligibility } }, `
        @{ Name = "RamCheckFailed";                 Expression = { $_.ramCheckFailed } }, `
        @{ Name = "StorageCheckFailed";             Expression = { $_.storageCheckFailed } }, `
        @{ Name = "ProcessorCoreCountCheckFailed";  Expression = { $_.processorCoreCountCheckFailed } }, `
        @{ Name = "ProcessorSpeedCheckFailed";      Expression = { $_.processorSpeedCheckFailed } }, `
        @{ Name = "TPMCheckFailed";                 Expression = { $_.tpmCheckFailed } }, `
        @{ Name = "SecureBootCheckFailed";          Expression = { $_.secureBootCheckFailed } }, `
        @{ Name = "ProcessorFamilyCheckFailed";     Expression = { $_.processorFamilyCheckFailed } }, `
        @{ Name = "Processor64BitCheckFailed";      Expression = { $_.processor64BitCheckFailed } }, `
        @{ Name = "OSCheckFailed";                  Expression = { $_.osCheckFailed } }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$reportName = "Intune_Devices_UpgradeEligibility"
    $csvFileName = "{0}_{1}.csv" -f $reportName, $timestamp
    $csvPath = Join-Path -Path $ScriptCsvLogFolderPath -ChildPath $csvFileName
    $latestCsvPath = if ([string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) { $null } else { Join-Path -Path $LatestCsvFolderPath -ChildPath "$reportName.csv" }

    WriteLogSmartM365 -Message ("Exporting report to CSV: {0}" -f $csvPath) -Level "INFO"

    $reportData | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
    if ($latestCsvPath) {
        Copy-Item -Path $csvPath -Destination $latestCsvPath -Force
        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $latestCsvPath
        WriteLogSmartM365 -Message ("Latest CSV updated: {0}" -f $latestCsvPath) -Level "SUCCESS"
    }
    else {
        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $csvPath
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



