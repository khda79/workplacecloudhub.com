<#
.SYNOPSIS
    Generates a device system inventory (SecureBoot, BIOS, FirmwareType, TPM, Encryption, DeviceGuard) for all Intune managed devices.

.DESCRIPTION
    Connects to Microsoft Graph Beta and retrieves:
    - Hardware security fields from hardwareInformation (TPM, encryption, DeviceGuard)
    - Recurring local observations from a detection-only Intune Remediations package
      via deviceHealthScripts/deviceRunStates, or legacy Platform Script snapshots until configured

    It:
    - Retrieves all managed devices with hardwareInformation in bulk
    - Correlates SecureBoot stdout from the Platform Script deviceRunStates
    - Compares registered third-party antivirus names with the tenant's ExpectedThirdPartyAV configuration
    - Exports results to CSV via ExportAndCopyCsv
    - Uses the shared framework (SmartM365.Core / InitializeScriptEnvironment)
    - Logs to text + transcript
    - Cleans old CSV/log files

.PARAMETER OutputPath
    Specifies the output directory for the generated CSV and log files.

.PARAMETER Connect
    Forces a (re)connection to Microsoft Graph (disconnects any existing session first).

.PARAMETER InteractiveAuth
    Uses interactive authentication instead of app-only certificate authentication.

.PARAMETER PlatformScriptId
    Legacy Platform Script GUID used only until a recurring DeviceHealthScriptId is configured.

.PARAMETER DeviceHealthScriptId
    Intune Remediations detection-only package GUID. When configured, its recurring
    deviceRunStates are the sole source of locally observed device system data.

.PARAMETER RunStatePageSize
    Number of deviceRunStates requested per Graph page. Default: 500. Allowed range: 1-1000.

.PARAMETER RunStatePagePauseMilliseconds
    Delay between Graph pages in milliseconds. Default: 500. Retry/backoff still handles throttling.
.PARAMETER ManagedDevicePageSize
    Number of managed devices requested per Graph page. Default: 500. A smaller page limits throttle pressure.
.PARAMETER ManagedDevicePagePauseMilliseconds
    Delay between managed-device Graph pages in milliseconds. Default: 500.
.VERSION
2.19
.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication.
    Minimum Graph application permissions: DeviceManagementManagedDevices.Read.All; DeviceManagementConfiguration.Read.All; DeviceManagementScripts.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Version : 2.17
    Author: https://github.com/khda79/workplacecloudhub.com
    Requires: SmartM365.Core module (logging, init, CSV, cleanup, cloud connectivity)
    Scopes: DeviceManagementManagedDevices.Read.All, DeviceManagementConfiguration.Read.All
    Minimum application permissions: DeviceManagementManagedDevices.Read.All, DeviceManagementConfiguration.Read.All, DeviceManagementScripts.Read.All
#>

param(
    [string]$Tenant = 'test',
[string]$OutputPath,
    [switch]$Connect,
    [switch]$InteractiveAuth,

    [Parameter(Mandatory = $false)]
    [string]$PlatformScriptId = "0d121c5c-65cc-480f-b07a-9ae79d2d928d",

    [string]$DeviceHealthScriptId = '',

    [ValidateRange(1, 1000)]
    [int]$RunStatePageSize = 500,

    [ValidateRange(0, 5000)]
    [int]$RunStatePagePauseMilliseconds = 500,

    [ValidateRange(1, 1000)]
    [int]$ManagedDevicePageSize = 500,

    [ValidateRange(0, 5000)]
    [int]$ManagedDevicePagePauseMilliseconds = 500,

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
Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot | Out-Null

function Get-ScriptLocalConfig {
    [CmdletBinding()]
    param()

    $configPath = Join-Path -Path $PSScriptRoot -ChildPath ("{0}.local.json" -f [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)); $configPath = Resolve-SmartM365JsonConfigurationPath -Path $configPath
    if (-not (Test-Path -LiteralPath $configPath)) {
        $templatePath = (Get-SmartM365JsonTemplateName -Path $configPath)
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

            Write-SmartM365JsonBytesAtomically -Path $configPath -Bytes ([IO.File]::ReadAllBytes($templatePath)) -ExpectedSHA256 'ABSENT' -Validate {param($document) if($document -isnot [pscustomobject]){throw 'Configuration template must be an object.'}} | Out-Null
            Write-Host ("Created script local configuration from template: {0}" -f $configPath) -ForegroundColor Yellow
            Write-Host 'Review the generated local JSON values; continuing with current file values.' -ForegroundColor Yellow
        }
    }

    try {
        $config = Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        return Sync-SmartM365JsonConfigWithTemplate -Config $config -Path $configPath
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
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json'; $globalConfigPath = Resolve-SmartM365JsonConfigurationPath -Path $globalConfigPath
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
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json'; $globalConfigPath = Resolve-SmartM365JsonConfigurationPath -Path $globalConfigPath
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
$DeviceHealthScriptId = if ($PSBoundParameters.ContainsKey('DeviceHealthScriptId')) { $DeviceHealthScriptId } else {
    [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DeviceHealthScriptId' -DefaultValue '')
}
if ($DeviceHealthScriptId -and $DeviceHealthScriptId -notmatch '^[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}$') {
    throw 'DeviceHealthScriptId must be an Intune script GUID.'
}
$UseRecurringDeviceSystem = -not [string]::IsNullOrWhiteSpace($DeviceHealthScriptId)



$global:RetentionMaxCSV = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxLogs' -DefaultValue 30)

$global:EnableSharePointUpload = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointTargetFolderPath' -DefaultValue ''
# ==========================================================
# PowerShell 7 minimum
# ==========================================================
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=(1); exit $smartM365RunGuard.ExitCode
}

# Avoid PS function-capacity issues
$MaximumFunctionCount = 32768

# ==========================================================
# App-only authentication parameters (same app as other scripts)
# ==========================================================
function Get-ScriptLocalConfig {
    [CmdletBinding()]
    param()

    $configPath = Join-Path -Path $PSScriptRoot -ChildPath ("{0}.local.json" -f [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)); $configPath = Resolve-SmartM365JsonConfigurationPath -Path $configPath
    if (-not (Test-Path -LiteralPath $configPath)) {
        $templatePath = (Get-SmartM365JsonTemplateName -Path $configPath)
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

            Write-SmartM365JsonBytesAtomically -Path $configPath -Bytes ([IO.File]::ReadAllBytes($templatePath)) -ExpectedSHA256 'ABSENT' -Validate {param($document) if($document -isnot [pscustomobject]){throw 'Configuration template must be an object.'}} | Out-Null
            Write-Host ("Created script local configuration from template: {0}" -f $configPath) -ForegroundColor Yellow
            Write-Host 'Review the generated local JSON values; continuing with current file values.' -ForegroundColor Yellow
        }
    }

    try {
        $config = Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
        return Sync-SmartM365JsonConfigWithTemplate -Config $config -Path $configPath
    }
    catch {
        throw ("Failed to read local configuration '{0}': {1}" -f $configPath, $_.Exception.Message)
    }
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
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json'; $globalConfigPath = Resolve-SmartM365JsonConfigurationPath -Path $globalConfigPath
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

function Resolve-ExpectedThirdPartyAVConfiguration {
    [CmdletBinding()]
    param([AllowNull()][object]$Value)

    $notConfigured = [pscustomobject]@{ DisplayName = ''; RegisteredNames = @(); IsConfigured = $false }
    if ($null -eq $Value) { return $notConfigured }
    if ($Value -isnot [pscustomobject] -and $Value -isnot [System.Collections.IDictionary]) {
        throw 'ExpectedThirdPartyAV must be an object with DisplayName and RegisteredNames.'
    }

    $displayName = ([string]$Value.DisplayName).Trim()
    $rawNames = $Value.RegisteredNames
    if ($null -ne $rawNames -and $rawNames -isnot [array]) {
        throw 'ExpectedThirdPartyAV.RegisteredNames must be a JSON array.'
    }

    $names = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($rawName in @($rawNames)) {
        $name = ([string]$rawName).Trim()
        if (-not $name) { throw 'ExpectedThirdPartyAV.RegisteredNames cannot contain empty names.' }
        [void]$names.Add($name)
    }

    if (-not $displayName -and $names.Count -eq 0) { return $notConfigured }
    if (-not $displayName -or $names.Count -eq 0) {
        throw 'ExpectedThirdPartyAV requires both DisplayName and at least one RegisteredNames value.'
    }

    return [pscustomobject]@{
        DisplayName = $displayName
        RegisteredNames = @($names)
        IsConfigured = $true
    }
}

function Get-ExpectedThirdPartyAVStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Configuration,
        [Parameter(Mandatory)][object]$PlatformScriptData
    )

    if (-not $Configuration.IsConfigured) { return 'NotConfigured' }
    switch ([string]$PlatformScriptData.ThirdPartyAVStatus) {
        'NotCollected' { return 'NotCollected' }
        'Unknown' { return 'Unknown' }
        'Error' { return 'Error' }
        'NotDetected' { return 'NotDetected' }
        'Detected' {
            if ([string]::IsNullOrWhiteSpace([string]$PlatformScriptData.ThirdPartyAVProducts)) {
                return 'Unknown'
            }
            foreach ($product in ([string]$PlatformScriptData.ThirdPartyAVProducts).Split(';')) {
                if ($Configuration.RegisteredNames -icontains $product.Trim()) { return 'Detected' }
            }
            return 'NotDetected'
        }
        default { return 'Unknown' }
    }
}

$ExpectedThirdPartyAV = Resolve-ExpectedThirdPartyAVConfiguration -Value (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ExpectedThirdPartyAV' -DefaultValue $null)

# ==========================================================
# Import SmartM365.Core module (psd1)
# ==========================================================
$modulePath = & { $d = $PSScriptRoot; while ($d) { $p = Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'; if (Test-Path -LiteralPath $p) { return $p }; $parent = Split-Path -Path $d -Parent; if ($parent -eq $d) { break }; $d = $parent }; throw 'SmartM365.Core module not found.' }
try {
    Import-Module -Name $modulePath -MinimumVersion '1.0.119' -ErrorAction Stop
} catch {
    Write-Host "Failed to import SmartM365.Core module from '$modulePath' : $_" -ForegroundColor Red
    $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=(1); exit $smartM365RunGuard.ExitCode
}

# ==========================================================
# Configuration
# ==========================================================

# Throttle-aware retry settings
$MaxRetries        = 8
$BaseDelaySeconds  = 2

# Graph API settings - hardwareInformation fields that are actually populated
$BulkEndpoint = "/beta/deviceManagement/managedDevices?`$select=id,deviceName,azureADDeviceId,userPrincipalName,managementState,complianceState,lastSyncDateTime,hardwareInformation&`$top=$ManagedDevicePageSize"

# ==========================================================
# Helpers
# ==========================================================

function ConvertTo-DeviceSystemGraphUtcTimestamp {
    [CmdletBinding()]
    param(
        [AllowNull()][AllowEmptyString()][object]$Value,
        [Parameter(Mandatory)][string]$FieldName
    )

    if ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value))) { return $null }

    if ($Value -is [datetimeoffset]) {
        return $Value.UtcDateTime.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
    }
    if ($Value -is [datetime]) {
        if ($Value.Kind -ne [DateTimeKind]::Unspecified) {
            return $Value.ToUniversalTime().ToString('o', [Globalization.CultureInfo]::InvariantCulture)
        }
    }
    elseif ($Value -is [string]) {
        # Only explicit Graph offsets are accepted; never infer the server/device time zone.
        $timestamp = $Value.Trim()
        $parsed = [datetimeoffset]::MinValue
        if ($timestamp -cmatch '^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{1,7})?(?:Z|[+-]\d{2}:\d{2})$' -and
            [datetimeoffset]::TryParse($timestamp, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$parsed)) {
            return $parsed.UtcDateTime.ToString('o', [Globalization.CultureInfo]::InvariantCulture)
        }
    }

    [void](WriteLog -Message "$FieldName is not a qualified Graph timestamp: a valid explicit offset or known DateTime Kind is required. Exporting an empty value; no time-zone fallback was applied." -Level 'WARNING')
    return $null
}

function Invoke-GraphSafe {
    <#
    .SYNOPSIS
        Throttle-aware wrapper around Invoke-MgGraphRequest with exponential back-off.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Method = 'GET',
        [int]$MaxRetries = 8,
        [int]$BaseDelaySeconds = 2
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $response = Invoke-MgGraphRequest -Uri $Uri -Method $Method -OutputType PSObject -ErrorAction Stop
            return $response
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }
            if ($null -eq $statusCode -and $_.Exception.Data -and $_.Exception.Data.Contains('StatusCode')) {
                $rawStatusCode = $_.Exception.Data['StatusCode']
                $parsedStatusCode = 0
                if ($null -ne $rawStatusCode -and [int]::TryParse([string]$rawStatusCode, [ref]$parsedStatusCode)) {
                    $statusCode = $parsedStatusCode
                }
            }

            # Fallback: detect SDK-wrapped 429 when Response is null (SDK exhausted its own retries)
            $isSdkWrapped429 = ($null -eq $statusCode) -and ($_.Exception.Message -match 'TooManyRequests|too many retries|429')

            # Throttled or transient HTTP failure
            if (($statusCode -in @(408, 409, 429, 500, 502, 503, 504) -or $isSdkWrapped429) -and $attempt -le $MaxRetries) {
                $retryAfter = $BaseDelaySeconds * [math]::Pow(2, $attempt - 1)

                # SDK-wrapped throttle: use a fixed 60s floor - the SDK already burned its own retries
                if ($isSdkWrapped429) {
                    $retryAfter = [math]::Max($retryAfter, 60)
                }
                # Check Retry-After header (only available on raw responses)
                else {
                    $retryAfterHeader = $null
                    try { $retryAfterHeader = @($_.Exception.Response.Headers.GetValues('Retry-After') | Select-Object -First 1)[0] } catch {}
                    if ($null -eq $retryAfterHeader) { try { $retryAfterHeader = $_.Exception.Data['Retry-After'] } catch {} }
                    $retryAfterSeconds = 0
                    if ($null -ne $retryAfterHeader -and [int]::TryParse([string]$retryAfterHeader, [ref]$retryAfterSeconds) -and $retryAfterSeconds -gt 0) {
                        $retryAfter = $retryAfterSeconds
                    }
                }
                $retryAfter = [Math]::Min(300, [Math]::Max(1, [int]$retryAfter))

                WriteLog -Message ("HTTP {0} on attempt {1}/{2}. Retrying in {3}s... URI: {4}" -f $(if ($isSdkWrapped429) { '429(sdk)' } else { $statusCode }), $attempt, $MaxRetries, $retryAfter, $Uri) "WARNING"
                Start-Sleep -Seconds $retryAfter
            }
            else {
                throw
            }
        }
    }
}

function Test-GraphResponseProperty {
    param([AllowNull()][object]$Response,[Parameter(Mandatory)][string]$Name)
    if($null-eq$Response){return $false}
    if($Response -is [System.Collections.IDictionary]){return $Response.Contains($Name)}
    return $null -ne $Response.PSObject.Properties[$Name]
}
function Get-GraphResponsePropertyValue {
    param([AllowNull()][object]$Response,[Parameter(Mandatory)][string]$Name)
    if($null-eq$Response){return $null}
    if($Response -is [System.Collections.IDictionary]){return $Response[$Name]}
    $property=$Response.PSObject.Properties[$Name]
    if($property){return $property.Value}
    return $null
}

function Get-PlatformScriptRunStates {
    <#
    .SYNOPSIS
        Retrieves deviceRunStates for a given Platform Script and builds a lookup hashtable
        keyed by managed device ID, with the stdout result as value.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ScriptId
    )

    WriteLog -Message ("Retrieving Platform Script deviceRunStates for script {0}. RequestedPageSize={1}; PagePauseMs={2}" -f $ScriptId, $RunStatePageSize, $RunStatePagePauseMilliseconds) "INFO"

    $map = @{}
    $pageCount = 0
    $visitedRunStateUris = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $uri = "/beta/deviceManagement/deviceManagementScripts/$ScriptId/deviceRunStates?`$expand=managedDevice(`$select=id)&`$select=resultMessage,lastStateUpdateDateTime&`$top=$RunStatePageSize"

    while ($uri) {
        if (-not $visitedRunStateUris.Add([string]$uri)) { throw 'Platform Script run-state pagination returned a repeated @odata.nextLink; collection is incomplete.' }
        $response = Invoke-GraphSafe -Uri $uri -MaxRetries $MaxRetries -BaseDelaySeconds $BaseDelaySeconds
        $pageCount++
        if (-not (Test-GraphResponseProperty -Response $response -Name 'value')) { throw "Platform Script run-state page $pageCount returned an invalid Graph collection response without a value property." }

        $runStates = Get-GraphResponsePropertyValue -Response $response -Name 'value'
        if ($runStates) {
            foreach ($runState in $runStates) {
                $deviceId = $null
                if ($runState.managedDevice -and $runState.managedDevice.id) {
                    $deviceId = $runState.managedDevice.id
                }

                if (-not [string]::IsNullOrWhiteSpace($deviceId) -and -not $map.ContainsKey($deviceId)) {
                    $map[$deviceId] = [PSCustomObject]@{
                        ResultMessage            = $runState.resultMessage
                        LastStateUpdateDateTime  = $runState.lastStateUpdateDateTime
                    }
                }
            }
        }

        $uri = Get-GraphResponsePropertyValue -Response $response -Name '@odata.nextLink'
        if ($uri) {
            if (($pageCount % 25) -eq 0) {
                WriteLog -Message ("Paging deviceRunStates... Pages={0}; devices={1}" -f $pageCount, $map.Count) "INFO"
            }
            if ($RunStatePagePauseMilliseconds -gt 0) {
                Start-Sleep -Milliseconds $RunStatePagePauseMilliseconds
            }
        }
    }

    WriteLog -Message ("Platform Script deviceRunStates retrieved: {0} devices in {1} page(s), requested page size {2}" -f $map.Count, $pageCount, $RunStatePageSize) "INFO"
    return $map
}

function Get-DeviceHealthScriptRunStates {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$ScriptId)

    WriteLog -Message ("Retrieving recurring Device System deviceRunStates. RequestedPageSize={0}; PagePauseMs={1}" -f $RunStatePageSize, $RunStatePagePauseMilliseconds) "INFO"
    $map = @{}
    $duplicateCount = 0
    $visited = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $uri = "/beta/deviceManagement/deviceHealthScripts/$ScriptId/deviceRunStates?`$expand=managedDevice(`$select=id)&`$select=detectionState,preRemediationDetectionScriptOutput,lastStateUpdateDateTime&`$top=$RunStatePageSize"
    $pageCount = 0
    while ($uri) {
        if (-not $visited.Add([string]$uri)) { throw 'Device Health Script pagination returned a repeated @odata.nextLink; collection is incomplete.' }
        $response = Invoke-GraphSafe -Uri $uri -MaxRetries $MaxRetries -BaseDelaySeconds $BaseDelaySeconds
        $pageCount++
        if (-not (Test-GraphResponseProperty -Response $response -Name 'value')) { throw "Device Health Script page $pageCount returned an invalid Graph collection response without a value property." }
        foreach ($runState in @(Get-GraphResponsePropertyValue -Response $response -Name 'value')) {
            if ($null -eq $runState) { continue }
            $deviceId = [string]$runState.managedDevice.id
            if ([string]::IsNullOrWhiteSpace($deviceId)) { throw 'Device Health Script run state lacks a managed device ID; collection is incomplete.' }
            $candidate = [pscustomobject]@{
                ResultMessage = $runState.preRemediationDetectionScriptOutput
                LastStateUpdateDateTime = $runState.lastStateUpdateDateTime
                DetectionState = [string]$runState.detectionState
            }
            if ($map.ContainsKey($deviceId)) {
                $duplicateCount++
                $existing = $map[$deviceId]
                $existingUtc = ConvertTo-DeviceSystemGraphUtcTimestamp -Value $existing.LastStateUpdateDateTime -FieldName 'DuplicateRunStateTime'
                $candidateUtc = ConvertTo-DeviceSystemGraphUtcTimestamp -Value $candidate.LastStateUpdateDateTime -FieldName 'DuplicateRunStateTime'
                if (-not $existingUtc -or -not $candidateUtc) {
                    throw 'Duplicate Device Health Script run states have a missing or invalid update timestamp; collection is ambiguous.'
                }
                $existingTime = [datetimeoffset]::Parse($existingUtc, [Globalization.CultureInfo]::InvariantCulture)
                $candidateTime = [datetimeoffset]::Parse($candidateUtc, [Globalization.CultureInfo]::InvariantCulture)
                if ($candidateTime -lt $existingTime) { continue }
                if ($candidateTime -eq $existingTime) {
                    if ($candidate.ResultMessage -cne $existing.ResultMessage -or $candidate.DetectionState -cne $existing.DetectionState) {
                        throw 'Duplicate Device Health Script run states conflict at the same update timestamp; collection is ambiguous.'
                    }
                    continue
                }
            }
            $map[$deviceId] = $candidate
        }
        $uri = Get-GraphResponsePropertyValue -Response $response -Name '@odata.nextLink'
        if ($uri -and $RunStatePagePauseMilliseconds -gt 0) { Start-Sleep -Milliseconds $RunStatePagePauseMilliseconds }
    }
    if ($map.Count -eq 0) { throw 'Device Health Script returned no device run states; the existing CSV was not replaced.' }
    WriteLog -Message ("Recurring Device System run states retrieved: {0} devices in {1} page(s); duplicate rows collapsed: {2}." -f $map.Count, $pageCount, $duplicateCount) "INFO"
    return $map
}

function Parse-PlatformScriptStdout {
    <#
    .SYNOPSIS
        Parses the pipe-delimited stdout from the SmartM365-Detect-DeviceSystemInfo device probe.
        The recurring source also has local console logon fields and a payload end marker.
        Returns a hashtable with each key/value pair.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$ResultMessage
    )

    $result = @{
        SecureBoot     = $null
        BIOSVersion    = $null
        BIOSDate       = $null
        FirmwareType   = $null
        LastBootUpTime = $null
        ThirdPartyAVStatus   = 'NotCollected'
        ThirdPartyAVProducts = $null
        WindowsUpdateRebootRequired = 'NotCollected'
        BitLockerOSProtectionStatus = 'NotCollected'
        BitLockerOSConversionStatus = 'NotCollected'
        LastConsoleLogonStatus = 'NotCollected'
        LastConsoleLogonUser = $null
        LastConsoleLogonSid = $null
        LastConsoleLogonUtc = $null
    }

    if ([string]::IsNullOrWhiteSpace($ResultMessage)) {
        return $result
    }

    $parts = $ResultMessage.Trim().Split('|')
    foreach ($part in $parts) {
        if ($part -match '^([^:]+):(.*)$') {
            $key   = $Matches[1].Trim()
            $value = $Matches[2].Trim()
            if ($result.ContainsKey($key)) {
                $result[$key] = $value
            }
        }
    }

    return $result
}

function Test-DeviceSystemRecurringPayload {
    [CmdletBinding()]
    param([AllowNull()][string]$ResultMessage)
    if ([string]::IsNullOrWhiteSpace($ResultMessage)) { return $false }
    $text = $ResultMessage.Trim()
    $hasSid = $text.EndsWith('|DeviceSystemPayloadEnd:1.6', [StringComparison]::Ordinal)
    if (-not $hasSid -and -not $text.EndsWith('|DeviceSystemPayloadEnd:1.5', [StringComparison]::Ordinal)) { return $false }
    if ($hasSid -and $text -notmatch '(?:^|\|)LastConsoleLogonSid:') { return $false }
    foreach ($field in @('SecureBoot','BIOSVersion','BIOSDate','FirmwareType','LastBootUpTime',
        'ThirdPartyAVStatus','ThirdPartyAVProducts','WindowsUpdateRebootRequired',
        'BitLockerOSProtectionStatus','BitLockerOSConversionStatus','LastConsoleLogonStatus',
        'LastConsoleLogonUser','LastConsoleLogonUtc')) {
        if ($text -notmatch "(?:^|\|)$($field):") { return $false }
    }
    return $true
}

function Get-DeviceSystemLocalFreshness {
    [CmdletBinding()]
    param([AllowNull()][string]$ObservationUtc, [bool]$IsRecurring, [bool]$HasOutput,
          [AllowNull()][string]$DetectionState, [bool]$PayloadValid = $true, [int]$MaxAgeHours = 192,
          [datetimeoffset]$AsOfUtc = [datetimeoffset]::UtcNow)
    if (-not $IsRecurring) { return 'LegacySnapshot' }
    if ([string]::IsNullOrWhiteSpace($DetectionState)) {
        if (-not $HasOutput) { return 'NotCollected' }
        return 'Unknown'
    }
    # Pending and non-applicable runs are not technical script failures.
    switch ($DetectionState.Trim().ToLowerInvariant()) {
        'success' { }
        'pending' { return 'Pending' }
        'notapplicable' { return 'NotApplicable' }
        'fail' { return 'Failed' }
        'scripterror' { return 'ScriptError' }
        default { return 'Unknown' }
    }
    if (-not $HasOutput) { return 'NotCollected' }
    if (-not $PayloadValid) { return 'InvalidOutput' }
    if (-not $ObservationUtc) { return 'InvalidTimestamp' }
    $observed = [datetimeoffset]::MinValue
    if (-not [datetimeoffset]::TryParse($ObservationUtc, [Globalization.CultureInfo]::InvariantCulture,
        [Globalization.DateTimeStyles]::None, [ref]$observed)) { return 'InvalidTimestamp' }
    $age = ($AsOfUtc - $observed).TotalHours
    if ($age -lt (-5.0 / 60.0)) { return 'FutureObservation' }
    if ($age -gt $MaxAgeHours) { return 'Stale' }
    return 'Fresh'
}

# ==========================================================
# Initialization via SmartM365.Core
# ==========================================================
$ScriptVersion = "2.19"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DeviceSystemCsvLogFolderPath' -DefaultValue $OutputPath
try {
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPathInit $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
    . (Join-Path (Split-Path $modulePath -Parent) 'SmartM365-SharePointUploadPolicy.ps1')
    Initialize-SmartM365CollectorUploadContext -LocalConfig $ScriptLocalConfig -GlobalConfig $global:SmartM365GlobalConfig -DataAllRoot $InitializeOutputPath -DataLastRoot (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') -LogAllRoot $global:LogPath -ExternalActionsAllowed $true
    Start-SmartM365SourceReceipt -ScriptPath $PSCommandPath -SourceRootPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') -ScopeParameters @{PlatformScriptId=$PlatformScriptId;DeviceHealthScriptId=$DeviceHealthScriptId}
    Start-SmartM365RunTranscript -Path $global:logTranscriptFile -Append

    WriteLog -Message "Script Environment initialized at $InitializeOutputPath"
    $OutputPath = $InitializeOutputPath
    WriteLog -Message "Starting $TaskName..."
} catch {
    Write-Host "Initialization failed: $_" -ForegroundColor Red
    $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=(1); exit $smartM365RunGuard.ExitCode
}

# ==========================================================
# MAIN TRY / CATCH / FINALLY
# ==========================================================
$connectedGraphInThisRun = $false

try {

    # ------------------------
    # Connect to Microsoft Graph
    # ------------------------
    function Test-GraphConnection {
        try {
            $org = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -ErrorAction Stop
            return $true
        } catch {
            return $false
        }
    }

    $graphContext = $null
    if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
        try {
            $graphContext = Get-MgContext -ErrorAction SilentlyContinue
        } catch { }
    }

    $needConnect = $false

    if ($Connect) {
        Write-Host "Connect switch specified: existing Graph session (if any) will be disconnected and reconnected..." -ForegroundColor Cyan
        if ($graphContext) {
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null; Set-SmartM365InventoryConnectionOwnership -Connection Graph -Owned $false } catch { }
        }
        $needConnect = $true
    } else {
        if ($graphContext -and (Test-GraphConnection)) {
            Write-Host "Existing Microsoft Graph session detected. Reusing current connection." -ForegroundColor Cyan
            $needConnect = $false
        } else {
            Write-Host "No existing Graph session detected. Will establish a new connection..." -ForegroundColor Cyan
            $needConnect = $true
        }
    }

    if ($needConnect) {
        if (-not $InteractiveAuth) {
            WriteLog -Message "Connecting to Microsoft Graph with app-only certificate authentication (direct Connect-MgGraph)." "INFO"
            Connect-MgGraph -TenantId $TenantId -ClientId $AppId -CertificateThumbprint $Thumb -NoWelcome | Out-Null; Set-SmartM365InventoryConnectionOwnership -Connection Graph -Owned $true
        } else {
            WriteLog -Message "Connecting to Microsoft Graph with interactive authentication (direct Connect-MgGraph)." "INFO"
            Connect-MgGraph -Scopes @("DeviceManagementManagedDevices.Read.All","DeviceManagementConfiguration.Read.All") -NoWelcome | Out-Null; Set-SmartM365InventoryConnectionOwnership -Connection Graph -Owned $true
        }

        if (-not (Test-GraphConnection)) {
            throw "Failed to connect to Microsoft Graph."
        }

        $connectedGraphInThisRun = $true
    }

    # ======================================================
    $runStateProbe = if ($UseRecurringDeviceSystem) {
        "https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts/$DeviceHealthScriptId/deviceRunStates?`$top=1"
    } else {
        'https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts?$top=1'
    }
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -RequiredGraphApplicationPermissions @('DeviceManagementManagedDevices.Read.All','DeviceManagementConfiguration.Read.All','DeviceManagementScripts.Read.All') -GraphProbeUris @(
        'https://graph.microsoft.com/beta/deviceManagement/managedDevices?$top=1', $runStateProbe
    ) | Out-Null

    # PHASE 1: BULK RETRIEVAL OF MANAGED DEVICES
    # ======================================================
    WriteLog -Message "Phase 1: Bulk retrieval of managed devices with hardwareInformation..." "INFO"

    $allDevices = [System.Collections.Generic.List[PSObject]]::new()
    $uri = $BulkEndpoint
    $visitedDeviceUris = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $devicePageNumber = 0

    while ($uri) {
        if (-not $visitedDeviceUris.Add([string]$uri)) { throw 'Managed-device pagination returned a repeated @odata.nextLink; collection is incomplete.' }
        $devicePageNumber++
        $response = Invoke-GraphSafe -Uri $uri -MaxRetries $MaxRetries -BaseDelaySeconds $BaseDelaySeconds
        if (-not (Test-GraphResponseProperty -Response $response -Name 'value')) { throw "Managed-device page $devicePageNumber returned an invalid Graph collection response without a value property." }
        $managedDevices = Get-GraphResponsePropertyValue -Response $response -Name 'value'
        if ($managedDevices) {
            $allDevices.AddRange([PSObject[]]$managedDevices)
            if ($MaxItems -gt 0 -and $allDevices.Count -ge $MaxItems) {
                while ($allDevices.Count -gt $MaxItems) { $allDevices.RemoveAt($allDevices.Count - 1) }
                WriteLog -Message ("MaxItems enabled: restricted managed devices to {0}." -f $allDevices.Count) "WARNING"
                break
            }
        }
        $uri = Get-GraphResponsePropertyValue -Response $response -Name '@odata.nextLink'
        if ($uri) {
            WriteLog -Message ("Paging... Total devices so far: {0}" -f $allDevices.Count) "INFO"
            if ($ManagedDevicePagePauseMilliseconds -gt 0) { Start-Sleep -Milliseconds $ManagedDevicePagePauseMilliseconds }
        }
    }

    WriteLog -Message ("Bulk retrieval complete. Total devices: {0}" -f $allDevices.Count) "INFO"

    if ($allDevices.Count -eq 0) {
        WriteLog -Message "No managed devices found. Exiting." "WARNING"
        $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
    }

    # ======================================================
    # PHASE 2: RETRIEVE LOCAL DEVICE OBSERVATIONS
    # ======================================================
    $secureBootMap = @{}

    if ($UseRecurringDeviceSystem) {
        WriteLog -Message 'Phase 2: Retrieving recurring Device System detection results...' 'INFO'
        $secureBootMap = Get-DeviceHealthScriptRunStates -ScriptId $DeviceHealthScriptId
    }
    elseif (-not [string]::IsNullOrWhiteSpace($PlatformScriptId)) {
        WriteLog -Message "Phase 2: Retrieving legacy Platform Script deviceRunStates..." "WARNING"
        $secureBootMap = Get-PlatformScriptRunStates -ScriptId $PlatformScriptId
    }
    else {
        WriteLog -Message "Phase 2: Skipped (no DeviceHealthScriptId or PlatformScriptId configured). Local observation columns will be empty." "WARNING"
    }

    # ======================================================
    # BUILD RESULTS
    # ======================================================
    WriteLog -Message "Building results..." "INFO"

    $results = [System.Collections.Generic.List[PSObject]]::new()

    foreach ($device in $allDevices) {
        $hw = $device.hardwareInformation

        # Locally observed system information, antivirus registration, and console logon.
        $emptyPsData = @{ SecureBoot = $null; BIOSVersion = $null; BIOSDate = $null; FirmwareType = $null; LastBootUpTime = $null; ThirdPartyAVStatus = 'NotCollected'; ThirdPartyAVProducts = $null; WindowsUpdateRebootRequired = 'NotCollected'; BitLockerOSProtectionStatus = 'NotCollected'; BitLockerOSConversionStatus = 'NotCollected'; LastConsoleLogonStatus = 'NotCollected'; LastConsoleLogonUser = $null; LastConsoleLogonSid = $null; LastConsoleLogonUtc = $null }
        $psData = $emptyPsData
        $psLastUpdate = $null
        $detectionState = $null
        $hasLocalOutput = $false
        $payloadValid = $false
        if ($secureBootMap.ContainsKey($device.id)) {
            $runState = $secureBootMap[$device.id]
            $detectionState = [string]$runState.DetectionState
            $hasLocalOutput = -not [string]::IsNullOrWhiteSpace([string]$runState.ResultMessage)
            $payloadValid = if ($UseRecurringDeviceSystem) { Test-DeviceSystemRecurringPayload -ResultMessage $runState.ResultMessage } else { $hasLocalOutput }
            if ((-not $UseRecurringDeviceSystem -or $detectionState -ieq 'success') -and $payloadValid) {
                $psData = Parse-PlatformScriptStdout -ResultMessage $runState.ResultMessage
            }
            $psLastUpdate = $runState.LastStateUpdateDateTime
        }

        $localObservationUtc = ConvertTo-DeviceSystemGraphUtcTimestamp -Value $psLastUpdate -FieldName 'LocalObservationUtc'
        $localFreshness = Get-DeviceSystemLocalFreshness -ObservationUtc $localObservationUtc -IsRecurring $UseRecurringDeviceSystem -HasOutput $hasLocalOutput -DetectionState $detectionState -PayloadValid $payloadValid
        if ($UseRecurringDeviceSystem -and $localFreshness -ne 'Fresh') {
            $psData = $emptyPsData
        }

        $record = [PSCustomObject]@{
            DeviceName                        = $device.deviceName
            DeviceId                          = $device.id
            AzureADDeviceId                   = $device.azureADDeviceId
            UserPrincipalName                 = $device.userPrincipalName
            SecureBootStatus                  = $psData.SecureBoot
            BIOSVersion                       = $psData.BIOSVersion
            BIOSDate                          = $psData.BIOSDate
            FirmwareType                      = $psData.FirmwareType
            PlatformScriptLastUpdate          = $localObservationUtc # Compatibility alias for existing prepared evidence.
            LocalObservationUtc               = $localObservationUtc
            LocalDataSource                   = if ($UseRecurringDeviceSystem) { 'IntuneRemediation' } else { 'LegacyPlatformScript' }
            LocalDataFreshness                = $localFreshness
            LastRebootDate                    = $psData.LastBootUpTime
            'Last Reboot Date'                = $psData.LastBootUpTime
            LastBootUpTime                    = $psData.LastBootUpTime
            LastBootTime                      = $psData.LastBootUpTime
            IsEncrypted                       = if ($hw) { $hw.isEncrypted } else { $null }
            DeviceGuardVBSState               = if ($hw) { $hw.deviceGuardVirtualizationBasedSecurityState } else { $null }
            DeviceGuardVBSHardwareRequirement = if ($hw) { $hw.deviceGuardVirtualizationBasedSecurityHardwareRequirementState } else { $null }
            DeviceGuardCredentialGuardState   = if ($hw) { $hw.deviceGuardLocalSystemAuthorityCredentialGuardState } else { $null }
            ManagementState                   = $device.managementState
            ComplianceState                   = $device.complianceState
            LastSyncDateTime                  = ConvertTo-DeviceSystemGraphUtcTimestamp -Value $device.lastSyncDateTime -FieldName 'LastSyncDateTime'
            ThirdPartyAVStatus                = $psData.ThirdPartyAVStatus
            ThirdPartyAVProducts              = $psData.ThirdPartyAVProducts
            ExpectedThirdPartyAVName          = $ExpectedThirdPartyAV.DisplayName
            ExpectedThirdPartyAVStatus        = Get-ExpectedThirdPartyAVStatus -Configuration $ExpectedThirdPartyAV -PlatformScriptData $psData
            WindowsUpdateRebootRequired       = $psData.WindowsUpdateRebootRequired
            BitLockerOSProtectionStatus       = $psData.BitLockerOSProtectionStatus
            BitLockerOSConversionStatus       = $psData.BitLockerOSConversionStatus
            LastConsoleLogonStatus            = $psData.LastConsoleLogonStatus
            LastConsoleLogonUser              = $psData.LastConsoleLogonUser
            LastConsoleLogonSid               = $psData.LastConsoleLogonSid
            LastConsoleLogonUtc               = ConvertTo-DeviceSystemGraphUtcTimestamp -Value $psData.LastConsoleLogonUtc -FieldName 'LastConsoleLogonUtc'
        }

        $results.Add($record)
    }

    # ======================================================
    # If no devices, exit gracefully
    # ======================================================
    if (-not $results -or $results.Count -eq 0) {
        WriteLog -Message "No devices found. Exiting." "WARNING"
        $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
    }

    # ======================================================
    # SUMMARY (before filtering)
    # ======================================================
    $sbEnabled      = ($results | Where-Object { $_.SecureBootStatus -eq 'Enabled' }).Count
    $sbDisabled     = ($results | Where-Object { $_.SecureBootStatus -eq 'Disabled' }).Count
    $sbNotSupported = ($results | Where-Object { $_.SecureBootStatus -eq 'NotSupported' }).Count
    $sbEmpty        = ($results | Where-Object { [string]::IsNullOrWhiteSpace($_.SecureBootStatus) }).Count
    $sbError        = ($results | Where-Object { $_.SecureBootStatus -eq 'Error' }).Count
    $fwUefi         = ($results | Where-Object { $_.FirmwareType -eq 'Uefi' }).Count
    $fwLegacy       = ($results | Where-Object { $_.FirmwareType -eq 'Bios' }).Count

    WriteLog -Message "=== SUMMARY ===" "INFO"
    WriteLog -Message ("Total devices           : {0}" -f $results.Count) "INFO"
    WriteLog -Message ("SecureBoot Enabled       : {0}" -f $sbEnabled) "INFO"
    WriteLog -Message ("SecureBoot Disabled      : {0}" -f $sbDisabled) "INFO"
    WriteLog -Message ("SecureBoot NotSupported  : {0}" -f $sbNotSupported) "INFO"
    WriteLog -Message ("SecureBoot Error         : {0}" -f $sbError) "INFO"
    WriteLog -Message ("SecureBoot Unknown/Empty : {0}" -f $sbEmpty) "INFO"
    WriteLog -Message ("Firmware UEFI            : {0}" -f $fwUefi) "INFO"
    WriteLog -Message ("Firmware Legacy BIOS     : {0}" -f $fwLegacy) "INFO"

    # ======================================================
    # Recurring mode retains all managed devices and marks unavailable local evidence.
    # ======================================================
    $filteredResults = if ($UseRecurringDeviceSystem) { @($results) } else {
        @($results | Where-Object { -not [string]::IsNullOrWhiteSpace($_.SecureBootStatus) })
    }
    $localStateSummary = (@($filteredResults | Group-Object LocalDataFreshness | Sort-Object Name |
        ForEach-Object { '{0}={1}' -f $_.Name, $_.Count }) -join '; ')
    WriteLog -Message ("Exportable results: {0} devices; recurring={1}; local states: {2}" -f
        $filteredResults.Count, $UseRecurringDeviceSystem, $localStateSummary) 'INFO'

    if ($filteredResults.Count -eq 0) {
        WriteLog -Message "No devices with local Device System data to export. CSV will not be generated." "WARNING"
        $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
    }

    # ======================================================
    # Export CSV
    # ======================================================
    Write-Host "`n--- Export CSV ---"
$BaseFileName = "Intune_Devices_LocalSystem"

    ExportAndCopyCsv -BaseFileName $BaseFileName `
       -OutputPath $OutputPath `
       -GlobalPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') `
       -Data $filteredResults `
       -Encoding "UTF8" `
       -NoTypeInformation

    WriteLog -Message "Export completed: $global:csvFilePath1"
    WriteLog -Message "Number of exported devices: $($filteredResults.Count)"


    $smartM365RunGuard.ReachedTerminal=$true
}
catch {
    if ($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]) { throw }
    $smartM365RunGuard.ErrorRecord=$_

    $globalError = $_
    WriteLog -Message ("Global error in Device System inventory: {0}" -f $globalError) "ERROR"
    Write-Host "A global error occurred. Check the log file for details." -ForegroundColor Red
}
finally {
    $smartM365RunGuard.InCleanup=$true

    # Disconnect Graph only if we connected it in this run
    if ($connectedGraphInThisRun) {
        Write-Host "`n--- Disconnect Cloud Services ---"
        try {
            if (Get-MgContext -ErrorAction SilentlyContinue) {
                try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null; Set-SmartM365InventoryConnectionOwnership -Connection Graph -Owned $false } catch { }
            }
        } catch {
            WriteLog -Message ("Error during Graph disconnect in finally: {0}" -f $_) "WARNING"
        }
    }

    # Cleanup
    try {
        Remove-SmartM365TimestampedFilesOlderThan -FolderPath $OutputPath -FilePattern '*.csv' -RetentionDays 7 -RequireCurrentRunPublication -LogFile $global:LogTextFile
        RemoveOldFiles -Path $global:LogPath -Filter "*.log" -KeepCount $global:RetentionMaxLogs -LogFile $global:LogTextFile
    } catch {
        WriteLog -Message ("Error during cleanup in finally: {0}" -f $_) "WARNING"
    }

    WriteLog -Message "$TaskName completed (finally block)."

    try {
        Stop-SmartM365RunTranscript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
        Complete-SmartM365ExecutionContext -Status Auto
    } catch {
        # ignore
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBqqiecDcEsCP+9
# Xm57Oy//0593RdKAIVAu+KCScaijiKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIGFYcPd8IC0z83ckFM5Gn+D10nqzE7HQbSiX/frfbIeBMA0GCSqG
# SIb3DQEBAQUABIIBgBF4KMSCx8TZX/OXkEdukEZQF6P8TpsgGHyoSwEgttPY00kG
# RmWjb4RnSc+PUG6pv7VG3p0v5ITC4rIkZBG4dEf4fYKpIrBIQeo3JSN8MIKRWxI0
# fgxGQFlNHGMv3THe+zBvxbHLU4g1VBaN999CJNXmkTZYQ/DmZ3HEO+bu8/HsCOMl
# XQ5MRjjQPzF1PRvuMJbIlRAMTEEn+FaH4gc9458PNJ5+ji+tNnR7lJhROgxt2pPY
# bo2mTOawyWNUupOgRPTl7RoUHRC2j9IxER0yW+sCHP3xQF4TkRetlmB4GZxAo1Dy
# vr4z8wiasrde0A1rV9du9+zTRB01rvKRzvLG+e58qujZZlLnWqYWACBVlqrWZFAE
# UJITaNqV1gdIn6a0r29NUIx9FvCKu/lXJzjrvOd2cfi0JU7KOpWsOH4TyuoxWggC
# uMPrIkbr+gc4aZ1ecBUeIZR8W8kiYfexKX12XwZpHM7DnzqPyR8ZZXhL8JYVlNlC
# XIvdEuxBNfdeplQC0KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMzExNDA3
# MjFaMC8GCSqGSIb3DQEJBDEiBCABJfmeFxz0VdsxItAbJsPJvjg2bXRRWO4dDKht
# G04kdjANBgkqhkiG9w0BAQEFAASCAgBRvYyd1IrhMFAirD2J7K8b3mM6JxO8XYRJ
# +PDmeXEXv/1tt49WkHiEVgKl2ruihnT3svyCdnAdrIlmCQK1ZIG14Slor0qIucvL
# Q4f3yPZnYEZhwB2K05JPKodftb6eamxFriZ5++jjHCBYkHBCy1kdpUwjoEiztSCS
# S0dtgI8FQO8bs1lq6uGltwmd3qTr94IwlzVis9mT3Uys8wPOcx3lwUOc+5xx6ZW/
# 53ufzw5YCdzRVfgW3Y99KF6Uykt77FKot9eloVrIQ29vWrLjKYPC9drQkrgyM5o/
# N869Ss6ZoqgT4UemGTrozdTYzK2JEHDLGWj+bwCooFmXBTpl9D1+TTP8Xq3mw2Ta
# 2OTKvfJG0OhEJ44/5l6svmw+CyJYs7k4yY3w1ka5frWYzR2q7x7ZSHKX0EeYUQbq
# acCkAgrbP7Vrwi0/G+A0wsnkj3oAZWqX4RY0RPsc0JrdhQ6bKLaaob/dquUNoPcx
# oK/XBQFwYu2l41Ya2B17ouL0hWUOxeNe/o5YTaeN43pbeBU8ZqO5aJTXVQS9P50E
# luvvN4vznMWlBMcynBorzVxIT72bAoXRzY2KpO8rlrCYH67jY38Nx/cToXNDNp3Q
# l0s/7YHMy3AUjRHtzhA8hT/SpEXTmXXlI3nfN08Bo+/7EZu7JGlA9yNH8hhgpmM4
# HoakTPyULg==
# SIG # End signature block
