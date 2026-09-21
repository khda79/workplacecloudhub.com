<#
.SYNOPSIS
M365 Devices Compliance Inventory (Graph-only, Windows-only).

.DESCRIPTION
    - All Windows devices by default (if no filter is supplied)
    - No RAM/Storage/DHA collection
    - Per-policy compliance with fixed category columns (always present, even empty)
    - AD_Domain / AD_OU / DirectorySource from Entra ID (Graph) only:
        * AD_OU from onPremisesDistinguishedName (Hybrid only)
        * AD_Domain from onPremisesDomainName or fallback to UPN suffix

.PARAMETER ManagedDeviceId
Optional: limit scope to a specific Intune managed device.

.PARAMETER DeviceName
Optional: limit scope to a specific Windows device name (exact or startswith).

.PARAMETER IncludeComplianceSettings
Fetch and process per-setting noncompliant details to compute category rollup when policy-state detail collection is enabled.

.PARAMETER IncludePolicyStates
    Fetch per-device compliance policy states and optional setting states. Disabled by default because this is expensive on large tenants.

.PARAMETER PolicyStateMaxRuntimeMinutes
    Maximum wall-clock duration for detailed policy-state collection before the circuit breaker preserves the summary-only result.

.PARAMETER EnableDirectoryEnrichment
Resolve Entra directory details such as on-premises OU/domain per device. This adds Graph calls and is disabled by default for full-tenant runs.

.PARAMETER MaxDevices
Optional cap for smoke tests. 0 means no cap.

.PARAMETER AllDevices
If present (or if no device filter is supplied), process all Windows managed devices.

.PARAMETER Connect
Forces a (re)connection to Microsoft Graph (disconnects any existing session first).

.PARAMETER InteractiveAuth
Uses interactive authentication instead of app-only certificate authentication.
    Version : 1.18

.VERSION
1.18


.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication.
    Minimum Graph application permissions: DeviceManagementManagedDevices.Read.All; DeviceManagementConfiguration.Read.All; Device.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Version : 1.18
Requires    : PowerShell 7+, SmartM365.Core, Microsoft Graph PowerShell SDK
Scopes      : DeviceManagementManagedDevices.Read.All, Directory.Read.All
    Minimum application permissions: DeviceManagementManagedDevices.Read.All, DeviceManagementConfiguration.Read.All, Device.Read.All
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
[Parameter(Mandatory = $false)]
    [string]$ManagedDeviceId,

    [Parameter(Mandatory = $false)]
    [string]$DeviceName,

    [Parameter(Mandatory = $false)]
    [bool]$IncludeComplianceSettings = $true,

    [Parameter(Mandatory = $false)]
    [bool]$IncludePolicyStates = $false,

    [Parameter(Mandatory = $false)]
    [bool]$EnableDirectoryEnrichment = $false,

    [Parameter(Mandatory = $false)]
    [int]$MaxDevices = 0,

    [Parameter(Mandatory = $false)]
    [int]$PolicyStateMaxRuntimeMinutes = 60,

    [Parameter(Mandatory = $false)]
    [switch]$AllDevices,

    [Parameter(Mandatory = $false)]
    [switch]$Connect,

    [Parameter(Mandatory = $false)]
    [switch]$InteractiveAuth,

    [Parameter(Mandatory = $false)]
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

# ==========================================================
# PowerShell 7 minimum
# ==========================================================
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    exit 1
}

# Avoid PS function-capacity issues
$MaximumFunctionCount = 32768

# ==========================================================
# App-only authentication parameters (same app as inventory script)
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
$AppId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AppId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$TenantId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$Thumb = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumb' -DefaultValue '0000000000000000000000000000000000000000'
$OrgDomain = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue 'contoso.onmicrosoft.com'
$LogAllRootPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LogAllRootPath' -DefaultValue ''

# ==========================================================
# Import SmartM365.Core module (psd1)
# ==========================================================
$modulePath = & { $d = $PSScriptRoot; while ($d) { $p = Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'; if (Test-Path -LiteralPath $p) { return $p }; $parent = Split-Path -Path $d -Parent; if ($parent -eq $d) { break }; $d = $parent }; throw 'SmartM365.Core module not found.' }
try {
    Import-Module -Name $modulePath -MinimumVersion '1.0.24' -ErrorAction Stop
} catch {
    Write-Host "Failed to import SmartM365.Core module from '$modulePath' : $_" -ForegroundColor Red
    exit 1
}

# ==========================================================
# Fixed output paths and transcript
# ==========================================================
$ScriptVersion = "1.18"
$ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$TaskName = "$ScriptName v$ScriptVersion"
$ts = Get-Date -Format 'yyyyMMdd_HHmmss'
$ScriptCsvLogFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ScriptCsvLogFolderPath' -DefaultValue ""
$LatestCsvFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ""
$OutputPath = $ScriptCsvLogFolderPath
$script:GraphRequestDelayMs = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'GraphRequestDelayMs' -DefaultValue 250)
$script:GraphMaxRetryAttempts = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'GraphMaxRetryAttempts' -DefaultValue 6)
$script:GraphRetryMaxSeconds = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'GraphRetryMaxSeconds' -DefaultValue 180)
$script:ManagedDevicePageSize = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ManagedDevicePageSize' -DefaultValue 999)
if ($script:ManagedDevicePageSize -lt 1) { $script:ManagedDevicePageSize = 999 }
if ($script:ManagedDevicePageSize -gt 999) { $script:ManagedDevicePageSize = 999 }
$script:MaxDevicesEffective = if ($PSBoundParameters.ContainsKey('MaxDevices')) {
    [int]$MaxDevices
} elseif ($PSBoundParameters.ContainsKey('MaxItems') -and $MaxItems -gt 0) {
    [int]$MaxItems
} else {
    [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'MaxDevices' -DefaultValue 0)
}
if ($script:MaxDevicesEffective -lt 0) { $script:MaxDevicesEffective = 0 }
$script:IncludePolicyStatesExplicit = $PSBoundParameters.ContainsKey('IncludePolicyStates')
$script:IncludePolicyStatesEffective = if ($script:IncludePolicyStatesExplicit) { [bool]$IncludePolicyStates } else { [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'IncludePolicyStates' -DefaultValue $false) }
$script:EnableDirectoryEnrichmentEffective = if ($PSBoundParameters.ContainsKey('EnableDirectoryEnrichment')) { [bool]$EnableDirectoryEnrichment } else { [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableDirectoryEnrichment' -DefaultValue $false) }
$script:PolicyStateMaxRuntimeMinutes = if ($PSBoundParameters.ContainsKey('PolicyStateMaxRuntimeMinutes')) { [int]$PolicyStateMaxRuntimeMinutes } else { [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'PolicyStateMaxRuntimeMinutes' -DefaultValue 60) }
if ($script:PolicyStateMaxRuntimeMinutes -lt 1) { $script:PolicyStateMaxRuntimeMinutes = 60 }
$script:PolicyStateAutoDisableDeviceThreshold = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'PolicyStateAutoDisableDeviceThreshold' -DefaultValue 5000)
if ($script:PolicyStateAutoDisableDeviceThreshold -lt 0) { $script:PolicyStateAutoDisableDeviceThreshold = 0 }
$script:PolicyStateBatchProgressInterval = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'PolicyStateBatchProgressInterval' -DefaultValue 25)
if ($script:PolicyStateBatchProgressInterval -lt 1) { $script:PolicyStateBatchProgressInterval = 25 }
$script:MaxPolicyStateFailures = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'MaxPolicyStateFailures' -DefaultValue 100)
$script:MaxConsecutivePolicyStateFailures = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'MaxConsecutivePolicyStateFailures' -DefaultValue 25)
$script:PolicyStateFailureCount = 0
$script:ConsecutivePolicyStateFailures = 0
$script:PolicyStateCollectionDisabled = -not $script:IncludePolicyStatesEffective
$script:PolicyDetailCollectionComplete = [bool]$script:IncludePolicyStatesEffective
$script:ComplianceFatalError = $null
$script:SettingBatchFallbackCounts = @{}
$script:SettingBatchFallbackExamples = [System.Collections.Generic.List[string]]::new()
$script:PolicyStateCollectionStartedAt = $null
$script:PolicyStateDeadlineUtc = $null
$script:PolicyStateCircuitBreakerLogged = $false
$script:PolicyStateBatchRetryCount = 0
$script:PolicyStateBatchThrottleCount = 0

$logDir = if ([string]::IsNullOrWhiteSpace($LogAllRootPath)) {
    Join-Path $ScriptCsvLogFolderPath "Log"
} else {
    Join-Path $LogAllRootPath $ScriptName
}

$mainCsv = Join-Path $ScriptCsvLogFolderPath "Intune_Devices_Compliance.csv"
$tsCsv = Join-Path $ScriptCsvLogFolderPath ("Intune_Devices_Compliance_{0}.csv" -f $ts)
$lastCsv = Join-Path $LatestCsvFolderPath "Intune_Devices_Compliance.csv"
$policyMainCsv = Join-Path $ScriptCsvLogFolderPath "Intune_Devices_Compliance_Policies.csv"
$policyTsCsv = Join-Path $ScriptCsvLogFolderPath ("Intune_Devices_Compliance_Policies_{0}.csv" -f $ts)
$policyLastCsv = Join-Path $LatestCsvFolderPath "Intune_Devices_Compliance_Policies.csv"

# These two canonical DATA-ALL files are written directly before the shared export
# helper runs. Suffix them explicitly in bounded mode so a smoke test cannot replace
# either production snapshot; timestamped and DATA-LAST paths are handled by Core.
if (Test-SmartM365MaxItemsMode) {
    $mainCsv = Add-SmartM365MaxItemsSuffixToCsvPath -Path $mainCsv
    $policyMainCsv = Add-SmartM365MaxItemsSuffixToCsvPath -Path $policyMainCsv
}

foreach ($dir in @($ScriptCsvLogFolderPath, $LatestCsvFolderPath, $logDir)) {
    try { New-Item -ItemType Directory -Force -Path $dir | Out-Null } catch { }
}

$global:LogPath = $logDir
$global:LogTextFile = Join-Path $logDir ("{0}-{1}.log" -f $ScriptName, (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'))
$global:logTranscriptFile = Join-Path $logDir ("{0}-{1}_Transcript.log" -f $ScriptName, (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'))
$global:SmartM365ExecutionStartTime = Get-Date
$global:SmartM365ExecutionSummaryWritten = $false
Set-SmartM365CoreContext -RunId $ts -RunOutputRoot $ScriptCsvLogFolderPath -LatestOutputRoot $LatestCsvFolderPath -LogPath $global:LogTextFile

function Write-ComplianceWarning {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Message)

    WriteLog -Message $Message -Level 'WARNING'
}

function Write-ComplianceInfo {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Message)

    WriteLog -Message $Message -Level 'INFO'
}

function Assert-PolicyStateRuntimeAvailable {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Operation)

    if ($null -eq $script:PolicyStateDeadlineUtc -or [datetime]::UtcNow -lt $script:PolicyStateDeadlineUtc) {
        return
    }

    $script:PolicyStateCollectionDisabled = $true
    $script:PolicyDetailCollectionComplete = $false
    $message = "Detailed compliance policy-state collection reached its $($script:PolicyStateMaxRuntimeMinutes)-minute maximum runtime during '$Operation'. Device summary processing will continue and the last valid detailed DATA-LAST export will be preserved."
    if (-not $script:PolicyStateCircuitBreakerLogged) {
        Write-ComplianceWarning -Message $message
        $script:PolicyStateCircuitBreakerLogged = $true
    }

    $exception = [System.TimeoutException]::new($message)
    $exception.Data['SmartM365PolicyStateCircuitBreaker'] = $true
    throw $exception
}

function Test-PolicyStateCircuitBreakerException {
    [CmdletBinding()]
    param([AllowNull()]$ErrorRecord)

    try {
        return [bool]$ErrorRecord.Exception.Data['SmartM365PolicyStateCircuitBreaker']
    }
    catch {
        return $false
    }
}

try {
    $transcriptPath = $global:logTranscriptFile
    Start-Transcript -Path $transcriptPath -Force | Out-Null
} catch {
    Write-ComplianceWarning -Message "Failed to start transcript. $_"
}

# ==========================================================
# Helpers
# ==========================================================
function Get-SafeProperty {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][object]$Object,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }
        $matchingKey = @($Object.Keys | Where-Object { [string]$_ -ieq $Name } | Select-Object -First 1)
        if ($matchingKey.Count -gt 0) { return $Object[$matchingKey[0]] }
    }

    $prop = $Object.PSObject.Properties[$Name]
    if ($prop) { return $prop.Value }

    $ap = $Object.PSObject.Properties['AdditionalProperties']
    if ($ap -and $ap.Value -is [System.Collections.IDictionary] -and $ap.Value.ContainsKey($Name)) {
        return $ap.Value[$Name]
    }

    $propCI = $Object.PSObject.Properties | Where-Object { $_.Name -ieq $Name }
    if ($propCI) { return $propCI.Value }

    return $null
}

function Get-ShortGraphErrorMessage {
    [CmdletBinding()]
    param([AllowNull()]$ErrorRecord)

    $message = if ($ErrorRecord -and $ErrorRecord.Exception) { [string]$ErrorRecord.Exception.Message } else { [string]$ErrorRecord }
    $message = ($message -replace '\s+', ' ').Trim()
    if ($message.Length -gt 500) { return ($message.Substring(0, 500) + '...') }
    return $message
}

function Get-GraphRetryDelaySeconds {
    [CmdletBinding()]
    param(
        [AllowNull()]$ErrorRecord,
        [int]$Attempt,
        [int]$MaximumSeconds = 180
    )

    $retryAfter = $null
    try {
        if ($ErrorRecord.Exception.Response -and $ErrorRecord.Exception.Response.Headers) {
            $retryAfter = @($ErrorRecord.Exception.Response.Headers.GetValues('Retry-After') | Select-Object -First 1)[0]
        }
    } catch {}
    if (-not $retryAfter) { try { $retryAfter = $ErrorRecord.Exception.Data['Retry-After'] } catch {} }

    $seconds = 0
    if ($retryAfter -and [int]::TryParse([string]$retryAfter, [ref]$seconds) -and $seconds -gt 0) {
        return [math]::Min($seconds, $MaximumSeconds)
    }

    $backoff = [math]::Min($MaximumSeconds, [math]::Pow(2, [math]::Min($Attempt, 8)) * 5)
    return [int]($backoff + (Get-Random -Minimum 0 -Maximum 5))
}

function Get-GraphBatchRetryDelaySeconds {
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()][object[]]$Responses,
        [int]$Attempt,
        [int]$MaximumSeconds = 180
    )

    $maximumRetryAfter = 0
    foreach ($response in @($Responses)) {
        $retryAfter = $null
        $headers = $response.headers
        if ($headers -is [System.Collections.IDictionary]) {
            foreach ($key in @($headers.Keys)) {
                if ([string]::Equals([string]$key, 'Retry-After', [System.StringComparison]::OrdinalIgnoreCase)) {
                    $retryAfter = @($headers[$key] | Select-Object -First 1)[0]
                    break
                }
            }
        } elseif ($headers) {
            $property = $headers.PSObject.Properties |
                Where-Object { [string]::Equals($_.Name, 'Retry-After', [System.StringComparison]::OrdinalIgnoreCase) } |
                Select-Object -First 1
            if ($property) { $retryAfter = @($property.Value | Select-Object -First 1)[0] }
        }

        $seconds = 0
        if ($retryAfter -and [int]::TryParse([string]$retryAfter, [ref]$seconds) -and $seconds -gt $maximumRetryAfter) {
            $maximumRetryAfter = $seconds
        }
    }

    if ($maximumRetryAfter -gt 0) {
        return [math]::Min($maximumRetryAfter, $MaximumSeconds)
    }

    return [int][math]::Min($MaximumSeconds, [math]::Pow(2, [math]::Min($Attempt, 8)))
}

function Invoke-GraphBatchWithSubRequestRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Requests,
        [Parameter(Mandatory = $true)][string]$Operation,
        [Parameter(Mandatory = $false)][int]$ServerErrorMaxAttempts = 2
    )

    $responseMap = @{}
    $maxAttempts = [math]::Max(1, $script:GraphMaxRetryAttempts)
    $serverErrorMaxAttempts = [math]::Max(1, [math]::Min($ServerErrorMaxAttempts, $maxAttempts))
    $batchUri = "https://graph.microsoft.com/v1.0/" + '$batch'
    $totalInitialBatches = if ($Requests.Count -gt 0) { [int][math]::Ceiling($Requests.Count / 20.0) } else { 0 }
    $batchNumber = 0
    $progressWatch = [System.Diagnostics.Stopwatch]::StartNew()

    # Retry each throttled slice before sending the rest of the collection. The former
    # whole-collection retry kept submitting thousands of requests after Graph had
    # started returning 429, then waited only after the throttle storm was complete.
    for ($initialOffset = 0; $initialOffset -lt $Requests.Count; $initialOffset += 20) {
        Assert-PolicyStateRuntimeAvailable -Operation $Operation
        $batchNumber++
        $initialLast = [math]::Min($initialOffset + 19, $Requests.Count - 1)
        $pending = @($Requests[$initialOffset..$initialLast])

        for ($attempt = 1; $attempt -le $maxAttempts -and $pending.Count -gt 0; $attempt++) {
            Assert-PolicyStateRuntimeAvailable -Operation $Operation
            $nextPending = [System.Collections.Generic.List[object]]::new()
            $retryResponses = [System.Collections.Generic.List[object]]::new()
            $batchSize = [math]::Max(1, [int][math]::Floor(20 / [math]::Pow(2, $attempt - 1)))

            for ($offset = 0; $offset -lt $pending.Count; $offset += $batchSize) {
                Assert-PolicyStateRuntimeAvailable -Operation $Operation
                $last = [math]::Min($offset + $batchSize - 1, $pending.Count - 1)
                $slice = @($pending[$offset..$last])
                $requestById = @{}
                foreach ($request in $slice) { $requestById[[string]$request.id] = $request }

                $body = @{ requests = $slice } | ConvertTo-Json -Depth 6
                $batchResponse = Invoke-WithRetry -Operation $Operation -Script {
                    Invoke-MgGraphRequest -Method POST -Uri $batchUri -Body $body -ContentType 'application/json' -ErrorAction Stop
                }

                $receivedIds = @{}
                foreach ($response in @($batchResponse.responses)) {
                    $requestId = [string]$response.id
                    $receivedIds[$requestId] = $true
                    $status = [int]$response.status
                    $isThrottle = $status -eq 429
                    $isServerError = $status -in @(500, 502, 503, 504)
                    $canRetry = ($isThrottle -and $attempt -lt $maxAttempts) -or
                        ($isServerError -and $attempt -lt $serverErrorMaxAttempts)

                    if ($isThrottle) { $script:PolicyStateBatchThrottleCount++ }

                    if ($status -eq 200 -or -not $canRetry) {
                        $responseMap[$requestId] = $response
                        continue
                    }

                    [void]$nextPending.Add($requestById[$requestId])
                    [void]$retryResponses.Add($response)
                }

                foreach ($request in $slice) {
                    $requestId = [string]$request.id
                    if ($receivedIds.ContainsKey($requestId)) { continue }
                    if ($attempt -lt $maxAttempts) {
                        [void]$nextPending.Add($request)
                    } else {
                        $responseMap[$requestId] = [pscustomobject]@{
                            id = $requestId
                            status = 0
                            body = $null
                            headers = $null
                        }
                    }
                }
            }

            if ($nextPending.Count -gt 0) {
                $script:PolicyStateBatchRetryCount++
                $delay = Get-GraphBatchRetryDelaySeconds -Responses @($retryResponses) -Attempt $attempt -MaximumSeconds $script:GraphRetryMaxSeconds
                Write-ComplianceWarning -Message ("{0}: batch {1}/{2} paused for {3}s before retry {4}/{5}; {6} sub-request(s) pending." -f $Operation, $batchNumber, $totalInitialBatches, $delay, ($attempt + 1), $maxAttempts, $nextPending.Count)
                Start-Sleep -Seconds $delay
                Assert-PolicyStateRuntimeAvailable -Operation $Operation
                Write-ComplianceInfo -Message ("{0}: batch {1}/{2} resumed after throttle/transient retry delay." -f $Operation, $batchNumber, $totalInitialBatches)
                $pending = @($nextPending)
            } else {
                $pending = @()
            }
        }

        if ($totalInitialBatches -gt 0 -and ($batchNumber -eq 1 -or $batchNumber -eq $totalInitialBatches -or ($batchNumber % $script:PolicyStateBatchProgressInterval) -eq 0)) {
            $elapsedSeconds = [math]::Max(0.001, $progressWatch.Elapsed.TotalSeconds)
            $completedRequests = [math]::Min($Requests.Count, $initialLast + 1)
            $rate = $completedRequests / $elapsedSeconds
            $remainingSeconds = if ($rate -gt 0) { [math]::Max(0, ($Requests.Count - $completedRequests) / $rate) } else { 0 }
            $percent = [math]::Round(100 * $completedRequests / $Requests.Count, 1)
            Write-ComplianceInfo -Message ("{0}: batch {1}/{2}; sub-requests {3}/{4} ({5}%); elapsed {6}; rate {7:N2}/s; ETA {8}." -f $Operation, $batchNumber, $totalInitialBatches, $completedRequests, $Requests.Count, $percent, $progressWatch.Elapsed.ToString('hh\:mm\:ss'), $rate, ([timespan]::FromSeconds($remainingSeconds).ToString('hh\:mm\:ss')))
        }
    }

    return $responseMap
}
function Invoke-WithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$Script,

        [Parameter(Mandatory = $false)]
        [int]$MaxAttempts = $script:GraphMaxRetryAttempts,

        [Parameter(Mandatory = $false)]
        [string]$Operation = 'Graph request'
    )

    if ($MaxAttempts -lt 1) { $MaxAttempts = 1 }

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            if ($script:GraphRequestDelayMs -gt 0) { Start-Sleep -Milliseconds $script:GraphRequestDelayMs }
            return & $Script
        } catch {
            $statusCode = $null
            try { if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode } } catch {}
            $message = Get-ShortGraphErrorMessage -ErrorRecord $_
            $isTransient = $statusCode -in @(429, 500, 502, 503, 504) -or $message -match 'TooManyRequests|throttl|timeout|temporarily unavailable|InternalServerError'

            if (-not $isTransient -or $attempt -ge $MaxAttempts) {
                $statusText = if ($statusCode) { $statusCode } else { 'unknown' }
                throw ("{0} failed. Status={1}; Attempts={2}; Message={3}" -f $Operation, $statusText, $attempt, $message)
            }

            $delay = Get-GraphRetryDelaySeconds -ErrorRecord $_ -Attempt $attempt -MaximumSeconds $script:GraphRetryMaxSeconds
            $statusRetryText = if ($statusCode) { $statusCode } else { 'unknown' }
            Write-ComplianceWarning -Message ("{0} transient failure. Status={1}; attempt {2}/{3}; retrying in {4}s." -f $Operation, $statusRetryText, $attempt, $MaxAttempts, $delay)
            Start-Sleep -Seconds $delay
        }
    }
}


function Test-ComplianceGraphProperty {
    [CmdletBinding()]
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $InputObject) { return $false }
    if ($InputObject -is [System.Collections.IDictionary]) { return $InputObject.Contains($Name) }
    return $null -ne $InputObject.PSObject.Properties[$Name]
}

function Get-ComplianceGraphPropertyValue {
    [CmdletBinding()]
    param(
        [AllowNull()][object]$InputObject,
        [Parameter(Mandatory)][string]$Name
    )

    if ($null -eq $InputObject) { return $null }
    if ($InputObject -is [System.Collections.IDictionary]) { return $InputObject[$Name] }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -ne $property) { return $property.Value }
    return $null
}

function Get-ComplianceGraphPageShape {
    [CmdletBinding()]
    param([AllowNull()][object]$InputObject)

    if ($null -eq $InputObject) { return 'Type=<null>; Keys=<none>' }
    $typeName = $InputObject.GetType().FullName
    $names = if ($InputObject -is [System.Collections.IDictionary]) {
        @($InputObject.Keys | ForEach-Object { [string]$_ })
    }
    else {
        @($InputObject.PSObject.Properties.Name)
    }
    $visibleNames = @($names | Select-Object -First 20)
    $nameText = if ($visibleNames.Count -gt 0) { $visibleNames -join ', ' } else { '<none>' }
    return "Type=$typeName; Keys=$nameText"
}

function Invoke-GraphPagedCollection {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri,

        [Parameter(Mandatory = $false)]
        [string]$Operation = 'Graph paged collection',

        [Parameter(Mandatory = $false)]
        [int]$MaxItems = 0
    )

    $items = New-Object System.Collections.Generic.List[object]
    $visited = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $nextLink = $Uri
    $pageNumber = 0

    while ($nextLink) {
        if (-not $visited.Add([string]$nextLink)) {
            throw "$Operation returned a repeated @odata.nextLink; collection is incomplete."
        }
        $pageNumber++
        $currentUri = $nextLink
        $page = Invoke-WithRetry -Operation $Operation -Script {
            Invoke-MgGraphRequest -Method GET -Uri $currentUri -ErrorAction Stop
        }

        if (-not (Test-ComplianceGraphProperty -InputObject $page -Name 'value')) {
            $pageShape = Get-ComplianceGraphPageShape -InputObject $page
            throw "$Operation page $pageNumber returned an invalid Graph collection response without a value property. $pageShape"
        }
        foreach ($item in @(Get-ComplianceGraphPropertyValue -InputObject $page -Name 'value')) {
            if ($null -ne $item) { $items.Add($item) | Out-Null }
            if ($MaxItems -gt 0 -and $items.Count -ge $MaxItems) { break }
        }

        Write-ComplianceInfo -Message ("{0}: page {1}, total {2}" -f $Operation, $pageNumber, $items.Count)

        if ($MaxItems -gt 0 -and $items.Count -ge $MaxItems) { break }
        $nextLink = if (Test-ComplianceGraphProperty -InputObject $page -Name '@odata.nextLink') { [string](Get-ComplianceGraphPropertyValue -InputObject $page -Name '@odata.nextLink') } else { $null }
    }

    return $items.ToArray()
}

function Get-ManagedWindowsDevicesFast {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ManagedDeviceId,

        [Parameter(Mandatory = $false)]
        [string]$DeviceName,

        [Parameter(Mandatory = $false)]
        [int]$MaxItems = 0
    )

    $select = 'id,deviceName,manufacturer,model,operatingSystem,lastSyncDateTime,complianceState,complianceGracePeriodExpirationDateTime,azureADDeviceId,userPrincipalName'

    if (-not [string]::IsNullOrWhiteSpace($ManagedDeviceId)) {
        $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$ManagedDeviceId`?`$select=$select"
        $device = Invoke-WithRetry -Operation 'Get Intune managed device' -Script {
            Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop
        }
        return @($device)
    }

    $pageSize = $script:ManagedDevicePageSize
    if ($pageSize -lt 1 -or $pageSize -gt 999) { $pageSize = 999 }

    if (-not [string]::IsNullOrWhiteSpace($DeviceName)) {
        $escaped = $DeviceName.Replace("'", "''")
        $exactUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows' and deviceName eq '$escaped'&`$select=$select&`$top=1"
        $exact = Invoke-GraphPagedCollection -Uri $exactUri -Operation 'Get Intune managedDevices exact-name page' -MaxItems 1
        if ($exact -and $exact.Count -gt 0) { return @($exact) }

        $startsWithUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows' and startswith(deviceName,'$escaped')&`$select=$select&`$top=1"
        return @(Invoke-GraphPagedCollection -Uri $startsWithUri -Operation 'Get Intune managedDevices startswith-name page' -MaxItems 1)
    }

    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'&`$select=$select&`$top=$pageSize"
    return @(Invoke-GraphPagedCollection -Uri $uri -Operation 'Get Intune managedDevices page' -MaxItems $MaxItems)
}

function Get-CompliancePolicyStateBatchMap {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][object[]]$Devices)

    $result = @{}
    $eligibleDevices = @($Devices | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.Id) })
    $requests = [System.Collections.Generic.List[object]]::new()
    foreach ($device in $eligibleDevices) {
        $deviceId = [string]$device.Id
        [void]$requests.Add(@{
            id = $deviceId
            method = 'GET'
            url = "/deviceManagement/managedDevices/$deviceId/deviceCompliancePolicyStates?`$top=200"
        })
    }

    $responseMap = Invoke-GraphBatchWithSubRequestRetry -Requests @($requests) -Operation 'Get Intune compliance policy states batch'
    $failureCounts = @{}
    $failureExamples = [System.Collections.Generic.List[string]]::new()

    foreach ($request in $requests) {
        $deviceId = [string]$request.id
        $response = $responseMap[$deviceId]
        $status = if ($response) { [int]$response.status } else { 0 }
        if ($status -ne 200) {
            $statusKey = [string]$status
            if (-not $failureCounts.ContainsKey($statusKey)) { $failureCounts[$statusKey] = 0 }
            $failureCounts[$statusKey]++
            if ($failureExamples.Count -lt 5) { [void]$failureExamples.Add($deviceId) }
            continue
        }

        $responseBody = Get-ComplianceGraphPropertyValue -InputObject $response -Name 'body'
        if (-not (Test-ComplianceGraphProperty -InputObject $responseBody -Name 'value')) {
            $failureCounts['InvalidBody'] = 1 + [int]$failureCounts['InvalidBody']
            if ($failureExamples.Count -lt 5) { [void]$failureExamples.Add($deviceId) }
            continue
        }
        $values = [System.Collections.Generic.List[object]]::new()
        foreach ($value in @(Get-ComplianceGraphPropertyValue -InputObject $responseBody -Name 'value')) { if ($null -ne $value) { [void]$values.Add($value) } }
        $nextLink = [string](Get-ComplianceGraphPropertyValue -InputObject $responseBody -Name '@odata.nextLink')
        $visited = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        while (-not [string]::IsNullOrWhiteSpace($nextLink)) {
            if (-not $visited.Add($nextLink)) { throw 'Compliance policy-state pagination returned a repeated @odata.nextLink; collection is incomplete.' }
            $page = Invoke-WithRetry -Operation 'Get Intune compliance policy-state continuation page' -Script {
                Invoke-MgGraphRequest -Method GET -Uri $nextLink -ErrorAction Stop
            }
            if (-not (Test-ComplianceGraphProperty -InputObject $page -Name 'value')) { throw 'Compliance policy-state continuation returned an invalid Graph collection response without a value property.' }
            foreach ($value in @(Get-ComplianceGraphPropertyValue -InputObject $page -Name 'value')) { if ($null -ne $value) { [void]$values.Add($value) } }
            $nextLink = [string](Get-ComplianceGraphPropertyValue -InputObject $page -Name '@odata.nextLink')
        }
        $result[$deviceId] = @($values)
    }

    if ($failureCounts.Count -gt 0) {
        $failureTotal = ($failureCounts.Values | Measure-Object -Sum).Sum
        $statusSummary = (@($failureCounts.GetEnumerator() | Sort-Object Name | ForEach-Object { "HTTP $($_.Name)=$($_.Value)" }) -join ', ')
        Write-ComplianceWarning -Message ("{0} compliance policy-state batch sub-request(s) still failed after targeted retries ({1}). Sequential fallback will be used. Sample managed device IDs: {2}" -f $failureTotal, $statusSummary, ($failureExamples -join ', '))
    }

    return $result
}

function Get-ComplianceSettingStateBatchMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$ManagedDeviceId,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Policies
    )

    $result = @{}
    $targets = @($Policies | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.id) })
    $requests = [System.Collections.Generic.List[object]]::new()
    foreach ($policy in $targets) {
        $policyId = [string]$policy.id
        $escapedPolicyId = [uri]::EscapeDataString($policyId)
        [void]$requests.Add(@{
            id = $policyId
            method = 'GET'
            url = "/deviceManagement/managedDevices/$ManagedDeviceId/deviceCompliancePolicyStates/$escapedPolicyId/settingStates?`$select=setting,state&`$top=200"
        })
    }

    # A 5xx returned only inside the batch commonly succeeds through the existing direct fallback.
    # Retry throttles here, but do not multiply per-device delays for a batch-only server error.
    $responseMap = Invoke-GraphBatchWithSubRequestRetry -Requests @($requests) -Operation 'Get Intune compliance setting states batch' -ServerErrorMaxAttempts 1
    $failureCounts = @{}
    $failureExamples = [System.Collections.Generic.List[string]]::new()

    foreach ($request in $requests) {
        $policyId = [string]$request.id
        $response = $responseMap[$policyId]
        $status = if ($response) { [int]$response.status } else { 0 }
        if ($status -ne 200) {
            $statusKey = [string]$status
            if (-not $failureCounts.ContainsKey($statusKey)) { $failureCounts[$statusKey] = 0 }
            $failureCounts[$statusKey]++
            if ($failureExamples.Count -lt 5) { [void]$failureExamples.Add($policyId) }
            continue
        }

        $responseBody = Get-ComplianceGraphPropertyValue -InputObject $response -Name 'body'
        if (-not (Test-ComplianceGraphProperty -InputObject $responseBody -Name 'value')) {
            $failureCounts['InvalidBody'] = 1 + [int]$failureCounts['InvalidBody']
            if ($failureExamples.Count -lt 5) { [void]$failureExamples.Add($policyId) }
            continue
        }
        $values = [System.Collections.Generic.List[object]]::new()
        foreach ($value in @(Get-ComplianceGraphPropertyValue -InputObject $responseBody -Name 'value')) { if ($null -ne $value) { [void]$values.Add($value) } }
        $nextLink = [string](Get-ComplianceGraphPropertyValue -InputObject $responseBody -Name '@odata.nextLink')
        $visited = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        while (-not [string]::IsNullOrWhiteSpace($nextLink)) {
            if (-not $visited.Add($nextLink)) { throw 'Compliance setting-state pagination returned a repeated @odata.nextLink; collection is incomplete.' }
            $page = Invoke-WithRetry -Operation 'Get Intune compliance setting-state continuation page' -Script {
                Invoke-MgGraphRequest -Method GET -Uri $nextLink -ErrorAction Stop
            }
            if (-not (Test-ComplianceGraphProperty -InputObject $page -Name 'value')) { throw 'Compliance setting-state continuation returned an invalid Graph collection response without a value property.' }
            foreach ($value in @(Get-ComplianceGraphPropertyValue -InputObject $page -Name 'value')) { if ($null -ne $value) { [void]$values.Add($value) } }
            $nextLink = [string](Get-ComplianceGraphPropertyValue -InputObject $page -Name '@odata.nextLink')
        }
        $result[$policyId] = @($values)
    }

    if ($failureCounts.Count -gt 0) {
        $failureTotal = ($failureCounts.Values | Measure-Object -Sum).Sum
        $statusSummary = (@($failureCounts.GetEnumerator() | Sort-Object Name | ForEach-Object { "HTTP $($_.Name)=$($_.Value)" }) -join ', ')
        foreach ($failure in $failureCounts.GetEnumerator()) {
            $statusKey = [string]$failure.Key
            if (-not $script:SettingBatchFallbackCounts.ContainsKey($statusKey)) {
                $script:SettingBatchFallbackCounts[$statusKey] = 0
            }
            $script:SettingBatchFallbackCounts[$statusKey] += [int]$failure.Value
        }
        foreach ($policyId in $failureExamples) {
            if ($script:SettingBatchFallbackExamples.Count -ge 5) { break }
            [void]$script:SettingBatchFallbackExamples.Add(("{0}/{1}" -f $ManagedDeviceId, $policyId))
        }
    }

    return $result
}

function Get-ADPartsFromDN {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$DN
    )

    if ([string]::IsNullOrWhiteSpace($DN)) {
        return [pscustomobject]@{
            OU     = $null
            Domain = $null
        }
    }

    $parts   = $DN -split ',' | ForEach-Object { $_.Trim() }
    $ouParts = @()
    $dcParts = @()

    foreach ($p in $parts) {
        if ($p -like 'OU=*') {
            $ouParts += ($p.Substring(3))
        } elseif ($p -like 'DC=*') {
            $dcParts += ($p.Substring(3))
        }
    }

    $ou     = if ($ouParts.Count -gt 0) { $ouParts -join '/' } else { $null }
    $domain = if ($dcParts.Count -gt 0) { $dcParts -join '.' } else { $null }

    return [pscustomobject]@{
        OU     = $ou
        Domain = $domain
    }
}

# Resolve directory info (OU and Domain) from Azure AD / Entra ID only (no on-prem AD calls)
function Resolve-DirInfoFromGraph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$AzureAdDeviceId,

        [Parameter(Mandatory = $false)]
        [string]$FallbackUpn
    )

    $out = [pscustomobject]@{
        AD_OU           = $null
        AD_Domain       = $null
        EntraObjectId   = $null
        DirectorySource = 'Unknown'
    }

    try {
        if ([string]::IsNullOrWhiteSpace($AzureAdDeviceId)) {
            if ($FallbackUpn) {
                $out.AD_Domain = ($FallbackUpn -split '@', 2)[1]
            }

            $out.DirectorySource = if ($out.AD_Domain) { 'AADOnly' } else { 'Unknown' }
            return $out
        }

        $uri  = "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$AzureAdDeviceId'&`$select=id,deviceId,trustType,onPremisesDomainName,onPremisesDistinguishedName"
        $resp = Invoke-WithRetry -Operation "Get Intune Graph page" -Script { Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop }

        $dev = $null
        if ($resp -and $resp.value) {
            if ($resp.value -is [System.Collections.IDictionary]) {
                $dev = $resp.value
            } elseif ($resp.value -is [System.Collections.IEnumerable]) {
                $dev = ($resp.value | Select-Object -First 1)
            } elseif ($resp.value -is [object[]]) {
                $dev = $resp.value[0]
            }
        }

        if ($dev) {
            $dn        = $dev.onPremisesDistinguishedName
            $domain    = $dev.onPremisesDomainName
            $trustType = $dev.trustType

            if ($dn) {
                $parts         = Get-ADPartsFromDN -DN $dn
                $out.AD_OU     = $parts.OU
                if (-not $domain -and $parts.Domain) {
                    $domain = $parts.Domain
                }
            }

            if (-not $domain -and $FallbackUpn) {
                $domain = ($FallbackUpn -split '@', 2)[1]
            }

            $out.AD_Domain       = $domain
            $out.EntraObjectId   = $dev.id
            $out.DirectorySource = if ($trustType -eq 'ServerAd' -or $dn -or $dev.onPremisesDomainName) {
                'Hybrid'
            } elseif ($trustType -eq 'Workplace') {
                'Registered'
            } elseif ($trustType -eq 'AzureAd' -or $domain) {
                'AADOnly'
            } else {
                'Unknown'
            }
        } elseif ($FallbackUpn) {
            $out.AD_Domain       = ($FallbackUpn -split '@', 2)[1]
            $out.DirectorySource = 'AADOnly'
        }
    } catch {
        Write-Verbose "Failed to resolve directory info from Graph: $_"
    }

    return $out
}

# ==========================================================
# Category mapping (for per-policy rollup)
# ==========================================================
$SettingRuleMap = @(
    @{ Pattern='secureboot(enabled)?';                          Category='SecureBoot' },
    @{ Pattern='bitlocker|encrypt';                             Category='BitLocker' },
    @{ Pattern='tpm|requiredtrustedplatformmodule';             Category='TPM' },
    @{ Pattern='defender|antivirus|antispyware|deviceThreat';   Category='Antivirus' },
    @{ Pattern='firewall';                                      Category='Firewall' },
    @{ Pattern='codeintegrity';                                 Category='CodeIntegrity' },
    @{ Pattern='os(version|minimum)|minosversion';              Category='OSVersion' },
    @{ Pattern='uefi(required)?';                               Category='UEFI' }
)

# Policy property name -> category (for Get-PolicyConfiguredCategories)
$PolicyPropertyCategoryMap = @(
    @{ Properties=@('secureBootEnabled');                                    Category='SecureBoot'     },
    @{ Properties=@('bitLockerEnabled','storageRequireEncryption');           Category='BitLocker'      },
    @{ Properties=@('tpmRequired');                                          Category='TPM'            },
    @{ Properties=@('antivirusRequired','antiSpywareRequired','defenderEnabled','rtpEnabled','signatureOutOfDate','deviceThreatProtectionEnabled'); Category='Antivirus' },
    @{ Properties=@('firewallEnabled','firewallBlockAllIncoming','firewallEnableStealthMode'); Category='Firewall' },
    @{ Properties=@('codeIntegrityEnabled');                                 Category='CodeIntegrity'  },
    @{ Properties=@('osMinimumVersion','osMaximumVersion','mobileOsMinimumVersion','mobileOsMaximumVersion','validOperatingSystemBuildRanges'); Category='OSVersion' },
    @{ Properties=@('uefiRequired');                                         Category='UEFI'           }
)

# Cache: policyId -> Set of configured category names
$policyDefCache = @{}

function Get-PolicyConfiguredCategories {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$PolicyId
    )

    if ($script:policyDefCache.ContainsKey($PolicyId)) {
        return $script:policyDefCache[$PolicyId]
    }

    $configured = [System.Collections.Generic.HashSet[string]]::new()

    try {
        $uri = "https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies/$PolicyId"
        $policy = Invoke-WithRetry -Operation "Get Intune Graph page" -Script { Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop }

        if ($policy) {
            foreach ($entry in $script:PolicyPropertyCategoryMap) {
                foreach ($prop in $entry.Properties) {
                    $val = $null
                    if (Test-ComplianceGraphProperty -InputObject $policy -Name $prop) {
                        $val = Get-ComplianceGraphPropertyValue -InputObject $policy -Name $prop
                    }
                    # A property is "configured" if it is true, or a non-empty/non-null string
                    $active = ($val -is [bool] -and $val -eq $true) -or
                              ($val -is [string] -and -not [string]::IsNullOrWhiteSpace($val)) -or
                              ($val -is [System.Collections.IEnumerable] -and ($val | Measure-Object).Count -gt 0)
                    if ($active) {
                        $null = $configured.Add($entry.Category)
                        Write-Verbose ("Get-PolicyConfiguredCategories: '{0}' property '{1}'='{2}' -> category '{3}' configured" -f $PolicyId, $prop, $val, $entry.Category)
                        break
                    }
                }
            }
        }
    } catch {
        Write-ComplianceWarning -Message ("Failed to retrieve policy definition for '{0}': {1}" -f $PolicyId, $_.Exception.Message)
    }

    $script:policyDefCache[$PolicyId] = $configured
    return $configured
}

function Map-SettingCategory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$SettingName
    )

    if ([string]::IsNullOrWhiteSpace($SettingName)) {
        return 'Other'
    }

    $n = $SettingName.ToLowerInvariant()
    foreach ($rule in $SettingRuleMap) {
        if ($n -match $rule.Pattern) {
            Write-Verbose ("Map-SettingCategory: '{0}' -> '{1}'" -f $SettingName, $rule.Category)
            return $rule.Category
        }
    }
    Write-Verbose ("Map-SettingCategory: '{0}' -> 'Other' (no pattern matched)" -f $SettingName)
    return 'Other'
}

# ==========================================================
# Graph connection via SmartM365.Core / Connect-SmartM365CloudSession
# ==========================================================
function Test-GraphConnection {
    try {
        # Works with delegated or app-only
        $null = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -ErrorAction Stop
        return $true
    } catch {
        return $false
    }
}

$connectedGraphInThisRun = $false

try {
    $graphContext = $null
    if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
        try { $graphContext = Get-MgContext -ErrorAction SilentlyContinue } catch { }
    }

    $needConnect = $false

    if ($Connect) {
        Write-ComplianceInfo -Message "Connect switch specified: existing Graph session (if any) will be disconnected and reconnected."
        Disconnect-SmartM365CloudSession -ExchangeOnline $false -Graph $true -VerboseDisconnect:$true
        $needConnect = $true
    } else {
        if ($graphContext -and (Test-GraphConnection)) {
            Write-ComplianceInfo -Message "Existing Microsoft Graph session detected. Reusing current connection."
            $needConnect = $false
        } else {
            Write-ComplianceInfo -Message "No existing Graph session detected. Will establish a new connection."
            $needConnect = $true
        }
    }

    if ($needConnect) {
        $connectParams = @{
            ExchangeOnline = $false
            Graph          = $true
            GraphScopes    = @("DeviceManagementManagedDevices.Read.All","Directory.Read.All")
        }

        if (-not $InteractiveAuth) {
            # App-only certificate authentication
            $connectParams.AppId        = $AppId
            $connectParams.Thumbprint   = $Thumb
            $connectParams.TenantId     = $TenantId
            $connectParams.Organization = $OrgDomain
            Write-ComplianceInfo -Message "Connecting to Microsoft Graph with app-only certificate authentication."
        } else {
            Write-ComplianceInfo -Message "Connecting to Microsoft Graph with interactive authentication."
        }

        $connectResult = Connect-SmartM365CloudSession @connectParams

        if (-not $connectResult.GraphConnected) {
            throw "Failed to connect to Microsoft Graph."
        }

        $connectedGraphInThisRun = $connectResult.GraphConnected
    }

    # ==========================================================
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -RequiredGraphApplicationPermissions @('DeviceManagementManagedDevices.Read.All','DeviceManagementConfiguration.Read.All','Device.Read.All') -GraphProbeUris @(
        'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$top=1',
        'https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies?$top=1'
    ) | Out-Null

    # MAIN LOGIC
    # ==========================================================

    # 1) Resolve device set
    $devices = @()

    try {
        $processAll = $AllDevices.IsPresent -or `
                      ([string]::IsNullOrWhiteSpace($ManagedDeviceId) -and [string]::IsNullOrWhiteSpace($DeviceName))

        if ($processAll) {
            Write-ComplianceInfo -Message "Retrieving all Intune managed Windows devices with explicit Graph paging."
            if ($script:MaxDevicesEffective -gt 0) {
                Write-ComplianceWarning -Message ("MaxDevices smoke cap active: {0}" -f $script:MaxDevicesEffective)
            }
            $devices = @(Get-ManagedWindowsDevicesFast -MaxItems $script:MaxDevicesEffective)
            if (-not $devices -or $devices.Count -eq 0) {
                Write-Error "No Windows managed devices found."
                throw "No Windows managed devices found."
            }
        } elseif (-not [string]::IsNullOrWhiteSpace($ManagedDeviceId)) {
            Write-ComplianceInfo -Message "Resolving device by ManagedDeviceId '$ManagedDeviceId'."
            $devices = @(Get-ManagedWindowsDevicesFast -ManagedDeviceId $ManagedDeviceId -MaxItems 1)
            if (-not $devices -or $devices.Count -eq 0) {
                throw ("No device found with ManagedDeviceId='{0}'." -f $ManagedDeviceId)
            }
        } else {
            Write-ComplianceInfo -Message "Resolving device by DeviceName '$DeviceName'."
            $devices = @(Get-ManagedWindowsDevicesFast -DeviceName $DeviceName -MaxItems 1)
            if (-not $devices -or $devices.Count -eq 0) {
                throw ("No Windows device found for DeviceName='{0}'." -f $DeviceName)
            }
        }

        Write-ComplianceInfo -Message ("Managed Windows devices selected for compliance summary: {0}" -f @($devices).Count)
    } catch {
        Write-Error "Failed to resolve target devices. $_"
        throw
    }

    # 2) Collect per-device summary rows and per-policy rows
    $aadCache = @{}
    $rows     = New-Object System.Collections.Generic.List[object]
    $polAll   = New-Object System.Collections.Generic.List[object]

    $policyStateBatchMap = @{}
    if ($script:IncludePolicyStatesEffective -and -not $script:IncludePolicyStatesExplicit -and
        $script:PolicyStateAutoDisableDeviceThreshold -gt 0 -and @($devices).Count -gt $script:PolicyStateAutoDisableDeviceThreshold) {
        $script:IncludePolicyStatesEffective = $false
        $script:PolicyStateCollectionDisabled = $true
        $script:PolicyDetailCollectionComplete = $true
        Write-ComplianceWarning -Message ("Detailed compliance policy-state collection was automatically disabled for {0} devices because the configured threshold is {1}. The device summary will continue. Use -IncludePolicyStates `$true to explicitly request the bounded detailed workflow." -f @($devices).Count, $script:PolicyStateAutoDisableDeviceThreshold)
    }

    if (-not $script:PolicyStateCollectionDisabled) {
        $script:PolicyStateCollectionStartedAt = Get-Date
        $script:PolicyStateDeadlineUtc = [datetime]::UtcNow.AddMinutes($script:PolicyStateMaxRuntimeMinutes)
        $expectedPolicyStateBatches = [int][math]::Ceiling(@($devices).Count / 20.0)
        Write-ComplianceInfo -Message ("Detailed compliance policy-state collection started: {0} devices, approximately {1} Graph batches, maximum runtime {2} minute(s)." -f @($devices).Count, $expectedPolicyStateBatches, $script:PolicyStateMaxRuntimeMinutes)
    }

    if (-not $script:PolicyStateCollectionDisabled) {
        try {
            $policyStateBatchMap = Get-CompliancePolicyStateBatchMap -Devices @($devices)
            Write-ComplianceInfo -Message ("Compliance policy-state batches completed: {0}/{1} devices prefetched." -f $policyStateBatchMap.Count, @($devices).Count)
        }
        catch {
            if (Test-PolicyStateCircuitBreakerException -ErrorRecord $_) {
                Write-ComplianceInfo -Message 'Compliance policy-state sequential fallback skipped because the runtime circuit breaker is active.'
            }
            else {
                Write-ComplianceWarning -Message ("Compliance policy-state batching failed; sequential retrieval will be used: {0}" -f $_.Exception.Message)
            }
            $policyStateBatchMap = @{}
        }
    }

    $total = ($devices | Measure-Object).Count
    $i     = 0

    foreach ($dev in $devices) {
        $i++
        Write-Progress -Id 1 -Activity "Processing devices" `
            -Status ("{0}/{1} - {2}" -f $i, $total, $dev.DeviceName) `
            -PercentComplete ([int](100 * $i / $total))

        $lastSync   = Get-SafeProperty -Object $dev -Name 'lastSyncDateTime'
        $azureId    = Get-SafeProperty -Object $dev -Name 'azureADDeviceId'
        $primaryUpn = Get-SafeProperty -Object $dev -Name 'userPrincipalName'

        # Directory info (Graph-only)
        $adOU         = $null
        $adDomain     = $null
        $dirSource    = $null
        $entraObjId   = $null

        if ($script:EnableDirectoryEnrichmentEffective -and $azureId) {
            if ($aadCache.ContainsKey($azureId)) {
                $adOU       = $aadCache[$azureId].AD_OU
                $adDomain   = $aadCache[$azureId].AD_Domain
                $dirSource  = $aadCache[$azureId].DirectorySource
                $entraObjId = $aadCache[$azureId].EntraObjectId
            } else {
                $info       = Resolve-DirInfoFromGraph -AzureAdDeviceId $azureId -FallbackUpn $primaryUpn
                $adOU       = $info.AD_OU
                $adDomain   = $info.AD_Domain
                $dirSource  = $info.DirectorySource
                $entraObjId = $info.EntraObjectId
                $aadCache[$azureId] = $info
            }
        } elseif ($primaryUpn) {
            $adDomain  = ($primaryUpn -split '@', 2)[1]
            $dirSource = if ($script:EnableDirectoryEnrichmentEffective) { 'AADOnly' } else { 'NotEnriched' }
        } else {
            $dirSource = if ($script:EnableDirectoryEnrichmentEffective) { $null } else { 'NotEnriched' }
        }

        # main row
        $rows.Add([pscustomobject]@{
            DeviceName                              = $dev.DeviceName
            AzureADDeviceId                         = $azureId
            EntraObjectId                           = $entraObjId
            Manufacturer                            = $dev.Manufacturer
            Model                                   = $dev.Model
            OperatingSystem                         = $dev.OperatingSystem
            LastSyncDateTime                        = $lastSync
            ComplianceState                         = $dev.ComplianceState
            ComplianceGracePeriodExpirationDateTime = $dev.ComplianceGracePeriodExpirationDateTime
            AD_Domain                               = $adDomain
            AD_OU                                   = $adOU
            DirectorySource                         = $dirSource
        })
        # Per-policy states
        if ($script:PolicyStateCollectionDisabled) {
            continue
        }

        $policyStates = $null
        try {
            Assert-PolicyStateRuntimeAvailable -Operation 'Process Intune compliance policy states'
            if ($policyStateBatchMap.ContainsKey([string]$dev.Id)) {
                $policyStates = @($policyStateBatchMap[[string]$dev.Id])
            }
            else {
                Assert-PolicyStateRuntimeAvailable -Operation 'Get Intune compliance policy states sequential fallback'
                $cmd = Get-Command -Name Get-MgDeviceManagementManagedDeviceDeviceCompliancePolicyState -ErrorAction SilentlyContinue
            if ($cmd) {
                $policyStates = Invoke-WithRetry -Script {
                    Get-MgDeviceManagementManagedDeviceDeviceCompliancePolicyState -ManagedDeviceId $dev.Id -All -ErrorAction Stop
                }
            } else {
                $uri  = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$($dev.Id)/deviceCompliancePolicyStates`?$top=200"
                $vals = @()
                $visitedPolicyStateUris = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
                while ($uri) {
                    if (-not $visitedPolicyStateUris.Add([string]$uri)) { throw 'Compliance policy-state pagination returned a repeated @odata.nextLink; collection is incomplete.' }
                    $resp = Invoke-WithRetry -Operation "Get Intune Graph page" -Script { Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop }
                    if (-not (Test-ComplianceGraphProperty -InputObject $resp -Name 'value')) { throw 'Compliance policy-state page returned an invalid Graph collection response without a value property.' }
                    $responseValues = @(Get-ComplianceGraphPropertyValue -InputObject $resp -Name 'value')
                    if ($responseValues.Count -gt 0) { $vals += $responseValues }
                    $uri = if (Test-ComplianceGraphProperty -InputObject $resp -Name '@odata.nextLink') { [string](Get-ComplianceGraphPropertyValue -InputObject $resp -Name '@odata.nextLink') } else { $null }
                }
                if ($vals.Count -gt 0) { $policyStates = $vals }
            }
            }
            $script:ConsecutivePolicyStateFailures = 0
        } catch {
            if (Test-PolicyStateCircuitBreakerException -ErrorRecord $_) {
                continue
            }
            $script:PolicyDetailCollectionComplete = $false
            $script:PolicyStateFailureCount++
            $script:ConsecutivePolicyStateFailures++
            $shortPolicyStateError = Get-ShortGraphErrorMessage -ErrorRecord $_
            Write-ComplianceWarning -Message ("Failed to retrieve policy states for device {0}: {1}" -f $dev.DeviceName, $shortPolicyStateError)
            if (($script:MaxPolicyStateFailures -gt 0 -and $script:PolicyStateFailureCount -ge $script:MaxPolicyStateFailures) -or
                ($script:MaxConsecutivePolicyStateFailures -gt 0 -and $script:ConsecutivePolicyStateFailures -ge $script:MaxConsecutivePolicyStateFailures)) {
                $script:PolicyStateCollectionDisabled = $true
                Write-ComplianceWarning -Message ("Policy state collection disabled for this run after {0} total failure(s), {1} consecutive. Device summary processing will continue." -f $script:PolicyStateFailureCount, $script:ConsecutivePolicyStateFailures)
            }
        }
        if ($policyStates) {
            $policyStates = $policyStates |
                Where-Object { $_.platformType -eq 'windows10AndLater' } |
                Sort-Object displayName, version -Unique

            # Build category rollup per policy
            $policyCategoryRollup = @{}

            if ($IncludeComplianceSettings) {
                $targets = $policyStates | Where-Object {
                    $_.nonCompliantSettingCount -gt 0 -or $_.state -eq 'nonCompliant'
                }

                $pIdx = 0
                $pTot = ($targets | Measure-Object).Count
                $settingStateBatchMap = @{}
                if ($pTot -gt 0) {
                    try {
                        Assert-PolicyStateRuntimeAvailable -Operation 'Get Intune compliance setting states batch'
                        $settingStateBatchMap = Get-ComplianceSettingStateBatchMap -ManagedDeviceId ([string]$dev.Id) -Policies @($targets)
                    }
                    catch {
                        if (Test-PolicyStateCircuitBreakerException -ErrorRecord $_) {
                            continue
                        }
                        Write-ComplianceWarning -Message ("Compliance setting-state batching failed for device '{0}'; sequential retrieval will be used: {1}" -f $dev.DeviceName, $_.Exception.Message)
                        $settingStateBatchMap = @{}
                    }
                }

                foreach ($p in $targets) {
                    $pIdx++
                    if ($pTot -gt 0) {
                        Write-Progress -Id 2 -ParentId 1 `
                            -Activity ("Policies for {0}" -f $dev.DeviceName) `
                            -Status ("{0}/{1} - {2}" -f $pIdx, $pTot, $p.displayName) `
                            -PercentComplete ([int](100 * $pIdx / $pTot))
                    }

                    try {
                        Assert-PolicyStateRuntimeAvailable -Operation 'Get Intune compliance setting states sequential fallback'
                        $s = @()
                        if ($settingStateBatchMap.ContainsKey([string]$p.id)) {
                            $s = @($settingStateBatchMap[[string]$p.id])
                        }
                        else {
                            $u = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$($dev.Id)/deviceCompliancePolicyStates/$([uri]::EscapeDataString($p.id))/settingStates`?$select=setting,state&`$top=200"
                        $visitedSettingStateUris = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
                        while ($u) {
                            if (-not $visitedSettingStateUris.Add([string]$u)) { throw 'Compliance setting-state pagination returned a repeated @odata.nextLink; collection is incomplete.' }
                            $page = Invoke-WithRetry -Operation "Get Intune compliance setting states" -Script { Invoke-MgGraphRequest -Method GET -Uri $u -ErrorAction Stop }
                            if (-not (Test-ComplianceGraphProperty -InputObject $page -Name 'value')) { throw 'Compliance setting-state page returned an invalid Graph collection response without a value property.' }
                            $pageValues = @(Get-ComplianceGraphPropertyValue -InputObject $page -Name 'value')
                            if ($pageValues.Count -gt 0) { $s += $pageValues }
                            $u = if (Test-ComplianceGraphProperty -InputObject $page -Name '@odata.nextLink') { [string](Get-ComplianceGraphPropertyValue -InputObject $page -Name '@odata.nextLink') } else { $null }
                        }
                        }

                        foreach ($v in ($s | Where-Object { $_.state -eq 'nonCompliant' })) {
                            $cat = Map-SettingCategory -SettingName $v.setting
                            if (-not $policyCategoryRollup.ContainsKey($p.displayName)) {
                                $policyCategoryRollup[$p.displayName] = @{}
                            }
                            $policyCategoryRollup[$p.displayName][$cat] = 'Fail'
                        }
                    } catch {
                        if (Test-PolicyStateCircuitBreakerException -ErrorRecord $_) {
                            break
                        }
                        $script:PolicyDetailCollectionComplete = $false
                        Write-ComplianceWarning -Message ("Failed to retrieve setting states for '{0}' on device '{1}': {2}" -f $p.displayName, $dev.DeviceName, $_.Exception.Message)
                    }
                }

                Write-Progress -Id 2 -ParentId 1 -Activity ("Policies for {0}" -f $dev.DeviceName) -Completed
            }

            # Emit per-policy rows (ensure all columns even if empty)
            foreach ($p in $policyStates) {
                # Resolve configured categories for this policy (cached)
                $policyId = Get-SafeProperty -Object $p -Name 'id'
                $configuredCats = if ($policyId) {
                    Get-PolicyConfiguredCategories -PolicyId $policyId
                } else {
                    [System.Collections.Generic.HashSet[string]]::new()
                }
                # Defensive null-guard: function may return $null in edge cases
                if ($null -eq $configuredCats) {
                    $configuredCats = [System.Collections.Generic.HashSet[string]]::new()
                }

                # Helper: resolve column value
                # - '' if category not configured in policy
                # - '' if state is error/unknown/notApplicable (indeterminate)
                # - 'Fail' if category is in nonCompliant rollup
                # - 'Pass' if category is configured and state is compliant/nonCompliant but not failed
                $determinable = $p.state -in @('compliant','nonCompliant')
                $resolveCol = {
                    param([string]$Cat)
                    if ($null -eq $configuredCats -or -not $configuredCats.Contains($Cat)) { return '' }
                    if (-not $determinable) { return '' }
                    $pName = $p.displayName
                    if ([string]::IsNullOrEmpty($pName)) { return '' }
                    if ($policyCategoryRollup.ContainsKey($pName) -and $policyCategoryRollup[$pName].ContainsKey($Cat)) {
                        return 'Fail'
                    }
                    return 'Pass'
                }
                $polAll.Add([pscustomobject]@{
                    DeviceName               = $dev.DeviceName
                    AzureADDeviceId          = $azureId
                    EntraObjectId            = $entraObjId
                    displayName              = $p.displayName
                    state                    = $p.state
                    version                  = $p.version
                    platformType             = $p.platformType
                    settingCount             = $p.settingCount
                    nonCompliantSettingCount = if ($p.nonCompliantSettingCount -ne $null) { $p.nonCompliantSettingCount } else { 0 }
                    lastReportedDateTime     = & {
                        $raw = Get-SafeProperty -Object $p -Name 'lastReportedDateTime'
                        if ([string]::IsNullOrWhiteSpace($raw)) { return '' }
                        $parsed = [datetime]::MinValue
                        if ([datetime]::TryParse($raw, [ref]$parsed) -and $parsed.Year -gt 1) {
                            $parsed.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
                        } else {
                            ''
                        }
                    }
                    SecureBoot               = & $resolveCol 'SecureBoot'
                    'BitLocker/Encryption'   = & $resolveCol 'BitLocker'
                    TPM                      = & $resolveCol 'TPM'
                    'Defender/Antivirus'     = & $resolveCol 'Antivirus'
                    Firewall                 = & $resolveCol 'Firewall'
                    CodeIntegrity            = & $resolveCol 'CodeIntegrity'
                    OSVersion                = & $resolveCol 'OSVersion'
                    UEFI                     = & $resolveCol 'UEFI'
                    AD_Domain                = $adDomain
                    AD_OU                    = $adOU
                    DirectorySource          = $dirSource
                })
            }
        }
    }

    Write-Progress -Id 1 -Activity "Processing devices" -Completed

    if ($script:PolicyStateCollectionStartedAt) {
        $policyStateElapsed = (Get-Date) - $script:PolicyStateCollectionStartedAt
        Write-ComplianceInfo -Message ("Detailed compliance policy-state collection finished after {0}; batch retries={1}; throttled sub-responses={2}; circuit breaker={3}." -f $policyStateElapsed.ToString('hh\:mm\:ss'), $script:PolicyStateBatchRetryCount, $script:PolicyStateBatchThrottleCount, $script:PolicyStateCircuitBreakerLogged)
    }

    if ($script:SettingBatchFallbackCounts.Count -gt 0) {
        $fallbackTotal = ($script:SettingBatchFallbackCounts.Values | Measure-Object -Sum).Sum
        $fallbackSummary = (@($script:SettingBatchFallbackCounts.GetEnumerator() |
            Sort-Object Name |
            ForEach-Object { "HTTP $($_.Name)=$($_.Value)" }) -join ', ')
        Write-ComplianceWarning -Message ("{0} compliance setting-state batch sub-request(s) required sequential fallback ({1}). Sample managed-device/policy IDs: {2}" -f
            $fallbackTotal, $fallbackSummary, ($script:SettingBatchFallbackExamples -join ', '))
    }

    # 3) Output
    if ($rows.Count -gt 0) {
        Write-ComplianceInfo -Message ("Devices compliance summary: {0} row(s)" -f $rows.Count)

        $summaryOut = $rows |
            Sort-Object DeviceName |
            Select-Object `
                DeviceName, AzureADDeviceId, EntraObjectId, Manufacturer, Model, OperatingSystem, LastSyncDateTime, `
                ComplianceState, ComplianceGracePeriodExpirationDateTime, `
                AD_Domain, AD_OU, DirectorySource

        $summaryOut | Select-Object -First 25 | Format-Table -AutoSize
        if ($summaryOut.Count -gt 25) {
            Write-ComplianceInfo -Message ("Displayed first 25 of {0} device summary rows." -f $summaryOut.Count)
        }

        try {
            Write-SmartM365CsvAtomically -Data @($summaryOut) -Path $mainCsv
            Export-SmartM365Csv -Data @($summaryOut) -TimestampedPath $tsCsv -LatestPath $lastCsv | Out-Null

            WriteLog -Message "Compliance summary CSV saved: $mainCsv" -Level 'SUCCESS'
        } catch {
            Write-ComplianceWarning -Message "Failed to export compliance summary CSVs: $_"
            throw
        }
    } else {
        Write-ComplianceWarning -Message "No devices to display or export."
    }

    if (-not $script:IncludePolicyStatesEffective) {
        Write-ComplianceInfo -Message "Compliance policy detail collection disabled by configuration, automatic large-tenant safeguard, or parameter. Set IncludePolicyStates to true explicitly to generate the detailed per-policy CSV."
    } elseif (-not $script:PolicyDetailCollectionComplete) {
        Write-ComplianceWarning -Message 'Compliance policy detail collection was incomplete. Detailed policy CSV publication is skipped so the last valid DATA-LAST export is preserved.'
    } elseif ($polAll.Count -gt 0) {
        Write-ComplianceInfo -Message ("Compliance details per policy: {0} row(s)" -f $polAll.Count)

        $polOut = $polAll |
            Sort-Object DeviceName, displayName, version |
            Select-Object `
                DeviceName, AzureADDeviceId, EntraObjectId, displayName, state, version, platformType, `
                settingCount, nonCompliantSettingCount, lastReportedDateTime, `
                SecureBoot, 'BitLocker/Encryption', TPM, 'Defender/Antivirus', `
                Firewall, CodeIntegrity, OSVersion, UEFI, `
                AD_Domain, AD_OU, DirectorySource

        $polOut | Select-Object -First 50 | Format-Table -AutoSize -Wrap
        if ($polOut.Count -gt 50) {
            Write-ComplianceInfo -Message ("Displayed first 50 of {0} policy detail rows." -f $polOut.Count)
        }

        try {
            Write-SmartM365CsvAtomically -Data @($polOut) -Path $policyMainCsv
            Export-SmartM365Csv -Data @($polOut) -TimestampedPath $policyTsCsv -LatestPath $policyLastCsv | Out-Null

            WriteLog -Message "Compliance policy CSV saved: $policyMainCsv" -Level 'SUCCESS'
        } catch {
            Write-ComplianceWarning -Message "Failed to export compliance policy CSVs: $_"
            throw
        }
    } else {
        Write-ComplianceWarning -Message "Compliance details: no policy states available or calls failed."
    }
}
catch {
    $script:ComplianceFatalError = $_
    WriteLog -Message "A global error occurred in SmartM365-Devices-Compliance-Inventory.ps1: $($_.Exception.Message)" -Level 'ERROR'
    $global:SmartM365ErrorCount = [Math]::Max(1, [int]$global:SmartM365ErrorCount)
    Write-Error $script:ComplianceFatalError
}
finally {
    # Disconnect Graph only if we connected it in this run
    if ($connectedGraphInThisRun) {
        Write-ComplianceInfo -Message 'Disconnecting cloud services.'
        try {
            Disconnect-SmartM365CloudSession -ExchangeOnline:$false -Graph:$true -VerboseDisconnect:$true
        } catch {
            Write-ComplianceWarning -Message ("Error during Graph disconnect in finally: {0}" -f $_)
        }
    }

    try {
        Stop-Transcript | Out-Null
        try {
            if ($transcriptPath) {
                Update-SmartM365TimestampedTranscript -Path $transcriptPath
            }
        } catch { }
    } catch { }

    try {
        $finalStatus = if ($script:ComplianceFatalError) { 'Failed' } else { 'Auto' }
        Complete-SmartM365ExecutionContext -Status $finalStatus -ErrorRecord $script:ComplianceFatalError -FailureStage 'ComplianceInventory'
    } catch { }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAsPG4NilpQ6x67
# lzg4wAzcrDzJua2Cny3f7usZN5i+4KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIB0OzgH6RGRrOi6bVTjoaQ5f27PmcH07CAk+jwH65o2HMA0GCSqG
# SIb3DQEBAQUABIIBgJvPiEaDngfmSStBYU0TCYKoBOSC4usNTePN1LMwagsJw8On
# 1zZ6DBbmNa2dTl69dKB2T/GZVzevnrie2ZZMxYv3+Nb//FlUIcGecsm2Pe4QV3LV
# 37dxnpAbQeDfAP2rSSEY1NaBiQzYM2COjMYUw1SlCxA5gxGG6EAqdSyhEs6EinmK
# aFs4aM9MWUgZazL7IowZJzkPDefOSd7SVpEfIbi+huYcK4LSEdk1qPrHH+rbpTa5
# YTr7RswLQYwlOfAttnb4BR6m8aAK+BA/WAE+XkkiC0EMJl0pR58hlZ1+TTZJP/oi
# nKJn+w50MsMIPdV7FAPV8Z0hRBPJ/8h+54d9gVXAeNLK7LPW0cHuLQz0q8f8HAwJ
# OLVI3rHQTEGEaSTLPHczHvDqpfIrRgxUCJQw93A+UQe1IhhnaH3GagamhpVlpqum
# 4XgQPhhXnvKXqlsZMwoj1WxRvWOfYoo6KnVcyrWFA1TvXKlXX1YtoUJf8mX0sGTg
# dBB+UEc8OZxLtV4A2KGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjEyMjI1
# MDVaMC8GCSqGSIb3DQEJBDEiBCCHJIGdACpBXpJEh07Qjn9iW521/Nm8bkjOiRRb
# MxIqnjANBgkqhkiG9w0BAQEFAASCAgAHbKw1yR46TcIwUBFe0B/vOqskDrSmzEds
# Im8EJscm69AZiZ2QnmyRqWm8mtlx2+pYQAVjuhE8k/i948iEdvYkSq5SfMEtXJnh
# YYCcPxcysnILek39HmE6nD9g8oSqyNDbq26mTRzHl0DhUXcl8+Mf9O+5xqh9hP8D
# d0nB+nido9em2dMp+iWktfVN/sf0479MACE5yLR1BlUcBgXy4dE/f/a+8e5xS3UN
# x35g/jt3jctHLZU4pt7SGu32dP6ec/YqlTd4uOnNlxRq9uGlFRyxgmq8zdOps6Gt
# tL62284udcE8Z0auFxqNzW4YW1JCDeERe0w5BxBU/deVWta1JqCPLLYrDXUiIkCB
# yKvqBYDUw6niXVL6d+OOmSkNvr7D0cl3Tvsky+0LFmB4p25fRj35T2pi8hArq5W8
# ZXy41lv/k9CdWexatFZKj2HceMp9PMO5rT+Abyvt2K3Z7H7NdO9dli0/QDblsofB
# kRIbBJRtIfiVVYPDm8ei84BXYnAR0sVbgWg9kIGGZxF8Qnar/jTEaQICnkYTYnEQ
# idB8TtlFh8SMyhFia4z+bl+ZNoZDDUzoYxit6/CuhxDsOCZLu/invtzzh0A0+pan
# hIqGvMx8rTaCSCd6h/C87BYhdhwvyK3BM9o9pSSuf9eDfTDs/sTH5ZCplgMX0khu
# ahN4IS9sIA==
# SIG # End signature block
