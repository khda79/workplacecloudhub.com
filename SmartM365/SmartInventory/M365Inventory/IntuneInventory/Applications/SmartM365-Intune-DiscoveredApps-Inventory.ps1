<#
.SYNOPSIS
    Intune-DiscoveredApps-Inventory
    Retrieves discovered Windows applications from Intune via Microsoft Graph API.

.DESCRIPTION
    Connects to Microsoft Graph using service principal with certificate authentication.
    Retrieves all discovered Windows applications and their associated managed devices.
    Produces two CSV exports:
      - Summary      : one row per discovered Windows app (name, version, publisher, device count)
      - DeviceDetail : one row per app / device pair (flat join)
    Both files are written to the DATA-ALL output folder (DiscoveredAppsCsvLogFolderPath) and copied to DATA-LAST (LatestCsvFolderPath).

.PARAMETER OutputPath
    Overrides the default output path (local configuration DiscoveredAppsCsvLogFolderPath).
    If omitted, the path is resolved from the local configuration (DiscoveredAppsCsvLogFolderPath).

.PARAMETER Connect
    Forces a (re)connection to Microsoft Graph (disconnects any existing session first).

.PARAMETER MaxApps
    Optional. Limits the number of Windows apps processed (0 = no limit).
    Use a small value (e.g. 10) for testing before a full run.

.PARAMETER DryRun
    If specified, collects and logs data but does not write CSV files.

.PARAMETER DelayMs
    Milliseconds to wait between each managedDevices Graph call to avoid throttling.
    Default: 300. Increase if 429 errors persist (e.g. 500 or 1000).
    Version : 1.15

.VERSION
1.15


.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication.
    Minimum Graph application permissions: DeviceManagementApps.Read.All; DeviceManagementManagedDevices.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Script  : Intune-DiscoveredApps-Inventory
    Version : 1.15
    Requires: Microsoft.Graph.Authentication module
              SmartM365.Core module (Modules\SmartM365.Core\SmartM365.Core.psd1)
    Local configuration: DiscoveredAppsCsvLogFolderPath -> output folder (DATA-ALL\M365-Inventory\Output-Windows-Discovered apps)
              LatestCsvFolderPath -> DATA-LAST folder (GlobalPath / copy destination)
    Graph permission required (application): DeviceManagementApps.Read.All
    Minimum application permissions: DeviceManagementApps.Read.All, DeviceManagementManagedDevices.Read.All
#>

param(
    [string]$Tenant = 'test',
[string]$OutputPath,
    [switch]$Connect,
    [ValidateRange(0, [int]::MaxValue)]
    [int]$MaxApps = 0,
    [switch]$DryRun,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, 5000)]
    [int]$DelayMs = 300,

    [Parameter(Mandatory = $false)]
    [ValidateSet('All', 'None', 'NonZero', 'Top')]
    [string]$DeviceDetailMode = 'All',

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$TopAppsByDeviceCount = 500,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, [int]::MaxValue)]
    [int]$ProgressEveryApps = 250,

    [Parameter(Mandatory = $false)]
    [switch]$ResetResume,

    [Parameter(Mandatory = $false)]
    [switch]$RefreshDeviceDetailCache,

    [Parameter(Mandatory = $false)]
    [ValidateRange(0, [int]::MaxValue)]
    [int]$DeviceDetailCacheMaxAgeDays = 7,
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

# ==========================================================
# Import SmartM365.Core module
# ==========================================================
$modulePath = & { $d = $PSScriptRoot; while ($d) { $p = Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'; if (Test-Path -LiteralPath $p) { return $p }; $parent = Split-Path -Path $d -Parent; if ($parent -eq $d) { break }; $d = $parent }; throw 'SmartM365.Core module not found.' }
try {
    Import-Module -Name $modulePath -MinimumVersion '1.0.24' -ErrorAction Stop
} catch {
    Write-Host "Failed to import SmartM365.Core module from '$modulePath': $_" -ForegroundColor Red
    exit 1
}

# ==========================================================
# Script metadata
# ==========================================================
$ScriptVersion = "1.15"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion"
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DiscoveredAppsCsvLogFolderPath' -DefaultValue $OutputPath
if (-not $PSBoundParameters.ContainsKey('DelayMs')) {
    $DelayMs = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'GraphRequestDelayMs' -DefaultValue 100)
}
if (-not $PSBoundParameters.ContainsKey('DeviceDetailMode')) {
    $DeviceDetailMode = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DeviceDetailMode' -DefaultValue 'All')
}
if ($DeviceDetailMode -notin @('All', 'None', 'NonZero', 'Top')) {
    throw "Invalid DeviceDetailMode '$DeviceDetailMode'. Valid values: All, None, NonZero, Top."
}
if (-not $PSBoundParameters.ContainsKey('TopAppsByDeviceCount')) {
    $TopAppsByDeviceCount = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'TopAppsByDeviceCount' -DefaultValue 500)
}
if ($TopAppsByDeviceCount -lt 1) {
    throw "TopAppsByDeviceCount must be greater than 0."
}
if (-not $PSBoundParameters.ContainsKey('ProgressEveryApps')) {
    $ProgressEveryApps = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ProgressEveryApps' -DefaultValue 250)
}
if (-not $PSBoundParameters.ContainsKey('DeviceDetailCacheMaxAgeDays')) {
    $DeviceDetailCacheMaxAgeDays = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DeviceDetailCacheMaxAgeDays' -DefaultValue 7)
}
$script:UsePreviousDeviceDetailCache = -not $RefreshDeviceDetailCache
$script:GraphMaxRetryAttempts = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'GraphMaxRetryAttempts' -DefaultValue 8)
$script:GraphRetryDefaultSeconds = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'GraphRetryDefaultSeconds' -DefaultValue 30)
$script:GraphRetryMaxSeconds = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'GraphRetryMaxSeconds' -DefaultValue 300)
# Statistics
$script:Stat_AppsTotal        = 0
$script:Stat_AppsWindows      = 0
$script:Stat_AppsProcessed    = 0
$script:Stat_AppsSkipped      = 0
$script:Stat_DeviceDetailRows = 0
$script:Stat_DetailAppsSkippedByResume = 0
$script:Stat_DetailAppsFromCache = 0
$script:Stat_DeviceDetailRowsFromCache = 0
$script:DeviceDetailResumePath = ''
$script:DeviceDetailPartialPath = ''
$script:DeviceDetailTimestampedPath = ''
$script:DeviceDetailCompleted = $false
$script:Stat_DetailAppsTargeted = 0
$script:Stat_GraphCalls       = 0
$script:Stat_ThrottleRetries  = 0
$script:DeviceDetailResumeContractVersion = 2

# ==========================================================
# Initialize script environment
# ==========================================================
try {
    $InitializeOutputPath = InitializeScriptEnvironment `
        -OutputPathInit $OutputPath `
        -LogFileName    ($MyInvocation.MyCommand.Name -replace '\.ps1$', '')
    Start-Transcript -Path $global:logTranscriptFile -Append

    WriteLog -Message "Script environment initialized at $InitializeOutputPath"
    $OutputPath = $InitializeOutputPath
    WriteLog -Message "Starting $TaskName..."
    WriteLog -Message "MaxApps            : $(if ($MaxApps -eq 0) { 'unlimited' } else { $MaxApps })"
    WriteLog -Message "DryRun             : $DryRun"
    WriteLog -Message "DelayMs            : $DelayMs"
    WriteLog -Message "DeviceDetailMode   : $DeviceDetailMode"
    WriteLog -Message "TopAppsByDeviceCount: $TopAppsByDeviceCount"
    WriteLog -Message "ProgressEveryApps  : $ProgressEveryApps"
    WriteLog -Message "Use previous DeviceDetail cache: $script:UsePreviousDeviceDetailCache"
    WriteLog -Message "DeviceDetail cache max age days: $DeviceDetailCacheMaxAgeDays"
    WriteLog -Message "RefreshDeviceDetailCache: $RefreshDeviceDetailCache"
    WriteLog -Message "ResetResume        : $ResetResume"
} catch {
    Write-Host "Initialization failed: $_" -ForegroundColor Red
    exit 1
}

# ==========================================================
# Helpers
# ==========================================================

# Graph connectivity check
function Test-GraphConnection {
    try {
        Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

# Paginated Graph request with transient retry
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
        [int]$DefaultSeconds = 30,
        [int]$MaximumSeconds = 300
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

    $backoff = [math]::Min($MaximumSeconds, $DefaultSeconds * [math]::Pow(2, [math]::Max(0, $Attempt - 1)))
    return [int]($backoff + (Get-Random -Minimum 0 -Maximum 5))
}

function Get-GraphBatchResponseRetryDelaySeconds {
    [CmdletBinding()]
    param(
        [AllowNull()]$Response,
        [int]$Attempt,
        [int]$DefaultSeconds = 30,
        [int]$MaximumSeconds = 300
    )

    $retryAfter = $null
    try {
        $headers = $Response.headers
        if ($headers -is [System.Collections.IDictionary]) {
            foreach ($key in $headers.Keys) {
                if ([string]$key -ieq 'Retry-After') {
                    $retryAfter = $headers[$key]
                    break
                }
            }
        }
        elseif ($headers) {
            $headerProperty = $headers.PSObject.Properties['Retry-After']
            if ($headerProperty) { $retryAfter = $headerProperty.Value }
        }
    } catch {}

    $seconds = 0
    if ($retryAfter -and [int]::TryParse([string]$retryAfter, [ref]$seconds) -and $seconds -gt 0) {
        return [math]::Min($seconds, $MaximumSeconds)
    }

    $backoff = [math]::Min($MaximumSeconds, $DefaultSeconds * [math]::Pow(2, [math]::Max(0, $Attempt - 1)))
    return [int]($backoff + (Get-Random -Minimum 0 -Maximum 5))
}
function Invoke-GraphPagedRequest {
    param(
        [Parameter(Mandatory = $true)] [string]$InitialUri,
        [Parameter(Mandatory = $false)] [int]$MaxRetries = $script:GraphMaxRetryAttempts,
        [Parameter(Mandatory = $false)] [int]$DefaultRetrySeconds = $script:GraphRetryDefaultSeconds
    )

    if ($MaxRetries -lt 1) { $MaxRetries = 1 }
    $allItems = [System.Collections.Generic.List[psobject]]::new()
    $currentUri = $InitialUri

    while ($null -ne $currentUri) {
        $success = $false

        for ($attempt = 1; -not $success -and $attempt -le $MaxRetries; $attempt++) {
            try {
                $script:Stat_GraphCalls++
                $response = Invoke-MgGraphRequest -Method GET -Uri $currentUri -OutputType PSObject -ErrorAction Stop

                if ($null -ne $response.value) {
                    foreach ($item in $response.value) { $allItems.Add($item) }
                }

                $currentUri = if ($response.PSObject.Properties.Name -contains '@odata.nextLink') {
                    $response.'@odata.nextLink'
                } else { $null }

                $success = $true
            } catch {
                $statusCode = $null
                try { if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode } } catch {}
                $message = Get-ShortGraphErrorMessage -ErrorRecord $_
                $isTransient = $statusCode -in @(429, 500, 502, 503, 504) -or $message -match 'TooManyRequests|throttl|timeout|temporarily unavailable|InternalServerError'

                if (-not $isTransient -or $attempt -ge $MaxRetries) {
                    $statusText = if ($statusCode) { $statusCode } else { 'unknown' }
                    throw ("Graph request failed. Status={0}; Attempts={1}; Uri={2}; Message={3}" -f $statusText, $attempt, $currentUri, $message)
                }

                $retryAfter = Get-GraphRetryDelaySeconds -ErrorRecord $_ -Attempt $attempt -DefaultSeconds $DefaultRetrySeconds -MaximumSeconds $script:GraphRetryMaxSeconds
                $script:Stat_ThrottleRetries++
                $statusRetryText = if ($statusCode) { $statusCode } else { 'unknown' }
                WriteLog -Message ("Graph transient failure on [$currentUri]. Status={0}; attempt {1}/{2}; waiting {3}s." -f $statusRetryText, $attempt, $MaxRetries, $retryAfter) "INFO"
                Start-Sleep -Seconds $retryAfter
            }
        }
    }

    return $allItems
}

function Stop-DiscoveredAppsTranscript {
    [CmdletBinding()]
    param()

    try {
        if ($global:logTranscriptFile -and (Test-Path -LiteralPath $global:logTranscriptFile)) {
            Stop-Transcript | Out-Null
        }
    } catch {}
}

function Get-DiscoveredAppDeviceRelationBatchMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Apps,
        [ValidateRange(0, 5000)][int]$DelayMs = 0,
        [int]$MaxRetries = $script:GraphMaxRetryAttempts
    )

    if ($MaxRetries -lt 1) { $MaxRetries = 1 }
    $result = @{}
    $batchUri = "https://graph.microsoft.com/v1.0/" + '$batch'
    for ($offset = 0; $offset -lt $Apps.Count; $offset += 20) {
        if ($offset -gt 0 -and $DelayMs -gt 0) {
            Start-Sleep -Milliseconds $DelayMs
        }
        $last = [math]::Min($offset + 19, $Apps.Count - 1)
        $pendingApps = @($Apps[$offset..$last])

        for ($attempt = 1; $pendingApps.Count -gt 0 -and $attempt -le $MaxRetries; $attempt++) {
            $requests = [System.Collections.Generic.List[object]]::new()
            $requestAppMap = @{}
            $requestId = 1

            foreach ($app in $pendingApps) {
                $localId = [string]$requestId
                $appId = [string]$app.id
                $requestAppMap[$localId] = $app
                [void]$requests.Add(@{
                    id = $localId
                    method = 'GET'
                    url = "/deviceManagement/detectedApps/$appId/managedDevices?`$top=999&`$select=id"
                })
                $requestId++
            }

            $body = @{ requests = $requests } | ConvertTo-Json -Depth 6
            try {
                $script:Stat_GraphCalls++
                $batchResponse = Invoke-MgGraphRequest -Method POST -Uri $batchUri -Body $body -ContentType 'application/json' -OutputType PSObject -ErrorAction Stop
            }
            catch {
                $statusCode = $null
                try { if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode } } catch {}
                $message = Get-ShortGraphErrorMessage -ErrorRecord $_
                $isTransient = $statusCode -in @(408, 429, 500, 502, 503, 504) -or $message -match 'TooManyRequests|throttl|timeout|temporarily unavailable|InternalServerError'
                if ($isTransient -and $attempt -lt $MaxRetries) {
                    $retryAfter = Get-GraphRetryDelaySeconds -ErrorRecord $_ -Attempt $attempt -DefaultSeconds $script:GraphRetryDefaultSeconds -MaximumSeconds $script:GraphRetryMaxSeconds
                    $script:Stat_ThrottleRetries++
                    WriteLog -Message ("Discovered Apps relation batch transient failure at app offset {0}; attempt {1}/{2}; retrying {3} app(s) in {4}s." -f $offset, $attempt, $MaxRetries, $pendingApps.Count, $retryAfter) 'INFO'
                    Start-Sleep -Seconds $retryAfter
                    continue
                }

                WriteLog -Message ("Discovered Apps relation batch failed at app offset {0} after {1} attempt(s); sequential fallback will be used for {2} app(s): {3}" -f $offset, $attempt, $pendingApps.Count, $message) 'WARNING'
                break
            }

            $responsesById = @{}
            foreach ($response in @($batchResponse.responses)) {
                $responsesById[[string]$response.id] = $response
            }

            $retryApps = @()
            $retryDelay = 0
            foreach ($requestIdKey in @($requestAppMap.Keys)) {
                $app = $requestAppMap[$requestIdKey]
                $appId = [string]$app.id
                $response = $responsesById[$requestIdKey]

                if (-not $response) {
                    if ($attempt -lt $MaxRetries) {
                        $retryApps += $app
                        $candidateDelay = Get-GraphBatchResponseRetryDelaySeconds -Response $null -Attempt $attempt -DefaultSeconds $script:GraphRetryDefaultSeconds -MaximumSeconds $script:GraphRetryMaxSeconds
                        $retryDelay = [math]::Max($retryDelay, $candidateDelay)
                    }
                    else {
                        WriteLog -Message ("Discovered Apps relation batch returned no sub-response: AppId={0}. Sequential fallback will be used." -f $appId) 'WARNING'
                    }
                    continue
                }

                $statusCode = [int]$response.status
                if ($statusCode -eq 200) {
                    $devices = [System.Collections.Generic.List[object]]::new()
                    foreach ($device in @($response.body.value)) { if ($null -ne $device) { [void]$devices.Add($device) } }
                    $nextLink = [string]$response.body.'@odata.nextLink'
                    if (-not [string]::IsNullOrWhiteSpace($nextLink)) {
                        foreach ($device in @(Invoke-GraphPagedRequest -InitialUri $nextLink)) { [void]$devices.Add($device) }
                    }
                    $result[$appId] = @($devices)
                    continue
                }

                $isTransient = $statusCode -in @(408, 429, 500, 502, 503, 504)
                if ($isTransient -and $attempt -lt $MaxRetries) {
                    $retryApps += $app
                    $candidateDelay = Get-GraphBatchResponseRetryDelaySeconds -Response $response -Attempt $attempt -DefaultSeconds $script:GraphRetryDefaultSeconds -MaximumSeconds $script:GraphRetryMaxSeconds
                    $retryDelay = [math]::Max($retryDelay, $candidateDelay)
                    continue
                }

                WriteLog -Message ("Discovered Apps relation batch sub-request failed: AppId={0}; HTTP={1}; Attempts={2}. Sequential fallback will be used." -f $appId, $statusCode, $attempt) 'WARNING'
            }

            if ($retryApps.Count -gt 0) {
                $script:Stat_ThrottleRetries++
                WriteLog -Message ("Discovered Apps relation batch transient sub-request(s): Apps={0}; attempt {1}/{2}; retrying in {3}s." -f $retryApps.Count, $attempt, $MaxRetries, $retryDelay) 'INFO'
                if ($retryDelay -gt 0) { Start-Sleep -Seconds $retryDelay }
                $pendingApps = @($retryApps)
            }
            else {
                $pendingApps = @()
            }
        }
    }

    return $result
}
function New-DiscoveredAppsDeviceDetailRecord {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][object]$App,
        [Parameter(Mandatory = $false)][object]$Device
    )

    [pscustomobject]@{
        AppId                  = $App.id
        AppName                = $App.displayName
        AppVersion             = $App.version
        AppPublisher           = $App.publisher
        Platform               = $App.platform
        DeviceId               = if ($Device) { $Device.id } else { '' }
        DeviceName             = if ($Device) { $Device.deviceName } else { '' }
        OperatingSystem        = if ($Device) { $Device.operatingSystem } else { '' }
        OSVersion              = if ($Device) { $Device.osVersion } else { '' }
        UserPrincipalName      = if ($Device) { $Device.userPrincipalName } else { '' }
        LastSyncDateTime       = if ($Device) { $Device.lastSyncDateTime } else { '' }
        EnrolledDateTime       = if ($Device) { $Device.enrolledDateTime } else { '' }
        ManagedDeviceOwnerType = if ($Device) { $Device.managedDeviceOwnerType } else { '' }
        ComplianceState        = if ($Device) { $Device.complianceState } else { '' }
        AzureADDeviceId        = if ($Device) { $Device.azureADDeviceId } else { '' }
    }
}

function Write-DiscoveredAppsCsvRows {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][object[]]$Rows
    )

    if (-not $Rows -or $Rows.Count -eq 0) { return }
    $folder = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($folder) -and -not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    if (Test-Path -LiteralPath $Path) {
        Repair-SmartM365CsvTenantKeySchema -Path $Path -Delimiter ',' -Encoding UTF8 | Out-Null
        $Rows | Add-SmartM365TenantKey | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Append
    } else {
        $Rows | Add-SmartM365TenantKey | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
    }
}


function Get-DiscoveredAppsAppDeviceCount {
    [CmdletBinding()]
    param([AllowNull()]$App)

    $count = 0
    try {
        if ($App -and $null -ne $App.deviceCount -and [int]::TryParse([string]$App.deviceCount, [ref]$count)) {
            return $count
        }
    } catch {}
    return 0
}

function Get-DiscoveredAppsTargetAppIdsHash {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$AppIds)

    $normalizedIds = @($AppIds |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
        ForEach-Object { ([string]$_).Trim().ToLowerInvariant() } |
        Sort-Object -Unique)
    $payload = [System.Text.Encoding]::UTF8.GetBytes(($normalizedIds -join "`n"))
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        return ([System.BitConverter]::ToString($sha256.ComputeHash($payload))).Replace('-', '').ToLowerInvariant()
    }
    finally {
        $sha256.Dispose()
    }
}
function Test-DiscoveredAppsDeviceDetailCacheRowEnrichment {
    [CmdletBinding()]
    param([AllowNull()]$Row)

    if (-not $Row -or [string]::IsNullOrWhiteSpace([string]$Row.DeviceId)) {
        return $true
    }

    foreach ($propertyName in @('UserPrincipalName', 'AzureADDeviceId')) {
        if (-not [string]::IsNullOrWhiteSpace([string]$Row.$propertyName)) {
            return $true
        }
    }

    foreach ($propertyName in @('LastSyncDateTime', 'EnrolledDateTime')) {
        $parsedDate = [datetime]::MinValue
        if ([datetime]::TryParse([string]$Row.$propertyName, [ref]$parsedDate) -and $parsedDate.Year -gt 2000) {
            return $true
        }
    }

    foreach ($propertyName in @('ManagedDeviceOwnerType', 'ComplianceState')) {
        $value = [string]$Row.$propertyName
        if (-not [string]::IsNullOrWhiteSpace($value) -and $value -ine 'unknown') {
            return $true
        }
    }

    return $false
}

function Use-DiscoveredAppsDeviceDetailCache {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$CachePath,
        [Parameter(Mandatory = $true)][object[]]$TargetApps,
        [Parameter(Mandatory = $true)][string]$PartialPath,
        [Parameter(Mandatory = $true)]$ProcessedAppIds,
        [Parameter(Mandatory = $true)]$CachedAppIds,
        [int]$MaxAgeDays = 7
    )

    $result = [ordered]@{
        CachePath = $CachePath
        Used      = $false
        Apps      = 0
        Rows      = 0
        RejectedEnrichmentApps = 0
        Reason    = ''
    }

    if (-not (Test-Path -LiteralPath $CachePath)) {
        $result.Reason = 'cache file not found'
        return [pscustomobject]$result
    }

    $cacheItem = Get-Item -LiteralPath $CachePath -ErrorAction Stop
    if ($MaxAgeDays -gt 0 -and $cacheItem.LastWriteTime -lt (Get-Date).AddDays(-1 * $MaxAgeDays)) {
        $result.Reason = ("cache file older than {0} day(s): {1}" -f $MaxAgeDays, $cacheItem.LastWriteTime)
        return [pscustomobject]$result
    }

    $targetById = @{}
    foreach ($app in @($TargetApps)) {
        $id = [string]$app.id
        if (-not [string]::IsNullOrWhiteSpace($id) -and -not $targetById.ContainsKey($id)) {
            $targetById[$id] = $app
        }
    }
    if ($targetById.Count -eq 0) {
        $result.Reason = 'no target apps'
        return [pscustomobject]$result
    }

    $statsById = @{}
    try {
        Import-Csv -LiteralPath $CachePath | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.AppId) -and $targetById.ContainsKey([string]$_.AppId) } | ForEach-Object {
            $id = [string]$_.AppId

            if (-not $statsById.ContainsKey($id)) {
                $statsById[$id] = [pscustomobject]@{
                    AppId      = $id
                    AppName    = [string]$_.AppName
                    AppVersion = [string]$_.AppVersion
                    Publisher  = [string]$_.AppPublisher
                    Platform   = [string]$_.Platform
                    Rows       = 0
                    DeviceRows = 0
                    MetadataOk = $true
                    EnrichmentOk = $true
                }
            }

            $stat = $statsById[$id]
            $stat.Rows++
            if (-not [string]::IsNullOrWhiteSpace([string]$_.DeviceId)) {
                $stat.DeviceRows++
                if ($stat.EnrichmentOk -and -not (Test-DiscoveredAppsDeviceDetailCacheRowEnrichment -Row $_)) {
                    $stat.EnrichmentOk = $false
                }
            }

            $app = $targetById[$id]
            if ([string]$_.AppName -ne [string]$app.displayName -or
                [string]$_.AppVersion -ne [string]$app.version -or
                [string]$_.AppPublisher -ne [string]$app.publisher -or
                [string]$_.Platform -ne [string]$app.platform) {
                $stat.MetadataOk = $false
            }
        }
    }
    catch {
        $result.Reason = "failed to read cache stats: $_"
        return [pscustomobject]$result
    }

    $cacheable = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($id in @($statsById.Keys)) {
        $stat = $statsById[$id]
        $app = $targetById[$id]
        $expectedDeviceCount = Get-DiscoveredAppsAppDeviceCount -App $app
        $deviceCountOk = if ($expectedDeviceCount -eq 0) {
            $stat.Rows -gt 0 -and $stat.DeviceRows -eq 0
        } else {
            $stat.DeviceRows -eq $expectedDeviceCount
        }

        if ($stat.MetadataOk -and $deviceCountOk -and $stat.EnrichmentOk) {
            [void]$cacheable.Add($id)
        }
        elseif ($stat.MetadataOk -and $deviceCountOk -and -not $stat.EnrichmentOk) {
            $result.RejectedEnrichmentApps++
        }
    }

    if ($cacheable.Count -eq 0) {
        $result.Reason = if ($result.RejectedEnrichmentApps -gt 0) {
            "no cache rows meet the current DeviceDetail enrichment contract; rejected apps: $($result.RejectedEnrichmentApps)"
        }
        else {
            'no cache rows match current app metadata/device counts'
        }
        return [pscustomobject]$result
    }

    $batch = New-Object 'System.Collections.Generic.List[object]'
    try {
        Import-Csv -LiteralPath $CachePath | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.AppId) -and $cacheable.Contains([string]$_.AppId) } | ForEach-Object {
            $id = [string]$_.AppId

            $batch.Add($_) | Out-Null
            $result.Rows++
            if ($batch.Count -ge 5000) {
                Write-DiscoveredAppsCsvRows -Path $PartialPath -Rows $batch.ToArray()
                $batch.Clear()
            }
        }
        if ($batch.Count -gt 0) {
            Write-DiscoveredAppsCsvRows -Path $PartialPath -Rows $batch.ToArray()
            $batch.Clear()
        }
    }
    catch {
        $result.Reason = "failed to copy cache rows: $_"
        return [pscustomobject]$result
    }

    foreach ($id in @($cacheable)) {
        [void]$ProcessedAppIds.Add($id)
        [void]$CachedAppIds.Add($id)
    }

    $result.Used = $true
    $result.Apps = $cacheable.Count
    $result.Reason = 'cache rows reused'
    return [pscustomobject]$result
}

function Get-DiscoveredAppsResumeState {
    [CmdletBinding()]
    param([Parameter(Mandatory = $true)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try {
        return (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop)
    } catch {
        WriteLog -Message "Ignoring unreadable resume state '$Path': $_" "WARNING"
        return $null
    }
}

function Test-DiscoveredAppsResumeStateCompatible {
    [CmdletBinding()]
    param(
        [AllowNull()]$State,
        [Parameter(Mandatory = $true)][string]$Mode,
        [Parameter(Mandatory = $true)][int]$TargetCount,
        [Parameter(Mandatory = $true)][string]$TargetAppIdsHash,
        [Parameter(Mandatory = $true)][int]$ResumeContractVersion
    )

    if (-not $State) { return $false }
    $propertyNames = @($State.PSObject.Properties.Name)
    foreach ($requiredProperty in @('ResumeContractVersion','TargetAppIdsHash','TargetCount','DeviceDetailMode','PartialPath','TimestampedPath','ProcessedAppIds','ProcessedCount','SkippedCount')) {
        if ($propertyNames -notcontains $requiredProperty) { return $false }
    }

    if ([int]$State.ResumeContractVersion -ne $ResumeContractVersion -or
        [string]$State.TargetAppIdsHash -ne $TargetAppIdsHash -or
        [int]$State.TargetCount -ne $TargetCount -or
        [string]$State.DeviceDetailMode -ne $Mode -or
        [string]::IsNullOrWhiteSpace([string]$State.PartialPath) -or
        [string]::IsNullOrWhiteSpace([string]$State.TimestampedPath) -or
        -not (Test-Path -LiteralPath ([string]$State.PartialPath))) {
        return $false
    }

    $processedIds = @($State.ProcessedAppIds | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    $expectedProcessedIds = [int]$State.ProcessedCount + [int]$State.SkippedCount
    if ($processedIds.Count -ne $expectedProcessedIds -or @($processedIds | Sort-Object -Unique).Count -ne $processedIds.Count) {
        return $false
    }

    return $true
}
function Save-DiscoveredAppsResumeState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$PartialPath,
        [Parameter(Mandatory = $true)][string]$TimestampedPath,
        [Parameter(Mandatory = $true)][string]$Mode,
        [Parameter(Mandatory = $true)][int]$TargetCount,
        [Parameter(Mandatory = $true)][int]$ResumeContractVersion,
        [Parameter(Mandatory = $true)][string]$TargetAppIdsHash,
        [Parameter(Mandatory = $true)][string[]]$ProcessedAppIds,
        [Parameter(Mandatory = $true)][int]$ProcessedCount,
        [Parameter(Mandatory = $true)][int]$SkippedCount,
        [Parameter(Mandatory = $true)][int]$DetailRows
    )

    $state = [ordered]@{
        Script           = $TaskName
        DeviceDetailMode = $Mode
        TargetCount      = $TargetCount
        ResumeContractVersion = $ResumeContractVersion
        TargetAppIdsHash = $TargetAppIdsHash
        PartialPath      = $PartialPath
        TimestampedPath  = $TimestampedPath
        ProcessedAppIds  = @($ProcessedAppIds)
        ProcessedCount   = $ProcessedCount
        SkippedCount     = $SkippedCount
        DeviceDetailRows = $DetailRows
        Updated          = (Get-Date).ToString('o')
    }
    $folder = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($folder) -and -not (Test-Path -LiteralPath $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
    $state | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Complete-DiscoveredAppsStreamExport {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$PartialPath,
        [Parameter(Mandatory = $true)][string]$TimestampedPath,
        [Parameter(Mandatory = $true)][string]$OutputPath,
        [Parameter(Mandatory = $true)][string]$GlobalPath,
        [Parameter(Mandatory = $true)][string]$BaseFileName
    )

    if (-not (Test-Path -LiteralPath $PartialPath)) {
        WriteLog -Message "DeviceDetail partial CSV not found; no DeviceDetail CSV to finalize: $PartialPath" "WARNING"
        return $null
    }

    if ([string]::IsNullOrWhiteSpace($GlobalPath)) { $GlobalPath = $OutputPath }
    if (-not (Test-Path -LiteralPath $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
    if (-not (Test-Path -LiteralPath $GlobalPath)) { New-Item -ItemType Directory -Path $GlobalPath -Force | Out-Null }

    Copy-Item -LiteralPath $PartialPath -Destination $TimestampedPath -Force
    $localLatestPath = Join-Path -Path $OutputPath -ChildPath ("$BaseFileName.csv")
    $globalLatestPath = Join-Path -Path $GlobalPath -ChildPath ("$BaseFileName.csv")
    Copy-Item -LiteralPath $TimestampedPath -Destination $localLatestPath -Force
    Copy-Item -LiteralPath $TimestampedPath -Destination $globalLatestPath -Force

    if (-not $global:csvGeneratedPaths) {
        $global:csvGeneratedPaths = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }
    [void]$global:csvGeneratedPaths.Add($TimestampedPath)
    [void]$global:csvGeneratedPaths.Add($localLatestPath)
    [void]$global:csvGeneratedPaths.Add($globalLatestPath)

    if (Get-Command -Name Invoke-SmartM365SharePointCsvUpload -ErrorAction SilentlyContinue) {
        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $globalLatestPath | Out-Null
    }
    if (Get-Command -Name Invoke-SmartM365WeeklyInventoryHistoryForCsv -ErrorAction SilentlyContinue) {
        Invoke-SmartM365WeeklyInventoryHistoryForCsv -SourceFiles @($globalLatestPath) -TimestampedPath $TimestampedPath | Out-Null
    }

    Remove-Item -LiteralPath $PartialPath -Force -ErrorAction SilentlyContinue
    WriteLog -Message "DeviceDetail CSV finalized: $TimestampedPath" "INFO"
    return $TimestampedPath
}
# ==========================================================
# MAIN TRY / CATCH / FINALLY
# ==========================================================
$connectedGraphInThisRun = $false
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
$script:RunError = $null

try {

    # ----------------------------------------------------------
    # Connect to Microsoft Graph
    # ----------------------------------------------------------
    $graphContext = $null
    if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
        try { $graphContext = Get-MgContext -ErrorAction SilentlyContinue } catch {}
    }

    $needConnect = $false

    if ($Connect) {
        if ($graphContext) {
            WriteLog -Message "Connect switch specified: disconnecting the existing Graph session before reconnecting." "INFO"
            try { Disconnect-MgGraph -ErrorAction Stop | Out-Null } catch {
                WriteLog -Message ("Existing Graph session could not be disconnected cleanly before reconnecting: {0}" -f $_) "WARNING"
            }
        }
        else {
            WriteLog -Message "Connect switch specified: no existing Graph session to disconnect; establishing a new connection." "INFO"
        }
        $needConnect = $true
    } elseif ($graphContext -and (Test-GraphConnection)) {
        WriteLog -Message "Existing Microsoft Graph session detected. Reusing current connection." "INFO"
        $needConnect = $false
    } else {
        WriteLog -Message "No existing Graph session detected. Establishing a new connection..." "INFO"
        $needConnect = $true
    }

    if ($needConnect) {
        WriteLog -Message "Connecting to Microsoft Graph with app-only certificate authentication." "INFO"
        Connect-MgGraph -TenantId $TenantId -ClientId $AppId -CertificateThumbprint $Thumb -NoWelcome | Out-Null

        if (-not (Test-GraphConnection)) {
            throw "Failed to connect to Microsoft Graph."
        }
        $connectedGraphInThisRun = $true
        WriteLog -Message "Connected to Microsoft Graph successfully." "INFO"
    }

    # ----------------------------------------------------------
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -RequiredGraphApplicationPermissions @('DeviceManagementApps.Read.All','DeviceManagementManagedDevices.Read.All') -GraphProbeUris @('https://graph.microsoft.com/v1.0/deviceManagement/detectedApps?$top=1') | Out-Null

    # Retrieve ALL discovered apps - filter Windows client-side
    # Note: $filter on platform is not guaranteed on this endpoint;
    #       client-side filtering ensures compatibility.
    # ----------------------------------------------------------
    WriteLog -Message "Retrieving all discovered apps from Intune (Windows filter applied client-side)..." "INFO"
    $appsUri    = 'https://graph.microsoft.com/v1.0/deviceManagement/detectedApps?$top=999&$select=id,displayName,version,publisher,deviceCount,platform'
    $allAppsRaw = Invoke-GraphPagedRequest -InitialUri $appsUri

    $script:Stat_AppsTotal   = $allAppsRaw.Count
    $windowsApps             = @($allAppsRaw |
        Where-Object { $_.platform -eq 'windows' } |
        Group-Object -Property id |
        ForEach-Object { $_.Group | Select-Object -First 1 })
    $script:Stat_AppsWindows = $windowsApps.Count

    WriteLog -Message "Total apps retrieved (all platforms) : $($script:Stat_AppsTotal)" "INFO"
    WriteLog -Message "Windows apps after filter            : $($script:Stat_AppsWindows)" "INFO"

    if ($script:Stat_AppsWindows -eq 0) {
        WriteLog -Message "No Windows apps found. Exiting." "WARNING"
        return
    }

    # Apply MaxApps limit if set
    if ($MaxApps -gt 0 -and $windowsApps.Count -gt $MaxApps) {
        WriteLog -Message "MaxApps=$($MaxApps): limiting processing to first $($MaxApps) Windows apps." "WARNING"
        $windowsApps = $windowsApps | Select-Object -First $MaxApps
    }

    # ----------------------------------------------------------
    # Build and export Summary records before long DeviceDetail calls
    # ----------------------------------------------------------
    WriteLog -Message "Building Summary records..." "INFO"
    $summaryRecords = [System.Collections.Generic.List[psobject]]::new()
    foreach ($app in $windowsApps) {
        $summaryRecords.Add([pscustomobject]@{
            AppId        = $app.id
            AppName      = $app.displayName
            AppVersion   = $app.version
            AppPublisher = $app.publisher
            Platform     = $app.platform
            DeviceCount  = $app.deviceCount
        })
    }
    WriteLog -Message "Summary records built: $($summaryRecords.Count) rows." "INFO"

    $globalPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''
    if ($DryRun) {
        WriteLog -Message "DryRun enabled - Summary CSV export skipped." "INFO"
    } else {
        WriteLog -Message "Exporting Summary CSV ($($summaryRecords.Count) rows)..." "INFO"
        ExportAndCopyCsv `
            -BaseFileName "Intune_DiscoveredApps_Summary" `
            -OutputPath $OutputPath `
            -GlobalPath $globalPath `
            -Data $summaryRecords `
            -Encoding "UTF8" `
            -NoTypeInformation
        WriteLog -Message "Summary export completed: $global:csvFilePath1" "INFO"
    }

    # ----------------------------------------------------------
    # Build DeviceDetail records with streaming export and resume support
    # ----------------------------------------------------------
    $detailRecords = [System.Collections.Generic.List[psobject]]::new()
    $detailApps = switch ($DeviceDetailMode) {
        'None' { @() }
        'NonZero' { @($windowsApps | Where-Object { [int]($_.deviceCount) -gt 0 }) }
        'Top' {
            @($windowsApps |
                Sort-Object -Property @{ Expression = { [int]($_.deviceCount) }; Descending = $true }, displayName, version |
                Select-Object -First $TopAppsByDeviceCount)
        }
        default { @($windowsApps) }
    }

    $script:Stat_DetailAppsTargeted = $detailApps.Count
    WriteLog -Message ("Device detail mode '{0}' selected {1} app(s) out of {2} Windows apps." -f $DeviceDetailMode, $detailApps.Count, $windowsApps.Count) "INFO"

    if ($DeviceDetailMode -eq 'None') {
        WriteLog -Message "DeviceDetailMode=None; skipping per-app managedDevices retrieval." "INFO"
    } elseif ($detailApps.Count -eq 0) {
        WriteLog -Message "No applications selected for DeviceDetail export (mode=$DeviceDetailMode)." "INFO"
    } else {
        WriteLog -Message "Building DeviceDetail records for $($detailApps.Count) applications (mode=$DeviceDetailMode)..." "INFO"

        $detailBaseFileName = Add-SmartM365MaxItemsSuffixToBaseName -BaseFileName 'Intune_DiscoveredApps_DeviceDetail'
        $detailTargetAppIdsHash = Get-DiscoveredAppsTargetAppIdsHash -AppIds @($detailApps | ForEach-Object { [string]$_.id })
        $detailTimestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $script:DeviceDetailResumePath = Join-Path -Path $OutputPath -ChildPath "$detailBaseFileName.resume.json"
        $processedAppIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $cachedAppIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $streamingEnabled = -not $DryRun

        if ($streamingEnabled -and $ResetResume) {
            $oldState = Get-DiscoveredAppsResumeState -Path $script:DeviceDetailResumePath
            if ($oldState -and $oldState.PartialPath) { Remove-Item -LiteralPath ([string]$oldState.PartialPath) -Force -ErrorAction SilentlyContinue }
            if ($oldState -and $oldState.TimestampedPath) { Remove-Item -LiteralPath ([string]$oldState.TimestampedPath) -Force -ErrorAction SilentlyContinue }
            Remove-Item -LiteralPath $script:DeviceDetailResumePath -Force -ErrorAction SilentlyContinue
            WriteLog -Message "ResetResume enabled; previous DeviceDetail resume state removed." "INFO"
        }

        $resumeState = if ($streamingEnabled) { Get-DiscoveredAppsResumeState -Path $script:DeviceDetailResumePath } else { $null }
        $resumeStateCompatible = Test-DiscoveredAppsResumeStateCompatible -State $resumeState -Mode $DeviceDetailMode -TargetCount $script:Stat_DetailAppsTargeted -TargetAppIdsHash $detailTargetAppIdsHash -ResumeContractVersion $script:DeviceDetailResumeContractVersion
        if ($resumeStateCompatible) {
            $script:DeviceDetailPartialPath = [string]$resumeState.PartialPath
            $script:DeviceDetailTimestampedPath = [string]$resumeState.TimestampedPath
            foreach ($id in @($resumeState.ProcessedAppIds)) {
                if (-not [string]::IsNullOrWhiteSpace([string]$id)) { [void]$processedAppIds.Add([string]$id) }
            }
            $script:Stat_AppsProcessed = [int]$resumeState.ProcessedCount
            $script:Stat_AppsSkipped = [int]$resumeState.SkippedCount
            $script:Stat_DeviceDetailRows = [int]$resumeState.DeviceDetailRows
            WriteLog -Message "Resuming DeviceDetail export from $($processedAppIds.Count) processed app ids; partial CSV: $script:DeviceDetailPartialPath" "INFO"
        } else {
            if ($resumeState) {
                WriteLog -Message "Existing DeviceDetail resume state is incompatible with the current contract or target app set; starting a new DeviceDetail export." "WARNING"
                if ($resumeState.PartialPath) { Remove-Item -LiteralPath ([string]$resumeState.PartialPath) -Force -ErrorAction SilentlyContinue }
            }
            $script:DeviceDetailPartialPath = Join-Path -Path $OutputPath -ChildPath "$detailBaseFileName`_$detailTimestamp.partial.csv"
            $script:DeviceDetailTimestampedPath = Join-Path -Path $OutputPath -ChildPath "$detailBaseFileName`_$detailTimestamp.csv"
            Remove-Item -LiteralPath $script:DeviceDetailPartialPath -Force -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath $script:DeviceDetailResumePath -Force -ErrorAction SilentlyContinue

            if ($streamingEnabled -and $script:UsePreviousDeviceDetailCache -and $DeviceDetailMode -ne 'None') {
                $cacheCandidates = New-Object 'System.Collections.Generic.List[string]'
                if (-not [string]::IsNullOrWhiteSpace($globalPath)) {
                    $cacheCandidates.Add((Join-Path -Path $globalPath -ChildPath "$detailBaseFileName.csv")) | Out-Null
                }
                if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
                    $cacheCandidates.Add((Join-Path -Path $OutputPath -ChildPath "$detailBaseFileName.csv")) | Out-Null
                }

                $cachePath = @($cacheCandidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1)[0]
                if ($cachePath) {
                    WriteLog -Message "Previous DeviceDetail cache candidate found: $cachePath" "INFO"
                    $cacheResult = Use-DiscoveredAppsDeviceDetailCache `
                        -CachePath $cachePath `
                        -TargetApps @($detailApps) `
                        -PartialPath $script:DeviceDetailPartialPath `
                        -ProcessedAppIds $processedAppIds `
                        -CachedAppIds $cachedAppIds `
                        -MaxAgeDays $DeviceDetailCacheMaxAgeDays

                    if ($cacheResult.Used) {
                        $script:Stat_DetailAppsFromCache = [int]$cacheResult.Apps
                        $script:Stat_DeviceDetailRowsFromCache = [int]$cacheResult.Rows
                        $script:Stat_AppsProcessed += [int]$cacheResult.Apps
                        $script:Stat_DeviceDetailRows += [int]$cacheResult.Rows
                        WriteLog -Message ("Previous DeviceDetail cache reused: Apps={0}; Rows={1}; EnrichmentRejectedApps={2}; Cache={3}" -f $cacheResult.Apps, $cacheResult.Rows, $cacheResult.RejectedEnrichmentApps, $cacheResult.CachePath) "INFO"
                    } else {
                        WriteLog -Message ("Previous DeviceDetail cache not reused: {0}" -f $cacheResult.Reason) "INFO"
                    }
                } else {
                    WriteLog -Message "No previous DeviceDetail cache found; falling back to DeviceDetailMode=$DeviceDetailMode Graph collection." "INFO"
                }
            } elseif ($RefreshDeviceDetailCache) {
                WriteLog -Message "RefreshDeviceDetailCache specified; previous DeviceDetail cache bypassed." "INFO"
            }
        }

        WriteLog -Message 'Preloading the managed-device snapshot once for DeviceDetail enrichment...' 'INFO'
        $managedDeviceCache = @{}
        $managedDeviceSnapshotUri = 'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$top=999&$select=id,deviceName,operatingSystem,osVersion,userPrincipalName,lastSyncDateTime,enrolledDateTime,managedDeviceOwnerType,complianceState,azureADDeviceId'
        foreach ($managedDevice in @(Invoke-GraphPagedRequest -InitialUri $managedDeviceSnapshotUri)) {
            if ($managedDevice.id) { $managedDeviceCache[[string]$managedDevice.id] = $managedDevice }
        }
        WriteLog -Message ("Managed-device snapshot cached: {0} unique device(s). App relation calls now retrieve IDs only." -f $managedDeviceCache.Count) 'INFO'
        $seenAppDevicePairs = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $appsForRelationBatch = @($detailApps | Where-Object { -not $processedAppIds.Contains([string]$_.id) })
        $appDeviceRelationMap = Get-DiscoveredAppDeviceRelationBatchMap -Apps $appsForRelationBatch -DelayMs $DelayMs
        WriteLog -Message ("Discovered Apps relation batches completed: {0}/{1} app(s) prefetched." -f $appDeviceRelationMap.Count, $appsForRelationBatch.Count) 'INFO'

        $appIndex = 0
        foreach ($app in $detailApps) {
            $appIndex++
            if ($processedAppIds.Contains([string]$app.id)) {
                if (-not $cachedAppIds.Contains([string]$app.id)) {
                    $script:Stat_DetailAppsSkippedByResume++
                }
                continue
            }

            $pctComplete = [math]::Round(($appIndex / $detailApps.Count) * 100, 1)
            Write-Progress `
                -Activity "Collecting Intune discovered app device details" `
                -Status "App $appIndex / $($detailApps.Count): $($app.displayName) v$($app.version) ($pctComplete%)" `
                -PercentComplete $pctComplete

            if ($ProgressEveryApps -gt 0 -and (($appIndex % $ProgressEveryApps -eq 0) -or $appIndex -eq $detailApps.Count)) {
                WriteLog -Message ("Device detail progress: {0}/{1} apps ({2}%). DetailRows={3}; Processed={4}; Skipped={5}; CacheApps={6}; ResumedSkipped={7}; GraphCalls={8}; ThrottleRetries={9}" -f $appIndex, $detailApps.Count, $pctComplete, $script:Stat_DeviceDetailRows, $script:Stat_AppsProcessed, $script:Stat_AppsSkipped, $script:Stat_DetailAppsFromCache, $script:Stat_DetailAppsSkippedByResume, $script:Stat_GraphCalls, $script:Stat_ThrottleRetries) "INFO"
            }

            try {
                $devicesUri = 'https://graph.microsoft.com/v1.0/deviceManagement/detectedApps/' + $app.id + '/managedDevices' +
                    '?$top=999&$select=id'
                $devices = if ($appDeviceRelationMap.ContainsKey([string]$app.id)) {
                    @($appDeviceRelationMap[[string]$app.id])
                }
                else {
                    if ($DelayMs -gt 0) { Start-Sleep -Milliseconds $DelayMs }
                    Invoke-GraphPagedRequest -InitialUri $devicesUri
                }

                $appRows = [System.Collections.Generic.List[psobject]]::new()
                if (-not $devices -or $devices.Count -eq 0) {
                    $appRows.Add((New-DiscoveredAppsDeviceDetailRecord -App $app -Device $null))
                    $script:Stat_DeviceDetailRows++
                } else {
                    foreach ($device in $devices) {
                        $deviceId = [string]$device.id
                        $pairKey = '{0}|{1}' -f $app.id, $deviceId
                        if (-not $seenAppDevicePairs.Add($pairKey)) { continue }
                        $resolvedDevice = if ($managedDeviceCache.ContainsKey($deviceId)) { $managedDeviceCache[$deviceId] } else { $device }
                        $appRows.Add((New-DiscoveredAppsDeviceDetailRecord -App $app -Device $resolvedDevice))
                        $script:Stat_DeviceDetailRows++
                    }
                }

                if ($streamingEnabled) {
                    Write-DiscoveredAppsCsvRows -Path $script:DeviceDetailPartialPath -Rows @($appRows)
                } else {
                    foreach ($row in $appRows) { $detailRecords.Add($row) }
                }

                $script:Stat_AppsProcessed++
                [void]$processedAppIds.Add([string]$app.id)

                if ($streamingEnabled -and (($script:Stat_AppsProcessed + $script:Stat_AppsSkipped) % 25 -eq 0 -or $appIndex -eq $detailApps.Count)) {
                    Save-DiscoveredAppsResumeState `
                        -Path $script:DeviceDetailResumePath `
                        -PartialPath $script:DeviceDetailPartialPath `
                        -TimestampedPath $script:DeviceDetailTimestampedPath `
                        -Mode $DeviceDetailMode `
                        -TargetCount $script:Stat_DetailAppsTargeted `
                        -ResumeContractVersion $script:DeviceDetailResumeContractVersion `
                        -TargetAppIdsHash $detailTargetAppIdsHash `
                        -ProcessedAppIds @($processedAppIds) `
                        -ProcessedCount $script:Stat_AppsProcessed `
                        -SkippedCount $script:Stat_AppsSkipped `
                        -DetailRows $script:Stat_DeviceDetailRows
                }
            } catch {
                WriteLog -Message "Failed to retrieve devices for app '$($app.displayName)' (Id=$($app.id)): $_" "WARNING"
                $script:Stat_AppsSkipped++
                [void]$processedAppIds.Add([string]$app.id)
                if ($streamingEnabled) {
                    Save-DiscoveredAppsResumeState `
                        -Path $script:DeviceDetailResumePath `
                        -PartialPath $script:DeviceDetailPartialPath `
                        -TimestampedPath $script:DeviceDetailTimestampedPath `
                        -Mode $DeviceDetailMode `
                        -TargetCount $script:Stat_DetailAppsTargeted `
                        -ResumeContractVersion $script:DeviceDetailResumeContractVersion `
                        -TargetAppIdsHash $detailTargetAppIdsHash `
                        -ProcessedAppIds @($processedAppIds) `
                        -ProcessedCount $script:Stat_AppsProcessed `
                        -SkippedCount $script:Stat_AppsSkipped `
                        -DetailRows $script:Stat_DeviceDetailRows
                }
            }
        }

        Write-Progress -Activity "Collecting Intune discovered app device details" -Completed

        if ($streamingEnabled) {
            Save-DiscoveredAppsResumeState `
                -Path $script:DeviceDetailResumePath `
                -PartialPath $script:DeviceDetailPartialPath `
                -TimestampedPath $script:DeviceDetailTimestampedPath `
                -Mode $DeviceDetailMode `
                -TargetCount $script:Stat_DetailAppsTargeted `
                -ResumeContractVersion $script:DeviceDetailResumeContractVersion `
                -TargetAppIdsHash $detailTargetAppIdsHash `
                -ProcessedAppIds @($processedAppIds) `
                -ProcessedCount $script:Stat_AppsProcessed `
                -SkippedCount $script:Stat_AppsSkipped `
                -DetailRows $script:Stat_DeviceDetailRows

            if (($script:Stat_AppsProcessed + $script:Stat_AppsSkipped) -eq $script:Stat_DetailAppsTargeted) {
                $completedPath = Complete-DiscoveredAppsStreamExport `
                    -PartialPath $script:DeviceDetailPartialPath `
                    -TimestampedPath $script:DeviceDetailTimestampedPath `
                    -OutputPath $OutputPath `
                    -GlobalPath $globalPath `
                    -BaseFileName $detailBaseFileName
                if ($completedPath) {
                    $script:DeviceDetailCompleted = $true
                    Remove-Item -LiteralPath $script:DeviceDetailResumePath -Force -ErrorAction SilentlyContinue
                }
            } else {
                WriteLog -Message "DeviceDetail export incomplete; partial CSV and resume state kept: $script:DeviceDetailPartialPath" "WARNING"
            }
        }
    }

    WriteLog -Message ("DeviceDetail completed. TargetApps={0}; Processed={1}; Skipped={2}; CacheApps={3}; ResumedSkipped={4}; Rows={5}; CacheRows={6}" -f $script:Stat_DetailAppsTargeted, $script:Stat_AppsProcessed, $script:Stat_AppsSkipped, $script:Stat_DetailAppsFromCache, $script:Stat_DetailAppsSkippedByResume, $script:Stat_DeviceDetailRows, $script:Stat_DeviceDetailRowsFromCache) "INFO"

} catch {
    $globalError = $_
    $script:RunError = $globalError
    WriteLog -Message ("Global error in $TaskName : {0}" -f $globalError) "ERROR"
    Write-Host "A global error occurred. Check the log file for details." -ForegroundColor Red

    try {
        $title = "Intune Windows Discovered Apps - ERROR"
        $msg   = @"
An error occurred in script $($MyInvocation.MyCommand.Name) on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss").

Error message:
$($globalError.Exception.Message)

See attached log file for details:
$($global:LogTextFile)
"@
        $bodyHtml    = NewSimpleEmailBody -Title $title -Message $msg
        $attachments = @()
        if ($global:LogTextFile -and (Test-Path $global:LogTextFile)) {
            $attachments = @($global:LogTextFile)
        }
        SendEmailHtmlReport -BodyHtml $bodyHtml -Subject $title -Attachments $attachments -VerboseLog
    } catch {
        WriteLog -Message ("Failed to send error notification email: {0}" -f $_) "ERROR"
    }

} finally {
    Write-Progress -Activity 'Retrieving device details per app' -Completed -ErrorAction SilentlyContinue

    if ($connectedGraphInThisRun) {
        Write-Host "`n--- Disconnect Cloud Services ---"
        try { Disconnect-MgGraph | Out-Null } catch {
            WriteLog -Message ("Error during Graph disconnect: {0}" -f $_) "WARNING"
        }
    }

    try {
        RemoveOldFiles -Path $OutputPath     -Filter "*.csv" -KeepCount 150 -LogFile $global:LogTextFile
        RemoveOldFiles -Path $global:LogPath -Filter "*.log" -KeepCount $global:RetentionMaxLogs  -LogFile $global:LogTextFile
    } catch {
        WriteLog -Message ("Error during cleanup: {0}" -f $_) "WARNING"
    }

    $stopwatch.Stop()
    $elapsed    = $stopwatch.Elapsed
    $elapsedStr = '{0:D2}h {1:D2}m {2:D2}s' -f $elapsed.Hours, $elapsed.Minutes, $elapsed.Seconds

    $detailCompletionCount = $script:Stat_AppsProcessed + $script:Stat_AppsSkipped
    if ($script:Stat_DetailAppsTargeted -gt 0 -and $detailCompletionCount -lt $script:Stat_DetailAppsTargeted -and -not $script:RunError) {
        $incompleteMessage = "DeviceDetail run incomplete: processed/skipped $detailCompletionCount of $($script:Stat_DetailAppsTargeted) targeted apps. Resume state: $($script:DeviceDetailResumePath)"
        WriteLog -Message $incompleteMessage "ERROR"
        $exception = [System.Exception]::new($incompleteMessage)
        $script:RunError = [System.Management.Automation.ErrorRecord]::new($exception, 'DiscoveredAppsIncomplete', [System.Management.Automation.ErrorCategory]::OperationStopped, $null)
    } elseif ($script:Stat_DetailAppsTargeted -gt 0 -and -not $DryRun -and -not $script:DeviceDetailCompleted -and -not $script:RunError) {
        $incompleteMessage = "DeviceDetail CSV was not finalized although all targeted apps were processed. Resume state: $($script:DeviceDetailResumePath)"
        WriteLog -Message $incompleteMessage "ERROR"
        $exception = [System.Exception]::new($incompleteMessage)
        $script:RunError = [System.Management.Automation.ErrorRecord]::new($exception, 'DiscoveredAppsExportIncomplete', [System.Management.Automation.ErrorCategory]::OperationStopped, $null)
    }
    WriteLog -Message "=== Run summary ==="
    WriteLog -Message "  Total apps retrieved (all platforms) : $($script:Stat_AppsTotal)"
    WriteLog -Message "  Windows apps found                   : $($script:Stat_AppsWindows)"
    WriteLog -Message "  Apps processed successfully          : $($script:Stat_AppsProcessed)"
    WriteLog -Message "  Apps skipped (errors)                : $($script:Stat_AppsSkipped)"
    WriteLog -Message "  Apps skipped by resume               : $($script:Stat_DetailAppsSkippedByResume)"
    WriteLog -Message "  Apps served from cache               : $($script:Stat_DetailAppsFromCache)"
    WriteLog -Message "  Device detail rows exported          : $($script:Stat_DeviceDetailRows)"
    WriteLog -Message "  Device detail rows from cache        : $($script:Stat_DeviceDetailRowsFromCache)"
    WriteLog -Message "  Graph API calls                      : $($script:Stat_GraphCalls)"
    WriteLog -Message "  Throttle retries (429)               : $($script:Stat_ThrottleRetries)"
    WriteLog -Message "  Elapsed time                         : $elapsedStr"
    WriteLog -Message "$TaskName completed."

    try {
        Stop-DiscoveredAppsTranscript
        $summaryStatus = if ($script:RunError) { 'Failed' } else { 'Auto' }
        Complete-SmartM365ExecutionContext -Status $summaryStatus -ErrorRecord $script:RunError
    } catch {
        WriteLog -Message ("Failed to write execution summary: {0}" -f $_) "WARNING"
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB1AjyDslDgSIJv
# eTSSgJHtSr8rrC6hYd2ly4iyx9IUDqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINDwmACVqUueWlxHnR5LGzgoe2tft50NWjL6T2ReeKPCMA0GCSqG
# SIb3DQEBAQUABIIBgD8829WXYWeww/y7+lH8JJjWvQTolqhQaoX0EyT4h4m+jMwM
# m6CcEXKa4WfknlphioEvCfKtrAvfcVPD69dthCHCtsNJkUvslSJ4bxCKBuU+woZx
# m2j9h+oNlaKG82JceXQifNsoEBtz//6d8WQqc3UO5wmNvcuY4clrFiKSqKT/QGBX
# ALSvV5XZcAd8C7vbdWHZyOpvB3GO04nvKJ7rA+w6FlIKfaas+U0eZ+VBX6bowNxa
# p8xRNk6E7JvdxIvUv9s1/09BPgR7KR3mnVFpdVW/+Dc4+flWUKxRYWPElrLfRNR1
# WukxqzMO1UZOZNOJmCnEZmiLNAgmWjMKzL2sr1d9wOiVCckz0EDMQo0XTyBSDSOz
# 5qSBZpiXfwa+ImKE000l74APG8nV0qqBNneUYrT57aXUucQmpPRc0LXnxMhrAACS
# n9Uu9GtvHzOxlfNc6G1v6fJAnLw5JHsiYygx1wwueTz3ICtrglq7FSKWCZ4Nsx70
# r211YwsjjoLSR5f5X6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTUxNjA1
# MTVaMC8GCSqGSIb3DQEJBDEiBCARditA2DYn6I3mATOuu7dv2gPExhPAFikNspGZ
# UqkYDDANBgkqhkiG9w0BAQEFAASCAgCDFRlmUIujnlTL4qJ/RItTugKm3H+iODt2
# 6DXO/yGMEivG/SuQQ4l8tF//ID5rRbqOPvpsWZXv0J5vZp9cGXYqSMAZ5jyLi5E9
# 6+1KabYKlT6iX5zL84bv3S3r2sJ/p8ptYlI1RDQxPl4VNj5/glg+fO7x9QqbdCZB
# dPwJNTcJJPDokR4ztdpC49Ob0PpesRmr+6Xpy7qtJ40/f/G0gzNcm58DrpJriDXz
# rjbb5IGd/sOiNWasz+bblYbOWo1mIWw9WBkd5O2oL43H4pOtoj41ZtwTv9zLE++Z
# 3ho0NQq3I2PnLvQfhcngvTqtejrhr5HgXg3L744iRL0J9KJiiWyBLXGp19+YfBiP
# Uvt5uWVnP/mVL/OJn877FWaXsFH5sNb2iPiNtGJHvFGzOtw58mMqZu+ndAmf/+O+
# mZO2AMPI3WOZ03Hd+1RPu5SGkfRIR2Yrr25MIz+WueeqI9d8iPNiLuG8Qy4B7aw/
# wDrFQ5dNKHTgVPYTD5NMbjQ/L/cSXZfTSESOi4kmaFd9npClR6x/Xoy4ntl1gd/v
# VjjnMOUsyFJ9ehR1Pchd3L8faUVjbGWo0Yg/0qOXsXjlDFspPfpzSIaOo6rq1rnc
# dd8bsMv8AQj3bneNyuQZfx9MhtU4oUmnp0dofACZCruiIivB/ViKrfEqnrWHZSCt
# ClrJTl6IfA==
# SIG # End signature block
