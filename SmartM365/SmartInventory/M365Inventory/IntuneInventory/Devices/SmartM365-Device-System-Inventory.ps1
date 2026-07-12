<#
.SYNOPSIS
    Generates a device system inventory (SecureBoot, BIOS, FirmwareType, TPM, Encryption, DeviceGuard) for all Intune managed devices.

.DESCRIPTION
    Connects to Microsoft Graph Beta and retrieves:
    - Hardware security fields from hardwareInformation (TPM, encryption, DeviceGuard)
    - SecureBoot status, BIOS version and firmware type from a deployed Platform Script
      via deviceManagementScripts/deviceRunStates

    It:
    - Retrieves all managed devices with hardwareInformation in bulk
    - Correlates SecureBoot stdout from the Platform Script deviceRunStates
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
    The GUID of the Intune Platform Script (Detect-DeviceSystemInfo) deployed to devices.
    Required to retrieve SecureBoot/BIOS/FirmwareType results from deviceRunStates.
.VERSION
1.6



.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication.
    Minimum Graph application permissions: DeviceManagementManagedDevices.Read.All; DeviceManagementConfiguration.Read.All; DeviceManagementScripts.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Version : 1.2
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
# App-only authentication parameters (same app as other scripts)
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

# ==========================================================
# Import SmartM365.Core module (psd1)
# ==========================================================
$modulePath = & { $d = $PSScriptRoot; while ($d) { $p = Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'; if (Test-Path -LiteralPath $p) { return $p }; $parent = Split-Path -Path $d -Parent; if ($parent -eq $d) { break }; $d = $parent }; throw 'SmartM365.Core module not found.' }
try {
    Import-Module $modulePath -ErrorAction Stop
} catch {
    Write-Host "Failed to import SmartM365.Core module from '$modulePath' : $_" -ForegroundColor Red
    exit 1
}

# ==========================================================
# Configuration
# ==========================================================

# Throttle-aware retry settings
$MaxRetries        = 5
$BaseDelaySeconds  = 2
$BatchSize         = 50
$BatchPauseSeconds = 2

# Graph API settings - hardwareInformation fields that are actually populated
$BulkEndpoint = '/beta/deviceManagement/managedDevices?$select=id,deviceName,azureADDeviceId,userPrincipalName,managementState,complianceState,lastSyncDateTime,hardwareInformation&$top=999'

# ==========================================================
# Helpers
# ==========================================================

function Invoke-GraphSafe {
    <#
    .SYNOPSIS
        Throttle-aware wrapper around Invoke-MgGraphRequest with exponential back-off.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Method = 'GET',
        [int]$MaxRetries = 5,
        [int]$BaseDelaySeconds = 2
    )

    $attempt = 0
    while ($true) {
        $attempt++
        try {
            $response = Invoke-MgGraphRequest -Uri $Uri -Method $Method -OutputType PSObject
            return $response
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            # Fallback: detect SDK-wrapped 429 when Response is null (SDK exhausted its own retries)
            $isSdkWrapped429 = ($null -eq $statusCode) -and ($_.Exception.Message -match 'TooManyRequests|too many retries|429')

            # Throttled (429) or transient server error (502, 503, 504)
            if (($statusCode -in @(429, 502, 503, 504) -or $isSdkWrapped429) -and $attempt -le $MaxRetries) {
                $retryAfter = $BaseDelaySeconds * [math]::Pow(2, $attempt - 1)

                # SDK-wrapped throttle: use a fixed 60s floor - the SDK already burned its own retries
                if ($isSdkWrapped429) {
                    $retryAfter = [math]::Max($retryAfter, 60)
                }
                # Check Retry-After header (only available on raw responses)
                elseif ($_.Exception.Response.Headers -and $_.Exception.Response.Headers['Retry-After']) {
                    $retryAfterHeader = $_.Exception.Response.Headers['Retry-After'] | Select-Object -First 1
                    if ([int]::TryParse($retryAfterHeader, [ref]$null)) {
                        $retryAfter = [int]$retryAfterHeader
                    }
                }

                WriteLog -Message ("HTTP {0} on attempt {1}/{2}. Retrying in {3}s... URI: {4}" -f $(if ($isSdkWrapped429) { '429(sdk)' } else { $statusCode }), $attempt, $MaxRetries, $retryAfter, $Uri) "WARNING"
                Start-Sleep -Seconds $retryAfter
            }
            else {
                throw
            }
        }
    }
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

    WriteLog -Message ("Retrieving Platform Script deviceRunStates for script {0}..." -f $ScriptId) "INFO"

    $map = @{}
    $uri = "/beta/deviceManagement/deviceManagementScripts/$ScriptId/deviceRunStates?`$expand=managedDevice(`$select=id)&`$select=resultMessage,lastStateUpdateDateTime&`$top=40"

    while ($uri) {
        $response = Invoke-GraphSafe -Uri $uri -MaxRetries $MaxRetries -BaseDelaySeconds $BaseDelaySeconds

        if ($response.value) {
            foreach ($runState in $response.value) {
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

        $uri = $response.'@odata.nextLink'
        if ($uri) {
            WriteLog -Message ("Paging deviceRunStates... Total so far: {0}" -f $map.Count) "INFO"
            Start-Sleep -Seconds $BatchPauseSeconds
        }
    }

    WriteLog -Message ("Platform Script deviceRunStates retrieved: {0} devices" -f $map.Count) "INFO"
    return $map
}

function Parse-PlatformScriptStdout {
    <#
    .SYNOPSIS
        Parses the pipe-delimited stdout from the Detect-DeviceSystemInfo Platform Script.
        Format: SecureBoot:<value>|BIOSVersion:<value>|BIOSDate:<value>|FirmwareType:<value>
        Returns a hashtable with each key/value pair.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [string]$ResultMessage
    )

    $result = @{
        SecureBoot   = $null
        BIOSVersion  = $null
        BIOSDate     = $null
        FirmwareType = $null
    }

    if ([string]::IsNullOrWhiteSpace($ResultMessage)) {
        return $result
    }

    $parts = $ResultMessage.Trim().Split('|')
    foreach ($part in $parts) {
        if ($part -match '^([^:]+):(.+)$') {
            $key   = $Matches[1].Trim()
            $value = $Matches[2].Trim()
            if ($result.ContainsKey($key)) {
                $result[$key] = $value
            }
        }
    }

    return $result
}

# ==========================================================
# Initialization via SmartM365.Core
# ==========================================================
$ScriptVersion = "1.6"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DeviceSystemCsvLogFolderPath' -DefaultValue $OutputPath
try {
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPathInit $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
    Start-Transcript -Path $global:logTranscriptFile -Append

    WriteLog -Message "Script Environment initialized at $InitializeOutputPath"
    $OutputPath = $InitializeOutputPath
    WriteLog -Message "Starting $TaskName..."
} catch {
    Write-Host "Initialization failed: $_" -ForegroundColor Red
    exit 1
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
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
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
            Connect-MgGraph -TenantId $TenantId -ClientId $AppId -CertificateThumbprint $Thumb -NoWelcome | Out-Null
        } else {
            WriteLog -Message "Connecting to Microsoft Graph with interactive authentication (direct Connect-MgGraph)." "INFO"
            Connect-MgGraph -Scopes @("DeviceManagementManagedDevices.Read.All","DeviceManagementConfiguration.Read.All") -NoWelcome | Out-Null
        }

        if (-not (Test-GraphConnection)) {
            throw "Failed to connect to Microsoft Graph."
        }

        $connectedGraphInThisRun = $true
    }

    # ======================================================
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -RequiredGraphApplicationPermissions @('DeviceManagementManagedDevices.Read.All','DeviceManagementConfiguration.Read.All','DeviceManagementScripts.Read.All') -GraphProbeUris @(
        'https://graph.microsoft.com/beta/deviceManagement/managedDevices?$top=1',
        'https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts?$top=1'
    ) | Out-Null

    # PHASE 1: BULK RETRIEVAL OF MANAGED DEVICES
    # ======================================================
    WriteLog -Message "Phase 1: Bulk retrieval of managed devices with hardwareInformation..." "INFO"

    $allDevices = [System.Collections.Generic.List[PSObject]]::new()
    $uri = $BulkEndpoint

    while ($uri) {
        $response = Invoke-GraphSafe -Uri $uri -MaxRetries $MaxRetries -BaseDelaySeconds $BaseDelaySeconds
        if ($response.value) {
            $allDevices.AddRange([PSObject[]]$response.value)
            if ($MaxItems -gt 0 -and $allDevices.Count -ge $MaxItems) {
                while ($allDevices.Count -gt $MaxItems) { $allDevices.RemoveAt($allDevices.Count - 1) }
                WriteLog -Message ("MaxItems enabled: restricted managed devices to {0}." -f $allDevices.Count) "WARNING"
                break
            }
        }
        $uri = $response.'@odata.nextLink'
        if ($uri) {
            WriteLog -Message ("Paging... Total devices so far: {0}" -f $allDevices.Count) "INFO"
        }
    }

    WriteLog -Message ("Bulk retrieval complete. Total devices: {0}" -f $allDevices.Count) "INFO"

    if ($allDevices.Count -eq 0) {
        WriteLog -Message "No managed devices found. Exiting." "WARNING"
        return
    }

    # ======================================================
    # PHASE 2: RETRIEVE SECUREBOOT FROM PLATFORM SCRIPT
    # ======================================================
    $secureBootMap = @{}

    if (-not [string]::IsNullOrWhiteSpace($PlatformScriptId)) {
        WriteLog -Message "Phase 2: Retrieving SecureBoot status from Platform Script deviceRunStates..." "INFO"
        $secureBootMap = Get-PlatformScriptRunStates -ScriptId $PlatformScriptId
    }
    else {
        WriteLog -Message "Phase 2: Skipped (no -PlatformScriptId provided). SecureBoot column will be empty." "WARNING"
        WriteLog -Message "To populate SecureBoot/BIOS/FirmwareType, deploy the Detect-DeviceSystemInfo Platform Script and pass its GUID via -PlatformScriptId." "WARNING"
    }

    # ======================================================
    # BUILD RESULTS
    # ======================================================
    WriteLog -Message "Building results..." "INFO"

    $results = [System.Collections.Generic.List[PSObject]]::new()

    foreach ($device in $allDevices) {
        $hw = $device.hardwareInformation

        # Platform Script data (SecureBoot, BIOSVersion, FirmwareType)
        $psData = @{ SecureBoot = $null; BIOSVersion = $null; BIOSDate = $null; FirmwareType = $null }
        $psLastUpdate = $null
        if ($secureBootMap.ContainsKey($device.id)) {
            $runState = $secureBootMap[$device.id]
            $psData = Parse-PlatformScriptStdout -ResultMessage $runState.ResultMessage
            $psLastUpdate = $runState.LastStateUpdateDateTime
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
            PlatformScriptLastUpdate          = $psLastUpdate
            IsEncrypted                       = if ($hw) { $hw.isEncrypted } else { $null }
            DeviceGuardVBSState               = if ($hw) { $hw.deviceGuardVirtualizationBasedSecurityState } else { $null }
            DeviceGuardVBSHardwareRequirement = if ($hw) { $hw.deviceGuardVirtualizationBasedSecurityHardwareRequirementState } else { $null }
            DeviceGuardCredentialGuardState   = if ($hw) { $hw.deviceGuardLocalSystemAuthorityCredentialGuardState } else { $null }
            ManagementState                   = $device.managementState
            ComplianceState                   = $device.complianceState
            LastSyncDateTime                  = $device.lastSyncDateTime
        }

        $results.Add($record)
    }

    # ======================================================
    # If no devices, exit gracefully
    # ======================================================
    if (-not $results -or $results.Count -eq 0) {
        WriteLog -Message "No devices found. Exiting." "WARNING"
        return
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
    # FILTER: Keep only devices with Platform Script data
    # ======================================================
    $filteredResults = @($results | Where-Object { -not [string]::IsNullOrWhiteSpace($_.SecureBootStatus) })
    WriteLog -Message ("Filtered results: {0} devices with Platform Script data (excluded {1} without)" -f $filteredResults.Count, ($results.Count - $filteredResults.Count)) "INFO"

    if ($filteredResults.Count -eq 0) {
        WriteLog -Message "No devices with Platform Script data to export. CSV will not be generated." "WARNING"
        WriteLog -Message "Ensure the Platform Script has been executed on target devices (allow 24-48h after deployment)." "WARNING"
        return
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

}
catch {
    $globalError = $_
    WriteLog -Message ("Global error in Device System inventory: {0}" -f $globalError) "ERROR"
    Write-Host "A global error occurred. Check the log file for details." -ForegroundColor Red
}
finally {
    # Disconnect Graph only if we connected it in this run
    if ($connectedGraphInThisRun) {
        Write-Host "`n--- Disconnect Cloud Services ---"
        try {
            if (Get-MgContext -ErrorAction SilentlyContinue) {
                try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
            }
        } catch {
            WriteLog -Message ("Error during Graph disconnect in finally: {0}" -f $_) "WARNING"
        }
    }

    # Cleanup
    try {
        RemoveOldFiles -Path $OutputPath     -Filter "*.csv" -KeepCount 150 -LogFile $global:LogTextFile
        RemoveOldFiles -Path $global:LogPath -Filter "*.log" -KeepCount $global:RetentionMaxLogs -LogFile $global:LogTextFile
    } catch {
        WriteLog -Message ("Error during cleanup in finally: {0}" -f $_) "WARNING"
    }

    WriteLog -Message "$TaskName completed (finally block)."

    try {
        Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
        Complete-SmartM365ExecutionContext -Status Auto
    } catch {
        # ignore
    }
}
# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB6m2N8VmSB6y66
# 2CZX+cyWTumXlTX3wxA1bftFXZJVS6CCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCmUkashzJOLcPyD3ge
# r9Qba56ZyuPdFxo9f3BLnpuCUzANBgkqhkiG9w0BAQEFAASCAYAwNTVT1EU/NJUu
# +eGV+PFPJl4uV8jqw19uQgw3DIdCW8EzE+BI5SJl8KS+3QHFA+rWbldGDXkzX/il
# BLAHc0hoyvUkUuhFOBkD3mzjWAZ/uoNsnXtIhoht7IvjTNxUzmv0fk4lBFQOi06S
# G8b7ESNGzbdwTrf7VmRO7xb1EAhV3TuS7soujaTJyALybIiL5h+JlLRPGiDGD2IL
# eHTrSqeuOrQhONAUa2wbNGmuUH+SBB1RMJqrSHSfYbZEIQvko2erwdQXeXfYEUmh
# 3Hf6InHJamYhcPRoXytKBuqcUJeliHcQwJo6v4LNVVrH4W/YiFBMh0VRyf8/ePtF
# zYYuKP1vvQ5o1yazUfQuILlHkc6YMi52ptJa03eKAGeTt0vDtn02uzjO0wEayt2F
# 4GPZoAkpHMDhniiE7R/YbMk4/CRKeh0tE/nGsIoadNrneyDfoWabcrZ1XvgIElJj
# K3DLayBceiu905q1FJ0r7SUOoHaOOnftVDFkPR/Pm3CIGSXBB0I=
# SIG # End signature block
