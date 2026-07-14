<#
.SYNOPSIS
Exports BIOS inventory (systemManagementBIOSVersion) for Intune-managed Windows devices using Microsoft Graph.

.DESCRIPTION
- shared pattern (PS7 check, SmartM365.Core init, logging, transcript, Graph connect/disconnect, CSV export, cleanup).
- Lists Intune managed Windows devices (Windows OS).
- Per-device BIOS retrieval using multi-step fallback strategy:
  1) /beta $select=hardwareInformation
  2) /beta $expand=hardwareInformation
  3) /v1.0 $select=hardwareInformation
  4) /v1.0 $expand=hardwareInformation
  5) Minimal GET (v1.0) to qualify object readability (FIXED URI: includes '?$select=')
- Captures HTTP status and request-id (best effort) on both success and failure.
- Retry strategy honors Retry-After when present (e.g., 429) and falls back to exponential backoff.
- No email notification.
- Optional -DebugHardwareInfo switch: dumps all non-null hardwareInformation properties to the log
  for the first N devices (controlled by -DebugDeviceCount). Useful to diagnose missing fields
  such as biosVersion or biosReleaseDateTime.

.PERMISSIONS
Microsoft Graph:
- DeviceManagementManagedDevices.Read.All

.PARAMETER OutputPath
Output folder (optional). If empty, default output path is used.

.PARAMETER Connect
Forces a Graph disconnect/reconnect.

.PARAMETER InteractiveAuth
Uses interactive authentication instead of app-only certificate authentication.

.PARAMETER DeviceNameContains
Optional filter on deviceName (case-insensitive contains).

.PARAMETER DebugHardwareInfo
If specified, dumps all non-null hardwareInformation properties (name + value) to the log
for the first N devices where hardwareInformation is successfully retrieved.
Stops automatically after DebugDeviceCount devices have been dumped.

.PARAMETER DebugDeviceCount
Number of devices for which to dump hardwareInformation properties when -DebugHardwareInfo is active.
Default: 3.
.VERSION
1.8


.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication.
    Minimum Graph application permissions: DeviceManagementManagedDevices.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Version : 1.6
    Minimum application permissions: DeviceManagementManagedDevices.Read.All
#>

param(
    [string]$Tenant = 'test',
[string]$OutputPath,
    [switch]$Connect,
    [switch]$InteractiveAuth,
    [string]$DeviceNameContains,
    [switch]$DebugHardwareInfo,
    [int]$DebugDeviceCount = 3,
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

$MaximumFunctionCount = 32768

# ==========================================================
# App-only authentication parameters (align with your inventory scripts)
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

# ==========================================================
# Import SmartM365.Core module (psd1)
# ==========================================================
$ModulePath = & { $d = $PSScriptRoot; while ($d) { $p = Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'; if (Test-Path -LiteralPath $p) { return $p }; $parent = Split-Path -Path $d -Parent; if ($parent -eq $d) { break }; $d = $parent }; throw 'SmartM365.Core module not found.' }
try {
    Import-Module -Name $ModulePath -MinimumVersion '1.0.24' -ErrorAction Stop
} catch {
    Write-Host "Failed to import SmartM365.Core module from '$ModulePath' : $_" -ForegroundColor Red
    exit 1
}

# ==========================================================
# Helpers
# ==========================================================
function Get-GraphErrorDetails {
    param([Parameter(Mandatory)]$ErrorRecord)

    $httpStatus = $null
    $requestId = $null
    $raw = $null
    $retryAfterSeconds = $null

    try {
        if ($ErrorRecord.Exception.PSObject.Properties.Name -contains "ResponseStatusCode") {
            $httpStatus = [string]$ErrorRecord.Exception.ResponseStatusCode
        } else {
            # Best effort: parse the standard MgGraph message "(BadRequest (Bad Request))"
            if ($ErrorRecord.Exception.Message -match ":\s*.*\((?<code>[^)]+)\s*\(") {
                $httpStatus = $matches["code"]
            }
        }
    } catch { }

    try {
        if ($ErrorRecord.Exception.PSObject.Properties.Name -contains "ResponseHeaders" -and $ErrorRecord.Exception.ResponseHeaders) {
            $headers = $ErrorRecord.Exception.ResponseHeaders

            if ($headers.ContainsKey("request-id")) { $requestId = ($headers["request-id"] | Select-Object -First 1) }
            if (-not $requestId -and $headers.ContainsKey("client-request-id")) { $requestId = ($headers["client-request-id"] | Select-Object -First 1) }

            if ($headers.ContainsKey("Retry-After")) {
                $ra = ($headers["Retry-After"] | Select-Object -First 1)
                if ($ra -match "^\d+$") { $retryAfterSeconds = [int]$ra }
            }
        }
    } catch { }

    try {
        if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
            $raw = $ErrorRecord.ErrorDetails.Message
        } else {
            $raw = $ErrorRecord.Exception.Message
        }
    } catch {
        $raw = $ErrorRecord.Exception.Message
    }

    [pscustomobject]@{
        HttpStatus         = $httpStatus
        RequestId          = $requestId
        RetryAfterSeconds  = $retryAfterSeconds
        RawError           = $raw
    }
}

function Invoke-GraphRequestWithRetry {
    param(
        [Parameter(Mandatory)][ValidateSet("GET","POST","PATCH","PUT","DELETE")][string]$Method,
        [Parameter(Mandatory)][string]$Uri,
        [int]$MaxRetries = 6
    )

    $attempt = 0
    while ($true) {
        try {
            $rh = $null
            $sc = $null

            $body = Invoke-MgGraphRequest -Method $Method -Uri $Uri -ErrorAction Stop `
                -ResponseHeadersVariable rh -StatusCodeVariable sc

            $requestId = $null
            try {
                if ($rh) {
                    if ($rh.ContainsKey("request-id")) { $requestId = ($rh["request-id"] | Select-Object -First 1) }
                    if (-not $requestId -and $rh.ContainsKey("client-request-id")) { $requestId = ($rh["client-request-id"] | Select-Object -First 1) }
                }
            } catch { }

            return [pscustomobject]@{
                Body       = $body
                StatusCode = [string]$sc
                Headers    = $rh
                RequestId  = $requestId
            }
        } catch {
            $attempt++
            $msg = $_.Exception.Message

            $d = Get-GraphErrorDetails -ErrorRecord $_

            $isRetryable = $false
            if ($msg -match "Too Many Requests" -or $msg -match "\b429\b") { $isRetryable = $true }
            if ($msg -match "temporarily unavailable" -or $msg -match "timeout" -or $msg -match "The remote server returned an error") { $isRetryable = $true }

            if (-not $isRetryable -or $attempt -gt $MaxRetries) {
                throw
            }

            $sleepSeconds = $null

            # Prefer server guidance when present (Retry-After)
            if ($d.RetryAfterSeconds -and $d.RetryAfterSeconds -gt 0) {
                $sleepSeconds = [int]$d.RetryAfterSeconds
            } else {
                $sleepSeconds = [Math]::Min(60, [Math]::Pow(2, $attempt))
            }

            WriteLog -Message "Graph request throttled/failed (attempt $attempt/$MaxRetries). Sleeping $sleepSeconds sec. Uri=$Uri. HttpStatus=$($d.HttpStatus). RequestId=$($d.RequestId). Error=$($d.RawError)" "WARNING"
            Start-Sleep -Seconds $sleepSeconds
        }
    }
}

function Get-InventoryColumns {
    [string[]]@(
        "ManagedDeviceId",
        "DeviceName",
        "AzureADDeviceId",
        "UserPrincipalName",
        "OperatingSystem",
        "OSVersion",
        "Manufacturer",
        "Model",
        "SerialNumber",
        "SystemManagementBIOSVersion",
        "BIOSVersionRaw",
        "BIOSReleaseDateTime",
        "HardwareInfoStatus",
        "HttpStatus",
        "GraphRequestId",
        "RawError"
    )
}

function Try-GetHardwareInfo {
    param(
        [Parameter(Mandatory)][string]$ManagedDeviceId
    )

    $attempts = @(
        @{ Name = "BETA_SELECT"; Uri = ("https://graph.microsoft.com/beta/deviceManagement/managedDevices/{0}?`$select=id,deviceName,azureADDeviceId,userPrincipalName,operatingSystem,osVersion,manufacturer,model,serialNumber,hardwareInformation" -f $ManagedDeviceId) },
        @{ Name = "BETA_EXPAND"; Uri = ("https://graph.microsoft.com/beta/deviceManagement/managedDevices/{0}?`$select=id,deviceName&`$expand=hardwareInformation" -f $ManagedDeviceId) },
        @{ Name = "V1_SELECT";   Uri = ("https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/{0}?`$select=id,deviceName,azureADDeviceId,userPrincipalName,operatingSystem,osVersion,manufacturer,model,serialNumber,hardwareInformation" -f $ManagedDeviceId) },
        @{ Name = "V1_EXPAND";   Uri = ("https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/{0}?`$select=id,deviceName&`$expand=hardwareInformation" -f $ManagedDeviceId) }
    )

    foreach ($a in $attempts) {
        try {
            $wrap = Invoke-GraphRequestWithRetry -Method GET -Uri $a.Uri
            $r = $wrap.Body

            if ($r -and $r.hardwareInformation) {
                return [pscustomobject]@{
                    Status     = "OK_$($a.Name)"
                    Response   = $r
                    HttpStatus = $wrap.StatusCode
                    RequestId  = $wrap.RequestId
                    RawError   = $null
                }
            } else {
                return [pscustomobject]@{
                    Status     = "NO_HWINFO_$($a.Name)"
                    Response   = $r
                    HttpStatus = $wrap.StatusCode
                    RequestId  = $wrap.RequestId
                    RawError   = $null
                }
            }
        } catch {
            $d = Get-GraphErrorDetails -ErrorRecord $_
            # Continue to next attempt
        }
    }

    # Qualify object readability with minimal GET (FIXED: include '?$select=')
    try {
        $minUri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/{0}?`$select=id,deviceName" -f $ManagedDeviceId
        $wrapMin = Invoke-GraphRequestWithRetry -Method GET -Uri $minUri
        $min = $wrapMin.Body

        return [pscustomobject]@{
            Status     = "UNSUPPORTED_HWINFO_BUT_READABLE"
            Response   = $min
            HttpStatus = $wrapMin.StatusCode
            RequestId  = $wrapMin.RequestId
            RawError   = $null
        }
    } catch {
        $d2 = Get-GraphErrorDetails -ErrorRecord $_
        return [pscustomobject]@{
            Status     = "UNREADABLE_MANAGEDDEVICE"
            Response   = $null
            HttpStatus = $d2.HttpStatus
            RequestId  = $d2.RequestId
            RawError   = $d2.RawError
        }
    }
}

# ==========================================================
# Initialization via SmartM365.Core
# ==========================================================
$ScriptVersion = "1.9"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DeviceBiosCsvLogFolderPath' -DefaultValue $OutputPath
try {
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPathInit $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
    Start-Transcript -Path $global:logTranscriptFile -Append

    WriteLog -Message "Script Environment initialized at $InitializeOutputPath" "INFO"
    $OutputPath = $InitializeOutputPath
    WriteLog -Message "Starting $TaskName..." "INFO"

    if ($DebugHardwareInfo) {
        WriteLog -Message "DebugHardwareInfo mode enabled. Will dump hardwareInformation properties for the first $DebugDeviceCount device(s) where hardwareInformation is retrieved." "INFO"
    }
} catch {
    Write-Host "Initialization failed: $_" -ForegroundColor Red
    exit 1
}

# ==========================================================
# MAIN TRY / CATCH / FINALLY
# ==========================================================
$connectedGraphInThisRun = $false
$results = New-Object System.Collections.Generic.List[object]

try {

    # ------------------------
    # Connect to Microsoft Graph via SmartM365.Core / Connect-SmartM365CloudSession (shared pattern)
    # ------------------------
    function Test-GraphConnection {
        try {
            $null = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -ErrorAction Stop
            return $true
        } catch {
            return $false
        }
    }

    $graphContext = $null
    if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
        try { $graphContext = Get-MgContext -ErrorAction SilentlyContinue } catch { }
    }

    $needConnect = $false

    if ($Connect) {
        WriteLog -Message "Connect switch specified: existing Graph session (if any) will be disconnected and reconnected..." "INFO"
        try { Disconnect-SmartM365CloudSession -ExchangeOnline $false -Graph $true -VerboseDisconnect:$true } catch { }
        $needConnect = $true
    } else {
        if ($graphContext -and (Test-GraphConnection)) {
            WriteLog -Message "Existing Microsoft Graph session detected. Reusing current connection." "INFO"
            $needConnect = $false
        } else {
            WriteLog -Message "No existing Graph session detected. Will establish a new connection..." "INFO"
            $needConnect = $true
        }
    }

    if ($needConnect) {
        $connectParams = @{
            ExchangeOnline = $false
            Graph          = $true
            GraphScopes    = @("DeviceManagementManagedDevices.Read.All")
        }

        if (-not $InteractiveAuth) {
            $connectParams.AppId        = $AppId
            $connectParams.Thumbprint   = $Thumb
            $connectParams.TenantId     = $TenantId
            $connectParams.Organization = $OrgDomain
            WriteLog -Message "Connecting to Microsoft Graph with app-only certificate authentication." "INFO"
        } else {
            WriteLog -Message "Connecting to Microsoft Graph with interactive authentication." "INFO"
        }

        $connectResult = Connect-SmartM365CloudSession @connectParams
        if (-not $connectResult.GraphConnected) {
            throw "Failed to connect to Microsoft Graph."
        }

        $connectedGraphInThisRun = $connectResult.GraphConnected
    }

    # ------------------------
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -RequiredGraphApplicationPermissions @('DeviceManagementManagedDevices.Read.All') -GraphProbeUris @('https://graph.microsoft.com/beta/deviceManagement/managedDevices?$top=1') | Out-Null

    # Retrieve Intune managed devices (Windows only)
    # ------------------------
    WriteLog -Message "Retrieving Intune managed Windows devices (LIST)..." "INFO"

    $listUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'&`$select=id,deviceName,azureADDeviceId,userPrincipalName,operatingSystem,osVersion,manufacturer,model,serialNumber,hardwareInformation&`$top=999"

    $allDevices = New-Object System.Collections.Generic.List[object]
    $nextLink = $listUri

    while ($nextLink) {
        $wrapPage = Invoke-GraphRequestWithRetry -Method GET -Uri $nextLink
        $page = $wrapPage.Body

        if ($page.value) {
            foreach ($d in $page.value) {
                if ($MaxItems -gt 0 -and $allDevices.Count -ge $MaxItems) { break }
                $allDevices.Add($d) | Out-Null
            }
            if ($MaxItems -gt 0 -and $allDevices.Count -ge $MaxItems) {
                WriteLog -Message ("MaxItems enabled: restricted managed devices to {0}." -f $allDevices.Count) "WARNING"
                break
            }
        }
        $nextLink = $page.'@odata.nextLink'
    }

    WriteLog -Message "Managed Windows devices retrieved: $($allDevices.Count)" "INFO"
    $bulkHardwareInfoCount = @($allDevices | Where-Object { $null -ne $_.hardwareInformation }).Count
    WriteLog -Message ("hardwareInformation returned by the bulk list for {0}/{1} devices; only missing entries will use per-device fallback calls." -f $bulkHardwareInfoCount, $allDevices.Count) "INFO"

    if ($DeviceNameContains) {
        $needle = $DeviceNameContains.Trim()
        if ($needle.Length -gt 0) {
            $filtered = $allDevices | Where-Object { $_.deviceName -and ($_.deviceName -like "*$needle*") }
            WriteLog -Message "DeviceNameContains filter applied: '$needle'. Remaining devices: $($filtered.Count)" "INFO"

            $tmp = New-Object System.Collections.Generic.List[object]
            foreach ($d in $filtered) { $tmp.Add($d) | Out-Null }
            $allDevices = $tmp
        }
    }

    if (-not $allDevices -or $allDevices.Count -eq 0) {
        WriteLog -Message "No Windows managed devices found. Exiting." "WARNING"
        return
    }

    $columns = Get-InventoryColumns

    $script:HwFailCount  = 0
    $script:DebugDumped  = 0

    $i = 0
    foreach ($d in $allDevices) {
        $i++
        $mdId = [string]$d.id
        $dn   = [string]$d.deviceName

        if ($i % 50 -eq 0) {
            WriteLog -Message "Progress: $i / $($allDevices.Count) devices processed..." "INFO"
        }

        $bios     = $null
        $biosRaw  = $null
        $biosDate = $null

        $hwStatus   = "OK"
        $httpStatus = $null
        $reqId      = $null
        $rawErr     = $null

        $hw = if ($null -ne $d.hardwareInformation) {
            [pscustomobject]@{
                Status     = 'OK_BETA_LIST'
                Response   = $d
                HttpStatus = 200
                RequestId  = $null
                RawError   = $null
            }
        } else {
            Try-GetHardwareInfo -ManagedDeviceId $mdId
        }

        $hwStatus   = $hw.Status
        $httpStatus = $hw.HttpStatus
        $reqId      = $hw.RequestId
        $rawErr     = $hw.RawError

        if ($hw.Response -and $hw.Response.hardwareInformation) {
            $bios     = $hw.Response.hardwareInformation.systemManagementBIOSVersion
            $biosRaw  = $hw.Response.hardwareInformation.biosVersion
            $biosDate = $hw.Response.hardwareInformation.biosReleaseDateTime

            # ------------------------
            # DebugHardwareInfo: dump all non-null properties of hardwareInformation
            # ------------------------
            if ($DebugHardwareInfo -and $script:DebugDumped -lt $DebugDeviceCount) {
                $script:DebugDumped++
                WriteLog -Message "[DEBUG_HW] Device $script:DebugDumped/$DebugDeviceCount : '$dn' ($mdId) - HardwareInfoStatus=$hwStatus" "INFO"

                $hwInfo = $hw.Response.hardwareInformation
                $props  = $null

                if ($hwInfo -is [System.Collections.IDictionary]) {
                    $props = $hwInfo.Keys | ForEach-Object { [pscustomobject]@{ Name = $_; Value = $hwInfo[$_] } }
                } elseif ($hwInfo -is [PSObject]) {
                    $props = $hwInfo.PSObject.Properties | Select-Object Name, Value
                }

                if ($props) {
                    $nonNullProps = $props | Where-Object { $null -ne $_.Value -and $_.Value -ne "" }
                    if ($nonNullProps) {
                        foreach ($p in $nonNullProps) {
                            WriteLog -Message "[DEBUG_HW]   $($p.Name) = $($p.Value)" "INFO"
                        }
                    } else {
                        WriteLog -Message "[DEBUG_HW]   (all hardwareInformation properties are null or empty)" "INFO"
                    }
                } else {
                    WriteLog -Message "[DEBUG_HW]   (unable to enumerate hardwareInformation properties)" "INFO"
                }
            }
        }

        if ($hwStatus -like "FAIL_*" -or $hwStatus -eq "UNREADABLE_MANAGEDDEVICE") {
            $script:HwFailCount++
            if ($script:HwFailCount -le 5 -or ($script:HwFailCount % 25 -eq 0)) {
                WriteLog -Message "hardwareInformation retrieval issue for '$dn' ($mdId): $hwStatus / $httpStatus / $rawErr" "WARNING"
            }
        }

        $row = [ordered]@{}
        foreach ($c in $columns) { $row[$c] = $null }

        $row["ManagedDeviceId"]             = $mdId
        $row["DeviceName"]                  = $dn
        $row["AzureADDeviceId"]             = $d.azureADDeviceId
        $row["UserPrincipalName"]           = $d.userPrincipalName
        $row["OperatingSystem"]             = $d.operatingSystem
        $row["OSVersion"]                   = $d.osVersion
        $row["Manufacturer"]                = $d.manufacturer
        $row["Model"]                       = $d.model
        $row["SerialNumber"]                = $d.serialNumber
        $row["SystemManagementBIOSVersion"] = $bios
        $row["BIOSVersionRaw"]              = $biosRaw
        $row["BIOSReleaseDateTime"]         = $biosDate
        $row["HardwareInfoStatus"]          = $hwStatus
        $row["HttpStatus"]                  = $httpStatus
        $row["GraphRequestId"]              = $reqId
        $row["RawError"]                    = $rawErr

        $results.Add([pscustomobject]$row) | Out-Null
    }

    WriteLog -Message "hardwareInformation failures/oddities total: $($script:HwFailCount)" "INFO"

    if ($DebugHardwareInfo) {
        WriteLog -Message "DebugHardwareInfo: $($script:DebugDumped) device(s) dumped out of $DebugDeviceCount requested." "INFO"
    }

    if (-not $results -or $results.Count -eq 0) {
        WriteLog -Message "No results generated. Exiting." "WARNING"
        return
    }

    # ------------------------
    # Export CSV (SmartM365)
    # ------------------------
    Write-Host "`n--- Export CSV ---"
$BaseFileName = "Intune_Devices_BIOS"

    ExportAndCopyCsv -BaseFileName $BaseFileName `
       -OutputPath $OutputPath `
       -GlobalPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') `
       -Data $results `
       -Encoding "UTF8" `
       -NoTypeInformation

    WriteLog -Message "Export completed: $global:csvFilePath1" "INFO"
    WriteLog -Message "Number of exported devices: $($results.Count)" "INFO"
}
catch {
    $globalError = $_
    WriteLog -Message ("Global error in Intune BIOS inventory: {0}" -f $globalError) "ERROR"
    Write-Host "A global error occurred. Check the log file for details." -ForegroundColor Red
}
finally {
    if ($connectedGraphInThisRun) {
        Write-Host "`n--- Disconnect Cloud Services ---"
        try {
            Disconnect-SmartM365CloudSession -ExchangeOnline:$false -Graph:$true -VerboseDisconnect:$true
        } catch {
            WriteLog -Message ("Error during Graph disconnect in finally: {0}" -f $_) "WARNING"
        }
    }

    try {
        RemoveOldFiles -Path $OutputPath     -Filter "*.csv" -KeepCount 150 -LogFile $global:LogTextFile
        RemoveOldFiles -Path $global:LogPath -Filter "*.log" -KeepCount $global:RetentionMaxLogs  -LogFile $global:LogTextFile
    } catch {
        WriteLog -Message ("Error during cleanup in finally: {0}" -f $_) "WARNING"
    }

    WriteLog -Message "$TaskName completed (finally block)." "INFO"

    try { Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {} } catch { }
    Complete-SmartM365ExecutionContext -Status Auto
}
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBndMI4wAdvLdQV
# usLH4nbNSYVwPa3GQVThaXju8L3tQKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAoVcv15TbvFB1CHx8+J2Gmrz6Y30mV+FJmLrDn+H2T/MA0GCSqG
# SIb3DQEBAQUABIIBgActPPhYlg2DMf/RpggN0Gp16icYpmtUkqeKZGtGYwJewx4r
# tP6hfP0dSvq5wyjCte33dbIRC5nUh4Ho1htMQ+sFfFe/NLUZvID9rCr4tGDxFxdJ
# vwAhWnLIyDQYHC49PTM+Gf9RsPj7VCGQYXd/f65m0WAZarzQT4BSNEd1hUuXIMF6
# SVCEMvOfM5GlMzzRW9x8JCyyaFrOZ0i6Yaqw0EE5B8TEa+1uELixMfmPCCooYbzW
# 5aVx8jcXqLa34DOw2u7eveeSfeM9OZn+b+I5PwGG8xfSets6AzcJ8zStLPCleGED
# kUkgxU4N+BjrOcWtynu329w5sJgA88YsErYucb7FJLgkGltLk/mkdOjx9dv5ausR
# JsHBSO61Ib1p6eJjKpRS75PkwBMb+ACv+4+0ONMBM6U1b+VXBr9pf9BWqE62GVM/
# MznDdtLDWvS0H3+AmC0x66Qf4fqMearWP9kc04cyhjX86G72S0xm7dhEckoqL/Ld
# +w5Qn61XjPYAuO1l2qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODUw
# MzBaMC8GCSqGSIb3DQEJBDEiBCCmwAVgmtHI//n1uUs95+zuFKrtARQ+Y5B50VAD
# cE5YIzANBgkqhkiG9w0BAQEFAASCAgCksr3ZTB5zQdG72x7ulkvVuGlwYd4p3/X0
# rYeup3GihLk708KtxI2TrePmXmo3aQoWVMsBOZXJQScPc0OEFeg8oaD7p9hr8ndP
# SqteBU3AUR0YYsAWDxxWmCjZUgYGMtlU5JX7Ku7epNULyc6/pm0iQ9AWwEbupxW6
# px379WXhI+f6qxTYgv/ljlDNwVhnPe3hijaLJp73dVpl2MLx7T/Y2F3BFp4Cpo7J
# UXOY8mR5SZU+aDsDXhq8eUEV888+W0eCtsRHUc7jkhCZGNO70tkBXAmQf+4wCP0d
# PgyU/gyRgK3ZHN8q5uFujiZKEXsr2BJmToRorjPNkv4NN3pTuy0WoBCWV5G6IsPH
# QUH0QS69BLwC/m+9OIhW5u9zWVNo9usuz5aU5XiOCMu9eKx0akkDLC8UvDmS2shI
# prWpGW1DON9EFgYADYdu1gz1JrdqnQ1ZZG1QT6ciKu/e2z8A6o7KHE0c/F95X7Is
# l2QzBG7Ri5C6hcDt9bIRg76qBjl3jFkcD4U6+OSmR8/IHNFJP3xQK82uVjwqxZBg
# e9or9jluMEA25IX+P99jq66h/fhd76wzNnVtcGkImHhxNn4HiB+OFYD0W0wO2eyV
# gkJDdovy9E/qxZXq2CdnNVrBXT8tCk9o5BabIX2DJe05a5hE1UzJGY5ZD7FMBml6
# H8DV0G0YDw==
# SIG # End signature block
