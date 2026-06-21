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
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
Version: 1.0
#>

param(
    [string]$Tenant = 'test',
[string]$OutputPath,
    [switch]$Connect,
    [switch]$InteractiveAuth,
    [string]$DeviceNameContains,
    [switch]$DebugHardwareInfo,
    [int]$DebugDeviceCount = 3
)
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
            Write-Host 'Edit the local JSON now if needed. Press Enter to continue with the current file values.' -ForegroundColor Yellow
            Read-Host 'Press Enter to continue'
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
    Import-Module $ModulePath -ErrorAction Stop
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
$ScriptVersion = "1.0"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DeviceBiosCsvLogFolderPath' -DefaultValue $OutputPath
try {
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPathInit $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
    Start-Transcript -Path $global:logTranscriptFile -Append

    WriteLog -Message "Script Environment initialized at $InitializeOutputPath" "INFO"
    $OutputPath = $InitializeOutputPath
    WriteLog -Message "Starting $TaskName..." "INFO"
    WriteLog -Message "PowerShell Version: $($PSVersionTable.PSVersion)" "INFO"

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
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -GraphProbeUris @('https://graph.microsoft.com/beta/deviceManagement/managedDevices?$top=1') | Out-Null

    # Retrieve Intune managed devices (Windows only)
    # ------------------------
    WriteLog -Message "Retrieving Intune managed Windows devices (LIST)..." "INFO"

    $listUri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'&`$select=id,deviceName,azureADDeviceId,userPrincipalName,operatingSystem,osVersion,manufacturer,model,serialNumber&`$top=999"

    $allDevices = New-Object System.Collections.Generic.List[object]
    $nextLink = $listUri

    while ($nextLink) {
        $wrapPage = Invoke-GraphRequestWithRetry -Method GET -Uri $nextLink
        $page = $wrapPage.Body

        if ($page.value) {
            foreach ($d in $page.value) { $allDevices.Add($d) | Out-Null }
        }
        $nextLink = $page.'@odata.nextLink'
    }

    WriteLog -Message "Managed Windows devices retrieved: $($allDevices.Count)" "INFO"

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

        $hw = Try-GetHardwareInfo -ManagedDeviceId $mdId

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
}
