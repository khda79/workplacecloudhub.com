<#
.SYNOPSIS
Generates a detailed inventory of Intune-managed devices using Microsoft Graph API (Windows-only) and includes PhysicalMemoryGB .

.DESCRIPTION
Connects to Microsoft Graph, retrieves all Windows managed devices from Intune,
adds PhysicalMemoryGB (RAM) via Graph JSON batching with $select=physicalMemoryInBytes and a per-device fallback.
Exports results to CSV.

.PARAMETER OutputPath
Specifies the output directory for the generated CSV and log files.

.PARAMETER Connect
Forces a (re)connection to Microsoft Graph (disconnects any existing session first).

.PARAMETER InteractiveAuth
Uses interactive authentication instead of app-only certificate authentication.
.VERSION
1.4


.NOTES
    Version : 1.2
    Author: https://github.com/khda79/workplacecloudhub.com
Requires: SmartM365.Core module (logging, init, CSV, cleanup, cloud connectivity)
Scopes: DeviceManagementManagedDevices.Read.All
#>

param(
    [string]$Tenant = 'test',
[string]$OutputPath,
    [switch]$Connect,
    [switch]$InteractiveAuth
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
$modulePath = & { $d = $PSScriptRoot; while ($d) { $p = Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'; if (Test-Path -LiteralPath $p) { return $p }; $parent = Split-Path -Path $d -Parent; if ($parent -eq $d) { break }; $d = $parent }; throw 'SmartM365.Core module not found.' }
try {
    Import-Module $modulePath -ErrorAction Stop
} catch {
    Write-Host "Failed to import SmartM365.Core module from '$modulePath' : $_" -ForegroundColor Red
    exit 1
}

# ==========================================================
# Helpers
# ==========================================================
function Get-SafeProperty {
    param(
        [Parameter(Mandatory)][object]$Object,
        [Parameter(Mandatory)][string]$Name
    )

    if (-not $Object) { return $null }

    # 1) Direct property (case-sensitive)
    $prop = $Object.PSObject.Properties[$Name]
    if ($prop) { return $prop.Value }

    # 2) Direct property (case-insensitive)
    $propCI = $Object.PSObject.Properties |
        Where-Object { $_.Name -ieq $Name } |
        Select-Object -First 1
    if ($propCI) { return $propCI.Value }

    # 3) AdditionalProperties dictionary
    $apProp = $Object.PSObject.Properties['AdditionalProperties']
    if ($apProp -and $apProp.Value -is [System.Collections.IDictionary]) {
        $dict = $apProp.Value

        # 3a) Exact key
        if ($dict.ContainsKey($Name)) {
            return $dict[$Name]
        }

        # 3b) Case-insensitive key (e.g. bitLockerStatus vs BitLockerStatus)
        $matchKey = $dict.Keys |
            Where-Object { $_ -is [string] -and $_ -ieq $Name } |
            Select-Object -First 1

        if ($matchKey) {
            return $dict[$matchKey]
        }
    }

    return $null
}

# Helper: RAM lookup per device (bytes) via unit GET + $select=physicalMemoryInBytes
function Get-ManagedDeviceRamBytes {
    param([Parameter(Mandatory)][string]$ManagedDeviceId)
    try {
        $d = Get-MgDeviceManagementManagedDevice -ManagedDeviceId $ManagedDeviceId -Property "id,physicalMemoryInBytes"
        if ($d.PSObject.Properties.Match('physicalMemoryInBytes').Count -gt 0 -and [int64]$d.physicalMemoryInBytes -gt 0) {
            return [int64]$d.physicalMemoryInBytes
        }
        if ($d.AdditionalProperties -and $d.AdditionalProperties.ContainsKey('physicalMemoryInBytes')) {
            $val = [int64]$d.AdditionalProperties['physicalMemoryInBytes']
            if ($val -gt 0) { return $val }
        }
    } catch {
        WriteLog -Message "RAM lookup failed for device $ManagedDeviceId - $($_.Exception.Message)" "WARNING"
    }
    return $null
}


# Helper: Bulk RAM lookup via Graph JSON batching ($batch, max 20 sub-requests per call)
# Returns hashtable: managedDeviceId -> physicalMemoryInBytes ([int64])
function Get-ManagedDeviceRamMap {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string[]]$ManagedDeviceIds,
        [int]$BatchSize  = 20,
        [int]$MaxRetries = 5
    )

    $ids = @($ManagedDeviceIds | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    $ramMap = @{}
    $total  = $ids.Count
    if ($total -eq 0) { return $ramMap }

    if ($BatchSize -lt 1) { $BatchSize = 20 }
    if ($BatchSize -gt 20) { $BatchSize = 20 }

    WriteLog -Message ("RAM bulk lookup via Graph batching: {0} devices, batch size {1}..." -f $total, $BatchSize) "INFO"

    $batchUri = "https://graph.microsoft.com/v1.0/" + '$batch'
    $swTotal = [System.Diagnostics.Stopwatch]::StartNew()
    $batchIdx = 0

    for ($i = 0; $i -lt $total; $i += $BatchSize) {
        $endIdx = [math]::Min($i + $BatchSize - 1, $total - 1)
        $pending = @($ids[$i..$endIdx])
        $attempt = 0
        $batchIdx++

        while ($pending.Count -gt 0 -and $attempt -le $MaxRetries) {
            $idToDevice = @{}
            $requests = New-Object System.Collections.Generic.List[object]
            $j = 1

            foreach ($id in $pending) {
                $localId = [string]$j
                $idToDevice[$localId] = $id
                $requests.Add(@{
                    id     = $localId
                    method = "GET"
                    url    = "/deviceManagement/managedDevices/$id" + '?$select=id,physicalMemoryInBytes'
                }) | Out-Null
                $j++
            }

            $body = @{ requests = $requests } | ConvertTo-Json -Depth 5

            try {
                $resp = Invoke-MgGraphRequest -Method POST -Uri $batchUri -Body $body -ContentType "application/json" -ErrorAction Stop
            }
            catch {
                $attempt++
                WriteLog -Message ("Graph batch RAM lookup failed (attempt {0}/{1}): {2}. Waiting 10s..." -f $attempt, $MaxRetries, $_.Exception.Message) "WARNING"
                Start-Sleep -Seconds 10
                continue
            }

            $retryIds = New-Object System.Collections.Generic.List[string]
            $retryAfter = 0

            foreach ($r in @($resp.responses)) {
                $managedDeviceId = $idToDevice[[string]$r.id]

                if ([int]$r.status -eq 200) {
                    $bytes = 0L
                    if ($r.body -and $r.body.physicalMemoryInBytes) {
                        $bytes = [int64]$r.body.physicalMemoryInBytes
                    }
                    if ($bytes -gt 0) {
                        $mapKey = if ($r.body.id) { [string]$r.body.id } else { [string]$managedDeviceId }
                        $ramMap[$mapKey] = $bytes
                    }
                }
                elseif ([int]$r.status -eq 429 -or [int]$r.status -ge 500) {
                    $retryIds.Add($managedDeviceId) | Out-Null
                    if ($r.headers -and $r.headers.'Retry-After') {
                        $ra = 0
                        if ([int]::TryParse([string]$r.headers.'Retry-After', [ref]$ra) -and $ra -gt $retryAfter) {
                            $retryAfter = $ra
                        }
                    }
                }
                else {
                    WriteLog -Message ("RAM lookup failed for device {0} - HTTP {1}" -f $managedDeviceId, $r.status) "WARNING"
                }
            }

            $pending = @($retryIds)
            if ($pending.Count -gt 0) {
                $attempt++
                $wait = if ($retryAfter -gt 0) { $retryAfter } else { 5 * $attempt }
                WriteLog -Message ("Throttled/transient RAM lookup errors on {0} sub-request(s). Waiting {1}s before retry (attempt {2}/{3})..." -f $pending.Count, $wait, $attempt, $MaxRetries) "WARNING"
                Start-Sleep -Seconds $wait
            }
        }

        foreach ($id in $pending) {
            $rb = Get-ManagedDeviceRamBytes -ManagedDeviceId $id
            if ($rb) { $ramMap[[string]$id] = $rb }
        }

        $processed = $endIdx + 1
        if (($batchIdx % 50) -eq 0 -or $processed -ge $total) {
            $pct = [math]::Round(100 * $processed / $total)
            WriteLog -Message ("RAM bulk lookup progress: {0} / {1} ({2}%)" -f $processed, $total, $pct) "INFO"
        }
    }

    $swTotal.Stop()
    WriteLog -Message ("RAM bulk lookup completed: {0} / {1} devices with RAM value in {2} s." -f $ramMap.Count, $total, [math]::Round($swTotal.Elapsed.TotalSeconds)) "INFO"

    return $ramMap
}
# Helper: Build Entra device map (deviceId GUID -> {Id,DeviceId,DisplayName,ApproximateLastSignInDateTime})
function Get-EntraDeviceMap {
    [CmdletBinding()]
    param()

    WriteLog -Message "Retrieving Entra ID devices (Id, deviceId, displayName, approximateLastSignInDateTime)..." "INFO"

    $map = @{}
    try {
        $entraDevices = Get-MgDevice -All -Property "id,deviceId,displayName,approximateLastSignInDateTime"

        foreach ($d in $entraDevices) {
            $did = Get-SafeProperty $d "DeviceId"
            if ([string]::IsNullOrWhiteSpace($did)) { continue }

            # Normalize as string for hashtable key
            $key = [string]$did

            if (-not $map.ContainsKey($key)) {
                $map[$key] = [PSCustomObject]@{
                    Id                           = Get-SafeProperty $d "Id"
                    DeviceId                      = $key
                    DisplayName                   = Get-SafeProperty $d "DisplayName"
                    ApproximateLastSignInDateTime = Get-SafeProperty $d "ApproximateLastSignInDateTime"
                }
            }
        }

        WriteLog -Message ("Entra devices retrieved: {0}" -f $map.Count) "INFO"
    }
    catch {
        WriteLog -Message "Failed to retrieve Entra ID devices - $($_.Exception.Message)" "ERROR"
        throw
    }

    return $map
}

# Returns the fixed inventory schema (column names) in the exact order
function Get-InventoryColumns {
    [string[]]@(
        "Device ID",
        "Device name",
        "Enrollment date",
        "Last check-in",
        "LastSyncDateTime",

        "Intune registered",
        "Azure AD Device ID",
        "Azure AD registered",

        "OS",
        "OS version",
        "Manufacturer",
        "Model",
        "Serial number",
        "Supervised",
        "Encrypted",

        "Compliance",
        "IsCompliant",
        "Device state",
        "Managed by",
        "Ownership",

        "Primary user UPN",
        "Primary user email address",
        "Primary user display name",
        "UserId",
        "Registered owners",

        "EAS activation ID",
        "EAS activated",
        "Last EAS sync time",
        "EAS status",
        "EAS reason",

        "Total storage",
        "Free storage",
        "Management name",
        "Category",

        "PhysicalMemoryGB",


        "Entra ObjectId",
        "Entra DeviceId",
        "Entra DisplayName",
        "Entra ApproximateLastSignInDateTime"
    )
}

# ==========================================================
# Initialization via SmartM365.Core
# ==========================================================
$ScriptVersion = "1.4"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DeviceUsersCsvLogFolderPath' -DefaultValue $OutputPath
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
$results = @()

try {

    # ------------------------
    # Connect to Microsoft Graph (direct Connect-MgGraph; avoids submodule load issues)
    # ------------------------
    function Test-GraphConnection {
        try {
            $org = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -ErrorAction Stop
            return $true
        } catch {
            return $false
        }
    }

    # Detect existing Graph session
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
            # App-only certificate authentication (no -Scopes required for app-only; permissions are granted on the app)
            WriteLog -Message "Connecting to Microsoft Graph with app-only certificate authentication (direct Connect-MgGraph)." "INFO"
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
            Connect-MgGraph -TenantId $TenantId -ClientId $AppId -CertificateThumbprint $Thumb -NoWelcome | Out-Null
        } else {
            # Interactive authentication (requires delegated scopes)
            WriteLog -Message "Connecting to Microsoft Graph with interactive authentication (direct Connect-MgGraph)." "INFO"
            try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }
            Connect-MgGraph -Scopes @("DeviceManagementManagedDevices.Read.All","Device.Read.All") -NoWelcome | Out-Null
        }

        # Basic connectivity validation
        if (-not (Test-GraphConnection)) {
            throw "Failed to connect to Microsoft Graph."
        }

        $connectedGraphInThisRun = $true
    }

    # ------------------------
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -GraphProbeUris @(
        'https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?$top=1',
        'https://graph.microsoft.com/v1.0/devices?$top=1'
    ) | Out-Null

    # Retrieve Intune managed devices (Windows only)
    # ------------------------
    WriteLog -Message "Retrieving Intune managed Windows devices..."
    $devices = Get-MgDeviceManagementManagedDevice -All -Filter "operatingSystem eq 'Windows'"
    WriteLog -Message "Managed Windows devices retrieved: $($devices.Count)"

    # ------------------------
    # Retrieve Entra devices map (for approximateLastSignInDateTime and Entra objectId/displayName)
    # Requires Graph permission: Device.Read.All (or Directory.Read.All)
    # ------------------------
    $entraMap = Get-EntraDeviceMap

    # ------------------------
    # Build result objects
    # ------------------------
    $deviceIds = @($devices | ForEach-Object { Get-SafeProperty $_ 'Id' } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $ramMap = Get-ManagedDeviceRamMap -ManagedDeviceIds $deviceIds

    $results = $devices | ForEach-Object {
        $devId   = Get-SafeProperty $_ 'Id'
        $physicalMemoryGB = $null
        if ($devId -and $ramMap.ContainsKey([string]$devId)) {
            $physicalMemoryGB = [math]::Round([double]$ramMap[[string]$devId] / 1GB, 2)
        }
        # Entra correlation (AzureADDeviceId == Entra deviceId GUID)
        $entraInfo = $null
        $aadDeviceId = Get-SafeProperty $_ 'AzureADDeviceId'
        if ($aadDeviceId) {
            $key = [string]$aadDeviceId
            if ($entraMap -and $entraMap.ContainsKey($key)) {
                $entraInfo = $entraMap[$key]
            }
        }


        [PSCustomObject]@{
            "Device ID"                              = $devId
            "Device name"                            = Get-SafeProperty $_ 'DeviceName'
            "Enrollment date"                        = Get-SafeProperty $_ 'EnrolledDateTime'
            "Last check-in"                          = Get-SafeProperty $_ 'LastSyncDateTime'
            "LastSyncDateTime"                       = Get-SafeProperty $_ 'LastSyncDateTime'

            "Intune registered"                      = [bool](Get-SafeProperty $_ 'EnrolledDateTime')
            "Azure AD Device ID"                     = Get-SafeProperty $_ 'AzureADDeviceId'
            "Azure AD registered"                    = Get-SafeProperty $_ 'AzureADRegistered'

            "OS"                                     = Get-SafeProperty $_ 'OperatingSystem'
            "OS version"                             = Get-SafeProperty $_ 'OsVersion'
            "Manufacturer"                           = Get-SafeProperty $_ 'Manufacturer'
            "Model"                                  = Get-SafeProperty $_ 'Model'
            "Serial number"                          = Get-SafeProperty $_ 'SerialNumber'
            "Supervised"                             = Get-SafeProperty $_ 'IsSupervised'
            "Encrypted"                              = Get-SafeProperty $_ 'IsEncrypted'

            "Compliance"                             = Get-SafeProperty $_ 'ComplianceState'
            "IsCompliant"                            = ((Get-SafeProperty $_ 'ComplianceState') -eq 'compliant')
            "Device state"                           = Get-SafeProperty $_ 'DeviceRegistrationState'
            "Managed by"                             = Get-SafeProperty $_ 'ManagementAgent'
            "Ownership"                              = Get-SafeProperty $_ 'ManagedDeviceOwnerType'

            "Primary user UPN"                       = Get-SafeProperty $_ 'UserPrincipalName'
            "Primary user email address"             = Get-SafeProperty $_ 'EmailAddress'
            "Primary user display name"              = Get-SafeProperty $_ 'UserDisplayName'
            "UserId"                                 = Get-SafeProperty $_ 'UserId'
            "Registered owners"                      = $null

            "EAS activation ID"                      = Get-SafeProperty $_ 'EasDeviceId'
            "EAS activated"                          = Get-SafeProperty $_ 'EasActivated'
            "Last EAS sync time"                     = Get-SafeProperty $_ 'ExchangeLastSuccessfulSyncDateTime'
            "EAS status"                             = Get-SafeProperty $_ 'ExchangeAccessState'
            "EAS reason"                             = Get-SafeProperty $_ 'ExchangeAccessStateReason'

            "Total storage"                          = Get-SafeProperty $_ 'TotalStorageSpaceInBytes'
            "Free storage"                           = Get-SafeProperty $_ 'FreeStorageSpaceInBytes'
            "Management name"                        = Get-SafeProperty $_ 'ManagedDeviceName'
            "Category"                               = Get-SafeProperty $_ 'DeviceCategoryDisplayName'

            "PhysicalMemoryGB"                       = $physicalMemoryGB

            "Entra ObjectId"                         = if ($entraInfo) { $entraInfo.Id } else { $null }
            "Entra DeviceId"                         = if ($entraInfo) { $entraInfo.DeviceId } else { $null }
            "Entra DisplayName"                      = if ($entraInfo) { $entraInfo.DisplayName } else { $null }
            "Entra ApproximateLastSignInDateTime"    = if ($entraInfo) { $entraInfo.ApproximateLastSignInDateTime } else { $null }
        }
    }

    # Determine column order baseline
    $columnOrder = @()
    if ($results -and $results.Count -gt 0) {
        $columnOrder = $results[0].PSObject.Properties.Name
    } else {
        $columnOrder = Get-InventoryColumns
    }

    # ------------------------
    # Deduplicate by "Device name" (keep most recent Last check-in)
    # ------------------------
    $sorted = $results | Sort-Object `
        @{ Expression = { -not $_.'Intune registered' } }, `
        @{ Expression = {
            if ($_.PSObject.Properties['Last check-in'] -and $_.'Last check-in') {
                [datetime]$_.'Last check-in'
            } else {
                [datetime]'1900-01-01'
            }
        }; Descending = $true }

    $seenNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $deduped   = New-Object System.Collections.Generic.List[object]

    foreach ($row in $sorted) {
        $name = $null
        if ($row.PSObject.Properties['Device name']) { $name = ($row.'Device name' -as [string]) }
        $name = if ($name) { $name.Trim() } else { $null }

        if ([string]::IsNullOrWhiteSpace($name)) {
            $deduped.Add($row) | Out-Null
        } elseif ($seenNames.Add($name)) {
            $deduped.Add($row) | Out-Null
        }
    }

    $dupCount = ($results.Count - $deduped.Count)
    WriteLog -Message "Dedup by Device name: removed $dupCount duplicate row(s). Kept $($deduped.Count)." "INFO"

    $results = $deduped

    # ------------------------
    # If no devices (in the chosen scope), exit gracefully
    # ------------------------
    if (-not $results -or $results.Count -eq 0) {
        WriteLog -Message "No devices found. Exiting." "WARNING"
        return
    }

    # ------------------------
    # Export CSV
    # ------------------------
    Write-Host "`n--- Export CSV ---"
$BaseFileName = "Intune_Devices_Inventory"

    ExportAndCopyCsv -BaseFileName $BaseFileName `
       -OutputPath $OutputPath `
       -GlobalPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') `
       -Data $results `
       -Encoding "UTF8" `
       -NoTypeInformation

    WriteLog -Message "Export completed: $global:csvFilePath1"
    WriteLog -Message "Number of exported devices: $($results.Count)"

}
catch {
    $globalError = $_
    WriteLog -Message ("Global error in Intune devices inventory: {0}" -f $globalError) "ERROR"
    Write-Host "A global error occurred. Check the log file for details." -ForegroundColor Red

    # -------- Envoi de mail en cas d'erreur globale --------
    try {
        $title = "Intune devices inventory - ERROR"
        $msg   = @"
An error occurred in script $($MyInvocation.MyCommand.Name) on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss").

Error message:
$($globalError.Exception.Message)

See attached log file for details:
$($global:LogTextFile)
"@

        # Genere un HTML simple via SmartM365.Core
        $bodyHtml = NewSimpleEmailBody -Title $title -Message $msg

        # Envoi du mail avec le log en piece jointe
        $attachments = @()
        if ($global:LogTextFile -and (Test-Path $global:LogTextFile)) {
            $attachments = @($global:LogTextFile)
        }

        SendEmailHtmlReport -BodyHtml $bodyHtml -Subject $title -Attachments $attachments -VerboseLog
    }
    catch {
        WriteLog -Message ("Failed to send error notification email: {0}" -f $_) "ERROR"
    }
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
