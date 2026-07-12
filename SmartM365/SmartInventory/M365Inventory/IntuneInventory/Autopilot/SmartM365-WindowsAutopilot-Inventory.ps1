<#
.SYNOPSIS
Generates a detailed inventory of Windows Autopilot devices using Microsoft Graph API.

.DESCRIPTION
Connects to Microsoft Graph, retrieves Windows Autopilot device identities,
optionally enriches them with Entra ID device information,
and exports the result to CSV.

.PARAMETER OutputPath
Specifies the output directory for the generated CSV and log files.

.PARAMETER Connect
Forces a (re)connection to Microsoft Graph (disconnects any existing session first).

.PARAMETER InteractiveAuth
Uses interactive authentication instead of app-only certificate authentication.
.VERSION
1.8



.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication.
    Minimum Graph application permissions: DeviceManagementServiceConfig.Read.All; Device.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Version : 1.7
    Author: https://github.com/khda79/workplacecloudhub.com
Requires: SmartM365.Core module (logging, init, CSV, cleanup, cloud connectivity)
Scopes: DeviceManagementServiceConfig.Read.All
    Minimum application permissions: DeviceManagementServiceConfig.Read.All, Device.Read.All
#>

param(
    [string]$Tenant = 'test',
[string]$OutputPath,
    [switch]$Connect,
    [switch]$InteractiveAuth,
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
    Import-Module -Name $modulePath -MinimumVersion '1.0.23' -ErrorAction Stop
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

        # 3b) Case-insensitive key
        $matchKey = $dict.Keys |
            Where-Object { $_ -is [string] -and $_ -ieq $Name } |
            Select-Object -First 1

        if ($matchKey) {
            return $dict[$matchKey]
        }
    }

    return $null
}

# Helper: RAM lookup per device (kept for compatibility with SmartM365 script pattern)
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

            $key = [string]$did

            if (-not $map.ContainsKey($key)) {
                $map[$key] = [PSCustomObject]@{
                    Id                           = Get-SafeProperty $d "Id"
                    DeviceId                     = $key
                    DisplayName                  = Get-SafeProperty $d "DisplayName"
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
        "Autopilot ID",
        "Display name",
        "Serial number",
        "Manufacturer",
        "Model",
        "Group tag",
        "Purchase order identifier",
        "Enrollment state",
        "Last contacted",
        "Addressable user name",
        "User principal name",
        "Resource name",
        "SKU number",
        "System family",
        "Azure AD Device ID",
        "Managed device ID",
        "Entra ObjectId",
        "Entra DeviceId",
        "Entra DisplayName",
        "Entra ApproximateLastSignInDateTime"
    )
}

# ==========================================================
# Initialization via SmartM365.Core
# ==========================================================
$ScriptVersion = "1.8"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AutopilotDevicesCsvLogFolderPath' -DefaultValue $OutputPath
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
            $graphContext = $null
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
            Connect-MgGraph -Scopes @("DeviceManagementServiceConfig.Read.All","Device.Read.All") -NoWelcome | Out-Null
        }

        # Basic connectivity validation
        if (-not (Test-GraphConnection)) {
            throw "Failed to connect to Microsoft Graph."
        }

        $connectedGraphInThisRun = $true
    }

    # ------------------------
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -RequiredGraphApplicationPermissions @('DeviceManagementServiceConfig.Read.All','Device.Read.All') -GraphProbeUris @(
        'https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?$top=1',
        'https://graph.microsoft.com/v1.0/devices?$top=1'
    ) | Out-Null

    # Retrieve Windows Autopilot devices
    # ------------------------
    WriteLog -Message "Retrieving Windows Autopilot device identities..."
    if ($MaxItems -gt 0) {
        WriteLog -Message ("MaxItems enabled: retrieving at most {0} Windows Autopilot devices from Graph." -f $MaxItems) "WARNING"
        $devices = @(Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -Top $MaxItems)
    }
    else {
        $devices = Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -All
    }
    WriteLog -Message "Windows Autopilot devices retrieved: $($devices.Count)"

    # ------------------------
    # Retrieve Entra devices map (for approximateLastSignInDateTime and Entra objectId/displayName)
    # Requires Graph permission: Device.Read.All (or Directory.Read.All)
    # ------------------------
    $entraMap = Get-EntraDeviceMap

    # ------------------------
    # Build result objects
    # ------------------------
    $results = $devices | ForEach-Object {
        $entraInfo = $null
        $aadDeviceId = Get-SafeProperty $_ 'AzureActiveDirectoryDeviceId'

        if ($aadDeviceId) {
            $key = [string]$aadDeviceId
            if ($entraMap -and $entraMap.ContainsKey($key)) {
                $entraInfo = $entraMap[$key]
            }
        }

        [PSCustomObject]@{
            "Autopilot ID"                        = Get-SafeProperty $_ 'Id'
            "Display name"                        = Get-SafeProperty $_ 'DisplayName'
            "Serial number"                       = Get-SafeProperty $_ 'SerialNumber'
            "Manufacturer"                        = Get-SafeProperty $_ 'Manufacturer'
            "Model"                               = Get-SafeProperty $_ 'Model'
            "Group tag"                           = Get-SafeProperty $_ 'GroupTag'
            "Purchase order identifier"           = Get-SafeProperty $_ 'PurchaseOrderIdentifier'
            "Enrollment state"                    = Get-SafeProperty $_ 'EnrollmentState'
            "Last contacted"                      = Get-SafeProperty $_ 'LastContactedDateTime'
            "Addressable user name"               = Get-SafeProperty $_ 'AddressableUserName'
            "User principal name"                 = Get-SafeProperty $_ 'UserPrincipalName'
            "Resource name"                       = Get-SafeProperty $_ 'ResourceName'
            "SKU number"                          = Get-SafeProperty $_ 'SkuNumber'
            "System family"                       = Get-SafeProperty $_ 'SystemFamily'
            "Azure AD Device ID"                  = $aadDeviceId
            "Managed device ID"                   = Get-SafeProperty $_ 'ManagedDeviceId'

            "Entra ObjectId"                      = if ($entraInfo) { $entraInfo.Id } else { $null }
            "Entra DeviceId"                      = if ($entraInfo) { $entraInfo.DeviceId } else { $null }
            "Entra DisplayName"                   = if ($entraInfo) { $entraInfo.DisplayName } else { $null }
            "Entra ApproximateLastSignInDateTime" = if ($entraInfo) { $entraInfo.ApproximateLastSignInDateTime } else { $null }
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
    # Deduplicate by "Serial number" (keep most recent Last contacted)
    # ------------------------
    $sorted = $results | Sort-Object @{
        Expression = {
            if ($_.PSObject.Properties['Last contacted'] -and $_.'Last contacted') {
                [datetime]$_.'Last contacted'
            } else {
                [datetime]'1900-01-01'
            }
        }
        Descending = $true
    }

    $seenSerials = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    $deduped     = New-Object System.Collections.Generic.List[object]

    foreach ($row in $sorted) {
        $serial = $null
        if ($row.PSObject.Properties['Serial number']) {
            $serial = ($row.'Serial number' -as [string])
        }
        $serial = if ($serial) { $serial.Trim() } else { $null }

        if ([string]::IsNullOrWhiteSpace($serial)) {
            $deduped.Add($row) | Out-Null
        } elseif ($seenSerials.Add($serial)) {
            $deduped.Add($row) | Out-Null
        }
    }

    $dupCount = ($results.Count - $deduped.Count)
    WriteLog -Message "Dedup by Serial number: removed $dupCount duplicate row(s). Kept $($deduped.Count)." "INFO"

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
$BaseFileName = "Intune_Autopilot_Devices"

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
    WriteLog -Message ("Global error in Windows Autopilot inventory: {0}" -f $globalError) "ERROR"
    Write-Host "A global error occurred. Check the log file for details." -ForegroundColor Red

    # -------- Envoi de mail en cas d'erreur globale --------
    try {
        $title = "Windows Autopilot inventory - ERROR"
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
# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAJMux3tfxsUlfC
# BMBbGXnCaxkpwCTROtNkLEXCIWBW6KCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCCUWBcNLaPqY/ZXyTRB35ovlY9jH8sw7QTE+A1crXsSSTANBgkqhkiG9w0B
# AQEFAASCAYCfVG2xxip2uT085vLus0Itmt0CSLLN8iwZg/YhzisTmR7TgLC8URaR
# k1ibSRe1Hf/2MxHJ95gP6u+a0x2/p0u0Rp0kHHu9Cmq2q0k4ydNp34N8tGdMU9YI
# s3sp9frwGm5/6pt2e8C6mS3rQKtPqyVbgD7lP1nEZZZ3H1RNe/TX1tsTHyHWY/7d
# Qk+HjQjaKGS9EZ2Vjya6NgGitm1yvu+anNnfbt8SgWCmFifq8GSPIp7M9f7alL4c
# 4/P+udEqmGULjiLGE2d2wsaf7/+eNo2cwuMAtspxvKhAQnrPMicz7dy6oNezj6qa
# ThutDCkUn7mCCsCUJIUkDEM9LmNqN1x4QAWfktWsX9gK/KVVYyLlMszPEs0T31LY
# v/cofFd0dKu2+4XFf0gwXE9VJTufc3hJZPwduRjxSn9PMq7EMqLOExXMmehz3AXe
# QnRDPcjhaVTOKfT4N8AWAy6Snck11cZqs6P3rw2tGLVgSSHmCOFGoW5AHMTxYBar
# aqbrzCRT6dM=
# SIG # End signature block
