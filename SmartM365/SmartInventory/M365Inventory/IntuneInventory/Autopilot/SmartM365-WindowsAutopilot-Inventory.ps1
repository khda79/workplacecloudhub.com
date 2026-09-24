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
1.10



.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication.
    Minimum Graph application permissions: DeviceManagementServiceConfig.Read.All; Device.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Version : 1.10
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
    Import-Module -Name $modulePath -MinimumVersion '1.0.57' -ErrorAction Stop
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
$ScriptVersion = "1.10"
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
        Remove-SmartM365TimestampedFilesOlderThan -FolderPath $OutputPath -FilePattern '*.csv' -RetentionDays 7 -RequireCurrentRunPublication -LogFile $global:LogTextFile
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAe0NYZeWOApULq
# +G7GpHPzHA+S0LfyV8/ME0pPSkAZ5KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJk19oJ3/1w9fDCW3eYK+gY2b+TsDva2+OWF9Uj5PFWIMA0GCSqG
# SIb3DQEBAQUABIIBgIWUMRldOxScGPuay2+whNM46VEXty8Eu3UpsM9oKnXtTzF1
# KcKzi3fFMlxsvf+/ZW0LAoySPVid0mTTGYyMLBB88zWsnlsKN3dLIIiQhwYgisda
# UAEo5PJ8l5+7WJwdmbC5ADUZK9XX/G9XwwLxW/5TqBy80kEGL9RzUGQ8BWUurkGg
# jLrybE9puoaKNAaINgg+TjWGWXFH2/rCcDcoemXj2wYDWQdJ8GrrtRp/PmFeoLNg
# Zy1McLp60N8p/oFMRynaFjHK/F3VYV27UT3kw2ghklof5R22mx7/N5cWjwkK8ita
# 1Z1aQ9t1+yObuEDq9ODf/JaO5R0huMSFO8HMkUXihQrccHfwGr553XGcZqaDgo5m
# iKqtvzKsNoxIK/LGUhgrp5/kqwqWPdT08sX/HaYlUKTroccQgDAl/KKss8Y0b1tU
# 8aZsX8J+aroLBwFkNmxBlQHLa7u1qKQR/f7u+xOPsuLUXt/0PrXryXQ4y+jRSgx8
# 1TPcS9i2EUmQ4LpfAqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjQwODQy
# NTlaMC8GCSqGSIb3DQEJBDEiBCCNuxOv+16GvJw7n28LiDnWoF/59tVlBNV4B49N
# /JDrfzANBgkqhkiG9w0BAQEFAASCAgAvsuHz9m6FAEhEBqPsS/3QsZkINimrGwrM
# PCkHM2dR3CtZYGJIvaDPwd+/dmw5M+I75gV8z+qa/INQpvylr7wy5ErLW039ZPI4
# 4FX+T/UbDwT0EUVDOGR9LMQYxdWab5+wihUC8FkvVT7BXscDrrBeoLNGVvFqAUuR
# o7FDXL9lrgLqBldkQStqQQSSYGdwjhIzRSATkF0FX4NTrXZBjAdq91r3ZkUfDibZ
# 4nNWwaj8gLhJ2Lxx+OiQfkSOHZKqf4hFhX+0sAfu7C6t/5h7x5NOOnsZfgog4w//
# X9nn6QtqwI5xmPdF3HPJZlC+vAOJlfriy2qJFEIkUpm192aL+1I3v88IPow5Pqq9
# W014umVBd+YoO3bJSl3/MEB+SPuji2WB7p0cF6EfxlnQaZ66ZZPnx4yCnzu7rAXt
# CQPwXuQmTIhVMKBN9UHd2OmL+h0dXPLDJo0FozD0KwWgAKXkS/tvukKaHDD+1yt8
# OHSe37ytZht6K8+jn7Yj3U3kO73d+FgvARCIjJYn9h+bLFkn4FG7Rq8j3hkq343K
# 36/uJRVGLLxBTCfXUqvNW9t7BDnSccIW0VU+li5byqf4eioce0leA8yjg7VWqznr
# OZkTd23viCdrTb7bSi2x0gv1mZLzG4etUt/IIsc0J7MtG87/dj6CJKzjdjfS6M3K
# zagkGVlowg==
# SIG # End signature block
