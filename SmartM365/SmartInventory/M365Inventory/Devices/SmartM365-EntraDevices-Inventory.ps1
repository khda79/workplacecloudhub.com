<#
.SYNOPSIS
Generates a detailed inventory of Azure Entra (Azure AD) devices using Microsoft Graph API (Windows-only host).
Does NOT rely on Intune managed devices; it reads directly from Entra ID / Azure AD (/devices).

.DESCRIPTION
Connects to Microsoft Graph, retrieves devices from Azure Entra (Azure AD) using /devices,
optionally filtered by OperatingSystem (default: "Windows") and TrustType (disabled by default),
and exports a CSV with relevant device properties:
- Object ID, DeviceId, DisplayName, OS, OS version
- Compliance / management flags, ownership, MDM app
- Manufacturer, Model
- On-prem sync details
- Registration and last sign-in timestamps
- TrustType, ProfileType
- SystemLabels, PhysicalIds, Hostnames
- Parsed internal IDs from SystemLabels/PhysicalIds: GroupId, HardwareId, UserGroupId, UserHardwareId
- Parsed Autopilot ZTDID from labels/PhysicalIds

Generates additional focused reports:
- Global "Registered pending" devices (official criterion: TrustType ServerAd with no AlternativeSecurityIds)
- HardwareId conflicts (same HardwareId mapped to multiple DeviceId)
- HardwareId conflicts that contain true pending devices (IsPending)
- Remediation suggestions: removal candidates for stale/duplicate devices, plus a PS1 script
  containing Remove-MgDevice -WhatIf commands (for manual review and execution).

Uses SmartM365.Core for initialization, logging, CSV export, cleanup and global error notification.

.PARAMETER Tenant
Tenant profile key to load from Config/Tenants. Defaults to test.

.PARAMETER OutputPath
Specifies the output directory for the generated CSV and log files.

.PARAMETER Connect
Forces a (re)connection to Microsoft Graph (disconnects any existing session first).

.PARAMETER InteractiveAuth
Uses interactive authentication instead of app-only certificate authentication.

.PARAMETER OperatingSystemFilter
Filters devices by OperatingSystem (exact match). Default: "Windows".
Use empty string "" to disable the OS filter.

.PARAMETER TrustTypeFilter
Filters devices by TrustType (exact match). Disabled by default.
Use "ServerAd" to target hybrid joined devices. Use empty string "" or "false" to disable the TrustType filter.
.VERSION
1.12

.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication; Microsoft.Graph.Identity.DirectoryManagement.
    Minimum Graph application permissions: Directory.Read.All; Device.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Version : 1.12
    Author: https://github.com/khda79/workplacecloudhub.com
Requires: SmartM365.Core module and Microsoft.Graph.Identity.DirectoryManagement
Minimum application permissions: Directory.Read.All, Device.Read.All
#>

param(
    [string]$Tenant = 'test',
    [string]$OutputPath,
    [switch]$Connect,
    [switch]$InteractiveAuth,
    [string]$OperatingSystemFilter = "Windows",
    [string]$TrustTypeFilter = "",
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
    Import-Module -Name $modulePath -MinimumVersion '1.0.24' -ErrorAction Stop
} catch {
    Write-Host "Failed to import SmartM365.Core module from '$modulePath' : $_" -ForegroundColor Red
    exit 1
}

try {
    Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop
} catch {
    Write-Host "Failed to import Microsoft.Graph.Identity.DirectoryManagement. Install the module before running this script." -ForegroundColor Red
    Write-Host $_ -ForegroundColor Yellow
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

    if ($Object -is [System.Collections.IDictionary]) {
        if ($Object.Contains($Name)) { return $Object[$Name] }

        foreach ($key in $Object.Keys) {
            if ([string]$key -ieq $Name) {
                return $Object[$key]
            }
        }

        if ($Object.Contains('AdditionalProperties') -and $Object['AdditionalProperties'] -is [System.Collections.IDictionary]) {
            $additionalProperties = $Object['AdditionalProperties']
            if ($additionalProperties.Contains($Name)) {
                return $additionalProperties[$Name]
            }
        }

        return $null
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

# Helper: flattens array properties into a string
function Join-OrNull {
    param(
        [object]$Value,
        [string]$Separator = ";"
    )
    if (-not $Value) { return $null }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        $items = @()
        foreach ($v in $Value) {
            if ($null -ne $v -and "$v".Trim().Length -gt 0) {
                $items += "$v"
            }
        }
        if ($items.Count -eq 0) { return $null }
        return ($items -join $Separator)
    }
    return "$Value"
}

# Helper: parse a label value from a semicolon-separated label string
# Expected patterns:
#   [GID]:g:<id>
#   [HWID]:h:<id>
#   [USER-GID]:<userObjectId>:<id>
#   [USER-HWID]:<userObjectId>:<id>
#   [ZTDID]:d:<id> (Autopilot)
function Get-LabelValue {
    param(
        [string]$Labels,
        [Parameter(Mandatory)][string]$TagName
    )

    if ([string]::IsNullOrWhiteSpace($Labels)) {
        return $null
    }

    $parts = $Labels -split ';'
    foreach ($p in $parts) {
        $t = $p.Trim()
        if (-not $t) { continue }

        if ($t -notmatch '^\[(?<tag>[^\]]+)\]:(?<rest>.+)$') { continue }

        $tag  = $Matches['tag']
        $rest = $Matches['rest']

        if ($tag -ne $TagName) { continue }

        $segments = $rest -split ':'
        if ($segments.Count -eq 0) { return $null }

        return $segments[-1]
    }

    return $null
}

# Helper: converts extensionAttributes (if present) to JSON string
function Get-ExtensionAttributesJson {
    param(
        [Parameter(Mandatory)][object]$Device
    )

    $ext = Get-SafeProperty -Object $Device -Name 'ExtensionAttributes'
    if (-not $ext) {
        $ext = Get-SafeProperty -Object $Device -Name 'extensionAttributes'
    }

    if (-not $ext) { return $null }

    try {
        return ($ext | ConvertTo-Json -Depth 5 -Compress)
    } catch {
        WriteLog -Message "Failed to serialize extensionAttributes for device $($Device.Id): $($_.Exception.Message)" "WARNING"
        return $null
    }
}

# Helper: parse DateTime or return $null (works with DateTime or string from Graph/CSV)
function ConvertTo-DateTimeOrNull {
    param(
        [object]$Value
    )

    if ($null -eq $Value) { return $null }

    if ($Value -is [datetime]) { return $Value }

    $str = "$Value"
    if ([string]::IsNullOrWhiteSpace($str)) { return $null }

    try {
        return [datetime]$str
    } catch {
        WriteLog -Message "Could not parse DateTime value '$str'. Returning null." "WARNING"
        return $null
    }
}

# Returns the fixed inventory schema (column names) in the exact order
function Get-EntraDeviceColumns {
    [string[]]@(
        "ObjectId",
        "DeviceId",
        "DisplayName",

        "OperatingSystem",
        "OperatingSystemVersion",

        "IsCompliant",
        "IsManaged",
        "ManagedBy",
        "DeviceOwnership",
        "DeviceCategory",
        "DeviceVersion",
        "MDMAppId",
        "EnrollmentProfileName",

        "Manufacturer",
        "Model",

        "AccountEnabled",
        "TrustType",
        "ProfileType",

        "OnPremisesSyncEnabled",
        "OnPremisesLastSyncDateTime",

        "RegistrationDateTime",
        "ComplianceExpirationDateTime",
        "ApproximateLastSignInDateTime",

        "SystemLabels",
        "PhysicalIds",
        "Hostnames",
        "ExtensionAttributesJson",

        "GroupId",
        "HardwareId",
        "UserGroupId",
        "UserHardwareId",
        "AutopilotZTDID",
        "HasAlternativeSecurityIds",
        "IsPending"
    )
}

function New-SmartM365AiHelpUrl {
    [CmdletBinding()]
    param(
        [string]$ScriptName,
        [string]$Phase,
        [string]$ErrorMessage
    )

    $query = "SmartM365 {0} {1} {2}" -f $ScriptName, $Phase, $ErrorMessage
    return "https://www.bing.com/search?q={0}" -f [uri]::EscapeDataString($query)
}

function Send-EntraDevicesTeamsInfo {
    [CmdletBinding()]
    param(
        [int]$TotalDevices,
        [int]$ExportedDevices,
        [int]$RegisteredPendingCount,
        [int]$HardwareConflictCount,
        [int]$RemovalCandidateCount,
        [string]$MainCsvPath
    )

    $resultSummary = "Entra devices inventory completed successfully. Total devices: {0}; exported devices: {1}; registered pending: {2}; hardware conflicts: {3}; removal candidates: {4}." -f $TotalDevices, $ExportedDevices, $RegisteredPendingCount, $HardwareConflictCount, $RemovalCandidateCount

    $facts = [ordered]@{
        'Script'             = $script:SmartM365ScriptName
        'Tenant'             = $Tenant
        'Organization'       = $OrgDomain
        'Computer'           = $env:COMPUTERNAME
        'Timestamp'          = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        'CSV path'           = $MainCsvPath
        'Result summary'     = $resultSummary
        'Total devices'      = $TotalDevices
        'Exported devices'   = $ExportedDevices
        'Registered pending' = $RegisteredPendingCount
        'Hardware conflicts' = $HardwareConflictCount
        'Removal candidates' = $RemovalCandidateCount
    }

    Send-SmartM365TeamsNotification `
        -Title 'Azure Entra devices inventory - SUCCESS' `
        -Message $resultSummary `
        -Level SUCCESS `
        -Channel Infos `
        -ResultSummary $resultSummary `
        -Facts $facts | Out-Null
}

function Send-EntraDevicesTeamsAlert {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$ErrorRecord,
        [string]$Phase = 'Unknown',
        [string]$CsvPath = ''
    )

    $innerMessages = New-Object System.Collections.Generic.List[string]
    $inner = $ErrorRecord.Exception.InnerException
    while ($inner) {
        $innerMessages.Add($inner.Message) | Out-Null
        $inner = $inner.InnerException
    }

    $message = $ErrorRecord.Exception.Message
    $helpUrl = New-SmartM365AiHelpUrl -ScriptName $script:SmartM365ScriptName -Phase $Phase -ErrorMessage $message
    $facts = [ordered]@{
        'Script'             = $script:SmartM365ScriptName
        'Tenant'             = $Tenant
        'Organization'       = $OrgDomain
        'Computer'           = $env:COMPUTERNAME
        'Timestamp'          = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        'Failed phase'       = $Phase
        'Exception message'  = $message
        'Inner exception'    = if ($innerMessages.Count -gt 0) { $innerMessages -join ' | ' } else { '' }
        'Log path'           = $global:LogTextFile
        'Transcript path'    = $global:logTranscriptFile
        'Output path'        = $OutputPath
        'CSV path'           = $CsvPath
        'AI help URL'        = $helpUrl
    }

    Send-SmartM365TeamsNotification `
        -Title 'Azure Entra devices inventory - ERROR' `
        -Message $message `
        -Level ERROR `
        -Channel Alerts `
        -HelpUrl $helpUrl `
        -Facts $facts | Out-Null
}

# ==========================================================
# Initialization via SmartM365.Core
# ==========================================================
$ScriptVersion = "1.12"
$script:SmartM365ScriptName = $MyInvocation.MyCommand.Name
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EntraDevicesCsvLogFolderPath' -DefaultValue $OutputPath
try {
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPathInit $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
    Start-Transcript -Path $global:logTranscriptFile -Append

    WriteLog -Message "Script Environment initialized at $InitializeOutputPath"
    $OutputPath = $InitializeOutputPath
    WriteLog -Message "Starting $TaskName..."
} catch {
    Write-Host "Initialization failed: $_" -ForegroundColor Red
    try {
        Send-EntraDevicesTeamsAlert -ErrorRecord $_ -Phase 'Initialization' -CsvPath ''
    } catch {
        Write-Host "Failed to send Teams alert for initialization failure: $_" -ForegroundColor Yellow
    }
    exit 1
}

# ==========================================================
# MAIN TRY / CATCH / FINALLY
# ==========================================================
$connectedGraphInThisRun = $false
$script:ExitCode = 0
$currentOperation = 'Starting main processing'
$totalDeviceCount = 0
$registeredPendingCount = 0
$hardwareConflictCount = 0
$removalCandidateCount = 0
$mainLatestCsvPath = ''
$results = @()

try {

    # ------------------------
    # Connect to Microsoft Graph via SmartM365.Core / Connect-SmartM365CloudSession
    # ------------------------
    $currentOperation = 'Disconnecting existing Microsoft Graph session'
    Write-Host "Existing Graph session (if any) will be disconnected before connection..." -ForegroundColor Cyan
    Disconnect-SmartM365CloudSession -ExchangeOnline $false -Graph $true -VerboseDisconnect:$true

    $currentOperation = 'Connecting to Microsoft Graph'
    $connectParams = @{
        ExchangeOnline = $false
        Graph          = $true
        GraphScopes    = @("Directory.Read.All")
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

    # ------------------------
    $currentOperation = 'Running preflight checks'
    Invoke-SmartM365Preflight `
        -ScriptName $TaskName `
        -RequiredModules @('Microsoft.Graph.Authentication','Microsoft.Graph.Identity.DirectoryManagement') `
        -RequiredCommands @('Get-MgDevice') `
        -OutputPaths @($OutputPath) `
        -RequiredGraphApplicationPermissions @('Directory.Read.All','Device.Read.All') -GraphProbeUris @(
        'https://graph.microsoft.com/v1.0/devices?$top=1',
        'https://graph.microsoft.com/v1.0/organization'
    ) | Out-Null

    # Retrieve Azure Entra (Azure AD) devices
    # ------------------------
    $currentOperation = 'Retrieving Azure Entra devices'
    WriteLog -Message "Retrieving Azure Entra (Azure AD) devices..."

    $deviceProperties = @(
        "id",
        "deviceId",
        "displayName",
        "operatingSystem",
        "operatingSystemVersion",
        "isCompliant",
        "isManaged",
        "managedBy",
        "deviceOwnership",
        "deviceCategory",
        "deviceVersion",
        "mdmAppId",
        "enrollmentProfileName",
        "manufacturer",
        "model",
        "accountEnabled",
        "trustType",
        "profileType",
        "onPremisesSyncEnabled",
        "onPremisesLastSyncDateTime",
        "registrationDateTime",
        "complianceExpirationDateTime",
        "approximateLastSignInDateTime",
        "systemLabels",
        "physicalIds",
        "hostnames",
        "alternativeSecurityIds",
        "extensionAttributes"
    )

    if ($MaxItems -gt 0) {
        WriteLog -Message ("MaxItems enabled: retrieving at most {0} Entra devices from Graph." -f $MaxItems) "WARNING"
        $devices = @(Get-MgDevice -Top $MaxItems -Property ($deviceProperties -join ','))
    }
    else {
        $devices = @(Get-MgDevice -All -Property ($deviceProperties -join ','))
    }
    $totalDeviceCount = $devices.Count

    WriteLog -Message "Azure Entra devices retrieved: $($devices.Count)"
    $serverAdDeviceCount = @($devices | Where-Object { (Get-SafeProperty $_ 'TrustType') -eq 'ServerAd' }).Count
    $alternativeSecurityIdDeviceCount = @($devices | Where-Object {
        $alternativeSecurityIds = @(Get-SafeProperty $_ 'AlternativeSecurityIds')
        @($alternativeSecurityIds | Where-Object { $null -ne $_ }).Count -gt 0
    }).Count
    WriteLog -Message ("AlternativeSecurityIds quality gate: ServerAd={0}; WithAlternativeSecurityIds={1}" -f $serverAdDeviceCount, $alternativeSecurityIdDeviceCount)
    if ($serverAdDeviceCount -gt 0 -and $alternativeSecurityIdDeviceCount -eq 0) {
        throw "Microsoft Graph returned ServerAd devices but no AlternativeSecurityIds. Registered-pending exports would be invalid, so publication is stopped."
    }

    if (-not $devices -or $devices.Count -eq 0) {
        WriteLog -Message "No Azure Entra devices found. Exiting." "WARNING"
        Send-EntraDevicesTeamsInfo -TotalDevices $totalDeviceCount -ExportedDevices 0 -RegisteredPendingCount 0 -HardwareConflictCount 0 -RemovalCandidateCount 0 -MainCsvPath ''
        return
    }

    # ------------------------
    # Default OS filter
    # ------------------------
    if ($OperatingSystemFilter -and $OperatingSystemFilter.Trim().Length -gt 0) {
        WriteLog -Message "Applying OperatingSystem filter: $OperatingSystemFilter"
        $devices = @($devices | Where-Object {
            (Get-SafeProperty $_ 'OperatingSystem') -eq $OperatingSystemFilter
        })
        WriteLog -Message "Devices after OperatingSystem filter: $($devices.Count)"
    }

    if (-not $devices -or $devices.Count -eq 0) {
        WriteLog -Message "No Azure Entra devices match the current OS filter. Exiting." "WARNING"
        Send-EntraDevicesTeamsInfo -TotalDevices $totalDeviceCount -ExportedDevices 0 -RegisteredPendingCount 0 -HardwareConflictCount 0 -RemovalCandidateCount 0 -MainCsvPath ''
        return
    }

    # ------------------------
    # TrustType filter
    # ------------------------
    if ($TrustTypeFilter -and $TrustTypeFilter.Trim().Length -gt 0 -and $TrustTypeFilter.Trim() -notin @('false', '$false', '0', 'no', 'off')) {
        WriteLog -Message "Applying TrustType filter: $TrustTypeFilter"
        $devices = @($devices | Where-Object {
            (Get-SafeProperty $_ 'TrustType') -eq $TrustTypeFilter
        })
        WriteLog -Message "Devices after TrustType filter: $($devices.Count)"

        if (-not $devices -or $devices.Count -eq 0) {
            WriteLog -Message "No Azure Entra devices with TrustType = '$TrustTypeFilter' after filters. Exiting." "WARNING"
            Send-EntraDevicesTeamsInfo -TotalDevices $totalDeviceCount -ExportedDevices 0 -RegisteredPendingCount 0 -HardwareConflictCount 0 -RemovalCandidateCount 0 -MainCsvPath ''
            return
        }
    }
    else {
        WriteLog -Message "TrustType filter disabled."
    }

    # ------------------------
    # Build result objects (Entra-only view, with parsed labels)
    # ------------------------
    $currentOperation = 'Building inventory results'
    $results = $devices | ForEach-Object {
        $objId        = Get-SafeProperty $_ 'Id'
        $systemLabels = Join-OrNull (Get-SafeProperty $_ 'SystemLabels')
        $physicalIds  = Join-OrNull (Get-SafeProperty $_ 'PhysicalIds')
        $hostnames    = Join-OrNull (Get-SafeProperty $_ 'Hostnames')

        # Combine SystemLabels + PhysicalIds for GID/HWID/USER-GID/USER-HWID
        $labelsForIds = $null
        if ($systemLabels -and $physicalIds) {
            $labelsForIds = "$systemLabels;$physicalIds"
        } elseif ($systemLabels) {
            $labelsForIds = $systemLabels
        } elseif ($physicalIds) {
            $labelsForIds = $physicalIds
        }

        $groupId        = Get-LabelValue -Labels $labelsForIds -TagName "GID"
        $hardwareId     = Get-LabelValue -Labels $labelsForIds -TagName "HWID"
        $userGroupId    = Get-LabelValue -Labels $labelsForIds -TagName "USER-GID"
        $userHardwareId = Get-LabelValue -Labels $labelsForIds -TagName "USER-HWID"

        # Autopilot ZTDID: priority PhysicalIds, then SystemLabels
        $ztdid = $null
        if ($physicalIds) {
            $ztdid = Get-LabelValue -Labels $physicalIds -TagName "ZTDID"
        }
        if (-not $ztdid -and $systemLabels) {
            $ztdid = Get-LabelValue -Labels $systemLabels -TagName "ZTDID"
        }

        $altSecIds = Get-SafeProperty $_ 'AlternativeSecurityIds'
        $hasAltSecIds = $false
        if ($altSecIds) {
            if ($altSecIds -is [System.Collections.IEnumerable] -and -not ($altSecIds -is [string])) {
                $hasAltSecIds = (@($altSecIds) | Where-Object { $null -ne $_ }).Count -gt 0
            }
            else {
                $hasAltSecIds = -not [string]::IsNullOrWhiteSpace([string]$altSecIds)
            }
        }
        $trustTypeVal = Get-SafeProperty $_ 'TrustType'
        $isPending = ($trustTypeVal -eq 'ServerAd') -and (-not $hasAltSecIds)
        [PSCustomObject]@{
            "ObjectId"                        = $objId
            "DeviceId"                        = Get-SafeProperty $_ 'DeviceId'
            "DisplayName"                     = Get-SafeProperty $_ 'DisplayName'

            "OperatingSystem"                 = Get-SafeProperty $_ 'OperatingSystem'
            "OperatingSystemVersion"          = Get-SafeProperty $_ 'OperatingSystemVersion'

            "IsCompliant"                     = Get-SafeProperty $_ 'IsCompliant'
            "IsManaged"                       = Get-SafeProperty $_ 'IsManaged'
            "ManagedBy"                       = Get-SafeProperty $_ 'ManagedBy'
            "DeviceOwnership"                 = Get-SafeProperty $_ 'DeviceOwnership'
            "DeviceCategory"                  = Get-SafeProperty $_ 'DeviceCategory'
            "DeviceVersion"                   = Get-SafeProperty $_ 'DeviceVersion'
            "MDMAppId"                        = Get-SafeProperty $_ 'MdmAppId'
            "EnrollmentProfileName"           = Get-SafeProperty $_ 'EnrollmentProfileName'

            "Manufacturer"                    = Get-SafeProperty $_ 'Manufacturer'
            "Model"                           = Get-SafeProperty $_ 'Model'

            "AccountEnabled"                  = Get-SafeProperty $_ 'AccountEnabled'
            "TrustType"                       = Get-SafeProperty $_ 'TrustType'
            "ProfileType"                     = Get-SafeProperty $_ 'ProfileType'

            "OnPremisesSyncEnabled"           = Get-SafeProperty $_ 'OnPremisesSyncEnabled'
            "OnPremisesLastSyncDateTime"      = Get-SafeProperty $_ 'OnPremisesLastSyncDateTime'

            "RegistrationDateTime"            = Get-SafeProperty $_ 'RegistrationDateTime'
            "ComplianceExpirationDateTime"    = Get-SafeProperty $_ 'ComplianceExpirationDateTime'
            "ApproximateLastSignInDateTime"   = Get-SafeProperty $_ 'ApproximateLastSignInDateTime'

            "SystemLabels"                    = $systemLabels
            "PhysicalIds"                     = $physicalIds
            "Hostnames"                       = $hostnames

            "ExtensionAttributesJson"         = Get-ExtensionAttributesJson -Device $_

            "GroupId"                         = $groupId
            "HardwareId"                      = $hardwareId
            "UserGroupId"                     = $userGroupId
            "UserHardwareId"                  = $userHardwareId
            "AutopilotZTDID"                  = $ztdid
            "HasAlternativeSecurityIds"       = $hasAltSecIds
            "IsPending"                       = $isPending
        }
    }

    # ------------------------
    # Ensure column order
    # ------------------------
    $columnOrder = Get-EntraDeviceColumns

    $projectedResults = foreach ($r in $results) {
        $row = [ordered]@{}
        foreach ($col in $columnOrder) {
            if ($r.PSObject.Properties[$col]) {
                $row[$col] = $r.$col
            } else {
                $row[$col] = $null
            }
        }
        [PSCustomObject]$row
    }

    $results = $projectedResults

    # ------------------------
    # Export CSV - main Entra inventory
    # ------------------------
    $currentOperation = 'Exporting main Entra devices CSV'
    Write-Host "`n--- Export CSV (main Entra devices inventory) ---"
    $BaseFileName = "M365_Entra_Devices"

    ExportAndCopyCsv -BaseFileName $BaseFileName `
       -OutputPath $OutputPath `
       -GlobalPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') `
       -Data $results `
       -Encoding "UTF8" `
       -NoTypeInformation

    WriteLog -Message "Export completed: $global:csvFilePath1"
    WriteLog -Message "Number of exported Entra devices: $($results.Count)"
    $mainLatestCsvPath = $global:csvFilePath3

    # ------------------------
    # Global report: Registered pending using the official Microsoft criterion.
    # IsPending = TrustType ServerAd and no AlternativeSecurityIds.
    # ------------------------
    WriteLog -Message "Building global Registered pending report (IsPending = true, official Microsoft criterion)..."

    $registeredPending = $results | Where-Object { $_.IsPending -eq $true }
    $legacyPending = $results | Where-Object { -not $_.ApproximateLastSignInDateTime }
    $registeredPendingCount = @($registeredPending).Count
    WriteLog -Message ("Registered pending official IsPending: {0}; legacy no LastSignIn heuristic: {1}." -f $registeredPendingCount, @($legacyPending).Count)

    if ($registeredPending -and $registeredPending.Count -gt 0) {
        WriteLog -Message ("Registered pending devices found: {0}" -f $registeredPending.Count)

        $BaseFileNamePending = "M365_Entra_Devices_RegisteredPending"

        ExportAndCopyCsv -BaseFileName $BaseFileNamePending `
            -OutputPath $OutputPath `
            -GlobalPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') `
            -Data $registeredPending `
            -Encoding "UTF8" `
            -NoTypeInformation

        WriteLog -Message "Global Registered pending report exported."
    }
    else {
        WriteLog -Message "No Registered pending devices found (all ServerAd devices have AlternativeSecurityIds)." "INFO"
    }

    # ------------------------
    # Targeted report: same HardwareId with multiple DeviceId
    # ------------------------
    WriteLog -Message "Building HardwareId conflict report (same HardwareId, multiple DeviceId)..."

    $hwConflicts = $results |
        Where-Object { $_.HardwareId -and $_.HardwareId.Trim().Length -gt 0 } |
        Group-Object -Property HardwareId |
        Where-Object {
            ($_.Group |
                Select-Object -ExpandProperty DeviceId -Unique |
                Where-Object { $_ -and $_.ToString().Trim().Length -gt 0 }
            ).Count -gt 1
        } |
        ForEach-Object {
            $hw    = $_.Name
            $group = $_.Group

            $deviceIds   = ($group | Select-Object -ExpandProperty DeviceId -Unique       | Where-Object { $_ }) -join ";"
            $objectIds   = ($group | Select-Object -ExpandProperty ObjectId -Unique       | Where-Object { $_ }) -join ";"
            $names       = ($group | Select-Object -ExpandProperty DisplayName -Unique    | Where-Object { $_ }) -join ";"
            $oses        = ($group | Select-Object -ExpandProperty OperatingSystem -Unique| Where-Object { $_ }) -join ";"
            $osVersions  = ($group | Select-Object -ExpandProperty OperatingSystemVersion -Unique | Where-Object { $_ }) -join ";"
            $trustTypes  = ($group | Select-Object -ExpandProperty TrustType -Unique      | Where-Object { $_ }) -join ";"
            $regDates    = ($group | Select-Object -ExpandProperty RegistrationDateTime   | Sort-Object | ForEach-Object { $_ }) -join ";"
            $lastSignIns = ($group | Select-Object -ExpandProperty ApproximateLastSignInDateTime | Sort-Object | ForEach-Object { $_ }) -join ";"
            $autopilotIds= ($group | Select-Object -ExpandProperty AutopilotZTDID -Unique | Where-Object { $_ }) -join ";"

            [PSCustomObject]@{
                "HardwareId"                     = $hw
                "DeviceCount"                    = $group.Count
                "DeviceIds"                      = $deviceIds
                "ObjectIds"                      = $objectIds
                "DisplayNames"                   = $names
                "OperatingSystems"               = $oses
                "OperatingSystemVersions"        = $osVersions
                "TrustTypes"                     = $trustTypes
                "RegistrationDateTimes"          = $regDates
                "ApproximateLastSignInDateTimes" = $lastSignIns
                "AutopilotZTDIDs"                = $autopilotIds
            }
        }

    if ($hwConflicts -and $hwConflicts.Count -gt 0) {
        $hardwareConflictCount = @($hwConflicts).Count
        WriteLog -Message "HardwareId conflict entries found: $($hwConflicts.Count)"

        $BaseFileNameHwConflicts = "M365_Entra_Devices_HardwareIdConflicts"

        ExportAndCopyCsv -BaseFileName $BaseFileNameHwConflicts `
            -OutputPath $OutputPath `
            -GlobalPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') `
            -Data $hwConflicts `
            -Encoding "UTF8" `
            -NoTypeInformation

        WriteLog -Message "HardwareId conflict report exported (HardwareId -> multiple DeviceId)."

        # ------------------------
        # Sub-report: HardwareId conflicts where some devices are truly pending
        # ------------------------
        WriteLog -Message "Building sub-report for HardwareId conflicts with pending devices (IsPending = true)..."

        $devicesByHardwareId = @{}
        foreach ($device in $results) {
            if (-not $device.HardwareId -or $device.HardwareId.Trim().Length -eq 0) { continue }
            if (-not $devicesByHardwareId.ContainsKey($device.HardwareId)) {
                $devicesByHardwareId[$device.HardwareId] = New-Object System.Collections.Generic.List[object]
            }
            $devicesByHardwareId[$device.HardwareId].Add($device) | Out-Null
        }

        $hwPending = New-Object System.Collections.Generic.List[object]

        foreach ($hwGroup in $hwConflicts) {
            $currentHw = $hwGroup.HardwareId
            if (-not $devicesByHardwareId.ContainsKey($currentHw)) { continue }

            foreach ($d in $devicesByHardwareId[$currentHw]) {
                if ($d.IsPending -ne $true) { continue }

                $hwPending.Add([PSCustomObject]@{
                    "HardwareId"                     = $currentHw
                    "ObjectId"                       = $d.ObjectId
                    "DeviceId"                       = $d.DeviceId
                    "DisplayName"                    = $d.DisplayName
                    "OperatingSystem"                = $d.OperatingSystem
                    "OperatingSystemVersion"         = $d.OperatingSystemVersion
                    "TrustType"                      = $d.TrustType
                    "OnPremisesSyncEnabled"          = $d.OnPremisesSyncEnabled
                    "OnPremisesLastSyncDateTime"     = $d.OnPremisesLastSyncDateTime
                    "RegistrationDateTime"           = $d.RegistrationDateTime
                    "ApproximateLastSignInDateTime"  = $d.ApproximateLastSignInDateTime
                    "AutopilotZTDID"                 = $d.AutopilotZTDID
                    "HasAlternativeSecurityIds"      = $d.HasAlternativeSecurityIds
                    "IsPending"                      = $d.IsPending
                }) | Out-Null
            }
        }

        $BaseFileNameHwPending = "M365_Entra_Devices_HardwareIdConflicts_RegisteredPending"
        $hwPendingColumns = @(
            "HardwareId",
            "ObjectId",
            "DeviceId",
            "DisplayName",
            "OperatingSystem",
            "OperatingSystemVersion",
            "TrustType",
            "OnPremisesSyncEnabled",
            "OnPremisesLastSyncDateTime",
            "RegistrationDateTime",
            "ApproximateLastSignInDateTime",
            "AutopilotZTDID",
            "HasAlternativeSecurityIds",
            "IsPending"
        )

        if ($hwPending.Count -gt 0) {
            WriteLog -Message "Pending devices inside HardwareId conflicts: $($hwPending.Count)"

            ExportAndCopyCsv -BaseFileName $BaseFileNameHwPending `
                -OutputPath $OutputPath `
                -GlobalPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') `
                -Data $hwPending `
                -Encoding "UTF8" `
                -NoTypeInformation

            WriteLog -Message "Pending devices in HardwareId conflict report exported."
        }
        else {
            $latestCsvFolder = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''
            $emptyPendingExport = Export-SmartM365Csv `
                -BaseFileName $BaseFileNameHwPending `
                -OutputPath $OutputPath `
                -GlobalPath $latestCsvFolder `
                -Data $null `
                -Columns $hwPendingColumns `
                -Encoding "UTF8"

            $currentBaseFileName = Add-SmartM365MaxItemsSuffixToBaseName -BaseFileName $BaseFileNameHwPending
            $currentPendingPath = Join-Path $OutputPath ("{0}.csv" -f $currentBaseFileName)
            Copy-Item -LiteralPath $emptyPendingExport.TimestampedPath -Destination $currentPendingPath -Force -ErrorAction Stop
            if (-not $global:csvGeneratedPaths) {
                $global:csvGeneratedPaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
            }
            [void]$global:csvGeneratedPaths.Add($currentPendingPath)
            WriteLog -Message "No pending devices found inside HardwareId conflicts. Header-only CSV published to replace stale snapshots." "INFO"
        }
    }
    else {
        WriteLog -Message "No HardwareId conflicts detected (no HardwareId mapped to more than one DeviceId)." "INFO"
    }

    # ======================================================
    # Remediation suggestions based on HardwareId duplicates
    # ======================================================
    WriteLog -Message "Building remediation suggestions (HardwareId duplicates)..."

    $allDevices = $results

    $devicesWithHwId = $allDevices | Where-Object {
        $_.HardwareId -and $_.HardwareId.Trim().Length -gt 0
    }

    WriteLog -Message ("Devices with non-empty HardwareId: {0}" -f $devicesWithHwId.Count)

    $hwGroupsForRem = $devicesWithHwId | Group-Object -Property HardwareId
    WriteLog -Message ("HardwareId groups (non-empty): {0}" -f $hwGroupsForRem.Count)

    $remediationRows = New-Object System.Collections.Generic.List[object]

    foreach ($group in $hwGroupsForRem) {

        $hardwareId = $group.Name
        $devices    = $group.Group

        if ($devices.Count -lt 2) { continue }

        $devicesEnriched = $devices | ForEach-Object {
            $lastSignIn  = ConvertTo-DateTimeOrNull $_.ApproximateLastSignInDateTime
            $regDate     = ConvertTo-DateTimeOrNull $_.RegistrationDateTime

            [PSCustomObject]@{
                Raw              = $_
                LastSignIn       = $lastSignIn
                RegistrationDate = $regDate
            }
        }

        $devicesSorted = $devicesEnriched | Sort-Object `
            @{ Expression = { if ($_.LastSignIn) { 0 } else { 1 } } }, `
            @{ Expression = { $_.LastSignIn }; Descending = $true }, `
            @{ Expression = { $_.RegistrationDate }; Descending = $true }

        $primary = $devicesSorted | Select-Object -First 1
        if (-not $primary) { continue }

        $primaryRaw     = $primary.Raw
        $primaryRegDate = $primary.RegistrationDate

        foreach ($dev in $devicesSorted) {
            $raw = $dev.Raw
            $ls  = $dev.LastSignIn
            $reg = $dev.RegistrationDate

            if ($raw.ObjectId -eq $primaryRaw.ObjectId) { continue }

            $reason = $null

            if (-not $ls) {
                $reason = "NeverSignedIn"
            }
            elseif ($primaryRegDate -and $reg -and ($reg -lt $primaryRegDate)) {
                $reason = "OlderRegistrationThanPrimary"
            }

            if ($reason) {
                $remediationRows.Add(
                    [PSCustomObject]@{
                        HardwareId                      = $hardwareId

                        CandidateObjectId               = $raw.ObjectId
                        CandidateDeviceId               = $raw.DeviceId
                        CandidateDisplayName            = $raw.DisplayName
                        CandidateRegistrationDateTime   = $raw.RegistrationDateTime
                        CandidateLastSignInDateTime     = $raw.ApproximateLastSignInDateTime

                        PrimaryObjectId                 = $primaryRaw.ObjectId
                        PrimaryDeviceId                 = $primaryRaw.DeviceId
                        PrimaryDisplayName              = $primaryRaw.DisplayName
                        PrimaryRegistrationDateTime     = $primaryRaw.RegistrationDateTime
                        PrimaryLastSignInDateTime       = $primaryRaw.ApproximateLastSignInDateTime

                        Reason                           = $reason
                    }
                ) | Out-Null
            }
        }
    }

    if ($remediationRows.Count -eq 0) {
        WriteLog -Message "No removal candidates identified based on HardwareId duplicates." "INFO"
    } else {
        $removalCandidateCount = $remediationRows.Count
        WriteLog -Message ("Removal candidates identified: {0}" -f $remediationRows.Count) "INFO"

        $BaseFileNameRem = "M365_Entra_Devices_RemovalCandidates"

        ExportAndCopyCsv -BaseFileName $BaseFileNameRem `
            -OutputPath $OutputPath `
            -GlobalPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') `
            -Data $remediationRows `
            -Encoding "UTF8" `
            -NoTypeInformation

        WriteLog -Message "Remediation candidates CSV exported."

        $timestamp             = Get-Date -Format "yyyyMMdd_HHmmss"
        $commandScriptFileName = "SmartM365-EntraDevices-RemoveCandidates_$timestamp.ps1"
        $commandScriptPath     = Join-Path $OutputPath $commandScriptFileName

        $lines = @()
        $lines += '# This script was generated automatically by the Entra inventory script.'
        $lines += '# It contains Remove-MgDevice commands for candidate devices.'
        $lines += '# REVIEW CAREFULLY before execution.'
        $lines += ''
        $lines += 'param('
        $lines += '    [switch]$WhatIf = $true,'
        $lines += '    [switch]$InteractiveAuth = $true,'
        $lines += '    [string[]]$Scopes = @("Directory.ReadWrite.All","Device.ReadWrite.All","DeviceManagementManagedDevices.Read.All"),'
        $lines += '    [switch]$SkipIfManagedInIntune = $true,'
        $lines += '    [int]$IntuneLookbackDays = 90,'
        $lines += '    [switch]$AllowDeleteIfIntuneCheckFails = $false'
        $lines += ')'
        $lines += ''
        $lines += '# Graph connection (interactive)'
        $lines += 'try {'
        $lines += '    $ctx = Get-MgContext'
        $lines += '    if (-not $ctx -or -not $ctx.Account) {'
        $lines += '        if (-not $InteractiveAuth) { throw "Not connected to Microsoft Graph. Re-run with -InteractiveAuth or connect manually." }'
        $lines += '        try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch { }'
        $lines += '        Connect-MgGraph -Scopes $Scopes -NoWelcome'
        $lines += '    }'
        $lines += '} catch {'
        $lines += '    throw ("Graph connection failed: {0}" -f $_.Exception.Message)'
        $lines += '}'
        $lines += ''
        $lines += 'function Test-ManagedInIntune {'
        $lines += '    param('
        $lines += '        [Parameter(Mandatory=$true)][string]$EntraObjectId,'
        $lines += '        [int]$LookbackDays = 180'
        $lines += '    )'
        $lines += '    $result = [ordered]@{ CheckSucceeded = $true; Found = $false; AzureAdDeviceId = $null; LastSyncDateTime = $null; IsRecentSync = $false; Reason = $null }'
        $lines += '    try {'
        $lines += '        $entra = Get-MgDevice -DeviceId $EntraObjectId -Property "id,deviceId,displayName" -ErrorAction Stop'
        $lines += '        $result.AzureAdDeviceId = $entra.DeviceId'
        $lines += '        if (-not $result.AzureAdDeviceId) {'
        $lines += '            $result.Reason = "Entra device has empty deviceId."'
        $lines += '            return [pscustomobject]$result'
        $lines += '        }'
        $lines += '        $filter = "azureADDeviceId eq ''{0}''" -f $result.AzureAdDeviceId'
        $lines += '        $md = Get-MgDeviceManagementManagedDevice -Filter $filter -Top 1 -ErrorAction Stop'
        $lines += '        if ($md) {'
        $lines += '            $result.Found = $true'
        $lines += '            $result.LastSyncDateTime = $md.LastSyncDateTime'
        $lines += '            if ($md.LastSyncDateTime) {'
        $lines += '                $cutoff = (Get-Date).AddDays(-1 * [math]::Abs($LookbackDays))'
        $lines += '                if ([datetime]$md.LastSyncDateTime -lt $cutoff) {'
        $lines += '                    $result.IsRecentSync = $false'
        $lines += '                    $result.Reason = "Found in Intune, but LastSyncDateTime older than lookback window."'
        $lines += '                } else {'
        $lines += '                    $result.IsRecentSync = $true'
        $lines += '                    $result.Reason = "Found in Intune and synced within lookback window."'
        $lines += '                }'
        $lines += '            } else {'
        $lines += '                $result.IsRecentSync = $true'
        $lines += '                $result.Reason = "Found in Intune (LastSyncDateTime unavailable). Treating as recent for safety."'
        $lines += '            }'
        $lines += '        }'
        $lines += '    } catch {'
        $lines += '        $result.CheckSucceeded = $false'
        $lines += '        $result.Reason = $_.Exception.Message'
        $lines += '    }'
        $lines += '    return [pscustomobject]$result'
        $lines += '}'
        $lines += ''

        foreach ($row in $remediationRows) {
            $candidateObjectId = $null
            try {
                if ($row -and $row.PSObject -and $row.PSObject.Properties) {
                    foreach ($name in @("CandidateObjectId","CandidateDeviceObjectId","CandidateDeviceId","Id")) {
                        if ($row.PSObject.Properties.Match($name).Count -gt 0) {
                            $v = $row.$name
                            if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v)) { $candidateObjectId = [string]$v; break }
                        }
                    }
                }
            } catch { $candidateObjectId = $null }

            $comment = "# HardwareId=$($row.HardwareId); PrimaryDeviceId=$($row.PrimaryDeviceId); Reason=$($row.Reason); CandidateDeviceId=$($row.CandidateDeviceId); CandidateObjectId=$candidateObjectId"
            $lines  += $comment

            if (-not $candidateObjectId) {
                $lines += 'Write-Warning "SKIP DELETE: CandidateObjectId is empty (cannot evaluate Intune / delete safely)."'
                $lines += ''
                continue
            }

            $lines  += ('$CandidateObjectId = ''{0}''' -f $candidateObjectId)
            $lines  += 'if ($WhatIf) {'
            $lines  += '    Write-Host ("WHATIF: Remove-MgDevice -DeviceId ''{0}''" -f $CandidateObjectId)'
            $lines  += '} else {'
            $lines  += '    $skipDelete = $false'
            $lines  += '    if ($SkipIfManagedInIntune) {'
            $lines  += '        $check = Test-ManagedInIntune -EntraObjectId $CandidateObjectId -LookbackDays $IntuneLookbackDays'
            $lines  += '        if (-not $check.CheckSucceeded) {'
            $lines  += '            if (-not $AllowDeleteIfIntuneCheckFails) {'
            $lines  += '                Write-Warning ("SKIP DELETE: Intune enrollment check failed (CandidateObjectId={0}). {1}" -f $CandidateObjectId, $check.Reason)'
            $lines  += '                $skipDelete = $true'
            $lines  += '            }'
            $lines  += '        } elseif ($check.Found -and $check.IsRecentSync) {'
            $lines  += '            Write-Warning ("SKIP DELETE: Device appears enrolled in Intune (AzureAdDeviceId={0}). {1}" -f $check.AzureAdDeviceId, $check.Reason)'
            $lines  += '            $skipDelete = $true'
            $lines  += '        } elseif ($check.Found -and -not $check.IsRecentSync) {'
            $lines  += '            Write-Warning ("Intune record found but appears stale (AzureAdDeviceId={0}). {1} Proceeding with Entra delete." -f $check.AzureAdDeviceId, $check.Reason)'
            $lines  += '        }'
            $lines  += '    }'
            $lines  += '    if (-not $skipDelete) {'
            $lines  += '        Remove-MgDevice -DeviceId $CandidateObjectId'
            $lines  += '    }'
            $lines  += '}'
            $lines  += ''
        }

        $lines += 'try { Disconnect-MgGraph | Out-Null } catch { }'
        $lines += ''

        Set-Content -Path $commandScriptPath -Value $lines -Encoding UTF8
        WriteLog -Message ("Remove-MgDevice command script generated: {0}" -f $commandScriptPath)
    }

    Send-EntraDevicesTeamsInfo `
        -TotalDevices $totalDeviceCount `
        -ExportedDevices (@($results).Count) `
        -RegisteredPendingCount $registeredPendingCount `
        -HardwareConflictCount $hardwareConflictCount `
        -RemovalCandidateCount $removalCandidateCount `
        -MainCsvPath $mainLatestCsvPath

}
catch {
    $globalError = $_
    $script:ExitCode = 1
    WriteLog -Message ("Global error in Azure Entra devices inventory: {0}" -f $globalError) "ERROR"
    Write-Host "A global error occurred. Check the log file for details." -ForegroundColor Red

    try {
        Send-EntraDevicesTeamsAlert -ErrorRecord $globalError -Phase $currentOperation -CsvPath $mainLatestCsvPath
    }
    catch {
        WriteLog -Message ("Failed to send Teams alert notification: {0}" -f $_) "ERROR"
    }

    try {
        $title = "Azure Entra devices inventory - ERROR"
        $msg   = @"
An error occurred in script $($MyInvocation.MyCommand.Name) on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss").

Error message:
$($globalError.Exception.Message)

See attached log file for details:
$($global:LogTextFile)
"@

        $bodyHtml = NewSimpleEmailBody -Title $title -Message $msg

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
    if ($connectedGraphInThisRun) {
        Write-Host "`n--- Disconnect Cloud Services ---"
        try {
            Disconnect-SmartM365CloudSession -ExchangeOnline:$false -Graph:$true -VerboseDisconnect:$true
        } catch {
            WriteLog -Message ("Error during Graph disconnect in finally: {0}" -f $_) "WARNING"
        }
    }

    try {
        RemoveOldFiles -Path $OutputPath -Filter "*.csv" -KeepCount $global:RetentionMaxCSV -LogFile $global:LogTextFile
        RemoveOldFiles -Path $global:LogPath -Filter "*.log" -KeepCount $global:RetentionMaxLogs -LogFile $global:LogTextFile
    } catch {
        WriteLog -Message ("Error during cleanup in finally: {0}" -f $_) "WARNING"
    }

    WriteLog -Message "$TaskName completed (finally block)."

    try {
        Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
        Complete-SmartM365ExecutionContext -Status Auto
    } catch {
    }

    if ($script:ExitCode -ne 0) {
        exit $script:ExitCode
    }
}
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCIYXZeMv/PEyEV
# MeLJCRgKjEuKPKE/rKsTlpsW6tTjP6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGKthuMECJTukX6g2gEum+ij6wXLrohw8QCgqVqlE8rBMA0GCSqG
# SIb3DQEBAQUABIIBgABE8VfO+evzpOHYTlrEwTpDH5SfK8D16dwww4T0c2TfIg4V
# 7/iXuprfIlxNYDi2kaBOI2qg7y8BmU92varg73FyE1bKp9WaDm3HmY+7jxeD1btH
# 3S/OuFkp5xQDwC936SZTrgL8VyczDg4ZgHROrUYR8Wyka9xCLbqH8I3DUWRjcWZG
# NWgjynJnuimSSbvgfppDI4D4bHvvfRjh8EkfYFQVmzOBh/SAoFX0G8IZtPyRO6X8
# nY4wgkm94PLEKpXvr7EcqncE7q6i/gyjZkWpmzFtrL+UnkZG7Q4lsErnkIcav+Xd
# ecvwpSZdWjBAdSarU37baafAZJTOvRUENvj8qPw48gIgsOjmPwRZmrEcC3lI5jQX
# nuE6WegRed0tcG1JhD+oms/0WQR+/VNigMeydPWxo3BAUv+cKjTgDQ5zkJDljEj0
# T/wgLPa3kexxaZ8AF6AQegLD2W40TzEv2Aiue4JSGy9yFElFlFkLo0vINWx87ouw
# a4lbgUV3YPKBR7WRo6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTYxODA5
# MzdaMC8GCSqGSIb3DQEJBDEiBCCn4ikZg2f8NPPrS8HU5sYzivWpGiclaz2lFWFh
# qr1mDDANBgkqhkiG9w0BAQEFAASCAgBS6r8U7ueX42/x7vC5tZuToJ7luhOr6sTs
# eaUmnBGgSGqqE2mD20iIV3d/cq++6Mlid/3fqHTB8KnGgrd8sSkbv/oZ9zM8D1If
# lNQtZ3eO0upc/YCfFfG7KcuX0fey736/SboPUlYtE/ftgRVKZyelFyZ/P399wcEJ
# p8EoaF1RVPcGZGl0jQVhy8eGOK9KYYbM40/y9Y5IEmFiyvzpNtjQDC4kNkgajbQm
# OPScFeLcY7Axa38ftONV/1URdRcLtnicKsNWa6Q9EHj5pRR0rG6KQB5FNygzNw5a
# zrunZc7X3qQPzTSpsflrFNDL+H/8ISLLjO2+9N8gWdZm2smtwwIKODGjr0Jc0PgM
# B4b8P7Rgka6fEVD6Y83jzQ0Y4QG2f3E8YcXIfJFV6FVMvlsOxyq6pGDYclKvld6a
# +g+q2/JvpJyAkf8u1e0XLZINzuBaTbxUB6ik/wZace4x7fi+PBAqqnHPxi1xGq4d
# shgmjF5VkjKTKy7nUq2/FMgmCaH6+EjAmPUifgu6EdCMy5WAPcGwtg+ZpXrqOrp4
# LHewUmdZL/qRnz7jzh5zIHi7Hww60t2++D+hMz/qKubldf+9vpkRqnOVuiEmFF16
# 1vKPC54DbAO/V4HuIG7Rjg1AJAT5q80Cx2XtDNitx/UxXH6XKd3VVuEaUt3petmR
# s+iGPbjPtw==
# SIG # End signature block
