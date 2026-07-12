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
Fetch per-device compliance policy states and optional setting states. This is detailed and can be slow on large tenants.

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
    Version : 1.6

.VERSION
1.6


.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication.
    Minimum Graph application permissions: DeviceManagementManagedDevices.Read.All; DeviceManagementConfiguration.Read.All; Device.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Version : 1.6
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
    Import-Module $modulePath -ErrorAction Stop
} catch {
    Write-Host "Failed to import SmartM365.Core module from '$modulePath' : $_" -ForegroundColor Red
    exit 1
}

# ==========================================================
# Fixed output paths and transcript
# ==========================================================
$ScriptVersion = "1.6"
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
$script:MaxDevicesEffective = if ($PSBoundParameters.ContainsKey('MaxDevices')) { [int]$MaxDevices } else { [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'MaxDevices' -DefaultValue 0) }
if ($script:MaxDevicesEffective -lt 0) { $script:MaxDevicesEffective = 0 }
$script:IncludePolicyStatesEffective = if ($PSBoundParameters.ContainsKey('IncludePolicyStates')) { [bool]$IncludePolicyStates } else { [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'IncludePolicyStates' -DefaultValue $false) }
$script:EnableDirectoryEnrichmentEffective = if ($PSBoundParameters.ContainsKey('EnableDirectoryEnrichment')) { [bool]$EnableDirectoryEnrichment } else { [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableDirectoryEnrichment' -DefaultValue $false) }
$script:MaxPolicyStateFailures = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'MaxPolicyStateFailures' -DefaultValue 100)
$script:MaxConsecutivePolicyStateFailures = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'MaxConsecutivePolicyStateFailures' -DefaultValue 25)
$script:PolicyStateFailureCount = 0
$script:ConsecutivePolicyStateFailures = 0
$script:PolicyStateCollectionDisabled = -not $script:IncludePolicyStatesEffective
$script:ComplianceFatalError = $null

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

foreach ($dir in @($ScriptCsvLogFolderPath, $LatestCsvFolderPath, $logDir)) {
    try { New-Item -ItemType Directory -Force -Path $dir | Out-Null } catch { }
}

$global:LogPath = $logDir
$global:LogTextFile = Join-Path $logDir ("{0}-{1}.log" -f $ScriptName, (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'))
$global:logTranscriptFile = Join-Path $logDir ("{0}-{1}_Transcript.log" -f $ScriptName, (Get-Date -Format 'yyyy-MM-dd_HH-mm-ss'))
Set-SmartM365CoreContext -RunId $ts -RunOutputRoot $ScriptCsvLogFolderPath -LatestOutputRoot $LatestCsvFolderPath -LogPath $global:LogTextFile

try {
    $transcriptPath = $global:logTranscriptFile
    Start-Transcript -Path $transcriptPath -Force | Out-Null
} catch {
    Write-Warning "Failed to start transcript. $_"
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
            Write-Warning ("{0} transient failure. Status={1}; attempt {2}/{3}; retrying in {4}s." -f $Operation, $statusRetryText, $attempt, $MaxAttempts, $delay)
            Start-Sleep -Seconds $delay
        }
    }
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
    $nextLink = $Uri
    $pageNumber = 0

    while ($nextLink) {
        $pageNumber++
        $currentUri = $nextLink
        $page = Invoke-WithRetry -Operation $Operation -Script {
            Invoke-MgGraphRequest -Method GET -Uri $currentUri -ErrorAction Stop
        }

        if ($page -and $page.value) {
            foreach ($item in @($page.value)) {
                $items.Add($item) | Out-Null
                if ($MaxItems -gt 0 -and $items.Count -ge $MaxItems) { break }
            }
        }

        Write-Host ("{0}: page {1}, total {2}" -f $Operation, $pageNumber, $items.Count) -ForegroundColor DarkCyan

        if ($MaxItems -gt 0 -and $items.Count -ge $MaxItems) { break }
        $nextLink = if ($page.'@odata.nextLink') { $page.'@odata.nextLink' } else { $null }
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
                    if ($policy.ContainsKey($prop)) {
                        $val = $policy[$prop]
                    } elseif ($policy.PSObject.Properties[$prop]) {
                        $val = $policy.PSObject.Properties[$prop].Value
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
        Write-Warning ("Failed to retrieve policy definition for '{0}': {1}" -f $PolicyId, $_.Exception.Message)
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
        Write-Host "Connect switch specified: existing Graph session (if any) will be disconnected and reconnected..." -ForegroundColor Cyan
        Disconnect-SmartM365CloudSession -ExchangeOnline $false -Graph $true -VerboseDisconnect:$true
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
            Write-Host "Connecting to Microsoft Graph with app-only certificate authentication..." -ForegroundColor Cyan
        } else {
            Write-Host "Connecting to Microsoft Graph with interactive authentication..." -ForegroundColor Cyan
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
            Write-Host "Retrieving all Intune managed Windows devices with explicit Graph paging..." -ForegroundColor Cyan
            if ($script:MaxDevicesEffective -gt 0) {
                Write-Host ("MaxDevices smoke cap active: {0}" -f $script:MaxDevicesEffective) -ForegroundColor Yellow
            }
            $devices = @(Get-ManagedWindowsDevicesFast -MaxItems $script:MaxDevicesEffective)
            if (-not $devices -or $devices.Count -eq 0) {
                Write-Error "No Windows managed devices found."
                throw "No Windows managed devices found."
            }
        } elseif (-not [string]::IsNullOrWhiteSpace($ManagedDeviceId)) {
            Write-Host "Resolving device by ManagedDeviceId '$ManagedDeviceId'..." -ForegroundColor Cyan
            $devices = @(Get-ManagedWindowsDevicesFast -ManagedDeviceId $ManagedDeviceId -MaxItems 1)
            if (-not $devices -or $devices.Count -eq 0) {
                throw ("No device found with ManagedDeviceId='{0}'." -f $ManagedDeviceId)
            }
        } else {
            Write-Host "Resolving device by DeviceName '$DeviceName'..." -ForegroundColor Cyan
            $devices = @(Get-ManagedWindowsDevicesFast -DeviceName $DeviceName -MaxItems 1)
            if (-not $devices -or $devices.Count -eq 0) {
                throw ("No Windows device found for DeviceName='{0}'." -f $DeviceName)
            }
        }

        Write-Host ("Managed Windows devices selected for compliance summary: {0}" -f @($devices).Count) -ForegroundColor Cyan
    } catch {
        Write-Error "Failed to resolve target devices. $_"
        throw
    }

    # 2) Collect per-device summary rows and per-policy rows
    $aadCache = @{}
    $rows     = New-Object System.Collections.Generic.List[object]
    $polAll   = New-Object System.Collections.Generic.List[object]

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
            $cmd = Get-Command -Name Get-MgDeviceManagementManagedDeviceDeviceCompliancePolicyState -ErrorAction SilentlyContinue
            if ($cmd) {
                $policyStates = Invoke-WithRetry -Script {
                    Get-MgDeviceManagementManagedDeviceDeviceCompliancePolicyState -ManagedDeviceId $dev.Id -All -ErrorAction Stop
                }
            } else {
                $uri  = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$($dev.Id)/deviceCompliancePolicyStates`?$top=200"
                $vals = @()
                while ($uri) {
                    $resp = Invoke-WithRetry -Operation "Get Intune Graph page" -Script { Invoke-MgGraphRequest -Method GET -Uri $uri -ErrorAction Stop }
                    if ($resp -and $resp.value) { $vals += $resp.value }
                    $uri = if ($resp.'@odata.nextLink') { $resp.'@odata.nextLink' } else { $null }
                }
                if ($vals.Count -gt 0) { $policyStates = $vals }
            }
            $script:ConsecutivePolicyStateFailures = 0
        } catch {
            $script:PolicyStateFailureCount++
            $script:ConsecutivePolicyStateFailures++
            $shortPolicyStateError = Get-ShortGraphErrorMessage -ErrorRecord $_
            Write-Warning ("Failed to retrieve policy states for device {0}: {1}" -f $dev.DeviceName, $shortPolicyStateError)
            if (($script:MaxPolicyStateFailures -gt 0 -and $script:PolicyStateFailureCount -ge $script:MaxPolicyStateFailures) -or
                ($script:MaxConsecutivePolicyStateFailures -gt 0 -and $script:ConsecutivePolicyStateFailures -ge $script:MaxConsecutivePolicyStateFailures)) {
                $script:PolicyStateCollectionDisabled = $true
                Write-Warning ("Policy state collection disabled for this run after {0} total failure(s), {1} consecutive. Device summary processing will continue." -f $script:PolicyStateFailureCount, $script:ConsecutivePolicyStateFailures)
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

                foreach ($p in $targets) {
                    $pIdx++
                    if ($pTot -gt 0) {
                        Write-Progress -Id 2 -ParentId 1 `
                            -Activity ("Policies for {0}" -f $dev.DeviceName) `
                            -Status ("{0}/{1} - {2}" -f $pIdx, $pTot, $p.displayName) `
                            -PercentComplete ([int](100 * $pIdx / $pTot))
                    }

                    try {
                        $u = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$($dev.Id)/deviceCompliancePolicyStates/$([uri]::EscapeDataString($p.id))/settingStates`?$select=setting,state&`$top=200"
                        $s = @()
                        while ($u) {
                            $page = Invoke-WithRetry -Operation "Get Intune compliance setting states" -Script { Invoke-MgGraphRequest -Method GET -Uri $u -ErrorAction Stop }
                            if ($page -and $page.value) { $s += $page.value }
                            $u = if ($page.'@odata.nextLink') { $page.'@odata.nextLink' } else { $null }
                        }

                        foreach ($v in ($s | Where-Object { $_.state -eq 'nonCompliant' })) {
                            $cat = Map-SettingCategory -SettingName $v.setting
                            if (-not $policyCategoryRollup.ContainsKey($p.displayName)) {
                                $policyCategoryRollup[$p.displayName] = @{}
                            }
                            $policyCategoryRollup[$p.displayName][$cat] = 'Fail'
                        }
                    } catch {
                        Write-Warning ("Failed to retrieve setting states for '{0}' on device '{1}': {2}" -f $p.displayName, $dev.DeviceName, $_.Exception.Message)
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

    # 3) Output
    if ($rows.Count -gt 0) {
        Write-Host ""
        Write-Host ("Devices compliance summary: {0} row(s)" -f $rows.Count) -ForegroundColor Cyan

        $summaryOut = $rows |
            Sort-Object DeviceName |
            Select-Object `
                DeviceName, AzureADDeviceId, EntraObjectId, Manufacturer, Model, OperatingSystem, LastSyncDateTime, `
                ComplianceState, ComplianceGracePeriodExpirationDateTime, `
                AD_Domain, AD_OU, DirectorySource

        $summaryOut | Select-Object -First 25 | Format-Table -AutoSize
        if ($summaryOut.Count -gt 25) {
            Write-Host ("Displayed first 25 of {0} device summary rows." -f $summaryOut.Count) -ForegroundColor DarkCyan
        }

        try {
            Write-SmartM365CsvAtomically -Data @($summaryOut) -Path $mainCsv
            Export-SmartM365Csv -Data @($summaryOut) -TimestampedPath $tsCsv -LatestPath $lastCsv | Out-Null

            Write-Host "Compliance summary CSV saved: $mainCsv"
            Write-Host "Compliance summary CSV timestamped: $tsCsv"
            Write-Host "Compliance summary CSV (last): $lastCsv"
        } catch {
            Write-Warning "Failed to export compliance summary CSVs: $_"
        }
    } else {
        Write-Host ""
        Write-Host "No devices to display or export." -ForegroundColor Yellow
    }

    if (-not $script:IncludePolicyStatesEffective) {
        Write-Host ""
        Write-Host "Compliance policy detail collection skipped by default. Use -IncludePolicyStates `$true for detailed per-policy CSV." -ForegroundColor Yellow
    } elseif ($polAll.Count -gt 0) {
        Write-Host ""
        Write-Host ("Compliance details per policy: {0} row(s)" -f $polAll.Count) -ForegroundColor Cyan

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
            Write-Host ("Displayed first 50 of {0} policy detail rows." -f $polOut.Count) -ForegroundColor DarkCyan
        }

        try {
            Write-SmartM365CsvAtomically -Data @($polOut) -Path $policyMainCsv
            Export-SmartM365Csv -Data @($polOut) -TimestampedPath $policyTsCsv -LatestPath $policyLastCsv | Out-Null

            Write-Host "Compliance policy CSV saved: $policyMainCsv"
            Write-Host "Compliance policy CSV timestamped: $policyTsCsv"
            Write-Host "Compliance policy CSV (last): $policyLastCsv"
        } catch {
            Write-Warning "Failed to export compliance policy CSVs: $_"
        }
    } else {
        Write-Host ""
        Write-Host "Compliance details: no policy states available or calls failed." -ForegroundColor Yellow
    }
}
catch {
    $script:ComplianceFatalError = $_
    Write-Host "A global error occurred in M365-Devices-Compliance.ps1 : $($_.Exception.Message)" -ForegroundColor Red
    Write-Error $script:ComplianceFatalError
}
finally {
    # Disconnect Graph only if we connected it in this run
    if ($connectedGraphInThisRun) {
        Write-Host "`n--- Disconnect Cloud Services ---"
        try {
            Disconnect-SmartM365CloudSession -ExchangeOnline:$false -Graph:$true -VerboseDisconnect:$true
        } catch {
            Write-Host ("Error during Graph disconnect in finally: {0}" -f $_) -ForegroundColor Yellow
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
        Complete-SmartM365ExecutionContext -Status $finalStatus
    } catch { }
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCC045PnznRTy0R3
# dZiXQ8PIWsgvf3PxjvYoJ3403yvjHaCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAoaOt/VhSrHE/9rk63
# Dou9w9qdgNQSGALgQFX+RP78+DANBgkqhkiG9w0BAQEFAASCAYAJWnfYpPYmr4FS
# OKoKDnRimB4DdrlLobP70eL/R5OqOO6Z7/hE0PY3iwrQo9gIok3RvuUPjEeSJYb4
# kxWZQC/E+5cFhzJTJUHaU1BtZoKab11Qfz513OkkhI0Jalm5qeEWdVV7RTa5hi1W
# kKK8WaAodd2d+jdS+WiEiBEV8tq7GknwIIuic50F9d8VcZWFGNx2o90mqSHdqFVS
# lOdDixy+Wyw6OWvY5ZvNCblx+TxisFcNbZ6ebFlRa0StKTsovrJfICsV97D9o2M0
# BAHYo1jgg2wZLJXfKV5IJiR7L088B33v3W/YEuOVP0b28oDOPkW9Lz7mOSD7Vdl+
# CL+YQWJ0VHAhoEalmzsafwltY4CQTAoroKtOVxCWhiGTEiuhEqY5dkt7ePxQ7W+7
# QZMMQFh49XhgPQ2Mr6S2rVbSgHHgCtUMtSddt8i0Q3Trrrhur+k8lee8T6U0w9UN
# J/02l0m1j90ukIOC8lqLSTS56Rh5Olg8Oh2sWGwMleL95Q8T/XA=
# SIG # End signature block
