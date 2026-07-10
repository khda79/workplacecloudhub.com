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
    Version : 1.5

.VERSION
1.5

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Script  : Intune-DiscoveredApps-Inventory
    Version : 1.5
    Requires: Microsoft.Graph.Authentication module
              SmartM365.Core module (Modules\SmartM365.Core\SmartM365.Core.psd1)
    Local configuration: DiscoveredAppsCsvLogFolderPath -> output folder (DATA-ALL\M365-Inventory\Output-Windows-Discovered apps)
              LatestCsvFolderPath -> DATA-LAST folder (GlobalPath / copy destination)
    Graph permission required (application): DeviceManagementApps.Read.All
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
    [switch]$ResetResume
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

# ==========================================================
# Import SmartM365.Core module
# ==========================================================
$modulePath = & { $d = $PSScriptRoot; while ($d) { $p = Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'; if (Test-Path -LiteralPath $p) { return $p }; $parent = Split-Path -Path $d -Parent; if ($parent -eq $d) { break }; $d = $parent }; throw 'SmartM365.Core module not found.' }
try {
    Import-Module $modulePath -ErrorAction Stop
} catch {
    Write-Host "Failed to import SmartM365.Core module from '$modulePath': $_" -ForegroundColor Red
    exit 1
}

# ==========================================================
# Script metadata
# ==========================================================
$ScriptVersion = "1.5"
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
$script:DeviceDetailResumePath = ''
$script:DeviceDetailPartialPath = ''
$script:DeviceDetailTimestampedPath = ''
$script:DeviceDetailCompleted = $false
$script:Stat_DetailAppsTargeted = 0
$script:Stat_GraphCalls       = 0
$script:Stat_ThrottleRetries  = 0

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
                WriteLog -Message ("Graph transient failure on [$currentUri]. Status={0}; attempt {1}/{2}; waiting {3}s." -f $statusRetryText, $attempt, $MaxRetries, $retryAfter) "WARNING"
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
        $Rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Append
    } else {
        $Rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
    }
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

function Save-DiscoveredAppsResumeState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$PartialPath,
        [Parameter(Mandatory = $true)][string]$TimestampedPath,
        [Parameter(Mandatory = $true)][string]$Mode,
        [Parameter(Mandatory = $true)][int]$TargetCount,
        [Parameter(Mandatory = $true)][string[]]$ProcessedAppIds,
        [Parameter(Mandatory = $true)][int]$ProcessedCount,
        [Parameter(Mandatory = $true)][int]$SkippedCount,
        [Parameter(Mandatory = $true)][int]$DetailRows
    )

    $state = [ordered]@{
        Script           = $TaskName
        DeviceDetailMode = $Mode
        TargetCount      = $TargetCount
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
        WriteLog -Message "Connect switch specified: disconnecting any existing Graph session and reconnecting." "INFO"
        try { Disconnect-MgGraph | Out-Null } catch {}
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
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -GraphProbeUris @('https://graph.microsoft.com/v1.0/deviceManagement/detectedApps?$top=1') | Out-Null

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

        $detailBaseFileName = 'Intune_DiscoveredApps_DeviceDetail'
        $detailTimestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
        $script:DeviceDetailResumePath = Join-Path -Path $OutputPath -ChildPath "$detailBaseFileName.resume.json"
        $processedAppIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        $streamingEnabled = -not $DryRun

        if ($streamingEnabled -and $ResetResume) {
            $oldState = Get-DiscoveredAppsResumeState -Path $script:DeviceDetailResumePath
            if ($oldState -and $oldState.PartialPath) { Remove-Item -LiteralPath ([string]$oldState.PartialPath) -Force -ErrorAction SilentlyContinue }
            if ($oldState -and $oldState.TimestampedPath) { Remove-Item -LiteralPath ([string]$oldState.TimestampedPath) -Force -ErrorAction SilentlyContinue }
            Remove-Item -LiteralPath $script:DeviceDetailResumePath -Force -ErrorAction SilentlyContinue
            WriteLog -Message "ResetResume enabled; previous DeviceDetail resume state removed." "INFO"
        }

        $resumeState = if ($streamingEnabled) { Get-DiscoveredAppsResumeState -Path $script:DeviceDetailResumePath } else { $null }
        if ($resumeState -and $resumeState.DeviceDetailMode -eq $DeviceDetailMode -and $resumeState.PartialPath -and (Test-Path -LiteralPath ([string]$resumeState.PartialPath))) {
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
            if ($resumeState) { WriteLog -Message "Existing DeviceDetail resume state is incompatible or partial CSV is missing; starting a new DeviceDetail export." "WARNING" }
            $script:DeviceDetailPartialPath = Join-Path -Path $OutputPath -ChildPath "$detailBaseFileName`_$detailTimestamp.partial.csv"
            $script:DeviceDetailTimestampedPath = Join-Path -Path $OutputPath -ChildPath "$detailBaseFileName`_$detailTimestamp.csv"
            Remove-Item -LiteralPath $script:DeviceDetailPartialPath -Force -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath $script:DeviceDetailResumePath -Force -ErrorAction SilentlyContinue
        }

        $appIndex = 0
        foreach ($app in $detailApps) {
            $appIndex++
            if ($processedAppIds.Contains([string]$app.id)) {
                $script:Stat_DetailAppsSkippedByResume++
                continue
            }

            $pctComplete = [math]::Round(($appIndex / $detailApps.Count) * 100, 1)
            Write-Progress `
                -Activity "Collecting Intune discovered app device details" `
                -Status "App $appIndex / $($detailApps.Count): $($app.displayName) v$($app.version) ($pctComplete%)" `
                -PercentComplete $pctComplete

            if ($ProgressEveryApps -gt 0 -and (($appIndex % $ProgressEveryApps -eq 0) -or $appIndex -eq $detailApps.Count)) {
                WriteLog -Message ("Device detail progress: {0}/{1} apps ({2}%). DetailRows={3}; Processed={4}; Skipped={5}; ResumedSkipped={6}; GraphCalls={7}; ThrottleRetries={8}" -f $appIndex, $detailApps.Count, $pctComplete, $script:Stat_DeviceDetailRows, $script:Stat_AppsProcessed, $script:Stat_AppsSkipped, $script:Stat_DetailAppsSkippedByResume, $script:Stat_GraphCalls, $script:Stat_ThrottleRetries) "INFO"
            }

            try {
                if ($DelayMs -gt 0) { Start-Sleep -Milliseconds $DelayMs }
                $devicesUri = 'https://graph.microsoft.com/v1.0/deviceManagement/detectedApps/' + $app.id + '/managedDevices' +
                    '?$top=999&$select=id,deviceName,operatingSystem,osVersion,userPrincipalName,' +
                    'lastSyncDateTime,enrolledDateTime,managedDeviceOwnerType,complianceState,azureADDeviceId'
                $devices = Invoke-GraphPagedRequest -InitialUri $devicesUri

                $appRows = [System.Collections.Generic.List[psobject]]::new()
                if (-not $devices -or $devices.Count -eq 0) {
                    $appRows.Add((New-DiscoveredAppsDeviceDetailRecord -App $app -Device $null))
                    $script:Stat_DeviceDetailRows++
                } else {
                    foreach ($device in $devices) {
                        $appRows.Add((New-DiscoveredAppsDeviceDetailRecord -App $app -Device $device))
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

    WriteLog -Message ("DeviceDetail completed. TargetApps={0}; Processed={1}; Skipped={2}; ResumedSkipped={3}; Rows={4}" -f $script:Stat_DetailAppsTargeted, $script:Stat_AppsProcessed, $script:Stat_AppsSkipped, $script:Stat_DetailAppsSkippedByResume, $script:Stat_DeviceDetailRows) "INFO"

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
    WriteLog -Message "  Device detail rows exported          : $($script:Stat_DeviceDetailRows)"
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
