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
.VERSION
1.1

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Script  : Intune-DiscoveredApps-Inventory
    Version : 1.1
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
    [int]$DelayMs = 300
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
$ScriptVersion = "1.1"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion"
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DiscoveredAppsCsvLogFolderPath' -DefaultValue $OutputPath
if (-not $PSBoundParameters.ContainsKey('DelayMs')) {
    $DelayMs = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'GraphRequestDelayMs' -DefaultValue 1000)
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
    WriteLog -Message "PowerShell version : $($PSVersionTable.PSVersion)"
    WriteLog -Message "MaxApps            : $(if ($MaxApps -eq 0) { 'unlimited' } else { $MaxApps })"
    WriteLog -Message "DryRun             : $DryRun"
    WriteLog -Message "DelayMs            : $DelayMs"
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
    # Build Summary records (one row per app)
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

    # ----------------------------------------------------------
    # Retrieve managed devices per app (DeviceDetail records)
    # ----------------------------------------------------------
    WriteLog -Message "Retrieving managed devices for $($windowsApps.Count) Windows apps..." "INFO"
    $detailRecords = [System.Collections.Generic.List[psobject]]::new()
    $appIndex      = 0

    foreach ($app in $windowsApps) {
        $appIndex++
        $pctComplete = [math]::Round(($appIndex / $windowsApps.Count) * 100, 1)

        Write-Progress `
            -Activity        'Retrieving device details per app' `
            -Status          "App $appIndex / $($windowsApps.Count): $($app.displayName) v$($app.version)  ($pctComplete%)" `
            -PercentComplete $pctComplete

        $devicesUri = "https://graph.microsoft.com/v1.0/deviceManagement/detectedApps/$($app.id)/managedDevices" +
                      '?$top=999&$select=id,deviceName,operatingSystem,osVersion,userPrincipalName,' +
                      'lastSyncDateTime,enrolledDateTime,managedDeviceOwnerType,complianceState'

        try {
            if ($DelayMs -gt 0) { Start-Sleep -Milliseconds $DelayMs }
            $devices = Invoke-GraphPagedRequest -InitialUri $devicesUri

            if ($devices.Count -eq 0) {
                # Preserve app row even when no device is returned
                $detailRecords.Add([pscustomobject]@{
                    AppId                  = $app.id
                    AppName                = $app.displayName
                    AppVersion             = $app.version
                    AppPublisher           = $app.publisher
                    Platform               = $app.platform
                    DeviceId               = ''
                    DeviceName             = ''
                    OperatingSystem        = ''
                    OSVersion              = ''
                    UserPrincipalName      = ''
                    LastSyncDateTime       = ''
                    EnrolledDateTime       = ''
                    ManagedDeviceOwnerType = ''
                    ComplianceState        = ''
                    AzureADDeviceId        = ''
                })
            } else {
                foreach ($device in $devices) {
                    $detailRecords.Add([pscustomobject]@{
                        AppId                  = $app.id
                        AppName                = $app.displayName
                        AppVersion             = $app.version
                        AppPublisher           = $app.publisher
                        Platform               = $app.platform
                        DeviceId               = $device.id
                        DeviceName             = $device.deviceName
                        OperatingSystem        = $device.operatingSystem
                        OSVersion              = $device.osVersion
                        UserPrincipalName      = $device.userPrincipalName
                        LastSyncDateTime       = $device.lastSyncDateTime
                        EnrolledDateTime       = $device.enrolledDateTime
                        ManagedDeviceOwnerType = $device.managedDeviceOwnerType
                        ComplianceState        = $device.complianceState
                    })
                    $script:Stat_DeviceDetailRows++
                }
            }
            $script:Stat_AppsProcessed++
        } catch {
            WriteLog -Message "Failed to retrieve devices for app '$($app.displayName)' (Id=$($app.id)): $_" "WARNING"
            $script:Stat_AppsSkipped++
        }
    }

    Write-Progress -Activity 'Retrieving device details per app' -Completed

    WriteLog -Message "Device detail retrieval complete." "INFO"
    WriteLog -Message "  Detail rows (app/device pairs) : $($script:Stat_DeviceDetailRows)" "INFO"
    WriteLog -Message "  Apps processed successfully    : $($script:Stat_AppsProcessed)" "INFO"
    WriteLog -Message "  Apps skipped (errors)          : $($script:Stat_AppsSkipped)" "INFO"

    # ----------------------------------------------------------
    # Export CSV files
    # ----------------------------------------------------------
    if ($DryRun) {
        WriteLog -Message "DryRun mode: CSV export skipped. Summary=$($summaryRecords.Count) rows, Detail=$($detailRecords.Count) rows." "WARNING"
    } else {
        $globalPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''

        WriteLog -Message "Exporting Summary CSV ($($summaryRecords.Count) rows)..." "INFO"
        ExportAndCopyCsv `
            -BaseFileName      "Intune_DiscoveredApps_Summary" `
            -OutputPath        $OutputPath `
            -GlobalPath        $globalPath `
            -Data              $summaryRecords `
            -Encoding          "UTF8" `
            -NoTypeInformation
        WriteLog -Message "Summary export completed: $global:csvFilePath1" "INFO"

        WriteLog -Message "Exporting DeviceDetail CSV ($($detailRecords.Count) rows)..." "INFO"
        ExportAndCopyCsv `
            -BaseFileName      "Intune_DiscoveredApps_DeviceDetail" `
            -OutputPath        $OutputPath `
            -GlobalPath        $globalPath `
            -Data              $detailRecords `
            -Encoding          "UTF8" `
            -NoTypeInformation
        WriteLog -Message "DeviceDetail export completed: $global:csvFilePath1" "INFO"
    }

} catch {
    $globalError = $_
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

    WriteLog -Message "=== Run summary ==="
    WriteLog -Message "  Total apps retrieved (all platforms) : $($script:Stat_AppsTotal)"
    WriteLog -Message "  Windows apps found                   : $($script:Stat_AppsWindows)"
    WriteLog -Message "  Apps processed successfully          : $($script:Stat_AppsProcessed)"
    WriteLog -Message "  Apps skipped (errors)                : $($script:Stat_AppsSkipped)"
    WriteLog -Message "  Device detail rows exported          : $($script:Stat_DeviceDetailRows)"
    WriteLog -Message "  Graph API calls                      : $($script:Stat_GraphCalls)"
    WriteLog -Message "  Throttle retries (429)               : $($script:Stat_ThrottleRetries)"
    WriteLog -Message "  Elapsed time                         : $elapsedStr"
    WriteLog -Message "$TaskName completed."

    try {
        $summaryStatus = if ($script:RunError) { 'Failed' } else { 'Auto' }
        Complete-SmartM365ExecutionContext -Status $summaryStatus -ErrorRecord $script:RunError
    } catch {
        WriteLog -Message ("Failed to write execution summary: {0}" -f $_) "WARNING"
    }
    try { Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {} } catch {}
}
