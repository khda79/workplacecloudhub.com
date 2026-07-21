#Requires -Version 7.0
<#
.SYNOPSIS
Exports Windows Autopatch-related alert data from Microsoft Intune / Microsoft Graph.

.DESCRIPTION
This script builds a practical Autopatch alert export by combining Intune report exports that are documented in Microsoft Graph:
- FeatureUpdateDeviceState
- QualityUpdateDeviceStatusByPolicy
- QualityUpdateDeviceErrorsByPolicy
- QualityUpdatePolicyStatusSummary

Important limitation:
The Intune export/report APIs expose alert messages and event timestamps, but they do not expose the exact same
UI-level Autopatch Alerts grid schema shown in the portal for all alert types. Because of that, the script returns:
- AlertName
- AlertCount
- FirstSeenUtc
- LastSeenUtc
- Category
- AffectedUpdateType
- Severity
- SourceReport
- PolicyId
- PolicyName

Severity and some UI-specific metadata are inferred when possible; otherwise they are set to Unknown.

.PARAMETER OutputFolder
Folder where CSV files are written.

.PARAMETER IncludeFeatureUpdates
Exports feature update alerts.

.PARAMETER IncludeQualityUpdates
Exports expedited quality update alerts.

.PARAMETER TenantId
Optional tenant ID used at connect time.

.PARAMETER UseDeviceCode
Uses device code authentication.

.EXAMPLE
pwsh -File .\SmartM365-Intune-WindowsAutopatch-Alerts-Inventory.ps1
.VERSION
1.15



.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication.
    Minimum Graph application permissions: DeviceManagementConfiguration.Read.All; DeviceManagementManagedDevices.Read.All; DeviceManagementApps.Read.All.
    Conditional: Mail.Send is required only when Graph mail is used; Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
Author    : https://github.com/khda79/workplacecloudhub.com
    Version : 1.15
    Minimum application permissions: DeviceManagementConfiguration.Read.All, DeviceManagementManagedDevices.Read.All, DeviceManagementApps.Read.All
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
[Parameter()]
    [string]$OutputFolder,

    [Parameter()]
    [switch]$IncludeFeatureUpdates = $true,

    [Parameter()]
    [switch]$IncludeQualityUpdates = $true,

    [Parameter()]
    [string]$TenantId,

    [Parameter()]
    [switch]$UseDeviceCode,
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
$script:SmartM365GlobalConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

if ($PSScriptRoot) {
    $ScriptRoot = $PSScriptRoot
}
elseif ($MyInvocation.MyCommand.Path) {
    $ScriptRoot = Split-Path -Path $MyInvocation.MyCommand.Path -Parent
}
else {
    throw 'Unable to determine script directory.'
}

if ($null -eq $script:SmartM365GlobalConfig) { $script:SmartM365GlobalConfig = [pscustomobject]@{} }

function Import-SmartM365CorePreflight {
    if (Get-Command Invoke-CoreSmartM365Preflight -ErrorAction SilentlyContinue) { return }

        $scriptRootValue = Get-Variable -Name ScriptRoot -ValueOnly -ErrorAction SilentlyContinue
    $searchRoot = if ($PSScriptRoot) { $PSScriptRoot } elseif ($scriptRootValue) { $scriptRootValue } elseif ($PSCommandPath) { Split-Path -Path $PSCommandPath -Parent } else { (Get-Location).Path }
    while ($searchRoot) {
        $modulePath = Join-Path -Path $searchRoot -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
        if (Test-Path -LiteralPath $modulePath) {
            Import-Module -Name $modulePath -MinimumVersion '1.0.24' -Prefix Core -ErrorAction Stop
            return
        }

        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }

    throw 'SmartM365.Core module was not found. Preflight checks cannot run.'
}

function Get-ScriptLocalConfig {
    [CmdletBinding()]
    param()

    $configPath = Join-Path -Path $ScriptRoot -ChildPath ("{0}.local.json" -f [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))
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
            $scriptRootValue = Get-Variable -Name ScriptRoot -ValueOnly -ErrorAction SilentlyContinue
    $searchRoot = if ($PSScriptRoot) { $PSScriptRoot } elseif ($scriptRootValue) { $scriptRootValue } elseif ($PSCommandPath) { Split-Path -Path $PSCommandPath -Parent } else { (Get-Location).Path }
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
        $searchRoot = $ScriptRoot
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
$OutputFolder = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'OutputFolder' -DefaultValue $OutputFolder
$TenantId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue $TenantId
$DataAllRootPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DataAllRootPath' -DefaultValue ''
$LatestCsvFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''
$AppId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AppId' -DefaultValue ''
$Thumb = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumb' -DefaultValue ''
$Thumbprint = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumbprint' -DefaultValue $Thumb
$global:RetentionMaxCSV = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxLogs' -DefaultValue 30)
$global:ErrorMailTo = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ErrorMailTo' -DefaultValue ''
$global:EnableSharePointUpload = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointTargetFolderPath' -DefaultValue ''
$LogAllRootPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LogAllRootPath' -DefaultValue ''

if ([string]::IsNullOrWhiteSpace($OutputFolder)) {
    if (-not [string]::IsNullOrWhiteSpace($DataAllRootPath)) {
        $OutputFolder = Join-Path -Path $DataAllRootPath -ChildPath 'Intune\WindowsUpdate\AutopatchAlerts'
    }
    else {
        $OutputFolder = $ScriptRoot
    }
}
if ([string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) {
    $LatestCsvFolderPath = $OutputFolder
}
$ScriptVersion = "1.15"
$ScriptName = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$StartTime = Get-Date
$RunStamp = $StartTime.ToString('yyyyMMdd_HHmmss')
$RunId = [guid]::NewGuid().Guid
$LogFolder = if ([string]::IsNullOrWhiteSpace($LogAllRootPath)) {
    Join-Path -Path $OutputFolder -ChildPath 'Logs'
} else {
    Join-Path -Path $LogAllRootPath -ChildPath $ScriptName
}
$LogFile = Join-Path -Path $LogFolder -ChildPath ("{0}_{1}_{2:yyyyMMdd_HHmmss}.log" -f $ScriptName, $ScriptVersion, $StartTime)
$TranscriptFile = Join-Path -Path $LogFolder -ChildPath ("{0}_{1}_{2:yyyyMMdd_HHmmss}_Transcript.log" -f $ScriptName, $ScriptVersion, $StartTime)
$global:SmartM365ExecutionStartTime = $StartTime
$global:SmartM365ExecutionSummaryWritten = $false
$global:SmartM365WarningCount = 0
$global:SmartM365ErrorCount = 0
$global:SmartM365ScriptName = $ScriptName
$global:BasePath = $OutputFolder
$global:LogPath = $LogFolder
$global:LogTextFile = $LogFile
$global:logTranscriptFile = $TranscriptFile
$SummaryCsvPath = Join-Path -Path $OutputFolder -ChildPath ("Intune_WindowsAutopatch_Alerts_Summary_{0}.csv" -f $RunStamp)
$DetailCsvPath = Join-Path -Path $OutputFolder -ChildPath ("Intune_WindowsAutopatch_Alerts_Detail_{0}.csv" -f $RunStamp)
$PolicyCsvPath = Join-Path -Path $OutputFolder -ChildPath ("Intune_WindowsAutopatch_Alerts_PolicySummary_{0}.csv" -f $RunStamp)
$SummaryLatestCsvPath = Join-Path -Path $LatestCsvFolderPath -ChildPath "Intune_WindowsAutopatch_Alerts_Summary.csv"
$DetailLatestCsvPath = Join-Path -Path $LatestCsvFolderPath -ChildPath "Intune_WindowsAutopatch_Alerts_Detail.csv"
$PolicyLatestCsvPath = Join-Path -Path $LatestCsvFolderPath -ChildPath "Intune_WindowsAutopatch_Alerts_PolicySummary.csv"
$PolicyColumns = @('PolicyId','PolicyName','ExpediteQUReleaseDate','CountDevicesErrorStatus','CountDevicesInProgressStatus','CountDevicesSuccessStatus')
$DetailColumns = @('AlertName','Severity','Category','AffectedUpdateType','DeviceName','DeviceId','PolicyId','PolicyName','EventDateUtc','LastScanUtc','AggregateState','CurrentStatus','SourceReport')
$SummaryColumns = @('AlertName','Severity','Category','AffectedUpdateType','Impact','FirstSeenUtc','LastSeenUtc','SourceReport')

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [Parameter()]
        [ValidateSet('INFO','WARN','ERROR')]
        [string]$Level = 'INFO'
    )

    if ($Level -eq 'WARN') { $global:SmartM365WarningCount = [int]$global:SmartM365WarningCount + 1 }
    elseif ($Level -eq 'ERROR') { $global:SmartM365ErrorCount = [int]$global:SmartM365ErrorCount + 1 }
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $lines = @([regex]::Split(([string]$Message), '\r?\n') | ForEach-Object { "{0} [{1}] [{2}] {3}" -f $timestamp, $Level, $RunId, $_ })
    $lines | ForEach-Object { Write-Host $_ }
    Add-Content -Path $LogFile -Value $lines -Encoding utf8
}

function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -Path $Path)) {
        $null = New-Item -Path $Path -ItemType Directory -Force
    }
}

function Test-Ps7 {
    if ($PSVersionTable.PSVersion.Major -lt 7) {
        throw 'This script requires PowerShell 7 or later.'
    }
}

function Import-RequiredModule {
    param([Parameter(Mandatory)][string]$Name)
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Required module not found: $Name. Install it with: Install-Module $Name -Scope CurrentUser"
    }
    Import-Module -Name $Name -Force -ErrorAction Stop
}

function Connect-GraphSession {
    $scopes = @('DeviceManagementConfiguration.ReadWrite.All','DeviceManagementApps.ReadWrite.All','DeviceManagementManagedDevices.ReadWrite.All')
    $certificateThumbprint = if (-not [string]::IsNullOrWhiteSpace($Thumbprint)) { $Thumbprint } else { $Thumb }
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}

    if (-not [string]::IsNullOrWhiteSpace($AppId) -and -not [string]::IsNullOrWhiteSpace($TenantId) -and -not [string]::IsNullOrWhiteSpace($certificateThumbprint)) {
        Write-Log -Message 'Connecting to Microsoft Graph using app-only certificate authentication.'
        Connect-MgGraph -TenantId $TenantId -ClientId $AppId -CertificateThumbprint $certificateThumbprint -NoWelcome | Out-Null
    }
    else {
        $connectParams = @{ Scopes = $scopes; NoWelcome = $true }
        if ($TenantId) { $connectParams['TenantId'] = $TenantId }
        if ($UseDeviceCode) { $connectParams['UseDeviceCode'] = $true }
        Write-Log -Message ("Connecting to Microsoft Graph with delegated scopes: {0}" -f ($scopes -join ', '))
        Connect-MgGraph @connectParams | Out-Null
    }

    $context = Get-MgContext
    Write-Log -Message ("Connected to tenant [{0}] as [{1}]" -f $context.TenantId, $context.Account)
}

function Get-AutopatchGraphStatusCode {
    param([Parameter(Mandatory = $true)]$ErrorRecord)

    try { if ($ErrorRecord.Exception.Response.StatusCode) { return [int]$ErrorRecord.Exception.Response.StatusCode } } catch {}
    $text = [string]$ErrorRecord
    if ($text -match '\b(408|409|429|500|502|503|504)\b') { return [int]$Matches[1] }
    if ($text -match '(?i)TooManyRequests|throttl') { return 429 }
    return 0
}

function Get-AutopatchGraphRetryDelaySeconds {
    param([Parameter(Mandatory = $true)]$ErrorRecord, [Parameter(Mandatory = $true)][int]$Attempt)

    $jitter = Get-Random -Minimum 0 -Maximum 6
    try {
        $retryAfter = $ErrorRecord.Exception.Response.Headers.RetryAfter
        if ($retryAfter.Delta) { return [math]::Max(1, [math]::Min(300, ([int][math]::Ceiling($retryAfter.Delta.TotalSeconds) + $jitter))) }
    } catch {}
    return [math]::Min(300, (([int][math]::Pow(2, [math]::Min($Attempt, 6)) * 5) + $jitter))
}

function Invoke-AutopatchGraphRequest {
    param(
        [Parameter(Mandatory = $true)][ValidateSet('GET', 'POST')][string]$Method,
        [Parameter(Mandatory = $true)][string]$Uri,
        [AllowNull()][string]$Body,
        [ValidateRange(1, 12)][int]$MaxAttempts = 6
    )

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            $parameters = @{ Method = $Method; Uri = $Uri; OutputType = 'PSObject'; ErrorAction = 'Stop' }
            if ($Method -eq 'POST') { $parameters.Body = $Body; $parameters.ContentType = 'application/json' }
            return Invoke-MgGraphRequest @parameters
        }
        catch {
            $statusCode = Get-AutopatchGraphStatusCode -ErrorRecord $_
            if ($statusCode -notin @(408, 409, 429, 500, 502, 503, 504) -or $attempt -ge $MaxAttempts) { throw }
            $delay = Get-AutopatchGraphRetryDelaySeconds -ErrorRecord $_ -Attempt $attempt
            Write-Log -Message ("Graph transient failure HTTP {0}; retry {1}/{2} in {3}s: {4}" -f $statusCode, $attempt, $MaxAttempts, $delay, $Uri) -Level WARN
            Start-Sleep -Seconds $delay
        }
    }
}

function Invoke-GraphGetAll {
    param([Parameter(Mandatory)][string]$Uri)

    $items = @()
    $nextUri = $Uri

    do {
        $response = Invoke-AutopatchGraphRequest -Method GET -Uri $nextUri

        if ($null -eq $response) {
            break
        }

        $hasValueProperty = $false
        $hasNextLinkProperty = $false
        $value = $null
        $nextLink = $null

        if ($response -is [System.Collections.IDictionary]) {
            if ($response.Contains('value')) {
                $hasValueProperty = $true
                $value = $response['value']
            }
            if ($response.Contains('@odata.nextLink')) {
                $hasNextLinkProperty = $true
                $nextLink = $response['@odata.nextLink']
            }
        }
        else {
            $propertyNames = @($response.PSObject.Properties.Name)
            if ($propertyNames -contains 'value') {
                $hasValueProperty = $true
                $value = $response.value
            }
            if ($propertyNames -contains '@odata.nextLink') {
                $hasNextLinkProperty = $true
                $nextLink = $response.'@odata.nextLink'
            }
        }

        if ($hasValueProperty -and $null -ne $value) {
            foreach ($item in @($value)) {
                $items += $item
            }
        }
        elseif ($response -is [System.Array]) {
            foreach ($item in $response) {
                $items += $item
            }
        }
        else {
            $items += $response
        }

        if ($hasNextLinkProperty -and -not [string]::IsNullOrWhiteSpace([string]$nextLink)) {
            $nextUri = [string]$nextLink
        }
        else {
            $nextUri = $null
        }
    }
    while ($nextUri)

    return ,$items
}

function Start-ExportJob {
    param(
        [Parameter(Mandatory)][string]$ReportName,
        [Parameter()][string[]]$Select,
        [Parameter()][string]$Filter
    )

    $body = @{ reportName = $ReportName; format = 'csv'; localizationType = 'localizedValuesAsAdditionalColumn' }
    if ($Select -and $Select.Count -gt 0) { $body['select'] = $Select }
    if ($Filter) { $body['filter'] = $Filter }

    Write-Log -Message ("Starting export job for report [{0}] with filter [{1}]" -f $ReportName, $(if ([string]::IsNullOrWhiteSpace($Filter)) { '<none>' } else { $Filter }))
    $job = Invoke-AutopatchGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/reports/exportJobs' -Body ($body | ConvertTo-Json -Depth 6)
    return $job
}

function Wait-ExportJob {
    param(
        [Parameter(Mandatory)][string]$JobId,
        [Parameter()][int]$TimeoutSeconds = 300,
        [Parameter()][int]$PollSeconds = 5
    )

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    do {
        $job = Invoke-AutopatchGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/deviceManagement/reports/exportJobs/{0}" -f $JobId)
        Write-Log -Message ("Export job [{0}] status: {1}" -f $JobId, $job.status)
        if ($job.status -eq 'completed') { return $job }
        if ($job.status -eq 'failed') { throw "Export job failed: $JobId" }
        Start-Sleep -Seconds $PollSeconds
    } while ((Get-Date) -lt $deadline)

    throw "Export job timed out after $TimeoutSeconds seconds: $JobId"
}

function Import-ExportedCsv {
    param(
        [Parameter(Mandatory)][string]$ReportName,
        [Parameter()][string[]]$Select,
        [Parameter()][string]$Filter
    )

    $job = Start-ExportJob -ReportName $ReportName -Select $Select -Filter $Filter
    $completedJob = Wait-ExportJob -JobId $job.id
    $tempZip = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("{0}_{1}.zip" -f $ReportName, $job.id)
    $tempFolder = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath ("{0}_{1}" -f $ReportName, $job.id)

    try {
        Invoke-WebRequest -Uri $completedJob.url -OutFile $tempZip
        Ensure-Folder -Path $tempFolder
        Expand-Archive -Path $tempZip -DestinationPath $tempFolder -Force
        $csvFile = Get-ChildItem -Path $tempFolder -Filter '*.csv' -File | Select-Object -First 1
        if (-not $csvFile) { throw "No CSV file found in exported package for report [$ReportName]." }
        Write-Log -Message ("Imported CSV for report [{0}] from [{1}]" -f $ReportName, $csvFile.FullName)
        return @(Import-Csv -Path $csvFile.FullName)
    }
    finally {
        if (Test-Path -Path $tempZip) { Remove-Item -Path $tempZip -Force -ErrorAction SilentlyContinue }
        if (Test-Path -Path $tempFolder) { Remove-Item -Path $tempFolder -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

function Get-FeatureUpdatePolicyMap {
    $uri = 'https://graph.microsoft.com/beta/deviceManagement/windowsFeatureUpdateProfiles?$select=id,displayName'
    $items = Invoke-GraphGetAll -Uri $uri
    $map = @{}
    foreach ($item in $items) {
        $map[$item.id] = $item.displayName
    }
    return $map
}

function Get-QualityUpdatePolicyMap {
    $uri = 'https://graph.microsoft.com/beta/deviceManagement/windowsQualityUpdateProfiles?$select=id,displayName'
    $items = Invoke-GraphGetAll -Uri $uri
    $map = @{}
    foreach ($item in $items) {
        $map[$item.id] = $item.displayName
    }
    return $map
}

function Convert-FeatureRowsToAlertDetails {
    param(
        [Parameter(Mandatory)][object[]]$Rows,
        [Parameter(Mandatory)][hashtable]$PolicyMap
    )

    $details = foreach ($row in $Rows) {
        if ([string]::IsNullOrWhiteSpace($row.LatestAlertMessage)) { continue }
        [pscustomobject]@{
            AlertName = $row.LatestAlertMessage
            Severity = 'Unknown'
            Category = 'Device'
            AffectedUpdateType = 'Feature'
            DeviceName = $row.DeviceName
            DeviceId = $row.DeviceId
            PolicyId = $row.PolicyId
            PolicyName = $(if ($PolicyMap.ContainsKey($row.PolicyId)) { $PolicyMap[$row.PolicyId] } else { $null })
            EventDateUtc = $row.EventDateTimeUTC
            LastScanUtc = $row.LastWUScanTimeUTC
            AggregateState = $row.AggregateState
            CurrentStatus = $row.CurrentDeviceUpdateStatus
            SourceReport = 'FeatureUpdateDeviceState'
        }
    }
    return @($details)
}

function Convert-QualityRowsToAlertDetails {
    param(
        [Parameter(Mandatory)][object[]]$Rows,
        [Parameter(Mandatory)][hashtable]$PolicyMap
    )

    $details = foreach ($row in $Rows) {
        if ([string]::IsNullOrWhiteSpace($row.LatestAlertMessage)) { continue }
        [pscustomobject]@{
            AlertName = $row.LatestAlertMessage
            Severity = 'Unknown'
            Category = 'Device'
            AffectedUpdateType = 'Quality'
            DeviceName = $row.DeviceName
            DeviceId = $row.DeviceId
            PolicyId = $row.PolicyId
            PolicyName = $(if ($PolicyMap.ContainsKey($row.PolicyId)) { $PolicyMap[$row.PolicyId] } else { $null })
            EventDateUtc = $row.EventDateTimeUTC
            LastScanUtc = $row.LastWUScanTimeUTC
            AggregateState = $row.AggregateState
            CurrentStatus = $row.CurrentDeviceUpdateStatus
            SourceReport = 'QualityUpdateDeviceStatusByPolicy'
        }
    }
    return @($details)
}

function Convert-QualityErrorRowsToAlertDetails {
    param(
        [Parameter(Mandatory)][object[]]$Rows,
        [Parameter(Mandatory)][hashtable]$PolicyMap
    )

    $details = foreach ($row in $Rows) {
        if ([string]::IsNullOrWhiteSpace($row.AlertMessage)) { continue }
        [pscustomobject]@{
            AlertName = $row.AlertMessage
            Severity = 'Critical'
            Category = 'Device'
            AffectedUpdateType = 'Quality'
            DeviceName = $row.DeviceName
            DeviceId = $row.DeviceId
            PolicyId = $row.PolicyId
            PolicyName = $(if ($PolicyMap.ContainsKey($row.PolicyId)) { $PolicyMap[$row.PolicyId] } else { $null })
            EventDateUtc = $row.ExpediteQUReleaseDate
            LastScanUtc = $null
            AggregateState = 'Error'
            CurrentStatus = 'Error'
            SourceReport = 'QualityUpdateDeviceErrorsByPolicy'
        }
    }
    return @($details)
}

function Group-AlertSummary {
    param([Parameter(Mandatory)][object[]]$Details)

    $summary = $Details | Group-Object -Property AlertName,Severity,Category,AffectedUpdateType,SourceReport | ForEach-Object {
        $first = $_.Group | Sort-Object -Property EventDateUtc | Select-Object -First 1
        $last = $_.Group | Sort-Object -Property EventDateUtc | Select-Object -Last 1
        [pscustomobject]@{
            AlertName = $first.AlertName
            Severity = $first.Severity
            Category = $first.Category
            AffectedUpdateType = $first.AffectedUpdateType
            Impact = $_.Count
            FirstSeenUtc = $first.EventDateUtc
            LastSeenUtc = $last.EventDateUtc
            SourceReport = $first.SourceReport
        }
    }

    return @($summary | Sort-Object -Property @{ Expression = 'Impact'; Descending = $true }, AlertName)
}

$script:CompletionStatus = 'Success'
$script:CompletionError = $null
$script:TranscriptStarted = $false
try {
    Ensure-Folder -Path $OutputFolder
    Ensure-Folder -Path $LogFolder
    try { Start-Transcript -Path $TranscriptFile -Force -ErrorAction Stop | Out-Null; $script:TranscriptStarted = $true } catch { Write-Log -Message ("Transcript start failed: {0}" -f $_.Exception.Message) -Level WARN }
    Test-Ps7
    Import-RequiredModule -Name Microsoft.Graph.Authentication

    Write-Log -Message ("Starting {0} v{1}" -f $ScriptName, $ScriptVersion)
    Write-Log -Message ("Using output folder [{0}]" -f $OutputFolder)
    Connect-GraphSession
    Import-SmartM365CorePreflight
    Invoke-CoreSmartM365Preflight -ScriptName $ScriptName -RequiredModules @('Microsoft.Graph.Authentication') -OutputPaths @($OutputFolder, $LogFolder) -RequiredGraphApplicationPermissions @('DeviceManagementConfiguration.Read.All','DeviceManagementManagedDevices.Read.All','DeviceManagementApps.Read.All') -GraphProbeUris @(
        'https://graph.microsoft.com/beta/deviceManagement/windowsFeatureUpdateProfiles?$top=1'
    ) | Out-Null


    $detailRows = New-Object System.Collections.Generic.List[object]

    if ($IncludeFeatureUpdates) {
        Write-Log -Message 'Collecting feature update report data.'
        $featurePolicyMap = Get-FeatureUpdatePolicyMap
        $featureSummaryRows = Import-ExportedCsv -ReportName 'FeatureUpdatePolicyStatusSummary' -Select @('PolicyId','PolicyName','FeatureUpdateVersion','CountDevicesErrorStatus','CountDevicesInProgressStatus','CountDevicesSuccessStatus')
        foreach ($policyRow in $featureSummaryRows) {
            if ([string]::IsNullOrWhiteSpace($policyRow.PolicyId)) { continue }
            $filter = "PolicyId eq '{0}'" -f $policyRow.PolicyId
            $featureRows = Import-ExportedCsv -ReportName 'FeatureUpdateDeviceState' -Select @('DeviceId','DeviceName','PolicyId','EventDateTimeUTC','LastWUScanTimeUTC','AggregateState','CurrentDeviceUpdateStatus','LatestAlertMessage') -Filter $filter
            foreach ($item in (Convert-FeatureRowsToAlertDetails -Rows $featureRows -PolicyMap $featurePolicyMap)) { $detailRows.Add($item) }
        }
    }

    if ($IncludeQualityUpdates) {
        Write-Log -Message 'Collecting quality update report data.'
        $qualityPolicyMap = Get-QualityUpdatePolicyMap
        $qualitySummaryRows = Import-ExportedCsv -ReportName 'QualityUpdatePolicyStatusSummary' -Select @('PolicyId','PolicyName','ExpediteQUReleaseDate','CountDevicesErrorStatus','CountDevicesInProgressStatus','CountDevicesSuccessStatus')
        $qualityPolicyRows = New-Object System.Collections.Generic.List[object]
        foreach ($policyRow in $qualitySummaryRows) {
            if ([string]::IsNullOrWhiteSpace($policyRow.PolicyId)) { continue }
            $qualityPolicyRows.Add([pscustomobject]@{
                PolicyId = $policyRow.PolicyId
                PolicyName = $policyRow.PolicyName
                ExpediteQUReleaseDate = $policyRow.ExpediteQUReleaseDate
                CountDevicesErrorStatus = $policyRow.CountDevicesErrorStatus
                CountDevicesInProgressStatus = $policyRow.CountDevicesInProgressStatus
                CountDevicesSuccessStatus = $policyRow.CountDevicesSuccessStatus
            })

            $filter = "PolicyId eq '{0}'" -f $policyRow.PolicyId
            $qualityRows = Import-ExportedCsv -ReportName 'QualityUpdateDeviceStatusByPolicy' -Select @('DeviceId','DeviceName','PolicyId','EventDateTimeUTC','LastWUScanTimeUTC','AggregateState','CurrentDeviceUpdateStatus','LatestAlertMessage') -Filter $filter
            foreach ($item in (Convert-QualityRowsToAlertDetails -Rows $qualityRows -PolicyMap $qualityPolicyMap)) { $detailRows.Add($item) }

            $qualityErrorRows = Import-ExportedCsv -ReportName 'QualityUpdateDeviceErrorsByPolicy' -Select @('DeviceId','DeviceName','PolicyId','ExpediteQUReleaseDate','AlertMessage','Win32ErrorCode') -Filter $filter
            foreach ($item in (Convert-QualityErrorRowsToAlertDetails -Rows $qualityErrorRows -PolicyMap $qualityPolicyMap)) { $detailRows.Add($item) }
        }
        Publish-CoreSmartM365Csv -Data $qualityPolicyRows.ToArray() -TimestampedPath $PolicyCsvPath -LatestPath $PolicyLatestCsvPath -Columns $PolicyColumns | Out-Null
        Write-Log -Message ("Policy summary CSV written to [{0}]" -f $PolicyCsvPath)
    }

    $detailOutput = @($detailRows.ToArray() | Sort-Object -Property @{ Expression = 'EventDateUtc'; Descending = $true }, DeviceName)
    if ($MaxItems -gt 0 -and $detailOutput.Count -gt $MaxItems) {
        Write-Log -Message ("MaxItems enabled: restricted Autopatch alert details from {0} to {1} before summary calculation." -f $detailOutput.Count, $MaxItems) -Level WARN
        $detailOutput = @($detailOutput | Select-Object -First $MaxItems)
    }
    $summaryOutput = Group-AlertSummary -Details $detailOutput

    Publish-CoreSmartM365Csv -Data $detailOutput -TimestampedPath $DetailCsvPath -LatestPath $DetailLatestCsvPath -Columns $DetailColumns | Out-Null
    Publish-CoreSmartM365Csv -Data $summaryOutput -TimestampedPath $SummaryCsvPath -LatestPath $SummaryLatestCsvPath -Columns $SummaryColumns | Out-Null

    Write-Log -Message ("Detail CSV written to [{0}]" -f $DetailCsvPath)
    Write-Log -Message ("Summary CSV written to [{0}]" -f $SummaryCsvPath)
    Write-Log -Message ("Completed successfully. Summary rows: {0}. Detail rows: {1}." -f $summaryOutput.Count, $detailOutput.Count)
}
catch {
    $script:CompletionStatus = 'Failed'
    $script:CompletionError = $_
    Write-Log -Message $_.Exception.Message -Level ERROR
    throw
}
finally {
    try {
        Disconnect-MgGraph | Out-Null
    }
    catch {
    }
    $duration = (Get-Date) - $StartTime
    Write-Log -Message ("Finished in {0:c}" -f $duration)
    if ($script:TranscriptStarted) { try { Stop-Transcript | Out-Null } catch {} }
    if (Get-Command Complete-CoreSmartM365ExecutionContext -ErrorAction SilentlyContinue) {
        Complete-CoreSmartM365ExecutionContext -Status $script:CompletionStatus -ErrorRecord $script:CompletionError
    }
}
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBSi5zx2cxBWqrl
# wFav9wTWWiimHeQ3wDYrzsnXXsU/AqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEICUPHn1O6erC+vquZ3DdE2HvlH5JGr9OyT/NHaCQgR06MA0GCSqG
# SIb3DQEBAQUABIIBgCcVW0qo9WXRnkjGEdTsfnnkef3OgVq1DczcztsstgobMEaL
# RpbzacplFUq9O9RgH6aoOIhqFk9yVuT9M9DCmsDi2pT5w5rVNsJvGaGoTL5K9Dwg
# AF2TWHfGDuqwENMnM9GCp4YsAfHPCaFdOFmDD7cc/EU4TPx+Ns9Ev6cExsEfoMeB
# DOOz+OUn3qBBwC3mnp9o2Skatx86yiV/TQvN7H1O4o8y1TKl0bTZTFrZdNcnR0gR
# XOXZJl9Y8JWZj6SnXyG67MXe7OvO3hQvUkbuUPcOkn35rXojnBdjvX4Lg2b5kbyu
# H9MdymRtyB9qvP72pNxN8o/izLOSQn1aTsEmS1Y8Z1kLMqpv2xbMdP7a9VTLFh6T
# pjX8QHU5km0hyWozCOkc7SA97iwvebwnvgSJGuLZieuAG1IyK4A+0sj7ewJxZnhS
# AFT691MKkjUy3YW/VT3Kr1gbdux84SZrRucXvnPx0SrQo+vSMnzx2e9uByslviKa
# rzmpRQH6QBnLMBco+qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjEwODM4
# MDJaMC8GCSqGSIb3DQEJBDEiBCCe2xWTjafNYygmh+kpjqEwwnIzdVTtb6/kwvXj
# FE4OsDANBgkqhkiG9w0BAQEFAASCAgCP9x0vm96eUxE5UUjFfJApi9bj8eDOMDKQ
# oQDEIYGByvsS537YN6Bp3Jbao2JfVqMeyLu+o/ik1LIKpmgSyPyOxuX6N40pfCKR
# fz79EDGyiQTtNdr0m15ZIEXwTMvdVnwpoPIWLsu7X/pCikO6TI0mF56OHlc31vEw
# eEDYAxMUNaurQcQbzBzwLuO6joh2Lez8HPRP+dC4tlEr50YcyN6dcXg00MupHeiO
# nTWa2ibkbgm/HRD2C0apeC5mYkULUQ1q4QUFotAvXXrpmQfHAR5SX2ndPZlvpBUv
# k++EQxw8vkp+glWGq4NueyD+i5QtGj7FCt6r9Os+hHtJaeN6rBDvLdKWE0FLrahv
# 5zOrOhr4EGXHqTEmOVVh39JJ2MAq6VsGXe66fKvC5Td1C/wE3Ul6k3TA3RIYZ+zB
# 7VqKxKMqXZkfcyEXhhHa1o3fms+1VJdCxu4SYTvJY5e/c0PF/M74+JJHxQ8YdiTo
# lY0iMWq64DCLCsUlBSFMR5N7VvdXWVDXt3ZzyqGzH4Aav7uEstC5G5VHMWSiDC7x
# 5mDeuV/T0a6gWdyhXOif9DFQZEclqqXgzrg+PXJJ5ApvAabqo6HyAqL1fN4OwbM8
# 8aHl7zTQmCQkX2QEoipDxO+eLJswE5sa1zT6RcRtXKaO07PFN55DJoDcJl0pmz3V
# ILjVSFviVQ==
# SIG # End signature block
