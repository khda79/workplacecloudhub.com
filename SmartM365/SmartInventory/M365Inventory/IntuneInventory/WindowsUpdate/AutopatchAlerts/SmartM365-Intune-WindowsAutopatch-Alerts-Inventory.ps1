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
1.10



.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication.
    Minimum Graph application permissions: DeviceManagementConfiguration.Read.All; DeviceManagementManagedDevices.Read.All; DeviceManagementApps.Read.All.
    Conditional: Mail.Send is required only when Graph mail is used; Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
Author    : https://github.com/khda79/workplacecloudhub.com
    Version : 1.11
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
            Import-Module $modulePath -Prefix Core -ErrorAction Stop
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
$ScriptVersion = "1.11"
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

function Invoke-GraphGetAll {
    param([Parameter(Mandatory)][string]$Uri)

    $items = @()
    $nextUri = $Uri

    do {
        $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType PSObject

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
    $job = Invoke-MgGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/deviceManagement/reports/exportJobs' -Body ($body | ConvertTo-Json -Depth 6) -ContentType 'application/json' -OutputType PSObject
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
        $job = Invoke-MgGraphRequest -Method GET -Uri ("https://graph.microsoft.com/v1.0/deviceManagement/reports/exportJobs/{0}" -f $JobId) -OutputType PSObject
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
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAvve9D7n6FLRjy
# 0VfupiQnq2ZISAzEurZDusxSjFm74KCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCDzn1a3QEnBSoVs/lyf
# UlNf4rzLRz4ItWMIsiM/cXsbeTANBgkqhkiG9w0BAQEFAASCAYCmFCaWLOoXbo/I
# hWnNLqPR0LNvQqkRPP5Q1BL2EG+a9D7qC86BYRJ7bWL5eRhvQRw4F6pYHITiCvIn
# 9rbl2m12kKPZZIJYOseRuXSRPHfKUAXM63P+jvQ7sr6pOvqwrLzzkB5ZnJ9uJXG2
# HD4FzzLxHJ1H0QTAV7O6E775NHubnzcuQa1gGUbqxa7H1rNOINocimN04CnmfSwS
# LdM3KVwBAdmdjoXsCBhVJrN11TzzCeNXLsCY7zUMixtCdiUn0iIdtaP8cf89Lhjg
# sbkAPzdUcKYbRKuSK2zjBWr02Y72Nygc8hgmLkOj2Yo8MG0oAtSKKY8DsbHRMW9V
# 5smmKWR2xsMu8WUuWwrhsj5nSy8RyYICo8mDt5vEvIaoc7vFTsTclGULIVKgXH7G
# 74teQR0LHfjpimrGKOETZJE+KrmbQHjsD16W2wzUD3KM1G5c0hXI34RVVSE07mcX
# SPYyYdkUUfFoTdNB8/Hrdu+CcMJTcPWC1zT+ybXpUnbnQzkvPsc=
# SIG # End signature block
