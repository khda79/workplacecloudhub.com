#Requires -Version 7.0
<#
.SYNOPSIS
    Azure AD Connect synchronization health inventory and freshness alert.

.DESCRIPTION
    Reads tenant organization synchronization state from Microsoft Graph, exports health checks to CSV,
    and sends an Alerts notification when Entra ID is not synchronized or when the last sync is older than
    the configured threshold. This protects downstream SmartInventory data consumers from stale hybrid data.

.PARAMETER Tenant
    Tenant profile key to load from Config/Tenants. Defaults to test.

.PARAMETER SyncAgeThresholdMinutes
    Maximum acceptable age in minutes for on-premises directory synchronization. Default: 180.

.PARAMETER Connect
    Kept for launcher consistency. The script always disconnects any existing Microsoft Graph session before connecting.

.PARAMETER InteractiveAuth
    Uses delegated interactive Microsoft Graph authentication for troubleshooting.

.PARAMETER OutputPath
    Optional output directory override. If omitted, ScriptCsvLogFolderPath from local JSON is used.

.VERSION
1.1

.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Required application permission: Directory.Read.All
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [int]$SyncAgeThresholdMinutes = 180,
    [switch]$Connect,
    [switch]$InteractiveAuth,
    [string]$OutputPath
)

$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidates = @(
            (Join-Path -Path $d -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $d -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )
        foreach ($p in $candidates) { if (Test-Path -LiteralPath $p) { return $p } }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $tenantContextPath
Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot | Out-Null

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$MaximumFunctionCount = 32768
$ScriptVersion = '1.1'
$TaskName = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion"
$CurrentOperation = 'Initialize'
$script:SmartM365GlobalConfig = $null

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host 'This script requires PowerShell 7 or later.' -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    exit 1
}

function Import-SmartM365CoreModule {
    [CmdletBinding()]
    param()

    $searchRoot = $PSScriptRoot
    while ($searchRoot) {
        $candidate = Join-Path -Path $searchRoot -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
        if (Test-Path -LiteralPath $candidate) {
            Import-Module $candidate -Force -ErrorAction Stop
            return
        }
        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }
    throw 'SmartM365.Core module manifest not found.'
}

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
            if (-not (Test-Path -LiteralPath $templatePath)) { throw "Missing local config and template: $configPath" }
            Copy-Item -LiteralPath $templatePath -Destination $configPath -ErrorAction Stop
        }
    }
    return Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
}

function Get-SmartM365GlobalConfig {
    [CmdletBinding()]
    param()

    if ($null -ne $script:SmartM365GlobalConfig) { return $script:SmartM365GlobalConfig }
    $script:SmartM365GlobalConfig = [pscustomobject]@{}
    $searchRoot = $PSScriptRoot
    while ($searchRoot) {
        $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'Config\SmartM365.global.local.json'
        if (Test-Path -LiteralPath $globalConfigPath) {
            $script:SmartM365GlobalConfig = Get-Content -LiteralPath $globalConfigPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
            break
        }
        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }
    return $script:SmartM365GlobalConfig
}

function Resolve-SmartM365ConfigValue {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value)) { return $Value }
    if ($Value -notmatch '\{\{[^}]+\}\}') { return $Value }
    $globalConfig = Get-SmartM365GlobalConfig
    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $tokenMatches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($tokenMatches.Count -eq 0) { break }
        $changed = $false
        foreach ($match in $tokenMatches) {
            $tokenProperty = $globalConfig.PSObject.Properties[$match.Groups['Name'].Value]
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
    param([Parameter(Mandatory)]$Config, [Parameter(Mandatory)][string]$Name, $DefaultValue)

    $property = $Config.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        if ($property.Value -is [string]) {
            $localValue = $property.Value.Trim()
            if ($localValue -and $localValue -notin @('__USE_GLOBAL__', 'USE_GLOBAL')) { return Resolve-SmartM365ConfigValue -Value $property.Value }
        }
        else { return Resolve-SmartM365ConfigValue -Value $property.Value }
    }

    $globalConfig = Get-SmartM365GlobalConfig
    $globalProperty = $globalConfig.PSObject.Properties[$Name]
    if ($null -ne $globalProperty -and $null -ne $globalProperty.Value) {
        if ($globalProperty.Value -is [string] -and [string]::IsNullOrWhiteSpace($globalProperty.Value)) { return $DefaultValue }
        return Resolve-SmartM365ConfigValue -Value $globalProperty.Value
    }
    return $DefaultValue
}

function Ensure-GraphModule {
    [CmdletBinding()]
    param()

    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
}

function Disconnect-GraphSafe {
    [CmdletBinding()]
    param()

    try {
        if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
            $context = Get-MgContext -ErrorAction SilentlyContinue
            if ($context) {
                Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
                WriteLog -Message 'Disconnected from Microsoft Graph.' -Level SUCCESS
            }
        }
    }
    catch { WriteLog -Message ("Disconnect-MgGraph failed (non-fatal): {0}" -f $_.Exception.Message) -Level WARNING }
}

function Connect-GraphForSyncHealth {
    [CmdletBinding()]
    param(
        [switch]$UseInteractiveAuth,
        [string]$AppId,
        [string]$TenantId,
        [string]$Thumbprint
    )

    Disconnect-GraphSafe
    if ($UseInteractiveAuth) {
        WriteLog -Message 'Connecting to Microsoft Graph using interactive authentication.' -Level INFO
        Connect-MgGraph -Scopes @('Directory.Read.All') -NoWelcome -ErrorAction Stop | Out-Null
    }
    else {
        WriteLog -Message 'Connecting to Microsoft Graph using app-only certificate authentication.' -Level INFO
        Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome -ErrorAction Stop | Out-Null
    }
    WriteLog -Message 'Connected to Microsoft Graph.' -Level SUCCESS
}

function Send-SyncHealthTeamsNotification {
    [CmdletBinding()]
    param(
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR')][string]$Level,
        [string]$Title,
        [string]$Message,
        [hashtable]$Facts
    )

    if (Get-Command Send-SmartM365TeamsNotification -ErrorAction SilentlyContinue) {
        Send-SmartM365TeamsNotification -Title $Title -Message $Message -Level $Level -Facts $Facts | Out-Null
    }
}

Import-SmartM365CoreModule
$ScriptLocalConfig = Get-ScriptLocalConfig

$global:RetentionMaxCSV = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxLogs' -DefaultValue 30)
$AppId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AppId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$TenantId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$Thumb = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumb' -DefaultValue '0000000000000000000000000000000000000000'
$OrgDomain = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue 'contoso.onmicrosoft.com'
$ScriptCsvLogFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ScriptCsvLogFolderPath' -DefaultValue ''
$LatestCsvFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''
$LogAllRootPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LogAllRootPath' -DefaultValue ''
$ConfiguredThreshold = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SyncAgeThresholdMinutes' -DefaultValue $SyncAgeThresholdMinutes)
if ($PSBoundParameters.ContainsKey('SyncAgeThresholdMinutes') -eq $false) { $SyncAgeThresholdMinutes = $ConfiguredThreshold }
$WeeklyHistoryFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'WeeklyHistoryFolderPath' -DefaultValue ''

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) { $ScriptCsvLogFolderPath = $OutputPath }
if ([string]::IsNullOrWhiteSpace($ScriptCsvLogFolderPath)) { $ScriptCsvLogFolderPath = Join-Path -Path $PSScriptRoot -ChildPath 'Output' }
if ([string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) { $LatestCsvFolderPath = $ScriptCsvLogFolderPath }

Ensure-GraphModule

$runId = [guid]::NewGuid().ToString()
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvName = 'M365_Entra_AzureADConnect_SyncHealth.csv'
$timestampedCsvPath = Join-Path -Path $ScriptCsvLogFolderPath -ChildPath ("M365_Entra_AzureADConnect_SyncHealth_{0}.csv" -f $timestamp)
$latestCsvPath = Join-Path -Path $LatestCsvFolderPath -ChildPath $csvName
$logRoot = if ([string]::IsNullOrWhiteSpace($LogAllRootPath)) { Join-Path $ScriptCsvLogFolderPath 'Logs' } else { Join-Path $LogAllRootPath 'AzureADConnect-SyncHealth' }
$logPath = Join-Path -Path $logRoot -ChildPath ("SmartM365-AzureADConnect-SyncHealth-Inventory_{0}.log" -f $timestamp)
foreach ($folder in @($ScriptCsvLogFolderPath, $LatestCsvFolderPath, $logRoot)) { if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null } }
Set-SmartM365CoreContext -RunId $runId -RunOutputRoot $ScriptCsvLogFolderPath -LatestOutputRoot $LatestCsvFolderPath -LogPath $logPath
$global:LogTextFile = $logPath

try {
    $CurrentOperation = 'InitializeScriptEnvironment'
    $initializedOutput = InitializeScriptEnvironment -OutputPath $ScriptCsvLogFolderPath -LogFileName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))
    $ScriptCsvLogFolderPath = $initializedOutput
    if ([string]::IsNullOrWhiteSpace($WeeklyHistoryFolderPath)) { $WeeklyHistoryFolderPath = Join-Path -Path $ScriptCsvLogFolderPath -ChildPath 'WeeklyHistory' }
    $timestampedCsvPath = Join-Path -Path $ScriptCsvLogFolderPath -ChildPath ("M365_Entra_AzureADConnect_SyncHealth_{0}.csv" -f $timestamp)
    $latestCsvPath = Join-Path -Path $LatestCsvFolderPath -ChildPath $csvName
    Start-Transcript -Path $global:logTranscriptFile -Append | Out-Null
    WriteLog -Message "Starting $TaskName. Threshold=$SyncAgeThresholdMinutes minute(s)." -Level INFO
    WriteLog -Message "Default WeeklyHistoryFolderPath: $WeeklyHistoryFolderPath" -Level INFO
    if ($Connect) { WriteLog -Message 'Connect switch specified; Graph connection will be established by this script.' -Level INFO }

    $CurrentOperation = 'ConnectGraph'
    Connect-GraphForSyncHealth -UseInteractiveAuth:$InteractiveAuth -AppId $AppId -TenantId $TenantId -Thumbprint $Thumb

    $CurrentOperation = 'Preflight'
    Invoke-SmartM365Preflight -ScriptName $TaskName -RequiredModules @('Microsoft.Graph.Authentication') -OutputPaths @($ScriptCsvLogFolderPath, $LatestCsvFolderPath) -GraphProbeUris @('https://graph.microsoft.com/v1.0/organization?$select=id,displayName,onPremisesSyncEnabled,onPremisesLastSyncDateTime') | Out-Null

    $CurrentOperation = 'ReadOrganizationSyncState'
    $orgResponse = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id,displayName,onPremisesSyncEnabled,onPremisesLastSyncDateTime' -OutputType PSObject -ErrorAction Stop
    $org = @($orgResponse.value) | Select-Object -First 1
    if (-not $org) { throw 'Graph /organization returned no tenant organization object.' }

    $syncEnabled = [bool]$org.onPremisesSyncEnabled
    $lastSyncText = [string]$org.onPremisesLastSyncDateTime
    $syncAgeMinutes = $null
    if (-not [string]::IsNullOrWhiteSpace($lastSyncText)) {
        $lastSyncUtc = [datetime]::Parse($lastSyncText, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::AssumeUniversal).ToUniversalTime()
        $syncAgeMinutes = [math]::Round(((Get-Date).ToUniversalTime() - $lastSyncUtc).TotalMinutes, 1)
    }

    $exportDateTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $rows = @(
        [pscustomobject][ordered]@{
            CheckName = 'SyncEnabled'
            Status = if ($syncEnabled) { 'OK' } else { 'ERROR' }
            Value = [string]$syncEnabled
            Threshold = 'true'
            Detail = if ($syncEnabled) { 'On-premises directory synchronization is enabled.' } else { 'On-premises directory synchronization is disabled. Hybrid SmartInventory data may be stale or incomplete.' }
            OrganizationId = [string]$org.id
            OrganizationName = [string]$org.displayName
            LastSyncDateTimeUtc = $lastSyncText
            SyncAgeMinutes = $syncAgeMinutes
            ExportDateTime = $exportDateTime
            RunId = $runId
        },
        [pscustomobject][ordered]@{
            CheckName = 'LastSyncAge'
            Status = if (-not $syncEnabled) { 'ERROR' } elseif ($null -eq $syncAgeMinutes) { 'WARNING' } elseif ($syncAgeMinutes -le $SyncAgeThresholdMinutes) { 'OK' } else { 'WARNING' }
            Value = if ($null -eq $syncAgeMinutes) { '' } else { "$syncAgeMinutes min" }
            Threshold = "$SyncAgeThresholdMinutes min"
            Detail = if (-not $syncEnabled) { 'Sync is disabled; last sync freshness cannot be trusted.' } elseif ($null -eq $syncAgeMinutes) { 'Last sync timestamp is empty.' } elseif ($syncAgeMinutes -le $SyncAgeThresholdMinutes) { "Last sync age is within threshold: $syncAgeMinutes minute(s)." } else { "Last sync is stale: $syncAgeMinutes minute(s), threshold is $SyncAgeThresholdMinutes minute(s). SmartInventory data may not be up to date." }
            OrganizationId = [string]$org.id
            OrganizationName = [string]$org.displayName
            LastSyncDateTimeUtc = $lastSyncText
            SyncAgeMinutes = $syncAgeMinutes
            ExportDateTime = $exportDateTime
            RunId = $runId
        }
    )

    $CurrentOperation = 'ExportCsv'
    Export-SmartM365Csv -Data $rows -TimestampedPath $timestampedCsvPath -LatestPath $latestCsvPath | Out-Null

    $overallStatus = if (@($rows | Where-Object Status -eq 'ERROR').Count -gt 0) { 'ERROR' } elseif (@($rows | Where-Object Status -eq 'WARNING').Count -gt 0) { 'WARNING' } else { 'OK' }
    $message = "Azure AD Connect sync health: $overallStatus. SyncEnabled=$syncEnabled; LastSyncAgeMinutes=$syncAgeMinutes; Threshold=$SyncAgeThresholdMinutes."
    WriteLog -Message $message -Level $(if ($overallStatus -eq 'OK') { 'SUCCESS' } else { 'WARNING' })

    $facts = @{
        Script = $MyInvocation.MyCommand.Name
        TenantOrOrganization = [string]$org.displayName
        OverallStatus = $overallStatus
        SyncEnabled = [string]$syncEnabled
        LastSyncDateTimeUtc = $lastSyncText
        SyncAgeMinutes = [string]$syncAgeMinutes
        ThresholdMinutes = [string]$SyncAgeThresholdMinutes
        LatestCsvPath = $latestCsvPath
        TimestampedCsvPath = $timestampedCsvPath
        WeeklyHistoryFolderPath = $WeeklyHistoryFolderPath
        LogFile = $global:LogTextFile
    }
    if ($overallStatus -eq 'OK') {
        Send-SyncHealthTeamsNotification -Level SUCCESS -Title 'Azure AD Connect sync health success' -Message $message -Facts $facts
    }
    else {
        Send-SyncHealthTeamsNotification -Level WARNING -Title 'Azure AD Connect sync health warning' -Message $message -Facts $facts
    }

    try { Complete-SmartM365ExecutionContext -Status Auto } catch {}
    Disconnect-GraphSafe
}
catch {
    $globalError = $_
    WriteLog -Message ("Global error during {0}: {1}" -f $CurrentOperation, $globalError.Exception.Message) -Level ERROR
    Send-SyncHealthTeamsNotification -Level ERROR -Title 'Azure AD Connect sync health failed' -Message $globalError.Exception.Message -Facts @{
        Script = $MyInvocation.MyCommand.Name
        TenantOrOrganization = $OrgDomain
        Operation = $CurrentOperation
        LogFile = $global:LogTextFile
    }
    try { Complete-SmartM365ExecutionContext -Status Failed -FailureStage $CurrentOperation } catch {}
    Disconnect-GraphSafe
    throw
}
finally {
    try { RemoveOldFiles -Path $ScriptCsvLogFolderPath -Filter '*.csv' -KeepCount $global:RetentionMaxCSV -LogFile $global:LogTextFile } catch {}
    try { Stop-SmartM365TranscriptSafely } catch {}
}