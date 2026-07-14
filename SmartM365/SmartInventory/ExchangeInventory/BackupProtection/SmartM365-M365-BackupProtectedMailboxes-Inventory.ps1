#Requires -Version 7.0
<#
.SYNOPSIS
    Microsoft 365 Backup protected mailbox inventory.

.DESCRIPTION
    Retrieves Microsoft 365 Backup mailbox protection units from Microsoft Graph Backup Restore APIs,
    keeps read-only protected mailbox evidence, exports timestamped and latest CSV files, and publishes
    the latest CSV through the shared SmartM365 SharePoint helper when enabled.

.PARAMETER Tenant
    Tenant profile key to load from Config/Tenants. Defaults to test.

.PARAMETER Connect
    Kept for launcher consistency. The script always disconnects any existing Microsoft Graph session before connecting.

.PARAMETER InteractiveAuth
    Uses delegated interactive Microsoft Graph authentication. Enabled by default because Microsoft 365
    Backup protection-unit inventory requires a delegated session in this SmartM365 workflow. Pass
    -InteractiveAuth:$false only to explicitly request app-only certificate authentication.

.PARAMETER IncludeNonProtected
    Include non-protected mailbox protection units in the CSV. By default only Status = protected is exported.

.PARAMETER OutputPath
    Optional output directory override. If omitted, ScriptCsvLogFolderPath from local JSON is used.

.PARAMETER MaxPages
    Maximum number of Graph pages to retrieve. Acts as a safety guard against endless pagination and
    as a quick smoke test limiter (e.g. -MaxPages 3). When the limit is reached the script logs a
    warning and exports the partial result set. Defaults to 2000.

.VERSION
1.14


.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication.
    Minimum delegated Graph permission: BackupRestore-Configuration.Read.All.
    Optional app-only mode requires the application permission BackupRestore-Configuration.Read.All
    and a supported Microsoft 365 Backup service-app/controller registration.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Version : 1.13
    Author: https://github.com/khda79/workplacecloudhub.com
    Required delegated permission by default: BackupRestore-Configuration.Read.All
    Uses the Graph v1.0 Backup Storage API (GA). Note: displayName and email properties of
    mailbox protection units are only returned with delegated permissions, not app-only.
    v1.5 changes:
      - Switched Graph endpoint from beta to v1.0.
      - Replaced Invoke-MgGraphRequest -OutputType PSObject with -OutputType Json plus
        ConvertFrom-Json to work around the SDK JSON-to-PSObject conversion failure
        ("Argument types do not match") observed on large result sets.
      - Added per-page progress logging (page number, page item count, cumulative count).
      - Added MaxPages pagination guard with graceful partial export.
      - Added Graph access probe with full error body logging for 403 diagnostics.
      - StrictMode-safe status property access in protected unit filtering.
    - v1.6: Return Graph collection as a flat object array to avoid Generic.List output binding failures after pagination.
    Optional app-only permission: BackupRestore-Configuration.Read.All
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [switch]$Connect,
    [switch]$InteractiveAuth = $true,
    [switch]$IncludeNonProtected,
    [string]$OutputPath,
    [ValidateRange(1, 100000)][int]$MaxPages = 2000,
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
        foreach ($p in $candidates) { if (Test-Path -LiteralPath $p) { return $p } }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $tenantContextPath
$script:SmartM365EffectiveConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$MaximumFunctionCount = 32768
$ScriptVersion = "1.14"
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
            Import-Module -Name $candidate -MinimumVersion '1.0.24' -Force -ErrorAction Stop
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
    if ($null -ne $script:SmartM365EffectiveConfig) { $script:SmartM365GlobalConfig = $script:SmartM365EffectiveConfig; return $script:SmartM365GlobalConfig }
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


function Stop-SmartM365TranscriptSafely {
    [CmdletBinding()]
    param()

    try { Stop-Transcript | Out-Null } catch {}
    try {
        if (-not [string]::IsNullOrWhiteSpace([string]$global:logTranscriptFile) -and
            (Test-Path -LiteralPath $global:logTranscriptFile -PathType Leaf)) {
            Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile
        }
    } catch {}
}

function Connect-GraphForBackupInventory {
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
        Connect-MgGraph -Scopes @('BackupRestore-Configuration.Read.All') -NoWelcome -ErrorAction Stop | Out-Null
    }
    else {
        WriteLog -Message 'Connecting to Microsoft Graph using app-only certificate authentication.' -Level INFO
        Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $Thumbprint -NoWelcome -ErrorAction Stop | Out-Null
    }
    WriteLog -Message 'Connected to Microsoft Graph.' -Level SUCCESS
}

function Get-SmartM365GraphErrorBody {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$ErrorRecord)

    try {
        if ($ErrorRecord.ErrorDetails -and -not [string]::IsNullOrWhiteSpace([string]$ErrorRecord.ErrorDetails.Message)) {
            return [string]$ErrorRecord.ErrorDetails.Message
        }
    } catch {}
    return ''
}

function Invoke-SmartM365GraphCollectionRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [ValidateRange(1, 100000)][int]$MaxPages = 2000,
        [ValidateRange(0, 1000000)][int]$MaxItems = 0
    )

    $items = New-Object System.Collections.ArrayList
    $nextUri = $Uri
    $pageIndex = 0
    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        if ($pageIndex -ge $MaxPages) {
            WriteLog -Message ("Graph pagination stopped at MaxPages limit ({0}). Result set is PARTIAL: {1} items retrieved." -f $MaxPages, $items.Count) -Level WARNING
            break
        }
        $pageIndex++
        WriteLog -Message ("Calling Graph (page {0}): {1}" -f $pageIndex, $nextUri) -Level DEBUG
        try {
            $rawJson = Invoke-MgGraphRequest -Method GET -Uri $nextUri -OutputType Json -ErrorAction Stop
            $response = $rawJson | ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            $errorBody = Get-SmartM365GraphErrorBody -ErrorRecord $_
            WriteLog -Message ("Graph request failed on page {0} (cumulative items so far: {1}). Error: {2} Body: {3}" -f $pageIndex, $items.Count, $_.Exception.Message, $errorBody) -Level ERROR
            throw
        }
        $pageItems = if ($response.PSObject.Properties['value']) { @($response.value) } else { @($response) }
        foreach ($item in $pageItems) {
            if ($MaxItems -gt 0 -and $items.Count -ge $MaxItems) { break }
            [void]$items.Add($item)
        }
        WriteLog -Message ("Graph page {0} retrieved: {1} items (cumulative exported/test items: {2})." -f $pageIndex, $pageItems.Count, $items.Count) -Level INFO
        if ($MaxItems -gt 0 -and $items.Count -ge $MaxItems) {
            WriteLog -Message ("MaxItems enabled: Graph collection stopped after {0} items." -f $MaxItems) -Level WARNING
            break
        }
        $nextLinkProperty = $response.PSObject.Properties['@odata.nextLink']
        $nextUri = if ($nextLinkProperty -and $null -ne $nextLinkProperty.Value) { [string]$nextLinkProperty.Value } else { '' }
    }
    WriteLog -Message ("Graph collection completed: {0} pages, {1} items." -f $pageIndex, $items.Count) -Level SUCCESS
    return @($items.ToArray())
}

function Test-SmartM365GraphAccess {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Uri)

    try {
        Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType Json -ErrorAction Stop | Out-Null
        WriteLog -Message ("Graph access probe OK: {0}" -f $Uri) -Level INFO
    }
    catch {
        $errorBody = Get-SmartM365GraphErrorBody -ErrorRecord $_
        WriteLog -Message ("Graph access probe failed: {0} Error: {1} Body: {2}" -f $Uri, $_.Exception.Message, $errorBody) -Level ERROR
        throw
    }
}

function Get-BackupMailboxProtectionUnits {
    [CmdletBinding()]
    param([ValidateRange(0, 1000000)][int]$MaxItems = 0)

    $uri = 'https://graph.microsoft.com/v1.0/solutions/backupRestore/protectionUnits/microsoft.graph.mailboxProtectionUnit?$top=999'
    return @(Invoke-SmartM365GraphCollectionRequest -Uri $uri -MaxPages $MaxPages -MaxItems $MaxItems)
}

function Get-SmartM365NestedValue {
    [CmdletBinding()]
    param([AllowNull()]$Object, [Parameter(Mandatory)][string[]]$Names)

    if ($null -eq $Object) { return '' }
    foreach ($name in $Names) {
        $property = $Object.PSObject.Properties[$name]
        if ($null -ne $property -and $null -ne $property.Value) { return [string]$property.Value }
    }
    return ''
}

function ConvertTo-BackupMailboxInventoryRow {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Unit, [Parameter(Mandatory)][string]$ExportDateTime, [Parameter(Mandatory)][string]$RunId)

    $mailboxInfo = $null
    foreach ($propertyName in @('mailboxInfo', 'mailbox', 'directoryObject', 'user')) {
        $property = $Unit.PSObject.Properties[$propertyName]
        if ($null -ne $property -and $null -ne $property.Value) { $mailboxInfo = $property.Value; break }
    }

    [pscustomobject][ordered]@{
        ExportDateTime          = $ExportDateTime
        RunId                   = $RunId
        ProtectionUnitId        = Get-SmartM365NestedValue -Object $Unit -Names @('id')
        Status                  = Get-SmartM365NestedValue -Object $Unit -Names @('status', 'protectionStatus')
        DisplayName             = if ($mailboxInfo) { Get-SmartM365NestedValue -Object $mailboxInfo -Names @('displayName', 'name') } else { Get-SmartM365NestedValue -Object $Unit -Names @('displayName') }
        UserPrincipalName       = if ($mailboxInfo) { Get-SmartM365NestedValue -Object $mailboxInfo -Names @('userPrincipalName', 'upn') } else { '' }
        Mail                    = if ($mailboxInfo) { Get-SmartM365NestedValue -Object $mailboxInfo -Names @('mail', 'emailAddress', 'primarySmtpAddress') } else { Get-SmartM365NestedValue -Object $Unit -Names @('email') }
        MailboxId               = if ($mailboxInfo) { Get-SmartM365NestedValue -Object $mailboxInfo -Names @('id', 'mailboxId', 'userId') } else { Get-SmartM365NestedValue -Object $Unit -Names @('directoryObjectId') }
        MailboxType             = if ($mailboxInfo) { Get-SmartM365NestedValue -Object $mailboxInfo -Names @('mailboxType', 'recipientTypeDetails') } else { Get-SmartM365NestedValue -Object $Unit -Names @('mailboxType') }
        ProtectionPolicyId      = Get-SmartM365NestedValue -Object $Unit -Names @('protectionPolicyId', 'policyId')
        ProtectionPolicyName    = Get-SmartM365NestedValue -Object $Unit -Names @('protectionPolicyName', 'policyName')
        CreatedDateTime         = Get-SmartM365NestedValue -Object $Unit -Names @('createdDateTime')
        LastModifiedDateTime    = Get-SmartM365NestedValue -Object $Unit -Names @('lastModifiedDateTime')
        ErrorCode               = Get-SmartM365NestedValue -Object $Unit -Names @('errorCode')
        ErrorMessage            = Get-SmartM365NestedValue -Object $Unit -Names @('errorMessage')
    }
}

function Send-BackupInventoryTeamsNotification {
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
$global:EnableSharePointUpload = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointTargetFolderPath' -DefaultValue ''
$AppId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AppId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$TenantId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$Thumb = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumb' -DefaultValue '0000000000000000000000000000000000000000'
$OrgDomain = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue 'contoso.onmicrosoft.com'
$ScriptCsvLogFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ScriptCsvLogFolderPath' -DefaultValue ''
$LatestCsvFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''
$LogAllRootPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LogAllRootPath' -DefaultValue ''
$WeeklyHistoryFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'WeeklyHistoryFolderPath' -DefaultValue ''

if (-not [string]::IsNullOrWhiteSpace($OutputPath)) { $ScriptCsvLogFolderPath = $OutputPath }
if ([string]::IsNullOrWhiteSpace($ScriptCsvLogFolderPath)) { $ScriptCsvLogFolderPath = Join-Path -Path $PSScriptRoot -ChildPath 'Output' }
if ([string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) { $LatestCsvFolderPath = $ScriptCsvLogFolderPath }

Ensure-GraphModule

$runId = [guid]::NewGuid().ToString()
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$csvName = 'M365_Backup_ProtectedMailboxes.csv'
$timestampedCsvPath = Join-Path -Path $ScriptCsvLogFolderPath -ChildPath ("M365_Backup_ProtectedMailboxes_{0}.csv" -f $timestamp)
$latestCsvPath = Join-Path -Path $LatestCsvFolderPath -ChildPath $csvName
$logFileBaseName = 'SmartM365-BackupProtectedMailboxes-Inventory'
$logRoot = if ([string]::IsNullOrWhiteSpace($LogAllRootPath)) { Join-Path $ScriptCsvLogFolderPath 'Logs' } else { Join-Path $LogAllRootPath $logFileBaseName }
$logPath = Join-Path -Path $logRoot -ChildPath ("{0}_{1}.log" -f $logFileBaseName,$timestamp)

foreach ($folder in @($ScriptCsvLogFolderPath, $LatestCsvFolderPath, $logRoot)) {
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
}
Set-SmartM365CoreContext -RunId $runId -RunOutputRoot $ScriptCsvLogFolderPath -LatestOutputRoot $LatestCsvFolderPath -LogPath $logPath
$global:LogTextFile = $logPath

$script:CompletionStatus = 'Auto'
try {
    $CurrentOperation = 'InitializeScriptEnvironment'
    $initializedOutput = InitializeScriptEnvironment -OutputPath $ScriptCsvLogFolderPath -LogFileName $logFileBaseName
    $ScriptCsvLogFolderPath = $initializedOutput
    if ([string]::IsNullOrWhiteSpace($WeeklyHistoryFolderPath)) { $WeeklyHistoryFolderPath = Join-Path -Path $ScriptCsvLogFolderPath -ChildPath 'WeeklyHistory' }
    $timestampedCsvPath = Join-Path -Path $ScriptCsvLogFolderPath -ChildPath ("M365_Backup_ProtectedMailboxes_{0}.csv" -f $timestamp)
    $latestCsvPath = Join-Path -Path $LatestCsvFolderPath -ChildPath $csvName
    Start-Transcript -Path $global:logTranscriptFile -Append | Out-Null

    WriteLog -Message "Starting $TaskName." -Level INFO
    WriteLog -Message "Tenant profile: $Tenant" -Level INFO
    WriteLog -Message "Default WeeklyHistoryFolderPath: $WeeklyHistoryFolderPath" -Level INFO
    WriteLog -Message "IncludeNonProtected: $($IncludeNonProtected.IsPresent)" -Level INFO
    WriteLog -Message "MaxPages: $MaxPages" -Level INFO
    WriteLog -Message ("Authentication mode: {0}" -f $(if ($InteractiveAuth) { 'Interactive delegated (default)' } else { 'App-only certificate (explicit override)' })) -Level INFO
    if ($Connect) { WriteLog -Message 'Connect switch specified; Graph connection will be established by this script.' -Level INFO }

    $CurrentOperation = 'ConnectGraph'
    Connect-GraphForBackupInventory -UseInteractiveAuth:$InteractiveAuth -AppId $AppId -TenantId $TenantId -Thumbprint $Thumb

    $CurrentOperation = 'GraphAccessProbe'
    Test-SmartM365GraphAccess -Uri "https://graph.microsoft.com/v1.0/solutions/backupRestore/protectionUnits/microsoft.graph.mailboxProtectionUnit?`$top=1"

    $CurrentOperation = 'Preflight'
    $preflightParameters = @{
        ScriptName      = $TaskName
        RequiredModules = @('Microsoft.Graph.Authentication')
        OutputPaths     = @($ScriptCsvLogFolderPath, $LatestCsvFolderPath)
        GraphProbeUris  = @('https://graph.microsoft.com/v1.0/solutions/backupRestore/protectionUnits/microsoft.graph.mailboxProtectionUnit?$top=1')
    }
    if (-not $InteractiveAuth) {
        $preflightParameters.RequiredGraphApplicationPermissions = @('BackupRestore-Configuration.Read.All')
    }
    Invoke-SmartM365Preflight @preflightParameters | Out-Null

    $CurrentOperation = 'RetrieveBackupProtectionUnits'
    $allUnits = @(Get-BackupMailboxProtectionUnits -MaxItems $MaxItems)
    $selectedUnits = if ($IncludeNonProtected) { $allUnits } else { @($allUnits | Where-Object { (Get-SmartM365NestedValue -Object $_ -Names @('status', 'protectionStatus')) -eq 'protected' }) }
    $exportDateTime = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    $rows = @($selectedUnits | ForEach-Object { ConvertTo-BackupMailboxInventoryRow -Unit $_ -ExportDateTime $exportDateTime -RunId $runId })

    $CurrentOperation = 'ExportCsv'
    Export-SmartM365Csv -Data $rows -TimestampedPath $timestampedCsvPath -LatestPath $latestCsvPath | Out-Null

    $protectedCount = @($allUnits | Where-Object { (Get-SmartM365NestedValue -Object $_ -Names @('status', 'protectionStatus')) -eq 'protected' }).Count
    $resultSummary = "M365 Backup protected mailbox inventory completed. Exported rows: {0}; protected units: {1}; total units: {2}." -f $rows.Count, $protectedCount, $allUnits.Count
    WriteLog -Message $resultSummary -Level SUCCESS

    Send-BackupInventoryTeamsNotification -Level SUCCESS -Title 'M365 Backup protected mailbox inventory success' -Message $resultSummary -Facts @{
        Script = $MyInvocation.MyCommand.Name
        TenantOrOrganization = $OrgDomain
        OutputPath = $ScriptCsvLogFolderPath
        LatestCsvPath = $latestCsvPath
        TimestampedCsvPath = $timestampedCsvPath
        WeeklyHistoryFolderPath = $WeeklyHistoryFolderPath
        TotalUnits = $allUnits.Count
        ProtectedUnits = $protectedCount
        ExportedRows = $rows.Count
        LogFile = $global:LogTextFile
    }

    try { Stop-SmartM365TranscriptSafely } catch {}
    Disconnect-GraphSafe
}
catch {
    $script:CompletionStatus = 'Failed'
    $globalError = $_
    WriteLog -Message ("Global error during {0}: {1}" -f $CurrentOperation, $globalError.Exception.Message) -Level ERROR
    Send-BackupInventoryTeamsNotification -Level ERROR -Title 'M365 Backup protected mailbox inventory failed' -Message $globalError.Exception.Message -Facts @{
        Script = $MyInvocation.MyCommand.Name
        TenantOrOrganization = $OrgDomain
        Operation = $CurrentOperation
        LogFile = $global:LogTextFile
    }
    try { Stop-SmartM365TranscriptSafely } catch {}
    Disconnect-GraphSafe
    throw
}
finally {
    try { RemoveOldFiles -Path $ScriptCsvLogFolderPath -Filter '*.csv' -KeepCount $global:RetentionMaxCSV -LogFile $global:LogTextFile } catch {}
    try { Stop-SmartM365TranscriptSafely } catch {}
    try { Complete-SmartM365ExecutionContext -Status $script:CompletionStatus -FailureStage $(if ($script:CompletionStatus -eq 'Failed') { $CurrentOperation } else { '' }) } catch {}
}
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAhJ6oKBTvMlXRZ
# 6U+BotLPr44NW/Yq/HmcDVQY8DUsGaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIhXHfYFPh54qGDzClwRMfWsL0iVRpAZuq6aZ4H0EvYtMA0GCSqG
# SIb3DQEBAQUABIIBgDMik9M5sBLHIIr0SMFk1CX/p8qtJQd3MeNf/80GmapWl3eR
# FnH4ICbHlmjhMl/ni9Sy+HhOWSuGc5fcJhgaDbV5GmKBypw28EpOJ3NQk35+9zKf
# PTTFxPp5mrhwDjfUk0gX/1isihTYYjcEEuhHgUx5XQ4ctMzW48tXF0Tv7JIp6Ibo
# UW3kjAR3H1/q3tbFY2HZ8p3AzInSHTJFbt4g0EueJBjnvNJTws9TMGbrQ2U3kw1U
# f/bLWv+jNSaIhUatC34/HHfUJXl1+buOVgzZ1NWKLZC+tR0ENqbkBDHFxsBmYA2b
# yJXOlMPm1v9XgLhJkTn7H9VGSfV6PqS/0grTv8w2ygFsKH/8NS8nunX2PVcZN1hk
# FFF95U7aUKn2bsjfpVcWkwJgBLzil2AFxElUfOmyYqt1JeOnQOTRSCcXL6zFuB7h
# af/0W96nZ5LN/wdhkUo7BB2wlpPERWZ4McW0dr1CpnO3UZlf8ma0Me7yLd+/jORb
# ZF3El8F2ypyBzuqRvaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTQwMTEy
# NTRaMC8GCSqGSIb3DQEJBDEiBCACRxcRmwt++iueebwrtbDU00cnbUT49KmQdRUJ
# VoqHQjANBgkqhkiG9w0BAQEFAASCAgAiu/i6AsU/C0zFWpMVougjSP9+HWhIs9ch
# zEFXeggfi4cAnCRlnQYMxOUXJp+XH4ZSWXiK7QqgnCIMyZz/Gmk5tfbCYbSQOI63
# FZJsJrQr2ClQJDboQzZ+VrLFBdVA+msVFU2dNQtV5c53cQT068HHdFO8+B9NP49b
# K5bsgF+r/SNuODxlEI02TkmJ34m/b/K+LZUpfYh7QquoN0HMJE0NT7sfjaS/0Ywd
# 8Id5WAywf6X1/gtBI4jvvV3oFSMdttwzMWUlAFgBfwRNVpHzUQ7dUrQ5qm3YTKE/
# 52iy2SHACO0c4s3kyiHaQJ5H6Gc7C0Jt+yDSPIo2GuWDJadAaP0g4u+0vgmibPYL
# ZcY090sd6+3cgMTWswv+hUDfKhip5Dfjc8r2JnY4YQsfP9I1v8qoqmDlZX4HlztN
# MOrZB4AF/vUwsgr71ln7XGc2TJxymgpwLlXOJVqKLpKaiKNz0I2ivUmOKyHqy2yZ
# JdXdy3QQUQW/xtmEGVmzbX4w9FkozrxNzLKayKU+ZQMgYYzyd9N5uNSkKk2Po7hd
# JGpx/l19AnEcbhirG13q9MbYx6XeIcAzpZJVb+JcQJOVzy4J96QHgLWh297lzJhZ
# 4OBznqgq8wKlXhzg/OehYRor+jayBFASzLPSwBG5zO/nvsP9edx3SN5N186gmVZ4
# J4mewZjyUA==
# SIG # End signature block
