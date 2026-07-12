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
    Uses delegated interactive Microsoft Graph authentication for troubleshooting.

.PARAMETER IncludeNonProtected
    Include non-protected mailbox protection units in the CSV. By default only Status = protected is exported.

.PARAMETER OutputPath
    Optional output directory override. If omitted, ScriptCsvLogFolderPath from local JSON is used.

.PARAMETER MaxPages
    Maximum number of Graph pages to retrieve. Acts as a safety guard against endless pagination and
    as a quick smoke test limiter (e.g. -MaxPages 3). When the limit is reached the script logs a
    warning and exports the partial result set. Defaults to 2000.

.VERSION
1.8


.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication.
    Minimum Graph application permissions: BackupRestore-Configuration.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Required application permission: BackupRestore-Configuration.Read.All
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
    Minimum application permissions: BackupRestore-Configuration.Read.All
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [switch]$Connect,
    [switch]$InteractiveAuth,
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
$ScriptVersion = '1.8'
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
$logRoot = if ([string]::IsNullOrWhiteSpace($LogAllRootPath)) { Join-Path $ScriptCsvLogFolderPath 'Logs' } else { Join-Path $LogAllRootPath 'M365-BackupProtectedMailboxes' }
$logPath = Join-Path -Path $logRoot -ChildPath ("SmartM365-M365-BackupProtectedMailboxes-Inventory_{0}.log" -f $timestamp)

foreach ($folder in @($ScriptCsvLogFolderPath, $LatestCsvFolderPath, $logRoot)) {
    if (-not (Test-Path -LiteralPath $folder)) { New-Item -Path $folder -ItemType Directory -Force | Out-Null }
}
Set-SmartM365CoreContext -RunId $runId -RunOutputRoot $ScriptCsvLogFolderPath -LatestOutputRoot $LatestCsvFolderPath -LogPath $logPath
$global:LogTextFile = $logPath

try {
    $CurrentOperation = 'InitializeScriptEnvironment'
    $initializedOutput = InitializeScriptEnvironment -OutputPath $ScriptCsvLogFolderPath -LogFileName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))
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
    if ($Connect) { WriteLog -Message 'Connect switch specified; Graph connection will be established by this script.' -Level INFO }

    $CurrentOperation = 'ConnectGraph'
    Connect-GraphForBackupInventory -UseInteractiveAuth:$InteractiveAuth -AppId $AppId -TenantId $TenantId -Thumbprint $Thumb

    $CurrentOperation = 'GraphAccessProbe'
    Test-SmartM365GraphAccess -Uri "https://graph.microsoft.com/v1.0/solutions/backupRestore/protectionUnits/microsoft.graph.mailboxProtectionUnit?`$top=1"

    $CurrentOperation = 'Preflight'
    Invoke-SmartM365Preflight `
        -ScriptName $TaskName `
        -RequiredModules @('Microsoft.Graph.Authentication') `
        -OutputPaths @($ScriptCsvLogFolderPath, $LatestCsvFolderPath) `
        -RequiredGraphApplicationPermissions @('BackupRestore-Configuration.Read.All') -GraphProbeUris @("https://graph.microsoft.com/v1.0/solutions/backupRestore/protectionUnits/microsoft.graph.mailboxProtectionUnit?`$top=1") | Out-Null

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
    try { Complete-SmartM365ExecutionContext -Status Auto } catch {}
    Disconnect-GraphSafe
}
catch {
    $globalError = $_
    WriteLog -Message ("Global error during {0}: {1}" -f $CurrentOperation, $globalError.Exception.Message) -Level ERROR
    Send-BackupInventoryTeamsNotification -Level ERROR -Title 'M365 Backup protected mailbox inventory failed' -Message $globalError.Exception.Message -Facts @{
        Script = $MyInvocation.MyCommand.Name
        TenantOrOrganization = $OrgDomain
        Operation = $CurrentOperation
        LogFile = $global:LogTextFile
    }
    try { Stop-SmartM365TranscriptSafely } catch {}
    try { Complete-SmartM365ExecutionContext -Status Failed -FailureStage $CurrentOperation } catch {}
    Disconnect-GraphSafe
    throw
}
finally {
    try { RemoveOldFiles -Path $ScriptCsvLogFolderPath -Filter '*.csv' -KeepCount $global:RetentionMaxCSV -LogFile $global:LogTextFile } catch {}
    try { Stop-SmartM365TranscriptSafely } catch {}
}

# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBzEksThNdn5rqr
# YFifi0sD8zYPM+uXXrdRmSKirlb+I6CCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCD/hI7riMOceFU/nVb3
# i+Xvj4sAl7Pn45IlqbbGsfQ4wTANBgkqhkiG9w0BAQEFAASCAYBvpakTRxeUlwDS
# xSbjYwrgXKIDT1dbErLOqMswB6SVbVoC5M24Q0vYuAReJlqYXoG5j2wDuUBWIPLJ
# aB4jyr5qyYkU8OGRYsWq57Do/DRZVTDgHCcrzrbB1Sq7mrFccBrLYC1PR9R/Moa+
# tL7sqGdnADIk8jcfKhNHyR97fF87WvFTylhF57Ncpa3OwNl6lj8+pLMDUy2yRh9f
# JsquIWVfT0Zg0mPTo/GjXug7yEN5GAlKPdTTdMSGP1CEkIQKCiS/5ZRI4NVMA+jD
# ujqL8HxrMSeVxsvsbKwqTvw598IHeLBFEuf9MPNqWsAVu64oEOt46YNI9dTIPS0W
# 7qQsDqQyhxq6UC+WR3YC2XF4ysIWMztOX6lekVk8j7THJ5tiaxG38+Ueh3538s89
# 0NvFyERLRhN23Xsq81rb/hmFjqvp2VTJCO6L4ib/8ZiX4hE5iAhp29migd4xLKc1
# GWRJpdivO0/jMc+7Ca1Pys1v/M2R+HaaiLjgiK99tHP658WVwDY=
# SIG # End signature block
