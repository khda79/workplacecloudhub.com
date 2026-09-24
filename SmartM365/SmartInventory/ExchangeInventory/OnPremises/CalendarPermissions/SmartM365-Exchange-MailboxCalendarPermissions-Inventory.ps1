<#
.SYNOPSIS
    Inventories Exchange on-premises mailbox calendar permissions.

.DESCRIPTION
    Retrieves calendar folder permissions for Exchange 2016 user, shared, room, and equipment mailboxes.
    The script runs sequentially in Windows PowerShell 5.1 with the Exchange Management Shell.
    Primary-calendar mode resolves the localized folder from mailbox statistics before querying permissions.
    It exports the stable Mailbox, UPN, CalendarFolder, User, and AccessRights schema.

.PARAMETER PrimaryOnly
    Scans only the main Calendar folder by default. Disable it to scan every Calendar folder.

.PARAMETER EmitNoPermRow
    Emits one "(none)" row when a calendar exists but has no explicit permissions.

.PARAMETER TopMailboxes
    Limits mailbox processing to the first N mailboxes for smoke tests. Default 0 processes all mailboxes.

.PARAMETER BackendPreflightTimeoutSeconds
    Bounds the Exchange MAPI connectivity preflight for each mailbox database. Set to 0 to disable it.

.VERSION
1.7

.REQUIREMENTS
    Windows PowerShell 5.1 on an Exchange 2016 management host.
    Modules/snap-ins: SmartM365 WindowsPowerShell5 compatibility module; Exchange Management snap-in.
    Minimum permissions: Exchange on-premises recipient and mailbox-folder permission read access.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled; Mail.Send is required only when Graph mail is enabled.

.NOTES
    Version : 1.7
    Author: https://github.com/khda79/workplacecloudhub.com
    Environment : Exchange 2016 On-Premises
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [switch]$PrimaryOnly = $true,
    [switch]$EmitNoPermRow = $true,
    [int]$TopMailboxes = 0,
    [ValidateRange(0, 300)][int]$BackendPreflightTimeoutSeconds = 30,
    [string]$OutputPath,
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
if ($PSVersionTable.PSEdition -ne 'Desktop' -or $PSVersionTable.PSVersion.Major -ne 5) {
    Write-Host 'This script requires Windows PowerShell 5.1.' -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion) ($($PSVersionTable.PSEdition))" -ForegroundColor Yellow
    exit 1
}
$MaximumFunctionCount = 32768
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
if (-not $PSBoundParameters.ContainsKey('BackendPreflightTimeoutSeconds')) {
    $BackendPreflightTimeoutSeconds = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'BackendPreflightTimeoutSeconds' -DefaultValue 30)
    if ($BackendPreflightTimeoutSeconds -lt 0 -or $BackendPreflightTimeoutSeconds -gt 300) {
        throw "BackendPreflightTimeoutSeconds must be between 0 and 300."
    }
}

$global:RetentionMaxCSV = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxLogs' -DefaultValue 30)

$global:EnableSharePointUpload = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointLibraryDisplayName' -DefaultValue 'Documents'
$global:SharePointTargetFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointTargetFolderPath' -DefaultValue ''
function Join-ModulePath {
    param([Parameter(Mandatory)][string]$FileName)
    $searchRoot = $PSScriptRoot
    while ($searchRoot) {
        $coreRoot = Join-Path -Path (Join-Path -Path $searchRoot -ChildPath 'Modules') -ChildPath 'SmartM365.Core'
        $candidate = Join-Path -Path (Join-Path -Path $coreRoot -ChildPath 'Compatibility\WindowsPowerShell5') -ChildPath $FileName
        if (Test-Path -LiteralPath $candidate) { return $candidate }
        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }
    throw "SmartM365 WindowsPowerShell5 compatibility module file not found: $FileName"
}

$ScriptVersion = '1.7'
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LocalCalendarPermissionsCsvLogFolderPath' -DefaultValue $OutputPath
$TaskName = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."

try {
    Write-Host 'Loading module SmartM365-WindowsPowerShell5.psd1...' -ForegroundColor Cyan
    Import-Module -Name (Join-ModulePath 'SmartM365-WindowsPowerShell5.psd1') -MinimumVersion '1.0.41' -ErrorAction Stop
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPath $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
    Start-Transcript -Path $global:logTranscriptFile -Append
    $logTextFile = Join-Path $logPath "$(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')-OnPrem-$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"
    WriteLog -Message "Script Environment initialized at $InitializeOutputPath"
    $OutputPath = $InitializeOutputPath
    WriteLog -Message "Starting $TaskName..."
}
catch {
    Write-Host "Initialization failed: $_" -ForegroundColor Red
    exit 1
}
Write-Host "Detecting Exchange environment..." -ForegroundColor Cyan

# Helper: detect if a command is available
function Test-HasCommand {
    param([Parameter(Mandatory)][string]$Name)
    return [bool](Get-Command -Name $Name -ErrorAction SilentlyContinue)
}

# Helper: silently run a scriptblock (suppresses Information/Verbose/Progress noise from EXO cmdlets)
function Invoke-Quiet {
    param([Parameter(Mandatory)][scriptblock]$Script)
    $oldInfo     = $InformationPreference
    $oldVerbose  = $VerbosePreference
    $oldProgress = $ProgressPreference
    try {
        $InformationPreference = 'SilentlyContinue'
        $VerbosePreference     = 'SilentlyContinue'
        $ProgressPreference    = 'SilentlyContinue'
        & $Script
    }
    finally {
        $InformationPreference = $oldInfo
        $VerbosePreference     = $oldVerbose
        $ProgressPreference    = $oldProgress
    }
}
$snapinName = 'Microsoft.Exchange.Management.PowerShell.SnapIn'
try {
    if (-not (Get-PSSnapin $snapinName -Registered -ErrorAction SilentlyContinue)) {
        throw "Exchange Management PSSnapin '$snapinName' is not registered on this server."
    }
    if (-not (Get-PSSnapin $snapinName -ErrorAction SilentlyContinue)) {
        Add-PSSnapin $snapinName -ErrorAction Stop
        WriteLog -Message 'Exchange On-Prem PSSnapin loaded.'
    }
    else {
        WriteLog -Message 'Exchange On-Prem PSSnapin detected.'
    }
    Set-ADServerSettings -ViewEntireForest $true -ErrorAction Stop
    WriteLog -Message 'Set-ADServerSettings -ViewEntireForest True applied.'
}
catch {
    WriteLog -Message ("Unable to initialize Exchange Management Shell: {0}" -f $_.Exception.Message) 'ERROR'
    Stop-Transcript | Out-Null
    try { if ($global:logTranscriptFile) { Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile } } catch {}
    Complete-SmartM365ExecutionContext -Status Auto
    exit 1
}

Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -RequireExchangeOnPrem | Out-Null
$results = New-Object 'System.Collections.Generic.List[object]'
$errors = New-Object 'System.Collections.Generic.List[object]'
$script:UnavailableCalendarBackends = @{}
$script:CalendarBackendPreflightResults = @{}
$processed = 0

try {
    $recipientTypes = @('UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox')
    $mailboxes = Invoke-Quiet {
        Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails $recipientTypes -ErrorAction Stop
    } | Select-Object DisplayName, PrimarySmtpAddress, UserPrincipalName, Identity, Guid, ServerName, Database
}
catch {
    WriteLog -Message "Mailbox enumeration failed : $($_.Exception.Message)" 'ERROR'
    Stop-Transcript | Out-Null
    try { if ($global:logTranscriptFile) { Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile } } catch {}
    Complete-SmartM365ExecutionContext -Status Auto
    exit 1
}

$total = $mailboxes.Count
WriteLog -Message "Total mailboxes found: $total"
if ($TopMailboxes -gt 0 -and $total -gt $TopMailboxes) {
    WriteLog -Message ("TopMailboxes enabled: processing first {0} of {1} mailboxes." -f $TopMailboxes, $total) 'WARNING'
    $mailboxes = @($mailboxes | Select-Object -First $TopMailboxes)
    $total = $mailboxes.Count
}
if ($total -eq 0) {
    WriteLog -Message 'No mailboxes found. Stopping.'
    Stop-Transcript | Out-Null
    try { if ($global:logTranscriptFile) { Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile } } catch {}
    Complete-SmartM365ExecutionContext -Status Auto
    exit 0
}

function Get-SmartM365CalendarFailureCategory {
    param([AllowEmptyString()][string]$Message)

    if ($Message -match '(?i)information store.*(is not available|inaccessible|unavailable)|cannot open mailbox.*microsoft system attendant') { return 'BackendUnavailable' }
    if ($Message -match '(?i)couldn''t find.*as a recipient|could not find.*recipient') { return 'RecipientNotFound' }
    if ($Message -match '(?i)doesn''t represent a unique recipient|isn''t unique|ambiguous') { return 'RecipientAmbiguous' }
    return 'CalendarFolderStatisticsFailure'
}

function Get-SmartM365CalendarBackendName {
    param(
        [AllowEmptyString()][string]$Message,
        [AllowNull()]$Mailbox
    )

    if ($Message -match "(?i)server\s+'(?<Server>[^']+)'" -and -not [string]::IsNullOrWhiteSpace($Matches['Server'])) { return $Matches['Server'] }
    if ($Message -match '(?i)cn=Servers/cn=(?<Server>[^/,"]+)' -and -not [string]::IsNullOrWhiteSpace($Matches['Server'])) { return $Matches['Server'] }
    if ($Mailbox -and -not [string]::IsNullOrWhiteSpace([string]$Mailbox.ServerName)) { return [string]$Mailbox.ServerName }
    return ''
}

function New-SmartM365CalendarLookupException {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter(Mandatory = $true)][string]$Category,
        [Parameter(Mandatory = $true)][int]$Attempts,
        [Parameter(Mandatory = $true)][double]$DurationSeconds,
        [AllowEmptyString()][string]$Backend = '',
        [AllowEmptyString()][string]$Database = '',
        [AllowNull()][System.Exception]$InnerException
    )

    $exception = New-Object System.InvalidOperationException($Message, $InnerException)
    $exception.Data['SmartM365Category'] = $Category
    $exception.Data['SmartM365Attempts'] = $Attempts
    $exception.Data['SmartM365DurationSeconds'] = [math]::Round($DurationSeconds, 3)
    $exception.Data['SmartM365Backend'] = $Backend
    $exception.Data['SmartM365Database'] = $Database
    return $exception
}

function Test-SmartM365CalendarBackendPreflight {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]$Mbx,
        [ValidateRange(0, 300)][int]$TimeoutSeconds = 30
    )

    if ($TimeoutSeconds -le 0 -or -not (Test-HasCommand -Name 'Test-MAPIConnectivity')) { return $true }

    $mailboxServer = [string]$Mbx.ServerName
    $mailboxDatabase = [string]$Mbx.Database
    $backendKey = if (-not [string]::IsNullOrWhiteSpace($mailboxDatabase)) {
        'database:' + $mailboxDatabase.Trim().ToLowerInvariant()
    }
    elseif (-not [string]::IsNullOrWhiteSpace($mailboxServer)) {
        'server:' + $mailboxServer.Trim().ToLowerInvariant()
    }
    else {
        return $true
    }

    if ($script:CalendarBackendPreflightResults.ContainsKey($backendKey)) { return $true }

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $probeParams = @{
        PerConnectionTimeout = $TimeoutSeconds
        AllConnectionsTimeout = $TimeoutSeconds
        ErrorAction = 'Stop'
    }
    if (-not [string]::IsNullOrWhiteSpace($mailboxDatabase)) { $probeParams['Database'] = $mailboxDatabase }
    else { $probeParams['Server'] = $mailboxServer }

    try {
        $probeRows = @(Invoke-Quiet { Test-MAPIConnectivity @probeParams })
    }
    catch {
        $message = [string]$_.Exception.Message
        $category = Get-SmartM365CalendarFailureCategory -Message $message
        $isConnectivityFailure = $category -eq 'BackendUnavailable' -or $message -match '(?i)(MAPI|RPC|information store|mailbox database).*(timed out|timeout|not available|unavailable|inaccessible|failed)'
        if ($isConnectivityFailure) {
            foreach ($cacheKey in @(
                if (-not [string]::IsNullOrWhiteSpace($mailboxServer)) { 'server:' + $mailboxServer.Trim().ToLowerInvariant() }
                if (-not [string]::IsNullOrWhiteSpace($mailboxDatabase)) { 'database:' + $mailboxDatabase.Trim().ToLowerInvariant() }
            )) {
                $script:UnavailableCalendarBackends[$cacheKey] = $message
            }
            throw (New-SmartM365CalendarLookupException -Message $message -Category 'BackendUnavailable' -Attempts 0 -DurationSeconds $stopwatch.Elapsed.TotalSeconds -Backend $mailboxServer -Database $mailboxDatabase -InnerException $_.Exception)
        }

        $script:CalendarBackendPreflightResults[$backendKey] = 'Inconclusive'
        WriteLog -Message ("MAPI connectivity preflight was inconclusive for backend '{0}' / database '{1}'; continuing with calendar statistics: {2}" -f $mailboxServer, $mailboxDatabase, $message) -Level 'WARNING'
        return $true
    }

    $failedProbe = @($probeRows | Where-Object {
        $errorText = [string]$_.Error
        $resultProperty = $_.PSObject.Properties['Result']
        $resultText = if ($resultProperty) { [string]$resultProperty.Value } else { '' }
        -not [string]::IsNullOrWhiteSpace($errorText) -or
            (-not [string]::IsNullOrWhiteSpace($resultText) -and $resultText -notmatch '^(?i:success|passed)$')
    } | Select-Object -First 1)

    if ($failedProbe.Count -gt 0) {
        $failureMessage = [string]$failedProbe[0].Error
        if ([string]::IsNullOrWhiteSpace($failureMessage)) {
            $failureMessage = "MAPI connectivity preflight returned result '$([string]$failedProbe[0].Result)'."
        }
        foreach ($cacheKey in @(
            if (-not [string]::IsNullOrWhiteSpace($mailboxServer)) { 'server:' + $mailboxServer.Trim().ToLowerInvariant() }
            if (-not [string]::IsNullOrWhiteSpace($mailboxDatabase)) { 'database:' + $mailboxDatabase.Trim().ToLowerInvariant() }
        )) {
            $script:UnavailableCalendarBackends[$cacheKey] = $failureMessage
        }
        throw (New-SmartM365CalendarLookupException -Message $failureMessage -Category 'BackendUnavailable' -Attempts 0 -DurationSeconds $stopwatch.Elapsed.TotalSeconds -Backend $mailboxServer -Database $mailboxDatabase)
    }

    if ($probeRows.Count -eq 0) {
        WriteLog -Message ("MAPI connectivity preflight returned no result for backend '{0}' / database '{1}'; continuing with calendar statistics." -f $mailboxServer, $mailboxDatabase) -Level 'WARNING'
        $script:CalendarBackendPreflightResults[$backendKey] = 'Inconclusive'
    }
    else {
        $script:CalendarBackendPreflightResults[$backendKey] = 'Healthy'
    }
    return $true
}

function New-SmartM365CalendarErrorRecord {
    param(
        [Parameter(Mandatory = $true)][int]$Index,
        [Parameter(Mandatory = $true)][string]$Mailbox,
        [AllowEmptyString()][string]$UPN = '',
        [Parameter(Mandatory = $true)][string]$Operation,
        [AllowNull()]$ErrorRecord,
        [AllowEmptyString()][string]$Category = '',
        [int]$Attempts = 1,
        [double]$DurationSeconds = 0,
        [AllowEmptyString()][string]$Backend = '',
        [AllowEmptyString()][string]$Database = '',
        [AllowEmptyString()][string]$Message = ''
    )

    $exception = $null
    if ($ErrorRecord) { $exception = $ErrorRecord.Exception }
    if ($exception) {
        if ([string]::IsNullOrWhiteSpace($Message)) { $Message = [string]$exception.Message }
        if ($exception.Data) {
            if ($exception.Data.Contains('SmartM365Category')) { $Category = [string]$exception.Data['SmartM365Category'] }
            if ($exception.Data.Contains('SmartM365Attempts')) { $Attempts = [int]$exception.Data['SmartM365Attempts'] }
            if ($exception.Data.Contains('SmartM365DurationSeconds')) { $DurationSeconds = [double]$exception.Data['SmartM365DurationSeconds'] }
            if ($exception.Data.Contains('SmartM365Backend')) { $Backend = [string]$exception.Data['SmartM365Backend'] }
            if ($exception.Data.Contains('SmartM365Database')) { $Database = [string]$exception.Data['SmartM365Database'] }
        }
    }
    if ([string]::IsNullOrWhiteSpace($Category)) { $Category = Get-SmartM365CalendarFailureCategory -Message $Message }

    return [pscustomobject][ordered]@{
        Index           = $Index
        Mailbox         = $Mailbox
        UPN             = $UPN
        Category        = $Category
        Operation       = $Operation
        Attempts        = $Attempts
        DurationSeconds = [math]::Round($DurationSeconds, 3)
        Backend         = $Backend
        Database        = $Database
        Message         = $Message
    }
}

function Get-CalendarFoldersSafe {
    param(
        [Parameter(Mandatory = $true)]$Mbx,
        [Parameter(Mandatory = $true)][bool]$PrimaryOnly
    )

    $identityCandidates = @(
        [string]$Mbx.Identity
        [string]$Mbx.UserPrincipalName
        [string]$Mbx.PrimarySmtpAddress
        if ($Mbx.Guid -and $Mbx.Guid -ne [guid]::Empty) { [string]$Mbx.Guid }
    ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $attempts = 0
    $mailboxServer = [string]$Mbx.ServerName
    $mailboxDatabase = [string]$Mbx.Database
    $backendCacheKeys = @()
    if (-not [string]::IsNullOrWhiteSpace($mailboxServer)) { $backendCacheKeys += 'server:' + $mailboxServer.Trim().ToLowerInvariant() }
    if (-not [string]::IsNullOrWhiteSpace($mailboxDatabase)) { $backendCacheKeys += 'database:' + $mailboxDatabase.Trim().ToLowerInvariant() }
    $cachedBackendKey = @($backendCacheKeys | Where-Object { $script:UnavailableCalendarBackends.ContainsKey($_) } | Select-Object -First 1)
    if ($cachedBackendKey.Count -gt 0) {
        $cachedMessage = "Calendar statistics skipped because backend '$mailboxServer' or database '$mailboxDatabase' was already marked unavailable during this run."
        $cachedException = New-SmartM365CalendarLookupException -Message $cachedMessage -Category 'BackendPreviouslyUnavailable' -Attempts 0 -DurationSeconds $stopwatch.Elapsed.TotalSeconds -Backend $mailboxServer -Database $mailboxDatabase
        throw $cachedException
    }

    [void](Test-SmartM365CalendarBackendPreflight -Mbx $Mbx -TimeoutSeconds $BackendPreflightTimeoutSeconds)

    $folders = @()
    $statisticsQuerySucceeded = $false
    $lastStatisticsError = $null
    $lastCategory = 'CalendarFolderStatisticsFailure'
    $lastBackend = $mailboxServer
    foreach ($id in $identityCandidates) {
        $attempts++
        try {
            $folders = @(Invoke-Quiet {
                Get-MailboxFolderStatistics -Identity $id -FolderScope Calendar -ErrorAction Stop
            })
            $statisticsQuerySucceeded = $true
            break
        }
        catch {
            $lastStatisticsError = $_
            $lastCategory = Get-SmartM365CalendarFailureCategory -Message $_.Exception.Message
            $lastBackend = Get-SmartM365CalendarBackendName -Message $_.Exception.Message -Mailbox $Mbx
            if ($lastCategory -eq 'BackendUnavailable') {
                foreach ($backendName in @($lastBackend, $mailboxServer)) {
                    if (-not [string]::IsNullOrWhiteSpace($backendName)) {
                        $script:UnavailableCalendarBackends['server:' + $backendName.Trim().ToLowerInvariant()] = [string]$_.Exception.Message
                    }
                }
                if (-not [string]::IsNullOrWhiteSpace($mailboxDatabase)) {
                    $script:UnavailableCalendarBackends['database:' + $mailboxDatabase.Trim().ToLowerInvariant()] = [string]$_.Exception.Message
                }
                break
            }
        }
    }

    if (-not $statisticsQuerySucceeded) {
        if ($lastStatisticsError) {
            $lookupException = New-SmartM365CalendarLookupException -Message ([string]$lastStatisticsError.Exception.Message) -Category $lastCategory -Attempts $attempts -DurationSeconds $stopwatch.Elapsed.TotalSeconds -Backend $lastBackend -Database $mailboxDatabase -InnerException $lastStatisticsError.Exception
            throw $lookupException
        }
        return @()
    }
    if (-not $folders) { return @() }
    if (-not $PrimaryOnly) { return $folders | Where-Object { $_.FolderType -eq 'Calendar' } }
    $rootCalendars = $folders | Where-Object { $_.FolderType -eq 'Calendar' -and $_.FolderPath -match '^/[^/]+$' }
    if ($rootCalendars) { return ,($rootCalendars | Sort-Object ItemsInFolder -Descending | Select-Object -First 1) }
    $anyCalendar = $folders | Where-Object { $_.FolderType -eq 'Calendar' } | Select-Object -First 1
    if ($anyCalendar) { return ,$anyCalendar }
    return @()
}
function Try-GetFolderPermission {
    param(
        [Parameter(Mandatory)][string[]]$MailboxIds,   # e.g. @($upn, $primarySMTP, $mbx.Identity)
        [Parameter(Mandatory)][string[]]$FolderNames,  # e.g. @('Calendar','Calendrier', $fromStats)
        [Parameter(Mandatory)][string]$PrimarySmtpForLog,
        [Parameter(Mandatory)][AllowEmptyString()][string]$UpnForLog
    )
    foreach ($mbId in $MailboxIds) {
        foreach ($fname in $FolderNames) {
            $identity = ("{0}:\{1}" -f $mbId, $fname)
            try {
                $perms = Invoke-Quiet {
                    Get-MailboxFolderPermission -Identity $identity -ErrorAction Stop
                } |
                    Where-Object { $_.User -notin @("Default","Anonymous") } |
                    Select-Object @{Name = "Mailbox";        Expression = { $PrimarySmtpForLog }},
                                  @{Name = "UPN";            Expression = { $UpnForLog }},
                                  @{Name = "CalendarFolder"; Expression = { $fname }},
                                  @{Name = "User";           Expression = { $_.User }},
                                  @{Name = "AccessRights";   Expression = { ($_.AccessRights -join ",") }}
                return [pscustomobject]@{
                    Ok          = $true
                    Permissions = @($perms)
                    Identity    = $identity
                }
            } catch {
                # Try next combination
            }
        }
    }
    return [pscustomobject]@{
        Ok          = $false
        Permissions = @()
        Identity    = $null
    }
}

function Publish-SmartM365CalendarWeeklyHistory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][AllowEmptyCollection()][string[]]$SourceCsvPaths,
        [Parameter(Mandatory = $true)]$Config,
        [Parameter(Mandatory = $true)][string]$FallbackRootPath
    )

    if (Test-SmartM365MaxItemsMode) {
        WriteLog -Message 'WeeklyHistory publication skipped because MaxItems test mode is active.' -Level 'WARNING'
        return
    }

    $weeklyHistoryEnabled = ConvertTo-SmartM365ConfigBoolean -Value (Get-ScriptLocalConfigValue -Config $Config -Name 'EnableWeeklyHistory' -DefaultValue $true) -DefaultValue $true
    if (-not $weeklyHistoryEnabled) {
        WriteLog -Message 'WeeklyHistory publication is disabled by configuration.' -Level 'INFO'
        return
    }

    $validatedSourcePaths = @($SourceCsvPaths |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and (Test-Path -LiteralPath $_ -PathType Leaf) } |
        Select-Object -Unique)
    if ($validatedSourcePaths.Count -eq 0) {
        WriteLog -Message 'WeeklyHistory publication skipped because no current calendar CSV was produced.' -Level 'INFO'
        return
    }

    $historyRootPath = [string](Get-ScriptLocalConfigValue -Config $Config -Name 'WeeklyHistoryFolderPath' -DefaultValue '')
    if ([string]::IsNullOrWhiteSpace($historyRootPath)) {
        $historyRootPath = Join-Path -Path $FallbackRootPath -ChildPath 'WeeklyHistory'
    }
    $retentionWeeks = [int](Get-ScriptLocalConfigValue -Config $Config -Name 'WeeklyHistoryRetentionWeeks' -DefaultValue 52)
    Add-SmartM365WeeklyHistory -SourceCsvPaths $validatedSourcePaths `
        -HistoryRootPath $historyRootPath `
        -RetentionWeeks $retentionWeeks `
        -HistoryLabel 'Exchange on-premises calendar permissions' `
        -UploadChangedFilesOnly
}

# Helper: emit a "(none)" row when calendar exists but no explicit permissions
function Add-NoPermissionRow {
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][AllowEmptyString()][string]$UPN,
        [Parameter(Mandatory)][string]$CalendarFolder
    )
    return [pscustomobject]@{
        Mailbox        = $Mailbox
        UPN            = $UPN
        CalendarFolder = $CalendarFolder
        User           = '(none)'
        AccessRights   = '(none)'
    }
}

function Add-CalendarResultRows {
    [CmdletBinding()]
    param([AllowNull()][object[]]$Rows)

    foreach ($row in @($Rows)) {
        if ($null -ne $row) { [void]$results.Add($row) }
    }
}

# ------------------------- Processing Loop -------------------------
$overallActivity = "Calendar permissions inventory"
$index = 0

foreach ($mbx in $mailboxes) {
    $index++; $processed++
    $primarySMTP = [string]$mbx.PrimarySmtpAddress
    if ([string]::IsNullOrWhiteSpace($primarySMTP)) {
        $primarySMTP = [string]$mbx.Identity
    }
    $upn = [string]$mbx.UserPrincipalName
    $percent     = [int](($index / $total) * 100)

    Write-Progress -Id 0 -Activity $overallActivity -Status ("[{0}/{1}] {2}" -f $index, $total, $primarySMTP) -PercentComplete $percent

    try {
        if ($PrimaryOnly) {
            $mailboxIds = @($mbx.Identity, $primarySMTP)
            $calendarFolders = @(Get-CalendarFoldersSafe -Mbx $mbx -PrimaryOnly:$true)
            $folderNames = @('Calendar','Calendrier','Kalender','Calendario')
            $lookupSource = 'common name fallback'
            if ($calendarFolders.Count -gt 0) {
                $folderNames = @($calendarFolders[0].FolderPath.TrimStart('/'))
                $lookupSource = 'folder statistics'
            }

            $permissionResult = Try-GetFolderPermission -MailboxIds $mailboxIds -FolderNames $folderNames -PrimarySmtpForLog $primarySMTP -UpnForLog $upn
            if ($permissionResult.Ok) {
                $permissions = @($permissionResult.Permissions)
                $usedIdentity = [string]$permissionResult.Identity
                $calendarFolder = [string]$folderNames[0]
                if ($usedIdentity -match ':\\(.+)$') {
                    $calendarFolder = $Matches[1]
                }
                WriteLog -Message "Calendar folder found for $primarySMTP (via $lookupSource : $usedIdentity)" "INFO"
                if ($permissions.Count -gt 0) {
                    Add-CalendarResultRows -Rows $permissions
                }
                elseif ($EmitNoPermRow) {
                    Add-CalendarResultRows -Rows (Add-NoPermissionRow -Mailbox $primarySMTP -UPN $upn -CalendarFolder $calendarFolder)
                }
            }
            else {
                WriteLog -Message "No calendar folder found for $primarySMTP" "WARNING"
            }
        } else {
            # Full mode
            $calendarFolders = Get-CalendarFoldersSafe -Mbx $mbx -PrimaryOnly:$false
            $fTotal = $calendarFolders.Count
            $fIndex = 0

            foreach ($folder in $calendarFolders) {
                $fIndex++
                $folderPath = $folder.FolderPath.TrimStart("/")
                WriteLog -Message "Calendar folder found for $primarySMTP ($folderPath)" "INFO"
                try {
                    $permMailboxId = $mbx.Identity
                    $permIdentity  = ("{0}:\{1}" -f $permMailboxId, $folderPath)
                    $permissions = Invoke-Quiet {
                        Get-MailboxFolderPermission -Identity $permIdentity -ErrorAction Stop
                    } |
                        Where-Object { $_.User -notin @("Default","Anonymous") } |
                        Select-Object @{Name = "Mailbox";        Expression = { $primarySMTP }},
                                      @{Name = "UPN";            Expression = { $upn }},
                                      @{Name = "CalendarFolder"; Expression = { $folderPath }},
                                      @{Name = "User";           Expression = { $_.User }},
                                      @{Name = "AccessRights";   Expression = { ($_.AccessRights -join ",") }}
                    if ($permissions -and $permissions.Count -gt 0) {
                        Add-CalendarResultRows -Rows $permissions
                    } elseif ($EmitNoPermRow) {
                        Add-CalendarResultRows -Rows (Add-NoPermissionRow -Mailbox $primarySMTP -UPN $upn -CalendarFolder $folderPath)
                    }
                } catch {
                    $errMsg = "Permission error for $primarySMTP ($folderPath) : $($_.Exception.Message)"
                    $errorRow = New-SmartM365CalendarErrorRecord -Index ($errors.Count + 1) -Mailbox $primarySMTP -UPN $upn -Operation 'Get-MailboxFolderPermission' -ErrorRecord $_ -Category 'CalendarPermissionLookupFailure' -Message $_.Exception.Message
                    WriteLog -Message $errMsg "WARNING"
                    [void]$errors.Add($errorRow)
                }
            }
        }
    } catch {
        $errorRow = New-SmartM365CalendarErrorRecord -Index ($errors.Count + 1) -Mailbox $primarySMTP -UPN $upn -Operation 'Get-MailboxFolderStatistics' -ErrorRecord $_
        $errMsg = "Calendar lookup error for $primarySMTP (category=$($errorRow.Category); attempts=$($errorRow.Attempts); duration=$($errorRow.DurationSeconds)s; backend=$($errorRow.Backend)) : $($errorRow.Message)"
        WriteLog -Message $errMsg "WARNING"
        [void]$errors.Add($errorRow)
    }
}
Write-Progress -Id 0 -Activity $overallActivity -Completed

# ------------------------- Export & Cleanup -------------------------
$BaseFileName = "Exchange_OnPrem_MailboxCalendarPermissions_AllDomains"
$weeklyHistorySourcePaths = New-Object 'System.Collections.Generic.List[string]'

Write-Host "`n--- Export CSV ---"
if ($results.Count -gt 0) {
    $requiredColumns = @('Mailbox','UPN','CalendarFolder','User','AccessRights')
    $missingColumns = @($requiredColumns | Where-Object { -not $results[0].PSObject.Properties[$_] })
    if ($missingColumns.Count -gt 0) {
        throw ("Calendar permissions export schema is incomplete. Missing column(s): {0}" -f ($missingColumns -join ', '))
    }
    $exportRows = @($results | Select-Object $requiredColumns)
    ExportAndCopyCsv -BaseFileName $BaseFileName `
        -OutputPath $OutputPath `
        -GlobalPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') `
        -Data $exportRows `
        -Encoding "UTF8" `
        -NoTypeInformation `
        -NoMaxItemsRowLimit `
        -SkipWeeklyHistory
    if (-not [string]::IsNullOrWhiteSpace([string]$global:csvFilePath3)) {
        [void]$weeklyHistorySourcePaths.Add([string]$global:csvFilePath3)
    }
} else {
    WriteLog -Message "No data to export (no calendars found or all skipped). Export step skipped." "INFO"
    Write-Host "No data to export. Skipping."
}

$errorBaseFileName = "Exchange_OnPrem_MailboxCalendarPermissions_Errors"
$errorLatestCsvFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''
if ($errors.Count -gt 0) {
    Add-Content -Path $logTextFile -Value "`n=== MAILBOX-LEVEL ERRORS ==="
    $errors | ForEach-Object {
        Add-Content -Path $logTextFile -Value ("[{0}] {1} | {2} | attempts={3} | duration={4}s | backend={5} | {6}" -f $_.Category, $_.Mailbox, $_.Operation, $_.Attempts, $_.DurationSeconds, $_.Backend, $_.Message)
    }

    $errorRows = @($errors | Select-Object Index,Mailbox,UPN,Category,Operation,Attempts,DurationSeconds,Backend,Database,Message)

    ExportAndCopyCsv -BaseFileName $errorBaseFileName `
        -OutputPath $OutputPath `
        -GlobalPath $errorLatestCsvFolderPath `
        -Data $errorRows `
        -Encoding "UTF8" `
        -NoTypeInformation `
        -SkipWeeklyHistory
    if (-not [string]::IsNullOrWhiteSpace([string]$global:csvFilePath3)) {
        [void]$weeklyHistorySourcePaths.Add([string]$global:csvFilePath3)
    }

    WriteLog -Message ("Calendar permissions completed with {0} mailbox-level error(s). CSV export was produced and final status is CompletedWithWarnings." -f $errors.Count) "INFO"
}
else {
    $errorRunBaseFileName = Add-SmartM365MaxItemsSuffixToBaseName -BaseFileName $errorBaseFileName
    $staleErrorLatestPaths = @(
        Join-Path -Path $OutputPath -ChildPath "$errorRunBaseFileName.csv"
        if (-not [string]::IsNullOrWhiteSpace($errorLatestCsvFolderPath)) {
            Join-Path -Path $errorLatestCsvFolderPath -ChildPath "$errorRunBaseFileName.csv"
        }
    ) | Select-Object -Unique
    foreach ($staleErrorLatestPath in $staleErrorLatestPaths) {
        if (Test-Path -LiteralPath $staleErrorLatestPath -PathType Leaf) {
            try {
                Remove-Item -LiteralPath $staleErrorLatestPath -Force -ErrorAction Stop
                WriteLog -Message "Removed stale error CSV after successful run: $staleErrorLatestPath" "INFO"
            }
            catch {
                WriteLog -Message "Unable to remove stale error CSV '$staleErrorLatestPath': $($_.Exception.Message)" "WARNING"
            }
        }
    }
    $sharePointStaleErrorPath = if (-not [string]::IsNullOrWhiteSpace($errorLatestCsvFolderPath)) {
        Join-Path -Path $errorLatestCsvFolderPath -ChildPath "$errorRunBaseFileName.csv"
    }
    else {
        Join-Path -Path $OutputPath -ChildPath "$errorRunBaseFileName.csv"
    }
    Remove-SmartM365SharePointFile -LocalFilePath $sharePointStaleErrorPath | Out-Null
}

Publish-SmartM365CalendarWeeklyHistory -SourceCsvPaths $weeklyHistorySourcePaths.ToArray() -Config $ScriptLocalConfig -FallbackRootPath $OutputPath

Write-Host "`n=== SUMMARY ===" -ForegroundColor Cyan
Write-Host "PrimaryOnly mode      : $PrimaryOnly"
Write-Host "EmitNoPermRow         : $EmitNoPermRow"
Write-Host "Mailboxes processed   : $processed"
Write-Host "Permissions rows      : $($results.Count)"
Write-Host "Errors                : $($errors.Count)"
Write-Host "Export completed (if any)."
Write-Host "- Log              : $global:logTextFile"

# Clean up old CSV files + old log files
# Automatically excludes all generated CSVs via global:csvGeneratedPaths + current transcript and log files via global variables
Remove-SmartM365TimestampedFilesOlderThan -FolderPath $OutputPath -FilePattern '*.csv' -RetentionDays 7 -LogFile $global:logTextFile
RemoveOldFiles -Path $logPath -Filter "*.log" -KeepCount $global:RetentionMaxLogs -LogFile $global:logTextFile
WriteLog -Message "$TaskName completed."
Stop-Transcript | Out-Null
try { if ($global:logTranscriptFile) { Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile } } catch {}
$finalStatus = if ($errors.Count -gt 0) { 'CompletedWithWarnings' } else { 'Auto' }
Complete-SmartM365ExecutionContext -Status $finalStatus

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA26AdHXmCDmRCE
# ve9DIMzjcLVEqtFdyz3F39sBWiBLjKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINuAgo/enIyXdxjVQJ+ghdZhsISNzt586IukcyBaQQecMA0GCSqG
# SIb3DQEBAQUABIIBgKrhpTuM1G6DhSnfCVch6vAj5KbE5JapSBqwrD5bmVgmTzWa
# 3Z8TnK4lpnEAC+VCrTGzYYtreze6ofwsPKOXWodwD2fpABbXPiH8BQs3xP1p8FzA
# D64UarnVxyOYQJ8nuqg8SNM8JH+IIhgNn9czC8ImgbSKTwjRyVi1DJla0dE3zBWB
# SO1R7sBFVEGUbHqtBxXRHtzGZEeOV4/oMSkdTi8cb5FOweay6PFSqbdes+6cKaCn
# tZrhGZ9ICJZKZB+lMmw4upYGxUGPUSdJUjkGKWcU0e53cKqQqFoa0e+lgCNq6KiR
# XyZ4Dx8DffNr8vJqXDYvbZ8m98q9K83vdc2Y5vuTJLdCsqUZqzKOhTyOgzSqbW65
# wpt9PH6k/+9p10eEwHgCslyqXkXb6WNLri+wIEzIsscmJaMCeKX5S5hCHVzsGKa6
# w7DhWEmgTZAm6IVwRwPldbV7gjm5TLRE6vVXL0VDL3WnMmRYLOFcPDl4AfXan2Q5
# s/5oGnNzlzED9z6yDKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjQwODQy
# NDdaMC8GCSqGSIb3DQEJBDEiBCDTN840LwsElrs14D8LDoe034kkv2xe7UYUhbor
# UyLCfzANBgkqhkiG9w0BAQEFAASCAgCPA5dlEndifR7ZjgceqcUt/VNNQHlcDOjJ
# 6vuXEBfpXHiWYGllYMT5wXfisN3htkn9FzKB5rZVXIC753SFsjz7gUtz7BOA3bC+
# NGL8VGg0C2EfUcOXLS23uxf3IXYzVAWWAh1XjIVMrHBugH+S/6CLceqjDinuccq5
# PrzqG6UvVbWR01zowODMjcD9lz/07SoixAk2EjvE9jS1PCFWv8oyAFrT6Fl7YbyK
# V+dT0KwBaGGDm8Zs4qEEShipE9Pu6mBVO76kUrucj0GX09+7oP+q1GQ5LfOjYWJp
# BhgVj1cFlb/e1Jl2mUjqEobnnIBPdZ1hZIF334IWBi7HK0807XFIOTwIoMcXW89M
# ivLmNx+TSQcVn+OwDvvczHQ6lIiW1Lqj04OPpuunGLHjAnDLJzPcxK6bOgO8koKK
# Q4zq3bGT1ykrcSlAYQHUNQ+eae6EBBWaA+ULVUik1JElkOrBKgu4jvkUo+rOjnUY
# Z3fvXPmZyCX2tpaY9vFTFmRqEBjZmjEEzL8fpE9f8vPCdMh8CohvB5F18HQFZ8+v
# ax92zh32UuEbnqCpkZxnaMDEg/DfH571J+iclK/cPPm7/cslTmH1YiuYaJcV23I3
# vZV5zujQK41GRmC66KSo4fvzsT9thWK2ypfFT4UCTVKfKhhrlgT6Tmo7MbhtGt2J
# H6iyJtuVEw==
# SIG # End signature block
