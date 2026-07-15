<#
.SYNOPSIS
    Inventories Exchange Online mailbox calendar permissions.

.DESCRIPTION
    Retrieves calendar folder permissions for Exchange Online user, shared, room, and equipment mailboxes.
    Primary-calendar mode uses persistent parallel Exchange Online worker sessions; full-folder mode remains sequential.
    The script tries the canonical Calendar name, folder statistics, then common localized names.
    It exports the stable Mailbox, UPN, CalendarFolder, User, and AccessRights schema.

.PARAMETER Connect
    Disconnects any existing Exchange Online session before establishing the app-only connection.

.PARAMETER PrimaryOnly
    Scans only the main Calendar folder by default. Disable it to scan every Calendar folder.

.PARAMETER EmitNoPermRow
    Emits one "(none)" row when a calendar exists but has no explicit permissions.

.PARAMETER TopMailboxes
    Limits mailbox processing to the first N mailboxes for smoke tests. Default 0 processes all mailboxes.

.VERSION
2.2

.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; ExchangeOnlineManagement.
    Minimum permissions: Exchange.ManageAsApp plus Exchange Online app-only RBAC allowing mailbox and calendar-folder permission reads; Global Reader is the default read-only service-principal role.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled; Mail.Send is required only when Graph mail is enabled.

.NOTES
    Version : 2.2
    Author: https://github.com/khda79/workplacecloudhub.com
    Environment : Exchange Online
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [switch]$Connect,
    [switch]$PrimaryOnly = $true,
    [switch]$EmitNoPermRow = $true,
    [int]$TopMailboxes = 0,
    [ValidateRange(1,20)][int]$ParallelThrottle = 4,
    [ValidateRange(1,10)][int]$PermissionWorkerConnectRetries = 3,
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
$Thumb = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumb' -DefaultValue '0000000000000000000000000000000000000000'
$OrgDomain = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue 'contoso.onmicrosoft.com'
function Join-ModulePath {
    param([Parameter(Mandatory)][string]$FileName)
    $searchRoot = $PSScriptRoot
    while ($searchRoot) {
        $candidate = Join-Path -Path (Join-Path -Path (Join-Path -Path $searchRoot -ChildPath 'Modules') -ChildPath 'SmartM365.Core') -ChildPath $FileName
        if (Test-Path -LiteralPath $candidate) { return $candidate }
        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }
    throw "SmartM365.Core module file not found: $FileName"
}

$ScriptVersion = '2.2'
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ExoCalendarPermissionsCsvLogFolderPath' -DefaultValue $OutputPath
$TaskName = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."

try {
    Write-Host 'Loading module SmartM365.Core.psd1...' -ForegroundColor Cyan
    Import-Module -Name (Join-ModulePath 'SmartM365.Core.psd1') -MinimumVersion '1.0.36' -ErrorAction Stop
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPath $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
    Start-Transcript -Path $global:logTranscriptFile -Append
    $logTextFile = Join-Path $logPath "$(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')-EXO-$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"
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

function Ensure-ExchangeOnlineModule {
    $m = Get-Module -ListAvailable -Name "ExchangeOnlineManagement" | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $m) {
        throw "Required module 'ExchangeOnlineManagement' is not installed. Install it with: Install-Module ExchangeOnlineManagement"
    }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
}
try {
    Ensure-ExchangeOnlineModule
    if ($Connect) {
        Write-Host 'Connect switch specified: existing Exchange Online session will be disconnected and reconnected...' -ForegroundColor Cyan
    }
    Disconnect-SmartM365CloudSession -ExchangeOnline $true -Graph $false -VerboseDisconnect:$true
    WriteLog -Message 'Connecting to Exchange Online with app-only certificate authentication.' 'INFO'
    Connect-ExchangeOnline -AppId $AppId -CertificateThumbprint $Thumb -Organization $OrgDomain -ShowBanner:$false -ErrorAction Stop | Out-Null
    WriteLog -Message 'Connected to Exchange Online.'
}
catch {
    WriteLog -Message ("Failed to connect to Exchange Online: {0}" -f $_.Exception.Message) 'ERROR'
    Stop-Transcript | Out-Null
    try { if ($global:logTranscriptFile) { Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile } } catch {}
    Complete-SmartM365ExecutionContext -Status Auto
    exit 1
}

Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -ExchangeOnlineProbeCommands @('Get-Mailbox') | Out-Null
$results = [System.Collections.Generic.List[object]]::new()
$errors = @()
$processed = 0

try {
    $recipientTypes = @('UserMailbox','SharedMailbox','RoomMailbox','EquipmentMailbox')
    if (Test-HasCommand -Name 'Get-EXOMailbox') {
        $mailboxes = Invoke-Quiet {
            Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails $recipientTypes -ErrorAction Stop
        } | Select-Object DisplayName, PrimarySmtpAddress, UserPrincipalName, Identity, ExchangeGuid
    }
    else {
        $mailboxes = Invoke-Quiet {
            Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails $recipientTypes -ErrorAction Stop
        } | Select-Object DisplayName, PrimarySmtpAddress, UserPrincipalName, Identity, ExchangeGuid
    }
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

function Get-CalendarFoldersSafe {
    param(
        [Parameter(Mandatory = $true)]$Mbx,
        [Parameter(Mandatory = $true)][bool]$PrimaryOnly
    )
    $id = $Mbx.ExchangeGuid
    if (-not $id -or $id -eq [guid]::Empty) { $id = $Mbx.UserPrincipalName }
    if (-not $id) { $id = $Mbx.PrimarySmtpAddress }
    $folders = Invoke-Quiet {
        if (Test-HasCommand -Name 'Get-EXOMailboxFolderStatistics') {
            Get-EXOMailboxFolderStatistics -Identity $id -FolderScope Calendar -ErrorAction Stop
        }
        else {
            Get-MailboxFolderStatistics -Identity $id -FolderScope Calendar -ErrorAction Stop
        }
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
        [Parameter(Mandatory)][string]$UpnForLog
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

# Helper: emit a "(none)" row when calendar exists but no explicit permissions
function Add-NoPermissionRow {
    param(
        [Parameter(Mandatory)][string]$Mailbox,
        [Parameter(Mandatory)][string]$UPN,
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

if ($PrimaryOnly -and $ParallelThrottle -gt 1) {
    WriteLog -Message ("Starting parallel primary-calendar permission collection with {0} persistent worker(s)." -f $ParallelThrottle) "INFO"

    $workerCount = [math]::Min($ParallelThrottle, [math]::Max(1, $mailboxes.Count))
    $calendarWorkItems = @()
    if ($mailboxes.Count -gt 0) {
        $chunkSize = [int][math]::Ceiling($mailboxes.Count / [double]$workerCount)
        for ($workerIndex = 0; $workerIndex -lt $workerCount; $workerIndex++) {
            $startIndex = $workerIndex * $chunkSize
            if ($startIndex -ge $mailboxes.Count) { break }
            $endIndex = [math]::Min($startIndex + $chunkSize - 1, $mailboxes.Count - 1)
            $calendarWorkItems += [pscustomobject]@{
                WorkerId  = $workerIndex + 1
                Mailboxes = @($mailboxes[$startIndex..$endIndex])
            }
        }
    }

    $p_AppId = $AppId
    $p_Thumb = $Thumb
    $p_OrgDomain = $OrgDomain
    $p_ConnectRetries = $PermissionWorkerConnectRetries
    $p_EmitNoPermRow = [bool]$EmitNoPermRow

    $parallelCalendarOutput = @($calendarWorkItems | ForEach-Object -ThrottleLimit $workerCount -Parallel {
        $workItem = $_
        $workerId = [int]$workItem.WorkerId
        $workerMailboxes = @($workItem.Mailboxes)
        $connected = $false
        $connectionError = ''

        Import-Module ExchangeOnlineManagement -ErrorAction Stop
        for ($connectAttempt = 1; $connectAttempt -le $using:p_ConnectRetries -and -not $connected; $connectAttempt++) {
            try {
                $workerConnectParams = @{
                    AppId                 = $using:p_AppId
                    CertificateThumbprint = $using:p_Thumb
                    Organization          = $using:p_OrgDomain
                    ShowBanner            = $false
                    ShowProgress          = $false
                    ErrorAction           = 'Stop'
                }
                Connect-ExchangeOnline @workerConnectParams | Out-Null
                $connected = $true
            }
            catch {
                $connectionError = $_.Exception.Message
                if ($connectAttempt -lt $using:p_ConnectRetries) {
                    Start-Sleep -Seconds ([math]::Min(30, 5 * $connectAttempt))
                }
            }
        }

        function Invoke-WorkerFolderPermissionQuery {
            param(
                [string[]]$MailboxIds,
                [string[]]$FolderNames,
                [string]$PrimarySmtp,
                [string]$UserPrincipalName
            )

            foreach ($mailboxId in @($MailboxIds | Where-Object { $_ } | Select-Object -Unique)) {
                foreach ($folderName in @($FolderNames | Where-Object { $_ } | Select-Object -Unique)) {
                    $identity = '{0}:\{1}' -f $mailboxId, $folderName
                    try {
                        $permissionRows = @(Get-MailboxFolderPermission -Identity $identity -ErrorAction Stop -WarningAction SilentlyContinue 6>$null |
                            Where-Object { [string]$_.User -notin @('Default','Anonymous') } |
                            ForEach-Object {
                                [pscustomobject]@{
                                    Mailbox        = $PrimarySmtp
                                    UPN            = $UserPrincipalName
                                    CalendarFolder = $folderName
                                    User           = [string]$_.User
                                    AccessRights   = ($_.AccessRights -join ',')
                                }
                            })
                        return [pscustomobject]@{ Found = $true; FolderName = $folderName; Rows = $permissionRows }
                    }
                    catch {}
                }
            }
            return [pscustomobject]@{ Found = $false; FolderName = ''; Rows = @() }
        }

        foreach ($mbx in $workerMailboxes) {
            $primarySmtp = [string]$mbx.PrimarySmtpAddress
            $upn = [string]$mbx.UserPrincipalName

            if (-not $connected) {
                [pscustomobject]@{
                    WorkerId = $workerId
                    Mailbox = $primarySmtp
                    Rows = @()
                    Warning = ''
                    Error = ("Worker {0} EXO connection failed after {1} attempt(s): {2}" -f $workerId, $using:p_ConnectRetries, $connectionError)
                    ConnectionFailure = $true
                }
                continue
            }

            try {
                $mailboxIds = @($upn, $primarySmtp)
                $query = Invoke-WorkerFolderPermissionQuery -MailboxIds $mailboxIds -FolderNames @('Calendar') -PrimarySmtp $primarySmtp -UserPrincipalName $upn

                if (-not $query.Found) {
                    try {
                        $statsIdentity = [string]$mbx.ExchangeGuid
                        if ([string]::IsNullOrWhiteSpace($statsIdentity) -or $statsIdentity -eq [guid]::Empty.ToString()) { $statsIdentity = $upn }
                        if ([string]::IsNullOrWhiteSpace($statsIdentity)) { $statsIdentity = $primarySmtp }
                        $calendarFolders = @(Get-EXOMailboxFolderStatistics -Identity $statsIdentity -FolderScope Calendar -ErrorAction Stop -WarningAction SilentlyContinue 6>$null |
                            Where-Object { $_.FolderType -eq 'Calendar' })
                        $calendarFolder = $calendarFolders | Where-Object { $_.FolderPath -match '^/[^/]+$' } |
                            Sort-Object ItemsInFolder -Descending | Select-Object -First 1
                        if (-not $calendarFolder) { $calendarFolder = $calendarFolders | Select-Object -First 1 }
                        if ($calendarFolder) {
                            $folderPath = ([string]$calendarFolder.FolderPath).TrimStart('/')
                            $query = Invoke-WorkerFolderPermissionQuery -MailboxIds $mailboxIds -FolderNames @($folderPath) -PrimarySmtp $primarySmtp -UserPrincipalName $upn
                        }
                    }
                    catch {}
                }

                if (-not $query.Found) {
                    $query = Invoke-WorkerFolderPermissionQuery -MailboxIds $mailboxIds -FolderNames @('Calendrier','Kalender','Calendario') -PrimarySmtp $primarySmtp -UserPrincipalName $upn
                }

                $rows = @($query.Rows)
                $warning = ''
                if ($query.Found -and $rows.Count -eq 0 -and $using:p_EmitNoPermRow) {
                    $rows = @([pscustomobject]@{
                        Mailbox = $primarySmtp
                        UPN = $upn
                        CalendarFolder = [string]$query.FolderName
                        User = '(none)'
                        AccessRights = '(none)'
                    })
                }
                elseif (-not $query.Found) {
                    $warning = "No calendar folder found for $primarySmtp"
                }

                [pscustomobject]@{
                    WorkerId = $workerId
                    Mailbox = $primarySmtp
                    Rows = $rows
                    Warning = $warning
                    Error = ''
                    ConnectionFailure = $false
                }
            }
            catch {
                [pscustomobject]@{
                    WorkerId = $workerId
                    Mailbox = $primarySmtp
                    Rows = @()
                    Warning = ''
                    Error = ("Error for {0}: {1}" -f $primarySmtp, $_.Exception.Message)
                    ConnectionFailure = $false
                }
            }
        }
        # Worker disconnects are intentionally omitted; one parent cleanup closes every EXO session.
    })

    $connectionFailures = @($parallelCalendarOutput | Where-Object { $_.ConnectionFailure })
    if ($connectionFailures.Count -gt 0) {
        Disconnect-SmartM365CloudSession -ExchangeOnline $true -Graph $false -VerboseDisconnect:$false
        throw ("Parallel calendar permission collection lost {0} mailbox row(s) to worker connection failures; publication refused." -f $connectionFailures.Count)
    }
    if ($parallelCalendarOutput.Count -ne $mailboxes.Count) {
        Disconnect-SmartM365CloudSession -ExchangeOnline $true -Graph $false -VerboseDisconnect:$false
        throw ("Parallel calendar permission mailbox-count mismatch: dispatched={0}, collected={1}." -f $mailboxes.Count, $parallelCalendarOutput.Count)
    }

    foreach ($calendarOutput in $parallelCalendarOutput) {
        Add-CalendarResultRows -Rows @($calendarOutput.Rows)
        if (-not [string]::IsNullOrWhiteSpace([string]$calendarOutput.Warning)) {
            WriteLog -Message ([string]$calendarOutput.Warning) "WARNING"
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$calendarOutput.Error)) {
            WriteLog -Message ([string]$calendarOutput.Error) "WARNING"
            $errors += [string]$calendarOutput.Error
        }
    }
    $processed = $parallelCalendarOutput.Count
    WriteLog -Message ("Parallel calendar permission collection completed with exact mailbox parity: {0}/{1}; permission rows={2}." -f $processed, $mailboxes.Count, $results.Count) "INFO"
}
else {
foreach ($mbx in $mailboxes) {
    $index++; $processed++
    $primarySMTP = $mbx.PrimarySmtpAddress.ToString()
    $upn         = $mbx.UserPrincipalName
    $percent     = [int](($index / $total) * 100)

    Write-Progress -Id 0 -Activity $overallActivity -Status ("[{0}/{1}] {2}" -f $index, $total, $primarySMTP) -PercentComplete $percent

    try {
        if ($PrimaryOnly) {
            $mailboxIds = @($upn, $primarySMTP)

            # 1) Canonical "Calendar"
            $permissionResult = Try-GetFolderPermission -MailboxIds $mailboxIds -FolderNames @('Calendar') -PrimarySmtpForLog $primarySMTP -UpnForLog $upn
            $ok = [bool]$permissionResult.Ok
            $permissions = @($permissionResult.Permissions)
            $usedIdentity = $permissionResult.Identity
            if ($ok) {
                WriteLog -Message "Calendar folder found for $primarySMTP (via canonical : $usedIdentity)" "INFO"
                if ($permissions -and $permissions.Count -gt 0) {
                    Add-CalendarResultRows -Rows $permissions
                } elseif ($EmitNoPermRow) {
                    Add-CalendarResultRows -Rows (Add-NoPermissionRow -Mailbox $primarySMTP -UPN $upn -CalendarFolder 'Calendar')
                }
            } else {
                # 2) Stats fallback
                $calendarFolders = Get-CalendarFoldersSafe -Mbx $mbx -PrimaryOnly:$true
                if ($calendarFolders -and $calendarFolders.Count -gt 0) {
                    $folderPath = $calendarFolders[0].FolderPath.TrimStart("/")
                    $permissionResult2 = Try-GetFolderPermission -MailboxIds $mailboxIds -FolderNames @($folderPath) -PrimarySmtpForLog $primarySMTP -UpnForLog $upn
                    $ok2 = [bool]$permissionResult2.Ok
                    $permissions2 = @($permissionResult2.Permissions)
                    $usedIdentity2 = $permissionResult2.Identity
                    if ($ok2) {
                        WriteLog -Message "Calendar folder found for $primarySMTP (via stats : $usedIdentity2)" "INFO"
                        if ($permissions2 -and $permissions2.Count -gt 0) {
                            Add-CalendarResultRows -Rows $permissions2
                        } elseif ($EmitNoPermRow) {
                            Add-CalendarResultRows -Rows (Add-NoPermissionRow -Mailbox $primarySMTP -UPN $upn -CalendarFolder $folderPath)
                        }
                    } else {
                        # 3) Common localized names
                        $permissionResult3 = Try-GetFolderPermission -MailboxIds $mailboxIds -FolderNames @('Calendrier','Kalender','Calendario') -PrimarySmtpForLog $primarySMTP -UpnForLog $upn
                        $ok3 = [bool]$permissionResult3.Ok
                        $permissions3 = @($permissionResult3.Permissions)
                        $usedIdentity3 = $permissionResult3.Identity
                        if ($ok3) {
                            WriteLog -Message "Calendar folder found for $primarySMTP (via common localized name : $usedIdentity3)" "INFO"
                            if ($permissions3 -and $permissions3.Count -gt 0) {
                                Add-CalendarResultRows -Rows $permissions3
                            } elseif ($EmitNoPermRow) {
                                # extract folder name from identity "mbId:\Name"
                                $folderName = ($usedIdentity3 -split ':\s*',2)[1]
                                Add-CalendarResultRows -Rows (Add-NoPermissionRow -Mailbox $primarySMTP -UPN $upn -CalendarFolder $folderName)
                            }
                        } else {
                            WriteLog -Message "No calendar folder found for $primarySMTP" "WARNING"
                        }
                    }
                } else {
                    # Last try with localized names if stats empty
                    $permissionResult4 = Try-GetFolderPermission -MailboxIds $mailboxIds -FolderNames @('Calendrier','Kalender','Calendario') -PrimarySmtpForLog $primarySMTP -UpnForLog $upn
                    $ok4 = [bool]$permissionResult4.Ok
                    $permissions4 = @($permissionResult4.Permissions)
                    $usedIdentity4 = $permissionResult4.Identity
                    if ($ok4) {
                        WriteLog -Message "Calendar folder found for $primarySMTP (via common localized name : $usedIdentity4)" "INFO"
                        if ($permissions4 -and $permissions4.Count -gt 0) {
                            Add-CalendarResultRows -Rows $permissions4
                        }
                        elseif ($EmitNoPermRow) {
                            $folderName = ($usedIdentity4 -split ':\s*',2)[1]
                            Add-CalendarResultRows -Rows (Add-NoPermissionRow -Mailbox $primarySMTP -UPN $upn -CalendarFolder $folderName)
                        }
                    } else {
                        WriteLog -Message "No calendar folder found for $primarySMTP" "WARNING"
                    }
                }
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
                    $permMailboxId = $upn
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
                    WriteLog -Message $errMsg "WARNING"
                    $errors += $errMsg
                }
            }
        }
    } catch {
        $errMsg = "Error for $primarySMTP : $($_.Exception.Message)"
        WriteLog -Message $errMsg "WARNING"
        $errors += $errMsg
    }
}
}
Write-Progress -Id 0 -Activity $overallActivity -Completed

# ------------------------- Export & Cleanup -------------------------
$BaseFileName = "Exchange_EXO_MailboxCalendarPermissions_AllDomains"

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
        -NoMaxItemsRowLimit
} else {
    WriteLog -Message "No data to export (no calendars found or all skipped). Export step skipped." "INFO"
    Write-Host "No data to export. Skipping."
}

$errorBaseFileName = "Exchange_EXO_MailboxCalendarPermissions_Errors"
$errorLatestCsvFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''
if ($errors.Count -gt 0) {
    Add-Content -Path $logTextFile -Value "`n=== MAILBOX-LEVEL ERRORS ==="
    $errors | ForEach-Object { Add-Content -Path $logTextFile -Value $_ }

    $errorRows = for ($i = 0; $i -lt $errors.Count; $i++) {
        [pscustomobject]@{
            Index   = $i + 1
            Message = [string]$errors[$i]
        }
    }

    ExportAndCopyCsv -BaseFileName $errorBaseFileName `
        -OutputPath $OutputPath `
        -GlobalPath $errorLatestCsvFolderPath `
        -Data $errorRows `
        -Encoding "UTF8" `
        -NoTypeInformation

    WriteLog -Message ("Calendar permissions completed with {0} mailbox-level error(s). CSV export was produced, but final status must be CompletedWithWarnings." -f $errors.Count) "WARNING"
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
RemoveOldFiles -Path $OutputPath -Filter "*.csv" -KeepCount $global:RetentionMaxCSV -LogFile $global:logTextFile
RemoveOldFiles -Path $logPath -Filter "*.log" -KeepCount $global:RetentionMaxLogs -LogFile $global:logTextFile
WriteLog -Message "$TaskName completed."
Disconnect-SmartM365CloudSession -ExchangeOnline $true -Graph $false -VerboseDisconnect:$false
Stop-Transcript | Out-Null
try { if ($global:logTranscriptFile) { Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile } } catch {}
$finalStatus = if ($errors.Count -gt 0) { 'CompletedWithWarnings' } else { 'Auto' }
Complete-SmartM365ExecutionContext -Status $finalStatus

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCJtBiFTTQdgYEB
# S9FZGiUGhoPysq+Hl3bfIJeECB8vIaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOAchulBcrCgt6NsaQeH3FJXEjaOkG/A8LD3CTOduW+5MA0GCSqG
# SIb3DQEBAQUABIIBgERa0Pd3Wf+fIDlA1LdZa/u5iwlGkYMKl1Tbly0Z+v+4XiBC
# db8h96gXHocEfDIWJTYlINl/9OhdQZPU722ho7UUWhdH+F8AabTrgJOYmwYAOiDb
# 1hA0KEnNB6TJ5zA2YCo2oC5OFnq6+RhmZFvRGtRr/aQiftPgHYMPcJ3GZ60Azaf0
# 3fxKFEAs4bu+RgOK7y/PYolJm7emEAohPeoFimUXP7ypj5GRrdcNFKBn6IJXgVXa
# l9nTOtOOYuktZnqRd1ZRzg+vfa0RenQzxz2aF4GfutQHcmvJ9JjSf8KVNGIerSk9
# FiaUXpf5IBlVcoSH7DO6XDqc7Hv+g4QBc6JXJRzVsVPG0ysOjJu9N4bs2k5q/0td
# qe4jGBqKLMnlrMWvyRlSGlRkIKrxIkDSgkw248f81Pd1aWD+1OYMyOdN+pY3hglT
# 5cdHw7l/vkUEs3t1qEx/GJ3EumEYLs359OLYfcmatvtaL8LGnQRYkZFVuBT6lx9m
# wbOEYWImjPX493sv26GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTUxMzM2
# MDBaMC8GCSqGSIb3DQEJBDEiBCDEIyKua80sgHBELbc8+AE2T1i9Pej/RihD8HJf
# NHFR6TANBgkqhkiG9w0BAQEFAASCAgBCjdLBJmHooZQ5yQwP7/PN14ZQrhkW61EV
# l2AM96e4Y0rglt0kMB99qoppHoCLVVEwWeXniPAwn3i7m84QTAeQpNXaURtYiHG/
# 95HJhkcVJPmfKfM/BJYL7HXNNUuDVSJGDmC9K8tpvn1S61A69udxGe3DF1hfnNbS
# V8RMAknohGkgMQOhnDoaxgYwtt/Vrrt1ByaKrnK+lC7ln0a7aCIMIXcDBskxfRT8
# LWHlCtmYHHrtdA7nsMk/kQpBgNJK8p+7khgJn/5C6IH/Y3v7qBamIva2MXM3FRFb
# Im6V7kGjRyzPZ6gNJ/RKhPJRmWSlIJpmmDnmDKD5hIfpU29v3l1nCxvjWWCB83HP
# +xfwV9fswcIH5/oIefTQThKpjgm0wu7vN4O4xETV021rIwQA/pkHEPtH1+Z4B7ZB
# UG+TX1CHmvQ0PfjlInAiEp0tilkNv3Ea8AMRLeRn3AsCYucOgJA9rtkXMZXsXPbq
# X3DC7govyngbU81O4sIHYd0FVr/6HNEPwmA88u/urRmY9/cYbT2txz1eOzTHhC/T
# 0auH1h1jftJ6wf69nqGwlzrMf9WN9WKjHBGY5lErESq+2oGlwQqLPVL97ZWNEYCY
# 4MkP2Z0kJ9nbHBK24O/ntCrHu8d+1VwL3KChqeClEOhjbRfjOsTFDMfIhJSub1c3
# XuVV8yPbkg==
# SIG # End signature block
