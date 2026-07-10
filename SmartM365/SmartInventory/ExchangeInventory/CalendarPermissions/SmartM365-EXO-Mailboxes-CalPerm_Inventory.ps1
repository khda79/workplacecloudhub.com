<#
.SYNOPSIS
    Inventory Exchange mailbox calendar permissions (Exchange 2016 On-Prem or Exchange Online).

.DESCRIPTION
    Retrieves calendar folder permissions for all mailboxes.
    - On-Prem (Exchange 2016): Get-Mailbox / Get-MailboxFolderStatistics.
    - Online (EXO): Prefer Get-EXOMailbox / Get-EXOMailboxFolderStatistics (v3 accelerated cmdlets),
                    fall back to classic Get-Mailbox / Get-MailboxFolderStatistics if v3 unavailable.
    Uses robust identities (ExchangeGuid/Guid with fallbacks to UPN/SMTP) and exports CSV with UPN column.
    Suppresses EXO v3 information noise via Invoke-Quiet. Adds informative logging.
    Can emit a row with "(none)" when a calendar exists but has no explicit permissions (see -EmitNoPermRow).

.PARAMETER Online
    Run against Exchange Online when specified. Otherwise runs against On-Prem Exchange 2016.

.PARAMETER Connect
    Force (re-)connection when running in Online mode.

.PARAMETER PrimaryOnly
    When enabled (default), scans only the main Calendar folder. It first targets the canonical path ":\Calendar",
    then falls back to statistics-based detection (root-level Calendar) if needed. When disabled, scans all Calendar folders.

.PARAMETER EmitNoPermRow
    When enabled (default), if a calendar is found but has NO explicit permissions (only Default/Anonymous),
    emit a single row with User="(none)" and AccessRights="(none)" so the CSV is not empty.

.PARAMETER InteractiveAuth
    Uses interactive authentication for Microsoft Graph instead of app-only certificate authentication.

.PARAMETER TopMailboxes
    Limits mailbox processing to the first N mailboxes for smoke tests. Default 0 processes all mailboxes.
.VERSION
1.5


.NOTES
    Version : 1.1
    Author: https://github.com/khda79/workplacecloudhub.com
    Environment : Hybrid (Online & On-Prem)
#>

param(
    [string]$Tenant = 'test',
    [switch]$Online = $true,
    [switch]$Connect,
    [switch]$PrimaryOnly = $true,
    [switch]$EmitNoPermRow = $true,
    [switch]$InteractiveAuth,
    [int]$TopMailboxes = 0,
    [string]$OutputPath
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
# App-only authentication parameters (same app as other inventory scripts)
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

# ------------------------- Init & Environment -------------------------
$ScriptVersion = "1.5"
if ($Online) {
    $OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ExoCalendarPermissionsCsvLogFolderPath' -DefaultValue $OutputPath
    $TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
} else {
    $OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LocalCalendarPermissionsCsvLogFolderPath' -DefaultValue $OutputPath
    $TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
}

try {
    Write-Host "Loading module SmartM365.Core.psd1..." -ForegroundColor Cyan
    Import-Module (Join-ModulePath 'SmartM365.Core.psd1') -ErrorAction Stop

    # Standard SmartM365 initialization (logging, paths, transcript, etc.)
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPath $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
    Start-Transcript -Path $global:logTranscriptFile -Append

    if ($Online) {
        $logTextFile = Join-Path $logPath "$(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')-Online-$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"
    } else {
        $logTextFile = Join-Path $logPath "$(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')-OnPrem-$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"
    }

    WriteLog -Message "Script Environment initialized at $InitializeOutputPath"
    $OutputPath = $InitializeOutputPath
    WriteLog -Message "Starting $TaskName..."
} catch {
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

function Disconnect-ExchangeOnlineSafe {
    try {
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop | Out-Null
    }
    catch {
        WriteLog -Message ("Disconnect-ExchangeOnline failed (non-fatal): {0}" -f $_.Exception.Message) "WARN"
    }
}

if ($Online) {
    # -------------------------
    # Exchange Online connection (EXO)
    # -------------------------
    try {
        Ensure-ExchangeOnlineModule

        if ($Connect) {
            Write-Host "Connect switch specified: existing Exchange Online session will be disconnected and reconnected..." -ForegroundColor Cyan
        }

        Disconnect-ExchangeOnlineSafe

        WriteLog -Message "Connecting to Exchange Online with app-only certificate authentication." "INFO"
        Connect-ExchangeOnline `
            -AppId $AppId `
            -CertificateThumbprint $Thumb `
            -Organization $OrgDomain `
            -ShowBanner:$false `
            -ErrorAction Stop | Out-Null
        WriteLog -Message "Connected to Exchange Online."
    }
    catch {
        WriteLog -Message ("Failed to connect to Exchange Online: {0}" -f $_.Exception.Message) "ERROR"
        Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
        Complete-SmartM365ExecutionContext -Status Auto
        exit 1
    }

    # -------------------------
    # -------------------------
    # Detect existing Graph session (standard SmartM365 block)
    # -------------------------
    $graphContext = $null
    if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
        try {
            $graphContext = Get-MgContext -ErrorAction SilentlyContinue
        } catch { }
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
            GraphScopes    = @("Directory.Read.All")
        }

        if (-not $InteractiveAuth) {
            # Default: app-only certificate authentication
            $connectParams.AppId        = $AppId
            $connectParams.Thumbprint   = $Thumb
            $connectParams.TenantId     = $TenantId
            $connectParams.Organization = $OrgDomain
            WriteLog -Message "Connecting to Microsoft Graph with app-only certificate authentication." "INFO"
        } else {
            WriteLog -Message "Connecting to Microsoft Graph with interactive authentication." "INFO"
        }

        $connectResult = Connect-SmartM365CloudSession @connectParams

        if (-not $connectResult.GraphConnected) {
            WriteLog -Message "Failed to connect to Microsoft Graph." "ERROR"
            Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
            Complete-SmartM365ExecutionContext -Status Auto
            throw "Failed to connect to Microsoft Graph."
        }

        $connectedGraphInThisRun = $connectResult.GraphConnected
    }

} else {
    # -------------------------
    # On-Prem Exchange
    # -------------------------
    if (-not (Get-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction SilentlyContinue)) {
        try {
            Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
            Write-Host "Exchange On-Prem PSSnapin loaded." -ForegroundColor Green
            try {
                Set-ADServerSettings -ViewEntireForest $true -ErrorAction Stop
                WriteLog -Message "Set-ADServerSettings -ViewEntireForest $true applied."
            }
            catch {
                WriteLog -Message "Unable to apply Set-ADServerSettings : $($_.Exception.Message)" "WARN"
            }
        } catch {
            WriteLog -Message "Unable to load Exchange Management Shell. Check your environment." "ERROR"
            Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
            Complete-SmartM365ExecutionContext -Status Auto
            exit 1
        }
    } else {
        Write-Host "Exchange On-Prem PSSnapin detected." -ForegroundColor Green
    }
}

if ($Online) {
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -ExchangeOnlineProbeCommands @('Get-Mailbox') | Out-Null
}
else {
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -RequireExchangeOnPrem | Out-Null
}

# ------------------------- Data Structures -------------------------
$results   = @()
$errors    = @()
$processed = 0

# ------------------------- Mailbox Enumeration (robust + filtered) -------------------------
try {
    $recipientTypes = @("UserMailbox","SharedMailbox","RoomMailbox","EquipmentMailbox")
    if ($Online) {
        if (Test-HasCommand -Name "Get-EXOMailbox") {
            $mailboxes = Invoke-Quiet {
                Get-EXOMailbox -ResultSize Unlimited -RecipientTypeDetails $recipientTypes -ErrorAction Stop
            } | Select-Object DisplayName, PrimarySmtpAddress, UserPrincipalName, Identity, ExchangeGuid
        } else {
            $mailboxes = Invoke-Quiet {
                Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails $recipientTypes -ErrorAction Stop
            } | Select-Object DisplayName, PrimarySmtpAddress, UserPrincipalName, Identity, ExchangeGuid
        }
    } else {
        $mailboxes = Invoke-Quiet {
            Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails $recipientTypes -ErrorAction Stop
        } | Select-Object DisplayName, PrimarySmtpAddress, UserPrincipalName, Identity, Guid
    }
} catch {
    WriteLog -Message "Mailbox enumeration failed : $($_.Exception.Message)" "ERROR"
    Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
    Complete-SmartM365ExecutionContext -Status Auto
    exit 1
}

$total = $mailboxes.Count
WriteLog -Message "Total mailboxes found: $total"
if ($TopMailboxes -gt 0 -and $total -gt $TopMailboxes) {
    WriteLog -Message ("TopMailboxes enabled: processing first {0} of {1} mailboxes." -f $TopMailboxes, $total) "WARN"
    $mailboxes = @($mailboxes | Select-Object -First $TopMailboxes)
    $total = $mailboxes.Count
}
if ($total -eq 0) {
    WriteLog -Message "No mailboxes found. Stopping."
    Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
    Complete-SmartM365ExecutionContext -Status Auto
    exit 0
}

# ------------------------- Helpers -------------------------
function Get-CalendarFoldersSafe {
    param(
        [Parameter(Mandatory = $true)]$Mbx,
        [Parameter(Mandatory = $true)][bool]$IsOnline,
        [Parameter(Mandatory = $true)][bool]$PrimaryOnly
    )

    $tryEXO = {
        param($id)
        if (Test-HasCommand -Name "Get-EXOMailboxFolderStatistics") {
            Get-EXOMailboxFolderStatistics -Identity $id -FolderScope Calendar -ErrorAction Stop
        } else {
            Get-MailboxFolderStatistics -Identity $id -FolderScope Calendar -ErrorAction Stop
        }
    }
    $tryOnPrem = {
        param($id)
        Get-MailboxFolderStatistics -Identity $id -FolderScope Calendar -ErrorAction Stop
    }

    $folders = @()
    if ($IsOnline) {
        $id = $Mbx.ExchangeGuid
        if (-not $id -or $id -eq [guid]::Empty) { $id = $Mbx.UserPrincipalName }
        if (-not $id) { $id = $Mbx.PrimarySmtpAddress }
        $folders = Invoke-Quiet { & $tryEXO $id }
    } else {
        $id = $Mbx.Guid
        if (-not $id -or $id -eq [guid]::Empty) { $id = $Mbx.UserPrincipalName }
        if (-not $id) { $id = $Mbx.PrimarySmtpAddress }
        $folders = Invoke-Quiet { & $tryOnPrem $id }
    }

    if (-not $folders) { return @() }
    if (-not $PrimaryOnly) { return $folders | Where-Object { $_.FolderType -eq "Calendar" } }

    $rootCalendars = $folders | Where-Object { $_.FolderType -eq "Calendar" -and ($_.FolderPath -match '^/[^/]+$') }
    if ($rootCalendars -and $rootCalendars.Count -gt 0) {
        return ,($rootCalendars | Sort-Object ItemsInFolder -Descending | Select-Object -First 1)
    }

    $anyCal = $folders | Where-Object { $_.FolderType -eq "Calendar" } | Select-Object -First 1
    if ($anyCal) { return ,$anyCal }
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
                return @($true, $perms, $identity)  # success flag, data, used identity
            } catch {
                # Try next combination
            }
        }
    }
    return @($false, $null, $null)
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

# ------------------------- Processing Loop -------------------------
$overallActivity = "Calendar permissions inventory"
$index = 0

foreach ($mbx in $mailboxes) {
    $index++; $processed++
    $primarySMTP = $mbx.PrimarySmtpAddress.ToString()
    $upn         = $mbx.UserPrincipalName
    $percent     = [int](($index / $total) * 100)

    Write-Progress -Id 0 -Activity $overallActivity -Status ("[{0}/{1}] {2}" -f $index, $total, $primarySMTP) -PercentComplete $percent

    try {
        if ($PrimaryOnly) {
            $mailboxIds = if ($Online) { @($upn, $primarySMTP) } else { @($mbx.Identity, $primarySMTP) }

            # 1) Canonical "Calendar"
            $ok, $permissions, $usedIdentity = Try-GetFolderPermission -MailboxIds $mailboxIds -FolderNames @('Calendar') -PrimarySmtpForLog $primarySMTP -UpnForLog $upn
            if ($ok) {
                WriteLog -Message "Calendar folder found for $primarySMTP (via canonical : $usedIdentity)" "INFO"
                if ($permissions -and $permissions.Count -gt 0) {
                    $results += $permissions
                } elseif ($EmitNoPermRow) {
                    $results += (Add-NoPermissionRow -Mailbox $primarySMTP -UPN $upn -CalendarFolder 'Calendar')
                }
            } else {
                # 2) Stats fallback
                $calendarFolders = Get-CalendarFoldersSafe -Mbx $mbx -IsOnline:$Online -PrimaryOnly:$true
                if ($calendarFolders -and $calendarFolders.Count -gt 0) {
                    $folderPath = $calendarFolders[0].FolderPath.TrimStart("/")
                    $ok2, $permissions2, $usedIdentity2 = Try-GetFolderPermission -MailboxIds $mailboxIds -FolderNames @($folderPath) -PrimarySmtpForLog $primarySMTP -UpnForLog $upn
                    if ($ok2) {
                        WriteLog -Message "Calendar folder found for $primarySMTP (via stats : $usedIdentity2)" "INFO"
                        if ($permissions2 -and $permissions2.Count -gt 0) {
                            $results += $permissions2
                        } elseif ($EmitNoPermRow) {
                            $results += (Add-NoPermissionRow -Mailbox $primarySMTP -UPN $upn -CalendarFolder $folderPath)
                        }
                    } else {
                        # 3) Common localized names
                        $ok3, $permissions3, $usedIdentity3 = Try-GetFolderPermission -MailboxIds $mailboxIds -FolderNames @('Calendrier','Kalender','Calendario') -PrimarySmtpForLog $primarySMTP -UpnForLog $upn
                        if ($ok3) {
                            WriteLog -Message "Calendar folder found for $primarySMTP (via common localized name : $usedIdentity3)" "INFO"
                            if ($permissions3 -and $permissions3.Count -gt 0) {
                                $results += $permissions3
                            } elseif ($EmitNoPermRow) {
                                # extract folder name from identity "mbId:\Name"
                                $folderName = ($usedIdentity3 -split ':\s*',2)[1]
                                $results += (Add-NoPermissionRow -Mailbox $primarySMTP -UPN $upn -CalendarFolder $folderName)
                            }
                        } else {
                            WriteLog -Message "No calendar folder found for $primarySMTP" "WARN"
                        }
                    }
                } else {
                    # Last try with localized names if stats empty
                    $ok4, $permissions4, $usedIdentity4 = Try-GetFolderPermission -MailboxIds $mailboxIds -FolderNames @('Calendrier','Kalender','Calendario') -PrimarySmtpForLog $primarySMTP -UpnForLog $upn
                    if ($ok4) {
                        WriteLog -Message "Calendar folder found for $primarySMTP (via common localized name : $usedIdentity4)" "INFO"
                        if ($permissions4 -and $permissions4.Count -gt 0) {
                            $results += $permissions4
                        }
                        elseif ($EmitNoPermRow) {
                            $folderName = ($usedIdentity4 -split ':\s*',2)[1]
                            $results += (Add-NoPermissionRow -Mailbox $primarySMTP -UPN $upn -CalendarFolder $folderName)
                        }
                    } else {
                        WriteLog -Message "No calendar folder found for $primarySMTP" "WARN"
                    }
                }
            }
        } else {
            # Full mode
            $calendarFolders = Get-CalendarFoldersSafe -Mbx $mbx -IsOnline:$Online -PrimaryOnly:$false
            $fTotal = $calendarFolders.Count
            $fIndex = 0

            foreach ($folder in $calendarFolders) {
                $fIndex++
                $folderPath = $folder.FolderPath.TrimStart("/")
                WriteLog -Message "Calendar folder found for $primarySMTP ($folderPath)" "INFO"
                try {
                    $permMailboxId = if ($Online) { $upn } else { $mbx.Identity }
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
                        $results += $permissions
                    } elseif ($EmitNoPermRow) {
                        $results += (Add-NoPermissionRow -Mailbox $primarySMTP -UPN $upn -CalendarFolder $folderPath)
                    }
                } catch {
                    $errMsg = "Permission error for $primarySMTP ($folderPath) : $($_.Exception.Message)"
                    Write-Warning $errMsg
                    $errors += $errMsg
                }
            }
        }
    } catch {
        $errMsg = "Error for $primarySMTP : $($_.Exception.Message)"
        Write-Warning $errMsg
        $errors += $errMsg
    }
}

Write-Progress -Id 0 -Activity $overallActivity -Completed

# ------------------------- Export & Cleanup -------------------------
$BaseFileName = if ($Online) { "Exchange_EXO_MailboxCalendarPermissions_AllDomains" } else { "Exchange_OnPrem_MailboxCalendarPermissions_AllDomains" }

Write-Host "`n--- Export CSV ---"
if ($results.Count -gt 0) {
    ExportAndCopyCsv -BaseFileName $BaseFileName `
        -OutputPath $OutputPath `
        -GlobalPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') `
        -Data $results `
        -Encoding "UTF8" `
        -NoTypeInformation
} else {
    WriteLog -Message "No data to export (no calendars found or all skipped). Export step skipped." "INFO"
    Write-Host "No data to export. Skipping."
}

if ($errors.Count -gt 0) {
    Add-Content -Path $logTextFile -Value "`n=== ERRORS ==="
    $errors | ForEach-Object { Add-Content -Path $logTextFile -Value $_ }
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
Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
Complete-SmartM365ExecutionContext -Status Auto
