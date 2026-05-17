<#
.SYNOPSIS
    Azure AD (Entra ID) verified domains inventory.

.DESCRIPTION
    Retrieves verified Azure AD domains from Microsoft Graph (Get-MgDomain),
    exports a timestamped CSV to OutputPath, and writes a non-timestamped "LAST"
    CSV to LatestCsvFolderPath.

.PARAMETER Connect
    Kept for compatibility. The script always disconnects any existing Microsoft Graph session before connecting.

.PARAMETER OutputPath
    Optional override output folder. If omitted, local configuration OutputPath is used.

.PARAMETER OutputFileName
    Base CSV file name (default: M365_Entra_VerifiedDomains.csv)
#>

param(
    [string]$Tenant = 'test',
[switch]$Connect,
    [string]$OutputPath,
    [string]$OutputFileName = "M365_Entra_VerifiedDomains.csv"
)
$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $p = Join-Path -Path $d -ChildPath 'SmartM365-TenantContext.ps1'
        if (Test-Path -LiteralPath $p) { return $p }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $tenantContextPath
Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot | Out-Null

function Get-ScriptLocalConfig {
    [CmdletBinding()]
    param()

    $configPath = Join-Path -Path $PSScriptRoot -ChildPath ("{0}.local.json" -f [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))
    if (-not (Test-Path -LiteralPath $configPath)) {
        return [pscustomobject]@{}
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
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'SmartM365.global.local.json'
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
            $globalConfigPath = Join-Path -Path $searchRoot -ChildPath 'SmartM365.global.local.json'
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

# ==========================================================
# App-only authentication parameters (same app as other scripts)
# ==========================================================
$AppId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AppId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$TenantId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$Thumb = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumb' -DefaultValue '0000000000000000000000000000000000000000'
$OrgDomain = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue 'contoso.onmicrosoft.com' # Not used by Graph, kept for consistency

# ==========================================================
# "LAST" share path (non-timestamped)
# ==========================================================
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'VerifiedDomainsCsvLogFolderPath' -DefaultValue $OutputPath
$LatestCsvFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ""

# ==========================================================
# PowerShell 7 minimum
# ==========================================================
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    exit 1
}

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$MaximumFunctionCount = 32768

# ==========================================================
# Import SmartM365.Core module (psd1)
# ==========================================================
$modulePath = & { $d = $PSScriptRoot; while ($d) { $p = Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'; if (Test-Path -LiteralPath $p) { return $p }; $parent = Split-Path -Path $d -Parent; if ($parent -eq $d) { break }; $d = $parent }; throw 'SmartM365.Core module not found.' }
try {
    Import-Module $modulePath -ErrorAction Stop
} catch {
    Write-Host "Failed to import SmartM365.Core module from '$modulePath' : $_" -ForegroundColor Red
    exit 1
}

function Write-CsvAtomically {
    param(
        [Parameter(Mandatory)][object[]]$InputObject,
        [Parameter(Mandatory)][string]$Path
    )

    $dir = Split-Path -Path $Path -Parent
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }

    $tmp = Join-Path $dir ("{0}.{1}.tmp" -f ([IO.Path]::GetFileNameWithoutExtension($Path)), ([guid]::NewGuid().ToString("N")))
    try {
        $InputObject | Export-Csv -Path $tmp -NoTypeInformation -Encoding UTF8
        Move-Item -Path $tmp -Destination $Path -Force
    } finally {
        if (Test-Path $tmp) { Remove-Item -Path $tmp -Force -ErrorAction SilentlyContinue }
    }
}

function Ensure-GraphModules {
    # Avoid importing Microsoft.Graph meta-module (slow).
    $required = @(
        "Microsoft.Graph.Authentication",
        "Microsoft.Graph.Identity.DirectoryManagement"
    )

    foreach ($name in $required) {
        $m = Get-Module -ListAvailable -Name $name | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $m) {
            throw "Required module '$name' is not installed. Install it with: Install-Module $name"
        }
        Import-Module $name -ErrorAction Stop | Out-Null
    }
}

function Disconnect-GraphSafe {
    try {
        $context = Get-MgContext -ErrorAction SilentlyContinue
        if ($null -eq $context) {
            WriteLog -Message "No active Microsoft Graph session to disconnect." "INFO"
            return
        }

        Disconnect-MgGraph -ErrorAction Stop | Out-Null
        WriteLog -Message "Disconnected from Microsoft Graph." "SUCCESS"
    } catch {
        WriteLog -Message ("Disconnect-MgGraph failed (non-fatal): {0}" -f $_.Exception.Message) "WARNING"
    }
}

function New-SmartM365AiHelpUrl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ScriptName,
        [Parameter(Mandatory)][string]$Operation,
        [Parameter(Mandatory)][string]$ErrorMessage,
        [string]$InnerException = ''
    )

    $prompt = @"
Help troubleshoot this SmartM365 script failure.

Script: $ScriptName
Operation: $Operation
Error: $ErrorMessage
Inner exception: $InnerException
Computer: $env:COMPUTERNAME
Timestamp: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
"@

    return 'https://chatgpt.com/?q=' + [uri]::EscapeDataString($prompt)
}

function Get-SmartM365GlobalValue {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Name)

    $variable = Get-Variable -Name $Name -Scope Global -ErrorAction SilentlyContinue
    if ($null -eq $variable -or $null -eq $variable.Value) {
        return ''
    }

    return [string]$variable.Value
}

function Send-VerifiedDomainsTeamsInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ScriptName,
        [Parameter(Mandatory)][string]$Organization,
        [Parameter(Mandatory)][string]$OutputPath,
        [Parameter(Mandatory)][string]$TimestampedCsvPath,
        [Parameter(Mandatory)][string]$LatestCsvPath,
        [int]$TotalDomains,
        [int]$VerifiedDomains
    )

    $facts = @{
        Script               = $ScriptName
        TenantOrOrganization = $Organization
        OutputPath           = $OutputPath
        TimestampedCsvPath   = $TimestampedCsvPath
        LatestCsvPath        = $LatestCsvPath
        LogFile              = Get-SmartM365GlobalValue -Name 'LogTextFile'
        TranscriptFile       = Get-SmartM365GlobalValue -Name 'logTranscriptFile'
        TotalDomains         = $TotalDomains
        VerifiedDomains      = $VerifiedDomains
    }

    Send-SmartM365TeamsNotification `
        -Title 'Azure AD verified domains inventory - SUCCESS' `
        -Message ("Verified domains inventory completed successfully. Total domains: {0}; verified domains: {1}." -f $TotalDomains, $VerifiedDomains) `
        -Level SUCCESS `
        -Channel Infos `
        -Facts $facts | Out-Null
}

function Send-VerifiedDomainsTeamsAlert {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ScriptName,
        [Parameter(Mandatory)][string]$Organization,
        [Parameter(Mandatory)][string]$Operation,
        [Parameter(Mandatory)][System.Management.Automation.ErrorRecord]$ErrorRecord,
        [string]$OutputPath = '',
        [string]$TimestampedCsvPath = '',
        [string]$LatestCsvPath = ''
    )

    $innerException = if ($ErrorRecord.Exception.InnerException) { $ErrorRecord.Exception.InnerException.Message } else { '' }
    $helpUrl = New-SmartM365AiHelpUrl `
        -ScriptName $ScriptName `
        -Operation $Operation `
        -ErrorMessage $ErrorRecord.Exception.Message `
        -InnerException $innerException

    $facts = @{
        Script               = $ScriptName
        TenantOrOrganization = $Organization
        Operation            = $Operation
        ExceptionMessage     = $ErrorRecord.Exception.Message
        InnerException       = $innerException
        OutputPath           = $OutputPath
        TimestampedCsvPath   = $TimestampedCsvPath
        LatestCsvPath        = $LatestCsvPath
        LogFile              = Get-SmartM365GlobalValue -Name 'LogTextFile'
        TranscriptFile       = Get-SmartM365GlobalValue -Name 'logTranscriptFile'
    }

    Send-SmartM365TeamsNotification `
        -Title 'Azure AD verified domains inventory - ERROR' `
        -Message ("Terminal error during {0}: {1}" -f $Operation, $ErrorRecord.Exception.Message) `
        -Level ERROR `
        -Channel Alerts `
        -Facts $facts `
        -HelpUrl $helpUrl | Out-Null
}

#region Init
$ScriptVersion = "1.0"
$TaskNameCore  = "Azure AD verified domains inventory"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
$currentOperation = 'InitializeScriptEnvironment'
$csvPathTimestamped = ''
$csvPathLast = ''
$countAll = 0
$countVerified = 0

try {
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPath $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
    Start-Transcript -Path $global:logTranscriptFile -Append
    WriteLog -Message "Script Environment initialized at $InitializeOutputPath"
    $OutputPath = $InitializeOutputPath
    WriteLog -Message "Starting $TaskName..."
    WriteLog -Message "PowerShell Version: $($PSVersionTable.PSVersion)"
} catch {
    Write-Host "Initialization failed: $_" -ForegroundColor Red
    try {
        Send-VerifiedDomainsTeamsAlert `
            -ScriptName $TaskName `
            -Organization $OrgDomain `
            -Operation $currentOperation `
            -ErrorRecord $_ `
            -OutputPath $OutputPath
    } catch {
        WriteLog -Message ("Failed to send Teams alert notification: {0}" -f $_.Exception.Message) "ERROR"
    }
    exit 1
}
#endregion

try {
    $logPath = Split-Path -Path $global:LogTextFile -Parent

    # ==========================================================
    # Connection management (Graph only - app-only certificate)
    # ==========================================================
    $currentOperation = 'EnsureGraphModules'
    Ensure-GraphModules

    $currentOperation = 'DisconnectExistingGraphSession'
    Write-Host "Existing Microsoft Graph session (if any) will be disconnected before app-only connection..." -ForegroundColor Cyan
    Disconnect-GraphSafe

    $currentOperation = 'ConnectGraph'
    WriteLog -Message "Connecting to Microsoft Graph with app-only certificate authentication." "INFO"

    Connect-MgGraph `
        -ClientId $AppId `
        -TenantId $TenantId `
        -CertificateThumbprint $Thumb `
        -NoWelcome `
        -ErrorAction Stop | Out-Null

    WriteLog -Message "Connected to Microsoft Graph."

    $currentOperation = 'Preflight'
    Invoke-SmartM365Preflight `
        -ScriptName $TaskName `
        -RequiredModules @('Microsoft.Graph.Authentication','Microsoft.Graph.Identity.DirectoryManagement') `
        -RequiredCommands @('Get-MgDomain') `
        -OutputPaths @($OutputPath) `
        -GraphProbeUris @('https://graph.microsoft.com/v1.0/domains?$top=1') | Out-Null

    # ==========================================================
    # Retrieve verified domains
    # ==========================================================
    $currentOperation = 'RetrieveDomains'
    WriteLog -Message "Retrieving Azure AD domains via Get-MgDomain..."
    $allDomains = Get-MgDomain -All

    $verifiedDomains = $allDomains | Where-Object { $_.IsVerified -eq $true } | ForEach-Object {
        [pscustomobject]@{
            Id                 = $_.Id
            IsVerified         = $_.IsVerified
            IsDefault          = $_.IsDefault
            IsInitial          = $_.IsInitial
            AuthenticationType = $_.AuthenticationType
            SupportedServices  = ($_.SupportedServices -join ";")
            AvailabilityStatus = $_.AvailabilityStatus
        }
    }

    $countAll      = if ($allDomains) { ($allDomains | Measure-Object).Count } else { 0 }
    $countVerified = if ($verifiedDomains) { ($verifiedDomains | Measure-Object).Count } else { 0 }

    WriteLog -Message ("Domains retrieved (total): {0}" -f $countAll)
    WriteLog -Message ("Domains retrieved (verified): {0}" -f $countVerified)

    # ==========================================================
    # Export CSV (timestamped) + write "LAST" to share (no timestamp)
    # ==========================================================
    $currentOperation = 'ExportCsv'
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $baseName  = [IO.Path]::GetFileNameWithoutExtension($OutputFileName)
    $ext       = [IO.Path]::GetExtension($OutputFileName)
    if ([string]::IsNullOrWhiteSpace($ext)) { $ext = ".csv" }

    $timestampedFileName = "{0}_{1}{2}" -f $baseName, $timestamp, $ext
    $csvPathTimestamped  = Join-Path $OutputPath $timestampedFileName

    WriteLog -Message ("Exporting timestamped CSV (atomic) to: {0}" -f $csvPathTimestamped)
    Write-CsvAtomically -InputObject @($verifiedDomains) -Path $csvPathTimestamped

    $csvPathLast = Join-Path $LatestCsvFolderPath $OutputFileName
    WriteLog -Message ("Writing LAST CSV to share (atomic): {0}" -f $csvPathLast)
    Write-CsvAtomically -InputObject @($verifiedDomains) -Path $csvPathLast
    Invoke-SmartM365SharePointCsvUpload -LocalFilePath $csvPathLast

    # ==========================================================
    # Summary
    # ==========================================================
    Write-Host "`n--- Execution Summary ---"
    Write-Host "Domains (total)          : $countAll"
    Write-Host "Domains (verified)       : $countVerified"
    Write-Host "CSV (timestamped)        : $csvPathTimestamped"
    Write-Host "CSV (LAST on share)      : $csvPathLast"
    Write-Host "Log file                 : $global:LogTextFile"
    Write-Host "-------------------------`n"

    $currentOperation = 'SendTeamsInfo'
    try {
        Send-VerifiedDomainsTeamsInfo `
            -ScriptName $MyInvocation.MyCommand.Name `
            -Organization $OrgDomain `
            -OutputPath $OutputPath `
            -TimestampedCsvPath $csvPathTimestamped `
            -LatestCsvPath $csvPathLast `
            -TotalDomains $countAll `
            -VerifiedDomains $countVerified
    } catch {
        WriteLog -Message ("Failed to send Teams info notification: {0}" -f $_.Exception.Message) "WARNING"
    }

    # ==========================================================
    # Disconnect + Cleanup
    # ==========================================================
    $currentOperation = 'DisconnectGraph'
    Write-Host "`n--- Disconnect Microsoft Graph ---"
    Disconnect-GraphSafe
    WriteLog -Message "$TaskName completed."

    try { Stop-Transcript | Out-Null } catch {}

    $currentOperation = 'Cleanup'
    RemoveOldFiles -Path $OutputPath -Filter "*.csv" -KeepCount $global:RetentionMaxCSV -LogFile $global:LogTextFile
    RemoveOldFiles -Path $logPath    -Filter "*.log" -KeepCount $global:RetentionMaxLogs -LogFile $global:LogTextFile
}
catch {
    $globalError = $_
    WriteLog -Message ("Global error in Azure AD verified domains inventory: {0}" -f $globalError) "ERROR"
    Write-Host "A global error occurred. Check the log file for details." -ForegroundColor Red

    try {
        Send-VerifiedDomainsTeamsAlert `
            -ScriptName $MyInvocation.MyCommand.Name `
            -Organization $OrgDomain `
            -Operation $currentOperation `
            -ErrorRecord $globalError `
            -OutputPath $OutputPath `
            -TimestampedCsvPath $csvPathTimestamped `
            -LatestCsvPath $csvPathLast
    } catch {
        WriteLog -Message ("Failed to send Teams alert notification: {0}" -f $_.Exception.Message) "ERROR"
    }

    # -------- Global error email via SmartM365.Core (SendEmailHtmlReport) --------
    try {
        $title = "Azure AD verified domains inventory - ERROR"
        $msg   = @"
An error occurred in script $($MyInvocation.MyCommand.Name) on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss").

Error message:
$($globalError.Exception.Message)

Log file:
$($global:LogTextFile)
"@

        $bodyHtml = NewSimpleEmailBody -Title $title -Message $msg

        $attachments = @()
        if ($global:LogTextFile -and (Test-Path $global:LogTextFile)) {
            $attachments = @($global:LogTextFile)
        }

        SendEmailHtmlReport -Subject $title -BodyHtml $bodyHtml -Attachments $attachments
    } catch {
        WriteLog -Message ("Failed to send global error email: {0}" -f $_.Exception.Message) "ERROR"
    }

    try { Disconnect-GraphSafe } catch {}
    try { Stop-Transcript | Out-Null } catch {}
    exit 1
}



