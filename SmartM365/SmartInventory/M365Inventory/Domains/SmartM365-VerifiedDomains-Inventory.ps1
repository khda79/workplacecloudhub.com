<#
.SYNOPSIS
    Azure AD (Entra ID) verified domains inventory.

.DESCRIPTION
    Retrieves verified Azure AD domains from Microsoft Graph (Get-MgDomain),
    exports a timestamped CSV to OutputPath, and writes a non-timestamped "LAST"
    CSV to LatestCsvFolderPath.

.PARAMETER Tenant
    Tenant profile key to load from Config/Tenants. Defaults to test.

.PARAMETER Connect
    Kept for compatibility. The script always disconnects any existing Microsoft Graph session before connecting.

.PARAMETER OutputPath
    Optional override output folder. If omitted, local configuration OutputPath is used.

.PARAMETER OutputFileName
    Base CSV file name (default: M365_Entra_VerifiedDomains.csv)
.VERSION
1.5


.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication; Microsoft.Graph.Identity.DirectoryManagement.
    Minimum Graph application permissions: Directory.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Requires: PowerShell 7+, Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement, SmartM365.Core.psd1
    Minimum application permissions: Directory.Read.All
#>

param(
    [string]$Tenant = 'test',
    [switch]$Connect,
    [string]$OutputPath,
    [string]$OutputFileName = "M365_Entra_VerifiedDomains.csv",
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

    $resultSummary = "Verified domains inventory completed successfully. Total domains: {0}; verified domains: {1}." -f $TotalDomains, $VerifiedDomains

    Send-SmartM365TeamsNotification `
        -Title 'Azure AD verified domains inventory - SUCCESS' `
        -Message $resultSummary `
        -Level SUCCESS `
        -Channel Infos `
        -ResultSummary $resultSummary `
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
$ScriptVersion = "1.5"
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
        -RequiredGraphApplicationPermissions @('Directory.Read.All') -GraphProbeUris @('https://graph.microsoft.com/v1.0/domains?$top=1') | Out-Null

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
    $verifiedDomains = @($verifiedDomains | Sort-Object Id)
    if ($MaxItems -gt 0 -and $verifiedDomains.Count -gt $MaxItems) {
        WriteLog -Message ("MaxItems enabled: restricted verified domains from {0} to {1}." -f $verifiedDomains.Count, $MaxItems) "WARN"
        $verifiedDomains = @($verifiedDomains | Select-Object -First $MaxItems)
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

    $csvPathLast = Join-Path $LatestCsvFolderPath $OutputFileName
    WriteLog -Message ("Publishing CSV. Timestamped: {0}; latest: {1}" -f $csvPathTimestamped, $csvPathLast)
    Export-SmartM365Csv -Data @($verifiedDomains) -TimestampedPath $csvPathTimestamped -LatestPath $csvPathLast | Out-Null

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

    try { Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {} } catch {}
    Complete-SmartM365ExecutionContext -Status Auto

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
    try { Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {} } catch {}
    Complete-SmartM365ExecutionContext -Status Auto
    exit 1
}
# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBpeBldyvDQ3enG
# ej64/w8CT7kedHhdSUDm6aGgIWnv9qCCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAKpaOv3oCiUYhXBdA9
# tAYGSb/61i6ZWunHsnLP+IEs8zANBgkqhkiG9w0BAQEFAASCAYBbzg6C+NobdTaC
# nG+e9FmLx11xzilioR05MGGaZRd+8oI/izMXDGvGZ4MVFDRF4rQ5IY6RIjTWSjCH
# PAjLtLFGcw5jNw7Z5YTH3AfNygq+Ta+dI9XZfE/2YfOQyTITfDj9BoJucpGHmKHT
# 28f3lG3X8t2DOGzFYAqTmd5vf2iM5JQkKit6c8LFnT68s/jAxk1pNc6xYA57TWMI
# lBvIVYHkLRUbMaRh+fQqGD/9VSbCmbmSAvVi6k3B69yXjjCvcVhQuGVi5HJHQVSJ
# eJICITGz7+DRQ1HP54aoKOSJLst9r9UwXxZcSvgPyvNpD3yjq8n1SAJYDJAvJpg1
# NfL28yGR7wDRuT8Q2sOcYA6QRk6rFOzBb2BCiXGtqaWV0HHf1QY03C1Hf0i7+sm3
# OFuLyOh3THjuqeH79f59a4ruNAyJXlusmfNQt7uHlT6h7QzNUL5hqOY5jqhojW58
# ce36M8ug9lf3+15Y7bJTZQ8wEV/jWeH53Ef+Dz9gUWspWPuG/7Q=
# SIG # End signature block
