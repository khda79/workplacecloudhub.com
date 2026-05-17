<#
.SYNOPSIS
    Exchange Online accepted domains inventory.

.DESCRIPTION
    Retrieves accepted domains from Exchange Online (Get-AcceptedDomain)
    and exports them to a UTF-8 CSV using an atomic write pattern.

    Outputs:
      - Timestamped CSV in OutputPath
      - Non-timestamped "LAST" CSV copied to <local-share-path>

.PARAMETER Connect
    Forces disconnect/reconnect to Exchange Online.

.PARAMETER OutputPath
    Optional override output folder. If omitted, local configuration OutputPath is used.

.PARAMETER OutputFileName
    Base CSV file name (default: Exchange_EXO_AcceptedDomains.csv)

.PARAMETER DefaultOnly
    Exports only the default accepted domain.
.NOTES
    Author: https://github.com/khda79/M365
#>

param(
    [string]$Tenant = 'test',
[switch]$Connect,
    [string]$OutputPath,
    [string]$OutputFileName = "Exchange_EXO_AcceptedDomains.csv",
    [switch]$DefaultOnly
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
$AppId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AppId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$TenantId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$Thumb = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumb' -DefaultValue '0000000000000000000000000000000000000000'
$OrgDomain = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue 'contoso.onmicrosoft.com'

# ==========================================================
# "LAST" share path (non-timestamped)
# ==========================================================
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
    } catch {
        WriteLog -Message ("Disconnect-ExchangeOnline failed (non-fatal): {0}" -f $_.Exception.Message) "WARN"
    }
}

function Send-AcceptedDomainsErrorNotification {
    param(
        [Parameter(Mandatory)]$ErrorRecord,
        [string]$Operation,
        [string]$TimestampedCsvPath,
        [string]$LatestCsvPath
    )

    try {
        $exception = $ErrorRecord.Exception
        $innerMessages = New-Object System.Collections.Generic.List[string]
        $inner = $exception.InnerException
        while ($null -ne $inner) {
            if (-not [string]::IsNullOrWhiteSpace($inner.Message)) {
                $innerMessages.Add($inner.Message) | Out-Null
            }
            $inner = $inner.InnerException
        }

        $scriptName = [System.IO.Path]::GetFileName($PSCommandPath)
        $errorContext = @(
            "Script: $scriptName"
            "Tenant/Organization: $OrgDomain"
            "Operation: $Operation"
            "Error: $($exception.Message)"
        ) -join "`n"

        $helpUrl = "https://chat.openai.com/?q={0}" -f [System.Uri]::EscapeDataString("Help troubleshoot this SmartM365 Exchange Online inventory error:`n$errorContext")

        $facts = @{
            "Script name"       = $scriptName
            "Tenant/Organization" = $OrgDomain
            "Failed operation"  = $Operation
            "Exception message" = $exception.Message
            "Inner exception"   = ($innerMessages -join " | ")
            "Log path"          = $global:LogTextFile
            "Transcript path"   = $global:logTranscriptFile
            "Output path"       = $OutputPath
            "Timestamped CSV"   = $TimestampedCsvPath
            "Latest CSV"        = $LatestCsvPath
        }

        Send-SmartM365TeamsNotification `
            -Title "SmartM365 Accepted Domains inventory failed" `
            -Message "A terminal error occurred in Exchange Online accepted domains inventory." `
            -Level "ERROR" `
            -Channel "Alerts" `
            -Facts $facts `
            -HelpUrl $helpUrl | Out-Null
    }
    catch {
        WriteLog -Message ("Failed to send Teams error notification: {0}" -f $_.Exception.Message) "ERROR"
    }
}

function Send-AcceptedDomainsSuccessNotification {
    param(
        [int]$AcceptedDomainCount,
        [string]$TimestampedCsvPath,
        [string]$LatestCsvPath
    )

    try {
        $scriptName = [System.IO.Path]::GetFileName($PSCommandPath)
        $facts = @{
            "Script name"       = $scriptName
            "Tenant/Organization" = $OrgDomain
            "Accepted domains"  = $AcceptedDomainCount
            "Output path"       = $OutputPath
            "Timestamped CSV"   = $TimestampedCsvPath
            "Latest CSV"        = $LatestCsvPath
            "Log path"          = $global:LogTextFile
            "Transcript path"   = $global:logTranscriptFile
        }

        Send-SmartM365TeamsNotification `
            -Title "SmartM365 Accepted Domains inventory completed" `
            -Message ("Exchange Online accepted domains inventory completed without error. Accepted domains exported: {0}." -f $AcceptedDomainCount) `
            -Level "SUCCESS" `
            -Channel "Infos" `
            -Facts $facts | Out-Null
    }
    catch {
        WriteLog -Message ("Failed to send Teams completion notification: {0}" -f $_.Exception.Message) "WARN"
    }
}

#region Init
$ScriptVersion = "1.0"
$TaskNameCore  = "Exchange Online accepted domains inventory"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
$currentOperation = "Initialize script environment"
$csvPathTimestamped = ""
$csvPathLast = ""

try {
    $OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AcceptedDomainsCsvLogFolderPath' -DefaultValue $OutputPath
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPath $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
    Start-Transcript -Path $global:logTranscriptFile -Append
    WriteLog -Message "Script Environment initialized at $InitializeOutputPath"
    $OutputPath = $InitializeOutputPath
    WriteLog -Message "Starting $TaskName..."
    WriteLog -Message "PowerShell Version: $($PSVersionTable.PSVersion)"
} catch {
    Send-AcceptedDomainsErrorNotification -ErrorRecord $_ -Operation $currentOperation -TimestampedCsvPath $csvPathTimestamped -LatestCsvPath $csvPathLast
    Write-Host "Initialization failed: $_" -ForegroundColor Red
    exit 1
}
#endregion

try {
    $logPath = Split-Path -Path $global:LogTextFile -Parent

    # ==========================================================
    # Connection management (EXO only - standalone app-only)
    # ==========================================================
    $currentOperation = "Load ExchangeOnlineManagement module"
    Ensure-ExchangeOnlineModule

    if ($Connect) {
        Write-Host "Connect switch specified: existing Exchange Online session (if any) will be disconnected and reconnected..." -ForegroundColor Cyan
    }

    $currentOperation = "Disconnect existing Exchange Online session"
    Disconnect-ExchangeOnlineSafe

    $currentOperation = "Connect to Exchange Online"
    WriteLog -Message "Connecting to Exchange Online with app-only certificate authentication." "INFO"

    Connect-ExchangeOnline `
        -AppId $AppId `
        -CertificateThumbprint $Thumb `
        -Organization $OrgDomain `
        -ShowBanner:$false `
        -ErrorAction Stop | Out-Null

    WriteLog -Message "Connected to Exchange Online."
    $currentOperation = "Run preflight checks"
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -ExchangeOnlineProbeCommands @('Get-AcceptedDomain') | Out-Null

    # ==========================================================
    # Retrieve accepted domains
    # ==========================================================
    $currentOperation = "Retrieve accepted domains"
    WriteLog -Message "Retrieving accepted domains via Get-AcceptedDomain..."
    $domains = Get-AcceptedDomain | Select-Object `
        Name,
        DomainName,
        DomainType,
        Default,
        AddressBookEnabled,
        WhenCreatedUTC,
        WhenChangedUTC

    if ($DefaultOnly) {
        $domains = $domains | Where-Object { $_.Default -eq $true }
        WriteLog -Message "DefaultOnly enabled: exporting only the default accepted domain."
    }

    $count = 0
    if ($domains) { $count = ($domains | Measure-Object).Count }
    WriteLog -Message ("Accepted domains retrieved: {0}" -f $count)

    # ==========================================================
    # Export CSV (timestamped) + copy "LAST" to share (no timestamp)
    # ==========================================================
    $currentOperation = "Export accepted domains CSV"
    $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $baseName  = [IO.Path]::GetFileNameWithoutExtension($OutputFileName)
    $ext       = [IO.Path]::GetExtension($OutputFileName)
    if ([string]::IsNullOrWhiteSpace($ext)) { $ext = ".csv" }

    $timestampedFileName = "{0}_{1}{2}" -f $baseName, $timestamp, $ext
    $csvPathTimestamped  = Join-Path $OutputPath $timestampedFileName

    WriteLog -Message ("Exporting timestamped CSV (atomic) to: {0}" -f $csvPathTimestamped)
    Write-CsvAtomically -InputObject @($domains) -Path $csvPathTimestamped

    $csvPathLast = Join-Path $LatestCsvFolderPath $OutputFileName
    WriteLog -Message ("Writing LAST CSV to share (atomic): {0}" -f $csvPathLast)
    Write-CsvAtomically -InputObject @($domains) -Path $csvPathLast
    $currentOperation = "Upload accepted domains CSV to SharePoint"
    Invoke-SmartM365SharePointCsvUpload -LocalFilePath $csvPathLast

    # ==========================================================
    # Summary
    # ==========================================================
    Write-Host "`n--- Execution Summary ---"
    Write-Host "Accepted domains exported : $count"
    Write-Host "CSV (timestamped)        : $csvPathTimestamped"
    Write-Host "CSV (LAST on share)      : $csvPathLast"
    Write-Host "Log file                 : $global:LogTextFile"
    Write-Host "-------------------------`n"

    # ==========================================================
    # Disconnect + Cleanup
    # ==========================================================
    $currentOperation = "Disconnect Exchange Online"
    Write-Host "`n--- Disconnect Exchange Online ---"
    Disconnect-ExchangeOnlineSafe

    $currentOperation = "Apply retention cleanup"
    RemoveOldFiles -Path $OutputPath -Filter "*.csv" -KeepCount $global:RetentionMaxCSV -LogFile $global:LogTextFile
    RemoveOldFiles -Path $logPath    -Filter "*.log" -KeepCount $global:RetentionMaxLogs -LogFile $global:LogTextFile

    $currentOperation = "Send Teams completion notification"
    Send-AcceptedDomainsSuccessNotification -AcceptedDomainCount $count -TimestampedCsvPath $csvPathTimestamped -LatestCsvPath $csvPathLast

    WriteLog -Message "$TaskName completed."
    try { Stop-Transcript | Out-Null } catch {}
}
catch {
    $globalError = $_
    WriteLog -Message ("Global error in Exchange Online accepted domains inventory: {0}" -f $globalError) "ERROR"
    Write-Host "A global error occurred. Check the log file for details." -ForegroundColor Red

    Send-AcceptedDomainsErrorNotification -ErrorRecord $globalError -Operation $currentOperation -TimestampedCsvPath $csvPathTimestamped -LatestCsvPath $csvPathLast

    try {
        $title = "Exchange Online accepted domains inventory - ERROR"
        $msg   = @"
An error occurred in script $($MyInvocation.MyCommand.Name) on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss").

Error message:
$($globalError.Exception.Message)

See attached log file for details:
$($global:LogTextFile)
"@

        $bodyHtml = NewSimpleEmailBody -Title $title -Message $msg

        $attachments = @()
        if ($global:LogTextFile -and (Test-Path $global:LogTextFile)) {
            $attachments = @($global:LogTextFile)
        }

        Send-SmartM365Mail -Subject $title -BodyHtml $bodyHtml -Attachments $attachments
    } catch {
        WriteLog -Message ("Failed to send global error email: {0}" -f $_) "ERROR"
    }

    try { Disconnect-ExchangeOnlineSafe } catch {}
    try { Stop-Transcript | Out-Null } catch {}
    exit 1
}



