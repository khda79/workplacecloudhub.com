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
.VERSION
1.5


.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; ExchangeOnlineManagement.
    Minimum permissions: Exchange.ManageAsApp plus an Exchange Online app-only RBAC role allowing Get-AcceptedDomain; Global Reader is the default read-only service-principal role.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Minimum permissions: Exchange Online app-only RBAC must allow Get-AcceptedDomain.
#>

param(
    [string]$Tenant = 'test',
[switch]$Connect,
    [string]$OutputPath,
    [string]$OutputFileName = "Exchange_EXO_AcceptedDomains.csv",
    [switch]$DefaultOnly,
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

        $resultSummary = "Exchange Online accepted domains inventory completed without error. Accepted domains exported: {0}." -f $AcceptedDomainCount

        Send-SmartM365TeamsNotification `
            -Title "SmartM365 Accepted Domains inventory success" `
            -Message $resultSummary `
            -Level "SUCCESS" `
            -Channel "Infos" `
            -ResultSummary $resultSummary `
            -Facts $facts | Out-Null
    }
    catch {
        WriteLog -Message ("Failed to send Teams completion notification: {0}" -f $_.Exception.Message) "WARN"
    }
}

#region Init
$ScriptVersion = "1.5"
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
    $domains = @($domains | Sort-Object DomainName, Name)
    if ($MaxItems -gt 0 -and $domains.Count -gt $MaxItems) {
        WriteLog -Message ("MaxItems enabled: restricted accepted domains from {0} to {1}." -f $domains.Count, $MaxItems) "WARN"
        $domains = @($domains | Select-Object -First $MaxItems)
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

    $csvPathLast = Join-Path $LatestCsvFolderPath $OutputFileName
    WriteLog -Message ("Publishing CSV. Timestamped: {0}; latest: {1}" -f $csvPathTimestamped, $csvPathLast)
    $currentOperation = "Publish accepted domains CSV"
    Export-SmartM365Csv -Data @($domains) -TimestampedPath $csvPathTimestamped -LatestPath $csvPathLast | Out-Null

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
    try { Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {} } catch {}
    Complete-SmartM365ExecutionContext -Status Auto
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
    try { Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {} } catch {}
    Complete-SmartM365ExecutionContext -Status Auto
    exit 1
}
# SIG # Begin signature block
# MIIHcgYJKoZIhvcNAQcCoIIHYzCCB18CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDKPdOMOdvNWrAG
# J+9gQVSanxeXjxEe5tBRXosHr19GL6CCBEgwggREMIICrKADAgECAhBxu0EivlCF
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
# AQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBXg/LbI40elj39r0t1
# 3khUG/0mAdzKwLC0Z0s6yE21FTANBgkqhkiG9w0BAQEFAASCAYAN06NWzSyZ36Uk
# c5adhSuXdcdcr+gO4vq2v+GApY7qbgt7oCz5MWnvzfPeBIXY6VgpIOZk03UAaLDF
# yVdLjk4EUbfcRmALhqI71GxxmaveoAKiQaeYL+5UQlvmDXBN6KkmRlzsZ/0crwom
# u+pVoPyJrz6r5th/hgzrhtxDICdwvtP2NrZu+JmFpzFEUotCD9TzqRsgsdB3tYQ6
# ZOEndHqG8QolpvGTyWv2k0d6YCoFb18b10Rzm+xdI7+wN4zNR6bqhxgkKqrvguXP
# 7Y9eHkdQbeB+O3ZCS8GYupGRd3VmY0ZrBEqTrAo5iGXXGXNc9nHQjIbEI1XEtcCe
# mVEXq+VkcK7na9qF/NFv5h/zzqhcuIEA7pVKXcqybHDwwIZglW6CdIPS3v18q7hL
# Tfac/PYs4ztL9ZvNNSYwAY5pxzz5M/4s58Z1fBGoSPpuyPek+72SIcwKjVNzhHph
# N/6Li1QuqHz4/GVKhUbIyqelzgLDN7/Q+R0izYTX0oGQHENTdo8=
# SIG # End signature block
