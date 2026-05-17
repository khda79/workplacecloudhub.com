#Requires -Version 7.0
<#
.SYNOPSIS
    Compares members of the configured M365 Backup protected mailboxes group with all EXO mailboxes.

.DESCRIPTION
    Retrieves the configured M365 Backup protected mailboxes group members via Microsoft Graph,
    and EXO mailboxes from a CSV export (EXO-Mailboxes-Inventory) or live via
    Connect-ExchangeOnline if no CSV is provided or found.

    Produces three output CSV files:
      - UnprotectedMailboxes : EXO mailboxes absent from the backup group
      - ProtectedMailboxes   : EXO mailboxes present in the backup group
      - MembersWithoutMailbox: Group members with no matching EXO mailbox

    Join key: ExternalDirectoryObjectId (EXO) <-> Id (Microsoft Graph)

.PARAMETER Tenant
    Tenant profile key to load from Config/Tenants. Defaults to test.

.PARAMETER TenantId
    Azure AD tenant ID (GUID) used for app-only authentication with Microsoft Graph.

.PARAMETER OrgDomain
    Tenant domain (e.g. contoso.onmicrosoft.com) used for Exchange Online app-only authentication.

.PARAMETER AppId
    Application (client) ID of the Azure AD App Registration.

.PARAMETER CertificateThumbprint
    Thumbprint of the certificate installed in the local certificate store,
    associated with the App Registration for app-only authentication.

.PARAMETER GroupDisplayName
    Display name of the Azure AD group for SmartM365 Backup protected mailboxes.
    Used to resolve GroupObjectId when GroupObjectId is not explicitly configured.

.PARAMETER GroupObjectId
    Optional ObjectId override for the Azure AD group for M365 Backup protected mailboxes.

.PARAMETER EXOMailboxesCsvPath
    Optional path to a CSV file from EXO-Mailboxes-Inventory.
    If omitted or not found, the script connects live to Exchange Online.

.PARAMETER ScriptCsvLogFolderPath
    Directory where timestamped output CSV files will be written (historical archive).

.PARAMETER LatestCsvFolderPath
    Directory where the latest CSV files are copied with fixed names (overwritten each run).

.PARAMETER ErrorMailTo
    Recipient address for error notification emails.

.PARAMETER From
    Sender address for error notification emails.

.NOTES
    Name      : EXO-BackupProtection-Comparison
    Version   : 1.0
    Author    : https://github.com/khda79/M365
    Requires  : Microsoft.Graph.Groups, Microsoft.Graph.Users, ExchangeOnlineManagement
#>

[CmdletBinding()]
param (
    [string]$Tenant = 'test',
    [string]$TenantId              = "00000000-0000-0000-0000-000000000000",
    [string]$OrgDomain             = "contoso.onmicrosoft.com",
    [string]$AppId                 = "00000000-0000-0000-0000-000000000000",
    [string]$CertificateThumbprint = "0000000000000000000000000000000000000000",
    [string]$GroupDisplayName      = "SmartM365 Backup - Protected Mailboxes",
    [string]$GroupObjectId         = "00000000-0000-0000-0000-000000000000",
    [string]$EXOMailboxesCsvPath = "",
    [string]$ScriptCsvLogFolderPath           = "",
    [string]$LatestCsvFolderPath       = "",
    [string]$ErrorMailTo        = "",
    [string]$From      = ""
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
            $scriptRootValue = Get-Variable -Name ScriptRoot -ValueOnly -ErrorAction SilentlyContinue
    $searchRoot = if ($PSScriptRoot) { $PSScriptRoot } elseif ($scriptRootValue) { $scriptRootValue } elseif ($PSCommandPath) { Split-Path -Path $PSCommandPath -Parent } else { (Get-Location).Path }
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
Set-StrictMode -Version Latest

foreach ($configName in @('TenantId','OrgDomain','AppId','CertificateThumbprint','GroupDisplayName','GroupObjectId','EXOMailboxesCsvPath','ScriptCsvLogFolderPath','LatestCsvFolderPath','ErrorMailTo','From')) {
    if (-not $PSBoundParameters.ContainsKey($configName)) {
        Set-Variable -Name $configName -Value (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name $configName -DefaultValue (Get-Variable -Name $configName -ValueOnly)) -Scope Local
    }
}

if (-not $PSBoundParameters.ContainsKey('CertificateThumbprint') -and ($CertificateThumbprint -eq '0000000000000000000000000000000000000000' -or [string]::IsNullOrWhiteSpace($CertificateThumbprint))) {
    $CertificateThumbprint = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumbprint' -DefaultValue $CertificateThumbprint
    if ($CertificateThumbprint -eq '0000000000000000000000000000000000000000' -or [string]::IsNullOrWhiteSpace($CertificateThumbprint)) {
        $CertificateThumbprint = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumb' -DefaultValue $CertificateThumbprint
    }
}
$LogAllRootPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LogAllRootPath' -DefaultValue ''
$ErrorActionPreference = "Stop"
$script:CurrentOperation = "Initialize"

function Assert-RequiredConfigValue {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)][string]$Name,
        [AllowNull()][string]$Value,
        [string[]]$InvalidValues = @()
    )

    if ([string]::IsNullOrWhiteSpace($Value) -or $InvalidValues -contains $Value) {
        throw "Required configuration value '$Name' is missing or still set to a placeholder."
    }
}

Assert-RequiredConfigValue -Name 'TenantId' -Value $TenantId -InvalidValues @('00000000-0000-0000-0000-000000000000')
Assert-RequiredConfigValue -Name 'OrgDomain' -Value $OrgDomain -InvalidValues @('contoso.onmicrosoft.com')
Assert-RequiredConfigValue -Name 'AppId' -Value $AppId -InvalidValues @('00000000-0000-0000-0000-000000000000')
Assert-RequiredConfigValue -Name 'CertificateThumbprint' -Value $CertificateThumbprint -InvalidValues @('0000000000000000000000000000000000000000')
Assert-RequiredConfigValue -Name 'GroupDisplayName' -Value $GroupDisplayName
Assert-RequiredConfigValue -Name 'ScriptCsvLogFolderPath' -Value $ScriptCsvLogFolderPath
Assert-RequiredConfigValue -Name 'LatestCsvFolderPath' -Value $LatestCsvFolderPath
Assert-RequiredConfigValue -Name 'ErrorMailTo' -Value $ErrorMailTo
Assert-RequiredConfigValue -Name 'From' -Value $From

function Import-SmartM365CorePreflight {
    if (Get-Command Invoke-CoreSmartM365Preflight -ErrorAction SilentlyContinue) { return }

        $scriptRootValue = Get-Variable -Name ScriptRoot -ValueOnly -ErrorAction SilentlyContinue
    $searchRoot = if ($PSScriptRoot) { $PSScriptRoot } elseif ($scriptRootValue) { $scriptRootValue } elseif ($PSCommandPath) { Split-Path -Path $PSCommandPath -Parent } else { (Get-Location).Path }
    while ($searchRoot) {
        $modulePath = Join-Path -Path $searchRoot -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
        if (Test-Path -LiteralPath $modulePath) {
            Import-Module $modulePath -Prefix Core -ErrorAction Stop
            return
        }

        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }

    throw 'SmartM365.Core module was not found. Preflight checks cannot run.'
}


# ─────────────────────────────────────────────────────────────────────────────
# REGION: Helpers
# ─────────────────────────────────────────────────────────────────────────────

function Write-Log {
    param (
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR")][string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"
    Write-Host $line
    if ($script:LogPath) {
        $line | Out-File -FilePath $script:LogPath -Append -Encoding UTF8
    }
}

function Send-ErrorEmail {
    param (
        [string]$ErrorMessage,
        [string]$ScriptName,
        [string]$GroupDisplayName = "",
        [string]$GroupObjectId
    )
    try {
        $groupLabel = if (-not [string]::IsNullOrWhiteSpace($GroupDisplayName)) { $GroupDisplayName } else { $GroupObjectId }
        $subject = "[SmartM365] ERROR - $ScriptName - Group $groupLabel"
        $body = @"
<html><body style='font-family:Segoe UI,Arial,sans-serif;font-size:13px;'>
<h2 style='color:#c0392b;'>Script Error Notification</h2>
<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;'>
<tr><td><b>Script</b></td><td>$ScriptName</td></tr>
<tr><td><b>Group DisplayName</b></td><td>$GroupDisplayName</td></tr>
<tr><td><b>Group ObjectId</b></td><td>$GroupObjectId</td></tr>
<tr><td><b>Timestamp</b></td><td>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</td></tr>
<tr><td><b>Host</b></td><td>$($env:COMPUTERNAME)</td></tr>
<tr><td><b>User</b></td><td>$($env:USERDOMAIN)\$($env:USERNAME)</td></tr>
<tr><td><b>Error</b></td><td style='color:#c0392b;'>$([System.Web.HttpUtility]::HtmlEncode($ErrorMessage))</td></tr>
</table>
</body></html>
"@
        Import-SmartM365CorePreflight
        Send-CoreSmartM365Mail -From $From -To $ErrorMailTo -Subject $subject -BodyHtml $body
        Write-Log "Error notification sent to $ErrorMailTo via Microsoft Graph" -Level "INFO"
    } catch {
        Write-Log "Failed to send error email: $_" -Level "WARN"
    }

    Send-BackupProtectionErrorNotification -ErrorMessage $ErrorMessage -Operation $script:CurrentOperation
}

function Export-CsvAtomic {
    param (
        [System.Collections.Generic.List[PSCustomObject]]$Data,
        [string]$Path,
        [string[]]$Columns = @()
    )
    $tempPath = $Path + ".tmp"
    if ($Data.Count -eq 0 -and $Columns.Count -gt 0) {
        $header = ($Columns | ForEach-Object { '"' + ($_ -replace '"', '""') + '"' }) -join ';'
        Set-Content -Path $tempPath -Value $header -Encoding UTF8
    }
    elseif ($Columns.Count -gt 0) {
        $Data | Select-Object -Property $Columns | Export-Csv -Path $tempPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    }
    else {
        $Data | Export-Csv -Path $tempPath -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    }
    Move-Item -Path $tempPath -Destination $Path -Force
}

function Resolve-BackupProtectionGroupObjectId {
    [CmdletBinding()]
    param (
        [string]$ConfiguredGroupObjectId,
        [string]$ConfiguredGroupDisplayName
    )

    $placeholderGroupObjectId = "00000000-0000-0000-0000-000000000000"

    if ([string]::IsNullOrWhiteSpace($ConfiguredGroupDisplayName)) {
        throw "GroupDisplayName must be configured."
    }

    $groupName = $ConfiguredGroupDisplayName.Trim()
    $escapedGroupName = $groupName.Replace("'", "''")
    $filter = [System.Uri]::EscapeDataString("displayName eq '$escapedGroupName'")
    $groupsByNameUri = "https://graph.microsoft.com/v1.0/groups?`$filter=$filter&`$select=id,displayName,mailNickname"
    $configuredObjectIdIsValid = -not [string]::IsNullOrWhiteSpace($ConfiguredGroupObjectId) -and $ConfiguredGroupObjectId -ne $placeholderGroupObjectId

    if ($configuredObjectIdIsValid) {
        $configuredObjectId = $ConfiguredGroupObjectId.Trim()
        Write-Log "Validating backup protection group ObjectId $configuredObjectId against display name '$groupName'."
        try {
            $groupById = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/groups/$configuredObjectId?`$select=id,displayName,mailNickname" -OutputType PSObject -ErrorAction Stop
        }
        catch {
            throw "Configured GroupObjectId '$configuredObjectId' was not found. Clear GroupObjectId to resolve by GroupDisplayName, or update it with the expected group ObjectId."
        }

        if ($groupById.displayName -ne $groupName) {
            throw "Configured GroupObjectId '$configuredObjectId' points to group '$($groupById.displayName)', but expected '$groupName'."
        }

        $groupsByNameResponse = Invoke-MgGraphRequest -Method GET -Uri $groupsByNameUri -OutputType PSObject -ErrorAction Stop
        $groupsByName = @($groupsByNameResponse.value)
        if ($groupsByName.Count -gt 1) {
            throw "Multiple Azure AD groups found with display name '$groupName'. Configure a unique GroupObjectId after removing duplicates."
        }
        if ($groupsByName.Count -eq 1 -and $groupsByName[0].id -ne $configuredObjectId) {
            throw "GroupDisplayName '$groupName' resolves to ObjectId '$($groupsByName[0].id)', but configured GroupObjectId is '$configuredObjectId'."
        }

        Write-Log "Backup protection group validated: '$groupName' ($configuredObjectId)"
        return $configuredObjectId
    }

    Write-Log "Resolving backup protection group by display name: $groupName"
    $response = Invoke-MgGraphRequest -Method GET -Uri $groupsByNameUri -OutputType PSObject -ErrorAction Stop
    $groups = @($response.value)

    if ($groups.Count -gt 1) {
        throw "Multiple Azure AD groups found with display name '$groupName'. Configure GroupObjectId to disambiguate."
    }

    if ($groups.Count -eq 1) {
        Write-Log "Resolved backup protection group '$($groups[0].displayName)' to ObjectId $($groups[0].id)"
        return [string]$groups[0].id
    }

    throw "No Azure AD group found with display name '$groupName'. Create or rename the group manually, or configure GroupObjectId with a group that has this exact display name."
}

function New-BackupProtectionAiHelpUrl {
    [CmdletBinding()]
    param(
        [string]$Operation,
        [string]$ErrorMessage
    )

    $prompt = @"
Help troubleshoot this SmartM365 Exchange Online BackupProtection comparison error.
Script: $ScriptName
Tenant: $Tenant
Organization: $OrgDomain
Operation: $Operation
GroupDisplayName: $GroupDisplayName
GroupObjectId: $GroupObjectId
Error: $ErrorMessage
"@

    return "https://chat.openai.com/?q={0}" -f [System.Uri]::EscapeDataString($prompt)
}

function Send-BackupProtectionErrorNotification {
    [CmdletBinding()]
    param(
        [string]$ErrorMessage,
        [string]$Operation
    )

    try {
        Import-SmartM365CorePreflight
        $logPathValue = Get-Variable -Name LogPath -Scope Script -ValueOnly -ErrorAction SilentlyContinue
        $helpUrl = New-BackupProtectionAiHelpUrl -Operation $Operation -ErrorMessage $ErrorMessage
        $facts = @{
            "Script name"         = [System.IO.Path]::GetFileName($PSCommandPath)
            "Tenant"              = $Tenant
            "Tenant/Organization" = $OrgDomain
            "Failed operation"    = $Operation
            "Exception message"   = $ErrorMessage
            "Group DisplayName"   = $GroupDisplayName
            "Group ObjectId"      = $GroupObjectId
            "Log path"            = $logPathValue
            "Output path"         = $ScriptCsvLogFolderPath
            "Latest CSV path"     = $LatestCsvFolderPath
        }

        Send-CoreSmartM365TeamsNotification `
            -Title "SmartM365 BackupProtection comparison failed" `
            -Message "A terminal error occurred in the Exchange Online BackupProtection comparison." `
            -Level "ERROR" `
            -Channel "Alerts" `
            -Facts $facts `
            -HelpUrl $helpUrl | Out-Null
    }
    catch {
        Write-Log "Failed to send Teams error notification: $_" -Level "WARN"
    }
}

function Send-BackupProtectionSuccessNotification {
    [CmdletBinding()]
    param(
        [int]$TotalMailboxes,
        [int]$TotalGroupMembers,
        [int]$ProtectedCount,
        [int]$UnprotectedCount,
        [int]$MembersWithoutMailboxCount,
        [string]$ProtectedCsvPath,
        [string]$UnprotectedCsvPath,
        [string]$MembersWithoutMailboxCsvPath
    )

    try {
        Import-SmartM365CorePreflight
        $logPathValue = Get-Variable -Name LogPath -Scope Script -ValueOnly -ErrorAction SilentlyContinue
        $resultSummary = "BackupProtection comparison completed without error. EXO mailboxes: {0}; group members: {1}; protected: {2}; unprotected: {3}; members without mailbox: {4}." -f $TotalMailboxes, $TotalGroupMembers, $ProtectedCount, $UnprotectedCount, $MembersWithoutMailboxCount
        $facts = @{
            "Script name"              = [System.IO.Path]::GetFileName($PSCommandPath)
            "Tenant"                   = $Tenant
            "Tenant/Organization"      = $OrgDomain
            "Group DisplayName"        = $GroupDisplayName
            "Group ObjectId"           = $GroupObjectId
            "Total EXO mailboxes"      = $TotalMailboxes
            "Total group members"      = $TotalGroupMembers
            "Protected mailboxes"      = $ProtectedCount
            "Unprotected mailboxes"    = $UnprotectedCount
            "Members without mailbox"  = $MembersWithoutMailboxCount
            "Protected CSV"            = $ProtectedCsvPath
            "Unprotected CSV"          = $UnprotectedCsvPath
            "Members without mailbox CSV" = $MembersWithoutMailboxCsvPath
            "Log path"                 = $logPathValue
            "Output path"              = $ScriptCsvLogFolderPath
            "Latest CSV path"          = $LatestCsvFolderPath
        }

        Send-CoreSmartM365TeamsNotification `
            -Title "SmartM365 BackupProtection comparison success" `
            -Message $resultSummary `
            -Level "SUCCESS" `
            -Channel "Infos" `
            -ResultSummary $resultSummary `
            -Facts $facts | Out-Null
    }
    catch {
        Write-Log "Failed to send Teams completion notification: $_" -Level "WARN"
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Initialization
# ─────────────────────────────────────────────────────────────────────────────

$ScriptName    = "EXO-BackupProtection-Comparison"
$ScriptVersion = "1.0"
$Timestamp     = Get-Date -Format "yyyy-MM-dd_HHmmss"
$LogDir        = if ([string]::IsNullOrWhiteSpace($LogAllRootPath)) {
    Join-Path $ScriptCsvLogFolderPath "Logs"
} else {
    Join-Path $LogAllRootPath $ScriptName
}
$script:LogPath = Join-Path $LogDir "$ScriptName`_$Timestamp.log"

$CsvPrefix           = "Exchange_EXO_BackupProtection"
$CsvUnprotected      = Join-Path $ScriptCsvLogFolderPath     "$CsvPrefix`_UnprotectedMailboxes_$Timestamp.csv"
$CsvMembersNoMailbox = Join-Path $ScriptCsvLogFolderPath     "$CsvPrefix`_MembersWithoutMailbox_$Timestamp.csv"
$CsvProtected        = Join-Path $ScriptCsvLogFolderPath     "$CsvPrefix`_ProtectedMailboxes_$Timestamp.csv"

try {
    $script:CurrentOperation = "Create output directories"
    if (-not (Test-Path $ScriptCsvLogFolderPath))     { New-Item -ItemType Directory -Path $ScriptCsvLogFolderPath     -Force | Out-Null }
    if (-not (Test-Path $LogDir))        { New-Item -ItemType Directory -Path $LogDir        -Force | Out-Null }
    if (-not (Test-Path $LatestCsvFolderPath)) { New-Item -ItemType Directory -Path $LatestCsvFolderPath -Force | Out-Null }
} catch {
    Write-Error "Failed to create output directories: $_"
    exit 1
}

Write-Log "========================================"
Write-Log "$ScriptName v$ScriptVersion started"
Write-Log "TenantId             : $TenantId"
Write-Log "OrgDomain            : $OrgDomain"
Write-Log "AppId                : $AppId"
Write-Log "CertThumbprint       : $CertificateThumbprint"
Write-Log "GroupDisplayName     : $GroupDisplayName"
Write-Log "GroupObjectId        : $GroupObjectId"
Write-Log "EXOMailboxesCsvPath  : $(if ($EXOMailboxesCsvPath) { $EXOMailboxesCsvPath } else { '(none - live fallback)' })"
Write-Log "ScriptCsvLogFolderPath            : $ScriptCsvLogFolderPath"
Write-Log "LatestCsvFolderPath        : $LatestCsvFolderPath"
Write-Log "========================================"

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Module check
# ─────────────────────────────────────────────────────────────────────────────

$requiredModules = @("Microsoft.Graph.Groups", "Microsoft.Graph.Users")
$useLiveEXO = (-not $EXOMailboxesCsvPath) -or (-not (Test-Path $EXOMailboxesCsvPath))
$script:CurrentOperation = "Check required modules"

if ($useLiveEXO) {
    $requiredModules += "ExchangeOnlineManagement"
    Write-Log "EXO source: live (CSV not provided or not found)" -Level "WARN"
} else {
    Write-Log "EXO source: CSV file ($EXOMailboxesCsvPath)"
}

foreach ($mod in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        $err = "Required module not found: $mod. Run: Install-Module $mod -Scope CurrentUser"
        Write-Log $err -Level "ERROR"
        Send-ErrorEmail -ErrorMessage $err -ScriptName $ScriptName -GroupDisplayName $GroupDisplayName -GroupObjectId $GroupObjectId
        exit 1
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Connect to Microsoft Graph
# ─────────────────────────────────────────────────────────────────────────────

try {
    $script:CurrentOperation = "Connect to Microsoft Graph and validate backup protection group"
    Write-Log "Connecting to Microsoft Graph (app-only)..."
    try { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null } catch {}
    Connect-MgGraph -ClientId $AppId -TenantId $TenantId -CertificateThumbprint $CertificateThumbprint -NoWelcome
    Write-Log "Connected to Microsoft Graph"
    Import-SmartM365CorePreflight
    $GroupObjectId = Resolve-BackupProtectionGroupObjectId -ConfiguredGroupObjectId $GroupObjectId -ConfiguredGroupDisplayName $GroupDisplayName
    Invoke-CoreSmartM365Preflight -ScriptName $ScriptName -RequiredModules @('Microsoft.Graph.Groups', 'Microsoft.Graph.Users') -OutputPaths @($ScriptCsvLogFolderPath, $LatestCsvFolderPath) -GraphProbeUris @("https://graph.microsoft.com/v1.0/groups/$GroupObjectId/members?`$top=1") | Out-Null
} catch {
    $err = "Failed to initialize Microsoft Graph access: $_"
    Write-Log $err -Level "ERROR"
    Send-ErrorEmail -ErrorMessage $err -ScriptName $ScriptName -GroupDisplayName $GroupDisplayName -GroupObjectId $GroupObjectId
    exit 1
}

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Retrieve group members from Microsoft Graph
# ─────────────────────────────────────────────────────────────────────────────

$groupMembersMap = @{}

try {
    $script:CurrentOperation = "Retrieve backup protection group members"
    Write-Log "Retrieving members for group '$GroupDisplayName' ($GroupObjectId)..."
    $rawMembers = @(Get-MgGroupMember -GroupId $GroupObjectId -All)
    Write-Log "Raw members retrieved: $($rawMembers.Count)"
} catch {
    $err = "Failed to retrieve group members: $_"
    Write-Log $err -Level "ERROR"
    Send-ErrorEmail -ErrorMessage $err -ScriptName $ScriptName -GroupDisplayName $GroupDisplayName -GroupObjectId $GroupObjectId
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    exit 1
}

$countGroupUsers  = 0
$countGroupOther  = 0
$countGroupErrors = 0

$groupMemberObjects = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($member in $rawMembers) {
    $objectType = $member.AdditionalProperties["@odata.type"] -replace "#microsoft.graph.", ""
    $upn  = ""
    $mail = ""

    if ($objectType -eq "user") {
        try {
            $user = Get-MgUser -UserId $member.Id -Property "Id,UserPrincipalName,Mail" -ErrorAction Stop
            $upn  = $user.UserPrincipalName
            $mail = $user.Mail
            $countGroupUsers++
        } catch {
            Write-Log "Failed to resolve user $($member.Id): $_" -Level "WARN"
            $countGroupErrors++
        }
    } else {
        Write-Log "Non-user member: Id=$($member.Id) Type=$objectType" -Level "WARN"
        $countGroupOther++
    }

    $normalizedId = $member.Id.ToLower()
    $groupMembersMap[$normalizedId] = $true

    $groupMemberObjects.Add([PSCustomObject]@{
        Id                = $member.Id
        UserPrincipalName = $upn
        Mail              = $mail
        ObjectType        = $objectType
    })
}

Write-Log "Group members resolved — Users: $countGroupUsers | Other: $countGroupOther | Errors: $countGroupErrors"

try {
    Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    Write-Log "Disconnected from Microsoft Graph"
} catch {}

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Load EXO mailboxes (CSV or live)
# ─────────────────────────────────────────────────────────────────────────────

$exoMailboxes = $null

if (-not $useLiveEXO) {
    try {
        $script:CurrentOperation = "Load EXO mailboxes from CSV"
        Write-Log "Loading EXO mailboxes from CSV: $EXOMailboxesCsvPath"
        $exoMailboxes = Import-Csv -Path $EXOMailboxesCsvPath -Delimiter ";" -Encoding UTF8

        # Validate required column
        $sampleRow = $exoMailboxes | Select-Object -First 1
        if (-not ($sampleRow.PSObject.Properties.Name -contains "ExternalDirectoryObjectId")) {
            $err = "CSV file does not contain the required column 'ExternalDirectoryObjectId'. Aborting."
            Write-Log $err -Level "ERROR"
            Send-ErrorEmail -ErrorMessage $err -ScriptName $ScriptName -GroupDisplayName $GroupDisplayName -GroupObjectId $GroupObjectId
            exit 1
        }

        Write-Log "EXO mailboxes loaded from CSV: $($exoMailboxes.Count) rows"
    } catch {
        $err = "Failed to load EXO CSV: $_"
        Write-Log $err -Level "ERROR"
        Send-ErrorEmail -ErrorMessage $err -ScriptName $ScriptName -GroupDisplayName $GroupDisplayName -GroupObjectId $GroupObjectId
        exit 1
    }
} else {
    try {
        $script:CurrentOperation = "Retrieve EXO mailboxes live"
        Write-Log "Connecting to Exchange Online (app-only, live)..."
        try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null } catch {}
        Connect-ExchangeOnline -AppId $AppId -Organization $OrgDomain -CertificateThumbprint $CertificateThumbprint -ShowBanner:$false
        Write-Log "Connected to Exchange Online"
        Invoke-CoreSmartM365Preflight -ScriptName $ScriptName -ExchangeOnlineProbeCommands @('Get-EXOMailbox') | Out-Null


        Write-Log "Retrieving all EXO mailboxes..."
        $exoMailboxes = Get-EXOMailbox -ResultSize Unlimited -Properties ExternalDirectoryObjectId, UserPrincipalName, PrimarySmtpAddress, RecipientTypeDetails
        Write-Log "EXO mailboxes retrieved: $($exoMailboxes.Count)"
    } catch {
        $err = "Failed to retrieve EXO mailboxes: $_"
        Write-Log $err -Level "ERROR"
        Send-ErrorEmail -ErrorMessage $err -ScriptName $ScriptName -GroupDisplayName $GroupDisplayName -GroupObjectId $GroupObjectId
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        exit 1
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Comparison
# ─────────────────────────────────────────────────────────────────────────────

$script:CurrentOperation = "Compare backup protection membership"
Write-Log "Starting comparison..."

# Build EXO lookup map: ExternalDirectoryObjectId (lowercased) -> mailbox object
$exoMap = @{}
foreach ($mbx in $exoMailboxes) {
    $exoId = ($mbx.ExternalDirectoryObjectId -replace "\s", "").ToLower()
    if ($exoId -and $exoId -ne "") {
        $exoMap[$exoId] = $mbx
    }
}

# List 1: EXO mailboxes NOT in the backup group
$unprotected = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($exoId in $exoMap.Keys) {
    if (-not $groupMembersMap.ContainsKey($exoId)) {
        $mbx = $exoMap[$exoId]
        $unprotected.Add([PSCustomObject]@{
            ExternalDirectoryObjectId = $mbx.ExternalDirectoryObjectId
            UserPrincipalName         = $mbx.UserPrincipalName
            PrimarySmtpAddress        = $mbx.PrimarySmtpAddress
            RecipientTypeDetails      = $mbx.RecipientTypeDetails
        })
    }
}

# List 2: Group members with no matching EXO mailbox
$membersNoMailbox = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($member in $groupMemberObjects) {
    $normalizedId = $member.Id.ToLower()
    if (-not $exoMap.ContainsKey($normalizedId)) {
        $membersNoMailbox.Add([PSCustomObject]@{
            Id                = $member.Id
            UserPrincipalName = $member.UserPrincipalName
            Mail              = $member.Mail
            ObjectType        = $member.ObjectType
        })
    }
}

# List 3: EXO mailboxes present in the backup group (protected)
$protected = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($exoId in $exoMap.Keys) {
    if ($groupMembersMap.ContainsKey($exoId)) {
        $mbx = $exoMap[$exoId]
        $protected.Add([PSCustomObject]@{
            ExternalDirectoryObjectId = $mbx.ExternalDirectoryObjectId
            UserPrincipalName         = $mbx.UserPrincipalName
            PrimarySmtpAddress        = $mbx.PrimarySmtpAddress
            RecipientTypeDetails      = $mbx.RecipientTypeDetails
        })
    }
}

Write-Log "Comparison complete:"
Write-Log "  Total EXO mailboxes    : $($exoMap.Count)"
Write-Log "  Total group members    : $($groupMemberObjects.Count)"
Write-Log "  Protected mailboxes    : $($protected.Count)"
Write-Log "  Unprotected mailboxes  : $($unprotected.Count)"
Write-Log "  Members without mailbox: $($membersNoMailbox.Count)"

# ─────────────────────────────────────────────────────────────────────────────
# REGION: CSV export (atomic)
# ─────────────────────────────────────────────────────────────────────────────

$mailboxColumns = @('ExternalDirectoryObjectId', 'UserPrincipalName', 'PrimarySmtpAddress', 'RecipientTypeDetails')
$memberColumns = @('Id', 'UserPrincipalName', 'Mail', 'ObjectType')

try {
    $script:CurrentOperation = "Export UnprotectedMailboxes CSV"
    Export-CsvAtomic -Data $unprotected -Path $CsvUnprotected -Columns $mailboxColumns
    Write-Log "Exported: $CsvUnprotected ($($unprotected.Count) rows)"
} catch {
    $err = "Failed to write UnprotectedMailboxes CSV: $_"
    Write-Log $err -Level "ERROR"
    Send-ErrorEmail -ErrorMessage $err -ScriptName $ScriptName -GroupDisplayName $GroupDisplayName -GroupObjectId $GroupObjectId
    exit 1
}

try {
    $script:CurrentOperation = "Export ProtectedMailboxes CSV"
    Export-CsvAtomic -Data $protected -Path $CsvProtected -Columns $mailboxColumns
    Write-Log "Exported: $CsvProtected ($($protected.Count) rows)"
} catch {
    $err = "Failed to write ProtectedMailboxes CSV: $_"
    Write-Log $err -Level "ERROR"
    Send-ErrorEmail -ErrorMessage $err -ScriptName $ScriptName -GroupDisplayName $GroupDisplayName -GroupObjectId $GroupObjectId
    exit 1
}

try {
    $script:CurrentOperation = "Export MembersWithoutMailbox CSV"
    Export-CsvAtomic -Data $membersNoMailbox -Path $CsvMembersNoMailbox -Columns $memberColumns
    Write-Log "Exported: $CsvMembersNoMailbox ($($membersNoMailbox.Count) rows)"
} catch {
    $err = "Failed to write MembersWithoutMailbox CSV: $_"
    Write-Log $err -Level "ERROR"
    Send-ErrorEmail -ErrorMessage $err -ScriptName $ScriptName -GroupDisplayName $GroupDisplayName -GroupObjectId $GroupObjectId
    exit 1
}

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Copy latest files to DATA-LAST (fixed names, overwritten each run)
# ─────────────────────────────────────────────────────────────────────────────

$lastUnprotected      = Join-Path $LatestCsvFolderPath "$CsvPrefix`_UnprotectedMailboxes.csv"
$lastProtected        = Join-Path $LatestCsvFolderPath "$CsvPrefix`_ProtectedMailboxes.csv"
$lastMembersNoMailbox = Join-Path $LatestCsvFolderPath "$CsvPrefix`_MembersWithoutMailbox.csv"

try {
    $script:CurrentOperation = "Copy latest CSV files and upload to SharePoint"
    Copy-Item -Path $CsvUnprotected      -Destination $lastUnprotected      -Force
    Copy-Item -Path $CsvProtected        -Destination $lastProtected        -Force
    Copy-Item -Path $CsvMembersNoMailbox -Destination $lastMembersNoMailbox -Force
    Write-Log "DATA-LAST updated: $lastUnprotected"
    Write-Log "DATA-LAST updated: $lastProtected"
    Write-Log "DATA-LAST updated: $lastMembersNoMailbox"
    Invoke-CoreSmartM365SharePointCsvUpload -LocalFilePath $lastUnprotected
    Invoke-CoreSmartM365SharePointCsvUpload -LocalFilePath $lastProtected
    Invoke-CoreSmartM365SharePointCsvUpload -LocalFilePath $lastMembersNoMailbox
} catch {
    Write-Log "Failed to copy files to DATA-LAST: $_" -Level "WARN"
}

# ─────────────────────────────────────────────────────────────────────────────
# REGION: Cleanup
# ─────────────────────────────────────────────────────────────────────────────

if ($useLiveEXO) {
    try {
        $script:CurrentOperation = "Disconnect from Exchange Online"
        Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        Write-Log "Disconnected from Exchange Online"
    } catch {}
}

$script:CurrentOperation = "Send completion notification"
Send-BackupProtectionSuccessNotification `
    -TotalMailboxes $exoMap.Count `
    -TotalGroupMembers $groupMemberObjects.Count `
    -ProtectedCount $protected.Count `
    -UnprotectedCount $unprotected.Count `
    -MembersWithoutMailboxCount $membersNoMailbox.Count `
    -ProtectedCsvPath $lastProtected `
    -UnprotectedCsvPath $lastUnprotected `
    -MembersWithoutMailboxCsvPath $lastMembersNoMailbox

Write-Log "========================================"
Write-Log "$ScriptName v$ScriptVersion completed"
Write-Log "  Protected mailboxes    : $($protected.Count) -> $CsvProtected"
Write-Log "  Unprotected mailboxes  : $($unprotected.Count) -> $CsvUnprotected"
Write-Log "  Members without mailbox: $($membersNoMailbox.Count) -> $CsvMembersNoMailbox"
Write-Log "  DATA-LAST (latest)     : $LatestCsvFolderPath"
Write-Log "========================================"



