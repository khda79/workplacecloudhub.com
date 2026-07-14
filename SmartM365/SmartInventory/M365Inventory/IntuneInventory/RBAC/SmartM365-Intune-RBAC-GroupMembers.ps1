<#
.SYNOPSIS
    Extracts members of Intune RBAC Azure AD groups for all configured countries.

.DESCRIPTION
    For each country (FR, DE, ES, PT, LU, IT, PL, BE) and each Intune RBAC role,
    the script resolves the group SG-RBAC-Intune-{Country}-{Role} via Microsoft Graph,
    retrieves all members (paginated), and exports a consolidated CSV with member details,
    role type and role description.
    Summary email delivery is disabled; unhandled errors still trigger an HTML notification.
    Logging, cleanup, and global error handling are handled via SmartM365.Core.

.PARAMETER OutputPath
    Output directory for the timestamped CSV and logs.
    Default: resolved from the local configuration file.

.PARAMETER Countries
    Array of country codes to process.
    Default: FR, DE, ES, PT, LU, IT, PL, BE

.PARAMETER Connect
    Forces a (re)connection to Microsoft Graph.

.PARAMETER InteractiveAuth
    Uses interactive authentication instead of app-only certificate authentication.

.PARAMETER DryRun
    Lists target groups without making any Graph API calls.
.VERSION
1.12



.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication.
    Minimum Graph application permissions: Group.Read.All; GroupMember.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Version : 1.11
    Author: https://github.com/khda79/workplacecloudhub.com
    Requires: PowerShell 7+, Microsoft.Graph PowerShell SDK, SmartM365.Core.psd1

.EXAMPLE
    # Dry run
    .\SmartM365-Intune-RBAC-GroupMembers.ps1 -DryRun

    # Full run - all countries
    .\SmartM365-Intune-RBAC-GroupMembers.ps1

    # Full run - subset of countries
    .\SmartM365-Intune-RBAC-GroupMembers.ps1 -Countries @("FR","DE")
    Minimum application permissions: Group.Read.All, GroupMember.Read.All
#>

param(
    [string]$Tenant = 'test',
[string]$OutputPath,
    [ValidateSet("FR","DE","ES","PT","LU","IT","PL","BE")]
    [string[]]$Countries = @("FR","DE","ES","PT","LU","IT","PL","BE"),
    [switch]$Connect,
    [switch]$InteractiveAuth,
    [switch]$DryRun,
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

if ($MaxItems -gt 0) {
    throw "-MaxItems is not supported by SmartM365-Intune-RBAC-GroupMembers because the RBAC matrix must include every configured country/role group."
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
# PowerShell 7 minimum
# ==========================================================
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    exit 1
}

$MaximumFunctionCount = 32768

# ==========================================================
# App-only authentication parameters
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
# Graph mail configuration
# ==========================================================
$MailFrom        = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'From' -DefaultValue ""
$ErrorMailTo     = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ErrorMailTo' -DefaultValue ""

# ==========================================================
# Local email helper
# ==========================================================
function Send-SmartM365Email {
    [CmdletBinding()]
    param (
        [string]$To = $ErrorMailTo,
        [string]$Subject,
        [string]$HtmlBody,
        [string[]]$Attachments = @()
    )
    try {
        Send-SmartM365Mail -From $MailFrom -To $To -Subject $Subject -BodyHtml $HtmlBody -Attachments $Attachments
    }
    catch {
        throw "Send-SmartM365Email failed: $($_.Exception.Message)"
    }
}


$ScriptCsvLogFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ScriptCsvLogFolderPath' -DefaultValue ""
$LatestCsvFolderPath      = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ""
$CsvBaseName       = "Intune_RBAC_GroupMembers"

# ==========================================================
# Import SmartM365.Core module
# ==========================================================
$modulePath = & { $d = $PSScriptRoot; while ($d) { $p = Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'; if (Test-Path -LiteralPath $p) { return $p }; $parent = Split-Path -Path $d -Parent; if ($parent -eq $d) { break }; $d = $parent }; throw 'SmartM365.Core module not found.' }
try {
    Import-Module -Name $modulePath -MinimumVersion '1.0.24' -ErrorAction Stop
}
catch {
    Write-Host "Failed to import SmartM365.Core module from '$modulePath' : $_" -ForegroundColor Red
    exit 1
}

# ==========================================================
# Role definitions (static mapping)
# ==========================================================
$RoleDefinitions = [ordered]@{
    "Policy-and-Profile-manager" = @{
        Label       = "Policy and Profile Manager"
        Type        = "Built-in Role"
        Description = "Policy and Profile Managers manage compliance policy, configuration profiles, Apple enrollment, Android Enterprise enrollment profiles and corporate device identifiers."
    }
    "Endpoint-Security-Manager" = @{
        Label       = "Endpoint Security Manager"
        Type        = "Built-in Role"
        Description = "Manages security and compliance features such as security baselines, device compliance, conditional access, and Microsoft Defender ATP."
    }
    "Windows-Autopatch-administrator" = @{
        Label       = "Windows Autopatch Administrator"
        Type        = "Built-in Role"
        Description = "Full permissions to manage updates and make tenant level changes."
    }
    "Read-Only-Operator" = @{
        Label       = "Read Only Operator"
        Type        = "Built-in Role"
        Description = "Read Only Operators view user, device, enrollment, configuration and application information and cannot make changes to Intune."
    }
    "Windows-Autopatch-reader" = @{
        Label       = "Windows Autopatch Reader"
        Type        = "Built-in Role"
        Description = "Read permissions on devices, reports, support requests, and tenant settings."
    }
    "Help-Desk-Operator" = @{
        Label       = "Help Desk Operator"
        Type        = "Built-in Role"
        Description = "Help Desk Operators perform remote tasks on users and devices and can assign applications or policies to users or devices."
    }
    "Application-Manager" = @{
        Label       = "Application Manager"
        Type        = "Built-in Role"
        Description = "Application Managers manage mobile and managed applications, can read device information and can view device configuration profiles."
    }
}

# ==========================================================
# GLOBAL TRY / CATCH / FINALLY WITH HTML ERROR EMAIL
# ==========================================================

$global:ScriptFailed     = $false
$connectedGraphInThisRun = $false

try {
    #region Initialization
$ScriptVersion = "1.12"
    $TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion"
    $OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RbacGroupMembersCsvLogFolderPath' -DefaultValue $OutputPath
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPath $ScriptCsvLogFolderPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')

    Start-Transcript -Path $global:logTranscriptFile -Append

    WriteLog -Message "Script environment initialized at $InitializeOutputPath"
    $OutputPath = $InitializeOutputPath
    $StartTime  = Get-Date

    WriteLog -Message "Starting $TaskName..."
    WriteLog -Message "Countries: $($Countries -join ', ')"
    WriteLog -Message "DryRun: $($DryRun.IsPresent)"

    #endregion Initialization

    #region DryRun

    if ($DryRun.IsPresent) {
        WriteLog -Message "DryRun mode - listing target groups only." "INFO"
        Write-Host "`n[DRY RUN] Target groups:`n" -ForegroundColor Cyan
        foreach ($Country in $Countries) {
            foreach ($RoleKey in $RoleDefinitions.Keys) {
                $GroupName = "SG-RBAC-Intune-$Country-$RoleKey"
                Write-Host "  $GroupName" -ForegroundColor Yellow
            }
        }
        $TotalGroups = $Countries.Count * $RoleDefinitions.Keys.Count
        Write-Host "`nTotal: $TotalGroups groups across $($Countries.Count) countries.`n" -ForegroundColor Cyan
        WriteLog -Message "DryRun complete. $TotalGroups groups listed." "INFO"
        exit 0
    }

    #endregion DryRun

    #region Connect to Microsoft Graph

    function Test-GraphConnection {
        try {
            Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -ErrorAction Stop | Out-Null
            return $true
        }
        catch {
            return $false
        }
    }

    $graphContext = $null
    if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
        try { $graphContext = Get-MgContext -ErrorAction SilentlyContinue } catch {}
    }

    $needConnect = $false

    if ($Connect) {
        WriteLog -Message "Connect switch specified: existing Graph session (if any) will be disconnected and reconnected." "INFO"
        if (Get-Command Disconnect-SmartM365CloudSession -ErrorAction SilentlyContinue) {
            Disconnect-SmartM365CloudSession -ExchangeOnline $false -Graph $true -VerboseDisconnect:$true
        }
        $needConnect = $true
    }
    elseif ($graphContext -and (Test-GraphConnection)) {
        WriteLog -Message "Existing Microsoft Graph session detected. Reusing current connection." "INFO"
        $needConnect = $false
    }
    else {
        WriteLog -Message "No existing Graph session detected. Will establish a new connection." "INFO"
        $needConnect = $true
    }

    if ($needConnect) {
        $connectParams = @{
            ExchangeOnline = $false
            Graph          = $true
            GraphScopes    = @("GroupMember.Read.All","Directory.Read.All")
        }

        if (-not $InteractiveAuth) {
            $connectParams.AppId        = $AppId
            $connectParams.Thumbprint   = $Thumb
            $connectParams.TenantId     = $TenantId
            $connectParams.Organization = $OrgDomain
            WriteLog -Message "Connecting to Microsoft Graph with app-only certificate authentication." "INFO"
        }
        else {
            WriteLog -Message "Connecting to Microsoft Graph with interactive authentication." "INFO"
        }

        $connectResult = Connect-SmartM365CloudSession @connectParams

        if (-not $connectResult.GraphConnected) {
            throw "Failed to connect to Microsoft Graph."
        }

        $connectedGraphInThisRun = $connectResult.GraphConnected
        WriteLog -Message "Microsoft Graph connection established successfully." "INFO"
    }

    #endregion Connect to Microsoft Graph

    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -RequiredGraphApplicationPermissions @('Group.Read.All','GroupMember.Read.All') -GraphProbeUris @('https://graph.microsoft.com/v1.0/groups?$top=1') | Out-Null

    #region Retrieve Group Members

    WriteLog -Message "Starting group member extraction for $($Countries.Count) countries x $($RoleDefinitions.Keys.Count) roles..." "INFO"

    $AllRows           = [System.Collections.Generic.List[PSCustomObject]]::new()
    $StatsGroupFound   = 0
    $StatsGroupMiss    = 0
    $StatsMembersTotal = 0

    # Per-country stats: Country -> @{ Found; Missing; Members }
    $StatsByCountry = @{}
    foreach ($C in $Countries) { $StatsByCountry[$C] = @{ Found = 0; Missing = 0; Members = 0 } }

    # Per-role stats: RoleKey -> @{ Found; Missing; Members }
    $StatsByRole = @{}
    foreach ($R in $RoleDefinitions.Keys) { $StatsByRole[$R] = @{ Found = 0; Missing = 0; Members = 0 } }

    # Per-role x per-country stats: RoleKey -> Country -> MemberCount
    $StatsByRoleAndCountry = @{}
    foreach ($R in $RoleDefinitions.Keys) {
        $StatsByRoleAndCountry[$R] = @{}
        foreach ($C in $Countries) { $StatsByRoleAndCountry[$R][$C] = 0 }
    }

    foreach ($Country in $Countries) {
        WriteLog -Message "Processing country: $Country" "INFO"

        foreach ($RoleKey in $RoleDefinitions.Keys) {
            $RoleDef   = $RoleDefinitions[$RoleKey]
            $GroupName = "SG-RBAC-Intune-$Country-$RoleKey"

            WriteLog -Message "  Resolving group: $GroupName" "INFO"

            $EncodedName   = [System.Uri]::EscapeDataString($GroupName)
            $GroupUri      = "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$EncodedName'&`$select=id,displayName"
            $GroupResponse = Invoke-MgGraphRequest -Method GET -Uri $GroupUri -ErrorAction Stop

            if (-not $GroupResponse.value -or $GroupResponse.value.Count -eq 0) {
                WriteLog -Message "  Group not found: $GroupName" "WARNING"
                $StatsGroupMiss++
                $StatsByCountry[$Country].Missing++
                $StatsByRole[$RoleKey].Missing++

                $AllRows.Add([PSCustomObject]@{
                    Country           = $Country
                    IntuneRole        = $RoleDef.Label
                    RoleType          = $RoleDef.Type
                    RoleDescription   = $RoleDef.Description
                    GroupName         = $GroupName
                    GroupFound        = "No"
                    DisplayName       = ""
                    UserPrincipalName = ""
                    UserId            = ""
                    AccountEnabled    = ""
                    JobTitle          = ""
                    Department        = ""
                })
                continue
            }

            $GroupId = $GroupResponse.value[0].id
            $StatsGroupFound++
            $StatsByCountry[$Country].Found++
            $StatsByRole[$RoleKey].Found++
            WriteLog -Message "  Group found: $GroupName (Id=$GroupId) - fetching members." "INFO"

            $Members    = [System.Collections.Generic.List[object]]::new()
            $MembersUri = "https://graph.microsoft.com/v1.0/groups/$GroupId/members?`$select=id,displayName,userPrincipalName,accountEnabled,jobTitle,department&`$top=999"

            while ($MembersUri) {
                $MembersResponse = Invoke-MgGraphRequest -Method GET -Uri $MembersUri -ErrorAction Stop
                foreach ($Member in $MembersResponse.value) { $Members.Add($Member) }
                $MembersUri = $MembersResponse.'@odata.nextLink'
            }

            WriteLog -Message "  Members retrieved: $($Members.Count)" "INFO"
            $StatsMembersTotal                         += $Members.Count
            $StatsByCountry[$Country].Members          += $Members.Count
            $StatsByRole[$RoleKey].Members             += $Members.Count
            $StatsByRoleAndCountry[$RoleKey][$Country] += $Members.Count

            if ($Members.Count -eq 0) {
                $AllRows.Add([PSCustomObject]@{
                    Country           = $Country
                    IntuneRole        = $RoleDef.Label
                    RoleType          = $RoleDef.Type
                    RoleDescription   = $RoleDef.Description
                    GroupName         = $GroupName
                    GroupFound        = "Yes"
                    DisplayName       = ""
                    UserPrincipalName = ""
                    UserId            = ""
                    AccountEnabled    = ""
                    JobTitle          = ""
                    Department        = ""
                })
            }
            else {
                foreach ($Member in $Members) {
                    $AllRows.Add([PSCustomObject]@{
                        Country           = $Country
                        IntuneRole        = $RoleDef.Label
                        RoleType          = $RoleDef.Type
                        RoleDescription   = $RoleDef.Description
                        GroupName         = $GroupName
                        GroupFound        = "Yes"
                        DisplayName       = if ($Member.displayName)       { $Member.displayName }       else { "" }
                        UserPrincipalName = if ($Member.userPrincipalName) { $Member.userPrincipalName } else { "" }
                        UserId            = if ($Member.id)                { $Member.id }                else { "" }
                        AccountEnabled    = if ($null -ne $Member.accountEnabled) { $Member.accountEnabled.ToString() } else { "" }
                        JobTitle          = if ($Member.jobTitle)          { $Member.jobTitle }          else { "" }
                        Department        = if ($Member.department)        { $Member.department }        else { "" }
                    })
                }
            }
        }
    }

    WriteLog -Message "Group extraction completed. GroupsFound=$StatsGroupFound GroupsMissing=$StatsGroupMiss TotalMembers=$StatsMembersTotal" "INFO"

    #endregion Retrieve Group Members

    #region Load Previous CSV and Compute Diff

    $PreviousCsvPath = Join-Path $LatestCsvFolderPath "$CsvBaseName.csv"
    $AddedMembers    = [System.Collections.Generic.List[PSCustomObject]]::new()
    $RemovedMembers  = [System.Collections.Generic.List[PSCustomObject]]::new()
    $DiffAvailable   = $false

    $resolvedPreviousCsvPath = Resolve-SmartM365CsvPathWithSharePointFallback -Path $PreviousCsvPath -Description 'Previous RBAC CSV'
    if (-not [string]::IsNullOrWhiteSpace($resolvedPreviousCsvPath)) {
        WriteLog -Message "Loading previous CSV for diff: $resolvedPreviousCsvPath" "INFO"

        $PreviousData    = Import-Csv -Path $resolvedPreviousCsvPath -Delimiter ";"
        $PreviousMembers = @{}
        foreach ($Row in $PreviousData) {
            if ($Row.UserPrincipalName -and $Row.GroupName) {
                $Key = "$($Row.GroupName)|$($Row.UserPrincipalName)"
                if (-not $PreviousMembers.ContainsKey($Key)) {
                    $PreviousMembers[$Key] = $Row
                }
            }
        }

        $CurrentMembers = @{}
        foreach ($Row in $AllRows) {
            if ($Row.UserPrincipalName -and $Row.GroupName) {
                $Key = "$($Row.GroupName)|$($Row.UserPrincipalName)"
                if (-not $CurrentMembers.ContainsKey($Key)) {
                    $CurrentMembers[$Key] = $Row
                }
            }
        }

        # Added: present in current but not in previous
        foreach ($Key in $CurrentMembers.Keys) {
            if (-not $PreviousMembers.ContainsKey($Key)) {
                $AddedMembers.Add($CurrentMembers[$Key])
            }
        }

        # Removed: present in previous but not in current
        foreach ($Key in $PreviousMembers.Keys) {
            if (-not $CurrentMembers.ContainsKey($Key)) {
                $RemovedMembers.Add($PreviousMembers[$Key])
            }
        }

        $DiffAvailable = $true
        WriteLog -Message "Diff computed. Added=$($AddedMembers.Count) Removed=$($RemovedMembers.Count)" "INFO"
    }
    else {
        WriteLog -Message "No previous CSV found locally or in SharePoint at $PreviousCsvPath - diff section will not be included in email." "INFO"
    }

    #endregion Load Previous CSV and Compute Diff

    #region Export CSV

    Write-Host "`n--- Export CSV ---"

    ExportAndCopyCsvFromConvert -BaseFileName $CsvBaseName `
                                -OutputPath   $OutputPath `
                                -GlobalPath   $LatestCsvFolderPath `
                                -Data         $AllRows `
                                -Encoding     "UTF8" `
                                -NoTypeInformation `
                                -Delimiter    ";"

    WriteLog -Message "CSV export completed. BaseFileName=$CsvBaseName Rows=$($AllRows.Count)" "INFO"

    #endregion Export CSV

    #region Build and Send Summary Email

    Write-Host "`n--- Send Summary Email ---"

    $RunDate  = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Duration = [math]::Round(((Get-Date) - $StartTime).TotalSeconds, 0)

    # --- Styles ---
    $StyleTable  = "border-collapse:collapse;width:100%;margin-bottom:20px;font-size:12px;"
    $StyleTh     = "padding:7px 12px;text-align:left;background:#1a5276;color:#ffffff;font-weight:bold;"
    $StyleThC    = "padding:7px 12px;text-align:center;background:#1a5276;color:#ffffff;font-weight:bold;"
    $StyleTdL    = "padding:5px 12px;border-bottom:1px solid #e0e0e0;"
    $StyleTdC    = "padding:5px 12px;border-bottom:1px solid #e0e0e0;text-align:center;"
    $StyleTdRed  = "padding:5px 12px;border-bottom:1px solid #e0e0e0;text-align:center;color:#c0392b;font-weight:bold;"
    $StyleRowAlt = "background:#f4f8fb;"

    # --- Run metadata ---
    $MetaHtml = @"
<table style="$StyleTable">
  <tr><td style="$StyleTdL font-weight:bold;width:200px;">Script</td><td style="$StyleTdL">$($MyInvocation.MyCommand.Name) v$ScriptVersion</td></tr>
  <tr style="$StyleRowAlt"><td style="$StyleTdL font-weight:bold;">Run date</td><td style="$StyleTdL">$RunDate</td></tr>
  <tr><td style="$StyleTdL font-weight:bold;">Execution time</td><td style="$StyleTdL">${Duration}s</td></tr>
  <tr style="$StyleRowAlt"><td style="$StyleTdL font-weight:bold;">Host</td><td style="$StyleTdL">$($env:COMPUTERNAME)</td></tr>
  <tr style="$StyleRowAlt"><td style="$StyleTdL font-weight:bold;">Countries processed</td><td style="$StyleTdL">$($Countries -join ", ")</td></tr>
  <tr><td style="$StyleTdL font-weight:bold;">Groups found</td><td style="$StyleTdL">$StatsGroupFound / $($StatsGroupFound + $StatsGroupMiss)</td></tr>
  <tr style="$StyleRowAlt"><td style="$StyleTdL font-weight:bold;">Groups missing</td><td style="$StyleTdL$(if ($StatsGroupMiss -gt 0) { ";color:#c0392b;font-weight:bold;" })">$StatsGroupMiss</td></tr>
  <tr><td style="$StyleTdL font-weight:bold;">Total members</td><td style="$StyleTdL">$StatsMembersTotal</td></tr>
  <tr style="$StyleRowAlt"><td style="$StyleTdL font-weight:bold;">CSV (timestamped)</td><td style="$StyleTdL font-family:monospace;font-size:11px;">$OutputPath\${CsvBaseName}_*.csv</td></tr>
  <tr><td style="$StyleTdL font-weight:bold;">CSV (latest)</td><td style="$StyleTdL font-family:monospace;font-size:11px;">$LatestCsvFolderPath\$CsvBaseName.csv</td></tr>
</table>
"@

    # --- Stats by country ---
    $CountryRowsHtml = ""
    $i = 0
    foreach ($C in $Countries) {
        $S       = $StatsByCountry[$C]
        $RowBg   = if ($i % 2 -eq 1) { $StyleRowAlt } else { "" }
        $MissTd  = if ($S.Missing -gt 0) { $StyleTdRed } else { $StyleTdC }
        $CountryRowsHtml += "<tr style='$RowBg'><td style='$StyleTdL'>$C</td><td style='$StyleTdC'>$($S.Found)</td><td style='$MissTd'>$($S.Missing)</td><td style='$StyleTdC'>$($S.Members)</td></tr>"
        $i++
    }
    $CountryHtml = @"
<table style="$StyleTable">
  <tr><th style="$StyleTh">Country</th><th style="$StyleThC">Groups Found</th><th style="$StyleThC">Groups Missing</th><th style="$StyleThC">Members</th></tr>
  $CountryRowsHtml
</table>
"@

    # --- Stats by role ---
    $RoleRowsHtml = ""
    $i = 0
    foreach ($RK in $RoleDefinitions.Keys) {
        $R      = $RoleDefinitions[$RK]
        $S      = $StatsByRole[$RK]
        $RowBg  = if ($i % 2 -eq 1) { $StyleRowAlt } else { "" }
        $MissTd = if ($S.Missing -gt 0) { $StyleTdRed } else { $StyleTdC }
        $RoleRowsHtml += "<tr style='$RowBg'><td style='$StyleTdL'>$($R.Label)</td><td style='$StyleTdL'>$($R.Type)</td><td style='$StyleTdC'>$($S.Found)</td><td style='$MissTd'>$($S.Missing)</td><td style='$StyleTdC'>$($S.Members)</td></tr>"
        $i++
    }
    $RoleHtml = @"
<table style="$StyleTable">
  <tr><th style="$StyleTh">Intune Role</th><th style="$StyleTh">Type</th><th style="$StyleThC">Groups Found</th><th style="$StyleThC">Groups Missing</th><th style="$StyleThC">Members</th></tr>
  $RoleRowsHtml
</table>
"@

    # --- Diff section ---
    $DiffHtml = ""
    if ($DiffAvailable) {
        if ($AddedMembers.Count -eq 0 -and $RemovedMembers.Count -eq 0) {
            $DiffHtml = "<p style='color:#1e8449;font-weight:bold;'>No changes detected since the previous run.</p>"
        }
        else {
            # Added members table
            $AddedHtml = ""
            if ($AddedMembers.Count -gt 0) {
                $AddedRowsHtml = ""
                $i = 0
                foreach ($M in ($AddedMembers | Sort-Object Country, IntuneRole, DisplayName)) {
                    $RowBg = if ($i % 2 -eq 1) { $StyleRowAlt } else { "" }
                    $AddedRowsHtml += "<tr style='$RowBg'><td style='$StyleTdL'>$($M.Country)</td><td style='$StyleTdL'>$($M.IntuneRole)</td><td style='$StyleTdL'>$($M.GroupName)</td><td style='$StyleTdL'>$($M.DisplayName)</td><td style='$StyleTdL'>$($M.UserPrincipalName)</td></tr>"
                    $i++
                }
                $AddedHtml = @"
<h3 style="color:#1e8449;margin-top:16px;">Members Added ($($AddedMembers.Count))</h3>
<table style="$StyleTable">
  <tr><th style="$StyleTh">Country</th><th style="$StyleTh">Intune Role</th><th style="$StyleTh">Group</th><th style="$StyleTh">Display Name</th><th style="$StyleTh">UPN</th></tr>
  $AddedRowsHtml
</table>
"@
            }

            # Removed members table
            $RemovedHtml = ""
            if ($RemovedMembers.Count -gt 0) {
                $RemovedRowsHtml = ""
                $i = 0
                foreach ($M in ($RemovedMembers | Sort-Object Country, IntuneRole, DisplayName)) {
                    $RowBg = if ($i % 2 -eq 1) { $StyleRowAlt } else { "" }
                    $RemovedRowsHtml += "<tr style='$RowBg'><td style='$StyleTdL'>$($M.Country)</td><td style='$StyleTdL'>$($M.IntuneRole)</td><td style='$StyleTdL'>$($M.GroupName)</td><td style='$StyleTdL'>$($M.DisplayName)</td><td style='$StyleTdL'>$($M.UserPrincipalName)</td></tr>"
                    $i++
                }
                $RemovedHtml = @"
<h3 style="color:#c0392b;margin-top:16px;">Members Removed ($($RemovedMembers.Count))</h3>
<table style="$StyleTable">
  <tr><th style="background:#c0392b;color:#fff;padding:7px 12px;">Country</th><th style="background:#c0392b;color:#fff;padding:7px 12px;">Intune Role</th><th style="background:#c0392b;color:#fff;padding:7px 12px;">Group</th><th style="background:#c0392b;color:#fff;padding:7px 12px;">Display Name</th><th style="background:#c0392b;color:#fff;padding:7px 12px;">UPN</th></tr>
  $RemovedRowsHtml
</table>
"@
            }

            $DiffHtml = $AddedHtml + $RemovedHtml
        }
    }
    else {
        $DiffHtml = "<p style='color:#888;font-style:italic;'>No previous CSV found - diff not available for this run.</p>"
    }

    # --- Members by Intune Role by Country (cross-table) ---
    $CrossHeaderCells = ($Countries | ForEach-Object { "<th style='$StyleThC'>$_</th>" }) -join ""
    $CrossRowsHtml = ""
    $i = 0
    foreach ($RK in $RoleDefinitions.Keys) {
        $R      = $RoleDefinitions[$RK]
        $RowBg  = if ($i % 2 -eq 1) { $StyleRowAlt } else { "" }
        $RowTotal = 0
        $Cells  = ""
        foreach ($C in $Countries) {
            $Count    = $StatsByRoleAndCountry[$RK][$C]
            $RowTotal += $Count
            $CellVal  = if ($Count -gt 0) { "$Count" } else { "<span style='color:#bbb;'>-</span>" }
            $Cells   += "<td style='$StyleTdC'>$CellVal</td>"
        }
        $CrossRowsHtml += "<tr style='$RowBg'><td style='$StyleTdL'>$($R.Label)</td>$Cells<td style='$StyleTdC font-weight:bold;'>$RowTotal</td></tr>"
        $i++
    }
    # Grand total row
    $GrandTotalCells = ""
    foreach ($C in $Countries) {
        $ColTotal = 0
        foreach ($RK in $RoleDefinitions.Keys) { $ColTotal += $StatsByRoleAndCountry[$RK][$C] }
        $GrandTotalCells += "<td style='$StyleTdC font-weight:bold;'>$ColTotal</td>"
    }
    $CrossRowsHtml += "<tr style='background:#e8f0f7;'><td style='$StyleTdL font-weight:bold;'>Total</td>$GrandTotalCells<td style='$StyleTdC font-weight:bold;'>$StatsMembersTotal</td></tr>"

    $CrossHtml = @"
<table style="$StyleTable">
  <tr><th style="$StyleTh">Intune Role</th>$CrossHeaderCells<th style="$StyleThC">Total</th></tr>
  $CrossRowsHtml
</table>
"@
    $BodyHtml = @"
<!DOCTYPE html>
<html>
<body style="font-family:Arial,sans-serif;font-size:13px;color:#222;max-width:900px;">

<h2 style="color:#1a5276;border-bottom:2px solid #1a5276;padding-bottom:6px;">
  Intune RBAC Group Members Inventory - Run Summary
</h2>

<h3 style="color:#1a5276;">Members by Intune Role</h3>
$RoleHtml

<h3 style="color:#1a5276;">Members by Intune Role by Country</h3>
$CrossHtml

<h3 style="color:#1a5276;border-top:2px solid #1a5276;padding-top:10px;">Changes Since Previous Run</h3>
$DiffHtml

<h3 style="color:#1a5276;border-top:2px solid #1a5276;padding-top:10px;">Run Details</h3>
$MetaHtml

</body>
</html>
"@

    # --- Determine subject line ---
    $SubjectStatus = if ($StatsGroupMiss -gt 0) { "WARNING - $StatsGroupMiss missing groups" } else { "OK" }
    $SubjectDiff   = if ($DiffAvailable -and ($AddedMembers.Count -gt 0 -or $RemovedMembers.Count -gt 0)) {
        " | +$($AddedMembers.Count) / -$($RemovedMembers.Count) members"
    } else { "" }
    $EmailSubject  = "[SmartM365] Intune RBAC GroupMembers | $SubjectStatus | $StatsMembersTotal members$SubjectDiff | $RunDate"

    # --- Generate Excel attachment ---
    $ExcelFile        = $null
    $CountryExcelFiles = [System.Collections.Generic.List[string]]::new()

    if (Get-Command Export-Excel -ErrorAction SilentlyContinue) {
        $ExcelTimestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

        # Consolidated Excel
        try {
            $ExcelFile = Join-Path $OutputPath "${CsvBaseName}_${ExcelTimestamp}.xlsx"
            $AllRows | Export-Excel -Path $ExcelFile `
                -WorksheetName "RBAC_Members" `
                -AutoSize -FreezeTopRow -BoldTopRow `
                -TableName "RBAC_Members" -TableStyle Medium2
            WriteLog -Message "Excel (consolidated) generated: $ExcelFile" "INFO"
        }
        catch {
            WriteLog -Message "Failed to generate consolidated Excel: $($_.Exception.Message)" "WARN"
            $ExcelFile = $null
        }

        # Per-country Excel files
        foreach ($Country in $Countries) {
            $CountryRows = $AllRows | Where-Object { $_.Country -eq $Country }
            if (-not $CountryRows) { continue }
            try {
                $CountryExcelFile = Join-Path $OutputPath "${CsvBaseName}_${Country}_${ExcelTimestamp}.xlsx"
                $CountryRows | Export-Excel -Path $CountryExcelFile `
                    -WorksheetName "RBAC_${Country}" `
                    -AutoSize -FreezeTopRow -BoldTopRow `
                    -TableName "RBAC_${Country}" -TableStyle Medium2
                $CountryExcelFiles.Add($CountryExcelFile)
                WriteLog -Message "Excel ($Country) generated: $CountryExcelFile" "INFO"
            }
            catch {
                WriteLog -Message "Failed to generate Excel for ${Country}: $($_.Exception.Message)" "WARN"
            }
        }
    }
    else {
        WriteLog -Message "ImportExcel module not found - Excel attachments skipped." "WARN"
    }

    $attachments = [System.Collections.Generic.List[string]]::new()
    if ($ExcelFile -and (Test-Path $ExcelFile)) { $attachments.Add($ExcelFile) }
    foreach ($F in $CountryExcelFiles) { if (Test-Path $F) { $attachments.Add($F) } }

    Write-Log "Summary email disabled; error emails only." "INFO"

    WriteLog -Message "Script main execution completed successfully." "INFO"

    #endregion Build and Send Summary Email
}
catch {
    $global:ScriptFailed = $true
    $globalError         = $_

    WriteLog -Message ("Global error in Intune RBAC Group Members Inventory: {0}" -f $globalError) "ERROR"
    Write-Host "A global error occurred. Check the log file for details." -ForegroundColor Red

    try {
        $ErrorSubject = "[SmartM365] Intune RBAC Group Members Inventory - ERROR - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        $ErrorHtml    = @"
<!DOCTYPE html><html><body style="font-family:Arial,sans-serif;font-size:13px;color:#222;">
<h2 style="color:#c0392b;">Intune RBAC Group Members Inventory - ERROR</h2>
<table style="border-collapse:collapse;">
  <tr><td style="padding:4px 10px;font-weight:bold;">Script</td><td style="padding:4px 10px;">$($MyInvocation.MyCommand.Name) v$ScriptVersion</td></tr>
  <tr><td style="padding:4px 10px;font-weight:bold;">Time</td><td style="padding:4px 10px;">$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</td></tr>
  <tr><td style="padding:4px 10px;font-weight:bold;">Host</td><td style="padding:4px 10px;">$($env:COMPUTERNAME)</td></tr>
  <tr><td style="padding:4px 10px;font-weight:bold;color:#c0392b;">Error</td><td style="padding:4px 10px;color:#c0392b;">$([System.Net.WebUtility]::HtmlEncode($globalError.Exception.Message))</td></tr>
</table>
<p>See attached log file for details.</p>
</body></html>
"@
        $attachments = @()
        Send-SmartM365Email -To $ErrorMailTo -Subject $ErrorSubject -HtmlBody $ErrorHtml -Attachments $attachments
        WriteLog -Message "Error notification email successfully sent." "INFO"
    }
    catch {
        WriteLog -Message ("Failed to send error notification email: {0}" -f $_) "ERROR"
    }
}
finally {
    if (Get-Command WriteLog -ErrorAction SilentlyContinue) {
        WriteLog -Message "Entering FINALLY block..." "INFO"
    }

    #region Disconnect Cloud Services

    try {
        if ($connectedGraphInThisRun -and (Get-Command Disconnect-SmartM365CloudSession -ErrorAction SilentlyContinue)) {
            Write-Host "`n--- Disconnect Cloud Services ---"
            Disconnect-SmartM365CloudSession -ExchangeOnline $false -Graph $true
            if (Get-Command WriteLog -ErrorAction SilentlyContinue) {
                WriteLog -Message "Disconnected from cloud services." "INFO"
            }
        }
    }
    catch {
        if (Get-Command WriteLog -ErrorAction SilentlyContinue) {
            WriteLog -Message ("Disconnect error: {0}" -f ($_ | Out-String)) "ERROR"
        }
        else {
            Write-Host ("Disconnect error: {0}" -f ($_ | Out-String)) -ForegroundColor Yellow
        }
    }

    #endregion Disconnect Cloud Services

    #region Cleanup

    try {
        if ((Get-Command RemoveOldFiles -ErrorAction SilentlyContinue) -and $OutputPath) {
            RemoveOldFiles -Path $OutputPath -Filter "*.csv" -KeepCount 150 -LogFile $global:logTextFile
        }
        if (Get-Command WriteLog -ErrorAction SilentlyContinue) {
            WriteLog -Message "Cleanup completed." "INFO"
        }
    }
    catch {
        if (Get-Command WriteLog -ErrorAction SilentlyContinue) {
            WriteLog -Message ("Cleanup error: {0}" -f ($_ | Out-String)) "ERROR"
        }
        else {
            Write-Host ("Cleanup error: {0}" -f ($_ | Out-String)) -ForegroundColor Yellow
        }
    }

    #endregion Cleanup

    #region Stop Transcript

    try { Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {} } catch {}

    if (Get-Command WriteLog -ErrorAction SilentlyContinue) {
        if ($global:ScriptFailed) {
            WriteLog -Message "SCRIPT FINISHED WITH ERRORS" "ERROR"
        }
        else {
            WriteLog -Message "SCRIPT COMPLETED SUCCESSFULLY" "SUCCESS"
        }
    }
    else {
        if ($global:ScriptFailed) {
            Write-Host "SCRIPT FINISHED WITH ERRORS" -ForegroundColor Red
        }
        else {
            Write-Host "SCRIPT COMPLETED SUCCESSFULLY" -ForegroundColor Green
        }
    }
    Complete-SmartM365ExecutionContext -Status Auto

    #endregion Stop Transcript
}
if ($global:ScriptFailed) {
    exit 1
}
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCH1AGvbrx51Lz+
# zh91nZhace5E8Ik7X5UH2ktISTWclaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILqqkps0wTNpFBfawWjpPic4J4sLftni29zxj4BJKGQGMA0GCSqG
# SIb3DQEBAQUABIIBgE9lSE0HPY0Euh2uvPIvsAy3M5sxS/0wYopXRthMnB7nbgtm
# aeEbz9fGBrO98sLGJ6L5XspcwABK43bfb6NaUVt90MfpeiU0NwI5GUcQDZm00yYV
# usv5gYKNyLnO3eZY4ic/d9xseG/frpGYZS566NdL8GBnBFewrxl6sO0O3RbuV8zK
# 58j5Buq49gDYz3AJ6n/2cJx5Ib0IMY3vBLYZhc4+/wzHeV2df4/Ha9x5ZmMbYN9H
# mfrh3ayd5T2DzW7NRwaNa/JLSYw59BY7idGrYzahrb2WB+3Ng6pETopqBqJJVFQN
# /aeHK1cwYYQpJCi9y9oXnScKkO4mQl6vIJGL5EKK6/UeruKb91eAZCCGGxrJvbKl
# DUo/BXb76yP/APnjyTDuhajKpQsfUpzt+tARAHvyRjZi+8TUGarpJ50MkPOUSS0G
# Pjp26B5YEFoQeeK4VN4z2Aqf+XxVH7i19IdzEGD8q0wt7VRg1AEQ3OsdE6aHK/04
# agXJImNuGf+/PIGjMqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTQwMTEy
# NTZaMC8GCSqGSIb3DQEJBDEiBCB1BYQ4YJG5eKp8Ds5RAWGvAC6FGGlERDD108OP
# IA32YDANBgkqhkiG9w0BAQEFAASCAgAydNNg8pSyPK8BQrAADQPRoAegmbgwSv5I
# l1fpXfTzMLVA+7a5Z5qiACHVWxUXLFPFbnWXjIL3vrt29p4+n9uSPyaw29yY+7K8
# PWsMxryCjuf2ZFJoRj95ryEtfLYt6c4HzZ0F63Vmg25QuJpjmcFzLTfP7GpLhxRZ
# 9k/biXhuVAEuxxCj6Z411XNuuh4Abye4S1ezruQIiv4ELRBl8/KPXBQNr+ljQvIo
# PugYV/8CEocMtBV5b6G5ETkf5spUH/OZwW0aTli4+14q8JdJUkIRkfwuuASk3qfu
# t9DgN9WATUtu+10dvhWuHxWIs28mxM7gR+FT62kVnAnta+IaEKI+gInvjC+dpxsy
# BaA7mQDs+mV4Fi5dzxZqwUzDFVoPXfwkvZkySL0CO4HRnJwhyIiLhnBx33X59SrV
# dtBrhWw/UE6lS9XMDDWgb9ugXf37aDnCTzbrnGebu98RsLiUkhmCDFv978schetR
# tWxzn9UfCl2JDPKyWyK+wIELz4k6XGw2YEfCYWy4vG//SFGYzrNH3PmEsmXq9tp4
# mvie7RZlg57du5u7maccLip8uOw/2XTT8CV9W+XLO70KoPKqWxuWGTPnAMst0BW1
# RS0IsMLpW66Zb/Uiz8kaKhZiuVgp56g3AhMLplbc3DZKK9bRE9kqXBZIcAhX4we3
# Bx56nIzUNQ==
# SIG # End signature block
