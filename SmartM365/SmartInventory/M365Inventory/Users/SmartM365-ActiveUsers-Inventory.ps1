<#
.SYNOPSIS
    Generates a detailed inventory of Microsoft 365 users with extended attributes and license information.

.DESCRIPTION
    This script connects to Microsoft Graph to retrieve user data from Microsoft 365, including synchronization status,
    licensing details, contact information, and organizational attributes. It enriches license data using a local CSV
    mapping file and exports the results to a structured CSV file. Logging, cleanup, and global error handling (with
    HTML email notification and log attachment) are handled via SmartM365.Core.

.PARAMETER OutputPath
    Specifies the output directory where CSV files will be saved. Logs are written under LogAllRootPath by SmartM365.Core.

.PARAMETER Connect
    Forces a (re)connection to Microsoft Graph (disconnects any existing Graph session first).

.PARAMETER InteractiveAuth
    Uses interactive authentication instead of app-only certificate authentication.
    Version : 1.6

.VERSION
1.8


.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Users.
    Minimum Graph application permissions: User.Read.All; AuditLog.Read.All; Directory.Read.All.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Version : 1.6
    Author: https://github.com/khda79/workplacecloudhub.com
    Requires: PowerShell 7+, Microsoft.Graph PowerShell SDK, SmartM365.Core.psd1
    Minimum application permissions: User.Read.All, Directory.Read.All, AuditLog.Read.All

.EXAMPLE
    .\SmartM365-ActiveUsers-Inventory.ps1
#>

param(
    [string]$Tenant = 'test',
[string]$OutputPath,
    [switch]$Connect,
    [switch]$InteractiveAuth,
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

# ==========================================================
# Import SmartM365.Core module (psd1)
# ==========================================================
$modulePath = & { $d = $PSScriptRoot; while ($d) { $p = Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'; if (Test-Path -LiteralPath $p) { return $p }; $parent = Split-Path -Path $d -Parent; if ($parent -eq $d) { break }; $d = $parent }; throw 'SmartM365.Core module not found.' }
try {
    Import-Module -Name $modulePath -MinimumVersion '1.0.24' -ErrorAction Stop
}
catch {
    Write-Host "Failed to import SmartM365.Core module from '$modulePath' : $_" -ForegroundColor Red
    exit 1
}

function Send-ActiveUsersErrorNotification {
    param(
        [Parameter(Mandatory)]$ErrorRecord,
        [string]$Operation,
        [string]$TimestampedCsvPath,
        [string]$CurrentCsvPath,
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
            "Log: $global:LogTextFile"
            "Transcript: $global:logTranscriptFile"
        ) -join "`n"

        $helpUrl = "https://chat.openai.com/?q={0}" -f [System.Uri]::EscapeDataString("Help troubleshoot this SmartM365 Microsoft 365 active users inventory error:`n$errorContext")

        $facts = @{
            "Script name"         = $scriptName
            "Tenant/Organization" = $OrgDomain
            "Computer"            = $env:COMPUTERNAME
            "Timestamp"           = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            "Failed operation"    = $Operation
            "Exception message"   = $exception.Message
            "Inner exception"     = ($innerMessages -join " | ")
            "Log path"            = $global:LogTextFile
            "Transcript path"     = $global:logTranscriptFile
            "Output path"         = $OutputPath
            "Timestamped CSV"     = $TimestampedCsvPath
            "Current CSV"         = $CurrentCsvPath
            "Latest CSV"          = $LatestCsvPath
        }

        Send-SmartM365TeamsNotification `
            -Title "SmartM365 Active Users inventory failed" `
            -Message "A terminal error occurred in Microsoft 365 active users inventory." `
            -Level "ERROR" `
            -Channel "Alerts" `
            -Facts $facts `
            -HelpUrl $helpUrl | Out-Null
    }
    catch {
        WriteLog -Message ("Failed to send Teams error notification: {0}" -f $_.Exception.Message) "ERROR"
    }
}

function Send-ActiveUsersSuccessNotification {
    param(
        [int]$UserCount,
        [int]$LastSignInCount,
        [int]$LastNonInteractiveSignInCount,
        [string]$TimestampedCsvPath,
        [string]$CurrentCsvPath,
        [string]$LatestCsvPath
    )

    try {
        $scriptName = [System.IO.Path]::GetFileName($PSCommandPath)
        $summary = "Microsoft 365 active users inventory completed without error. Users exported: {0}. Last sign-in values: {1}. Last non-interactive sign-in values: {2}." -f $UserCount, $LastSignInCount, $LastNonInteractiveSignInCount
        $facts = @{
            "Script name"                  = $scriptName
            "Tenant/Organization"          = $OrgDomain
            "Computer"                     = $env:COMPUTERNAME
            "Timestamp"                    = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            "Users exported"               = $UserCount
            "Last sign-in values"          = $LastSignInCount
            "Last non-interactive values"  = $LastNonInteractiveSignInCount
            "Output path"                  = $OutputPath
            "Timestamped CSV"              = $TimestampedCsvPath
            "Current CSV"                  = $CurrentCsvPath
            "Latest CSV"                   = $LatestCsvPath
            "Log path"                     = $global:LogTextFile
            "Transcript path"              = $global:logTranscriptFile
        }

        Send-SmartM365TeamsNotification `
            -Title "SmartM365 Active Users inventory completed" `
            -Message $summary `
            -Level "SUCCESS" `
            -Channel "Infos" `
            -ResultSummary $summary `
            -Facts $facts | Out-Null
    }
    catch {
        WriteLog -Message ("Failed to send Teams completion notification: {0}" -f $_.Exception.Message) "WARN"
    }
}

# ======================================================================
# GLOBAL TRY / CATCH / FINALLY WITH HTML ERROR EMAIL
# ======================================================================

$global:ScriptFailed      = $false
$connectedGraphInThisRun  = $false
$currentOperation         = "Initialize script environment"
$csvPathTimestamped       = ""
$csvPathCurrent           = ""
$csvPathLatest            = ""

try {
    #region Initialization
$ScriptVersion = "1.8"
    $TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
    $currentOperation = "Resolve output path"
    $OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ActiveUsersCsvLogFolderPath' -DefaultValue $OutputPath
    # Initialize environment (paths, logs, global variables, etc.)
    $currentOperation = "Initialize script environment"
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPath $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')

    $currentOperation = "Start transcript"
    Start-Transcript -Path $global:logTranscriptFile -Append

    WriteLog -Message "Script environment initialized at $InitializeOutputPath"
    $OutputPath = $InitializeOutputPath

    WriteLog -Message "Starting $TaskName..."

    #endregion Initialization

    #region Connect to Microsoft Graph via SmartM365.Core / Connect-SmartM365CloudSession

    $currentOperation = "Prepare Microsoft Graph connection"

    function Test-GraphConnection {
        try {
            # Works in delegated or app-only (avoid /me)
            $org = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/organization" -ErrorAction Stop
            return $true
        }
        catch {
            return $false
        }
    }

    # Detect existing Graph session
    $graphContext = $null
    if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
        try {
            $graphContext = Get-MgContext -ErrorAction SilentlyContinue
        }
        catch { }
    }

    $needConnect = $false

    if ($Connect) {
        Write-Host "Connect switch specified: existing Graph session (if any) will be disconnected and reconnected..." -ForegroundColor Cyan
        if (Get-Command Disconnect-SmartM365CloudSession -ErrorAction SilentlyContinue) {
            Disconnect-SmartM365CloudSession -ExchangeOnline $false -Graph $true -VerboseDisconnect:$true
        }
        $needConnect = $true
    }
    else {
        if ($graphContext -and (Test-GraphConnection)) {
            Write-Host "Existing Microsoft Graph session detected. Reusing current connection." -ForegroundColor Cyan
            $needConnect = $false
        }
        else {
            Write-Host "No existing Graph session detected. Will establish a new connection..." -ForegroundColor Cyan
            $needConnect = $true
        }
    }

    if ($needConnect) {
        $currentOperation = "Connect to Microsoft Graph"
        $connectParams = @{
            ExchangeOnline = $false
            Graph          = $true
            GraphScopes    = @("User.Read.All","Directory.Read.All","AuditLog.Read.All")
        }

        if (-not $InteractiveAuth) {
            # Default: app-only certificate authentication
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
        $currentOperation = "Run Microsoft Graph preflight"
        Invoke-SmartM365Preflight -ScriptName $TaskName -RequiredModules @('Microsoft.Graph.Users') -RequiredCommands @('Get-MgUser') -OutputPaths @($OutputPath) -RequiredGraphApplicationPermissions @('User.Read.All','AuditLog.Read.All','Directory.Read.All') -GraphProbeUris @(
            'https://graph.microsoft.com/v1.0/users?$select=id,signInActivity&$top=1',
            'https://graph.microsoft.com/v1.0/organization'
        ) | Out-Null
    }

    #endregion Connect to Microsoft Graph

    #region Load License Dictionary (CSV)

    WriteLog -Message "Reading local CSV file containing license information..."
    $currentOperation = "Load local license dictionary"
    $LocalCsvFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'M365InventoryRootFolderPath' -DefaultValue ''
    $licenseDictionaryFileName = "Product names and service plan identifiers for licensing.csv"
    $candidateLicenseDictionaryPaths = New-Object System.Collections.Generic.List[string]
    $configuredLicenseDictionaryPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SkuNameCsvPath' -DefaultValue ''
    if (-not [string]::IsNullOrWhiteSpace($configuredLicenseDictionaryPath)) { $candidateLicenseDictionaryPaths.Add($configuredLicenseDictionaryPath) | Out-Null }
    if (-not [string]::IsNullOrWhiteSpace($LocalCsvFolderPath)) { $candidateLicenseDictionaryPaths.Add((Join-Path $LocalCsvFolderPath $licenseDictionaryFileName)) | Out-Null }
    $candidateLicenseDictionaryPaths.Add((Join-Path $PSScriptRoot $licenseDictionaryFileName)) | Out-Null
    $candidateLicenseDictionaryPaths.Add((Join-Path (Join-Path (Split-Path -Path $PSScriptRoot -Parent) 'Licensing') $licenseDictionaryFileName)) | Out-Null
    $localCsvPath = @($candidateLicenseDictionaryPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1)[0]
    if ([string]::IsNullOrWhiteSpace($localCsvPath)) {
        $localCsvPath = @($candidateLicenseDictionaryPaths | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1)[0]
    }

    $skuMap = @{}

    try {
        if ($localCsvPath -and (Test-Path -LiteralPath $localCsvPath)) {
            $csvData = Import-Csv -LiteralPath $localCsvPath
            foreach ($row in $csvData) {
                $guid = $row.GUID.ToLower()
                $name = $row.'Product_Display_Name'
                if ($guid -and $name -and (-not $skuMap.ContainsKey($guid))) {
                    $skuMap[$guid] = $name
                }
            }
            WriteLog -Message ("License dictionary loaded from local file: {0}" -f $localCsvPath)
        }
        else {
            WriteLog -Message "Local CSV file not found: $localCsvPath" "WARNING"
        }
    }
    catch {
        WriteLog -Message ("Error while loading local CSV file: {0}" -f ($_ | Out-String)) "ERROR"
        throw
    }

    #endregion Load License Dictionary (CSV)

    #region Retrieve Users from Graph

    WriteLog -Message "Retrieving users with extended properties from Microsoft Graph..."
    $currentOperation = "Retrieve users from Microsoft Graph"

    $userSelectProperties = @(
        'DisplayName',
        'OnPremisesSyncEnabled',
        'OnPremisesImmutableId',
        'OnPremisesSecurityIdentifier',
        'UserType',
        'UserPrincipalName',
        'Id',
        'GivenName',
        'Surname',
        'CreatedDateTime',
        'DeletedDateTime',
        'JobTitle',
        'Department',
        'PreferredDataLocation',
        'City',
        'Country',
        'OfficeLocation',
        'State',
        'UsageLocation',
        'OnPremisesLastSyncDateTime',
        'AccountEnabled',
        'AssignedLicenses',
        'PasswordPolicies',
        'PasswordLastModifiedDateTime',
        'MobilePhone',
        'BusinessPhones',
        'PostalCode',
        'PreferredLanguage',
        'StreetAddress',
        'FaxNumber',
        'ProxyAddresses',
        'LastPasswordChangeDateTime',
        'SignInActivity'
    )

    if ($MaxItems -gt 0) {
        WriteLog -Message ("MaxItems enabled: retrieving at most {0} users from Graph." -f $MaxItems) "WARNING"
        $users = Get-MgUser -Top $MaxItems -Property $userSelectProperties -ErrorAction Stop
        $users = @($users | Sort-Object UserPrincipalName | Select-Object -First $MaxItems)
    }
    else {
        $users = Get-MgUser -All -Property $userSelectProperties -ErrorAction Stop
    }

    WriteLog -Message ("Number of users retrieved: {0}" -f $users.Count) "INFO"

    #endregion Retrieve Users from Graph

    #region Transform and Clean Data

    WriteLog -Message "Transforming and cleaning user data..."
    $currentOperation = "Transform user data"

    $cleanResults = $users | ForEach-Object {
        $licenseNames = ($_.AssignedLicenses | ForEach-Object {
            $skuMap[$_.SkuId.ToString().ToLower()]
        }) -join ", "

        $obj = [ordered]@{
            "Display name"                     = $_.DisplayName
            "DirSyncEnabled"                   = $_.OnPremisesSyncEnabled
            "User principal name"              = $_.UserPrincipalName
            "Object Id"                        = $_.Id
            "First name"                       = $_.GivenName
            "Last name"                        = $_.Surname
            "When created"                     = $_.CreatedDateTime
            "Soft deletion time stamp"         = $_.DeletedDateTime
            "Title"                            = $_.JobTitle
            "Department"                       = $_.Department
            "Preferred data location"          = $_.PreferredDataLocation
            "City"                             = $_.City
            "CountryOrRegion"                  = $_.Country
            "Office"                           = $_.OfficeLocation
            "StateOrProvince"                  = $_.State
            "Usage location"                   = $_.UsageLocation
            "Last dirsync time"                = $_.OnPremisesLastSyncDateTime
            "Block credential"                 = ($_.AccountEnabled -eq $false)
            "Licenses"                         = $licenseNames
            "Password never expires"           = ($_.PasswordPolicies -like "*DisablePasswordExpiration*")
            "Last password change time stamp"  = $_.LastPasswordChangeDateTime
            "Mobile Phone"                     = $_.MobilePhone
            "Phone number"                     = $_.BusinessPhones -join ","
            "Postal code"                      = $_.PostalCode
            "Preferred language"               = $_.PreferredLanguage
            "Street address"                   = $_.StreetAddress
            "Fax"                              = $_.FaxNumber
            "Proxy addresses"                  = $_.ProxyAddresses -join ","

            # --- Added columns (no changes to existing order/names above) ---
            "OnPremisesImmutableId"            = $_.OnPremisesImmutableId
            "OnPremisesSecurityIdentifier"     = $_.OnPremisesSecurityIdentifier
            "OnPremisesSyncEnabled"            = $_.OnPremisesSyncEnabled
            "AccountEnabled"                   = $_.AccountEnabled
            "UserType"                         = $_.UserType
            "LastSignInDateTime"               = $_.SignInActivity.LastSignInDateTime
            "LastNonInteractiveSignInDateTime"  = $_.SignInActivity.LastNonInteractiveSignInDateTime
        }

        # Clean string values
        $keys = @($obj.Keys)
        foreach ($key in $keys) {
            if ($obj[$key] -is [string]) {
                $obj[$key] = $obj[$key] -replace "`r`n|`n|`r", " "
                $obj[$key] = $obj[$key] -replace '"', "'"
            }
        }

        [PSCustomObject]$obj
    }

    WriteLog -Message ("Data transformation completed. Number of exported users: {0}" -f $cleanResults.Count) "INFO"

    #endregion Transform and Clean Data

    #region Export CSV

    Write-Host "`n--- Export CSV ---"
    $currentOperation = "Export CSV"
    $BaseFileName = "M365_Users_Active"

    ExportAndCopyCsvFromConvert -BaseFileName $BaseFileName `
                                -OutputPath $OutputPath `
                                -GlobalPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') `
                                -Data $cleanResults `
                                -Encoding "UTF8" `
                                -NoTypeInformation `
                                -Delimiter ","

    WriteLog -Message "CSV export completed for base file name '$BaseFileName'." "INFO"
    $csvPathTimestamped = $global:csvFilePath1
    $csvPathCurrent = $global:csvFilePath2
    $csvPathLatest = $global:csvFilePath3

    $lastSignInCount = @($cleanResults | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_."LastSignInDateTime") }).Count
    $lastNonInteractiveSignInCount = @($cleanResults | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_."LastNonInteractiveSignInDateTime") }).Count
    Send-ActiveUsersSuccessNotification `
        -UserCount $cleanResults.Count `
        -LastSignInCount $lastSignInCount `
        -LastNonInteractiveSignInCount $lastNonInteractiveSignInCount `
        -TimestampedCsvPath $csvPathTimestamped `
        -CurrentCsvPath $csvPathCurrent `
        -LatestCsvPath $csvPathLatest

    WriteLog -Message "Script main execution completed successfully." "INFO"

    #endregion Export CSV
}
catch {
    $global:ScriptFailed = $true
    $globalError         = $_

    WriteLog -Message ("Global error in M365 users inventory: {0}" -f $globalError) "ERROR"
    Write-Host "A global error occurred. Check the log file for details." -ForegroundColor Red
    Send-ActiveUsersErrorNotification -ErrorRecord $globalError -Operation $currentOperation -TimestampedCsvPath $csvPathTimestamped -CurrentCsvPath $csvPathCurrent -LatestCsvPath $csvPathLatest

    # -------- Send error email (same pattern style as Devices Inventory) --------
    try {
        $title = "M365 Users Inventory - ERROR"
        $msg   = @"
An error occurred in script $($MyInvocation.MyCommand.Name) on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss").

Error message:
$($globalError.Exception.Message)

See attached log file for details:
$($global:LogTextFile)
"@

        # Generate simple HTML body via SmartM365.Core
        $bodyHtml = NewSimpleEmailBody -Title $title -Message $msg

        # Attach log file if available
        $attachments = @()
        if ($global:LogTextFile -and (Test-Path $global:LogTextFile)) {
            $attachments = @($global:LogTextFile)
        }

        # Send HTML email with log attachment
        SendEmailHtmlReport -BodyHtml $bodyHtml -Subject $title -Attachments $attachments -VerboseLog

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
            RemoveOldFiles -Path $OutputPath -Filter "*.csv" -KeepCount $global:RetentionMaxCSV -LogFile $global:logTextFile
        }

        if ((Get-Command RemoveOldFiles -ErrorAction SilentlyContinue) -and $global:LogPath) {
            RemoveOldFiles -Path $global:LogPath -Filter "*.log" -KeepCount $global:RetentionMaxLogs -LogFile $global:logTextFile
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

    #region Stop Transcript + final status

    try {
        Stop-Transcript | Out-Null
        try {
            $smartM365TranscriptPath = $null
            $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue
            if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) {
                $smartM365TranscriptPath = $smartM365TranscriptVariable.Value
            }
            else {
                $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue
                if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) {
                    $smartM365TranscriptPath = $smartM365TranscriptVariable.Value
                }
            }
            if ($smartM365TranscriptPath) {
                Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath
            }
        }
        catch {}
    }
    catch {
        # Ignore transcript stop errors
    }

    try {
        $summaryStatus = if ($global:ScriptFailed) { 'Failed' } else { 'Auto' }
        $summaryError = if ($global:ScriptFailed) { $globalError } else { $null }
        Complete-SmartM365ExecutionContext -Status $summaryStatus -ErrorRecord $summaryError -FailureStage $currentOperation
    }
    catch {
        WriteLog -Message ("Failed to write execution summary: {0}" -f $_) "WARNING"
    }

    #endregion Stop Transcript + final status
}
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA3pylUv2K9yk7z
# ocxN8j0vYCJn7eZaKQK8R298B9rYR6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIE0FC+afHMo8GynHaFLPP/cgpt9X39acmTllNMQq3CtzMA0GCSqG
# SIb3DQEBAQUABIIBgALwEVmLkUXFzdP0+bs8R+JrALkfzSQWUKH6AlCdzUqBgvpX
# GNnS072nvaF1cUW+NEokHyUTgrqre3/9GjWIuDmpZHjcOFCktnNwoIqIt7TXxhaQ
# r2XoolAH4odJxwqAa3xjCLhkO8q6LfoTHc/wR+tORbghg39CGXJhAuWmi/D8Va6q
# 0Qd7ov/4raDN7P6dGW+DTnPKF8qFrKHXPJABOjXIEPNSHj2VHGAwkh79iu7oyolv
# b5k8YNyvx2T5eJyO+0oqmYTjQv25c0VXrr6owcxsn8A0EFCxNID+w0Sy2pDEv1ab
# ipQ/n35ZcvDLTRK5yNemDttgsAI+/EqHBJKIILK6YZ6ElZDWb14zL5ocqFA8AVhO
# LLdZwkcaBI+LBBPj+yG2CxOY1YPx081gOWshKlZmPMrOJA6yuT6Z7dviLEtPaKIO
# 3AinnDaBjIOhV2495/ZxF/LaJQYyKfRqRtK1a6NBnm1ULR00Sm9mGCKSmf6NPLqW
# cwuFRiwwV+eAu7F3RaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODUw
# MzZaMC8GCSqGSIb3DQEJBDEiBCA6v15tmpy5vblA+lc5MsaWo3DPMe9r6zrNQnkS
# WA86zzANBgkqhkiG9w0BAQEFAASCAgB87MtKXH4smvAlRtsviVhH6Q5DZEKXuy7Z
# WP2p5Ag+juh1/ZnfCHM1OsC2aF5nPSlBJjr4UMHlnVvHfC8EhWGivvNEuUBut1iU
# KgsRsupejStQoayKDAuYEZkjl92gSF9RQPPHw/x6AVjK6ger0Vf7byfAltapEloa
# Hp6q43XodTSh/ynzZNTZAG2mlcKbai9juLkvPZ1KHGYSTw/Q8NIzxCDGXqsQ8j2B
# BMyMr+FUIiCLDR2oiVFwmd9h8jjkf20BbXt8HKqfPPXqbQA4VyJxWB+1fhKPbDec
# QJBIGNbTmcTtQCs99fswnfxVEZP9HwXqgKIkNrHyTFUq4goZVxOJaqz9+AVmMGji
# AxlWP+iy7YPtUAw0pEZ2ljbhod+82om0jbFg9VhBNoitjqUW/4iVBRghU9cP3Kxh
# WF48XOLAr0GadR1SQl1Ohi6mnd8hyCegED4RO2WqE309m0Rw1ZDh1GR1/ULZR961
# Pm35o7OPiOBdda1TPl5tuuZHj8S06+RAk7aXKAwaqhaHqB3OuOwtejSJQwYm4Vet
# IfM42Jpg0ysZK8ZNZNvnhmDJBZvRugchUVcNprpArBlzwIO1D4yHG5kd9GIko/SY
# ToTMtFVat7ucGxLewXPZdOzlcO1XC/UrXEgPfZu29szeIaxAwisBm7PjJc1v4cg5
# uEDsWhBnyQ==
# SIG # End signature block
