<#
.SYNOPSIS
    Generates a detailed inventory of Microsoft 365 users with extended attributes and license information.

.DESCRIPTION
    This script connects to Microsoft Graph to retrieve user data from Microsoft 365, including synchronization status,
    licensing details, contact information, and organizational attributes. It enriches license data using a local CSV
    mapping file and exports the results to a structured CSV file. Logging, cleanup, and global error handling (with
    HTML email notification and log attachment) are handled via SmartM365.Core.

.PARAMETER OutputPath
    Specifies the output directory where the CSV file and logs will be saved. If not provided, the script will initialize
    a default path using SmartM365.Core.

.PARAMETER Connect
    Forces a (re)connection to Microsoft Graph (disconnects any existing Graph session first).

.PARAMETER InteractiveAuth
    Uses interactive authentication instead of app-only certificate authentication.

.NOTES
    Version: 1.0
    Author: https://github.com/khda79/M365
    Requires: PowerShell 7+, Microsoft.Graph PowerShell SDK, SmartM365.Core.psd1
    Scopes: User.Read.All, Directory.Read.All, AuditLog.Read.All

.EXAMPLE
    .\SmartM365-ActiveUsers-Inventory.ps1
#>

param(
    [string]$OutputPath,
    [switch]$Connect,
    [switch]$InteractiveAuth
)

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
$AppId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'AppId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$TenantId = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'TenantId' -DefaultValue '00000000-0000-0000-0000-000000000000'
$Thumb = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'Thumb' -DefaultValue '0000000000000000000000000000000000000000'
$OrgDomain = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue 'contoso.onmicrosoft.com'

# ==========================================================
# Import SmartM365.Core module (psd1)
# ==========================================================
$modulePath = & { $d = $PSScriptRoot; while ($d) { $p = Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'; if (Test-Path -LiteralPath $p) { return $p }; $parent = Split-Path -Path $d -Parent; if ($parent -eq $d) { break }; $d = $parent }; throw 'SmartM365.Core module not found.' }
try {
    Import-Module $modulePath -ErrorAction Stop
}
catch {
    Write-Host "Failed to import SmartM365.Core module from '$modulePath' : $_" -ForegroundColor Red
    exit 1
}

# ======================================================================
# GLOBAL TRY / CATCH / FINALLY WITH HTML ERROR EMAIL
# ======================================================================

$global:ScriptFailed      = $false
$connectedGraphInThisRun  = $false

try {
    #region Initialization

    $ScriptVersion = "1.0"
    $TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
    $OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ActiveUsersCsvLogFolderPath' -DefaultValue $OutputPath
    # Initialize environment (paths, logs, global variables, etc.)
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPath $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')

    Start-Transcript -Path $global:logTranscriptFile -Append

    WriteLog -Message "Script environment initialized at $InitializeOutputPath"
    $OutputPath = $InitializeOutputPath

    WriteLog -Message "Starting $TaskName..."
    WriteLog -Message "PowerShell Version: $($PSVersionTable.PSVersion)"

    #endregion Initialization

    #region Connect to Microsoft Graph via SmartM365.Core / Connect-SmartM365CloudSession

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
        Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -GraphProbeUris @(
            'https://graph.microsoft.com/v1.0/users?$top=1',
            'https://graph.microsoft.com/v1.0/organization'
        ) | Out-Null
    }

    #endregion Connect to Microsoft Graph

    #region Load License Dictionary (CSV)

    WriteLog -Message "Reading local CSV file containing license information..."
    $LocalCsvFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'M365InventoryRootFolderPath' -DefaultValue ''
    $localCsvPath       = Join-Path $LocalCsvFolderPath "Product names and service plan identifiers for licensing.csv"

    $skuMap = @{}

    try {
        if (Test-Path $localCsvPath) {
            $csvData = Import-Csv -Path $localCsvPath
            foreach ($row in $csvData) {
                $guid = $row.GUID.ToLower()
                $name = $row.'Product_Display_Name'
                if ($guid -and $name -and (-not $skuMap.ContainsKey($guid))) {
                    $skuMap[$guid] = $name
                }
            }
            WriteLog -Message "License dictionary loaded from local file."
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

    function Invoke-GraphPagedRequest {
        [CmdletBinding()]
        param([Parameter(Mandatory)][string]$Uri)

        $items = @()
        $nextUri = $Uri
        while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
            $response = Invoke-MgGraphRequest -Method GET -Uri $nextUri -ErrorAction Stop
            $items += @($response.value)
            $nextUri = $response.'@odata.nextLink'
        }

        return @($items)
    }

    $userSelectProperties = @(
        'displayName',
        'onPremisesSyncEnabled',
        'onPremisesImmutableId',
        'onPremisesSecurityIdentifier',
        'userType',
        'userPrincipalName',
        'id',
        'givenName',
        'surname',
        'createdDateTime',
        'deletedDateTime',
        'jobTitle',
        'department',
        'preferredDataLocation',
        'city',
        'country',
        'officeLocation',
        'state',
        'usageLocation',
        'onPremisesLastSyncDateTime',
        'accountEnabled',
        'assignedLicenses',
        'passwordPolicies',
        'passwordLastModifiedDateTime',
        'mobilePhone',
        'businessPhones',
        'postalCode',
        'preferredLanguage',
        'streetAddress',
        'faxNumber',
        'proxyAddresses',
        'lastPasswordChangeDateTime',
        'signInActivity'
    )

    $userSelect = $userSelectProperties -join ','
    $usersUri = "https://graph.microsoft.com/v1.0/users?`$select=$userSelect&`$top=999"

    try {
        $users = Invoke-GraphPagedRequest -Uri $usersUri
    }
    catch {
        if ($_.Exception.Message -match 'signInActivity|AuditLog|Authorization_RequestDenied|Forbidden|Insufficient privileges') {
            WriteLog -Message "Could not retrieve signInActivity. Retrying users inventory without sign-in activity columns. Add AuditLog.Read.All if those values are required." "WARNING"
            $userSelect = ($userSelectProperties | Where-Object { $_ -ne 'signInActivity' }) -join ','
            $usersUri = "https://graph.microsoft.com/v1.0/users?`$select=$userSelect&`$top=999"
            $users = Invoke-GraphPagedRequest -Uri $usersUri
        }
        else {
            throw
        }
    }

    WriteLog -Message ("Number of users retrieved: {0}" -f $users.Count) "INFO"

    #endregion Retrieve Users from Graph

    #region Transform and Clean Data

    WriteLog -Message "Transforming and cleaning user data..."

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
$BaseFileName = "M365_Users_Active"

    ExportAndCopyCsvFromConvert -BaseFileName $BaseFileName `
                                -OutputPath $OutputPath `
                                -GlobalPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') `
                                -Data $cleanResults `
                                -Encoding "UTF8" `
                                -NoTypeInformation `
                                -Delimiter ","

    WriteLog -Message "CSV export completed for base file name '$BaseFileName'." "INFO"
    WriteLog -Message "Script main execution completed successfully." "INFO"

    #endregion Export CSV
}
catch {
    $global:ScriptFailed = $true
    $globalError         = $_

    WriteLog -Message ("Global error in M365 users inventory: {0}" -f $globalError) "ERROR"
    Write-Host "A global error occurred. Check the log file for details." -ForegroundColor Red

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
            RemoveOldFiles -Path $OutputPath -Filter "*.csv" -KeepCount 150 -LogFile $global:logTextFile
        }

        if ((Get-Command RemoveOldFiles -ErrorAction SilentlyContinue) -and $logPath) {
            RemoveOldFiles -Path $logPath -Filter "*.log" -KeepCount $global:RetentionMaxLogs -LogFile $global:logTextFile
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
    }
    catch {
        # Ignore transcript stop errors
    }

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

    if ($global:ScriptFailed) {
        exit 1
    }

    #endregion Stop Transcript + final status
}


