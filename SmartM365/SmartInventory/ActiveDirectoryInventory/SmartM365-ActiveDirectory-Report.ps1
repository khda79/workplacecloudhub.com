<#
.SYNOPSIS
Generates a daily report of Active Directory computer objects across all domains in the forest, filtered by operating system and account status.

.DESCRIPTION
This script reuses fresh Active Directory inventory CSV files from LatestCsvFolderPath when available, or collects computer
and user objects from all domains in the Active Directory forest (or a subset specified via -TargetDomains) as a fallback.
Computer objects are filtered based on specified operating systems (Windows 7, 8, 10, 11) and categorized by enabled/disabled status.
It then generates a CSV report summarizing the data per domain, including OS distribution.

The script:
- Reuses fresh AD inventory CSV files when UseLatestInventoryCsvForReport is enabled
- Falls back to live AD reads when the inventory CSV files are missing or older than LatestInventoryCsvMaxAgeMinutes
- Uses the shared initialization framework (SmartM365.Core / InitializeScriptEnvironment)
- Supports optional output path configuration via local configuration or parameter
- Logs to both text log and transcript
- Copies the final CSV to an additional local configuration path (LatestCsvFolderPath)
- Uploads generated CSV files to SharePoint when enabled in configuration
- Cleans up old CSV and log files
- Sends an email notification in case of a global error (using SendEmailHtmlReport)

.VERSION
1.0

.AUTHOR
    https://github.com/khda79/workplacecloudhub.com
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
[Parameter(Mandatory = $false)]
    [string[]]$TargetDomains,

    [Parameter(Mandatory = $false)]
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
$UseLatestInventoryCsvForReport = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'UseLatestInventoryCsvForReport' -DefaultValue $true)
$LatestInventoryCsvMaxAgeMinutes = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestInventoryCsvMaxAgeMinutes' -DefaultValue 720)
# ==========================================================
# PowerShell 7 minimum
# ==========================================================
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    exit 1
}

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

# ==========================================================
# Initialization via SmartM365.Core
# ==========================================================
$ScriptVersion = "1.0"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ActiveDirectoryReportCsvLogFolderPath' -DefaultValue $OutputPath
try {
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPathInit $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
    Start-Transcript -Path $global:LogTranscriptFile -Append

    WriteLog -Message "Script environment initialized at $InitializeOutputPath"

#region Daily Success Lock
$LockRoot     = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LockRoot' -DefaultValue "C:\ProgramData\SmartM365\Locks"
$DailyLockName = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'DailyLockName' -DefaultValue "SmartM365-ActiveDirectory-Report"
$Today        = Get-Date -Format "yyyy-MM-dd"
$DailyLockFile = Join-Path $LockRoot "$DailyLockName-SUCCESS-$Today.lock"

if (!(Test-Path $LockRoot)) {
    New-Item -Path $LockRoot -ItemType Directory -Force | Out-Null
}

if (Test-Path $DailyLockFile) {
    WriteLog -Message "[$DailyLockName] Already succeeded on $Today. Skipping execution." "INFO"
    exit 0
}

$ScriptSucceeded = $false
#endregion

    $OutputPath = $InitializeOutputPath
    WriteLog -Message "Starting $TaskName..."
    WriteLog -Message "PowerShell Version: $($PSVersionTable.PSVersion)"
}
catch {
    Write-Host "Initialization failed: $_" -ForegroundColor Red
    exit 1
}

function ConvertTo-ReportBoolean {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -is [bool]) {
        return $Value
    }

    if ($null -eq $Value) {
        return $false
    }

    return ([string]$Value).Trim() -eq 'True'
}

function ConvertTo-ReportDate {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    if ($Value -is [datetime]) {
        return $Value
    }

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return $null
    }

    $parsedDate = [datetime]::MinValue
    if ([datetime]::TryParse([string]$Value, [ref]$parsedDate)) {
        return $parsedDate
    }

    return $null
}

function Get-ReportSimpleOS {
    [CmdletBinding()]
    param([AllowNull()][string]$OperatingSystem)

    switch -Wildcard ($OperatingSystem) {
        "*Windows 11*" { return "Windows 11" }
        "*Windows 10*" { return "Windows 10" }
        "*Windows 8*"  { return "Windows 8" }
        "*Windows 7*"  { return "Windows 7" }
        default         { return "Unknown" }
    }
}

function Test-ReportWildcardMatch {
    [CmdletBinding()]
    param(
        [AllowNull()][string]$Value,
        [string[]]$Patterns
    )

    foreach ($pattern in @($Patterns)) {
        if ($Value -like $pattern) {
            return $true
        }
    }

    return $false
}

function Get-LatestInventoryReportSource {
    [CmdletBinding()]
    param(
        [string]$LatestCsvFolderPath,
        [string[]]$AllowedOS,
        [string[]]$TargetDomains,
        [int]$MaxAgeMinutes
    )

    if ([string]::IsNullOrWhiteSpace($LatestCsvFolderPath)) {
        WriteLog -Message "Latest inventory CSV source skipped: LatestCsvFolderPath is empty." "INFO"
        return $null
    }

    $computersCsv = Join-Path -Path $LatestCsvFolderPath -ChildPath "AD_Computers_AllDomains.csv"
    $usersCsv = Join-Path -Path $LatestCsvFolderPath -ChildPath "AD_Users_AllDomains.csv"

    foreach ($path in @($computersCsv, $usersCsv)) {
        if (-not (Test-Path -LiteralPath $path)) {
            WriteLog -Message ("Latest inventory CSV source unavailable. Missing file: {0}" -f $path) "INFO"
            return $null
        }
    }

    if ($MaxAgeMinutes -gt 0) {
        $freshAfter = (Get-Date).AddMinutes(-1 * $MaxAgeMinutes)
        foreach ($path in @($computersCsv, $usersCsv)) {
            $item = Get-Item -LiteralPath $path -ErrorAction Stop
            if ($item.LastWriteTime -lt $freshAfter) {
                WriteLog -Message ("Latest inventory CSV source is too old. File: {0}; LastWriteTime: {1}; MaxAgeMinutes: {2}" -f $path, $item.LastWriteTime, $MaxAgeMinutes) "INFO"
                return $null
            }
        }
    }

    $targetDomainSet = @{}
    foreach ($domain in @($TargetDomains | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $targetDomainSet[$domain.ToLowerInvariant()] = $true
    }
    $hasTargetDomainFilter = $targetDomainSet.Count -gt 0

    $computers = @(Import-Csv -LiteralPath $computersCsv -Encoding UTF8 | Where-Object {
        $domainName = [string]$_.DomainName
        ((-not $hasTargetDomainFilter) -or $targetDomainSet.ContainsKey($domainName.ToLowerInvariant())) -and
        (Test-ReportWildcardMatch -Value ([string]$_.OperatingSystem) -Patterns $AllowedOS)
    } | ForEach-Object {
        [PSCustomObject]@{
            Enabled                = ConvertTo-ReportBoolean $_.Enabled
            SimpleOS               = Get-ReportSimpleOS -OperatingSystem ([string]$_.OperatingSystem)
            ADDomain               = [string]$_.DomainName
            OperatingSystem        = [string]$_.OperatingSystem
            OperatingSystemVersion = [string]$_.OperatingSystemVersion
            LastLogonDate          = ConvertTo-ReportDate $_.LastLogonDate
        }
    })

    $users = @(Import-Csv -LiteralPath $usersCsv -Encoding UTF8 | Where-Object {
        $domainName = [string]$_.DomainName
        ((-not $hasTargetDomainFilter) -or $targetDomainSet.ContainsKey($domainName.ToLowerInvariant())) -and
        (-not [string]::IsNullOrWhiteSpace([string]$_.UserPrincipalName))
    } | ForEach-Object {
        [PSCustomObject]@{
            ADDomain          = [string]$_.DomainName
            Enabled           = ConvertTo-ReportBoolean $_.Enabled
            LastLogonDate     = ConvertTo-ReportDate $_.LastLogonDate
            UserPrincipalName = [string]$_.UserPrincipalName
        }
    })

    if ($computers.Count -eq 0 -and $users.Count -eq 0) {
        WriteLog -Message "Latest inventory CSV source produced no reportable rows. Falling back to Active Directory." "INFO"
        return $null
    }

    $domains = @(@($computers | Select-Object -ExpandProperty ADDomain) + @($users | Select-Object -ExpandProperty ADDomain) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object -Unique)
    WriteLog -Message ("Using latest inventory CSV source. Computers: {0}; Users: {1}; Domains: {2}" -f $computers.Count, $users.Count, ($domains -join ', ')) "INFO"

    return [PSCustomObject]@{
        Domains   = $domains
        Computers = $computers
        Users     = $users
    }
}

# ==========================================================
# MAIN TRY / CATCH / FINALLY
# ==========================================================
try {
    Write-Output "------------------------------------------------------------"
    Write-Output "Starting Active Directory Computer Report script at $(Get-Date)."
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) | Out-Null

    # --- FILTERS ---
    $allowedOS = @('Windows 7*', 'Windows 8*', 'Windows 10*', 'Windows 11*')
    $localCopyPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''

    # Optional OU filter (applied identically to each domain if specified)
    $ou = ""

    Write-Output "Filtering on Operating Systems: $($allowedOS -join ', ')"
    if (-not [string]::IsNullOrEmpty($ou)) {
        Write-Output "Filtering on specific OU: $ou"
    }

    $latestInventorySource = $null
    if ($UseLatestInventoryCsvForReport) {
        $latestInventorySource = Get-LatestInventoryReportSource -LatestCsvFolderPath $localCopyPath -AllowedOS $allowedOS -TargetDomains $TargetDomains -MaxAgeMinutes $LatestInventoryCsvMaxAgeMinutes
    }
    else {
        WriteLog -Message "Latest inventory CSV source disabled by configuration." "INFO"
    }

    if ($null -ne $latestInventorySource) {
        Write-Output "Step 1/4: Using fresh inventory CSV files from LatestCsvFolderPath."
        $allDomains = @($latestInventorySource.Domains)
        $processedComputers = @($latestInventorySource.Computers)
        $allUsers = @($latestInventorySource.Users)
    }
    else {
        #region Load Active Directory module
        Write-Output "Checking for Active Directory module availability..."
        if (-not (Get-Command Get-ADComputer -ErrorAction SilentlyContinue)) {
            throw "The Get-ADComputer cmdlet is not available. This script must be run on a server with RSAT for Active Directory (AD DS) tools installed."
        }

        Import-Module ActiveDirectory
        Write-Output "Active Directory module loaded successfully."
        Invoke-SmartM365Preflight -ScriptName $TaskName -RequiredModules @('ActiveDirectory') -RequireActiveDirectoryRead | Out-Null
        #endregion

        # Splatting parameters for Get-ADComputer
        $getComputerParams = @{
            Properties    = 'OperatingSystem', 'OperatingSystemVersion', 'Enabled', 'LastLogonDate'
            ResultSetSize = $null
            ErrorAction   = 'SilentlyContinue'
        }

        $osFilter = ($allowedOS | ForEach-Object { "(OperatingSystem -like '$_')" }) -join " -or "
        $getComputerParams['Filter'] = $osFilter

        # --- STEP 1: DATA GATHERING ACROSS THE FOREST (OR TARGET DOMAINS) ---
        Write-Output "Step 1/4: Getting all computers from Active Directory..."

        $forestDomains = (Get-ADForest).Domains
        Write-Output "Forest domains discovered: $($forestDomains -join ', ')"

        if ($TargetDomains -and $TargetDomains.Count -gt 0) {
            $allDomains = @()
            foreach ($td in $TargetDomains) {
                if ($forestDomains -contains $td) {
                    $allDomains += $td
                }
                else {
                    Write-Warning ("Target domain '{0}' is not part of the forest and will be ignored." -f $td)
                }
            }

            if ($allDomains.Count -eq 0) {
                throw "None of the specified TargetDomains are part of the forest. Aborting."
            }

            Write-Output "Using filtered domains: $($allDomains -join ', ')"
        }
        else {
            $allDomains = $forestDomains
            Write-Output "No TargetDomains specified. Using all forest domains: $($allDomains -join ', ')"
        }

        Write-Output "Processing $($allDomains.Count) domain(s)."

        $allComputers = foreach ($domain in $allDomains) {
            Write-Output "--> Searching in domain: $domain"
            $getComputerParams['Server'] = $domain

            if (-not [string]::IsNullOrEmpty($ou)) {
                $getComputerParams['SearchBase'] = $ou
            }
            else {
                $getComputerParams.Remove('SearchBase')
            }

            $computersInDomain = Get-ADComputer @getComputerParams

            if ($null -ne $computersInDomain -and $computersInDomain.Count -gt 0) {
                $count = $computersInDomain.Count
                Write-Output "    Found $count computers in $domain."

                $computersInDomain | Add-Member -MemberType NoteProperty -Name "ADDomain" -Value $domain -PassThru -Force
            }
            else {
                Write-Output "    Found 0 computers in $domain."
            }
        }

        if ($null -eq $allComputers) {
            $allComputers = @()
        }

        Write-Output "Successfully found $($allComputers.Count) total computer objects across the selected domains."

        # Filter out objects with no valid ADDomain
        $initialCount  = $allComputers.Count
        $allComputers  = $allComputers | Where-Object { -not [string]::IsNullOrEmpty($_.ADDomain) }
        $filteredCount = $initialCount - $allComputers.Count

        if ($filteredCount -gt 0) {
            Write-Warning ("{0} object(s) were ignored because they could not be associated with a valid domain (likely corrupted or non-standard entries)." -f $filteredCount)
        }

        Write-Output "Proceeding with $($allComputers.Count) valid computer objects."

        if ($allComputers.Count -eq 0) {
            Write-Warning "No computers were found that match the filter criteria. The report will not be generated."
            return
        }

        # --- STEP 2: DATA PROCESSING ---
        Write-Output "Step 2/4: Processing and standardizing data..."

        $processedComputers = $allComputers | ForEach-Object {
            $osName = $_.OperatingSystem

            [PSCustomObject]@{
                Enabled                = $_.Enabled
                SimpleOS               = Get-ReportSimpleOS -OperatingSystem $osName
                ADDomain               = $_.ADDomain
                OperatingSystem        = $osName
                OperatingSystemVersion = $_.OperatingSystemVersion
                LastLogonDate          = $_.LastLogonDate
            }
        }

        # --- USERS DAILY STATS SOURCE (per domain, UPN required) ---
        Write-Output "Getting all users with non-empty UPN from Active Directory..."

        $getUserParams = @{ Properties = 'Enabled','LastLogonDate','UserPrincipalName'; ResultSetSize = $null; ErrorAction = 'SilentlyContinue' }

        $allUsers = foreach ($domain in $allDomains) {
            Write-Output "--> Searching users in domain: $domain"
            $getUserParams['Server'] = $domain
            $usersInDomain = Get-ADUser -Filter * @getUserParams | Where-Object { -not [string]::IsNullOrWhiteSpace($_.UserPrincipalName) }
            if ($null -ne $usersInDomain -and $usersInDomain.Count -gt 0) { $usersInDomain | Add-Member -MemberType NoteProperty -Name "ADDomain" -Value $domain -PassThru -Force }
            else { Write-Output "    Found 0 users with UPN in $domain." }
        }

        if ($null -eq $allUsers) { $allUsers = @() }
    }

    if ($processedComputers.Count -eq 0) {
        Write-Warning "No computers were found that match the filter criteria. The report will not be generated."
        return
    }

    # --- STEP 3: REPORT GENERATION ---
    Write-Output "Step 3/4: Grouping data and generating the report..."

    $groupedByDomain = $processedComputers | Group-Object -Property ADDomain
    $osReportColumns = @('Windows 7', 'Windows 8', 'Windows 10', 'Windows 11')

    $report = $groupedByDomain | ForEach-Object {
        $domainName       = $_.Name
        $computersInGroup = $_.Group

        $outputRow = [ordered]@{
            "Date"             = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            "DomainName"       = $domainName
            "TotalComputers"   = $computersInGroup.Count
            "EnabledAccounts"  = ($computersInGroup | Where-Object { $_.Enabled -eq $true }).Count
            "DisabledAccounts" = ($computersInGroup | Where-Object { $_.Enabled -eq $false }).Count
        }

        $enabledComputersInGroup  = $computersInGroup | Where-Object { $_.Enabled -eq $true }
        $disabledComputersInGroup = $computersInGroup | Where-Object { $_.Enabled -eq $false }

        $countsByEnabledOS  = $enabledComputersInGroup  | Group-Object -Property SimpleOS -NoElement
        $countsByDisabledOS = $disabledComputersInGroup | Group-Object -Property SimpleOS -NoElement

        foreach ($os in $osReportColumns) {
            $enabledOsGroup  = $countsByEnabledOS  | Where-Object { $_.Name -eq $os }
            $disabledOsGroup = $countsByDisabledOS | Where-Object { $_.Name -eq $os }

            $outputRow["$os Enabled"]  = if ($null -ne $enabledOsGroup)  { $enabledOsGroup.Count }  else { 0 }
            $outputRow["$os Disabled"] = if ($null -ne $disabledOsGroup) { $disabledOsGroup.Count } else { 0 }
        }

        [PSCustomObject]$outputRow
    }

    Write-Output "Report object generated with $($report.Count) rows."

    # --- CSV OUTPUT (single file, append mode) ---
$csvOutputPath    = Join-Path -Path $OutputPath -ChildPath "AD_Computers_DailyStats.csv"
$csvCopyPath      = if (-not [string]::IsNullOrEmpty($localCopyPath)) { Join-Path -Path $localCopyPath -ChildPath "AD_Computers_DailyStats.csv" } else { '' }

    # Register CSV path for cleanup exclusion (shared convention)
    try {
        if (-not $global:csvGeneratedPaths) {
            $global:csvGeneratedPaths = @()
        }
        $global:csvGeneratedPaths += $csvOutputPath
    }
    catch {
        WriteLog -Message ("Failed to register CSV path in global csvGeneratedPaths: {0}" -f $_) "WARNING"
    }

    try {
        if (Test-Path -Path $csvOutputPath) {
            Write-Output "CSV file exists. Appending new data..."
            $report |
                ConvertTo-Csv -NoTypeInformation -Delimiter ";" |
                Select-Object -Skip 1 |
                Add-Content -Path $csvOutputPath -Encoding UTF8
        }
        else {
            Write-Output "CSV file does not exist. Creating new file with headers..."
            $report | Export-Csv -Path $csvOutputPath -NoTypeInformation -Delimiter ";" -Encoding UTF8
        }

        Write-Output "Successfully finished writing to CSV file: $csvOutputPath"

        try {
            if (-not [string]::IsNullOrEmpty($localCopyPath)) {
                Copy-Item -Path $csvOutputPath -Destination $csvCopyPath -Force
                Write-Output "Successfully copied CSV to: $csvCopyPath"
            }
            else {
                Write-Warning "Local configuration copy path (LatestCsvFolderPath) is null or empty. CSV copy will be skipped."
            }
        }
        catch {
            Write-Warning ("Failed to copy CSV to '{0}'. Error: {1}" -f $csvCopyPath, $_.Exception.Message)
        }

        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $csvOutputPath
    }
    catch {
        throw ("Failed to write the report to '{0}'. Please check file permissions. Error: {1}" -f $csvOutputPath, $_.Exception.Message)
    }


    # --- USERS DAILY STATS (per domain, UPN required) ---
    Write-Output "Step 4/4: Generating users stats..."

    if ($null -eq $allUsers) { $allUsers = @() }

    $inactiveThreshold = (Get-Date).AddDays(-90)

    $usersGroupedByDomain = $allUsers | Where-Object { -not [string]::IsNullOrEmpty($_.ADDomain) } | Group-Object -Property ADDomain
    $usersReport = $usersGroupedByDomain | ForEach-Object {
        $domainName = $_.Name
        $usersInGroup = $_.Group

        $enabledUsers = $usersInGroup | Where-Object { $_.Enabled -eq $true }
        $disabledUsers = $usersInGroup | Where-Object { $_.Enabled -eq $false }
        $inactiveUsers = $usersInGroup | Where-Object { $null -eq $_.LastLogonDate -or $_.LastLogonDate -lt $inactiveThreshold }

        [PSCustomObject]([ordered]@{
            Date = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            DomainName = $domainName
            TotalUsers = $usersInGroup.Count
            EnabledUsers = $enabledUsers.Count
            DisabledUsers = $disabledUsers.Count
            Inactive90DaysUsers = $inactiveUsers.Count
        })
    }

$usersCsvOutputPath = Join-Path -Path $OutputPath -ChildPath "AD_Users_DailyStats.csv"
$usersCsvCopyPath = if (-not [string]::IsNullOrEmpty($localCopyPath)) { Join-Path -Path $localCopyPath -ChildPath "AD_Users_DailyStats.csv" } else { '' }

    # Register CSV path for cleanup exclusion (shared convention)
    try { if (-not $global:csvGeneratedPaths) { $global:csvGeneratedPaths = @() }; $global:csvGeneratedPaths += $usersCsvOutputPath } catch { WriteLog -Message ("Failed to register Users CSV path in global csvGeneratedPaths: {0}" -f $_) "WARNING" }

    try {
        if (Test-Path -Path $usersCsvOutputPath) { Write-Output "Users CSV exists. Appending new data..."; $usersReport | ConvertTo-Csv -NoTypeInformation -Delimiter ";" | Select-Object -Skip 1 | Add-Content -Path $usersCsvOutputPath -Encoding UTF8 }
        else { Write-Output "Users CSV does not exist. Creating new file with headers..."; $usersReport | Export-Csv -Path $usersCsvOutputPath -NoTypeInformation -Delimiter ";" -Encoding UTF8 }

        Write-Output "Successfully finished writing Users CSV file: $usersCsvOutputPath"

        try { if (-not [string]::IsNullOrEmpty($localCopyPath)) { Copy-Item -Path $usersCsvOutputPath -Destination $usersCsvCopyPath -Force; Write-Output "Successfully copied Users CSV to: $usersCsvCopyPath" } else { Write-Warning "Local configuration copy path (LatestCsvFolderPath) is null or empty. Users CSV copy will be skipped." } }
        catch { Write-Warning ("Failed to copy Users CSV to '{0}'. Error: {1}" -f $usersCsvCopyPath, $_.Exception.Message) }

        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $usersCsvOutputPath
    }
    catch {
        throw ("Failed to write the users report to '{0}'. Please check file permissions. Error: {1}" -f $usersCsvOutputPath, $_.Exception.Message)
    }

    WriteLog -Message "AD computer inventory report successfully generated."
    $ScriptSucceeded = $true
}
catch {
    $globalError = $_
    WriteLog -Message ("Global error in AD computer inventory: {0}" -f $globalError) "ERROR"
    Write-Host "A global error occurred. Check the log file for details." -ForegroundColor Red

    # -------- Global error email notification (shared pattern) --------
    try {
        $title = "AD computers inventory - ERROR"
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

        SendEmailHtmlReport -BodyHtml $bodyHtml -Subject $title -Attachments $attachments -VerboseLog
    }
    catch {
        WriteLog -Message ("Failed to send error notification email: {0}" -f $_) "ERROR"
    }
}
finally {
    # Cleanup
    try {
        RemoveOldFiles -Path $OutputPath     -Filter "*.csv" -KeepCount $global:RetentionMaxCSV -LogFile $global:LogTextFile
        RemoveOldFiles -Path $global:LogPath -Filter "*.log" -KeepCount $global:RetentionMaxLogs -LogFile $global:LogTextFile
    }
    catch {
        WriteLog -Message ("Error during cleanup in finally: {0}" -f $_) "WARNING"
    }

    #region Mark Daily Success
    if ($ScriptSucceeded) {
        try {
            New-Item -Path $DailyLockFile -ItemType File -Force | Out-Null
            WriteLog -Message "[$DailyLockName] Success lock created for $Today." "INFO"
        }
        catch {
            WriteLog -Message ("[$DailyLockName] Failed to create daily success lock: {0}" -f $_.Exception.Message) "WARNING"
        }
    }
    #endregion


    WriteLog -Message "$TaskName completed (finally block)."

    try {
        Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
    }
    catch {
        # Ignore transcript stop errors
    }
}
