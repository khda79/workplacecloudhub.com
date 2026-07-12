<#
.SYNOPSIS
Generates a detailed inventory of Exchange Online mailbox migration jobs with history preservation.

.DESCRIPTION
This script connects to Exchange Online and retrieves enriched information about all active and historical
mailbox migration batches and their associated users. It merges results with ALL prior CSVs (from OutputPath only)
to preserve historical rows (jobs/users no longer present in EXO), then exports the merged dataset to CSV.
It includes retry logic for throttling, consistent ErrorAction handling, logging, transcript, and cleanup.

.PARAMETER OutputPath
Specifies the output directory where the CSV and log files will be stored.

.PARAMETER Connect
Forces a (reconnection) to cloud services when specified, or auto-connects if not connected.

.PARAMETER InteractiveAuth
Forces interactive authentication instead of app-only certificate authentication.
.VERSION
1.8



.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; ExchangeOnlineManagement.
    Minimum permissions: Exchange.ManageAsApp plus Exchange Online app-only RBAC allowing Get-MigrationBatch and related migration user/detail reads; Global Reader is the default read-only service-principal role.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Version : 1.6
    Author: https://github.com/khda79/workplacecloudhub.com
Date: October 2025
Dependencies: SmartM365.Core module
Required Modules: ExchangeOnlineManagement, Microsoft.Graph (handled by SmartM365.Core)
    Minimum permissions: Exchange Online app-only RBAC must allow Get-MigrationBatch and migration user read cmdlets.
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

# Ensure we are running on PowerShell 7+ (Core)
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    exit 1
}

# Avoid PS function-capacity issues
$MaximumFunctionCount = 32768

# App-only authentication parameters (same app as your other scripts)
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

# Import SmartM365.Core module (psd1). 
$modulePath = & { $d = $PSScriptRoot; while ($d) { $p = Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'; if (Test-Path -LiteralPath $p) { return $p }; $parent = Split-Path -Path $d -Parent; if ($parent -eq $d) { break }; $d = $parent }; throw 'SmartM365.Core module not found.' }
try {
    Import-Module -Name $modulePath -MinimumVersion '1.0.24' -ErrorAction Stop
} catch {
    Write-Host "Failed to import SmartM365.Core module from '$modulePath' : $_" -ForegroundColor Red
    exit 1
}

#region Helpers: Retry & Merge (script-specific)

function Invoke-WithRetry {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][scriptblock]$Script,
        [int]$MaxAttempts = 5,
        [int]$BaseDelaySeconds = 2
    )
    $attempt = 0
    $rnd = [System.Random]::new()
    while ($true) {
        try {
            return & $Script
        } catch {
            $attempt++
            $msg = $_.Exception.Message
            $isTransient =
                $msg -match 'throttl' -or
                $msg -match 'TooManyRequests' -or
                $msg -match 'temporar|transient|try again' -or
                $msg -match '5\d\d' -or
                $msg -match 'The remote server returned an error'
            if (-not $isTransient -or $attempt -ge $MaxAttempts) {
                throw
            }
            $delay = [int]([math]::Pow(2, $attempt - 1) * $BaseDelaySeconds + $rnd.Next(0,1000) / 1000.0)
            WriteLog -Message "Transient error detected (attempt $attempt/$MaxAttempts). Backing off for $delay s. Error: $msg" "WARN"
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-KeyFromRow {
    param([Parameter(Mandatory)]$Row)
    # Stable uniqueness: BatchGuid + MigrationUser; fallback BatchName + EmailAddress when Guid missing
    $guid = $Row.BatchGuid
    $user = $Row.MigrationUser
    if ([string]::IsNullOrWhiteSpace($guid)) {
        "{0}|{1}" -f $Row.BatchName, $Row.EmailAddress
    } else {
        "{0}|{1}" -f $guid, $user
    }
}

# Import CSV with delimiter auto-detection (',' vs ';') and header name trimming
function Import-CsvSmart {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path
    )
    if (-not (Test-Path -LiteralPath $Path)) { return @() }

    $firstLine = Get-Content -LiteralPath $Path -TotalCount 5 |
                 Where-Object { $_ -and ($_.Trim() -ne "") } |
                 Select-Object -First 1

    $delimiter = ','
    if ($firstLine) {
        $commaCount = ($firstLine.ToCharArray() | Where-Object { $_ -eq ',' }).Count
        $semiCount  = ($firstLine.ToCharArray() | Where-Object { $_ -eq ';' }).Count
        if ($semiCount -gt $commaCount) { $delimiter = ';' }
    }

    $rows = Import-Csv -LiteralPath $Path -Delimiter $delimiter -Encoding UTF8 -ErrorAction SilentlyContinue

    foreach ($r in $rows) {
        foreach ($p in @($r.PSObject.Properties)) {
            $old = $p.Name
            $new = $old.Trim()
            if ($new -ne $old) {
                $val = $r.$old
                Add-Member -InputObject $r -NotePropertyName $new -NotePropertyValue $val -Force
                $r.PSObject.Properties.Remove($old) | Out-Null
            }
        }
    }
    return $rows
}

# Enumerate ALL CSVs in OutputPath only (no recursion)
function Get-AllPreviousCsvPaths {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][string]$BaseFileName
    )
    $candidates = New-Object System.Collections.Generic.List[string]

    if (Test-Path -LiteralPath $Folder) {
        $candidates += Get-ChildItem -LiteralPath $Folder -Filter "$BaseFileName*.csv" -File -ErrorAction SilentlyContinue | Select-Object -ExpandProperty FullName
    }

    # Sort oldest -> newest so the latest overwrites previous entries
    $sorted = $candidates | ForEach-Object {
        $fi = Get-Item -LiteralPath $_
        [PSCustomObject]@{ Path = $_; LastWriteTime = $fi.LastWriteTime }
    } | Sort-Object LastWriteTime, Path

    return $sorted.Path
}

#endregion

#region Module Import and Initialization
$ScriptVersion = "1.8"
$TaskName      = "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) v$ScriptVersion ..."
$OutputPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ExoMigrationJobCsvLogFolderPath' -DefaultValue $OutputPath
try {
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPathInit $OutputPath -LogFileName $(($MyInvocation.MyCommand.Name) -replace '\.ps1$','')
    Start-Transcript -Path $global:logTranscriptFile -Append

    WriteLog -Message "Script Environment initialized at $InitializeOutputPath"
    $OutputPath = $InitializeOutputPath
    WriteLog -Message "Starting $TaskName..."
} catch {
    Write-Host "Initialization failed: $_" -ForegroundColor Red
    exit
}

#endregion

#region Error email helper (using SmartM365.Core mail functions)

function Send-InventoryErrorEmail {
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)][string]$ErrorMessage
    )

    try {
        $htmlBody = NewSimpleEmailBody -Title $Title -Message $ErrorMessage

        # SMTP / From / To / Cc / Port are controlled globally by SmartM365.Core defaults
        SendEmailHtmlReport -Subject $Title -BodyHtml $htmlBody -VerboseLog
    } catch {
        WriteLog -Message ("Failed to send error notification email: {0}" -f $_.Exception.Message) "ERROR"
    }
}

#endregion

# Top-level protection to ensure cleanup is always attempted
try {

    #region Connect (Exchange Online + Graph via SmartM365.Core / Connect-SmartM365CloudSession)

    $connectedExoInThisRun   = $false
    $connectedGraphInThisRun = $false

    # Detect existing Exchange Online session if cmdlet is available
    $exoInfo = $null
    if (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue) {
        try {
            $exoInfo = Get-ConnectionInformation -ErrorAction SilentlyContinue
        } catch { }
    }

    # Detect existing Microsoft Graph session
    $graphContext = $null
    if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
        try {
            $graphContext = Get-MgContext -ErrorAction SilentlyContinue
        } catch { }
    }

    # Logic:
    # - If -Connect is specified: always disconnect EXO + Graph then reconnect.
    # - If -Connect is not specified:
    #       * If EXO or Graph already connected: reuse existing sessions.
    #       * If neither EXO nor Graph connected: connect once.
    $needConnect = $false

    if ($Connect) {
        Write-Host "Connect switch specified: existing cloud sessions will be disconnected and reconnected..." -ForegroundColor Cyan
        Disconnect-SmartM365CloudSession -ExchangeOnline $true -Graph $true -VerboseDisconnect:$true
        $needConnect = $true
    } else {
        if ($exoInfo -or $graphContext) {
            if ($exoInfo) {
                Write-Host "Existing Exchange Online session detected. Reusing current connection." -ForegroundColor Cyan
            }
            if ($graphContext) {
                Write-Host "Existing Microsoft Graph session detected. Reusing current connection." -ForegroundColor Cyan
            }
            $needConnect = $false
        } else {
            Write-Host "No existing EXO/Graph sessions detected. Will establish new connections..." -ForegroundColor Cyan
            $needConnect = $true
        }
    }

    if ($needConnect) {
        $connectParams = @{
            ExchangeOnline = $true
            Graph          = $true   # Keep Graph enabled so the same logic applies to both services
        }

        if (-not $InteractiveAuth) {
            # Default: app-only certificate authentication
            $connectParams.AppId        = $AppId
            $connectParams.Thumbprint   = $Thumb
            $connectParams.TenantId     = $TenantId
            $connectParams.Organization = $OrgDomain
            WriteLog -Message "Connecting to EXO and Graph with app-only certificate authentication." "INFO"
        } else {
            # Interactive authentication
            WriteLog -Message "Connecting to EXO and Graph with interactive authentication." "INFO"
        }

        $connectResult = Connect-SmartM365CloudSession @connectParams

        if (-not $connectResult.ExchangeOnlineConnected) {
            Write-Host "Failed to connect to Exchange Online. Aborting." -ForegroundColor Red
            WriteLog -Message "Failed to connect to Exchange Online. Script aborted." "ERROR"

            # Even ici, on peut prevenir par mail
            Send-InventoryErrorEmail -Title "SmartM365 EXO migration inventory - Connection error" -ErrorMessage "Failed to connect to Exchange Online. Script aborted."
            exit
        }

        $connectedExoInThisRun   = $connectResult.ExchangeOnlineConnected
        $connectedGraphInThisRun = $connectResult.GraphConnected
    }

    #endregion

    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputPath) -ExchangeOnlineProbeCommands @('Get-MigrationBatch') | Out-Null

    #region Main Script Logic

    WriteLog -Message "Starting enriched Exchange Online migration jobs inventory"

    try {
        $results       = @()
        $migrationJobs = Invoke-WithRetry { Get-MigrationBatch -ErrorAction Stop }
        $maxItemsReached = $false
        if ($MaxItems -gt 0) {
            WriteLog -Message ("MaxItems enabled: migration user export will stop after {0} rows." -f $MaxItems) "WARN"
        }

        foreach ($batch in $migrationJobs) {
            $batchDetails = $batch
            try {
                $users = Invoke-WithRetry { Get-MigrationUser -BatchId $batch.Identity -ErrorAction Stop }
            } catch {
                $users = @()
                WriteLog -Message "Error in Get-MigrationUser for batch '$($batch.Name)': $_" "ERROR"
            }

            foreach ($user in $users) {
                if ($MaxItems -gt 0 -and $results.Count -ge $MaxItems) { $maxItemsReached = $true; break }
                $results += [PSCustomObject]@{
                    Timestamp                   = Get-Date
                    BatchName                   = $batchDetails.Identity
                    BatchStatus                 = $batchDetails.Status
                    State                       = $batchDetails.State
                    DataConsistencyScore        = $batchDetails.DataConsistencyScore
                    Flags                       = $batchDetails.Flags
                    WorkflowStage               = $batchDetails.WorkflowStage
                    TriggeredAction             = $batchDetails.TriggeredAction
                    BatchGuid                   = $batchDetails.BatchGuid
                    TotalCount                  = $batchDetails.TotalCount
                    ActiveCount                 = $batchDetails.ActiveCount
                    StoppedCount                = $batchDetails.StoppedCount
                    SyncedCount                 = $batchDetails.SyncedCount
                    FinalizedCount              = $batchDetails.FinalizedCount
                    FailedCount                 = $batchDetails.FailedCount
                    CompletedWithWarningCount   = $batchDetails.CompletedWithWarningCount
                    PostMigrationSyncCount      = $batchDetails.PostMigrationSyncCount
                    PendingCount                = $batchDetails.PendingCount
                    ProvisionedCount            = $batchDetails.ProvisionedCount
                    ValidationWarningCount      = $batchDetails.ValidationWarningCount
                    ValidationWarnings          = ($batchDetails.ValidationWarnings -join "; ")
                    Message                     = $batchDetails.Message
                    CreationDateTime            = $batchDetails.CreationDateTime
                    CreationDateTimeUTC         = $batchDetails.CreationDateTimeUTC
                    StartDateTime               = $batchDetails.StartDateTime
                    StartDateTimeUTC            = $batchDetails.StartDateTimeUTC
                    ExpirationDateTime          = $batchDetails.ExpirationDateTime
                    ExpirationDateTimeUTC       = $batchDetails.ExpirationDateTimeUTC
                    LastSyncedDateTime          = $batchDetails.LastSyncedDateTime
                    LastSyncedDateTimeUTC       = $batchDetails.LastSyncedDateTimeUTC
                    SubmittedByUser             = $batchDetails.SubmittedByUser
                    OwnerId                     = $batchDetails.OwnerId
                    OwnerExchangeObjectId       = $batchDetails.OwnerExchangeObjectId
                    NotificationEmails          = ($batchDetails.NotificationEmails -join "; ")
                    MigrationType               = $batchDetails.MigrationType
                    BatchDirection              = $batchDetails.BatchDirection
                    Locale                      = $batchDetails.Locale
                    Reports                     = ($batchDetails.Reports -join "; ")
                    IsProvisioning              = $batchDetails.IsProvisioning
                    IsLargeArchiveOnboarding    = $batchDetails.IsLargeArchiveOnboarding
                    IsOnboardingRestore         = $batchDetails.IsOnboardingRestore
                    CompleteAfter               = $batchDetails.CompleteAfter
                    CompleteAfterUTC            = $batchDetails.CompleteAfterUTC
                    MigrationUser               = $user.Identity
                    UserStatus                  = $user.Status
                    EmailAddress                = $user.EmailAddress
                    EstimatedTotalItems         = $user.EstimatedTotalItems
                    EstimatedTotalSize          = $user.EstimatedTotalSize
                    FinalSyncTime               = $user.FinalSyncTime
                    LastSyncTime                = $user.LastSyncTime
                    CompleteAfterUser           = $user.CompleteAfter
                    CompleteAfterUTCUser        = $user.CompleteAfterUTC
                }
                if ($MaxItems -gt 0 -and $results.Count -ge $MaxItems) { $maxItemsReached = $true; break }
            }
            if ($maxItemsReached) { WriteLog -Message ("MaxItems reached: stopped after {0} exported migration user rows." -f $results.Count) "WARN"; break }
        }

        Write-Host "`n--- Export CSV ---"

        $baseName        = "Exchange_EXO_MigrationJobs"
        $allPrevCsvPaths = @()
        if ($MaxItems -gt 0) {
            WriteLog -Message "MaxItems enabled: skipping previous CSV merge so the test export only contains current-run rows." "WARN"
        }
        else {
            $allPrevCsvPaths = Get-AllPreviousCsvPaths -Folder $OutputPath -BaseFileName $baseName
        }

        if ($allPrevCsvPaths.Count -gt 0) {
            # Compact log: only count, no huge list
            WriteLog -Message ("Previous CSVs found in OutputPath: {0} file(s)." -f $allPrevCsvPaths.Count)
        } else {
            WriteLog -Message "No previous CSVs found in OutputPath for base '$baseName'." "WARN"
        }

        # Determine column order (strict)
        $propertyOrder = @()
        if ($results.Count -gt 0) {
            $propertyOrder = $results[0].psobject.Properties.Name
        } else {
            $latestPrev = $allPrevCsvPaths | Select-Object -Last 1
            if ($latestPrev) {
                $firstPrev = Import-CsvSmart -Path $latestPrev | Select-Object -First 1
                if ($firstPrev) { $propertyOrder = $firstPrev.psobject.Properties.Name }
            }
        }

        # Build merged dictionary
        $dict = @{}

        foreach ($csvPath in $allPrevCsvPaths) {
            $prevRows = Import-CsvSmart -Path $csvPath
            foreach ($r in $prevRows) {
                $k = Get-KeyFromRow -Row $r
                $dict[$k] = $r
            }
        }

        foreach ($r in $results) {
            $k = Get-KeyFromRow -Row $r
            $dict[$k] = $r
        }

        $merged = @()
        if ($dict.Count -gt 0) { $merged = $dict.Values }

        if (($merged.Count -eq 0) -and ($propertyOrder.Count -gt 0)) {
            $empty = [ordered]@{}
            foreach ($h in $propertyOrder) { $empty[$h] = $null }
            $merged = ,([pscustomobject]$empty) | Where-Object { $false }
        }

        if ($propertyOrder.Count -gt 0 -and $merged.Count -gt 0) {
            $merged = $merged | Select-Object -Property $propertyOrder
        }

        ExportAndCopyCsv -BaseFileName $baseName `
            -OutputPath $OutputPath `
            -GlobalPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') `
            -Data $merged `
            -Encoding "UTF8" `
            -NoTypeInformation

    } catch {
        $msg = "Global error while retrieving migration jobs: $($_.Exception.Message)"
        WriteLog -Message $msg "ERROR"
        Send-InventoryErrorEmail -Title "SmartM365 EXO migration inventory - Global error" -ErrorMessage $msg
    }

    WriteLog -Message ("Script completed. {0} migration jobs processed. {1} users exported." -f `
        ( ($migrationJobs | Measure-Object).Count ), ( ($results | Measure-Object).Count ) )

    #endregion

}
catch {
    # Top-level catch for any unhandled errors in the main workflow
    $fatal = "Unhandled fatal error in script: $($_.Exception.Message)"
    WriteLog -Message $fatal "ERROR"
    Send-InventoryErrorEmail -Title "SmartM365 EXO migration inventory - Fatal error" -ErrorMessage $fatal
}
finally {

    #region Disconnect + Cleanup

    # Only disconnect services that were connected by this script (not reused sessions)
    if ($connectedExoInThisRun -or $connectedGraphInThisRun) {
        Write-Host "`n--- Disconnect Cloud Services ---"
        Disconnect-SmartM365CloudSession -ExchangeOnline:$connectedExoInThisRun -Graph:$connectedGraphInThisRun -VerboseDisconnect:$true
    }

    # Cleanup of old files (CSV + logs)
    RemoveOldFiles -Path $OutputPath     -Filter "*.csv" -KeepCount $global:RetentionMaxCSV -LogFile $global:LogTextFile
    RemoveOldFiles -Path $global:LogPath -Filter "*.log" -KeepCount $global:RetentionMaxLogs -LogFile $global:LogTextFile

    WriteLog -Message "$TaskName completed."
    Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {}
    Complete-SmartM365ExecutionContext -Status Auto
    #endregion

}
# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD880ZkA8Z6GbP2
# X+D8gAogzrPHJGpweHq899i9zZeF/aCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
# v0GFVsTsys9PMA0GCSqGSIb3DQEBCwUAMCAxHjAcBgNVBAMMFXdvcmtwbGFjZWNs
# b3VkaHViLmNvbTAeFw0yNjA3MTIwNjM5MTZaFw0yOTA3MTIwNjQ5MTZaMCAxHjAc
# BgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNvbTCCAaIwDQYJKoZIhvcNAQEBBQAD
# ggGPADCCAYoCggGBAMJqEmY4V9VM4HhTovXPXHSWb44jVYMj05xJIZf2f/NxQLR/
# vfka/0JbdTSRJ03Yy3OIulBP5DqbnfAyzv+9eulPVX/BUFM6b2lENxZpVrvj55TZ
# levsXyzHuK0xs7/FFpbLQ2Ts3LGPJTLlneOfuEWKRT6xTotD1RnElDCumiOnQHOD
# 6qtPSRuwoxaVwSDw2QFJ8hp4RGHKsDAMRLgaRBhBM7e9A3/k7bA541DrWt19Cq5d
# IY1LUII3pVolF3YUtot7wFU2BbfpM0WiDEPXDWBUAvHNF0FDDukwuXUtn9J2n1f/
# 8EzDznON1GuNhrPP7cWJh6hywJgBzeR7ZHf2tsk76sKqY75u+qWoe4xQJXK7V2N7
# UJW7i6YC2W+/LrOaUYB9JykD88Jk+OJ2eLDtLSqzYAnJXYTIq7/mju5E8twyNZrN
# tQHqKUxUKhkeVgezgKoc4t12dgkTryl9efMy3qyxNesN34RR2i6eK8+6UtiW2ae5
# GESynl96l1E9+UWlRQIDAQABo0YwRDAOBgNVHQ8BAf8EBAMCB4AwEwYDVR0lBAww
# CgYIKwYBBQUHAwMwHQYDVR0OBBYEFEooM+aK7XCOIsSi0oFRhXyVQqdzMA0GCSqG
# SIb3DQEBCwUAA4IBgQC08zIpMh0vUuvfMcIUpwX3lABvT3V9Rf6swy8xuWHjJyJz
# hZVt0hOHeCBWF2RxYeJ2iY4hyH4FSkwwLCHmmM6kV3eLY2uibsYCUdwm1mwbtSws
# i4YAzGZF0Ueap2TC94d9O/dcpzYILKPdJwqAd3MprkWEbyFSfEkhy5NCmxZ2wQFd
# LtOU6YHMI9v6P8tIhGXpZbp3QjK9mZif6LZ9ZgXEzi4whxDwQ2RMTUVaf7kamyjc
# gGmO32gRcNr0qsGwTog7TUTcbTd/RVc0DEUMMrUZVWMcBwrBIFUWqnD4i/oZuHdH
# pMytQjZQcZBOzrJ/YcWxMNmdf09gq44kFs1QHiG+FFnATyglOs8SR3fJwJdPI+KN
# qpK0zo9FhCyl37qSpKpyS9QNZdl+isj7YQncfqCmadjY1y6nZhLzaEoDW0oHdv/s
# NzjZ54ieDALCH69wCbeCYk1lrI3ggu0t22QG1sHN7NmOm3T6SL2w7cF+TpeYXIfv
# FCGIHWHVGbQtK/TtwJMxggJmMIICYgIBATA0MCAxHjAcBgNVBAMMFXdvcmtwbGFj
# ZWNsb3VkaHViLmNvbQIQcCHy1SICVr9BhVbE7MrPTzANBglghkgBZQMEAgEFAKCB
# hDAYBgorBgEEAYI3AgEMMQowCKACgAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEE
# AYI3AgEEMBwGCisGAQQBgjcCAQsxDjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJ
# BDEiBCCuRE69UPfGWeiC9s4mrhYujHA2zxmLQxpRsbPXq3gX+zANBgkqhkiG9w0B
# AQEFAASCAYB6N+gOHL97Ue76GYstf6q+uOv0xChvSHTdRqTjIbWuGEKs0iEx0Com
# IxGvBddVzap6wgLLBgpXw4kziE6MhbVkOWcmiuGm+ziNn3elLzP7ynFc0l789EY6
# Dfa1KBEHTgZYA9qd2uekAQ8MPGNBw0dwVhRNFSbKRZP6zHHuH17fNGuUpUPB8K95
# U0KIkF1ifL+AxOlRrUj7Qu9rvGfLnRiBU/1JTGURDwb+HO8N4yyyIya7aqs4Rq3a
# mo7BUA+1neDX6hSpq9dSlC6++0qrHyAbDePKikXDTghnC1XftLqHwZR71F1MXqPY
# zcu4b3ho0yn65Ns9d7rkbXjQCu1vDV4PQod/zDL5r8PnvJyKUnT16mDoWj6c1qGy
# 9OHGjNTMZxApGqHj4/WPi0GnXi8KGczLmWFzT6Pbpj8xFX4ae103NGTtQwyEMz6p
# e1ZnRklE5sqiFERwZ/AGektvhFXVMU+uoEjFSTFo4DSSZ7EfgBCtuR87X6Rzc1N4
# 0tcW9uOolQA=
# SIG # End signature block
