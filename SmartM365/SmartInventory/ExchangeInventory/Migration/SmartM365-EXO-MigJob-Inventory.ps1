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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD880ZkA8Z6GbP2
# X+D8gAogzrPHJGpweHq899i9zZeF/aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIK5ETr1Q98ZZ6IL2ziauFi6McDbPGYtDGlGxs9ereBf7MA0GCSqG
# SIb3DQEBAQUABIIBgJJPGo4sG9TPVy15Bo39TItI/a1Os4A5hBaNlrWmdusfAuL1
# /5u3RJeYza7nhuUU5OFHu/v6ndmHBRUgi9O4SPAyejgLZclnMAYqimPTca9fsOqT
# bFnlLoc/x5/wdyLO9GCJ3wygoJkk40gl4o3TxW0kAnFC9liV8Fuqa+fPCQZct8to
# 6jturzW/KY5Ymx9bk5dzCuGLgxXY7IPh0I7tKVa6Ci/FSt+Z5YNEuDWPxgu0AECS
# v2Nb1ZU++hg+0Bw/bb5pvsBH5uP+3FEAftW0j13PsOMbe/uPTnP0vVjLql1y8Su+
# 2oZR0GtJ6VNv8JYbLoAot9YUB24pQx4YztlXmH5RnhboQp0wMZWNR2Kutw67v22S
# 2h3dYCHDmhS6s3c8n281kj4lOVPDzijtSdEdnHbQWNiKbjCNk02OFx8qTAyjR+hU
# Cm3cbrPaI49H37ORUiYtcLQ7fxt2rqQP00ClxDA57uFBChvISOsJwV4kXruguEyg
# B5ndC4+jP3T/CelIlKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODUw
# MjZaMC8GCSqGSIb3DQEJBDEiBCBlPKvwi4r1raCTzvM3hi3KtT1/e7Xya74qMMVT
# JVosyTANBgkqhkiG9w0BAQEFAASCAgCc3xdeBTVyztJnggyLEs0CJBUAgjZj7UqU
# VJiDTCBvpr6zFs2PcXIMFwa+Bz8kgTAb+oSNn9xdtrtAWLPYOOZO5uuTNrQNEb4O
# MIHsV+U8oK7KGIPOgu6w29mp/KdsCv1MOH/R13PpQmJ8D6AdGFKdOjgqcijFGBI6
# w4K+TkZNEIMwiH0ul7N9lmAPFl28mgMqD/tsvLvAeEYS/rDfsv/+9jZ8eKo3xaa1
# uAeZ2O0sql5A28DroZk4IZTYXKanBJjawzZuyXI3rJDRhlQlV6knQG5foF+QS8aP
# fBmiNBons9qR9SU8z51hqtqYdLu3GMiX8FmxynLaxIPfNNdIa16J+w+LmsgPy0XL
# DvwauVWk6v5FFl12IdELoYeFVUqrCplIYtgb7zxFAfSJO9bsKus1LemBAvPSKsvQ
# gLs6BcW7UsTa95s9wRRfZV6K2lgmVgcdMkMbFiciuA13iWdA8YEpYPWiwVnTT7lA
# e/Fn5uz+O0ecejbgH+m3JXju+og1aYBepyou4iXs7k/HHiS8Hf0Rx7viYETc0L6j
# RGJvTVphlOIc7JzYwiJmETcmpDPGYS5TbrecjgY8IInMNkF8yy6IkV8+ez0EhKuq
# 6zfYTy1+9nLn+LKyrvhWw/X1yqWRdp0nK6fSMhUSRJBqbOIIDkEhmHzsiSYMDU5r
# jL9IHj5U7Q==
# SIG # End signature block
