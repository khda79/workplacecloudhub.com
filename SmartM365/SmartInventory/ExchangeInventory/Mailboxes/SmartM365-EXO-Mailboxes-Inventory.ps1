<#
.SYNOPSIS
    Exchange Online mailbox inventory with Graph enrichment.
    Produces TWO CSVs every run:
      - Exchange_EXO_Mailboxes_AllDomains.csv           (detailed default export, NO archive metrics; columns EXACT as original)
      - Exchange_EXO_Mailboxes_AllDomains_Archive.csv   (minimal archive metrics export)
    Optional third CSV if -IncludePerm is used:
      - Exchange_EXO_Mailboxes_AllDomains_Permissions.csv

.DESCRIPTION
    Optimizations:
      - Single pass over mailboxes
      - Primary stats via EXO (optional extra call only if -IncludeLastUserActionTime) to fetch LastUserActionTime; legacy fallbacks used for LastLogonTime/ItemCount/Size
      - Size computed via TotalItemSize.Value.ToBytes() (priority), then fallbacks (incl. FR parsing)
      - Permissions live only with -IncludePerm; otherwise via snapshot CSV by UPN (NoPermInDefault)
      - Robust DisplayName (EXO -> Get-User -> Graph -> Get-Recipient -> UPN)
      - Robust RecipientTypeDetails/MailboxType via fallback Get-Recipient/Get-EXOMailbox/Get-Mailbox if null
      - -MaxItems <N> for bounded smoke tests (preferred); -Top100 is kept for legacy quick tests
      - -ResolveManager enables Graph manager resolution (DISABLED by default; column kept, may be empty)
      - SchedulingMailbox added to system types filter
      - AssignedLicenses no longer fetched (column kept empty)
      - Graph HTTP code logging and NOC-style error summary
      - -PermissionsOnly (with -IncludePerm) accelerates by collecting permissions only; other columns are left empty (NOT "N/A")
      - LastUserActionTime: populated from Stats snapshot when available. In live stats mode (-IncludeStats),
        requires ALSO passing -IncludeLastUserActionTime (triggers an extra Get-EXOMailboxStatistics call per mailbox,
        expensive at scale ~9800 mailboxes). Without -IncludeLastUserActionTime, the column is intentionally empty
        even when -IncludeStats is active.
.VERSION
1.15


.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; ExchangeOnlineManagement; Microsoft.Graph.Authentication when Graph enrichment is used. SmartM365 mail/SharePoint upload uses Graph REST direct from SmartM365.Core.
    Minimum permissions: Exchange.ManageAsApp plus Exchange Online app-only RBAC allowing Get-Mailbox and optional mailbox statistics/permission cmdlets; Global Reader is the default read-only service-principal role.
    Minimum Graph application permissions: User.Read.All for user enrichment. Not required in -PermissionsOnly mode.
    Conditional: Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Minimum application permissions: User.Read.All; Exchange Online RBAC must allow Get-Mailbox and related read cmdlets.
#>

param(
    [string]$Tenant = 'test',
[switch]$IncludePerm,
    [switch]$Top100,
    [ValidateSet("Legacy","Auto","Off","Force")] [string]$RegionalConfigMode = "Off",
    [switch]$Connect,
    [string]$OutputPath,
    [string]$DebugUPN,
    # Enable Graph manager resolution (disabled by default)
    [switch]$ResolveManager = $false,
    # Resolve RecipientTypeDetails using live EXO cmdlets when missing (disabled by default; expensive)
    [switch]$ResolveRecipientTypeDetailsLive = $false,
    # Include LastUserActionTime via an extra EXO call (disabled by default)
    [switch]$IncludeLastUserActionTime = $false,
    # Include live mailbox statistics collection via Get-EXOMailboxStatistics (default: use Stats snapshot)
    [switch]$IncludeStats,
    # Exclude live mailbox statistics collection (use Stats snapshot CSV instead) [DEPRECATED: kept for backward compatibility]
    [switch]$ExcludeStats,
    # Use Graph interactive auth instead of app-only certificate auth
    [switch]$InteractiveAuth,
    # Collect permissions only (still exports same CSV columns/order; non-permission fields are empty)
    [switch]$PermissionsOnly = $false,
    # Maximum run duration in minutes before a clean stop with partial export (0 = no limit)
    [int]$MaxRunMinutes = 2880,
    # Number of parallel threads for permission collection in -PermissionsOnly mode (1 = sequential)
    [ValidateRange(1,20)][int]$ParallelThrottle = 10,
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
# Centralized paths (snapshots + export)
# ==========================================================
$LatestCsvFolderPath          = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ""
$AdUsersCsvPath        = Join-Path $LatestCsvFolderPath "AD_Users_AllDomains.csv"
$M365UsersCsvPath      = Join-Path $LatestCsvFolderPath "M365_Users_Active.csv"
$PermSnapshotDir       = $LatestCsvFolderPath
$StatsSnapshotCsvPath  = Join-Path $LatestCsvFolderPath "Exchange_EXO_Mailboxes_AllDomains_Stats.csv"

# ==========================================================
# Import SmartM365.Core module (psd1)
# ==========================================================
$modulePath = & { $d = $PSScriptRoot; while ($d) { $p = Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'; if (Test-Path -LiteralPath $p) { return $p }; $parent = Split-Path -Path $d -Parent; if ($parent -eq $d) { break }; $d = $parent }; throw 'SmartM365.Core module not found.' }
try {
    Import-Module -Name $modulePath -MinimumVersion '1.0.24' -ErrorAction Stop
} catch {
    Write-Host "Failed to import SmartM365.Core module from '$modulePath' : $_" -ForegroundColor Red
    exit 1
}


# ==========================================================
# Graph mail configuration
# Mail must be sent only to configured recipients.
# ==========================================================
$From              = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'From' -DefaultValue ""
$ErrorMailTo   = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ErrorMailTo' -DefaultValue ""
$EnableErrorEmail  = $true
$EnableSummaryEmail = $false

function Send-FatalErrorEmail {
    param(
        [Parameter(Mandatory=$true)][string]$Subject,
        [Parameter(Mandatory=$true)][string]$HtmlBody,
        [Parameter()][string[]]$Attachments = $null
    )

    if (-not $EnableErrorEmail) { return }

    if ([string]::IsNullOrWhiteSpace($ErrorMailTo)) {
        WriteLog -Message "Error email enabled but ErrorMailTo is empty. Skipping." "WARNING"
        return
    }

    try {
        Send-SmartM365Mail -From $From -To $ErrorMailTo -Subject $Subject -BodyHtml $HtmlBody -Attachments $Attachments
        WriteLog -Message ("Fatal error HTML email sent to {0} via Microsoft Graph." -f $ErrorMailTo) "INFO"
    }
    catch {
        WriteLog -Message ("Failed to send fatal error HTML email: {0}" -f $_.Exception.Message) "WARNING"
    }
}

function ConvertTo-MailboxInventoryHtmlText {
    param([AllowNull()]$Value)
    if ($null -eq $Value) { return '' }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Add-MailboxInventoryHtmlSection {
    param(
        [AllowNull()][string]$BodyHtml,
        [AllowNull()][string]$SectionHtml
    )

    if ([string]::IsNullOrWhiteSpace($SectionHtml)) { return $BodyHtml }
    if ([string]::IsNullOrWhiteSpace($BodyHtml)) { return $SectionHtml }
    if ($BodyHtml -match '(?is)</body>') {
        return ([regex]::Replace($BodyHtml, '(?is)</body>', ($SectionHtml + "`n</body>"), 1))
    }
    return ([string]$BodyHtml + "`n" + $SectionHtml)
}

function New-StatsCompletenessIssuesHtml {
    param(
        [Parameter(Mandatory)][object[]]$Issues,
        [int]$MaxRows = 50
    )

    $rows = New-Object System.Collections.Generic.List[string]
    foreach ($issue in ($Issues | Select-Object -First $MaxRows)) {
        $rows.Add(("<tr><td style=`"border-top:1px solid #e2e8f0;padding:6px 8px;font-size:11px;color:#334155;`">{0}</td><td style=`"border-top:1px solid #e2e8f0;padding:6px 8px;font-size:11px;color:#334155;word-break:break-all;`">{1}</td><td style=`"border-top:1px solid #e2e8f0;padding:6px 8px;font-size:11px;color:#334155;word-break:break-all;`">{2}</td><td style=`"border-top:1px solid #e2e8f0;padding:6px 8px;font-size:11px;color:#b45309;font-weight:700;`">{3}</td><td style=`"border-top:1px solid #e2e8f0;padding:6px 8px;font-size:11px;color:#334155;`">{4}</td><td style=`"border-top:1px solid #e2e8f0;padding:6px 8px;font-size:11px;color:#334155;`">{5}</td></tr>" -f `
            (ConvertTo-MailboxInventoryHtmlText $issue.RowNumber),
            (ConvertTo-MailboxInventoryHtmlText $issue.UserPrincipalName),
            (ConvertTo-MailboxInventoryHtmlText $issue.PrimarySmtpAddress),
            (ConvertTo-MailboxInventoryHtmlText $issue.MissingFields),
            (ConvertTo-MailboxInventoryHtmlText $issue.StatsSource),
            (ConvertTo-MailboxInventoryHtmlText $issue.Details))) | Out-Null
    }

    if ($rows.Count -eq 0) { return '' }
    $extra = ''
    if ($Issues.Count -gt $MaxRows) {
        $extra = "<p style=`"margin:8px 0 0 0;font-size:11px;color:#64748b;`">Only the first $MaxRows issue rows are shown in this mail.</p>"
    }

    return @"
<div style=""margin-top:18px;"">
  <h3 style=""margin:0 0 8px 0;font-size:14px;line-height:20px;color:#0f172a;"">Live mailbox statistics completeness issues</h3>
  <p style=""margin:0 0 10px 0;font-size:12px;line-height:18px;color:#334155;"">The run was stopped before DATA-LAST publication because live Exchange Online statistics returned incomplete critical fields.</p>
  <table role=""presentation"" cellpadding=""0"" cellspacing=""0"" style=""width:100%;border-collapse:collapse;border:1px solid #e2e8f0;"">
    <tr><th align=""left"" style=""background:#f8fafc;padding:6px 8px;font-size:10px;color:#475569;text-transform:uppercase;"">Row</th><th align=""left"" style=""background:#f8fafc;padding:6px 8px;font-size:10px;color:#475569;text-transform:uppercase;"">UPN</th><th align=""left"" style=""background:#f8fafc;padding:6px 8px;font-size:10px;color:#475569;text-transform:uppercase;"">Primary SMTP</th><th align=""left"" style=""background:#f8fafc;padding:6px 8px;font-size:10px;color:#475569;text-transform:uppercase;"">Missing fields</th><th align=""left"" style=""background:#f8fafc;padding:6px 8px;font-size:10px;color:#475569;text-transform:uppercase;"">Stats source</th><th align=""left"" style=""background:#f8fafc;padding:6px 8px;font-size:10px;color:#475569;text-transform:uppercase;"">Details</th></tr>
    $($rows -join "`n")
  </table>
  $extra
</div>
"@
}

function Publish-MailboxInventoryDiagnosticCsv {
    param([Parameter(Mandatory)][string]$Path)

    try {
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return }
        $uploadCommand = Get-Command -Name Invoke-SmartM365SharePointCsvUpload -ErrorAction SilentlyContinue
        if (-not $uploadCommand) { return }
        Invoke-SmartM365SharePointCsvUpload -LocalFilePath $Path | Out-Null
        WriteLog -Message ("Diagnostic CSV uploaded to SharePoint: {0}" -f $Path) "SUCCESS"
    }
    catch {
        WriteLog -Message ("Diagnostic CSV SharePoint upload failed; local diagnostic remains available: {0}. Error: {1}" -f $Path, $_.Exception.Message) "WARNING"
    }
}
#region Init
$ScriptVersion = "1.15"
$script:StatsCompletenessDiagnosticPath = $null
$script:StatsCompletenessIssueRows = @()
$StatsCompletenessFailMinRows = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'StatsCompletenessFailMinRows' -DefaultValue 50)
$StatsCompletenessFailPercent = [double](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'StatsCompletenessFailPercent' -DefaultValue 0.5)
$IsMaxItemsRun = ($MaxItems -gt 0)
$IsBoundedMailboxRun = ($IsMaxItemsRun -or $Top100)
$CsvSuffix = if ($IsMaxItemsRun) { "_MAXITEMS-$MaxItems" } elseif ($Top100) { "_top100" } else { "" }
if ($PermissionsOnly -and (-not $IncludePerm)) {
    # PermissionsOnly is meaningful only when permissions are being collected live
    $IncludePerm = $true
}

$TaskNameCore = [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
if ($IncludePerm) { $TaskNameCore = "$TaskNameCore (DEFAULT + ARCHIVE + PERMISSIONS)" }
if ($PermissionsOnly) { $TaskNameCore = "$TaskNameCore (PERMISSIONS ONLY)" }

$TaskName = if ($IsMaxItemsRun) { "$TaskNameCore [MAXITEMS-$MaxItems TEST] v$ScriptVersion ..." } elseif ($Top100) { "$TaskNameCore [TOP100 MODE] v$ScriptVersion ..." } else { "$TaskNameCore v$ScriptVersion ..." }

$OutputPath = if ($PermissionsOnly) { Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ExoMailboxPermissionsCsvLogFolderPath' -DefaultValue $OutputPath } else { Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ExoMailboxCsvLogFolderPath' -DefaultValue $OutputPath }

try {
    $logMode = if ($PermissionsOnly) { 'PermissionsOnly' } elseif ($IncludeStats -and (-not $ExcludeStats)) { 'LiveStats' } else { 'Snapshot' }
    $logFileName = '{0}-{1}' -f (($MyInvocation.MyCommand.Name) -replace '\.ps1$',''), $logMode
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPath $OutputPath -LogFileName $logFileName
    Start-Transcript -Path $global:logTranscriptFile -Append
    WriteLog -Message "Script Environment initialized at $InitializeOutputPath"
    $OutputPath = $InitializeOutputPath
    $GlobalExportPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue ''
    WriteLog -Message "Starting $TaskName..."
    if ($PermissionsOnly) { WriteLog -Message "PermissionsOnly enabled: non-permission fields will be exported as empty strings." "INFO" }
} catch {
    Write-Host "Initialization failed: $_" -ForegroundColor Red
    exit 1
}
#endregion

try {
    $swTotalRun = [System.Diagnostics.Stopwatch]::StartNew()
    # ================= Graph HTTP code counters & error buckets ===================
    $global:GraphHttpCodeCounts = @{}
    $global:ErrorBuckets = @{}

    function Add-ErrorBucket {
        param([string]$Key)
        if ([string]::IsNullOrWhiteSpace($Key)) { $Key = "Unknown" }
        if (-not $global:ErrorBuckets.ContainsKey($Key)) { $global:ErrorBuckets[$Key] = 0 }
        $global:ErrorBuckets[$Key]++
    }

    function Log-GraphHttpError {
        param(
            [Parameter(Mandatory)]$ExceptionObj,
            [string]$Context
        )

        $code = $null; $detail = $null
        try {
            if ($ExceptionObj.ResponseStatusCode) { $code = [string]$ExceptionObj.ResponseStatusCode }
            elseif ($ExceptionObj.StatusCode)     { $code = [string]$ExceptionObj.StatusCode }
            elseif ($ExceptionObj.ErrorResponse -and $ExceptionObj.ErrorResponse.Code) { $code = [string]$ExceptionObj.ErrorResponse.Code }
            elseif ($ExceptionObj.HResult)        { $code = [string]$ExceptionObj.HResult }

            if ($ExceptionObj.Message) { $detail = $ExceptionObj.Message }
            elseif ($ExceptionObj.ToString()) { $detail = $ExceptionObj.ToString() }
        } catch { }

        if ([string]::IsNullOrWhiteSpace($code)) { $code = "Unknown" }

        # Downgrade "manager not found" to INFO (expected)
        $isManagerMissing = ($detail -match 'Request_ResourceNotFound' -or $detail -match "Resource 'manager' does not exist")
        if ($isManagerMissing) {
            WriteLog ("Graph manager missing in {0} | Detail={1}" -f $Context,$detail) "INFO"
            if (-not $global:GraphHttpCodeCounts.ContainsKey('ManagerMissing')) { $global:GraphHttpCodeCounts['ManagerMissing'] = 0 }
            $global:GraphHttpCodeCounts['ManagerMissing']++
            return
        }

        if (-not $global:GraphHttpCodeCounts.ContainsKey($code)) { $global:GraphHttpCodeCounts[$code] = 0 }
        $global:GraphHttpCodeCounts[$code]++

        WriteLog ("Graph HTTP error in {0} | Code={1} | Detail={2}" -f $Context,$code,$detail) "WARNING"
        Add-ErrorBucket ("Graph:{0}" -f $code)
    }
    # ==============================================================================


# --- Performance Profiling -------------------------------------------------------
$global:PerfCounters = [ordered]@{
    Total_Run_Seconds                 = 0
    Stage_LoadSnapshots_Seconds       = 0
    Stage_GetMailboxes_Seconds        = 0
    Stage_MailboxLoop_Seconds         = 0
    Stage_Export_Seconds              = 0
    Stage_DisconnectCleanup_Seconds   = 0
    EXO_SafeCalls_Seconds             = 0
    Graph_SafeCalls_Seconds           = 0
    Loop_PermissionsLookup_Seconds    = 0
    Loop_OULookup_Seconds             = 0
    Loop_RegionalConfig_Seconds       = 0
    Loop_Stats_Seconds                = 0
    Loop_ObjectBuild_Seconds          = 0
}

$global:PerfMailboxTimes = New-Object System.Collections.Generic.List[object]

function Add-PerfSeconds {
    param([string]$Key, [double]$Seconds)
    if ([string]::IsNullOrWhiteSpace($Key)) { return }
    # normalize keys to prevent accidental creation of duplicates (e.g. \Key\ or \"Key\")
    $Key = $Key.Trim().Trim('"','\')
    if (-not $global:PerfCounters.Contains($Key)) { $global:PerfCounters[$Key] = 0 }
    $global:PerfCounters[$Key] += $Seconds
}
# -----------------------------------------------------------------------------
# --- Resilient wrappers (Graph/EXO) -------------------------------------------
function Invoke-GraphSafe {
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [Parameter(Mandatory)][string]$Context,
        [int]$MaxRetries = 4,
        [int]$BaseDelaySeconds = 2
    )

    for ($i = 0; $i -le $MaxRetries; $i++) {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            $result = & $ScriptBlock
            $sw.Stop()
            Add-PerfSeconds -Key "Graph_SafeCalls_Seconds" -Seconds $sw.Elapsed.TotalSeconds
            return $result
        } catch {
            $sw.Stop()
            Add-PerfSeconds -Key "Graph_SafeCalls_Seconds" -Seconds $sw.Elapsed.TotalSeconds
            $msg = $_.Exception.Message
            Log-GraphHttpError -ExceptionObj $_.Exception -Context $Context
            if ($i -ge $MaxRetries) { return $null }
            $isThrottle = ($msg -match '429|TooManyRequests|throttl|ResponseEnded|prematurely')
            $delay = if ($isThrottle) { [Math]::Min(120, (60 * ($i + 1))) } else { [Math]::Min(60, ($BaseDelaySeconds * [Math]::Pow(2, $i))) }
            if ($isThrottle) { WriteLog ("Graph throttle detected in {0} - back-off {1}s" -f $Context,$delay) "WARNING" }
            Start-Sleep -Seconds $delay
        }
    }
    return $null
}

function Invoke-ExoSafe {
    param(
        [Parameter(Mandatory)][scriptblock]$ScriptBlock,
        [Parameter(Mandatory)][string]$Context,
        [int]$MaxRetries = 4,
        [int]$BaseDelaySeconds = 2
    )

    for ($i = 0; $i -le $MaxRetries; $i++) {
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        try {
            $result = & $ScriptBlock
            $sw.Stop()
            Add-PerfSeconds -Key "EXO_SafeCalls_Seconds" -Seconds $sw.Elapsed.TotalSeconds
            return $result
        } catch {
            $sw.Stop()
            Add-PerfSeconds -Key "EXO_SafeCalls_Seconds" -Seconds $sw.Elapsed.TotalSeconds
            $msg = $_.Exception.Message
            WriteLog ("EXO error in {0} (attempt {1}/{2}) | Detail={3}" -f $Context,($i+1),($MaxRetries+1),$msg) "WARNING"
            Add-ErrorBucket ("EXO:{0}" -f $Context)
            if ($i -ge $MaxRetries) { return $null }
            # --- Auth failure: token expired - attempt proactive reconnection then retry once ---
            # Only for app-only cert auth; interactive auth cannot reconnect silently (MFA required).
            $isAuthFailure = ($msg -match 'Auth failed|oAuthEventOperationId|authentication.*failed.*inner error')
            if ($isAuthFailure) {
                WriteLog ("EXO auth failure detected in {0} - attempting EXO reconnection before retry." -f $Context) "WARNING"
                Add-ErrorBucket "EXO:AuthFailure"
                if (-not $InteractiveAuth) {
                    try { Invoke-EXOProactiveReconnect } catch { WriteLog ("Reconnect attempt failed: {0}" -f $_.Exception.Message) "WARNING" }
                    Start-Sleep -Seconds 5
                    continue
                }
                return $null
            }
            # --- Non-retryable: permanent server-side errors - skip immediately ---
            $isNonRetryable = ($msg -match 'CmdletProxyException|Object reference not set|NullReferenceException|MdbAdminTaskException|critical property.*missing|Database.*missing in the ADUser|ManagementObjectNotFoundException|couldn.t be found on|Identity is a mandatory value|HttpStatusCode=401')
            if ($isNonRetryable) {
                WriteLog ("EXO non-retryable error in {0} - skipping retries." -f $Context) "WARNING"
                Add-ErrorBucket ("EXO:NonRetryable:{0}" -f $Context)
                return $null
            }
            # --- Classify retryable errors ---
            $isThrottle   = ($msg -match '429|TooManyRequests|throttl')
            $isTransient  = ($msg -match 'ResponseEnded|prematurely|pipeline.*stop|server side error|MapiNetworkErrorException|Information Store.*inaccessible|MapiExceptionProtocolError|No such host')
            $delay = if ($isThrottle) {
                [Math]::Min(120, (30 * ($i + 1)))
            } elseif ($isTransient) {
                [Math]::Min(15, ($BaseDelaySeconds * ($i + 1)))
            } else {
                [Math]::Min(60, ($BaseDelaySeconds * [Math]::Pow(2, $i)))
            }
            if ($isThrottle)  { WriteLog ("EXO throttle detected in {0} - back-off {1}s" -f $Context,$delay) "WARNING" }
            if ($isTransient) { WriteLog ("EXO transient server error in {0} - short retry {1}s" -f $Context,$delay) "WARNING" }
            Start-Sleep -Seconds $delay
        }
    }
    return $null
}
# -----------------------------------------------------------------------------

    # --- Proactive EXO reconnect (called on auth failure or periodically during long runs) ---
    function Invoke-EXOProactiveReconnect {
        WriteLog "Proactive EXO reconnect: disconnecting current session..." "INFO"
        try { Disconnect-SmartM365CloudSession -ExchangeOnline $true -Graph $false -VerboseDisconnect:$false } catch {}

        $reParams = @{ ExchangeOnline = $true; Graph = $false }
        if (-not $InteractiveAuth) {
            $reParams.AppId        = $AppId
            $reParams.Thumbprint   = $Thumb
            $reParams.TenantId     = $TenantId
            $reParams.Organization = $OrgDomain
        }
        $reResult = Connect-SmartM365CloudSession @reParams
        if (-not $reResult.ExchangeOnlineConnected) {
            WriteLog "Proactive EXO reconnect FAILED." "ERROR"
            throw "EXO reconnect failed."
        }
        WriteLog "Proactive EXO reconnect successful." "INFO"
        $script:LastEXOReconnect = Get-Date
    }

    # --- Helper to test Graph connection (used in session reuse logic) ---
    function Test-GraphConnection {
        try {
            $null = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/me" -ErrorAction Stop
            $true
        } catch {
            $false
        }
    }

    $preflightModules = @('ExchangeOnlineManagement')
    if (-not $PermissionsOnly) {
        $preflightModules += 'Microsoft.Graph.Authentication'
    }

    Invoke-SmartM365Preflight -ScriptName $TaskName -RequiredModules $preflightModules -OutputPaths @($OutputPath) | Out-Null
    #region Connection management (EXO + Graph via Connect-SmartM365CloudSession)

    # --- Exchange Online: detect need for connection ---
    $needExoConnect = $false
    if ($Connect) {
        Write-Host "Connect switch specified: existing Exchange Online session (if any) will be disconnected and reconnected..." -ForegroundColor Cyan
        Disconnect-SmartM365CloudSession -ExchangeOnline $true -Graph $false -VerboseDisconnect:$true
        $needExoConnect = $true
    } else {
        if (-not (Get-ConnectionInformation)) {
            Write-Host "No existing Exchange Online session detected. Will establish a new connection..." -ForegroundColor Cyan
            $needExoConnect = $true
        }
    }

    # --- Detect existing Graph session ---
    $graphContext = $null
    if (Get-Command Get-MgContext -ErrorAction SilentlyContinue) {
        try {
            $graphContext = Get-MgContext -ErrorAction SilentlyContinue
        } catch { }
    }

    $needGraphConnect = $false
    $connectedGraphInThisRun = $false

    if ($PermissionsOnly) {
        # PermissionsOnly does not require Graph
        $needGraphConnect = $false
    }
    else {
        # Always allow Graph connection (even if snapshots are available), so operators can reuse the same session
        # and optionally enable Graph-dependent enrichment later without changing the runbook.
        if ($Connect) {
            Write-Host "Connect switch specified: existing Graph session (if any) will be disconnected and reconnected..." -ForegroundColor Cyan
            Disconnect-SmartM365CloudSession -ExchangeOnline $false -Graph $true -VerboseDisconnect:$true
            $needGraphConnect = $true
        } else {
            if ($graphContext -and (Test-GraphConnection)) {
                Write-Host "Existing Microsoft Graph session detected. Reusing current connection." -ForegroundColor Cyan
                $needGraphConnect = $false
            } else {
                Write-Host "No existing Graph session detected. Will establish a new connection..." -ForegroundColor Cyan
                $needGraphConnect = $true
            }
        }
    }

    if ($needExoConnect -or $needGraphConnect) {
        $connectParams = @{
            ExchangeOnline = $needExoConnect
            Graph          = $needGraphConnect
        }

        # Default connection mode is app-only (certificate) for BOTH EXO and Graph.
        # Interactive authentication is used only when -InteractiveAuth is specified.
        if (-not $InteractiveAuth) {
            $connectParams.AppId        = $AppId
            $connectParams.Thumbprint   = $Thumb
            $connectParams.TenantId     = $TenantId
            $connectParams.Organization = $OrgDomain
            WriteLog -Message "Connecting with app-only certificate authentication (EXO and/or Graph as requested)." "INFO"
        } else {
            WriteLog -Message "Connecting with interactive authentication (EXO and/or Graph as requested)." "INFO"
        }

        if ($needGraphConnect) {
            $connectParams.GraphScopes = @("Directory.Read.All")
        }

        $connectResult = Connect-SmartM365CloudSession @connectParams

        if ($needExoConnect -and (-not $connectResult.ExchangeOnlineConnected)) {
            throw "Failed to connect to Exchange Online."
        }

        if ($needGraphConnect -and (-not $connectResult.GraphConnected)) {
            throw "Failed to connect to Microsoft Graph."
        }

        $connectedGraphInThisRun = $connectResult.GraphConnected
    } else {
        WriteLog -Message "Reusing existing Exchange Online/Graph connections." "INFO"
    }


    #endregion

    $graphRequiredPermissionsForPreflight = @()
    $graphProbeUrisForPreflight = @()
    if (-not $PermissionsOnly) {
        $graphRequiredPermissionsForPreflight = @('User.Read.All')
        $graphProbeUrisForPreflight = @('https://graph.microsoft.com/v1.0/users?$top=1')
    }

    Invoke-SmartM365Preflight -ScriptName $TaskName -ExchangeOnlineProbeCommands @('Get-Mailbox') -RequiredGraphApplicationPermissions $graphRequiredPermissionsForPreflight -GraphProbeUris $graphProbeUrisForPreflight | Out-Null

    # --- Retry helper ----------------------------------------------------
    function Invoke-WithRetry {
        param(
            [Parameter(Mandatory)][scriptblock]$ScriptBlock,
            [int]$MaxRetries=3,
            [int]$DelaySeconds=2,
            [object[]]$Args
        )
        for ($i=1;$i -le $MaxRetries;$i++) {
            try {
                if ($Args){
                    return & $ScriptBlock @Args
                } else {
                    return & $ScriptBlock
                }
            }
            catch {
                if ($i -lt $MaxRetries){
                    Write-Warning "Attempt $i failed: $($_.Exception.Message)"
                    Start-Sleep -Seconds ($DelaySeconds*$i)
                } else {
                    return $null
                }
            }
        }
    }

    # ====== CACHES ======

    # ================= Local enrichment caches (CSV snapshots) =================
    # These caches accelerate enrichment and reduce Graph dependency.
    # They are skipped in PermissionsOnly mode.
            $global:M365UsersCache   = @{}   # key: UPN (lower)
    $global:AdUsersUpnCache  = @{}   # key: UPN (lower) -> row
    $global:AdUsersDnCache   = @{}   # key: DN  (lower) -> DisplayName

    function Get-OUFromDistinguishedName {
        param([string]$DistinguishedName)
        if ([string]::IsNullOrWhiteSpace($DistinguishedName)) { return "" }
        try {
            if ($DistinguishedName -match '^[^,]+,(?<ou>.+)$') { return $matches['ou'] }
        } catch {}
        return $DistinguishedName
    }


    function Normalize-Key {
        param([string]$Value)
        if ([string]::IsNullOrWhiteSpace($Value)) { return "" }
        try { return $Value.Trim().ToLowerInvariant() } catch { return ($Value.ToLower()) }
    }


    function Load-M365UsersCache {
        param([string]$Path)
        $dict = @{}
        try {
            $resolvedPath = Resolve-SmartM365CsvPathWithSharePointFallback -Path $Path -Description 'M365 users snapshot'
            if ([string]::IsNullOrWhiteSpace($resolvedPath)) {
                WriteLog -Message "M365 users snapshot not found at $Path. Will continue without M365 enrichment." "WARNING"
                return $dict
            }
            $fi = Get-Item -LiteralPath $resolvedPath -ErrorAction Stop
    WriteLog -Message "Loading M365 users snapshot: $($fi.FullName) (LastWrite: $($fi.LastWriteTime))" "INFO"

            $rows = Invoke-WithRetry -ScriptBlock { param($p) Import-Csv -LiteralPath $p -ErrorAction Stop } -Args @($resolvedPath)
            $count = 0
            foreach ($r in $rows) {
                $upn = [string]$r.'User principal name'
                if ([string]::IsNullOrWhiteSpace($upn)) { continue }
                $dict[$upn.ToLower()] = $r
                $count++
            }
            WriteLog -Message "M365 users snapshot loaded: $count entries." "INFO"
        } catch {
            WriteLog -Message "Failed to load M365 users snapshot: $($_.Exception.Message). Continuing without M365 enrichment." "WARNING"
        }
        return $dict
    }

    function Load-AdUsersCache {
        param([string]$Path)

        $upnDict = @{}
        $dnDict  = @{}

        try {
            $resolvedPath = Resolve-SmartM365CsvPathWithSharePointFallback -Path $Path -Description 'AD users snapshot'
            if ([string]::IsNullOrWhiteSpace($resolvedPath)) {
                WriteLog -Message "AD users snapshot not found at $Path. Will continue without AD enrichment." "WARNING"
                return @{ Upn = $upnDict; Dn = $dnDict }
            }
            $fi = Get-Item -LiteralPath $resolvedPath -ErrorAction Stop
            WriteLog -Message "Loading AD users snapshot: $($fi.FullName) (LastWrite: $($fi.LastWriteTime))" "INFO"

            $rows = Invoke-WithRetry -ScriptBlock { param($p) Import-Csv -LiteralPath $p -ErrorAction Stop } -Args @($resolvedPath)
            $count = 0
            foreach ($r in $rows) {
                $upn = [string]$r.UserPrincipalName
                if (-not [string]::IsNullOrWhiteSpace($upn)) {
                    $upnDict[$upn.ToLower()] = $r
                }
                $dn = [string]$r.DistinguishedName
                $dnDisp = [string]$r.DisplayName
                if (-not [string]::IsNullOrWhiteSpace($dn) -and -not [string]::IsNullOrWhiteSpace($dnDisp)) {
                    $dnDict[$dn.ToLower()] = $dnDisp
                }
                $count++
            }
            WriteLog -Message "AD users snapshot loaded: $count rows. Indexes: UPN=$($upnDict.Count), DN=$($dnDict.Count)." "INFO"
        } catch {
            WriteLog -Message "Failed to load AD users snapshot: $($_.Exception.Message). Continuing without AD enrichment." "WARNING"
        }

        return @{ Upn = $upnDict; Dn = $dnDict }
    }

    function Resolve-ManagerDisplayNameFromDn {
        param(
            [string]$ManagerDn
        )
        if ([string]::IsNullOrWhiteSpace($ManagerDn)) { return "" }
        try {
            $k = $ManagerDn.ToLower()
            if ($global:AdUsersDnCache.ContainsKey($k)) { return [string]$global:AdUsersDnCache[$k] }
        } catch {}
        # Fallback: return DN raw (requested)
        return $ManagerDn
    }

    if (-not $PermissionsOnly) {
        $swStageSnapshots = [System.Diagnostics.Stopwatch]::StartNew()
        $global:M365UsersCache = Load-M365UsersCache -Path $M365UsersCsvPath
        $ad = Load-AdUsersCache -Path $AdUsersCsvPath
        $global:AdUsersUpnCache = $ad.Upn
        $global:AdUsersDnCache  = $ad.Dn
    } else {
        WriteLog -Message "PermissionsOnly enabled: skipping AD/M365 snapshot loads." "INFO"
    }
    # ===========================================================================

    $OuCache=@{}
    $adSnapshotAvailable = $false
    try {
        if ($global:AdUsersUpnCache -and $global:AdUsersUpnCache.Count -gt 0) { $adSnapshotAvailable = $true }
        elseif ($global:AdUsersDnCache -and $global:AdUsersDnCache.Count -gt 0) { $adSnapshotAvailable = $true }
    } catch { $adSnapshotAvailable = $false }

    # Bulk prefill OuCache from AD snapshot to avoid per-mailbox re-processing in Get-OU-Fast
    if ($adSnapshotAvailable -and $global:AdUsersUpnCache -and $global:AdUsersUpnCache.Count -gt 0) {
        $ouPrefillCount = 0
        foreach ($kv in $global:AdUsersUpnCache.GetEnumerator()) {
            try {
                $dn = [string]$kv.Value.DistinguishedName
                if (-not [string]::IsNullOrWhiteSpace($dn)) {
                    $OuCache[$kv.Key] = Get-OUFromDistinguishedName $dn
                    $ouPrefillCount++
                }
            } catch {}
        }
        WriteLog -Message "OuCache prefilled from AD snapshot: $ouPrefillCount entries." "INFO"
    }

    if ($needGraphConnect -and (-not $PermissionsOnly)) {
        if ($adSnapshotAvailable) {
            WriteLog -Message "AD snapshot available: skipping bulk Graph prefetch; Graph will be used only on-demand for missing values."
        }
        elseif (-not $Top100) {
            try {
                WriteLog -Message "Prefetching Graph users (id, upn, onPremisesDN)..."
                $allUsers = Invoke-WithRetry -ScriptBlock { Get-MgUser -All -Property "id,userPrincipalName,onPremisesDistinguishedName" }
                foreach ($u in $allUsers) {
                    if ($u.Id) { $OuCache[$u.Id] = $u.OnPremisesDistinguishedName }
                    if ($u.UserPrincipalName) { $OuCache[$u.UserPrincipalName] = $u.OnPremisesDistinguishedName }
                }
                WriteLog -Message "Graph users prefetched: $($OuCache.Count) keys loaded."
            } catch {
                Log-GraphHttpError -ExceptionObj $_.Exception -Context "Get-MgUser(All) prefetch"
                WriteLog -Message "Bulk Graph prefetch failed: $($_.Exception.Message). Fallback on-demand." "WARNING"
                Add-ErrorBucket "GraphPrefetch"
            }
        } else {
            WriteLog -Message "Top100 enabled: skipping bulk Graph prefetch; OU will be resolved on-demand."
        }
    } else {
        WriteLog -Message "PermissionsOnly enabled or Graph not requested: skipping Graph prefetch and OU resolution."
    }

    if ($swStageSnapshots) { try { $swStageSnapshots.Stop() } catch {} ; Add-PerfSeconds -Key "Stage_LoadSnapshots_Seconds" -Seconds $swStageSnapshots.Elapsed.TotalSeconds }
function Get-OU-Fast {
        param([string]$Key)

        if ([string]::IsNullOrWhiteSpace($Key)) { return $null }

        $kNorm = Normalize-Key $Key
        if ($OuCache.ContainsKey($kNorm)) { return $OuCache[$kNorm] }

        # Prefer AD snapshot (fast, no network)
        try {
            if ($global:AdUsersUpnCache -and $global:AdUsersUpnCache.Count -gt 0) {
                if ($global:AdUsersUpnCache.ContainsKey($kNorm)) {
                    $adRow = $global:AdUsersUpnCache[$kNorm]
                    $dn = [string]$adRow.DistinguishedName
                    if (-not [string]::IsNullOrWhiteSpace($dn)) {
                        $val = Get-OUFromDistinguishedName $dn
                        $OuCache[$kNorm] = $val
                        return $val
                    }
                }
            }
        } catch {}

        # Prefer M365 snapshot if it contains an on-premises DN field
        try {
            if ($global:M365UsersCache -and $global:M365UsersCache.Count -gt 0) {
                if ($global:M365UsersCache.ContainsKey($kNorm)) {
                    $m365Row = $global:M365UsersCache[$kNorm]
                    $dn = ""
                    if ($m365Row.PSObject.Properties.Name -contains 'OnPremisesDistinguishedName') { $dn = [string]$m365Row.OnPremisesDistinguishedName }
                    elseif ($m365Row.PSObject.Properties.Name -contains 'On-Premises distinguished name') { $dn = [string]$m365Row.'On-Premises distinguished name' }
                    elseif ($m365Row.PSObject.Properties.Name -contains 'On-premises distinguished name') { $dn = [string]$m365Row.'On-premises distinguished name' }
                    elseif ($m365Row.PSObject.Properties.Name -contains 'onPremisesDistinguishedName') { $dn = [string]$m365Row.onPremisesDistinguishedName }

                    if (-not [string]::IsNullOrWhiteSpace($dn)) {
                        $val = Get-OUFromDistinguishedName $dn
                        $OuCache[$kNorm] = $val
                        return $val
                    }
                }
            }
        } catch {}

        # Fallback to Graph on-demand (network)
        if (-not $needGraphConnect) { return $null }

        try {
            $u = Invoke-GraphSafe -Context ("Get-MgUser({0}) onPremisesDN" -f $Key) -ScriptBlock {
                Get-MgUser -UserId $Key -Property "userPrincipalName,onPremisesDistinguishedName" -ErrorAction Stop
            }

            if ($u) {
                $dn = [string]$u.OnPremisesDistinguishedName
                $val = if ([string]::IsNullOrWhiteSpace($dn)) { "" } else { Get-OUFromDistinguishedName $dn }

                $OuCache[$kNorm] = $val

                # Also cache by UPN if returned
                try {
                    if ($u.UserPrincipalName) {
                        $upnNorm = Normalize-Key ([string]$u.UserPrincipalName)
                        if (-not [string]::IsNullOrWhiteSpace($upnNorm)) { $OuCache[$upnNorm] = $val }
                    }
                } catch {}

                return $val
            }
        } catch {
            Log-GraphHttpError -ExceptionObj $_.Exception -Context ("Get-MgUser({0}) on-demand OU" -f $Key)
        }

        return $null
    }

    $RecipientCache=@{}
    function Resolve-RecipientCached {
        param([string]$Identity)
        if ([string]::IsNullOrWhiteSpace($Identity)){return $null}
        if ($RecipientCache.ContainsKey($Identity)){return $RecipientCache[$Identity]}
        try{
            $r=Get-Recipient -Identity $Identity -ErrorAction Stop
            $val=if($r.PrimarySmtpAddress){$r.PrimarySmtpAddress.ToString()}elseif($r.ExternalEmailAddress){$r.ExternalEmailAddress.AddressString}else{$r.DisplayName}
            $RecipientCache[$Identity]=$val
            $val
        } catch {
            $RecipientCache[$Identity]=$Identity
            $Identity
        }
    }

    # -------- FR/EN size conversions --------
    function Convert-QuotaToGB {
        param($input)
        if ($null -eq $input) { return "" }
        try {
            if ($input.PSObject.Methods.Name -contains 'ToBytes') { return [math]::Round(([double]$input.ToBytes())/1GB, 2) }
            elseif ($input.PSObject.Properties.Name -contains 'Value' -and $input.Value -and ($input.Value.PSObject.Methods.Name -contains 'ToBytes')) { return [math]::Round(([double]$input.Value.ToBytes())/1GB, 2) }
            elseif ($input.PSObject.Properties.Name -contains 'Bytes') { return [math]::Round(([double]$input.Bytes)/1GB, 2) }

            if ($input -is [double] -or $input -is [int64] -or $input -is [decimal]) { return [math]::Round(([double]$input)/1GB, 2) }

            $s = ([string]$input) -replace [char]0xA0, ' '
            $s = $s -replace [char]0x202F, ' '
            $s = $s -replace [char]0x2009, ' '
            $s = $s.Trim()

            if ($s -match '(?i)unlimited|illimit') { return 'Unlimited' }

            if ($s -match '\((?<bytes>[\d\.\,\s]+)\s*(bytes|octets)\)') {
                $raw = $matches['bytes'] -replace '[\s\.,]', ''
                if ($raw -match '^\d+$') { return [math]::Round(([double]$raw)/1GB, 2) }
            }

            if ($s -match '^\s*(?<num>[\d\.,]+)\s*(?<unit>TB|GB|MB|KB|B|To|Go|Mo|Ko|o)\b') {
                $n = [double](($matches['num'] -replace ',','.') )
                switch -regex ($matches['unit']) {
                    'TB|To' { $bytes = $n * 1TB }
                    'GB|Go' { $bytes = $n * 1GB }
                    'MB|Mo' { $bytes = $n * 1MB }
                    'KB|Ko' { $bytes = $n * 1KB }
                    'B|o'   { $bytes = $n }
                }
                if ($bytes -ne $null) { return [math]::Round($bytes/1GB, 2) }
            }
            return ""
        } catch { "" }
    }

    # -------- MB integer conversion (no separators, rounded) --------
    
    # -------- MB integer conversion from GB (fast path) --------
    function Convert-GBToMBInteger {
        param($Gb)

        if ($null -eq $Gb) { return "" }

        try {
            if ($Gb -is [double] -or $Gb -is [int] -or $Gb -is [int64] -or $Gb -is [decimal]) {
                return [int][math]::Round(([double]$Gb * 1024), 0)
            }

            $s = ([string]$Gb).Trim()
            if ([string]::IsNullOrWhiteSpace($s)) { return "" }
            if ($s -match '(?i)unlimited|illimit') { return "" }

            $n = [double](($s -replace ',','.') )
            return [int][math]::Round(($n * 1024), 0)
        } catch {
            return ""
        }
    }

function Convert-ToMBInteger {
        param($input)

        if ($null -eq $input) { return "" }

        try {
            # Case 1: EXO/legacy ByteQuantifiedSize-like objects
            if ($input.PSObject.Properties.Name -contains 'Value' -and $input.Value -and ($input.Value.PSObject.Methods.Name -contains 'ToBytes')) {
                return [int][math]::Round(($input.Value.ToBytes() / 1MB), 0)
            }
            if ($input.PSObject.Methods.Name -contains 'ToBytes') {
                return [int][math]::Round(($input.ToBytes() / 1MB), 0)
            }

            # Case 2: objects exposing Bytes
            if ($input.PSObject.Properties.Name -contains 'Bytes' -and $input.Bytes -ne $null) {
                return [int][math]::Round((([double]$input.Bytes) / 1MB), 0)
            }

            # Case 3: numeric already bytes
            if ($input -is [double] -or $input -is [int64] -or $input -is [decimal] -or $input -is [int]) {
                return [int][math]::Round((([double]$input) / 1MB), 0)
            }

            # Case 4: string parsing (EN/FR)
            $s = ([string]$input) -replace [char]0xA0, ' '
            $s = $s -replace [char]0x202F, ' '
            $s = $s -replace [char]0x2009, ' '
            $s = $s.Trim()

            if ($s -match '(?i)unlimited|illimit') { return "" }

            # Prefer explicit bytes/octet count in parentheses
            if ($s -match '\((?<bytes>[\d\.\,\s]+)\s*(bytes|octets)\)') {
                $raw = $matches['bytes'] -replace '[\s\.,]', ''
                if ($raw -match '^\d+$') {
                    return [int][math]::Round((([double]$raw) / 1MB), 0)
                }
            }

            # Fallback: "<num> <unit>"
            if ($s -match '^\s*(?<num>[\d\.,]+)\s*(?<unit>TB|GB|MB|KB|B|To|Go|Mo|Ko|o)\b') {
                $n = [double](($matches['num'] -replace ',','.') )
                switch -regex ($matches['unit']) {
                    'TB|To' { $bytes = $n * 1TB }
                    'GB|Go' { $bytes = $n * 1GB }
                    'MB|Mo' { $bytes = $n * 1MB }
                    'KB|Ko' { $bytes = $n * 1KB }
                    'B|o'   { $bytes = $n }
                }
                if ($bytes -ne $null) { return [int][math]::Round(($bytes / 1MB), 0) }
            }
        } catch {}

        return ""
    }

    # ====== Exact EXO stats (optional extra call for LUAT) ======
    function Get-ExoStatsExact {
        param(
            [Parameter(Mandatory)][string]$Identity,
            [switch]$Archive,
            [switch]$IncludeLastUserActionTime
        )

        # Detect -ShowProgress availability
        $supportsShowProgress = $false
        try {
            $cmd = Get-Command Get-EXOMailboxStatistics -ErrorAction Stop
            if ($cmd -and $cmd.Parameters.ContainsKey('ShowProgress')) { $supportsShowProgress = $true }
        } catch { }

        # Common params (mute EXO progress if supported)
        $commonParams = @{
            ErrorAction       = 'SilentlyContinue'
            InformationAction = 'SilentlyContinue'
        }
        if ($supportsShowProgress) { $commonParams['ShowProgress'] = $false }

        $oldProg = $ProgressPreference
        $ProgressPreference = 'SilentlyContinue'
        try {
            $baseCall = {
                param($id,$isArchive,$p)
                if ($isArchive) {
                    Get-EXOMailboxStatistics -Identity $id -Archive @p 2>$null 6>$null
                } else {
                    Get-EXOMailboxStatistics -Identity $id @p 2>$null 6>$null
                }
            }
            $extraCall = {
                param($id,$isArchive,$p)
                if ($isArchive) {
                    Get-EXOMailboxStatistics -Identity $id -Archive -Properties LastUserActionTime @p 2>$null 6>$null
                } else {
                    Get-EXOMailboxStatistics -Identity $id -Properties LastUserActionTime @p 2>$null 6>$null
                }
            }

            $res = & $baseCall $Identity $Archive.IsPresent $commonParams
            if ($res -is [System.Array]) { $res = $res | Select-Object -First 1 }

            # Only perform the extra call if requested
            if ($IncludeLastUserActionTime) {
                $res2 = & $extraCall $Identity $Archive.IsPresent $commonParams
                if ($res2 -is [System.Array]) { $res2 = $res2 | Select-Object -First 1 }
                if ($res2) { $res = $res2 }
            }

            return $res
        }
        finally {
            $ProgressPreference = $oldProg
        }
    }

    # ====== Permissions snapshot (UPN) ======
    function Load-PermissionsSnapshot {
        param([string]$ScriptCsvLogFolderPath)
        $dict=@{}
        try{
            $permFile=Join-Path $ScriptCsvLogFolderPath "Exchange_EXO_Mailboxes_AllDomains_Permissions.csv"
            $resolvedPermFile = Resolve-SmartM365CsvPathWithSharePointFallback -Path $permFile -Description 'Permissions snapshot'
            if ([string]::IsNullOrWhiteSpace($resolvedPermFile)){
                WriteLog -Message "Permissions snapshot not found at $permFile. Default CSV will have blank permission columns." "WARNING"
                return $dict
            }
            $fi=Get-Item -LiteralPath $resolvedPermFile -ErrorAction Stop
            WriteLog -Message "Loading permissions snapshot: $($fi.FullName) (LastWrite: $($fi.LastWriteTime))"
            $rows=Invoke-WithRetry -ScriptBlock { param($p) Import-Csv -LiteralPath $p -ErrorAction Stop } -Args @($resolvedPermFile)
            $count=0
            foreach($row in $rows){
                $upn=[string]$row.UserPrincipalName
                if([string]::IsNullOrWhiteSpace($upn)){continue}
                $key=$upn.ToLower()
                $dict[$key] = [PSCustomObject]@{
                    SendAs             = [string]$row.SendAs
                    FullAccess         = [string]$row.FullAccess
                    GrantSendOnBehalfTo= [string]$row.GrantSendOnBehalfTo
                }
                $count++
            }
            WriteLog -Message "Permissions snapshot loaded: $count entries."
            return $dict
        } catch {
            WriteLog -Message "Failed to load permissions snapshot: $($_.Exception.Message). Default CSV will have blank permission columns." "WARNING"
            return $dict
        }
    }

    
    # ====== Stats snapshot (PrimarySmtpAddress) ======
    function Load-StatsSnapshot {
        param([string]$StatsCsvPath)
        $dict=@{}
        try{
            if ([string]::IsNullOrWhiteSpace($StatsCsvPath)) { return $dict }
            $resolvedStatsCsvPath = Resolve-SmartM365CsvPathWithSharePointFallback -Path $StatsCsvPath -Description 'Stats snapshot'
            if ([string]::IsNullOrWhiteSpace($resolvedStatsCsvPath)){
                WriteLog -Message "Stats snapshot not found at $StatsCsvPath. Stats columns will be blank." "WARNING"
                return $dict
            }

            $fi=Get-Item -LiteralPath $resolvedStatsCsvPath -ErrorAction Stop
            WriteLog -Message "Loading stats snapshot: $($fi.FullName) (LastWrite: $($fi.LastWriteTime))"

            $rows=Invoke-WithRetry -ScriptBlock { param($p) Import-Csv -LiteralPath $p -ErrorAction Stop } -Args @($resolvedStatsCsvPath)
            $count=0
            foreach($row in $rows){
                $smtp=[string]$row.PrimarySmtpAddress
                if([string]::IsNullOrWhiteSpace($smtp)){continue}
                $key = Normalize-Key $smtp

                $dict[$key] = [PSCustomObject]@{
                    Identity                      = [string]$row.Identity
                    UserPrincipalName             = [string]$row.UserPrincipalName
                    PrimarySmtpAddress            = [string]$row.PrimarySmtpAddress

                    TotalItemSizeGB               = [string]$row.TotalItemSizeGB
                    ItemCount                     = [string]$row.ItemCount
                    LastLogonTime                 = [string]$row.LastLogonTime
                    LastUserActionTime            = [string]$row.LastUserActionTime

                    Archive_TotalItemSizeGB       = [string]$row.Archive_TotalItemSizeGB
                    Archive_ItemCount             = [string]$row.Archive_ItemCount
                    Archive_LastLogonTime         = [string]$row.Archive_LastLogonTime
                }
                $count++
            }
            WriteLog -Message "Stats snapshot loaded: $count entries."
            return $dict
        } catch {
            WriteLog -Message "Failed to load stats snapshot: $($_.Exception.Message). Stats columns will be blank." "WARNING"
            return $dict
        }
    }


    # ====== Graph enrichment (manager optional; Enabled only if -ResolveManager) ======
    function Get-GraphUserDetails {
param (
    [string]$UserId,
    [switch]$ResolveManager
)

if ([string]::IsNullOrWhiteSpace($UserId)) { return $null }

$response = Invoke-GraphSafe -Context ("Get-MgUser({0})" -f $UserId) -ScriptBlock {
    Get-MgUser -UserId $UserId -Property "displayName,department,jobTitle,companyName,onPremisesDistinguishedName,userType,accountEnabled,officeLocation,onPremisesSyncEnabled,onPremisesImmutableId" -ErrorAction Stop
}

if (-not $response) { return $null }

$managerName = ""
if ($ResolveManager) {
    $manager = Invoke-GraphSafe -Context ("Get-MgUserManager({0})" -f $UserId) -ScriptBlock {
        Get-MgUserManager -UserId $UserId -ErrorAction Stop
    }

    if ($manager) {
        if ($manager.PSObject.Properties['DisplayName']) {
            $managerName = [string]$manager.DisplayName
        } elseif ($manager.PSObject.Properties['AdditionalProperties'] -and $manager.AdditionalProperties.displayName) {
            $managerName = [string]$manager.AdditionalProperties.displayName
        }
    }
}

return @{
    DisplayName          = $response.DisplayName
    Department           = $response.Department
    Title                = $response.JobTitle
    Company              = $response.CompanyName
    OrganizationalUnit   = $response.OnPremisesDistinguishedName
    UserType             = $response.UserType
    AccountEnabled       = $response.AccountEnabled
    OfficeLocation       = $response.OfficeLocation
    Manager              = $managerName
    IsDirSynced          = $response.OnPremisesSyncEnabled
    OnPremisesImmutableId = if ($response.OnPremisesImmutableId) { $response.OnPremisesImmutableId } else { "" }
}
    }
    # ====== Robust RecipientTypeDetails ======
    function Get-RecipientTypeDetailsSafe {
        param(
            [Parameter(Mandatory)][object]$MailboxObj
        )

        # Fast path: already present on the object returned by Get-EXOMailbox
        try {
            if ($MailboxObj.PSObject.Properties.Name -contains 'RecipientTypeDetails' -and $MailboxObj.RecipientTypeDetails) {
                return $MailboxObj.RecipientTypeDetails.ToString()
            }
        } catch {}

        # Optional (expensive) live resolution
        if ($ResolveRecipientTypeDetailsLive) {
            try {
                $rec = Get-Recipient -Identity $MailboxObj.UserPrincipalName -ErrorAction SilentlyContinue 2>$null
                if ($rec -and $rec.RecipientTypeDetails) { return $rec.RecipientTypeDetails.ToString() }
            } catch {}

            try {
                $oldInfo = $InformationPreference
                $InformationPreference = 'SilentlyContinue'
                try {
                    $mbById = Get-EXOMailbox -Identity $MailboxObj.UserPrincipalName -ErrorAction SilentlyContinue 2>$null 6>$null
                    if ($mbById -and $mbById.RecipientTypeDetails) { return $mbById.RecipientTypeDetails.ToString() }
                }
                finally {
                    $InformationPreference = $oldInfo
                }
            } catch {}

            try {
                $mbx = Get-Mailbox -Identity $MailboxObj.UserPrincipalName -ErrorAction SilentlyContinue 2>$null
                if ($mbx -and $mbx.RecipientTypeDetails) { return $mbx.RecipientTypeDetails.ToString() }
            } catch {}
        }

        return ""
    }
# ================= Column orders (must remain stable) =================
    $columnOrderDefault = @(
        "DisplayName","UserPrincipalName","PrimarySmtpAddress","RecipientTypeDetails","MailboxType","SamAccountName","MailboxGuid","EmailAddresses",
        "Company","Department","Title","OfficeLocation","Manager","OrganizationalUnit","UserType","AccountEnabled","AssignedLicenses",
        "ForwardingSmtpAddress","ForwardingAddress","DeliverToMailboxAndForward","IsMailboxEnabled","IsEmailAddressPolicyEnabled","HiddenFromAddressListsEnabled","IsDirSynced",
        "RetentionHoldEnabled","LitigationHoldEnabled","RetentionPolicy","ThrottlingPolicy","RoleAssignmentPolicy","UseDatabaseQuotaDefaults",
        "IssueWarningQuotaGB","ProhibitSendQuotaGB","ProhibitSendReceiveQuotaGB","WhenMailboxCreated","LastExchangeChangedTime","Languages",
        "LastUserActionTime","TotalItemSizeGB","ItemCount","LastLogonTime","SendAs","FullAccess","GrantSendOnBehalfTo",
        "ArchiveStatus","ArchiveName","ArchiveQuotaGB","ArchiveWarningQuotaGB",
        "CustomAttribute1","CustomAttribute2","CustomAttribute3","CustomAttribute4","CustomAttribute5","CustomAttribute6","CustomAttribute7","CustomAttribute8","CustomAttribute9","CustomAttribute10",
        "CustomAttribute11","CustomAttribute12","CustomAttribute13","CustomAttribute14","CustomAttribute15","ExternalDirectoryObjectId","MailboxPlan",
        "ImmutableId","OnPremisesImmutableId"
    )
    $columnOrderArchive = @('UserPrincipalName','PrimarySmtpAddress','OrganizationalUnit','Archive_TotalItemSizeGB','Archive_ItemCount','Archive_LastLogonTime','ImmutableId')
    # =====================================================================

    #region Get mailboxes
    try {
        WriteLog "Retrieving Exchange Online mailboxes..." "INFO"
        $swStageGetMailboxes = [System.Diagnostics.Stopwatch]::StartNew()


        # Mute Information stream during EXO bulk call
        $oldInfo = $InformationPreference
        $InformationPreference = 'SilentlyContinue'
        try {
            $exoSupportsRtdFilter = $false
            try {
                $cmd = Get-Command Get-EXOMailbox -ErrorAction Stop
                if ($cmd -and $cmd.Parameters.ContainsKey("RecipientTypeDetails")) { $exoSupportsRtdFilter = $true }
            } catch {}
            $RecipientTypeDetailsFilter = $null
            if ($exoSupportsRtdFilter) {
                # Exclude known system mailbox types at source to avoid expensive per-mailbox RTD resolution.
                $RecipientTypeDetailsFilter = @("UserMailbox","SharedMailbox","RoomMailbox","EquipmentMailbox","LinkedMailbox")
            }

            if ($PermissionsOnly) {
                # Minimal properties for permissions collection (keep RTD + GrantSendOnBehalfTo for delegation)
                $mailboxes = if ($RecipientTypeDetailsFilter) {
                    Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop -RecipientTypeDetails $RecipientTypeDetailsFilter -Properties `
                        RecipientTypeDetails, GrantSendOnBehalfTo, ExternalDirectoryObjectId, ArchiveStatus 6>$null
                } else {
                    Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop -Properties `
                        RecipientTypeDetails, GrantSendOnBehalfTo, ExternalDirectoryObjectId, ArchiveStatus 6>$null
                }
            } else {
                $mailboxes = if ($RecipientTypeDetailsFilter) {
                    Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop -RecipientTypeDetails $RecipientTypeDetailsFilter -Properties `
                        RecipientTypeDetails, WhenMailboxCreated, LastExchangeChangedTime, Languages, GrantSendOnBehalfTo, ExternalDirectoryObjectId, ArchiveStatus, `
                        EmailAddresses, UseDatabaseQuotaDefaults, MailboxPlan, ExchangeGuid, EmailAddressPolicyEnabled, HiddenFromAddressListsEnabled, `
                        RetentionHoldEnabled, LitigationHoldEnabled, RetentionPolicy, ThrottlingPolicy, RoleAssignmentPolicy `
                        -PropertySets Retention,StatisticsSeed,Delivery 6>$null
                } else {
                    Get-EXOMailbox -ResultSize Unlimited -ErrorAction Stop -Properties `
                        RecipientTypeDetails, WhenMailboxCreated, LastExchangeChangedTime, Languages, GrantSendOnBehalfTo, ExternalDirectoryObjectId, ArchiveStatus, `
                        EmailAddresses, UseDatabaseQuotaDefaults, MailboxPlan, ExchangeGuid, EmailAddressPolicyEnabled, HiddenFromAddressListsEnabled, `
                        RetentionHoldEnabled, LitigationHoldEnabled, RetentionPolicy, ThrottlingPolicy, RoleAssignmentPolicy `
                        -PropertySets Retention,StatisticsSeed,Delivery 6>$null
                }
            }
        }
        finally { $InformationPreference = $oldInfo }

        if ($IsMaxItemsRun) {
            $pre = $mailboxes.Count
            $mailboxes = $mailboxes | Sort-Object UserPrincipalName | Select-Object -First $MaxItems
            WriteLog "MaxItems enabled: restricted from $pre to $($mailboxes.Count) mailboxes."
        }
        elseif ($Top100) {
            $pre = $mailboxes.Count
            $mailboxes = $mailboxes | Sort-Object UserPrincipalName | Select-Object -First 100
            WriteLog "Top100 enabled: restricted from $pre to $($mailboxes.Count) mailboxes."
        }

        try {
            $peek = $mailboxes | Select-Object -First 3 | Select-Object UserPrincipalName,DisplayName,RecipientTypeDetails
            foreach($p in $peek){
                WriteLog -Message ("Peek MBX: UPN={0} | DN={1} | RTD={2}" -f $p.UserPrincipalName,$p.DisplayName,$p.RecipientTypeDetails)
            }
        } catch {
            if ($swStageGetMailboxes) { try { $swStageGetMailboxes.Stop() } catch {} ; Add-PerfSeconds -Key "Stage_GetMailboxes_Seconds" -Seconds $swStageGetMailboxes.Elapsed.TotalSeconds }
        }

        WriteLog "Successfully retrieved $($mailboxes.Count) mailboxes for processing." "INFO"
        if ($swStageGetMailboxes) { $swStageGetMailboxes.Stop(); Add-PerfSeconds -Key "Stage_GetMailboxes_Seconds" -Seconds $swStageGetMailboxes.Elapsed.TotalSeconds }
    } catch {
        WriteLog "Failed to retrieve mailboxes: $_" "ERROR"
        Write-Host "Unable to retrieve mailboxes. Stopping script." -ForegroundColor Red
        throw
    }
    #endregion

    # Snapshot permissions if -IncludePerm is not used (NoPermInDefault)
    $PermSnapshot = $null
    if (-not $IncludePerm) { $PermSnapshot = Load-PermissionsSnapshot -ScriptCsvLogFolderPath $PermSnapshotDir }

    # Snapshot mailbox stats by default (standard runs do NOT call Get-EXOMailboxStatistics).
    # Live stats are collected only when -IncludeStats is used (unless -ExcludeStats is explicitly set).
    $StatsSnapshot = $null
    $StatsSnapshotPath = $StatsSnapshotCsvPath

    $UseLiveStats = ($IncludeStats -and (-not $ExcludeStats) -and (-not $PermissionsOnly))

    # Compute stats source label (written to every row in the Stats CSV, only when -IncludeStats)
    $statsSourceLabel = if ($UseLiveStats) { "Live:" + (Get-Date -Format "yyyy-MM-dd HH:mm:ss") } else { "" }

    if ((-not $UseLiveStats) -and (-not $PermissionsOnly)) {
        $StatsSnapshot = Load-StatsSnapshot -StatsCsvPath $StatsSnapshotPath
        if (-not $StatsSnapshot -or $StatsSnapshot.Count -eq 0) {
            WriteLog -Message "Stats snapshot is empty or missing. Stats columns will be blank." "WARNING"
        }
    }
# Fallback Get-Mailbox for specific fields
    function Get-MailboxFallbackFields {
        param([string]$Upn)
        try {
            $mbx = Get-Mailbox -Identity $Upn -ErrorAction SilentlyContinue 2>$null
            if ($mbx) {
                return @{
                    EmailAddressPolicyEnabled     = $mbx.EmailAddressPolicyEnabled
                    MailboxPlan                   = $mbx.MailboxPlan
                    HiddenFromAddressListsEnabled = $mbx.HiddenFromAddressListsEnabled
                    RetentionHoldEnabled          = $mbx.RetentionHoldEnabled
                    LitigationHoldEnabled         = $mbx.LitigationHoldEnabled
                    RetentionPolicy               = $mbx.RetentionPolicy
                    ThrottlingPolicy              = $mbx.ThrottlingPolicy
                    RoleAssignmentPolicy          = $mbx.RoleAssignmentPolicy
                }
            }
        } catch {}
        return @{}
    }

    $errors=@()
    $resultsDetailed=@()
    $resultsArchiveOnly=@()
    $resultsPerm=@()
    $resultsStats=@()
    $total=$mailboxes.Count
    $counter=0
    $regionalAttempted = 0
    $regionalOkWithData = 0
    $regionalOkEmpty = 0
    $regionalFailed = 0
    $regionalFromCache = 0
    $regionalSkippedEmptyCache = 0

    # Regional configuration caches (per-run)
    $script:RegionalConfigCache = @{}      # key: upn (lower) -> region object
    $script:RegionalEmptyCache  = New-Object 'System.Collections.Generic.HashSet[string]'
    $systemTypesFinal = @("DiscoveryMailbox","ArbitrationMailbox","AuditLogMailbox","MailboxPlan","PublicFolderMailbox","SchedulingMailbox")

    $swStageLoop = [System.Diagnostics.Stopwatch]::StartNew()
    $script:LastEXOReconnect   = Get-Date
    $ReconnectIntervalMinutes  = 45
    $swRunTotal       = [System.Diagnostics.Stopwatch]::StartNew()
    $ProgressInterval = 500
    $MaxRunExceeded   = $false
    $ManualStop       = $false

    # -----------------------------------------------------------------------
    # -PermissionsOnly + ParallelThrottle > 1 : parallel permission collection
    # Each thread opens its own EXO session (app-only cert) to avoid throttling.
    # -----------------------------------------------------------------------
    if ($PermissionsOnly -and $ParallelThrottle -gt 1 -and (-not $InteractiveAuth)) {
        WriteLog ("Starting parallel permission collection - ThrottleLimit={0}." -f $ParallelThrottle) "INFO"

        # Filter system mailboxes before dispatching to threads (avoids wasted connections)
        $systemTypesFinalP = @("DiscoveryMailbox","ArbitrationMailbox","AuditLogMailbox","MailboxPlan","PublicFolderMailbox","SchedulingMailbox")
        $mailboxesForPerm  = $mailboxes | Where-Object {
            $rtdP = if ($_.PSObject.Properties.Name -contains 'RecipientTypeDetails' -and $_.RecipientTypeDetails) { $_.RecipientTypeDetails.ToString() } else { "" }
            $systemTypesFinalP -notcontains $rtdP
        }
        WriteLog ("Parallel: {0} mailboxes dispatched to {1} threads (filtered from {2} total)." -f $mailboxesForPerm.Count, $ParallelThrottle, $mailboxes.Count) "INFO"

        # Thread-safe result bag
        $permBag = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
        $permCounter = [System.Threading.Interlocked]::Exchange([ref]$counter, 0)

        # Capture variables needed inside parallel scope
        $p_AppId     = $AppId
        $p_Thumb     = $Thumb
        $p_TenantId  = $TenantId
        $p_OrgDomain = $OrgDomain
        $p_Total     = $mailboxesForPerm.Count
        $p_Bag       = $permBag

        $mailboxesForPerm | ForEach-Object -ThrottleLimit $ParallelThrottle -Parallel {
            $mb       = $_
            $bag      = $using:p_Bag
            $appId    = $using:p_AppId
            $thumb    = $using:p_Thumb
            $tenantId = $using:p_TenantId
            $org      = $using:p_OrgDomain

            # Each thread establishes its own EXO session
            $connected = $false
            try {
                Import-Module ExchangeOnlineManagement -ErrorAction Stop
                Connect-ExchangeOnline -AppId $appId -CertificateThumbprint $thumb `
                    -Organization $org -ShowBanner:$false -ErrorAction Stop
                $connected = $true
            } catch {
                # Thread cannot connect - skip this mailbox silently
                return
            }

            try {
                $sendAsVals    = @()
                $fullAccessVals = @()
                $delegationVals = @()

                try {
                    $sendAsVals = Get-EXORecipientPermission -Identity $mb.Identity -ErrorAction SilentlyContinue 2>$null |
                                  Where-Object { $_.Trustee -ne "NT AUTHORITY\SELF" -and $_.AccessRights -contains "SendAs" } |
                                  Select-Object -ExpandProperty Trustee
                } catch {}
                try {
                    $fullAccessVals = Get-EXOMailboxPermission -Identity $mb.Identity -ErrorAction SilentlyContinue 2>$null |
                                      Where-Object { $_.User -ne "NT AUTHORITY\SELF" -and $_.AccessRights -contains "FullAccess" -and -not $_.IsInherited } |
                                      Select-Object -ExpandProperty User
                } catch {}
                try {
                    if ($mb.GrantSendOnBehalfTo) {
                        $delegationVals = $mb.GrantSendOnBehalfTo | ForEach-Object { $_.ToString() } | Where-Object { $_ } | Sort-Object -Unique
                    }
                } catch {}

                $bag.Add([PSCustomObject]@{
                    UserPrincipalName   = $mb.UserPrincipalName
                    PrimarySmtpAddress  = $mb.PrimarySmtpAddress
                    OrganizationalUnit  = $null
                    SendAs              = ($sendAsVals -join ";")
                    FullAccess          = ($fullAccessVals -join ";")
                    GrantSendOnBehalfTo = ($delegationVals -join ";")
                })
            } finally {
                if ($connected) {
                    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}
                }
            }
        }

        $resultsPerm = $permBag.ToArray()
        $counter     = $resultsPerm.Count
        WriteLog ("Parallel permission collection completed: {0} results collected." -f $counter) "INFO"

    } else {
    # -----------------------------------------------------------------------
    # Sequential loop (all modes except PermissionsOnly+Parallel)
    # -----------------------------------------------------------------------
    try { foreach ($mb in $mailboxes) {
        $swMailbox = [System.Diagnostics.Stopwatch]::StartNew()
        $counter++
        if ($counter % 50 -eq 0 -or $IsBoundedMailboxRun) {
            Write-Progress -Activity "Processing mailboxes" -Status "$counter of $total" -PercentComplete (($counter/$total)*100)
        }

        # Transcript progress: log a summary line every $ProgressInterval mailboxes.
        if ($counter % $ProgressInterval -eq 0) {
            $elapsedMin  = [Math]::Round($swRunTotal.Elapsed.TotalMinutes, 1)
            $avgSec      = if ($counter -gt 0) { [Math]::Round($swRunTotal.Elapsed.TotalSeconds / $counter, 2) } else { 0 }
            $remaining   = $total - $counter
            $etaSec      = $avgSec * $remaining
            $etaStr      = if ($etaSec -gt 0) { "{0}h {1}min" -f [Math]::Floor($etaSec/3600), [Math]::Floor(($etaSec%3600)/60) } else { "N/A" }
            WriteLog ("[Progress] {0}/{1} mailboxes processed | Elapsed: {2} min | Avg: {3}s/mbx | ETA: {4}" -f $counter, $total, $elapsedMin, $avgSec, $etaStr) "INFO"
        }

        # MaxRunMinutes guard: stop cleanly and export whatever has been collected so far.
        if ($MaxRunMinutes -gt 0 -and $swRunTotal.Elapsed.TotalMinutes -ge $MaxRunMinutes) {
            WriteLog ("[TIMEOUT] MaxRunMinutes limit reached ({0} min). Stopping after {1}/{2} mailboxes - exporting partial results." -f $MaxRunMinutes, $counter, $total) "WARNING"
            $MaxRunExceeded = $true
            break
        }

        # Proactive reconnect: refresh EXO session if the reconnect interval has elapsed.
        # Only applies to app-only cert auth - interactive auth requires MFA and cannot reconnect silently.
        if (-not $InteractiveAuth -and ((Get-Date) - $script:LastEXOReconnect).TotalMinutes -ge $ReconnectIntervalMinutes) {
            try { Invoke-EXOProactiveReconnect } catch { WriteLog ("Periodic EXO reconnect failed: {0}" -f $_.Exception.Message) "WARNING" }
        }

        # Robust RecipientTypeDetails (used for filtering)
        $rtd = if ($PermissionsOnly) {
            if ($mb.PSObject.Properties.Name -contains 'RecipientTypeDetails' -and $mb.RecipientTypeDetails) { $mb.RecipientTypeDetails.ToString() } else { "" }
        } else {
            Get-RecipientTypeDetailsSafe -MailboxObj $mb
        }

        if ($systemTypesFinal -contains $rtd) { continue }

        try {
            # -------- Permissions (live if -IncludePerm, else snapshot) --------
            $sendAsStr=""; $fullAccessStr=""; $delegationStr=""
            if ($IncludePerm) {
                $sendAsVals=@(); $fullAccessVals=@(); $delegationVals=@()
                try {
                    $sendAsVals = Get-EXORecipientPermission -Identity $mb.Identity -ErrorAction SilentlyContinue 2>$null |
                                  Where-Object { $_.Trustee -ne "NT AUTHORITY\SELF" -and $_.AccessRights -contains "SendAs" } |
                                  Select-Object -ExpandProperty Trustee
                } catch {}
                try {
                    $fullAccessVals = Get-EXOMailboxPermission -Identity $mb.Identity -ErrorAction SilentlyContinue 2>$null |
                                      Where-Object { $_.User -ne "NT AUTHORITY\SELF" -and $_.AccessRights -contains "FullAccess" -and -not $_.IsInherited } |
                                      Select-Object -ExpandProperty User
                } catch {}
                try {
                    if ($mb.GrantSendOnBehalfTo) {
                        $delegationVals = $mb.GrantSendOnBehalfTo |
                                          ForEach-Object { Resolve-RecipientCached $_ } |
                                          Where-Object { $_ } |
                                          Sort-Object -Unique
                    }
                } catch {}
                $sendAsStr=($sendAsVals -join ";")
                $fullAccessStr=($fullAccessVals -join ";")
                $delegationStr=($delegationVals -join ";")

                $ouPerm=$null
                if (-not $PermissionsOnly) {
                    $swOU = [System.Diagnostics.Stopwatch]::StartNew()
                    if ($mb.ExternalDirectoryObjectId){$ouPerm=Get-OU-Fast $mb.ExternalDirectoryObjectId}
                    if (-not $ouPerm -and $mb.UserPrincipalName){$ouPerm=Get-OU-Fast $mb.UserPrincipalName}
                    $swOU.Stop(); Add-PerfSeconds -Key "Loop_OULookup_Seconds" -Seconds $swOU.Elapsed.TotalSeconds
                }

                $resultsPerm += [PSCustomObject]@{
                    UserPrincipalName    = $mb.UserPrincipalName
                    PrimarySmtpAddress   = $mb.PrimarySmtpAddress
                    OrganizationalUnit   = $ouPerm
                    SendAs               = $sendAsStr
                    FullAccess           = $fullAccessStr
                    GrantSendOnBehalfTo  = $delegationStr
                }
            } else {
                $swPerm = [System.Diagnostics.Stopwatch]::StartNew()
                $snap = $null
                if ($PermSnapshot -and $mb.UserPrincipalName) {
                    $pk = Normalize-Key $mb.UserPrincipalName
                    if ($PermSnapshot.ContainsKey($pk)) { $snap = $PermSnapshot[$pk] }
                }
                $swPerm.Stop(); Add-PerfSeconds -Key "Loop_PermissionsLookup_Seconds" -Seconds $swPerm.Elapsed.TotalSeconds
                if ($snap) {
                    $sendAsStr=[string]$snap.SendAs
                    $fullAccessStr=[string]$snap.FullAccess
                    $delegationStr=[string]$snap.GrantSendOnBehalfTo
                }
}

            if ($PermissionsOnly) {
                # Build DEFAULT row with ALL columns empty (same columns/order), then fill only what we have.
                $swObj = [System.Diagnostics.Stopwatch]::StartNew()
                                $properties = [ordered]@{}
                foreach ($col in $columnOrderDefault) { $properties[$col] = "" }

                $properties["DisplayName"]          = if ($mb.DisplayName) { $mb.DisplayName } else { "" }
                $properties["UserPrincipalName"]    = $mb.UserPrincipalName
                $properties["PrimarySmtpAddress"]   = $mb.PrimarySmtpAddress
                $properties["RecipientTypeDetails"] = $rtd
                $properties["MailboxType"]          = $rtd
                $properties["ExternalDirectoryObjectId"] = $mb.ExternalDirectoryObjectId
                $properties["MailboxPlan"]          = if ($mb.PSObject.Properties.Name -contains 'MailboxPlan') { $mb.MailboxPlan } else { "" }

                $properties["SendAs"]              = $sendAsStr
                $properties["FullAccess"]          = $fullAccessStr
                $properties["GrantSendOnBehalfTo"] = $delegationStr
                $properties["ImmutableId"]         = if ($mb.ExternalDirectoryObjectId) {
                                                         try { [Convert]::ToBase64String(([GUID]$mb.ExternalDirectoryObjectId).ToByteArray()) } catch { "" }
                                                     } else { "" }
                $properties["OnPremisesImmutableId"] = ""

                $resultsDetailed += [PSCustomObject]$properties
                $swObj.Stop(); Add-PerfSeconds -Key "Loop_ObjectBuild_Seconds" -Seconds $swObj.Elapsed.TotalSeconds

            $swMailbox.Stop(); $global:PerfMailboxTimes.Add([PSCustomObject]@{ UserPrincipalName=$mb.UserPrincipalName; Seconds=$swMailbox.Elapsed.TotalSeconds }) | Out-Null

                # Build ARCHIVE row with ALL columns empty (same columns/order), fill minimal identity.
                $arch = [ordered]@{}
                foreach ($col in $columnOrderArchive) { $arch[$col] = "" }
                $arch["UserPrincipalName"]  = $mb.UserPrincipalName
                $arch["PrimarySmtpAddress"] = $mb.PrimarySmtpAddress
                $arch["ImmutableId"]        = if ($mb.ExternalDirectoryObjectId) {
                                                  try { [Convert]::ToBase64String(([GUID]$mb.ExternalDirectoryObjectId).ToByteArray()) } catch { "" }
                                              } else { "" }
                $resultsArchiveOnly += [PSCustomObject]$arch

                continue
            }
            # -------- Archive + Primary stats --------
            $swStats = $null
            if (-not $PermissionsOnly) { $swStats = [System.Diagnostics.Stopwatch]::StartNew() }

            # -------- Archive (minimal CSV) --------
            $archiveSizeGB = ""; $archiveCount = ""; $archiveLlt = ""
            $hasArchive = ($mb.PSObject.Properties.Name -contains 'ArchiveStatus' -and $mb.ArchiveStatus -and $mb.ArchiveStatus.ToString() -ne 'None')
            if ($PermissionsOnly) {
                # PermissionsOnly: stats are intentionally blank
            }
            elseif (-not $UseLiveStats) {
                $snapStat = $null
                try {
                    if ($StatsSnapshot -and $mb.PrimarySmtpAddress) {
                        $k = Normalize-Key ($mb.PrimarySmtpAddress.ToString())
                        if ($StatsSnapshot.ContainsKey($k)) { $snapStat = $StatsSnapshot[$k] }
                    }
                } catch { }

                if ($snapStat) {
                    $archiveSizeGB = [string]$snapStat.Archive_TotalItemSizeGB
                    $archiveCount  = [string]$snapStat.Archive_ItemCount
                    $archiveLlt    = [string]$snapStat.Archive_LastLogonTime
                }
            }
            else {
                if ($hasArchive) {
                $exoArch = Get-ExoStatsExact -Identity $mb.UserPrincipalName -Archive -IncludeLastUserActionTime:$IncludeLastUserActionTime.IsPresent
                if ($exoArch) {
                    try {
                        if ($exoArch.PSObject.Properties.Name -contains 'TotalItemSize' -and $exoArch.TotalItemSize) {
                            if ($exoArch.TotalItemSize.PSObject.Properties.Name -contains 'Value' -and $exoArch.TotalItemSize.Value -and ($exoArch.TotalItemSize.Value.PSObject.Methods.Name -contains 'ToBytes')) {
                                $archiveSizeGB = [math]::Round(([double]$exoArch.TotalItemSize.Value.ToBytes())/1GB, 2)
                            } elseif ($exoArch.TotalItemSize.PSObject.Methods.Name -contains 'ToBytes') {
                                $archiveSizeGB = [math]::Round(($exoArch.TotalItemSize.ToBytes())/1GB, 2)
                            } else {
                                $archiveSizeGB = Convert-QuotaToGB $exoArch.TotalItemSize
                            }
                        }
                        elseif ($exoArch.PSObject.Properties.Name -contains 'TotalItemSizeInBytes' -and $exoArch.TotalItemSizeInBytes) {
                            $archiveSizeGB = [math]::Round(($exoArch.TotalItemSizeInBytes)/1GB, 2)
                        }
                    } catch {
                        if ($exoArch.PSObject.Properties.Name -contains 'TotalItemSize') { $archiveSizeGB = Convert-QuotaToGB $exoArch.TotalItemSize }
                    }


                    if ($exoArch.PSObject.Properties.Name -contains 'ItemCount' -and $exoArch.ItemCount -ne $null) { $archiveCount = $exoArch.ItemCount }
                    if ($exoArch.PSObject.Properties.Name -contains 'LastLogonTime' -and $exoArch.LastLogonTime) {
                        $archiveLlt = (Get-Date $exoArch.LastLogonTime -Format "yyyy-MM-dd HH:mm:ss")
                    } else {
                        try {
                            $legacyArch = Get-MailboxStatistics -Identity $mb.UserPrincipalName -Archive -ErrorAction SilentlyContinue 2>$null
                            if ($legacyArch -and $legacyArch.LastLogonTime) { $archiveLlt = (Get-Date $legacyArch.LastLogonTime -Format "yyyy-MM-dd HH:mm:ss") }
                            if ([string]::IsNullOrWhiteSpace($archiveCount) -and $legacyArch -and $legacyArch.PSObject.Properties.Name -contains 'ItemCount') { $archiveCount = $legacyArch.ItemCount }
                            if ([string]::IsNullOrWhiteSpace($archiveSizeGB) -and $legacyArch -and $legacyArch.PSObject.Properties.Name -contains 'TotalItemSize') { $archiveSizeGB = Convert-QuotaToGB $legacyArch.TotalItemSize }

                        } catch {}
                    }
                }
            }
            }
            $ouArch=$null
            $swOU2 = [System.Diagnostics.Stopwatch]::StartNew()
            if ($mb.ExternalDirectoryObjectId){$ouArch=Get-OU-Fast $mb.ExternalDirectoryObjectId}
            if (-not $ouArch -and $mb.UserPrincipalName){$ouArch=Get-OU-Fast $mb.UserPrincipalName}
            $swOU2.Stop(); Add-PerfSeconds -Key "Loop_OULookup_Seconds" -Seconds $swOU2.Elapsed.TotalSeconds
            $resultsArchiveOnly += [PSCustomObject]@{
                UserPrincipalName              = $mb.UserPrincipalName
                PrimarySmtpAddress             = $mb.PrimarySmtpAddress
                OrganizationalUnit             = $ouArch
                Archive_TotalItemSizeGB        = $archiveSizeGB
                Archive_ItemCount              = $archiveCount
                Archive_LastLogonTime          = $archiveLlt
                ImmutableId                    = if ($mb.ExternalDirectoryObjectId) {
                                                     try { [Convert]::ToBase64String(([GUID]$mb.ExternalDirectoryObjectId).ToByteArray()) } catch { "" }
                                                 } else { "" }
            }

            # -------- Primary stats (optional LUAT) --------
            $exoPri = $null
            $legacyPri = $null

            $lastAction  = ""
            $totalSizeGB = ""
            $itemCount   = ""
            $lastLogon   = ""

            if ($PermissionsOnly) {
                # PermissionsOnly: stats are intentionally blank
            }
            elseif (-not $UseLiveStats) {
                $snapStat = $null
                try {
                    if ($StatsSnapshot -and $mb.PrimarySmtpAddress) {
                        $k = Normalize-Key ($mb.PrimarySmtpAddress.ToString())
                        if ($StatsSnapshot.ContainsKey($k)) { $snapStat = $StatsSnapshot[$k] }
                    }
                } catch { }

                if ($snapStat) {
                    $lastAction  = [string]$snapStat.LastUserActionTime
                    $totalSizeGB = [string]$snapStat.TotalItemSizeGB
                    $itemCount   = [string]$snapStat.ItemCount
                    $lastLogon   = [string]$snapStat.LastLogonTime
                }
            }
            else {
                            $exoPri = Get-ExoStatsExact -Identity $mb.UserPrincipalName -IncludeLastUserActionTime:$IncludeLastUserActionTime.IsPresent
                            $legacyPri = $null
                            $lastAction  = ""
                            $totalSizeGB = ""
                            $itemCount   = ""
                            $lastLogon   = ""

                            if ($exoPri) {
                                if ($DebugUPN -and $mb.UserPrincipalName -ieq $DebugUPN) {
                                    WriteLog -Message ("DEBUG EXO: UPN={0} | LUA={1} | LLT={2} | SizeRaw={3} | HasBytes={4}" -f $mb.UserPrincipalName, $exoPri.LastUserActionTime, $exoPri.LastLogonTime, $exoPri.TotalItemSize, ($exoPri.PSObject.Properties.Name -contains 'TotalItemSizeInBytes'))
                                }

                                if ($IncludeLastUserActionTime -and $exoPri.PSObject.Properties.Name -contains 'LastUserActionTime' -and $exoPri.LastUserActionTime) {
                                    $lastAction = (Get-Date $exoPri.LastUserActionTime -Format "yyyy-MM-dd HH:mm:ss")
                                } else {
                                    $lastAction = ""
                                }

                                try {
                                    if ($exoPri.PSObject.Properties.Name -contains 'TotalItemSize' -and $exoPri.TotalItemSize) {
                                        if ($exoPri.TotalItemSize.PSObject.Properties.Name -contains 'Value' -and $exoPri.TotalItemSize.Value -and ($exoPri.TotalItemSize.Value.PSObject.Methods.Name -contains 'ToBytes')) {
                                            $totalSizeGB = [math]::Round(([double]$exoPri.TotalItemSize.Value.ToBytes())/1GB, 2)
                                        }
                                        elseif ($exoPri.TotalItemSize.PSObject.Methods.Name -contains 'ToBytes') {
                                            $totalSizeGB = [math]::Round(([double]$exoPri.TotalItemSize.ToBytes())/1GB, 2)
                                        }
                                        else {
                                            $totalSizeGB = Convert-QuotaToGB $exoPri.TotalItemSize
                                        }
                                    }
                                    elseif ($exoPri.PSObject.Properties.Name -contains 'TotalItemSizeInBytes' -and $exoPri.TotalItemSizeInBytes) {
                                        $totalSizeGB = [math]::Round(([double]$exoPri.TotalItemSizeInBytes)/1GB, 2)
                                    }
                                } catch {
                                    if ($exoPri.PSObject.Properties.Name -contains 'TotalItemSize') { $totalSizeGB = Convert-QuotaToGB $exoPri.TotalItemSize }
                                }

                                if ($exoPri.PSObject.Properties.Name -contains 'ItemCount' -and $exoPri.ItemCount -ne $null) { $itemCount = $exoPri.ItemCount }

                                if ($exoPri.PSObject.Properties.Name -contains 'LastLogonTime' -and $exoPri.LastLogonTime) {
                                    $lastLogon = (Get-Date $exoPri.LastLogonTime -Format "yyyy-MM-dd HH:mm:ss")
                                } else {
                                    try {
                                        $legacyPri = Get-MailboxStatistics -Identity $mb.UserPrincipalName -ErrorAction SilentlyContinue 2>$null
                                        if ($legacyPri -and $legacyPri.PSObject.Properties.Name -contains 'LastLogonTime' -and $legacyPri.LastLogonTime) {
                                            $lastLogon = (Get-Date $legacyPri.LastLogonTime -Format "yyyy-MM-dd HH:mm:ss")
                                        }
                                        if ([string]::IsNullOrWhiteSpace($itemCount) -and $legacyPri -and $legacyPri.PSObject.Properties.Name -contains 'ItemCount') {
                                            $itemCount = $legacyPri.ItemCount
                                        }
                                        if ([string]::IsNullOrWhiteSpace($totalSizeGB) -and $legacyPri -and $legacyPri.PSObject.Properties.Name -contains 'TotalItemSize') {
                                            $totalSizeGB = Convert-QuotaToGB $legacyPri.TotalItemSize
                                        }
                                    } catch {}
                                }
                            }
                            else {
                                try {
                                    $legacyPri = Get-MailboxStatistics -Identity $mb.UserPrincipalName -ErrorAction SilentlyContinue 2>$null
                                    if ($legacyPri) {
                                        if ($legacyPri.PSObject.Properties.Name -contains 'TotalItemSize') { $totalSizeGB = Convert-QuotaToGB $legacyPri.TotalItemSize }
                                        if ($legacyPri.PSObject.Properties.Name -contains 'ItemCount') { $itemCount = $legacyPri.ItemCount }
                                        if ($legacyPri.PSObject.Properties.Name -contains 'LastLogonTime' -and $legacyPri.LastLogonTime) {
                                            $lastLogon = (Get-Date $legacyPri.LastLogonTime -Format "yyyy-MM-dd HH:mm:ss")
                                        }
                                    }
                                } catch {}
                            }

            
            }
            if ($swStats) { $swStats.Stop(); Add-PerfSeconds -Key "Loop_Stats_Seconds" -Seconds $swStats.Elapsed.TotalSeconds }

            # -------- Stats export row (only when live stats collected) --------
            if ($UseLiveStats) {
                $resultsStats += [PSCustomObject]@{
                    Identity                        = $mb.Identity
                    UserPrincipalName               = $mb.UserPrincipalName
                    PrimarySmtpAddress              = $mb.PrimarySmtpAddress
                    StatsSource                     = $statsSourceLabel
                    TotalItemSizeGB                 = $totalSizeGB
                    ItemCount                       = $itemCount
                    LastLogonTime                   = $lastLogon
                    LastUserActionTime              = $lastAction
                    Archive_TotalItemSizeGB         = $archiveSizeGB
                    Archive_ItemCount               = $archiveCount
                    Archive_LastLogonTime           = $archiveLlt
                }
            }

            # -------- User/Graph enrichment --------
            # Local enrichment from CSV snapshots (if loaded)
            $upnKey = ""
            $m365Row = $null
            $adRow   = $null
            $adOu    = ""
            $adEnabled = $null
            $adManager = ""
            try {
                if ($mb.UserPrincipalName) {
                    $upnKey = $mb.UserPrincipalName.ToLower()
                    if ($global:M365UsersCache -and $global:M365UsersCache.ContainsKey($upnKey)) { $m365Row = $global:M365UsersCache[$upnKey] }
                    if ($global:AdUsersUpnCache -and $global:AdUsersUpnCache.ContainsKey($upnKey)) {
                        $adRow = $global:AdUsersUpnCache[$upnKey]
                        $adOu = Get-OUFromDistinguishedName -DistinguishedName ([string]$adRow.DistinguishedName)
                        if ($adRow.PSObject.Properties.Name -contains 'Enabled') { $adEnabled = $adRow.Enabled }
                        if ($adRow.PSObject.Properties.Name -contains 'manager') { $adManager = Resolve-ManagerDisplayNameFromDn -ManagerDn ([string]$adRow.manager) }
                    }
                }
            } catch { }

            $userDetails = $null
            $needGetUser = (-not $adRow) -or [string]::IsNullOrWhiteSpace([string]$adRow.SamAccountName)
            if ($needGetUser) {
                $userDetails = Invoke-ExoSafe -Context ("Get-User:{0}" -f $mb.UserPrincipalName) -ScriptBlock { Get-User -Identity $mb.UserPrincipalName -ErrorAction Stop } 2>$null
            }
            $region = $null
            if ($mb.UserPrincipalName) {
                $doRegional = $false
                switch ($RegionalConfigMode) {
                    "Off"   { $doRegional = $false }
                    "Force" { $doRegional = $true }
                    "Auto"  {
                        # Optimized mode: call only for UserMailbox when Languages is empty (reduces EXO load)
                        if (-not $mb.Languages -and ($rtd -eq "UserMailbox")) { $doRegional = $true }
                    }
                    default {
                        # Legacy mode: preserve previous behavior
                        if (-not $mb.Languages) { $doRegional = $true }
                    }
                }

                if ($doRegional) {

                    $rk = $mb.UserPrincipalName.ToLower()

                    # 1) Empty cache => skip EXO call
                    if ($script:RegionalEmptyCache.Contains($rk)) {
                        $regionalSkippedEmptyCache++
                        $regionalOkEmpty++
                        $region = $null
                    }
                    # 2) Result cache => reuse
                    elseif ($script:RegionalConfigCache.ContainsKey($rk)) {
                        $regionalFromCache++
                        $region = $script:RegionalConfigCache[$rk]
                        if ($null -eq $region) {
                            $regionalOkEmpty++
                        } else {
                            $isEmpty = [string]::IsNullOrWhiteSpace($region.Language) -and [string]::IsNullOrWhiteSpace($region.TimeZone) -and [string]::IsNullOrWhiteSpace($region.DateFormat) -and [string]::IsNullOrWhiteSpace($region.TimeFormat)
                            if ($isEmpty) { $regionalOkEmpty++ } else { $regionalOkWithData++ }
                        }
                    }
                    # 3) Live call
                    else {
                        $regionalAttempted++
                        $swRc = [System.Diagnostics.Stopwatch]::StartNew()
                        $region = Invoke-ExoSafe -Context ("Get-MailboxRegionalConfiguration:{0}" -f $mb.UserPrincipalName) -ScriptBlock { Get-MailboxRegionalConfiguration -Identity $mb.UserPrincipalName -ErrorAction Stop }
                        $swRc.Stop(); Add-PerfSeconds -Key "Loop_RegionalConfig_Seconds" -Seconds $swRc.Elapsed.TotalSeconds

                        if ($null -eq $region) {
                            $regionalFailed++
                            # Do not cache failures
                        } else {
                            $isEmpty = [string]::IsNullOrWhiteSpace($region.Language) -and [string]::IsNullOrWhiteSpace($region.TimeZone) -and [string]::IsNullOrWhiteSpace($region.DateFormat) -and [string]::IsNullOrWhiteSpace($region.TimeFormat)
                            if ($isEmpty) {
                                $regionalOkEmpty++
                                [void]$script:RegionalEmptyCache.Add($rk)
                                $script:RegionalConfigCache[$rk] = $null
                            } else {
                                $regionalOkWithData++
                                $script:RegionalConfigCache[$rk] = $region
                            }
                        }
                    }
                }
}
            $graphData = $null
            if ($needGraphConnect) {
                try {
                    if ($mb.ExternalDirectoryObjectId) {
                        $graphData = Get-GraphUserDetails -UserId $mb.ExternalDirectoryObjectId -ResolveManager:$ResolveManager.IsPresent
                    } elseif ($mb.UserPrincipalName) {
                        $graphData = Get-GraphUserDetails -UserId $mb.UserPrincipalName -ResolveManager:$ResolveManager.IsPresent
                    }
                } catch {
                    Add-ErrorBucket "Graph:Unknown"
                }
            }

# Robust DisplayName
            $displayNameVal = $null
            if ($mb.PSObject.Properties.Name -contains 'DisplayName' -and $mb.DisplayName) { $displayNameVal = $mb.DisplayName }
            elseif ($userDetails -and $userDetails.DisplayName) { $displayNameVal = $userDetails.DisplayName }
            elseif ($m365Row -and $m365Row.'Display name') { $displayNameVal = $m365Row.'Display name' }
            elseif ($adRow -and $adRow.DisplayName) { $displayNameVal = $adRow.DisplayName }
            elseif ($graphData -and $graphData['DisplayName']) { $displayNameVal = $graphData['DisplayName'] }
            else {
                try {
                    $rec = Get-Recipient -Identity $mb.UserPrincipalName -ErrorAction SilentlyContinue 2>$null
                    if ($rec -and $rec.DisplayName) { $displayNameVal = $rec.DisplayName }
                } catch {}
                if (-not $displayNameVal) { $displayNameVal = $mb.UserPrincipalName }
            }

            # Fallback for missing REST fields
            $needFallback = $false
            foreach ($chk in @('EmailAddressPolicyEnabled','MailboxPlan','HiddenFromAddressListsEnabled','RetentionHoldEnabled','LitigationHoldEnabled','RetentionPolicy','ThrottlingPolicy','RoleAssignmentPolicy')) {
                if (-not ($mb.PSObject.Properties.Name -contains $chk)) { $needFallback = $true; break }
            }
            $mailboxFallback=@{}
            if ($needFallback) {
                $mailboxFallback = Get-MailboxFallbackFields -Upn $mb.UserPrincipalName
            }

            # -------- Output line --------
            $swObj = [System.Diagnostics.Stopwatch]::StartNew()
            $properties = [ordered]@{
                DisplayName            = $displayNameVal
                UserPrincipalName      = $mb.UserPrincipalName
                PrimarySmtpAddress     = $mb.PrimarySmtpAddress
                RecipientTypeDetails   = $rtd
                MailboxType            = $rtd
                SamAccountName         = if ($adRow -and $adRow.PSObject.Properties.Name -contains 'SamAccountName' -and $adRow.SamAccountName) { [string]$adRow.SamAccountName }
                                         elseif ($userDetails) { $userDetails.SamAccountName } else { "" }
                MailboxGuid            = if ($mb.PSObject.Properties.Name -contains 'MailboxGuid' -and $mb.MailboxGuid) {
                                             $mb.MailboxGuid
                                         }
                                         elseif ($mb.PSObject.Properties.Name -contains 'ExchangeGuid' -and $mb.ExchangeGuid) {
                                             $mb.ExchangeGuid
                                         } else { "" }
                EmailAddresses         = ($mb.EmailAddresses | Where-Object { $_ }) -join ";"

                Company                = if ($mb.Company) { $mb.Company } elseif ($m365Row -and $m365Row.Company) { $m365Row.Company } elseif ($graphData) { $graphData.Company } else { $null }
                Department             = if ($mb.Department) { $mb.Department } elseif ($m365Row -and $m365Row.Department) { $m365Row.Department } elseif ($adRow -and $adRow.Department) { $adRow.Department } elseif ($graphData) { $graphData.Department } else { $null }
                Title                  = if ($mb.Title) { $mb.Title } elseif ($m365Row -and $m365Row.Title) { $m365Row.Title } elseif ($adRow -and $adRow.Title) { $adRow.Title } elseif ($graphData) { $graphData.Title } else { $null }
                OfficeLocation         = if ($graphData) { $graphData.OfficeLocation } else { $null }
                Manager                = if ($adManager) { $adManager } elseif ($graphData) { $graphData.Manager } else { $null }
                OrganizationalUnit     = if ($adOu) { $adOu } elseif ($graphData) { $graphData.OrganizationalUnit } else { $null }
                UserType               = if ($graphData) { $graphData.UserType } else { $null }
                AccountEnabled         = if ($adEnabled -ne $null -and $adEnabled -ne "") {
                                             $adEnabled
                                         } elseif ($graphData -and $graphData.AccountEnabled -ne $null) {
                                             $graphData.AccountEnabled
                                         } elseif ($userDetails -and $userDetails.Enabled -ne $null) {
                                             $userDetails.Enabled
                                         } else { "" }
                AssignedLicenses       = ""

                ForwardingSmtpAddress  = $mb.ForwardingSmtpAddress
                ForwardingAddress      = $mb.ForwardingAddress
                DeliverToMailboxAndForward = $mb.DeliverToMailboxAndForward
                IsMailboxEnabled       = if ($adEnabled -ne $null -and $adEnabled -ne "") {
                                             $adEnabled
                                         } elseif ($graphData -and $graphData.AccountEnabled -ne $null) {
                                             $graphData.AccountEnabled
                                         } elseif ($userDetails -and $userDetails.Enabled -ne $null) {
                                             $userDetails.Enabled
                                         } else { "" }

                IsEmailAddressPolicyEnabled = if ($mb.PSObject.Properties.Name -contains 'EmailAddressPolicyEnabled') {
                                                  $mb.EmailAddressPolicyEnabled
                                              }
                                              elseif ($mailboxFallback.ContainsKey('EmailAddressPolicyEnabled')) {
                                                  $mailboxFallback['EmailAddressPolicyEnabled']
                                              } else { "" }

                HiddenFromAddressListsEnabled = if ($mb.PSObject.Properties.Name -contains 'HiddenFromAddressListsEnabled') {
                                                    $mb.HiddenFromAddressListsEnabled
                                                }
                                                elseif ($mailboxFallback.ContainsKey('HiddenFromAddressListsEnabled')) {
                                                    $mailboxFallback['HiddenFromAddressListsEnabled']
                                                } else { "" }

                IsDirSynced            = if ($m365Row -and $m365Row.DirSyncEnabled -ne $null -and $m365Row.DirSyncEnabled -ne "") {
                                             $m365Row.DirSyncEnabled
                                         }
                                         elseif ($mb.PSObject.Properties.Name -contains 'IsDirSynced' -and $null -ne $mb.IsDirSynced) {
                                             $mb.IsDirSynced
                                         }
                                         elseif ($graphData -and $graphData['IsDirSynced'] -ne $null) {
                                             $graphData['IsDirSynced']
                                         } else { "" }

                RetentionHoldEnabled   = if ($mb.PSObject.Properties.Name -contains 'RetentionHoldEnabled') {
                                             $mb.RetentionHoldEnabled
                                         }
                                         elseif ($mailboxFallback.ContainsKey('RetentionHoldEnabled')) {
                                             $mailboxFallback['RetentionHoldEnabled']
                                         } else { "" }

                LitigationHoldEnabled  = if ($mb.PSObject.Properties.Name -contains 'LitigationHoldEnabled') {
                                             $mb.LitigationHoldEnabled
                                         }
                                         elseif ($mailboxFallback.ContainsKey('LitigationHoldEnabled')) {
                                             $mailboxFallback['LitigationHoldEnabled']
                                         } else { "" }

                RetentionPolicy        = if ($mb.PSObject.Properties.Name -contains 'RetentionPolicy') {
                                             $mb.RetentionPolicy
                                         }
                                         elseif ($mailboxFallback.ContainsKey('RetentionPolicy')) {
                                             $mailboxFallback['RetentionPolicy']
                                         } else { "" }

                ThrottlingPolicy       = if ($mb.PSObject.Properties.Name -contains 'ThrottlingPolicy') {
                                             $mb.ThrottlingPolicy
                                         }
                                         elseif ($mailboxFallback.ContainsKey('ThrottlingPolicy')) {
                                             $mailboxFallback['ThrottlingPolicy']
                                         } else { "" }

                RoleAssignmentPolicy   = if ($mb.PSObject.Properties.Name -contains 'RoleAssignmentPolicy') {
                                             $mb.RoleAssignmentPolicy
                                         }
                                         elseif ($mailboxFallback.ContainsKey('RoleAssignmentPolicy')) {
                                             $mailboxFallback['RoleAssignmentPolicy']
                                         } else { "" }

                UseDatabaseQuotaDefaults   = $mb.UseDatabaseQuotaDefaults
                IssueWarningQuotaGB        = Convert-QuotaToGB $mb.IssueWarningQuota
                ProhibitSendQuotaGB        = Convert-QuotaToGB $mb.ProhibitSendQuota
                ProhibitSendReceiveQuotaGB = Convert-QuotaToGB $mb.ProhibitSendReceiveQuota

                WhenMailboxCreated     = $mb.WhenMailboxCreated
                LastExchangeChangedTime= if ($mb.PSObject.Properties.Name -contains 'LastExchangeChangedTime') { $mb.LastExchangeChangedTime } else { "" }
                Languages              = if ($mb.Languages) {
                                             $mb.Languages -join ";"
                                         } elseif ($region -and $region.Language) {
                                             $region.Language.Name
                                         } else { "" }
                LastUserActionTime     = $lastAction
                TotalItemSizeGB        = $totalSizeGB
                ItemCount              = $itemCount
                LastLogonTime          = $lastLogon

                SendAs                 = $sendAsStr
                FullAccess             = $fullAccessStr
                GrantSendOnBehalfTo    = $delegationStr

                ArchiveStatus          = $mb.ArchiveStatus
                ArchiveName            = $mb.ArchiveName
                ArchiveQuotaGB         = Convert-QuotaToGB $mb.ArchiveQuota
                ArchiveWarningQuotaGB  = Convert-QuotaToGB $mb.ArchiveWarningQuota

                CustomAttribute1  = $mb.CustomAttribute1
                CustomAttribute2  = $mb.CustomAttribute2
                CustomAttribute3  = $mb.CustomAttribute3
                CustomAttribute4  = $mb.CustomAttribute4
                CustomAttribute5  = $mb.CustomAttribute5
                CustomAttribute6  = $mb.CustomAttribute6
                CustomAttribute7  = $mb.CustomAttribute7
                CustomAttribute8  = $mb.CustomAttribute8
                CustomAttribute9  = $mb.CustomAttribute9
                CustomAttribute10 = $mb.CustomAttribute10
                CustomAttribute11 = $mb.CustomAttribute11
                CustomAttribute12 = $mb.CustomAttribute12
                CustomAttribute13 = $mb.CustomAttribute13
                CustomAttribute14 = $mb.CustomAttribute14
                CustomAttribute15 = $mb.CustomAttribute15

                ExternalDirectoryObjectId = $mb.ExternalDirectoryObjectId
                MailboxPlan               = if ($mb.PSObject.Properties.Name -contains 'MailboxPlan' -and $mb.MailboxPlan) {
                                                 $mb.MailboxPlan
                                             }
                                             elseif ($mailboxFallback.ContainsKey('MailboxPlan')) {
                                                 $mailboxFallback['MailboxPlan']
                                             }
                                             else { $null }

                ImmutableId                      = if ($mb.ExternalDirectoryObjectId) {
                                                       try { [Convert]::ToBase64String(([GUID]$mb.ExternalDirectoryObjectId).ToByteArray()) } catch { "" }
                                                   } else { "" }
                OnPremisesImmutableId            = if ($graphData -and $graphData['OnPremisesImmutableId']) { $graphData['OnPremisesImmutableId'] } else { "" }
            }

            $expectedKeys=@("Department","Title","Company","OrganizationalUnit","UserType","AccountEnabled","AssignedLicenses","Manager","OfficeLocation","IsDirSynced","OnPremisesImmutableId")
            foreach($key in $expectedKeys){
                if (-not $properties.Contains($key)) { $properties[$key] = "" }
            }

            $resultsDetailed += [PSCustomObject]$properties
            $swObj.Stop(); Add-PerfSeconds -Key "Loop_ObjectBuild_Seconds" -Seconds $swObj.Elapsed.TotalSeconds
            try { $swMailbox.Stop(); $global:PerfMailboxTimes.Add([PSCustomObject]@{ UserPrincipalName=$mb.UserPrincipalName; Seconds=$swMailbox.Elapsed.TotalSeconds }) | Out-Null } catch {}

        } catch {
            WriteLog "Error processing mailbox $($mb.UserPrincipalName): $($_.Exception.Message)" "ERROR"
            $errors += [PSCustomObject]@{
                UserPrincipalName = $mb.UserPrincipalName
                Error             = $_.Exception.Message
            }
            Add-ErrorBucket "MailboxProcessing"
            try { $swMailbox.Stop(); $global:PerfMailboxTimes.Add([PSCustomObject]@{ UserPrincipalName=$mb.UserPrincipalName; Seconds=$swMailbox.Elapsed.TotalSeconds }) | Out-Null } catch {}
            continue
        }
    } } catch [System.Management.Automation.PipelineStoppedException] {
        $ManualStop = $true
        WriteLog ("[STOP] Script interrupted manually after {0}/{1} mailboxes - exporting partial results." -f $counter, $total) "WARNING"
    } catch {
        $ManualStop = $true
        WriteLog ("[STOP] Unexpected termination after {0}/{1} mailboxes: {2} - exporting partial results." -f $counter, $total, $_.Exception.Message) "ERROR"
    }
    } # end else (sequential loop)

    $swStageLoop.Stop(); Add-PerfSeconds -Key "Stage_MailboxLoop_Seconds" -Seconds $swStageLoop.Elapsed.TotalSeconds

    #region Export
    if ($MaxRunExceeded) {
        WriteLog ("[TIMEOUT] Partial export: {0}/{1} mailboxes collected." -f $resultsDetailed.Count, $total) "WARNING"
    }
    if ($ManualStop -and -not $MaxRunExceeded) {
        WriteLog ("[STOP] Partial export: {0}/{1} mailboxes collected." -f $resultsDetailed.Count, $total) "WARNING"
    }
    # Append _PARTIAL to all CSV filenames when the run did not complete normally.
    if ($MaxRunExceeded -or $ManualStop) { $CsvSuffix += "_PARTIAL" }
    $swStageExport = [System.Diagnostics.Stopwatch]::StartNew()
    $exportData = $resultsDetailed | Select-Object -Property $columnOrderDefault
    $exportDataArchive = $resultsArchiveOnly | Select-Object -Property $columnOrderArchive
    $columnOrderStats = @(
        'Identity','UserPrincipalName','PrimarySmtpAddress','StatsSource',
        'TotalItemSizeGB','ItemCount','LastLogonTime','LastUserActionTime',
        'Archive_TotalItemSizeGB','Archive_ItemCount','Archive_LastLogonTime'
    )
    $exportDataStats = @()
    if ($UseLiveStats) {
        $exportDataStats = @($resultsStats | Select-Object -Property $columnOrderStats)
        $statsCompletenessIssues = New-Object System.Collections.Generic.List[object]
        $expectedStatsRows = @($exportData).Count
        $actualStatsRows = @($exportDataStats).Count
        if ($actualStatsRows -ne $expectedStatsRows) {
            $statsCompletenessIssues.Add([PSCustomObject]@{
                RowNumber = ''
                UserPrincipalName = ''
                PrimarySmtpAddress = ''
                Identity = ''
                MissingFields = 'StatsRowCount'
                StatsSource = ''
                Details = ("Expected {0} stats rows but collected {1}." -f $expectedStatsRows, $actualStatsRows)
            }) | Out-Null
        }
        $statsRowNumber = 0
        foreach ($statsRow in $exportDataStats) {
            $statsRowNumber++
            $missingFields = @()
            foreach ($fieldName in @('TotalItemSizeGB','ItemCount')) {
                if ([string]::IsNullOrWhiteSpace([string]$statsRow.$fieldName)) {
                    $missingFields += $fieldName
                }
            }
            if ($missingFields.Count -gt 0) {
                $statsCompletenessIssues.Add([PSCustomObject]@{
                    RowNumber = $statsRowNumber
                    UserPrincipalName = $statsRow.UserPrincipalName
                    PrimarySmtpAddress = $statsRow.PrimarySmtpAddress
                    Identity = $statsRow.Identity
                    MissingFields = ($missingFields -join ';')
                    StatsSource = $statsRow.StatsSource
                    Details = 'Critical live mailbox statistics fields are missing.'
                }) | Out-Null
            }
        }
        if ($statsCompletenessIssues.Count -gt 0) {
            Add-ErrorBucket "StatsCompleteness"
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $statsDiagnosticPath = Join-Path $OutputPath ("Exchange_EXO_Mailboxes_AllDomains_Stats_Incomplete_{0}.csv" -f $timestamp)
            $statsCompletenessIssues | Export-Csv -Path $statsDiagnosticPath -NoTypeInformation -Encoding UTF8
            $issueRows = foreach ($issue in $statsCompletenessIssues) { $issue }
            $script:StatsCompletenessDiagnosticPath = $statsDiagnosticPath
            $script:StatsCompletenessIssueRows = @($issueRows)
            Publish-MailboxInventoryDiagnosticCsv -Path $statsDiagnosticPath
            $issuePercent = if ($expectedStatsRows -gt 0) { [math]::Round((100.0 * $statsCompletenessIssues.Count / $expectedStatsRows), 4) } else { 100 }
            $isBlockingStatsCompleteness = (($StatsCompletenessFailMinRows -gt 0 -and $statsCompletenessIssues.Count -ge $StatsCompletenessFailMinRows) -or ($StatsCompletenessFailPercent -gt 0 -and $issuePercent -ge $StatsCompletenessFailPercent))
            $thresholdText = ("fail thresholds: rows>={0}, percent>={1}%" -f $StatsCompletenessFailMinRows, $StatsCompletenessFailPercent)
            if ($isBlockingStatsCompleteness) {
                WriteLog -Message ("Live stats completeness validation failed for {0}/{1} issue row(s), {2}%. {3}. Diagnostic exported to: {4}. DATA-LAST and SharePoint CSV publication are skipped." -f $statsCompletenessIssues.Count, $expectedStatsRows, $issuePercent, $thresholdText, $statsDiagnosticPath) "ERROR"
                throw ("Live mailbox statistics export is incomplete for {0}/{1} issue row(s), {2}%. DATA-LAST CSVs were not updated. Diagnostic: {3}" -f $statsCompletenessIssues.Count, $expectedStatsRows, $issuePercent, $statsDiagnosticPath)
            }
            WriteLog -Message ("Live stats completeness warning for {0}/{1} issue row(s), {2}%. {3}. Diagnostic exported to: {4}. DATA-LAST publication continues." -f $statsCompletenessIssues.Count, $expectedStatsRows, $issuePercent, $thresholdText, $statsDiagnosticPath) "WARNING"
        }
    }

    $globalPathForStdExports = $GlobalExportPath
    $globalPathForPermExport = $GlobalExportPath

    Write-Host "`n--- Export CSV (DEFAULT) ---"
    if ($PermissionsOnly) {
        WriteLog "CSV DEFAULT skipped: -PermissionsOnly mode, no mailbox data collected." "INFO"
        Write-Host "CSV DEFAULT skipped (-PermissionsOnly). See log for details." -ForegroundColor Yellow
    } else {
        ExportAndCopyCsv -BaseFileName "Exchange_EXO_Mailboxes_AllDomains$CsvSuffix" `
           -OutputPath $OutputPath `
           -GlobalPath $globalPathForStdExports `
           -Data $exportData `
           -Encoding "UTF8" `
           -NoTypeInformation
    }

    Write-Host "`n--- Export CSV (ARCHIVE) ---"
    if ($PermissionsOnly) {
        WriteLog "CSV ARCHIVE skipped: -PermissionsOnly mode, no mailbox data collected." "INFO"
        Write-Host "CSV ARCHIVE skipped (-PermissionsOnly). See log for details." -ForegroundColor Yellow
    } else {
        ExportAndCopyCsv -BaseFileName "Exchange_EXO_Mailboxes_AllDomains_Archive$CsvSuffix" `
           -OutputPath $OutputPath `
           -GlobalPath $globalPathForStdExports `
           -Data $exportDataArchive `
           -Encoding "UTF8" `
           -NoTypeInformation
    }

    Write-Host "`n--- Export CSV (STATS) ---"
    if (-not $UseLiveStats) {
        WriteLog -Message "CSV STATS not exported: -IncludeStats was not specified. Snapshot/default mode completed as configured." "INFO"
        Write-Host "CSV STATS skipped (no -IncludeStats)." -ForegroundColor DarkGray
    } else {
        ExportAndCopyCsv -BaseFileName "Exchange_EXO_Mailboxes_AllDomains_Stats$CsvSuffix" `
           -OutputPath $OutputPath `
           -GlobalPath $globalPathForStdExports `
           -Data $exportDataStats `
           -Encoding "UTF8" `
           -NoTypeInformation
    }

    if ($IncludePerm) {
        $columnOrderPerm = @('UserPrincipalName','PrimarySmtpAddress','OrganizationalUnit','SendAs','FullAccess','GrantSendOnBehalfTo')
        $exportDataPerm = $resultsPerm | Select-Object -Property $columnOrderPerm
        Write-Host "`n--- Export CSV (PERMISSIONS) ---"
        ExportAndCopyCsv -BaseFileName "Exchange_EXO_Mailboxes_AllDomains_Permissions$CsvSuffix" `
           -OutputPath $OutputPath `
           -GlobalPath $globalPathForPermExport `
           -Data $exportDataPerm `
           -Encoding "UTF8" `
           -NoTypeInformation
    }

    if ($errors.Count -gt 0) {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $errorPath  = Join-Path $OutputPath ("Exchange_EXO_Mailboxes_AllDomains${CsvSuffix}_Errors.csv")
        $errorPath2 = Join-Path $OutputPath ("Exchange_EXO_Mailboxes_AllDomains${CsvSuffix}_Errors_${timestamp}.csv")
        $errors | Export-Csv -Path $errorPath  -NoTypeInformation -Encoding UTF8
        Write-Host "Error log exported to: $errorPath"
        $errors | Export-Csv -Path $errorPath2 -NoTypeInformation -Encoding UTF8
        Write-Host "Error log exported to: $errorPath2"
    }

    $swStageExport.Stop(); Add-PerfSeconds -Key "Stage_Export_Seconds" -Seconds $swStageExport.Elapsed.TotalSeconds

    Write-Host "`n--- Execution Summary ---"
    Write-Host "Mailboxes processed : $total"
        Write-Host ("RegionalConfig calls  : {0} | WithData={1} | Empty={2} | Failed={3} | CacheHit={4} | EmptyCacheSkip={5}" -f $regionalAttempted,$regionalOkWithData,$regionalOkEmpty,$regionalFailed,$regionalFromCache,$regionalSkippedEmptyCache)
    Write-Host "Log file             : $logTextFile"
    Write-Host "-------------------------`n"
    #endregion

    # ================= NOC-style Error Summary =================
    Write-Host "`n--- Error Summary (NOC) ---"
    if ($errors.Count -gt 0) {
        Write-Host "By error message (top buckets):"
        $grouped = $errors | Group-Object -Property Error | Sort-Object Count -Descending | Select-Object -First 10
        foreach ($g in $grouped) {
            Write-Host ("  x{0}  {1}" -f $g.Count, $g.Name)
        }
    } else {
        Write-Host "No per-mailbox processing errors captured."
    }

    if ($global:GraphHttpCodeCounts.Keys.Count -gt 0) {
        Write-Host "Graph HTTP codes:"
        foreach ($k in $global:GraphHttpCodeCounts.Keys) {
            Write-Host ("  {0} : {1}" -f $k, $global:GraphHttpCodeCounts[$k])
        }
    } else {
        Write-Host "No Graph HTTP errors recorded."
    }

    if ($global:ErrorBuckets.Keys.Count -gt 0) {
        Write-Host "Error buckets:"
        foreach ($k in $global:ErrorBuckets.Keys) {
            Write-Host ("  {0} : {1}" -f $k, $global:ErrorBuckets[$k])
        }
    }
    Write-Host "------------------------------`n"
    # =================================================================

    # ================= Performance Summary =================
    try {
        $swTotalRun.Stop()
        Add-PerfSeconds -Key "Total_Run_Seconds" -Seconds $swTotalRun.Elapsed.TotalSeconds

        Write-Host "`n--- Performance Summary ---"
        foreach ($k in $global:PerfCounters.Keys) {
                        $sec = [double]$global:PerfCounters[$k]
            $ms  = [int][math]::Round($sec * 1000, 0)
            Write-Host ("{0,-32} : {1,10:N3} sec  ({2,8} ms)" -f $k, $sec, $ms)
        }

        if ($global:PerfMailboxTimes -and $global:PerfMailboxTimes.Count -gt 0) {
            $top = $global:PerfMailboxTimes | Sort-Object Seconds -Descending | Select-Object -First 5
            Write-Host "`nTop 5 slowest mailboxes:"
            foreach ($t in $top) {
                Write-Host ("  {0,-55} {1,8:N2} sec" -f $t.UserPrincipalName, $t.Seconds)
            }

            $avg = ($global:PerfMailboxTimes | Measure-Object -Property Seconds -Average).Average
            Write-Host ("`nAvg per-mailbox processing: {0:N3} sec" -f [double]$avg)
        }

        Write-Host "---------------------------`n"
    } catch {
        WriteLog -Message ("Failed to produce performance summary: {0}" -f $_.Exception.Message) "WARNING"
    }
    # =======================================================

    #region Disconnect + Cleanup
    $swStageCleanup = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Host "`n--- Disconnect Cloud Services ---"
    Disconnect-SmartM365CloudSession -ExchangeOnline $true -Graph $true
    WriteLog -Message "$TaskName completed."
    try { Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {} } catch {}
    Complete-SmartM365ExecutionContext -Status Auto
    RemoveOldFiles -Path $OutputPath -Filter "*.csv" -KeepCount $global:RetentionMaxCSV -LogFile $global:logTextFile
    RemoveOldFiles -Path $logPath -Filter "*.log" -KeepCount $global:RetentionMaxLogs -LogFile $global:logTextFile
    $swStageCleanup.Stop(); Add-PerfSeconds -Key "Stage_DisconnectCleanup_Seconds" -Seconds $swStageCleanup.Elapsed.TotalSeconds
    #endregion
}
catch {
    $globalError = $_
    WriteLog -Message ("Global error in Exchange Online mailboxes inventory: {0}" -f $globalError) "ERROR"
    Write-Host "A global error occurred. Check the log file for details." -ForegroundColor Red

    # -------- Global error email via SmartM365.Core --------
    try {
        $title = "Exchange Online mailboxes inventory - ERROR"
        $msg   = @"
An error occurred in script $($MyInvocation.MyCommand.Name) on $(Get-Date -Format "yyyy-MM-dd HH:mm:ss").

Error message:
$($globalError.Exception.Message)

See attached log file for details:
$($global:LogTextFile)
"@

        $bodyHtml = NewSimpleEmailBody -Title $title -Message $msg
        if ($script:StatsCompletenessIssueRows -and $script:StatsCompletenessIssueRows.Count -gt 0) {
            $bodyHtml = Add-MailboxInventoryHtmlSection -BodyHtml $bodyHtml -SectionHtml (New-StatsCompletenessIssuesHtml -Issues $script:StatsCompletenessIssueRows -MaxRows 50)
        }

        $attachments = @()
        if ($global:LogTextFile -and (Test-Path $global:LogTextFile)) {
            $attachments += $global:LogTextFile
        }
        if ($script:StatsCompletenessDiagnosticPath -and (Test-Path -LiteralPath $script:StatsCompletenessDiagnosticPath -PathType Leaf)) {
            $attachments += $script:StatsCompletenessDiagnosticPath
        }

        Send-FatalErrorEmail -Subject $title -HtmlBody $bodyHtml -Attachments $attachments
    } catch {
        WriteLog -Message ("Failed to send global error email: {0}" -f $_) "ERROR"
    }

    try {
        Disconnect-SmartM365CloudSession -ExchangeOnline $true -Graph $true
    } catch {}

    try { Stop-Transcript | Out-Null; try { $smartM365TranscriptPath = $null; $smartM365TranscriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } else { $smartM365TranscriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue; if ($smartM365TranscriptVariable -and $smartM365TranscriptVariable.Value) { $smartM365TranscriptPath = $smartM365TranscriptVariable.Value } }; if ($smartM365TranscriptPath) { Update-SmartM365TimestampedTranscript -Path $smartM365TranscriptPath } } catch {} } catch {}
    Complete-SmartM365ExecutionContext -Status Auto
    exit 1
}
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBHQ/9W6PQKqZLs
# fhFbanc70/kI0VArz8T9kaY6BvX3n6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGf4q3QapdRLLWYPBQKC30Xlh9q+d4dvB1WD3KDSNy+HMA0GCSqG
# SIb3DQEBAQUABIIBgJCHTiQNuM8IfdKhYM+xGkCjyon3LmuAMBedDzfW3yHD2m0r
# 3Ect/tdZJzsmBMFTI9kNkZyNn2txmaEQ0PwFMP9sAmkAbYQq8u4zRcFBfoEOgUNB
# zgV0n7wSlnA54UMhDgF8AsDZc8BLnXbU7mSRTmwpHRRF5+LntKIF43RlNOLv6Rzh
# nwjCndRwHdErqttPRznyUxzM0ln+Sb+P4hjlzKRmKx/FyGkAY0gVtCY7cNifmLUt
# pHhi9+Xa+6G84+sMS+mkLB4U13RBpC6Lw95HXjLUWzPCXBbfM07n2IRUXoVZLkzU
# mAqrab13mW4A1bQvJpdPffohlog2F2zgfLw6NKgzlJ4fAYvTPCbTJz7Y1lDkh89R
# touaWTk5RaWQ6uHrBQK92rx1Ws2PVvYHqJ/IOLOxmxMTtqclb1D43JvVmKoazePc
# sQlybETAFpmGF8YGbLEgHpP81s+6V5pQVcLkYG6Abb1NPH1O6f/WQl9wgNxETDMM
# X05QvEDCQnqzf6FS6qGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODUw
# MjRaMC8GCSqGSIb3DQEJBDEiBCDklHFyoZDOHzyKrhwdvHC3shTTJ+1tZd393iEy
# a3fJ3TANBgkqhkiG9w0BAQEFAASCAgArVX+R5WbU4RdTaHUumJ0FyVmc2cRck6QL
# v1O5k64ePdxQtcatJYBZ/Gg7Qn7h4r6oktHnhAa2tx8zGsTrKwkV4i59Wld+gFwY
# MPNcLlABPor1Wgdrsk9KRQ/K6P+31NZ8tH3j0eSxCNQobwwyEPGAy2ZshsH4DeQ+
# pisoO69QYohzjteHhzzgc2q7CEQg41c0+K+HFvnx7bwcL4Z2xY5+saZ4CM8Su+gm
# ontoa82K03/x4SDYrZZN/I4+0tUsD9lY1uvBGp7b7sal5fwG5LD7rK0qT9KDp3Wj
# pe2AYC+si64pIy5ix4v1vF4a+i0ZwlebcHMLLRJwv6O1T/jSm+jSGlCczCUis73D
# oXItuYTXJ+R75mSn3BaOARJ9BOUf6HXzijck52GwajBBsGGeWXguCWTe5W61QVdd
# SscQpMrSJB8QQ6Ff+RIuKhI+Ro2s6rIplIRhfJNlWG/NuNAGl1/Tn82UU8Lf2yY8
# iLn4byNuAzX6NzI9SYMl9uBncKR4usayUae3fdiAfMRwIi0OlW0qrj7I6ZR0BETi
# RSBtaffEaO0kHrm2VR2b4foWDm2AA25hlhMruKtbxfcflpwkZ1xKlDUg7SIt1i5p
# qVagMeSkEBRCL9ldR+borBibxcEIcamkK4HistlHZl58yk/92G8hgEPKDrQJ6/xC
# veDZpb9VJQ==
# SIG # End signature block
