#Requires -Version 7.0
<#
.SYNOPSIS
    SharePoint Online tenant inventory with CSV exports and HTML alert summary.

.DESCRIPTION
    Uses Microsoft Graph app-only authentication by default to inventory
    SharePoint Online sites, storage, activity, lists where Graph allows access,
    owner signals, inactive sites, orphaned sites, and tenant-level summary
    statistics. Tenant capacity is also queried through PnP.PowerShell when
    available; deep sharing scans remain an explicit optional feature.

.PARAMETER Tenant
    Tenant profile key to load from Config/Tenants. Defaults to test.

.PARAMETER IncludeOneDrive
    Includes OneDrive personal sites in the tenant site collection inventory.

.PARAMETER InactiveDays
    Number of days without activity before a site is flagged as inactive. Defaults to 180.

.PARAMETER MaxSites
    Limits the number of sites processed. Intended for test runs.

.PARAMETER DryRun
    Exports CSVs and writes console/log output only. No email is sent.

.PARAMETER AlwaysSend
    Bypasses the daily mail marker while retaining the exclusive send lock.

.PARAMETER MailOnly
    Sends the HTML summary from existing latest CSVs without collecting inventory.

.PARAMETER AppendHistory
    Appends rows to per-entity history CSV files in addition to the normal timestamped/latest exports.

.PARAMETER QuotaCriticalPercent
    Storage quota usage percentage that raises a Critical alert. Defaults to 90.

.PARAMETER ListItemWarningThreshold
    List/library item count that raises a Warning alert. Defaults to 5000.

.PARAMETER SkipDeepSharingScan
    Skips the optional deep per-site sharing link scan when -UsePnPDeepScan is enabled.

.PARAMETER UsePnPDeepScan
    Enables optional PnP.PowerShell per-site sharing scans. This is best-effort and may require broader SharePoint permissions than the default Graph inventory.

.PARAMETER UsePnPTenantCapacity
    Requests the licensed SharePoint tenant storage capacity through Get-PnPTenant.
    This is also attempted by default when PnP.PowerShell is installed.
    SharePointAdminUrl is derived automatically when omitted.

.PARAMETER SkipPnPTenantCapacity
    Skips the optional Get-PnPTenant call and uses the license-based mail estimate.

.PARAMETER InteractiveAuth
    Uses delegated interactive Graph authentication instead of app-only certificate authentication.

.VERSION
 0.44
.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication; ImportExcel. PnP.PowerShell enables tenant capacity collection and optional sharing scans.
    Minimum Graph application permissions for default inventory: Reports.Read.All; Sites.Read.All; Directory.Read.All.
    Optional PnP deep sharing scan may require SharePoint site-level access for scanned sites.
    Conditional: Mail.Send is required only when Graph mail is used; Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Requires: PowerShell 7+, Microsoft.Graph.Authentication, ImportExcel, SmartM365.Core.psd1
    PnP.PowerShell is used for capacity when installed, unless -SkipPnPTenantCapacity is set.
    Minimum Microsoft Graph application permissions for default inventory:
      - Reports.Read.All for SharePoint and OneDrive usage reports.
      - Sites.Read.All for site and list inventory.
      - Directory.Read.All for organization context and owner/directory enrichment.
    Optional PnP deep sharing scan may require SharePoint site-level access for the scanned sites.
    Optional unattended tenant capacity collection requires the SharePoint API application permission Sites.FullControl.All with administrator consent.
    Microsoft Graph Sites.Read.All does not authorize Get-PnPTenant.
    The Graph site/list inventory does not require SharePoint Administrator or Sites.FullControl.All.
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [switch]$IncludeOneDrive,
    [int]$InactiveDays = 180,
    [int]$MaxSites = 0,
    [switch]$DryRun,
    [switch]$AlwaysSend,
    [switch]$MailOnly,
    [switch]$AppendHistory,
    [double]$QuotaCriticalPercent = 90,
    [int]$ListItemWarningThreshold = 5000,
    [switch]$SkipDeepSharingScan,
    [switch]$UsePnPDeepScan,
    [switch]$UsePnPTenantCapacity,
    [switch]$SkipPnPTenantCapacity,
    [int]$SharingScanItemLimitPerSite = 0,
    [switch]$InteractiveAuth,
    [string]$OutputPath,
    [int]$MaxItems = 0
)
# Local lifecycle only; no collection or publication is performed by this guard.
$smartM365LifecycleRoot=$PSScriptRoot
while ($smartM365LifecycleRoot -and -not (Test-Path -LiteralPath (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1') -PathType Leaf)) {
    $smartM365LifecycleParent=Split-Path $smartM365LifecycleRoot -Parent
    if($smartM365LifecycleParent -eq $smartM365LifecycleRoot){break}
    $smartM365LifecycleRoot=$smartM365LifecycleParent
}
if(-not $smartM365LifecycleRoot){throw 'SmartM365 lifecycle support is missing; update the complete runtime.'}
. (Join-Path $smartM365LifecycleRoot 'Modules/SmartM365.Core/SmartM365-InventoryLifecycle.ps1')
$smartM365RunGuard=Start-SmartM365InventoryRun -ScriptPath $PSCommandPath -RequireReturn
try {
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

Set-StrictMode -Version Latest
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::InvariantCulture
[System.Threading.Thread]::CurrentThread.CurrentUICulture = [System.Globalization.CultureInfo]::InvariantCulture
$ErrorActionPreference = 'Stop'
$MaximumFunctionCount = 32768
$ScriptVersion = "0.44"
$TenantCapacityEnabled = -not [bool]$SkipPnPTenantCapacity -and ([bool]$UsePnPTenantCapacity -or [bool](Get-Module -ListAvailable PnP.PowerShell | Select-Object -First 1))
$CurrentOperation = 'Initialize'

$tenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        $candidates = @(
            (Join-Path -Path $d -ChildPath 'SmartM365-TenantContext.ps1'),
            (Join-Path -Path $d -ChildPath 'Config\SmartM365-TenantContext.ps1')
        )
        foreach ($p in $candidates) { if (Test-Path -LiteralPath $p) { return $p } }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $tenantContextPath
$script:SmartM365GlobalConfig = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host 'This script requires PowerShell 7 or later.' -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; $smartM365RunGuard.ExitCode=(1); exit $smartM365RunGuard.ExitCode
}

function Import-SmartM365CoreModule {
    $searchRoot = $PSScriptRoot
    while ($searchRoot) {
        $candidate = Join-Path -Path $searchRoot -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
        if (Test-Path -LiteralPath $candidate) { Import-Module -Name $candidate -MinimumVersion '1.0.95' -Force -ErrorAction Stop; return }
        $parent = Split-Path -Path $searchRoot -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $searchRoot) { break }
        $searchRoot = $parent
    }
    throw 'SmartM365.Core module manifest not found.'
}

function Get-SmartM365GlobalConfig { if ($null -ne $script:SmartM365GlobalConfig) { return $script:SmartM365GlobalConfig }; $script:SmartM365GlobalConfig = [pscustomobject]@{}; return $script:SmartM365GlobalConfig }

function Resolve-SmartM365ConfigValue {
    param([AllowNull()]$Value)
    if ($Value -isnot [string] -or [string]::IsNullOrWhiteSpace($Value) -or $Value -notmatch '\{\{[^}]+\}\}') { return $Value }
    $globalConfig = Get-SmartM365GlobalConfig
    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $tokenMatches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($tokenMatches.Count -eq 0) { break }
        $changed = $false
        foreach ($match in $tokenMatches) {
            $tokenProperty = $globalConfig.PSObject.Properties[$match.Groups['Name'].Value]
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

function Get-ScriptLocalConfig {
    $configPath = Join-Path -Path $PSScriptRoot -ChildPath ("{0}.local.json" -f [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)); $configPath = Resolve-SmartM365JsonConfigurationPath -Path $configPath
    if (-not (Test-Path -LiteralPath $configPath)) {
        $templatePath = (Get-SmartM365JsonTemplateName -Path $configPath)
        if (Get-Command Initialize-SmartM365LocalJsonFromTemplate -ErrorAction SilentlyContinue) { Initialize-SmartM365LocalJsonFromTemplate -Path $configPath -TemplatePath $templatePath -ConfigDescription 'script local configuration' | Out-Null }
        else { if (-not (Test-Path -LiteralPath $templatePath)) { throw "Missing local config and template: $configPath" }; Write-SmartM365JsonBytesAtomically -Path $configPath -Bytes ([IO.File]::ReadAllBytes($templatePath)) -ExpectedSHA256 'ABSENT' -Validate {param($document) if($document -isnot [pscustomobject]){throw 'Configuration template must be an object.'}} | Out-Null }
    }
    $config = Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
    return (Sync-SmartM365JsonConfigWithTemplate -Config $config -Path $configPath -TemplatePath (Get-SmartM365JsonTemplateName -Path $configPath))
}

function Get-ScriptLocalConfigValue {
    param([Parameter(Mandatory)]$Config,[Parameter(Mandatory)][string]$Name,$DefaultValue)
    $property = $Config.PSObject.Properties[$Name]
    if ($null -ne $property -and $null -ne $property.Value) {
        if ($property.Value -is [string]) { $localValue = $property.Value.Trim(); if ($localValue -and $localValue -notin @('__USE_GLOBAL__','USE_GLOBAL')) { return Resolve-SmartM365ConfigValue -Value $property.Value } }
        else { return Resolve-SmartM365ConfigValue -Value $property.Value }
    }
    $globalConfig = Get-SmartM365GlobalConfig
    $globalProperty = $globalConfig.PSObject.Properties[$Name]
    if ($null -ne $globalProperty -and $null -ne $globalProperty.Value) { if ($globalProperty.Value -is [string] -and [string]::IsNullOrWhiteSpace($globalProperty.Value)) { return $DefaultValue }; return Resolve-SmartM365ConfigValue -Value $globalProperty.Value }
    return $DefaultValue
}

function Write-SpoLog { param([Parameter(Mandatory)][string]$Message,[ValidateSet('INFO','SUCCESS','WARNING','ERROR')][string]$Level='INFO') if (Get-Command WriteLog -ErrorAction SilentlyContinue) { WriteLog -Message $Message -Level $Level; return }; Write-Host "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message" }
function ConvertTo-SpoIsoUtc { param([AllowNull()]$Value) if ($null -eq $Value) { return '' }; try { $date=[datetime]$Value; if($date -eq [datetime]::MinValue){return ''}; return $date.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ',[System.Globalization.CultureInfo]::InvariantCulture) } catch { return '' } }
function ConvertTo-SpoDouble { param([AllowNull()]$Value) if ($null -eq $Value) { return $null }; try { return [double]::Parse(([string]$Value),[System.Globalization.CultureInfo]::InvariantCulture) } catch { try { return [double]$Value } catch { return $null } } }
function Get-SpoPropertyValue { param([AllowNull()]$Object,[Parameter(Mandatory)][string[]]$Names,[AllowNull()]$DefaultValue=$null) if ($null -eq $Object) { return $DefaultValue }; foreach($name in $Names){ if($Object -is [System.Collections.IDictionary] -and $Object.Contains($name)){return $Object[$name]}; $property=$Object.PSObject.Properties[$name]; if($null -ne $property){return $property.Value} }; return $DefaultValue }
function ConvertTo-SpoText { param([AllowNull()]$Value) if ($null -eq $Value) { return '' }; if ($Value -is [array]) { return (@($Value) | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_ }) -join ';' }; return [string]$Value }
function ConvertTo-SpoHtml { param([AllowNull()]$Value) if ($null -eq $Value) { return '' }; return [System.Net.WebUtility]::HtmlEncode([string]$Value) }

function Invoke-SpoDailySummaryMail {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$MarkerPath,
        [Parameter(Mandatory)][scriptblock]$SendAction,
        [switch]$ForceSend
    )

    $today = (Get-Date).ToString('yyyy-MM-dd', [System.Globalization.CultureInfo]::InvariantCulture)
    $markerParent = Split-Path -Path $MarkerPath -Parent
    if (-not (Test-Path -LiteralPath $markerParent -PathType Container)) {
        New-Item -Path $markerParent -ItemType Directory -Force | Out-Null
    }

    $lockPath = "$MarkerPath.lock"
    $lockStream = $null
    try {
        try {
            $lockStream = [System.IO.File]::Open($lockPath, [System.IO.FileMode]::OpenOrCreate, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        }
        catch [System.IO.IOException] {
            Write-SpoLog -Message "Daily SharePoint summary email is already being evaluated by another run: $lockPath" -Level INFO
            return $false
        }

        $lastSentDate = if (Test-Path -LiteralPath $MarkerPath -PathType Leaf) {
            [string](Get-Content -LiteralPath $MarkerPath -Raw -ErrorAction SilentlyContinue)
        }
        else { '' }
        if ($lastSentDate.Trim() -eq $today -and -not $ForceSend) {
            Write-SpoLog -Message "Daily SharePoint summary email already sent for $today; email skipped." -Level INFO
            return $false
        }

        $null = & $SendAction
        [System.IO.File]::WriteAllText($MarkerPath, $today, [System.Text.UTF8Encoding]::new($false))
        return $true
    }
    finally {
        if ($null -ne $lockStream) { $lockStream.Dispose() }
    }
}

function Invoke-SpoWithRetry {
    param([Parameter(Mandatory)][scriptblock]$ScriptBlock,[string]$Operation='operation',[int]$MaxAttempts=8,[int]$BaseDelaySeconds=3,[int]$MaxDelaySeconds=90)
    for($attempt=1;$attempt -le $MaxAttempts;$attempt++){
        try { return & $ScriptBlock } catch {
            $message=$_.Exception.Message; $status=''; try { if($_.Exception.Response -and $_.Exception.Response.StatusCode){$status=[string][int]$_.Exception.Response.StatusCode} } catch {}
            if(-not $status){try{$status=[string][int]$_.Exception.Data['StatusCode']}catch{}}
            $isTransient=($status -in @('408','409','429','500','502','503','504')) -or ($message -match '(?i)throttl|too many requests|temporar|timeout|timed out|503|429')
            if(-not $isTransient -or $attempt -ge $MaxAttempts){throw}
            $retryAfter=$null; try{$retryAfter=@($_.Exception.Response.Headers.GetValues('Retry-After')|Select-Object -First 1)[0]}catch{}
            if($null-eq$retryAfter){try{$retryAfter=$_.Exception.Data['Retry-After']}catch{}}
            $retrySeconds=0
            $delay=if($null-ne$retryAfter -and [int]::TryParse([string]$retryAfter,[ref]$retrySeconds) -and $retrySeconds -gt 0){[Math]::Min($MaxDelaySeconds,$retrySeconds)}else{[Math]::Min($MaxDelaySeconds,[int]($BaseDelaySeconds*[Math]::Pow(2,($attempt-1))))+(Get-Random -Minimum 0 -Maximum 4)}
            Write-SpoLog -Message ("Transient failure on {0}. Status={1}; attempt {2}/{3}; waiting {4}s. {5}" -f $Operation,$(if($status){$status}else{'unknown'}),$attempt,$MaxAttempts,$delay,$message) -Level WARNING
            Start-Sleep -Seconds $delay
        }
    }
}

function Add-SpoHistoryCsv { param([AllowEmptyCollection()][object[]]$Data,[Parameter(Mandatory)][string]$Path,[Parameter(Mandatory)][string[]]$Columns) $Columns=@('TenantKey','OrganizationKey','EnvironmentKey','TenantId')+@($Columns|Where-Object{$_-inotmatch'^(TenantKey|OrganizationKey|EnvironmentKey|TenantId)$'}); $parent=Split-Path -Path $Path -Parent; if(-not(Test-Path -LiteralPath $parent)){New-Item -Path $parent -ItemType Directory -Force|Out-Null}; $rows=@($Data|Select-Object -Property $Columns); if(Test-Path -LiteralPath $Path){Repair-SmartM365CsvTenantKeySchema -Path $Path -Delimiter ',' -Encoding UTF8|Out-Null}; Add-SmartM365CsvRowsAtomically -Data $rows -Path $Path -Columns $Columns -Encoding utf8BOM -Delimiter ',' }
function Export-SpoEntityCsv {
    param(
        [Parameter(Mandatory)][string]$BaseFileName,
        [AllowEmptyCollection()][object[]]$Data,
        [Parameter(Mandatory)][string[]]$Columns,
        [Parameter(Mandatory)][string]$TimestampedFolder,
        [Parameter(Mandatory)][string]$LatestFolder,
        [Parameter(Mandatory)][string]$Timestamp,
        [switch]$AppendHistoryMode,
        [switch]$NoWeeklyHistory
    )
    $Columns = @('TenantKey','OrganizationKey','EnvironmentKey','TenantId') + @($Columns | Where-Object { $_ -inotmatch '^(TenantKey|OrganizationKey|EnvironmentKey|TenantId)$' })
    $timestampedPath = Join-Path -Path $TimestampedFolder -ChildPath ("{0}_{1}.csv" -f $BaseFileName,$Timestamp)
    $latestPath = Join-Path -Path $LatestFolder -ChildPath ("{0}.csv" -f $BaseFileName)
    if (@($Data).Count -eq 0) {
        Write-SmartM365CsvAtomically -Data @() -Path $timestampedPath -Columns $Columns -Encoding utf8BOM
        # A staged run writes its DATA-LAST copy to source staging; it is promoted with the receipt.
        $latestWritePath = if (Get-Command -Name Resolve-SmartM365SourcePublicationPath -ErrorAction SilentlyContinue) { Resolve-SmartM365SourcePublicationPath -Path $latestPath } else { $latestPath }
        Write-SmartM365CsvAtomically -Data @() -Path $latestWritePath -Columns $Columns -Encoding utf8BOM
        if (-not $DryRun) {
            $timestampedUpload = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $timestampedPath
            if ($latestWritePath -cne $latestPath) { Add-SmartM365StagedPublicationAction -Path $latestPath -SharePointUpload }
            else { Invoke-SmartM365SharePointCsvUpload -LocalFilePath $latestPath | Out-Null }
            # Non-empty exports go through Export-SmartM365Csv, which applies the same retention.
            if ($timestampedUpload) { Remove-SmartM365SharePointTimestampedCsvOlderThan -TimestampedPath $timestampedPath -RetentionDays 7 | Out-Null }
        }
        return [pscustomobject]@{ TimestampedPath=$timestampedPath; LatestPath=$latestPath; PublishedPath=$latestWritePath }
    }
    # Final encoding must precede uploads, weekly snapshots and receipt hashing.
    $result = Export-SmartM365Csv -Data $Data -TimestampedPath $timestampedPath -LatestPath $latestPath -Columns $Columns -Encoding utf8BOM -NoWeeklyHistory:$NoWeeklyHistory
    if ($AppendHistoryMode) {
        $historyPath = Join-Path -Path $TimestampedFolder -ChildPath ("{0}_History.csv" -f $BaseFileName)
        Add-SpoHistoryCsv -Data $Data -Path $historyPath -Columns $Columns
        Write-SpoLog -Message ("History CSV appended: {0}" -f $historyPath)
    }
    return $result
}
function Get-SpoCsvColumnNames {
    param([Parameter(Mandatory)][string]$Path)
    $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path)
    try {
        $parser.TextFieldType = [Microsoft.VisualBasic.FileIO.FieldType]::Delimited
        $parser.SetDelimiters(',')
        $parser.HasFieldsEnclosedInQuotes = $true
        if ($parser.EndOfData) { return @() }
        return @($parser.ReadFields())
    }
    finally { $parser.Dispose() }
}

function Get-SpoLicenseCapacityEstimate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][string]$ExpectedTenantKey,
        [Parameter(Mandatory)][AllowEmptyString()][string]$ExpectedTenantId,
        [Parameter(Mandatory)][string]$ReferenceDateUtc
    )

    $unavailable = [pscustomobject]@{Available=$false;CapacityGB='';EligibleUnits=0;AsOfUtc='';Reason='A current, matching license snapshot is unavailable.'}
    $path = Join-Path $Folder 'M365_Licenses_Tenant.csv'
    if ([string]::IsNullOrWhiteSpace($ExpectedTenantId) -or -not (Test-Path -LiteralPath $path -PathType Leaf)) { return $unavailable }
    try {
        $required = @('TenantKey','TenantId','TenantSkuPartNumber','TenantPrepaidEnabled','CapabilityStatus','AppliesTo','CollectedAtUtc')
        $headers = @(Get-SpoCsvColumnNames -Path $path)
        if (@($required | Where-Object { $headers -notcontains $_ }).Count -gt 0) { return $unavailable }
        $beforeHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        $rows = @(Import-Csv -LiteralPath $path -ErrorAction Stop)
        $afterHash = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        if ($beforeHash -ne $afterHash -or $rows.Count -eq 0) { return $unavailable }
        if (@($rows | Where-Object { [string]$_.TenantKey -ine $ExpectedTenantKey -or [string]$_.TenantId -ine $ExpectedTenantId }).Count -gt 0) { return $unavailable }
        $dates = @($rows | ForEach-Object { [string]$_.CollectedAtUtc } | Sort-Object -Unique)
        if ($dates.Count -ne 1) { return $unavailable }
        $collected = [datetimeoffset]::MinValue
        $reference = [datetimeoffset]::MinValue
        if (-not [datetimeoffset]::TryParse($dates[0],[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AssumeUniversal,[ref]$collected) -or
            -not [datetimeoffset]::TryParse($ReferenceDateUtc,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AssumeUniversal,[ref]$reference)) { return $unavailable }
        if ([math]::Abs(($collected - $reference).TotalHours) -gt 48 -or ($collected - [datetimeoffset]::UtcNow).TotalMinutes -gt 10) { return $unavailable }

        # Only known SharePoint-entitled purchased seats contribute to this conservative estimate.
        $eligibleSkus = @('SPE_E3','SPE_E5','ENTERPRISEPACK','ENTERPRISEPREMIUM','STANDARDPACK','SHAREPOINTSTANDARD','SHAREPOINTENTERPRISE','SPB','O365_BUSINESS_PREMIUM','O365_BUSINESS_ESSENTIALS')
        $eligibleUnits = [long]0
        foreach ($row in $rows) {
            if ([string]$row.TenantSkuPartNumber -notin $eligibleSkus -or [string]$row.CapabilityStatus -ine 'Enabled' -or [string]$row.AppliesTo -ine 'User') { continue }
            $units = [long]0
            if (-not [long]::TryParse([string]$row.TenantPrepaidEnabled,[Globalization.NumberStyles]::Integer,[Globalization.CultureInfo]::InvariantCulture,[ref]$units) -or $units -lt 0) { return $unavailable }
            $eligibleUnits += $units
        }
        if ($eligibleUnits -le 0) { return $unavailable }
        return [pscustomobject]@{Available=$true;CapacityGB=(1024 + 10 * $eligibleUnits);EligibleUnits=$eligibleUnits;AsOfUtc=$collected.UtcDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ',[Globalization.CultureInfo]::InvariantCulture);Reason='Base allocation plus known eligible purchased licenses; additional storage is excluded.'}
    }
    catch { return $unavailable }
}

function Add-SpoLicenseEstimateToSummary {
    param(
        [Parameter(Mandatory)][System.Collections.IDictionary]$Summary,
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][string]$ExpectedTenantKey,
        [Parameter(Mandatory)][AllowEmptyString()][string]$ExpectedTenantId,
        [Parameter(Mandatory)][string]$ReferenceDateUtc,
        [bool]$PartialInventory = $false
    )
    $Summary['EstimatedTenantCapacityGB'] = ''
    $Summary['EstimatedTenantUtilizationPercent'] = ''
    $Summary['CapacityEstimateEligibleUnits'] = 0
    $Summary['CapacityEstimateAsOfUtc'] = ''
    if (-not [string]::IsNullOrWhiteSpace([string]$Summary.TenantStorageCapacityGB)) { return }
    $estimate = Get-SpoLicenseCapacityEstimate -Folder $Folder -ExpectedTenantKey $ExpectedTenantKey -ExpectedTenantId $ExpectedTenantId -ReferenceDateUtc $ReferenceDateUtc
    if (-not $estimate.Available) {
        Write-SpoLog -Message $estimate.Reason -Level WARNING
        return
    }
    $Summary['EstimatedTenantCapacityGB'] = $estimate.CapacityGB
    $Summary['CapacityEstimateEligibleUnits'] = $estimate.EligibleUnits
    $Summary['CapacityEstimateAsOfUtc'] = $estimate.AsOfUtc
    if (-not $PartialInventory -and $Summary.Contains('StorageUsedGB')) {
        $usedGB = [double]0
        if ([double]::TryParse([string]$Summary.StorageUsedGB,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$usedGB) -and
            [double]::IsFinite($usedGB) -and $usedGB -ge 0 -and [double]$estimate.CapacityGB -gt 0) {
            $Summary['EstimatedTenantUtilizationPercent'] = [math]::Round(100 * $usedGB / [double]$estimate.CapacityGB,2)
        }
    }
}

function ConvertTo-SpoSiteUrlKey {
    param([AllowEmptyString()][string]$Url)
    if ([string]::IsNullOrWhiteSpace($Url)) { return '' }
    $uri = $null
    if (-not [uri]::TryCreate($Url,[UriKind]::Absolute,[ref]$uri) -or $uri.Scheme -ne 'https') { return '' }
    return $uri.GetLeftPart([UriPartial]::Path).TrimEnd('/').ToLowerInvariant()
}

function Measure-SpoTeamsSiteLinks {
    param(
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$SiteRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$TeamRows,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$ChannelRows,
        [Parameter(Mandatory)][string]$TeamsRunDateUtc
    )

    $sites = @($SiteRows | Where-Object { [string]$_.IsOneDrive -ne 'True' })
    if ($sites.Count -eq 0) { throw 'No SharePoint sites are available for the Teams site ratio.' }
    if (@($sites | Where-Object { -not $_.SiteId -or [string]$_.SiteIdentitySource -ne 'GraphSite' -or [string]$_.GraphEnrichmentStatus -ne 'Collected' }).Count -gt 0) {
        throw 'Some SPO sites lack a verified Graph site identity.'
    }
    $byId = @{}
    $byUrl = @{}
    foreach ($site in $sites) {
        $urlKey = ConvertTo-SpoSiteUrlKey -Url ([string]$site.SiteUrl)
        if (-not $urlKey -or $byUrl.ContainsKey($urlKey)) { throw 'SPO site URLs are missing or duplicated.' }
        $byUrl[$urlKey] = $site
        $idKey = ([string]$site.SiteId).Trim().ToLowerInvariant()
        if ($idKey) {
            if ($byId.ContainsKey($idKey)) { throw 'SPO Graph site IDs are duplicated.' }
            $byId[$idKey] = $site
        }
    }
    $matched = @{}
    $unmatchedParent = 0
    $unmatchedChannel = 0
    $links = [System.Collections.Generic.List[object]]::new()
    foreach ($team in $TeamRows) {
        [void]$links.Add([pscustomobject]@{SiteId=[string]$team.SharePointSiteId;SiteUrl=[string]$team.SharePointSiteUrl;LinkType='Team site';TeamId=[string]$team.TeamId;TeamDisplayName=[string]$team.TeamDisplayName;ChannelId='';ChannelDisplayName=''})
    }
    $teamById = @{}
    foreach ($team in $TeamRows) { $teamById[[string]$team.TeamId] = $team }
    foreach ($channel in $ChannelRows) {
        if ([string]$channel.MembershipType -notin @('private','shared')) { continue }
        [void]$links.Add([pscustomobject]@{SiteId=[string]$channel.SharePointSiteId;SiteUrl='';LinkType=([string]$channel.MembershipType + ' channel');TeamId=[string]$channel.TeamId;TeamDisplayName=[string]$teamById[[string]$channel.TeamId].TeamDisplayName;ChannelId=[string]$channel.ChannelId;ChannelDisplayName=[string]$channel.ChannelDisplayName})
    }
    foreach ($link in $links) {
        $idKey = $link.SiteId.Trim().ToLowerInvariant()
        $urlKey = ConvertTo-SpoSiteUrlKey -Url $link.SiteUrl
        $siteById = if ($idKey -and $byId.ContainsKey($idKey)) { $byId[$idKey] } else { $null }
        $siteByUrl = if ($urlKey -and $byUrl.ContainsKey($urlKey)) { $byUrl[$urlKey] } else { $null }
        if ($siteById -and $siteByUrl -and [string]$siteById.SiteUrl -ine [string]$siteByUrl.SiteUrl) { throw 'Teams site ID and URL identify different SPO sites.' }
        $site = if ($siteById) { $siteById } else { $siteByUrl }
        if (-not $site) {
            if ($link.LinkType -eq 'Team site') { $unmatchedParent++ } else { $unmatchedChannel++ }
            continue
        }
        $matchKey = ConvertTo-SpoSiteUrlKey -Url ([string]$site.SiteUrl)
        if ($matched.ContainsKey($matchKey)) { throw 'One SPO site is linked to multiple Teams resources.' }
        $matched[$matchKey] = [pscustomobject]@{SiteUrl=[string]$site.SiteUrl;SiteId=[string]$site.SiteId;LinkType=$link.LinkType;TeamId=$link.TeamId;TeamDisplayName=$link.TeamDisplayName;ChannelId=$link.ChannelId;ChannelDisplayName=$link.ChannelDisplayName;TeamsRunDateUtc=$TeamsRunDateUtc}
    }
    $rows = @($matched.Values | Sort-Object SiteUrl)
    $parentCount = @($rows | Where-Object { $_.LinkType -eq 'Team site' }).Count
    $channelCount = $rows.Count - $parentCount
    return [pscustomobject]@{
        Available=$true;Reason='';TeamsRunDateUtc=$TeamsRunDateUtc;TotalSites=$sites.Count;LinkedSites=$rows.Count
        ParentSites=$parentCount;ChannelSites=$channelCount;Percent=[math]::Round(100 * $rows.Count / $sites.Count,2)
        UnmatchedParentSites=$unmatchedParent;UnmatchedChannelSites=$unmatchedChannel;Rows=$rows
    }
}

function Get-SpoTeamsSiteLinksFromCsv {
    param(
        [Parameter(Mandatory)][string]$Folder,
        [Parameter(Mandatory)][string]$ExpectedTenantKey,
        [Parameter(Mandatory)][AllowEmptyString()][string]$ExpectedTenantId,
        [Parameter(Mandatory)][string]$ReferenceDateUtc,
        [Parameter(Mandatory)][AllowEmptyCollection()][object[]]$SiteRows,
        [bool]$PartialInventory = $false
    )
    $unavailable = [pscustomobject]@{Available=$false;Reason='Teams site-link evidence is unavailable.';TeamsRunDateUtc='';TotalSites=0;LinkedSites=0;ParentSites=0;ChannelSites=0;Percent='';UnmatchedParentSites=0;UnmatchedChannelSites=0;Rows=@()}
    if ($PartialInventory) { $unavailable.Reason = 'The SPO inventory is limited.'; return $unavailable }
    if (-not $ExpectedTenantId) { $unavailable.Reason = 'Tenant identity is unavailable.'; return $unavailable }
    $teamPath = Join-Path $Folder 'M365_Teams_Teams.csv'
    $channelPath = Join-Path $Folder 'M365_Teams_Channels.csv'
    if (-not (Test-Path -LiteralPath $teamPath -PathType Leaf) -or -not (Test-Path -LiteralPath $channelPath -PathType Leaf)) { $unavailable.Reason = 'Teams site-link CSVs are missing.'; return $unavailable }
    try {
        $teamHeaders = @(Get-SpoCsvColumnNames -Path $teamPath)
        $channelHeaders = @(Get-SpoCsvColumnNames -Path $channelPath)
        if (@(@('TenantKey','TenantId','RunId','RunDateUtc','TeamId','SharePointSiteId','SharePointSiteUrl','SharePointSiteLinkStatus','IsPartialInventory') | Where-Object { $teamHeaders -notcontains $_ }).Count -gt 0 -or
            @(@('TenantKey','TenantId','RunId','RunDateUtc','TeamId','ChannelId','MembershipType','SharePointSiteId','SharePointSiteLinkStatus') | Where-Object { $channelHeaders -notcontains $_ }).Count -gt 0) {
            $unavailable.Reason = 'Teams site-link columns are missing; run the updated Teams inventory.'
            return $unavailable
        }
        $beforeTeam = (Get-FileHash -LiteralPath $teamPath -Algorithm SHA256).Hash
        $beforeChannel = (Get-FileHash -LiteralPath $channelPath -Algorithm SHA256).Hash
        $teams = @(Import-Csv -LiteralPath $teamPath -ErrorAction Stop)
        $channels = @(Import-Csv -LiteralPath $channelPath -ErrorAction Stop)
        if ($beforeTeam -ne (Get-FileHash -LiteralPath $teamPath -Algorithm SHA256).Hash -or
            $beforeChannel -ne (Get-FileHash -LiteralPath $channelPath -Algorithm SHA256).Hash) { throw 'Teams site-link CSVs changed while being read.' }
        if ($teams.Count -eq 0) { throw 'Teams inventory has no team rows.' }
        $sourceRunId = [string]$teams[0].RunId
        $sourceRunDate = [string]$teams[0].RunDateUtc
        if (-not $sourceRunId -or -not $sourceRunDate -or
            @($teams | Where-Object { [string]$_.TenantKey -ine $ExpectedTenantKey -or [string]$_.TenantId -ine $ExpectedTenantId -or [string]$_.RunId -ne $sourceRunId -or [string]$_.RunDateUtc -ne $sourceRunDate -or [string]$_.IsPartialInventory -ne 'False' }).Count -gt 0 -or
            @($channels | Where-Object { [string]$_.TenantKey -ine $ExpectedTenantKey -or [string]$_.TenantId -ine $ExpectedTenantId -or [string]$_.RunId -ne $sourceRunId -or [string]$_.RunDateUtc -ne $sourceRunDate }).Count -gt 0) { throw 'Teams site-link CSVs have different tenant, run, or scope identities.' }
        if ($channels.Count -eq 0 -and [math]::Abs(((Get-Item -LiteralPath $channelPath).LastWriteTimeUtc - (Get-Item -LiteralPath $teamPath).LastWriteTimeUtc).TotalMinutes) -gt 30) { throw 'Empty Teams channel CSV is not from the team run.' }
        $sourceDate = [datetimeoffset]::MinValue
        $referenceDate = [datetimeoffset]::MinValue
        if (-not [datetimeoffset]::TryParse($sourceRunDate,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AssumeUniversal,[ref]$sourceDate) -or
            -not [datetimeoffset]::TryParse($ReferenceDateUtc,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AssumeUniversal,[ref]$referenceDate) -or
            $sourceDate -gt $referenceDate -or ($referenceDate - $sourceDate).TotalHours -gt 24) { throw 'Teams site-link snapshot is newer than or more than 24 hours older than SPO.' }
        if (@($teams | Where-Object { -not $_.TeamId -or [string]$_.SharePointSiteLinkStatus -ne 'Collected' -or -not $_.SharePointSiteId }).Count -gt 0 -or
            @($teams.TeamId | Sort-Object -Unique).Count -ne $teams.Count) { throw 'Teams primary site identities are incomplete or duplicated.' }
        $teamIds = @{}
        foreach ($team in $teams) { $teamIds[[string]$team.TeamId] = $true }
        $channelKeys = @{}
        foreach ($channel in $channels) {
            $key = '{0}|{1}' -f [string]$channel.TeamId,[string]$channel.ChannelId
            if (-not $channel.TeamId -or -not $channel.ChannelId -or -not $teamIds.ContainsKey([string]$channel.TeamId) -or $channelKeys.ContainsKey($key)) { throw 'Teams channel identities are incomplete or duplicated.' }
            $channelKeys[$key] = $true
            $type = [string]$channel.MembershipType
            if ($type -notin @('','standard','private','shared')) { throw 'Teams channel type is unknown.' }
            if ($type -in @('private','shared') -and ([string]$channel.SharePointSiteLinkStatus -ne 'Collected' -or -not $channel.SharePointSiteId)) { throw 'A private or shared channel site identity is unavailable.' }
        }
        return Measure-SpoTeamsSiteLinks -SiteRows $SiteRows -TeamRows $teams -ChannelRows $channels -TeamsRunDateUtc $sourceRunDate
    }
    catch {
        $unavailable.Reason = $_.Exception.Message
        Write-SpoLog -Message ("Teams site-link ratio unavailable: {0}" -f $unavailable.Reason) -Level WARNING
        return $unavailable
    }
}

function Import-SpoMailCsvSet {
    param([Parameter(Mandatory)][string]$Folder,[Parameter(Mandatory)][string]$ExpectedTenantKey)
    $names=@('Sites','Lists','Permissions','ExternalSharing','Tenant')
    $sets=@{}; $files=New-Object 'System.Collections.Generic.List[object]'
    $sourceRunId=''; $sourceRunDate=''; $sourceTenantId=''; $anchorWriteTime=$null
    foreach($name in $names){
        $path=Join-Path $Folder "M365_SPO_$name.csv"
        if(-not(Test-Path -LiteralPath $path -PathType Leaf)){throw "MailOnly requires the latest CSV: $path"}
        $headers=@(Get-SpoCsvColumnNames -Path $path)
        $requiredColumns=@('TenantKey','RunId','RunDateUtc')
        if($name -eq 'Sites'){$requiredColumns+=@('SiteUrl','IsOneDrive','IsInactive','IsOrphaned','StorageQuotaPercent','Status','Details','Threshold','GraphEnrichmentStatus','LastActivityUtc')}
        if($name -eq 'Lists'){$requiredColumns+=@('SiteUrl','ListTitle','ItemCount','Status','Threshold','Details')}
        if($name -eq 'Tenant'){$requiredColumns+=@('IsPartialInventory','StorageCapacityGB','StorageUtilizationPercent','CapacitySource')}
        foreach($required in $requiredColumns){if($headers -notcontains $required){throw "MailOnly CSV $path is missing $required."}}
        $beforeHash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        $rows=@(Import-Csv -LiteralPath $path -ErrorAction Stop)
        $afterHash=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
        if($beforeHash -ne $afterHash){throw "MailOnly CSV changed while being read: $path"}
        $sets[$name]=$rows
        $file=Get-Item -LiteralPath $path
        if($name -eq 'Tenant'){
            if($rows.Count -ne 1){throw 'MailOnly requires exactly one SPO tenant row to establish the source run.'}
            $anchorWriteTime=$file.LastWriteTimeUtc
        }
        $files.Add([pscustomobject]@{Path=$path;WorksheetName=$name;TableName="Spo$name";LastWriteTimeUtc=$file.LastWriteTimeUtc;Empty=($rows.Count -eq 0);Sha256=$afterHash})
        foreach($row in $rows){
            $key=[string]$row.TenantKey; $id=[string]$row.RunId; $date=[string]$row.RunDateUtc
            if([string]::IsNullOrWhiteSpace($key) -or $key -ine $ExpectedTenantKey -or [string]::IsNullOrWhiteSpace($id) -or [string]::IsNullOrWhiteSpace($date)){throw "MailOnly CSV $path contains a row with missing or unexpected tenant/run identity."}
            if(-not $sourceRunId){$sourceRunId=$id; $sourceRunDate=$date; $sourceTenantId=[string]$row.TenantId}
            if($id -ne $sourceRunId -or $date -ne $sourceRunDate -or ([string]$row.TenantId -ne $sourceTenantId)){throw "MailOnly CSV files do not belong to one source run: $path"}
        }
    }
    foreach($file in $files){
        if($file.Empty -and [math]::Abs(($file.LastWriteTimeUtc-$anchorWriteTime).TotalMinutes) -gt 30){throw "MailOnly empty CSV is not timestamp-consistent with the tenant CSV: $($file.Path)"}
    }
    $parsedDate=[datetimeoffset]::MinValue
    if(-not [datetimeoffset]::TryParse($sourceRunDate,[Globalization.CultureInfo]::InvariantCulture,[Globalization.DateTimeStyles]::AssumeUniversal,[ref]$parsedDate)){throw "MailOnly source RunDateUtc is invalid: $sourceRunDate"}
    [pscustomobject]@{Rows=$sets;Files=$files.ToArray();RunId=$sourceRunId;RunDateUtc=$sourceRunDate;Started=$parsedDate.UtcDateTime}
}

function Ensure-SpoImportExcelModule {
    [CmdletBinding()]
    param()
    Initialize-SmartM365ImportExcel -InstallIfMissing | Out-Null
}
function New-SpoTimestampedWorkbook {
    param(
        [Parameter(Mandatory)][object[]]$CsvFiles,
        [Parameter(Mandatory)][string]$Path
    )

    Initialize-SmartM365ImportExcel -InstallIfMissing | Out-Null
    if (Test-Path -LiteralPath $Path) { Remove-Item -LiteralPath $Path -Force }
    foreach ($csv in $CsvFiles) {
        $rows = @(Import-Csv -LiteralPath $csv.Path)
        $isEmpty = $rows.Count -eq 0
        if ($isEmpty) {
            $placeholder = [ordered]@{}
            foreach ($column in @(Get-SpoCsvColumnNames -Path $csv.Path)) { $placeholder[$column] = '' }
            $rows = @([pscustomobject]$placeholder)
        }
        $rows | Export-Excel -Path $Path -WorksheetName $csv.WorksheetName -TableName $csv.TableName -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter
        if ($isEmpty) {
            $package = Open-ExcelPackage -Path $Path
            try { $package.Workbook.Worksheets[$csv.WorksheetName].DeleteRow(2) }
            finally { Close-ExcelPackage -ExcelPackage $package }
        }
    }
    return $Path
}
function Get-SpoMailStatus {
    param([Parameter(Mandatory)][System.Collections.IDictionary]$Summary,[bool]$PartialInventory,[double]$CapacityWarningPercent=80)
    $total=[int]$Summary.SitesProcessed
    $capacity=[string]$Summary.TenantStorageUtilizationPercent
    if (-not $PartialInventory -and -not [string]::IsNullOrWhiteSpace($capacity) -and [double]$capacity -ge 95) { return 'Critical' }
    if ($total -gt 0 -and -not $PartialInventory -and (([double]$Summary.OwnerlessSites / $total) -ge 0.1 -or ([double]$Summary.HighQuotaSites / $total) -ge 0.1)) { return 'Critical' }
    if (-not [string]::IsNullOrWhiteSpace($capacity) -and [double]$capacity -ge $CapacityWarningPercent) { return 'Warning' }
    if (($Summary.CriticalAlerts + $Summary.WarningAlerts) -gt 0) { return 'Warning' }
    return 'OK'
}
function Add-SpoWorkbookOverview {
    param([Parameter(Mandatory)][string]$Path,[Parameter(Mandatory)][System.Collections.IDictionary]$Summary,[Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Alerts,[Parameter(Mandatory)][string]$Status)
    $teamsLinks = if ($Summary.Contains('TeamsSiteLinks')) { $Summary.TeamsSiteLinks } else { $null }
    $overview = @(
        [pscustomobject]@{Metric='Inventory status';Value=$Status}
        [pscustomobject]@{Metric='Sites inventoried';Value=$Summary.SitesProcessed}
        [pscustomobject]@{Metric='Sites without valid owner';Value=$Summary.OwnerlessSites}
        [pscustomobject]@{Metric='Inactive sites';Value=$Summary.InactiveSites}
        [pscustomobject]@{Metric='Sites above storage quota threshold';Value=$Summary.HighQuotaSites}
        [pscustomobject]@{Metric='SharePoint sites';Value=$Summary.SharePointSites}
        [pscustomobject]@{Metric='OneDrive sites';Value=$Summary.OneDriveSites}
        [pscustomobject]@{Metric='Lists inventoried';Value=$Summary.ListsProcessed}
        [pscustomobject]@{Metric='Storage used GB';Value=$Summary.StorageUsedGB}
        [pscustomobject]@{Metric=if(-not[string]::IsNullOrWhiteSpace([string]$Summary.TenantStorageCapacityGB)){'Tenant capacity measured GB'}elseif(-not[string]::IsNullOrWhiteSpace([string]$Summary.EstimatedTenantCapacityGB)){'Tenant capacity estimated from licenses GB'}else{'Tenant capacity'};Value=if(-not[string]::IsNullOrWhiteSpace([string]$Summary.TenantStorageCapacityGB)){$Summary.TenantStorageCapacityGB}elseif(-not[string]::IsNullOrWhiteSpace([string]$Summary.EstimatedTenantCapacityGB)){$Summary.EstimatedTenantCapacityGB}else{'Estimate unavailable'}}
        [pscustomobject]@{Metric=if(-not[string]::IsNullOrWhiteSpace([string]$Summary.TenantStorageUtilizationPercent)){'Tenant storage utilization measured percent'}elseif(-not[string]::IsNullOrWhiteSpace([string]$Summary.EstimatedTenantUtilizationPercent)){'Tenant storage utilization estimated percent'}else{'Tenant storage utilization percent'};Value=if(-not[string]::IsNullOrWhiteSpace([string]$Summary.TenantStorageUtilizationPercent)){$Summary.TenantStorageUtilizationPercent}else{$Summary.EstimatedTenantUtilizationPercent}}
        [pscustomobject]@{Metric='Object-level critical findings';Value=$Summary.CriticalAlerts}
        [pscustomobject]@{Metric='Object-level warning findings';Value=$Summary.WarningAlerts}
    )
    if ($teamsLinks) {
        $overview += @(
            [pscustomobject]@{Metric='Teams site mapping status';Value=if($teamsLinks.Available){'Qualified'}else{$teamsLinks.Reason}}
            [pscustomobject]@{Metric='SharePoint sites linked to Teams';Value=if($teamsLinks.Available){$teamsLinks.LinkedSites}else{''}}
            [pscustomobject]@{Metric='SharePoint sites linked to Teams percent';Value=if($teamsLinks.Available){$teamsLinks.Percent}else{''}}
            [pscustomobject]@{Metric='Team main sites';Value=if($teamsLinks.Available){$teamsLinks.ParentSites}else{''}}
            [pscustomobject]@{Metric='Private or shared channel sites';Value=if($teamsLinks.Available){$teamsLinks.ChannelSites}else{''}}
            [pscustomobject]@{Metric='Teams source run UTC';Value=$teamsLinks.TeamsRunDateUtc}
        )
    }
    $overview | Export-Excel -Path $Path -WorksheetName 'Summary' -TableName 'SpoSummary' -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter
    $findings = @($Alerts | Select-Object Severity,Category,SiteUrl,ObjectName,Metric,Value,Threshold,Details)
    if ($findings.Count -gt 0) {
        $findings | Export-Excel -Path $Path -WorksheetName 'Findings' -TableName 'SpoFindings' -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter
    }
    else {
        [pscustomobject]@{Severity='';Category='';SiteUrl='';ObjectName='';Metric='';Value='';Threshold='';Details=''} | Export-Excel -Path $Path -WorksheetName 'Findings' -TableName 'SpoFindings' -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter
        $package = Open-ExcelPackage -Path $Path
        try { $package.Workbook.Worksheets['Findings'].DeleteRow(2) }
        finally { Close-ExcelPackage -ExcelPackage $package }
    }
    if ($teamsLinks -and $teamsLinks.Available) {
        $linkRows = @($teamsLinks.Rows)
        if ($linkRows.Count -eq 0) {
            $linkRows = @([pscustomobject]@{SiteUrl='';SiteId='';LinkType='';TeamId='';TeamDisplayName='';ChannelId='';ChannelDisplayName='';TeamsRunDateUtc=''})
        }
        $linkRows | Export-Excel -Path $Path -WorksheetName 'Teams site links' -TableName 'SpoTeamsSiteLinks' -AutoSize -FreezeTopRow -BoldTopRow -AutoFilter
        if ($teamsLinks.Rows.Count -eq 0) {
            $package = Open-ExcelPackage -Path $Path
            try { $package.Workbook.Worksheets['Teams site links'].DeleteRow(2) }
            finally { Close-ExcelPackage -ExcelPackage $package }
        }
    }
}

function Ensure-SpoSharePointUploadRecord {
    param([Parameter(Mandatory)][string]$Path)
    $record = Get-SmartM365SharePointUploadRecordForLocalFile -FilePath $Path
    if (-not $record -and -not $DryRun) {
        $record = Invoke-SmartM365SharePointCsvUpload -LocalFilePath $Path
    }
    return $record
}

function New-SpoSharePointLinksHtml {
    param([Parameter(Mandatory)][string[]]$Paths)
    $links = foreach ($path in $Paths) {
        $record = Ensure-SpoSharePointUploadRecord -Path $path
        if (-not $record -or [string]::IsNullOrWhiteSpace([string]$record.WebUrl)) {
            Write-SpoLog -Message ("Mail link omitted because no SharePoint WebUrl is available for {0}." -f [IO.Path]::GetFileName($path)) -Level WARNING
            continue
        }
        $name = ConvertTo-SpoHtml ([IO.Path]::GetFileName($path))
        $url = ConvertTo-SpoHtml ([string]$record.WebUrl)
        "<li style='margin:0 0 6px;'><a href='$url' style='color:#075985;text-decoration:underline;'>$name</a></li>"
    }
    if (@($links).Count -eq 0) { return '' }
    return "<div style='margin-top:18px;padding:14px;border:1px solid #d9e2ec;background:#f8fafc;'><h2 style='font-size:15px;margin:0 0 8px;'>Timestamped exports</h2><p style='margin:0 0 8px;'>SharePoint links for this run:</p><ul style='margin:0;padding-left:20px;'>$($links -join '')</ul></div>"
}
function New-SpoAlertRow { param([Parameter(Mandatory)][string]$Severity,[Parameter(Mandatory)][string]$Category,[Parameter(Mandatory)][string]$SiteUrl,[string]$ObjectName='',[string]$Metric='',[AllowNull()]$Value=$null,[string]$Threshold='',[string]$Details='') [pscustomobject]@{Severity=$Severity;Category=$Category;SiteUrl=$SiteUrl;ObjectName=$ObjectName;Metric=$Metric;Value=if($null -eq $Value){''}else{[string]$Value};Threshold=$Threshold;Details=$Details} }

function New-SpoHtmlSummary {
    param([Parameter(Mandatory)][string]$WorstStatus,[Parameter(Mandatory)][AllowEmptyCollection()][object[]]$Alerts,[AllowEmptyCollection()][object[]]$SiteRows=@(),[Parameter(Mandatory)][System.Collections.IDictionary]$Summary,[bool]$PartialInventory,[string]$FileLinksHtml='',[string]$SourceNote='')
    $statusColor=switch($WorstStatus){'Critical'{'#991b1b'}'Warning'{'#92400e'}default{'#166534'}}; $statusBg=switch($WorstStatus){'Critical'{'#fee2e2'}'Warning'{'#fef3c7'}default{'#dcfce7'}}
    $affectedSites=@($Alerts|Where-Object{-not[string]::IsNullOrWhiteSpace([string]$_.SiteUrl)}|Select-Object -ExpandProperty SiteUrl -Unique).Count
    $inactiveWithDate=@($SiteRows|Where-Object{$_.IsInactive -and -not[string]::IsNullOrWhiteSpace([string]$_.LastActivityUtc)}).Count
    $unknownActivity=@($SiteRows|Where-Object{[string]::IsNullOrWhiteSpace([string]$_.LastActivityUtc)}).Count
    $graphUnavailable=@($SiteRows|Where-Object{$_.GraphEnrichmentStatus -ne 'Collected'}).Count
    $refreshDates=@($SiteRows|Where-Object{-not[string]::IsNullOrWhiteSpace([string]$_.ReportRefreshDate)}|ForEach-Object{[string]$_.ReportRefreshDate}|Sort-Object -Unique)
    $refreshText=if($refreshDates.Count -eq 0){'Not available'}elseif($refreshDates.Count -eq 1){$refreshDates[0]}else{"$($refreshDates[0]) to $($refreshDates[-1])"}
    $listWarnings=@($Alerts|Where-Object{$_.Category -eq 'ListViewThreshold'}).Count
    $measuredCapacity = -not [string]::IsNullOrWhiteSpace([string]$Summary.TenantStorageCapacityGB)
    $estimatedCapacity = -not [string]::IsNullOrWhiteSpace([string]$Summary.EstimatedTenantCapacityGB)
    if ($measuredCapacity -and -not [string]::IsNullOrWhiteSpace([string]$Summary.TenantStorageUtilizationPercent)) {
        $capacityLabel = 'Tenant utilization'
        $capacityValue = "$($Summary.TenantStorageUtilizationPercent)%"
        $capacityDetail = "Measured capacity: $($Summary.TenantStorageCapacityGB) GB"
    }
    elseif ($measuredCapacity) {
        $capacityLabel = 'Tenant capacity (measured)'
        $capacityValue = "$($Summary.TenantStorageCapacityGB) GB"
        $capacityDetail = 'Utilization unavailable for this run'
    }
    elseif ($estimatedCapacity) {
        $approx = [string][char]0x2248
        if (-not [string]::IsNullOrWhiteSpace([string]$Summary.EstimatedTenantUtilizationPercent)) {
            $capacityLabel = 'Tenant utilization (estimated)'
            $capacityValue = "$approx " + ([double]$Summary.EstimatedTenantUtilizationPercent).ToString('N2') + '%'
            $capacityDetail = ([double]$Summary.StorageUsedGB).ToString('N2') + " GB used / $approx " + ([double]$Summary.EstimatedTenantCapacityGB).ToString('N0') + ' GB estimated capacity'
        }
        else {
            $capacityLabel = 'Tenant capacity (estimated)'
            $capacityValue = "$approx " + ([double]$Summary.EstimatedTenantCapacityGB).ToString('N0') + ' GB'
            $capacityDetail = 'Utilization unavailable for this inventory scope'
        }
    }
    else {
        $capacityLabel = 'Tenant capacity'
        $capacityValue = 'Estimate unavailable'
        $capacityDetail = 'Current matching license evidence is unavailable'
    }
    $cards=@(
        [pscustomobject]@{Label='Sites inventoried';Value=([int]$Summary.SitesProcessed).ToString('N0');Detail='usage-report sites';Background='#f8fafc';Border='#dbe3ef';Accent='#0f172a';Span=1},
        [pscustomobject]@{Label='Sites to review';Value=$affectedSites.ToString('N0');Detail='distinct sites';Background='#eff6ff';Border='#bfdbfe';Accent='#1d4ed8';Span=1},
        [pscustomobject]@{Label='Without valid owner';Value=([int]$Summary.OwnerlessSites).ToString('N0');Detail='ownership review';Background='#fff7ed';Border='#fed7aa';Accent='#9a3412';Span=1},
        [pscustomobject]@{Label='Above storage threshold';Value=([int]$Summary.HighQuotaSites).ToString('N0');Detail='site quotas';Background='#faf5ff';Border='#e9d5ff';Accent='#7e22ce';Span=1},
        [pscustomobject]@{Label='Inactive with date';Value=$inactiveWithDate.ToString('N0');Detail='recorded activity';Background='#f0fdfa';Border='#99f6e4';Accent='#0f766e';Span=1},
        [pscustomobject]@{Label='Large lists';Value=$listWarnings.ToString('N0');Detail='above item threshold';Background='#f0fdf4';Border='#bbf7d0';Accent='#166534';Span=1},
        [pscustomobject]@{Label=$capacityLabel;Value=$capacityValue;Detail=$capacityDetail;Background='#eef2ff';Border='#c7d2fe';Accent='#4338ca';Span=2}
    )
    $siteThreshold=if(@($SiteRows).Count -gt 0){[string]$SiteRows[0].Threshold}else{''}
    $storageLabel=if($siteThreshold -match '^(?<value>[\d.]+)%'){"Site storage above $($Matches.value)%"}else{'Site storage above threshold'}
    $inactiveLabel=if($siteThreshold -match '; (?<value>\d+) days'){"No activity recorded for over $($Matches.value) days"}else{'Inactive with dated evidence'}
    $listAlert=@($Alerts|Where-Object Category -eq ListViewThreshold|Select-Object -First 1)
    $listLabel=if($listAlert.Count -gt 0 -and [string]$listAlert[0].Threshold -match '^\d+$'){"List above $($listAlert[0].Threshold) items"}else{'List above item threshold'}
    $categoryRows=New-Object 'System.Collections.Generic.List[object]'
    $categoryRows.Add(@('High','Site without a valid owner',$Summary.OwnerlessSites,'Verify ownership'))
    $categoryRows.Add(@('High',$storageLabel,$Summary.HighQuotaSites,'Review storage and quota'))
    $categoryRows.Add(@('Review',$inactiveLabel,$inactiveWithDate,'Confirm use before cleanup'))
    $categoryRows.Add(@('Review',$listLabel,$listWarnings,'Review list design'))
    $actionRows=@($categoryRows.ToArray()|Where-Object{[int]$_[2] -gt 0})
    $categoriesHtml=if($actionRows.Count -gt 0){New-SmartM365EmailTableHtml -Headers @('Priority','Finding','Sites or lists','Suggested review') -Rows $actionRows}else{'<p>No action was flagged by the configured checks.</p>'}
    $rankedAlerts=@($Alerts|Sort-Object @{Expression={if($_.Severity -eq 'Critical'){0}else{1}}},@{Expression={switch($_.Category){'Ownership'{0}'Storage'{1}'InactiveSite'{2}default{3}}}},SiteUrl)
    $topAlerts=New-Object 'System.Collections.Generic.List[object]'
    $seenObjects=@{}
    foreach($alert in $rankedAlerts){
        $key=('{0}|{1}' -f [string]$alert.SiteUrl,[string]$alert.ObjectName).ToLowerInvariant()
        if($seenObjects.ContainsKey($key)){continue}
        $seenObjects[$key]=$true
        $topAlerts.Add($alert)
        if($topAlerts.Count -eq 5){break}
    }
    $topRows=foreach($alert in $topAlerts){
        $object=if(-not[string]::IsNullOrWhiteSpace([string]$alert.ObjectName)){"$($alert.SiteUrl) / $($alert.ObjectName)"}else{[string]$alert.SiteUrl}
        $finding=switch([string]$alert.Category){'Ownership'{'No valid owner'}'Storage'{'Storage above threshold'}'InactiveSite'{'Inactive or unknown activity'}'ListViewThreshold'{'List above item threshold'}default{[string]$alert.Category}}
        ,@($object,$finding,[string]$alert.Value,[string]$alert.Threshold)
    }
    $topHtml=if($topRows){New-SmartM365EmailTableHtml -Headers @('Site or list','Finding','Measured','Threshold') -Rows @($topRows)}else{'<p>No priority object to review.</p>'}
    $scope=if($PartialInventory){'Limited test run of usage-report sites'}else{'Usage-report sites; not a full tenant site-discovery claim'}
    $oneDrive=if($Summary.InventoryMode -eq 'CSVSnapshot' -and -not $Summary.IncludeOneDrive){'Not verifiable from CSV'}elseif($Summary.IncludeOneDrive){'Yes'}else{'No'}
    $sharing=if($Summary.InventoryMode -eq 'CSVSnapshot'){'Not verifiable from CSV'}elseif($Summary.DeepSharingScan){'Enabled'}else{'Not enabled'}
    $capacitySource=if([string]::IsNullOrWhiteSpace([string]$Summary.TenantCapacitySource)){'Unavailable'}else{[string]$Summary.TenantCapacitySource}
    $estimateEvidence=if($estimatedCapacity -and -not $measuredCapacity){"<br>License estimate: 1 TB + 10 GB x $($Summary.CapacityEstimateEligibleUnits) known eligible purchased licenses; extra storage excluded. License snapshot: $(ConvertTo-SpoHtml $Summary.CapacityEstimateAsOfUtc)."}else{''}
    $teamsLinks=if($Summary.Contains('TeamsSiteLinks')){$Summary.TeamsSiteLinks}else{$null}
    $teamsSitesHtml=if($teamsLinks -and $teamsLinks.Available){
        "<p style='margin:12px 0 0;'><b>Teams-linked sites:</b> $(ConvertTo-SpoHtml $teamsLinks.LinkedSites) / $(ConvertTo-SpoHtml $teamsLinks.TotalSites) ($(ConvertTo-SpoHtml $teamsLinks.Percent)%) SharePoint sites. Team sites: $(ConvertTo-SpoHtml $teamsLinks.ParentSites); private/shared channel sites: $(ConvertTo-SpoHtml $teamsLinks.ChannelSites).</p>"
    }else{
        $reason=if($teamsLinks){[string]$teamsLinks.Reason}else{'Teams site-link evidence was not assessed.'}
        "<p style='margin:12px 0 0;'><b>Teams-linked sites:</b> unavailable $([char]0x2014) $(ConvertTo-SpoHtml $reason)</p>"
    }
    $teamsEvidence=if($teamsLinks -and $teamsLinks.Available){"<br>Teams site-link source run: $(ConvertTo-SpoHtml $teamsLinks.TeamsRunDateUtc); linked sites absent from the SPO usage-report scope: $(ConvertTo-SpoHtml ($teamsLinks.UnmatchedParentSites + $teamsLinks.UnmatchedChannelSites))."}else{''}
    $technical = 'RunId: {0}<br>Duration: {1}<br>Mode: {2}<br>Sites processed: {3}' -f (ConvertTo-SpoHtml $Summary.RunId),(ConvertTo-SpoHtml $Summary.Duration),(ConvertTo-SpoHtml $Summary.InventoryMode),(ConvertTo-SpoHtml $Summary.SitesProcessed)
    return "<div style='font-family:Segoe UI,Arial,sans-serif;color:#1f2937;'><div style='background:#fff;border:1px solid #dde7f0;border-radius:8px;padding:16px;margin-bottom:16px;'><h1>SharePoint Online Inventory <span style='display:inline-block;border-radius:999px;background:$statusBg;color:$statusColor;border:1px solid $statusColor;padding:4px 12px;font-size:14px;'>$WorstStatus</span></h1><p>$($Summary.SitesProcessed) sites processed from usage reports; $affectedSites distinct sites require review. Findings can overlap. The complete detail is in the attached workbook.</p></div><div style='background:#fff;border:1px solid #dde7f0;border-radius:8px;padding:16px;margin-bottom:16px;'><h2>Summary</h2>$(New-SmartM365EmailKpiGridHtml -Cards $cards -RowSizes @(4,3))$teamsSitesHtml</div><div style='background:#fff;border:1px solid #dde7f0;border-radius:8px;padding:16px;margin-bottom:16px;'><h2>Action required</h2>$categoriesHtml<h3>First objects to review</h3>$topHtml</div><div style='background:#fff;border:1px solid #dde7f0;border-radius:8px;padding:16px;margin-bottom:16px;'><h2>Evidence coverage</h2><p>Scope: $scope<br>OneDrive included: $oneDrive<br>SharePoint sites: $($Summary.SharePointSites); OneDrive sites: $($Summary.OneDriveSites)<br>Usage report refresh: $(ConvertTo-SpoHtml $refreshText)<br>Sites without an activity date: $unknownActivity<br>Sites without Graph enrichment: $graphUnavailable<br>Tenant capacity source: $(ConvertTo-SpoHtml $capacitySource)$estimateEvidence$teamsEvidence<br>Deep sharing scan: $sharing<br>$(ConvertTo-SpoHtml $SourceNote)</p></div>$FileLinksHtml<div style='margin-top:18px;padding:14px;border:1px solid #d9e2ec;color:#64748b;font-size:12px;'><b>Run details</b><br>$technical</div></div>"
}

function Connect-SpoGraph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Config,
        [switch]$UseInteractiveAuth
    )

    $scopes = @('Reports.Read.All','Sites.Read.All','Directory.Read.All')
    if ($UseInteractiveAuth) {
        Write-SpoLog -Message 'Connecting to Microsoft Graph with delegated authentication.'
        Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop | Out-Null; Set-SmartM365InventoryConnectionOwnership -Connection Graph -Owned $true
    }
    else {
        $appId = [string](Get-ScriptLocalConfigValue -Config $Config -Name 'AppId' -DefaultValue '')
        $tenantId = [string](Get-ScriptLocalConfigValue -Config $Config -Name 'TenantId' -DefaultValue '')
        $thumbprint = [string](Get-ScriptLocalConfigValue -Config $Config -Name 'Thumbprint' -DefaultValue '')
        if ([string]::IsNullOrWhiteSpace($thumbprint)) {
            $thumbprint = [string](Get-ScriptLocalConfigValue -Config $Config -Name 'Thumb' -DefaultValue '')
        }
        if ([string]::IsNullOrWhiteSpace($appId) -or [string]::IsNullOrWhiteSpace($tenantId) -or [string]::IsNullOrWhiteSpace($thumbprint)) {
            throw 'Microsoft Graph app-only authentication requires AppId, TenantId, and Thumbprint/Thumb in SmartM365 configuration.'
        }
        Write-SpoLog -Message 'Connecting to Microsoft Graph with app-only certificate authentication.'
        Connect-MgGraph -ClientId $appId -TenantId $tenantId -CertificateThumbprint $thumbprint -NoWelcome -ErrorAction Stop | Out-Null; Set-SmartM365InventoryConnectionOwnership -Connection Graph -Owned $true
    }

    Invoke-SpoWithRetry -Operation 'Graph organization probe' -ScriptBlock {
        Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id,displayName' -OutputType PSObject -ErrorAction Stop
    } | Out-Null
    Write-SpoLog -Message 'Connected to Microsoft Graph.' -Level SUCCESS
}

function Invoke-SpoGraphRequest {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Operation = 'Graph request',
        [string]$OutputFilePath = ''
    )

    Invoke-SpoWithRetry -Operation $Operation -ScriptBlock {
        $parameters = @{ Method = 'GET'; Uri = $Uri; ErrorAction = 'Stop' }
        if ([string]::IsNullOrWhiteSpace($OutputFilePath)) {
            $parameters.OutputType = 'PSObject'
        }
        else {
            $parameters.OutputFilePath = $OutputFilePath
        }
        Invoke-MgGraphRequest @parameters
    }
}

function Get-SpoGraphPagedValues {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Uri,
        [string]$Operation = 'Graph paged request',
        [scriptblock]$RequestInvoker
    )

    $values = New-Object System.Collections.Generic.List[object]
    $visited = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $nextUri = $Uri
    $pageNumber = 0
    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        if (-not $visited.Add($nextUri)) { throw "$Operation returned a repeated @odata.nextLink; collection is incomplete." }
        $pageNumber++
        $page = if ($null -ne $RequestInvoker) { & $RequestInvoker $nextUri } else { Invoke-SpoGraphRequest -Uri $nextUri -Operation "$Operation page $pageNumber" }
        if ($null -eq $page -or $null -eq $page.PSObject.Properties['value']) {
            throw "$Operation page $pageNumber returned an invalid Graph collection response without a value property."
        }
        foreach ($row in @($page.value)) { $values.Add($row) | Out-Null }
        $nextUri = ConvertTo-SpoText (Get-SpoPropertyValue -Object $page -Names @('@odata.nextLink','odata.nextLink'))
    }
    return $values.ToArray()
}


function Get-SpoGraphSiteLookupByReportId {
    $lookup = @{}
    try {
        $searchTerm = ''
        try { if (-not [string]::IsNullOrWhiteSpace([string]$global:SharePointSiteHostname)) { $searchTerm = ([string]$global:SharePointSiteHostname -split '\.')[0] } } catch { $searchTerm = '' }
        if ([string]::IsNullOrWhiteSpace($searchTerm)) { $searchTerm = (([string]$TenantName) -split '\.')[0] }
        if ([string]::IsNullOrWhiteSpace($searchTerm)) { $searchTerm = 'sharepoint' }
        $uri = 'https://graph.microsoft.com/v1.0/sites?search=' + [System.Uri]::EscapeDataString($searchTerm)
        $sites = @(Get-SpoGraphPagedValues -Uri $uri -Operation 'Graph sites search')
        foreach ($site in $sites) {
            $graphId = ConvertTo-SpoText (Get-SpoPropertyValue -Object $site -Names @('id'))
            if ([string]::IsNullOrWhiteSpace($graphId)) { continue }
            $parts = $graphId -split ','
            if ($parts.Count -ge 2 -and -not [string]::IsNullOrWhiteSpace($parts[1])) {
                $lookup[$parts[1].ToLowerInvariant()] = $site
            }
        }
        Write-SpoLog -Message ("Graph sites search lookup loaded: {0} site ids." -f $lookup.Count) -Level INFO
    }
    catch {
        Write-SpoLog -Message ("Graph sites search lookup unavailable: {0}" -f $_.Exception.Message) -Level INFO
    }
    return $lookup
}
function Get-SpoGraphReportCsv {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ReportName,
        [Parameter(Mandatory)][string]$OutputFolder
    )

    if (-not (Test-Path -LiteralPath $OutputFolder)) { New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null }
    $safeName = $ReportName -replace '[^A-Za-z0-9_-]+', '_'
    $rawPath = Join-Path -Path $OutputFolder -ChildPath ("{0}_{1}.csv" -f $safeName, (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $uri = "https://graph.microsoft.com/v1.0/reports/{0}(period='D180')" -f $ReportName
    Invoke-SpoGraphRequest -Uri $uri -Operation $ReportName -OutputFilePath $rawPath | Out-Null
    if (-not (Test-Path -LiteralPath $rawPath -PathType Leaf) -or (Get-Item -LiteralPath $rawPath).Length -eq 0) { throw "Graph report did not create non-empty CSV content: $rawPath" }
    return @(Import-Csv -LiteralPath $rawPath -ErrorAction Stop)
}

function Get-SpoReportValue {
    [CmdletBinding()]
    param(
        [AllowNull()]$Row,
        [Parameter(Mandatory)][string[]]$Names,
        [AllowNull()]$DefaultValue = $null
    )

    return Get-SpoPropertyValue -Object $Row -Names $Names -DefaultValue $DefaultValue
}

function ConvertFrom-SpoBytesToMb {
    [CmdletBinding()]
    param([AllowNull()]$Value)

    $bytes = ConvertTo-SpoDouble $Value
    if ($null -eq $bytes) { return 0 }
    return [Math]::Round($bytes / 1MB, 2)
}

function Resolve-SpoGraphSite {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$SiteUrl)

    try {
        $siteUri = [uri]$SiteUrl
        $graphUri = if ([string]::IsNullOrWhiteSpace($siteUri.AbsolutePath) -or $siteUri.AbsolutePath -eq '/') {
            "https://graph.microsoft.com/v1.0/sites/$($siteUri.Host)?`$select=id,displayName,webUrl,createdDateTime,lastModifiedDateTime"
        }
        else {
            "https://graph.microsoft.com/v1.0/sites/$($siteUri.Host):$($siteUri.AbsolutePath)?`$select=id,displayName,webUrl,createdDateTime,lastModifiedDateTime"
        }
        return Invoke-SpoGraphRequest -Uri $graphUri -Operation "Resolve Graph site $SiteUrl"
    }
    catch {
        Write-SpoLog -Message ("Graph site enrichment unavailable for {0}: {1}" -f $SiteUrl, $_.Exception.Message) -Level INFO
        return $null
    }
}

function Get-SpoListInventoryGraph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RunId,
        [Parameter(Mandatory)][string]$RunDateUtc,
        [Parameter(Mandatory)][string]$TenantName,
        [Parameter(Mandatory)][string]$SiteUrl,
        [AllowNull()]$GraphSite,
        [int]$WarningThreshold
    )

    if ($null -eq $GraphSite -or [string]::IsNullOrWhiteSpace([string]$GraphSite.id)) {
        return @([pscustomobject]@{
            RunId = $RunId; RunDateUtc = $RunDateUtc; TenantName = $TenantName; SiteUrl = $SiteUrl
            ListId = ''; ListTitle = ''; ListUrl = ''; BaseTemplate = ''; BaseType = ''; Hidden = ''
            ItemCount = ''; SizeMB = ''; VersioningEnabled = ''; MajorVersionLimit = ''
            Status = 'OK'; NumericValue = 0; TextValue = 'NotAvailableGraphOnly'; Threshold = $WarningThreshold
            UnavailableFields = 'BaseType;ItemCount;SizeMB;VersioningEnabled;MajorVersionLimit'
            Details = 'List inventory unavailable because the site could not be resolved through Graph.'
        })
    }

    $encodedSiteId = [System.Uri]::EscapeDataString([string]$GraphSite.id)
    $listsUri = "https://graph.microsoft.com/v1.0/sites/$encodedSiteId/lists?`$select=id,displayName,webUrl,createdDateTime,lastModifiedDateTime,list"
    try {
        $lists = @(Get-SpoGraphPagedValues -Uri $listsUri -Operation "Graph lists $SiteUrl")
    }
    catch {
        Write-SpoLog -Message ("Graph list inventory unavailable for {0}: {1}" -f $SiteUrl, $_.Exception.Message) -Level INFO
        return @([pscustomobject]@{
            RunId = $RunId; RunDateUtc = $RunDateUtc; TenantName = $TenantName; SiteUrl = $SiteUrl
            ListId = ''; ListTitle = ''; ListUrl = ''; BaseTemplate = ''; BaseType = ''; Hidden = ''
            ItemCount = ''; SizeMB = ''; VersioningEnabled = ''; MajorVersionLimit = ''
            Status = 'OK'; NumericValue = 0; TextValue = 'NotAvailableGraphOnly'; Threshold = $WarningThreshold
            UnavailableFields = 'BaseType;ItemCount;SizeMB;VersioningEnabled;MajorVersionLimit'
            Details = 'List inventory unavailable through Graph for this site.'
        })
    }

    foreach ($list in $lists) {
        $listInfo = Get-SpoPropertyValue -Object $list -Names @('list') -DefaultValue $null
        $itemCount = ConvertTo-SpoDouble (Get-SpoPropertyValue -Object $listInfo -Names @('itemCount') -DefaultValue $null)
        $numericValue = if ($null -eq $itemCount) { 0 } else { [int]$itemCount }
        $status = if ($null -ne $itemCount -and $itemCount -gt $WarningThreshold) { 'Warning' } else { 'OK' }
        [pscustomobject]@{
            RunId = $RunId
            RunDateUtc = $RunDateUtc
            TenantName = $TenantName
            SiteUrl = $SiteUrl
            ListId = ConvertTo-SpoText (Get-SpoPropertyValue -Object $list -Names @('id'))
            ListTitle = ConvertTo-SpoText (Get-SpoPropertyValue -Object $list -Names @('displayName','name'))
            ListUrl = ConvertTo-SpoText (Get-SpoPropertyValue -Object $list -Names @('webUrl'))
            BaseTemplate = ConvertTo-SpoText (Get-SpoPropertyValue -Object $listInfo -Names @('template'))
            BaseType = ''
            Hidden = ConvertTo-SpoText (Get-SpoPropertyValue -Object $listInfo -Names @('hidden') -DefaultValue '')
            ItemCount = if ($null -eq $itemCount) { '' } else { [int]$itemCount }
            SizeMB = ''
            VersioningEnabled = ''
            MajorVersionLimit = ''
            Status = $status
            NumericValue = $numericValue
            TextValue = ConvertTo-SpoText (Get-SpoPropertyValue -Object $list -Names @('displayName','name'))
            Threshold = $WarningThreshold
            UnavailableFields = 'BaseType;SizeMB;VersioningEnabled;MajorVersionLimit'
            Details = if ($status -eq 'Warning') { "List exceeds $WarningThreshold items." } else { 'Versioning and list size are not exposed by Graph list inventory.' }
        }
    }
}

function Get-SpoPermissionInventoryGraph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RunId,
        [Parameter(Mandatory)][string]$RunDateUtc,
        [Parameter(Mandatory)][string]$TenantName,
        [Parameter(Mandatory)][string]$SiteUrl,
        [string]$Owner = ''
    )

    $isOrphaned = [string]::IsNullOrWhiteSpace($Owner) -or $Owner -match 'deleted|unknown|system account'
    [pscustomobject]@{
        RunId = $RunId
        RunDateUtc = $RunDateUtc
        TenantName = $TenantName
        SiteUrl = $SiteUrl
        PrincipalType = 'ReportedOwner'
        PrincipalName = $Owner
        PrincipalLoginName = $Owner
        IsSiteAdmin = ''
        IsExternal = ($Owner -match '#EXT#|@.*#ext#')
        IsDisabled = ''
        Status = if ($isOrphaned) { 'Critical' } else { 'OK' }
        NumericValue = if ($isOrphaned) { 0 } else { 1 }
        TextValue = if ([string]::IsNullOrWhiteSpace($Owner)) { 'No owner reported' } else { $Owner }
        Threshold = 'Valid owner required'
        UnavailableFields = 'IsSiteAdmin;IsDisabled'
        Details = if ($isOrphaned) { 'No valid owner was reported by Graph usage data.' } else { 'Owner reported by Graph usage data. Site collection admin enumeration requires optional deep mode or SPO admin APIs.' }
    }
}

function Get-SpoExternalSharingInventoryGraph {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RunId,
        [Parameter(Mandatory)][string]$RunDateUtc,
        [Parameter(Mandatory)][string]$TenantName,
        [Parameter(Mandatory)][string]$SiteUrl
    )

    [pscustomobject]@{
        RunId = $RunId
        RunDateUtc = $RunDateUtc
        TenantName = $TenantName
        SiteUrl = $SiteUrl
        ObjectType = 'Site'
        ObjectTitle = ''
        ObjectUrl = $SiteUrl
        SharingSignal = 'GraphOnlyMode'
        SharingValue = 'NotAvailableGraphOnly'
        LinkScope = ''
        Principal = ''
        IsAnonymous = ''
        IsExternal = ''
        Status = 'OK'
        NumericValue = 0
        TextValue = 'NotAvailableGraphOnly'
        Threshold = 'No anonymous sharing'
        Details = 'Tenant-wide anonymous/external sharing link discovery is not available in least-privilege Graph-only mode.'
    }
}

function Convert-SpoReportRowToSiteSeed {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Row,
        [hashtable]$SiteLookupById = @{},
        [switch]$IsOneDrive
    )

    $siteId = ConvertTo-SpoText (Get-SpoReportValue -Row $Row -Names @('Site Id','Site ID','Id'))
    $siteUrl = ConvertTo-SpoText (Get-SpoReportValue -Row $Row -Names @('Site URL','Site Url','URL','Url'))
    $lookupSite = $null
    if ([string]::IsNullOrWhiteSpace($siteUrl) -and -not [string]::IsNullOrWhiteSpace($siteId) -and $SiteLookupById.ContainsKey($siteId.ToLowerInvariant())) {
        $lookupSite = $SiteLookupById[$siteId.ToLowerInvariant()]
        $siteUrl = ConvertTo-SpoText (Get-SpoPropertyValue -Object $lookupSite -Names @('webUrl'))
    }
    if ([string]::IsNullOrWhiteSpace($siteUrl)) { return $null }
    $owner = ConvertTo-SpoText (Get-SpoReportValue -Row $Row -Names @('Owner Principal Name','Owner PrincipalName','Owner Display Name','Owner DisplayName','Owner'))
    $storageUsedMb = ConvertFrom-SpoBytesToMb (Get-SpoReportValue -Row $Row -Names @('Storage Used (Byte)','Storage Used Byte','Storage Used'))
    $storageQuotaMb = ConvertFrom-SpoBytesToMb (Get-SpoReportValue -Row $Row -Names @('Storage Allocated (Byte)','Storage Allocated Byte','Storage Allocated'))
    $title = ConvertTo-SpoText (Get-SpoReportValue -Row $Row -Names @('Site Name','Site Display Name','Display Name','Site URL'))
    if ([string]::IsNullOrWhiteSpace($title) -and $null -ne $lookupSite) { $title = ConvertTo-SpoText (Get-SpoPropertyValue -Object $lookupSite -Names @('displayName')) }
    [pscustomobject]@{
        SiteUrl = $siteUrl
        SiteId = $siteId
        ReportRefreshDate = ConvertTo-SpoText (Get-SpoReportValue -Row $Row -Names @('Report Refresh Date','ReportRefreshDate'))
        Title = $title
        Owner = $owner
        Template = ConvertTo-SpoText (Get-SpoReportValue -Row $Row -Names @('Root Web Template','RootWebTemplate','Template'))
        LastActivityRaw = Get-SpoReportValue -Row $Row -Names @('Last Activity Date','LastActivityDate') -DefaultValue $null
        StorageUsedMB = $storageUsedMb
        StorageQuotaMB = $storageQuotaMb
        IsOneDrive = [bool]$IsOneDrive
    }
}

function ConvertTo-SpoTenantAdminUrl {
    [CmdletBinding()]
    param([AllowNull()][string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }
    $candidate = $Value.Trim()
    if ($candidate -notmatch '^https?://') { $candidate = "https://$candidate" }
    try { $hostName = ([uri]$candidate).DnsSafeHost.ToLowerInvariant().TrimEnd('.') }
    catch { return '' }

    if ($hostName -match '^(?<Tenant>[^.]+?)(?:-admin|-my)?\.sharepoint\.com$') {
        return "https://$($Matches.Tenant)-admin.sharepoint.com"
    }
    return ''
}

function Resolve-SpoTenantAdminUrl {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Config,
        [Parameter(Mandatory)][string]$TenantName,
        [AllowEmptyCollection()][string[]]$SiteUrls = @()
    )

    $configuredUrl = [string](Get-ScriptLocalConfigValue -Config $Config -Name 'SharePointAdminUrl' -DefaultValue '')
    if (-not [string]::IsNullOrWhiteSpace($configuredUrl)) {
        $normalizedConfiguredUrl = ConvertTo-SpoTenantAdminUrl -Value $configuredUrl
        if ([string]::IsNullOrWhiteSpace($normalizedConfiguredUrl)) { throw "Invalid SharePointAdminUrl: $configuredUrl" }
        return [pscustomobject]@{ Url = $normalizedConfiguredUrl; Source = 'SharePointAdminUrl' }
    }

    $siteHostname = [string](Get-ScriptLocalConfigValue -Config $Config -Name 'SharePointSiteHostname' -DefaultValue '')
    foreach ($candidate in @($siteHostname) + @($SiteUrls | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
        $derivedUrl = ConvertTo-SpoTenantAdminUrl -Value $candidate
        if (-not [string]::IsNullOrWhiteSpace($derivedUrl)) {
            return [pscustomobject]@{ Url = $derivedUrl; Source = 'DerivedFromSharePointHostname' }
        }
    }

    if ($TenantName -match '^(?<Tenant>[^.]+)\.onmicrosoft\.com$') {
        return [pscustomobject]@{ Url = "https://$($Matches.Tenant)-admin.sharepoint.com"; Source = 'DerivedFromTenantName' }
    }

    throw 'Unable to derive the SharePoint administration URL from SharePointSiteHostname, collected site URLs, or TenantName. Configure SharePointAdminUrl as an override.'
}

function Get-SpoTenantCapacityRow {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$RunId,
        [Parameter(Mandatory)][string]$RunDateUtc,
        [Parameter(Mandatory)][string]$TenantName,
        [Parameter(Mandatory)][double]$StorageUsedGB,
        [Parameter(Mandatory)]$Config,
        [AllowEmptyCollection()][string[]]$SiteUrls = @(),
        [switch]$Enabled,
        [switch]$Interactive,
        [switch]$PartialInventory,
        [double]$WarningPercent = 80,
        [double]$CriticalPercent = 90
    )

    $capacityMb = $null
    $allocatedMb = $null
    $capacitySource = 'NotCollectedGraphOnly'
    $status = 'OK'
    $details = 'Licensed tenant capacity was not collected through PnP. The email and workbook may show a separate estimate from M365_Licenses_Tenant.csv.'

    if ($Enabled) {
        try {
            $adminEndpoint = Resolve-SpoTenantAdminUrl -Config $Config -TenantName $TenantName -SiteUrls $SiteUrls
            $adminUrl = [string]$adminEndpoint.Url
            Write-SpoLog -Message ("SharePoint tenant administration URL: {0} ({1})." -f $adminUrl, $adminEndpoint.Source)

            $appId = [string](Get-ScriptLocalConfigValue -Config $Config -Name 'AppId' -DefaultValue '')
            $tenantId = [string](Get-ScriptLocalConfigValue -Config $Config -Name 'TenantId' -DefaultValue '')
            $thumbprint = [string](Get-ScriptLocalConfigValue -Config $Config -Name 'Thumbprint' -DefaultValue (Get-ScriptLocalConfigValue -Config $Config -Name 'Thumb' -DefaultValue ''))
            $connectParameters = @{ Url = $adminUrl; ReturnConnection = $true; ErrorAction = 'Stop' }
            if (-not [string]::IsNullOrWhiteSpace($appId)) { $connectParameters.ClientId = $appId }
            if ($Interactive) {
                $connectParameters.Interactive = $true
            }
            else {
                if ([string]::IsNullOrWhiteSpace($appId) -or [string]::IsNullOrWhiteSpace($tenantId) -or [string]::IsNullOrWhiteSpace($thumbprint)) {
                    throw 'AppId, TenantId and Thumbprint are required for app-only SharePoint tenant capacity collection.'
                }
                $connectParameters.Tenant = $tenantId
                $connectParameters.Thumbprint = $thumbprint
            }

            $connection = Connect-PnPOnline @connectParameters
            $tenantProperties = Get-PnPTenant -Connection $connection -ErrorAction Stop
            $capacityMb = ConvertTo-SpoDouble (Get-SpoPropertyValue -Object $tenantProperties -Names @('StorageQuota') -DefaultValue $null)
            $allocatedMb = ConvertTo-SpoDouble (Get-SpoPropertyValue -Object $tenantProperties -Names @('StorageQuotaAllocated') -DefaultValue $null)
            if ($null -eq $capacityMb -or $capacityMb -le 0) {
                throw 'Get-PnPTenant did not return a positive StorageQuota value.'
            }
            $capacitySource = 'Get-PnPTenant.StorageQuota'
            $status = 'OK'
            $details = "Tenant storage capacity collected from $adminUrl; URL source: $($adminEndpoint.Source)."
        }
        catch {
            $status = 'Warning'
            $capacityErrorMessage = [string]$_.Exception.Message
            $capacityAccessDenied = $capacityErrorMessage -match '(?i)unauthori[sz]ed|access\s+denied|forbidden|\b401\b|\b403\b'
            if ($capacityAccessDenied) {
                $capacitySource = 'Get-PnPTenantUnauthorized'
                $details = "Tenant capacity collection was denied by SharePoint: $capacityErrorMessage Microsoft Graph Sites.Read.All does not authorize Get-PnPTenant. Unattended app-only access to the SharePoint administration API requires the SharePoint API application permission Sites.FullControl.All with administrator consent. SharePointAdminUrl is not the cause because the URL was resolved successfully."
            }
            else {
                $capacitySource = 'Get-PnPTenantUnavailable'
                $details = "Tenant capacity collection failed: $capacityErrorMessage SharePointAdminUrl is only required as an override when automatic derivation is not suitable."
            }
            Write-SpoLog -Message $details -Level WARNING
        }
    }

    $usedMb = [Math]::Round($StorageUsedGB * 1024, 2)
    $utilizationPercent = if ($null -ne $capacityMb -and $capacityMb -gt 0 -and -not $PartialInventory) { [Math]::Round(($usedMb / $capacityMb) * 100, 2) } else { $null }
    if ($PartialInventory) {
        $status = 'Warning'
        $details = 'Site inventory is limited by MaxSites/MaxItems; tenant utilization is intentionally left blank.'
    }
    elseif ($null -ne $utilizationPercent -and $utilizationPercent -ge $CriticalPercent) {
        $status = 'Critical'
        $details = "Tenant SharePoint storage utilization is at or above $CriticalPercent percent."
    }
    elseif ($null -ne $utilizationPercent -and $utilizationPercent -ge $WarningPercent) {
        $status = 'Warning'
        $details = "Tenant SharePoint storage utilization is at or above $WarningPercent percent."
    }

    [pscustomobject]@{
        RunId = $RunId
        RunDateUtc = $RunDateUtc
        TenantName = $TenantName
        StorageUsedMB = $usedMb
        StorageUsedGB = [Math]::Round($StorageUsedGB, 2)
        StorageUsedTB = [Math]::Round($StorageUsedGB / 1024, 4)
        StorageCapacityMB = if ($null -eq $capacityMb) { '' } else { [Math]::Round($capacityMb, 2) }
        StorageCapacityGB = if ($null -eq $capacityMb) { '' } else { [Math]::Round($capacityMb / 1024, 2) }
        StorageCapacityTB = if ($null -eq $capacityMb) { '' } else { [Math]::Round($capacityMb / 1048576, 4) }
        StorageQuotaAllocatedMB = if ($null -eq $allocatedMb) { '' } else { [Math]::Round($allocatedMb, 2) }
        StorageQuotaAllocatedGB = if ($null -eq $allocatedMb) { '' } else { [Math]::Round($allocatedMb / 1024, 2) }
        StorageUtilizationPercent = if ($null -eq $utilizationPercent) { '' } else { $utilizationPercent }
        CapacitySource = $capacitySource
        IsPartialInventory = [bool]$PartialInventory
        Status = $status
        NumericValue = if ($null -eq $utilizationPercent) { '' } else { $utilizationPercent }
        TextValue = $capacitySource
        Threshold = "Warning >= $WarningPercent%; Critical >= $CriticalPercent%"
        Details = $details
    }
}
Import-SmartM365CoreModule
Ensure-SpoImportExcelModule
$requiredSpoModules = if ($MailOnly) { @('ImportExcel') } else { @('Microsoft.Graph.Authentication','ImportExcel') }
if (-not $MailOnly -and ($UsePnPDeepScan -or $TenantCapacityEnabled)) { $requiredSpoModules += 'PnP.PowerShell' }
Invoke-SmartM365Preflight `
    -ScriptName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) `
    -RequiredModules $requiredSpoModules | Out-Null
if (-not $MailOnly) { Import-Module Microsoft.Graph.Authentication -ErrorAction Stop }
    Initialize-SmartM365ImportExcel -InstallIfMissing | Out-Null
if (-not $MailOnly -and ($UsePnPDeepScan -or $TenantCapacityEnabled)) { Import-Module PnP.PowerShell -ErrorAction Stop }

$ScriptLocalConfig = Get-ScriptLocalConfig
$TenantName = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'OrgDomain' -DefaultValue $Tenant); if ([string]::IsNullOrWhiteSpace($TenantName)) { $TenantName = $Tenant }
$RunId = [guid]::NewGuid().ToString()
$RunDateUtc = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ', [System.Globalization.CultureInfo]::InvariantCulture)
$runStart = Get-Date
$DeepSharingScanEnabled = [bool]$UsePnPDeepScan -and -not [bool]$SkipDeepSharingScan

$global:RetentionMaxCSV = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxCSV' -DefaultValue 30)
$global:RetentionMaxLogs = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'RetentionMaxLogs' -DefaultValue 30)
$global:EnableSharePointUpload = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableSharePointUpload' -DefaultValue $false)
$global:SharePointSiteHostname = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSiteHostname' -DefaultValue ''
$global:SharePointSitePath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointSitePath' -DefaultValue ''
$global:SharePointLibraryDisplayName = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointLibraryDisplayName' -DefaultValue ''
$global:SharePointTargetFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'SharePointTargetFolderPath' -DefaultValue ''
$global:EnableWeeklyHistory = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'EnableWeeklyHistory' -DefaultValue $true)
$global:WeeklyHistoryFolderPath = Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'WeeklyHistoryFolderPath' -DefaultValue ''
$global:WeeklyHistoryRetentionWeeks = [int](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'WeeklyHistoryRetentionWeeks' -DefaultValue 52)
$ShowMailLinks = [bool](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ShowMailLinks' -DefaultValue $false)

$scriptOutputPath = if (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath
}
else {
    [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'ScriptCsvLogFolderPath' -DefaultValue "{{DataAllRootPath}}\M365\SharePoint\Inventory")
}
$latestFolder = [string](Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '')
if ([string]::IsNullOrWhiteSpace($latestFolder)) { throw 'LatestCsvFolderPath could not be resolved from configuration.' }

$CurrentOperation = 'InitializeScriptEnvironment'
$initializedOutput = InitializeScriptEnvironment -OutputPath $scriptOutputPath -LogFileName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))
if (-not $MailOnly) {
    Start-SmartM365CmdbSourceReceipt -ScriptPath $PSCommandPath -SourceRootPath (Get-ScriptLocalConfigValue -Config $ScriptLocalConfig -Name 'LatestCsvFolderPath' -DefaultValue '') -ReadOnly:$DryRun
    # DATA-LAST changes only when this run completes; DryRun and MailOnly hold no receipt.
    Enable-SmartM365StagedSourcePublication
}
if ([string]::IsNullOrWhiteSpace($global:WeeklyHistoryFolderPath)) {
    $global:WeeklyHistoryFolderPath = Join-Path -Path $initializedOutput -ChildPath 'WeeklyHistory'
}

$siteRows = New-Object System.Collections.Generic.List[object]
$listRows = New-Object System.Collections.Generic.List[object]
$permissionRows = New-Object System.Collections.Generic.List[object]
$sharingRows = New-Object System.Collections.Generic.List[object]
$tenantRows = New-Object System.Collections.Generic.List[object]
$alerts = New-Object System.Collections.Generic.List[object]
$generatedCsvPaths = New-Object System.Collections.Generic.List[string]
$scriptError = $null

try {
    if ($MailOnly) {
        $CurrentOperation = 'Validate existing SPO CSVs'
        if ($MaxSites -gt 0 -or $MaxItems -gt 0) { throw 'MailOnly cannot be combined with MaxSites or MaxItems.' }
        $tenantKey = if ($script:SmartM365GlobalConfig.PSObject.Properties['TenantKey']) { [string]$script:SmartM365GlobalConfig.TenantKey } else { $Tenant }
        $snapshot = Import-SpoMailCsvSet -Folder $latestFolder -ExpectedTenantKey $tenantKey
        $siteArray = @($snapshot.Rows.Sites); $listArray = @($snapshot.Rows.Lists)
        foreach($site in $siteArray){foreach($field in @('IsOneDrive','IsInactive','IsOrphaned')){$site.$field=([string]$site.$field -eq 'True')}}
        $permissionArray = @($snapshot.Rows.Permissions); $sharingArray = @($snapshot.Rows.ExternalSharing)
        $tenantArray = @($snapshot.Rows.Tenant); $tenantCapacity = $tenantArray[0]
        $sourceThreshold = if ($siteArray.Count -gt 0 -and [string]$siteArray[0].Threshold -match '^(?<value>[\d.]+)%') { [double]::Parse($Matches.value,[Globalization.CultureInfo]::InvariantCulture) } else { $QuotaCriticalPercent }
        $alertArray = New-Object 'System.Collections.Generic.List[object]'
        foreach ($site in $siteArray) {
            $category = if ([string]$site.Details -like 'Storage quota usage is above*') { 'Storage' } elseif ([string]$site.Details -like 'Site has no valid owner*') { 'Ownership' } elseif ([string]$site.Details -like 'Site has no activity within*') { 'InactiveSite' } else { '' }
            if (-not $category) { continue }
            $severity = if ($category -eq 'InactiveSite') { 'Warning' } else { 'Critical' }
            $value = switch ($category) { 'Storage' { $site.StorageQuotaPercent } 'Ownership' { $site.Owner } default { $site.TextValue } }
            $threshold = switch ($category) { 'Storage' { "$sourceThreshold%" } 'Ownership' { 'Valid owner required' } default { if ([string]$site.Threshold -match '; (?<days>\d+ days)') { $Matches.days } else { "$InactiveDays days" } } }
            $alertArray.Add((New-SpoAlertRow -Severity $severity -Category $category -SiteUrl ([string]$site.SiteUrl) -Metric ([string]$category) -Value $value -Threshold $threshold -Details ([string]$site.Details))) | Out-Null
        }
        foreach ($list in $listArray) {
            if ([string]$list.Status -eq 'Warning') { $alertArray.Add((New-SpoAlertRow -Severity Warning -Category ListViewThreshold -SiteUrl ([string]$list.SiteUrl) -ObjectName ([string]$list.ListTitle) -Metric ItemCount -Value $list.ItemCount -Threshold ([string]$list.Threshold) -Details ([string]$list.Details))) | Out-Null }
        }
        $RunId = $snapshot.RunId
        $partialInventory = [string]$tenantCapacity.IsPartialInventory -eq 'True'
        $quotaMeasure = $siteArray | Measure-Object -Property StorageQuotaGB -Sum
        $usedMeasure = $siteArray | Measure-Object -Property StorageUsedGB -Sum
        $summary = [ordered]@{
            Tenant=$TenantName; RunId=$RunId; Duration='Source CSV snapshot'; InventoryMode='CSVSnapshot'
            SitesProcessed=$siteArray.Count; SharePointSites=@($siteArray|Where-Object{[string]$_.IsOneDrive -ne 'True'}).Count; OneDriveSites=@($siteArray|Where-Object{[string]$_.IsOneDrive -eq 'True'}).Count
            OwnerlessSites=@($siteArray|Where-Object{[string]$_.IsOrphaned -eq 'True'}).Count; InactiveSites=@($siteArray|Where-Object{[string]$_.IsInactive -eq 'True'}).Count
            HighQuotaSites=@($siteArray|Where-Object{-not[string]::IsNullOrWhiteSpace([string]$_.StorageQuotaPercent) -and [double]$_.StorageQuotaPercent -gt $sourceThreshold}).Count
            ListsProcessed=$listArray.Count; PermissionRows=$permissionArray.Count; ExternalSharingRows=$sharingArray.Count
            CriticalAlerts=@($alertArray|Where-Object Severity -eq Critical).Count; WarningAlerts=@($alertArray|Where-Object Severity -eq Warning).Count
            StorageUsedGB=if($null -ne $usedMeasure.Sum){[math]::Round([double]$usedMeasure.Sum,2)}else{0}; StorageQuotaGB=if($null -ne $quotaMeasure.Sum){[math]::Round([double]$quotaMeasure.Sum,2)}else{0}
            TenantStorageCapacityGB=[string]$tenantCapacity.StorageCapacityGB; TenantStorageUtilizationPercent=[string]$tenantCapacity.StorageUtilizationPercent; TenantCapacitySource=[string]$tenantCapacity.CapacitySource
            IncludeOneDrive=@($siteArray|Where-Object{[string]$_.IsOneDrive -eq 'True'}).Count -gt 0; DeepSharingScan=$false; SharingScanItemLimitPerSite=0
        }
        Add-SpoLicenseEstimateToSummary -Summary $summary -Folder $latestFolder -ExpectedTenantKey $tenantKey -ExpectedTenantId ([string]$tenantCapacity.TenantId) -ReferenceDateUtc $snapshot.RunDateUtc -PartialInventory:$partialInventory
        $summary['TeamsSiteLinks'] = Get-SpoTeamsSiteLinksFromCsv -Folder $latestFolder -ExpectedTenantKey $tenantKey -ExpectedTenantId ([string]$tenantCapacity.TenantId) -ReferenceDateUtc $snapshot.RunDateUtc -SiteRows $siteArray -PartialInventory:$partialInventory
        $worstStatus = Get-SpoMailStatus -Summary $summary -PartialInventory:$partialInventory
        $CurrentOperation = 'Build SPO mail from existing CSVs'
        $stamp = '{0}_{1}' -f (Get-Date).ToUniversalTime().ToString('yyyyMMdd_HHmmss_fff',[Globalization.CultureInfo]::InvariantCulture),[guid]::NewGuid().ToString('N').Substring(0,8)
        $workbookPath = Join-Path $initializedOutput "M365_SPO_Inventory_MailOnly_$stamp.xlsx"
        New-SpoTimestampedWorkbook -CsvFiles $snapshot.Files -Path $workbookPath | Out-Null
        foreach($file in $snapshot.Files){if((Get-FileHash -LiteralPath $file.Path -Algorithm SHA256).Hash -ne $file.Sha256){throw "MailOnly CSV changed before the email was prepared: $($file.Path)"}}
        Add-SpoWorkbookOverview -Path $workbookPath -Summary $summary -Alerts $alertArray.ToArray() -Status $worstStatus
        $prefix = "[SNAPSHOT][$($worstStatus.ToUpperInvariant())]"
        $subject = "$prefix SmartM365 SharePoint Online inventory | $($summary.OwnerlessSites) without valid owner | $($summary.HighQuotaSites) above quota threshold"
        $sourceNote = "Mail reconstructed from existing CSVs collected $($snapshot.RunDateUtc). No new inventory collection; source freshness was not reverified."
        $bodyHtml = New-SpoHtmlSummary -WorstStatus $worstStatus -Alerts $alertArray.ToArray() -SiteRows $siteArray -Summary $summary -PartialInventory:$partialInventory -SourceNote $sourceNote
        if ($DryRun) { Write-SpoLog -Message 'DryRun enabled: MailOnly summary email skipped.' }
        else {
            $dailySummaryMarkerPath = Join-Path (Split-Path $global:LogTextFile -Parent) "$([IO.Path]::GetFileNameWithoutExtension($PSCommandPath))-DailySummary-LastSent.txt"
            $sent = Invoke-SpoDailySummaryMail -MarkerPath $dailySummaryMarkerPath -ForceSend:$AlwaysSend -SendAction { Send-SmartM365Mail -Subject $subject -BodyHtml $bodyHtml -Attachments @($workbookPath) -AllowAttachments -SuppressAttachmentLinks }
            if ($sent) { Write-SpoLog -Message "MailOnly SPO summary email sent from source run $RunId." -Level SUCCESS }
        }
        Complete-SmartM365ExecutionContext -Status Success
        $smartM365RunGuard.ReachedTerminal=$true; $smartM365RunGuard.ExplicitExit=$true; return
    }
    $CurrentOperation = 'Connect Microsoft Graph'
    Connect-SpoGraph -Config $ScriptLocalConfig -UseInteractiveAuth:$InteractiveAuth

    $CurrentOperation = 'Run Graph permission probes'
    Invoke-SmartM365Preflight `
        -ScriptName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) `
        -RequiredModules $requiredSpoModules `
        -RequiredGraphApplicationPermissions @('Sites.Read.All','Reports.Read.All') `
        -GraphProbeUris @(
            'https://graph.microsoft.com/v1.0/organization?$select=id,displayName',
            'https://graph.microsoft.com/v1.0/sites/root?$select=id,webUrl'
        ) | Out-Null

    $CurrentOperation = 'Get Graph usage reports'
    $rawReportFolder = Join-Path -Path $initializedOutput -ChildPath 'RawGraphReports'
    $tenantSiteReportRows = @(Get-SpoGraphReportCsv -ReportName 'getSharePointSiteUsageDetail' -OutputFolder $rawReportFolder)
    $siteLookupById = @{}
    if (@($tenantSiteReportRows | Where-Object { [string]::IsNullOrWhiteSpace([string](Get-SpoPropertyValue -Object $_ -Names @('Site URL','Site Url'))) }).Count -gt 0) {
        $siteLookupById = Get-SpoGraphSiteLookupByReportId
    }
    $siteSeeds = New-Object System.Collections.Generic.List[object]
    foreach ($row in $tenantSiteReportRows) {
        $seed = Convert-SpoReportRowToSiteSeed -Row $row -SiteLookupById $siteLookupById
        if ($null -ne $seed) { $siteSeeds.Add($seed) | Out-Null }
    }
    if ($IncludeOneDrive) {
        try {
            $oneDriveReportRows = @(Get-SpoGraphReportCsv -ReportName 'getOneDriveUsageAccountDetail' -OutputFolder $rawReportFolder)
            foreach ($row in $oneDriveReportRows) {
                $seed = Convert-SpoReportRowToSiteSeed -Row $row -SiteLookupById $siteLookupById -IsOneDrive
                if ($null -ne $seed) { $siteSeeds.Add($seed) | Out-Null }
            }
        }
        catch {
            Write-SpoLog -Message ("OneDrive usage report unavailable: {0}" -f $_.Exception.Message) -Level INFO
        }
    }

    $tenantSites = @($siteSeeds.ToArray() | Sort-Object -Property SiteUrl -Unique)
    if (-not $IncludeOneDrive) {
        $tenantSites = @($tenantSites | Where-Object { -not [bool]$_.IsOneDrive -and [string]$_.SiteUrl -notmatch '-my\.sharepoint\.com/personal/' })
    }
    if ($MaxSites -gt 0) { $tenantSites = @($tenantSites | Select-Object -First $MaxSites) }

    Write-SpoLog -Message ("Tenant sites discovered through Graph reports: {0}" -f $tenantSites.Count) -Level SUCCESS
    $processed = 0

    foreach ($site in $tenantSites) {
        $processed++
        $siteUrl = ConvertTo-SpoText $site.SiteUrl
        if ([string]::IsNullOrWhiteSpace($siteUrl)) { continue }
        Write-SpoLog -Message ("Processing site {0}/{1}: {2}" -f $processed, $tenantSites.Count, $siteUrl)

        $graphSite = Resolve-SpoGraphSite -SiteUrl $siteUrl
        $owner = ConvertTo-SpoText $site.Owner
        $template = ConvertTo-SpoText $site.Template
        if ([string]::IsNullOrWhiteSpace($template) -and $site.IsOneDrive) { $template = 'SPSPERS' }
        $storageUsedMb = ConvertTo-SpoDouble $site.StorageUsedMB
        $storageQuotaMb = ConvertTo-SpoDouble $site.StorageQuotaMB
        $quotaPercent = if ($storageQuotaMb -and $storageQuotaMb -gt 0) { [Math]::Round(($storageUsedMb / $storageQuotaMb) * 100, 2) } else { 0 }
        $lastActivityRaw = $site.LastActivityRaw
        $lastActivityDate = if ($lastActivityRaw) { try { [datetime]$lastActivityRaw } catch { $null } } else { $null }
        $daysInactive = if ($lastActivityDate) { [int]((Get-Date).ToUniversalTime().Date - $lastActivityDate.ToUniversalTime().Date).TotalDays } else { $null }
        $isInactive = ($null -eq $daysInactive) -or ($daysInactive -ge $InactiveDays)
        $isOrphaned = [string]::IsNullOrWhiteSpace($owner) -or $owner -match 'deleted|unknown|system account'
        $createdUtc = if ($null -ne $graphSite) { ConvertTo-SpoIsoUtc (Get-SpoPropertyValue -Object $graphSite -Names @('createdDateTime') -DefaultValue $null) } else { '' }
        $title = if ($null -ne $graphSite) { ConvertTo-SpoText (Get-SpoPropertyValue -Object $graphSite -Names @('displayName') -DefaultValue '') } else { '' }
        if ([string]::IsNullOrWhiteSpace($title)) { $title = ConvertTo-SpoText $site.Title }

        $siteStatus = 'OK'
        $siteDetails = 'OK'
        $siteNumeric = $quotaPercent
        $siteText = ''
        if ($quotaPercent -gt $QuotaCriticalPercent) {
            $siteStatus = 'Critical'
            $siteText = "$quotaPercent%"
            $siteDetails = "Storage quota usage is above $QuotaCriticalPercent%."
            $alerts.Add((New-SpoAlertRow -Severity Critical -Category Storage -SiteUrl $siteUrl -Metric 'StorageQuotaPercent' -Value $quotaPercent -Threshold "$QuotaCriticalPercent%" -Details $siteDetails)) | Out-Null
        }
        elseif ($isOrphaned) {
            $siteStatus = 'Critical'
            $siteText = $owner
            $siteDetails = 'Site has no valid owner in Graph usage data.'
            $alerts.Add((New-SpoAlertRow -Severity Critical -Category Ownership -SiteUrl $siteUrl -Metric Owner -Value $owner -Threshold 'Valid owner required' -Details $siteDetails)) | Out-Null
        }
        elseif ($isInactive) {
            $siteStatus = 'Warning'
            $siteText = if ($null -eq $daysInactive) { 'Unknown last activity' } else { "$daysInactive days" }
            $siteDetails = "Site has no activity within $InactiveDays days."
            $alerts.Add((New-SpoAlertRow -Severity Warning -Category InactiveSite -SiteUrl $siteUrl -Metric DaysInactive -Value $siteText -Threshold "$InactiveDays days" -Details $siteDetails)) | Out-Null
        }

        $nativeGraphSiteId = if ($graphSite) { ConvertTo-SpoText (Get-SpoPropertyValue $graphSite @('id')) } else { '' }
        $siteRows.Add([pscustomobject]@{
            RunId = $RunId
            RunDateUtc = $RunDateUtc
            TenantName = $TenantName
            SiteUrl = $siteUrl
            SiteId = if ($nativeGraphSiteId) { $nativeGraphSiteId } else { ConvertTo-SpoText $site.SiteId }
            ReportSiteId = ConvertTo-SpoText $site.SiteId
            SiteIdentitySource = if ($nativeGraphSiteId) { 'GraphSite' } elseif ($site.SiteId) { 'UsageReportSiteId' } else { 'Unavailable' }
            GraphEnrichmentStatus = if ($graphSite) { 'Collected' } else { 'Unavailable' }
            CollectionScope = if ($IncludeOneDrive) { 'UsageReportSitesIncludingOneDrive' } else { 'UsageReportSitesExcludingOneDrive' }
            ReportRefreshDate = $site.ReportRefreshDate
            CollectedAtUtc = [datetime]::UtcNow.ToString('o')
            Title = $title
            Template = $template
            CreatedUtc = $createdUtc
            LastActivityUtc = ConvertTo-SpoIsoUtc $lastActivityRaw
            DaysSinceLastActivity = if ($null -eq $daysInactive) { '' } else { $daysInactive }
            Owner = $owner
            LockState = 'NotAvailableGraphOnly'
            SharingCapability = 'NotAvailableGraphOnly'
            ExternalSharingEnabled = ''
            StorageQuotaMB = if ($null -eq $storageQuotaMb) { 0 } else { [Math]::Round($storageQuotaMb, 2) }
            StorageUsedMB = if ($null -eq $storageUsedMb) { 0 } else { [Math]::Round($storageUsedMb, 2) }
            StorageQuotaPercent = $quotaPercent
            StorageQuotaGB = if ($storageQuotaMb) { [Math]::Round($storageQuotaMb / 1024, 2) } else { 0 }
            StorageUsedGB = if ($storageUsedMb) { [Math]::Round($storageUsedMb / 1024, 2) } else { 0 }
            IsOneDrive = [bool]$site.IsOneDrive
            IsHubSite = ''
            HubSiteId = 'NotAvailableGraphOnly'
            RelatedGroupId = ''
            IsInactive = [bool]$isInactive
            IsOrphaned = [bool]$isOrphaned
            Status = $siteStatus
            NumericValue = $siteNumeric
            TextValue = $siteText
            Threshold = "$QuotaCriticalPercent%; $InactiveDays days"
            UnavailableFields = 'ExternalSharingEnabled;IsHubSite;RelatedGroupId'
            Details = $siteDetails
        }) | Out-Null

        foreach ($row in @(Get-SpoListInventoryGraph -RunId $RunId -RunDateUtc $RunDateUtc -TenantName $TenantName -SiteUrl $siteUrl -GraphSite $graphSite -WarningThreshold $ListItemWarningThreshold)) {
            $listRows.Add($row) | Out-Null
            if ($row.Status -eq 'Warning') {
                $alerts.Add((New-SpoAlertRow -Severity Warning -Category ListViewThreshold -SiteUrl $siteUrl -ObjectName $row.ListTitle -Metric ItemCount -Value $row.ItemCount -Threshold $ListItemWarningThreshold -Details $row.Details)) | Out-Null
            }
        }
        foreach ($row in @(Get-SpoPermissionInventoryGraph -RunId $RunId -RunDateUtc $RunDateUtc -TenantName $TenantName -SiteUrl $siteUrl -Owner $owner)) {
            $permissionRows.Add($row) | Out-Null
        }
        foreach ($row in @(Get-SpoExternalSharingInventoryGraph -RunId $RunId -RunDateUtc $RunDateUtc -TenantName $TenantName -SiteUrl $siteUrl)) {
            $sharingRows.Add($row) | Out-Null
        }

        if ($DeepSharingScanEnabled) {
            Write-SpoLog -Message 'UsePnPDeepScan is enabled. The site/list baseline remains Graph-based; tenant capacity is handled separately.' -Level INFO
        }
    }

    $CurrentOperation = 'Collect tenant storage capacity'
    $siteArrayForTenant = @($siteRows.ToArray())
    $usedMeasureForTenant = $siteArrayForTenant | Measure-Object -Property StorageUsedGB -Sum
    $storageUsedGbForTenant = if ($siteArrayForTenant.Count -gt 0 -and $null -ne $usedMeasureForTenant.Sum) { [Math]::Round([double]$usedMeasureForTenant.Sum, 2) } else { 0 }
    $tenantRows.Add((Get-SpoTenantCapacityRow -RunId $RunId -RunDateUtc $RunDateUtc -TenantName $TenantName -StorageUsedGB $storageUsedGbForTenant -Config $ScriptLocalConfig -SiteUrls @($siteArrayForTenant.SiteUrl) -Enabled:$TenantCapacityEnabled -Interactive:$InteractiveAuth -PartialInventory:($MaxSites -gt 0 -or (Test-SmartM365MaxItemsMode)) -CriticalPercent $QuotaCriticalPercent)) | Out-Null
    $CurrentOperation = 'Export CSV files'
    $siteColumns = @('RunId','RunDateUtc','TenantName','SiteUrl','Title','Template','CreatedUtc','LastActivityUtc','DaysSinceLastActivity','Owner','LockState','SharingCapability','ExternalSharingEnabled','StorageQuotaMB','StorageUsedMB','StorageQuotaPercent','StorageQuotaGB','StorageUsedGB','IsOneDrive','IsHubSite','HubSiteId','RelatedGroupId','IsInactive','IsOrphaned','Status','NumericValue','TextValue','Threshold','UnavailableFields','Details')
    $listColumns = @('RunId','RunDateUtc','TenantName','SiteUrl','ListId','ListTitle','ListUrl','BaseTemplate','BaseType','Hidden','ItemCount','SizeMB','VersioningEnabled','MajorVersionLimit','Status','NumericValue','TextValue','Threshold','UnavailableFields','Details')
    $siteColumns += @('SiteId','ReportSiteId','SiteIdentitySource','GraphEnrichmentStatus','CollectionScope','ReportRefreshDate','CollectedAtUtc')
    $permissionColumns = @('RunId','RunDateUtc','TenantName','SiteUrl','PrincipalType','PrincipalName','PrincipalLoginName','IsSiteAdmin','IsExternal','IsDisabled','Status','NumericValue','TextValue','Threshold','UnavailableFields','Details')
    $sharingColumns = @('RunId','RunDateUtc','TenantName','SiteUrl','ObjectType','ObjectTitle','ObjectUrl','SharingSignal','SharingValue','LinkScope','Principal','IsAnonymous','IsExternal','Status','NumericValue','TextValue','Threshold','Details')
    $tenantColumns = @('RunId','RunDateUtc','TenantName','StorageUsedMB','StorageUsedGB','StorageUsedTB','StorageCapacityMB','StorageCapacityGB','StorageCapacityTB','StorageQuotaAllocatedMB','StorageQuotaAllocatedGB','StorageUtilizationPercent','CapacitySource','IsPartialInventory','Status','NumericValue','TextValue','Threshold','Details')

    $exportStamp = (Get-Date).ToUniversalTime().ToString('yyyyMMdd_HHmmss',[Globalization.CultureInfo]::InvariantCulture)
    $timestampedCsvFiles = New-Object System.Collections.Generic.List[object]
    $weeklyHistorySourcePaths = New-Object System.Collections.Generic.List[string]
    foreach ($export in @(
        @{ Base = 'M365_SPO_Sites'; Data = $siteRows.ToArray(); Columns = $siteColumns; WorksheetName='Sites'; TableName='SpoSites' },
        @{ Base = 'M365_SPO_Lists'; Data = $listRows.ToArray(); Columns = $listColumns; WorksheetName='Lists'; TableName='SpoLists' },
        @{ Base = 'M365_SPO_Permissions'; Data = $permissionRows.ToArray(); Columns = $permissionColumns; WorksheetName='Permissions'; TableName='SpoPermissions' },
        @{ Base = 'M365_SPO_ExternalSharing'; Data = $sharingRows.ToArray(); Columns = $sharingColumns; WorksheetName='External sharing'; TableName='SpoExternalSharing' },
        @{ Base = 'M365_SPO_Tenant'; Data = $tenantRows.ToArray(); Columns = $tenantColumns; WorksheetName='Tenant capacity'; TableName='SpoTenant' }
    )) {
        $exportResult = Export-SpoEntityCsv -BaseFileName $export.Base -Data $export.Data -Columns $export.Columns -TimestampedFolder $initializedOutput -LatestFolder $latestFolder -Timestamp $exportStamp -AppendHistoryMode:$AppendHistory -NoWeeklyHistory:$global:EnableWeeklyHistory
        if ($exportResult.TimestampedPath) {
            $generatedCsvPaths.Add($exportResult.TimestampedPath) | Out-Null
            $timestampedCsvFiles.Add([pscustomobject]@{Path=$exportResult.TimestampedPath;WorksheetName=$export.WorksheetName;TableName=$export.TableName}) | Out-Null
        }
        if ($exportResult.LatestPath) {
            $generatedCsvPaths.Add($exportResult.LatestPath) | Out-Null
            # WeeklyHistory reads this run's copy (staged until promotion), never the previous run's file.
            $runCopyPath = if ($exportResult.PSObject.Properties['PublishedPath'] -and $exportResult.PublishedPath) { [string]$exportResult.PublishedPath } else { [string]$exportResult.LatestPath }
            $weeklyHistorySourcePaths.Add($runCopyPath) | Out-Null
        }
    }
    if ($global:EnableWeeklyHistory -and -not (Test-SmartM365MaxItemsMode)) {
        Add-SmartM365WeeklyHistory -SourceCsvPaths $weeklyHistorySourcePaths.ToArray() -HistoryRootPath $global:WeeklyHistoryFolderPath -RetentionWeeks $global:WeeklyHistoryRetentionWeeks -HistoryLabel 'SmartM365 SharePoint Online inventory' -UploadChangedFilesOnly
    }
    $workbookPath = Join-Path -Path $initializedOutput -ChildPath ("M365_SPO_Inventory_{0}.xlsx" -f $exportStamp)
    New-SpoTimestampedWorkbook -CsvFiles $timestampedCsvFiles.ToArray() -Path $workbookPath | Out-Null
    Remove-SmartM365TimestampedFilesOlderThan -FolderPath $initializedOutput -FilePattern 'M365_SPO_Inventory_*.xlsx' -RetentionDays 7 -LogFile $global:LogTextFile
    # Raw Graph report downloads are working copies: the published CSVs and WeeklyHistory keep the data.
    if (-not [string]::IsNullOrWhiteSpace($rawReportFolder)) { Remove-SmartM365TimestampedFilesOlderThan -FolderPath $rawReportFolder -FilePattern '*.csv' -RetentionDays 7 -LogFile $global:LogTextFile }
    $CurrentOperation = 'Build and send notification'
    $siteArray = @($siteRows.ToArray())
    $listArray = @($listRows.ToArray())
    $permissionArray = @($permissionRows.ToArray())
    $sharingArray = @($sharingRows.ToArray())
    $tenantArray = @($tenantRows.ToArray())
    $tenantCapacity = $tenantArray | Select-Object -First 1
    $alertArray = @($alerts.ToArray())
    $criticalAlertCount = @($alertArray | Where-Object { $_.Severity -eq 'Critical' }).Count
    $warningAlertCount = @($alertArray | Where-Object { $_.Severity -eq 'Warning' }).Count
    $duration = New-TimeSpan -Start $runStart -End (Get-Date)
    $quotaMeasure = $siteArray | Measure-Object -Property StorageQuotaGB -Sum
    $usedMeasure = $siteArray | Measure-Object -Property StorageUsedGB -Sum
    $totalQuotaGb = if ($siteArray.Count -gt 0 -and $null -ne $quotaMeasure -and $null -ne $quotaMeasure.PSObject.Properties['Sum'] -and $null -ne $quotaMeasure.Sum) { [Math]::Round([double]$quotaMeasure.Sum, 2) } else { 0 }
    $totalUsedGb = if ($siteArray.Count -gt 0 -and $null -ne $usedMeasure -and $null -ne $usedMeasure.PSObject.Properties['Sum'] -and $null -ne $usedMeasure.Sum) { [Math]::Round([double]$usedMeasure.Sum, 2) } else { 0 }
    $summary = [ordered]@{
        Tenant = $TenantName
        RunId = $RunId
        Duration = ('{0:hh\:mm\:ss}' -f $duration)
        InventoryMode = if ($TenantCapacityEnabled -and $DeepSharingScanEnabled) { 'Graph+PnPTenantCapacity+PnPDeepScan' } elseif ($TenantCapacityEnabled) { 'Graph+PnPTenantCapacity' } elseif ($DeepSharingScanEnabled) { 'Graph+PnPDeepScan' } else { 'GraphOnly' }
        SitesProcessed = $siteArray.Count
        SharePointSites = @($siteArray | Where-Object { -not [bool]$_.IsOneDrive }).Count
        OneDriveSites = @($siteArray | Where-Object { [bool]$_.IsOneDrive }).Count
        OwnerlessSites = @($siteArray | Where-Object { [bool]$_.IsOrphaned }).Count
        InactiveSites = @($siteArray | Where-Object { [bool]$_.IsInactive }).Count
        HighQuotaSites = @($siteArray | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_.StorageQuotaPercent) -and [double]$_.StorageQuotaPercent -gt $QuotaCriticalPercent }).Count
        ListsProcessed = $listArray.Count
        PermissionRows = $permissionArray.Count
        ExternalSharingRows = $sharingArray.Count
        CriticalAlerts = $criticalAlertCount
        WarningAlerts = $warningAlertCount
        StorageUsedGB = $totalUsedGb
        StorageQuotaGB = $totalQuotaGb
        TenantStorageCapacityGB = if ($null -ne $tenantCapacity -and -not [string]::IsNullOrWhiteSpace([string]$tenantCapacity.StorageCapacityGB)) { $tenantCapacity.StorageCapacityGB } else { '' }
        TenantStorageUtilizationPercent = if ($null -ne $tenantCapacity -and -not [string]::IsNullOrWhiteSpace([string]$tenantCapacity.StorageUtilizationPercent)) { $tenantCapacity.StorageUtilizationPercent } else { '' }
        TenantCapacitySource = if ($null -ne $tenantCapacity) { $tenantCapacity.CapacitySource } else { 'Unavailable' }
        IncludeOneDrive = [bool]$IncludeOneDrive
        DeepSharingScan = $DeepSharingScanEnabled
        SharingScanItemLimitPerSite = $SharingScanItemLimitPerSite
    }
    $tenantKey = [string]$script:SmartM365GlobalConfig.TenantKey
    $tenantId = [string]$script:SmartM365GlobalConfig.TenantId
    Add-SpoLicenseEstimateToSummary -Summary $summary -Folder $latestFolder -ExpectedTenantKey $tenantKey -ExpectedTenantId $tenantId -ReferenceDateUtc $RunDateUtc -PartialInventory:($MaxSites -gt 0 -or (Test-SmartM365MaxItemsMode))
    $summary['TeamsSiteLinks'] = Get-SpoTeamsSiteLinksFromCsv -Folder $latestFolder -ExpectedTenantKey $tenantKey -ExpectedTenantId $tenantId -ReferenceDateUtc $RunDateUtc -SiteRows $siteArray -PartialInventory:($MaxSites -gt 0 -or (Test-SmartM365MaxItemsMode))

    $worstStatus = Get-SpoMailStatus -Summary $summary -PartialInventory ($MaxSites -gt 0 -or (Test-SmartM365MaxItemsMode))
    Add-SpoWorkbookOverview -Path $workbookPath -Summary $summary -Alerts $alertArray -Status $worstStatus
    if (-not $DryRun) {
        $workbookUpload = Ensure-SpoSharePointUploadRecord -Path $workbookPath
        if ($workbookUpload) { Remove-SmartM365SharePointTimestampedCsvOlderThan -TimestampedPath $workbookPath -RetentionDays 7 | Out-Null }
    }

    $resultSummary = "SPO inventory $worstStatus; mode=$($summary.InventoryMode); sites=$($siteRows.Count); lists=$($listRows.Count); critical=$($summary.CriticalAlerts); warnings=$($summary.WarningAlerts)."
    if (-not $DryRun) {
        $prefix = if ($worstStatus -eq 'Critical') { '[CRITICAL]' } elseif ($worstStatus -eq 'Warning') { '[WARNING]' } else { '[OK]' }
        $subject = "$prefix SmartM365 SharePoint Online inventory | $($summary.OwnerlessSites) without valid owner | $($summary.HighQuotaSites) above quota threshold"
        $mailFileLinks = if ($ShowMailLinks) { New-SpoSharePointLinksHtml -Paths (@($timestampedCsvFiles.Path) + @($workbookPath)) } else { '' }
        $bodyHtml = New-SpoHtmlSummary -WorstStatus $worstStatus -Alerts $alertArray -SiteRows $siteArray -Summary $summary -PartialInventory:($MaxSites -gt 0 -or (Test-SmartM365MaxItemsMode)) -FileLinksHtml $mailFileLinks
        $dailySummaryMarkerPath = Join-Path -Path (Split-Path -Path $global:LogTextFile -Parent) -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))-DailySummary-LastSent.txt"
        $dailySummarySent = Invoke-SpoDailySummaryMail -MarkerPath $dailySummaryMarkerPath -ForceSend:$AlwaysSend -SendAction {
            Send-SmartM365Mail -Subject $subject -BodyHtml $bodyHtml -Attachments @($workbookPath) -AllowAttachments -SuppressAttachmentLinks
        }
        if ($dailySummarySent) {
            Write-SpoLog -Message ("Daily SharePoint summary email sent: {0}" -f $subject) -Level SUCCESS
        }
    }
    else {
        Write-SpoLog -Message 'DryRun enabled: daily summary email skipped.'
    }
    Write-SpoLog -Message $resultSummary -Level SUCCESS

    Set-SmartM365CmdbSourceScope -CompleteScope ($MaxSites -eq 0 -and $MaxItems -eq 0) -Scope 'CMDB:sites' -Qualifications @("Usage report site scope; IncludeOneDrive=$([bool]$IncludeOneDrive); not a full tenant site-discovery claim.")
    Complete-SmartM365ExecutionContext -Status Success

    $smartM365RunGuard.ReachedTerminal=$true
}
catch {
    if ($_.Exception -is [Management.Automation.PipelineStoppedException] -or $_.Exception -is [OperationCanceledException]) { throw }
    $smartM365RunGuard.ErrorRecord=$_

    $scriptError = $_
    Write-SpoLog -Message ("SPO inventory failed during {0}: {1}" -f $CurrentOperation, $_.Exception.Message) -Level ERROR
    try { Complete-SmartM365ExecutionContext -Status Failed -ErrorRecord $_ -FailureStage $CurrentOperation } catch { Write-SpoLog -Message ('Complete-SmartM365ExecutionContext failed after script error: {0}' -f $_.Exception.Message) -Level WARNING }
    throw
}
finally {
    $smartM365RunGuard.InCleanup=$true

    try {
        if (Get-Command Disconnect-PnPOnline -ErrorAction SilentlyContinue) { Disconnect-PnPOnline -ErrorAction SilentlyContinue }
    }
    catch { Write-SpoLog -Message ('Disconnect-PnPOnline failed during cleanup: {0}' -f $_.Exception.Message) -Level WARNING }
    try {
        if (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null; Set-SmartM365InventoryConnectionOwnership -Connection Graph -Owned $false }
    }
    catch { Write-SpoLog -Message ('Disconnect-MgGraph failed during cleanup: {0}' -f $_.Exception.Message) -Level WARNING }
    if ($null -ne $scriptError) { exit 1 }
}
    $smartM365RunGuard.ReachedTerminal=$true
    $smartM365RunGuard.ReturnedNormally=$true
} catch {
    if ($_.Exception -isnot [Management.Automation.PipelineStoppedException] -and $_.Exception -isnot [OperationCanceledException]) { $smartM365RunGuard.ErrorRecord=$_ }
    throw
} finally {
    Complete-SmartM365InventoryRun -State $smartM365RunGuard
    if($smartM365RunGuard.ExitCode -eq 130){exit 130}
}
if($smartM365RunGuard.ExitCode -ne 0){exit $smartM365RunGuard.ExitCode}


# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAK8uEDz6R8VoTe
# /FMiSxl4E6GCQE5aDwpvyeogD1ZzDqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIM5uD1Wc59cdmeNKPlSsCMCbwSDp1+3WtLvLCGflLv0sMA0GCSqG
# SIb3DQEBAQUABIIBgAjY9o7OiLW+0xCTZ8kSS1lQ1eUW2WTYYls4KdZU9c8+fTY2
# ZAgXhh6fSvEPNoeDyOp/8vEEpw8w84GOWH55eqzwawoETeLClSCHSaz5lheNhbre
# s2UBCs0nTCLOJ78ITCMb15qUSMhyO+20gdUvRGPJJmOYasjmXSH5JLcX2Mt4r/b5
# BN9o9kUcBhprMTWltDBfRc4lKRrgd9eKdYyZONv+on0PntzyErWFhFV7xdqqpvou
# 7YGy1o1IBKIYW1YvsLt/P8r3DL3yBLsxQ7cTLruoS0yj9/I5M8HMZQLqP7XTvdrk
# K77dcO4YloTLn4kriBc/qvLun7fv+g13r+HMCbwDtbtlBaDy2F/00O5xJkWM4u64
# T61O5YI5UZ/KbWCGjfg/RHmNqGgIFP3I08ke65P8113hjUMWqcc9YG8Q1dO4AXYi
# hEaYM+LVC+jqDdHfyR723rctGeVTXVag4J8tab7nLfel51+3em3dUFYmTb7AYzGQ
# VUfRk3WPMEy/ZaVlaKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjkyMDU4
# MDhaMC8GCSqGSIb3DQEJBDEiBCCcWpi6kiW1KynuaNXI6EPo3Uauoe1PtEEKoMJx
# 8wG0CDANBgkqhkiG9w0BAQEFAASCAgBSd7ojPClRUYY9V9k5hoA7WQK6uD4LJ5Uy
# F+pk3o++RiBPhbs89+r8+Ulef+cN1t0F+odKmoKof7XC2CIjGetXI6U1Z4npZWO9
# DNqgbSU+ibvBZAeL68+oMdInnzNvUyCzyq2It0ToOHDrWqW2XXE0YO336NtaVCkm
# Bw8l3piNo8S9hdwEudsJPR8230ZEq95/iOOPhH1o6B093DzNxRpYUHpg+XhCvFQE
# 8yfE3GJFZ4wYfmnxf0Zd2hYonneBZ27acAu4Lp2hKNljv4+P83QfshV9hf14ryrE
# i7bjI3gtfg7+B7EoADieEFeb8MadkrE+ZFFGIrxHJJULPZlY662+TKU0k6QJdnuK
# z9+nHSQVAN0FMJVxLkyjOGuJydmPIKVDI2iiLMOpEHSuv9jWL4dmbs0VWL39vpEV
# qMTunbcAE5S6vTFRS2HpGa+mO+5WkirPWNh3L3ewYSQ0jWM+NVmR67efEsItPy41
# 9rggCQ2CnmAOGgoJlbGiC3dVtZbjdgRQI8fwiJSSJIwnsQ2FShq9ZQ2IB+ow5DBF
# QGi/HpJb3KGVFctUPJROTAumGlFgrkpjm1Cq6wBZjv23lihSp9+9WsiYDXdfXRLK
# 4LodfryXN0F0x06AVJrgzEjEaWpjZgFwwsqLc2ubMpiUJxNSz0OjNmJrIGuDjt2K
# 1krQf4yQ0Q==
# SIG # End signature block
