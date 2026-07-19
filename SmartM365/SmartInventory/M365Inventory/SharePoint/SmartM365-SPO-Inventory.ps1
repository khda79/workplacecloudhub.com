#Requires -Version 7.0
<#
.SYNOPSIS
    SharePoint Online tenant inventory with CSV exports and HTML alert summary.

.DESCRIPTION
    Uses Microsoft Graph app-only authentication by default to inventory
    SharePoint Online sites, storage, activity, lists where Graph allows access,
    owner signals, inactive sites, orphaned sites, and tenant-level summary
    statistics. The default collection is Graph-only. PnP.PowerShell tenant
    capacity and deep sharing scans remain explicit optional features.

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
    Retained for launcher compatibility. The HTML summary is sent once per day even when the status is OK.

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
    Optionally collects the licensed SharePoint tenant storage capacity through
    Get-PnPTenant. Disabled by default. SharePointAdminUrl is derived automatically
    when omitted.

.PARAMETER SkipPnPTenantCapacity
    Retained for backward compatibility. Tenant capacity is already disabled by
    default unless -UsePnPTenantCapacity is specified.

.PARAMETER InteractiveAuth
    Uses delegated interactive Graph authentication instead of app-only certificate authentication.

.VERSION
0.24


.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication; ImportExcel. PnP.PowerShell is required only for optional PnP features.
    Minimum Graph application permissions for default inventory: Reports.Read.All; Sites.Read.All; Directory.Read.All.
    Optional PnP deep sharing scan may require SharePoint site-level access for scanned sites.
    Conditional: Mail.Send is required only when Graph mail is used; Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Requires: PowerShell 7+, Microsoft.Graph.Authentication, ImportExcel, SmartM365.Core.psd1
    PnP.PowerShell is used only with -UsePnPTenantCapacity or -UsePnPDeepScan.
    Minimum Microsoft Graph application permissions for default inventory:
      - Reports.Read.All for SharePoint and OneDrive usage reports.
      - Sites.Read.All for site and list inventory.
      - Directory.Read.All for organization context and owner/directory enrichment.
    Optional PnP deep sharing scan may require SharePoint site-level access for the scanned sites.
    Optional unattended tenant capacity collection requires the SharePoint API application permission Sites.FullControl.All with administrator consent.
    Microsoft Graph Sites.Read.All does not authorize Get-PnPTenant.
    The default Graph-only mode does not require SharePoint Administrator or Sites.FullControl.All.
#>

[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [switch]$IncludeOneDrive,
    [int]$InactiveDays = 180,
    [int]$MaxSites = 0,
    [switch]$DryRun,
    [switch]$AlwaysSend,
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
$ScriptVersion = "0.24"
$TenantCapacityEnabled = [bool]$UsePnPTenantCapacity -and -not [bool]$SkipPnPTenantCapacity
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
    exit 1
}

function Import-SmartM365CoreModule {
    $searchRoot = $PSScriptRoot
    while ($searchRoot) {
        $candidate = Join-Path -Path $searchRoot -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
        if (Test-Path -LiteralPath $candidate) { Import-Module -Name $candidate -MinimumVersion '1.0.38' -Force -ErrorAction Stop; return }
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
    $configPath = Join-Path -Path $PSScriptRoot -ChildPath ("{0}.local.json" -f [System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))
    if (-not (Test-Path -LiteralPath $configPath)) {
        $templatePath = '{0}.template' -f $configPath
        if (Get-Command Initialize-SmartM365LocalJsonFromTemplate -ErrorAction SilentlyContinue) { Initialize-SmartM365LocalJsonFromTemplate -Path $configPath -TemplatePath $templatePath -ConfigDescription 'script local configuration' | Out-Null }
        else { if (-not (Test-Path -LiteralPath $templatePath)) { throw "Missing local config and template: $configPath" }; Copy-Item -LiteralPath $templatePath -Destination $configPath -ErrorAction Stop }
    }
    return Get-Content -LiteralPath $configPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
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
        [Parameter(Mandatory)][scriptblock]$SendAction
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
        if ($lastSentDate.Trim() -eq $today) {
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
            $isTransient=($status -in @('429','500','502','503','504')) -or ($message -match '(?i)throttl|too many requests|temporar|timeout|503|429')
            if(-not $isTransient -or $attempt -ge $MaxAttempts){throw}
            $delay=[Math]::Min($MaxDelaySeconds,[int]($BaseDelaySeconds*[Math]::Pow(2,($attempt-1))))+(Get-Random -Minimum 0 -Maximum 4)
            Write-SpoLog -Message ("Transient failure on {0}. Status={1}; attempt {2}/{3}; waiting {4}s. {5}" -f $Operation,$(if($status){$status}else{'unknown'}),$attempt,$MaxAttempts,$delay,$message) -Level WARNING
            Start-Sleep -Seconds $delay
        }
    }
}

function Ensure-SpoUtf8Bom { param([Parameter(Mandatory)][string]$Path) if(-not(Test-Path -LiteralPath $Path -PathType Leaf)){return}; $bytes=[System.IO.File]::ReadAllBytes($Path); if($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF){return}; $bom=[byte[]](0xEF,0xBB,0xBF); $combined=[byte[]]::new($bom.Length+$bytes.Length); [Array]::Copy($bom,0,$combined,0,$bom.Length); [Array]::Copy($bytes,0,$combined,$bom.Length,$bytes.Length); [System.IO.File]::WriteAllBytes($Path,$combined) }
function Add-SpoHistoryCsv { param([AllowEmptyCollection()][object[]]$Data,[Parameter(Mandatory)][string]$Path,[Parameter(Mandatory)][string[]]$Columns) $Columns=@('TenantKey','OrganizationKey','EnvironmentKey','TenantId')+@($Columns|Where-Object{$_-inotmatch'^(TenantKey|OrganizationKey|EnvironmentKey|TenantId)$'}); $parent=Split-Path -Path $Path -Parent; if(-not(Test-Path -LiteralPath $parent)){New-Item -Path $parent -ItemType Directory -Force|Out-Null}; $rows=@($Data|Select-Object -Property $Columns); if(-not(Test-Path -LiteralPath $Path)){ $rows|Add-SmartM365TenantKey | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Delimiter ','; Ensure-SpoUtf8Bom -Path $Path; return }; Repair-SmartM365CsvTenantKeySchema -Path $Path -Delimiter ',' -Encoding UTF8|Out-Null; $rows|Add-SmartM365TenantKey | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Delimiter ',' -Append; Ensure-SpoUtf8Bom -Path $Path }
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
        $header = '"' + ($Columns -join '","') + '"'
        Set-Content -LiteralPath $timestampedPath -Value $header -Encoding UTF8
        Set-Content -LiteralPath $latestPath -Value $header -Encoding UTF8
        Ensure-SpoUtf8Bom -Path $timestampedPath
        Ensure-SpoUtf8Bom -Path $latestPath
        if (-not $DryRun) {
            Invoke-SmartM365SharePointCsvUpload -LocalFilePath $timestampedPath | Out-Null
            Invoke-SmartM365SharePointCsvUpload -LocalFilePath $latestPath | Out-Null
        }
        return [pscustomobject]@{ TimestampedPath=$timestampedPath; LatestPath=$latestPath }
    }
    $result = Export-SmartM365Csv -Data $Data -TimestampedPath $timestampedPath -LatestPath $latestPath -Columns $Columns -NoWeeklyHistory:$NoWeeklyHistory
    Ensure-SpoUtf8Bom -Path $timestampedPath
    Ensure-SpoUtf8Bom -Path $latestPath
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

function Ensure-SpoImportExcelModule {
    [CmdletBinding()]
    param()

    $module = Get-Module -ListAvailable -Name ImportExcel | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $module) {
        $installCommand = 'Install-Module ImportExcel -Scope CurrentUser -Repository PSGallery -Force -AllowClobber'
        Write-SpoLog -Message ("ImportExcel is not installed. Automatic installation starting: {0}" -f $installCommand) -Level WARNING
        if (-not (Get-Command Install-Module -ErrorAction SilentlyContinue)) {
            throw "ImportExcel is missing and Install-Module is unavailable. Install PowerShellGet, then run: $installCommand"
        }
        try {
            Install-Module -Name ImportExcel -Scope CurrentUser -Repository PSGallery -Force -AllowClobber -ErrorAction Stop
        }
        catch {
            throw "Automatic ImportExcel installation failed. Run '$installCommand' with the SmartM365 execution account. $($_.Exception.Message)"
        }
        $module = Get-Module -ListAvailable -Name ImportExcel | Sort-Object Version -Descending | Select-Object -First 1
        if (-not $module) { throw 'ImportExcel installation completed but the module is still unavailable in PSModulePath.' }
        Write-SpoLog -Message ("ImportExcel installed automatically: version={0}; path={1}" -f $module.Version,$module.Path) -Level SUCCESS
    }
    Import-Module -Name $module.Path -Force -ErrorAction Stop
    Write-SpoLog -Message ("ImportExcel module loaded: version={0}; path={1}" -f $module.Version,$module.Path) -Level INFO
}
function New-SpoTimestampedWorkbook {
    param(
        [Parameter(Mandatory)][object[]]$CsvFiles,
        [Parameter(Mandatory)][string]$Path
    )

    Import-Module ImportExcel -ErrorAction Stop
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
    param([Parameter(Mandatory)][string]$Title,[Parameter(Mandatory)][string]$WorstStatus,[Parameter(Mandatory)][object[]]$Alerts,[Parameter(Mandatory)][hashtable]$Summary,[Parameter(Mandatory)][object[]]$TopSites,[string]$FileLinksHtml='')
    $statusColor=switch($WorstStatus){'Critical'{'#991b1b'}'Warning'{'#92400e'}default{'#166534'}}; $statusBg=switch($WorstStatus){'Critical'{'#fee2e2'}'Warning'{'#fef3c7'}default{'#dcfce7'}}
    $alertRows=foreach($alert in @($Alerts|Sort-Object @{Expression={if($_.Severity -eq 'Critical'){0}else{1}}},Category,SiteUrl|Select-Object -First 100)){ $rowColor=if($alert.Severity -eq 'Critical'){'#fee2e2'}else{'#fef3c7'}; '<tr><td style="padding:8px;border-bottom:1px solid #e5edf5;background:{0};font-weight:700;">{1}</td><td style="padding:8px;border-bottom:1px solid #e5edf5;">{2}</td><td style="padding:8px;border-bottom:1px solid #e5edf5;word-break:break-all;">{3}</td><td style="padding:8px;border-bottom:1px solid #e5edf5;">{4}</td><td style="padding:8px;border-bottom:1px solid #e5edf5;">{5}</td></tr>' -f $rowColor,(ConvertTo-SpoHtml $alert.Severity),(ConvertTo-SpoHtml $alert.Category),(ConvertTo-SpoHtml $alert.SiteUrl),(ConvertTo-SpoHtml $alert.Value),(ConvertTo-SpoHtml $alert.Details) }
    if(-not $alertRows){$alertRows=@('<tr><td colspan="5" style="padding:10px;color:#166534;">No Warning or Critical alert detected.</td></tr>')}
    $summaryRows=foreach($key in ($Summary.Keys|Sort-Object)){ '<tr><td style="padding:8px;border-bottom:1px solid #e5edf5;color:#475569;">{0}</td><td align="right" style="padding:8px;border-bottom:1px solid #e5edf5;font-weight:700;">{1}</td></tr>' -f (ConvertTo-SpoHtml $key),(ConvertTo-SpoHtml $Summary[$key]) }
    $globalRows = @('SharePointSites','OneDriveSites','ListsProcessed','StorageUsedGB','StorageQuotaGB','CriticalAlerts','WarningAlerts','InventoryMode') | ForEach-Object { if ($Summary.Contains($_)) { '<td style="padding:10px 12px;border:1px solid #d9e2ec;background:#f8fafc;min-width:120px;"><div style="font-size:11px;color:#64748b;text-transform:uppercase;">{0}</div><div style="font-size:20px;font-weight:700;color:#0f172a;">{1}</div></td>' -f (ConvertTo-SpoHtml $_),(ConvertTo-SpoHtml $Summary[$_]) } }
    $topRows=foreach($site in @($TopSites|Select-Object -First 20)){ '<tr><td style="padding:8px;border-bottom:1px solid #e5edf5;word-break:break-all;">{0}</td><td align="right" style="padding:8px;border-bottom:1px solid #e5edf5;font-weight:700;">{1}</td><td align="right" style="padding:8px;border-bottom:1px solid #e5edf5;">{2}</td></tr>' -f (ConvertTo-SpoHtml $site.SiteUrl),(ConvertTo-SpoHtml $site.StorageUsedGB),(ConvertTo-SpoHtml $site.StorageQuotaPercent) }

    return "<div style='margin:0 0 16px 0;'><span style='display:inline-block;border-radius:999px;background:$statusBg;color:$statusColor;border:1px solid $statusColor;padding:4px 12px;font-size:12px;font-weight:700;'>$WorstStatus</span></div><h2 style='font-size:15px;margin:0 0 8px;'>Global summary</h2><table width='100%' style='border-collapse:collapse;margin-bottom:16px;'><tr>$($globalRows -join '')</tr></table><h2 style='font-size:15px;margin:0 0 8px;'>Alerts</h2><table width='100%' style='border-collapse:collapse;border:1px solid #d9e2ec;font-size:12px;margin-bottom:16px;'><tr><th align='left' style='padding:8px;background:#f8fafc;'>Severity</th><th align='left' style='padding:8px;background:#f8fafc;'>Category</th><th align='left' style='padding:8px;background:#f8fafc;'>Site</th><th align='left' style='padding:8px;background:#f8fafc;'>Value</th><th align='left' style='padding:8px;background:#f8fafc;'>Details</th></tr>$($alertRows -join '')</table><h2 style='font-size:15px;margin:0 0 8px;'>Summary</h2><table width='100%' style='border-collapse:collapse;border:1px solid #d9e2ec;font-size:12px;margin-bottom:16px;'>$($summaryRows -join '')</table><h2 style='font-size:15px;margin:0 0 8px;'>Top 20 largest sites</h2><table width='100%' style='border-collapse:collapse;border:1px solid #d9e2ec;font-size:12px;'><tr><th align='left' style='padding:8px;background:#f8fafc;'>Site URL</th><th align='right' style='padding:8px;background:#f8fafc;'>Used GB</th><th align='right' style='padding:8px;background:#f8fafc;'>Quota %</th></tr>$($topRows -join '')</table>$FileLinksHtml"
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
        Connect-MgGraph -Scopes $scopes -NoWelcome -ErrorAction Stop | Out-Null
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
        Connect-MgGraph -ClientId $appId -TenantId $tenantId -CertificateThumbprint $thumbprint -NoWelcome -ErrorAction Stop | Out-Null
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
        [string]$Operation = 'Graph paged request'
    )

    $values = New-Object System.Collections.Generic.List[object]
    $nextUri = $Uri
    while (-not [string]::IsNullOrWhiteSpace($nextUri)) {
        $page = Invoke-SpoGraphRequest -Uri $nextUri -Operation $Operation
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
    if (-not (Test-Path -LiteralPath $rawPath -PathType Leaf)) { throw "Graph report did not create expected CSV: $rawPath" }
    return @(Import-Csv -LiteralPath $rawPath)
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
    $details = 'Licensed tenant capacity is not collected in Graph-only mode. Power BI estimates capacity from M365_Licenses_Tenant.csv.'

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
$requiredSpoModules = @('Microsoft.Graph.Authentication','ImportExcel')
if ($UsePnPDeepScan -or $TenantCapacityEnabled) { $requiredSpoModules += 'PnP.PowerShell' }
Invoke-SmartM365Preflight `
    -ScriptName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) `
    -RequiredModules $requiredSpoModules | Out-Null
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module ImportExcel -ErrorAction Stop
if ($UsePnPDeepScan -or $TenantCapacityEnabled) { Import-Module PnP.PowerShell -ErrorAction Stop }

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

        $siteRows.Add([pscustomobject]@{
            RunId = $RunId
            RunDateUtc = $RunDateUtc
            TenantName = $TenantName
            SiteUrl = $siteUrl
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
            $weeklyHistorySourcePaths.Add($exportResult.LatestPath) | Out-Null
        }
    }
    if ($global:EnableWeeklyHistory -and -not (Test-SmartM365MaxItemsMode)) {
        Add-SmartM365WeeklyHistory -SourceCsvPaths $weeklyHistorySourcePaths.ToArray() -HistoryRootPath $global:WeeklyHistoryFolderPath -RetentionWeeks $global:WeeklyHistoryRetentionWeeks -HistoryLabel 'SmartM365 SharePoint Online inventory'
    }
    $workbookPath = Join-Path -Path $initializedOutput -ChildPath ("M365_SPO_Inventory_{0}.xlsx" -f $exportStamp)
    New-SpoTimestampedWorkbook -CsvFiles $timestampedCsvFiles.ToArray() -Path $workbookPath | Out-Null
    if ($global:RetentionMaxCSV -gt 0) { RemoveOldFiles -Path $initializedOutput -Filter 'M365_SPO_Inventory_*.xlsx' -KeepCount $global:RetentionMaxCSV }
    if (-not $DryRun) { Ensure-SpoSharePointUploadRecord -Path $workbookPath | Out-Null }

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
    $worstStatus = if ($criticalAlertCount -gt 0) { 'Critical' } elseif ($warningAlertCount -gt 0) { 'Warning' } else { 'OK' }
    $duration = New-TimeSpan -Start $runStart -End (Get-Date)
    $quotaMeasure = $siteArray | Measure-Object -Property StorageQuotaGB -Sum
    $usedMeasure = $siteArray | Measure-Object -Property StorageUsedGB -Sum
    $totalQuotaGb = if ($siteArray.Count -gt 0 -and $null -ne $quotaMeasure -and $null -ne $quotaMeasure.PSObject.Properties['Sum'] -and $null -ne $quotaMeasure.Sum) { [Math]::Round([double]$quotaMeasure.Sum, 2) } else { 0 }
    $totalUsedGb = if ($siteArray.Count -gt 0 -and $null -ne $usedMeasure -and $null -ne $usedMeasure.PSObject.Properties['Sum'] -and $null -ne $usedMeasure.Sum) { [Math]::Round([double]$usedMeasure.Sum, 2) } else { 0 }
    $topSites = @($siteArray | Sort-Object -Property StorageUsedGB -Descending | Select-Object -First 20)
    $summary = [ordered]@{
        Tenant = $TenantName
        RunId = $RunId
        Duration = ('{0:hh\:mm\:ss}' -f $duration)
        InventoryMode = if ($TenantCapacityEnabled -and $DeepSharingScanEnabled) { 'Graph+PnPTenantCapacity+PnPDeepScan' } elseif ($TenantCapacityEnabled) { 'Graph+PnPTenantCapacity' } elseif ($DeepSharingScanEnabled) { 'Graph+PnPDeepScan' } else { 'GraphOnly' }
        SitesProcessed = $siteArray.Count
        SharePointSites = @($siteArray | Where-Object { -not [bool]$_.IsOneDrive }).Count
        OneDriveSites = @($siteArray | Where-Object { [bool]$_.IsOneDrive }).Count
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

    $resultSummary = "SPO inventory $worstStatus; mode=$($summary.InventoryMode); sites=$($siteRows.Count); lists=$($listRows.Count); critical=$($summary.CriticalAlerts); warnings=$($summary.WarningAlerts)."
    if (-not $DryRun) {
        $prefix = if ($worstStatus -eq 'Critical') { '[CRITICAL]' } elseif ($worstStatus -eq 'Warning') { '[WARNING]' } else { '[OK]' }
        $subject = "$prefix SmartM365 SharePoint Online inventory - $TenantName"
        $mailFileLinks = New-SpoSharePointLinksHtml -Paths (@($timestampedCsvFiles.Path) + @($workbookPath))
        $bodyHtml = New-SpoHtmlSummary -Title $subject -WorstStatus $worstStatus -Alerts $alertArray -Summary $summary -TopSites $topSites -FileLinksHtml $mailFileLinks
        $dailySummaryMarkerPath = Join-Path -Path (Split-Path -Path $global:LogTextFile -Parent) -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath))-DailySummary-LastSent.txt"
        $dailySummarySent = Invoke-SpoDailySummaryMail -MarkerPath $dailySummaryMarkerPath -SendAction {
            Send-SmartM365Mail -Subject $subject -BodyHtml $bodyHtml
        }
        if ($dailySummarySent) {
            Write-SpoLog -Message ("Daily SharePoint summary email sent: {0}" -f $subject) -Level SUCCESS
        }
    }
    else {
        Write-SpoLog -Message 'DryRun enabled: daily summary email skipped.'
    }
    Write-SpoLog -Message $resultSummary -Level SUCCESS

    Complete-SmartM365ExecutionContext -Status Success
}
catch {
    $scriptError = $_
    Write-SpoLog -Message ("SPO inventory failed during {0}: {1}" -f $CurrentOperation, $_.Exception.Message) -Level ERROR
    try { Complete-SmartM365ExecutionContext -Status Failed -ErrorRecord $_ -FailureStage $CurrentOperation } catch { Write-SpoLog -Message ('Complete-SmartM365ExecutionContext failed after script error: {0}' -f $_.Exception.Message) -Level WARNING }
    throw
}
finally {
    try {
        if (Get-Command Disconnect-PnPOnline -ErrorAction SilentlyContinue) { Disconnect-PnPOnline -ErrorAction SilentlyContinue }
    }
    catch { Write-SpoLog -Message ('Disconnect-PnPOnline failed during cleanup: {0}' -f $_.Exception.Message) -Level WARNING }
    try {
        if (Get-Command Disconnect-MgGraph -ErrorAction SilentlyContinue) { Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null }
    }
    catch { Write-SpoLog -Message ('Disconnect-MgGraph failed during cleanup: {0}' -f $_.Exception.Message) -Level WARNING }
    if ($null -ne $scriptError) { exit 1 }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCClIe2WK3gYJJ28
# gZP8CTq/a+hGEhklDKxQebnmNNBu5aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIBsL9Ueh5YhQNlZD6FSSAAPaq8IzMUJL7lHjSIsb37avMA0GCSqG
# SIb3DQEBAQUABIIBgIoh4jqmjezQxu2xgMLNCkV2vG4WG4ad+AyREKCFka+BXDmc
# itdZ7bKKhC8dpYTXf5cQxCqCW5USCiT7C+VFLe1GDOgnWoykbEKJoS3L5zvK6hNt
# F1JQP0wYg9+WVeRurC1K14jJzcC9VFr19Qfhaxj46A15eppbtBFO/0nj/FlyktWv
# lzqFsZmU8m8mp1074iTPuX7rHQ+VolQfrCZTTtcPgDPSg/A5Gzvoy9sougmY5cFP
# MvX4yr//9VTnIGLnBG6+mxJUEEh0lhK9/ZN6Al1xrB14DPbHLsC1Qo4K/N2c4soM
# CtxrfquutvARTRQ1ZWAZyKaJljG5S1jbkOeNjexbZULEU/RT4Ne0M4ywltzVghq0
# E0nicHpol161CQWwaJTq+b6E8xll9x5qAI8H26XXB3QkA+GSaF21jafIgX92vzof
# 7WSpIQgGv9OunGSb7OnMHtLNrPFyTNhbX5yydPA9zh/WZ/8UlFtEgJbdyRuVkSGW
# X6Wd4/VvTHAbP6KbiaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTkyMjA0
# NTVaMC8GCSqGSIb3DQEJBDEiBCCFra1UWF1AbHqLaxiBNjwC1Oa57wF7Vn/E8JiY
# BuXmIjANBgkqhkiG9w0BAQEFAASCAgAwEa89p+xxKMYlt8ymcd7DYS6netdu7MPN
# PzjdA1+TrXVEI2Kl9C2zI/VDFb5J1dx1Z/poiyDohPT9MFunT+pMp7cAWbtsM3mz
# qVrJWwSrSRoOcsuDRZtmZhDxGFelAax6KHxLCKMfnP+BVelUBMP6pYFshOnDNt8p
# ziNZv5Ss8anljHvwYvXwGvhZ3bbIYYJ3xMRIQ8aZoSylu/uOvmZOYrldsIDnlQiM
# kcfGIVu2wAe3CKrL6NHwLgw+8WvHHH52+658c77HqrAm2n7b8NF1INkLCU0PNriV
# 7FRaUM85YEpp8I1PUeK1sej1Mp0cwyy+yTI6Zsc2rRPSXXUWX7N/VyZN1GXAv3DV
# JFS0ZTIq/nJ0vfd/5SZsCt29KPIkP2pn/vrU+kvNG2MlHWBSyB/SKkoMV4tQVOhc
# Ng8zTVXN5mRGc3ouGyNfL0bXkWYoqHt26sE1oMAX6arh9JsPf4kD3cfUjIqryTWx
# ghi/gLAOgRSqcJ+OlPahgc3/e3zX/JIFc0GOCiGMOdG050dBxDl3EGzmI9oUJYS/
# jw9uzz0lEKNLkfWQLOj2pvi+IagFRMVt2lcv6UiGysc3wgzSi9sadE3qe6y2wM7Y
# 0zalQwMokjr8FQ+Nxr0TsGtON1FR6SV+VvfpYrfcqxILRB7zlkSnSqIiw9p+0GNH
# VE2bsuliKQ==
# SIG # End signature block
