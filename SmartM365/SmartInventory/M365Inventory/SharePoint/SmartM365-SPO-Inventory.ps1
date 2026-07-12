#Requires -Version 7.0
<#
.SYNOPSIS
    SharePoint Online tenant inventory with CSV exports and HTML alert summary.

.DESCRIPTION
    Uses Microsoft Graph app-only authentication by default to inventory
    SharePoint Online sites, storage, activity, lists where Graph allows access,
    owner signals, inactive sites, orphaned sites, and tenant-level summary
    statistics. PnP.PowerShell deep sharing scans are optional and never required
    for the default inventory mode.

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
    Sends the HTML email even when no Warning/Critical condition is detected.

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

.PARAMETER InteractiveAuth
    Uses delegated interactive Graph authentication instead of app-only certificate authentication.

.VERSION
0.8


.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication; optional PnP.PowerShell only when -UsePnPDeepScan is used.
    Minimum Graph application permissions for default inventory: Reports.Read.All; Sites.Read.All; Directory.Read.All.
    Optional PnP deep sharing scan may require SharePoint site-level access for scanned sites.
    Conditional: Mail.Send is required only when Graph mail is used; Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Author: https://github.com/khda79/workplacecloudhub.com
    Requires: PowerShell 7+, Microsoft.Graph.Authentication, SmartM365.Core.psd1
    Optional: PnP.PowerShell only when -UsePnPDeepScan is used.
    Minimum Microsoft Graph application permissions for default inventory:
      - Reports.Read.All for SharePoint and OneDrive usage reports.
      - Sites.Read.All for site and list inventory.
      - Directory.Read.All for organization context and owner/directory enrichment.
    Optional PnP deep sharing scan may require SharePoint site-level access for the scanned sites. Do not grant SharePoint Administrator or Sites.FullControl.All only for the default Graph inventory.
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
$ScriptVersion = '0.8'
$CurrentOperation = 'Initialize'

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host 'This script requires PowerShell 7 or later.' -ForegroundColor Red
    Write-Host "Current PowerShell version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    exit 1
}

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

function Import-SmartM365CoreModule {
    $searchRoot = $PSScriptRoot
    while ($searchRoot) {
        $candidate = Join-Path -Path $searchRoot -ChildPath 'Modules\SmartM365.Core\SmartM365.Core.psd1'
        if (Test-Path -LiteralPath $candidate) { Import-Module $candidate -Force -ErrorAction Stop; return }
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
function Add-SpoHistoryCsv { param([AllowEmptyCollection()][object[]]$Data,[Parameter(Mandatory)][string]$Path,[Parameter(Mandatory)][string[]]$Columns) $parent=Split-Path -Path $Path -Parent; if(-not(Test-Path -LiteralPath $parent)){New-Item -Path $parent -ItemType Directory -Force|Out-Null}; $rows=@($Data|Select-Object -Property $Columns); if(-not(Test-Path -LiteralPath $Path)){ $rows|Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Delimiter ','; Ensure-SpoUtf8Bom -Path $Path; return }; $rows|Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Delimiter ',' -Append; Ensure-SpoUtf8Bom -Path $Path }
function Export-SpoEntityCsv { param([Parameter(Mandatory)][string]$BaseFileName,[AllowEmptyCollection()][object[]]$Data,[Parameter(Mandatory)][string[]]$Columns,[Parameter(Mandatory)][string]$TimestampedFolder,[Parameter(Mandatory)][string]$LatestFolder,[switch]$AppendHistoryMode) $timestamp=Get-Date -Format 'yyyyMMdd_HHmmss'; $timestampedPath=Join-Path -Path $TimestampedFolder -ChildPath ("{0}_{1}.csv" -f $BaseFileName,$timestamp); $latestPath=Join-Path -Path $LatestFolder -ChildPath ("{0}.csv" -f $BaseFileName); if(@($Data).Count -eq 0){$header='"'+($Columns -join '","')+'"'; Set-Content -LiteralPath $timestampedPath -Value $header -Encoding UTF8; Set-Content -LiteralPath $latestPath -Value $header -Encoding UTF8; Ensure-SpoUtf8Bom -Path $timestampedPath; Ensure-SpoUtf8Bom -Path $latestPath; return [pscustomobject]@{TimestampedPath=$timestampedPath;LatestPath=$latestPath}}; $result=Export-SmartM365Csv -Data $Data -TimestampedPath $timestampedPath -LatestPath $latestPath -Columns $Columns; Ensure-SpoUtf8Bom -Path $timestampedPath; Ensure-SpoUtf8Bom -Path $latestPath; if($AppendHistoryMode){$historyPath=Join-Path -Path $TimestampedFolder -ChildPath ("{0}_History.csv" -f $BaseFileName); Add-SpoHistoryCsv -Data $Data -Path $historyPath -Columns $Columns; Write-SpoLog -Message ("History CSV appended: {0}" -f $historyPath)}; return $result }
function New-SpoAlertRow { param([Parameter(Mandatory)][string]$Severity,[Parameter(Mandatory)][string]$Category,[Parameter(Mandatory)][string]$SiteUrl,[string]$ObjectName='',[string]$Metric='',[AllowNull()]$Value=$null,[string]$Threshold='',[string]$Details='') [pscustomobject]@{Severity=$Severity;Category=$Category;SiteUrl=$SiteUrl;ObjectName=$ObjectName;Metric=$Metric;Value=if($null -eq $Value){''}else{[string]$Value};Threshold=$Threshold;Details=$Details} }

function New-SpoHtmlSummary {
    param([Parameter(Mandatory)][string]$Title,[Parameter(Mandatory)][string]$WorstStatus,[Parameter(Mandatory)][object[]]$Alerts,[Parameter(Mandatory)][hashtable]$Summary,[Parameter(Mandatory)][object[]]$TopSites,[string]$LogFilePath='',[string]$CsvFolderPath='')
    $statusColor=switch($WorstStatus){'Critical'{'#991b1b'}'Warning'{'#92400e'}default{'#166534'}}; $statusBg=switch($WorstStatus){'Critical'{'#fee2e2'}'Warning'{'#fef3c7'}default{'#dcfce7'}}
    $alertRows=foreach($alert in @($Alerts|Sort-Object @{Expression={if($_.Severity -eq 'Critical'){0}else{1}}},Category,SiteUrl|Select-Object -First 100)){ $rowColor=if($alert.Severity -eq 'Critical'){'#fee2e2'}else{'#fef3c7'}; '<tr><td style="padding:8px;border-bottom:1px solid #e5edf5;background:{0};font-weight:700;">{1}</td><td style="padding:8px;border-bottom:1px solid #e5edf5;">{2}</td><td style="padding:8px;border-bottom:1px solid #e5edf5;word-break:break-all;">{3}</td><td style="padding:8px;border-bottom:1px solid #e5edf5;">{4}</td><td style="padding:8px;border-bottom:1px solid #e5edf5;">{5}</td></tr>' -f $rowColor,(ConvertTo-SpoHtml $alert.Severity),(ConvertTo-SpoHtml $alert.Category),(ConvertTo-SpoHtml $alert.SiteUrl),(ConvertTo-SpoHtml $alert.Value),(ConvertTo-SpoHtml $alert.Details) }
    if(-not $alertRows){$alertRows=@('<tr><td colspan="5" style="padding:10px;color:#166534;">No Warning or Critical alert detected.</td></tr>')}
    $summaryRows=foreach($key in ($Summary.Keys|Sort-Object)){ '<tr><td style="padding:8px;border-bottom:1px solid #e5edf5;color:#475569;">{0}</td><td align="right" style="padding:8px;border-bottom:1px solid #e5edf5;font-weight:700;">{1}</td></tr>' -f (ConvertTo-SpoHtml $key),(ConvertTo-SpoHtml $Summary[$key]) }
    $globalRows = @('SharePointSites','OneDriveSites','ListsProcessed','StorageUsedGB','StorageQuotaGB','CriticalAlerts','WarningAlerts','InventoryMode') | ForEach-Object { if ($Summary.Contains($_)) { '<td style="padding:10px 12px;border:1px solid #d9e2ec;background:#f8fafc;min-width:120px;"><div style="font-size:11px;color:#64748b;text-transform:uppercase;">{0}</div><div style="font-size:20px;font-weight:700;color:#0f172a;">{1}</div></td>' -f (ConvertTo-SpoHtml $_),(ConvertTo-SpoHtml $Summary[$_]) } }
    $topRows=foreach($site in @($TopSites|Select-Object -First 20)){ '<tr><td style="padding:8px;border-bottom:1px solid #e5edf5;word-break:break-all;">{0}</td><td align="right" style="padding:8px;border-bottom:1px solid #e5edf5;font-weight:700;">{1}</td><td align="right" style="padding:8px;border-bottom:1px solid #e5edf5;">{2}</td></tr>' -f (ConvertTo-SpoHtml $site.SiteUrl),(ConvertTo-SpoHtml $site.StorageUsedGB),(ConvertTo-SpoHtml $site.StorageQuotaPercent) }
    $technicalHtml="<div style='margin-top:18px;font-size:11px;line-height:16px;color:#64748b;'><div><strong>CSV:</strong> <span style='font-family:Consolas,monospace;word-break:break-all;'>$(ConvertTo-SpoHtml $CsvFolderPath)</span></div><div><strong>Log:</strong> <span style='font-family:Consolas,monospace;word-break:break-all;'>$(ConvertTo-SpoHtml $LogFilePath)</span></div></div>"
    return "<div style='margin:0 0 16px 0;'><span style='display:inline-block;border-radius:999px;background:$statusBg;color:$statusColor;border:1px solid $statusColor;padding:4px 12px;font-size:12px;font-weight:700;'>$WorstStatus</span></div><h2 style='font-size:15px;margin:0 0 8px;'>Global summary</h2><table width='100%' style='border-collapse:collapse;margin-bottom:16px;'><tr>$($globalRows -join '')</tr></table><h2 style='font-size:15px;margin:0 0 8px;'>Alerts</h2><table width='100%' style='border-collapse:collapse;border:1px solid #d9e2ec;font-size:12px;margin-bottom:16px;'><tr><th align='left' style='padding:8px;background:#f8fafc;'>Severity</th><th align='left' style='padding:8px;background:#f8fafc;'>Category</th><th align='left' style='padding:8px;background:#f8fafc;'>Site</th><th align='left' style='padding:8px;background:#f8fafc;'>Value</th><th align='left' style='padding:8px;background:#f8fafc;'>Details</th></tr>$($alertRows -join '')</table><h2 style='font-size:15px;margin:0 0 8px;'>Summary</h2><table width='100%' style='border-collapse:collapse;border:1px solid #d9e2ec;font-size:12px;margin-bottom:16px;'>$($summaryRows -join '')</table><h2 style='font-size:15px;margin:0 0 8px;'>Top 20 largest sites</h2><table width='100%' style='border-collapse:collapse;border:1px solid #d9e2ec;font-size:12px;'><tr><th align='left' style='padding:8px;background:#f8fafc;'>Site URL</th><th align='right' style='padding:8px;background:#f8fafc;'>Used GB</th><th align='right' style='padding:8px;background:#f8fafc;'>Quota %</th></tr>$($topRows -join '')</table>$technicalHtml"
}

function Get-SpoMailAttachmentPaths {
    param(
        [Parameter(Mandatory)][string]$LatestFolder,
        [long]$MaxTotalBytes = 2MB
    )

    $candidateNames = @(
        'M365_SPO_Sites.csv',
        'M365_SPO_ExternalSharing.csv',
        'M365_SPO_Permissions.csv',
        'M365_SPO_Lists.csv'
    )
    $attachments = New-Object System.Collections.Generic.List[string]
    [long]$totalBytes = 0
    foreach ($name in $candidateNames) {
        $candidate = Join-Path -Path $LatestFolder -ChildPath $name
        if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) { continue }
        $sizeBytes = (Get-Item -LiteralPath $candidate).Length
        if ($sizeBytes -gt 3MB) {
            Write-SpoLog -Message ("Mail attachment skipped because it exceeds Graph inline attachment limit: {0} ({1:n1} MB)." -f $candidate, ($sizeBytes / 1MB)) -Level WARNING
            continue
        }
        if (($totalBytes + $sizeBytes) -gt $MaxTotalBytes) {
            Write-SpoLog -Message ("Mail attachment skipped to keep Graph sendMail payload small: {0} ({1:n1} MB); current total {2:n1} MB; limit {3:n1} MB." -f $candidate, ($sizeBytes / 1MB), ($totalBytes / 1MB), ($MaxTotalBytes / 1MB)) -Level WARNING
            continue
        }
        $attachments.Add($candidate) | Out-Null
        $totalBytes += $sizeBytes
    }
    if ($attachments.Count -gt 0) {
        Write-SpoLog -Message ("Mail attachments selected: {0} file(s), {1:n1} MB total." -f $attachments.Count, ($totalBytes / 1MB)) -Level INFO
    }
    return @($attachments)
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

Import-SmartM365CoreModule
$requiredSpoModules = @('Microsoft.Graph.Authentication')
if ($UsePnPDeepScan) { $requiredSpoModules += 'PnP.PowerShell' }
Invoke-SmartM365Preflight `
    -ScriptName ([System.IO.Path]::GetFileNameWithoutExtension($PSCommandPath)) `
    -RequiredModules $requiredSpoModules | Out-Null
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
if ($UsePnPDeepScan) { Import-Module PnP.PowerShell -ErrorAction Stop }

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
        -GraphProbeUris @(
            'https://graph.microsoft.com/v1.0/organization?$select=id,displayName',
            'https://graph.microsoft.com/v1.0/reports/getSharePointSiteUsageDetail(period=''D7'')',
            'https://graph.microsoft.com/v1.0/sites?search=*&$top=1'
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
            Write-SpoLog -Message 'UsePnPDeepScan is enabled, but default SPO inventory remains Graph-only; no SharePoint Administrator dependency is introduced.' -Level INFO
        }
    }

    $CurrentOperation = 'Export CSV files'
    $siteColumns = @('RunId','RunDateUtc','TenantName','SiteUrl','Title','Template','CreatedUtc','LastActivityUtc','DaysSinceLastActivity','Owner','LockState','SharingCapability','ExternalSharingEnabled','StorageQuotaMB','StorageUsedMB','StorageQuotaPercent','StorageQuotaGB','StorageUsedGB','IsOneDrive','IsHubSite','HubSiteId','RelatedGroupId','IsInactive','IsOrphaned','Status','NumericValue','TextValue','Threshold','Details')
    $listColumns = @('RunId','RunDateUtc','TenantName','SiteUrl','ListId','ListTitle','ListUrl','BaseTemplate','BaseType','Hidden','ItemCount','SizeMB','VersioningEnabled','MajorVersionLimit','Status','NumericValue','TextValue','Threshold','Details')
    $permissionColumns = @('RunId','RunDateUtc','TenantName','SiteUrl','PrincipalType','PrincipalName','PrincipalLoginName','IsSiteAdmin','IsExternal','IsDisabled','Status','NumericValue','TextValue','Threshold','Details')
    $sharingColumns = @('RunId','RunDateUtc','TenantName','SiteUrl','ObjectType','ObjectTitle','ObjectUrl','SharingSignal','SharingValue','LinkScope','Principal','IsAnonymous','IsExternal','Status','NumericValue','TextValue','Threshold','Details')

    foreach ($export in @(
        @{ Base = 'M365_SPO_Sites'; Data = $siteRows.ToArray(); Columns = $siteColumns },
        @{ Base = 'M365_SPO_Lists'; Data = $listRows.ToArray(); Columns = $listColumns },
        @{ Base = 'M365_SPO_Permissions'; Data = $permissionRows.ToArray(); Columns = $permissionColumns },
        @{ Base = 'M365_SPO_ExternalSharing'; Data = $sharingRows.ToArray(); Columns = $sharingColumns }
    )) {
        $exportResult = Export-SpoEntityCsv -BaseFileName $export.Base -Data $export.Data -Columns $export.Columns -TimestampedFolder $initializedOutput -LatestFolder $latestFolder -AppendHistoryMode:$AppendHistory
        if ($exportResult.TimestampedPath) { $generatedCsvPaths.Add($exportResult.TimestampedPath) | Out-Null }
        if ($exportResult.LatestPath) { $generatedCsvPaths.Add($exportResult.LatestPath) | Out-Null }
    }

    $CurrentOperation = 'Build and send notification'
    $siteArray = @($siteRows.ToArray())
    $listArray = @($listRows.ToArray())
    $permissionArray = @($permissionRows.ToArray())
    $sharingArray = @($sharingRows.ToArray())
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
        InventoryMode = if ($DeepSharingScanEnabled) { 'GraphAppOnly+OptionalPnPDeepScan' } else { 'GraphAppOnly' }
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
        IncludeOneDrive = [bool]$IncludeOneDrive
        DeepSharingScan = $DeepSharingScanEnabled
        SharingScanItemLimitPerSite = $SharingScanItemLimitPerSite
    }

    $resultSummary = "SPO inventory $worstStatus; mode=$($summary.InventoryMode); sites=$($siteRows.Count); lists=$($listRows.Count); critical=$($summary.CriticalAlerts); warnings=$($summary.WarningAlerts)."
    $mailShouldSend = (-not $DryRun) -and ($AlwaysSend -or $worstStatus -in @('Critical','Warning'))
    if ($mailShouldSend) {
        $prefix = if ($worstStatus -eq 'Critical') { '[CRITICAL]' } elseif ($worstStatus -eq 'Warning') { '[WARNING]' } else { '[OK]' }
        $subject = "$prefix SmartM365 SharePoint Online inventory - $TenantName"
        $bodyHtml = New-SpoHtmlSummary -Title $subject -WorstStatus $worstStatus -Alerts $alertArray -Summary $summary -TopSites $topSites -LogFilePath $global:LogTextFile -CsvFolderPath $latestFolder
        $mailAttachments = Get-SpoMailAttachmentPaths -LatestFolder $latestFolder
        Send-SmartM365Mail -Subject $subject -BodyHtml $bodyHtml -Attachments $mailAttachments
        Write-SpoLog -Message ("Notification email sent: {0}" -f $subject) -Level SUCCESS
    }
    elseif ($DryRun) {
        Write-SpoLog -Message 'DryRun enabled: email notification skipped.'
    }
    else {
        Write-SpoLog -Message 'No Warning/Critical alert detected and AlwaysSend is not enabled: email notification skipped.'
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
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBTwPlmFZAt63eQ
# n8QLFHzqkz1tXGlkfXxiaJ7tWgQYYaCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCBIAOg9AVGkE8+RgGoWjtCgqPfXqGskxK5pj5ByOco8pjANBgkqhkiG9w0B
# AQEFAASCAYBvSaTBrIOpqJgdEp0DuiL2lZFjIu5tBphEO6Eyt93RPFHZR72Bf3KO
# HZUgoxcRJcUQZd/ZfqsGRLw3sm9+9jVMugBA9ZGhhaMk1kmJwj2MbhwSH79KHEsR
# Oaayz75sQgOukHPLZ3QyPiw9CI7TtI1eOIqe0hcIqy4ro4eWfXrpQjIed30wArGL
# IVAo/hL5vVRPaKRSc6CjdrHgfPcq6rJlGMn/rpQzsDbH8u+4LcYpqfG+m4i0M7CS
# Dqo9enWKdCN8eRyY6QeAABPKJ+2FS30DWF7evqKf4wyXS3HTVIwa0FaPTLei+zMt
# qQGNo65vjz7vyNjcQX5hhawFIqdz8hMcz0GQO5+vviR85TeynUoLrFAMAN6VlbkG
# QQfPOsXY4IK5F2jYlDPxnQj3CqxCPE9JcoCX8UcxVc5XbzSJEdfDtzJDrEJAMZkK
# lj9v0hoRRpGranwogSwKsOdX+ab73hTRZQy6JJ/O7jsjZtD2CbW1KINwo9UkAJuW
# xWR4bMsTVwQ=
# SIG # End signature block
