<#
.SYNOPSIS
    Active Directory forest health check for PowerShell 7 and RSAT ActiveDirectory.
.VERSION
    1.0.17
.DESCRIPTION
    Discovers every domain with Get-ADForest, audits domain controllers and domain health,
    exports a flat Power BI-ready CSV, and sends an HTML summary email on warnings or critical alerts.
    Minimum permissions: PowerShell 7+, RSAT ActiveDirectory module, and read access to the AD forest/domains. Remote DC admin checks require explicit T0/admin rights only when -EnableRemoteDcAdminChecks is used.

.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; ActiveDirectory RSAT/Windows Server module.
    Minimum permissions: read access to the AD forest/domains through Get-ADForest, Get-ADDomain, Get-ADDomainController, Get-ADReplication* and related read-only AD cmdlets.
    Optional: -EnableRemoteDcAdminChecks requires explicit tier-0/admin rights for remote DC service, event, disk and AD database checks.
    Conditional: Mail.Send is required only when Graph mail is used; Sites.Selected write is required only when SharePoint upload is enabled.
#>
#requires -Version 7.0
#requires -Modules ActiveDirectory
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSReviewUnusedParameter','',Justification='Parameters are consumed by script-scope helper functions.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions','',Justification='This read-only health script does not change AD state.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='Final console status is intentional for command-line use.')]
[CmdletBinding()]
param(
    [string]$Tenant = 'test',
    [string]$OutputFolder = '',
    [switch]$AppendHistory,
    [string]$HistoryCsvPath,
    [string[]]$To = @(),
    [string]$From = '',
    [string]$SmtpServer = '',
    [int]$SmtpPort = 25,
    [switch]$UseSsl,
    [pscredential]$SmtpCredential,
    [switch]$AlwaysSend,
    [int]$RetryCount = 2,
    [int]$RetryDelaySeconds = 3,
    [int]$TcpTimeoutMs = 2000,
    [double]$TimeOffsetWarningMinutes = 5,
    [double]$DiskFreePercentCritical = 15,
    [double]$DiskFreeGbCritical = 5,
    [int]$ReplicationDelayWarningHours = 24,
    [int]$DfsrBacklogWarningCount = 100,
    [int]$DfsrBacklogCriticalCount = 1000,
    [switch]$SkipDfsrBacklog,
    [switch]$EnableRemoteDcAdminChecks,
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
    throw "-MaxItems is not supported by SmartM365-ActiveDirectory-HealthCheck because it produces a forest/domain health report where partial data would be misleading."
}
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[System.Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::InvariantCulture
[System.Threading.Thread]::CurrentThread.CurrentUICulture = [Globalization.CultureInfo]::InvariantCulture
$RunStarted = Get-Date
$RunDateUtc = $RunStarted.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ',[Globalization.CultureInfo]::InvariantCulture)
$RunId = [guid]::NewGuid().ToString()
$Rows = [System.Collections.ArrayList]::new()
$DomainFacts = [System.Collections.ArrayList]::new()
$ScriptBaseName = [IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ScriptVersion = "1.0.17"
$TaskName = "$ScriptBaseName v$ScriptVersion"
$TenantContextPath = & {
    $d = $PSScriptRoot
    while ($d) {
        foreach ($candidate in @((Join-Path $d 'SmartM365-TenantContext.ps1'), (Join-Path $d 'Config\SmartM365-TenantContext.ps1'))) {
            if (Test-Path -LiteralPath $candidate) { return $candidate }
        }
        $parent = Split-Path -Path $d -Parent
        if ([string]::IsNullOrWhiteSpace($parent) -or $parent -eq $d) { break }
        $d = $parent
    }
    throw 'SmartM365-TenantContext.ps1 not found.'
}
. $TenantContextPath
$TenantContext = Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$TenantContextDirectory = Split-Path -Path $TenantContextPath -Parent
$SmartM365ProjectRoot = if ((Split-Path -Path $TenantContextDirectory -Leaf) -ieq 'Config') {
    Split-Path -Path $TenantContextDirectory -Parent
}
else {
    $TenantContextDirectory
}
$CoreModulePath = Join-Path $SmartM365ProjectRoot 'Modules\SmartM365.Core\SmartM365.Core.psd1'
Import-Module -Name $CoreModulePath -MinimumVersion '1.0.24' -Force -ErrorAction Stop
$LocalConfigPath = Join-Path $PSScriptRoot ("$ScriptBaseName.local.json")
$LegacyLocalConfigPath = Join-Path $PSScriptRoot 'AD-HealthCheck.local.json'
$LocalTemplatePath = "$LocalConfigPath.template"
if (-not (Test-Path -LiteralPath $LocalConfigPath) -and (Test-Path -LiteralPath $LegacyLocalConfigPath)) {
    $LocalConfigPath = $LegacyLocalConfigPath
    $LocalTemplatePath = "$LocalConfigPath.template"
}
if (-not (Test-Path -LiteralPath $LocalConfigPath)) {
    if (Get-Command Initialize-SmartM365LocalJsonFromTemplate -ErrorAction SilentlyContinue) {
        Initialize-SmartM365LocalJsonFromTemplate -Path $LocalConfigPath -TemplatePath $LocalTemplatePath -ConfigDescription 'script local configuration' | Out-Null
    }
    elseif (Test-Path -LiteralPath $LocalTemplatePath) {
        Copy-Item -LiteralPath $LocalTemplatePath -Destination $LocalConfigPath -ErrorAction Stop
    }
}
$ScriptConfig = if (Test-Path -LiteralPath $LocalConfigPath) {
    $config = Get-Content -LiteralPath $LocalConfigPath -Raw | ConvertFrom-Json
    if (Get-Command Sync-SmartM365JsonConfigWithTemplate -ErrorAction SilentlyContinue) {
        Sync-SmartM365JsonConfigWithTemplate -Config $config -Path $LocalConfigPath
    }
    else {
        $config
    }
}
else {
    [pscustomobject]@{}
}
function Resolve-ConfigToken([AllowNull()][object]$Value) {
    if ($Value -isnot [string]) { return $Value }
    $resolved = $Value
    for ($i = 0; $i -lt 10; $i++) {
        $tokenMatches = [regex]::Matches($resolved, '\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}')
        if ($tokenMatches.Count -eq 0) { break }
        foreach ($match in $tokenMatches) {
            $name = $match.Groups['Name'].Value
            $prop = $TenantContext.PSObject.Properties[$name]
            if ($prop -and $null -ne $prop.Value) { $resolved = $resolved.Replace($match.Value, [string]$prop.Value) }
        }
    }
    return $resolved
}
function Get-LocalConfigValue([string]$Name, [AllowNull()][object]$DefaultValue) {
    $prop = $ScriptConfig.PSObject.Properties[$Name]
    if ($prop -and $null -ne $prop.Value) {
        if ($prop.Value -isnot [string] -or ($prop.Value.Trim() -and $prop.Value.Trim() -notin @('__USE_GLOBAL__','USE_GLOBAL'))) {
            return Resolve-ConfigToken $prop.Value
        }
    }
    $ctxProp = $TenantContext.PSObject.Properties[$Name]
    if ($ctxProp -and $null -ne $ctxProp.Value) { return Resolve-ConfigToken $ctxProp.Value }
    return Resolve-ConfigToken $DefaultValue
}
$global:AppId = [string](Get-LocalConfigValue 'AppId' '')
$global:TenantId = [string](Get-LocalConfigValue 'TenantId' '')
$global:Thumb = [string](Get-LocalConfigValue 'Thumb' '')
$global:Thumbprint = [string](Get-LocalConfigValue 'Thumbprint' $global:Thumb)
$global:EnableSharePointUpload = [bool](Get-LocalConfigValue 'EnableSharePointUpload' $false)
$global:SharePointSiteHostname = [string](Get-LocalConfigValue 'SharePointSiteHostname' '')
$global:SharePointSitePath = [string](Get-LocalConfigValue 'SharePointSitePath' '')
$global:SharePointLibraryDisplayName = [string](Get-LocalConfigValue 'SharePointLibraryDisplayName' 'Documents')
$global:SharePointTargetFolderPath = [string](Get-LocalConfigValue 'SharePointTargetFolderPath' '')
if ([string]::IsNullOrWhiteSpace($OutputFolder)) { $OutputFolder = [string](Get-LocalConfigValue 'ADHealthCheckCsvLogFolderPath' '{{DataAllRootPath}}\ActiveDirectory\HealthCheck') }
$LatestCsvFolderPath = [string](Get-LocalConfigValue 'LatestCsvFolderPath' $TenantContext.LatestCsvFolderPath)
if (-not $PSBoundParameters.ContainsKey('AlwaysSend')) { $AlwaysSend = [bool](Get-LocalConfigValue 'AlwaysSend' $false) }
if (-not $PSBoundParameters.ContainsKey('AppendHistory')) { $AppendHistory = [bool](Get-LocalConfigValue 'AppendHistory' $false) }
if ([string]::IsNullOrWhiteSpace($HistoryCsvPath)) { $HistoryCsvPath = [string](Get-LocalConfigValue 'HistoryCsvPath' '{{DataAllRootPath}}\ActiveDirectory\HealthCheck\AD_HealthCheck_History.csv') }
if (-not $PSBoundParameters.ContainsKey('EnableRemoteDcAdminChecks')) { $EnableRemoteDcAdminChecks = [bool](Get-LocalConfigValue 'EnableRemoteDcAdminChecks' $false) }
if ([string]::IsNullOrWhiteSpace($From)) { $From = [string](Get-LocalConfigValue 'From' '') }
if ([string]::IsNullOrWhiteSpace($SmtpServer)) { $SmtpServer = [string](Get-LocalConfigValue 'SmtpServer' '') }
$SendMailMode = [string](Get-LocalConfigValue 'SendMailMode' '')
$Cc = [string](Get-LocalConfigValue 'Cc' '')
if ($To.Count -eq 0) {
    $configuredTo = [string](Get-LocalConfigValue 'To' '')
    if ([string]::IsNullOrWhiteSpace($configuredTo)) { $configuredTo = [string](Get-LocalConfigValue 'ErrorMailTo' '') }
    if (-not [string]::IsNullOrWhiteSpace($configuredTo)) { $To = @($configuredTo -split '[;,]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
}
function ConvertTo-IsoUtc([datetime]$d){ if($null -eq $d -or $d -eq [datetime]::MinValue){''}else{$d.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ',[Globalization.CultureInfo]::InvariantCulture)} }
function Num($v){ if($null -eq $v -or [string]::IsNullOrWhiteSpace([string]$v)){''}else{ try{([double]$v).ToString('0.########',[Globalization.CultureInfo]::InvariantCulture)}catch{[string]$v} } }
function Ms($s){ [int64][math]::Max(0,((Get-Date)-$s).TotalMilliseconds) }
function Add-Row{
    param(
        [AllowNull()][object]$Forest,
        [AllowNull()][object]$Domain,
        [AllowNull()][object]$DC,
        [string]$Category,
        [string]$Check,
        [ValidateSet('OK','Warning','Critical')][string]$Status,
        [AllowNull()][object]$NumericValue,
        [AllowNull()][object]$TextValue,
        [AllowNull()][object]$Threshold,
        [AllowNull()][object]$Details,
        [int64]$DurationMs
    )
    $row = [ordered]@{
        RunId = $RunId
        RunDateUtc = $RunDateUtc
        Forest = [string]$Forest
        Domain = [string]$Domain
        DC = [string]$DC
        Category = $Category
        Check = $Check
        Status = $Status
        NumericValue = (Num $NumericValue)
        TextValue = [string]$TextValue
        Threshold = [string]$Threshold
        Details = [string]$Details
        DurationMs = $DurationMs
    }
    [void]$Rows.Add([pscustomobject]$row)
}
function Get-RowSnapshot{
    @($Rows.ToArray())
}
function Invoke-Retry([scriptblock]$ScriptBlock){
    for($i=1;$i -le [math]::Max(1,$RetryCount);$i++){
        try{ return & $ScriptBlock }catch{ $m=[string]$_.Exception.Message; $transient=$m -match 'timeout|temporar|server is not operational|RPC server|unavailable|busy|could not be contacted|network path|WinRM|timed out'; if(-not $transient -or $i -ge $RetryCount){throw}; Start-Sleep -Seconds $RetryDelaySeconds }
    }
}
function Test-Port([string]$Computer,[int]$Port){
    $c=[Net.Sockets.TcpClient]::new(); try{ $a=$c.BeginConnect($Computer,$Port,$null,$null); if(-not $a.AsyncWaitHandle.WaitOne($TcpTimeoutMs,$false)){return $false}; $c.EndConnect($a); $true }catch{$false}finally{try{$c.Close()}catch{ $null = $_ }}
}
function Worst($r){ if($r|Where-Object Status -eq Critical){'Critical'}elseif($r|Where-Object Status -eq Warning){'Warning'}else{'OK'} }
function Rank($s){ if($s -eq 'Critical'){0}elseif($s -eq 'Warning'){1}else{2} }
function ConvertTo-HtmlSafe($s){ [Net.WebUtility]::HtmlEncode([string]$s) }
function Get-ObjectPropertyValue([object]$Object,[string[]]$Names){
    foreach($name in $Names){
        $prop = if($null -ne $Object){$Object.PSObject.Properties[$name]}else{$null}
        if($prop -and $null -ne $prop.Value){return [string]$prop.Value}
    }
    ''
}
function Get-TimeOffsetMinute([string]$DC){
    $out=& w32tm.exe /stripchart /computer:$DC /samples:3 /dataonly 2>&1
    $vals=@(); foreach($line in $out){ foreach($m in [regex]::Matches([string]$line,'([+-]?\d+(?:[\.,]\d+)?)s')){ $raw=$m.Groups[1].Value.Replace(',','.'); $d=0.0; if([double]::TryParse($raw,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$d)){$vals += [math]::Abs($d)} } }
    if($vals.Count -eq 0){ return [pscustomobject]@{ OffsetMinutes = $null; Error = "Unable to parse w32tm output: $($out -join ' | ')" } }
    [pscustomobject]@{ OffsetMinutes = (($vals|Measure-Object -Maximum).Maximum/60); Error = '' }
}
function Get-ADGroupMembersBySid([string]$Sid,[string]$Server){
    $group = Get-ADGroup -Identity $Sid -Server $Server -ErrorAction SilentlyContinue
    if (-not $group) { return @() }
    @(Get-ADGroupMember -Identity $group.DistinguishedName -Recursive -Server $Server -ErrorAction SilentlyContinue)
}
function Get-ADDbVolume([string]$DC){
    try{ $v=Invoke-Command -ComputerName $DC -ScriptBlock { $o=& ntdsutil.exe 'activate instance ntds' 'files' 'info' 'quit' 'quit' 2>&1; $l=$o|Where-Object{[string]$_ -match 'Database directory|DB Path|Database path'}|Select-Object -First 1; if($l -and [string]$l -match '([A-Za-z]:)'){$matches[1]}else{$env:SystemDrive} } -ErrorAction Stop; if($v){return ([string]$v).TrimEnd('\')} }catch{ $null = $_ }
    $os=Get-CimInstance Win32_OperatingSystem -ComputerName $DC -ErrorAction Stop; ([string]$os.SystemDrive).TrimEnd('\')
}
function Get-Disk([string]$DC,[string]$DeviceId){
    $d=Get-CimInstance Win32_LogicalDisk -ComputerName $DC -Filter "DeviceID='$DeviceId'" -ErrorAction Stop
    if($null -eq $d -or [double]$d.Size -le 0){throw "Logical disk $DeviceId not found"}
    [pscustomobject]@{FreeGb=[math]::Round(([double]$d.FreeSpace/1GB),2);FreePct=[math]::Round(([double]$d.FreeSpace/[double]$d.Size*100),2);SizeGb=[math]::Round(([double]$d.Size/1GB),2)}
}
function Get-DfsrBacklog([string]$Src,[string]$Dst){
    if(-not (Get-Command dfsrdiag.exe -ErrorAction SilentlyContinue)){return $null}
    $o=& dfsrdiag.exe backlog /rgname:'Domain System Volume' /rfname:'SYSVOL Share' /smem:$Src /rmem:$Dst 2>&1
    foreach($l in $o){$t=[string]$l; if($t -match 'Backlog File Count\s*:\s*(\d+)'){return [int]$matches[1]}; if($t -match 'No Backlog'){return 0}}
    $null
}
function Send-ReportMail([string]$Subject,[string]$Body){
    $mailParams = @{ Subject = $Subject; BodyHtml = $Body; VerboseLog = $true }
    if (-not [string]::IsNullOrWhiteSpace($SmtpServer)) { $mailParams.SmtpServer = $SmtpServer }
    if (-not [string]::IsNullOrWhiteSpace($SendMailMode)) { $mailParams.SendMailMode = $SendMailMode }
    if (-not [string]::IsNullOrWhiteSpace($From)) { $mailParams.From = $From }
    if ($To.Count -gt 0) { $mailParams.To = ($To -join ';') }
    if (-not [string]::IsNullOrWhiteSpace($Cc)) { $mailParams.Cc = $Cc }
    if ($SmtpPort -ne 25) { $mailParams.SmtpPort = $SmtpPort }
    SendEmailHtmlReport @mailParams
}
function ConvertTo-ReportHtml([object[]]$r,[string]$status,[datetime]$started,[datetime]$ended,[string]$csv,[AllowNull()][object]$forestInfo=$null,[bool]$RemoteDcAdminChecksEnabled=$false){
    $r = @($r)
    $facts = @($DomainFacts.ToArray())
    $c = @{OK='#107c10';Warning='#ff8c00';Critical='#d13438'}[$status]
    $criticalRows = @($r | Where-Object Status -eq Critical)
    $warningRows = @($r | Where-Object Status -eq Warning)
    $okRows = @($r | Where-Object Status -eq OK)
    $find = @($r | Where-Object Status -ne OK | Sort-Object @{Expression={Rank $_.Status}},Domain,DC,Category,Check | Select-Object -First 200)
    $forestName = if ($forestInfo -and $forestInfo.Name) { [string]$forestInfo.Name } else { [string](@($r | Where-Object Forest | Select-Object -ExpandProperty Forest -First 1)[0]) }
    $rootDomain = if ($forestInfo -and $forestInfo.RootDomain) { [string]$forestInfo.RootDomain } else { [string](@($r | Where-Object Domain | Select-Object -ExpandProperty Domain -First 1)[0]) }
    $forestMode = if ($forestInfo -and $forestInfo.ForestMode) { [string]$forestInfo.ForestMode } else { 'NotAvailable' }
    $domains = if ($forestInfo -and $forestInfo.Domains) { @($forestInfo.Domains | Sort-Object) } else { @($r | Where-Object Domain | Select-Object -ExpandProperty Domain -Unique | Sort-Object) }
    $domainBadges = (@($domains) | ForEach-Object { '<span class="tag">' + (ConvertTo-HtmlSafe $_) + '</span>' }) -join ' '
    $dcNames = @($r | Where-Object { $_.Category -eq 'Connectivity' -and $_.Check -eq 'Ping' -and $_.DC } | Select-Object -ExpandProperty DC -Unique | Sort-Object)
    $gcCount = if ($facts.Count -gt 0) { ($facts | Measure-Object -Property GlobalCatalogCount -Sum).Sum } else { '' }
    $lockedTotal = 0
    foreach ($row in @($r | Where-Object { $_.Category -eq 'DomainStats' -and $_.Check -eq 'LockedUserAccounts' })) {
        $value = 0.0
        if ([double]::TryParse([string]$row.NumericValue,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$value)) { $lockedTotal += [int]$value }
    }
    $domainsWithCritical = @($criticalRows | Where-Object Domain | Select-Object -ExpandProperty Domain -Unique).Count
    $dcsWithCritical = @($criticalRows | Where-Object DC | Select-Object -ExpandProperty DC -Unique).Count
    $unreachableDcs = @($criticalRows | Where-Object { $_.Category -eq 'Connectivity' -and $_.DC } | Select-Object -ExpandProperty DC -Unique).Count
    $replicationCritical = @($criticalRows | Where-Object Category -eq Replication).Count
    $sysvolCritical = @($criticalRows | Where-Object { $_.Category -eq 'SYSVOL' -and $_.Check -in @('SYSVOL','NETLOGON') }).Count
    $b = [Text.StringBuilder]::new()
    [void]$b.AppendLine('<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:Segoe UI,Arial;background:#f5f8fb;color:#1f2937;padding:24px}.card{background:#fff;border:1px solid #dde7f0;border-radius:8px;padding:16px;margin:0 0 16px}.muted{color:#52657a}.small{font-size:11px}.metric{font-size:22px;font-weight:700;color:#0f172a}.tag{display:inline-block;border:1px solid #cbd8e6;background:#f8fbfe;border-radius:999px;padding:3px 8px;margin:2px;font-size:11px}.pill{color:#fff;border-radius:999px;padding:4px 10px;font-weight:600}.grid{width:100%;border-collapse:separate;border-spacing:8px}.grid td{border:1px solid #dde7f0;background:#f8fbfe;border-radius:6px;padding:10px}.riskCritical{color:#b42318;font-weight:700}.riskWarning{color:#b54708;font-weight:700}.riskOK{color:#067647;font-weight:700}table{border-collapse:collapse;width:100%}th,td{border:1px solid #dde7f0;padding:7px;font-size:12px;text-align:left;vertical-align:top}th{background:#eef6fc}.rowOK{background:#f3fbf3}.rowWarning{background:#fff7e6}.rowCritical{background:#fde7e9}</style></head><body>')
    [void]$b.AppendLine(("<div class='card'><h1>Active Directory Health Check <span class='pill' style='background:{0}'>{1}</span></h1><p class='muted'>RunId: {2}<br>Machine: {3}<br>Started UTC: {4}<br>Ended UTC: {5}<br>Duration: {6}</p></div>" -f $c,$status,(ConvertTo-HtmlSafe $RunId),(ConvertTo-HtmlSafe $env:COMPUTERNAME),(ConvertTo-IsoUtc $started),(ConvertTo-IsoUtc $ended),(ConvertTo-HtmlSafe ((New-TimeSpan -Start $started -End $ended).ToString()))))
    [void]$b.AppendLine(("<div class='card'><h2>Forest recap</h2><table><tr><th>Forest</th><td>{0}</td><th>Root domain</th><td>{1}</td></tr><tr><th>Forest mode</th><td>{2}</td><th>Remote DC admin checks</th><td>{3}</td></tr><tr><th>Domains</th><td>{4}</td><th>Domain controllers</th><td>{5}</td></tr><tr><th>Global catalogs</th><td>{6}</td><th>Domains scanned</th><td>{7}</td></tr></table></div>" -f (ConvertTo-HtmlSafe $forestName),(ConvertTo-HtmlSafe $rootDomain),(ConvertTo-HtmlSafe $forestMode),$(if($RemoteDcAdminChecksEnabled){'Enabled'}else{'Disabled'}),@($domains).Count,$dcNames.Count,$gcCount,$domainBadges))
    [void]$b.AppendLine(("<div class='card'><h2>Operational recap</h2><table class='grid'><tr><td><div class='metric'>{0}</div><div>Total checks</div></td><td><div class='metric riskCritical'>{1}</div><div>Critical</div></td><td><div class='metric riskWarning'>{2}</div><div>Warning</div></td><td><div class='metric riskOK'>{3}</div><div>OK</div></td></tr><tr><td><div class='metric'>{4}</div><div>Domains with critical findings</div></td><td><div class='metric'>{5}</div><div>DCs with critical findings</div></td><td><div class='metric'>{6}</div><div>Unreachable DCs</div></td><td><div class='metric'>{7}</div><div>Locked user accounts</div></td></tr><tr><td><div class='metric'>{8}</div><div>Replication critical rows</div></td><td><div class='metric'>{9}</div><div>SYSVOL/NETLOGON critical rows</div></td><td><div class='metric'>{10}</div><div>Total DCs</div></td><td><div class='metric'>{11}</div><div>Domains</div></td></tr></table></div>" -f $r.Count,$criticalRows.Count,$warningRows.Count,$okRows.Count,$domainsWithCritical,$dcsWithCritical,$unreachableDcs,$lockedTotal,$replicationCritical,$sysvolCritical,$dcNames.Count,@($domains).Count))
    [void]$b.AppendLine('<div class="card"><h2>FSMO roles</h2><table><tr><th>Scope</th><th>Domain</th><th>Role</th><th>Holder</th><th>Status</th></tr>')
    foreach ($x in @($r | Where-Object Category -eq FSMO | Sort-Object Domain,Check)) { $scope = if ($x.Check -in @('SchemaMaster','DomainNamingMaster')) { 'Forest' } else { 'Domain' }; [void]$b.AppendLine(("<tr class='row{0}'><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{0}</td></tr>" -f (ConvertTo-HtmlSafe $x.Status),(ConvertTo-HtmlSafe $scope),(ConvertTo-HtmlSafe $x.Domain),(ConvertTo-HtmlSafe $x.Check),(ConvertTo-HtmlSafe $x.DC))) }
    [void]$b.AppendLine('</table></div>')
    [void]$b.AppendLine('<div class="card"><h2>Priority findings</h2><table><tr><th>Type</th><th>Name</th><th>Critical</th><th>Warning</th></tr>')
    $priorityGroups = @()
    $priorityGroups += @($r | Where-Object DC | Group-Object DC | ForEach-Object { $it = @($_.Group); [pscustomobject]@{ Type='DC'; Name=$_.Name; Critical=@($it | Where-Object Status -eq Critical).Count; Warning=@($it | Where-Object Status -eq Warning).Count } })
    $priorityGroups += @($r | Group-Object Category | ForEach-Object { $it = @($_.Group); [pscustomobject]@{ Type='Category'; Name=$_.Name; Critical=@($it | Where-Object Status -eq Critical).Count; Warning=@($it | Where-Object Status -eq Warning).Count } })
    foreach ($g in @($priorityGroups | Where-Object { $_.Critical -gt 0 -or $_.Warning -gt 0 } | Sort-Object @{Expression='Critical';Descending=$true},@{Expression='Warning';Descending=$true},Type,Name | Select-Object -First 20)) { [void]$b.AppendLine(("<tr><td>{0}</td><td>{1}</td><td class='riskCritical'>{2}</td><td class='riskWarning'>{3}</td></tr>" -f (ConvertTo-HtmlSafe $g.Type),(ConvertTo-HtmlSafe $g.Name),$g.Critical,$g.Warning)) }
    [void]$b.AppendLine('</table></div>')
    [void]$b.AppendLine('<div class="card"><h2>Domain status summary</h2><table><tr><th>Domain</th><th>DCs</th><th>OK</th><th>Warning</th><th>Critical</th><th>Locked users</th><th>FSMO</th><th>Replication</th><th>SYSVOL</th></tr>')
    foreach ($g in ($r | Where-Object Domain | Group-Object Domain | Sort-Object Name)) { $it = @($g.Group); $fact = @($facts | Where-Object Domain -eq $g.Name | Select-Object -First 1)[0]; $locked = @($it | Where-Object { $_.Category -eq 'DomainStats' -and $_.Check -eq 'LockedUserAccounts' } | Select-Object -First 1)[0]; $lockedValue = if ($locked) { $locked.NumericValue } else { '' }; $fsmoWorst = Worst @($it | Where-Object Category -eq FSMO); $repWorst = Worst @($it | Where-Object Category -eq Replication); $sysvolWorst = Worst @($it | Where-Object Category -eq SYSVOL); $dcCount = if ($fact) { $fact.DCCount } else { @($it | Where-Object DC | Select-Object -ExpandProperty DC -Unique).Count }; [void]$b.AppendLine(("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td class='risk{6}'>{6}</td><td class='risk{7}'>{7}</td><td class='risk{8}'>{8}</td></tr>" -f (ConvertTo-HtmlSafe $g.Name),$dcCount,@($it | Where-Object Status -eq OK).Count,@($it | Where-Object Status -eq Warning).Count,@($it | Where-Object Status -eq Critical).Count,(ConvertTo-HtmlSafe $lockedValue),$fsmoWorst,$repWorst,$sysvolWorst)) }
    [void]$b.AppendLine('</table></div>')
    [void]$b.AppendLine('<div class="card"><h2>Critical and warning findings</h2><table><tr><th>Status</th><th>Domain</th><th>DC</th><th>Category</th><th>Check</th><th>Value</th><th>Threshold</th><th>Details</th></tr>')
    if ($find.Count -eq 0) { [void]$b.AppendLine('<tr class="rowOK"><td colspan="8">No critical or warning findings.</td></tr>') }
    foreach ($x in $find) { [void]$b.AppendLine(("<tr class='row{0}'><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td></tr>" -f (ConvertTo-HtmlSafe $x.Status),(ConvertTo-HtmlSafe $x.Domain),(ConvertTo-HtmlSafe $x.DC),(ConvertTo-HtmlSafe $x.Category),(ConvertTo-HtmlSafe $x.Check),(ConvertTo-HtmlSafe $x.NumericValue),(ConvertTo-HtmlSafe $x.Threshold),(ConvertTo-HtmlSafe $x.Details))) }
    [void]$b.AppendLine('</table></div><div class="card small"><h2>Technical files</h2><p>CSV: ' + (ConvertTo-HtmlSafe $csv) + '</p></div></body></html>')
    $b.ToString()
}
function Invoke-DcCheck($ForestName,$DomainName,$DC,$AllDcs){
    $dcName=if($DC.HostName){[string]$DC.HostName}else{[string]$DC.Name}
    foreach($port in 389,3268,445){$s=Get-Date;try{$ok=Invoke-Retry {Test-Port $dcName $port};$stat=if($ok -or ($port -eq 3268 -and -not $DC.IsGlobalCatalog)){'OK'}else{'Critical'};Add-Row $ForestName $DomainName $dcName Connectivity "TCP $port" $stat ([int]$ok) "Reachable=$ok" 'Open' "TCP port $port" (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName Connectivity "TCP $port" Critical 0 Error Open $_.Exception.Message (Ms $s)}}
    $s=Get-Date;try{$ok=Invoke-Retry {Test-Connection -ComputerName $dcName -Count 1 -Quiet -ErrorAction Stop};Add-Row $ForestName $DomainName $dcName Connectivity Ping $(if($ok){'OK'}else{'Critical'}) ([int]$ok) "Reachable=$ok" Reachable 'ICMP echo test' (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName Connectivity Ping Critical 0 Error Reachable $_.Exception.Message (Ms $s)}
    if ($EnableRemoteDcAdminChecks) {
        foreach($svcName in 'NTDS','DNS','Netlogon','KDC','W32Time','DFSR') {
            $s=Get-Date;try{$svc=Invoke-Retry {Get-Service -ComputerName $dcName -Name $svcName -ErrorAction Stop};$ok=$svc.Status -eq 'Running';Add-Row $ForestName $DomainName $dcName Services $svcName $(if($ok){'OK'}else{'Critical'}) ([int]$ok) $svc.Status Running "Service $svcName" (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName Services $svcName Critical 0 Error Running $_.Exception.Message (Ms $s)}
        }
    }
    else {
        foreach($svcName in 'NTDS','DNS','Netlogon','KDC','W32Time','DFSR') {
            $s=Get-Date;Add-Row $ForestName $DomainName $dcName Services $svcName OK '' NotMeasured 'Requires T0 remote DC admin' 'Skipped by default to avoid remote service-control access to domain controllers. Use -EnableRemoteDcAdminChecks with a T0 account.' (Ms $s)
        }
    }
$s=Get-Date;$repErrors=$null;$f=@(Get-ADReplicationFailure -Target $dcName -Scope Server -ErrorAction SilentlyContinue -ErrorVariable repErrors);if($repErrors){Add-Row $ForestName $DomainName $dcName Replication ReplicationFailures Critical '' Error 0 ([string]$repErrors[0].Exception.Message) (Ms $s)}else{$det=($f|Select-Object -First 5|ForEach-Object{"$($_.Partner): $($_.FailureCount) since $(ConvertTo-IsoUtc $_.FirstFailureTime)"}) -join '; ';Add-Row $ForestName $DomainName $dcName Replication ReplicationFailures $(if($f.Count -eq 0){'OK'}else{'Critical'}) $f.Count "Failures=$($f.Count)" 0 $det (Ms $s)}
$s=Get-Date;$metaErrors=$null;$m=@(Get-ADReplicationPartnerMetadata -Target $dcName -Scope Server -ErrorAction SilentlyContinue -ErrorVariable metaErrors);if($metaErrors){Add-Row $ForestName $DomainName $dcName Replication MaxLastSuccessDelayHours Critical '' Error "<= $ReplicationDelayWarningHours hours" ([string]$metaErrors[0].Exception.Message) (Ms $s)}else{$old=@($m|Where-Object{$_.LastReplicationSuccess -and $_.LastReplicationSuccess -ne [datetime]::MinValue}|Sort-Object LastReplicationSuccess|Select-Object -First 1)[0];$max=0;if($old){$max=[math]::Round(((Get-Date).ToUniversalTime()-$old.LastReplicationSuccess.ToUniversalTime()).TotalHours,2)};$stat=if($max -gt $ReplicationDelayWarningHours){'Warning'}else{'OK'};Add-Row $ForestName $DomainName $dcName Replication MaxLastSuccessDelayHours $stat $max $(if($old){ConvertTo-IsoUtc $old.LastReplicationSuccess}else{''}) "<= $ReplicationDelayWarningHours hours" "Partners=$($m.Count)" (Ms $s)}
    foreach($sh in 'SYSVOL','NETLOGON'){$s=Get-Date;$unc="\\$dcName\$sh";try{$ok=Invoke-Retry {Test-Path -LiteralPath $unc -PathType Container};Add-Row $ForestName $DomainName $dcName SYSVOL $sh $(if($ok){'OK'}else{'Critical'}) ([int]$ok) $unc Available "Share available=$ok" (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName SYSVOL $sh Critical 0 $unc Available $_.Exception.Message (Ms $s)}}
    if(-not $SkipDfsrBacklog -and $AllDcs.Count -gt 1){$src=@($AllDcs|Where-Object{$_ -ine $dcName}|Select-Object -First 1)[0];$s=Get-Date;try{$n=Get-DfsrBacklog $src $dcName;if($null -eq $n){Add-Row $ForestName $DomainName $dcName SYSVOL DFSRBacklog OK '' NotMeasured "Warning>$DfsrBacklogWarningCount; Critical>$DfsrBacklogCriticalCount" "Not measurable from $src to $dcName" (Ms $s)}else{$st=if($n -gt $DfsrBacklogCriticalCount){'Critical'}elseif($n -gt $DfsrBacklogWarningCount){'Warning'}else{'OK'};Add-Row $ForestName $DomainName $dcName SYSVOL DFSRBacklog $st $n "$src->$dcName" "Warning>$DfsrBacklogWarningCount; Critical>$DfsrBacklogCriticalCount" 'DFSR backlog measured' (Ms $s)}}catch{Add-Row $ForestName $DomainName $dcName SYSVOL DFSRBacklog OK '' NotMeasured "Warning>$DfsrBacklogWarningCount; Critical>$DfsrBacklogCriticalCount" $_.Exception.Message (Ms $s)}}
    $s=Get-Date;try{$r=@(Resolve-DnsName -Name $dcName -ErrorAction Stop);$targets=@($r|Select-Object -First 3|ForEach-Object{$v=Get-ObjectPropertyValue $_ @('IPAddress','NameHost','NameTarget','Target','Name');if($v){$v}});Add-Row $ForestName $DomainName $dcName DNS ResolveDC $(if($r.Count -gt 0){'OK'}else{'Critical'}) $r.Count $dcName '>= 1 record' ($targets -join '; ') (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName DNS ResolveDC Critical 0 $dcName '>= 1 record' $_.Exception.Message (Ms $s)}
$s=Get-Date;$time=Get-TimeOffsetMinute $dcName;if($null -eq $time.OffsetMinutes){Add-Row $ForestName $DomainName $dcName Time W32TimeOffsetMinutes OK '' NotMeasured "<= $TimeOffsetWarningMinutes minutes" $time.Error (Ms $s)}else{Add-Row $ForestName $DomainName $dcName Time W32TimeOffsetMinutes $(if($time.OffsetMinutes -gt $TimeOffsetWarningMinutes){'Warning'}else{'OK'}) $time.OffsetMinutes 'Absolute max sample offset' "<= $TimeOffsetWarningMinutes minutes" 'w32tm /stripchart samples=3' (Ms $s)}
    $s=Get-Date
    if ($EnableRemoteDcAdminChecks) {
        try{$vol=Invoke-Retry {Get-ADDbVolume $dcName};$disk=Invoke-Retry {Get-Disk $dcName $vol};$st=if($disk.FreePct -lt $DiskFreePercentCritical -or $disk.FreeGb -lt $DiskFreeGbCritical){'Critical'}else{'OK'};Add-Row $ForestName $DomainName $dcName Disk ADDatabaseVolumeFreePercent $st $disk.FreePct "$vol free $($disk.FreeGb) GB" ">= $DiskFreePercentCritical percent and >= $DiskFreeGbCritical GB" 'AD DB volume via ntdsutil, fallback system drive' (Ms $s);Add-Row $ForestName $DomainName $dcName Disk ADDatabaseVolumeFreeGB $st $disk.FreeGb "$vol free $($disk.FreePct) percent" ">= $DiskFreeGbCritical GB and >= $DiskFreePercentCritical percent" 'Same volume as percent check' (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName Disk ADDatabaseVolumeFreePercent Critical '' Error ">= $DiskFreePercentCritical percent and >= $DiskFreeGbCritical GB" $_.Exception.Message (Ms $s)}
    }
    else {
        Add-Row $ForestName $DomainName $dcName Disk ADDatabaseVolumeFreePercent OK '' NotMeasured ">= $DiskFreePercentCritical percent and >= $DiskFreeGbCritical GB" 'Skipped by default to avoid WinRM/CIM remote admin logon to domain controllers. Use -EnableRemoteDcAdminChecks with a T0 account.' (Ms $s)
        Add-Row $ForestName $DomainName $dcName Disk ADDatabaseVolumeFreeGB OK '' NotMeasured ">= $DiskFreeGbCritical GB and >= $DiskFreePercentCritical percent" 'Skipped by default to avoid WinRM/CIM remote admin logon to domain controllers. Use -EnableRemoteDcAdminChecks with a T0 account.' (Ms $s)
    }
}
function Invoke-DomainCheck($ForestName,[string]$DomainName,$ForestInfo){
    $domain=Invoke-Retry {Get-ADDomain -Identity $DomainName -Server $DomainName -ErrorAction Stop}
    $dcs=@(Invoke-Retry {Get-ADDomainController -Filter * -Server $DomainName -ErrorAction Stop}|Sort-Object HostName)
    $dcNames=@($dcs|ForEach-Object{if($_.HostName){[string]$_.HostName}else{[string]$_.Name}})
    [void]$DomainFacts.Add([pscustomobject]@{Domain=$DomainName;DomainMode=[string]$domain.DomainMode;DCCount=@($dcs).Count;GlobalCatalogCount=@($dcs|Where-Object IsGlobalCatalog).Count;PDCEmulator=[string]$domain.PDCEmulator;RIDMaster=[string]$domain.RIDMaster;InfrastructureMaster=[string]$domain.InfrastructureMaster})
    $s=Get-Date;try{$srv=@(Resolve-DnsName -Name "_ldap._tcp.$DomainName" -Type SRV -ErrorAction Stop);$targets=@($srv|Select-Object -First 5|ForEach-Object{$target=Get-ObjectPropertyValue $_ @('NameTarget','NameHost','Target','Name');$port=Get-ObjectPropertyValue $_ @('Port');if($port){"$target`:$port"}else{$target}});Add-Row $ForestName $DomainName '' DNS LDAP_SRV $(if($srv.Count -gt 0){'OK'}else{'Critical'}) $srv.Count "_ldap._tcp.$DomainName" '>= 1 SRV record' ($targets -join '; ') (Ms $s)}catch{Add-Row $ForestName $DomainName '' DNS LDAP_SRV Critical 0 "_ldap._tcp.$DomainName" '>= 1 SRV record' $_.Exception.Message (Ms $s)}
    foreach($dc in $dcs){Invoke-DcCheck $ForestName $DomainName $dc $dcNames}
    foreach($role in 'PDCEmulator','RIDMaster','InfrastructureMaster'){$s=Get-Date;try{$h=[string]$domain.$role;$ok=Test-Port $h 389;Add-Row $ForestName $DomainName $h FSMO $role $(if($ok){'OK'}else{'Critical'}) ([int]$ok) $h 'LDAP 389 reachable' "Domain FSMO holder for $role" (Ms $s)}catch{Add-Row $ForestName $DomainName '' FSMO $role Critical 0 Error 'LDAP 389 reachable' $_.Exception.Message (Ms $s)}}
    $s=Get-Date;try{$trusts=@(Get-ADTrust -Filter * -Server $DomainName -ErrorAction Stop);$trustTestCommand=Get-Command Test-ADTrustRelationship -ErrorAction SilentlyContinue;foreach($t in $trusts){$ts=Get-Date;if($trustTestCommand){try{$ok=Test-ADTrustRelationship -Identity $t.Name -Server $DomainName -ErrorAction Stop;Add-Row $ForestName $DomainName '' Trusts $t.Name $(if($ok){'OK'}else{'Critical'}) ([int]$ok) $t.TrustType 'Trust validation succeeds' "Direction=$($t.Direction); Transitive=$($t.IsTransitive)" (Ms $ts)}catch{Add-Row $ForestName $DomainName '' Trusts $t.Name Warning 0 $t.TrustType 'Trust validation attempted' $_.Exception.Message (Ms $ts)}}else{Add-Row $ForestName $DomainName '' Trusts $t.Name OK '' $t.TrustType 'Trust validation unavailable' "Direction=$($t.Direction); Transitive=$($t.IsTransitive); Test-ADTrustRelationship not available in this PowerShell session" (Ms $ts)}};if($trusts.Count -eq 0){Add-Row $ForestName $DomainName '' Trusts TrustCount OK 0 'No trusts discovered' 'Inventory only' 'No trust returned by Get-ADTrust' (Ms $s)}}catch{Add-Row $ForestName $DomainName '' Trusts TrustEnumeration Critical '' Error 'Enumeration succeeds' $_.Exception.Message (Ms $s)}
    $s=Get-Date;try{$locked=@(Search-ADAccount -LockedOut -UsersOnly -Server $DomainName -ErrorAction Stop);Add-Row $ForestName $DomainName '' DomainStats LockedUserAccounts $(if($locked.Count -gt 0){'Warning'}else{'OK'}) $locked.Count 'Locked user accounts' '0 preferred' 'Search-ADAccount -LockedOut -UsersOnly' (Ms $s)}catch{Add-Row $ForestName $DomainName '' DomainStats LockedUserAccounts Warning '' Error '0 preferred' $_.Exception.Message (Ms $s)}
    $s=Get-Date;try{
        $domainAdminsSid = $domain.DomainSID.Value + '-512'
        $da=@(Get-ADGroupMembersBySid -Sid $domainAdminsSid -Server $DomainName);Add-Row $ForestName $DomainName '' DomainStats DomainAdminsMemberCount $(if($da.Count -gt 0){'OK'}else{'Warning'}) $da.Count 'Domain Admins recursive members' 'Inventory only' 'Group resolved by RID 512' (Ms $s)
        $dns=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach($m in $da){if($m.objectClass -eq 'user'){[void]$dns.Add($m.DistinguishedName)}}
        $builtinAdmins=@(Get-ADGroupMembersBySid -Sid 'S-1-5-32-544' -Server $DomainName)
        foreach($m in $builtinAdmins){if($m.objectClass -eq 'user'){[void]$dns.Add($m.DistinguishedName)}}
        if($DomainName -ieq $ForestInfo.RootDomain){foreach($rid in '518','519'){foreach($m in @(Get-ADGroupMembersBySid -Sid ($domain.DomainSID.Value+'-'+$rid) -Server $DomainName)){if($m.objectClass -eq 'user'){[void]$dns.Add($m.DistinguishedName)}}}}
        $never=0;foreach($dn in $dns){try{$u=Get-ADUser -Identity $dn -Server $DomainName -Properties PasswordNeverExpires -ErrorAction Stop;if($u.PasswordNeverExpires){$never++}}catch{ $null = $_ }}
        Add-Row $ForestName $DomainName '' DomainStats PrivilegedPasswordNeverExpires $(if($never -gt 0){'Warning'}else{'OK'}) $never "PrivilegedUsers=$($dns.Count)" '0 preferred' 'Domain Admins, Builtin Administrators when resolvable, Enterprise Admins, Schema Admins when applicable' (Ms $s)
    }catch{Add-Row $ForestName $DomainName '' DomainStats PrivilegedAccountStats Warning '' Error 'Inventory succeeds' $_.Exception.Message (Ms $s)}
}
try{
    Import-Module ActiveDirectory -ErrorAction Stop
    $InitializeOutputPath = InitializeScriptEnvironment -OutputPathInit $OutputFolder -LogFileName $ScriptBaseName
    $OutputFolder = $InitializeOutputPath
    $transcriptPath = $null
    $transcriptVariable = Get-Variable -Name logTranscriptFile -Scope Global -ErrorAction SilentlyContinue
    if ($transcriptVariable -and $transcriptVariable.Value) { $transcriptPath = [string]$transcriptVariable.Value }
    else {
        $transcriptVariable = Get-Variable -Name LogTranscriptFile -Scope Global -ErrorAction SilentlyContinue
        if ($transcriptVariable -and $transcriptVariable.Value) { $transcriptPath = [string]$transcriptVariable.Value }
    }
    if ($transcriptPath) { Start-Transcript -Path $transcriptPath -Append | Out-Null }
    Write-SmartM365LoadedModuleVersions
    WriteLog -Message ("Script environment initialized at {0}" -f $OutputFolder)
    WriteLog -Message ("Starting {0}" -f $TaskName)
    Invoke-SmartM365Preflight -ScriptName $TaskName -OutputPaths @($OutputFolder) -RequiredModules @('ActiveDirectory') -RequireActiveDirectoryRead | Out-Null
    if (-not $EnableRemoteDcAdminChecks) { WriteLog -Message 'Remote DC admin checks are disabled. Service status and AD database disk checks will be marked NotMeasured; use -EnableRemoteDcAdminChecks only with a T0 account.' -Level 'INFO' }
    $forest=Invoke-Retry {Get-ADForest -ErrorAction Stop}; $forestName=[string]$forest.Name
    WriteLog -Message ("Forest discovered: {0}; Domains: {1}" -f $forestName, @($forest.Domains).Count) -Level 'INFO'
    WriteLog -Message 'Collecting forest FSMO checks.' -Level 'INFO'
    foreach($role in 'SchemaMaster','DomainNamingMaster'){$s=Get-Date;try{$h=[string]$forest.$role;$ok=Test-Port $h 389;Add-Row $forestName $forest.RootDomain $h FSMO $role $(if($ok){'OK'}else{'Critical'}) ([int]$ok) $h 'LDAP 389 reachable' "Forest FSMO holder for $role" (Ms $s)}catch{Add-Row $forestName $forest.RootDomain '' FSMO $role Critical 0 Error 'LDAP 389 reachable' $_.Exception.Message (Ms $s)}}
    WriteLog -Message 'Collecting tombstone lifetime.' -Level 'INFO'
    $s=Get-Date;try{$cfg=(Get-ADRootDSE -Server $forest.RootDomain -ErrorAction Stop).configurationNamingContext;$ds="CN=Directory Service,CN=Windows NT,CN=Services,$cfg";$obj=Get-ADObject -Identity $ds -Properties tombstoneLifetime -Server $forest.RootDomain -ErrorAction Stop;$tomb=if($obj.tombstoneLifetime){[int]$obj.tombstoneLifetime}else{180};Add-Row $forestName $forest.RootDomain '' Tombstone TombstoneLifetimeDays OK $tomb 'Forest tombstone lifetime' 'Compare with oldest replication failure' $ds (Ms $s)}catch{$tomb=180;Add-Row $forestName $forest.RootDomain '' Tombstone TombstoneLifetimeDays Warning $tomb Defaulted 'Compare with oldest replication failure' $_.Exception.Message (Ms $s)}
    foreach($d in @($forest.Domains|Sort-Object)){WriteLog -Message ("Collecting domain checks: {0}" -f $d) -Level 'INFO';try{Invoke-DomainCheck $forestName ([string]$d) $forest}catch{Add-Row $forestName ([string]$d) '' Domain DomainScan Critical 0 Failed 'Domain scan succeeds' $_.Exception.Message 0}}
    $oldest=0.0;foreach($d in @($forest.Domains|Sort-Object)){$oldestErrors=$null;foreach($f in @(Get-ADReplicationFailure -Target $d -Scope Domain -ErrorAction SilentlyContinue -ErrorVariable oldestErrors)){if($f.FirstFailureTime -and $f.FirstFailureTime -ne [datetime]::MinValue){$days=((Get-Date).ToUniversalTime()-$f.FirstFailureTime.ToUniversalTime()).TotalDays;if($days -gt $oldest){$oldest=$days}}}}
    $st=if($oldest -gt $tomb){'Critical'}elseif($oldest -gt ($tomb*.8)){'Warning'}else{'OK'};Add-Row $forestName $forest.RootDomain '' Tombstone OldestReplicationFailureAgeDays $st ([math]::Round($oldest,2)) "TombstoneLifetimeDays=$tomb" "Critical > $tomb days; Warning > 80 percent" 'Compared with forest tombstone lifetime' 0
    $stamp=(Get-Date).ToUniversalTime().ToString('yyyyMMdd_HHmmss',[Globalization.CultureInfo]::InvariantCulture)
    $all = Get-RowSnapshot
    if(-not (Test-Path -LiteralPath $LatestCsvFolderPath)){New-Item -Path $LatestCsvFolderPath -ItemType Directory -Force|Out-Null}
    $latestCsv=Join-Path $LatestCsvFolderPath 'AD_HealthCheck.csv'
    if($AppendHistory){$csv=$HistoryCsvPath}else{$csv=Join-Path $OutputFolder "AD_HealthCheck_$stamp.csv"}
    Assert-SmartM365CsvDataCompleteness -Data $all -TimestampedPath $csv -LatestPath $latestCsv
    if($AppendHistory){if(Test-Path $csv){$all|Export-Csv $csv -NoTypeInformation -Append -Encoding utf8BOM}else{$all|Export-Csv $csv -NoTypeInformation -Encoding utf8BOM}}else{$all|Export-Csv $csv -NoTypeInformation -Encoding utf8BOM}
    $all|Export-Csv $latestCsv -NoTypeInformation -Encoding utf8BOM
    Invoke-SmartM365SharePointCsvUpload -LocalFilePath $csv
    Invoke-SmartM365SharePointCsvUpload -LocalFilePath $latestCsv
    $end=Get-Date;$worst=Worst $all;$subject="[$($worst.ToUpperInvariant())] Active Directory Health Check - $forestName - $RunDateUtc";$html=ConvertTo-ReportHtml -r $all -status $worst -started $RunStarted -ended $end -csv $csv -forestInfo $forest -RemoteDcAdminChecksEnabled:$EnableRemoteDcAdminChecks;if($AlwaysSend -or $worst -ne 'OK'){Send-ReportMail $subject $html}
    $summaryStatus=if($worst -eq 'OK'){'Success'}else{'CompletedWithWarnings'}
    try { Stop-Transcript | Out-Null } catch { $null = $_ }
    $global:csvGeneratedPaths = @($csv, $latestCsv)
    Complete-SmartM365ExecutionContext -Status $summaryStatus
    Write-Host "AD Health Check completed. Status=$worst; Rows=$($all.Count); Csv=$csv; Latest=$latestCsv"
    if($worst -eq 'Critical'){exit 2};if($worst -eq 'Warning'){exit 1};exit 0
}catch{
    $runError = $_
    $runErrorDetail = @(
        "Exception: $($runError.Exception.GetType().FullName): $($runError.Exception.Message)"
        if ($runError.InvocationInfo -and $runError.InvocationInfo.PositionMessage) { "Position: $($runError.InvocationInfo.PositionMessage)" }
        if ($runError.ScriptStackTrace) { "Stack: $($runError.ScriptStackTrace)" }
    ) -join ' | '
    try { WriteLog -Message $runErrorDetail -Level 'ERROR' } catch { Write-Warning $runErrorDetail }
    Add-Row '' '' '' Script UnhandledError Critical 0 Failed 'Script completes' $runErrorDetail 0
    if(-not (Test-Path -LiteralPath $OutputFolder)){New-Item -Path $OutputFolder -ItemType Directory -Force|Out-Null}
    $csv=Join-Path $OutputFolder ("AD_HealthCheck_FAILED_{0}.csv" -f (Get-Date).ToUniversalTime().ToString('yyyyMMdd_HHmmss',[Globalization.CultureInfo]::InvariantCulture));$failedRows=Get-RowSnapshot;$failedRows|Export-Csv $csv -NoTypeInformation -Encoding utf8BOM
    try{$html=ConvertTo-ReportHtml -r $failedRows -status 'Critical' -started $RunStarted -ended (Get-Date) -csv $csv;Send-ReportMail "[CRITICAL] Active Directory Health Check failed - $RunDateUtc" $html}catch{Write-Warning $_.Exception.Message}
    try { Stop-Transcript | Out-Null } catch { $null = $_ }
    $global:csvGeneratedPaths = @($csv)
    try { Complete-SmartM365ExecutionContext -Status Failed -FailureStage 'ADHealthCheck' } catch { $null = $_ }
    Write-Error $runErrorDetail;exit 2
}
# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB+jzIoUBUne5Lv
# rClXeEFi+eLUB2J/QyKkGMo2ss6o3qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIPDujPnPXPyK/00JzW9LX/UikoRLnBjEsiEq5AikJ18vMA0GCSqG
# SIb3DQEBAQUABIIBgAqINUcj8luVbl0N49HoC7Dre5TTmNWDAuBbYyGcBRVya0ED
# awoQjDJKlkl3yg0K+WXo0s8pQ3vXJX1vKfte30MltuXAx8/IeTPn6XlleFEA6Krw
# aCeGUB6mQ8UErIlP2mSwl1PJUGqooTkijl8AVzu7TR1UY83c4IaX9jePBTaOuNsZ
# tiFAoXHZiHoYNddob6uGTyAdB8xYzA8gggv22g9YyY8vI/A3PMZJ5iGZagwhy43Q
# gX533MDG0j4+6lujmsPW+kL5nX6w2TqgNtx9lafXOXnxG3GQOoZnd7mQIbCeRMya
# wm9611E2Q15DeBLsQcT8FsyX9mLafz/QG2pbveuBUz1I9flLyvwDcfFdwqjOgEef
# tzR+B3xfYg5ZGPxqmHncMSUnZPQZPqoJxDIHSU86ONu2jHp5t7dZ4AjqZ07l7gcC
# C45yzDJxmHopBHnvjz8Fu2VRCHTtqDSQWP/FSC46wdHb7Dad5RyEkAppXkCROWNx
# 1ZJBHSq9cwKw4AVt26GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMwODUw
# MjJaMC8GCSqGSIb3DQEJBDEiBCAxtxvOzwfgHFrdk2Ut8DW1cJubKc3FakJHwqPt
# Wa7SCzANBgkqhkiG9w0BAQEFAASCAgCe0cf0eKaZVDD6sKiD3Pnp4kXX2O7n2SIS
# SN24Cq0vXCf+c+/1/g8PMchTy3LpFU2TZiwqqX1JYYJe62Ipr0ze06YDwCFABnQq
# Rgn4YY4Dw3Det+UYfn1tmb/yuQ0eO7ThI0+exHYRiPOnuB1SgG6YPTtC0pZCLuFF
# 6qzJ+Lkc8vq1QWee5oC2+lLMV++e+DJgz1USQIF8n/+PlCbe/NXo0orKw7JkKmhD
# 6CjkDWMU6FXJOifXVrQyn7ASmLYAX311Sov+oI0/j42q7N3jNeheezzMnCUnI8ME
# 0GxQlh1xybmse7W3O1PDSZVhG9nrNts0xv+w6AJzC6nuC2vHLZ0+Y+38uXcccpG2
# OALQNbf+YO7SlvF33ZatYMI18J5ekQQl+7Z6MJluDigzNrVfvKCALuoZzgoirXkA
# eMjXOHPtUwi76T3NOm/EmqWGgYMQWGrkz11Yku6rOpFUylU7ScqPm/iw8Z4A3FCa
# e4KQwSPc3orqOS+e6699xM4BVLxJQ9tN3fCHm/AoL3n209n9hI3BG2PkiAtfIrQR
# ac7Zf2+9QYg51PdroC2eHhHtSYl/7p8D9m+qHGmIu3poGfX/LUiJR+2jP6JTRp3t
# UNwjavJHphfcBl9AdTZWlk/ziMyemQnAkjBcYkYE4qOTHUoTBe+0Ej+68cdYaDnp
# JQJFJWsiCw==
# SIG # End signature block
