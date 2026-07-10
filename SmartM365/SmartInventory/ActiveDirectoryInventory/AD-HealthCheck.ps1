<#
.SYNOPSIS
    Active Directory forest health check for PowerShell 7 and RSAT ActiveDirectory.
.VERSION
    1.0.3
.DESCRIPTION
    Discovers every domain with Get-ADForest, audits domain controllers and domain health,
    exports a flat Power BI-ready CSV, and sends an HTML summary email on warnings or critical alerts.
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
    [switch]$SkipDfsrBacklog
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
[System.Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::InvariantCulture
[System.Threading.Thread]::CurrentThread.CurrentUICulture = [Globalization.CultureInfo]::InvariantCulture
$RunStarted = Get-Date
$RunDateUtc = $RunStarted.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ',[Globalization.CultureInfo]::InvariantCulture)
$RunId = [guid]::NewGuid().ToString()
$Rows = New-Object 'System.Collections.Generic.List[object]'
$ScriptBaseName = [IO.Path]::GetFileNameWithoutExtension($PSCommandPath)
$ScriptVersion = '1.0.3'
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
Import-Module $CoreModulePath -Force -ErrorAction Stop
$LocalConfigPath = Join-Path $PSScriptRoot ("$ScriptBaseName.local.json")
$LocalTemplatePath = "$LocalConfigPath.template"
if (-not (Test-Path -LiteralPath $LocalConfigPath)) {
    if (Get-Command Initialize-SmartM365LocalJsonFromTemplate -ErrorAction SilentlyContinue) {
        Initialize-SmartM365LocalJsonFromTemplate -Path $LocalConfigPath -TemplatePath $LocalTemplatePath -ConfigDescription 'script local configuration' | Out-Null
    }
    elseif (Test-Path -LiteralPath $LocalTemplatePath) {
        Copy-Item -LiteralPath $LocalTemplatePath -Destination $LocalConfigPath -ErrorAction Stop
    }
}
$ScriptConfig = if (Test-Path -LiteralPath $LocalConfigPath) { Get-Content -LiteralPath $LocalConfigPath -Raw | ConvertFrom-Json } else { [pscustomobject]@{} }
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
if ([string]::IsNullOrWhiteSpace($OutputFolder)) { $OutputFolder = [string](Get-LocalConfigValue 'ADHealthCheckCsvLogFolderPath' '{{DataAllRootPath}}\ActiveDirectory\HealthCheck') }
$LatestCsvFolderPath = [string](Get-LocalConfigValue 'LatestCsvFolderPath' $TenantContext.LatestCsvFolderPath)
if (-not $PSBoundParameters.ContainsKey('AlwaysSend')) { $AlwaysSend = [bool](Get-LocalConfigValue 'AlwaysSend' $false) }
if (-not $PSBoundParameters.ContainsKey('AppendHistory')) { $AppendHistory = [bool](Get-LocalConfigValue 'AppendHistory' $false) }
if ([string]::IsNullOrWhiteSpace($HistoryCsvPath)) { $HistoryCsvPath = [string](Get-LocalConfigValue 'HistoryCsvPath' '{{DataAllRootPath}}\ActiveDirectory\HealthCheck\AD_HealthCheck_History.csv') }
function ConvertTo-IsoUtc([datetime]$d){ if($null -eq $d -or $d -eq [datetime]::MinValue){''}else{$d.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ',[Globalization.CultureInfo]::InvariantCulture)} }
function Num($v){ if($null -eq $v -or [string]::IsNullOrWhiteSpace([string]$v)){''}else{ try{([double]$v).ToString('0.########',[Globalization.CultureInfo]::InvariantCulture)}catch{[string]$v} } }
function Ms($s){ [int64][math]::Max(0,((Get-Date)-$s).TotalMilliseconds) }
function Add-Row{ param($Forest,$Domain,$DC,[string]$Category,[string]$Check,[ValidateSet('OK','Warning','Critical')]$Status,$NumericValue,$TextValue,$Threshold,$Details,[int64]$DurationMs)
    [void]$Rows.Add([pscustomobject][ordered]@{RunId=$RunId;RunDateUtc=$RunDateUtc;Forest=[string]$Forest;Domain=[string]$Domain;DC=[string]$DC;Category=$Category;Check=$Check;Status=$Status;NumericValue=(Num $NumericValue);TextValue=[string]$TextValue;Threshold=[string]$Threshold;Details=[string]$Details;DurationMs=$DurationMs})
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
function Get-TimeOffsetMinute([string]$DC){
    $out=& w32tm.exe /stripchart /computer:$DC /samples:3 /dataonly 2>&1
    $vals=@(); foreach($line in $out){ foreach($m in [regex]::Matches([string]$line,'([+-]?\d+(?:[\.,]\d+)?)s')){ $raw=$m.Groups[1].Value.Replace(',','.'); $d=0.0; if([double]::TryParse($raw,[Globalization.NumberStyles]::Float,[Globalization.CultureInfo]::InvariantCulture,[ref]$d)){$vals += [math]::Abs($d)} } }
    if($vals.Count -eq 0){ throw "Unable to parse w32tm output: $($out -join ' | ')" }
    (($vals|Measure-Object -Maximum).Maximum/60)
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
function ConvertTo-ReportHtml($r,[string]$status,[datetime]$started,[datetime]$ended,[string]$csv){
    $c=@{OK='#107c10';Warning='#ff8c00';Critical='#d13438'}[$status]; $find=@($r|Where-Object Status -ne OK|Sort-Object @{Expression={Rank $_.Status}},Domain,DC,Category,Check|Select-Object -First 200); $b=[Text.StringBuilder]::new()
    [void]$b.AppendLine('<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:Segoe UI,Arial;background:#f5f8fb;color:#1f2937;padding:24px}.card{background:#fff;border:1px solid #dde7f0;border-radius:8px;padding:16px;margin:0 0 16px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #dde7f0;padding:7px;font-size:12px;text-align:left;vertical-align:top}th{background:#eef6fc}.rowOK{background:#f3fbf3}.rowWarning{background:#fff7e6}.rowCritical{background:#fde7e9}.pill{color:#fff;border-radius:999px;padding:4px 10px;font-weight:600}</style></head><body>')
    [void]$b.AppendLine(("<div class='card'><h1>Active Directory Health Check <span class='pill' style='background:{0}'>{1}</span></h1><p>RunId: {2}<br>Machine: {3}<br>Started UTC: {4}<br>Ended UTC: {5}<br>Duration: {6}<br>CSV: {7}</p></div>" -f $c,$status,(ConvertTo-HtmlSafe $RunId),(ConvertTo-HtmlSafe $env:COMPUTERNAME),(ConvertTo-IsoUtc $started),(ConvertTo-IsoUtc $ended),(ConvertTo-HtmlSafe ((New-TimeSpan -Start $started -End $ended).ToString())),(ConvertTo-HtmlSafe $csv)))
    [void]$b.AppendLine(("<div class='card'><b>Summary</b>: Critical={0}; Warning={1}; OK={2}; Total={3}</div>" -f @($r|Where-Object Status -eq Critical).Count,@($r|Where-Object Status -eq Warning).Count,@($r|Where-Object Status -eq OK).Count,$r.Count))
    [void]$b.AppendLine('<div class="card"><h2>Critical and warning findings</h2><table><tr><th>Status</th><th>Domain</th><th>DC</th><th>Category</th><th>Check</th><th>Value</th><th>Threshold</th><th>Details</th></tr>')
    if($find.Count -eq 0){[void]$b.AppendLine('<tr class="rowOK"><td colspan="8">No critical or warning findings.</td></tr>')}
    foreach($x in $find){[void]$b.AppendLine(("<tr class='row{0}'><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td><td>{4}</td><td>{5}</td><td>{6}</td><td>{7}</td></tr>" -f (ConvertTo-HtmlSafe $x.Status),(ConvertTo-HtmlSafe $x.Domain),(ConvertTo-HtmlSafe $x.DC),(ConvertTo-HtmlSafe $x.Category),(ConvertTo-HtmlSafe $x.Check),(ConvertTo-HtmlSafe $x.NumericValue),(ConvertTo-HtmlSafe $x.Threshold),(ConvertTo-HtmlSafe $x.Details)))}
    [void]$b.AppendLine('</table></div><div class="card"><h2>OK summary by domain</h2><table><tr><th>Domain</th><th>OK</th><th>Warning</th><th>Critical</th></tr>')
    foreach($g in ($r|Where-Object Domain|Group-Object Domain|Sort-Object Name)){ $it=@($g.Group); [void]$b.AppendLine(("<tr><td>{0}</td><td>{1}</td><td>{2}</td><td>{3}</td></tr>" -f (ConvertTo-HtmlSafe $g.Name),@($it|Where-Object Status -eq OK).Count,@($it|Where-Object Status -eq Warning).Count,@($it|Where-Object Status -eq Critical).Count)) }
    [void]$b.AppendLine('</table></div></body></html>'); $b.ToString()
}
function Invoke-DcCheck($ForestName,$DomainName,$DC,$AllDcs){
    $dcName=if($DC.HostName){[string]$DC.HostName}else{[string]$DC.Name}
    foreach($port in 389,3268,445){$s=Get-Date;try{$ok=Invoke-Retry {Test-Port $dcName $port};$stat=if($ok -or ($port -eq 3268 -and -not $DC.IsGlobalCatalog)){'OK'}else{'Critical'};Add-Row $ForestName $DomainName $dcName Connectivity "TCP $port" $stat ([int]$ok) "Reachable=$ok" 'Open' "TCP port $port" (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName Connectivity "TCP $port" Critical 0 Error Open $_.Exception.Message (Ms $s)}}
    $s=Get-Date;try{$ok=Invoke-Retry {Test-Connection -ComputerName $dcName -Count 1 -Quiet -ErrorAction Stop};Add-Row $ForestName $DomainName $dcName Connectivity Ping $(if($ok){'OK'}else{'Critical'}) ([int]$ok) "Reachable=$ok" Reachable 'ICMP echo test' (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName Connectivity Ping Critical 0 Error Reachable $_.Exception.Message (Ms $s)}
    foreach($svcName in 'NTDS','DNS','Netlogon','KDC','W32Time','DFSR'){$s=Get-Date;try{$svc=Invoke-Retry {Get-Service -ComputerName $dcName -Name $svcName -ErrorAction Stop};$ok=$svc.Status -eq 'Running';Add-Row $ForestName $DomainName $dcName Services $svcName $(if($ok){'OK'}else{'Critical'}) ([int]$ok) $svc.Status Running "Service $svcName" (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName Services $svcName Critical 0 Error Running $_.Exception.Message (Ms $s)}}
    $s=Get-Date;try{$f=@(Invoke-Retry {Get-ADReplicationFailure -Target $dcName -Scope Server -ErrorAction Stop});$det=($f|Select-Object -First 5|ForEach-Object{"$($_.Partner): $($_.FailureCount) since $(ConvertTo-IsoUtc $_.FirstFailureTime)"}) -join '; ';Add-Row $ForestName $DomainName $dcName Replication ReplicationFailures $(if($f.Count -eq 0){'OK'}else{'Critical'}) $f.Count "Failures=$($f.Count)" 0 $det (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName Replication ReplicationFailures Critical '' Error 0 $_.Exception.Message (Ms $s)}
    $s=Get-Date;try{$m=@(Invoke-Retry {Get-ADReplicationPartnerMetadata -Target $dcName -Scope Server -ErrorAction Stop});$old=@($m|Where-Object{$_.LastReplicationSuccess -and $_.LastReplicationSuccess -ne [datetime]::MinValue}|Sort-Object LastReplicationSuccess|Select-Object -First 1)[0];$max=0;if($old){$max=[math]::Round(((Get-Date).ToUniversalTime()-$old.LastReplicationSuccess.ToUniversalTime()).TotalHours,2)};$stat=if($max -gt $ReplicationDelayWarningHours){'Warning'}else{'OK'};Add-Row $ForestName $DomainName $dcName Replication MaxLastSuccessDelayHours $stat $max $(if($old){ConvertTo-IsoUtc $old.LastReplicationSuccess}else{''}) "<= $ReplicationDelayWarningHours hours" "Partners=$($m.Count)" (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName Replication MaxLastSuccessDelayHours Critical '' Error "<= $ReplicationDelayWarningHours hours" $_.Exception.Message (Ms $s)}
    foreach($sh in 'SYSVOL','NETLOGON'){$s=Get-Date;$unc="\\$dcName\$sh";try{$ok=Invoke-Retry {Test-Path -LiteralPath $unc -PathType Container};Add-Row $ForestName $DomainName $dcName SYSVOL $sh $(if($ok){'OK'}else{'Critical'}) ([int]$ok) $unc Available "Share available=$ok" (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName SYSVOL $sh Critical 0 $unc Available $_.Exception.Message (Ms $s)}}
    if(-not $SkipDfsrBacklog -and $AllDcs.Count -gt 1){$src=@($AllDcs|Where-Object{$_ -ine $dcName}|Select-Object -First 1)[0];$s=Get-Date;try{$n=Get-DfsrBacklog $src $dcName;if($null -eq $n){Add-Row $ForestName $DomainName $dcName SYSVOL DFSRBacklog OK '' NotMeasured "Warning>$DfsrBacklogWarningCount; Critical>$DfsrBacklogCriticalCount" "Not measurable from $src to $dcName" (Ms $s)}else{$st=if($n -gt $DfsrBacklogCriticalCount){'Critical'}elseif($n -gt $DfsrBacklogWarningCount){'Warning'}else{'OK'};Add-Row $ForestName $DomainName $dcName SYSVOL DFSRBacklog $st $n "$src->$dcName" "Warning>$DfsrBacklogWarningCount; Critical>$DfsrBacklogCriticalCount" 'DFSR backlog measured' (Ms $s)}}catch{Add-Row $ForestName $DomainName $dcName SYSVOL DFSRBacklog OK '' NotMeasured "Warning>$DfsrBacklogWarningCount; Critical>$DfsrBacklogCriticalCount" $_.Exception.Message (Ms $s)}}
    $s=Get-Date;try{$r=Resolve-DnsName -Name $dcName -ErrorAction Stop;Add-Row $ForestName $DomainName $dcName DNS ResolveDC OK $r.Count $dcName '>= 1 record' (($r|Select-Object -First 3|ForEach-Object{$_.IPAddress}) -join '; ') (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName DNS ResolveDC Critical 0 $dcName '>= 1 record' $_.Exception.Message (Ms $s)}
    $s=Get-Date;try{$off=Invoke-Retry {Get-TimeOffsetMinute $dcName};Add-Row $ForestName $DomainName $dcName Time W32TimeOffsetMinutes $(if($off -gt $TimeOffsetWarningMinutes){'Warning'}else{'OK'}) $off 'Absolute max sample offset' "<= $TimeOffsetWarningMinutes minutes" 'w32tm /stripchart samples=3' (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName Time W32TimeOffsetMinutes Warning '' NotMeasured "<= $TimeOffsetWarningMinutes minutes" $_.Exception.Message (Ms $s)}
    $s=Get-Date;try{$vol=Invoke-Retry {Get-ADDbVolume $dcName};$disk=Invoke-Retry {Get-Disk $dcName $vol};$st=if($disk.FreePct -lt $DiskFreePercentCritical -or $disk.FreeGb -lt $DiskFreeGbCritical){'Critical'}else{'OK'};Add-Row $ForestName $DomainName $dcName Disk ADDatabaseVolumeFreePercent $st $disk.FreePct "$vol free $($disk.FreeGb) GB" ">= $DiskFreePercentCritical percent and >= $DiskFreeGbCritical GB" 'AD DB volume via ntdsutil, fallback system drive' (Ms $s);Add-Row $ForestName $DomainName $dcName Disk ADDatabaseVolumeFreeGB $st $disk.FreeGb "$vol free $($disk.FreePct) percent" ">= $DiskFreeGbCritical GB and >= $DiskFreePercentCritical percent" 'Same volume as percent check' (Ms $s)}catch{Add-Row $ForestName $DomainName $dcName Disk ADDatabaseVolumeFreePercent Critical '' Error ">= $DiskFreePercentCritical percent and >= $DiskFreeGbCritical GB" $_.Exception.Message (Ms $s)}
}
function Invoke-DomainCheck($ForestName,[string]$DomainName,$ForestInfo){
    $domain=Invoke-Retry {Get-ADDomain -Identity $DomainName -Server $DomainName -ErrorAction Stop}
    $dcs=@(Invoke-Retry {Get-ADDomainController -Filter * -Server $DomainName -ErrorAction Stop}|Sort-Object HostName)
    $dcNames=@($dcs|ForEach-Object{if($_.HostName){[string]$_.HostName}else{[string]$_.Name}})
    $s=Get-Date;try{$srv=@(Resolve-DnsName -Name "_ldap._tcp.$DomainName" -Type SRV -ErrorAction Stop);Add-Row $ForestName $DomainName '' DNS LDAP_SRV $(if($srv.Count -gt 0){'OK'}else{'Critical'}) $srv.Count "_ldap._tcp.$DomainName" '>= 1 SRV record' (($srv|Select-Object -First 5|ForEach-Object{"$($_.NameTarget):$($_.Port)"}) -join '; ') (Ms $s)}catch{Add-Row $ForestName $DomainName '' DNS LDAP_SRV Critical 0 "_ldap._tcp.$DomainName" '>= 1 SRV record' $_.Exception.Message (Ms $s)}
    foreach($dc in $dcs){Invoke-DcCheck $ForestName $DomainName $dc $dcNames}
    foreach($role in 'PDCEmulator','RIDMaster','InfrastructureMaster'){$s=Get-Date;try{$h=[string]$domain.$role;$ok=Test-Port $h 389;Add-Row $ForestName $DomainName $h FSMO $role $(if($ok){'OK'}else{'Critical'}) ([int]$ok) $h 'LDAP 389 reachable' "Domain FSMO holder for $role" (Ms $s)}catch{Add-Row $ForestName $DomainName '' FSMO $role Critical 0 Error 'LDAP 389 reachable' $_.Exception.Message (Ms $s)}}
    $s=Get-Date;try{$trusts=@(Get-ADTrust -Filter * -Server $DomainName -ErrorAction Stop);foreach($t in $trusts){$ts=Get-Date;try{$ok=Test-ADTrustRelationship -Identity $t.Name -Server $DomainName -ErrorAction Stop;Add-Row $ForestName $DomainName '' Trusts $t.Name $(if($ok){'OK'}else{'Critical'}) ([int]$ok) $t.TrustType 'Trust validation succeeds' "Direction=$($t.Direction); Transitive=$($t.IsTransitive)" (Ms $ts)}catch{Add-Row $ForestName $DomainName '' Trusts $t.Name Critical 0 $t.TrustType 'Trust validation succeeds' $_.Exception.Message (Ms $ts)}};if($trusts.Count -eq 0){Add-Row $ForestName $DomainName '' Trusts TrustCount OK 0 'No trusts discovered' 'Inventory only' 'No trust returned by Get-ADTrust' (Ms $s)}}catch{Add-Row $ForestName $DomainName '' Trusts TrustEnumeration Critical '' Error 'Enumeration succeeds' $_.Exception.Message (Ms $s)}
    $s=Get-Date;try{$locked=@(Search-ADAccount -LockedOut -UsersOnly -Server $DomainName -ErrorAction Stop);Add-Row $ForestName $DomainName '' DomainStats LockedUserAccounts $(if($locked.Count -gt 0){'Warning'}else{'OK'}) $locked.Count 'Locked user accounts' '0 preferred' 'Search-ADAccount -LockedOut -UsersOnly' (Ms $s)}catch{Add-Row $ForestName $DomainName '' DomainStats LockedUserAccounts Warning '' Error '0 preferred' $_.Exception.Message (Ms $s)}
    $s=Get-Date;try{
        $da=@(Get-ADGroupMember -Identity ($domain.DomainSID.Value+'-512') -Recursive -Server $DomainName -ErrorAction Stop);Add-Row $ForestName $DomainName '' DomainStats DomainAdminsMemberCount OK $da.Count 'Domain Admins recursive members' 'Inventory only' 'Group resolved by RID 512' (Ms $s)
        $dns=New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
        foreach($rid in '512','544'){try{foreach($m in @(Get-ADGroupMember -Identity ($domain.DomainSID.Value+'-'+$rid) -Recursive -Server $DomainName -ErrorAction Stop)){if($m.objectClass -eq 'user'){[void]$dns.Add($m.DistinguishedName)}}}catch{ $null = $_ }}
        if($DomainName -ieq $ForestInfo.RootDomain){foreach($rid in '518','519'){try{foreach($m in @(Get-ADGroupMember -Identity ($domain.DomainSID.Value+'-'+$rid) -Recursive -Server $DomainName -ErrorAction Stop)){if($m.objectClass -eq 'user'){[void]$dns.Add($m.DistinguishedName)}}}catch{ $null = $_ }}}
        $never=0;foreach($dn in $dns){try{$u=Get-ADUser -Identity $dn -Server $DomainName -Properties PasswordNeverExpires -ErrorAction Stop;if($u.PasswordNeverExpires){$never++}}catch{ $null = $_ }}
        Add-Row $ForestName $DomainName '' DomainStats PrivilegedPasswordNeverExpires $(if($never -gt 0){'Warning'}else{'OK'}) $never "PrivilegedUsers=$($dns.Count)" '0 preferred' 'Domain Admins, Administrators, Enterprise Admins, Schema Admins when applicable' (Ms $s)
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
    $forest=Invoke-Retry {Get-ADForest -ErrorAction Stop}; $forestName=[string]$forest.Name
    foreach($role in 'SchemaMaster','DomainNamingMaster'){$s=Get-Date;try{$h=[string]$forest.$role;$ok=Test-Port $h 389;Add-Row $forestName $forest.RootDomain $h FSMO $role $(if($ok){'OK'}else{'Critical'}) ([int]$ok) $h 'LDAP 389 reachable' "Forest FSMO holder for $role" (Ms $s)}catch{Add-Row $forestName $forest.RootDomain '' FSMO $role Critical 0 Error 'LDAP 389 reachable' $_.Exception.Message (Ms $s)}}
    $s=Get-Date;try{$cfg=(Get-ADRootDSE -Server $forest.RootDomain -ErrorAction Stop).configurationNamingContext;$ds="CN=Directory Service,CN=Windows NT,CN=Services,$cfg";$obj=Get-ADObject -Identity $ds -Properties tombstoneLifetime -Server $forest.RootDomain -ErrorAction Stop;$tomb=if($obj.tombstoneLifetime){[int]$obj.tombstoneLifetime}else{180};Add-Row $forestName $forest.RootDomain '' Tombstone TombstoneLifetimeDays OK $tomb 'Forest tombstone lifetime' 'Compare with oldest replication failure' $ds (Ms $s)}catch{$tomb=180;Add-Row $forestName $forest.RootDomain '' Tombstone TombstoneLifetimeDays Warning $tomb Defaulted 'Compare with oldest replication failure' $_.Exception.Message (Ms $s)}
    foreach($d in @($forest.Domains|Sort-Object)){try{Invoke-DomainCheck $forestName ([string]$d) $forest}catch{Add-Row $forestName ([string]$d) '' Domain DomainScan Critical 0 Failed 'Domain scan succeeds' $_.Exception.Message 0}}
    $oldest=0.0;foreach($d in @($forest.Domains|Sort-Object)){try{foreach($f in @(Get-ADReplicationFailure -Target $d -Scope Domain -ErrorAction Stop)){if($f.FirstFailureTime -and $f.FirstFailureTime -ne [datetime]::MinValue){$days=((Get-Date).ToUniversalTime()-$f.FirstFailureTime.ToUniversalTime()).TotalDays;if($days -gt $oldest){$oldest=$days}}}}catch{ $null = $_ }}
    $st=if($oldest -gt $tomb){'Critical'}elseif($oldest -gt ($tomb*.8)){'Warning'}else{'OK'};Add-Row $forestName $forest.RootDomain '' Tombstone OldestReplicationFailureAgeDays $st ([math]::Round($oldest,2)) "TombstoneLifetimeDays=$tomb" "Critical > $tomb days; Warning > 80 percent" 'Compared with forest tombstone lifetime' 0
    $stamp=(Get-Date).ToUniversalTime().ToString('yyyyMMdd_HHmmss',[Globalization.CultureInfo]::InvariantCulture)
    if($AppendHistory){$csv=$HistoryCsvPath;if(Test-Path $csv){$Rows|Export-Csv $csv -NoTypeInformation -Append -Encoding utf8BOM}else{$Rows|Export-Csv $csv -NoTypeInformation -Encoding utf8BOM}}else{$csv=Join-Path $OutputFolder "AD_HealthCheck_$stamp.csv";$Rows|Export-Csv $csv -NoTypeInformation -Encoding utf8BOM}
    if(-not (Test-Path -LiteralPath $LatestCsvFolderPath)){New-Item -Path $LatestCsvFolderPath -ItemType Directory -Force|Out-Null}
    $latestCsv=Join-Path $LatestCsvFolderPath 'AD_HealthCheck.csv'
    $Rows|Export-Csv $latestCsv -NoTypeInformation -Encoding utf8BOM
    Invoke-SmartM365SharePointCsvUpload -LocalFilePath $csv
    Invoke-SmartM365SharePointCsvUpload -LocalFilePath $latestCsv
    $end=Get-Date;$all=@($Rows);$worst=Worst $all;$subject="[$($worst.ToUpperInvariant())] Active Directory Health Check - $forestName - $RunDateUtc";$html=ConvertTo-ReportHtml $all $worst $RunStarted $end $csv;if($AlwaysSend -or $worst -ne 'OK'){Send-ReportMail $subject $html}
    $summaryStatus=if($worst -eq 'OK'){'Success'}else{'CompletedWithWarnings'}
    Complete-SmartM365ExecutionContext -Status $summaryStatus -GeneratedCsvPaths @($csv,$latestCsv) -ResultSummary "AD health worst status: $worst; rows: $($all.Count)"
    try { Stop-Transcript | Out-Null } catch { $null = $_ }
    Write-Host "AD Health Check completed. Status=$worst; Rows=$($all.Count); Csv=$csv; Latest=$latestCsv"
    if($worst -eq 'Critical'){exit 2};if($worst -eq 'Warning'){exit 1};exit 0
}catch{
    Add-Row '' '' '' Script UnhandledError Critical 0 Failed 'Script completes' $_.Exception.Message 0
    if(-not (Test-Path -LiteralPath $OutputFolder)){New-Item -Path $OutputFolder -ItemType Directory -Force|Out-Null}
    $csv=Join-Path $OutputFolder ("AD_HealthCheck_FAILED_{0}.csv" -f (Get-Date).ToUniversalTime().ToString('yyyyMMdd_HHmmss',[Globalization.CultureInfo]::InvariantCulture));$Rows|Export-Csv $csv -NoTypeInformation -Encoding utf8BOM
    try{$html=ConvertTo-ReportHtml @($Rows) Critical $RunStarted (Get-Date) $csv;Send-ReportMail "[CRITICAL] Active Directory Health Check failed - $RunDateUtc" $html}catch{Write-Warning $_.Exception.Message}
    try { Complete-SmartM365ExecutionContext -Status Failed -GeneratedCsvPaths @($csv) -ResultSummary 'AD health check failed before completion.' } catch { $null = $_ }
    try { Stop-Transcript | Out-Null } catch { $null = $_ }
    Write-Error $_;exit 2
}
