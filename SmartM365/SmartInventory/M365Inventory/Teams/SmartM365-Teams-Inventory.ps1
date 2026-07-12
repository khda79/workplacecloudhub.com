#Requires -Version 7.0
<#
.SYNOPSIS
    Microsoft Teams tenant inventory with CSV exports and HTML alert summary.
.VERSION
0.13

.REQUIREMENTS
    PowerShell 7+.
    Modules: SmartM365.Core; Microsoft.Graph.Authentication.
    Minimum Graph application permissions: Team.ReadBasic.All; TeamMember.Read.All; Channel.ReadBasic.All; Group.Read.All; Reports.Read.All; Sites.Read.All.
    Optional: ChannelMember.Read.All is required only when private/shared channel member or owner expansion is enabled.
    Conditional: Mail.Send is required only when Graph mail is used; Sites.Selected write is required only when SharePoint upload is enabled.
.NOTES
    Requires: PowerShell 7+, Microsoft.Graph.Authentication, SmartM365.Core.psd1
    Minimum application permissions: Team.ReadBasic.All, TeamMember.Read.All, Channel.ReadBasic.All, Group.Read.All, Reports.Read.All, Sites.Read.All.
    Optional: ChannelMember.Read.All for private/shared channel owners when -IncludeChannelOwners is used.
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidGlobalVars','',Justification='SmartM365.Core uses global execution context variables for logs and generated CSV tracking.')]
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWriteHost','',Justification='Final console status is intentional for command-line use.')]
[CmdletBinding()]
param(
    [string]$Tenant='test',
    [int]$InactiveDays=180,
    [int]$MaxTeams=0,
    [switch]$DryRun,
    [switch]$AlwaysSend,
    [switch]$AppendHistory,
    [double]$QuotaCriticalPercent=90,
    [int]$MinOwners=2,
    [int]$GuestWarningThreshold=25,
    [switch]$RequireSensitivityLabel,
    [switch]$IncludeChannelOwners,
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
$ErrorActionPreference='Stop'; Set-StrictMode -Version Latest
$ScriptVersion="0.13"
$RunStarted=Get-Date; $RunDateUtc=$RunStarted.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ',[Globalization.CultureInfo]::InvariantCulture); $RunId=[guid]::NewGuid().ToString(); $CurrentOperation='Initialize'
$TeamsRows=New-Object 'System.Collections.Generic.List[object]'; $MembersRows=New-Object 'System.Collections.Generic.List[object]'; $ChannelsRows=New-Object 'System.Collections.Generic.List[object]'; $GuestsRows=New-Object 'System.Collections.Generic.List[object]'; $Alerts=New-Object 'System.Collections.Generic.List[object]'; $GeneratedCsvPaths=New-Object 'System.Collections.Generic.List[string]'
if($PSVersionTable.PSVersion.Major -lt 7){throw 'This script requires PowerShell 7 or later.'}
$tenantContextPath=&{ $d=$PSScriptRoot; while($d){ foreach($c in @((Join-Path $d 'SmartM365-TenantContext.ps1'),(Join-Path $d 'Config\SmartM365-TenantContext.ps1'))){ if(Test-Path -LiteralPath $c){return $c} }; $p=Split-Path $d -Parent; if([string]::IsNullOrWhiteSpace($p)-or$p-eq$d){break}; $d=$p }; throw 'SmartM365-TenantContext.ps1 not found.' }
. $tenantContextPath
$TenantContext=Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot
$ctxDir=Split-Path $tenantContextPath -Parent; $SmartM365Root=if((Split-Path $ctxDir -Leaf)-ieq 'Config'){Split-Path $ctxDir -Parent}else{$ctxDir}
Import-Module -Name (Join-Path $SmartM365Root 'Modules\SmartM365.Core\SmartM365.Core.psd1') -MinimumVersion '1.0.24' -Force -ErrorAction Stop
$LocalConfigPath=Join-Path $PSScriptRoot "$ScriptBaseName.local.json"; $LocalTemplatePath="$LocalConfigPath.template"
if(-not(Test-Path -LiteralPath $LocalConfigPath)){Initialize-SmartM365LocalJsonFromTemplate -Path $LocalConfigPath -TemplatePath $LocalTemplatePath -ConfigDescription 'script local configuration'|Out-Null}
$ScriptConfig=Get-Content -LiteralPath $LocalConfigPath -Raw|ConvertFrom-Json
function Resolve-ConfigToken{param([AllowNull()][object]$Value) if($Value -isnot [string]){return $Value}; $r=$Value; for($i=0;$i-lt 10;$i++){ $m=[regex]::Matches($r,'\{\{(?<Name>[A-Za-z0-9_.-]+)\}\}'); if($m.Count-eq 0){break}; foreach($x in $m){$p=$TenantContext.PSObject.Properties[$x.Groups['Name'].Value]; if($p-and$null-ne$p.Value){$r=$r.Replace($x.Value,[string]$p.Value)}}}; $r}
function Get-ConfigValue{param([string]$Name,[AllowNull()][object]$DefaultValue) $p=$ScriptConfig.PSObject.Properties[$Name]; if($p-and$null-ne$p.Value){ if($p.Value -isnot [string] -or ($p.Value.Trim() -and $p.Value.Trim() -notin @('__USE_GLOBAL__','USE_GLOBAL'))){return Resolve-ConfigToken $p.Value}}; $c=$TenantContext.PSObject.Properties[$Name]; if($c-and$null-ne$c.Value){return Resolve-ConfigToken $c.Value}; Resolve-ConfigToken $DefaultValue}
function IsoUtc{param([AllowNull()][object]$Value) if($null-eq$Value -or [string]::IsNullOrWhiteSpace([string]$Value)){return ''}; try{([datetime]$Value).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ',[Globalization.CultureInfo]::InvariantCulture)}catch{''}}
function Num{param([AllowNull()][object]$Value) if($null-eq$Value -or [string]::IsNullOrWhiteSpace([string]$Value)){return ''}; try{([double]$Value).ToString('0.########',[Globalization.CultureInfo]::InvariantCulture)}catch{[string]$Value}}
function Prop{param([AllowNull()][object]$Object,[string[]]$Names) if($null-eq$Object){return $null}; foreach($n in $Names){$p=$Object.PSObject.Properties[$n]; if($p){return $p.Value}}; $null}
function JoinVals{param([AllowNull()][object[]]$Values) @($Values|Where-Object{-not[string]::IsNullOrWhiteSpace([string]$_)}) -join '; '}

function Add-Alert{param([string]$TeamId,[string]$TeamDisplayName,[ValidateSet('Warning','Critical')][string]$Status,[string]$Check,[AllowNull()][object]$NumericValue,[string]$TextValue,[string]$Threshold,[string]$Details) [void]$Alerts.Add([pscustomobject]@{TeamId=$TeamId;TeamDisplayName=$TeamDisplayName;Status=$Status;Check=$Check;NumericValue=(Num $NumericValue);TextValue=$TextValue;Threshold=$Threshold;Details=$Details})}
function WorstStatus{param([object[]]$Rows) if(@($Rows|Where-Object Status -eq Critical).Count){'Critical'}elseif(@($Rows|Where-Object Status -eq Warning).Count){'Warning'}else{'OK'}}
function Invoke-Graph{param([string]$Uri,[string]$Operation='Graph request',[string]$OutputFilePath='') for($a=1;$a-le 5;$a++){try{$p=@{Method='GET';Uri=$Uri;ErrorAction='Stop'}; if($OutputFilePath){$p.OutputFilePath=$OutputFilePath}; return Invoke-MgGraphRequest @p}catch{$sc=$null; try{if($_.Exception.Response){$sc=[int]$_.Exception.Response.StatusCode}}catch{$null=$_}; $transient=$sc-in@(429,500,502,503,504)-or([string]$_.Exception.Message-match'throttl|TooManyRequests|temporarily|timeout'); if(-not$transient-or$a-ge 5){throw}; $delay=[Math]::Min(300,[Math]::Pow(2,$a)*5); WriteLog -Message ("$Operation transient/throttled. Status=$sc; attempt $a/5; retry in $delay s.") -Level WARNING; Start-Sleep -Seconds $delay}}}
function Get-GraphCollection{param([string]$Uri,[string]$Operation) $items=New-Object 'System.Collections.Generic.List[object]'; $next=$Uri; while($next){$r=Invoke-Graph -Uri $next -Operation $Operation; foreach($i in @($r.value)){[void]$items.Add($i)}; $p=$r.PSObject.Properties['@odata.nextLink']; $next=if($p){[string]$p.Value}else{''}}; return $items.ToArray()}
function Get-ReportRow{param([string]$ReportName,[string]$Period='D180') $tmp=Join-Path ([IO.Path]::GetTempPath()) ("SmartM365-$ReportName-$([guid]::NewGuid().ToString('N')).csv"); try{Invoke-Graph -Uri ("https://graph.microsoft.com/v1.0/reports/{0}(period='{1}')" -f $ReportName,$Period) -Operation $ReportName -OutputFilePath $tmp|Out-Null; if(Test-Path -LiteralPath $tmp){return @(Import-Csv -LiteralPath $tmp)}}catch{WriteLog -Message ("Report $ReportName could not be loaded: $($_.Exception.Message)") -Level WARNING}finally{if(Test-Path -LiteralPath $tmp){Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue}}; @()}
function Export-InventoryCsv{param([object[]]$Rows,[string[]]$Columns,[string]$TimestampedPath,[string]$LatestPath,[string]$HistoryPath) Assert-SmartM365CsvDataCompleteness -Data $Rows -Columns $Columns -TimestampedPath $TimestampedPath -LatestPath $LatestPath; foreach($folder in @((Split-Path $TimestampedPath -Parent),(Split-Path $LatestPath -Parent))){if(-not(Test-Path -LiteralPath $folder)){New-Item -Path $folder -ItemType Directory -Force|Out-Null}}; if($Rows.Count-eq 0){$h=($Columns|ForEach-Object{'"'+($_-replace'"','""')+'"'})-join','; Set-Content -LiteralPath $TimestampedPath -Value $h -Encoding utf8BOM; Set-Content -LiteralPath $LatestPath -Value $h -Encoding utf8BOM}else{$Rows|Select-Object -Property $Columns|Export-Csv -LiteralPath $TimestampedPath -NoTypeInformation -Encoding utf8BOM; $Rows|Select-Object -Property $Columns|Export-Csv -LiteralPath $LatestPath -NoTypeInformation -Encoding utf8BOM}; [void]$GeneratedCsvPaths.Add($TimestampedPath); [void]$GeneratedCsvPaths.Add($LatestPath); if(-not$global:csvGeneratedPaths){$global:csvGeneratedPaths=New-Object 'System.Collections.Generic.HashSet[string]'([StringComparer]::OrdinalIgnoreCase)}; [void]$global:csvGeneratedPaths.Add($TimestampedPath); [void]$global:csvGeneratedPaths.Add($LatestPath); if($DryRun){WriteLog -Message 'DryRun enabled: SharePoint CSV upload skipped.' -Level INFO}else{Invoke-SmartM365SharePointCsvUpload -LocalFilePath $TimestampedPath|Out-Null; Invoke-SmartM365SharePointCsvUpload -LocalFilePath $LatestPath|Out-Null}; if($AppendHistory-and$HistoryPath){$hp=Split-Path $HistoryPath -Parent; if(-not(Test-Path -LiteralPath $hp)){New-Item -Path $hp -ItemType Directory -Force|Out-Null}; if($Rows.Count-gt 0){$Rows|Select-Object -Property $Columns|Export-Csv -LiteralPath $HistoryPath -NoTypeInformation -Encoding utf8BOM -Append:(Test-Path -LiteralPath $HistoryPath)}}}
function ConvertTo-HtmlReport {
    param([object[]]$AlertRows,[hashtable]$Summary,[string]$Worst,[datetime]$Started,[datetime]$Ended)
    $color=@{OK='#107c10';Warning='#ff8c00';Critical='#d13438'}
    $sb=[Text.StringBuilder]::new()
    [void]$sb.AppendLine('<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:Segoe UI,Arial;background:#f5f8fb;color:#1f2937;padding:24px}.card{background:#fff;border:1px solid #dde7f0;border-radius:8px;padding:16px;margin:0 0 16px}table{border-collapse:collapse;width:100%}th,td{border:1px solid #dde7f0;padding:7px;font-size:12px;text-align:left;vertical-align:top}th{background:#eef6fc}.rowWarning{background:#fff7e6}.rowCritical{background:#fde7e9}.pill{color:#fff;border-radius:999px;padding:4px 10px;font-weight:600}.kpi td{background:#f8fafc}.kpiLabel{font-size:11px;color:#64748b;text-transform:uppercase}.kpiValue{font-size:20px;font-weight:700;color:#0f172a}</style></head><body>')
    [void]$sb.AppendLine(("<div class='card'><h1>Microsoft Teams Inventory <span class='pill' style='background:{0}'>{1}</span></h1><p>RunId: {2}<br>Machine: {3}<br>Started UTC: {4}<br>Ended UTC: {5}<br>Duration: {6}<br>Teams processed: {7}</p></div>" -f $color[$Worst],$Worst,$RunId,$env:COMPUTERNAME,(IsoUtc $Started),(IsoUtc $Ended),((New-TimeSpan -Start $Started -End $Ended).ToString()),$Summary.TotalTeams))
    $kpiKeys=@('TotalTeams','ActiveTeams','InactiveTeams','ArchivedTeams','PublicTeams','PrivateTeams','TeamsWithGuests','MemberRows','ChannelRows','GuestRows','CriticalCount','WarningCount')
    $kpiCells=@($kpiKeys|ForEach-Object{if($Summary.ContainsKey($_)){"<td><div class='kpiLabel'>{0}</div><div class='kpiValue'>{1}</div></td>" -f [Net.WebUtility]::HtmlEncode($_),[Net.WebUtility]::HtmlEncode([string]$Summary[$_])}})
    [void]$sb.AppendLine(("<div class='card'><h2>Global summary</h2><table class='kpi'><tr>{0}</tr></table></div>" -f ($kpiCells -join '')))
    [void]$sb.AppendLine(("<div class='card'><b>Summary</b>: Total={0}; Active={1}; Inactive={2}; Archived={3}; Public={4}; Private={5}; TeamsWithGuests={6}; Critical={7}; Warnings={8}</div>" -f $Summary.TotalTeams,$Summary.ActiveTeams,$Summary.InactiveTeams,$Summary.ArchivedTeams,$Summary.PublicTeams,$Summary.PrivateTeams,$Summary.TeamsWithGuests,$Summary.CriticalCount,$Summary.WarningCount))
    [void]$sb.AppendLine('<div class="card"><h2>Critical and warning findings</h2><table><tr><th>Status</th><th>Team</th><th>Check</th><th>Value</th><th>Threshold</th><th>Details</th></tr>')
    foreach($a in @($AlertRows|Sort-Object @{Expression={if($_.Status-eq'Critical'){0}else{1}}},TeamDisplayName,Check|Select-Object -First 200)){
        [void]$sb.AppendLine(("<tr class='row{0}'><td>{0}</td><td>{1}</td><td>{2}</td><td>{3} {4}</td><td>{5}</td><td>{6}</td></tr>" -f $a.Status,[Net.WebUtility]::HtmlEncode($a.TeamDisplayName),[Net.WebUtility]::HtmlEncode($a.Check),[Net.WebUtility]::HtmlEncode([string]$a.NumericValue),[Net.WebUtility]::HtmlEncode([string]$a.TextValue),[Net.WebUtility]::HtmlEncode($a.Threshold),[Net.WebUtility]::HtmlEncode($a.Details)))
    }
    [void]$sb.AppendLine('</table></div></body></html>')
    $sb.ToString()
}
if([string]::IsNullOrWhiteSpace($OutputPath)){$OutputPath=[string](Get-ConfigValue 'TeamsInventoryCsvLogFolderPath' '{{DataAllRootPath}}\M365\Teams\Inventory')}
$LatestCsvFolderPath=[string](Get-ConfigValue 'LatestCsvFolderPath' $TenantContext.LatestCsvFolderPath); $WeeklyHistoryFolderPath=[string](Get-ConfigValue 'WeeklyHistoryFolderPath' (Join-Path $OutputPath 'WeeklyHistory')); $WeeklyHistoryRetentionWeeks=[int](Get-ConfigValue 'WeeklyHistoryRetentionWeeks' 52); $EnableWeeklyHistory=[bool](Get-ConfigValue 'EnableWeeklyHistory' $true)
if(-not$PSBoundParameters.ContainsKey('AlwaysSend')){$AlwaysSend=[bool](Get-ConfigValue 'AlwaysSend' $false)}; if(-not$PSBoundParameters.ContainsKey('RequireSensitivityLabel')){$RequireSensitivityLabel=[bool](Get-ConfigValue 'RequireSensitivityLabel' $false)}; if(-not$PSBoundParameters.ContainsKey('IncludeChannelOwners')){$IncludeChannelOwners=[bool](Get-ConfigValue 'IncludeChannelOwners' $false)}
$global:RetentionMaxCSV=[int](Get-ConfigValue 'RetentionMaxCSV' 30); $global:RetentionMaxLogs=[int](Get-ConfigValue 'RetentionMaxLogs' 30)
$global:EnableSharePointUpload=[bool](Get-ConfigValue 'EnableSharePointUpload' $false); $global:SharePointSiteHostname=[string](Get-ConfigValue 'SharePointSiteHostname' ''); $global:SharePointSitePath=[string](Get-ConfigValue 'SharePointSitePath' ''); $global:SharePointLibraryDisplayName=[string](Get-ConfigValue 'SharePointLibraryDisplayName' 'Documents'); $global:SharePointTargetFolderPath=[string](Get-ConfigValue 'SharePointTargetFolderPath' '')
$AppId=[string](Get-ConfigValue 'AppId' ''); $TenantId=[string](Get-ConfigValue 'TenantId' ''); $OrgDomain=[string](Get-ConfigValue 'OrgDomain' ''); $Thumb=[string](Get-ConfigValue 'Thumbprint' (Get-ConfigValue 'Thumb' ''))
$global:AppId=$AppId; $global:TenantId=$TenantId; $global:OrgDomain=$OrgDomain; $global:Thumb=$Thumb; $global:Thumbprint=$Thumb
$teamColumns=@('RunId','RunDateUtc','TenantName','TeamId','TeamDisplayName','Description','Visibility','CreatedDateTimeUtc','Classification','SensitivityLabel','IsArchived','OwnerCount','MemberCount','GuestCount','StandardChannelCount','PrivateChannelCount','SharedChannelCount','LastActivityDateUtc','InactiveDays','StorageUsedGB','StorageQuotaGB','StorageQuotaPercent','Status','NumericValue','TextValue','Threshold','Details')
$memberColumns=@('RunId','RunDateUtc','TenantName','TeamId','TeamDisplayName','UserId','DisplayName','UserPrincipalName','Mail','UserType','Role','Status','NumericValue','TextValue','Threshold','Details')
$channelColumns=@('RunId','RunDateUtc','TenantName','TeamId','TeamDisplayName','ChannelId','ChannelDisplayName','MembershipType','CreatedDateTimeUtc','PrivateChannelOwners','Status','NumericValue','TextValue','Threshold','Details')
$guestColumns=@('RunId','RunDateUtc','TenantName','TeamId','TeamDisplayName','UserId','DisplayName','UserPrincipalName','Mail','Status','NumericValue','TextValue','Threshold','Details')
try{
 $CurrentOperation='Initialize script environment'; $OutputPath=InitializeScriptEnvironment -OutputPathInit $OutputPath -LogFileName $ScriptBaseName; Start-Transcript -Path $global:logTranscriptFile -Append|Out-Null; Write-SmartM365LoadedModuleVersions; WriteLog -Message "Starting $TaskName"
 $CurrentOperation='Connect Microsoft Graph'; Disconnect-SmartM365CloudSession -ExchangeOnline:$false -Graph:$true -VerboseDisconnect:$true; $conn=Connect-SmartM365CloudSession -ExchangeOnline:$false -Graph:$true -AppId $AppId -Thumbprint $Thumb -TenantId $TenantId -Organization $OrgDomain -GraphScopes @('Team.ReadBasic.All','TeamMember.Read.All','Channel.ReadBasic.All','Group.Read.All','Reports.Read.All'); if(-not$conn.GraphConnected){throw 'Microsoft Graph app-only connection failed.'}
 $CurrentOperation='Run preflight'; Invoke-SmartM365Preflight -ScriptName $TaskName -RequiredModules @('Microsoft.Graph.Authentication') -OutputPaths @($OutputPath) -RequiredGraphApplicationPermissions @('Team.ReadBasic.All','TeamMember.Read.All','Channel.ReadBasic.All','Group.Read.All','Reports.Read.All','Sites.Read.All') -GraphProbeUris @('https://graph.microsoft.com/v1.0/organization','https://graph.microsoft.com/v1.0/groups?$top=1')|Out-Null
 $CurrentOperation='Load tenant metadata'; $org=Invoke-Graph -Uri 'https://graph.microsoft.com/v1.0/organization?$select=displayName' -Operation 'Get organization'; $TenantName=[string]@($org.value)[0].displayName; if([string]::IsNullOrWhiteSpace($TenantName)){$TenantName=$Tenant}
 $activityById=@{}; foreach($r in (Get-ReportRow -ReportName 'getTeamsTeamActivityDetail' -Period 'D180')){$id=[string](Prop $r @('Team Id','TeamId','Team ID')); if($id){$activityById[$id]=$r}}
 $teamFilter=[uri]::EscapeDataString("resourceProvisioningOptions/Any(x:x eq 'Team')"); $teamsUri="https://graph.microsoft.com/v1.0/groups?`$filter=$teamFilter&`$select=id,displayName,description,visibility,createdDateTime,classification,assignedLabels,mail,webUrl&`$top=999"; $teams=@(Get-GraphCollection -Uri $teamsUri -Operation 'Get team groups'); if($MaxTeams-gt 0){$teams=@($teams|Select-Object -First $MaxTeams)}; WriteLog -Message ("Teams discovered: {0}" -f $teams.Count)
 $i=0; foreach($g in $teams){$i++; $teamId=[string]$g.id; $teamName=[string]$g.displayName; WriteLog -Message ("Processing team {0}/{1}: {2}" -f $i,$teams.Count,$teamName); $CurrentOperation="Process team $teamName"
  $details=$null; try{$details=Invoke-Graph -Uri ("https://graph.microsoft.com/v1.0/teams/{0}" -f $teamId) -Operation 'Get team details'}catch{WriteLog -Message ("Team details unavailable for {0}: {1}" -f $teamName,$_.Exception.Message) -Level WARNING}; $archived=if($details-and$details.PSObject.Properties['isArchived']){[bool]$details.isArchived}else{$false}
  $labels=@(); foreach($l in @($g.assignedLabels)){$labels += [string](if($l.displayName){$l.displayName}else{$l.labelId})}; $label=JoinVals $labels
  $owners=@(Get-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/owners/microsoft.graph.user?`$select=id,displayName,userPrincipalName,mail,userType&`$top=999" -f $teamId) -Operation 'Get owners'); $members=@(Get-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/members/microsoft.graph.user?`$select=id,displayName,userPrincipalName,mail,userType&`$top=999" -f $teamId) -Operation 'Get members')
  $ownerIds=@{}; foreach($o in $owners){$ownerIds[[string]$o.id]=$true}; $guests=@($members|Where-Object{[string]$_.userType-eq'Guest'})
  foreach($m in $members){$role=if($ownerIds.ContainsKey([string]$m.id)){'Owner'}else{'Member'}; [void]$MembersRows.Add([pscustomobject]@{RunId=$RunId;RunDateUtc=$RunDateUtc;TenantName=$TenantName;TeamId=$teamId;TeamDisplayName=$teamName;UserId=[string]$m.id;DisplayName=[string]$m.displayName;UserPrincipalName=[string]$m.userPrincipalName;Mail=[string]$m.mail;UserType=[string]$m.userType;Role=$role;Status='OK';NumericValue='';TextValue=$role;Threshold='Inventory only';Details=''})}
  foreach($guest in $guests){[void]$GuestsRows.Add([pscustomobject]@{RunId=$RunId;RunDateUtc=$RunDateUtc;TenantName=$TenantName;TeamId=$teamId;TeamDisplayName=$teamName;UserId=[string]$guest.id;DisplayName=[string]$guest.displayName;UserPrincipalName=[string]$guest.userPrincipalName;Mail=[string]$guest.mail;Status='Warning';NumericValue='1';TextValue='Guest';Threshold="Guests <= $GuestWarningThreshold";Details='External guest member'})}
  $channels=@(); try{$channels=@(Get-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/teams/{0}/channels" -f $teamId) -Operation 'Get channels')}catch{WriteLog -Message ("Channels unavailable for {0}: {1}" -f $teamName,$_.Exception.Message) -Level WARNING}
  $standard=@($channels|Where-Object{[string]$_.membershipType-in@('','standard')}).Count; $private=@($channels|Where-Object{[string]$_.membershipType-eq'private'}).Count; $shared=@($channels|Where-Object{[string]$_.membershipType-eq'shared'}).Count
  foreach($ch in $channels){$chOwners=@(); if($IncludeChannelOwners-and [string]$ch.membershipType-in@('private','shared')){try{$cm=@(Get-GraphCollection -Uri ("https://graph.microsoft.com/v1.0/teams/{0}/channels/{1}/members?`$top=200" -f $teamId,$ch.id) -Operation 'Get channel members'); $chOwners=@($cm|Where-Object{@($_.roles)-contains'owner'}|ForEach-Object{$_.displayName})}catch{$chOwners=@('NotMeasured: ChannelMember.Read.All may be required')}}; [void]$ChannelsRows.Add([pscustomobject]@{RunId=$RunId;RunDateUtc=$RunDateUtc;TenantName=$TenantName;TeamId=$teamId;TeamDisplayName=$teamName;ChannelId=[string]$ch.id;ChannelDisplayName=[string]$ch.displayName;MembershipType=[string]$ch.membershipType;CreatedDateTimeUtc=(IsoUtc $ch.createdDateTime);PrivateChannelOwners=(JoinVals $chOwners);Status='OK';NumericValue='1';TextValue=[string]$ch.membershipType;Threshold='Inventory only';Details=''})}
  $last=''; $inactive=''; $act=$activityById[$teamId]; if($act){$la=Prop $act @('Last Activity Date','LastActivityDate'); $last=IsoUtc $la; if($last){$inactive=[math]::Round(((Get-Date).ToUniversalTime()-([datetime]$la).ToUniversalTime()).TotalDays,0)}}
  $usedGb=''; $quotaGb=''; $quotaPct=''; $storageDetail='Storage not measured'; try{$drive=Invoke-Graph -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/sites/root/drive?`$select=quota,webUrl" -f $teamId) -Operation 'Get team SharePoint quota'; if($drive.quota-and[double]$drive.quota.total-gt 0){$usedGb=[math]::Round([double]$drive.quota.used/1GB,2); $quotaGb=[math]::Round([double]$drive.quota.total/1GB,2); $quotaPct=[math]::Round(([double]$drive.quota.used/[double]$drive.quota.total)*100,2); $storageDetail=[string]$drive.webUrl}}catch{$storageDetail='NotMeasured: '+$_.Exception.Message}
  $status='OK'; $notes=New-Object 'System.Collections.Generic.List[string]'; if($owners.Count-eq 0){$status='Critical'; [void]$notes.Add('No owner'); Add-Alert $teamId $teamName Critical Owners $owners.Count 'NoOwner' "Owners >= $MinOwners" 'Team has no owner.'}elseif($owners.Count-lt$MinOwners){if($status-ne'Critical'){$status='Warning'}; [void]$notes.Add('Owner count below threshold'); Add-Alert $teamId $teamName Warning Owners $owners.Count 'LowOwnerCount' "Owners >= $MinOwners" 'Team has too few owners.'}
  if($inactive-ne'' -and [double]$inactive-gt$InactiveDays -and -not$archived){if($status-ne'Critical'){$status='Warning'}; [void]$notes.Add('Inactive team'); Add-Alert $teamId $teamName Warning Inactivity $inactive $last "<= $InactiveDays days" 'Team is a candidate for archival.'}
  if($quotaPct-ne'' -and [double]$quotaPct-gt$QuotaCriticalPercent){$status='Critical'; [void]$notes.Add('Storage quota critical'); Add-Alert $teamId $teamName Critical StorageQuotaPercent $quotaPct $storageDetail "<= $QuotaCriticalPercent percent" 'SharePoint storage quota usage is above threshold.'}
  if($guests.Count-gt$GuestWarningThreshold){if($status-ne'Critical'){$status='Warning'}; [void]$notes.Add('High guest count'); Add-Alert $teamId $teamName Warning GuestCount $guests.Count Guests "<= $GuestWarningThreshold" 'Team has many external guests.'}
  if([string]$g.visibility-eq'Public' -and -not[string]::IsNullOrWhiteSpace($label)){if($status-ne'Critical'){$status='Warning'}; [void]$notes.Add('Public team with sensitivity label'); Add-Alert $teamId $teamName Warning PublicSensitiveLabel 1 $label 'Review public sensitive teams' 'Public team has a sensitivity/classification label.'}
  if($RequireSensitivityLabel -and [string]::IsNullOrWhiteSpace($label)){if($status-ne'Critical'){$status='Warning'}; [void]$notes.Add('Missing sensitivity label'); Add-Alert $teamId $teamName Warning MissingSensitivityLabel 0 NoLabel 'Sensitivity label required' 'Tenant policy expects labels.'}
  [void]$TeamsRows.Add([pscustomobject]@{RunId=$RunId;RunDateUtc=$RunDateUtc;TenantName=$TenantName;TeamId=$teamId;TeamDisplayName=$teamName;Description=[string]$g.description;Visibility=[string]$g.visibility;CreatedDateTimeUtc=(IsoUtc $g.createdDateTime);Classification=[string]$g.classification;SensitivityLabel=$label;IsArchived=[string]$archived;OwnerCount=$owners.Count;MemberCount=$members.Count;GuestCount=$guests.Count;StandardChannelCount=$standard;PrivateChannelCount=$private;SharedChannelCount=$shared;LastActivityDateUtc=$last;InactiveDays=(Num $inactive);StorageUsedGB=(Num $usedGb);StorageQuotaGB=(Num $quotaGb);StorageQuotaPercent=(Num $quotaPct);Status=$status;NumericValue=(Num $members.Count);TextValue="Owners=$($owners.Count); Members=$($members.Count); Guests=$($guests.Count)";Threshold="Owners >= $MinOwners; inactive <= $InactiveDays days; storage <= $QuotaCriticalPercent percent; guests <= $GuestWarningThreshold";Details=(JoinVals $notes)})
 }
 $stamp=(Get-Date).ToUniversalTime().ToString('yyyyMMdd_HHmmss',[Globalization.CultureInfo]::InvariantCulture)
 Export-InventoryCsv -Rows $TeamsRows.ToArray() -Columns $teamColumns -TimestampedPath (Join-Path $OutputPath "M365_Teams_Teams_$stamp.csv") -LatestPath (Join-Path $LatestCsvFolderPath 'M365_Teams_Teams.csv') -HistoryPath (Join-Path $OutputPath 'M365_Teams_Teams_History.csv')
 Export-InventoryCsv -Rows $MembersRows.ToArray() -Columns $memberColumns -TimestampedPath (Join-Path $OutputPath "M365_Teams_Members_$stamp.csv") -LatestPath (Join-Path $LatestCsvFolderPath 'M365_Teams_Members.csv') -HistoryPath (Join-Path $OutputPath 'M365_Teams_Members_History.csv')
 Export-InventoryCsv -Rows $ChannelsRows.ToArray() -Columns $channelColumns -TimestampedPath (Join-Path $OutputPath "M365_Teams_Channels_$stamp.csv") -LatestPath (Join-Path $LatestCsvFolderPath 'M365_Teams_Channels.csv') -HistoryPath (Join-Path $OutputPath 'M365_Teams_Channels_History.csv')
 Export-InventoryCsv -Rows $GuestsRows.ToArray() -Columns $guestColumns -TimestampedPath (Join-Path $OutputPath "M365_Teams_Guests_$stamp.csv") -LatestPath (Join-Path $LatestCsvFolderPath 'M365_Teams_Guests.csv') -HistoryPath (Join-Path $OutputPath 'M365_Teams_Guests_History.csv')
 if($EnableWeeklyHistory-and-not$DryRun){Add-SmartM365WeeklyHistory -SourceCsvPaths $GeneratedCsvPaths.ToArray() -HistoryRootPath $WeeklyHistoryFolderPath -RetentionWeeks $WeeklyHistoryRetentionWeeks -HistoryLabel 'Microsoft Teams inventory'|Out-Null}elseif($DryRun){WriteLog -Message 'DryRun enabled: WeeklyHistory skipped.' -Level INFO}
 $teamArray=$TeamsRows.ToArray(); $memberArray=$MembersRows.ToArray(); $channelArray=$ChannelsRows.ToArray(); $guestArray=$GuestsRows.ToArray(); $alertArray=$Alerts.ToArray(); $csvArray=$GeneratedCsvPaths.ToArray()
 $summary=@{TotalTeams=$teamArray.Count;ActiveTeams=@($teamArray|Where-Object{$_.IsArchived-ne'True' -and ([string]::IsNullOrWhiteSpace([string]$_.InactiveDays)-or [double]$_.InactiveDays-le$InactiveDays)}).Count;InactiveTeams=@($teamArray|Where-Object{-not[string]::IsNullOrWhiteSpace([string]$_.InactiveDays)-and [double]$_.InactiveDays-gt$InactiveDays}).Count;ArchivedTeams=@($teamArray|Where-Object{$_.IsArchived-eq'True'}).Count;PublicTeams=@($teamArray|Where-Object{$_.Visibility-eq'Public'}).Count;PrivateTeams=@($teamArray|Where-Object{$_.Visibility-eq'Private'}).Count;TeamsWithGuests=@($teamArray|Where-Object{[int]$_.GuestCount-gt 0}).Count;CriticalCount=@($alertArray|Where-Object Status -eq Critical).Count;WarningCount=@($alertArray|Where-Object Status -eq Warning).Count;MemberRows=$memberArray.Count;ChannelRows=$channelArray.Count;GuestRows=$guestArray.Count}
 $worst=WorstStatus $alertArray; $subject="[$($worst.ToUpperInvariant())] Microsoft Teams Inventory - $TenantName - $RunDateUtc"; $html=ConvertTo-HtmlReport -AlertRows $alertArray -Summary $summary -Worst $worst -Started $RunStarted -Ended (Get-Date)
 if(-not$DryRun -and ($AlwaysSend -or $worst-ne'OK')){Send-SmartM365Mail -Subject $subject -BodyHtml $html -Attachments @($csvArray|Select-Object -First 4)}elseif($DryRun){WriteLog -Message 'DryRun enabled: email skipped.' -Level INFO}else{WriteLog -Message 'No warning/critical finding and AlwaysSend disabled: email skipped.' -Level INFO}
 $result="Teams=$($summary.TotalTeams); Critical=$($summary.CriticalCount); Warnings=$($summary.WarningCount); Members=$($memberArray.Count); Channels=$($channelArray.Count); Guests=$($guestArray.Count)"; try{Stop-Transcript|Out-Null; Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile}catch{$null=$_}; WriteLog -Message ("Result summary: $result") -Level INFO; Complete-SmartM365ExecutionContext -Status $(if($worst-eq'OK'){'Success'}else{'CompletedWithWarnings'}); Write-Host "Teams inventory completed. Status=$worst; $result"
}catch{ $err=$_; try{WriteLog -Message ("Teams inventory failed during {0}: {1}" -f $CurrentOperation,$err.Exception.Message) -Level ERROR}catch{$null=$_}; try{Stop-Transcript|Out-Null; Update-SmartM365TimestampedTranscript -Path $global:logTranscriptFile}catch{$null=$_}; try{Complete-SmartM365ExecutionContext -Status Failed -ErrorRecord $err -FailureStage $CurrentOperation}catch{$null=$_}; throw }

# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAF4KQZn2Oy9KpI
# trncsx+EQheI10OaN/wkIBHfi2gSZ6CCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCDw1Pw0kO091BcLzLgzgc44VlABFCRj+qjQai2t/hRybDANBgkqhkiG9w0B
# AQEFAASCAYBIo4+5TatluaJy6mXntaalbssVS2Oufs1kYO1N4iJMHxPr1V3jKIMI
# PcsST1z+xd/UlMkwxybS8EySW/TSFkajShLqeQsSOD1Drr30M7sEOkKL9h/f3BTx
# 6nJT1/kEYVesRxzhT5bdfoxVRCPXuV+Q3h2OFor6lUJPuaeDjzUc+sQ5PqegFVWr
# sq6OvOwjrJ6euqcY2F4EJMup+uyhMO6BTbahfjQK4cRmZXPuBmNWDw3SNAVA2XwS
# 0DRHa61rioF/b8qexToyhuNbhQAf76zyCkIO5LNqJAP0bhA6rdp6105+puHGvVG5
# nx7juNrIAalkPvVw61s1QUhRfmYiXvGH0mIa1CLsJd2AtCDg+qSEfnkkOD212wZW
# j6K0BasRmnD2mxbbprKuzjwv1MwmHvaSv74oLPmW4pXNgqrmCkQZQTN3UYF+oSMK
# zJrXkiz6O/2gadjw/ZHpe4BDTUSITPrDYXO6shiFo2FSwx95kPWrxV4/G5l+z8nj
# AsOPR3I+wns=
# SIG # End signature block
