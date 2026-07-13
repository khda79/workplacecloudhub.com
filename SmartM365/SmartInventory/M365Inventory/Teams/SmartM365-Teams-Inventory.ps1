#Requires -Version 7.0
<#
.SYNOPSIS
    Microsoft Teams tenant inventory with CSV exports and HTML alert summary.
.VERSION
0.14

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
$ScriptVersion="0.15"
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
function Export-InventoryCsv{param([object[]]$Rows,[string[]]$Columns,[string]$TimestampedPath,[string]$LatestPath,[string]$HistoryPath) $Columns=@('TenantKey','OrganizationKey','EnvironmentKey','TenantId')+@($Columns|Where-Object{$_-inotmatch'^(TenantKey|OrganizationKey|EnvironmentKey|TenantId)$'}); Assert-SmartM365CsvDataCompleteness -Data $Rows -Columns $Columns -TimestampedPath $TimestampedPath -LatestPath $LatestPath; foreach($folder in @((Split-Path $TimestampedPath -Parent),(Split-Path $LatestPath -Parent))){if(-not(Test-Path -LiteralPath $folder)){New-Item -Path $folder -ItemType Directory -Force|Out-Null}}; if($Rows.Count-eq 0){$h=($Columns|ForEach-Object{'"'+($_-replace'"','""')+'"'})-join','; Set-Content -LiteralPath $TimestampedPath -Value $h -Encoding utf8BOM; Set-Content -LiteralPath $LatestPath -Value $h -Encoding utf8BOM}else{$Rows|Select-Object -Property $Columns|Add-SmartM365TenantKey | Export-Csv -LiteralPath $TimestampedPath -NoTypeInformation -Encoding utf8BOM; $Rows|Select-Object -Property $Columns|Add-SmartM365TenantKey | Export-Csv -LiteralPath $LatestPath -NoTypeInformation -Encoding utf8BOM}; [void]$GeneratedCsvPaths.Add($TimestampedPath); [void]$GeneratedCsvPaths.Add($LatestPath); if(-not$global:csvGeneratedPaths){$global:csvGeneratedPaths=New-Object 'System.Collections.Generic.HashSet[string]'([StringComparer]::OrdinalIgnoreCase)}; [void]$global:csvGeneratedPaths.Add($TimestampedPath); [void]$global:csvGeneratedPaths.Add($LatestPath); if($DryRun){WriteLog -Message 'DryRun enabled: SharePoint CSV upload skipped.' -Level INFO}else{Invoke-SmartM365SharePointCsvUpload -LocalFilePath $TimestampedPath|Out-Null; Invoke-SmartM365SharePointCsvUpload -LocalFilePath $LatestPath|Out-Null}; if($AppendHistory-and$HistoryPath){$hp=Split-Path $HistoryPath -Parent; if(-not(Test-Path -LiteralPath $hp)){New-Item -Path $hp -ItemType Directory -Force|Out-Null}; if($Rows.Count-gt 0){if(Test-Path -LiteralPath $HistoryPath){Repair-SmartM365CsvTenantKeySchema -Path $HistoryPath -Delimiter ',' -Encoding UTF8|Out-Null}; $Rows|Select-Object -Property $Columns|Add-SmartM365TenantKey | Export-Csv -LiteralPath $HistoryPath -NoTypeInformation -Encoding utf8BOM -Append:(Test-Path -LiteralPath $HistoryPath)}}}
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCMqnvCs7smn9gm
# 68frg2RCkCvuNUdEd03MdwE9gTv4vqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIHQLojhrl3Mmkoc1gFWy5PpWeucf5E3bqQ6JepSUwZtPMA0GCSqG
# SIb3DQEBAQUABIIBgKj+vPtLL4fS2F2GCUfDNQY9o4ECZoSk+Af9lCOPKLjI6FPm
# FVsdqj08C62pj5g08FqoHw2gFHskNn6EBL6CAnKLEuWBaWSWuxKpqvOBiFc+BjCk
# 2Fkr7k7a6pYPe1vH+P6VVbsrGnnRWvIXqs8UIdDZh99g35cC2P/C8XLNk+4En9vk
# pIGW/+t1f4URb47TAAZU+li+9ybpRaOkw6WnZ1UEjiWWB1lUq/2Jy+rsbT8ch3xC
# 07mCESSfpRcWDFeQoRrk9SxMyq1VbWJh60JZ2GQ00g/Zaqq88Dsw3cr+X59oV7YN
# uCSGQE3fRhkhL8jkQD0rJLRHTyAH8ugVxb5glyC7XEFSVrUBjzWGUcEy0+/JtAUq
# /GzZRTPSJ3RO7POeSqvpHl6AQvwTaMdtvbUD6YCUgKFbGfpGmZP0oiG88EFti+73
# 2CxzCd01hyH+VlOyySJJ1D2GSERdpVIT5Eb5PQZSBlADYbqWHQgnCwN+fAt1GXLl
# ohwzPfVKdNnGgMlnkqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MTMxMjAz
# MjdaMC8GCSqGSIb3DQEJBDEiBCB0B+2zNvxEq3zsQAAm/A3aO3J1VvQlcm97j8RC
# 4cTSoTANBgkqhkiG9w0BAQEFAASCAgBT/Oom7n1ln305B3eOdATBnYusxFk5cYo+
# Lwp9noxKLIIjqvXSR+nJf2At2o0bJ8Ig1Mse6HCbWbYKAP6rkBIry1Yrx81xft/b
# SBz2mAUrxyEDxN9C7TV9Bgm50vzM8XWYkxWIwWnD3Kz8gyzz+ltZ5JMchus0OQAJ
# LT+D/iua3ivXUdusqh0P0a9Zbai6RZrKbyMf/FZbS4+a/wlW4R02Hz0QF3EsOnGr
# HwoovkBdFl1TELvtf+0W+LP/iBfj2r50uljC7ToWRaVKWZmxJS/OGtIlwiq6tOqq
# 4pvgmP9CeYxa6HlXEoxkYUUmJELx5wU2b021m3Y8kySUpjTRrnZCKyipZQaFPm0J
# CjJcjUh1QZdYIKq3pujqqWNqk4mxMQTgC/T/9hgL/aKoQ97sRfpCVtrji/0PiHku
# K/RGU4dxiY1mmk4JbyoRm0NwABdMBryBYgh6LaVy5fx4klBAoLrN6C2XE8zCVQ5+
# /+X1PW7n8zPJkW5OC8XN1LM0jWX1nV2G4HtwUUzxl6a+iOYS5rAWu3TQIpNvwAZ9
# tM+zP5ih1qthtVA9xfWlbZJ6+dsKpFsnDd0CFM0PMw9ksX/OOH9FatnOpBbpAcf0
# z6in06bvIW2J01gNlba8B+IJJUBeq/MdRUzItuJismFVofGWLcVmt8pn/C+a5Zsw
# lk3AmetDmA==
# SIG # End signature block
