<#
.SYNOPSIS
Generates Exchange hybrid identity issue tables for PowerBI from SmartInventory CSV exports.

.REQUIREMENTS
- PowerShell 7+
- SmartM365.Core module
- Read access to DATA-LAST SmartInventory CSV exports
- Write access to the configured DATA-ALL, DATA-LAST, and WeeklyHistory output folders
- Required input CSV files: AD_Users_AllDomains_Enriched.csv when available, otherwise AD_Users_AllDomains.csv,
  Exchange_OnPrem_Mailboxes_AllDomains.csv, Exchange_OnPrem_RemoteMailboxes_AllDomains.csv,
  Exchange_EXO_Mailboxes_AllDomains.csv,
  M365_Users_Active.csv

.VERSION
1.11
#>
#requires -Version 7.0
[CmdletBinding()]
param(
  [string]$Tenant='test',
  [string]$DataLastFolder='',
  [string]$OutputFolder='',
  [string]$LatestFolder='',
  [switch]$DisableSharePointUpload,
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
    throw "-MaxItems is not supported by SmartM365-Exchange-HybridIdentity-Issues-Inventory because issue detection must compare complete source CSV snapshots."
}
$ErrorActionPreference='Stop'
$ScriptName='SmartM365-Exchange-HybridIdentity-Issues-Inventory'
$ScriptVersion="1.11"
$RunStamp=Get-Date -Format 'yyyyMMdd-HHmmss'
function Log([string]$m){ Write-Host ("{0} [INFO] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$m) }
function Warn([string]$m){ Write-Warning $m }
function Root(){ $d=$PSScriptRoot; while($d){ if((Test-Path (Join-Path $d 'Config\SmartM365-TenantContext.ps1')) -and (Test-Path (Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'))){return $d}; $p=Split-Path $d -Parent; if(!$p -or $p -eq $d){break}; $d=$p }; throw 'SmartM365 root not found.' }
function LocalConfig(){ $p=Join-Path $PSScriptRoot (([IO.Path]::GetFileNameWithoutExtension($PSCommandPath))+'.local.json'); if(!(Test-Path $p)){ $t="$p.template"; if(!(Test-Path $t)){throw "Missing config template: $t"}; Copy-Item $t $p; Log "Created local config: $p" }; Get-Content $p -Raw | ConvertFrom-Json }
function ResolveToken($v){ if($v -isnot [string] -or [string]::IsNullOrWhiteSpace($v)){return $v}; $r=[string]$v; for($iteration=0;$iteration -lt 10;$iteration++){ $ms=[regex]::Matches($r,'\{\{(?<n>[A-Za-z0-9_.-]+)\}\}'); if($ms.Count -eq 0){break}; $changed=$false; foreach($m in $ms){ $pr=$script:Cfg.PSObject.Properties[$m.Groups['n'].Value]; if($pr){$replacement=ResolveToken $pr.Value; if($replacement -is [array]){throw "Configuration token '$($m.Value)' resolved to multiple values."}; $next=$r.Replace($m.Value,[string]$replacement); if($next -ne $r){$changed=$true;$r=$next}} }; if(!$changed){break} }; return $r }
function AssertResolvedPath([string]$Name,[string]$Path){ if([string]::IsNullOrWhiteSpace($Path)){throw "Configuration path '$Name' is empty."}; if($Path -match '\{\{[A-Za-z0-9_.-]+\}\}'){throw "Configuration path '$Name' contains an unresolved token: $Path"}; if(([regex]::Matches($Path,'(?i)(?:[A-Z]:\\|\\\\)').Count) -gt 1){throw "Configuration path '$Name' contains multiple path roots: $Path"} }
function Cfg($c,$n,$d){ $p=$c.PSObject.Properties[$n]; if($p -and $null -ne $p.Value){ if($p.Value -isnot [string]){return ResolveToken $p.Value}; $txt=$p.Value.Trim(); if($txt -and $txt -notin @('__USE_GLOBAL__','USE_GLOBAL')){return ResolveToken $p.Value} }; $gp=$script:Cfg.PSObject.Properties[$n]; if($gp -and $null -ne $gp.Value){ if($gp.Value -is [string] -and [string]::IsNullOrWhiteSpace($gp.Value)){return $d}; return ResolveToken $gp.Value }; $d }
function P($r,[string[]]$names){ if($null -eq $r){return $null}; foreach($n in $names){$p=$r.PSObject.Properties[$n]; if($p -and $null -ne $p.Value){return $p.Value}}; $null }
function T($v){ if($null -eq $v){''}else{([string]$v).Trim()} }
function K($v){ (T $v).ToLowerInvariant() }
function B($v){ (K $v) -in @('true','1','yes','y','enabled') }
function CB($c,$n,[bool]$d){ $v=Cfg $c $n $d; if($v -is [bool]){return $v}; if($null -eq $v){return $d}; $s=K $v; if(!$s){return $d}; $s -in @('true','1','yes','y','enabled') }
function Dbl($v){ $s=(T $v)-replace ',','.'; $o=0.0; if([double]::TryParse($s,[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$o)){$o}else{0.0} }
function Dt($v){ $s=T $v; if(!$s){return $null}; foreach($c in @([Globalization.CultureInfo]::InvariantCulture,[Globalization.CultureInfo]::GetCultureInfo('fr-FR'),[Globalization.CultureInfo]::GetCultureInfo('en-US'))){$d=[datetime]::MinValue; if([datetime]::TryParse($s,$c,[Globalization.DateTimeStyles]::AssumeLocal,[ref]$d)){return $d}}; $null }
function AddMap($map,$key,$val){ $k=K $key; if($k -and !$map.ContainsKey($k)){$map[$k]=$val} }
function Csv($name,[switch]$Req){ $p=Join-Path $DataLastFolder $name; if(!(Test-Path $p)){ if($Req){throw "Required CSV not found: $p"}; Warn "Optional CSV not found: $p"; return @()}; $r=@(Import-Csv $p); Log "Loaded $name : $($r.Count) row(s)"; $r }
function AddIssue($list,[int]$num,[string]$guid,[string]$issue,[string]$cat,[string]$action){
  if(!$guid){$guid='UNRESOLVED'}
  $key="{0}|{1}" -f $guid,$num
  if($script:IssueKeys -and $script:IssueKeys.Contains($key)){return}
  if($script:IssueKeys){[void]$script:IssueKeys.Add($key)}
  [void]$list.Add([pscustomobject]@{IssueNumber=$num;ObjectGUID=$guid;Potential_Issue=$issue;IssueCategory=$cat;RecommendedAction=$action})
}
function SplitAddr($v){ $s=T $v; if(!$s){@()}else{@($s -split '[;|,]'|%{(T $_)-replace '^(smtp|SMTP):',''}|?{$_})} }
function NewSet(){[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)}
function AddSet($set,$value){$v=T $value; if($v){[void]$set.Add($v)}}
function DomainPart($value){$s=(T $value) -replace '^(smtp|SMTP):',''; if($s -match '@'){$s.Split('@')[-1].ToLowerInvariant()}else{''}}
function LocalPart($value){$s=(T $value) -replace '^(smtp|SMTP):',''; if($s -match '@'){$s.Split('@')[0]}else{$s}}
function HasSkuLike($licenses,[string]$pattern){foreach($l in @($licenses)){if((T $l) -match $pattern){return $true}}; $false}
function IsUserMailbox($lm,$rm,$em){((T(P $lm RecipientTypeDetails,RecipientType))+' '+(T(P $rm RecipientTypeDetails,RecipientType))+' '+(T(P $em RecipientTypeDetails,MailboxType))) -match 'UserMailbox|RemoteUserMailbox'}
function IsSharedMailbox($lm,$rm,$em){((T(P $lm RecipientTypeDetails,RecipientType))+' '+(T(P $rm RecipientTypeDetails,RecipientType))+' '+(T(P $em RecipientTypeDetails,MailboxType))) -match 'Shared'}
function IsRoomMailbox($lm,$rm,$em){((T(P $lm RecipientTypeDetails,RecipientType))+' '+(T(P $rm RecipientTypeDetails,RecipientType))+' '+(T(P $em RecipientTypeDetails,MailboxType))) -match 'Room|Equipment'}
function ContainsInvalidUpnChar($value){
  $s=T $value
  if(!$s){return $false}
  foreach($charValue in @(' ','#','%','&','~','+','/','=','?','*','{','}','|','<','>','(',')',';',':',',','[',']','"','\','`')){
    if($s.Contains($charValue)){return $true}
  }
  $false
}
function ContainsInvalidSmtpLocalChar($value){(LocalPart $value) -match '[\\/\[\]:;\|=,+*?<>\s]'}
function PrimaryProxyCount($proxyText){ return @((T $proxyText) -split '[;|,]' | Where-Object { $_ -cmatch '^SMTP:' }).Count }
function ProxyContains($proxyText,$address){$a=(T $address) -replace '^(smtp|SMTP):',''; foreach($p in SplitAddr $proxyText){if($p -ieq $a){return $true}}; $false}
function IsFalseLike($value){$s=K $value; $s -in @('false','0','no','n','disabled')}
function LicensesFor($upn,$userId=''){ $id=K $userId; if($id -and $licUserId.ContainsKey($id)){return @($licUserId[$id])}; $k=K $upn; if($k -and $licUpn.ContainsKey($k)){return @($licUpn[$k])}; @() }
function GetM365User($upn,$immutableId){ $id=K $immutableId; if($id -and $m365Immutable.ContainsKey($id)){return $m365Immutable[$id]}; $u=K $upn; if($u -and $m365Upn.ContainsKey($u)){return $m365Upn[$u]}; $null }
function GetRecipientTypeFromMailboxes($primary,$lm,$rm,$em){ $smtp=K $primary; if($smtp -and $localSmtp.ContainsKey($smtp)){return T(P $localSmtp[$smtp] RecipientType,RecipientTypeDetails)}; if($smtp -and $remoteSmtp.ContainsKey($smtp)){return T(P $remoteSmtp[$smtp] RecipientTypeDetails,RecipientType)}; if($smtp -and $exoSmtp.ContainsKey($smtp)){return T(P $exoSmtp[$smtp] RecipientTypeDetails,MailboxType)}; $fallback=T(P $lm RecipientType,RecipientTypeDetails); if($fallback){return $fallback}; $fallback=T(P $rm RecipientTypeDetails,RecipientType); if($fallback){return $fallback}; $fallback=T(P $em RecipientTypeDetails,MailboxType); if($fallback){return $fallback}; 'NoMailboxes' }
function TestAllowedLicenseGroup($licenseGroup){ (T $licenseGroup) -in @('Microsoft 365 E3','Microsoft 365 F3','Microsoft 365 F1') }
function GetLicenseGroupFromM365($recipientType,$licenses,$inM365){
  if((T $recipientType) -in @('SharedMailbox','RemoteSharedMailbox')){return 'No License (Shared)'}
  if((T $recipientType) -in @('RoomMailbox','RemoteRoomMailbox','EquipmentMailbox','RemoteEquipmentMailbox')){return 'No License (Resource Mailbox)'}
  if(-not $inM365){return 'Not in M365'}
  $text=((@($licenses)|ForEach-Object{T $_}) -join ';')
  if([string]::IsNullOrWhiteSpace($text)){return 'No License (Unlicensed)'}
  if($text -match 'Microsoft 365 E3|ENTERPRISEPACK|SPE_E3'){return 'Microsoft 365 E3'}
  if($text -match 'Microsoft 365 F3|SPE_F1|SPE_F3|M365_F3|O365_F3'){return 'Microsoft 365 F3'}
  if($text -match 'Microsoft 365 F1|M365_F1|O365_F1'){return 'Microsoft 365 F1'}
  'No License (Unlicensed)'
}
function ExchangePlanStatus($userId){
  $id=K $userId
  if($id -and $planUserId.ContainsKey($id)){
    foreach($plan in @($planUserId[$id])){
      if((T(P $plan PlanName)) -in @('EXCHANGE_S_ENTERPRISE','EXCHANGE_S_DESKLESS') -and (B(P $plan IsEnabled))){
        $display=T(P $plan PlanDisplayName); if($display){return $display}; return T(P $plan PlanName)
      }
    }
  }
  'Not Exchange Plan Enabled'
}
function FlattenPath($value){
  foreach($item in @($value)){
    if($item -is [System.Array]){FlattenPath $item; continue}
    if($null -ne $item -and -not [string]::IsNullOrWhiteSpace([string]$item)){[string]$item}
  }
}
function CopyCsv($s,$d){
  $sources=@(FlattenPath $s)
  $dests=@(FlattenPath $d)
  foreach($source in $sources){
    foreach($dest in $dests){
      $dir=Split-Path $dest -Parent
      if(!(Test-Path $dir)){New-Item -ItemType Directory -Path $dir -Force|Out-Null}
      Copy-Item -LiteralPath $source -Destination $dest -Force
    }
  }
}
function ExportIssues($rows){
  foreach($f in @($OutputFolder,$LatestFolder,(Join-Path $OutputFolder 'Archive'))){ if(!(Test-Path $f)){New-Item -ItemType Directory -Path $f -Force|Out-Null} }
  $main=Join-Path $OutputFolder 'Exchange_HybridIdentity_Issues.csv'; $latest=Join-Path $LatestFolder 'Exchange_HybridIdentity_Issues.csv'; $archive=Join-Path (Join-Path $OutputFolder 'Archive') "Exchange_HybridIdentity_Issues_$RunStamp.csv"
  Assert-SmartM365CsvDataCompleteness -Data $rows -TimestampedPath $main -LatestPath $latest; $rows|Sort-Object IssueNumber,ObjectGUID,Potential_Issue|Add-SmartM365TenantKey | Export-Csv $main -NoTypeInformation -Encoding UTF8; CopyCsv $main $latest; CopyCsv $main $archive
  $summary=@($rows|Group-Object IssueNumber,Potential_Issue,IssueCategory|Sort-Object Count -Descending|%{ $p=$_.Name -split ', ',3; [pscustomobject]@{IssueNumber=[int]$p[0];Potential_Issue=$p[1];IssueCategory=$p[2];Count=$_.Count} })
  $sm=Join-Path $OutputFolder 'Exchange_HybridIdentity_Issues_Summary.csv'; $sl=Join-Path $LatestFolder 'Exchange_HybridIdentity_Issues_Summary.csv'; Assert-SmartM365CsvDataCompleteness -Data $summary -TimestampedPath $sm -LatestPath $sl; $summary|Add-SmartM365TenantKey | Export-Csv $sm -NoTypeInformation -Encoding UTF8; CopyCsv $sm $sl
  $files=@($main,$latest,$archive,$sm,$sl)
  $files
}
function PublishWeeklyHistory($files){
  if(!(CB $lc 'EnableWeeklyHistory' $true)){return @()}
  $base=Cfg $lc 'WeeklyHistoryFolderPath' (Join-Path $OutputFolder 'WeeklyHistory')
  if([string]::IsNullOrWhiteSpace($base)){return @()}
  $year=[Globalization.ISOWeek]::GetYear((Get-Date)); $week=[Globalization.ISOWeek]::GetWeekOfYear((Get-Date)); $folder=Join-Path $base ('{0}-W{1:00}' -f $year,$week)
  if(!(Test-Path $folder)){New-Item -ItemType Directory -Path $folder -Force|Out-Null}
  $published=[Collections.Generic.List[string]]::new()
  $publishedNames=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach($f in @($files)){
    if(!$f -or !(Test-Path -LiteralPath $f)){continue}
    $leaf=Split-Path $f -Leaf
    if(-not $publishedNames.Add($leaf)){continue}
    $dest=Join-Path $folder $leaf
    CopyCsv $f $dest
    [void]$published.Add($dest)
  }
  $manifest=Join-Path $folder 'manifest.json'
  [pscustomobject]@{ScriptName=$ScriptName;ScriptVersion=$ScriptVersion;Tenant=$Tenant;GeneratedOn=(Get-Date).ToString('o');Week=('{0}-W{1:00}' -f $year,$week);Files=@($published|ForEach-Object{Split-Path $_ -Leaf})} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifest -Encoding UTF8
  [void]$published.Add($manifest)
  @($published)
}
try{
  Log "Starting $ScriptName v$ScriptVersion"
  $sr=Root; . (Join-Path $sr 'Config\SmartM365-TenantContext.ps1'); $script:Cfg=Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot; Import-Module -Name (Join-Path $sr 'Modules\SmartM365.Core\SmartM365.Core.psd1') -MinimumVersion '1.0.24' -Force; Initialize-SmartM365DefaultCsvValidationRules
  $lc=LocalConfig
  if(!$DataLastFolder){$DataLastFolder=Cfg $lc 'InputDataLastFolder' (Cfg $lc 'LatestCsvFolderPath' $PSScriptRoot)}
  if(!$OutputFolder){$OutputFolder=Cfg $lc 'ScriptCsvLogFolderPath' (Join-Path (Cfg $lc 'DataAllRootPath' $PSScriptRoot) 'Exchange\Issues\HybridIdentity')}
  if(!$LatestFolder){$LatestFolder=Cfg $lc 'LatestCsvFolderPath' $OutputFolder}
  AssertResolvedPath 'DataLastFolder' $DataLastFolder
  AssertResolvedPath 'OutputFolder' $OutputFolder
  AssertResolvedPath 'LatestFolder' $LatestFolder
  $global:EnableSharePointUpload=CB $lc 'EnableSharePointUpload' $false; if($DisableSharePointUpload){$global:EnableSharePointUpload=$false}
  $global:SharePointSiteHostname=Cfg $lc 'SharePointSiteHostname' ''; $global:SharePointSitePath=Cfg $lc 'SharePointSitePath' ''; $global:SharePointLibraryDisplayName=Cfg $lc 'SharePointLibraryDisplayName' 'Documents'; $global:SharePointTargetFolderPath=Cfg $lc 'SharePointTargetFolderPath' ''; $global:AppId=Cfg $lc 'AppId' ''; $global:TenantId=Cfg $lc 'TenantId' ''; $global:Thumbprint=Cfg $lc 'Thumbprint' (Cfg $lc 'Thumb' '')
  Log "DataLastFolder: $DataLastFolder"; Log "OutputFolder: $OutputFolder"; Log "LatestFolder: $LatestFolder"
  Invoke-SmartM365Preflight -ScriptName $ScriptName -OutputPaths @($OutputFolder,$LatestFolder) | Out-Null
  $ad=Csv 'AD_Users_AllDomains_Enriched.csv'; if($ad.Count -eq 0){$ad=Csv 'AD_Users_AllDomains.csv' -Req}; $dupUpn=Csv 'AD_Users_DuplicateUPN.csv'; $dupSmtp=Csv 'AD_Users_DuplicateSMTP.csv'; $local=Csv 'Exchange_OnPrem_Mailboxes_AllDomains.csv' -Req; $remote=Csv 'Exchange_OnPrem_RemoteMailboxes_AllDomains.csv' -Req; $exo=Csv 'Exchange_EXO_Mailboxes_AllDomains.csv' -Req; $stats=Csv 'Exchange_EXO_Mailboxes_AllDomains_Stats.csv'; $arch=Csv 'Exchange_EXO_Mailboxes_AllDomains_Archive.csv'; $perms=Csv 'Exchange_EXO_Mailboxes_AllDomains_Permissions.csv'; $m365=Csv 'M365_Users_Active.csv' -Req; $lic=Csv 'M365_Licenses_Users.csv'; $plans=Csv 'M365_Licenses_ServicePlans.csv'; $domains=Csv 'Exchange_EXO_AcceptedDomains.csv'; $verified=Csv 'M365_Entra_VerifiedDomains.csv'
  $localGuid=@{};$localUpn=@{};$localSmtp=@{};$localDomainSam=@{}; foreach($r in $local){AddMap $localGuid (P $r ObjectGUID) $r; AddMap $localUpn (P $r UserPrincipalName) $r; AddMap $localSmtp (P $r PrimarySMTPaddress,PrimarySmtpAddress) $r; AddMap $localDomainSam (P $r DomainAndSam) $r}
  $remoteGuid=@{};$remoteUpn=@{};$remoteSmtp=@{}; foreach($r in $remote){AddMap $remoteGuid (P $r ObjectGuid,ObjectGUID) $r; AddMap $remoteUpn (P $r UserPrincipalName) $r; AddMap $remoteSmtp (P $r PrimarySmtpAddress,PrimarySMTPaddress) $r}
  $exoUpn=@{};$exoSmtp=@{};$exoImmutable=@{}; foreach($r in $exo){AddMap $exoUpn (P $r UserPrincipalName) $r; AddMap $exoSmtp (P $r PrimarySmtpAddress,PrimarySMTPaddress) $r; AddMap $exoImmutable (P $r OnPremisesImmutableId,ImmutableId) $r}
  $statsUpn=@{}; foreach($r in $stats){AddMap $statsUpn (P $r UserPrincipalName) $r}; $archUpn=@{}; foreach($r in $arch){AddMap $archUpn (P $r UserPrincipalName) $r}; $permsUpn=@{}; foreach($r in $perms){AddMap $permsUpn (P $r UserPrincipalName) $r}
  $m365Upn=@{};$m365Immutable=@{}; foreach($r in $m365){AddMap $m365Upn (P $r 'User principal name',UserPrincipalName) $r; AddMap $m365Immutable (P $r OnPremisesImmutableId) $r}
  $licUpn=@{};$licUserId=@{}; foreach($r in $lic){$u=K (P $r 'User principal name',UserPrincipalName); $id=K(P $r UserId,'Object Id'); $sku=T (P $r 'SKU name',SkuPartNumber); if(!$sku){continue}; if($u){if(!$licUpn.ContainsKey($u)){$licUpn[$u]=[Collections.Generic.List[string]]::new()}; [void]$licUpn[$u].Add($sku)}; if($id){if(!$licUserId.ContainsKey($id)){$licUserId[$id]=[Collections.Generic.List[string]]::new()}; [void]$licUserId[$id].Add($sku)}}
  $planUserId=@{}; foreach($r in $plans){$id=K(P $r UserId,'Object Id'); if(!$id){continue}; if(!$planUserId.ContainsKey($id)){$planUserId[$id]=[Collections.Generic.List[object]]::new()}; [void]$planUserId[$id].Add($r)}
  $dupUpnSet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); foreach($r in $dupUpn){[void]$dupUpnSet.Add((T(P $r UPN_norm,UserPrincipalName)))}
  $dupSmtpSet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); foreach($r in $dupSmtp){[void]$dupSmtpSet.Add((T(P $r PrimarySmtp,SmtpAddress)))}
  $domSet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); foreach($r in $domains){[void]$domSet.Add((T(P $r DomainName,Name)))}
  $domType=@{}; $domAddressBook=@{}; foreach($r in $domains){$dn=K(P $r DomainName,Name); if(!$dn){continue}; $domType[$dn]=T(P $r DomainType); $domAddressBook[$dn]=B(P $r AddressBookEnabled)}
  $verifiedSet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); foreach($r in $verified){[void]$verifiedSet.Add((T(P $r Id,Name,DomainName)))}
  $samGroups=@{}; $displayGroups=@{}; $targetGroups=@{}; $adUpn=@{}; $adSmtp=@{}; $adImmutableGuid=@{}
  foreach($r in $ad){$sam=K(P $r SamAccountName); if($sam){if(!$samGroups.ContainsKey($sam)){$samGroups[$sam]=0}; $samGroups[$sam]++}; $dn=K(P $r DisplayName); if($dn){if(!$displayGroups.ContainsKey($dn)){$displayGroups[$dn]=0}; $displayGroups[$dn]++}; $ta=K((T(P $r TargetAddress,targetAddress)) -replace '^smtp:',''); if($ta){if(!$targetGroups.ContainsKey($ta)){$targetGroups[$ta]=0}; $targetGroups[$ta]++}; AddMap $adUpn (P $r UserPrincipalName) $r; AddMap $adSmtp (P $r EmailAddress,PrimarySmtp,PrimarySMTPaddress) $r; $imm=K(P $r ImmutableId_AD); if($imm -and !$adImmutableGuid.ContainsKey($imm)){$adImmutableGuid[$imm]=T(P $r ObjectGUID)}}
  $dupSamSet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); foreach($k in $samGroups.Keys){if($samGroups[$k] -gt 1){[void]$dupSamSet.Add($k)}}
  $dupDisplaySet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); foreach($k in $displayGroups.Keys){if($displayGroups[$k] -gt 1){[void]$dupDisplaySet.Add($k)}}
  $dupTargetSet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); foreach($k in $targetGroups.Keys){if($targetGroups[$k] -gt 1){[void]$dupTargetSet.Add($k)}}
  $userMailboxTypes=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); @('UserMailbox','RemoteUserMailbox')|ForEach-Object{[void]$userMailboxTypes.Add($_)}
  $remoteMailboxTypes=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); @('RemoteUserMailbox','RemoteSharedMailbox','RemoteRoomMailbox')|ForEach-Object{[void]$remoteMailboxTypes.Add($_)}
  $targetRecipientTypes=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); @('UserMailbox','SharedMailbox','RoomMailbox','RemoteUserMailbox','RemoteSharedMailbox','RemoteRoomMailbox')|ForEach-Object{[void]$targetRecipientTypes.Add($_)}
  $sharedMailboxTypes=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); @('SharedMailbox','RemoteSharedMailbox')|ForEach-Object{[void]$sharedMailboxTypes.Add($_)}
  $roomMailboxTypes=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); @('RoomMailbox','RemoteRoomMailbox')|ForEach-Object{[void]$roomMailboxTypes.Add($_)}
  $nonPersonalAccountTypes=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); @('Service Account','Admin Account','System Account','Generic Account')|ForEach-Object{[void]$nonPersonalAccountTypes.Add($_)}
  $allowedTargetDomains=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); @('orpea.mail.onmicrosoft.com','orpea.com')|ForEach-Object{[void]$allowedTargetDomains.Add($_)}
  $issues=[Collections.Generic.List[object]]::new()
  foreach($u in $ad){
    $guid=T(P $u ObjectGUID); $upn=T(P $u UserPrincipalName); $uk=K $upn; $immutableId=T(P $u ImmutableId_AD); $domainAndSam=T(P $u DomainAndSam); $enabled=B(P $u Enabled); $primary=T(P $u EmailAddress,PrimarySmtp,PrimarySMTPaddress); if(!$primary){foreach($addr in SplitAddr(P $u ProxyAddresses)){if($addr -match '@'){$primary=$addr; break}}}; $target=T(P $u TargetAddress); $proxyText=T(P $u ProxyAddresses)
    $lm=if($domainAndSam -and $localDomainSam.ContainsKey((K $domainAndSam))){$localDomainSam[(K $domainAndSam)]}elseif($localGuid.ContainsKey((K $guid))){$localGuid[(K $guid)]}elseif($localUpn.ContainsKey($uk)){$localUpn[$uk]}elseif($localSmtp.ContainsKey((K $primary))){$localSmtp[(K $primary)]}else{$null}
    $rm=if($remoteGuid.ContainsKey((K $guid))){$remoteGuid[(K $guid)]}elseif($remoteUpn.ContainsKey($uk)){$remoteUpn[$uk]}elseif($remoteSmtp.ContainsKey((K $primary))){$remoteSmtp[(K $primary)]}else{$null}
    $em=if($exoImmutable.ContainsKey((K $immutableId))){$exoImmutable[(K $immutableId)]}elseif($exoUpn.ContainsKey($uk)){$exoUpn[$uk]}elseif($exoSmtp.ContainsKey((K $primary))){$exoSmtp[(K $primary)]}else{$null}
    $st=if($statsUpn.ContainsKey($uk)){$statsUpn[$uk]}else{$null}; $ar=if($archUpn.ContainsKey($uk)){$archUpn[$uk]}else{$null}; $pm=if($permsUpn.ContainsKey($uk)){$permsUpn[$uk]}else{$null}
    $recipientType=GetRecipientTypeFromMailboxes $primary $lm $rm $em; $isMailbox=($recipientType -ne 'NoMailboxes' -or $lm -or $rm -or $em); $hasSmtp=-not [string]::IsNullOrWhiteSpace($primary)
    $m365User=GetM365User $upn $immutableId; $inM365=$null -ne $m365User; $m365UserId=T(P $m365User 'Object Id',Id,UserId)
    $userLicenses=@(LicensesFor $upn $m365UserId); $licenseText=@($userLicenses) -join ';'; $licenseGroup=GetLicenseGroupFromM365 $recipientType $userLicenses $inM365; $hasAllowedLicense=TestAllowedLicenseGroup $licenseGroup
    $localSizeMb=Dbl(P $lm 'TotalItemSize-In-MB'); $exoSizeMb=[math]::Max((Dbl(P $em TotalItemSizeGB))*1024.0,(Dbl(P $st TotalItemSizeGB))*1024.0); $sizeMb=[math]::Max($localSizeMb,$exoSizeMb); $size=$sizeMb/1024.0
    $arc=[math]::Max((Dbl(P $lm 'ArchiveTotalItemSize-In-MB')/1024.0),(Dbl(P $ar Archive_TotalItemSizeGB)))
    $pc=[int]((Dbl(P $lm FullAccessCount))+(Dbl(P $lm SendAsCount))+(SplitAddr(P $pm FullAccess)).Count+(SplitAddr(P $pm SendAs)).Count)
    $ll=Dt(P $u LastLogonDate); foreach($candidateDate in @((Dt(P $lm LastLogonTime)),(Dt(P $st LastLogonTime,LastUserActionTime)))){if($candidateDate -and (!$ll -or $candidateDate -gt $ll)){$ll=$candidateDate}}
    $itemCount=[math]::Max((Dbl(P $lm ItemCount,TotalItemCount)),(Dbl(P $st ItemCount,TotalItemCount)))
    $baseScope=$isMailbox -and $hasSmtp
    if($baseScope){
      if(-not $inM365){AddIssue $issues 1 $guid 'User not in Azure Entra' '1.Critical' 'Sync user account with Azure Entra'}
      if($itemCount -gt 999999){AddIssue $issues 2 $guid 'Item count > 999,999' '2.High' 'Reduce mailbox item count'}
      if($proxyText -notmatch '@orpea\.mail\.onmicrosoft\.com'){AddIssue $issues 3 $guid 'Missing proxy address' '1.Critical' 'Add required proxy address'}
      if($size -gt 80 -and $recipientType -eq 'UserMailbox'){AddIssue $issues 4 $guid 'Mailbox size > 80 GB' '2.High' 'Archive or clean mailbox'}
      if(($licenseGroup -in @('Microsoft 365 F3','Microsoft 365 F1')) -and $sizeMb -ge 2000){AddIssue $issues 6 $guid 'F3/F1 user with mailbox >= 2 GB' '2.High' 'Review F3/F1 license or reduce mailbox size'}
      if(!$enabled -and $userMailboxTypes.Contains($recipientType)){AddIssue $issues 7 $guid 'Account disabled' '4.Low' 'Review account status'}
      if((!$hasAllowedLicense) -and $enabled -and $userMailboxTypes.Contains($recipientType)){AddIssue $issues 8 $guid 'No M365 license assigned (enabled account)' '1.Critical' 'Assign appropriate M365 license to this enabled account'}
      if($upn -and $dupUpnSet.Contains($upn)){AddIssue $issues 9 $guid 'Duplicate UPN across domains' '1.Critical' 'Resolve UPN conflict before AAD Connect sync'}
      $exchangePlanValue=ExchangePlanStatus $m365UserId; if($exchangePlanValue -eq 'Not Exchange Plan Enabled' -and $hasAllowedLicense -and $userMailboxTypes.Contains($recipientType)){AddIssue $issues 12 $guid 'No Exchange Plan enabled' '1.Critical' 'Enable Exchange Plan only for user migrated'}
      if($arc -gt 80 -and $recipientType -eq 'UserMailbox'){AddIssue $issues 13 $guid 'Archive size > 80 GB' '2.High' 'Review archive mailbox size'}
      if($pc -gt 10){AddIssue $issues 14 $guid 'Too many mailbox permissions' '3.Medium' 'Review mailbox delegation and permissions'}
      if(!$ll -or ((Get-Date).Date-$ll.Date).Days -gt 180){AddIssue $issues 15 $guid 'Mailbox inactive > 6 months' '4.Low' 'Review mailbox usage and consider decommissioning'}
      foreach($a in @((SplitAddr $proxyText)+@($primary))){ if($dupSmtpSet.Contains($a)){AddIssue $issues 16 $guid 'Conflicting proxy address' '1.Critical' 'Resolve proxy address conflicts'; break} }
      $sharedAccessCount=[int][math]::Max((Dbl(P $u SharedAccessCount)), $pc); if($sharedAccessCount -gt 5){AddIssue $issues 17 $guid 'Excessive shared access' '3.Medium' 'Review shared access model'}
      if(!(T(P $u DisplayName))){AddIssue $issues 18 $guid 'Missing DisplayName' '4.Low' 'Complete user metadata'}
      $normTarget=(T $target) -replace '(?i)^smtp:',''; if($remoteMailboxTypes.Contains($recipientType) -and !$normTarget){AddIssue $issues 19 $guid 'Missing TargetAddress for mailbox' '1.Critical' 'Set TargetAddress for mail-enabled user'}
      if($remoteMailboxTypes.Contains($recipientType) -and $normTarget){$targetDomain=DomainPart $normTarget; if($normTarget -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$'){AddIssue $issues 20 $guid 'Invalid TargetAddress format' '2.High' 'Fix TargetAddress format to a valid SMTP address'}; if($targetDomain -and !$allowedTargetDomains.Contains($targetDomain)){AddIssue $issues 21 $guid 'Disallowed TargetAddress domain' '1.Critical' 'Update TargetAddress to an allowed routing domain'}; if(!$proxyText -or -not (ProxyContains $proxyText $normTarget)){AddIssue $issues 22 $guid 'TargetAddress absent from ProxyAddresses' '3.Medium' 'Add TargetAddress as a proxy entry in ProxyAddresses'}; if($dupTargetSet.Contains((K $normTarget))){AddIssue $issues 23 $guid 'Duplicate TargetAddress' '1.Critical' 'Resolve duplicate TargetAddress routing'}}
      $primaryProxyCount=PrimaryProxyCount $proxyText; if($targetRecipientTypes.Contains($recipientType) -and $primaryProxyCount -eq 0){AddIssue $issues 24 $guid 'Missing primary SMTP in proxyAddresses' '1.Critical' 'Add a primary SMTP entry (SMTP:) in proxyAddresses'}
      if($targetRecipientTypes.Contains($recipientType) -and $primaryProxyCount -gt 1){AddIssue $issues 25 $guid 'Multiple primary SMTP in proxyAddresses' '1.Critical' 'Keep only one primary SMTP (SMTP:) in proxyAddresses'}
      if($targetRecipientTypes.Contains($recipientType) -and $primary -and $proxyText -and -not (ProxyContains $proxyText $primary)){AddIssue $issues 26 $guid 'PrimarySMTP not present in proxyAddresses' '2.High' 'Ensure proxyAddresses contains the PrimarySMTPaddress'}
      if($sharedMailboxTypes.Contains($recipientType) -and $inM365 -and $hasAllowedLicense -and $size -lt 50){AddIssue $issues 27 $guid 'Shared mailbox with M365 license (< 50 GB)' '2.High' 'Remove M365 license: shared mailboxes under 50 GB do not require a paid license in EXO'}
      if($roomMailboxTypes.Contains($recipientType) -and $inM365 -and $hasAllowedLicense){AddIssue $issues 28 $guid 'Room mailbox with M365 license' '3.Medium' 'Convert room mailbox to unlicensed or review license assignment'}
      if(!$targetRecipientTypes.Contains($recipientType)){AddIssue $issues 29 $guid 'Unexpected recipient type in scope' '3.Medium' 'Review mailbox type and migration scope for this recipient'}
      $primaryDomain=DomainPart $primary; if($primaryDomain -and $domSet.Count -gt 0 -and !$domSet.Contains($primaryDomain)){AddIssue $issues 30 $guid 'Primary SMTP domain not in EXO accepted domains' '1.Critical' 'Fix Primary SMTP domain or add the domain to EXO Accepted Domains'}
      if($primaryDomain -like '*.onmicrosoft.com' -or $primaryDomain -eq 'onmicrosoft.com'){AddIssue $issues 31 $guid 'Primary SMTP uses onmicrosoft.com domain' '3.Medium' 'Move Primary SMTP from onmicrosoft.com to corporate accepted domain'}
      if($primaryDomain -and $domType.ContainsKey($primaryDomain) -and $domType[$primaryDomain] -eq 'InternalRelay' -and $userMailboxTypes.Contains($recipientType)){AddIssue $issues 32 $guid 'Primary SMTP domain type is InternalRelay' '4.Low' 'Review why an InternalRelay domain is used as Primary SMTP (routing/coexistence)'}
      if($primaryDomain -and $domSet.Count -gt 0 -and !$domType.ContainsKey($primaryDomain)){AddIssue $issues 33 $guid 'Primary SMTP domain type is blank in EXO (not found)' '4.Low' 'Validate domain inventory refresh and EXO accepted domains data source'}
      if($primaryDomain -and $domAddressBook.ContainsKey($primaryDomain) -and -not $domAddressBook[$primaryDomain]){AddIssue $issues 34 $guid 'Domain allowed in EXO but AddressBookEnabled is False' '4.Low' 'Enable AddressBookEnabled on the accepted domain in EXO or move Primary SMTP to an AddressBook-enabled domain'}
      $upnDomain=DomainPart $upn; if($upnDomain -and $verifiedSet.Count -gt 0 -and !$verifiedSet.Contains($upnDomain)){AddIssue $issues 35 $guid 'UPN domain not verified in Azure AD' '1.Critical' 'Add and verify UPN domain in Azure AD or update user UPN'}
      if($upnDomain -and $verifiedSet.Contains($upnDomain) -and $domSet.Count -gt 0 -and !$domSet.Contains($upnDomain)){AddIssue $issues 36 $guid 'UPN domain verified but not EXO accepted' '1.Critical' 'Review UPN domain usage; consider aligning with an EXO accepted domain'}
      if($upnDomain -and $primaryDomain -and $upnDomain -ne $primaryDomain){AddIssue $issues 37 $guid 'UPN domain not aligned with Primary SMTP domain' '4.Low' 'Align UPN domain with Primary SMTP domain'}
      if($lm -and $em -and $recipientType -eq 'UserMailbox'){AddIssue $issues 38 $guid 'Split-brain mailbox (Local UserMailbox + EXO mailbox)' '1.Critical' 'Remediate split-brain: ensure a single authoritative mailbox'}
      if($upn -and (ContainsInvalidUpnChar $upn)){AddIssue $issues 39 $guid 'UPN contains invalid Azure AD characters' '1.Critical' 'Remove or replace invalid characters in UPN before AAD Connect sync'}
      if($primary -and (ContainsInvalidSmtpLocalChar $primary)){AddIssue $issues 40 $guid 'Primary SMTP contains invalid characters' '1.Critical' 'Remove or replace invalid characters in Primary SMTP address'}
      $sam=T(P $u SamAccountName); if($sam -and $dupSamSet.Contains($sam)){AddIssue $issues 41 $guid 'Duplicate SamAccountName across domains' '4.Low' 'Resolve SAM conflict to avoid GPO and profile collisions'}
      if((@($userLicenses) -join ';') -match 'Microsoft 365 E3|ENTERPRISEPACK|SPE_E3' -and (@($userLicenses) -join ';') -match 'Microsoft 365 F1|Microsoft 365 F3|SPE_F1|SPE_F3|M365_F1|M365_F3|O365_F1|O365_F3'){AddIssue $issues 47 $guid 'Conflicting M365 licenses (E3 + F3/F1)' '2.High' 'Review license assignment and keep only one SKU per user'}
      $displayName=K(P $u DisplayName); if($displayName -and $dupDisplaySet.Contains($displayName)){AddIssue $issues 48 $guid 'Duplicate DisplayName across accounts' '4.Low' 'Disambiguate DisplayNames for GAL clarity and migration tooling'}
      if($userMailboxTypes.Contains($recipientType) -and !$ll -and $size -eq 0){AddIssue $issues 50 $guid 'Ghost mailbox (no logon, zero size)' '4.Low' 'Verify mailbox existence and consider decommissioning'}
      $proxyLower=$proxyText.ToLowerInvariant(); $addrLower=$primary.ToLowerInvariant().Trim(); if($targetRecipientTypes.Contains($recipientType) -and $proxyLower -and $addrLower -and ([regex]::Matches($proxyLower,[regex]::Escape($addrLower)).Count -gt 1)){AddIssue $issues 51 $guid 'Duplicate address in ProxyAddresses (case conflict)' '2.High' 'Remove duplicate proxy entries keeping only one instance per address'}
      if($sharedMailboxTypes.Contains($recipientType)){ $noSendAs=((SplitAddr(P $lm SendAs)).Count+(SplitAddr(P $pm SendAs)).Count) -eq 0; $noFullAccess=((SplitAddr(P $lm FullAccess)).Count+(SplitAddr(P $pm FullAccess)).Count) -eq 0; if($noSendAs -and $noFullAccess){AddIssue $issues 53 $guid 'Shared mailbox with no active delegation' '3.Medium' 'Assign at least one FullAccess or SendAs delegate or decommission'}}
      $accountType=T(P $u AccountType); if(!$enabled -and $hasAllowedLicense -and $accountType -ne 'Shared Mailbox' -and $userMailboxTypes.Contains($recipientType)){AddIssue $issues 55 $guid 'Disabled account with active M365 license' '2.High' 'Review license assignment: disabled accounts should not retain paid M365 licenses'}
      if($sharedMailboxTypes.Contains($recipientType) -and (!$hasAllowedLicense) -and $size -ge 50){AddIssue $issues 57 $guid 'Large shared mailbox (>= 50 GB) with no M365 license' '2.High' 'Assign an appropriate M365 license to enable archive and avoid EXO quota issues post-migration'}
      if($accountType -and $nonPersonalAccountTypes.Contains($accountType) -and $hasAllowedLicense){AddIssue $issues 58 $guid 'Non-personal account type with M365 license (E3/F3/F1)' '2.High' 'Review license assignment: service/admin/generic/system accounts should not hold E3 or F3 licenses'}
    }
    if($hasAllowedLicense -and (-not $isMailbox) -and $upn){AddIssue $issues 56 $guid 'M365 license assigned but no mailbox' '3.Medium' 'Review license assignment: user has M365 license but no mailbox is provisioned.'}
  }
  foreach($e in $exo){
    $u=T(P $e UserPrincipalName); $s=T(P $e PrimarySmtpAddress,PrimarySMTPaddress); $imm=K(P $e OnPremisesImmutableId,ImmutableId)
    if($localUpn.ContainsKey((K $u)) -or $remoteUpn.ContainsKey((K $u)) -or $localSmtp.ContainsKey((K $s)) -or $remoteSmtp.ContainsKey((K $s))){continue}
    $adRow=if($imm -and $adImmutableGuid.ContainsKey($imm)){$adImmutableGuid[$imm]}elseif($adUpn.ContainsKey((K $u))){T(P $adUpn[(K $u)] ObjectGUID)}elseif($adSmtp.ContainsKey((K $s))){T(P $adSmtp[(K $s)] ObjectGUID)}else{''}
    if($adRow){AddIssue $issues 54 $adRow 'EXO-only mailbox with AD account (no local or remote mailbox on-premises)' '1.Critical' 'AD account exists but has no on-premises mailbox counterpart (UserMailbox or RemoteMailbox) - verify if migration is complete or if RemoteMailbox conversion was skipped'}
  }
  $files=ExportIssues -rows @($issues); $files += PublishWeeklyHistory -files $files; Log "Generated Exchange hybrid identity issues: $($issues.Count) row(s)"
  if($global:EnableSharePointUpload){foreach($f in $files){try{Invoke-SmartM365SharePointCsvUpload -LocalFilePath $f|Out-Null; Log "SharePoint upload completed: $f"}catch{Warn "SharePoint upload failed for $f : $($_.Exception.Message)"}}}
  Log "$ScriptName completed successfully."
}catch{Write-Error $_; throw}


# SIG # Begin signature block
# MIIH/wYJKoZIhvcNAQcCoIIH8DCCB+wCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBTQR8FTlerJD+N
# p8msufmnQIBrg+IjhZ3IdEqCSiDZx6CCBMEwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjGCApQw
# ggKQAgEBMGIwTjEeMBwGA1UEAwwVd29ya3BsYWNlY2xvdWRodWIuY29tMSwwKgYJ
# KoZIhvcNAQkBFh1jb250YWN0QHdvcmtwbGFjZWNsb3VkaHViLmNvbQIQHm7vO8c4
# 4bNEOMjxAx/iaDANBglghkgBZQMEAgEFAKCBhDAYBgorBgEEAYI3AgEMMQowCKAC
# gAChAoAAMBkGCSqGSIb3DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
# DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCCiOXocZE8ZCv+SE5NrXKjo
# ZWW32QVuvXlaj+w/AqUQ7DANBgkqhkiG9w0BAQEFAASCAYB7XbNOafxdqCOcPAAo
# pzZIv5CUDPD+0+jgpWptE4SgxN6Xnwe5QKuAmCindiFaZ6Ef6hYf7JmggN9LM8NH
# eoDrEqArvYFy23F8LDLjG3PWV6AmQvAIz21T2yJIaZ2q3R2qUiZ0bDb24P3In2gI
# FaDh7dKdPjTDoDNHqv6QH7eJNIcozdeUFXwWL2IQhbzHpwkR+/4/+fUb/zlqFa7c
# 5i6VRQ+dmUuJDWvhqfeIahtRWDDxkb0pRhkh6FdrP7cR/Szo/ZzLwatKgjATiJjU
# Ano2LvPODjU4XK40srBz9ofBQ0KhVkugRncCnDDvnYtr9ffca/S+EAvkFg9qiEo2
# J0BmlB85jHdvdekW8yaruZAWBHugo+oBDPTsvyOjDFxuVi7Mbgcgs8pXmGdOYZO2
# dGeuyFCccq8uHpQ+VDJF1XPlSSjMUDbKBns3hoRmOzZRyJln3JVmb4ZeEBtjllpc
# hCJ8gkq/tv0Wc7hBU3lXJXfggWRvE/+zj5lD4jHnqG0hI48=
# SIG # End signature block
