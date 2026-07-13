<#
.SYNOPSIS
Generates Exchange hybrid identity issue tables for PowerBI from SmartInventory CSV exports.

.REQUIREMENTS
- PowerShell 7+
- SmartM365.Core module
- Read access to DATA-LAST SmartInventory CSV exports
- Write access to the configured DATA-ALL, DATA-LAST, and WeeklyHistory output folders
- Required input CSV files: AD_Users_AllDomains.csv, Exchange_OnPrem_Mailboxes_AllDomains.csv,
  Exchange_OnPrem_RemoteMailboxes_AllDomains.csv, Exchange_EXO_Mailboxes_AllDomains.csv,
  M365_Users_Active.csv

.VERSION
1.5
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
$ScriptVersion="1.7"
$RunStamp=Get-Date -Format 'yyyyMMdd-HHmmss'
function Log([string]$m){ Write-Host ("{0} [INFO] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$m) }
function Warn([string]$m){ Write-Warning $m }
function Root(){ $d=$PSScriptRoot; while($d){ if((Test-Path (Join-Path $d 'Config\SmartM365-TenantContext.ps1')) -and (Test-Path (Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'))){return $d}; $p=Split-Path $d -Parent; if(!$p -or $p -eq $d){break}; $d=$p }; throw 'SmartM365 root not found.' }
function LocalConfig(){ $p=Join-Path $PSScriptRoot (([IO.Path]::GetFileNameWithoutExtension($PSCommandPath))+'.local.json'); if(!(Test-Path $p)){ $t="$p.template"; if(!(Test-Path $t)){throw "Missing config template: $t"}; Copy-Item $t $p; Log "Created local config: $p" }; Get-Content $p -Raw | ConvertFrom-Json }
function ResolveToken($v){ if($v -isnot [string] -or [string]::IsNullOrWhiteSpace($v)){return $v}; $r=$v; 0..9|%{ $ms=[regex]::Matches($r,'\{\{(?<n>[A-Za-z0-9_.-]+)\}\}'); if($ms.Count -eq 0){return $r}; foreach($m in $ms){ $pr=$script:Cfg.PSObject.Properties[$m.Groups['n'].Value]; if($pr){$r=$r.Replace($m.Value,[string](ResolveToken $pr.Value))} } }; $r }
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
function ContainsInvalidUpnChar($value){(T $value) -match '[\\/\[\]:;\|=,+*?<>]'}
function ContainsInvalidSmtpLocalChar($value){(LocalPart $value) -match '[\\/\[\]:;\|=,+*?<>\s]'}
function PrimaryProxyCount($proxyText){ return @((T $proxyText) -split '[;|,]' | Where-Object { $_ -cmatch '^SMTP:' }).Count }
function ProxyContains($proxyText,$address){$a=(T $address) -replace '^(smtp|SMTP):',''; foreach($p in SplitAddr $proxyText){if($p -ieq $a){return $true}}; $false}
function IsFalseLike($value){$s=K $value; $s -in @('false','0','no','n','disabled')}
function LicensesFor($upn){$k=K $upn; if($licUpn.ContainsKey($k)){return @($licUpn[$k])}; @()}
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
  Assert-SmartM365CsvDataCompleteness -Data $rows -TimestampedPath $main -LatestPath $latest; $rows|Sort-Object IssueNumber,ObjectGUID,Potential_Issue|Export-Csv $main -NoTypeInformation -Encoding UTF8; CopyCsv $main $latest; CopyCsv $main $archive
  $summary=@($rows|Group-Object IssueNumber,Potential_Issue,IssueCategory|Sort-Object Count -Descending|%{ $p=$_.Name -split ', ',3; [pscustomobject]@{IssueNumber=[int]$p[0];Potential_Issue=$p[1];IssueCategory=$p[2];Count=$_.Count} })
  $sm=Join-Path $OutputFolder 'Exchange_HybridIdentity_Issues_Summary.csv'; $sl=Join-Path $LatestFolder 'Exchange_HybridIdentity_Issues_Summary.csv'; Assert-SmartM365CsvDataCompleteness -Data $summary -TimestampedPath $sm -LatestPath $sl; $summary|Export-Csv $sm -NoTypeInformation -Encoding UTF8; CopyCsv $sm $sl
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
  foreach($f in @($files)){
    if(!$f -or !(Test-Path -LiteralPath $f)){continue}
    $dest=Join-Path $folder (Split-Path $f -Leaf)
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
  $global:EnableSharePointUpload=CB $lc 'EnableSharePointUpload' $false; if($DisableSharePointUpload){$global:EnableSharePointUpload=$false}
  $global:SharePointSiteHostname=Cfg $lc 'SharePointSiteHostname' ''; $global:SharePointSitePath=Cfg $lc 'SharePointSitePath' ''; $global:SharePointLibraryDisplayName=Cfg $lc 'SharePointLibraryDisplayName' 'Documents'; $global:SharePointTargetFolderPath=Cfg $lc 'SharePointTargetFolderPath' ''; $global:AppId=Cfg $lc 'AppId' ''; $global:TenantId=Cfg $lc 'TenantId' ''; $global:Thumbprint=Cfg $lc 'Thumbprint' (Cfg $lc 'Thumb' '')
  Log "DataLastFolder: $DataLastFolder"; Log "OutputFolder: $OutputFolder"; Log "LatestFolder: $LatestFolder"
  Invoke-SmartM365Preflight -ScriptName $ScriptName -OutputPaths @($OutputFolder,$LatestFolder) | Out-Null
  $ad=Csv 'AD_Users_AllDomains.csv' -Req; $dupUpn=Csv 'AD_Users_DuplicateUPN.csv'; $dupSmtp=Csv 'AD_Users_DuplicateSMTP.csv'; $local=Csv 'Exchange_OnPrem_Mailboxes_AllDomains.csv' -Req; $remote=Csv 'Exchange_OnPrem_RemoteMailboxes_AllDomains.csv' -Req; $exo=Csv 'Exchange_EXO_Mailboxes_AllDomains.csv' -Req; $stats=Csv 'Exchange_EXO_Mailboxes_AllDomains_Stats.csv'; $arch=Csv 'Exchange_EXO_Mailboxes_AllDomains_Archive.csv'; $perms=Csv 'Exchange_EXO_Mailboxes_AllDomains_Permissions.csv'; $m365=Csv 'M365_Users_Active.csv' -Req; $lic=Csv 'M365_Licenses_Users.csv'; $plans=Csv 'M365_Licenses_ServicePlans.csv'; $domains=Csv 'Exchange_EXO_AcceptedDomains.csv'; $verified=Csv 'M365_Entra_VerifiedDomains.csv'
  $localGuid=@{};$localUpn=@{};$localSmtp=@{}; foreach($r in $local){AddMap $localGuid (P $r ObjectGUID) $r; AddMap $localUpn (P $r UserPrincipalName) $r; AddMap $localSmtp (P $r PrimarySMTPaddress,PrimarySmtpAddress) $r}
  $remoteGuid=@{};$remoteUpn=@{};$remoteSmtp=@{}; foreach($r in $remote){AddMap $remoteGuid (P $r ObjectGuid,ObjectGUID) $r; AddMap $remoteUpn (P $r UserPrincipalName) $r; AddMap $remoteSmtp (P $r PrimarySmtpAddress,PrimarySMTPaddress) $r}
  $exoUpn=@{};$exoSmtp=@{}; foreach($r in $exo){AddMap $exoUpn (P $r UserPrincipalName) $r; AddMap $exoSmtp (P $r PrimarySmtpAddress,PrimarySMTPaddress) $r}
  $statsUpn=@{}; foreach($r in $stats){AddMap $statsUpn (P $r UserPrincipalName) $r}; $archUpn=@{}; foreach($r in $arch){AddMap $archUpn (P $r UserPrincipalName) $r}; $permsUpn=@{}; foreach($r in $perms){AddMap $permsUpn (P $r UserPrincipalName) $r}; $m365Upn=@{}; foreach($r in $m365){AddMap $m365Upn (P $r 'User principal name',UserPrincipalName) $r}
  $licUpn=@{}; foreach($r in $lic){$u=K (P $r 'User principal name',UserPrincipalName); if(!$u){continue}; if(!$licUpn.ContainsKey($u)){$licUpn[$u]=[Collections.Generic.List[string]]::new()}; $sku=T (P $r SkuPartNumber,'SKU name'); if($sku){[void]$licUpn[$u].Add($sku)}}
  $dupUpnSet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); foreach($r in $dupUpn){[void]$dupUpnSet.Add((T(P $r UserPrincipalName)))}
  $dupSmtpSet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); foreach($r in $dupSmtp){[void]$dupSmtpSet.Add((T(P $r SmtpAddress)))}
  $domSet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); foreach($r in $domains){[void]$domSet.Add((T(P $r DomainName,Name)))}
  $domType=@{}; $domAddressBook=@{}; foreach($r in $domains){$dn=K(P $r DomainName,Name); if(!$dn){continue}; $domType[$dn]=T(P $r DomainType); $domAddressBook[$dn]=B(P $r AddressBookEnabled)}
  $verifiedSet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); foreach($r in $verified){[void]$verifiedSet.Add((T(P $r Id,Name,DomainName)))}
  $samGroups=@{}; $displayGroups=@{}; $targetGroups=@{}
  foreach($r in $ad){$sam=K(P $r SamAccountName); if($sam){if(!$samGroups.ContainsKey($sam)){$samGroups[$sam]=0}; $samGroups[$sam]++}; $dn=K(P $r DisplayName); if($dn){if(!$displayGroups.ContainsKey($dn)){$displayGroups[$dn]=0}; $displayGroups[$dn]++}; $ta=K((T(P $r TargetAddress,targetAddress)) -replace '^smtp:',''); if($ta){if(!$targetGroups.ContainsKey($ta)){$targetGroups[$ta]=0}; $targetGroups[$ta]++}}
  $dupSamSet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); foreach($k in $samGroups.Keys){if($samGroups[$k] -gt 1){[void]$dupSamSet.Add($k)}}
  $dupDisplaySet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); foreach($k in $displayGroups.Keys){if($displayGroups[$k] -gt 1){[void]$dupDisplaySet.Add($k)}}
  $dupTargetSet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); foreach($k in $targetGroups.Keys){if($targetGroups[$k] -gt 1){[void]$dupTargetSet.Add($k)}}
  $issues=[Collections.Generic.List[object]]::new()
  foreach($u in $ad){
    $guid=T(P $u ObjectGUID); $upn=T(P $u UserPrincipalName); $uk=K $upn; $enabled=B(P $u Enabled); $primary=T(P $u EmailAddress,PrimarySmtp,PrimarySMTPaddress); if(!$primary){$primary=SplitAddr(P $u ProxyAddresses)|?{$_ -match '@'}|select -First 1}; $target=T(P $u TargetAddress)
    $lm=if($localGuid.ContainsKey((K $guid))){$localGuid[(K $guid)]}elseif($localUpn.ContainsKey($uk)){$localUpn[$uk]}else{$null}; $rm=if($remoteGuid.ContainsKey((K $guid))){$remoteGuid[(K $guid)]}elseif($remoteUpn.ContainsKey($uk)){$remoteUpn[$uk]}else{$null}; $em=if($exoUpn.ContainsKey($uk)){$exoUpn[$uk]}elseif($exoSmtp.ContainsKey((K $primary))){$exoSmtp[(K $primary)]}else{$null}
    $st=if($statsUpn.ContainsKey($uk)){$statsUpn[$uk]}else{$null}; $ar=if($archUpn.ContainsKey($uk)){$archUpn[$uk]}else{$null}; $pm=if($permsUpn.ContainsKey($uk)){$permsUpn[$uk]}else{$null}; $hasMbx=($lm -or $rm -or $em); $hasLic=$licUpn.ContainsKey($uk)
    if($hasMbx -and !$m365Upn.ContainsKey($uk)){AddIssue $issues 1 $guid 'User not found in M365 active users' '1.Critical' 'Review Entra ID synchronization and cloud account state.'}
    if($hasMbx -and !(T(P $u ProxyAddresses))){AddIssue $issues 3 $guid 'Missing proxy address' '1.Critical' 'Add required proxy addresses before coexistence, migration, or daily mail routing decisions.'}
    $size=[math]::Max((Dbl(P $lm 'TotalItemSize-In-MB')/1024.0),(Dbl(P $st TotalItemSizeGB))); if($size -gt 80){AddIssue $issues 4 $guid 'Mailbox size > 80 GB' '2.High' 'Review archive, retention, cleanup, or target license/quota before migration or operations.'}
    if($hasMbx -and !$enabled){AddIssue $issues 7 $guid 'Account disabled with mailbox' '4.Low' 'Review whether the mailbox should remain active, be converted, archived, or decommissioned.'}
    if($hasMbx -and $enabled -and !$hasLic){AddIssue $issues 8 $guid 'No M365 license assigned (enabled mailbox)' '1.Critical' 'Assign or validate the expected license for this enabled mailbox.'}
    if($dupUpnSet.Contains($upn)){AddIssue $issues 9 $guid 'Duplicate UPN across domains' '1.Critical' 'Resolve UPN duplicates before synchronization, migration, or identity cleanup.'}
    $arc=[math]::Max((Dbl(P $lm 'ArchiveTotalItemSize-In-MB')/1024.0),(Dbl(P $ar Archive_TotalItemSizeGB))); if($arc -gt 80){AddIssue $issues 13 $guid 'Archive size > 80 GB' '2.High' 'Review archive size and quota before migration or operational quota decisions.'}
    $pc=(SplitAddr(P $lm FullAccess)).Count+(SplitAddr(P $lm SendAs)).Count+(SplitAddr(P $pm FullAccess)).Count+(SplitAddr(P $pm SendAs)).Count; if($pc -gt 25){AddIssue $issues 14 $guid 'Too many mailbox permissions' '3.Medium' 'Review mailbox delegation and remove obsolete permissions.'}
    $ll=Dt(P $lm LastLogonTime); if(!$ll){$ll=Dt(P $st LastLogonTime,LastUserActionTime)}; if($hasMbx -and $ll -and ((Get-Date).Date-$ll.Date).Days -gt 180){AddIssue $issues 15 $guid 'Mailbox inactive > 6 months' '4.Low' 'Review mailbox usage and consider cleanup, archive, or decommissioning.'}
    foreach($a in @((SplitAddr(P $u ProxyAddresses))+@($primary))){ if($dupSmtpSet.Contains($a)){AddIssue $issues 16 $guid 'Duplicate SMTP address' '1.Critical' 'Resolve duplicate SMTP/proxy addresses before synchronization or mail routing changes.'; break} }
    if($rm -and !(T(P $rm RemoteRoutingAddress))){AddIssue $issues 19 $guid 'Remote mailbox missing remote routing address' '1.Critical' 'Repair RemoteRoutingAddress/targetAddress for this remote mailbox.'}
    $itemCount=[math]::Max((Dbl(P $lm ItemCount,TotalItemCount)),(Dbl(P $st ItemCount,TotalItemCount)))
    if((B(P $u IsItemCountOver999999)) -or $itemCount -gt 999999){AddIssue $issues 2 $guid 'Item count > 999,999' '2.High' 'Reduce mailbox item count'}
    $userLicenses=@(LicensesFor $upn)
    if($hasMbx -and $size -ge 2 -and (HasSkuLike $userLicenses '(^|_)F(1|3)($|_)|SPE_F1|SPE_F3|M365_F1|M365_F3|O365_F1|O365_F3')){AddIssue $issues 6 $guid 'F3/F1 user with mailbox >= 2 GB' '2.High' 'Review F3/F1 license or reduce mailbox size'}
    $exchangePlanValue=T(P $u LicenseExchangePlanEnabledFromM365,ExchangePlanEnabled,HasExchangePlan)
    if($hasLic -and $hasMbx -and $exchangePlanValue -and (IsFalseLike $exchangePlanValue)){AddIssue $issues 12 $guid 'No Exchange Plan enabled' '1.Critical' 'Enable Exchange Plan only for user migrated'}
    $sharedAccessCount=[int][math]::Max((Dbl(P $u SharedAccessCount)), $pc)
    if($sharedAccessCount -gt 25){AddIssue $issues 17 $guid 'Excessive shared access' '3.Medium' 'Review shared access model'}
    if($hasMbx -and !(T(P $u DisplayName))){AddIssue $issues 18 $guid 'Missing DisplayName' '4.Low' 'Complete user metadata'}
    $normTarget=(T $target) -replace '^smtp:',''
    if($normTarget){
      if($normTarget -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$'){AddIssue $issues 20 $guid 'Invalid TargetAddress format' '2.High' 'Fix TargetAddress format to a valid SMTP address'}
      $targetDomain=DomainPart $normTarget
      $allowedDomains=@('mail.onmicrosoft.com','onmicrosoft.com')
      if($targetDomain -and $allowedDomains.Count -gt 0 -and -not ($allowedDomains | Where-Object { $targetDomain -like "*$_" })){AddIssue $issues 21 $guid 'Disallowed TargetAddress domain' '1.Critical' 'Update TargetAddress to an allowed routing domain'}
      if((T(P $u ProxyAddresses)) -and -not (ProxyContains (P $u ProxyAddresses) $normTarget)){AddIssue $issues 22 $guid 'TargetAddress absent from ProxyAddresses' '3.Medium' 'Add TargetAddress as a proxy entry in ProxyAddresses'}
      if($dupTargetSet.Contains((K $normTarget))){AddIssue $issues 23 $guid 'Duplicate TargetAddress' '1.Critical' 'Resolve duplicate TargetAddress routing'}
    }
    $proxyText=T(P $u ProxyAddresses)
    $primaryProxyCount=PrimaryProxyCount $proxyText
    if($hasMbx -and $proxyText -and $primaryProxyCount -eq 0){AddIssue $issues 24 $guid 'Missing primary SMTP in proxyAddresses' '1.Critical' 'Add a primary SMTP entry (SMTP:) in proxyAddresses'}
    if($primaryProxyCount -gt 1){AddIssue $issues 25 $guid 'Multiple primary SMTP in proxyAddresses' '1.Critical' 'Keep only one primary SMTP (SMTP:) in proxyAddresses'}
    if($primary -and $proxyText -and -not (ProxyContains $proxyText $primary)){AddIssue $issues 26 $guid 'PrimarySMTP not present in proxyAddresses' '2.High' 'Ensure proxyAddresses contains the PrimarySMTPaddress'}
    if((IsSharedMailbox $lm $rm $em) -and $size -lt 50 -and $hasLic){AddIssue $issues 27 $guid 'Shared mailbox with M365 license (< 50 GB)' '2.High' 'Remove M365 license: shared mailboxes under 50 GB do not require a paid license in EXO'}
    if((IsRoomMailbox $lm $rm $em) -and $hasLic){AddIssue $issues 28 $guid 'Room mailbox with M365 license' '3.Medium' 'Convert room mailbox to unlicensed or review license assignment'}
    if($hasMbx -and -not (IsUserMailbox $lm $rm $em) -and -not (IsSharedMailbox $lm $rm $em) -and -not (IsRoomMailbox $lm $rm $em)){AddIssue $issues 29 $guid 'Unexpected recipient type in scope' '3.Medium' 'Review mailbox type and migration scope for this recipient'}
    $primaryDomain=DomainPart $primary
    if($primaryDomain -and $domSet.Count -gt 0 -and !$domSet.Contains($primaryDomain)){AddIssue $issues 30 $guid 'Primary SMTP domain not in EXO accepted domains' '1.Critical' 'Fix Primary SMTP domain or add the domain to EXO Accepted Domains'}
    if($primaryDomain -like '*.onmicrosoft.com' -or $primaryDomain -eq 'onmicrosoft.com'){AddIssue $issues 31 $guid 'Primary SMTP uses onmicrosoft.com domain' '3.Medium' 'Move Primary SMTP from onmicrosoft.com to corporate accepted domain'}
    if($primaryDomain -and $domType.ContainsKey($primaryDomain) -and $domType[$primaryDomain] -eq 'InternalRelay'){AddIssue $issues 32 $guid 'Primary SMTP domain type is InternalRelay' '4.Low' 'Review why an InternalRelay domain is used as Primary SMTP (routing/coexistence)'}
    if($primaryDomain -and $domSet.Count -gt 0 -and !$domType.ContainsKey($primaryDomain)){AddIssue $issues 33 $guid 'Primary SMTP domain type is blank in EXO (not found)' '4.Low' 'Validate domain inventory refresh and EXO accepted domains data source'}
    if($primaryDomain -and $domAddressBook.ContainsKey($primaryDomain) -and -not $domAddressBook[$primaryDomain]){AddIssue $issues 34 $guid 'Domain allowed in EXO but AddressBookEnabled is False' '4.Low' 'Enable AddressBookEnabled on the accepted domain in EXO or move Primary SMTP to an AddressBook-enabled domain'}
    $upnDomain=DomainPart $upn
    if($upnDomain -and $verifiedSet.Count -gt 0 -and !$verifiedSet.Contains($upnDomain)){AddIssue $issues 35 $guid 'UPN domain not verified in Azure AD' '1.Critical' 'Add and verify UPN domain in Azure AD or update user UPN'}
    if($upnDomain -and $verifiedSet.Contains($upnDomain) -and $domSet.Count -gt 0 -and !$domSet.Contains($upnDomain)){AddIssue $issues 36 $guid 'UPN domain verified but not EXO accepted' '1.Critical' 'Review UPN domain usage; consider aligning with an EXO accepted domain'}
    if($upnDomain -and $primaryDomain -and $upnDomain -ne $primaryDomain){AddIssue $issues 37 $guid 'UPN domain not aligned with Primary SMTP domain' '4.Low' 'Align UPN domain with Primary SMTP domain'}
    if($lm -and $em -and (IsUserMailbox $lm $null $em)){AddIssue $issues 38 $guid 'Split-brain mailbox (Local UserMailbox + EXO mailbox)' '1.Critical' 'Remediate split-brain: ensure a single authoritative mailbox'}
    if($upn -and (ContainsInvalidUpnChar $upn)){AddIssue $issues 39 $guid 'UPN contains invalid Azure AD characters' '1.Critical' 'Remove or replace invalid characters in UPN before AAD Connect sync'}
    if($primary -and (ContainsInvalidSmtpLocalChar $primary)){AddIssue $issues 40 $guid 'Primary SMTP contains invalid characters' '1.Critical' 'Remove or replace invalid characters in Primary SMTP address'}
    if($sam -and $dupSamSet.Contains($sam)){AddIssue $issues 41 $guid 'Duplicate SamAccountName across domains' '4.Low' 'Resolve SAM conflict to avoid GPO and profile collisions'}
    if((HasSkuLike $userLicenses 'E3|ENTERPRISEPACK|SPE_E3') -and (HasSkuLike $userLicenses 'F1|F3|SPE_F1|SPE_F3|M365_F1|M365_F3')){AddIssue $issues 47 $guid 'Conflicting M365 licenses (E3 + F3/F1)' '2.High' 'Review license assignment and keep only one SKU per user'}
    $displayName=K(P $u DisplayName); if($displayName -and $dupDisplaySet.Contains($displayName)){AddIssue $issues 48 $guid 'Duplicate DisplayName across accounts' '4.Low' 'Disambiguate DisplayNames for GAL clarity and migration tooling'}
    if($hasMbx -and (!$ll) -and $size -eq 0 -and $itemCount -eq 0){AddIssue $issues 50 $guid 'Ghost mailbox (no logon, zero size)' '4.Low' 'Verify mailbox existence and consider decommissioning'}
    $seenProxy=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); foreach($pa in SplitAddr $proxyText){if($seenProxy.Contains($pa)){AddIssue $issues 51 $guid 'Duplicate address in ProxyAddresses (case conflict)' '2.High' 'Remove duplicate proxy entries keeping only one instance per address'; break}; [void]$seenProxy.Add($pa)}
    if((IsSharedMailbox $lm $rm $em) -and $pc -eq 0){AddIssue $issues 53 $guid 'Shared mailbox with no active delegation' '3.Medium' 'Assign at least one FullAccess or SendAs delegate or decommission'}
    if(!$enabled -and $hasLic){AddIssue $issues 55 $guid 'Disabled account with active M365 license' '2.High' 'Review license assignment: disabled accounts should not keep paid licenses unless justified.'}
    if($hasLic -and !$hasMbx){AddIssue $issues 56 $guid 'M365 license assigned but no mailbox' '3.Medium' 'Review license assignment: user has M365 license but no mailbox is provisioned.'}
    $rtype=((T(P $lm RecipientType,RecipientTypeDetails))+' '+(T(P $rm RecipientTypeDetails,RecipientType))+' '+(T(P $em RecipientTypeDetails,MailboxType))).ToLowerInvariant(); if($rtype -match 'shared' -and $size -ge 50 -and !$hasLic){AddIssue $issues 57 $guid 'Large shared mailbox (>= 50 GB) with no M365 license' '2.High' 'Review shared mailbox license/quota/archive requirements.'}
    $sam=T(P $u SamAccountName); if($hasLic -and ($sam -match '^(adm|admin|svc|service|sa_|sys|test)' -or $upn -match 'admin|service|svc|test')){AddIssue $issues 58 $guid 'Non-personal account with M365 license' '2.High' 'Review license assignment on service/admin/generic accounts.'}
    foreach($d in @(($upn -split '@')[-1],($primary -split '@')[-1],(($target -replace '^smtp:','') -split '@')[-1])){ if($d -and $domSet.Count -gt 0 -and !$domSet.Contains($d)){AddIssue $issues 60 $guid 'Mail domain not found in EXO accepted domains' '2.High' 'Validate UPN, primary SMTP, targetAddress, and accepted domain configuration.'; break} }
  }
  foreach($e in $exo){$u=T(P $e UserPrincipalName); $s=T(P $e PrimarySmtpAddress,PrimarySMTPaddress); if($localUpn.ContainsKey((K $u)) -or $remoteUpn.ContainsKey((K $u)) -or $localSmtp.ContainsKey((K $s)) -or $remoteSmtp.ContainsKey((K $s))){continue}; AddIssue $issues 54 (T(P $e OnPremisesImmutableId,ImmutableId,ExternalDirectoryObjectId)) 'EXO-only mailbox with no on-premises mailbox counterpart' '1.Critical' 'Verify whether migration is complete or RemoteMailbox/on-premises representation is missing.'}
  $files=ExportIssues -rows @($issues); $files += PublishWeeklyHistory -files $files; Log "Generated Exchange hybrid identity issues: $($issues.Count) row(s)"
  if($global:EnableSharePointUpload){foreach($f in $files){try{Invoke-SmartM365SharePointCsvUpload -LocalFilePath $f|Out-Null; Log "SharePoint upload completed: $f"}catch{Warn "SharePoint upload failed for $f : $($_.Exception.Message)"}}}
  Log "$ScriptName completed successfully."
}catch{Write-Error $_; throw}


# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCV1KnG3taLgkhm
# or7ubcZcXcHuHJHiUwkWkndmGygmcKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCAI6mB6mbmjV6eHFNyb86XCmqF+XN/dZN1i1dtcgO7DojANBgkqhkiG9w0B
# AQEFAASCAYBJoESsrTsWCIvfi3UHMlgdR+pkJC2tCcpBF4M+o4fzscRcKd9zqklX
# fIzOm0G95GGZ8uyitfNJHoCzJV03nOqqIzj1z3fN2iPnwqKvKGLLB/ldTQuSRr8u
# 3YU+MmHsQBtd/QHb8egyL5ZNJJujZxUF1Jkps0Fv7cq7mviq15XvhuWkYJG4bdvx
# sF5SiefJ8DF8WFVh4lY9dAlgKEUsPiy14wW5Y4M3ZCn7NgjpL0VAeTExdZGvKL9X
# xuATE/6GJnaizAYsuhw3zD7QfFjMiUBdgdmO5x5sTqxwLd3EI9U5p71UDBOC/tHo
# 4+xowy/TEBxP43DAdR/+l8tKZ1bRBJbGEhk+I1Wxfwr0ghbdZ2zYm2gzU6blbW1i
# ziX/ckBEmUeK7fn7ZFWahVBaFXZVHfq3O33QScM5n7yXqrNrv60mdeCt0hzQKZFp
# fBS3LByOrbmg91v2RvoMXgGbIeBm3oCNiH8GAa9dEzntcJOpiudYTX0tGqHG3JTS
# W94Tr+ma7yg=
# SIG # End signature block
