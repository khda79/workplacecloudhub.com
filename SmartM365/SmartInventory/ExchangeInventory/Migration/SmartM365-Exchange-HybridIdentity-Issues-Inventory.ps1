<#
.SYNOPSIS
Generates Exchange hybrid identity issue tables for PowerBI from SmartInventory CSV exports.
.VERSION
1.0
#>
#requires -Version 7.0
[CmdletBinding()]
param(
  [string]$Tenant='test',
  [string]$DataLastFolder='',
  [string]$OutputFolder='',
  [string]$LatestFolder='',
  [switch]$DisableSharePointUpload,
  [switch]$SkipLegacyAliases
)
$ErrorActionPreference='Stop'
$ScriptName='SmartM365-Exchange-HybridIdentity-Issues-Inventory'
$ScriptVersion='1.0'
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
function AddIssue($list,[int]$num,[string]$guid,[string]$issue,[string]$cat,[string]$action){ if(!$guid){$guid='UNRESOLVED'}; [void]$list.Add([pscustomobject]@{IssueNumber=$num;ObjectGUID=$guid;Potential_Issue=$issue;IssueCategory=$cat;RecommendedAction=$action}) }
function SplitAddr($v){ $s=T $v; if(!$s){@()}else{@($s -split '[;|,]'|%{(T $_)-replace '^(smtp|SMTP):',''}|?{$_})} }
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
  $rows|Sort-Object IssueNumber,ObjectGUID,Potential_Issue|Export-Csv $main -NoTypeInformation -Encoding UTF8; CopyCsv $main $latest; CopyCsv $main $archive
  $summary=@($rows|Group-Object IssueNumber,Potential_Issue,IssueCategory|Sort-Object Count -Descending|%{ $p=$_.Name -split ', ',3; [pscustomobject]@{IssueNumber=[int]$p[0];Potential_Issue=$p[1];IssueCategory=$p[2];Count=$_.Count} })
  $sm=Join-Path $OutputFolder 'Exchange_HybridIdentity_Issues_Summary.csv'; $sl=Join-Path $LatestFolder 'Exchange_HybridIdentity_Issues_Summary.csv'; $summary|Export-Csv $sm -NoTypeInformation -Encoding UTF8; CopyCsv $sm $sl
  $files=@($main,$latest,$archive,$sm,$sl)
  if(!$SkipLegacyAliases){ $lm=Join-Path $OutputFolder 'Mig_Potential_Issues_Expanded.csv'; $ll=Join-Path $LatestFolder 'Mig_Potential_Issues_Expanded.csv'; CopyCsv $main $lm; CopyCsv $main $ll; $files+=@($lm,$ll) }
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
  $sr=Root; . (Join-Path $sr 'Config\SmartM365-TenantContext.ps1'); $script:Cfg=Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot; Import-Module (Join-Path $sr 'Modules\SmartM365.Core\SmartM365.Core.psd1') -Force
  $lc=LocalConfig
  if(!$DataLastFolder){$DataLastFolder=Cfg $lc 'InputDataLastFolder' (Cfg $lc 'LatestCsvFolderPath' $PSScriptRoot)}
  if(!$OutputFolder){$OutputFolder=Cfg $lc 'ScriptCsvLogFolderPath' (Join-Path (Cfg $lc 'DataAllRootPath' $PSScriptRoot) 'Exchange\Issues\HybridIdentity')}
  if(!$LatestFolder){$LatestFolder=Cfg $lc 'LatestCsvFolderPath' $OutputFolder}
  $global:EnableSharePointUpload=CB $lc 'EnableSharePointUpload' $false; if($DisableSharePointUpload){$global:EnableSharePointUpload=$false}; if(!(CB $lc 'EmitLegacyPowerBIAlias' $true)){$SkipLegacyAliases=$true}
  $global:SharePointSiteHostname=Cfg $lc 'SharePointSiteHostname' ''; $global:SharePointSitePath=Cfg $lc 'SharePointSitePath' ''; $global:SharePointLibraryDisplayName=Cfg $lc 'SharePointLibraryDisplayName' 'Documents'; $global:SharePointTargetFolderPath=Cfg $lc 'SharePointTargetFolderPath' ''; $global:AppId=Cfg $lc 'AppId' ''; $global:TenantId=Cfg $lc 'TenantId' ''; $global:Thumbprint=Cfg $lc 'Thumbprint' (Cfg $lc 'Thumb' '')
  Log "DataLastFolder: $DataLastFolder"; Log "OutputFolder: $OutputFolder"; Log "LatestFolder: $LatestFolder"
  $ad=Csv 'AD_Users_AllDomains.csv' -Req; $dupUpn=Csv 'AD_Users_DuplicateUPN.csv'; $dupSmtp=Csv 'AD_Users_DuplicateSMTP.csv'; $local=Csv 'Exchange_OnPrem_Mailboxes_AllDomains.csv'; $remote=Csv 'Exchange_OnPrem_RemoteMailboxes_AllDomains.csv'; $exo=Csv 'Exchange_EXO_Mailboxes_AllDomains.csv'; $stats=Csv 'Exchange_EXO_Mailboxes_AllDomains_Stats.csv'; $arch=Csv 'Exchange_EXO_Mailboxes_AllDomains_Archive.csv'; $perms=Csv 'Exchange_EXO_Mailboxes_AllDomains_Permissions.csv'; $m365=Csv 'M365_Users_Active.csv'; $lic=Csv 'M365_Licenses_Users.csv'; $domains=Csv 'Exchange_EXO_AcceptedDomains.csv'
  $localGuid=@{};$localUpn=@{};$localSmtp=@{}; foreach($r in $local){AddMap $localGuid (P $r ObjectGUID) $r; AddMap $localUpn (P $r UserPrincipalName) $r; AddMap $localSmtp (P $r PrimarySMTPaddress,PrimarySmtpAddress) $r}
  $remoteGuid=@{};$remoteUpn=@{};$remoteSmtp=@{}; foreach($r in $remote){AddMap $remoteGuid (P $r ObjectGuid,ObjectGUID) $r; AddMap $remoteUpn (P $r UserPrincipalName) $r; AddMap $remoteSmtp (P $r PrimarySmtpAddress,PrimarySMTPaddress) $r}
  $exoUpn=@{};$exoSmtp=@{}; foreach($r in $exo){AddMap $exoUpn (P $r UserPrincipalName) $r; AddMap $exoSmtp (P $r PrimarySmtpAddress,PrimarySMTPaddress) $r}
  $statsUpn=@{}; foreach($r in $stats){AddMap $statsUpn (P $r UserPrincipalName) $r}; $archUpn=@{}; foreach($r in $arch){AddMap $archUpn (P $r UserPrincipalName) $r}; $permsUpn=@{}; foreach($r in $perms){AddMap $permsUpn (P $r UserPrincipalName) $r}; $m365Upn=@{}; foreach($r in $m365){AddMap $m365Upn (P $r 'User principal name',UserPrincipalName) $r}
  $licUpn=@{}; foreach($r in $lic){$u=K (P $r 'User principal name',UserPrincipalName); if(!$u){continue}; if(!$licUpn.ContainsKey($u)){$licUpn[$u]=[Collections.Generic.List[string]]::new()}; $sku=T (P $r SkuPartNumber,'SKU name'); if($sku){[void]$licUpn[$u].Add($sku)}}
  $dupUpnSet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); foreach($r in $dupUpn){[void]$dupUpnSet.Add((T(P $r UserPrincipalName)))}
  $dupSmtpSet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); foreach($r in $dupSmtp){[void]$dupSmtpSet.Add((T(P $r SmtpAddress)))}
  $domSet=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase); foreach($r in $domains){[void]$domSet.Add((T(P $r DomainName,Name)))}
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
