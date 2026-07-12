<#
.SYNOPSIS
Generates Windows 11 readiness issue tables for PowerBI from SmartInventory CSV exports.

.REQUIREMENTS
- PowerShell 7+
- SmartM365.Core module
- Read access to DATA-LAST SmartInventory CSV exports
- Write access to the configured DATA-ALL, DATA-LAST, and WeeklyHistory output folders
- Required input CSV files: AD_Computers_AllDomains.csv, Intune_Devices_LocalSystem.csv,
  Intune_WindowsUpdate_Status.csv, Intune_Devices_Win11Readiness.csv,
  Intune_Devices_Inventory.csv, M365_Entra_Devices.csv, Intune_Devices_BIOS.csv,
  Intune_Devices_Compliance.csv

.VERSION
1.2
#>
#requires -Version 7.0
[CmdletBinding()]
param(
  [string]$Tenant='test',
  [string]$DataLastFolder='',
  [string]$OutputFolder='',
  [string]$LatestFolder='',
  [switch]$DisableSharePointUpload,
  [switch]$SkipLegacyAliases,
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
    throw "-MaxItems is not supported by SmartM365-Intune-Windows11-Readiness-Issues-Inventory because issue detection must be generated from complete readiness and inventory CSV snapshots."
}
$ErrorActionPreference='Stop'
$ScriptName='SmartM365-Intune-Windows11-Readiness-Issues-Inventory'
$ScriptVersion='1.2'
$RunStamp=Get-Date -Format 'yyyyMMdd-HHmmss'
function Log([string]$m){Write-Host ("{0} [INFO] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$m)}
function Warn([string]$m){Write-Warning $m}
function Root(){ $d=$PSScriptRoot; while($d){ if((Test-Path (Join-Path $d 'Config\SmartM365-TenantContext.ps1')) -and (Test-Path (Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'))){return $d}; $p=Split-Path $d -Parent; if(!$p -or $p -eq $d){break}; $d=$p }; throw 'SmartM365 root not found.' }
function LocalConfig(){ $p=Join-Path $PSScriptRoot (([IO.Path]::GetFileNameWithoutExtension($PSCommandPath))+'.local.json'); if(!(Test-Path $p)){ $t="$p.template"; if(!(Test-Path $t)){throw "Missing config template: $t"}; Copy-Item $t $p; Log "Created local config: $p" }; Get-Content $p -Raw | ConvertFrom-Json }
function ResolveToken($v){ if($v -isnot [string] -or [string]::IsNullOrWhiteSpace($v)){return $v}; $r=$v; 0..9|%{ $ms=[regex]::Matches($r,'\{\{(?<n>[A-Za-z0-9_.-]+)\}\}'); if($ms.Count -eq 0){return $r}; foreach($m in $ms){$pr=$script:Cfg.PSObject.Properties[$m.Groups['n'].Value]; if($pr){$r=$r.Replace($m.Value,[string](ResolveToken $pr.Value))}}}; $r }
function Cfg($c,$n,$d){ $p=$c.PSObject.Properties[$n]; if($p -and $null -ne $p.Value){ if($p.Value -isnot [string]){return ResolveToken $p.Value}; $txt=$p.Value.Trim(); if($txt -and $txt -notin @('__USE_GLOBAL__','USE_GLOBAL')){return ResolveToken $p.Value} }; $gp=$script:Cfg.PSObject.Properties[$n]; if($gp -and $null -ne $gp.Value){ if($gp.Value -is [string] -and [string]::IsNullOrWhiteSpace($gp.Value)){return $d}; return ResolveToken $gp.Value}; $d }
function P($r,[string[]]$names){ if($null -eq $r){return $null}; foreach($n in $names){$p=$r.PSObject.Properties[$n]; if($p -and $null -ne $p.Value){return $p.Value}}; $null }
function T($v){if($null -eq $v){''}else{([string]$v).Trim()}}
function K($v){(T $v).ToLowerInvariant()}
function N($v){(T $v).TrimEnd('$').ToUpperInvariant()}
function B($v){(K $v) -in @('true','1','yes','y','enabled','compliant')}
function CB($c,$n,[bool]$d){$v=Cfg $c $n $d; if($v -is [bool]){return $v}; if($null -eq $v){return $d}; $s=K $v; if(!$s){return $d}; $s -in @('true','1','yes','y','enabled')}
function Num($v){$s=(T $v)-replace ',','.'; $o=0.0; if([double]::TryParse($s,[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$o)){$o}else{$null}}
function Dt($v){$s=T $v; if(!$s){return $null}; foreach($c in @([Globalization.CultureInfo]::InvariantCulture,[Globalization.CultureInfo]::GetCultureInfo('fr-FR'),[Globalization.CultureInfo]::GetCultureInfo('en-US'))){$d=[datetime]::MinValue; if([datetime]::TryParse($s,$c,[Globalization.DateTimeStyles]::AssumeLocal,[ref]$d)){return $d}}; $null}
function Csv($name,[switch]$Req){$p=Join-Path $DataLastFolder $name; if(!(Test-Path $p)){if($Req){throw "Required CSV not found: $p"}; Warn "Optional CSV not found: $p"; return @()}; $r=@(Import-Csv $p); Log "Loaded $name : $($r.Count) row(s)"; $r}
function AddMap($map,$key,$val){$k=K $key; if($k -and !$map.ContainsKey($k)){$map[$k]=$val}}
function AddIssue($list,$code,$area,$id,$issue,$cat,[int]$prio,[bool]$block,$action,$impact){if(!$id){$id='UNRESOLVED'}; [void]$list.Add([pscustomobject]@{IssueCode=$code;Area=$area;ObjectGUID_Norm=$id;Potential_Issue=$issue;IssueCategory=$cat;PriorityScore=$prio;IsBlocking=$block;RecommendedAction=$action;ImpactMigration=$impact})}
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
  foreach($f in @($OutputFolder,$LatestFolder,(Join-Path $OutputFolder 'Archive'))){if(!(Test-Path $f)){New-Item -ItemType Directory -Path $f -Force|Out-Null}}
  $main=Join-Path $OutputFolder 'Intune_Windows11_Readiness_Issues.csv'; $latest=Join-Path $LatestFolder 'Intune_Windows11_Readiness_Issues.csv'; $archive=Join-Path (Join-Path $OutputFolder 'Archive') "Intune_Windows11_Readiness_Issues_$RunStamp.csv"
  Assert-SmartM365CsvDataCompleteness -Data $rows -TimestampedPath $main -LatestPath $latest; $rows|Sort-Object PriorityScore,IssueCode,ObjectGUID_Norm|Export-Csv $main -NoTypeInformation -Encoding UTF8; CopyCsv $main $latest; CopyCsv $main $archive
  $summary=@($rows|Group-Object IssueCode,Area,Potential_Issue,IssueCategory,PriorityScore,IsBlocking|Sort-Object Count -Descending|%{$p=$_.Name -split ', ',6; [pscustomobject]@{IssueCode=$p[0];Area=$p[1];Potential_Issue=$p[2];IssueCategory=$p[3];PriorityScore=[int]$p[4];IsBlocking=[bool]::Parse($p[5]);Count=$_.Count}})
  $sm=Join-Path $OutputFolder 'Intune_Windows11_Readiness_Issues_Summary.csv'; $sl=Join-Path $LatestFolder 'Intune_Windows11_Readiness_Issues_Summary.csv'; Assert-SmartM365CsvDataCompleteness -Data $summary -TimestampedPath $sm -LatestPath $sl; $summary|Export-Csv $sm -NoTypeInformation -Encoding UTF8; CopyCsv $sm $sl
  $files=@($main,$latest,$archive,$sm,$sl)
  if(!$SkipLegacyAliases){$lm=Join-Path $OutputFolder 'Mig_Win11Migration_Issues_Expanded.csv'; $ll=Join-Path $LatestFolder 'Mig_Win11Migration_Issues_Expanded.csv'; CopyCsv $main $lm; CopyCsv $main $ll; $files+=@($lm,$ll)}
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
  $sr=Root; . (Join-Path $sr 'Config\SmartM365-TenantContext.ps1'); $script:Cfg=Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot; Import-Module (Join-Path $sr 'Modules\SmartM365.Core\SmartM365.Core.psd1') -Force; Initialize-SmartM365DefaultCsvValidationRules
  $lc=LocalConfig
  if(!$DataLastFolder){$DataLastFolder=Cfg $lc 'InputDataLastFolder' (Cfg $lc 'LatestCsvFolderPath' $PSScriptRoot)}
  if(!$OutputFolder){$OutputFolder=Cfg $lc 'ScriptCsvLogFolderPath' (Join-Path (Cfg $lc 'DataAllRootPath' $PSScriptRoot) 'Intune\WindowsUpdate\Windows11ReadinessIssues')}
  if(!$LatestFolder){$LatestFolder=Cfg $lc 'LatestCsvFolderPath' $OutputFolder}
  $global:EnableSharePointUpload=CB $lc 'EnableSharePointUpload' $false; if($DisableSharePointUpload){$global:EnableSharePointUpload=$false}; if(!(CB $lc 'EmitLegacyPowerBIAlias' $true)){$SkipLegacyAliases=$true}
  $global:SharePointSiteHostname=Cfg $lc 'SharePointSiteHostname' ''; $global:SharePointSitePath=Cfg $lc 'SharePointSitePath' ''; $global:SharePointLibraryDisplayName=Cfg $lc 'SharePointLibraryDisplayName' 'Documents'; $global:SharePointTargetFolderPath=Cfg $lc 'SharePointTargetFolderPath' ''; $global:AppId=Cfg $lc 'AppId' ''; $global:TenantId=Cfg $lc 'TenantId' ''; $global:Thumbprint=Cfg $lc 'Thumbprint' (Cfg $lc 'Thumb' '')
  Log "DataLastFolder: $DataLastFolder"; Log "OutputFolder: $OutputFolder"; Log "LatestFolder: $LatestFolder"
  Invoke-SmartM365Preflight -ScriptName $ScriptName -OutputPaths @($OutputFolder,$LatestFolder) | Out-Null
  $ad=Csv 'AD_Computers_AllDomains.csv' -Req; $local=Csv 'Intune_Devices_LocalSystem.csv' -Req; $wu=Csv 'Intune_WindowsUpdate_Status.csv' -Req; $ready=Csv 'Intune_Devices_Win11Readiness.csv' -Req; $intune=Csv 'Intune_Devices_Inventory.csv' -Req; $entra=Csv 'M365_Entra_Devices.csv' -Req; $bios=Csv 'Intune_Devices_BIOS.csv' -Req; $comp=Csv 'Intune_Devices_Compliance.csv' -Req
  $adName=@{}; foreach($r in $ad){AddMap $adName (N(P $r Name,SamAccountName)) $r}
  $localAad=@{};$localName=@{}; foreach($r in $local){AddMap $localAad (P $r AzureADDeviceId) $r; AddMap $localName (N(P $r DeviceName)) $r}
  $readyName=@{};$readyGraph=@{}; foreach($r in $ready){AddMap $readyName (N(P $r NormalizedDeviceName,DeviceName)) $r; AddMap $readyGraph (P $r GraphId) $r}
  $intuneName=@{};$intuneAad=@{};$intuneId=@{}; foreach($r in $intune){AddMap $intuneName (N(P $r 'Device name',DeviceName)) $r; AddMap $intuneAad (P $r 'Azure AD Device ID',AzureADDeviceId,'Entra DeviceId') $r; AddMap $intuneId (P $r 'Device ID',DeviceId) $r}
  $entraId=@{};$entraName=@{}; foreach($r in $entra){AddMap $entraId (P $r DeviceId) $r; AddMap $entraName (N(P $r DisplayName)) $r}
  $biosAad=@{};$biosName=@{}; foreach($r in $bios){AddMap $biosAad (P $r AzureADDeviceId) $r; AddMap $biosName (N(P $r DeviceName)) $r}
  $compAad=@{};$compName=@{}; foreach($r in $comp){AddMap $compAad (P $r AzureADDeviceId) $r; AddMap $compName (N(P $r DeviceName,displayName)) $r}
  $wuId=@{};$wuName=@{}; foreach($r in $wu){$ed=Dt(P $r ExportDateTime); $did=K(P $r DeviceId); $nm=N(P $r NormalizedDeviceName,DeviceName); foreach($item in @(@($wuId,$did),@($wuName,$nm))){$m=$item[0];$k=$item[1]; if(!$k){continue}; $old=if($m.ContainsKey($k)){$m[$k]}else{$null}; if(!$old -or ($ed -and (Dt(P $old ExportDateTime)) -lt $ed)){$m[$k]=$r}}}
  $issues=[Collections.Generic.List[object]]::new(); $seen=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach($a in $ad){
    $name=N(P $a Name,SamAccountName); [void]$seen.Add($name); $guid=T(P $a ObjectGUID); $enabled=B(P $a Enabled); $os=T(P $a OperatingSystem); $osu=$os.ToUpperInvariant(); $is10=$osu.Contains('WINDOWS 10'); $is11=$osu.Contains('WINDOWS 11'); $last=Dt(P $a LastLogonDate)
    $in=if($intuneName.ContainsKey($name)){$intuneName[$name]}else{$null}; $aad=T(P $in 'Azure AD Device ID',AzureADDeviceId,'Entra DeviceId'); $devId=T(P $in 'Device ID',DeviceId)
    $ls=if($localAad.ContainsKey((K $aad))){$localAad[(K $aad)]}elseif($localName.ContainsKey($name)){$localName[$name]}else{$null}; $rd=if($readyName.ContainsKey($name)){$readyName[$name]}elseif($readyGraph.ContainsKey((K $aad))){$readyGraph[(K $aad)]}else{$null}; $en=if($entraId.ContainsKey((K $aad))){$entraId[(K $aad)]}elseif($entraName.ContainsKey($name)){$entraName[$name]}else{$null}; $bi=if($biosAad.ContainsKey((K $aad))){$biosAad[(K $aad)]}elseif($biosName.ContainsKey($name)){$biosName[$name]}else{$null}; $co=if($compAad.ContainsKey((K $aad))){$compAad[(K $aad)]}elseif($compName.ContainsKey($name)){$compName[$name]}else{$null}; $wr=if($wuId.ContainsKey((K $devId))){$wuId[(K $devId)]}elseif($wuName.ContainsKey($name)){$wuName[$name]}else{$null}
    $elig=(T(P $rd UpgradeEligibility,UpgradeEligibilityLabel)).ToUpperInvariant(); $cap=$elig -match 'CAPABLE|ELIGIBLE|TRUE'; $badElig=$elig -match 'NOT|UNKNOWN|FALSE|INELIGIBLE'; $check=Dt(P $in 'Last check-in',LastSyncDateTime); $entraSign=Dt(P $en ApproximateLastSignInDateTime); $wuDate=Dt(P $wr ExportDateTime); $wuState=(T(P $wr AggregateState_loc,AggregateState)).ToUpperInvariant(); $wuStatus=(T(P $wr CurrentDeviceUpdateStatus_loc,CurrentDeviceUpdateStatus)).ToUpperInvariant(); $alert=T(P $wr LatestAlertMessage_loc,LatestAlertMessage,BlockingReason); $mem=Num(P $in PhysicalMemoryGB); $free=Num(P $in 'Free storage'); $biosDate=Dt(P $bi BIOSReleaseDateTime,BIOSDate)
    if(!$enabled){AddIssue $issues 'ISSUE-006' 'Directory' $guid 'Account disabled' '4.Low' 4 $false 'Review device account status.' 'Device disabled; upgrade or migration is likely not applicable.'}
    if($enabled -and ($is10 -or $is11) -and !$in){AddIssue $issues 'ISSUE-010' 'Intune' $guid 'Windows device not found in Intune inventory' '2.High' 2 $true 'Enroll or reconcile the device in Intune before Windows 11 rollout decisions.' 'Blocking: no Intune inventory signal for this active Windows device.'}
    if($last -and ((Get-Date).Date-$last.Date).Days -gt 56){AddIssue $issues 'ISSUE-011-AD' 'Directory' $guid 'Device inactive for more than 56 days (LastLogon AD)' '4.Low' 4 $false 'Review device activity in AD; retire or decommission if obsolete.' 'Device inactive in AD; migration or upgrade may be unnecessary.'}
    if($check -and ((Get-Date).Date-$check.Date).Days -gt 56){AddIssue $issues 'ISSUE-011-M365' 'Intune' $guid 'Device inactive for more than 56 days (Last check-in Intune)' '4.Low' 4 $false 'Review device activity in Intune; retire or remediate if still required.' 'Device inactive in Intune; state may be stale for readiness decisions.'}
    if($entraSign -and ((Get-Date).Date-$entraSign.Date).Days -gt 90){AddIssue $issues 'ISSUE-024' 'Entra' $guid 'Azure Entra stale: last sign-in older than 90 days' '4.Low' 4 $false 'Review device usage and retire or remediate if still required.' 'Cloud device appears inactive; readiness decision may be unnecessary or stale.'}
    if($is10 -and $cap){AddIssue $issues 'ISSUE-038' 'OS' $guid 'Windows 10 device capable for Windows 11 but not yet upgraded' '4.Low' 4 $false 'Prioritize this device for Windows 11 rollout when business timing allows.' 'Good upgrade candidate; no technical blocker detected in readiness source.'}
    if($null -eq $mem){AddIssue $issues 'ISSUE-042' 'Hardware' $guid 'RAM capacity unknown' '4.Low' 4 $false 'Fix inventory enrichment to report memory before scheduling upgrade.' 'Advisory: inventory incomplete; validate device capacity.'}
    if(!$os){AddIssue $issues 'ISSUE-043' 'OS' $guid 'Operating system identification missing' '4.Low' 4 $false 'Fix OS inventory reporting to identify Windows version/build.' 'Advisory: OS not identified; cannot determine eligibility or baseline.'}
    if($is11 -and $badElig){AddIssue $issues 'ISSUE-051' 'OS' $guid 'Inconsistent inventory: Windows 11 device flagged Not Capable or Unknown eligibility' '4.Low' 4 $false 'Re-evaluate eligibility signals and correct readiness pipeline.' 'Advisory: data quality issue; eligibility flags should reflect actual state.'}
    if(!$ls){AddIssue $issues 'ISSUE-071' 'BIOS' $guid 'No Secure Boot / firmware data available' '4.Low' 4 $false 'Ensure the local system inventory/remediation script is assigned and has run.' 'Advisory: cannot assess BIOS/Secure Boot readiness until platform data is available.'}
    if($biosDate -and $biosDate -lt [datetime]'2019-01-01'){AddIssue $issues 'ISSUE-074' 'BIOS' $guid 'BIOS firmware older than 01/01/2019 - update recommended before Windows 11 upgrade' '4.Low' 4 $false 'Update BIOS firmware to latest vendor version; verify Secure Boot and TPM settings after update.' 'Advisory: outdated BIOS may cause Windows 11 compatibility or firmware setting issues.'}
    if($null -ne $free -and $free -lt 20){AddIssue $issues 'ISSUE-054' 'Storage' $guid 'Low free storage for Windows 11 upgrade' '2.High' 2 $true 'Free disk space or remediate storage before scheduling upgrade.' 'Blocking: low storage can prevent Windows feature update completion.'}
    $cs=(T(P $co state,Compliance,ComplianceState)).ToUpperInvariant(); if($cs -match 'NONCOMPLIANT|ERROR|FAILED'){AddIssue $issues 'ISSUE-055' 'Compliance' $guid 'Device is non-compliant' '2.High' 2 $true 'Resolve compliance issues before Windows 11 rollout.' 'Blocking: non-compliance may prevent policy or update success.'}
    if($wr){ if($wuState -match 'ERROR|FAILED|ATTENTION|OFFERFAILED' -or $wuStatus -match 'ERROR|FAILED|ATTENTION|ROLLBACK'){AddIssue $issues 'ISSUE-062' 'Windows Update' $guid 'Windows Update reports failed or attention-required state' '2.High' 2 $true 'Review Windows Update for Business status and remediate the reported error.' 'Blocking: update service reports failure or attention required.'}; if($alert){AddIssue $issues 'ISSUE-063' 'Windows Update' $guid 'Windows Update latest alert message present' '3.Medium' 3 $true 'Review the latest alert message and remediate if still current.' 'Potential blocker: Windows Update reported an alert for this device.'}; if($wuDate -and ((Get-Date).Date-$wuDate.Date).Days -ge 14){AddIssue $issues 'ISSUE-068' 'Windows Update' $guid 'Windows Update data is stale (>= 14 days since last export)' '4.Low' 4 $false 'Refresh Windows Update status export from Intune and validate pipeline health.' 'Advisory: WU-based issues may not reflect current device state.'} }
    elseif($is10 -and $cap -and $in){AddIssue $issues 'ISSUE-069' 'Windows Update' $guid 'Windows 10 capable device enrolled in Intune but no Windows Update policy signal detected' '3.Medium' 3 $true 'Assign or verify Windows Update for Business policy coverage in Intune.' 'Device cannot be upgraded via WUfB without active update policy assignment.'}
  }
  foreach($i in $intune){$name=N(P $i 'Device name',DeviceName); if(!$name -or $seen.Contains($name)){continue}; AddIssue $issues 'ISSUE-002' 'Intune' (T(P $i 'Azure AD Device ID','Device ID','Entra DeviceId')) 'Device exists in Intune but not found in AD (orphan MDM device)' '2.High' 2 $true 'Identify device owner and re-join to AD or retire from Intune.' 'Orphan MDM device cannot be managed through expected hybrid scope.'}
  $files=ExportIssues -rows @($issues); $files += PublishWeeklyHistory -files $files; Log "Generated Windows 11 readiness issues: $($issues.Count) row(s)"
  if($global:EnableSharePointUpload){foreach($f in $files){try{Invoke-SmartM365SharePointCsvUpload -LocalFilePath $f|Out-Null; Log "SharePoint upload completed: $f"}catch{Warn "SharePoint upload failed for $f : $($_.Exception.Message)"}}}
  Log "$ScriptName completed successfully."
}catch{Write-Error $_; throw}
# SIG # Begin signature block
# MIIHJAYJKoZIhvcNAQcCoIIHFTCCBxECAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCgF4lxYP6A4kvW
# JT07w3fnnbV1+KkS5N1/8ytkgD42jKCCBBQwggQQMIICeKADAgECAhBwIfLVIgJW
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
# BDEiBCBiMotm9MoxdX0s6nNjgWAKN9xJosPJpFM2QLcwNM7X1TANBgkqhkiG9w0B
# AQEFAASCAYAvseRzKtvUmH2JCuYyLK+2ZS3ltbpXM1PmcL9jwMXHgS0wUEvrQT+3
# qkngmWr4mFy2PmZBT/zz5fIyW6ZLMhVQsdkkIOaYFrdkBfZxtAmGA/T4aKYGW8SJ
# lNccgAYx20G6NZtbAyWrdnMQhe9Nsp3EkLRhLaDdOGLthBSp36BQYLX5cwtjicVL
# V5h0Q1BlHQOBOt4fTwJOhLxQSRRWYc8vQgHsie29yeS4EBs183aSqMrNAxKVJdkJ
# vPzmdvfvhjF2K6OpLuozEr8HO4mkIEkywYcjYvzs1r5H5HqSgNHKqhsFwzwbKyoe
# ej/siWmlM2CZSxq+UfS2LnnCgPicOCvI17xdlXMO2fCh2EHHPhQ+XsjyC7IRoVcH
# nq+5W1+mvJkfCMqa+zFHMvkKb048Opn8jujdkADVMgGfiXCiQATd6eH6Xr/oAodH
# H16LXLOlfp/ACmMMAjx2RXQ0f4GCRHTh7zTNGVwCML6xhSZLvez6osxlFACHClpE
# fsOllvgl9gc=
# SIG # End signature block
