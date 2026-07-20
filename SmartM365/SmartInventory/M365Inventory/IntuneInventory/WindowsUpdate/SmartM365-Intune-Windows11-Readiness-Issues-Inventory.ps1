<#
.SYNOPSIS
Generates Windows 11 readiness issue tables from SmartInventory CSV exports using the legacy Power BI rule set.

.REQUIREMENTS
- PowerShell 7+
- SmartM365.Core module
- Read access to DATA-LAST SmartInventory CSV exports
- Write access to the configured DATA-ALL, DATA-LAST, and WeeklyHistory output folders
- Required input CSV files: AD_Computers_AllDomains.csv or AD_Computers_AllDomains_Brut.csv, Intune_Devices_Inventory.csv, M365_Entra_Devices.csv,
  Intune_WindowsUpdate_Status.csv. Optional enrichment CSV files: AD_Users_AllDomains.csv,
  M365_Licenses_Users.csv, Intune_Devices_LocalSystem.csv, Intune_Devices_BIOS.csv,
  Intune_Devices_Compliance.csv, Intune_Devices_UpgradeEligibility.csv, M365_Entra_Devices_HardwareIdConflicts.csv

.VERSION
1.20
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
    throw "-MaxItems is not supported by SmartM365-Intune-Windows11-Readiness-Issues-Inventory because issue detection must be generated from complete readiness and inventory CSV snapshots."
}
$ErrorActionPreference='Stop'
$ScriptName='SmartM365-Intune-Windows11-Readiness-Issues-Inventory'
$ScriptVersion="1.20"
$RunStamp=Get-Date -Format 'yyyyMMdd-HHmmss'
$RunStartedAt=Get-Date
$script:WarningCount=0
$script:ErrorCount=0
$script:GeneratedFileCount=0
function Log([string]$m){Write-Host ("{0} [INFO] {1}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$m)}
function Warn([string]$m){$script:WarningCount++;Write-Warning $m}
function Root(){ $d=$PSScriptRoot; while($d){ if((Test-Path (Join-Path $d 'Config\SmartM365-TenantContext.ps1')) -and (Test-Path (Join-Path $d 'Modules\SmartM365.Core\SmartM365.Core.psd1'))){return $d}; $p=Split-Path $d -Parent; if(!$p -or $p -eq $d){break}; $d=$p }; throw 'SmartM365 root not found.' }
function LocalConfig(){ $p=Join-Path $PSScriptRoot (([IO.Path]::GetFileNameWithoutExtension($PSCommandPath))+'.local.json'); if(!(Test-Path $p)){ $t="$p.template"; if(!(Test-Path $t)){throw "Missing config template: $t"}; Copy-Item $t $p; Log "Created local config: $p" }; Get-Content $p -Raw | ConvertFrom-Json }
function ResolveToken($v){ if($v -isnot [string] -or [string]::IsNullOrWhiteSpace($v)){return $v}; $r=[string]$v; for($iteration=0;$iteration -lt 10;$iteration++){ $ms=[regex]::Matches($r,'\{\{(?<n>[A-Za-z0-9_.-]+)\}\}'); if($ms.Count -eq 0){break}; $changed=$false; foreach($m in $ms){$pr=$script:Cfg.PSObject.Properties[$m.Groups['n'].Value]; if($pr){$replacement=ResolveToken $pr.Value; if($replacement -is [array]){throw "Configuration token '$($m.Value)' resolved to multiple values."}; $next=$r.Replace($m.Value,[string]$replacement); if($next -ne $r){$changed=$true;$r=$next}}}; if(!$changed){break} }; return $r }
function AssertResolvedPath([string]$Name,[string]$Path){ if([string]::IsNullOrWhiteSpace($Path)){throw "Configuration path '$Name' is empty."}; if($Path -match '\{\{[A-Za-z0-9_.-]+\}\}'){throw "Configuration path '$Name' contains an unresolved token: $Path"}; if(([regex]::Matches($Path,'(?i)(?:[A-Z]:\\|\\\\)').Count) -gt 1){throw "Configuration path '$Name' contains multiple path roots: $Path"} }
function Cfg($c,$n,$d){ $p=$c.PSObject.Properties[$n]; if($p -and $null -ne $p.Value){ if($p.Value -isnot [string]){return ResolveToken $p.Value}; $txt=$p.Value.Trim(); if($txt -and $txt -notin @('__USE_GLOBAL__','USE_GLOBAL')){return ResolveToken $p.Value} }; $gp=$script:Cfg.PSObject.Properties[$n]; if($gp -and $null -ne $gp.Value){ if($gp.Value -is [string] -and [string]::IsNullOrWhiteSpace($gp.Value)){return $d}; return ResolveToken $gp.Value}; $d }
function P($r,[string[]]$names){ if($null -eq $r){return $null}; foreach($n in $names){$p=$r.PSObject.Properties[$n]; if($p -and $null -ne $p.Value){return $p.Value}}; $null }
function T($v){if($null -eq $v){''}else{([string]$v).Trim()}}
function K($v){(T $v).ToLowerInvariant()}
function N($v){(T $v).TrimEnd('$').ToUpperInvariant()}
function G($v){
  $s=(T $v).Trim('{','}')
  if(!$s){return ''}
  $g=[guid]::Empty
  if([guid]::TryParse($s,[ref]$g)){return $g.ToString('D').ToLowerInvariant()}
  $s.ToLowerInvariant()
}
function B($v){(K $v) -in @('true','1','yes','y','enabled','compliant')}
function CB($c,$n,[bool]$d){$v=Cfg $c $n $d; if($v -is [bool]){return $v}; if($null -eq $v){return $d}; $s=K $v; if(!$s){return $d}; $s -in @('true','1','yes','y','enabled')}
function Num($v){$s=(T $v)-replace ',','.'; $o=0.0; if([double]::TryParse($s,[Globalization.NumberStyles]::Any,[Globalization.CultureInfo]::InvariantCulture,[ref]$o)){$o}else{$null}}
function Dt($v){$s=T $v; if(!$s){return $null}; foreach($c in @([Globalization.CultureInfo]::InvariantCulture,[Globalization.CultureInfo]::GetCultureInfo('fr-FR'),[Globalization.CultureInfo]::GetCultureInfo('en-US'))){$d=[datetime]::MinValue; if([datetime]::TryParse($s,$c,[Globalization.DateTimeStyles]::AssumeLocal,[ref]$d)){return $d}}; $null}
function Csv($name,[switch]$Req){$p=Join-Path $DataLastFolder $name; if(!(Test-Path $p)){if($Req){throw "Required CSV not found: $p"}; Log "Optional CSV not found: $p"; return @()}; $r=@(Import-Csv $p); Log "Loaded $name : $($r.Count) row(s)"; $r}

function CsvAny([string[]]$Names,[switch]$Req){
  foreach($name in $Names){
    $path=Join-Path $DataLastFolder $name
    if(Test-Path -LiteralPath $path){$rows=@(Import-Csv -LiteralPath $path); Log "Loaded $name : $($rows.Count) row(s)"; return $rows}
  }
  if($Req){throw "Required CSV not found. Checked: $($Names -join ', ') in $DataLastFolder"}
  Log "Optional CSV not found. Checked: $($Names -join ', ') in $DataLastFolder"
  @()
}
function CsvAnyProjected([string[]]$Names,[string[]]$Columns,[switch]$Req){
  $uniqueColumns=[Collections.Generic.List[string]]::new()
  $seenColumns=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach($column in $Columns){if($seenColumns.Add($column)){[void]$uniqueColumns.Add($column)}}
  foreach($name in $Names){
    $path=Join-Path $DataLastFolder $name
    if(Test-Path -LiteralPath $path){
      $rows=@(Import-Csv -LiteralPath $path | Select-Object -Property $uniqueColumns.ToArray())
      Log "Loaded $name : $($rows.Count) row(s), projected to $($uniqueColumns.Count) column(s)"
      return $rows
    }
  }
  if($Req){throw "Required CSV not found. Checked: $($Names -join ', ') in $DataLastFolder"}
  Log "Optional CSV not found. Checked: $($Names -join ', ') in $DataLastFolder"
  @()
}


function Add-IssueRow{
  param([Collections.Generic.List[object]]$Target,[string]$IssueCode,[string]$Area,[string]$ObjectGUID_Norm,[string]$Potential_Issue,[string]$IssueCategory,[int]$PriorityScore,[bool]$IsBlocking,[string]$RecommendedAction,[string]$ImpactMigration)
  if(!$ObjectGUID_Norm){$ObjectGUID_Norm='UNRESOLVED'}
  [void]$Target.Add([pscustomobject]@{IssueCode=$IssueCode;Area=$Area;ObjectGUID_Norm=$ObjectGUID_Norm;Potential_Issue=$Potential_Issue;IssueCategory=$IssueCategory;PriorityScore=$PriorityScore;IsBlocking=$IsBlocking;RecommendedAction=$RecommendedAction;ImpactMigration=$ImpactMigration})
}
function FirstValue($row,[string[]]$names){P $row $names}
function Get-UpperTrimmedString($Value){(T $Value).ToUpperInvariant()}
function ParseBuild($adVersion,$intuneVersion){
  $iv=T $intuneVersion
  if($iv -match '^\d+\.\d+\.(\d{4,5})(?:\.|$)'){return $matches[1]}
  $rv=T $adVersion
  if($rv -match '\((\d{4,5})(?:\.|\))'){return $matches[1]}
  if($rv -match '^(\d{4,5})(?:\.|$)'){return $matches[1]}
  ''
}
function FriendlyOs($adRow,$intuneRow){
  $adOs=T(FirstValue $adRow @('OperatingSystem'))
  $build=ParseBuild (FirstValue $adRow @('operatingSystemVersion','OperatingSystemVersion')) (FirstValue $intuneRow @('OS version','OSVersion','OperatingSystemVersion'))
  if(!$build){if($adOs){return $adOs}; return '?'}
  $b5=if($build.Length -ge 5){$build.Substring(0,5)}else{$build}
  $b4=if($build.Length -ge 4){$build.Substring(0,4)}else{$build}
  $b2=if($build.Length -ge 2){$build.Substring(0,2)}else{$build}
  switch($b5){
    '28000'{return 'Windows 11 26H1'} '26200'{return 'Windows 11 25H2'} '26100'{return 'Windows 11 24H2'} '22631'{return 'Windows 11 23H2'} '22621'{return 'Windows 11 22H2'} '22000'{return 'Windows 11 21H2'}
    '19045'{return 'Windows 10 22H2'} '19044'{return 'Windows 10 21H2'} '19043'{return 'Windows 10 21H1'} '19042'{return 'Windows 10 20H2'} '19041'{return 'Windows 10 2004'} '18363'{return 'Windows 10 1909'} '18362'{return 'Windows 10 1903'} '17763'{return 'Windows 10 1809'} '17134'{return 'Windows 10 1803'} '16299'{return 'Windows 10 1709'} '15063'{return 'Windows 10 1703'} '14393'{return 'Windows 10 1607'} '10586'{return 'Windows 10 1511'} '10240'{return 'Windows 10 1507'}
  }
  if($b2 -in @('28','26')){return "Windows 11 Insider ($build)"}
  switch($b4){'9600'{return 'Windows 8.1'} '9200'{return 'Windows 8'} '7601'{return 'Windows 7 SP1'} '7600'{return 'Windows 7'}}
  if($adOs){return "$adOs ($build)"}
  "OS Unknown ($build)"
}
function EligibilityLabel($value){
  $s=(T $value).Replace([char]0xA0,' ').Trim().ToUpperInvariant()
  $compact=$s -replace '\s',''
  switch($compact){'0'{return 'UPGRADED'} '2'{return 'NOT CAPABLE'} '3'{return 'CAPABLE'} 'UPGRADED'{return 'UPGRADED'} 'CAPABLE'{return 'CAPABLE'} 'NOTCAPABLE'{return 'NOT CAPABLE'}}
  if($s -in @('UNKNOWN','UNDETERMINED')){return $s}
  ''
}

function FreeGb($value){
  $n=Num $value
  if($null -eq $n){return $null}
  if($n -gt 1048576){return [math]::Round($n / 1GB,2)}
  if($n -gt 1024){return [math]::Round($n / 1024,2)}
  [math]::Round($n,2)
}
function DaysSince($dateValue){$dt=Dt $dateValue; if($null -eq $dt){return $null}; ([datetime]::Today - $dt.Date).Days}
function AddAllMap($map,$key,$val){$k=K $key; if(!$k){return}; if(!$map.ContainsKey($k)){$map[$k]=[Collections.Generic.List[object]]::new()}; [void]$map[$k].Add($val)}
function LatestByDate($rows,[string[]]$dateNames){
  $latest=$null; $latestDate=$null; $latestRun=-1
  foreach($row in @($rows)){
    $dt=$null; foreach($dn in $dateNames){$dt=Dt(FirstValue $row @($dn)); if($dt){break}}
    $run=Num(FirstValue $row @('RunId','RunID')) ; if($null -eq $run){$run=0}
    if(!$latest -or ($dt -and (!$latestDate -or $dt -gt $latestDate)) -or ($dt -and $latestDate -and $dt -eq $latestDate -and $run -gt $latestRun)){$latest=$row; $latestDate=$dt; $latestRun=$run}
  }
  $latest
}
function GetMapValue($map,$key){$k=T $key; if($k -and $map.ContainsKey($k)){$map[$k]}else{$null}}
function CountMapValue($map,$key){$k=T $key; if($k -and $map.ContainsKey($k)){@($map[$k]).Count}else{0}}
function NormalizeLicenseStatus($value){
  $s=T $value
  if(!$s){return 'LICENSE_SIGNAL_MISSING'}
  if($s -eq '?'){return 'LICENSE_SIGNAL_MISSING'}
  if($s -in @('None','Not in M365','No License (Unlicensed)','No License (Shared)')){return 'USER_NO_LICENSE'}
  'USER_LICENSED'
}

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
  Assert-SmartM365CsvDataCompleteness -Data $rows -TimestampedPath $main -LatestPath $latest; $rows|Sort-Object PriorityScore,IssueCode,ObjectGUID_Norm|Add-SmartM365TenantKey | Export-Csv $main -NoTypeInformation -Encoding UTF8; CopyCsv $main $latest; CopyCsv $main $archive
  $summary=@($rows|Group-Object IssueCode,Area,Potential_Issue,IssueCategory,PriorityScore,IsBlocking|Sort-Object Count -Descending|%{$p=$_.Name -split ', ',6; [pscustomobject]@{IssueCode=$p[0];Area=$p[1];Potential_Issue=$p[2];IssueCategory=$p[3];PriorityScore=[int]$p[4];IsBlocking=[bool]::Parse($p[5]);Count=$_.Count}})
  $sm=Join-Path $OutputFolder 'Intune_Windows11_Readiness_Issues_Summary.csv'; $sl=Join-Path $LatestFolder 'Intune_Windows11_Readiness_Issues_Summary.csv'; Assert-SmartM365CsvDataCompleteness -Data $summary -TimestampedPath $sm -LatestPath $sl; $summary|Add-SmartM365TenantKey | Export-Csv $sm -NoTypeInformation -Encoding UTF8; CopyCsv $sm $sl
  @($main,$latest,$archive,$sm,$sl)
}
function PublishWeeklyHistory($files){
  if(!(CB $lc 'EnableWeeklyHistory' $true)){return @()}
  $base=Cfg $lc 'WeeklyHistoryFolderPath' (Join-Path $OutputFolder 'WeeklyHistory')
  if([string]::IsNullOrWhiteSpace($base)){return @()}
  $year=[Globalization.ISOWeek]::GetYear((Get-Date)); $week=[Globalization.ISOWeek]::GetWeekOfYear((Get-Date)); $weekName=('{0}-W{1:00}' -f $year,$week); $folder=Join-Path $base $weekName
  if(!(Test-Path $folder)){New-Item -ItemType Directory -Path $folder -Force|Out-Null}

  $outputRoot=[IO.Path]::GetFullPath($OutputFolder).TrimEnd('\')
  $sourceByLeaf=[ordered]@{}
  foreach($f in @($files)){
    if(!$f -or !(Test-Path -LiteralPath $f)){continue}
    $parent=[IO.Path]::GetFullPath((Split-Path $f -Parent)).TrimEnd('\')
    if(-not [string]::Equals($parent,$outputRoot,[StringComparison]::OrdinalIgnoreCase)){continue}
    $leaf=Split-Path $f -Leaf
    if(-not $sourceByLeaf.Contains($leaf)){$sourceByLeaf[$leaf]=$f}
  }
  if($sourceByLeaf.Count -eq 0){return @()}

  $manifest=Join-Path $folder 'manifest.json'
  $snapshotComplete=(Test-Path -LiteralPath $manifest)
  if($snapshotComplete){
    foreach($leaf in $sourceByLeaf.Keys){
      if(!(Test-Path -LiteralPath (Join-Path $folder $leaf))){$snapshotComplete=$false;break}
    }
  }
  if($snapshotComplete){
    Log "Weekly history already complete for $weekName. Snapshot and SharePoint republication skipped."
    return @()
  }

  $published=[Collections.Generic.List[string]]::new()
  foreach($leaf in $sourceByLeaf.Keys){
    $dest=Join-Path $folder $leaf
    CopyCsv $sourceByLeaf[$leaf] $dest
    [void]$published.Add($dest)
  }
  [pscustomobject]@{ScriptName=$ScriptName;ScriptVersion=$ScriptVersion;Tenant=$Tenant;GeneratedOn=(Get-Date).ToString('o');Week=$weekName;Files=@($published|ForEach-Object{Split-Path $_ -Leaf})} | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $manifest -Encoding UTF8
  [void]$published.Add($manifest)
  @($published)
}
$script:CompletionStatus = 'Auto'
try{
  Log "Starting $ScriptName v$ScriptVersion"
  $sr=Root; . (Join-Path $sr 'Config\SmartM365-TenantContext.ps1'); $script:Cfg=Initialize-SmartM365TenantContext -Tenant $Tenant -StartPath $PSScriptRoot; Import-Module -Name (Join-Path $sr 'Modules\SmartM365.Core\SmartM365.Core.psd1') -MinimumVersion '1.0.24' -Force; Initialize-SmartM365DefaultCsvValidationRules
  $lc=LocalConfig
  if(!$DataLastFolder){$DataLastFolder=Cfg $lc 'InputDataLastFolder' (Cfg $lc 'LatestCsvFolderPath' $PSScriptRoot)}
  if(!$OutputFolder){$OutputFolder=Cfg $lc 'ScriptCsvLogFolderPath' (Join-Path (Cfg $lc 'DataAllRootPath' $PSScriptRoot) 'Intune\WindowsUpdate\Windows11ReadinessIssues')}
  if(!$LatestFolder){$LatestFolder=Cfg $lc 'LatestCsvFolderPath' $OutputFolder}
  AssertResolvedPath 'DataLastFolder' $DataLastFolder
  AssertResolvedPath 'OutputFolder' $OutputFolder
  AssertResolvedPath 'LatestFolder' $LatestFolder
  $global:EnableSharePointUpload=CB $lc 'EnableSharePointUpload' $false; if($DisableSharePointUpload){$global:EnableSharePointUpload=$false}
  $global:SharePointSiteHostname=Cfg $lc 'SharePointSiteHostname' ''; $global:SharePointSitePath=Cfg $lc 'SharePointSitePath' ''; $global:SharePointLibraryDisplayName=Cfg $lc 'SharePointLibraryDisplayName' 'Documents'; $global:SharePointTargetFolderPath=Cfg $lc 'SharePointTargetFolderPath' ''; $global:AppId=Cfg $lc 'AppId' ''; $global:TenantId=Cfg $lc 'TenantId' ''; $global:Thumbprint=Cfg $lc 'Thumbprint' (Cfg $lc 'Thumb' '')
  Log "DataLastFolder: $DataLastFolder"; Log "OutputFolder: $OutputFolder"; Log "LatestFolder: $LatestFolder"
  Invoke-SmartM365Preflight -ScriptName $ScriptName -OutputPaths @($OutputFolder,$LatestFolder) | Out-Null

  $ad=CsvAnyProjected @('AD_Computers_AllDomains.csv','AD_Computers_AllDomains_Brut.csv') @('ObjectGuidNormalized','ObjectGUID','IntuneDeviceId','EntraDeviceIdNormalized','AzureADDeviceId','AzureADDeviceId_Norm','DeviceId_Norm','Name','SamAccountName','OperatingSystem','operatingSystemVersion','OperatingSystemVersion','LastLogonDate','Enabled','OrganizationalUnit','CanonicalName','DistinguishedName','LastRebootDateTime','LastRebootDate','LastBootUpTime','LastLoggedUser','LastLoggedUserDomain') -Req
  $users=CsvAnyProjected @('AD_Users_AllDomains.csv','AD_Users_AllDomains_Brut.csv') @('SamAccountName','Name','UserPrincipalName','EmailAddress','OrganizationalUnit','CanonicalName','DistinguishedName')
  $licenses=CsvAnyProjected @('M365_Licenses_Users.csv') @('User principal name','UserPrincipalName','primarysmtp')
  $local=CsvAnyProjected @('Intune_Devices_LocalSystem.csv','M365_Inventory_Device_LocalSystem.csv') @('AzureADDeviceId','DeviceId','DeviceName','SecureBootStatus','FirmwareType','BIOSDate','Last Reboot Date','LastRebootDate','LastBootUpTime')
  $wu=CsvAnyProjected @('Intune_WindowsUpdate_Status.csv','M365_WindowsUpdate_Status_From_Intune.csv') @('PolicyId','DeviceId','ReadinessGraphId','NormalizedDeviceName','DeviceName','ExportDateTime','ReadinessExportDateTime','RunId','ReadinessRunId','AggregateState_loc','AggregateState','CurrentDeviceUpdateStatus_loc','CurrentDeviceUpdateStatus','LatestAlertMessage_loc','LatestAlertMessage','BlockingReason','UpgradeEligibilityLabel','UpgradeEligibility','RiskBucket') -Req
  $ready=CsvAnyProjected @('Intune_Devices_UpgradeEligibility.csv') @('TenantKey','NormalizedDeviceName','DeviceName','GraphId','AzureADDeviceId','DeviceId','Device ID','ReadinessGraphId','ExportDateTime','UpgradeEligibilityLabel','UpgradeEligibility')
  $intune=CsvAnyProjected @('Intune_Devices_Inventory.csv','M365_Inventory_Devices.csv') @('Device name','DeviceName','displayName','Azure AD Device ID','AzureADDeviceId','Entra DeviceId','EntraDeviceId','Device ID','DeviceId','ManagedDeviceId','OS version','OSVersion','OperatingSystemVersion','Free storage','FreeStorage','FreeStorageGB','Primary user UPN','Primary user email address','UserPrincipalName','UPN','Compliance','ComplianceState','IsCompliant','Last check-in','LastSyncDateTime','PhysicalMemoryGB','Memory_GB_Number','Memory','Model','Manufacturer') -Req
  $entra=CsvAnyProjected @('M365_Entra_Devices.csv','M365_Inventory_EntraDevices.csv') @('DeviceId','DeviceId_Norm','DisplayName','DeviceName','IsPending','ApproximateLastSignInDateTime','RegistrationDateTime','HardwareId') -Req
  $bios=CsvAnyProjected @('Intune_Devices_BIOS.csv') @('AzureADDeviceId','DeviceName','BIOSReleaseDateTime','BIOSDate')
  $comp=CsvAnyProjected @('Intune_Devices_Compliance.csv') @('AzureADDeviceId','EntraObjectId','DeviceName','displayName','ComplianceState','Compliance')
  $hwConflicts=CsvAnyProjected @('M365_Entra_Devices_HardwareIdConflicts.csv') @('HardwareId','DeviceCount')

  $anchorPolicyId=T(Cfg $lc 'SourceAutopatchFeatureUpdateAnchorPolicyId' '38ad040b-08ca-41cd-bd86-5da5ef0b740e')
  $policy24Id=T(Cfg $lc 'SourceWindows1124H2PolicyId' '82e1d3e6-bbc0-4ddd-b36d-415979dadec6')
  $policy25Id=T(Cfg $lc 'SourceWindows1125H2PolicyId' '41046c77-bc66-44af-b4cd-7bbf2c7d343e')
  $targetPolicyIds=@($anchorPolicyId,$policy24Id,$policy25Id) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
  Log "Windows Update policy IDs: Anchor=$anchorPolicyId; 24H2=$policy24Id; 25H2=$policy25Id"

  $indexStopwatch=[Diagnostics.Stopwatch]::StartNew()
  function LogIndexPhase([string]$Name){
    Log ("Windows 11 index phase completed: {0}; elapsed={1:n1}s" -f $Name,$indexStopwatch.Elapsed.TotalSeconds)
    $indexStopwatch.Restart()
  }

  $adObjectGuid=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
  $adDeviceIdFromM365Map=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
  $adEntraDeviceId=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach($r in $ad){
    $key=G(FirstValue $r @('ObjectGuidNormalized','ObjectGUID'))
    if($key -and -not $adObjectGuid.ContainsKey($key)){$adObjectGuid[$key]=$r}
    $key=G(FirstValue $r @('IntuneDeviceId'))
    if($key -and -not $adDeviceIdFromM365Map.ContainsKey($key)){$adDeviceIdFromM365Map[$key]=$r}
    $key=G(FirstValue $r @('ObjectGuidNormalized','EntraDeviceIdNormalized','AzureADDeviceId','AzureADDeviceId_Norm','DeviceId_Norm'))
    if($key -and -not $adEntraDeviceId.ContainsKey($key)){$adEntraDeviceId[$key]=$r}
  }
  LogIndexPhase 'AD computers by ObjectGUID, managed device ID and Entra device ID'

  $usersBySam=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase); $usersByUpn=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach($r in $users){
    $key=([string]$r.SamAccountName).Trim(); if(!$key){$key=([string]$r.Name).Trim()}; $key=$key.TrimEnd('$')
    if($key -and -not $usersBySam.ContainsKey($key)){$usersBySam[$key]=$r}
    $key=([string]$r.UserPrincipalName).Trim(); if(!$key){$key=([string]$r.EmailAddress).Trim()}
    if($key -and -not $usersByUpn.ContainsKey($key)){$usersByUpn[$key]=$r}
  }
  LogIndexPhase 'AD users by SAM and UPN'

  $licensedUpn=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach($r in $licenses){$upn=([string]$r.'User principal name').Trim(); if(!$upn){$upn=([string]$r.UserPrincipalName).Trim()}; if(!$upn){$upn=([string]$r.primarysmtp).Trim()}; if($upn){[void]$licensedUpn.Add($upn)}}
  LogIndexPhase 'licensed users'

  $intuneAad=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase); $intuneId=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach($r in $intune){
    $key=G(FirstValue $r @('Azure AD Device ID','AzureADDeviceId','Entra DeviceId','EntraDeviceId'))
    if($key -and -not $intuneAad.ContainsKey($key)){$intuneAad[$key]=$r}
    $key=G(FirstValue $r @('Device ID','DeviceId','ManagedDeviceId'))
    if($key -and -not $intuneId.ContainsKey($key)){$intuneId[$key]=$r}
  }
  LogIndexPhase 'Intune devices by Azure AD device ID and managed device ID'

  $entraId=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase); $entraName=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase); $entraNameAll=[Collections.Generic.Dictionary[string,Collections.Generic.List[object]]]::new([StringComparer]::OrdinalIgnoreCase); $entraHardwareCounts=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach($r in $entra){
    $key=G(FirstValue $r @('DeviceId','DeviceId_Norm'))
    if($key -and -not $entraId.ContainsKey($key)){$entraId[$key]=$r}
    $key=([string]$r.DisplayName).Trim(); if(!$key){$key=([string]$r.DeviceName).Trim()}; $key=$key.TrimEnd('$')
    if($key -and -not $entraName.ContainsKey($key)){$entraName[$key]=$r}
    if($key){if(-not $entraNameAll.ContainsKey($key)){$entraNameAll[$key]=[Collections.Generic.List[object]]::new()}; [void]$entraNameAll[$key].Add($r)}
  }
  LogIndexPhase 'Entra devices by ID and name'

  foreach($r in $hwConflicts){$hid=([string]$r.HardwareId).Trim(); if($hid){$entraHardwareCounts[$hid]=Num($r.DeviceCount)}}
  LogIndexPhase 'Entra hardware conflict counts'

  $localAad=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach($r in $local){
    $key=G(FirstValue $r @('AzureADDeviceId','DeviceId'))
    if($key -and -not $localAad.ContainsKey($key)){$localAad[$key]=$r}
  }
  LogIndexPhase 'Local system rows by Azure AD device ID'

  $biosAad=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach($r in $bios){
    $key=G(FirstValue $r @('AzureADDeviceId'))
    if($key -and -not $biosAad.ContainsKey($key)){$biosAad[$key]=$r}
  }
  LogIndexPhase 'BIOS rows by Azure AD device ID'

  $compAad=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach($r in $comp){
    $key=G(FirstValue $r @('AzureADDeviceId','EntraObjectId'))
    if($key -and -not $compAad.ContainsKey($key)){$compAad[$key]=$r}
  }
  LogIndexPhase 'Compliance rows by Azure AD device ID'

  $readyIdAll=[Collections.Generic.Dictionary[string,Collections.Generic.List[object]]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach($r in $ready){
    $key=G(FirstValue $r @('GraphId','AzureADDeviceId','DeviceId','Device ID','ReadinessGraphId'))
    if($key){if(-not $readyIdAll.ContainsKey($key)){$readyIdAll[$key]=[Collections.Generic.List[object]]::new()}; [void]$readyIdAll[$key].Add($r)}
  }
  LogIndexPhase 'readiness rows by ID'

  function SetLatestWuRowFast($map,$dateMap,$runMap,[string]$key,$row,$candidateDate,[double]$candidateRun){
    if(!$key){return}
    if(!$map.ContainsKey($key)){$map[$key]=$row; $dateMap[$key]=$candidateDate; $runMap[$key]=$candidateRun; return}
    $currentDate=$dateMap[$key]
    $currentRun=$runMap[$key]
    if(($candidateDate -and (!$currentDate -or $candidateDate -gt $currentDate)) -or ($candidateDate -and $currentDate -and $candidateDate -eq $currentDate -and $candidateRun -gt $currentRun)){$map[$key]=$row; $dateMap[$key]=$candidateDate; $runMap[$key]=$candidateRun}
  }
  $wuByDevice=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase); $wuAnchorByDevice=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase); $wu24ByDevice=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase); $wu25ByDevice=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
  $wuDateByDevice=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase); $wuAnchorDateByDevice=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase); $wu24DateByDevice=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase); $wu25DateByDevice=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
  $wuRunByDevice=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase); $wuAnchorRunByDevice=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase); $wu24RunByDevice=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase); $wu25RunByDevice=[Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
  foreach($r in $wu){
    $policy=([string]$r.PolicyId).Trim()
    if($targetPolicyIds.Count -gt 0 -and $policy -notin $targetPolicyIds){continue}
    $did=G(FirstValue $r @('DeviceId','ReadinessGraphId'))
    $candidateDate=Dt($r.ExportDateTime); if(!$candidateDate){$candidateDate=Dt($r.ReadinessExportDateTime)}
    $candidateRun=Num($r.RunId); if($null -eq $candidateRun){$candidateRun=Num($r.ReadinessRunId)}; if($null -eq $candidateRun){$candidateRun=0}
    SetLatestWuRowFast $wuByDevice $wuDateByDevice $wuRunByDevice $did $r $candidateDate $candidateRun
    if($policy -eq $anchorPolicyId){
      SetLatestWuRowFast $wuAnchorByDevice $wuAnchorDateByDevice $wuAnchorRunByDevice $did $r $candidateDate $candidateRun
    } elseif($policy -eq $policy24Id){
      SetLatestWuRowFast $wu24ByDevice $wu24DateByDevice $wu24RunByDevice $did $r $candidateDate $candidateRun
    } elseif($policy -eq $policy25Id){
      SetLatestWuRowFast $wu25ByDevice $wu25DateByDevice $wu25RunByDevice $did $r $candidateDate $candidateRun
    }
  }
  LogIndexPhase 'Windows Update status latest rows by device ID'
  $unknownUserSet = @('', 'UNKNOWN', 'N/A', 'NA', '-', 'UNDETERMINED', 'SYSTEM', 'LOCAL SYSTEM')
  $today = (Get-Date).Date
  $now = Get-Date
  $issues=[Collections.Generic.List[object]]::new()

  if(($adDeviceIdFromM365Map.Count -gt 0) -or ($adEntraDeviceId.Count -gt 0)){
    foreach($i in $intune){
      $managedId=G(FirstValue $i @('Device ID','DeviceId','ManagedDeviceId'))
      $aadId=G(FirstValue $i @('Azure AD Device ID','AzureADDeviceId','Entra DeviceId','EntraDeviceId'))
      $existsInAd = ($managedId -and $adDeviceIdFromM365Map.ContainsKey($managedId)) -or ($aadId -and $adEntraDeviceId.ContainsKey($aadId))
      if($existsInAd){continue}
      $orphanId=$aadId; if(!$orphanId){$orphanId=$managedId}
      Add-IssueRow -Target $issues -IssueCode 'ISSUE-002' -Area 'Intune' -ObjectGUID_Norm $orphanId -Potential_Issue 'Device exists in Intune but not found in AD (orphan MDM device)' -IssueCategory '2.High' -PriorityScore 2 -IsBlocking $true -RecommendedAction 'Identify device owner and re-join to AD or retire from Intune' -ImpactMigration 'Orphan MDM device cannot be managed via hybrid join; migration scope unclear'
    }
  } else {
    Warn "Intune orphan check skipped: AD relationship keys IntuneDeviceId and EntraDeviceIdNormalized are not available in the input CSV."
  }
  LogIndexPhase 'Intune orphan check by managed device ID and Entra device ID'

  $analysisStopwatch=[Diagnostics.Stopwatch]::StartNew()
  $processedAdRows=0
  $totalAdRows=@($ad).Count
  Log "Windows 11 issue analysis started: $totalAdRows AD computer row(s)."

  foreach($row in $ad){
    $processedAdRows++
    if(($processedAdRows % 1000) -eq 0){Log ("Windows 11 issue analysis progress: {0}/{1} AD rows; issues={2}; elapsed={3:n1}s" -f $processedAdRows,$totalAdRows,$issues.Count,$analysisStopwatch.Elapsed.TotalSeconds)}
    $objectGuidNorm = ([string]$row.ObjectGuidNormalized).Trim(); if(!$objectGuidNorm){$objectGuidNorm=([string]$row.ObjectGUID).Trim()}
    if([string]::IsNullOrWhiteSpace($objectGuidNorm)){continue}
    $name=([string]$row.Name).Trim(); if(!$name){$name=([string]$row.SamAccountName).Trim()}; $name=$name.TrimEnd('$')
    $adDeviceIdFromM365Key=G(FirstValue $row @('IntuneDeviceId'))
    $adEntraDeviceIdKey=G(FirstValue $row @('ObjectGuidNormalized','EntraDeviceIdNormalized','AzureADDeviceId','AzureADDeviceId_Norm','DeviceId_Norm'))
    $in=$null
    if($adDeviceIdFromM365Key -and $intuneId.ContainsKey($adDeviceIdFromM365Key)){$in=$intuneId[$adDeviceIdFromM365Key]}
    elseif($adEntraDeviceIdKey -and $intuneAad.ContainsKey($adEntraDeviceIdKey)){$in=$intuneAad[$adEntraDeviceIdKey]}
    $aad=G(FirstValue $in @('Azure AD Device ID','AzureADDeviceId','Entra DeviceId','EntraDeviceId')); if(!$aad){$aad=$adEntraDeviceIdKey}
    $deviceIdFromM365=G(FirstValue $in @('Device ID','DeviceId','ManagedDeviceId')); if(!$deviceIdFromM365){$deviceIdFromM365=$adDeviceIdFromM365Key}
    $en=if($aad -and $entraId.ContainsKey($aad)){$entraId[$aad]}else{$null}
    $localSystemRow=if($aad -and $localAad.ContainsKey($aad)){$localAad[$aad]}else{$null}
    $biosRow=if($aad -and $biosAad.ContainsKey($aad)){$biosAad[$aad]}else{$null}
    $compRow=if($aad -and $compAad.ContainsKey($aad)){$compAad[$aad]}else{$null}
    $readyRows=@(); if($aad -and $readyIdAll.ContainsKey($aad)){$readyRows+=@($readyIdAll[$aad])}; if($deviceIdFromM365 -and $readyIdAll.ContainsKey($deviceIdFromM365)){$readyRows+=@($readyIdAll[$deviceIdFromM365])}
    $readyRow=LatestByDate $readyRows @('ExportDateTime','ReadinessExportDateTime')
    $wuAny=if($deviceIdFromM365 -and $wuByDevice.ContainsKey($deviceIdFromM365)){$wuByDevice[$deviceIdFromM365]}else{$null}
    $wuAnchor=if($deviceIdFromM365 -and $wuAnchorByDevice.ContainsKey($deviceIdFromM365)){$wuAnchorByDevice[$deviceIdFromM365]}else{$null}
    $wu24=if($deviceIdFromM365 -and $wu24ByDevice.ContainsKey($deviceIdFromM365)){$wu24ByDevice[$deviceIdFromM365]}else{$null}
    $wu25=if($deviceIdFromM365 -and $wu25ByDevice.ContainsKey($deviceIdFromM365)){$wu25ByDevice[$deviceIdFromM365]}else{$null}

    $friendlyOs = FriendlyOs $row $in
    $osNameUpper = $friendlyOs.ToUpperInvariant()
    $isWin10 = $osNameUpper.Contains('WINDOWS 10')
    $isWin11 = $osNameUpper.Contains('WINDOWS 11')
    $isWin1124H2 = ($friendlyOs -eq 'Windows 11 24H2')
    $isWin1125H2 = ($friendlyOs -eq 'Windows 11 25H2')
    $eligNorm = if($isWin11){'UPGRADED'}else{EligibilityLabel($readyRow.UpgradeEligibilityLabel)}
    if(!$eligNorm){$eligNorm=EligibilityLabel($readyRow.UpgradeEligibility)}
    if(!$eligNorm){$eligNorm=EligibilityLabel($readyRow.ReadinessStatus)}
    if(!$eligNorm){$eligNorm=EligibilityLabel($wuAny.UpgradeEligibilityLabel)}
    if(!$eligNorm){$eligNorm=EligibilityLabel($wuAny.UpgradeEligibility)}
    if(!$eligNorm){$eligNorm=EligibilityLabel($in.UpgradeEligibility)}
    if(!$eligNorm){$eligNorm=EligibilityLabel($in.UpgradeEligibilityLabel)}
    if(!$eligNorm){$eligNorm='UNKNOWN'}
    $isEligCapable = ($eligNorm -eq 'CAPABLE')
    $isEligNotCapable = ($eligNorm -eq 'NOT CAPABLE')
    $isEligUnknown = $eligNorm -in @('UNKNOWN','UNDETERMINED')
    $build=ParseBuild $row.operatingSystemVersion $in.'OS version'; if(!$build){$build=ParseBuild $row.OperatingSystemVersion $in.OSVersion}; if(!$build){$build=ParseBuild $row.OperatingSystemVersion $in.OperatingSystemVersion}
    $hasMinW11Base = $false; if($build -match '^\d+$'){$hasMinW11Base=([int]$build -ge 19041)}
    $freeValue=$in.'Free storage'; if(!$freeValue){$freeValue=$in.FreeStorage}; if(!$freeValue){$freeValue=$in.FreeStorageGB}
    $freeGb=FreeGb($freeValue)
    $storageRaw=if($null -eq $freeGb){''}else{[string]$freeGb}
    $isStorageUnknown=[string]::IsNullOrWhiteSpace($storageRaw)
    $isFreeStorageNotEnough=($null -ne $freeGb -and $freeGb -lt 40)

    $lastLoggedUserAdNorm = Get-UpperTrimmedString $row.LastLoggedUser; if(!$lastLoggedUserAdNorm){$lastLoggedUserAdNorm=Get-UpperTrimmedString $row.LastLoggedUserDomain}
    $primaryUser = ([string]$in.'Primary user UPN').Trim(); if(!$primaryUser){$primaryUser=([string]$in.'Primary user email address').Trim()}; if(!$primaryUser){$primaryUser=([string]$in.UserPrincipalName).Trim()}; if(!$primaryUser){$primaryUser=([string]$in.UPN).Trim()}
    if(!$lastLoggedUserAdNorm -and $primaryUser){$lastLoggedUserAdNorm=$primaryUser.ToUpperInvariant()}
    if(!$lastLoggedUserAdNorm){$lastLoggedUserAdNorm='UNKNOWN'}
    $licenseValue='?'
    if($primaryUser -and $licensedUpn.Contains($primaryUser)){$licenseValue='Licensed'}
    elseif($primaryUser -and $licenses.Count -gt 0){$licenseValue='No License (Unlicensed)'}
    $licenseAdStatus=NormalizeLicenseStatus $licenseValue
    $complianceValue=([string]$in.Compliance).Trim(); if(!$complianceValue){$complianceValue=([string]$in.ComplianceState).Trim()}; if(!$complianceValue){$complianceValue=([string]$in.IsCompliant).Trim()}
    $compValue=([string]$compRow.ComplianceState).Trim(); if(!$compValue){$compValue=([string]$compRow.Compliance).Trim()}
    $isNonCompliant = ($complianceValue.ToUpperInvariant() -in @('NONCOMPLIANT','FALSE')) -or ($compValue.ToUpperInvariant() -eq 'NONCOMPLIANT')

    $lastCheckInRaw=$in.'Last check-in'; if(!$lastCheckInRaw){$lastCheckInRaw=$in.LastSyncDateTime}
    $lastCheckInDateTime=Dt($lastCheckInRaw)
    $lastCheckInDate=if($lastCheckInDateTime){$lastCheckInDateTime.Date}else{$null}
    $intuneDaysSinceLastCheckIn=if($lastCheckInDateTime){([datetime]::Today - $lastCheckInDateTime.Date).Days}else{$null}
    $intuneCheckInRecent7d=($null -ne $intuneDaysSinceLastCheckIn) -and ($intuneDaysSinceLastCheckIn -le 7)

    $wuStateAnchorNorm = Get-UpperTrimmedString $wuAnchor.AggregateState_loc; if(!$wuStateAnchorNorm){$wuStateAnchorNorm=Get-UpperTrimmedString $wuAnchor.AggregateState}
    $wuState24Norm = Get-UpperTrimmedString $wu24.AggregateState_loc; if(!$wuState24Norm){$wuState24Norm=Get-UpperTrimmedString $wu24.AggregateState}
    $wuState25Norm = Get-UpperTrimmedString $wu25.AggregateState_loc; if(!$wuState25Norm){$wuState25Norm=Get-UpperTrimmedString $wu25.AggregateState}
    $wuStatusAnchorNorm = Get-UpperTrimmedString $wuAnchor.CurrentDeviceUpdateStatus_loc; if(!$wuStatusAnchorNorm){$wuStatusAnchorNorm=Get-UpperTrimmedString $wuAnchor.CurrentDeviceUpdateStatus}
    $wuStatus24Norm = Get-UpperTrimmedString $wu24.CurrentDeviceUpdateStatus_loc; if(!$wuStatus24Norm){$wuStatus24Norm=Get-UpperTrimmedString $wu24.CurrentDeviceUpdateStatus}
    $wuStatus25Norm = Get-UpperTrimmedString $wu25.CurrentDeviceUpdateStatus_loc; if(!$wuStatus25Norm){$wuStatus25Norm=Get-UpperTrimmedString $wu25.CurrentDeviceUpdateStatus}
    $wuAlertAnchorNorm = Get-UpperTrimmedString $wuAnchor.LatestAlertMessage_loc; if(!$wuAlertAnchorNorm){$wuAlertAnchorNorm=Get-UpperTrimmedString $wuAnchor.LatestAlertMessage}; if(!$wuAlertAnchorNorm){$wuAlertAnchorNorm=Get-UpperTrimmedString $wuAnchor.BlockingReason}
    $wuAlert24Norm = Get-UpperTrimmedString $wu24.LatestAlertMessage_loc; if(!$wuAlert24Norm){$wuAlert24Norm=Get-UpperTrimmedString $wu24.LatestAlertMessage}; if(!$wuAlert24Norm){$wuAlert24Norm=Get-UpperTrimmedString $wu24.BlockingReason}
    $wuAlert25Norm = Get-UpperTrimmedString $wu25.LatestAlertMessage_loc; if(!$wuAlert25Norm){$wuAlert25Norm=Get-UpperTrimmedString $wu25.LatestAlertMessage}; if(!$wuAlert25Norm){$wuAlert25Norm=Get-UpperTrimmedString $wu25.BlockingReason}
    $wuStateError = ($wuStateAnchorNorm -eq 'ERROR') -or ($wuState24Norm -eq 'ERROR') -or ($wuState25Norm -eq 'ERROR')
    $wuStateRollback = ($wuStateAnchorNorm -eq 'ROLLBACK INITIATED OR COMPLETED') -or ($wuState24Norm -eq 'ROLLBACK INITIATED OR COMPLETED') -or ($wuState25Norm -eq 'ROLLBACK INITIATED OR COMPLETED')
    $wuStateInProgress = ($wuStateAnchorNorm -eq 'IN PROGRESS') -or ($wuState24Norm -eq 'IN PROGRESS') -or ($wuState25Norm -eq 'IN PROGRESS')
    $wuStatusNeedsAttention = ($wuStatusAnchorNorm -eq 'NEEDS ATTENTION') -or ($wuStatus24Norm -eq 'NEEDS ATTENTION') -or ($wuStatus25Norm -eq 'NEEDS ATTENTION')
    $wuAlertMessage = $null
    foreach ($candidate in @($wuAlertAnchorNorm, $wuAlert24Norm, $wuAlert25Norm)) {
      if ($candidate -eq 'DOWNLOAD ISSUE') { $wuAlertMessage = 'Download Issue'; break }
      if ($candidate -eq 'INSUFFICIENT UPDATE CONNECTIVITY') { $wuAlertMessage = 'Insufficient Update Connectivity'; break }
    }
    $wuAttentionCause = $null; $wuAttentionHint = $null
    if ($wuStatusNeedsAttention -and $wuAlertMessage -eq 'Download Issue') {$wuAttentionCause='CONTENT_DOWNLOAD'; $wuAttentionHint='Check Delivery Optimization / cache / proxy/CDN. Validate WUfB content reachability.'}
    elseif ($wuStatusNeedsAttention -and $wuAlertMessage -eq 'Insufficient Update Connectivity') {$wuAttentionCause='CONNECTIVITY'; $wuAttentionHint='Check VPN/proxy/TLS inspection. Validate Microsoft update endpoints reachability.'}
    elseif ($wuStatusNeedsAttention) {$wuAttentionCause='NEEDS_ATTENTION_OTHER'; $wuAttentionHint='Check disk space, reboot pending, WU service health, update ring assignment.'}
    $wuDate=Dt($wuAny.ExportDateTime)
    $wuDaysSinceLastExport=if($wuDate){([datetime]::Today - $wuDate.Date).Days}else{$null}
    $wuNoPolicySignal = [string]::IsNullOrWhiteSpace($wuStateAnchorNorm) -and [string]::IsNullOrWhiteSpace($wuState24Norm) -and [string]::IsNullOrWhiteSpace($wuState25Norm)

    $existsInIntune = $null -ne $in
    $enabled = B($row.Enabled)
    $entraExistsByGuid = $null -ne $en
    $entraNameMatchesCountByName = if($name -and $entraNameAll.ContainsKey($name)){$entraNameAll[$name].Count}else{0}
    $entraRegisteredPending = B($en.IsPending)
    $memoryValue=$in.PhysicalMemoryGB; if(!$memoryValue){$memoryValue=$in.Memory_GB_Number}; if(!$memoryValue){$memoryValue=$in.Memory}
    $memoryGb = Num($memoryValue)
    $azureEntraDaysSinceLastSignIn = DaysSince($en.ApproximateLastSignInDateTime)
    $azureEntraRegistrationDateTime = Dt($en.RegistrationDateTime)
    $hardwareId=([string]$en.HardwareId).Trim()
    $hardwareIdDeviceCount=$null
    if($hardwareId){if($entraHardwareCounts.ContainsKey($hardwareId)){$hardwareIdDeviceCount=[int]$entraHardwareCounts[$hardwareId]}else{$hardwareIdDeviceCount=1}}
    $isWindows10Ltsc = (([string]$row.OperatingSystem).ToUpperInvariant() -match 'LTSC|LTSB') -or ($osNameUpper -match 'LTSC|LTSB')
    $lastLogonDate = Dt($row.LastLogonDate)
    if($lastLogonDate){$lastLogonDate=$lastLogonDate.Date}
    $lastRebootRaw = $row.LastRebootDateTime; if(!$lastRebootRaw){$lastRebootRaw=$row.LastRebootDate}; if(!$lastRebootRaw){$lastRebootRaw=$row.LastBootUpTime}
    if([string]::IsNullOrWhiteSpace(([string]$lastRebootRaw).Trim())){$lastRebootRaw = $localSystemRow.'Last Reboot Date'; if(!$lastRebootRaw){$lastRebootRaw=$localSystemRow.LastRebootDate}; if(!$lastRebootRaw){$lastRebootRaw=$localSystemRow.LastBootUpTime}}
    $lastRebootDateNorm = Dt($lastRebootRaw)
    if($lastRebootDateNorm){$lastRebootDateNorm=$lastRebootDateNorm.Date}
    $daysSinceLastReboot = if($lastRebootDateNorm){([datetime]::Today - $lastRebootDateNorm.Date).Days}else{$null}
    $organizationalUnit = ([string]$row.OrganizationalUnit).Trim(); if(!$organizationalUnit){$organizationalUnit=([string]$row.CanonicalName).Trim()}; if(!$organizationalUnit){$organizationalUnit=([string]$row.DistinguishedName).Trim()}
    $userRow=if($primaryUser -and $usersByUpn.ContainsKey($primaryUser)){$usersByUpn[$primaryUser]}else{$null}
    $organizationalUnitUser = ([string]$userRow.OrganizationalUnit).Trim(); if(!$organizationalUnitUser){$organizationalUnitUser=([string]$userRow.CanonicalName).Trim()}; if(!$organizationalUnitUser){$organizationalUnitUser=([string]$userRow.DistinguishedName).Trim()}
    $userOuMismatch = (-not [string]::IsNullOrWhiteSpace($organizationalUnit)) -and (-not [string]::IsNullOrWhiteSpace($organizationalUnitUser)) -and ($organizationalUnit -ne $organizationalUnitUser)
    $sbStatus = Get-UpperTrimmedString $localSystemRow.SecureBootStatus
    $sbFirmwareType = Get-UpperTrimmedString $localSystemRow.FirmwareType
    $sbHasData = -not [string]::IsNullOrWhiteSpace(([string]$localSystemRow.SecureBootStatus).Trim())
    $sbIsLegacyBios = ($sbFirmwareType -eq 'BIOS')
    $sbNotSupported = ($sbStatus -eq 'NOTSUPPORTED')
    $sbDisabled = ($sbStatus -eq 'DISABLED')
    $sbError = ($sbStatus -eq 'ERROR')
    $biosDateRaw = ([string]$localSystemRow.BIOSDate).Trim()
    if(!$biosDateRaw){$biosDateRaw=([string]$biosRow.BIOSReleaseDateTime).Trim()}
    $biosDateParsed = if([string]::IsNullOrWhiteSpace($biosDateRaw)){$null}else{Dt $biosDateRaw}
    if($biosDateParsed){$biosDateParsed=$biosDateParsed.Date}
    $biosIsOld = ($null -ne $biosDateParsed) -and ($biosDateParsed -lt (Get-Date '2019-01-01').Date) -and (-not $sbIsLegacyBios)
    $correlationStatusRobust = if($entraExistsByGuid){if($azureEntraRegistrationDateTime){'Found (correlated, registered)'}else{'Found (correlated, missing registration date)'}}elseif($entraNameMatchesCountByName -gt 0){'Not found in Entra'}else{'Not found in Entra'}
        # Critical
        if ((-not ($unknownUserSet -contains $lastLoggedUserAdNorm)) -and ($licenseAdStatus -eq "USER_NO_LICENSE")) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-001" -Area "Licensing" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "No M365 license assigned (assigned user)" -IssueCategory "1.Critical" -PriorityScore 1 -IsBlocking $true -RecommendedAction "Assign the appropriate M365 license to the assigned user" -ImpactMigration "Blocks migration of managed scenarios depending on M365 licensing (Intune/MDM)"
        }

        if (($unknownUserSet -contains $lastLoggedUserAdNorm) -and ($licenseAdStatus -eq "USER_NO_LICENSE")) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-001B" -Area "Licensing" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Assigned user is Unknown and has no M365 license" -IssueCategory "1.Critical" -PriorityScore 1 -IsBlocking $true -RecommendedAction "Identify the true last interactive user (logon audit/Entra sign-ins), then assign the appropriate M365 license" -ImpactMigration "Blocks migration planning (ownership unclear) and any managed scenarios depending on M365 licensing (Intune/MDM)"
        }

        if ($isWin10 -and $isEligNotCapable) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-005" -Area "OS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Device not eligible for Windows 11" -IssueCategory "1.Critical" -PriorityScore 1 -IsBlocking $true -RecommendedAction "Review hardware compatibility for Windows 11" -ImpactMigration "Migration impossible without hardware upgrade"
        }

        if ((-not $entraExistsByGuid) -and ($entraNameMatchesCountByName -eq 0)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-019" -Area "Entra" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Device not found in Azure Entra ID" -IssueCategory "1.Critical" -PriorityScore 1 -IsBlocking $true -RecommendedAction "Investigate why the device is missing from Azure Entra (sync issue?)" -ImpactMigration "Device missing in Azure; migration cannot proceed"
        }

        if ($isWin10 -and (-not $isEligNotCapable) -and ($wuStateError -or $wuStateRollback)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-062" -Area "Windows Update" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Windows Update failed or rollback detected during Windows 11 upgrade attempt" -IssueCategory "1.Critical" -PriorityScore 1 -IsBlocking $true -RecommendedAction "Investigate Windows Update failure (logs, connectivity, compatibility), remediate and retry" -ImpactMigration "Windows 11 upgrade cannot complete until update failure or rollback is resolved"
        }

        if ($isWin10 -and (-not $isEligNotCapable) -and $sbHasData -and $sbIsLegacyBios) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-070" -Area "BIOS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Legacy BIOS firmware detected - Windows 11 requires UEFI" -IssueCategory "1.Critical" -PriorityScore 1 -IsBlocking $true -RecommendedAction "Verify if UEFI mode is available in firmware settings; if legacy BIOS only, plan hardware refresh" -ImpactMigration "Windows 11 requires UEFI firmware - legacy BIOS devices cannot be upgraded"
        }

        if ($isWin10 -and (-not $isEligNotCapable) -and $sbHasData -and $sbNotSupported) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-070B" -Area "BIOS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Secure Boot not supported by hardware - Windows 11 requires Secure Boot capability" -IssueCategory "1.Critical" -PriorityScore 1 -IsBlocking $true -RecommendedAction "Validate hardware compatibility; device likely requires hardware refresh for Windows 11" -ImpactMigration "Windows 11 requires Secure Boot support - hardware without this capability cannot be upgraded"
        }

        # High
        if ($isFreeStorageNotEnough -and $isWin10 -and (-not $isEligNotCapable)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-004" -Area "Storage" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Free storage not sufficient for Win11 update" -IssueCategory "2.High" -PriorityScore 2 -IsBlocking $true -RecommendedAction "Free up disk space to meet Windows 11 minimum requirement (>= 40 GB free)" -ImpactMigration "Upgrade will fail due to insufficient disk space"
        }

        if ($isWin10 -and (-not $hasMinW11Base) -and (-not $isEligNotCapable)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-007A" -Area "OS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Windows 10 below minimum required version for Windows 11 update (needs 2004+)" -IssueCategory "2.High" -PriorityScore 2 -IsBlocking $true -RecommendedAction "Update Windows 10 to minimum baseline (2004/20H2/21H2/22H2) before Windows 11 upgrade" -ImpactMigration "Device is on a Windows 10 build older than 2004; upgrade to at least 2004 before Windows 11 rollout"
        }

        if ($entraRegisteredPending) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-032" -Area "Entra" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Azure Entra registered pending (no last sign-in)" -IssueCategory "2.High" -PriorityScore 2 -IsBlocking $true -RecommendedAction "Trigger a device sign-in and validate hybrid join status" -ImpactMigration "No last sign-in; cloud policies and upgrade flows may not apply until the device signs in"
        }

        if ($isWin10 -and $isEligUnknown -and (-not $existsInIntune) -and (-not $isEligNotCapable)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-052" -Area "OS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Windows 10 device has unknown eligibility and is not enrolled in Intune" -IssueCategory "2.High" -PriorityScore 2 -IsBlocking $true -RecommendedAction "Enroll device to Intune or complete eligibility assessment before scheduling upgrade" -ImpactMigration "Potential blocker: eligibility and management status unclear"
        }

        if (($null -ne $daysSinceLastReboot) -and ($daysSinceLastReboot -ge 30) -and (-not $isEligNotCapable)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-055" -Area "OS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "No reboot in the last 30 days" -IssueCategory "2.High" -PriorityScore 2 -IsBlocking $false -RecommendedAction "Force reboot immediately before scheduling any upgrade or feature update" -ImpactMigration "High risk: extended uptime significantly increases upgrade failure rate and policy application issues"
        }


        if ($isWin10 -and (-not $isEligNotCapable) -and (-not $wuStateError) -and (-not $wuStateRollback) -and $wuStatusNeedsAttention) {
            $potential = "Windows Update requires attention"
            if ($wuAlertMessage) { $potential += (" - {0}" -f $wuAlertMessage) }
            if ($wuAttentionCause) { $potential += (" ({0})" -f $wuAttentionCause) }
            if ([string]::IsNullOrWhiteSpace($wuAttentionHint)) { $wuAttentionHint = "Resolve update attention state (connectivity/content/disk), then retry deployment" }

            Add-IssueRow -Target $issues -IssueCode "ISSUE-063" -Area "Windows Update" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue $potential -IssueCategory "2.High" -PriorityScore 2 -IsBlocking $true -RecommendedAction $wuAttentionHint -ImpactMigration "Upgrade is blocked until Windows Update attention state is resolved"
        }

        if ($isWin10 -and (-not $isEligNotCapable) -and (-not $wuStateError) -and (-not $wuStateRollback) -and (-not $wuStatusNeedsAttention) -and $wuStateInProgress -and $intuneCheckInRecent7d -and ($null -ne $wuDaysSinceLastExport) -and ($wuDaysSinceLastExport -ge 5)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-066" -Area "Windows Update" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Windows Update deployment appears stuck in progress (>= 5 days without completion)" -IssueCategory "2.High" -PriorityScore 2 -IsBlocking $true -RecommendedAction "Investigate stalled deployment (WU service health, content download, reboot pending). Consider retry or redeploy." -ImpactMigration "Upgrade likely stalled; migration completion is at risk until progress resumes or is remediated"
        }

        if ($isWin10 -and (-not $isEligNotCapable) -and (-not $wuStateError) -and (-not $wuStateRollback) -and (-not $wuStatusNeedsAttention) -and $wuStateInProgress -and (($null -eq $intuneDaysSinceLastCheckIn) -or ($intuneDaysSinceLastCheckIn -gt 7))) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-067" -Area "Windows Update" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Windows 11 upgrade in progress but device is not checking in recently" -IssueCategory "2.High" -PriorityScore 2 -IsBlocking $true -RecommendedAction "Bring device online and force Intune check-in; validate MDM channel/VPN connectivity" -ImpactMigration "Upgrade progress cannot be confirmed or completed while device is offline/not reporting"
        }

        if ($isWin10 -and (-not $isEligNotCapable) -and $sbHasData -and $sbDisabled) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-072" -Area "BIOS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Secure Boot is disabled - must be enabled before Windows 11 upgrade" -IssueCategory "2.High" -PriorityScore 2 -IsBlocking $true -RecommendedAction "Enable Secure Boot in UEFI firmware settings; verify TPM 2.0 is also active" -ImpactMigration "Windows 11 upgrade will fail or be blocked until Secure Boot is enabled"
        }

        # Medium
        if ($isWin10 -and (-not $isEligNotCapable) -and $sbHasData -and $sbError) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-073" -Area "BIOS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Secure Boot status could not be determined (Platform Script returned Error)" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Investigate Platform Script execution failure on this device; re-run Detect-DeviceSystemInfo and review result" -ImpactMigration "Advisory: Secure Boot status unknown due to script error; validate before scheduling upgrade"
        }

        if ($isWin10 -and $isStorageUnknown -and (-not $isEligNotCapable)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-039" -Area "Storage" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Unknown free storage status for Windows 11 update" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Verify disk free space reporting and update inventory before scheduling upgrade" -ImpactMigration "Storage status unknown; upgrade scheduling should be deferred until disk space is confirmed"
        }

        if ($isWin10 -and $hasMinW11Base -and (-not $osNameUpper.Contains("22H2")) -and (-not $isEligNotCapable)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-007" -Area "OS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Windows 10 not at 22H2 (below recommended baseline)" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $true -RecommendedAction "Update to Windows 10 22H2" -ImpactMigration "Upgrade readiness: standardize on Windows 10 22H2 before Windows 11 rollout to reduce failure rate and support variance"
        }

        if ($isWin10 -and $isEligUnknown) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-008" -Area "OS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Unknown device eligibility for Windows 11" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Run hardware readiness assessment (PC Health Check or Intune eligibility report)" -ImpactMigration "Migration status unclear; requires eligibility check"
        }

        if ((-not $existsInIntune) -and $entraExistsByGuid -and (-not $entraRegisteredPending)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-009" -Area "Intune" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Device registered in Azure but not enrolled in Intune" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Check why the device is not enrolled in Intune (GPO not applied?)" -ImpactMigration "Cannot upgrade via Intune until enrollment is done"
        }

        if ($existsInIntune -and ($null -ne $intuneDaysSinceLastCheckIn) -and ($intuneDaysSinceLastCheckIn -gt 7)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-010" -Area "Intune" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Intune enrolled but no check-in in the last 7 days" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Bring device online and force Intune sync; validate MDM channel and VPN connectivity" -ImpactMigration "Stale MDM check-in; policies, apps and upgrade rings may not apply reliably"
        }

        if ($isNonCompliant) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-017" -Area "Compliance" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Device not compliant with Intune policies" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $true -RecommendedAction "Review compliance policies and remediate" -ImpactMigration "Upgrade blocked until compliance is restored"
        }

        if (($null -ne $memoryGb) -and ($memoryGb -lt 8) -and (-not $isEligNotCapable)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-018" -Area "Hardware" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Device has less than 8 GB RAM" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Upgrade RAM to at least 8 GB (enterprise performance guideline)" -ImpactMigration "Not blocked: Microsoft minimum is 4 GB RAM. Recommended to reach >= 8 GB for performance and future workloads."
        }

        if ($entraExistsByGuid -and ($null -eq $azureEntraRegistrationDateTime)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-022" -Area "Entra" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Azure Entra registration date missing" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Verify hybrid join (ServerAd trust), re-register if needed" -ImpactMigration "Hybrid-joined device seems unregistered or not matched; may impact cloud management"
        }

        if ($isWin11 -and (-not $isWin1124H2) -and (-not $isWin1125H2)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-025" -Area "OS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Windows 11 device not yet on latest feature update (24H2 or 25H2)" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Schedule feature update to Windows 11 24H2 or 25H2 via Intune update ring" -ImpactMigration "Device is on Windows 11 but not on the target version; standardize before end of support"
        }

        if ($entraExistsByGuid -and (-not $isWin11) -and ($null -ne $azureEntraDaysSinceLastSignIn) -and ($azureEntraDaysSinceLastSignIn -gt 30)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-026" -Area "Entra" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Not on Windows 11: Azure Entra last sign-in older than 30 days" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Investigate stalled upgrade; ensure recent sign-in and policy/app flow" -ImpactMigration "Upgrade likely stalled due to lack of recent cloud sign-in; investigate policy/app delivery"
        }

        if ($entraExistsByGuid -and $isWin11 -and (-not $isWin1124H2) -and (-not $isWin1125H2) -and ($null -ne $azureEntraDaysSinceLastSignIn) -and ($azureEntraDaysSinceLastSignIn -gt 30)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-026B" -Area "Entra" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "On Windows 11 but not 24H2/25H2: Azure Entra last sign-in older than 30 days" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Investigate stalled feature update; ensure recent Entra sign-in and update ring assignment" -ImpactMigration "Feature update to 24H2/25H2 likely stalled due to lack of recent cloud sign-in"
        }

        if ($entraExistsByGuid -and ($null -ne $hardwareIdDeviceCount) -and ($hardwareIdDeviceCount -ge 2)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-027" -Area "Entra" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Azure Entra duplicate HardwareId: multiple devices share the same ID" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Identify duplicate Azure Entra device objects and consolidate; re-register device if needed" -ImpactMigration "Duplicate HardwareId can cause conflicting policies, reporting and upgrade targeting"
        }

        if ($entraExistsByGuid -and ($null -eq $hardwareIdDeviceCount)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-028" -Area "Entra" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Azure Entra HardwareId missing" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Re-evaluate hybrid join; ensure Hardware Hash/ID is properly reported" -ImpactMigration "Missing HardwareId may limit device correlation and reliable upgrade management"
        }

        if ($entraExistsByGuid -and (-not $isWin11) -and ($null -ne $hardwareIdDeviceCount) -and ($hardwareIdDeviceCount -ge 2)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-030" -Area "Entra" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Azure Entra duplicate HardwareId and device not on Windows 11" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Remediate HardwareId duplicates; re-trigger enrollment and upgrade" -ImpactMigration "Duplicate HardwareId likely prevents consistent policy/app delivery needed to complete Windows 11 upgrade"
        }

        if ($entraExistsByGuid -and $isWin11 -and (-not $isWin1124H2) -and (-not $isWin1125H2) -and ($null -ne $hardwareIdDeviceCount) -and ($hardwareIdDeviceCount -ge 2)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-030B" -Area "Entra" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Azure Entra duplicate HardwareId and device not yet on 24H2/25H2" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Remediate HardwareId duplicates; re-trigger feature update ring assignment" -ImpactMigration "Duplicate HardwareId likely prevents consistent policy delivery needed to complete 24H2/25H2 feature update"
        }
        if (($entraNameMatchesCountByName -gt 0) -and ($correlationStatusRobust -eq "Not found in Entra") -and (-not $entraRegisteredPending)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-031" -Area "Entra" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Azure Entra correlation failed" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Normalize device names or adjust matching keys; verify hybrid join chain" -ImpactMigration "Not blocked: improve correlation to ensure accurate policy/reporting"
        }

        if ($isWin10 -and $isWindows10Ltsc -and $hasMinW11Base -and (-not $isEligNotCapable)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-041" -Area "OS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Device running Windows 10 LTSC" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Plan dedicated servicing path for Windows 10 LTSC device (in-place upgrade or hardware refresh)" -ImpactMigration "Windows 10 LTSC device: verify support policy and define the appropriate migration or replacement path before Windows 11 rollout"
        }

        if ($entraExistsByGuid -and ($entraNameMatchesCountByName -gt 1) -and (-not $entraRegisteredPending)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-047" -Area "Entra" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Azure Entra duplicate device names" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Normalize device naming and keep a single active Entra device record" -ImpactMigration "Name duplicates can confuse targeting/policies and degrade reporting quality"
        }

        if ($existsInIntune -and ($null -eq $lastCheckInDateTime)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-049" -Area "Intune" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Intune last check-in is missing" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Trigger/check device check-in; ensure MDM channel health" -ImpactMigration "Without recent check-in, policies/apps/upgrade rings may not apply reliably"
        }

        if (($null -ne $daysSinceLastReboot) -and ($daysSinceLastReboot -ge 10) -and $isWin10 -and (-not $isEligNotCapable)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-054" -Area "OS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "No reboot in the last 10 days" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Schedule a reboot before upgrade to ensure pending updates/policies are applied" -ImpactMigration "Advisory: stale uptime can cause upgrade failures or incomplete policy/applications state"
        }

        if (($null -ne $daysSinceLastReboot) -and ($daysSinceLastReboot -ge 10) -and $isWin11 -and (-not $isWin1124H2) -and (-not $isWin1125H2)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-054B" -Area "OS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "No reboot in the last 10 days (Windows 11 device not yet on 24H2/25H2)" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Schedule a reboot before feature update to ensure pending policies are applied" -ImpactMigration "Advisory: stale uptime can cause feature update failures or incomplete policy state"
        }



        if ($isWin10 -and (-not $isEligNotCapable) -and (-not $wuStateError) -and (-not $wuStateRollback) -and (-not $wuStatusNeedsAttention) -and $wuStateInProgress -and $intuneCheckInRecent7d -and (($null -eq $wuDaysSinceLastExport) -or ($wuDaysSinceLastExport -lt 5))) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-064" -Area "Windows Update" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Windows Update / Windows 11 deployment in progress" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $false -RecommendedAction "Monitor progression; investigate if status remains unchanged for multiple days" -ImpactMigration "Not blocked: device is currently upgrading"
        }

        if ($isWin10 -and $isEligCapable -and $existsInIntune -and (-not $isEligNotCapable) -and $wuNoPolicySignal) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-069" -Area "Windows Update" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Windows 10 capable device enrolled in Intune but no Windows Update policy signal detected" -IssueCategory "3.Medium" -PriorityScore 3 -IsBlocking $true -RecommendedAction "Assign a Windows Update for Business ring or feature update policy in Intune" -ImpactMigration "Device cannot be upgraded via WUfB without an active update policy assignment"
        }

        # Low
        if ($licenseAdStatus -eq "LICENSE_SIGNAL_MISSING") {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-001C" -Area "Licensing" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "M365 license signal missing for assigned user" -IssueCategory "4.Low" -PriorityScore 4 -IsBlocking $false -RecommendedAction "Fix the license enrichment pipeline for LicenseFromM365_Users_From_AD" -ImpactMigration "Advisory: licensing status cannot be assessed reliably until signal is available"
        }

        if (-not $enabled) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-006" -Area "Directory" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Account disabled" -IssueCategory "4.Low" -PriorityScore 4 -IsBlocking $false -RecommendedAction "Review device account status" -ImpactMigration "Device disabled; migration not applicable"
        }

        if (($null -ne $lastLogonDate) -and ($lastLogonDate -le $today.AddDays(-56))) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-011-AD" -Area "Directory" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Device inactive for more than 56 days (LastLogon AD)" -IssueCategory "4.Low" -PriorityScore 4 -IsBlocking $false -RecommendedAction "Review device activity in AD; retire or decommission if obsolete" -ImpactMigration "Device inactive in AD; migration likely unnecessary"
        }

        if (($null -ne $lastCheckInDate) -and ($lastCheckInDate -le $today.AddDays(-56))) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-011-M365" -Area "Intune" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Device inactive for more than 56 days (LastcheckIn Intune)" -IssueCategory "4.Low" -PriorityScore 4 -IsBlocking $false -RecommendedAction "Review device activity in Intune; retire or decommission if obsolete" -ImpactMigration "Device inactive in Intune; migration likely unnecessary"
        }

        if ($entraExistsByGuid -and ($null -ne $azureEntraDaysSinceLastSignIn) -and ($azureEntraDaysSinceLastSignIn -ge 90)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-024" -Area "Entra" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Azure Entra stale: last sign-in older than 90 days" -IssueCategory "4.Low" -PriorityScore 4 -IsBlocking $false -RecommendedAction "Review device usage; retire or remediate if still required" -ImpactMigration "Cloud device appears inactive; migration may be unnecessary or needs re-engagement"
        }

        if ($isWin10 -and $osNameUpper.Contains("22H2") -and $isEligCapable -and $existsInIntune -and (-not $isNonCompliant) -and (-not $isStorageUnknown) -and (-not $isFreeStorageNotEnough) -and $entraExistsByGuid -and (($null -eq $hardwareIdDeviceCount) -or ($hardwareIdDeviceCount -lt 2)) -and (-not $entraRegisteredPending)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-038" -Area "OS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Windows 10 device fully capable for Windows 11 but not yet upgraded" -IssueCategory "4.Low" -PriorityScore 4 -IsBlocking $false -RecommendedAction "Prioritize this device for early migration to latest Windows 11" -ImpactMigration "Ideal candidate for migration; no technical blockers detected"
        }

        if ($userOuMismatch) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-040" -Area "Directory" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Device OrganizationalUnit does not match user OrganizationalUnit" -IssueCategory "4.Low" -PriorityScore 4 -IsBlocking $false -RecommendedAction "Verify user's site assignment and ensure device is joined in the correct Organizational Unit" -ImpactMigration "User and device belong to different sites, which may impact support model, policy assignment, or migration grouping"
        }

        if (($null -eq $memoryGb) -and (-not $isEligNotCapable)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-042" -Area "Hardware" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "RAM capacity unknown" -IssueCategory "4.Low" -PriorityScore 4 -IsBlocking $false -RecommendedAction "Fix hardware inventory to report RAM before scheduling upgrade" -ImpactMigration "Advisory: inventory incomplete; validate device capacity"
        }

        if ([string]::IsNullOrWhiteSpace($friendlyOs)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-043" -Area "OS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Operating system identification missing" -IssueCategory "4.Low" -PriorityScore 4 -IsBlocking $false -RecommendedAction "Fix OS inventory reporting to identify Windows version/build" -ImpactMigration "Advisory: OS not identified; cannot determine eligibility or baseline"
        }

        if ($entraExistsByGuid -and ($null -eq $azureEntraDaysSinceLastSignIn)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-044" -Area "Entra" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Azure Entra last sign-in missing" -IssueCategory "4.Low" -PriorityScore 4 -IsBlocking $false -RecommendedAction "Validate Entra sign-in logs ingestion and device registration status" -ImpactMigration "Advisory: missing last sign-in; use with caution for staleness decisions"
        }

        if ($isWin11 -and ($isEligNotCapable -or $isEligUnknown)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-051" -Area "OS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Inconsistent inventory: Windows 11 device flagged Not Capable or Unknown eligibility" -IssueCategory "4.Low" -PriorityScore 4 -IsBlocking $false -RecommendedAction "Re-evaluate eligibility signals and correct data pipeline (TPM/SB/CPU checks)" -ImpactMigration "Advisory: data quality issue; eligibility flags should reflect actual state"
        }

        if (($null -eq $lastRebootDateNorm) -and (-not $isEligNotCapable)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-056" -Area "OS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Missing last reboot date" -IssueCategory "4.Low" -PriorityScore 4 -IsBlocking $false -RecommendedAction "Fix inventory pipeline to reliably collect the last reboot time" -ImpactMigration "Advisory: missing signal can hide stale uptime risks before upgrade"
        }


        if (($null -ne $wuDaysSinceLastExport) -and ($wuDaysSinceLastExport -ge 14) -and (-not $isEligNotCapable)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-068" -Area "Windows Update" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "Windows Update data is stale (>= 14 days since last export)" -IssueCategory "4.Low" -PriorityScore 4 -IsBlocking $false -RecommendedAction "Refresh Windows Update status export from Intune; validate data pipeline health" -ImpactMigration "Advisory: WU-based issues (ISSUE-062/063/064/066/067) may not reflect current device state"
        }

        if ($isWin10 -and (-not $isEligNotCapable) -and $existsInIntune -and (-not $sbHasData)) {
            Add-IssueRow -Target $issues -IssueCode "ISSUE-071" -Area "BIOS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue "No Secure Boot / Firmware data available (Platform Script not yet executed)" -IssueCategory "4.Low" -PriorityScore 4 -IsBlocking $false -RecommendedAction "Ensure the Detect-DeviceSystemInfo Platform Script is assigned and has run on this device (allow 24-48h after deployment)" -ImpactMigration "Advisory: cannot assess BIOS/SecureBoot readiness until Platform Script data is available"
        }

        if ($isWin10 -and (-not $isEligNotCapable) -and $sbHasData -and $biosIsOld) {
            $potential = "BIOS firmware older than 01/01/2019 - update recommended before Windows 11 upgrade"
            if ($null -ne $biosDateParsed) {
                $potential += (" (BIOS date: {0})" -f $biosDateParsed.ToString("MM/dd/yyyy"))
            }

            Add-IssueRow -Target $issues -IssueCode "ISSUE-074" -Area "BIOS" -ObjectGUID_Norm $objectGuidNorm -Potential_Issue $potential -IssueCategory "4.Low" -PriorityScore 4 -IsBlocking $false -RecommendedAction "Update BIOS firmware to latest vendor version; verify Secure Boot and TPM 2.0 settings after update" -ImpactMigration "Advisory: outdated BIOS may cause Windows 11 compatibility issues or prevent enabling Secure Boot / TPM 2.0"
        }
    }

  Log ("Windows 11 issue analysis completed: {0}/{1} AD rows; issues={2}; elapsed={3:n1}s" -f $processedAdRows,$totalAdRows,$issues.Count,$analysisStopwatch.Elapsed.TotalSeconds)
  $files=ExportIssues -rows @($issues)
  $files += PublishWeeklyHistory -files $files
  $script:GeneratedFileCount=@($files).Count
  Log "Generated Windows 11 readiness issues: $($issues.Count) row(s)"
  if($global:EnableSharePointUpload){foreach($f in $files){try{Invoke-SmartM365SharePointCsvUpload -LocalFilePath $f|Out-Null; Log "SharePoint upload completed: $f"}catch{Warn "SharePoint upload failed for $f : $($_.Exception.Message)"}}}
  Log "$ScriptName completed successfully."
}
catch {
  $script:ErrorCount++
  $script:CompletionStatus = 'Failed'
  Write-Error $_
  throw
}
finally {
  Write-SmartM365CompletionBanner -Status $script:CompletionStatus -ScriptName $ScriptName -StartedAt $RunStartedAt -WarningCount $script:WarningCount -ErrorCount $script:ErrorCount -GeneratedCsvFiles $script:GeneratedFileCount
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCATfPgXxkAzYBHT
# wefRxE9Mv4ZtPJatDdQX1/Y+vCeur6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOftZISEM7XTjcmUlmrAtSU1/aOb92me77isRkUk8jHsMA0GCSqG
# SIb3DQEBAQUABIIBgF1WDaGsuBKE1g4i2IrM5J99uomcRcxsX+Ty1OgR4ElBoOFZ
# YA/5gCFIE+dunf/blu0kk1DlGZDDUSlM6kh/MxZGk0yVkOSgEtj7vamD279fsKJy
# DGDkwvGMfshZEOXYnxQfsj46ZWbTIbg69a7COD02/uo6awkPtIpcqXqy2THHDBqP
# eZlau9Zk0GgxxUhE/PokwtCPL+smKmYI1dH1GYUWFQSNfZvbR1X88zBMk2+Se0bz
# WwUPjSFgCt32FZMxp/oG6b6X6m1iaXidb8VaiqUwBkzkojzLz6k0VggLfPVNSlvZ
# vYhkq38NlI3CBUC9yHwl2484Ua74LmEflBN+kvh4y7OfUi00F3onDtd5t+VBcpiC
# +2ViheSncx5QFks4CcEFAKIb2+KRBqvMzakLRHqJ7L5XUqzOgb1X0g72QyfbM/Qc
# MV2lRj+MAOEKge+Cq05YoJN/pJGA30dSLiovWbjOnW7YrVtMzEEN/h1pTOKo2Fm2
# eUVjFadwOTOOZ8jU26GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAqA7xhLjfEFgtHEdqeVdGgwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA3MjAyMTAz
# MzRaMC8GCSqGSIb3DQEJBDEiBCCpj6b4vDpQu43Lz+L3M3K1r3xKkm6rPGZ3LIB7
# uKjn+zANBgkqhkiG9w0BAQEFAASCAgBBHqL3UC5+VKcvxNoHfWT7rGh+OCeUol9z
# knUdSETh/sqEyjIPZJfCz5lFX2/CkUMxHkvbziUn3Axg/1z/ZyO/urbT1DToP1CB
# xP0PDnlxOaKnTDw1RsWAk6KYgMD44D9jHZujlvV/m5f2orawOgN0ePxBJiFeNAUE
# 6yUvojUsrE3ckDL3GhpgnXZgb23v7n5Eu5iHckKN1yJBb6Fdf5n9IOgye5jzjpXX
# VaYQnLz0Qy2JshmOIATC2HkG2qlMvhcH3bhQPiT6N/vfT64jCV6olfrRFfadvoMc
# 7zOmt+vWOd9pXksYGfufXBu5lj3RnCw7eOsf/GZMKHhO7nKCAbXzIyPAiEN0Zj0r
# AhzR6jkOja6sGUvocnbL4YWBfgSo2yTt49Grh1SgkLvEV2QD9HqnH9H9cX7OEQ+R
# /ASGRs2KVEizjynukmF6QjLJPOyYiPQLFt78H9f9FJRPToxriJ5muHA9lNme3wg2
# 1gZVI03NGK0YAFzmOqyMqwKB1f3b6QorRQ9Z1R0aoiJQBL0lkKobJ/ccXKYm27Vd
# 5U+yJEzbXyQkAn+WqfZ1sEA0DbYzOzKVC9fRHDrB11hSYJG0I86K73iqjBgYIbPP
# FCd9KAnxm19ZZ22oVlblTZGD5x82CapAWLLfjaTTtUZfMzDNhx9v4n8olk7hcyNu
# Mob87zrFiA==
# SIG # End signature block
