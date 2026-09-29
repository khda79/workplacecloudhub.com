[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$HistoryPath,
    [Parameter(Mandatory)][string]$OutputPath
)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$weeks=@{}
$keys=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
Import-Csv -LiteralPath $HistoryPath | ForEach-Object {
    $row=$_; $week=[string]$row.'Week Label'
    $date=[datetime]::Parse($row.'Snapshot Date',[Globalization.CultureInfo]::InvariantCulture)
    if(-not $week -or -not $row.'User Source ID') {throw 'Missing history key'}
    if(-not $keys.Add($row.'Date Key'+'|'+$row.'User Source ID')) {throw 'Duplicate date/user history key'}
    if(-not $weeks.ContainsKey($week)) {$weeks[$week]=@{Date=$date;Enabled=0;Human=0;Covered=0;Active=0;Inactive=0}}
    $totals=$weeks[$week]
    if($date -ne $totals.Date) {throw "Multiple snapshots in week $week; choose a single complete snapshot before aggregating."}
    $totals.Enabled++
    if([bool]::Parse($row.'Confirmed Human')) {$totals.Human++}
    if($row.'Activity Evidence Status' -eq 'Observed') {$totals.Covered++}
    if($row.'Activity State' -eq 'Active <=30D') {$totals.Active++}
    if($row.'Activity State' -in @('Dormant 31-90D','Stale >90D','Observed never used')) {$totals.Inactive++}
}
if(-not $weeks.Count) {throw 'History is empty'}
$output=foreach($week in $weeks.Keys | Sort-Object) {
    $t=$weeks[$week]
    [pscustomobject][ordered]@{
        'Snapshot Date'=$t.Date.ToString('yyyy-MM-dd');'Date Key'=[int]$t.Date.ToString('yyyyMMdd');'Week Label'=$week
        'Enabled Workforce Accounts'=$t.Enabled;'Confirmed Human Users'=$t.Human;'Activity Covered Users'=$t.Covered
        'Activity Evidence Coverage'=([math]::Round($t.Covered/[double]$t.Enabled,6)).ToString('0.######',[Globalization.CultureInfo]::InvariantCulture)
        'Active Workforce'=$t.Active;'Accounts Without Activity >30D'=$t.Inactive;'Evidence Status'='Observed'
    }
}
$output | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8NoBOM
[pscustomobject]@{Weeks=$weeks.Count;Rows=$keys.Count;Output=$OutputPath}
