[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$DataRoot,
    [Parameter(Mandatory)][string]$OutputPath
)
$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
$result = [Collections.Generic.List[object]]::new()
function Add-Metric([datetime]$Date, [string]$Name, [double]$Value) {
    $result.Add([pscustomobject][ordered]@{
        'Date Key' = [int]$Date.ToString('yyyyMMdd')
        'Tenant Key' = 1
        'Environment Key' = 1
        'Metric Name' = $Name
        'Metric Value' = $Value.ToString('R',[Globalization.CultureInfo]::InvariantCulture)
        'Evidence Status' = 'Observed'
    })
}
function Read-Required([string]$Path, [char]$Delimiter = ',') {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Missing required trend source: $Path" }
    Import-Csv -LiteralPath $Path -Delimiter $Delimiter
}
function Week-Files([string]$RelativeRoot, [string]$Name) {
    $root = Join-Path $DataRoot $RelativeRoot
    if (-not (Test-Path -LiteralPath $root)) { throw "Historical source directory is missing: $root" }
    $latest = @{}
    foreach ($f in Get-ChildItem -LiteralPath $root -Recurse -File -Filter $Name) {
        $match = [regex]::Match($f.FullName,'WeeklyHistory[\\/](\d{4}-W\d{2})[\\/]')
        if (-not $match.Success) { continue }
        $key = $match.Groups[1].Value
        if (-not $latest.ContainsKey($key) -or $f.LastWriteTimeUtc -gt $latest[$key].LastWriteTimeUtc) { $latest[$key] = $f }
    }
    @($latest.Values | Sort-Object LastWriteTimeUtc)
}
function Parse-Date([string]$Value) {
    [datetime]::Parse($Value,[Globalization.CultureInfo]::InvariantCulture)
}
function Read-DailyStats([string]$Name, [string[]]$ValueFields) {
    $path = Join-Path $DataRoot ('DATA-ALL\ActiveDirectory\Inventory\' + $Name)
    $latest = @{}
    foreach ($row in Read-Required $path ';') {
        $date = Parse-Date $row.Date
        $key = @($date.ToString('yyyy-MM-dd'),$row.TenantKey,$row.OrganizationKey,$row.EnvironmentKey,$row.TenantId,$row.DomainName) -join '|'
        if (-not $latest.ContainsKey($key) -or $date -gt $latest[$key].Date) {
            $latest[$key] = @{ Date=$date; Row=$row }
        }
    }
    foreach ($day in $latest.Values | Group-Object { $_.Date.Date }) {
        $values = @{}
        foreach ($field in $ValueFields) {
            $values[$field] = [double]0
            foreach ($entry in $day.Group) {
                $values[$field] += [double]::Parse([string]$entry.Row.$field,[Globalization.CultureInfo]::InvariantCulture)
            }
        }
        [pscustomobject]@{ Date=$day.Group[0].Date.Date; Values=$values }
    }
}
foreach ($day in Read-DailyStats 'AD_Users_DailyStats.csv' @('TotalUsers')) {
    Add-Metric $day.Date '# Users' $day.Values.TotalUsers
}
foreach ($day in Read-DailyStats 'AD_Computers_DailyStats.csv' @('TotalComputers','EnabledAccounts','Windows 11 Enabled')) {
    Add-Metric $day.Date '# Devices' $day.Values.TotalComputers
    if ($day.Values.EnabledAccounts -gt 0) { Add-Metric $day.Date 'Windows 11 Adoption (%)' ($day.Values['Windows 11 Enabled'] / $day.Values.EnabledAccounts) }
}
$usageFiles = @(Week-Files 'DATA-ALL\M365\Usage' 'M365_Users_Activity.csv') + @(Get-Item -LiteralPath (Join-Path $DataRoot 'DATA-LAST\M365_Users_Activity.csv'))
$usageByDate = @{}
foreach ($file in $usageFiles | Sort-Object LastWriteTimeUtc) {
    $days = @{}
    foreach ($row in Read-Required $file.FullName) {
        $date = (Parse-Date $row.ReportRefreshDate).ToString('yyyy-MM-dd')
        if (-not $days.ContainsKey($date)) { $days[$date] = @{} }
        $upn = ([string]$row.UserPrincipalName).Trim().ToLowerInvariant()
        $flag = ([string]$row.HasAnyM365Activity).Trim().ToLowerInvariant()
        # Match the prior distinct report-date/UPN/activity grain.
        $days[$date][($upn + '|' + $flag)] = $flag -eq 'true'
    }
    foreach ($date in $days.Keys) { $usageByDate[$date] = $days[$date] }
}
foreach ($date in $usageByDate.Keys) {
    $values = @($usageByDate[$date].Values)
    if ($values.Count -gt 0) { Add-Metric (Parse-Date $date) 'M365 Active Use (30D) (%)' (@($values | Where-Object { $_ }).Count / [double]$values.Count) }
}
$counts = @{}
foreach ($spec in @(
    @{Root='DATA-ALL\Exchange\EXO'; Name='Exchange_EXO_Mailboxes_AllDomains.csv'; Hosting='Online'},
    @{Root='DATA-ALL\Exchange\OnPrem\Mailboxes'; Name='Exchange_OnPrem_Mailboxes_AllDomains.csv'; Hosting='OnPrem'}
)) {
    $files = @(Week-Files $spec.Root $spec.Name)
    $current = Get-Item -LiteralPath (Join-Path $DataRoot ('DATA-LAST\' + $spec.Name))
    foreach ($file in @($files) + @($current)) {
        $date = $file.LastWriteTime.Date
        if ($file.FullName -ne $current.FullName) { $date = $date.AddDays(-(([int]$date.DayOfWeek + 6) % 7)) }
        $key = $date.ToString('yyyy-MM-dd')
        if (-not $counts.ContainsKey($key)) { $counts[$key] = @{} }
        $counts[$key][$spec.Hosting] = (Read-Required $file.FullName | Measure-Object).Count
    }
}
foreach ($date in $counts.Keys) {
    $pair = $counts[$date]
    # One missing hosting source is unknown, never assumed zero.
    if (-not $pair.ContainsKey('Online') -or -not $pair.ContainsKey('OnPrem')) { continue }
    $total = $pair.Online + $pair.OnPrem
    if ($total -gt 0) { Add-Metric (Parse-Date $date) 'Exchange Online Adoption (%)' ($pair.Online / [double]$total) }
}
if ($result.Count -eq 0) { throw 'No observed executive trend evidence was produced.' }
$duplicates = @($result | Group-Object 'Date Key','Metric Name' | Where-Object Count -gt 1)
if ($duplicates.Count) { throw 'Duplicate executive trend keys.' }
New-Item -ItemType Directory -Path (Split-Path -Parent $OutputPath) -Force | Out-Null
$result | Sort-Object 'Date Key','Metric Name' | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding utf8NoBOM
[pscustomobject]@{Output=$OutputPath;Rows=$result.Count;Metrics=@($result.'Metric Name' | Sort-Object -Unique)}
