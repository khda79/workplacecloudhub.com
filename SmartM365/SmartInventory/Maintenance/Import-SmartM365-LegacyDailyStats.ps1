<#
.SYNOPSIS
Imports legacy AD and Exchange on-premises daily statistics into the SmartInventory daily CSVs.

.DESCRIPTION
Older collectors of the same organization produced AD_Users_Daily_Stats.csv,
AD_Computers_Daily_Stats.csv and Exchange_Daily_Stats.csv with the same columns as the
SmartInventory daily files, without the tenant columns. This script merges them into:
- DATA-ALL\ActiveDirectory\Inventory\AD_Users_DailyStats.csv
- DATA-ALL\ActiveDirectory\Inventory\AD_Computers_DailyStats.csv
- DATA-ALL\Exchange\OnPrem\MailboxReports\Exchange_OnPrem_Mailboxes_DailyStats.csv
and refreshes the matching DATA-LAST copy when it exists.

SmartInventory rows always win:
- a legacy row is imported only when SmartInventory has no row at all for that day;
- existing SmartInventory rows are never removed or reordered within a day, and their values
  are kept, except the Exchange mailbox type columns written by collector versions before 1.66
  (one 0/1 value per type instead of a count). Such a row is detected because its type columns
  do not sum to TotalMailboxCount. Its type columns are taken from the legacy row of the same
  day and domain when that row has the same TotalMailboxCount, TotalLocalMailboxCount and
  TotalRemoteMailboxCount and is itself consistent; otherwise they are left empty.

Imported legacy rows receive the tenant columns of the existing SmartInventory rows (exactly one
tenant identity is required). Non-numeric legacy values are written as empty, never guessed.
When exactly one part of Total = Enabled + Disabled (or Exchange Total = Local + Remote) is
missing in a legacy row, it is computed from the other two, but only when that identity holds on
every complete row of both files.
The merge is refused when the legacy file contains a domain that the SmartInventory file does not
know, or when the totals of the overlapping days differ by more than MaxMedianDifferencePercent.

Without -Execute the script is read-only and prints the plan. With -Execute it backs up every
file it replaces, writes each merged file atomically, refuses to replace a file that changed
during the run, and writes a CSV report. Run it where the collectors write (the server share),
outside the AD and Exchange on-premises collector windows; the next collector runs republish
the merged files to SharePoint. The script is idempotent: a second run imports nothing.

.PARAMETER DataRootPath
Tenant data root that contains DATA-ALL and DATA-LAST.

.PARAMETER LegacyFolderPath
Folder that contains the legacy daily CSV files.

.PARAMETER ExchangeDailyStatsPath
Exchange on-premises daily CSV when LocalMailboxReportCsvLogFolderPath is not the default.

.PARAMETER MaxMedianDifferencePercent
Maximum median difference of the tenant total between both sources on overlapping days. Default 5.

.PARAMETER BackupFolderPath
Backup folder. Defaults to %TEMP%\SmartM365-LegacyDailyStats_<timestamp>.

.PARAMETER ReportPath
CSV report. Defaults to %TEMP%\SmartM365-LegacyDailyStats_<timestamp>.csv.

.PARAMETER Execute
Writes the merged files. Omit for preview only.

.VERSION
1.0

.NOTES
Author: https://github.com/khda79/workplacecloudhub.com
#>
#requires -Version 7.2

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$DataRootPath,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$LegacyFolderPath,

    [string]$ExchangeDailyStatsPath = '',

    [ValidateRange(0, 100)]
    [double]$MaxMedianDifferencePercent = 5,

    [string]$BackupFolderPath = '',

    [string]$ReportPath = '',

    [switch]$Execute
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ScriptVersion = '1.0'
$startedAt = Get-Date

$tenantContextPath = Join-Path -Path $PSScriptRoot -ChildPath '..\..\Config\SmartM365-TenantContext.ps1'
if (-not (Test-Path -LiteralPath $tenantContextPath -PathType Leaf)) { throw "SmartM365 tenant context not found: $tenantContextPath" }
. $tenantContextPath
Write-SmartM365StartupBanner

$dataRoot = [IO.Path]::GetFullPath($DataRootPath).TrimEnd('\', '/')
$dataAllRoot = Join-Path -Path $dataRoot -ChildPath 'DATA-ALL'
$dataLastRoot = Join-Path -Path $dataRoot -ChildPath 'DATA-LAST'
foreach ($folder in @($dataAllRoot, $dataLastRoot)) {
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) { throw "Folder not found under DataRootPath: $folder" }
}
$legacyRoot = [IO.Path]::GetFullPath($LegacyFolderPath)
if (-not (Test-Path -LiteralPath $legacyRoot -PathType Container)) { throw "Legacy folder not found: $legacyRoot" }
$stamp = $startedAt.ToString('yyyyMMdd_HHmmss')
if ([string]::IsNullOrWhiteSpace($BackupFolderPath)) { $BackupFolderPath = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath "SmartM365-LegacyDailyStats_$stamp" }
if ([string]::IsNullOrWhiteSpace($ReportPath)) { $ReportPath = Join-Path -Path ([IO.Path]::GetTempPath()) -ChildPath "SmartM365-LegacyDailyStats_$stamp.csv" }
if ([string]::IsNullOrWhiteSpace($ExchangeDailyStatsPath)) {
    $ExchangeDailyStatsPath = Join-Path -Path $dataAllRoot -ChildPath 'Exchange\OnPrem\MailboxReports\Exchange_OnPrem_Mailboxes_DailyStats.csv'
}

$tenantColumns = @('TenantKey', 'OrganizationKey', 'EnvironmentKey', 'TenantId')
$mailboxTypeColumns = @('EquipmentMailbox', 'RemoteSharedMailbox', 'RemoteUserMailbox', 'RoomMailbox', 'SharedMailbox', 'UserMailbox')
$invariant = [Globalization.CultureInfo]::InvariantCulture
$datasets = @(
    [pscustomobject]@{ Name = 'AD_Users'; LegacyFile = 'AD_Users_Daily_Stats.csv'; Target = (Join-Path $dataAllRoot 'ActiveDirectory\Inventory\AD_Users_DailyStats.csv'); TotalColumn = 'TotalUsers'; Complements = @(, @('TotalUsers', 'EnabledUsers', 'DisabledUsers')); RepairMailboxTypes = $false }
    [pscustomobject]@{ Name = 'AD_Computers'; LegacyFile = 'AD_Computers_Daily_Stats.csv'; Target = (Join-Path $dataAllRoot 'ActiveDirectory\Inventory\AD_Computers_DailyStats.csv'); TotalColumn = 'TotalComputers'; Complements = @(, @('TotalComputers', 'EnabledAccounts', 'DisabledAccounts')); RepairMailboxTypes = $false }
    [pscustomobject]@{ Name = 'Exchange_OnPrem'; LegacyFile = 'Exchange_Daily_Stats.csv'; Target = [IO.Path]::GetFullPath($ExchangeDailyStatsPath); TotalColumn = 'TotalMailboxCount'; Complements = @(@('TotalMailboxCount', 'EnabledAccounts', 'DisabledAccounts'), @('TotalMailboxCount', 'TotalLocalMailboxCount', 'TotalRemoteMailboxCount')); RepairMailboxTypes = $true }
)

function Get-FileSha256([string]$Path) { (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash }

function Read-DailyCsv {
    param([string]$Path)
    $bytes = [IO.File]::ReadAllBytes($Path)
    $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
    $text = [Text.Encoding]::UTF8.GetString($bytes)
    if ($hasBom) { $text = $text.Substring(1) }
    $lines = @($text -split "`r?`n" | Where-Object { $_ -ne '' })
    if ($lines.Count -lt 2) { throw "CSV has no data rows: $Path" }
    $items = @($lines | ConvertFrom-Csv -Delimiter ';')
    $header = @($items[0].PSObject.Properties.Name)
    $rows = [Collections.Generic.List[object]]::new()
    $sequence = 0
    foreach ($item in $items) {
        $row = [ordered]@{}
        foreach ($column in $header) { $row[$column] = if ($null -eq $item.$column) { '' } else { [string]$item.$column } }
        $rows.Add([pscustomobject]@{ Values = $row; Sequence = $sequence; Source = 'SmartInventory' })
        $sequence++
    }
    [pscustomobject]@{ Header = $header; Rows = $rows; HasBom = $hasBom; Hash = Get-FileSha256 $Path }
}

function ConvertTo-RowDate([string]$Value) {
    $parsed = [datetime]::MinValue
    if ([datetime]::TryParseExact($Value, 'yyyy-MM-dd HH:mm:ss', $invariant, [Globalization.DateTimeStyles]::None, [ref]$parsed)) { return $parsed }
    return $null
}

function ConvertTo-Count([string]$Value) {
    $parsed = 0L
    if ([long]::TryParse($Value, [Globalization.NumberStyles]::None, $invariant, [ref]$parsed)) { return $parsed }
    return $null
}

function Test-NumericValue([string]$Column, [string]$Value) {
    if ($Column -like '*SizeGB') { return $Value -match '^\d+([.,]\d+)?$' }
    return $Value -match '^\d+$'
}

function Get-MailboxTypeVector($Values) {
    $counts = @(foreach ($column in $mailboxTypeColumns) { ConvertTo-Count $Values[$column] })
    if (@($counts | Where-Object { $null -eq $_ }).Count -gt 0) { return $null }
    $total = ConvertTo-Count $Values['TotalMailboxCount']
    [pscustomobject]@{ Counts = $counts; Sum = ($counts | Measure-Object -Sum).Sum; Total = $total; Key = ($counts -join '|') }
}

function Format-CsvLine([string[]]$Values) {
    ($Values | ForEach-Object { '"' + ([string]$_).Replace('"', '""') + '"' }) -join ';'
}

function Write-FileAtomically([string]$Path, [byte[]]$Bytes) {
    $temp = '{0}.import-{1}.tmp' -f $Path, ([guid]::NewGuid().ToString('N'))
    [IO.File]::WriteAllBytes($temp, $Bytes)
    for ($attempt = 1; ; $attempt++) {
        try { Move-Item -LiteralPath $temp -Destination $Path -Force; break }
        catch {
            if ($attempt -ge 5) { Remove-Item -LiteralPath $temp -Force -ErrorAction SilentlyContinue; throw "Cannot replace $Path : $($_.Exception.Message)" }
            Start-Sleep -Seconds ([Math]::Min(15, 2 * $attempt))
        }
    }
    $expected = [BitConverter]::ToString([Security.Cryptography.SHA256]::HashData($Bytes)).Replace('-', '')
    if ((Get-FileSha256 $Path) -ne $expected) { throw "Written file does not match the merged content: $Path" }
}

$plans = [Collections.Generic.List[object]]::new()
foreach ($dataset in $datasets) {
    $legacyPath = Join-Path -Path $legacyRoot -ChildPath $dataset.LegacyFile
    if (-not (Test-Path -LiteralPath $legacyPath -PathType Leaf)) { Write-Host ("[{0}] Legacy file not found, skipped: {1}" -f $dataset.Name, $legacyPath); continue }
    if (-not (Test-Path -LiteralPath $dataset.Target -PathType Leaf)) { throw "[$($dataset.Name)] SmartInventory daily CSV not found: $($dataset.Target)" }

    $target = Read-DailyCsv $dataset.Target
    $legacy = Read-DailyCsv $legacyPath
    $expectedLegacyHeader = @($target.Header | Select-Object -Skip $tenantColumns.Count)
    if ((@($target.Header | Select-Object -First $tenantColumns.Count) -join '|') -ne ($tenantColumns -join '|')) { throw "[$($dataset.Name)] SmartInventory CSV does not start with the tenant columns: $($dataset.Target)" }
    if (($legacy.Header -join '|') -ne ($expectedLegacyHeader -join '|')) { throw "[$($dataset.Name)] Legacy columns differ from the SmartInventory columns without tenant columns: $legacyPath" }
    $valueColumns = @($expectedLegacyHeader | Where-Object { $_ -notin @('Date', 'DomainName') })

    $identities = @($target.Rows | ForEach-Object { $values = $_.Values; @(foreach ($c in $tenantColumns) { $values[$c] }) -join '|' } | Sort-Object -Unique)
    if ($identities.Count -ne 1 -or ($identities[0].Split('|') | Where-Object { [string]::IsNullOrWhiteSpace($_) })) { throw "[$($dataset.Name)] SmartInventory CSV must hold exactly one complete tenant identity; found $($identities.Count)." }
    $identity = $identities[0].Split('|')

    $targetDays = [Collections.Generic.HashSet[string]]::new()
    $targetDomains = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($row in $target.Rows) {
        $date = ConvertTo-RowDate $row.Values['Date']
        if ($null -eq $date) { throw "[$($dataset.Name)] Invalid SmartInventory Date '$($row.Values['Date'])' in $($dataset.Target)" }
        $row | Add-Member -NotePropertyName DateValue -NotePropertyValue $date
        [void]$targetDays.Add($date.ToString('yyyy-MM-dd'))
        [void]$targetDomains.Add($row.Values['DomainName'])
    }
    $firstTargetDay = @($targetDays | Sort-Object)[0]

    # Legacy values: non-numeric text becomes empty, never a guessed number.
    $legacyInvalidDates = 0
    $blankedByColumn = [ordered]@{}
    $legacyRows = [Collections.Generic.List[object]]::new()
    foreach ($row in $legacy.Rows) {
        $date = ConvertTo-RowDate $row.Values['Date']
        if ($null -eq $date) { $legacyInvalidDates++; continue }
        foreach ($column in $expectedLegacyHeader) {
            $value = $row.Values[$column].Trim()
            if ($column -in $valueColumns -and $value -ne '' -and -not (Test-NumericValue $column $value)) {
                $blankedByColumn[$column] = 1 + $(if ($blankedByColumn.Contains($column)) { $blankedByColumn[$column] } else { 0 })
                $value = ''
            }
            $row.Values[$column] = $value
        }
        $row | Add-Member -NotePropertyName DateValue -NotePropertyValue $date
        $row.Source = 'Legacy'
        $legacyRows.Add($row)
    }

    # Total = PartA + PartB holds on every complete row of both sources: one missing part is then exact.
    $derivedByColumn = [ordered]@{}
    foreach ($rule in $dataset.Complements) {
        $total, $partA, $partB = $rule
        $violations = @(@($target.Rows) + @($legacyRows) | Where-Object {
            $t = ConvertTo-Count $_.Values[$total]; $a = ConvertTo-Count $_.Values[$partA]; $b = ConvertTo-Count $_.Values[$partB]
            $null -ne $t -and $null -ne $a -and $null -ne $b -and $t -ne ($a + $b)
        }).Count
        if ($violations -gt 0) { Write-Host ("[{0}] {1} = {2} + {3} does not hold on {4} complete row(s); missing values stay empty." -f $dataset.Name, $total, $partA, $partB, $violations); continue }
        foreach ($row in $legacyRows) {
            $t = ConvertTo-Count $row.Values[$total]; $a = ConvertTo-Count $row.Values[$partA]; $b = ConvertTo-Count $row.Values[$partB]
            if ($null -eq $t) { continue }
            foreach ($pair in @(@($partA, $a, $b), @($partB, $b, $a))) {
                $column, $missing, $known = $pair
                if ($null -eq $missing -and $null -ne $known -and $t -ge $known) {
                    $row.Values[$column] = [string]($t - $known)
                    $derivedByColumn[$column] = 1 + $(if ($derivedByColumn.Contains($column)) { $derivedByColumn[$column] } else { 0 })
                }
            }
        }
    }
    $unknownDomains = @($legacyRows | ForEach-Object { $_.Values['DomainName'] } | Sort-Object -Unique | Where-Object { -not $targetDomains.Contains($_) })
    if ($unknownDomains.Count -gt 0) { throw "[$($dataset.Name)] Legacy file holds $($unknownDomains.Count) domain(s) unknown to SmartInventory; refusing a possible cross-tenant import." }

    # Evidence that both files measure the same population: tenant total of the latest row per day and domain.
    $latestTotals = { param($rows) $byKey = @{}; foreach ($r in $rows) { $k = $r.DateValue.ToString('yyyy-MM-dd') + '|' + $r.Values['DomainName'].ToLowerInvariant(); if (-not $byKey.ContainsKey($k) -or $r.DateValue -gt $byKey[$k].DateValue) { $byKey[$k] = $r } }; $byKey }
    $targetLatest = & $latestTotals $target.Rows
    $legacyLatest = & $latestTotals $legacyRows
    $dayTotals = @{}
    foreach ($key in $targetLatest.Keys) {
        if (-not $legacyLatest.ContainsKey($key)) { continue }
        $a = ConvertTo-Count $legacyLatest[$key].Values[$dataset.TotalColumn]; $b = ConvertTo-Count $targetLatest[$key].Values[$dataset.TotalColumn]
        if ($null -eq $a -or $null -eq $b) { continue }
        $day = $key.Split('|')[0]
        if (-not $dayTotals.ContainsKey($day)) { $dayTotals[$day] = @(0L, 0L) }
        $dayTotals[$day] = @(($dayTotals[$day][0] + $a), ($dayTotals[$day][1] + $b))
    }
    $differences = @($dayTotals.Values | Where-Object { $_[1] -gt 0 } | ForEach-Object { [Math]::Abs($_[0] - $_[1]) * 100.0 / $_[1] } | Sort-Object)
    $medianDifference = if ($differences.Count -gt 0) { $differences[[int][Math]::Floor($differences.Count / 2)] } else { $null }
    $maxDifference = if ($differences.Count -gt 0) { $differences[-1] } else { $null }
    if ($null -ne $medianDifference -and $medianDifference -gt $MaxMedianDifferencePercent) { throw ("[{0}] Median {1} difference on {2} overlapping days is {3:N2}% (> {4}%); refusing the import." -f $dataset.Name, $dataset.TotalColumn, $differences.Count, $medianDifference, $MaxMedianDifferencePercent) }

    # Legacy rows only for days SmartInventory does not cover at all.
    $imported = [Collections.Generic.List[object]]::new()
    $skippedCovered = 0
    foreach ($row in $legacyRows) {
        if ($targetDays.Contains($row.DateValue.ToString('yyyy-MM-dd'))) { $skippedCovered++; continue }
        $values = [ordered]@{}
        for ($i = 0; $i -lt $tenantColumns.Count; $i++) { $values[$tenantColumns[$i]] = $identity[$i] }
        foreach ($column in $expectedLegacyHeader) { $values[$column] = $row.Values[$column] }
        $imported.Add([pscustomobject]@{ Values = $values; Sequence = $row.Sequence; Source = 'Legacy'; DateValue = $row.DateValue })
    }
    $importedDays = @($imported | ForEach-Object { $_.DateValue.ToString('yyyy-MM-dd') } | Sort-Object -Unique)

    # Exchange rows written before collector 1.66: type columns are not counts.
    $typeRepaired = 0; $typeBlanked = 0
    $before = @($target.Rows | ForEach-Object { Format-CsvLine @($_.Values.Values) })
    $repairedSequences = [Collections.Generic.HashSet[int]]::new()
    if ($dataset.RepairMailboxTypes) {
        $legacyByKey = @{}
        foreach ($row in $legacyRows) {
            $k = $row.DateValue.ToString('yyyy-MM-dd') + '|' + $row.Values['DomainName'].ToLowerInvariant()
            if (-not $legacyByKey.ContainsKey($k)) { $legacyByKey[$k] = [Collections.Generic.List[object]]::new() }
            $legacyByKey[$k].Add($row)
        }
        foreach ($row in $target.Rows) {
            $vector = Get-MailboxTypeVector $row.Values
            if ($null -eq $vector -or $null -eq $vector.Total -or $vector.Sum -eq $vector.Total) { continue }
            $k = $row.DateValue.ToString('yyyy-MM-dd') + '|' + $row.Values['DomainName'].ToLowerInvariant()
            $candidates = @()
            if ($legacyByKey.ContainsKey($k)) {
                $candidates = @($legacyByKey[$k] | Where-Object {
                    $legacyVector = Get-MailboxTypeVector $_.Values
                    $null -ne $legacyVector -and $legacyVector.Sum -eq $legacyVector.Total -and
                    $_.Values['TotalMailboxCount'] -eq $row.Values['TotalMailboxCount'] -and
                    $_.Values['TotalLocalMailboxCount'] -eq $row.Values['TotalLocalMailboxCount'] -and
                    $_.Values['TotalRemoteMailboxCount'] -eq $row.Values['TotalRemoteMailboxCount']
                } | ForEach-Object { Get-MailboxTypeVector $_.Values })
            }
            $distinct = @($candidates | ForEach-Object { $_.Key } | Sort-Object -Unique)
            if ($distinct.Count -eq 1) {
                for ($i = 0; $i -lt $mailboxTypeColumns.Count; $i++) { $row.Values[$mailboxTypeColumns[$i]] = [string]$candidates[0].Counts[$i] }
                $typeRepaired++
            }
            else {
                foreach ($column in $mailboxTypeColumns) { $row.Values[$column] = '' }
                $typeBlanked++
            }
            [void]$repairedSequences.Add($row.Sequence)
        }
    }

    # SmartInventory rows must be unchanged apart from the repaired type columns.
    foreach ($row in $target.Rows) {
        if ($repairedSequences.Contains($row.Sequence)) {
            $original = ($before[$row.Sequence] | ConvertFrom-Csv -Delimiter ';' -Header $target.Header)
            foreach ($column in $target.Header) {
                if ($column -in $mailboxTypeColumns) { continue }
                if ([string]$original.$column -ne $row.Values[$column]) { throw "[$($dataset.Name)] Internal check failed: SmartInventory value changed in column $column." }
            }
        }
        elseif ((Format-CsvLine @($row.Values.Values)) -ne $before[$row.Sequence]) { throw "[$($dataset.Name)] Internal check failed: SmartInventory row $($row.Sequence) changed." }
    }

    $merged = @(@($target.Rows) + @($imported) | Sort-Object -Property @{ Expression = 'DateValue' }, @{ Expression = { if ($_.Source -eq 'Legacy') { 0 } else { 1 } } }, @{ Expression = 'Sequence' })
    $lines = [Collections.Generic.List[string]]::new()
    $lines.Add((Format-CsvLine $target.Header))
    foreach ($row in $merged) { $lines.Add((Format-CsvLine @(foreach ($c in $target.Header) { $row.Values[$c] }))) }
    $text = ($lines -join "`r`n") + "`r`n"
    $encoding = [Text.UTF8Encoding]::new($target.HasBom)
    $preamble = $encoding.GetPreamble()
    $body = $encoding.GetBytes($text)
    $bytes = [byte[]]::new($preamble.Length + $body.Length)
    $preamble.CopyTo($bytes, 0)
    $body.CopyTo($bytes, $preamble.Length)

    $latestPath = Join-Path -Path $dataLastRoot -ChildPath ([IO.Path]::GetFileName($dataset.Target))
    $plans.Add([pscustomobject]@{
        Dataset = $dataset.Name
        Target = $dataset.Target
        TargetHash = $target.Hash
        LatestPath = $(if (Test-Path -LiteralPath $latestPath -PathType Leaf) { $latestPath } else { '' })
        Bytes = $bytes
        Changed = ($imported.Count -gt 0 -or $typeRepaired -gt 0 -or $typeBlanked -gt 0)
        SmartInventoryRows = $target.Rows.Count
        SmartInventoryDays = $targetDays.Count
        SmartInventoryFirstDay = $firstTargetDay
        LegacyRows = $legacy.Rows.Count
        LegacyInvalidDates = $legacyInvalidDates
        LegacySkippedCoveredDays = $skippedCovered
        ImportedRows = $imported.Count
        ImportedDays = $importedDays.Count
        ImportedFirstDay = $(if ($importedDays.Count -gt 0) { $importedDays[0] } else { '' })
        ImportedDaysAfterFirstSmartInventoryDay = @($importedDays | Where-Object { $_ -gt $firstTargetDay }).Count
        LegacyBlankedNonNumericValues = $(($blankedByColumn.Values | Measure-Object -Sum).Sum -as [int])
        LegacyBlankedColumns = (@($blankedByColumn.GetEnumerator() | ForEach-Object { '{0}={1}' -f $_.Key, $_.Value }) -join '; ')
        LegacyDerivedValues = $(($derivedByColumn.Values | Measure-Object -Sum).Sum -as [int])
        LegacyDerivedColumns = (@($derivedByColumn.GetEnumerator() | ForEach-Object { '{0}={1}' -f $_.Key, $_.Value }) -join '; ')
        MailboxTypesRepaired = $typeRepaired
        MailboxTypesBlanked = $typeBlanked
        OverlapDays = $differences.Count
        OverlapMedianDifferencePercent = $(if ($null -ne $medianDifference) { [Math]::Round($medianDifference, 3) } else { '' })
        OverlapMaxDifferencePercent = $(if ($null -ne $maxDifference) { [Math]::Round($maxDifference, 3) } else { '' })
        MergedRows = $merged.Count
        Status = 'Planned'
    })
}

if ($plans.Count -eq 0) { throw "No legacy daily CSV found in $legacyRoot" }
$plans | Select-Object Dataset, SmartInventoryRows, SmartInventoryFirstDay, LegacyRows, ImportedRows, ImportedDays, ImportedFirstDay, ImportedDaysAfterFirstSmartInventoryDay, LegacyBlankedNonNumericValues, LegacyDerivedValues, MailboxTypesRepaired, MailboxTypesBlanked, OverlapDays, OverlapMedianDifferencePercent, OverlapMaxDifferencePercent, MergedRows |
    Format-List | Out-String -Width 240 | Write-Host
foreach ($plan in $plans) {
    if ($plan.LegacyBlankedColumns) { Write-Host ("[{0}] Non-numeric legacy values made empty: {1}" -f $plan.Dataset, $plan.LegacyBlankedColumns) }
    if ($plan.LegacyDerivedColumns) { Write-Host ("[{0}] Missing legacy values computed from the total: {1}" -f $plan.Dataset, $plan.LegacyDerivedColumns) }
}

if ($Execute) {
    foreach ($plan in $plans) {
        if (-not $plan.Changed) { $plan.Status = 'NothingToDo'; continue }
        if ((Get-FileSha256 $plan.Target) -ne $plan.TargetHash) { $plan.Status = 'ChangedDuringRun'; throw "[$($plan.Dataset)] $($plan.Target) changed during the run (collector running?). Nothing written for it; rerun outside the collector window." }
        $backupDir = Join-Path -Path $BackupFolderPath -ChildPath $plan.Dataset
        New-Item -ItemType Directory -Path $backupDir -Force | Out-Null
        foreach ($source in @($plan.Target, $plan.LatestPath) | Where-Object { $_ }) {
            $backupName = '{0}_{1}' -f $(if ($source -eq $plan.Target) { 'DATA-ALL' } else { 'DATA-LAST' }), [IO.Path]::GetFileName($source)
            $backup = Join-Path -Path $backupDir -ChildPath $backupName
            Copy-Item -LiteralPath $source -Destination $backup -Force
            if ((Get-FileSha256 $backup) -ne (Get-FileSha256 $source)) { throw "Backup verification failed: $backup" }
        }
        Write-FileAtomically -Path $plan.Target -Bytes $plan.Bytes
        if ($plan.LatestPath) { Write-FileAtomically -Path $plan.LatestPath -Bytes $plan.Bytes }
        $plan.Status = 'Merged'
        Write-Host ("[{0}] Merged: {1} rows written to {2}{3}" -f $plan.Dataset, $plan.MergedRows, $plan.Target, $(if ($plan.LatestPath) { " and $($plan.LatestPath)" } else { '' }))
    }
    Write-Host ("Backup: {0}" -f $BackupFolderPath)
}
else {
    Write-Host 'Preview only. Run again with -Execute to write the merged files.'
}

$plans | Select-Object -Property * -ExcludeProperty Bytes, TargetHash | Export-Csv -LiteralPath $ReportPath -NoTypeInformation -Encoding utf8
Write-Host ("Report: {0} (script v{1})" -f $ReportPath, $ScriptVersion)
Write-SmartM365CompletionBanner -Status Success -ScriptName 'Import-SmartM365-LegacyDailyStats' -StartedAt $startedAt -EndedAt (Get-Date)

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAN/h1j0cbnrTSF
# 9AOvJYAUqW6eXNNeZW+To54ZovkSwqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# M9LHJmyrxaFtoza2zNaQ9k+5t1wwggbtMIIE1aADAgECAhAIT9wzT35FTtvDD4/5
# khg1MA0GCSqGSIb3DQEBCwUAMGkxCzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdp
# Q2VydCwgSW5jLjFBMD8GA1UEAxM4RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3Rh
# bXBpbmcgUlNBNDA5NiBTSEEyNTYgMjAyNSBDQTEwHhcNMjYwODA1MDAwMDAwWhcN
# MzcxMTA0MjM1OTU5WjBjMQswCQYDVQQGEwJVUzEXMBUGA1UEChMORGlnaUNlcnQs
# IEluYy4xOzA5BgNVBAMTMkRpZ2lDZXJ0IFNIQTI1NiBSU0E0MDk2IFRpbWVzdGFt
# cCBSZXNwb25kZXIgMjAyNiAxMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEAtnum8sn+zUr41JtMZbP9OMYw+HwJDpG5xkIu/lqcfNYmMX81YmsUiHLbh9yk
# peWBGKTLhYBrAN9Tdg/QEzG32XcObmgIblnr0CoQ3WSAeDZ6nH6X6VkFyYkJw3QB
# JREwvm4UhLzSxmwPA7cFKRTEOMsmEEj6qJk/dqLEAL+oQYuOwE2UuiX1Vnul8YRe
# IyWd4kgLn9gq6LNXM0UplkR6jL/QHxmb6fMoGBJYbnaUI7XD6cKDpekK2SVMld4i
# DbzeHDtOaaxldH5IxuNusQ69nd8/ZXEiB5Hbxj3RlK13cX1W4DlFXKdv/CEhM8Cj
# 1vvlmvhNroyPdRGbbpBlgyf8Wdu5N6ByhFwURn0U6ozlPoxN22v+fviUhP+6DR54
# 7OZnpBMWDfei1f5sVGwiiW/KQTWOK97g+4RJpPzPNV4VYMAwO2jM2Aty2QYPVmOQ
# TJm0msuXnJrSbl2gf9JylpkJlWXqk1Q4LJsxz+TELoQCZIljbgvTJgoPU2R12ydv
# 8i1UqL/adelA0y7U9Pmmtbze9Xx3rtajC5SzQd1jgfwAwsa90v9YcSPdmeoyoBBA
# /27cCL237l5DTYYPDLQ4ON3OLTGWnvRb6jDrf/T75gMRfUzSLCBQfBusm9+mSWRl
# C/Df6S/e9Q8i13CuhzOT2Jx+V/nlbXM4QoBwlUAhelwwJT0CAwEAAaOCAZUwggGR
# MAwGA1UdEwEB/wQCMAAwHQYDVR0OBBYEFBTJY4owLtRK+26U8+bjQH717M3iMB8G
# A1UdIwQYMBaAFO9vU0rp5AZ8esrikFb2L9RJ7MtOMA4GA1UdDwEB/wQEAwIHgDAW
# BgNVHSUBAf8EDDAKBggrBgEFBQcDCDCBlQYIKwYBBQUHAQEEgYgwgYUwJAYIKwYB
# BQUHMAGGGGh0dHA6Ly9vY3NwLmRpZ2ljZXJ0LmNvbTBdBggrBgEFBQcwAoZRaHR0
# cDovL2NhY2VydHMuZGlnaWNlcnQuY29tL0RpZ2lDZXJ0VHJ1c3RlZEc0VGltZVN0
# YW1waW5nUlNBNDA5NlNIQTI1NjIwMjVDQTEuY3J0MF8GA1UdHwRYMFYwVKBSoFCG
# Tmh0dHA6Ly9jcmwzLmRpZ2ljZXJ0LmNvbS9EaWdpQ2VydFRydXN0ZWRHNFRpbWVT
# dGFtcGluZ1JTQTQwOTZTSEEyNTYyMDI1Q0ExLmNybDAgBgNVHSAEGTAXMAgGBmeB
# DAEEAjALBglghkgBhv1sBwEwDQYJKoZIhvcNAQELBQADggIBAI3FOmEenVIK35ms
# CYB+fShAsWvSYvLBItoNdAgQ2jIqrGsVsluXMJU/+mRebBc52s6lbKAvOVPXaizm
# KkMLLflEEKDZQx4CkS2t8aHPjkXha3hYZ010htFa3dhNgmalH5vuWvh3tTCf4frT
# S7gPtGc4Z/xaPhQ2AB1mR8eEe/WbH0RWHvVIl6VwQ3+g5FKNfN2N/DWJkf13w2H+
# 2GfqEfbd35Ww8CvoYBjLNIDTadcPWdgsjsiOaK/7EsKJgLjUNIVgvcaFOLLQ/Glr
# A+0ZHJoFUbOr5SJN8zykPspXIXlpDJY/gqFUZRROeab9GVgmhbdOJcD/63RhxPah
# FUGbckRONqMe6DYAv6/mOG0pWd3cPStsdcS7buj5DyniwRY8yooMH6ptx5vpP/pZ
# zBPBeZD2U4IsthyxB5Jaa8qrOkB5z160TXiM5ADMspZ0TfD9MJoq0tFpFPssKRFh
# WeEDYPvcUuN7U7lvcdHl4ezQ3NT/7Ffs1sR1yh/LRbdZ3B3Vc6q2WmD8mDC0p9kz
# l2o73iVtS946IkEj7FkRsZGww1teYxERROC745xrtjvcw9ZyyUjHZWGRIpJeMNsP
# quCDf0fkyHtB+J4AiNZqCQk23rxh+KbpyMTNVKItJ5l92Svl20U9NbqMBOVYl1h5
# 4NEYLJq1/xHWFKPNK903zJZA9P2DMYIFvjCCBboCAQEwYjBOMR4wHAYDVQQDDBV3
# b3JrcGxhY2VjbG91ZGh1Yi5jb20xLDAqBgkqhkiG9w0BCQEWHWNvbnRhY3RAd29y
# a3BsYWNlY2xvdWRodWIuY29tAhAebu87xzjhs0Q4yPEDH+JoMA0GCWCGSAFlAwQC
# AQUAoIGEMBgGCisGAQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwG
# CisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwLwYJKoZI
# hvcNAQkEMSIEIKSz+WecXRWqddtFHDLQsD5lBRffsC//cVuU4n4fZzRhMA0GCSqG
# SIb3DQEBAQUABIIBgGJuDVJf3ohSLDw1Boyc5ZaIgvyzJx6Hme48qDJjUi8RuVdS
# BvjyoOFcLXVwTUH9yRDOhvRwRpOG4tY9760NPPOM6Tqo7hntfpeCOm7r8u2l0L+h
# LZPnuTpoHyNKkj5i+VStnWcru5upNZ7GN5Mu9rzHE2GLT0bkgPku8MXeLETM2hGL
# H24qA6YENjIqVMFUYF7ZiNjPnZ/NlpkxyhsMQpyQKW7jBoeoTstedn/jCOrRNv6m
# jy2nrivEHXX/yvexD56yQ5T37SSQU3Z7UetEB9OO0ADiiaZ8rFlx7NmIsEwo23gS
# F+FTXfwhOmUKbAeGZ1sRMKVOw2A5q8YG5VItFYV8j2d8ZLIXiP6YHTddofFFjFHP
# Xb1TBu1NS+zW47Vrs7xEIYuIxZkjoqh/oB/I+5Jb9qt7Bdc5qd6Mfgo+jitdyN1q
# H5m3WWMRLFuEBuxz9ybZpk0jRZOza4WUh2m5Q7JkjF32seMFiPheETzZq2nXzbhu
# f21pfoLl0+35eaYzraGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTYxMzM0
# MTFaMC8GCSqGSIb3DQEJBDEiBCAX2g/WXEw180q7gOCok6RanlL9Z2PLtimc1xTt
# wj/CtjANBgkqhkiG9w0BAQEFAASCAgCLkNrIm7xTAyWqu1Cl0RBPVtM6wJC9FKgD
# 44drG5/7jjNv1o2yLbqphNB15UEvsulucfwn7Yo+oawto+OQ48cB1xT1ccY+rlDw
# vIoef5dvAknWsD3KZhyyYlpo8uUa5xDlwvK/mGqLRB4u/H6qqwy36JefnDcILowN
# 0TTkh3GxxH1RFHPiK6TTJn1Gc5wf/Xe8xfbyLQCy0amjj3g6ECZzL8olRlsm1uIv
# rz+xMQyPOKxhaNRf/csOdARE/PrBcY6XQQDnTVS+W5HhpNLsLKj6OZ5UIcLb8JUK
# tnUG5QJezHr2LLvD9eSjb4sl1LjZrKu5EmTY6olTnDaG6SluJdNlKgrfwN2GvhNH
# U2TDBlWu9XUARRe0Ekoss7XLhg6Hku2EItNICNQk7D77YJAD0V0eOyIVACy34TZp
# cCv7//nt955eaYQzB2tRTdYBz/587f3YbBWRuZyYghV8M0l4gmekQRukLjlPdxRe
# EB6mJJMztKRlDibxPV8vfGYhStrOOVxPZfMQecg4pRHjew3jE5e/1L6pA0DBqgGZ
# Ti5DfxXPHd0VPdFhCWc08YkDKcMQMQq7hhdOq7qiQhlxZmqzY1FoMOIGrKq73lGr
# 6T5E2Gh1WuDLnlhzJIRnk5WQEDD8nX/+c4MrT6YVewPa9XSQV9xMgGR3ojqhr1UT
# jxem5IM8lA==
# SIG # End signature block
