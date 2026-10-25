# Version 1.1.1
# Offline comparison of the legacy hybrid identity / Windows 11 readiness issue CSVs with the new
# prepared outputs (docs/HYBRID-IDENTITY-DEVICE-READINESS-FINDINGS.md). Read-only on its inputs.
# A difference is justified only by evidence checked here; everything else stays unexplained.
# Hybrid: legacy verdicts versus new findings, per retained legacy rule and account.
# Devices: the new output holds facts, not verdicts. Each legacy condition is re-applied to the
# facts only where every input of its applicability is known; otherwise the device is undetermined.
[CmdletBinding()]
param(
    [string]$LegacyHybridPath,
    [string]$HybridFindingsPath,
    # Evidence for rules 51 and 58: the AD users and the governed classification rules.
    [string]$AdUsersPath,
    [string]$AccountClassificationConfigPath,
    [string]$LegacyReadinessPath,
    [string]$ReadinessGapPath,
    [string]$DeviceInventoryPath,
    # Date on which the legacy readiness CSV computed its reboot ages (its run date).
    [Nullable[datetime]]$LegacyReferenceDate,
    [Parameter(Mandatory)][string]$ReportPath,
    [string]$DetailPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-Key([AllowNull()][object]$Value) { ([string]$Value).Trim().Trim('{', '}').ToLowerInvariant() }
function New-Set { , [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) }
function Read-Rows([string]$Path, [string]$Name) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "$Name was not found: $Path" }
    return @(Import-Csv -LiteralPath $Path)
}
function Get-Text([AllowNull()][object]$Row, [string[]]$Names) {
    if ($null -eq $Row) { return '' }
    foreach ($name in $Names) { $property = $Row.PSObject.Properties[$name]; if ($property -and -not [string]::IsNullOrWhiteSpace([string]$property.Value)) { return ([string]$property.Value).Trim() } }
    return ''
}

$retainedHybrid = @(1, 2, 3, 6, 9, 16, 19, 20, 21, 22, 23, 24, 25, 26, 29, 30, 31, 32, 35, 38, 39, 40, 51, 54, 58)
$report = [ordered]@{ Hybrid = $null; Devices = $null }
$details = [Collections.Generic.List[object]]::new()
function Add-Detail([string]$Area, [string]$Rule, [string]$Key, [string]$Difference, [bool]$Justified, [string]$Explanation) {
    $script:details.Add([pscustomobject][ordered]@{Area = $Area; Rule = $Rule; Key = $Key; Difference = $Difference; Justified = $Justified; Explanation = $Explanation })
}

if ($LegacyHybridPath -and $HybridFindingsPath) {
    $legacy = Read-Rows $LegacyHybridPath 'Legacy hybrid identity issues'
    $new = Read-Rows $HybridFindingsPath 'Hybrid identity findings'
    $adByGuid = @{}
    if ($AdUsersPath) { foreach ($row in (Read-Rows $AdUsersPath 'AD users')) { $guid = Get-Key $row.ObjectGUID; if ($guid -and -not $adByGuid.ContainsKey($guid)) { $adByGuid[$guid] = $row } } }
    $populationRules = $null
    if ($AccountClassificationConfigPath) {
        Import-Module (Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) 'SmartM365\SmartInventory\Common\SmartM365.AccountClassification.psd1') -MinimumVersion '1.2.0' -ErrorAction Stop
        $populationRules = (Read-SmartM365AccountClassification -Path $AccountClassificationConfigPath).Population
    }
    # Legacy-only rule 51: the legacy producer counted regex matches of the primary SMTP in the whole
    # proxy text. Justified only when at most one whole entry equals it and the substring count is higher.
    function Test-Rule51Justification([string]$Guid) {
        if (-not $adByGuid.ContainsKey($Guid)) { return $null }
        $row = $adByGuid[$Guid]
        $proxy = Get-Text $row @('ProxyAddresses')
        $primary = Get-Text $row @('EmailAddress', 'PrimarySmtpAddress', 'PrimarySMTPaddress')
        if (-not $primary) { foreach ($entry in @($proxy -split '[;|,]')) { $address = $entry.Trim() -replace '^(smtp|SMTP):', ''; if ($address -match '@') { $primary = $address; break } } }
        if (-not $primary -or -not $proxy) { return $null }
        $whole = @($proxy -split '[;|,]' | Where-Object { ($_.Trim() -replace '^(smtp|SMTP):', '') -ieq $primary }).Count
        $substring = [regex]::Matches($proxy.ToLowerInvariant(), [regex]::Escape($primary.ToLowerInvariant())).Count
        if ($whole -le 1 -and $substring -gt 1) { return "Legacy substring match only: $whole whole proxy entry, $substring substring matches." }
        return $null
    }
    # Legacy-only rule 58: justified only when the account type was in the legacy list and the governed
    # configuration does not classify it as Non-human.
    function Test-Rule58Justification([string]$Guid) {
        if ($null -eq $populationRules -or -not $adByGuid.ContainsKey($Guid)) { return $null }
        $type = Get-Text $adByGuid[$Guid] @('AccountType')
        if ($type -notin @('Service Account', 'Admin Account', 'System Account', 'Generic Account')) { return $null }
        $population = Get-SmartM365AccountPopulation -PopulationRules $populationRules -AccountType $type
        if ($population -ne 'Non-human') { return "AccountType '$type' was in the legacy list; the governed population is '$population'." }
        return $null
    }
    $legacyByRule = @{}; $newByRule = @{}
    foreach ($row in $legacy) { $n = [int]$row.IssueNumber; if (-not $legacyByRule.ContainsKey($n)) { $legacyByRule[$n] = New-Set }; [void]$legacyByRule[$n].Add((Get-Key $row.ObjectGUID)) }
    foreach ($row in $new) { $n = [int]$row.'Legacy Issue Number'; if (-not $newByRule.ContainsKey($n)) { $newByRule[$n] = New-Set }; $key = Get-Key $row.'AD Object GUID'; if (-not $key) { $key = 'unresolved:' + (Get-Key $row.'Finding Key') }; [void]$newByRule[$n].Add($key) }
    $rules = [Collections.Generic.List[object]]::new()
    foreach ($n in $retainedHybrid) {
        # Plain assignments: an if expression would enumerate the set.
        if (-not $legacyByRule.ContainsKey($n)) { $legacyByRule[$n] = New-Set }
        if (-not $newByRule.ContainsKey($n)) { $newByRule[$n] = New-Set }
        $l = $legacyByRule[$n]; $r = $newByRule[$n]
        $both = @($l | Where-Object { $r.Contains($_) }).Count
        $justified = 0; $unexplained = 0
        foreach ($key in @($l | Where-Object { -not $r.Contains($_) } | Sort-Object)) {
            $reason = switch ($n) { 51 { Test-Rule51Justification $key } 58 { Test-Rule58Justification $key } default { $null } }
            if ($reason) { $justified++ } else { $unexplained++ }
            Add-Detail 'Hybrid' ([string]$n) $key 'Legacy only' ([bool]$reason) $(if ($reason) { $reason } else { 'No supporting evidence.' })
        }
        # A new finding is never justified by the legacy design: it is a regression until reviewed.
        $newOnly = @($r | Where-Object { -not $l.Contains($_) } | Sort-Object)
        foreach ($key in $newOnly) { $unexplained++; Add-Detail 'Hybrid' ([string]$n) $key 'New only' $false 'No supporting evidence.' }
        $rules.Add([ordered]@{LegacyIssueNumber = $n; Legacy = $l.Count; New = $r.Count; Both = $both; LegacyOnly = ($l.Count - $both); NewOnly = $newOnly.Count; Justified = $justified; Unexplained = $unexplained })
    }
    $notCarried = [ordered]@{}
    foreach ($n in @($legacyByRule.Keys | Where-Object { $_ -notin $retainedHybrid } | Sort-Object)) { $notCarried[[string]$n] = $legacyByRule[$n].Count }
    $report.Hybrid = [ordered]@{ Rules = $rules.ToArray(); NotCarriedOverLegacyAccounts = $notCarried; UnexplainedDifferences = (@($rules | ForEach-Object { $_.Unexplained }) | Measure-Object -Sum).Sum }
}

if ($LegacyReadinessPath -and $ReadinessGapPath -and $DeviceInventoryPath) {
    Import-Module (Join-Path $PSScriptRoot 'DeviceReadinessGapEvidence.psm1') -Force -ErrorAction Stop
    $legacy = Read-Rows $LegacyReadinessPath 'Legacy Windows 11 readiness issues'
    $facts = Read-Rows $ReadinessGapPath 'Device readiness gap evidence'
    $inventory = Read-Rows $DeviceInventoryPath 'Device inventory evidence'
    $inventoryById = @{}; foreach ($row in $inventory) { $id = Get-Key $row.'Device Source ID'; if ($id -and -not $inventoryById.ContainsKey($id)) { $inventoryById[$id] = $row } }
    $factsByGuid = @{}; foreach ($row in $facts) { $guid = Get-Key $row.'AD Object GUID'; if ($guid -and -not $factsByGuid.ContainsKey($guid)) { $factsByGuid[$guid] = $row } }
    $codes = [ordered]@{
        'ISSUE-032' = 'Entra registration pending'
        'ISSUE-018' = 'Physical memory below 8 GB, eligibility not Not capable'
        'ISSUE-054' = 'No reboot for 10 days or more, Windows 10, eligibility not Not capable'
        'ISSUE-055' = 'No reboot for 30 days or more, eligibility not Not capable'
        'ISSUE-070' = 'BIOS firmware, Windows 10, Secure Boot data present, eligibility not Not capable'
        'ISSUE-070B' = 'Secure Boot NotSupported, Windows 10, eligibility not Not capable'
    }
    $legacyByCode = @{}; foreach ($code in $codes.Keys) { $legacyByCode[$code] = New-Set }
    foreach ($row in $legacy) { if ($legacyByCode.ContainsKey($row.IssueCode)) { [void]$legacyByCode[$row.IssueCode].Add((Get-Key $row.ObjectGUID_Norm)) } }

    # Re-applies a legacy condition to the facts: $true, $false, or $null when an input is unknown.
    # Eligibility and Windows generation come from a usable Device Inventory Evidence row only.
    function Test-LegacyCondition([string]$Code, [object]$Fact, [Nullable[int]]$Days) {
        $item = if ($Fact.'Device Source ID') { $inventoryById[(Get-Key $Fact.'Device Source ID')] } else { $null }
        $eligibility = Get-Text $item @('Windows 11 Upgrade Eligibility')
        $generation = Get-Text $item @('Windows Generation')
        $firmware = [string]$Fact.'Firmware Type (source)'; $secureBoot = [string]$Fact.'Secure Boot Status (source)'
        if ($Code -eq 'ISSUE-032') { if ($Fact.'Entra Registration Pending' -eq '') { return $null }; return $Fact.'Entra Registration Pending' -eq 'True' }
        if (-not $eligibility) { return $null }
        $notCapable = $eligibility -eq 'Not capable'
        switch ($Code) {
            'ISSUE-018' { if ($Fact.'Physical Memory (GB)' -eq '') { return $null }; return ([double]$Fact.'Physical Memory (GB)' -lt 8) -and -not $notCapable }
            'ISSUE-055' { if ($null -eq $Days) { return $null }; return ($Days -ge 30) -and -not $notCapable }
            'ISSUE-054' { if ($null -eq $Days -or -not $generation) { return $null }; return ($Days -ge 10) -and ($generation -eq 'Windows 10') -and -not $notCapable }
            'ISSUE-070' { if (-not $generation -or -not $secureBoot -or -not $firmware) { return $null }; return ($generation -eq 'Windows 10') -and -not $notCapable -and ($firmware -ieq 'BIOS') }
            'ISSUE-070B' { if (-not $generation -or -not $secureBoot) { return $null }; return ($generation -eq 'Windows 10') -and -not $notCapable -and ($secureBoot -ieq 'NotSupported') }
        }
    }
    function Get-FactDays([object]$Fact) { if ($Fact.'Days Since Last Reboot' -ne '') { [int]$Fact.'Days Since Last Reboot' } else { $null } }
    # A reboot-age difference is attributed to the reference dates only when both dates are known and
    # differ, both ages are recomputed from the same source reboot value, the recomputed new age equals
    # the age declared in the facts, and the two ages reproduce the legacy and the new verdicts.
    function Get-ReferenceDateJustification([string]$Code, [object]$Fact, [bool]$LegacyVerdict) {
        if ($Code -notin @('ISSUE-054', 'ISSUE-055') -or $null -eq $LegacyReferenceDate) { return $null }
        $source = [string]$Fact.'Last Reboot (source)'; $declared = Get-FactDays $Fact
        $evidenceDate = [datetime]::MinValue
        if (-not $source -or $null -eq $declared -or -not [datetime]::TryParseExact([string]$Fact.'Evidence Date', 'yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$evidenceDate)) { return $null }
        $legacyDate = $LegacyReferenceDate.Date
        if ($legacyDate -eq $evidenceDate.Date) { return $null }
        # End of each reference day: the age is in whole calendar days; a same-day reboot is not future.
        $atNew = ConvertTo-RebootFact -Value $source -Reference ([datetimeoffset]::new([datetime]::SpecifyKind($evidenceDate.Date.AddDays(1).AddTicks(-1), [DateTimeKind]::Utc)))
        $atLegacy = ConvertTo-RebootFact -Value $source -Reference ([datetimeoffset]::new([datetime]::SpecifyKind($legacyDate.AddDays(1).AddTicks(-1), [DateTimeKind]::Utc)))
        if ($null -eq $atNew.Days -or $null -eq $atLegacy.Days -or $atNew.Days -ne $declared) { return $null }
        $legacyCheck = Test-LegacyCondition $Code $Fact $atLegacy.Days
        $newCheck = Test-LegacyCondition $Code $Fact $atNew.Days
        if ($null -eq $legacyCheck -or $null -eq $newCheck -or $legacyCheck -ne $LegacyVerdict -or $newCheck -eq $LegacyVerdict) { return $null }
        return ('Reference date differs: {0} days at the legacy date {1:yyyy-MM-dd}, {2} days at the evidence date {3:yyyy-MM-dd}, both from the same source reboot value.' -f $atLegacy.Days, $legacyDate, $atNew.Days, $evidenceDate)
    }
    $results = [Collections.Generic.List[object]]::new()
    foreach ($code in $codes.Keys) {
        $derived = New-Set; $undetermined = New-Set
        foreach ($fact in $facts) {
            $guid = Get-Key $fact.'AD Object GUID'; if (-not $guid) { continue }
            $verdict = Test-LegacyCondition $code $fact (Get-FactDays $fact)
            if ($null -eq $verdict) { [void]$undetermined.Add($guid) } elseif ($verdict) { [void]$derived.Add($guid) }
        }
        $l = $legacyByCode[$code]
        $buckets = [ordered]@{ Both = 0; 'Legacy only: undetermined' = 0; 'Legacy only: not in prepared device scope' = 0; 'Legacy only: justified' = 0; 'Legacy only: unexplained' = 0; 'New only: justified' = 0; 'New only: unexplained' = 0 }
        foreach ($guid in $l) {
            if ($derived.Contains($guid)) { $buckets.Both++; continue }
            if (-not $factsByGuid.ContainsKey($guid)) { $bucket = 'Legacy only: not in prepared device scope'; $reason = 'AD object GUID absent from the device facts.' }
            elseif ($undetermined.Contains($guid)) { $bucket = 'Legacy only: undetermined'; $reason = 'An input of the legacy applicability is unknown in the facts.' }
            else {
                $reason = Get-ReferenceDateJustification $code $factsByGuid[$guid] $true
                $bucket = if ($reason) { 'Legacy only: justified' } else { 'Legacy only: unexplained' }
                if (-not $reason) { $reason = 'No supporting evidence.' }
            }
            $buckets[$bucket]++
            Add-Detail 'Devices' $code $guid $bucket ($bucket -eq 'Legacy only: justified') $reason
        }
        foreach ($guid in $derived) {
            if ($l.Contains($guid)) { continue }
            $reason = Get-ReferenceDateJustification $code $factsByGuid[$guid] $false
            $bucket = if ($reason) { 'New only: justified' } else { 'New only: unexplained' }
            $buckets[$bucket]++
            Add-Detail 'Devices' $code $guid $bucket ([bool]$reason) $(if ($reason) { $reason } else { 'No supporting evidence.' })
        }
        $results.Add([ordered]@{ LegacyIssueCode = $code; Condition = $codes[$code]; LegacyVerdicts = $l.Count; ConditionMetOnFacts = $derived.Count; Undetermined = $undetermined.Count; Unexplained = $buckets.'Legacy only: unexplained' + $buckets.'New only: unexplained'; Buckets = $buckets })
    }
    $report.Devices = [ordered]@{
        Note = 'New output holds facts; legacy conditions are re-applied for comparison only.'
        LegacyReferenceDate = $(if ($null -ne $LegacyReferenceDate) { $LegacyReferenceDate.ToString('yyyy-MM-dd') } else { '' })
        Codes = $results.ToArray()
        UnexplainedDifferences = (@($results | ForEach-Object { $_.Unexplained }) | Measure-Object -Sum).Sum
    }
}

$ReportPath = [IO.Path]::GetFullPath($ReportPath)
New-Item -ItemType Directory -Path (Split-Path -Parent $ReportPath) -Force | Out-Null
[IO.File]::WriteAllText($ReportPath, ($report | ConvertTo-Json -Depth 8), [Text.UTF8Encoding]::new($false))
if ($DetailPath) {
    $DetailPath = [IO.Path]::GetFullPath($DetailPath)
    New-Item -ItemType Directory -Path (Split-Path -Parent $DetailPath) -Force | Out-Null
    if ($details.Count) { $details | Sort-Object Area, Rule, Difference, Key | Export-Csv -LiteralPath $DetailPath -NoTypeInformation -Encoding utf8NoBOM }
    else { [IO.File]::WriteAllText($DetailPath, '"Area","Rule","Key","Difference","Justified","Explanation"' + [Environment]::NewLine, [Text.UTF8Encoding]::new($false)) }
}
$report
