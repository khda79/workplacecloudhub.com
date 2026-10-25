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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBE7q4Q46dS2ZNH
# 95YRijJFVgOlkxnNl0YNkpEEK/np86CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMDjmT/8zNZUqSA2vWqowmMDVf4HgVXY507/zNWZa3wfMA0GCSqG
# SIb3DQEBAQUABIIBgGzy5F4P00KiTAm+ciVn3LPCNvWhHgvXjCjHnjdOp5k9cf7h
# +btzSdvuEOr4kdNZyYpcEjIUR9cxeRHJQPs96P+Vua9t8YVTX+g8VQZJNhrjrvhq
# YcMxFKdJqwVDRb2Ip2wXaOlqE6hGAqH70UbAtGLCTLES7Wf1t+pNuWcUCZ6yz2Au
# tAUUl8UVtWKxKKa3tHW4z2QpbroclyY1QKDsMZ8R+Px0Va6c6/2KxKG84FLzBxgV
# YXI5e1g1Ih9klVFGQxlaqnSwttjiCRVNGmt9sMMQ7sYsn/xEFqE9+YkfB7iutLlp
# gbR3YLE/mKGjo24eI6e9AWXVy2+cpb5QDQm2kNAZaeCMHyFYB5hJe3/IlyeLL4wG
# 5B/E9bHTgFAl3rEnb868UBRPS/20Nq6vO+tDXBbBhhZC38GL/0j4K9xilwxtcDsy
# RI424NLDQMusxiMUSlpJVVJrfVDm9HX9oOK1h8DJOgW7G7DrQhm3HGzqX4E4tu8k
# Tn1R+yE4m7vFwXY0LaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMjU3
# NTNaMC8GCSqGSIb3DQEJBDEiBCC+3PIaIKtKqX/ZGV0V3Er/MRaNOSx08j9llKpk
# 73XJ2DANBgkqhkiG9w0BAQEFAASCAgBYxAB6TKLVkcQkmfTghOyqgBHe2WtMjq5v
# RLvoTDhcbHs+tPwQA0IXf/1K0vaNNpHaY77RGpBKpefhMRexzn5/vsn5YnOM0H0v
# McBfHLWAACYQZzIgYD/RB2HpcIpNp21bUDSrEuqqS6fIyB4c1NpHwoxzaIwYbYeE
# jef8prepLiSO1KGv3zuMeC6g9IIbCpOCAcI4kLTCsrS7K7FSxMZ8J7ylq4ACJeed
# fJw5SEqhjRM1I0eAL4UG/S0/n19D7kW5luNziQFktk3njqoExkNGhPTF+9+SPd+G
# 80+y5ychtolpD3ogJ/1jhYWWRBRum7e79Q2+EN9xOpBEkX3Cd5uBszMdTRgHtmG3
# YLO1NwnUC1CXWFVIt8q5CLpzdx2+HDdPaZlj6yS85Uzzl2A0kNfWqyTdrb7HWFSk
# ywp5t/whkp4MiQoAasf8kcfkwQXul7CXlQTdkHaHUbrSA9HBuV5G9tkHoLc77CXN
# bKAz9yHJRmd9y2ngDZzNFbAympeZ4US4ikPlYxa4hUfRvVr85zc+XfDKRW54D2Xo
# HxI7nhXjCMrr1IXfA++8BSvamTyhCs2ObDVGScrpYhnPQyre69vrO7j/7Bw06oue
# YbPHsrCGDqSnVkm8hJzFy+ecgMMSJKdTNSQVLw0CFihxWbW3STVW9I4ZDas+TsbD
# nHZdHiZuRQ==
# SIG # End signature block
