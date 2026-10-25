#Requires -Version 7.0
# Version 1.1.1
[CmdletBinding()]
param()
# Synthetic fixtures only. No tenant initialization, collectors, production files or Power BI.
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$script:checks = 0
function Check([bool]$Pass, [string]$Message) { $script:checks++; if (-not $Pass) { throw $Message } }
$root = Join-Path ([IO.Path]::GetTempPath()) ('CompareReadiness-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root -Force | Out-Null
$comparator = Join-Path $PSScriptRoot 'Compare-ReadinessFindings.ps1'
function Write-Rows([string]$Name, [object[]]$Rows) { $path = Join-Path $root $Name; $Rows | Export-Csv -LiteralPath $path -NoTypeInformation -Encoding utf8NoBOM; $path }
function Get-Detail($Rows, [string]$Rule, [string]$Key) { @($Rows | Where-Object { $_.Rule -eq $Rule -and $_.Key -eq $Key }) }
try {
    $rules = Join-Path $root 'rules.json.txt'
    [IO.File]::WriteAllText($rules, '{"SchemaVersion":"1.0","RuleVersion":"SYNTHETIC-1","Population":{"Human":["Named Account"],"NonHuman":["Service Account"],"ReviewRequired":["Generic Account","Admin Account"]}}')
    $legacyHybrid = Write-Rows 'legacy-hybrid.csv' @(
        [pscustomobject]@{IssueNumber = '1'; ObjectGUID = '{G-1}'}
        [pscustomobject]@{IssueNumber = '1'; ObjectGUID = 'g-2'}
        [pscustomobject]@{IssueNumber = '51'; ObjectGUID = 'g-sip'}
        [pscustomobject]@{IssueNumber = '51'; ObjectGUID = 'g-real'}
        [pscustomobject]@{IssueNumber = '51'; ObjectGUID = 'g-noad'}
        [pscustomobject]@{IssueNumber = '58'; ObjectGUID = 'g-generic'}
        [pscustomobject]@{IssueNumber = '58'; ObjectGUID = 'g-service'}
        [pscustomobject]@{IssueNumber = '7'; ObjectGUID = 'g-5'})
    $newHybrid = Write-Rows 'new-hybrid.csv' @(
        [pscustomobject]@{'Legacy Issue Number' = '1'; 'AD Object GUID' = 'g-1'; 'Finding Key' = 'HYB-001|g-1'}
        [pscustomobject]@{'Legacy Issue Number' = '1'; 'AD Object GUID' = 'g-6'; 'Finding Key' = 'HYB-001|g-6'}
        [pscustomobject]@{'Legacy Issue Number' = '51'; 'AD Object GUID' = 'g-new51'; 'Finding Key' = 'HYB-051|g-new51'})
    $adUsers = Write-Rows 'ad.csv' @(
        [pscustomobject]@{ObjectGUID = 'g-sip'; EmailAddress = 'sip@contoso.com'; ProxyAddresses = 'SMTP:sip@contoso.com;SIP:sip@contoso.com'; AccountType = 'Named Account'}
        [pscustomobject]@{ObjectGUID = 'g-real'; EmailAddress = 'real@contoso.com'; ProxyAddresses = 'SMTP:real@contoso.com;smtp:REAL@contoso.com'; AccountType = 'Named Account'}
        [pscustomobject]@{ObjectGUID = 'g-new51'; EmailAddress = 'new@contoso.com'; ProxyAddresses = 'SMTP:new@contoso.com;SIP:new@contoso.com'; AccountType = 'Named Account'}
        [pscustomobject]@{ObjectGUID = 'g-generic'; EmailAddress = 'generic@contoso.com'; ProxyAddresses = 'SMTP:generic@contoso.com'; AccountType = 'Generic Account'}
        [pscustomobject]@{ObjectGUID = 'g-service'; EmailAddress = 'service@contoso.com'; ProxyAddresses = 'SMTP:service@contoso.com'; AccountType = 'Service Account'})
    $legacyReadiness = Write-Rows 'legacy-readiness.csv' @(
        [pscustomobject]@{IssueCode = 'ISSUE-018'; ObjectGUID_Norm = 'g-low'}
        [pscustomobject]@{IssueCode = 'ISSUE-018'; ObjectGUID_Norm = 'g-notcapable'}
        [pscustomobject]@{IssueCode = 'ISSUE-018'; ObjectGUID_Norm = 'g-missing'}
        [pscustomobject]@{IssueCode = 'ISSUE-018'; ObjectGUID_Norm = 'g-noinventory'}
        [pscustomobject]@{IssueCode = 'ISSUE-055'; ObjectGUID_Norm = 'g-noinventory'}
        [pscustomobject]@{IssueCode = 'ISSUE-055'; ObjectGUID_Norm = 'g-blankelig'}
        [pscustomobject]@{IssueCode = 'ISSUE-054'; ObjectGUID_Norm = 'g-adonly'}
        [pscustomobject]@{IssueCode = 'ISSUE-055'; ObjectGUID_Norm = 'g-refdate'}
        [pscustomobject]@{IssueCode = 'ISSUE-055'; ObjectGUID_Norm = 'g-near'}
        [pscustomobject]@{IssueCode = 'ISSUE-055'; ObjectGUID_Norm = 'g-mismatch'}
        [pscustomobject]@{IssueCode = 'ISSUE-055'; ObjectGUID_Norm = 'g-noevidence'}
        [pscustomobject]@{IssueCode = 'ISSUE-070'; ObjectGUID_Norm = 'g-nofirmware'}
        [pscustomobject]@{IssueCode = 'ISSUE-032'; ObjectGUID_Norm = 'g-adonly'}
        [pscustomobject]@{IssueCode = 'ISSUE-070B'; ObjectGUID_Norm = 'g-low'})
    function New-Fact([string]$Id, [string]$Guid, [string]$Memory = '16', [string]$Firmware = 'Uefi', [string]$SecureBoot = 'Enabled', [string]$Days = '2', [string]$Reboot = '', [string]$Pending = 'False', [string]$Evidence = '2026-10-25') {
        [pscustomobject]@{'Device Source ID' = $Id; 'AD Object GUID' = $Guid; 'Entra Registration Pending' = $Pending; 'Physical Memory (GB)' = $Memory; 'Firmware Type (source)' = $Firmware; 'Secure Boot Status (source)' = $SecureBoot; 'Last Reboot (source)' = $Reboot; 'Days Since Last Reboot' = $Days; 'Evidence Date' = $Evidence }
    }
    $gap = Write-Rows 'gap.csv' @(
        (New-Fact 'i-low' 'g-low' -Memory '4' -SecureBoot 'NotSupported')
        (New-Fact 'i-notcapable' 'g-notcapable' -Memory '4')
        (New-Fact 'i-new' 'g-new' -Memory '6' -Days '45')
        (New-Fact 'i-noinventory' 'g-noinventory' -Memory '4' -Days '40')
        (New-Fact 'i-blankelig' 'g-blankelig' -Days '40')
        (New-Fact '' 'g-adonly' -Memory '' -Firmware '' -SecureBoot '' -Days '12' -Pending 'True')
        # Real date change: 29 days at the evidence date, 30 days at the legacy run date one day later.
        (New-Fact 'i-refdate' 'g-refdate' -Days '29' -Reboot '2026-09-26 08:00:00')
        # Near the threshold, consistent, but 29 days at the legacy date: the legacy verdict is not reproduced.
        (New-Fact 'i-near' 'g-near' -Days '28' -Reboot '2026-09-27 08:00:00')
        # Reproduced defect: the source value gives 30 days at 2026-10-25 but the facts declare 29.
        (New-Fact 'i-mismatch' 'g-mismatch' -Days '29' -Reboot '2026-09-25T00:00:00Z')
        # No evidence date in the facts.
        (New-Fact 'i-noevidence' 'g-noevidence' -Days '29' -Reboot '2026-09-26 08:00:00' -Evidence '')
        (New-Fact 'i-nofirmware' 'g-nofirmware' -Firmware ''))
    $inventory = Write-Rows 'inventory.csv' @(
        [pscustomobject]@{'Device Source ID' = 'i-low'; 'Windows Generation' = 'Windows 10'; 'Windows 11 Upgrade Eligibility' = 'Unknown'}
        [pscustomobject]@{'Device Source ID' = 'i-notcapable'; 'Windows Generation' = 'Windows 10'; 'Windows 11 Upgrade Eligibility' = 'Not capable'}
        [pscustomobject]@{'Device Source ID' = 'i-new'; 'Windows Generation' = 'Windows 11'; 'Windows 11 Upgrade Eligibility' = 'Upgraded'}
        [pscustomobject]@{'Device Source ID' = 'i-blankelig'; 'Windows Generation' = 'Windows 10'; 'Windows 11 Upgrade Eligibility' = ''}
        [pscustomobject]@{'Device Source ID' = 'i-refdate'; 'Windows Generation' = 'Windows 11'; 'Windows 11 Upgrade Eligibility' = 'Upgraded'}
        [pscustomobject]@{'Device Source ID' = 'i-near'; 'Windows Generation' = 'Windows 11'; 'Windows 11 Upgrade Eligibility' = 'Upgraded'}
        [pscustomobject]@{'Device Source ID' = 'i-mismatch'; 'Windows Generation' = 'Windows 11'; 'Windows 11 Upgrade Eligibility' = 'Upgraded'}
        [pscustomobject]@{'Device Source ID' = 'i-noevidence'; 'Windows Generation' = 'Windows 11'; 'Windows 11 Upgrade Eligibility' = 'Upgraded'}
        [pscustomobject]@{'Device Source ID' = 'i-nofirmware'; 'Windows Generation' = 'Windows 10'; 'Windows 11 Upgrade Eligibility' = 'Capable'})
    $inputs = @($legacyHybrid, $newHybrid, $adUsers, $legacyReadiness, $gap, $inventory, $rules)
    $before = @{}; foreach ($path in $inputs) { $before[$path] = (Get-FileHash $path).Hash }
    $common = @{LegacyHybridPath = $legacyHybrid; HybridFindingsPath = $newHybrid; LegacyReadinessPath = $legacyReadiness; ReadinessGapPath = $gap; DeviceInventoryPath = $inventory}
    $report = & $comparator @common -AdUsersPath $adUsers -AccountClassificationConfigPath $rules -LegacyReferenceDate '2026-10-26' -ReportPath (Join-Path $root 'out/report.json') -DetailPath (Join-Path $root 'out/details.csv')
    foreach ($path in $before.Keys) { Check ((Get-FileHash $path).Hash -eq $before[$path]) "Comparator modified its input $path" }
    $details = @(Import-Csv -LiteralPath (Join-Path $root 'out/details.csv'))

    # Hybrid: only evidence justifies a difference; a new rule 51 finding is a regression.
    $rule1 = $report.Hybrid.Rules | Where-Object { $_.LegacyIssueNumber -eq 1 }
    Check ($rule1.Both -eq 1 -and $rule1.LegacyOnly -eq 1 -and $rule1.NewOnly -eq 1 -and $rule1.Unexplained -eq 2) 'Hybrid key normalization or set comparison changed'
    $rule51 = $report.Hybrid.Rules | Where-Object { $_.LegacyIssueNumber -eq 51 }
    Check ($rule51.Justified -eq 1 -and $rule51.Unexplained -eq 3) "Rule 51 justification changed: $($rule51 | ConvertTo-Json -Compress)"
    Check ((Get-Detail $details '51' 'g-sip')[0].Justified -eq 'True' -and (Get-Detail $details '51' 'g-sip')[0].Explanation -match 'substring') 'Legacy substring artifact not justified by its evidence'
    Check ((Get-Detail $details '51' 'g-real')[0].Justified -eq 'False') 'A real legacy case duplicate was justified'
    Check ((Get-Detail $details '51' 'g-noad')[0].Justified -eq 'False') 'A difference without AD evidence was justified'
    $regression = Get-Detail $details '51' 'g-new51'
    Check ($regression.Count -eq 1 -and $regression[0].Difference -eq 'New only' -and $regression[0].Justified -eq 'False') 'An unjustified new rule 51 finding was not reported as a regression'
    $rule58 = $report.Hybrid.Rules | Where-Object { $_.LegacyIssueNumber -eq 58 }
    Check ($rule58.Justified -eq 1 -and $rule58.Unexplained -eq 1 -and (Get-Detail $details '58' 'g-generic')[0].Justified -eq 'True' -and (Get-Detail $details '58' 'g-service')[0].Justified -eq 'False') 'Rule 58 justification must depend on the governed population'
    Check ($report.Hybrid.UnexplainedDifferences -eq 6 -and $report.Hybrid.NotCarriedOverLegacyAccounts['7'] -eq 1) 'Hybrid unexplained total or excluded rules changed'
    $noEvidence = & $comparator -LegacyHybridPath $legacyHybrid -HybridFindingsPath $newHybrid -ReportPath (Join-Path $root 'out/no-evidence.json')
    Check ((($noEvidence.Hybrid.Rules | Where-Object { $_.LegacyIssueNumber -in 51, 58 }) | ForEach-Object { $_.Justified } | Measure-Object -Sum).Sum -eq 0) 'Rules 51/58 were justified without evidence'

    # Devices: undetermined when an applicability input is missing; reboot-age differences need evidence.
    $codes = @{}; foreach ($code in $report.Devices.Codes) { $codes[$code.LegacyIssueCode] = $code }
    Check ($report.Devices.Note -match 'facts' -and $report.Devices.LegacyReferenceDate -eq '2026-10-26') 'Device report metadata changed'
    $c018 = $codes['ISSUE-018']
    Check ($c018.Buckets.Both -eq 1 -and $c018.Buckets.'Legacy only: unexplained' -eq 1 -and $c018.Buckets.'Legacy only: not in prepared device scope' -eq 1 -and $c018.Buckets.'Legacy only: undetermined' -eq 1 -and $c018.Buckets.'New only: unexplained' -eq 1) "Rule 018 buckets changed: $($c018.Buckets | ConvertTo-Json -Compress)"
    Check ((Get-Detail $details 'ISSUE-018' 'g-noinventory')[0].Difference -eq 'Legacy only: undetermined') 'Rule 018 without an inventory row must stay undetermined'
    $c055 = $codes['ISSUE-055']
    Check ((Get-Detail $details 'ISSUE-055' 'g-noinventory')[0].Difference -eq 'Legacy only: undetermined' -and (Get-Detail $details 'ISSUE-055' 'g-blankelig')[0].Difference -eq 'Legacy only: undetermined') 'Rule 055 without usable inventory must stay undetermined'
    Check ((Get-Detail $details 'ISSUE-055' 'g-refdate')[0].Difference -eq 'Legacy only: justified' -and (Get-Detail $details 'ISSUE-055' 'g-refdate')[0].Explanation -match 'Reference date differs') 'Reference-date difference not justified by its source value'
    Check ((Get-Detail $details 'ISSUE-055' 'g-near')[0].Difference -eq 'Legacy only: unexplained') 'A difference near the threshold was justified without evidence'
    Check ((Get-Detail $details 'ISSUE-055' 'g-mismatch')[0].Difference -eq 'Legacy only: unexplained') 'A declared age that differs from the recomputed age was attributed to the dates'
    Check ((Get-Detail $details 'ISSUE-055' 'g-noevidence')[0].Difference -eq 'Legacy only: unexplained') 'A difference without an evidence date was attributed to the dates'
    Check ((Get-Detail $details 'ISSUE-055' 'g-refdate')[0].Explanation -match '30 days at the legacy date 2026-10-26, 29 days at the evidence date 2026-10-25') 'Justification does not state both recomputed ages'
    Check ((Get-Detail $details 'ISSUE-055' 'g-new')[0].Difference -eq 'New only: unexplained') 'An unjustified new rule 055 verdict was hidden'
    Check ((Get-Detail $details 'ISSUE-070' 'g-nofirmware')[0].Difference -eq 'Legacy only: undetermined' -and $codes['ISSUE-070'].Undetermined -ge 1) 'Rule 070 without firmware must stay undetermined'
    Check ((Get-Detail $details 'ISSUE-054' 'g-adonly')[0].Difference -eq 'Legacy only: undetermined') 'Rule 054 without inventory must stay undetermined'
    Check ($codes['ISSUE-032'].Buckets.Both -eq 1 -and $codes['ISSUE-070B'].Buckets.Both -eq 1) 'Pending registration or Secure Boot NotSupported condition changed'
    Check ($report.Devices.UnexplainedDifferences -eq 6) "Device unexplained total changed: $($report.Devices.UnexplainedDifferences)"
    $noReference = & $comparator @common -ReportPath (Join-Path $root 'out/no-reference.json') -DetailPath (Join-Path $root 'out/no-reference.csv')
    Check ((Get-Detail @(Import-Csv (Join-Path $root 'out/no-reference.csv')) 'ISSUE-055' 'g-refdate')[0].Difference -eq 'Legacy only: unexplained' -and $noReference.Devices.UnexplainedDifferences -eq 7) 'Reference-date justification used without a legacy reference date'

    # Same legacy and evidence dates (2026-10-25): no difference can be attributed to the dates.
    $sameDate = & $comparator @common -LegacyReferenceDate '2026-10-25' -ReportPath (Join-Path $root 'out/same-date.json') -DetailPath (Join-Path $root 'out/same-date.csv')
    $sameDetails = @(Import-Csv (Join-Path $root 'out/same-date.csv'))
    Check ((Get-Detail $sameDetails 'ISSUE-055' 'g-mismatch')[0].Difference -eq 'Legacy only: unexplained' -and (Get-Detail $sameDetails 'ISSUE-055' 'g-refdate')[0].Difference -eq 'Legacy only: unexplained' -and @($sameDetails | Where-Object Justified -EQ 'True' | Where-Object Area -EQ 'Devices').Count -eq 0) 'A difference was attributed to identical reference dates'

    # Hybrid only: the device section is omitted rather than invented.
    $hybridOnly = & $comparator -LegacyHybridPath $legacyHybrid -HybridFindingsPath $newHybrid -ReportPath (Join-Path $root 'out/hybrid.json')
    Check ($null -eq $hybridOnly.Devices -and $null -ne $hybridOnly.Hybrid) 'Partial comparison invented a device section'
    Write-Host "PASS: $script:checks readiness comparison checks (evidence-based justification, regressions, undetermined applicability, read-only inputs). Synthetic fixtures only."
} finally {
    Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue
}
