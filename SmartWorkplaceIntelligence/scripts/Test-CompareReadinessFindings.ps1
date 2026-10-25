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

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDsxFgSnojW9Jmy
# OJ/AJ4wBS9pStvzNDBGR3uzsHisqCqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIP12UevW9Jt9dh8/H2nBj4D4iE8p/irCX1YZzdm5PVmOMA0GCSqG
# SIb3DQEBAQUABIIBgFByFuCwvvuk9pjLP355NR89+vG2NFxfT4F4sLNgEzW1oKdv
# xZN2PJR0CBd2mL51WT+Jdo7nK8I7gfzwIPdYKebztSoFk4mUhnVg9KJ4USrWCNdR
# DSERphWNwz9KUM4iY+ht/z9zZWrm1xf+pJ81UMnbYPQ1n9XrHzwFRUCn2F331PEq
# cAnyk6y/bKVBDAlymAte1A+tvhhG8n+dHMw4kIkNd65TWqWI7Og0MtmCMyyXa9js
# q6WieAwCJfflAfmNJsyzCDRqEsbV4U0TpAHklRSjevdyM3GgAcaLtkrPeJFmOLcn
# EFEvf0QONn5oxqXXTljVlvagHaExm9on+Tm0MFHgupgbmjTwzP4Cj2RwSrDIe/DD
# 3Ss5HyOsRZMHyNhnTAb144lB3j1gVYgdE1s0/6wuSgo1Ce/vBuoOa41v3UjgECvB
# fgp9JwZcK2mS3zx+/tpXS2XZ+Q8SLXx81Vy2XrURfcIExzr0FxahKhcjX+STJNzd
# MP4hXjNAKVGBime1L6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjUxMjU3
# NTVaMC8GCSqGSIb3DQEJBDEiBCCHPecLuAAssz/P45hqnDCCFsYXFjFpF6VYWK7c
# wK5H6TANBgkqhkiG9w0BAQEFAASCAgAQEJbyBTPWvnhDxkLBCBM9ntCVqWjFgTzU
# Gt0j8bX1cnesIUl9jtrdqOc/tcrQiYJML0X7fldAq3e42/DhG+2bD375rvSXl030
# H5V+msjwST6KJA7yxcO/t+5X9WGD+HcQAIGwyXActNmi+bObR3TLAVaO/axb3eyI
# xWzyrXbk+0KpcSxYGVZXzD9fHdWZerb8zcBieaevn8uLhpgVeTUM3ZPAY+41pVYl
# Cgwql/CGv/URgqC0C1aJuWjqzKuf0DT5uBdtb9hlAjuIjrr+c6tOw8TgKKHycFsV
# ALuCOXe/wHUAKiqiTQh/y5w/RtwKO74xEoRuqNxmXws+pysxJecljakU6MRIXGba
# jIJDzT/ExfskorHiDZfdz5K9UyBkEPQ/DTQ+gKh28skX9jthqeZ4fBTkmzooeZgg
# fG8FEUcPcOGhVqjyG+DHgGDxuKlwjiixQpZuRqCqsg3A6HN91XOF1MPJQZnl6Npk
# 6i52zaq5OpqLkFnDj+fzlUzkYmQ5ip12CXMYjSXmEflSIVrv7tT0Jh1Mw/QM01eZ
# YYkbgcVt1rEGPyTZt/hZ1xOFilO8QT8rYXVYwILaUadin3oJxMorL80ENb6viv6O
# k+9NCoS1Nj/4UOxmp8HRMT9aKKTYZMIRaQF5Re3PblzhV68cuASCR3dnZXQgd6Z+
# jxkgt6p7ZQ==
# SIG # End signature block
