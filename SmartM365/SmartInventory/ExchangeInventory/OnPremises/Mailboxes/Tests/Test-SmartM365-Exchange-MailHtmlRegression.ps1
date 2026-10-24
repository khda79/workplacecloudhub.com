[CmdletBinding()]
param([switch]$RecordBaseline, [string]$ExchangeSourcePath, [string]$ProxyRenderPath)
# Freeze fixture formatting, not the production renderer's locale.
$savedCulture = [Threading.Thread]::CurrentThread.CurrentCulture
try {
[Threading.Thread]::CurrentThread.CurrentCulture = [Globalization.CultureInfo]::InvariantCulture
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$folder = Split-Path $PSScriptRoot -Parent
$scriptPath = if ($ExchangeSourcePath) { $ExchangeSourcePath } else { Join-Path $folder 'SmartM365-Exchange-Local-Mailboxes-Inventory.ps1' }
$tokens = $null
$errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($scriptPath, [ref]$tokens, [ref]$errors)
if (@($errors).Count -gt 0) { throw 'Exchange collector does not parse.' }
$functionAst = @($ast.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'New-SmartM365ExchangeLocalMailboxReportEmailBody' }, $true))
if ($functionAst.Count -ne 1) { throw 'Exchange mail renderer was not found.' }
. ([scriptblock]::Create($functionAst[0].Extent.Text))

function ConvertTo-SmartM365EmailHtmlText { param($Value) return [System.Net.WebUtility]::HtmlEncode([string]$Value) }
$corePath = Join-Path $folder '..\..\..\..\Modules\SmartM365.Core\Compatibility\WindowsPowerShell5\SmartM365-WindowsPowerShell5.psm1'
$coreTokens = $null
$coreErrors = $null
$coreAst = [Management.Automation.Language.Parser]::ParseFile($corePath, [ref]$coreTokens, [ref]$coreErrors)
$gridAst = @($coreAst.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'New-SmartM365EmailKpiGridHtml' }, $true))
if ($gridAst.Count -ne 1) { throw 'Shared Exchange card grid helper was not found.' }
. ([scriptblock]::Create($gridAst[0].Extent.Text))
$bodyAst = @($coreAst.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'New-SmartM365EmailBody' }, $true))
if ($bodyAst.Count -ne 1) { throw 'Shared Exchange email frame was not found.' }
. ([scriptblock]::Create($bodyAst[0].Extent.Text.Replace('function New-SmartM365EmailBody {','function New-SmartM365EmailBodyFixture {')))
function Get-SmartM365MailboxReportCsvRowCount { param($Path) return 2 }
function New-SmartM365LocalMailboxIssueEmailSection { param($Issues, $IssuesCsvPath) return $null }
function New-SmartM365EmailBody { param($Title, $Category, $Severity, $Tenant, $Message, $Sections, $Footer) return New-SmartM365EmailBodyFixture -Title $Title -Category $Category -Severity $Severity -Tenant $Tenant -Message $Message -Sections $Sections -Footer $Footer -HostName 'fixture-host' -GeneratedAt '2026-10-09 00:00:00 +02:00' }

$Tenant = 'fixture'
$rows = @(
    [pscustomobject]@{ DomainName='a.example'; TotalMailboxCount=4; TotalLocalMailboxCount=1; TotalRemoteMailboxCount=3; EnabledAccounts=3; DisabledAccounts=1; TotalLocalMailboxSizeGB=12.5 },
    [pscustomobject]@{ DomainName='b.example'; TotalMailboxCount=6; TotalLocalMailboxCount=2; TotalRemoteMailboxCount=4; EnabledAccounts=5; DisabledAccounts=1; TotalLocalMailboxSizeGB=7.25 }
)
$rendered = New-SmartM365ExchangeLocalMailboxReportEmailBody -ReportRows $rows -DailyStatsCsv 'daily.csv' -SummaryCsv 'summary.csv' -Title 'Fixture' -ShowMailLinks:$false
# Semantic checks keep a baseline refresh from accepting wrong totals or unsafe HTML.
$plain = ($rendered -replace '<[^>]+>', ' ') -replace '\s+', ' '
$plain = [Net.WebUtility]::HtmlDecode($plain)
foreach ($expected in @('Domains 2', 'Total mailboxes 10', 'Local mailboxes 3', 'Remote mailboxes 7', 'Enabled accounts 8', 'Disabled accounts 2', '19.75 GB', 'daily.csv', 'summary.csv')) {
    if ($plain -notlike ('*' + $expected + '*')) { throw "Missing mailbox report evidence: $expected" }
}
$unsafeRows = @([pscustomobject]@{DomainName='<script>alert(1)</script>';TotalMailboxCount=0;TotalLocalMailboxCount=0;TotalRemoteMailboxCount=0;EnabledAccounts=0;DisabledAccounts=0;TotalLocalMailboxSizeGB=0})
$escaped = New-SmartM365ExchangeLocalMailboxReportEmailBody -ReportRows $unsafeRows -DailyStatsCsv 'daily.csv' -SummaryCsv 'summary.csv' -Title 'Fixture' -ShowMailLinks:$false
if ($escaped -match '<script>alert' -or $escaped -notmatch '&lt;script&gt;') { throw 'Domain name HTML escaping failed.' }
$bytes = [Text.Encoding]::UTF8.GetBytes($rendered)
$hasher = [Security.Cryptography.SHA256]::Create()
try { $hash = ([BitConverter]::ToString($hasher.ComputeHash($bytes))).Replace('-', '') } finally { $hasher.Dispose() }
$baseline = '4B1F6548006A16A684F21CAD8BF7A9F1D83CD543B5DF853CCF8695517CCF8EF0'
if ($RecordBaseline) { Write-Output "EXCHANGE_HTML_BASELINE=$hash"; return }
if ($hash -ne $baseline) { throw "Exchange HTML changed: expected $baseline, got $hash" }
Write-Output "PASS: Exchange HTML unchanged ($hash)."

# ProxyAddresses sections: synthetic evidence with long unbreakable addresses and statuses must stay inside the 760px frame.
Import-Module (Join-Path $folder 'SmartM365.ExchangeProxyMail.psm1') -Force
$proxyChecks = 0
function Assert-ProxyMail { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw "Proxy mail layout: $Message" }; $script:proxyChecks++ }
$longSuffix = 'very-long-synthetic-subdomain-name-for-layout.example.invalid'
$proxyDetail = @(
    foreach ($index in 1..60) {
        [pscustomobject]@{ Identity = "id$index"; Alias = ('sharedalias{0}' -f ($index % 3)); DisplayName = "Synthetic User $index"; PrimaryAddress = ('firstname.lastname.number{0}@{1}' -f $index, $longSuffix); ExpectedAddress = ('sharedalias{0}@{1}' -f ($index % 3), $longSuffix); SuggestedUniqueAddress = ('sharedalias{0}.{1}@{2}' -f ($index % 3), $index, $longSuffix); Status = 'Missing -> RemoteRoutingAddressUnavailable' }
    }
    foreach ($index in 61..70) {
        [pscustomobject]@{ Identity = "id$index"; Alias = "unique$index"; DisplayName = "Synthetic User $index"; PrimaryAddress = "unique$index@$longSuffix"; ExpectedAddress = "unique$index@$longSuffix"; SuggestedUniqueAddress = ''; Status = 'OK' }
    }
)
$proxyEvidence = [pscustomobject]@{
    Available = $true; Reason = ''; CompletedAt = [datetimeoffset]'2026-10-09T00:00:00+02:00'; Version = 'synthetic'; Mode = 'Audit'; Scope = 'Synthetic scope'; ExpectedSuffix = $longSuffix
    Summary = @([pscustomobject]@{ Summary = 'With expected address missing'; Count = 60 }, [pscustomobject]@{ Summary = 'Total recipients processed'; Count = 70 })
    Detail = $proxyDetail
    DetailPath = 'C:\Synthetic\DATA-LAST\Exchange_OnPrem_ProxyAddresses_Check.csv'; SummaryPath = 'C:\Synthetic\DATA-LAST\Exchange_OnPrem_ProxyAddresses_Summary.csv'; WorkbookPath = 'C:\Synthetic\DATA-LAST\Exchange_OnPrem_ProxyAddresses.xlsx'
    Links = @([pscustomobject]@{ FileName = 'Exchange_OnPrem_ProxyAddresses.xlsx'; WebUrl = 'https://synthetic.invalid/sites/smartm365/Exchange_OnPrem_ProxyAddresses.xlsx' })
}
foreach ($showLinks in @($false, $true)) {
    $proxySections = @(New-SmartM365ProxyMailSections -Evidence $proxyEvidence -ShowMailLinks:$showLinks)
    Assert-ProxyMail ((@($proxySections | ForEach-Object Title) -join '|') -eq 'Check ProxyAddresses Report|Proxy address summary|Top 50 duplicate aliases|Top 50 missing proxy addresses|Proxy address files') 'unexpected section list.'
    foreach ($section in $proxySections | Where-Object { $_.Html -match '<table' }) {
        foreach ($table in [regex]::Matches($section.Html, '<table[^>]*>')) {
            Assert-ProxyMail ($table.Value -match 'table-layout:fixed' -and $table.Value -match 'width:712px' -and $table.Value -match 'border:1px solid #d9e2ec') "$($section.Title) table is not a fixed 712px shared-style table."
        }
        $headerWidths = @([regex]::Matches($section.Html, '<th [^>]*width="(\d+)"') | ForEach-Object { [int]$_.Groups[1].Value })
        Assert-ProxyMail ($headerWidths.Count -gt 0 -and ($headerWidths | Measure-Object -Sum).Sum + (16 * $headerWidths.Count) -eq 712) "$($section.Title) column widths do not total 712px."
        Assert-ProxyMail (@([regex]::Matches($section.Html, '<th [^>]*background:#f8fafc[^>]*text-transform:uppercase')).Count -eq $headerWidths.Count) "$($section.Title) headers do not use the shared style."
        foreach ($cell in [regex]::Matches($section.Html, '<td [^>]*>(?:(?!</td>).)*@(?:(?!</td>).)*</td>')) {
            Assert-ProxyMail ($cell.Value -match 'word-break:break-all' -and $cell.Value -match 'overflow-wrap:anywhere') "$($section.Title) address cell cannot wrap."
        }
        foreach ($cell in [regex]::Matches($section.Html, '<td [^>]*>(?:(?!</td>).)*-&gt; (?:(?!</td>).)*</td>')) {
            Assert-ProxyMail ($cell.Value -match 'word-break:break-all') "$($section.Title) status cell cannot wrap."
        }
    }
    $duplicateHtml = @($proxySections | Where-Object Title -eq 'Top 50 duplicate aliases')[0].Html
    Assert-ProxyMail (([regex]::Matches($duplicateHtml, '<tr>')).Count -eq 51) 'duplicate preview is not limited to 50 rows.'
    Assert-ProxyMail ($duplicateHtml -notmatch '>Display name</th>' -and $duplicateHtml -notmatch 'Synthetic User') 'duplicate preview still shows display names.'
    $missingHtml = @($proxySections | Where-Object Title -eq 'Top 50 missing proxy addresses')[0].Html
    Assert-ProxyMail (([regex]::Matches($missingHtml, '<tr>')).Count -eq 51 -and $missingHtml -match '>Display name</th>') 'missing preview changed.'
    $filesHtml = @($proxySections | Where-Object Title -eq 'Proxy address files')[0].Html
    $expectedLinks = if ($showLinks) { 1 } else { 0 }
    Assert-ProxyMail (([regex]::Matches($filesHtml, '<a href=')).Count -eq $expectedLinks -and ([regex]::Matches($filesHtml, '<tr>')).Count -eq 4) 'file links do not follow ShowMailLinks.'
}
$proxyRendered = New-SmartM365ExchangeLocalMailboxReportEmailBody -ReportRows $rows -DailyStatsCsv 'daily.csv' -SummaryCsv 'summary.csv' -Title 'Fixture' -ShowMailLinks:$false -ProxyEvidence $proxyEvidence
$declaredWidths = @([regex]::Matches($proxyRendered, '(?:width="|width:)(\d+)(?:px)?') | ForEach-Object { [int]$_.Groups[1].Value })
Assert-ProxyMail (@($declaredWidths | Where-Object { $_ -gt 760 }).Count -eq 0) 'an element is declared wider than 760px.'
Assert-ProxyMail ($proxyRendered -match 'width:760px;max-width:760px') 'the 760px frame is missing.'
if ($ProxyRenderPath) { [IO.File]::WriteAllText([IO.Path]::GetFullPath($ProxyRenderPath), $proxyRendered, (New-Object Text.UTF8Encoding($false))); Write-Output "Proxy mail render written to $ProxyRenderPath" }
Write-Output "PASS: $proxyChecks ProxyAddresses mail layout checks (synthetic data)."

} finally { [Threading.Thread]::CurrentThread.CurrentCulture = $savedCulture }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCZ6Nnwwu699+wC
# gdbJ78sSn+BuWtTnQJaNt/F2bBDZXaCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIJrMS8DKn8/o2DBOoEAo0PfWEaTPHaE3sWEJZD7Ce5KoMA0GCSqG
# SIb3DQEBAQUABIIBgK0efXQ2Hn06z5s/s+XyZ1Sg8FFugim2vWrLm0uCBCr/JaWX
# CGxk7bm7pUbNzvLCBrHWPXnos2OWIgHYollRLo7XkTPGa4sWnJA8xeZi6+0zTH5/
# GMaSje9rTwTOdvr9kGx+KyORaQilvWs9ZocveMyGCVwFVPCI2eoDD8jMoxr/iGH8
# 9MrseQ+6d0UEQWcbiVilt9guqtdcsijfNuxyuEN48tNW3wLmKutAu9sHIZYtG3Tb
# k+UKbVDVZTEpDz80loBlBEnUrs6e5ttONaOylz+wGWlF2ml9zwJJDn7jgoGRhmMM
# bwaqqx0HFq8d+Z9Stzjoq7GJE1WIRN9vGukLcKunj4OTBuPOtGK0yFuuMgQUTYjX
# dMUNX/OFq/9BMR8m+Sw1rlroy7+iCA7G10+F/X6BHqeVPvIIfjYVDCjwGdHpcGye
# QyusHHw9x9uTJ0d9i2Smo8qVnpD0Ch8P891J/91lcKUIE1aX4qpd/U9U616FE+gi
# cNo9rLkTVWGA1FLx96GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjQxNDIy
# NTBaMC8GCSqGSIb3DQEJBDEiBCAsa68+M1o6tBUW++LBhgxXPy6C9g2sAP7J2PxO
# hOFdiDANBgkqhkiG9w0BAQEFAASCAgBMNGzeIlQJdd3URLgMhjj3ikEn9wBbAe6h
# PcMTo1n9w9h5mPsuHUGgpPciLKOzd3TK0UT6l39Utcmb7O0YT6UneB7LIY7VItpw
# 4oDIYf0YfpulRYI3wJKzJCRTxFEimta7E0hlQqrbDeB1ML38qUoequAR4v0dbAnb
# usjtYdIPPUWznd3QMA3t33NaQeIvNzT85ecuWqOpf+i+cacYvmSMdow+b/OZMeS7
# jq+JG7FngfS5JY0qXEcC0UuqUSidwTU5LpQvW0vjvhG95VbcqtOM92DMq0YHZIds
# IKq17TMkh05k/M9/QlC9UYjJrv+ui9pdY8mfPHPXo5Cn6hsVrPjXaFAL8vlEI334
# iIr5MQHsZmpsGymXEPkJrK0TWdwtkUGOuMGWR3MvV6DZ5hBXSFDFTLRaL6NwyD03
# PEp829YJ7jX8e8wpRLxkF0LSmUUT+ty53BIXlQYOpfzDZv0V5cFC2O9ieuWVGoJJ
# +ALTFarwkNSElVA8p37kSCivp0qS8//JkrgA71H5Q8mGJN9390apEy4/S/vRnixa
# TQVc+/rWI8GAeZpUSl7JiNUvPnKuOyGo1dLqTvU35IUwDLrg5du+0Tt5207Q82V/
# DlAwjaYgkZKivqSJsIQ0QGoMNETf8i8JwIi7DCqqgMvmYXGUzAPJNvRiVBy1lZ4K
# ErAhRxpYBQ==
# SIG # End signature block
