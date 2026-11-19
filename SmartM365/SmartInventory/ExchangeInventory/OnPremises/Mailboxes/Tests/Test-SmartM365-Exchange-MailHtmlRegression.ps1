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
$reportAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path (Join-Path $folder '..\..\..\..') 'SmartInventory/Common/SmartM365.ReportRenderers.ps1'),[ref]$null,[ref]$null)
$functionAst = @($reportAst.FindAll({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'New-SmartM365ExchangeLocalMailboxReportEmailBody' }, $true))
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
$rendered = New-SmartM365ExchangeLocalMailboxReportEmailBody -ReportRows $rows -DailyStatsCsv 'daily.csv' -SummaryCsv 'summary.csv' -Title 'Fixture' -LocalMailboxWithDatabaseCount 3 -ShowMailLinks:$false
# Semantic checks keep a baseline refresh from accepting wrong totals or unsafe HTML.
$plain = ($rendered -replace '<[^>]+>', ' ') -replace '\s+', ' '
$plain = [Net.WebUtility]::HtmlDecode($plain)
foreach ($expected in @('Domains 2', 'Inventory objects 10', 'Local mailboxes 3', 'distinct, with database', 'Remote mailboxes 7', 'Enabled accounts 8', 'Disabled accounts 2', '19.75 GB', 'daily.csv', 'summary.csv')) {
    if ($plain -notlike ('*' + $expected + '*')) { throw "Missing mailbox report evidence: $expected" }
}
$unsafeRows = @([pscustomobject]@{DomainName='<script>alert(1)</script>';TotalMailboxCount=0;TotalLocalMailboxCount=0;TotalRemoteMailboxCount=0;EnabledAccounts=0;DisabledAccounts=0;TotalLocalMailboxSizeGB=0})
$escaped = New-SmartM365ExchangeLocalMailboxReportEmailBody -ReportRows $unsafeRows -DailyStatsCsv 'daily.csv' -SummaryCsv 'summary.csv' -Title 'Fixture' -ShowMailLinks:$false
if ($escaped -match '<script>alert' -or $escaped -notmatch '&lt;script&gt;') { throw 'Domain name HTML escaping failed.' }
$bytes = [Text.Encoding]::UTF8.GetBytes($rendered)
$hasher = [Security.Cryptography.SHA256]::Create()
try { $hash = ([BitConverter]::ToString($hasher.ComputeHash($bytes))).Replace('-', '') } finally { $hasher.Dispose() }
$baseline = '4D326EE41C9F16F768328068B50EF841B35F6141EE891222EA760C3E593C08AA'
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
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCPUFBf6D0yEqeq
# fWSeIpDK/pdsGuXEdIsXDGq/aLwnVKCCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# ztcaoVD7a8ggHP1Vdp/rnafM4GtyCAE6b7U9Yzgvp1/a1kh7XffmqVhRRjCCBoIw
# ggRqoAMCAQICEDbCsL18Gzrno7PdNsvJdWgwDQYJKoZIhvcNAQEMBQAwgYgxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpOZXcgSmVyc2V5MRQwEgYDVQQHEwtKZXJzZXkg
# Q2l0eTEeMBwGA1UEChMVVGhlIFVTRVJUUlVTVCBOZXR3b3JrMS4wLAYDVQQDEyVV
# U0VSVHJ1c3QgUlNBIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MB4XDTIxMDMyMjAw
# MDAwMFoXDTM4MDExODIzNTk1OVowVzELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1Nl
# Y3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFt
# cGluZyBSb290IFI0NjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAIid
# 2LlFZ50d3ei5JoGaVFTAfEkFm8xaFQ/ZlBBEtEFAgXcUmanU5HYsyAhTXiDQkiUv
# pVdYqZ1uYoZEMgtHES1l1Cc6HaqZzEbOOp6YiTx63ywTon434aXVydmhx7Dx4IBr
# Aou7hNGsKioIBPy5GMN7KmgYmuu4f92sKKjbxqohUSfjk1mJlAjthgF7Hjx4vvyV
# DQGsd5KarLW5d73E3ThobSkob2SL48LpUR/O627pDchxll+bTSv1gASn/hp6IuHJ
# orEu6EopoB1CNFp/+HpTXeNARXUmdRMKbnXWflq+/g36NJXB35ZvxQw6zid61qmr
# lD/IbKJA6COw/8lFSPQwBP1ityZdwuCysCKZ9ZjczMqbUcLFyq6KdOpuzVDR3ZUw
# xDKL1wCAxgL2Mpz7eZbrb/JWXiOcNzDpQsmwGQ6Stw8tTCqPumhLRPb7YkzM8/6N
# nWH3T9ClmcGSF22LEyJYNWCHrQqYubNeKolzqUbCqhSqmr/UdUeb49zYHr7ALL8b
# AJyPDmubNqMtuaobKASBqP84uhqcRY/pjnYd+V5/dcu9ieERjiRKKsxCG1t6tG9o
# j7liwPddXEcYGOUiWLm742st50jGwTzxbMpepmOP1mLnJskvZaN5e45NuzAHteOR
# lsSuDt5t4BBRCJL+5EZnnw0ezntk9R8QJyAkL6/bAgMBAAGjggEWMIIBEjAfBgNV
# HSMEGDAWgBRTeb9aqitKz1SA4dibwJ3ysgNmyzAdBgNVHQ4EFgQU9ndq3T/9ARP/
# FqFsggIv0Ao9FCUwDgYDVR0PAQH/BAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wEwYD
# VR0lBAwwCgYIKwYBBQUHAwgwEQYDVR0gBAowCDAGBgRVHSAAMFAGA1UdHwRJMEcw
# RaBDoEGGP2h0dHA6Ly9jcmwudXNlcnRydXN0LmNvbS9VU0VSVHJ1c3RSU0FDZXJ0
# aWZpY2F0aW9uQXV0aG9yaXR5LmNybDA1BggrBgEFBQcBAQQpMCcwJQYIKwYBBQUH
# MAGGGWh0dHA6Ly9vY3NwLnVzZXJ0cnVzdC5jb20wDQYJKoZIhvcNAQEMBQADggIB
# AA6+ZUHtaES45aHF1BGH5Lc7JYzrftrIF5Ht2PFDxKKFOct/awAEWgHQMVHol9ZL
# Syd/pYMbaC0IZ+XBW9xhdkkmUV/KbUOiL7g98M/yzRyqUOZ1/IY7Ay0YbMniIibJ
# rPcgFp73WDnRDKtVutShPSZQZAdtFwXnuiWl8eFARK3PmLqEm9UsVX+55DbVIz33
# Mbhba0HUTEYv3yJ1fwKGxPBsP/MgTECimh7eXomvMm0/GPxX2uhwCcs/YLxDnBdV
# VlxvDjHjO1cuwbOpkiJGHmLXXVNbsdXUC2xBrq9fLrfe8IBsA4hopwsCj8hTuwKX
# JlSTrZcPRVSccP5i9U28gZ7OMzoJGlxZ5384OKm0r568Mo9TYrqzKeKZgFo0fj2/
# 0iHbj55hc20jfxvK3mQi+H7xpbzxZOFGm/yVQkpo+ffv5gdhp+hv1GDsvJOtJinJ
# mgGbBFZIThbqI+MHvAmMmkfb3fTxmSkop2mSJL1Y2x/955S29Gu0gSJIkc3z30vU
# /iXrMpWx2tS7UVfVP+5tKuzGtgkP7d/doqDrLF1u6Ci3TpjAZdeLLlRQZm867eVe
# XED58LXd1Dk6UvaAhvmWYXoiLz4JA5gPBcz7J311uahxCweNxE+xxxR3kT0WKzAS
# o5G/PyDez6NHdIUKBeE3jDPs2ACc6CkJ1Sji4PKWVT0/MIIGpzCCBI+gAwIBAgIR
# AJCsCHIg/cWnxGtcxw33PQYwDQYJKoZIhvcNAQEMBQAwVzELMAkGA1UEBhMCR0Ix
# GDAWBgNVBAoTD1NlY3RpZ28gTGltaXRlZDEuMCwGA1UEAxMlU2VjdGlnbyBQdWJs
# aWMgVGltZSBTdGFtcGluZyBSb290IFI0NjAeFw0yNjAzMjUwMDAwMDBaFw00MTAz
# MjQyMzU5NTlaMFUxCzAJBgNVBAYTAkdCMRgwFgYDVQQKEw9TZWN0aWdvIExpbWl0
# ZWQxLDAqBgNVBAMTI1NlY3RpZ28gUHVibGljIFRpbWUgU3RhbXBpbmcgQ0EgUjQx
# MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAruRKogGtghxi+WYtW5oD
# zPDVF8GGSfgbUKh6bONxi0wvrI1S8qbAYfvLr/ky5ILVRg//70pgNKq8xC3/WEQo
# djEwAP2hmkGShoNAUQps4kd6Wwp74Fo7RlwQ1Mp949ytpWQDvCsYBbZccDmBAJC/
# ggqiuL/c805fGcMw6TzIgyBWuUx5PGp9YnheSNPXFzaz0MPtREdZYk4WhtM+hazq
# asMWVpj0WUAcNhN9vO/FAdWy9Gafdb7lmYLDKTTYjwqAY9P9RfixPPjUaJH6mnBS
# NBdrX7a0Qdlux0ApS0fc48RW1m+W3tq3HiHzch1FHyhiLzCNjc6MUpcV5xalBvPO
# w/FtQo/AxaJOvPCSsVrx0f/WkMpEm3fvVbrY9+oo9rIKv9ducE6VGfwIAtKYedG0
# bO4Ba1MmlxPcErDqjLwggvrBJu73fwXpkhtE0hzV0psgm2vhQs3pHll9N00SHBdy
# 2qndEcNuDh+46XouM2hoXCO533YQQOHPEUnMTWOo3hyxx5kjDE5PVqp+x+HS4VAT
# +WBMG4GzeLr9YvZbU5x5YvLdcR1dErV/QRYK55rp019fZFF2NR+TkSW0WcmQ3b5t
# aGcrXg49EpzKM6/mEpnSJXg1E13X6GO29rWs/LNvkGzsS8XGoRCGBls6ruofeebS
# sHADR3GeIE5gIU927bjokLECAwEAAaOCAW4wggFqMB8GA1UdIwQYMBaAFPZ3at0/
# /QET/xahbIICL9AKPRQlMB0GA1UdDgQWBBQ6dKUMZ8ZCUML9tfzHuyk0gvR6uTAO
# BgNVHQ8BAf8EBAMCAYYwEgYDVR0TAQH/BAgwBgEB/wIBADATBgNVHSUEDDAKBggr
# BgEFBQcDCDAjBgNVHSAEHDAaMAgGBmeBDAEEAjAOBgwrBgEEAbIxAQIBAwgwTAYD
# VR0fBEUwQzBBoD+gPYY7aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVi
# bGljVGltZVN0YW1waW5nUm9vdFI0Ni5jcmwwfAYIKwYBBQUHAQEEcDBuMEcGCCsG
# AQUFBzAChjtodHRwOi8vY3J0LnNlY3RpZ28uY29tL1NlY3RpZ29QdWJsaWNUaW1l
# U3RhbXBpbmdSb290UjQ2LnA3YzAjBggrBgEFBQcwAYYXaHR0cDovL29jc3Auc2Vj
# dGlnby5jb20wDQYJKoZIhvcNAQEMBQADggIBADLeUkdm8Z4DZvjfKHOhqu+hsdXN
# t+X5F+48PB6PTAJCRgA3qxxO3YbAV7baps4K/+2WWoWUspBkT4T2NXK49NvJAfCs
# ztHSgtqkAQMLX4KjCypJF/+m2Ktrk993g+gcgKdv1yg9C3JmYdJCnnL0ga+pZ/Wo
# 1+rtXZ8dnwO8RCstTN6gYX0ElFi7Y7NpxbdBC1S6bc05V/SA9HC/ojj33W6Gdwnp
# U/iVylSkdkoHtHeGIhQLT2ZH0qPM9Wdce8v2fZsDCJQQJ8rll7OGLDbsXa2CLf0M
# RN9TwzifQ3rEuAXOx/TkzkZRFfwL34hf1XqSmaYq2tTMy2LgsPrqC2Z/6ZKb3fgr
# zU0vphB4wSTWulitY/KlxbvoyKvrBvUCCx4sgeqf8aR65CbvM5MN/d/lahfXipU2
# NlY0cXcnGS61XpmeGKd8It92/lufApZR9x6o5qMJWe0jq4JsfGMGDpIKx7FzkB8g
# aejuBUW/CJ9Phc40+xJRonvVewn4S9yJVRWeM47irGbR9YlN3xruM/yZzhk+rAm9
# AW06nv7ob6RQkAXR+cTxiAPy620FF41NrViYB4UyKpzfx7x8jh4ubTOMz954YIdq
# yeiqqtsbBwXjWLP0dfMUPA3iIPnPdBKGnodGJTdSlPAMmKJdyvTPqmOXs/LMnf+2
# Za0Z6FXsIB9z9aXLMIIG4jCCBMqgAwIBAgIRAOdO8lWwUE/626bf9/yLoxUwDQYJ
# KoZIhvcNAQEMBQAwVTELMAkGA1UEBhMCR0IxGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBDQSBS
# NDEwHhcNMjYwMzI1MDAwMDAwWhcNMzcwNjI0MjM1OTU5WjByMQswCQYDVQQGEwJH
# QjEXMBUGA1UECBMOR3JlYXRlciBMb25kb24xGDAWBgNVBAoTD1NlY3RpZ28gTGlt
# aXRlZDEwMC4GA1UEAxMnU2VjdGlnbyBQdWJsaWMgVGltZSBTdGFtcGluZyBTaWdu
# ZXIgUjM3MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAsv/DbUvcUNlF
# LQURd9m4+1St5+JudFKo5P803Iks4mFeNB9SymodP6BJJWBuNhOFQj9w77AVAeg5
# qQpA2dIwp2QTyBHr2h9eWSTkMBVj9mV6+WI5SaW+vDZW7PhJTbysd9v9WB3Xt6ql
# Ei8m47pcTy8+k/OfhziKiuzNQXqfC7KcoRD/6up8OZBsU0qxr7n5nh/iRfAp1QXF
# TBQONBZSGIdHAyVRYYX033VoC8v71rizEKCpH97Pxbwcn9eq9K7W8h5v4npsMUoq
# CS/c8mQwylDQGx15dHYV6NlcVFdjXD11l7qCrIy/unH5OlZtgx58QJRXRbGgQyBd
# STpEpwuj3i5Qc52Z9m7hd7yCGCXKujf83hUQpOPx1w8+84EbEUTHVAfq4cpORaGW
# gY8NJy6txmd3wpS1MeXrOaVAMczTgzAZ+yZBWIqdgQBgTxEeXldEToZOrRkxvn1I
# jIlfr4I4NWJz+Rb52FshLVnkA/wdoad789Eb7XZDNKd4oMmnc636TgauaaVZP2LL
# oU0JD/fYr53hwBn4uXu5ZsSfpnqAT60S7szJm/Na882xEoyRzLJ+UVbXOlHLO63D
# KkAtdz1CDuwWxgRE1drnwplepT06dz+1yTr5p1AkUz21bzE6cT/8/kjh4OPzggYY
# qrOBQPfuKEL5ZJPcN9jRgEpYvRlq5ucCAwEAAaOCAY4wggGKMB8GA1UdIwQYMBaA
# FDp0pQxnxkJQwv21/Me7KTSC9Hq5MB0GA1UdDgQWBBRhEOl6Eq9RxIXU8s+kdA9Q
# zSCv+DAOBgNVHQ8BAf8EBAMCBsAwDAYDVR0TAQH/BAIwADAWBgNVHSUBAf8EDDAK
# BggrBgEFBQcDCDBKBgNVHSAEQzBBMAgGBmeBDAEEAjA1BgwrBgEEAbIxAQIBAwgw
# JTAjBggrBgEFBQcCARYXaHR0cHM6Ly9zZWN0aWdvLmNvbS9DUFMwSgYDVR0fBEMw
# QTA/oD2gO4Y5aHR0cDovL2NybC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGlt
# ZVN0YW1waW5nQ0FSNDEuY3JsMHoGCCsGAQUFBwEBBG4wbDBFBggrBgEFBQcwAoY5
# aHR0cDovL2NydC5zZWN0aWdvLmNvbS9TZWN0aWdvUHVibGljVGltZVN0YW1waW5n
# Q0FSNDEuY3J0MCMGCCsGAQUFBzABhhdodHRwOi8vb2NzcC5zZWN0aWdvLmNvbTAN
# BgkqhkiG9w0BAQwFAAOCAgEAA+o9jdGszfoZepOmygef1OlbkjrPd2QW9z3M8vVb
# QSCruPeO2eRsC9GhZ4CMZfhkrixayYD67gQkbyiRCbJu5L/i0NQjlQhBvbWfiEba
# +KHFKGud5YHRWhDZUtDeMIJGZG0BD7/sftZUo2Ifk+CXi/ZlM50+xK3OkqeXVi5G
# ubDD/5txmYuqCT3T3LAilmoB+5th9sQxiMhyQuT3R/aYb4vypoZJLYklUzTalXle
# W1nV9s4UROlE389CHDKAi/fepRSMnV8TghODDQxwzNGrOJZ04k/yhzHHDupfHPU5
# 1FYJqXIvWq9SAAWdlNV1JGIxhkp/TAtxBwz/Vd/VbgVb2d9/wRFfxFkka39O0+4x
# aZSl/oEK/1DqjxjJRO2Se9lGlJDScu21Zd23Cys3aYyB8y5H/+DFWtVe8PMKgr+V
# uIDp0Rk5bneVDAEW0TPAT8Ufwl2F6DJiDg/KZk5NmsYES+CxvF7bnISEnQh0ZrWn
# AJixquV0mElUx01wA5TuPIgyodxzNq/fC0hen9LBtdnfFfSZ+wt8A1Injsbio+DH
# Vq1voYiVNpBfO7+nh9NB4AhRXNldPgr3zgjJ+47s0uNYy2iDXAZSlkP3ym/7gy31
# jlu989SNpRWO14/LUNV2LSuXkRI1iLTPI6ZdXG0DnPPG7UftF0tk5m6BP9eNfr2t
# j1sxggW7MIIFtwIBATBiME4xHjAcBgNVBAMMFXdvcmtwbGFjZWNsb3VkaHViLmNv
# bTEsMCoGCSqGSIb3DQEJARYdY29udGFjdEB3b3JrcGxhY2VjbG91ZGh1Yi5jb20C
# EB5u7zvHOOGzRDjI8QMf4mgwDQYJYIZIAWUDBAIBBQCggYQwGAYKKwYBBAGCNwIB
# DDEKMAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEE
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgBt2UCruK5KTb
# iTVAkBsjKalg/H7bfHBNU1cbWloSvicwDQYJKoZIhvcNAQEBBQAEggGAdjbd3Y1e
# dTNCvbtbWOJ4/gAnRDctP0j9SuzXIRhUisxnsbrkGxjQZ5XzBEdAnq/Jxojlej6Z
# l+iJMH6AUO6MXCnfs/eRtxcIzLrDVRCKvDgmATPo4JSQ6Gmlez1as+l9jPGRKha4
# HhpXPKWVOOYpjOti3/AOGN+JEf01DuesHjJENu5Q8YxlxJYVzonWRKHU2wU/pvhP
# zdLLjqFetqKRgJUPQiibdJa1bI+WBktj+ZroVzCHrR5UD3sQzANIbRctvI4gc4He
# 8C1iB0JrLxKIOhNTW1CPw0i07Vg+Q+A0XJEUm9JRxb9wo/XPwDLJhz0zc1D16fuk
# +Wstn0TSeeXojlnQ/yYxlbwR2uP2tTZ75sPNtzivmmjdd6lcTUspdQWQtD1Y8kcB
# QS8+LDLzSokFKUf1hZp3NIKT3AHtDsyupOln9zXPvUB73UDl8vh5pFsFmdQv0cgC
# bE0ZYGxenaWKcGut+9RFA+iZ2WR3AbJeYKD1kUuFKxxC9TklPadITBrAoYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTE5MDkxNTI0WjA/BgkqhkiG9w0BCQQxMgQwRCHXyuLlWuhPmaoY2VIvmgu+
# Uclr+eSzbgTz6tpnht/eeY27HNEVzNvtp8lu5rusMA0GCSqGSIb3DQEBAQUABIIC
# ADFvu//vLxF+s0wmv4Sq7Q3HR1Eyikz+UG/Ejrd3T0e/WYA126UIQ7/fa0VZJHIC
# d9Va+YOHQXdiUpa510k6WTd/ISAM6UJq4uHU7TXsoAPjRQDP6MpvuWFu5hcxNUtv
# Ki0qW7d0Oq14T9FaJwXZh0h9PBlU+XssKlyEL22JjbMIxAiW5F/ih/zpqdSES/m2
# bX7ciLy16kdIA2tFy/7B7W2/nXvIHEeI0UMmmPFalUtNECktpwSRC6IVK7dtv0mR
# Rw0hMH16kyI4no8rYhFwMAPujVGOIc2vTy44lb3R9N6yhZaZTZhL6S6YmVZnBuQc
# lgc9LPo7Vb79iGVmN3nUmFARl4tq33tajpamc7kvfbZI+R5yI4ZxtjGSYw+LsqtF
# 0xB+i7+dl8zjGn9uPsZoV7dLjCpax0KrvOemSKbBlRrzvxt3HMICO3KSLs9cuGNS
# BTZ8XE2iep7fsiAsblcLU1zqAMnZapxokiy7+Q/TxpGI71sUdd9vy6/FcyaId0BP
# Vm9TUGwfJGkLrQzre/k+bYJ06XTkLGGvs6T4VaW+IevU6iEeZlBTRwcLaHBXLOad
# hr5a9OTbesTaGOBc3mLMFduPJZ1aeWQqIpQ+jNbCUXls+rMSwH9VU6Xmbs29fa6X
# LAmw1gAHy0pe5ZeM9Vi68b4qvarcQ9eKoPDwzI1WMe6n
# SIG # End signature block
