<#
.SYNOPSIS
Synthetic database email tests; extracts functions without running the collector.
.VERSION
1.0.1
#>
[CmdletBinding()]
param([string]$SourceRoot)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
if (-not $SourceRoot) { $SourceRoot = Split-Path $PSScriptRoot -Parent }
$collectorPath = Join-Path $SourceRoot 'SmartInventory/ExchangeInventory/OnPremises/ServersAndStorage/SmartM365-Exchange-OnPrem-InfrastructureAndReadiness-Inventory.ps1'
$tokens = $null; $parseErrors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile($collectorPath, [ref]$tokens, [ref]$parseErrors)
if ($parseErrors.Count) { throw "Collector parse failed: $parseErrors" }
$reportAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $SourceRoot 'SmartInventory/Common/SmartM365.ReportRenderers.ps1'),[ref]$null,[ref]$null)
$functionNames = @('Format-HtmlValue', 'ConvertTo-ServersAndStorageDatabaseSectionHtml', 'New-HtmlExecutiveSummary',
    'Import-ServersAndStorageExistingCsv', 'Invoke-ServersAndStorageHtmlReportRegeneration', 'Send-ServersAndStorageHtmlReport',
    'Get-ObjectPropertyValue', 'ConvertTo-ServersAndStorageEmailHtmlText', 'New-ServersAndStorageEmailLinkHtml',
    'ConvertTo-ServersAndStorageMailNumber', 'Format-ServersAndStorageMailNumber', 'Limit-ServersAndStorageMailText',
    'ConvertTo-ServersAndStorageMailDate', 'New-ServersAndStorageMailTableHtml', 'New-ServersAndStorageMailNoteHtml',
    'Get-ServersAndStorageMailFile', 'New-ServersAndStorageMailBodyHtml')
$definitions = foreach ($name in $functionNames) {
    $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    if (-not $node) {
        $node = $reportAst.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
    }
    if ($null -eq $node) { throw "Missing function: $name" }
    $node.Extent.Text
}
# Shared mail template functions, exposed under the -Prefix Core names used by the collector.
$coreAst = [Management.Automation.Language.Parser]::ParseFile((Join-Path $SourceRoot 'Modules/SmartM365.Core/SmartM365.Core.psm1'), [ref]$null, [ref]$null)
$definitions = @($definitions) + @(foreach ($name in @('ConvertTo-SmartM365EmailHtmlText', 'New-SmartM365EmailKpiGridHtml', 'New-SmartM365EmailTableHtml', 'New-SmartM365EmailBody')) {
    $coreAst.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true).Extent.Text
    'Set-Alias -Name {0} -Value {1}' -f ($name -replace '-SmartM365', '-CoreSmartM365'), $name
})
# Never import the runtime Core module or load tenant configuration. All external
# operations reachable from regeneration/mail are replaced inside this module.
$module = New-Module -ScriptBlock ([scriptblock]::Create(($definitions -join "`n")))
& $module {
    Set-StrictMode -Version Latest
    $script:CapturedMail = $null
    function script:Import-SmartM365CoreModule { }
    function script:Write-Log { param($Message) }
    function script:Invoke-ServersAndStorageSharePointUpload { param($LocalFilePath) }
    function script:Get-SmartM365EffectiveConfigValue { param($Name, $DefaultValue) return $DefaultValue }
    function script:CoreSendEmailHtmlReport {
        param($From, $To, $Subject, $BodyHtml, $Attachments)
        $script:CapturedMail = [pscustomobject]@{ Body=$BodyHtml; Attachments=$Attachments }
    }
    function script:Get-MailboxDatabase { throw 'A synthetic report must not collect Exchange data.' }
}
$script:Checks = 0
function Assert-DatabaseMail {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
    $script:Checks++
}
function New-ReadinessRow {
    param([string]$Name, [string]$Setting, [AllowNull()][object]$Value, [string]$Status = 'OK', [string]$Category = 'MailboxDatabase')
    [pscustomobject]@{Category=$Category; ObjectName=$Name; Setting=$Setting; Value=$Value; CollectionStatus=$Status; Importance='Review'; MigrationFocus='Synthetic fixture'; ErrorMessage=''}
}
function Get-DatabaseSection {
    param([AllowNull()][object[]]$Rows)
    & $module { param($r,$s) ConvertTo-ServersAndStorageDatabaseSectionHtml -ReadinessInventory $r -Summary $s } $Rows $summary
}
function Assert-Counters {
    param([string]$Html, [string]$Total, [string]$Mounted, [string]$NotMounted, [string]$Unknown)
    foreach ($pair in @(@('Total databases observed',$Total), @('Mounted',$Mounted), @('Not mounted',$NotMounted), @('Unknown status',$Unknown))) {
        Assert-DatabaseMail ($Html -cmatch ([regex]::Escape($pair[0]) + ': <strong>' + [regex]::Escape($pair[1]) + '</strong>')) "Wrong counter: $($pair[0]) = $($pair[1])"
    }
}

$summary = [pscustomobject]@{ExecutionDate='2020-01-02 03:04:05';RunId='SYNTHETIC-OLD-RUN';ExchangeServersCount=1;TotalLogicalProcessorCount=4;TotalMemoryGB=16;TotalDiskDriveCount=1;TotalDiskDriveSizeTB=1;ExchangeSchemaRangeUpper='N/A';ExchangeOrgObjectVersion='N/A';ExcludedServers='EX-EXCLUDED'}
$server = [pscustomobject]@{ExchangeServerName='EX-01';ServerRole='Mailbox';LogicalProcessorCount=4;MemoryGB=16;DiskDriveCount=1;DiskDriveTotalSizeGB=1024;ComputeCollectionStatus='OK';DiskDriveCollectionStatus='OK';LogicalDiskCollectionStatus='OK';LowSpaceLogicalDiskCount=0}
$testRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-DatabaseMail-' + [guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)
try {
    $rows = @(
        New-ReadinessRow 'DB-A' 'Mounted' $true
        New-ReadinessRow 'DB-A' 'Server' 'EX-01'
        New-ReadinessRow 'DB-A' 'DatabaseSize' '10 GB (10,737,418,240 bytes)'
        New-ReadinessRow 'DB-A' 'AvailableNewMailboxSpace' '0 B (0 bytes)'
        New-ReadinessRow 'DB-B' 'Mounted' $false
        New-ReadinessRow 'DB-B' 'Server' 'EX-EXCLUDED'
        New-ReadinessRow 'DB-C' 'Mounted' ''
        New-ReadinessRow 'DB-D' 'Mounted' 'not-a-boolean'
        New-ReadinessRow 'DB-D' 'Recovery' 'True'
        New-ReadinessRow 'DAG-A' 'Servers' 'EX-01; EX-02' -Category 'DatabaseAvailabilityGroup'
        New-ReadinessRow 'DAG-B' 'Servers' 'EX-03; EX-04' -Category 'DatabaseAvailabilityGroup'
    )
    $html = Get-DatabaseSection $rows
    Assert-Counters $html 4 1 1 2
    Assert-DatabaseMail ($html -match '<td>DB-B</td><td>EX-EXCLUDED</td><td[^>]*>Not mounted</td>') 'Boolean false or excluded server association lost.'
    Assert-DatabaseMail ($html -match '<td>10 GB \(10,737,418,240 bytes\)</td><td>0 B \(0 bytes\)</td>') 'Raw Exchange sizes or measured zero lost.'
    Assert-DatabaseMail ($html -match 'DatabaseSize: missing\.' -and $html -match 'Mounted: unrecognized value\.') 'Missing values and invalid boolean not disclosed.'
    Assert-DatabaseMail ($html -notmatch '<td>DAG-A</td>|<td>DAG-B</td>') 'DAGs counted as databases.'
    foreach ($notice in @('not free disk space', 'detailed DAG copy states and per-copy servers are not collected', 'Mailboxes per database: N/A', 'Zero observed mailboxes would not prove an empty database', 'SYNTHETIC-OLD-RUN', '2020-01-02 03:04:05', 'Exchange_OnPrem_MigrationReadiness_Config.csv')) {
        Assert-DatabaseMail ($html.Contains($notice)) "Missing provenance or limitation: $notice"
    }

    $csvPath = Join-Path $testRoot 'Exchange_OnPrem_MigrationReadiness_Config.csv'
    $rows | Export-Csv -LiteralPath $csvPath -Delimiter ';' -NoTypeInformation -Encoding UTF8
    $roundTrip = @(Import-Csv -LiteralPath $csvPath -Delimiter ';')
    Assert-DatabaseMail ((Get-DatabaseSection $roundTrip) -eq $html) 'Boolean values changed after CSV serialization.'
    Assert-Counters (Get-DatabaseSection @(New-ReadinessRow 'DB-ONLY' 'Mounted' 'fAlSe')) 1 0 1 0
    Assert-Counters (Get-DatabaseSection @(New-ReadinessRow 'DB-ONLY' 'Server' 'EX-01')) 1 0 0 1

    foreach ($empty in @(@(), @($null), @(New-ReadinessRow 'DAG' 'Servers' 'EX-01' -Category 'DatabaseAvailabilityGroup'))) {
        $absent = Get-DatabaseSection $empty
        Assert-Counters $absent 'N/A' 'N/A' 'N/A' 'N/A'
        Assert-DatabaseMail ($absent.Contains('this does not prove that there are no databases')) 'Absent evidence represented as zero.'
    }
    $failed = @(New-ReadinessRow 'MailboxDatabase' 'Collection' 'Failed' -Status 'ERROR' -Category 'CollectorError')
    $failedHtml = Get-DatabaseSection $failed
    Assert-Counters $failedHtml 'N/A' 'N/A' 'N/A' 'N/A'
    Assert-DatabaseMail ($failedHtml.Contains('coverage is incomplete')) 'Failed collection not disclosed.'
    $partial = Get-DatabaseSection ($rows + $failed)
    Assert-Counters $partial 4 1 1 2
    Assert-DatabaseMail ($partial.Contains('observed counts are not estate totals')) 'Partial counts presented as estate totals.'

    $duplicates = Get-DatabaseSection @(
        New-ReadinessRow 'DB-A' 'Mounted' 'True'
        New-ReadinessRow 'db-a' 'Mounted' 'False'
        New-ReadinessRow 'DB-A' 'Server' 'EX-01'
        New-ReadinessRow 'DB-A' 'Server' 'EX-02'
    )
    Assert-Counters $duplicates 1 0 0 1
    Assert-DatabaseMail ($duplicates.Contains('Mounted: ambiguous duplicate records.') -and $duplicates.Contains('Server: ambiguous duplicate records.')) 'Ambiguous records silently picked.'
    $invalid = Get-DatabaseSection @(
        New-ReadinessRow 'DB-A' 'Mounted' 'True' -Status 'ERROR'
        New-ReadinessRow 'DB-A' 'DatabaseSize' '20 GB' -Status 'WARNING'
        New-ReadinessRow '' 'Mounted' 'True'
    )
    Assert-Counters $invalid 1 0 0 1
    Assert-DatabaseMail ($invalid.Contains('1 record(s) have no database name') -and $invalid.Contains('DatabaseSize: collection not OK.')) 'Invalid data not disclosed.'
    Assert-DatabaseMail (-not $invalid.Contains('<td>20 GB</td>')) 'Failed measurement rendered as valid.'
    $escaped = Get-DatabaseSection @(New-ReadinessRow '<script>DB</script>' 'Server' 'EX<&>')
    Assert-DatabaseMail ($escaped.Contains('&lt;script&gt;DB&lt;/script&gt;') -and $escaped.Contains('EX&lt;&amp;&gt;') -and -not $escaped.Contains('<script>')) 'Database names/server HTML not escaped.'

    # Exercise the actual regeneration and mail body paths on a historical fixture.
    $summary | Export-Csv -LiteralPath (Join-Path $testRoot 'Exchange_OnPrem_Servers_Inventory_Summary.csv') -Delimiter ';' -NoTypeInformation -Encoding UTF8
    $server | Export-Csv -LiteralPath (Join-Path $testRoot 'Exchange_OnPrem_Infrastructure_PerServerSummary.csv') -Delimiter ';' -NoTypeInformation -Encoding UTF8
    & $module { param($root) Invoke-ServersAndStorageHtmlReportRegeneration -RunFolder $root -SendMail } $testRoot
    $report = Get-Content -LiteralPath (Join-Path $testRoot 'Exchange_OnPrem_InfrastructureAndReadiness_Report.html') -Raw
    $mail = & $module { $script:CapturedMail }
    Assert-DatabaseMail ($report.Contains($html)) 'Database summary missing from the regenerated full report.'
    Assert-DatabaseMail ($mail.Body.Contains('SmartM365EmailTemplate:v1') -and -not $mail.Body.Contains($html) -and $null -eq $mail.Attachments) 'Email is not the shared summary template.'
    foreach ($fileName in @('Exchange_OnPrem_InfrastructureAndReadiness_Report.html', 'Exchange_OnPrem_Servers_Inventory_Summary.csv', 'Exchange_OnPrem_Infrastructure_PerServerSummary.csv', 'Exchange_OnPrem_MigrationReadiness_Config.csv')) {
        Assert-DatabaseMail ($mail.Body.Contains($fileName)) "Email Files section does not link $fileName."
    }
    Assert-DatabaseMail ($report.Contains('Excluded</td>') -and $report.Contains('Per-server infrastructure summary') -and $report.Contains('Exchange readiness configuration')) 'Existing report sections regressed.'
    Assert-Counters $report 4 1 1 2
    Write-Host "PASS: $script:Checks synthetic Exchange database mail checks (PowerShell $($PSVersionTable.PSVersion)). No collector or real mail executed."
}
finally {
    Remove-Module $module -Force
    $resolved = [IO.Path]::GetFullPath($testRoot)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (-not $resolved.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'SmartM365-DatabaseMail-*') { throw 'Unsafe test cleanup path.' }
    Remove-Item -LiteralPath $resolved -Recurse -Force
}

# SIG # Begin signature block
# MIIfPQYJKoZIhvcNAQcCoIIfLjCCHyoCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCB+ePcGB8mArHQH
# A99/bKqDjHnzSIdTT2mUKX25JwSLV6CCGNgwggS9MIIDJaADAgECAhAebu87xzjh
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
# AYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQg1O7i4I9S7NZM
# pRySuxz4Roph8F09/Xm/0M0tR4hUbRcwDQYJKoZIhvcNAQEBBQAEggGAofm+iFku
# YtLfnXPSqszNElalBqLm1NgKT5FVERybNfSVpShLbv7YZarob2R9dATWGkDcYrsh
# NaBpYg4BvItTW4I12Vee+ATVGYBsxNDA7ylinghY1qcw2R+LNZpnKobihb7Bmj5q
# c99rH4XRIf5QxDR3o/hp9IGsaXHIkqf1HSL7xKsKzlxGhRfFT2DKdrQaPoYgkVIy
# UMGk9LHoWWtYt4X4acaT9VmSU0MH/nANKJkLbAm6fka1xMpDcmCVEl07/KBb9WSV
# xqCpjLVaGCUv71N4X03ntmRoC+9O1gDVK8annDBxWfWvzeyNXP1vH0pDjnjqmfSM
# j2w2Btpt0BBvYz83cVDvQcZURqquKPRrX1hOgEmLkDJGttf1Bpg/vkJrMkCdTXfi
# u5NY4UGL1acJe7WKQauYt6HblsClNwiOJymZ745AomZ6NlSH98ly88CIizd97okH
# hwmqkbpr+FLdlZw7LTQCYZd9HrjYvcp7OAvo8AhVbjl4N2PJ6YngTVeRoYIDIzCC
# Ax8GCSqGSIb3DQEJBjGCAxAwggMMAgEBMGowVTELMAkGA1UEBhMCR0IxGDAWBgNV
# BAoTD1NlY3RpZ28gTGltaXRlZDEsMCoGA1UEAxMjU2VjdGlnbyBQdWJsaWMgVGlt
# ZSBTdGFtcGluZyBDQSBSNDECEQDnTvJVsFBP+tum3/f8i6MVMA0GCWCGSAFlAwQC
# AgUAoHkwGAYJKoZIhvcNAQkDMQsGCSqGSIb3DQEHATAcBgkqhkiG9w0BCQUxDxcN
# MjYxMTE5MDkxNTMwWjA/BgkqhkiG9w0BCQQxMgQwTQE6cFLOdvY4RWqWVcCK7qCe
# iI1ysr2TOLx2Rnz5B+ImIZz0cu5tdIZKimLhqqjhMA0GCSqGSIb3DQEBAQUABIIC
# AGCvBwRMtUURMx6lPmtsn0JHeFibbKPebE837AevM9prhIzcSnVAff3/q+RtYlQI
# AAhYOap4M48+XeMBAkIgY48jf4cMK5k6KYOY+HDfa10viq7ndyK9b+8p25ZUv+G7
# T7kbvCm3u4gZVOvkaJZyKC7Z6uGCNI0E9XpU3aDLqxJ+PXSXWSBvE9HoKtG3oeJU
# TUD6Q88hA3Owe0Mwt0NCs1qvkhBdgBxNUpHOI6eMTBreAhCroPwndzGG+K3sor65
# qzJDsZL4QMtfvKBgAprEHeHCO3eRYwHOAejmR2OXrqBvg+/twbl9AtHZxZhIOA4g
# ohaoDgN6MkkKhuANa65qZtckp8A3h/ZxCpGlYah92pmYs/PZUpBrTIErTbbBr00v
# bQr2kdVTJ8d+NK/VXWpYMuVLy4S+3nVr8TYhMkuTI0Yr6HYk2QJUj9uNxSVqt71K
# 12otzKBHU8Tukj2nRJGqtaB3SucIZ3yMgV6Hbt1b5XhBHzLCPBYxxjiWAXDlz/Wh
# EKdYLA0WylaUGYetPH8kTEe41oUKDlQwLk9yTWWy7ko28OSs8pB3eG2cdhvyWtZs
# n0Q4cjuLlavX9lJX1k9AMDwef6OjsQ4FVHnQh2MP3vGvBvQl6a03PIl3FJnSvbje
# iOFdTJHwdz0c884Oc1Ixn02eY2+4r7X8Rkww+fewBtS1
# SIG # End signature block
