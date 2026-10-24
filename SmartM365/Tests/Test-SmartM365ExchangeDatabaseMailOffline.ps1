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
$functionNames = @('Format-HtmlValue', 'ConvertTo-ServersAndStorageDatabaseSectionHtml', 'New-HtmlExecutiveSummary',
    'Import-ServersAndStorageExistingCsv', 'Invoke-ServersAndStorageHtmlReportRegeneration', 'Send-ServersAndStorageHtmlReport',
    'Get-ObjectPropertyValue', 'ConvertTo-ServersAndStorageEmailHtmlText', 'New-ServersAndStorageEmailLinkHtml',
    'ConvertTo-ServersAndStorageMailNumber', 'Format-ServersAndStorageMailNumber', 'Limit-ServersAndStorageMailText',
    'ConvertTo-ServersAndStorageMailDate', 'New-ServersAndStorageMailTableHtml', 'New-ServersAndStorageMailNoteHtml',
    'Get-ServersAndStorageMailFile', 'New-ServersAndStorageMailBodyHtml')
$definitions = foreach ($name in $functionNames) {
    $node = $ast.Find({ param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name }, $true)
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
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAJfKEcGI0gZoHX
# TrEK1pXcKPhVguU0L0vKKMl3CPDq6aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIka78F+mIe22S4rmfJyq+fxG1G0FW4KeJlBQ0qkl598MA0GCSqG
# SIb3DQEBAQUABIIBgC703K3CegczQod3mQf9Bni5wzJW2yl/e6rqpnyqXoTsk5b7
# +rERZRBGrw72ou4TiBfvZKjqe7ijvCbIXyhh17MDQeDOqKC5feW3F9FDVow3kWfW
# sI+bvJbtMPCTHLsJJGQ+iDr8rWZFiRAOMWb9suxFlE3cxnYkYepWthonBxKNN/1F
# CPp4cqNiOuvKp51VSZNVSAFjk1FPpsQhCDTwfIfCjX4J1aO3MGSKJQhvQCm1mNqs
# bJEqDR098VJWTAXj4yYwfKHIQ4XLhGF+r8wtaL/jV5l0dcEFXbQiFwL6V5L4qAbv
# 7RR3Qnv+NRAuccclQkpQwycQMptOSZshY7JsKvqc1BfyVUWQUBY+7gcSIsmH1YfM
# xEm+fA2t9Ah6lPBoQ1gaR/cPJkJRtedvh+3w669pbXhKnDRGlP1Ezl2mSJHsUzU2
# 1Q8hgJAQAFQg307kaYYuMHCMfpBy70/GRZlSq2XKSB6FKCxDmJcDyui2wPPwBe2b
# fTfuxHoBX7UaYeMZdKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjQxMzU5
# MzdaMC8GCSqGSIb3DQEJBDEiBCDLxTXTU3H4K6MBv9gMdZA0P3eXs5rnYUxO8r0X
# C0i7DTANBgkqhkiG9w0BAQEFAASCAgArpR4otWvaBWlTwS45xA5o2ITqur5YHsZC
# M9O1/ygpRpuNBrjgxus+0uOZtsLQXIkzrMc8l/TKSEpwjShr11ATJAwu8d9W3X3M
# 9RCl1rlpVRD8zi+Mq3DfX9JSq+rwakjI3d1qNspnScweRsgqrAuC0ySDRZK/OfbQ
# bu+mYkOFfmjj/B6+2JTP7q5vbHemJsk5vBh9oWlw50jY23FsGP2b0MOfXFXqbHDn
# VnIy+3pmKtT7a8WB6DR+9AyUjrEWiOgF0F7w0dWOy/wpEtP2c18yggPrhftsP+uK
# CY0CrJIpUp0SqCUDzOwbDP4lLz6vVeO6V9QFEdSLI8UShMZr+hzYYOMqJBQWxy3S
# bDbiZvcwqy+NSQVOtv5yIOX0Kb22EXlw1sWcI/yQluOOMOaD40DBxR6K7GcEcFSS
# +cgl5AKuzueVAbtqJ9M3fbiagVxK/MsIvIlNVY0bdCnPQYME33Mr4uM5yLhPCNwm
# eknBNc6XAD3oFkOLjRdloWM2XyOXYxattKJ9YS+KczEhmJZv5lMckdgvN+DGJk4K
# lxh2IVRTibGbzbnQuczv1LqudzIHURvg+abydAYwbRNDfReCTdeqP8ynH8DaobgH
# Pd40UIQPtVpFBX36cXiN4s6Rxqj4SrKcNx/sEAB7qrhJ18YCkteR65AUsF7Ux7An
# lDm5IekHJA==
# SIG # End signature block
