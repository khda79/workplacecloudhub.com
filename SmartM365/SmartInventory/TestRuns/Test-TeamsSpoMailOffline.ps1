#Requires -Version 7.0
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
Import-Module ImportExcel -ErrorAction Stop

function Import-ScriptFunction {
    param([string]$Path,[string[]]$Names)
    $tokens = $null
    $errors = $null
    $ast = [Management.Automation.Language.Parser]::ParseFile($Path,[ref]$tokens,[ref]$errors)
    if (@($errors).Count -gt 0) { throw "Parser errors in $Path" }
    foreach ($name in $Names) {
        $definition = $ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name},$true)
        if (-not $definition) { throw "Missing function: $name" }
        Invoke-Expression ("function global:{0} {1}" -f $name,$definition.Body.Extent.Text)
    }
}
function Assert-Equal {
    param($Actual,$Expected,[string]$Message)
    if ($Actual -ne $Expected) { throw "$Message. Expected '$Expected'; got '$Actual'." }
}

$inventoryRoot = Split-Path $PSScriptRoot -Parent
$teamsPath = Join-Path $inventoryRoot 'M365Inventory/Teams/SmartM365-Teams-Inventory.ps1'
$spoPath = Join-Path $inventoryRoot 'M365Inventory/SharePoint/SmartM365-SPO-Inventory.ps1'
Import-ScriptFunction -Path $teamsPath -Names @('IsoUtc','Get-TeamsCsvColumnNames','New-TeamsTimestampedWorkbook','Add-TeamsWorkbookOverview','Get-TeamsMailStatus','ConvertTo-HtmlReport')
Import-ScriptFunction -Path $spoPath -Names @('ConvertTo-SpoHtml','Get-SpoCsvColumnNames','New-SpoTimestampedWorkbook','Add-SpoWorkbookOverview','Get-SpoMailStatus','New-SpoHtmlSummary')

$teamsSummary = @{TotalTeams=769;OwnerlessTeams=60;HighQuotaTeams=2;InactiveTeams=100;ActiveTeams=600;ArchivedTeams=69;PublicTeams=20;TeamsWithGuests=50;CriticalCount=62;WarningCount=384}
Assert-Equal (Get-TeamsMailStatus -Summary $teamsSummary -PartialInventory $false) 'Warning' 'Localized Teams findings'
$teamsSummary.OwnerlessTeams = 77
Assert-Equal (Get-TeamsMailStatus -Summary $teamsSummary -PartialInventory $false) 'Critical' 'Systemic Teams findings'
Assert-Equal (Get-TeamsMailStatus -Summary $teamsSummary -PartialInventory $true) 'Warning' 'Partial Teams inventory'
$teamsSummary.OwnerlessTeams = 60

$spoSummary = [ordered]@{SitesProcessed=1915;OwnerlessSites=5;HighQuotaSites=3;InactiveSites=789;SharePointSites=1915;OneDriveSites=0;ListsProcessed=200;StorageUsedGB=500;TenantStorageUtilizationPercent='';CriticalAlerts=8;WarningAlerts=789;RunId='synthetic-run';Duration='00:01:00';InventoryMode='GraphOnly'}
Assert-Equal (Get-SpoMailStatus -Summary $spoSummary -PartialInventory $false) 'Warning' 'Localized SPO findings'
$spoSummary.CriticalAlerts = 0
$spoSummary.WarningAlerts = 0
$spoSummary.TenantStorageUtilizationPercent = 82
Assert-Equal (Get-SpoMailStatus -Summary $spoSummary -PartialInventory $false) 'Warning' 'Measured tenant capacity warning'
$spoSummary.CriticalAlerts = 8
$spoSummary.WarningAlerts = 789
$spoSummary.TenantStorageUtilizationPercent = 95
Assert-Equal (Get-SpoMailStatus -Summary $spoSummary -PartialInventory $false) 'Critical' 'Measured tenant capacity'
$spoSummary.TenantStorageUtilizationPercent = ''

$teamsHtml = ConvertTo-HtmlReport -AlertRows @([pscustomobject]@{Status='Critical'}) -Summary $teamsSummary -Worst 'Warning' -Started (Get-Date).AddMinutes(-1) -Ended (Get-Date)
if ($teamsHtml -notmatch 'Without owner' -or $teamsHtml -match 'Critical and warning findings|Timestamped exports|<a\s') { throw 'Teams mail body is not concise or contains a data link.' }
if ($teamsHtml.IndexOf('Run details') -lt $teamsHtml.IndexOf('Global summary')) { throw 'Teams run details must follow the summary.' }
$teamsHtmlWithLinks = ConvertTo-HtmlReport -AlertRows @() -Summary $teamsSummary -Worst 'OK' -Started (Get-Date).AddMinutes(-1) -Ended (Get-Date) -FileLinksHtml '<a href="https://example.invalid/export">Export</a>'
if ($teamsHtmlWithLinks -notmatch 'https://example.invalid/export') { throw 'Teams mail links cannot be enabled.' }
$spoHtml = New-SpoHtmlSummary -WorstStatus 'Warning' -Alerts @([pscustomobject]@{Severity='Critical'}) -Summary $spoSummary
if ($spoHtml -notmatch 'Without valid owner' -or $spoHtml -match 'Top 20 largest sites|Timestamped exports|<a\s') { throw 'SPO mail body is not concise or contains a data link.' }
if ($spoHtml.IndexOf('Run details') -lt $spoHtml.IndexOf('Global summary')) { throw 'SPO run details must follow the summary.' }
$spoHtmlWithLinks = New-SpoHtmlSummary -WorstStatus 'OK' -Alerts @() -Summary $spoSummary -FileLinksHtml '<a href="https://example.invalid/export">Export</a>'
if ($spoHtmlWithLinks -notmatch 'https://example.invalid/export') { throw 'SPO mail links cannot be enabled.' }

$temporaryRoot = Join-Path ([IO.Path]::GetTempPath()) ("SmartM365-MailOffline-{0}" -f [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporaryRoot -Force | Out-Null
try {
    $teamsFiles = @()
    foreach ($name in @('Teams','Members','Channels','Guests')) {
        $path = Join-Path $temporaryRoot "$name.csv"
        [pscustomobject]@{Name='Synthetic';Status='OK'} | Export-Csv -LiteralPath $path -NoTypeInformation
        $teamsFiles += [pscustomobject]@{Path=$path;WorksheetName=$name;TableName="TestTeams$name"}
    }
    $teamsBook = Join-Path $temporaryRoot 'Teams.xlsx'
    New-TeamsTimestampedWorkbook -CsvFiles $teamsFiles -Path $teamsBook | Out-Null
    Add-TeamsWorkbookOverview -Path $teamsBook -Summary $teamsSummary -Alerts @([pscustomobject]@{Status='Critical';TeamDisplayName='Synthetic';Check='Owners';NumericValue='0';TextValue='';Threshold='1';Details='No owner'}) -Status 'Warning'
    $book = Open-ExcelPackage -Path $teamsBook
    try { Assert-Equal $book.Workbook.Worksheets.Count 6 'Teams worksheet count'; Assert-Equal $book.Workbook.Worksheets['Findings'].Cells[2,1].Text 'Critical' 'Teams findings worksheet' }
    finally { Close-ExcelPackage -ExcelPackage $book -NoSave }
    $teamsEmptyBook = Join-Path $temporaryRoot 'Teams-empty.xlsx'
    New-TeamsTimestampedWorkbook -CsvFiles $teamsFiles -Path $teamsEmptyBook | Out-Null
    Add-TeamsWorkbookOverview -Path $teamsEmptyBook -Summary $teamsSummary -Alerts @() -Status 'OK'
    $book = Open-ExcelPackage -Path $teamsEmptyBook
    try { Assert-Equal $book.Workbook.Worksheets['Findings'].Cells[1,1].Text 'Status' 'Empty Teams findings worksheet header' }
    finally { Close-ExcelPackage -ExcelPackage $book -NoSave }

    $spoFiles = @()
    foreach ($name in @('Sites','Lists','Permissions','External sharing','Tenant capacity')) {
        $path = Join-Path $temporaryRoot (($name -replace ' ','_') + '.csv')
        [pscustomobject]@{Name='Synthetic';Status='OK'} | Export-Csv -LiteralPath $path -NoTypeInformation
        $spoFiles += [pscustomobject]@{Path=$path;WorksheetName=$name;TableName=('TestSpo' + ($name -replace ' ',''))}
    }
    $spoBook = Join-Path $temporaryRoot 'SPO.xlsx'
    New-SpoTimestampedWorkbook -CsvFiles $spoFiles -Path $spoBook | Out-Null
    Add-SpoWorkbookOverview -Path $spoBook -Summary $spoSummary -Alerts @([pscustomobject]@{Severity='Critical';Category='Ownership';SiteUrl='https://example.invalid';ObjectName='';Metric='Owner';Value='';Threshold='Valid owner';Details='No owner'}) -Status 'Warning'
    $book = Open-ExcelPackage -Path $spoBook
    try { Assert-Equal $book.Workbook.Worksheets.Count 7 'SPO worksheet count'; Assert-Equal $book.Workbook.Worksheets['Findings'].Cells[2,1].Text 'Critical' 'SPO findings worksheet' }
    finally { Close-ExcelPackage -ExcelPackage $book -NoSave }
    $spoEmptyBook = Join-Path $temporaryRoot 'SPO-empty.xlsx'
    New-SpoTimestampedWorkbook -CsvFiles $spoFiles -Path $spoEmptyBook | Out-Null
    Add-SpoWorkbookOverview -Path $spoEmptyBook -Summary $spoSummary -Alerts @() -Status 'OK'
    $book = Open-ExcelPackage -Path $spoEmptyBook
    try { Assert-Equal $book.Workbook.Worksheets['Findings'].Cells[1,1].Text 'Severity' 'Empty SPO findings worksheet header' }
    finally { Close-ExcelPackage -ExcelPackage $book -NoSave }
}
finally {
    $resolvedTemporaryRoot = [IO.Path]::GetFullPath($temporaryRoot)
    $resolvedTempParent = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\','/') + [IO.Path]::DirectorySeparatorChar
    if ($resolvedTemporaryRoot.StartsWith($resolvedTempParent,[StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolvedTemporaryRoot) -like 'SmartM365-MailOffline-*') {
        Remove-Item -LiteralPath $resolvedTemporaryRoot -Recurse -Force -ErrorAction SilentlyContinue
    }
}
Write-Host 'Teams/SPO mail and workbook offline checks passed.'

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCDfXU0EANaaB9n
# DwFAxJudUEqGjEhoMY5jugmC1tPA9qCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDInd9llcphdCcGgYtRizwtfMdsEbkl2SygrdGWt+mPzMA0GCSqG
# SIb3DQEBAQUABIIBgHEa0NaRDjXxPSotHkCD2MnXxdgOjE9/0qntOJMiIrzZdK/e
# F2fGLaSbPleFM3yYWBMxS8KWt3sI7KDYp3qf71VRAgD22y9KdXMmsvnDwYUm5jLp
# s4brDA8wV5WxZGiV6t0GteeLTG9VVx6ynXALwumj4fktn8eTsaY7UjahT7rvPEU4
# zeBjUy8TOyruO2dt7S+hKitHbYtzGAhYUSb5CuTDy3MoEFdUTGdQgDaAxGPWBUDC
# 1Jsg/govpFktLQ82RkFp2UZ5DuKHI/JtdEB5+0j9FoSgx81XswRJfExaLFjAnhwp
# 2gXM3zzs/c2eJqZe8tOaVnaf1uTX9ZTb3hjXLj3ZhIkbhI85kCHXI9kcuwuOtyMU
# 8fPOZAiVxijhojUJxVUS+LeTdbvAGIS74scPWwgELtNEZl9BQ2/ftMpRnGHtdT/Z
# EAO/69kMFdOyxh7Cc3A32gBjnWB22OHre8xnGOE6D0tRUmAWp1xYkAtvBQn/KAHb
# poQqvzU57G4FYq8AxaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDYwNzQ1
# MTVaMC8GCSqGSIb3DQEJBDEiBCCWYR0VegMoLe+f4hfbpeIgn9f7K10jg9O61P2E
# b/kuEjANBgkqhkiG9w0BAQEFAASCAgAVJ7Vfp4NnzG6zaL3BbVtFUwtHuK3neMmt
# rLJHEQN70cZdehd/ysiMl0Zup+bfAGK/mqVn6Ca68DL5+Drj8AKn0AGAW+Wzc9YT
# 565eoG6Har3QUdTD7Vp1uWVdKSJxfc7z8nFi6HWDE5qxKRblhdHArrGIEikEt7Kz
# WAyXD7eq7C53U6AF+L3ChQgmRQfCRPbm6bXwyzsDWu3YKoLJQU61Kg/MD+ITeaYy
# fuayN0XQzUvIikHk5f8ColRDHcL49AByZWsDHAR1XJPGTPiL71doxc4apaMbQGUR
# 1wxgWvXzYoodbDNvjyQUGeenkJmwvUNdoHaiznMXUluxt4iaCLh3xhuNJRiAKA4n
# u84obVzF9YLOFxaBsJbyNFpsIQLpwgiHvm1n914pAer66UShsiHd9UWWiGu6G8jX
# FSFC4uUJKfrR9r8uU+/vc5VTCm8LemQ9umBatGaf+QxTgnJOKDyhDBdbgoUcOiTN
# tXUy84HOEbFk91ElWWDDa6V9HDO5RAf5A1BoeSvwpOMgpZSsl2HyA+mIBIXLYkTj
# PUeLVwT/PwRA1q8Jjw4sAkMkfbsjXttkzoq3sX5gguyXLUcSHsPe3350q8VIhkg9
# /QEItq/LpsY5zD2AQp7o5XPGzQupWB8PsgASR6fGg9Q6oq5l2/LotdQ7/oTn6L1H
# q1uDzH5zXw==
# SIG # End signature block
