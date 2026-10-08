<#
.SYNOPSIS
  Offline checks for failed license assignments in the mail and Power BI source.
#>

[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$inventoryPath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\SmartM365-Licences-Inventory.ps1')).Path
$tokens = $null
$parseErrors = $null
$ast = [System.Management.Automation.Language.Parser]::ParseFile($inventoryPath,[ref]$tokens,[ref]$parseErrors)
if ($parseErrors.Count) { throw ($parseErrors | Out-String) }
$corePath = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..\..\Modules\SmartM365.Core\SmartM365.Core.psd1')).Path
Import-Module -Name $corePath -Force -ErrorAction Stop
$names = @('Get-LicensesCsvSource','Import-LicensesSourceCsv','Read-LicensesIndexedSource','Get-LicensesAssignmentIssues','Get-LicensesAlertSummary','New-LicensesRecoveryWorkbook','Publish-LicensesReportCsv')
foreach ($name in $names) {
  $definition = $ast.Find({param($node) $node -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name},$true)
  if (-not $definition) { throw "Missing function: $name" }
  Invoke-Expression $definition.Extent.Text
}
function Assert-Equal($Actual,$Expected,[string]$Label) {
  if ($Actual -ne $Expected) { throw "$Label expected '$Expected', got '$Actual'." }
}
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-AssignmentIssues-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root -Force | Out-Null
try {
  $skuF1 = '11111111-1111-1111-1111-111111111111'
  $skuF3 = '33333333-3333-3333-3333-333333333333'
  $now = [datetimeoffset]::UtcNow
  $tenant = @(
    [pscustomobject]@{Id=$skuF1;TenantSkuPartNumber='M365_F1_COMM'},
    [pscustomobject]@{Id=$skuF3;TenantSkuPartNumber='SPE_F1'}
  )
  $paths = @(
    [pscustomobject]@{TenantKey='test';UserId='u1';SkuId=$skuF3;AssignedByGroupId='g1';AssignmentRoute='Group';AssignmentState='Error';AssignmentError='CountViolation';LastUpdatedDateTime=$now.ToString('o')},
    [pscustomobject]@{TenantKey='test';UserId='u1';SkuId=$skuF3;AssignedByGroupId='g2';AssignmentRoute='Group';AssignmentState='Error';AssignmentError='CountViolation';LastUpdatedDateTime=$now.ToString('o')},
    [pscustomobject]@{TenantKey='test';UserId='u2';SkuId=$skuF3;AssignedByGroupId='g1';AssignmentRoute='Group';AssignmentState='Error';AssignmentError='UniquenessViolation';LastUpdatedDateTime=$now.ToString('o')},
    [pscustomobject]@{TenantKey='test';UserId='u3';SkuId=$skuF3;AssignedByGroupId='g1';AssignmentRoute='Group';AssignmentState='Error';AssignmentError='CountViolation';LastUpdatedDateTime=$now.ToString('o')},
    [pscustomobject]@{TenantKey='test';UserId='u4';SkuId=$skuF1;AssignedByGroupId='g1';AssignmentRoute='Group';AssignmentState='Error';AssignmentError='CountViolation';LastUpdatedDateTime=$now.ToString('o')},
    [pscustomobject]@{TenantKey='test';UserId='u5';SkuId=$skuF3;AssignedByGroupId='';AssignmentRoute='Direct';AssignmentState='ActiveWithError';AssignmentError='Other';LastUpdatedDateTime=$now.ToString('o')}
  )
  $effective = @(
    [pscustomobject]@{TenantKey='test';UserId='u3';SkuPartNumber='SPE_F1'},
    [pscustomobject]@{TenantKey='test';UserId='u4';SkuPartNumber='M365_F1_COMM'}
  )
  $active = @(1..5 | ForEach-Object {
    [pscustomobject]@{TenantKey='test';'Object Id'="u$_";'User principal name'="u$_@example.invalid";'Display name'=if ($_ -eq 2) { '=2+2' } else { "User $_" };AccountEnabled='True';UserType='Member'}
  })
  $groups = @(
    [pscustomobject]@{TenantKey='test';GroupId='g1';DisplayName='Frontline'},
    [pscustomobject]@{TenantKey='test';GroupId='g2';DisplayName='Regional'}
  )
  $data = @(
    @{Name='M365_Licenses_AssignmentPaths.csv';Rows=$paths},
    @{Name='M365_Licenses_Users.csv';Rows=$effective},
    @{Name='M365_Users_Active.csv';Rows=$active},
    @{Name='M365_EntraGroups_All.csv';Rows=$groups}
  )
  foreach ($item in $data) { $item.Rows | Export-Csv -LiteralPath (Join-Path $root $item.Name) -NoTypeInformation }
  $licenseFiles = @($data | Where-Object { $_.Name -ne 'M365_Users_Active.csv' } | ForEach-Object {
    @{File=$_.Name;Status='Success';IsPartialInventory=$false;SHA256=(Get-FileHash -LiteralPath (Join-Path $root $_.Name) -Algorithm SHA256).Hash}
  })
  @{RunId='same-run';Status='Completed';IsPartialInventory=$false;Files=$licenseFiles} |
    ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $root 'SmartInventory_SmartM365-Licences-Inventory.current.json.txt')
  @{RunId='active-run';Status='Completed';IsPartialInventory=$false;Files=@(@{File='M365_Users_Active.csv';Status='Success';IsPartialInventory=$false;SHA256=(Get-FileHash -LiteralPath (Join-Path $root 'M365_Users_Active.csv') -Algorithm SHA256).Hash})} |
    ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $root 'SmartInventory_SmartM365-ActiveUsers-Inventory.current.json.txt')

  $issues = Get-LicensesAssignmentIssues -Folder $root -TenantKey 'test' -TenantRows $tenant -LicenseSnapshotUtc $now
  Assert-Equal $issues.Available $true 'Qualified issue list'
  Assert-Equal $issues.Counts['Microsoft 365 F3'].CapacityWait 1 'F3 users awaiting capacity'
  Assert-Equal $issues.Counts['Microsoft 365 F3'].OtherErrors 3 'F3 other error users'
  Assert-Equal $issues.Counts['Microsoft 365 F1'].CapacityWait 0 'F1 covered user is not waiting'
  Assert-Equal $issues.Counts['Microsoft 365 F1'].OtherErrors 1 'F1 covered path remains an error'
  Assert-Equal @($issues.Rows).Count 5 'Distinct issue rows'
  $waiting = @($issues.Rows | Where-Object Category -eq 'Capacity wait')
  Assert-Equal $waiting[0].UserId 'u1' 'Waiting identity'
  Assert-Equal $waiting[0].AssignmentPaths 2 'Two group paths merged'
  Assert-Equal $waiting[0].AssignedByGroupNames 'Frontline;Regional' 'Assigning group names'
  $alerts = Get-LicensesAlertSummary -AssignmentIssues $issues
  Assert-Equal $alerts.AwaitingCapacity '1' 'Top alert waiting distinct users'
  Assert-Equal $alerts.LicensingErrors '4' 'Top alert errors distinct users'
  Assert-Equal $alerts.AwaitingBreakdown 'F1 0 · F3 1 · E3 0 · E5 0' 'Top alert waiting suite breakdown'
  Assert-Equal $alerts.ErrorsBreakdown 'F1 1 · F3 3 · E3 0 · E5 0' 'Top alert error suite breakdown'
  $duplicateSuite = [pscustomobject]@{ Product='Microsoft 365 F1'; Category='Licensing error'; UserId='u2' }
  $issues.Rows = @($issues.Rows) + @($duplicateSuite)
  $issues.Counts['Microsoft 365 F1'].OtherErrors = 2
  $deduplicatedAlerts = Get-LicensesAlertSummary -AssignmentIssues $issues
  Assert-Equal $deduplicatedAlerts.LicensingErrors '4' 'Top alert deduplicates a user across suites'
  Assert-Equal $deduplicatedAlerts.ErrorsBreakdown 'F1 2 · F3 3 · E3 0 · E5 0' 'Top alert retains per-suite counts'
  $issues.Rows = @($issues.Rows | Where-Object { $_ -ne $duplicateSuite })
  $issues.Counts['Microsoft 365 F1'].OtherErrors = 1
  $unavailableAlerts = Get-LicensesAlertSummary -AssignmentIssues $null
  Assert-Equal $unavailableAlerts.AwaitingCapacity 'N/D' 'Top alert missing source is N/D'
  Assert-Equal $unavailableAlerts.LicensingErrors 'N/D' 'Top alert missing errors are N/D'

  $book = Join-Path $root 'AssignmentIssues.xlsx'
  $summary = @('Microsoft 365 F1','Microsoft 365 F3','Microsoft 365 E3','Microsoft 365 E5' | ForEach-Object { [pscustomobject]@{Product=$_} })
  [void](New-LicensesRecoveryWorkbook -Path $book -SummaryRows $summary -Usage $null -AssignmentIssues $issues -CollectedAtUtc $now.ToString('o'))
  Assert-Equal @(Import-Excel -Path $book -WorksheetName 'Awaiting licenses').Count 1 'Workbook awaiting sheet'
  Assert-Equal @(Import-Excel -Path $book -WorksheetName 'Licensing errors').Count 4 'Workbook errors sheet'
  $excelErrors = @(Import-Excel -Path $book -WorksheetName 'Licensing errors')
  Assert-Equal @($excelErrors | Where-Object UserId -eq 'u2')[0].DisplayName "'=2+2" 'Workbook formula text is escaped'
  $package = Open-ExcelPackage -Path $book -ErrorAction Stop
  try {
    $sheet = $package.Workbook.Worksheets['Awaiting licenses']
    if ($sheet.Cells[2,14].Value -isnot [datetime] -and $sheet.Cells[2,14].Value -isnot [double]) { throw 'Assignment update timestamp is not a sortable Excel date.' }
    Assert-Equal $sheet.Cells[2,14].Style.Numberformat.Format 'yyyy-mm-dd hh:mm:ss' 'Assignment update date format'
  }
  finally { $package.Dispose() }
  $bookSummary = @(Import-Excel -Path $book -WorksheetName Summary)
  Assert-Equal $bookSummary[1].AwaitingLicense 1 'Workbook F3 waiting KPI'
  Assert-Equal $bookSummary[1].LicensingErrors 3 'Workbook F3 error KPI'

  $snapshotId = [guid]::NewGuid().ToString('N')
  $snapshot = [pscustomobject]@{
    SnapshotId=$snapshotId; TenantKey='test'; GeneratedAtUtc=$now.ToString('o'); LicenseCollectedAtUtc=$now.ToString('o')
    Products=@($summary | ForEach-Object { [pscustomobject]@{Product=$_.Product;Enabled=10;Consumed=8;Subscribed=$true;UsageAvailable=$false;Counts=$null} })
    OtherProducts=@('Copilot','Dynamics 365','Power BI' | ForEach-Object { [pscustomobject]@{Product=$_;Enabled=0;Consumed=0;Subscribed=$false} })
    RecoveryCandidates=@(); DowngradeCandidates=@(); E3ToF3Review=$null; AssignmentIssues=$issues
    MailboxGap=$null; AdGapActivity=$null; Sources=@()
  }
  Publish-LicensesReportCsv -Folder $root -Snapshot $snapshot
  $publishedIssues = @(Import-Csv -LiteralPath (Join-Path $root 'M365_Licenses_ReportAssignmentIssues.csv'))
  $publishedSummary = @(Import-Csv -LiteralPath (Join-Path $root 'M365_Licenses_ReportSummary.csv'))
  Assert-Equal $publishedIssues.Count 5 'Published issue rows'
  Assert-Equal @($publishedIssues | Where-Object { $_.SnapshotId -eq $snapshotId }).Count 5 'Published issue snapshot ID'
  Assert-Equal $publishedSummary[1].AssignmentCapacityWait 1 'Published F3 waiting count'
  Assert-Equal $publishedSummary[1].AssignmentOtherErrors 3 'Published F3 error count'
  $issues.Counts['Microsoft 365 F3'].CapacityWait = 2
  $reconciliationRejected = $false
  try { Publish-LicensesReportCsv -Folder $root -Snapshot $snapshot }
  catch { $reconciliationRejected = $_.Exception.Message -like '*does not reconcile*' }
  Assert-Equal $reconciliationRejected $true 'Inconsistent assignment KPI is rejected before publication'
  $issues.Counts['Microsoft 365 F3'].CapacityWait = 1

  $paths[0].AssignmentError = 'Other'
  $paths | Export-Csv -LiteralPath (Join-Path $root 'M365_Licenses_AssignmentPaths.csv') -NoTypeInformation
  $unqualified = Get-LicensesAssignmentIssues -Folder $root -TenantKey 'test' -TenantRows $tenant -LicenseSnapshotUtc $now
  Assert-Equal $unqualified.Available $false 'Changed assignment CSV is rejected'
  Assert-Equal @($unqualified.Rows).Count 0 'No unqualified user list'
  'PASS: License assignment issue extraction and workbook offline checks.'
}
finally { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBpxOP0yhDcq+Yu
# Iv0JTKEc48r8RG/RfuFi0PtubeCBmKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIOKvLBJK5dbR/dJ9gF+P9wnIbu6QiuzBIVvAaEfsebzSMA0GCSqG
# SIb3DQEBAQUABIIBgKq3EIagaON0v5+HjtmqeMoWG/tye1o2l96yacVs08J8kGQb
# lOnry33p/DdfuOQGBenHQA5X4MXjiZyTYWkXRM8qRBDltR4eMpz2IiQyfAPRQ+PS
# 3IOQjpbXdRIOPYmeJFSBdyktqnVnvQOtbfjexq3UgqQ9YTiA+/IFIew+WWlN5++h
# dLY/nc1tn1K7btIfi+h5HFzIDflD4jM1VMmruK7eaS9f8G2acnIcl1WlnWBVlpY6
# +qrb+JnPOJQGleH2ex+p8/0LCdARv8M04c+JsgT7PRnrCdjo3GyMOpylrlUiAzSF
# +n+mbCMHdqOTFRxRuMhgAUyku6oDPwg66Ph48KlCEs3q1yQEqbP3Sm6vfKJUIrVF
# 2G6ZclTyS/rX9BEZvK7E22JGtqSuFd/1fG547D40kpRepXx19wQuQt62e9gsKzmp
# BdyDdA4rYLY76Qrx/BqklKl8+h/ygyfpGoijStN95othxs2qYEiGMSVyPPNfZSYH
# kY3WBqA7Fg2oXF4IAKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDgxMTIy
# MDhaMC8GCSqGSIb3DQEJBDEiBCBO3oOdRbnBC2vFz/2pE3LjTuElZ3NLByBnwRgw
# 0OhyoDANBgkqhkiG9w0BAQEFAASCAgAcfXmAJ60/lWTVbll5FqWt21w0jYWDSedH
# peR+vBWTr39MebT/0JBg3dGVDrb6ddCxJCLs3KTPkZSLI2hl/pJ9eJmweHnVd5I1
# FAG8hgwdI32v9SRSXEZirJ4YwOTjkNt4OKsN06PbCwQNGuAkaDLpAWWs7kBi278o
# FpOeB4FwqtPaJzBH3xfo0J3sCj1p7Ii5qJwXnn7IIugYzYrofBQ4yrlcgx0+GZqz
# zFNt0keGkJn8s27eLuZjlPeKxJAsYmgT86byonpm4D4SIZTTKqqvFaw7uPROXp/D
# tI6UG3O53p8aOJ9709kB5e9oEaYy6SIBJaEVsGOnrG1c21kR/mKMj8sReEo2/i4F
# LJxQsmIZli39Dr4WjqTzHAN55esVqvnqhO3WFB5nxZKhAOSlP9EB8fuaRSFYz6tz
# +KQ5J63XpTL0wjRVeoHCHu4E94J0r1RwX3DXR7qlhUd1e8vkFC6T1cqGp9mF1t/Y
# GDEF3oECg55UIyXUR/WHK2W0ErgKfMR0Uc/vxN5PClzasVrWAkuZuASDbANYlz3p
# GEX6QxC6GJGRrXULBo64hpY4t0e67GSVuPHJaNVrX4IL3hdO0XT+FBh3IrL7P6wF
# KUjgnMkHS2bX0KtikkjVZcXiuLTKAJBKyrLaHpn0MCFHsA+pB4f6DOvkB2Ymaykq
# b5KsbCqWqg==
# SIG # End signature block
