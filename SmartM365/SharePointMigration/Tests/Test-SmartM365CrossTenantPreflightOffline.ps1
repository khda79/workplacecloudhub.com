<#
.SYNOPSIS
    Exercise launcher identity mapping preflight with synthetic offline inventories.
.VERSION
    1.0.2
#>
#requires -Version 7.4
$ErrorActionPreference = 'Stop'
$toolkit = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-CrossTenantPreflight-' + [guid]::NewGuid().ToString('N'))
try {
    foreach ($directory in @('Migrations\Synthetic\scans', 'Migrations\Synthetic\target-scans', 'Scripts\Launchers\Generic', 'Scripts\Compare', 'Scripts\Inventory')) {
        New-Item -ItemType Directory -Path (Join-Path $root $directory) -Force | Out-Null
    }
    foreach ($file in @('SmartM365-SharePointMigration-LauncherCommon.ps1', 'SmartM365-SharePointMigration-Storage.ps1')) {
        Copy-Item -LiteralPath (Join-Path $toolkit "Scripts\Launchers\$file") -Destination (Join-Path $root 'Scripts\Launchers')
    }
    $launcher = Join-Path $root 'Scripts\Launchers\Generic\SmartM365-SharePointMigration-Launcher.ps1'
    Copy-Item -LiteralPath (Join-Path $toolkit 'Scripts\Launchers\Generic\SmartM365-SharePointMigration-Launcher.ps1') -Destination $launcher
    Get-ChildItem -LiteralPath (Join-Path $toolkit 'Scripts\Compare') -Filter '*.py' -File | Copy-Item -Destination (Join-Path $root 'Scripts\Compare')
    Copy-Item -LiteralPath (Join-Path $toolkit 'Scripts\console_lifecycle.py') -Destination (Join-Path $root 'Scripts')
    $migration = Join-Path $root 'Migrations\Synthetic'
    @'
@{
 Name='Synthetic'
 Source=@{Type='SPO'; SiteUrl='https://source.example.test/sites/a'; PermissionRootPath='/sites/a'; AuthConfigPath='source.local.psd1'}
 Target=@{Type='SPO'; SiteUrl='https://target.example.test/sites/a'; PermissionRootPath='/sites/a'; AuthConfigPath='target.local.psd1'}
 Comparison=@{CrossTenantIdentityMappingPath='identity-map.csv'; PathMappingsFile='mapping.txt'; ShareGateReplacementCharacter='_'; EntraUsersCacheEnabled=$true}
 Permissions=@{SourceDocumentLibrariesOnly=$false; TargetDocumentLibrariesOnly=$false}
 Output=@{SourcePermissionScans='scans'; TargetPermissionScans='target-scans'; PermissionComparisons='comparisons'; Logs='logs'}
}
'@ | Set-Content -LiteralPath (Join-Path $migration 'migration.config.psd1') -Encoding UTF8
    "@{ TenantId='11111111-1111-1111-1111-111111111111' }" | Set-Content -LiteralPath (Join-Path $migration 'source.local.psd1') -Encoding UTF8
    "@{ TenantId='22222222-2222-2222-2222-222222222222' }" | Set-Content -LiteralPath (Join-Path $migration 'target.local.psd1') -Encoding UTF8
    'https://source.example.test/sites/a https://target.example.test/sites/a' | Set-Content -LiteralPath (Join-Path $migration 'mapping.txt') -Encoding UTF8
    @'
param([string]$OutputPath,[string]$TenantHost,[string]$MaxCacheAgeHours,[string]$TenantId,[switch]$Connect,[switch]$InteractiveAuth,[switch]$DeviceLogin)
throw 'Cross-tenant comparison must not query the target directory cache.'
'@ | Set-Content -LiteralPath (Join-Path $root 'Scripts\Inventory\SmartM365-EntraUsers-CacheExport.ps1') -Encoding UTF8
    $inventory = Join-Path $migration 'scans\SPO-PermissionInventory-Synthetic-20261104-100000.csv'
    @'
WebUrl,ObjectScope,ObjectServerRelativeUrl,PermissionLevels,PrincipalType,PrincipalLoginName
https://source.example.test/sites/a,Web,/sites/a,Read,User,one@source.example.test
https://source.example.test/sites/a,Web,/sites/a,Read,User,two@source.example.test
'@ | Set-Content -LiteralPath $inventory -Encoding UTF8
    $mapping = Join-Path $migration 'identity-map.csv'
    @'
SourceWebUrl,SourcePrincipalType,SourcePrincipalLoginName,TargetWebUrl,TargetPrincipalType,TargetPrincipalLoginName,Status
https://source.example.test/sites/a,User,one@source.example.test,https://target.example.test/sites/a,User,one@target.example.test,Confirmed
'@ | Set-Content -LiteralPath $mapping -Encoding UTF8
    $originalHash = (Get-FileHash -LiteralPath $inventory -Algorithm SHA256).Hash
    $output = & pwsh -NoProfile -File $launcher -MigrationName Synthetic -Action ValidateCrossTenantIdentityMappings -NonInteractive 2>&1 | Out-String
    if ($LASTEXITCODE -eq 0 -or $output -notmatch 'unresolved source rows') { throw "Incomplete mapping was not rejected: $output" }
    $reports = @(Get-ChildItem -LiteralPath (Join-Path $migration 'identity-mapping-reviews') -Filter '*.csv')
    if ($reports.Count -ne 1 -or @(Import-Csv -LiteralPath $reports[0].FullName).Count -ne 1) { throw 'Unresolved evidence was not saved.' }
    Add-Content -LiteralPath $mapping -Encoding UTF8 -Value 'https://source.example.test/sites/a,User,two@source.example.test,https://target.example.test/sites/a,User,two@target.example.test,Confirmed'
    $target = Join-Path $migration 'target.csv'
    @'
WebUrl,ObjectScope,ObjectServerRelativeUrl,PermissionLevels,PrincipalType,PrincipalLoginName
https://target.example.test/sites/a,Web,/sites/a,Read,User,one@target.example.test
'@ | Set-Content -LiteralPath $target -Encoding UTF8
    $output = & pwsh -NoProfile -File $launcher -MigrationName Synthetic -Action ValidateCrossTenantIdentityMappings -TargetCsv $target -NonInteractive 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0 -or $output -notmatch 'Permissions and group membership were not compared' -or $output -notmatch '"MappedSourceRowsWithTargetPrincipalObserved": 1') { throw "Complete mapping preflight failed: $output" }
    $targetInventory = Join-Path $migration 'target-scans\SPO-PermissionInventory-Synthetic-20261104-100000.csv'
    Copy-Item -LiteralPath $target -Destination $targetInventory
    $output = & pwsh -NoProfile -File $launcher -MigrationName Synthetic -Action ComparePermissions -SourceCsv $inventory -TargetCsv $targetInventory -Force -NonInteractive 2>&1 | Out-String
    if ($LASTEXITCODE -ne 0) { throw "Cross-tenant permission comparison failed: $output" }
    $summaryPath = Get-ChildItem -LiteralPath (Join-Path $migration 'comparisons') -Recurse -Filter 'Synthetic-Summary.csv' | Where-Object { $_.DirectoryName -notmatch 'DocumentLibraries|OtherScopes' } | Select-Object -First 1
    if (-not $summaryPath) { throw 'Comparison summary was not created.' }
    $summary = Import-Csv -LiteralPath $summaryPath.FullName
    if ($summary.CrossTenantIdentityMode -ne 'ExplicitMapping' -or $summary.MatchedPermissions -ne '1' -or $summary.MissingInSPO -ne '1' -or $summary.ValidationStatus -ne 'ReviewNeeded') { throw 'Cross-tenant summary did not preserve mapped and missing assignments with review status.' }
    $mappingLines = @(Get-Content -LiteralPath $mapping)
    $mappingLines[0..1] | Set-Content -LiteralPath $mapping -Encoding UTF8
    $output = & pwsh -NoProfile -File $launcher -MigrationName Synthetic -Action ComparePermissions -SourceCsv $inventory -TargetCsv $targetInventory -Force -NonInteractive 2>&1 | Out-String
    if ($LASTEXITCODE -eq 0 -or $output -notmatch 'unresolved source rows') { throw 'Incomplete cross-tenant mapping was not blocked before comparison.' }
    $topSummaries = @(Get-ChildItem -LiteralPath (Join-Path $migration 'comparisons') -Recurse -Filter 'Synthetic-Summary.csv' | Where-Object { $_.DirectoryName -notmatch 'DocumentLibraries|OtherScopes' })
    if ($topSummaries.Count -ne 1) { throw 'Incomplete mapping published a new comparison summary.' }
    if ((Get-FileHash -LiteralPath $inventory -Algorithm SHA256).Hash -ne $originalHash) { throw 'Source inventory was modified.' }
    'PASS: unresolved and complete mapping preflight through the launcher; no tenant access.'
}
finally {
    $resolved = [IO.Path]::GetFullPath($root)
    $allowed = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if ($resolved.StartsWith($allowed, [StringComparison]::OrdinalIgnoreCase) -and (Split-Path $resolved -Leaf) -like 'SmartM365-CrossTenantPreflight-*') {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD5IK+rg+/jElTl
# tHUCGOL5ph2pWkCWS/wvWueUXLIMlqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIAvr6RiVPOTAubYXPzLgf7R4t5+hTFOOLBAuSSiZGq8CMA0GCSqG
# SIb3DQEBAQUABIIBgAU3Rm2JBvQ9Lvpy+5Phj97+BP3VSP383ZNAAL9JVjQhwwKU
# neWpWxM7iA2BXmQmw8kCXWa/d8rX/DeUblX4ajlnjRWveIbzHyVitjno6o+UJrTp
# 5zPQPxc/1WvWGj4IhTLGeAeJONCQHoges3lFxSgUbY/ZwwX4fJssRw8hXyWYp5T4
# CWASGZWk12wr/qfUw+tLIrUj9ToWYQO8xcwZeCaUHm+P8iOZXnIzfcOQnJMju2pN
# lDhx5PQcJB0aLXk7DsN8xUah4L5UvPWTxxiZwq9ybIqZgSyds1VfqgzFKAOPMmAU
# DkGhAfmr3aVZ/uTAdiczf2sxz+58uEdLYfyi635YzO7j8nom2riK29Hvf9zrZLc9
# zbGE/PQAtTyEtw5gRESF/ZNgamf+PK1gba9QWTitMQ7EqeWDE90r0mGE2yzwOJbO
# 6dEkM7b5U36WxyQFHo8he+kR6KVZzOUE9bXydcaqp+XLUS7MUxiWJ855k+oel5Ja
# Bac5+jW9L6a2sBDeiKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDQyMjM4
# MzBaMC8GCSqGSIb3DQEJBDEiBCAezMAeu2MLUAgF2W8bKzVGMyjmPe7/3LeX4nUs
# lgXHCTANBgkqhkiG9w0BAQEFAASCAgAAcx6CX5BgStFlvq1WByXf5+V3o5Lz9vYR
# LwwQQOD/YXlJgvHn45F5Ef0gnZzYKAsz2pW9rAKZIwsbX9ziFGek39whnFBPwdGO
# X3p5DwFwf/qPT1hfHKnOiW3etbEalETX5OxQT67X0UeGPnCMMN9ik07n2ng9tipQ
# ilcDkg1xhusOQPac3s6OUH3znJ6+1EdQJ0vrcEcqGVBIh7up3bAbe3ddKnTnIFJY
# GcQmUSUACAnqKR3D65tBeT1xi3u7zlRILdP26fj36ZcCoJU/YpWYyoJNnEBX6fXB
# X69ODR/RzireqxLIWovFpbpml1osUI1RRkMpDhZ7dyu0PtVEw7GFDCnxccNxIaH8
# AKcWObAtFzDwvQ0RYNovh12EqmD6URwKgpJh5Np8KgE2rhl/kJ3gicqUxuJtMT5f
# PiFWkfm5yooUyogq9iDp0pb/4FBxSqiOsVYQ9EUNnKUHsy6XKukSiNaW7eIRBoJa
# hKE37cp47wn1ueMZpjvzOatS34uYIGNbg5raCVpzfmFXcva+u3HRTgkjQrHLnbTF
# 86/UV5JFQyv1YoCwyIYFrhY4iMhtuxfR4HbXSwTQtHNRFicYubrG5jdM7QwZ4KKO
# iwVd2PEwy2JCGybDXbrtxaHSbgA5gYFO6jfqVvSC65lhhYkViu6fsECZnRY4xJnU
# CIQimFrPug==
# SIG # End signature block
