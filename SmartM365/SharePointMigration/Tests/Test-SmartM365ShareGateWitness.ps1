<#
.SYNOPSIS
    Offline five-item ShareGate witness safety and CSV reader contract.
.VERSION
    1.0.0
#>
#Requires -Version 5.1
[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$testsRoot = (Resolve-Path -LiteralPath $PSScriptRoot).ProviderPath
$scratch = Join-Path $testsRoot ('.witness-test-' + [guid]::NewGuid().ToString('N'))
if (-not $scratch.StartsWith($testsRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) { throw 'Unsafe test path.' }
$oldModulePath = $env:PSModulePath

try {
    $moduleDir = Join-Path $scratch 'Modules\ShareGate'
    $analysis = Join-Path $scratch 'Project\ShareGate\Diagnostics\Analysis'
    New-Item -ItemType Directory -Path $moduleDir,$analysis -Force | Out-Null
    @'
@{ RootModule='ShareGate.psm1'; ModuleVersion='99.0.0'; GUID='5d5f7e6e-9039-4c51-ad50-9e28d3df67c3'; FunctionsToExport=@('Connect-Site','Get-List','Copy-Content','Export-Report') }
'@ | Set-Content -LiteralPath (Join-Path $moduleDir 'ShareGate.psd1') -Encoding UTF8
    @'
function Connect-Site { [CmdletBinding()] param([string]$Url,[switch]$Browser) if ($Url -match 'target' -and -not $Browser) { throw 'Destination Browser missing' }; if ($Url -match 'source' -and $Browser) { throw 'Source must use current Windows user' }; [pscustomobject]@{ Url=$Url } }
function Get-List { [CmdletBinding()] param($Site,[string[]]$Name) [pscustomobject]@{ Title=$Name[0]; Site=$Site } }
function Copy-Content { [CmdletBinding()] param($SourceList,$DestinationList,[int[]]$SourceItemId,[switch]$WhatIf) if (-not $WhatIf) { throw 'Copy-Content without WhatIf' }; [pscustomobject]@{ Id=('261003-'+$SourceItemId[0]); SourceId=$SourceItemId[0]; Marker='synthetic' } }
function Export-Report { [CmdletBinding()] param($CopyResult,[string]$Path) if ($CopyResult.SourceId -in @(2,25)) { [pscustomobject]@{ Result='Warning'; Type='File'; Title='sample'; Path='/sample'; Version='1.0'; 'Site name'='sample'; 'Site address'='https://source.example'; 'List title'='sample'; Error=''; 'Help links'=''; Details='Known permanent warning' } | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8 } else { Set-Content -LiteralPath $Path -Value 'Result,Type,Title,Path,Version,Site name,Site address,List title,Error,Help links,Details' -Encoding UTF8 } }
Export-ModuleMember -Function *
'@ | Set-Content -LiteralPath (Join-Path $moduleDir 'ShareGate.psm1') -Encoding UTF8
    @'
SessionId,RowId,RuleId,ItemKey,SourceUrl,SourceList,SourceItemId,DestinationUrl,DestinationList,ObjectType,Status
260930-6,1,SG-SHORTCUT,shortcut-2,https://source.example/sites/a,Media,2,https://target.example/sites/a,Media,File,Warning
260930-6,2,SG-MODERN-LINK,modern-25,https://source.example/sites/a,Pages,25,https://target.example/sites/a,Pages,File,Warning
260930-6,3,SG-ACCESS-SOURCE,access-88,https://source.example/sites/a,Documents,88,https://target.example/sites/a,Documents,File,Error
260930-6,4,SG-ACCESS-SOURCE,access-868,https://source.example/sites/a,Documents,868,https://target.example/sites/a,Documents,File,Error
260930-6,5,SG-ACCESS-SOURCE,access-1,https://source.example/sites/a,Pages,1,https://target.example/sites/a,Pages,File,Error
'@ | Set-Content -LiteralPath (Join-Path $analysis 'ClassifiedRows.csv') -Encoding UTF8
    $env:PSModulePath = (Join-Path $scratch 'Modules') + [IO.Path]::PathSeparator + $oldModulePath
    $scriptPath = Join-Path $testsRoot '..\Scripts\Diagnostics\SmartM365-SharePointMigration-ShareGateWitness.ps1'
    $project = Join-Path $scratch 'Project'
    try {
        & $scriptPath -ProjectRoot $project -AnalysisDirectory $analysis -SessionId '260930-6' -Run | Out-Null
        throw 'Expected mandatory WhatIf guard.'
    }
    catch { if ($_.Exception.Message -notmatch 'WhatIf is mandatory') { throw } }
    try {
        & $scriptPath -ProjectRoot $project -AnalysisDirectory $analysis -SessionId '260930-6' -ExpectedAnalysisHash ('0' * 64) -WhatIf -DryRun | Out-Null
        throw 'Expected analysis hash guard.'
    }
    catch { if ($_.Exception.Message -notmatch 'ClassifiedRows.csv changed') { throw } }
    $dry = @(& $scriptPath -ProjectRoot $project -AnalysisDirectory $analysis -SessionId '260930-6' -WhatIf -DryRun)
    if (@($dry | Where-Object { $_ -match 'WitnessShortcut: ID=2|WitnessModernLink: ID=25|AccessDominantFirst: ID=88|AccessDominantLast: ID=868|AccessOtherList: ID=1' }).Count -ne 5) { throw 'DryRun did not select the intended five items.' }
    if (@(Get-ChildItem -LiteralPath (Join-Path $project 'ShareGate\Diagnostics') -Directory -Filter 'Witness-*').Count) { throw 'DryRun created witness output.' }
    & $scriptPath -ProjectRoot $project -AnalysisDirectory $analysis -SessionId '260930-6' -WhatIf -Run | Out-Null
    $output = Get-ChildItem -LiteralPath (Join-Path $project 'ShareGate\Diagnostics') -Directory -Filter 'Witness-*' | Select-Object -First 1
    if (-not $output) { throw 'Witness output was not created.' }
    $results = @(Import-Csv -LiteralPath (Join-Path $output.FullName 'Witness-Results.csv'))
    if ($results.Count -ne 5) { throw "Expected 5 results, got $($results.Count)." }
    if (@($results | Where-Object { $_.Role -like 'Witness*' -and $_.Status -eq 'Pre-check warning' -and $_.ReportRows -eq '1' }).Count -ne 2) { throw 'Positive witnesses were not read through Result alias.' }
    if (@($results | Where-Object { $_.Role -like 'Access*' -and $_.Status -eq 'Undetermined - empty report' -and $_.ReportRows -eq '0' }).Count -ne 3) { throw 'Header-only access reports were misclassified.' }
    if (@($results | Where-Object { $_.PrecheckSessionId -match '^261003-' }).Count -ne 5) { throw 'CopyResult session IDs were not captured.' }
    if (@(Get-ChildItem -LiteralPath (Join-Path $output.FullName 'CopyResults') -File -Filter '*.txt').Count -ne 5) { throw 'CopyResult properties were not saved for all calls.' }
    if (@(Get-ChildItem -LiteralPath (Join-Path $output.FullName 'Reports') -File -Filter '*.csv').Count -ne 5) { throw 'Export-Report was not called for all results.' }
    $log = Get-Content -LiteralPath (Join-Path $output.FullName 'Witness.log') -Raw
    if ($log -notmatch 'Both positive witnesses produced rows while all three prior 401 reports were empty') { throw 'The witness interpretation is missing.' }
    . (Join-Path $testsRoot '..\Scripts\Diagnostics\SmartM365-SharePointMigration-ShareGateReportReader.ps1')
    $config = Join-Path $scratch 'Config'
    New-Item -ItemType Directory -Path $config -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $testsRoot '..\Config\sharegate-diagnostics.columns.json.template') -Destination $config
    '{"Status":["Outcome"],"Errors":["Failure"]}' | Set-Content -LiteralPath (Join-Path $config 'sharegate-diagnostics.columns.json.txt') -Encoding UTF8
    $customAliases = Get-SmartM365ShareGateReportAliases -ConfigRoot $config
    $customRows = @([pscustomobject]@{ Outcome='Error'; Failure="401 Unauthorized. URL 'https://source.example/sites/a/_vti_bin/client.svc/ProcessQuery' was not authorized" })
    $customAssessment = Get-SmartM365ShareGatePrecheckAssessment -Rows $customRows -SourceItemId 88 -SourceUrl 'https://source.example/sites/a' -DestinationUrl 'https://target.example/sites/a' -Aliases $customAliases
    if ($customAssessment.Status -ne 'Source 401 observed') { throw 'Custom configured column aliases were not honored.' }
    Write-Output 'ShareGate witness offline contract test passed.'
}
finally {
    $env:PSModulePath = $oldModulePath
    if ((Test-Path -LiteralPath $scratch) -and $scratch.StartsWith($testsRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $scratch -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD/795awzUsf2HD
# H17YURu/5xocDk/nda/ZPcge6iXLH6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDjzRXLxdNxtFR6EjrRBs4P3bIGBvOJzSEnJjGr21vsoMA0GCSqG
# SIb3DQEBAQUABIIBgKbHMYkSZvzpACAXNL7UDRzH52ezxFU+ko/QsGj9WDgveZAJ
# okcVpaCLrDu1psAeYUM4Nf4SVkStUKphcWyik5yz45/PoikiUY2gwCHLlzEbkC5T
# FpzYmkNZ5gK0hPjprChvHaFBpjT0Nx/sMOXZaNavRqA8N6FPUZmDe+K7MsRUcYp9
# B8Fxu26BlRyM3H4nO3jp1tQFh3hu/OLED60csDJlVrwoobvVIehlAG+vE9oneRUy
# jlfYaa+8vCDxDrq83+jKI3yPERtg5B9ShSVpl7xqXnv3nyMZSdAf62h6aUcNU+1T
# rubhzhb5yvHnol2jKRiiLzZCrLbstO6KtYC0/N0V57ZlGoe90FF/UL42Lu0zR2Q6
# kZPVTvMnxZzY/ygxFUtwNLwPlJv2Y9mSxVk3n/ccKG1AlZRoAPvcoNg+ZRMvEH7F
# pNaB6NPRRK5e54UHXwvo1LkiNJMlIp2Oi7N4q4LZKzXjRZwHXcXvLdDrKNber/lE
# xBJMEkBMiiTXFKiW46GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxMDQy
# NDVaMC8GCSqGSIb3DQEJBDEiBCDyvUgyHls+whTLTAkTO+YZ3FYu+NU5okE/TuZ0
# m1S2BzANBgkqhkiG9w0BAQEFAASCAgCvFIN6jPuvW+3NYaYMe3wbsVFai6PlW4Mz
# tIkNStryKCz8Def7PxLfO+rOFH5GQAvOZ50zyE6ATSxf7fVBZIFPDEPSxUvVGBik
# iZtpPoBHOdjXq9x1ExT14z2YcQKMkpFrmcxN0dxnZYaSbp7/korUE48CTGIZe7zH
# NaAMLkkrkJjDR5qkozuTO0M+1aF5jMdQwaeT1wbfUT0hunlOSiytSRrMEbKqxTae
# zejzOLRaPOe0XRvbD0N02+iwNeYk/qKRHTUhF3LUmZGmQPGUmH6hYSKbZzmaqU/i
# MGVzu5xzpj1udn6GRgvHXKgoPN64YGbzNVtUCG2C1xTtFZE+GSNGYKi1ZsQe/JkZ
# K1GnBHy2AsjWvY5L+eFQCgnnrNsb0ZZsQ/xN3xUKxNbB8EGFwWMy4C+R1xBldlGh
# uG21hxek2JOjr6gOAXXZfkV94WqQYNCHcQrm4mrSPvOWcsgWnchZN23uXdrlRBrd
# iFm3LvTIjuvnczVQ/AMao/btVgBJCWJHjqVAzboaarPhBeowhf6M/xunJ9UkXUhH
# bGobOhgUZ8sN+PlgJRsjJL5YhStD2cXHC28P38hvAoy0VOCqc60Zwimi70xasHUP
# vl0xpvZlv0g8cVqJL+yHspCc//KPKVgKW9Uaasc0OGv+CnGuZ4jmEs5kFDt1Qsl3
# DFDYQlBw6Q==
# SIG # End signature block
