<#
.SYNOPSIS
Offline checks of Publish-SmartM365-DataAllBackfill.ps1 (file selection, preview default, guards).

.VERSION
1.1.0
#>
#requires -Version 7.2
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path $PSScriptRoot '..\SmartInventory\Maintenance\Publish-SmartM365-DataAllBackfill.ps1'
$checks = 0
function Check([bool]$Condition, [string]$Message) { if (-not $Condition) { throw "FAIL: $Message" }; $script:checks++ }

$tokens = $null; $errors = $null
$ast = [Management.Automation.Language.Parser]::ParseFile((Resolve-Path $scriptPath), [ref]$tokens, [ref]$errors)
Check ($errors.Count -eq 0) 'Script does not parse.'
$planFunction = $ast.Find({ param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Get-DataAllBackfillPlan' }, $true)
Check ($null -ne $planFunction) 'Get-DataAllBackfillPlan not found.'
. ([scriptblock]::Create($planFunction.Extent.Text))

$text = Get-Content -LiteralPath $scriptPath -Raw
Check ($ast.ParamBlock.Parameters.Name.VariablePath.UserPath -contains 'Execute') 'No -Execute switch.'
Check ($text -match '\.VERSION\s+1\.1 ') 'Help .VERSION missing.'
Check ($text -match "if \(\`$Execute -and -not \`$policy\.Enabled\) \{ throw") 'Execute is not refused when DATA-ALL uploads are disabled.'
Check ($text -match "Invoke-SmartM365SharePointCsvUpload -LocalFilePath \`$entry\.FullName -Enabled \`$true -EnsureParentFolders") 'Upload does not reuse the Core upload.'
Check ($text -notmatch 'Remove-SmartM365SharePoint') 'The backfill must never delete on SharePoint.'
Check (([regex]::Matches($text, 'Invoke-SmartM365SharePointCsvUpload')).Count -eq 1) 'Upload called outside the -Execute branch.'

$root = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-BackfillTest-' + [guid]::NewGuid().ToString('N'))
$dataAll = Join-Path $root 'DATA-ALL'
try {
    $now = [datetime]'2026-11-17T10:00:00'
    function New-TestFile([string]$Relative, [datetime]$Written, [int]$Size = 10) {
        $path = Join-Path $dataAll $Relative
        [IO.Directory]::CreateDirectory((Split-Path $path -Parent)) | Out-Null
        [IO.File]::WriteAllBytes($path, [byte[]]::new($Size))
        (Get-Item -LiteralPath $path).LastWriteTime = $Written
    }
    New-TestFile 'Intune\Devices\Devices_20261116_010000.csv' $now.AddHours(-33)
    New-TestFile 'Intune\Devices\Devices_20261105_010000.csv' $now.AddDays(-12)
    New-TestFile 'Intune\Devices\Devices_20261020_010000.csv' $now.AddDays(-28)
    New-TestFile 'Intune\Devices\Devices_20261116-010000.xlsx' $now.AddHours(-33)
    New-TestFile 'Intune\Devices\WeeklyHistory\2026-W45\Devices.csv' $now.AddDays(-9)
    New-TestFile 'Intune\Devices\WeeklyHistory\2026-W45\manifest.json.txt' $now.AddDays(-9)
    New-TestFile 'Usage\20261110_044500\Usage.csv' $now.AddDays(-7)
    New-TestFile 'Orchestrator\Jobs.json.txt' $now.AddDays(-1)
    New-TestFile 'Intune\.staging\Devices.csv' $now.AddDays(-1)
    New-TestFile 'Intune\Devices\Devices.csv.lock' $now.AddDays(-1)
    New-TestFile 'Intune\Devices\Devices_20261117_095000.csv' $now.AddMinutes(-10)
    New-TestFile 'Intune\Remediations\Exports\20261115-043032\Scripts\Fix_0001\detect.ps1' $now.AddDays(-2)
    New-TestFile 'Intune\Remediations\Exports\WeeklyHistory\2026-W46\Remediations.csv' $now.AddDays(-2)
    New-TestFile 'Intune\Remediations\Remediations.csv' $now.AddDays(-2)
    New-TestFile 'Intune\Remediations\ExportsArchive\Remediations.csv' $now.AddDays(-2)
    New-TestFile 'Intune\WindowsUpdate\Status\Work\Extract_20261116_020016_abc\part.csv' $now.AddDays(-1)
    New-TestFile 'Intune\WindowsUpdate\Status\Status.csv' $now.AddDays(-1)
    New-TestFile 'Exchange\Work\Mailboxes.csv' $now.AddDays(-1)
    New-TestFile 'M365\Licensing\Licenses\DailyCaptureRecords\SmartM365-Licences-Inventory\record.json.txt' $now.AddDays(-1)
    New-TestFile 'M365\Licensing\Licenses\DailyCaptures\0da4e1ada690\Licenses.csv' $now.AddDays(-1)

    $plan = @(Get-DataAllBackfillPlan -DataAllRoot $dataAll -Since ([datetime]'2026-10-24') -TimestampedCutoff $now.AddDays(-7) -NewestAllowed $now.AddMinutes(-30))
    $byName = @{}; foreach ($entry in $plan) { $byName[$entry.Folder + '|' + $entry.Name] = $entry }
    function Expect([string]$Key, [string]$Action, [string]$Reason) {
        Check ($byName.ContainsKey($Key)) "Missing plan entry $Key."
        Check ($byName[$Key].Action -eq $Action -and $byName[$Key].Reason -eq $Reason) ("{0}: expected {1}/{2}, got {3}/{4}." -f $Key, $Action, $Reason, $byName[$Key].Action, $byName[$Key].Reason)
    }
    Check (-not $byName.ContainsKey('Intune\Devices|Devices_20261020_010000.csv')) 'A file written before -Since was planned.'
    Expect 'Intune\Devices|Devices_20261116_010000.csv' 'Upload' ''
    Expect 'Intune\Devices|Devices_20261116-010000.xlsx' 'Upload' ''
    Expect 'Intune\Devices|Devices_20261105_010000.csv' 'Skip' 'BeyondSharePointRetention'
    Expect 'Intune\Devices\WeeklyHistory\2026-W45|Devices.csv' 'Upload' ''
    Expect 'Intune\Devices\WeeklyHistory\2026-W45|manifest.json.txt' 'Upload' ''
    Expect 'Usage\20261110_044500|Usage.csv' 'Upload' ''
    Expect 'Orchestrator|Jobs.json.txt' 'Skip' 'OrchestratorFolder'
    Expect 'Intune\.staging|Devices.csv' 'Skip' 'HiddenFolder'
    Expect 'Intune\Devices|Devices.csv.lock' 'Skip' 'TransientFile'
    Expect 'Intune\Devices|Devices_20261117_095000.csv' 'Skip' 'TooRecent'
    Expect 'Intune\Remediations\Exports\20261115-043032\Scripts\Fix_0001|detect.ps1' 'Skip' 'NotPublishedByCollector'
    Expect 'Intune\Remediations\Exports\WeeklyHistory\2026-W46|Remediations.csv' 'Skip' 'NotPublishedByCollector'
    Expect 'Intune\Remediations|Remediations.csv' 'Upload' ''
    Expect 'Intune\Remediations\ExportsArchive|Remediations.csv' 'Upload' ''
    Expect 'Intune\WindowsUpdate\Status\Work\Extract_20261116_020016_abc|part.csv' 'Skip' 'NotPublishedByCollector'
    Expect 'Intune\WindowsUpdate\Status|Status.csv' 'Upload' ''
    Expect 'Exchange\Work|Mailboxes.csv' 'Upload' ''
    Expect 'M365\Licensing\Licenses\DailyCaptureRecords\SmartM365-Licences-Inventory|record.json.txt' 'Skip' 'NotPublishedByCollector'
    Expect 'M365\Licensing\Licenses\DailyCaptures\0da4e1ada690|Licenses.csv' 'Upload' ''
    $ordered = @($plan | ForEach-Object LastWriteTime)
    Check ((@($ordered | Sort-Object) -join '|') -eq ($ordered -join '|')) 'Plan is not ordered oldest first.'
}
finally { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }

Write-Host "DATA-ALL backfill offline: $checks checks passed."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCA0Ti/QdEUCNNR7
# h2F5ky96wjgzkHwwou8iURSt0wdaM6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEILU7JsR+eXrErajDPWATG2JAf6ubzJ/2ZD60M9jPHOJuMA0GCSqG
# SIb3DQEBAQUABIIBgCKGTbvOPD54BYlNi5JE4XWoUAwqidmJOJsjHeJynrQFSkB5
# JukyskA/0L3VMREvUkc6AwkgcYMYa5bbs8ObQ6YIlQg28W4++HjM0Fyao/vrNKZ5
# duu/MbcLykCXJqfU2VgsyhU20C2s6Cf2qgEAIIjjQmT4BuloxsWxax9KkST35KhP
# NMMs4ydUCXF5wIWVsi8rVFocpTw6tKFScna63d1Q8W5ERAwPjXOG4yUIDKgQPAbf
# dNmtPEr6cYT8kBdrva5faOsyXjp/roTYXv71wjs+BAWGrX1237LwXY+ZCCQee+8S
# kQJwAMdCZ4aSPNGWh5I3kTeKgD7NlPqK7ctWESaJ5igRLeODzK2pihkMXnmh2hcS
# NivmmaggwraafXLMjG+ZcfCvPthG5C3OFAWFs5gwCqcZrk+ArNIHiAGPeYUmIOqo
# 0dpzHqNMExvBPVPVqLxl2AU0KXlRjlEjrMWuegx8mdoNTs2lY/KYCwE8yVxaAeHf
# UNzww5KSBVKSbFWVRaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMTcxMTU0
# NDJaMC8GCSqGSIb3DQEJBDEiBCBsskEi1eSNtgCeU41Ucg2Pm1zsuHXDc7M30+m+
# rlepMDANBgkqhkiG9w0BAQEFAASCAgBPrsIj3Y7jTwE4cIJ+P8LlPKM1uYWn97Tb
# 9xlkPrhqe68hzzWdw1vqe35wN+cToW0iY+IPAAMltN1rtny8ITYSE5cKt0tc+OTw
# Ea7Bcrx+rChIOnniUlc6M4j//S5Id2TuFD7pSAnN1hzuFC9pVpyxyw0oq6LCvjtt
# iuOiUQEK/uLvGVzmfZVi5V4fbEsd8Z9nku6VZHSI2/tfkrHJEl1O2RP0Ewalseew
# ujhHx5C8ad8AE4LGiBCFjQqGRLsfiiv0WopyNLCwtQbr2LqgdJMnjmMVhn7SGpNF
# N3yvoCFAi131NYuUgw6ebiNMQlWpyOYdJwKGvBYaJHcy2aAYoH5mtxBQ8bEWa35g
# 2wMOO5mGzBow7WIzPlNv+a7PkBiiMOl0s+uI8iojmIKSLmPT57TDiMlc4e+E2jJv
# fhql7mGjvwoGuvdvXaOiC+RyL/FSwmQLYwKKNnpfrYxwQ14KJZYN6G4N8nvoghwy
# Agh88C5G5vLDKu0/6wys6tQcOZuQLY4XKIx6ar2+E78zhQKa1CLI5IV8AF5FvGyR
# srK6gVB4ZePAPgu/Ia+etZrfhJzJUmH0N02ExZwuIy24AMsu73vdSzgT/KpzxCT8
# quJyzciyTWCuNtLjvM7hdWLNHHtDlyNWKY+uum60eL4wmJlvMCNXxJBF3wcWsy/J
# kLKZ+gYDJQ==
# SIG # End signature block
