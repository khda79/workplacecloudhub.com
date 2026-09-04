#Requires -Version 5.1
# Synthetic collection and finalization contracts for 0.3.0.
[CmdletBinding()]param()
$ErrorActionPreference='Stop'
$app=Split-Path -Parent $PSScriptRoot
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $app 'SmartM365-EndpointDiagnosticsAnalyzer-GUI.ps1'),[ref]$tokens,[ref]$errors)
if($errors.Count){throw ($errors.Message -join '; ')}
foreach($name in @('New-SedaObject','Get-SedaTextContent','Get-SedaResultsXml','Write-SedaCollectionResults','copy-collect-files','Set-SedaAnalysisProgress')) {
    $node=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $name},$true)
    if(-not$node){throw "Missing helper $name"};Invoke-Expression $node.Extent.Text
}
$root=Join-Path ([IO.Path]::GetTempPath()) ('seda_collection_'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$checks=New-Object 'System.Collections.Generic.List[object]'
$script:messages=New-Object 'System.Collections.Generic.List[string]'
function Write-SedaLog {param($Level,$Message,$Exception)$script:messages.Add("$Level $Message")}
function check([string]$Name,[bool]$Pass){$checks.Add([pscustomobject]@{Name=$Name;Passed=$Pass});if(-not$Pass){throw "FAIL: $Name"}}
function fixture([string]$Relative,[string]$Text='fixture') {
    $path=Join-Path $root $Relative;New-Item -ItemType Directory -Path (Split-Path -Parent $path) -Force|Out-Null
    [IO.File]::WriteAllText($path,$Text,[Text.UTF8Encoding]::new($false));return $path
}
for($i=1;$i -le 201;$i++){ $file=fixture "source\$i.log"; (Get-Item $file).LastWriteTime=(Get-Date).AddDays(-180) }
[void](fixture 'source\excluded.bin')
$count=copy-collect-files -Source (Join-Path $root 'source') -Destination (Join-Path $root 'all') -Extensions @('.log')
$metadata=Get-Content (Get-ChildItem (Join-Path $root 'all') -Filter '*.collection.json').FullName -Raw|ConvertFrom-Json
check 'All 201 old eligible logs retained without count/age cap' ($count -eq 201 -and $metadata.Copied -eq 201 -and $metadata.ExcludedByAge -eq 0 -and $metadata.ExcludedByCount -eq 0)
check 'Extension exclusion is not a copy failure' ($metadata.Status -eq 'Excluded' -and $metadata.ExcludedByExtension -eq 1 -and @($metadata.Errors).Count -eq 0)
Write-SedaCollectionResults -Root (Join-Path $root 'all')
$parsed=Get-SedaResultsXml @((Join-Path $root 'all\results.xml'))
check 'Excluded source survives XML as information with reason' ($parsed.Errors.Count -eq 0 -and $parsed.Excluded.Count -eq 1 -and $parsed.Excluded[0].Detail -match 'ExcludedByExtension')
[void](copy-collect-files (Join-Path $root 'source') (Join-Path $root 'limited') -MaxFiles 80 -Extensions @('.log'))
Write-SedaCollectionResults -Root (Join-Path $root 'limited')
$parsed=Get-SedaResultsXml @((Join-Path $root 'limited\results.xml'))
check 'Explicit cap is partial not success or generic failure' ($parsed.Partial.Count -eq 1 -and $parsed.Errors.Count -eq 0 -and $parsed.Partial[0].Detail -match '"ExcludedByCount":121')
[void](copy-collect-files (Join-Path $root 'source') (Join-Path $root 'aged') -MaxAgeDays 90 -Extensions @('.log'))
Write-SedaCollectionResults -Root (Join-Path $root 'aged')
check 'Explicit age exclusion is partial even when zero copied' ((Get-SedaResultsXml @((Join-Path $root 'aged\results.xml'))).Partial.Count -eq 1)
[void](copy-collect-files (Join-Path $root 'missing') (Join-Path $root 'absent'))
Write-SedaCollectionResults -Root (Join-Path $root 'absent')
check 'Absent directory is reported not silently ignored' ((Get-SedaResultsXml @((Join-Path $root 'absent\results.xml'))).Unavailable.Count -eq 1)
[void](fixture 'statuses\key.reg.unavailable.txt' 'Optional registry source absent')
[void](fixture 'statuses\denied.error.txt' 'Access denied')
[void](fixture 'statuses\command.timeout.txt' 'Timed out')
[void](fixture 'statuses\mdmdiag_error.txt' 'Native command failed')
[void](fixture 'statuses\broken.collection.json' '{')
Write-SedaCollectionResults -Root (Join-Path $root 'statuses')
$parsed=Get-SedaResultsXml @((Join-Path $root 'statuses\results.xml'))
check 'Absent registry is INFO; failures/timeouts/malformed metadata remain ERROR' ($parsed.Unavailable.Count -eq 1 -and $parsed.Errors.Count -eq 4)
# Simulate a native copy refusal, not an ACL change to the workstation.
function Copy-Item {param($LiteralPath,$Destination,[switch]$Force,$ErrorAction)throw 'Synthetic access denied'}
[void](copy-collect-files (Join-Path $root 'source') (Join-Path $root 'failed') -Extensions @('.log'))
Remove-Item Function:\Copy-Item
Write-SedaCollectionResults -Root (Join-Path $root 'failed')
$parsed=Get-SedaResultsXml @((Join-Path $root 'failed\results.xml'))
check 'Copy refusal remains an error with evidence' ($parsed.Errors.Count -eq 1 -and $parsed.Errors[0].Detail -match 'Synthetic access denied')
check 'Copy failures produce a log warning' (@($script:messages|Where-Object {$_ -match '^WARN Collection copy: status=Failed'}).Count -eq 1)
# Exercise the real registry loop against an absent fixture and simulated access denial.
$loop=$ast.Find({param($n)$n -is [Management.Automation.Language.ForEachStatementAst] -and $n.Variable.Extent.Text -eq '$rk'},$true)
function next-name {param($Kind,$Name,$Ext)Join-Path $root ('registry-'+$Name+'.'+$Ext)}
function set-collect-progress {param($Step)}
function write-collect-text {param($Path,$Text,$Header)[IO.File]::WriteAllText($Path,$Header+"`n"+$Text)}
function Invoke-SedaProcessWithTimeout {throw 'Native export must not run for absent/inaccessible registry'}
function Test-Path {param($LiteralPath,$ErrorAction)if($LiteralPath -match 'Denied'){throw [UnauthorizedAccessException]::new('Synthetic registry access denied')};return $false}
$registryKeys=@(@{Name='Absent';Path='HKLM\SOFTWARE\Absent'},@{Name='Denied';Path='HKLM\SOFTWARE\Denied'})
Invoke-Expression $loop.Extent.Text
Remove-Item Function:\Test-Path
check 'Absent registry uses unavailable marker' (Test-Path (Join-Path $root 'registry-Absent export.reg.unavailable.txt'))
check 'Denied registry is not classified absent' (Test-Path (Join-Path $root 'registry-Denied export.reg.error.txt'))
$calls=$ast.FindAll({param($n)$n -is [Management.Automation.Language.CommandAst] -and $n.GetCommandName() -eq 'copy-collect-files'},$true)
check 'Production collection has no count or age limits' (@($calls|Where-Object {$_.Extent.Text -match '-MaxFiles|-MaxAgeDays'}).Count -eq 0)
# Execute the real serialization branch on complete synthetic evidence.
$exportNode=$ast.Find({param($n)$n -is [Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text -eq '$ExportAnalysisClixmlPath'},$true)
$AnalysisProgressPath=Join-Path ([IO.Path]::GetTempPath()) ('seda_analysis_'+[guid]::NewGuid().ToString('N')+'.progress.txt')
$ExportAnalysisClixmlPath=Join-Path $root 'analysis.clixml'
$analysis=[pscustomobject]@{Events=@(1..6001);FullMessage=('Evidence '*1000)}
Invoke-Expression $exportNode.Extent.Text
$restored=Import-Clixml $ExportAnalysisClixmlPath
check 'Finalization serialization preserves all records and full messages' ($restored.Events.Count -eq 6001 -and $restored.FullMessage -ceq $analysis.FullMessage)
check 'Finalization progress identifies step 1 of 3' ((Get-Content $AnalysisProgressPath -Raw) -match '\|1\|3\|Finalization: exporting complete results')
check 'Finalization logs measured completion' (@($script:messages|Where-Object {$_ -match 'Finalization 1/3 completed:.*duration=.*bytes='}).Count -eq 1)
[pscustomobject]@{Result='PASS';Engine=[string]$PSVersionTable.PSVersion;Checks=$checks.Count;FixturePath=$root}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBNfu6HFZwgZ58u
# nsXP9OVFyIWSZjyp7/BVmiUeOvtDjKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIL/2wBJaM5BhgXFOH4kiyk6WbkoWdRMd8yqIcvYeAGH2MA0GCSqG
# SIb3DQEBAQUABIIBgDbNGZ4kBB3/vXTjpndRoNzYUOZB8FdKYLweba1QEtb211Oe
# bDW/RbC9uR9L0zn/BlSuhSaccYRMogZwFBme1KtyQo6YfuMtk9r6HVbTtTPRQJwb
# Op9moelP7pDQNO5m/H3y9uPkYAMCB75ml/CtH7BZBk38KxOCz8ocjfUV0P736dco
# JrHxFNK57xUd/H1tVZaIHDB2kG3TNa/RDFAWWsjbS3bTk4ucWbzmQO9tKftC0Ott
# qGe6CZh+RNHnBg2ZmDhldO+SS1z0JtriF0E6ud/zgINq9SLhggEntDWbsVzM/QLz
# qbORSbmyhvcJIL/EC34y+r00IS6+BHapbJ2wUKCY7KXLFhvzrhqLWvnIaebVftwj
# MJHFmlhqTMi+8tkd/nS58N2dso+MgkPhhhSmFZ6pAEpYGI6Oo9WjPCIbZrHJlQFK
# zE6cgDHn35vMd/eH1m9l6CYpC5beX0OnzZj90VUESntHdQ7DaBDIvjeAPrzp3jks
# N509VFA1CFQ/rnNl1aGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MDQxNTQw
# MzhaMC8GCSqGSIb3DQEJBDEiBCDh/eVUaTibv4Uej3yyIJWaGVbhSwh10IsnH0xx
# nlTm7DANBgkqhkiG9w0BAQEFAASCAgAGVCl6waEs2YQWkpDXL2IadpxEUBZYjn8I
# jB3HhNLJ/Plv9mDsYYpLTXpVzs2j9APnmoZkUM3c3+QXlK0FV/DoEXjGN0lSP4/Y
# RZGUdmty/WaNmhvVUwnc0ZywOJEZRJNPzXKU8/HsE/fqNhpiosh9Fs1umVDoP27g
# MS9rVY90xXcwT3kCod7eu39303U3F04RdUCIEb0dxxpZ394cEyTIAtyh63ffge2E
# 8B/zFzAJ8ZIJ8tv7USbTXBMc6W6HQOS+yFeoujANe4ia2xrkd0jO6gND2QmkoXVT
# YWdK5Pv6HiqCxb6hQMMLQfhjfgJEOS/X4eM0yuerQrPIBbkxSXvdcS75+ZyWOmUv
# H7lp2MVJDiGFv7XNswEyPm89Na6HeIKonCK110xJnxfNt6hhFOg3f1oPdZ2/D7/a
# J+3tuahzirW6nYwmwU3EDMja+fCiWQVTV6hzqFLjFNk6Kkq8x9C8dSRt0k4hi+iw
# oNk+rue4YXUD6fgDMPZ+Pp9WMbRWZXmvXA6Q33hJjI/PR4r+7iV9XlWofEP5c8lw
# tKzQkyamqAnQWZur/OvZzT1gJ1rF2uJOC+uQnxqLe9xYgwHeleHAs1qFkhEQGgaG
# EjZenQnEa6STTfq4wnbqxQaTFPgyIrhWCHo83v7L5P6yQWu19dTJFEhqB5tI5deM
# 9tLLFxO2/w==
# SIG # End signature block
