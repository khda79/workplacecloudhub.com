[CmdletBinding()]
param()
# Offline synthetic regressions. No collectors, tenant APIs or publication.
$ErrorActionPreference='Stop'
Import-Module (Join-Path $PSScriptRoot 'WindowsMigrationEvidence.psm1') -Force
$selectors=Get-Content -LiteralPath (Join-Path (Split-Path $PSScriptRoot -Parent) 'config/windows-migration-policy-selectors.json.txt') -Raw | ConvertFrom-Json
function Inventory([string]$Id,[string]$Name,[string]$OS='10.0.19045.1') { [pscustomobject]@{TenantKey='test';'Device ID'=$Id;'Device name'=$Name;'OS version'=$OS} }
function Prepared($R,[string]$Eligibility='Capable') {
    [pscustomobject]@{'Device Source ID'=$R.'Device ID';'Device Name'=$R.'Device name';'Operating System Version'=$R.'OS version';'Windows Generation'=$(if($R.'OS version' -like '10.0.226*'){'Windows 11'}else{'Windows 10'});'Windows 11 Upgrade Eligibility'=$Eligibility;Country='FR';'Free Disk Space (GB)'='19.5';'Windows Update State'='Error'}
}
function AD([string]$Guid,[string]$Name,[string]$IntuneId='',[string]$Enabled='True') { [pscustomobject]@{TenantKey='test';ObjectGUID=$Guid;Name=$Name;IntuneDeviceId=$IntuneId;Enabled=$Enabled;OperatingSystemShortName='Windows 10';DomainName='be.example.test';operatingSystemVersion='10.0 (19045)'} }
function Run($I,$D,$A=@(),$R=@(),$P=@()) { New-WindowsMigrationProjection -Inventory $I -Prepared $D -AD $A -Readiness $R -Policy $P -TenantKey 'test' -IntunePublicationUtc '2026-10-08T00:00:00.000Z' -ADPublicationUtc '2026-10-09T00:00:00.000Z' -PolicySelectors $selectors }
$script:checks=0
function Check([string]$Name,[bool]$Condition) { if(-not $Condition){throw "FAILED: $Name"};$script:checks++ }
function Reject([string]$Name,[scriptblock]$Action) { $failed=$false;try{& $Action | Out-Null}catch{$failed=$true};Check $Name $failed }
$i=@((Inventory 'a' 'BE-PC1'),(Inventory 'b' 'UPGRADED' '10.0.22631.1'))
$d=@($i | ForEach-Object { Prepared $_ })
$a=@((AD 'ad1' 'BE-PC1' 'a'),(AD 'ad2' 'UPGRADED'),(AD 'ad3' 'ORPHAN'),(AD 'ad4' 'DISABLED' '' 'False'))
$x=Run $i $d $a
Check 'Union / matched Windows 11 / disabled excluded' ($x.Audit.Total -eq 2 -and $x.Audit.ADOnly -eq 1 -and $x.Audit.MatchedADIntuneWindows11 -eq 1)
Check 'Country independent of domain / AD Unknown' ($x.Windows[0].Country -eq 'FR' -and $x.Windows[0].'AD Domain' -eq 'be.example.test' -and $x.Windows[1].Country -eq 'Unknown')
Check 'AD no invented disk/policy/eligibility' ($null -eq $x.Windows[1].'Free Disk GB' -and $x.Windows[1].'Autopatch Policy State' -eq 'Not observed' -and $x.Windows[1].'Upgrade Eligibility' -eq 'Not assessed')
Check 'Disk / error state retained' ($x.Windows[0].'Disk Review' -eq 'Below 20 GB planning guardrail' -and $x.Windows[0].'Update State' -eq 'Error')
Reject 'Duplicate Intune ID' { Run @($i[0],$i[0]) @($d[0],$d[0]) }
$foreign=Inventory 'c' 'FOREIGN';$foreign.TenantKey='other'
Reject 'Mixed tenant' { Run @($i[0],$foreign) @($d[0],(Prepared $foreign)) }
$foreignAD=AD 'foreign' 'FOREIGN';$foreignAD.TenantKey='other'
Reject 'Foreign AD tenant' { Run $i $d @($foreignAD) }
$stale=Prepared $i[0];$stale.'Operating System Version'='10.0.19044.1'
Reject 'Stale prepared OS' { Run $i @($stale,$d[1]) }
$x=Run $i $d @((AD 'dupe' 'ORPHAN1'),(AD 'dupe' 'ORPHAN2'))
Check 'Duplicate AD GUID excluded, not arbitrary winner' ($x.Audit.ADOnly -eq 0 -and $x.Audit.ExcludedAD['Missing or duplicate AD GUID'] -eq 2)
$x=Run $i $d @((AD 'n1' 'AD-SAME'),(AD 'n2' 'AD-SAME'))
Check 'Duplicate AD name excluded' ($x.Audit.ADOnly -eq 0 -and $x.Audit.ExcludedAD['Ambiguous name reconciliation'] -eq 2)
$j=@((Inventory 'old' 'SAME'),(Inventory 'new' 'SAME' '10.0.22631.1'))
$jd=@((Prepared $j[0] 'Already upgraded'),(Prepared $j[1]))
$ready=[pscustomobject]@{TenantKey='test';GraphId='unlinked';DeviceName='SAME';OSVersion='10.0.22631.1';ExportDateTime='2026-10-09 02:00:00'}
$x=Run $j $jd @((AD 'amb' 'SAME')) @($ready)
Check 'Ambiguous name / conflict attribution visible' ($x.Audit.Conflicting -eq 1 -and $x.Windows[0].'Eligibility Match Status' -eq 'Readiness name only - ambiguous Intune identity' -and $x.Audit.ADOnly -eq 0 -and $x.Audit.ExcludedAD['Ambiguous name reconciliation'] -eq 1)
$ready.GraphId='old';$x=Run $j $jd @() @($ready)
Check 'Readiness ID precedes ambiguous name' ($x.Windows[0].'Eligibility Match Status' -eq 'Unique readiness ID')
$x=Run $j $jd @() @($ready,$ready)
Check 'Duplicate readiness ID has no arbitrary attribution' ($x.Windows[0].'Eligibility Match Status' -eq 'Ambiguous readiness match')
$p=[pscustomobject]@{TenantKey='test';PolicyId='anchor';PolicyName='Windows Autopatch - Feature Update Anchor Policy - TEST';DeviceId='a';AggregateState='Success';BlockingReason='Completed'}
$x=Run $i $d $a @() @($p)
Check 'Anchor is context, does not migrate OS' ($x.Audit.AnchorObserved -eq 1 -and $x.Audit.IntuneWindows10 -eq 1 -and $x.Windows[0].'OS Version' -eq '10.0.19045.1')
$x=Run $i $d @() @() @($p,$p)
Check 'Duplicate policy observation unknown' ($x.Windows[0].'Autopatch Policy State' -eq 'Ambiguous policy match' -and $x.Audit.AnchorObserved -eq 0)
$p2=$p.PSObject.Copy();$p2.PolicyId='another'
Reject 'Ambiguous anchor selector' { Run $i $d @() @() @($p,$p2) }
$only11=@($i[1]);$x=Run $only11 @($d[1])
Check 'Qualified zero union keeps full schema' ($x.Windows.Count -eq 0 -and $x.Columns.Count -eq 38 -and $x.ADColumns.Count -eq 9)
$bad=Prepared $i[0];$bad.'Free Disk Space (GB)'='-1'
Reject 'Invalid negative disk' { Run @($i[0]) @($bad) }
$empty=Prepared $i[0];$empty.'Free Disk Space (GB)'='';$empty.'Windows 11 Upgrade Eligibility'='Unknown';$x=Run @($i[0]) @($empty)
Check 'Missing disk / eligibility never zero' ($null -eq $x.Windows[0].'Free Disk GB' -and $x.Windows[0].'Disk Review' -eq 'Unknown' -and $x.Audit.Unknown -eq 1)
Write-Host "PASS: $script:checks offline migration checks. No collection, publication or Power BI refresh."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCVeEjMiasfsS69
# 6mu6nIU9LBaopsT3SI9Y7mL4YV2P1aCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIIURThu/dlGZY0K+RM/rxLgA4q8fricvYiiUjtQ3f1uqMA0GCSqG
# SIb3DQEBAQUABIIBgHzmCuMqyo96rTlB4O9HsoQk0WkZpX6fbw2yih2y+Bl7qiRR
# g/8uezHVyfi81vqQW+nPoNT1rW88DOMGs1wyESRfYNoL8w6QR4yq7ZTGOtX0wxIa
# 6aky730cWSH+dsN3d0i3I+M+3PYzAUu4ZphsyUqj2ousmCNCLcLbUI/7UCU0OmEV
# x9DJYv3TL+HuLQMx1/eam8sbEvDql4peT7yXm2/mJtJpxExFBl8b9lKF+MErCrzQ
# VL2+S7ojniH6FncHzsOzEyqNN8/oFfvanm+qHR6EZMg+IgQTqc7uG4YM5KytzehV
# /q14yO4yhC8RBRANj9DyDuZub2x8acUSisHTvwKgEK4J7ZGzLz/kLxHcVMw0ScSu
# lhP11cckW1T3cXsRQdmE4fIS408nFrC4Wwm/2G4oKKRMmPjP35w3+tu6HENMOvPr
# 2EIWj9m7dehqBxcEmJZ9bpcVGm0EgZWRQS94hS3gRwg5WRsHzCu9H0WUw50pAvl5
# K4anqysuQeNsKeaMg6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDkxNDA4
# MzZaMC8GCSqGSIb3DQEJBDEiBCCanjXQQkQktrpQCTw82KsMmCEY0rmjbZdfFcUc
# c5JenzANBgkqhkiG9w0BAQEFAASCAgCZLtfx1LQVKw6i8tfB+hOorRsGw/Q02mGH
# jN4BDXVW68H33YF4MwgMija7KnH2wuRusSyH1KT9+hnIpM7S9QjRbdOOoNlRyhfe
# zWXGYHMYMfjrk1U+VQxP4znArilPoCBtK3GdbehqVkTdmAQFD09dB1L/sLTljJ6J
# RwEVUJW+OpE6e/EAIN9umSOYYJb2tUMe1a6tVdj47rda4c6voFBsv/XR/Z6tWnnf
# CUcCvc1nMlZVrYq0mHF7BJpSXmJgs3UBtnznu1/4ffN1IHnvJQFW4Rpx6gKIJmAe
# EbiuwmKYob8HY8bUD6UQhA7VU+KwG9Zz0skF8cMzXrvt5tWGcQobhMaGnDphgZlf
# j8mWCgi0qx4bCxMQM+h3trmIFOX7U8Bf7NyP8bRwxCntekmpxblDDefGfpeS9+2f
# T0J3doU58XO4I7nCs/ZSqF5mfH3xMYEJ3qb1guukV3poW3na5k3sOvjlhxH/4pxD
# bSPL/JXySPK4J9ueK4QQxNYpwVOIoqTJkl+yUbPMKbVB0UTEIuWSsQgof6m8z6sZ
# dA2++7dGLnGWahMiHuTTyH4PkZ9Am1UtVqI3R+zWzIYUKMlcqPMdw61x3fN6KV2D
# 6Ni12VH0xyfAZ0mEorhWSCCTfqnE1rsLqL5UWYPLkWpiSk1b2xC+XUkqXyLo4JnW
# njgc8b51rQ==
# SIG # End signature block
