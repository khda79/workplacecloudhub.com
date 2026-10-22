<# Synthetic reader -> consolidation -> business rules; no collector or data reads. #>
[CmdletBinding()]
param()
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot '../SmartInventory/Common/SmartM365.IdentityDateEvidence.ps1')
$checks=0
function Check([bool]$Ok,[string]$Message) { if(-not $Ok){throw $Message}; $script:checks++ }
$now=[datetimeoffset]'2026-10-22T12:00:00Z'
function Evidence([string]$Iso) { Read-IdentityDate ([pscustomobject]@{Date='legacy';DateIso8601=$Iso}) Date -Now $now }
$unknown=Read-IdentityDate ([pscustomobject]@{Date='2026-10-21T00:00:00Z'}) Date -Now $now
Check ($unknown.Status -eq 'Unqualified legacy value' -and $null -eq $unknown.Utc) 'Old inferred Z was trusted'
foreach($raw in @('2026-10-21T00:00:00Z','21/10/2026 00:00:00')) {
    Check ((ConvertTo-IdentityNativeIso $raw) -eq '') 'String/deserialized value was promoted to native proof'
}
$typed=[datetimeoffset]'2026-10-21T02:00:00+02:00'
Check ((Evidence (ConvertTo-IdentityNativeIso $typed)).Utc -eq [datetimeoffset]'2026-10-21T00:00:00Z') 'Offset not preserved'
$unspecified=[datetime]::SpecifyKind($typed.DateTime,[DateTimeKind]::Unspecified)
Check ((Evidence (ConvertTo-IdentityNativeIso $unspecified)).Status -eq 'Timezone not established') 'Unspecified acquired a timezone'
$serialized=$typed | ConvertTo-Json -Compress | ConvertFrom-Json -DateKind String
Check ((ConvertTo-IdentityNativeIso $serialized) -eq '') 'Serialization without Kind proof was trusted'
foreach($case in @(@('bad','Invalid'),@('2026-10-21T00:00:00','Timezone not established'),@('1601-01-01T00:00:00Z','Sentinel'),@('2027-01-01T00:00:00Z','Future'))) {
    Check ((Evidence $case[0]).Status -eq $case[1]) "Unexpected qualification: $($case[1])"
}
$missing=Read-IdentityDate $null Date -Now $now
Check ($missing.Status -eq 'Absent') 'Missing evidence not distinct'
$filetime=([datetime]'2026-10-21T00:00:00Z').ToFileTimeUtc()
$ft=Read-IdentityDate ([pscustomobject]@{Date='not-parseable';Ticks=$filetime}) Date Ticks -Now $now
Check ($ft.Status -eq 'Qualified' -and $ft.Proof -eq 'AD FILETIME') 'FILETIME proof lost'
foreach($days in @(30,45,90,180)) {
    $recent=Evidence ($now.AddDays(-$days).ToString('o'))
    $old=Evidence ($now.AddDays(-$days-1).ToString('o'))
    foreach($bad in @($unknown,$missing,(Evidence 'bad'),(Evidence '2027-01-01T00:00:00Z'),(Evidence '1601-01-01T00:00:00Z'))) {
        $allUnknown=Merge-IdentityActivity @($bad,$bad) $days -Now $now
        Check ($null -eq $allUnknown.Recent -and $null -eq $allUnknown.Old) 'Unknown became inactive/active'
        $mixedOld=Merge-IdentityActivity @($old,$bad) $days -Now $now
        Check ($null -eq $mixedOld.Old -and $mixedOld.Coverage -eq 'Partial') 'Partial old evidence became global inactivity'
        $mixedRecent=Merge-IdentityActivity @($recent,$bad) $days -Now $now
        Check ($mixedRecent.Recent -eq $true -and $mixedRecent.Old -eq $false -and $mixedRecent.Coverage -eq 'Partial') 'Positive recent observation lost'
        Check ((Format-IdentityActivityCategory $mixedOld $days) -eq 'Unknown / unqualified activity') 'Unknown category says never/inactive'
    }
    $completeOld=Merge-IdentityActivity @($old,$old) $days -Now $now
    Check ($completeOld.Old -eq $true -and $completeOld.Coverage -eq 'Complete') 'Complete old evidence not retained'
}
$crossDay=Evidence '2026-10-22T00:30:00+02:00'
Check ($crossDay.Utc.ToString('yyyy-MM-dd') -eq '2026-10-21') 'Offset crossing UTC day lost'

# Inspect actual consumer rules: never execute entry points or their external actions.
$root=Split-Path $PSScriptRoot -Parent
$paths=@('SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-UsersEnrichment.ps1',
    'SmartInventory/ActiveDirectoryInventory/SmartM365-ActiveDirectory-Enrichment.ps1',
    'SmartInventory/ExchangeInventory/Migration/SmartM365-Exchange-HybridIdentity-Issues-Inventory.ps1',
    'SmartInventory/ExchangeInventory/Mailboxes/SmartM365-EXO-Mailboxes-Inventory.ps1',
    'SmartInventory/ExchangeInventory/OnPremises/Mailboxes/SmartM365-Exchange-Local-Mailboxes-Inventory.ps1')
foreach($path in $paths) {
    $tokens=$null;$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $root $path),[ref]$tokens,[ref]$errors)
    Check ($errors.Count -eq 0) "Parser: $path"
    if($path -like '*UsersEnrichment*') {
        $assignments=$ast.FindAll({param($n) $n -is [Management.Automation.Language.AssignmentStatementAst] -and
            $n.Left.Extent.Text -match "^\`$out\['(?:ResolvedLastLogonStatus|HasNoResolvedLastLogon|ResolvedLastLogonDateTime|IsEnabledWithNoResolvedLastLogon)'\]$"},$true)
        Check ($assignments.Count -eq 4) 'User activity flags not covered'
        foreach($adActivity in @($unknown,$missing,(Evidence 'bad'),(Evidence '2026-10-21T00:00:00Z'))) {
            $out=[ordered]@{}; $isEnabled=$true
            foreach($assignment in $assignments){. ([scriptblock]::Create($assignment.Extent.Text))}
            if($adActivity.Status -in @('Invalid','Unqualified legacy value')) {
                Check ($out.HasNoResolvedLastLogon -eq '' -and $out.IsEnabledWithNoResolvedLastLogon -eq '') 'Actual user flags turn unknown into never-connected'
            }
            Check ($out.ResolvedLastLogonStatus -eq $adActivity.Status) 'Actual user status lost'
        }
        Check ($ast.Extent.Text.Contains('$combinedActivity = @($adActivity, $entraSuccessful)')) 'Attempted sign-in substituted for successful activity'
    }
    if($path -like '*HybridIdentity*') {
        $rule=$ast.Find({param($n) $n -is [Management.Automation.Language.IfStatementAst] -and $n.Clauses[0].Item1.Extent.Text.Contains('$activity180.Old')},$true)
        Check ($null -ne $rule) 'Missing inactivity consumer'
        $condition=[scriptblock]::Create($rule.Clauses[0].Item1.Extent.Text)
        foreach($items in @(@($unknown,$missing),@((Evidence '2025-01-01T00:00:00Z'),$unknown),@((Evidence '2026-10-21T00:00:00Z'),$unknown))) {
            $activity180=Merge-IdentityActivity $items 180 -Now $now
            Check (-not (& $condition)) 'Actual rule 15 would flag unknown/partial/recent activity'
        }
        $activity180=Merge-IdentityActivity @((Evidence '2025-01-01T00:00:00Z')) 180 -Now $now
        Check ([bool](& $condition)) 'Actual rule 15 lost qualified inactivity'
        Check (-not $ast.Extent.Text.Contains("50 `$guid 'Ghost mailbox")) 'Missing logon still creates ghost issue'
    }
}
Write-Output "PASS: $checks identity activity checks; synthetic only."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAKqrhWF6ysxrls
# /QrXJLFCmMM7Mb231Fmcx9QujfEVsKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIEkSe8UoBvS9i1ZxYFfvPveNN0zJVMoALWB9hCiu1rUJMA0GCSqG
# SIb3DQEBAQUABIIBgJajP6jeaF/24r6d18Jc+vPB74HjI4IG3Asc2g7LwEkHZwFV
# lByyLOZm29lTDnbKuYnWGOZsfRx1fQ/qKe5ggAtEF/yNNldByoQMAA2Dn9DZCfyE
# jPEAUiViarvd15HVkiX06SeGRg3OdaHbPALBrZ/WiBkOd3R4h3u5c6pgJMelu+sc
# B6gpLAn6+Kou02SgwFlYGHgdrW/q3in7g277VdFzVZsoUlhhu2cy8mRgk3fPmMZv
# slEZvqT6KJ9+B8RldIM8GkmVDtZ2UOEUg5uLyhMWMZjmjW6UI1B7LgNhn8HXJDZO
# XiqPO4IhtlZuO7qLfyczztXvmiGxbFaiww6JM4+VWc6v04yoZFSEe12amGclVfcS
# atYzKA0OT0nL0HbFHP2MULliHHNn0ve6vFUBp9UEB4ST8RT7ZMkU4m/3SEsgoEJ6
# 2Hm99TD7yTlzeFja8QCcurMUNEl2kiMAPSaEXMYy8qEqq7eUz2hi/rcrMs0lkKFW
# 78W7hxKJlCL5+ewzmKGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjIyMTI2
# MjhaMC8GCSqGSIb3DQEJBDEiBCDqjAB6Do6DuBpVyM9soZ7tRG5l3jiXIsL9/xFu
# VdgeDjANBgkqhkiG9w0BAQEFAASCAgBphuds3t+Uggo+tONnCG4pETt9Z43+Z00l
# KtsE46NXjxMb/DvZx0QFvVkvFyXy9zWOSfhu5HxoCXYta+jCOw/WIvRXV4yt0we2
# QV16/NVLWRDrqv8rGVqZsfH4fRFfjHmugV/3LD2C9yIMNQzuM4X87MDfkJ5vpdc4
# 6OxYGrqh21DkfD49Ly0U/U2ecjzhr8D2J0JI2gdHOXyJFU/jwfFO8gE/Mx4jabQw
# nNvQL4I4vFV/CLb29XxCVa1kef3h/7b0188kpZzC7Mq1R7qkNa5tPOxunfiYuXgv
# MHtyEKi0oKjAZgLWgbF+RaMWOpuQ8+KM1rY1Dn+XN+w2K88tnhg5vo9CmBbCefCO
# H9uUua9n027Z3Qk4ALecsFqMnUiCMHefPTQRPPyP3EXPeOqzHxQNWIrkqsPIdDIh
# 7MHwWxafyl+L2otdZhDBxs/+Eywn5dUQIJNy0G5dNC/w++zw106xbq4tFpZ2ueaq
# hq5mqZWPrOFv6Z9KtLIq+MxGHnfEoSErT9dn8MwK7PaikjuJdCEBv/elm1smGB7p
# c4y0wEj4lYqJrWqpfBfqus1uEA4VSjg1+TUR6rE756kf8TcS3MoKi7PHCi79FcTI
# tOK5kvLL5Eaq+hMDE2OtEjmscchGS1x8k3ArCd+6mZV7upgvViaLKfI2NStfkY2p
# OAJiXgCAxA==
# SIG # End signature block
