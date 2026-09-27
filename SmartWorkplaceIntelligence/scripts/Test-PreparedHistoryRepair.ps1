[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot)
# Local synthetic fixtures only. Never execute the tenant entry point or access a share.
$ErrorActionPreference='Stop'
$fixture=Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
$source=Join-Path $fixture 'source';$work=Join-Path $fixture 'work';$output=Join-Path $source 'DATA-POWERBI'
$config=Join-Path $fixture 'config';$contractPath=Join-Path $config 'prepared-source-contract.json'
$current=Join-Path $source 'DATA-LAST/Example.csv'
$history=Join-Path $source 'DATA-ALL/Example/WeeklyHistory/2026-W28/Example.csv'
$empty=Join-Path $source 'DATA-ALL/Example/WeeklyHistory/2026-W29/Example.csv'
$outside=Join-Path $source 'DATA-ALL/Example/WeeklyHistory/2026-W30/Example.csv'
foreach($p in $current,$history,$empty,$outside,$contractPath){New-Item -ItemType Directory -Path (Split-Path $p -Parent) -Force | Out-Null}
function WriteFixture([string]$Path,[string]$Text){[IO.File]::WriteAllText($Path,$Text,[Text.UTF8Encoding]::new($true))}
$valid="`"TenantKey`",`"Value`"`r`n`"synthetic-test`",`"current`"`r`n"
$old="`"Value`",`"Note`"`r`n`" 001 `",`"multiline`r`ntext, quoted `"`"value`"`"`"`r`n`"été`",`"`"`r`n"
WriteFixture $current $valid;WriteFixture $history $old;WriteFixture $empty "`"Value`",`"Note`"`r`n";WriteFixture $outside $valid
[IO.File]::SetLastWriteTimeUtc($history,[datetime]::new(2026,7,10,12,0,0,[DateTimeKind]::Utc))
$oldTime=(Get-Item $history).LastWriteTimeUtc
$originalHash=(Get-FileHash $history).Hash;$currentHash=(Get-FileHash $current).Hash;$outsideHash=(Get-FileHash $outside).Hash
@{currentFiles=@('Example.csv');mappingFiles=@();dailyFiles=@();history=@(@{root='DATA-ALL/Example';file='Example.csv'})} |
    ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $contractPath
$module=Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force -PassThru
$repairModule=Import-Module (Join-Path $PSScriptRoot 'Repair-PreparedHistoryTenantKeys.psm1') -Force -PassThru
& $module {param($path) $script:ProductRoot=$path} $fixture
$pipeline=@{DataRoot=$source;OutputRoot=$output;WorkRoot=$work;TenantKey='synthetic-test';AccountClassificationConfigPath=(Join-Path $fixture 'unused.psd1');ValidateOnly=$true}
$repair=@{DataRoot=$source;TenantKey='synthetic-test';Weeks=@('2026-W28','2026-W29');ExpectedFileCount=2;SourceContractPath=$contractPath}
$script:checks=0
function Reject([scriptblock]$Action,[string]$Pattern){
    $message='';try {& $Action | Out-Null} catch {$message=$_.Exception.Message}
    if($message -notmatch $Pattern){throw "Expected '$Pattern', got '$message'"};$script:checks++
}
function Check([bool]$Condition,[string]$Message){if(-not $Condition){throw $Message};$script:checks++}
try {
    Reject {Invoke-PreparedEvidencePipeline @pipeline} 'failed \(2 files\)'
    Reject {Invoke-PreparedEvidencePipeline @pipeline} 'DATA-ALL/Example/WeeklyHistory/2026-W28/Example.csv: TenantKey column missing'
    Reject {Invoke-PreparedEvidencePipeline @pipeline -AllowLegacyTenantless} 'Global tenantless bypass'
    $preview=Invoke-PreparedHistoryTenantRepair @repair
    Check ($preview.Status -eq 'PreviewOnly' -and $preview.Files -eq 2 -and -not $preview.Applied) 'Preview failed'
    Check ((Get-FileHash $history).Hash -eq $originalHash -and -not(Test-Path (Join-Path $source 'DATA-REPAIR-BACKUPS'))) 'Preview changed source/created backups'
    $bad=$repair.Clone();$bad.ExpectedFileCount=3
    Reject {Invoke-PreparedHistoryTenantRepair @bad -Apply} 'Repair count differs'
    $bad=$repair.Clone();$bad.Weeks=@('../DATA-LAST')
    Reject {Invoke-PreparedHistoryTenantRepair @bad -Apply} 'Explicit valid'
    WriteFixture $empty "`"TenantKey`",`"Value`"`r`n`"wrong`",`"row`"`r`n"
    Reject {Invoke-PreparedHistoryTenantRepair @repair -Apply} 'incompatible TenantKey'
    Check ((Get-FileHash $history).Hash -eq $originalHash) 'Wrong-tenant failure modified another source'
    WriteFixture $empty "`"Value`",`"Note`"`r`n`"missing field`"`r`n"
    Reject {Invoke-PreparedHistoryTenantRepair @repair -Apply} 'Malformed record'
    WriteFixture $empty "`"Value`",`"Value`"`r`n"
    Reject {Invoke-PreparedHistoryTenantRepair @repair -Apply} 'duplicate CSV header'
    WriteFixture $empty "`"Value`",`"Note`"`r`n"
    $lock=[IO.File]::Open((Join-Path $source '.prepared-source.lock'),'OpenOrCreate','ReadWrite','None')
    try {
        Reject {Invoke-PreparedHistoryTenantRepair @repair -Apply} 'another process|autre processus'
        Reject {Invoke-PreparedEvidencePipeline @pipeline} 'another process|autre processus'
    } finally {$lock.Dispose()}
    [IO.File]::WriteAllText($empty,"`"Value`",`"Note`"`r`n",[Text.Encoding]::Unicode)
    $result=Invoke-PreparedHistoryTenantRepair @repair -Apply
    Check ($result.Applied -and $result.Files -eq 2 -and $result.Status -eq 'Completed') 'Repair result failed'
    $audit=Get-Content (Join-Path $result.BackupPath 'repair.json') -Raw | ConvertFrom-Json
    Check ($audit.Status -eq 'Completed' -and @($audit.Files | Where-Object Status -ne 'Repaired').Count -eq 0) 'Audit incomplete'
    $backup=Join-Path $result.BackupPath 'originals/DATA-ALL/Example/WeeklyHistory/2026-W28/Example.csv'
    Check ((Get-FileHash $backup).Hash -eq $originalHash) 'Original-byte backup missing'
    Check ((Get-Item $history).LastWriteTimeUtc -eq $oldTime) 'Historical timestamp was not preserved'
    Check ((Get-FileHash $current).Hash -eq $currentHash -and (Get-FileHash $outside).Hash -eq $outsideHash) 'Current or non-selected week changed'
    $rows=@(Import-Csv $history)
    Check ($rows.Count -eq 2 -and $rows[0].Value -ceq ' 001 ' -and $rows[1].Value -ceq 'été' -and $rows[0].Note -ceq "multiline`r`ntext, quoted `"value`"") 'Original decoded values changed'
    $repeat=Invoke-PreparedHistoryTenantRepair @repair -Apply
    Check ($repeat.Status -eq 'NoMissingTenantKey' -and -not $repeat.Applied) 'Repeat not idempotent'
    $preflight=Invoke-PreparedEvidencePipeline @pipeline
    Check ($preflight.Identity.CsvFiles -eq 4 -and $preflight.Identity.Rows -eq 4 -and -not $preflight.Publication) 'Full row preflight failed after repair'
    Check (-not(Test-Path $output)) 'Repair/preflight generated prepared outputs'
    $plan=@(Get-PreparedSourcePlan -DataRoot $source);$snapshot=Join-Path $fixture 'snapshot'
    foreach($entry in $plan){$to=Join-Path $snapshot $entry.Relative;New-Item -ItemType Directory -Path (Split-Path $to -Parent) -Force | Out-Null;Copy-Item $entry.SourcePath $to}
    $snap=& $module {param($p,$s) Test-PreparedSourceTenants -Plan $p -TenantKey 'synthetic-test' -SnapshotRoot $s} $plan $snapshot
    Check ($snap.Rows -eq $preflight.Identity.Rows) 'Snapshot and preflight differ'
    WriteFixture $current "`"TenantKey`",`"Value`"`r`n`"wrong`",`"row`"`r`n"
    Reject {Invoke-PreparedEvidencePipeline @pipeline} 'incompatible TenantKey'
    WriteFixture $current "`"TenantKey`",`"Value`"`r`n`"`",`"row`"`r`n"
    Reject {Invoke-PreparedEvidencePipeline @pipeline} 'incompatible TenantKey'
    WriteFixture $current "`"Value`"`r`n`"row`"`r`n"
    Reject {Invoke-PreparedEvidencePipeline @pipeline} 'DATA-LAST/Example.csv: TenantKey column missing'
    Write-Host "PASS: $script:checks repair, exact-value preservation, backup, lock, preflight and snapshot checks. Synthetic data only."
} finally {Remove-Module $module;Remove-Module $repairModule}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAKNVjvs6mHttLC
# ZZNJen+YtQfFTwQBjAl0yYzLrBE+sqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIMnfBVrATZKqAvOV82ujX11abvfTa9CXoNe5LWqSL3ZSMA0GCSqG
# SIb3DQEBAQUABIIBgDIsl9oTto6EXU5Lj4qFn1vyOrmZQdm5JLupQ9bscgW1bP76
# SNIZHWZoYmVKVrhv6elSDzwxg7EUqcB/6rXEXjqEsIrC0kc1rvMDYPCLuGzhPIjV
# ireF6sB5SzELUXlRAw73ZgjIOEB5iWzkLhZVP9BHCRKHCl6QflXlNK+difTJ7RHs
# j4FBTMkFu1cPUoQow1+qUNTny6q7WX6Bj1Xcm324yTEXTfBiwkLZ4GgUyQJyQuoq
# fgOe5vx8z0zLxE9P6i7tTJZ2axGWvLX/++KZbPzIL+CEZ70HQYJSBbkXY5PLTLlt
# qwLmbiM9O6AwMlcO9fmP4fesk9sflX7xbDgNLHdXVbkBMl0Ou9UvawxWFbtWHLKv
# Qt3FIuyJiCcs9zLNOg7sTeHShopBUlC1tkppSZOZxtQPeP93TgNPK0A2/tuhqqT9
# MS8y7c34/MsnEF+60f4tsnMawEewyS/f93/qQ8qX4BBaR3/xz/rW//pLZFaFaahX
# Ov8AdztPW650R0/zAqGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjA5MjcwMDMz
# NTFaMC8GCSqGSIb3DQEJBDEiBCC470lVlSF3uYnLhUZPmn2/ewpqFmvuq5z74NRC
# e5jk0jANBgkqhkiG9w0BAQEFAASCAgBBCpF7AJaFk7r5roB5zxF+IRkqF5H+ZRr5
# SlvpgpwbipNLk3bOtlCdwddw+hmKxSvxVey0wpq2jofzSprzux2R6izrAAP69KWq
# ptUAAEcT2o45LLJ2dWWQKuR4L9brIzlWD3YoXM7hiaZMcFD/hCg86EAWBGHfWIlL
# Cp5GX948fD/LJHSSab137LOBLxnyriklhMdn5v6+Ba2OYIFossoadxUjoKcD2nbE
# 06UgL/ZnxOMISmC0xjXfPFiTXmt/aaxC/Mdg6fgybqRZseKe1uViyoVUaI7mUAag
# xSYhzaIC8v7Wj+egzFf+qJkWiRMcnvPKPiDPV4md1GNp3iAzu0ZWIcsLggG5uC+G
# ZIC7tUBTWS/K5Q2yy5dq6/phBl9PqjVFBAM7Uky3LwOmpSHKAuand7Ii3Nj2j/JU
# ly2nrHBluf5AjyWDAyOKkprysHOwZ5grLL2CYbDWMZlz1/hZLT+0xrJYCAtZPoXg
# ZwFBu9QbXChhCpU/ntDauFx5frQ6wm8QID2wYk8L9QBIUXDClcMBQOOrOWWnGFBH
# wI40Qm+9tS0I9VuibdzygK8G5ENGH11rVEHHuvaxPhdDe2yCt7Xx/Cjfim93OlZn
# MeUk2JoMNm9KjS5dljTx4Ci9lRcvpkcRLmAM38luRja5v0giSGcGnhKEsQJMuiXN
# SJJBiM3jdA==
# SIG # End signature block
