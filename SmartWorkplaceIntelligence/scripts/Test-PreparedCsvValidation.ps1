[CmdletBinding()]
param([int]$BenchmarkRows=20000)
# Only generated fixtures in a uniquely owned temporary directory. No production run.
$ErrorActionPreference='Stop'
Add-Type -Path (Join-Path $PSScriptRoot 'PreparedCsvValidation.cs') -ReferencedAssemblies @('System.Runtime','System.Collections','System.Console','System.Security.Cryptography','Microsoft.VisualBasic.Core','System.IO','System.Text.RegularExpressions','System.Diagnostics.TraceSource')
$root=Join-Path ([IO.Path]::GetTempPath()) ('PreparedCsvTests-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root | Out-Null
$culture=[Globalization.CultureInfo]::CurrentCulture
$checks=0
function Assert($Condition,[string]$Message) { if(-not $Condition){throw $Message};$script:checks++ }
function Legacy([string]$Path,[string[]]$Names,[string[]]$Types,[string[]]$Keys) {
    $state=@{Rows=0L;InvalidValues=0L;DuplicateKeys=0L}
    $seen=[Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    Import-Csv -LiteralPath $Path | ForEach-Object {
        $row=$_;$state.Rows++
        $values=@($Keys | ForEach-Object {([string]$row.$_).Trim()})
        if(@($values | Where-Object {$_ -eq ''}).Count){$state.InvalidValues++}
        $key=($values | ForEach-Object {$_.Length.ToString()+':'+$_}) -join '|'
        if(-not $seen.Add($key)){$state.DuplicateKeys++}
        for($i=0;$i -lt $Names.Count;$i++){
            $value=[string]$row.($Names[$i]);$type=$Types[$i]
            if([string]::IsNullOrWhiteSpace($value) -or $type -eq 'string'){continue}
            try {
                switch($type){
                    int64 {$null=[long]::Parse($value,[Globalization.CultureInfo]::InvariantCulture)}
                    double {$n=[double]::Parse($value,[Globalization.CultureInfo]::InvariantCulture);if([double]::IsNaN($n) -or [double]::IsInfinity($n)){throw 'nonfinite'}}
                    decimal {$null=[decimal]::Parse($value,[Globalization.CultureInfo]::InvariantCulture)}
                    dateTime {$null=[datetime]::Parse($value,[Globalization.CultureInfo]::InvariantCulture)}
                    boolean {$null=[bool]::Parse($value)}
                }
            } catch {$state.InvalidValues++}
        }
    }
    $state
}
try {
    $path=Join-Path $root 'types.csv'
    $names=[string[]]@('TenantKey','Key','Integer','Number','Amount','Date','Flag','Note')
    $types=[string[]]@('string','string','int64','double','decimal','dateTime','boolean','string')
    $keys=[string[]]@('TenantKey','Key')
    $rows=@(
        [pscustomobject]@{TenantKey='synthetic';Key='a';Integer='9223372036854775807';Number='1.25';Amount='1.25';Date='2026-10-05T09:08:07+02:00';Flag='true';Note="comma, quote `" and`nnewline"},
        [pscustomobject]@{TenantKey='synthetic';Key=' A ';Integer='9223372036854775808';Number='NaN';Amount='invalid';Date='31/07/2026 00:00:00';Flag='yes';Note='x'},
        [pscustomobject]@{TenantKey='synthetic';Key='';Integer='8794.0';Number='Infinity';Amount='';Date='05/10/2026 00:00:00';Flag='';Note=''},
        [pscustomobject]@{TenantKey='synthetic';Key='b';Integer='-9223372036854775808';Number='-1e3';Amount='1,234.50';Date='2026-10-05';Flag='FALSE';Note=''}
    )
    $rows | Export-Csv -LiteralPath $path -NoTypeInformation
    foreach($locale in @('fr-FR','en-US')){
        [Globalization.CultureInfo]::CurrentCulture=[Globalization.CultureInfo]::GetCultureInfo($locale)
        $old=Legacy $path $names $types $keys
        $new=[Workplace.Prepared.CsvValidation]::Check($path,$names,$types,$keys,'synthetic',$false)
        foreach($field in @('Rows','InvalidValues','DuplicateKeys')) { Assert ($old.$field -eq $new.$field) "Parity failed: $locale/$field" }
        Assert ($new.SHA256 -eq (Get-FileHash -LiteralPath $path).Hash) 'Hash differs from scanned file'
        Assert (@($new.Issues | Where-Object Column -eq 'Integer').Count -eq 1) 'Column diagnostics missing'
        Assert (($new.Issues | Where-Object Column -eq 'Integer').Count -eq 2) 'Column rejection count differs'
        Assert (($new.Issues | ConvertTo-Json -Depth 5) -notmatch 'synthetic|8794|922337') 'Diagnostics leaked raw data'
    }
    $wrong=[Workplace.Prepared.CsvValidation]::Check($path,$names,$types,$keys,'other',$false)
    Assert (($wrong.Issues | Where-Object Rule -eq 'TenantIdentity').Count -eq 4) 'Cross-tenant rows accepted'
    $schema=[Workplace.Prepared.CsvValidation]::Check($path,$names,$types,$keys,'synthetic',$true)
    Assert ($schema.Rows -eq 0 -and $schema.Columns -eq 8) 'Schema-only contract changed'
    $report=Join-Path $root 'contract.json.txt'
    @{tables=@(@{table='Types';file='types.csv';uniqueKey=$keys;columns=@(for($i=0;$i -lt $names.Count;$i++){@{name=$names[$i];type=$types[$i]}})})} | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $report
    $json=& (Get-Process -Id $PID).Path -NoProfile -File (Join-Path $PSScriptRoot 'Test-PreparedEvidence.ps1') -Root $root -ContractPath $report
    Assert ($LASTEXITCODE -eq 1) 'CLI failed to reject invalid rows'
    $parsed=($json -join "`n") | ConvertFrom-Json
    Assert (-not $parsed.Passed -and $parsed.Files[0].Issues.Count -gt 0) 'CLI JSON contract or issue details broken'
    foreach($bad in @('"Key","Key"' + "`n" + '"a","b"', '"Key","Note"' + "`n" + '"a"', '"Key","Note"' + "`n" + '"a","unterminated')){
        [IO.File]::WriteAllText($path,$bad)
        $rejected=$false
        try{$null=[Workplace.Prepared.CsvValidation]::Check($path,@('Key'),@('string'),@('Key'),'',$false)}catch{$rejected=$true}
        Assert $rejected 'Malformed CSV accepted'
    }
    if($BenchmarkRows -gt 0){
        $writer=[IO.StreamWriter]::new($path,$false,[Text.UTF8Encoding]::new($false))
        try{
            $writer.WriteLine('TenantKey,Key,Integer,Number,Amount,Date,Flag,Note')
            for($i=0;$i -lt $BenchmarkRows;$i++){$writer.WriteLine('synthetic,k'+$i+',8794,1.25,2.5,2026-10-05T09:08:07Z,true,synthetic')}
        }finally{$writer.Dispose()}
        $timer=[Diagnostics.Stopwatch]::StartNew();$old=Legacy $path $names $types $keys;$oldSeconds=$timer.Elapsed.TotalSeconds
        $timer.Restart();$new=[Workplace.Prepared.CsvValidation]::Check($path,$names,$types,$keys,'synthetic',$false);$newSeconds=$timer.Elapsed.TotalSeconds
        Assert ($old.Rows -eq $new.Rows -and $new.Rows -eq $BenchmarkRows -and $new.InvalidValues -eq 0 -and $new.DuplicateKeys -eq 0) 'Benchmark equivalence failed'
        [pscustomobject]@{Mode='Synthetic local; compilation excluded';Rows=$BenchmarkRows;LegacySeconds=$oldSeconds;CompiledScanAndHashSeconds=$newSeconds;Ratio=$oldSeconds/$newSeconds}|Format-List
    }
    Write-Host "PASS: $checks streaming validation checks. Synthetic fixtures only."
} finally {
    [Globalization.CultureInfo]::CurrentCulture=$culture
    # Unique directory allocated by this test, never a caller-supplied data root.
    if((Split-Path $root -Parent) -eq ([IO.Path]::GetTempPath().TrimEnd('\')) -and (Split-Path $root -Leaf) -like 'PreparedCsvTests-*') {Remove-Item -LiteralPath $root -Recurse -Force}
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCABNMPCLVgE4oDD
# iBfo5Ed7UYi9bQNpsLFvrZRs7liUAKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIECJRv9NYNG30nK1pQVRMTHLiFR3CNPTslOwpbPXGS46MA0GCSqG
# SIb3DQEBAQUABIIBgGH6nf+mDKtR442ESGWIr9GHX6ZhU+Vc8spFW68lLW+IZENj
# wcJ6uu91JRQX9H88xDPIWNrDpvNPdB1Ql+J+6ViAY/bRLvDr62YcwDiwL9pGBDar
# wclZo4y/EpBvP3Pgg/rF1rsppVdbMGVbMW+2HbIUPi7olq0OphQimF7nXKqFgjOg
# zKkXXLuD0UdEF6H5oKDwQKCd1lO/6xiYnnlNeSHYBKO8hnKxKFt3sBz6Gh+SrX4b
# Qw0CsuAGM7wEXddLtNPZpgNMNBnqjXjvRQPaaoCW6TTKVKmcvT791SM5ap//yu+l
# oujtLSfJo/rQF2RBuMMOUuG9fZTsTl2HOZ80TS3u3dtAIBbuZ6QHY/BOQnmhil3m
# ZHWWhR8HbIhzGVai+PebJFznbRzZKvdn/E0A6v6yys6VGcz98F6BSh58k5eIlDDL
# NQb+6WJm+Jks7aaf6g9m2mxbWUa5gnPiunHyQNV1H24YeHErnkFMHepQZ8VQOEUp
# sVYz5dklFJv4lGhwi6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTUxNzA5
# MzhaMC8GCSqGSIb3DQEJBDEiBCBtgTfqNEQtRt7NLcFKuPc5FnesRHNE5WJPUlr6
# IKGisjANBgkqhkiG9w0BAQEFAASCAgBON6kIIz+1n5IS+OUVLXGtUcVmPnbMdXqg
# bT+OSeqF1JfA7BTEa+m34yVRPedkW7toHUure0GIe+QsW1oIMSx1IMtM7ShZtvxv
# NtPH/wNZnkEbI97yJ/GLQVEluI7V2Idpxn5581QKeMEi5hwUyX6hPs1HA0S5KBfh
# DqTrNQ4o00q4DE0ornt3wzEK4kuDlXXnLtTFXeqAbnHAVjd7j8Ou2qVb04eojgY5
# I9eydy6MklHidAkSCIA639JylYARDCHtrhvHZuGVpJGFuDuCU/YSrmvvTDaYmikv
# maL7vnSpYgcANh4ZEf+OO7nFDB1XsrIhZivs9Vy2APAyeo6YsbCKupPo+85dBwqi
# YPfhvIW2788t+ne+e3y9bqItH4WI47EPKKEt0tXB8dEfVzHxNqybgA+ggxs3RZyA
# eQIzUFJP/zDEBVMdul1G6XKGSjkJxPFpl7fn9Rj8S8cR8xBo/xEwv6xJpsKFBS1J
# 0IuGxtTstVmgTVc/yAHj64Vb7HSbWonQr/9pwcv6SWskefK18emajc4cTqOyIP20
# Y3ajaRgk7+ID0M4gTagX92ZNdHOtswsngewkwTeQzKV9NSu1Pv1YdCWMJL+D8EJU
# Q+sxdKrz/eYANE7XqGEyewYLKE6U2w2m2a3rDx/Dz0UJMl6xHqX83Xwcj0GfgOFk
# P7nR+cx95Q==
# SIG # End signature block
