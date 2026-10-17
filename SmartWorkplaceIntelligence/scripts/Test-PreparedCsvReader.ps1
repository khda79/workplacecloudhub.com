[CmdletBinding()]
param([Parameter(Mandatory)][string]$TestRoot)
# Synthetic fixtures only. The compiled strict reader must either return exactly what the
# reference parsers return (TextFieldParser tenant checks, Import-Csv column selection) or
# return null so that the reference parser is used.
$ErrorActionPreference = 'Stop'
$fixture = Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $fixture -Force | Out-Null
$module = Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force -PassThru
$script:checks = 0
function Check([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }; $script:checks++
}
function WriteFixture([string]$Name, [string]$Text, [bool]$Bom = $true) {
    $path = Join-Path $fixture $Name
    [IO.File]::WriteAllText($path, $Text, [Text.UTF8Encoding]::new($Bom))
    $path
}
# Reference tenant check: the TextFieldParser implementation used before the strict reader.
function ReferenceTenant([string]$Path, [string]$TenantKey, [string]$Delimiter, [bool]$AllowMissingTenantKey) {
    $parser = [Microsoft.VisualBasic.FileIO.TextFieldParser]::new($Path)
    try {
        $parser.SetDelimiters($Delimiter); $parser.HasFieldsEnclosedInQuotes = $true; $parser.TrimWhiteSpace = $false
        $header = $parser.ReadFields()
        if (-not $header -or @($header | Sort-Object -Unique).Count -ne $header.Length -or @($header | Where-Object {[string]::IsNullOrWhiteSpace($_)}).Count) { throw 'Empty or duplicate CSV header.' }
        $index = [array]::IndexOf($header,'TenantKey')
        if ($index -lt 0 -and -not $AllowMissingTenantKey) { throw 'TenantKey column missing.' }
        $rows = 0L
        while (-not $parser.EndOfData) {
            $fields = $parser.ReadFields(); $rows++
            if ($fields.Length -ne $header.Length) { throw "Malformed CSV at record $rows." }
            if ($index -ge 0 -and ([string]::IsNullOrWhiteSpace($fields[$index]) -or $fields[$index] -ne $TenantKey)) { throw "Empty or incompatible TenantKey at record $rows." }
        }
        "Rows=$rows;HasTenantKey=$($index -ge 0)"
    } finally { $parser.Dispose() }
}
function Outcome([scriptblock]$Action) {
    try { & $Action } catch { 'Error: ' + $_.Exception.Message }
}
$tenant = 'synthetic-test'
$t = '"synthetic-test"'
$cases = @(
    @{Name='quoted'; Text="`"TenantKey`",`"Name`",`"Note`"`r`n$t,`"a, b`",`"say `"`"hi`"`"`"`r`n$t,`"line1`r`nline2`",`"`u{00E9}t`u{00E9}`"`r`n"; Strict=$true; Columns=@('TenantKey','Note')},
    @{Name='unquoted-lf'; Text="TenantKey,Name,Note`nsynthetic-test,alpha,`nsynthetic-test,,omega"; Strict=$true; Columns=@('Name','Note')},
    @{Name='mixed'; Text="`"TenantKey`",Name,`"Note`"`r`n$t,plain value,`"`"`r`n$t,`"`",x`r`n"; Strict=$true; Columns=@('Name')},
    @{Name='header-only'; Text="`"TenantKey`",`"Name`"`r`n"; Strict=$true; Columns=@('Name')},
    @{Name='no-bom'; Text="`"TenantKey`",`"Name`"`r`n$t,`"x`"`r`n"; Bom=$false; Strict=$true; Columns=@('name')},
    @{Name='blank-line'; Text="`"TenantKey`",`"Name`"`r`n$t,`"x`"`r`n`r`n$t,`"y`"`r`n"; Strict=$false; Columns=@('Name')},
    @{Name='white-line'; Text="`"TenantKey`",`"Name`"`r`n$t,`"x`"`r`n   `r`n$t,`"y`"`r`n"; Strict=$false; Columns=@('Name')},
    @{Name='lone-cr'; Text="`"TenantKey`",`"Name`"`r$t,`"x`"`r"; Strict=$false; Columns=@('Name')},
    @{Name='inner-quote'; Text="`"TenantKey`",`"Name`"`r`n$t,ab`"c`r`n"; Strict=$false; Columns=@('Name')},
    @{Name='after-quote'; Text="`"TenantKey`",`"Name`"`r`n$t,`"ab`"c`r`n"; Strict=$false; Columns=@('Name')},
    @{Name='space-before-quote'; Text="`"TenantKey`",`"Name`"`r`n$t, `"x`"`r`n"; Strict=$false; Columns=@('Name')},
    @{Name='edge-space-value'; Text="`"TenantKey`",Name`r`n$t, padded `r`n"; Strict=$true; Columns=@('Name'); ColumnsStrict=$false},
    @{Name='tenant-edge-space'; Text="TenantKey,Name`r`n synthetic-test,x`r`n"; Strict=$false; Columns=@('Name'); ColumnsStrict=$true},
    @{Name='tenant-case'; Text="`"TenantKey`",`"Name`"`r`n`"Synthetic-Test`",`"x`"`r`n"; Strict=$false; Columns=@('Name'); ColumnsStrict=$true},
    @{Name='tenant-wrong'; Text="`"TenantKey`",`"Name`"`r`n$t,`"x`"`r`n`"other`",`"y`"`r`n"; Strict=$false; Columns=@('Name'); ColumnsStrict=$true},
    @{Name='tenant-empty'; Text="`"TenantKey`",`"Name`"`r`n`"`",`"x`"`r`n"; Strict=$false; Columns=@('Name'); ColumnsStrict=$true},
    @{Name='duplicate-header'; Text="`"TenantKey`",`"Name`",`"name`"`r`n$t,`"x`",`"y`"`r`n"; Strict=$false; Columns=@('Name')},
    @{Name='missing-tenant'; Text="`"Name`",`"Note`"`r`n`"x`",`"y`"`r`n"; Strict=$false; Columns=@('Name'); ColumnsStrict=$true},
    @{Name='short-row'; Text="`"TenantKey`",`"Name`"`r`n$t`r`n"; Strict=$false; Columns=@('Name')},
    @{Name='unclosed'; Text="`"TenantKey`",`"Name`"`r`n$t,`"open"; Strict=$false; Columns=@('Name')},
    @{Name='type-line'; Text="#TYPE Synthetic`r`n`"TenantKey`",`"Name`"`r`n$t,`"x`"`r`n"; Strict=$false; Columns=@('Name')},
    @{Name='missing-column'; Text="`"TenantKey`",`"Name`"`r`n$t,`"x`"`r`n"; Strict=$true; Columns=@('Absent'); ColumnsStrict=$false},
    @{Name='multiline-selected'; Text="`"TenantKey`",`"Name`"`r`n$t,`"a`r`nb`"`r`n"; Strict=$true; Columns=@('Name'); ColumnsStrict=$false}
)
try {
    foreach ($case in $cases) {
        $bom = if ($case.ContainsKey('Bom')) { $case.Bom } else { $true }
        foreach ($delimiter in ',', ';') {
            $name = if ($delimiter -eq ';') { "$($case.Name)_DailyStats.csv" } else { "$($case.Name).csv" }
            $path = WriteFixture $name ($case.Text.Replace(',', $delimiter)) $bom
            $reference = Outcome { ReferenceTenant $path $tenant $delimiter $false }
            $fast = [Workplace.Prepared.StrictCsv]::CheckTenant($path, [char]$delimiter, $tenant, $false)
            Check (($null -ne $fast) -eq $case.Strict) "Strict tenant path expectation failed: $name"
            if ($null -ne $fast) { Check ("Rows=$($fast.Rows);HasTenantKey=$($fast.HasTenantKey)" -ceq $reference) "Strict tenant result differs: $name" }
            $pipeline = Outcome { $r = & $module { param($p, $k) Test-PreparedTenant -Path $p -TenantKey $k } $path $tenant; "Rows=$($r.Rows);HasTenantKey=$($r.HasTenantKey)" }
            Check ($pipeline -ceq $reference) "Pipeline tenant outcome differs: $name ($pipeline / $reference)"
            $allowed = Outcome { ReferenceTenant $path $tenant $delimiter $true }
            $fastAllowed = [Workplace.Prepared.StrictCsv]::CheckTenant($path, [char]$delimiter, $tenant, $true)
            if ($null -ne $fastAllowed) { Check ("Rows=$($fastAllowed.Rows);HasTenantKey=$($fastAllowed.HasTenantKey)" -ceq $allowed) "Strict legacy result differs: $name" }
        }
        $path = WriteFixture "$($case.Name)-columns.csv" $case.Text $bom
        $columnsStrict = if ($case.ContainsKey('ColumnsStrict')) { $case.ColumnsStrict } else { $case.Strict }
        $fastRows = [Workplace.Prepared.StrictCsv]::ReadColumns($path, [char]',', [string[]]$case.Columns)
        Check (($null -ne $fastRows) -eq $columnsStrict) "Strict column path expectation failed: $($case.Name)"
        if ($null -ne $fastRows) {
            $referenceRows = @(Import-Csv -LiteralPath $path | Select-Object -Property $case.Columns)
            Check ($fastRows.Length -eq $referenceRows.Count) "Column row count differs: $($case.Name)"
            for ($i = 0; $i -lt $referenceRows.Count; $i++) {
                foreach ($column in $case.Columns) {
                    Check ([string]$fastRows[$i].PSObject.Properties[$column].Value -ceq [string]$referenceRows[$i].PSObject.Properties[$column].Value) "Column value differs: $($case.Name) row $i $column"
                    Check ($null -ne $fastRows[$i].PSObject.Properties[$column].Value) "Column value is null: $($case.Name)"
                }
            }
        }
    }
    $before = (Get-FileHash -LiteralPath $path).Hash
    [void][Workplace.Prepared.StrictCsv]::CheckTenant($path, [char]',', $tenant, $false)
    Check ((Get-FileHash -LiteralPath $path).Hash -ceq $before) 'Reader changed source bytes'
    Write-Host "PASS: $script:checks strict CSV reader checks against TextFieldParser and Import-Csv references. Synthetic data only."
} finally { Remove-Module $module; Remove-Item -LiteralPath $fixture -Recurse -Force -ErrorAction SilentlyContinue }

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAQUI72Hn+QtIBr
# gu4GsHHkUDAlqEI48sxbQbY6Vioj36CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIM37nVt3MXlSdwf1MWt1kjHeY4DF4ZKPPB6l3hLQ2lt8MA0GCSqG
# SIb3DQEBAQUABIIBgDp8z0NDxP3KC+fefV4cmXJu3GQONevWWt1HcpOc2CVx8zeS
# r0XaMS6dzcAEG3wjtaDbmNUpmeLwKfenGr1s51Uo6SKBejysPhGII8F/JQIARCLK
# XhlTTv60mOhvGQiagKhhecWRlPy++aOItmq6xS/X0eeloI+0kX4v6s5qWvg4lk40
# nVz61F2Vxfkyb9D+efn+BzxvAIGR6EhTeSiNNmiT0c6kpIx1SjkWY/SiVdRhGXX8
# cA00MbMHCAFjPFujALUoykEHQqGXJN3fysKgIa0cTA/EH1ixiF0vH2/7bkMT9n9F
# lSqSN91FIosmWTfB977qFgmSZzZ28wa7tqLcNO6G/qaHnk0yl3pWrpP4ODzmRYDx
# oqW/cn4NoOQmV5BCZjR0Te+yR9MG+TUsJvr25tXqADBMwrzl1SEcX8B0ppz3lYq8
# odH4Bq0O+GvvZx837f7VwzhgRNvtSvH3T0yACkMJC6R0jcg0zgHFczJciza2lslL
# +fLaxPE7H1kIY7j9n6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMTcyMTA1
# MzNaMC8GCSqGSIb3DQEJBDEiBCAR34r6HKwgCNfLCZHJwG56P1caCEv5wn3bMdWp
# l6CsKjANBgkqhkiG9w0BAQEFAASCAgAck2PjE5F2y5LHFYnJPCmfJnI6EIcrUYLB
# mT0IBwk4uSY8Zv90KH8SVhf9DAHLYatDWAmAiuMhH9ZRN/24gJ+CsWaMZJ0BMHWI
# 18LymNKx1n+EQh9VqWyy4FD6OiENzxh5RcKAlS/pdd5vWWcUHzQvTtjQ6lwXq8NC
# DVPIWd9e126Md5TiOYTm+AT0uxvDSVdNnwiUfqUwrbDzSSiwDpfiDGIc/+3BXEZj
# 9+j87zxON/eXq2Ez8HOPhvknzLnEW9Atsndn1OkSZdyk3yVm4x6JqDrc+E02d87k
# KJ4GMi6J7dpgiCU09i4o0BbqOsKyFZisc7fFezBFa0D5tPQtM2BRZ9M+E/DyrSHk
# ebYslV0SuLo8ZF4UGUy3l5N5X1jjPLVnjNogC+aPcga4lygvrx1+6/jd6pnjQg9O
# m3pe9/7zoVTTwc1vAz0h5UkI1BNDdPOWKAL4iHjc0s/lPNxxOb2Q81XJZnjqW2+D
# F96sAlQpRQ0kBGmYs/zrlq2bOX7+iiWU3S7O+9ttxE+sinmkq5PKsnQo99mvuccI
# UGzpLaLAV4P2GcL9khkpmc1AxczCGI6zNAlamnWodJNZkNPnDGt9cBPA4oRLqWGO
# f9Xtr0jr+9cUlA8TI55WTpxtV1VWvJ+7RFKNKNSq+dOrjHa2lFpcWwR76Iks4Atq
# EotKNmRM2Q==
# SIG # End signature block
