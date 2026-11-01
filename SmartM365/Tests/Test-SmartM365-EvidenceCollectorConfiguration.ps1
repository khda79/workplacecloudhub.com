#Requires -Version 7.0
<#
.SYNOPSIS
Tests evidence collector configuration precedence with isolated offline fixtures.
.VERSION
1.0.1
#>
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$modulePath = Join-Path $PSScriptRoot '../SmartInventory/Common/SmartM365.EvidenceCollector.Common.psd1'
Import-Module $modulePath -MinimumVersion '1.0.4' -Force
$fixtureRoot = Join-Path ([IO.Path]::GetTempPath()) ('SmartM365-EvidenceConfig-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $fixtureRoot
$preferred = '{"Marker":"preferred","DataAllRootPath":"__USE_GLOBAL__","EnableSharePointUpload":"__USE_GLOBAL__","DerivedPath":"{{DataAllRootPath}}/evidence"}'
$legacy = '{"Marker":"legacy"}'
$local = '{"Marker":"local","EnableSharePointUpload":false}'
$cases = @(
    @{ Name='preferred template and global inheritance'; Files=@{ 'json.txt.template'=$preferred }; Marker='preferred' },
    @{ Name='legacy template fallback'; Files=@{ 'json.template'=$legacy }; Marker='legacy' },
    @{ Name='preferred template wins'; Files=@{ 'json.txt.template'=$preferred; 'json.template'=$legacy }; Marker='preferred' },
    @{ Name='preferred local configuration wins'; Files=@{ 'json.txt'=$local; 'json.txt.template'=$preferred; 'json.template'=$legacy }; Marker='local' },
    @{ Name='legacy local configuration wins'; Files=@{ 'json'=$local; 'json.txt.template'=$preferred }; Marker='local' },
    @{ Name='invalid template blocks required enrichment of an existing configuration'; Files=@{ 'json.txt'=$local; 'json.txt.template'='invalid' }; Error='*template*' },
    @{ Name='invalid preferred template cannot fall back'; Files=@{ 'json.txt.template'='invalid'; 'json.template'=$legacy }; Error='*' },
    @{ Name='preferred template directory cannot fall back'; Files=@{ 'json.template'=$legacy }; Directory='json.txt.template'; Error='*not a file*' },
    @{ Name='missing templates fail explicitly'; Files=@{}; Error='*template*' },
    @{ Name='invalid local configuration cannot fall back'; Files=@{ 'json.txt'='invalid'; 'json.txt.template'=$preferred }; Error='*' },
    @{ Name='conflicting local configurations fail'; Files=@{ 'json.txt'=$local; 'json'=$legacy; 'json.txt.template'=$preferred }; Error='*Conflicting JSON transport names*' },
    @{ Name='template must be a JSON object'; Files=@{ 'json.txt.template'='[]'; 'json.template'=$legacy }; Error='*JSON object*' },
    @{ Name='one-object array is not a configuration object'; Files=@{ 'json.txt.template'='[{"Marker":"array"}]'; 'json.template'=$legacy }; Error='*JSON object*' },
    @{ Name='null template cannot fall back'; Files=@{ 'json.txt.template'='null'; 'json.template'=$legacy }; Error='*JSON object*' }
)

try {
    $index = 0
    foreach ($case in $cases) {
        $index++
        $folder = Join-Path $fixtureRoot ([string]$index)
        $null = New-Item -ItemType Directory -Path $folder
        $scriptPath = Join-Path $folder 'SmartM365-Fixture-Inventory.ps1'
        $configBase = Join-Path $folder 'SmartM365-Fixture-Inventory.local.'
        foreach ($suffix in $case.Files.Keys) {
            [IO.File]::WriteAllText($configBase + $suffix, $case.Files[$suffix], [Text.UTF8Encoding]::new($false))
        }
        if ($case.ContainsKey('Directory')) {
            $null = New-Item -ItemType Directory -Path ($configBase + $case.Directory)
        }
        $effective = [pscustomobject]@{ DataAllRootPath='fixture-root'; EnableSharePointUpload=$true }
        $caught = $null
        $actual = $null
        try { $actual = Get-SmartM365EvidenceConfig -ScriptPath $scriptPath -EffectiveConfig $effective }
        catch { $caught = $_ }

        if ($case.ContainsKey('Error')) {
            if ($null -eq $caught -or $caught.Exception.Message -notlike $case.Error) {
                throw "Expected failure was not observed: $($case.Name); actual=$caught"
            }
        }
        else {
            if ($null -ne $caught) { throw $caught }
            if ($actual.Marker -cne $case.Marker) { throw "Wrong configuration selected: $($case.Name)" }
            if ($actual.Marker -eq 'preferred' -and
                ($actual.DataAllRootPath -ne 'fixture-root' -or -not $actual.EnableSharePointUpload -or
                 $actual.DerivedPath -ne 'fixture-root/evidence')) {
                throw "Global inheritance or token resolution regressed: $($case.Name)"
            }
            if ($actual.Marker -eq 'local' -and $actual.EnableSharePointUpload) {
                throw "Explicit local override was lost: $($case.Name)"
            }
        }

        foreach ($suffix in @('json.txt.template','json.template')) {
            if ($case.Files.ContainsKey($suffix) -and
                [IO.File]::ReadAllText($configBase + $suffix) -cne $case.Files[$suffix]) {
                throw "Template content was modified: $($case.Name)"
            }
        }
        if (-not $case.Files.ContainsKey('json') -and -not $case.Files.ContainsKey('json.txt') -and
            ((Test-Path -LiteralPath ($configBase + 'json')) -or (Test-Path -LiteralPath ($configBase + 'json.txt')))) {
            throw "Template-only resolution unexpectedly created local configuration: $($case.Name)"
        }
        Write-Output "PASS: $($case.Name)"
    }
    # Exercise the actual committed WorkplaceScope template without reading private config.
    $index++
    $folder = Join-Path $fixtureRoot ([string]$index)
    $null = New-Item -ItemType Directory -Path $folder
    $templateSource = Join-Path $PSScriptRoot '../SmartInventory/M365Inventory/WorkplaceScope/SmartM365-WorkplaceScope-Inventory.local.json.txt.template'
    $templateBytes = [IO.File]::ReadAllBytes($templateSource)
    $templateHash = (Get-FileHash -LiteralPath $templateSource -Algorithm SHA256).Hash
    [IO.File]::WriteAllBytes((Join-Path $folder 'SmartM365-WorkplaceScope-Inventory.local.json.txt.template'), $templateBytes)
    $workplaceConfig = Get-SmartM365EvidenceConfig -ScriptPath (Join-Path $folder 'SmartM365-WorkplaceScope-Inventory.ps1') -EffectiveConfig (
        [pscustomobject]@{ DataAllRootPath='fixture-root'; LatestCsvFolderPath='fixture-latest'; LogAllRootPath='fixture-logs'; EnableSharePointUpload=$true }
    )
    if ($workplaceConfig.ScriptCsvLogFolderPath -cne 'fixture-root\M365\WorkplaceScope' -or
        $workplaceConfig.LatestCsvFolderPath -cne 'fixture-latest' -or -not $workplaceConfig.EnableSharePointUpload -or
        (Get-FileHash -LiteralPath $templateSource -Algorithm SHA256).Hash -cne $templateHash) {
        throw 'Committed WorkplaceScope template inheritance or read-only contract failed.'
    }
    Write-Output 'PASS: committed WorkplaceScope template with inherited output and publication settings'
    Write-Output "PASS: $index offline evidence configuration cases; no tenant connection or collection."
}
finally {
    $resolvedFixtureRoot = [IO.Path]::GetFullPath($fixtureRoot)
    $resolvedTempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\','/') + [IO.Path]::DirectorySeparatorChar
    if (-not $resolvedFixtureRoot.StartsWith($resolvedTempRoot, [StringComparison]::OrdinalIgnoreCase) -or
        [IO.Path]::GetFileName($resolvedFixtureRoot) -notmatch '^SmartM365-EvidenceConfig-[0-9a-f]{32}$') {
        throw 'Unsafe fixture cleanup path.'
    }
    Remove-Item -LiteralPath $resolvedFixtureRoot -Recurse -Force
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCABFOeoEA2SbFgX
# u6NEdMZv+XNvciK27V0JQNS6Baiio6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIM+7FvxUTjsq2+SjJMlFuMnThGr4nP3rHyLztBIBA9xmMA0GCSqG
# SIb3DQEBAQUABIIBgKjW7Z4KfkRM3HtW9y9TC+94AUNsrsprlHxFG5ifBVSoqyH6
# aX/EBAoNRKHatLhWNkg45k7Auy2aTSgz6vPK0kpvZY8rCpHmBLzdX3wi7OlNRicF
# b6NO7H0Hm/WiVA7QT8pXIFJxng0idZR4lffPFMB+6UVQMvau/saeBzcelMUfgSNs
# pvYzOQniWjN/gJ4WjEPHDulJS4aUkxrDFKy/p8BRB8TCIPhybIMqvUH7Y1d7PDW0
# HkRsKgG4dnYS/t6hNAR8CMnVOkFaW/4DsSf7o4RRR5N+zl4nptREMEj2/c3pGhA5
# 7bjoPXtrHGfMpwL08sWW98044zzGyoWkNBtR5+DjTjdgcgkSZeTii7IsS0U2+N1J
# rOettwyta76abzxlkW2/Gj1Ekwq+foI5XC+i9EyICeyh3e0C8DI89ip5T+swulZZ
# bvryyEOVDgpvdsvFQwzRbfgdsjqonyxvEu7/LZ8NKJaL6DYQsEuHBMHlB3Gym5Vm
# znvZi47LScyCvH85z6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjExMDExMzA1
# MzRaMC8GCSqGSIb3DQEJBDEiBCCho0Zap+MjD/2dP7ReIIFDv02La5vghjArnDKh
# wiiioDANBgkqhkiG9w0BAQEFAASCAgAH8FHqnK8khLlRx0Ew0/hook6XqQUaDp4z
# dH3ICvABs9kASlciuA/s17191GxJz5rLaxPCT57ozDy0HV5tFD3V/Ft1mKCdhCmo
# BdaudRnlcfjPsvDy5iCAEG4tewmtezJZOlPwKxNFdAwaRCSCUeE/H9A+HkPH7/zA
# xI0Y3WZ3kLz3WUA4BI7biPdfRHWwPTYo7MFi/5NIWMPcQZ6WGLfmU4B4quwGpCmZ
# Dj0rtLdJnPV+5u8iZUekCjtS5vPCfsMLG5BPXnZLSATLInxF+UC9hyuM0Wf0342p
# hjN9rSXtcU4WSivKPp2zmNtQr3u9wt8qVdmA+RKANFJza8+ZLzYvg6Hp6SwMcvPl
# qDC4EJ8zmJNc2vYbND0Xg4utsXY03GQadA6PYdYjRNtnHiEYrMThL3gh0bK517Ie
# yZKF0Qz9y/Isp8aGf0G0m9cn0a+9Uy0eEt/nn9Kdm6X8joJVBkJJ7oTGiDM+EEU6
# 6Z9aCHVTlb8Iu57XGRxph+YC8zidTXtlFSJeUdQ0K8uNy8vP4flrP9Dd7nBmvluu
# +05W10B/rX9H/BdgMTVqra/SU8gF1uCUqt4vLUD5zQ74SZUiy8NLCXC+AQTwKqVc
# 2ZEkPqJr1JUtFDZ2QDIDD8OHs7qQkam8FztjPkIlaoH8qNWj8W/+/8TRKJWtTR4S
# ULU1GuMiUg==
# SIG # End signature block
