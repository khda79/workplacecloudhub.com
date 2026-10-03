<#
.SYNOPSIS
    Offline Windows PowerShell 5.1 farm diagnostic launcher routing test.
.VERSION
    1.0.1
#>
#Requires -Version 5.1
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$toolkit = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$launcher = Join-Path $toolkit 'Start-SmartM365-SharePointMigration-FarmDiagnostic.cmd'
$runLauncher = Join-Path $toolkit 'Start-SmartM365-SharePointMigration-FarmDiagnostic-Run.cmd'
$root = Join-Path $PSScriptRoot ('.farm-launcher-test-' + [guid]::NewGuid().ToString('N'))
if (-not $root.StartsWith($PSScriptRoot + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) {
    throw 'Unsafe test path.'
}
try {
    $project = Join-Path $root 'Migrations\Synthetic'
    $older = Join-Path $project 'ShareGate\Diagnostics\Analysis-old'
    $newer = Join-Path $project 'ShareGate\Diagnostics\Analysis-new'
    $alpha = Join-Path $root 'Migrations\Alpha\ShareGate\Diagnostics\Analysis'
    $missing = Join-Path $root 'Migrations\Missing'
    $logs = Join-Path $root 'Migrations\logs'
    $reports = Join-Path $root 'Migrations\reports'
    $template = Join-Path $root 'Migrations\_Template\ShareGate\Diagnostics\Analysis'
    $invalidCsv = Join-Path $root 'Migrations\Z-Invalid\ShareGate\Diagnostics\Analysis'
    $scripts = Join-Path $root 'Scripts\Diagnostics'
    [void](New-Item -ItemType Directory -Path $older,$newer,$alpha,$missing,$logs,$reports,$template,$invalidCsv,$scripts -Force)
    $oldCsv = Join-Path $older 'AccessFailures-5min.csv'
    $newCsv = Join-Path $newer 'AccessFailures-5min.csv'
    $newline = [Environment]::NewLine
    ('WindowUtc,Lines' + $newline + '2026-10-02 20:00 UTC,1') | Set-Content -LiteralPath $oldCsv -Encoding UTF8
    ('WindowUtc,Lines' + $newline + '2026-10-02 22:00 UTC,9' + $newline + '2026-10-02 22:05 UTC,4') | Set-Content -LiteralPath $newCsv -Encoding UTF8
    ('WindowUtc,Lines' + $newline + '2026-10-02 21:00 UTC,2') | Set-Content -LiteralPath (Join-Path $alpha 'AccessFailures-5min.csv') -Encoding UTF8
    ('WindowUtc,Lines' + $newline + '2026-10-02 23:00 UTC,3') | Set-Content -LiteralPath (Join-Path $template 'AccessFailures-5min.csv') -Encoding UTF8
    ('WindowUtc,Lines' + $newline + 'invalid,1') | Set-Content -LiteralPath (Join-Path $invalidCsv 'AccessFailures-5min.csv') -Encoding UTF8
    (Get-Item -LiteralPath $oldCsv).LastWriteTimeUtc = [datetime]::UtcNow.AddHours(-2)
    (Get-Item -LiteralPath $newCsv).LastWriteTimeUtc = [datetime]::UtcNow
    $fake = Join-Path $scripts 'SmartM365-SharePointMigration-FarmDiagnostic.ps1'
    @'
[CmdletBinding()]
param([string]$Project,[string]$ToolkitRoot,[string]$ShareGatePeaksCsv,[int]$WindowMinutes,[switch]$DryRun)
$value = '{0}|{1}|{2}|{3}' -f $Project,$ShareGatePeaksCsv,$WindowMinutes,[bool]$DryRun
[IO.File]::WriteAllText((Join-Path $ToolkitRoot 'invoked.txt'),$value)
if (Test-Path -LiteralPath (Join-Path $ToolkitRoot 'force-error.txt')) { Write-Host 'Simulated diagnostic failure.'; exit 23 }
'@ | Set-Content -LiteralPath $fake -Encoding UTF8
    $marker = Join-Path $root 'invoked.txt'

    $preview = @(& $launcher -Project Synthetic -ToolkitRoot $root -PreviewOnly 2>&1)
    if ($LASTEXITCODE -ne 0) { throw "Launcher preview failed: $($preview -join ' ')" }
    if (Test-Path -LiteralPath $marker) { throw 'PreviewOnly executed the diagnostic.' }
    if (($preview -join ' ') -notmatch 'Host: Windows PowerShell 5\.1' -or ($preview -join ' ') -notmatch '2 UTC windows') {
        throw 'Launcher did not select Windows PowerShell 5.1 and both CSV windows.'
    }
    if (-not ($preview -join ' ').Contains($newCsv)) { throw 'Latest ShareGate CSV was not selected.' }
    $prompted = @('3' | & $launcher -ToolkitRoot $root -PreviewOnly 2>&1)
    $menuText = $prompted -join ' '
    if ($LASTEXITCODE -ne 0 -or -not $menuText.Contains('Project: Synthetic')) {
        throw "Prompted project selection failed: $($prompted -join ' ')"
    }
    if ($menuText -notmatch 'Available migration projects:' -or $menuText -notmatch '1\. Alpha \| CSV .* \(1 window\)' -or $menuText -notmatch '2\. Missing \| CSV missing' -or $menuText -notmatch '3\. Synthetic' -or $menuText -notmatch '4\. Z-Invalid \| CSV invalid or inaccessible' -or $menuText -notmatch '0\. Cancel' -or $menuText -match '_Template' -or $menuText -match '\d+\.\s+(logs|reports)\s+\|') {
        throw 'The project menu was not ordered, annotated or filtered correctly.'
    }
    $launcherSource = [IO.File]::ReadAllText((Join-Path $toolkit 'Scripts\Diagnostics\SmartM365-SharePointMigration-FarmDiagnosticLauncher.ps1'))
    if (-not $launcherSource.Contains('Select a project number')) { throw 'The interactive prompt is not in English.' }
    $cancelled = @('0' | & $runLauncher -ToolkitRoot $root 2>&1)
    if ($LASTEXITCODE -ne 0 -or -not ($cancelled -join ' ').Contains('Selection cancelled') -or (Test-Path -LiteralPath $marker)) {
        throw 'Cancellation did not stop the Run launcher before collection.'
    }
    $unavailable = @('2' | & $runLauncher -ToolkitRoot $root 2>&1)
    if ($LASTEXITCODE -eq 0 -or -not ($unavailable -join ' ').Contains('CSV missing') -or (Test-Path -LiteralPath $marker)) {
        throw 'A project without ShareGate CSV was accepted.'
    }
    $invalid = @(@('99','99','99') | & $launcher -ToolkitRoot $root -PreviewOnly 2>&1)
    if ($LASTEXITCODE -eq 0 -or -not ($invalid -join ' ').Contains('No valid project number') -or (Test-Path -LiteralPath $marker)) {
        throw 'Invalid menu choices were accepted.'
    }

    $dry = @(& $launcher -Project Synthetic -ToolkitRoot $root 2>&1)
    if ($LASTEXITCODE -ne 0) { throw "Launcher DryRun failed: $($dry -join ' ')" }
    if ((Get-Content -LiteralPath $marker -Raw) -cne ('Synthetic|{0}|30|True' -f $newCsv)) {
        throw 'DryRun did not forward the project, latest CSV and window margin.'
    }
    $real = @('3' | & $runLauncher -ToolkitRoot $root 2>&1)
    if ($LASTEXITCODE -ne 0) { throw "Dedicated Run launcher failed: $($real -join ' ')" }
    if ((Get-Content -LiteralPath $marker -Raw) -cne ('Synthetic|{0}|30|False' -f $newCsv)) {
        throw 'Dedicated Run launcher did not remove DryRun.'
    }
    [IO.File]::WriteAllText((Join-Path $root 'force-error.txt'),'1')
    $failed = @(& $launcher -Project Synthetic -ToolkitRoot $root 2>&1)
    if ($LASTEXITCODE -eq 0 -or ($failed -join ' ').Contains('Farm diagnostic completed.') -or -not ($failed -join ' ').Contains('exit code 23')) {
        throw 'The launcher reported success after the diagnostic failed.'
    }
    Write-Output 'Farm diagnostic launcher offline host and routing tests passed.'
}
finally {
    if ((Test-Path -LiteralPath $root) -and $root.StartsWith($PSScriptRoot + [IO.Path]::DirectorySeparatorChar,[StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $root -Recurse -Force
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCdt1O/AIycV2hB
# rkG45rpPNiNMZYZk2iRPoAVPzYBhQKCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIDW6OVtxtY1kE3nKRZoiBo9mcf4dpHk6QhV024KAhLo3MA0GCSqG
# SIb3DQEBAQUABIIBgCCyzEiixA6NKiAIYJ4wBBeQZ5rahDxqolHlBQ9y+XbM2JWZ
# K/fLaSYs/JQmBMjUEJhLeEyBiNtMc0tQQqqhf3/Bz/D8IXiB8bYZ38tb/++UsZ1g
# 7/51xowUm5FrapqDtVZlbhhSxP3n9Wx7zDejiLsxCnnpY2s2c8INCAX30cPLF0ei
# ddXv+ZrqCr4i1ADrZM+Itocn2afdm0zE/0n0YlRCQ5GiVs1wC/rOC0rj8eGVeF/7
# 0VMzBhAw/L6Idk8PvY02lXoxmBJ6Tk0lHjiGZ4pUJDbvb3mexNPKbDZuFrHUYau1
# nH4+gkzp/3kLJmGnc7ra5eacaPsC3AS186o2CF59AASp5ufXxlCYmsax4lSFudTy
# 4gIVO5qAW2oqFkWWaBvtme2PgUP5qcPTIIJHpBZRIsRWBS9Fc3ittFbbKN/XZy98
# LtLSXrG32IpS4fSn2uQ55IdgzHnGmU13huDc+jFCLgQl4pHPZUbQegL4MLmREddC
# gEjfv8q1togcEStzFaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMDMxMzI2
# MDFaMC8GCSqGSIb3DQEJBDEiBCBaqnWKRJDS99pmI4Z0i84aiVnc+lWtXNG3r0U+
# 5ThiOjANBgkqhkiG9w0BAQEFAASCAgA2LMTd9eRkQyYhAv60lyD/uyFJVnea/JdO
# bvhcWdx77DsFmmgEamZujSkN2cuURHNOIU90D7sVtqPPV1dBmFpOpyhcOGR6D11v
# VhqAEK03q6dc068XpeTsPEKtQJ9Ar1K5Nu5uxRjSmEa8fctcAKKHLQlAI4Jh74xC
# HehGHMYucrm9zBR2gZexmQsP1GhwGTO6AC7XfDMh4L477ysGRf7u8XJ7N4SXPgqT
# mn5rqGdwMEnb3JOEL8Zo1SfYjNUbR3+GgnDd/DkDz2/EDmX5W+pzjNHXQuiqfbXA
# M+pUGPjSP6atlwNGRFiac4ZFWl//4paoJwNp0gp3/ZqaQJMirX03Lft79+VRI3Cb
# y9bgK0P1YOvPz4eVCuN5vO2aY+a49Xd/azhxDUQqTUaGqbDYvl0HvZ0cyIrCWm7g
# j19+IhOuKxZR6ZyIOyvVcQhhCw38ZHpEGsx9QvB/BAaQkfAIRM+aGbd21M4Kxfbq
# OimNk1W/IiVd1dEHWtrkuNlJWbGisEcqbwjhhKMhK93JZCKIFmo87s+f73G4HCfE
# Tz/WoTGkQz3yxwR/OfIrl3/43cTh+XkXH8rlz4lZ3FX9IehUom33s2MuIygfykf9
# ZKabWWjodggq67hrZmcz56NKJXjp3H9jnoVxJ6nMMLbj561If7hPkC1m+3ixiNa5
# zruGgPT7uQ==
# SIG # End signature block
