<#
.SYNOPSIS
Offline check of the orphan atomic-write temporary rule of Invoke-SmartM365-ServerDataCleanup.
.VERSION
1.0.0
.NOTES
Builds a synthetic tenant DATA root under TEMP and runs the cleanup script in a child PowerShell 7
process, so a transport lock held here is a real lock from another process. Checks that an old
orphan is removed, a recent one and an old one beside a held transport lock are kept, a .previous
copy uses its longer age, weekly upload.pending markers and near-miss names are never touched,
and every removal is logged. No tenant, network or synchronized data access.
#>
#requires -Version 7.2
[CmdletBinding()]
param()
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$smartM365Root = Split-Path $PSScriptRoot -Parent
$cleanupScript = Join-Path $smartM365Root 'SmartInventory/Maintenance/Invoke-SmartM365-ServerDataCleanup.ps1'
$script:checks = 0
function Check { param([bool]$Condition, [string]$Message) if (-not $Condition) { throw $Message }; $script:checks++ }

$base = Join-Path ([IO.Path]::GetTempPath()) ('sm-orphan-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$root = Join-Path $base 'prod'
$reportFolder = Join-Path $base 'report'
$old = (Get-Date).AddDays(-10)
$older = (Get-Date).AddDays(-40)
$recent = (Get-Date).AddDays(-1)

function Write-SyntheticFile {
    param([string]$RelativePath, [datetime]$Time)
    $path = Join-Path $root $RelativePath
    [void][IO.Directory]::CreateDirectory((Split-Path $path -Parent))
    [IO.File]::WriteAllText($path, 'synthetic')
    [IO.File]::SetCreationTime($path, $Time)
    [IO.File]::SetLastWriteTime($path, $Time)
    return $path
}

function Invoke-Cleanup {
    param([switch]$Execute)
    $arguments = @('-NoProfile', '-NonInteractive', '-File', $cleanupScript, '-Tenant', 'prod', '-DataRootPath', $root,
        '-ReportDirectory', $reportFolder, '-EffectiveJobsPath', (Join-Path $base 'no-jobs.json'))
    if ($Execute) { $arguments += '-Execute' }
    $output = & pwsh @arguments 2>&1 | Out-String
    return [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = $output }
}

$lockHandle = $null
try {
    $files = [ordered]@{
        OldOrphan       = Write-SyntheticFile 'DATA-LAST/~sm0123456789ab.pending' $old
        OldCsvOrphan    = Write-SyntheticFile 'DATA-LAST/~sm0123456789ac.csv' $old
        RecentOrphan    = Write-SyntheticFile 'DATA-LAST/~smaaaaaaaaaaaa.tmp' $recent
        RecentPrevious  = Write-SyntheticFile 'DATA-ALL/Orchestrator/SRV/~smcccccccccccc.previous' $old
        OldPrevious     = Write-SyntheticFile 'DATA-ALL/Orchestrator/SRV/~smdddddddddddd.previous' $older
        LegacyOrphan    = Write-SyntheticFile ('DATA-ALL/Config/state.json.txt.' + ('e' * 32) + '.pending') $old
        LegacyProbe     = Write-SyntheticFile ('DATA-ALL/.smartm365-write-test-' + ('f' * 32) + '.tmp') $old
        LockedOrphan    = Write-SyntheticFile 'DATA-ALL/Locked/~smbbbbbbbbbbbb.pending' $old
        LockFile        = Write-SyntheticFile 'DATA-ALL/Locked/state.json.transport.lock' $old
        UploadMarker    = Write-SyntheticFile 'DATA-ALL/Intune/Apps/WeeklyHistory/2026-W30/upload.pending' $older
        BadHex          = Write-SyntheticFile 'DATA-LAST/~smXYZXYZXYZXYZ.pending' $old
        LongHex         = Write-SyntheticFile 'DATA-LAST/~sm0123456789abc.tmp' $old
        OtherExtension  = Write-SyntheticFile 'DATA-LAST/~sm0123456789ab.exe' $old
        LegacyCsvGuid   = Write-SyntheticFile ('DATA-LAST/Export.' + ('a' * 32) + '.csv') $old
        CanonicalCsv    = Write-SyntheticFile 'DATA-LAST/Export.csv' $old
    }
    $removed = @('OldOrphan', 'OldCsvOrphan', 'OldPrevious', 'LegacyOrphan', 'LegacyProbe')

    # Preview never deletes and plans exactly the eligible names (no lock is held yet).
    $preview = Invoke-Cleanup
    Check ($preview.ExitCode -eq 0) "Preview failed: $($preview.Output)"
    foreach ($name in $files.Keys) { Check (Test-Path -LiteralPath $files[$name]) "Preview removed $name." }
    $planned = @(Import-Csv (Join-Path $reportFolder 'orphan-temporaries.csv') | Where-Object Action -eq 'Remove' | ForEach-Object Name)
    $expectedNames = @(@($removed) + 'LockedOrphan' | ForEach-Object { Split-Path $files[$_] -Leaf })
    Check (@(Compare-Object ($planned | Sort-Object) ($expectedNames | Sort-Object)).Count -eq 0) "Preview planned: $($planned -join ', ')"
    $report = @(Import-Csv (Join-Path $reportFolder 'orphan-temporaries.csv'))
    Check (-not @($report | Where-Object Name -eq 'upload.pending').Count) 'upload.pending appears in the orphan report.'
    Check (@($report | Where-Object { $_.Name -eq '~smcccccccccccc.previous' -and $_.Reason -eq 'Recent' -and $_.MinimumAgeDays -eq '30' }).Count -eq 1) 'A 10-day .previous copy is not kept by its 30-day age.'

    # Execute while another process holds the transport lock of one folder.
    $lockHandle = [IO.File]::Open($files.LockFile, [IO.FileMode]::Open, [IO.FileAccess]::ReadWrite, [IO.FileShare]::None)
    $run = Invoke-Cleanup -Execute
    Check ($run.ExitCode -eq 0) "Execute failed: $($run.Output)"
    foreach ($name in $files.Keys) {
        $exists = Test-Path -LiteralPath $files[$name]
        if ($name -in $removed) { Check (-not $exists) "$name was not removed." }
        else { Check $exists "$name was removed." }
    }
    foreach ($name in $removed) {
        Check ($run.Output -match [regex]::Escape('Removed orphan temporary') + '.*' + [regex]::Escape((Split-Path $files[$name] -Leaf))) "Removal of $name was not logged."
    }
    $locked = @(Import-Csv (Join-Path $reportFolder 'orphan-temporaries.csv') | Where-Object Name -eq '~smbbbbbbbbbbbb.pending')
    Check ($locked.Count -eq 1 -and $locked[0].Reason -eq 'TransportLockHeld') 'An orphan beside a held transport lock was not kept as TransportLockHeld.'

    # Once the lock is released, the same old orphan becomes eligible; upload.pending still stays.
    $lockHandle.Dispose(); $lockHandle = $null
    $run = Invoke-Cleanup -Execute
    Check ($run.ExitCode -eq 0) "Second execute failed: $($run.Output)"
    Check (-not (Test-Path -LiteralPath $files.LockedOrphan)) 'Orphan was not removed after its folder lock was released.'
    Check (Test-Path -LiteralPath $files.LockFile) 'The transport lock file itself was removed.'
    Check (Test-Path -LiteralPath $files.UploadMarker) 'upload.pending was removed.'
    Check (Test-Path -LiteralPath $files.RecentOrphan) 'A recent orphan was removed.'
}
finally {
    if ($null -ne $lockHandle) { $lockHandle.Dispose() }
    if (Test-Path -LiteralPath $base) { Remove-Item -LiteralPath $base -Recurse -Force }
}
"Test-SmartM365ServerDataCleanupOrphanTemporariesOffline: $script:checks checks passed."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCChcKINtxbyGTyo
# n1b/YU7c3VuZfvY511ONNh4DnNl+M6CCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEINqd/dIisQGDlIkP5SxnDKfSJ23rSx9BiO8ezWehTLF+MA0GCSqG
# SIb3DQEBAQUABIIBgJ3mV/81o4ZD0nl0vYda+mMzE47jBDGnadSREGnk4D6muMb6
# scfxVPbzMr29fOXYhYz+wzeAQq3ZgrS/2uO15BF4IAKQVKtM7yX1CEJT5LuFCaUs
# s++JFMwzIafi1x/PF912mTBuy9519EgHElPzrmslC0LkBk9perCW+uYe4fSrWBl8
# s/j0XRqn0FU+b9eNuFsUkB9/DqmmeZHu2Rfyxdz/wAWR3ih/90v2xFXl9v1E8S0A
# Wh/u7zgQ/UNb9iEepbwyFw4zW3uT+AyH4aEUVbY12qPEGwsbsd/6c3zNz91tQEG3
# psUtraWnAI/VDVBvPlsG56MNiveOhx15Es4x9O7XWgI3Ar6Mgo6TxBctsscaAa7d
# y8iFPPmwnXMf3uAG6DyA0fpVwwGikK8h7vvtlqmVkyRFNmRhnzLEJXRxRD9Pajgr
# nJYmdvQDG2pz+cH8AVxMg946vzMGIBC8Ra79HubT7Mvi7hu/3VeiQtZ1h0SwQZWu
# aNEODHUtvoO8TXZtfaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjIxNDM3
# MzdaMC8GCSqGSIb3DQEJBDEiBCCWdCDXjptcGPEdzTOwTHTw29uUkZB66VQEzfE7
# F3qdXTANBgkqhkiG9w0BAQEFAASCAgA7KeZzHfWU7erx2uwIHxF31Y2h20vFqMnk
# 5QdS97V86ZmiGY5x7GeHIGMU+D9OJud4OJB5Yey+jLa/OefaEtIHWseCqGaDN/FX
# ZHZeHT4EgnBnGGAAFDQergQsO6/A5lfMYsdQa2jjch6Ccpyv7xoUyfDIPGufyB9N
# FSNK10prEjZOVhCH9jhD+Uv1K/iupOxKMNoo76bHnNMrhdeo1eyZOpHqlQYq/MeX
# Tm4mlK48DjARtRmKcJ+85ax8rI+O4zFdd3F+6+p73fT+NX5lT9zRmuusSufRBtJa
# iLTHtIahHv2op+/0kWIFMGj+ekAzKdfjzMgzgfqPFIaun4pK/FZL5tBwWoR6gO8Z
# qcExIvXMfWZZfzflyZH8ATXgG/uZx3bsNsSh0WPpN58ixxw67tPQk0gTytv0LvST
# rvBtn0OTn0NICxhil3zJL3yc1JRZLiqKTam6K09r7uZGS1AUTbci32g8w8OIti/S
# 574qpOw2LPcnSD89iSJbY7trYLbXJtKsN8/LjTl4wXee8e9bLPV5RPONLcA6MWhB
# FyUEQFwpxbcYbCk52YNoyxW97c8VcNy1lRjK0UOFc82WkSj1E33QQEQ9xn4Wf+j+
# 95XDBwyDixjD/ZTqyb/lD1rjtEqPBDtSMiWJLzpEyZ6C9Yc18qji2PLkvs0tmH9z
# yh379uNqXA==
# SIG # End signature block
