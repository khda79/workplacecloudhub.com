<#
.SYNOPSIS
Offline check of the linked prepared-batch retention of Invoke-SmartM365-ServerDataCleanup.
.VERSION
1.0.0
.NOTES
Builds a synthetic tenant DATA root under TEMP with a published chain current -> previous -> three
older batches and runs the cleanup script in a child PowerShell 7 process. Checks that by default
no linked batch is planned or removed and all are inventoried as kept; that enabled retirement
honours the retained count and the minimum age, never touches current and previous, and preserves
SHA256-identical audit receipts. No tenant, network or synchronized data access.
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

$base = Join-Path ([IO.Path]::GetTempPath()) ('sm-batches-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
$root = Join-Path $base 'prod'
$reportFolder = Join-Path $base 'report'
$prepared = Join-Path $root 'DATA-POWERBI'
$tenantKey = 'synthetic-tenant-key'

function Get-SyntheticBatchId {
    param([int]$AgeDays, [string]$Suffix)
    return [datetime]::UtcNow.AddDays(-$AgeDays).ToString('yyyyMMddTHHmmssfff', [Globalization.CultureInfo]::InvariantCulture) + 'Z-' + $Suffix
}

function Write-Json {
    param([string]$Path, [object]$Value)
    [void][IO.Directory]::CreateDirectory((Split-Path $Path -Parent))
    [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 4))
}

function Write-SyntheticBatch {
    param([string]$Id, [string]$PreviousId)
    $folder = Join-Path (Join-Path $prepared 'batches') $Id
    Write-Json (Join-Path $folder 'batch.json.txt') ([ordered]@{ BatchId = $Id; TenantKey = $tenantKey })
    Write-Json (Join-Path $folder 'validation.json.txt') ([ordered]@{ BatchId = $Id; Status = 'Valid' })
    Write-Json (Join-Path $folder 'tables/Table.csv') 'synthetic'
    $manifestHash = (Get-FileHash -LiteralPath (Join-Path $folder 'batch.json.txt') -Algorithm SHA256).Hash
    Write-Json (Join-Path $folder 'current.json.txt') ([ordered]@{
        BatchId = $Id; PreviousBatchId = $PreviousId; TenantKey = $tenantKey; ManifestSHA256 = $manifestHash })
    return $folder
}

function Invoke-Cleanup {
    param([string[]]$Extra = @())
    $arguments = @('-NoProfile', '-NonInteractive', '-File', $cleanupScript, '-Tenant', 'prod', '-DataRootPath', $root,
        '-ReportDirectory', $reportFolder, '-EffectiveJobsPath', (Join-Path $base 'no-jobs.json')) + $Extra
    $output = & pwsh @arguments 2>&1 | Out-String
    return [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = $output }
}

function Get-PlannedBatchId {
    return @(Import-Csv (Join-Path $reportFolder 'verified-actions.csv') |
        Where-Object Category -eq 'LinkedPreparedBatch' | ForEach-Object { Split-Path $_.RelativePath -Leaf } | Sort-Object)
}

try {
    [void][IO.Directory]::CreateDirectory((Join-Path $root 'DATA-ALL'))
    [void][IO.Directory]::CreateDirectory((Join-Path $root 'DATA-LAST'))
    $ids = [ordered]@{
        Current  = Get-SyntheticBatchId 1 'aaaaaaa1'
        Previous = Get-SyntheticBatchId 2 'aaaaaaa2'
        Older1   = Get-SyntheticBatchId 5 'aaaaaaa3'
        Older2   = Get-SyntheticBatchId 100 'aaaaaaa4'
        Older3   = Get-SyntheticBatchId 200 'aaaaaaa5'
    }
    $folders = [ordered]@{
        Current  = Write-SyntheticBatch $ids.Current $ids.Previous
        Previous = Write-SyntheticBatch $ids.Previous $ids.Older1
        Older1   = Write-SyntheticBatch $ids.Older1 $ids.Older2
        Older2   = Write-SyntheticBatch $ids.Older2 $ids.Older3
        Older3   = Write-SyntheticBatch $ids.Older3 ''
    }
    Write-Json (Join-Path $prepared 'current.json.txt') ([ordered]@{
        BatchId = $ids.Current; PreviousBatchId = $ids.Previous; TenantKey = $tenantKey })

    # Default: the share keeps the whole linked history, in Preview and Execute.
    $preview = Invoke-Cleanup
    Check ($preview.ExitCode -eq 0) "Default preview failed: $($preview.Output)"
    Check (@(Get-PlannedBatchId).Count -eq 0) 'Default preview planned a linked prepared batch.'
    Check ($preview.Output -match '3 found, 0 planned for retirement, 3 kept .*retirement not enabled') "Default retention was not logged: $($preview.Output)"
    $inventory = @(Import-Csv (Join-Path $reportFolder 'all-files.csv'))
    foreach ($name in 'Older1', 'Older2', 'Older3') {
        $rows = @($inventory | Where-Object { $_.RelativePath -like ('DATA-POWERBI\batches\' + $ids[$name] + '\*') })
        Check ($rows.Count -eq 4) "Inventory misses the $name batch files."
        Check (-not @($rows | Where-Object { $_.Decision -ne 'Keep' -or $_.Reason -notlike 'Linked prepared batch kept: share retains*' }).Count) "$name batch is not inventoried as kept by default."
    }
    $run = Invoke-Cleanup @('-Execute')
    Check ($run.ExitCode -eq 0) "Default execute failed: $($run.Output)"
    foreach ($name in $folders.Keys) { Check (Test-Path -LiteralPath $folders[$name]) "Default execute removed the $name batch." }
    Check (-not (Test-Path -LiteralPath (Join-Path $prepared 'retired'))) 'Default execute wrote retired receipts.'

    # Enabled: the retained count and the minimum age both keep batches.
    $preview = Invoke-Cleanup @('-RetireLinkedPreparedBatches', '-RetainedLinkedPreparedBatchCount', '1')
    Check ($preview.ExitCode -eq 0) "Count preview failed: $($preview.Output)"
    Check (@(Compare-Object @(Get-PlannedBatchId) (@($ids.Older2, $ids.Older3) | Sort-Object)).Count -eq 0) "Count preview planned: $(@(Get-PlannedBatchId) -join ', ')"
    $kept = @(Import-Csv (Join-Path $reportFolder 'all-files.csv') | Where-Object { $_.RelativePath -like ('DATA-POWERBI\batches\' + $ids.Older1 + '\*') })
    Check (-not @($kept | Where-Object { $_.Decision -ne 'Keep' -or $_.Reason -notlike '*within the 1 retained older batches*' }).Count) 'Retained-count batch is not inventoried as kept.'

    $retireOldest = @('-RetireLinkedPreparedBatches', '-RetainedLinkedPreparedBatchCount', '1', '-LinkedPreparedBatchMinimumAgeDays', '150')
    $preview = Invoke-Cleanup $retireOldest
    Check ($preview.ExitCode -eq 0) "Age preview failed: $($preview.Output)"
    Check (@(Compare-Object @(Get-PlannedBatchId) @($ids.Older3)).Count -eq 0) "Age preview planned: $(@(Get-PlannedBatchId) -join ', ')"

    $receiptHashes = @{}
    foreach ($receipt in 'current', 'batch', 'validation') {
        $receiptHashes[$receipt] = (Get-FileHash -LiteralPath (Join-Path $folders.Older3 ($receipt + '.json.txt')) -Algorithm SHA256).Hash
    }
    $run = Invoke-Cleanup ($retireOldest + '-Execute')
    Check ($run.ExitCode -eq 0) "Age execute failed: $($run.Output)"
    Check (-not (Test-Path -LiteralPath $folders.Older3)) 'The oldest batch was not retired.'
    foreach ($name in 'Current', 'Previous', 'Older1', 'Older2') { Check (Test-Path -LiteralPath $folders[$name]) "Age execute removed the $name batch." }
    foreach ($receipt in $receiptHashes.Keys) {
        $copy = Join-Path (Join-Path (Join-Path $prepared 'retired') $ids.Older3) ($receipt + '.json.txt')
        Check ((Test-Path -LiteralPath $copy) -and (Get-FileHash -LiteralPath $copy -Algorithm SHA256).Hash -eq $receiptHashes[$receipt]) "Audit receipt $receipt is missing or differs."
    }

    # Fully enabled: every older linked batch goes, current and previous always stay.
    $run = Invoke-Cleanup @('-RetireLinkedPreparedBatches', '-Execute')
    Check ($run.ExitCode -eq 0) "Full execute failed: $($run.Output)"
    foreach ($name in 'Older1', 'Older2') { Check (-not (Test-Path -LiteralPath $folders[$name])) "Full execute kept the $name batch." }
    foreach ($name in 'Current', 'Previous') { Check (Test-Path -LiteralPath (Join-Path $folders[$name] 'tables/Table.csv')) "Full execute touched the $name batch." }
    foreach ($name in 'Older1', 'Older2', 'Older3') {
        Check (@(Get-ChildItem -LiteralPath (Join-Path (Join-Path $prepared 'retired') $ids[$name]) -File).Count -eq 3) "Retired receipts of $name are incomplete."
    }
}
finally {
    if (Test-Path -LiteralPath $base) { Remove-Item -LiteralPath $base -Recurse -Force }
}
"Test-SmartM365ServerDataCleanupPreparedBatchRetentionOffline: $script:checks checks passed."

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCCxSZOnmtrVvOYE
# KFXaQr8eAbCpPnGXVMW7Am85rQsBjqCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIGW6w1xspHDt3wSZCHiOCgjozkc0xOqv/1CQ8vxavmCEMA0GCSqG
# SIb3DQEBAQUABIIBgBwUEa2zeHy35rGkSwpEYuAYOx+GXLOTKkQcsqPBqQWqKyNI
# 7oipQYEP8FShlAQMDDiXCiVeLiWxQtHJprzV2Z6LZG1hwWyAnkrfQv6Ohbf+e2EL
# Fnolp8FUxsu5V82Z0vnScTKwHRkVQJQItLNzrensiLNibSgEbc89uaTnbUUJl+TO
# 2QB6Y5GJzMLlubwR7YjVtMMz8m2fQ2xpa8IZv0rAvpz7rdkbD+V9z5CXu8ajJCGd
# 4WCnEMOyPY7Cx9yJ44gsOGoZE3hcyLzAx2s0SjMadso2S/FM6gsk9Ud4VSD/fvKa
# M5X7A4Ok3ekHroOPDKIO4Bp0o4B6NG9MlOqtKxviq1O5e5syX/F34V0AYlmSuLZ0
# xIg4bvv3wNpwuCGyeK/r1SyG5LONcBVYcGlz51i3bYwDX9Onj6WyEreQmMhWaPfe
# Fr4CRHgJShgouZGNgdZySwY5PLCIZN09EBDholaHpVJEmbwMH3EeNjkl9awz5ZWY
# 5Amv06I0cXqE+DE7WaGCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjIxNTM4
# MzdaMC8GCSqGSIb3DQEJBDEiBCBkDHztwH1t8XZsAbgcfIluURqW/royYBConosU
# TNfg9jANBgkqhkiG9w0BAQEFAASCAgAZ7BFSqF/Z+Kwp+//EC4bj0ITUk9TgG6nc
# nyjFl6zdJlrfC8HY4DZXxDSnBziHnap0ROo7nAjvbACct1FBQMBx/19OJNxVNAr2
# 2NeTniD9k3SQZ5DYCbUZotLkQCAyF/9T70JpRNlVO427FoX+IpLHfbq7VMelgSlC
# khtnWN/63vjLDHNlJeAdFSWIWwu3rDOVDDS79Ql86RmrPJGPHPsN4UVXOj/BkL06
# 2gEyip6jVcVRUMUgg7XZOxJ0YF6UbGqHDIkWqwzK8R7mNqkLldGQDYZoe01cGFc6
# OouZtNF6nnWvhdUa+aLSKijyh2QBvpwl/GZPygIW1s/GmP1Wf/KZgpXJ7d2wpmel
# KKBCQxlPILdguWv63Otrwl8wcO8+Vkwg3ugo5QrxakqEYtZDSbYkH2nLbrkPkIDY
# TKG4rZ3Yh8Pe6muEGJDCo0ux8aw+RS0im3CQ0lmduo5+oCzwx+hycfotEhSp8i7w
# zXvZYSOpoz3vceabz/l620T2YZ8Aqp6mjFNnRY8/TFUM3tmh6k1coDhJoh5ie/OI
# u9EOyYWE9nTrMZcaSk9eElLRZFCmZWQ/IahJfC3iCZCB1AWU/R6LSWr72adRo8/G
# bBLhAvwUmo5N0H9QwiGjvmScMnGgUDImgC2zmWP7/vbskIikjPtoM/Hwez9sD8JP
# jHVWGMFrCA==
# SIG # End signature block
