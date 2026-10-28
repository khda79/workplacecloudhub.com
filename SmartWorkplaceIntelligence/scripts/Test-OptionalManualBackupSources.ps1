[CmdletBinding()]
param([string]$TestRoot = (Join-Path ([IO.Path]::GetTempPath()) 'SmartM365-OfflineTests'))

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$fixture = Join-Path ([IO.Path]::GetFullPath($TestRoot)) ([guid]::NewGuid().ToString('N'))
$last = Join-Path $fixture 'DATA-LAST'
$contract = Join-Path $fixture 'prepared-source-contract.json.txt'
$manualNames = @(
    'M365_Backup_ProtectedMailboxes.csv',
    'M365_Backup_ProtectedSharePointSites.csv',
    'M365_Backup_ProtectedOneDriveAccounts.csv'
)
$policyNames = @('M365_BackupPolicyScope_MailboxCoverage.csv', 'M365_BackupPolicyScope_GroupMembers.csv')
$now = [datetime]::UtcNow
$checks = 0

function Check([bool]$Condition, [string]$Message) {
    if (-not $Condition) { throw $Message }
    $script:checks++
}

function Get-Plan {
    param([Collections.Generic.List[object]]$Issues)
    @(Get-PreparedSourcePlan -DataRoot $fixture -SourceContractPath $contract -MetadataOnly -OptionalSourceIssues $Issues)
}

function Run-Generator {
    param([string]$Output)
    & (Get-Process -Id $PID).Path -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'New-BackupResilienceEvidence.ps1') -DataRoot $last -OutputRoot $Output -AsOfDate $now.Date | Out-Null
    if ($LASTEXITCODE -ne 0) { throw "Backup generator failed with exit code $LASTEXITCODE." }
}

try {
    New-Item -ItemType Directory -Path $last -Force | Out-Null
    @{
        currentFiles = $policyNames
        optionalCurrentFiles = @($manualNames | ForEach-Object { @{ file = $_; maxAgeHours = 168 } })
        mappingFiles = @()
        dailyFiles = @()
        history = @()
    } | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $contract

    @([pscustomobject]@{
        TenantKey = 'synthetic-test'; MemberId = 'user@example.invalid';
        PrimarySmtpAddress = 'user@example.invalid'; MailboxFoundInInventory = 'true';
        RunDateUtc = $now.ToString('O'); GroupDisplayName = 'Backup policy';
        MailboxUserPrincipalName = 'user@example.invalid'; MemberUserPrincipalName = 'user@example.invalid';
        MemberMail = 'user@example.invalid'; MailboxDisplayName = 'Example User';
        MemberDisplayName = 'Example User'; RecipientTypeDetails = 'UserMailbox'
    }) | Export-Csv -LiteralPath (Join-Path $last $policyNames[0]) -NoTypeInformation
    @([pscustomobject]@{
        TenantKey = 'synthetic-test'; MemberId = 'user@example.invalid';
        MemberMail = 'user@example.invalid'; MemberUserPrincipalName = 'user@example.invalid';
        AccountEnabled = 'true'
    }) | Export-Csv -LiteralPath (Join-Path $last $policyNames[1]) -NoTypeInformation

    $module = Import-Module (Join-Path $PSScriptRoot 'PreparedEvidencePipeline.psm1') -Force -PassThru
    & $module { param($root) $script:ProductRoot = $root } $fixture

    $issues = [Collections.Generic.List[object]]::new()
    $plan = Get-Plan -Issues $issues
    Check ($plan.Count -eq 2 -and $issues.Count -eq 3) 'Missing interactive sources blocked selection or lost diagnostics.'

    $output = Join-Path $fixture 'output'
    Run-Generator -Output $output
    $rows = @(Import-Csv -LiteralPath (Join-Path $output 'BackupMailboxEvidence.csv'))
    Check ($rows.Count -eq 1 -and $rows[0].'Protection State' -eq 'Unknown' -and
        $rows[0].'Risk Priority' -eq 'Unknown' -and $rows[0].'Evidence Status' -eq 'Not measured') 'Missing interactive protection was misclassified as an unprotected mailbox.'

    @([pscustomobject]@{
        TenantKey = 'synthetic-test'; ProtectionUnitId = 'unit-1'; MailboxId = 'user@example.invalid';
        Mail = 'user@example.invalid'; UserPrincipalName = 'user@example.invalid';
        DisplayName = 'Example User'; MailboxType = 'UserMailbox'; ProtectionPolicyName = 'Backup policy';
        ErrorCode = ''; ErrorMessage = ''; ExportDateTime = $now.ToString('O'); Status = 'Protected'
    }) | Export-Csv -LiteralPath (Join-Path $last $manualNames[0]) -NoTypeInformation
    foreach ($name in $manualNames[1..2]) {
        'TenantKey,Workload,ProtectionUnitId,DirectoryObjectId,Status' | Set-Content -LiteralPath (Join-Path $last $name)
    }
    foreach ($name in $manualNames) {
        [IO.File]::SetLastWriteTimeUtc((Join-Path $last $name), $now.AddHours(-169))
    }
    $issues = [Collections.Generic.List[object]]::new()
    $plan = Get-Plan -Issues $issues
    Check ($plan.Count -eq 5 -and $issues.Count -eq 3 -and @($issues | Where-Object Phase -ne 'StaleUsed').Count -eq 0) 'Stale interactive sources were not kept as the latest available data, or blocked preparation.'

    foreach ($name in $manualNames) {
        [IO.File]::SetLastWriteTimeUtc((Join-Path $last $name), $now.AddHours(-1))
    }
    $issues = [Collections.Generic.List[object]]::new()
    $plan = Get-Plan -Issues $issues
    Check ($plan.Count -eq 5 -and $issues.Count -eq 0) 'Fresh interactive sources were not selected.'
    Run-Generator -Output $output
    $rows = @(Import-Csv -LiteralPath (Join-Path $output 'BackupMailboxEvidence.csv'))
    Check ($rows.Count -eq 1 -and $rows[0].'Protection State' -eq 'Protected') 'Fresh protected-mailbox evidence was not preserved.'

    $unmatched = @(Import-Csv -LiteralPath (Join-Path $last $manualNames[0]))
    $unmatched[0].MailboxId = 'other@example.invalid'
    $unmatched[0].Mail = 'other@example.invalid'
    $unmatched[0].UserPrincipalName = 'other@example.invalid'
    $unmatched | Export-Csv -LiteralPath (Join-Path $last $manualNames[0]) -NoTypeInformation
    Run-Generator -Output $output
    $rows = @(Import-Csv -LiteralPath (Join-Path $output 'BackupMailboxEvidence.csv'))
    Check (@($rows | Where-Object { $_.'Reconciliation State' -eq 'Expected not protected' -and $_.'Risk Priority' -eq 'High' }).Count -eq 1) 'Available protection evidence stopped identifying actual gaps.'

    [IO.File]::SetLastWriteTimeUtc((Join-Path $last $policyNames[0]), $now.AddHours(-169))
    $failure = ''
    try { Get-Plan -Issues ([Collections.Generic.List[object]]::new()) | Out-Null } catch { $failure = $_.Exception.Message }
    Check ($failure -match 'publication-age limit.*M365_BackupPolicyScope_MailboxCoverage') 'Required noninteractive policy scope lost its freshness gate.'

    "PASS: $checks optional manual backup source and unknown-protection checks."
}
finally {
    $resolvedRoot = [IO.Path]::GetFullPath($TestRoot).TrimEnd('\', '/')
    $resolvedFixture = [IO.Path]::GetFullPath($fixture)
    if ($resolvedFixture.StartsWith($resolvedRoot + [IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path $resolvedFixture -Leaf) -match '^[0-9a-f]{32}$') {
        Remove-Item -LiteralPath $resolvedFixture -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# SIG # Begin signature block
# MIIeYwYJKoZIhvcNAQcCoIIeVDCCHlACAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCD/7NO0GzueQuqE
# OWGEztiXf5zV79aEO6F1R2TJ90Br5KCCF/swggS9MIIDJaADAgECAhAebu87xzjh
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
# hvcNAQkEMSIEIL2qWBPuqDoz2YyN5Rtu53vqIyHwJ51CeHRv9EhBj6eLMA0GCSqG
# SIb3DQEBAQUABIIBgIwL0NydqGxZNIc5nJNQX7PRKMX2kV4XxVzYtagsVJAqleF+
# BsPRkIAN6ewBqPlgQE/Txp25VIRbV+1o3ICF3oGtC9cXCg2Bh92bgSqrydbYIttg
# LcojSB0XuIGxgxFrM/ltanYaqBEL1geeG6xNIYpHrOTWhrtV16Fw7wJ8v+2W+VP7
# UP8BMTNEXq2M5Hjs0CGnsZLagAaRgM0p5zeVUhNPNVAGclwzb9bSmMbI4bVovacV
# qYsOcD3+wwwA19wSMzisvxvZWqqKKwn6iyZDRbT1VL85PerkbIibiEGnoOc4nN92
# XovaB7mlq+2iPNcvVWVE/FmjJqwvsQVteoFIA5wYKi5KozrDoo/F5hlmqzaELw1E
# 65S2nKRlN+1+1Eb9mX4b9N8qaTUesqPf6WZdBCYv/SNwxG5j3QN+ASRLECRfTlRu
# y/VECHYoaA/vgpPYJq3SaxhzY59sDUVB/ewBN0W5RErdFmUABqWzyjiN3leAs1xp
# sO3jpCFU3YB9F9Aks6GCAyYwggMiBgkqhkiG9w0BCQYxggMTMIIDDwIBATB9MGkx
# CzAJBgNVBAYTAlVTMRcwFQYDVQQKEw5EaWdpQ2VydCwgSW5jLjFBMD8GA1UEAxM4
# RGlnaUNlcnQgVHJ1c3RlZCBHNCBUaW1lU3RhbXBpbmcgUlNBNDA5NiBTSEEyNTYg
# MjAyNSBDQTECEAhP3DNPfkVO28MPj/mSGDUwDQYJYIZIAWUDBAIBBQCgaTAYBgkq
# hkiG9w0BCQMxCwYJKoZIhvcNAQcBMBwGCSqGSIb3DQEJBTEPFw0yNjEwMjgxODIw
# MjJaMC8GCSqGSIb3DQEJBDEiBCBqdeOPOU45jUVJfo+gVWgs1BzQiJJmxlJvsK1r
# GmlRejANBgkqhkiG9w0BAQEFAASCAgBEtydjrGymZaNdM7VPAVFIXtcKBwgweQPm
# e75XfNIxCYDY/2dwiSjRKXVN9f563OvwZP1cHuPqB7YwOEviu75VK2SQiXexypCq
# NlgmQKlPQ1CqIV6/OxFWfvxamvgglmsVxh2Yue3zPZpE2WwVHSOSO9uf/zairbTw
# 6ICEf1SP7Vq7Kw5NGeiQQLw/TuTHifGEPn8M4yRnRVqPMLXDzkd+5UBG/DWwAEjg
# zo3IdjP5WBVR45TRfqbMzJTr+L3IAlNh7IzsZs9EB8lSOojh/4nCOFqzaUGx804E
# GC+B6R9llh6PWn1uJsw7Vpe3WvZsvmrqfnpxwt7bu+cNVxXR9a+OYJj5oc2STPIZ
# 1qExyfC654QzMz7xdCsq527y2SpmvJa7sWbpvytrG/qSzQoxWETqyExM2BI1HNj0
# +beG1tXou0dgTxPr+gDr12e/gAkjkWXNEXpcaYVGRI4mip3HDJdkSsZnMSt2HfC1
# 4MSynN5XKBBKcLNCTLDDkJya+Fy5WWJyfGGmIRgITUXRmPUSQM23ekRZQnR2NPig
# jsQMuTViTilWWq+19PwsKXtvEFjI00op0EQvhzkL+7v73SAgxV2q3kw8RR/0uqEo
# Cv3mwCf9pTjPLJA50OWUY4wubDeY9CxDzl9R7VFSzStwBnD6885Ciy9nsTVR5JqG
# 8fh//rWxYg==
# SIG # End signature block
